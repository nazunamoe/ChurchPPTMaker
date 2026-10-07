"""예배 순서 PPT 제작 프로그램 v1.0 — 메인 윈도우.

레이아웃은 UI 초안 기준 (CLAUDE.md §4 참고).
"""
from __future__ import annotations

import datetime as dt
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from bible_bridge import Bible2PPT, parse_references
from config_manager import MAX_PRESETS, ConfigManager, PresetManager
from models import (KIND_BIBLE, KIND_CUSTOM, KIND_HYMN, KIND_LABELS, KIND_PRAYER, ServiceItem, make_item,
                    validate_item)
from ui_widgets import (ACCENT, BG, BORDER, BTN_HOVER, CARD, DANGER, KIND_COLORS, PILL_BORDER, SUBTLE, TEXT,
                        TEXT2, BibleField, CalendarPopup, DragController, ItemDialog, LogDialog, SettingsDialog,
                        accent_button, card, font, post_to_ui, pump_ui_queue, run_in_thread, secondary_button)
from worship_builder import BuildError, BuildRequest, WorshipBuilder, list_ppt_files, output_filename

APP_TITLE = "예배 순서 PPT 제작 프로그램 v1.0"
DATE_FMT = "%Y-%m-%d"
MODES = {"라이트": "Light", "다크": "Dark", "시스템": "System"}
# PyInstaller onefile 은 번들 파일을 sys._MEIPASS 에 풀어 둠
ICON_PATH = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / "assets" / "app.ico"


def _set_app_user_model_id():
    """작업 표시줄이 python.exe 아이콘 대신 창 아이콘을 쓰도록 별도 앱으로 등록 (Tk 생성 전 호출)."""
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ChurchPPTMaker.App")
    except (AttributeError, OSError):
        pass


def _powerpoint_exe() -> Path | None:
    """레지스트리 App Paths 에서 POWERPNT.EXE 경로 (Show 동사가 없을 때 대체용)."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe") as k:
            p = Path(winreg.QueryValue(k, None))
        return p if p.is_file() else None
    except OSError:
        return None


class App(ctk.CTk):
    def __init__(self):
        self.cfg = ConfigManager()
        self.presets = PresetManager()
        ctk.set_appearance_mode(self.cfg.get("appearance_mode") or "System")
        ctk.set_default_color_theme("blue")
        super().__init__(fg_color=BG)
        self.title(APP_TITLE)
        if ICON_PATH.is_file():
            self.iconbitmap(default=str(ICON_PATH))  # CTk 기본 아이콘 대신 (default= 는 다이얼로그 등에도 적용)
        self._fit_to_screen(1220, 820)

        self.items: list[ServiceItem] = []
        self.row_widgets: list[ctk.CTkFrame] = []
        self.file_buttons: dict[str, ctk.CTkButton] = {}
        self.selected_file: str | None = None
        self.bible_values: list[str] = list(self.cfg.get("bible2ppt", "bible_list") or [])
        self._bible_loading = False
        self.log_lines: list[str] = []
        self.building = False
        self._save_job = None
        self._calendar: CalendarPopup | None = None

        self.drag = DragController(self, self._on_drag_hover, self._on_drag_drop, self._on_drag_click)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self._build_header()
        self._build_info_bar()
        self._build_body()
        self._build_bottom_bar()

        self._refresh_files()
        self._refresh_presets()
        self._render_list()
        self._update_bible_status()
        self._load_bible_list_if_needed()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        pump_ui_queue(self)

    def _fit_to_screen(self, w: int, h: int):
        """화면 배율(DPI)을 고려해 창이 화면 밖으로 나가지 않게 크기/위치를 정한다."""
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(w, sw - 40), min(h, sh - 90)
        scale = ctk.ScalingTracker.get_window_scaling(self)  # 크기는 논리 단위, 위치는 실제 픽셀
        self.geometry(f"{w}x{h}+{int(max((sw - w) // 2, 0) * scale)}+{int(max((sh - h) // 3, 0) * scale)}")
        self.minsize(min(1000, w), min(660, h))

    # ================================================================ 상단
    def _build_header(self):
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 8))
        ctk.CTkLabel(bar, text=APP_TITLE, font=font(20, "bold"), text_color=TEXT).pack(side="left")
        secondary_button(bar, "⚙  설정", self._open_settings, width=88).pack(side="right")
        current = {v: k for k, v in MODES.items()}.get(self.cfg.get("appearance_mode"), "시스템")
        self.mode_seg = ctk.CTkSegmentedButton(bar, values=list(MODES), command=self._set_mode, font=font(12),
                                               selected_color=ACCENT, height=32)
        self.mode_seg.set(current)
        self.mode_seg.pack(side="right", padx=10)

    def _build_info_bar(self):
        bar = card(self)
        bar.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))

        def label(text):
            return ctk.CTkLabel(bar, text=text, font=font(13, "bold"), text_color=TEXT)

        label("교회 이름").pack(side="left", padx=(18, 8), pady=14)
        self.church_entry = ctk.CTkEntry(bar, width=230, height=34, font=font(13), placeholder_text="예: 사랑의교회")
        self.church_entry.insert(0, self.cfg.get("church_name") or "")
        self.church_entry.pack(side="left")
        self.church_entry.bind("<KeyRelease>", lambda e: self._schedule_save())

        label("날짜").pack(side="left", padx=(28, 8))
        self.date_entry = ctk.CTkEntry(bar, width=130, height=34, font=font(13), placeholder_text="YYYY-MM-DD")
        self.date_entry.insert(0, dt.date.today().strftime(DATE_FMT))
        self.date_entry.pack(side="left")
        self.cal_btn = secondary_button(bar, "📅", self._open_calendar, width=36)
        self.cal_btn.configure(height=34)
        self.cal_btn.pack(side="left", padx=(4, 4))
        secondary_button(bar, "금일 지정", self._set_today, width=80).pack(side="left")

        label("예배 종류").pack(side="left", padx=(28, 8))
        self.worship_entry = ctk.CTkEntry(bar, width=170, height=34, font=font(13), placeholder_text="예: 주일오전예배")
        self.worship_entry.insert(0, self.cfg.get("worship_type") or "")
        self.worship_entry.pack(side="left", padx=(0, 18))
        self.worship_entry.bind("<KeyRelease>", lambda e: self._schedule_save())

    # ================================================================ 본문
    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=2, column=0, sticky="nsew", padx=20)
        body.grid_columnconfigure(0, weight=0, minsize=380)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)
        self._build_left(body)
        self._build_order_panel(body)

    # ---------------------------------------------------------------- 좌측 패널
    def _build_left(self, parent):
        left = ctk.CTkFrame(parent, fg_color="transparent", width=380)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)

        # 폴더
        fc = card(left)
        fc.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        fc.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(fc, text="폴더", font=font(15, "bold"), text_color=TEXT).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=14, pady=(10, 4))
        self.path_vars: dict[str, tk.StringVar] = {}
        for r, (key, name) in enumerate([("template_dir", "템플릿"), ("hymn_dir", "찬송가"), ("output_dir", "출력")], 1):
            ctk.CTkLabel(fc, text=name, font=font(12), text_color=TEXT2, width=44, anchor="w").grid(
                row=r, column=0, padx=(14, 4), pady=3, sticky="w")
            var = tk.StringVar(value=str(self.cfg.resolve_path(key)))
            self.path_vars[key] = var
            ctk.CTkEntry(fc, textvariable=var, state="readonly", height=30, font=font(11),
                         text_color=TEXT2).grid(row=r, column=1, sticky="ew", pady=3)
            b1 = secondary_button(fc, "📁", lambda k=key: self._choose_folder(k), width=32)
            b1.grid(row=r, column=2, padx=(4, 0), pady=3)
            b2 = secondary_button(fc, "↗", lambda k=key: self._open_folder(k), width=32)
            b2.grid(row=r, column=3, padx=(4, 14), pady=3)
        ctk.CTkFrame(fc, width=1, height=6, fg_color="transparent").grid(row=4, column=0)

        # 템플릿 파일 목록
        tc = card(left)
        tc.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        tc.grid_columnconfigure(0, weight=1)
        tc.grid_rowconfigure(2, weight=1)
        head = ctk.CTkFrame(tc, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 4))
        ctk.CTkLabel(head, text="템플릿 파일", font=font(15, "bold"), text_color=TEXT).pack(side="left")
        self.file_count = ctk.CTkLabel(head, text="", font=font(12), text_color=TEXT2)
        self.file_count.pack(side="left", padx=8)
        secondary_button(head, "↻", self._refresh_files, width=32).pack(side="right")
        self.file_search = ctk.CTkEntry(tc, height=30, font=font(12), placeholder_text="🔍  파일 이름 검색")
        self.file_search.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 6))
        self.file_search.bind("<KeyRelease>", lambda e: self._render_files())
        self.file_list = ctk.CTkScrollableFrame(tc, fg_color=SUBTLE, corner_radius=8, height=120)
        self.file_list.grid(row=2, column=0, sticky="nsew", padx=14)
        foot = ctk.CTkFrame(tc, fg_color="transparent")
        foot.grid(row=3, column=0, sticky="ew", padx=14, pady=(6, 12))
        ctk.CTkLabel(foot, text="끌어다 놓기 · 더블클릭으로 추가", font=font(11), text_color=TEXT2).pack(side="left")
        secondary_button(foot, "순서에 추가", self._add_selected_file, width=90).pack(side="right")

        # 성경 구절 추가 (Bible2PPT 연동)
        bc = card(left)
        bc.grid(row=2, column=0, sticky="ew")
        bc.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(bc, text="성경 구절 추가  (Bible2PPT 연동)", font=font(15, "bold"), text_color=TEXT).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(10, 4))
        self.verse_entry = ctk.CTkEntry(bc, height=34, font=font(13), placeholder_text="예: 창1:1-5, 롬1-3, 요3:16")
        self.verse_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 6))
        self.verse_entry.bind("<Return>", lambda e: self._add_bible_from_form())
        brow = ctk.CTkFrame(bc, fg_color="transparent")
        brow.grid(row=2, column=0, columnspan=2, sticky="ew", padx=14)
        self.bible_field = BibleField(brow, values=self.bible_values, width=200, font_size=12, placeholder="역본 (비우면 기본값)")
        self.bible_field.set(self.cfg.get("bible2ppt", "bible") or "")
        self.bible_field.pack(side="left", fill="x", expand=True)
        secondary_button(brow, "↻", self._load_bible_list, width=32).pack(side="left", padx=4)
        accent_button(brow, "순서에 추가", self._add_bible_from_form, width=96).pack(side="left")
        self.bible_status = ctk.CTkLabel(bc, text="", font=font(11), text_color=TEXT2, anchor="w", justify="left")
        self.bible_status.grid(row=3, column=0, columnspan=2, sticky="w", padx=14, pady=(4, 10))

    # ---------------------------------------------------------------- 우측: 순서 패널 (CLAUDE.md §4 초안)
    def _build_order_panel(self, parent):
        oc = card(parent)
        oc.grid(row=0, column=1, sticky="nsew")
        oc.grid_columnconfigure(0, weight=1)
        oc.grid_rowconfigure(2, weight=1)
        self.order_card = oc

        head = ctk.CTkFrame(oc, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 6))
        ctk.CTkLabel(head, text="순서", font=font(17, "bold"), text_color=TEXT).pack(side="left")
        self.count_label = ctk.CTkLabel(head, text="", font=font(12), text_color=TEXT2)
        self.count_label.pack(side="left", padx=8)
        ctk.CTkButton(head, text="전체 비우기", width=80, height=28, fg_color="transparent", hover_color=BTN_HOVER,
                      text_color=DANGER, font=font(12), command=self._clear_items).pack(side="right")

        # 프리셋 행: [프리셋 ▼] [저장] [불러오기] [삭제]
        prow = ctk.CTkFrame(oc, fg_color="transparent")
        prow.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 8))
        self.preset_combo = ctk.CTkComboBox(prow, height=34, font=font(13), dropdown_font=font(13), values=[])
        self.preset_combo.pack(side="left", fill="x", expand=True)
        self.preset_combo.set("")
        for text, cmd in (("저장", self._save_preset), ("불러오기", self._load_preset), ("삭제", self._delete_preset)):
            b = secondary_button(prow, text, cmd, width=72 if text != "불러오기" else 84)
            b.configure(height=34)
            b.pack(side="left", padx=(6, 0))

        # 순서 목록
        self.list_zone = ctk.CTkFrame(oc, fg_color=SUBTLE, corner_radius=10, border_width=1, border_color=BORDER)
        self.list_zone.grid(row=2, column=0, sticky="nsew", padx=16)
        self.list_zone.grid_columnconfigure(0, weight=1)
        self.list_zone.grid_rowconfigure(0, weight=1)
        self.scroll = ctk.CTkScrollableFrame(self.list_zone, fg_color="transparent", corner_radius=10)
        self.scroll.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self.scroll.grid_columnconfigure(0, weight=1)

        # 팔레트: [기도] [찬송] [말씀] [사용자 지정] — 드래그&드롭
        pal = ctk.CTkFrame(oc, fg_color="transparent")
        pal.grid(row=3, column=0, sticky="ew", padx=16, pady=(10, 4))
        for kind in (KIND_PRAYER, KIND_HYMN, KIND_BIBLE, KIND_CUSTOM):
            b = ctk.CTkButton(pal, text=KIND_LABELS[kind], width=96 if kind != KIND_CUSTOM else 120, height=38,
                              corner_radius=8, fg_color=CARD, hover_color=BTN_HOVER, border_width=2,
                              border_color=KIND_COLORS[kind], text_color=TEXT, font=font(14, "bold"))
            b.pack(side="left", padx=(0, 8))
            b.configure(cursor="hand2")
            self.drag.attach(b, ("new", kind), KIND_LABELS[kind])
        ctk.CTkLabel(oc, text="버튼을 위 순서 목록으로 끌어다 놓으면 해당 위치에 추가됩니다 (클릭하면 맨 아래에 추가)",
                     font=font(11), text_color=TEXT2).grid(row=4, column=0, sticky="w", padx=16, pady=(0, 12))

    # ================================================================ 하단
    def _build_bottom_bar(self):
        bar = card(self)
        bar.grid(row=3, column=0, sticky="ew", padx=20, pady=(10, 16))
        bar.grid_columnconfigure(0, weight=1)
        left = ctk.CTkFrame(bar, fg_color="transparent")
        left.grid(row=0, column=0, sticky="ew", padx=16, pady=12)
        self.status = ctk.CTkLabel(left, text="준비됨", font=font(12), text_color=TEXT2, anchor="w")
        self.status.pack(fill="x")
        self.progress = ctk.CTkProgressBar(left, height=6, progress_color=ACCENT)
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(4, 0))
        secondary_button(bar, "로그", self._show_log, width=60).grid(row=0, column=1, padx=(0, 8))
        self.open_after = ctk.BooleanVar(value=bool(self.cfg.get("ppt", "open_after_build")))
        self.slideshow = ctk.BooleanVar(value=bool(self.cfg.get("ppt", "start_slideshow")))
        ctk.CTkCheckBox(bar, text="생성 후 열기", variable=self.open_after, font=font(12), checkbox_width=20,
                        checkbox_height=20, fg_color=ACCENT, command=self._on_open_after).grid(row=0, column=2, padx=8)
        self.slideshow_cb = ctk.CTkCheckBox(
            bar, text="슬라이드 쇼로 시작", variable=self.slideshow, font=font(12), checkbox_width=20,
            checkbox_height=20, fg_color=ACCENT,
            command=lambda: self.cfg.set("ppt", "start_slideshow", self.slideshow.get()))
        self.slideshow_cb.grid(row=0, column=3, padx=8)
        self._on_open_after(save=False)
        self.build_btn = accent_button(bar, "예배 PPT 병합 및 생성", self._build, width=220, height=42,
                                       font=font(15, "bold"))
        self.build_btn.grid(row=0, column=4, padx=(8, 16), pady=12)

    def _on_open_after(self, save: bool = True):
        """'슬라이드 쇼로 시작'은 '생성 후 열기'가 켜져 있을 때만 선택 가능."""
        if save:
            self.cfg.set("ppt", "open_after_build", self.open_after.get())
        self.slideshow_cb.configure(state="normal" if self.open_after.get() else "disabled")

    # ================================================================ 순서 목록 렌더링
    def _render_list(self):
        for w in self.scroll.winfo_children():
            w.destroy()
        self.row_widgets = []
        self.indicator = ctk.CTkFrame(self.scroll, height=4, fg_color=ACCENT, corner_radius=0, border_width=0)

        if not self.items:
            ctk.CTkLabel(self.scroll, text="아래 [기도] [찬송] [말씀] [사용자 지정] 버튼을 끌어다 놓아\n예배 순서를 추가하세요.",
                         font=font(13), text_color=TEXT2, justify="center").pack(pady=60)
        for idx, item in enumerate(self.items):
            self.row_widgets.append(self._make_row(idx, item))
        self.count_label.configure(text=f"{len(self.items)}개 항목")

    def _make_row(self, idx: int, item: ServiceItem) -> ctk.CTkFrame:
        row = ctk.CTkFrame(self.scroll, fg_color="transparent", height=46)
        row.pack(fill="x", padx=6, pady=3)

        handle = ctk.CTkLabel(row, text="⋮⋮", width=22, font=font(14, "bold"), text_color=TEXT2, cursor="fleur")
        handle.pack(side="left", padx=(2, 6))
        self.drag.attach(handle, ("move", idx), f"{item.label}  {item.value}".strip())

        left = ctk.CTkButton(row, text=item.label, width=140, height=36, corner_radius=10, fg_color=CARD,
                             hover_color=BTN_HOVER, border_width=2, border_color=KIND_COLORS[item.kind],
                             text_color=TEXT, font=font(13, "bold"), command=lambda i=idx: self._edit_item(i))
        left.pack(side="left")
        ctk.CTkFrame(row, height=3, fg_color=PILL_BORDER, corner_radius=0, border_width=0).pack(side="left", fill="x", expand=True)
        right = ctk.CTkButton(row, text=item.display_value, width=170, height=36, corner_radius=10, fg_color=CARD,
                              hover_color=BTN_HOVER, border_width=1, border_color=PILL_BORDER,
                              text_color=TEXT if item.value else TEXT2, font=font(13),
                              command=lambda i=idx: self._edit_item(i))
        right.pack(side="left")

        tools = ctk.CTkFrame(row, fg_color="transparent")
        tools.pack(side="left", padx=(8, 0))
        for text, cmd, color in (("▲", lambda i=idx: self._move_item(i, -1), TEXT2),
                                 ("▼", lambda i=idx: self._move_item(i, 1), TEXT2),
                                 ("✕", lambda i=idx: self._remove_item(i), DANGER)):
            ctk.CTkButton(tools, text=text, width=28, height=28, corner_radius=6, fg_color="transparent",
                          hover_color=BTN_HOVER, text_color=color, font=font(12), command=cmd).pack(side="left")
        return row

    # ================================================================ 드래그 & 드롭
    def _in_zone(self, x: int, y: int) -> bool:
        z = self.order_card
        return z.winfo_rootx() <= x <= z.winfo_rootx() + z.winfo_width() and \
            z.winfo_rooty() <= y <= z.winfo_rooty() + z.winfo_height()

    def _drop_index(self, y: int) -> int:
        for i, row in enumerate(self.row_widgets):
            if y < row.winfo_rooty() + row.winfo_height() / 2:
                return i
        return len(self.row_widgets)

    def _show_indicator(self, idx: int | None):
        if idx is None or not self.row_widgets:
            self.indicator.pack_forget()
            return
        if idx < len(self.row_widgets):
            self.indicator.pack(fill="x", padx=10, pady=0, before=self.row_widgets[idx])
        else:
            self.indicator.pack(fill="x", padx=10, pady=0, after=self.row_widgets[-1])

    def _autoscroll(self, y: int):
        z = self.list_zone
        top, bottom = z.winfo_rooty(), z.winfo_rooty() + z.winfo_height()
        canvas = getattr(self.scroll, "_parent_canvas", None)
        if canvas is None:
            return
        if top <= y < top + 36:
            canvas.yview_scroll(-1, "units")
        elif bottom - 36 < y <= bottom:
            canvas.yview_scroll(1, "units")

    def _on_drag_hover(self, x, y, payload):
        if self._in_zone(x, y):
            self._autoscroll(y)
            self._show_indicator(self._drop_index(y))
        else:
            self._show_indicator(None)

    def _on_drag_drop(self, x, y, payload):
        self._show_indicator(None)
        if not self._in_zone(x, y):
            return
        idx = self._drop_index(y)
        kind, data = payload
        if kind == "new":
            self._prompt_new(data, idx)
        elif kind == "move":
            if idx > data:
                idx -= 1
            if idx != data:
                self.items.insert(idx, self.items.pop(data))
                self._render_list()
        elif kind == "file":
            self._insert(ServiceItem(KIND_CUSTOM, data), idx)

    def _on_drag_click(self, payload):
        kind, data = payload
        if kind == "new":
            self._prompt_new(data, len(self.items))
        elif kind == "file":
            self._select_file(data)

    # ================================================================ 항목 조작
    def _template_names(self) -> list[str]:
        return [p.stem for p in list_ppt_files(self.cfg.resolve_path("template_dir"))]

    def _dialog(self, kind: str, item: ServiceItem | None = None) -> ServiceItem | None:
        return ItemDialog(self, kind, item, template_names=self._template_names(),
                          bible_values=self.bible_values, default_bible=self.bible_field.get().strip()).show()

    def _prompt_new(self, kind: str, idx: int):
        item = self._dialog(kind)
        if item:
            self._insert(item, idx)

    def _insert(self, item: ServiceItem, idx: int | None = None):
        idx = len(self.items) if idx is None else max(0, min(idx, len(self.items)))
        self.items.insert(idx, item)
        self._render_list()
        self._set_status(f"추가됨: {item.label} {item.value}".strip())
        if idx == len(self.items) - 1:
            self.after(30, lambda: self.scroll._parent_canvas.yview_moveto(1.0))

    def _edit_item(self, idx: int):
        if self.drag.dragging:
            return
        item = self._dialog(self.items[idx].kind, self.items[idx])
        if item:
            self.items[idx] = item
            self._render_list()

    def _move_item(self, idx: int, delta: int):
        j = idx + delta
        if 0 <= j < len(self.items):
            self.items[idx], self.items[j] = self.items[j], self.items[idx]
            self._render_list()

    def _remove_item(self, idx: int):
        removed = self.items.pop(idx)
        self._render_list()
        self._set_status(f"삭제됨: {removed.label} {removed.value}".strip())

    def _clear_items(self):
        if self.items and messagebox.askyesno("전체 비우기", "순서 목록을 모두 비울까요?", parent=self):
            self.items.clear()
            self._render_list()

    # ================================================================ 좌측: 파일 / 폴더 / 성경
    def _refresh_files(self):
        self._files = list_ppt_files(self.cfg.resolve_path("template_dir"))
        self._render_files()

    def _render_files(self):
        for w in self.file_list.winfo_children():
            w.destroy()
        self.file_buttons = {}
        q = self.file_search.get().strip().lower()
        shown = [p for p in self._files if q in p.stem.lower()]
        self.file_count.configure(text=f"{len(shown)}개")
        if not shown:
            ctk.CTkLabel(self.file_list, text="PPT 파일이 없습니다.\n템플릿 폴더에 .pptx 를 넣어 주세요." if not q
                         else "검색 결과가 없습니다.", font=font(12), text_color=TEXT2).pack(pady=20)
        for p in shown:
            b = ctk.CTkButton(self.file_list, text=f"📄  {p.stem}", anchor="w", height=30, corner_radius=6,
                              fg_color=ACCENT if p.stem == self.selected_file else "transparent",
                              hover_color=BTN_HOVER, font=font(12),
                              text_color="#FFFFFF" if p.stem == self.selected_file else TEXT)
            b.pack(fill="x", pady=1)
            self.drag.attach(b, ("file", p.stem), p.stem)
            b.bind("<Double-Button-1>", lambda e, s=p.stem: self._insert(ServiceItem(KIND_CUSTOM, s)), add="+")
            self.file_buttons[p.stem] = b

    def _select_file(self, stem: str):
        self.selected_file = stem
        for s, b in self.file_buttons.items():
            sel = s == stem
            b.configure(fg_color=ACCENT if sel else "transparent", text_color="#FFFFFF" if sel else TEXT)

    def _add_selected_file(self):
        if not self.selected_file:
            messagebox.showerror("선택 필요", "추가할 템플릿 파일을 먼저 선택해 주세요.", parent=self)
            return
        self._insert(ServiceItem(KIND_CUSTOM, self.selected_file))

    def _choose_folder(self, key: str):
        path = filedialog.askdirectory(parent=self, initialdir=self.path_vars[key].get(), mustexist=True)
        if not path:
            return
        self.cfg.set("paths", key, path)
        self.path_vars[key].set(str(self.cfg.resolve_path(key)))
        if key == "template_dir":
            self._refresh_files()

    def _open_result(self, path):
        """생성된 PPT 열기. 슬라이드 쇼 옵션이면 PowerPoint 의 Show 동작(POWERPNT.EXE /s)으로 바로 쇼 시작."""
        if self.slideshow.get():
            try:
                os.startfile(path, "show")  # .pptx 에 등록된 Show 동사 = 탐색기 우클릭 [쇼]
                return
            except OSError as e:
                exe = _powerpoint_exe()
                if exe:
                    try:
                        subprocess.Popen([str(exe), "/s", str(path)])
                        return
                    except OSError:
                        pass
                self._log(f"⚠ 슬라이드 쇼로 열지 못해 일반 열기로 대신합니다: {e}")
        try:
            os.startfile(path)
        except OSError as e:
            messagebox.showerror("열기 실패", f"생성된 파일을 열 수 없습니다.\n{path}\n\n{e}", parent=self)

    def _open_folder(self, key: str):
        path = self.cfg.resolve_path(key)
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(path)

    def _add_bible_from_form(self):
        text = self.verse_entry.get().strip()
        if not text:
            messagebox.showerror("입력 오류", "성경 구절을 입력해 주세요.", parent=self)
            return
        try:
            refs = parse_references(text)
        except ValueError as e:
            messagebox.showerror("입력 오류", str(e), parent=self)
            return
        bible = self.bible_field.get().strip()
        item = make_item(KIND_BIBLE, "말씀", ", ".join(refs), {"bible": bible} if bible else {})
        self._insert(item)
        self.verse_entry.delete(0, "end")

    def _bible2ppt(self) -> Bible2PPT:
        b = self.cfg.get("bible2ppt")
        return Bible2PPT(self.cfg.resolve_file(b.get("exe_path", "")), int(b.get("timeout_sec") or 180))

    def _bible_exe_ok(self) -> bool:
        exe = self.cfg.resolve_file(self.cfg.get("bible2ppt", "exe_path"))
        return bool(exe and exe.is_file())

    def _update_bible_status(self, note: str | None = None, error: bool = False):
        exe = self.cfg.resolve_file(self.cfg.get("bible2ppt", "exe_path"))
        if not self._bible_exe_ok():
            self.bible_status.configure(text="⚠ Bible2PPT.exe 경로가 설정되지 않았습니다 — 상단 ⚙ 설정에서 지정",
                                        text_color=DANGER)
            return
        if note is None:
            note = f"역본 {len(self.bible_values)}종" if self.bible_values else "역본 목록 없음 — ↻ 로 불러오기"
            error = not self.bible_values
        self.bible_status.configure(text=f"✔ 연결됨: {exe.name} · {note}",
                                    text_color=DANGER if error else KIND_COLORS[KIND_PRAYER])

    def _set_bible_values(self, values: list[str]):
        self.bible_values = values
        self.bible_field.set_values(values)
        cur = self.bible_field.get().strip()
        if not cur or cur not in values:  # 비었거나 예전 형식이면 설정의 기본 역본으로 맞춤
            default = self.cfg.get("bible2ppt", "bible") or ""
            if not cur or default in values:
                self.bible_field.set(default)

    def _load_bible_list_if_needed(self):
        """exe 는 있는데 캐시된 역본 목록이 없으면 (예: 첫 실행·이전 버전 설정) 조용히 불러온다."""
        if self._bible_exe_ok() and not self.bible_values:
            self._load_bible_list(silent=True)

    def _load_bible_list(self, silent: bool = False):
        if self._bible_loading:
            return
        if not self._bible_exe_ok():
            if not silent:
                messagebox.showerror("Bible2PPT", "Bible2PPT.exe 경로가 설정되지 않았습니다. ⚙ 설정에서 지정해 주세요.",
                                     parent=self)
            return
        self._bible_loading = True
        self._update_bible_status("역본 목록을 불러오는 중…")
        if not silent:
            self._set_status("성경 목록을 불러오는 중…")

        def done(res, err):
            self._bible_loading = False
            if err:
                self._update_bible_status("역본 목록을 불러오지 못함 — ↻ 로 다시 시도", error=True)
                if not silent:
                    messagebox.showerror("Bible2PPT", str(err), parent=self)
                    self._set_status("성경 목록 불러오기 실패")
                return
            self.cfg.set("bible2ppt", "bible_list", res)
            self._set_bible_values(res)
            self._update_bible_status()
            if not silent:
                self._set_status(f"성경 {len(res)}종을 불러왔습니다.")

        run_in_thread(lambda: self._bible2ppt().list_bibles(), done)

    # ================================================================ 날짜 / 설정
    def _open_calendar(self):
        try:
            initial = dt.datetime.strptime(self.date_entry.get().strip(), DATE_FMT).date()
        except ValueError:
            initial = None
        if self._calendar is not None and self._calendar.winfo_exists():
            self._calendar.destroy()  # 다시 누르면 닫기
            self._calendar = None
            return
        self._calendar = CalendarPopup(self, self.cal_btn, initial, self._set_date)

    def _set_date(self, d: dt.date):
        self.date_entry.delete(0, "end")
        self.date_entry.insert(0, d.strftime(DATE_FMT))

    def _set_today(self):
        self._set_date(dt.date.today())

    def _set_mode(self, label: str):
        mode = MODES[label]
        ctk.set_appearance_mode(mode)
        self.cfg.set("appearance_mode", None, mode)
        self.after(50, self._render_files)  # 선택 강조색 갱신

    def _open_settings(self):
        if SettingsDialog(self, self.cfg).show():
            self._set_bible_values(list(self.cfg.get("bible2ppt", "bible_list") or []))
            if not self.bible_field.get().strip():
                self.bible_field.set(self.cfg.get("bible2ppt", "bible") or "")
            self._update_bible_status()
            self._load_bible_list_if_needed()  # exe 를 바꿨는데 설정 창에서 목록을 못 받았으면 다시 시도
            self._set_status("설정을 저장했습니다.")

    def _schedule_save(self):
        if self._save_job:
            self.after_cancel(self._save_job)
        self._save_job = self.after(600, self._save_basic)

    def _save_basic(self):
        self._save_job = None
        self.cfg.set("church_name", None, self.church_entry.get().strip(), save=False)
        self.cfg.set("worship_type", None, self.worship_entry.get().strip())

    # ================================================================ 프리셋
    def _refresh_presets(self):
        self.preset_combo.configure(values=self.presets.names())

    def _save_preset(self):
        name = self.preset_combo.get().strip()
        err = self.presets.validate_name(name)
        if err:
            messagebox.showerror("프리셋 저장", err, parent=self)
            return
        if not self.items:
            messagebox.showerror("프리셋 저장", "저장할 예배 순서가 없습니다.", parent=self)
            return
        if self.presets.exists(name):
            if not messagebox.askyesno("프리셋 저장", f"'{name}' 프리셋이 이미 있습니다. 덮어쓸까요?", parent=self):
                return
        elif not self.presets.can_add(name):
            messagebox.showerror("프리셋 저장", f"프리셋은 최대 {MAX_PRESETS}개까지 저장할 수 있습니다.\n"
                                            "기존 프리셋을 삭제한 뒤 다시 시도해 주세요.", parent=self)
            return
        self.presets.upsert(name, self.items)
        self._refresh_presets()
        self.preset_combo.set(name)
        self._set_status(f"프리셋 '{name}' 저장됨 ({len(self.presets.presets)}/{MAX_PRESETS})")

    def _load_preset(self):
        name = self.preset_combo.get().strip()
        if not self.presets.exists(name):
            messagebox.showerror("프리셋 불러오기", "불러올 프리셋을 목록에서 선택해 주세요.", parent=self)
            return
        if self.items and not messagebox.askyesno("프리셋 불러오기",
                                                  f"현재 순서를 '{name}' 프리셋으로 바꿀까요?", parent=self):
            return
        self.items = self.presets.load(name)
        self.worship_entry.delete(0, "end")
        self.worship_entry.insert(0, name)
        self._save_basic()
        self._render_list()
        self._set_status(f"프리셋 '{name}' 불러옴")

    def _delete_preset(self):
        name = self.preset_combo.get().strip()
        if not self.presets.exists(name):
            messagebox.showerror("프리셋 삭제", "삭제할 프리셋을 목록에서 선택해 주세요.", parent=self)
            return
        if not messagebox.askyesno("프리셋 삭제", f"'{name}' 프리셋을 삭제할까요?\n이 작업은 되돌릴 수 없습니다.",
                                   icon="warning", parent=self):
            return
        self.presets.delete(name)
        self._refresh_presets()
        self.preset_combo.set("")
        self._set_status(f"프리셋 '{name}' 삭제됨")

    # ================================================================ 생성
    def _build(self):
        if self.building:
            return
        church = self.church_entry.get().strip()
        worship = self.worship_entry.get().strip()
        date = self.date_entry.get().strip()
        errors = []
        if not church:
            errors.append("교회 이름을 입력해 주세요.")
        try:
            dt.datetime.strptime(date, DATE_FMT)
        except ValueError:
            errors.append("날짜를 YYYY-MM-DD 형식으로 입력해 주세요.")
        if not worship:
            errors.append("예배 종류를 입력해 주세요.")
        if not self.items:
            errors.append("예배 순서가 비어 있습니다.")
        for i, it in enumerate(self.items, 1):
            err = validate_item(it)
            if err:
                errors.append(f"{i}번 항목: {err}")
        if errors:
            messagebox.showerror("생성할 수 없습니다", "\n".join(errors), parent=self)
            return
        self._save_basic()

        builder = WorshipBuilder(self.cfg, lambda m: post_to_ui(lambda: self._log(m)),
                                 lambda v, m: post_to_ui(lambda: self._set_progress(v, m)))
        _, problems = builder.resolve(self.items)
        if problems:
            messagebox.showerror("파일을 찾을 수 없습니다", "\n".join(problems), parent=self)
            return
        out = self.cfg.resolve_path("output_dir") / output_filename(date, church, worship)
        if out.exists() and not messagebox.askyesno("덮어쓰기", f"이미 같은 이름의 파일이 있습니다.\n{out.name}\n\n덮어쓸까요?",
                                                     parent=self):
            return

        self.building = True
        self.build_btn.configure(state="disabled", text="생성 중…")
        self.log_lines.clear()
        self._log(f"생성 시작: {out.name}")
        self._set_progress(0, "준비 중…")
        items = [ServiceItem.from_dict(i.to_dict()) for i in self.items]  # 작업 중 UI 변경과 분리

        def done(result, err):
            self.building = False
            self.build_btn.configure(state="normal", text="예배 PPT 병합 및 생성")
            if err:
                self._set_progress(0, "생성 실패")
                self._log(f"오류: {err}")
                messagebox.showerror("생성 실패", str(err) if isinstance(err, BuildError) else f"{type(err).__name__}: {err}",
                                     parent=self)
                return
            warns = [l for l in self.log_lines if "⚠" in l]
            self._set_progress(1, f"완료: {result.name}" + (f"  (경고 {len(warns)}건 — 로그 확인)" if warns else ""))
            if self.open_after.get():
                self._open_result(result)
            elif messagebox.askyesno("생성 완료", f"생성되었습니다.\n{result}\n\n폴더를 열까요?", parent=self):
                os.startfile(result.parent)

        run_in_thread(lambda: builder.build(BuildRequest(items, church, date, worship)), done)

    # ================================================================ 상태 / 로그
    def _set_status(self, text: str):
        self.status.configure(text=text)

    def _set_progress(self, value: float, text: str):
        self.progress.set(max(0.0, min(1.0, value)))
        self._set_status(text)

    def _log(self, msg: str):
        self.log_lines.append(f"[{dt.datetime.now():%H:%M:%S}] {msg}")
        if "⚠" in msg:
            self._set_status(msg.strip())

    def _show_log(self):
        LogDialog(self, self.log_lines).show()

    def _on_close(self):
        if self.building and not messagebox.askyesno("종료", "PPT 생성 중입니다. 그래도 종료할까요?", parent=self):
            return
        self._save_basic()
        self.destroy()


if __name__ == "__main__":
    _set_app_user_model_id()
    App().mainloop()
