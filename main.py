"""예배 순서 PPT 제작 프로그램 — 메인 윈도우. (버전은 version.py)

레이아웃은 UI 초안 기준 (CLAUDE.md §4 참고).
"""
from __future__ import annotations

import datetime as dt
import math
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from bible_bridge import Bible2PPT, Bible2PPTError, parse_references
from config_manager import MAX_PRESETS, ConfigManager, PresetManager, save_worship_to_text, load_worship_from_text
from models import (DEFAULT_FILENAME_FORMAT, FILENAME_FIELDS, KIND_BIBLE, KIND_CCM, KIND_CUSTOM, KIND_HYMN, KIND_LABELS, KIND_PRAYER, KIND_SERMON, ServiceItem, make_item,
                    validate_item)
from ui_widgets import (ACCENT, BG, BORDER, BTN_HOVER, CARD, DANGER, KIND_COLORS, PILL_BORDER, SUBTLE, TEXT,
                        TEXT2, BibleField, CalendarPopup, DragController, ItemDialog, LogDialog, SettingsDialog,
                        accent_button, card, font, mode_color, post_to_ui, pump_ui_queue, run_in_thread, scaled,
                        secondary_button, tk_font)
from version import APP_NAME, __version__
from worship_builder import (BuildError, BuildRequest, WorshipBuilder, list_ccm_files, list_ppt_files, output_filename,
                             prayer_template_names, sermon_template, validate_filename_format)

APP_TITLE = f"{APP_NAME} v{__version__}"
DATE_FMT = "%Y-%m-%d"
MODES = {"라이트": "Light", "다크": "Dark", "시스템": "System"}
MIN_FILE_ROWS = 4  # 최소 창 크기에서도 템플릿 파일 목록이 이만큼은 보이도록
FILE_FONT_SIZE = 13.2  # 템플릿 파일 목록 글자 크기 (기본 12 의 110% — 항목 줄 높이도 글자 크기를 따라 커짐)
RESIZE_SETTLE_MS = 150  # 창 크기 조절이 이 시간 동안 멈추면 그때 내용 배치를 새 크기로 맞춤
# PyInstaller onefile 은 번들 파일을 sys._MEIPASS 에 풀어 둠
ICON_PATH = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / "assets" / "app.ico"


def _enable_per_monitor_dpi_v2():
    """Windows 배율이 바뀔 때 창 틀(제목 표시줄·테두리)까지 다시 그려지도록 Per-Monitor V2 DPI 모드를 켠다 (Tk 생성 전).
    CTk 는 V1(SetProcessDpiAwareness(2))을 쓰는데, V1 에선 창 틀을 Windows 가 다시 맞춰 주지 않아 옛 배율 크기로 남는다.
    먼저 V2 로 정해 두면 CTk 의 V1 호출은 '이미 설정됨'으로 무시되고, CTk 의 DPI 감지(GetDpiForMonitor)는 그대로 동작한다.
    V2 를 지원하지 않는 Windows(10 1703 이전)면 아무것도 하지 않음 → CTk 의 V1 로 동작."""
    try:
        import ctypes
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))  # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    except (AttributeError, OSError):
        pass


def _set_app_user_model_id():
    """작업 표시줄이 python.exe 아이콘 대신 창 아이콘을 쓰도록 별도 앱으로 등록 (Tk 생성 전 호출)."""
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ChurchPPTMaker.App")
    except (AttributeError, OSError):
        pass


def _disable_full_redraw_on_resize(win: tk.Misc):
    """Tk 창 클래스(TkTopLevel/TkChild)의 CS_HREDRAW·CS_VREDRAW 를 끈다.
    켜져 있으면 창 크기가 1px 바뀔 때마다 Windows 가 모든 하위 위젯(수백 개 HWND)을 통째로 다시 그려
    드래그 중 한 단계에 0.15초 이상 걸렸다(실측 152ms → 6ms). Tk 위젯은 자기 크기가 바뀌면 스스로 다시 그리므로 불필요.
    클래스 단위 설정이라 이 프로세스의 모든 Tk 창에 적용된다."""
    try:
        import ctypes
        from ctypes import wintypes
        u = ctypes.windll.user32
        u.GetParent.restype = wintypes.HWND
        u.GetParent.argtypes = [wintypes.HWND]
        u.GetClassLongPtrW.restype = ctypes.c_size_t
        u.GetClassLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
        u.SetClassLongPtrW.restype = ctypes.c_size_t
        u.SetClassLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_size_t]
        gcl_style, cs_redraw = -26, 0x0001 | 0x0002  # GCL_STYLE, CS_VREDRAW | CS_HREDRAW
        inner = win.winfo_id()                       # TkChild
        for hwnd in (u.GetParent(inner), inner):     # TkTopLevel(바깥 틀), TkChild
            style = u.GetClassLongPtrW(hwnd, gcl_style)
            if style & cs_redraw:
                u.SetClassLongPtrW(hwnd, gcl_style, style & ~cs_redraw)
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
        _enable_per_monitor_dpi_v2()
        self.cfg = ConfigManager()
        self.presets = PresetManager()
        ctk.set_appearance_mode(self.cfg.get("appearance_mode") or "System")
        ctk.set_default_color_theme("blue")
        super().__init__(fg_color=BG)
        self.title(APP_TITLE)
        if ICON_PATH.is_file():
            self.iconbitmap(default=str(ICON_PATH))  # CTk 기본 아이콘 대신 (default= 는 다이얼로그 등에도 적용)
        self._fit_to_screen(1220, 820)
        _disable_full_redraw_on_resize(self)

        self.items: list[ServiceItem] = []
        self.path_vars: dict[str, tk.StringVar] = {}  # 폴더 키 → 화면에 표시하는 경로 (_choose_folder 가 갱신)
        self.row_widgets: list[tk.Frame] = []
        self._shown_files: list[str] = []  # 템플릿 목록 상자에 보이는 순서대로의 파일 이름
        self.selected_file: str | None = None
        self.bible_values: list[str] = list(self.cfg.get("bible2ppt", "bible_list") or [])
        self._bible_loading = False
        self.log_lines: list[str] = []
        self.building = False
        self._save_job = None
        self._calendar: CalendarPopup | None = None

        self.drag = DragController(self, self._on_drag_hover, self._on_drag_drop, self._on_drag_click)

        # 창 크기 조절 렉 방지: 모든 내용을 컨테이너 하나에 담고, 크기 조절이 잠시 멈췄을 때만 새 크기로 맞춘다.
        # (Windows 에서는 위젯 수백 개를 매 Configure 마다 재배치·재그리기 하느라 한 번에 0.5초 가까이 걸렸음)
        self._content = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        tk.Place.place_configure(self._content, x=0, y=0, relwidth=1, relheight=1)  # CTk place 는 width/height 금지
        self._content_size: tuple[int, int] | None = None
        self._resize_job = None
        self.bind("<Configure>", self._on_window_configure, add="+")
        self._content.grid_columnconfigure(0, weight=1)
        self._content.grid_rowconfigure(2, weight=1)
        self._build_header()
        self._build_info_bar()
        self._build_body()
        self._build_bottom_bar()

        self._refresh_files()
        self._refresh_presets()
        self._render_list()
        self._update_bible_status()
        self._load_bible_list_if_needed()
        ctk.AppearanceModeTracker.add(self._on_appearance_change)  # '시스템' 모드에서 Windows 테마가 바뀔 때도 호출됨
        ctk.ScalingTracker.add_widget(self._on_scaling_change, self)  # Windows 배율 변경·다른 배율 모니터로 이동
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        pump_ui_queue(self)

    def _on_window_configure(self, e):
        if e.widget is not self or (e.width, e.height) == self._content_size:
            return
        if self._content_size is None:  # 처음 표시될 때는 바로 맞추고, 실제 배치를 보고 최소 높이 결정
            self._fit_content()
            self.after(100, self._apply_min_size)
            return
        if self._resize_job is not None:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(RESIZE_SETTLE_MS, self._fit_content)

    def _fit_content(self):
        """컨테이너를 현재 창 크기(실제 픽셀)로 고정 → 내용 재배치는 여기서 한 번만 일어난다."""
        self._resize_job = None
        self._content_size = (self.winfo_width(), self.winfo_height())
        tk.Place.place_configure(self._content, relwidth="", relheight="",
                                 width=self._content_size[0], height=self._content_size[1])

    def _work_area(self) -> tuple[int, int, int, int]:
        """창이 있는(아직 없으면 가장 가까운) 모니터의 작업 영역 (왼, 위, 오른, 아래) — 실제 픽셀, 작업 표시줄 제외.
        winfo_screenwidth 는 DPI 모드에 따라 논리/실제 단위가 달라서(V1: 논리, V2: 실제) 쓰지 않는다."""
        try:
            import ctypes
            from ctypes import wintypes

            class MONITORINFO(ctypes.Structure):
                _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT),
                            ("rcWork", wintypes.RECT), ("dwFlags", wintypes.DWORD)]

            u = ctypes.windll.user32
            u.MonitorFromWindow.restype = wintypes.HMONITOR
            u.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
            info = MONITORINFO(cbSize=ctypes.sizeof(MONITORINFO))
            if u.GetMonitorInfoW(u.MonitorFromWindow(self.winfo_id(), 2), ctypes.byref(info)):  # MONITOR_DEFAULTTONEAREST
                r = info.rcWork
                return r.left, r.top, r.right, r.bottom
        except (AttributeError, OSError):
            pass
        return 0, 0, self.winfo_screenwidth(), self.winfo_screenheight()

    def _available_size(self) -> tuple[int, int]:
        """창이 쓸 수 있는 최대 크기 (CTk 논리 단위, 제목 표시줄·여유 제외)."""
        left, top, right, bottom = self._work_area()
        scale = ctk.ScalingTracker.get_window_scaling(self)
        return int((right - left) / scale) - 20, int((bottom - top) / scale) - 50

    def _fit_to_screen(self, w: int, h: int):
        """화면 배율(DPI)을 고려해 창이 화면 밖으로 나가지 않게 크기/위치를 정한다."""
        self.update_idletasks()
        aw, ah = self._available_size()
        w, h = min(w, aw), min(h, ah)
        left, top, right, bottom = self._work_area()
        scale = ctk.ScalingTracker.get_window_scaling(self)  # 크기는 논리 단위, 위치는 실제 픽셀
        x = left + max((right - left - int(w * scale)) // 2, 0)
        y = top + max((bottom - top - int(h * scale)) // 3, 0)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(min(1000, w), min(660, h))

    def _apply_min_size(self):
        """최소 높이를 '템플릿 목록이 MIN_FILE_ROWS 줄 보이는 높이'로 맞춘다.
        고정값 660 은 왼쪽 패널보다 낮아서 최소 크기에서 템플릿 목록이 0줄로 눌려 사라졌다.
        창 높이가 줄면 왼쪽에서는 템플릿 목록만 줄어드므로 (현재 창 높이 - 현재 목록 높이 + 목록 MIN_FILE_ROWS 줄 높이)가
        필요한 최소 높이다. CTk 위젯 요청 크기(reqheight)는 실제보다 커서 실제 배치로 계산. 화면보다 크게 잡지는 않음."""
        self.update_idletasks()
        lb = self.file_listbox
        if not lb.winfo_ismapped():
            return
        scale = ctk.ScalingTracker.get_window_scaling(self)
        need = math.ceil((self.winfo_height() - lb.winfo_height() + lb.winfo_reqheight()) / scale)  # 픽셀 → 논리 단위
        screen = self._available_size()[1]  # _fit_to_screen 과 같은 기준 (논리 단위)
        self.minsize(round(self.wm_minsize()[0] / scale), max(660, min(need, screen)))

    # ================================================================ 상단
    def _build_header(self):
        bar = ctk.CTkFrame(self._content, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 8))
        ctk.CTkLabel(bar, text=APP_TITLE, font=font(20, "bold"), text_color=TEXT).pack(side="left")
        secondary_button(bar, "⚙  설정", self._open_settings, width=88).pack(side="right")
        current = {v: k for k, v in MODES.items()}.get(self.cfg.get("appearance_mode"), "시스템")
        self.mode_seg = ctk.CTkSegmentedButton(bar, values=list(MODES), command=self._set_mode, font=font(12),
                                               selected_color=ACCENT, height=32)
        self.mode_seg.set(current)
        self.mode_seg.pack(side="right", padx=10)

    def _build_info_bar(self):
        bar = card(self._content)
        bar.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        # 줄 틀은 카드 색으로 칠해지므로 둥근 테두리를 덮지 않도록 안쪽으로 들여 배치
        line1 = ctk.CTkFrame(bar, fg_color="transparent")
        line1.pack(fill="x", padx=10, pady=(8, 0))
        line2 = ctk.CTkFrame(bar, fg_color="transparent")
        line2.pack(fill="x", padx=10, pady=(0, 8))

        def label(master, text):
            return ctk.CTkLabel(master, text=text, font=font(13, "bold"), text_color=TEXT)

        def on_change(_=None):
            self._update_filename_preview()
            self._schedule_save()

        label(line1, "교회 이름").pack(side="left", padx=(8, 8), pady=(6, 5))
        self.church_entry = ctk.CTkEntry(line1, width=230, height=34, font=font(13), placeholder_text="예: 사랑의교회")
        self.church_entry.insert(0, self.cfg.get("church_name") or "")
        self.church_entry.pack(side="left")
        self.church_entry.bind("<KeyRelease>", on_change)

        label(line1, "날짜").pack(side="left", padx=(28, 8))
        self.date_entry = ctk.CTkEntry(line1, width=130, height=34, font=font(13), placeholder_text="YYYY-MM-DD")
        self.date_entry.insert(0, dt.date.today().strftime(DATE_FMT))
        self.date_entry.pack(side="left")
        self.date_entry.bind("<KeyRelease>", lambda e: self._update_filename_preview())
        self.cal_btn = secondary_button(line1, "📅", self._open_calendar, width=36)
        self.cal_btn.configure(height=34)
        self.cal_btn.pack(side="left", padx=(4, 4))
        secondary_button(line1, "금일 지정", self._set_today, width=80).pack(side="left")

        label(line1, "예배 종류").pack(side="left", padx=(28, 8))
        self.worship_entry = ctk.CTkEntry(line1, width=170, height=34, font=font(13), placeholder_text="예: 주일오전예배")
        self.worship_entry.insert(0, self.cfg.get("worship_type") or "")
        self.worship_entry.pack(side="left", padx=(0, 8))
        self.worship_entry.bind("<KeyRelease>", on_change)

        # 결과 파일 이름 형식: {항목} 자리에 값이 들어감. 프리셋을 저장하면 함께 저장되고, 불러오면 프리셋의 형식으로 바뀜
        # 출력 폴더: 결과 파일이 저장되는 곳이라 파일 이름 형식 바로 왼쪽에 둔다
        label(line2, "출력 폴더").pack(side="left", padx=(8, 8), pady=(5, 6))
        var = tk.StringVar(value=str(self.cfg.resolve_path("output_dir")))
        self.path_vars["output_dir"] = var
        ctk.CTkEntry(line2, textvariable=var, state="readonly", width=180, height=34, font=font(12),
                     text_color=TEXT2).pack(side="left")
        for text, cmd in (("📁", lambda: self._choose_folder("output_dir")), ("↗", lambda: self._open_folder("output_dir"))):
            b = secondary_button(line2, text, cmd, width=34)
            b.configure(height=34)
            b.pack(side="left", padx=(4, 0))

        label(line2, "파일 이름").pack(side="left", padx=(20, 8))
        self.format_entry = ctk.CTkEntry(line2, width=230, height=34, font=font(13),
                                         placeholder_text=DEFAULT_FILENAME_FORMAT)
        self.format_entry.insert(0, self.cfg.get("filename_format") or DEFAULT_FILENAME_FORMAT)
        self.format_entry.pack(side="left")
        self.format_entry.bind("<KeyRelease>", on_change)
        self.format_menu_btn = secondary_button(line2, "항목 ▾", self._open_format_menu, width=64)
        self.format_menu_btn.configure(height=34)
        self.format_menu_btn.pack(side="left", padx=(4, 4))
        b = secondary_button(line2, "기본값", lambda: self._set_filename_format(DEFAULT_FILENAME_FORMAT), width=64)
        b.configure(height=34)
        b.pack(side="left")
        self.filename_preview = ctk.CTkLabel(line2, text="", font=font(12), text_color=TEXT2, anchor="w")
        self.filename_preview.pack(side="left", fill="x", expand=True, padx=(12, 8))
        self._update_filename_preview()

    def _open_format_menu(self):
        """[항목 ▾] → 커서 위치에 {항목} 넣기 (사용할 수 있는 항목 안내도 겸함)."""
        menu = tk.Menu(self, tearoff=0, font=tk_font(self, 12))
        for f in FILENAME_FIELDS:
            menu.add_command(label=f"{{{f}}}", command=lambda f=f: self._insert_format_field(f))
        btn = self.format_menu_btn
        menu.tk_popup(btn.winfo_rootx(), btn.winfo_rooty() + btn.winfo_height())

    def _insert_format_field(self, field: str):
        self.format_entry.insert("insert", f"{{{field}}}")
        self.format_entry.focus_set()
        self._update_filename_preview()
        self._schedule_save()

    def _set_filename_format(self, fmt: str):
        self.format_entry.delete(0, "end")
        self.format_entry.insert(0, fmt)
        self._update_filename_preview()
        self._schedule_save()

    def _update_filename_preview(self):
        fmt = self.format_entry.get().strip()
        err = validate_filename_format(fmt)
        if err:
            self.filename_preview.configure(text="⚠ " + err.splitlines()[0], text_color=DANGER)
            return
        name = output_filename(self.date_entry.get().strip(), self.church_entry.get().strip(),
                               self.worship_entry.get().strip(), fmt)
        self.filename_preview.configure(text=f"→  {name}", text_color=TEXT2)

    # ================================================================ 본문
    def _build_body(self):
        body = ctk.CTkFrame(self._content, fg_color="transparent")
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
        # 템플릿 폴더는 [템플릿 파일] 카드 맨 위, 출력 폴더는 상단 파일 이름 줄에 둔다
        folders = [("prayer_dir", "기도"), ("hymn_dir", "찬송가"), ("ccm_dir", "CCM")]
        for r, (key, name) in enumerate(folders, 1):
            self._folder_row(fc, r, key, name)
        ctk.CTkFrame(fc, width=1, height=6, fg_color="transparent").grid(row=len(folders) + 1, column=0)

        # 템플릿 파일 목록
        tc = card(left)
        tc.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        tc.grid_columnconfigure(0, weight=1)
        tc.grid_rowconfigure(3, weight=1)
        head = ctk.CTkFrame(tc, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 4))
        ctk.CTkLabel(head, text="템플릿 파일", font=font(15, "bold"), text_color=TEXT).pack(side="left")
        self.file_count = ctk.CTkLabel(head, text="", font=font(12), text_color=TEXT2)
        self.file_count.pack(side="left", padx=8)
        secondary_button(head, "↻", self._refresh_files, width=32).pack(side="right")
        frow = ctk.CTkFrame(tc, fg_color="transparent")
        frow.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        frow.grid_columnconfigure(1, weight=1)
        self._folder_row(frow, 0, "template_dir", "폴더")
        self.file_search = ctk.CTkEntry(tc, height=30, font=font(12), placeholder_text="🔍  파일 이름 검색")
        self.file_search.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 6))
        self.file_search.bind("<KeyRelease>", lambda e: self._render_files())
        # 파일마다 CTkButton 을 만들면 파일 수만큼 위젯(HWND)이 늘어 창 크기 조절이 느려진다 → tk.Listbox 하나로 그림
        box = ctk.CTkFrame(tc, fg_color=SUBTLE, corner_radius=8)
        box.grid(row=3, column=0, sticky="nsew", padx=14)
        self.file_listbox = tk.Listbox(box, height=MIN_FILE_ROWS, activestyle="none", borderwidth=0, highlightthickness=0,
                                       exportselection=False, selectmode="browse", font=tk_font(self, FILE_FONT_SIZE),
                                       cursor="hand2")
        sb = ctk.CTkScrollbar(box, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=sb.set)
        self.file_listbox.pack(side="left", fill="both", expand=True, padx=(scaled(self, 8), 0), pady=scaled(self, 6))
        sb.pack(side="right", fill="y", padx=2, pady=4)
        self.file_empty = ctk.CTkLabel(box, text="", font=font(12), text_color=TEXT2, fg_color=SUBTLE)
        self.drag.attach(self.file_listbox, self._file_at)
        # 기본 동작(누르기/끌기 = 커서 아래 항목 선택)은 끈다 → 선택 표시는 _select_file 만 바꿈 (빈 곳 클릭 시 엉뚱한 선택 방지)
        for seq in ("<ButtonPress-1>", "<B1-Motion>"):
            self.file_listbox.bind(seq, lambda e: "break", add="+")
        self.file_listbox.bind("<Double-Button-1>", self._on_file_double_click)
        self._style_file_listbox()
        foot = ctk.CTkFrame(tc, fg_color="transparent")
        foot.grid(row=4, column=0, sticky="ew", padx=14, pady=(6, 12))
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

    def _folder_row(self, parent, r: int, key: str, name: str):
        """[이름] [경로(읽기 전용)] [📁 지정] [↗ 열기] 한 줄 (parent 의 grid r 행, 1열이 늘어남)."""
        ctk.CTkLabel(parent, text=name, font=font(12), text_color=TEXT2, width=44, anchor="w").grid(
            row=r, column=0, padx=(14, 4), pady=3, sticky="w")
        var = tk.StringVar(value=str(self.cfg.resolve_path(key)))
        self.path_vars[key] = var
        ctk.CTkEntry(parent, textvariable=var, state="readonly", height=30, font=font(11),
                     text_color=TEXT2).grid(row=r, column=1, sticky="ew", pady=3)
        secondary_button(parent, "📁", lambda: self._choose_folder(key), width=32).grid(
            row=r, column=2, padx=(4, 0), pady=3)
        secondary_button(parent, "↗", lambda: self._open_folder(key), width=32).grid(
            row=r, column=3, padx=(4, 14), pady=3)

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
        # 파일 입출력: 프리셋 행과 헷갈리지 않도록 제목 줄에 둠 → [파일 불러오기] [파일 내보내기] [전체 비우기]
        for text, cmd in (("파일 내보내기", self._export_worship_file), ("파일 불러오기", self._import_worship_file)):
            b = secondary_button(head, text, cmd, width=96)
            b.configure(height=28, font=font(12))
            b.pack(side="right", padx=(0, 6))

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

        # 팔레트: [기도] [찬송] [말씀] [설교] [CCM] [사용자 지정] — 드래그&드롭
        # 버튼 폭 합계가 최소 창 폭의 순서 패널(약 570)을 넘지 않도록 할 것
        pal = ctk.CTkFrame(oc, fg_color="transparent")
        pal.grid(row=3, column=0, sticky="ew", padx=16, pady=(10, 4))
        for kind in (KIND_PRAYER, KIND_HYMN, KIND_BIBLE, KIND_SERMON, KIND_CCM, KIND_CUSTOM):
            b = ctk.CTkButton(pal, text=KIND_LABELS[kind], width=78 if kind != KIND_CUSTOM else 108, height=38,
                              corner_radius=8, fg_color=CARD, hover_color=BTN_HOVER, border_width=2,
                              border_color=KIND_COLORS[kind], text_color=TEXT, font=font(14, "bold"))
            b.pack(side="left", padx=(0, 8))
            b.configure(cursor="hand2")
            self.drag.attach(b, ("new", kind), KIND_LABELS[kind])
        ctk.CTkLabel(oc, text="버튼을 위 순서 목록으로 끌어다 놓으면 해당 위치에 추가됩니다 (클릭하면 맨 아래에 추가)",
                     font=font(11), text_color=TEXT2).grid(row=4, column=0, sticky="w", padx=16, pady=(0, 12))

    # ================================================================ 하단
    def _build_bottom_bar(self):
        bar = card(self._content)
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
            ctk.CTkLabel(self.scroll, text="아래 [기도] [찬송] [말씀] [설교] [CCM] [사용자 지정] 버튼을 끌어다 놓아\n예배 순서를 추가하세요.",
                         font=font(13), text_color=TEXT2, justify="center").pack(pady=60)
        for idx, item in enumerate(self.items):
            self.row_widgets.append(self._make_row(idx, item))
        self.count_label.configure(text=f"{len(self.items)}개 항목")

    def _make_row(self, idx: int, item: ServiceItem) -> tk.Frame:
        # Windows 에선 위젯(HWND) 수가 곧 창 크기 조절·목록 갱신 비용이라(CLAUDE.md §11) 둥근 알약 2개만 CTk 로 그리고
        # 나머지(행 틀·손잡이·연결선·▲▼✕)는 일반 tk 위젯으로 그린다. 색은 모드가 바뀌면 _on_appearance_change 가 다시 그림.
        bg = mode_color(SUBTLE)
        row = tk.Frame(self.scroll, bg=bg)
        row.pack(fill="x", padx=scaled(self, 6), pady=scaled(self, 3))

        handle = tk.Label(row, text="⋮⋮", bg=bg, fg=mode_color(TEXT2), font=tk_font(self, 14, "bold"),
                          width=2, cursor="fleur")
        handle.pack(side="left", padx=(scaled(self, 2), scaled(self, 6)))
        self.drag.attach(handle, ("move", idx), f"{item.label}  {item.display_value}".strip())

        left = ctk.CTkButton(row, text=item.label, width=140, height=36, corner_radius=10, fg_color=CARD, bg_color=SUBTLE,
                             hover_color=BTN_HOVER, border_width=2, border_color=KIND_COLORS[item.kind],
                             text_color=TEXT, font=font(13, "bold"), command=lambda i=idx: self._edit_item(i))
        left.pack(side="left")
        tk.Frame(row, height=max(1, scaled(self, 1.5)), bg=mode_color(PILL_BORDER)).pack(side="left", fill="x", expand=True)
        right = ctk.CTkButton(row, text=item.display_value, width=170, height=36, corner_radius=10, fg_color=CARD,
                              bg_color=SUBTLE, hover_color=BTN_HOVER, border_width=1, border_color=PILL_BORDER,
                              text_color=TEXT if item.display_value != "—" else TEXT2, font=font(13),
                              command=lambda i=idx: self._edit_item(i))
        right.pack(side="left")

        tools = tk.Frame(row, bg=bg)
        tools.pack(side="left", padx=(scaled(self, 8), 0))
        hover = mode_color(BTN_HOVER)
        for text, cmd, color in (("▲", lambda i=idx: self._move_item(i, -1), TEXT2),
                                 ("▼", lambda i=idx: self._move_item(i, 1), TEXT2),
                                 ("✕", lambda i=idx: self._remove_item(i), DANGER)):
            b = tk.Label(tools, text=text, bg=bg, fg=mode_color(color), font=tk_font(self, 12), cursor="hand2",
                         padx=scaled(self, 7), pady=scaled(self, 4))
            b.pack(side="left")
            b.bind("<Enter>", lambda e, w=b: w.configure(bg=hover))
            b.bind("<Leave>", lambda e, w=b: w.configure(bg=bg))
            # 클릭 즉시 목록을 다시 그리면 이벤트 처리 중인 위젯이 사라지므로 다음 루프에서 실행
            b.bind("<ButtonRelease-1>", lambda e, c=cmd: self.after_idle(c))
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
        ccm = list_ccm_files(self.cfg.resolve_path("ccm_dir")) if kind == KIND_CCM else None
        prayer = prayer_template_names(self.cfg) if kind == KIND_PRAYER else None
        return ItemDialog(self, kind, item, template_names=self._template_names(), bible_values=self.bible_values,
                          default_bible=self.bible_field.get().strip(), ccm_files=ccm, prayer_names=prayer).show()

    def _can_add(self, kind: str) -> bool:
        """새 항목을 추가하기 전에 필요한 설정이 됐는지 확인 (안 됐으면 에러 창)."""
        if kind == KIND_PRAYER:
            names = prayer_template_names(self.cfg)
            if names is None:
                messagebox.showerror("기도 폴더 없음", "기도 템플릿 폴더를 찾을 수 없습니다.\n"
                                     "왼쪽 [폴더]의 '기도' 줄에서 📁 로 폴더를 지정해 주세요.", parent=self)
                return False
            if not names:
                messagebox.showerror("기도 템플릿 없음", "기도 폴더에 PPT 파일이 없습니다.\n"
                                     "합심기도.pptx, 중보기도.pptx 처럼 기도 종류별 PPT 를 넣거나\n"
                                     "왼쪽 [폴더]의 '기도' 줄에서 📁 로 다른 폴더를 지정해 주세요.", parent=self)
                return False
        if kind == KIND_BIBLE:
            try:
                self._bible2ppt().check()  # exe 미설정·없음·.exe 아님·명령줄 미지원 버전
            except Bible2PPTError as e:
                msg = f"말씀을 추가하려면 Bible2PPT 가 필요합니다.\n\n{e}"
                if "⚙ 설정" not in str(e):  # 미설정 메시지에는 이미 안내가 들어 있음
                    msg += "\n\n상단 ⚙ 설정 → Bible2PPT 연동에서 Bible2PPT.exe 를 지정해 주세요."
                messagebox.showerror("Bible2PPT 설정 필요", msg, parent=self)
                return False
        if kind == KIND_SERMON and sermon_template(self.cfg) is None:
            messagebox.showerror("설교 템플릿 설정 필요", "설교 템플릿 PPT 가 설정되지 않았거나 찾을 수 없습니다.\n"
                                 "상단 ⚙ 설정 → 설교 템플릿에서 설교 템플릿 PPT 를 지정해 주세요.", parent=self)
            return False
        return True

    def _prompt_new(self, kind: str, idx: int):
        if not self._can_add(kind):
            return
        item = self._dialog(kind)
        if item:
            self._insert(item, idx)

    def _insert(self, item: ServiceItem, idx: int | None = None):
        idx = len(self.items) if idx is None else max(0, min(idx, len(self.items)))
        self.items.insert(idx, item)
        self._render_list()
        self._set_status(f"추가됨: {item.label} {item.display_value}".strip())
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
        self._set_status(f"삭제됨: {removed.label} {removed.display_value}".strip())

    def _clear_items(self):
        if self.items and messagebox.askyesno("전체 비우기", "순서 목록을 모두 비울까요?", parent=self):
            self.items.clear()
            self._render_list()

    # ================================================================ 좌측: 파일 / 폴더 / 성경
    def _refresh_files(self):
        self._files = list_ppt_files(self.cfg.resolve_path("template_dir"))
        self._render_files()

    def _render_files(self):
        q = self.file_search.get().strip().lower()
        self._shown_files = [p.stem for p in self._files if q in p.stem.lower()]
        self.file_count.configure(text=f"{len(self._shown_files)}개")
        lb = self.file_listbox
        lb.delete(0, "end")
        for stem in self._shown_files:
            lb.insert("end", f"📄  {stem}")
        if self.selected_file in self._shown_files:
            lb.selection_set(self._shown_files.index(self.selected_file))
        if self._shown_files:
            self.file_empty.place_forget()
        else:
            self.file_empty.configure(text="PPT 파일이 없습니다.\n템플릿 폴더에 .pptx 를 넣어 주세요." if not q
                                      else "검색 결과가 없습니다.")
            self.file_empty.place(relx=0.5, rely=0.5, anchor="center")

    def _style_file_listbox(self):
        self.file_listbox.configure(bg=mode_color(SUBTLE), fg=mode_color(TEXT), selectbackground=mode_color(ACCENT),
                                    selectforeground="#FFFFFF")

    def _file_at(self, e) -> tuple[tuple[str, str], str] | None:
        """목록 상자에서 누른 위치의 파일 → 드래그 payload (빈 공간이면 None)."""
        lb = self.file_listbox
        i = lb.nearest(e.y)
        box = lb.bbox(i) if i >= 0 else None
        if not box or not box[1] <= e.y < box[1] + box[3]:
            return None
        stem = self._shown_files[i]
        return ("file", stem), stem

    def _on_file_double_click(self, e):
        if (got := self._file_at(e)) is not None:
            self._insert(ServiceItem(KIND_CUSTOM, got[1]))

    def _select_file(self, stem: str):
        self.selected_file = stem
        lb = self.file_listbox
        lb.selection_clear(0, "end")
        if stem in self._shown_files:
            lb.selection_set(self._shown_files.index(stem))

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
        if not self._can_add(KIND_BIBLE):
            return
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
        return Bible2PPT(self.cfg.resolve_file(b.get("exe_path", "")), b.get("timeout_sec"))

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
        self._update_filename_preview()

    def _set_today(self):
        self._set_date(dt.date.today())

    def _on_appearance_change(self, _mode: str):
        """일반 tk 위젯(순서 행·템플릿 목록)은 CTk 처럼 색이 자동으로 바뀌지 않으므로 다시 칠한다.
        CTk 가 콜백 목록을 도는 중이라 위젯 파괴(콜백 제거)는 다음 루프에서."""
        self.after_idle(lambda: (self._style_file_listbox(), self._render_list()))

    def _on_scaling_change(self, _widget_scaling: float, _window_scaling: float):
        """일반 tk 위젯(순서 행·템플릿 목록)의 글꼴·여백은 만들 때의 배율로 고정되므로 새 배율로 다시 만든다.
        CTk 위젯은 CTk 가 알아서 다시 그림. 콜백 순회 중이므로 위젯 파괴는 다음 루프에서."""
        self.after_idle(self._rescale_plain_widgets)

    def _rescale_plain_widgets(self):
        lb = self.file_listbox
        lb.configure(font=tk_font(self, FILE_FONT_SIZE))
        lb.pack_configure(padx=(scaled(self, 8), 0), pady=scaled(self, 6))
        self._render_list()

    def _set_mode(self, label: str):
        mode = MODES[label]
        ctk.set_appearance_mode(mode)
        self.cfg.set("appearance_mode", None, mode)

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
        self.cfg.set("worship_type", None, self.worship_entry.get().strip(), save=False)
        self.cfg.set("filename_format", None, self.format_entry.get().strip())

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
        fmt = self.format_entry.get().strip()
        err = validate_filename_format(fmt)
        if err:
            messagebox.showerror("프리셋 저장", err, parent=self)
            return
        if self.presets.exists(name):
            if not messagebox.askyesno("프리셋 저장", f"'{name}' 프리셋이 이미 있습니다. 덮어쓸까요?", parent=self):
                return
        elif not self.presets.can_add(name):
            messagebox.showerror("프리셋 저장", f"프리셋은 최대 {MAX_PRESETS}개까지 저장할 수 있습니다.\n"
                                            "기존 프리셋을 삭제한 뒤 다시 시도해 주세요.", parent=self)
            return
        self.presets.upsert(name, self.items, fmt)
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
        self._set_filename_format(self.presets.filename_format(name))
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

    def _import_worship_file(self):
        """외부 텍스트 파일에서 예배 순서를 불러옴."""
        path = filedialog.askopenfilename(
            parent=self,
            title="예배 순서 파일 선택",
            filetypes=[("텍스트 파일", "*.txt"), ("모든 파일", "*.*")]
        )
        if not path:
            return

        try:
            items, meta = load_worship_from_text(Path(path))
            if not items:
                messagebox.showwarning("불러오기 실패", "유효한 예배 순서 항목이 파일에 없습니다.", parent=self)
                return

            if self.items and not messagebox.askyesno("불러오기", "현재 순서를 파일 내용으로 대체할까요?\n'아니오'를 누르면 아래에 추가됩니다.", parent=self):
                self.items = self.items + items
            else:
                self.items = items

            # 메타 정보 반영
            if "교회이름" in meta:
                self.church_entry.delete(0, "end")
                self.church_entry.insert(0, meta["교회이름"])
            if "예배종류" in meta:
                self.worship_entry.delete(0, "end")
                self.worship_entry.insert(0, meta["예배종류"])
            if "날짜" in meta:
                self.date_entry.delete(0, "end")
                self.date_entry.insert(0, meta["날짜"])
            if meta.get("파일이름형식") and not validate_filename_format(meta["파일이름형식"]):
                self.format_entry.delete(0, "end")
                self.format_entry.insert(0, meta["파일이름형식"])
            self._update_filename_preview()

            self._save_basic()
            self._render_list()
            self._set_status(f"파일에서 {len(items)}개 항목을 불러왔습니다.")

        except Exception as e:
            messagebox.showerror("불러오기 오류", f"파일을 읽는 중 오류가 발생했습니다:\n{e}", parent=self)

    def _export_worship_file(self):
        """현재 예배 순서를 텍스트 파일로 내보냄."""
        if not self.items:
            messagebox.showerror("내보내기 실패", "내보낼 예배 순서가 없습니다.", parent=self)
            return

        path = filedialog.asksaveasfilename(
            parent=self,
            title="예배 순서 내보내기",
            defaultextension=".txt",
            filetypes=[("텍스트 파일", "*.txt"), ("모든 파일", "*.*")]
        )
        if not path:
            return

        try:
            meta = {
                "church_name": self.church_entry.get().strip(),
                "worship_type": self.worship_entry.get().strip(),
                "date": self.date_entry.get().strip(),
                "filename_format": self.format_entry.get().strip(),
            }
            save_worship_to_text(Path(path), self.items, meta)
            self._set_status(f"파일로 저장되었습니다: {Path(path).name}")
        except Exception as e:
            messagebox.showerror("내보내기 오류", f"파일을 저장하는 중 오류가 발생했습니다:\n{e}", parent=self)

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
        fmt = self.format_entry.get().strip()
        if err := validate_filename_format(fmt):
            errors.append(err)
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
        out = self.cfg.resolve_path("output_dir") / output_filename(date, church, worship, fmt)
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

        run_in_thread(lambda: builder.build(BuildRequest(items, church, date, worship, fmt)), done)

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
