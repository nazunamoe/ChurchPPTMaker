"""예배 순서 PPT 제작 프로그램 — 메인 윈도우. (버전은 version.py)

레이아웃은 UI 초안 기준 (CLAUDE.md §4 참고).
"""
from __future__ import annotations

import ctypes
import datetime as dt
import math
import os
import shutil
import subprocess
import sys
import tempfile
import time
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from bible_bridge import Bible2PPT, Bible2PPTError, match_bible, parse_references
from config_manager import MAX_PRESETS, ConfigManager, clean_temp_keep, PresetManager, save_worship_to_text, load_worship_from_text
from models import (DEFAULT_FILENAME_FORMAT, FILENAME_FIELDS, KIND_BIBLE, KIND_CCM, KIND_CUSTOM, KIND_HYMN, KIND_LABELS,
                    KIND_PRAYER, KIND_SERMON, MAX_FILL, ServiceItem, make_item, numbered_token, validate_item)
from ui_widgets import (ACCENT, BG, BORDER, BTN_HOVER, CARD, DANGER, FONT_FAMILY, KIND_COLORS, PILL_BORDER, SUBTLE, TEXT,
                        TEXT2, BibleField, CalendarPopup, DragController, HelpDialog, ItemDialog, LogDialog, SettingsDialog, TempFilesDialog,
                        accent_button, card, font, mode_color, post_to_ui, pump_ui_queue, run_in_thread, scaled,
                        secondary_button, tk_font)
from version import APP_NAME, __version__
from window_frame import CustomFrame
from worship_builder import (PPT_EXTS, BuildError, BuildRequest, WorshipBuilder, find_template, list_ccm_files,
                             list_ppt_files, list_temp_results, output_conflict, output_filename, prayer_template_names,
                             prune_temp_results, sanitize_filename, saved_name_for_temp, sermon_template, temp_dir_conflict,
                             temp_result_dir, template_fill_slots, validate_filename_format)

APP_TITLE = f"{APP_NAME} v{__version__}"
DATE_FMT = "%Y-%m-%d"
MODES = {"라이트": "Light", "다크": "Dark", "시스템": "System"}
SLIDE_SIZE_LABELS = {"원본": "original", "16:9": "16:9", "4:3": "4:3"}  # 화면 표시 → config.ppt.slide_size
MIN_FILE_ROWS = 4  # 최소 창 크기에서도 템플릿 파일 목록이 이만큼은 보이도록
FILE_FONT_SIZE = 13.2  # 템플릿 파일 목록 글자 크기 (기본 12 의 110% — 항목 줄 높이도 글자 크기를 따라 커짐)
ORDER_SCROLL_UNIT = 2.4  # 순서 목록 스크롤 단위(논리 px). 휠 한 칸 = 20단위 ≈ 48 = 행 하나 (CTk 기본 1px → 20px)
RESIZE_SETTLE_MS = 150  # 창 크기 조절이 이 시간 동안 멈추면 그때 내용 배치를 새 크기로 맞춤
MIN_WIDTH = 1000  # 최소 창 너비 (논리 단위). 처음 열 때도 이 너비 × 최소 높이로 연다
TOP_RESIZE_EDGE = 5  # 제목 줄 맨 위 이 높이(논리 px)에서는 창 위쪽 크기 조절
CAPTION_FONT = "Segoe MDL2 Assets"  # Windows 기본 창 버튼 모양 (Win10/11 기본 글꼴)
CAPTION_GLYPHS = {"min": "\uE921", "max": "\uE922", "restore": "\uE923", "close": "\uE8BB"}
CLOSE_HOVER = ("#C42B1C", "#C42B1C")
START_REVEAL_MS = 2000  # 시작 크기 계산이 늦어져도 이 시간 뒤에는 창을 보이게 함
# 최소 높이를 재기 전 (투명한) 첫 창 높이. 왼쪽 패널보다 낮으면 템플릿 목록이 0줄로 눌려 사라져서(unmapped)
# _apply_min_size 가 잴 수 없음 — 820 은 성경 영역에 역본2 줄을 넣은 뒤 부족했음. 화면보다 크면 화면에 맞춰짐
START_MEASURE_HEIGHT = 1000
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


def _double_click_ms() -> int:
    """Windows 더블클릭 간격(ms). 순서 행 알약은 한 번 클릭=수정, 두 번=원본 열기라 이만큼 기다렸다가 수정 창을 띄운다."""
    try:
        import ctypes
        return max(200, min(int(ctypes.windll.user32.GetDoubleClickTime()), 500))
    except (AttributeError, OSError):
        return 400


def _send_to_recycle_bin(path: Path) -> str | None:
    """파일을 휴지통으로 이동 (실수로 지워도 되살릴 수 있게). 실패하면 오류 메시지."""
    try:
        import ctypes
        from ctypes import wintypes

        class SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [("hwnd", wintypes.HWND), ("wFunc", wintypes.UINT), ("pFrom", wintypes.LPCWSTR),
                        ("pTo", wintypes.LPCWSTR), ("fFlags", ctypes.c_uint16), ("fAnyOperationsAborted", wintypes.BOOL),
                        ("hNameMappings", ctypes.c_void_p), ("lpszProgressTitle", wintypes.LPCWSTR)]

        FO_DELETE, FOF_SILENT, FOF_NOCONFIRMATION, FOF_ALLOWUNDO, FOF_NOERRORUI = 3, 0x4, 0x10, 0x40, 0x400
        op = SHFILEOPSTRUCTW(wFunc=FO_DELETE, pFrom=str(path) + "\0",  # 목록은 NUL 두 개로 끝나야 함
                             fFlags=FOF_SILENT | FOF_NOCONFIRMATION | FOF_ALLOWUNDO | FOF_NOERRORUI)
        rc = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    except (AttributeError, OSError) as e:
        return f"휴지통으로 옮기지 못했습니다: {e}"
    if rc != 0 or op.fAnyOperationsAborted or path.exists():
        return "휴지통으로 옮기지 못했습니다. 파일이 PowerPoint 에서 열려 있으면 닫고 다시 시도해 주세요."
    return None


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
        # 처음에는 최소 크기로 연다. 최소 높이는 첫 배치 뒤에야 알 수 있으므로(_apply_min_size) 그때까지 창을 투명하게 두어
        # 큰 창이 보였다가 줄어드는 깜빡임을 없앤다. 혹시 크기 계산이 안 되더라도 START_REVEAL_MS 뒤에는 반드시 보이게 함.
        self._revealed = False
        try:
            self.attributes("-alpha", 0.0)
        except tk.TclError:
            pass
        self.after(START_REVEAL_MS, self._reveal)
        self._fit_to_screen(MIN_WIDTH, START_MEASURE_HEIGHT)
        _disable_full_redraw_on_resize(self)

        self.items: list[ServiceItem] = []
        self.path_vars: dict[str, tk.StringVar] = {}  # 폴더 키 → 화면에 표시하는 경로 (_choose_folder 가 갱신)
        self.row_widgets: list[tk.Frame] = []
        self._row_parts: list[dict] = []  # 행별 위젯·전체 글자·줄 바꿈 상태 (_layout_rows)
        self._layout_job = None
        self._shown_files: list[str] = []  # 템플릿 목록 상자에 보이는 순서대로의 파일 이름
        self.selected_file: str | None = None
        self.bible_values: list[str] = list(self.cfg.get("bible2ppt", "bible_list") or [])
        self._bible_loading = False
        self.log_lines: list[str] = []
        self.building = False
        self._save_job = None
        self._calendar: CalendarPopup | None = None
        self._pending_click: tuple[int, str] | None = None  # 순서 알약 한 번 클릭 → (행, 수정 창 예약 after id)
        self._bible_preview = False  # 말씀 더블클릭으로 Bible2PPT 실행 중 (PowerPoint 를 쓰므로 생성과 동시에 하지 않음)

        self.drag = DragController(self, self._on_drag_hover, self._on_drag_drop, self._on_drag_click)
        self.frame = CustomFrame(self)  # Windows 제목 표시줄 대신 프로그램 안의 제목 줄 (처음 표시될 때 적용)
        self._last_state = "normal"
        self._normal_size: tuple[int, int] | None = None  # 일반 상태의 마지막 창 크기 (실제 픽셀)

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
        if e.widget is not self:
            return
        self._on_state_change()
        if (e.width, e.height) == self._content_size:
            return
        if self._content_size is None:  # 처음 표시될 때는 바로 맞추고, 실제 배치를 보고 최소 높이 결정
            # 아직 투명할 때 Windows 제목 표시줄을 없앰 → 클라이언트가 그만큼 커지므로 새 크기를 반영한 뒤 맞춤
            if self.frame.install():
                self.update_idletasks()
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
        self.minsize(min(MIN_WIDTH, w), min(660, h))

    def _apply_min_size(self):
        """최소 높이를 '템플릿 목록이 MIN_FILE_ROWS 줄 보이는 높이'로 맞춘다.
        고정값 660 은 왼쪽 패널보다 낮아서 최소 크기에서 템플릿 목록이 0줄로 눌려 사라졌다.
        창 높이가 줄면 왼쪽에서는 템플릿 목록만 줄어드므로 (현재 창 높이 - 현재 목록 높이 + 목록 MIN_FILE_ROWS 줄 높이)가
        필요한 최소 높이다. CTk 위젯 요청 크기(reqheight)는 실제보다 커서 실제 배치로 계산. 화면보다 크게 잡지는 않음."""
        self.update_idletasks()
        lb = self.file_listbox
        if not lb.winfo_ismapped():
            self._reveal()
            return
        scale = ctk.ScalingTracker.get_window_scaling(self)
        need = math.ceil((self.winfo_height() - lb.winfo_height() + lb.winfo_reqheight()) / scale)  # 픽셀 → 논리 단위
        screen = self._available_size()[1]  # _fit_to_screen 과 같은 기준 (논리 단위)
        min_w, min_h = round(self.wm_minsize()[0] / scale), max(660, min(need, screen))
        if not self._revealed:  # 처음 열 때: 최소 크기로 맞추고 화면 가운데에 둔 뒤 보이기
            self._fit_to_screen(min_w, min_h)
        self.minsize(min_w, min_h)
        self.after(50, self._reveal)

    def _reveal(self):
        if self._revealed:
            return
        self._revealed = True
        try:
            self.attributes("-alpha", 1.0)
        except tk.TclError:
            pass

    # ================================================================ 상단
    def _build_header(self):
        bar = ctk.CTkFrame(self._content, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", padx=(20, 0), pady=(0, 8))  # 창 버튼이 오른쪽 위 끝에 붙도록
        title = ctk.CTkLabel(bar, text=APP_TITLE, font=font(20, "bold"), text_color=TEXT)
        title.pack(side="left", pady=(14, 0))
        secondary_button(bar, "?  도움말", self._show_help, width=84).pack(side="left", padx=(12, 0), pady=(14, 0))
        # Windows 제목 표시줄 대신 쓰는 창 버튼 (window_frame.CustomFrame — 창 오른쪽 위 끝에 붙임)
        caption = ctk.CTkFrame(bar, fg_color="transparent")
        caption.pack(side="right", anchor="n")
        self.max_btn = None
        for glyph, cmd, hover in ((CAPTION_GLYPHS["min"], self.iconify, BTN_HOVER),
                                  (CAPTION_GLYPHS["max"], self._toggle_maximize, BTN_HOVER),
                                  (CAPTION_GLYPHS["close"], self._on_close, CLOSE_HOVER)):
            b = ctk.CTkButton(caption, text=glyph, width=46, height=34, corner_radius=0, fg_color="transparent",
                              hover_color=hover, text_color=TEXT, font=(CAPTION_FONT, 10), command=cmd)
            b.pack(side="left")
            if cmd == self._toggle_maximize:
                self.max_btn = b
        secondary_button(bar, "⚙  설정", self._open_settings, width=88).pack(side="right", padx=(0, 12), pady=(14, 0))
        current = {v: k for k, v in MODES.items()}.get(self.cfg.get("appearance_mode"), "시스템")
        self.mode_seg = ctk.CTkSegmentedButton(bar, values=list(MODES), command=self._set_mode, font=font(12),
                                               selected_color=ACCENT, height=32)
        self.mode_seg.set(current)
        self.mode_seg.pack(side="right", padx=10, pady=(14, 0))
        # 제목 줄 빈 곳·제목 글자: 끌면 창 옮기기, 두 번 누르면 최대화/복원, 맨 위 가장자리는 크기 조절
        self._title_widgets = (bar, title)
        self._last_title_press = (0, 0, 0)
        for w in self._title_widgets:
            w.bind("<ButtonPress-1>", self._on_title_press)
            w.bind("<Motion>", self._on_title_motion)

    # ---- 창 틀 (Windows 제목 표시줄 없이)
    def _top_edge(self, y_root: int) -> bool:
        return not self.frame.is_maximized() and y_root - self.winfo_rooty() < scaled(self, TOP_RESIZE_EDGE)

    def _on_title_motion(self, e):
        cursor = "sb_v_double_arrow" if self._top_edge(e.y_root) else ""
        for w in self._title_widgets:
            try:
                w.configure(cursor=cursor)
            except (tk.TclError, ValueError):
                pass

    def _on_title_press(self, e):
        if self._top_edge(e.y_root):
            self.frame.start_resize_top(e.x_root)
            return
        # 두 번 누르기: 창 옮기기(Windows 이동 루프)에 들어가면 Tk 가 더블클릭을 받지 못하므로 직접 판정
        t, x, y = self._last_title_press
        now = time.monotonic()
        limit = ctypes.windll.user32.GetDoubleClickTime() / 1000
        if now - t <= limit and abs(e.x_root - x) <= scaled(self, 4) and abs(e.y_root - y) <= scaled(self, 4):
            self._last_title_press = (0, 0, 0)
            self._toggle_maximize()
            return
        self._last_title_press = (now, e.x_root, e.y_root)
        self.frame.start_move()

    def _toggle_maximize(self):
        self.state("normal" if self.state() == "zoomed" else "zoomed")

    def _on_state_change(self):
        """최대화 버튼 모양 맞추기 + 최대화에서 돌아올 때 원래 크기로.
        Tk 는 최대화 때 잰 틀 두께(위쪽을 안으로 들인 만큼 더 큼)를 복원할 때도 써서 창이 한 번 11px 쯤 커졌다(실측)
        → 일반 상태의 마지막 크기를 기억해 두었다가 복원하면 그 크기로 되돌린다."""
        state = self.state()
        if self.max_btn is not None:
            glyph = CAPTION_GLYPHS["restore" if state == "zoomed" else "max"]
            if self.max_btn.cget("text") != glyph:
                self.max_btn.configure(text=glyph)
        if state == "normal":
            size = (self.winfo_width(), self.winfo_height())
            if self._last_state == "zoomed" and self._normal_size and size != self._normal_size:
                scale = ctk.ScalingTracker.get_window_scaling(self)
                w, h = (round(v / scale) for v in self._normal_size)
                self.after_idle(lambda: self.geometry(f"{w}x{h}"))
            else:
                self._normal_size = size
        self._last_state = state

    def _build_info_bar(self):
        bar = card(self._content)
        bar.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        # 줄 틀은 카드 색으로 칠해지므로 둥근 테두리를 덮지 않도록 안쪽으로 들여 배치
        # 입력칸 너비는 최소 창 너비(MIN_WIDTH)에서 두 줄이 오른쪽 여백 8 까지 꽉 차도록 맞춘 값 (두 줄 오른쪽 끝도 일치).
        # 줄에 요소를 더하거나 너비를 바꾸면 최소 너비에서 넘치지 않는지 함께 확인할 것.
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
        self.church_entry = ctk.CTkEntry(line1, width=255, height=34, font=font(13), placeholder_text="예: 사랑의교회")
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
        self.worship_entry = ctk.CTkEntry(line1, width=192, height=34, font=font(13), placeholder_text="예: 주일오전예배")
        self.worship_entry.insert(0, self.cfg.get("worship_type") or "")
        self.worship_entry.pack(side="left", padx=(0, 8))
        self.worship_entry.bind("<KeyRelease>", on_change)

        # 결과 파일 이름 형식: {항목} 자리에 값이 들어감. 프리셋을 저장하면 함께 저장되고, 불러오면 프리셋의 형식으로 바뀜
        # 출력 폴더: 결과 파일이 저장되는 곳이라 파일 이름 형식 바로 왼쪽에 둔다
        # 실제 파일 이름 미리보기는 길이가 교회 이름·예배 종류에 따라 늘어나 같은 줄에 두면 상단 카드(창) 너비를 밀어내므로
        # 아랫줄의 형식 입력칸 밑에 두고, 남는 너비보다 길면 줄여서(…) 표시한다 (_update_filename_preview).
        line2.grid_columnconfigure(8, weight=1)
        label(line2, "출력 폴더").grid(row=0, column=0, padx=(8, 8), pady=(5, 2))
        var = tk.StringVar(value=str(self.cfg.resolve_path("output_dir")))
        self.path_vars["output_dir"] = var
        ctk.CTkEntry(line2, textvariable=var, state="readonly", width=245, height=34, font=font(12),
                     text_color=TEXT2).grid(row=0, column=1)
        for col, (text, cmd) in enumerate((("📁", lambda: self._choose_folder("output_dir")),
                                           ("↗", lambda: self._open_folder("output_dir"))), 2):
            b = secondary_button(line2, text, cmd, width=34)
            b.configure(height=34)
            b.grid(row=0, column=col, padx=(4, 0))

        label(line2, "파일 이름").grid(row=0, column=4, padx=(28, 8))
        self.format_entry = ctk.CTkEntry(line2, width=307, height=34, font=font(13),
                                         placeholder_text=DEFAULT_FILENAME_FORMAT)
        self.format_entry.insert(0, self.cfg.get("filename_format") or DEFAULT_FILENAME_FORMAT)
        self.format_entry.grid(row=0, column=5)
        self.format_entry.bind("<KeyRelease>", on_change)
        self.format_menu_btn = secondary_button(line2, "항목 ▾", self._open_format_menu, width=64)
        self.format_menu_btn.configure(height=34)
        self.format_menu_btn.grid(row=0, column=6, padx=(4, 4))
        b = secondary_button(line2, "기본값", lambda: self._set_filename_format(DEFAULT_FILENAME_FORMAT), width=64)
        b.configure(height=34)
        b.grid(row=0, column=7)
        self.filename_preview = ctk.CTkLabel(line2, text="", font=font(12), text_color=TEXT2, anchor="w", height=20)
        self.filename_preview.grid(row=1, column=5, columnspan=4, sticky="w", padx=(2, 0), pady=(2, 4))
        self._preview_text = ("", TEXT2)  # 줄이기 전 전체 문구 (창 크기가 바뀌면 다시 맞춤)
        # 줄 너비가 실제로 바뀐 뒤에 다시 줄인다 (창 크기 조절은 _fit_content 가 멈췄을 때 한 번만 반영하므로 여기서도 한 번)
        line2.bind("<Configure>", lambda e: self._fit_filename_preview(), add="+")
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
            self._preview_text = ("⚠ " + err.splitlines()[0], DANGER)
        else:
            name = output_filename(self.date_entry.get().strip(), self.church_entry.get().strip(),
                                   self.worship_entry.get().strip(), fmt)
            self._preview_text = (f"→  {name}", TEXT2)
        self._fit_filename_preview()

    def _fit_filename_preview(self):
        """미리보기가 형식 입력칸 왼쪽 끝부터 카드 오른쪽 끝까지 남는 너비보다 길면 가운데를 '…' 로 줄인다.
        (긴 글자가 라벨 요청 너비를 키워 상단 카드·창을 넓히지 않도록)"""
        text, color = self._preview_text
        lb = self.filename_preview
        avail = lb.master.winfo_width() - self.format_entry.winfo_x() - scaled(self, 12)
        if avail > 50:  # 아직 배치 전(너비 1)이면 그대로
            f = tkfont.Font(family=FONT_FAMILY, size=-scaled(self, 12))
            if f.measure(text) > avail:
                head, tail = text[:len(text) // 2], text[len(text) // 2:]
                while (head or tail) and f.measure(head + "…" + tail) > avail:
                    head, tail = head[:-1], tail[1:]
                text = head + "…" + tail
        lb.configure(text=text, text_color=color)

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
        ctk.CTkLabel(head, text="끌어다 놓기 · 더블클릭으로 추가", font=font(11), text_color=TEXT2).pack(
            side="right", padx=(0, 8))
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
        # [불러오기] [열기] [삭제] [순서에 추가] — 템플릿 폴더 파일 관리
        secondary_button(foot, "순서에 추가", self._add_selected_file, width=90).pack(side="right")
        del_btn = secondary_button(foot, "삭제", self._delete_selected_template, width=52)
        del_btn.configure(text_color=DANGER)
        del_btn.pack(side="right", padx=(0, 6))
        secondary_button(foot, "열기", self._open_selected_template, width=52).pack(side="right", padx=(0, 6))
        secondary_button(foot, "불러오기", self._import_templates, width=72).pack(side="right", padx=(0, 6))

        # 성경 구절 추가 (Bible2PPT 연동)
        bc = card(left)
        bc.grid(row=2, column=0, sticky="ew")
        bc.grid_columnconfigure(0, weight=1)
        # 제목 줄: 성경 구절 추가 (Bible2PPT) ······ [↻ 역본 목록] [순서에 추가]
        bhead = ctk.CTkFrame(bc, fg_color="transparent")
        bhead.grid(row=0, column=0, columnspan=2, sticky="ew", padx=14, pady=(10, 6))
        ctk.CTkLabel(bhead, text="성경 구절 추가", font=font(15, "bold"), text_color=TEXT).pack(side="left")
        ctk.CTkLabel(bhead, text="Bible2PPT", font=font(11), text_color=TEXT2).pack(side="left", padx=(6, 0))
        accent_button(bhead, "순서에 추가", self._add_bible_from_form, width=96).pack(side="right")
        secondary_button(bhead, "↻", self._load_bible_list, width=32).pack(side="right", padx=(0, 6))
        self.verse_entry = ctk.CTkEntry(bc, height=34, font=font(13), placeholder_text="예: 창1:1-5, 롬1-3, 요3:16")
        self.verse_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 6))
        self.verse_entry.bind("<Return>", lambda e: self._add_bible_from_form())
        # [기본 성경 역본][역본1 🔍] / [두 번째 성경 역본][역본2 🔍] — 역본2 를 고르면 Bible2PPT 가 두 역본을 나란히 배치 (-b 두 번)
        brow = ctk.CTkFrame(bc, fg_color="transparent")
        brow.grid(row=2, column=0, columnspan=2, sticky="ew", padx=14)
        brow.grid_columnconfigure(1, weight=1)
        for r, text in enumerate(("기본 성경 역본", "두 번째 성경 역본")):
            ctk.CTkLabel(brow, text=text, font=font(12), text_color=TEXT2, anchor="w").grid(
                row=r, column=0, sticky="w", padx=(0, 8), pady=(6 if r else 0, 0))
        self.bible_field = BibleField(brow, values=self.bible_values, width=180, font_size=12,
                                      placeholder="비우면 ⚙ 설정의 기본값")
        self.bible_field.set(self.cfg.get("bible2ppt", "bible") or "")
        self.bible_field.grid(row=0, column=1, sticky="ew")
        self.bible2_field = BibleField(brow, values=self.bible_values, width=180, font_size=12,
                                       placeholder="선택 · 두 역본을 나란히 표시", clear_text="지정 안 함")
        self.bible2_field.set(self.cfg.get("bible2ppt", "bible2") or "")
        self.bible2_field.grid(row=1, column=1, sticky="ew", pady=(6, 0))
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
        clear_btn = secondary_button(head, "전체 비우기", self._clear_items, width=80)  # 옆 버튼과 같은 배경, 글자만 빨강
        clear_btn.configure(height=28, font=font(12), text_color=DANGER)
        clear_btn.pack(side="right")
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
        self._set_order_scroll_speed()
        # 목록 너비가 바뀌면 긴 행의 줄 바꿈을 다시 정함 (창 크기 조절은 멈춘 뒤 한 번만 배치되므로 _fit_content 후 1회)
        self.scroll._parent_canvas.bind("<Configure>", self._schedule_layout_rows, add="+")

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
        left.grid(row=0, column=0, rowspan=2, sticky="ew", padx=16, pady=12)
        self.status = ctk.CTkLabel(left, text="준비됨", font=font(12), text_color=TEXT2, anchor="w")
        self.status.pack(fill="x")
        self.progress = ctk.CTkProgressBar(left, height=6, progress_color=ACCENT)
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(4, 0))
        # 상태 문구의 부연 설명(저장 방법·경고 수 등)은 진행 막대 아랫줄에 — 한 줄에 이어 쓰면 상태 칸을 넘어 잘렸음
        self.status_detail = ctk.CTkLabel(left, text="", font=font(11), text_color=TEXT2, anchor="w", justify="left",
                                          height=18)
        self.status_detail.pack(fill="x", pady=(2, 0))
        # 칸보다 길면 잘리지 않고 다음 줄로 (하단 바 높이는 오른쪽 버튼 두 줄이 정하므로 두 줄까지는 바가 커지지 않음)
        left.bind("<Configure>", lambda e: self.status_detail.configure(
            wraplength=max(100, e.width / ctk.ScalingTracker.get_widget_scaling(self))), add="+")
        # [로그] 위 / [임시 파일](임시 실행으로 만든 PPT 다시 열기·저장) 아래
        secondary_button(bar, "로그", self._show_log, width=76, height=30).grid(
            row=0, column=1, padx=(0, 8), pady=(12, 2), sticky="s")
        secondary_button(bar, "임시 파일", self._show_temp_files, width=76, height=30).grid(
            row=1, column=1, padx=(0, 8), pady=(2, 12), sticky="n")
        self.open_after = ctk.BooleanVar(value=bool(self.cfg.get("ppt", "open_after_build")))
        self.slideshow = ctk.BooleanVar(value=bool(self.cfg.get("ppt", "start_slideshow")))
        self.temp_mode = ctk.BooleanVar(value=bool(self.cfg.get("ppt", "temp_mode")))
        self.open_after_cb = ctk.CTkCheckBox(
            bar, text="생성 후 열기", variable=self.open_after, font=font(12), checkbox_width=20,
            checkbox_height=20, fg_color=ACCENT, command=self._on_open_after)
        self.open_after_cb.grid(row=0, column=2, padx=8, pady=(12, 4), sticky="w")
        self.slideshow_cb = ctk.CTkCheckBox(
            bar, text="슬라이드 쇼로 시작", variable=self.slideshow, font=font(12), checkbox_width=20,
            checkbox_height=20, fg_color=ACCENT,
            command=lambda: self.cfg.set("ppt", "start_slideshow", self.slideshow.get()))
        self.slideshow_cb.grid(row=0, column=3, padx=8, pady=(12, 4), sticky="w")
        # 두 체크박스 아래: 결과 슬라이드 비율 (원본 = 첫 PPT 크기)
        size_row = ctk.CTkFrame(bar, fg_color="transparent")
        size_row.grid(row=1, column=2, columnspan=2, padx=8, pady=(4, 12), sticky="w")
        ctk.CTkLabel(size_row, text="슬라이드 비율", font=font(12), text_color=TEXT).pack(side="left", padx=(0, 8))
        self.size_seg = ctk.CTkSegmentedButton(size_row, values=list(SLIDE_SIZE_LABELS), font=font(12), height=28,
                                               selected_color=ACCENT, command=self._set_slide_size)
        cur = self.cfg.get("ppt", "slide_size")
        self.size_seg.set(next((k for k, v in SLIDE_SIZE_LABELS.items() if v == cur), "원본"))
        self.size_seg.pack(side="left")
        self.build_btn = accent_button(bar, "예배 PPT 병합 및 생성", self._build, width=220, height=42,
                                       font=font(15, "bold"))
        self.build_btn.grid(row=0, column=4, padx=(8, 16), pady=(12, 2), sticky="s")
        # 생성 버튼 아래: 저장 없이 임시 폴더에 만들어 바로 실행 (최근 ppt.temp_keep 개 보관).
        # 버튼 열(220)에 두어 하단 바가 넓어지지 않게 함 — 옆 열에 두면 상태 표시 칸이 약 2/3 로 줄었음
        ctk.CTkCheckBox(bar, text="저장 없이 임시 파일로 실행", variable=self.temp_mode, font=font(12), checkbox_width=18,
                        checkbox_height=18, fg_color=ACCENT, command=self._on_temp_mode).grid(
            row=1, column=4, padx=(8, 16), pady=(2, 10), sticky="n")
        self._on_temp_mode(save=False)  # 생성 버튼 글자·'생성 후 열기' 상태 맞춤

    def _set_slide_size(self, label: str):
        self.cfg.set("ppt", "slide_size", SLIDE_SIZE_LABELS[label])
        self._set_status("슬라이드 비율: " + ("첫 PPT 크기 그대로" if label == "원본" else label))

    def _on_open_after(self, save: bool = True):
        """'슬라이드 쇼로 시작'은 결과를 열 때('생성 후 열기' 또는 임시 실행)만 선택 가능."""
        if save:
            self.cfg.set("ppt", "open_after_build", self.open_after.get())
        opens = self.open_after.get() or self.temp_mode.get()
        self.slideshow_cb.configure(state="normal" if opens else "disabled")

    def _on_temp_mode(self, save: bool = True):
        """임시 실행은 항상 바로 열므로 '생성 후 열기'는 의미가 없어 잠그고, 생성 버튼 글자로 모드를 알린다."""
        temp = self.temp_mode.get()
        if save:
            self.cfg.set("ppt", "temp_mode", temp)
            self._set_status("임시 실행: 저장하지 않고 임시 파일로 만들어 바로 엽니다." if temp
                             else "생성하면 출력 폴더에 저장합니다.")
        self.open_after_cb.configure(state="disabled" if temp else "normal")
        if not self.building:
            self.build_btn.configure(text=self._build_label())
        self._on_open_after(save=False)

    def _build_label(self) -> str:
        return "예배 PPT 임시 생성 및 실행" if self.temp_mode.get() else "예배 PPT 병합 및 생성"

    # ================================================================ 순서 목록 렌더링
    def _render_list(self):
        self._cancel_pending_click()  # 행 번호가 바뀌므로 예약된 수정 창은 취소
        for w in self.scroll.winfo_children():
            w.destroy()
        self.row_widgets = []
        self._row_parts = []
        self.indicator = ctk.CTkFrame(self.scroll, height=4, fg_color=ACCENT, corner_radius=0, border_width=0)

        if not self.items:
            ctk.CTkLabel(self.scroll, text="아래 [기도] [찬송] [말씀] [설교] [CCM] [사용자 지정] 버튼을 끌어다 놓아\n예배 순서를 추가하세요.",
                         font=font(13), text_color=TEXT2, justify="center").pack(pady=60)
        for idx, item in enumerate(self.items):
            self.row_widgets.append(self._make_row(idx, item))
        self.count_label.configure(text=f"{len(self.items)}개 항목")
        if self.items:
            self.scroll.update_idletasks()  # 알약 너비(reqwidth)를 알아야 줄 바꿈 여부를 정할 수 있음
            self._layout_rows()

    def _make_row(self, idx: int, item: ServiceItem) -> tk.Frame:
        # Windows 에선 위젯(HWND) 수가 곧 창 크기 조절·목록 갱신 비용이라(CLAUDE.md §11) 둥근 알약 2개만 CTk 로 그리고
        # 나머지(행 틀·손잡이·연결선·▲▼✕)는 일반 tk 위젯으로 그린다. 색은 모드가 바뀌면 _on_appearance_change 가 다시 그림.
        # grid: [손잡이][왼쪽 알약][연결선 ─ 늘어남][오른쪽 알약][▲▼✕] — 한 줄에 안 들어가면 _layout_rows 가 오른쪽 알약을
        # 아랫줄로 내림.
        bg = mode_color(SUBTLE)
        row = tk.Frame(self.scroll, bg=bg)
        row.pack(fill="x", padx=scaled(self, 6), pady=scaled(self, 3))
        row.grid_columnconfigure(2, weight=1)

        handle = tk.Label(row, text="⋮⋮", bg=bg, fg=mode_color(TEXT2), font=tk_font(self, 14, "bold"),
                          width=2, cursor="fleur")
        handle.grid(row=0, column=0, padx=(scaled(self, 2), scaled(self, 6)))
        self.drag.attach(handle, ("move", idx), f"{item.label}  {item.display_value}".strip())

        left = ctk.CTkButton(row, text=item.label, width=140, height=36, corner_radius=10, fg_color=CARD, bg_color=SUBTLE,
                             hover_color=BTN_HOVER, border_width=2, border_color=KIND_COLORS[item.kind],
                             text_color=TEXT, font=font(13, "bold"), command=lambda i=idx: self._on_pill_click(i))
        left.grid(row=0, column=1)
        line = tk.Frame(row, height=max(1, scaled(self, 1.5)), bg=mode_color(PILL_BORDER))
        line.grid(row=0, column=2, sticky="ew")
        for w in (row, line):  # 알약 밖(행 빈 곳·연결선) 더블클릭도 원본 열기
            w.bind("<Double-Button-1>", lambda e, i=idx: self.after_idle(lambda: self._open_item(i)))
        right = ctk.CTkButton(row, text=item.display_value, width=170, height=36, corner_radius=10, fg_color=CARD,
                              bg_color=SUBTLE, hover_color=BTN_HOVER, border_width=1, border_color=PILL_BORDER,
                              text_color=TEXT if item.display_value != "—" else TEXT2, font=font(13),
                              command=lambda i=idx: self._on_pill_click(i))
        right.grid(row=0, column=3)

        tools = tk.Frame(row, bg=bg)
        tools.grid(row=0, column=4, padx=(scaled(self, 8), 0))
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
        self._row_parts.append({"row": row, "handle": handle, "left": left, "right": right, "tools": tools,
                                "texts": (item.label, item.display_value), "wrapped": False, "widths": None,
                                "shown": (item.label, item.display_value)})
        return row

    def _layout_rows(self, _event=None):
        """오른쪽 알약까지 한 줄에 안 들어가는 행은 오른쪽 알약을 아랫줄(오른쪽 정렬)로 내린다.
        아랫줄에도 안 들어갈 만큼 길면 글자를 '…' 로 줄인다 (전체 내용은 알약을 눌러 수정 창에서 확인).
        목록을 그릴 때와 목록 너비가 바뀔 때(창 크기 조절이 끝난 뒤 한 번) 실행 — 다시 만들지 않고 grid 만 바꿈."""
        canvas = getattr(self.scroll, "_parent_canvas", None)
        avail = (canvas.winfo_width() if canvas is not None else self.scroll.winfo_width()) - 2 * scaled(self, 6)
        if avail < scaled(self, 200):
            return  # 아직 배치 전
        min_line = scaled(self, 24)
        for part in self._row_parts:
            try:
                if not part["row"].winfo_exists():
                    continue
            except tk.TclError:
                continue
            if part["widths"] is None:  # 처음 한 번: 전체 글자일 때의 너비
                part["widths"] = (part["handle"].winfo_reqwidth() + scaled(self, 8), part["left"].winfo_reqwidth(),
                                  part["right"].winfo_reqwidth(), part["tools"].winfo_reqwidth() + scaled(self, 8))
            hw, lw, rw, tw = part["widths"]
            wrapped = hw + lw + min_line + rw + tw > avail
            left_max = avail - hw - min_line - tw
            right_max = avail - hw - tw if wrapped else rw
            shown = (self._fit_pill(part["texts"][0], lw, left_max, bold=True),
                     self._fit_pill(part["texts"][1], rw, right_max))
            if shown != part["shown"]:
                part["left"].configure(text=shown[0])
                part["right"].configure(text=shown[1])
                part["shown"] = shown
            if wrapped != part["wrapped"]:
                part["wrapped"] = wrapped
                if wrapped:
                    # 아랫줄 오른쪽 끝을 다른 행의 오른쪽 알약과 맞춤 (▲▼✕ 열 앞까지)
                    part["right"].grid(row=1, column=1, columnspan=3, sticky="e", pady=(scaled(self, 4), 0))
                else:
                    part["right"].grid(row=0, column=3, columnspan=1, sticky="", pady=0)

    def _schedule_layout_rows(self, _event=None):
        if self._layout_job is not None:
            self.after_cancel(self._layout_job)
        self._layout_job = self.after_idle(self._run_layout_rows)

    def _run_layout_rows(self):
        self._layout_job = None
        self._layout_rows()

    def _set_order_scroll_speed(self):
        """휠 한 칸 = 행 하나쯤. CTk 는 Windows 에서 휠 한 칸(delta 120)을 20단위 × 1px 로 스크롤해 행 높이(약 48)의
        절반도 안 움직였고, 배율이 크면 더 느렸다 → 단위를 배율에 맞춰 키움 (드래그 중 자동 스크롤도 같이 빨라짐)."""
        canvas = getattr(self.scroll, "_parent_canvas", None)
        if canvas is not None:
            canvas.configure(yscrollincrement=max(1, scaled(self, ORDER_SCROLL_UNIT)))

    def _fit_pill(self, text: str, full_width: int, max_width: int, bold: bool = False) -> str:
        """알약 너비가 max_width 를 넘으면 글자 끝을 '…' 로 줄인 문구 (알약 여백은 전체 글자일 때 너비로 계산)."""
        if full_width <= max_width or len(text) <= 1:
            return text
        f = tkfont.Font(font=tk_font(self, 13, "bold" if bold else "normal"))
        margin = full_width - f.measure(text)
        room = max_width - margin
        lo, hi = 0, len(text)
        while lo < hi:  # room 안에 들어가는 가장 긴 앞부분
            mid = (lo + hi + 1) // 2
            if f.measure(text[:mid].rstrip() + "…") <= room:
                lo = mid
            else:
                hi = mid - 1
        return text[:lo].rstrip() + "…" if lo else text[:1] + "…"

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
            self._add_template(data, idx)

    def _on_drag_click(self, payload):
        kind, data = payload
        if kind == "new":
            self._prompt_new(data, len(self.items))
        elif kind == "file":
            self._select_file(data)

    # ================================================================ 항목 조작
    def _template_names(self) -> list[str]:
        return [p.stem for p in list_ppt_files(self.cfg.resolve_path("template_dir"))]

    def _dialog(self, kind: str, item: ServiceItem | None = None,
                initial: ServiceItem | None = None) -> ServiceItem | None:
        ccm = list_ccm_files(self.cfg.resolve_path("ccm_dir")) if kind == KIND_CCM else None
        prayer = prayer_template_names(self.cfg) if kind == KIND_PRAYER else None
        token = self.cfg.get("ppt", "text_placeholder") or ""
        return ItemDialog(self, kind, item, template_names=self._template_names(), bible_values=self.bible_values,
                          default_bible=self.bible_field.get().strip(),
                          default_bible2=self.bible2_field.get().strip(),
                          fallback_bible=self.cfg.get("bible2ppt", "bible") or "", ccm_files=ccm, prayer_names=prayer,
                          initial=initial, fill_slots=self._custom_slots,
                          slot_names=[numbered_token(token, n) for n in range(1, MAX_FILL + 1)] if token else None).show()

    def _custom_slots(self, name: str) -> set[int] | None:
        """템플릿 폴더의 name PPT 에 있는 텍스트 입력 자리 번호 (파일이 없으면 빈 집합, 형식상 알 수 없으면 None)."""
        path = find_template(self.cfg.resolve_path("template_dir"), name)
        if path is None:
            return set()
        return template_fill_slots(path, self.cfg.get("ppt", "text_placeholder") or "",
                                   self.cfg.get("ppt", "text_shape_name") or "")

    def _add_template(self, stem: str, idx: int | None = None):
        """템플릿 목록에서 고른 PPT 를 사용자 지정 항목으로 추가. 템플릿에 {{내용n}} 자리가 있으면 내용을 입력받는다."""
        initial = ServiceItem(KIND_CUSTOM, stem)
        if self._custom_slots(stem):
            initial = self._dialog(KIND_CUSTOM, initial=initial)
            if initial is None:  # 취소
                return
        self._insert(initial, idx)

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

    # ---- 순서 알약: 한 번 클릭 = 수정 창, 두 번 클릭 = 원본 PPT 열기
    # CTkButton 은 누를 때마다 command 를 부르므로 더블클릭 간격만큼 기다렸다가 두 번째 클릭이 없으면 수정 창을 연다.
    def _on_pill_click(self, idx: int):
        if self.drag.dragging:
            return
        pending = self._pending_click
        self._cancel_pending_click()
        if pending is not None and pending[0] == idx:
            self._open_item(idx)
            return
        self._pending_click = (idx, self.after(_double_click_ms(), lambda: self._run_pending_click(idx)))

    def _run_pending_click(self, idx: int):
        self._pending_click = None
        self._edit_item(idx)

    def _cancel_pending_click(self):
        if self._pending_click is not None:
            self.after_cancel(self._pending_click[1])
            self._pending_click = None

    def _open_item(self, idx: int):
        """순서 항목의 원본을 PowerPoint 로 연다. 기도·설교·사용자 지정 = 템플릿, 찬송·CCM = 그 PPT,
        말씀 = Bible2PPT 로 구절 PPT 를 만들어 연다."""
        if not 0 <= idx < len(self.items):
            return
        item = self.items[idx]
        if item.kind == KIND_BIBLE:
            self._open_bible(item)
            return
        path, problem = WorshipBuilder(self.cfg, lambda m: None, lambda v, m: None).resolve_item(item)
        if path is None:
            messagebox.showerror("열 수 없습니다", problem or "PPT 파일을 찾지 못했습니다.", parent=self)
            return
        self._open_ppt(path)

    def _busy_merging(self) -> bool:
        """생성(병합) 중이면 에러 창 + True. 병합의 [원본 서식 유지] 붙여넣기(ExecuteMso)는 PowerPoint 의 '활성 창'에
        붙여넣으므로, 그사이 사용자가 연 PPT 가 활성 창이 되면 그 파일에 슬라이드가 붙을 수 있다."""
        if self.building:
            messagebox.showerror("잠시 후 다시", "PPT 를 만드는 중에는 PowerPoint 로 파일을 열 수 없습니다.\n"
                                 "끝난 뒤 다시 시도해 주세요.", parent=self)
        return self.building

    def _open_ppt(self, path: Path) -> bool:
        """PPT 를 PowerPoint 로 열어 수정 (기본 '열기' 동작 = 편집 모드)."""
        if self._busy_merging():
            return False
        if path.suffix.lower() not in PPT_EXTS or not path.is_file():  # PPT 가 아닌 파일을 실행하지 않도록
            messagebox.showerror("열 수 없습니다", f"PPT 파일이 아닙니다:\n{path}", parent=self)
            return False
        try:
            os.startfile(path)
        except OSError as e:
            messagebox.showerror("열 수 없습니다", f"파일을 열 수 없습니다.\n{path}\n\n{e}", parent=self)
            return False
        self._set_status(f"PowerPoint 로 열었습니다: {path.name}", "수정 후 저장하면 다음 생성부터 반영됩니다.")
        return True

    def _open_bible(self, item: ServiceItem):
        """말씀 구절을 Bible2PPT 로 만들어 연다. 결과는 임시 폴더(%TEMP%\\ChurchPPTMaker_말씀)에 저장."""
        if self.building or self._bible_preview:
            messagebox.showerror("잠시 후 다시", "PPT 를 만드는 중입니다. 끝난 뒤 다시 시도해 주세요.", parent=self)
            return
        builder = WorshipBuilder(self.cfg, lambda m: None, lambda v, m: None)
        if (err := builder.check_bible()) is not None:
            messagebox.showerror("Bible2PPT 설정 필요", err, parent=self)
            return
        out_dir = Path(tempfile.gettempdir()) / "ChurchPPTMaker_말씀"
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            for old in out_dir.glob("*.pptx"):  # 하루 지난 미리보기는 정리 (열려 있으면 건너뜀)
                if time.time() - old.stat().st_mtime > 86400:
                    old.unlink(missing_ok=True)
        except OSError:
            pass
        out = out_dir / f"{sanitize_filename('말씀 ' + item.value)}_{dt.datetime.now():%H%M%S}.pptx"
        self._bible_preview = True
        self._set_status(f"Bible2PPT 로 '{item.value}' 만드는 중…")

        def done(result, err):
            self._bible_preview = False
            if err:
                self._set_status("성경 PPT 만들기 실패")
                messagebox.showerror("Bible2PPT", str(err), parent=self)
                return
            self._open_ppt(Path(result))

        run_in_thread(lambda: builder.generate_bible(item, out), done)

    def _edit_item(self, idx: int):
        if self.drag.dragging or not 0 <= idx < len(self.items):
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
            self._add_template(got[1])

    def _select_file(self, stem: str):
        self.selected_file = stem
        lb = self.file_listbox
        lb.selection_clear(0, "end")
        if stem in self._shown_files:
            lb.selection_set(self._shown_files.index(stem))

    def _selected_template(self) -> Path | None:
        """목록에서 고른 템플릿 파일 경로 (안 골랐으면 에러 창 후 None)."""
        path = find_template(self.cfg.resolve_path("template_dir"), self.selected_file) if self.selected_file else None
        if path is None:
            messagebox.showerror("선택 필요", "템플릿 파일 목록에서 파일을 먼저 선택해 주세요.", parent=self)
        return path

    def _import_templates(self):
        """다른 곳의 PPT 를 템플릿 폴더로 복사해 목록에 추가 (원본은 그대로)."""
        paths = filedialog.askopenfilenames(parent=self, title="템플릿으로 불러올 PPT 선택",
                                            filetypes=[("PowerPoint", "*.pptx *.pptm *.ppt"), ("모든 파일", "*.*")])
        if not paths:
            return
        tdir = self.cfg.resolve_path("template_dir")
        try:
            tdir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            messagebox.showerror("불러오기", f"템플릿 폴더를 만들 수 없습니다:\n{tdir}\n\n{e}", parent=self)
            return
        added, errors = [], []
        for raw in paths:
            src = Path(raw)
            if src.suffix.lower() not in PPT_EXTS:
                errors.append(f"{src.name}: PPT 파일(.pptx/.pptm/.ppt)이 아닙니다")
                continue
            dest = tdir / src.name
            if dest.resolve() == src.resolve():
                continue  # 이미 템플릿 폴더에 있는 파일
            if dest.exists() and not messagebox.askyesno(
                    "불러오기", f"템플릿 폴더에 '{src.name}' 이 이미 있습니다. 덮어쓸까요?", parent=self):
                continue
            try:
                shutil.copy2(src, dest)
                added.append(src.stem)
            except OSError as e:
                errors.append(f"{src.name}: {e}")
        self._refresh_files()
        if added:
            self._select_file(added[-1])
            self._set_status(f"템플릿 {len(added)}개를 불러왔습니다: {', '.join(added)}")
        if errors:
            messagebox.showerror("불러오기", "다음 파일을 불러오지 못했습니다:\n\n" + "\n".join(errors), parent=self)

    def _open_selected_template(self):
        if (path := self._selected_template()) is not None:
            self._open_ppt(path)

    def _delete_selected_template(self):
        path = self._selected_template()
        if path is None:
            return
        if not messagebox.askyesno("템플릿 삭제", f"'{path.name}' 을(를) 템플릿 폴더에서 삭제할까요?\n"
                                   "파일은 휴지통으로 이동합니다.", icon="warning", parent=self):
            return
        if (err := _send_to_recycle_bin(path)) is not None:
            messagebox.showerror("템플릿 삭제", err, parent=self)
            return
        self.selected_file = None
        self._refresh_files()
        self._set_status(f"삭제됨(휴지통): {path.name}")

    def _add_selected_file(self):
        if not self.selected_file:
            messagebox.showerror("선택 필요", "추가할 템플릿 파일을 먼저 선택해 주세요.", parent=self)
            return
        self._add_template(self.selected_file)

    def _choose_folder(self, key: str):
        path = filedialog.askdirectory(parent=self, initialdir=self.path_vars[key].get(), mustexist=True)
        if not path:
            return
        try:  # 임시 폴더 안을 고르면 임시 파일 정리 때 그 안의 PPT 가 지워질 수 있음
            in_temp = Path(path).resolve().is_relative_to(temp_result_dir().resolve())
        except OSError:
            in_temp = False
        if in_temp:
            messagebox.showerror("폴더 지정", f"임시 파일 폴더 안은 지정할 수 없습니다 (오래된 PPT 가 자동으로 지워짐):\n{path}",
                                 parent=self)
            return
        if key == "output_dir":
            probe = Path(path) / "_.pptx"  # 이 폴더에 결과를 저장해도 원본을 덮어쓰지 않는지 미리 확인
            if (err := output_conflict(self.cfg, probe)) is not None:
                messagebox.showerror("출력 폴더", err, parent=self)
                return
        self.cfg.set("paths", key, path)
        self.path_vars[key].set(str(self.cfg.resolve_path(key)))
        if key == "template_dir":
            self._refresh_files()

    def _open_result(self, path, show: bool | None = None) -> bool:
        """생성된 PPT 열기. 슬라이드 쇼(show, 생략 시 하단 옵션)면 PowerPoint 의 Show 동작(POWERPNT.EXE /s)으로 바로 쇼 시작."""
        if self.slideshow.get() if show is None else show:
            try:
                os.startfile(path, "show")  # .pptx 에 등록된 Show 동사 = 탐색기 우클릭 [쇼]
                return True
            except OSError as e:
                exe = _powerpoint_exe()
                if exe:
                    try:
                        subprocess.Popen([str(exe), "/s", str(path)])
                        return True
                    except OSError:
                        pass
                self._log(f"⚠ 슬라이드 쇼로 열지 못해 일반 열기로 대신합니다: {e}")
        try:
            os.startfile(path)
        except OSError as e:
            messagebox.showerror("열기 실패", f"생성된 파일을 열 수 없습니다.\n{path}\n\n{e}", parent=self)
            return False
        return True

    # ---- 임시 실행으로 만든 PPT (worship_builder.temp_result_dir)
    def _show_temp_files(self):
        if (err := temp_dir_conflict(self.cfg)) is not None:  # 목록·정리가 사용자 PPT 를 다루지 않도록
            messagebox.showerror("임시 파일", err, parent=self)
            return
        TempFilesDialog(self, list_temp_results(), clean_temp_keep(self.cfg.get("ppt", "temp_keep")),
                        on_open=lambda p: self._open_temp(p, False), on_show=lambda p: self._open_temp(p, True),
                        on_save=self._save_temp_as, on_folder=self._open_temp_folder).show()

    def _temp_file_ok(self, path: Path) -> bool:
        """목록을 연 뒤 지워졌거나(보관 개수 정리·직접 삭제) 임시 폴더 밖 파일이면 거부."""
        try:
            ok = (path.suffix.lower() == ".pptx" and path.is_file() and not path.is_symlink()
                  and path.resolve().parent == temp_result_dir().resolve())
        except OSError:
            ok = False
        if not ok:
            messagebox.showerror("임시 파일", f"파일을 찾을 수 없습니다 (이미 정리되었을 수 있음):\n{path.name}", parent=self)
        return ok

    def _open_temp(self, path: Path, show: bool) -> bool:
        if self._busy_merging() or not self._temp_file_ok(path) or not self._open_result(path, show):
            return False
        self._set_status(f"임시 파일을 열었습니다: {path.stem}")
        return True

    def _save_temp_as(self, path: Path) -> bool:
        """임시 파일을 원하는 곳에 복사해 저장 (기본: 출력 폴더, _시분초 를 뗀 이름)."""
        if not self._temp_file_ok(path):
            return False
        out_dir = self.cfg.resolve_path("output_dir")
        dest = filedialog.asksaveasfilename(parent=self, title="임시 파일 저장", defaultextension=".pptx",
                                            initialdir=str(out_dir) if out_dir.is_dir() else None,
                                            initialfile=saved_name_for_temp(path),
                                            filetypes=[("PowerPoint", "*.pptx")])
        if not dest:
            return False
        dest = Path(dest)
        if dest.suffix.lower() != ".pptx":  # 다른 확장자로 저장하면 PowerPoint 가 열지 못함
            dest = dest.with_name(dest.name + ".pptx")
        if not dest.stem or sanitize_filename(dest.stem) != dest.stem:  # CON·NUL 같은 장치 이름, 너무 긴 이름 등
            messagebox.showerror("저장할 수 없습니다", f"쓸 수 없는 파일 이름입니다:\n{dest.name}", parent=self)
            return False
        try:
            in_temp = dest.resolve().parent == temp_result_dir().resolve()
        except OSError:
            in_temp = False
        if in_temp:  # 임시 폴더 안이면 보관 개수 정리 때 지워짐
            messagebox.showerror("저장할 수 없습니다", "임시 폴더 안에는 저장할 수 없습니다 (오래되면 자동으로 지워짐).\n"
                                 "출력 폴더 등 다른 곳을 골라 주세요.", parent=self)
            return False
        if (err := output_conflict(self.cfg, dest)) is not None:
            messagebox.showerror("저장할 수 없습니다", err, parent=self)
            return False
        try:
            shutil.copy2(path, dest)
        except OSError as e:  # 같은 이름의 파일이 PowerPoint 에서 열려 있는 경우 등
            messagebox.showerror("저장 실패", f"저장하지 못했습니다.\n{dest}\n\n{e}", parent=self)
            return False
        self._set_status(f"저장했습니다: {dest}")
        return True

    def _open_temp_folder(self):
        folder = temp_result_dir()
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        if folder.is_dir():
            os.startfile(folder)

    def _open_folder(self, key: str):
        path = self.cfg.resolve_path(key)
        try:
            path.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        # 설정 파일이 고쳐져 폴더 대신 실행 파일 등을 가리키면 startfile 이 그것을 실행하므로 폴더일 때만 연다
        if not path.is_dir():
            messagebox.showerror("폴더 열기", f"폴더를 찾을 수 없습니다:\n{path}", parent=self)
            return
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
        bible, bible2 = self.bible_field.get().strip(), self.bible2_field.get().strip()
        if bible2 and not bible and not (self.cfg.get("bible2ppt", "bible") or "").strip():  # 역본1 칸이 비면 기본 역본
            messagebox.showerror("입력 오류", "두 번째 역본을 쓰려면 첫 번째 역본도 지정해 주세요\n"
                                 "(또는 ⚙ 설정에서 기본 성경 역본을 지정).", parent=self)
            return
        item = make_item(KIND_BIBLE, "말씀", ", ".join(refs), {"bible": bible, "bible2": bible2})
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
        for field in (self.bible_field, self.bible2_field):
            field.set_values(values)
            cur = field.get().strip()
            # 직접 입력한 이름(예전 형식)은 목록의 같은 이름 항목으로 맞춤. 목록에 없는 이름·빈 칸은 그대로 둔다
            # (빈 역본1 = 기본 역본 사용, 빈 역본2 = 한 역본만 — 사용자가 지운 칸을 기본값으로 다시 채우지 않음)
            if cur and (match := match_bible(cur, values)) and match != cur:
                field.set(match)

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
        self._set_order_scroll_speed()
        self._render_list()

    def _set_mode(self, label: str):
        mode = MODES[label]
        ctk.set_appearance_mode(mode)
        self.cfg.set("appearance_mode", None, mode)

    def _open_settings(self):
        keys = ("bible", "bible2")
        old = {k: self.cfg.get("bible2ppt", k) or "" for k in keys}
        if SettingsDialog(self, self.cfg).show():
            self._set_bible_values(list(self.cfg.get("bible2ppt", "bible_list") or []))
            # 기본 역본을 바꿨으면, 그 기본값을 그대로 쓰던(비었거나 예전 기본값인) 칸만 새 기본값으로
            for field, k in zip((self.bible_field, self.bible2_field), keys):
                new = self.cfg.get("bible2ppt", k) or ""
                if new != old[k] and field.get().strip() in ("", old[k]):
                    field.set(new)
            self._update_bible_status()
            self._load_bible_list_if_needed()  # exe 를 바꿨는데 설정 창에서 목록을 못 받았으면 다시 시도
            # 보관 개수를 줄였으면 바로 정리 (임시 폴더가 원본 폴더와 겹치면 정리하지 않음)
            removed = [] if temp_dir_conflict(self.cfg) else prune_temp_results(self.cfg.get("ppt", "temp_keep"))
            self._set_status("설정을 저장했습니다." + (f" 오래된 임시 파일 {len(removed)}개를 지웠습니다." if removed else ""))

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
        if self._bible_preview:  # Bible2PPT 와 병합이 동시에 PowerPoint 를 쓰지 않도록
            messagebox.showerror("잠시 후 다시", "말씀 PPT 를 만드는 중입니다. 끝난 뒤 다시 시도해 주세요.", parent=self)
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
        temp = self.temp_mode.get()  # 생성 중에 체크를 바꿔도 이번 생성은 시작할 때의 모드로
        out = self.cfg.resolve_path("output_dir") / output_filename(date, church, worship, fmt)
        if temp and (err := temp_dir_conflict(self.cfg)) is not None:
            messagebox.showerror("임시 실행을 할 수 없습니다", err, parent=self)
            return
        if not temp:  # 임시 실행은 매번 새 이름(_시분초)으로 임시 폴더에 만들므로 덮어쓰기 확인이 필요 없음
            if (err := output_conflict(self.cfg, out)) is not None:
                messagebox.showerror("저장할 수 없습니다", err, parent=self)
                return
            if out.exists() and not messagebox.askyesno("덮어쓰기", f"이미 같은 이름의 파일이 있습니다.\n{out.name}\n\n덮어쓸까요?",
                                                         parent=self):
                return

        self.building = True
        self.build_btn.configure(state="disabled", text="생성 중…")
        self.log_lines.clear()
        self._log(f"생성 시작: {out.stem} (임시 실행 — 저장하지 않음)" if temp else f"생성 시작: {out.name}")
        self._set_progress(0, "준비 중…")
        items = [ServiceItem.from_dict(i.to_dict()) for i in self.items]  # 작업 중 UI 변경과 분리
        # 열기·쇼 여부도 시작할 때 정함: 병합에 쓴 PowerPoint 에서 결과를 바로 보여 주므로(끄고 다시 켜지 않음) 빠름
        opens = temp or self.open_after.get()
        req = BuildRequest(items, church, date, worship, fmt, temp=temp,
                           open_mode=("show" if self.slideshow.get() else "open") if opens else "")

        def done(result, err):
            self.building = False
            self.build_btn.configure(state="normal", text=self._build_label())
            if err:
                self._set_progress(0, "생성 실패")
                self._log(f"오류: {err}")
                messagebox.showerror("생성 실패", str(err) if isinstance(err, BuildError) else f"{type(err).__name__}: {err}",
                                     parent=self)
                return
            warns = [l for l in self.log_lines if "⚠" in l]
            warn_text = f"⚠ 경고 {len(warns)}건 — [로그]에서 확인" if warns else ""
            if temp:  # 임시 실행: 항상 바로 열기. 나중에 [임시 파일]에서 다시 열거나 저장
                self._set_progress(1, f"임시 실행: {result.stem}", "  ·  ".join(
                    x for x in ("저장하려면 [임시 파일]에서 다른 이름으로 저장", warn_text) if x))
                if not req.opened:
                    self._open_result(result, req.open_mode == "show")
                return
            self._set_progress(1, f"완료: {result.name}", warn_text)
            if opens:
                if not req.opened:  # 창 없는 병합 방식 등으로 바로 보여 주지 못했으면 파일로 연다
                    self._open_result(result, req.open_mode == "show")
            elif messagebox.askyesno("생성 완료", f"생성되었습니다.\n{result}\n\n폴더를 열까요?", parent=self):
                os.startfile(result.parent)

        run_in_thread(lambda: builder.build(req), done)

    # ================================================================ 상태 / 로그
    def _set_status(self, text: str, detail: str = ""):
        """text: 상태 한 줄, detail: 진행 막대 아랫줄의 부연 설명 (비우면 지움)."""
        self.status.configure(text=text)
        self.status_detail.configure(text=detail)

    def _set_progress(self, value: float, text: str, detail: str = ""):
        self.progress.set(max(0.0, min(1.0, value)))
        self._set_status(text, detail)

    def _log(self, msg: str):
        self.log_lines.append(f"[{dt.datetime.now():%H:%M:%S}] {msg}")
        if "⚠" in msg:
            self._set_status(msg.strip())

    def _show_help(self):
        HelpDialog(self).show()

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
