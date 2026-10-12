"""재사용 UI 위젯: 테마 상수, 달력 팝업, 항목 입력 다이얼로그, 설정 창, 드래그&드롭."""
from __future__ import annotations

import calendar
import copy
import datetime as dt
import os
import queue
import re
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any, Callable

import customtkinter as ctk

from bible_bridge import Bible2PPT, bible_sources, filter_bibles, match_bible, search_names, short_source, split_bible
from models import (BIBLE_KEYS, CUSTOM_TEXT_KEYS, DEFAULT_FILENAME_FORMAT, FILENAME_FIELDS, KIND_BIBLE, KIND_CCM, KIND_CUSTOM,
                    KIND_HYMN, KIND_LABELS, KIND_PRAYER, KIND_SERMON, MAX_FILL, ServiceItem, ccm_title, make_item,
                    numbered_token, validate_item)
from config_manager import TEMP_KEEP_MAX, clean_temp_keep
from worship_builder import (PPT_EXTS, count_ppt_files, folder_conflict, output_conflict, output_filename, temp_dir_conflict,
                             validate_filename_format)

APPEARANCE_MODES = {"라이트": "Light", "다크": "Dark", "시스템": "System"}  # 화면 표시 → config.appearance_mode
SLIDE_SIZE_LABELS = {"원본": "original", "16:9": "16:9", "4:3": "4:3"}  # 화면 표시 → config.ppt.slide_size
SETTINGS_FOLDERS = (  # (설정 키, 이름) — ⚙ 설정의 '폴더' (설정 마법사는 WIZARD_FOLDERS)
    ("template_dir", "템플릿"), ("prayer_dir", "기도"), ("hymn_dir", "찬송가"), ("ccm_dir", "CCM"),
    ("output_dir", "출력 (결과 저장)"),
)


def open_folder(master, path: Path) -> None:
    """폴더를 탐색기로 연다. 설정이 고쳐져 폴더 대신 실행 파일 등을 가리키면 startfile 이 그것을 실행하므로 폴더일 때만."""
    try:
        path.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    if not path.is_dir():
        messagebox.showerror("폴더 열기", f"폴더를 찾을 수 없습니다:\n{path}", parent=master)
        return
    os.startfile(path)

# ---------------------------------------------------------------- Win11 Fluent 톤 색상 (light, dark)
FONT_FAMILY = "Malgun Gothic"
# 강조 색은 apply_accent() 가 Windows 강조 색으로 바꿀 수 있도록 리스트(같은 개체를 제자리에서 바꿈 — 다른 모듈이
# from ui_widgets import ACCENT 로 가져간 이름도 같이 바뀜). 위젯을 만들기 전에 정해야 한다 (CTk 는 만들 때 색을 읽음).
DEFAULT_ACCENT = ("#0067C0", "#2B88D8")
DEFAULT_ACCENT_HOVER = ("#1975C5", "#3C95E0")
ACCENT = list(DEFAULT_ACCENT)
ACCENT_HOVER = list(DEFAULT_ACCENT_HOVER)
BG = ("#F3F3F3", "#202020")
CARD = ("#FFFFFF", "#2B2B2B")
SUBTLE = ("#F9F9F9", "#323232")
BORDER = ("#E3E3E3", "#3C3C3C")
PILL_BORDER = ("#BDBDBD", "#555555")
TEXT = ("#1A1A1A", "#F3F3F3")
TEXT2 = ("#616161", "#A8A8A8")
BTN = ("#FBFBFB", "#383838")
BTN_HOVER = ("#F0F0F0", "#424242")
DANGER = ("#C42B1C", "#FF99A4")
KIND_COLORS = {
    KIND_PRAYER: ("#0F7B0F", "#6CCB5F"),
    KIND_HYMN: ("#8764B8", "#B69CE0"),
    KIND_BIBLE: ("#C35A00", "#F7A35C"),
    KIND_SERMON: ("#B4235A", "#F28AB0"),
    KIND_CCM: ("#00787A", "#4CC9C0"),
    KIND_CUSTOM: ("#0067C0", "#60CDFF"),
}


MERGE_MODES = {
    "원본 서식 유지 (권장)": "source_formatting",
    "창 없이 병합 (배경은 복제/이미지화)": "windowless",
}


def _luminance(hex_color: str) -> float:
    """WCAG 상대 휘도."""
    def ch(v: int) -> float:
        c = v / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (1, 3, 5))
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def _contrast_with_white(hex_color: str) -> float:
    return 1.05 / (_luminance(hex_color) + 0.05)


def _blend(hex_color: str, toward: str, t: float) -> str:
    a = [int(hex_color[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(toward[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))


def windows_accent_palette() -> list[str] | None:
    """Windows 강조 색 팔레트 (밝은 것 → 어두운 것 7단계: Light3, Light2, Light1, 강조 색, Dark1, Dark2, Dark3).
    HKCU\\…\\Explorer\\Accent\\AccentPalette (색마다 RGBA 4바이트, 8개 중 마지막은 강조 색과 무관). 읽을 수 없으면 None."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\Accent") as k:
            data, kind = winreg.QueryValueEx(k, "AccentPalette")
        if kind != winreg.REG_BINARY or len(data) < 28:
            return None
        return [f"#{data[i]:02X}{data[i + 1]:02X}{data[i + 2]:02X}" for i in range(0, 28, 4)]
    except (OSError, ImportError):
        return None


def pick_accent(palette: list[str]) -> tuple[tuple[str, str], tuple[str, str]]:
    """팔레트 → ((라이트, 다크) 강조 색, (라이트, 다크) 마우스 올림 색).
    Win11 앱처럼 라이트 모드는 Dark1, 다크 모드는 강조 색에서 시작하되, 강조 버튼 글자가 흰색이므로 흰 글자와의
    명암비가 4.5:1 이상이 될 때까지 더 어두운 단계로 내려간다 (노랑·연두처럼 밝은 강조 색에서도 글자가 읽히게).
    끝까지 모자라면 가장 어두운 단계를 더 어둡게 섞는다. 마우스 올림 색은 흰색 쪽으로 10% 섞은 색."""
    def readable(start: int) -> str:
        for c in palette[start:7]:
            if _contrast_with_white(c) >= 4.5:
                return c
        c = palette[6]
        for _ in range(10):
            if _contrast_with_white(c) >= 4.5:
                break
            c = _blend(c, "#000000", 0.15)
        return c
    light, dark = readable(4), readable(3)
    return (light, dark), (_blend(light, "#FFFFFF", 0.1), _blend(dark, "#FFFFFF", 0.1))


def apply_accent(use_windows: bool) -> bool:
    """강조 색을 Windows 강조 색(use_windows) 또는 기본 파랑으로. 위젯을 만들기 전에 부를 것. Windows 색을 썼으면 True."""
    palette = windows_accent_palette() if use_windows else None
    accent, hover = pick_accent(palette) if palette else (DEFAULT_ACCENT, DEFAULT_ACCENT_HOVER)
    ACCENT[:] = accent
    ACCENT_HOVER[:] = hover
    return palette is not None


_FONT_CACHE: dict[tuple[int, str], ctk.CTkFont] = {}


def font(size: int = 13, weight: str = "normal") -> ctk.CTkFont:
    """(크기, 굵기)별로 CTkFont 를 한 번만 만들어 재사용한다.
    매번 새로 만들면 목록을 다시 그릴 때마다 Tcl 명명 폰트가 쌓이고, GC 가 이를 삭제(font delete)하는
    시점이 위젯 생성과 겹치면 Tk 가 접근 위반으로 종료될 수 있다."""
    key = (size, weight)
    f = _FONT_CACHE.get(key)
    if f is None:
        f = _FONT_CACHE[key] = ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight)
    return f


def secondary_button(master, text: str, command=None, width: int = 72, **kw) -> ctk.CTkButton:
    kw.setdefault("height", 32)
    return ctk.CTkButton(master, text=text, command=command, width=width, corner_radius=6,
                         fg_color=BTN, hover_color=BTN_HOVER, border_width=1, border_color=BORDER,
                         text_color=TEXT, font=font(13), **kw)


def accent_button(master, text: str, command=None, width: int = 72, **kw) -> ctk.CTkButton:
    kw.setdefault("height", 32)
    kw.setdefault("font", font(13, "bold"))
    return ctk.CTkButton(master, text=text, command=command, width=width, corner_radius=6,
                         fg_color=ACCENT, hover_color=ACCENT_HOVER, text_color="#FFFFFF", **kw)


def card(master, **kw) -> ctk.CTkFrame:
    return ctk.CTkFrame(master, corner_radius=10, fg_color=CARD, border_width=1, border_color=BORDER, **kw)


def _is_dark() -> bool:
    return ctk.get_appearance_mode() == "Dark"


def _pick(color: tuple[str, str]) -> str:
    return color[1] if _is_dark() else color[0]


mode_color = _pick  # (라이트, 다크) 색 중 현재 모드 색 — 일반 tk 위젯용


def scaled(widget: tk.Misc, value: float) -> int:
    """CTk 논리 크기 → 실제 픽셀 (일반 tk 위젯은 CTk 처럼 배율을 자동 적용하지 않음)."""
    return round(value * ctk.ScalingTracker.get_widget_scaling(widget))


def tk_font(widget: tk.Misc, size: float = 13, weight: str = "normal") -> tuple:
    """일반 tk 위젯용 글꼴 — CTk 위젯의 font(size) 와 같은 크기(음수 = 픽셀)."""
    return (FONT_FAMILY, -scaled(widget, size), weight)


class _Modal(ctk.CTkFrame):
    """메인 창 안에 뜨는 모달 다이얼로그 (Win11 ContentDialog 스타일).

    별도 창(CTkToplevel)을 쓰지 않는다. CTkToplevel 은 생성 시 창 숨김/강제 update/ctypes 타이틀바 조작과
    해제되지 않는 bind_all 을 수행하는데, 반복 생성·파괴 시 Tk 접근 위반 크래시가 재현되었다.
    다이얼로그에 grab 을 걸어 바깥 위젯 클릭을 막고, 메인 화면은 뒤에 그대로 보이게 한다.
    사용: result = SomeDialog(app, ...).show()   (닫힐 때까지 대기)

    다이얼로그 위에 다이얼로그를 띄울 수 있다(예: 항목 입력 → 역본 검색). Esc/Enter 는 창 단위로 바인딩되므로
    맨 위 다이얼로그만 반응하도록 _stack 으로 확인하고, 닫힐 때 아래 다이얼로그의 grab 을 되돌린다.
    """
    _stack: list["_Modal"] = []

    def __init__(self, master, title: str):
        self._win = master.winfo_toplevel()
        super().__init__(self._win, corner_radius=12, fg_color=BG, border_width=2, border_color=PILL_BORDER)
        self.result: Any = None
        self._closing = False
        self._prev_grab = None
        ctk.CTkLabel(self, text=title, font=font(16, "bold"), text_color=TEXT, anchor="w").pack(
            fill="x", padx=20, pady=(16, 0))
        self._keys = [("<Escape>", self._win.bind("<Escape>", lambda e: self._if_top(self.destroy), add="+"))]

    def _if_top(self, fn):
        if _Modal._stack and _Modal._stack[-1] is self:
            return fn()

    def show(self):
        self.place(relx=0.5, rely=0.45, anchor="center")
        self.lift()
        self.update_idletasks()
        self._prev_grab = self.grab_current()
        _Modal._stack.append(self)
        try:
            self.grab_set()  # 다이얼로그 바깥 위젯 클릭 차단
        except tk.TclError:
            pass
        self.after(30, self.on_shown)
        self._win.wait_window(self)
        return self.result

    def bind_return(self, fn):
        self._keys.append(("<Return>", self._win.bind("<Return>", lambda e: self._if_top(fn), add="+")))

    def on_shown(self):
        pass

    def destroy(self):
        if self._closing:
            return
        self._closing = True
        if self in _Modal._stack:
            _Modal._stack.remove(self)
        for seq, funcid in self._keys:
            _unbind(self._win, seq, funcid)
        try:
            self.grab_release()
        except tk.TclError:
            pass
        prev = self._prev_grab
        super().destroy()
        if prev is not None:
            try:
                if prev.winfo_exists():
                    prev.grab_set()  # 아래 다이얼로그로 grab 복귀
            except tk.TclError:
                pass


WARN_BG = ("#FFF4CE", "#433519")  # Win11 InfoBar '경고' 배경
WARN_ICON = ("#9D5D00", "#FCE100")


class BuildOverlay(ctk.CTkFrame):
    """생성(병합) 중 화면. 제목 줄 아래 전체를 덮어 순서·설정을 못 바꾸게 하고, PowerPoint 를 쓰지 말라고 알린다.
    [원본 서식 유지] 붙여넣기(ExecuteMso)는 PowerPoint 의 '활성 창'에 붙여넣으므로, 그사이 사용자가 PowerPoint 에서
    다른 파일을 열거나 누르면 슬라이드가 그 파일에 붙을 수 있다 (프로그램 밖 조작은 막을 수 없어 안내만).
    _Modal 과 달리 grab 을 걸지 않는다 — 제목 줄(창 옮기기·최소화·닫기)은 생성 중에도 쓸 수 있어야 하므로 덮기만 함.
    사용: App 이 _content 의 1~3행에 grid 로 겹쳐 놓고 set_progress 로 갱신, 끝나면 destroy."""

    STALL_SEC = 60  # 진행이 이만큼 없으면 'PowerPoint 가 응답하지 않음' 안내 (set_progress 의 stall_sec 로 단계별 조정)

    def __init__(self, master, opens: bool, on_show_ppt: Callable[[], None] | None = None,
                 on_cancel: Callable[[], None] | None = None):
        super().__init__(master, fg_color=BG, corner_radius=0)
        self._box = box = card(self)
        box.place(relx=0.5, rely=0.42, anchor="center")
        ctk.CTkLabel(box, text="예배 PPT 를 만드는 중입니다", font=font(18, "bold"), text_color=TEXT, anchor="w").pack(
            fill="x", padx=28, pady=(24, 0))
        self._step = ctk.CTkLabel(box, text="준비 중…", font=font(13), text_color=TEXT2, anchor="w")
        self._step.pack(fill="x", padx=28, pady=(10, 0))
        self._bar = ctk.CTkProgressBar(box, width=460, height=8, progress_color=ACCENT)
        self._bar.set(0)
        self._bar.pack(fill="x", padx=28, pady=(8, 0))
        self._time = ctk.CTkLabel(box, text="", font=font(12), text_color=TEXT2, anchor="e")
        self._time.pack(fill="x", padx=28, pady=(4, 0))

        self._warn = warn = ctk.CTkFrame(box, fg_color=WARN_BG, corner_radius=8)
        warn.pack(fill="x", padx=28, pady=(12, 26))
        ctk.CTkLabel(warn, text="⚠", font=font(22, "bold"), text_color=WARN_ICON).pack(
            side="left", anchor="n", padx=(14, 10), pady=10)
        body = ctk.CTkFrame(warn, fg_color="transparent")
        body.pack(side="left", fill="x", expand=True, padx=(0, 14), pady=12)
        ctk.CTkLabel(body, text="다 만들 때까지 PowerPoint 를 사용하지 마세요", font=font(14, "bold"), text_color=TEXT,
                     anchor="w").pack(fill="x")
        after = "PowerPoint 창은 다 만든 뒤 자동으로 열립니다." if opens else "다 만들면 이 화면이 닫히고 알려 드립니다."
        ctk.CTkLabel(body, text="만드는 동안 PowerPoint 에서 다른 파일을 열거나 누르면 슬라이드가 그 파일에 "
                                f"붙을 수 있습니다.\n{after}",
                     font=font(12), text_color=TEXT, anchor="w", justify="left", wraplength=380).pack(fill="x", pady=(4, 0))

        # 진행이 오래 없을 때만 보이는 안내 (처음 설치한 Office 의 라이선스 동의·로그인 창 등을 PowerPoint 가 기다리는 경우)
        self._stall = ctk.CTkFrame(box, fg_color=SUBTLE, corner_radius=8)
        self._stall_text = ctk.CTkLabel(self._stall, text="", font=font(12), text_color=TEXT, anchor="w",
                                        justify="left", wraplength=400)
        self._stall_text.pack(fill="x", padx=14, pady=(10, 6))
        btns = ctk.CTkFrame(self._stall, fg_color="transparent")
        btns.pack(fill="x", padx=14, pady=(0, 12))
        if on_show_ppt is not None:
            secondary_button(btns, "PowerPoint 창 보기", on_show_ppt, width=130).pack(side="left")
        if on_cancel is not None:
            stop = secondary_button(btns, "생성 중단", on_cancel, width=90)
            stop.configure(text_color=DANGER)
            stop.pack(side="left", padx=(8, 0))
        self._stall_sec = self.STALL_SEC
        self._last_progress = time.monotonic()

        # 덮인 위젯으로 Tab 이동 금지 (Tk 는 가려진 위젯도 '보이는' 것으로 보고 포커스를 옮겨, 생성 중에 입력칸에 글자가 들어감)
        for seq in ("<Tab>", "<Shift-Tab>"):
            self.bind(seq, lambda e: "break")
        self._start = dt.datetime.now()
        self._tick_job = None
        self._tick()

    def _tick(self):
        sec = int((dt.datetime.now() - self._start).total_seconds())
        self._time.configure(text=f"경과 {sec // 60}:{sec % 60:02d}")
        idle = time.monotonic() - self._last_progress
        if idle >= self._stall_sec and not self._stall.winfo_ismapped():
            wait = f"{int(idle) // 60}분" if idle >= 60 else f"{int(idle)}초"
            self._stall_text.configure(
                text=f"⏳ {wait} 넘게 진행이 없습니다. PowerPoint 가 다른 창의 응답을 기다리고 있을 수 있습니다 "
                     "(처음 설치한 Office 의 라이선스 동의·로그인 창, PowerPoint 오류 창 등).\n"
                     "[PowerPoint 창 보기]로 떠 있는 창을 확인해 닫으면 계속됩니다.")
            self._warn.pack_configure(pady=(12, 10))
            self._stall.pack(fill="x", padx=28, pady=(0, 26), after=self._warn)
        self._tick_job = self.after(1000, self._tick)

    def set_progress(self, value: float, text: str, stall_sec: float | None = None) -> None:
        self._bar.set(max(0.0, min(1.0, value)))
        self._step.configure(text=text)
        self._last_progress = time.monotonic()
        self._stall_sec = stall_sec or self.STALL_SEC
        if self._stall.winfo_ismapped():  # 다시 진행되면 안내를 거둠
            self._stall.pack_forget()
            self._warn.pack_configure(pady=(12, 26))

    def set_step(self, text: str) -> None:
        self._step.configure(text=text)

    def destroy(self):
        if self._tick_job is not None:
            self.after_cancel(self._tick_job)
            self._tick_job = None
        super().destroy()


def _unbind(widget: tk.Misc, seq: str, funcid: str) -> None:
    """seq 에 묶인 여러 콜백 중 funcid 하나만 해제.
    (Python 3.12 이하의 Misc.unbind(seq, funcid) 는 seq 의 바인딩을 전부 지워 아래 다이얼로그의 Esc/Enter 까지 사라짐)"""
    try:
        script = widget.tk.call("bind", widget._w, seq) or ""
        keep = "\n".join(line for line in str(script).split("\n") if funcid not in line)
        widget.tk.call("bind", widget._w, seq, keep)
        widget.deletecommand(funcid)
    except tk.TclError:
        pass


# ---------------------------------------------------------------- 달력 팝업
class CalendarPopup(ctk.CTkFrame):
    """날짜 입력칸 아래에 뜨는 달력 (메인 창 안의 플로팅 패널). 바깥 클릭·Esc 로 닫힘."""
    WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"]

    def __init__(self, master, anchor: tk.Widget, initial: dt.date | None, on_pick: Callable[[dt.date], None]):
        self._win = master.winfo_toplevel()
        super().__init__(self._win, fg_color=CARD, corner_radius=10, border_width=1, border_color=PILL_BORDER)
        self.on_pick = on_pick
        self.anchor = anchor
        self.today = dt.date.today()
        self.selected = initial
        d = initial or self.today
        self.year, self.month = d.year, d.month
        self._closed = False

        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", padx=10, pady=(10, 4))
        ctk.CTkButton(head, text="‹", width=30, height=28, fg_color="transparent", hover_color=BTN_HOVER,
                      text_color=TEXT, font=font(16), command=lambda: self._move(-1)).pack(side="left")
        self.title_lbl = ctk.CTkLabel(head, text="", font=font(14, "bold"), text_color=TEXT)
        self.title_lbl.pack(side="left", expand=True)
        ctk.CTkButton(head, text="›", width=30, height=28, fg_color="transparent", hover_color=BTN_HOVER,
                      text_color=TEXT, font=font(16), command=lambda: self._move(1)).pack(side="right")
        self.grid_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.grid_frame.pack(padx=10, pady=(0, 6))
        foot = ctk.CTkFrame(self, fg_color="transparent")
        foot.pack(fill="x", padx=10, pady=(0, 10))
        secondary_button(foot, "오늘", lambda: self._pick(self.today), width=60).pack(side="left")
        secondary_button(foot, "닫기", self.destroy, width=60).pack(side="right")

        self._render()
        self.update_idletasks()
        # CTk 의 place 는 좌표에 배율을 곱하므로 실제 픽셀 좌표를 배율로 나눠 전달
        scale = ctk.ScalingTracker.get_widget_scaling(self)
        x = anchor.winfo_rootx() - self._win.winfo_rootx() + anchor.winfo_width() - self.winfo_reqwidth()
        y = anchor.winfo_rooty() - self._win.winfo_rooty() + anchor.winfo_height() + 4
        self.place(x=max(x, 0) / scale, y=y / scale)
        self.lift()
        # 팝업 바깥 클릭 / Esc → 닫기 (창 전체에 임시 바인딩, 닫을 때 해제)
        self._binds = [("<Button-1>", self._win.bind("<Button-1>", self._on_click_anywhere, add="+")),
                       ("<Escape>", self._win.bind("<Escape>", lambda e: self.destroy(), add="+"))]

    def _on_click_anywhere(self, event):
        x, y = event.x_root, event.y_root

        def inside(w):
            return (w.winfo_rootx() <= x <= w.winfo_rootx() + w.winfo_width()
                    and w.winfo_rooty() <= y <= w.winfo_rooty() + w.winfo_height())

        if not inside(self) and not inside(self.anchor):
            self.after_idle(self.destroy)

    def destroy(self):
        if self._closed:
            return
        self._closed = True
        for seq, funcid in getattr(self, "_binds", []):
            try:
                self._win.unbind(seq, funcid)
            except tk.TclError:
                pass
        super().destroy()

    def _move(self, delta: int):
        m = self.month + delta
        self.year += (m - 1) // 12
        self.month = (m - 1) % 12 + 1
        self._render()

    def _pick(self, d: dt.date):
        self.on_pick(d)
        self.destroy()

    def _render(self):
        for w in self.grid_frame.winfo_children():
            w.destroy()
        self.title_lbl.configure(text=f"{self.year}년 {self.month}월")
        for c, wd in enumerate(self.WEEKDAYS):
            color = DANGER if c == 0 else (("#0067C0", "#60CDFF") if c == 6 else TEXT2)
            ctk.CTkLabel(self.grid_frame, text=wd, width=36, height=24, font=font(12, "bold"),
                         text_color=color).grid(row=0, column=c)
        cal = calendar.Calendar(firstweekday=6)  # 일요일 시작
        for r, week in enumerate(cal.monthdatescalendar(self.year, self.month), start=1):
            for c, d in enumerate(week):
                in_month = d.month == self.month
                is_sel = d == self.selected
                is_today = d == self.today
                fg = ACCENT if is_sel else "transparent"
                tc = "#FFFFFF" if is_sel else (TEXT if in_month else TEXT2)
                if not is_sel and in_month and c == 0:
                    tc = DANGER
                ctk.CTkButton(self.grid_frame, text=str(d.day), width=36, height=30, corner_radius=15,
                              fg_color=fg, hover_color=ACCENT_HOVER if is_sel else BTN_HOVER, text_color=tc,
                              border_width=1 if (is_today and not is_sel) else 0, border_color=ACCENT,
                              font=font(12, "bold" if is_today else "normal"),
                              command=lambda d=d: self._pick(d)).grid(row=r, column=c, padx=1, pady=1)


# ---------------------------------------------------------------- 성경 역본 검색
class BiblePicker(_Modal):
    """성경 역본 검색·선택 (이름·약어·ID, 초성 검색, 소스 필터, ↑↓/Enter).
    show() → 선택한 표시 문자열 / '' (기본값 사용) / None (취소).
    역본이 수백 개라 CTk 위젯 대신 tk.Listbox 로 그린다 (CTkButton 수백 개는 그리는 데 수 초 걸림)."""
    MAX_ROWS = 500

    def __init__(self, master, values: list[str], current: str = "", clear_text: str | None = "기본값 사용"):
        super().__init__(master, "성경 역본 선택")
        self.values = list(values)
        self._shown: list[str] = []
        self._source: str | None = None
        self._job = None

        body = card(self)
        body.pack(fill="both", expand=True, padx=16, pady=(12, 8))
        self.search = ctk.CTkEntry(body, width=480, height=34, font=font(13),
                                   placeholder_text="이름 · 약어 · ID 검색   (초성 가능: ㄱㅇㄱㅈ → 개역개정)")
        self.search.pack(fill="x", padx=14, pady=(14, 8))
        sources = bible_sources(self.values)
        if len(sources) > 1:
            seg = ctk.CTkSegmentedButton(body, values=["전체", *sources], font=font(12), command=self._on_source)
            seg.set("전체")
            seg.pack(anchor="w", padx=14, pady=(0, 8))

        box = ctk.CTkFrame(body, fg_color=SUBTLE, corner_radius=8, border_width=1, border_color=BORDER)
        box.pack(fill="both", expand=True, padx=14)
        self.listbox = tk.Listbox(box, height=12, width=60, activestyle="none", borderwidth=0, highlightthickness=0,
                                  exportselection=False, font=tk_font(self, 15), bg=_pick(SUBTLE), fg=_pick(TEXT),
                                  selectbackground=_pick(ACCENT), selectforeground="#FFFFFF")
        sb = ctk.CTkScrollbar(box, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=8)
        sb.pack(side="right", fill="y", padx=2, pady=6)
        self.count = ctk.CTkLabel(body, text="", font=font(11), text_color=TEXT2, anchor="w")
        self.count.pack(fill="x", padx=14, pady=(4, 10))

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=16, pady=(4, 16))
        if clear_text:
            secondary_button(btns, clear_text, self._clear, width=110).pack(side="left")
        accent_button(btns, "선택", self._ok, width=96).pack(side="right")
        secondary_button(btns, "취소", self.destroy, width=96).pack(side="right", padx=8)

        self.search.bind("<KeyRelease>", self._on_key)
        self.search.bind("<Down>", lambda e: self._move(1))
        self.search.bind("<Up>", lambda e: self._move(-1))
        self.search.bind("<Next>", lambda e: self._move(10))
        self.search.bind("<Prior>", lambda e: self._move(-10))
        self.listbox.bind("<Double-Button-1>", lambda e: self._ok())
        self.bind_return(self._ok)

        current = current.strip()
        if current and current not in self.values:  # 직접 입력한 이름이면 그걸로 검색해 보여 줌
            self.search.insert(0, current)
        self._refresh(select=current)

    def on_shown(self):
        self.search.focus_set()

    def _on_key(self, event):
        if event.keysym in ("Up", "Down", "Prior", "Next", "Return", "KP_Enter", "Escape"):
            return
        if self._job:
            self.after_cancel(self._job)
        self._job = self.after(120, self._refresh)  # 한글 조합 중 매 키마다 다시 그리지 않도록

    def _on_source(self, value: str):
        self._source = None if value == "전체" else value
        self._refresh()

    def _refresh(self, select: str | None = None):
        self._job = None
        matches = filter_bibles(self.values, self.search.get(), self._source)
        self._shown = matches[:self.MAX_ROWS]
        self.listbox.delete(0, "end")
        for e in self._shown:
            name, src, bid = split_bible(e)
            self.listbox.insert("end", f"  {name}" + (f"      {short_source(src)} · {bid}" if bid else ""))
        if not self.values:
            text = "역본 목록이 없습니다. ↻ 로 먼저 불러오세요."
        elif not matches:
            text = "검색 결과가 없습니다."
        else:
            text = f"{len(matches)}개" + (f" (처음 {self.MAX_ROWS}개 표시)" if len(matches) > self.MAX_ROWS else "") \
                   + f"  /  전체 {len(self.values)}개   ·   ↑↓ 이동, Enter 선택, 더블클릭 선택"
        self.count.configure(text=text)
        if self._shown:
            self._select(self._shown.index(select) if select in self._shown else 0)

    def _select(self, i: int):
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(i)
        self.listbox.activate(i)
        self.listbox.see(i)

    def _move(self, delta: int):
        if self._shown:
            cur = self.listbox.curselection()
            i = (cur[0] + delta) if cur else 0
            self._select(max(0, min(len(self._shown) - 1, i)))
        return "break"

    def _ok(self):
        if self._job:  # 입력 직후 Enter → 검색 결과부터 반영
            self.after_cancel(self._job)
            self._refresh()
        cur = self.listbox.curselection()
        if not cur:
            self.bell()
            return
        self.result = self._shown[cur[0]]
        self.destroy()

    def _clear(self):
        self.result = ""
        self.destroy()


class BibleField(ctk.CTkFrame):
    """역본 입력칸 + 🔍(역본 검색 창). 직접 입력도 가능 (이름만 입력하면 Bible2PPT 가 이름으로 찾음)."""

    def __init__(self, master, values: list[str] | None = None, width: int = 260, placeholder: str = "",
                 font_size: int = 13, clear_text: str | None = "기본값 사용"):
        super().__init__(master, fg_color="transparent")
        self.values = list(values or [])
        self.clear_text = clear_text
        self.entry = ctk.CTkEntry(self, width=width - 38, height=32, font=font(font_size), placeholder_text=placeholder)
        self.entry.pack(side="left", fill="x", expand=True)
        self.button = secondary_button(self, "🔍", self.open_picker, width=32)
        self.button.pack(side="left", padx=(6, 0))

    def get(self) -> str:
        return self.entry.get()

    def set(self, value: str):
        self.entry.delete(0, "end")
        if value:
            self.entry.insert(0, value)

    def set_values(self, values: list[str]):
        self.values = list(values)

    def open_picker(self):
        res = BiblePicker(self, self.values, current=self.get(), clear_text=self.clear_text).show()
        if res is not None:
            self.set(res)


# ---------------------------------------------------------------- 항목 입력 다이얼로그
class TextPrompt(_Modal):
    """이름 하나 입력 (예: 프리셋으로 저장). validate(값) → 오류 메시지(없으면 None). 결과 = 앞뒤 공백을 뗀 값."""

    def __init__(self, master, title: str, label: str, initial: str = "", hint: str = "",
                 validate: Callable[[str], str | None] | None = None):
        super().__init__(master, title)
        self._validate = validate
        self._prompt_title = title
        body = card(self)
        body.pack(fill="both", expand=True, padx=16, pady=(16, 8))
        ctk.CTkLabel(body, text=label, font=font(13), text_color=TEXT, anchor="w").grid(
            row=0, column=0, padx=16, pady=(14, 6), sticky="w")
        self.entry = ctk.CTkEntry(body, width=260, font=font(13))
        self.entry.insert(0, initial)
        self.entry.grid(row=0, column=1, padx=(0, 16), pady=(14, 6), sticky="w")
        if hint:
            ctk.CTkLabel(body, text=hint, font=font(11), text_color=TEXT2, anchor="w", justify="left").grid(
                row=1, column=0, columnspan=2, padx=16, pady=(0, 12), sticky="w")
        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=16, pady=(4, 16))
        accent_button(btns, "확인", self._ok, width=96).pack(side="right")
        secondary_button(btns, "취소", self.destroy, width=96).pack(side="right", padx=8)
        self.bind_return(self._ok)

    def on_shown(self):
        self.entry.focus_set()
        self.entry.select_range(0, "end")

    def _ok(self):
        value = self.entry.get().strip()
        if self._validate is not None and (err := self._validate(value)):
            messagebox.showerror(self._prompt_title, err, parent=self)
            return
        self.result = value
        self.destroy()


class _SearchList:
    """다이얼로그 안의 [검색칸] + 목록 (CCM 곡·사용자 지정 템플릿 고르기). 공백 AND·초성 검색(search_names).
    항목이 수백 개일 수 있어 CTk 위젯 대신 tk.Listbox 하나로 그림 (BiblePicker 와 같은 이유).
    ↑↓·PgUp/PgDn 이동, 더블클릭 = 다이얼로그 확인(_ok), Enter 는 다이얼로그가 처리. get() = 고른 값 (없으면 '')."""

    def __init__(self, dialog: "ItemDialog", body, row: int, label: str, values: list[str], titles: list[str],
                 current: str, missing_text: str, icon: str, rows: int, empty_text: str,
                 on_pick: Callable[[], None] | None = None):
        self._dialog = dialog
        self._values, self._titles = list(values), list(titles)
        self._icon, self._empty_text, self._on_pick = icon, empty_text, on_pick
        self._shown: list[int] = []
        self._job = None
        dialog._lbl(body, row, label)
        self.entry = ctk.CTkEntry(body, width=400, placeholder_text="🔍  파일 이름 검색 (초성 가능)", font=font(13))
        self.entry.grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")
        box = ctk.CTkFrame(body, fg_color=SUBTLE, corner_radius=8, border_width=1, border_color=BORDER)
        box.grid(row=row + 1, column=0, columnspan=2, sticky="nsew", padx=16, pady=(2, 4))
        self.listbox = tk.Listbox(box, height=rows, width=56, activestyle="none", borderwidth=0,
                                  highlightthickness=0, exportselection=False, font=tk_font(dialog, 14),
                                  bg=_pick(SUBTLE), fg=_pick(TEXT), selectbackground=_pick(ACCENT),
                                  selectforeground="#FFFFFF")
        sb = ctk.CTkScrollbar(box, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=6)
        sb.pack(side="right", fill="y", padx=2, pady=4)
        self.count = ctk.CTkLabel(body, text="", font=font(11), text_color=TEXT2, anchor="w")
        self.count.grid(row=row + 2, column=0, columnspan=2, sticky="w", padx=16)

        self.entry.bind("<KeyRelease>", self._on_key)
        for seq, d in (("<Down>", 1), ("<Up>", -1), ("<Next>", 10), ("<Prior>", -10)):
            self.entry.bind(seq, lambda e, d=d: self._move(d))
        self.listbox.bind("<<ListboxSelect>>", lambda e: self._picked())
        self.listbox.bind("<Double-Button-1>", lambda e: dialog._ok())
        if current and current not in self._values:  # 폴더에서 사라진 파일 → 이름으로 검색해 보여 줌
            self.entry.insert(0, missing_text)
        self.refresh(select=current)

    def _on_key(self, event):
        if event.keysym in ("Up", "Down", "Prior", "Next", "Return", "KP_Enter", "Escape"):
            return
        if self._job:
            self._dialog.after_cancel(self._job)
        self._job = self._dialog.after(120, self.refresh)  # 한글 조합 중 매 키마다 다시 그리지 않도록

    def refresh(self, select: str = ""):
        self._job = None
        self._shown = search_names(self._titles, self.entry.get())
        lb = self.listbox
        lb.delete(0, "end")
        for i in self._shown:
            lb.insert("end", f"  {self._icon}  {self._titles[i]}")
        if not self._values:
            text = self._empty_text
        elif not self._shown:
            text = "검색 결과가 없습니다."
        else:
            text = f"{len(self._shown)}개 / 전체 {len(self._values)}개   ·   ↑↓ 이동, Enter 확인, 더블클릭 선택"
        self.count.configure(text=text)
        if self._shown:
            self._select(next((k for k, i in enumerate(self._shown) if self._values[i] == select), 0))
        else:
            self._picked()

    def _select(self, pos: int):
        lb = self.listbox
        lb.selection_clear(0, "end")
        lb.selection_set(pos)
        lb.activate(pos)
        lb.see(pos)
        self._picked()

    def _picked(self):
        if self._on_pick is not None:
            self._on_pick()

    def _move(self, delta: int):
        if self._shown:
            cur = self.listbox.curselection()
            self._select(max(0, min(len(self._shown) - 1, (cur[0] + delta) if cur else 0)))
        return "break"

    def get(self) -> str:
        if self._job:  # 입력 직후 Enter → 검색 결과부터 반영
            self._dialog.after_cancel(self._job)
            self.refresh()
        cur = self.listbox.curselection()
        return self._values[self._shown[cur[0]]] if cur and self._shown else ""


class ItemDialog(_Modal):
    """기도/찬송/말씀/설교/CCM/사용자 지정 입력. 필수값 미입력 시 에러 창.
    item 이 있으면 수정, initial 이 있으면 그 값으로 채운 새 항목 추가 (예: 템플릿 목록에서 고른 PPT).
    사용자 지정은 내용1~3 칸을 항상 보여 주고, fill_slots(템플릿 이름) → 템플릿에 있는 입력 자리 번호(모르면 None)를 안내한다."""
    CCM_ROWS = 10
    TEMPLATE_ROWS = 7  # 사용자 지정은 아래에 내용1~3 칸이 더 있어 목록을 조금 낮게

    def __init__(self, master, kind: str, item: ServiceItem | None = None,
                 template_names: list[str] | None = None, bible_values: list[str] | None = None,
                 default_bible: str = "", default_bible2: str = "", fallback_bible: str = "", ccm_files: list[str] | None = None,
                 prayer_names: list[str] | None = None, initial: ServiceItem | None = None,
                 fill_slots: Callable[[str], set[int] | None] | None = None, slot_names: list[str] | None = None):
        editing = item is not None
        item = item or initial
        super().__init__(master, f"{KIND_LABELS[kind]} {'수정' if editing else '추가'}")
        self.kind = kind
        self.options = dict(item.options) if item else {}

        body = card(self)
        body.pack(fill="both", expand=True, padx=16, pady=(16, 8))
        ctk.CTkLabel(body, text=f"{KIND_LABELS[kind]} 항목", font=font(15, "bold"),
                     text_color=KIND_COLORS[kind]).grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 8))

        self.label_widget = None
        self.text_entries: dict[int, ctk.CTkEntry] = {}  # 사용자 지정 내용2·3 입력칸
        self._text_labels: dict[int, ctk.CTkLabel] = {}
        self._fill_slots = fill_slots
        self._slot_names = slot_names or [f"{{{{내용{n}}}}}" for n in range(1, MAX_FILL + 1)]
        self._slots_job = None
        self.bible_widget = None
        self.bible2_widget = None
        self._fallback_bible = fallback_bible.strip()  # 역본1 칸이 비었을 때 쓰는 설정의 기본 역본
        self.preacher_entry = None
        self.ccm_picker: _SearchList | None = None
        self.template_picker: _SearchList | None = None
        self._template_original = ""
        self.slot_hint = None
        row = 1
        if kind == KIND_CCM:
            self._build_ccm(body, row, ccm_files or [], item.value if item else "")
            row += 3
        if kind == KIND_CUSTOM:
            self._build_template_picker(body, row, template_names or [], item.label if item else "")
            row += 3
        if kind == KIND_PRAYER:
            # 기도 종류 = 설정의 기도 템플릿 폴더에 있는 PPT (합심기도, 중보기도 …)
            names = list(prayer_names or [])
            cur = item.label if item else ("기도" if "기도" in names else (names[0] if names else ""))
            if cur and cur not in names:  # 폴더에서 사라진 템플릿이어도 기존 값은 보여 줌
                names.insert(0, cur)
            self._lbl(body, row, "기도 종류 *")
            self.label_widget = ctk.CTkComboBox(body, width=260, values=names, state="readonly", font=font(13),
                                                dropdown_font=font(13))
            self.label_widget.set(cur)
            self.label_widget.grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")
            row += 1

        prompt, hint = {
            KIND_PRAYER: ("기도자 이름 *", "예: 홍길동 장로"),
            KIND_HYMN: ("찬송가 장 *", "예: 29  또는  29장"),
            KIND_BIBLE: ("성경 구절 *", "예: 창1:1-5, 롬1-3"),
            KIND_SERMON: ("설교 제목 *", "예: 믿음의 사람"),
            KIND_CUSTOM: ("내용1", "예: 인도자 (선택)"),
        }.get(kind, (None, None))
        if prompt:
            self._lbl(body, row, prompt)
            self.value_entry = ctk.CTkEntry(body, width=260, placeholder_text=hint, font=font(13))
            if item and item.value:
                self.value_entry.insert(0, item.value)
            self.value_entry.grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")
            row += 1

        if kind == KIND_PRAYER:
            ctk.CTkLabel(body, text="기도 폴더(⚙ 설정 → 폴더)에서 고른 PPT 의 {{내용1}} 자리에 기도자 이름이 들어갑니다.",
                         font=font(11), text_color=TEXT2).grid(row=row, column=0, columnspan=2, padx=16, sticky="w")
            row += 1
        if kind == KIND_SERMON:
            self._lbl(body, row, "설교자 *")
            self.preacher_entry = ctk.CTkEntry(body, width=260, placeholder_text="예: 홍길동 목사", font=font(13))
            if self.options.get("preacher"):
                self.preacher_entry.insert(0, self.options["preacher"])
            self.preacher_entry.grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")
            row += 1
            ctk.CTkLabel(body, text="설정의 설교 템플릿 PPT 에서 제목은 {{내용1}}, 설교자는 {{내용2}} 자리에 들어갑니다.",
                         font=font(11), text_color=TEXT2).grid(row=row, column=0, columnspan=2, padx=16, sticky="w")
            row += 1
        if kind == KIND_BIBLE:
            # 역본을 지정한 항목이면 그 값, 아니면(새 항목·기본 역본 사용 항목) 빈 칸 = ⚙ 설정의 기본 역본을 따름
            own = any(self.options.get(k) for k in BIBLE_KEYS)
            self._lbl(body, row, "성경 역본 (선택)")
            self.bible_widget = BibleField(body, values=bible_values, width=260, placeholder="비우면 설정의 기본 역본")
            self.bible_widget.set(self.options.get("bible", "") if own else default_bible)
            self.bible_widget.grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")
            row += 1
            self._lbl(body, row, "두 번째 역본 (선택)")
            self.bible2_widget = BibleField(body, values=bible_values, width=260, placeholder="나란히 표시할 역본",
                                            clear_text="지정 안 함")
            self.bible2_widget.set(self.options.get("bible2", "") if own else default_bible2)
            self.bible2_widget.grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")
            row += 1
            ctk.CTkLabel(body, text="여러 구절은 쉼표로 구분합니다. 역본을 비우면 설정의 기본 역본을 사용합니다.\n"
                                    "두 번째 역본을 고르면 한 슬라이드에 두 역본이 나란히 들어갑니다.",
                         font=font(11), text_color=TEXT2, justify="left").grid(
                row=row, column=0, columnspan=2, padx=16, sticky="w")
            row += 1
        if kind == KIND_CUSTOM:
            # 내용2·3 칸 (내용1 = 위의 value_entry). 템플릿에 자리가 없어도 항상 보여 줌 — 어떤 자리가 있는지는 아래 안내
            for n, key in CUSTOM_TEXT_KEYS.items():
                lbl = ctk.CTkLabel(body, text=f"내용{n}", font=font(13), text_color=TEXT, anchor="w", width=130)
                lbl.grid(row=row, column=0, padx=16, pady=6, sticky="w")
                e = ctk.CTkEntry(body, width=260, placeholder_text="(선택)", font=font(13))
                if self.options.get(key):
                    e.insert(0, self.options[key])
                e.grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")
                self.text_entries[n] = e
                self._text_labels[n] = lbl
                row += 1
            self.slot_hint = ctk.CTkLabel(body, text="", font=font(11), text_color=TEXT2, justify="left")
            self.slot_hint.grid(row=row, column=0, columnspan=2, padx=16, sticky="w")
            row += 1
            self._update_custom_slots()
        ctk.CTkFrame(body, width=1, height=8, fg_color="transparent").grid(row=row, column=0)

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=16, pady=(4, 16))
        accent_button(btns, "확인", self._ok, width=96).pack(side="right")
        secondary_button(btns, "취소", self.destroy, width=96).pack(side="right", padx=8)
        self.bind_return(self._ok)

    def on_shown(self):
        # 사용자 지정: 새로 추가할 때는 템플릿 검색칸, 수정할 때는 내용 칸부터
        if self.kind == KIND_CUSTOM and not self._template_original:
            self.template_picker.entry.focus_set()
        else:
            self.value_entry.focus_set()

    # ---- 사용자 지정: 템플릿에 있는 {{내용n}} 자리에 맞춰 입력칸 보이기
    def _update_custom_slots(self):
        self._slots_job = None
        if self.slot_hint is None:  # 목록을 처음 채울 때 (안내 칸을 만들기 전)
            return
        name = self._custom_label().strip()
        slots = self._fill_slots(name) if (self._fill_slots and name) else set()
        tokens = "~".join((self._slot_names[0], self._slot_names[-1]))
        if not name:
            text = f"위 목록에서 템플릿 PPT 를 고르세요.\n내용1~3 은 템플릿의 {tokens} 자리에 들어갑니다."
        elif slots is None:
            text = "이 템플릿 형식(.ppt)은 입력 자리를 미리 확인할 수 없습니다."
        elif slots:
            text = "템플릿에서 찾은 입력 자리: " + ", ".join(self._slot_names[n - 1] for n in sorted(slots))
        else:
            text = "템플릿에 입력 자리가 없습니다.\n내용1 은 순서 목록 오른쪽에만 표시됩니다."
        self.slot_hint.configure(text=text)

    # ---- CCM: 파일명이 정형화돼 있지 않으므로 CCM 폴더의 파일을 이름으로 검색해 고른다 (항목 값 = 상대 경로)
    def _build_ccm(self, body, row: int, files: list[str], current: str):
        current = current.replace("\\", "/")
        self.ccm_picker = _SearchList(
            self, body, row, "CCM 파일 *", files,
            [ccm_title(f) + (f"   ({f.rsplit('/', 1)[0]})" if "/" in f else "") for f in files], current,
            missing_text=ccm_title(current), icon="🎵", rows=self.CCM_ROWS,
            empty_text="CCM 폴더에 PPT 파일이 없습니다. ⚙ 설정 → 폴더에서 CCM 폴더를 지정해 주세요.")
        self.value_entry = self.ccm_picker.entry

    # ---- 사용자 지정: 템플릿 폴더의 PPT 를 검색해 고른다 (항목 label = 파일 이름)
    def _build_template_picker(self, body, row: int, names: list[str], current: str):
        self._template_original = current
        self.template_picker = _SearchList(
            self, body, row, "템플릿 PPT *", names, names, current, missing_text=current, icon="📄",
            rows=self.TEMPLATE_ROWS, on_pick=self._update_custom_slots,
            empty_text="템플릿 폴더에 PPT 파일이 없습니다. 왼쪽 [템플릿 파일]의 [불러오기]로 넣어 주세요.")
        self.label_widget = self.template_picker

    def _custom_label(self) -> str:
        """고른 템플릿 이름. 목록에 없는 템플릿(폴더에서 사라짐)을 쓰던 항목은 검색칸을 그대로 두면 원래 이름 유지."""
        name = self.template_picker.get()
        if not name and self._template_original and self.template_picker.entry.get().strip() == self._template_original:
            return self._template_original
        return name

    @staticmethod
    def _lbl(master, row, text):
        ctk.CTkLabel(master, text=text, font=font(13), text_color=TEXT, anchor="w", width=130).grid(
            row=row, column=0, padx=16, pady=6, sticky="w")

    def _ok(self):
        if self.template_picker is not None:
            label = self._custom_label()
        else:
            label = self.label_widget.get() if self.label_widget is not None else KIND_LABELS[self.kind]
        if self.kind == KIND_CUSTOM and not label.strip():
            messagebox.showerror("입력 오류", "목록에서 템플릿 PPT 를 골라 주세요.", parent=self)
            return
        if self.kind == KIND_PRAYER and not label.strip():
            messagebox.showerror("입력 오류", "기도 종류(템플릿)를 선택해 주세요.", parent=self)
            return
        if self.bible_widget is not None:
            b1, b2 = self.bible_widget.get().strip(), self.bible2_widget.get().strip()
            if b2 and not b1 and not self._fallback_bible:
                messagebox.showerror("입력 오류", "두 번째 역본을 쓰려면 첫 번째 역본도 지정해 주세요\n"
                                     "(또는 ⚙ 설정에서 기본 성경 역본을 지정).", parent=self)
                return
            self.options["bible"], self.options["bible2"] = b1, b2  # 빈 값은 make_item 이 지움
        if self.preacher_entry is not None:
            self.options["preacher"] = self.preacher_entry.get()
        for n, e in self.text_entries.items():
            self.options[CUSTOM_TEXT_KEYS[n]] = e.get()
        value = self.ccm_picker.get() if self.ccm_picker is not None else self.value_entry.get()
        item = make_item(self.kind, label, value, self.options)
        err = validate_item(item)
        if err:
            messagebox.showerror("입력 오류", err, parent=self)
            return
        self.result = item
        self.destroy()


# ---------------------------------------------------------------- 파일 찾기·역본 목록 (설정 창·설정 마법사 공용)
def path_text(entry) -> str:
    """경로 입력칸의 값. 탐색기 [경로로 복사]는 "C:\…" 처럼 따옴표를 붙이므로 앞뒤 따옴표·공백을 뗀다."""
    return entry.get().strip().strip('"').strip()


class _BibleListMixin:
    """Bible2PPT.exe 를 지정하면 --list-bibles 로 역본 목록을 불러와 기본 역본 칸(bible·bible2)의 🔍 목록을 채운다.
    쓰는 쪽이 갖출 것: cfg, exe(입력칸), bible·bible2(BibleField), bible_hint(라벨), bible_values, _loaded_for,
    _loading_for, _closing(_Modal), _bible_timeout()."""

    def _browse(self, entry: ctk.CTkEntry, types) -> bool:
        path = filedialog.askopenfilename(parent=self, filetypes=types + [("모든 파일", "*.*")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)
        return bool(path)

    def _browse_exe(self):
        if self._browse(self.exe, [("Bible2PPT", "*.exe")]):
            self._load_bibles_if_needed()

    def _exe(self):
        exe = self.cfg.resolve_file(path_text(self.exe))
        return exe if exe and exe.is_file() else None

    def _show_bible_hint(self, text: str | None = None, error: bool = False):
        if text is None:
            if self._exe() is None:
                text = "Bible2PPT.exe 를 지정하면 사용할 수 있는 역본 목록을 불러옵니다."
            elif self.bible_values:
                text = f"역본 {len(self.bible_values)}종 — 🔍 로 검색해 고르거나 비워 두면 Bible2PPT 마지막 선택값을 사용합니다."
            else:
                text = "↻ 를 눌러 역본 목록을 불러오세요."
        self.bible_hint.configure(text=text, text_color=DANGER if error else TEXT2)

    def _load_bibles_if_needed(self):
        exe = self._exe()
        if exe is not None and exe != self._loaded_for and exe != self._loading_for:
            self._load_bibles()
        elif self._loading_for is None:
            self._show_bible_hint()

    def _load_bibles(self):
        exe = self._exe()
        if exe is None:
            self._show_bible_hint("Bible2PPT.exe 경로를 먼저 지정해 주세요.", error=True)
            return
        if exe == self._loading_for:
            return
        self._loading_for = exe
        timeout = self._bible_timeout()
        self._show_bible_hint("역본 목록을 불러오는 중… (인터넷에서 받아오므로 잠시 걸릴 수 있습니다)")
        run_in_thread(lambda: Bible2PPT(exe, timeout).list_bibles(), lambda res, err: self._on_bibles(exe, res, err))

    def _on_bibles(self, exe, res, err):
        if self._closing or not self.winfo_exists():
            return
        if exe != self._loading_for:  # 그사이 다른 경로를 불러오는 중
            return
        self._loading_for = None
        if exe != self._exe():  # 그사이 경로가 바뀜 → 새 경로로 다시
            self._load_bibles_if_needed()
            return
        if err:
            first = str(err).strip().splitlines()[0] if str(err).strip() else type(err).__name__
            self._show_bible_hint(f"⚠ 역본 목록을 불러오지 못했습니다: {first}", error=True)
            return
        self.bible_values = res
        self._loaded_for = exe
        for field in (self.bible, self.bible2):
            field.set_values(res)
            cur = field.get().strip()
            if cur and (match := match_bible(cur, res)) and match != cur:  # 예전 형식(이름만 입력) → 목록 항목
                field.set(match)
        self._show_bible_hint()


# ---------------------------------------------------------------- 설정 창
class SettingsDialog(_BibleListMixin, _Modal):
    """설정 창. Bible2PPT.exe 를 지정하면 --list-bibles 로 역본 목록을 바로 불러와 기본 역본 목록을 채운다.
    저장 시 불러온 목록은 config.bible2ppt.bible_list 에 캐시된다.
    항목이 많아 한 화면에 다 보이지 않으므로 한 쪽에 모두 두고 세로로 스크롤한다(높이는 메인 창에 맞춤, _fit_body)."""

    def __init__(self, master, cfg, preview: tuple[str, str, str] = ("", "", "")):
        """preview: 파일 이름 형식 미리보기에 쓸 (날짜, 교회 이름, 예배 종류) — 메인 화면의 지금 값."""
        super().__init__(master, "설정")
        self.cfg = cfg
        self._preview_values = preview
        b = cfg.get("bible2ppt")
        p = cfg.get("ppt")
        t = cfg.get("item_templates")
        self.bible_values: list[str] = list(b.get("bible_list") or [])
        # 목록을 불러온 exe 경로 (캐시가 있으면 현재 설정된 exe 기준)
        self._loaded_for = cfg.resolve_file(b.get("exe_path", "")) if self.bible_values else None
        self._loading_for = None

        self.body = body = ctk.CTkScrollableFrame(self, fg_color=CARD, corner_radius=10, border_width=1,
                                                  border_color=BORDER)
        body.pack(fill="both", expand=True, padx=16, pady=(12, 8))
        body._parent_canvas.configure(yscrollincrement=max(1, scaled(self, 2.4)))  # 휠 한 칸 ≈ 48px (CTk 기본 20px)
        r = 0

        def section(text):
            nonlocal r
            ctk.CTkLabel(body, text=text, font=font(14, "bold"), text_color=TEXT).grid(
                row=r, column=0, columnspan=3, sticky="w", padx=16, pady=(14, 4))
            r += 1

        def row(label, widget, browse=None, button_text="찾기"):
            nonlocal r
            ctk.CTkLabel(body, text=label, font=font(13), text_color=TEXT, anchor="w", width=150).grid(
                row=r, column=0, padx=16, pady=5, sticky="w")
            widget.grid(row=r, column=1, pady=5, sticky="w")
            if browse:
                secondary_button(body, button_text, browse, width=56).grid(row=r, column=2, padx=(6, 16), pady=5)
            r += 1

        def entry(value, width=320):
            e = ctk.CTkEntry(body, width=width, font=font(13))
            e.insert(0, value or "")
            return e

        # 폴더: 한 번 정하면 거의 바꾸지 않으므로 메인 화면이 아니라 여기서 (예전에는 메인 왼쪽·상단에 5줄)
        section("폴더")
        self.folder_entries: dict[str, tuple[ctk.CTkEntry, str]] = {}  # 키 → (입력칸, 처음 보인 값)
        for key, name in SETTINGS_FOLDERS:
            shown = str(cfg.resolve_path(key))
            e = entry(shown)
            btns = ctk.CTkFrame(body, fg_color="transparent")
            secondary_button(btns, "📁", lambda k=key: self._choose_folder(k), width=36).pack(side="left")
            secondary_button(btns, "↗", lambda k=key: self._open_folder(k), width=36).pack(side="left", padx=(4, 0))
            ctk.CTkLabel(body, text=name, font=font(13), text_color=TEXT, anchor="w", width=150).grid(
                row=r, column=0, padx=16, pady=5, sticky="w")
            e.grid(row=r, column=1, pady=5, sticky="w")
            btns.grid(row=r, column=2, padx=(6, 16), pady=5, sticky="w")
            self.folder_entries[key] = (e, shown)
            r += 1
        ctk.CTkLabel(body, text="템플릿·기도·찬송가·CCM 폴더의 PPT 를 가져와 하나로 합치고, 결과는 출력 폴더에 저장합니다.\n"
                                "출력 폴더는 원본 폴더 밖이어야 합니다 (원본을 덮어쓰지 않도록). ↗ 는 폴더 열기.",
                     font=font(11), text_color=TEXT2, anchor="w", justify="left").grid(
            row=r, column=1, columnspan=2, sticky="w", pady=(0, 6))
        r += 1

        section("결과 파일 이름")
        fbox = ctk.CTkFrame(body, fg_color="transparent")
        self.format_entry = ctk.CTkEntry(fbox, width=320, font=font(13), placeholder_text=DEFAULT_FILENAME_FORMAT)
        self.format_entry.insert(0, cfg.get("filename_format") or DEFAULT_FILENAME_FORMAT)
        self.format_entry.pack(side="left")
        self.format_entry.bind("<KeyRelease>", lambda e: self._update_format_preview(), add="+")
        self.format_menu_btn = secondary_button(fbox, "항목 ▾", self._open_format_menu, width=64)
        self.format_menu_btn.pack(side="left", padx=(6, 0))
        secondary_button(fbox, "기본값", lambda: self._set_format(DEFAULT_FILENAME_FORMAT), width=64).pack(
            side="left", padx=(4, 0))
        ctk.CTkLabel(body, text="형식", font=font(13), text_color=TEXT, anchor="w", width=150).grid(
            row=r, column=0, padx=16, pady=5, sticky="w")
        fbox.grid(row=r, column=1, columnspan=2, pady=5, sticky="w")
        r += 1
        self.format_preview = ctk.CTkLabel(body, text="", font=font(12), anchor="w", justify="left")
        self.format_preview.grid(row=r, column=1, columnspan=2, sticky="w")
        r += 1
        ctk.CTkLabel(body, text="{항목} 자리에 값이 들어갑니다. [항목 ▾]로 넣을 수 있습니다.\n"
                                "프리셋을 저장하면 이 형식도 함께 저장되고, 프리셋을 불러오면 그 프리셋의 형식으로 바뀝니다.",
                     font=font(11), text_color=TEXT2, anchor="w", justify="left").grid(
            row=r, column=1, columnspan=2, sticky="w", pady=(0, 14))
        r += 1
        self._update_format_preview()

        section("Bible2PPT 연동")
        self.exe = entry(b.get("exe_path"))
        row("Bible2PPT.exe 경로", self.exe, self._browse_exe)
        self.exe.bind("<FocusOut>", lambda e: self._load_bibles_if_needed())
        self.exe.bind("<Return>", lambda e: self._load_bibles_if_needed())
        self.bible = BibleField(body, values=self.bible_values, width=320, placeholder="비우면 Bible2PPT 마지막 선택값",
                                clear_text="지정 안 함")
        self.bible.set(b.get("bible") or "")
        row("기본 성경 역본", self.bible, self._load_bibles, button_text="↻")
        self.bible2 = BibleField(body, values=self.bible_values, width=320, placeholder="나란히 표시할 역본 (선택)",
                                 clear_text="지정 안 함")
        self.bible2.set(b.get("bible2") or "")
        row("두 번째 기본 역본", self.bible2)
        self.bible_hint = ctk.CTkLabel(body, text="", font=font(11), text_color=TEXT2, anchor="w", justify="left")
        self.bible_hint.grid(row=r, column=1, columnspan=2, sticky="w", pady=(0, 4))
        r += 1
        self._show_bible_hint()
        self.tpl = entry(b.get("template"))
        row("성경 템플릿 (.pptx)", self.tpl, lambda: self._browse(self.tpl, [("PowerPoint", "*.pptx")]))
        self.lines = ctk.CTkComboBox(body, width=160, values=["프로그램 설정"] + [str(i) for i in range(10)],
                                     state="readonly", font=font(13), dropdown_font=font(13))
        self.lines.set("프로그램 설정" if b.get("lines") is None else str(b.get("lines")))
        row("슬라이드당 줄 수", self.lines)
        self.timeout = entry(str(b.get("timeout_sec", 180)), width=80)
        row("실행 제한 시간(초)", self.timeout)

        section("설교 템플릿")
        self.sermon_file = entry(t.get("sermon_file"))
        row("설교 템플릿 (.pptx)", self.sermon_file,
            lambda: self._browse(self.sermon_file, [("PowerPoint", "*.pptx *.pptm *.ppt")]))
        ctk.CTkLabel(body, text="설교 제목은 {{내용1}}, 설교자는 {{내용2}} 자리에 들어갑니다.", font=font(11),
                     text_color=TEXT2, anchor="w").grid(row=r, column=1, columnspan=2, sticky="w", pady=(0, 14))
        r += 1

        section("PPT 병합")
        self.merge_mode = ctk.CTkComboBox(body, width=320, values=list(MERGE_MODES), state="readonly",
                                          font=font(13), dropdown_font=font(13))
        self.merge_mode.set(next((k for k, v in MERGE_MODES.items() if v == p.get("merge_mode")),
                                 next(iter(MERGE_MODES))))
        row("병합 방식", self.merge_mode)
        # 결과 슬라이드 비율 (원본 = 첫 PPT 크기). 예전에는 메인 하단 생성 버튼 옆 — 한 번 정하면 거의 안 바꿔서 여기로
        self.size_seg = ctk.CTkSegmentedButton(body, values=list(SLIDE_SIZE_LABELS), font=font(12), height=30,
                                               selected_color=ACCENT)
        self.size_seg.set(next((k for k, v in SLIDE_SIZE_LABELS.items() if v == p.get("slide_size")), "원본"))
        row("슬라이드 비율", self.size_seg)
        ctk.CTkLabel(body, text="원본 = 첫 번째 순서 PPT 크기. 16:9·4:3 을 고르면 모든 슬라이드를 그 비율로 맞춥니다 (내용은 잘리지 않음).",
                     font=font(11), text_color=TEXT2, anchor="w").grid(row=r, column=1, columnspan=2, sticky="w", pady=(0, 6))
        r += 1

        section("텍스트 입력 규칙")

        def numbered_row(label, value, make):
            """기본형 입력칸 + 실제로 쓰일 번호 붙은 이름 미리보기 (입력할 때마다 갱신)."""
            box = ctk.CTkFrame(body, fg_color="transparent")
            e = ctk.CTkEntry(box, width=140, font=font(13))
            e.insert(0, value or "")
            e.pack(side="left")
            preview = ctk.CTkLabel(box, text="", font=font(12), text_color=ACCENT)
            preview.pack(side="left", padx=(10, 0))

            def update(_=None):
                base = e.get().strip()
                preview.configure(text="→  " + "   ".join(make(base, n) for n in range(1, MAX_FILL + 1)) if base else "")
            e.bind("<KeyRelease>", update, add="+")
            update()
            row(label, box)
            return e

        self.placeholder = numbered_row("치환 토큰", p.get("text_placeholder"), numbered_token)
        self.shape_name = numbered_row("텍스트 상자 이름", p.get("text_shape_name"), lambda b, n: f"{b}{n}")
        # 이 안내가 설정 창에서 가장 긴 줄이라 창 너비를 정한다 → 기존 둘째 줄(약 550)보다 길어지지 않게 나눠 쓸 것
        ctk.CTkLabel(body, text="기본형에 번호를 붙여 사용합니다.\n"
                                "  1번: 기도자 이름 · 찬송 장 · 설교 제목 · 사용자 지정 내용1\n"
                                "  2번: 설교자 · 사용자 지정 내용2      3번: 사용자 지정 내용3\n"
                                "토큰이 없으면 같은 이름의 텍스트 상자에 입력합니다. 번호 없는 기본형은 1번으로 취급합니다(예전 템플릿 호환).\n"
                                "공통 토큰(모든 슬라이드): {{날짜}} {{년}} {{월}} {{일}} {{요일}} {{교회이름}} {{예배종류}}",
                     font=font(11), text_color=TEXT2, justify="left").grid(
            row=r, column=0, columnspan=3, padx=16, pady=(2, 6), sticky="w")
        r += 1

        section("빈 칸일 때 기본값")
        defaults = list(p.get("fill_defaults") or [])
        self.fill_defaults: list[ctk.CTkEntry] = []
        self._default_labels: list[ctk.CTkLabel] = []
        for n in range(1, MAX_FILL + 1):
            lbl = ctk.CTkLabel(body, text="", font=font(13), text_color=TEXT, anchor="w", width=150)
            lbl.grid(row=r, column=0, padx=16, pady=4, sticky="w")
            e = ctk.CTkEntry(body, width=320, font=font(13), placeholder_text="값이 비었을 때 넣을 문구 (비우면 토큰 지움)")
            if n <= len(defaults) and defaults[n - 1]:
                e.insert(0, defaults[n - 1])
            e.grid(row=r, column=1, pady=(4, 14 if n == MAX_FILL else 4), sticky="w")
            self.fill_defaults.append(e)
            self._default_labels.append(lbl)
            r += 1
        # 기본값 칸 이름은 치환 토큰을 바꾸면 따라 바뀜 ({{내용1}} …)
        self.placeholder.bind("<KeyRelease>", lambda e: self._update_default_labels(), add="+")
        self._update_default_labels()

        section("화면")
        current = {v: k for k, v in APPEARANCE_MODES.items()}.get(cfg.get("appearance_mode"), "시스템")
        self.mode_seg = ctk.CTkSegmentedButton(body, values=list(APPEARANCE_MODES), font=font(12), selected_color=ACCENT,
                                               height=32)
        self.mode_seg.set(current)
        row("화면 모드", self.mode_seg)
        self.use_accent = ctk.BooleanVar(value=bool(cfg.get("use_system_accent")))
        ctk.CTkCheckBox(body, text="Windows 강조 색 사용", variable=self.use_accent, font=font(13), checkbox_width=20,
                        checkbox_height=20, fg_color=ACCENT).grid(row=r, column=0, columnspan=2, padx=16, pady=5, sticky="w")
        r += 1
        palette = windows_accent_palette()
        ctk.CTkLabel(body, text=("Windows 설정 → 개인 설정 → 색의 강조 색을 버튼·선택 표시 등에 씁니다. "
                                 "끄면 기본 파란색. 프로그램을 다시 켜면 바뀝니다.")
                     + ("" if palette else "\n(이 PC 에서는 Windows 강조 색을 읽을 수 없어 기본 파란색을 씁니다.)"),
                     font=font(11), text_color=TEXT2, anchor="w", justify="left").grid(
            row=r, column=0, columnspan=3, padx=16, pady=(0, 10), sticky="w")
        r += 1

        section("임시 파일")
        self.temp_keep = entry(str(clean_temp_keep(p.get("temp_keep"))), width=80)
        row("최대 보관 개수", self.temp_keep)
        ctk.CTkLabel(body, text=f"[미리 보기 (저장 안 함)]로 만든 PPT 를 최근 것부터 이만큼 보관합니다 (1~{TEMP_KEEP_MAX}).\n"
                                "넘으면 오래된 것부터 지웁니다. 하단 [임시 파일]에서 다시 열거나 저장할 수 있습니다.",
                     font=font(11), text_color=TEXT2, anchor="w", justify="left").grid(
            row=r, column=1, columnspan=2, sticky="w", pady=(0, 14))
        r += 1
        self._fit_body()

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=16, pady=(4, 16))
        accent_button(btns, "저장", self._save, width=96).pack(side="right")
        secondary_button(btns, "취소", self.destroy, width=96).pack(side="right", padx=8)

    def on_shown(self):
        self._load_bibles_if_needed()  # exe 는 지정돼 있는데 목록이 없으면 바로 불러옴

    def _update_default_labels(self):
        base = self.placeholder.get().strip()
        for n, lbl in enumerate(self._default_labels, 1):
            lbl.configure(text=numbered_token(base, n) if base else f"{n}번")

    # ---- 폴더
    def _folder_raw(self, key: str) -> str:
        """저장할 값: 처음 보인 값(절대 경로) 그대로면 원래 값 (기본 'templates' 같은 상대 경로는 exe 를 옮겨도 따라가도록)."""
        e, shown = self.folder_entries[key]
        raw = path_text(e)
        return self.cfg.data["paths"].get(key, "") if raw == shown else raw

    def _choose_folder(self, key: str):
        e = self.folder_entries[key][0]
        p = self.cfg.resolve_file(path_text(e)) if path_text(e) else None
        path = filedialog.askdirectory(parent=self, initialdir=str(p) if p and p.is_dir() else None, mustexist=True)
        if path:
            e.delete(0, "end")
            e.insert(0, str(Path(path)))

    def _open_folder(self, key: str):
        raw = path_text(self.folder_entries[key][0])
        if raw:
            open_folder(self, self.cfg.resolve_file(raw))

    def _check_folders(self):
        """(문제 칸, 메시지) 또는 None. 설정 마법사와 같은 규칙 + 출력 폴더."""
        probe = copy.copy(self.cfg)
        probe.data = copy.deepcopy(self.cfg.data)
        for key, _name in SETTINGS_FOLDERS:
            probe.data["paths"][key] = self._folder_raw(key)
        for key, name in SETTINGS_FOLDERS:
            e = self.folder_entries[key][0]
            raw = path_text(e)
            if not raw:
                return e, f"{name} 폴더를 지정해 주세요."
            p = self.cfg.resolve_file(raw)
            if key == "output_dir":
                try:  # 출력 폴더는 없으면 만든다 (생성 때도 만듦)
                    p.mkdir(parents=True, exist_ok=True)
                except OSError as err:
                    return e, f"출력 폴더를 만들 수 없습니다:\n{p}\n{err}"
            if not p.is_dir():
                return e, f"{name} 폴더를 찾을 수 없습니다:\n{p}"
            if (err := folder_conflict(probe, key, p)) is not None:
                return e, err
        if (err := output_conflict(probe, probe.resolve_path("output_dir") / "_.pptx")) is not None:
            return self.folder_entries["output_dir"][0], err
        return None

    # ---- 결과 파일 이름 형식
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
        self._update_format_preview()

    def _set_format(self, fmt: str):
        self.format_entry.delete(0, "end")
        self.format_entry.insert(0, fmt)
        self._update_format_preview()

    def _update_format_preview(self):
        fmt = self.format_entry.get().strip()
        if err := validate_filename_format(fmt):
            self.format_preview.configure(text="⚠ " + err.splitlines()[0], text_color=DANGER)
        else:
            date, church, worship = self._preview_values
            self.format_preview.configure(text="→  " + output_filename(date, church, worship, fmt), text_color=TEXT2)

    def _fit_body(self):
        """스크롤 영역 크기: 너비는 내용에 맞추고, 높이는 내용과 메인 창(제목·버튼 자리를 뺀 만큼) 중 작은 쪽."""
        self.body.update_idletasks()
        s = ctk.ScalingTracker.get_widget_scaling(self)
        inner = self.body  # CTkScrollableFrame 자체가 내용 틀 (캔버스 안에 배치됨)
        req_w, req_h = inner.winfo_reqwidth() / s, inner.winfo_reqheight() / s
        win_h = self._win.winfo_height() / s
        # 다이얼로그는 창 높이 45% 지점에 가운데 맞춤(_Modal.show) → 전체가 창 높이의 90% 안이어야 위가 잘리지 않음.
        # 스크롤 영역 밖(제목·버튼·여백)은 약 170
        self.body.configure(width=int(req_w) + 4, height=int(max(240, min(req_h, win_h * 0.9 - 170))))

    def _reveal(self, widget) -> None:
        """입력 오류가 난 칸이 보이도록 스크롤하고 포커스 (예전 두 쪽 설정 창의 쪽 넘기기 대신)."""
        try:
            canvas = self.body._parent_canvas
            total = max(self.body.winfo_height(), 1)
            y = widget.winfo_rooty() - self.body.winfo_rooty()
            canvas.yview_moveto(max(0.0, (y - scaled(self, 40)) / total))
            widget.focus_set()
        except tk.TclError:
            pass

    def _bible_timeout(self) -> int:
        try:
            return max(int(self.timeout.get().strip() or 180), 10)
        except ValueError:
            return 180

    def _save(self):
        if (problem := self._check_folders()) is not None:
            widget, msg = problem
            self._reveal(widget)
            messagebox.showerror("폴더", msg, parent=self)
            return
        fmt = self.format_entry.get().strip()
        if err := validate_filename_format(fmt):
            self._reveal(self.format_entry)
            messagebox.showerror("결과 파일 이름", err, parent=self)
            return
        try:
            timeout = int(self.timeout.get().strip() or 180)
            if timeout < 10:
                raise ValueError
        except ValueError:
            self._reveal(self.timeout)
            messagebox.showerror("입력 오류", "실행 제한 시간은 10 이상의 숫자로 입력해 주세요.", parent=self)
            return
        exe_raw = path_text(self.exe)
        if exe_raw and not exe_raw.lower().endswith(".exe"):
            self._reveal(self.exe)
            messagebox.showerror("입력 오류", "Bible2PPT 경로는 .exe 실행 파일이어야 합니다.", parent=self)
            return
        tpl_raw = path_text(self.tpl)
        if tpl_raw:
            tp = self.cfg.resolve_file(tpl_raw)
            if tp.suffix.lower() != ".pptx" or not tp.is_file():
                self._reveal(self.tpl)
                messagebox.showerror("입력 오류", "성경 템플릿은 존재하는 .pptx 파일이어야 합니다\n"
                                     f"(매크로가 들어갈 수 있는 .pptm 등은 사용할 수 없음):\n{tpl_raw}", parent=self)
                return
        sermon_raw = path_text(self.sermon_file)
        if sermon_raw:
            sp = self.cfg.resolve_file(sermon_raw)
            if not sp.is_file() or sp.suffix.lower() not in PPT_EXTS:
                self._reveal(self.sermon_file)
                messagebox.showerror("입력 오류", f"설교 템플릿 PPT 파일을 찾을 수 없습니다:\n{sermon_raw}", parent=self)
                return
        bible1, bible2 = self.bible.get().strip(), self.bible2.get().strip()
        if bible2 and not bible1:
            self._reveal(self.bible)
            messagebox.showerror("입력 오류", "두 번째 기본 역본을 쓰려면 기본 성경 역본도 지정해 주세요.", parent=self)
            return
        keep_raw = self.temp_keep.get().strip()
        if not (keep_raw.isascii() and keep_raw.isdigit()) or not 1 <= int(keep_raw) <= TEMP_KEEP_MAX:
            self._reveal(self.temp_keep)
            messagebox.showerror("입력 오류", f"임시 파일 최대 보관 개수는 1~{TEMP_KEEP_MAX} 사이 숫자로 입력해 주세요.",
                                 parent=self)
            return
        lines = self.lines.get()
        b = self.cfg.get("bible2ppt")
        b.update(exe_path=path_text(self.exe), bible=bible1, bible2=bible2, template=tpl_raw,
                 lines=None if lines == "프로그램 설정" else int(lines), timeout_sec=timeout,
                 # 현재 exe 로 불러온 목록만 캐시 (exe 가 바뀌었는데 아직 못 불러왔으면 비움 → 메인에서 다시 불러옴)
                 bible_list=self.bible_values if self._loaded_for is not None and self._loaded_for == self._exe() else [])
        self.cfg.get("item_templates").update(sermon_file=sermon_raw)
        self.cfg.data["use_system_accent"] = bool(self.use_accent.get())
        for key, _name in SETTINGS_FOLDERS:
            self.cfg.data["paths"][key] = self._folder_raw(key)
        self.cfg.data["filename_format"] = fmt
        self.cfg.data["appearance_mode"] = APPEARANCE_MODES.get(self.mode_seg.get(), "System")
        p = self.cfg.get("ppt")
        p.update(text_placeholder=self.placeholder.get().strip(), text_shape_name=self.shape_name.get().strip(),
                 fill_defaults=[e.get().strip() for e in self.fill_defaults],
                 merge_mode=MERGE_MODES[self.merge_mode.get()], temp_keep=int(keep_raw),
                 slide_size=SLIDE_SIZE_LABELS.get(self.size_seg.get(), "original"))
        self.cfg.save()
        self.result = True
        self.destroy()


# ---------------------------------------------------------------- 설정 마법사
BIBLE2PPT_URL = "https://github.com/nazunamoe/Bible2PPT"
WIZARD_FOLDERS = (  # (설정 키, 이름, 설명, 하위 폴더까지 찾는지)
    ("template_dir", "템플릿 폴더", "예배로 부름·축도처럼 순서마다 쓰는 PPT (파일 이름 = 순서 이름)", False),
    ("prayer_dir", "기도 폴더", "합심기도.pptx 처럼 기도 종류별 PPT", False),
    ("hymn_dir", "찬송가 폴더", "파일 이름에 장 번호가 있는 찬송가 PPT (예: 029장.pptx, 하위 폴더 포함)", True),
    ("ccm_dir", "CCM 폴더", "곡명이 파일 이름인 CCM PPT (하위 폴더 포함)", True),
)
LINES_DEFAULT = "프로그램 설정"  # 슬라이드당 줄 수를 Bible2PPT 에 맡김 (bible2ppt.lines = None)


class SetupWizard(_BibleListMixin, _Modal):
    """설정 마법사: 처음 쓰는 사람이 폴더(템플릿·기도·찬송가·CCM) → 설교 템플릿 → 말씀(Bible2PPT·역본·성경 템플릿·줄 수)
    → 확인 순서로 지정한다. 처음 실행할 때 자동으로, 이후에는 제목 줄 [설정 마법사]로 띄운다.
    단계마다 [다음]에서 그 단계 값을 검사하고, [마침]에서만 저장한다(중간에 닫으면 아무것도 바뀌지 않음).
    값은 ⚙ 설정과 같은 config 키를 쓰고, 검사 규칙도 같다(folder_conflict·output_conflict·temp_dir_conflict)."""
    STEPS = ("폴더", "설교", "말씀", "확인")
    FIELD_W = 400  # 입력칸 너비 (모든 단계 공통 — 단계를 넘겨도 다이얼로그 너비가 그대로)

    def __init__(self, master, cfg):
        super().__init__(master, "설정 마법사")
        self.cfg = cfg
        b = cfg.get("bible2ppt")
        self.bible_values: list[str] = list(b.get("bible_list") or [])
        self._loaded_for = cfg.resolve_file(b.get("exe_path", "")) if self.bible_values else None
        self._loading_for = None
        self.step = 0

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=20, pady=(6, 0))
        self._step_labels: list[ctk.CTkLabel] = []
        for i, name in enumerate(self.STEPS):
            if i:
                ctk.CTkLabel(bar, text="›", font=font(13), text_color=TEXT2).pack(side="left", padx=8)
            lbl = ctk.CTkLabel(bar, text=f"{i + 1}. {name}", font=font(13))
            lbl.pack(side="left")
            self._step_labels.append(lbl)

        # 단계 쪽은 모두 같은 칸에 두고 지금 쪽만 보인다. 칸 크기를 가장 큰 쪽에 맞춰 단계를 넘겨도 다이얼로그 크기가 그대로.
        # 작은 화면에서는 창보다 커질 수 있어 스크롤 영역 (_fit_body)
        self.body = ctk.CTkScrollableFrame(self, fg_color=CARD, corner_radius=10, border_width=1, border_color=BORDER)
        self.body.pack(fill="both", expand=True, padx=16, pady=(10, 8))
        self.body._parent_canvas.configure(yscrollincrement=max(1, scaled(self, 2.4)))  # 설정 창과 같은 휠 속도
        self.pages = [self._build_folders(), self._build_sermon(), self._build_bible(), self._build_summary()]
        for page in self.pages:
            page.grid(row=0, column=0, sticky="nsew", padx=20, pady=(14, 16))
        self.body.update_idletasks()
        # 칸 최소 크기는 쪽 바깥 여백(padx·pady)까지 포함한 크기여야 함 (빠뜨리면 큰 쪽에서 칸이 다시 커짐)
        self.body.grid_rowconfigure(0, minsize=max(p.winfo_reqheight() for p in self.pages) + scaled(self, 30))
        self.body.grid_columnconfigure(0, minsize=max(p.winfo_reqwidth() for p in self.pages) + scaled(self, 40))

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=16, pady=(4, 16))
        secondary_button(btns, "나중에 하기", self.destroy, width=96).pack(side="left")
        self.next_btn = accent_button(btns, "다음", self._next, width=96)
        self.next_btn.pack(side="right")
        self.prev_btn = secondary_button(btns, "이전", self._prev, width=96)
        self.prev_btn.pack(side="right", padx=8)
        self.bind_return(self._next)
        self._show_step(0)
        self._fit_body()

    def _fit_body(self):
        """스크롤 영역 크기: 너비는 가장 큰 쪽, 높이는 내용과 (창 높이 90% − 제목·단계·버튼) 중 작은 쪽.
        다이얼로그는 창 45% 지점에 가운데 맞춤(_Modal.show)이라 이보다 크면 위가 잘린다 — 1366×768·125% 노트북이면
        창 높이가 약 530 인데 마법사는 약 574 였음."""
        s = ctk.ScalingTracker.get_widget_scaling(self)
        self.update_idletasks()
        req_w, req_h = self.body.winfo_reqwidth() / s, self.body.winfo_reqheight() / s
        self.body.configure(width=int(req_w) + 4, height=int(req_h))
        self.update_idletasks()
        # height 로 지정되는 것은 안쪽 캔버스 — 바깥 틀의 테두리·둥근 모서리 여백(약 21)까지 '나머지'에 넣어야 맞음
        extra = (self.winfo_reqheight() - self.body._parent_canvas.winfo_reqheight()) / s
        avail = self._win.winfo_height() / s * 0.9 - extra
        self.body.configure(height=int(max(160, min(req_h, avail))))

    # ---- 단계 쪽 만들기
    def _page(self) -> ctk.CTkFrame:
        return ctk.CTkFrame(self.body, fg_color="transparent")

    def _intro(self, page, text: str):
        ctk.CTkLabel(page, text=text, font=font(13), text_color=TEXT, anchor="w", justify="left",
                     wraplength=self.FIELD_W + 160).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

    def _note(self, page, r: int, text: str = "") -> ctk.CTkLabel:
        lbl = ctk.CTkLabel(page, text=text, font=font(11), text_color=TEXT2, anchor="w", justify="left",
                           wraplength=self.FIELD_W + 40)
        lbl.grid(row=r, column=1, columnspan=2, sticky="w", pady=(0, 8))
        return lbl

    def _label(self, page, r: int, text: str):
        ctk.CTkLabel(page, text=text, font=font(13, "bold"), text_color=TEXT, anchor="w", width=130).grid(
            row=r, column=0, sticky="w", pady=(4, 2))

    def _entry(self, page, value: str) -> ctk.CTkEntry:
        e = ctk.CTkEntry(page, width=self.FIELD_W, font=font(13))
        e.insert(0, value or "")
        return e

    def _build_folders(self):
        page = self._page()
        self._intro(page, "예배 PPT 를 만들 때 가져올 PPT 들이 있는 폴더를 지정합니다.\n"
                          "아직 PPT 가 없다면 그대로 두고, 나중에 이 폴더에 넣어도 됩니다.")
        self.folders: dict[str, tuple[ctk.CTkEntry, str, ctk.CTkLabel, bool]] = {}  # 키 → (입력칸, 처음 값, 안내, 하위 포함)
        self._desc = {key: desc for key, _n, desc, _r in WIZARD_FOLDERS}
        r = 1
        for key, name, desc, recursive in WIZARD_FOLDERS:
            shown = str(self.cfg.resolve_path(key))
            self._label(page, r, name)
            e = self._entry(page, shown)
            e.grid(row=r, column=1, sticky="w", pady=(4, 2))
            secondary_button(page, "📁", lambda k=key: self._choose_folder(k), width=36).grid(
                row=r, column=2, padx=(6, 0), pady=(4, 2))
            note = self._note(page, r + 1)
            self.folders[key] = (e, shown, note, recursive)
            e.bind("<FocusOut>", lambda _e, k=key: self._update_folder_note(k), add="+")
            self._update_folder_note(key)
            r += 2
        return page

    def _build_sermon(self):
        page = self._page()
        t = self.cfg.get("item_templates")
        self._intro(page, "설교 순서에 쓸 PPT 템플릿 파일 하나를 지정합니다.\n"
                          "템플릿 안의 {{내용1}} 자리에 설교 제목, {{내용2}} 자리에 설교자가 들어갑니다.")
        self._label(page, 1, "설교 템플릿")
        self.sermon_file = self._entry(page, t.get("sermon_file"))
        self.sermon_file.grid(row=1, column=1, sticky="w", pady=(4, 2))
        secondary_button(page, "찾기", lambda: self._browse(self.sermon_file, [("PowerPoint", "*.pptx *.pptm *.ppt")]),
                         width=56).grid(row=1, column=2, padx=(6, 0), pady=(4, 2))
        self._note(page, 2, "비워 두면 설교 순서를 추가할 수 없습니다. 나중에 ⚙ 설정에서 지정할 수 있습니다.")
        return page

    def _build_bible(self):
        page = self._page()
        b = self.cfg.get("bible2ppt")
        self._intro(page, "말씀(성경 구절) PPT 는 Bible2PPT 프로그램이 만듭니다. 명령줄 실행(CLI)을 지원하는 버전이 필요합니다.\n"
                          "Bible2PPT 를 지정하면 사용할 수 있는 성경 역본 목록을 불러옵니다.")
        self._label(page, 1, "Bible2PPT.exe")
        self.exe = self._entry(page, b.get("exe_path"))
        self.exe.grid(row=1, column=1, sticky="w", pady=(4, 2))
        self.exe.bind("<FocusOut>", lambda e: self._load_bibles_if_needed(), add="+")
        secondary_button(page, "찾기", self._browse_exe, width=56).grid(row=1, column=2, padx=(6, 0), pady=(4, 2))
        link = ctk.CTkFrame(page, fg_color="transparent")
        link.grid(row=2, column=1, columnspan=2, sticky="w", pady=(0, 8))
        ctk.CTkLabel(link, text="아직 없다면:", font=font(11), text_color=TEXT2).pack(side="left")
        secondary_button(link, "Bible2PPT 받기 ↗", lambda: webbrowser.open(BIBLE2PPT_URL), width=120, height=26).pack(
            side="left", padx=(8, 0))

        self._label(page, 3, "기본 성경 역본")
        self.bible = BibleField(page, values=self.bible_values, width=self.FIELD_W,
                                placeholder="비우면 Bible2PPT 마지막 선택값", clear_text="지정 안 함")
        self.bible.set(b.get("bible") or "")
        self.bible.grid(row=3, column=1, sticky="w", pady=(4, 2))
        secondary_button(page, "↻", self._load_bibles, width=36).grid(row=3, column=2, padx=(6, 0), pady=(4, 2))
        self._label(page, 4, "두 번째 역본")
        self.bible2 = BibleField(page, values=self.bible_values, width=self.FIELD_W,
                                 placeholder="나란히 표시할 역본 (선택)", clear_text="지정 안 함")
        self.bible2.set(b.get("bible2") or "")
        self.bible2.grid(row=4, column=1, sticky="w", pady=(4, 2))
        self.bible_hint = self._note(page, 5)
        self._show_bible_hint()

        self._label(page, 6, "성경 템플릿")
        self.tpl = self._entry(page, b.get("template"))
        self.tpl.grid(row=6, column=1, sticky="w", pady=(4, 2))
        secondary_button(page, "찾기", lambda: self._browse(self.tpl, [("PowerPoint", "*.pptx")]), width=56).grid(
            row=6, column=2, padx=(6, 0), pady=(4, 2))
        self._note(page, 7, "선택 — 말씀 슬라이드 모양을 정하는 .pptx. 비우면 Bible2PPT 기본 템플릿을 씁니다.")
        self._label(page, 8, "슬라이드당 줄 수")
        self.lines = ctk.CTkComboBox(page, width=160, values=[LINES_DEFAULT] + [str(i) for i in range(10)],
                                     state="readonly", font=font(13), dropdown_font=font(13))
        self.lines.set(LINES_DEFAULT if b.get("lines") is None else str(b.get("lines")))
        self.lines.grid(row=8, column=1, sticky="w", pady=(4, 2))
        self._note(page, 9, "한 슬라이드에 넣을 구절 줄 수. 프로그램 설정 = Bible2PPT 에서 정한 값, 0 = 제한 없음.")
        return page

    def _build_summary(self):
        page = self._page()
        self._intro(page, "아래 내용으로 저장합니다. 고칠 것이 있으면 [이전]으로 돌아가세요.")
        self.summary = ctk.CTkFrame(page, fg_color="transparent")
        self.summary.grid(row=1, column=0, columnspan=3, sticky="nw")
        ctk.CTkLabel(page, text="[마침]을 누르면 저장합니다. 도움말 옆 [설정 마법사]로 언제든 다시 할 수 있고,\n"
                                "자세한 설정은 ⚙ 설정에서 바꿀 수 있습니다.",
                     font=font(11), text_color=TEXT2, anchor="w", justify="left").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))
        return page

    def _fill_summary(self):
        for w in self.summary.winfo_children():
            w.destroy()
        rows: list[tuple[bool, str, str]] = []
        for key, name, _desc, _rec in WIZARD_FOLDERS:
            n = self._ppt_count(key)
            rows.append((bool(n), name, self._count_text(n) if n != 0 else "PPT 파일이 없습니다 — 나중에 넣어 주세요"))
        sermon = path_text(self.sermon_file)
        rows.append((bool(sermon), "설교 템플릿", Path(sermon).name if sermon else "지정 안 함 — 설교 순서를 추가할 수 없음"))
        exe = path_text(self.exe)
        rows.append((bool(exe), "Bible2PPT", Path(exe).name if exe else "지정 안 함 — 말씀 순서를 추가할 수 없음"))
        bibles = [v for v in (self.bible.get().strip(), self.bible2.get().strip()) if v]
        rows.append((True, "기본 역본", " + ".join(bibles) if bibles else "지정 안 함 (Bible2PPT 마지막 선택값)"))
        tpl = path_text(self.tpl)
        rows.append((True, "성경 템플릿", Path(tpl).name if tpl else "Bible2PPT 기본"))
        rows.append((True, "슬라이드당 줄 수", self.lines.get()))
        for r, (ok, name, detail) in enumerate(rows):
            ctk.CTkLabel(self.summary, text="✔" if ok else "⚠", font=font(13, "bold"),
                         text_color=KIND_COLORS[KIND_PRAYER] if ok else WARN_ICON, width=22).grid(row=r, column=0, pady=3)
            ctk.CTkLabel(self.summary, text=name, font=font(13, "bold"), text_color=TEXT, anchor="w", width=130).grid(
                row=r, column=1, sticky="w", padx=(6, 0), pady=3)
            ctk.CTkLabel(self.summary, text=detail, font=font(13), text_color=TEXT if ok else TEXT2, anchor="w",
                         wraplength=self.FIELD_W).grid(row=r, column=2, sticky="w", pady=3)

    # ---- 폴더
    def _folder_path(self, key: str) -> Path | None:
        raw = path_text(self.folders[key][0])
        return self.cfg.resolve_file(raw) if raw else None

    def _ppt_count(self, key: str) -> int | None:
        """PPT 개수, 폴더가 너무 커서 세다가 그만두면 None (화면이 멈추지 않게 count_ppt_files 의 시간 제한)."""
        p = self._folder_path(key)
        return count_ppt_files(p, recursive=self.folders[key][3]) if p else 0

    @staticmethod
    def _count_text(n: int | None) -> str:
        if n is None:
            return "폴더가 너무 커서 개수를 세지 않았습니다 — PPT 만 모아 둔 폴더를 지정하면 찾기가 빨라집니다"
        return f"PPT {n}개" if n else "아직 PPT 가 없습니다"

    def _update_folder_note(self, key: str):
        desc, note = self._desc[key], self.folders[key][2]
        p = self._folder_path(key)
        if p is None or not p.is_dir():
            note.configure(text=f"{desc}\n⚠ 폴더를 찾을 수 없습니다", text_color=DANGER)
            return
        n = self._ppt_count(key)
        note.configure(text=f"{desc}\n{self._count_text(n)}", text_color=TEXT2 if n is not None else WARN_ICON)

    def _choose_folder(self, key: str):
        e = self.folders[key][0]
        p = self._folder_path(key)
        path = filedialog.askdirectory(parent=self, initialdir=str(p) if p and p.is_dir() else None, mustexist=True)
        if path:
            e.delete(0, "end")
            e.insert(0, str(Path(path)))
            self._update_folder_note(key)

    # ---- 단계 이동
    def _show_step(self, i: int):
        self.step = i
        for j, page in enumerate(self.pages):
            if j == i:
                page.grid()
            else:
                page.grid_remove()
        for j, lbl in enumerate(self._step_labels):
            lbl.configure(text_color=ACCENT if j == i else (TEXT if j < i else TEXT2),
                          font=font(13, "bold") if j == i else font(13))
        self.prev_btn.configure(state="normal" if i else "disabled")
        self.next_btn.configure(text="마침" if i == len(self.pages) - 1 else "다음")
        if i == len(self.pages) - 1:
            self._fill_summary()
        # 숨긴 쪽의 입력칸에 포커스가 남으면 보이지 않는 칸에 계속 입력됨(검사를 마친 값이 몰래 바뀜) → 다이얼로그로 옮김.
        # Tab 이동은 보이지 않는 위젯을 건너뛰므로 여기서부터는 지금 쪽 안에서만 움직인다.
        self.focus_set()

    def _prev(self):
        if self.step:
            self._show_step(self.step - 1)

    def _next(self):
        last = self.step == len(self.pages) - 1
        # [마침]에서는 모든 단계를 다시 검사 — 앞 단계 값은 [다음] 때 검사했지만 그 뒤에 바뀌었을 수 있으므로
        # 저장하는 값은 반드시 지금 검사한 값이어야 한다 (실측: 숨긴 칸에 남은 포커스로 .bat 경로가 검사 없이 저장됐었음).
        for i in (range(len(self.pages)) if last else (self.step,)):
            if (problem := self._check_step(i)) is not None:
                widget, msg = problem
                if i != self.step:
                    self._show_step(i)
                if widget is not None:
                    widget.focus_set()
                messagebox.showerror("입력 확인", msg, parent=self)
                return
        if not last:
            self._show_step(self.step + 1)
        else:
            self._apply(self.cfg)
            self.cfg.save()
            self.result = True
            self.destroy()

    def _check_step(self, i: int):
        """(문제 칸, 메시지) 또는 None."""
        if i == 0:
            for key, name, _desc, _rec in WIZARD_FOLDERS:
                e = self.folders[key][0]
                p = self._folder_path(key)
                if p is None:
                    return e, f"{name}를 지정해 주세요."
                if not p.is_dir():
                    return e, f"{name}를 찾을 수 없습니다:\n{p}"
                if (err := folder_conflict(self.cfg, key, p)) is not None:
                    return e, err
            probe = self._probe_cfg()  # 바꾼 폴더가 출력 폴더와 겹쳐 결과가 원본을 덮어쓰지 않는지
            if (err := output_conflict(probe, probe.resolve_path("output_dir") / "_.pptx")) is not None:
                return None, err + "\n(출력 폴더는 ⚙ 설정 → 폴더에서 바꿀 수 있습니다.)"
        elif i == 1:
            raw = path_text(self.sermon_file)
            if raw:
                p = self.cfg.resolve_file(raw)
                if not p.is_file() or p.suffix.lower() not in PPT_EXTS:
                    return self.sermon_file, f"설교 템플릿 PPT 파일을 찾을 수 없습니다:\n{raw}"
        elif i == 2:
            raw = path_text(self.exe)
            if raw and (not raw.lower().endswith(".exe") or self._exe() is None):
                return self.exe, f"Bible2PPT 실행 파일(.exe)을 찾을 수 없습니다:\n{raw}"
            raw = path_text(self.tpl)
            if raw:
                p = self.cfg.resolve_file(raw)
                if p.suffix.lower() != ".pptx" or not p.is_file():
                    return self.tpl, ("성경 템플릿은 존재하는 .pptx 파일이어야 합니다\n"
                                      f"(매크로가 들어갈 수 있는 .pptm 등은 사용할 수 없음):\n{raw}")
            if self.bible2.get().strip() and not self.bible.get().strip():
                return self.bible.entry, "두 번째 역본을 쓰려면 기본 성경 역본도 지정해 주세요."
        else:
            if (err := temp_dir_conflict(self._probe_cfg())) is not None:
                return None, err
        return None

    # ---- 저장
    def _probe_cfg(self):
        """지금 입력한 값을 넣은 설정 사본 (검사용 — 파일에 저장하지 않음)."""
        probe = copy.copy(self.cfg)
        probe.data = copy.deepcopy(self.cfg.data)
        self._apply(probe)
        return probe

    def _apply(self, cfg) -> None:
        paths = cfg.data["paths"]
        for key, (e, shown, _note, _rec) in self.folders.items():
            raw = path_text(e)
            if raw != shown:  # 그대로면 원래 값 유지 (기본값 'templates' 같은 상대 경로는 exe 를 옮겨도 따라가도록)
                paths[key] = raw
        cfg.data["item_templates"]["sermon_file"] = path_text(self.sermon_file)
        lines = self.lines.get()
        cfg.data["bible2ppt"].update(
            exe_path=path_text(self.exe), bible=self.bible.get().strip(), bible2=self.bible2.get().strip(),
            template=path_text(self.tpl), lines=None if lines == LINES_DEFAULT else int(lines),
            # 설정 창과 같은 규칙: 지금 exe 로 불러온 목록만 캐시 (못 불러왔으면 비움 → 메인에서 다시 불러옴)
            bible_list=self.bible_values if self._loaded_for is not None and self._loaded_for == self._exe() else [])

    def _bible_timeout(self) -> int:
        return max(int(self.cfg.get("bible2ppt", "timeout_sec") or 180), 10)

    def on_shown(self):
        self._load_bibles_if_needed()


# ---------------------------------------------------------------- 로그
class TempFilesDialog(_Modal):
    """[미리 보기 (저장 안 함)]로 만든 PPT 목록 (최근 것부터). 열기·쇼 시작·다른 이름으로 저장·폴더 열기.
    실제 동작은 App 콜백(path → 성공 여부)이 하고, 성공하면 창을 닫는다. 목록은 tk.Listbox 하나 (CCM 선택과 같은 이유)."""
    ROWS = 10

    def __init__(self, master, files: list[Path], keep: int, on_open: Callable[[Path], bool],
                 on_show: Callable[[Path], bool], on_save: Callable[[Path], bool], on_folder: Callable[[], None]):
        super().__init__(master, "임시 파일 불러오기")
        self.files = list(files)
        self._actions = {"open": on_open, "show": on_show, "save": on_save}
        body = card(self)
        body.pack(fill="both", expand=True, padx=16, pady=(12, 8))
        box = ctk.CTkFrame(body, fg_color=SUBTLE, corner_radius=8, border_width=1, border_color=BORDER)
        box.pack(fill="both", expand=True, padx=12, pady=(12, 4))
        self.listbox = tk.Listbox(box, height=self.ROWS, width=72, activestyle="none", borderwidth=0,
                                  highlightthickness=0, exportselection=False, font=tk_font(self, 13),
                                  bg=_pick(SUBTLE), fg=_pick(TEXT), selectbackground=_pick(ACCENT),
                                  selectforeground="#FFFFFF")
        sb = ctk.CTkScrollbar(box, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=6)
        sb.pack(side="right", fill="y", padx=2, pady=4)
        for p in self.files:
            self.listbox.insert("end", f"  {self._describe(p)}")
        if self.files:
            self.listbox.selection_set(0)
            self.listbox.activate(0)
            info = f"{len(self.files)}개 · 최근 {keep}개까지 보관 (⚙ 설정에서 변경)   ·   더블클릭 또는 Enter 로 열기"
        else:
            info = "임시 파일이 없습니다. 하단 [미리 보기 (저장 안 함)]로 만들면 여기에 쌓입니다."
        ctk.CTkLabel(body, text=info, font=font(11), text_color=TEXT2, anchor="w").pack(
            fill="x", padx=14, pady=(0, 10))
        self.listbox.bind("<Double-Button-1>", lambda e: self._run("open"))

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=16, pady=(4, 16))
        secondary_button(btns, "폴더 열기", on_folder, width=88).pack(side="left")
        actions = [accent_button(btns, "열기", lambda: self._run("open"), width=88),
                   secondary_button(btns, "슬라이드 쇼", lambda: self._run("show"), width=96),
                   secondary_button(btns, "다른 이름으로 저장", lambda: self._run("save"), width=128)]
        for b in actions:
            b.pack(side="right", padx=(8, 0))
            if not self.files:
                b.configure(state="disabled")
        secondary_button(btns, "닫기", self.destroy, width=72).pack(side="right", padx=(8, 0))
        self.bind_return(lambda: self._run("open"))

    @staticmethod
    def _describe(p: Path) -> str:
        try:
            st = p.stat()
            when = dt.datetime.fromtimestamp(st.st_mtime)
            return f"{when:%m-%d %H:%M}   {p.stem}   ({st.st_size / 1048576:.1f}MB)"
        except OSError:
            return p.stem

    def on_shown(self):
        self.listbox.focus_set()

    def _run(self, key: str):
        cur = self.listbox.curselection()
        if not cur:
            return
        if self._actions[key](self.files[cur[0]]):
            self.destroy()


class LogDialog(_Modal):
    def __init__(self, master, lines: list[str]):
        super().__init__(master, "작업 로그")
        box = ctk.CTkTextbox(self, width=680, height=340, font=("Consolas", 12), corner_radius=8,
                             border_width=1, border_color=BORDER)
        box.pack(padx=20, pady=12)
        box.insert("1.0", "\n".join(lines) or "로그가 없습니다.")
        box.configure(state="disabled")
        accent_button(self, "닫기", self.destroy, width=96).pack(anchor="e", padx=20, pady=(0, 16))


# ---------------------------------------------------------------- 도움말
class HelpDialog(_Modal):
    """도움말 (help_content.HELP_SECTIONS). 왼쪽 목차 클릭 → 그 절로 이동, 위 검색칸 Enter → 다음 찾기.
    본문은 Text 위젯 하나에 태그로 꾸며 그린다 (절·표마다 위젯을 만들면 수백 개가 되어 느림 — CLAUDE.md §11).
    CTkTextbox 는 태그 글꼴을 막아 두었으므로(배율 때문) 안쪽 tk.Text 에 배율을 반영한 픽셀 글꼴(tk_font)로 직접 지정."""
    _MARKUP = re.compile(r"(`[^`]+`|\*\*[^*]+\*\*)")

    def __init__(self, master):
        from help_content import HELP_SECTIONS  # 도움말을 열 때만 읽음
        super().__init__(master, "도움말")
        self._sections = HELP_SECTIONS
        self._last_found = "1.0"
        self._chosen: int | None = None  # 목차에서 마지막으로 고른 절
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=(10, 0))
        # 창이 뜨면 검색칸에 포커스가 가서 안내 문구(placeholder)가 안 보이므로 앞에 이름표를 둠
        ctk.CTkLabel(top, text="🔍 검색", font=font(13), text_color=TEXT2).pack(side="left", padx=(0, 8))
        self.search = ctk.CTkEntry(top, width=260, height=32, font=font(13), placeholder_text="찾을 낱말 (Enter: 다음)")
        self.search.pack(side="left")
        self.search.bind("<Return>", lambda e: self._find_next())
        self.search.bind("<KeyRelease>", self._on_search_key)
        self.found_label = ctk.CTkLabel(top, text="", font=font(12), text_color=TEXT2)
        self.found_label.pack(side="left", padx=10)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=10)
        nav_box = ctk.CTkFrame(body, fg_color=CARD, corner_radius=8, border_width=1, border_color=BORDER)
        nav_box.pack(side="left", fill="y", padx=(0, 10))
        self.nav = tk.Listbox(nav_box, activestyle="none", borderwidth=0, highlightthickness=0, exportselection=False,
                              font=tk_font(self, 13), width=0, cursor="hand2", bg=_pick(CARD), fg=_pick(TEXT),
                              selectbackground=_pick(ACCENT), selectforeground="#FFFFFF")
        self.nav.pack(fill="both", expand=True, padx=scaled(self, 6), pady=scaled(self, 6))
        for title, _ in self._sections:
            self.nav.insert("end", "  " + title)
        self.nav.bind("<<ListboxSelect>>", lambda e: self._go(self.nav.curselection()))

        win_h = self._win.winfo_height() / ctk.ScalingTracker.get_widget_scaling(self)
        self.box = ctk.CTkTextbox(body, width=640, height=int(max(300, win_h * 0.9 - 190)), font=font(13),
                                  corner_radius=8, border_width=1, border_color=BORDER, fg_color=CARD, wrap="word",
                                  activate_scrollbars=True)
        self.box.pack(side="left", fill="both", expand=True)
        self.text = self.box._textbox
        self.text.configure(spacing1=scaled(self, 2), spacing3=scaled(self, 2), padx=scaled(self, 14),
                            pady=scaled(self, 10), cursor="arrow")
        sb_set = self.box._y_scrollbar.set
        self.text.configure(yscrollcommand=lambda a, b: (sb_set(a, b), self._sync_nav()))
        self._render()
        self.text.configure(state="disabled")
        self.box.bind("<Button-1>", lambda e: self.text.focus_set())  # 휠·키보드 스크롤
        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=20, pady=(0, 16))
        accent_button(btns, "닫기", self.destroy, width=96).pack(side="right")
        self.nav.selection_set(0)

    # ---- 그리기
    def _render(self):
        t = self.text
        base = tk_font(self, 13)
        t.tag_configure("h1", font=tk_font(self, 19, "bold"), foreground=_pick(TEXT),
                        spacing1=scaled(self, 18), spacing3=scaled(self, 8))
        t.tag_configure("h2", font=tk_font(self, 15, "bold"), foreground=_pick(TEXT),
                        spacing1=scaled(self, 14), spacing3=scaled(self, 4))
        t.tag_configure("bold", font=tk_font(self, 13, "bold"))
        t.tag_configure("code", foreground=_pick(ACCENT), font=tk_font(self, 13, "bold"))
        t.tag_configure("note", foreground=_pick(TEXT2), font=tk_font(self, 12),
                        lmargin1=scaled(self, 10), lmargin2=scaled(self, 10))
        t.tag_configure("found", background=_pick(("#FFE58F", "#7A5C00")))
        t.tag_raise("code", "note")  # 참고 문장(흐린 글씨) 안의 `글자` 도 강조색으로
        for level in (0, 1):
            ind = scaled(self, 8 + 18 * level)
            t.tag_configure(f"bullet{level}", lmargin1=ind, lmargin2=ind + scaled(self, 16))
        measure = tkfont.Font(font=base)
        bold_measure = tkfont.Font(font=tk_font(self, 13, "bold"))
        for i, (title, text) in enumerate(self._sections):
            if i:
                t.insert("end", "\n")
            t.mark_set(f"sec{i}", "end-1c")
            t.mark_gravity(f"sec{i}", "left")
            t.insert("end", title + "\n", "h1")
            lines = text.rstrip("\n").split("\n")
            n = 0
            while n < len(lines):
                line = lines[n]
                if line.startswith("|"):  # 표: 이어지는 | 줄을 모아 열 너비를 맞춤
                    rows = []
                    while n < len(lines) and lines[n].startswith("|"):
                        rows.append([c.strip() for c in lines[n].strip().strip("|").split("|")])
                        n += 1
                    self._table(rows, f"tbl{i}_{n}", measure, bold_measure)
                    continue
                if line.startswith("## "):
                    t.insert("end", line[3:] + "\n", "h2")
                elif line.lstrip().startswith("- "):
                    level = 1 if line.startswith("  ") else 0
                    self._inline("•  " + line.lstrip()[2:], (f"bullet{level}",))
                elif line.startswith("> "):
                    self._inline(line[2:], ("note",))
                else:
                    self._inline(line, ())
                n += 1

    def _inline(self, text: str, tags: tuple, end: str = "\n"):
        for part in self._MARKUP.split(text):
            if not part:
                continue
            if part.startswith("`") and part.endswith("`") and len(part) > 1:
                self.text.insert("end", part[1:-1], tags + ("code",))
            elif part.startswith("**") and part.endswith("**") and len(part) > 3:
                self.text.insert("end", part[2:-2], tags + ("bold",))
            else:
                self.text.insert("end", part, tags)
        if end:
            self.text.insert("end", end, tags)

    def _table(self, rows: list[list[str]], tag: str, measure, bold_measure):
        """칸 사이는 탭, 탭 위치는 각 열의 가장 긴 글자에 맞춤 (너무 긴 열은 잘라서 다음 열이 밀려나지 않게).
        마지막 열이 길어 줄이 바뀌면 그 열 시작 위치에 맞춰 이어짐(lmargin2)."""
        plain = [[self._MARKUP.sub(lambda m: m.group(0).strip("`*"), c) for c in r] for r in rows]
        cols = max(len(r) for r in plain)
        gap, cap = scaled(self, 22), scaled(self, 300)
        indent = scaled(self, 8)
        stops, x = [], indent
        for c in range(cols - 1):
            w = max((bold_measure if ri == 0 else measure).measure(r[c]) for ri, r in enumerate(plain) if c < len(r))
            x += min(w, cap) + gap
            stops.append(x)
        self.text.tag_configure(tag, tabs=tuple(stops) or None, lmargin1=indent, lmargin2=stops[-1] if stops else indent,
                                spacing1=scaled(self, 3), spacing3=scaled(self, 3))
        self.text.tag_configure(tag + "_head", font=tk_font(self, 13, "bold"))
        for ri, r in enumerate(rows):
            extra = (tag + "_head",) if ri == 0 else ()
            for c, cell in enumerate(r):
                if c:
                    self.text.insert("end", "\t", (tag,) + extra)
                if ri == 0:
                    self.text.insert("end", self._MARKUP.sub(lambda m: m.group(0).strip("`*"), cell), (tag,) + extra)
                else:
                    self._inline(cell, (tag,), end="")
            self.text.insert("end", "\n", (tag,) + extra)

    # ---- 목차·검색
    def _go(self, sel):
        if sel:
            self._chosen = sel[0]
            self.text.yview(f"sec{sel[0]}")
            self._sync_nav()  # 이미 끝까지 내려가 있으면 스크롤 이벤트가 없으므로 직접 맞춤

    def _sync_nav(self):
        """보이는 맨 위 줄이 속한 절을 목차에서 선택 표시. 끝까지 내렸으면 마지막 절은 맨 위로 올 수 없으므로
        화면에 제목이 보이는 마지막 절을 고름 (목차에서 마지막 절을 눌렀을 때 그 앞 절이 선택되지 않게)."""
        top = self.text.index("@0,0")
        at_end = self.text.yview()[1] >= 0.999
        bottom = self.text.index(f"@0,{self.text.winfo_height()}")
        current = 0
        for i in range(len(self._sections)):
            if self.text.compare(f"sec{i}", "<=", bottom if at_end else top):
                current = i
        # 끝부분 절을 목차에서 골랐으면(맨 위로 올 수 없어 뒤 절 제목도 보임) 고른 절 제목이 보이는 동안은 그 절로
        chosen = self._chosen
        if at_end and chosen is not None and self.text.compare(f"sec{chosen}", ">=", top)                 and self.text.compare(f"sec{chosen}", "<=", bottom):
            current = chosen
        if self.nav.curselection() != (current,):
            self.nav.selection_clear(0, "end")
            self.nav.selection_set(current)
            self.nav.see(current)

    def _on_search_key(self, event):
        if event.keysym in ("Return", "Up", "Down", "Left", "Right"):
            return
        self._last_found = "1.0"  # 검색어가 바뀌면 처음부터
        self._highlight()

    def _highlight(self) -> int:
        t = self.text
        t.tag_remove("found", "1.0", "end")
        word = self.search.get().strip()
        if not word:
            self.found_label.configure(text="")
            return 0
        count, pos = 0, "1.0"
        while True:
            pos = t.search(word, pos, "end", nocase=True)
            if not pos:
                break
            end = f"{pos}+{len(word)}c"
            t.tag_add("found", pos, end)
            pos, count = end, count + 1
        self.found_label.configure(text=f"{count}건" if count else "찾는 내용이 없습니다", text_color=TEXT2 if count else DANGER)
        return count

    def _find_next(self):
        if not self._highlight():
            return
        t = self.text
        start = self._last_found
        pos = t.search(self.search.get().strip(), f"{start}+1c", "end", nocase=True) or \
            t.search(self.search.get().strip(), "1.0", "end", nocase=True)
        if pos:
            self._last_found = pos
            t.see(pos)

    def on_shown(self):
        self.search.focus_set()


# ---------------------------------------------------------------- 드래그 & 드롭
class DragController:
    """마우스 이벤트 기반 드래그&드롭 (tkinter 는 위젯 간 DnD 를 기본 제공하지 않음).

    attach(widget, payload, text) 로 드래그 소스를 등록한다.
    payload 가 함수면 누를 때 event 로 (payload, text) 를 구한다 (목록 상자처럼 한 위젯에 항목이 여럿일 때). None 이면 무시.
    - 6px 이상 움직이면 드래그 시작 → 반투명 고스트 라벨이 커서를 따라감
    - on_hover(x_root, y_root, payload) : 드래그 중 호출 (드롭 위치 표시/자동 스크롤)
    - on_drop(x_root, y_root, payload)  : 놓았을 때
    - on_click(payload)                 : 드래그 없이 클릭만 했을 때
    """

    THRESHOLD = 6

    def __init__(self, root: tk.Misc, on_hover, on_drop, on_click=None, on_cancel=None):
        self.root = root
        self.on_hover = on_hover
        self.on_drop = on_drop
        self.on_click = on_click
        self.on_cancel = on_cancel
        self._start = None
        self._payload = None
        self._text = ""
        self._ghost: tk.Toplevel | None = None

    def attach(self, widget, payload, text: str = ""):
        def press(e):
            if not callable(payload):
                self._press(e, payload, text)
            elif (got := payload(e)) is not None:
                self._press(e, *got)
        widget.bind("<ButtonPress-1>", press, add="+")
        widget.bind("<B1-Motion>", self._motion, add="+")
        widget.bind("<ButtonRelease-1>", self._release, add="+")

    @property
    def dragging(self) -> bool:
        return self._ghost is not None

    def _press(self, e, payload, text):
        self._start = (e.x_root, e.y_root)
        self._payload = payload
        self._text = text

    def _motion(self, e):
        if self._start is None:
            return
        if self._ghost is None:
            if abs(e.x_root - self._start[0]) + abs(e.y_root - self._start[1]) < self.THRESHOLD:
                return
            self._make_ghost()
        self._ghost.geometry(f"+{e.x_root + 14}+{e.y_root + 10}")
        self.on_hover(e.x_root, e.y_root, self._payload)

    def _release(self, e):
        if self._start is None:
            return
        payload, was_drag = self._payload, self._ghost is not None
        self._start = None
        self._payload = None
        if self._ghost is not None:
            self._ghost.destroy()
            self._ghost = None
        try:
            self.root.configure(cursor="")
        except tk.TclError:
            pass
        if was_drag:
            # 위젯 재구성이 이벤트 처리 중에 일어나지 않도록 다음 루프에서 처리
            self.root.after_idle(lambda: self.on_drop(e.x_root, e.y_root, payload))
        elif self.on_click:
            self.root.after_idle(lambda: self.on_click(payload))

    def _make_ghost(self):
        g = tk.Toplevel(self.root)
        g.overrideredirect(True)
        g.attributes("-topmost", True)
        try:
            g.attributes("-alpha", 0.88)
        except tk.TclError:
            pass
        tk.Label(g, text=f"  {self._text}  ", bg=_pick(ACCENT), fg="#FFFFFF",
                 font=tk_font(self.root, 15, "bold"), padx=scaled(self.root, 6), pady=scaled(self.root, 4)).pack()
        self._ghost = g
        try:
            self.root.configure(cursor="fleur")
        except tk.TclError:
            pass


_ui_queue: "queue.Queue[Callable[[], None]]" = queue.Queue()


def post_to_ui(fn: Callable[[], None]) -> None:
    """작업 스레드 → 메인 스레드 호출 (tkinter 는 스레드 안전하지 않음)."""
    _ui_queue.put(fn)


def pump_ui_queue(root: tk.Misc, interval_ms: int = 50) -> None:
    """메인 스레드에서 한 번 호출하면 주기적으로 큐를 비운다."""
    while True:
        try:
            fn = _ui_queue.get_nowait()
        except queue.Empty:
            break
        try:
            fn()
        except Exception:  # noqa: BLE001
            import traceback
            traceback.print_exc()
    root.after(interval_ms, lambda: pump_ui_queue(root, interval_ms))


def run_in_thread(work: Callable[[], Any], done: Callable[[Any, BaseException | None], None]):
    """work() 를 작업 스레드에서 실행하고 결과를 메인 스레드의 done(result, error) 로 전달."""

    def runner():
        try:
            res, err = work(), None
        except BaseException as ex:  # noqa: BLE001 — UI 로 그대로 전달
            res, err = None, ex
        post_to_ui(lambda: done(res, err))

    threading.Thread(target=runner, daemon=True).start()
