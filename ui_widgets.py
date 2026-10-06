"""재사용 UI 위젯: 테마 상수, 달력 팝업, 항목 입력 다이얼로그, 설정 창, 드래그&드롭."""
from __future__ import annotations

import calendar
import datetime as dt
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Any, Callable

import customtkinter as ctk

from models import KIND_BIBLE, KIND_CUSTOM, KIND_HYMN, KIND_LABELS, KIND_PRAYER, ServiceItem, make_item, validate_item

# ---------------------------------------------------------------- Win11 Fluent 톤 색상 (light, dark)
FONT_FAMILY = "Malgun Gothic"
ACCENT = ("#0067C0", "#2B88D8")
ACCENT_HOVER = ("#1975C5", "#3C95E0")
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
    KIND_CUSTOM: ("#0067C0", "#60CDFF"),
}


MERGE_MODES = {
    "원본 서식 유지 (권장, PowerPoint 최소화 창 사용)": "source_formatting",
    "창 없이 병합 (배경은 복제/이미지화)": "windowless",
}


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
    return ctk.CTkButton(master, text=text, command=command, width=width, height=32, corner_radius=6,
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


class _Modal(ctk.CTkFrame):
    """메인 창 안에 뜨는 모달 다이얼로그 (Win11 ContentDialog 스타일).

    별도 창(CTkToplevel)을 쓰지 않는다. CTkToplevel 은 생성 시 창 숨김/강제 update/ctypes 타이틀바 조작과
    해제되지 않는 bind_all 을 수행하는데, 반복 생성·파괴 시 Tk 접근 위반 크래시가 재현되었다.
    다이얼로그에 grab 을 걸어 바깥 위젯 클릭을 막고, 메인 화면은 뒤에 그대로 보이게 한다.
    사용: result = SomeDialog(app, ...).show()   (닫힐 때까지 대기)
    """

    def __init__(self, master, title: str):
        self._win = master.winfo_toplevel()
        super().__init__(self._win, corner_radius=12, fg_color=BG, border_width=2, border_color=PILL_BORDER)
        self.result: Any = None
        self._closing = False
        ctk.CTkLabel(self, text=title, font=font(16, "bold"), text_color=TEXT, anchor="w").pack(
            fill="x", padx=20, pady=(16, 0))
        self._keys = [("<Escape>", self._win.bind("<Escape>", lambda e: self.destroy(), add="+"))]

    def show(self):
        self.place(relx=0.5, rely=0.45, anchor="center")
        self.lift()
        self.update_idletasks()
        try:
            self.grab_set()  # 다이얼로그 바깥 위젯 클릭 차단
        except tk.TclError:
            pass
        self.after(30, self.on_shown)
        self._win.wait_window(self)
        return self.result

    def bind_return(self, fn):
        self._keys.append(("<Return>", self._win.bind("<Return>", lambda e: fn(), add="+")))

    def on_shown(self):
        pass

    def destroy(self):
        if self._closing:
            return
        self._closing = True
        for seq, funcid in self._keys:
            try:
                self._win.unbind(seq, funcid)
            except tk.TclError:
                pass
        try:
            self.grab_release()
        except tk.TclError:
            pass
        super().destroy()


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


# ---------------------------------------------------------------- 항목 입력 다이얼로그
class ItemDialog(_Modal):
    """기도/찬송/말씀/사용자 지정 입력. 필수값 미입력 시 에러 창."""

    def __init__(self, master, kind: str, item: ServiceItem | None = None,
                 template_names: list[str] | None = None, bible_values: list[str] | None = None,
                 default_bible: str = ""):
        editing = item is not None
        super().__init__(master, f"{KIND_LABELS[kind]} {'수정' if editing else '추가'}")
        self.kind = kind
        self.options = dict(item.options) if item else {}

        body = card(self)
        body.pack(fill="both", expand=True, padx=16, pady=(16, 8))
        ctk.CTkLabel(body, text=f"{KIND_LABELS[kind]} 항목", font=font(15, "bold"),
                     text_color=KIND_COLORS[kind]).grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 8))

        self.label_widget = None
        self.bible_widget = None
        row = 1
        if kind == KIND_CUSTOM:
            self._lbl(body, row, "호출할 PPT 이름 *")
            self.label_widget = ctk.CTkComboBox(body, width=260, values=template_names or [], font=font(13),
                                                dropdown_font=font(13))
            self.label_widget.set(item.label if item else "")
            self.label_widget.grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")
            row += 1

        prompt, hint = {
            KIND_PRAYER: ("기도자 이름 *", "예: 홍길동 장로"),
            KIND_HYMN: ("찬송가 장 *", "예: 29  또는  29장"),
            KIND_BIBLE: ("성경 구절 *", "예: 창1:1-5, 롬1-3"),
            KIND_CUSTOM: ("오른쪽 문구 (선택)", "예: 인도자"),
        }[kind]
        self._lbl(body, row, prompt)
        self.value_entry = ctk.CTkEntry(body, width=260, placeholder_text=hint, font=font(13))
        if item and item.value:
            self.value_entry.insert(0, item.value)
        self.value_entry.grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")
        row += 1

        if kind == KIND_BIBLE:
            self._lbl(body, row, "성경 역본 (선택)")
            self.bible_widget = ctk.CTkComboBox(body, width=260, values=bible_values or [], font=font(13),
                                                dropdown_font=font(13))
            self.bible_widget.set(self.options.get("bible", default_bible))
            self.bible_widget.grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")
            row += 1
            ctk.CTkLabel(body, text="여러 구절은 쉼표로 구분합니다. 역본을 비우면 설정의 기본 역본을 사용합니다.",
                         font=font(11), text_color=TEXT2).grid(row=row, column=0, columnspan=2, padx=16, sticky="w")
            row += 1
        if kind == KIND_CUSTOM:
            ctk.CTkLabel(body, text="템플릿 폴더에서 이 이름의 PPT 를 찾아 호출합니다.",
                         font=font(11), text_color=TEXT2).grid(row=row, column=0, columnspan=2, padx=16, sticky="w")
            row += 1
        ctk.CTkFrame(body, width=1, height=8, fg_color="transparent").grid(row=row, column=0)

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=16, pady=(4, 16))
        accent_button(btns, "확인", self._ok, width=96).pack(side="right")
        secondary_button(btns, "취소", self.destroy, width=96).pack(side="right", padx=8)
        self.bind_return(self._ok)

    def on_shown(self):
        (self.label_widget if self.label_widget is not None else self.value_entry).focus_set()

    @staticmethod
    def _lbl(master, row, text):
        ctk.CTkLabel(master, text=text, font=font(13), text_color=TEXT, anchor="w", width=130).grid(
            row=row, column=0, padx=16, pady=6, sticky="w")

    def _ok(self):
        label = self.label_widget.get() if self.label_widget is not None else KIND_LABELS[self.kind]
        if self.bible_widget is not None:
            b = self.bible_widget.get().strip()
            if b:
                self.options["bible"] = b
            else:
                self.options.pop("bible", None)
        item = make_item(self.kind, label, self.value_entry.get(), self.options)
        err = validate_item(item)
        if err:
            messagebox.showerror("입력 오류", err, parent=self)
            return
        self.result = item
        self.destroy()


# ---------------------------------------------------------------- 설정 창
class SettingsDialog(_Modal):
    def __init__(self, master, cfg):
        super().__init__(master, "설정")
        self.cfg = cfg
        b = cfg.get("bible2ppt")
        p = cfg.get("ppt")

        body = card(self)
        body.pack(fill="both", expand=True, padx=16, pady=(16, 8))
        r = 0

        def section(text):
            nonlocal r
            ctk.CTkLabel(body, text=text, font=font(14, "bold"), text_color=TEXT).grid(
                row=r, column=0, columnspan=3, sticky="w", padx=16, pady=(14, 4))
            r += 1

        def row(label, widget, browse=None):
            nonlocal r
            ctk.CTkLabel(body, text=label, font=font(13), text_color=TEXT, anchor="w", width=150).grid(
                row=r, column=0, padx=16, pady=5, sticky="w")
            widget.grid(row=r, column=1, pady=5, sticky="w")
            if browse:
                secondary_button(body, "찾기", browse, width=56).grid(row=r, column=2, padx=(6, 16), pady=5)
            r += 1

        def entry(value, width=320):
            e = ctk.CTkEntry(body, width=width, font=font(13))
            e.insert(0, value or "")
            return e

        section("Bible2PPT 연동")
        self.exe = entry(b.get("exe_path"))
        row("Bible2PPT.exe 경로", self.exe, lambda: self._browse(self.exe, [("Bible2PPT", "*.exe")]))
        self.bible = entry(b.get("bible"))
        row("기본 성경 역본", self.bible)
        self.tpl = entry(b.get("template"))
        row("성경 템플릿 (.pptx)", self.tpl, lambda: self._browse(self.tpl, [("PowerPoint", "*.pptx")]))
        self.lines = ctk.CTkComboBox(body, width=160, values=["프로그램 설정"] + [str(i) for i in range(10)],
                                     state="readonly", font=font(13), dropdown_font=font(13))
        self.lines.set("프로그램 설정" if b.get("lines") is None else str(b.get("lines")))
        row("슬라이드당 줄 수", self.lines)
        self.timeout = entry(str(b.get("timeout_sec", 180)), width=80)
        row("실행 제한 시간(초)", self.timeout)

        section("PPT 병합")
        self.merge_mode = ctk.CTkComboBox(body, width=320, values=list(MERGE_MODES), state="readonly",
                                          font=font(13), dropdown_font=font(13))
        self.merge_mode.set(next((k for k, v in MERGE_MODES.items() if v == p.get("merge_mode")),
                                 next(iter(MERGE_MODES))))
        row("병합 방식", self.merge_mode)

        section("텍스트 입력 규칙")
        self.placeholder = entry(p.get("text_placeholder"), width=200)
        row("치환 토큰", self.placeholder)
        self.shape_name = entry(p.get("text_shape_name"), width=200)
        row("텍스트 상자 이름", self.shape_name)
        ctk.CTkLabel(body, text="템플릿의 치환 토큰을 오른쪽 문구로 바꿉니다. 토큰이 없으면 같은 이름의 텍스트 상자에 입력합니다.\n"
                                "공통 토큰: {{교회이름}} {{날짜}} {{예배종류}}",
                     font=font(11), text_color=TEXT2, justify="left").grid(
            row=r, column=0, columnspan=3, padx=16, pady=(2, 14), sticky="w")

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=16, pady=(4, 16))
        accent_button(btns, "저장", self._save, width=96).pack(side="right")
        secondary_button(btns, "취소", self.destroy, width=96).pack(side="right", padx=8)

    def _browse(self, entry: ctk.CTkEntry, types):
        path = filedialog.askopenfilename(parent=self, filetypes=types + [("모든 파일", "*.*")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _save(self):
        try:
            timeout = int(self.timeout.get().strip() or 180)
            if timeout < 10:
                raise ValueError
        except ValueError:
            messagebox.showerror("입력 오류", "실행 제한 시간은 10 이상의 숫자로 입력해 주세요.", parent=self)
            return
        lines = self.lines.get()
        b = self.cfg.get("bible2ppt")
        b.update(exe_path=self.exe.get().strip(), bible=self.bible.get().strip(), template=self.tpl.get().strip(),
                 lines=None if lines == "프로그램 설정" else int(lines), timeout_sec=timeout)
        p = self.cfg.get("ppt")
        p.update(text_placeholder=self.placeholder.get().strip(), text_shape_name=self.shape_name.get().strip(),
                 merge_mode=MERGE_MODES[self.merge_mode.get()])
        self.cfg.save()
        self.result = True
        self.destroy()


# ---------------------------------------------------------------- 로그
class LogDialog(_Modal):
    def __init__(self, master, lines: list[str]):
        super().__init__(master, "작업 로그")
        box = ctk.CTkTextbox(self, width=680, height=340, font=("Consolas", 12), corner_radius=8,
                             border_width=1, border_color=BORDER)
        box.pack(padx=20, pady=12)
        box.insert("1.0", "\n".join(lines) or "로그가 없습니다.")
        box.configure(state="disabled")
        accent_button(self, "닫기", self.destroy, width=96).pack(anchor="e", padx=20, pady=(0, 16))


# ---------------------------------------------------------------- 드래그 & 드롭
class DragController:
    """마우스 이벤트 기반 드래그&드롭 (tkinter 는 위젯 간 DnD 를 기본 제공하지 않음).

    attach(widget, payload, text) 로 드래그 소스를 등록한다.
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

    def attach(self, widget, payload, text: str):
        widget.bind("<ButtonPress-1>", lambda e: self._press(e, payload, text), add="+")
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
                 font=(FONT_FAMILY, 11, "bold"), padx=6, pady=4).pack()
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
