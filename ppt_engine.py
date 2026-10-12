"""win32com 기반 PowerPoint 병합/치환 엔진.

원칙
- 원본 PPT는 ReadOnly 로만 연다 (절대 저장하지 않음).
- 첫 원본을 SaveAs 로 베이스 생성 (원본 100% 유지).
- 나머지 원본은 merge_mode 에 따라 이어 붙인다.
  * "source_formatting" (기본): PowerPoint 의 [원본 서식 유지] 붙여넣기 명령(PasteSourceFormatting)을 실행.
    테마·마스터·레이아웃·슬라이드 배경(그림 포함)까지 PowerPoint 가 직접 보존한다.
    명령 실행에 문서 창이 필요하므로 창을 하나 열되, 나타나는 즉시 숨긴다(_FrameHider — 화면에 깜빡이지 않음).
  * "windowless": 창 없이 Slides.Paste 후 슬라이드별 Design/CustomLayout 을 원본 것으로 지정하고
    슬라이드 배경을 복제(단색은 정확히, 그라데이션/그림 등은 고해상도 이미지로)한다.
  "source_formatting" 이 실패하면 해당 파일은 자동으로 "windowless" 로 처리한다.
- 애니메이션·전환·폰트 서식은 슬라이드와 함께 복사된다.
- 반드시 작업 스레드에서 with 문으로 사용한다 (내부에서 CoInitialize).
"""
from __future__ import annotations

import ctypes
import os
import shutil
import tempfile
import threading
import time
from ctypes import wintypes
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import pythoncom
import pywintypes
import win32com.client
import win32con
import win32gui
import win32process

from models import MAX_FILL, numbered_token

MSO_TRUE = -1
MSO_FALSE = 0
MSO_GROUP = 6
MSO_FILL_SOLID = 1
PP_SAVE_AS_PPTX = 24
PP_ALERTS_NONE = 1
PP_WINDOW_MAXIMIZED = 3
PP_VIEW_NORMAL = 9
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3  # 자동화로 여는 파일의 매크로 강제 비활성
_OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # OLE 복합 문서 (암호화된 .pptx 는 ZIP 대신 이 형식)

MERGE_SOURCE_FORMATTING = "source_formatting"
MERGE_WINDOWLESS = "windowless"

# 결과 슬라이드 비율 → (가로, 세로) pt. 16:9 는 현재 PowerPoint 기본(13.333×7.5in), 4:3 은 10×7.5in.
# 그 외 값(original)은 첫 PPT 크기 그대로.
SLIDE_SIZES = {"16:9": (960.0, 540.0), "4:3": (720.0, 540.0)}

LogFn = Callable[[str], None]

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_WINEVENTPROC = ctypes.WINFUNCTYPE(None, wintypes.HANDLE, wintypes.DWORD, wintypes.HWND, wintypes.LONG,
                                   wintypes.LONG, wintypes.DWORD, wintypes.DWORD)
_user32.SetWinEventHook.argtypes = (wintypes.DWORD, wintypes.DWORD, wintypes.HMODULE, _WINEVENTPROC,
                                    wintypes.DWORD, wintypes.DWORD, wintypes.DWORD)
_user32.SetWinEventHook.restype = wintypes.HANDLE
_user32.UnhookWinEvent.argtypes = (wintypes.HANDLE,)
_user32.GetClassNameW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
_user32.ShowWindow.argtypes = (wintypes.HWND, ctypes.c_int)
_user32.SetForegroundWindow.argtypes = (wintypes.HWND,)
_user32.PostThreadMessageW.argtypes = (wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
_EVENT_OBJECT_SHOW = 0x8002
_PPT_FRAME_CLASS = "PPTFrameClass"  # PowerPoint 문서 창(최상위 창) 클래스
_SW_HIDE, _SW_SHOW, _SW_SHOWNA = 0, 5, 8
_WM_QUIT = 0x0012


class _FrameHider:
    """with 블록 동안 화면에 나타나는 PowerPoint 문서 창을 그 즉시 숨긴다 (hidden 에 창 핸들을 모음).
    [원본 서식 유지] 붙여넣기에 필요한 문서 창(NewWindow)·결과를 닫을 때 잠깐 뜨는 빈 PowerPoint 창이 깜빡이지 않게 쓴다.
    - WinEvent 훅 콜백은 메시지 루프를 도는 스레드에서만 불린다 — 작업 스레드는 COM 호출 동안 전달받지 못하므로(실측)
      전용 스레드에서 훅을 건다.
    - 다른 프로세스 창의 ShowWindow 는 그 창의 스레드가 처리해야 끝나는데, 창을 처음 띄우는 PowerPoint 는 바빠서 약 0.25초
      늦게 숨겨진다(실측). PowerPoint 는 문서 창을 미리 숨겨 만들어 두고 다시 쓰므로, 블록에 들어갈 때(PowerPoint 가 한가할
      때) 숨어 있는 문서 창을 완전히 투명하게 해 두면 그사이에도 보이지 않는다. 나갈 때 원래 스타일로 되돌린다.
    - 블록 동안 사용자가 우연히 연 PowerPoint 창도 숨겨지므로, 나갈 때 keep(창 제목)이 False 인 창은 다시 보이게 한다
      (keep 을 주지 않으면 숨긴 창을 모두 그대로 둠 — 종료 직전처럼 사용자 창이 있을 수 없을 때).
    블록은 짧게 유지할 것."""

    def __init__(self, keep: Callable[[str], bool] | None = None):
        self.hidden: list[int] = []
        self._keep = keep
        self._transparent: dict[int, int] = {}  # 투명하게 바꾼 창 → 원래 확장 스타일
        self._proc = _WINEVENTPROC(self._on_show)  # 콜백 개체가 사라지지 않게 보관
        self._ready = threading.Event()
        self._thread_id = 0
        self._thread: threading.Thread | None = None

    def _on_show(self, hook, event, hwnd, obj_id, child_id, thread, time_ms):
        if obj_id != 0 or not hwnd:  # OBJID_WINDOW 만 (창 안의 컨트롤·캐럿 등 제외)
            return
        buf = ctypes.create_unicode_buffer(64)
        if _user32.GetClassNameW(hwnd, buf, 64) and buf.value == _PPT_FRAME_CLASS:
            _user32.ShowWindow(hwnd, _SW_HIDE)
            if hwnd not in self.hidden:
                self.hidden.append(hwnd)

    def _run(self):
        self._thread_id = ctypes.windll.kernel32.GetCurrentThreadId()
        hook = _user32.SetWinEventHook(_EVENT_OBJECT_SHOW, _EVENT_OBJECT_SHOW, None, self._proc, 0, 0, 0)
        self._ready.set()
        if not hook:
            return
        msg = wintypes.MSG()
        try:
            while _user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                _user32.TranslateMessage(ctypes.byref(msg))
                _user32.DispatchMessageW(ctypes.byref(msg))
        finally:
            _user32.UnhookWinEvent(hook)

    def _make_hidden_frames_transparent(self) -> None:
        frames = []
        try:
            win32gui.EnumWindows(lambda h, _: frames.append(h) if win32gui.GetClassName(h) == _PPT_FRAME_CLASS
                                 and not win32gui.IsWindowVisible(h) else None, None)
        except pywintypes.error:
            return
        for hwnd in frames:
            try:
                ex = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                if ex & win32con.WS_EX_LAYERED:
                    continue  # PowerPoint 가 스스로 쓰는 반투명 창일 수 있음 → 건드리지 않음
                win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex | win32con.WS_EX_LAYERED)
                self._transparent[hwnd] = ex
                win32gui.SetLayeredWindowAttributes(hwnd, 0, 0, win32con.LWA_ALPHA)
            except pywintypes.error:
                continue

    def __enter__(self):
        self._thread = threading.Thread(target=self._run, name="FrameHider", daemon=True)
        self._thread.start()
        self._ready.wait(2)
        self._make_hidden_frames_transparent()
        return self

    def __exit__(self, *exc):
        # 숨김 이벤트가 큐에 남아 있을 수 있으므로 메시지 루프가 그것까지 처리한 뒤 끝나도록 WM_QUIT 를 뒤에 넣음
        if self._thread_id:
            _user32.PostThreadMessageW(self._thread_id, _WM_QUIT, 0, 0)
        if self._thread is not None:
            self._thread.join(2)
        for hwnd, ex in self._transparent.items():  # 숨긴 뒤에 되돌리므로 보이지 않음. 이미 닫힌 창은 무시
            try:
                win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex)
            except pywintypes.error:
                pass
        if self._keep is not None:  # 이 작업과 관계없는 창(사용자가 그사이 연 문서 등)은 되살림
            for hwnd in list(self.hidden):
                try:
                    if not win32gui.IsWindow(hwnd):
                        self.hidden.remove(hwnd)
                    elif not self._keep(win32gui.GetWindowText(hwnd)):
                        _user32.ShowWindow(hwnd, _SW_SHOWNA)
                        self.hidden.remove(hwnd)
                except Exception:
                    continue
        return False


@dataclass
class MergeJob:
    source: Path
    title: str                                   # 로그/오류 표시용 (예: "찬송 29장")
    fill_values: list[str] = field(default_factory=list)  # {{내용1}}~{{내용3}} 에 넣을 문구 (빈 값은 건너뜀)
    warn_missing: bool = True                    # 입력 위치가 없을 때 경고 (찬송가 PPT 는 토큰이 없는 게 보통)
    replacements: dict[str, str] = field(default_factory=dict)


class PptEngineError(RuntimeError):
    pass


class PptEngine:
    def __init__(self, log: LogFn | None = None, text_placeholder: str = "{{내용}}",
                 text_shape_name: str = "내용", merge_mode: str = MERGE_SOURCE_FORMATTING,
                 fill_defaults: list[str] | None = None, slide_size: str = "original"):
        self.log = log or (lambda m: None)
        self.text_placeholder = text_placeholder
        self.text_shape_name = text_shape_name
        self.fill_defaults = list(fill_defaults or [])  # n 번 값이 비었을 때 넣을 문구 (없거나 비면 토큰을 지움)
        self.slide_size = slide_size
        self.merge_mode = merge_mode
        self.app = None
        self._was_running = False
        self._window = None          # [원본 서식 유지] 붙여넣기용 문서 창 (숨겨 둠)
        self._window_hwnds: list[int] = []
        self._anchor = None          # hold() 가 연 빈 프레젠테이션
        self._prev_alerts = None
        self._prev_security = None
        self._design_cache: dict[tuple[str, str], object] = {}
        self._tmp: Path | None = None

    # ------------------------------------------------------------------ 수명 관리
    def __enter__(self):
        pythoncom.CoInitialize()
        try:
            self.app = win32com.client.GetActiveObject("PowerPoint.Application")
            self._was_running = True
        except pywintypes.com_error:
            try:
                self.app = win32com.client.Dispatch("PowerPoint.Application")
            except pywintypes.com_error as e:
                pythoncom.CoUninitialize()
                raise PptEngineError("PowerPoint 를 실행할 수 없습니다. Microsoft PowerPoint 설치를 확인해 주세요.") from e
            self._was_running = False
        try:
            self._prev_alerts = self.app.DisplayAlerts
            self.app.DisplayAlerts = PP_ALERTS_NONE
        except pywintypes.com_error:
            pass
        # COM 으로 연 파일은 보호된 보기 없이 열리고 기본값(LowByDefault)이면 매크로도 허용된다.
        # 인터넷에서 받은 .pptm/.ppt 의 매크로·ActiveX 코드가 실행되지 않도록 작업 동안 강제로 막는다.
        try:
            self._prev_security = self.app.AutomationSecurity
            self.app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except (pywintypes.com_error, AttributeError):
            self.log("⚠ PowerPoint 매크로 차단 설정을 적용하지 못했습니다.")
        self._tmp = Path(tempfile.mkdtemp(prefix="ChurchPPT_bg_"))
        return self

    def __exit__(self, *exc):
        # 사용자가 쓰던 PowerPoint 라면 알림/매크로 보안 설정을 원래대로 (전역 설정이라 그대로 두면 영향이 남음)
        for attr, prev in (("AutomationSecurity", self._prev_security), ("DisplayAlerts", self._prev_alerts)):
            try:
                if prev is not None and self.app is not None:
                    setattr(self.app, attr, prev)
            except Exception:
                pass
        self.release()
        try:
            if self.app is not None and not self._was_running and self.app.Presentations.Count == 0:
                _empty_clipboard()  # 붙여넣기로 남은 슬라이드가 클립보드에 있으면 종료가 수 초 늦어진다 (실측 5.5초)
                with _FrameHider():  # 종료 직전에 잠깐 뜨는 빈 PowerPoint 창 숨김
                    self.app.Quit()
        except Exception:
            pass
        self.app = None
        self._window = None
        if self._tmp:
            shutil.rmtree(self._tmp, ignore_errors=True)
        pythoncom.CoUninitialize()
        return False

    def hold(self) -> None:
        """Bible2PPT 를 실행하는 동안 PowerPoint 를 켜 둔다. Bible2PPT 는 실행 중인 PowerPoint 를 같이 쓰고, 열린
        프레젠테이션이 있으면 끝날 때 종료하지 않는다 → 말씀마다 PowerPoint 를 새로 켜고 끄는 시간을 아낀다."""
        if self._anchor is None:
            try:
                self._anchor = self.app.Presentations.Add(MSO_FALSE)  # WithWindow=False
            except pywintypes.com_error:
                self._anchor = None

    def release(self) -> None:
        if self._anchor is not None:
            try:
                self._anchor.Close()
            except Exception:
                pass
            self._anchor = None

    def _already_open(self, path: Path, by_name: bool = True) -> bool:
        """사용자가 PowerPoint 에서 열어 둔 파일인지 (템플릿을 열어 고치는 중에 생성하는 경우 등).
        OneDrive 폴더의 파일은 FullName 이 'https://d.docs.live.net/…/이름.pptx' 라서(실측) 경로로는 비교할 수 없다
        → 주소로 열린 문서는 파일 이름이 같으면 열린 것으로 본다 (다른 파일이어도 복사본을 쓸 뿐이라 해가 없음)."""
        want = os.path.normcase(str(Path(path).resolve()))
        name = Path(path).name.casefold()
        try:
            pres = self.app.Presentations
            for i in range(1, pres.Count + 1):
                try:
                    full = pres(i).FullName
                    if "://" in full:
                        if by_name and pres(i).Name.casefold() == name:
                            return True
                    elif os.path.normcase(str(Path(full).resolve())) == want:
                        return True
                except (pywintypes.com_error, OSError, ValueError):
                    continue
        except pywintypes.com_error:
            pass
        return False

    def _open(self, path: Path):
        path = Path(path)
        # 암호가 걸린 PPT 는 PowerPoint 가 암호 창을 띄우고 답을 기다려 생성이 끝나지 않는다(실측).
        # 암호화된 .pptx/.pptm 은 ZIP 이 아니라 OLE 파일로 저장되므로 열기 전에 알아본다 (.ppt 는 원래 OLE 라 구별 못 함)
        if path.suffix.lower() in (".pptx", ".pptm"):
            try:
                with open(path, "rb") as f:
                    head = f.read(8)
            except OSError:
                head = b""
            if head == _OLE_SIGNATURE:
                raise PptEngineError(f"암호가 걸린 PPT 는 사용할 수 없습니다. PowerPoint 에서 암호를 해제한 뒤 다시 시도해 주세요:\n{path}")
        # 이미 열린 파일을 Presentations.Open 하면 새로 열지 않고 사용자의 그 문서를 돌려준다(실측) → 그대로 쓰면
        # 첫 원본은 SaveAs 로 결과 파일 이름이 되고, 나머지는 Close 로 닫혀 저장하지 않은 수정이 사라졌다.
        # 저장된 파일을 복사해 연다 (사용자 문서는 건드리지 않음).
        if self._already_open(path):
            copy = self._tmp / f"open_{len(list(self._tmp.glob('open_*')))}{path.suffix}"
            try:
                shutil.copy2(path, copy)
            except OSError as e:
                raise PptEngineError(f"PowerPoint 에서 열려 있는 파일을 읽을 수 없습니다. 닫고 다시 시도해 주세요:\n{path}\n{e}") from e
            self.log(f"  ⚠ {path.name}: PowerPoint 에서 열려 있어 저장된 내용으로 만듭니다 (저장하지 않은 수정은 빠짐)")
            path = copy
        try:
            # ReadOnly, Untitled, WithWindow
            return self.app.Presentations.Open(str(path), MSO_TRUE, MSO_FALSE, MSO_FALSE)
        except pywintypes.com_error as e:
            raise PptEngineError(f"PPT 파일을 열 수 없습니다:\n{path}\n{_com_msg(e)}") from e

    # ------------------------------------------------------------------ 병합
    def merge(self, jobs: list[MergeJob], output: Path, global_replacements: dict[str, str] | None = None,
              progress: Callable[[int, int, str], None] | None = None, open_mode: str = "") -> bool:
        """jobs 를 이어 붙여 output 에 저장한다. open_mode 가 "open"/"show" 면 저장한 결과를 이 PowerPoint 에서 바로 보여 준다
        (PowerPoint 를 끄고 다시 켜서 여는 것보다 훨씬 빠름). 보여 줬으면 True — 아니면 False 이고 호출한 쪽이 파일로 연다."""
        if not jobs:
            raise PptEngineError("병합할 PPT 가 없습니다.")
        progress = progress or (lambda *a: None)
        output = Path(output).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        if self._already_open(output, by_name=False):  # 다 만든 뒤 마지막 저장에서야 실패하지 않도록 미리 확인
            raise PptEngineError(f"결과 파일이 PowerPoint 에서 열려 있습니다. 닫고 다시 시도해 주세요:\n{output}")
        # 작업 파일에 만들고 마지막에 output 으로 SaveAs 한다. OneDrive 폴더에 처음부터 SaveAs 하면 PowerPoint 가 그 뒤로
        # 클라우드 문서로 다뤄 마지막 Save 가 내 PC 의 파일에 들어가지 않았다 (실측: 60초 뒤에도 첫 PPT 슬라이드만 남음)
        work = self._tmp / "result" / output.name
        work.parent.mkdir(exist_ok=True)
        self._design_cache.clear()

        dest = None
        try:
            # 1) 베이스: 첫 원본을 그대로 새 파일로 저장 (서식 100% 유지)
            progress(0, len(jobs), jobs[0].title)
            first = self._open(jobs[0].source)
            try:
                first.SaveAs(str(work), PP_SAVE_AS_PPTX)
            except pywintypes.com_error as e:
                first.Close()
                raise PptEngineError(f"작업 파일을 만들 수 없습니다:\n{work}\n{_com_msg(e)}") from e
            dest = first
            self._apply_slide_size(dest)
            ranges = [(1, dest.Slides.Count)]
            self.log(f"[1/{len(jobs)}] {jobs[0].title} — {dest.Slides.Count}장")

            # 2) 나머지 원본 이어 붙이기
            for n, job in enumerate(jobs[1:], start=2):
                progress(n - 1, len(jobs), job.title)
                start = dest.Slides.Count + 1
                count = self._append(dest, job.source)
                ranges.append((start, start + count - 1))
                self.log(f"[{n}/{len(jobs)}] {job.title} — {count}장")

            # 3) 항목별 텍스트 입력 / 치환 → 4) 공통 토큰 치환. 항목 슬라이드들이 전체를 빈틈없이 나누므로 공통 토큰도
            #    항목별로 같이 처리해 텍스트를 한 번만 읽는다.
            progress(len(jobs), len(jobs), "텍스트 입력")
            for job, (a, b) in zip(jobs, ranges):
                if b < a:
                    continue
                slides = [dest.Slides(i) for i in range(a, b + 1)]
                texts = self._collect_texts(slides)
                for k, v in job.replacements.items():
                    self._replace_in_texts(texts, k, v)
                self._fill_text(slides, texts, job)  # 값이 없어도 실행: 기본값을 넣거나 남은 토큰을 지움
                for k, v in (global_replacements or {}).items():
                    self._replace_in_texts(texts, k, v)

            progress(len(jobs), len(jobs), "저장 중")
            try:
                dest.SaveAs(str(output), PP_SAVE_AS_PPTX)
            except pywintypes.com_error as e:
                raise PptEngineError(f"결과 파일을 저장할 수 없습니다 (파일이 열려 있는지 확인):\n{output}\n{_com_msg(e)}") from e
            self.log(f"저장 완료: {output} (총 {dest.Slides.Count}장)")
            if open_mode and self._reveal(dest, open_mode == "show"):
                dest = None  # 사용자가 보는 창이 됐으므로 닫지 않음
                return True
            return False
        finally:
            if dest is not None:
                mine = Path(output).name
                try:
                    if self._window is not None:
                        mine = self._window.Caption
                except Exception:
                    pass
                try:
                    # 마지막 창이 닫힐 때 잠깐 뜨는 빈 PowerPoint 창(제목 'PowerPoint' 또는 빈 제목)은 숨긴 채로
                    with _FrameHider(keep=lambda t: t in ("", "PowerPoint") or _is_title_of(t, mine)):
                        dest.Close()
                except Exception:
                    pass
                self._window = None

    def _apply_slide_size(self, dest) -> None:
        """결과 비율을 지정했으면 베이스(첫 PPT 사본)의 슬라이드 크기를 바꾼다. 이어 붙이는 슬라이드는 이 크기에 맞춰진다.
        COM 으로 크기를 바꾸면 PowerPoint 의 [맞춤 확인]처럼 동작한다 (실측: 4:3→16:9 는 내용을 그대로 두고 가운데로,
        16:9→4:3 은 내용을 비율대로 줄여 위아래 여백) → 내용이 잘리지 않음."""
        size = SLIDE_SIZES.get(self.slide_size)
        if size is None:
            return
        ps = dest.PageSetup
        if abs(ps.SlideWidth - size[0]) <= 1 and abs(ps.SlideHeight - size[1]) <= 1:
            return
        try:
            ps.SlideWidth, ps.SlideHeight = size
            self.log(f"슬라이드 비율을 {self.slide_size} 로 맞춤")
        except pywintypes.com_error as e:
            self.log(f"⚠ 슬라이드 비율을 {self.slide_size} 로 바꾸지 못했습니다: {_com_msg(e)}")

    def _append(self, dest, source: Path) -> int:
        src = self._open(source)
        try:
            n = src.Slides.Count
            if n == 0:
                return 0
            if (abs(src.PageSetup.SlideWidth - dest.PageSetup.SlideWidth) > 1
                    or abs(src.PageSetup.SlideHeight - dest.PageSetup.SlideHeight) > 1):
                if self.slide_size in SLIDE_SIZES:  # 비율을 직접 골랐으면 맞추는 게 의도한 동작 → 경고 아님
                    self.log(f"  {source.name}: 슬라이드 크기를 {self.slide_size} 에 맞춤")
                else:
                    self.log(f"  ⚠ 슬라이드 크기가 다릅니다: {source.name} (첫 PPT 크기에 맞춰짐)")
            if self.merge_mode == MERGE_SOURCE_FORMATTING:
                try:
                    self._paste_source_formatting(src, dest)
                    return n
                except PptEngineError as e:
                    self.log(f"  ⚠ 원본 서식 유지 붙여넣기 실패 → 창 없는 방식으로 대체 ({e})")
            self._paste_windowless(src, dest, source)
            return n
        finally:
            src.Close()

    # ---- 방식 A: PowerPoint [원본 서식 유지] 붙여넣기
    def _ensure_window(self, dest):
        """붙여넣기용 문서 창. 나타나는 순간 숨기므로 화면에 뜨지 않고, 사용자의 다른 PowerPoint 창도 건드리지 않는다
        (예전에는 PowerPoint 를 최소화했는데 붙여넣을 때마다 Activate 가 창을 되살려 생성 내내 창이 보였음)."""
        if self._window is not None:
            return self._window
        try:
            caption = []
            with _FrameHider(keep=lambda t: bool(caption) and _is_title_of(t, caption[0])) as hider:
                win = dest.NewWindow()
                caption.append(win.Caption)
                win.ViewType = PP_VIEW_NORMAL
                win.Activate()
            self._window = win
            self._window_hwnds = hider.hidden
            return win
        except pywintypes.com_error as e:
            raise PptEngineError(f"작업 창을 열 수 없음: {_com_msg(e)}") from e

    def _activate(self, win) -> None:
        """ExecuteMso 는 PowerPoint 의 활성 창에 붙여넣는다. 숨긴 창도 활성 상태는 유지되므로 이미 활성이면 그대로 두고
        (Activate 는 숨긴 창을 다시 보이게 함), 다른 창이 활성이 됐을 때만 숨긴 채로 되돌린다."""
        try:
            if self.app.ActiveWindow.Caption == win.Caption:
                return
        except pywintypes.com_error:
            pass
        caption = win.Caption
        with _FrameHider(keep=lambda t: _is_title_of(t, caption)):
            win.Activate()

    def _reveal(self, dest, show: bool) -> bool:
        """숨겨 둔 작업 창을 결과 창으로 보여 준다 (파일을 다시 열 필요 없이 바로 표시). show 면 슬라이드 쇼 시작."""
        win = self._window
        if win is None or not self._window_hwnds:
            return False  # 창 없는 방식으로만 병합함 → 호출한 쪽이 파일로 연다
        try:
            win.View.GotoSlide(1)
            try:  # PowerPoint 창이 앞으로 나올 수 있게 — 그 프로세스에만 허용 (ASFW_ANY 는 아무 프로세스나 허용)
                pid = win32process.GetWindowThreadProcessId(self._window_hwnds[-1])[1]
                ctypes.windll.user32.AllowSetForegroundWindow(pid)
            except pywintypes.error:
                pass
            for hwnd in self._window_hwnds:  # 숨겨 둔 작업 창
                _user32.ShowWindow(hwnd, _SW_SHOW)
            win.WindowState = PP_WINDOW_MAXIMIZED
            win.Activate()
            _user32.SetForegroundWindow(self._window_hwnds[-1])
            if show:
                dest.SlideShowSettings.Run()
            self._window = None
            return True
        except pywintypes.com_error as e:
            self.log(f"⚠ 결과를 바로 열지 못했습니다: {_com_msg(e)}")
            return False

    def _paste_source_formatting(self, src, dest, retries: int = 3) -> None:
        n = src.Slides.Count
        last = None
        for attempt in range(retries):
            before = dest.Slides.Count
            try:
                win = self._ensure_window(dest)
                src.Slides.Range().Copy()
                time.sleep(0.05)
                self._activate(win)
                win.View.GotoSlide(before)
                try:
                    win.Panes(1).Activate()  # 썸네일 창에 포커스 → 슬라이드 단위로 붙여넣기
                except pywintypes.com_error:
                    pass
                self.app.CommandBars.ExecuteMso("PasteSourceFormatting")
            except pywintypes.com_error as e:
                last = _com_msg(e)
                time.sleep(0.4 * (attempt + 1))
                continue
            # ExecuteMso 는 비동기일 수 있으므로 슬라이드 수가 늘 때까지 대기
            deadline = time.time() + 10
            while time.time() < deadline and dest.Slides.Count < before + n:
                time.sleep(0.1)
                pythoncom.PumpWaitingMessages()
            added = dest.Slides.Count - before
            if added == n:
                return
            for _ in range(max(added, 0)):  # 일부만 붙은 경우 되돌림
                dest.Slides(before + 1).Delete()
            last = f"붙여넣은 슬라이드 수 불일치 ({added}/{n})"
            time.sleep(0.4 * (attempt + 1))
        raise PptEngineError(last or "알 수 없는 오류")

    # ---- 방식 B: 창 없이 붙여넣기 + 디자인/레이아웃/배경 복원
    def _paste_windowless(self, src, dest, source: Path) -> None:
        n = src.Slides.Count
        backgrounds = [self._capture_background(src.Slides(i), i) for i in range(1, n + 1)]
        start = dest.Slides.Count + 1
        try:
            pasted = self._copy_paste(src, dest)
            new_slides = [pasted(i) for i in range(1, pasted.Count + 1)]
        except PptEngineError as e:
            self.log(f"  ⚠ 복사/붙여넣기 실패 → 파일 삽입 방식으로 대체 ({e})")
            dest.Slides.InsertFromFile(str(source), dest.Slides.Count, 1, n)
            new_slides = [dest.Slides(start + i) for i in range(n)]
        for i, new in enumerate(new_slides):
            self._apply_design(new, src.Slides(i + 1), source)
            self._apply_background(new, backgrounds[i])

    def _copy_paste(self, src, dest, retries: int = 6):
        last = None
        for attempt in range(retries):
            try:
                src.Slides.Range().Copy()
                time.sleep(0.15)
                pasted = dest.Slides.Paste(dest.Slides.Count + 1)
                if pasted.Count == src.Slides.Count:
                    return pasted
                pasted.Delete()
                last = "붙여넣은 슬라이드 수 불일치"
            except pywintypes.com_error as e:
                last = _com_msg(e)
            time.sleep(0.3 * (attempt + 1))
        raise PptEngineError(f"클립보드 오류: {last}")

    def _apply_design(self, new_slide, src_slide, source: Path) -> None:
        try:
            src_design = src_slide.Design
            key = (str(source).lower(), src_design.Name)
            cached = self._design_cache.get(key)
            if cached is not None:
                new_slide.Design = cached
            else:
                new_slide.Design = src_design          # 다른 프레젠테이션의 디자인 → 가져오기
                self._design_cache[key] = new_slide.Design
        except pywintypes.com_error as e:
            self.log(f"  ⚠ 디자인 적용 실패: {_com_msg(e)}")
            return
        try:
            layout_name = src_slide.CustomLayout.Name
            layouts = new_slide.Design.SlideMaster.CustomLayouts
            for i in range(1, layouts.Count + 1):
                if layouts(i).Name == layout_name:
                    new_slide.CustomLayout = layouts(i)
                    break
        except pywintypes.com_error:
            pass

    def _capture_background(self, slide, idx: int):
        """슬라이드 자체 배경을 기록. 반환: None(마스터 배경) | ("solid", rgb, transparency) | ("image", path)"""
        try:
            if slide.FollowMasterBackground:
                return None
            fill = slide.Background.Fill
            if fill.Type == MSO_FILL_SOLID:
                return ("solid", fill.ForeColor.RGB, fill.Transparency)
            # 그라데이션/그림/질감 등: 도형을 잠시 숨기고 배경만 고해상도 이미지로 내보냄 (원본은 저장하지 않음)
            hidden = []
            for i in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes(i)
                if shp.Visible:
                    shp.Visible = MSO_FALSE
                    hidden.append(shp)
            master_shapes = slide.DisplayMasterShapes
            slide.DisplayMasterShapes = MSO_FALSE
            ps = slide.Parent.PageSetup
            scale = 2560 / max(ps.SlideWidth, 1)
            path = self._tmp / f"bg_{id(slide)}_{idx}.png"
            slide.Export(str(path), "PNG", int(ps.SlideWidth * scale), int(ps.SlideHeight * scale))
            slide.DisplayMasterShapes = master_shapes
            for shp in hidden:
                shp.Visible = MSO_TRUE
            return ("image", path)
        except pywintypes.com_error as e:
            self.log(f"  ⚠ 배경 읽기 실패: {_com_msg(e)}")
            return None

    def _apply_background(self, new_slide, bg) -> None:
        try:
            if bg is None:
                new_slide.FollowMasterBackground = MSO_TRUE
                return
            new_slide.FollowMasterBackground = MSO_FALSE
            fill = new_slide.Background.Fill
            if bg[0] == "solid":
                fill.Solid()
                fill.ForeColor.RGB = bg[1]
                fill.Transparency = bg[2]
            else:
                fill.UserPicture(str(bg[1]))
        except pywintypes.com_error as e:
            self.log(f"  ⚠ 배경 적용 실패: {_com_msg(e)}")

    # ------------------------------------------------------------------ 텍스트
    def _shape_text_ranges(self, shape):
        try:
            if shape.Type == MSO_GROUP:
                items = shape.GroupItems
                for i in range(1, items.Count + 1):
                    yield from self._shape_text_ranges(items(i))
                return
            if shape.HasTable:
                table = shape.Table
                for r in range(1, table.Rows.Count + 1):
                    for c in range(1, table.Columns.Count + 1):
                        tf = table.Cell(r, c).Shape.TextFrame
                        if tf.HasText:
                            yield tf.TextRange
                return
            if shape.HasTextFrame and shape.TextFrame.HasText:
                yield shape.TextFrame.TextRange
        except pywintypes.com_error:
            return

    def _collect_texts(self, slides) -> list[list]:
        """슬라이드의 모든 텍스트 범위와 그 글자 [[TextRange, text], ...]. COM 호출은 한 번에 수 ms 라서, 예전처럼 토큰마다
        모든 도형을 다시 훑고 Replace 를 시도하면 37장 기준 약 10초가 걸렸다 → 글자를 한 번 읽어 두고 토큰이 든 곳만 Replace."""
        texts = []
        for slide in slides:
            shapes = slide.Shapes
            for i in range(1, shapes.Count + 1):
                for tr in self._shape_text_ranges(shapes(i)):
                    try:
                        texts.append([tr, tr.Text])
                    except pywintypes.com_error:
                        continue
        return texts

    @staticmethod
    def _replace_in_texts(texts: list[list], find: str, repl: str) -> int:
        if not find:
            return 0
        key = find.lower()  # TextRange.Replace 는 기본이 대소문자 무시 → 같은 기준으로 후보를 고름
        total = 0
        for entry in texts:
            if key not in entry[1].lower():
                continue
            n = _replace_all(entry[0], find, repl)
            if n:
                total += n
                try:
                    entry[1] = entry[0].Text
                except pywintypes.com_error:
                    pass
        return total

    def _fill_text(self, slides, texts: list[list], job: MergeJob) -> None:
        """값 n 번째 → {{내용n}} 치환, 없으면 이름이 '내용n' 인 텍스트 상자에 입력.
        1번은 예전 템플릿 호환을 위해 {{내용}} / '내용' 도 받는다.
        값이 비었으면 설정의 n 번 기본값을 넣고, 기본값도 비었으면 토큰 글자를 지운다 (결과에 '{{내용2}}' 가 남지 않도록).
        이름이 '내용n' 인 텍스트 상자는 값이 없으면 그대로 둔다 (상자 글자는 템플릿이 정한 문구라서)."""
        missing = []
        for n in range(1, MAX_FILL + 1):
            given = job.fill_values[n - 1] if n <= len(job.fill_values) else ""
            value = given or (self.fill_defaults[n - 1] if n <= len(self.fill_defaults) else "")
            tokens = [numbered_token(self.text_placeholder, n)]
            names = [f"{self.text_shape_name}{n}"] if self.text_shape_name else []
            if n == 1:
                tokens.append(self.text_placeholder)
                names.append(self.text_shape_name)
            tokens = [t for t in tokens if t]
            if not value:
                for t in tokens:
                    self._replace_in_texts(texts, t, "")
                continue
            if sum(self._replace_in_texts(texts, t, value) for t in tokens):
                continue
            if self._fill_shape(slides, [nm for nm in names if nm], value):
                texts[:] = self._collect_texts(slides)  # 상자 글자가 바뀌었으므로 다시 읽음 (공통 토큰 치환용)
            elif given:  # 기본값은 자리가 없어도 경고 안 함
                missing.append(tokens[0] if tokens else (names[0] if names else str(n)))
        if missing and job.warn_missing:
            self.log(f"  ⚠ '{job.title}': 텍스트 입력 위치를 찾지 못했습니다 ({', '.join(missing)}). "
                     f"템플릿에 해당 토큰 또는 '{self.text_shape_name}1'~'{self.text_shape_name}{MAX_FILL}' "
                     "이름의 텍스트 상자를 넣어 주세요.")

    @staticmethod
    def _fill_shape(slides, names: list[str], value: str) -> bool:
        for slide in slides:
            for i in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes(i)
                try:
                    if shp.Name in names and shp.HasTextFrame:
                        shp.TextFrame.TextRange.Text = value
                        return True
                except pywintypes.com_error:
                    continue
        return False


def _replace_all(text_range, find: str, repl: str, limit: int = 200) -> int:
    """TextRange.Replace 는 한 번에 하나만 바꾸므로 반복 (글자 서식 유지)."""
    count = 0
    after = 0
    while count < limit:
        try:
            r = text_range.Replace(find, repl, after)
        except pywintypes.com_error:
            break
        if r is None:
            break
        count += 1
        after = r.Start + r.Length - 1  # 치환 결과 뒤부터 다시 찾기 (repl 에 find 포함 시 무한루프 방지)
    return count


def _is_title_of(title: str, caption: str) -> bool:
    """창 제목이 그 문서 창의 것인지. 제목은 '이름 - PowerPoint' 이고, 확장자는 탐색기 설정에 따라 빠지기도 한다."""
    if not caption:
        return False
    stem = caption.rsplit(".", 1)[0] if caption.lower().endswith((".pptx", ".pptm", ".ppt")) else caption
    return title.startswith(caption) or title.startswith(stem + " ") or title == stem


# ---- 생성이 멈췄을 때 (App 의 생성 중 화면에서 사용): 처음 설치한 Office 의 라이선스 동의·로그인 창 등이 뜨면
#      PowerPoint 가 그 창의 응답을 기다리느라 COM 호출이 끝나지 않는다. COM 호출은 중간에 끊을 수 없으므로
#      대화 상자를 앞으로 보여 주거나, 이 프로그램이 켠 PowerPoint 를 강제로 끄는(호출이 오류로 끝남) 수밖에 없다.
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
_kernel32.OpenProcess.restype = wintypes.HANDLE
_kernel32.QueryFullProcessImageNameW.argtypes = (wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
                                                 ctypes.POINTER(wintypes.DWORD))
_kernel32.TerminateProcess.argtypes = (wintypes.HANDLE, wintypes.UINT)
_kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
_PROCESS_TERMINATE, _PROCESS_QUERY_LIMITED_INFORMATION = 0x0001, 0x1000


def _image_name(pid: int) -> str:
    h = _kernel32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return ""
    try:
        buf, size = ctypes.create_unicode_buffer(1024), wintypes.DWORD(1024)
        return buf.value if _kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)) else ""
    finally:
        _kernel32.CloseHandle(h)


def powerpoint_pids() -> set[int]:
    """실행 중인 POWERPNT.EXE 프로세스 ID (PowerPoint 는 하나만 실행되므로 보통 0~1개)."""
    try:
        pids = win32process.EnumProcesses()
    except pywintypes.error:
        return set()
    return {pid for pid in pids if pid and _image_name(pid).lower().endswith("\\powerpnt.exe")}


def show_powerpoint_dialogs(pids: set[int]) -> int:
    """그 PowerPoint 프로세스의 화면에 보이는 창(대화 상자) 중 문서 창이 아닌 것을 앞으로 가져온다 → 개수.
    문서 창(PPTFrameClass)은 작업 중 숨겨 둔 창이라 건드리지 않는다."""
    found = []

    def visit(hwnd, _):
        try:
            if (win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) != _PPT_FRAME_CLASS
                    and win32process.GetWindowThreadProcessId(hwnd)[1] in pids):
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                if right - left > 20 and bottom - top > 20:  # 크기 없는 보조 창 제외
                    found.append(hwnd)
        except pywintypes.error:
            pass
        return True

    try:
        win32gui.EnumWindows(visit, None)
    except pywintypes.error:
        return 0
    for pid in pids:
        try:
            ctypes.windll.user32.AllowSetForegroundWindow(pid)
        except OSError:
            pass
    for hwnd in found:
        _user32.ShowWindow(hwnd, _SW_SHOW)
        _user32.SetForegroundWindow(hwnd)
    return len(found)


def kill_powerpoint(pids: set[int]) -> int:
    """PowerPoint 강제 종료 → 끈 개수. 프로세스 ID 가 다른 프로그램에 다시 쓰였을 수 있으므로 지금도 PowerPoint 인 것만."""
    killed = 0
    for pid in pids & powerpoint_pids():
        h = _kernel32.OpenProcess(_PROCESS_TERMINATE, False, pid)
        if h:
            try:
                killed += bool(_kernel32.TerminateProcess(h, 1))
            finally:
                _kernel32.CloseHandle(h)
    return killed


def _empty_clipboard() -> None:
    try:
        if _user32.OpenClipboard(None):
            try:
                _user32.EmptyClipboard()
            finally:
                _user32.CloseClipboard()
    except OSError:
        pass


def _com_msg(e: pywintypes.com_error) -> str:
    try:
        exc = e.excepinfo
        if exc and exc[2]:
            return str(exc[2]).strip()
    except Exception:
        pass
    return str(e)
