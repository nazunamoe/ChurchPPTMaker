"""Windows 기본 제목 표시줄을 없애고 프로그램 안에 제목 줄·창 버튼을 두기 위한 창 틀 처리.

overrideredirect 로 틀을 통째로 없애면 작업 표시줄 아이콘·Alt+Tab·가장자리 크기 조절·화면 끝 끌어 붙이기(스냅)·그림자가
모두 사라진다. 대신 창은 보통 창 그대로 두고 WM_NCCALCSIZE 에서 **위쪽 제목 표시줄만** 클라이언트 영역으로 돌린다
(Windows Terminal 과 같은 방식). 좌우·아래의 보이지 않는 크기 조절 테두리, Win11 둥근 모서리·테두리·그림자는 그대로.
- 창 옮기기: 제목 줄을 누르면 WM_NCLBUTTONDOWN(HTCAPTION) 을 보내 Windows 가 직접 옮긴다 (스냅·Win+화살표 그대로).
  ⚠ 반드시 PostMessage. Tk 이벤트 처리(=Python 콜백) 안에서 SendMessage 로 보내면 Windows 의 옮기기 루프가 그 안에서
  Tk 이벤트를 다시 돌려 Python 콜백이 겹쳐 불리고, tkinter 의 GIL 상태가 맞지 않아 프로그램이 강제 종료된다(실측:
  "PyEval_RestoreThread: ... GIL is released"). PostMessage 면 콜백이 끝난 뒤 Tk 메시지 루프에서 시작된다.
- 위쪽 가장자리 크기 조절: 위쪽은 클라이언트라 Windows 가 판정하지 못하므로 제목 줄 맨 위 몇 픽셀에서 HTTOP 을 보낸다.
- 최대화하면 Windows 는 창을 틀 두께만큼 화면 밖으로 키우므로 그만큼 위쪽을 안으로 들인다 (아래·좌우는 기본 처리).
창 프로시저는 comctl32 SetWindowSubclass 로 감싼다 (SetWindowLongPtr 보다 안전, 창이 없어질 때 WM_NCDESTROY 에서 해제).
"""
from __future__ import annotations

import ctypes
import tkinter as tk
from ctypes import wintypes

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_comctl = ctypes.WinDLL("comctl32")
_LRESULT = ctypes.c_ssize_t
_SUBCLASSPROC = ctypes.WINFUNCTYPE(_LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
                                   ctypes.c_size_t, ctypes.c_size_t)
_comctl.SetWindowSubclass.argtypes = (wintypes.HWND, _SUBCLASSPROC, ctypes.c_size_t, ctypes.c_size_t)
_comctl.SetWindowSubclass.restype = wintypes.BOOL
_comctl.RemoveWindowSubclass.argtypes = (wintypes.HWND, _SUBCLASSPROC, ctypes.c_size_t)
_comctl.DefSubclassProc.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
_comctl.DefSubclassProc.restype = _LRESULT
_user32.GetParent.argtypes = (wintypes.HWND,)
_user32.GetParent.restype = wintypes.HWND
_user32.IsZoomed.argtypes = (wintypes.HWND,)
try:  # Windows 10 1607 이상에만 있음 — 없으면 시스템 배율 값으로 대신 (불러오기 단계에서 꺼지지 않게)
    _user32.GetDpiForWindow.argtypes = (wintypes.HWND,)
    _user32.GetSystemMetricsForDpi.argtypes = (ctypes.c_int, wintypes.UINT)
    _HAS_DPI_API = True
except AttributeError:
    _HAS_DPI_API = False
_user32.SetWindowPos.argtypes = (wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                 wintypes.UINT)
_user32.PostMessageW.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

_WM_NCCALCSIZE, _WM_NCDESTROY, _WM_NCLBUTTONDOWN = 0x0083, 0x0082, 0x00A1
_HTCAPTION, _HTTOP, _HTTOPLEFT, _HTTOPRIGHT = 2, 12, 13, 14
_SM_CYFRAME, _SM_CXPADDEDBORDER = 33, 92
_SWP_FRAMECHANGED_ONLY = 0x0020 | 0x0002 | 0x0001 | 0x0004 | 0x0010  # FRAMECHANGED|NOMOVE|NOSIZE|NOZORDER|NOACTIVATE
_SUBCLASS_ID = 0xC4A3


def _window_dpi(hwnd) -> int:
    return _user32.GetDpiForWindow(hwnd) if _HAS_DPI_API else 96


def _frame_metric(index: int, dpi: int) -> int:
    return _user32.GetSystemMetricsForDpi(index, dpi) if _HAS_DPI_API else _user32.GetSystemMetrics(index)


class _NCCALCSIZE_PARAMS(ctypes.Structure):
    _fields_ = [("rgrc", wintypes.RECT * 3), ("lppos", ctypes.c_void_p)]


class CustomFrame:
    """win(Tk 최상위 창)의 Windows 제목 표시줄을 없앤다. install() 은 창이 화면에 매핑된 뒤 한 번 부른다."""

    def __init__(self, win: tk.Misc):
        self.win = win
        self.hwnd = 0
        self.active = False
        self._proc = _SUBCLASSPROC(self._wndproc)  # 콜백 개체가 사라지지 않게 보관

    def install(self) -> bool:
        if self.active:
            return True
        try:
            self.hwnd = _user32.GetParent(self.win.winfo_id())  # Tk 최상위 창의 실제 틀(wrapper) 창
            if not self.hwnd or not _comctl.SetWindowSubclass(self.hwnd, self._proc, _SUBCLASS_ID, 0):
                return False
            self.active = True
            _user32.SetWindowPos(self.hwnd, None, 0, 0, 0, 0, _SWP_FRAMECHANGED_ONLY)  # 틀 다시 계산
            return True
        except (OSError, AttributeError, tk.TclError):
            return False

    def _wndproc(self, hwnd, msg, wparam, lparam, uid, ref):
        try:
            if msg == _WM_NCCALCSIZE and wparam:
                params = ctypes.cast(lparam, ctypes.POINTER(_NCCALCSIZE_PARAMS)).contents
                top = params.rgrc[0].top
                result = _comctl.DefSubclassProc(hwnd, msg, wparam, lparam)  # 좌우·아래 틀은 기본대로
                rect = params.rgrc[0]
                if _user32.IsZoomed(hwnd):  # 최대화: 화면 밖으로 넘친 위쪽 틀 두께만큼 안으로
                    dpi = _window_dpi(hwnd)
                    rect.top = top + _frame_metric(_SM_CYFRAME, dpi) + _frame_metric(_SM_CXPADDEDBORDER, dpi)
                else:
                    rect.top = top  # 위쪽 제목 표시줄만 없앰
                params.rgrc[0] = rect
                return result
            if msg == _WM_NCDESTROY:
                _comctl.RemoveWindowSubclass(hwnd, self._proc, _SUBCLASS_ID)
                self.active = False
        except Exception:  # 창 프로시저에서 예외가 새면 안 됨 → 기본 처리로
            pass
        return _comctl.DefSubclassProc(hwnd, msg, wparam, lparam)

    # ---- 제목 줄 동작
    def is_maximized(self) -> bool:
        return bool(self.hwnd and _user32.IsZoomed(self.hwnd))

    def start_move(self, _event=None):
        """제목 줄을 누르면 Windows 기본 창 옮기기 (끌어서 화면 위쪽에 놓으면 최대화 등 스냅 포함)."""
        if self.active:
            _user32.ReleaseCapture()
            _user32.PostMessageW(self.hwnd, _WM_NCLBUTTONDOWN, _HTCAPTION, 0)

    def start_resize_top(self, x_root: int):
        """위쪽 가장자리 크기 조절 (모서리면 대각선)."""
        if not self.active or self.is_maximized():
            return
        left = self.win.winfo_rootx()
        right = left + self.win.winfo_width()
        corner = round(16 * _window_dpi(self.hwnd) / 96)  # 이 폭 안이면 모서리(대각선) 크기 조절
        hit = _HTTOPLEFT if x_root - left < corner else _HTTOPRIGHT if right - x_root < corner else _HTTOP
        _user32.ReleaseCapture()
        _user32.PostMessageW(self.hwnd, _WM_NCLBUTTONDOWN, hit, 0)
