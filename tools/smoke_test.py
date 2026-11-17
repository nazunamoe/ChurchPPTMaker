"""실행 시험 — PowerPoint·Bible2PPT 없이, 이 PC 의 CPU 에서 프로그램이 뜨는지 확인한다.
ARM64 PC 가 없어 GitHub Actions 의 ARM64 실행기(windows-11-arm)에서 돌리려고 만듦 (.github/workflows/build.yml).

    python tools/smoke_test.py [--expect win-arm64]          소스로: 모듈·Win32 함수·ctypes 콜백·메인 창
    python tools/smoke_test.py --exe dist/ChurchPPTMaker-arm64.exe [--expect win-arm64]   빌드한 exe 가 창을 띄우는지

CPU 마다 달라질 수 있는 부분만 본다 (PPT 병합은 PowerPoint 가 있어야 해서 확인하지 못함):
- 이 Python 의 CPU(sysconfig 플랫폼), pywin32(COM 포함)·CustomTkinter 불러오기
- 모듈을 불러올 때 이름으로 찾는 Win32 함수(user32·kernel32·comctl32)
- ctypes 콜백 두 가지: 창 틀(SetWindowSubclass — 위쪽 제목 표시줄이 실제로 없어졌는지), WinEvent 훅(_FrameHider)
- 메인 창을 만들어 몇 초 돌린 뒤 콜백 오류([로그]의 '프로그램 오류')가 없는지
- exe: PE 헤더의 CPU, 창이 뜨고 닫기(WM_CLOSE)로 정상 종료하는지
설정 파일을 건드리지 않도록 소스·exe 를 임시 폴더에 복사해 거기서 실행한다 (프로그램 폴더 = 소스·exe 가 있는 폴더라
config.json·presets.json·templates 폴더 등이 생김). 처음 실행이면 설정 마법사가 떠서 멈추므로 마친 것으로 해 둔다.
"""
from __future__ import annotations

import argparse
import ctypes
import json
import shutil
import subprocess
import sys
import sysconfig
import tempfile
import time
from ctypes import wintypes
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RUN_SECONDS = 4          # 메인 창을 돌려 보는 시간 (시작 크기 계산·표시·창 틀 처리가 끝나기에 충분)
EXE_START_TIMEOUT = 120  # exe 창이 뜰 때까지 (onefile 은 처음에 임시 폴더로 풀어서 느림)
EXE_EXIT_TIMEOUT = 30
PE_MACHINES = {0x8664: "win-amd64", 0xAA64: "win-arm64", 0x014C: "win32"}
WM_CLOSE = 0x0010


class SmokeError(Exception):
    pass


def check(ok: bool, msg: str) -> None:
    if not ok:
        raise SmokeError(msg)
    print(f"  ✓ {msg}")


def _isolated_copy(dest: Path) -> None:
    for p in ROOT.glob("*.py"):
        shutil.copy2(p, dest / p.name)
    shutil.copytree(ROOT / "assets", dest / "assets")
    (dest / "tools").mkdir()
    shutil.copy2(Path(__file__), dest / "tools" / Path(__file__).name)
    (dest / "config.json").write_text(json.dumps({"setup_done": True}), encoding="utf-8")


def _pump(app, seconds: float) -> None:
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        app.update()
        time.sleep(0.02)


def run_source(expect: str | None) -> None:
    """임시 폴더에 복사한 소스로 자식 프로세스에서 시험 (config_manager 가 불러올 때 프로그램 폴더를 정하므로)."""
    with tempfile.TemporaryDirectory(prefix="ChurchPPTMaker_smoke_", ignore_cleanup_errors=True) as tmp:
        _isolated_copy(Path(tmp))
        cmd = [sys.executable, str(Path(tmp) / "tools" / Path(__file__).name), "--child"]
        if expect:
            cmd += ["--expect", expect]
        r = subprocess.run(cmd, cwd=tmp, timeout=180)
        if r.returncode != 0:
            raise SmokeError(f"소스 시험 실패 (종료 코드 {r.returncode})")


def child(expect: str | None) -> None:
    sys.path.insert(0, str(ROOT))
    platform = sysconfig.get_platform()
    print(f"Python {sys.version.split()[0]} ({platform}) — {ROOT}")
    if expect:
        check(platform == expect, f"Python CPU = {expect}")

    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    d = win32com.client.Dispatch("Scripting.Dictionary")  # PowerPoint 와 같은 늦은 바인딩(IDispatch) 경로
    d.Add("찬송", 29)
    check(d.Count == 1 and d.Item("찬송") == 29, "pywin32 COM (Scripting.Dictionary)")
    del d
    pythoncom.CoUninitialize()

    import tkinter as tk
    import main          # 모든 모듈 + 불러올 때 지정하는 Win32 함수
    import ppt_engine
    check(True, "모듈 불러오기 (main·ppt_engine·window_frame·ui_widgets …)")
    check(isinstance(ppt_engine.powerpoint_pids(), set), "프로세스 목록 (OpenProcess·QueryFullProcessImageNameW)")

    # WinEvent 훅 콜백이 실제로 불리는지 세기 (불릴 때마다 원래 처리도 그대로)
    events: list[int] = []
    original = ppt_engine._FrameHider._on_show

    def counting(self, *args):
        events.append(args[2])
        return original(self, *args)

    ppt_engine._FrameHider._on_show = counting

    app = main.App()
    try:
        _pump(app, RUN_SECONDS)
        check(app._revealed, "메인 창 표시")
        check(app.frame.active, "창 틀 처리 설치 (SetWindowSubclass)")
        u32 = ctypes.WinDLL("user32")
        u32.GetWindowRect.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.RECT))
        u32.ClientToScreen.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.POINT))
        rect, pt = wintypes.RECT(), wintypes.POINT(0, 0)
        u32.GetWindowRect(app.frame.hwnd, ctypes.byref(rect))
        u32.ClientToScreen(app.frame.hwnd, ctypes.byref(pt))
        check(abs(pt.y - rect.top) <= 1,
              f"창 틀 콜백(WM_NCCALCSIZE)으로 제목 표시줄 없앰 (창 위 {rect.top}, 내용 위 {pt.y})")

        with ppt_engine._FrameHider(keep=lambda _title: False):
            top = tk.Toplevel(app)
            top.title("smoke")
            top.geometry("200x100")
            _pump(app, 1.0)
            top.destroy()
            _pump(app, 0.3)
        check(len(events) > 0, f"WinEvent 훅 콜백 (EVENT_OBJECT_SHOW {len(events)}회)")

        errors = [line for line in app.log_lines if "프로그램 오류" in line]
        for e in errors:
            print(e)
        check(not errors, "콜백 오류 없음")
    finally:
        ppt_engine._FrameHider._on_show = original
        app.destroy()


def _pe_platform(exe: Path) -> str:
    data = exe.read_bytes()[:4096]
    pe = int.from_bytes(data[0x3C:0x40], "little")
    return PE_MACHINES.get(int.from_bytes(data[pe + 4:pe + 6], "little"), "unknown")


def run_exe(exe: Path, expect: str | None) -> None:
    sys.path.insert(0, str(ROOT))
    from version import APP_NAME, __version__
    title = f"{APP_NAME} v{__version__}"
    exe = exe.resolve()
    check(exe.is_file(), f"exe 있음: {exe}")
    platform = _pe_platform(exe)
    print(f"exe CPU: {platform}")
    if expect:
        check(platform == expect, f"exe CPU = {expect}")

    u32 = ctypes.WinDLL("user32")
    u32.FindWindowW.restype = wintypes.HWND
    u32.FindWindowW.argtypes = (wintypes.LPCWSTR, wintypes.LPCWSTR)
    u32.PostMessageW.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
    # 이미 떠 있으면 새 exe 는 그 창을 앞으로 가져오고 끝나므로(한 번에 하나만 실행) 시험이 되지 않음
    check(not u32.FindWindowW(None, title), "같은 프로그램이 실행 중이 아님")

    with tempfile.TemporaryDirectory(prefix="ChurchPPTMaker_smoke_exe_", ignore_cleanup_errors=True) as tmp:
        run = Path(tmp) / exe.name
        shutil.copy2(exe, run)
        (Path(tmp) / "config.json").write_text(json.dumps({"setup_done": True}), encoding="utf-8")
        proc = subprocess.Popen([str(run)], cwd=tmp)
        try:
            start = time.monotonic()
            hwnd = None
            while time.monotonic() - start < EXE_START_TIMEOUT and proc.poll() is None:
                hwnd = u32.FindWindowW(None, title)
                if hwnd:
                    break
                time.sleep(0.25)
            check(bool(hwnd), f"exe 창 뜸 ({time.monotonic() - start:.1f}초, 종료 코드 {proc.poll()})")
            time.sleep(RUN_SECONDS)
            check(proc.poll() is None, "몇 초 뒤에도 실행 중")
            check((Path(tmp) / "presets.json").is_file(), "exe 폴더에 설정 파일 만듦 (프로그램 폴더 = exe 폴더)")
            u32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
            try:
                code = proc.wait(EXE_EXIT_TIMEOUT)
            except subprocess.TimeoutExpired:
                code = None
            check(code == 0, f"닫기로 정상 종료 (종료 코드 {code})")
        finally:
            if proc.poll() is None:  # onefile 은 프로세스 둘(풀기 + 실행) → 함께 끔
                subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], capture_output=True)
                proc.wait(10)
            time.sleep(1)  # onefile 이 풀어 둔 임시 폴더를 지우는 동안 기다림 (그 뒤 임시 폴더 정리)


def main_cli() -> int:
    try:  # GitHub Actions 등 콘솔이 아닌 곳(cp1252 등)에서도 한글·✓ 출력
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass
    ap = argparse.ArgumentParser(description="ChurchPPTMaker 실행 시험 (PowerPoint 없이)")
    ap.add_argument("--expect", choices=sorted(set(PE_MACHINES.values())), help="이 CPU 여야 함 (예: win-arm64)")
    ap.add_argument("--exe", type=Path, help="빌드한 exe 를 시험 (생략하면 소스)")
    ap.add_argument("--child", action="store_true", help=argparse.SUPPRESS)
    args = ap.parse_args()
    try:
        if args.child:
            child(args.expect)
        elif args.exe:
            run_exe(args.exe, args.expect)
            print("exe 시험 통과")
        else:
            run_source(args.expect)
            print("소스 시험 통과")
    except SmokeError as e:
        print(f"  ✗ {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main_cli())
