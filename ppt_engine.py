"""win32com 기반 PowerPoint 병합/치환 엔진.

원칙
- 원본 PPT는 ReadOnly 로만 연다 (절대 저장하지 않음).
- 첫 원본을 SaveAs 로 베이스 생성 (원본 100% 유지).
- 나머지 원본은 merge_mode 에 따라 이어 붙인다.
  * "source_formatting" (기본): PowerPoint 의 [원본 서식 유지] 붙여넣기 명령(PasteSourceFormatting)을 실행.
    테마·마스터·레이아웃·슬라이드 배경(그림 포함)까지 PowerPoint 가 직접 보존한다.
    명령 실행에 문서 창이 필요하므로 PowerPoint 를 최소화한 상태로 창을 하나 연다.
  * "windowless": 창 없이 Slides.Paste 후 슬라이드별 Design/CustomLayout 을 원본 것으로 지정하고
    슬라이드 배경을 복제(단색은 정확히, 그라데이션/그림 등은 고해상도 이미지로)한다.
  "source_formatting" 이 실패하면 해당 파일은 자동으로 "windowless" 로 처리한다.
- 애니메이션·전환·폰트 서식은 슬라이드와 함께 복사된다.
- 반드시 작업 스레드에서 with 문으로 사용한다 (내부에서 CoInitialize).
"""
from __future__ import annotations

import shutil
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import pythoncom
import pywintypes
import win32com.client

MSO_TRUE = -1
MSO_FALSE = 0
MSO_GROUP = 6
MSO_FILL_SOLID = 1
PP_SAVE_AS_PPTX = 24
PP_ALERTS_NONE = 1
PP_WINDOW_MINIMIZED = 2
PP_VIEW_NORMAL = 9

MERGE_SOURCE_FORMATTING = "source_formatting"
MERGE_WINDOWLESS = "windowless"

LogFn = Callable[[str], None]


@dataclass
class MergeJob:
    source: Path
    title: str                                   # 로그/오류 표시용 (예: "찬송 29장")
    fill_text: str | None = None                 # 지정 텍스트 박스에 넣을 문구
    replacements: dict[str, str] = field(default_factory=dict)


class PptEngineError(RuntimeError):
    pass


class PptEngine:
    def __init__(self, log: LogFn | None = None, text_placeholder: str = "{{내용}}",
                 text_shape_name: str = "내용", merge_mode: str = MERGE_SOURCE_FORMATTING):
        self.log = log or (lambda m: None)
        self.text_placeholder = text_placeholder
        self.text_shape_name = text_shape_name
        self.merge_mode = merge_mode
        self.app = None
        self._was_running = False
        self._prev_window_state = None
        self._window = None
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
            self.app.DisplayAlerts = PP_ALERTS_NONE
        except pywintypes.com_error:
            pass
        self._tmp = Path(tempfile.mkdtemp(prefix="ChurchPPT_bg_"))
        return self

    def __exit__(self, *exc):
        try:
            if self._prev_window_state is not None and self._was_running:
                self.app.WindowState = self._prev_window_state
        except Exception:
            pass
        try:
            if self.app is not None and not self._was_running and self.app.Presentations.Count == 0:
                self.app.Quit()
        except Exception:
            pass
        self.app = None
        self._window = None
        if self._tmp:
            shutil.rmtree(self._tmp, ignore_errors=True)
        pythoncom.CoUninitialize()
        return False

    def _open(self, path: Path):
        try:
            # ReadOnly, Untitled, WithWindow
            return self.app.Presentations.Open(str(path), MSO_TRUE, MSO_FALSE, MSO_FALSE)
        except pywintypes.com_error as e:
            raise PptEngineError(f"PPT 파일을 열 수 없습니다:\n{path}\n{_com_msg(e)}") from e

    # ------------------------------------------------------------------ 병합
    def merge(self, jobs: list[MergeJob], output: Path, global_replacements: dict[str, str] | None = None,
              progress: Callable[[int, int, str], None] | None = None) -> Path:
        if not jobs:
            raise PptEngineError("병합할 PPT 가 없습니다.")
        progress = progress or (lambda *a: None)
        output = Path(output).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        self._design_cache.clear()

        dest = None
        try:
            # 1) 베이스: 첫 원본을 그대로 새 파일로 저장 (서식 100% 유지)
            progress(0, len(jobs), jobs[0].title)
            first = self._open(jobs[0].source)
            try:
                first.SaveAs(str(output), PP_SAVE_AS_PPTX)
            except pywintypes.com_error as e:
                first.Close()
                raise PptEngineError(f"결과 파일을 저장할 수 없습니다 (파일이 열려 있는지 확인):\n{output}\n{_com_msg(e)}") from e
            dest = first
            ranges = [(1, dest.Slides.Count)]
            self.log(f"[1/{len(jobs)}] {jobs[0].title} — {dest.Slides.Count}장")

            # 2) 나머지 원본 이어 붙이기
            for n, job in enumerate(jobs[1:], start=2):
                progress(n - 1, len(jobs), job.title)
                start = dest.Slides.Count + 1
                count = self._append(dest, job.source)
                ranges.append((start, start + count - 1))
                self.log(f"[{n}/{len(jobs)}] {job.title} — {count}장")

            # 3) 항목별 텍스트 입력 / 치환
            for job, (a, b) in zip(jobs, ranges):
                if b < a:
                    continue
                slides = [dest.Slides(i) for i in range(a, b + 1)]
                for k, v in job.replacements.items():
                    self._replace_in_slides(slides, k, v)
                if job.fill_text:
                    self._fill_text(slides, job)

            # 4) 공통 토큰 치환
            if global_replacements:
                all_slides = [dest.Slides(i) for i in range(1, dest.Slides.Count + 1)]
                for k, v in global_replacements.items():
                    self._replace_in_slides(all_slides, k, v)

            progress(len(jobs), len(jobs), "저장 중")
            dest.Save()
            self.log(f"저장 완료: {output} (총 {dest.Slides.Count}장)")
            return output
        finally:
            if dest is not None:
                try:
                    dest.Close()
                except Exception:
                    pass

    def _append(self, dest, source: Path) -> int:
        src = self._open(source)
        try:
            n = src.Slides.Count
            if n == 0:
                return 0
            if (abs(src.PageSetup.SlideWidth - dest.PageSetup.SlideWidth) > 1
                    or abs(src.PageSetup.SlideHeight - dest.PageSetup.SlideHeight) > 1):
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
        if self._window is not None:
            return self._window
        try:
            if self._was_running and self._prev_window_state is None:
                self._prev_window_state = self.app.WindowState
            win = dest.NewWindow()
            self.app.WindowState = PP_WINDOW_MINIMIZED
            win.ViewType = PP_VIEW_NORMAL
            self._window = win
            return win
        except pywintypes.com_error as e:
            raise PptEngineError(f"작업 창을 열 수 없음: {_com_msg(e)}") from e

    def _paste_source_formatting(self, src, dest, retries: int = 3) -> None:
        n = src.Slides.Count
        last = None
        for attempt in range(retries):
            before = dest.Slides.Count
            try:
                win = self._ensure_window(dest)
                src.Slides.Range().Copy()
                time.sleep(0.2)
                win.Activate()
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

    def _replace_in_slides(self, slides, find: str, repl: str) -> int:
        if not find:
            return 0
        total = 0
        for slide in slides:
            shapes = slide.Shapes
            for i in range(1, shapes.Count + 1):
                for tr in self._shape_text_ranges(shapes(i)):
                    total += _replace_all(tr, find, repl)
        return total

    def _fill_text(self, slides, job: MergeJob) -> None:
        if self.text_placeholder and self._replace_in_slides(slides, self.text_placeholder, job.fill_text):
            return
        if self.text_shape_name:
            for slide in slides:
                for i in range(1, slide.Shapes.Count + 1):
                    shp = slide.Shapes(i)
                    try:
                        if shp.Name == self.text_shape_name and shp.HasTextFrame:
                            shp.TextFrame.TextRange.Text = job.fill_text
                            return
                    except pywintypes.com_error:
                        continue
        self.log(f"  ⚠ '{job.title}': 텍스트 입력 위치를 찾지 못했습니다. "
                 f"템플릿에 '{self.text_placeholder}' 문구 또는 '{self.text_shape_name}' 이름의 텍스트 상자를 넣어 주세요.")


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


def _com_msg(e: pywintypes.com_error) -> str:
    try:
        exc = e.excepinfo
        if exc and exc[2]:
            return str(exc[2]).strip()
    except Exception:
        pass
    return str(e)
