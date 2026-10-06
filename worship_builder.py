"""예배 PPT 생성 파이프라인: 항목 → 원본 파일 해석 → Bible2PPT 생성 → 병합."""
from __future__ import annotations

import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from bible_bridge import Bible2PPT, Bible2PPTError
from config_manager import ConfigManager
from models import KIND_BIBLE, KIND_CUSTOM, KIND_HYMN, KIND_PRAYER, ServiceItem, hymn_number
from ppt_engine import MERGE_SOURCE_FORMATTING, MergeJob, PptEngine, PptEngineError

PPT_EXTS = (".pptx", ".pptm", ".ppt")
_INVALID_FN = re.compile(r'[\\/:*?"<>|\r\n\t]+')


class BuildError(RuntimeError):
    pass


def sanitize_filename(text: str) -> str:
    return _INVALID_FN.sub("", text).strip().rstrip(".")


def output_filename(date: str, church: str, worship: str) -> str:
    """[YYYY-MM-DD]_[교회이름]_[예배종류].pptx"""
    return f"{sanitize_filename(date)}_{sanitize_filename(church)}_{sanitize_filename(worship)}.pptx"


def list_ppt_files(folder: Path, recursive: bool = False) -> list[Path]:
    if not folder.is_dir():
        return []
    it = folder.rglob("*") if recursive else folder.iterdir()
    return sorted((p for p in it if p.is_file() and p.suffix.lower() in PPT_EXTS
                   and not p.name.startswith("~$")), key=lambda p: p.name)


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", s).lower()


def find_template(folder: Path, name: str) -> Path | None:
    files = list_ppt_files(folder)
    for p in files:
        if p.stem == name:
            return p
    target = _norm(name)
    for p in files:
        if _norm(p.stem) == target:
            return p
    return None


def find_hymn(folder: Path, value: str) -> Path | None:
    files = list_ppt_files(folder, recursive=True)
    num = hymn_number(value)
    if num is None:  # 곡명으로 검색
        target = _norm(value)
        return next((p for p in files if target and target in _norm(p.stem)), None)
    exact = {str(num), f"{num:03d}", f"{num}장", f"{num:03d}장"}
    for p in files:
        if _norm(p.stem) in exact:
            return p
    for p in files:
        m = re.search(r"\d+", p.stem)
        if m and int(m.group()) == num:
            return p
    return None


@dataclass
class BuildRequest:
    items: list[ServiceItem]
    church: str
    date: str
    worship: str


class WorshipBuilder:
    def __init__(self, cfg: ConfigManager, log: Callable[[str], None],
                 progress: Callable[[float, str], None]):
        self.cfg = cfg
        self.log = log
        self.progress = progress

    # 사전 점검: 생성 전에 UI 에서 누락 파일을 한 번에 알려주기 위함
    def resolve(self, items: list[ServiceItem]) -> tuple[list[Path | None], list[str]]:
        tdir = self.cfg.resolve_path("template_dir")
        hdir = self.cfg.resolve_path("hymn_dir")
        paths: list[Path | None] = []
        problems: list[str] = []
        for idx, it in enumerate(items, 1):
            if it.kind == KIND_BIBLE:
                paths.append(None)  # 생성 시 Bible2PPT 가 만든다
                continue
            if it.kind == KIND_HYMN:
                p = find_hymn(hdir, it.value)
                if not p:
                    problems.append(f"{idx}. 찬송 '{it.value}' — 찬송가 폴더에서 PPT 를 찾지 못함")
            else:
                name = "기도" if it.kind == KIND_PRAYER else it.label
                p = find_template(tdir, name)
                if not p:
                    problems.append(f"{idx}. '{name}' — 템플릿 폴더에 '{name}.pptx' 없음")
            paths.append(p)
        if any(it.kind == KIND_BIBLE for it in items):
            try:
                self._bible().check()
            except Bible2PPTError as e:
                problems.append(f"말씀 — {e}")
        return paths, problems

    def _bible(self) -> Bible2PPT:
        b = self.cfg.get("bible2ppt")
        return Bible2PPT(self.cfg.resolve_file(b.get("exe_path", "")), int(b.get("timeout_sec") or 180))

    def build(self, req: BuildRequest) -> Path:
        if not req.items:
            raise BuildError("예배 순서가 비어 있습니다.")
        paths, problems = self.resolve(req.items)
        if problems:
            raise BuildError("다음 항목을 처리할 수 없습니다:\n\n" + "\n".join(problems))

        out_dir = self.cfg.resolve_path("output_dir")
        output = out_dir / output_filename(req.date, req.church, req.worship)
        tmp = Path(tempfile.mkdtemp(prefix="ChurchPPT_"))
        bcfg = self.cfg.get("bible2ppt")
        n_bible = sum(1 for it in req.items if it.kind == KIND_BIBLE)
        try:
            # 1) Bible2PPT 로 성경 PPT 생성 (PowerPoint 를 쓰므로 병합 전에 순차 실행)
            done = 0
            for idx, it in enumerate(req.items):
                if it.kind != KIND_BIBLE:
                    continue
                self.progress(0.4 * done / max(n_bible, 1), f"성경 PPT 생성 중: {it.value}")
                self.log(f"Bible2PPT: {it.value}")
                try:
                    paths[idx] = self._bible().generate(
                        it.value, tmp / f"bible_{idx:02d}.pptx",
                        bible=it.options.get("bible") or bcfg.get("bible") or None,
                        template=self.cfg.resolve_file(bcfg.get("template", "")),
                        lines=bcfg.get("lines"),
                    )
                except (Bible2PPTError, ValueError) as e:
                    raise BuildError(f"말씀 '{it.value}' 생성 실패:\n{e}") from e
                done += 1

            # 2) 병합 작업 목록
            pcfg = self.cfg.get("ppt")
            jobs = []
            for it, p in zip(req.items, paths):
                title = f"{it.label} {it.value}".strip()
                fill = it.value if it.kind in (KIND_PRAYER, KIND_CUSTOM) and it.value else None
                jobs.append(MergeJob(source=p, title=title, fill_text=fill))

            base = 0.4 if n_bible else 0.0

            def on_progress(i, total, label):
                self.progress(base + (1 - base) * i / max(total, 1), f"병합 중: {label}")

            with PptEngine(self.log, pcfg.get("text_placeholder", ""), pcfg.get("text_shape_name", ""),
                           pcfg.get("merge_mode") or MERGE_SOURCE_FORMATTING) as eng:
                result = eng.merge(jobs, output, {
                    "{{교회이름}}": req.church, "{{날짜}}": req.date, "{{예배종류}}": req.worship,
                }, on_progress)
            self.progress(1.0, "완료")
            return result
        except PptEngineError as e:
            raise BuildError(str(e)) from e
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
