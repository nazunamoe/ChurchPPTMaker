"""예배 PPT 생성 파이프라인: 항목 → 원본 파일 해석 → Bible2PPT 생성 → 병합."""
from __future__ import annotations

import datetime as dt
import html
import os
import re
import shutil
import stat
import tempfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from bible_bridge import Bible2PPT, Bible2PPTError, bible_arg
from config_manager import ConfigManager, clean_temp_keep
from models import (BIBLE_KEYS, DEFAULT_FILENAME_FORMAT, FILENAME_FIELDS, KIND_BIBLE, KIND_CCM, KIND_CUSTOM, KIND_HYMN, KIND_PRAYER,
                    KIND_SERMON, MAX_FILL, ServiceItem, ccm_title, hymn_number, numbered_token)
from ppt_engine import MERGE_SOURCE_FORMATTING, MergeJob, PptEngine, PptEngineError

PPT_EXTS = (".pptx", ".pptm", ".ppt")
_INVALID_FN = re.compile(r'[\\/:*?"<>|\r\n\t]+')


class BuildError(RuntimeError):
    pass


# Windows 예약 장치 이름. 확장자가 붙어도('CON.txt') 장치로 해석되므로 첫 '.' 앞부분으로 판단한다.
_RESERVED = re.compile(r"^(CON|PRN|AUX|NUL|CONIN\$|CONOUT\$|COM[0-9¹²³]|LPT[0-9¹²³])$", re.IGNORECASE)
MAX_FILENAME_STEM = 150  # 파일 이름(확장자 제외) 최대 글자 수 — 255자·MAX_PATH 초과로 저장이 실패하지 않도록


def sanitize_filename(text: str) -> str:
    """파일명에 쓸 수 없는 문자(경로 구분자 포함)·제어 문자를 제거 → 다른 폴더로 빠져나가는 경로 생성 방지."""
    text = "".join(c for c in text if ord(c) >= 32)
    text = _INVALID_FN.sub("", text).strip()[:MAX_FILENAME_STEM].rstrip(". ")
    return f"_{text}" if _RESERVED.match(text.split(".", 1)[0].rstrip(" ")) else text


_FIELD_RE = re.compile(r"\{([^{}]*)\}")


def validate_filename_format(fmt: str) -> str | None:
    """파일 이름 형식 오류 메시지 (정상이면 None)."""
    if not fmt.strip():
        return "파일 이름 형식을 입력해 주세요."
    unknown = [f"{{{f}}}" for f in _FIELD_RE.findall(fmt) if f not in FILENAME_FIELDS]
    if unknown:
        return (f"파일 이름 형식에 알 수 없는 항목이 있습니다: {' '.join(unknown)}\n"
                f"사용할 수 있는 항목: {' '.join(f'{{{f}}}' for f in FILENAME_FIELDS)}")
    if not output_stem(fmt, "2000-01-01", "교회", "예배"):  # 예: '///' 처럼 쓸 수 없는 문자뿐인 형식
        return "파일 이름 형식으로 만든 이름이 비어 있습니다. 글자나 항목을 넣어 주세요."
    return None


def field_values(date: str, church: str, worship: str, pad: bool = True) -> dict[str, str]:
    """{항목} 이름(models.FILENAME_FIELDS) → 값. 파일 이름(pad=True: 월·일 0 채움 '07')과 PPT 공통 토큰
    (pad=False: '7' — 슬라이드에 '10월 7일' 처럼 쓰므로)이 같은 항목을 쓴다."""
    y, m, d = date.split("-") if re.fullmatch(r"\d{4}-\d{2}-\d{2}", date) else ("", "", "")
    try:
        weekday = "월화수목금토일"[dt.date(int(y), int(m), int(d)).weekday()]  # 한 글자: '일'
    except ValueError:  # 빈 날짜·없는 날짜(2026-02-30 등)
        weekday = ""
    if not pad:
        m, d = m.lstrip("0") or m, d.lstrip("0") or d
    return {"날짜": date, "년": y, "월": m, "일": d, "요일": weekday, "교회이름": church, "예배종류": worship}


def ppt_replacements(date: str, church: str, worship: str) -> dict[str, str]:
    """모든 슬라이드에서 바꾸는 공통 토큰 {{날짜}} {{년}} {{월}} {{일}} {{요일}} {{교회이름}} {{예배종류}}."""
    return {f"{{{{{k}}}}}": v for k, v in field_values(date, church, worship, pad=False).items()}


def output_stem(fmt: str, date: str, church: str, worship: str) -> str:
    """형식의 {항목}을 값으로 바꾼 파일 이름(확장자 제외). 전체를 sanitize 해 경로 구분자 등은 제거된다."""
    values = field_values(date, church, worship)
    name = _FIELD_RE.sub(lambda mo: values.get(mo.group(1), mo.group(0)), fmt).strip()
    if name.lower().endswith(".pptx"):  # 확장자까지 적었으면 한 번만 붙도록
        name = name[:-5]
    return sanitize_filename(name)


def output_filename(date: str, church: str, worship: str, fmt: str = DEFAULT_FILENAME_FORMAT) -> str:
    """파일 이름 형식(기본 {날짜}_{교회이름}_{예배종류}) → 'YYYY-MM-DD_교회이름_예배종류.pptx'.
    형식이 잘못돼 이름이 비면 기본 형식을 쓴다."""
    stem = output_stem(fmt, date, church, worship) or output_stem(DEFAULT_FILENAME_FORMAT, date, church, worship)
    return f"{stem}.pptx"


# ---- 임시 실행 모드: 결과를 출력 폴더 대신 임시 폴더에 만들고 바로 연다. 최근 ppt.temp_keep 개만 보관
def temp_result_dir() -> Path:
    """임시 결과 폴더. %TEMP% 는 Windows(저장소 센스·디스크 정리)가 지울 수 있어 보관 개수를 지킬 수 없으므로
    %LOCALAPPDATA%\\ChurchPPTMaker\\임시 (없으면 %TEMP% 아래)."""
    base = os.environ.get("LOCALAPPDATA")
    return (Path(base) if base else Path(tempfile.gettempdir())) / "ChurchPPTMaker" / "임시"


def _has_reparse_point(folder: Path) -> bool:
    """임시 폴더나 그 위 ChurchPPTMaker 폴더가 정션·심볼릭 링크면 True.
    링크를 따라가면 보관 개수 정리가 링크 대상(예: 문서 폴더)의 PPT 를 영구 삭제할 수 있으므로 그런 폴더는 쓰지 않는다."""
    for p in (folder, folder.parent):
        try:
            st = os.lstat(p)
        except OSError:  # 아직 없음
            continue
        if stat.S_ISLNK(st.st_mode) or getattr(st, "st_file_attributes", 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT:
            return True
    return False


def temp_dir_conflict(cfg: ConfigManager) -> str | None:
    """임시 폴더를 쓸 수 없으면 오류 메시지. 정리(prune_temp_results)는 임시 폴더의 .pptx 를 영구 삭제하므로
    원본·출력 폴더나 설교·성경 템플릿이 그 안에 있으면(설정 파일을 직접 고친 경우 등) 사용자 PPT 가 지워질 수 있다."""
    folder = temp_result_dir()
    if _has_reparse_point(folder):
        return f"임시 폴더가 바로 가기(정션·심볼릭 링크)로 되어 있어 사용하지 않습니다:\n{folder}"
    tmp = folder.resolve()
    for key, name in (("template_dir", "템플릿"), ("prayer_dir", "기도"), ("hymn_dir", "찬송가"), ("ccm_dir", "CCM"),
                      ("output_dir", "출력")):
        p = cfg.resolve_path(key).resolve()
        if p.is_relative_to(tmp):
            return (f"{name} 폴더가 임시 폴더 안에 있어, 임시 파일을 정리할 때 그 안의 PPT 가 지워질 수 있습니다:\n{p}\n"
                    "폴더를 다른 곳으로 지정해 주세요.")
    for p, name in ((cfg.resolve_file(cfg.get("item_templates", "sermon_file", "")), "설교 템플릿"),
                    (cfg.resolve_file(cfg.get("bible2ppt", "template", "")), "성경 템플릿")):
        if p is not None and p.resolve().is_relative_to(tmp):
            return f"{name}이 임시 폴더 안에 있어 임시 파일을 정리할 때 지워질 수 있습니다:\n{p}\n다른 곳으로 옮겨 주세요."
    return None


def _mtime(p: Path) -> float:
    try:
        return p.stat().st_mtime
    except OSError:  # 목록을 만드는 사이 지워짐
        return 0.0


def list_temp_results(folder: Path | None = None) -> list[Path]:
    """임시 결과 PPT (최근 것부터). PowerPoint 가 열어 둔 파일 옆에 만드는 잠금 파일 '~$이름.pptx' 와 링크는 제외.
    폴더가 정션·심볼릭 링크면 빈 목록 (정리가 링크 대상의 파일을 지우지 않도록)."""
    folder = folder or temp_result_dir()
    if _has_reparse_point(folder):
        return []
    try:
        files = [p for p in folder.glob("*.pptx")
                 if not p.name.startswith("~$") and p.is_file() and not p.is_symlink()]
    except OSError:
        return []
    return sorted(files, key=_mtime, reverse=True)


def prune_temp_results(keep: int, folder: Path | None = None) -> list[Path]:
    """최근 keep 개만 남기고 지운다. PowerPoint 에서 열려 있는 파일은 지울 수 없으므로 건너뜀(다음 정리 때 다시 시도)."""
    removed = []
    for p in list_temp_results(folder)[clean_temp_keep(keep):]:
        try:
            p.unlink()
            removed.append(p)
        except OSError:
            pass
    return removed


def temp_output_path(folder: Path, date: str, church: str, worship: str, fmt: str) -> Path:
    """임시 결과 경로: 형식대로 만든 이름 + _시분초 (앞서 연 임시 파일이 PowerPoint 에서 열려 있어도 겹치지 않게)."""
    stem = output_filename(date, church, worship, fmt)[:-5]
    base = f"{stem}_{dt.datetime.now():%H%M%S}"
    out, n = folder / f"{base}.pptx", 2
    while out.exists() or out.is_symlink():
        out, n = folder / f"{base}_{n}.pptx", n + 1
    return out


_TEMP_SUFFIX = re.compile(r"_\d{6}(_\d+)?$")


def saved_name_for_temp(path: Path) -> str:
    """임시 파일을 저장할 때 제안할 이름 (붙였던 _시분초 를 뗌)."""
    return f"{_TEMP_SUFFIX.sub('', path.stem) or path.stem}.pptx"


_IO_REPARSE_TAG_MOUNT_POINT = 0xA0000003  # 정션


def _is_link_dir(entry: os.DirEntry) -> bool:
    """심볼릭 링크·정션 폴더. 따라 들어가면 자기 자신을 가리키는 링크에서 끝없이 반복한다(실측: 정션 2개에 10초 넘게
    15000개 이상 — 찬송가 찾기가 화면 스레드를 멈춤). OneDrive 같은 다른 재분석 지점 폴더는 그대로 들어간다."""
    if entry.is_symlink():
        return True
    try:
        return os.lstat(entry.path).st_reparse_tag == _IO_REPARSE_TAG_MOUNT_POINT
    except (OSError, AttributeError):
        return False


def _walk_files(folder: Path, recursive: bool):
    """폴더의 파일 (recursive 면 하위 폴더까지, 링크 폴더는 들어가지 않음). 읽을 수 없는 폴더는 건너뜀."""
    stack = [str(folder)]
    while stack:
        try:
            with os.scandir(stack.pop()) as it:
                entries = list(it)
        except OSError:
            continue
        for e in entries:
            try:
                if e.is_dir():
                    if recursive and not _is_link_dir(e):
                        stack.append(e.path)
                elif e.is_file():
                    yield e
            except OSError:
                continue


def list_ppt_files(folder: Path, recursive: bool = False) -> list[Path]:
    if not folder.is_dir():
        return []
    return sorted((Path(e.path) for e in _walk_files(folder, recursive)
                   if e.name.lower().endswith(PPT_EXTS) and not e.name.startswith("~$")), key=lambda p: p.name)


def count_ppt_files(folder: Path, recursive: bool = False, budget: float = 0.3) -> int | None:
    """폴더의 PPT 개수 (list_ppt_files 와 같은 기준). budget 초 안에 다 못 세면 None.
    설정 마법사가 개수를 보여 줄 때 사용 — 사용자 폴더처럼 큰 폴더를 하위까지 세면 화면이 수십 초 멈췄음(실측 19초)."""
    deadline = time.monotonic() + budget
    n, stack = 0, [folder]
    while stack:
        try:
            with os.scandir(stack.pop()) as it:
                for e in it:
                    if time.monotonic() > deadline:
                        return None
                    try:
                        if e.is_dir():
                            if recursive and not _is_link_dir(e):
                                stack.append(e.path)
                        elif e.name.lower().endswith(PPT_EXTS) and not e.name.startswith("~$") and e.is_file():
                            n += 1
                    except OSError:
                        continue
        except OSError:  # 권한 없는 하위 폴더 등은 건너뜀
            continue
    return n


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", s).lower()


class FileLists:
    """폴더의 PPT 목록을 한 번만 읽어 다시 쓴다 (list_ppt_files 와 같은 호출 형태 — find_* 의 lister 로 넘김).
    생성 전 점검은 찬송·CCM 항목마다 폴더를 하위까지 다시 훑었다(항목 12개 = 12번, 화면 스레드에서 한 번 + 작업
    스레드에서 한 번). 점검 한 번 동안만 쓰고 버리므로 그사이 넣은 파일을 못 보는 일은 없다."""

    def __init__(self):
        self._lists: dict[tuple[str, bool], list[Path]] = {}

    def __call__(self, folder: Path, recursive: bool = False) -> list[Path]:
        key = (str(folder), recursive)
        if key not in self._lists:
            self._lists[key] = list_ppt_files(folder, recursive)
        return self._lists[key]


def find_template(folder: Path, name: str, lister=list_ppt_files) -> Path | None:
    files = lister(folder)
    for p in files:
        if p.stem == name:
            return p
    target = _norm(name)
    for p in files:
        if _norm(p.stem) == target:
            return p
    return None


_SLIDE_XML = re.compile(r"ppt/slides/slide\d+\.xml")
_A_T = re.compile(r"<a:t(?:\s[^>]*)?>([^<]*)</a:t>")
_SHAPE_NAME = re.compile(r'<p:cNvPr\b[^>]*?\sname="([^"]*)"')
_MAX_SLIDE_XML = 20_000_000  # 슬라이드 XML 하나의 최대 크기 (비정상적으로 큰 파일은 읽지 않음)
_MAX_TOTAL_XML = 100_000_000  # 슬라이드 XML 합계 최대 크기


def template_fill_slots(path: Path, placeholder: str, shape_name: str) -> set[int] | None:
    """템플릿 PPT 의 슬라이드에서 텍스트 입력 자리 번호(1~MAX_FILL)를 찾는다 — 사용자 지정 항목 입력칸 표시용.
    {{내용n}} 토큰(설정의 치환 토큰 기본형에 번호) 또는 이름이 '내용n' 인 도형이 있으면 n. 1번은 번호 없는 기본형도 인정.
    PowerPoint 를 띄우지 않도록 .pptx/.pptm 의 슬라이드 XML 을 직접 읽는다. 토큰이 여러 글자 조각(run)으로 나뉘어
    저장되는 경우가 많아 문단 단위로 글자를 이어 붙여 찾는다. 읽을 수 없는 형식(.ppt)·손상 파일이면 None."""
    if path.suffix.lower() not in (".pptx", ".pptm"):
        return None
    texts: list[str] = []
    names: set[str] = set()
    budget = _MAX_TOTAL_XML
    try:
        with zipfile.ZipFile(path) as z:
            for info in z.infolist():
                if not _SLIDE_XML.fullmatch(info.filename) or info.file_size > _MAX_SLIDE_XML:
                    continue
                budget -= info.file_size
                if budget < 0:  # 압축 폭탄처럼 풀면 아주 커지는 파일 — 화면이 멈추지 않게 읽지 않음
                    return None
                xml = z.read(info).decode("utf-8", "replace")
                texts += [html.unescape("".join(_A_T.findall(chunk))) for chunk in xml.split("</a:p>")]
                names.update(html.unescape(n) for n in _SHAPE_NAME.findall(xml))
    except (zipfile.BadZipFile, OSError, KeyError, RuntimeError, ValueError):
        return None
    slots = set()
    for n in range(1, MAX_FILL + 1):
        token = numbered_token(placeholder, n)
        if (token and any(token in t for t in texts)) or (shape_name and f"{shape_name}{n}" in names):
            slots.add(n)
    if (placeholder and any(placeholder in t for t in texts)) or (shape_name and shape_name in names):
        slots.add(1)  # 예전 템플릿의 {{내용}} / '내용' 은 1번
    return slots


def prayer_template_names(cfg: ConfigManager) -> list[str] | None:
    """설정의 기도 템플릿 폴더에 있는 PPT 이름 (기도 종류). 폴더가 지정되지 않았거나 없으면 None."""
    folder = cfg.resolve_path("prayer_dir")
    if not folder.is_dir():
        return None
    return [p.stem for p in list_ppt_files(folder)]


def sermon_template(cfg: ConfigManager) -> Path | None:
    """설정의 설교 템플릿 PPT. 지정되지 않았거나 파일이 없으면 None."""
    p = cfg.resolve_file(cfg.get("item_templates", "sermon_file", ""))
    return p if p and p.is_file() and p.suffix.lower() in PPT_EXTS else None


def item_bibles(it: ServiceItem, cfg: ConfigManager) -> list[str]:
    """말씀 항목의 역본 목록 (-b 순서, 2개면 Bible2PPT 가 나란히 배치).
    항목에 역본을 지정했으면 그 역본1(비면 설정 기본 역본1)·역본2, 아니면 설정의 기본 역본1·역본2."""
    b = cfg.get("bible2ppt")
    own = [(it.options.get(k) or "").strip() for k in BIBLE_KEYS]  # 공백뿐인 값(직접 고친 프리셋 등)은 빈 값
    if any(own):
        pair = (own[0] or (b.get("bible") or "").strip(), own[1])
    else:
        pair = ((b.get("bible") or "").strip(), (b.get("bible2") or "").strip())
    out: list[str] = []
    seen: set[str] = set()
    for v in pair:
        key = bible_arg(v).casefold()  # '개역개정 · 갓피아 [1]' 과 '1', 'KJV' 와 'kjv' 는 같은 역본 → 한 번만 넘김
        if v and key not in seen:
            seen.add(key)
            out.append(v)
    return out


def bible_template(cfg: ConfigManager) -> tuple[Path | None, str | None]:
    """설정의 성경 템플릿 → (경로, 오류). 지정 안 했으면 (None, None).
    Bible2PPT 는 자기 PowerPoint 세션에서 이 파일을 열어 이 프로그램의 매크로 차단(AutomationSecurity)이 미치지 않으므로
    매크로를 담을 수 없는 .pptx 만 허용한다."""
    raw = (cfg.get("bible2ppt", "template", "") or "").strip()
    if not raw:
        return None, None
    p = cfg.resolve_file(raw)
    if p.suffix.lower() != ".pptx":
        return None, f"성경 템플릿은 .pptx 파일만 사용할 수 있습니다: {raw}"
    if not p.is_file():
        return None, f"성경 템플릿 파일을 찾을 수 없습니다: {raw}"
    return p, None


def output_conflict(cfg: ConfigManager, output: Path, sources=()) -> str | None:
    """결과 파일을 저장하면 원본(템플릿·찬송가·CCM·기도·설교·성경 템플릿)을 덮어쓸 수 있으면 오류 메시지.
    파일 이름 형식을 바꿀 수 있어서, 출력 폴더가 원본 폴더면 이번 생성에 쓰지 않는 원본도 같은 이름으로 덮어쓸 수 있다."""
    if output.is_symlink():  # 링크를 따라가 엉뚱한 파일을 덮어쓰지 않도록
        return f"결과 파일 자리에 바로 가기(심볼릭 링크)가 있어 저장하지 않습니다:\n{output}"
    out = output.resolve()
    named = [(p, "원본 PPT") for p in sources if p is not None]
    named += [(cfg.resolve_file(cfg.get("item_templates", "sermon_file", "")), "설교 템플릿"),
              (cfg.resolve_file(cfg.get("bible2ppt", "template", "")), "성경 템플릿")]
    for p, name in named:
        if p is not None and p.resolve() == out:
            return f"결과 파일 경로가 {name}와 같아 원본을 덮어쓰게 됩니다:\n{output}\n출력 폴더나 파일 이름 형식을 바꿔 주세요."
    for key, name in (("template_dir", "템플릿"), ("prayer_dir", "기도"), ("hymn_dir", "찬송가"), ("ccm_dir", "CCM")):
        folder = cfg.resolve_path(key).resolve()
        if out.parent.is_relative_to(folder):
            return (f"출력 폴더가 {name} 폴더와 같거나 그 안에 있어 원본 PPT 를 덮어쓸 수 있습니다:\n{out.parent}\n"
                    "출력 폴더를 다른 곳으로 지정해 주세요.")
    return None


def folder_conflict(cfg: ConfigManager, key: str, path: str | Path) -> str | None:
    """폴더 지정(📁·설정 마법사) 때 쓸 수 없는 폴더면 오류 메시지. 임시 폴더 안은 임시 파일 정리 때 그 안의 PPT 가
    지워질 수 있어 거부하고, 출력 폴더는 원본 폴더와 겹쳐 원본을 덮어쓰지 않는지 미리 확인한다 (cfg 의 다른 폴더 기준)."""
    try:
        in_temp = Path(path).resolve().is_relative_to(temp_result_dir().resolve())
    except OSError:
        in_temp = False
    if in_temp:
        return f"임시 파일 폴더 안은 지정할 수 없습니다 (오래된 PPT 가 자동으로 지워짐):\n{path}"
    if key == "output_dir":
        return output_conflict(cfg, Path(path) / "_.pptx")
    return None


def find_hymn(folder: Path, value: str, lister=list_ppt_files) -> Path | None:
    files = lister(folder, recursive=True)
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


def list_ccm_files(folder: Path) -> list[str]:
    """CCM 폴더(하위 포함)의 PPT → 폴더 기준 상대 경로('/' 구분). 파일명이 정형화돼 있지 않아 검색해서 고른다."""
    files = list_ppt_files(folder, recursive=True)
    return sorted((p.relative_to(folder).as_posix() for p in files), key=lambda r: (ccm_title(r).lower(), r))


def find_ccm(folder: Path, value: str, lister=list_ppt_files) -> Path | None:
    """① 저장된 상대 경로 그대로 ② (파일이 다른 하위 폴더로 옮겨졌으면) 같은 파일 이름 ③ 공백·대소문자 무시 이름.
    presets.json 등은 외부 입력이므로 CCM 폴더 밖을 가리키는 경로('..', 절대 경로)는 인정하지 않는다."""
    value = value.replace("\\", "/").strip()
    if not value or not folder.is_dir():
        return None
    root = folder.resolve()
    rel = Path(value)
    # 절대 경로·드라이브·'..' 는 resolve 전에 거른다 — resolve 가 '\\서버\공유' 경로를 열려고 네트워크에 접속해
    # (실측 21초 멈춤) Windows 로그인 정보(NTLM)가 그 서버로 보내질 수 있음
    if not rel.anchor and ".." not in rel.parts:
        p = (root / rel).resolve()
        if p.is_relative_to(root) and p.is_file() and p.suffix.lower() in PPT_EXTS:
            return p
    files = lister(folder, recursive=True)
    name = value.rsplit("/", 1)[-1]
    for q in files:
        if q.name == name:
            return q
    title = _norm(ccm_title(value))
    return next((q for q in files if _norm(q.stem) == title), None)


@dataclass
class BuildRequest:
    items: list[ServiceItem]
    church: str
    date: str
    worship: str
    filename_format: str = DEFAULT_FILENAME_FORMAT
    temp: bool = False  # True: 출력 폴더 대신 임시 폴더에 만들고 오래된 임시 파일 정리
    open_mode: str = ""  # "open"/"show": 다 만들면 병합에 쓴 PowerPoint 에서 바로 열기(쇼)
    opened: bool = False  # build 가 결과를 열었으면 True (못 열었으면 호출한 쪽이 파일로 연다)


class WorshipBuilder:
    def __init__(self, cfg: ConfigManager, log: Callable[[str], None],
                 progress: Callable[[float, str], None]):
        self.cfg = cfg
        self.log = log
        self.progress = progress

    # 사전 점검: 생성 전에 UI 에서 누락 파일을 한 번에 알려주기 위함
    def resolve_item(self, it: ServiceItem, lister=list_ppt_files) -> tuple[Path | None, str | None]:
        """항목 하나의 원본 PPT → (경로, 문제). 말씀은 생성할 때 만들어지므로 (None, None).
        생성 전 점검과 순서 목록 [열기](원본 열기)가 같은 규칙으로 파일을 찾도록 한 곳에 둔다.
        lister: 폴더 목록 읽기 (여러 항목을 한꺼번에 볼 때는 FileLists 로 폴더를 한 번만 읽음)."""
        if it.kind == KIND_BIBLE:
            return None, None
        if it.kind == KIND_HYMN:
            p = find_hymn(self.cfg.resolve_path("hymn_dir"), it.value, lister)
            return p, None if p else f"찬송 '{it.value}' — 찬송가 폴더에서 PPT 를 찾지 못함"
        if it.kind == KIND_CCM:
            p = find_ccm(self.cfg.resolve_path("ccm_dir"), it.value, lister)
            return p, None if p else f"CCM '{it.display_value}' — CCM 폴더에서 PPT 를 찾지 못함"
        if it.kind == KIND_PRAYER:
            p = find_template(self.cfg.resolve_path("prayer_dir"), it.label, lister)
            return p, None if p else f"'{it.label}' — 기도 템플릿 폴더에 '{it.label}.pptx' 없음"
        if it.kind == KIND_SERMON:
            p = sermon_template(self.cfg)
            if not self.cfg.get("item_templates", "sermon_file", ""):
                return None, "설교 — ⚙ 설정에서 설교 템플릿 PPT 를 지정해 주세요"
            return p, None if p else ("설교 — 설교 템플릿 PPT 를 찾을 수 없음: "
                                      f"{self.cfg.get('item_templates', 'sermon_file')}")
        p = find_template(self.cfg.resolve_path("template_dir"), it.label, lister)
        return p, None if p else f"'{it.label}' — 템플릿 폴더에 '{it.label}.pptx' 없음"

    def check_bible(self) -> str | None:
        """말씀 생성에 필요한 설정 문제 (Bible2PPT exe·성경 템플릿). 없으면 None."""
        try:
            self._bible().check()
        except Bible2PPTError as e:
            return str(e)
        return bible_template(self.cfg)[1]

    def generate_bible(self, it: ServiceItem, output: Path) -> Path:
        """말씀 항목 → Bible2PPT 로 output 에 PPT 생성 (역본은 item_bibles, 없으면 Bible2PPT 마지막 선택값)."""
        bcfg = self.cfg.get("bible2ppt")
        return self._bible().generate(it.value, output, bible=item_bibles(it, self.cfg) or None,
                                      template=bible_template(self.cfg)[0], lines=bcfg.get("lines"))

    def resolve(self, items: list[ServiceItem]) -> tuple[list[Path | None], list[str]]:
        paths: list[Path | None] = []
        problems: list[str] = []
        lister = FileLists()  # 이 점검 동안 폴더마다 한 번만 읽음
        for idx, it in enumerate(items, 1):
            p, problem = self.resolve_item(it, lister)
            if problem:
                problems.append(f"{idx}. {problem}")
            paths.append(p)
        if any(it.kind == KIND_BIBLE for it in items) and (err := self.check_bible()) is not None:
            problems.append(f"말씀 — {err}")
        return paths, problems

    def _bible(self) -> Bible2PPT:
        b = self.cfg.get("bible2ppt")
        return Bible2PPT(self.cfg.resolve_file(b.get("exe_path", "")), b.get("timeout_sec"))

    def build(self, req: BuildRequest) -> Path:
        if not req.items:
            raise BuildError("예배 순서가 비어 있습니다.")
        paths, problems = self.resolve(req.items)
        if problems:
            raise BuildError("다음 항목을 처리할 수 없습니다:\n\n" + "\n".join(problems))

        if req.temp:
            if (err := temp_dir_conflict(self.cfg)) is not None:
                raise BuildError(err)
            out_dir = temp_result_dir()
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
            except OSError as e:
                raise BuildError(f"임시 폴더를 만들 수 없습니다:\n{out_dir}\n{e}") from e
            output = temp_output_path(out_dir, req.date, req.church, req.worship, req.filename_format)
        else:
            out_dir = self.cfg.resolve_path("output_dir")
            output = out_dir / output_filename(req.date, req.church, req.worship, req.filename_format)
        if (err := output_conflict(self.cfg, output, paths)) is not None:
            raise BuildError(err)
        tmp = Path(tempfile.mkdtemp(prefix="ChurchPPT_"))
        n_bible = sum(1 for it in req.items if it.kind == KIND_BIBLE)
        pcfg = self.cfg.get("ppt")
        try:
            # PowerPoint 를 먼저 켜 두고 Bible2PPT 가 같이 쓰게 한다 (말씀마다 PowerPoint 를 켜고 끄지 않음, PptEngine.hold)
            with PptEngine(self.log, pcfg.get("text_placeholder", ""), pcfg.get("text_shape_name", ""),
                           pcfg.get("merge_mode") or MERGE_SOURCE_FORMATTING, pcfg.get("fill_defaults"),
                           pcfg.get("slide_size") or "original") as eng:
                # 1) Bible2PPT 로 성경 PPT 생성 (PowerPoint 를 쓰므로 병합 전에 순차 실행)
                if n_bible:
                    eng.hold()
                done = 0
                for idx, it in enumerate(req.items):
                    if it.kind != KIND_BIBLE:
                        continue
                    self.progress(0.4 * done / max(n_bible, 1), f"성경 PPT 생성 중: {it.value}")
                    self.log(f"Bible2PPT: {it.value}")
                    try:
                        paths[idx] = self.generate_bible(it, tmp / f"bible_{idx:02d}.pptx")
                    except (Bible2PPTError, ValueError) as e:
                        raise BuildError(f"말씀 '{it.value}' 생성 실패:\n{e}") from e
                    done += 1
                eng.release()

                # 2) 병합 작업 목록
                jobs = []
                for it, p in zip(req.items, paths):
                    title = f"{it.label} {it.display_value if it.kind in (KIND_SERMON, KIND_CCM, KIND_CUSTOM) else it.value}".strip()
                    jobs.append(MergeJob(source=p, title=title, fill_values=it.fill_values,
                                         warn_missing=it.kind not in (KIND_HYMN, KIND_CCM)))

                base = 0.4 if n_bible else 0.0

                def on_progress(i, total, label):
                    self.progress(base + (1 - base) * i / max(total, 1), f"병합 중: {label}")

                req.opened = eng.merge(jobs, output, ppt_replacements(req.date, req.church, req.worship),
                                       on_progress, req.open_mode)
            result = output
            if req.temp:
                removed = prune_temp_results(pcfg.get("temp_keep"), out_dir)
                if removed:
                    self.log(f"오래된 임시 파일 {len(removed)}개 정리 (최근 {clean_temp_keep(pcfg.get('temp_keep'))}개 보관)")
            self.progress(1.0, "완료")
            return result
        except BaseException as e:
            if req.temp:  # 실패하면 반쯤 만든 임시 결과가 [임시 파일] 목록에 남지 않게 지움 (일반 생성은 사용자 폴더라 그대로)
                try:
                    output.unlink(missing_ok=True)
                except OSError:
                    pass
            if isinstance(e, PptEngineError):
                raise BuildError(str(e)) from e
            raise
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
