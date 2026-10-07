"""Bible2PPT (https://github.com/nazunamoe/Bible2PPT) CLI 연동.

Bible2PPT.exe <성경 구절...> [-o 경로] [-b 성경] [-t 템플릿] [-l 0-9] [-s] [--open]
Bible2PPT.exe --list-bibles
종료 코드: 0 성공, 1 인자 오류, 2 PPT 생성 실패, 3 파워포인트 초기화 실패
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

# 책약자[시작장[:시작절][-[끝장[:끝절]|끝절]]]  — 예: 창, 창1, 롬1-3, 전1:3, 창1:1-5, 사1:3-3:9
_REF_RE = re.compile(r"^(\d?[가-힣A-Za-z]+)(\d+(:\d+)?(-\d+(:\d+)?)?)?$")

EXIT_MESSAGES = {
    1: "잘못된 인자입니다. 구절 형식이나 성경 이름을 확인해 주세요.",
    2: "성경 PPT 생성에 실패했습니다.",
    3: "파워포인트를 초기화하지 못했습니다.",
}

_NO_WINDOW = 0x08000000 if sys.platform == "win32" else 0  # CREATE_NO_WINDOW


class Bible2PPTError(RuntimeError):
    pass


def parse_references(text: str) -> list[str]:
    """'요 3:16, 롬8:28-39' → ['요3:16', '롬8:28-39']. 형식 오류 시 ValueError."""
    text = text.replace("：", ":").replace("～", "-").replace("~", "-").replace("–", "-")
    refs = []
    for part in re.split(r"[,;，；、/\n]+", text):
        ref = re.sub(r"\s+", "", part)
        if not ref:
            continue
        if not _REF_RE.match(ref):
            raise ValueError(
                f"'{part.strip()}' 은(는) 올바른 성경 구절 형식이 아닙니다.\n"
                "예) 창1:1-5, 롬1-3, 요3:16, 시23 (여러 구절은 쉼표로 구분)"
            )
        refs.append(ref)
    if not refs:
        raise ValueError("성경 구절을 입력해 주세요.")
    return refs


def format_bible(bible_id: str, source: str, name: str) -> str:
    """역본 목록 표시 문자열: '개역개정 · 갓피아 [12]' (bible_arg 가 [ID] 를 꺼내 -b 로 넘김)."""
    return f"{name} · {source} [{bible_id}]" if source else f"{name} [{bible_id}]"


def parse_bible_list(stdout: str) -> list[str]:
    """--list-bibles 출력('ID\\t소스\\t성경' 헤더 + 'ID\\t소스\\t이름' 줄)을 표시 문자열 목록으로."""
    result = []
    for line in stdout.splitlines():
        cols = [c.strip() for c in line.split("\t")]
        if not cols[0] or cols[:1] == ["ID"]:  # 빈 줄·헤더
            continue
        if len(cols) >= 3:
            result.append(format_bible(cols[0], cols[1], cols[2]))
        elif len(cols) == 2:
            result.append(format_bible(cols[0], "", cols[1]))
        else:
            result.append(cols[0])
    return result


_ENTRY_RE = re.compile(r"^(?P<name>.*?)(?: · (?P<source>.*?))? \[(?P<id>[^\[\]]+)\]$")


def split_bible(entry: str) -> tuple[str, str, str]:
    """표시 문자열 → (이름, 소스, ID). 형식이 아니면 (entry, '', '')."""
    m = _ENTRY_RE.match(entry.strip())
    if not m:
        return entry.strip(), "", ""
    return m.group("name"), m.group("source") or "", m.group("id")


def short_source(source: str) -> str:
    """'갓피아 성경' → '갓피아' (소스 필터 버튼용)."""
    return re.sub(r"\s*성경$", "", source.strip())


def bible_sources(entries: list[str]) -> list[str]:
    """목록에 나온 순서대로 소스 이름(짧은 형태) 목록."""
    seen: dict[str, None] = {}
    for e in entries:
        src = short_source(split_bible(e)[1])
        if src:
            seen.setdefault(src)
    return list(seen)


_CHO = "ㄱㄲㄴㄷㄸㄹㅁㅂㅃㅅㅆㅇㅈㅉㅊㅋㅌㅍㅎ"


def _char_eq(q: str, t: str) -> bool:
    """검색 글자 q 가 대상 글자 t 와 같거나, q 가 초성이고 t 의 초성과 같으면 일치."""
    if q == t:
        return True
    if q in _CHO and "가" <= t <= "힣":
        return _CHO[(ord(t) - 0xAC00) // 588] == q
    return False


def _contains(text: str, token: str) -> bool:
    n, m = len(text), len(token)
    return any(all(_char_eq(token[j], text[i + j]) for j in range(m)) for i in range(n - m + 1))


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", s).lower()


def filter_bibles(entries: list[str], query: str = "", source: str | None = None) -> list[str]:
    """역본 검색. 공백으로 나눈 검색어가 모두 이름·소스·ID 중 하나에 포함되면 일치 (대소문자·공백 무시, 초성 검색 가능).
    정렬: 이름이 검색어로 시작 → 이름에 포함 → 그 외(소스/ID 일치), 같은 순위는 원래 순서."""
    tokens = [_norm(t) for t in query.split() if t.strip()]
    ranked = []
    for idx, e in enumerate(entries):
        name, src, bid = split_bible(e)
        if source and short_source(src) != source:
            continue
        fields = [_norm(name), _norm(src), _norm(bid)]
        if not all(any(_contains(f, t) for f in fields) for t in tokens):
            continue
        rank = 0
        if tokens:
            first = tokens[0]
            nm = fields[0]
            if nm.startswith(first) or (len(nm) >= len(first) and all(_char_eq(first[j], nm[j]) for j in range(len(first)))):
                rank = 0
            elif _contains(nm, first):
                rank = 1
            else:
                rank = 2
        ranked.append((rank, idx, e))
    ranked.sort()
    return [e for _, _, e in ranked]


def search_names(names: list[str], query: str = "") -> list[int]:
    """파일 이름 등 일반 문자열 검색 → 일치하는 항목의 인덱스 (filter_bibles 와 같은 규칙: 공백 구분 AND,
    대소문자·공백 무시, 초성 가능). 정렬: 검색어로 시작 → 그 외, 같은 순위는 원래 순서."""
    tokens = [_norm(t) for t in query.split() if t.strip()]
    if not tokens:
        return list(range(len(names)))
    ranked = []
    for idx, name in enumerate(names):
        nm = _norm(name)
        if not all(_contains(nm, t) for t in tokens):
            continue
        first = tokens[0]
        starts = len(nm) >= len(first) and all(_char_eq(first[j], nm[j]) for j in range(len(first)))
        ranked.append((0 if starts else 1, idx, name))
    ranked.sort()
    return [i for _, i, _ in ranked]


def bible_arg(entry: str) -> str:
    """역본 표시 문자열에서 -b 에 넘길 값 추출.
    '개역개정 · 갓피아 [12]' → '12', 'ID\\t소스\\t이름' 원본 줄 → ID, 그 외에는 이름으로 간주."""
    entry = entry.strip()
    m = re.search(r"\[([^\[\]]+)\]$", entry)
    if m:
        return m.group(1).strip()
    if "\t" in entry:
        return entry.split("\t", 1)[0].strip()
    m = re.match(r"^(\S+)\s*[:|]?\s+\S", entry)
    if m and re.fullmatch(r"[\w.-]+", m.group(1)) and re.search(r"\d", m.group(1)):
        return m.group(1)
    return entry


def supports_cli(exe: Path) -> bool:
    """명령줄 지원 여부를 실행하지 않고 판별 (GUI 전용 버전은 인자를 무시하고 창을 띄운 채 멈춤).
    .NET Framework 어셈블리면 CLI 도움말 문자열(UTF-16)이 exe 에 그대로 들어 있으므로 그 유무로 판단.
    그 외 형식(단일 파일 압축 배포 등)은 판별할 수 없으므로 지원한다고 본다."""
    try:
        data = exe.read_bytes()
    except OSError:
        return True
    if b".NETFramework,Version=" not in data:
        return True
    return "--list-bibles".encode("utf-16-le") in data


def safe_bible_arg(entry: str) -> str:
    """-b 로 넘길 값. 옵션처럼 보이는 값('-'로 시작)이나 제어 문자는 거부 (프리셋 파일 조작 등으로 인한 인자 주입 방지)."""
    arg = bible_arg(entry)
    if arg.startswith("-") or any(ord(c) < 32 for c in arg):
        raise Bible2PPTError(f"올바르지 않은 성경 역본 값입니다: {entry!r}")
    return arg


def _timeout(value) -> int:
    try:
        return max(int(value), 10)
    except (TypeError, ValueError):
        return 180


class Bible2PPT:
    def __init__(self, exe_path: Path | None, timeout_sec=180):
        self.exe_path = exe_path
        self.timeout_sec = _timeout(timeout_sec)

    def check(self) -> None:
        if not self.exe_path:
            raise Bible2PPTError("Bible2PPT.exe 경로가 설정되지 않았습니다. ⚙ 설정에서 지정해 주세요.")
        if not Path(self.exe_path).is_file():
            raise Bible2PPTError(f"Bible2PPT.exe 를 찾을 수 없습니다:\n{self.exe_path}")
        if Path(self.exe_path).suffix.lower() != ".exe":
            # .bat/.cmd 는 cmd.exe 가 인자를 다시 해석하므로(& | 등) 인자 주입 위험 → 실행 파일만 허용
            raise Bible2PPTError(f"Bible2PPT 경로는 .exe 파일이어야 합니다:\n{self.exe_path}")
        if not supports_cli(Path(self.exe_path)):
            raise Bible2PPTError(
                "이 Bible2PPT 는 명령줄 실행을 지원하지 않는 버전입니다 (인자를 줘도 창만 열림).\n"
                "명령줄(--list-bibles 등)을 지원하는 Bible2PPT 를 지정해 주세요:\n"
                "https://github.com/nazunamoe/Bible2PPT")

    def _run(self, args: list[str]) -> subprocess.CompletedProcess:
        self.check()
        try:
            return subprocess.run(
                [str(self.exe_path), *args],
                capture_output=True, encoding="utf-8", errors="replace",
                timeout=self.timeout_sec, creationflags=_NO_WINDOW,
                cwd=str(Path(self.exe_path).parent),
            )
        except subprocess.TimeoutExpired as e:
            raise Bible2PPTError(f"Bible2PPT 실행 시간이 초과되었습니다 ({self.timeout_sec}초).") from e
        except OSError as e:
            raise Bible2PPTError(f"Bible2PPT 실행 실패: {e}") from e

    def generate(self, references: str, output: Path, bible: str | list[str] | None = None,
                 template: Path | None = None, lines: int | None = None) -> Path:
        refs = parse_references(references)
        args = [*refs, "-o", str(output)]
        bibles = [bible] if isinstance(bible, str) else (bible or [])
        for b in bibles:
            if b and b.strip():
                args += ["-b", safe_bible_arg(b)]
        if template:
            if Path(template).suffix.lower() != ".pptx" or not Path(template).is_file():
                raise Bible2PPTError(f"성경 템플릿은 존재하는 .pptx 파일이어야 합니다:\n{template}")
            args += ["-t", str(Path(template).resolve())]  # 절대 경로 → '-' 로 시작해 옵션으로 해석될 수 없음
        if lines is not None and 0 <= int(lines) <= 9:
            args += ["-l", str(int(lines))]

        proc = self._run(args)
        if proc.returncode != 0:
            msg = EXIT_MESSAGES.get(proc.returncode, f"알 수 없는 오류 (종료 코드 {proc.returncode})")
            detail = (proc.stderr or proc.stdout or "").strip()
            raise Bible2PPTError(f"{msg}\n구절: {' '.join(refs)}" + (f"\n\n{detail}" if detail else ""))

        # 성공 시 저장 경로가 stdout 으로 출력됨 — 마지막 유효 경로를 우선 사용.
        # 단, 지정한 출력 폴더 안의 파일만 인정 (엉뚱한 경로의 파일을 병합에 끌어들이지 않도록)
        out_dir = output.parent.resolve()
        for line in reversed((proc.stdout or "").splitlines()):
            p = Path(line.strip().strip('"'))
            if (line.strip() and p.suffix.lower() in (".pptx", ".ppt") and p.is_file()
                    and p.resolve().parent == out_dir):
                return p
        if output.is_file():
            return output
        raise Bible2PPTError(f"Bible2PPT 가 성공을 보고했지만 결과 파일이 없습니다:\n{output}")

    def list_bibles(self) -> list[str]:
        """사용 가능한 역본 표시 문자열 목록. 소스별 실패는 stderr 로만 보고되므로 목록이 비면 그 내용을 알림."""
        proc = self._run(["--list-bibles"])
        detail = (proc.stderr or "").strip()
        if proc.returncode != 0:
            msg = EXIT_MESSAGES[3] if proc.returncode == 3 else (
                f"성경 목록을 가져오지 못했습니다 (종료 코드 {proc.returncode}). "
                "--list-bibles 를 지원하는 Bible2PPT 인지 확인해 주세요.")
            raise Bible2PPTError(msg + (f"\n\n{detail}" if detail else ""))
        bibles = parse_bible_list(proc.stdout or "")
        if not bibles:
            raise Bible2PPTError("사용할 수 있는 성경 역본이 없습니다. 인터넷 연결을 확인해 주세요."
                                 + (f"\n\n{detail}" if detail else ""))
        return bibles
