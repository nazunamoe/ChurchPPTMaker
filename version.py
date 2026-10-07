"""프로그램 버전 — 유일한 기준. 화면 제목(main.APP_TITLE)과 빌드(build.ps1 → exe 파일 속성)가 모두 여기서 읽는다.

명령줄 (build.ps1 이 사용):
  python version.py                          현재 버전 출력
  python version.py --set 1.2.0              버전 변경 (이 파일의 __version__ 줄을 고침)
  python version.py --write-info <path>      PyInstaller --version-file 용 exe 버전 정보 파일 생성
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

__version__ = "1.0.3"

APP_NAME = "예배 순서 PPT 제작 프로그램"
EXE_NAME = "ChurchPPTMaker"

_VERSION_RE = re.compile(r"^\d+\.\d+\.\d+$")  # 주.부.수정 (예: 1.2.0)


def validate(version: str) -> str | None:
    """오류 메시지 반환 (정상이면 None)."""
    if not _VERSION_RE.match(version):
        return f"버전 형식이 잘못되었습니다: '{version}' (예: 1.2.0)"
    return None


def set_version(version: str) -> None:
    path = Path(__file__)
    text = path.read_text(encoding="utf-8")
    new, n = re.subn(r'^__version__ = "[^"]*"$', f'__version__ = "{version}"', text, count=1, flags=re.M)
    if n != 1:
        raise RuntimeError("version.py 에서 __version__ 줄을 찾지 못했습니다.")
    path.write_bytes(new.encode("utf-8"))


def version_info_text(version: str) -> str:
    """PyInstaller --version-file 형식 (탐색기 [속성] → [자세히] 에 표시되는 파일/제품 버전)."""
    nums = tuple(int(x) for x in version.split(".")) + (0,)
    fields = {
        "CompanyName": "",
        "FileDescription": APP_NAME,
        "FileVersion": version,
        "InternalName": EXE_NAME,
        "LegalCopyright": "",
        "OriginalFilename": f"{EXE_NAME}.exe",
        "ProductName": APP_NAME,
        "ProductVersion": version,
    }
    strings = ",\n".join(f"          StringStruct({k!r}, {v!r})" for k, v in fields.items())
    return f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={nums}, prodvers={nums}, mask=0x3f, flags=0x0, OS=0x40004,
                    fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([
      StringTable('041204B0', [
{strings}
      ])
    ]),
    VarFileInfo([VarStruct('Translation', [0x0412, 1200])])
  ]
)
"""


def _main(argv: list[str]) -> int:
    if not argv:
        print(__version__)
        return 0
    if argv[0] == "--set" and len(argv) == 2:
        err = validate(argv[1])
        if err:
            print(err, file=sys.stderr)
            return 1
        set_version(argv[1])
        print(argv[1])
        return 0
    if argv[0] == "--write-info" and len(argv) == 2:
        out = Path(argv[1])
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(version_info_text(__version__), encoding="utf-8")
        print(__version__)
        return 0
    print(__doc__, file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
