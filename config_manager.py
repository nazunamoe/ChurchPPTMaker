"""config.json / presets.json 관리."""
from __future__ import annotations

import copy
import json
import os
import sys
from pathlib import Path

from models import (DEFAULT_FILENAME_FORMAT, KIND_BIBLE, KIND_CUSTOM, KIND_HYMN, KIND_PRAYER, KIND_SERMON, MAX_FILL,
                    ServiceItem, validate_item)

MAX_PRESETS = 10
MAX_PRESET_NAME = 10
MAX_TEXT_IMPORT_BYTES = 1_000_000  # 예배 순서 텍스트 파일 불러오기 최대 크기 (다른 사람이 준 파일일 수 있음)


def app_dir() -> Path:
    if getattr(sys, "frozen", False):  # PyInstaller 등으로 패키징된 경우
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CONFIG_PATH = app_dir() / "config.json"
PRESETS_PATH = app_dir() / "presets.json"

DEFAULT_CONFIG: dict = {
    "church_name": "",
    "worship_type": "주일오전예배",
    "filename_format": DEFAULT_FILENAME_FORMAT,  # 결과 파일 이름 형식 (프리셋을 불러오면 프리셋의 형식으로 바뀜)
    "appearance_mode": "System",
    "use_system_accent": True,  # Windows 강조 색을 프로그램 강조 색으로 (다시 켤 때 적용)
    "setup_done": False,  # 설정 마법사를 한 번 보였음 (처음 실행할 때만 자동으로 띄움)
    "last_items": [],  # 마지막으로 작업하던 순서 (다시 켜면 그대로 보임 — 프리셋으로 저장하지 않아도 사라지지 않게)
    "paths": {
        "template_dir": "templates",
        "hymn_dir": "hymns",
        "ccm_dir": "ccm",
        "prayer_dir": "prayers",  # 기도 템플릿 폴더 (합심기도.pptx, 중보기도.pptx … — 기도 추가 시 하나 선택)
        "output_dir": "output",
    },
    "bible2ppt": {
        "exe_path": "",
        "bible": "",
        "bible2": "",  # 두 번째 기본 역본 (선택, 첫 역본과 나란히 배치)
        "template": "",
        "lines": None,
        "timeout_sec": 180,
        "bible_list": [],  # --list-bibles 결과 캐시 (exe 지정·↻ 시 갱신)
    },
    "item_templates": {
        "sermon_file": "",  # 설교 템플릿 PPT 파일 (하나)
    },
    "ppt": {
        "text_placeholder": "{{내용}}",
        "text_shape_name": "내용",
        "fill_defaults": ["", "", ""],  # {{내용1}}~{{내용3}} 값이 비었을 때 넣을 문구 (비우면 토큰을 지움)
        "open_after_build": True,
        "start_slideshow": False,  # 생성 후 열 때 바로 슬라이드 쇼로 시작
        "merge_mode": "source_formatting",  # source_formatting | windowless
        "slide_size": "original",  # 결과 슬라이드 비율: original(첫 PPT 크기) | 16:9 | 4:3
        "temp_keep": 10,  # 임시 파일 최대 보관 개수 (1~100, 넘으면 오래된 것부터 삭제)
    },
}


def clean_fill_defaults(raw) -> list[str]:
    """빈 칸 기본값을 정확히 MAX_FILL 개의 문자열로 (외부에서 고친 설정의 잘못된 값은 빈 문자열)."""
    raw = raw if isinstance(raw, list) else []
    return [(v if isinstance(v, str) else "") for v in raw[:MAX_FILL]] + [""] * (MAX_FILL - min(len(raw), MAX_FILL))


MAX_SAVED_ITEMS = 200  # 설정에 보관하는 '마지막 순서' 최대 항목 수 (예배 순서는 보통 10~30개)


def clean_items(raw) -> list[dict]:
    """외부에서 고칠 수 있는 항목 목록(설정의 last_items) → 검증된 항목 dict 목록.
    화면에서 만든 항목은 모두 validate_item 을 통과한 것이므로, 통과하지 못하는 항목(직접 고친 값)은 버린다."""
    if not isinstance(raw, list):
        return []
    items = (ServiceItem.from_dict(d) for d in raw if isinstance(d, dict))
    return [it.to_dict() for it in items if validate_item(it) is None][:MAX_SAVED_ITEMS]


TEMP_KEEP_DEFAULT = 10  # 임시 파일 최대 보관 개수 (ppt.temp_keep)
TEMP_KEEP_MAX = 100


def clean_temp_keep(value) -> int:
    """보관 개수를 1~TEMP_KEEP_MAX 정수로 (외부에서 고친 설정의 잘못된 값은 기본값)."""
    if isinstance(value, bool) or not isinstance(value, int):
        return TEMP_KEEP_DEFAULT
    return max(1, min(TEMP_KEEP_MAX, value))


def _item(kind, label, value="", **options):
    return ServiceItem(kind, label, value, options).to_dict()


DEFAULT_PRESETS = [
    {"name": "주일오전예배", "items": [
        _item(KIND_CUSTOM, "예배로 부름", "인도자"),
        _item(KIND_HYMN, "찬송", "8장"),
        _item(KIND_PRAYER, "기도", "OOO장로"),
        _item(KIND_BIBLE, "말씀", "요3:16"),
        _item(KIND_SERMON, "설교", "설교 제목", preacher="OOO목사"),
        _item(KIND_HYMN, "찬송", "29장"),
        _item(KIND_CUSTOM, "축도"),
    ]},
    {"name": "주일오후예배", "items": [
        _item(KIND_CUSTOM, "예배로 부름", "인도자"),
        _item(KIND_HYMN, "찬송", "29장"),
        _item(KIND_PRAYER, "기도", "OOO집사"),
        _item(KIND_BIBLE, "말씀", "롬8:28"),
        _item(KIND_CUSTOM, "주기도문"),
    ]},
    {"name": "수요예배", "items": [
        _item(KIND_HYMN, "찬송", "29장"),
        _item(KIND_PRAYER, "기도", "OOO권사"),
        _item(KIND_BIBLE, "말씀", "시23"),
        _item(KIND_CUSTOM, "주기도문"),
    ]},
]


def _deep_merge(base: dict, override: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in (override if isinstance(override, dict) else {}).items():
        if isinstance(out.get(k), dict):
            if isinstance(v, dict):
                out[k] = _deep_merge(out[k], v)
            # 기본값이 dict 인데 다른 타입이면 무시 (손상된 설정으로 인한 오류 방지)
        elif k in base and base[k] is not None and v is not None and not isinstance(v, type(base[k])):
            continue  # 타입 불일치 → 기본값 유지
        else:
            out[k] = v
    return out


def _atomic_write_json(path: Path, data) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except OSError:
        try:
            tmp.unlink(missing_ok=True)  # 디스크가 꽉 차는 등으로 반쯤 쓴 임시 파일
        except OSError:
            pass
        raise


SAVE_ERROR_HINT = ("프로그램 폴더에 쓸 수 없어 설정·프리셋이 저장되지 않습니다.\n"
                   "C:\\Program Files 처럼 쓰기가 막힌 곳이면 문서 폴더 등 다른 곳으로 프로그램을 옮겨 주세요.")


MAX_JSON_BYTES = 5_000_000  # 설정·프리셋 파일 최대 크기 (정상 파일은 수십 KB — 잘못 바뀐 아주 큰 파일을 읽느라 멈추지 않게)


def _decode_json(raw: bytes) -> dict:
    """설정 파일 내용 → dict. 메모장으로 고쳐 저장한 파일도 읽는다: UTF-8 BOM(예전 메모장 기본), UTF-16('유니코드'),
    ANSI(한글 Windows 의 CP949). 예전에는 BOM 만 붙어도 '손상'으로 보고 설정·프리셋을 기본값으로 되돌렸고,
    ANSI·UTF-16 은 시작하자마자 UnicodeDecodeError 로 꺼졌다. 맨 위가 {…} 가 아니면 ValueError."""
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        text = raw.decode("utf-16")
    else:
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("mbcs" if sys.platform == "win32" else "cp949")
    data = json.loads(text)
    if not isinstance(data, dict):  # [], "글자", 숫자, null — 예전에는 .get 호출에서 AttributeError 로 꺼졌음
        raise ValueError("JSON 맨 위가 {…} 형식이 아닙니다")
    return data


def _read_json(path: Path) -> tuple[dict | None, str | None]:
    """(내용, 알림). 파일이 없으면 (None, None). 읽을 수 없는(손상된) 파일은 '.broken' 으로 옮겨 보관하고
    (None, 사용자에게 보일 알림) — 예전에는 아무 말 없이 기본값으로 시작해 설정이 왜 사라졌는지 알 수 없었다."""
    try:
        if path.stat().st_size > MAX_JSON_BYTES:
            raise ValueError("파일이 너무 큽니다")
        return _decode_json(path.read_bytes()), None
    except FileNotFoundError:
        return None, None
    except (ValueError, OSError, RecursionError) as e:  # JSONDecodeError·UnicodeDecodeError 는 ValueError, 깊은 중첩은 RecursionError
        backup = path.with_suffix(path.suffix + ".broken")
        try:
            os.replace(path, backup)
            kept = f"원래 파일은 '{backup.name}' 이름으로 같은 폴더에 보관했습니다."
        except OSError:
            kept = "원래 파일은 보관하지 못했습니다."
        reason = str(e).splitlines()[0][:200] if str(e) else type(e).__name__
        return None, f"'{path.name}' 파일을 읽을 수 없어 기본값으로 시작했습니다. {kept}\n({reason})"


class ConfigManager:
    def __init__(self, path: Path = CONFIG_PATH):
        self.path = path
        self.save_error: str | None = None
        raw, self.load_notice = _read_json(path)  # load_notice: 손상된 파일을 기본값으로 바꿨으면 그 알림 (시작 후 한 번 표시)
        raw = raw or {}
        self.data = _deep_merge(DEFAULT_CONFIG, raw)
        # 설정 마법사 이전 버전을 쓰던 사람(설정 파일은 있는데 키가 없음)은 이미 설정했으므로 자동으로 띄우지 않음
        if raw and "setup_done" not in raw:
            self.data["setup_done"] = True
        # 1.0.3 초기 형식: 기도 폴더가 ⚙ 설정(item_templates.prayer_dir)에 있었음 → 왼쪽 [폴더](paths.prayer_dir)로 옮김
        old = (raw.get("item_templates") or {}).get("prayer_dir") if isinstance(raw.get("item_templates"), dict) else None
        if isinstance(old, str) and old.strip() and "prayer_dir" not in (raw.get("paths") or {}):
            self.data["paths"]["prayer_dir"] = old.strip()
        self.data["item_templates"].pop("prayer_dir", None)
        # 외부에서 고친 설정의 역본 목록에 문자열이 아닌 값이 있으면 화면(검색·표시)에서 오류가 나므로 걸러 냄
        bl = self.data["bible2ppt"].get("bible_list")
        self.data["bible2ppt"]["bible_list"] = [x for x in bl if isinstance(x, str)] if isinstance(bl, list) else []
        self.data["last_items"] = clean_items(self.data.get("last_items"))
        self.data["ppt"]["fill_defaults"] = clean_fill_defaults(self.data["ppt"].get("fill_defaults"))
        self.data["ppt"]["temp_keep"] = clean_temp_keep(self.data["ppt"].get("temp_keep"))
        # 기본값이 None 이라 타입 검사를 받지 않는 값 — 숫자가 아니면 말씀 생성 때 TypeError 가 났음
        lines = self.data["bible2ppt"].get("lines")
        if isinstance(lines, bool) or not isinstance(lines, int) or not 0 <= lines <= 9:
            self.data["bible2ppt"]["lines"] = None
        self.save()
        for key in ("template_dir", "hymn_dir", "ccm_dir", "prayer_dir", "output_dir"):
            try:
                self.resolve_path(key).mkdir(parents=True, exist_ok=True)
            except OSError:
                pass

    def save(self) -> None:
        """저장 실패(쓰기가 막힌 프로그램 폴더 등)는 save_error 에 남기고 계속한다 — 예전에는 시작하자마자 오류로 꺼졌음."""
        try:
            _atomic_write_json(self.path, self.data)
            self.save_error = None
        except OSError as e:
            self.save_error = f"{self.path}\n{e}"

    def get(self, section: str, key: str | None = None, default=None):
        if key is None:
            return self.data.get(section, default)
        return self.data.get(section, {}).get(key, default)

    def set(self, section: str, key: str | None, value, save: bool = True) -> None:
        if key is None:
            self.data[section] = value
        else:
            self.data.setdefault(section, {})[key] = value
        if save:
            self.save()

    def resolve_path(self, key: str) -> Path:
        """paths.<key> 를 절대 경로로 (상대 경로는 프로그램 폴더 기준)."""
        p = Path(os.path.expandvars(self.data["paths"].get(key) or DEFAULT_CONFIG["paths"][key]))
        return p if p.is_absolute() else (app_dir() / p)

    def resolve_file(self, raw: str) -> Path | None:
        if not raw:
            return None
        p = Path(os.path.expandvars(raw))
        return p if p.is_absolute() else (app_dir() / p)


def _clean_presets(raw: list) -> list[dict]:
    """presets.json 검증: 형식이 잘못된 프리셋은 버리고, 항목은 ServiceItem 으로 정규화."""
    out, seen = [], set()
    for p in raw:
        if not isinstance(p, dict):
            continue
        name = p.get("name")
        if not isinstance(name, str) or PresetManager.validate_name(name) or name in seen:
            continue
        items = p.get("items") if isinstance(p.get("items"), list) else []
        seen.add(name)
        entry = {"name": name, "items": [ServiceItem.from_dict(d).to_dict() for d in items if isinstance(d, dict)]}
        if isinstance(p.get("filename_format"), str) and p["filename_format"].strip():
            entry["filename_format"] = p["filename_format"]
        out.append(entry)
    return out


class PresetManager:
    def __init__(self, path: Path = PRESETS_PATH):
        self.path = path
        data, self.load_notice = _read_json(path)
        if not data or not isinstance(data.get("presets"), list):
            data = {"presets": copy.deepcopy(DEFAULT_PRESETS)}
        self.presets: list[dict] = _clean_presets(data["presets"])[:MAX_PRESETS]
        self.save_error: str | None = None
        try:
            self.save()
        except OSError as e:  # 쓰기가 막힌 프로그램 폴더 — 불러온 프리셋으로 계속 (저장할 때 다시 알림)
            self.save_error = f"{self.path}\n{e}"

    def save(self) -> None:
        """저장 실패는 OSError 그대로 (프리셋 저장·삭제를 누른 사람에게 알려야 함)."""
        _atomic_write_json(self.path, {"presets": self.presets})

    def names(self) -> list[str]:
        return [p["name"] for p in self.presets]

    def exists(self, name: str) -> bool:
        return name in self.names()

    @staticmethod
    def validate_name(name: str) -> str | None:
        if not name:
            return "프리셋 이름을 입력해 주세요."
        if any(ord(c) < 32 or ord(c) == 127 for c in name):  # 붙여 넣은 줄 바꿈·탭 (목록에서 이름이 깨져 보임)
            return "프리셋 이름에 줄 바꿈·탭 같은 문자는 쓸 수 없습니다."
        if len(name) > MAX_PRESET_NAME:
            return f"프리셋 이름은 {MAX_PRESET_NAME}글자 이내로 입력해 주세요. (현재 {len(name)}글자)"
        return None

    def can_add(self, name: str) -> bool:
        return self.exists(name) or len(self.presets) < MAX_PRESETS

    def upsert(self, name: str, items: list[ServiceItem], filename_format: str = DEFAULT_FILENAME_FORMAT) -> None:
        entry = {"name": name, "items": [i.to_dict() for i in items], "filename_format": filename_format}
        for idx, p in enumerate(self.presets):
            if p["name"] == name:
                self.presets[idx] = entry
                break
        else:
            if len(self.presets) >= MAX_PRESETS:
                raise ValueError(f"프리셋은 최대 {MAX_PRESETS}개까지 저장할 수 있습니다.")
            self.presets.append(entry)
        self.save()

    def load(self, name: str) -> list[ServiceItem]:
        for p in self.presets:
            if p["name"] == name:
                return [ServiceItem.from_dict(d) for d in p.get("items", [])]
        raise KeyError(name)

    def filename_format(self, name: str) -> str:
        """프리셋의 결과 파일 이름 형식 (저장된 형식이 없는 예전 프리셋은 기본 형식)."""
        p = next((p for p in self.presets if p["name"] == name), {})
        return p.get("filename_format") or DEFAULT_FILENAME_FORMAT

    def delete(self, name: str) -> None:
        self.presets = [p for p in self.presets if p["name"] != name]
        self.save()


def save_worship_to_text(path: Path, items: list[ServiceItem], metadata: dict) -> None:
    """예배 순서를 텍스트 파일로 저장."""
    from models import export_to_text
    text = export_to_text(
        items,
        church_name=metadata.get("church_name", ""),
        worship_type=metadata.get("worship_type", ""),
        date_str=metadata.get("date", ""),
        filename_format=metadata.get("filename_format", ""),
    )
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write(text)


def load_worship_from_text(path: Path) -> tuple[list[ServiceItem], dict]:
    """텍스트 파일에서 예배 순서를 읽어옴."""
    from models import import_from_text
    if path.stat().st_size > MAX_TEXT_IMPORT_BYTES:
        raise ValueError(f"파일이 너무 큽니다 (최대 {MAX_TEXT_IMPORT_BYTES // 1_000_000}MB). 예배 순서 텍스트 파일이 맞는지 확인해 주세요.")
    with open(path, "r", encoding="utf-8-sig") as f:
        text = f.read()
    return import_from_text(text)
