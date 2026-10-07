"""config.json / presets.json 관리."""
from __future__ import annotations

import copy
import json
import os
import sys
from pathlib import Path

from models import KIND_BIBLE, KIND_CUSTOM, KIND_HYMN, KIND_PRAYER, ServiceItem

MAX_PRESETS = 10
MAX_PRESET_NAME = 10


def app_dir() -> Path:
    if getattr(sys, "frozen", False):  # PyInstaller 등으로 패키징된 경우
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CONFIG_PATH = app_dir() / "config.json"
PRESETS_PATH = app_dir() / "presets.json"

DEFAULT_CONFIG: dict = {
    "church_name": "",
    "worship_type": "주일오전예배",
    "appearance_mode": "System",
    "paths": {
        "template_dir": "templates",
        "hymn_dir": "hymns",
        "output_dir": "output",
    },
    "bible2ppt": {
        "exe_path": "",
        "bible": "",
        "template": "",
        "lines": None,
        "timeout_sec": 180,
        "bible_list": [],  # --list-bibles 결과 캐시 (exe 지정·↻ 시 갱신)
    },
    "ppt": {
        "text_placeholder": "{{내용}}",
        "text_shape_name": "내용",
        "open_after_build": True,
        "start_slideshow": False,  # 생성 후 열 때 바로 슬라이드 쇼로 시작
        "merge_mode": "source_formatting",  # source_formatting | windowless
    },
}


def _item(kind, label, value=""):
    return ServiceItem(kind, label, value).to_dict()


DEFAULT_PRESETS = [
    {"name": "주일오전예배", "items": [
        _item(KIND_CUSTOM, "예배로 부름", "인도자"),
        _item(KIND_HYMN, "찬송", "8장"),
        _item(KIND_PRAYER, "기도", "OOO장로"),
        _item(KIND_BIBLE, "말씀", "요3:16"),
        _item(KIND_CUSTOM, "설교"),
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
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _read_json(path: Path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return None
    except (json.JSONDecodeError, OSError):
        # 손상된 파일은 백업해 두고 기본값으로 재시작
        try:
            os.replace(path, path.with_suffix(path.suffix + ".broken"))
        except OSError:
            pass
        return None


class ConfigManager:
    def __init__(self, path: Path = CONFIG_PATH):
        self.path = path
        self.data = _deep_merge(DEFAULT_CONFIG, _read_json(path) or {})
        self.save()
        for key in ("template_dir", "hymn_dir", "output_dir"):
            try:
                self.resolve_path(key).mkdir(parents=True, exist_ok=True)
            except OSError:
                pass

    def save(self) -> None:
        _atomic_write_json(self.path, self.data)

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
        out.append({"name": name, "items": [ServiceItem.from_dict(d).to_dict() for d in items if isinstance(d, dict)]})
    return out


class PresetManager:
    def __init__(self, path: Path = PRESETS_PATH):
        self.path = path
        data = _read_json(path)
        if not data or not isinstance(data.get("presets"), list):
            data = {"presets": copy.deepcopy(DEFAULT_PRESETS)}
        self.presets: list[dict] = _clean_presets(data["presets"])[:MAX_PRESETS]
        self.save()

    def save(self) -> None:
        _atomic_write_json(self.path, {"presets": self.presets})

    def names(self) -> list[str]:
        return [p["name"] for p in self.presets]

    def exists(self, name: str) -> bool:
        return name in self.names()

    @staticmethod
    def validate_name(name: str) -> str | None:
        if not name:
            return "프리셋 이름을 입력해 주세요."
        if len(name) > MAX_PRESET_NAME:
            return f"프리셋 이름은 {MAX_PRESET_NAME}글자 이내로 입력해 주세요. (현재 {len(name)}글자)"
        return None

    def can_add(self, name: str) -> bool:
        return self.exists(name) or len(self.presets) < MAX_PRESETS

    def upsert(self, name: str, items: list[ServiceItem]) -> None:
        entry = {"name": name, "items": [i.to_dict() for i in items]}
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
        date_str=metadata.get("date", "")
    )
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write(text)


def load_worship_from_text(path: Path) -> tuple[list[ServiceItem], dict]:
    """텍스트 파일에서 예배 순서를 읽어옴."""
    from models import import_from_text
    with open(path, "r", encoding="utf-8-sig") as f:
        text = f.read()
    return import_from_text(text)
