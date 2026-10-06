"""예배 순서 항목 데이터 모델."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

KIND_CUSTOM = "custom"   # 왼쪽 이름의 PPT 호출 (+선택 문구)
KIND_PRAYER = "prayer"   # 기도.pptx 호출 + 기도자 이름 입력
KIND_HYMN = "hymn"       # 오른쪽 장 번호로 찬송가 폴더에서 검색
KIND_BIBLE = "bible"     # Bible2PPT 로 구절 PPT 생성

KIND_LABELS = {
    KIND_PRAYER: "기도",
    KIND_HYMN: "찬송",
    KIND_BIBLE: "말씀",
    KIND_CUSTOM: "사용자 지정",
}


@dataclass
class ServiceItem:
    kind: str
    label: str
    value: str = ""
    options: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {"kind": self.kind, "label": self.label, "value": self.value, "options": dict(self.options)}

    @classmethod
    def from_dict(cls, data: dict) -> "ServiceItem":
        kind = data.get("kind", KIND_CUSTOM)
        if kind not in KIND_LABELS:
            kind = KIND_CUSTOM
        return cls(kind=kind, label=str(data.get("label", "")), value=str(data.get("value", "")),
                   options=dict(data.get("options") or {}))

    @property
    def display_value(self) -> str:
        return self.value if self.value else "—"


def normalize_hymn(value: str) -> str:
    """'29' → '29장', '29 장' → '29장'. 숫자가 없으면 그대로(곡명 검색용)."""
    v = value.strip().replace(" ", "")
    if re.fullmatch(r"\d+", v):
        return f"{int(v)}장"
    m = re.fullmatch(r"(\d+)장", v)
    if m:
        return f"{int(m.group(1))}장"
    return value.strip()


def hymn_number(value: str) -> int | None:
    m = re.search(r"\d+", value)
    return int(m.group()) if m else None


def make_item(kind: str, label: str, value: str, options: dict | None = None) -> ServiceItem:
    """UI 다이얼로그 입력값으로 항목 생성. 검증은 validate_item 에서."""
    label = label.strip()
    value = value.strip()
    if kind == KIND_PRAYER:
        label = "기도"
    elif kind == KIND_HYMN:
        label = "찬송"
        value = normalize_hymn(value)
    elif kind == KIND_BIBLE:
        label = "말씀"
    return ServiceItem(kind, label, value, dict(options or {}))


def validate_item(item: ServiceItem) -> str | None:
    """오류 메시지 반환 (정상이면 None)."""
    from bible_bridge import parse_references  # 순환 import 방지

    if item.kind == KIND_PRAYER and not item.value:
        return "기도자 이름을 입력해 주세요."
    if item.kind == KIND_HYMN:
        if not item.value:
            return "찬송가 장을 입력해 주세요."
    if item.kind == KIND_BIBLE:
        if not item.value:
            return "성경 구절을 입력해 주세요."
        try:
            parse_references(item.value)
        except ValueError as e:
            return str(e)
    if item.kind == KIND_CUSTOM and not item.label:
        return "호출할 PPT 이름을 입력해 주세요."
    return None
