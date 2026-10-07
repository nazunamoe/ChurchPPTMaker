"""예배 순서 항목 데이터 모델."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

KIND_CUSTOM = "custom"   # 왼쪽 이름의 PPT 호출 (+선택 문구)
KIND_PRAYER = "prayer"   # 기도 템플릿 폴더에서 label 이름의 PPT 호출 + 기도자 이름 입력
KIND_HYMN = "hymn"       # 오른쪽 장 번호로 찬송가 폴더에서 검색
KIND_BIBLE = "bible"     # Bible2PPT 로 구절 PPT 생성
KIND_SERMON = "sermon"   # 설교.pptx 호출 + 설교 제목({{내용1}})·설교자({{내용2}}) 입력
KIND_CCM = "ccm"         # CCM 폴더에서 검색해 고른 파일 (value = CCM 폴더 기준 상대 경로)

MAX_FILL = 3             # 항목당 텍스트 입력칸 수 ({{내용1}} ~ {{내용3}})

# 결과 파일 이름 형식 (프리셋별로 다르게 저장 가능). {항목} 자리에 값이 들어감 — worship_builder.output_filename
FILENAME_FIELDS = ("날짜", "년", "월", "일", "교회이름", "예배종류")
DEFAULT_FILENAME_FORMAT = "{날짜}_{교회이름}_{예배종류}"

KIND_LABELS = {
    KIND_PRAYER: "기도",
    KIND_HYMN: "찬송",
    KIND_BIBLE: "말씀",
    KIND_SERMON: "설교",
    KIND_CCM: "CCM",
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
        opts = data.get("options")
        opts = {str(k): v for k, v in opts.items() if isinstance(v, str)} if isinstance(opts, dict) else {}
        return cls(kind=kind, label=str(data.get("label", "")), value=str(data.get("value", "")), options=opts)

    @property
    def display_value(self) -> str:
        if self.kind == KIND_SERMON:
            return " · ".join(v for v in (self.value, self.options.get("preacher", "")) if v) or "—"
        if self.kind == KIND_CCM and self.value:
            return ccm_title(self.value)
        return self.value if self.value else "—"

    @property
    def fill_values(self) -> list[str]:
        """템플릿의 {{내용1}}~{{내용N}} 에 들어갈 값 (빈 값은 입력하지 않음)."""
        if self.kind == KIND_SERMON:
            return [self.value, self.options.get("preacher", "")]
        if self.kind in (KIND_PRAYER, KIND_HYMN, KIND_CUSTOM):
            return [self.value]
        return []


def numbered_token(base: str, n: int) -> str:
    """'{{내용}}' → '{{내용1}}' (닫는 '}}' 앞에 번호). '}}' 로 끝나지 않으면 끝에 붙임."""
    if not base:
        return ""
    if base.endswith("}}") and len(base) > 2:
        return f"{base[:-2]}{n}}}}}"
    return f"{base}{n}"


def ccm_title(value: str) -> str:
    """CCM 상대 경로 'a/주님 다시 오실 때.pptx' → '주님 다시 오실 때' (화면 표시용)."""
    name = value.replace("\\", "/").rsplit("/", 1)[-1]
    stem, dot, ext = name.rpartition(".")
    return stem if dot and ext.lower() in ("pptx", "pptm", "ppt") else name


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
        label = label or KIND_LABELS[KIND_PRAYER]  # 기도 템플릿 이름 (합심기도, 중보기도 …)
    elif kind == KIND_HYMN:
        label = "찬송"
        value = normalize_hymn(value)
    elif kind == KIND_BIBLE:
        label = "말씀"
    elif kind == KIND_CCM:
        label = "CCM"
        value = value.replace("\\", "/")
    elif kind == KIND_SERMON:
        label = "설교"
        options = dict(options or {})
        options["preacher"] = str(options.get("preacher", "")).strip()
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
    if item.kind == KIND_SERMON:
        if not item.value:
            return "설교 제목을 입력해 주세요."
        if not item.options.get("preacher"):
            return "설교자를 입력해 주세요."
    if item.kind == KIND_CCM and not item.value:
        return "CCM 파일을 선택해 주세요."
    if item.kind == KIND_CUSTOM and not item.label:
        return "호출할 PPT 이름을 입력해 주세요."
    return None


def export_to_text(items: list[ServiceItem], church_name: str = "", worship_type: str = "", date_str: str = "",
                   filename_format: str = "") -> str:
    """예배 순서를 텍스트 형식으로 변환."""
    lines = []
    if church_name: lines.append(f"# 교회이름: {church_name}")
    if worship_type: lines.append(f"# 예배종류: {worship_type}")
    if date_str: lines.append(f"# 날짜: {date_str}")
    if filename_format: lines.append(f"# 파일이름형식: {filename_format}")
    lines.append("")

    for item in items:
        # [kind]label|value|역본(말씀) 또는 설교자(설교)
        opt_str = ""
        if item.kind == KIND_BIBLE and "bible" in item.options:
            opt_str = f"|{item.options['bible']}"
        elif item.kind == KIND_SERMON:
            opt_str = f"|{item.options.get('preacher', '')}"

        lines.append(f"[{item.kind}]{item.label}|{item.value}{opt_str}")

    return "\n".join(lines)


def import_from_text(text: str) -> tuple[list[ServiceItem], dict[str, str]]:
    """텍스트 형식을 읽어 ServiceItem 리스트와 메타 정보 반환."""
    items = []
    meta = {}

    for line in text.splitlines():
        line = line.strip()
        if not line: continue

        if line.startswith("#"):
            if ":" in line:
                k, v = line[1:].split(":", 1)
                meta[k.strip()] = v.strip()
            continue

        if line.startswith("[") and "]" in line:
            try:
                end_idx = line.find("]")
                kind = line[1:end_idx]
                rest = line[end_idx+1:].strip()

                if not rest: continue

                if kind not in KIND_LABELS:
                    kind = KIND_CUSTOM
                parts = rest.split("|", 2)
                label = parts[0].strip()
                value = parts[1].strip() if len(parts) > 1 else ""

                opts = {}
                if kind == KIND_BIBLE and len(parts) > 2 and parts[2].strip():
                    opts["bible"] = parts[2].strip()
                elif kind == KIND_SERMON and len(parts) > 2:
                    opts["preacher"] = parts[2].strip()

                # 기본 검증/정규화 적용 (검증 실패 항목은 건너뜀)
                item = make_item(kind, label, value, opts)
                if validate_item(item) is None:
                    items.append(item)
            except Exception:
                continue # 잘못된 라인은 건너뜀

    return items, meta
