# CLAUDE.md — 예배 순서 PPT 제작 프로그램 (ChurchPPTMaker)

> 이 파일은 프로젝트의 핵심 규칙/구조를 기록하는 기준 문서다. 기능·구조·규칙이 바뀌면 **반드시 함께 갱신**한다.

## 1. 프로젝트 개요 및 목표
- 교회 예배 PPT **전자동화** 데스크톱 프로그램 (Windows 전용).
- 예배 순서(기도·찬송·말씀·사용자 지정)를 UI에서 조립하면, 각 순서에 맞는 PPT 템플릿/찬송가 PPT/Bible2PPT 성경 PPT를
  찾아서 **원본 서식·폰트·애니메이션·레이아웃을 유지한 채 하나의 .pptx로 병합**한다.
- 최종 파일명: `YYYY-MM-DD_교회이름_예배종류.pptx` (요구사항 표기 `[YYYY-MM-DD]_[교회이름]_[예배종류].pptx`의 대괄호는 자리표시일 뿐 실제 파일명에는 넣지 않음). 출력 폴더에 저장.

## 2. 기술 스택
- Python 3.10+ (개발 환경: 3.14)
- **CustomTkinter 6.x** — UI (Win11 Fluent 느낌: 둥근 모서리, 액센트 `#0067C0`, 맑은 고딕, 다크/라이트/시스템 모드)
- **pywin32 (`win32com.client`)** — PowerPoint COM 자동화 (백그라운드, `WithWindow=False`)
- **Bible2PPT** (https://github.com/nazunamoe/Bible2PPT) — 외부 실행 파일을 CLI로 호출
- 실행: `pip install -r requirements.txt` → `python main.py`
- 빌드: `build.bat` (→ `build.ps1`: `.venv` 생성 → `requirements-build.txt` 설치 → PyInstaller `--onefile --windowed --collect-data customtkinter`) → `dist\ChurchPPTMaker.exe`
  - `build.ps1` 은 **UTF-8 BOM + CRLF** 로 유지 (Windows PowerShell 5.1 은 BOM 없는 파일을 ANSI로 읽어 한글이 깨짐).
  - 패키징 시 `config_manager` 가 `sys.frozen` 을 보고 exe 폴더를 프로그램 폴더로 사용.
  - 아이콘: `--icon assets/app.ico`(exe) + `--add-data assets/app.ico;assets`(창). `main.ICON_PATH` 는 `sys._MEIPASS` 기준.
    CTk 는 200ms 뒤 기본 아이콘으로 덮어쓰므로 생성 직후 `iconbitmap()` 호출 필수. 작업 표시줄 아이콘은 Tk 생성 전
    `SetCurrentProcessExplicitAppUserModelID` 로 python.exe 와 분리.

## 3. 디렉토리 구조 및 핵심 파일 역할
```
ChurchPPTMaker/
├─ CLAUDE.md           이 문서
├─ README.md           사용자용 소개·빌드·사용 방법 (기능/설정이 바뀌면 함께 갱신)
├─ requirements.txt    실행 패키지 / requirements-build.txt  빌드용(+PyInstaller)
├─ build.bat, build.ps1  exe 빌드 스크립트 (.venv 자동 생성)
├─ assets/app.ico     프로그램 아이콘 (창·작업 표시줄·exe, 16~256px 다중 크기)
├─ main.py             진입점 + 메인 윈도우(App), 패널 구성, 드래그&드롭, 프리셋/생성 이벤트
├─ ui_widgets.py       재사용 위젯: 공통 색/폰트, _Modal(창 내부 오버레이 다이얼로그), CalendarPopup, ItemDialog,
│                      BiblePicker(역본 검색 창)/BibleField(입력칸+🔍),
│                      SettingsDialog, LogDialog, DragController, 스레드→UI 큐(post_to_ui/run_in_thread)
├─ models.py           ServiceItem 데이터 모델 + 항목 종류(KIND_*) 정의, 입력 검증
├─ config_manager.py   config.json / presets.json 읽기·쓰기, 기본값 병합, 프리셋 규칙 검증
├─ bible_bridge.py     Bible2PPT CLI 래퍼 (구절 정규화/검증, 실행, --list-bibles)
├─ ppt_engine.py       win32com PowerPoint 엔진: 병합(원본 서식 유지), 텍스트 치환/입력
├─ worship_builder.py  생성 파이프라인: 항목→파일 해석, Bible2PPT 생성, 엔진 병합, 파일명 규칙
├─ config.json         사용자 설정 (첫 실행 시 자동 생성)
├─ presets.json        예배 순서 프리셋 (첫 실행 시 기본 3종 생성)
├─ templates/          순서별 PPT 템플릿 (기본 폴더, 설정에서 변경 가능)
├─ hymns/              찬송가 PPT 폴더 (기본, 하위 폴더까지 검색)
├─ output/             생성 결과 저장 폴더 (기본)
└─ data/               실제 예배 PPT를 순서별로 나눈 자료 (git 제외). templates/·hymns/ 하위 폴더, 순서.txt에 원본 순서 기록
```

## 4. UI 레이아웃 (최초 UI 초안 기준)
초안의 단일 창 구조를 **우측 "순서" 패널**로 그대로 옮기고, 추가 요구사항(폴더/파일/성경 폼)을 좌측 패널로 확장했다.
```
┌ 예배 순서 PPT 제작 프로그램 v1.0 ───────────────── [라이트|다크|시스템] [⚙ 설정] ┐
│ 교회 이름 [________]  날짜 [YYYY-MM-DD][📅][금일 지정]  예배 종류 [________]     │
├ 좌측 패널 ──────────────┬ 우측 패널 "순서" (초안 그대로) ─────────────────────┤
│                         │ 순서 (N개)   [파일 불러오기][파일 내보내기][전체 비우기] │
│ 폴더                    │ [프리셋 ▼(입력 가능)] [저장] [불러오기] [삭제]       │
│  템플릿/찬송가/출력 [📁]│ ┌──────────────────────────────────────────────┐ │
│ 템플릿 파일 (검색/목록) │ │ ⋮⋮ [예배로 부름]────────[인도자]   ▲ ▼ ✕      │ │
│  → 드래그/더블클릭 추가 │ │ ⋮⋮ [기도]──────────────[OOO장로]  ▲ ▼ ✕      │ │
│ 성경 구절 추가          │ │ ⋮⋮ [찬송]──────────────[29장]     ▲ ▼ ✕      │ │
│  (Bible2PPT 연동)       │ │ ⋮⋮ [말씀]──────────────[요3:16]   ▲ ▼ ✕      │ │
│  [구절][역본▼][추가]    │ └──────────────────────────────────────────────┘ │
│                         │ [기도] [찬송] [말씀] [사용자 지정]  ← 드래그&드롭   │
├─────────────────────────┴───────────────────────────────────────────────────┤
│ 상태 메시지 / 진행바  [로그] [✓ 생성 후 열기] [☐ 슬라이드 쇼로 시작] [예배 PPT 병합 및 생성] │
└─────────────────────────────────────────────────────────────────────────────┘
```
- 날짜: 입력칸 + 📅 아이콘(달력 팝업, 날짜 클릭 시 입력, 다시 누르거나 바깥 클릭/Esc 시 닫힘) + [금일 지정](오늘 날짜).
- 다이얼로그(항목 입력/설정/로그)와 달력은 **별도 창이 아니라 메인 창 안의 오버레이**(Win11 ContentDialog 스타일). Enter=확인, Esc=취소.
- 좌측 템플릿 파일 목록: 클릭=선택, 더블클릭=맨 아래 추가, 순서 목록으로 드래그=해당 위치 삽입(사용자 지정 항목).
- 순서 행: `⋮⋮` 손잡이로 드래그 재정렬, ▲▼ 이동, ✕ 개별 삭제, **오른쪽 값(알약) 클릭 시 수정**.
- 하단 팔레트 버튼(기도/찬송/말씀/사용자 지정)을 순서 목록으로 **드래그&드롭**하면 놓은 위치에 삽입, 단순 클릭은 맨 아래 추가.
  - 추가 시 입력 다이얼로그: 기도→기도자 이름, 찬송→찬송가 장, 말씀→구절 (**미입력 시 에러**)
  - 사용자 지정→호출할 PPT 이름(필수, 미입력 시 에러) + 오른쪽 문구(**선택**)

## 5. 순서 항목 종류와 PPT 호출 규칙 (models.py / worship_builder.py)
| kind     | 왼쪽(label)      | 오른쪽(value)  | PPT 호출 규칙 |
|----------|------------------|----------------|---------------|
| `custom` | PPT 이름 (예: 예배로 부름) | 커스텀 문구(선택) | 템플릿 폴더에서 **왼쪽 이름**의 PPT. 문구가 있으면 텍스트 입력 |
| `prayer` | 기도             | 기도자 이름(필수) | 템플릿 폴더의 `기도.pptx` 호출 + 지정 텍스트 박스에 이름 입력 |
| `hymn`   | 찬송             | `29장`(필수)   | 찬송가 폴더(하위 포함)에서 **오른쪽 장 번호**의 PPT 검색 |
| `bible`  | 말씀             | `요3:16`(필수) | Bible2PPT CLI로 구절 PPT 생성 후 병합 |

- 템플릿 파일 매칭: 파일명(확장자 제외)이 label과 같거나, 공백 제거 후 같으면 일치. 확장자 `.pptx/.pptm/.ppt`.
- 찬송가 매칭 우선순위: ① 파일명이 `29`, `029`, `29장`, `029장` ② 파일명의 **첫 번째 숫자**가 29 ③ (숫자 없는 값) 파일명에 값 포함.
- 텍스트 입력 규칙("지정 텍스트 박스"):
  1. 해당 항목 슬라이드에서 치환 토큰(기본 `{{내용}}`, 설정 가능)을 찾아 value로 치환 (서식 유지, `TextRange.Replace`).
  2. 토큰이 없으면 이름이 `내용`(설정 가능)인 도형의 텍스트를 value로 교체.
  3. 둘 다 없으면 경고 로그만 남기고 진행.
- 전체 공통 토큰(모든 슬라이드): `{{교회이름}}`, `{{날짜}}`, `{{예배종류}}`.

## 6. Bible2PPT 연동 규칙 (bible_bridge.py)
- 원본 C# 프로그램의 CLI를 subprocess로 호출 (창 숨김, UTF-8 stdout).
```
Bible2PPT.exe <성경 구절...> [옵션]
  -o, --output <path>    저장 경로 (--split이면 폴더). 생략 시 임시 파일+파워포인트로 열림 → 본 프로그램은 항상 지정
  -b, --bible <bible>    성경 ID 또는 이름 (여러 번 지정 시 나란히 배치). 생략 시 마지막 선택 성경
  -t, --template <path>  템플릿 .pptx (생략 시 프로그램 템플릿)
  -l, --lines <0-9>      슬라이드당 구절 줄 수 (0: 제한 없음)
  -s, --split            장별로 PPT 나누기 (본 프로그램은 사용하지 않음)
  --open                 완료 후 열기 (사용하지 않음)
  --list-bibles          사용 가능한 성경 ID/이름 출력
종료 코드: 0 성공 / 1 인자 오류 / 2 PPT 생성 실패 / 3 파워포인트 초기화 실패
성공 시 저장 경로를 stdout으로 출력
```
- 구절 문법: `책약자[시작장[:시작절][-[끝장[:끝절]|끝절]]]` 예) `창`, `창1`, `롬1-3`, `전1:3`, `창1:1-5`, `사1:3-3:9`
- 입력 정규화: 쉼표/세미콜론으로 여러 구절 구분, 각 구절 내부 공백 제거(`요 3:16` → `요3:16`), 전각 콜론 `：` → `:`.
- 역본은 항목별(`options.bible`) → 없으면 설정 기본값 → 없으면 Bible2PPT 마지막 선택값.
- `--list-bibles` 출력(포크 `CommandLine.cs` 확인): 헤더 `ID\t소스\t성경` + 줄마다 `ID\t소스\t이름`, 리다이렉트 시 UTF-8.
  소스별 실패는 stderr 로만 알림(종료 코드 0). → `parse_bible_list` 가 `이름 · 소스 [ID]` 로 표시, `bible_arg` 가 `[ID]` 를 `-b` 로 전달.
- 역본 선택은 모두 `BibleField`(직접 입력 + 🔍 → `BiblePicker`): 메인 성경 영역, 말씀 항목 입력, 설정 기본 역본.
  역본이 수백 개(YouVersion 10개 언어만 약 180개)라 드롭다운 대신 검색 창. 검색 로직은 `bible_bridge.filter_bibles`
  (공백 구분 AND, 이름·소스·ID, 대소문자/공백 무시, 초성 `ㄱㅇㄱㅈ`, 이름 접두 일치 우선). 목록은 tk.Listbox (CTk 위젯 수백 개는 느림).
- 역본 목록은 `config.bible2ppt.bible_list` 에 캐시. 설정에서 exe 지정(찾기/입력 후 포커스 이동) 시 즉시 불러오고,
  시작 시에는 exe 가 있고 캐시가 비었을 때만 조용히 불러옴(매 실행마다 Bible2PPT·PowerPoint 를 띄우지 않기 위함).
- 원본 Bible2PPT 2.0.4(GUI 전용)는 인자를 무시하고 창을 띄운 채 멈춘다(실측: 180초 타임아웃).
  → `supports_cli()` 가 실행 전에 .NET Framework exe 의 UTF-16 문자열 `--list-bibles` 유무로 판별해 즉시 안내.
- Bible2PPT는 PowerPoint를 사용하므로 **병합 COM 작업 시작 전에 순차 실행**한다.

## 7. PPT 엔진 규칙 (ppt_engine.py)
- `win32com.client.Dispatch("PowerPoint.Application")`, 이미 실행 중이었으면 종료하지 않음.
- 베이스: 첫 원본을 `SaveAs(출력, 24=pptx)` (원본 100% 유지). 이후 원본은 `config.ppt.merge_mode` 에 따라 이어 붙임:
  - `source_formatting` (기본): `CommandBars.ExecuteMso("PasteSourceFormatting")` = PowerPoint의 [원본 서식 유지] 붙여넣기.
    테마·마스터·레이아웃·**슬라이드별 배경(그림 포함)**까지 보존. 문서 창이 필요해서 `dest.NewWindow()` 후 PowerPoint를 **최소화**하고,
    이미 실행 중이던 PowerPoint면 작업 후 창 상태를 복원. 비동기라 슬라이드 수 증가를 폴링으로 확인, 실패 시 해당 파일만 아래 방식으로 대체.
  - `windowless`: `Slides.Paste` → 슬라이드별 `Design`(파일+디자인명 캐시)·`CustomLayout` 원본 지정 → 배경 복제
    (단색은 RGB 정확 복사, 그라데이션/그림 등은 원본 도형을 잠시 숨기고 2560px PNG로 내보내 `UserPicture`). 클립보드 실패 시 `InsertFromFile`.
  - ⚠ 그냥 `Slides.Paste`/`InsertFromFile` 만 쓰면 **슬라이드별 배경이 사라진다** (대상 테마 사용과 동일) — 2026-10-06 실험으로 확인.
- 애니메이션/전환/폰트는 슬라이드 복사 시 그대로 따라온다. 원본 파일은 절대 수정하지 않는다(ReadOnly 열기).
- UI 스레드를 막지 않도록 작업 스레드에서 `pythoncom.CoInitialize()` 후 실행.

## 8. 사용자 설정 (config.json)
```json
{
  "church_name": "", "worship_type": "주일오전예배", "appearance_mode": "System",
  "paths": {"template_dir": "templates", "hymn_dir": "hymns", "output_dir": "output"},
  "bible2ppt": {"exe_path": "", "bible": "", "template": "", "lines": null, "timeout_sec": 180,
                "bible_list": []},
  "ppt": {"text_placeholder": "{{내용}}", "text_shape_name": "내용", "open_after_build": true, "start_slideshow": false,
          "merge_mode": "source_formatting"}
}
```
- `start_slideshow`: 생성 후 열 때 `os.startfile(path, "show")`(= .pptx 의 Show 동사, `POWERPNT.EXE /s`)로 바로 쇼 시작.
  Show 동사가 없으면 App Paths 의 POWERPNT.EXE `/s`, 그래도 실패하면 일반 열기 + 로그 경고. `open_after_build` 가 꺼지면 비활성.
- 상대 경로는 프로그램 폴더 기준. 누락 키는 기본값으로 자동 병합. 교회 이름/예배 종류/폴더는 메인 UI에서 바로 저장됨, 나머지는 ⚙ 설정.

## 9. 프리셋 규칙 (presets.json)
- 형식: `{"presets": [{"name": "주일오전예배", "items": [{"kind","label","value","options"}...]}]}`
- **최대 10개**, 이름 **1~10글자**, 같은 이름 저장 시 덮어쓰기 확인, **삭제 전 확인 창** 필수.
- 불러오기 시 순서 목록을 교체하고 예배 종류를 프리셋 이름으로 설정.
- 기본 프리셋: 주일오전예배, 주일오후예배, 수요예배 (최초 실행 시 생성).
- 텍스트 파일 입출력(순서 제목 줄 [파일 불러오기]/[파일 내보내기], `models.export_to_text`/`import_from_text`, UTF-8 BOM):
  `# 교회이름: …`/`# 예배종류: …`/`# 날짜: …` 메타 줄 + 항목 줄 `[kind]label|value[|역본]`.
  불러올 때 모르는 kind 는 custom, `validate_item` 실패 줄은 건너뜀.

## 10. 개발 히스토리
- 2026-10-06 v1.0: UI 초안 기반 최초 구현 (UI, 프리셋, 드래그&드롭, Bible2PPT CLI 연동, COM 병합 엔진).
  - 병합 시 슬라이드 배경 소실 발견 → `PasteSourceFormatting` 기본 + windowless 배경 복제 대체 경로 추가.
  - CTkToplevel 반복 생성/파괴 시 Tk 접근 위반(크래시) 재현 → 모든 다이얼로그를 창 내부 오버레이(_Modal)로 전환, 12회 반복 테스트 무크래시.
  - 검증: 실제 PowerPoint 16.0으로 7개 항목 병합(텍스트 치환·애니메이션·단색/그림 배경 유지, 원본 미수정), UI 자동 테스트 25항목.
    Bible2PPT는 실제 exe 없이 가짜 생성기로만 검증됨 → 실제 exe 연동 확인 필요.

## 11. 구현 시 주의사항 (검증된 함정)
- **CTkToplevel 사용 금지** → `ui_widgets._Modal`(창 내부 오버레이) 상속. CTkToplevel 은 생성 때마다 `bind_all("<Button-1>")`를 누적하고
  창 숨김·강제 update·ctypes 호출을 하며, 반복 사용 시 접근 위반 크래시가 났다.
- tkinter 위젯 서브클래스에 `self._root` 속성을 만들지 말 것 (`Misc._root()` 메서드를 덮어써 이벤트 처리가 깨짐). → `self._win` 사용.
- `CTkFrame` 은 높이 2px 이하를 0으로 그려 보이지 않는다 → 구분선/연결선은 `height=3` 이상.
- CTk 의 `place(x=, y=)` 는 좌표에 배율을 곱한다 → 실제 픽셀 좌표는 `ScalingTracker.get_widget_scaling` 으로 나눠 전달.
- 폰트는 `ui_widgets.font()` 캐시만 사용 (렌더링마다 CTkFont 생성 금지).
- 작업 스레드에서 tkinter 호출 금지 → `post_to_ui()` / `run_in_thread()`.
- `_Modal` 은 중첩 가능(예: 항목 입력 → 역본 검색). Esc/Enter 는 창 단위 바인딩이라 `_Modal._stack` 맨 위만 반응하고,
  닫힐 때 아래 다이얼로그 grab 을 복구한다. 바인딩 해제는 `_unbind()` 사용 — Python 3.12 이하 `Misc.unbind(seq, funcid)` 는
  seq 의 바인딩을 **전부** 지워 아래 다이얼로그의 Esc/Enter 가 사라진다.
- `PptEngine` 은 내부에서 CoInitialize/CoUninitialize 하므로 반드시 작업 스레드에서 `with` 로 사용.

## 11-1. 보안 규칙 (2026-10-07 점검)
- COM 으로 연 파일은 **보호된 보기 없이** 열리고, 자동화 기본값(`AutomationSecurity=1`)은 매크로를 허용한다(실측).
  → `PptEngine` 이 작업 동안 `AutomationSecurity=3`(강제 비활성)으로 바꾸고, `DisplayAlerts` 와 함께 종료 시 원래 값으로 복원.
- 외부 프로그램 실행은 `subprocess` 리스트 인자만(`shell=True` 금지). Bible2PPT 경로는 **`.exe` 만 허용**
  (.bat/.cmd 는 cmd.exe 가 인자를 재해석해 `&` 등으로 명령 주입 가능). `-b` 값은 `safe_bible_arg` 로 `-` 시작·제어 문자 거부.
  구절은 `_REF_RE` 로 검증돼 옵션처럼 시작할 수 없다.
- Bible2PPT stdout 이 알려준 경로는 지정한 출력 폴더 안의 파일만 인정.
- 출력 파일명은 `sanitize_filename`(경로 구분자·제어 문자 제거, CON/NUL 등 예약어 회피)으로만 만들고,
  결과 경로가 원본 PPT 와 같으면 생성 거부(원본 덮어쓰기 방지).
- config.json / presets.json 은 외부에서 수정될 수 있는 입력으로 취급: 기본값과 타입이 다른 값은 무시, 잘못된 프리셋/항목은 버림,
  `options` 는 문자열 값만 유지.

## 12. 작업 규칙
- UI 배치를 바꿀 때는 §4 초안 구조(순서 패널 = 프리셋 행 / 순서 목록 / 팔레트 행)를 유지한다.
- 사용자 메시지는 한국어. 모든 입력 검증 실패는 `messagebox.showerror`로 알린다.
- COM 호출은 반드시 작업 스레드에서, UI 갱신은 `after()`로 메인 스레드에서.
