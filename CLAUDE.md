# CLAUDE.md — 예배 순서 PPT 제작 프로그램 (ChurchPPTMaker)

> 이 파일은 프로젝트의 핵심 규칙/구조를 기록하는 기준 문서다. 기능·구조·규칙이 바뀌면 **반드시 함께 갱신**한다.

## 1. 프로젝트 개요 및 목표
- 교회 예배 PPT **전자동화** 데스크톱 프로그램 (Windows 전용).
- 예배 순서(기도·찬송·말씀·설교·CCM·사용자 지정)를 UI에서 조립하면, 각 순서에 맞는 PPT 템플릿/찬송가 PPT/Bible2PPT 성경 PPT를
  찾아서 **원본 서식·폰트·애니메이션·레이아웃을 유지한 채 하나의 .pptx로 병합**한다.
- 최종 파일명: 기본 `YYYY-MM-DD_교회이름_예배종류.pptx` (요구사항 표기 `[YYYY-MM-DD]_[교회이름]_[예배종류].pptx`의 대괄호는 자리표시일 뿐 실제 파일명에는 넣지 않음). 출력 폴더에 저장.
  상단 [파일 이름] 형식으로 바꿀 수 있고 프리셋별로 저장된다 (§9-1).

## 2. 기술 스택
- Python 3.10+ (개발 환경: 3.14)
- **CustomTkinter 6.x** — UI (Win11 Fluent 느낌: 둥근 모서리, 액센트 = Windows 강조 색(끄면 `#0067C0`), 맑은 고딕, 다크/라이트/시스템 모드)
- **pywin32 (`win32com.client`)** — PowerPoint COM 자동화 (백그라운드, `WithWindow=False`)
- **Bible2PPT** (https://github.com/nazunamoe/Bible2PPT) — 외부 실행 파일을 CLI로 호출
- 실행: `pip install -r requirements.txt` → `python main.py`
- 빌드: `build.bat` (→ `build.ps1`: `.venv` 생성 → `requirements-build.txt` 설치 → PyInstaller `--onefile --windowed --collect-data customtkinter`) → `dist\ChurchPPTMaker.exe`
  - **버전**: `version.py` 의 `__version__`(주.부.수정)이 유일한 기준. `main.APP_TITLE` 이 `APP_NAME v버전` 으로 표시하고,
    `build.ps1` 이 `python version.py --write-info build\version_info.txt` → PyInstaller `--version-file` 로 exe 파일/제품 버전에 반영.
    `build.bat -Version 1.2.0` 은 `version.py --set` 으로 파일을 고친 뒤 빌드. 버전을 화면·문서에 하드코딩하지 말 것.
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
├─ window_frame.py     Windows 제목 표시줄을 없애는 창 틀 처리(CustomFrame: WM_NCCALCSIZE 서브클래스, 창 옮기기·위쪽 크기 조절)
├─ help_content.py     프로그램 안 도움말 내용 (HELP_SECTIONS: 절 제목 + 간단한 표기). 사용법이 바뀌면 README 와 함께 갱신
├─ requirements.txt    실행 패키지 / requirements-build.txt  빌드용(+PyInstaller)
├─ build.bat, build.ps1  exe 빌드 스크립트 (.venv 자동 생성, -Version 으로 버전 변경)
├─ version.py          프로그램 버전·이름 (화면 제목과 exe 버전 정보의 기준)
├─ assets/app.ico     프로그램 아이콘 (창·작업 표시줄·exe, 16~256px 다중 크기)
├─ main.py             진입점 + 메인 윈도우(App), 패널 구성, 드래그&드롭, 프리셋/생성 이벤트
├─ ui_widgets.py       재사용 위젯: 공통 색/폰트, _Modal(창 내부 오버레이 다이얼로그), CalendarPopup, ItemDialog,
│                      BiblePicker(역본 검색 창)/BibleField(입력칸+🔍),
│                      SettingsDialog(한 쪽 세로 스크롤), SetupWizard(설정 마법사), BuildOverlay(생성 중 화면), HelpDialog, TempFilesDialog, LogDialog, DragController, 스레드→UI 큐(post_to_ui/run_in_thread)
├─ models.py           ServiceItem 데이터 모델 + 항목 종류(KIND_*) 정의, 입력 검증
├─ config_manager.py   config.json / presets.json 읽기·쓰기, 기본값 병합, 프리셋 규칙 검증
├─ bible_bridge.py     Bible2PPT CLI 래퍼 (구절 정규화/검증, 실행, --list-bibles)
├─ ppt_engine.py       win32com PowerPoint 엔진: 병합(원본 서식 유지), 텍스트 치환/입력
├─ worship_builder.py  생성 파이프라인: 항목→파일 해석, Bible2PPT 생성, 엔진 병합, 파일명 규칙
├─ config.json         사용자 설정 (첫 실행 시 자동 생성)
├─ presets.json        예배 순서 프리셋 (첫 실행 시 기본 3종 생성)
├─ templates/          순서별 PPT 템플릿 (기본 폴더, 설정에서 변경 가능)
├─ hymns/              찬송가 PPT 폴더 (기본, 하위 폴더까지 검색)
├─ ccm/                CCM PPT 폴더 (기본, 하위 폴더까지 검색, 파일명 자유 → 검색해서 선택)
├─ prayers/            기도 템플릿 폴더 (기본, 합심기도.pptx 등 — 기도 추가 시 하나 선택)
├─ output/             생성 결과 저장 폴더 (기본)
└─ data/               실제 예배 PPT를 순서별로 나눈 자료 (git 제외). templates/·hymns/ 하위 폴더, 순서.txt에 원본 순서 기록
```

## 4. UI 레이아웃 (최초 UI 초안 기준)
초안의 단일 창 구조를 **우측 "순서" 패널**로 그대로 옮기고, 추가 요구사항(폴더/파일/성경 폼)을 좌측 패널로 확장했다.
```
┌ 예배 순서 PPT 제작 프로그램 v1.0.0 [? 도움말][설정 마법사] ─ [라이트|다크|시스템] [⚙ 설정] ─ □ ✕┐
│ 교회 이름 [________]  날짜 [YYYY-MM-DD][📅][금일 지정]  예배 종류 [________]     │
│ 출력 폴더 [____][📁][↗]  파일 이름 [{날짜}_{교회이름}_{예배종류}][항목 ▾][기본값]      │
│                                    → 2026-10-11_사랑교회_주일오전예배.pptx          │
├ 좌측 패널 ──────────────┬ 우측 패널 "순서" (초안 그대로) ─────────────────────┤
│                         │ 순서 (N개)   [파일 불러오기][파일 내보내기][전체 비우기] │
│ 폴더                    │ [프리셋 ▼(입력 가능)] [저장] [불러오기] [삭제]       │
│  기도/찬송가/CCM        │ ┌──────────────────────────────────────────────┐ │
│ 템플릿 파일 (폴더/검색) │ │ ⋮⋮ [예배로 부름]────────[인도자]   ▲ ▼ ✕      │ │
│  → 드래그/더블클릭 추가 │ │ ⋮⋮ [기도]──────────────[OOO장로]  ▲ ▼ ✕      │ │
│ 성경 구절 추가 [↻][추가]│ │ ⋮⋮ [찬송]──────────────[29장]     ▲ ▼ ✕      │ │
│  [구절]                 │ │ ⋮⋮ [CCM]───────────────────────   ▲ ▼ ✕      │ │
│  기본 성경 역본 [  ][🔍] │ │        [아주 긴 곡 제목 … (아랫줄)]           │ │
│  두 번째 성경 역본[  ][🔍]│ └──────────────────────────────────────────────┘ │
│                         │ [기도][찬송][말씀][설교][CCM][사용자 지정] ← 드래그&드롭 │
├─────────────────────────┴───────────────────────────────────────────────────┤
│ 상태 메시지 / 진행바  [로그]     [✓ 생성 후 열기] [☐ 슬라이드 쇼로 시작] [예배 PPT 병합 및 생성] │
│ 부연 설명(최대 2줄)  [임시 파일] 슬라이드 비율 [원본|16:9|4:3]    [☐ 저장 없이 임시 파일로 실행] │
└─────────────────────────────────────────────────────────────────────────────┘
```
- 파일 이름: 결과 파일 이름 형식 입력칸 + [항목 ▾](tk.Menu 로 커서 위치에 `{항목}` 삽입) + [기본값] + 실시간 미리보기(잘못되면 ⚠).
  상단 카드는 두 줄 틀(line1/line2)을 카드 안쪽으로 들여 배치 — 투명 틀이 카드 색으로 칠해져 둥근 테두리를 덮기 때문.
- 하단 상태: `_set_status(text, detail)` — text 는 상태 한 줄, detail(저장 방법·경고 수 등 부연)은 진행 막대 아랫줄
  (`status_detail`, 칸 너비로 줄 바꿈 — 하단 바 높이는 오른쪽 버튼 두 줄이 정해서 두 줄까지는 바가 커지지 않음).
  한 줄에 이어 쓰면 최소 창 폭에서 뒷부분이 잘렸음.
- 성경 구절 추가: 제목 줄 오른쪽에 [↻](역본 목록 다시 불러오기)·[순서에 추가], 아래 역본 두 칸 왼쪽에 '기본 성경 역본'·'두 번째 성경 역본'.
- 창 틀(`window_frame.CustomFrame`): Windows 제목 표시줄 없이 프로그램 제목 줄이 맨 위. 오른쪽 위 끝에 ─ □ ✕
  (`CAPTION_GLYPHS`, Segoe MDL2 Assets — 최대화면 □ 가 복원 모양으로). 제목 줄 빈 곳·제목 글자를 끌면 창 옮기기(스냅 포함),
  두 번 누르면 최대화/복원(옮기기 루프에 들어가면 Tk 가 더블클릭을 못 받아 `_on_title_press` 가 시간·위치로 직접 판정),
  맨 위 `TOP_RESIZE_EDGE`(5) 안은 위쪽 크기 조절. 좌우·아래 크기 조절·둥근 모서리·그림자·작업 표시줄·Alt+Tab 은 Windows 그대로.
  창 틀 처리는 처음 표시될 때(아직 투명할 때 `_on_window_configure`) 적용 → 깜빡임 없음.
  다이얼로그가 떠 있으면 grab 때문에 제목 줄로 창을 옮길 수 없음(Esc 로 닫은 뒤). Win11 최대화 버튼 위 스냅 레이아웃 메뉴는 없음.
  Win10: `GetDpiForWindow`·`GetSystemMetricsForDpi` 는 1607 이상에만 있어 없으면 시스템 배율 값(`_HAS_DPI_API`)으로 대신하고,
  창 틀 처리(`install`)가 실패하면 Windows 기본 제목 표시줄 그대로 + 프로그램 안 창 버튼 숨김. 실제 Win10 에서는 아직 확인 안 함
  (이 PC 에서 함수 없음·처리 실패를 흉내 내 확인). Win10 은 모서리가 각지고, 위쪽 틀이 없어 위 테두리 선이 안 보일 수 있음.
- 도움말(`HelpDialog`, 제목 오른쪽 [? 도움말]): 왼쪽 목차(tk.Listbox, 클릭하면 그 절로) + 본문 Text 하나(태그로 제목·표·
  글머리·강조) + 위 검색칸(입력하면 모두 표시, Enter 다음 찾기). 내용은 `help_content.HELP_SECTIONS` — 표기는 그 파일 머리말.
  CTkTextbox 는 태그 font 를 막으므로 안쪽 `_textbox` 에 `tk_font` 로 지정. 표는 열마다 가장 긴 글자로 탭 위치를 정함
  (열 최대 300 논리 px, 넘는 칸은 다음 열을 밀어냄 → 첫 열들은 짧게 쓸 것). 높이는 설정 창과 같은 이유로 창 높이×0.9 기준.
- 날짜: 입력칸 + 📅 아이콘(달력 팝업, 날짜 클릭 시 입력, 다시 누르거나 바깥 클릭/Esc 시 닫힘) + [금일 지정](오늘 날짜).
- 다이얼로그(항목 입력/설정/로그)와 달력은 **별도 창이 아니라 메인 창 안의 오버레이**(Win11 ContentDialog 스타일). Enter=확인, Esc=취소.
- 생성 중 화면(`ui_widgets.BuildOverlay`, `App._show_build_overlay`/`_hide_build_overlay`): 생성하는 동안 `_content` 1~3행
  (상단 입력·본문·하단 바)을 grid 로 겹쳐 덮고 진행 단계·막대·경과 시간 + "PowerPoint 를 사용하지 마세요" 경고(§9-2 활성 창 문제)를 보인다.
  `_Modal` 과 달리 grab 을 걸지 않음 — 제목 줄(창 옮기기·최소화·닫기)은 생성 중에도 써야 하므로 덮기만 하고, [⚙ 설정]·
  [설정 마법사]만 잠근다. 덮기만 하면 키보드 포커스가 덮인 입력칸에 남아 글자가 입력되므로 화면에 포커스를 가져온다.
  `_set_progress` 가 하단 진행 막대와 함께 갱신, `done` 에서 결과 창·에러 창보다 먼저 닫는다.
  **멈춤 안내**: 진행이 `STALL_SEC`(60초, '성경 PPT 생성 중' 단계는 Bible2PPT 제한 시간 + 30초) 동안 없으면 경고 아래에 안내 +
  [PowerPoint 창 보기](`ppt_engine.show_powerpoint_dialogs` — POWERPNT 프로세스의 보이는 창 중 문서 창이 아닌 것을 앞으로)·
  [생성 중단](`App._cancel_build` → `kill_powerpoint`). COM 호출은 중간에 끊을 수 없어(처음 설치한 Office 의 라이선스 동의 창,
  암호 창 등을 PowerPoint 가 기다리면 끝나지 않음) 프로세스를 끄는 것뿐 — 끄면 멈춘 호출이 오류로 끝나 정상 정리된다(실측).
  생성 시작 때 `powerpoint_pids()` 를 기억해 **그 뒤에 켜진 PowerPoint 만** 끈다(PowerPoint 는 프로세스 하나라 사용자가 켜 둔
  것이면 사용자 문서가 같이 사라짐). 중단한 생성은 실패 창 대신 상태 줄에 '생성을 중단했습니다'.
- 설정 마법사(`ui_widgets.SetupWizard`, 제목 [? 도움말] 옆 [설정 마법사] → `App._open_setup_wizard`): 폴더(템플릿·기도·
  찬송가·CCM) → 설교 템플릿 → 말씀(Bible2PPT.exe·역본1·2·성경 템플릿·줄 수) → 확인(요약, 빠진 것 ⚠) 4단계 `_Modal`.
  [다음]마다 그 단계 검사(`_check_step` — 폴더는 `worship_builder.folder_conflict` + 바꾼 값을 넣은 설정 사본 `_probe_cfg` 로
  `output_conflict`, 파일은 ⚙ 설정과 같은 규칙), **[마침]에서만 저장**. 폴더 칸에 처음 보인 값(절대 경로) 그대로면 원래 값
  (기본 `templates` 같은 상대 경로)을 유지. 역본 목록 불러오기는 ⚙ 설정과 `_BibleListMixin` 공용.
  단계 쪽은 한 칸에 모두 grid 해 두고 지금 쪽만 보임 — 칸 minsize = 가장 큰 쪽 + 바깥 여백(여백을 빼면 큰 쪽에서 다시 커짐).
  본문은 CTkScrollableFrame — 높이 = min(내용, 창 높이×0.9 − 나머지)(`_fit_body`, 1366×768·125% 처럼 창이 약 530 이면
  마법사(574)의 위가 잘렸음). '나머지'는 다이얼로그 − **안쪽 캔버스** 높이로 잴 것(바깥 틀 기준이면 테두리 여백 약 21 만큼 넘침).
  PPT 개수는 `worship_builder.count_ppt_files`(0.3초 제한, 넘으면 None → "폴더가 너무 커서…") — 사용자 폴더를 찬송가
  폴더로 고르면 하위까지 세느라 19초 멈췄음(실측). 경로 칸은 `path_text` 로 앞뒤 따옴표를 뗌(탐색기 [경로로 복사]) — ⚙ 설정도 같음.
  처음 실행(`config.setup_done` False)이면 `_reveal` 뒤 자동으로 띄우고, 닫는 방법과 상관없이 `setup_done` True.
  `ConfigManager` 는 설정 파일은 있는데 키가 없으면(마법사 이전 버전 사용자) True 로 둔다. 생성 중에는 버튼을 잠근다.
- 좌측 폴더 줄(`App._folder_row`: 이름·경로·📁·↗): [폴더] 카드에 기도/찬송가/CCM, 템플릿 폴더는 그 파일을 보여 주는
  [템플릿 파일] 카드 맨 위(검색칸 위)에 둔다. 출력 폴더는 상단 둘째 줄의 파일 이름 형식 왼쪽에 둔다.
  상단 둘째 줄은 grid: 파일 이름 미리보기는 형식 입력칸 **아랫줄**(같은 열). 미리보기 길이가 교회 이름·예배 종류에 따라
  늘어 같은 줄에 두면 상단 카드 요청 너비가 1013→1248 로 창보다 넓어졌음(실측) → 아랫줄로 내리고, 남는 너비보다 길면
  가운데를 '…' 로 줄임(`_fit_filename_preview`, 줄의 `<Configure>` 마다 다시 맞춤). 지금 상단 카드 요청 너비는 913 고정.
  `App.path_vars`(폴더 키 → 표시 경로)는 상단·좌측 양쪽에서 쓰므로 `__init__` 에서 만든다.
- 좌측 템플릿 파일 목록: 클릭=선택, 더블클릭=맨 아래 추가, 순서 목록으로 드래그=해당 위치 삽입(사용자 지정 항목).
  아래 버튼 [불러오기](다른 곳의 PPT 를 템플릿 폴더로 **복사**, 같은 이름은 덮어쓰기 확인) [열기](PowerPoint 로 편집)
  [삭제](확인 후 **휴지통**으로 이동 — `main._send_to_recycle_bin`, SHFileOperationW+FOF_ALLOWUNDO) [순서에 추가].
  버튼 4개가 들어가도록 '끌어다 놓기 · 더블클릭으로 추가' 안내는 카드 제목 줄로 옮김.
  글자 크기 `main.FILE_FONT_SIZE`(13.2 = 기본 12의 110%, 줄 높이도 따라 커짐).
- 순서 행: `⋮⋮` 손잡이로 드래그 재정렬, ▲▼ 이동, ✕ 개별 삭제, **알약 클릭 시 수정, 더블클릭 시 원본 PPT 열기**
  (`App._open_item`: 기도·설교·사용자 지정=템플릿, 찬송·CCM=그 PPT — 파일 찾기는 생성과 같은 `WorshipBuilder.resolve_item`,
  말씀=`generate_bible` 로 `%TEMP%\ChurchPPTMaker_말씀` 에 만들어 열기, 하루 지난 파일은 정리).
  CTkButton 은 클릭마다 command 를 부르므로 `_on_pill_click` 이 Windows 더블클릭 간격(최대 500ms)만큼 수정 창을 미뤘다가
  같은 행이 다시 눌리면 열기로 바꾼다. 행 빈 곳·연결선은 `<Double-Button-1>`. 목록을 다시 그리면 예약 취소.
  말씀 열기(Bible2PPT)와 생성은 둘 다 PowerPoint 를 쓰므로 동시에 하지 않음(`_bible_preview`/`building`).
- 순서 행 줄 바꿈(`App._layout_rows`): 행은 grid `[손잡이][왼쪽 알약][연결선][오른쪽 알약][▲▼✕]`. 한 줄에 안 들어가면
  오른쪽 알약을 아랫줄(columnspan 1~3, 오른쪽 정렬 — 다른 행의 오른쪽 알약과 끝을 맞춤)로 내리고, 그래도 넘치면 글자를
  '…' 로 줄임(`_fit_pill`, 알약 여백 = 전체 글자일 때 reqwidth − 글자 폭). 너비는 처음 그릴 때 한 번 재 두고, 목록 캔버스
  `<Configure>` 마다 grid 만 바꿈(위젯을 다시 만들지 않음). 왼쪽 알약도 너무 길면 줄임.
- 순서 목록 스크롤: CTk 는 Windows 에서 휠 한 칸 = 20단위 × `yscrollincrement`(기본 1px)라 느렸음 →
  `ORDER_SCROLL_UNIT`(2.4 논리 px, 배율 반영)로 한 칸 ≈ 행 하나. 배율이 바뀌면 다시 지정. 설정 창 스크롤도 같은 값.
- 하단 팔레트 버튼(기도/찬송/말씀/설교/CCM/사용자 지정)을 순서 목록으로 **드래그&드롭**하면 놓은 위치에 삽입, 단순 클릭은 맨 아래 추가.
  - 추가 시 입력 다이얼로그: 기도→기도 종류(기도 폴더 PPT 중 선택)+기도자 이름, 찬송→찬송가 장, 말씀→구절, 설교→설교 제목+설교자 (**미입력 시 에러**)
  - CCM→CCM 폴더 파일 검색(검색칸 + tk.Listbox, `bible_bridge.search_names`: 공백 AND·초성) 후 선택 (**미선택 시 에러**)
  - 팔레트 버튼 6개 폭 합계는 최소 창 폭의 순서 패널(약 570)을 넘지 않게 유지 (현재 78×5 + 108).
  - 사용자 지정→호출할 PPT 이름(필수, 미입력 시 에러) + 내용1(=오른쪽 문구)·내용2·내용3(**선택**).
    내용1~3 칸은 항상 보이고, 고른 템플릿에서 찾은 입력 자리를 아래에 안내(`ItemDialog._update_custom_slots`, 이름 바꾸면 갱신).
    템플릿 목록에서 추가(더블클릭·드래그·[순서에 추가], `App._add_template`)할 때 입력 자리가 있으면 이 다이얼로그를 띄우고, 없으면 바로 추가.

## 5. 순서 항목 종류와 PPT 호출 규칙 (models.py / worship_builder.py)
| kind     | 왼쪽(label)      | 오른쪽(value)  | PPT 호출 규칙 |
|----------|------------------|----------------|---------------|
| `custom` | PPT 이름 (예: 예배로 부름) | 내용1(선택) + `options.text2`·`text3` 내용2·3(선택) | 템플릿 폴더에서 **왼쪽 이름**의 PPT. 내용이 있으면 텍스트 입력 |
| `prayer` | 기도 종류 = 기도 템플릿 이름 (예: 합심기도) | 기도자 이름(필수) | **기도 폴더**(`paths.prayer_dir`)에서 **왼쪽 이름**의 PPT 호출 + 지정 텍스트 박스에 이름 입력 |
| `hymn`   | 찬송             | `29장`(필수)   | 찬송가 폴더(하위 포함)에서 **오른쪽 장 번호**의 PPT 검색 |
| `bible`  | 말씀             | `요3:16`(필수) | Bible2PPT CLI로 구절 PPT 생성 후 병합 |
| `sermon` | 설교             | 설교 제목(필수) + `options.preacher` 설교자(필수) | 설정의 **설교 템플릿 PPT**(파일 하나) 호출 + 제목→내용1, 설교자→내용2 |
| `ccm`    | CCM              | CCM 폴더 기준 상대 경로 `봄/곡명.pptx`(필수, 화면엔 `ccm_title`=파일명) | CCM 폴더(하위 포함)의 그 파일. 텍스트 입력 없음 |

- 기도 폴더(`paths.prayer_dir`, 기본 `prayers`)가 없거나 비었으면 기도 추가 시 에러(`App._can_add`).
  1.0.3 초기 설정의 `item_templates.prayer_dir` 는 `ConfigManager` 가 시작 시 `paths.prayer_dir` 로 옮긴다.
  예전 프리셋의 기도 항목(label `기도`)은 그 폴더의 `기도.pptx` 를 찾는다.
- 설교 템플릿(`item_templates.sermon_file`, `worship_builder.sermon_template`)이 미설정/없으면 설교 추가 시 에러.
- 템플릿 파일 매칭: 파일명(확장자 제외)이 label과 같거나, 공백 제거 후 같으면 일치. 확장자 `.pptx/.pptm/.ppt`.
- 찬송가 매칭 우선순위: ① 파일명이 `29`, `029`, `29장`, `029장` ② 파일명의 **첫 번째 숫자**가 29 ③ (숫자 없는 값) 파일명에 값 포함.
- CCM 매칭(`find_ccm`): ① 저장된 상대 경로 ② 같은 파일명(다른 하위 폴더로 이동) ③ 공백·대소문자 무시 곡명. CCM 폴더 밖 경로(`..`·절대 경로)는 거부.
- 텍스트 입력 규칙("지정 텍스트 박스", 항목당 최대 `models.MAX_FILL`=3칸, 값 목록은 `ServiceItem.fill_values`):
  - 기도·찬송: 내용1=value / 사용자 지정: 내용1=value, 내용2·3=`options` 의 `models.CUSTOM_TEXT_KEYS`(text2·text3)
    / 설교: 내용1=제목, 내용2=설교자 / 말씀·CCM: 없음. 빈 값은 건너뜀.
  - 입력 자리 감지(`worship_builder.template_fill_slots`): PowerPoint 를 띄우지 않고 .pptx/.pptm 슬라이드 XML 을 읽어
    `{{내용n}}` 토큰·이름 `내용n` 도형을 찾는다. PowerPoint 는 `{{내용1}}` 을 `{{`/`내용`/`1}}` 처럼 글자 조각(run)으로 나눠
    저장하므로(한글 언어 run 분리, 부분 서식) **문단 단위로 이어 붙여** 찾을 것. .ppt·손상 파일은 None → 모든 칸 표시.
    XML 파서 대신 정규식(외부 엔티티·엔티티 폭탄 영향 없음), 슬라이드 XML 20MB 초과는 건너뜀.
  1. 해당 항목 슬라이드에서 번호 토큰 `{{내용n}}`을 찾아 치환 (서식 유지, `TextRange.Replace`).
     설정의 치환 토큰(기본 `{{내용}}`)이 기본형이고 `models.numbered_token` 이 닫는 `}}` 앞에 번호를 넣는다. 설정 창은 입력칸 옆에 실제 이름(`{{내용1}}`…, `내용1`…)을 실시간 미리보기.
  2. 토큰이 없으면 이름이 `내용n`(설정의 텍스트 상자 이름 + 번호)인 도형의 텍스트를 교체.
  3. 하위 호환: 1번 값은 예전 `{{내용}}` 토큰 / `내용` 도형에도 들어간다.
  4. 못 찾으면 경고 로그만 남기고 진행. 찬송가·CCM PPT 는 토큰이 없는 게 보통이라 경고하지 않음(`MergeJob.warn_missing`).
  5. n 번 값이 비었으면 설정의 빈 칸 기본값(`ppt.fill_defaults[n-1]`)을 넣고, 기본값도 비면 **토큰 글자를 지움**
     (결과에 `{{내용2}}` 가 남지 않게). 그래서 `_fill_text` 는 값이 없어도 모든 항목에 실행. 이름이 `내용n` 인 상자는
     값·기본값이 없으면 그대로 둔다(상자 글자는 템플릿 문구). 기본값으로만 채운 자리는 못 찾아도 경고하지 않음.
     `TextRange.Replace(…, "")` 로 지우는 동작은 실제 PowerPoint 로 확인함.
- 전체 공통 토큰(모든 슬라이드): 파일 이름 항목과 같은 `{{날짜}}` `{{년}}` `{{월}}` `{{일}}` `{{요일}}` `{{교회이름}}` `{{예배종류}}`
  (`worship_builder.ppt_replacements` ← `field_values(pad=False)`). 파일 이름은 `{월}`·`{일}` 을 0 채움(`04`, 정렬용),
  PPT 는 0 없이(`4` — `{{월}}월 {{일}}일` 처럼 쓰므로). 항목을 늘릴 때는 `FILENAME_FIELDS`·`field_values` 한 곳만 고치면 둘 다 반영.

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
- 말씀 추가(팔레트·왼쪽 성경 영역) 전에 `Bible2PPT.check()` 로 exe 를 확인하고 실패하면 에러 창(`App._can_add`).
  프리셋/파일로 불러온 말씀 항목은 막지 않고 생성 전 점검에서 알린다.
- 역본은 최대 2개(역본2 를 고르면 `-b 역본1 -b 역본2` → Bible2PPT 가 한 슬라이드에 나란히 배치).
  `worship_builder.item_bibles`: 항목에 역본(`options.bible`·`bible2`, `models.BIBLE_KEYS`)이 하나라도 있으면
  [항목 역본1(비면 설정 기본 역본1), 항목 역본2], 없으면 설정 기본값 [`bible2ppt.bible`, `bible2ppt.bible2`].
  공백뿐인 값은 빈 값, 같은 역본(`bible_arg` 대소문자 무시 — `개역개정 · 갓피아 [1]`·`1`)은 한 번만, 아무것도 없으면
  `-b` 없이 Bible2PPT 마지막 선택값. 예전 항목(역본1만)은 설정의 기본 역본2 를 쓰지 않고 한 역본으로 생성.
  역본1 칸이 비면 설정 기본 역본1 을 쓰므로, 역본2 만 지정은 **기본 역본1 도 없을 때만** 에러 창(메인 성경 영역·말씀 항목 입력,
  `ItemDialog(fallback_bible=)`). 설정 창은 기본 역본2 만 지정하면 에러.
  메인 성경 영역의 빈 칸은 의미가 있으므로(빈 역본1 = 기본값, 빈 역본2 = 한 역본) 역본 목록을 다시 불러와도 채우지 않고,
  직접 입력한 이름만 `bible_bridge.match_bible` 로 목록 항목에 맞춘다. 설정 저장 후에는 기본 역본이 **바뀐 경우에만**
  그 기본값을 쓰던 칸(빈 칸·예전 기본값)을 새 기본값으로 바꾼다 (사용자가 지운 역본2 가 몰래 되살아나지 않게).
- `--list-bibles` 출력(포크 `CommandLine.cs` 확인): 헤더 `ID\t소스\t성경` + 줄마다 `ID\t소스\t이름`, 리다이렉트 시 UTF-8.
  소스별 실패는 stderr 로만 알림(종료 코드 0). → `parse_bible_list` 가 `이름 · 소스 [ID]` 로 표시, `bible_arg` 가 `[ID]` 를 `-b` 로 전달.
- 역본 선택은 모두 `BibleField`(직접 입력 + 🔍 → `BiblePicker`): 메인 성경 영역(역본1 줄 아래 역본2 줄), 말씀 항목 입력,
  설정 기본 역본 — 각각 역본1·역본2 두 칸.
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
    테마·마스터·레이아웃·**슬라이드별 배경(그림 포함)**까지 보존. 문서 창이 필요해서 `dest.NewWindow()` 를 쓰되 **화면에 띄우지 않는다**
    (`_FrameHider`, 아래). 비동기라 슬라이드 수 증가를 폴링으로 확인, 실패 시 해당 파일만 아래 방식으로 대체.
    ExecuteMso 는 활성 창에 붙여넣으므로 `_activate` 가 `ActiveWindow.Caption` 이 작업 창일 때는 Activate 하지 않는다
    (Activate 는 숨긴 창을 다시 보이게 함 — 예전엔 최소화해도 붙여넣을 때마다 되살아나 생성 내내 창이 보였음).
  - `windowless`: `Slides.Paste` → 슬라이드별 `Design`(파일+디자인명 캐시)·`CustomLayout` 원본 지정 → 배경 복제
    (단색은 RGB 정확 복사, 그라데이션/그림 등은 원본 도형을 잠시 숨기고 2560px PNG로 내보내 `UserPicture`). 클립보드 실패 시 `InsertFromFile`.
  - ⚠ 그냥 `Slides.Paste`/`InsertFromFile` 만 쓰면 **슬라이드별 배경이 사라진다** (대상 테마 사용과 동일) — 2026-10-06 실험으로 확인.
- 애니메이션/전환/폰트는 슬라이드 복사 시 그대로 따라온다. 원본 파일은 절대 수정하지 않는다(ReadOnly 열기).
- **원본이 이미 PowerPoint 에 열려 있으면**(템플릿을 열어 고치는 중 생성 등) `Presentations.Open` 은 새로 열지 않고 사용자의
  그 문서를 돌려준다(실측) → 첫 원본이면 SaveAs 로 결과 이름이 되어 닫히고, 나머지면 Close 로 닫혀 저장 안 한 수정이 사라졌음.
  `_open` 이 `_already_open` 으로 확인해 저장된 파일을 작업 폴더에 복사해 연다(경고 로그). 사용자 문서는 그대로 둔다.
  OneDrive 폴더의 파일은 `FullName` 이 `https://d.docs.live.net/…` 라서(실측) 주소로 열린 문서는 **파일 이름**으로 비교한다.
- **OneDrive 출력 폴더**: 처음부터 결과 경로로 SaveAs 하면 PowerPoint 가 그 뒤로 클라우드 문서로 다뤄 마지막 `Save` 가 내 PC 의
  파일에 들어가지 않았다(실측: 60초 뒤에도 첫 PPT 슬라이드만). → 작업 폴더(`_tmp/result/`)에 만들고 **마지막에 결과 경로로
  SaveAs**. 결과 파일이 이미 PowerPoint 에 열려 있으면 다 만든 뒤가 아니라 시작할 때 오류(`_already_open(by_name=False)`).
- 암호가 걸린 .pptx/.pptm(ZIP 대신 OLE 형식 — `_OLE_SIGNATURE`)은 열기 전에 오류. 그대로 열면 PowerPoint 가 암호 창을 띄우고
  끝없이 기다렸음(실측). .ppt 는 원래 OLE 라 구별 못 함 → 멈춤 안내로 처리.
- 창 숨기기 `_FrameHider`(with 블록, NewWindow·Activate·Close·Quit 때만 짧게): 전용 스레드의 WinEvent 훅(EVENT_OBJECT_SHOW)으로
  `PPTFrameClass` 창이 보이는 즉시 `SW_HIDE`. 작업 스레드에 걸면 COM 호출 중엔 콜백이 오지 않음(실측). 다른 프로세스 창의
  ShowWindow 는 PowerPoint UI 스레드가 바쁘면 ~0.25초 늦으므로, 블록에 들어갈 때 PowerPoint 가 미리 숨겨 만들어 둔 문서 창을
  WS_EX_LAYERED+알파 0 으로 투명하게 해 두고 나갈 때 원래 스타일로 되돌린다(실측: 결과를 보여 주기 전까지 불투명한 창 0회).
- 결과 열기(`merge(open_mode="open"|"show")`): 저장 후 숨겨 둔 작업 창을 그대로 보여 준다(`_reveal`: 최대화·앞으로·쇼면
  `SlideShowSettings.Run()`). PowerPoint 를 끄고(`Quit`) `os.startfile` 로 다시 켜는 시간이 없어짐. 창 없는 방식이라 작업 창이
  없으면 False → `App._open_result` 가 예전처럼 파일로 연다. 쇼를 Esc 로 끝내면 편집 화면이 남는다(Show 동사 `/s` 는 PowerPoint 도 닫힘).
- 텍스트 입력은 항목 슬라이드의 텍스트를 `_collect_texts` 로 한 번만 읽고 토큰이 든 범위에만 `Replace`(대소문자 무시 비교 —
  `TextRange.Replace` 기본과 같음). 공통 토큰도 항목별로 같이 처리. 예전처럼 토큰마다 모든 도형을 COM 으로 훑으면 37장에 약 10초.
- 종료 전에 클립보드를 비운다 — 붙여넣은 슬라이드가 클립보드에 남아 있으면 `Quit` 가 5초 넘게 걸림(실측).
- 말씀이 있으면 `PptEngine.hold()` 가 빈 프레젠테이션(창 없음)을 열어 둔 채 Bible2PPT 를 실행 → Bible2PPT 는 실행 중인
  PowerPoint 를 같이 쓰고, 열린 프레젠테이션이 있으면 끝날 때 종료하지 않음(`HiddenPowerPoint.Dispose` 확인). 말씀마다 켜고 끄지 않음.
- 슬라이드 비율(`config.ppt.slide_size`: original | 16:9 | 4:3, 하단 생성 옵션 아래 `App.size_seg`):
  original 은 첫 PPT 크기. 16:9(960×540pt)·4:3(720×540pt)이면 베이스를 SaveAs 한 직후 `_apply_slide_size` 로
  PageSetup 을 바꾼 뒤 나머지를 이어 붙인다(붙인 슬라이드는 결과 크기에 맞춰짐). COM 크기 변경은 [맞춤 확인]처럼 동작(실측:
  4:3→16:9 내용 그대로 가운데, 16:9→4:3 내용을 0.75배로 줄여 위아래 여백 — 잘리지 않음). 비율을 직접 골랐으면
  크기가 다른 원본에 ⚠ 경고를 남기지 않음(의도한 동작).
- UI 스레드를 막지 않도록 작업 스레드에서 `pythoncom.CoInitialize()` 후 실행.

## 8. 사용자 설정 (config.json)
```json
{
  "church_name": "", "worship_type": "주일오전예배", "filename_format": "{날짜}_{교회이름}_{예배종류}",
  "appearance_mode": "System", "use_system_accent": true, "setup_done": false,
  "paths": {"template_dir": "templates", "hymn_dir": "hymns", "ccm_dir": "ccm", "prayer_dir": "prayers",
            "output_dir": "output"},
  "bible2ppt": {"exe_path": "", "bible": "", "bible2": "", "template": "", "lines": null, "timeout_sec": 180,
                "bible_list": []},
  "item_templates": {"sermon_file": ""},
  "ppt": {"text_placeholder": "{{내용}}", "text_shape_name": "내용", "fill_defaults": ["", "", ""],
          "open_after_build": true, "start_slideshow": false,
          "merge_mode": "source_formatting", "slide_size": "original", "temp_mode": false, "temp_keep": 10}
}
```
- `open_after_build`·`start_slideshow`: 생성 시작 때 `BuildRequest.open_mode` 로 정해 병합한 PowerPoint 에서 바로 연다(§7 결과 열기).
  그게 안 됐을 때(`req.opened` False)·[임시 파일]의 쇼는 `os.startfile(path, "show")`(= .pptx 의 Show 동사, `POWERPNT.EXE /s`).
  Show 동사가 없으면 App Paths 의 POWERPNT.EXE `/s`, 그래도 실패하면 일반 열기 + 로그 경고. `open_after_build`·`temp_mode` 가 모두 꺼지면 비활성.
- `use_system_accent`: Windows 강조 색을 프로그램 강조 색으로 (⚙ 설정 '화면', 다시 켤 때 적용). `ui_widgets.apply_accent` 가
  `App.__init__` 에서 위젯을 만들기 전에 `ACCENT`·`ACCENT_HOVER` 를 **제자리에서** 바꾼다(리스트 — `from ui_widgets import ACCENT`
  로 가져간 이름도 같은 개체). 색은 `HKCU\...\Explorer\Accent\AccentPalette`(7단계)에서 라이트 모드 Dark1·다크 모드 강조 색부터,
  흰 글자 명암비 4.5:1 이상이 될 때까지 어두운 단계로(`pick_accent` — 노랑 등 밝은 색도 버튼 글자가 읽히게). 못 읽으면 기본 파랑.
  항목 종류 색(`KIND_COLORS`)·삭제 빨강(`DANGER`)은 강조 색과 별개 — 강조 색이 빨강이면 삭제 글자와 비슷해 보일 수 있음.
- `temp_mode`/`temp_keep`: §9-2 임시 실행 모드. `temp_keep` 은 `config_manager.clean_temp_keep` 으로 1~100 정수(잘못된 값은 10).
- 상대 경로는 프로그램 폴더 기준. 누락 키는 기본값으로 자동 병합. 교회 이름/예배 종류/파일 이름 형식/폴더는 메인 UI에서 바로 저장됨, 나머지는 ⚙ 설정.

## 9. 프리셋 규칙 (presets.json)
- 형식: `{"presets": [{"name": "주일오전예배", "items": [{"kind","label","value","options"}...], "filename_format": "…"}]}`
- **최대 10개**, 이름 **1~10글자**, 같은 이름 저장 시 덮어쓰기 확인, **삭제 전 확인 창** 필수.
- 불러오기 시 순서 목록을 교체하고 예배 종류를 프리셋 이름으로, 파일 이름 형식을 프리셋의 형식으로 설정
  (`filename_format` 이 없는 예전 프리셋은 기본 형식). 저장 시 현재 파일 이름 형식도 함께 저장.
- 기본 프리셋: 주일오전예배, 주일오후예배, 수요예배 (최초 실행 시 생성).
- 텍스트 파일 입출력(순서 제목 줄 [파일 불러오기]/[파일 내보내기], `models.export_to_text`/`import_from_text`, UTF-8 BOM):
  `# 교회이름: …`/`# 예배종류: …`/`# 날짜: …`/`# 파일이름형식: …` 메타 줄 + 항목 줄 `[kind]label|value[|역본1|역본2(말씀) 또는 설교자(설교) 또는 내용2|내용3(사용자 지정)]`.
  불러올 때 모르는 kind 는 custom, `validate_item` 실패 줄은 건너뜀.
  값 안의 `|`·`\` 는 `\|`·`\\` 로 내보냄(`_esc`/`_split_fields` — 그 외 `\` 는 글자 그대로라 예전 파일도 읽힘. 예전 설교 줄의
  셋째 칸 뒤 `|` 는 설교자 글자). 한 줄 `MAX_IMPORT_LINE`(2000자) 초과는 건너뜀. `make_item` 이 줄 바꿈·탭 등 제어 문자를
  공백으로 바꿈(입력칸에 붙여 넣은 줄 바꿈이 내보낸 파일의 한 줄 = 한 항목 형식을 깨뜨렸음).

## 9-1. 결과 파일 이름 형식 (worship_builder.output_filename)
- 항목(`models.FILENAME_FIELDS`): `{날짜}`(YYYY-MM-DD) `{년}` `{월}` `{일}`(0 채움, 예 `10`·`07`) `{요일}`(한 글자, 예 `일`) `{교회이름}` `{예배종류}`. 기본 `models.DEFAULT_FILENAME_FORMAT`.
- 값을 넣은 **전체 문자열**을 `sanitize_filename` 하고 `.pptx` 를 붙임(형식 끝의 `.pptx` 는 한 번만). 결과가 비면 기본 형식.
- `validate_filename_format`: 빈 형식·모르는 `{항목}`·쓸 수 없는 문자뿐인 형식은 오류 → 미리보기 ⚠, 프리셋 저장·생성 시 에러 창.

## 9-2. 임시 실행 모드 (worship_builder.temp_result_dir)
- 하단 [저장 없이 임시 파일로 실행](`App.temp_mode`, 생성 버튼 바로 아래 — 옆 열에 두면 하단 바가 넓어져 상태 표시 칸이
  약 2/3 로 줄었음)을 켜면 출력 폴더 대신 `%LOCALAPPDATA%\ChurchPPTMaker\임시` 에 `형식대로 만든 이름_시분초.pptx` 로 만들고
  **항상 바로 연다**(슬라이드 쇼 옵션 적용, [생성 후 열기]는 잠김). 생성 버튼 글자가 '예배 PPT 임시 생성 및 실행'으로 바뀜.
  %TEMP% 는 Windows(저장소 센스·디스크 정리)가 지울 수 있어 보관 개수를 지킬 수 없으므로 쓰지 않음.
- 이름에 시분초(같으면 `_2`…)를 붙이는 이유: 앞서 연 임시 파일이 PowerPoint 에서 열려 있어도 다음 생성이 덮어쓰지 않게.
  같은 이유로 덮어쓰기 확인 없음. 모드는 생성 시작 때 값으로 고정(생성 중 체크를 바꿔도 이번 생성은 그대로).
- 보관: 임시 생성이 끝날 때마다 `prune_temp_results(ppt.temp_keep)` 가 최근 N 개만 남김(⚙ 설정 '최대 보관 개수' 1~100,
  설정을 저장할 때도 바로 정리). PowerPoint 에서 열려 있는 파일은 지울 수 없어 건너뛰고 다음 정리 때 다시 시도.
  목록(`list_temp_results`)은 PowerPoint 잠금 파일 `~$이름.pptx`·링크·.pptx 외 파일을 제외.
- 하단 [임시 파일](`TempFilesDialog`): 최근순 목록(수정 시각·이름·크기) → [열기]·[슬라이드 쇼]·[다른 이름으로 저장]·[폴더 열기].
  저장은 출력 폴더·`_시분초` 를 뗀 이름(`saved_name_for_temp`)을 제안하고 복사(`shutil.copy2`). 임시 폴더 안(곧 정리됨)·
  `output_conflict`(원본 폴더·원본 파일)에 걸리는 곳은 거부, 확장자는 .pptx 로 맞춤. 목록을 연 뒤 지워진 파일·임시 폴더 밖
  경로는 `App._temp_file_ok` 가 거부.
- `secondary_button` 은 `height` 를 덮어쓸 수 있음(하단 [로그]·[임시 파일] 은 30).
- 정리는 **영구 삭제**라서 임시 폴더를 쓸 수 없는 경우를 `temp_dir_conflict(cfg)` 로 막는다 — 임시 생성·[임시 파일] 목록·
  설정 저장 때 정리 모두 이 검사를 거친다: ① 임시 폴더나 그 위 `ChurchPPTMaker` 폴더가 정션·심볼릭 링크(`_has_reparse_point`,
  `list_temp_results` 도 빈 목록 반환 — 링크 대상의 PPT 를 지우지 않게) ② 원본·출력 폴더나 설교·성경 템플릿이 임시 폴더 안
  (설정 파일 직접 수정 등). 📁 폴더 지정에서도 임시 폴더 안은 거부(`App._choose_folder`).
- 임시 생성이 실패하면 반쯤 만든 결과를 지운다(목록에 깨진 PPT 가 남지 않게). 일반 생성은 사용자 폴더라 그대로 둔다.
- 생성(병합) 중에는 PowerPoint 로 파일을 열지 않는다(`App._busy_merging` — `_open_ppt`·임시 파일 열기). [원본 서식 유지]
  붙여넣기(`ExecuteMso`)는 PowerPoint 의 **활성 창**에 붙여넣으므로, 그사이 연 파일이 활성 창이 되면 그 파일에 슬라이드가 붙을 수 있음.
- 다른 이름으로 저장은 `sanitize_filename(이름) == 이름` 일 때만 (CON·NUL 같은 장치 이름·150자 초과 거부).

## 10. 개발 히스토리
- 2026-10-06 v1.0: UI 초안 기반 최초 구현 (UI, 프리셋, 드래그&드롭, Bible2PPT CLI 연동, COM 병합 엔진).
  - 병합 시 슬라이드 배경 소실 발견 → `PasteSourceFormatting` 기본 + windowless 배경 복제 대체 경로 추가.
  - CTkToplevel 반복 생성/파괴 시 Tk 접근 위반(크래시) 재현 → 모든 다이얼로그를 창 내부 오버레이(_Modal)로 전환, 12회 반복 테스트 무크래시.
  - 검증: 실제 PowerPoint 16.0으로 7개 항목 병합(텍스트 치환·애니메이션·단색/그림 배경 유지, 원본 미수정), UI 자동 테스트 25항목.
    Bible2PPT는 실제 exe 없이 가짜 생성기로만 검증됨 → 실제 exe 연동 확인 필요.
- 2026-10-07 v1.0.3: 기도 폴더(왼쪽 [폴더], 추가 시 기도 종류 선택)·설교 템플릿 파일(⚙ 설정, 설정 창 2쪽으로 분리) 지정,
  템플릿 폴더 줄을 [템플릿 파일] 카드 위로 이동,
  Bible2PPT 미설정 시 말씀 추가 차단, 템플릿 파일 목록 10% 확대, 결과 파일 이름 형식(프리셋별 저장).
  - 검증: 격리 설정으로 UI 동작·에러 창 자동 테스트, 실제 PowerPoint 로 합심기도·설교 병합(이름·제목·설교자 입력, 형식대로 파일명).
- 2026-10-07 v1.0.4: 사용자 지정 내용1~3 입력(템플릿 입력 자리 자동 감지), 빈 칸일 때 `{{내용1}}~{{내용3}}` 기본값(없으면 토큰 삭제),
  결과 슬라이드 비율(원본/16:9/4:3) 선택.
  - 검증: 실제 PowerPoint 로 쪼개진 토큰·'내용3' 상자 치환, 토큰 삭제/기본값, 4:3↔16:9 변환(내용 잘림 없음) 확인.
- 2026-10-07 v1.0.5: 템플릿 파일 불러오기·열기·삭제(휴지통), 순서 항목 더블클릭으로 원본 열기(말씀은 Bible2PPT 로 생성),
  파일 이름 미리보기 아랫줄 이동, 처음 켤 때 최소 크기, 최소 크기에서 상단 입력칸이 꽉 차도록 조정.
  - 검증: 항목 종류별 열리는 파일·클릭/더블클릭 구분·불러오기·휴지통 삭제 자동 테스트(startfile 가로채기),
    말씀 열기는 Bible2PPT 없이 생성 단계만 대체해 확인. 시작 크기 1000×892 표시, 상단 요청 너비 959 고정.
- 2026-10-09 v1.0.6: 말씀 역본 2개 지정(Bible2PPT `-b` 두 번 → 나란히 배치) — 메인 성경 영역·말씀 항목 입력·설정 기본 역본에 역본2 칸.
  - 검증: `-b` 인자 순서·중복 제거·기본값 대체, 텍스트 내보내기/불러오기 왕복, UI 입력(역본2만 지정 시 에러) 자동 테스트.
    실제 Bible2PPT 로 두 역본 생성은 아직 확인 안 함. 성경 영역이 한 줄 늘어 시작 크기 1000×930.
  - 엣지 케이스 점검(25건 자동 테스트)으로 수정: 같은 역본 다른 표기 중복 전달, 공백뿐인 역본 값, 역본1 칸 비움+역본2 가
    기본 역본이 있는데도 막히던 문제, 지운 역본2 가 목록 재로드·설정 저장 때 되살아나던 문제, 직접 입력한 역본 이름이
    목록 재로드 때 기본 역본으로 덮이던 문제(역본1 은 이전부터 있던 문제).
- 2026-10-09 v1.0.6: 저장 없이 임시 파일로 실행하는 모드, 임시 파일 최대 보관 개수 설정, [임시 파일] 목록(열기·쇼·저장) (§9-2).
  - 검증: 실제 PowerPoint 로 임시 생성 3회(보관 2개 → 가장 오래된 것 삭제, 출력 폴더 비어 있음)·일반 생성 복귀,
    실제 [생성] 버튼 경로(임시 폴더에 만들어 쇼로 열기), 함수 19건·UI 18건 자동 테스트(잠금 파일 제외, 열린 파일 건너뜀,
    잘못된 보관 개수, 템플릿 폴더·임시 폴더로 저장 거부, 목록 연 뒤 삭제된 파일). 하단 바 높이 +16px(시작 1000×946).
  - 엣지 케이스·보안 점검(28건 자동 테스트, 실제 `mklink /J` 정션)으로 수정: 임시 폴더가 정션이면 정리가 링크 대상 PPT 를
    영구 삭제하던 문제, 원본·출력 폴더가 임시 폴더 안이면 사용자 PPT 가 정리되던 문제, 실패한 임시 생성의 반쯤 만든 파일,
    병합 중 PPT 를 열면 붙여넣기가 그 파일로 갈 수 있던 문제(템플릿 열기는 이전부터), 저장 이름의 장치 이름(CON 등).
- 2026-10-09: 병합 속도·창 깜빡임 개선 — 작업 창을 숨김(투명+즉시 숨김), 결과를 같은 PowerPoint 에서 바로 열기,
  텍스트 입력 COM 호출 축소, 종료 전 클립보드 비우기, Bible2PPT 와 PowerPoint 공유.
  - 검증(실제 자료 11항목·37장, 말씀 1개): 생성 30.7초 + 다시 켜서 열기 → 16.7초에 결과 창 표시까지(쇼 13.5초,
    PowerPoint 를 이미 켜 둔 경우 11.3초). 창 감시(5ms)로 결과 표시 전 불투명 창 0회(이미 켜 둔 경우 Activate 때 10ms 1회),
    사용자가 열어 둔 문서 창은 그대로. 이전 엔진 결과와 슬라이드 텍스트 동일·남은 토큰 없음, 결과 창을 닫으면 PowerPoint 종료.
- 2026-10-09: UI 다듬기 — 상태 부연 설명을 아랫줄로, 성경 영역 버튼을 제목 줄로·역본 칸 이름 표시, 긴 순서 행 줄 바꿈,
  순서 목록 스크롤 속도, 설정 창을 한 쪽(세로 스크롤)으로 합침(높이 = min(내용, 창 높이×0.9−170): 다이얼로그가 창 45% 지점에
  가운데 정렬이라 이보다 크면 위가 제목 표시줄 밑으로 잘렸음). 입력 오류 시 `_reveal` 이 그 칸으로 스크롤·포커스.
  - 검증: 실제 앱 화면 캡처(최소 창: 긴 CCM·설교·사용자 지정 행 아랫줄로, 넓은 창: 한 줄·전체 글자로 복귀, 하단 바 높이 그대로),
    설정 창 위/아래 스크롤 캡처, 잘못된 보관 개수·제한 시간 저장 시 해당 칸으로 스크롤되고 저장 안 됨.

- 2026-10-09: 제목 옆 [? 도움말] — 처음 설정·순서·말씀·템플릿·자동 문구·파일 이름·찬송가/CCM·프리셋·생성·설정·문제 해결 11절.
  - 검증: 라이트/다크 화면 캡처, 목차 이동(마지막 절 포함)·검색('요일' 8건) 자동 확인.
- 2026-10-09: Windows 제목 표시줄 대신 프로그램 안 제목 줄 + 창 버튼 (`window_frame.py`).
  - 검증: 실제 앱 캡처(일반·최대화), 가장자리 판정(좌·우·아래·모서리), 최대화↔복원 3회 크기 그대로, 최소화·복원,
    실제 마우스 입력(SendInput)으로 제목 줄 끌기·두 번 눌러 최대화/복원·위쪽 가장자리 크기 조절, UI 테스트 39건.
  - 찾아서 고친 것: 제목 줄 끌기에서 SendMessage 로 옮기기 루프를 시작하면 GIL 오류로 강제 종료 → PostMessage.
    최대화 후 복원하면 창이 한 번 11px 커짐 → 일반 상태 크기를 기억했다가 되돌림.
    제목 줄 위에서 마우스를 움직이면 제목 줄이 4배쯤(69→300px) 커짐 → 커서 지정이 CTkFrame.configure 를 거쳐 틀 크기를
    기본값으로 되돌렸음. 안쪽 Tk 위젯에 직접, 바뀔 때만 지정 (§11).
- 2026-10-09: Windows 강조 색을 프로그램 강조 색으로 (⚙ 설정 '화면'에서 끄기 가능).
  - 검증: 이 PC 빨강 팔레트(라이트 #D20E1E·다크 #E81123, 흰 글자 5.5·4.6:1), 노랑·연두·기본 파랑·회색 팔레트 선택 결과,
    라이트/다크 화면 캡처, 설정 끄기 저장 → 다시 켤 때 기본 파랑, 잘못된 설정 값 → 기본값(켬), UI 테스트 39건.
- 2026-10-09 v1.0.7: 위 변경 전체의 엣지 케이스·보안 점검 (함수 20건, 실제 PowerPoint 엔진 21건, UI 39건, 실제 [생성] 버튼 경로).
  - 엔진: 글자 조각으로 나뉜 토큰·표 칸·그룹·이름 상자(`내용2`) 안 공통 토큰·빈 `{{내용3}}` 삭제, 결과 창이 읽기 전용이 아니어서
    고쳐 저장 가능, 결과를 닫으면 PowerPoint 종료, 창 없는 방식은 열기 요청에 False(파일로 엶), 손상 파일·말씀 실패 시
    숨은 창·빈 프레젠테이션·PowerPoint 가 남지 않음, 나중에 연 문서 창이 투명으로 남지 않음.
  - 수정: `AllowSetForegroundWindow` 를 PowerPoint 프로세스로 한정, 창 숨기는 동안 우연히 숨겨진 사용자 PowerPoint 창 되살리기,
    작업 창 제목을 못 읽어도 결과 닫기는 계속, 도움말 목차에서 끝부분 절을 고르면 다음 절이 선택되던 문제.
  - 테스트 주의: 엔진은 끝날 때 자기 스레드의 COM 을 해제하므로, 같은 스레드에서 미리 얻은 PowerPoint 개체는 쓸 수 없게 되고
    남은 참조는 PowerPoint 종료를 막는다 → 앱처럼 엔진은 작업 스레드(또는 별도 프로세스)에서 돌려 시험할 것.

- 2026-10-10: 생성 중 화면(진행·경과 시간·PowerPoint 사용 금지 안내), 설정 마법사(처음 실행 시 자동, 제목 줄 버튼).
  - 검증: 임시 설정으로 마법사 흐름 21건 + 엣지 케이스 35건(큰 폴더·따옴표 경로·환경 변수·없는 드라이브·권한 없는 폴더·
    임시 폴더·출력 폴더와 겹침·.exe 인 폴더·.pptm·역본2만·목록 불러오는 중 마침·닫힌 뒤 목록 도착·작은 창·생성 중 열기 차단,
    생성 중 키 입력·생성 실패/성공·생성 중 닫기 취소). 실제 Bible2PPT 로 역본 불러오기와 실제 PowerPoint 생성은 이번에 안 함.
  - 찾아서 고친 것: 큰 폴더 개수 세기 19초 멈춤, 작은 창에서 마법사 위 잘림, 생성 중 덮인 입력칸에 글자 입력됨,
    단계마다 다이얼로그 너비가 바뀜, 출력 폴더와 겹칠 때 어디서 바꾸는지 안내 없음.

- 2026-10-12: 전체 엣지 케이스·보안 점검 (함수 28건, UI 5건, 실제 PowerPoint 엔진·생성 경로).
  - 고친 것: 열려 있는(저장 안 한) 템플릿으로 생성하면 그 문서가 닫히며 수정이 사라짐(가장 심각), 정션 순환으로 찬송가 찾기·CCM
    목록이 멈춤, CCM 값의 UNC 경로로 네트워크 접속(21초 멈춤·NTLM 노출), 쓰기가 막힌 프로그램 폴더에서 시작 시 꺼짐,
    `2026-1-5` 날짜로 {년}{월}{일}{요일} 빈 값, 내보낸 텍스트에서 `|` 가 든 값이 칸이 밀림, 붙여 넣은 줄 바꿈, 잘못된 `lines`
    설정으로 말씀 생성 TypeError, 템플릿 입력 자리 감지의 압축 폭탄, 말씀 미리보기 정리.
  - 검증: 실제 PowerPoint 로 열린 문서가 첫·중간 원본일 때 + 결과 열기 모드(사용자 문서·수정 유지, 결과에 저장 안 한 수정 없음),
    격리 설정 5항목 생성(정션 순환 찬송가 폴더·UNC CCM 값·제어 문자 이름·{년}-{월}-{일}_{요일} 형식, 남은 토큰 없음),
    쓰기 금지 폴더(icacls)로 앱 시작·경고·프리셋 저장 에러, 휴지통 삭제 회귀.
- 2026-10-12: PowerPoint 버전·환경 차이 대비 — OneDrive 파일·출력 폴더, 생성 멈춤 안내·중단, 암호 PPT.
  - 검증(실제 PowerPoint 16.0 + 이 PC 의 OneDrive): OneDrive 의 열린 템플릿(첫·중간 원본) 보존, OneDrive 출력 폴더 결과가 2장
    그대로(고치기 전 1장), 결과 열기 모드도 같음. 암호 PPT 병합 → 고치기 전 '암호' 창(NUIDialog)에서 멈춤, [PowerPoint 창 보기]가
    그 창을 찾아 앞으로·강제 종료로 생성 정리 확인, 고친 뒤 즉시 오류. UI 14건(3초로 줄인 멈춤 안내 표시·진행 재개 시 숨김·
    중단 후 상태·PowerPoint 종료, 생성 전부터 켜진 PowerPoint 는 끄지 않음), 라이트/다크 화면 캡처, 앞선 회귀 테스트 전부.
  - 확인 못 함: 실제 라이선스 동의·로그인 창(처음 설치한 Office 가 없어 암호 창으로 대신), PowerPoint 2016 이하.

## 11. 구현 시 주의사항 (검증된 함정)
- **CTkToplevel 사용 금지** → `ui_widgets._Modal`(창 내부 오버레이) 상속. CTkToplevel 은 생성 때마다 `bind_all("<Button-1>")`를 누적하고
  창 숨김·강제 update·ctypes 호출을 하며, 반복 사용 시 접근 위반 크래시가 났다.
- tkinter 위젯 서브클래스에 `self._root` 속성을 만들지 말 것 (`Misc._root()` 메서드를 덮어써 이벤트 처리가 깨짐). → `self._win` 사용.
- `CTkFrame.configure(...)` 는 무엇을 바꾸든(커서 등) 틀 크기를 기본값(200×배율)으로 다시 지정한다 → 자식에 맞춰 줄어 있던
  틀이 커짐(실측: 제목 줄 위에서 마우스를 움직일 때마다 커서를 바꿔 69→300px). 커서처럼 자주 바꾸는 속성은 CTk 를 거치지 말고
  안쪽 Tk 위젯(`_canvas`, `_label`)에 `tk.call(w._w, "configure", "-cursor", …)` 로 직접, 그리고 값이 바뀔 때만.
- `CTkFrame` 은 높이 2px 이하를 0으로 그려 보이지 않는다 → 구분선/연결선은 `height=3` 이상.
- CTk 의 `place(x=, y=)` 는 좌표에 배율을 곱한다 → 실제 픽셀 좌표는 `ScalingTracker.get_widget_scaling` 으로 나눠 전달.
- 폰트는 `ui_widgets.font()` 캐시만 사용 (렌더링마다 CTkFont 생성 금지).
- 작업 스레드에서 tkinter 호출 금지 → `post_to_ui()` / `run_in_thread()`.
- `_Modal` 은 중첩 가능(예: 항목 입력 → 역본 검색). Esc/Enter 는 창 단위 바인딩이라 `_Modal._stack` 맨 위만 반응하고,
  닫힐 때 아래 다이얼로그 grab 을 복구한다. 바인딩 해제는 `_unbind()` 사용 — Python 3.12 이하 `Misc.unbind(seq, funcid)` 는
  seq 의 바인딩을 **전부** 지워 아래 다이얼로그의 Esc/Enter 가 사라진다.
- `PptEngine` 은 내부에서 CoInitialize/CoUninitialize 하므로 반드시 작업 스레드에서 `with` 로 사용.
- **창 틀(`window_frame`)**: Tk 이벤트 처리(Python 콜백) 안에서 다른 메시지 루프를 도는 Win32 호출(SendMessage 로 창 옮기기
  시작 등)을 ctypes 로 부르지 말 것 — 그 루프 안에서 Tk 가 다시 Python 콜백을 부르면 tkinter 의 GIL 상태가 어긋나 강제 종료(실측).
  PostMessage 로 미룬다. 창 프로시저(서브클래스) 콜백은 예외를 밖으로 내지 않고 DefSubclassProc 로 넘김, WM_NCDESTROY 에서 해제.
  최대화 때 Tk 가 잰 틀 두께가 복원에도 쓰여 창이 커지므로 `App._on_state_change` 가 일반 상태 크기를 되돌린다.
- **창 크기 조절 렉** (2026-10-07 실측, 순서 10개 기준 드래그 한 단계 약 480ms → 7ms):
  1. Tk 창 클래스 `TkTopLevel`/`TkChild` 의 `CS_HREDRAW|CS_VREDRAW` 때문에 1px 만 바뀌어도 모든 하위 HWND 를 다시 그림
     → `main._disable_full_redraw_on_resize` 가 시작 시 끔 (클래스 단위라 프로세스 전체 Tk 창에 적용, Tk 위젯은 스스로 다시 그림).
  2. 위젯 수백 개 재배치 자체가 수백 ms → 모든 내용은 `App._content`(CTkFrame) 안에 두고, 창 `<Configure>` 가
     `RESIZE_SETTLE_MS`(150ms) 동안 멈췄을 때만 `_fit_content` 로 실제 픽셀 크기를 지정(재배치 1회). 새 위젯은 `self` 가 아니라
     `self._content` 에 grid 할 것 (오버레이 `_Modal`·달력은 그대로 창(`winfo_toplevel`)에 place).
  3. 위젯(HWND) 수 줄이기 (리사이즈 후 재배치 557ms → 322ms, 템플릿 파일 수와 무관):
     순서 행은 둥근 알약 2개만 CTkButton, 행 틀·손잡이·연결선·▲▼✕ 는 일반 tk 위젯(`mode_color`/`scaled`/`tk_font` 로 CTk 와 같은 색·크기).
     템플릿 파일 목록은 `tk.Listbox` 하나 (`DragController.attach` 에 payload 함수를 넘겨 누른 위치의 항목을 드래그,
     기본 클릭 선택 바인딩은 "break" 로 끄고 `_select_file` 만 선택 표시). 일반 tk 위젯은 모드 전환 시
     `AppearanceModeTracker` 콜백(`_on_appearance_change`)에서 다시 칠함 — 콜백 순회 중이므로 `after_idle` 로 지연.
  - 참고: 일반 tk.Label 200개도 리사이즈당 ~200ms → Windows 에선 위젯(HWND) 수가 곧 비용. 목록형 UI 는 위젯 수를 줄일 것.
- **Windows 배율(DPI) 변경**: CTk 위젯은 `ScalingTracker` 가 100ms 마다 모니터 DPI 를 확인해 알아서 다시 그리지만,
  일반 tk 위젯의 `tk_font`/`scaled` 값은 만들 때 배율로 고정 → `App._on_scaling_change`(ScalingTracker.add_widget)에서
  템플릿 목록 글꼴·여백을 다시 지정하고 순서 행을 다시 만든다. tk 포인트 크기 글꼴 `(FONT_FAMILY, 11)` 은 시작 시
  `tk scaling` 에 고정되므로 쓰지 말고 `tk_font`(픽셀) 사용. 검증은 `get_window_dpi_scaling` 을 바꿔치기한 시뮬레이션으로만 함.
  창 틀(제목 표시줄·테두리)은 CTk 기본 Per-Monitor **V1** 에선 배율이 바뀌어도 그대로 → Tk 생성 전
  `main._enable_per_monitor_dpi_v2` 로 **V2** 지정 (CTk 의 V1 호출은 무시됨). V2 에선 `winfo_screenwidth/height` 가
  실제 픽셀이 되므로(V1 은 논리 단위) 화면 크기는 `App._work_area`(모니터 작업 영역, 실제 픽셀)/배율로만 계산할 것.
- **최소 창 높이**: 고정 660 은 왼쪽 패널보다 낮아 템플릿 목록이 0줄로 사라졌다. `App._apply_min_size` 가 첫 배치 후
  `창 높이 - 목록 높이 + 목록 MIN_FILE_ROWS(4)줄 높이` 로 계산해 지정 (CTk 위젯의 reqheight 는 실제보다 커서 쓰지 않음).
  왼쪽 열에서 세로로 늘고 주는 행은 템플릿 목록뿐이라는 전제 — 왼쪽 패널에 가변 높이 요소를 추가하면 함께 고칠 것.
- **시작 크기 = 최소 크기** (`main.MIN_WIDTH`=1000 × 위 최소 높이, 화면 가운데). 최소 높이는 첫 배치 뒤에야 알 수 있어
  창을 `-alpha 0` 으로 만들어 두고 `_apply_min_size` 가 크기를 맞춘 뒤 `_reveal` 로 보인다(큰 창→작은 창 깜빡임 방지).
  계산이 실패해도 `START_REVEAL_MS`(2초) 뒤에는 반드시 보이게 함. 실측: 1000×1000 → 1000×946 로 맞춘 뒤 표시 (역본2 줄 +38, 하단 임시 실행 체크 +16).
  창 틀을 없앤 뒤로는 크기가 곧 창 전체(위쪽 틀 없음) — 실측 1000×950 (상태 부연 설명 줄·제목 줄 여백 변화).
  첫 (투명) 창 높이 `START_MEASURE_HEIGHT` 는 왼쪽 패널보다 커야 함 — 낮으면 템플릿 목록이 0줄로 눌려 unmapped 가 되고
  `_apply_min_size` 가 재지 못해 최소 크기 없이 그대로 보인다 (역본2 줄 추가 때 820 에서 실제로 발생).

## 11-1. 보안 규칙 (2026-10-07 점검, v1.0.3 기능 추가 후 재점검)
- COM 으로 연 파일은 **보호된 보기 없이** 열리고, 자동화 기본값(`AutomationSecurity=1`)은 매크로를 허용한다(실측).
  → `PptEngine` 이 작업 동안 `AutomationSecurity=3`(강제 비활성)으로 바꾸고, `DisplayAlerts` 와 함께 종료 시 원래 값으로 복원.
- 외부 프로그램 실행은 `subprocess` 리스트 인자만(`shell=True` 금지). Bible2PPT 경로는 **`.exe` 만 허용**
  (.bat/.cmd 는 cmd.exe 가 인자를 재해석해 `&` 등으로 명령 주입 가능). `-b` 값은 `safe_bible_arg` 로 `-` 시작·제어 문자 거부.
  구절은 `_REF_RE` 로 검증돼 옵션처럼 시작할 수 없다.
- Bible2PPT stdout 이 알려준 경로는 지정한 출력 폴더 안의 파일만 인정.
- 출력 파일명은 사용자가 정한 형식이라도 `sanitize_filename`(경로 구분자·제어 문자 제거, 이름 `MAX_FILENAME_STEM`=150자 제한,
  예약어 회피 — CON/NUL/COM¹/CONIN$ 등, `con.txt` 처럼 확장자가 붙어도 장치로 해석되므로 첫 `.` 앞부분으로 판단)으로만 만든다.
- `worship_builder.output_conflict`: 출력 폴더가 템플릿·기도·찬송가·CCM 폴더와 같거나 그 안이면 거부(파일 이름 형식을 바꿀 수 있어
  이번 생성에 쓰지 않는 원본도 같은 이름으로 덮어쓸 수 있음), 결과 경로가 이번 원본·설교 템플릿·성경 템플릿과 같거나
  심볼릭 링크면 거부. 출력 폴더 지정(📁) 때와 생성 전(UI·`build`) 모두 확인.
- 성경 템플릿(`-t`)은 존재하는 **.pptx 만** (`worship_builder.bible_template`, `Bible2PPT.generate`, 설정 저장).
  Bible2PPT 는 자기 PowerPoint 세션에서 열어 이 프로그램의 `AutomationSecurity` 가 미치지 않으므로 .pptm 등 매크로 형식 금지.
  (생성 때는 `PptEngine.hold()` 로 같은 PowerPoint 를 쓰게 되어 매크로 차단이 미치지만, 말씀 열기·사용자가 켠 PowerPoint 등
  그렇지 않은 경우가 있으므로 이 제한은 그대로 둔다.)
- 창 숨기기(`_FrameHider`)는 다른 프로세스 창을 건드리므로: PowerPoint 문서 창 클래스(`PPTFrameClass`)만, NewWindow·Activate·
  Close·Quit 동안만. 블록 동안 우연히 숨겨진 관계없는 창(제목이 작업 창·빈 PowerPoint 창이 아닌 것)은 블록이 끝나면 다시 보이게
  하고(`keep`, `_is_title_of`), 투명하게 바꾼 창 스타일은 항상 되돌린다.
- 결과 창을 앞으로 가져올 때 `AllowSetForegroundWindow` 는 그 PowerPoint 프로세스 ID 에만 (ASFW_ANY 는 아무 프로세스나
  포커스를 가져갈 수 있게 열어 둠).
- `os.startfile` 로 폴더를 열 때는 `is_dir()` 확인 후에만 (설정이 실행 파일을 가리키게 조작되면 그 파일이 실행됨).
  PPT 를 열 때(`App._open_ppt`)는 확장자가 `.pptx/.pptm/.ppt` 이고 실제 파일일 때만.
- 템플릿 삭제는 영구 삭제하지 않고 휴지통으로(되돌릴 수 있게), 삭제 전 확인 창 필수.
- 임시 파일 열기·저장은 임시 폴더 바로 안의 실제 .pptx(링크 아님)만. 저장 대상도 `output_conflict` 로 원본 덮어쓰기 차단.
- 임시 파일 정리(영구 삭제)는 `temp_dir_conflict` 를 통과할 때만 — 임시 폴더가 정션·심볼릭 링크이거나 원본·출력 폴더·템플릿
  파일이 그 안에 있으면 하지 않는다 (§9-2). 임시 폴더는 %LOCALAPPDATA% 라 사용자 본인만 접근(같은 사용자의 다른 프로세스가
  넣은 파일은 이 프로그램과 같은 신뢰 수준 — 막을 수 없음).
- 텍스트 파일 불러오기는 `MAX_TEXT_IMPORT_BYTES`(1MB) 이하만. `bible2ppt.bible_list` 의 문자열이 아닌 값은 시작 시 버림.
- 남은 위험(코드로 막을 수 없음): COM 으로 연 PPT 는 보호된 보기 없이 열린다 → 매크로는 막지만 PowerPoint 자체 취약점을 노린
  파일은 막지 못함. 출처를 알 수 없는 PPT 는 템플릿·찬송가·CCM 폴더에 넣지 않도록 안내. 빌드 패키지 버전은 `>=` 라 고정되지 않음.
- config.json / presets.json 은 외부에서 수정될 수 있는 입력으로 취급: 기본값과 타입이 다른 값은 무시, 잘못된 프리셋/항목은 버림,
  `options` 는 문자열 값만 유지.
- 설정 마법사(2026-10-10 점검): **[마침]에서 모든 단계를 다시 검사**한 값만 저장. 단계마다 [다음]에서만 검사하던 때는
  숨긴 쪽 입력칸에 포커스가 남아(grid_remove 해도 포커스는 유지) 확인 단계에서 친 `.bat` 경로가 검사 없이 저장됐음(실측) →
  단계를 넘길 때 포커스를 다이얼로그로 옮기고, 저장 직전 전체 재검사. (실행은 `Bible2PPT._run` 의 `.exe` 검사가 막아 실제
  명령 실행까지는 안 됐음 — 실행 직전 검사는 그대로 유지할 것.) 쪽을 숨기는 다이얼로그를 만들 때 같은 함정 주의.
- 생성 중 화면은 덮기만 하므로 Tab 으로 가려진 입력칸에 포커스가 갈 수 있었음 → 화면에 포커스 + Tab·Shift-Tab 차단.
- 폴더 하위 검색(`list_ppt_files(recursive)`·`count_ppt_files`, `_walk_files`)은 심볼릭 링크·**정션** 폴더에 들어가지 않는다
  (`_is_link_dir` — reparse tag 가 MOUNT_POINT 인 것만, OneDrive 폴더 같은 다른 재분석 지점은 들어감). 예전 `rglob` 은 정션을
  따라가 자기 자신을 가리키는 정션 2개에 10초 넘게 15000개 이상을 돌며 화면 스레드(생성 전 점검·CCM 목록)를 멈췄음(실측).
- `find_ccm` 은 값(프리셋·텍스트 파일 = 외부 입력)이 드라이브·절대 경로·`..` 이면 `resolve` 전에 거른다 — `resolve` 가
  `\\서버\공유` 를 열려고 접속해(실측 21초 멈춤) Windows 로그인 정보(NTLM)가 그 서버로 갈 수 있음. 외부 입력을 경로로 만들 때는
  `resolve`/`is_file` 전에 같은 검사를 할 것.
- `template_fill_slots` 는 슬라이드 XML 합계 `_MAX_TOTAL_XML`(100MB) 초과면 읽지 않음(압축 폭탄으로 화면 멈춤 방지).
- 설정·프리셋 파일을 쓸 수 없으면(쓰기가 막힌 프로그램 폴더) `ConfigManager.save` 는 `save_error` 만 남기고 계속, 시작 후
  한 번 경고(`App._after_reveal`). 예전에는 시작하자마자 PermissionError 로 꺼졌음. 프리셋 저장·삭제 실패는 에러 창.
  `bible2ppt.lines` 는 기본값이 None 이라 `_deep_merge` 타입 검사를 받지 않으므로 시작 시 0~9 정수가 아니면 None.
- 날짜는 생성 때 `YYYY-MM-DD` 로 맞춰 고친다 (`strptime` 은 `2026-1-5` 도 받지만 `{년}`·`{월}`·`{일}`·`{요일}` 이 빈 값이 됐음).
- 템플릿 삭제는 `FOF_WANTNUKEWARNING` — 휴지통이 없는 USB·네트워크 드라이브면 영구 삭제 전에 Windows 가 묻는다 (이 PC 에서는
  일반 드라이브 휴지통 이동만 확인, 휴지통 없는 드라이브는 아직 확인 안 함). 템플릿 불러오기는 같은 이름이 링크면 거부.
- 말씀 미리보기 폴더(`%TEMP%\ChurchPPTMaker_말씀`) 정리는 폴더가 정션·링크면 하지 않고, 지울 수 없는 파일이 있어도 나머지는 계속.
- [Bible2PPT 받기 ↗]는 고정 주소(`BIBLE2PPT_URL`)만 연다 — 설정·입력값으로 주소를 만들지 말 것.

## 12. 작업 규칙
- 사용자가 쓰는 기능·규칙·설정을 바꾸면 README.md 와 `help_content.py`(프로그램 안 도움말)를 함께 고친다.
- UI 배치를 바꿀 때는 §4 초안 구조(순서 패널 = 프리셋 행 / 순서 목록 / 팔레트 행)를 유지한다.
- 사용자 메시지는 한국어. 모든 입력 검증 실패는 `messagebox.showerror`로 알린다.
- COM 호출은 반드시 작업 스레드에서, UI 갱신은 `after()`로 메인 스레드에서.
