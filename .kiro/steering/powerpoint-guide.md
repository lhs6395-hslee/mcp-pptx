# PowerPoint Generation System - Complete Specification

**Version**: v5.3 (2026-03-06)
**Purpose**: 이 문서 시리즈만으로 AI가 모든 Python 파일을 100% 동일하게 재생성할 수 있는 완전한 시스템 명세

## Related Documents

| File | Contents |
|------|----------|
| `powerpoint-guide.md` | 아키텍처, 스티어링 포맷, 레이아웃 참조 (이 문서) |
| `powerpoint-code-generate.md` | `generate.py` + `generate_ppt.sh` 소스코드 |
| `powerpoint-code-cover-toc.md` | `powerpoint_cover.py` + `powerpoint_toc.py` 소스코드 |
| `powerpoint-code-content.md` | `powerpoint_content.py` 소스코드 Part 1 (유틸리티 + 레이아웃 1~13 + phased_columns) |
| `powerpoint-code-content-2.md` | `powerpoint_content.py` 소스코드 Part 2 (다이어그램 헬퍼 + 레이아웃 14~40 + 라우터) |

---

## System Architecture

```
[Steering File]     →  [generate.py]      →  [Rendering Modules]  →  [.pptx]
rayhli-eks_guide_2026.py      orchestration         powerpoint_cover.py
ss_db_migration_      (template copy,        powerpoint_toc.py
resume.py              section removal,       powerpoint_content.py
(data only)            slide management)      (41 layout renderers)
```

## Dependencies

```bash
pip install python-pptx lxml pillow
```

- `python-pptx`: PowerPoint 생성/수정
- `lxml`: XML 파싱 (섹션 제거용)
- `pillow` (PIL): 이미지 종횡비 계산 (선택사항)
- `duckduckgo_search`: 이미지 검색 (선택사항, 현재 비활성)

## File Structure

```
ppt-mcp/
├── generate_ppt.sh              # Shell wrapper (one-line 실행)
├── rayhli-*.py                  # Steering files (data only — 영구 보존)
│
├── code/                        # ← Python 엔진 파일 (MD에서 재생성 가능, 휘발성)
│   ├── generate.py              #   Orchestration script
│   ├── powerpoint_content.py    #   41 layout renderers + utility functions
│   ├── powerpoint_cover.py      #   Cover slide renderer
│   └── powerpoint_toc.py        #   TOC slide renderer
│
├── icons/                       # PNG 아이콘 (512×512, 40종)
├── architecture/                # 아키텍처 다이어그램 PNG + draw.io 원본
│   ├── eks_architecture_wide.drawio   # draw.io 원본 (수정용)
│   ├── eks_architecture_wide.png      # draw.io → PNG (3405×665, AR=3.62)
│   └── karpenter_scheduling.png       # Karpenter 파드 스케줄링 (3600×1000, AR=3.60)
│   # ⚠️ detail_image/full_image의 search_q는 이 폴더 파일명(확장자 제외)과 정확히 일치해야 함
├── screenshots/                 # UI 스크린샷 PNG
├── template/
│   └── 2025_PPT_Template_FINAL.pptx  # PPT 템플릿 (13.33" × 7.50")
└── results/                     # 생성된 PPT 출력
```

> **code/ 폴더 원칙**: `code/` 안의 Python 파일은 MD 스티어링 파일로부터 재생성 가능한 휘발성 파일입니다.
> MD 파일이 항상 소스 오브 트루스 (Source of Truth)이며, Python을 직접 수정했다면 반드시 해당 MD도 동시에 업데이트해야 합니다.
>
> **실행 명령**: `python3 code/generate.py <steering_file>.py` 또는 `./generate_ppt.sh <steering_file>.py`

> **아키텍처 이미지 워크플로**: draw.io에서 편집 → `drawio --export --format png --scale 2 --output architecture/xxx.png architecture/xxx.drawio`

## Template Requirements

PPT 템플릿은 다음 슬라이드 구조를 가져야 함:
- **Index 0**: Cover slide (표지)
- **Index 1**: TOC slide (목차)
- **Index 7**: Body slide (본문 - 이 레이아웃을 복제하여 본문 슬라이드 생성)
- **Last slide**: Ending slide (감사합니다 - 보존됨)
- **Slide dimensions**: 13.333" × 7.500"

## Icons (40 Total)

```
analysis, aurora, auto_mode, aws_account, billing, chat, cicd, cli,
cluster_delete, config, console, container, cutover, dashboard, database,
deploy, dms, eks, eksctl, encryption, gitops, helm, k8s_version, kubectl,
kubernetes, load_balancer, microservices, migration, monitoring, network,
performance, pipeline, schema, security, server, service, storage,
terraform, timeline, verification
```

- Format: PNG, 512×512 pixels, transparent background
- Naming: lowercase, underscores for spaces (e.g., `load_balancer.png`)
- Location: `icons/` folder
- Fallback: 아이콘 파일 없으면 파란색 원형 표시

---

## Constants & Design System

### Fonts (FONTS dict)

| Key | Value | Usage |
|-----|-------|-------|
| `HEAD_TITLE` | "프리젠테이션 7 Bold" | 슬라이드 제목 (28pt) |
| `HEAD_DESC` | "프리젠테이션 5 Medium" | 슬라이드 설명 (12pt) |
| `BODY_TITLE` | "Freesentation" | 본문 제목/강조 |
| `BODY_TEXT` | "Freesentation" | 본문 텍스트 |

### Colors (COLORS dict)

| Key | RGB | Usage |
|-----|-----|-------|
| `PRIMARY` | (0, 67, 218) | 제목, 강조, 배지 |
| `BLACK` | (0, 0, 0) | 본문 텍스트 |
| `DARK_GRAY` | (33, 33, 33) | 진한 회색 텍스트 |
| `GRAY` | (80, 80, 80) | 설명글 |
| `BG_BOX` | (248, 249, 250) | 박스 배경 |
| `BG_WHITE` | (255, 255, 255) | 흰색 배경 |
| `BORDER` | (220, 220, 220) | 테두리 |
| `TERMINAL_BG` | (48, 10, 36) | 터미널 배경 (Ubuntu 보라색) |
| `TERMINAL_TITLEBAR` | (44, 44, 44) | 터미널 타이틀 바 |
| `TERMINAL_TEXT` | (102, 204, 102) | 터미널 텍스트 (초록) |
| `TERMINAL_COMMENT` | (150, 150, 150) | 터미널 주석 (회색) |
| `TERMINAL_RED` | (255, 95, 86) | macOS 빨강 버튼 |
| `TERMINAL_YELLOW` | (255, 189, 46) | macOS 노랑 버튼 |
| `TERMINAL_GREEN` | (39, 201, 63) | macOS 초록 버튼 |
| **Semantic Colors** | | |
| `SEM_RED` / `_BG` / `_TEXT` | (185,28,28) / (254,242,242) / (127,29,29) | 주의/필수 |
| `SEM_ORANGE` / `_BG` / `_TEXT` | (194,65,12) / (255,247,237) / (154,52,18) | 경고/핵심 |
| `SEM_GREEN` / `_BG` / `_TEXT` | (4,120,87) / (236,253,245) / (6,95,70) | 긍정/완료 |
| `SEM_BLUE` / `_BG` / `_TEXT` | (30,58,138) / (239,246,255) / (30,64,175) | 참조/조건 |
| `CALLOUT_BG` | (30, 58, 138) | 콜아웃 배경 (진한 파랑) |
| `CALLOUT_TEXT` | (219, 234, 254) | 콜아웃 본문 (밝은 파랑) |

### Layout Coordinates (LAYOUT dict)

| Key | Value | Description |
|-----|-------|-------------|
| `SLIDE_TITLE_Y` | 0.6" | 헤더 상단 |
| `SLIDE_DESC_Y` | 0.6" | 설명글 상단 |
| `BODY_START_Y` | 2.0" | 본문 시작점 |
| `BODY_LIMIT_Y` | 7.2" | 본문 한계선 |
| `MARGIN_X` | 0.5" | 좌우 여백 |
| `SLIDE_W` | 13.333" | 슬라이드 너비 |

---

## Steering File Format

steering file은 `presentation_data` 딕셔너리 하나만 정의하는 순수 데이터 파일입니다.

```python
# -*- coding: utf-8 -*-
"""Presentation Data File"""

presentation_data = {
    "cover": {
        "title": "프레젠테이션 제목",
        "subtitle": "부제목"
    },
    "sections": [
        {
            "section_title": "1. 섹션 제목",
            "slides": [
                {
                    "l": "layout_name",    # 레이아웃 종류
                    "t": "슬라이드 제목",   # 헤더 좌측
                    "d": "슬라이드 설명",   # 헤더 우측
                    "data": {              # 레이아웃별 데이터
                        # ...
                    }
                }
            ]
        }
    ]
}
```

### Data Nesting Pattern

모든 레이아웃은 `data.data.data` 3중 중첩 구조:
- Level 1: `slide_info` (t, d, l, data)
- Level 2: `wrapper = data.get('data', {})` → body_title, body_desc, data
- Level 3: `content = wrapper.get('data', {})` → 실제 콘텐츠

예외: `challenge_solution`, `before_after` — Level 2에서 직접 데이터 읽음

### 개조식(Bullet List) 자동 처리 규칙

`create_content_box`, `render_3_cards`, `render_before_after` 렌더러는 body에 **여러 줄(`\n`)이 있으면 자동으로 `•`를 앞에 추가**합니다.

- 데이터에 `•`를 명시하지 않아도 됨
- 이미 `•`로 시작하는 줄은 중복 추가하지 않음
- 실제 번호 목록 패턴(`1. `, `2) `, `3: `)은 `•` 추가하지 않음
- 단일 줄 body(단락 형태)는 `•` 추가하지 않음

**올바른 데이터 작성:**
```python
"body": "IAM 최소 권한 원칙 적용\nVPC 네트워크 격리\nKMS 암호화"
# → 렌더링: "• IAM 최소 권한 원칙 적용 / • VPC 네트워크 격리 / • KMS 암호화"
```

---

## ⚠️ 데이터 키 불일치 주의사항 (CRITICAL)

> 스티어링 파일 작성 시 가장 많이 발생하는 오류 유형.
> 아래 목록은 실제 발생한 사례 기반입니다.

### 예외 1: wrapper 레벨 렌더러 (2종)

`challenge_solution`과 `before_after`는 `content = wrapper` (Level 3 생략).
실제 데이터 키를 `data['data']`(wrapper) 안에 직접 넣어야 함.

**올바른 구조 — challenge_solution:**
```python
"data": {
    "body_title": "...",
    "body_desc": "...",
    "challenge": {          # ← wrapper 바로 아래 (data.data가 아님)
        "title": "...",
        "body": "..."
    },
    "solution": {           # ← wrapper 바로 아래
        "title": "...",
        "body": "..."
    }
}
```

**❌ 잘못된 구조 (data.data.data로 넣으면 CHALLENGE/SOLUTION 레이블만 표시됨):**
```python
"data": {
    "body_title": "...",
    "data": {               # ← 이 레벨 추가하면 안 됨
        "challenge": {...},
        "solution": {...}
    }
}
```

**올바른 구조 — before_after:**
```python
"data": {
    "body_title": "...",
    "body_desc": "...",
    "before_title": "...",  # ← wrapper 바로 아래
    "before_body": "...",
    "after_title": "...",
    "after_body": "..."
}
```

---

### 예외 2: icon_grid — `icon` 키 (search_q 아님)

`icon_grid` 렌더러는 `item.get('icon', '')` 사용. `search_q` 키를 쓰면 빈 문자열 → 파란 원만 표시됨.

**올바른 구조:**
```python
"items": [
    {"icon": "eks", "title": "...", "desc": "..."},   # ← icon 키
    {"icon": "aurora", "title": "...", "desc": "..."},
]
```

**❌ 잘못된 구조:**
```python
{"search_q": "eks", "title": "...", "desc": "..."}  # ← search_q 쓰면 안 됨
```

---

### 예외 3: quote_highlight — author에 "—" 포함 금지

렌더러가 `f"— {author}"` 형식으로 자동 추가. author에 "—"를 포함하면 "— — 이름"이 됨.

**올바른 구조:**
```python
"author": "김지훈 CTO"              # ← 대시 없이 이름만
```

**❌ 잘못된 구조:**
```python
"author": "— 김지훈 CTO"            # ← 이중 대시 발생
```

---

### 예외 4: funnel — stages 키명 (title이 아닌 label/value)

렌더러는 `stage.get('label', '')`, `stage.get('value', '')`, `stage.get('desc', '')` 사용.
`title` 키를 쓰면 단계명이 표시되지 않음.

**올바른 구조:**
```python
"stages": [
    {"label": "인지 단계", "value": "500개 사", "desc": "AWS 도입 검토 시작", "color": "blue"},
    {"label": "관심 단계", "value": "280개 사", "desc": "PoC 제안서 접수",    "color": "green"},
    {"label": "검토 단계", "value": "120개 사", "desc": "MRA 평가 수행",       "color": "orange"},
    {"label": "계약 단계", "value": "48개 사",  "desc": "계약 체결 및 착수",    "color": "red"},
    {"label": "완료",      "value": "31개 사",  "desc": "안정 운영 중",         "color": "primary"},
]
```

> ⚠️ `color` 미지정 시 모든 단계가 `primary`(파란색 단일)로 렌더링됨.
> 단계별 색상 다양화 권장: blue → green → orange → red → primary

**❌ 잘못된 구조:**
```python
{"title": "인지 단계", "desc": "500개 사 — AWS 도입 검토"}  # ← title 쓰면 단계명 미표시, color 없으면 단색
```

---

### 전체 예외 목록 요약

| Layout | 예외 유형 | 주의 키 |
|--------|----------|---------|
| `challenge_solution` | wrapper 레벨 (data 한 단계 생략) | challenge{title,body}, solution{title,body} |
| `before_after` | wrapper 레벨 (data 한 단계 생략) | before_title, before_body, after_title, after_body |
| `icon_grid` | 아이콘 키명 다름 | `icon` (NOT `search_q`) |
| `quote_highlight` | author 자동 "— " 추가 | author에 "—" 포함 금지 |
| `funnel` | stages 키명 다름 | `label`, `value` (NOT `title`) |

---

## Available Layouts (41종: 고유 38종 + alias 3종)

| # | Layout | Code | Data Keys | 비고 |
|---|--------|------|-----------|------|
| 1 | Bento Grid | `bento_grid` | main, sub1, sub2 | 좌 50% + 우 2분할 |
| 2 | Three Cards | `3_cards` | card_1, card_2, card_3 | 아이콘+제목+본문 |
| 3 | Grid 2×2 | `grid_2x2` | item1, item2, item3, item4 | compact 모드 |
| 4 | Quad Matrix | `quad_matrix` | (grid_2x2와 동일) | `grid_2x2` alias |
| 5 | Process Arrow | `process_arrow` | steps[]{title,body,search_q} | 쉐브론+본문 박스 |
| 6 | Phased Columns | `phased_columns` | steps[]{title,body,search_q} | 단계별 컬럼+그라데이션 |
| 7 | Timeline | `timeline_steps` | steps[]{date,desc} | 숫자 배지+카드. date에 `\n` 사용 가능 |
| 8 | Challenge/Solution | `challenge_solution` | challenge{title,body}, solution{title,body} | 좌우+화살표 (wrapper 레벨). points[] 사용 불가, body 문자열 필수 |
| 9 | Comparison VS | `comparison_vs` | item_a_title/body, item_b_title/body | VS 원형 |
| 10 | Comparison Table | `comparison_table` | columns[], rows[] | 3열 표 |
| 11 | Detail Image | `detail_image` | title, body, search_q | 상단 텍스트+하단 이미지. **`search_q`는 반드시 `architecture/` 폴더에 실제 파일이 있는 이름** (없으면 fallback 박스) |
| 12 | Image Left | `image_left` | image_path, bullets[] | 좌 이미지+우 불릿 |
| 13 | Architecture Wide | `architecture_wide` | col1, col2, col3 | 상단 다이어그램+하단 3열 |
| 14 | Key Metric | `key_metric` | (3_cards와 동일) | `3_cards` alias |
| 15 | Detail Sections | `detail_sections` | overview, highlight, condition, diagram | 좌 멀티섹션+우 다이어그램 |
| 16 | Table Callout | `table_callout` | columns[], rows[], callout{icon,title,body} | 테이블+추천박스. callout은 반드시 dict |
| 17 | Full Image | `full_image` | image_path/search_q, caption | 풀와이드 이미지. 아키텍처 다이어그램은 **가로 비율 ≥ 2.5:1** 이미지 사용 권장 (슬라이드 본문 영역 ≈ 13.3" × 5.2", AR ≈ 2.56) |
| 18 | Before/After | `before_after` | before_title/body, after_title/body | 전/후 비교 (wrapper 레벨) |
| 19 | Icon Grid | `icon_grid` | items[]{**icon**,title,desc} | 3열×N행 아이콘 그리드. **`search_q` 아님, 반드시 `icon` 키** |
| 20 | Numbered List | `numbered_list` | items[]{title,desc} | 번호형 세로 리스트 |
| 21 | Stats Dashboard | `stats_dashboard` | metrics[]{value,unit,label,desc} | KPI 대형 숫자 |
| 22 | Quote Highlight | `quote_highlight` | quote, author, role | 인용문 강조. **author에 "—" 포함 금지** (렌더러가 자동 추가) |
| 23 | Pros & Cons | `pros_cons` | subject, pros[], cons[] | 장단점 비교 |
| 24 | Do / Don't | `do_dont` | do_items[], dont_items[] | Best Practice |
| 25 | Split Text+Code | `split_text_code` | description, bullets[], code_title, code | 설명+코드 병렬. **코드 14줄 초과 시 슬라이드 분리 필수** — 분리 규칙은 아래 참조 |
| 26 | Pyramid Hierarchy | `pyramid_hierarchy` | levels[]{label,desc,color} | 피라미드 계층 |
| 27 | Cycle Loop | `cycle_loop` | steps[]{label,desc}, center_label | 순환 프로세스 |
| 28 | Venn Diagram | `venn_diagram` | circles[]{label,desc,color}, center_label | 좌측 3원 벤 + 우측 설명 카드 |
| 29 | SWOT Matrix | `swot_matrix` | quadrants[]{label,title,items[],color} | 2×2 SWOT 분석 |
| 30 | Center Radial | `center_radial` | center{label,desc}, directions[]{label,desc,color} | 중심 ROUNDED_RECT + 4방향 카드 |
| 31 | Funnel | `funnel` | stages[]{**label**,**value**,desc,color} | 퍼널 다이어그램. **`title` 아님, 반드시 `label`+`value`** |
| 32 | Zigzag Timeline | `zigzag_timeline` | steps[]{date,title,desc,color} | 지그재그 타임라인 |
| 33 | Fishbone Cause-Effect | `fishbone_cause_effect` | effect, categories[]{label,causes[],color} | 피쉬본 원인-결과 |
| 34 | Org Chart | `org_chart` | root{label,desc}, children[]{label,desc,items[],color} | 조직도/트리 |
| 35 | Temple Pillars | `temple_pillars` | roof{label}, pillars[]{label,desc,color}, foundation{label} | 기둥형 구조도 |
| 36 | Infinity Loop | `infinity_loop` | left_loop[],right_loop[],left_label,right_label,center_label | 무한 순환 루프 |
| 37 | Speedometer Gauge | `speedometer_gauge` | value,segments[]{label,color},title | 스피도미터 게이지 |
| 38 | Mind Map | `mind_map` | center{label}, branches[]{label,sub_branches[],desc,color} | 좌측 방사형 맵 + 우측 설명 카드 |
| 39 | Checklist 2-Column | `checklist_2col` | items[]{title, status(done/in_progress/todo), subitems[]{text,badge}} | 2열 체크리스트 그리드 |
| 40 | Kanban Board | `kanban_board` | columns[]{title, color(gray/orange/green), cards[]{title,badge}} | 열당 카드 칸반 보드 |
| 41 | Executive Summary | `exec_summary` | sections[]{label, body, color(gray/blue/green/orange/red)} | 전체 폭 레이블+본문 섹션 |

### Detail Sections Diagram Types

`detail_sections` 레이아웃의 우측 다이어그램은 4가지 type 지원:

| Type | Description | Data Key |
|------|-------------|----------|
| `flow` | 수직 박스+화살표 흐름도 (기본값) | items[]{label} — label에 `\n` 사용해 2줄 표시 가능 |
| `layers` | 수평 계층 다이어그램 | layers[]{title,desc,color,items[]} |
| `compare` | 좌우 비교 다이어그램 | sides[] |
| `process` | 좌→우 가로 프로세스 | steps[] |

> ⚠️ `flow` 타입 items는 반드시 `label` 키 사용. `title`/`body` 키는 무시됨.

### Phased Columns Gradient Palette

7-step gradient (dark navy → light blue):
```python
[(0,27,94), (0,45,143), (0,67,218), (59,122,237), (123,167,247), (160,195,250), (190,215,252)]
```
N개 컬럼에 대해 균등 샘플링하여 색상 배정.

### Semantic Box Styles (\_SEM\_BOX\_STYLES)

다이어그램/detail_sections에서 사용하는 의미 기반 박스 스타일:

| Key | Fill | Line | Text |
|-----|------|------|------|
| `gray` | (248,249,250) | (150,150,150) | (33,33,33) |
| `red` | (254,242,242) | (185,28,28) | (127,29,29) |
| `orange` | (255,247,237) | (194,65,12) | (154,52,18) |
| `green` | (236,253,245) | (4,120,87) | (6,95,70) |
| `blue` | (239,246,255) | (30,58,138) | (30,64,175) |
| `primary` | (239,246,255) | (0,67,218) | (30,64,175) |

---

## Generation Flow

1. `generate.py`가 steering file의 `presentation_data`를 `exec()`로 로드
2. 템플릿 복사 → 섹션 제거 (`remove_all_sections`)
3. Cover slide 업데이트 (`powerpoint_cover.update_cover_slide`)
4. TOC slide 업데이트 (`powerpoint_toc.update_toc_slide`)
   - section_title에서 `^\d+\.\s*` prefix 제거 후 전달
5. 각 section의 각 slide에 대해:
   - Body layout 복제하여 새 슬라이드 생성
   - `set_slide_title_area()`로 헤더 설정
   - `render_slide_content()`로 본문 렌더링 (layout→renderer 라우팅)
6. 불필요한 템플릿 슬라이드 삭제 (keeper_ids 기반)
7. Ending slide를 마지막으로 이동
8. 저장 → `results/{steering_basename}.pptx`

---

## ⚠️ 텍스트 렌더링 3대 금지 규칙 (CRITICAL — 생성·검증 필수)

> **이 규칙을 어기면 PPT를 다시 생성해야 합니다.**
> 데이터 파일 작성 시 사전 준수 + 생성 후 반드시 검증.

### 규칙 1: 슬라이드 타이틀 단어 잘림 금지

**증상**: 표지·섹션 제목·슬라이드 제목에서 단어가 중간에 끊겨 다음 줄로 넘어감
**원인**: `set_slide_title_area()`가 4.5인치 영역에 고정 렌더링 → 긴 타이틀 = 단어 경계 무시 자동 줄바꿈

**렌더러 자동 폰트 축소 (v5.3~)**:
`set_slide_title_area()`가 타이틀 글자 수에 따라 자동으로 폰트를 줄입니다:
- ≤ 18자: **28pt** (기본)
- 19~23자: **25pt**
- 24자 이상: **22pt**

→ 데이터에서 타이틀을 억지로 `\n`으로 쪼갤 필요 없음. 단, **24자 이내**로 작성 권장 (가독성).

**데이터 파일 작성 기준**:
- 슬라이드 제목 `"t"` 영문 기준 **최대 23자** (폰트 자동 축소 없이 1줄 유지)
- 24자 이상 시 22pt로 축소되어 표시됨 (허용범위)
- 예: `"1-1. 프로젝트 전체 로드맵"` ✓ / `"4-1. Challenge & Solution"` (25자, 22pt 적용) ✓

**검증 스크립트**:
```python
from pptx import Presentation
prs = Presentation('results/output.pptx')
for i, slide in enumerate(prs.slides):
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        text = shape.text_frame.text.strip()
        if not text or '\n' in text: continue
        if len(text) > 30:
            print(f'⚠️  Slide {i+1} LONG TITLE [{len(text)}자]: {text}')
```

---

### 규칙 2: 본문 텍스트 박스 오버플로우 금지

**증상**: 본문 텍스트가 박스 경계를 넘어 다음 도형 위에 겹쳐 보임
**원인**: 레이아웃별 박스 높이가 고정되어 있어 텍스트가 많으면 박스 밖으로 넘침

#### 레이아웃별 텍스트 한도 (이 기준을 데이터 파일 작성 시 반드시 준수)

| 레이아웃 | 한도 | 세부 기준 |
|----------|------|----------|
| `zigzag_timeline` | **date 1줄 + title 1줄 + desc ≤ 4줄** | 각 줄 한글 ~8자, 영문 ~15자 이내 |
| `stats_dashboard` | **value+unit 1줄 + desc ≤ 2줄** | desc 3줄 이상 시 반드시 단축 |
| `3_cards` | **card body ≤ 8줄** | 카드당 8줄 초과 시 슬라이드 추가 |
| `numbered_list` | **item title 1줄 + desc ≤ 2줄** | 5개 항목 시 desc 1줄 권장 |
| `comparison_table` | **셀당 ≤ 2줄** | 4열 이상 시 셀 폭 좁아짐, 1줄 권장 |
| `detail_sections` | **overview ≤ 2줄, highlight ≤ 3줄, bullets ≤ 3개** | 각 섹션 높이 비율 30/45/25% 고정 |
| `bento_grid` | **main ≤ 10줄, sub ≤ 6줄** | 빈 줄(\n\n) 포함 시 줄 수에 합산 |
| `grid_2x2` | **셀 body ≤ 6줄** | 셀 높이 약 2.5" |
| `phased_columns` | **단계 body ≤ 5줄** | 5단계 시 열 폭 좁아짐, 4줄 권장 |
| `checklist_2col` | **subitems ≤ 3개/item, text ≤ 1줄** | 6개 item 이하, 각 text 1줄 |
| `kanban_board` | **cards ≤ 5개/열, title ≤ 2줄** | 카드 높이 자동계산, title 2줄 초과 금지 |
| `exec_summary` | **sections ≤ 5개, body ≤ 2줄/섹션** | body 3줄 이상 시 섹션 추가 분리 |
| `split_text_code` | **코드 ≤ 14줄** | 14줄 초과 시 슬라이드 분리 — 분리 규칙 아래 참조 |

#### split_text_code 슬라이드 분리 규칙 (CRITICAL)

코드가 14줄을 초과하면 슬라이드를 2장(이상)으로 수동 분리합니다.

**슬라이드 분리 시 필수 사항**:
1. **코드를 논리 단위로 분할** — 단순히 줄 수로 자르지 않고, 의미 있는 블록 경계(함수, 리소스, 섹션 주석 등)에서 분리
2. **각 슬라이드의 description/bullets는 해당 슬라이드의 코드 내용에 맞게 독립적으로 작성** — 슬라이드 1의 내용을 슬라이드 2에 그대로 복사하는 것은 절대 금지
3. **슬라이드 제목에 순서 표기** — `code_title`은 `"eks_cluster.tf (1/2)"`, `"eks_cluster.tf (2/2)"` 형식

**올바른 분리 예시**:
```python
# ✅ 슬라이드 1: 기본 설정 코드 (14줄 이내)
{
  "l": "split_text_code",
  "t": "Terraform EKS 클러스터 — 기본 설정",
  "data": {
    "data": {
      "description": "EKS 클러스터 기본 구성: 네트워크, 버전, VPC 설정",
      "bullets": ["VPC와 서브넷 연결", "Kubernetes v1.29 지정", "클러스터명 prod-eks-cluster"],
      "code_title": "eks_cluster.tf (1/2)",
      "code": "# 기본 클러스터 설정 블록\n..."
    }
  }
}

# ✅ 슬라이드 2: 노드 그룹 코드 (14줄 이내)
{
  "l": "split_text_code",
  "t": "Terraform EKS 클러스터 — 노드 그룹",
  "data": {
    "data": {
      "description": "노드 그룹 설정: On-Demand/Spot 혼합 비율과 Autoscaler 정책",
      "bullets": ["On-Demand 30% + Spot 70% 구성", "min=3 / max=50 자동 확장"],
      "code_title": "eks_cluster.tf (2/2)",
      "code": "# 노드 그룹 설정 블록\n..."
    }
  }
}
```

**❌ 잘못된 분리 예시 (description 복사)**:
```python
# 슬라이드 2에서 슬라이드 1과 동일한 description/bullets를 그대로 복사 — 절대 금지
"description": "EKS 클러스터는 Terraform 모듈로 선언형 관리합니다."  # 슬라이드 1과 동일 → 무의미
```

**오버플로우 발생 시 처리 순서**:
1. 텍스트 단축 (수식어·반복 표현 제거)
2. 단축 불가 시 슬라이드를 **1장 추가**하여 내용 분리
3. 절대 텍스트를 박스 밖으로 내보내지 않음

---

### 규칙 3: 슬라이드 경계 이탈 금지

**증상**: 텍스트 박스·도형·이미지가 슬라이드(13.333" × 7.5") 바깥으로 나감
**원인**: `calculate_dynamic_rect()`가 반환하는 `bh`를 초과하는 높이로 요소를 그리거나, 마지막 요소의 y좌표 + height > 7.2" 초과

**데이터 파일 작성 기준**:
- 항목 수가 많을수록 한 슬라이드에서 처리하려 하지 말 것
- `numbered_list` 항목 ≤ 7개, `3_cards` 카드 body ≤ 8줄
- 슬라이드 세로 경계(`BODY_LIMIT_Y` = 7.2") 초과 여지가 있으면 슬라이드를 분리

**검증 기준**:
```
모든 도형의 top + height ≤ 7.2인치 (BODY_LIMIT_Y)
슬라이드 너비 초과: left + width ≤ 13.333인치 (SLIDE_W)
```

---

### 규칙 4: 개조식 body 텍스트 — `•` 자동 추가 동작

렌더러가 body에 여러 줄(`\n`)이 있으면 **자동으로 `•`를 추가**합니다. 데이터에 `•`를 수동으로 넣지 않아도 됩니다.

적용 렌더러: `create_content_box`, `render_3_cards`, `render_before_after`

**자동 추가 규칙**:
- 다중 줄(`\n`) body → 각 줄 앞에 `•` 자동 추가
- 이미 `•`로 시작하는 줄 → 중복 추가 안 함
- 번호 목록 패턴(`1. `, `2) `, `3: `) → `•` 추가 안 함
- 단일 줄 단락 → `•` 추가 안 함

**Before/After 정렬**: body 텍스트는 좌측 정렬(`PP_ALIGN.LEFT`)로 렌더링. 기존에 가운데 정렬 버그 있었음 → v5.3에서 수정.

---

### 규칙 5: `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE` 오동작 주의 (CRITICAL)

**증상**: `auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE` 를 설정해도 텍스트가 여전히 박스 밖으로 넘침
**원인**: python-pptx의 `TEXT_TO_FIT_SHAPE`(값 2)는 XML에 `<a:normAutofit/>`을 씁니다. 이는 텍스트를 축소하는 게 아니라 **텍스트 넘침을 허용(overflow: visible)**하는 설정입니다. 동시에 `p.font.size = Pt(N)`로 명시적 폰트 크기를 설정하면 그 값이 우선되어 자동 축소가 전혀 일어나지 않습니다.

**핵심**: `auto_size = TEXT_TO_FIT_SHAPE`는 텍스트 오버플로우 방지 수단이 아닙니다. PowerPoint가 파일을 열 때 자체적으로 fontScale을 계산하여 축소할 수도 있지만, 명시적 `sz` 속성이 있으면 이를 무시합니다.

**올바른 오버플로우 방지법**: 박스 높이에 맞는 폰트 크기와 라인 수를 Python 코드에서 직접 계산합니다.
```
사용 가능 높이 = text_h - margin_top - margin_bottom
라인 당 높이  = font_size_pt + space_after_pt
최대 라인 수  = 사용 가능 높이(pt) / 라인당 높이
```
코드 박스(`create_terminal_box`) 기준:
- 사용 가능 높이 ≈ 216pt (text_h 3.4" 에서 상하마진 0.4" 제외)
- Pt(14) + Pt(6) → 최대 10줄 | Pt(11) + Pt(4) → 최대 14줄 | Pt(10) + Pt(3) → 최대 16줄 | Pt(9) + Pt(2) → 최대 19줄

**엔진 코드에서 `TEXT_TO_FIT_SHAPE` 유지 여부**: 일부 PowerPoint 버전에서 부분적으로 동작하는 경우가 있어 코드에 남겨두되, 신뢰하지 말 것. 실제 오버플로우 방지는 반드시 라인 수 계산으로 보장해야 합니다.

---

### 규칙 5: EMU 좌표는 반드시 정수 (CRITICAL)

**증상**: PowerPoint 복구 대화상자 ("프레젠테이션 복구가 시도될 수 있습니다.")
**원인**: python-pptx로 EMU 값을 계산할 때 `/` 연산자가 float을 반환, XML에 `y=3017520.0` 같은 소수점 좌표가 저장됨

**규칙**: 좌표·크기를 계산하는 모든 나눗셈은 `int()` 로 감쌀 것
```python
# ❌ 잘못된 예
sec_h = (bh - gap * (n - 1)) / n          # → float EMU → XML 손상
col_w = (bw - col_gap) / 2                 # → float EMU → 복구 대화상자

# ✅ 올바른 예
sec_h = int((bh - gap * (n - 1)) / n)     # 정수 EMU
col_w = int((bw - col_gap) / 2)            # 정수 EMU
# 또는 정수 나눗셈 연산자 사용:
sec_h = (bh - gap * (n - 1)) // n         # 정수 나눗셈
```

**적용 필수 패턴**:
- 높이 분할: `row_h = int(bh / n_rows)`, `sec_h = int(...)`, `card_h = int(min(...))`
- 너비 분할: `col_w = int(bw / n_cols)`, `sub_w = int((bw - gap) / 2)`
- 서브 항목 높이: `sub_item_h = int((parent_h - header_h) / n_items)`

---

### 규칙 6: Task 완료 상태 표시 정확성 (CRITICAL)

> **이 규칙을 어기면 PPT를 다시 생성해야 합니다.**
> 계획 PPT·완료 보고 PPT 모두 해당.

**원칙: ✅는 실제 완료된 Task에만 사용한다.**

#### ✅ 허용 — 실제 완료
- WBS에서 progress = 100% 로 확인된 Task
- Trello에서 dueComplete = true 로 확인된 Task
- 이전 주차에 이미 Done/아카이브 처리된 Task

#### ❌ 금지 — 미완료·예정
| 잘못된 표현 | 올바른 표현 |
|------------|------------|
| `"1.5 Aurora 파라미터 ✅ 완료"` (미완료) | `"1.5 Aurora 파라미터 🔄 완료 목표"` |
| `"Phase 1: 분석 ✅"` (일부 미완료 포함) | `"Phase 1: 분석 🔄"` |
| `"보안/NW 점검 🔄→✅"` (예정) | `"보안/NW 점검 🔄"` |
| `"Aurora 파라미터 확인 🔄→✅"` (예정) | `"Aurora 파라미터 확인 🔄"` |

#### 상태 이모지 기준
| 이모지 | 의미 | 조건 |
|--------|------|------|
| ✅ | 완료 | WBS 100% 또는 Trello dueComplete=true **확인 후** |
| 🔄 | 진행 중 / 완료 목표 | WBS 0~99% (이번 주 완료 예정 포함) |
| 🆕 | 신규 착수 | 이번 주 새로 시작하는 Task |
| 📅 | 예정 | 아직 시작 전 (To Do) |

#### 계획 PPT 작성 시 특별 주의
- **이월 Task**: 이전 주에서 미완료로 넘어온 경우 → `🔄 완료 목표` (✅ 금지)
- **Phase 완료 표시**: 해당 Phase의 모든 Task가 WBS 100%일 때만 `Phase N: 완료 ✅` 사용
- **일별 계획 desc**: 앞으로 할 일이면 `🔄 완료 목표` / 이미 한 일이면 `✅ 완료`

---

### 통합 검증 체크리스트 (생성 후 PPT 열어서 확인)

```
생성 즉시 PowerPoint/LibreOffice로 열어서 슬라이드별 확인:

□ 파일 열기: PowerPoint 복구 대화상자 없이 정상 열림 (float EMU 없음)
□ Cover  : 제목·부제목 단어 중간 잘림 없음
□ TOC    : 목차 항목이 슬라이드 안에 완전히 표시됨
□ 모든 슬라이드:
  □ 슬라이드 제목이 단어 경계에서 줄바꿈됨 (중간 잘림 없음)
  □ 본문 텍스트가 박스 안에 완전히 들어있음 (넘침 없음)
  □ 어떤 도형·텍스트도 슬라이드 하단(7.2") 아래로 나가지 않음
  □ 어떤 도형·텍스트도 슬라이드 오른쪽(13.3") 밖으로 나가지 않음
□ Ending : "Thank You" 슬라이드 정상 렌더링
□ 완료 상태 정확성 (계획·완료 보고 PPT):
  □ ✅ 표시 항목이 WBS 100% 또는 Trello dueComplete=true 확인된 Task인가?
  □ 이번 주 완료 목표(이월 포함)에 ✅ 사용 여부 → 🔄로 수정
  □ "🔄→✅" 패턴 사용 여부 → 단순 🔄로 수정 (예정이므로)
  □ Phase 완료 표시(Phase N ✅)가 해당 Phase 전체 Task 100% 기준인가?

문제 발견 시:
1. 해당 슬라이드의 steering 데이터 파일 수정 (텍스트 단축 또는 슬라이드 분리)
2. python3 generate.py <steering_file>.py 재실행
3. 재확인 반복
```

---

### Post-Generation: 검증 스크립트 (생성 즉시 실행 — 업로드 전 필수)

> **PPT 생성 후 반드시 아래 스크립트를 실행하고, 이상 없음을 확인한 후에만 업로드한다.**

```python
from pptx import Presentation

prs = Presentation('results/output.pptx')
SLIDE_W = int(13.33 * 914400)
SLIDE_H = int(7.5 * 914400)
issues_total = []

for i, slide in enumerate(prs.slides):
    slide_issues = []
    title_text = ""
    for s in slide.shapes:
        # 실제 슬라이드 타이틀: 상단 1.5" 이내, 너비 3~6" (body_desc 7"+ 제외)
        if (s.has_text_frame and s.top < int(1.5*914400)
                and int(3*914400) < s.width < int(6*914400)):
            t = s.text_frame.text.strip()
            if t and len(t) < 50:
                title_text = t

    for shape in slide.shapes:
        r = shape.left + shape.width
        b = shape.top + shape.height
        if r > SLIDE_W + 91440:
            slide_issues.append(f"  ❌ 우측 이탈: {shape.name}")
        if b > SLIDE_H + 91440:
            slide_issues.append(f"  ❌ 하단 이탈: {shape.name} (bot={b/914400:.2f}\")")

        if not shape.has_text_frame:
            continue
        full_text = shape.text_frame.text.strip()
        if not full_text:
            continue

        # 슬라이드 타이틀 한글 가중 길이 검사 (body_desc 제외)
        if (shape.top < int(1.5*914400)
                and int(3*914400) < shape.width < int(6*914400)
                and len(full_text) < 50):
            weighted = sum(2 if ord(c) > 0x3000 else 1 for c in full_text)
            if weighted > 36:
                slide_issues.append(f"  ⚠️ 타이틀 길이 초과: '{full_text}' (가중 {weighted})")

        # 텍스트 줄 수 vs 박스 높이
        box_h_pt = shape.height / 12700
        lines = [ln for p in shape.text_frame.paragraphs
                 for ln in p.text.split('\n') if ln.strip()]
        sizes = [r.font.size for p in shape.text_frame.paragraphs
                 for r in p.runs if r.font.size]
        avg_pt = (sum(s/12700 for s in sizes)/len(sizes)) if sizes else 10
        max_lines = box_h_pt / (avg_pt * 1.35)
        if len(lines) > max_lines + 1:
            slide_issues.append(
                f"  ⚠️ 오버플로우: {shape.name} "
                f"({len(lines)}줄 > 한계 {max_lines:.1f}줄, {avg_pt:.0f}pt)")

    if slide_issues:
        print(f"슬라이드 {i+1} [{title_text[:35]}]")
        for iss in slide_issues:
            print(iss)
        issues_total.extend(slide_issues)

print("✅ 이상 없음" if not issues_total else f"❌ {len(issues_total)}건 — 수정 후 재생성")
```

**슬라이드 타이틀 단어 잘림 판정 기준**:
- 실질 임계값: **~340pt** (4.5인치 영역, 실제 렌더링 기준)
- 가중 36자 초과 시 steering 파일의 `"t"` 값을 단축 (단어 경계에서만 개행)
- **body_desc** (`d` 필드, 너비 7"이상) 는 별도 박스라 검사 대상 아님
- **사전에 미리 자르지 말 것** — 생성 후 검증하여 초과분만 수정

---

## Layout Diversity Rule

- 최대 3장까지 같은 레이아웃 허용
- 단, 동일 주제/로직/다른 데이터(예: 주차별 일정)는 같은 레이아웃 허용
- 예: 1주차/2주차/3주차 작업 → 모두 `process_arrow` 사용 OK

---

## Utility Functions Summary

| Function | Description |
|----------|-------------|
| `draw_body_header_and_get_y()` | 본문 헤더(제목+설명) 그리고 시작 Y 반환 |
| `calculate_dynamic_rect()` | 남은 공간 (x, y, w, h) 계산 |
| `create_content_box()` | 만능 박스 (normal/compact/terminal 모드) |
| `create_terminal_box()` | macOS 스타일 터미널 박스 |
| `draw_icon_search()` | 로컬 아이콘 로드 (없으면 파란 원형) |
| `clean_body_placeholders()` | 본문 영역(2.0"~7.2") 기존 도형 제거 |
| `_place_image_centered()` | 이미지 비율 유지 중앙 배치 |
| `_diagram_box()` | 다이어그램용 의미색상 박스 |
| `_diagram_arrow_label()` | 화살표 라벨 (⬇/➡/⬅/⬆) |
| `_diagram_shape_arrow()` | 실제 화살표 shape |
