# PowerPoint 자동 생성 시스템 - Design

## Overview

스티어링 파일(Python 데이터) → 오케스트레이션 → 모듈별 렌더링 → .pptx 출력의 파이프라인 아키텍처.
모든 렌더링 모듈은 독립적이며, 공유 상수(FONTS, COLORS, LAYOUT)와 유틸리티 함수 체인으로 일관성을 보장한다.

## Architecture

```
[Steering File]     →  [generate.py]        →  [Rendering Modules]     →  [.pptx]
rayhli-eks_guide_2026.py      exec() 로드              powerpoint_cover.py
ss_db_migration_       템플릿 복사               powerpoint_toc.py
resume.py              섹션 제거                 powerpoint_content.py
(data only)            슬라이드 관리              (27 layout renderers)
```

### Generation Flow

```
1. exec(steering_file) → presentation_data 추출
2. template 복사 → remove_all_sections() (lxml)
3. update_cover_slide(slide[0], title, subtitle)
4. update_toc_slide(slide[1], section_titles)
5. for each section → for each slide:
   a. Body layout(slide[7]) 복제
   b. set_slide_title_area() → 헤더 설정
   c. render_slide_content() → 라우터 → 레이아웃 렌더러
6. 불필요한 템플릿 슬라이드 삭제 (keeper_ids 기반)
7. Ending slide를 마지막으로 이동
8. 저장 → results/{steering_basename}.pptx
```

## Module Design

### 1. generate.py — 오케스트레이션

**책임**: 스티어링 파일 로드, 템플릿 복사, 슬라이드 생성 순서 관리, 저장

**핵심 상수**:
- `IDX_COVER = 0`, `IDX_TOC = 1`, `IDX_BODY_SRC = 7`
- `TEMPLATE_PATH = "template/2025_PPT_Template_FINAL.pptx"`

**핵심 함수**:
- `remove_all_sections(prs)` — lxml로 `<p:sectionLst>` 제거
- `duplicate_slide(prs, template_slide)` — 본문 슬라이드 복제
- `move_slide_to_end(prs, slide)` — 엔딩 슬라이드 이동

**의존성**: powerpoint_cover, powerpoint_toc, powerpoint_content

### 2. powerpoint_cover.py — 표지 렌더러

**책임**: 표지 슬라이드의 제목/부제목/날짜 업데이트, 스타일 보존

**핵심 함수**:
| Function | 역할 |
|----------|------|
| `update_cover_slide(slide, title, subtitle)` | 메인 진입점. 키워드 매칭으로 도형 분류 |
| `get_original_style(shape)` | RGB/테마 색상 추출 |
| `apply_text_with_style(shape, text, style)` | 스타일 승계 후 텍스트 교체 |
| `center_shape_horizontally(shape)` | 수평 중앙 정렬 |
| `find_shapes_by_keywords(shapes, keywords)` | 재귀 도형 검색 |

**설계 결정**:
- 키워드 기반 도형 식별 (위치 기반 대비 템플릿 변경에 강건)
- `int()` 변환으로 좌표 정수 보장 (python-pptx EMU 호환)
- `\\n` 리터럴 + `\n` 모두 줄바꿈 처리

### 3. powerpoint_toc.py — 목차 렌더러

**책임**: 목차 슬라이드 업데이트, 줄 간격 보존

**핵심 함수**:
| Function | 역할 |
|----------|------|
| `update_toc_slide(slide, toc_items)` | 메인 진입점. 두 가지 모드 자동 감지 |
| `update_paragraph_text_only(paragraph, text)` | 문단 삭제 없이 텍스트만 교체 |
| `iter_shapes(shapes)` | 그룹 내부까지 재귀 탐색 |

**설계 결정**:
- **다중 문단 모드** (3줄+ 텍스트박스): 기존 문단 객체 재활용 → 줄 간격 유지
- **개별 박스 모드** (fallback): top 좌표 기반 행 그룹핑 (0.2" 임계값)
- 빈 항목에 `buNone` 추가하여 불릿 표시 제거

### 4. powerpoint_content.py — 27종 레이아웃 렌더러

**책임**: 본문 슬라이드 콘텐츠 렌더링

#### 상수 (파일 상단)

```python
FONTS = {"HEAD_TITLE": "프리젠테이션 7 Bold", "HEAD_DESC": "프리젠테이션 5 Medium",
         "BODY_TITLE": "Freesentation", "BODY_TEXT": "Freesentation"}

COLORS = {"PRIMARY": (0,67,218), "BLACK": (0,0,0), "GRAY": (80,80,80),
          "BG_BOX": (248,249,250), "BORDER": (220,220,220), ...}

LAYOUT = {"SLIDE_TITLE_Y": 0.6, "BODY_START_Y": 2.0, "BODY_LIMIT_Y": 7.2,
          "MARGIN_X": 0.5, "SLIDE_W": 13.333}
```

#### 유틸리티 함수 체인

```
set_slide_title_area(slide, t, d)     → 헤더 영역 설정 (Y=0.6")
    ↓
draw_body_header_and_get_y(slide, wrapper)  → body_title/body_desc 렌더링 + 시작 Y 반환
    ↓
calculate_dynamic_rect(start_y)       → 남은 공간 (x, y, w, h) 계산
    ↓
[Layout Renderer]                     → 각 레이아웃별 렌더링
```

#### 핵심 유틸리티 함수

| Function | 역할 |
|----------|------|
| `create_content_box(slide, ...)` | 만능 박스 (normal/compact/terminal 3모드) |
| `create_terminal_box(slide, ...)` | macOS 스타일 터미널 (보라색 배경, 3색 버튼) |
| `draw_icon_search(slide, search_q, ...)` | 로컬 아이콘 로드 또는 파란 원형 fallback |
| `clean_body_placeholders(slide)` | 본문 영역(2.0"~7.2") 기존 도형 제거 |
| `_place_image_centered(slide, path, ...)` | 종횡비 유지 중앙 배치 |

#### 다이어그램 헬퍼 (detail_sections용)

| Function | 역할 |
|----------|------|
| `_diagram_box(slide, ...)` | 의미색상 기반 박스 (6종: gray/red/orange/green/blue/primary) |
| `_diagram_arrow_label(slide, ...)` | 화살표 텍스트 라벨 |
| `_diagram_shape_arrow(slide, ...)` | 실제 화살표 shape |
| `_draw_diagram_flow/layers/compare/process()` | 4종 다이어그램 렌더러 |
| `_draw_right_diagram(slide, diagram_data, ...)` | type 기반 다이어그램 라우터 |

#### 라우터

`render_slide_content(slide, slide_info)` — `slide_info["l"]`을 키로 딕셔너리에서 렌더러 함수 매핑.
존재하지 않는 레이아웃 요청 시 경고 메시지 출력.

## Data Flow

### 3중 중첩 패턴 (표준, 22종 레이아웃)

```
slide_info["data"]                    # Level 1: slide 메타데이터
    └── wrapper = data["data"]        # Level 2: body_title, body_desc
        └── content = wrapper["data"] # Level 3: 실제 콘텐츠
```

### 예외 패턴 (wrapper 레벨 직접 접근, 2종)

- `challenge_solution`: wrapper에서 challenge, solution 직접 접근
- `before_after`: wrapper에서 before_title/body, after_title/body 직접 접근

## Design Constants

### Phased Columns Gradient (7-step navy→blue)
```python
[(0,27,94), (0,45,143), (0,67,218), (59,122,237), (123,167,247), (160,195,250), (190,215,252)]
```
N개 컬럼에 대해 균등 샘플링하여 색상 배정.

### Semantic Box Styles (_SEM_BOX_STYLES)

| Key | Fill | Line | Text |
|-----|------|------|------|
| `gray` | (248,249,250) | (150,150,150) | (33,33,33) |
| `red` | (254,242,242) | (185,28,28) | (127,29,29) |
| `orange` | (255,247,237) | (194,65,12) | (154,52,18) |
| `green` | (236,253,245) | (4,120,87) | (6,95,70) |
| `blue` | (239,246,255) | (30,58,138) | (30,64,175) |
| `primary` | (239,246,255) | (0,67,218) | (30,64,175) |

## Steering MD Files (.kiro/steering/)

코드 재생성을 위한 완전한 명세. 5개 파일로 분할:

| File | 내용 | 대상 Python 파일 |
|------|------|-----------------|
| `powerpoint-guide.md` | 아키텍처, 상수, 레이아웃 참조 | (참조용) |
| `powerpoint-code-generate.md` | 오케스트레이션 코드 | generate.py, generate_ppt.sh |
| `powerpoint-code-cover-toc.md` | 표지/목차 코드 | powerpoint_cover.py, powerpoint_toc.py |
| `powerpoint-code-content.md` | 본문 Part 1 (유틸리티+레이아웃 1~13) | powerpoint_content.py |
| `powerpoint-code-content-2.md` | 본문 Part 2 (레이아웃 14~27+라우터) | powerpoint_content.py |
