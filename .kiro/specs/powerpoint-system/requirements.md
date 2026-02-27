# PowerPoint 자동 생성 시스템 - Requirements

## Introduction

PowerPoint 자동 생성 시스템은 Python 기반의 스티어링 파일(steering file)을 입력으로 받아 GS Neotek 템플릿 기반의 프레젠테이션 파일(.pptx)을 자동으로 생성하는 시스템입니다.

이 시스템은 다음과 같은 핵심 가치를 제공합니다:
- **재현 가능성**: 스티어링 파일만으로 동일한 PPT를 100% 재생성 가능
- **확장성**: 새로운 레이아웃을 쉽게 추가할 수 있는 모듈형 구조
- **일관성**: 템플릿 기반 디자인 시스템으로 브랜드 일관성 유지
- **효율성**: 수십 장의 슬라이드를 수 초 내에 자동 생성

## Functional Requirements

### 1. 스티어링 파일 기반 PPT 생성

1.1 시스템은 `presentation_data` 딕셔너리를 포함하는 Python 파일(스티어링 파일)을 입력으로 받아야 한다

1.2 스티어링 파일은 다음 구조를 따라야 한다:
```python
presentation_data = {
    "cover": {"title": "...", "subtitle": "..."},
    "sections": [
        {
            "section_title": "1. 섹션명",
            "slides": [
                {"l": "layout_name", "t": "제목", "d": "설명", "data": {...}}
            ]
        }
    ]
}
```

1.3 시스템은 스티어링 파일명을 기반으로 출력 파일명을 자동 생성해야 한다 (예: `rayhli-eks_guide_2026.py` → `results/eks_guide_2026.pptx`)

1.4 시스템은 스티어링 파일 파싱 실패 시 명확한 오류 메시지를 출력해야 한다

### 2. 템플릿 기반 렌더링

2.1 시스템은 `template/2025_PPT_Template_FINAL.pptx` 파일을 기본 템플릿으로 사용해야 한다

2.2 템플릿은 다음 슬라이드 구조를 가져야 한다:
- Index 0: Cover slide (표지)
- Index 1: TOC slide (목차)
- Index 7: Body slide (본문 레이아웃 원본)
- Last slide: Ending slide (감사합니다)

2.3 시스템은 슬라이드 크기 13.333" × 7.500"를 유지해야 한다

2.4 시스템은 템플릿의 섹션(Section) 구조를 제거하고 평면 슬라이드 구조로 변환해야 한다

### 3. 표지 슬라이드 생성

3.1 시스템은 표지 슬라이드의 제목과 부제목을 스티어링 파일의 `cover.title`, `cover.subtitle`로 업데이트해야 한다

3.2 시스템은 표지 슬라이드의 날짜를 현재 날짜로 자동 업데이트해야 한다 (연도: YYYY, 월/일: MM/DD)

3.3 시스템은 템플릿의 기존 텍스트 스타일(폰트, 색상, 크기)을 보존해야 한다

3.4 시스템은 제목과 부제목을 슬라이드 중앙에 수직 정렬해야 한다

### 4. 목차 슬라이드 생성

4.1 시스템은 목차 슬라이드를 스티어링 파일의 `sections[].section_title`로 업데이트해야 한다

4.2 시스템은 section_title에서 숫자 prefix(예: "1. ", "2.1 ")를 제거하고 순수 제목만 표시해야 한다

4.3 시스템은 템플릿의 기존 줄 간격과 스타일을 보존해야 한다

4.4 시스템은 다중 문단 모드(숫자통/제목통 분리)와 개별 박스 모드를 자동 감지해야 한다

### 5. 본문 슬라이드 생성

5.1 시스템은 각 슬라이드의 헤더 영역에 제목(`t`)과 설명(`d`)을 표시해야 한다

5.2 시스템은 헤더 영역을 고정 좌표(Y=0.6")에 배치해야 한다

5.3 시스템은 본문 영역(Y=2.0" ~ 7.2")에 레이아웃별 콘텐츠를 렌더링해야 한다

5.4 시스템은 기존 템플릿 도형을 본문 영역에서 제거해야 한다

### 6. 27종 레이아웃 지원

6.1 시스템은 다음 24종 고유 레이아웃을 지원해야 한다:
- bento_grid, 3_cards, grid_2x2, process_arrow, phased_columns
- timeline_steps, challenge_solution, comparison_vs, comparison_table
- detail_image, image_left, architecture_wide, detail_sections
- table_callout, full_image, before_after, icon_grid
- numbered_list, stats_dashboard, quote_highlight, pros_cons
- do_dont, split_text_code, pyramid_hierarchy, cycle_loop

6.2 시스템은 다음 3종 alias 레이아웃을 지원해야 한다:
- quad_matrix → grid_2x2
- key_metric → 3_cards

6.3 시스템은 레이아웃 다양성 규칙을 적용해야 한다: 동일 레이아웃 최대 3장 (단, 동일 주제/다른 데이터는 예외)

6.4 시스템은 존재하지 않는 레이아웃 요청 시 오류 메시지를 표시해야 한다

### 7. 아이콘 및 이미지 처리

7.1 시스템은 `icons/` 폴더에서 로컬 아이콘 파일을 우선 로드해야 한다 (파일명: `{search_term}.png`)

7.2 시스템은 아이콘 파일이 없을 경우 파란색 원형으로 폴백해야 한다

7.3 시스템은 `architecture/` 폴더에서 다이어그램 이미지를 로드해야 한다

7.4 시스템은 `screenshots/` 폴더에서 UI 스크린샷을 로드해야 한다

7.5 시스템은 이미지 로드 실패 시 회색 박스 placeholder를 표시해야 한다

7.6 시스템은 이미지 종횡비를 유지하며 중앙 정렬해야 한다

### 8. 디자인 시스템

8.1 시스템은 다음 폰트를 사용해야 한다:
- HEAD_TITLE: "프리젠테이션 7 Bold" (28pt)
- HEAD_DESC: "프리젠테이션 5 Medium" (12pt)
- BODY_TITLE: "Freesentation" (16pt)
- BODY_TEXT: "Freesentation" (14pt)

8.2 시스템은 다음 색상 팔레트를 사용해야 한다:
- PRIMARY: RGB(0, 67, 218) - 제목, 강조
- BLACK: RGB(0, 0, 0) - 본문 텍스트
- GRAY: RGB(80, 80, 80) - 설명글
- BG_BOX: RGB(248, 249, 250) - 박스 배경
- Semantic colors: RED, ORANGE, GREEN, BLUE (각 3단계: 제목/배경/본문)

8.3 시스템은 compact 모드에서 폰트 크기를 축소해야 한다 (제목 15pt, 본문 13pt)

8.4 시스템은 터미널 박스에 macOS 스타일 UI를 적용해야 한다 (빨강/노랑/초록 버튼)

### 9. 출력 및 저장

9.1 시스템은 생성된 PPT를 `results/` 폴더에 저장해야 한다

9.2 시스템은 생성 과정을 콘솔에 출력해야 한다 (진행 상황, 슬라이드 제목)

9.3 시스템은 생성 완료 시 총 슬라이드 수를 출력해야 한다

9.4 시스템은 오류 발생 시 스택 트레이스를 출력해야 한다

### 10. 엔딩 슬라이드 보존

10.1 시스템은 템플릿의 마지막 슬라이드(감사합니다)를 보존해야 한다

10.2 시스템은 엔딩 슬라이드를 생성된 PPT의 마지막 위치로 이동해야 한다

10.3 시스템은 엔딩 슬라이드의 내용을 수정하지 않아야 한다

## Non-Functional Requirements

### 11. 재생성 가능성

11.1 시스템은 `.kiro/steering/` 폴더의 markdown 문서만으로 모든 Python 코드를 100% 동일하게 재생성할 수 있어야 한다

11.2 시스템은 코드와 문서 간 양방향 동기화를 지원해야 한다

### 12. 확장성

12.1 새 레이아웃 추가 시 `powerpoint_content.py`에 렌더러 함수 추가 + 라우터 딕셔너리 등록만 필요해야 한다

12.2 시스템은 레이아웃 간 의존성이 없어야 한다 (독립적 렌더링)

### 13. 성능

13.1 시스템은 50장 슬라이드를 10초 이내에 생성해야 한다

13.2 시스템은 메모리 사용량을 500MB 이하로 유지해야 한다

### 14. 호환성

14.1 시스템은 Python 3.8 이상에서 동작해야 한다

14.2 시스템은 다음 라이브러리에 의존해야 한다:
- python-pptx: PowerPoint 생성/수정
- lxml: XML 파싱 (섹션 제거)
- pillow: 이미지 종횡비 계산

### 15. 유지보수성

15.1 시스템은 모듈별로 명확히 분리되어야 한다:
- generate.py: 오케스트레이션
- powerpoint_cover.py: 표지 렌더링
- powerpoint_toc.py: 목차 렌더링
- powerpoint_content.py: 본문 레이아웃 렌더링

15.2 시스템은 상수(FONTS, COLORS, LAYOUT)를 파일 상단에 집중해야 한다

15.3 시스템은 각 함수에 명확한 docstring을 포함해야 한다
