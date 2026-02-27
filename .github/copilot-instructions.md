# PowerPoint 자동 생성 시스템 (PPT-MCP)

## 프로젝트 개요
스티어링 파일(데이터)과 python-pptx 렌더러로 프레젠테이션을 자동 생성하는 엔진.
**Version**: v5.1 | **Template**: 13.33" × 7.50" | **레이아웃**: 27종

## 핵심 규칙

### 코드 생성/수정 시
- `.kiro/steering/` md 파일이 **Single Source of Truth** — Python 코드 수정 후 반드시 해당 md도 동기화
- `powerpoint-guide.md`는 아키텍처/상수/레이아웃 참조, 나머지 4개는 소스코드 전문
- 코드 수정 시 md 파일 → Python 파일 **양방향 동기화** 필수

### 스티어링 파일 작성 시
- `presentation_data` 딕셔너리 하나만 정의하는 순수 데이터 파일
- 데이터 구조: `data.data.data` 3중 중첩 (slide → wrapper → content)
- 예외: `challenge_solution`, `before_after`는 wrapper 레벨에서 직접 읽음
- 레이아웃 다양성: 동일 레이아웃 최대 3장 (같은 주제/다른 데이터는 예외)

### 디자인 시스템
- **폰트**: 프리젠테이션 7 Bold (제목), Freesentation (본문)
- **Primary 색상**: RGB(0, 67, 218)
- **본문 영역**: Y 2.0"~7.2", X 여백 0.5"
- **의미 색상**: red=주의, orange=경고, green=긍정, blue=참조

## 파일 구조
```
generate_ppt.sh          → Shell wrapper (one-line 실행)
generate.py              → 오케스트레이션 (템플릿 복사→섹션 제거→렌더링→저장)
powerpoint_content.py    → 27종 레이아웃 렌더러 + 유틸리티
powerpoint_cover.py      → 표지 렌더러
powerpoint_toc.py        → 목차 렌더러
rayhli-eks_guide_2026.py        → 스티어링 파일 (AWS EKS Guide)
rayhli-ss_db_migration_resume.py → 스티어링 파일 (DB Migration)
```

## 실행 방법
```bash
./generate_ppt.sh                           # 기본 (rayhli-eks_guide_2026.py)
./generate_ppt.sh rayhli-ss_db_migration_resume.py # 다른 스티어링 파일
python3 generate.py my_presentation.py      # 직접 실행
```

## 의존성
```
python-pptx, lxml, pillow
```

## 스티어링 md 파일 (.kiro/steering/)
| File | 내용 |
|------|------|
| `powerpoint-guide.md` | 아키텍처, 상수, 디자인 시스템, 레이아웃 참조 |
| `powerpoint-code-generate.md` | generate.py + generate_ppt.sh |
| `powerpoint-code-cover-toc.md` | powerpoint_cover.py + powerpoint_toc.py |
| `powerpoint-code-content.md` | powerpoint_content.py Part 1 (유틸리티 + 레이아웃 1~13) |
| `powerpoint-code-content-2.md` | powerpoint_content.py Part 2 (레이아웃 14~27 + 라우터) |

## 27종 레이아웃
bento_grid, 3_cards, grid_2x2, quad_matrix(alias), process_arrow, phased_columns,
timeline_steps, challenge_solution, comparison_vs, comparison_table, detail_image,
image_left, architecture_wide, key_metric(alias), detail_sections, table_callout,
full_image, before_after, icon_grid, numbered_list, stats_dashboard, quote_highlight,
pros_cons, do_dont, split_text_code, pyramid_hierarchy, cycle_loop
