# PowerPoint 자동 생성 시스템 - Tasks

## Task 1: 오케스트레이션 엔진 구현 (generate.py + generate_ppt.sh)

- [x] `generate_ppt.sh` 쉘 래퍼 작성 (venv 활성화 + python3 실행)
- [x] 스티어링 파일 `exec()` 로드 및 `presentation_data` 추출
- [x] 템플릿 복사 및 `remove_all_sections()` 구현 (lxml)
- [x] `duplicate_slide()` 구현 (본문 레이아웃 복제)
- [x] keeper_ids 기반 불필요 슬라이드 삭제
- [x] `move_slide_to_end()` 구현 (엔딩 슬라이드 이동)
- [x] 출력 파일명 자동 생성 (`results/{steering_basename}.pptx`)

## Task 2: 표지 렌더러 구현 (powerpoint_cover.py)

- [x] 키워드 기반 도형 분류 (제목/부제목/날짜)
- [x] `get_original_style()` 구현 (RGB + 테마 색상 추출)
- [x] `apply_text_with_style()` 구현 (스타일 승계 + `\\n` 줄바꿈 처리)
- [x] `center_shape_horizontally()` 구현 (수평 중앙 정렬)
- [x] 제목+부제목 수직 중앙 정렬 (`int()` 변환 적용)
- [x] 날짜 자동 업데이트 (연도 YYYY, 월/일 MM/DD)

## Task 3: 목차 렌더러 구현 (powerpoint_toc.py)

- [x] `iter_shapes()` 구현 (그룹 내부 재귀 탐색)
- [x] `update_paragraph_text_only()` 구현 (문단 보존 텍스트 교체)
- [x] 다중 문단 모드 구현 (3줄+ 텍스트박스, 숫자통/제목통 분리)
- [x] 개별 박스 모드 구현 (fallback, top 좌표 기반 행 그룹핑)
- [x] section_title에서 숫자 prefix 제거 (`^\d+\.\s*`)
- [x] 빈 항목 불릿 제거 (`buNone` 처리)

## Task 4: 본문 렌더러 상수 및 유틸리티 구현 (powerpoint_content.py)

- [x] FONTS, COLORS, LAYOUT 상수 딕셔너리 정의
- [x] `set_slide_title_area()` 구현 (헤더 영역 설정)
- [x] `draw_body_header_and_get_y()` 구현 (body_title/body_desc + 시작 Y 반환)
- [x] `calculate_dynamic_rect()` 구현 (남은 공간 계산)
- [x] `clean_body_placeholders()` 구현 (본문 영역 기존 도형 제거)
- [x] `create_content_box()` 구현 (normal/compact/terminal 3모드)
- [x] `create_terminal_box()` 구현 (macOS 스타일, 3색 버튼)
- [x] `draw_icon_search()` 구현 (로컬 아이콘 로드 + 파란 원형 fallback)
- [x] `_place_image_centered()` 구현 (종횡비 유지 중앙 배치)

## Task 5: 카드/그리드 레이아웃 구현 (레이아웃 1~4)

- [x] `render_bento_grid()` — 좌 50% + 우 2분할 (main, sub1, sub2)
- [x] `render_3_cards()` — 아이콘+제목+본문 카드 3개
- [x] `render_grid_2x2()` — 4분할 compact 모드 (item1~4)
- [x] `quad_matrix` → `grid_2x2` alias 등록

## Task 6: 프로세스/타임라인 레이아웃 구현 (레이아웃 5~7)

- [x] `render_process_arrow()` — 쉐브론+본문 박스 (steps[])
- [x] `render_phased_columns()` — 단계별 컬럼+그라데이션 헤더 (7-step palette)
- [x] `render_timeline_steps()` — 숫자 배지+카드 (steps[])

## Task 7: 비교/테이블 레이아웃 구현 (레이아웃 8~10)

- [x] `render_challenge_solution()` — 좌우+화살표 (wrapper 레벨 직접 접근)
- [x] `render_comparison_vs()` — VS 원형 (item_a/b_title/body)
- [x] `render_comparison_table()` — 3열 비교 표 (columns[], rows[])

## Task 8: 이미지/아키텍처 레이아웃 구현 (레이아웃 11~13)

- [x] `render_detail_image()` — 상단 텍스트+하단 이미지 (title, body, search_q)
- [x] `render_image_left()` — 좌 이미지+우 불릿 (image_path, bullets[])
- [x] `render_architecture_wide()` — 상단 다이어그램+하단 3열 (col1~3)

## Task 9: 확장 레이아웃 구현 Part 1 (레이아웃 14~18)

- [x] `key_metric` → `3_cards` alias 등록
- [x] `render_detail_sections()` — 좌 멀티섹션+우 다이어그램 (4종 diagram type)
- [x] 다이어그램 헬퍼 구현 (_diagram_box, _diagram_arrow_label, _diagram_shape_arrow)
- [x] 4종 다이어그램 렌더러 구현 (flow, layers, compare, process)
- [x] `render_table_callout()` — 테이블+콜아웃 박스
- [x] `render_full_image()` — 풀와이드 이미지+캡션
- [x] `render_before_after()` — 전/후 비교 (wrapper 레벨 직접 접근)

## Task 10: 확장 레이아웃 구현 Part 2 (레이아웃 19~23)

- [x] `render_icon_grid()` — 3열×N행 아이콘 그리드
- [x] `render_numbered_list()` — 번호형 세로 리스트
- [x] `render_stats_dashboard()` — KPI 대형 숫자 (metrics[])
- [x] `render_quote_highlight()` — 인용문 강조 (quote, author, role)
- [x] `render_pros_cons()` — 장단점 비교 (subject, pros[], cons[])

## Task 11: 확장 레이아웃 구현 Part 3 (레이아웃 24~27)

- [x] `render_do_dont()` — Best Practice (do_items[], dont_items[])
- [x] `render_split_text_code()` — 설명+코드 병렬 (description, bullets[], code)
- [x] `render_pyramid_hierarchy()` — 피라미드 계층 (levels[])
- [x] `render_cycle_loop()` — 순환 프로세스 (steps[], center_label)

## Task 12: 고급 다이어그램 레이아웃 구현 (레이아웃 28~38)

- [x] `render_venn_diagram()` — 좌측 3원 벤 + 우측 설명 카드 (circles[], center_label)
- [x] `render_swot_matrix()` — 2×2 SWOT 분석 매트릭스 (quadrants[])
- [x] `render_center_radial()` — 중심 ROUNDED_RECT + 4방향 카드 (center, directions[])
- [x] `render_funnel()` — 퍼널 다이어그램 (stages[])
- [x] `render_zigzag_timeline()` — 지그재그 타임라인 (steps[])
- [x] `render_fishbone_cause_effect()` — 피쉬본 원인-결과 (effect, categories[])
- [x] `render_org_chart()` — 조직도/트리 계층 (root, children[])
- [x] `render_temple_pillars()` — 기둥형 구조도 (roof, pillars[], foundation)
- [x] `render_infinity_loop()` — 무한 순환 루프 (left_loop[], right_loop[], labels)
- [x] `render_speedometer_gauge()` — 스피도미터 게이지 (value, segments[], title)
- [x] `render_mind_map()` — 좌측 방사형 맵 + 우측 설명 카드 (center, branches[])

## Task 13: 라우터 및 통합

- [x] `render_slide_content()` 라우터 딕셔너리 구현 (38종 매핑)
- [x] 존재하지 않는 레이아웃 요청 시 경고 메시지 처리
- [x] 전체 스티어링 파일(rayhli-eks_guide_2026.py)로 통합 테스트
- [x] 전체 스티어링 파일(rayhli-ss_db_migration_resume.py)로 통합 테스트

## Task 14: 스티어링 MD 문서화

- [x] `powerpoint-guide.md` 작성 (아키텍처, 상수, 38종 레이아웃 참조)
- [x] `powerpoint-code-generate.md` 작성 (generate.py + generate_ppt.sh 전문)
- [x] `powerpoint-code-cover-toc.md` 작성 (표지/목차 코드 전문)
- [x] `powerpoint-code-content.md` 작성 (본문 Part 1: 유틸리티+레이아웃 1~13)
- [x] `powerpoint-code-content-2.md` 작성 (본문 Part 2: 레이아웃 14~38+라우터)

## Task 15: 멀티 IDE 지원

- [x] `AGENTS.md` 작성 (공통 프로젝트 가이드, 38종 레이아웃 테이블)
- [x] `CLAUDE.md` 작성 (Claude Code 프로젝트 가이드)
- [x] `GEMINI.md` 작성 (Antigravity 프로젝트 가이드)
- [x] `.gemini/settings.json` 작성 (AGENTS.md 병행 읽기 설정)
- [x] `.github/copilot-instructions.md` 작성 (VS Code Copilot 가이드)
- [x] `.kiro/specs/powerpoint-system/` Kiro spec 작성 (requirements/design/tasks)

## Task 15: Cross-IDE Hook 동기화

- [x] AGENTS.md에 Hooks (자동화) 섹션 추가 (H1~H6 정규표 + IDE별 매핑)
- [x] `.claude/settings.json`에 H5(Stop/prompt), H6(Cross-IDE 감지) 추가
- [x] 기존 Claude Code hook 메시지에 `[Hook:Hx]` 라벨 추가
- [x] `.kiro/hooks/cross-ide-sync.kiro.hook` 생성 (양방향 감지)
- [x] `.kiro/hooks/git-push-on-complete.kiro.hook` agentStop→userTriggered 확정
