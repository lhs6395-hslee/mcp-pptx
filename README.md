# PowerPoint 자동 생성 시스템

**스티어링 MD 명세만으로 Python 코드를 재생성하고, 프레젠테이션을 자동 생성하는 엔진**

**Version**: v5.1 (2026-02-27) | **Repository**: [mcp-pptx](https://github.com/lhs6395-hslee/mcp-pptx)

## 핵심 컨셉

```
.kiro/steering/ (MD 명세)  →  Python 코드 재생성  →  rayhli-*.py (데이터)  →  rayhli-*.pptx
```

- **Git에는 시스템(엔진)만 관리** — Python 코드와 개별 프레젠테이션 데이터는 제외
- **Steering MD가 Single Source of Truth** — 5개 md 파일에서 모든 Python 코드를 100% 재생성
- **멀티 IDE 지원** — Kiro, Claude Code, Antigravity, VS Code Copilot

## 빠른 시작

```bash
# 1. 클론
git clone https://github.com/lhs6395-hslee/mcp-pptx.git
cd mcp-pptx

# 2. 환경 설정
python3 -m venv venv && source venv/bin/activate
pip install python-pptx lxml pillow

# 3. Python 코드 재생성 (AI에게 steering md 파일을 읽고 코드 생성 요청)
#    → generate.py, powerpoint_content.py, powerpoint_cover.py, powerpoint_toc.py 생성

# 4. 스티어링 데이터 파일 작성 (rayhli-*.py)
#    → presentation_data 딕셔너리 정의

# 5. PPT 생성
./generate_ppt.sh rayhli-my_presentation.py
open results/rayhli-my_presentation.pptx
```

## 프로젝트 구조

```
mcp-pptx/
│
├── .kiro/
│   ├── steering/                        # 코드 재생성 명세 (Single Source of Truth)
│   │   ├── powerpoint-guide.md          #   아키텍처, 상수, 디자인 시스템, 레이아웃 참조
│   │   ├── powerpoint-code-generate.md  #   generate.py + generate_ppt.sh
│   │   ├── powerpoint-code-cover-toc.md #   powerpoint_cover.py + powerpoint_toc.py
│   │   ├── powerpoint-code-content.md   #   powerpoint_content.py Part 1 (유틸리티 + 1~13)
│   │   └── powerpoint-code-content-2.md #   powerpoint_content.py Part 2 (14~27 + 라우터)
│   │
│   ├── specs/powerpoint-system/         # 정형 명세 (Kiro IDE)
│   │   ├── requirements.md              #   FR 10개 + NFR 5개
│   │   ├── design.md                    #   아키텍처, 모듈 설계, 데이터 플로우
│   │   └── tasks.md                     #   태스크 추적
│   │
│   └── hooks/                           # Kiro Agent Hooks
│       ├── sync-steering-md.kiro.hook   #   Python 수정 → steering md 동기화 알림
│       ├── validate-steering-file.kiro.hook  # 스티어링 파일 구문 검증
│       ├── update-specs-on-change.kiro.hook  # tasks.md 자동 업데이트
│       └── git-push-on-complete.kiro.hook    # 작업 완료 시 git push
│
├── .claude/settings.json               # Claude Code Hooks
├── .gemini/settings.json               # Antigravity 설정
├── .github/copilot-instructions.md     # VS Code Copilot 가이드
├── CLAUDE.md                           # Claude Code 프로젝트 가이드
├── GEMINI.md                           # Antigravity 프로젝트 가이드
│
├── template/                           # PPT 템플릿 (13.33" × 7.50")
├── icons/                              # PNG 아이콘 (512×512)
├── architecture/                       # 다이어그램 PNG
├── screenshots/                        # UI 스크린샷
│
├── .gitignore                          # Python 코드 + rayhli-*.py + results/ 제외
└── README.md
```

### Git 추적 vs 제외

| 추적 (시스템/엔진) | 제외 (생성물/데이터) |
|-------------------|-------------------|
| `.kiro/steering/` (코드 명세) | `*.py` (Python 코드 — steering md에서 재생성) |
| `.kiro/specs/` (요구사항/설계) | `rayhli-*.py` (개별 프레젠테이션 데이터) |
| `.kiro/hooks/` (자동화) | `results/*.pptx` (PPT 출력물) |
| IDE 설정 (CLAUDE.md 등) | `venv/`, `__pycache__/` |
| 에셋 (template, icons 등) | |

## 스티어링 파일 작성법

파일명 규칙: `rayhli-{주제}.py`

```python
# rayhli-my_topic.py
presentation_data = {
    "cover": {"title": "제목", "subtitle": "부제목"},
    "sections": [
        {
            "section_title": "1. 섹션",
            "slides": [
                {
                    "l": "3_cards",          # 레이아웃 (27종 중 택 1)
                    "t": "슬라이드 제목",     # 헤더 좌측
                    "d": "설명",             # 헤더 우측
                    "data": { ... }          # 레이아웃별 데이터
                }
            ]
        }
    ]
}
```

## 27종 레이아웃

| Layout | 용도 | Data Keys |
|--------|------|-----------|
| `bento_grid` | 메인+서브 2분할 | main, sub1, sub2 |
| `3_cards` | 아이콘 카드 3개 | card_1, card_2, card_3 |
| `grid_2x2` | 4분할 compact | item1~item4 |
| `quad_matrix` | grid_2x2 alias | (동일) |
| `process_arrow` | 쉐브론 프로세스 | steps[] |
| `phased_columns` | 단계별 그라데이션 컬럼 | steps[] |
| `timeline_steps` | 숫자 배지 타임라인 | steps[] |
| `challenge_solution` | 문제→해결 | challenge, solution |
| `comparison_vs` | A vs B | item_a/b_title/body |
| `comparison_table` | 3열 비교 표 | columns[], rows[] |
| `detail_image` | 텍스트+이미지 | title, body, search_q |
| `image_left` | 좌 이미지+우 불릿 | image_path, bullets[] |
| `architecture_wide` | 다이어그램+3열 | col1, col2, col3 |
| `key_metric` | 3_cards alias | (동일) |
| `detail_sections` | 멀티섹션+다이어그램 | overview, highlight, condition, diagram |
| `table_callout` | 테이블+콜아웃 | columns[], rows[], callout |
| `full_image` | 풀와이드 이미지 | image_path/search_q, caption |
| `before_after` | 전/후 비교 | before/after_title/body |
| `icon_grid` | 3열 아이콘 그리드 | items[]{icon,title,desc} |
| `numbered_list` | 번호형 리스트 | items[]{title,desc} |
| `stats_dashboard` | KPI 대형 숫자 | metrics[]{value,unit,label,desc} |
| `quote_highlight` | 인용문 강조 | quote, author, role |
| `pros_cons` | 장단점 비교 | subject, pros[], cons[] |
| `do_dont` | Best Practice | do_items[], dont_items[] |
| `split_text_code` | 설명+코드 | description, bullets[], code |
| `pyramid_hierarchy` | 피라미드 계층 | levels[]{label,desc,color} |
| `cycle_loop` | 순환 프로세스 | steps[]{label,desc}, center_label |

## Agent Hooks

### Kiro IDE (4개)

| Hook | 트리거 | 동작 |
|------|--------|------|
| Steering MD 동기화 | Python 파일 저장 | 대응 md 파일 동기화 요청 |
| 스티어링 파일 검증 | `rayhli-*.py` 저장 | 구문/구조 자동 검증 |
| Spec 태스크 업데이트 | 에이전트 완료 | tasks.md 체크리스트 업데이트 |
| Git Push | 에이전트 완료 | 변경사항 commit + push |

### Claude Code (3개)

| Hook | 트리거 | 동작 |
|------|--------|------|
| Steering MD 동기화 | Python 파일 Edit/Write | 경고 메시지 |
| 스티어링 파일 검증 | `rayhli-*.py` Edit/Write | python3 구문 검증 |
| Git Push | 대화 종료 (Stop) | commit + push 요청 |

## 디자인 시스템

| 항목 | 값 |
|------|-----|
| 폰트 (제목) | 프리젠테이션 7 Bold (28pt) |
| 폰트 (본문) | Freesentation (14pt) |
| Primary 색상 | RGB(0, 67, 218) |
| 템플릿 | GS Neotek 2025 (13.33" × 7.50") |
| 본문 영역 | Y 2.0"~7.2", X 여백 0.5" |
| 터미널 박스 | macOS 스타일 (Ubuntu 보라색, 3색 버튼) |
| 의미 색상 | red=주의, orange=경고, green=긍정, blue=참조 |

## 의존성

```
python-pptx   # PowerPoint 생성
lxml          # XML 파싱 (섹션 제거)
pillow        # 이미지 종횡비 (선택)
```
