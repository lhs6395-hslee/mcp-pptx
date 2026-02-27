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
│       ├── git-push-on-complete.kiro.hook    # 작업 완료 시 git push
│       └── cross-ide-sync.kiro.hook          # Cross-IDE hook 동기화 감지
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

## 프레젠테이션 생성 Best Practices

### 1. 스티어링 파일 설계

**섹션 구성부터 시작하세요.**
```python
# 먼저 목차(섹션)를 확정하고, 각 섹션에 슬라이드를 배분합니다.
"sections": [
    {"section_title": "1. 개요",      "slides": [...]},   # 2~3장
    {"section_title": "2. 아키텍처",   "slides": [...]},   # 3~4장
    {"section_title": "3. 구현 상세",  "slides": [...]},   # 4~5장
    {"section_title": "4. 운영 가이드", "slides": [...]},  # 3~4장
]
```

**슬라이드 수 가이드:**
- 10분 발표: 10~15장 (3~4 섹션)
- 30분 발표: 20~30장 (5~8 섹션)
- 가이드 문서: 제한 없음 (논리적 섹션 단위로 구성)

### 2. 레이아웃 선택 전략

**콘텐츠 유형에 따라 레이아웃을 선택하세요:**

| 콘텐츠 유형 | 추천 레이아웃 | 피해야 할 레이아웃 |
|-------------|-------------|-------------------|
| 개요/소개 | `bento_grid`, `3_cards`, `icon_grid` | `comparison_table`, `split_text_code` |
| 프로세스/절차 | `process_arrow`, `timeline_steps`, `phased_columns` | `grid_2x2`, `quote_highlight` |
| 비교/분석 | `comparison_vs`, `pros_cons`, `before_after` | `numbered_list`, `cycle_loop` |
| 기술 상세 | `split_text_code`, `detail_sections`, `table_callout` | `stats_dashboard`, `quote_highlight` |
| 성과/KPI | `stats_dashboard`, `3_cards` (key_metric) | `split_text_code`, `pyramid_hierarchy` |
| 아키텍처 | `architecture_wide`, `full_image`, `detail_image` | `numbered_list`, `do_dont` |

**레이아웃 다양성 규칙:**
- 동일 레이아웃은 **최대 3장**까지 허용
- 예외: 같은 주제/로직/다른 데이터 (예: 1주차/2주차/3주차)는 같은 레이아웃 허용
- 연속된 슬라이드에 같은 레이아웃 배치를 피하세요

### 3. 데이터 작성 패턴

**3중 중첩 구조를 지키세요:**
```python
{
    "l": "3_cards",
    "t": "슬라이드 제목",           # 헤더 좌측 (필수)
    "d": "간단한 설명",             # 헤더 우측 (선택)
    "data": {                      # Level 1
        "body_title": "본문 제목",  # 본문 헤더 (선택)
        "body_desc": "본문 설명",   # 본문 서브 (선택)
        "data": {                  # Level 2 → Level 3: 실제 콘텐츠
            "card_1": {"icon": "cloud", "title": "제목", "body": "설명"},
            "card_2": {"icon": "server", "title": "제목", "body": "설명"},
            "card_3": {"icon": "database", "title": "제목", "body": "설명"}
        }
    }
}
```

**예외 레이아웃** (`challenge_solution`, `before_after`):
```python
# Level 2(wrapper)에서 직접 데이터를 읽음
"data": {
    "challenge": "문제점 설명",    # wrapper 레벨
    "solution": "해결책 설명",     # wrapper 레벨
    "data": {}                    # 빈 객체도 가능
}
```

### 4. 아이콘과 이미지

**아이콘 사용 (`search_q` / `icon` 키):**
- `icons/` 폴더의 40종 아이콘 사용 (아래 전체 목록 참조)
- 매칭 실패 시 파란 원형 fallback 자동 적용
- 아이콘명은 소문자 + 언더스코어: `load_balancer`, `aws_account`

**아이콘 전체 목록 (40종):**
```
analysis, aurora, auto_mode, aws_account, billing, chat, cicd, cli,
cluster_delete, config, console, container, cutover, dashboard, database,
deploy, dms, eks, eksctl, encryption, gitops, helm, k8s_version, kubectl,
kubernetes, load_balancer, microservices, migration, monitoring, network,
performance, pipeline, schema, security, server, service, storage,
terraform, timeline, verification
```

**이미지 사용 (`image_path` 키):**
- `architecture/`, `screenshots/` 폴더에 PNG 파일 배치
- 경로는 프로젝트 루트 기준: `"image_path": "architecture/eks_arch.png"`
- 종횡비 자동 유지, 영역 내 중앙 배치

### 5. 텍스트 작성 가이드

- **제목 (`t`)**: 핵심 키워드 중심, 15자 이내
- **설명 (`d`)**: 부연 설명, 25자 이내
- **본문 텍스트**: 줄바꿈은 `\n` 사용
- **불릿 리스트**: 항목당 1~2줄, 3~5개 항목이 적정
- **터미널 코드**: `code_title`과 `code` 키로 macOS 스타일 터미널 박스 표시

### 6. 흔한 실수와 해결

| 실수 | 원인 | 해결 |
|------|------|------|
| 빈 슬라이드 | `data.data.data` 중첩 누락 | 3단계 중첩 확인 |
| 아이콘 미표시 | 아이콘명 오타 또는 대문자 | `icons/` 폴더 파일명 확인 |
| 레이아웃 무시됨 | `l` 키 오타 | 27종 레이아웃명 정확히 확인 |
| 텍스트 잘림 | 본문 영역(Y 2.0"~7.2") 초과 | 텍스트 축소 또는 레이아웃 분할 |
| 이미지 안 나옴 | `image_path` 경로 오류 | 프로젝트 루트 기준 상대경로 확인 |

### 7. AI 에이전트 활용 팁

```
1. PDF/문서 분석 → 섹션 구조 설계
2. 각 섹션별 적합한 레이아웃 선택
3. rayhli-{주제}.py 스티어링 파일 작성
4. ./generate_ppt.sh rayhli-{주제}.py 실행
5. 결과 확인 후 데이터 수정 → 재생성
```

- 스티어링 파일 저장 시 **H1 Hook이 자동 검증** (sections/slides 카운트)
- 한번에 완벽하게 만들려 하지 말고, **반복 수정**으로 완성도를 높이세요
- 레이아웃별 Data Keys는 위 테이블 또는 `AGENTS.md` 참조

## Agent Hooks (6종)

자세한 Hook 설정은 `AGENTS.md`의 **Hooks (자동화)** 섹션을 참조하세요.

| ID | Hook | 트리거 | 동작 |
|----|------|--------|------|
| H1 | 스티어링 파일 검증 | `rayhli-*.py` 저장 시 | sections/slides 카운트 검증 |
| H2 | Steering MD 동기화 | 핵심 Python 파일 수정 시 | 대응 steering md 동기화 안내 |
| H3 | MCP 서버 검증 | 세션 시작 시 | 필수 MCP 서버 누락 확인 |
| H4 | Git Commit & Push | 사용자 요청 시 | commit(한국어) → push |
| H5 | Spec 태스크 업데이트 | 에이전트 작업 완료 시 | tasks.md 체크리스트 갱신 |
| H6 | Cross-IDE 동기화 감지 | Hook 설정 파일 변경 시 | 다른 IDE hook 동기화 안내 |

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
