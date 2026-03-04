# PowerPoint 자동 생성 시스템 (PPT-MCP)

## 프로젝트 개요

스티어링 MD 명세만으로 Python 코드를 재생성하고, 프레젠테이션을 자동 생성하는 엔진.

```
.kiro/steering/ (MD 명세)  →  Python 코드 재생성  →  rayhli-*.py (데이터)  →  rayhli-*.pptx
```

**Version**: v5.1 | **Template**: 13.33" × 7.50" | **레이아웃**: 38종

## 핵심 규칙

### 코드 생성/수정 시
- `.kiro/steering/` md 파일이 **Single Source of Truth**
- Python 코드 수정 후 반드시 해당 steering md도 동기화
- 코드 재생성 시 아래 5개 파일을 순서대로 참조:

| File | 내용 | 생성 대상 |
|------|------|----------|
| `powerpoint-guide.md` | 아키텍처, 상수, 디자인 시스템, 레이아웃 참조 | (참조용) |
| `powerpoint-code-generate.md` | 오케스트레이션 코드 | generate.py, generate_ppt.sh |
| `powerpoint-code-cover-toc.md` | 표지/목차 코드 | powerpoint_cover.py, powerpoint_toc.py |
| `powerpoint-code-content.md` | 본문 Part 1 (유틸리티 + 레이아웃 1~13) | powerpoint_content.py |
| `powerpoint-code-content-2.md` | 본문 Part 2 (레이아웃 14~38 + 라우터) | powerpoint_content.py |

### 스티어링 데이터 파일 작성 시
- 파일명 규칙: `rayhli-{주제}.py`
- `presentation_data` 딕셔너리 하나만 정의하는 순수 데이터 파일
- 데이터 구조: `data.data.data` 3중 중첩 (slide → wrapper → content)
- 예외: `challenge_solution`, `before_after`는 wrapper 레벨에서 직접 읽음
- 레이아웃 다양성: 동일 레이아웃 최대 3장 (같은 주제/다른 데이터는 예외)

### 디자인 시스템
- **폰트**: 프리젠테이션 7 Bold (제목 28pt), Freesentation (본문 14pt)
- **Primary 색상**: RGB(0, 67, 218)
- **본문 영역**: Y 2.0"~7.2", X 여백 0.5"
- **의미 색상**: red=주의, orange=경고, green=긍정, blue=참조
- **터미널 박스**: macOS 스타일 (Ubuntu 보라색 배경, 3색 버튼)

## 프로젝트 구조

```
mcp-pptx/
├── AGENTS.md                           # 공통 프로젝트 가이드 (이 파일)
├── CLAUDE.md                           # Claude Code → AGENTS.md 참조
├── GEMINI.md                           # Antigravity → AGENTS.md 참조
├── .github/copilot-instructions.md     # VS Code Copilot → AGENTS.md 참조
│
├── .kiro/
│   ├── steering/                       # 코드 재생성 명세 (Single Source of Truth)
│   ├── specs/powerpoint-system/        # 정형 명세 (requirements/design/tasks)
│   └── hooks/                          # Kiro Agent Hooks
│
├── .claude/settings.json               # Claude Code Hooks
├── .gemini/settings.json               # Antigravity 설정
│
├── template/                           # PPT 템플릿 (13.33" × 7.50")
├── icons/                              # PNG 아이콘 (512×512)
├── architecture/                       # 다이어그램 PNG
├── screenshots/                        # UI 스크린샷
└── .gitignore                          # Python 코드 + rayhli-*.py + results/ 제외
```

### Git 추적 vs 제외
- **추적**: steering md, specs, hooks, IDE 설정, 에셋(template/icons/architecture/screenshots)
- **제외**: Python 코드(steering md에서 재생성), rayhli-*.py(개별 데이터), results/(출력물)

## 실행 방법

```bash
# 환경 설정
python3 -m venv venv && source venv/bin/activate
pip install python-pptx lxml pillow

# Python 코드 재생성 (AI에게 .kiro/steering/ 읽고 코드 생성 요청)
# PPT 생성
./generate_ppt.sh rayhli-my_presentation.py
```

## 38종 레이아웃

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
| `venn_diagram` | 좌측 3원 벤 + 우측 설명 카드 | circles[]{label,desc,color}, center_label |
| `swot_matrix` | SWOT 분석 매트릭스 | quadrants[]{label,title,items[],color} |
| `center_radial` | 중심 ROUNDED_RECT + 4방향 카드 | center{label,desc}, directions[]{label,desc,color} |
| `funnel` | 퍼널 다이어그램 | stages[]{label,value,desc,color} |
| `zigzag_timeline` | 지그재그 타임라인 | steps[]{date,title,desc,color} |
| `fishbone_cause_effect` | 피쉬본 원인-결과 | effect, categories[]{label,causes[],color} |
| `org_chart` | 조직도/트리 | root{label,desc}, children[]{label,desc,items[],color} |
| `temple_pillars` | 기둥형 구조도 | roof{label}, pillars[]{label,desc,color}, foundation{label} |
| `infinity_loop` | 무한 순환 루프 (상단 라벨+흐름 화살표) | left_loop[],right_loop[],center_label,left_label,right_label |
| `speedometer_gauge` | 스피도미터 게이지 | value, segments[]{label,color}, title |
| `mind_map` | 좌측 방사형 맵 + 우측 설명 카드 | center{label}, branches[]{label,sub_branches[],desc,color} |

## 의존성

```
python-pptx   # PowerPoint 생성
lxml          # XML 파싱 (섹션 제거)
pillow        # 이미지 종횡비 (선택)
```

## MCP 서버 (필수)

이 프로젝트에서 사용하는 MCP 서버 목록입니다. IDE에 설정되어 있지 않으면 추가해주세요.

| 서버 | 명령어 | 용도 |
|------|--------|------|
| `context7` | `npx -y @upstash/context7-mcp@latest` | 라이브러리 문서 검색 |
| `infrastructure-diagrams` | `uvx infrastructure-diagram-mcp-server` | 인프라 다이어그램 생성 |
| `powerpoint` | `uvx --from office-powerpoint-mcp-server ppt_mcp_server` | PPT 직접 조작 |
| `fetch` | `uvx mcp-server-fetch` | 웹 페이지 fetch |

### IDE별 설정 방법

**Claude Code:**
```bash
claude mcp add context7 -- npx -y @upstash/context7-mcp@latest
claude mcp add infrastructure-diagrams -- uvx infrastructure-diagram-mcp-server
claude mcp add powerpoint -- uvx --from office-powerpoint-mcp-server ppt_mcp_server
claude mcp add fetch -- uvx mcp-server-fetch
```

**Kiro:** `.kiro/settings/mcp.json`
```json
{
  "mcpServers": {
    "context7": { "command": "npx", "args": ["-y", "@upstash/context7-mcp@latest"] },
    "infrastructure-diagrams": { "command": "uvx", "args": ["infrastructure-diagram-mcp-server"] },
    "powerpoint": { "command": "uvx", "args": ["--from", "office-powerpoint-mcp-server", "ppt_mcp_server"] },
    "fetch": { "command": "uvx", "args": ["mcp-server-fetch"] }
  }
}
```

**VS Code Copilot:** `.vscode/mcp.json`
```json
{
  "servers": {
    "context7": { "type": "stdio", "command": "npx", "args": ["-y", "@upstash/context7-mcp@latest"] },
    "infrastructure-diagrams": { "type": "stdio", "command": "uvx", "args": ["infrastructure-diagram-mcp-server"] },
    "powerpoint": { "type": "stdio", "command": "uvx", "args": ["--from", "office-powerpoint-mcp-server", "ppt_mcp_server"] },
    "fetch": { "type": "stdio", "command": "uvx", "args": ["mcp-server-fetch"] }
  }
}
```

**Antigravity:** `~/.gemini/antigravity/mcp_config.json`에 수동 추가

## Hooks (자동화)

이 프로젝트에서 사용하는 자동화 Hook 목록입니다. IDE에 설정되어 있지 않으면 추가해주세요.

| ID | Hook | 트리거 | 동작 | 비고 |
|----|------|--------|------|------|
| H1 | 스티어링 파일 검증 | `rayhli-*.py` 저장 시 | Python exec → sections/slides 카운트 | 구문 오류 즉시 감지 |
| H2 | Steering MD 동기화 알림 | 핵심 Python 파일 수정 시 | 대응 steering md 동기화 안내 | 매핑 규칙 아래 참조 |
| H3 | MCP 서버 검증 | 세션 시작 시 | 필수 MCP 서버 누락 확인 | context7, infra-diagrams, ppt, fetch |
| H4 | Git Commit & Push | 사용자 요청 시 | git add → commit(한국어) → push | 수동 트리거만 (자동 X) |
| H5 | Spec 태스크 업데이트 | 에이전트 작업 완료 시 | tasks.md 체크리스트 갱신 | .kiro/specs/ 전용 |
| H6 | Cross-IDE Hook 동기화 감지 | Hook 설정 파일 변경 시 | 다른 IDE hook 동기화 안내 | 양방향 감지 |

### H2 파일→MD 매핑 규칙

| Python 파일 | Steering MD |
|-------------|-------------|
| `generate.py` / `generate_ppt.sh` | `powerpoint-code-generate.md` |
| `powerpoint_cover.py` / `powerpoint_toc.py` | `powerpoint-code-cover-toc.md` |
| `powerpoint_content.py` (유틸리티+레이아웃 1~13) | `powerpoint-code-content.md` |
| `powerpoint_content.py` (레이아웃 14~27+라우터) | `powerpoint-code-content-2.md` |
| 상수/아키텍처 변경 | `powerpoint-guide.md` |

### IDE별 Hook 매핑

| ID | Kiro 트리거 | Kiro 타입 | Claude Code 이벤트 | Claude Code 타입 |
|----|------------|----------|-------------------|-----------------|
| H1 | `fileEdited` (rayhli-*.py) | `runCommand` | `PostToolUse` (Edit\|Write) | `command` |
| H2 | `fileEdited` (powerpoint_*.py 등) | `askAgent` | `PostToolUse` (Edit\|Write) | `command` |
| H3 | `userTriggered` | `askAgent` | `SessionStart` | `command` |
| H4 | `userTriggered` | `askAgent` | _(사용자 요청)_ | _(해당 없음)_ |
| H5 | `agentStop` | `askAgent` | `Stop` | `command` |
| H6 | `fileEdited` (.claude/settings.json) | `askAgent` | `PostToolUse` (Edit\|Write) | `command` |

### IDE별 설정 방법

**Claude Code:** `.claude/settings.json`의 `hooks` 섹션

```json
{
  "hooks": {
    "PostToolUse": [
      { "matcher": "Edit|Write", "hooks": [{ "type": "command", "command": "H2: 핵심 Python 파일 감지" }] },
      { "matcher": "Edit|Write", "hooks": [{ "type": "command", "command": "H1: rayhli-*.py 검증" }] },
      { "matcher": "Edit|Write", "hooks": [{ "type": "command", "command": "H6: .kiro/hooks/ 또는 .claude/settings.json 변경 감지" }] }
    ],
    "SessionStart": [{ "hooks": [{ "type": "command", "command": "H3: MCP 서버 누락 확인" }] }],
    "Stop": [{ "hooks": [{ "type": "command", "command": "H5: tasks.md 체크리스트 갱신 알림" }] }]
  }
}
```

**Kiro:** `.kiro/hooks/*.kiro.hook` 파일

| Hook ID | 파일명 |
|---------|--------|
| H1 | `validate-steering-file.kiro.hook` |
| H2 | `sync-steering-md.kiro.hook` |
| H3 | `sync-mcp-servers.kiro.hook` |
| H4 | `git-push-on-complete.kiro.hook` |
| H5 | `update-specs-on-change.kiro.hook` |
| H6 | `cross-ide-sync.kiro.hook` |

### Cross-IDE 동기화 원칙

- **AGENTS.md**가 Hook의 Single Source of Truth
- Hook 추가/수정 시: AGENTS.md 먼저 업데이트 → 현재 IDE에 구현 → H6이 다른 IDE 동기화 안내
- Claude Code → `.kiro/hooks/` 편집 감지 시 `.claude/settings.json` 동기화 안내
- Kiro → `.claude/settings.json` 편집 감지 시 `.kiro/hooks/` 동기화 안내
