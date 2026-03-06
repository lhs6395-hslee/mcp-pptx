# PowerPoint Generation - Content Renderers Source Code (Part 3)

**Part of**: [powerpoint-guide.md](./powerpoint-guide.md) 시스템 명세
**Continues from**: [powerpoint-code-content-2.md](./powerpoint-code-content-2.md)
**File**: `powerpoint_content.py` — Layouts 28~41 + Router

---

## powerpoint_content.py (continued) — Layouts 28~41 + Router

아래 코드는 `powerpoint-code-content-2.md`의 코드에 이어서 같은 파일(`powerpoint_content.py`)에 포함됩니다.

```python
# 28. SWOT Matrix (SWOT 분석 매트릭스)
def render_swot_matrix(slide, data):
    """2×2 그리드 + 중앙 축 라벨 (S/W/O/T 또는 커스텀)

    data.data.data:
        quadrants: [
            {"label": "S", "title": "Strengths", "items": ["기술력", "경험"], "color": "blue"},
            {"label": "W", "title": "Weaknesses", "items": ["인력부족"], "color": "red"},
            {"label": "O", "title": "Opportunities", "items": ["시장확대"], "color": "green"},
            {"label": "T", "title": "Threats", "items": ["경쟁심화"], "color": "orange"}
        ]
    """
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    quadrants = content.get('quadrants', [])
    if len(quadrants) < 4: return

    gap = Inches(0.5)   # 중앙 라벨 영역
    cell_w = (bw - gap) / 2; cell_h = (bh - gap) / 2

    positions = [
        (bx, by),                           # 좌상 (S)
        (bx + cell_w + gap, by),             # 우상 (W)
        (bx, by + cell_h + gap),             # 좌하 (O)
        (bx + cell_w + gap, by + cell_h + gap),  # 우하 (T)
    ]

    default_colors = ['blue', 'red', 'green', 'orange']

    for i, q in enumerate(quadrants[:4]):
        qx, qy = positions[i]
        color_key = q.get('color', default_colors[i])
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(color_key, _SEM_BOX_STYLES['primary'])

        # 사분면 박스
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, int(qx), int(qy), int(cell_w), int(cell_h))
        shp.fill.solid(); shp.fill.fore_color.rgb = fill_c
        shp.line.color.rgb = line_c; shp.line.width = Pt(2.0)

        tf = shp.text_frame; tf.clear(); tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.2); tf.margin_right = Inches(0.2); tf.margin_top = Inches(0.15)

        # 제목 (예: "Strengths")
        title = q.get('title', '')
        p = tf.paragraphs[0]; p.text = title; p.font.name = FONTS["BODY_TITLE"]
        p.font.bold = True; p.font.size = Pt(15); p.font.color.rgb = line_c; p.space_after = Pt(6)

        # 항목 리스트
        items = q.get('items', [])
        for item in items:
            p2 = tf.add_paragraph(); p2.text = f"• {item}"; p2.font.name = FONTS["BODY_TEXT"]
            p2.font.size = Pt(12); p2.font.color.rgb = text_c; p2.space_after = Pt(3); p2.line_spacing = 1.2

    # 중앙 라벨 (S/W/O/T)
    label_size = Inches(0.45)
    label_positions = [
        (bx + cell_w - label_size / 2 + gap / 2, by + cell_h - label_size / 2 + gap / 2),  # 중앙
    ]
    # 4개 라벨을 중앙 교차점에 배치
    cx = int(bx + cell_w + gap / 2); cy_mid = int(by + cell_h + gap / 2)
    labels_pos = [
        (cx - int(label_size) - Inches(0.02), cy_mid - int(label_size) - Inches(0.02)),  # S (좌상)
        (cx + Inches(0.02), cy_mid - int(label_size) - Inches(0.02)),                    # W (우상)
        (cx - int(label_size) - Inches(0.02), cy_mid + Inches(0.02)),                    # O (좌하)
        (cx + Inches(0.02), cy_mid + Inches(0.02)),                                       # T (우하)
    ]

    for i, q in enumerate(quadrants[:4]):
        lx, ly = labels_pos[i]
        color_key = q.get('color', default_colors[i])
        _, line_c, _ = _SEM_BOX_STYLES.get(color_key, _SEM_BOX_STYLES['primary'])

        lbl = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, int(lx), int(ly), int(label_size), int(label_size))
        lbl.fill.solid(); lbl.fill.fore_color.rgb = line_c; lbl.line.color.rgb = line_c
        tf = lbl.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.text = q.get('label', ''); p.font.name = FONTS["BODY_TITLE"]
        p.font.bold = True; p.font.size = Pt(18); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment = PP_ALIGN.CENTER


# 29. Center Radial (중심 방사형 관계도)
def render_center_radial(slide, data):
    """중앙 원 + 4방향 화살표 + 코너 라벨/설명

    data.data.data:
        center: {"label": "핵심 전략", "desc": "디지털 트랜스포메이션"}
        directions: [
            {"label": "기술", "desc": "클라우드, AI, DevOps", "color": "blue"},
            {"label": "프로세스", "desc": "자동화, 표준화", "color": "green"},
            {"label": "인력", "desc": "역량강화, 교육", "color": "orange"},
            {"label": "문화", "desc": "혁신, 협업", "color": "red"}
        ]
    """
    import math
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    center = content.get('center', {}); directions = content.get('directions', [])
    if not directions: return

    cx = int(bx) + int(bw) // 2; cy = int(by) + int(bh) // 2

    # 중앙 노드 (ROUNDED_RECTANGLE — OVAL은 내접 텍스트영역이 좁아 단어 잘림 발생)
    center_w = int(int(bh) * 0.42); center_h = int(int(bh) * 0.30)
    center_shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx - center_w // 2, cy - center_h // 2, center_w, center_h)
    center_shp.fill.solid(); center_shp.fill.fore_color.rgb = COLORS["PRIMARY"]
    center_shp.line.color.rgb = COLORS["PRIMARY"]; center_shp.line.width = Pt(3.0)
    tf = center_shp.text_frame; tf.clear(); tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.15); tf.margin_right = Inches(0.15)
    tf.margin_top = Inches(0.06); tf.margin_bottom = Inches(0.06)
    c_label = center.get('label', '')
    p = tf.paragraphs[0]; p.text = c_label; p.font.name = FONTS["BODY_TITLE"]
    c_fs = Pt(11) if len(c_label) > 16 else Pt(13) if len(c_label) > 10 else Pt(15)
    p.font.bold = True; p.font.size = c_fs; p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment = PP_ALIGN.CENTER
    c_desc = center.get('desc', '')
    if c_desc:
        p2 = tf.add_paragraph(); p2.text = c_desc; p2.font.name = FONTS["BODY_TEXT"]
        p2.font.size = Pt(9); p2.font.color.rgb = RGBColor(200, 215, 255); p2.alignment = PP_ALIGN.CENTER

    # 4방향: 상, 우, 하, 좌 — 슬라이드 경계 보장 + 균일 간격
    n = min(len(directions), 4)
    default_colors_r = ['blue', 'green', 'orange', 'red']
    card_w = Inches(2.5); card_h = Inches(1.1)
    cr_v = center_h // 2; cr_h = center_w // 2  # 상하/좌우 반경 다름
    # 동적 간격: 상하 가용 공간 기준 + 좌우는 조금 더 길게
    avail_v = int(bh) // 2 - cr_v - int(card_h)
    v_gap = max(Inches(0.15), avail_v - Inches(0.08))
    h_gap = v_gap + Inches(0.4)

    card_positions = [
        (cx, cy - cr_v - int(v_gap) - int(card_h) // 2, 'top'),
        (cx + cr_h + int(h_gap) + int(card_w) // 2, cy, 'right'),
        (cx, cy + cr_v + int(v_gap) + int(card_h) // 2, 'bottom'),
        (cx - cr_h - int(h_gap) - int(card_w) // 2, cy, 'left'),
    ]

    for i in range(n):
        color_key = directions[i].get('color', default_colors_r[i])
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(color_key, _SEM_BOX_STYLES['primary'])

        card_cx, card_cy, direction = card_positions[i]
        card_x = card_cx - int(card_w) // 2; card_y = card_cy - int(card_h) // 2

        # 연결선 (중앙 노드 테두리 → 카드 테두리)
        if direction == 'top':
            lx1, ly1 = cx, cy - cr_v
            lx2, ly2 = cx, card_cy + int(card_h) // 2
        elif direction == 'bottom':
            lx1, ly1 = cx, cy + cr_v
            lx2, ly2 = cx, card_cy - int(card_h) // 2
        elif direction == 'right':
            lx1, ly1 = cx + cr_h, cy
            lx2, ly2 = card_cx - int(card_w) // 2, cy
        else:  # left
            lx1, ly1 = cx - cr_h, cy
            lx2, ly2 = card_cx + int(card_w) // 2, cy

        connector = slide.shapes.add_connector(1, lx1, ly1, lx2, ly2)
        connector.line.color.rgb = line_c; connector.line.width = Pt(2.5)

        # 카드
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_x, card_y, int(card_w), int(card_h))
        card.fill.solid(); card.fill.fore_color.rgb = fill_c
        card.line.color.rgb = line_c; card.line.width = Pt(2.0)
        tf = card.text_frame; tf.clear(); tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.15); tf.margin_right = Inches(0.15)
        p = tf.paragraphs[0]; p.text = directions[i].get('label', ''); p.font.name = FONTS["BODY_TITLE"]
        p.font.bold = True; p.font.size = Pt(14); p.font.color.rgb = line_c; p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(4)
        d_desc = directions[i].get('desc', '')
        if d_desc:
            p2 = tf.add_paragraph(); p2.text = d_desc; p2.font.name = FONTS["BODY_TEXT"]
            p2.font.size = Pt(11); p2.font.color.rgb = text_c; p2.alignment = PP_ALIGN.CENTER


# 30. Funnel (퍼널 다이어그램)
def render_funnel(slide, data):
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    stages = content.get('stages', [])
    if not stages: return
    n = len(stages); gap = Inches(0.06); level_h = (bh - gap * (n - 1)) / n; center_x = bx + bw / 2
    max_w = bw * 0.95; min_w = bw * 0.25

    for i, stage in enumerate(stages):
        ratio = i / max(n - 1, 1); level_w = max_w - (max_w - min_w) * ratio
        level_x = center_x - level_w / 2; level_y = by + i * (level_h + gap)
        color_key = stage.get('color', 'primary')
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(color_key, _SEM_BOX_STYLES['primary'])

        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, int(level_x), int(level_y), int(level_w), int(level_h))
        shp.fill.solid(); shp.fill.fore_color.rgb = fill_c; shp.line.color.rgb = line_c; shp.line.width = Pt(2.0)
        tf = shp.text_frame; tf.clear(); tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.2); tf.margin_right = Inches(0.2)

        value = stage.get('value', '')
        label = stage.get('label', '')
        if value:
            p = tf.paragraphs[0]; p.text = f"{label}  {value}"; p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = line_c; p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(2)
        else:
            p = tf.paragraphs[0]; p.text = label; p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = line_c; p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(2)
        desc = stage.get('desc', '')
        if desc:
            p2 = tf.add_paragraph(); p2.text = desc; p2.font.name = FONTS["BODY_TEXT"]; p2.font.size = Pt(11); p2.font.color.rgb = text_c; p2.alignment = PP_ALIGN.CENTER


# 31. Zigzag Timeline (지그재그 타임라인)
def render_zigzag_timeline(slide, data):
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    steps = content.get('steps', [])
    if not steps: return
    n = len(steps)

    # 카드 크기 및 레이아웃 계산
    card_w = min(Inches(2.0), int(bw / max(n, 4) * 1.3)); card_h = Inches(1.3)
    top_y = int(by); bottom_y = int(by + bh) - card_h
    mid_y = int(by) + int(bh) // 2
    col_step = int(bw) / max(n, 1)

    step_colors = [
        (COLORS["PRIMARY"], COLORS["SEM_BLUE_BG"]), (COLORS["SEM_GREEN"], COLORS["SEM_GREEN_BG"]),
        (COLORS["SEM_ORANGE"], COLORS["SEM_ORANGE_BG"]), (COLORS["SEM_RED"], COLORS["SEM_RED_BG"]),
        (RGBColor(30, 58, 138), RGBColor(239, 246, 255)), (RGBColor(4, 120, 87), RGBColor(236, 253, 245)),
        (RGBColor(194, 65, 12), RGBColor(255, 247, 237)), (RGBColor(185, 28, 28), RGBColor(254, 242, 242)),
    ]

    # 날짜 기반 파란색 자동 감지 (MM/DD 형식 날짜가 오늘 이전이면 PRIMARY 파란색 솔리드)
    import datetime as _dt; import re as _re
    _today = _dt.date.today()
    def _step_start_date(ds):
        m = _re.match(r'(\d{2})/(\d{2})', ds.strip())
        if m:
            try: return _dt.date(2026, int(m.group(1)), int(m.group(2)))
            except: pass
        return None

    # 중앙 가로선 (배경 타임라인)
    line_shp = slide.shapes.add_connector(1, int(bx) + Inches(0.2), mid_y, int(bx + bw) - Inches(0.2), mid_y)
    line_shp.line.color.rgb = COLORS["BORDER"]; line_shp.line.width = Pt(2.0)
    line_shp.line.dash_style = 2  # dash

    for i, step in enumerate(steps):
        accent, bg = step_colors[i % len(step_colors)]

        # MM/DD 날짜 기반 3단계 색상 (원본 p3 스타일)
        # 과거: SEM_BLUE_BG(EFF6FF) fill + PRIMARY(0043DA) border (시작된 단계)
        # 미래: BG_BOX(F8F9FA) fill + GRAY(505050) border (예정 단계)
        # 날짜 없음: 기존 인덱스 색상 유지
        _sd = _step_start_date(step.get('date', ''))
        _is_past = _sd is not None and _sd <= _today
        if _is_past:
            fill_c = COLORS["SEM_BLUE_BG"]; line_c = COLORS["PRIMARY"]  # EFF6FF fill, 0043DA border
            dt_clr = COLORS["PRIMARY"]; title_clr = COLORS["PRIMARY"]; desc_clr = COLORS["GRAY"]
        elif _sd is not None:
            fill_c = COLORS["BG_BOX"]; line_c = COLORS["GRAY"]  # F8F9FA fill, 505050 border
            dt_clr = COLORS["GRAY"]; title_clr = COLORS["GRAY"]; desc_clr = COLORS["GRAY"]
        else:
            fill_c = bg; line_c = accent  # 날짜 없는 경우: 인덱스 색상
            dt_clr = accent; title_clr = accent; desc_clr = COLORS["GRAY"]

        cx = int(bx) + int(col_step * i) + int(col_step - card_w) // 2
        is_top = (i % 2 == 0)
        card_y = top_y if is_top else bottom_y

        # 수직 연결선 (카드 → 중앙선)
        conn_x = cx + card_w // 2
        if is_top:
            conn_y1 = card_y + card_h; conn_y2 = mid_y
        else:
            conn_y1 = mid_y; conn_y2 = card_y
        connector = slide.shapes.add_connector(1, conn_x, conn_y1, conn_x, conn_y2)
        connector.line.color.rgb = line_c; connector.line.width = Pt(1.5)

        # 중앙선 위 마커 원
        marker_size = Inches(0.2)
        marker = slide.shapes.add_shape(MSO_SHAPE.OVAL, conn_x - int(marker_size) // 2, mid_y - int(marker_size) // 2, int(marker_size), int(marker_size))
        marker.fill.solid(); marker.fill.fore_color.rgb = line_c; marker.line.color.rgb = line_c

        # 카드
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, card_h)
        shp.fill.solid(); shp.fill.fore_color.rgb = fill_c; shp.line.color.rgb = line_c; shp.line.width = Pt(2.0)
        tf = shp.text_frame; tf.clear(); tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)

        date = step.get('date', '')
        if date:
            p0 = tf.paragraphs[0]; p0.text = date; p0.font.name = FONTS["BODY_TEXT"]; p0.font.size = Pt(9); p0.font.color.rgb = dt_clr; p0.alignment = PP_ALIGN.CENTER; p0.space_after = Pt(2)
            p1 = tf.add_paragraph()
        else:
            p1 = tf.paragraphs[0]
        p1.text = step.get('title', ''); p1.font.name = FONTS["BODY_TITLE"]; p1.font.bold = True; p1.font.size = Pt(12); p1.font.color.rgb = title_clr; p1.alignment = PP_ALIGN.CENTER; p1.space_after = Pt(2)
        desc = step.get('desc', '')
        if desc:
            p2 = tf.add_paragraph(); p2.text = desc; p2.font.name = FONTS["BODY_TEXT"]; p2.font.size = Pt(9); p2.font.color.rgb = desc_clr; p2.alignment = PP_ALIGN.CENTER


# 32. Fishbone Cause-Effect (피쉬본 원인-결과)
def render_fishbone_cause_effect(slide, data):
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    effect = content.get('effect', ''); categories = content.get('categories', [])
    if not categories: return
    n = len(categories)

    spine_y = int(by) + int(bh) // 2
    spine_x1 = int(bx) + Inches(0.2); spine_x2 = int(bx + bw) - Inches(2.2)
    spine = slide.shapes.add_connector(1, spine_x1, spine_y, spine_x2, spine_y)
    spine.line.color.rgb = COLORS["PRIMARY"]; spine.line.width = Pt(3.0)

    # 효과 박스 (오른쪽 화살촉)
    eff_w = Inches(2.2); eff_h = Inches(0.9)
    eff_shp = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, spine_x2 - Inches(0.2), spine_y - int(eff_h) // 2, int(eff_w), int(eff_h))
    eff_shp.fill.solid(); eff_shp.fill.fore_color.rgb = COLORS["PRIMARY"]; eff_shp.line.color.rgb = COLORS["PRIMARY"]
    tf = eff_shp.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE; tf.margin_left = Inches(0.3)
    p = tf.paragraphs[0]; p.text = effect; p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(14); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment = PP_ALIGN.CENTER

    usable_w = spine_x2 - spine_x1 - Inches(0.3); spacing = int(usable_w) / max(n, 1)
    branch_h = int(bh * 0.18); default_colors = ['blue', 'green', 'orange', 'red', 'primary', 'gray']

    for i, cat in enumerate(categories):
        is_top = (i % 2 == 0)
        cx = spine_x1 + Inches(0.3) + int(spacing * i) + int(spacing * 0.5)
        color_key = cat.get('color', default_colors[i % len(default_colors)])
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(color_key, _SEM_BOX_STYLES['primary'])
        end_y = spine_y - branch_h if is_top else spine_y + branch_h

        conn = slide.shapes.add_connector(1, cx, spine_y, cx, end_y)
        conn.line.color.rgb = line_c; conn.line.width = Pt(2.0)

        card_w = min(Inches(1.8), int(spacing * 0.85)); card_h = Inches(0.9)
        card_x = cx - int(card_w) // 2; card_y = (end_y - int(card_h)) if is_top else end_y
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_x, card_y, int(card_w), int(card_h))
        card.fill.solid(); card.fill.fore_color.rgb = fill_c; card.line.color.rgb = line_c; card.line.width = Pt(1.5)
        tf = card.text_frame; tf.clear(); tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.08); tf.margin_right = Inches(0.08); tf.margin_top = Inches(0.05)
        p = tf.paragraphs[0]; p.text = cat.get('label', ''); p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(10); p.font.color.rgb = line_c; p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(2)
        for cause in cat.get('causes', [])[:3]:
            p2 = tf.add_paragraph(); p2.text = f"• {cause}"; p2.font.name = FONTS["BODY_TEXT"]
            p2.font.size = Pt(7); p2.font.color.rgb = text_c; p2.alignment = PP_ALIGN.LEFT; p2.space_after = Pt(0)


# 33. Org Chart (조직도/트리 계층)
def render_org_chart(slide, data):
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    root = content.get('root', {}); children = content.get('children', [])
    if not root: return

    # 루트 노드
    root_w = Inches(2.5); root_h = Inches(0.8)
    root_x = int(bx) + int(bw) // 2 - int(root_w) // 2; root_y = int(by)
    root_shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, root_x, root_y, int(root_w), int(root_h))
    root_shp.fill.solid(); root_shp.fill.fore_color.rgb = COLORS["PRIMARY"]; root_shp.line.color.rgb = COLORS["PRIMARY"]
    tf = root_shp.text_frame; tf.clear(); tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = root.get('label', ''); p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
    p.font.size = Pt(18); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment = PP_ALIGN.CENTER
    if root.get('desc'):
        p2 = tf.add_paragraph(); p2.text = root['desc']; p2.font.name = FONTS["BODY_TEXT"]
        p2.font.size = Pt(12); p2.font.color.rgb = COLORS["BG_WHITE"]; p2.alignment = PP_ALIGN.CENTER

    if not children: return
    n = len(children)

    # 자식 노드 배치 (전체 폭 사용)
    connector_gap = Inches(0.5)
    child_y = int(by) + int(root_h) + int(connector_gap)
    gap_between = Inches(0.2)
    child_w = (int(bw) - int(gap_between) * (n - 1)) // max(n, 1)
    child_h = int(bh) - int(root_h) - int(connector_gap)
    total_w = child_w * n + int(gap_between) * (n - 1)
    start_x = int(bx) + (int(bw) - total_w) // 2
    root_bottom_y = root_y + int(root_h)
    mid_y = root_bottom_y + int(connector_gap) // 2
    root_cx = root_x + int(root_w) // 2

    # 수직선 + 수평 연결선
    vert = slide.shapes.add_connector(1, root_cx, root_bottom_y, root_cx, mid_y)
    vert.line.color.rgb = COLORS["PRIMARY"]; vert.line.width = Pt(2.0)
    first_cx = start_x + child_w // 2; last_cx = start_x + total_w - child_w // 2
    if n > 1:
        horiz = slide.shapes.add_connector(1, first_cx, mid_y, last_cx, mid_y)
        horiz.line.color.rgb = COLORS["PRIMARY"]; horiz.line.width = Pt(2.0)

    # 자식 수에 따라 폰트 크기 동적 조절
    if n >= 5:
        label_sz = Pt(12); desc_sz = Pt(10); item_sz = Pt(9)
    elif n >= 4:
        label_sz = Pt(14); desc_sz = Pt(11); item_sz = Pt(10)
    else:
        label_sz = Pt(16); desc_sz = Pt(12); item_sz = Pt(11)

    default_colors = ['blue', 'green', 'orange', 'red', 'primary', 'gray']
    for i, child in enumerate(children):
        cx = start_x + int((child_w + int(gap_between)) * i)
        ccx = cx + child_w // 2
        cv = slide.shapes.add_connector(1, ccx, mid_y, ccx, child_y)
        cv.line.color.rgb = COLORS["PRIMARY"]; cv.line.width = Pt(2.0)
        color_key = child.get('color', default_colors[i % len(default_colors)])
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(color_key, _SEM_BOX_STYLES['primary'])

        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, child_y, child_w, child_h)
        shp.fill.solid(); shp.fill.fore_color.rgb = fill_c; shp.line.color.rgb = line_c; shp.line.width = Pt(2.0)
        tf = shp.text_frame; tf.clear(); tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.12); tf.margin_right = Inches(0.12); tf.margin_top = Inches(0.15)
        p = tf.paragraphs[0]; p.text = child.get('label', ''); p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = label_sz; p.font.color.rgb = line_c; p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(4)
        desc = child.get('desc', '')
        if desc:
            p2 = tf.add_paragraph(); p2.text = desc; p2.font.name = FONTS["BODY_TEXT"]
            p2.font.size = desc_sz; p2.font.color.rgb = text_c; p2.alignment = PP_ALIGN.CENTER; p2.space_after = Pt(6)
        for item in child.get('items', [])[:3]:
            p3 = tf.add_paragraph(); p3.text = f"• {item}"; p3.font.name = FONTS["BODY_TEXT"]
            p3.font.size = item_sz; p3.font.color.rgb = text_c; p3.alignment = PP_ALIGN.LEFT; p3.space_after = Pt(2)


```
