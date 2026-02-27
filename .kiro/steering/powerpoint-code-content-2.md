# PowerPoint Generation - Content Renderers Source Code (Part 2)

**Part of**: [powerpoint-guide.md](./powerpoint-guide.md) 시스템 명세
**Continues from**: [powerpoint-code-content.md](./powerpoint-code-content.md)
**File**: `powerpoint_content.py` — Layouts 14~27 + Diagram helpers + Router

---

## powerpoint_content.py (continued) — Diagram Helpers & Layouts 14~27

아래 코드는 `powerpoint-code-content.md`의 코드에 이어서 같은 파일(`powerpoint_content.py`)에 포함됩니다.

```python
# 14. Detail Sections (KMS PPT 슬라이드 2~4 참조)

# ── 다이어그램 공통 색상 팔레트 ──
_SEM_BOX_STYLES = {
    'gray':    (RGBColor(248, 249, 250), RGBColor(150, 150, 150), RGBColor(33, 33, 33)),
    'red':     (RGBColor(254, 242, 242), RGBColor(185, 28, 28), RGBColor(127, 29, 29)),
    'orange':  (RGBColor(255, 247, 237), RGBColor(194, 65, 12), RGBColor(154, 52, 18)),
    'green':   (RGBColor(236, 253, 245), RGBColor(4, 120, 87), RGBColor(6, 95, 70)),
    'blue':    (RGBColor(239, 246, 255), RGBColor(30, 58, 138), RGBColor(30, 64, 175)),
    'primary': (RGBColor(239, 246, 255), RGBColor(0, 67, 218), RGBColor(30, 64, 175)),
}

def _diagram_box(slide, x, y, w, h, label, color='gray', font_size=13):
    """공통 다이어그램 박스 그리기"""
    fill_c, line_c, text_c = _SEM_BOX_STYLES.get(color, _SEM_BOX_STYLES['gray'])
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = fill_c
    shp.line.color.rgb = line_c; shp.line.width = Pt(1.5)

    tf = shp.text_frame; tf.clear(); tf.word_wrap = True
    tf.margin_left = Inches(0.12); tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.06); tf.margin_bottom = Inches(0.06)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    lines = label.split('\n')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = FONTS["BODY_TEXT"]
        p.font.size = Pt(font_size) if i == 0 else Pt(font_size - 2)
        p.font.bold = (i == 0)
        p.font.color.rgb = text_c if i == 0 else COLORS["GRAY"]
        p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(2)
    return shp

def _diagram_arrow_label(slide, x, y, w, h, label, direction='down'):
    """화살표 라벨 (방향 지원)"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.clear(); tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    prefix = {'down': '⬇', 'right': '➡', 'left': '⬅', 'up': '⬆'}.get(direction, '⬇')
    p = tf.paragraphs[0]
    p.text = f"{prefix} {label}" if label else prefix
    p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(11)
    p.font.color.rgb = COLORS["GRAY"]; p.alignment = PP_ALIGN.CENTER

def _diagram_shape_arrow(slide, x, y, w, h, direction='down'):
    """실제 화살표 shape (방향 지원)"""
    shape_type = {
        'down': MSO_SHAPE.DOWN_ARROW, 'right': MSO_SHAPE.RIGHT_ARROW,
        'left': MSO_SHAPE.LEFT_ARROW, 'up': MSO_SHAPE.UP_ARROW,
    }.get(direction, MSO_SHAPE.DOWN_ARROW)
    arrow = slide.shapes.add_shape(shape_type, x, y, w, h)
    arrow.fill.solid(); arrow.fill.fore_color.rgb = COLORS["PRIMARY"]
    arrow.line.color.rgb = COLORS["PRIMARY"]
    return arrow


def _draw_diagram_flow(slide, rx, ry, rw, rh, items):
    """type=flow: 수직 흐름도 (박스 + 화살표 라벨, 개수 자동 대응)"""
    boxes = [it for it in items if it.get('type') != 'arrow']
    arrows = [it for it in items if it.get('type') == 'arrow']

    arrow_h = Inches(0.3)
    gap = Inches(0.06)
    total_h = arrow_h * len(arrows) + gap * (len(items) - 1)
    box_h = (rh - total_h) / max(len(boxes), 1)

    pad_x = Inches(0.1)
    bw = rw - pad_x * 2; bx = rx + pad_x
    cy = ry

    for item in items:
        if item.get('type') == 'arrow':
            _diagram_arrow_label(slide, bx, cy, bw, arrow_h, item.get('label', ''))
            cy += arrow_h + gap
        else:
            _diagram_box(slide, bx, cy, bw, box_h, item.get('label', ''), item.get('color', 'gray'))
            cy += box_h + gap


def _draw_diagram_layers(slide, rx, ry, rw, rh, layers):
    """type=layers: 수평 계층 다이어그램 (아키텍처 티어, 분리된 영역)

    layers: [
        {"title": "Data Layer", "desc": "Encrypted — 변경 없음", "color": "green"},
        {"title": "Key Layer", "desc": "KMS CMK로 보호", "color": "blue", "items": ["CMK", "Data Key"]},
    ]
    """
    n = len(layers)
    gap = Inches(0.15)
    layer_h = (rh - gap * (n - 1)) / n
    pad_x = Inches(0.1)

    for i, layer in enumerate(layers):
        ly = ry + i * (layer_h + gap)
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(layer.get('color', 'gray'), _SEM_BOX_STYLES['gray'])

        # 외곽 테두리
        outer = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, rx + pad_x, ly, rw - pad_x * 2, layer_h)
        outer.fill.solid(); outer.fill.fore_color.rgb = fill_c
        outer.line.color.rgb = line_c; outer.line.width = Pt(2.0)

        # 레이어 제목
        title_h = Inches(0.3)
        tb = slide.shapes.add_textbox(rx + pad_x + Inches(0.15), ly + Inches(0.08), rw - pad_x * 2 - Inches(0.3), title_h)
        p = tb.text_frame.paragraphs[0]
        p.text = layer.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(12); p.font.color.rgb = line_c

        # 레이어 설명
        desc = layer.get('desc', '')
        if desc:
            tb_d = slide.shapes.add_textbox(rx + pad_x + Inches(0.15), ly + Inches(0.35), rw - pad_x * 2 - Inches(0.3), Inches(0.25))
            p_d = tb_d.text_frame.paragraphs[0]
            p_d.text = desc
            p_d.font.name = FONTS["BODY_TEXT"]; p_d.font.size = Pt(10)
            p_d.font.color.rgb = text_c

        # 내부 아이템 박스 (가로 배치)
        sub_items = layer.get('items', [])
        if sub_items:
            inner_y = ly + Inches(0.65)
            inner_h = layer_h - Inches(0.8)
            inner_gap = Inches(0.1)
            inner_w = (rw - pad_x * 2 - Inches(0.3) - inner_gap * (len(sub_items) - 1)) / len(sub_items)

            for j, sub in enumerate(sub_items):
                sx = rx + pad_x + Inches(0.15) + j * (inner_w + inner_gap)
                sub_label = sub if isinstance(sub, str) else sub.get('label', '')
                sub_color = 'gray' if isinstance(sub, str) else sub.get('color', 'gray')
                _diagram_box(slide, sx, inner_y, inner_w, inner_h, sub_label, sub_color, font_size=11)


def _draw_diagram_compare(slide, rx, ry, rw, rh, sides):
    """type=compare: 좌우 비교 다이어그램

    sides: [
        {"title": "Before", "items": [...], "color": "red"},
        {"title": "After", "items": [...], "color": "green"},
    ]
    """
    n = len(sides)
    gap = Inches(0.15)
    side_w = (rw - gap * (n - 1)) / n
    pad_x = Inches(0.05)

    for i, side in enumerate(sides):
        sx = rx + i * (side_w + gap)
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(side.get('color', 'gray'), _SEM_BOX_STYLES['gray'])

        # 헤더
        header_h = Inches(0.45)
        hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, ry, side_w, header_h)
        hdr.fill.solid(); hdr.fill.fore_color.rgb = line_c
        hdr.line.color.rgb = line_c
        tf = hdr.text_frame; tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.text = side.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(13); p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

        # 아이템 박스들 (수직 배치)
        items = side.get('items', [])
        if items:
            item_y = ry + header_h + Inches(0.1)
            item_gap = Inches(0.08)
            item_h = (rh - header_h - Inches(0.1) - item_gap * (len(items) - 1)) / len(items)

            for j, item in enumerate(items):
                iy = item_y + j * (item_h + item_gap)
                label = item if isinstance(item, str) else item.get('label', '')
                color = side.get('color', 'gray') if isinstance(item, str) else item.get('color', side.get('color', 'gray'))
                _diagram_box(slide, sx + pad_x, iy, side_w - pad_x * 2, item_h, label, color, font_size=11)


def _draw_diagram_process(slide, rx, ry, rw, rh, steps):
    """type=process: 좌→우 가로 프로세스 (쉐브론 + 설명)

    steps: [
        {"title": "Step 1", "desc": "설명", "color": "gray"},
        ...
    ]
    """
    n = len(steps)
    gap = Inches(0.08)
    step_w = (rw - gap * (n - 1)) / n
    chevron_h = Inches(0.6)

    for i, step in enumerate(steps):
        sx = rx + i * (step_w + gap)
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(step.get('color', 'gray'), _SEM_BOX_STYLES['gray'])

        # 쉐브론 헤더
        chv = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, sx, ry, step_w, chevron_h)
        chv.fill.solid(); chv.fill.fore_color.rgb = line_c
        chv.line.color.rgb = line_c
        tf = chv.text_frame; tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.text = step.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(10); p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

        # 설명 박스
        desc = step.get('desc', '')
        if desc:
            desc_y = ry + chevron_h + Inches(0.08)
            desc_h = rh - chevron_h - Inches(0.08)
            _diagram_box(slide, sx, desc_y, step_w, desc_h, desc, step.get('color', 'gray'), font_size=10)


def _draw_right_diagram(slide, rx, ry, rw, rh, diagram_data):
    """우측 다이어그램 라우터 — type에 따라 다른 렌더러 호출

    지원 type:
    - flow: 수직 박스+화살표 흐름도 (기본값)
    - layers: 수평 계층 다이어그램 (아키텍처 티어)
    - compare: 좌우 비교 다이어그램
    - process: 좌→우 가로 프로세스
    """
    # dict 형태 (type 지정)
    if isinstance(diagram_data, dict):
        d_type = diagram_data.get('type', 'flow')
        items = diagram_data.get('items', diagram_data.get('steps', diagram_data.get('layers', diagram_data.get('sides', []))))

        if d_type == 'layers':
            _draw_diagram_layers(slide, rx, ry, rw, rh, items)
        elif d_type == 'compare':
            _draw_diagram_compare(slide, rx, ry, rw, rh, items)
        elif d_type == 'process':
            _draw_diagram_process(slide, rx, ry, rw, rh, items)
        else:
            _draw_diagram_flow(slide, rx, ry, rw, rh, items)

    # list 형태 (하위 호환: flow로 처리)
    elif isinstance(diagram_data, list):
        _draw_diagram_flow(slide, rx, ry, rw, rh, diagram_data)


def render_detail_sections(slide, data):
    """좌측 멀티섹션 텍스트 + 우측 다이어그램/이미지 레이아웃

    KMS PPT 참조: 개요 → 강조 박스(의미 색상) → 조건/불릿 구조
    우측: diagram 데이터 → shape 직접 그리기 / image_path → 이미지 로드
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    gap = Inches(0.3)
    w_left = (bw - gap) * 0.5
    w_right = (bw - gap) * 0.5

    # ── 좌측 콘텐츠 높이 사전 계산 ──
    overview = content.get('overview', {})
    highlight = content.get('highlight', {})
    condition = content.get('condition', {})

    section_count = sum([1 for s in [overview, highlight, condition] if s])
    if section_count == 0:
        return

    section_gap = Inches(0.12)
    total_gap = section_gap * (section_count - 1)
    available_h = bh - total_gap

    ratios = []
    if overview: ratios.append(('overview', 0.30))
    if highlight: ratios.append(('highlight', 0.45))
    if condition: ratios.append(('condition', 0.25))

    total_ratio = sum(r[1] for r in ratios)
    section_heights = {}
    for name, ratio in ratios:
        section_heights[name] = available_h * (ratio / total_ratio)

    current_y = by

    # (1) 개요 섹션
    if overview:
        sec_h = section_heights['overview']
        tb = slide.shapes.add_textbox(bx, current_y, w_left, sec_h)
        tf = tb.text_frame; tf.word_wrap = True; tf.clear()
        tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)
        tf.margin_top = Inches(0.05); tf.margin_bottom = Inches(0.05)
        tf.vertical_anchor = MSO_ANCHOR.TOP

        p = tf.paragraphs[0]
        p.text = overview.get('title', '개요')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(16); p.font.color.rgb = COLORS["DARK_GRAY"]
        p.space_after = Pt(6)

        for line in overview.get('body', '').split('\n'):
            p = tf.add_paragraph()
            p.text = line
            p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(13)
            p.font.color.rgb = COLORS["GRAY"]; p.space_after = Pt(3)

        current_y += sec_h + section_gap

    # (2) 강조 박스 (의미 기반 색상)
    if highlight:
        sec_h = section_heights['highlight']
        color_key = highlight.get('color', 'red')
        sem_colors = {
            'red':    ("SEM_RED", "SEM_RED_BG", "SEM_RED_TEXT"),
            'orange': ("SEM_ORANGE", "SEM_ORANGE_BG", "SEM_ORANGE_TEXT"),
            'green':  ("SEM_GREEN", "SEM_GREEN_BG", "SEM_GREEN_TEXT"),
            'blue':   ("SEM_BLUE", "SEM_BLUE_BG", "SEM_BLUE_TEXT"),
        }
        title_c, bg_c, text_c = sem_colors.get(color_key, sem_colors['red'])

        hl_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, current_y, w_left, sec_h)
        hl_box.fill.solid()
        hl_box.fill.fore_color.rgb = COLORS[bg_c]
        hl_box.line.color.rgb = COLORS[title_c]
        hl_box.line.width = Pt(1.5)

        tf = hl_box.text_frame; tf.clear(); tf.word_wrap = True
        tf.margin_left = Inches(0.2); tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.04); tf.margin_bottom = Inches(0.04)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        p = tf.paragraphs[0]
        p.text = highlight.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(12); p.font.color.rgb = COLORS[title_c]
        p.space_after = Pt(3)

        for line in highlight.get('body', '').split('\n'):
            p = tf.add_paragraph()
            p.text = line
            p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(12)
            p.font.color.rgb = COLORS[text_c]; p.space_after = Pt(2)

        current_y += sec_h + section_gap

    # (3) 조건/불릿 섹션
    if condition:
        sec_h = section_heights['condition']
        tb = slide.shapes.add_textbox(bx, current_y, w_left, sec_h)
        tf = tb.text_frame; tf.word_wrap = True; tf.clear()
        tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)
        tf.margin_top = Inches(0.05); tf.margin_bottom = Inches(0.05)
        tf.vertical_anchor = MSO_ANCHOR.TOP

        p = tf.paragraphs[0]
        p.text = condition.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(12); p.font.color.rgb = COLORS["SEM_BLUE"]
        p.space_after = Pt(6)

        for bullet in condition.get('bullets', []):
            p = tf.add_paragraph()
            p.text = f"• {bullet}"
            p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(12)
            p.font.color.rgb = COLORS["GRAY"]; p.space_after = Pt(3)

    # ── 우측: diagram shape 또는 이미지 ──
    right_x = bx + w_left + gap
    rendered = False

    # 1순위: diagram 데이터 → shape로 직접 그리기
    diagram = content.get('diagram', [])
    if diagram:
        _draw_right_diagram(slide, right_x, by, w_right, bh, diagram)
        rendered = True

    # 2순위: image_path 직접 지정
    if not rendered:
        image_path = content.get('image_path', '')
        if image_path and os.path.exists(image_path):
            try:
                from PIL import Image as PILImage
                with PILImage.open(image_path) as img:
                    orig_w, orig_h = img.size
                aspect = orig_w / orig_h
                aw, ah = int(w_right), int(bh)
                if aw / aspect <= ah:
                    fw, fh = aw, int(aw / aspect)
                else:
                    fh, fw = ah, int(ah * aspect)
                cx = int(right_x) + (aw - fw) // 2
                cy = int(by) + (ah - fh) // 2
                slide.shapes.add_picture(image_path, cx, cy, width=fw, height=fh)
                rendered = True
            except ImportError:
                slide.shapes.add_picture(image_path, int(right_x), int(by), width=int(w_right), height=int(bh))
                rendered = True
            except Exception as e:
                print(f"   ⚠️ [이미지 로드 실패] {str(e)[:50]}")

    # 3순위: architecture/ 폴더 검색
    if not rendered:
        search_q = content.get('search_q', '')
        if search_q:
            img_file = os.path.join('architecture', search_q.replace(' ', '_') + '.png')
            if os.path.exists(img_file):
                try:
                    from PIL import Image as PILImage
                    with PILImage.open(img_file) as img:
                        orig_w, orig_h = img.size
                    aspect = orig_w / orig_h
                    aw, ah = int(w_right), int(bh)
                    if aw / aspect <= ah:
                        fw, fh = aw, int(aw / aspect)
                    else:
                        fh, fw = ah, int(ah * aspect)
                    cx = int(right_x) + (aw - fw) // 2
                    cy = int(by) + (ah - fh) // 2
                    slide.shapes.add_picture(img_file, cx, cy, width=fw, height=fh)
                    rendered = True
                except:
                    pass

    if not rendered:
        print(f"   ⚠️ [detail_sections] 우측 콘텐츠 없음 — diagram, image_path, 또는 architecture/ 이미지를 지정해주세요")


# 15. Table + Callout (KMS PPT 슬라이드 6 참조)
def render_table_callout(slide, data):
    """비교 테이블 + 하단 추천/결론 콜아웃 박스 레이아웃

    KMS PPT 참조: 상단에 비교 테이블, 하단에 결론/추천 강조 박스
    테이블 열 수 자동 대응 (2~5열)
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    columns = content.get('columns', [])
    rows = content.get('rows', [])
    callout = content.get('callout', {})

    if not columns:
        return

    n_cols = len(columns)

    # 공간 분배: 테이블 65%, 콜아웃 35% (콜아웃 없으면 테이블 100%)
    callout_h = Inches(1.3) if callout else 0
    callout_gap = Inches(0.2) if callout else 0
    table_h = bh - callout_h - callout_gap

    # ── 테이블 영역 ──
    gap = Inches(0.15)
    w_col = (bw - (gap * (n_cols - 1))) / n_cols

    # 헤더 행
    header_h = Inches(0.7)
    for i, col in enumerate(columns):
        x = bx + i * (w_col + gap)
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, by, w_col, header_h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = COLORS["PRIMARY"]
        shp.line.color.rgb = COLORS["PRIMARY"]
        shp.line.width = Pt(1.0)

        tf = shp.text_frame; tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = col.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(15); p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

    # 데이터 행
    if rows:
        row_area_h = table_h - header_h - Inches(0.15)
        row_h = row_area_h / len(rows)

        for row_idx, row in enumerate(rows):
            row_y = by + header_h + Inches(0.15) + (row_idx * row_h)
            values = row.get('values', [])

            for col_idx in range(n_cols):
                x = bx + col_idx * (w_col + gap)
                value = values[col_idx] if col_idx < len(values) else ''

                shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, row_y, w_col, row_h - Inches(0.05))
                shp.fill.solid()
                shp.fill.fore_color.rgb = COLORS["BG_WHITE"]
                shp.line.color.rgb = COLORS["BORDER"]
                shp.line.width = Pt(1.0)

                tf = shp.text_frame; tf.clear()
                tf.margin_left = Inches(0.15); tf.margin_right = Inches(0.15)
                tf.margin_top = Inches(0.08); tf.margin_bottom = Inches(0.08)
                tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE

                p = tf.paragraphs[0]
                p.text = str(value)
                p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(13)
                p.font.color.rgb = COLORS["BLACK"]; p.alignment = PP_ALIGN.CENTER

    # ── 콜아웃 박스 (하단 추천/결론) ──
    if callout:
        callout_y = by + table_h + callout_gap

        # 배경 박스
        cb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, callout_y, bw, callout_h)
        cb.fill.solid()
        cb.fill.fore_color.rgb = COLORS["CALLOUT_BG"]
        cb.line.color.rgb = COLORS["CALLOUT_BG"]

        # 아이콘 (이모지 또는 텍스트)
        icon_text = callout.get('icon', '💡')
        icon_w = Inches(0.7)
        tb_icon = slide.shapes.add_textbox(bx + Inches(0.25), callout_y + Inches(0.15), icon_w, Inches(0.6))
        p = tb_icon.text_frame.paragraphs[0]
        p.text = icon_text
        p.font.size = Pt(30); p.alignment = PP_ALIGN.CENTER

        # 제목 + 본문
        text_x = bx + Inches(1.1)
        text_w = bw - Inches(1.5)

        # 콜아웃 제목
        tb_title = slide.shapes.add_textbox(text_x, callout_y + Inches(0.15), text_w, Inches(0.4))
        p = tb_title.text_frame.paragraphs[0]
        p.text = callout.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(16); p.font.color.rgb = COLORS["BG_WHITE"]

        # 콜아웃 본문
        callout_body = callout.get('body', '')
        if callout_body:
            tb_body = slide.shapes.add_textbox(text_x, callout_y + Inches(0.55), text_w, callout_h - Inches(0.7))
            tf = tb_body.text_frame; tf.word_wrap = True
            for i, line in enumerate(callout_body.split('\n')):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = line
                p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(14)
                p.font.color.rgb = COLORS["CALLOUT_TEXT"]; p.space_after = Pt(3)


# 16. Full Image (풀와이드 이미지/다이어그램)
def render_full_image(slide, data):
    """이미지/다이어그램이 슬라이드 본문 전체를 차지하는 레이아웃

    data.data.data:
        image_path: 이미지 파일 경로
        search_q: architecture/ 폴더 검색어 (image_path 없을 때)
        caption: 하단 캡션 (선택)
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    caption = content.get('caption', '')
    caption_h = Inches(0.45) if caption else 0
    caption_gap = Inches(0.1) if caption else 0
    img_h = bh - caption_h - caption_gap

    # 이미지 로드 시도
    img_loaded = False
    image_path = content.get('image_path', '')

    # 1순위: 직접 경로
    if image_path and os.path.exists(image_path):
        img_loaded = _place_image_centered(slide, image_path, bx, by, bw, img_h)

    # 2순위: architecture/ 폴더 검색
    if not img_loaded:
        search_q = content.get('search_q', '')
        if search_q:
            img_file = os.path.join('architecture', search_q.replace(' ', '_') + '.png')
            if os.path.exists(img_file):
                img_loaded = _place_image_centered(slide, img_file, bx, by, bw, img_h)

    # 3순위: screenshots/ 폴더 검색
    if not img_loaded:
        search_q = content.get('search_q', '')
        if search_q:
            img_file = os.path.join('screenshots', search_q.replace(' ', '_') + '.png')
            if os.path.exists(img_file):
                img_loaded = _place_image_centered(slide, img_file, bx, by, bw, img_h)

    # 폴백: 회색 박스
    if not img_loaded:
        placeholder = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, bw, img_h)
        placeholder.fill.solid()
        placeholder.fill.fore_color.rgb = COLORS["BG_BOX"]
        placeholder.line.color.rgb = COLORS["BORDER"]
        tf = placeholder.text_frame; tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = f"Image: {content.get('image_path', '') or content.get('search_q', 'N/A')}"
        p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(14)
        p.font.color.rgb = COLORS["GRAY"]; p.alignment = PP_ALIGN.CENTER

    # 캡션
    if caption:
        cap_y = by + img_h + caption_gap
        tb = slide.shapes.add_textbox(bx, cap_y, bw, caption_h)
        tf = tb.text_frame; tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = caption
        p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(12)
        p.font.color.rgb = COLORS["GRAY"]; p.alignment = PP_ALIGN.CENTER


def _place_image_centered(slide, image_path, area_x, area_y, area_w, area_h):
    """이미지를 영역 내 중앙에 비율 유지하며 배치 (공통 유틸리티)"""
    try:
        from PIL import Image as PILImage
        with PILImage.open(image_path) as img:
            orig_w, orig_h = img.size
        aspect = orig_w / orig_h
        aw, ah = int(area_w), int(area_h)
        if aw / aspect <= ah:
            fw, fh = aw, int(aw / aspect)
        else:
            fh, fw = ah, int(ah * aspect)
        cx = int(area_x) + (aw - fw) // 2
        cy = int(area_y) + (ah - fh) // 2
        slide.shapes.add_picture(image_path, cx, cy, width=fw, height=fh)
        return True
    except ImportError:
        slide.shapes.add_picture(image_path, int(area_x), int(area_y),
                                 width=int(area_w), height=int(area_h))
        return True
    except Exception as e:
        print(f"   ⚠️ [이미지 로드 실패] {str(e)[:50]}")
        return False


# 17. Before/After (전/후 비교)
def render_before_after(slide, data):
    """Before/After 비교 레이아웃

    좌측: Before (회색/빨강 톤) / 우측: After (파랑/녹색 톤)
    중앙에 화살표

    data.data:
        before_title, before_body: Before 패널
        after_title, after_body: After 패널
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    arrow_gap = Inches(0.8)
    w_half = (bw - arrow_gap) / 2
    label_h = Inches(0.55)
    body_gap = Inches(0.1)

    # ── Before 패널 (좌측) ──
    before_label = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, w_half, label_h)
    before_label.fill.solid()
    before_label.fill.fore_color.rgb = COLORS["SEM_RED"]
    before_label.line.color.rgb = COLORS["SEM_RED"]
    tf = before_label.text_frame; tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = content.get('before_title', 'Before')
    p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
    p.font.size = Pt(18); p.font.color.rgb = COLORS["BG_WHITE"]
    p.alignment = PP_ALIGN.CENTER

    before_body_y = by + label_h + body_gap
    before_body_h = bh - label_h - body_gap
    before_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, before_body_y, w_half, before_body_h)
    before_box.fill.solid()
    before_box.fill.fore_color.rgb = COLORS["SEM_RED_BG"]
    before_box.line.color.rgb = COLORS["SEM_RED"]
    before_box.line.width = Pt(1.5)

    tf = before_box.text_frame; tf.clear(); tf.word_wrap = True
    tf.margin_left = Inches(0.25); tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.2); tf.margin_bottom = Inches(0.2)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    for i, line in enumerate(str(content.get('before_body', '')).split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(14)
        p.font.color.rgb = COLORS["SEM_RED_TEXT"]
        p.space_after = Pt(6); p.line_spacing = 1.3

    # ── After 패널 (우측) ──
    after_x = bx + w_half + arrow_gap

    after_label = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, after_x, by, w_half, label_h)
    after_label.fill.solid()
    after_label.fill.fore_color.rgb = COLORS["SEM_GREEN"]
    after_label.line.color.rgb = COLORS["SEM_GREEN"]
    tf = after_label.text_frame; tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = content.get('after_title', 'After')
    p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
    p.font.size = Pt(18); p.font.color.rgb = COLORS["BG_WHITE"]
    p.alignment = PP_ALIGN.CENTER

    after_body_y = by + label_h + body_gap
    after_body_h = bh - label_h - body_gap
    after_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, after_x, after_body_y, w_half, after_body_h)
    after_box.fill.solid()
    after_box.fill.fore_color.rgb = COLORS["SEM_GREEN_BG"]
    after_box.line.color.rgb = COLORS["SEM_GREEN"]
    after_box.line.width = Pt(1.5)

    tf = after_box.text_frame; tf.clear(); tf.word_wrap = True
    tf.margin_left = Inches(0.25); tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.2); tf.margin_bottom = Inches(0.2)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    for i, line in enumerate(str(content.get('after_body', '')).split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(14)
        p.font.color.rgb = COLORS["SEM_GREEN_TEXT"]
        p.space_after = Pt(6); p.line_spacing = 1.3

    # ── 중앙 화살표 ──
    arrow_w = Inches(1.0); arrow_h_size = Inches(1.0)
    arrow_x = bx + w_half + (arrow_gap - arrow_w) / 2
    arrow_y = by + (bh / 2) - (arrow_h_size / 2)
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, arrow_x, arrow_y, arrow_w, arrow_h_size)
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = COLORS["PRIMARY"]
    arrow.line.color.rgb = COLORS["PRIMARY"]


# 18. Icon Grid (6~9 아이콘 그리드)
def render_icon_grid(slide, data):
    """아이콘 + 제목 + 설명 그리드 (6~9개 아이템)

    자동 레이아웃: 3열 x N행 (아이템 수에 따라)

    data.data.data.items: [
        {"icon": "kubernetes", "title": "제목", "desc": "설명"},
        ...
    ]
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    items = content.get('items', [])
    if not items:
        return

    # 그리드 계산: 항상 3열
    n_cols = 3
    n_rows = (len(items) + n_cols - 1) // n_cols  # ceil

    gap_x = Inches(0.25)
    gap_y = Inches(0.2)
    cell_w = (bw - gap_x * (n_cols - 1)) / n_cols
    cell_h = (bh - gap_y * (n_rows - 1)) / n_rows

    icon_size = Inches(0.55)
    text_left_margin = icon_size + Inches(0.2)

    for idx, item in enumerate(items):
        col = idx % n_cols
        row = idx // n_cols
        x = bx + col * (cell_w + gap_x)
        y = by + row * (cell_h + gap_y)

        # 셀 배경 박스
        cell_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, cell_w, cell_h)
        cell_box.fill.solid()
        cell_box.fill.fore_color.rgb = COLORS["BG_WHITE"]
        cell_box.line.color.rgb = COLORS["BORDER"]
        cell_box.line.width = Pt(1.0)

        # 아이콘 (좌측 상단)
        icon_x = x + Inches(0.15)
        icon_y = y + (cell_h - icon_size) / 2
        draw_icon_search(slide, icon_x, icon_y, icon_size, item.get('icon', ''))

        # 텍스트 (아이콘 우측)
        text_x = x + text_left_margin
        text_w = cell_w - text_left_margin - Inches(0.1)
        tb = slide.shapes.add_textbox(text_x, y + Inches(0.1), text_w, cell_h - Inches(0.2))
        tf = tb.text_frame; tf.word_wrap = True; tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)

        # 제목
        p = tf.paragraphs[0]
        p.text = item.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(14); p.font.color.rgb = COLORS["PRIMARY"]
        p.space_after = Pt(4)

        # 설명
        desc = item.get('desc', '')
        if desc:
            p2 = tf.add_paragraph()
            p2.text = desc
            p2.font.name = FONTS["BODY_TEXT"]; p2.font.size = Pt(11)
            p2.font.color.rgb = COLORS["GRAY"]
            p2.line_spacing = 1.2


# 19. Numbered List (번호형 세로 리스트)
def render_numbered_list(slide, data):
    """번호형 세로 스텝 리스트

    좌측 큰 번호 원형 + 우측 제목/설명

    data.data.data.items: [
        {"title": "항목 제목", "desc": "항목 설명"},
        ...
    ]
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    items = content.get('items', [])
    if not items:
        return

    gap = Inches(0.15)
    item_h = (bh - gap * (len(items) - 1)) / len(items)
    badge_size = Inches(0.65)
    text_left = badge_size + Inches(0.3)

    for i, item in enumerate(items):
        y = by + i * (item_h + gap)

        # 배경 바 (연한 회색)
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, y, bw, item_h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = COLORS["BG_BOX"] if i % 2 == 0 else COLORS["BG_WHITE"]
        bar.line.color.rgb = COLORS["BORDER"]
        bar.line.width = Pt(1.0)

        # 번호 배지
        badge_x = bx + Inches(0.2)
        badge_y = y + (item_h - badge_size) / 2
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, badge_x, badge_y, badge_size, badge_size)
        badge.fill.solid()
        badge.fill.fore_color.rgb = COLORS["PRIMARY"]
        badge.line.color.rgb = COLORS["PRIMARY"]

        tf_badge = badge.text_frame; tf_badge.clear()
        tf_badge.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_badge.paragraphs[0]
        p.text = str(i + 1)
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(22); p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

        # 텍스트 (제목 + 설명)
        text_x = bx + text_left + Inches(0.1)
        text_w = bw - text_left - Inches(0.3)
        tb = slide.shapes.add_textbox(text_x, y + Inches(0.1), text_w, item_h - Inches(0.2))
        tf = tb.text_frame; tf.word_wrap = True; tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)

        # 제목
        p = tf.paragraphs[0]
        p.text = item.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(16); p.font.color.rgb = COLORS["DARK_GRAY"]
        p.space_after = Pt(4)

        # 설명
        desc = item.get('desc', '')
        if desc:
            for j, line in enumerate(desc.split('\n')):
                p2 = tf.add_paragraph()
                p2.text = line
                p2.font.name = FONTS["BODY_TEXT"]; p2.font.size = Pt(13)
                p2.font.color.rgb = COLORS["GRAY"]
                p2.space_after = Pt(2); p2.line_spacing = 1.2


# 20. Stats Dashboard (KPI/대형 숫자 표시)
def render_stats_dashboard(slide, data):
    """KPI/대형 숫자 대시보드 레이아웃

    3~4개 메트릭을 큰 숫자로 강조 표시

    data.data.data.metrics: [
        {"value": "99.9", "unit": "%", "label": "가용성", "desc": "연간 SLA 기준"},
        ...
    ]
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    metrics = content.get('metrics', [])
    if not metrics:
        return

    n = len(metrics)
    gap = Inches(0.25)
    card_w = (bw - gap * (n - 1)) / n

    # 색상 팔레트 (순환)
    accent_colors = [
        (COLORS["PRIMARY"], COLORS["SEM_BLUE_BG"]),
        (COLORS["SEM_GREEN"], COLORS["SEM_GREEN_BG"]),
        (COLORS["SEM_ORANGE"], COLORS["SEM_ORANGE_BG"]),
        (COLORS["SEM_RED"], COLORS["SEM_RED_BG"]),
    ]

    for i, metric in enumerate(metrics):
        x = bx + i * (card_w + gap)
        accent, bg = accent_colors[i % len(accent_colors)]

        # 카드 배경
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, by, card_w, bh)
        card.fill.solid()
        card.fill.fore_color.rgb = bg
        card.line.color.rgb = accent
        card.line.width = Pt(2.0)

        # 레이아웃: 상단 65% 숫자, 하단 35% 라벨+설명
        number_h = bh * 0.55
        label_h = bh * 0.45

        # 큰 숫자 + 단위
        tb_num = slide.shapes.add_textbox(x, by + Inches(0.2), card_w, number_h)
        tf = tb_num.text_frame; tf.clear(); tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        p = tf.paragraphs[0]
        value_text = str(metric.get('value', ''))
        unit_text = metric.get('unit', '')
        p.alignment = PP_ALIGN.CENTER

        # value와 unit을 별도 run으로 분리 (크기 차이)
        run_val = p.add_run()
        run_val.text = value_text
        run_val.font.name = FONTS["BODY_TITLE"]; run_val.font.bold = True
        run_val.font.size = Pt(44); run_val.font.color.rgb = accent

        if unit_text:
            run_unit = p.add_run()
            run_unit.text = unit_text
            run_unit.font.name = FONTS["BODY_TITLE"]; run_unit.font.bold = True
            run_unit.font.size = Pt(24); run_unit.font.color.rgb = accent

        # 라벨
        tb_label = slide.shapes.add_textbox(x, by + number_h, card_w, label_h)
        tf = tb_label.text_frame; tf.clear(); tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.15); tf.margin_right = Inches(0.15)

        p = tf.paragraphs[0]
        p.text = metric.get('label', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(16); p.font.color.rgb = COLORS["DARK_GRAY"]
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(6)

        # 설명
        desc = metric.get('desc', '')
        if desc:
            p2 = tf.add_paragraph()
            p2.text = desc
            p2.font.name = FONTS["BODY_TEXT"]; p2.font.size = Pt(12)
            p2.font.color.rgb = COLORS["GRAY"]
            p2.alignment = PP_ALIGN.CENTER
            p2.line_spacing = 1.2


# 21. Quote Highlight (인용문/핵심 메시지 강조)
def render_quote_highlight(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, bw, bh)
    bg.fill.solid(); bg.fill.fore_color.rgb = COLORS["SEM_BLUE_BG"]
    bg.line.color.rgb = COLORS["PRIMARY"]; bg.line.width = Pt(2.0)

    bar_w = Inches(0.08)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, bx + Inches(0.4), by + Inches(0.5), bar_w, bh - Inches(1.0))
    bar.fill.solid(); bar.fill.fore_color.rgb = COLORS["PRIMARY"]; bar.line.color.rgb = COLORS["PRIMARY"]

    tb_mark = slide.shapes.add_textbox(bx + Inches(0.7), by + Inches(0.2), Inches(1.0), Inches(0.8))
    p = tb_mark.text_frame.paragraphs[0]; p.text = "\u201C"
    p.font.size = Pt(72); p.font.bold = True; p.font.color.rgb = COLORS["PRIMARY"]

    quote_x = bx + Inches(0.8); quote_w = bw - Inches(1.6); quote_h = bh * 0.6
    tb_quote = slide.shapes.add_textbox(quote_x, by + Inches(0.8), quote_w, quote_h)
    tf = tb_quote.text_frame; tf.word_wrap = True; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = content.get('quote', '')
    p.font.name = FONTS["BODY_TITLE"]; p.font.size = Pt(22)
    p.font.italic = True; p.font.color.rgb = COLORS["DARK_GRAY"]
    p.alignment = PP_ALIGN.LEFT; p.line_spacing = 1.4

    author = content.get('author', ''); role = content.get('role', '')
    if author:
        author_y = by + bh - Inches(0.9)
        tb_author = slide.shapes.add_textbox(quote_x, author_y, quote_w, Inches(0.7))
        tf = tb_author.text_frame; tf.word_wrap = True; tf.clear()
        p = tf.paragraphs[0]; p.text = f"— {author}"
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = COLORS["PRIMARY"]
        if role:
            p2 = tf.add_paragraph(); p2.text = role
            p2.font.name = FONTS["BODY_TEXT"]; p2.font.size = Pt(13); p2.font.color.rgb = COLORS["GRAY"]


# 22. Pros & Cons (장단점 비교)
def render_pros_cons(slide, data):
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    gap = Inches(0.3); w_half = (bw - gap) / 2

    subject = content.get('subject', ''); subject_h = Inches(0.55) if subject else 0
    if subject:
        tb = slide.shapes.add_textbox(bx, by, bw, subject_h); tf = tb.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.text = subject; p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(20); p.font.color.rgb = COLORS["DARK_GRAY"]; p.alignment = PP_ALIGN.CENTER

    panel_y = by + subject_h + (Inches(0.1) if subject else 0); panel_h = bh - subject_h - (Inches(0.1) if subject else 0); label_h = Inches(0.5)

    # Pros (좌측 - 녹색)
    pros_label = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, panel_y, w_half, label_h)
    pros_label.fill.solid(); pros_label.fill.fore_color.rgb = COLORS["SEM_GREEN"]; pros_label.line.color.rgb = COLORS["SEM_GREEN"]
    tf = pros_label.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = "PROS"; p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment = PP_ALIGN.CENTER

    pros_body_y = panel_y + label_h + Inches(0.1); pros_body_h = panel_h - label_h - Inches(0.1)
    pros_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, pros_body_y, w_half, pros_body_h)
    pros_box.fill.solid(); pros_box.fill.fore_color.rgb = COLORS["SEM_GREEN_BG"]; pros_box.line.color.rgb = COLORS["SEM_GREEN"]; pros_box.line.width = Pt(1.5)

    tb = slide.shapes.add_textbox(bx + Inches(0.2), pros_body_y + Inches(0.15), w_half - Inches(0.4), pros_body_h - Inches(0.3))
    tf = tb.text_frame; tf.word_wrap = True; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.TOP
    for i, item in enumerate(content.get('pros', [])):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"\u2714  {item}"; p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(14); p.font.color.rgb = COLORS["SEM_GREEN_TEXT"]; p.space_after = Pt(8); p.line_spacing = 1.3

    # Cons (우측 - 빨강)
    cons_x = bx + w_half + gap
    cons_label = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cons_x, panel_y, w_half, label_h)
    cons_label.fill.solid(); cons_label.fill.fore_color.rgb = COLORS["SEM_RED"]; cons_label.line.color.rgb = COLORS["SEM_RED"]
    tf = cons_label.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = "CONS"; p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment = PP_ALIGN.CENTER

    cons_body_y = panel_y + label_h + Inches(0.1); cons_body_h = panel_h - label_h - Inches(0.1)
    cons_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cons_x, cons_body_y, w_half, cons_body_h)
    cons_box.fill.solid(); cons_box.fill.fore_color.rgb = COLORS["SEM_RED_BG"]; cons_box.line.color.rgb = COLORS["SEM_RED"]; cons_box.line.width = Pt(1.5)

    tb = slide.shapes.add_textbox(cons_x + Inches(0.2), cons_body_y + Inches(0.15), w_half - Inches(0.4), cons_body_h - Inches(0.3))
    tf = tb.text_frame; tf.word_wrap = True; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.TOP
    for i, item in enumerate(content.get('cons', [])):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"\u2718  {item}"; p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(14); p.font.color.rgb = COLORS["SEM_RED_TEXT"]; p.space_after = Pt(8); p.line_spacing = 1.3


# 23. Do / Don't (가이드라인 레이아웃)
def render_do_dont(slide, data):
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    gap = Inches(0.3); w_half = (bw - gap) / 2; label_h = Inches(0.6)

    # DO 패널 (좌측 - 파랑)
    do_label = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, w_half, label_h)
    do_label.fill.solid(); do_label.fill.fore_color.rgb = COLORS["PRIMARY"]; do_label.line.color.rgb = COLORS["PRIMARY"]
    tf = do_label.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = "\u2714  DO — 이렇게 하세요"
    p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment = PP_ALIGN.CENTER

    do_items = content.get('do_items', [])
    if do_items:
        item_y = by + label_h + Inches(0.15); item_gap = Inches(0.1)
        item_h = (bh - label_h - Inches(0.15) - item_gap * (len(do_items) - 1)) / len(do_items)
        for i, item in enumerate(do_items):
            iy = item_y + i * (item_h + item_gap)
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, iy, w_half, item_h)
            box.fill.solid(); box.fill.fore_color.rgb = COLORS["SEM_BLUE_BG"]; box.line.color.rgb = COLORS["PRIMARY"]; box.line.width = Pt(1.0)
            text = item if isinstance(item, str) else item.get('text', ''); detail = '' if isinstance(item, str) else item.get('detail', '')
            tb = slide.shapes.add_textbox(bx + Inches(0.2), iy + Inches(0.08), w_half - Inches(0.4), item_h - Inches(0.16))
            tf = tb.text_frame; tf.word_wrap = True; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]; p.text = f"\u2714  {text}"; p.font.name = FONTS["BODY_TEXT"]; p.font.bold = True; p.font.size = Pt(13); p.font.color.rgb = COLORS["PRIMARY"]; p.space_after = Pt(2)
            if detail:
                p2 = tf.add_paragraph(); p2.text = detail; p2.font.name = FONTS["BODY_TEXT"]; p2.font.size = Pt(11); p2.font.color.rgb = COLORS["GRAY"]

    # DON'T 패널 (우측 - 빨강)
    dont_x = bx + w_half + gap
    dont_label = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, dont_x, by, w_half, label_h)
    dont_label.fill.solid(); dont_label.fill.fore_color.rgb = COLORS["SEM_RED"]; dont_label.line.color.rgb = COLORS["SEM_RED"]
    tf = dont_label.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = "\u2718  DON'T — 이렇게 하지 마세요"
    p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment = PP_ALIGN.CENTER

    dont_items = content.get('dont_items', [])
    if dont_items:
        item_y = by + label_h + Inches(0.15); item_gap = Inches(0.1)
        item_h = (bh - label_h - Inches(0.15) - item_gap * (len(dont_items) - 1)) / len(dont_items)
        for i, item in enumerate(dont_items):
            iy = item_y + i * (item_h + item_gap)
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, dont_x, iy, w_half, item_h)
            box.fill.solid(); box.fill.fore_color.rgb = COLORS["SEM_RED_BG"]; box.line.color.rgb = COLORS["SEM_RED"]; box.line.width = Pt(1.0)
            text = item if isinstance(item, str) else item.get('text', ''); detail = '' if isinstance(item, str) else item.get('detail', '')
            tb = slide.shapes.add_textbox(dont_x + Inches(0.2), iy + Inches(0.08), w_half - Inches(0.4), item_h - Inches(0.16))
            tf = tb.text_frame; tf.word_wrap = True; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]; p.text = f"\u2718  {text}"; p.font.name = FONTS["BODY_TEXT"]; p.font.bold = True; p.font.size = Pt(13); p.font.color.rgb = COLORS["SEM_RED"]; p.space_after = Pt(2)
            if detail:
                p2 = tf.add_paragraph(); p2.text = detail; p2.font.name = FONTS["BODY_TEXT"]; p2.font.size = Pt(11); p2.font.color.rgb = COLORS["GRAY"]


# 24. Split Text + Code (좌측 설명 + 우측 코드 블록)
def render_split_text_code(slide, data):
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    gap = Inches(0.3); w_left = (bw - gap) * 0.4; w_right = (bw - gap) * 0.6

    tb = slide.shapes.add_textbox(bx, by, w_left, bh)
    tf = tb.text_frame; tf.word_wrap = True; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1); tf.margin_top = Inches(0.2); tf.margin_bottom = Inches(0.2)

    desc = content.get('description', '')
    if desc:
        for i, line in enumerate(desc.split('\n')):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line; p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(15); p.font.color.rgb = COLORS["DARK_GRAY"]; p.space_after = Pt(6); p.line_spacing = 1.3

    bullets = content.get('bullets', [])
    if bullets:
        if desc: p_gap = tf.add_paragraph(); p_gap.text = ""; p_gap.space_after = Pt(8)
        for i, bullet in enumerate(bullets):
            p = tf.add_paragraph() if (desc or i > 0) else tf.paragraphs[0]
            p.text = f"• {bullet}"; p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(14); p.font.color.rgb = COLORS["BLACK"]; p.space_after = Pt(6); p.line_spacing = 1.2

    code_x = bx + w_left + gap
    create_terminal_box(slide, code_x, by, w_right, bh, content.get('code_title', 'code'), content.get('code', ''))


# 25. Pyramid Hierarchy (피라미드 계층 구조)
def render_pyramid_hierarchy(slide, data):
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    levels = content.get('levels', [])
    if not levels: return
    n = len(levels); gap = Inches(0.08); level_h = (bh - gap * (n - 1)) / n; center_x = bx + bw / 2
    min_w = bw * 0.3; max_w = bw * 0.95

    for i, level in enumerate(levels):
        ratio = i / max(n - 1, 1); level_w = min_w + (max_w - min_w) * ratio
        level_x = center_x - level_w / 2; level_y = by + i * (level_h + gap)
        color_key = level.get('color', 'primary')
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(color_key, _SEM_BOX_STYLES['primary'])

        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, int(level_x), int(level_y), int(level_w), int(level_h))
        shp.fill.solid(); shp.fill.fore_color.rgb = fill_c; shp.line.color.rgb = line_c; shp.line.width = Pt(2.0)
        tf = shp.text_frame; tf.clear(); tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; tf.margin_left = Inches(0.2); tf.margin_right = Inches(0.2)
        p = tf.paragraphs[0]; p.text = level.get('label', ''); p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = line_c; p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(2)
        desc = level.get('desc', '')
        if desc:
            p2 = tf.add_paragraph(); p2.text = desc; p2.font.name = FONTS["BODY_TEXT"]; p2.font.size = Pt(12); p2.font.color.rgb = text_c; p2.alignment = PP_ALIGN.CENTER


# 26. Cycle Loop (순환형 프로세스)
def render_cycle_loop(slide, data):
    import math
    wrapper = data.get('data', {}); y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start); content = wrapper.get('data', {})
    steps = content.get('steps', [])
    if not steps: return
    n = len(steps); center_label = content.get('center_label', '')
    side = min(int(bw), int(bh)); cx = int(bx) + int(bw) // 2; cy = int(by) + int(bh) // 2
    radius = side // 2 - Inches(0.8)

    if center_label:
        center_size = Inches(1.6)
        center_shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx - int(center_size) // 2, cy - int(center_size) // 2, int(center_size), int(center_size))
        center_shape.fill.solid(); center_shape.fill.fore_color.rgb = COLORS["PRIMARY"]; center_shape.line.color.rgb = COLORS["PRIMARY"]
        tf = center_shape.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.text = center_label; p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment = PP_ALIGN.CENTER

    step_colors = [
        (COLORS["PRIMARY"], COLORS["SEM_BLUE_BG"]), (COLORS["SEM_GREEN"], COLORS["SEM_GREEN_BG"]),
        (COLORS["SEM_ORANGE"], COLORS["SEM_ORANGE_BG"]), (COLORS["SEM_RED"], COLORS["SEM_RED_BG"]),
        (RGBColor(30, 58, 138), RGBColor(239, 246, 255)), (RGBColor(4, 120, 87), RGBColor(236, 253, 245)),
        (RGBColor(194, 65, 12), RGBColor(255, 247, 237)), (RGBColor(185, 28, 28), RGBColor(254, 242, 242)),
    ]

    box_w = Inches(1.8); box_h = Inches(1.2)
    for i, step in enumerate(steps):
        angle = -math.pi / 2 + (2 * math.pi * i / n)
        sx = cx + int(radius * math.cos(angle)) - int(box_w) // 2
        sy = cy + int(radius * math.sin(angle)) - int(box_h) // 2
        accent, bg = step_colors[i % len(step_colors)]

        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, sy, int(box_w), int(box_h))
        shp.fill.solid(); shp.fill.fore_color.rgb = bg; shp.line.color.rgb = accent; shp.line.width = Pt(2.0)
        tf = shp.text_frame; tf.clear(); tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)
        p = tf.paragraphs[0]; p.text = step.get('label', ''); p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True; p.font.size = Pt(13); p.font.color.rgb = accent; p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(2)
        desc = step.get('desc', '')
        if desc:
            p2 = tf.add_paragraph(); p2.text = desc; p2.font.name = FONTS["BODY_TEXT"]; p2.font.size = Pt(9); p2.font.color.rgb = COLORS["GRAY"]; p2.alignment = PP_ALIGN.CENTER

        # 화살표 (현재 → 다음 단계 방향)
        next_i = (i + 1) % n; next_angle = -math.pi / 2 + (2 * math.pi * next_i / n)
        mid_angle = (angle + next_angle) / 2
        if next_i == 0 and i == n - 1: mid_angle = angle + (2 * math.pi / n) / 2
        arrow_r = int(radius * 0.65)
        arrow_x = cx + int(arrow_r * math.cos(mid_angle)) - Inches(0.15)
        arrow_y = cy + int(arrow_r * math.sin(mid_angle)) - Inches(0.15)
        arrow_size = Inches(0.3)
        arrow = slide.shapes.add_shape(MSO_SHAPE.OVAL, arrow_x, arrow_y, int(arrow_size), int(arrow_size))
        arrow.fill.solid(); arrow.fill.fore_color.rgb = COLORS["PRIMARY"]; arrow.line.color.rgb = COLORS["PRIMARY"]
        tf = arrow.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.text = "\u27A4"; p.font.size = Pt(10); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment = PP_ALIGN.CENTER


# ==========================================
# 5. 메인 라우터
# ==========================================
def render_slide_content(slide, layout, data):
    clean_body_placeholders(slide)

    renderers = {
        "bento_grid": render_bento_grid, "3_cards": render_3_cards,
        "grid_2x2": render_grid_2x2, "quad_matrix": render_quad_matrix,
        "timeline_steps": render_timeline_steps, "process_arrow": render_process_arrow, "phased_columns": render_phased_columns,
        "architecture_wide": render_architecture_wide, "image_left": render_image_left,
        "comparison_vs": render_comparison_vs, "key_metric": render_key_metric,
        "challenge_solution": render_challenge_solution, "detail_image": render_detail_image,
        "comparison_table": render_comparison_table,
        "detail_sections": render_detail_sections, "table_callout": render_table_callout,
        "full_image": render_full_image, "before_after": render_before_after,
        "icon_grid": render_icon_grid, "numbered_list": render_numbered_list,
        "stats_dashboard": render_stats_dashboard,
        "quote_highlight": render_quote_highlight, "pros_cons": render_pros_cons,
        "do_dont": render_do_dont, "split_text_code": render_split_text_code,
        "pyramid_hierarchy": render_pyramid_hierarchy, "cycle_loop": render_cycle_loop,
    }

    func = renderers.get(layout)
    if func:
        try: func(slide, data)
        except Exception as e: create_content_box(slide, Inches(1), Inches(3), Inches(10), Inches(2), "Error", str(e))
    else:
        create_content_box(slide, Inches(1), Inches(3), Inches(10), Inches(2), "Layout Not Found", str(data))
```

---

**NOTE**: `powerpoint_content.py`에는 `create_msk_architecture_diagram_with_icons()` 및 `download_aws_icons()`, `find_aws_icon()` 함수도 포함되어 있으나, 이들은 MSK 아키텍처 다이어그램 전용으로 레이아웃 렌더링과 무관합니다. 필요시 원본 파일을 참조하세요.
