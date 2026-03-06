---
inclusion: manual
---

# PowerPoint Generation - Content Renderers Source Code

**Part of**: [powerpoint-guide.md](./powerpoint-guide.md) 시스템 명세
**File**: `powerpoint_content.py` — 38 layout renderers (35 unique + 3 aliases)

---

## powerpoint_content.py - Complete Source Code

```python
# -*- coding: utf-8 -*-
import random
import urllib.request
import os
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

# [A] 폰트 & 색상 (가시성 최우선)
FONTS = {
    "HEAD_TITLE": "프리젠테이션 7 Bold",
    "HEAD_DESC":  "프리젠테이션 5 Medium",
    "BODY_TITLE": "Freesentation",
    "BODY_TEXT":  "Freesentation"
}

COLORS = {
    "PRIMARY":    RGBColor(0, 67, 218),    # 파랑 (제목)
    "BLACK":      RGBColor(0, 0, 0),       # 검정 (본문 가시성 강제)
    "DARK_GRAY":  RGBColor(33, 33, 33),    # 진한 회색
    "GRAY":       RGBColor(80, 80, 80),    # 설명글
    "BG_BOX":     RGBColor(248, 249, 250), # 박스 배경
    "BG_WHITE":   RGBColor(255, 255, 255), # 흰색 배경
    "BORDER":     RGBColor(220, 220, 220), # 테두리
    "TERMINAL_BG": RGBColor(48, 10, 36),   # 터미널 배경 (Ubuntu 보라색)
    "TERMINAL_TITLEBAR": RGBColor(44, 44, 44), # 터미널 타이틀 바 (어두운 회색)
    "TERMINAL_TEXT": RGBColor(102, 204, 102),  # 터미널 텍스트 (초록색)
    "TERMINAL_COMMENT": RGBColor(150, 150, 150), # 터미널 주석 (회색)
    "TERMINAL_RED": RGBColor(255, 95, 86),    # macOS 빨강 버튼
    "TERMINAL_YELLOW": RGBColor(255, 189, 46), # macOS 노랑 버튼
    "TERMINAL_GREEN": RGBColor(39, 201, 63),  # macOS 초록 버튼
    # 의미 기반 색상 (Semantic Colors - KMS PPT 참조)
    "SEM_RED":       RGBColor(185, 28, 28),   # 주의/필수 (제목)
    "SEM_RED_BG":    RGBColor(254, 242, 242), # 주의/필수 (배경)
    "SEM_RED_TEXT":  RGBColor(127, 29, 29),   # 주의/필수 (본문)
    "SEM_ORANGE":    RGBColor(194, 65, 12),   # 경고/핵심 (제목)
    "SEM_ORANGE_BG": RGBColor(255, 247, 237), # 경고/핵심 (배경)
    "SEM_ORANGE_TEXT": RGBColor(154, 52, 18), # 경고/핵심 (본문)
    "SEM_GREEN":     RGBColor(4, 120, 87),    # 긍정/완료 (제목)
    "SEM_GREEN_BG":  RGBColor(236, 253, 245), # 긍정/완료 (배경)
    "SEM_GREEN_TEXT": RGBColor(6, 95, 70),    # 긍정/완료 (본문)
    "SEM_BLUE":      RGBColor(30, 58, 138),   # 참조/조건 (제목)
    "SEM_BLUE_BG":   RGBColor(239, 246, 255), # 참조/조건 (배경)
    "SEM_BLUE_TEXT": RGBColor(30, 64, 175),   # 참조/조건 (본문)
    "CALLOUT_BG":    RGBColor(30, 58, 138),   # 콜아웃 배경 (진한 파랑)
    "CALLOUT_TEXT":  RGBColor(219, 234, 254), # 콜아웃 본문 (밝은 파랑)
}

# [B] 레이아웃 고정 좌표 (템플릿 원본 준수)
LAYOUT = {
    "SLIDE_TITLE_Y": Inches(0.6),      # 헤더 (상단 고정)
    "SLIDE_DESC_Y":  Inches(0.6),      # 설명글 (상단 고정)
    "BODY_START_Y":  Inches(2.0),      # 본문 시작점
    "BODY_LIMIT_Y":  Inches(7.2),      # 본문 한계선
    "MARGIN_X":      Inches(0.5),
    "SLIDE_W":       Inches(13.333)
}

# [C] 라이브러리 로드
try:
    from duckduckgo_search import DDGS
    HAS_SEARCH_LIB = True
except ImportError:
    HAS_SEARCH_LIB = False

def get_image_from_web(query):
    """이미지 검색 (배경 채우기용)"""
    if not HAS_SEARCH_LIB or not query: return None
    try:
        with DDGS() as ddgs:
            r = list(ddgs.images(f"{query} wallpaper minimal business technology", max_results=1))
            if r:
                req = urllib.request.Request(r[0]['image'], headers={'User-Agent': 'Mozilla/5.0'})
                path = f"img_{random.randint(0,9999)}.jpg"
                with urllib.request.urlopen(req, timeout=3) as res, open(path, 'wb') as f: f.write(res.read())
                return path
    except: pass
    return None

def draw_icon_search(slide, x, y, size, search_term):
    """
    아이콘 로컬 우선 로드 (v3.0 - 웹 다운로드 제거)

    전략:
    - icons/ 폴더에서 로컬 아이콘 파일 우선 검색
    - 없으면 파란색 원형으로 폴백 (웹 다운로드 없음)
    - 레이트 리미트 문제 완전 해결

    파일명 규칙:
    - "upload arrow" → "icons/upload_arrow.png"
    - "server" → "icons/server.png"
    """
    if not search_term:
        oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
        oval.fill.solid(); oval.fill.fore_color.rgb = COLORS["PRIMARY"]
        return

    # 로컬 아이콘 파일 검색
    icon_filename = search_term.replace(" ", "_") + ".png"
    icon_path = os.path.join("icons", icon_filename)

    if os.path.exists(icon_path):
        try:
            slide.shapes.add_picture(icon_path, x, y, width=size, height=size)
            print(f"   ✅ [로컬 아이콘 추가됨] '{search_term}' → {icon_path}")
            return
        except Exception as e:
            print(f"   ⚠️ [로컬 아이콘 로드 실패] '{search_term}': {str(e)[:50]}")

    # 로컬 파일 없음 → 파란색 원형 폴백
    print(f"   ⚠️ [아이콘 없음] '{search_term}' (파란색 원형 표시)")
    oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    oval.fill.solid(); oval.fill.fore_color.rgb = COLORS["PRIMARY"]

def clean_body_placeholders(slide):
    """본문 영역(2.0~7.2)만 청소"""
    for shape in list(slide.shapes):
        if shape.top > Inches(2.0) and shape.top < LAYOUT["BODY_LIMIT_Y"]:
            try: shape._element.getparent().remove(shape._element)
            except: pass

def create_content_box(slide, x, y, w, h, title, body, style="gray", search_q=None, compact=False, terminal=False):
    """
    [만능 박스 생성기]
    - 폰트: 제목 16pt / 본문 14pt (가시성 확보)
    - compact=True: grid_2x2용 작은 폰트 (제목 14pt / 본문 12pt)
    - terminal=True: 터미널 스타일 (macOS 터미널 UI)
    - 이미지: 텍스트가 적고 검색어가 있으면 배경 이미지 자동 삽입
    """
    if w < Inches(1.0): w = Inches(1.0)
    if h < Inches(1.0): h = Inches(1.0)

    # 터미널 모드: 별도 함수 호출
    if terminal:
        create_terminal_box(slide, x, y, w, h, title, body, compact=compact)
        return

    # 일반 모드
    bg = COLORS["BG_BOX"] if style=="gray" else COLORS["BG_WHITE"]
    line = COLORS["BORDER"] if style=="gray" else COLORS["PRIMARY"]

    # 이미지 자동 채우기 비활성화 (안정성 우선)
    filled_image = False
    # 웹 이미지 다운로드는 레이트 리미트 문제로 비활성화

    # 박스 생성
    if not filled_image:
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = bg
        shp.line.color.rgb = line
        shp.line.width = Pt(1.0)
        text_shape = shp
    else:
        text_shape = slide.shapes.add_textbox(x, y, w, h)

    # 텍스트 설정 (compact 모드에 따라 마진 조정)
    tf = text_shape.text_frame; tf.clear()
    if compact:
        tf.margin_left = Inches(0.3); tf.margin_right = Inches(0.8)  # 오른쪽 여백 증가 (아이콘 공간)
        tf.margin_top = Inches(0.3); tf.margin_bottom = Inches(0.3)
    else:
        tf.margin_left = Inches(0.25); tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.3); tf.margin_bottom = Inches(0.3)
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # 텍스트 오버플로우 방지 (Shrink text on overflow)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE  # 수직 중앙 정렬

    # 색상 및 폰트
    title_color = COLORS["BG_WHITE"] if filled_image else COLORS["PRIMARY"]
    body_color  = COLORS["BG_WHITE"] if filled_image else COLORS["BLACK"]
    title_font = FONTS["BODY_TITLE"]
    body_font = FONTS["BODY_TEXT"]

    # compact 모드에 따라 폰트 크기 조정
    title_size = Pt(15) if compact else Pt(16)
    body_size = Pt(13) if compact else Pt(14)
    line_spacing = Pt(6) if compact else Pt(8)

    if title:
        p = tf.paragraphs[0]; p.text = str(title)
        p.font.name = title_font; p.font.bold = True; p.font.size = title_size
        p.font.color.rgb = title_color
        p.space_after = line_spacing

    if body:
        # 줄바꿈(\n) 처리를 위해 각 줄을 별도 paragraph로 추가
        lines = str(body).split('\n')
        # 여러 줄이면 개조식(bullet list)으로 간주 → • 자동 추가
        is_list = len(lines) > 1
        for i, line in enumerate(lines):
            if i == 0 and not title:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            stripped = line.strip()
            # 개조식 자동 • 추가: 다중 줄이고, 이미 기호/번호로 시작하지 않는 경우
            # 번호 목록 패턴: "1. ", "2) ", "3: " — 실제 번호 목록만 제외
            import re as _re
            is_numbered = bool(_re.match(r'^\d+[.):\s]\s', stripped))
            if is_list and stripped and not stripped.startswith('•') and not is_numbered:
                p.text = "• " + line
            else:
                p.text = line
            p.font.name = body_font; p.font.size = body_size
            p.font.color.rgb = body_color
            p.alignment = PP_ALIGN.LEFT
            p.space_after = line_spacing

    # 아이콘 추가 (오른쪽 상단, 박스 크기 충분할 때만)
    if search_q:
        icon_size = Inches(0.6)
        # 박스가 아이콘을 수용할 수 있을 때만 표시 (최소 높이 1.2", 너비 1.5")
        if h >= Inches(1.2) and w >= Inches(1.5):
            icon_x = x + w - icon_size - Inches(0.25)
            icon_y_pos = y + Inches(0.2)
            draw_icon_search(slide, icon_x, icon_y_pos, icon_size, search_q)

def create_terminal_box(slide, x, y, w, h, title, body, compact=False):
    """Ubuntu 스타일 터미널 박스 생성 (사각형)"""
    titlebar_h = Inches(0.3)

    # 1. 전체 배경 박스 (Ubuntu 보라색, 사각형)
    background = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    background.fill.solid()
    background.fill.fore_color.rgb = COLORS["TERMINAL_BG"]
    background.line.color.rgb = COLORS["TERMINAL_BG"]

    # 2. 타이틀 바 (어두운 회색)
    titlebar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, titlebar_h)
    titlebar.fill.solid()
    titlebar.fill.fore_color.rgb = COLORS["TERMINAL_TITLEBAR"]
    titlebar.line.color.rgb = COLORS["TERMINAL_TITLEBAR"]

    # 3. macOS 버튼 3개
    btn_size = Inches(0.11)
    btn_y = y + (titlebar_h - btn_size) / 2
    btn_colors = [COLORS["TERMINAL_RED"], COLORS["TERMINAL_YELLOW"], COLORS["TERMINAL_GREEN"]]
    btn_gap = Inches(0.06)

    for i, color in enumerate(btn_colors):
        btn_x = x + Inches(0.12) + i * (btn_size + btn_gap)
        btn = slide.shapes.add_shape(MSO_SHAPE.OVAL, btn_x, btn_y, btn_size, btn_size)
        btn.fill.solid()
        btn.fill.fore_color.rgb = color
        btn.line.color.rgb = color

    # 4. "bash" 타이틀
    title_tb = slide.shapes.add_textbox(x + Inches(0.5), y, w - Inches(0.5), titlebar_h)
    tf_title = title_tb.text_frame
    tf_title.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_title.margin_left = Inches(0)
    p_title = tf_title.paragraphs[0]
    p_title.text = title if title else "bash"
    p_title.font.name = "Courier New"
    p_title.font.size = Pt(10)
    p_title.font.bold = True
    p_title.font.color.rgb = RGBColor(200, 200, 200)
    p_title.alignment = PP_ALIGN.CENTER

    # 5. 코드 텍스트 (여백 충분히 확보)
    text_y = y + titlebar_h + Inches(0.15)
    text_h = h - titlebar_h - Inches(0.3)

    text_tb = slide.shapes.add_textbox(x + Inches(0.25), text_y, w - Inches(0.5), text_h)
    tf = text_tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # 텍스트 오버플로우 방지
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE  # 수직 중앙 정렬
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)
    tf.margin_top = Inches(0.2)
    tf.margin_bottom = Inches(0.2)

    lines = str(body).split('\n')
    n_lines = len(lines)

    # 라인 수에 따라 동적 폰트 크기 조절
    # generate.py에서 split_text_code는 MAX_LINES_PER_SLIDE=14 이하로 분할 전달됨
    if compact or n_lines > 20:
        font_size = Pt(9); line_spacing = Pt(2)
    elif n_lines > 15:
        font_size = Pt(10); line_spacing = Pt(3)
    elif n_lines > 10:
        font_size = Pt(11); line_spacing = Pt(4)
    else:
        font_size = Pt(14); line_spacing = Pt(6)
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.name = FONTS["BODY_TEXT"]  # 본문과 동일한 폰트
        p.font.size = font_size

        # 주석(#으로 시작)은 회색으로 표시
        if line.strip().startswith('#'):
            p.font.color.rgb = COLORS["TERMINAL_COMMENT"]
        else:
            p.font.color.rgb = COLORS["TERMINAL_TEXT"]

        p.alignment = PP_ALIGN.LEFT
        p.space_after = line_spacing

def set_slide_title_area(slide, title_text, desc_text=""):
    """헤더 설정 (템플릿 좌표 준수)"""
    # 1. 제목 (Left)
    title_shape = slide.shapes.title or slide.shapes.add_textbox(Inches(0.5), LAYOUT["SLIDE_TITLE_Y"], Inches(4.5), Inches(1.0))
    title_shape.left, title_shape.top = Inches(0.5), LAYOUT["SLIDE_TITLE_Y"]
    title_shape.width = Inches(4.5)

    tf = title_shape.text_frame; tf.clear(); tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = str(title_text)
    # 제목 길이에 따라 폰트 자동 축소 (단어 잘림 방지)
    # "프리젠테이션 7 Bold" 28pt: Latin 글자 ~17pt 너비 → 4.5" = 324pt → 최대 ~19자
    _char_count = len(str(title_text))
    if _char_count > 23:
        _title_pt = Pt(22)
    elif _char_count > 18:
        _title_pt = Pt(25)
    else:
        _title_pt = Pt(28)
    p.font.name = FONTS["HEAD_TITLE"]; p.font.size = _title_pt; p.font.bold = True
    p.font.color.rgb = COLORS["PRIMARY"]; p.alignment = PP_ALIGN.LEFT

    # 2. 설명 (Right)
    desc_box = None
    for s in slide.shapes:
        if s.has_text_frame and s.left > Inches(5.0) and s.top < Inches(1.5):
            desc_box = s; break
    if not desc_box:
        desc_box = slide.shapes.add_textbox(Inches(5.2), LAYOUT["SLIDE_DESC_Y"], Inches(7.6), Inches(1.2))

    tf_d = desc_box.text_frame; tf_d.clear(); tf_d.word_wrap = True
    p_d = tf_d.paragraphs[0]; p_d.text = str(desc_text)
    p_d.font.name = FONTS["HEAD_DESC"]; p_d.font.size = Pt(12); p_d.font.color.rgb = COLORS["GRAY"]
    p_d.alignment = PP_ALIGN.LEFT

def draw_body_header_and_get_y(slide, title, desc):
    """본문 헤더 (동적 위치 계산)"""
    current_y = LAYOUT["BODY_START_Y"]
    content_w = LAYOUT["SLIDE_W"] - (LAYOUT["MARGIN_X"] * 2)

    if title:
        tb = slide.shapes.add_textbox(LAYOUT["MARGIN_X"], current_y, content_w, Inches(0.6))
        p = tb.text_frame.paragraphs[0]; p.text = "• " + str(title)
        p.font.name = FONTS["BODY_TITLE"]; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = COLORS["DARK_GRAY"]
        current_y += Inches(0.5)

        if desc:
            tb_d = slide.shapes.add_textbox(LAYOUT["MARGIN_X"], current_y, content_w, Inches(0.5))
            tb_d.text_frame.word_wrap = True
            p_d = tb_d.text_frame.paragraphs[0]; p_d.text = str(desc)
            p_d.font.name = FONTS["BODY_TEXT"]; p_d.font.size = Pt(12); p_d.font.color.rgb = COLORS["GRAY"]
            current_y += Inches(0.5)
        current_y += Inches(0.2)
    return current_y

def calculate_dynamic_rect(start_y):
    """남은 공간 계산"""
    available_h = LAYOUT["BODY_LIMIT_Y"] - start_y
    if available_h < Inches(1.5): available_h = Inches(1.5)
    return LAYOUT["MARGIN_X"], start_y, LAYOUT["SLIDE_W"] - (LAYOUT["MARGIN_X"] * 2), available_h

def render_3_cards(slide, data):
    """
    3개 카드 레이아웃 (동적 높이 계산)

    개선사항:
    - 모든 카드의 최대 본문 줄 수 계산
    - 각 줄당 0.28인치 (11pt 폰트 + 라인 간격 1.1 + 여백)
    - 3줄 이상의 본문 텍스트 완벽 지원
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    # 카드는 항상 가용 높이 전체를 채움 (auto_size로 오버플로우 방지)
    card_h = bh
    gap = Inches(0.3); w_card = (bw - (gap * 2)) / 3
    for i, key in enumerate(['card_1', 'card_2', 'card_3']):
        item = content.get(key, {})
        x = bx + i * (w_card + gap)
        create_content_box(slide, x, by, w_card, card_h, "", "", "white")

        # 아이콘+텍스트 그룹을 카드 내에서 수직 중앙 정렬
        icon_size = Inches(0.8)
        icon_text_gap = Inches(0.15)
        text_area_height = card_h * 0.65

        total_content_height = icon_size + icon_text_gap + text_area_height
        top_margin = (card_h - total_content_height) / 2

        # 아이콘 배치
        icon_y = by + top_margin
        draw_icon_search(slide, x + w_card/2 - icon_size/2, icon_y, icon_size, item.get('search_q'))

        # 텍스트박스 배치 (auto_size로 오버플로우 방지)
        text_y = icon_y + icon_size + icon_text_gap
        text_height = text_area_height
        tb = slide.shapes.add_textbox(x, text_y, w_card, text_height)
        tb.text_frame.word_wrap = True
        tb.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tb.text_frame.margin_left = Inches(0.2)
        tb.text_frame.margin_right = Inches(0.2)
        tb.text_frame.margin_top = Inches(0.1)
        tb.text_frame.margin_bottom = Inches(0.1)
        p = tb.text_frame.paragraphs[0]; p.text = item.get('title',''); p.font.bold=True; p.font.size=Pt(17); p.alignment=PP_ALIGN.CENTER; p.font.color.rgb=COLORS["PRIMARY"]; p.font.name=FONTS["BODY_TITLE"]
        body_lines = [l for l in item.get('body', '').split('\n') if l.strip()] or ['']
        is_list = len(body_lines) > 1
        import re as _re
        for li, body_line in enumerate(body_lines):
            p2 = tb.text_frame.add_paragraph()
            stripped = body_line.strip()
            is_numbered = bool(_re.match(r'^\d+[.):\s]\s', stripped))
            if is_list and stripped and not stripped.startswith('•') and not is_numbered:
                p2.text = "• " + body_line
            else:
                p2.text = body_line
            p2.font.size=Pt(13); p2.alignment=PP_ALIGN.CENTER; p2.font.color.rgb=COLORS["BLACK"]; p2.font.name=FONTS["BODY_TEXT"]
            if li == 0: p2.space_before = Pt(8)

# 1. Bento Grid
def render_bento_grid(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    gap = Inches(0.2); w_main = (bw - gap) / 2
    # main도 터미널이면 compact 적용
    main_compact = content.get('main',{}).get('terminal', False)
    create_content_box(slide, bx, by, w_main, bh, content.get('main',{}).get('title'), content.get('main',{}).get('body'), "gray", content.get('main',{}).get('search_q'), compact=main_compact, terminal=content.get('main',{}).get('terminal', False))
    h_sub = (bh - gap) / 2
    create_content_box(slide, bx+w_main+gap, by, w_main, h_sub, content.get('sub1',{}).get('title'), content.get('sub1',{}).get('body'), "white", content.get('sub1',{}).get('search_q'), compact=True, terminal=content.get('sub1',{}).get('terminal', False))
    create_content_box(slide, bx+w_main+gap, by+h_sub+gap, w_main, h_sub, content.get('sub2',{}).get('title'), content.get('sub2',{}).get('body'), "white", content.get('sub2',{}).get('search_q'), compact=True, terminal=content.get('sub2',{}).get('terminal', False))

# 3. Grid 2x2
def render_grid_2x2(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    gap = Inches(0.2); w_half = (bw - gap) / 2; h_half = (bh - gap) / 2
    coords = [(0,0), (1,0), (0,1), (1,1)]
    for i, key in enumerate(['item1', 'item2', 'item3', 'item4']):
        item = content.get(key, {})
        c, r = coords[i]
        create_content_box(slide, bx + c*(w_half+gap), by + r*(h_half+gap), w_half, h_half, item.get('title'), item.get('body'), "white", item.get('search_q'), compact=True, terminal=item.get('terminal', False))

# 4. Quad Matrix (Alias)
def render_quad_matrix(slide, data): render_grid_2x2(slide, data)

# 5. Challenge & Solution
def render_challenge_solution(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    # challenge와 solution은 wrapper 레벨에 있음
    content = wrapper

    # challenge/solution은 dict 또는 string일 수 있음
    ch = content.get('challenge', {})
    sol = content.get('solution', {})
    ch_title = ch.get('title', 'CHALLENGE') if isinstance(ch, dict) else 'CHALLENGE'
    ch_body = ch.get('body', '') if isinstance(ch, dict) else str(ch)
    sol_title = sol.get('title', 'SOLUTION') if isinstance(sol, dict) else 'SOLUTION'
    sol_body = sol.get('body', '') if isinstance(sol, dict) else str(sol)

    gap = Inches(0.6); w_half = (bw - gap) / 2
    create_content_box(slide, bx, by, w_half, bh, ch_title, ch_body, "gray")
    create_content_box(slide, bx+w_half+gap, by, w_half, bh, sol_title, sol_body, "white")
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, bx+w_half-Inches(0.5)+(gap/2), by+(bh/2)-Inches(0.5), Inches(1.0), Inches(1.0))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = COLORS["PRIMARY"]

# 6. Timeline Steps
def render_timeline_steps(slide, data):
    """카드 형태 타임라인 (가시성 최적화)"""
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    steps = content.get('steps', [])
    if not steps: return

    # 카드 간 간격
    arrow_gap = Inches(0.4)
    card_width = (bw - (arrow_gap * (len(steps) - 1))) / len(steps)

    for i, step in enumerate(steps):
        x = bx + i * (card_width + arrow_gap)

        # 카드 박스 (명확한 배경)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, by, card_width, bh)
        card.fill.solid()
        card.fill.fore_color.rgb = COLORS["BG_BOX"]
        card.line.color.rgb = COLORS["PRIMARY"]
        card.line.width = Pt(2.0)

        # 숫자 배지 (큰 원형)
        badge_size = Inches(0.8)
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, x + (card_width/2) - (badge_size/2), by + Inches(0.3), badge_size, badge_size)
        badge.fill.solid()
        badge.fill.fore_color.rgb = COLORS["PRIMARY"]
        badge.line.color.rgb = COLORS["PRIMARY"]

        # 배지 숫자
        tf_badge = badge.text_frame
        tf_badge.clear()
        tf_badge.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_badge = tf_badge.paragraphs[0]
        p_badge.text = str(i + 1)
        p_badge.font.name = FONTS["BODY_TITLE"]
        p_badge.font.bold = True
        p_badge.font.size = Pt(28)
        p_badge.font.color.rgb = COLORS["BG_WHITE"]
        p_badge.alignment = PP_ALIGN.CENTER

        # 텍스트 영역 (아이콘 제거 → 배지 바로 아래에서 시작)
        text_y = by + Inches(1.3)
        text_h = bh - Inches(1.3) - Inches(0.3)
        tb = slide.shapes.add_textbox(x + Inches(0.2), text_y, card_width - Inches(0.4), text_h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.15)
        tf.margin_right = Inches(0.15)

        # 날짜/기간
        p = tf.paragraphs[0]
        p.text = step.get('date','')
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(16)
        p.font.color.rgb = COLORS["PRIMARY"]
        p.font.name = FONTS["BODY_TITLE"]
        p.space_after = Pt(10)

        # 설명
        p2 = tf.add_paragraph()
        p2.text = step.get('desc','')
        p2.font.size = Pt(14)
        p2.alignment = PP_ALIGN.CENTER
        p2.font.color.rgb = COLORS["BLACK"]
        p2.font.name = FONTS["BODY_TEXT"]
        p2.line_spacing = 1.3

        # 단계 간 화살표 (마지막 단계 제외)
        if i < len(steps) - 1:
            arrow_x = x + card_width + Inches(0.05)
            arrow_y = by + (bh / 2) - Inches(0.3)
            arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, arrow_x, arrow_y, arrow_gap - Inches(0.1), Inches(0.6))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = COLORS["PRIMARY"]
            arrow.line.color.rgb = COLORS["PRIMARY"]

# 7. Process Arrow
def render_process_arrow(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    steps = content.get('steps', [])
    if not steps: return
    gap = Inches(0.3); w_step = (bw - (gap * (len(steps)-1))) / len(steps)
    for i, step in enumerate(steps):
        x = bx + i*(w_step+gap)
        shp = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, x, by, w_step, Inches(0.8))
        shp.fill.solid(); shp.fill.fore_color.rgb = COLORS["PRIMARY"]
        p = shp.text_frame.paragraphs[0]; p.text = step.get('title',''); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment=PP_ALIGN.CENTER; p.font.size=Pt(14); p.font.bold=True
        create_content_box(slide, x, by + Inches(1.0), w_step, bh - Inches(1.0), "", step.get('body',''), "white", step.get('search_q'), terminal=step.get('terminal', False))

# 7-2. Phased Columns (단계별 컬럼 + 의미 기반 색상)
def render_phased_columns(slide, data):
    """단계별 컬럼 레이아웃 (의미 기반 색상)

    N개 세로 컬럼 나란히 배치, 의미 기반 고유 색상.
    각 컬럼: 색상 헤더 스트립 + 본문 내용 + 아이콘

    data.data.data.steps: [
        {"title": "1. 현황분석", "body": "...", "search_q": "..."},
        ...
    ]
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    steps = content.get('steps', [])
    if not steps:
        return

    n = len(steps)
    gap = Inches(0.15)
    col_w = (bw - gap * (n - 1)) / n
    header_h = Inches(0.7)

    # 의미 기반 색상 팔레트 (각 단계별 고유 색상)
    _phase_colors = [
        COLORS["PRIMARY"],          # 파랑
        RGBColor(4, 120, 87),       # 초록
        RGBColor(194, 65, 12),      # 주황
        RGBColor(185, 28, 28),      # 빨강
        RGBColor(30, 58, 138),      # 진파랑
        RGBColor(120, 53, 15),      # 갈색
        RGBColor(88, 28, 135),      # 보라
    ]
    colors = [_phase_colors[i % len(_phase_colors)] for i in range(n)]

    for i, step in enumerate(steps):
        x = bx + i * (col_w + gap)

        # 헤더 스트립 (색상)
        header = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, by, col_w, header_h)
        header.fill.solid()
        header.fill.fore_color.rgb = colors[i]
        header.line.color.rgb = colors[i]
        header.line.width = Pt(0.5)

        tf = header.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.1)
        tf.margin_right = Inches(0.1)
        p = tf.paragraphs[0]
        p.text = step.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

        # 본문 박스 (헤더 아래)
        body_y = by + header_h + Inches(0.1)
        body_h = bh - header_h - Inches(0.1)

        body_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, body_y, col_w, body_h)
        body_box.fill.solid()
        body_box.fill.fore_color.rgb = COLORS["BG_BOX"]
        body_box.line.color.rgb = COLORS["BORDER"]
        body_box.line.width = Pt(1.0)

        tf_body = body_box.text_frame
        tf_body.clear()
        tf_body.word_wrap = True
        tf_body.vertical_anchor = MSO_ANCHOR.TOP
        tf_body.margin_left = Inches(0.15)
        tf_body.margin_right = Inches(0.15)
        tf_body.margin_top = Inches(0.15)
        tf_body.margin_bottom = Inches(0.15)

        body_text = step.get('body', '')
        lines = str(body_text).split('\n')
        for j, line in enumerate(lines):
            p = tf_body.paragraphs[0] if j == 0 else tf_body.add_paragraph()
            p.text = line
            p.font.name = FONTS["BODY_TEXT"]
            p.font.size = Pt(12)
            p.font.color.rgb = COLORS["BLACK"]
            p.alignment = PP_ALIGN.LEFT
            p.space_after = Pt(4)

        # 아이콘 (본문 박스 우하단)
        if step.get('search_q'):
            icon_size = Inches(0.5)
            icon_x = x + col_w - icon_size - Inches(0.1)
            icon_y = body_y + body_h - icon_size - Inches(0.1)
            draw_icon_search(slide, icon_x, icon_y, icon_size, step['search_q'])

# 8. Architecture Wide
def render_architecture_wide(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    h_diag = bh * 0.45
    # 다이어그램 영역: 로컬 이미지 또는 아이콘+화살표 폴백
    diagram_src = content.get('diagram_path', '')
    diag_loaded = False
    if diagram_src and os.path.exists(diagram_src):
        try:
            slide.shapes.add_picture(diagram_src, bx, by, width=bw, height=h_diag)
            diag_loaded = True
        except: pass
    if not diag_loaded:
        # 아이콘+화살표 폴백: 컬럼 아이콘들을 가로로 배치
        cols_data = [content.get(f'col{i+1}', {}) for i in range(3)]
        icon_keys = [c.get('search_q', '') for c in cols_data if isinstance(c, dict)]
        if icon_keys:
            icon_n = len(icon_keys); icon_size = Inches(1.0)
            icon_gap = (bw - icon_size * icon_n) / max(icon_n + 1, 1)
            for idx, sq in enumerate(icon_keys):
                ix = bx + icon_gap * (idx + 1) + icon_size * idx
                iy = by + (h_diag - icon_size) / 2
                draw_icon_search(slide, ix, iy, icon_size, sq)
                if idx < icon_n - 1:
                    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, ix + icon_size + Inches(0.1), by + h_diag / 2 - Inches(0.15), icon_gap - Inches(0.2), Inches(0.3))
                    arrow.fill.solid(); arrow.fill.fore_color.rgb = COLORS["PRIMARY"]; arrow.line.color.rgb = COLORS["PRIMARY"]
        else:
            ph = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, bx, by, bw, h_diag)
            ph.fill.solid(); ph.fill.fore_color.rgb = RGBColor(230,230,230); ph.text_frame.text = "Diagram Area"

    y_desc = by + h_diag + Inches(0.2); h_desc = bh - h_diag - Inches(0.2); gap = Inches(0.15); w_col = (bw - (gap*2)) / 3
    for i, k in enumerate(['col1', 'col2', 'col3']):
        if k in content:
            item = content[k]
            create_content_box(slide, bx + i*(w_col+gap), y_desc, w_col, h_desc, item.get('title',''), item.get('body',''), "white", item.get('search_q'), compact=True)

# 9. Image Left
def render_image_left(slide, data):
    """좌측 이미지 + 우측 텍스트 레이아웃 (개조식)"""
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)

    content = wrapper.get('data', {})
    gap = Inches(0.25)
    w_half = (bw - gap) / 2

    # 좌측: 이미지 (image_path 또는 search_q 폴더 체인)
    image_path = content.get('image_path')
    img_loaded = False
    if image_path and os.path.exists(image_path):
        try:
            slide.shapes.add_picture(image_path, bx, by, width=w_half, height=bh)
            img_loaded = True
        except Exception as e:
            print(f"⚠️ 이미지 로드 실패: {str(e)[:50]}")
    if not img_loaded:
        sq = content.get('search_q', '')
        if sq:
            for folder in ['architecture', 'screenshots', 'icons']:
                candidate = os.path.join(folder, sq.replace(' ', '_') + '.png')
                if os.path.exists(candidate):
                    try:
                        slide.shapes.add_picture(candidate, bx, by, width=w_half, height=bh)
                        img_loaded = True
                    except: pass
                    break
    if not img_loaded:
        placeholder = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, bx, by, w_half, bh)
        placeholder.fill.solid()
        placeholder.fill.fore_color.rgb = COLORS["BG_BOX"]
        placeholder.line.color.rgb = COLORS["BORDER"]

    # 우측: 텍스트 (개조식 - 불릿 포인트)
    text_x = bx + w_half + gap
    text_box = slide.shapes.add_textbox(text_x, by, w_half, bh)
    tf = text_box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.3)
    tf.margin_bottom = Inches(0.3)

    # bullets 배열 처리
    bullets = content.get('bullets', [])
    if bullets:
        for i, bullet in enumerate(bullets):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            p.text = f"• {bullet}"
            p.font.name = FONTS["BODY_TEXT"]
            p.font.size = Pt(16)
            p.font.color.rgb = COLORS["BLACK"]
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.3
            p.space_after = Pt(12)
    else:
        # 하위 호환성: body 필드 지원
        body_text = content.get('body', '')
        if body_text:
            lines = body_text.split('\n')
            for i, line in enumerate(lines):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()

                p.text = line.strip()
                p.font.name = FONTS["BODY_TEXT"]
                p.font.size = Pt(14)
                p.font.color.rgb = COLORS["BLACK"]
                p.alignment = PP_ALIGN.LEFT
                p.line_spacing = 1.2
                p.space_after = Pt(8)

# 10. Comparison VS
def render_comparison_vs(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    # VS 원형을 위한 충분한 간격 확보 (더 증가)
    gap = Inches(0.8); w_half = (bw - gap) / 2
    # 아이콘 없이 텍스트만 표시 (comparison_vs는 텍스트 비교가 목적)
    create_content_box(slide, bx, by, w_half, bh, content.get('item_a_title','A'), content.get('item_a_body',''), "gray")
    create_content_box(slide, bx + w_half + gap, by, w_half, bh, content.get('item_b_title','B'), content.get('item_b_body',''), "white")

    # VS 원형 + 텍스트
    oval_x = bx + w_half - Inches(0.5) + (gap/2)
    oval_y = by + (bh/2) - Inches(0.5)
    oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, oval_x, oval_y, Inches(1.0), Inches(1.0))
    oval.fill.solid()
    oval.fill.fore_color.rgb = COLORS["PRIMARY"]
    oval.line.color.rgb = COLORS["PRIMARY"]

    # VS 텍스트 추가
    tf = oval.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = "VS"
    p.font.name = FONTS["BODY_TITLE"]
    p.font.bold = True
    p.font.size = Pt(20)
    p.font.color.rgb = COLORS["BG_WHITE"]
    p.alignment = PP_ALIGN.CENTER

# 11. Key Metric
def render_key_metric(slide, data): render_3_cards(slide, data)

# 12. Detail Image
def render_detail_image(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    h_text = bh * 0.25
    create_content_box(slide, bx, by, bw, h_text, content.get('title',''), content.get('body',''), "gray")

    img_y = by + h_text + Inches(0.2); img_h = bh - h_text - Inches(0.2)

    # 로컬 이미지 우선 로드 (architecture 폴더 우선, 없으면 icons 폴더)
    search_q = content.get('search_q')
    img_loaded = False

    if search_q:
        # 로컬 파일 검색 (architecture 폴더 우선)
        img_filename = search_q.replace(" ", "_") + ".png"

        # 1. architecture 폴더 확인
        img_path = os.path.join("architecture", img_filename)
        if os.path.exists(img_path):
            try:
                try:
                    from PIL import Image
                    with Image.open(img_path) as img:
                        orig_width, orig_height = img.size
                    aspect_ratio = orig_width / orig_height
                    available_width = bw
                    available_height = img_h
                    if available_width / aspect_ratio <= available_height:
                        final_width = available_width
                        final_height = available_width / aspect_ratio
                    else:
                        final_height = available_height
                        final_width = available_height * aspect_ratio
                    centered_x = bx + (available_width - final_width) / 2
                    centered_y = img_y + (available_height - final_height) / 2
                    slide.shapes.add_picture(img_path, int(centered_x), int(centered_y),
                                            width=int(final_width), height=int(final_height))
                    print(f"   ✅ [아키텍처 다이어그램 추가됨 - 중앙 정렬] '{search_q}' → {img_path}")
                except ImportError:
                    slide.shapes.add_picture(img_path, bx, img_y, width=bw, height=img_h)
                    print(f"   ✅ [아키텍처 다이어그램 추가됨] '{search_q}' → {img_path}")
                img_loaded = True
            except Exception as e:
                print(f"   ⚠️ [아키텍처 다이어그램 로드 실패] '{search_q}': {str(e)[:50]}")

        # 2. icons 폴더 확인 (architecture에 없으면) — 아이콘은 적절한 크기로 + 라벨
        if not img_loaded:
            img_path = os.path.join("icons", img_filename)
            if os.path.exists(img_path):
                try:
                    icon_max = min(Inches(2.5), img_h * 0.6)
                    icon_cx = bx + (bw - icon_max) / 2
                    icon_cy = img_y + (img_h - icon_max - Inches(0.4)) / 2
                    slide.shapes.add_picture(img_path, int(icon_cx), int(icon_cy),
                                            width=int(icon_max), height=int(icon_max))
                    # 아이콘 아래 라벨 추가
                    label_y = int(icon_cy) + int(icon_max) + Inches(0.1)
                    label_tb = slide.shapes.add_textbox(bx, label_y, bw, Inches(0.35))
                    label_tf = label_tb.text_frame; label_tf.word_wrap = True
                    label_p = label_tf.paragraphs[0]
                    label_p.text = search_q.replace('_', ' ').title()
                    label_p.font.name = FONTS["BODY_TEXT"]; label_p.font.size = Pt(12)
                    label_p.font.color.rgb = COLORS["DARK_GRAY"]; label_p.alignment = PP_ALIGN.CENTER
                    print(f"   ✅ [로컬 다이어그램 추가됨 - 중앙 정렬] '{search_q}' → {img_path}")
                    img_loaded = True
                except Exception as e:
                    print(f"   ⚠️ [로컬 다이어그램 로드 실패] '{search_q}': {str(e)[:50]}")

    # 로컬 파일 없으면 폴백
    if not img_loaded:
        placeholder = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, img_y, bw, img_h)
        placeholder.fill.solid()
        placeholder.fill.fore_color.rgb = COLORS["BG_BOX"]
        placeholder.text_frame.text = f"Diagram: {search_q or 'N/A'}"

# 13. Comparison Table
def render_comparison_table(slide, data):
    """표 형태 비교 레이아웃 (3열 비교)"""
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    # 3개 열 데이터
    columns = content.get('columns', [])
    if not columns or len(columns) != 3:
        return

    gap = Inches(0.2)
    w_col = (bw - (gap * 2)) / 3

    # 헤더 행 (제목)
    header_h = Inches(0.8)
    for i, col in enumerate(columns):
        x = bx + i * (w_col + gap)
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, by, w_col, header_h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = COLORS["PRIMARY"]
        shp.line.color.rgb = COLORS["PRIMARY"]
        shp.line.width = Pt(1.0)

        tf = shp.text_frame
        tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = col.get('title', '') if isinstance(col, dict) else str(col)
        p.font.name = FONTS["BODY_TITLE"]
        p.font.bold = True
        p.font.size = Pt(16)
        p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

    # 데이터 행들
    rows = content.get('rows', [])
    row_h = (bh - header_h - Inches(0.2)) / len(rows) if rows else Inches(1.0)

    for row_idx, row in enumerate(rows):
        row_y = by + header_h + Inches(0.2) + (row_idx * row_h)
        values = row if isinstance(row, list) else row.get('values', ['', '', ''])

        for col_idx, value in enumerate(values):
            x = bx + col_idx * (w_col + gap)

            shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, row_y, w_col, row_h - Inches(0.05))
            shp.fill.solid()
            shp.fill.fore_color.rgb = COLORS["BG_WHITE"]
            shp.line.color.rgb = COLORS["BORDER"]
            shp.line.width = Pt(1.0)

            tf = shp.text_frame
            tf.clear()
            tf.margin_left = Inches(0.15)
            tf.margin_right = Inches(0.15)
            tf.margin_top = Inches(0.1)
            tf.margin_bottom = Inches(0.1)
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE

            p = tf.paragraphs[0]
            p.text = str(value)
            p.font.name = FONTS["BODY_TEXT"]
            p.font.size = Pt(14)
            p.font.color.rgb = COLORS["BLACK"]
            p.alignment = PP_ALIGN.CENTER

```

---

**(계속: 14~41번 레이아웃 + 다이어그램 헬퍼 + 라우터)**

→ [powerpoint-code-content-2.md](./powerpoint-code-content-2.md)
