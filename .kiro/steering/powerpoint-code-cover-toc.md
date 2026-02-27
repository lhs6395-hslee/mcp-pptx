# PowerPoint Generation - Cover & TOC Source Code

**Part of**: [powerpoint-guide.md](./powerpoint-guide.md) 시스템 명세

---

## powerpoint_cover.py - Complete Source Code

```python
# -*- coding: utf-8 -*-
from datetime import datetime
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR # 테마 색상 인식용

def replace_text_preserving_style(shape, new_text):
    """
    [헬퍼] 텍스트 교체 (첫 문단 스타일 유지, 나머지 삭제)
    """
    if not shape.has_text_frame: return
    tf = shape.text_frame
    tf.word_wrap = True # 줄바꿈 허용

    if not tf.paragraphs:
        tf.text = new_text; return

    p0 = tf.paragraphs[0]
    if not p0.runs:
        p0.text = new_text
    else:
        p0.runs[0].text = new_text
        for i in range(len(p0.runs) - 1, 0, -1):
            p0._p.remove(p0.runs[i]._r)

    # 첫 문단 이후의 나머지 문단(잔존 텍스트) 모두 삭제
    for i in range(len(tf.paragraphs) - 1, 0, -1):
        p_element = tf.paragraphs[i]._p
        p_element.getparent().remove(p_element)

def find_shapes_by_keywords(shapes, keywords):
    """키워드 포함 도형 검색 (재귀)"""
    found = []
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            found.extend(find_shapes_by_keywords(shape.shapes, keywords))
            continue
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            for kw in keywords:
                if kw in text:
                    found.append(shape)
                    break
    return found

def center_shape_horizontally(shape, slide_width_inch=13.333, shape_width_inch=None):
    """
    [NEW] 도형을 지정된 너비로 설정하고, 슬라이드 정중앙에 배치하는 함수
    """
    if shape_width_inch:
        shape.width = Inches(shape_width_inch)

    # 중앙 좌표 계산: (슬라이드너비 - 도형너비) / 2
    slide_width = Inches(slide_width_inch)
    shape.left = int((slide_width - shape.width) / 2)

    # 텍스트 내부 중앙 정렬
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER

# [핵심 1] 템플릿의 색상 정보(RGB 또는 테마 색상) 추출
def get_original_style(shape):
    style = {
        'name': None, 'size': None, 'bold': None,
        'italic': None, 'color_rgb': None, 'color_theme': None, 'brightness': 0
    }

    if shape.has_text_frame and shape.text_frame.paragraphs:
        try:
            p = shape.text_frame.paragraphs[0]
            if p.runs:
                r = p.runs[0]
                style['name'] = r.font.name
                style['size'] = r.font.size
                style['bold'] = r.font.bold
                style['italic'] = r.font.italic

                if r.font.color.type == MSO_COLOR_TYPE.RGB:
                    style['color_rgb'] = r.font.color.rgb
                elif r.font.color.type == MSO_COLOR_TYPE.THEME:
                    style['color_theme'] = r.font.color.theme_color
                    style['brightness'] = r.font.color.brightness
        except: pass
    return style

# [핵심 2] 스타일 승계 및 텍스트 교체
def apply_text_with_style(shape, text, inherited_style, force_center=False):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    p = tf.paragraphs[0]

    if force_center:
        p.alignment = PP_ALIGN.CENTER
    else:
        p.alignment = inherited_style.get('alignment', PP_ALIGN.CENTER)

    lines = text.split('\\n') if '\\n' in text else text.split('\n')

    for i, line in enumerate(lines):
        run = p.add_run()
        run.text = line

        # 스타일 복원
        if inherited_style['name']: run.font.name = inherited_style['name']
        if inherited_style['size']: run.font.size = inherited_style['size']
        if inherited_style['bold'] is not None: run.font.bold = inherited_style['bold']

        # 색상 복원
        if inherited_style['color_rgb']:
            run.font.color.rgb = inherited_style['color_rgb']
        elif inherited_style['color_theme']:
            run.font.color.theme_color = inherited_style['color_theme']
            if inherited_style['brightness']:
                run.font.color.brightness = inherited_style['brightness']
        else:
            run.font.color.rgb = RGBColor(255, 255, 255)

        if i < len(lines) - 1:
            run.text += '\n'

def center_shape_horizontally(shape, slide_width_inch=13.333, fixed_width_inch=None):
    if fixed_width_inch:
        shape.width = Inches(fixed_width_inch)
    shape.left = int((Inches(slide_width_inch) - shape.width) / 2)

def update_cover_slide(slide, title_text, subtitle_text):
    now = datetime.now()
    current_year = str(now.year)
    current_md = now.strftime("%m/%d")

    title_candidates = []
    subtitle_candidates = []

    # 1. 도형 분류
    for shape in list(slide.shapes):
        if not shape.has_text_frame: continue
        txt = shape.text_frame.text

        # (A) 날짜
        if any(k in txt for k in ["2025", "2026", "02/06", "02.06", "00/00"]):
            style = get_original_style(shape)
            new_text = current_year if ("20" in txt) else current_md
            apply_text_with_style(shape, new_text, style, force_center=False)
            continue

        # (B) 부제목
        if any(k in txt for k in ["설계", "원칙", "부제목", "Subtitle", "소제목"]):
            subtitle_candidates.append(shape)
            continue

        # (C) 제목
        if any(k in txt for k in ["가이드라인", "GS", "Template", "제목", "AWS"]):
            title_candidates.append(shape)

    # 2. 중복 제거 및 스타일 적용
    target_title = None
    if title_candidates:
        target_title = title_candidates[0]
        saved_style = get_original_style(target_title)
        for trash in title_candidates[1:]:
            try: trash._element.getparent().remove(trash._element)
            except: pass
        apply_text_with_style(target_title, title_text, saved_style, force_center=True)
        center_shape_horizontally(target_title, fixed_width_inch=8.0)

    target_subtitle = None
    if subtitle_candidates:
        target_subtitle = subtitle_candidates[0]
        saved_style = get_original_style(target_subtitle)
        for trash in subtitle_candidates[1:]:
            try: trash._element.getparent().remove(trash._element)
            except: pass
        apply_text_with_style(target_subtitle, subtitle_text, saved_style, force_center=True)
        center_shape_horizontally(target_subtitle, fixed_width_inch=11.333)

    # 3. [FIXED] 수직 정렬 계산 (int 변환 추가)
    slide_height = Inches(7.5)
    gap = Inches(0.3)

    if target_title and target_subtitle:
        total_block_height = target_title.height + gap + target_subtitle.height

        # [수정됨] 나눗셈 결과를 int()로 감싸서 정수로 변환
        start_top = int((slide_height - total_block_height) / 2)

        target_title.top = start_top
        target_subtitle.top = target_title.top + target_title.height + gap

        print(f"✅ 표지 완료: 수직 중앙 정렬 적용 (Top: {target_title.top/914400:.2f}in)")

    elif target_title:
        # 제목만 있을 때도 int() 변환 필요
        target_title.top = int((slide_height - target_title.height) / 2)
```

---

## powerpoint_toc.py - Complete Source Code

```python
# -*- coding: utf-8 -*-
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE_TYPE

def iter_shapes(shapes):
    """그룹 내부까지 재귀 탐색하여 모든 도형을 반환"""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)
        else:
            yield shape

def update_paragraph_text_only(paragraph, new_text):
    """
    [핵심] 문단 객체를 삭제하거나 새로 만들지 않고,
    기존 문단 안에 있는 텍스트(Run)만 교체하여 '줄 간격'과 '스타일'을 보존합니다.
    """
    # 1. 런(Run)이 없으면 최소한 하나 생성
    if not paragraph.runs:
        paragraph.add_run()

    # 2. 첫 번째 런에 텍스트 덮어쓰기
    paragraph.runs[0].text = new_text

    # 3. 뒤따르는 런(잔여 텍스트)만 제거 (문단 자체는 건드리지 않음)
    for i in range(len(paragraph.runs) - 1, 0, -1):
        paragraph._p.remove(paragraph.runs[i]._r)

    # 4. 텍스트가 비어있으면 불렛(점) 제거 처리
    if new_text == "":
        pPr = paragraph._p.get_or_add_pPr()
        buNone = pPr.find(qn('a:buNone'))
        if buNone is None:
            buNone = pPr.makeelement(qn('a:buNone'))
            pPr.insert(0, buNone)

def update_toc_slide(slide, toc_items):
    HEADER_LIMIT = Inches(1.8)

    # 1. 텍스트 박스 수집
    candidates = []
    for s in iter_shapes(slide.shapes):
        if not s.has_text_frame: continue
        if s.top < HEADER_LIMIT: continue # 헤더 보호

        # 잡음 제거
        txt = s.text_frame.text.strip()
        if any(x in txt for x in ["GS Neotek", "PAGE", "00/00"]): continue

        candidates.append(s)

    if not candidates:
        print("   ⚠️ 목차 영역을 찾지 못했습니다.")
        return

    # 2. [패턴 인식] "줄이 많은 상자(3줄 이상)" 우선 탐색
    # 템플릿의 '숫자통', '제목통'을 찾습니다.
    multiline_boxes = [s for s in candidates if len(s.text_frame.paragraphs) >= 3]
    multiline_boxes.sort(key=lambda s: s.left) # 좌->우 정렬 (왼쪽:숫자, 오른쪽:제목)

    # -------------------------------------------------------
    # [CASE A] 다중 문단 모드 (기존 줄 간격 유지 필수)
    # -------------------------------------------------------
    if len(multiline_boxes) > 0:
        print(f"   🚀 [다중 문단 모드] 기존 문단 객체를 재활용하여 줄 간격을 유지합니다.")

        # 숫자 박스와 제목 박스가 분리된 구조 (가장 일반적)
        if len(multiline_boxes) >= 2:
            num_box = multiline_boxes[0]
            title_box = multiline_boxes[1]

            # [중요] 기존 문단 개수만큼 루프를 돌거나, 데이터 개수만큼 돕니다.
            # 템플릿이 5줄이고 데이터가 4개면 -> 4개 채우고 1개 비움

            # 1. 숫자통 처리
            num_paragraphs = num_box.text_frame.paragraphs
            for i in range(len(num_paragraphs)):
                if i < len(toc_items):
                    # 데이터 채우기 (1, 2, 3...)
                    update_paragraph_text_only(num_paragraphs[i], str(i + 1))
                else:
                    # 남는 줄 비우기 (지우지 않고 빈칸으로 둠 -> 줄 간격 유지)
                    update_paragraph_text_only(num_paragraphs[i], "")

            # 2. 제목통 처리
            title_paragraphs = title_box.text_frame.paragraphs
            for i in range(len(title_paragraphs)):
                if i < len(toc_items):
                    # 제목 채우기
                    update_paragraph_text_only(title_paragraphs[i], toc_items[i])
                else:
                    # 남는 줄 비우기
                    update_paragraph_text_only(title_paragraphs[i], "")

        # 통짜 박스 하나인 경우
        elif len(multiline_boxes) == 1:
            box = multiline_boxes[0]
            paragraphs = box.text_frame.paragraphs
            for i in range(len(paragraphs)):
                if i < len(toc_items):
                    update_paragraph_text_only(paragraphs[i], toc_items[i])
                else:
                    update_paragraph_text_only(paragraphs[i], "")

    # -------------------------------------------------------
    # [CASE B] 개별 박스 모드 (Fallback)
    # -------------------------------------------------------
    else:
        print("   🚀 [개별 박스 모드] 행 단위로 처리합니다.")
        candidates.sort(key=lambda s: s.top)
        rows = []
        if candidates:
            current = [candidates[0]]
            for i in range(1, len(candidates)):
                if abs(candidates[i].top - candidates[i-1].top) < Inches(0.2):
                    current.append(candidates[i])
                else:
                    rows.append(current)
                    current = [candidates[i]]
            rows.append(current)

        for i, row in enumerate(rows):
            row.sort(key=lambda s: s.left)
            if i < len(toc_items):
                if len(row) >= 2:
                    update_paragraph_text_only(row[0].text_frame.paragraphs[0], str(i+1))
                    update_paragraph_text_only(row[1].text_frame.paragraphs[0], toc_items[i])
                    for extra in row[2:]: extra.text_frame.clear()
                elif len(row) == 1:
                    update_paragraph_text_only(row[0].text_frame.paragraphs[0], toc_items[i])
            else:
                for s in row: s.text_frame.clear()

    print(f"   ✅ 목차 업데이트 완료.")
```
