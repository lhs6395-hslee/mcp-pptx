# PowerPoint Generation - Orchestration Source Code

**Part of**: [powerpoint-guide.md](./powerpoint-guide.md) 시스템 명세

---

## generate_ppt.sh - Complete Source Code

```bash
#!/bin/bash
# PowerPoint 생성 스크립트

STEERING_FILE="${1:-rayhli-eks_guide_2026.py}"

echo "🎓 프레젠테이션 생성 중... ($STEERING_FILE)"
echo ""

# 가상환경 활성화 (있는 경우)
if [ -d "venv" ]; then
    source venv/bin/activate
fi

if [ ! -f "$STEERING_FILE" ]; then
    echo "❌ 스티어링 파일을 찾을 수 없습니다: $STEERING_FILE"
    echo "   데이터를 정의한 Python 파일이 필요합니다."
    exit 1
fi

# 생성 스크립트 실행
if [ -f "generate.py" ]; then
    python3 generate.py "$STEERING_FILE"
else
    echo "❌ 생성 스크립트를 찾을 수 없습니다: generate.py"
    exit 1
fi

OUTPUT_NAME=$(basename "$STEERING_FILE" .py)
echo ""
echo "✅ 완료! results/rayhli-${OUTPUT_NAME}.pptx를 확인하세요."
```

---

## generate.py - Complete Source Code

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PowerPoint 생성 스크립트

스티어링 파일(rayhli-eks_guide_2026.py)에서 데이터를 읽어 PPT를 자동 생성합니다.

사용법:
    python3 generate.py                    # 기본 스티어링 파일 사용
    python3 generate.py my_presentation.py # 커스텀 스티어링 파일 사용
"""

import os
import sys
import shutil
import zipfile
import re
from pptx import Presentation
from lxml import etree

from powerpoint_cover import update_cover_slide
from powerpoint_toc import update_toc_slide
from powerpoint_content import render_slide_content, set_slide_title_area


def remove_all_sections(pptx_file):
    """PowerPoint 파일에서 모든 섹션 제거"""
    temp_dir = 'temp_rm_sec'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    try:
        with zipfile.ZipFile(pptx_file, 'r') as z:
            z.extractall(temp_dir)
        xml_path = os.path.join(temp_dir, 'ppt', 'presentation.xml')
        if os.path.exists(xml_path):
            tree = etree.parse(xml_path)
            for elem in list(tree.getroot().iter()):
                if 'sectionLst' in elem.tag:
                    elem.getparent().remove(elem)
            tree.write(xml_path, xml_declaration=True, encoding='UTF-8', standalone=True)
        with zipfile.ZipFile(pptx_file, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(temp_dir):
                for f in files:
                    z.write(os.path.join(root, f), os.path.relpath(os.path.join(root, f), temp_dir))
    except:
        pass
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


def move_slide(prs, old_index, new_index):
    """슬라이드 위치 이동"""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[old_index])
    xml_slides.insert(new_index, slides[old_index])


def load_presentation_data(steering_file):
    """스티어링 파일에서 presentation_data 로드"""
    if not os.path.exists(steering_file):
        print(f"❌ 스티어링 파일을 찾을 수 없습니다: {steering_file}")
        sys.exit(1)

    print(f"📂 스티어링 파일: {steering_file}")

    # 스티어링 파일 실행하여 presentation_data 추출
    exec_globals = {}
    with open(steering_file, 'r', encoding='utf-8') as f:
        exec(f.read(), exec_globals)

    presentation_data = exec_globals.get('presentation_data')
    if not presentation_data:
        print(f"❌ presentation_data를 찾을 수 없습니다: {steering_file}")
        sys.exit(1)

    return presentation_data


def generate_presentation(steering_file='rayhli-eks_guide_2026.py'):
    """프레젠테이션 생성"""

    # 데이터 로드
    presentation_data = load_presentation_data(steering_file)

    # 템플릿 및 출력 설정
    TEMPLATE_FILE = 'template/2025_PPT_Template_FINAL.pptx'
    # 스티어링 파일명에서 출력 파일명 자동 생성 (rayhli- prefix)
    steering_basename = os.path.splitext(os.path.basename(steering_file))[0]
    OUTPUT_FILE = f'results/rayhli-{steering_basename}.pptx'

    IDX_COVER = 0
    IDX_TOC = 1
    IDX_BODY_SRC = 7

    if not os.path.exists('results'):
        os.makedirs('results')

    if not os.path.exists(TEMPLATE_FILE):
        print(f"❌ 템플릿 파일을 찾을 수 없습니다: {TEMPLATE_FILE}")
        sys.exit(1)

    cover_title = presentation_data['cover']['title'].replace('\n', ' ')
    print(f"🎓 [{cover_title}] 프레젠테이션 생성 시작...")
    print(f"📄 출력: {OUTPUT_FILE}")
    print()

    # 템플릿 복사
    shutil.copy(TEMPLATE_FILE, OUTPUT_FILE)
    remove_all_sections(OUTPUT_FILE)
    prs = Presentation(OUTPUT_FILE)

    if len(prs.slides) <= IDX_BODY_SRC:
        print("❌ 템플릿 슬라이드 부족")
        sys.exit(1)

    keeper_ids = []

    # 표지
    cover = prs.slides[IDX_COVER]
    keeper_ids.append(cover.slide_id)
    update_cover_slide(cover, presentation_data['cover']['title'], presentation_data['cover']['subtitle'])
    print(f"   ✅ [표지] 완료")

    # 목차
    toc = prs.slides[IDX_TOC]
    keeper_ids.append(toc.slide_id)
    clean_toc = [re.sub(r'^[\d\.]+\s*', '', s['section_title']) for s in presentation_data['sections']]
    update_toc_slide(toc, clean_toc)
    print(f"   ✅ [목차] 완료")

    # 엔딩 슬라이드 보존
    ending = prs.slides[len(prs.slides)-1]
    ending_id = ending.slide_id
    keeper_ids.append(ending_id)

    # 본문 생성
    insert_idx = 2
    body_layout = prs.slides[IDX_BODY_SRC].slide_layout

    for section in presentation_data['sections']:
        for slide_info in section['slides']:
            slide = prs.slides.add_slide(body_layout)
            move_slide(prs, len(prs.slides)-1, insert_idx)

            set_slide_title_area(slide, slide_info.get('t', ''), slide_info.get('d', ''))
            render_slide_content(slide, slide_info.get('l', 'bento_grid'), slide_info)

            keeper_ids.append(slide.slide_id)
            insert_idx += 1
            print(f"   ✅ [Page {insert_idx-1}] {slide_info.get('t')}")

    # 불필요한 슬라이드 삭제
    xml_slides = prs.slides._sldIdLst
    for i in range(len(prs.slides) - 1, -1, -1):
        if prs.slides[i].slide_id not in keeper_ids:
            rId = xml_slides[i].rId
            prs.part.drop_rel(rId)
            del xml_slides[i]

    # 엔딩 슬라이드를 마지막으로 이동
    for i, s in enumerate(prs.slides):
        if s.slide_id == ending_id:
            move_slide(prs, i, len(prs.slides)-1)
            break

    # 저장
    prs.save(OUTPUT_FILE)

    print(f"\n🎓 완료: {OUTPUT_FILE}")
    print(f"📊 총 {len(prs.slides)}장 생성")

    return OUTPUT_FILE


if __name__ == "__main__":
    # 명령행 인자 처리
    if len(sys.argv) > 1:
        steering_file = sys.argv[1]
    else:
        steering_file = 'rayhli-eks_guide_2026.py'

    try:
        output_file = generate_presentation(steering_file)
        print(f"\n✨ 성공적으로 생성되었습니다!")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
```
