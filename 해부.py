"""해부: 참고 PPTX 한 개 → 내 스타일(색·글자체·장식) 한 개.

참고 덱을 고치지 않고 읽기만 한다. 읽는 것:
- 색: 테마 색 + 실제 장 바탕 + 자주 쓴 도형 색
- 글자체: 테마 글자체, 없으면 글에 가장 많이 쓴 글자체
- 장식: 위쪽 가로 띠 / 제목 왼쪽 세로 띠 / 제목 아래 선, 표지의 큰 사진·색 판, 강조색 바탕 장
만든 스타일은 앱의 '내스타일/' 폴더에 JSON으로 남아 다른 덱에서도 고를 수 있다.
"""

from __future__ import annotations

import json
import re
from collections import Counter
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

import 바탕pptx
import 스타일 as 스

_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
내스타일폴더 = 스.내스타일폴더


def _테마색(prs) -> dict[str, str]:
    from lxml import etree
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    theme = etree.fromstring(prs.slide_master.part.part_related_by(RT.THEME).blob)
    cs = theme.find(f".//{_A}clrScheme")
    c = {el.tag.split("}")[1]: 바탕pptx._색값(el) for el in cs} if cs is not None else {}
    # 바탕/글 이름(bg1·tx1)도 같은 색을 가리키게
    c.update({"bg1": c.get("lt1"), "tx1": c.get("dk1"), "bg2": c.get("lt2"), "tx2": c.get("dk2")})
    return {k: v for k, v in c.items() if v}


def _채움색(el, 테마: dict) -> str | None:
    """<a:solidFill> 를 품은 요소에서 색 하나."""
    if el is None:
        return None
    sf = el.find(f".//{_A}solidFill")
    if sf is None:
        return None
    for c in sf:
        if c.tag == f"{_A}srgbClr":
            return "#" + c.get("val").upper()
        if c.tag == f"{_A}schemeClr":
            return 테마.get(c.get("val"))
    return None


def _바탕색(part_el, 테마: dict) -> str | None:
    bg = part_el.find(f".//{_P}bg")
    if bg is None:
        return None
    bgPr = bg.find(f"{_P}bgPr")
    if bgPr is not None:
        return _채움색(bgPr, 테마)
    ref = bg.find(f"{_P}bgRef")       # 테마 배경 목록을 가리키면 색만 읽는다
    if ref is not None:
        for c in ref:
            if c.tag == f"{_A}schemeClr":
                return 테마.get(c.get("val"))
            if c.tag == f"{_A}srgbClr":
                return "#" + c.get("val").upper()
    return None


def _도형들(prs):
    """(어디, 도형) — 마스터·레이아웃·장의 도형 전부."""
    for s in prs.slide_master.shapes:
        yield "마스터", s
    for lay in prs.slide_layouts:
        for s in lay.shapes:
            yield "레이아웃", s
    for i, sl in enumerate(prs.slides):
        for s in sl.shapes:
            yield i, s


def _대비(a: str, b: str) -> float:
    la, lb = sorted((스._밝기(a), 스._밝기(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def _거리(a: str, b: str) -> float:
    ra = [int(a[i:i + 2], 16) for i in (1, 3, 5)]
    rb = [int(b[i:i + 2], 16) for i in (1, 3, 5)]
    return sum((x - y) ** 2 for x, y in zip(ra, rb)) ** 0.5


def 해부(pptx: Path, 이름: str | None = None, 원래이름: str | None = None) -> dict:
    """참고 PPTX → 스타일 사전 {"이름","설명","색","계열","글자체","장식","읽은것"}."""
    prs = Presentation(str(pptx))
    W, H = prs.slide_width, prs.slide_height
    테마 = _테마색(prs)
    기본 = 바탕pptx.테마읽기(pptx)
    읽은것: list[str] = []

    # 바탕: 장마다 실제로 보이는 바탕(장 → 레이아웃 → 마스터 차례)
    마스터바탕 = _바탕색(prs.slide_master._element, 테마)
    바탕들 = Counter()
    for sl in prs.slides:
        c = (_바탕색(sl._element, 테마) or _바탕색(sl.slide_layout._element, 테마) or 마스터바탕
             or 테마.get("lt1"))
        if c:
            바탕들[c] += 1
    바탕 = 바탕들.most_common(1)[0][0] if 바탕들 else (마스터바탕 or 기본["색"]["바탕"])
    어두움 = 스._밝기(바탕) < 0.2
    글 = (테마.get("lt1") if 어두움 else 테마.get("dk1")) or ("#F4F4F0" if 어두움 else "#111111")
    if _대비(글, 바탕) < 4.5:
        글 = "#F4F4F0" if 어두움 else "#111111"
    읽은것.append(f"바탕 {바탕} ({'어두운' if 어두움 else '밝은'} 바탕), 글 {글}")

    # 강조: 글·바탕이 아닌 도형 채움 중 가장 넓게 쓴 색, 없으면 테마 accent1
    넓이 = Counter()
    도형목록 = list(_도형들(prs))
    for _, s in 도형목록:
        if s.shape_type in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.FREEFORM):
            c = _채움색(s._element.find(f".//{_P}spPr"), 테마)
            if c and _거리(c, 바탕) > 60 and _거리(c, 글) > 60 and (s.width or 0) and (s.height or 0):
                넓이[c] += int(s.width) * int(s.height)
    강조 = 넓이.most_common(1)[0][0] if 넓이 else (테마.get("accent1") or 기본["색"]["강조"])
    if _대비(강조, 바탕) < 2:      # 바탕에 묻히는 강조는 쓸 수 없다
        강조 = 테마.get("accent1") or 강조
    읽은것.append(f"강조 {강조}" + (" (도형에 가장 넓게 쓴 색)" if 넓이 else " (테마 강조1)"))

    # 계열(도표): 테마 강조 1~4가 서로 다르고 바탕에서 보이면 쓰고, 아니면 검사 통과한 기본 계열
    계열 = [테마.get(f"accent{i}") for i in range(1, 5)]
    좋음 = all(계열) and all(_대비(c, 바탕) >= 1.6 for c in 계열) and all(
        _거리(a, b) > 90 for i, a in enumerate(계열) for b in 계열[i + 1:])
    if not 좋음:
        계열 = list(스._어두운계열 if 어두움 else 스._밝은계열)
        읽은것.append("도표 색: 테마 색이 서로 비슷하거나 바탕에 묻혀서 기본 도표 색을 씀")
    else:
        읽은것.append("도표 색: 테마 강조 1~4")

    # 글자체
    # 글자체: 장 글에 직접 지정해 가장 많이 쓴 것 → 없으면 테마 글자체
    쓰임 = Counter()
    for _, s in 도형목록:
        if s.has_text_frame:
            for p in s.text_frame.paragraphs:
                for r in p.runs:
                    rPr = r._r.find(f"{_A}rPr")
                    for t in ("ea", "latin"):
                        el = rPr.find(f"{_A}{t}") if rPr is not None else None
                        if el is not None and el.get("typeface") and not el.get("typeface").startswith("+"):
                            쓰임[el.get("typeface")] += len(r.text)
    글자체 = 쓰임.most_common(1)[0][0] if 쓰임 else None
    테마글자체 = 기본.get("글자체")
    글자체 = 글자체 or (테마글자체 if 테마글자체 and not 테마글자체.startswith("+") else 스.기본글자체)
    읽은것.append(f"글자체 {글자체}")

    # 장식
    장식 = {"제목": "밑줄", "표지": "그림오른쪽", "한마디": "보통"}
    for 어디, s in 도형목록:
        if s.is_placeholder or s.shape_type not in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.LINE):
            continue
        x, y, w, h = (int(v or 0) for v in (s.left, s.top, s.width, s.height))
        if w >= 0.9 * W and h <= 0.1 * H and y <= 0.05 * H:
            장식["제목"] = "위띠"
            break
        if w <= 0.02 * W and 0.03 * H <= h <= 0.2 * H and y <= 0.3 * H:
            장식["제목"] = "왼띠"
        elif w >= 0.6 * W and h <= 0.006 * H and 0.1 * H <= y <= 0.3 * H and 장식["제목"] == "밑줄":
            장식["제목"] = "선"
    if len(prs.slides):
        첫장 = prs.slides[0]
        for s in 첫장.shapes:
            x, y, w, h = (int(v or 0) for v in (s.left, s.top, s.width, s.height))
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE and w * h >= 0.6 * W * H:
                장식["표지"] = "그림배경"
            elif (s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and h >= 0.9 * H and 0.35 * W <= w <= 0.8 * W
                  and 장식["표지"] != "그림배경"):
                장식["표지"] = "띠"
    for sl in prs.slides:
        c = _바탕색(sl._element, 테마)
        if c and _거리(c, 강조) < 40:
            장식["한마디"] = "반전"
    읽은것.append(f"장식: 제목 {장식['제목']}, 표지 {장식['표지']}, 한마디 {장식['한마디']}")

    색 = {"바탕": 바탕, "면": 바탕pptx._섞기(바탕, 글, 0.07), "글": 글,
          "흐린글": 바탕pptx._섞기(글, 바탕, 0.32), "강조": 강조, "선": 바탕pptx._섞기(바탕, 글, 0.18)}
    이름 = re.sub(r"[_\- ]?v\d+$", "", 이름 or pptx.stem, flags=re.I) or "내 스타일"
    return {"이름": 이름, "설명": f"{원래이름 or pptx.name}에서 해부한 스타일", "색": 색, "계열": 계열,
            "글자체": 글자체, "장식": 장식, "읽은것": 읽은것}


# ---------------------------------------------------------------- 내 스타일 보관
def _파일이름(이름: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", 이름).strip() or "내 스타일"


def 남기기(스타일: dict) -> dict:
    """내스타일/<이름>.json 에 남긴다. 같은 이름이 있으면 (2), (3)… 을 붙인다(덮어쓰지 않음)."""
    내스타일폴더.mkdir(exist_ok=True)
    밑 = _파일이름(스타일["이름"])
    이름, n = 밑, 2
    while (내스타일폴더 / f"{이름}.json").exists() or 이름 in 스.스타일들 or 이름 == "v1 템플릿":
        이름, n = f"{밑} ({n})", n + 1
    스타일 = {**스타일, "이름": 이름}
    (내스타일폴더 / f"{이름}.json").write_text(json.dumps(스타일, ensure_ascii=False, indent=1), "utf-8")
    return 스타일
