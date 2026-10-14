"""박스 목록 → HTML. 한 장은 1920×1080 <section>."""

from __future__ import annotations

import html

import 도표
import 스타일 as 스
from 배치 import 글머리, 글폭


def 글자체css(주소: str) -> str:
    """주소: 글자체 폴더를 가리키는 URL(서버면 /글자체, 굽기 파일이면 file:///...)."""
    return "".join(
        f"@font-face{{font-family:Pretendard;font-weight:{w};src:url('{주소}/{f}') format('opentype')}}"
        for w, f in ((400, "Pretendard-Regular.otf"), (600, "Pretendard-SemiBold.otf"),
                     (700, "Pretendard-Bold.otf")))


장css = (
    ".장{position:relative;width:1920px;height:1080px;overflow:hidden}"
    ".장 .글{position:absolute;display:flex;flex-direction:column;word-break:keep-all;"
    "overflow-wrap:anywhere}"
    ".장 .글 p{margin:0}"
    ".장 .그림{position:absolute;overflow:hidden;display:flex;align-items:center;justify-content:center}"
    ".장 .그림 img{width:100%;height:100%;display:block}"
    ".장 .네모{position:absolute}"
    ".장 .도표{position:absolute}"
    ".장 .표{position:absolute;border-collapse:collapse;table-layout:fixed}"
    ".장 .표 td,.장 .표 th{padding:0 18px;text-align:left;font-weight:400;overflow:hidden;"
    "white-space:nowrap;text-overflow:ellipsis}"
    ".장 .표 th{font-weight:700}"
)


# 파워포인트 '오른쪽 화살표' 도형과 같은 모양(몸통 높이 50%, 머리 길이 = 높이의 절반)
_화살 = "polygon(0 25%,50% 25%,50% 0,100% 50%,50% 100%,50% 75%,0 75%)"
# 꾸밈 '부드럽게' 그림자(PPTX는 같은 값의 바깥 그림자)
_그림자 = "box-shadow:0 8px 28px rgba(0,0,0,.16);"
_오른쪽 = " style='text-align:right'"     # 표의 숫자 칸


def _박스(b: dict, 색: dict, 그림주소, 글자체: str) -> str:
    자리 = (f"left:{b['x']:.1f}px;top:{b['y']:.1f}px;width:{b['w']:.1f}px;"
          f"height:{b['h']:.1f}px;")
    data = f" data-자리='{html.escape(b['자리'])}'" if b["자리"] else ""
    if b["종류"] == "네모":
        모양 = {"원": "border-radius:50%;", "화살": f"clip-path:{_화살};"}.get(
            b.get("모양", ""), f"border-radius:{b['둥글기']}px;")
        투명 = f"opacity:{1 - b['투명']:.2f};" if b.get("투명") else ""
        return (f"<div class='네모' style='{자리}background:{색[b['채움'] or '강조']};"
                f"{모양}{투명}{_그림자 if b.get('그림자') else ''}'></div>")
    if b["종류"] == "그림":
        바탕 = f"background:{색[b['채움']]};" if b["채움"] else ""
        fit = "contain" if b["맞춤"] == "담기" else "cover"
        src = html.escape(그림주소(b["그림"]))
        필터 = {"흑백": "filter:grayscale(1);", "옅게": "opacity:.7;"}.get(b.get("필터", ""), "")
        # 테두리는 그림 위에 얹는 틀(PPTX도 같은 자리에 선만 있는 도형을 얹는다)
        테두리 = (f"<div style='position:absolute;inset:0;border:{b['테두리폭']}px solid {색[b['테두리']]};"
               f"border-radius:{b['둥글기']}px'></div>" if b.get("테두리") else "")
        return (f"<div class='그림'{data} style='{자리}{바탕}border-radius:{b['둥글기']}px;"
                f"{_그림자 if b.get('그림자') else ''}'>"
                f"<img src='{src}' style='object-fit:{fit};{필터}' alt=''>{테두리}</div>")
    if b["종류"] == "도표":
        return (f"<div class='도표'{data} style='{자리}'>"
                f"{도표.svg(b['도표'], b['w'], b['h'], 색)}</div>")
    if b["종류"] == "표":
        줄들 = b["표"]
        줄높이 = b["h"] / max(len(줄들), 1)
        정렬들 = b.get("정렬들") or []
        칸 = []
        if b.get("칸너비들"):          # 칸 너비는 배치가 글 길이로 정한 값(PPTX와 같게)
            칸.append("<colgroup>" + "".join(f"<col style='width:{w:.1f}px'>" for w in b["칸너비들"])
                     + "</colgroup>")
        for i, 줄 in enumerate(줄들):
            태그 = "th" if i == 0 else "td"
            바탕 = (f"background:{색['강조']};color:{색['바탕']}" if i == 0 else
                  f"background:{색['면'] if i % 2 == 0 else 색['바탕']};color:{색['글']}")
            칸.append(f"<tr style='height:{줄높이:.1f}px;{바탕}'>" + "".join(
                f"<{태그}{_오른쪽 if j < len(정렬들) and 정렬들[j] == '오른' else ''}>"
                f"{html.escape(str(v))}</{태그}>" for j, v in enumerate(줄)) + "</tr>")
        return (f"<table class='표'{data} style='{자리}font-size:{b['크기']:.1f}px;"
                f"border-bottom:2px solid {색['선']}'>{''.join(칸)}</table>")
    # 글
    s = b["크기"]
    세로 = {"위": "flex-start", "가운데": "center", "아래": "flex-end"}[b["세로"]]
    정렬 = {"왼": "left", "가운데": "center"}[b["정렬"]]
    들여 = 글폭(글머리, s, b["굵기"], 글자체) if b["점"] else 0
    문단 = []
    for i, p in enumerate(b["글"]):
        아래 = f"margin-bottom:{0.45 * s:.1f}px;" if i < len(b["글"]) - 1 else ""
        머리 = (f"<span style='color:{색['강조']}'>{글머리}</span>" if b["점"] else "")
        문단.append(f"<p style='{아래}padding-left:{들여:.1f}px;text-indent:-{들여:.1f}px'>"
                  f"{머리}{html.escape(p)}</p>")
    테두리 = ";outline:4px dashed #e5484d" if b.get("넘침") else ""
    return (f"<div class='글'{data} style='{자리}justify-content:{세로};text-align:{정렬};"
            f"font-size:{s:.1f}px;font-weight:{b['굵기']};line-height:{스.줄간격};"
            f"color:{색[b['색']]}{테두리}'>{''.join(문단)}</div>")


def 장html(박스들: list[dict], 색: dict, 그림주소=lambda p: p, 표시넘침: bool = True,
          글자체: str | None = None) -> str:
    이름, 꼴 = 스.글자체(글자체)
    속 = "".join(_박스({**b, "넘침": b.get("넘침") and 표시넘침}, 색, 그림주소, 이름)
                for b in 박스들)
    글꼴 = html.escape(꼴["css"], quote=True)   # 글꼴 이름의 따옴표가 속성을 끊지 않게
    return (f"<section class='장' style='background:{색['바탕']};font-family:{글꼴}'>"
            f"{속}</section>")


def 문서(장들: list[str], 글자체주소: str, 제목: str = "덱") -> str:
    """굽기용 전체 문서: 장마다 한 쪽(PDF)으로 나뉜다."""
    return (
        "<!doctype html><html lang='ko'><head><meta charset='utf-8'>"
        f"<title>{html.escape(제목)}</title><style>{글자체css(글자체주소)}{장css}"
        "@page{size:1920px 1080px;margin:0}html,body{margin:0;padding:0;background:#000}"
        ".장{break-after:page}</style></head><body>" + "".join(장들) + "</body></html>")
