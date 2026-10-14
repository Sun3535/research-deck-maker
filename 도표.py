"""도표 그리기(화면·PNG·PDF용 SVG). PPTX에는 같은 값으로 파워포인트 기본 차트를 넣는다.

규칙(dataviz): 얇은 표시, 막대 끝 둥글게, 격자는 흐리게, 계열 색은 정해진 순서,
계열이 둘 이상이면 범례 + 값 직접 표시, 글자는 글 색(계열 색으로 쓰지 않음).
막대는 0에서 시작한다(숫자의 정직).
"""

from __future__ import annotations

import html
import math

from 배치 import 글폭


def 눈금(lo: float, hi: float, 개수: int = 5) -> list[float]:
    if hi <= lo:
        hi = lo + 1
    거친 = (hi - lo) / 개수
    자리 = 10 ** math.floor(math.log10(거친))
    간격 = min((m * 자리 for m in (1, 2, 2.5, 5, 10) if m * 자리 >= 거친), default=자리 * 10)
    시작 = math.floor(lo / 간격) * 간격
    out, v = [], 시작
    while v <= hi + 간격 * 0.001:
        out.append(round(v, 10))
        v += 간격
    if out[-1] < hi:
        out.append(round(out[-1] + 간격, 10))
    return out


def 수글(v: float) -> str:
    if v is None:
        return ""
    if abs(v) >= 1000:
        return f"{v:,.0f}"
    return f"{v:.3g}" if v != int(v) else f"{int(v)}"


def 범위(도표: dict) -> tuple[float, float]:
    값들 = [v for s in 도표["계열"] for v in s["값"] if v is not None]
    lo, hi = (min(값들), max(값들)) if 값들 else (0, 1)
    if 도표.get("축0") or 도표["모양"] == "막대":
        lo, hi = min(0, lo), max(0, hi)
    else:
        여유 = (hi - lo) * 0.08 or abs(hi) * 0.1 or 1
        lo, hi = lo - 여유, hi + 여유
    return lo, hi


def svg(도표: dict, w: float, h: float, 색: dict) -> str:
    if not 도표 or not 도표.get("계열"):
        return (f"<svg width='{w}' height='{h}'><text x='{w / 2}' y='{h / 2}' fill='{색['흐린글']}' "
                f"font-size='32' text-anchor='middle'>표를 고르면 도표가 그려져요</text></svg>")
    계열색 = 색.get("계열") or [색["강조"]]
    여럿 = len(도표["계열"]) > 1
    왼, 오른, 위, 아래 = 130, 50, (76 if 여럿 else 30), 96
    pw, ph = w - 왼 - 오른, h - 위 - 아래
    lo, hi = 범위(도표)
    yt = 눈금(lo, hi)
    lo, hi = yt[0], yt[-1]
    Y = lambda v: 위 + ph * (1 - (v - lo) / (hi - lo))  # noqa: E731
    p = [f"<svg xmlns='http://www.w3.org/2000/svg' width='{w:.0f}' height='{h:.0f}' "
         f"viewBox='0 0 {w:.0f} {h:.0f}' font-family='inherit'>"]
    for t in yt:   # 흐린 격자 + 눈금 글
        p.append(f"<line x1='{왼}' x2='{w - 오른}' y1='{Y(t):.1f}' y2='{Y(t):.1f}' "
                 f"stroke='{색['선']}' stroke-width='{2 if t == 0 else 1.5}'/>"
                 f"<text x='{왼 - 16}' y='{Y(t) + 9:.1f}' fill='{색['흐린글']}' font-size='26' "
                 f"text-anchor='end'>{수글(t)}</text>")
    n = len(도표["가로"])
    모양 = 도표["모양"]
    if 모양 == "점":
        xs = [x for x in (_수(v) for v in 도표["가로"]) if x is not None] or [0, 1]
        xt = 눈금(min(xs), max(xs))
        X = lambda v: 왼 + pw * (_수(v) - xt[0]) / (xt[-1] - xt[0])  # noqa: E731
        for t in xt:
            p.append(f"<text x='{X(t):.1f}' y='{h - 아래 + 40}' fill='{색['흐린글']}' font-size='26' "
                     f"text-anchor='middle'>{수글(t)}</text>")
    else:
        띠 = pw / max(n, 1)
        X = lambda i: 왼 + 띠 * (i + 0.5)  # noqa: E731
        for i, 이름 in enumerate(도표["가로"]):
            p.append(f"<text x='{X(i):.1f}' y='{h - 아래 + 40}' fill='{색['흐린글']}' font-size='26' "
                     f"text-anchor='middle'>{html.escape(str(이름))}</text>")
    p.append(f"<text x='{왼 + pw / 2:.1f}' y='{h - 14}' fill='{색['흐린글']}' font-size='26' "
             f"text-anchor='middle'>{html.escape(str(도표['가로이름']))}</text>")
    k = len(도표["계열"])
    for j, s in enumerate(도표["계열"]):
        c = 계열색[j % len(계열색)]
        if 모양 == "막대":
            bw = 띠 * 0.72 / k
            for i, v in enumerate(s["값"]):
                if v is None:
                    continue
                x0 = X(i) - 띠 * 0.36 + j * bw + 2
                y0, y1 = sorted((Y(v), Y(0)))
                r = min(6, bw / 3, (y1 - y0) / 2)
                bw2 = bw - 4
                p.append(f"<path fill='{c}' d='M{x0:.1f},{y1:.1f} V{y0 + r:.1f} Q{x0:.1f},{y0:.1f} "
                         f"{x0 + r:.1f},{y0:.1f} H{x0 + bw2 - r:.1f} Q{x0 + bw2:.1f},{y0:.1f} "
                         f"{x0 + bw2:.1f},{y0 + r:.1f} V{y1:.1f} Z'/>")
                if n * k <= 16:   # 막대가 적으면 값을 직접 적는다
                    p.append(f"<text x='{x0 + bw2 / 2:.1f}' y='{y0 - 10:.1f}' fill='{색['글']}' "
                             f"font-size='24' text-anchor='middle'>{수글(v)}</text>")
        else:
            점들 = [((X(i) if 모양 == "선" else X(도표["가로"][i])), Y(v))
                  for i, v in enumerate(s["값"]) if v is not None]
            if 모양 == "선" and len(점들) > 1:
                p.append(f"<polyline fill='none' stroke='{c}' stroke-width='5' stroke-linejoin='round' "
                         f"points='{' '.join(f'{x:.1f},{y:.1f}' for x, y in 점들)}'/>")
            for x, y in 점들:
                p.append(f"<circle cx='{x:.1f}' cy='{y:.1f}' r='9' fill='{c}' stroke='{색['바탕']}' "
                         f"stroke-width='3'/>")
            if 점들 and 모양 == "선":   # 끝 값만 직접 표시
                x, y = 점들[-1]
                마지막 = [v for v in s["값"] if v is not None][-1]
                p.append(f"<text x='{x + 16:.1f}' y='{y + 9:.1f}' fill='{색['글']}' font-size='26'>"
                         f"{수글(마지막)}</text>")
    if 여럿:   # 범례: 표시 + 이름(글 색)
        x = 왼
        for j, s in enumerate(도표["계열"]):
            c = 계열색[j % len(계열색)]
            p.append(f"<rect x='{x}' y='18' width='26' height='26' rx='6' fill='{c}'/>"
                     f"<text x='{x + 38}' y='40' fill='{색['글']}' font-size='28'>{html.escape(s['이름'])}</text>")
            x += 38 + 글폭(s["이름"], 28) + 48
    p.append("</svg>")
    return "".join(p)


def _수(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None
