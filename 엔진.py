"""엔진: 프로젝트 읽고 쓰기, 덱 배치, 굽기.

프로젝트 폴더: 프로젝트/<이름>/ (덱설계.json · 소재/ · 내보내기/ · .이전/)
지우지 않는다: 설계를 저장할 때마다 이전 판을 .이전/ 에 남기고, 지운 장은 설계의 "휴지통"에 둔다.
"""

from __future__ import annotations

import json
import os
import time
from pathlib import Path
from urllib.parse import quote

import base64

import 그리기_html as 화면
import 재료받기
import 그리기_pptx as 파포
import 감독관
import 꾸밈
import 내보내기
import 바탕pptx
import 해부
import 굽기
import 배치
import 소재
import 스타일 as 스

뿌리 = Path(__file__).parent
프로젝트들 = 뿌리 / "프로젝트"
설계파일 = "덱설계.json"


# ---------------------------------------------------------------- 읽고 쓰기
def 폴더(이름: str) -> Path:
    p = (프로젝트들 / 이름).resolve()
    if 프로젝트들.resolve() not in p.parents:
        raise ValueError("프로젝트 이름이 올바르지 않아요")
    return p


def 목록() -> list[dict]:
    out = []
    for d in sorted(프로젝트들.iterdir()) if 프로젝트들.exists() else []:
        f = d / 설계파일
        if f.exists():
            try:
                장수 = len(json.loads(f.read_text("utf-8")).get("장", []))
            except (OSError, ValueError, AttributeError):
                장수 = "?"          # 덱 하나가 깨져도 목록은 열린다(.이전 폴더에 이전 판이 있다)
            out.append({"이름": d.name, "장수": 장수,
                        "고친때": time.strftime("%Y-%m-%d %H:%M", time.localtime(f.stat().st_mtime))})
    return out


def 읽기(이름: str) -> dict:
    return json.loads((폴더(이름) / 설계파일).read_text("utf-8"))


def 쓰기(이름: str, 설계: dict) -> None:
    f = 폴더(이름) / 설계파일
    if f.exists():
        이전 = f.parent / ".이전"
        이전.mkdir(exist_ok=True)
        때 = time.strftime("%Y%m%d_%H%M%S")
        백업, n = 이전 / f"덱설계_{때}.json", 2
        while 백업.exists():                  # 같은 초에 여러 번 저장해도 덮어쓰지 않는다
            백업, n = 이전 / f"덱설계_{때}_{n}.json", n + 1
        백업.write_bytes(f.read_bytes())
    # 임시 파일에 다 쓴 뒤 바꿔 넣는다: 쓰는 도중 꺼지거나 굽기가 읽어도 반쯤 쓴 파일이 없다
    임시 = f.with_name(f".{설계파일}.쓰는중")
    임시.write_text(json.dumps(설계, ensure_ascii=False, indent=2), "utf-8")
    os.replace(임시, f)


def 소재목록(이름: str) -> list[dict]:
    p = 폴더(이름)
    out = []
    후보 = sorted((p / "소재").rglob("*")) if (p / "소재").exists() else []
    후보 += sorted((p / "재료").glob("*/그림/*")) if (p / "재료").exists() else []
    for f in 후보:
        if f.is_file() and ".미리보기" not in f.parts and f.suffix.lower() in (
                소재.브라우저그림 | 소재.파워포인트그림 | {".tif", ".tiff"}):
            상대 = str(f.relative_to(폴더(이름))).replace("\\", "/")
            w, h = 소재.크기안전(f) or (None, None)      # 못 읽는 그림도 목록에는 남긴다(감독관이 알림)
            out.append({"경로": 상대, "이름": f.name, "너비": w, "높이": h,
                        "보기": 그림주소(이름)(상대)})
    return out


# ---------------------------------------------------------------- 배치
def _표붙이기(장: dict, 프로젝트: Path | None, 설계: dict | None = None) -> dict:
    """설계 파일엔 경로만 두고, 그릴 때 필요한 것을 붙인다.
    표(재료/.../표/x.json) → _표, 원본장 → _쪽(v1 장을 구운 그림)."""
    if 장.get("종류") == "원본장" and 프로젝트 is not None and 설계 and 설계.get("바탕"):
        쪽 = f"재료/{설계['바탕']['재료']}/쪽/{int(장['원본']):02d}.png"
        return {**장, "_쪽": 쪽} if (프로젝트 / 쪽).exists() else 장
    if not 장.get("표") or 프로젝트 is None:
        return 장
    f = (프로젝트 / 장["표"]).resolve()
    if 프로젝트.resolve() in f.parents and f.exists():
        return {**장, "_표": json.loads(f.read_text("utf-8"))}
    return 장


def 덱배치(설계: dict, 프로젝트: Path | None = None) -> dict:
    스타일이름, 스타일 = 스.고르기(설계.get("스타일"))
    글자체이름, _ = 스.글자체(설계.get("글자체"))
    장들 = []
    for 장 in 설계.get("장", []):
        장 = _표붙이기(장, 프로젝트, 설계)
        박스들 = 꾸밈.입히기([b.사전() for b in 배치.장배치(장, 글자체이름, 스타일.get("장식"))],
                         설계.get("꾸밈"))
        장들.append({"종류": 장.get("종류"), "박스": 박스들, "노트": 장.get("노트", ""),
                    "넘침": sum(b["넘침"] for b in 박스들)})
    색 = 스.색갖추기(스타일)
    return {"스타일": 스타일이름, "색": 색, "글자체": 글자체이름, "장들": 장들}


# 그림 파일이 없을 때 자리에 그리는 표시(덱이 안 열리면 안 된다. 감독관이 '그림없음'으로 잡는다)
없는그림 = ("data:image/svg+xml;charset=utf-8," + quote(
    "<svg xmlns='http://www.w3.org/2000/svg' width='480' height='270'><rect width='480' height='270' "
    "fill='#e9e9e9' stroke='#e5484d' stroke-width='6' stroke-dasharray='18 12'/><text x='240' y='148' "
    "font-size='34' text-anchor='middle' fill='#e5484d' font-family='sans-serif'>그림 없음</text></svg>"))


def 그림주소(이름: str):
    """화면(서버)용 그림 주소. TIFF 등은 화면용 PNG 미리보기로."""
    def 주소(상대: str) -> str:
        if not (폴더(이름) / 상대).is_file():
            return 없는그림
        try:
            return f"/프로젝트/{quote(이름)}/{quote(소재.미리보기(폴더(이름), 상대))}"
        except Exception:  # noqa: BLE001 — 깨진 그림이어도 덱은 열린다
            return 없는그림
    return 주소


def 화면자료(이름: str, 설계: dict | None = None) -> dict:
    설계 = 설계 if 설계 is not None else 읽기(이름)
    덱 = 덱배치(설계, 폴더(이름))
    주소 = 그림주소(이름)
    for 장 in 덱["장들"]:
        장["html"] = 화면.장html(장["박스"], 덱["색"], 주소, 글자체=덱["글자체"])
    원본글 = 바탕pptx.원본장들(폴더(이름) / 설계["바탕"]["파일"]) if 설계.get("바탕") else []
    return {"이름": 이름, "설계": 설계, **덱, "원본글": 원본글,
            "감독": 감독관.살피기(설계, 덱, 폴더(이름)),
            "스타일들": {k: {"설명": v["설명"], "색": 스.색갖추기(v)} for k, v in 스.스타일들.items()},
            "내스타일들": {k: {"설명": v.get("설명", ""), "읽은것": v.get("읽은것", []),
                          "원본": {a: b for a, b in v.items() if a != "읽은것"}}
                       for k, v in 스.내스타일들().items()},
            "글자체들": list(스.글자체들), "장종류": 배치.장종류, "꾸밈들": 꾸밈.고를것, "소재": 소재목록(이름),
            "재료": 재료목록(이름)}


def 재료목록(이름: str) -> list[dict]:
    """재료함: 쪽·그림은 화면 주소를 붙여 돌려준다."""
    주소 = 그림주소(이름)
    out = []
    for m in 재료받기.목록(폴더(이름)):
        m["쪽보기"] = [주소(p) for p in m["쪽"]]
        m["그림보기"] = [{"경로": p, "보기": 주소(p)} for p in m["그림"]]
        out.append(m)
    return out


def 재료넣기(이름: str, 파일이름: str | None = None, 내용64: str | None = None,
          경로: str | None = None) -> dict:
    """브라우저에서 올린 파일(base64) 또는 이 PC의 경로를 재료로 넣는다."""
    p = 폴더(이름)
    if 경로:
        f = Path(경로.strip().strip('"'))
        if not f.is_file():
            raise ValueError(f"파일을 찾지 못했어요: {f}")
        return 재료받기.넣기(p, f)
    return 재료받기.넣기(p, 이름=파일이름, 내용=base64.b64decode(내용64 or ""))


def 미리보기(이름: str, 장: dict, 스타일이름, 글자체: str | None = None, 꾸밈값: dict | None = None) -> str:
    """테마·꾸미기 타일: 장 하나를 다른 스타일·꾸밈으로 그린 HTML."""
    _, 스타일 = 스.고르기(스타일이름)
    박스들 = 꾸밈.입히기([b.사전() for b in 배치.장배치(_표붙이기(장, 폴더(이름)), 글자체, 스타일.get("장식"))],
                     꾸밈값)
    색 = 스.색갖추기(스타일)
    return 화면.장html(박스들, 색, 그림주소(이름), 표시넘침=False, 글자체=글자체)


# ---------------------------------------------------------------- 굽기
def 굽기실행(이름: str, 형식들: list[str], 강행: bool = False) -> dict:
    """감독관이 '멈춤'을 찾으면 굽지 않고 돌려준다(강행=True 면 사람이 알고 굽는 것).
    굽고 나서는 완성 점검(장 수, PPTX 글상자)을 붙인다."""
    설계 = 읽기(이름)
    p = 폴더(이름)
    덱 = 덱배치(설계, p)
    찾음 = 감독관.살피기(설계, 덱, p)
    if 감독관.멈춤(찾음) and not 강행:
        return {"멈춤": 감독관.멈춤(찾음), "파일": [], "알림": []}
    결과 = _굽기(이름, 설계, p, 덱, 형식들)
    결과["점검"] = 감독관.완성점검(설계, p / 결과["폴더"], 덱)
    return 결과


def 파일이름(이름: str) -> str:
    """덱 이름 → 파일 이름. 윈도우에서 못 쓰는 글자(/ \\ : * ? " < > |)와 앞뒤 점·빈칸만 바꾼다."""
    import re
    깨끗 = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "_", str(이름)).strip(" .")
    return 깨끗[:80] or "덱"


형식목록 = ["pptx", "pdf", "png", "mp4", "대본", "데이터", "소재"]


def _굽기(이름: str, 설계: dict, p: Path, 덱: dict, 형식들: list[str]) -> dict:
    때 = time.strftime("%Y%m%d_%H%M%S")
    out = p / "내보내기" / 때
    n = 2
    while out.exists():          # 같은 초에 두 번 구워도 덮어쓰지 않는다
        out, n = p / "내보내기" / f"{때}_{n}", n + 1
    out.mkdir(parents=True)
    덱이름 = 파일이름(설계.get("이름") or 이름)
    결과: dict = {"폴더": str(out.relative_to(p)).replace("\\", "/"), "전체폴더": str(out.resolve()),
                "파일": [], "알림": []}
    형식들 = [f for f in 형식목록 if f in 형식들]

    def 파일주소(path: Path) -> dict:
        상대 = str(path.relative_to(p)).replace("\\", "/")
        return {"이름": path.name, "주소": f"/프로젝트/{quote(이름)}/{quote(상대)}"}

    그림필요 = bool({"png", "mp4"} & set(형식들))
    pdf필요 = 그림필요 or "pdf" in 형식들
    pdf = out / f"{덱이름}.pdf" if "pdf" in 형식들 else out / "장" / "전체.pdf"   # 안 고르면 작업 폴더에
    if 설계.get("바탕"):        # v1 템플릿을 그대로 두고 v2 를 만든다(PDF·그림도 v2 PPTX에서)
        _v2굽기(설계, 덱, p, out, 결과, 파일주소, pdf if pdf필요 else None)
    else:
        if "pptx" in 형식들:
            f = out / f"{덱이름}.pptx"
            결과["알림"] += 파포.만들기([{**장, "색": 덱["색"]} for 장 in 덱["장들"]], p, f, 덱["글자체"])
            결과["파일"].append(파일주소(f))
        if pdf필요:
            글자체주소 = (뿌리 / "글자체").resolve().as_uri()

            def 파일그림(상대: str) -> str:
                if not (p / 상대).is_file():
                    return 없는그림
                try:
                    return (p / 소재.미리보기(p, 상대)).resolve().as_uri()
                except Exception:  # noqa: BLE001
                    return 없는그림
            조각들 = [화면.장html(장["박스"], 덱["색"], 파일그림, 표시넘침=False, 글자체=덱["글자체"])
                    for 장 in 덱["장들"]]
            h = out / "장" / "전체.html"
            h.parent.mkdir(parents=True, exist_ok=True)
            h.write_text(화면.문서(조각들, 글자체주소, 덱이름), "utf-8")
            굽기.pdf(h, pdf)
    if not pdf.exists():                 # v2에서 PDF를 못 만든 경우(알림은 이미 남겼다)
        형식들 = [x for x in 형식들 if x not in ("pdf", "png", "mp4")]
        그림필요 = False
    if "pdf" in 형식들:
        결과["파일"].append(파일주소(pdf))
    그림들 = 굽기.pdf그림들(pdf, out / "그림") if 그림필요 else []
    if "png" in 형식들:
        결과["파일"].append(파일주소(내보내기.그림묶음(그림들, out / f"{덱이름}_그림묶음.zip")))
    if "mp4" in 형식들:
        f = out / f"{덱이름}.mp4"
        말 = 내보내기.영상(그림들, 내보내기.장시간들(설계), f)
        if 말:
            결과["알림"].append(말)
        else:
            결과["파일"].append(파일주소(f))
    if "대본" in 형식들:
        결과["파일"].append(파일주소(내보내기.대본(설계, out / "대본.md")))
    if "데이터" in 형식들:
        f = 내보내기.데이터(설계, lambda 장: _표붙이기(장, p, 설계), out / "데이터.xlsx")
        if f:
            결과["파일"].append(파일주소(f))
        else:
            결과["알림"].append("도표·표 장이 없어 데이터.xlsx는 만들지 않았어요")
    if "소재" in 형식들:
        모은것 = 내보내기.소재모으기(설계, p, out / "소재")
        결과["알림"].append(f"덱에 쓴 그림 원본 {len(모은것)}개를 소재 폴더에 그대로 복사했어요" if 모은것
                         else "덱에 쓴 그림이 없어요")
    return 결과


# ---------------------------------------------------------------- v1 → v2
def _v2굽기(설계, 덱, p: Path, out: Path, 결과, 파일주소, pdf: Path | None) -> None:
    v1 = p / 설계["바탕"]["파일"]
    이름 = 바탕pptx.다음판이름(Path(설계["바탕"]["원래이름"]))
    f = out / f"{이름}.pptx"
    알림: list[str] = []

    def 새장(prs, slide, 장자료):
        파포.판맞추기(prs)
        파포.장그리기(slide, 장자료, p, 덱["글자체"], 알림, 바탕칠=False)

    장들 = [원래 if 원래.get("종류") == "원본장" else
          {**배치된, "색": 덱["색"], "제목": 원래.get("제목") or 원래.get("큰말") or ""}
          for 원래, 배치된 in zip(설계["장"], 덱["장들"])]
    기록 = 바탕pptx.v2만들기(v1, 장들, 새장, f)
    (out / "변경기록.md").write_text(
        f"# {f.name}\n\n바탕: {설계['바탕']['원래이름']} (v1은 그대로 두었어요)\n\n" + "\n".join(기록) + "\n",
        "utf-8")
    결과["파일"] += [파일주소(f), 파일주소(out / "변경기록.md")]
    결과["알림"] += 알림
    if pdf is not None:
        pdf.parent.mkdir(parents=True, exist_ok=True)
        try:
            결과["알림"].append(f"PDF·그림은 {굽기.pptx_pdf(f, pdf)}로 v2를 그대로 구웠어요")
        except Exception as e:  # noqa: BLE001 — v2 PPTX는 이미 만들어졌으니 그것은 돌려준다
            결과["알림"].append(f"v2 PPTX는 만들었지만 PDF·그림·영상은 못 만들었어요: {e}")


def 폴더열기(이름: str, 상대: str) -> dict:
    """구운 결과 폴더를 탐색기(Finder)로 연다. 이 덱 폴더 안만 연다."""
    import subprocess
    import sys
    p = 폴더(이름).resolve()
    f = (p / 상대).resolve()
    if f != p and p not in f.parents or not f.is_dir():
        raise ValueError("열 수 없는 폴더예요")
    if sys.platform.startswith("win"):
        os.startfile(str(f))  # noqa: S606 — 사용자가 누른 단추로 자기 폴더를 연다
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(f)])
    else:
        import shutil
        if shutil.which("xdg-open"):
            subprocess.Popen(["xdg-open", str(f)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return {"폴더": str(f)}


def 감독고치기(이름: str, 규칙: str, 장: int, 고침: str) -> dict:
    설계, 말 = 감독관.고치기(읽기(이름), 규칙, int(장), 고침)
    쓰기(이름, 설계)
    return {**화면자료(이름, 설계), "말": 말}


def 감독무시(이름: str, 규칙: str, 장: int, 말: str | None = None) -> dict:
    설계 = 감독관.무시하기(읽기(이름), 규칙, int(장), 말)
    쓰기(이름, 설계)
    return 화면자료(이름, 설계)


def 해부하기(이름: str, 파일이름: str | None = None, 내용64: str | None = None,
          경로: str | None = None, 덱에쓰기: bool = True) -> dict:
    """참고 PPTX → 내 스타일. 참고 PPTX는 재료로 들어가고(원본 그대로), 스타일은 내스타일/에 남는다."""
    p = 폴더(이름)
    m = (재료받기.넣기(p, Path(경로.strip().strip('"'))) if 경로 else
         재료받기.넣기(p, 이름=파일이름, 내용=base64.b64decode(내용64 or "")))
    if m["종류"] != "PPTX":
        raise ValueError("해부는 PPTX 파일만 돼요")
    스타일 = 해부.남기기(해부.해부(p / "재료" / m["아이디"] / "원본.pptx",
                             Path(m["원래이름"]).stem, m["원래이름"]))
    if 덱에쓰기:
        설계 = 읽기(이름)
        설계["스타일"] = {k: v for k, v in 스타일.items() if k != "읽은것"}
        if 스타일.get("글자체"):
            설계["글자체"] = 스타일["글자체"]
        쓰기(이름, 설계)
    return 스타일


def 새덱(덱이름: str, v1경로: str | None = None, 파일이름: str | None = None,
       내용64: str | None = None) -> str:
    """새 프로젝트. v1 PPTX를 주면 그 템플릿을 바탕으로(장은 원본장으로 들어온다)."""
    이름 = 재료받기._이름다듬기(덱이름)
    p = 프로젝트들 / 이름
    if (p / 설계파일).exists():
        raise ValueError(f"같은 이름의 덱이 있어요: {이름}. 덱 이름 칸에 다른 이름을 넣어 주세요")
    # 전에 만들다 만 폴더(덱설계.json 없음)는 지우지 않고 그대로 이어 쓴다
    (p / "소재").mkdir(parents=True, exist_ok=True)
    설계: dict = {"이름": 덱이름, "장": []}
    if v1경로 or 내용64:
        m = (재료받기.넣기(p, Path(v1경로.strip().strip('"'))) if v1경로 else
             재료받기.넣기(p, 이름=파일이름, 내용=base64.b64decode(내용64)))
        if m["종류"] != "PPTX":
            raise ValueError("바탕은 PPTX 파일이어야 해요")
        v1 = p / "재료" / m["아이디"] / "원본.pptx"
        테마 = 바탕pptx.테마읽기(v1)
        설계.update({
            "바탕": {"재료": m["아이디"], "파일": str(v1.relative_to(p)).replace("\\", "/"),
                     "원래이름": m["원래이름"]},
            "스타일": {"이름": "v1 템플릿", **{k: 테마[k] for k in ("설명", "색", "계열")}},
            "글자체": 테마["글자체"],
            "장": [{"종류": "원본장", "원본": n} for n in range(1, (m.get("쪽수") or 0) + 1)],
        })
    쓰기(이름, 설계)
    return 이름
