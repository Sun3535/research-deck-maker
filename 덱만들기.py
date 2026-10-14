"""덱만들기: 덱설계 JSON 하나 → 결과 전부를 한 명령으로(화면 없이).

    python 덱만들기.py 프로젝트\\내덱\\덱설계.json
    python 덱만들기.py 프로젝트\\내덱                  (폴더를 줘도 된다)
    python 덱만들기.py 내덱 --형식 pptx,pdf          (프로젝트 이름만 줘도 된다)
    python 덱만들기.py 내덱 --강행                    (감독관이 멈춰도 굽는다)

형식: pptx, pdf, png, mp4, 대본, 데이터, 소재 (기본: 전부)
감독관이 '멈춤'을 찾으면 굽지 않고 이유를 보여 준다(끝 코드 2).
결과는 그 프로젝트의 내보내기/<날짜_시간>/ 에 쌓인다(지우거나 덮어쓰지 않는다).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import 감독관
import 엔진


def _프로젝트(대상: str) -> str:
    """경로나 이름 → 엔진이 쓰는 프로젝트 이름. 프로젝트 폴더 밖의 덱이면 그 부모 폴더를 프로젝트 모음으로 쓴다."""
    p = Path(대상.strip().strip('"'))
    if p.suffix.lower() == ".json":
        p = p.parent
    if p.is_dir() and (p / "덱설계.json").is_file():
        엔진.프로젝트들 = p.resolve().parent
        return p.name
    if (엔진.프로젝트들 / 대상 / "덱설계.json").is_file():
        return 대상
    raise SystemExit(f"덱설계.json 을 찾지 못했어요: {대상}")


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):       # 윈도우 콘솔에서도 한글이 깨지지 않게
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="덱설계 JSON → PPTX·PDF·PNG·mp4·대본·데이터")
    ap.add_argument("대상", help="덱설계.json, 프로젝트 폴더, 또는 프로젝트 이름")
    ap.add_argument("--형식", nargs="+", default=[",".join(엔진.형식목록)],
                    help="쉼표나 빈칸으로: " + ",".join(엔진.형식목록))
    ap.add_argument("--강행", action="store_true", help="감독관이 멈춰도 굽는다")
    a = ap.parse_args(argv)
    이름 = _프로젝트(a.대상)
    형식들 = [x.strip() for 묶음 in a.형식 for x in 묶음.split(",") if x.strip()]   # cmd는 쉼표에서 인자를 나눈다
    모름 = [x for x in 형식들 if x not in 엔진.형식목록]
    if 모름:
        raise SystemExit(f"모르는 형식: {', '.join(모름)} (쓸 수 있는 것: {', '.join(엔진.형식목록)})")

    설계 = 엔진.읽기(이름)
    p = 엔진.폴더(이름)
    찾음 = 감독관.살피기(설계, 엔진.덱배치(설계, p), p)
    print(f"■ {설계.get('이름', 이름)} · {len(설계.get('장', []))}장")
    for f in 찾음:
        print(f"  [{f['수준']}] {str(f['장']) + '장' if f['장'] else '덱'}: {f['말']}")
    if not 찾음:
        print("  감독관: 찾은 것 없음 ✓")
    r = 엔진.굽기실행(이름, 형식들, 강행=a.강행)
    if r.get("멈춤"):
        print("\n감독관이 멈췄어요. 위의 [멈춤]을 고치거나, 알고 구우려면 --강행 을 붙이세요.")
        return 2
    폴더 = p / r["폴더"]
    print(f"\n결과: {폴더}")
    for f in r["파일"]:
        print(f"  ✓ {f['이름']}")
    for 말 in r["알림"]:
        print(f"  · {말}")
    for f in r.get("점검", []):
        print(f"  [완성 점검·{f['수준']}] {f['말']}")
    return 1 if any(f["수준"] == "멈춤" for f in r.get("점검", [])) else 0


if __name__ == "__main__":
    sys.exit(main())
