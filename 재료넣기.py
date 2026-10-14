"""재료넣기: 파일들을 덱의 재료함에 넣고, Claude가 읽을 '재료요약.md'를 만든다(화면 없이).

    python 재료넣기.py 내덱 실험메모.md 데이터.xlsx 회의록.pdf SEM_1.tif
    python 재료넣기.py 내덱                     (넣지 않고 요약만 다시 만들기)

덱(프로젝트)이 없으면 빈 덱을 만든다. 넣은 파일은 그대로 두고 복사본을 푼다.
재료요약.md: 재료마다 글 덩어리, 표(머리·줄 수·경로), 그림(경로·크기). 덱설계 JSON에 그대로 쓸 경로가 적혀 있다.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import 소재
import 엔진
import 재료받기


def 요약(이름: str) -> Path:
    p = 엔진.폴더(이름)
    줄 = [f"# 재료 요약: {이름}", "",
          "덱설계 JSON에는 아래 경로를 그대로 쓴다(프로젝트 폴더 기준).", ""]
    for m in 재료받기.목록(p):
        줄 += [f"## {m['아이디']} · {m['종류']} · {m['원래이름']}", ""]
        if m.get("꺼낸것"):
            줄.append(f"- 꺼낸 것: {m['꺼낸것']}")
        for 말 in m.get("알림", []):
            줄.append(f"- 알림: {말}")
        if m.get("글"):
            줄 += ["", "### 글", ""]
            for d in m["글"]:
                글 = d["글"].replace("\n", " / ")
                줄.append(f"- (쪽 {d.get('쪽', 1)}) {글}")
        for t in m.get("표", []):           # 숫자는 재료에 있는 것만 쓰도록 표를 그대로 보여 준다
            전체 = json.loads((p / t["경로"]).read_text("utf-8"))
            줄 += ["", f"### 표 `{t['경로']}` ({len(전체['줄들'])}줄)", "",
                  "| " + " | ".join(map(str, 전체["머리"])) + " |", "|" + "---|" * len(전체["머리"])]
            줄 += ["| " + " | ".join("" if v is None else str(v) for v in r) + " |" for r in 전체["줄들"][:12]]
            if len(전체["줄들"]) > 12:
                줄.append(f"(… {len(전체['줄들']) - 12}줄 더)")
        if m.get("그림"):
            줄 += ["", "### 그림", ""]
            for g in m["그림"]:
                try:
                    w, h = 소재.크기(p / g)
                    줄.append(f"- `{g}` ({w}×{h})")
                except OSError:
                    줄.append(f"- `{g}`")
        줄.append("")
    f = p / "재료요약.md"
    f.write_text("\n".join(줄), "utf-8")
    return f


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    a = sys.argv[1:] if argv is None else argv
    if not a:
        print(__doc__)
        return 1
    이름 = 재료받기._이름다듬기(a[0])
    if not (엔진.프로젝트들 / 이름 / "덱설계.json").exists():
        이름 = 엔진.새덱(a[0])
        print(f"새 덱을 만들었어요: {엔진.폴더(이름)}")
    for 파일 in a[1:]:
        f = Path(파일.strip().strip('"'))
        if not f.is_file():
            print(f"✗ 파일이 없어요: {f}")
            continue
        try:
            m = 엔진.재료넣기(이름, 경로=str(f))
        except ValueError as e:
            print(f"✗ {f.name}: {e}")
            continue
        print(f"✓ {f.name} → 재료/{m['아이디']} ({m['종류']}: {m.get('꺼낸것', '')})")
    print(f"재료 요약: {요약(이름)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
