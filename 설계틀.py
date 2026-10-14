"""설계틀: 덱설계 JSON의 스키마(덱설계.schema.json)를 앱의 장 종류 표에서 만든다.

    python 설계틀.py          → 덱설계.schema.json 과 덱설계_쓰는법.md 의 장 종류 표 다시 쓰기
장 종류나 칸을 바꾸면 이걸 다시 돌린다(시험이 어긋나면 알려 준다).
"""

from __future__ import annotations

import json
from pathlib import Path

import 꾸밈
import 배치
import 스타일 as 스

파일 = Path(__file__).parent / "덱설계.schema.json"
쓰는법 = Path(__file__).parent / "덱설계_쓰는법.md"
_시작, _끝 = "<!-- 장종류 시작 -->", "<!-- 장종류 끝 -->"


def _칸틀(칸: dict) -> dict:
    한도 = 칸.get("한도")
    설명 = 칸["이름"] + (f" (한 줄 {한도}자 안)" if 한도 else "")
    형태 = 칸["형태"]
    if 형태 == "한줄":
        return {"type": "string", "description": 설명}
    if 형태 == "여러줄":
        return {"description": 설명 + f", {칸.get('줄수', '몇')}줄까지",
                "anyOf": [{"type": "array", "items": {"type": "string"}}, {"type": "string"}]}
    if 형태 == "그림":
        return {"type": "string", "description": "그림 경로(프로젝트 폴더 기준, 예: 재료/003_SEM/그림/SEM.tif)"}
    if 형태 == "그림여럿":
        return {"type": "array", "items": {"type": "string"}, "maxItems": 칸.get("최대", 6),
                "description": "그림 경로 목록(누른 차례대로)"}
    if 형태 == "표":
        return {"type": "string", "description": "표 경로(재료/NNN_이름/표/<이름>.json)"}
    if 형태 == "고르기":
        return {"enum": 칸["고를것"], "description": 칸["이름"]}
    if 형태 == "원본글":
        return {"type": "object", "additionalProperties": {"type": "string"},
                "description": "v1 글상자 아이디 → 새 글"}
    return {}


def 스키마() -> dict:
    공통 = {"노트": {"type": "string", "description": "발표자 노트(대본.md에 모인다)"},
          "시간": {"type": "number", "minimum": 1, "description": "영상에서 이 장을 보여 줄 초"}}
    도표덧 = {"축0": {"type": "boolean", "description": "축을 0에서 시작(막대는 늘 참)"},
            "합100": {"type": "boolean", "description": "줄마다 합이 100이어야 함(비율)"},
            "가로": {"type": "string", "description": "가로축으로 쓸 열 이름(기본: 첫 열)"},
            "세로": {"type": "array", "items": {"type": "string"}, "description": "그릴 열 이름들(기본: 숫자 열 전부)"},
            "단위": {"type": "string"}}
    장들 = []
    for 종류, 틀 in 배치.장종류.items():
        속성 = {"종류": {"const": 종류}, **공통, **{c["이름"]: _칸틀(c) for c in 틀["칸"]}}
        if 종류 == "도표":
            속성.update(도표덧)
        if 종류 == "원본장":
            속성["원본"] = {"type": "integer", "minimum": 1, "description": "v1의 몇 번째 장"}
        필수 = ["종류"] + [c["이름"] for c in 틀["칸"] if c.get("필수")]
        장들.append({"title": 종류, "description": 틀["설명"], "type": "object", "properties": 속성,
                    "required": 필수, "additionalProperties": False})
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "$id": "덱설계.schema.json",
        "title": "덱설계",
        "description": "덱편집기가 받는 발표 덱 설계. 덱설계_쓰는법.md 를 함께 본다.",
        "type": "object",
        "required": ["이름", "장"],
        "properties": {
            "이름": {"type": "string", "description": "덱 이름(결과 파일 이름)"},
            "스타일": {"description": "스타일 이름 또는 해부로 만든 스타일 사전",
                     "anyOf": [{"enum": list(스.스타일들)}, {"type": "string"}, {"type": "object"}]},
            "글자체": {"type": "string", "description": "기본 '맑은 고딕'"},
            "발표분": {"type": "number", "description": "발표 시간(분). 감독관이 장 수를 살핀다"},
            "장시간": {"type": "number", "description": "영상에서 한 장 기본 초(기본 5)"},
            "꾸밈": {"type": "object", "additionalProperties": False,
                   "properties": {k: {"enum": list(v)} for k, v in 꾸밈.고를것.items()}},
            "바탕": {"type": "object", "description": "v1 PPTX 바탕(앱이 채운다)"},
            "감독관무시": {"type": "array", "items": {"type": "object"}},
            "휴지통": {"type": "array", "items": {"type": "object"}},
            "장": {"type": "array", "minItems": 1, "items": {"oneOf": 장들}},
        },
    }


_형태말 = {"한줄": "글", "여러줄": "글 목록", "그림": "그림 경로", "그림여럿": "그림 경로 목록", "표": "표 경로",
         "고르기": "고르기", "원본글": "v1 글상자 고침"}


def 장종류표() -> str:
    """덱설계_쓰는법.md 에 들어갈 장 종류 표(앱의 장 종류에서 만든다)."""
    줄 = ["(이 표는 `python 설계틀.py`가 앱의 장 종류에서 만든다. 손으로 고치지 않는다.)", "",
         "| 종류 | 쓰는 때 | 칸 |", "|---|---|---|"]
    for 종류, 틀 in 배치.장종류.items():
        if 틀.get("추가안함"):
            continue
        칸들 = []
        for c in 틀["칸"]:
            덧 = []
            if c.get("한도"):
                덧.append(f"{c['한도']}자")
            if c.get("줄수"):
                덧.append(f"{c['줄수']}줄")
            if c.get("최대"):
                덧.append(f"{c['최대']}장")
            if c.get("고를것"):
                덧.append("/".join(c["고를것"]))
            칸들.append(f"`{c['이름']}`{'*' if c.get('필수') else ''} {_형태말[c['형태']]}"
                       + (f"({', '.join(덧)})" if 덧 else ""))
        설명 = 틀["설명"].replace("|", "\\|")          # 표 안의 | 는 \| 로
        줄.append(f"| {종류} | {설명} | {' · '.join(칸들)} |")
    return "\n".join(줄)


def 쓰는법글() -> str:
    글 = 쓰는법.read_text("utf-8")
    a, b = 글.index(_시작) + len(_시작), 글.index(_끝)
    return 글[:a] + "\n" + 장종류표() + "\n" + 글[b:]


def 쓰기() -> Path:
    파일.write_text(json.dumps(스키마(), ensure_ascii=False, indent=1) + "\n", "utf-8")
    쓰는법.write_text(쓰는법글(), "utf-8")
    return 파일


if __name__ == "__main__":
    print(쓰기())
