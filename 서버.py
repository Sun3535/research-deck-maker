"""덱편집기 서버: http://127.0.0.1:8780

- 내 PC 안(127.0.0.1)에서만 열린다. 바깥 인터넷을 쓰지 않는다.
- 덧붙는 길(/api/...)은 전부 POST 로 받는다.
- 파이썬 표준 라이브러리만 쓴다(프레임워크 설치 없음).
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import sys
import threading
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

sys.path.insert(0, str(Path(__file__).parent))
import 엔진  # noqa: E402

뿌리 = Path(__file__).parent
_자물쇠 = threading.Lock()   # 엔진은 한 번에 하나씩


def _안전한경로(밑: Path, 상대: str) -> Path | None:
    p = (밑 / 상대).resolve()
    return p if (p == 밑.resolve() or 밑.resolve() in p.parents) and p.is_file() else None


# ---------------------------------------------------------------- 길(API)
def 길_목록(_):
    return {"프로젝트": 엔진.목록()}


def 길_열기(요청):
    return 엔진.화면자료(요청["이름"])


def 길_저장(요청):
    엔진.쓰기(요청["이름"], 요청["설계"])
    return 엔진.화면자료(요청["이름"], 요청["설계"])


def 길_배치(요청):
    """저장하지 않고 다시 그리기만(입력 중 미리보기)."""
    return 엔진.화면자료(요청["이름"], 요청["설계"])


def 길_미리보기(요청):
    return {"html": 엔진.미리보기(요청["이름"], 요청["장"], 요청["스타일"], 요청.get("글자체"), 요청.get("꾸밈"))}


def 길_굽기(요청):
    return 엔진.굽기실행(요청["이름"], 요청.get("형식", ["pptx"]), bool(요청.get("강행")))


def 길_감독고치기(요청):
    return 엔진.감독고치기(요청["이름"], 요청["규칙"], 요청["장"], 요청["고침"])


def 길_폴더열기(요청):
    return 엔진.폴더열기(요청["이름"], 요청["폴더"])


def 길_감독무시(요청):
    return 엔진.감독무시(요청["이름"], 요청["규칙"], 요청["장"], 요청.get("말"))


def 길_재료넣기(요청):
    엔진.재료넣기(요청["이름"], 요청.get("파일이름"), 요청.get("내용"), 요청.get("경로"))
    return 엔진.화면자료(요청["이름"])


def 길_해부(요청):
    스타일 = 엔진.해부하기(요청["이름"], 요청.get("파일이름"), 요청.get("내용"), 요청.get("경로"))
    return {**엔진.화면자료(요청["이름"]), "해부": 스타일}


def 길_새덱(요청):
    이름 = 엔진.새덱(요청["덱이름"], 요청.get("경로"), 요청.get("파일이름"), 요청.get("내용"))
    return {"이름": 이름}


길들 = {"새덱": 길_새덱, "해부": 길_해부, "감독고치기": 길_감독고치기, "감독무시": 길_감독무시, "폴더열기": 길_폴더열기, "재료넣기": 길_재료넣기, "목록": 길_목록, "열기": 길_열기, "저장": 길_저장, "배치": 길_배치,
       "미리보기": 길_미리보기, "굽기": 길_굽기}


class 처리기(BaseHTTPRequestHandler):
    server_version = "DeckEditor/0.1"   # 머리글은 영문만 된다

    def log_message(self, fmt, *args):  # 조용히
        pass

    def _보내기(self, 코드: int, 몸: bytes, 종류: str) -> None:
        self.send_response(코드)
        self.send_header("Content-Type", 종류)
        self.send_header("Content-Length", str(len(몸)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(몸)

    def _json(self, 코드: int, 값) -> None:
        self._보내기(코드, json.dumps(값, ensure_ascii=False).encode(), "application/json; charset=utf-8")

    def _이PC에서(self) -> bool:
        """이 PC의 이 화면에서 온 요청만 받는다. 다른 웹사이트가 몰래 부르는 것(CSRF·DNS rebinding)을 막는다."""
        포트 = self.server.server_address[1]
        허용 = {f"127.0.0.1:{포트}", f"localhost:{포트}"}
        if self.headers.get("Host", "") not in 허용:
            return False
        출처 = self.headers.get("Origin")
        return 출처 is None or 출처 in {f"http://{h}" for h in 허용}

    def do_GET(self):
        if not self._이PC에서():
            return self._json(403, {"오류": "이 PC의 덱편집기 화면에서만 열 수 있어요"})
        길 = unquote(urlparse(self.path).path)
        if 길 in ("/", "/index.html"):
            길 = "/화면/index.html"
        if 길 == "/장.css":  # 장을 그리는 CSS는 굽기와 같은 것을 쓴다
            import 그리기_html
            css = 그리기_html.글자체css("/글자체") + 그리기_html.장css
            return self._보내기(200, css.encode(), "text/css; charset=utf-8")
        for 머리, 밑 in (("/화면/", 뿌리 / "화면"), ("/글자체/", 뿌리 / "글자체"),
                       ("/프로젝트/", 엔진.프로젝트들)):
            if 길.startswith(머리):
                p = _안전한경로(밑, 길[len(머리):])
                if p:
                    종류 = mimetypes.guess_type(p.name)[0] or "application/octet-stream"
                    if 종류.startswith("text/") or 종류 in ("application/javascript",):
                        종류 += "; charset=utf-8"
                    return self._보내기(200, p.read_bytes(), 종류)
        self._json(404, {"오류": "없는 주소예요"})

    def do_POST(self):
        if not self._이PC에서():
            return self._json(403, {"오류": "이 PC의 덱편집기 화면에서만 쓸 수 있어요"})
        if not self.headers.get("Content-Type", "").startswith("application/json"):
            return self._json(415, {"오류": "JSON으로만 받아요"})   # 다른 사이트의 단순 폼 전송을 막는다
        길 = unquote(urlparse(self.path).path)
        if not 길.startswith("/api/") or 길[5:] not in 길들:
            return self._json(404, {"오류": "없는 길이에요"})
        try:
            n = int(self.headers.get("Content-Length", 0))
            요청 = json.loads(self.rfile.read(n) or b"{}")
            with _자물쇠:
                self._json(200, 길들[길[5:]](요청))
        except Exception as exc:  # 화면에 쉬운 말로 알린다
            traceback.print_exc()
            self._json(500, {"오류": str(exc)})


def main(argv=None) -> None:
    ap = argparse.ArgumentParser(description="덱편집기 서버")
    ap.add_argument("--포트", type=int, default=8780)
    ap.add_argument("--창안열기", action="store_true", help="브라우저를 자동으로 열지 않기")
    a = ap.parse_args(argv)
    mimetypes.add_type("application/javascript", ".js")
    mimetypes.add_type("font/otf", ".otf")
    서버 = ThreadingHTTPServer(("127.0.0.1", a.포트), 처리기)
    주소 = f"http://127.0.0.1:{a.포트}/"
    print(f"덱편집기가 켜졌어요: {주소}  (끄려면 이 창에서 Ctrl+C)")
    if not a.창안열기:
        webbrowser.open(주소)
    try:
        서버.serve_forever()
    except KeyboardInterrupt:
        print("\n덱편집기를 껐어요.")


if __name__ == "__main__":
    main()
