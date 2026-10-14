"""굽기: HTML → PNG·PDF (Edge/크롬 헤드리스), PPTX → 장 그림 (PowerPoint 또는 LibreOffice).

설치할 것이 없도록 PC에 이미 있는 프로그램만 부른다.
- 회사 PC: Edge는 윈도우에 기본으로 있다. PowerPoint가 있으면 PPTX 확인에 쓴다.
- PowerPoint를 부를 때 사용자가 열어 둔 파일이 있으면 PowerPoint를 끄지 않는다.
"""

from __future__ import annotations

import glob
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

_후보 = [
    os.environ.get("덱편집기_브라우저", ""),
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    *sorted(glob.glob("/opt/pw-browsers/chromium-*/chrome-linux/chrome"), reverse=True),
]


def 브라우저() -> str:
    for p in _후보:
        if p and Path(p).exists():
            return p
    for 이름 in ("msedge", "google-chrome", "chromium", "chromium-browser", "chrome"):
        if shutil.which(이름):
            return shutil.which(이름)
    raise RuntimeError("Edge나 크롬을 찾지 못했어요. 환경변수 덱편집기_브라우저 에 경로를 적어 주세요.")


def _실행(추가: list[str], 시간: int = 120) -> None:
    with tempfile.TemporaryDirectory() as 임시:
        명령 = [브라우저(), "--headless=new", "--disable-gpu", "--hide-scrollbars",
              "--force-device-scale-factor=1", "--allow-file-access-from-files",
              "--no-first-run", "--disable-extensions", "--virtual-time-budget=4000",
              f"--user-data-dir={임시}"]
        if sys.platform.startswith("linux") and os.geteuid() == 0:
            명령.append("--no-sandbox")
        subprocess.run(명령 + 추가, check=True, timeout=시간,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def pdf(html경로: Path, pdf경로: Path) -> Path:
    """HTML(장마다 1920×1080 한 쪽) → PDF. 글자는 PDF 안에 글자로 남는다."""
    pdf경로.parent.mkdir(parents=True, exist_ok=True)
    _실행([f"--print-to-pdf={pdf경로}", "--no-pdf-header-footer", html경로.resolve().as_uri()])
    return pdf경로


def pdf그림들(pdf경로: Path, 폴더: Path, 폭: int = 1920) -> list[Path]:
    """PDF 쪽마다 PNG. (크롬 스크린샷은 창 높이가 어긋나서 PDF를 거친다)"""
    import pypdfium2 as pdfium
    폴더.mkdir(parents=True, exist_ok=True)
    doc = pdfium.PdfDocument(str(pdf경로))
    out = []
    try:
        for i, page in enumerate(doc, start=1):
            f = 폴더 / f"{i:02d}.png"
            page.render(scale=폭 / page.get_width()).to_pil().convert("RGB").save(f)
            out.append(f)
    finally:
        doc.close()
    return out


# ---------------------------------------------------------------- PPTX → 장 그림
# PowerPoint 를 부르는 PowerShell. 실패해도 연 문서는 닫고, 이 스크립트가 켠 PowerPoint 만 끈다
# (사용자가 이미 켜 둔 PowerPoint·문서는 건드리지 않는다).
_파워포인트틀 = r"""
$ErrorActionPreference = 'Stop'
$wasOpen = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
$app = $null; $p = $null
try {{
  $app = New-Object -ComObject PowerPoint.Application
  $p = $app.Presentations.Open('{pptx}', $true, $false, $false)
  {일}
}} finally {{
  if ($p) {{ $p.Close() }}
  if ($app -and -not $wasOpen -and $app.Presentations.Count -eq 0) {{ $app.Quit() }}
}}
"""
_그림으로 = r"""for ($i = 1; $i -le $p.Slides.Count; $i++) {{
    $p.Slides.Item($i).Export(('{out}\장{{0:D2}}.png' -f $i), 'PNG', 1920, 1080)
  }}"""
_pdf로 = "$p.SaveAs('{pdf}', 32)"
마지막오류 = ""        # 실패 이유(화면 알림에 덧붙인다)


def _ps(일: str, **값) -> bool:
    """PowerPoint 일 하나를 PowerShell 로. 회사 정책(제한 언어 모드 등)으로 막히면 False."""
    global 마지막오류
    if not sys.platform.startswith("win") or not shutil.which("powershell"):
        return False
    값 = {k: str(v).replace("'", "''") for k, v in 값.items()}
    스크립트 = _파워포인트틀.format(일=일.format(**값), **값)
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", 스크립트],
                           timeout=300, capture_output=True)
    except (subprocess.SubprocessError, OSError) as e:
        마지막오류 = f"PowerPoint: {e}"
        return False
    if r.returncode != 0:
        마지막오류 = "PowerPoint: " + r.stderr.decode("utf-8", "replace").strip()[-200:]
        return False
    return True


def _soffice() -> str | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if not soffice and Path(p).exists():
            soffice = p
    return soffice


def _리브레pdf(pptx: Path, 임시: str) -> Path:
    """LibreOffice 로 PDF. 설정 폴더는 임시 폴더에(사용자 설정을 건드리지 않음)."""
    soffice = _soffice()
    if not soffice:
        raise RuntimeError("PPTX를 PDF로 바꿀 프로그램(PowerPoint/LibreOffice)이 없어요"
                           + (f" ({마지막오류})" if 마지막오류 else ""))
    subprocess.run([soffice, f"-env:UserInstallation={(Path(임시) / 'u').as_uri()}", "--headless",
                    "--convert-to", "pdf", "--outdir", 임시, str(pptx)],
                   check=True, timeout=300, capture_output=True)
    return Path(임시) / (pptx.stem + ".pdf")


def _파워포인트로(pptx: Path, 폴더: Path) -> bool:
    return _ps(_그림으로, pptx=pptx.resolve(), out=폴더.resolve())


def _리브레로(pptx: Path, 폴더: Path) -> bool:
    if not _soffice():
        return False
    with tempfile.TemporaryDirectory() as 임시:
        pdf = _리브레pdf(pptx, 임시)
        for i, f in enumerate(pdf그림들(pdf, Path(임시) / "쪽"), 1):
            f.replace(폴더 / f"장{i:02d}.png")
    return True


def pptx그림들(pptx: Path, 폴더: Path) -> tuple[list[Path], str]:
    """PPTX를 장마다 PNG로. 돌려주는 것: (그림들, 쓴 프로그램 이름)"""
    폴더.mkdir(parents=True, exist_ok=True)
    for 이름, 함수 in (("PowerPoint", _파워포인트로), ("LibreOffice", _리브레로)):
        if 함수(pptx, 폴더):
            return sorted(폴더.glob("장*.png")), 이름
    return [], "없음"


# ---------------------------------------------------------------- PPTX → PDF
def pptx_pdf(pptx: Path, pdf경로: Path) -> str:
    """PPTX → PDF (PowerPoint 있으면 PowerPoint, 없으면 LibreOffice). 쓴 프로그램 이름을 돌려준다."""
    pdf경로.parent.mkdir(parents=True, exist_ok=True)
    if _ps(_pdf로, pptx=pptx.resolve(), pdf=pdf경로.resolve()):
        return "PowerPoint"
    with tempfile.TemporaryDirectory() as 임시:
        shutil.move(str(_리브레pdf(pptx, 임시)), pdf경로)
    return "LibreOffice"
