"""소재(그림) 다루기.

원본은 절대 바꾸지 않는다. 브라우저가 못 여는 TIFF(16-bit 포함)는 화면용 PNG 미리보기만
따로 만든다. 파워포인트에는 원본 화소 그대로(무손실) 들어간다.
"""

from __future__ import annotations

import hashlib
import io
from pathlib import Path

import numpy as np
from PIL import Image

브라우저그림 = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"}
파워포인트그림 = {".png", ".jpg", ".jpeg", ".gif", ".bmp"}


def 크기(경로: Path) -> tuple[int, int]:
    with Image.open(경로) as im:
        return im.size


def 크기안전(경로: Path) -> tuple[int, int] | None:
    """크기를 못 읽으면(깨진 파일, svg 등) None. 그림 하나 때문에 덱이 안 열리면 안 된다."""
    try:
        return 크기(경로)
    except Exception:  # noqa: BLE001 — PIL은 형식마다 다른 예외를 낸다
        return None


def _8비트(im: Image.Image) -> Image.Image:
    """화면 미리보기용: 16-bit 등은 밝기 범위를 늘려 8-bit로. (원본은 그대로 둔다)"""
    if im.mode in ("I;16", "I;16B", "I;16L", "I", "F"):
        a = np.asarray(im, dtype=np.float64)
        lo, hi = np.percentile(a, [0.1, 99.9])
        a = np.clip((a - lo) / max(hi - lo, 1e-9) * 255, 0, 255).astype(np.uint8)
        return Image.fromarray(a, "L")
    if im.mode not in ("RGB", "RGBA", "L", "LA"):
        return im.convert("RGBA" if "A" in im.mode else "RGB")
    return im


def 미리보기(프로젝트: Path, 상대: str, 최대: int = 2400) -> str:
    """화면에 띄울 그림의 상대 경로. 브라우저가 열 수 있으면 원본 그대로."""
    원본 = 프로젝트 / 상대
    if 원본.suffix.lower() in 브라우저그림:
        return 상대
    # 이름이 같은 그림(다른 폴더, x.tif와 x.tiff)이 섞이지 않게 경로로 이름을 가른다
    표 = hashlib.sha1(상대.replace("\\", "/").encode("utf-8")).hexdigest()[:8]
    보기 = 프로젝트 / "소재" / ".미리보기" / f"{원본.stem}_{표}.png"
    if not 보기.exists() or 보기.stat().st_mtime < 원본.stat().st_mtime:
        보기.parent.mkdir(parents=True, exist_ok=True)
        with Image.open(원본) as im:
            im = _8비트(im)
            im.thumbnail((최대, 최대))
            im.save(보기)
    return str(보기.relative_to(프로젝트)).replace("\\", "/")


def 파워포인트용(경로: Path) -> tuple[io.BytesIO | str, str | None]:
    """파워포인트에 넣을 그림. 원본 형식을 못 넣으면 화소를 그대로 PNG로 옮긴다(무손실).

    돌려주는 것: (파일 경로나 바이트, 알릴 말)
    """
    if 경로.suffix.lower() in 파워포인트그림:
        return str(경로), None
    with Image.open(경로) as im:
        im.load()
        말 = None
        if im.mode == "RGB" and getattr(im, "tag_v2", {}).get(258, (8,))[0] > 8:
            말 = "16-bit 컬러 TIFF는 8-bit로 넣었어요(원본은 소재 폴더에 그대로)"
        if im.mode in ("I;16B", "I;16L", "I;16N"):          # 16비트(바이트 순서만 다름) → 16비트 PNG 그대로
            im = Image.fromarray(np.asarray(im).astype(np.uint16))
        elif im.mode in ("I", "F"):                          # 32비트 정수·실수
            a = np.asarray(im)
            if np.all(np.isfinite(a)) and a.min() >= 0 and a.max() <= 65535 and np.all(a == np.round(a)):
                im = Image.fromarray(a.astype(np.uint16))    # 값이 16비트 안의 정수면 그대로(무손실)
            else:
                im = _8비트(im)
                말 = "32-bit TIFF는 PPTX에 화면용(8-bit)으로 넣었어요(원본은 소재 폴더에 그대로)"
        elif im.mode not in ("1", "L", "LA", "P", "RGB", "RGBA", "I;16"):
            im = im.convert("RGBA" if "A" in im.mode else "RGB")
        buf = io.BytesIO()
        im.save(buf, "PNG")
        buf.seek(0)
        return buf, 말
