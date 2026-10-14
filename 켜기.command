#!/bin/bash
# Mac: 처음이면(또는 설치가 덜 끝났으면) 가상환경을 만들고, 덱편집기를 켠다
cd "$(dirname "$0")" || exit 1
if [ ! -f .venv/ready.txt ]; then
  { [ -x .venv/bin/python ] || python3 -m venv .venv; } \
    && .venv/bin/python -m pip install -r requirements.txt \
    && .venv/bin/python -c "import pptx, PIL, pypdfium2, openpyxl" \
    && echo ok > .venv/ready.txt \
    || { read -r -p "준비하지 못했어요(인터넷·파이썬 3.10 이상 확인). Enter"; exit 1; }
fi
.venv/bin/python 서버.py "$@"
