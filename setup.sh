#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
printf 'Ready. Run ./refresh_clockify.sh to update Excel and PDF.\n'
