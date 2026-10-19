#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
if [[ ! -x .venv/bin/python ]]; then
  printf 'Run ./setup.sh first to create the Python environment.\n' >&2
  exit 1
fi
if [[ ! -f .env ]]; then
  printf 'Missing .env; add your Clockify API key and workspace ID.\n' >&2
  exit 1
fi
since="${1:-2026-01-01}"
shift $(( $# > 0 ? 1 : 0 ))
exec .venv/bin/python sync_clockify_controlling.py --since "$since" "$@"
