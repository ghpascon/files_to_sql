#!/usr/bin/env bash
set -euo pipefail

poetry run python -m ruff check --fix
poetry run python -m ruff format
poetry run python scripts/commit.py
