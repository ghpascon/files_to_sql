#!/usr/bin/env bash
set -euo pipefail

read -r -p "[WARN] This will reset your local branch to match remote. Continue? [y/N] " CONFIRM
case "$CONFIRM" in
  [yY][eE][sS]|[yY]) ;;
  *) echo "[ABORTED] No changes made."; exit 0 ;;
esac

echo "[INFO] Fetching remote..."
git fetch origin

echo "[INFO] Resetting local branch to match remote..."
BRANCH="$(git rev-parse --abbrev-ref HEAD)"
git reset --hard "origin/${BRANCH}"

echo "[INFO] Cleaning untracked files and directories..."
git clean -fd

echo "[DONE] Local repo is now in sync with origin/${BRANCH}"
