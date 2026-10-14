#!/usr/bin/env bash
# Пересобирает карту проекта Graphify при старте сессии (только код и Markdown, без LLM, ~1 с).
# Карта не коммитится: в облачной сессии git-хуки Graphify не сохраняются, а закоммиченная карта устаревает.
set -u
cd "${CLAUDE_PROJECT_DIR:-$(pwd)}" || exit 0
if ! command -v graphify >/dev/null 2>&1; then
  pip install --quiet "graphifyy==0.9.90" >/dev/null 2>&1 || { echo "graphify: установка не удалась, работаю без карты"; exit 0; }
fi
graphify update . >/dev/null 2>&1 && echo "graphify: карта проекта обновлена (graphify-out/GRAPH_REPORT.md)" || echo "graphify: карта не построена"
exit 0
