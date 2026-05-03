#!/usr/bin/env bash
# Simple launcher for non-technical users: generates map and opens it in default browser
set -e
python3 Route.py
sleep 1
latest=$(ls -t /workspaces/RnD/Site_Map_*.html 2>/dev/null | head -n1 || true)
if [ -n "$latest" ]; then
  # Start a simple HTTP server in background so the browser can fetch JSON assets.
  (cd /workspaces/RnD && python3 -m http.server 8000 >/dev/null 2>&1 &) || true
  url="http://localhost:8000/$(basename "$latest")"
  echo "Opening map in browser: $url"
  if command -v xdg-open >/dev/null 2>&1; then
    xdg-open "$url" || echo "Open the map at: $url"
  else
    echo "Open the map at: $url"
  fi
else
  echo "Map file not found. Check script output for errors."
fi
