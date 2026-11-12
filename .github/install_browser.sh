#!/usr/bin/env bash
# Installs Playwright's Chromium for the bots (maya-*, market-data, funds-*) WITHOUT apt.
#
# Why: `playwright install --with-deps chromium` runs apt against the Ubuntu mirrors on every run, only to add
# a few font packages the bots do not need (they call JSON APIs inside the page; nothing is rendered for output).
# On 11.11.2026 the runner could not resolve the Ubuntu mirrors for ~20 minutes and both daily maya-watch runs
# failed in this step before the watcher even started.
#
# Now: download the browser, try to start it, and call apt (`playwright install-deps`) only if it cannot start.
set -euo pipefail

browser_starts() {
  python - <<'EOF'
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    b.new_page().goto("about:blank")
    b.close()
EOF
}

python -m playwright install chromium

if browser_starts; then
  echo "install_browser: Chromium starts — no system packages needed"
  exit 0
fi

echo "install_browser: Chromium did not start — installing system dependencies (apt)"
python -m playwright install-deps chromium
browser_starts
echo "install_browser: Chromium starts after installing system dependencies"
