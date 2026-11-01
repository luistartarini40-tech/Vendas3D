#!/usr/bin/env bash
# Recompila o Tailwind e embute o CSS dentro do <style id="tailwind"> do index.html.
# Uso (na raiz do repositório): bash tools/build-css.sh
set -euo pipefail
cd "$(dirname "$0")/.."
printf '@tailwind base;\n@tailwind components;\n@tailwind utilities;\n' > /tmp/tw-input.css
${TAILWIND_BIN:-npx --yes tailwindcss@3} -c tools/tailwind.config.js -i /tmp/tw-input.css -o /tmp/tw-output.css --minify
python3 - <<'PY'
import re
h = open('index.html', encoding='utf-8').read()
css = open('/tmp/tw-output.css', encoding='utf-8').read().strip()
h = re.sub(r'<style id="tailwind">.*?</style>', lambda m: '<style id="tailwind">' + css + '</style>', h, count=1, flags=re.S)
open('index.html', 'w', encoding='utf-8').write(h)
print('CSS embutido:', len(css), 'bytes')
PY
