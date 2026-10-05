#!/bin/bash
# בדיקות לפני כל שינוי שעולה לאוויר. שימוש: qa/qa_run.sh [רוחב חלון]
# 1) הסקריפט של גוגל מול גיליון מדומה  2) הדף האמיתי מול הסקריפט האמיתי, בכרום ללא ממשק
echo "--- סקריפט גוגל (הרצה יבשה) ---"; node "$(cd "$(dirname "$0")" && pwd)/gs_test.js" || exit 1
echo "--- הדף מקצה לקצה ---"
S="$(cd "$(dirname "$0")" && pwd)"
/opt/homebrew/bin/python3.13 "$S/qa_build.py" >/dev/null || exit 1
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --disable-gpu --window-size=${1:-1280},900 --virtual-time-budget=30000 --dump-dom "file://$S/qa.html" 2>/dev/null > "$S/qa_dom.html"
/opt/homebrew/bin/python3.13 - "$S/qa_dom.html" <<'PY'
import io, sys, re, json, html
d=io.open(sys.argv[1],encoding='utf-8').read()
m=re.search(r'<pre id="qaOut">(.*?)</pre>', d, flags=re.S)
if not m: print('NO RESULT - harness did not finish'); sys.exit(1)
r=json.loads(html.unescape(m.group(1)))
print('width', r['vw'], '|', r['pass'], '/', r['total'], 'passed')
for t in r['results']:
    print(('PASS ' if t['ok'] else 'FAIL ')+str(t['id']).rjust(2)+' '+t['name']+('' if t['ok'] else '\n        -> '+t['detail'][:400]))
PY
