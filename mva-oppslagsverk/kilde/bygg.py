"""Setter sammen de fem HTML-filene: tema + felles CSS/JS + data + variant. Kjør: python -I bygg.py"""
import json, pathlib
here = pathlib.Path(__file__).parent
root = here.parent
_d = json.load(open(root/'innhold.json', encoding='utf8'))
for _s in _d['scenarioer']:
    _s.pop('fotnote_intern', None); _s.pop('intern_merknad', None)   # interne notater skal ikke ligge i filen
data = json.dumps(_d, ensure_ascii=False).replace('</', '<\/')
tema = (root/'bauda-tema.css').read_text(encoding='utf8')
cj = (here/'common.js').read_text(encoding='utf8')
# --- Ny hovedfil: MVA-oppslagsverk.html (beslutningstre som app-skjerm) ---
pr = (here/'print.css').read_text(encoding='utf8')
v = (here/'app.del').read_text(encoding='utf8')
css, body, js = v.split('<!--CSS-->')[1].split('<!--BODY-->')[0], v.split('<!--BODY-->')[1].split('<!--JS-->')[0], v.split('<!--JS-->')[1]
html = f'''<!doctype html>
<html lang="nb"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>MVA-oppslagsverk – Bauda</title>
<style>{tema}
{css}
{pr}</style></head>
<body>
{body}
<script>window.DATA={data};</script>
<script>{cj}</script>
<script>{js}</script>
</body></html>'''
import base64
bil = 'data:image/jpeg;base64,' + base64.b64encode((root/'bilder'/'bil-header.jpg').read_bytes()).decode()
html = html.replace('%%BIL%%', bil)
import re as _re
svg = (root/'bilder'/'bdo-logo.svg').read_text(encoding='utf8')
svg = _re.sub(r' xmlns:xlink="[^"]*"| id="[^"]*"| overflow="[^"]*"', '', svg).replace('<defs></defs>', '').replace('<svg ', '<svg aria-hidden="true" focusable="false" ', 1)
_farger = iter(['#E3243A', '#1F3A93'])   # BDO: rød strek, blå bokstaver
svg = _re.sub(r'fill="#FFFFFF"', lambda m: 'fill="' + next(_farger) + '"', svg)
html = html.replace('%%BDO%%', svg)
(root/'MVA-oppslagsverk.html').write_text(html, encoding='utf8')
print('skrev MVA-oppslagsverk', len(html)//1024, 'kB')

# --- Redigeringsverktøy: Rediger.html (malen er hovedfilen med data byttet ut) ---
mal = html.replace('window.DATA=' + data + ';', 'window.DATA=__DATA__;')
assert '__DATA__' in mal
ed = (here/'rediger.html').read_text(encoding='utf8')
ed = ed.split('%%MAL%%')
ed = (json.dumps(mal, ensure_ascii=False).replace('</', '<\\/')).join(ed)
ed = ed.replace('%%DATA%%', data)
(root/'Rediger.html').write_text(ed, encoding='utf8')
print('skrev Rediger', len(ed)//1024, 'kB')
