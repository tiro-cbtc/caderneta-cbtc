# Gera o site (index.html, sw.js, manifest, ícones) a partir de fonte/app.html.
# Uso: python3 fonte/gerar.py   (antes, aumente VERSION abaixo para os aparelhos pegarem a versão nova)
import os, re, shutil, json
from PIL import Image, ImageDraw

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # raiz do repositório
src = open(f"{ROOT}/fonte/app.html", encoding="utf-8").read()

# 1) Versão artifact: o próprio arquivo
open(f"{ROOT}/fonte/artifact.html", "w", encoding="utf-8").write(src)  # versão publicada no Claude

# 2) Versão PWA
out = ROOT  # o site é a raiz do repositório (GitHub Pages)
os.makedirs(f"{out}/icons", exist_ok=True)  # vendor/xlsx.full.min.js (SheetJS 0.18.5) já fica no repositório

title = re.search(r"<title>.*?</title>", src).group(0)
body = src.replace(title, "", 1).replace('<meta name="claude:viewport-fit-sides" content="cover">\n', "")
body = body.replace("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js", "vendor/xlsx.full.min.js")
# separa <link>/<style> para o head
head_bits = re.findall(r'<link[^>]*>\n?', body)
style = re.search(r"<style>.*?</style>", body, re.S).group(0)
for h in head_bits: body = body.replace(h, "", 1)
body = body.replace(style, "", 1)
reset = "*,*::before,*::after{box-sizing:border-box}:root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0}img{max-width:100%}[hidden]{display:none!important}"
html = f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
{title}
<meta name="description" content="Provas da CBTC por tipo de alvo, lançamento de resultados, checklist de material e exportação para planilha.">
<meta name="theme-color" content="#1b2a6b">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Caderneta">
<link rel="manifest" href="manifest.webmanifest">
<link rel="icon" href="icons/icon-192.png">
<link rel="apple-touch-icon" href="icons/apple-touch-icon.png">
{''.join(head_bits)}<style>{reset}</style>
{style}
</head>
<body>
{body.strip()}
</body>
</html>
"""
open(f"{out}/index.html", "w", encoding="utf-8").write(html)

manifest = {
  "name": "Caderneta CBTC", "short_name": "Caderneta", "lang": "pt-BR",
  "description": "Provas da CBTC, resultados e material da etapa",
  "start_url": "./", "scope": "./", "display": "standalone",
  "background_color": "#f3f4f1", "theme_color": "#1b2a6b",
  "icons": [
    {"src": "icons/icon-192.png", "sizes": "192x192", "type": "image/png"},
    {"src": "icons/icon-512.png", "sizes": "512x512", "type": "image/png"},
    {"src": "icons/icon-maskable-512.png", "sizes": "512x512", "type": "image/png", "purpose": "maskable"}
  ]
}
json.dump(manifest, open(f"{out}/manifest.webmanifest", "w", encoding="utf-8"), ensure_ascii=False, indent=2)

VERSION = "v8-2026-11-01"
sw = f"""// Caderneta CBTC: funciona sem internet depois da primeira abertura.
const CACHE = 'caderneta-{VERSION}';
const ARQUIVOS = ['./', './index.html', './manifest.webmanifest', './vendor/xlsx.full.min.js',
  './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png'];
self.addEventListener('install', e => {{ e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARQUIVOS)).then(() => self.skipWaiting())); }});
self.addEventListener('activate', e => {{ e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); }});
self.addEventListener('fetch', e => {{
  const req = e.request;
  if (req.method !== 'GET') return;            // envio para o Google Planilhas passa direto
  const url = new URL(req.url);
  if (url.origin === location.origin) {{
    // rede primeiro (pega atualização), cache se estiver sem sinal
    e.respondWith(fetch(req).then(r => {{ const c = r.clone(); caches.open(CACHE).then(k => k.put(req, c)); return r; }})
      .catch(() => caches.match(req).then(r => r || caches.match('./index.html'))));
  }} else if (/fonts\\.(googleapis|gstatic)\\.com$/.test(url.hostname)) {{
    e.respondWith(caches.match(req).then(r => r || fetch(req).then(n => {{ const c = n.clone(); caches.open(CACHE).then(k => k.put(req, c)); return n; }})));
  }}
}});
"""
open(f"{out}/sw.js", "w", encoding="utf-8").write(sw)

# Ícones: alvo com anéis, fundo azul CBTC
def icone(size, pad=0.0, path=None):
    S = size * 4
    img = Image.new("RGB", (S, S), "#1b2a6b")
    d = ImageDraw.Draw(img)
    c = S / 2; R = S * (0.40 - pad)
    cores = ["#f3f4f1", "#1b2a6b", "#f3f4f1", "#1b2a6b", "#0f7a3d"]
    for i, col in enumerate(cores):
        r = R * (1 - i * 0.19)
        d.ellipse([c - r, c - r, c + r, c + r], fill=col)
    r = R * 0.12
    d.ellipse([c - r, c - r, c + r, c + r], fill="#e0b04f")
    img = img.resize((size, size), Image.LANCZOS)
    img.save(path)
icone(192, path=f"{out}/icons/icon-192.png")
icone(512, path=f"{out}/icons/icon-512.png")
icone(512, pad=0.08, path=f"{out}/icons/icon-maskable-512.png")
icone(180, path=f"{out}/icons/apple-touch-icon.png")

print("ok")
