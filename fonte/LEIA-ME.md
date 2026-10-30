# Código-fonte da Caderneta CBTC

- `app.html`: o app inteiro (estilos, telas, dados das provas e lógica). É o único arquivo que se edita.
- `gerar.py`: gera o site na raiz do repositório (`index.html`, `sw.js`, `manifest.webmanifest`, ícones) e `artifact.html` (versão publicada no Claude). Antes de rodar, aumente `VERSION` para os aparelhos instalados pegarem a atualização.
- `planilha.py`: gera `Caderneta_CBTC_planilha.xlsx`, a planilha modelo, a partir das provas de `app.html`.
- `Caderneta_CBTC_planilha.xlsx`: planilha modelo atual, para mandar ao grupo.

Fluxo: editar `app.html` → `python3 fonte/gerar.py` → `python3 fonte/planilha.py` (se mudaram provas ou colunas) → commit e push. O GitHub Pages publica em https://tiro-cbtc.github.io/caderneta-cbtc/ em um ou dois minutos.

Requisitos: Python 3 com `Pillow` (ícones) e `openpyxl` (planilha). A biblioteca de Excel do app (SheetJS 0.18.5) fica em `vendor/`.

Decisões de conteúdo e pendências: ver o documento de passagem no Projeto "Tiro Desportivo - Campeonatos CBTC" no Claude.
