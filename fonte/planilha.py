# Gera a planilha modelo a partir dos dados das provas em fonte/app.html.
# Uso: python3 fonte/planilha.py
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import os
AQUI=os.path.dirname(os.path.abspath(__file__))
src=open(os.path.join(AQUI,'app.html'),encoding='utf-8').read()
provas=re.findall(r'P\("([^"]+)","(CBTC|ISA)","([^"]+)","([^"]+)",(\d+),(\d+),(null|\d+),"[^"]*",\[[^\]]*\],"([^"]+)"',src)
RANK={"hf":"Hit factor","pontos":"Pontos","?":"A confirmar"}
ORD={"CBTC":["TRP","4 Cores","Humanoide"],"ISA":["B-27","4x2","Ring of Fire"]}
provas.sort(key=lambda p:(0 if p[1]=="CBTC" else 1, ORD[p[1]].index(p[3]), int(p[4])))
COLS=re.search(r'const COLS = \[(.*?)\];',src).group(1); COLS=re.findall(r'"([^"]+)"',COLS)
assert len(COLS)==22, COLS
N=2001  # linhas preparadas

F="Arial"; NAVY="1B2A6B"
hfont=Font(name=F,bold=True,color="FFFFFF"); base=Font(name=F); bold=Font(name=F,bold=True)
hfill=PatternFill("solid",fgColor=NAVY); infill=PatternFill("solid",fgColor="FFF59D"); cinza=PatternFill("solid",fgColor="EEF0EC")
thin=Border(bottom=Side(style="thin",color="D9DDD6"))
wb=Workbook()
h=wb.active; h.title="Histórico"
h.append(COLS+["Vale"])
for c in h[1]: c.font=hfont; c.fill=hfill
widths=[16,11,9,9,34,11,22,16,11,10,22,8,10,9,6,6,10,30,16,9,10,15,6]
for i,w in enumerate(widths,1): h.column_dimensions[get_column_letter(i)].width=w
for r in range(2,N+1):
    h.cell(r,3).number_format="@"; h.cell(r,2).number_format="DD/MM/YYYY"
    h.cell(r,13).number_format="0.00"; h.cell(r,17).number_format="0.0000"; h.cell(r,22).number_format="0"
    # Vale = 1 só na versão mais recente de cada ID, uma vez, e se não foi apagada
    h.cell(r,23,f'=IF(A{r}="","",IF(OR(T{r}="Apagada",T{r}="Teste",COUNTIFS($A$2:$A${N},A{r},$V$2:$V${N},">"&V{r})>0,COUNTIFS($A$2:A{r},A{r},$V$2:V{r},V{r})>1),0,1))')
    h.cell(r,23).fill=cinza
h.freeze_panes="A2"; h.auto_filter.ref="A1:W1"
h.cell(1,23).comment=None

R="Histórico!"
rng=lambda col:f"{R}${col}$2:${col}${N}"
s=wb.create_sheet("Resumo")
s["A1"]="Resumo da etapa"; s["A1"].font=Font(name=F,bold=True,size=14)
s["A2"]="Etapa (aaaa-mm)"; s["B2"]="2026-10"; s["B2"].number_format="@"
s["A3"]="Categoria (* = todas)"; s["B3"]="*"
for c in ("B2","B3"): s[c].fill=infill; s[c].font=bold
s["D2"]="Edite as células amarelas. Etapa no formato 2026-10. Categoria igual à do app (ex.: Curta restrita) ou * para todas."
s["D2"].font=Font(name=F,italic=True,color="5B6670")
hdr=["Federação","Prova","Alvo","Disparos","Ranking","Passadas","Feita","Lançada na CBTC","Melhor pontos","Melhor hit factor","Menor tempo (s)"]
s.append([]); s.append(hdr)
for c in s[5]: c.font=hfont; c.fill=hfill
crit=lambda r:f'{rng("E")},$B{r},{rng("C")},$B$2,{rng("G")},$B$3,{rng("W")},1'
for i,p in enumerate(provas):
    r=6+i
    s.cell(r,1,p[1]); s.cell(r,2,p[2]); s.cell(r,3,p[3]); s.cell(r,4,int(p[4])); s.cell(r,5,RANK[p[7]])
    s.cell(r,6,f'=COUNTIFS({crit(r)})')
    s.cell(r,7,f'=IF(F{r}>0,"Sim","")')
    s.cell(r,8,f'=IF(F{r}=0,"",IF(COUNTIFS({crit(r)},{rng("U")},"Sim")>0,"Sim",IF(A{r}="CBTC","Falta","")))')
    s.cell(r,9,f'=IF(F{r}=0,"",_xlfn.MAXIFS({rng("L")},{crit(r)}))')
    s.cell(r,10,f'=IF(F{r}=0,"",IF(_xlfn.MAXIFS({rng("Q")},{crit(r)})=0,"",_xlfn.MAXIFS({rng("Q")},{crit(r)})))')
    s.cell(r,11,f'=IF(F{r}=0,"",_xlfn.MINIFS({rng("M")},{crit(r)}))')
    s.cell(r,10).number_format="0.0000"; s.cell(r,11).number_format="0.00"
    for c in range(1,12): s.cell(r,c).font=base; s.cell(r,c).border=thin
last=6+len(provas)-1
s.cell(last+2,1,"Provas feitas (CBTC)").font=bold
s.cell(last+2,6,f'=COUNTIFS(A6:A{last},"CBTC",F6:F{last},">0")').font=bold
s.cell(last+3,1,"Treinos lançados na etapa").font=bold
s.cell(last+3,6,f'=COUNTIFS({rng("T")},"Treino",{rng("C")},$B$2,{rng("G")},$B$3,{rng("W")},1)').font=bold
for col,w in zip("ABCDEFGHIJK",[11,36,13,9,14,10,7,15,14,16,15]): s.column_dimensions[col].width=w
s.freeze_panes="C6"

e=wb.create_sheet("Evolução")
e["A1"]="Evolução mês a mês"; e["A1"].font=Font(name=F,bold=True,size=14)
e["A2"]="Melhor resultado de cada etapa: hit factor nas provas por hit factor, pontos nas demais. Troque as etapas da linha 4 se quiser outro período."
e["A2"].font=Font(name=F,italic=True,color="5B6670")
e["A3"]="Categoria (* = todas)"; e["C3"]="*"; e["C3"].fill=infill; e["C3"].font=bold
meses=[f"2026-{m:02d}" for m in range(1,13)]
e.cell(4,1,"Prova"); e.cell(4,2,"Mede")
for j,m in enumerate(meses):
    c=e.cell(4,3+j,m); c.number_format="@"
for c in e[4]: c.font=hfont; c.fill=hfill
for i,p in enumerate(provas):
    r=5+i; e.cell(r,1,f"{p[2]} ({p[1]})"); e.cell(r,2,f'=IF(Resumo!E{6+i}="Hit factor","HF","Pontos")')
    for j in range(12):
        col=get_column_letter(3+j)
        cr=f'{rng("E")},Resumo!$B{6+i},{rng("C")},{col}$4,{rng("G")},$C$3,{rng("W")},1'
        e.cell(r,3+j,f'=IF(COUNTIFS({cr})=0,"",IF($B{r}="HF",_xlfn.MAXIFS({rng("Q")},{cr}),_xlfn.MAXIFS({rng("L")},{cr})))')
        e.cell(r,3+j).number_format="0.00"
    for c in range(1,15): e.cell(r,c).font=base; e.cell(r,c).border=thin
e.column_dimensions["A"].width=42; e.column_dimensions["B"].width=8
for j in range(12): e.column_dimensions[get_column_letter(3+j)].width=9
e.freeze_panes="C5"

ins=wb.create_sheet("Instruções")
linhas=[("Caderneta CBTC: planilha de controle (versão 2.0)",True),("",False),
("COMO LANÇAR (padrão: manual)",True),
("1. No app, vá em Mais → Exportar e toque em 'Copiar para colar na planilha'. Por padrão vão as passadas e treinos novos, os editados e os apagados desde a última exportação.",False),
("2. Aqui, na aba Histórico, clique na primeira linha vazia, coluna A, e cole. Cada valor cai na sua coluna.",False),
("3. Passada editada ou apagada entra de novo como linha nova. A coluna Vale (cinza) marca 1 só na versão mais recente de cada passada; Resumo e Evolução contam só essas.",False),
("4. As abas Resumo e Evolução se atualizam sozinhas. Na Resumo, troque a etapa e a categoria nas células amarelas.",False),
("",False),
("ENVIO AUTOMÁTICO (opcional, só Google Planilhas)",True),
("1. Envie esta planilha para o Google Drive e abra com o Google Planilhas.",False),
("2. Extensões > Apps Script. Apague o que estiver lá e cole o código que o app mostra em Mais → Envio automático (botão 'Copiar código do script').",False),
("3. Implantar > Nova implantação > App da Web. Executar como: eu. Quem pode acessar: qualquer pessoa. Autorize.",False),
("4. Copie o endereço que termina em /exec e cole no app. Toque em Testar: o app confirma quando a planilha recebeu.",False),
("Com o envio automático, a passada editada ou apagada atualiza a própria linha, sem criar linha nova.",False),
("Já usava a versão anterior do script? Cole o código novo e use Implantar > Gerenciar implantações > editar > Nova versão. O endereço continua o mesmo.",False),
("",False),
("VINDO DA PLANILHA VERSÃO 1",True),
("A versão 2.0 tem 3 colunas novas no fim (Tipo, Lançada na CBTC, Versão) e a coluna Vale. Comece esta planilha nova e, no app, exporte 'Tudo' uma vez para preenchê-la.",False),
("",False),
("CUIDADOS",True),
("Não mude a ordem das colunas da aba Histórico: o app e as fórmulas dependem dela.",False),
("A coluna Etapa está formatada como texto (ex.: 2026-10). Mantenha assim.",False),
(f"A coluna Vale vem pronta até a linha {N}. Passando disso, arraste a fórmula para baixo.",False),
("A contagem de habitualidade fica no próprio app, na aba Habitualidade.",False)]
for i,(t,b) in enumerate(linhas,1):
    c=ins.cell(i,1,t); c.font=Font(name=F,bold=b,size=13 if i==1 else 11); c.alignment=Alignment(wrap_text=True,vertical="top")
ins.column_dimensions["A"].width=110
wb.move_sheet("Instruções",-3); wb.active=0
wb.save(os.path.join(AQUI,"Caderneta_CBTC_planilha.xlsx"))
print(len(provas),"provas")
