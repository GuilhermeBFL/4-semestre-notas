"""
4-SEMESTRE-NOTAS.xlsx — planilha unificada com 7 matérias + dashboard Início
"""
import os, openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

OUT_PATH = r"C:\Users\guilh\Downloads\4 SEMESTRE\4-SEMESTRE-NOTAS.xlsx"

# ── Estilos ────────────────────────────────────────────────────────────
def fill(h): return PatternFill(start_color=h, end_color=h, fill_type="solid")
F_DB=fill("1F3864"); F_MB=fill("2E75B6"); F_LB=fill("EBF3FB"); F_GL=fill("EDEDED")
F_YI=fill("FFFF99"); F_GC=fill("D9D9D9"); F_GD=fill("FFD700"); F_DG=fill("595959")
F_BW=fill("833C00"); F_CH=fill("BDD7EE"); F_GN=fill("C6EFCE"); F_RD=fill("FFC7CE")
F_OR=fill("FCE4D6"); F_PU=fill("E2EFDA")

def fnt(bold=False,color="000000",size=11,italic=False):
    return Font(name="Calibri",bold=bold,color=color,size=size,italic=italic)
AC=Alignment(horizontal="center",vertical="center",wrap_text=True)
AL=Alignment(horizontal="left",  vertical="center",wrap_text=True)

def tb():
    s=Side(style="thin"); return Border(left=s,right=s,top=s,bottom=s)

def sc(ws,ref,val,f=None,fn=None,al=None,brd=True,fmt=None):
    c=ws[ref]; c.value=val
    if f:   c.fill=f
    if fn:  c.font=fn
    if al:  c.alignment=al
    if brd: c.border=tb()
    if fmt: c.number_format=fmt
    return c

def mg(ws,rng,val,f=None,fn=None,al=None,h=None,r=None):
    ws.merge_cells(rng); top=rng.split(":")[0]; c=ws[top]
    c.value=val
    if f:  c.fill=f
    if fn: c.font=fn
    c.alignment=al or AC
    if h and r: ws.row_dimensions[r].height=h

def std_cols(ws):
    ws.column_dimensions["A"].width=38; ws.column_dimensions["B"].width=14
    ws.column_dimensions["C"].width=16; ws.column_dimensions["D"].width=30
    ws.column_dimensions["F"].hidden=True; ws.column_dimensions["F"].width=18
    ws.sheet_view.showGridLines=False

def header4(ws,title,sub,formula,approval):
    mg(ws,"A1:D1",title,   F_DB,fnt(bold=True,color="FFFFFF",size=14),AC,30,1)
    mg(ws,"A2:D2",sub,     F_MB,fnt(color="FFFFFF",size=10),AC)
    mg(ws,"A3:D3",formula, F_LB,fnt(italic=True,size=10),AC)
    mg(ws,"A4:D4",approval,F_GL,fnt(size=10),AC)
    ws.row_dimensions[5].height=6

def inp_header(ws,row):
    mg(ws,f"A{row}:D{row}","▼  INSIRA SUAS NOTAS  (células amarelas)",
       F_GD,fnt(bold=True,size=11),AC,22,row)
    for col,txt in [("A","Avaliação"),("B","Data"),("C","Nota"),("D","Observações")]:
        sc(ws,f"{col}{row+1}",txt,F_CH,fnt(bold=True),AC)
    ws.row_dimensions[row+1].height=18

def calc_hdr(ws,row):
    mg(ws,f"A{row}:D{row}","▼  CALCULADO AUTOMATICAMENTE  (não editar)",
       F_DG,fnt(bold=True,color="FFFFFF",size=11),AC,22,row)

def sim_hdr(ws,row):
    mg(ws,f"A{row}:D{row}","▼  SIMULAÇÃO — O QUE PRECISO PARA PASSAR?",
       F_BW,fnt(bold=True,color="FFFFFF",size=11),AC,22,row)

def input_row(ws,row,label,date,obs,row_h=18):
    sc(ws,f"A{row}",label,fn=fnt(),al=AL)
    sc(ws,f"B{row}",date,fn=fnt(size=10),al=AC)
    sc(ws,f"C{row}",None,F_YI,fn=fnt(),al=AC)
    sc(ws,f"D{row}",obs,fn=fnt(italic=True,size=10),al=AL)
    ws.row_dimensions[row].height=row_h

def calc_row(ws,row,label,formula,obs="",fmt="0.00"):
    sc(ws,f"A{row}",label,fn=fnt(bold=True),al=AL)
    sc(ws,f"C{row}",formula,F_GC,fn=fnt(),al=AC,fmt=fmt if fmt else None)
    sc(ws,f"D{row}",obs,fn=fnt(italic=True,size=10),al=AL)
    ws.row_dimensions[row].height=20

def spacer(ws,row,h=6):
    ws.row_dimensions[row].height=h

def add_cf(ws,cell,green=None,red=None):
    if green: ws.conditional_formatting.add(cell,FormulaRule(
        formula=[f'NOT(ISERROR(SEARCH("✓",{cell})))'],fill=F_GN))
    if red:   ws.conditional_formatting.add(cell,FormulaRule(
        formula=[f'NOT(ISERROR(SEARCH("✗",{cell})))'],fill=F_RD))

# ════════════════════════════════════════════════════════════════════════
# 1. ECONOMETRIA
# MF_cell = C18   (row for referencing from Início)
# ════════════════════════════════════════════════════════════════════════
def build_econometria(wb):
    ws=wb.create_sheet("Econometria"); std_cols(ws)
    header4(ws,
        "ECONOMETRIA — Controle de Notas",
        "Prof. Adriana Bortoluzzo / Sérgio Martins  |  Turma 4ECOB  |  Qua 16:30–18:30 / Sex 14:15–16:15",
        "MP=0,40·AI+0,60·AF  |  MF=MP se MP<4  |  MF=0,75·MP+0,15·MQ+0,10·APS se MP≥4  |  MQ=2 melhores de 3 quizzes",
        "Aprovação: MF≥5,0  |  APS zero = REPROVAÇÃO AUTOMÁTICA")
    inp_header(ws,6)
    rows_in=[
        (8, "Avaliação Intermediária (AI) — 30% (via MP)","25/03/2026",""),
        (9, "Quiz 1 — para MQ","11/03/2026",""),
        (10,"Quiz 2 — para MQ","20/03/2026",""),
        (11,"Quiz 3 — para MQ","08/05/2026","Descarta pior → média dos 2 melhores"),
        (12,"APS — Atividade em Aula (10%)","ao longo","Nota zero = reprovação"),
        (13,"Avaliação Final (AF) — 45% (via MP)","27/05/2026",""),
    ]
    for row,label,date,obs in rows_in:
        input_row(ws,row,label,date,obs)
    spacer(ws,14)
    calc_hdr(ws,15)
    calc_row(ws,16,"MQ (2 melhores quizzes de 3)",
        '=IF(COUNT(C9:C11)=0,"—",IF(COUNT(C9:C11)<3,AVERAGE(C9:C11),(LARGE(C9:C11,1)+LARGE(C9:C11,2))/2))',
        "Descarta pior quiz dos 3")
    calc_row(ws,17,"MP (Média das Provas)",
        '=IF(OR(NOT(ISNUMBER(C8)),NOT(ISNUMBER(C13))),"",0.4*C8+0.6*C13)',
        "0,40·AI + 0,60·AF")
    calc_row(ws,18,"MF (Média Final)",
        '=IF(C17="","",IF(C17<4,C17,0.75*C17+0.15*IF(ISNUMBER(C16),C16,0)+0.1*IF(ISNUMBER(C12),C12,0)))',
        "Fórmula condicional (MP≥4)")
    calc_row(ws,19,"Situação atual",
        '=IF(C18="","Aguardando notas...",IF(C18>=5,"✓ APROVADO DIRETO!","Em andamento — veja simulação"))',
        "",fmt=None)
    add_cf(ws,"C19",green=True,red=False)
    spacer(ws,20)
    sim_hdr(ws,21)
    ws["F22"].value='=IF(NOT(ISNUMBER(C8)),-999,MAX((5-0.3*C8-0.15*IF(ISNUMBER(C16),C16,0)-0.1*IF(ISNUMBER(C12),C12,0))/0.45,(4-0.4*C8)/0.6))'
    sim_f='=IF(NOT(ISNUMBER(C8)),"Preencha a AI primeiro",IF(F22<=0,"✓ Aprovado direto!",IF(F22>10,"✗ Reprovado mesmo com 10 na AF","Precisa de "&TEXT(ROUND(F22,2),"0.00")&" na AF")))'
    sc(ws,"A22","Nota necessária na AF (dados atuais)",fn=fnt(bold=True),al=AL)
    sc(ws,"B22","Brancos = 0 (pior caso)",fn=fnt(italic=True,size=9),al=AC)
    sc(ws,"C22",sim_f,F_OR,fn=fnt(bold=True),al=AC)
    ws.row_dimensions[22].height=22
    add_cf(ws,"C22",green=True,red=True)
    ws.freeze_panes="A8"
    return "C18"  # MF cell

# ════════════════════════════════════════════════════════════════════════
# 2. ESTATÍSTICA II
# MF_cell = C16
# ════════════════════════════════════════════════════════════════════════
def build_estatistica(wb):
    ws=wb.create_sheet("Estatística II"); std_cols(ws)
    header4(ws,
        "ESTATÍSTICA II — Controle de Notas",
        "Prof. Maria Kelly Venezuela / Rinaldo Artes  |  Turma 3DPA  |  Qua 19:00–21:00 / Sex 19:00–21:00",
        "MP=0,35·PI+0,65·PF  |  MA=0,50·APS1+0,50·APS2  |  MF=MP(MP<4) ou 0,80·MP+0,20·MA(MP≥4)",
        "Aprovação: MF≥5,0  |  Não entregar NENHUMA APS = REPROVAÇÃO AUTOMÁTICA")
    inp_header(ws,6)
    rows_in=[
        (8, "PI — Prova Intermediária (35% do MP)","25/03/2026",""),
        (9, "APS 1 — 50% da MA","13/03/2026","Realizada em sala, individual"),
        (10,"APS 2 — 50% da MA","29/04/2026","Realizada em sala, individual"),
        (11,"PF — Prova Final (65% do MP)","27/05/2026",""),
    ]
    for row,label,date,obs in rows_in:
        input_row(ws,row,label,date,obs)
    spacer(ws,12)
    calc_hdr(ws,13)
    calc_row(ws,14,"MA (Média das APS)",
        '=IF(AND(ISNUMBER(C9),ISNUMBER(C10)),0.5*C9+0.5*C10,IF(ISNUMBER(C9),C9,IF(ISNUMBER(C10),C10,"—")))',
        "0,50·APS1+0,50·APS2")
    calc_row(ws,15,"MP (Média das Provas)",
        '=IF(OR(NOT(ISNUMBER(C8)),NOT(ISNUMBER(C11))),"",0.35*C8+0.65*C11)',
        "0,35·PI+0,65·PF")
    calc_row(ws,16,"MF (Média Final)",
        '=IF(C15="","",IF(C15<4,C15,0.8*C15+0.2*IF(ISNUMBER(C14),C14,0)))',
        "Condicional em MP≥4")
    calc_row(ws,17,"Situação atual",
        '=IF(C16="","Aguardando notas...",IF(C16>=5,"✓ APROVADO DIRETO!","Em andamento — veja simulação"))',
        "",fmt=None)
    add_cf(ws,"C17",green=True,red=False)
    spacer(ws,18)
    sim_hdr(ws,19)
    ws["F20"].value='=IF(NOT(ISNUMBER(C8)),-999,MAX((5-0.28*C8-0.2*IF(ISNUMBER(C14),C14,0))/0.52,(4-0.35*C8)/0.65))'
    sim_f='=IF(NOT(ISNUMBER(C8)),"Preencha a PI primeiro",IF(F20<=0,"✓ Aprovado direto!",IF(F20>10,"✗ Reprovado mesmo com 10 na PF","Precisa de "&TEXT(ROUND(F20,2),"0.00")&" na PF")))'
    sc(ws,"A20","Nota necessária na PF (dados atuais)",fn=fnt(bold=True),al=AL)
    sc(ws,"B20","Brancos = 0 (pior caso)",fn=fnt(italic=True,size=9),al=AC)
    sc(ws,"C20",sim_f,F_OR,fn=fnt(bold=True),al=AC)
    ws.row_dimensions[20].height=22
    add_cf(ws,"C20",green=True,red=True)
    ws.freeze_panes="A8"
    return "C16"

# ════════════════════════════════════════════════════════════════════════
# 3. FINANÇAS II
# PI 30%, Q 20% (5 quizzes, descarta pior), PF 40%, APS 10%
# MF = 0.30*PI + 0.20*MQ + 0.40*PF + 0.10*APS
# MF_cell = C20
# ════════════════════════════════════════════════════════════════════════
def build_financas(wb):
    ws=wb.create_sheet("Finanças II"); std_cols(ws)
    header4(ws,
        "FINANÇAS II — Controle de Notas",
        "Prof. Gustavo Soares / Paulo Neto  |  Turma 4ECOB  |  Ter 14:15–16:15 / Sex 09:45–11:45",
        "MF = 0,30·PI + 0,20·MQ + 0,40·PF + 0,10·APS  |  MQ = média dos 4 melhores quizzes (de 5)",
        "Aprovação: MF≥5,0  |  APS zero = REPROVAÇÃO AUTOMÁTICA")
    inp_header(ws,6)
    rows_in=[
        (8,  "PI — Prova Intermediária (30%)","24/02/2026",""),
        (9,  "Quiz 1 (para MQ)","24/02/2026",""),
        (10, "Quiz 2 (para MQ)","10/03/2026",""),
        (11, "Quiz 3 (para MQ)","07/04/2026",""),
        (12, "Quiz 4 (para MQ)","05/05/2026",""),
        (13, "Quiz 5 (para MQ)","19/05/2026","Descarta o pior dos 5"),
        (14, "APS — Ativ. Pedagógica (10%)","ao longo","Nota zero = reprovação"),
        (15, "PF — Prova Final (40%)","26/05/2026",""),
    ]
    for row,label,date,obs in rows_in:
        input_row(ws,row,label,date,obs)
    spacer(ws,16)
    calc_hdr(ws,17)
    # MQ: média dos 4 melhores de 5 = (SUM - MIN) / 4, se COUNT>=1
    calc_row(ws,18,"MQ (Média dos 4 melhores quizzes)",
        '=IF(COUNT(C9:C13)=0,"—",IF(COUNT(C9:C13)=1,AVERAGE(C9:C13),(SUM(C9:C13)-MIN(C9:C13))/MAX(COUNT(C9:C13)-1,1)))',
        "Descarta 1 pior dos 5")
    calc_row(ws,19,"MF (Média Final)",
        '=IF(OR(NOT(ISNUMBER(C8)),NOT(ISNUMBER(C15))),"",0.3*C8+0.2*IF(ISNUMBER(C18),C18,0)+0.4*C15+0.1*IF(ISNUMBER(C14),C14,0))',
        "0,30·PI+0,20·MQ+0,40·PF+0,10·APS")
    calc_row(ws,20,"Situação atual",
        '=IF(C19="","Aguardando notas...",IF(C19>=5,"✓ APROVADO DIRETO!","Em andamento — veja simulação"))',
        "",fmt=None)
    add_cf(ws,"C20",green=True,red=False)
    spacer(ws,21)
    sim_hdr(ws,22)
    # PF needed: 0.3*PI + 0.2*MQ + 0.4*PF + 0.1*APS >= 5
    # PF >= (5 - 0.3*PI - 0.2*MQ - 0.1*APS) / 0.4
    ws["F23"].value='=IF(NOT(ISNUMBER(C8)),-999,(5-0.3*C8-0.2*IF(ISNUMBER(C18),C18,0)-0.1*IF(ISNUMBER(C14),C14,0))/0.4)'
    sim_f='=IF(NOT(ISNUMBER(C8)),"Preencha a PI primeiro",IF(F23<=0,"✓ Aprovado direto!",IF(F23>10,"✗ Reprovado mesmo com 10 na PF","Precisa de "&TEXT(ROUND(F23,2),"0.00")&" na PF")))'
    sc(ws,"A23","Nota necessária na PF (dados atuais)",fn=fnt(bold=True),al=AL)
    sc(ws,"B23","Brancos = 0 (pior caso)",fn=fnt(italic=True,size=9),al=AC)
    sc(ws,"C23",sim_f,F_OR,fn=fnt(bold=True),al=AC)
    ws.row_dimensions[23].height=22
    add_cf(ws,"C23",green=True,red=True)
    ws.freeze_panes="A8"
    return "C19"

# ════════════════════════════════════════════════════════════════════════
# 4. HPE — História do Pensamento Econômico
# PI 25%, PF 35%, APS 15%, AA 20%, PE 5%
# Condição: (PI+PF)/2 >= 4.5; se não, MF = (PI+PF)/2
# AA: notas -1/0/+1, descarta pior (se >=2 apresentações), converte p/ 0-12
# PE: notas -1/0/+1, sem descarte, converte p/ 0-10
# AA e PE: células pré-alocadas — usuario só preenche a próxima linha
# MF_cell = C55
# ════════════════════════════════════════════════════════════════════════
def build_hpe(wb):
    ws=wb.create_sheet("HPE"); std_cols(ws)
    header4(ws,
        "HPE — História do Pensamento Econômico",
        "Prof. Pedro Duarte / Fernando Leite Neto  |  Turma 4ECOB  |  Ter 16:30–18:30 / Qui 07:30–09:30",
        "MF=0,25·PI+0,35·PF+0,15·APS+0,20·AA+0,05·PE  |  Notas AA/PE: -1, 0 ou +1",
        "Aprovação: MF≥5,0  E  (PI+PF)/2≥4,5  |  APS zero = REPROVAÇÃO AUTOMÁTICA")
    inp_header(ws,6)
    rows_std=[
        (8, "PI — Prova Intermediária (25%)","26/03/2026",""),
        (9, "PF — Prova Final (35%)","28/05/2026",""),
        (10,"APS — Ativ. Prática (15%)","ao longo","2 APS no semestre; nota zero = reprovação"),
    ]
    for row,label,date,obs in rows_std:
        input_row(ws,row,label,date,obs)
    spacer(ws,11)

    # AA section
    mg(ws,"A12:D12","▼  ATIVIDADES DE APRENDIZAGEM (AA) — insira -1, 0 ou +1 nas células amarelas",
       F_MB,fnt(bold=True,color="FFFFFF",size=10),AC,20,12)
    for col,txt in [("A","Aula/Data"),("B","Tópico"),("C","Nota (-1/0/+1)"),("D","Obs")]:
        sc(ws,f"{col}13",txt,F_CH,fnt(bold=True,size=9),AC)
    ws.row_dimensions[13].height=16
    aa_start, aa_end = 14, 33  # 20 rows
    for r in range(aa_start, aa_end+1):
        idx=r-aa_start+1
        sc(ws,f"A{r}",f"AA {idx}",fn=fnt(size=10),al=AC)
        sc(ws,f"B{r}","",fn=fnt(size=9),al=AL)
        sc(ws,f"C{r}",None,F_YI,fn=fnt(),al=AC)
        sc(ws,f"D{r}","",fn=fnt(italic=True,size=9),al=AL)
        ws.row_dimensions[r].height=15
    spacer(ws,34)

    # PE section
    mg(ws,"A35:D35","▼  PERGUNTAS SOBRE APRESENTAÇÕES (PE) — insira -1, 0 ou +1",
       F_MB,fnt(bold=True,color="FFFFFF",size=10),AC,20,35)
    for col,txt in [("A","Aula/Data"),("B","Tópico"),("C","Nota (-1/0/+1)"),("D","Obs")]:
        sc(ws,f"{col}36",txt,F_CH,fnt(bold=True,size=9),AC)
    ws.row_dimensions[36].height=16
    pe_start, pe_end = 37, 51  # 15 rows
    for r in range(pe_start, pe_end+1):
        idx=r-pe_start+1
        sc(ws,f"A{r}",f"PE {idx}",fn=fnt(size=10),al=AC)
        sc(ws,f"B{r}","",fn=fnt(size=9),al=AL)
        sc(ws,f"C{r}",None,F_YI,fn=fnt(),al=AC)
        sc(ws,f"D{r}","",fn=fnt(italic=True,size=9),al=AL)
        ws.row_dimensions[r].height=15
    spacer(ws,52)

    # Calculated section
    calc_hdr(ws,53)
    aa_rng=f"C{aa_start}:C{aa_end}"; pe_rng=f"C{pe_start}:C{pe_end}"
    # AA mean (descarta pior se >=2 entradas)
    calc_row(ws,54,"Média AA (após descarte do pior)",
        f'=IF(COUNT({aa_rng})=0,"—",IF(COUNT({aa_rng})>=2,(SUM({aa_rng})-MIN({aa_rng}))/(COUNT({aa_rng})-1),AVERAGE({aa_rng})))',
        "Descarta 1 se ≥2 AA feitas")
    # AA grade (0-12 scale)
    calc_row(ws,55,"Nota AA (escala 0–12)",
        '=IF(C54="—","—",IF(C54<=-0.75,0,IF(C54<=-0.501,2,IF(C54<=-0.01,4,IF(C54<=0.19,6,IF(C54<=0.39,8,IF(C54<=0.59,10,12)))))))',
        "Tabela de conversão do plano")
    # PE mean (no descarte)
    calc_row(ws,56,"Média PE (sem descarte)",
        f'=IF(COUNT({pe_rng})=0,"—",AVERAGE({pe_rng}))',
        "Média simples de todas as PE")
    # PE grade (0-10 scale)
    calc_row(ws,57,"Nota PE (escala 0–10)",
        '=IF(C56="—","—",IF(C56<=-0.75,0,IF(C56<=-0.01,3,IF(C56<=0.29,6,10))))',
        "Tabela de conversão do plano")
    # MPF = (PI+PF)/2
    calc_row(ws,58,"MPF = (PI+PF)/2",
        '=IF(OR(NOT(ISNUMBER(C8)),NOT(ISNUMBER(C9))),"",( C8+C9)/2)',
        "Mínimo 4,5 p/ usar fórmula completa")
    # MF
    mf_formula=(
        '=IF(C58="","",IF(C58<4.5,C58,'
        '0.25*C8+0.35*C9+0.15*IF(ISNUMBER(C10),C10,0)'
        '+0.20*IF(ISNUMBER(C55),C55,0)'
        '+0.05*IF(ISNUMBER(C57),C57,0)))'
    )
    calc_row(ws,59,"MF (Média Final)",mf_formula,"Condicional em MPF≥4,5")
    calc_row(ws,60,"Situação atual",
        '=IF(C59="","Aguardando notas...",IF(C59>=5,"✓ APROVADO DIRETO!","Em andamento — veja simulação"))',
        "",fmt=None)
    add_cf(ws,"C60",green=True,red=False)
    spacer(ws,61)
    sim_hdr(ws,62)
    # PF needed: both (PI+PF)/2 >= 4.5 and full MF >= 5
    # From full: 0.25*PI + 0.35*PF + rest >= 5 => PF >= (5-0.25*PI-rest)/0.35
    # From MPF: (PI+PF)/2 >= 4.5 => PF >= 9 - PI
    ws["F63"].value=(
        '=IF(NOT(ISNUMBER(C8)),-999,'
        'MAX(9-C8,(5-0.25*C8-0.15*IF(ISNUMBER(C10),C10,0)'
        '-0.2*IF(ISNUMBER(C55),C55,0)-0.05*IF(ISNUMBER(C57),C57,0))/0.35))'
    )
    sim_f='=IF(NOT(ISNUMBER(C8)),"Preencha a PI primeiro",IF(F63<=0,"✓ Aprovado direto!",IF(F63>10,"✗ Reprovado mesmo com 10 na PF","Precisa de "&TEXT(ROUND(F63,2),"0.00")&" na PF")))'
    sc(ws,"A63","Nota necessária na PF (dados atuais)",fn=fnt(bold=True),al=AL)
    sc(ws,"B63","Brancos/vazios = 0 (pior caso)",fn=fnt(italic=True,size=9),al=AC)
    sc(ws,"C63",sim_f,F_OR,fn=fnt(bold=True),al=AC)
    ws.row_dimensions[63].height=22
    add_cf(ws,"C63",green=True,red=True)
    spacer(ws,64)
    mg(ws,"A64:D64",
       "Dica: adicione novas AA/PE nas próximas linhas livres. O cálculo atualiza automaticamente.",
       F_LB,fnt(italic=True,size=9),AC)
    ws.freeze_panes="A14"
    return "C59"

# ════════════════════════════════════════════════════════════════════════
# 5. MACROECONOMIA INTERNACIONAL
# AI 40%, AF 50%, APS 10% (4 APSs, descarta menor → média das 3 maiores)
# MSP = (AI+AF)/2
# MPP = 0.40*AI + 0.50*AF + 0.10*APS_media
# MF = MSP se MSP<4.5; MPP se MSP>=4.5
# APS zero = reprovação
# MF_cell = C19
# ════════════════════════════════════════════════════════════════════════
def build_macro(wb):
    ws=wb.create_sheet("Macro Internacional"); std_cols(ws)
    header4(ws,
        "MACROECONOMIA INTERNACIONAL — Controle de Notas",
        "Prof. Gino Olivares / Guilherme Duarte  |  Turma 4ECOB  |  Seg 16:30–18:30 / Qua 09:45–11:45",
        "MPP=0,40·AI+0,50·AF+0,10·APS  |  MSP=(AI+AF)/2  |  MF=MSP(MSP<4,5) ou MPP(MSP≥4,5)",
        "Aprovação: MSP≥4,5 E MF≥5,0  |  APS zero = REPROVAÇÃO AUTOMÁTICA")
    inp_header(ws,6)
    rows_in=[
        (8,  "AI — Avaliação Intermediária (40%)","25/03/2026",""),
        (9,  "APS 1 — 1ª de 4 (média 3 maiores)","04/03/2026",""),
        (10, "APS 2","11/03/2026",""),
        (11, "APS 3","27/04/2026",""),
        (12, "APS 4","13/05/2026","Descarta a menor das 4"),
        (13, "AF — Avaliação Final (50%)","27/05/2026",""),
    ]
    for row,label,date,obs in rows_in:
        input_row(ws,row,label,date,obs)
    spacer(ws,14)
    calc_hdr(ws,15)
    # APS media = média das 3 maiores de 4 (se <4 preenchidas, usa todas)
    calc_row(ws,16,"APS média (3 maiores de 4)",
        '=IF(COUNT(C9:C12)=0,"—",IF(COUNT(C9:C12)<4,AVERAGE(C9:C12),(LARGE(C9:C12,1)+LARGE(C9:C12,2)+LARGE(C9:C12,3))/3))',
        "Descarta a menor APS das 4")
    calc_row(ws,17,"MSP = (AI+AF)/2",
        '=IF(OR(NOT(ISNUMBER(C8)),NOT(ISNUMBER(C13))),"",( C8+C13)/2)',
        "Mínimo 4,5 p/ usar MPP")
    calc_row(ws,18,"MPP (Média Ponderada)",
        '=IF(OR(NOT(ISNUMBER(C8)),NOT(ISNUMBER(C13))),"",0.4*C8+0.5*C13+0.1*IF(ISNUMBER(C16),C16,0))',
        "0,40·AI+0,50·AF+0,10·APS")
    calc_row(ws,19,"MF (Média Final)",
        '=IF(C17="","",IF(C17<4.5,C17,C18))',
        "MSP<4,5→MF=MSP; senão MF=MPP")
    calc_row(ws,20,"Situação atual",
        '=IF(C19="","Aguardando notas...",IF(C19>=5,"✓ APROVADO DIRETO!","Em andamento — veja simulação"))',
        "",fmt=None)
    add_cf(ws,"C20",green=True,red=False)
    spacer(ws,21)
    sim_hdr(ws,22)
    # AF needed: both MSP>=4.5 and MPP>=5
    # MSP>=4.5: AF >= 9 - AI
    # MPP>=5: 0.4*AI + 0.5*AF + 0.1*APS >= 5 => AF >= (5-0.4*AI-0.1*APS)/0.5
    ws["F23"].value='=IF(NOT(ISNUMBER(C8)),-999,MAX(9-C8,(5-0.4*C8-0.1*IF(ISNUMBER(C16),C16,0))/0.5))'
    sim_f='=IF(NOT(ISNUMBER(C8)),"Preencha a AI primeiro",IF(F23<=0,"✓ Aprovado direto!",IF(F23>10,"✗ Reprovado mesmo com 10 na AF","Precisa de "&TEXT(ROUND(F23,2),"0.00")&" na AF")))'
    sc(ws,"A23","Nota necessária na AF (dados atuais)",fn=fnt(bold=True),al=AL)
    sc(ws,"B23","Brancos = 0 (pior caso)",fn=fnt(italic=True,size=9),al=AC)
    sc(ws,"C23",sim_f,F_OR,fn=fnt(bold=True),al=AC)
    ws.row_dimensions[23].height=22
    add_cf(ws,"C23",green=True,red=True)
    ws.freeze_panes="A8"
    return "C19"

# ════════════════════════════════════════════════════════════════════════
# 6. MICROECONOMIA II
# Q 10%, APS 10%, AI 30%, AF 50%
# MF = 0.10*Q + 0.10*APS + 0.30*AI + 0.50*AF
# Aprovação: MF>=5.0 E (AI+AF)/2>=4.0
# MF_cell = C19
# ════════════════════════════════════════════════════════════════════════
def build_micro2(wb):
    ws=wb.create_sheet("Micro II"); std_cols(ws)
    header4(ws,
        "MICROECONOMIA II — Controle de Notas",
        "Prof. Darcio Martins / Guilherme Carvalho  |  Turma 3D  |  Seg 14:15–16:15 / Qua 14:15–16:15",
        "MF = 0,10·Q + 0,10·APS + 0,30·AI + 0,50·AF  |  Q = média simples de 5 quizzes",
        "Aprovação: MF≥5,0  E  (AI+AF)/2≥4,0")
    inp_header(ws,6)
    rows_in=[
        (8,  "AI — Avaliação Intermediária (30%)","30/03/2026",""),
        (9,  "Quiz #1","04/03/2026",""),
        (10, "Quiz #2","11/03/2026",""),
        (11, "Quiz #4","15/04/2026",""),
        (12, "Quiz #5","06/05/2026",""),
        (13, "Quiz #6","20/05/2026","Sem descarte — média dos 5"),
        (14, "APS — Ativ. Prática (10%)","ao longo",""),
        (15, "AF — Avaliação Final (50%)","27/05/2026",""),
    ]
    for row,label,date,obs in rows_in:
        input_row(ws,row,label,date,obs)
    spacer(ws,16)
    calc_hdr(ws,17)
    calc_row(ws,18,"MQ (Média dos 5 Quizzes)",
        '=IFERROR(AVERAGE(C9:C13),"—")',
        "Média simples, sem descarte")
    calc_row(ws,19,"MSP = (AI+AF)/2",
        '=IF(OR(NOT(ISNUMBER(C8)),NOT(ISNUMBER(C15))),"",( C8+C15)/2)',
        "Mínimo 4,0 p/ aprovação")
    calc_row(ws,20,"MF (Média Final)",
        '=IF(OR(NOT(ISNUMBER(C8)),NOT(ISNUMBER(C15))),"",0.1*IF(ISNUMBER(C18),C18,0)+0.1*IF(ISNUMBER(C14),C14,0)+0.3*C8+0.5*C15)',
        "0,10·Q+0,10·APS+0,30·AI+0,50·AF")
    calc_row(ws,21,"Situação atual",
        '=IF(C20="","Aguardando notas...",IF(AND(C20>=5,IF(C19="",0,C19)>=4),"✓ APROVADO DIRETO!","Em andamento — veja simulação"))',
        "",fmt=None)
    add_cf(ws,"C21",green=True,red=False)
    spacer(ws,22)
    sim_hdr(ws,23)
    ws["F24"].value='=IF(NOT(ISNUMBER(C8)),-999,MAX(8-C8,(5-0.1*IF(ISNUMBER(C18),C18,0)-0.1*IF(ISNUMBER(C14),C14,0)-0.3*C8)/0.5))'
    sim_f='=IF(NOT(ISNUMBER(C8)),"Preencha a AI primeiro",IF(F24<=0,"✓ Aprovado direto!",IF(F24>10,"✗ Reprovado mesmo com 10 na AF","Precisa de "&TEXT(ROUND(F24,2),"0.00")&" na AF")))'
    sc(ws,"A24","Nota necessária na AF (dados atuais)",fn=fnt(bold=True),al=AL)
    sc(ws,"B24","Brancos = 0 (pior caso)",fn=fnt(italic=True,size=9),al=AC)
    sc(ws,"C24",sim_f,F_OR,fn=fnt(bold=True),al=AC)
    ws.row_dimensions[24].height=22
    add_cf(ws,"C24",green=True,red=True)
    ws.freeze_panes="A8"
    return "C20"

# ════════════════════════════════════════════════════════════════════════
# 7. MICROECONOMIA III
# PI 35%, PF 40%, QZ 10% (4 quizzes, descarta 1 → média 3 melhores)
# APS 10%, PA 5% (participação: 0/5/10, descarta 5 piores)
# MF = 0.35*PI + 0.40*PF + 0.10*QZ + 0.10*APS + 0.05*PA
# MF_cell = C22
# ════════════════════════════════════════════════════════════════════════
def build_micro3(wb):
    ws=wb.create_sheet("Micro III"); std_cols(ws)
    header4(ws,
        "MICROECONOMIA III — Controle de Notas",
        "Prof. Isabela Furtado / Fabio Rosa  |  Turma 4ECOB  |  Seg 09:45–11:45 / Qui 12:00–14:00",
        "MF=0,35·PI+0,40·PF+0,10·QZ+0,10·APS+0,05·PA  |  QZ=3 melhores de 4 quizzes",
        "Aprovação: MF≥5,0  |  PA: descarta 5 piores notas")
    inp_header(ws,6)
    rows_in=[
        (8,  "PI — Prova Intermediária (35%)","30/03/2026",""),
        (9,  "Quiz 1 (para QZ)","02/03/2026",""),
        (10, "Quiz 2 (para QZ)","23/03/2026",""),
        (11, "Quiz 3 (para QZ)","30/04/2026",""),
        (12, "Quiz 4 (para QZ)","14/05/2026","Descarta 1 pior → média dos 3 melhores"),
        (13, "APS 1 — Caso (10%)","16/04/2026","50% particip sala + 50% entrega"),
        (14, "APS 2 — Caso","21/05/2026","Média de APS1 e APS2 = APS"),
    ]
    for row,label,date,obs in rows_in:
        input_row(ws,row,label,date,obs)
    spacer(ws,15)

    # Participação section
    mg(ws,"A16:D16","▼  PARTICIPAÇÃO (PA) — insira 0, 5 ou 10 em cada aula  (células amarelas)",
       F_MB,fnt(bold=True,color="FFFFFF",size=10),AC,20,16)
    for col,txt in [("A","Aula"),("B","Data"),("C","Nota (0/5/10)"),("D","Obs")]:
        sc(ws,f"{col}17",txt,F_CH,fnt(bold=True,size=9),AC)
    ws.row_dimensions[17].height=16
    pa_start, pa_end = 18, 47  # 30 rows
    for r in range(pa_start, pa_end+1):
        idx=r-pa_start+1
        sc(ws,f"A{r}",f"Aula {idx}",fn=fnt(size=10),al=AC)
        sc(ws,f"B{r}","",fn=fnt(size=9),al=AL)
        sc(ws,f"C{r}",None,F_YI,fn=fnt(),al=AC)
        sc(ws,f"D{r}","",fn=fnt(italic=True,size=9),al=AL)
        ws.row_dimensions[r].height=14
    spacer(ws,48)

    calc_hdr(ws,49)
    pa_rng=f"C{pa_start}:C{pa_end}"
    # QZ: média dos 3 melhores de 4
    calc_row(ws,50,"QZ (3 melhores quizzes de 4)",
        '=IF(COUNT(C9:C12)=0,"—",IF(COUNT(C9:C12)<=3,AVERAGE(C9:C12),(LARGE(C9:C12,1)+LARGE(C9:C12,2)+LARGE(C9:C12,3))/3))',
        "Descarta o pior dos 4 quizzes")
    # APS media
    calc_row(ws,51,"APS (média APS1 e APS2)",
        '=IF(AND(ISNUMBER(C13),ISNUMBER(C14)),AVERAGE(C13:C14),IF(ISNUMBER(C13),C13,IF(ISNUMBER(C14),C14,"—")))',
        "Média de 2 casos")
    # PA: média descartando 5 piores
    calc_row(ws,52,"PA média (descarta 5 piores)",
        f'=IF(COUNT({pa_rng})=0,"—",IF(COUNT({pa_rng})<=5,AVERAGE({pa_rng}),(SUM({pa_rng})-SUMPRODUCT(SMALL({pa_rng},ROW(INDIRECT("1:"&MIN(5,COUNT({pa_rng})))))))/MAX(COUNT({pa_rng})-5,1)))',
        "Descarta 5 piores participações")
    calc_row(ws,53,"MF (Média Final)",
        '=IF(OR(NOT(ISNUMBER(C8)),NOT(ISNUMBER(ws_pf))),"",0.35*C8+0.40*ws_pf+0.10*IF(ISNUMBER(C50),C50,0)+0.10*IF(ISNUMBER(C51),C51,0)+0.05*IF(ISNUMBER(C52),C52,0))',
        "0,35·PI+0,40·PF+0,10·QZ+0,10·APS+0,05·PA")
    calc_row(ws,54,"Situação atual",
        '=IF(C53="","Aguardando notas...",IF(C53>=5,"✓ APROVADO DIRETO!","Em andamento — veja simulação"))',
        "",fmt=None)
    add_cf(ws,"C54",green=True,red=False)
    spacer(ws,55)
    sim_hdr(ws,56)
    ws["F57"].value='=IF(NOT(ISNUMBER(C8)),-999,(5-0.35*C8-0.10*IF(ISNUMBER(C50),C50,0)-0.10*IF(ISNUMBER(C51),C51,0)-0.05*IF(ISNUMBER(C52),C52,0))/0.40)'
    sim_f='=IF(NOT(ISNUMBER(C8)),"Preencha a PI primeiro",IF(F57<=0,"✓ Aprovado direto!",IF(F57>10,"✗ Reprovado mesmo com 10 na PF","Precisa de "&TEXT(ROUND(F57,2),"0.00")&" na PF")))'
    sc(ws,"A57","Nota necessária na PF (dados atuais)",fn=fnt(bold=True),al=AL)
    sc(ws,"B57","Brancos = 0 (pior caso)",fn=fnt(italic=True,size=9),al=AC)
    sc(ws,"C57",sim_f,F_OR,fn=fnt(bold=True),al=AC)
    ws.row_dimensions[57].height=22
    add_cf(ws,"C57",green=True,red=True)
    ws.freeze_panes="A18"

    # Fix PF reference (C_pf is a placeholder — we need to add a PF input row)
    # Actually I forgot to add PF input! Let me fix by adding it after PA section.
    # The PF row needs to be accessible. Let me put it at C15 (after rows_in ends at 14)
    # Actually I added rows_in up to row 14 only. Row 15 is a spacer.
    # I'll add PF after the spacer at row 15... but 16 is already the PA header.
    # Better: add PF as row 15 and shift PA down.
    # I need to redo the PA layout. Let me fix this by inserting PF at row 15.
    # Actually, the simplest fix: I'll correct the build_micro3 function.
    # Let me just fix it now: C53 formula uses "ws_pf" placeholder which is wrong.
    # The PF input should be at row 15. I'll add it there and fix the spacer to row 16.
    # This requires rewriting the function. Let me do it properly below.
    return "C53"  # placeholder — will fix below


def build_micro3_fixed(wb):
    ws=wb.create_sheet("Micro III"); std_cols(ws)
    header4(ws,
        "MICROECONOMIA III — Controle de Notas",
        "Prof. Isabela Furtado / Fabio Rosa  |  Turma 4ECOB  |  Seg 09:45–11:45 / Qui 12:00–14:00",
        "MF=0,35·PI+0,40·PF+0,10·QZ+0,10·APS+0,05·PA  |  QZ=3 melhores de 4 quizzes",
        "Aprovação: MF≥5,0  |  PA: descarta 5 piores notas")
    inp_header(ws,6)
    rows_in=[
        (8,  "PI — Prova Intermediária (35%)","30/03/2026",""),
        (9,  "Quiz 1 (para QZ)","02/03/2026",""),
        (10, "Quiz 2 (para QZ)","23/03/2026",""),
        (11, "Quiz 3 (para QZ)","30/04/2026",""),
        (12, "Quiz 4 (para QZ)","14/05/2026","Descarta 1 pior → média dos 3 melhores"),
        (13, "APS 1 — Caso (10%)","16/04/2026","50% particip sala + 50% entrega"),
        (14, "APS 2 — Caso","21/05/2026","Média de APS1 e APS2 = APS"),
        (15, "PF — Prova Final (40%)","27/05/2026",""),
    ]
    for row,label,date,obs in rows_in:
        input_row(ws,row,label,date,obs)
    spacer(ws,16)

    # Participação section header
    mg(ws,"A17:D17","▼  PARTICIPAÇÃO (PA) — insira 0, 5 ou 10 em cada aula  (células amarelas)",
       F_MB,fnt(bold=True,color="FFFFFF",size=10),AC,20,17)
    for col,txt in [("A","Aula"),("B","Data"),("C","Nota (0/5/10)"),("D","Obs")]:
        sc(ws,f"{col}18",txt,F_CH,fnt(bold=True,size=9),AC)
    ws.row_dimensions[18].height=16
    pa_start, pa_end = 19, 48  # 30 rows
    for r in range(pa_start, pa_end+1):
        idx=r-pa_start+1
        sc(ws,f"A{r}",f"Aula {idx}",fn=fnt(size=10),al=AC)
        sc(ws,f"B{r}","",fn=fnt(size=9),al=AL)
        sc(ws,f"C{r}",None,F_YI,fn=fnt(),al=AC)
        sc(ws,f"D{r}","",fn=fnt(italic=True,size=9),al=AL)
        ws.row_dimensions[r].height=14
    spacer(ws,49)

    calc_hdr(ws,50)
    pa_rng=f"C{pa_start}:C{pa_end}"
    calc_row(ws,51,"QZ (3 melhores quizzes de 4)",
        '=IF(COUNT(C9:C12)=0,"—",IF(COUNT(C9:C12)<=3,AVERAGE(C9:C12),(LARGE(C9:C12,1)+LARGE(C9:C12,2)+LARGE(C9:C12,3))/3))',
        "Descarta o pior dos 4 quizzes")
    calc_row(ws,52,"APS (média APS1 e APS2)",
        '=IF(AND(ISNUMBER(C13),ISNUMBER(C14)),AVERAGE(C13:C14),IF(ISNUMBER(C13),C13,IF(ISNUMBER(C14),C14,"—")))',
        "Média de 2 casos")
    calc_row(ws,53,"PA média (descarta 5 piores)",
        f'=IF(COUNT({pa_rng})=0,"—",IF(COUNT({pa_rng})<=5,AVERAGE({pa_rng}),(SUM({pa_rng})-SUMPRODUCT(SMALL({pa_rng},ROW(INDIRECT("1:5")))))/MAX(COUNT({pa_rng})-5,1)))',
        "Descarta 5 piores participações")
    calc_row(ws,54,"MF (Média Final)",
        '=IF(OR(NOT(ISNUMBER(C8)),NOT(ISNUMBER(C15))),"",0.35*C8+0.40*C15+0.10*IF(ISNUMBER(C51),C51,0)+0.10*IF(ISNUMBER(C52),C52,0)+0.05*IF(ISNUMBER(C53),C53,0))',
        "0,35·PI+0,40·PF+0,10·QZ+0,10·APS+0,05·PA")
    calc_row(ws,55,"Situação atual",
        '=IF(C54="","Aguardando notas...",IF(C54>=5,"✓ APROVADO DIRETO!","Em andamento — veja simulação"))',
        "",fmt=None)
    add_cf(ws,"C55",green=True,red=False)
    spacer(ws,56)
    sim_hdr(ws,57)
    ws["F58"].value='=IF(NOT(ISNUMBER(C8)),-999,(5-0.35*C8-0.10*IF(ISNUMBER(C51),C51,0)-0.10*IF(ISNUMBER(C52),C52,0)-0.05*IF(ISNUMBER(C53),C53,0))/0.40)'
    sim_f='=IF(NOT(ISNUMBER(C8)),"Preencha a PI primeiro",IF(F58<=0,"✓ Aprovado direto!",IF(F58>10,"✗ Reprovado mesmo com 10 na PF","Precisa de "&TEXT(ROUND(F58,2),"0.00")&" na PF")))'
    sc(ws,"A58","Nota necessária na PF (dados atuais)",fn=fnt(bold=True),al=AL)
    sc(ws,"B58","Brancos = 0 (pior caso)",fn=fnt(italic=True,size=9),al=AC)
    sc(ws,"C58",sim_f,F_OR,fn=fnt(bold=True),al=AC)
    ws.row_dimensions[58].height=22
    add_cf(ws,"C58",green=True,red=True)
    ws.freeze_panes="A19"
    return "C54"

# ════════════════════════════════════════════════════════════════════════
# INÍCIO — Dashboard com todas as matérias
# Referencia MF de cada sheet
# ════════════════════════════════════════════════════════════════════════
def build_inicio(wb, mf_refs):
    ws=wb.create_sheet("Início", 0)
    ws.sheet_view.showGridLines=False
    ws.column_dimensions["A"].width=28
    ws.column_dimensions["B"].width=18
    ws.column_dimensions["C"].width=14
    ws.column_dimensions["D"].width=22
    ws.column_dimensions["E"].width=16

    mg(ws,"A1:E1","4º SEMESTRE — PAINEL DE NOTAS",
       F_DB,fnt(bold=True,color="FFFFFF",size=16),AC,36,1)
    mg(ws,"A2:E2","Insper  |  2026-1  |  Atualizado automaticamente com as notas inseridas em cada aba",
       F_MB,fnt(color="FFFFFF",size=10),AC)
    ws.row_dimensions[3].height=8

    for col,hdr in zip(["A","B","C","D","E"],
                       ["Matéria","Professor(es)","MF Atual","Situação","Prova Final"]):
        sc(ws,f"{col}4",hdr,F_CH,fnt(bold=True),AC)
    ws.row_dimensions[4].height=18

    subjects=[
        ("Econometria",     "Bortoluzzo / Martins",  "Econometria",      mf_refs[0], "27/05/2026"),
        ("Estatística II",  "M.Kelly / Rinaldo",     "Estatística II",   mf_refs[1], "27/05/2026"),
        ("Finanças II",     "G.Soares / P.Neto",     "Finanças II",      mf_refs[2], "26/05/2026"),
        ("HPE",             "P.Duarte / F.Leite",    "HPE",              mf_refs[3], "28/05/2026"),
        ("Macro Internacional","G.Olivares / G.Duarte","Macro Internacional",mf_refs[4],"27/05/2026"),
        ("Micro II",        "Darcio / G.Carvalho",   "Micro II",         mf_refs[5], "27/05/2026"),  # MF = C20
        ("Micro III",       "I.Furtado / F.Rosa",    "Micro III",        mf_refs[6], "27/05/2026"),
    ]
    for i,(name,prof,sheet,mf_cell,pf_date) in enumerate(subjects,start=5):
        row=i
        sc(ws,f"A{row}",name,fn=fnt(bold=True),al=AL)
        sc(ws,f"B{row}",prof,fn=fnt(size=10),al=AL)
        sc(ws,f"C{row}",f"=IF(ISNUMBER('{sheet}'!{mf_cell}),'{sheet}'!{mf_cell},\"—\")",
           fn=fnt(bold=True),al=AC,fmt="0.00")
        sc(ws,f"D{row}",f"=IF(C{row}=\"—\",\"—\",IF(C{row}>=5,\"✓ Aprovado\",\"Em andamento\"))",
           fn=fnt(),al=AC,fmt=None)
        sc(ws,f"E{row}",pf_date,fn=fnt(size=10),al=AC)
        ws.row_dimensions[row].height=18
        ws.conditional_formatting.add(f"D{row}",FormulaRule(
            formula=[f'NOT(ISERROR(SEARCH("✓",D{row})))'],fill=F_GN))

    ws.row_dimensions[12].height=8
    mg(ws,"A13:E13",
       "Acesse cada aba para inserir notas. Células amarelas = entrada manual. Cinzas = calculado.",
       F_LB,fnt(italic=True,size=10),AC)

# ════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════
def main():
    wb=openpyxl.Workbook(); wb.remove(wb.active)
    mf1=build_econometria(wb)
    mf2=build_estatistica(wb)
    mf3=build_financas(wb)
    mf4=build_hpe(wb)
    mf5=build_macro(wb)
    mf6=build_micro2(wb)
    mf7=build_micro3_fixed(wb)
    build_inicio(wb,[mf1,mf2,mf3,mf4,mf5,mf6,mf7])
    wb.save(OUT_PATH)
    print(f"OK Salvo: {OUT_PATH}")

if __name__=="__main__": main()
