"""
Cria a planilha Excel de gestão de estoque da lanchonete.
Todos os indicadores são FÓRMULAS do Excel: ao alterar parâmetros,
vendas ou estoque atual, a planilha recalcula sozinha.
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parents[1]
DATA_DIR = BASE / "data"
SAIDA = BASE / "excel" / "gestao_estoque_lanchonete.xlsx"

FONTE = "Arial"
AZUL_ESCURO = "1F3A5F"
F_TITULO = Font(name=FONTE, size=16, bold=True, color=AZUL_ESCURO)
F_SUB = Font(name=FONTE, size=10, italic=True, color="6B6B6B")
F_CAB = Font(name=FONTE, size=10, bold=True, color="FFFFFF")
F_NORMAL = Font(name=FONTE, size=10)
F_INPUT = Font(name=FONTE, size=10, color="0000FF")
F_LINK = Font(name=FONTE, size=10, color="008000")
F_NEGRITO = Font(name=FONTE, size=10, bold=True)
FILL_CAB = PatternFill("solid", fgColor=AZUL_ESCURO)
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_KPI = PatternFill("solid", fgColor="EEF4FB")
BORDA = Border(bottom=Side(style="thin", color="D9D9D9"))
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)

produtos = pd.read_csv(DATA_DIR / "produtos.csv")
vendas = pd.read_csv(DATA_DIR / "vendas_diarias.csv", parse_dates=["data"])
vendas_wide = vendas.pivot(index="data", columns="codigo", values="quantidade").reset_index()
N_PROD = len(produtos)
N_DIAS = len(vendas_wide)

wb = Workbook()


def cabecalho(ws, linha, titulos, larguras=None):
    for i, t in enumerate(titulos, start=1):
        c = ws.cell(row=linha, column=i, value=t)
        c.font, c.fill, c.alignment = F_CAB, FILL_CAB, CENTRO
    ws.row_dimensions[linha].height = 32
    if larguras:
        for i, w in enumerate(larguras, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def titulo(ws, texto, sub):
    ws["A1"] = texto
    ws["A1"].font = F_TITULO
    ws["A2"] = sub
    ws["A2"].font = F_SUB


# ---------------------------------------------------------------- Leia-me
ws = wb.active
ws.title = "Leia-me"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 90
ws["B2"] = "Gestão de Estoque e Estoque de Segurança — Lanchonete"
ws["B2"].font = F_TITULO
ws["B3"] = "Projeto de Ciência de Dados · dados simulados para fins de estudo"
ws["B3"].font = F_SUB

linhas = [
    ("OBJETIVO", None),
    ("", "Evitar a falta de insumos (ruptura) sem deixar dinheiro parado em estoque, calculando quando e quanto comprar."),
    ("ABAS", None),
    ("Parâmetros", "Nível de serviço e dias de cobertura. Altere aqui e todos os cálculos se atualizam."),
    ("Produtos", "Cadastro dos insumos: custo, lead time (prazo do fornecedor) e estoque atual."),
    ("Vendas", f"Consumo diário de cada insumo nos últimos {N_DIAS} dias (uma coluna por produto)."),
    ("Análise", "Demanda média, desvio padrão, estoque de segurança, ponto de pedido, status e sugestão de compra."),
    ("Curva ABC", "Classificação dos insumos pela importância no custo mensal."),
    ("Dashboard", "Indicadores-chave e gráficos."),
    ("FÓRMULAS", None),
    ("Estoque de segurança (ES)", "ES = Z × σ × √LT   (Z do nível de serviço, σ = desvio padrão diário, LT = lead time em dias)"),
    ("Ponto de pedido (PP)", "PP = demanda média diária × LT + ES   → quando o estoque chega aqui, é hora de comprar"),
    ("Estoque máximo", "Estoque máximo = PP + demanda média × dias de cobertura"),
    ("Quantidade sugerida", "Se estoque atual ≤ PP: comprar (estoque máximo − estoque atual); senão, 0"),
    ("Status", "CRÍTICO: estoque ≤ ES  ·  REPOR: estoque ≤ PP  ·  OK: acima do PP"),
    ("COMO USAR", None),
    ("Células amarelas, texto azul", "São entradas: pode editar (parâmetros, estoque atual, custos, lead time e vendas)."),
    ("Texto preto", "Fórmulas: não editar."),
    ("Texto verde", "Valores puxados de outra aba."),
    ("Rotina sugerida", "1) Lançar o consumo do dia na aba Vendas  2) Atualizar o estoque atual em Produtos  3) Ver a aba Análise / Dashboard e fazer os pedidos com status REPOR ou CRÍTICO."),
]
r = 5
for a, b in linhas:
    if b is None:
        ws.cell(row=r, column=2, value=a).font = Font(name=FONTE, size=11, bold=True, color=AZUL_ESCURO)
        r += 1
        continue
    ws.cell(row=r, column=2, value=a).font = F_NEGRITO
    c = ws.cell(row=r, column=3, value=b)
    c.font, c.alignment = F_NORMAL, Alignment(wrap_text=True, vertical="top")
    r += 1
ws["B21"].fill = FILL_INPUT
ws["B21"].font = Font(name=FONTE, size=10, bold=True, color="0000FF")
ws["B23"].font = Font(name=FONTE, size=10, bold=True, color="008000")

# ---------------------------------------------------------------- Parâmetros
ws = wb.create_sheet("Parâmetros")
ws.sheet_view.showGridLines = False
titulo(ws, "Parâmetros", "Células amarelas são editáveis")
ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 70
params = [
    ("Nível de serviço", 0.95, "0.0%", "Probabilidade de NÃO faltar o insumo durante o lead time. Use entre 50% e 99,9%."),
    ("Fator Z", "=NORMSINV(B4)", "0.000", "Calculado a partir do nível de serviço (distribuição normal padrão). 95% → 1,645"),
    ("Dias de cobertura do pedido", 7, "0", "Quantos dias de consumo cada compra deve cobrir além do ponto de pedido."),
    ("Dias de histórico de vendas", f"=COUNT(Vendas!A5:A{4 + N_DIAS})", "0", "Contado automaticamente na aba Vendas."),
]
for i, (nome, val, fmt, nota) in enumerate(params, start=4):
    ws.cell(row=i, column=1, value=nome).font = F_NEGRITO
    c = ws.cell(row=i, column=2, value=val)
    c.number_format = fmt
    c.alignment = Alignment(horizontal="center")
    if isinstance(val, str):
        c.font = F_NORMAL
    else:
        c.font, c.fill = F_INPUT, FILL_INPUT
    ws.cell(row=i, column=3, value=nota).font = F_SUB
dv = DataValidation(type="decimal", operator="between", formula1="0.5", formula2="0.999",
                    showErrorMessage=True, error="Use um valor entre 50% e 99,9%.")
ws.add_data_validation(dv)
dv.add("B4")

# ---------------------------------------------------------------- Produtos
ws = wb.create_sheet("Produtos")
titulo(ws, "Cadastro de insumos", "Custo, lead time e estoque atual são entradas (amarelo)")
cab = ["Código", "Produto", "Categoria", "Unidade", "Custo unitário (R$)", "Lead time (dias)", "Fornecedor", "Estoque atual"]
cabecalho(ws, 4, cab, [9, 28, 15, 9, 13, 11, 24, 12])
for i, p in enumerate(produtos.itertuples(index=False), start=5):
    vals = [p.codigo, p.produto, p.categoria, p.unidade, p.custo_unitario, p.lead_time_dias, p.fornecedor, p.estoque_atual]
    for j, v in enumerate(vals, start=1):
        c = ws.cell(row=i, column=j, value=v)
        c.font, c.border = F_NORMAL, BORDA
        if j in (5, 6, 8):
            c.font, c.fill = F_INPUT, FILL_INPUT
            c.alignment = Alignment(horizontal="center")
    ws.cell(row=i, column=5).number_format = 'R$ #,##0.00'
    ws.cell(row=i, column=8).number_format = '#,##0.0'
ws.freeze_panes = "C5"
ws["H4"].comment = Comment("Atualize com a contagem física do estoque. Valores iniciais simulados.", "Projeto")
ULT_P = 4 + N_PROD

# ---------------------------------------------------------------- Vendas
ws = wb.create_sheet("Vendas")
titulo(ws, "Consumo diário por insumo", "Quantidade consumida por dia, na unidade de cada produto (dados simulados)")
ws.cell(row=4, column=1, value="Data")
for j, cod in enumerate(produtos.codigo, start=2):
    ws.cell(row=3, column=j, value=produtos.loc[produtos.codigo == cod, "produto"].iloc[0])
    ws.cell(row=4, column=j, value=cod)
ws.cell(row=4, column=N_PROD + 2, value="Custo do dia (R$)")
cabecalho(ws, 4, [ws.cell(row=4, column=k).value for k in range(1, N_PROD + 3)])
ws.column_dimensions["A"].width = 12
for j in range(2, N_PROD + 2):
    ws.column_dimensions[get_column_letter(j)].width = 11
    c = ws.cell(row=3, column=j)
    c.font, c.alignment = Font(name=FONTE, size=8, color="6B6B6B"), CENTRO
ws.row_dimensions[3].height = 36
ws.column_dimensions[get_column_letter(N_PROD + 2)].width = 15
col_custo = get_column_letter(N_PROD + 2)
ult_col_prod = get_column_letter(N_PROD + 1)
for i, row in enumerate(vendas_wide.itertuples(index=False), start=5):
    c = ws.cell(row=i, column=1, value=row[0].to_pydatetime())
    c.number_format, c.font = "dd/mm/yyyy", F_NORMAL
    for j, v in enumerate(row[1:], start=2):
        c = ws.cell(row=i, column=j, value=float(v))
        c.font, c.number_format = F_INPUT, "#,##0.0"
    c = ws.cell(row=i, column=N_PROD + 2,
                value="=" + "+".join(f"{get_column_letter(j)}{i}*Produtos!$E${3 + j}" for j in range(2, N_PROD + 2)))
    c.font, c.number_format = F_NORMAL, 'R$ #,##0.00'
ws.freeze_panes = "B5"
ULT_V = 4 + N_DIAS

# ---------------------------------------------------------------- Análise
ws = wb.create_sheet("Análise")
titulo(ws, "Análise de estoque e estoque de segurança", "Tudo calculado por fórmula — não editar")
cab = ["Código", "Produto", "Unidade", "Demanda média diária", "Desvio padrão diário", "Lead time (dias)",
       "Estoque de segurança", "Ponto de pedido", "Estoque máximo", "Estoque atual", "Dias de cobertura",
       "Status", "Qtd. sugerida de compra", "Custo do pedido (R$)", "Valor em estoque (R$)"]
cabecalho(ws, 4, cab, [9, 28, 9, 12, 12, 10, 12, 12, 12, 11, 11, 11, 13, 14, 14])
VR = f"Vendas!$B$5:${ult_col_prod}${ULT_V}"
VH = f"Vendas!$B$4:${ult_col_prod}$4"
for k in range(N_PROD):
    i = 5 + k
    pi = 5 + k
    col_v = f"INDEX({VR},0,MATCH($A{i},{VH},0))"
    f = {
        "A": f"=Produtos!A{pi}",
        "B": f"=Produtos!B{pi}",
        "C": f"=Produtos!D{pi}",
        "D": f"=AVERAGE({col_v})",
        "E": f"=STDEV({col_v})",
        "F": f"=Produtos!F{pi}",
        "G": f"=ROUNDUP('Parâmetros'!$B$5*E{i}*SQRT(F{i}),0)",
        "H": f"=ROUNDUP(D{i}*F{i}+G{i},0)",
        "I": f"=ROUNDUP(H{i}+D{i}*'Parâmetros'!$B$6,0)",
        "J": f"=Produtos!H{pi}",
        "K": f"=IFERROR(J{i}/D{i},0)",
        "L": f'=IF(J{i}<=G{i},"CRÍTICO",IF(J{i}<=H{i},"REPOR","OK"))',
        "M": f'=IF(L{i}="OK",0,I{i}-J{i})',
        "N": f"=M{i}*Produtos!E{pi}",
        "O": f"=J{i}*Produtos!E{pi}",
    }
    fmts = {"D": "#,##0.0", "E": "#,##0.0", "G": "#,##0", "H": "#,##0", "I": "#,##0", "J": "#,##0.0",
            "K": '0.0" d"', "M": "#,##0;-#,##0;-", "N": 'R$ #,##0.00;-R$ #,##0.00;-', "O": 'R$ #,##0.00'}
    for col, formula in f.items():
        c = ws[f"{col}{i}"]
        c.value = formula
        c.font = F_LINK if col in "ABCFJ" else F_NORMAL
        c.border = BORDA
        c.number_format = fmts.get(col, "General")
        if col not in "AB":
            c.alignment = Alignment(horizontal="center")
ULT_A = 4 + N_PROD
tot = ULT_A + 1
ws[f"B{tot}"] = "TOTAL"
ws[f"B{tot}"].font = F_NEGRITO
for col in "NO":
    c = ws[f"{col}{tot}"]
    c.value = f"=SUM({col}5:{col}{ULT_A})"
    c.font, c.number_format = F_NEGRITO, 'R$ #,##0.00'
    c.alignment = Alignment(horizontal="center")
faixa = f"A5:O{ULT_A}"
ws.conditional_formatting.add(faixa, FormulaRule(formula=['$L5="CRÍTICO"'], fill=PatternFill("solid", fgColor="F8D7D7")))
ws.conditional_formatting.add(faixa, FormulaRule(formula=['$L5="REPOR"'], fill=PatternFill("solid", fgColor="FFF0C2")))
ws.conditional_formatting.add(f"L5:L{ULT_A}", FormulaRule(formula=['$L5="OK"'], fill=PatternFill("solid", fgColor="D9F2DC")))
ws.freeze_panes = "C5"
ws["G4"].comment = Comment("ES = Z × desvio padrão × √lead time, arredondado para cima.", "Projeto")
ws["H4"].comment = Comment("PP = demanda média × lead time + ES. Ao atingir este nível, faça o pedido.", "Projeto")

# ---------------------------------------------------------------- Curva ABC
# Ordenada em Python pelo valor de consumo; os valores continuam sendo fórmulas.
ws = wb.create_sheet("Curva ABC")
titulo(ws, "Curva ABC dos insumos", "Pelo custo de consumo mensal (30 dias). A: até 80% acumulado · B: até 95% · C: restante")
cab = ["Código", "Produto", "Consumo mensal (R$)", "% do total", "% acumulado", "Classe"]
cabecalho(ws, 4, cab, [9, 28, 16, 12, 13, 9])
ordem = (produtos.merge(vendas.groupby("codigo").quantidade.mean().rename("d"), on="codigo")
         .assign(v=lambda x: x.d * x.custo_unitario).sort_values("v", ascending=False).codigo.tolist())
ULT_C = 4 + N_PROD
for k, cod in enumerate(ordem):
    i = 5 + k
    m = f"MATCH($A{i},'Análise'!$A$5:$A${ULT_A},0)"
    pm = f"MATCH($A{i},Produtos!$A$5:$A${ULT_P},0)"
    ws[f"A{i}"] = cod
    ws[f"B{i}"] = f"=INDEX('Análise'!$B$5:$B${ULT_A},{m})"
    ws[f"C{i}"] = f"=INDEX('Análise'!$D$5:$D${ULT_A},{m})*30*INDEX(Produtos!$E$5:$E${ULT_P},{pm})"
    ws[f"D{i}"] = f"=C{i}/SUM($C$5:$C${ULT_C})"
    ws[f"E{i}"] = f'=SUMIF($C$5:$C${ULT_C},">="&C{i})/SUM($C$5:$C${ULT_C})'
    ws[f"F{i}"] = f'=IF(E{i}<=0.8,"A",IF(E{i}<=0.95,"B","C"))'
    for col, fmt in zip("ABCDEF", ["General", "General", 'R$ #,##0.00', "0.0%", "0.0%", "General"]):
        c = ws[f"{col}{i}"]
        c.number_format, c.border = fmt, BORDA
        c.font = F_LINK if col == "B" else F_NORMAL
        if col in "CDEF":
            c.alignment = Alignment(horizontal="center")
ws[f"B{ULT_C + 1}"] = "TOTAL"
ws[f"B{ULT_C + 1}"].font = F_NEGRITO
ws[f"C{ULT_C + 1}"] = f"=SUM(C5:C{ULT_C})"
ws[f"C{ULT_C + 1}"].font, ws[f"C{ULT_C + 1}"].number_format = F_NEGRITO, 'R$ #,##0.00'
ws[f"A{ULT_C + 3}"] = ("Obs.: a ordem das linhas foi definida pelo histórico atual. Se os dados mudarem muito, "
                      "use Dados > Classificar pela coluna C (maior para menor) para refazer a curva.")
ws[f"A{ULT_C + 3}"].font = F_SUB
for cls, cor in (("A", "D6E6FA"), ("B", "FDE3D4"), ("C", "D4F3EC")):
    ws.conditional_formatting.add(f"F5:F{ULT_C}", FormulaRule(formula=[f'$F5="{cls}"'], fill=PatternFill("solid", fgColor=cor)))

# ---------------------------------------------------------------- Dashboard
ws = wb.create_sheet("Dashboard", 1)
ws.sheet_view.showGridLines = False
titulo(ws, "Dashboard de estoque — Lanchonete", "Atualiza automaticamente com as abas de dados")
for col in "ABCDEFGHIJ":
    ws.column_dimensions[col].width = 16
kpis = [
    ("Itens CRÍTICOS", f'=COUNTIF(\'Análise\'!L5:L{ULT_A},"CRÍTICO")', "0"),
    ("Itens para REPOR", f'=COUNTIF(\'Análise\'!L5:L{ULT_A},"REPOR")', "0"),
    ("Itens OK", f'=COUNTIF(\'Análise\'!L5:L{ULT_A},"OK")', "0"),
    ("Custo do pedido sugerido", f"='Análise'!N{tot}", 'R$ #,##0'),
    ("Valor em estoque", f"='Análise'!O{tot}", 'R$ #,##0'),
]
for k, (nome, formula, fmt) in enumerate(kpis):
    col = 1 + k * 2
    ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
    ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
    a = ws.cell(row=4, column=col, value=nome)
    a.font = Font(name=FONTE, size=9, bold=True, color="6B6B6B")
    a.alignment, a.fill = CENTRO, FILL_KPI
    ws.cell(row=4, column=col + 1).fill = FILL_KPI
    b = ws.cell(row=5, column=col, value=formula)
    b.font = Font(name=FONTE, size=20, bold=True, color=AZUL_ESCURO)
    b.alignment, b.number_format, b.fill = CENTRO, fmt, FILL_KPI
    ws.cell(row=5, column=col + 1).fill = FILL_KPI
ws.row_dimensions[5].height = 40

# Tabela de compras pendentes (fórmulas que listam até 12 itens com status ≠ OK)
ws["A8"] = "Lista de compras (itens com status REPOR ou CRÍTICO)"
ws["A8"].font = Font(name=FONTE, size=11, bold=True, color=AZUL_ESCURO)
for j, t in enumerate(["Produto", "", "Status", "Comprar", "Unidade", "Custo (R$)"], start=1):
    c = ws.cell(row=9, column=j, value=t)
    c.font, c.fill, c.alignment = F_CAB, FILL_CAB, CENTRO
ws.merge_cells("A9:B9")
aux = f"'Análise'!$P$5:$P${ULT_A}"
# coluna auxiliar P na Análise: número sequencial dos itens a comprar
wa = wb["Análise"]
wa["P4"] = "Aux. lista"
wa["P4"].font = Font(name=FONTE, size=8, color="9C9C9C")
for i in range(5, ULT_A + 1):
    wa[f"P{i}"] = f'=IF(L{i}="OK","",COUNTIF($L$5:L{i},"<>OK"))'
    wa[f"P{i}"].font = Font(name=FONTE, size=8, color="9C9C9C")
for n in range(1, N_PROD + 1):
    r = 9 + n
    m = f"MATCH({n},{aux},0)"
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws[f"A{r}"] = f'=IFERROR(INDEX(\'Análise\'!$B$5:$B${ULT_A},{m}),"")'
    ws[f"C{r}"] = f'=IFERROR(INDEX(\'Análise\'!$L$5:$L${ULT_A},{m}),"")'
    ws[f"D{r}"] = f'=IFERROR(INDEX(\'Análise\'!$M$5:$M${ULT_A},{m}),"")'
    ws[f"E{r}"] = f'=IFERROR(INDEX(\'Análise\'!$C$5:$C${ULT_A},{m}),"")'
    ws[f"F{r}"] = f'=IFERROR(INDEX(\'Análise\'!$N$5:$N${ULT_A},{m}),"")'
    for col, fmt in zip("ACDEF", ["General", "General", "#,##0", "General", 'R$ #,##0.00']):
        c = ws[f"{col}{r}"]
        c.font, c.number_format, c.border = F_NORMAL, fmt, BORDA
        if col != "A":
            c.alignment = Alignment(horizontal="center")
ultima_lista = 9 + N_PROD
ws.conditional_formatting.add(f"C10:C{ultima_lista}", FormulaRule(formula=['$C10="CRÍTICO"'], fill=PatternFill("solid", fgColor="F8D7D7")))
ws.conditional_formatting.add(f"C10:C{ultima_lista}", FormulaRule(formula=['$C10="REPOR"'], fill=PatternFill("solid", fgColor="FFF0C2")))


def estilo(chart, titulo_txt, y_titulo=None):
    chart.title = titulo_txt
    chart.roundedCorners = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    if y_titulo:
        chart.y_axis.title = y_titulo
        chart.y_axis.title.overlay = False
    chart.title.overlay = False
    if chart.legend is not None:
        chart.legend.overlay = False
        chart.legend.position = "b"
    chart.y_axis.majorGridlines.spPr = None


# Gráfico 1: dias de cobertura × ponto de pedido em dias (normaliza as unidades)
wa["Q4"] = "PP em dias"
wa["Q4"].font = Font(name=FONTE, size=8, color="9C9C9C")
for i in range(5, ULT_A + 1):
    wa[f"Q{i}"] = f"=IFERROR(H{i}/D{i},0)"
    wa[f"Q{i}"].number_format = "0.0"
    wa[f"Q{i}"].font = Font(name=FONTE, size=8, color="9C9C9C")
g1 = BarChart()
g1.type = "bar"
g1.add_data(Reference(wa, min_col=11, min_row=4, max_row=ULT_A), titles_from_data=True)
g1.add_data(Reference(wa, min_col=17, min_row=4, max_row=ULT_A), titles_from_data=True)
g1.set_categories(Reference(wa, min_col=2, min_row=5, max_row=ULT_A))
estilo(g1, "Cobertura atual × ponto de pedido (em dias)", "dias")
g1.series[0].graphicalProperties.solidFill = "257DE0"
g1.series[1].graphicalProperties.solidFill = "F47230"
g1.series[0].graphicalProperties.line.noFill = True
g1.series[1].graphicalProperties.line.noFill = True
g1.height, g1.width = 12, 18
ws.add_chart(g1, "H8")

# Gráfico 2: custo diário do consumo
wv = wb["Vendas"]
g2 = LineChart()
g2.add_data(Reference(wv, min_col=N_PROD + 2, min_row=4, max_row=ULT_V), titles_from_data=True)
g2.set_categories(Reference(wv, min_col=1, min_row=5, max_row=ULT_V))
estilo(g2, "Custo diário dos insumos consumidos (R$)")
g2.legend = None
g2.series[0].graphicalProperties.line.solidFill = "257DE0"
g2.series[0].graphicalProperties.line.width = 20000
g2.series[0].smooth = False
g2.x_axis.number_format = "dd/mm"
g2.x_axis.tickLblSkip = 14
g2.height, g2.width = 8, 18
ws.add_chart(g2, "H33")

# Gráfico 3: curva ABC
wc = wb["Curva ABC"]
g3 = BarChart()
g3.add_data(Reference(wc, min_col=3, min_row=4, max_row=ULT_C), titles_from_data=True)
g3.set_categories(Reference(wc, min_col=2, min_row=5, max_row=ULT_C))
estilo(g3, "Curva ABC — consumo mensal por insumo (R$)")
g3.legend = None
g3.series[0].graphicalProperties.solidFill = "257DE0"
g3.height, g3.width = 8, 15
ws.add_chart(g3, "A24")

SAIDA.parent.mkdir(exist_ok=True)
wb.save(SAIDA)
print(f"Planilha salva em {SAIDA}")
