"""
Gerador de Planilha de Controle Financeiro
Gera controle_financeiro_2025.xlsx com:
  - Dashboard anual + gráfico gastos vs orçamento
  - 12 abas mensais com coluna Cartão, gráfico por categoria
  - Aba Orçamento para definir limites por categoria
  - Aba Instruções
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
import os

# ── Paleta ───────────────────────────────────────────────────────────────────
C_BG_DARK    = "1E1E2E"
C_BG_CARD    = "2A2A3E"
C_BG_HEADER  = "12122A"
C_ACCENT1    = "7C6AF7"
C_ACCENT2    = "F75C8D"
C_ACCENT3    = "4ECDC4"
C_ACCENT4    = "FFD93D"
C_ACCENT5    = "6BCB77"
C_TEXT_LIGHT = "FFFFFF"
C_TEXT_DIM   = "A0A0C0"
C_BORDER     = "3A3A5C"

MESES = [
    "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
    "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"
]

CATEGORIAS = [
    "Alimentação","Moradia","Transporte","Saúde","Educação",
    "Lazer","Vestuário","Assinaturas","Investimentos","Outros"
]

# Cor por categoria (usada no gráfico)
CAT_CORES = {
    "Alimentação":  "F75C8D",
    "Moradia":      "7C6AF7",
    "Transporte":   "4ECDC4",
    "Saúde":        "6BCB77",
    "Educação":     "FFD93D",
    "Lazer":        "FF9A3C",
    "Vestuário":    "A78BFA",
    "Assinaturas":  "F472B6",
    "Investimentos":"34D399",
    "Outros":       "94A3B8",
}

# Cor por banco/cartão (fundo da célula na coluna Cartão)
BANCO_CORES = {
    "nubank":    "6A0DAD",
    "itaú":      "E86213",
    "itau":      "E86213",
    "bradesco":  "CC092F",
    "santander": "EC0000",
    "inter":     "FF7A00",
    "caixa":     "005CA9",
    "btg":       "1E3A5F",
    "c6":        "242424",
    "xp":        "1F1F1F",
    "sicoob":    "00713E",
    "sicredi":   "007A3D",
    "next":      "00B300",
    "picpay":    "21C25E",
}

FORMAS_PAGAMENTO = ["Débito","Crédito","Pix","Dinheiro","Boleto","Transferência"]
TIPOS = ["Despesa","Receita","Transferência"]

# Colunas da aba mensal
# Índice 1-based para referência em fórmulas
COLUNAS_MES = [
    ("Data",         12, "DD/MM/YYYY"),   # col A = 1
    ("Descrição",    35, "@"),             # col B = 2
    ("Categoria",    18, "@"),             # col C = 3
    ("Tipo",         14, "@"),             # col D = 4
    ("Valor (R$)",   14, '#,##0.00'),      # col E = 5
    ("Forma Pagto",  16, "@"),             # col F = 6
    ("Cartão",       16, "@"),             # col G = 7  ← NOVO
    ("Parcelas",     10, "0"),             # col H = 8
    ("Observações",  30, "@"),             # col I = 9
]
N_COLUNAS_DADOS = len(COLUNAS_MES)  # 9
LINHA_DADOS     = 6   # primeira linha de dados
COL_RESUMO      = 11  # coluna K – início da área de resumo/gráfico


# ── Helpers de estilo ─────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, size=11, color=C_TEXT_LIGHT, italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name="Segoe UI")

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def border_thin(color=C_BORDER):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def border_bottom(color=C_BORDER):
    return Border(bottom=Side(style="thin", color=color))

def set_cell(ws, row, col, value, bold=False, size=11, color=C_TEXT_LIGHT,
             bg=None, h_align="left", italic=False, wrap=False,
             fmt=None, border=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = font(bold=bold, size=size, color=color, italic=italic)
    c.alignment = align(h_align, wrap=wrap)
    if bg:    c.fill   = fill(bg)
    if fmt:   c.number_format = fmt
    if border: c.border = border
    return c


# ── Aba mensal ────────────────────────────────────────────────────────────────
def criar_aba_mes(wb, nome_mes, mes_num):
    ws = wb.create_sheet(title=nome_mes)
    ws.sheet_view.showGridLines = False
    ws.tab_color = C_ACCENT1

    # Linha 1 – Título
    ws.merge_cells(f"A1:{get_column_letter(N_COLUNAS_DADOS)}1")
    set_cell(ws, 1, 1, f"  {nome_mes.upper()} – CONTROLE FINANCEIRO",
             bold=True, size=15, bg=C_BG_HEADER)
    ws.row_dimensions[1].height = 36

    # Linhas 2-3 – Cards de resumo
    cards = [
        (1, 2, "RECEITAS",    C_ACCENT5,
         f'=SUMIF(D{LINHA_DADOS}:D2000,"Receita",E{LINHA_DADOS}:E2000)', '#,##0.00'),
        (3, 4, "DESPESAS",    C_ACCENT2,
         f'=SUMIF(D{LINHA_DADOS}:D2000,"Despesa",E{LINHA_DADOS}:E2000)', '#,##0.00'),
        (5, 6, "SALDO",       C_ACCENT1,
         '=C3-E3', '#,##0.00'),
        (7, 8, "LANÇAMENTOS", C_ACCENT4,
         f'=COUNTA(A{LINHA_DADOS}:A2000)', '0'),
        (9, 9, "CARTÕES",     C_ACCENT3,
         f'=COUNTA(UNIQUE(FILTER(G{LINHA_DADOS}:G2000,G{LINHA_DADOS}:G2000<>"")))','0'),
    ]
    for c1, c2, lbl, cor, formula, fmt_card in cards:
        ws.merge_cells(f"{get_column_letter(c1)}2:{get_column_letter(c2)}2")
        ws.merge_cells(f"{get_column_letter(c1)}3:{get_column_letter(c2)}3")
        set_cell(ws, 2, c1, lbl, bold=True, size=8, color=cor, bg=C_BG_CARD, h_align="center")
        set_cell(ws, 3, c1, formula, bold=True, size=13, color=cor, bg=C_BG_CARD,
                 h_align="center", fmt=fmt_card)
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 28

    # Linha 4 – separador
    for col in range(1, N_COLUNAS_DADOS + 1):
        ws.cell(4, col).fill = fill(C_BG_DARK)
    ws.row_dimensions[4].height = 8

    # Linha 5 – Cabeçalhos
    ws.row_dimensions[5].height = 24
    for idx, (titulo, largura, _) in enumerate(COLUNAS_MES, start=1):
        set_cell(ws, 5, idx, titulo, bold=True, size=10,
                 color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center",
                 border=border_thin())
        ws.column_dimensions[get_column_letter(idx)].width = largura

    # Linhas de dados
    for r in range(LINHA_DADOS, 2001):
        bg = C_BG_DARK if r % 2 == 0 else C_BG_CARD
        ws.row_dimensions[r].height = 18
        for idx, (_, _, fmt_str) in enumerate(COLUNAS_MES, start=1):
            c = ws.cell(row=r, column=idx)
            c.fill      = fill(bg)
            c.font      = font(size=10)
            c.alignment = align("center" if idx in [1,3,4,6,7,8] else "left")
            c.border    = border_bottom()
            if fmt_str not in ("@",):
                c.number_format = fmt_str

    # Validações
    dv_cat = DataValidation(type="list",
        formula1='"' + ",".join(CATEGORIAS) + '"', allow_blank=True)
    dv_cat.sqref = f"C{LINHA_DADOS}:C2000"
    ws.add_data_validation(dv_cat)

    dv_tipo = DataValidation(type="list",
        formula1='"' + ",".join(TIPOS) + '"', allow_blank=True)
    dv_tipo.sqref = f"D{LINHA_DADOS}:D2000"
    ws.add_data_validation(dv_tipo)

    dv_pag = DataValidation(type="list",
        formula1='"' + ",".join(FORMAS_PAGAMENTO) + '"', allow_blank=True)
    dv_pag.sqref = f"F{LINHA_DADOS}:F2000"
    ws.add_data_validation(dv_pag)

    # Área de resumo por categoria (colunas K-M) + gráfico
    _adicionar_resumo_e_grafico(ws, nome_mes)

    # Cores por banco na coluna Cartão (G)
    _aplicar_cores_bancos(ws)

    ws.freeze_panes = "A6"
    return ws


def _aplicar_cores_bancos(ws):
    """Formatação condicional: pinta a célula da coluna Cartão com a cor do banco."""
    # Texto branco em todos os casos
    font_branco = Font(color="FFFFFF", bold=True, name="Segoe UI", size=10)
    # Ordem importa: bancos com nomes mais específicos primeiro
    bancos_ordenados = [
        ("c6",        "242424"),
        ("btg",       "1E3A5F"),
        ("xp",        "1F1F1F"),
        ("nubank",    "6A0DAD"),
        ("bradesco",  "CC092F"),
        ("santander", "EC0000"),
        ("sicoob",    "00713E"),
        ("sicredi",   "007A3D"),
        ("picpay",    "21C25E"),
        ("next",      "00B300"),
        ("inter",     "FF7A00"),
        ("caixa",     "005CA9"),
        ("itaú",      "E86213"),
        ("itau",      "E86213"),
    ]
    col_cartao = LINHA_DADOS  # primeira linha de dados (usada como âncora na fórmula)
    faixa = f"G{LINHA_DADOS}:G2000"
    for banco_key, hex_color in bancos_ordenados:
        # SEARCH é case-insensitive; a fórmula é relativa à primeira célula da faixa
        formula = [f'=ISNUMBER(SEARCH("{banco_key}",G{LINHA_DADOS}))']
        rule = FormulaRule(
            formula=formula,
            fill=PatternFill("solid", fgColor=hex_color),
            font=font_branco,
            stopIfTrue=False,
        )
        ws.conditional_formatting.add(faixa, rule)


def _adicionar_resumo_e_grafico(ws, nome_mes):
    """Insere tabela de resumo por categoria e gráfico de barras na aba."""
    col_cat  = COL_RESUMO       # K = 11
    col_gst  = COL_RESUMO + 1   # L
    col_orc  = COL_RESUMO + 2   # M
    col_pct  = COL_RESUMO + 3   # N

    # Larguras
    ws.column_dimensions[get_column_letter(col_cat)].width  = 2   # espaço
    ws.column_dimensions[get_column_letter(col_cat+1)].width = 18
    ws.column_dimensions[get_column_letter(col_gst+1)].width = 14
    ws.column_dimensions[get_column_letter(col_orc+1)].width = 14
    ws.column_dimensions[get_column_letter(col_pct+1)].width = 10

    # Cabeçalho da tabela de resumo (linha 2)
    set_cell(ws, 2, col_cat+1, "CATEGORIA",   bold=True, size=9,
             color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center")
    set_cell(ws, 2, col_gst+1, "GASTO (R$)",  bold=True, size=9,
             color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center")
    set_cell(ws, 2, col_orc+1, "ORÇAMENTO",   bold=True, size=9,
             color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center")
    set_cell(ws, 2, col_pct+1, "USO %",       bold=True, size=9,
             color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center")

    # Linhas por categoria
    for i, cat in enumerate(CATEGORIAS):
        row = 3 + i
        cor = CAT_CORES.get(cat, C_TEXT_DIM)
        bg  = C_BG_DARK if i % 2 == 0 else C_BG_CARD

        set_cell(ws, row, col_cat+1, cat,  size=10, color=cor, bg=bg)
        set_cell(ws, row, col_gst+1,
                 f'=SUMIF(C{LINHA_DADOS}:C2000,"{cat}",E{LINHA_DADOS}:E2000)',
                 size=10, color=C_TEXT_LIGHT, bg=bg,
                 h_align="right", fmt='#,##0.00')
        set_cell(ws, row, col_orc+1,
                 f'=IFERROR(VLOOKUP("{cat}",\'Orçamento\'!$A:$B,2,0),0)',
                 size=10, color=C_TEXT_DIM, bg=bg,
                 h_align="right", fmt='#,##0.00')
        set_cell(ws, row, col_pct+1,
                 f'=IF({get_column_letter(col_orc+1)}{row}=0,0,{get_column_letter(col_gst+1)}{row}/{get_column_letter(col_orc+1)}{row})',
                 size=10, color=C_ACCENT4, bg=bg,
                 h_align="right", fmt='0%')
        ws.row_dimensions[row].height = 18

    # Gráfico de barras agrupado: Gasto vs Orçamento por categoria
    chart = BarChart()
    chart.type      = "col"
    chart.grouping  = "clustered"
    chart.title     = f"Gastos vs Orçamento – {nome_mes}"
    chart.y_axis.title = "R$"
    chart.style     = 2
    chart.width     = 22
    chart.height    = 14

    n_cats = len(CATEGORIAS)
    row_ini = 3
    row_fim = row_ini + n_cats - 1

    cats_ref = Reference(ws, min_col=col_cat+1, min_row=row_ini, max_row=row_fim)
    gasto_ref = Reference(ws, min_col=col_gst+1, min_row=2, max_row=row_fim)
    orc_ref   = Reference(ws, min_col=col_orc+1, min_row=2, max_row=row_fim)

    chart.add_data(gasto_ref, titles_from_data=True)
    chart.add_data(orc_ref,   titles_from_data=True)
    chart.set_categories(cats_ref)

    # Cores das séries
    chart.series[0].graphicalProperties.solidFill = "F75C8D"  # Gasto
    chart.series[0].graphicalProperties.line.solidFill = "F75C8D"
    chart.series[1].graphicalProperties.solidFill = "7C6AF7"  # Orçamento
    chart.series[1].graphicalProperties.line.solidFill = "7C6AF7"

    anchor_col = get_column_letter(col_cat + 1)
    ws.add_chart(chart, f"{anchor_col}14")


# ── Aba Orçamento ─────────────────────────────────────────────────────────────
ORCAMENTO_PADRAO = {
    "Alimentação":  800,
    "Moradia":      2000,
    "Transporte":   500,
    "Saúde":        400,
    "Educação":     300,
    "Lazer":        400,
    "Vestuário":    200,
    "Assinaturas":  150,
    "Investimentos":1000,
    "Outros":       200,
}

def criar_orcamento(wb):
    ws = wb.create_sheet(title="Orçamento")
    ws.sheet_view.showGridLines = False
    ws.tab_color = C_ACCENT4

    for r in range(1, 30):
        ws.row_dimensions[r].height = 18
        for c in range(1, 7):
            ws.cell(r, c).fill = fill(C_BG_DARK)

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 3

    # Título
    ws.merge_cells("B1:E2")
    set_cell(ws, 1, 2, "DEFINIÇÃO DE ORÇAMENTO MENSAL",
             bold=True, size=15, bg=C_BG_HEADER)
    ws.row_dimensions[1].height = 34

    # Cabeçalho
    headers = ["Categoria", "Limite Mensal (R$)", "Gasto Atual (R$)", "Uso %"]
    for ci, h in enumerate(headers, start=2):
        set_cell(ws, 3, ci, h, bold=True, size=10,
                 color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center",
                 border=border_thin())
    ws.row_dimensions[3].height = 22

    # Linhas por categoria
    for i, cat in enumerate(CATEGORIAS):
        row = 4 + i
        bg  = C_BG_DARK if i % 2 == 0 else C_BG_CARD
        cor = CAT_CORES.get(cat, C_TEXT_DIM)
        limite = ORCAMENTO_PADRAO.get(cat, 0)

        set_cell(ws, row, 2, cat,   size=10, color=cor, bg=bg)
        set_cell(ws, row, 3, limite, size=11, color=C_ACCENT4, bg=bg,
                 bold=True, h_align="right", fmt='#,##0.00')

        # Gasto atual = soma de todos os meses (referência cruzada)
        partes = []
        for m in MESES:
            partes.append(f'SUMIF(\'{m}\'!C{LINHA_DADOS}:C2000,"{cat}",\'{m}\'!E{LINHA_DADOS}:E2000)')
        formula_gasto = "=" + "+".join(partes)
        set_cell(ws, row, 4, formula_gasto, size=10, color=C_ACCENT2, bg=bg,
                 h_align="right", fmt='#,##0.00')

        # Uso %
        set_cell(ws, row, 5,
                 f'=IF(C{row}=0,0,D{row}/C{row})',
                 size=10, color=C_ACCENT3, bg=bg,
                 h_align="right", fmt='0%')
        ws.row_dimensions[row].height = 20

    # Nota explicativa
    row_nota = 4 + len(CATEGORIAS) + 1
    ws.merge_cells(f"B{row_nota}:E{row_nota}")
    set_cell(ws, row_nota, 2,
             "  ⚠  Edite os valores em 'Limite Mensal' para personalizar seus limites de gastos.",
             size=9, color=C_ACCENT4, bg=C_BG_CARD, italic=True, wrap=True)
    ws.row_dimensions[row_nota].height = 24

    return ws


# ── Dashboard ─────────────────────────────────────────────────────────────────
def criar_dashboard(wb):
    ws = wb.create_sheet(title="Dashboard", index=0)
    ws.sheet_view.showGridLines = False
    ws.tab_color = C_ACCENT2

    col_widths = [2, 18, 14, 14, 14, 14, 14, 14, 2]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for r in range(1, 130):
        ws.row_dimensions[r].height = 16
        for c in range(1, 10):
            ws.cell(r, c).fill = fill(C_BG_DARK)

    # Título
    ws.merge_cells("B2:H3")
    set_cell(ws, 2, 2, "CONTROLE FINANCEIRO 2025", bold=True,
             size=20, color=C_TEXT_LIGHT, bg=C_BG_DARK)
    ws.row_dimensions[2].height = 34

    ws.merge_cells("B4:H4")
    set_cell(ws, 4, 2,
             "Dashboard Anual  •  Resumo por Categoria  •  Evolução Mensal",
             size=10, color=C_TEXT_DIM, bg=C_BG_DARK, italic=True)
    ws.row_dimensions[4].height = 18

    # ── Fórmulas anuais ──
    def soma_meses(tipo):
        return "=" + "+".join(
            f'SUMIF(\'{m}\'!D{LINHA_DADOS}:D2000,"{tipo}",\'{m}\'!E{LINHA_DADOS}:E2000)'
            for m in MESES
        )
    receita_f  = soma_meses("Receita")
    despesa_f  = soma_meses("Despesa")
    saldo_f    = f"={receita_f[1:]}-{despesa_f[1:]}"

    saldo_por_mes = []
    for m in MESES:
        r = f'SUMIF(\'{m}\'!D{LINHA_DADOS}:D2000,"Receita",\'{m}\'!E{LINHA_DADOS}:E2000)'
        d = f'SUMIF(\'{m}\'!D{LINHA_DADOS}:D2000,"Despesa",\'{m}\'!E{LINHA_DADOS}:E2000)'
        saldo_por_mes.append(f"({r}-{d})")
    melhor_f = (
        '=INDEX({"' + '","'.join(MESES) + '"},'
        'MATCH(MAX(' + ','.join(saldo_por_mes) + '),'
        '{' + ','.join(saldo_por_mes) + '},0))'
    )

    # Cards resumo
    ws.row_dimensions[6].height = 10
    ws.row_dimensions[7].height = 22
    ws.row_dimensions[8].height = 32
    ws.row_dimensions[9].height = 10

    cards = [
        ("B","C", "RECEITA ANUAL",  receita_f,  C_ACCENT5, '#,##0.00'),
        ("D","E", "DESPESA ANUAL",  despesa_f,  C_ACCENT2, '#,##0.00'),
        ("F","G", "SALDO ANUAL",    saldo_f,    C_ACCENT1, '#,##0.00'),
        ("H","H", "MELHOR MÊS",     melhor_f,   C_ACCENT4, '@'),
    ]
    for c1, c2, lbl, formula, cor, fmt_c in cards:
        ci = column_index_from_string(c1)
        ws.merge_cells(f"{c1}7:{c2}7")
        ws.merge_cells(f"{c1}8:{c2}8")
        set_cell(ws, 7, ci, lbl, bold=True, size=8, color=cor,
                 bg=C_BG_CARD, h_align="center")
        set_cell(ws, 8, ci, formula, bold=True, size=16, color=cor,
                 bg=C_BG_CARD, h_align="center", fmt=fmt_c)

    # ── Tabela: Despesas por Categoria vs Orçamento ──
    ws.merge_cells("B11:E11")
    set_cell(ws, 11, 2, "DESPESAS POR CATEGORIA vs ORÇAMENTO",
             bold=True, size=11, color=C_TEXT_LIGHT, bg=C_BG_CARD)
    ws.row_dimensions[11].height = 26

    hdrs = ["Categoria", "Gasto (R$)", "Orçamento", "Uso %"]
    for ci, h in enumerate(hdrs, start=2):
        set_cell(ws, 12, ci, h, bold=True, size=9,
                 color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center",
                 border=border_thin())
    ws.row_dimensions[12].height = 20

    cat_rows = {}
    for ri, cat in enumerate(CATEGORIAS, start=13):
        bg  = C_BG_DARK if ri % 2 == 0 else C_BG_CARD
        cor = CAT_CORES.get(cat, C_TEXT_DIM)
        partes_cat = [
            f'SUMIF(\'{m}\'!C{LINHA_DADOS}:C2000,"{cat}",\'{m}\'!E{LINHA_DADOS}:E2000)'
            for m in MESES
        ]
        gasto_f = "=" + "+".join(partes_cat)
        orc_f   = f'=IFERROR(VLOOKUP("{cat}",\'Orçamento\'!$B:$C,2,0),0)'

        set_cell(ws, ri, 2, cat,    size=10, color=cor,         bg=bg)
        set_cell(ws, ri, 3, gasto_f,size=10, color=C_TEXT_LIGHT,bg=bg,
                 h_align="right", fmt='#,##0.00')
        set_cell(ws, ri, 4, orc_f,  size=10, color=C_TEXT_DIM,  bg=bg,
                 h_align="right", fmt='#,##0.00')
        set_cell(ws, ri, 5,
                 f'=IF(D{ri}=0,0,C{ri}/D{ri})',
                 size=10, color=C_ACCENT4, bg=bg,
                 h_align="right", fmt='0%')
        ws.row_dimensions[ri].height = 18
        cat_rows[cat] = ri

    ultima_cat = 13 + len(CATEGORIAS) - 1
    set_cell(ws, ultima_cat+1, 2, "TOTAL", bold=True, size=10,
             color=C_ACCENT2, bg=C_BG_HEADER)
    set_cell(ws, ultima_cat+1, 3, despesa_f, bold=True, size=10,
             color=C_ACCENT2, bg=C_BG_HEADER, h_align="right", fmt='#,##0.00')
    ws.row_dimensions[ultima_cat+1].height = 22

    # Gráfico: gastos vs orçamento por categoria (dashboard)
    chart_dash = BarChart()
    chart_dash.type     = "col"
    chart_dash.grouping = "clustered"
    chart_dash.title    = "Gastos vs Orçamento por Categoria (Anual)"
    chart_dash.style    = 2
    chart_dash.width    = 26
    chart_dash.height   = 14

    cats_ref_d = Reference(ws, min_col=2, min_row=13, max_row=ultima_cat)
    gasto_ref_d = Reference(ws, min_col=3, min_row=12, max_row=ultima_cat)
    orc_ref_d   = Reference(ws, min_col=4, min_row=12, max_row=ultima_cat)
    chart_dash.add_data(gasto_ref_d, titles_from_data=True)
    chart_dash.add_data(orc_ref_d,   titles_from_data=True)
    chart_dash.set_categories(cats_ref_d)
    chart_dash.series[0].graphicalProperties.solidFill = "F75C8D"
    chart_dash.series[0].graphicalProperties.line.solidFill = "F75C8D"
    chart_dash.series[1].graphicalProperties.solidFill = "7C6AF7"
    chart_dash.series[1].graphicalProperties.line.solidFill = "7C6AF7"
    ws.add_chart(chart_dash, "F11")

    # ── Tabela: Evolução mensal ──
    tbl_row = ultima_cat + 3
    ws.merge_cells(f"B{tbl_row}:H{tbl_row}")
    set_cell(ws, tbl_row, 2, "EVOLUÇÃO MENSAL",
             bold=True, size=11, color=C_TEXT_LIGHT, bg=C_BG_CARD)
    ws.row_dimensions[tbl_row].height = 26

    ev_hdrs = ["Mês", "Receitas", "Despesas", "Saldo", "Economia %"]
    for ci, h in enumerate(ev_hdrs, start=2):
        set_cell(ws, tbl_row+1, ci, h, bold=True, size=9,
                 color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center",
                 border=border_thin())
    ws.row_dimensions[tbl_row+1].height = 20

    for mi, mes in enumerate(MESES):
        ri2 = tbl_row + 2 + mi
        bg  = C_BG_DARK if mi % 2 == 0 else C_BG_CARD
        f_rec = f'=SUMIF(\'{mes}\'!D{LINHA_DADOS}:D2000,"Receita",\'{mes}\'!E{LINHA_DADOS}:E2000)'
        f_dep = f'=SUMIF(\'{mes}\'!D{LINHA_DADOS}:D2000,"Despesa",\'{mes}\'!E{LINHA_DADOS}:E2000)'
        f_sal = f"=C{ri2}-D{ri2}"
        f_eco = f"=IF(C{ri2}=0,0,E{ri2}/C{ri2})"
        set_cell(ws, ri2, 2, mes,   size=10, color=C_TEXT_LIGHT, bg=bg)
        set_cell(ws, ri2, 3, f_rec, size=10, color=C_ACCENT5,    bg=bg,
                 h_align="right", fmt='#,##0.00')
        set_cell(ws, ri2, 4, f_dep, size=10, color=C_ACCENT2,    bg=bg,
                 h_align="right", fmt='#,##0.00')
        set_cell(ws, ri2, 5, f_sal, size=10, color=C_ACCENT1,    bg=bg,
                 h_align="right", fmt='#,##0.00')
        set_cell(ws, ri2, 6, f_eco, size=10, color=C_ACCENT4,    bg=bg,
                 h_align="right", fmt='0.0%')
        ws.row_dimensions[ri2].height = 18

    ri_tot = tbl_row + 2 + 12
    for ci, (val, cor, fmt_t) in enumerate([
        ("TOTAL ANUAL", C_TEXT_LIGHT, "@"),
        (receita_f, C_ACCENT5, '#,##0.00'),
        (despesa_f, C_ACCENT2, '#,##0.00'),
        (saldo_f,   C_ACCENT1, '#,##0.00'),
    ], start=2):
        set_cell(ws, ri_tot, ci, val, bold=True, size=10,
                 color=cor, bg=C_BG_HEADER, h_align="right" if ci > 2 else "left",
                 fmt=fmt_t)
    ws.row_dimensions[ri_tot].height = 22

    # ── Gráfico de torres: Receitas vs Despesas por mês ──
    # (linhas tbl_row+2 .. ri_tot-1 contêm Mês, Receitas, Despesas, Saldo, Economia%)
    ev_row_ini = tbl_row + 2
    ev_row_fim = ri_tot - 1  # exclui a linha TOTAL

    chart_mensal = BarChart()
    chart_mensal.type      = "col"
    chart_mensal.grouping  = "clustered"
    chart_mensal.title     = "Receitas vs Despesas por Mês"
    chart_mensal.y_axis.title = "R$"
    chart_mensal.style     = 2
    chart_mensal.width     = 30
    chart_mensal.height    = 14

    meses_ref  = Reference(ws, min_col=2, min_row=ev_row_ini, max_row=ev_row_fim)
    rec_ref    = Reference(ws, min_col=3, min_row=ev_row_ini - 1, max_row=ev_row_fim)
    dep_ref    = Reference(ws, min_col=4, min_row=ev_row_ini - 1, max_row=ev_row_fim)
    saldo_ref2 = Reference(ws, min_col=5, min_row=ev_row_ini - 1, max_row=ev_row_fim)

    chart_mensal.add_data(rec_ref,    titles_from_data=True)
    chart_mensal.add_data(dep_ref,    titles_from_data=True)
    chart_mensal.add_data(saldo_ref2, titles_from_data=True)
    chart_mensal.set_categories(meses_ref)

    chart_mensal.series[0].graphicalProperties.solidFill = "6BCB77"   # Receitas – verde
    chart_mensal.series[0].graphicalProperties.line.solidFill = "6BCB77"
    chart_mensal.series[1].graphicalProperties.solidFill = "F75C8D"   # Despesas – rosa
    chart_mensal.series[1].graphicalProperties.line.solidFill = "F75C8D"
    chart_mensal.series[2].graphicalProperties.solidFill = "7C6AF7"   # Saldo – roxo
    chart_mensal.series[2].graphicalProperties.line.solidFill = "7C6AF7"

    chart_anchor = f"B{ri_tot + 2}"
    ws.add_chart(chart_mensal, chart_anchor)

    ws.freeze_panes = "B6"
    return ws


# ── Aba Instruções ────────────────────────────────────────────────────────────
def criar_instrucoes(wb):
    ws = wb.create_sheet(title="ℹ Instruções")
    ws.sheet_view.showGridLines = False
    ws.tab_color = C_ACCENT3

    for r in range(1, 70):
        ws.row_dimensions[r].height = 18
        for c in range(1, 6):
            ws.cell(r, c).fill = fill(C_BG_DARK)

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 3

    ws.merge_cells("B1:D2")
    set_cell(ws, 1, 2, "GUIA DE USO DA PLANILHA", bold=True, size=16,
             color=C_TEXT_LIGHT, bg=C_BG_HEADER)

    instrucoes = [
        ("LANÇAMENTO MANUAL", [
            ("Data",            "Formato DD/MM/AAAA. Pode digitar diretamente."),
            ("Categoria",       "Lista suspensa. Defina limites na aba Orçamento."),
            ("Tipo",            "Despesa / Receita / Transferência."),
            ("Valor",           "Sempre positivo. O Tipo define entrada ou saída."),
            ("Cartão",          "Nome do banco/cartão (Nubank, Itaú, etc.). Preenchido automaticamente pelo N8N."),
            ("Forma de Pagto",  "Débito, Crédito, Pix, Dinheiro, Boleto."),
        ]),
        ("IMPORTAÇÃO VIA N8N (WhatsApp ou Formulário)", [
            ("Como enviar",     "Envie o PDF da fatura pelo WhatsApp configurado OU acesse o link do formulário N8N."),
            ("O que acontece",  "Claude lê a fatura, categoriza os lançamentos e preenche o Google Sheets automaticamente."),
            ("Cartão",          "O banco é detectado automaticamente e a coluna Cartão é preenchida."),
            ("Ajuste manual",   "Após importação, revise a aba do mês para corrigir categorizações se necessário."),
        ]),
        ("ORÇAMENTO E ALERTAS", [
            ("Definir limites",  "Vá na aba 'Orçamento' e edite os valores na coluna 'Limite Mensal'."),
            ("Alertas Wpp",      "O N8N verifica os limites após cada importação e avisa via WhatsApp se ultrapassar."),
            ("Agente diário",    "Todo dia às 22h o agente IA analisa os gastos e envia um resumo no WhatsApp."),
        ]),
        ("DASHBOARD E GRÁFICOS", [
            ("Dashboard",        "Atualiza automaticamente via fórmulas. Mostra gastos vs orçamento por categoria."),
            ("Gráfico mensal",   "Cada aba de mês tem um gráfico Gastos vs Orçamento por categoria (colunas K em diante)."),
            ("Cores dos bancos", "Nubank = roxo, Itaú = laranja, Bradesco = vermelho, Inter = laranja escuro, etc."),
        ]),
    ]

    row = 4
    for secao, itens in instrucoes:
        ws.row_dimensions[row].height = 26
        ws.merge_cells(f"B{row}:D{row}")
        set_cell(ws, row, 2, secao, bold=True, size=11,
                 color=C_ACCENT3, bg=C_BG_CARD)
        row += 1
        for campo, descricao in itens:
            ws.row_dimensions[row].height = 18
            set_cell(ws, row, 2, campo,    bold=True, size=10,
                     color=C_ACCENT4, bg=C_BG_DARK)
            set_cell(ws, row, 3, descricao, size=10,
                     color=C_TEXT_LIGHT, bg=C_BG_DARK, wrap=True)
            row += 1
        row += 1
    return ws


# ── Main ─────────────────────────────────────────────────────────────────────
def gerar_planilha(output_path="controle_financeiro_2025.xlsx"):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    print("Dashboard...")
    criar_dashboard(wb)

    print("Abas mensais...")
    for i, mes in enumerate(MESES, start=1):
        print(f"  {mes}")
        criar_aba_mes(wb, mes, i)

    print("Orçamento...")
    criar_orcamento(wb)

    print("Instruções...")
    criar_instrucoes(wb)

    print(f"Salvando {output_path}...")
    wb.save(output_path)
    print(f"✓ {os.path.abspath(output_path)}")


if __name__ == "__main__":
    gerar_planilha()
