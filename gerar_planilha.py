"""
Gerador de Planilha de Controle Financeiro
Gera um arquivo Excel (.xlsx) com abas mensais e dashboard.
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
import os

# ── Paleta de cores (tema escuro similar ao print) ───────────────────────────
C_BG_DARK    = "1E1E2E"   # fundo principal
C_BG_CARD    = "2A2A3E"   # card/painel
C_BG_HEADER  = "12122A"   # cabeçalho de tabela
C_ACCENT1    = "7C6AF7"   # roxo principal
C_ACCENT2    = "F75C8D"   # rosa/destaque
C_ACCENT3    = "4ECDC4"   # teal
C_ACCENT4    = "FFD93D"   # amarelo
C_ACCENT5    = "6BCB77"   # verde
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

FORMAS_PAGAMENTO = [
    "Débito","Crédito","Pix","Dinheiro","Boleto","Transferência"
]

TIPOS = ["Despesa","Receita","Transferência"]


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, size=11, color=C_TEXT_LIGHT, italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic,
                name="Segoe UI")

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def border_thin(color=C_BORDER):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def border_bottom(color=C_BORDER):
    s = Side(style="thin", color=color)
    return Border(bottom=s)


def set_cell(ws, row, col, value, bold=False, size=11,
             color=C_TEXT_LIGHT, bg=None, h_align="left",
             italic=False, wrap=False, fmt=None, border=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font(bold=bold, size=size, color=color, italic=italic)
    cell.alignment = align(h_align, wrap=wrap)
    if bg:
        cell.fill = fill(bg)
    if fmt:
        cell.number_format = fmt
    if border:
        cell.border = border
    return cell


# ── Aba mensal ────────────────────────────────────────────────────────────────
COLUNAS_MES = [
    ("Data",          12, "DD/MM/YYYY"),
    ("Descrição",     35, "@"),
    ("Categoria",     18, "@"),
    ("Tipo",          14, "@"),
    ("Valor (R$)",    14, '#,##0.00'),
    ("Forma Pagto",   18, "@"),
    ("Parcelas",      10, "0"),
    ("Observações",   30, "@"),
]

def criar_aba_mes(wb, nome_mes, mes_num):
    ws = wb.create_sheet(title=nome_mes)
    ws.sheet_view.showGridLines = False
    ws.tab_color = C_ACCENT1

    # Fundo geral
    ws.sheet_properties.tabColor = C_ACCENT1

    # ── Título ──
    ws.merge_cells("A1:H1")
    set_cell(ws, 1, 1, f"  {nome_mes.upper()} – CONTROLE FINANCEIRO",
             bold=True, size=15, bg=C_BG_HEADER, h_align="left")
    ws.row_dimensions[1].height = 36

    # ── Linha de resumo (fórmulas, preenchidas depois) ──
    ws.merge_cells("A2:B2")
    ws.merge_cells("C2:D2")
    ws.merge_cells("E2:F2")
    ws.merge_cells("G2:H2")

    labels_resumo = [
        (1, "RECEITAS",    C_ACCENT5),
        (3, "DESPESAS",    C_ACCENT2),
        (5, "SALDO",       C_ACCENT1),
        (7, "LANÇAMENTOS", C_ACCENT4),
    ]
    for col, lbl, cor in labels_resumo:
        set_cell(ws, 2, col, lbl, bold=True, size=9,
                 color=cor, bg=C_BG_CARD, h_align="center")
    ws.row_dimensions[2].height = 18

    ws.merge_cells("A3:B3")
    ws.merge_cells("C3:D3")
    ws.merge_cells("E3:F3")
    ws.merge_cells("G3:H3")

    linha_dados_inicio = 6  # dados começam na linha 6

    # Fórmulas de resumo (referem-se à coluna E = Valor)
    # Receitas = soma dos valores onde Tipo = Receita
    # Usamos SUMIF
    resumo_formulas = [
        (1, f'=SUMIF(D{linha_dados_inicio}:D2000,"Receita",E{linha_dados_inicio}:E2000)', C_ACCENT5),
        (3, f'=SUMIF(D{linha_dados_inicio}:D2000,"Despesa",E{linha_dados_inicio}:E2000)', C_ACCENT2),
        (5, f'=C3-E3', C_ACCENT1),
        (7, f'=COUNTA(A{linha_dados_inicio}:A2000)', C_ACCENT4),
    ]
    for col, formula, cor in resumo_formulas:
        c = set_cell(ws, 3, col, formula, bold=True, size=13,
                     color=cor, bg=C_BG_CARD, h_align="center",
                     fmt='#,##0.00')
    # Lançamentos é inteiro
    ws.cell(3, 7).number_format = "0"
    ws.row_dimensions[3].height = 28

    # ── Linha em branco separadora ──
    for col in range(1, 9):
        ws.cell(row=4, column=col).fill = fill(C_BG_DARK)
    ws.row_dimensions[4].height = 8

    # ── Cabeçalho da tabela ──
    ws.row_dimensions[5].height = 24
    for idx, (titulo, largura, _) in enumerate(COLUNAS_MES, start=1):
        set_cell(ws, 5, idx, titulo, bold=True, size=10,
                 color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center",
                 border=border_thin())
        ws.column_dimensions[get_column_letter(idx)].width = largura

    # ── Linhas de dados (2000 linhas pré-formatadas) ──
    for r in range(linha_dados_inicio, 2001):
        bg = C_BG_DARK if r % 2 == 0 else C_BG_CARD
        ws.row_dimensions[r].height = 18
        for idx, (_, _, fmt_str) in enumerate(COLUNAS_MES, start=1):
            c = ws.cell(row=r, column=idx)
            c.fill = fill(bg)
            c.font = font(size=10)
            c.alignment = align("center" if idx in [1,3,4,6,7] else "left")
            c.border = border_bottom()
            if fmt_str not in ("@",):
                c.number_format = fmt_str

    # ── Validações de dados ──
    # Categoria
    dv_cat = DataValidation(
        type="list",
        formula1='"' + ",".join(CATEGORIAS) + '"',
        allow_blank=True,
        showErrorMessage=True,
        error="Selecione uma categoria válida.",
        errorTitle="Categoria inválida"
    )
    dv_cat.sqref = f"C{linha_dados_inicio}:C2000"
    ws.add_data_validation(dv_cat)

    # Tipo
    dv_tipo = DataValidation(
        type="list",
        formula1='"' + ",".join(TIPOS) + '"',
        allow_blank=True
    )
    dv_tipo.sqref = f"D{linha_dados_inicio}:D2000"
    ws.add_data_validation(dv_tipo)

    # Forma de pagamento
    dv_pag = DataValidation(
        type="list",
        formula1='"' + ",".join(FORMAS_PAGAMENTO) + '"',
        allow_blank=True
    )
    dv_pag.sqref = f"F{linha_dados_inicio}:F2000"
    ws.add_data_validation(dv_pag)

    # ── Congelar painel ──
    ws.freeze_panes = "A6"

    return ws


# ── Aba Dashboard ─────────────────────────────────────────────────────────────
def criar_dashboard(wb):
    ws = wb.create_sheet(title="Dashboard", index=0)
    ws.sheet_view.showGridLines = False
    ws.tab_color = C_ACCENT2

    # Larguras de coluna
    col_widths = [2, 18, 14, 14, 14, 14, 14, 14, 2]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Altura das linhas principais
    for r in range(1, 120):
        ws.row_dimensions[r].height = 16

    # ── Fundo ──
    for r in range(1, 120):
        for c in range(1, 10):
            ws.cell(r, c).fill = fill(C_BG_DARK)

    # ── Título principal ──
    ws.merge_cells("B1:H1")
    ws.row_dimensions[1].height = 10

    ws.merge_cells("B2:H3")
    set_cell(ws, 2, 2, "CONTROLE FINANCEIRO 2025", bold=True,
             size=20, color=C_TEXT_LIGHT, bg=C_BG_DARK, h_align="left")
    ws.row_dimensions[2].height = 34

    ws.merge_cells("B4:H4")
    set_cell(ws, 4, 2, "Dashboard Anual  •  Resumo por Categoria  •  Evolução Mensal",
             size=10, color=C_TEXT_DIM, bg=C_BG_DARK, italic=True)
    ws.row_dimensions[4].height = 18

    # ── Cards de resumo anual ──
    # Linha 6-9 : 4 cards lado a lado
    card_defs = [
        ("B", "C", "RECEITA ANUAL",    "Receita",    C_ACCENT5),
        ("D", "E", "DESPESA ANUAL",    "Despesa",    C_ACCENT2),
        ("F", "G", "SALDO ANUAL",      "Saldo",      C_ACCENT1),
        ("H", "H", "MELHOR MÊS",       "MelhorMes",  C_ACCENT4),
    ]

    # Construir fórmulas de receita/despesa somando de todos os meses
    def soma_meses(tipo, col="E"):
        """Gera fórmula SUMIF somando coluna col de todos os meses."""
        partes = []
        for m in MESES:
            safe = m.replace("ç", "ç")  # nome da aba já está correto
            partes.append(
                f"SUMIF('{m}'!D6:D2000,\"{tipo}\",'{m}'!{col}6:{col}2000)"
            )
        return "=" + "+".join(partes)

    receita_formula = soma_meses("Receita")
    despesa_formula = soma_meses("Despesa")
    saldo_formula   = f"={receita_formula[1:]}-{despesa_formula[1:]}"

    # Melhor mês: mês com maior saldo (receita - despesa)
    saldo_por_mes = []
    for m in MESES:
        r = f"SUMIF('{m}'!D6:D2000,\"Receita\",'{m}'!E6:E2000)"
        d = f"SUMIF('{m}'!D6:D2000,\"Despesa\",'{m}'!E6:E2000)"
        saldo_por_mes.append(f"({r}-{d})")
    melhor_mes_formula = (
        '=INDEX({"' + '","'.join(MESES) + '"},'
        'MATCH(MAX(' + ','.join(saldo_por_mes) + '),'
        '{' + ','.join(saldo_por_mes) + '},0))'
    )

    ws.row_dimensions[6].height = 10
    ws.row_dimensions[7].height = 22
    ws.row_dimensions[8].height = 32
    ws.row_dimensions[9].height = 10

    formulas_cards = [receita_formula, despesa_formula, saldo_formula, melhor_mes_formula]
    fmts_cards     = ['#,##0.00', '#,##0.00', '#,##0.00', '@']

    card_cols = [("B","C"), ("D","E"), ("F","G"), ("H","H")]
    card_labels = ["RECEITA ANUAL", "DESPESA ANUAL", "SALDO ANUAL", "MELHOR MÊS"]
    card_cores  = [C_ACCENT5, C_ACCENT2, C_ACCENT1, C_ACCENT4]

    for i, ((c1, c2), label, cor, formula, fmt_card) in enumerate(
            zip(card_cols, card_labels, card_cores, formulas_cards, fmts_cards)):
        ws.merge_cells(f"{c1}7:{c2}7")
        ws.merge_cells(f"{c1}8:{c2}8")
        set_cell(ws, 7, openpyxl.utils.column_index_from_string(c1),
                 label, bold=True, size=8, color=cor, bg=C_BG_CARD, h_align="center")
        c = set_cell(ws, 8, openpyxl.utils.column_index_from_string(c1),
                     formula, bold=True, size=16, color=cor, bg=C_BG_CARD,
                     h_align="center", fmt=fmt_card)

    # ── Tabela: Resumo por Categoria ──
    ws.merge_cells("B11:D11")
    set_cell(ws, 11, 2, "DESPESAS POR CATEGORIA", bold=True, size=11,
             color=C_TEXT_LIGHT, bg=C_BG_CARD, h_align="left")
    ws.row_dimensions[11].height = 26

    headers_cat = ["Categoria", "Valor (R$)", "% Total"]
    for ci, h in enumerate(headers_cat, start=2):
        set_cell(ws, 12, ci, h, bold=True, size=9,
                 color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center",
                 border=border_thin())
    ws.row_dimensions[12].height = 20

    for ri, cat in enumerate(CATEGORIAS, start=13):
        bg = C_BG_DARK if ri % 2 == 0 else C_BG_CARD
        cor = CAT_CORES.get(cat, C_TEXT_DIM)
        set_cell(ws, ri, 2, cat, size=10, color=cor, bg=bg, h_align="left")

        # Soma de todos os meses para esta categoria
        partes_cat = []
        for m in MESES:
            partes_cat.append(
                f"SUMIF('{m}'!C6:C2000,\"{cat}\",'{m}'!E6:E2000)"
            )
        formula_cat = "=" + "+".join(partes_cat)
        set_cell(ws, ri, 3, formula_cat, size=10, color=C_TEXT_LIGHT,
                 bg=bg, h_align="right", fmt='#,##0.00')

        # % do total de despesas
        set_cell(ws, ri, 4,
                 f"=IF({despesa_formula[1:]}=0,0,C{ri}/{despesa_formula[1:]})",
                 size=10, color=C_TEXT_DIM, bg=bg, h_align="right",
                 fmt='0.0%')
        ws.row_dimensions[ri].height = 18

    # Total
    ultima_cat = 13 + len(CATEGORIAS) - 1
    set_cell(ws, ultima_cat+1, 2, "TOTAL", bold=True, size=10,
             color=C_ACCENT2, bg=C_BG_HEADER, h_align="left")
    set_cell(ws, ultima_cat+1, 3, despesa_formula, bold=True, size=10,
             color=C_ACCENT2, bg=C_BG_HEADER, h_align="right", fmt='#,##0.00')
    ws.row_dimensions[ultima_cat+1].height = 20

    # ── Tabela: Evolução mensal ──
    tbl_row = ultima_cat + 3
    ws.merge_cells(f"B{tbl_row}:H{tbl_row}")
    set_cell(ws, tbl_row, 2, "EVOLUÇÃO MENSAL", bold=True, size=11,
             color=C_TEXT_LIGHT, bg=C_BG_CARD, h_align="left")
    ws.row_dimensions[tbl_row].height = 26

    headers_ev = ["Mês", "Receitas", "Despesas", "Saldo", "Economia %"]
    for ci, h in enumerate(headers_ev, start=2):
        set_cell(ws, tbl_row+1, ci, h, bold=True, size=9,
                 color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="center",
                 border=border_thin())
    ws.row_dimensions[tbl_row+1].height = 20

    for mi, mes in enumerate(MESES):
        ri2 = tbl_row + 2 + mi
        bg = C_BG_DARK if mi % 2 == 0 else C_BG_CARD
        f_rec = f"=SUMIF('{mes}'!D6:D2000,\"Receita\",'{mes}'!E6:E2000)"
        f_dep = f"=SUMIF('{mes}'!D6:D2000,\"Despesa\",'{mes}'!E6:E2000)"
        f_sal = f"=C{ri2}-D{ri2}"
        f_eco = f"=IF(C{ri2}=0,0,E{ri2}/C{ri2})"

        set_cell(ws, ri2, 2, mes, size=10, color=C_TEXT_LIGHT, bg=bg)
        set_cell(ws, ri2, 3, f_rec, size=10, color=C_ACCENT5, bg=bg,
                 h_align="right", fmt='#,##0.00')
        set_cell(ws, ri2, 4, f_dep, size=10, color=C_ACCENT2, bg=bg,
                 h_align="right", fmt='#,##0.00')
        set_cell(ws, ri2, 5, f_sal, size=10, color=C_ACCENT1, bg=bg,
                 h_align="right", fmt='#,##0.00')
        set_cell(ws, ri2, 6, f_eco, size=10, color=C_ACCENT4, bg=bg,
                 h_align="right", fmt='0.0%')
        ws.row_dimensions[ri2].height = 18

    # ── Fórmulas de total anual na última linha ──
    ri_tot = tbl_row + 2 + 12
    set_cell(ws, ri_tot, 2, "TOTAL ANUAL", bold=True, size=10,
             color=C_TEXT_LIGHT, bg=C_BG_HEADER)
    set_cell(ws, ri_tot, 3, receita_formula, bold=True, size=10,
             color=C_ACCENT5, bg=C_BG_HEADER, h_align="right", fmt='#,##0.00')
    set_cell(ws, ri_tot, 4, despesa_formula, bold=True, size=10,
             color=C_ACCENT2, bg=C_BG_HEADER, h_align="right", fmt='#,##0.00')
    set_cell(ws, ri_tot, 5, saldo_formula, bold=True, size=10,
             color=C_ACCENT1, bg=C_BG_HEADER, h_align="right", fmt='#,##0.00')
    ws.row_dimensions[ri_tot].height = 22

    ws.freeze_panes = "B6"
    return ws


# ── Aba de instruções ────────────────────────────────────────────────────────
def criar_instrucoes(wb):
    ws = wb.create_sheet(title="ℹ Instruções")
    ws.sheet_view.showGridLines = False
    ws.tab_color = C_ACCENT3

    for r in range(1, 60):
        ws.row_dimensions[r].height = 18
        for c in range(1, 6):
            ws.cell(r, c).fill = fill(C_BG_DARK)

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 3

    ws.merge_cells("B1:D2")
    set_cell(ws, 1, 2, "GUIA DE USO DA PLANILHA", bold=True, size=16,
             color=C_TEXT_LIGHT, bg=C_BG_HEADER, h_align="left")

    instrucoes = [
        ("LANÇAMENTO MANUAL", [
            ("Como lançar",     "Vá à aba do mês desejado e preencha as colunas na primeira linha livre."),
            ("Data",            "Use o formato DD/MM/AAAA. O Excel aceita digitação direta."),
            ("Categoria",       "Escolha da lista suspensa. Novas categorias devem ser adicionadas à lista de validação."),
            ("Tipo",            "Despesa, Receita ou Transferência."),
            ("Valor",           "Digite sempre positivo. O Tipo define se é entrada ou saída."),
            ("Forma de Pagto",  "Débito, Crédito, Pix, Dinheiro, Boleto ou Transferência."),
            ("Parcelas",        "Para compras parceladas, coloque o total de parcelas. Ex: 3 (para 3x)."),
        ]),
        ("IMPORTAÇÃO VIA N8N (PDF)", [
            ("Workflow N8N",    "Use o arquivo n8n_workflow.json incluído neste projeto para configurar a automação."),
            ("Envio do PDF",    "Envie a fatura em PDF pelo webhook configurado no N8N (ex: WhatsApp, e-mail ou upload manual)."),
            ("Resultado",       "O N8N extrai os lançamentos via IA (Claude/OpenAI) e preenche a aba do mês correto."),
            ("Ajuste manual",   "Após importação, revise os lançamentos na aba do mês para corrigir categorizações."),
        ]),
        ("DASHBOARD", [
            ("Atualização",     "O Dashboard atualiza automaticamente via fórmulas — não há necessidade de ação manual."),
            ("Saldo",           "Saldo = soma de Receitas - soma de Despesas do período."),
            ("Economia %",      "Percentual do saldo em relação às receitas do mês."),
        ]),
        ("DICAS", [
            ("Backup",          "Salve uma cópia no OneDrive/Google Drive para histórico."),
            ("Proteção",        "Proteja as abas do Dashboard para evitar edição acidental."),
            ("Cores",           "Cada categoria tem uma cor associada no Dashboard."),
        ]),
    ]

    row = 4
    for secao, itens in instrucoes:
        ws.row_dimensions[row].height = 26
        ws.merge_cells(f"B{row}:D{row}")
        set_cell(ws, row, 2, secao, bold=True, size=11,
                 color=C_ACCENT3, bg=C_BG_CARD, h_align="left")
        row += 1
        for campo, descricao in itens:
            ws.row_dimensions[row].height = 18
            set_cell(ws, row, 2, campo, bold=True, size=10,
                     color=C_ACCENT4, bg=C_BG_DARK, h_align="left")
            set_cell(ws, row, 3, descricao, size=10,
                     color=C_TEXT_LIGHT, bg=C_BG_DARK, h_align="left", wrap=True)
            row += 1
        row += 1  # espaço entre seções

    return ws


# ── Main ─────────────────────────────────────────────────────────────────────
def gerar_planilha(output_path="controle_financeiro_2025.xlsx"):
    wb = openpyxl.Workbook()
    # Remove a aba padrão
    wb.remove(wb.active)

    print("Criando Dashboard...")
    criar_dashboard(wb)

    print("Criando abas mensais...")
    for i, mes in enumerate(MESES, start=1):
        print(f"  {mes}")
        criar_aba_mes(wb, mes, i)

    print("Criando instruções...")
    criar_instrucoes(wb)

    print(f"Salvando em {output_path}...")
    wb.save(output_path)
    print(f"Planilha gerada com sucesso: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    gerar_planilha()
