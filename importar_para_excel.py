"""
Importador de Lançamentos para Excel Local
Recebe o JSON retornado pelo webhook N8N e insere os lançamentos
diretamente no arquivo .xlsx gerado pelo gerar_planilha.py

Uso:
  # Opção 1 – via arquivo JSON
  python importar_para_excel.py resposta_n8n.json

  # Opção 2 – via pipe (saída do curl)
  curl -s -X POST http://localhost:5678/webhook/fatura-pdf \
       -F arquivo=@fatura.pdf | python importar_para_excel.py

  # Opção 3 – processar PDF direto (sem N8N, usando Claude API local)
  python importar_para_excel.py --pdf fatura.pdf --api-key sk-ant-...
"""

import sys
import json
import os
import argparse
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("Erro: instale as dependências com 'pip install -r requirements.txt'")
    sys.exit(1)

PLANILHA_PADRAO = "controle_financeiro_2025.xlsx"

COLUNAS = [
    "Data", "Descrição", "Categoria", "Tipo",
    "Valor (R$)", "Forma Pagto", "Parcelas", "Observações"
]

COL_MAP = {
    "data":            1,
    "descricao":       2,
    "categoria":       3,
    "tipo":            4,
    "valor":           5,
    "forma_pagamento": 6,
    "parcelas":        7,
    "observacoes":     8,
}

LINHA_INICIO_DADOS = 6  # mesma do gerar_planilha.py


def encontrar_proxima_linha(ws):
    """Retorna a próxima linha vazia na coluna A a partir de LINHA_INICIO_DADOS."""
    for row in range(LINHA_INICIO_DADOS, 2001):
        if ws.cell(row=row, column=1).value is None:
            return row
    raise ValueError("Planilha cheia! Todas as 2000 linhas estão preenchidas.")


def inserir_lancamentos(planilha_path, mes, lancamentos):
    if not os.path.exists(planilha_path):
        print(f"Erro: arquivo '{planilha_path}' não encontrado.")
        print("Execute primeiro: python gerar_planilha.py")
        sys.exit(1)

    wb = openpyxl.load_workbook(planilha_path)

    if mes not in wb.sheetnames:
        print(f"Erro: aba '{mes}' não encontrada na planilha.")
        print(f"Abas disponíveis: {', '.join(wb.sheetnames)}")
        sys.exit(1)

    ws = wb[mes]
    linha = encontrar_proxima_linha(ws)
    inseridos = 0
    ignorados = 0

    for l in lancamentos:
        # Validar campos obrigatórios
        if not l.get("data") or not l.get("descricao"):
            ignorados += 1
            continue

        # Converter data
        data_str = l.get("data", "")
        try:
            data_val = datetime.strptime(data_str, "%d/%m/%Y").date()
        except ValueError:
            data_val = data_str  # mantém string se não puder converter

        ws.cell(linha, COL_MAP["data"]).value            = data_val
        ws.cell(linha, COL_MAP["descricao"]).value       = str(l.get("descricao", ""))[:50]
        ws.cell(linha, COL_MAP["categoria"]).value       = l.get("categoria", "Outros")
        ws.cell(linha, COL_MAP["tipo"]).value            = l.get("tipo", "Despesa")
        ws.cell(linha, COL_MAP["valor"]).value           = float(l.get("valor", 0))
        ws.cell(linha, COL_MAP["forma_pagamento"]).value = l.get("forma_pagamento", "")
        ws.cell(linha, COL_MAP["parcelas"]).value        = int(l.get("parcelas", 1))
        ws.cell(linha, COL_MAP["observacoes"]).value     = str(l.get("observacoes", ""))

        # Formatar data
        ws.cell(linha, COL_MAP["data"]).number_format = "DD/MM/YYYY"
        ws.cell(linha, COL_MAP["valor"]).number_format = '#,##0.00'

        linha += 1
        inseridos += 1

    wb.save(planilha_path)
    return inseridos, ignorados


def processar_pdf_direto(pdf_path, api_key):
    """Processa PDF diretamente sem N8N, usando a API do Claude."""
    try:
        import anthropic
        import base64
    except ImportError:
        print("Erro: instale 'anthropic' com 'pip install -r requirements.txt'")
        sys.exit(1)

    print(f"Lendo PDF: {pdf_path}")
    with open(pdf_path, "rb") as f:
        pdf_data = base64.standard_b64encode(f.read()).decode("utf-8")

    client = anthropic.Anthropic(api_key=api_key)

    prompt = """Você é um assistente especialista em análise de faturas bancárias e de cartão de crédito brasileiras.

Analise o PDF desta fatura e extraia TODOS os lançamentos financeiros.

Para cada lançamento retorne um JSON no array "lancamentos" com os campos:
- data: string no formato DD/MM/AAAA
- descricao: string (nome do estabelecimento/descrição limpa, máx 50 chars)
- categoria: uma das opções: Alimentação, Moradia, Transporte, Saúde, Educação, Lazer, Vestuário, Assinaturas, Investimentos, Outros
- tipo: "Despesa" ou "Receita"
- valor: número positivo (sem R$, sem vírgula, use ponto decimal)
- forma_pagamento: "Crédito", "Débito", "Pix", "Dinheiro", "Boleto" ou "Transferência"
- parcelas: número inteiro (1 se não for parcelado)
- observacoes: string com informações adicionais ou string vazia

Regras: Ignore totais, subtotais, pagamento mínimo, saldo anterior. Retorne SOMENTE o JSON.

Formato:
{
  "mes": "NomeDomes",
  "banco": "NomeDoBanco",
  "periodo": "DD/MM/AAAA a DD/MM/AAAA",
  "total_despesas": 0.00,
  "total_receitas": 0.00,
  "lancamentos": [...]
}"""

    print("Enviando para Claude API...")
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=8000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_data,
                    },
                },
                {"type": "text", "text": prompt}
            ],
        }]
    )

    resposta = message.content[0].text
    import re
    json_match = re.search(r'\{[\s\S]*\}', resposta)
    if not json_match:
        raise ValueError("JSON não encontrado na resposta da API")

    return json.loads(json_match.group(0))


def main():
    parser = argparse.ArgumentParser(
        description="Importa lançamentos de faturas para a planilha Excel de controle financeiro."
    )
    parser.add_argument("json_file", nargs="?",
                        help="Arquivo JSON com resposta do N8N (ou stdin se omitido)")
    parser.add_argument("--planilha", default=PLANILHA_PADRAO,
                        help=f"Caminho da planilha Excel (padrão: {PLANILHA_PADRAO})")
    parser.add_argument("--pdf", help="Processa PDF diretamente sem N8N")
    parser.add_argument("--api-key", help="Chave API Anthropic (para --pdf)")
    parser.add_argument("--mes", help="Nome do mês para sobrescrever o detectado (ex: Janeiro)")

    args = parser.parse_args()

    # ── Obter dados ──
    if args.pdf:
        api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            print("Erro: informe a chave API com --api-key ou variável ANTHROPIC_API_KEY")
            sys.exit(1)
        dados = processar_pdf_direto(args.pdf, api_key)
    elif args.json_file:
        with open(args.json_file, "r", encoding="utf-8") as f:
            dados = json.load(f)
    elif not sys.stdin.isatty():
        dados = json.load(sys.stdin)
    else:
        parser.print_help()
        sys.exit(0)

    # ── Extrair lançamentos ──
    # Suporta tanto o formato do webhook N8N quanto o do processamento direto
    lancamentos = dados.get("lancamentos", [])
    mes = args.mes or dados.get("mes", "")

    if not mes:
        print("Erro: mês não identificado. Use --mes NomeDoMes para forçar.")
        sys.exit(1)

    if not lancamentos:
        print("Nenhum lançamento encontrado no JSON.")
        sys.exit(0)

    print(f"\nImportando {len(lancamentos)} lançamentos para aba '{mes}'...")
    inseridos, ignorados = inserir_lancamentos(args.planilha, mes, lancamentos)

    print(f"\n{'='*50}")
    print(f"  Planilha : {args.planilha}")
    print(f"  Mes      : {mes}")
    print(f"  Inseridos: {inseridos}")
    print(f"  Ignorados: {ignorados}")
    if dados.get("banco"):
        print(f"  Banco    : {dados['banco']}")
    if dados.get("total_despesas"):
        print(f"  Despesas : R$ {dados['total_despesas']}")
    if dados.get("total_receitas"):
        print(f"  Receitas : R$ {dados['total_receitas']}")
    print(f"{'='*50}")
    print("Importação concluída! Abra a planilha para revisar os lançamentos.")


if __name__ == "__main__":
    main()
