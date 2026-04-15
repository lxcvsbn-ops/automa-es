# Controle Financeiro – Planilha + Automação N8N

Sistema de controle financeiro pessoal com:
- Planilha Excel (`.xlsx`) com abas mensais e dashboard automático
- Workflow N8N para importar faturas em PDF automaticamente via IA (Claude)
- Script Python para importação local sem N8N

---

## Estrutura

```
.
├── gerar_planilha.py       # Gera o arquivo Excel
├── importar_para_excel.py  # Importa JSON de lançamentos para o Excel
├── n8n_workflow.json       # Workflow N8N pronto para importar
├── requirements.txt
└── README.md
```

---

## 1. Gerar a Planilha Excel

```bash
pip install -r requirements.txt
python gerar_planilha.py
# Gera: controle_financeiro_2025.xlsx
```

A planilha contém:
- **Dashboard** – resumo anual, evolução mensal, despesas por categoria (atualização automática via fórmulas)
- **Janeiro … Dezembro** – abas de lançamento com listas suspensas validadas
- **Instruções** – guia de uso

### Colunas de cada aba mensal

| Coluna | Descrição |
|---|---|
| Data | DD/MM/AAAA |
| Descrição | Nome do estabelecimento |
| Categoria | Lista: Alimentação, Moradia, Transporte… |
| Tipo | Despesa / Receita / Transferência |
| Valor (R$) | Sempre positivo |
| Forma Pagto | Débito, Crédito, Pix, Dinheiro, Boleto |
| Parcelas | 1 se à vista |
| Observações | Livre |

---

## 2. Automação com N8N (PDF para Excel)

### Requisitos
- N8N instalado (self-hosted ou cloud)
- Conta na Anthropic API (Claude)

### Configuração

1. Importe `n8n_workflow.json` no N8N: **Settings → Import Workflow**
2. Configure a credencial **Anthropic API** com sua chave
3. *(Opcional)* Configure **Google Sheets OAuth2** se quiser preencher o Google Sheets em vez do Excel local
4. Ative o workflow — o webhook ficará em: `https://SEU_N8N/webhook/fatura-pdf`

### Uso via curl

```bash
curl -X POST https://SEU_N8N/webhook/fatura-pdf \
     -F "arquivo=@fatura_cartao.pdf"
```

A resposta JSON conterá os lançamentos extraídos e o resumo por categoria.

### Importar a resposta para o Excel local

```bash
# Via arquivo
curl -s -X POST https://SEU_N8N/webhook/fatura-pdf \
     -F "arquivo=@fatura.pdf" > resposta.json
python importar_para_excel.py resposta.json

# Via pipe
curl -s -X POST https://SEU_N8N/webhook/fatura-pdf \
     -F "arquivo=@fatura.pdf" | python importar_para_excel.py
```

---

## 3. Sem N8N – Processar PDF Direto

```bash
# Com variável de ambiente
export ANTHROPIC_API_KEY=sk-ant-...
python importar_para_excel.py --pdf fatura.pdf

# Com chave inline
python importar_para_excel.py --pdf fatura.pdf --api-key sk-ant-...

# Forçar mês específico
python importar_para_excel.py --pdf fatura.pdf --mes Marco
```

---

## Fluxo N8N

```
PDF recebido (webhook POST)
     |
     v
Verificar se é PDF
     |
     v
Extrair texto (ExtractFromFile)
     |
     v
Claude AI – Analisar e categorizar lançamentos
     |
     v
Parsear JSON de resposta
     |
     |-- erro --> resposta 422
     |
     v
Dividir em lançamentos individuais
     |
     v
Google Sheets – Inserir linha por linha
     |
     v
Agregar resultados
     |
     v
Resposta JSON com resumo
```

---

## Categorias Suportadas

Alimentação, Moradia, Transporte, Saúde, Educação, Lazer, Vestuário, Assinaturas, Investimentos, Outros
