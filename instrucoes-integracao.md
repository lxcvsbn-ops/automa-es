# Instruções de Integração — Follow-up por Status

## 1. Colunas necessárias na Google Sheet

Sua planilha precisa ter as seguintes colunas (o nome deve ser exatamente igual):

| Coluna | Descrição |
|---|---|
| `id_chamado` | ID único do chamado (ex: CHM-001) |
| `titulo` | Título/assunto do chamado |
| `nome_solicitante` | Nome de quem abriu o chamado |
| `email` | E-mail do solicitante |
| `status` | Status atual: `aberto`, `Em análise`, `Resolvido`, `Aguardando retorno` |
| `ultimo_status_enviado` | Último status para o qual o e-mail foi enviado (começa vazio) |
| `gmail_message_id` | ID da mensagem do Gmail do primeiro e-mail enviado |

---

## 2. O que adicionar no workflow de criação do chamado

Após o nó que envia o primeiro e-mail (Gmail), adicione um nó **Google Sheets (Update)** para salvar o `gmail_message_id` na planilha:

```
Coluna a atualizar:  gmail_message_id
Valor:               ={{ $json.id }}
Coluna de busca:     id_chamado
```

O campo `$json.id` é o ID da mensagem retornado pelo nó Gmail após o envio.

Também defina `ultimo_status_enviado` = `aberto` na criação do chamado.

---

## 3. Como importar o workflow no N8N

1. Acesse seu N8N
2. Clique em **+ New Workflow** → **Import from file**
3. Selecione o arquivo `workflow-followup-status.json`
4. Após importar, configure as credenciais:
   - Substitua `SEU_GOOGLE_SHEET_ID_AQUI` pelo ID da sua planilha
   - Selecione suas credenciais do Google Sheets e Gmail nos nós correspondentes
5. Ative o workflow

---

## 4. Como funciona o sistema

```
A cada 1 minuto
   └─ Busca todas as linhas da planilha
         └─ Filtra linhas onde:
               • status ≠ ultimo_status_enviado
               • status ≠ "aberto"
               • email preenchido
               • gmail_message_id preenchido
                  └─ Switch por status
                        ├─ Em análise       → Template 1 → Responder na thread → Atualizar planilha
                        ├─ Resolvido        → Template 2 → Responder na thread → Atualizar planilha
                        └─ Aguardando ret.  → Template 3 → Responder na thread → Atualizar planilha
```

O campo `ultimo_status_enviado` é atualizado após cada envio, garantindo que o mesmo e-mail **não seja enviado duas vezes**.
