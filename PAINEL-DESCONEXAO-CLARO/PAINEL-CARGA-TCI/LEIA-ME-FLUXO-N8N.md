# PAINEL CARGA TCI — Fluxo n8n (WhatsApp/WAHA → Google Sheets)

Arquivo do fluxo: **`n8n-painel-carga-tci.json`** (importar no n8n).

Este fluxo lê mensagens de grupo do WhatsApp (via WAHA), interpreta planilhas
(SALDO GESTECH / GESTECH serializados) e textos de BASE TCI, e grava na planilha
do Google Sheets. Inclui **resgate offline**: ao religar o PC, recupera o que
chegou enquanto estava desligado.

---

## Como importar
1. n8n → **Workflows** → **Import from File** → selecione `n8n-painel-carga-tci.json`.
2. Em **cada** nó do Google Sheets (`Append BASE TCI`, `Limpar aba MATERIAIS`,
   `Gravar valores MATERIAIS`), vincule a credencial **Conta do Google Sheets 2**
   (ela vem marcada como `VINCULAR_NO_N8N`).
3. **Ative** o workflow (botão *Active*). A URL de produção só funciona ativo:
   `http://n8n:5678/webhook/painel-carga-waha`

## Configuração do WAHA (sessão `default`)
O webhook precisa entregar SUAS mensagens (`message.any`) e o status de reconexão
(`session.status`). PowerShell:

```powershell
$body = @{
  config = @{
    webhooks = @( @{
      url    = "http://n8n:5678/webhook/painel-carga-waha"
      events = @("message.any", "session.status")
    } )
    metadata = @{ numero = "551985670380"; projeto = "painel-carga-tci" }
  }
} | ConvertTo-Json -Depth 6

Invoke-RestMethod -Uri "http://localhost:3000/api/sessions/default" -Method Put `
  -Headers @{ "X-Api-Key" = "minha_chave_123" } -ContentType "application/json" -Body $body

Invoke-RestMethod -Uri "http://localhost:3000/api/sessions/default/restart" -Method Post `
  -Headers @{ "X-Api-Key" = "minha_chave_123" }
```

---

## O que mudou nesta versão (v8 — resgate offline)

1. **Append BASE TCI corrigido**
   Era um *Code node* usando `helpers.httpRequestWithAuthentication`, que **não é
   suportado** em Code node (causava o erro na linha 22). Agora é um **HTTP Request**
   com `predefinedCredentialType = googleSheetsOAuth2Api` (mesmo padrão do nó
   "Gravar valores MATERIAIS").

2. **Suas mensagens passam a contar (`fromMe`)**
   Removidos os dois blocos que descartavam mensagens enviadas pelo número
   conectado. Agora suas planilhas e textos no grupo também são processados.
   (Requer o evento `message.any` no WAHA — veja acima.)

3. **Resgate offline (ao religar)**
   - O nó **"Filtrar Grupo WAHA"** aprende e guarda o ID de cada grupo (`gruposVistos`)
     toda vez que chega uma mensagem ao vivo — **não precisa configurar ID na mão**.
   - O nó **"Resgate offline (WAHA history)"** roda em 2 gatilhos:
     - **"Resgate a cada 10 min"** (Schedule) — assim que o PC liga e o n8n sobe,
       ele recupera o backlog no próximo ciclo;
     - **"Reconectou?"** — quando o WAHA volta a `WORKING` (evento `session.status`),
       dispara o resgate na hora.
   - Ele lê o histórico recente de cada grupo (`/api/default/chats/{grupo}/messages`)
     e **re-injeta** cada mensagem no próprio webhook. O **dedup** (`seenMsg`)
     garante que nada seja gravado duas vezes.

### Limite honesto do resgate
Só recupera o que o **WhatsApp ainda tiver no histórico** quando o WAHA reconecta.
Cobre bem quedas curtas / uma noite (engine NOWEB recupera melhor). Para offline
muito longo, parte pode se perder — a solução 100% é hospedar WAHA + n8n num
servidor 24/7.

---

## Ajustes rápidos (dentro dos Code nodes)
- Intervalo do resgate: nó **"Resgate a cada 10 min"** → campo *Minutes*.
- Limitar a grupos específicos: no nó **"Resgate offline"**, preencha
  `GRUPOS_FIXOS = ['120363XXXX@g.us']`.
- Chave/URL do WAHA e URL do webhook do n8n: constantes no topo dos Code nodes
  `Filtrar Grupo WAHA` e `Resgate offline`.
