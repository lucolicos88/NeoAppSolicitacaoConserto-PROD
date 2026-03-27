# CLAUDE.md — Instruções para o App_Solicitacao_Conserto

## Sobre o projeto
Sistema de Solicitação de Consertos da **NeoFormula**. Registra e acompanha erros em requisições farmacêuticas, com fluxo: Solicitação → Erro → Resposta/Correção.

- **Backend:** Google Apps Script (V8) — arquivos `.gs`
- **Frontend:** `ui.html` — SPA vanilla (sem frameworks, sem npm)
- **Banco de dados:** Google Sheets (abas = tabelas, colunas lidas por índice numérico)
- **Deploy:** Google Web App via CLASP (`clasp push` + publicar nova versão no GAS)

---

## Regras críticas de desenvolvimento

### Backend (.gs)
- **Nunca use `console.log`** — use `safeLogDebug_()` ou `logDebug_()`
- **Funções internas** terminam com `_` (ex: `getConfigData_()`)
- **Funções de API pública** não têm sufixo (ex: `submitRequest()`)
- **Todo retorno de API** segue: `{ ok: boolean, data: any, errors: string[], debugId: string }`
- **Colunas da planilha são lidas por índice numérico** — ex: `erros[i][5]` — nunca por nome
- **Novos campos** na planilha `Erros` devem ser adicionados **ao final** (após `criado_em` no índice [6]) para não quebrar dados históricos. O campo `confirmacao_medica` está no índice [7].
- **Cache:** após alterar configurações, chamar `invalidateConfigCache_()`
- **Validações** ficam em `logic.gs` (`validateRequestPayload_`, `validateResponsePayload_`)
- **Operações de banco** ficam em `db.gs`

### Frontend (ui.html)
- Arquivo único com ~13.000 linhas — CSS, HTML e JS todos inline
- **Não usar frameworks** — só vanilla JS
- Warnings de "CSS inline styles" são esperados e pré-existentes — não corrija
- **Payload do formulário** usa camelCase: `confirmacaoMedica`, `diferencaNoValor`
- Ao adicionar campo no formulário: adicionar também em `clearSolicitacaoForm()`
- Ao adicionar coluna na tabela `#openErrorsTable`: atualizar CSS `nth-child`, header `<th>`, renderização da linha e `colspan` da mensagem vazia

---

## Schema atual da planilha Erros (índices de coluna)
```
[0] id_solicitacao
[1] sequencia_erro
[2] erro
[3] detalhamento
[4] setor_local
[5] diferenca_valor
[6] criado_em
[7] confirmacao_medica   ← adicionado v213
```

## Schema atual da planilha Solicitacoes
```
[0] id_solicitacao
[1] requisicao
[2] solicitante
[3] data_hora_pedido
[4] status (ABERTO / EM_CORRECAO / CORRIGIDO)
[5] criado_por_email
[6] criado_em
```

## Schema atual da planilha Respostas
```
[0] id_resposta
[1] id_solicitacao
[2] sequencia_erro
[3] nome_responsavel
[4] email_responsavel
[5] erro_corrigido
[6] houve_diferenca_valor
[7] diferenca_valor_resposta
[8] observacoes
[9] data_hora_correcao
[10] criado_em
```

---

## Perfis de usuário
| Perfil | Permissões |
|--------|-----------|
| ADMIN | Acesso total |
| CONFERENTE | Edita solicitações e respostas |
| RESPOSTA | Responde solicitações do seu setor |
| ESPECTADOR | Apenas visualização |

---

## Links fixos do projeto

| Recurso | URL |
| ------- | --- |
| **Web App (FIXA)** | `https://script.google.com/macros/s/AKfycbwOM3Dl61OZT0dfSmZjv41-2b12t-lBet3MNYrpXGTL8CmsFNvAMTXyKIVI0Y2mzxynVg/exec` |
| Planilha de dados | `https://docs.google.com/spreadsheets/d/1YALSvcJ8ETZSHm6wFvhS7uNnbvBUfT_9fnoC6O5-WBc/edit` |
| Projeto GAS | `https://script.google.com/u/0/home/projects/1JlKPUx3SFy09mmcFlNBFgcmjgoPqXJkADQdPqAa8Vkcefdgn0cG9J57F/edit` |

---

## Fluxo de deploy — TOTALMENTE AUTOMATIZADO PELO CLAUDE

> O Claude executa todos os passos abaixo. O usuário não precisa fazer nada manualmente.

### O Claude executa na ordem

```bash
# 1. Commit e push no GitHub
git add .
git commit -m "feat(vXXX): descrição"
git push

# 2. Envia código ao GAS
clasp push --force

# 3. Cria nova versão no GAS
clasp version "vXXX: descrição"

# 4. Atualiza o deployment ativo (URL não muda)
clasp deploy \
  --deploymentId "AKfycbwOM3Dl61OZT0dfSmZjv41-2b12t-lBet3MNYrpXGTL8CmsFNvAMTXyKIVI0Y2mzxynVg" \
  --versionNumber XXX \
  --description "vXXX: descrição"

# 5. Arquivar deployments antigos (manter apenas @HEAD e o ativo atual)
# Listar: clasp deployments
# Arquivar cada ID antigo: clasp undeploy "<ID_ANTIGO>"
```

### IDs fixos para os comandos clasp

- **Deployment ID (URL fixa):** `AKfycbwOM3Dl61OZT0dfSmZjv41-2b12t-lBet3MNYrpXGTL8CmsFNvAMTXyKIVI0Y2mzxynVg`
- O `@HEAD` é automático do GAS — **nunca arquivar**
- O número da versão é sequencial — verificar com `clasp deployments` e incrementar

> ⚠️ **NUNCA usar** `clasp deploy` sem `--deploymentId` — isso cria um novo deployment com URL diferente.

---

## SLA padrão
- Alerta: 15 minutos
- Crítico: 30 minutos
- Fuso: America/Sao_Paulo
- Configurável via aba `Limiares_SLA`

---

## Convenções de commit
```
feat(vXXX): descrição da nova funcionalidade
fix(vXXX): descrição da correção
refactor(vXXX): descrição da refatoração
perf(vXXX): melhoria de performance
debug(vXXX): ajustes de debug
```
