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

## Ambiente

Este é o repositório de **PRODUÇÃO**. O desenvolvimento acontece no DEV e é promovido para cá.

| Ambiente | Repositório GitHub | Pasta local |
|----------|-------------------|-------------|
| **DEV** | `NeoAppSolicitacaoConserto-DEV` | `App_Solicitacao_Conserto-DEV` |
| **PROD** | `NeoAppSolicitacaoConserto-PROD` | `App_Solicitacao_Conserto-PROD` |

---

## Links fixos do projeto — PROD

| Recurso | URL |
| ------- | --- |
| **Web App PROD (FIXA)** | `https://script.google.com/macros/s/AKfycbwPnDnjPdcMSALvbeQHAMfHWraUO03xjBHQq-cGGfcRH5DT0_Qw3OMmCAXdcwcUirWKsw/exec` |
| Planilha PROD | `https://docs.google.com/spreadsheets/d/1h4bCllbefqsmsjXpMSSRXVR6avdeXgDS3uGA3NercH8/edit` |
| Projeto GAS PROD | `https://script.google.com/u/0/home/projects/1z9htiRIBjfPCDVbxmghhOafZHI5gNddIeFdPrbp8Wb-1RPiasZed3FFD/edit` |

## Links fixos do projeto — DEV (referência)

| Recurso | URL |
| ------- | --- |
| **Web App DEV (FIXA)** | `https://script.google.com/macros/s/AKfycbwOM3Dl61OZT0dfSmZjv41-2b12t-lBet3MNYrpXGTL8CmsFNvAMTXyKIVI0Y2mzxynVg/exec` |
| Planilha DEV | `https://docs.google.com/spreadsheets/d/1YALSvcJ8ETZSHm6wFvhS7uNnbvBUfT_9fnoC6O5-WBc/edit` |

---

## Fluxo de promoção DEV → PROD — TOTALMENTE AUTOMATIZADO PELO CLAUDE

> Quando o usuário pedir "subir para PROD" ou "promover para PROD", o Claude executa a partir da pasta PROD:

```bash
# Pasta: App_Solicitacao_Conserto-PROD

# 1. Copiar arquivos de código do DEV (NÃO copiar .clasp.json nem appsscript.json)
cp ../App_Solicitacao_Conserto-DEV/api.gs .
cp ../App_Solicitacao_Conserto-DEV/audit.gs .
cp ../App_Solicitacao_Conserto-DEV/db.gs .
cp ../App_Solicitacao_Conserto-DEV/logic.gs .
cp ../App_Solicitacao_Conserto-DEV/ui.html .

# Para Code.gs: copiar e substituir SPREADSHEET_ID
# DEV usa: 1YALSvcJ8ETZSHm6wFvhS7uNnbvBUfT_9fnoC6O5-WBc
# PROD usa: 1h4bCllbefqsmsjXpMSSRXVR6avdeXgDS3uGA3NercH8

# 2. Commit e push no GitHub PROD
git add .
git commit -m "feat(vXXX-PROD): Promoção DEV vXXX → PROD"
git push

# 3. Envia código ao GAS PROD
clasp push --force

# 4. Cria nova versão no GAS PROD
clasp version "vXXX-PROD: descrição"

# 5. Atualiza o deployment ativo PROD (URL não muda)
clasp deploy \
  --deploymentId "AKfycbwPnDnjPdcMSALvbeQHAMfHWraUO03xjBHQq-cGGfcRH5DT0_Qw3OMmCAXdcwcUirWKsw" \
  --versionNumber XXX \
  --description "vXXX-PROD: descrição"

# 6. Arquivar deployments antigos (manter apenas @HEAD e o ativo atual)
# Listar: clasp deployments
# Arquivar cada ID antigo: clasp undeploy "<ID_ANTIGO>"
```

### IDs fixos PROD para os comandos clasp

- **Deployment ID PROD (URL fixa):** `AKfycbwPnDnjPdcMSALvbeQHAMfHWraUO03xjBHQq-cGGfcRH5DT0_Qw3OMmCAXdcwcUirWKsw`
- **Script ID PROD:** `1z9htiRIBjfPCDVbxmghhOafZHI5gNddIeFdPrbp8Wb-1RPiasZed3FFD`
- **Spreadsheet ID PROD:** `1h4bCllbefqsmsjXpMSSRXVR6avdeXgDS3uGA3NercH8`
- O `@HEAD` é automático do GAS — **nunca arquivar**
- O número da versão é sequencial — verificar com `clasp deployments` e incrementar

> ⚠️ **NUNCA usar** `clasp deploy` sem `--deploymentId` — isso cria um novo deployment com URL diferente.
> ⚠️ **NUNCA sobrescrever** `.clasp.json` com o do DEV — têm Script IDs diferentes.
> ⚠️ **NUNCA sobrescrever** o `SPREADSHEET_ID` no Code.gs com o ID do DEV.

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
