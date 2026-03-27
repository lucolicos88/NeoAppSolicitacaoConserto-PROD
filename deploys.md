# Histórico de Deploys — App_Solicitacao_Conserto

## Links fixos

| Recurso | URL |
| ------- | --- |
| **Web App (FIXA)** | `https://script.google.com/macros/s/AKfycbwOM3Dl61OZT0dfSmZjv41-2b12t-lBet3MNYrpXGTL8CmsFNvAMTXyKIVI0Y2mzxynVg/exec` |
| Planilha de dados | `https://docs.google.com/spreadsheets/d/1YALSvcJ8ETZSHm6wFvhS7uNnbvBUfT_9fnoC6O5-WBc/edit` |
| Projeto GAS | `https://script.google.com/u/0/home/projects/1JlKPUx3SFy09mmcFlNBFgcmjgoPqXJkADQdPqAa8Vkcefdgn0cG9J57F/edit` |

---

## Como fazer deploy (manter URL fixa)

> ⚠️ **NUNCA usar "Nova implantação"** — isso gera uma URL diferente.
> Sempre editar (✏️) a implantação existente para manter a URL fixa.

### Passo a passo

1. Editar arquivos localmente
2. Commit + push no **GitHub**:

   ```bash
   git add .
   git commit -m "feat(vXXX): descrição"
   git push
   ```

3. `clasp push` (dentro da pasta do projeto)
4. No editor GAS: **Implantar → Gerenciar implantações**
5. Clicar no ✏️ **(lápis)** da implantação ativa
6. Em **Versão**: selecionar **"Nova versão"** com descrição da mudança
7. Clicar **"Implantar"** — a URL **não muda**
8. Registrar abaixo no histórico

---

## Registro de versões

### v213 — 2026-03-26 — feat

**Campo: Necessário confirmação Médica?**

- Novo campo `confirmacao_medica` (SIM/NAO) adicionado à planilha `Erros` (índice [7])
- Formulário de solicitação com novo campo obrigatório
- Tabela de solicitações com nova coluna "Conf. Médica?"
- Linhas com confirmação médica = SIM exibem fundo lilás (`#f3e5f5`) em vez do verde padrão
- Badge roxo (`#7b1fa2`) para identificar visualmente as solicitações que precisam de confirmação médica

---

### v212 — fix

- Aumenta diâmetro do círculo para acomodar logo quadrada no loading

### v211 — fix

- Remove border-radius da logo para não cortar no loading

### v210 — feat

- Ajusta logo, arquiva versões antigas e atualiza docs

### v209 — feat

- Melhora qualidade visual da logo no loading

### v208 — feat

- Usa imagem real `logo gotinha.jpg` em base64 no loading

### v207 — feat

- Implementa loading NeoExpedicao e redesenha campo Setores

### v2.8.31 — debug

- Retorna info de debug para o frontend

### v2.8.30 — fix

- Melhoria na verificação de permissões de usuário

### v2.8.29 — feat

- Corrige ui.alert e adiciona formatação de planilhas

### v2.8.28 — fix

- Corrige erro ao limpar dados de teste

### v2.8.27 — fix

- Gerador de dados configura usuário ADMIN automaticamente

### v2.8.26 — feat

- Menu App Solicitações e gerador de dados de teste

### v2.8.25 — feat

- Novas abas Ranking e Erros no Dashboard

### v2.8.23 — feat

- Espaço maior labels gráfico e ano input texto

### v2.8.22 — feat

- Tabela Resposta com colunas diferenciadas e valor R$

### v2.8.21 — feat

- Card variação mostra quantidade e valor total

### v2.8.20 — fix

- Correção contagem variação valor e exibição na tabela

### v2.8.19 — feat

- Botão atualizar dashboard e refresh automático

### v2.8.18 — fix

- Correção habilitação campo diferença de valor

### v2.8.17 — feat

- Fechar modal ao clicar fora e campo diferença condicional

### v2.8.16 — fix

- Fontes ainda maiores para impressão PDF

### v2.8.15 — fix

- Melhoria resolução gráficos PDF para impressão

### v2.8.14 — fix

- Cor verde no status Corrigida e gráficos PDF maiores

### v2.8.13 — fix

- Correção salvar resposta (coluna id_solicitacao) e PDF compacto

### v2.8.12 — fix

- Correção salvar resposta e PDF páginas

### v2.8.11 — fix

- Correção erro console e layout PDF 2 páginas

### v2.8.10 — fix

- Redução das margens do relatório PDF

### v2.8.9 — feat

- Gráficos profissionais e informações de suporte

### v2.8.8 — feat

- Redesign cards com indicadores de tendência

### v2.8.7 — feat

- PDF com logo Neoformula e gráficos maiores

### v2.8.6 — feat

- Relatório PDF profissional completo

### v2.8.5 — fix

- PDF corrigido: filtros, tabela, rótulos e resolução

### v2.8.4 — feat

- PDF melhorado com filtros, títulos e tabela detalhada

### v2.8.3 — feat

- Rótulos internos nos gráficos, PDF melhorado e correções

### v2.8.2 — fix

- Cor verde nas sub-abas e correção do gráfico SLA

### v2.8.1 — feat

- Sub-abas na Auditoria com botões de limpar

### v2.8.0 — feat

- Consulta por requisição na aba Auditoria

### v2.7.4 — fix

- Correção robusta para resposta null na aba Auditoria

### v2.7.3 — feat

- Loading em todos os botões, campo diferença valor numérico e correções

### v2.7.1 — fix

- Correções de bugs e ajustes de permissões

### v2.7.0 — feat

- Modal de edição para SOLICITANTE/ADMIN e melhorias na auditoria
- Otimização massiva de performance

### v2.6.3 — fix

- Corrige backup de respostas e adiciona escopo do Drive

### v2.6.1 — fix

- Debug para testar pasta de backup

### v2.6.0 — feat

- Sistema de arquivamento mensal para Looker Studio

### v2.5.1 — feat

- Melhorias no backup e organização da UI

### v2.5.0 — feat

- Perfil Conferente, cores de classificação, backup e limpeza

### v2.4.2 — fix

- Corrige carregamento de solicitações e dashboard inicial

### v2.4.1 — refactor

- Consolida estilos inline em classes CSS reutilizáveis
