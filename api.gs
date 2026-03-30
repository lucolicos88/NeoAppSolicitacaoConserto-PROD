/**
 * ╔══════════════════════════════════════════════════════════════════════════════╗
 * ║                    SISTEMA DE SOLICITAÇÃO DE CONSERTOS                       ║
 * ║                              API de Serviços                                 ║
 * ╠══════════════════════════════════════════════════════════════════════════════╣
 * ║  Arquivo: api.gs                                                             ║
 * ║  Versão: 2.8.31                                                              ║
 * ║  Última atualização: Janeiro 2026                                            ║
 * ╠══════════════════════════════════════════════════════════════════════════════╣
 * ║                                                                              ║
 * ║  DESCRIÇÃO:                                                                  ║
 * ║  Este arquivo contém todas as funções de API do sistema, responsáveis        ║
 * ║  pela comunicação entre o frontend (ui.html) e o backend (Google Sheets).    ║
 * ║  Inclui operações CRUD, validações, auditoria e relatórios.                  ║
 * ║                                                                              ║
 * ║  ESTRUTURA DO ARQUIVO:                                                       ║
 * ║  ├── Funções de Diagnóstico (testConnection, debugSnapshot)                  ║
 * ║  ├── Funções de Bootstrap (getBootstrapData)                                 ║
 * ║  ├── Funções de Leitura (getOpenErros, getDashboardRecords, getAuditLogs)   ║
 * ║  ├── Funções de Escrita (submitRequest, submitResponse)                      ║
 * ║  ├── Funções de Atualização (updateSolicitacao, updateUsuario)              ║
 * ║  ├── Funções de Exclusão (deleteSolicitacao, deleteUsuario)                 ║
 * ║  ├── Funções de Configuração (addConfigItem, deleteConfigItem)              ║
 * ║  ├── Funções de Gestão de Usuários                                           ║
 * ║  ├── Funções de Manutenção (executarLimpeza, executarArquivamentoMensal)    ║
 * ║  └── Funções Auxiliares Internas (getThresholds_, usuarioPodeVerSetor_)     ║
 * ║                                                                              ║
 * ║  PADRÃO DE RESPOSTA:                                                         ║
 * ║  Todas as funções retornam objetos no formato:                               ║
 * ║  {                                                                            ║
 * ║    ok: boolean,        // Sucesso da operação                                ║
 * ║    debugId: string,    // ID único para rastreamento                         ║
 * ║    data: any,          // Dados retornados (quando ok=true)                  ║
 * ║    errors: string[],   // Lista de erros (quando ok=false)                   ║
 * ║    _debug: object      // Info de debug (opcional)                           ║
 * ║  }                                                                            ║
 * ║                                                                              ║
 * ║  SISTEMA DE PERMISSÕES:                                                      ║
 * ║  - ADMIN: Acesso total a todas as funcionalidades                            ║
 * ║  - CONFERENTE: Pode editar solicitações e respostas                          ║
 * ║  - RESPOSTA: Pode responder solicitações do seu setor                        ║
 * ║  - ESPECTADOR: Apenas visualização                                           ║
 * ║                                                                              ║
 * ╚══════════════════════════════════════════════════════════════════════════════╝
 */

// ============================================================================
// FUNÇÕES DE DIAGNÓSTICO E TESTE
// ============================================================================

/**
 * Testa conectividade básica com o backend
 * Útil para verificar se o Web App está funcionando
 *
 * @returns {Object} Status da conexão com timestamp
 * @example
 * // Retorno esperado:
 * { ok: true, message: "Conexão funcionando!", timestamp: "2026-01-29T10:00:00.000Z" }
 */
function testConnection() {
  return { ok: true, message: "Conexão funcionando!", timestamp: new Date().toISOString() };
}

// ============================================================================
// DIAGNÓSTICO / HEALTH CHECK
// ============================================================================

/**
 * Executa diagnóstico completo do sistema: banco de dados, configurações e app.
 * @param {boolean} [autoFix=false] - Se true, tenta corrigir automaticamente os problemas encontrados.
 * @returns {Object} Resultado detalhado de cada verificação e lista de correções aplicadas
 */
function healthCheck(autoFix) {
  autoFix = autoFix === true;
  const debugId = Utilities.getUuid();
  const inicio = new Date();
  const checks = [];
  const fixes = [];

  function addCheck(categoria, nome, status, detalhe, valor) {
    checks.push({ categoria: categoria, nome: nome, status: status, detalhe: detalhe, valor: valor || null });
  }
  function addFix(descricao, sucesso) {
    fixes.push({ descricao: descricao, sucesso: sucesso });
  }

  // --- 1. Conexão com a Planilha ---
  let ss = null;
  try {
    ss = getSpreadsheet_();
    addCheck("Banco de Dados", "Conexão com a Planilha", "ok", "Planilha aberta com sucesso", ss.getId());
  } catch (e) {
    addCheck("Banco de Dados", "Conexão com a Planilha", "erro",
      "Não foi possível abrir a planilha: " + e + " — Verifique o SPREADSHEET_ID no código.", null);
    return { ok: false, debugId: debugId, checks: checks, fixes: fixes, duracaoMs: new Date() - inicio };
  }

  // --- 2. Verificação de cada aba (estrutura + cabeçalhos) ---
  // headers: lista completa de cabeçalhos esperados na linha 1 (em ordem)
  const abasEsperadas = [
    { chave: "SOLICITACOES",   nome: "Solicitacoes",   headers: ["id_solicitacao","requisicao","solicitante","data_hora_pedido","status","criado_por_email","criado_em"] },
    { chave: "ERROS",          nome: "Erros",          headers: ["id_solicitacao","sequencia_erro","erro","detalhamento","setor_local","diferenca_valor","criado_em","confirmacao_medica"] },
    { chave: "RESPOSTAS",      nome: "Respostas",      headers: ["id_resposta","id_solicitacao","sequencia_erro","nome_responsavel","email_responsavel","erro_corrigido","houve_diferenca_valor","diferenca_valor_resposta","observacoes","data_hora_correcao","criado_em"] },
    { chave: "USUARIOS",       nome: "Usuarios",       headers: ["email","nome","perfil","setores"] },
    { chave: "SOLICITANTES",   nome: "Solicitantes",   headers: ["solicitante"] },
    { chave: "SETORES",        nome: "Setores_Local",  headers: ["setor_local"] },
    { chave: "ERROS_CADASTRO", nome: "Erros_Cadastro", headers: ["erro"] },
    { chave: "LIMIARES",       nome: "Limiares_SLA",   headers: ["horas_alerta","horas_critico","fuso_padrao"] },
    { chave: "AUDITORIA",      nome: "Auditoria",      headers: ["id_auditoria","email_usuario","acao","tabela","chave_registro","campo","valor_anterior","valor_novo","ip","data_hora"] },
    { chave: "LOGS_DEBUG",     nome: "Logs_Debug",     headers: ["data_hora","funcao","email","mensagem","payload"] },
    { chave: "CONFIG_GERAL",   nome: "Config_Geral",   headers: ["chave","valor"] }
  ];

  var abaProblema = false;
  const contagensAbas = {};
  for (var ai = 0; ai < abasEsperadas.length; ai++) {
    var aba = abasEsperadas[ai];
    try {
      var sheet = ss.getSheetByName(CONFIG.SHEETS[aba.chave]);
      if (!sheet) {
        abaProblema = true;
        addCheck("Banco de Dados", "Aba: " + aba.nome, "erro", "Aba não encontrada na planilha", null);
        continue;
      }
      var lastRow = sheet.getLastRow();
      if (lastRow < 1) {
        abaProblema = true;
        addCheck("Banco de Dados", "Aba: " + aba.nome, "erro", "Aba existe mas está vazia (sem cabeçalho)", null);
        continue;
      }

      // Ler cabeçalhos reais da linha 1
      var numEsperados = aba.headers.length;
      var headersReais = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), numEsperados)).getValues()[0];
      var erradosCols = [];
      for (var hi = 0; hi < numEsperados; hi++) {
        var esperado = aba.headers[hi];
        var real = String(headersReais[hi] || "").trim();
        if (real !== esperado) {
          erradosCols.push({ col: hi + 1, esperado: esperado, real: real || "(vazio)" });
        }
      }

      var registros = lastRow - 1;
      contagensAbas[aba.chave] = registros;

      if (erradosCols.length === 0) {
        addCheck("Banco de Dados", "Aba: " + aba.nome, "ok", registros + " registro(s)", registros);
      } else {
        var descErros = erradosCols.map(function(e) {
          return "col " + e.col + ": '" + e.real + "' → esperado '" + e.esperado + "'";
        }).join("; ");

        if (autoFix) {
          try {
            for (var fi = 0; fi < erradosCols.length; fi++) {
              sheet.getRange(1, erradosCols[fi].col).setValue(erradosCols[fi].esperado);
            }
            addCheck("Banco de Dados", "Aba: " + aba.nome, "ok",
              registros + " registro(s) — " + erradosCols.length + " cabeçalho(s) corrigido(s)", registros);
            addFix("Aba " + aba.nome + ": cabeçalhos corrigidos — " + descErros, true);
          } catch (fixErr) {
            addCheck("Banco de Dados", "Aba: " + aba.nome, "aviso",
              erradosCols.length + " cabeçalho(s) incorreto(s) — falha ao corrigir: " + fixErr, registros);
            addFix("Falha ao corrigir cabeçalhos da aba " + aba.nome + " — " + fixErr, false);
          }
        } else {
          addCheck("Banco de Dados", "Aba: " + aba.nome, "aviso",
            erradosCols.length + " cabeçalho(s) incorreto(s): " + descErros + " — Use 'Diagnosticar e Corrigir'.", registros);
        }
      }
    } catch (e) {
      abaProblema = true;
      addCheck("Banco de Dados", "Aba: " + aba.nome, "erro", "Erro ao acessar: " + e, null);
    }
  }

  // Auto-corrigir abas completamente ausentes ou vazias
  if (abaProblema && autoFix) {
    try {
      ensureSheets_();
      addFix("Abas ausentes ou vazias recriadas com cabeçalhos corretos (ensureSheets_)", true);
    } catch (fixErr) {
      addFix("Falha ao recriar abas — " + fixErr, false);
    }
  }

  // --- 3. Configurações essenciais ---
  var configProblema = false;
  try {
    invalidateConfigCache_();
    var configData = getConfigData_();
    var numFarm = configData.pharmaceuticas ? configData.pharmaceuticas.length : 0;
    var numSetores = configData.setores ? configData.setores.length : 0;
    var numErros = configData.erros ? configData.erros.length : 0;

    if (numFarm === 0 || numSetores === 0 || numErros === 0) configProblema = true;

    addCheck("Configurações", "Solicitantes cadastrados", numFarm > 0 ? "ok" : "aviso",
      numFarm > 0 ? numFarm + " solicitante(s)" : "Nenhum solicitante — use Configurações > Solicitantes para adicionar", numFarm);
    addCheck("Configurações", "Setores cadastrados", numSetores > 0 ? "ok" : "aviso",
      numSetores > 0 ? numSetores + " setor(es)" : "Nenhum setor — use Configurações > Setores para adicionar", numSetores);
    addCheck("Configurações", "Tipos de erro cadastrados", numErros > 0 ? "ok" : "aviso",
      numErros > 0 ? numErros + " tipo(s)" : "Nenhum tipo de erro — use Configurações > Tipos de Erro para adicionar", numErros);
  } catch (e) {
    configProblema = true;
    addCheck("Configurações", "Leitura de configurações", "erro", "Falha ao carregar config: " + e, null);
  }

  if (configProblema && autoFix) {
    try {
      ensureConfigDefaults_();
      invalidateConfigCache_();
      addFix("Configurações padrão restauradas (solicitantes, setores e tipos de erro de exemplo inseridos)", true);
    } catch (fixErr) {
      addFix("Falha ao restaurar configurações padrão — " + fixErr, false);
    }
  }

  // --- 4. SLA ---
  try {
    var thresholds = getThresholds_();
    var warnOk = thresholds.warnMinutes > 0;
    var critOk = thresholds.criticalMinutes > thresholds.warnMinutes;
    if (warnOk && critOk) {
      addCheck("Configurações", "Limiares de SLA", "ok",
        "Alerta: " + thresholds.warnMinutes + "min | Crítico: " + thresholds.criticalMinutes + "min",
        thresholds.warnMinutes + "/" + thresholds.criticalMinutes);
    } else {
      if (autoFix) {
        try {
          ensureThresholdDefault_();
          addCheck("Configurações", "Limiares de SLA", "ok", "Limiares padrão restaurados (15min alerta / 30min crítico)", "15/30");
          addFix("Limiares de SLA inválidos restaurados para padrão (15min alerta, 30min crítico)", true);
        } catch (fixErr) {
          addCheck("Configurações", "Limiares de SLA", "aviso", "Valores inválidos — falha ao corrigir: " + fixErr, null);
          addFix("Falha ao restaurar limiares de SLA — " + fixErr, false);
        }
      } else {
        addCheck("Configurações", "Limiares de SLA", "aviso",
          "Valores inválidos — Alerta: " + thresholds.warnMinutes + " | Crítico: " + thresholds.criticalMinutes +
          " — Use 'Diagnosticar e Corrigir' para restaurar o padrão.", null);
      }
    }
  } catch (e) {
    addCheck("Configurações", "Limiares de SLA", "erro", "Falha ao ler limiares: " + e, null);
  }

  // --- 5. Usuários ---
  try {
    var usuarios = getUsuarios_();
    var numUsuarios = usuarios ? usuarios.length : 0;
    var admins = usuarios ? usuarios.filter(function(u) { return u.perfil === "ADMIN"; }).length : 0;
    if (numUsuarios === 0) {
      addCheck("Configurações", "Usuários cadastrados", "erro",
        "Nenhum usuário cadastrado — Acesse Configurações > Usuários e cadastre um ADMIN", 0);
    } else if (admins === 0) {
      addCheck("Configurações", "Usuários cadastrados", "aviso",
        numUsuarios + " usuário(s) mas nenhum ADMIN — Acesse Configurações > Usuários e defina um perfil ADMIN", numUsuarios);
    } else {
      addCheck("Configurações", "Usuários cadastrados", "ok",
        numUsuarios + " usuário(s) | " + admins + " ADMIN(s)", numUsuarios);
    }
  } catch (e) {
    addCheck("Configurações", "Usuários cadastrados", "erro", "Falha ao ler usuários: " + e, null);
  }

  // --- 6. Usuário atual ---
  try {
    var emailAtual = Session.getActiveUser().getEmail() || "";
    var usuarioAtual = getUsuarioContexto_(emailAtual);
    if (!emailAtual) {
      addCheck("Aplicativo", "Usuário atual", "aviso", "Não foi possível identificar o e-mail do usuário", null);
    } else {
      addCheck("Aplicativo", "Usuário atual", "ok",
        emailAtual + " | Perfil: " + (usuarioAtual.perfil || "?"), emailAtual);
    }
  } catch (e) {
    addCheck("Aplicativo", "Usuário atual", "erro", "Falha ao identificar usuário: " + e, null);
  }

  // --- 7. Cache ---
  try {
    var cache = CacheService.getScriptCache();
    var testKey = "healthcheck_" + debugId;
    cache.put(testKey, "ok", 10);
    var testVal = cache.get(testKey);
    cache.remove(testKey);
    if (testVal === "ok") {
      addCheck("Aplicativo", "Cache do sistema", "ok", "Leitura e escrita funcionando", null);
    } else {
      if (autoFix) {
        try {
          invalidateConfigCache_();
          addCheck("Aplicativo", "Cache do sistema", "ok", "Cache limpo e reiniciado com sucesso", null);
          addFix("Cache com falha: limpo e reiniciado", true);
        } catch (fixErr) {
          addCheck("Aplicativo", "Cache do sistema", "aviso", "Cache com falha — não foi possível corrigir: " + fixErr, null);
          addFix("Falha ao limpar cache — " + fixErr, false);
        }
      } else {
        addCheck("Aplicativo", "Cache do sistema", "aviso",
          "Cache não retornou o valor esperado — use 'Diagnosticar e Corrigir' ou o botão 'Limpar Cache'.", null);
      }
    }
  } catch (e) {
    addCheck("Aplicativo", "Cache do sistema", "aviso", "Cache indisponível: " + e, null);
  }

  // --- 8. Resumo de solicitações ---
  try {
    var solSheet = ss.getSheetByName(CONFIG.SHEETS.SOLICITACOES);
    if (solSheet && solSheet.getLastRow() > 1) {
      var solData = solSheet.getDataRange().getValues();
      var abertas = 0, emCorrecao = 0, corrigidas = 0;
      for (var si = 1; si < solData.length; si++) {
        var st = solData[si][4];
        if (st === "ABERTO") abertas++;
        else if (st === "EM_CORRECAO") emCorrecao++;
        else if (st === "CORRIGIDO") corrigidas++;
      }
      var total = solData.length - 1;
      addCheck("Aplicativo", "Resumo de solicitações",
        abertas + emCorrecao > 0 ? "aviso" : "ok",
        "Total: " + total + " | Abertas: " + abertas + " | Em correção: " + emCorrecao + " | Corrigidas: " + corrigidas,
        total);
    } else {
      addCheck("Aplicativo", "Resumo de solicitações", "ok", "Nenhuma solicitação registrada ainda", 0);
    }
  } catch (e) {
    addCheck("Aplicativo", "Resumo de solicitações", "erro", "Falha ao ler solicitações: " + e, null);
  }

  // --- Resultado final ---
  var totalErros = checks.filter(function(c) { return c.status === "erro"; }).length;
  var totalAvisos = checks.filter(function(c) { return c.status === "aviso"; }).length;
  var statusGeral = totalErros > 0 ? "erro" : (totalAvisos > 0 ? "aviso" : "ok");

  safeLogDebug_("healthCheck", "complete", {
    debugId: debugId, autoFix: autoFix,
    statusGeral: statusGeral, totalErros: totalErros, totalAvisos: totalAvisos,
    totalFixes: fixes.length
  });

  return {
    ok: true,
    debugId: debugId,
    autoFix: autoFix,
    statusGeral: statusGeral,
    totalErros: totalErros,
    totalAvisos: totalAvisos,
    totalOk: checks.filter(function(c) { return c.status === "ok"; }).length,
    checks: checks,
    fixes: fixes,
    duracaoMs: new Date() - inicio,
    timestamp: inicio.toISOString()
  };
}

// ============================================================================
// HELPERS DE SEGURANÇA
// ============================================================================

/**
 * Retorna o email do usuário autenticado. Lança erro se sessão inválida.
 * @returns {string} Email do usuário
 */
function getEmailValidado_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error("Sessão inválida. Faça login no Google e tente novamente.");
  return email;
}

/**
 * Rate limiting simples: máximo 30 requisições de escrita por usuário por minuto.
 * Usa CacheService.getUserCache() (por usuário, não compartilhado).
 * @param {string} email - Email do usuário
 */
function checkRateLimit_(email) {
  try {
    const cache = CacheService.getUserCache();
    const key = "rl_w_" + email.replace(/[^a-zA-Z0-9]/g, "_");
    const count = parseInt(cache.get(key) || "0", 10);
    if (count >= 15) {
      throw new Error("Muitas requisições. Aguarde 1 minuto e tente novamente.");
    }
    cache.put(key, String(count + 1), 60);
  } catch (e) {
    if (e.message && e.message.indexOf("Muitas requisições") !== -1) throw e;
    // CacheService indisponível: registra aviso mas não bloqueia a operação
    Logger.log("AVISO: checkRateLimit_ — CacheService falhou: " + e);
  }
}

// ============================================================================
// FUNÇÕES DE BOOTSTRAP E INICIALIZAÇÃO
// ============================================================================

/**
 * Carrega todos os dados iniciais necessários para a aplicação
 * Chamada uma única vez quando o usuário abre a aplicação
 *
 * @returns {Object} Objeto contendo:
 *   - ok: boolean - Sucesso da operação
 *   - debugId: string - ID para rastreamento
 *   - userEmail: string - Email do usuário logado
 *   - usuario: Object - Dados e permissões do usuário
 *   - config: Object - Configurações (tipos de erro, setores, solicitantes)
 *   - listas: Object - Listas de colaboradores e setores
 *   - responsaveis: string[] - Lista de responsáveis disponíveis
 *   - usuarios: Object[] - Lista de usuários (apenas para ADMIN)
 *
 * @description
 * Esta função é otimizada para carregar apenas dados essenciais.
 * Dados pesados (solicitações, erros abertos) são carregados
 * separadamente via getDashboardRecords() e getOpenErros().
 */
function getBootstrapData() {
  const debugId = Utilities.getUuid();
  try {
    safeLogDebug_("getBootstrapData", "start", { debugId: debugId });
    const result = {
      ok: true,
      debugId: debugId,
      errors: [],
      userEmail: "",
      usuario: null,
      config: null,
      listas: { colaboradores: [], setores: [] },
      configGeral: { pastaBackupId: "" },
      responsaveis: [],
      usuarios: [],
      openErros: [],
      solicitacoes: [],
      dashboard: { totalSolicitacoes: 0, abertas: 0, emCorrecao: 0, corrigidas: 0, alerta: 0, critico: 0 }
    };

    try {
      ensureSheets_();
    } catch (error) {
      result.errors.push("ensureSheets_: " + error);
    }

    try {
      result.config = getConfigData_();
    } catch (error) {
      result.errors.push("getConfigData_: " + error);
    }

    try {
      result.userEmail = Session.getActiveUser().getEmail() || "";
      result.usuario = getUsuarioContexto_(result.userEmail);
    } catch (error) {
      result.errors.push("getUsuarioContexto_: " + error);
    }

    try {
      result.listas = getListasData_();
    } catch (error) {
      result.errors.push("getListasData_: " + error);
    }

    try {
      result.configGeral = { pastaBackupId: getConfigGeral_("pasta_backup_id") };
    } catch (error) {
      result.errors.push("getConfigGeral_: " + error);
    }

    try {
      result.responsaveis = getResponsaveis_();
    } catch (error) {
      result.errors.push("getResponsaveis_: " + error);
    }

    try {
      result.usuarios = getUsuarios_();
    } catch (error) {
      result.errors.push("getUsuarios_: " + error);
    }

    // Skip heavy data loading on bootstrap - load separately to avoid timeout
    // These will be loaded by getDashboardRecords() call after bootstrap
    result.openErros = [];
    result.solicitacoes = [];
    result.dashboard = { totalSolicitacoes: 0, abertas: 0, emCorrecao: 0, corrigidas: 0, alerta: 0, critico: 0 };

    // Only fail if we don't have essential data (usuario or config)
    if (result.usuario === null) {
      result.errors.push("CRITICAL: usuario is null");
    }
    if (result.config === null) {
      result.errors.push("CRITICAL: config is null");
    }

    const hasCriticalData = result.usuario !== null && result.config !== null;
    result.ok = hasCriticalData;

    safeLogDebug_("getBootstrapData", "loaded", {
      debugId: debugId,
      ok: result.ok,
      errors: result.errors,
      hasCriticalData: hasCriticalData,
      hasUsuario: result.usuario !== null,
      hasConfig: result.config !== null,
      perfil: result.usuario && result.usuario.perfil ? result.usuario.perfil : "",
      setores: result.usuario && result.usuario.setores ? result.usuario.setores : []
    });
    return result;
  } catch (error) {
    safeLogDebug_("getBootstrapData", "error", { debugId: debugId, error: String(error) });
    return { ok: false, error: String(error), debugId: debugId };
  }
}

// ============================================================================
// FUNÇÕES DE LEITURA DE DADOS
// ============================================================================

/**
 * Retorna lista de erros em aberto (pendentes de resposta)
 * Filtrada por permissões de setor do usuário
 *
 * @param {number} [limit=3000] - Limite máximo de registros a retornar
 * @returns {Object} Lista de erros abertos com informações de SLA
 *
 * @description
 * Otimizada para performance com grandes volumes de dados:
 * - Usa lookup maps O(1) para cruzamento de dados
 * - Filtra por setores do usuário
 * - Calcula SLA em batch (não por registro)
 * - Ordena por data de pedido (mais antigos primeiro)
 */
function getOpenErros(limit) {
  const debugId = Utilities.getUuid();
  const startTime = new Date().getTime();
  const MAX_LIMIT = limit || 3000; // Default limit increased to handle more records

  try {
    ensureSheets_();
    const userEmail = Session.getActiveUser().getEmail() || "";
    const usuario = getUsuarioContexto_(userEmail);

    // Debug: count rows in each sheet
    const errosSheet = getSheet_(CONFIG.SHEETS.ERROS);
    const solicitacoesSheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);
    const errosCount = errosSheet ? errosSheet.getLastRow() - 1 : 0;
    const solicitacoesCount = solicitacoesSheet ? solicitacoesSheet.getLastRow() - 1 : 0;

    const openErros = getOpenErrosForUsuario_(usuario, MAX_LIMIT);
    const elapsed = new Date().getTime() - startTime;

    return {
      ok: true,
      debugId: debugId,
      openErros: openErros,
      _debug: {
        usuarioPerfil: usuario ? usuario.perfil : null,
        errosSheetRows: errosCount,
        solicitacoesSheetRows: solicitacoesCount,
        returnedCount: openErros.length,
        limit: MAX_LIMIT,
        elapsedMs: elapsed
      }
    };
  } catch (error) {
    return { ok: false, error: String(error), stack: error.stack, debugId: debugId };
  }
}

/**
 * Retorna registros consolidados para o Dashboard
 * Inclui solicitações, erros e respostas unificados
 *
 * @param {number} [limit=3000] - Limite máximo de registros
 * @returns {Object} Registros consolidados para análise
 *
 * @description
 * Cada registro retornado contém:
 * - Dados da solicitação (requisicao, solicitante, data)
 * - Dados do erro (tipo, detalhamento, setor)
 * - Dados da resposta (responsável, data correção, observações)
 * - Status SLA calculado
 */
function getDashboardRecords(limit) {
  const debugId = Utilities.getUuid();
  const startTime = new Date().getTime();
  const MAX_LIMIT = limit || 3000; // Default limit increased to handle more records

  try {
    ensureSheets_();
    // With O(n) lookup maps optimization, we can load records efficiently
    const records = getDashboardRecords_(MAX_LIMIT);
    const elapsed = new Date().getTime() - startTime;

    return {
      ok: true,
      debugId: debugId,
      records: records,
      _debug: {
        limit: MAX_LIMIT,
        returnedCount: records.length,
        elapsedMs: elapsed
      }
    };
  } catch (error) {
    return { ok: false, error: String(error), stack: error.stack, debugId: debugId };
  }
}

/**
 * Retorna logs de auditoria dentro de um período
 * Registra todas as alterações realizadas no sistema
 *
 * @param {string} dataInicial - Data inicial no formato YYYY-MM-DD
 * @param {string} dataFinal - Data final no formato YYYY-MM-DD
 * @returns {Object} Lista de logs de auditoria
 *
 * @description
 * Cada log contém:
 * - auditId: ID único do registro
 * - userEmail: Quem realizou a ação
 * - actionType: Tipo (CREATE, UPDATE, DELETE, ARCHIVE)
 * - tableName: Tabela afetada
 * - recordKey: Chave do registro
 * - fieldName: Campo alterado
 * - oldValue/newValue: Valores antes/depois
 * - timestamp: Data/hora da alteração
 */
function getAuditLogs(dataInicial, dataFinal) {
  // CRITICAL: Wrap EVERYTHING in try-catch at the very top level
  // and ensure we ALWAYS return a plain object
  var debugId = "";
  try {
    debugId = Utilities.getUuid();
  } catch (e) {
    debugId = "fallback-" + new Date().getTime();
  }

  Logger.log("getAuditLogs START - debugId: " + debugId);
  Logger.log("getAuditLogs params: dataInicial=" + dataInicial + ", dataFinal=" + dataFinal);

  try {
    // Validate parameters
    if (!dataInicial || !dataFinal) {
      Logger.log("getAuditLogs: Missing dates");
      return { ok: false, errors: ["Data inicial e final são obrigatórias."], debugId: debugId, logs: [], total: 0 };
    }

    // Ensure sheets exist
    try {
      ensureSheets_();
    } catch (sheetError) {
      Logger.log("getAuditLogs: ensureSheets_ error: " + String(sheetError));
      // Continue anyway - sheet might already exist
    }

    var auditSheet = null;
    try {
      auditSheet = getSheet_(CONFIG.SHEETS.AUDITORIA);
    } catch (getSheetError) {
      Logger.log("getAuditLogs: getSheet_ error: " + String(getSheetError));
      return { ok: false, errors: ["Erro ao acessar planilha: " + String(getSheetError)], debugId: debugId, logs: [], total: 0 };
    }

    if (!auditSheet) {
      Logger.log("getAuditLogs: Sheet not found");
      return { ok: false, errors: ["Planilha de auditoria não encontrada."], debugId: debugId, logs: [], total: 0 };
    }

    var values = [];
    try {
      values = auditSheet.getDataRange().getValues();
      Logger.log("getAuditLogs: Got " + values.length + " rows");
    } catch (valuesError) {
      Logger.log("getAuditLogs: getValues error: " + String(valuesError));
      return { ok: false, errors: ["Erro ao ler dados: " + String(valuesError)], debugId: debugId, logs: [], total: 0 };
    }

    if (values.length <= 1) {
      Logger.log("getAuditLogs: No data rows");
      return { ok: true, debugId: debugId, logs: [], total: 0 };
    }

    // Parse dates
    var startDate = new Date(dataInicial);
    if (isNaN(startDate.getTime())) {
      var parts = String(dataInicial).split('-');
      if (parts.length === 3) {
        startDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      }
    }

    if (isNaN(startDate.getTime())) {
      Logger.log("getAuditLogs: Invalid start date");
      return { ok: false, errors: ["Data inicial inválida: " + dataInicial], debugId: debugId, logs: [], total: 0 };
    }
    startDate.setHours(0, 0, 0, 0);

    var endDate = new Date(dataFinal);
    if (isNaN(endDate.getTime())) {
      var parts = String(dataFinal).split('-');
      if (parts.length === 3) {
        endDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      }
    }

    if (isNaN(endDate.getTime())) {
      Logger.log("getAuditLogs: Invalid end date");
      return { ok: false, errors: ["Data final inválida: " + dataFinal], debugId: debugId, logs: [], total: 0 };
    }
    endDate.setHours(23, 59, 59, 999);

    Logger.log("getAuditLogs: Date range: " + startDate.toISOString() + " to " + endDate.toISOString());

    var logs = [];

    // Process rows - use simple for loop to avoid any issues
    for (var i = 1; i < values.length; i++) {
      try {
        var row = values[i];
        var timestamp = row[9]; // Date column (index 9)

        if (!timestamp) continue;

        var logDate;
        if (timestamp instanceof Date) {
          logDate = timestamp;
        } else {
          logDate = new Date(timestamp);
        }

        if (isNaN(logDate.getTime())) continue;

        // Filter by date range
        if (logDate >= startDate && logDate <= endDate) {
          // Create plain object with only string/number primitives
          var logEntry = {
            auditId: String(row[0] || ""),
            userEmail: String(row[1] || ""),
            actionType: String(row[2] || ""),
            tableName: String(row[3] || ""),
            recordKey: String(row[4] || ""),
            fieldName: String(row[5] || ""),
            oldValue: String(row[6] !== undefined && row[6] !== null ? row[6] : ""),
            newValue: String(row[7] !== undefined && row[7] !== null ? row[7] : ""),
            ipAddress: String(row[8] || ""),
            timestamp: logDate.toISOString()
          };
          logs.push(logEntry);
        }
      } catch (rowError) {
        Logger.log("getAuditLogs: Error processing row " + i + ": " + String(rowError));
        // Continue processing other rows
      }
    }

    Logger.log("getAuditLogs: Found " + logs.length + " logs in date range");

    // Create result object
    var result = {
      ok: true,
      debugId: debugId,
      logs: logs,
      total: logs.length
    };

    Logger.log("getAuditLogs: Returning result with " + result.total + " logs");
    return result;

  } catch (error) {
    Logger.log("getAuditLogs FATAL ERROR: " + String(error));
    Logger.log("getAuditLogs error stack: " + (error && error.stack ? error.stack : "no stack"));

    // Return a minimal valid response
    return {
      ok: false,
      errors: ["Erro fatal: " + String(error)],
      debugId: debugId,
      logs: [],
      total: 0
    };
  }
}

/**
 * Busca o histórico completo de uma requisição específica
 * Retorna: solicitação, erros, respostas e logs de auditoria
 */
function getHistoricoRequisicao(requisicao) {
  var debugId = "";
  try {
    debugId = Utilities.getUuid();
  } catch (e) {
    debugId = "fallback-" + new Date().getTime();
  }

  Logger.log("getHistoricoRequisicao START - requisicao: " + requisicao);

  try {
    // Validate parameter
    if (!requisicao || String(requisicao).trim() === "") {
      return { ok: false, errors: ["Informe o número da requisição."], debugId: debugId };
    }

    var requisicaoTrimmed = String(requisicao).trim();

    // Ensure sheets exist
    try {
      ensureSheets_();
    } catch (e) {
      // Continue anyway
    }

    // 1. Find the solicitação by requisicao number
    var solicitacoesSheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);
    var solicitacoes = solicitacoesSheet.getDataRange().getValues();

    var solicitacaoData = null;
    var solicitacaoId = null;

    for (var i = 1; i < solicitacoes.length; i++) {
      var reqNum = String(solicitacoes[i][1] || "").trim();
      if (reqNum === requisicaoTrimmed) {
        solicitacaoId = solicitacoes[i][0];
        var dataPedido = solicitacoes[i][3];
        solicitacaoData = {
          id: solicitacaoId,
          requisicao: solicitacoes[i][1],
          solicitante: solicitacoes[i][2],
          dataHoraPedido: dataPedido instanceof Date ? dataPedido.toISOString() : String(dataPedido || ""),
          status: solicitacoes[i][4],
          criadoPorEmail: solicitacoes[i][5],
          criadoEm: solicitacoes[i][6] instanceof Date ? solicitacoes[i][6].toISOString() : String(solicitacoes[i][6] || "")
        };
        break;
      }
    }

    if (!solicitacaoData) {
      return { ok: false, errors: ["Requisição não encontrada: " + requisicaoTrimmed], debugId: debugId };
    }

    Logger.log("getHistoricoRequisicao: Found solicitacao ID=" + solicitacaoId);

    // 2. Get all erros for this solicitação
    var errosSheet = getSheet_(CONFIG.SHEETS.ERROS);
    var erros = errosSheet.getDataRange().getValues();
    var errosData = [];

    for (var i = 1; i < erros.length; i++) {
      if (erros[i][0] === solicitacaoId) {
        errosData.push({
          solicitacaoId: erros[i][0],
          erroSeq: erros[i][1],
          erro: erros[i][2],
          detalhamento: erros[i][3],
          setorLocal: erros[i][4],
          diferencaNoValor: erros[i][5],
          criadoEm: erros[i][6] instanceof Date ? erros[i][6].toISOString() : String(erros[i][6] || "")
        });
      }
    }

    Logger.log("getHistoricoRequisicao: Found " + errosData.length + " erros");

    // 3. Get all respostas for this solicitação
    var respostasSheet = getSheet_(CONFIG.SHEETS.RESPOSTAS);
    var respostas = respostasSheet.getDataRange().getValues();
    var respostasData = [];

    for (var i = 1; i < respostas.length; i++) {
      if (respostas[i][1] === solicitacaoId) {
        var dataCorrecao = respostas[i][9];
        var criadoEm = respostas[i][10];
        respostasData.push({
          idResposta: respostas[i][0],
          solicitacaoId: respostas[i][1],
          erroSeq: respostas[i][2],
          nomeResponsavel: respostas[i][3],
          emailResponsavel: respostas[i][4],
          correcaoFinalizada: respostas[i][5],
          houveDiferencaValor: respostas[i][6],
          diferencaValorResposta: respostas[i][7],
          observacoes: respostas[i][8],
          dataHoraCorrecao: dataCorrecao instanceof Date ? dataCorrecao.toISOString() : String(dataCorrecao || ""),
          criadoEm: criadoEm instanceof Date ? criadoEm.toISOString() : String(criadoEm || "")
        });
      }
    }

    Logger.log("getHistoricoRequisicao: Found " + respostasData.length + " respostas");

    // 4. Get all audit logs related to this solicitação
    var auditSheet = getSheet_(CONFIG.SHEETS.AUDITORIA);
    var audits = auditSheet.getDataRange().getValues();
    var auditLogs = [];

    for (var i = 1; i < audits.length; i++) {
      var recordKey = String(audits[i][4] || "");
      // Match by solicitacaoId in recordKey (e.g., "id_solicitacao=xxx" or "xxx_1")
      if (recordKey.indexOf(solicitacaoId) !== -1 || recordKey.indexOf(requisicaoTrimmed) !== -1) {
        var timestamp = audits[i][9];
        auditLogs.push({
          auditId: String(audits[i][0] || ""),
          userEmail: String(audits[i][1] || ""),
          actionType: String(audits[i][2] || ""),
          tableName: String(audits[i][3] || ""),
          recordKey: recordKey,
          fieldName: String(audits[i][5] || ""),
          oldValue: String(audits[i][6] !== undefined && audits[i][6] !== null ? audits[i][6] : ""),
          newValue: String(audits[i][7] !== undefined && audits[i][7] !== null ? audits[i][7] : ""),
          ipAddress: String(audits[i][8] || ""),
          timestamp: timestamp instanceof Date ? timestamp.toISOString() : String(timestamp || "")
        });
      }
    }

    // Sort audit logs by timestamp (most recent first)
    auditLogs.sort(function(a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });

    Logger.log("getHistoricoRequisicao: Found " + auditLogs.length + " audit logs");

    return {
      ok: true,
      debugId: debugId,
      solicitacao: solicitacaoData,
      erros: errosData,
      respostas: respostasData,
      auditLogs: auditLogs
    };

  } catch (error) {
    Logger.log("getHistoricoRequisicao ERROR: " + String(error));
    return {
      ok: false,
      errors: ["Erro ao buscar histórico: " + String(error)],
      debugId: debugId
    };
  }
}

function getListas() {
  const debugId = Utilities.getUuid();
  safeLogDebug_("getListas", "start", { debugId: debugId });
  try {
    ensureSheets_();
    const listas = getListasData_();
    safeLogDebug_("getListas", "loaded", {
      debugId: debugId,
      colaboradores: listas.colaboradores.length,
      setores: listas.setores.length
    });
    return { ok: true, debugId: debugId, listas: listas };
  } catch (error) {
    safeLogDebug_("getListas", "error", { debugId: debugId, error: String(error) });
    return { ok: false, error: String(error), debugId: debugId };
  }
}

function getConfig() {
  const debugId = Utilities.getUuid();
  safeLogDebug_("getConfig", "start", { debugId: debugId });
  try {
    ensureSheets_();
    const config = getConfigData_();
    safeLogDebug_("getConfig", "loaded", {
      debugId: debugId,
      solicitantes: config.pharmaceuticas.length,
      setores: config.setores.length,
      erros: config.erros.length
    });
    return { ok: true, debugId: debugId, config: config };
  } catch (error) {
    safeLogDebug_("getConfig", "error", { debugId: debugId, error: String(error) });
    return { ok: false, error: String(error), debugId: debugId };
  }
}

function getUsuarios() {
  const debugId = Utilities.getUuid();
  safeLogDebug_("getUsuarios", "start", { debugId: debugId });
  try {
    ensureSheets_();
    const usuarios = getUsuarios_();
    safeLogDebug_("getUsuarios", "loaded", { debugId: debugId, total: usuarios.length });
    return { ok: true, debugId: debugId, usuarios: usuarios };
  } catch (error) {
    safeLogDebug_("getUsuarios", "error", { debugId: debugId, error: String(error) });
    return { ok: false, error: String(error), debugId: debugId };
  }
}

function getUsuario() {
  const debugId = Utilities.getUuid();
  safeLogDebug_("getUsuario", "start", { debugId: debugId });
  try {
    ensureSheets_();
    const userEmail = Session.getActiveUser().getEmail() || "";
    const usuario = getUsuarioContexto_(userEmail);
    safeLogDebug_("getUsuario", "loaded", {
      debugId: debugId,
      email: userEmail,
      perfil: usuario ? usuario.perfil : null
    });
    return { ok: true, debugId: debugId, usuario: usuario };
  } catch (error) {
    safeLogDebug_("getUsuario", "error", { debugId: debugId, error: String(error) });
    return { ok: false, error: String(error), debugId: debugId };
  }
}

function debugSnapshot() {
  const debugId = Utilities.getUuid();
  safeLogDebug_("debugSnapshot", "start", { debugId: debugId });
  try {
    ensureSheets_();
    const ss = getSpreadsheet_();
    const sheetNames = ss.getSheets().map((s) => s.getName());
    const listasData = getListasData_();
    const config = getConfigData_();
    const usuarios = getUsuarios_();
    const snapshot = {
      debugId: debugId,
      sheets: sheetNames,
      listas: {
        colaboradores: listasData.colaboradores.length,
        setores: listasData.setores.length
      },
      config: {
        solicitantes: config.pharmaceuticas.length,
        setores: config.setores.length,
        erros: config.erros.length
      },
      usuarios: usuarios.length
    };
    safeLogDebug_("debugSnapshot", "loaded", snapshot);
    return { ok: true, debugId: debugId, snapshot: snapshot };
  } catch (error) {
    safeLogDebug_("debugSnapshot", "error", { debugId: debugId, error: String(error) });
    return { ok: false, error: String(error), debugId: debugId };
  }
}

// ============================================================================
// FUNÇÕES DE ESCRITA DE DADOS
// ============================================================================

/**
 * Cria uma nova solicitação de conserto
 *
 * @param {Object} payload - Dados da solicitação
 * @param {string} payload.requisicao - Número da requisição
 * @param {string} payload.solicitante - Nome do solicitante
 * @param {string} payload.erro - Tipo de erro
 * @param {string} payload.detalhamento - Descrição detalhada
 * @param {string} payload.setorLocal - Setor/Local
 * @param {string} payload.diferencaNoValor - Se houve diferença de valor (SIM/NÃO)
 * @returns {Object} ID da solicitação criada
 *
 * @sideeffect Adiciona registro nas planilhas Solicitações e Erros
 * @sideeffect Registra ação na auditoria
 */
function submitRequest(payload) {
  const debugId = Utilities.getUuid();
  safeLogDebug_("submitRequest", "start", { debugId: debugId });
  ensureSheets_();
  const config = getConfigData_();

  payload.dataHoraPedido = new Date().toISOString();
  safeLogDebug_("submitRequest", "payload", payload);
  const errors = validateRequestPayload_(payload, config);
  if (errors.length) {
    safeLogDebug_("submitRequest", "validation_failed", errors);
    return { ok: false, errors: errors, debugId: debugId };
  }

  const userEmail = getEmailValidado_();
  checkRateLimit_(userEmail);
  const result = insertSolicitacaoEErro_(payload, userEmail);
  safeLogDebug_("submitRequest", "inserted", result);
  return { ok: true, data: result, debugId: debugId };
}

/**
 * Registra resposta/correção para uma solicitação
 *
 * @param {Object} payload - Dados da resposta
 * @param {string} payload.solicitacaoId - ID da solicitação
 * @param {number} payload.erroSeq - Sequência do erro
 * @param {string} payload.responsavel - Nome do responsável pela correção
 * @param {string} payload.correcaoFinalizada - Se a correção foi finalizada (SIM/NÃO)
 * @param {string} [payload.diferencaValorResposta] - Valor da diferença (se houver)
 * @param {string} [payload.observacoes] - Observações adicionais
 * @param {boolean} [payload.isUpdate] - Se é atualização de resposta existente
 * @returns {Object} Confirmação da operação
 *
 * @sideeffect Adiciona/atualiza registro na planilha Respostas
 * @sideeffect Atualiza status da solicitação para CORRIGIDO
 * @sideeffect Registra ação na auditoria
 */
function submitResponse(payload) {
  const debugId = Utilities.getUuid();
  safeLogDebug_("submitResponse", "start", { debugId: debugId });
  ensureSheets_();
  if (!payload || typeof payload !== "object") {
    return { ok: false, errors: ["Payload inválido."], debugId: debugId };
  }
  const config = getConfigData_();
  payload.dataHoraCorrecao = new Date().toISOString();
  safeLogDebug_("submitResponse", "payload", payload);
  const errors = validateResponsePayload_(payload, config);
  if (errors.length) {
    safeLogDebug_("submitResponse", "validation_failed", errors);
    return { ok: false, errors: errors, debugId: debugId };
  }

  const userEmail = getEmailValidado_();
  checkRateLimit_(userEmail);

  // Check if this is an update operation
  if (payload.isUpdate === true) {
    safeLogDebug_("submitResponse", "update_mode", { solicitacaoId: payload.solicitacaoId, erroSeq: payload.erroSeq });
    const result = updateResposta_(payload, userEmail);
    if (!result) {
      return { ok: false, errors: ["Resposta não encontrada para atualização."], debugId: debugId };
    }
    safeLogDebug_("submitResponse", "updated", result);
    return { ok: true, data: result, debugId: debugId };
  }

  // Otherwise, insert new response
  const result = insertResposta_(payload, userEmail);
  safeLogDebug_("submitResponse", "inserted", result);
  return { ok: true, data: result, debugId: debugId };
}

function deleteSolicitacao(solicitacaoId) {
  const debugId = Utilities.getUuid();
  safeLogDebug_("deleteSolicitacao", "start", { debugId: debugId, solicitacaoId: solicitacaoId });
  ensureSheets_();

  const userEmail = getEmailValidado_();
  checkRateLimit_(userEmail);
  const usuario = getUsuarioContexto_(userEmail);

  // Check permissions: only ADMIN or CONFERENTE can delete
  if (usuario.perfil !== "ADMIN" && usuario.perfil !== "CONFERENTE") {
    return { ok: false, errors: ["Apenas ADMIN e CONFERENTE podem excluir solicitações."], debugId: debugId };
  }

  if (!solicitacaoId) {
    return { ok: false, errors: ["ID da solicitação inválido."], debugId: debugId };
  }

  const result = deleteSolicitacao_(solicitacaoId);
  if (!result) {
    return { ok: false, errors: ["Solicitação não encontrada."], debugId: debugId };
  }

  safeLogDebug_("deleteSolicitacao", "deleted", { solicitacaoId: solicitacaoId });
  return { ok: true, debugId: debugId };
}

/**
 * Atualiza uma solicitação existente
 * Permissões:
 * - ADMIN/CONFERENTE: Edita solicitação + resposta
 * - RESPOSTA: Edita apenas resposta
 * - ESPECTADOR: Não pode editar
 * Registra as alterações na auditoria com antes/depois
 */
function updateSolicitacao(payload) {
  const debugId = Utilities.getUuid();

  try {
    ensureSheets_();

    const userEmail = Session.getActiveUser().getEmail() || "";
    const usuario = getUsuarioContexto_(userEmail);
    const perfil = usuario.perfil ? usuario.perfil.toUpperCase() : "";

    // Check permissions based on profile
    // ADMIN/CONFERENTE can edit solicitation AND response; RESPOSTA can edit response only
    const canEditSolicitation = perfil === "ADMIN" || perfil === "CONFERENTE";
    const canEditResponse = perfil === "ADMIN" || perfil === "CONFERENTE" || perfil === "RESPOSTA";

    if (!canEditSolicitation && !canEditResponse) {
      return { ok: false, errors: ["Sem permissão para editar."], debugId: debugId };
    }

    // Validate required fields
    const errors = [];
    if (!payload.solicitacaoId) errors.push("ID da solicitação é obrigatório.");

    // Only validate solicitation fields if user can edit them
    if (canEditSolicitation && payload.requisicao !== undefined) {
      if (!payload.requisicao) errors.push("Requisição é obrigatória.");
      if (!payload.solicitante) errors.push("Solicitante é obrigatório.");
      if (!payload.erro) errors.push("Tipo de erro é obrigatório.");
      if (!payload.setorLocal) errors.push("Setor/Local é obrigatório.");
      if (!payload.diferencaNoValor) errors.push("Diferença no valor é obrigatório.");
      if (!payload.detalhamento) errors.push("Detalhamento é obrigatório.");
    }

    if (errors.length) {
      return { ok: false, errors: errors, debugId: debugId };
    }

    // Find and update the solicitation
    const solicitacoesSheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);
    const solicitacoes = solicitacoesSheet.getDataRange().getValues();
    const headers = solicitacoes[0];

    // Find column indices - column name is "id_solicitacao" not "id"
    const colId = headers.indexOf("id_solicitacao");
    const colRequisicao = headers.indexOf("requisicao");
    const colSolicitante = headers.indexOf("solicitante");
    const colSetorLocal = headers.indexOf("setorLocal");
    const colDiferencaNoValor = headers.indexOf("diferencaNoValor");

    safeLogDebug_("updateSolicitacao", "columns", { colId: colId, headers: headers.slice(0, 10) });

    let solicitacaoRowIndex = -1;
    const searchId = String(payload.solicitacaoId).trim();
    for (let i = 1; i < solicitacoes.length; i++) {
      const rowId = String(solicitacoes[i][colId]).trim();
      if (rowId === searchId) {
        solicitacaoRowIndex = i + 1; // 1-based row number
        safeLogDebug_("updateSolicitacao", "found", { row: solicitacaoRowIndex, rowId: rowId });
        break;
      }
    }

    if (solicitacaoRowIndex === -1) {
      safeLogDebug_("updateSolicitacao", "notFound", { searchId: searchId, colId: colId, totalRows: solicitacoes.length });
      return { ok: false, errors: ["Solicitação não encontrada. ID: " + searchId], debugId: debugId };
    }

    const originalValues = payload.originalValues || {};
    const recordKey = payload.solicitacaoId;

    // Update Solicitacoes sheet (only if user can edit solicitation)
    if (canEditSolicitation) {
      if (originalValues.requisicao !== payload.requisicao) {
        solicitacoesSheet.getRange(solicitacaoRowIndex, colRequisicao + 1).setValue(payload.requisicao);
        auditLog_("UPDATE", "Solicitacoes", recordKey, "requisicao", originalValues.requisicao, payload.requisicao);
      }
      if (originalValues.solicitante !== payload.solicitante) {
        solicitacoesSheet.getRange(solicitacaoRowIndex, colSolicitante + 1).setValue(payload.solicitante);
        auditLog_("UPDATE", "Solicitacoes", recordKey, "solicitante", originalValues.solicitante, payload.solicitante);
      }
      if (originalValues.setorLocal !== payload.setorLocal) {
        solicitacoesSheet.getRange(solicitacaoRowIndex, colSetorLocal + 1).setValue(payload.setorLocal);
        auditLog_("UPDATE", "Solicitacoes", recordKey, "setorLocal", originalValues.setorLocal, payload.setorLocal);
      }
      if (originalValues.diferencaNoValor !== payload.diferencaNoValor) {
        solicitacoesSheet.getRange(solicitacaoRowIndex, colDiferencaNoValor + 1).setValue(payload.diferencaNoValor);
        auditLog_("UPDATE", "Solicitacoes", recordKey, "diferencaNoValor", originalValues.diferencaNoValor, payload.diferencaNoValor);
      }

      // Update Erros sheet
      const errosSheet = getSheet_(CONFIG.SHEETS.ERROS);
      const erros = errosSheet.getDataRange().getValues();
      const errosHeaders = erros[0];

      const erroColSolId = errosHeaders.indexOf("solicitacaoId");
      const erroColSeq = errosHeaders.indexOf("erroSeq");
      const erroColErro = errosHeaders.indexOf("erro");
      const erroColDetalhamento = errosHeaders.indexOf("detalhamento");
      const erroColSetorLocal = errosHeaders.indexOf("setorLocal");
      const erroColDiferencaNoValor = errosHeaders.indexOf("diferencaNoValor");

      for (let i = 1; i < erros.length; i++) {
        if (erros[i][erroColSolId] === payload.solicitacaoId &&
            String(erros[i][erroColSeq]) === String(payload.erroSeq)) {
          const erroRowIndex = i + 1;
          const erroRecordKey = payload.solicitacaoId + "_" + payload.erroSeq;

          if (originalValues.erro !== payload.erro) {
            errosSheet.getRange(erroRowIndex, erroColErro + 1).setValue(payload.erro);
            auditLog_("UPDATE", "Erros", erroRecordKey, "erro", originalValues.erro, payload.erro);
          }
          if (originalValues.detalhamento !== payload.detalhamento) {
            errosSheet.getRange(erroRowIndex, erroColDetalhamento + 1).setValue(payload.detalhamento);
            auditLog_("UPDATE", "Erros", erroRecordKey, "detalhamento", originalValues.detalhamento, payload.detalhamento);
          }
          if (erroColSetorLocal >= 0 && originalValues.setorLocal !== payload.setorLocal) {
            errosSheet.getRange(erroRowIndex, erroColSetorLocal + 1).setValue(payload.setorLocal);
          }
          if (erroColDiferencaNoValor >= 0 && originalValues.diferencaNoValor !== payload.diferencaNoValor) {
            errosSheet.getRange(erroRowIndex, erroColDiferencaNoValor + 1).setValue(payload.diferencaNoValor);
          }
          break;
        }
      }
    }

    // Update Respostas sheet (only if user can edit response and responseData is provided)
    if (canEditResponse && payload.responseData) {
      const respostasSheet = getSheet_(CONFIG.SHEETS.RESPOSTAS);
      const respostas = respostasSheet.getDataRange().getValues();
      const respHeaders = respostas[0];

      // Find column indices for Respostas
      const respColId = respHeaders.indexOf("id_resposta") >= 0 ? respHeaders.indexOf("id_resposta") : 0;
      const respColSolId = respHeaders.indexOf("solicitacaoId") >= 0 ? respHeaders.indexOf("solicitacaoId") : 1;
      const respColErroSeq = respHeaders.indexOf("erroSeq") >= 0 ? respHeaders.indexOf("erroSeq") : 2;
      const respColResponsavel = respHeaders.indexOf("nomeResponsavel") >= 0 ? respHeaders.indexOf("nomeResponsavel") : 3;
      const respColCorrecaoFinalizada = respHeaders.indexOf("correcaoFinalizada") >= 0 ? respHeaders.indexOf("correcaoFinalizada") : 5;
      const respColDiferencaValor = respHeaders.indexOf("diferenca_valor_resposta") >= 0 ? respHeaders.indexOf("diferenca_valor_resposta") : 7;
      const respColObservacoes = respHeaders.indexOf("observacoes") >= 0 ? respHeaders.indexOf("observacoes") : 8;
      const respColDataCorrecao = respHeaders.indexOf("dataHoraCorrecao") >= 0 ? respHeaders.indexOf("dataHoraCorrecao") : 9;

      const respostaRecordKey = payload.solicitacaoId + "_" + payload.erroSeq;
      const origResp = payload.originalResponseValues || {};
      const newResp = payload.responseData;

      // Find existing response row
      let respostaRowIndex = -1;
      for (let i = 1; i < respostas.length; i++) {
        if (respostas[i][respColSolId] === payload.solicitacaoId &&
            String(respostas[i][respColErroSeq]) === String(payload.erroSeq)) {
          respostaRowIndex = i + 1;
        }
      }

      // If response exists, update it
      if (respostaRowIndex > 0) {
        // Update responsavel
        if (origResp.responsavel !== newResp.responsavel) {
          respostasSheet.getRange(respostaRowIndex, respColResponsavel + 1).setValue(newResp.responsavel || "");
          auditLog_("UPDATE", "Respostas", respostaRecordKey, "responsavel", origResp.responsavel, newResp.responsavel);
        }
        // Update dataHoraCorrecao
        const newDataCorrecao = newResp.dataHoraCorrecao ? new Date(newResp.dataHoraCorrecao) : "";
        const origDataStr = origResp.dataHoraCorrecao ? new Date(origResp.dataHoraCorrecao).toISOString() : "";
        const newDataStr = newDataCorrecao ? newDataCorrecao.toISOString() : "";
        if (origDataStr !== newDataStr) {
          respostasSheet.getRange(respostaRowIndex, respColDataCorrecao + 1).setValue(newDataCorrecao);
          auditLog_("UPDATE", "Respostas", respostaRecordKey, "dataHoraCorrecao", origResp.dataHoraCorrecao, newResp.dataHoraCorrecao);
        }
        // Update correcaoFinalizada
        if (origResp.correcaoFinalizada !== newResp.correcaoFinalizada) {
          respostasSheet.getRange(respostaRowIndex, respColCorrecaoFinalizada + 1).setValue(newResp.correcaoFinalizada || "");
          auditLog_("UPDATE", "Respostas", respostaRecordKey, "correcaoFinalizada", origResp.correcaoFinalizada, newResp.correcaoFinalizada);
        }
        // Update diferencaValorResposta
        if (origResp.diferencaValorResposta !== newResp.diferencaValorResposta) {
          respostasSheet.getRange(respostaRowIndex, respColDiferencaValor + 1).setValue(newResp.diferencaValorResposta || "");
          auditLog_("UPDATE", "Respostas", respostaRecordKey, "diferencaValorResposta", origResp.diferencaValorResposta, newResp.diferencaValorResposta);
        }
        // Update observacoes
        if (origResp.observacoes !== newResp.observacoes) {
          respostasSheet.getRange(respostaRowIndex, respColObservacoes + 1).setValue(newResp.observacoes || "");
          auditLog_("UPDATE", "Respostas", respostaRecordKey, "observacoes", origResp.observacoes, newResp.observacoes);
        }
      } else if (newResp.responsavel || newResp.dataHoraCorrecao || newResp.observacoes) {
        // Create new response if data is provided and no existing response
        const newRespostaId = Utilities.getUuid();
        const now = new Date();
        respostasSheet.appendRow([
          newRespostaId,
          payload.solicitacaoId,
          payload.erroSeq,
          newResp.responsavel || "",
          userEmail,
          newResp.correcaoFinalizada || "",
          "", // houveDiferencaValor (legacy, unused)
          newResp.diferencaValorResposta || "",
          newResp.observacoes || "",
          newResp.dataHoraCorrecao ? new Date(newResp.dataHoraCorrecao) : "",
          now
        ]);
        auditLog_("CREATE", "Respostas", respostaRecordKey, null, null, "Nova resposta criada via edição");
      }
    }

    return { ok: true, debugId: debugId };

  } catch (error) {
    return { ok: false, errors: ["Erro interno: " + String(error)], debugId: debugId };
  }
}

// ============================================================================
// FUNÇÕES DE CONFIGURAÇÃO DO SISTEMA
// ============================================================================

/**
 * Adiciona novo item de configuração (solicitante, setor ou tipo de erro)
 *
 * @param {string} type - Tipo do item: "solicitante", "setor" ou "erro"
 * @param {string} value - Valor a ser adicionado
 * @returns {Object} Confirmação da operação
 *
 * @requires Permissão ADMIN
 * @sideeffect Adiciona registro na planilha correspondente
 * @sideeffect Invalida cache de configurações
 */
// Tipos de configuração permitidos (whitelist explícita)
const VALID_CONFIG_TYPES_ = ["solicitante", "setor", "erro"];

function addConfigItem(type, value) {
  const debugId = Utilities.getUuid();
  safeLogDebug_("addConfigItem", "start", { debugId: debugId });
  ensureSheets_();
  checkRateLimit_(getEmailValidado_());

  // Validar tipo contra whitelist
  if (VALID_CONFIG_TYPES_.indexOf(String(type || "")) === -1) {
    return { ok: false, errors: ["Tipo inválido."], debugId: debugId };
  }

  const trimmed = String(value || "").trim();
  if (!trimmed) {
    return { ok: false, errors: ["Informe um valor válido."], debugId: debugId };
  }
  safeLogDebug_("addConfigItem", "payload", { type: type });

  let sheetName = "";
  if (type === "solicitante") sheetName = CONFIG.SHEETS.SOLICITANTES;
  if (type === "setor") sheetName = CONFIG.SHEETS.SETORES;
  if (type === "erro") sheetName = CONFIG.SHEETS.ERROS_CADASTRO;

  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  // Comparação case-insensitive para evitar duplicatas como "Erro" e "erro"
  const trimmedLower = trimmed.toLowerCase();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || "").trim().toLowerCase() === trimmedLower) {
      return { ok: false, errors: ["Item já cadastrado."], debugId: debugId };
    }
  }
  sheet.appendRow([trimmed]);
  if (type === "erro") sortErrosCadastro_();
  invalidateConfigCache_();
  return { ok: true, debugId: debugId };
}

function deleteConfigItem(type, value) {
  const debugId = Utilities.getUuid();
  safeLogDebug_("deleteConfigItem", "start", { debugId: debugId });
  ensureSheets_();
  checkRateLimit_(getEmailValidado_());
  const trimmed = String(value || "").trim();
  if (!trimmed) {
    return { ok: false, errors: ["Valor inválido."], debugId: debugId };
  }
  safeLogDebug_("deleteConfigItem", "payload", { type: type, value: trimmed });

  let sheetName = "";
  if (type === "solicitante") sheetName = CONFIG.SHEETS.SOLICITANTES;
  if (type === "setor") sheetName = CONFIG.SHEETS.SETORES;
  if (type === "erro") sheetName = CONFIG.SHEETS.ERROS_CADASTRO;
  if (!sheetName) {
    return { ok: false, errors: ["Tipo inválido."], debugId: debugId };
  }

  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  let rowToDelete = -1;

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === trimmed) {
      rowToDelete = i + 1;
      break;
    }
  }

  if (rowToDelete === -1) {
    return { ok: false, errors: ["Item não encontrado."], debugId: debugId };
  }

  sheet.deleteRow(rowToDelete);
  invalidateConfigCache_(); // Clear cache after deleting item
  safeLogDebug_("deleteConfigItem", "deleted", { row: rowToDelete });
  return { ok: true, debugId: debugId };
}

function updateErrorClassification(nomeErro, classificacao) {
  const debugId = Utilities.getUuid();
  safeLogDebug_("updateErrorClassification", "start", { debugId: debugId });
  ensureSheets_();
  checkRateLimit_(getEmailValidado_());

  const VALID_CLASSIFICATIONS_ = ["NA", "Leve", "Grave", "Gravíssimo"];

  const trimmedNome = String(nomeErro || "").trim();
  const trimmedClass = String(classificacao || "").trim();

  if (!trimmedNome) {
    return { ok: false, errors: ["Nome do erro inválido."], debugId: debugId };
  }
  if (!trimmedClass || VALID_CLASSIFICATIONS_.indexOf(trimmedClass) === -1) {
    return { ok: false, errors: ["Classificação inválida. Use: NA, Leve, Grave ou Gravíssimo."], debugId: debugId };
  }

  safeLogDebug_("updateErrorClassification", "payload", {
    nomeErro: trimmedNome,
    classificacao: trimmedClass
  });

  const sheet = getSheet_(CONFIG.SHEETS.ERROS_CADASTRO);
  const values = sheet.getDataRange().getValues();
  let rowToUpdate = -1;

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === trimmedNome) {
      rowToUpdate = i + 1;
      break;
    }
  }

  if (rowToUpdate === -1) {
    return { ok: false, errors: ["Erro não encontrado."], debugId: debugId };
  }

  sheet.getRange(rowToUpdate, 2).setValue(trimmedClass);
  sortErrosCadastro_();
  invalidateConfigCache_(); // Clear cache after updating classification
  safeLogDebug_("updateErrorClassification", "updated", { row: rowToUpdate });
  return { ok: true, debugId: debugId };
}

function sortErrosCadastro_() {
  const sheet = getSheet_(CONFIG.SHEETS.ERROS_CADASTRO);
  const dataRange = sheet.getDataRange();
  const allValues = dataRange.getValues();
  if (allValues.length <= 2) return;
  const header = allValues[0];
  const dataRows = allValues.slice(1);
  const classOrder = { 'NA': 0, 'Leve': 1, 'Grave': 2, 'Gravíssimo': 3 };
  dataRows.sort(function(a, b) {
    const ca = classOrder[String(a[1] || '').trim()];
    const cb = classOrder[String(b[1] || '').trim()];
    const orderA = ca !== undefined ? ca : 99;
    const orderB = cb !== undefined ? cb : 99;
    if (orderA !== orderB) return orderA - orderB;
    return String(a[0] || '').localeCompare(String(b[0] || ''), 'pt-BR', { sensitivity: 'base' });
  });
  dataRange.setValues([header].concat(dataRows));
}

// ============================================================================
// FUNÇÕES DE GESTÃO DE USUÁRIOS
// ============================================================================

/**
 * Adiciona novo usuário ao sistema
 *
 * @param {Object} payload - Dados do usuário
 * @param {string} payload.email - Email do usuário (chave única)
 * @param {string} payload.nome - Nome completo
 * @param {string} payload.perfil - Perfil: ADMIN, CONFERENTE, RESPOSTA ou ESPECTADOR
 * @param {string} payload.setores - Setores permitidos (separados por vírgula ou "*" para todos)
 * @returns {Object} Confirmação da operação
 *
 * @requires Permissão ADMIN
 * @note ADMIN e CONFERENTE recebem automaticamente acesso a todos os setores (*)
 * @sideeffect Adiciona registro na planilha Usuarios
 * @sideeffect Registra ação na auditoria
 */
function addUsuario(payload) {
  const debugId = Utilities.getUuid();
  safeLogDebug_("addUsuario", "start", { debugId: debugId });
  ensureSheets_();
  const email = String(payload.email || "").trim();
  const nome = String(payload.nome || "").trim();
  const perfil = String(payload.perfil || "").trim();
  let setores = String(payload.setores || "").trim();
  safeLogDebug_("addUsuario", "payload", payload);

  const errors = [];
  if (!email) errors.push("Informe o email.");
  if (!nome) errors.push("Informe o nome.");
  if (!perfil) errors.push("Informe o perfil.");
  // ADMIN e CONFERENTE têm acesso a todos os setores
  if (perfil === "ADMIN" || perfil === "CONFERENTE") setores = "*";
  if (!setores) errors.push("Informe os setores.");
  if (errors.length) return { ok: false, errors: errors, debugId: debugId };

  const sheet = getSheet_(CONFIG.SHEETS.USUARIOS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === email) {
      return { ok: false, errors: ["Email já cadastrado."], debugId: debugId };
    }
  }

  sheet.appendRow([email, nome, perfil, setores]);
  invalidateUsuarioCache_(email);
  auditLog_("CREATE", "usuarios", "email=" + email, null, null, null);
  return { ok: true, debugId: debugId };
}

function updateUsuario(payload) {
  const debugId = Utilities.getUuid();
  safeLogDebug_("updateUsuario", "start", { debugId: debugId });
  ensureSheets_();
  const email = String(payload.email || "").trim();
  const nome = String(payload.nome || "").trim();
  const perfil = String(payload.perfil || "").trim();
  let setores = String(payload.setores || "").trim();
  const errors = [];
  if (!email) errors.push("Informe o email.");
  if (!nome) errors.push("Informe o nome.");
  if (!perfil) errors.push("Informe o perfil.");
  // ADMIN e CONFERENTE têm acesso a todos os setores
  if (perfil === "ADMIN" || perfil === "CONFERENTE") setores = "*";
  if (!setores) errors.push("Informe os setores.");
  if (errors.length) return { ok: false, errors: errors, debugId: debugId };

  const sheet = getSheet_(CONFIG.SHEETS.USUARIOS);
  const values = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === email) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) {
    return { ok: false, errors: ["Usuário não encontrado."], debugId: debugId };
  }

  const oldRow = values[rowIndex - 1];
  sheet.getRange(rowIndex, 1, 1, 4).setValues([[email, nome, perfil, setores]]);
  invalidateUsuarioCache_(email);
  auditLog_("UPDATE", "usuarios", "email=" + email, "nome", oldRow[1], nome);
  auditLog_("UPDATE", "usuarios", "email=" + email, "perfil", oldRow[2], perfil);
  auditLog_("UPDATE", "usuarios", "email=" + email, "setores", oldRow[3], setores);
  safeLogDebug_("updateUsuario", "updated", { debugId: debugId });
  return { ok: true, debugId: debugId };
}

function deleteUsuario(payload) {
  const debugId = Utilities.getUuid();
  safeLogDebug_("deleteUsuario", "start", { debugId: debugId });
  ensureSheets_();
  const email = String(payload.email || "").trim();
  if (!email) return { ok: false, errors: ["Informe o email."], debugId: debugId };
  const sheet = getSheet_(CONFIG.SHEETS.USUARIOS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === email) {
      sheet.deleteRow(i + 1);
      invalidateUsuarioCache_(email);
      auditLog_("DELETE", "usuarios", "email=" + email, null, null, null);
      safeLogDebug_("deleteUsuario", "deleted", { debugId: debugId });
      return { ok: true, debugId: debugId };
    }
  }
  return { ok: false, errors: ["Usuário não encontrado."], debugId: debugId };
}

function limparCache() {
  CacheService.getScriptCache().removeAll([]);
  safeLogDebug_("limparCache", "cache cleared");
  return { ok: true };
}

function salvarPastaBackup(payload) {
  const debugId = Utilities.getUuid();
  const folderId = String(payload.folderId || "").trim();
  if (!folderId) {
    return { ok: false, errors: ["Informe o ID da pasta."], debugId: debugId };
  }
  setConfigGeral_("pasta_backup_id", folderId);
  safeLogDebug_("salvarPastaBackup", "saved", { folderId: folderId, debugId: debugId });
  return { ok: true, debugId: debugId };
}

function atualizarLimiares(payload) {
  const debugId = Utilities.getUuid();
  const minutosAlerta = Number(payload.minutosAlerta);
  const minutosCritico = Number(payload.minutosCritico);
  if (!minutosAlerta || !minutosCritico) {
    return { ok: false, errors: ["Informe os minutos de alerta e crítico."], debugId: debugId };
  }
  safeLogDebug_("atualizarLimiares", "payload", payload);
  const sheet = getSheet_(CONFIG.SHEETS.LIMIARES);
  if (sheet.getLastRow() < 2) {
    sheet.appendRow([
      minutosAlerta,
      minutosCritico,
      CONFIG.DEFAULTS.TIMEZONE,
      Session.getActiveUser().getEmail(),
      new Date()
    ]);
  } else {
    sheet.getRange(2, 1, 1, 5).setValues([[
      minutosAlerta,
      minutosCritico,
      CONFIG.DEFAULTS.TIMEZONE,
      Session.getActiveUser().getEmail(),
      new Date()
    ]]);
  }
  invalidateConfigCache_(); // Limpar cache para carregar novos valores
  return { ok: true, debugId: debugId };
}

function testarPastaBackup() {
  const debugId = Utilities.getUuid();
  Logger.log("testarPastaBackup - START");

  try {
    // Ler ID salvo
    const pastaBackupIdRaw = getConfigGeral_("pasta_backup_id");
    Logger.log("testarPastaBackup - ID raw: [" + pastaBackupIdRaw + "]");
    Logger.log("testarPastaBackup - ID length: " + (pastaBackupIdRaw ? pastaBackupIdRaw.length : 0));
    Logger.log("testarPastaBackup - ID type: " + typeof pastaBackupIdRaw);

    // Limpar ID
    const pastaBackupId = String(pastaBackupIdRaw || "").trim();
    Logger.log("testarPastaBackup - ID limpo: [" + pastaBackupId + "]");
    Logger.log("testarPastaBackup - ID limpo length: " + pastaBackupId.length);

    if (!pastaBackupId) {
      return { ok: false, errors: ["ID da pasta não configurado."], debugId: debugId };
    }

    // Testar acesso à pasta
    try {
      const folder = DriveApp.getFolderById(pastaBackupId);
      const folderName = folder.getName();
      Logger.log("testarPastaBackup - Pasta encontrada: " + folderName);
      return {
        ok: true,
        debugId: debugId,
        idSalvo: pastaBackupIdRaw,
        idLimpo: pastaBackupId,
        nomePasta: folderName
      };
    } catch (e) {
      Logger.log("testarPastaBackup - Erro ao acessar pasta: " + String(e));
      return {
        ok: false,
        errors: ["Erro ao acessar pasta: " + String(e)],
        debugId: debugId,
        idSalvo: pastaBackupIdRaw,
        idLimpo: pastaBackupId
      };
    }
  } catch (error) {
    Logger.log("testarPastaBackup - Erro geral: " + String(error));
    return { ok: false, errors: [String(error)], debugId: debugId };
  }
}

// ============================================================================
// FUNÇÕES DE MANUTENÇÃO E ARQUIVAMENTO
// ============================================================================

/**
 * Executa arquivamento mensal de dados para Looker Studio
 * Move dados de um mês específico para planilha de backup
 *
 * @param {Object} payload - Parâmetros do arquivamento
 * @param {number} payload.mes - Mês a arquivar (1-12)
 * @param {number} payload.ano - Ano a arquivar
 * @returns {Object} Quantidade de registros arquivados
 *
 * @requires Permissão ADMIN
 * @requires Pasta de backup configurada em Config_Geral
 *
 * @description
 * O processo de arquivamento:
 * 1. Identifica solicitações do período especificado
 * 2. Cria/atualiza planilha "Backup_Looker_Studio" na pasta de backup
 * 3. Consolida dados de Solicitações + Erros + Respostas
 * 4. Adiciona campos calculados (ano, mês, semana, tempo de resposta)
 * 5. Remove dados originais das planilhas do app
 *
 * @sideeffect Move dados para planilha de backup
 * @sideeffect Remove dados das planilhas originais
 * @sideeffect Registra ação na auditoria
 */
function executarArquivamentoMensal(payload) {
  const debugId = Utilities.getUuid();
  const mes = Number(payload.mes);
  const ano = Number(payload.ano);

  try {
    // Verificar permissão (apenas ADMIN)
    const userEmail = Session.getActiveUser().getEmail() || "";
    const usuario = getUsuarioContexto_(userEmail);
    if (!usuario || usuario.perfil !== "ADMIN") {
      return { ok: false, errors: ["Apenas ADMIN pode executar arquivamento."], debugId: debugId };
    }

    // Validar mês e ano
    if (!mes || mes < 1 || mes > 12 || !ano || ano < 2020 || ano > 2030) {
      return { ok: false, errors: ["Mês ou ano inválido."], debugId: debugId };
    }

    // Obter pasta de backup
    const pastaBackupId = String(getConfigGeral_("pasta_backup_id") || "").trim();
    if (!pastaBackupId) {
      return { ok: false, errors: ["Configure a pasta de backup primeiro."], debugId: debugId };
    }

    // Verificar pasta
    let folder;
    try {
      folder = DriveApp.getFolderById(pastaBackupId);
      folder.getName();
    } catch (e) {
      return { ok: false, errors: ["Pasta de backup não encontrada. Verifique o ID."], debugId: debugId };
    }

    // Definir período
    const dataInicio = new Date(ano, mes - 1, 1, 0, 0, 0);
    const dataFim = new Date(ano, mes, 0, 23, 59, 59);

    // Obter dados uma única vez
    const ss = getSpreadsheet_();
    const solicitacoesSheet = ss.getSheetByName(CONFIG.SHEETS.SOLICITACOES);
    const errosSheet = ss.getSheetByName(CONFIG.SHEETS.ERROS);
    const respostasSheet = ss.getSheetByName(CONFIG.SHEETS.RESPOSTAS);
    const limiaresSheet = ss.getSheetByName(CONFIG.SHEETS.LIMIARES);

    // Salvar e remover filtros das planilhas antes de ler os dados
    const sheetsComFiltro = [];
    [solicitacoesSheet, errosSheet, respostasSheet].forEach(sheet => {
      const filter = sheet.getFilter();
      if (filter) {
        sheetsComFiltro.push(sheet);
        filter.remove();
      }
    });

    const solicitacoes = solicitacoesSheet.getDataRange().getValues();
    const erros = errosSheet.getDataRange().getValues();
    const respostas = respostasSheet.getDataRange().getValues();
    const limiaresData = limiaresSheet.getDataRange().getValues();

    const horasAlerta = limiaresData.length > 1 ? Number(limiaresData[1][0]) || 1 : 1;
    const horasCritico = limiaresData.length > 1 ? Number(limiaresData[1][1]) || 2 : 2;

    // Identificar IDs do período (usar Set para busca O(1))
    const idsDoMesSet = new Set();
    for (let i = 1; i < solicitacoes.length; i++) {
      const d = solicitacoes[i][3];
      if (d instanceof Date && d >= dataInicio && d <= dataFim) {
        idsDoMesSet.add(solicitacoes[i][0]);
      }
    }

    if (idsDoMesSet.size === 0) {
      const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
      return { ok: false, errors: ["Nenhuma solicitação encontrada em " + meses[mes-1] + "/" + ano + "."], debugId: debugId };
    }

    // Validar se todas as solicitações do período estão respondidas
    const solicitacoesAbertas = [];
    for (let i = 1; i < solicitacoes.length; i++) {
      const id = solicitacoes[i][0];
      const status = solicitacoes[i][4];
      const requisicao = solicitacoes[i][1];
      if (idsDoMesSet.has(id) && status === "ABERTO") {
        solicitacoesAbertas.push(requisicao);
      }
    }

    if (solicitacoesAbertas.length > 0) {
      const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
      const lista = solicitacoesAbertas.length <= 10
        ? solicitacoesAbertas.join(", ")
        : solicitacoesAbertas.slice(0, 10).join(", ") + " e mais " + (solicitacoesAbertas.length - 10) + " outras";
      return {
        ok: false,
        errors: [
          "Não é possível arquivar " + meses[mes-1] + "/" + ano + " pois existem " + solicitacoesAbertas.length + " solicitação(ões) em ABERTO.",
          "Requisições pendentes: " + lista,
          "Por favor, responda todas as solicitações do período antes de arquivar."
        ],
        debugId: debugId
      };
    }

    // Criar mapas otimizados
    const solicitacoesMap = {};
    for (let i = 1; i < solicitacoes.length; i++) {
      solicitacoesMap[solicitacoes[i][0]] = solicitacoes[i];
    }

    const respostasMap = {};
    for (let i = 1; i < respostas.length; i++) {
      respostasMap[respostas[i][1] + "_" + respostas[i][2]] = respostas[i];
    }

    // === PREPARAR BACKUP ===
    const backupFileName = "Backup_Looker_Studio";
    let backupSS = obterOuCriarPlanilhaBackup_(folder, backupFileName);
    let dadosSheet = backupSS.getSheetByName("Dados");
    if (!dadosSheet) {
      dadosSheet = backupSS.insertSheet("Dados");
    }

    // Header do backup
    const HEADER = [
      "ID_Solicitacao", "Sequencia_Erro", "ID_Resposta",
      "Requisicao", "Solicitante", "Data_Hora_Pedido", "Status_Solicitacao", "Criado_Por_Email", "Data_Criacao_Solicitacao",
      "Tipo_Erro", "Detalhamento", "Setor_Local", "Diferenca_Valor_Solicitada", "Data_Criacao_Erro",
      "Nome_Responsavel", "Email_Responsavel", "Erro_Corrigido", "Houve_Diferenca_Valor", "Valor_Diferenca_R$", "Observacoes_Resposta", "Data_Hora_Correcao", "Data_Criacao_Resposta",
      "Ano_Pedido", "Mes_Pedido", "Mes_Nome_Pedido", "Trimestre_Pedido", "Semana_Ano_Pedido", "Dia_Mes_Pedido", "Dia_Semana_Pedido", "Dia_Semana_Num_Pedido", "Hora_Pedido", "Periodo_Dia_Pedido",
      "Ano_Correcao", "Mes_Correcao", "Mes_Nome_Correcao", "Dia_Semana_Correcao", "Hora_Correcao",
      "Tempo_Resposta_Minutos", "Tempo_Resposta_Horas", "Tempo_Resposta_Dias", "Status_Final", "Status_SLA", "Dentro_SLA", "SLA_Alerta_Horas", "SLA_Critico_Horas",
      "Ano_Mes_Pedido", "Ano_Mes_Correcao", "Tem_Resposta", "Tem_Diferenca_Valor"
    ];

    const diasSemana = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
    const mesesNome = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

    // Processar dados para backup
    const dadosParaBackup = [];
    for (let i = 1; i < erros.length; i++) {
      const solicitacaoId = erros[i][0];
      if (!idsDoMesSet.has(solicitacaoId)) continue;

      const sol = solicitacoesMap[solicitacaoId];
      if (!sol) continue;

      const erroSeq = erros[i][1];
      const dataPedido = sol[3];
      const resposta = respostasMap[solicitacaoId + "_" + erroSeq];
      const dataCorrecao = resposta ? resposta[9] : null;

      // Campos de data - Pedido
      let anoPedido = "", mesPedido = "", mesNomePedido = "", trimestrePedido = "", semanaAnoPedido = "";
      let diaMesPedido = "", diaSemaPedido = "", diaSemanNumPedido = "", horaPedido = "", periodoDiaPedido = "";
      if (dataPedido instanceof Date) {
        anoPedido = dataPedido.getFullYear();
        mesPedido = dataPedido.getMonth() + 1;
        mesNomePedido = mesesNome[dataPedido.getMonth()];
        trimestrePedido = "T" + Math.ceil(mesPedido / 3);
        semanaAnoPedido = getWeekNumber_(dataPedido);
        diaMesPedido = dataPedido.getDate();
        diaSemaPedido = diasSemana[dataPedido.getDay()];
        diaSemanNumPedido = dataPedido.getDay();
        horaPedido = dataPedido.getHours();
        periodoDiaPedido = horaPedido < 6 ? "Madrugada" : horaPedido < 12 ? "Manhã" : horaPedido < 18 ? "Tarde" : "Noite";
      }

      // Campos de data - Correção
      let anoCorrecao = "", mesCorrecao = "", mesNomeCorrecao = "", diaSemanaCorrecao = "", horaCorrecao = "";
      if (dataCorrecao instanceof Date) {
        anoCorrecao = dataCorrecao.getFullYear();
        mesCorrecao = dataCorrecao.getMonth() + 1;
        mesNomeCorrecao = mesesNome[dataCorrecao.getMonth()];
        diaSemanaCorrecao = diasSemana[dataCorrecao.getDay()];
        horaCorrecao = dataCorrecao.getHours();
      }

      // Métricas de tempo (usar vírgula como separador decimal - padrão BR)
      let tempoMinutos = "", tempoHoras = "", tempoDias = "";
      if (dataCorrecao instanceof Date && dataPedido instanceof Date) {
        const diffMs = dataCorrecao - dataPedido;
        tempoMinutos = Math.round(diffMs / 60000);
        tempoHoras = (diffMs / 3600000).toFixed(2).replace(".", ",");
        tempoDias = (diffMs / 86400000).toFixed(2).replace(".", ",");
      }

      // Status
      const statusFinal = resposta ? (resposta[5] === "SIM" ? "Corrigido" : "Em Correção") : "Aberto";
      let statusSLA = "OK", dentroSLA = "SIM";
      if (tempoHoras !== "") {
        const h = parseFloat(tempoHoras);
        if (h >= horasCritico) { statusSLA = "CRITICO"; dentroSLA = "NÃO"; }
        else if (h >= horasAlerta) { statusSLA = "ALERTA"; dentroSLA = "NÃO"; }
      }

      // Campos auxiliares
      const anoMesPedido = anoPedido ? anoPedido + "-" + String(mesPedido).padStart(2, "0") : "";
      const anoMesCorrecao = anoCorrecao ? anoCorrecao + "-" + String(mesCorrecao).padStart(2, "0") : "";

      dadosParaBackup.push([
        solicitacaoId, erroSeq, resposta ? resposta[0] : "",
        sol[1], sol[2], dataPedido, sol[4], sol[5], sol[6],
        erros[i][2], erros[i][3], erros[i][4], erros[i][5], erros[i][6],
        resposta ? resposta[3] : "", resposta ? resposta[4] : "", resposta ? resposta[5] : "",
        resposta ? resposta[6] : "", resposta ? resposta[7] : "", resposta ? resposta[8] : "",
        dataCorrecao || "", resposta ? resposta[10] : "",
        anoPedido, mesPedido, mesNomePedido, trimestrePedido, semanaAnoPedido, diaMesPedido, diaSemaPedido, diaSemanNumPedido, horaPedido, periodoDiaPedido,
        anoCorrecao, mesCorrecao, mesNomeCorrecao, diaSemanaCorrecao, horaCorrecao,
        tempoMinutos, tempoHoras, tempoDias, statusFinal, statusSLA, dentroSLA, horasAlerta, horasCritico,
        anoMesPedido, anoMesCorrecao, resposta ? "SIM" : "NÃO", erros[i][5] === "SIM" ? "SIM" : "NÃO"
      ]);
    }

    // Gravar no backup (com cabeçalho se necessário)
    if (dadosParaBackup.length > 0) {
      const lastRow = dadosSheet.getLastRow();

      // Verificar se precisa adicionar cabeçalho (linha 1 vazia ou sem cabeçalho correto)
      let precisaHeader = false;
      if (lastRow === 0) {
        precisaHeader = true;
      } else {
        // Verificar se a primeira célula tem o cabeçalho correto
        const primeiraCell = dadosSheet.getRange(1, 1).getValue();
        if (primeiraCell !== "ID_Solicitacao") {
          precisaHeader = true;
        }
      }

      if (precisaHeader) {
        // Limpar planilha e adicionar cabeçalho formatado
        dadosSheet.clearContents();
        dadosSheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
        dadosSheet.getRange(1, 1, 1, HEADER.length)
          .setFontWeight("bold")
          .setBackground("#009688")
          .setFontColor("#FFFFFF")
          .setHorizontalAlignment("center");
        dadosSheet.setFrozenRows(1);
        // Adicionar dados na linha 2
        dadosSheet.getRange(2, 1, dadosParaBackup.length, HEADER.length).setValues(dadosParaBackup);
      } else {
        // Adicionar dados após a última linha existente
        dadosSheet.getRange(lastRow + 1, 1, dadosParaBackup.length, HEADER.length).setValues(dadosParaBackup);
      }

      // Ajustar largura das colunas automaticamente
      dadosSheet.autoResizeColumns(1, HEADER.length);
    }

    // === REMOÇÃO EM MASSA (filtrar e reescrever) ===
    // Filtrar solicitações que NÃO são do período (manter apenas essas)
    const solicitacoesHeader = solicitacoes[0];
    const solicitacoesManter = [solicitacoesHeader];
    for (let i = 1; i < solicitacoes.length; i++) {
      if (!idsDoMesSet.has(solicitacoes[i][0])) {
        solicitacoesManter.push(solicitacoes[i]);
      }
    }

    // Filtrar erros que NÃO são do período
    const errosHeader = erros[0];
    const errosManter = [errosHeader];
    for (let i = 1; i < erros.length; i++) {
      if (!idsDoMesSet.has(erros[i][0])) {
        errosManter.push(erros[i]);
      }
    }

    // Filtrar respostas que NÃO são do período
    const respostasHeader = respostas[0];
    const respostasManter = [respostasHeader];
    for (let i = 1; i < respostas.length; i++) {
      if (!idsDoMesSet.has(respostas[i][1])) {
        respostasManter.push(respostas[i]);
      }
    }

    // Reescrever planilhas (muito mais rápido que deletar linha a linha)
    // Limpar e reescrever Solicitações
    solicitacoesSheet.clearContents();
    if (solicitacoesManter.length > 0) {
      solicitacoesSheet.getRange(1, 1, solicitacoesManter.length, solicitacoesManter[0].length).setValues(solicitacoesManter);
    }

    // Limpar e reescrever Erros
    errosSheet.clearContents();
    if (errosManter.length > 0) {
      errosSheet.getRange(1, 1, errosManter.length, errosManter[0].length).setValues(errosManter);
    }

    // Limpar e reescrever Respostas
    respostasSheet.clearContents();
    if (respostasManter.length > 0) {
      respostasSheet.getRange(1, 1, respostasManter.length, respostasManter[0].length).setValues(respostasManter);
    }

    // Restaurar filtros nas planilhas que tinham
    sheetsComFiltro.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow > 0 && lastCol > 0) {
        sheet.getRange(1, 1, lastRow, lastCol).createFilter();
      }
    });

    const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
    auditLog_("ARCHIVE", "sistema", "arquivamento_mensal", meses[mes-1] + "/" + ano, null, idsDoMesSet.size + " registros");

    return {
      ok: true,
      debugId: debugId,
      arquivados: idsDoMesSet.size,
      errosArquivados: dadosParaBackup.length,
      backupFile: backupFileName
    };

  } catch (error) {
    // Tentar restaurar filtros mesmo em caso de erro
    try {
      if (typeof sheetsComFiltro !== 'undefined') {
        sheetsComFiltro.forEach(sheet => {
          const lastRow = sheet.getLastRow();
          const lastCol = sheet.getLastColumn();
          if (lastRow > 0 && lastCol > 0 && !sheet.getFilter()) {
            sheet.getRange(1, 1, lastRow, lastCol).createFilter();
          }
        });
      }
    } catch (e) { /* ignora erro na restauração */ }

    safeLogDebug_("executarArquivamentoMensal", "error", { debugId: debugId, error: String(error) });
    return { ok: false, errors: ["Erro no arquivamento: " + String(error)], debugId: debugId };
  }
}

// Obtém ou cria a planilha de backup para Looker Studio
function obterOuCriarPlanilhaBackup_(folder, fileName) {
  // Procurar planilha existente na pasta
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.open(file);
  }

  // Criar nova planilha
  const ss = SpreadsheetApp.create(fileName);
  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(folder);

  // Remover aba padrão "Sheet1"
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }

  return ss;
}

// Calcula o número da semana do ano (ISO 8601)
function getWeekNumber_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/**
 * Executa limpeza de dados antigos do sistema
 * Remove registros mais antigos que o período especificado
 *
 * @param {Object} payload - Parâmetros da limpeza
 * @param {number} [payload.meses=6] - Quantidade de meses a manter
 * @returns {Object} Quantidade de registros removidos e data limite
 *
 * @requires Permissão ADMIN
 *
 * @description
 * O processo de limpeza:
 * 1. Calcula data limite (hoje - X meses)
 * 2. Identifica solicitações anteriores à data limite
 * 3. Remove respostas relacionadas
 * 4. Remove erros relacionados
 * 5. Remove solicitações
 *
 * @warning Esta operação é IRREVERSÍVEL!
 * @sideeffect Remove dados permanentemente
 * @sideeffect Registra ação na auditoria
 */
function executarLimpeza(payload) {
  const debugId = Utilities.getUuid();
  const meses = Number(payload.meses) || 6;
  safeLogDebug_("executarLimpeza", "start", { debugId: debugId, meses: meses });

  try {
    // Verificar permissão (apenas ADMIN)
    const userEmail = Session.getActiveUser().getEmail() || "";
    const usuario = getUsuarioContexto_(userEmail);
    if (!usuario || usuario.perfil !== "ADMIN") {
      return { ok: false, errors: ["Apenas ADMIN pode executar limpeza."], debugId: debugId };
    }

    // Calcular data limite
    const dataLimite = new Date();
    dataLimite.setMonth(dataLimite.getMonth() - meses);
    dataLimite.setHours(0, 0, 0, 0);

    const dataLimiteStr = Utilities.formatDate(dataLimite, "America/Sao_Paulo", "dd/MM/yyyy");

    safeLogDebug_("executarLimpeza", "dataLimite", { dataLimite: dataLimiteStr });

    // Obter planilhas
    const solicitacoesSheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);
    const errosSheet = getSheet_(CONFIG.SHEETS.ERROS);
    const respostasSheet = getSheet_(CONFIG.SHEETS.RESPOSTAS);

    // Encontrar solicitações antigas
    const solicitacoes = solicitacoesSheet.getDataRange().getValues();
    const solicitacoesARemover = [];

    for (let i = 1; i < solicitacoes.length; i++) {
      const dataHoraPedido = solicitacoes[i][3];
      if (dataHoraPedido && dataHoraPedido instanceof Date) {
        if (dataHoraPedido < dataLimite) {
          solicitacoesARemover.push({
            row: i + 1,
            id: solicitacoes[i][0]
          });
        }
      }
    }

    safeLogDebug_("executarLimpeza", "encontradas", { total: solicitacoesARemover.length });

    if (solicitacoesARemover.length === 0) {
      return { ok: true, debugId: debugId, removidos: 0, dataLimite: dataLimiteStr };
    }

    // Coletar IDs para remover erros e respostas relacionados
    const idsARemover = solicitacoesARemover.map(s => s.id);

    // Remover respostas relacionadas (coluna 1 = id_solicitacao)
    const respostas = respostasSheet.getDataRange().getValues();
    const respostasRowsARemover = [];
    for (let i = 1; i < respostas.length; i++) {
      if (idsARemover.indexOf(respostas[i][1]) !== -1) {
        respostasRowsARemover.push(i + 1);
      }
    }
    // Remover de baixo para cima para não afetar os índices
    respostasRowsARemover.sort((a, b) => b - a).forEach(row => {
      respostasSheet.deleteRow(row);
    });

    // Remover erros relacionados
    const erros = errosSheet.getDataRange().getValues();
    const errosRowsARemover = [];
    for (let i = 1; i < erros.length; i++) {
      if (idsARemover.indexOf(erros[i][0]) !== -1) {
        errosRowsARemover.push(i + 1);
      }
    }
    errosRowsARemover.sort((a, b) => b - a).forEach(row => {
      errosSheet.deleteRow(row);
    });

    // Remover solicitações (de baixo para cima)
    solicitacoesARemover.sort((a, b) => b.row - a.row).forEach(s => {
      solicitacoesSheet.deleteRow(s.row);
    });

    const totalRemovidos = solicitacoesARemover.length;

    safeLogDebug_("executarLimpeza", "success", {
      debugId: debugId,
      removidos: totalRemovidos,
      dataLimite: dataLimiteStr
    });

    auditLog_("CLEANUP", "sistema", "limpeza_dados", "meses", meses, totalRemovidos + " registros");

    return { ok: true, debugId: debugId, removidos: totalRemovidos, dataLimite: dataLimiteStr };

  } catch (error) {
    safeLogDebug_("executarLimpeza", "error", { debugId: debugId, error: String(error) });
    return { ok: false, errors: [String(error)], debugId: debugId };
  }
}

// ============================================================================
// FUNÇÕES AUXILIARES INTERNAS
// ============================================================================

/**
 * Retorna erros abertos filtrados por permissões do usuário
 * Função interna otimizada para grandes volumes de dados
 *
 * @param {Object} usuario - Objeto do usuário com perfil e setores
 * @param {number} [limit=3000] - Limite máximo de registros
 * @returns {Array} Lista de erros abertos
 *
 * @private
 * @description
 * Otimizações implementadas:
 * - Criação de lookup maps O(1) para solicitações e respostas
 * - Filtragem em duas passagens (índices primeiro, objetos depois)
 * - Cálculo de SLA em batch usando thresholds pré-carregados
 * - Ordenação por timestamp para priorização de SLA
 */
function getOpenErrosForUsuario_(usuario, limit) {
  const startTime = new Date().getTime();
  const MAX_LIMIT = limit || 3000;

  // Log user info for debugging
  safeLogDebug_("getOpenErrosForUsuario_", "usuario", {
    email: usuario ? usuario.email : "null",
    perfil: usuario ? usuario.perfil : "null",
    setores: usuario ? usuario.setores : "null",
    limit: MAX_LIMIT
  });

  const errosSheet = getSheet_(CONFIG.SHEETS.ERROS);
  const respostasSheet = getSheet_(CONFIG.SHEETS.RESPOSTAS);
  const solicitacoesSheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);

  const erros = errosSheet.getDataRange().getValues();
  const respostas = respostasSheet.getDataRange().getValues();
  const solicitacoesValues = solicitacoesSheet.getDataRange().getValues();

  // OPTIMIZATION: Read thresholds ONCE (not for each record)
  const thresholds = getThresholds_();
  const now = new Date();

  // Create lookup maps
  const solicitacoesMap = {};
  for (let i = 1; i < solicitacoesValues.length; i++) {
    solicitacoesMap[solicitacoesValues[i][0]] = solicitacoesValues[i];
  }

  const respostasMap = {};
  for (let i = 1; i < respostas.length; i++) {
    const key = respostas[i][1] + "_" + respostas[i][2];
    respostasMap[key] = respostas[i];
  }

  const mapsTime = new Date().getTime();
  safeLogDebug_("getOpenErrosForUsuario_", "maps_created", {
    elapsed: mapsTime - startTime,
    errosCount: erros.length
  });

  // OPTIMIZATION: First pass - collect only minimal data for sorting
  // This avoids creating full objects for all 4000 records
  const minimalList = [];
  let filteredBySetor = 0;

  for (let i = 1; i < erros.length; i++) {
    const setorLocal = erros[i][4];
    if (!usuarioPodeVerSetor_(usuario, setorLocal)) {
      filteredBySetor++;
      continue;
    }
    const solicitacaoId = erros[i][0];
    const solicitacaoData = solicitacoesMap[solicitacaoId];
    if (!solicitacaoData) continue;

    // Store only index and date timestamp for sorting
    const dataHoraPedido = solicitacaoData[3];
    const timestamp = dataHoraPedido instanceof Date ? dataHoraPedido.getTime() : new Date(dataHoraPedido).getTime();
    minimalList.push({ idx: i, ts: timestamp });
  }

  const filterTime = new Date().getTime();
  safeLogDebug_("getOpenErrosForUsuario_", "filtered", {
    elapsed: filterTime - mapsTime,
    total: erros.length - 1,
    filtered: filteredBySetor,
    passed: minimalList.length
  });

  // Sort by timestamp (oldest first for SLA priority)
  minimalList.sort((a, b) => a.ts - b.ts);

  // Take only the records we need
  const limitedList = minimalList.slice(0, MAX_LIMIT);

  const sortTime = new Date().getTime();
  safeLogDebug_("getOpenErrosForUsuario_", "sorted", {
    elapsed: sortTime - filterTime,
    limited: limitedList.length
  });

  // OPTIMIZATION: Second pass - build full objects only for limited records
  const result = [];
  for (let j = 0; j < limitedList.length; j++) {
    const i = limitedList[j].idx;
    const solicitacaoId = erros[i][0];
    const erroSeq = erros[i][1];
    const setorLocal = erros[i][4];
    const solicitacaoData = solicitacoesMap[solicitacaoId];
    const latestResposta = respostasMap[solicitacaoId + "_" + erroSeq] || null;

    const dataHoraPedido = solicitacaoData[3];
    const dataHoraPedidoStr = dataHoraPedido instanceof Date ? dataHoraPedido.toISOString() : String(dataHoraPedido);

    let responsavel = "";
    let dataHoraCorrecaoStr = "";
    let correcaoFinalizada = false;
    let houveDiferencaValorResposta = "";
    let diferencaValorResposta = "";
    let observacoesResposta = "";

    if (latestResposta) {
      responsavel = latestResposta[3] || "";
      correcaoFinalizada = latestResposta[5] === "SIM";
      houveDiferencaValorResposta = latestResposta[6] || "";
      diferencaValorResposta = latestResposta[7] || "";
      observacoesResposta = latestResposta[8] || "";
      if (latestResposta[9]) {
        const dataHoraCorrecao = latestResposta[9];
        dataHoraCorrecaoStr = dataHoraCorrecao instanceof Date ? dataHoraCorrecao.toISOString() : String(dataHoraCorrecao);
      }
    }

    // Calculate SLA inline (avoid calling getSlaStatus_ which reads sheet each time)
    const diffMs = now - new Date(dataHoraPedido);
    const diffMinutes = diffMs / 60000;
    let slaStatus = "OK";
    if (diffMinutes >= thresholds.criticalMinutes) slaStatus = "CRITICO";
    else if (diffMinutes >= thresholds.warnMinutes) slaStatus = "ALERTA";

    result.push({
      rowIndex: i,
      solicitacaoId: solicitacaoId,
      erroSeq: erroSeq,
      requisicao: solicitacaoData[1],
      solicitante: solicitacaoData[2],
      erro: erros[i][2],
      detalhamento: erros[i][3],
      setorLocal: setorLocal,
      diferencaNoValor: erros[i][5],
      confirmacaoMedica: erros[i][7] || "NAO",
      dataHoraPedido: dataHoraPedidoStr,
      slaStatus: slaStatus,
      responsavel: responsavel,
      dataHoraCorrecao: dataHoraCorrecaoStr,
      correcaoFinalizada: correcaoFinalizada,
      houveDiferencaValorResposta: houveDiferencaValorResposta,
      diferencaValorResposta: diferencaValorResposta,
      observacoesResposta: observacoesResposta
    });
  }

  const endTime = new Date().getTime();
  safeLogDebug_("getOpenErrosForUsuario_", "complete", {
    totalElapsed: endTime - startTime,
    returned: result.length
  });

  return result;
}

function getSolicitacoesList_() {
  const solicitacoesSheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);
  const values = solicitacoesSheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    const solicitacaoId = values[i][0];
    const status = values[i][4];
    const dataHoraPedido = values[i][3];
    list.push({
      solicitacaoId: solicitacaoId,
      requisicao: values[i][1],
      solicitante: values[i][2],
      status: status,
      dataHoraPedido: dataHoraPedido,
      slaStatus: getSlaStatus_(dataHoraPedido, null)
    });
  }
  const ordered = list.sort((a, b) => new Date(a.dataHoraPedido) - new Date(b.dataHoraPedido));
  logDebug_("getSolicitacoesList_", "count", { total: ordered.length });
  return ordered;
}

function getDashboardData_(records) {
  const solicitacoesSheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);
  const solicitacoes = solicitacoesSheet.getDataRange().getValues();
  const erros = records || [];

  let totalSolicitacoes = 0;
  let abertas = 0;
  let emCorrecao = 0;
  let corrigidas = 0;
  let alerta = 0;
  let critico = 0;

  for (let i = 1; i < solicitacoes.length; i++) {
    totalSolicitacoes++;
    const status = solicitacoes[i][4];
    if (status === "ABERTO") abertas++;
    if (status === "EM_CORRECAO") emCorrecao++;
    if (status === "CORRIGIDO") corrigidas++;
  }

  for (let i = 0; i < erros.length; i++) {
    const status = getSlaStatus_(erros[i].dataHoraPedido, erros[i].dataHoraCorrecao || null);
    if (status === "ALERTA") alerta++;
    if (status === "CRITICO") critico++;
  }

  return {
    totalSolicitacoes: totalSolicitacoes,
    abertas: abertas,
    emCorrecao: emCorrecao,
    corrigidas: corrigidas,
    alerta: alerta,
    critico: critico
  };
}

function usuarioPodeVerSetor_(usuario, setorLocal) {
  if (!usuario) return false;

  const perfil = (usuario.perfil || "").toUpperCase();
  if (perfil === "ADMIN" || perfil === "CONFERENTE") return true;

  // Registros sem setor: apenas usuários com acesso total (*) podem ver.
  // Evita que usuários com setor restrito vejam registros sem setor como bypass.
  if (!setorLocal || setorLocal === "") {
    return !!(usuario.setores && usuario.setores.indexOf("*") !== -1);
  }

  if (usuario.setores && usuario.setores.indexOf("*") !== -1) return true;
  if (usuario.setores && usuario.setores.indexOf(setorLocal) !== -1) return true;

  return false;
}

function getResponsaveis_() {
  const sheet = getSheet_(CONFIG.SHEETS.USUARIOS);
  const values = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    const nome = values[i][1];
    const perfil = values[i][2];
    if (!nome) continue;
    if (perfil === "ADMIN" || perfil === "RESPOSTA") {
      list.push(nome);
    }
  }
  return list;
}

function getUsuarios_() {
  const sheet = getSheet_(CONFIG.SHEETS.USUARIOS);
  const values = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    list.push({
      email: values[i][0],
      nome: values[i][1],
      perfil: values[i][2],
      setores: values[i][3]
    });
  }
  return list;
}

function getDashboardRecords_(limit) {
  try {
    const startTime = new Date().getTime();
    const MAX_LIMIT = limit || 3000;
    safeLogDebug_("getDashboardRecords_", "start", { limit: MAX_LIMIT });

    const errosSheet = getSheet_(CONFIG.SHEETS.ERROS);
    const respostasSheet = getSheet_(CONFIG.SHEETS.RESPOSTAS);
    const solicitacoesSheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);

    const erros = errosSheet.getDataRange().getValues();
    const respostas = respostasSheet.getDataRange().getValues();
    const solicitacoesValues = solicitacoesSheet.getDataRange().getValues();

    // OPTIMIZATION: Read thresholds ONCE (not for each record)
    const thresholds = getThresholds_();
    const now = new Date();

    // Create lookup maps
    const solicitacoesMap = {};
    for (let i = 1; i < solicitacoesValues.length; i++) {
      solicitacoesMap[solicitacoesValues[i][0]] = solicitacoesValues[i];
    }

    const respostasMap = {};
    for (let i = 1; i < respostas.length; i++) {
      const key = respostas[i][1] + "_" + respostas[i][2];
      respostasMap[key] = respostas[i];
    }

    const mapsTime = new Date().getTime();
    safeLogDebug_("getDashboardRecords_", "maps_created", {
      elapsed: mapsTime - startTime,
      errosCount: erros.length
    });

    // OPTIMIZATION: First pass - collect only minimal data for sorting
    const minimalList = [];
    for (let i = 1; i < erros.length; i++) {
      const solicitacaoId = erros[i][0];
      const solicitacaoData = solicitacoesMap[solicitacaoId];
      if (!solicitacaoData) continue;

      const dataHoraPedido = solicitacaoData[3];
      const timestamp = dataHoraPedido instanceof Date ? dataHoraPedido.getTime() : new Date(dataHoraPedido).getTime();
      minimalList.push({ idx: i, ts: timestamp });
    }

    const filterTime = new Date().getTime();
    safeLogDebug_("getDashboardRecords_", "filtered", {
      elapsed: filterTime - mapsTime,
      count: minimalList.length
    });

    // Sort by timestamp (oldest first for SLA priority)
    minimalList.sort((a, b) => a.ts - b.ts);

    // Take only the records we need
    const limitedList = minimalList.slice(0, MAX_LIMIT);

    const sortTime = new Date().getTime();
    safeLogDebug_("getDashboardRecords_", "sorted", {
      elapsed: sortTime - filterTime,
      limited: limitedList.length
    });

    // OPTIMIZATION: Second pass - build full objects only for limited records
    const result = [];
    for (let j = 0; j < limitedList.length; j++) {
      const i = limitedList[j].idx;
      const solicitacaoId = erros[i][0];
      const erroSeq = erros[i][1];
      const solicitacaoData = solicitacoesMap[solicitacaoId];
      const resposta = respostasMap[solicitacaoId + "_" + erroSeq] || null;

      const dataHoraPedido = solicitacaoData[3];
      const dataHoraPedidoStr = dataHoraPedido instanceof Date ? dataHoraPedido.toISOString() : String(dataHoraPedido);

      let dataHoraCorrecaoStr = "";
      let correcaoDate = null;
      if (resposta && resposta[9]) {
        const dataHoraCorrecao = resposta[9];
        correcaoDate = dataHoraCorrecao instanceof Date ? dataHoraCorrecao : new Date(dataHoraCorrecao);
        dataHoraCorrecaoStr = correcaoDate.toISOString();
      }

      // Calculate SLA inline (avoid calling getSlaStatus_ which reads sheet each time)
      const base = correcaoDate || now;
      const diffMs = base - new Date(dataHoraPedido);
      const diffMinutes = diffMs / 60000;
      let slaStatus = "OK";
      if (diffMinutes >= thresholds.criticalMinutes) slaStatus = "CRITICO";
      else if (diffMinutes >= thresholds.warnMinutes) slaStatus = "ALERTA";

      result.push({
        solicitacaoId: solicitacaoId,
        erroSeq: erroSeq,
        requisicao: solicitacaoData[1],
        solicitante: solicitacaoData[2],
        dataHoraPedido: dataHoraPedidoStr,
        statusSolicitacao: solicitacaoData[4],
        erro: erros[i][2],
        detalhamento: erros[i][3],
        setorLocal: erros[i][4],
        diferencaNoValor: erros[i][5],
        confirmacaoMedica: erros[i][7] || "NAO",
        responsavel: resposta ? resposta[3] : "",
        houveDiferencaValorResposta: resposta ? resposta[6] : "",
        diferencaValorResposta: resposta ? resposta[7] : "",
        observacoes: resposta ? resposta[8] : "",
        dataHoraCorrecao: dataHoraCorrecaoStr,
        slaStatus: slaStatus
      });
    }

    const endTime = new Date().getTime();
    safeLogDebug_("getDashboardRecords_", "complete", {
      totalElapsed: endTime - startTime,
      returned: result.length
    });

    return result;

  } catch (error) {
    safeLogDebug_("getDashboardRecords_", "error", { error: String(error) });
    throw error;
  }
}
