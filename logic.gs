// Cache keys
const CACHE_KEY_CONFIG = "app_config_data";
const CACHE_KEY_LISTAS = "app_listas_data";
const CACHE_DURATION = 600; // 10 minutes in seconds

function getConfigData_() {
  // Try to get from cache first
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY_CONFIG);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Cache corrupted, continue to load from sheet
    }
  }

  const solicitantesSheet = getSheet_(CONFIG.SHEETS.SOLICITANTES);
  const setoresSheet = getSheet_(CONFIG.SHEETS.SETORES);
  const errosSheet = getSheet_(CONFIG.SHEETS.ERROS_CADASTRO);

  const solicitantesValues = solicitantesSheet.getDataRange().getValues();
  const setoresValues = setoresSheet.getDataRange().getValues();
  const errosValues = errosSheet.getDataRange().getValues();

  const pharmaceuticas = [];
  const erros = [];
  const setores = [];

  for (let i = 1; i < solicitantesValues.length; i++) {
    if (solicitantesValues[i][0]) pharmaceuticas.push(solicitantesValues[i][0]);
  }
  for (let i = 1; i < setoresValues.length; i++) {
    if (setoresValues[i][0]) setores.push(setoresValues[i][0]);
  }
  for (let i = 1; i < errosValues.length; i++) {
    if (errosValues[i][0]) {
      erros.push({
        nome: errosValues[i][0],
        classificacao: errosValues[i][1] || ''
      });
    }
  }

  const thresholds = getThresholds_();
  const result = {
    pharmaceuticas: pharmaceuticas,
    erros: erros,
    setores: setores,
    thresholds: thresholds
  };

  // Store in cache
  try {
    cache.put(CACHE_KEY_CONFIG, JSON.stringify(result), CACHE_DURATION);
  } catch (e) {
    // Cache write failed, continue anyway
  }

  return result;
}

// Invalidate config cache (call this when config changes)
function invalidateConfigCache_() {
  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_KEY_CONFIG);
  cache.remove(CACHE_KEY_LISTAS);
}

function getListasSheet_() {
  const ss = getSpreadsheet_();
  const candidates = [
    CONFIG.SHEETS.LISTAS,
    "Colaboradores",
    "colaboradores",
    "Lista",
    "LISTAS",
    "listas"
  ];
  for (let i = 0; i < candidates.length; i++) {
    const sheet = ss.getSheetByName(candidates[i]);
    if (sheet) return sheet;
  }
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const name = String(sheets[i].getName() || "").trim().toLowerCase();
    if (name === "listas" || name === "colaboradores") return sheets[i];
  }
  return null;
}

function getListasData_() {
  // Try to get from cache first
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY_LISTAS);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Cache corrupted, continue to load from sheet
    }
  }

  const listasSheet = getListasSheet_();
  if (!listasSheet) {
    return { colaboradores: [], setores: [] };
  }
  const values = listasSheet.getDataRange().getValues();
  let colColaborador = 0;
  let colSetor = 1;
  if (values.length) {
    const header = values[0].map((v) => String(v || "").trim().toLowerCase());
    const idxColab = header.indexOf("colaboradores");
    const idxColabAlt = header.indexOf("colaborador");
    const idxSetor = header.indexOf("setores");
    if (idxColab >= 0) colColaborador = idxColab;
    else if (idxColabAlt >= 0) colColaborador = idxColabAlt;
    if (idxSetor >= 0) colSetor = idxSetor;
  }
  const colaboradores = new Set();
  const setores = new Set();

  for (let i = 1; i < values.length; i++) {
    if (values[i][colColaborador]) colaboradores.add(values[i][colColaborador]);
    if (values[i][colSetor]) setores.add(values[i][colSetor]);
  }

  // If no setores found in Listas sheet, try Setores_Local sheet
  if (setores.size === 0) {
    try {
      const setoresSheet = getSheet_(CONFIG.SHEETS.SETORES);
      const setoresValues = setoresSheet.getDataRange().getValues();
      for (let i = 1; i < setoresValues.length; i++) {
        if (setoresValues[i][0]) setores.add(setoresValues[i][0]);
      }
    } catch (error) {
      // Ignore error
    }
  }

  const result = {
    colaboradores: Array.from(colaboradores),
    setores: Array.from(setores)
  };

  // Store in cache
  try {
    cache.put(CACHE_KEY_LISTAS, JSON.stringify(result), CACHE_DURATION);
  } catch (e) {
    // Cache write failed, continue anyway
  }
  return result;
}

function getThresholds_() {
  const sheet = getSheet_(CONFIG.SHEETS.LIMIARES);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return {
      warnMinutes: CONFIG.DEFAULTS.WARN_MINUTES,
      criticalMinutes: CONFIG.DEFAULTS.CRITICAL_MINUTES,
      timezone: CONFIG.DEFAULTS.TIMEZONE
    };
  }
  return {
    warnMinutes: Number(values[1][0]) || CONFIG.DEFAULTS.WARN_MINUTES,
    criticalMinutes: Number(values[1][1]) || CONFIG.DEFAULTS.CRITICAL_MINUTES,
    timezone: values[1][2] || CONFIG.DEFAULTS.TIMEZONE
  };
}

// Limites de tamanho de campo (caracteres)
const MAX_REQUISICAO_LEN = 30;
const MAX_DETALHAMENTO_LEN = 2000;
const MAX_OBSERVACOES_LEN = 1000;
const MAX_DIFERENCA_VALOR_LEN = 50;

function validateRequestPayload_(payload, config) {
  const errors = [];
  if (!payload.requisicao) {
    errors.push("Informe a requisição.");
  } else if (String(payload.requisicao).length > MAX_REQUISICAO_LEN) {
    errors.push("Requisição não pode exceder " + MAX_REQUISICAO_LEN + " caracteres.");
  }
  if (!payload.solicitante || config.pharmaceuticas.indexOf(payload.solicitante) === -1) {
    errors.push("Selecione uma farmacêutica válida.");
  }
  // Extract error names from objects or strings
  const errorNames = config.erros.map(function(e) { return typeof e === 'object' ? e.nome : e; });
  if (!payload.erro || errorNames.indexOf(payload.erro) === -1) {
    errors.push("Selecione um tipo de erro válido.");
  }
  if (!payload.detalhamento) {
    errors.push("Detalhamento do erro é obrigatório.");
  } else if (String(payload.detalhamento).length > MAX_DETALHAMENTO_LEN) {
    errors.push("Detalhamento não pode exceder " + MAX_DETALHAMENTO_LEN + " caracteres.");
  }
  if (!payload.setorLocal || config.setores.indexOf(payload.setorLocal) === -1) {
    errors.push("Selecione um setor/local válido.");
  }
  if (payload.diferencaNoValor !== "SIM" && payload.diferencaNoValor !== "NAO") {
    errors.push("Informe se houve diferença no valor.");
  }
  if (payload.confirmacaoMedica !== "SIM" && payload.confirmacaoMedica !== "NAO") {
    errors.push("Informe se é necessário confirmação médica.");
  }
  if (!payload.dataHoraPedido) {
    errors.push("Data/hora do pedido não foi registrada.");
  }
  if (errors.length) {
    logDebug_("validateRequestPayload_", "errors", errors);
  }
  return errors;
}

function validateResponsePayload_(payload, config) {
  const errors = [];
  if (!payload.solicitacaoId) {
    errors.push("Informe a solicitação.");
  }
  if (!payload.erroSeq) {
    errors.push("Informe o número do erro.");
  }
  if (!payload.nomeResponsavel) {
    errors.push("Informe o responsável pelo erro.");
  }
  if (payload.houveDiferencaValor !== "SIM" && payload.houveDiferencaValor !== "NAO") {
    errors.push("Informe se houve diferença no valor.");
  }
  if (payload.houveDiferencaValor === "SIM") {
    if (payload.diferencaValorResposta === "" || payload.diferencaValorResposta === null || payload.diferencaValorResposta === undefined) {
      errors.push("Informe a diferença do valor.");
    } else if (isNaN(Number(payload.diferencaValorResposta))) {
      errors.push("A diferença do valor precisa ser numérica.");
    }
  }
  if (!payload.dataHoraCorrecao) {
    errors.push("Informe a data/hora da correção.");
  }
  if (payload.observacoes && String(payload.observacoes).length > MAX_OBSERVACOES_LEN) {
    errors.push("Observações não podem exceder " + MAX_OBSERVACOES_LEN + " caracteres.");
  }
  if (payload.diferencaValorResposta && String(payload.diferencaValorResposta).length > MAX_DIFERENCA_VALOR_LEN) {
    errors.push("Diferença de valor não pode exceder " + MAX_DIFERENCA_VALOR_LEN + " caracteres.");
  }
  if (errors.length) {
    logDebug_("validateResponsePayload_", "errors", errors);
  }
  return errors;
}

function getSlaStatus_(requestDatetime, correcaoDatetime) {
  const thresholds = getThresholds_();
  const now = new Date();
  const base = correcaoDatetime ? new Date(correcaoDatetime) : now;
  const diffMs = base - new Date(requestDatetime);
  const diffMinutes = diffMs / 60000;
  if (diffMinutes >= thresholds.criticalMinutes) return "CRITICO";
  if (diffMinutes >= thresholds.warnMinutes) return "ALERTA";
  return "OK";
}

// Permissões padrão por perfil — espelhado no frontend como PROFILE_PERM_DEFAULTS_UI_
const PROFILE_PERM_DEFAULTS_ = {
  ADMIN: {
    tabs: ['resposta', 'solicitacao', 'dashboard', 'config', 'auditoria', 'ajuda'],
    actions: ['submitRequest', 'respond', 'editSolicitation', 'delete', 'manageUsers', 'manageConfig', 'archive', 'viewAudit'],
    dashTabs: ['colaboradores', 'solicitantes', 'ranking', 'erros']
  },
  CONFERENTE: {
    tabs: ['resposta', 'solicitacao', 'dashboard', 'config', 'ajuda'],
    actions: ['submitRequest', 'respond', 'editSolicitation', 'delete'],
    dashTabs: ['colaboradores', 'solicitantes', 'ranking', 'erros']
  },
  RESPOSTA: {
    tabs: ['resposta', 'dashboard', 'ajuda'],
    actions: ['submitRequest', 'respond'],
    dashTabs: ['colaboradores', 'solicitantes', 'ranking', 'erros']
  },
  ESPECTADOR: {
    tabs: ['dashboard', 'ajuda'],
    actions: [],
    dashTabs: ['colaboradores', 'solicitantes', 'ranking', 'erros']
  }
};

// Gera chave de cache segura a partir do email
function usuarioCacheKey_(email) {
  return "uctx_" + email.replace(/[^a-zA-Z0-9]/g, "_").substring(0, 60);
}

function getUsuarioContexto_(email) {
  // Cache de 5 min: leitura da sheet de usuários é cara e o perfil muda raramente
  const cache = CacheService.getScriptCache();
  const cacheKey = usuarioCacheKey_(email);
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch(e) { /* cache miss ou parse error — continua */ }

  const sheet = getSheet_(CONFIG.SHEETS.USUARIOS);
  const values = sheet.getDataRange().getValues();
  let result = null;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === email) {
      const perfil = values[i][2] || 'ESPECTADOR';
      // Lê permissões customizadas da coluna [4]; fallback para padrão do perfil
      let permissoes = null;
      const rawPerms = values[i][4];
      if (rawPerms) {
        try { permissoes = JSON.parse(rawPerms); } catch(e) {}
      }
      if (!permissoes || !permissoes.tabs || !permissoes.actions) {
        permissoes = PROFILE_PERM_DEFAULTS_[perfil] || PROFILE_PERM_DEFAULTS_.ESPECTADOR;
      }
      // Compatibilidade: usuários sem dashTabs recebem padrão do perfil
      if (!permissoes.dashTabs) {
        permissoes.dashTabs = (PROFILE_PERM_DEFAULTS_[perfil] || PROFILE_PERM_DEFAULTS_.ESPECTADOR).dashTabs;
      }
      result = {
        email: values[i][0],
        nome: values[i][1],
        perfil: perfil,
        setores: parseSetores_(values[i][3]),
        permissoes: permissoes
      };
      break;
    }
  }
  if (!result) {
    logDebug_("getUsuarioContexto_", "usuario_nao_encontrado", {});
    result = {
      email: email,
      nome: "",
      perfil: "ESPECTADOR",
      setores: [],
      permissoes: PROFILE_PERM_DEFAULTS_.ESPECTADOR
    };
  }

  try { cache.put(cacheKey, JSON.stringify(result), 300); } catch(e) { /* ignore */ }
  return result;
}

/**
 * Guard de autorização centralizado — verifica se o usuário logado tem a permissão exigida.
 * Lança Error se não autorizado; retorna o objeto usuario se autorizado.
 * Use ao início de qualquer endpoint de escrita sensível.
 *
 * @param {string|string[]} acaoOuAcoes - Uma ação ou array de ações necessárias (todas devem estar presentes)
 * @param {Object} [options]
 * @param {boolean} [options.adminOnly=false] - Se true, exige perfil ADMIN independente das actions
 * @returns {Object} Objeto usuario autenticado
 * @throws {Error} Se o usuário não tiver permissão
 */
function requirePermissao_(acaoOuAcoes, options) {
  const email = getEmailValidado_();
  const usuario = getUsuarioContexto_(email);
  const opts = options || {};

  if (opts.adminOnly) {
    if (!usuario.perfil || usuario.perfil.toUpperCase() !== 'ADMIN') {
      throw new Error('Acesso negado: apenas administradores podem realizar esta operação.');
    }
    return usuario;
  }

  const acoes = Array.isArray(acaoOuAcoes) ? acaoOuAcoes : [acaoOuAcoes];
  const permActions = (usuario.permissoes && usuario.permissoes.actions) ? usuario.permissoes.actions : [];

  for (let i = 0; i < acoes.length; i++) {
    if (permActions.indexOf(acoes[i]) === -1) {
      throw new Error('Acesso negado: permissão "' + acoes[i] + '" necessária.');
    }
  }
  return usuario;
}

// Invalida o cache de um usuário específico após alteração de perfil/setores
function invalidateUsuarioCache_(email) {
  try {
    CacheService.getScriptCache().remove(usuarioCacheKey_(email));
  } catch(e) { /* ignore */ }
}

function parseSetores_(raw) {
  if (!raw) return [];
  if (raw === "*") return ["*"];
  return raw
    .split(/[,;|]/)
    .map((item) => item.trim())
    .filter((item) => item);
}
