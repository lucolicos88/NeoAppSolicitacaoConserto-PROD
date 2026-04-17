// Buffer de auditoria: entradas acumuladas durante a execução e gravadas em lote no final.
// Cada execução GAS começa com o buffer vazio (variável de módulo reiniciada a cada request).
let _auditBuffer_ = [];

/**
 * Enfileira uma entrada de auditoria no buffer em memória.
 * NÃO grava na planilha — chame flushAuditLogs_() ao final da operação.
 * PERF: elimina N appendRow individuais; flushAuditLogs_() faz UMA chamada setValues.
 */
function auditLog_(actionType, tableName, recordKey, fieldName, oldValue, newValue) {
  try {
    const userEmail = Session.getActiveUser().getEmail() || "unknown";
    _auditBuffer_.push([
      Utilities.getUuid(),
      userEmail,
      actionType,
      tableName,
      recordKey,
      fieldName || "",
      String(oldValue !== null && oldValue !== undefined ? oldValue : ""),
      String(newValue !== null && newValue !== undefined ? newValue : ""),
      getClientIp_(),
      new Date()
    ]);
  } catch(e) {
    Logger.log("auditLog_ buffer failed: " + e);
  }
}

/**
 * Grava todas as entradas de auditoria acumuladas em UMA única chamada setValues.
 * Deve ser chamada no final de cada função de API pública (submitRequest, submitResponse, etc.).
 * PERF: 1 API call em vez de N appendRow individuais.
 */
function flushAuditLogs_() {
  if (_auditBuffer_.length === 0) return;
  try {
    const sheet = getSheet_(CONFIG.SHEETS.AUDITORIA);
    if (_auditBuffer_.length === 1) {
      // PERF: appendRow = 1 API call vs getLastRow + setValues = 2 calls
      sheet.appendRow(_auditBuffer_[0]);
    } else {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, _auditBuffer_.length, 10).setValues(_auditBuffer_);
    }
  } catch (e) {
    Logger.log("flushAuditLogs_ failed: " + e);
  }
  _auditBuffer_ = [];
}

function getClientIp_() {
  // GAS Web Apps não expõem o IP real do cliente.
  // Session.getTemporaryActiveUserKey() retorna uma chave opaca por usuário/sessão —
  // útil para correlacionar ações do mesmo usuário, mas não é um IP.
  try {
    return Session.getTemporaryActiveUserKey();
  } catch (error) {
    return "unknown";
  }
}

/**
 * Mascara email para logs: joao.silva@empresa.com → jo***@empresa.com
 */
function maskEmail_(email) {
  if (!email || typeof email !== "string") return "***";
  const at = email.indexOf("@");
  if (at < 0) return "***";
  const local = email.substring(0, at);
  const domain = email.substring(at);
  const visible = local.length <= 2 ? local : local.substring(0, 2);
  return visible + "***" + domain;
}

/**
 * Remove valores de chaves sensíveis de um objeto antes de gravar em log.
 * Ofusca campos como email, nome, solicitante, requisicao.
 */
function sanitizePayloadForLog_(payload) {
  if (!payload || typeof payload !== "object") return payload;
  const SENSITIVE = ["email", "nome", "solicitante", "requisicao", "password", "senha"];
  const result = {};
  Object.keys(payload).forEach(function(key) {
    const lk = key.toLowerCase();
    const isSensitive = SENSITIVE.some(function(s) { return lk.indexOf(s) !== -1; });
    if (isSensitive) {
      result[key] = "***";
    } else if (typeof payload[key] === "object" && payload[key] !== null && !Array.isArray(payload[key])) {
      result[key] = sanitizePayloadForLog_(payload[key]);
    } else {
      result[key] = payload[key];
    }
  });
  return result;
}

/**
 * Registra mensagem de debug.
 * PERF: Por padrão escreve APENAS em Logger.log() (rápido, sem chamada Sheets).
 * Para gravar na planilha, ative CONFIG.DEBUG_SHEETS = true (uso pontual em depuração).
 */
function logDebug_(context, message, payload) {
  if (!CONFIG.DEBUG_ENABLED) return;
  const sanitized = sanitizePayloadForLog_(payload);
  // Logger.log() é em memória — zero API calls, sem custo de latência
  Logger.log("[%s] %s %s", context, message || "", sanitized ? JSON.stringify(sanitized) : "");

  // Gravação na planilha: DESATIVADA por padrão (CONFIG.DEBUG_SHEETS = false).
  // Ligue somente para depuração pontual — cada chamada custa ~1s de latência.
  if (!CONFIG.DEBUG_SHEETS) return;
  try {
    const sheet = getSheet_(CONFIG.SHEETS.LOGS_DEBUG);
    const email = Session.getActiveUser().getEmail() || "unknown";
    sheet.appendRow([
      new Date(),
      context,
      maskEmail_(email),
      message || "",
      sanitized ? JSON.stringify(sanitized) : ""
    ]);
    // Manter no máximo 1000 linhas de log (exclui as mais antigas)
    const MAX_LOG_ROWS = 1000;
    const lastRow = sheet.getLastRow();
    if (lastRow > MAX_LOG_ROWS + 1) {
      sheet.deleteRows(2, lastRow - MAX_LOG_ROWS - 1);
    }
  } catch (error) {
    Logger.log("logDebug_ sheet write failed: " + error);
  }
}

function safeLogDebug_(context, message, payload) {
  try {
    logDebug_(context, message, payload);
  } catch (error) {
    Logger.log("safeLogDebug_ failed: " + error);
  }
}
