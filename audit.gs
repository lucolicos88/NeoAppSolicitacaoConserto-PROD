function auditLog_(actionType, tableName, recordKey, fieldName, oldValue, newValue) {
  const sheet = getSheet_(CONFIG.SHEETS.AUDITORIA);
  const userEmail = Session.getActiveUser().getEmail() || "unknown";
  sheet.appendRow([
    Utilities.getUuid(),
    userEmail,
    actionType,
    tableName,
    recordKey,
    fieldName || "",
    oldValue || "",
    newValue || "",
    getClientIp_(),
    new Date()
  ]);
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

function logDebug_(context, message, payload) {
  if (!CONFIG.DEBUG_ENABLED) return;
  try {
    const sheet = getSheet_(CONFIG.SHEETS.LOGS_DEBUG);
    const email = Session.getActiveUser().getEmail() || "unknown";
    const sanitized = sanitizePayloadForLog_(payload);
    sheet.appendRow([
      new Date(),
      context,
      maskEmail_(email),
      message || "",
      sanitized ? JSON.stringify(sanitized) : ""
    ]);
    Logger.log("[%s] %s", context, message || "");
    // Manter no máximo 1000 linhas de log (exclui as mais antigas)
    const MAX_LOG_ROWS = 1000;
    const lastRow = sheet.getLastRow();
    if (lastRow > MAX_LOG_ROWS + 1) {
      sheet.deleteRows(2, lastRow - MAX_LOG_ROWS - 1);
    }
  } catch (error) {
    Logger.log("logDebug_ failed: " + error);
  }
}

function safeLogDebug_(context, message, payload) {
  try {
    logDebug_(context, message, payload);
  } catch (error) {
    Logger.log("safeLogDebug_ failed: " + error);
  }
}
