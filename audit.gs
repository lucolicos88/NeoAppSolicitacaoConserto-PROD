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
  try {
    return Session.getTemporaryActiveUserKey();
  } catch (error) {
    return "unknown";
  }
}

function logDebug_(context, message, payload) {
  if (!CONFIG.DEBUG_ENABLED) return;
  try {
    const sheet = getSheet_(CONFIG.SHEETS.LOGS_DEBUG);
    const email = Session.getActiveUser().getEmail() || "unknown";
    sheet.appendRow([
      new Date(),
      context,
      email,
      message || "",
      payload ? JSON.stringify(payload) : ""
    ]);
    Logger.log("[%s] %s %s", context, message, payload ? JSON.stringify(payload) : "");
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
