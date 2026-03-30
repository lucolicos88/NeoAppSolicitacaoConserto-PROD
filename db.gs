// Memoiza a instância da planilha durante uma execução.
// SpreadsheetApp.openById() é uma chamada HTTP cara — chamá-la uma vez por execução economiza segundos.
let _ssInstance_ = null;
function getSpreadsheet_() {
  if (!_ssInstance_) {
    // PropertiesService é preferido: o ID não fica exposto no código-fonte do repositório.
    // Fallback para CONFIG.SPREADSHEET_ID durante a transição (remover após rodar setupPropriedades_).
    const props = PropertiesService.getScriptProperties();
    const spreadsheetId = props.getProperty("SPREADSHEET_ID") || CONFIG.SPREADSHEET_ID;
    if (!spreadsheetId) {
      throw new Error("SPREADSHEET_ID não configurado. Execute 'Configurar Propriedades' no menu App Solicitações.");
    }
    _ssInstance_ = SpreadsheetApp.openById(spreadsheetId);
  }
  return _ssInstance_;
}

function getSheet_(name) {
  return getSpreadsheet_().getSheetByName(name);
}

// Cache key para flag de sheets inicializadas (5 minutos)
const CACHE_KEY_SHEETS_OK_ = "app_sheets_ok_v1";

function ensureSheets_() {
  // Pula verificação se sheets já foram confirmadas recentemente
  const cache = CacheService.getScriptCache();
  if (cache.get(CACHE_KEY_SHEETS_OK_) === "1") return;

  const ss = getSpreadsheet_();
  ensureSheet_(ss, CONFIG.SHEETS.SOLICITACOES, [
    "id_solicitacao",
    "requisicao",
    "solicitante",
    "data_hora_pedido",
    "status",
    "criado_por_email",
    "criado_em"
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.ERROS, [
    "id_solicitacao",
    "sequencia_erro",
    "erro",
    "detalhamento",
    "setor_local",
    "diferenca_valor",
    "criado_em",
    "confirmacao_medica"
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.RESPOSTAS, [
    "id_resposta",
    "id_solicitacao",
    "sequencia_erro",
    "nome_responsavel",
    "email_responsavel",
    "erro_corrigido",
    "houve_diferenca_valor",
    "diferenca_valor_resposta",
    "observacoes",
    "data_hora_correcao",
    "criado_em"
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.AUDITORIA, [
    "id_auditoria",
    "email_usuario",
    "acao",
    "tabela",
    "chave_registro",
    "campo",
    "valor_anterior",
    "valor_novo",
    "ip",
    "data_hora"
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.LOGS_DEBUG, [
    "data_hora",
    "funcao",
    "email",
    "mensagem",
    "payload"
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.LIMIARES, [
    "horas_alerta",
    "horas_critico",
    "fuso_padrao",
    "atualizado_por",
    "atualizado_em"
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.SOLICITANTES, ["solicitante"]);
  ensureSheet_(ss, CONFIG.SHEETS.SETORES, ["setor_local"]);
  ensureSheet_(ss, CONFIG.SHEETS.ERROS_CADASTRO, ["erro"]);
  ensureSheet_(ss, CONFIG.SHEETS.USUARIOS, [
    "email",
    "nome",
    "perfil",
    "setores"
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.LISTAS, [
    "colaboradores",
    "setores"
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.CONFIG_GERAL, [
    "chave",
    "valor"
  ]);

  ensureThresholdDefault_();
  ensureConfigDefaults_();

  // Marca sheets como verificadas por 5 minutos
  try { cache.put(CACHE_KEY_SHEETS_OK_, "1", 300); } catch(e) { /* ignore */ }
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function ensureThresholdDefault_() {
  const sheet = getSheet_(CONFIG.SHEETS.LIMIARES);
  if (sheet.getLastRow() < 2) {
    sheet.appendRow([
      CONFIG.DEFAULTS.WARN_HOURS,
      CONFIG.DEFAULTS.CRITICAL_HOURS,
      CONFIG.DEFAULTS.TIMEZONE,
      Session.getActiveUser().getEmail(),
      new Date()
    ]);
  }
}

function ensureConfigDefaults_() {
  const solicitantes = getSheet_(CONFIG.SHEETS.SOLICITANTES);
  const setores = getSheet_(CONFIG.SHEETS.SETORES);
  const erros = getSheet_(CONFIG.SHEETS.ERROS_CADASTRO);
  const usuarios = getSheet_(CONFIG.SHEETS.USUARIOS);
  const listas = getSheet_(CONFIG.SHEETS.LISTAS);
  const configGeral = getSheet_(CONFIG.SHEETS.CONFIG_GERAL);

  if (solicitantes.getLastRow() < 2) {
    solicitantes.appendRow(["FARMÁCIA A"]);
    solicitantes.appendRow(["FARMÁCIA B"]);
  }
  if (setores.getLastRow() < 2) {
    setores.appendRow(["SETOR 01"]);
    setores.appendRow(["SETOR 02"]);
  }
  if (erros.getLastRow() < 2) {
    erros.appendRow(["ERRO PADRÃO"]);
    erros.appendRow(["ERRO CADASTRAL"]);
  }
  if (usuarios.getLastRow() < 2) {
    usuarios.appendRow([
      Session.getActiveUser().getEmail(),
      "Administrador",
      "ADMIN",
      "*"
    ]);
  }
  if (listas.getLastRow() < 2) {
    listas.appendRow(["Administrador", "SETOR 01"]);
  }
  if (configGeral.getLastRow() < 2) {
    configGeral.appendRow(["pasta_backup_id", ""]);
  }
}

function resetPlanilha_() {
  // Invalidar cache de sheets ao resetar para forçar recriação
  try { CacheService.getScriptCache().remove(CACHE_KEY_SHEETS_OK_); } catch(e) { /* ignore */ }

  const ss = getSpreadsheet_();
  const sheets = [
    CONFIG.SHEETS.SOLICITACOES,
    CONFIG.SHEETS.ERROS,
    CONFIG.SHEETS.RESPOSTAS,
    CONFIG.SHEETS.AUDITORIA,
    CONFIG.SHEETS.LOGS_DEBUG,
    CONFIG.SHEETS.LIMIARES,
    CONFIG.SHEETS.SOLICITANTES,
    CONFIG.SHEETS.SETORES,
    CONFIG.SHEETS.ERROS_CADASTRO,
    CONFIG.SHEETS.USUARIOS,
    CONFIG.SHEETS.LISTAS,
    CONFIG.SHEETS.CONFIG_GERAL
  ];

  sheets.forEach((name) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  });

  ensureSheets_();
  SpreadsheetApp.getUi().alert("Planilha resetada com as abas necessárias.");
}

function findSolicitacaoRow_(solicitacaoId) {
  const sheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === solicitacaoId) {
      return { row: i + 1, data: values[i] };
    }
  }
  return null;
}

function getNextErroSeq_(solicitacaoId) {
  const sheet = getSheet_(CONFIG.SHEETS.ERROS);
  const values = sheet.getDataRange().getValues();
  let maxSeq = 0;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === solicitacaoId) {
      maxSeq = Math.max(maxSeq, Number(values[i][1]) || 0);
    }
  }
  return maxSeq + 1;
}

function insertSolicitacaoEErro_(payload, userEmail) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    const solicitacoesSheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);
    const errosSheet = getSheet_(CONFIG.SHEETS.ERROS);

    let solicitacaoId = payload.solicitacaoId || Utilities.getUuid();
    let solicitacaoRow = findSolicitacaoRow_(solicitacaoId);
    if (!solicitacaoRow) {
      solicitacoesSheet.appendRow([
        solicitacaoId,
        payload.requisicao,
        payload.solicitante,
        new Date(payload.dataHoraPedido),
        "ABERTO",
        userEmail,
        now
      ]);
      logDebug_("insertSolicitacaoEErro_", "nova_solicitacao", {
        solicitacaoId: solicitacaoId,
        requisicao: payload.requisicao
      });
      auditLog_(
        "CREATE",
        "solicitacoes",
        "id_solicitacao=" + solicitacaoId,
        null,
        null,
        null
      );
    }

    const erroSeq = payload.erroSeq || getNextErroSeq_(solicitacaoId);
    errosSheet.appendRow([
      solicitacaoId,
      erroSeq,
      payload.erro,
      payload.detalhamento,
      payload.setorLocal,
      payload.diferencaNoValor || "",
      now,
      payload.confirmacaoMedica || "NAO"
    ]);
    logDebug_("insertSolicitacaoEErro_", "novo_erro", {
      solicitacaoId: solicitacaoId,
      erroSeq: erroSeq,
      setorLocal: payload.setorLocal
    });
    auditLog_(
      "CREATE",
      "erros",
      "id_solicitacao=" + solicitacaoId + ";sequencia_erro=" + erroSeq,
      null,
      null,
      null
    );

    return { solicitacaoId: solicitacaoId, erroSeq: erroSeq };
  } finally {
    lock.releaseLock();
  }
}

function deleteSolicitacao_(solicitacaoId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const solicitacoesSheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);
    const errosSheet = getSheet_(CONFIG.SHEETS.ERROS);
    const respostasSheet = getSheet_(CONFIG.SHEETS.RESPOSTAS);

    // Find and delete solicitation
    const solicitacaoRow = findSolicitacaoRow_(solicitacaoId);
    if (!solicitacaoRow) {
      logDebug_("deleteSolicitacao_", "solicitacao_not_found", { solicitacaoId: solicitacaoId });
      return null;
    }

    // Batch delete erros: filtra as linhas a manter e reescreve de uma vez
    // Muito mais rápido que N chamadas deleteRow() individuais
    const errosData = errosSheet.getDataRange().getValues();
    const errosRemaining = errosData.filter(function(row, idx) {
      return idx === 0 || row[0] !== solicitacaoId;
    });
    const errosDeleted = errosData.length - errosRemaining.length;
    if (errosDeleted > 0) {
      errosSheet.clearContents();
      errosSheet.getRange(1, 1, errosRemaining.length, errosRemaining[0].length).setValues(errosRemaining);
    }

    // Batch delete respostas: mesmo padrão
    const respostasData = respostasSheet.getDataRange().getValues();
    const respostasRemaining = respostasData.filter(function(row, idx) {
      return idx === 0 || row[1] !== solicitacaoId;
    });
    const respostasDeleted = respostasData.length - respostasRemaining.length;
    if (respostasDeleted > 0) {
      respostasSheet.clearContents();
      respostasSheet.getRange(1, 1, respostasRemaining.length, respostasRemaining[0].length).setValues(respostasRemaining);
    }

    // Finally, delete the solicitation itself
    solicitacoesSheet.deleteRow(solicitacaoRow.row);

    auditLog_(
      "DELETE",
      "solicitacoes",
      "id_solicitacao=" + solicitacaoId,
      null,
      null,
      null
    );

    logDebug_("deleteSolicitacao_", "solicitacao_deleted", {
      solicitacaoId: solicitacaoId,
      errosDeleted: errosDeleted,
      respostasDeleted: respostasDeleted
    });

    return { solicitacaoId: solicitacaoId };
  } finally {
    lock.releaseLock();
  }
}

function insertResposta_(payload, userEmail) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    const respostasSheet = getSheet_(CONFIG.SHEETS.RESPOSTAS);
    const respostaId = Utilities.getUuid();

    respostasSheet.appendRow([
      respostaId,
      payload.solicitacaoId,
      payload.erroSeq,
      payload.nomeResponsavel || "",
      userEmail,
      "SIM",
      payload.houveDiferencaValor || "",
      payload.diferencaValorResposta || "",
      payload.observacoes || "",
      payload.dataHoraCorrecao ? new Date(payload.dataHoraCorrecao) : "",
      now
    ]);
    logDebug_("insertResposta_", "nova_resposta", {
      respostaId: respostaId,
      solicitacaoId: payload.solicitacaoId,
      erroSeq: payload.erroSeq
    });
    auditLog_(
      "CREATE",
      "respostas",
      "id_resposta=" + respostaId,
      null,
      null,
      null
    );

    updateSolicitacaoStatus_(payload.solicitacaoId);
    return { respostaId: respostaId };
  } finally {
    lock.releaseLock();
  }
}

function updateResposta_(payload, userEmail) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const respostasSheet = getSheet_(CONFIG.SHEETS.RESPOSTAS);
    const respostas = respostasSheet.getDataRange().getValues();

    // Find the latest response for this solicitacaoId + erroSeq
    let rowToUpdate = -1;
    let respostaId = null;
    for (let i = respostas.length - 1; i >= 1; i--) {
      if (respostas[i][1] === payload.solicitacaoId && respostas[i][2] === payload.erroSeq) {
        rowToUpdate = i + 1; // Convert to 1-indexed
        respostaId = respostas[i][0];
        break;
      }
    }

    if (rowToUpdate === -1) {
      logDebug_("updateResposta_", "resposta_not_found", {
        solicitacaoId: payload.solicitacaoId,
        erroSeq: payload.erroSeq
      });
      return null;
    }

    const now = new Date();

    // Atualiza cols 4–11 em uma única chamada ao Sheets (7x mais rápido que setValue individual)
    respostasSheet.getRange(rowToUpdate, 4, 1, 8).setValues([[
      payload.nomeResponsavel || "",                                          // col 4: nome_responsavel
      userEmail,                                                               // col 5: email_responsavel
      "SIM",                                                                   // col 6: erro_corrigido (mantém SIM)
      payload.houveDiferencaValor || "",                                      // col 7: houve_diferenca_valor
      payload.diferencaValorResposta || "",                                   // col 8: diferenca_valor_resposta
      payload.observacoes || "",                                               // col 9: observacoes
      payload.dataHoraCorrecao ? new Date(payload.dataHoraCorrecao) : "",    // col 10: data_hora_correcao
      now                                                                      // col 11: criado_em (timestamp update)
    ]]);

    logDebug_("updateResposta_", "resposta_atualizada", {
      respostaId: respostaId,
      solicitacaoId: payload.solicitacaoId,
      erroSeq: payload.erroSeq,
      row: rowToUpdate
    });

    auditLog_(
      "UPDATE",
      "respostas",
      "id_resposta=" + respostaId,
      "resposta_atualizada",
      null,
      null
    );

    updateSolicitacaoStatus_(payload.solicitacaoId);
    return { respostaId: respostaId };
  } finally {
    lock.releaseLock();
  }
}

function updateSolicitacaoStatus_(solicitacaoId) {
  const solicitacaoRow = findSolicitacaoRow_(solicitacaoId);
  if (!solicitacaoRow) {
    return;
  }

  const errosSheet = getSheet_(CONFIG.SHEETS.ERROS);
  const respostasSheet = getSheet_(CONFIG.SHEETS.RESPOSTAS);
  const erros = errosSheet.getDataRange().getValues();
  const respostas = respostasSheet.getDataRange().getValues();

  let totalErros = 0;
  let errosCorrigidos = 0;
  for (let i = 1; i < erros.length; i++) {
    if (erros[i][0] === solicitacaoId) {
      totalErros++;
      const erroSeq = erros[i][1];
      const resposta = findLatestResposta_(respostas, solicitacaoId, erroSeq);
      if (resposta && resposta[5] === "SIM") {
        errosCorrigidos++;
      }
    }
  }

  const newStatus = totalErros > 0 && errosCorrigidos === totalErros ? "CORRIGIDO" : "EM_CORRECAO";
  const sheet = getSheet_(CONFIG.SHEETS.SOLICITACOES);
  const statusCell = sheet.getRange(solicitacaoRow.row, 5);
  const oldValue = statusCell.getValue();
  if (oldValue !== newStatus) {
    statusCell.setValue(newStatus);
    auditLog_(
      "UPDATE",
      "solicitacoes",
      "id_solicitacao=" + solicitacaoId,
      "status",
      oldValue,
      newStatus
    );
  }
}

function findLatestResposta_(respostas, solicitacaoId, erroSeq) {
  let latest = null;
  for (let i = 1; i < respostas.length; i++) {
    if (respostas[i][1] === solicitacaoId && respostas[i][2] === erroSeq) {
      latest = respostas[i];
    }
  }
  return latest;
}

function getConfigGeral_(chave) {
  const sheet = getSheet_(CONFIG.SHEETS.CONFIG_GERAL);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === chave) {
      return values[i][1];
    }
  }
  return "";
}

function setConfigGeral_(chave, valor) {
  const sheet = getSheet_(CONFIG.SHEETS.CONFIG_GERAL);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === chave) {
      sheet.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }
  sheet.appendRow([chave, valor]);
}
