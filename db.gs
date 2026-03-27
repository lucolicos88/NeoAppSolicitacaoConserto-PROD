function getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet_(name) {
  return getSpreadsheet_().getSheetByName(name);
}

function ensureSheets_() {
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

    // Delete all associated errors
    const errosData = errosSheet.getDataRange().getValues();
    const errosToDelete = [];
    for (let i = errosData.length - 1; i >= 1; i--) {
      if (errosData[i][0] === solicitacaoId) {
        errosToDelete.push(i + 1); // Convert to 1-indexed
      }
    }

    for (const rowNum of errosToDelete) {
      errosSheet.deleteRow(rowNum);
      logDebug_("deleteSolicitacao_", "erro_deleted", { solicitacaoId: solicitacaoId, row: rowNum });
    }

    // Delete all associated responses
    const respostasData = respostasSheet.getDataRange().getValues();
    const respostasToDelete = [];
    for (let i = respostasData.length - 1; i >= 1; i--) {
      if (respostasData[i][1] === solicitacaoId) {
        respostasToDelete.push(i + 1); // Convert to 1-indexed
      }
    }

    for (const rowNum of respostasToDelete) {
      respostasSheet.deleteRow(rowNum);
      logDebug_("deleteSolicitacao_", "resposta_deleted", { solicitacaoId: solicitacaoId, row: rowNum });
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
      errosDeleted: errosToDelete.length,
      respostasDeleted: respostasToDelete.length
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

    // Update the existing row (columns D, E, G, H, I, J)
    respostasSheet.getRange(rowToUpdate, 4).setValue(payload.nomeResponsavel || ""); // Nome Responsável
    respostasSheet.getRange(rowToUpdate, 5).setValue(userEmail); // Email Responsável
    // Column 6 (Correção Finalizada) stays as "SIM"
    respostasSheet.getRange(rowToUpdate, 7).setValue(payload.houveDiferencaValor || ""); // Houve Diferença Valor
    respostasSheet.getRange(rowToUpdate, 8).setValue(payload.diferencaValorResposta || ""); // Diferença Valor Resposta
    respostasSheet.getRange(rowToUpdate, 9).setValue(payload.observacoes || ""); // Observações
    respostasSheet.getRange(rowToUpdate, 10).setValue(payload.dataHoraCorrecao ? new Date(payload.dataHoraCorrecao) : ""); // Data/Hora Correção
    respostasSheet.getRange(rowToUpdate, 11).setValue(now); // Data/Hora Cadastro (update timestamp)

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
