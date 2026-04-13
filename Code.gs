/**
 * ╔══════════════════════════════════════════════════════════════════════════════╗
 * ║                    SISTEMA DE SOLICITAÇÃO DE CONSERTOS                       ║
 * ║                           Biblioteca Principal                               ║
 * ╠══════════════════════════════════════════════════════════════════════════════╣
 * ║  Arquivo: Code.gs                                                            ║
 * ║  Versão: 2.8.31                                                              ║
 * ║  Última atualização: Janeiro 2026                                            ║
 * ╠══════════════════════════════════════════════════════════════════════════════╣
 * ║                                                                              ║
 * ║  DESCRIÇÃO:                                                                  ║
 * ║  Este arquivo contém as funções principais do backend do sistema de          ║
 * ║  solicitações de consertos. Inclui configurações globais, funções de         ║
 * ║  inicialização, formatação de planilhas e utilitários de teste.              ║
 * ║                                                                              ║
 * ║  ESTRUTURA DO ARQUIVO:                                                       ║
 * ║  ├── Configurações Globais (CONFIG)                                          ║
 * ║  ├── Funções de Inicialização (doGet, onOpen)                               ║
 * ║  ├── Funções de Teste de Performance                                         ║
 * ║  ├── Funções de Formatação de Planilhas                                      ║
 * ║  └── Funções Auxiliares Internas                                             ║
 * ║                                                                              ║
 * ║  DEPENDÊNCIAS:                                                               ║
 * ║  - api.gs: Funções de API e operações de dados                              ║
 * ║  - ui.html: Interface do usuário                                            ║
 * ║                                                                              ║
 * ║  PLANILHAS UTILIZADAS:                                                       ║
 * ║  - Solicitacoes: Registro de todas as solicitações                          ║
 * ║  - Erros: Tipos de erro de cada solicitação                                 ║
 * ║  - Respostas: Respostas/correções realizadas                                ║
 * ║  - Auditoria: Log de alterações do sistema                                   ║
 * ║  - Usuarios: Cadastro de usuários e permissões                              ║
 * ║  - Config_Geral: Configurações gerais do sistema                            ║
 * ║                                                                              ║
 * ╚══════════════════════════════════════════════════════════════════════════════╝
 */

// ============================================================================
// CONFIGURAÇÕES GLOBAIS
// ============================================================================

/**
 * URL do logotipo exibido no cabeçalho da aplicação
 * @constant {string}
 */
const LOGO_URL = "https://neoformula.com.br/cdn/shop/files/Logotipo-NeoFormula-Manipulacao-Homeopatia_76b2fa98-5ffa-4cc3-ac0a-6d41e1bc8810.png?height=100&v=1677088468";

/**
 * Configurações centrais do sistema
 * @constant {Object}
 * @property {string} SPREADSHEET_ID - ID da planilha Google Sheets principal
 * @property {Object} SHEETS - Nomes das abas da planilha
 * @property {Object} DEFAULTS - Valores padrão do sistema
 * @property {boolean} DEBUG_ENABLED - Flag para ativar logs de debug
 */
const CONFIG = {
  SPREADSHEET_ID: "", // Configurado via PropertiesService — execute "Configurar Propriedades" no menu
  SHEETS: {
    SOLICITACOES: "Solicitacoes",
    ERROS: "Erros",
    RESPOSTAS: "Respostas",
    AUDITORIA: "Auditoria",
    LOGS_DEBUG: "Logs_Debug",
    LIMIARES: "Limiares_SLA",
    SOLICITANTES: "Solicitantes",
    SETORES: "Setores_Local",
    ERROS_CADASTRO: "Erros_Cadastro",
    USUARIOS: "Usuarios",
    LISTAS: "Colaboradores",
    CONFIG_GERAL: "Config_Geral"
  },
  DEFAULTS: {
    WARN_MINUTES: 15,
    CRITICAL_MINUTES: 30,
    TIMEZONE: "America/Sao_Paulo"
  },
  APP_VERSION: "v247",
  APP_ENV: "PROD",
  // ID da planilha PROD — usado apenas no ambiente DEV para importação de dados
  PROD_SPREADSHEET_ID: "1h4bCllbefqsmsjXpMSSRXVR6avdeXgDS3uGA3NercH8"
};
// DEBUG_ENABLED automático: ativo em DEV, desativado em PROD
CONFIG.DEBUG_ENABLED = CONFIG.APP_ENV === "DEV";

// ============================================================================
// FUNÇÕES DE INICIALIZAÇÃO
// ============================================================================

/**
 * Ponto de entrada principal da aplicação web
 * Chamada automaticamente quando o usuário acessa a URL do Web App
 *
 * @returns {HtmlOutput} Página HTML renderizada da aplicação
 *
 * @description
 * Esta função:
 * 1. Garante que todas as planilhas necessárias existam
 * 2. Carrega o template HTML (ui.html)
 * 3. Injeta variáveis de contexto (logo, email do usuário)
 * 4. Configura metadados da página (título, viewport)
 */
function doGet() {
  try {
    ensureSheets_();
  } catch(e) {
    Logger.log('Error in ensureSheets_: ' + e.toString());
  }

  const tpl = HtmlService.createTemplateFromFile("ui");
  tpl.logoUrl = LOGO_URL || "";
  tpl.appVersion = CONFIG.APP_VERSION || "v???";
  tpl.appEnv = CONFIG.APP_ENV || "DEV";

  try {
    tpl.userEmail = Session.getActiveUser().getEmail() || "usuario@exemplo.com";
  } catch(e) {
    tpl.userEmail = "usuario@exemplo.com";
  }

  return tpl
    .evaluate()
    .setTitle("Solicitacoes de Conserto")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Inclui conteúdo de arquivo HTML no template
 * Usado para modularização do código HTML
 *
 * @param {string} filename - Nome do arquivo HTML a ser incluído
 * @returns {string} Conteúdo HTML do arquivo
 * @private
 */
function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Cria menu personalizado na planilha Google Sheets
 * Chamada automaticamente ao abrir a planilha
 *
 * @description
 * Adiciona menu "App Solicitações" com opções:
 * - Setup/Reset Planilha: Inicializa estrutura de dados
 * - Formatar Planilhas: Aplica formatação visual
 * - Gerar Dados de Teste: Cria dados para testes de performance
 * - Limpar Dados de Teste: Remove todos os dados de teste
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("App Solicitações")
    .addItem("Configurar Propriedades (rodar 1x)", "setupPropriedades_")
    .addSeparator()
    .addItem("Setup/Reset Planilha", "resetPlanilha_")
    .addItem("Formatar Planilhas", "formatarPlanilhas_")
    .addSeparator()
    .addItem("Gerar Dados de Teste (2000+)", "gerarDadosTeste_")
    .addItem("Limpar Dados de Teste", "limparDadosTeste_")
    .addToUi();
}

/**
 * Salva o SPREADSHEET_ID no PropertiesService do script.
 * Rodar UMA VEZ após o primeiro deploy em cada ambiente (DEV e PROD).
 * Após rodar, o ID não precisa mais ficar exposto no código-fonte.
 */
function setupPropriedades_() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  // Verifica se já está configurado
  const atual = props.getProperty("SPREADSHEET_ID");
  if (atual) {
    const resp = ui.alert(
      "Propriedades já configuradas",
      "SPREADSHEET_ID já está salvo:\n" + atual + "\n\nDeseja sobrescrever com o valor do código (" + CONFIG.SPREADSHEET_ID + ")?",
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) {
      ui.alert("Operação cancelada. Propriedades mantidas.");
      return;
    }
  }

  if (!CONFIG.SPREADSHEET_ID) {
    ui.alert("Erro: CONFIG.SPREADSHEET_ID está vazio no código. Nada foi salvo.");
    return;
  }

  props.setProperty("SPREADSHEET_ID", CONFIG.SPREADSHEET_ID);
  ui.alert(
    "Propriedades configuradas!",
    "SPREADSHEET_ID salvo no PropertiesService:\n" + CONFIG.SPREADSHEET_ID +
    "\n\nO código agora lê o ID das propriedades — ele pode ser removido do CONFIG com segurança.",
    ui.ButtonSet.OK
  );
}

// ============================================================================
// FUNÇÕES DE TESTE DE PERFORMANCE
// ============================================================================

/**
 * Gera dados de teste para validação de performance
 * Acessível via menu da planilha
 *
 * @description
 * Cria automaticamente:
 * - 3000 solicitações com datas aleatórias (últimos 90 dias)
 * - 3000 erros (1 por solicitação)
 * - 2950 respostas (~98.3% das solicitações)
 *
 * @requires Permissão ADMIN
 * @sideeffect Adiciona dados às planilhas Solicitações, Erros e Respostas
 */
function gerarDadosTeste_() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Gerar Dados de Teste',
    'Isso irá criar:\n' +
    '- 3000 solicitações\n' +
    '- 3000 erros (1 por solicitação)\n' +
    '- 2950 respostas\n\n' +
    'Os dados existentes NÃO serão apagados.\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  const startTime = new Date();
  ui.alert('Iniciando geração de dados. Isso pode levar alguns minutos...');

  try {
    const stats = gerarDadosTesteInterno_(3000);
    const elapsed = Math.round((new Date() - startTime) / 1000);

    ui.alert(
      'Dados de Teste Gerados!',
      'Tempo: ' + elapsed + ' segundos\n\n' +
      'Solicitações criadas: ' + stats.solicitacoes + '\n' +
      'Erros criados: ' + stats.erros + '\n' +
      'Respostas criadas: ' + stats.respostas,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro ao gerar dados: ' + e.toString());
  }
}

function gerarDadosTesteInterno_(quantidade) {
  const ss = getSpreadsheet_();

  // Carregar listas existentes ou usar defaults
  const solicitantesSheet = ss.getSheetByName(CONFIG.SHEETS.SOLICITANTES);
  const setoresSheet = ss.getSheetByName(CONFIG.SHEETS.SETORES);
  const errosSheet = ss.getSheetByName(CONFIG.SHEETS.ERROS_CADASTRO);
  const colaboradoresSheet = ss.getSheetByName(CONFIG.SHEETS.LISTAS);

  let solicitantes = solicitantesSheet ?
    solicitantesSheet.getDataRange().getValues().slice(1).map(r => r[0]).filter(v => v) : [];
  let setores = setoresSheet ?
    setoresSheet.getDataRange().getValues().slice(1).map(r => r[0]).filter(v => v) : [];
  let tiposErro = errosSheet ?
    errosSheet.getDataRange().getValues().slice(1).map(r => r[0]).filter(v => v) : [];
  let colaboradores = colaboradoresSheet ?
    colaboradoresSheet.getDataRange().getValues().slice(1).map(r => r[0]).filter(v => v) : [];

  // Defaults se listas vazias
  if (solicitantes.length === 0) {
    solicitantes = ['Ana Silva', 'João Santos', 'Maria Oliveira', 'Pedro Costa', 'Carla Souza',
                   'Lucas Ferreira', 'Juliana Lima', 'Roberto Alves', 'Fernanda Gomes', 'Carlos Pereira'];
  }
  if (setores.length === 0) {
    setores = ['Bosque', 'Glicério', 'Cambuí', 'Guanabara', 'Convênio', 'Virtual', 'e-Commerce'];
  }
  if (tiposErro.length === 0) {
    tiposErro = [
      'Fórmula Excluída / Incluída fora do horário',
      'Cadastro Incompleto',
      'Falta de observações',
      'Forma farmacêutica errada',
      'Nome do paciente',
      'Nome do médico',
      'Número da Notificação',
      'Via de administração',
      'Dose',
      'Prescrição proibida por profissional'
    ];
  }
  if (colaboradores.length === 0) {
    colaboradores = ['Andre Vinicius Rossi de Sousa Ribeiro', 'Ana Júlia Lopes Silva',
                    'Patricia Magalhães', 'Beatriz Santos', 'Camila Oliveira'];
  }

  const detalhamentos = [
    'Erro identificado na conferência',
    'Cliente reclamou do problema',
    'Verificado durante auditoria',
    'Detectado no sistema',
    'Reportado pelo farmacêutico',
    'Encontrado na revisão',
    'Notificado pelo cliente',
    'Identificado na entrega'
  ];

  const observacoes = [
    'Corrigido conforme procedimento',
    'Ajustado e verificado',
    'Problema resolvido',
    'Correção aplicada com sucesso',
    'Situação normalizada',
    'Erro sanado',
    'Conferido e aprovado',
    ''
  ];

  // Gerar dados em lotes
  const solicitacoesData = [];
  const errosData = [];
  const respostasData = [];

  const now = new Date();
  // Email genérico para dados de teste (LGPD: não gravar email real em registros de teste)
  const testEmail = 'teste@neoformula.dev';
  // Email real apenas para configurar o admin ao final
  const adminEmail = Session.getActiveUser().getEmail() || 'admin@neoformula.dev';

  // Calcular quantas respostas gerar (2950 de 3000 = 98.33%)
  const qtdRespostas = Math.min(2950, quantidade);

  for (let i = 0; i < quantidade; i++) {
    const solicitacaoId = Utilities.getUuid();
    const diasAtras = Math.floor(Math.random() * 90); // Últimos 90 dias
    const dataHoraPedido = new Date(now.getTime() - diasAtras * 24 * 60 * 60 * 1000);
    dataHoraPedido.setHours(Math.floor(Math.random() * 12) + 7); // 7h às 19h
    dataHoraPedido.setMinutes(Math.floor(Math.random() * 60));

    const requisicao = String(100000 + Math.floor(Math.random() * 900000));
    const solicitante = solicitantes[Math.floor(Math.random() * solicitantes.length)];

    // Primeiras qtdRespostas serão respondidas, restante fica ABERTO
    const temResposta = i < qtdRespostas;
    const status = temResposta ? 'CORRIGIDO' : 'ABERTO';

    // Solicitação
    solicitacoesData.push([
      solicitacaoId,
      requisicao,
      solicitante,
      dataHoraPedido,
      status,
      testEmail,
      dataHoraPedido
    ]);

    // Exatamente 1 erro por solicitação
    const tipoErro = tiposErro[Math.floor(Math.random() * tiposErro.length)];
    const setor = setores[Math.floor(Math.random() * setores.length)];
    const detalhamento = detalhamentos[Math.floor(Math.random() * detalhamentos.length)];
    const diferencaValor = Math.random() < 0.1 ? 'SIM' : 'NÃO'; // 10% com diferença

    errosData.push([
      solicitacaoId,
      1, // Sempre sequência 1 (1 erro por solicitação)
      tipoErro,
      detalhamento,
      setor,
      diferencaValor,
      dataHoraPedido
    ]);

    // Criar resposta se status não for ABERTO
    if (temResposta) {
      const horasParaCorrecao = Math.floor(Math.random() * 48) + 1; // 1 a 48 horas
      const dataCorrecao = new Date(dataHoraPedido.getTime() + horasParaCorrecao * 60 * 60 * 1000);
      const responsavel = colaboradores[Math.floor(Math.random() * colaboradores.length)];
      const obs = observacoes[Math.floor(Math.random() * observacoes.length)];
      const valorDif = diferencaValor === 'SIM' ? (Math.random() * 100).toFixed(2).replace(".", ",") : '';

      respostasData.push([
        Utilities.getUuid(),
        solicitacaoId,
        1, // Sempre sequência 1
        responsavel,
        testEmail,
        'SIM',
        diferencaValor,
        valorDif,
        obs,
        dataCorrecao,
        dataCorrecao
      ]);
    }
  }

  // Inserir em lote (muito mais rápido que appendRow)
  const solicitacoesSheet = ss.getSheetByName(CONFIG.SHEETS.SOLICITACOES);
  const errosSheetData = ss.getSheetByName(CONFIG.SHEETS.ERROS);
  const respostasSheetData = ss.getSheetByName(CONFIG.SHEETS.RESPOSTAS);

  if (solicitacoesData.length > 0) {
    const lastRow = solicitacoesSheet.getLastRow();
    solicitacoesSheet.getRange(lastRow + 1, 1, solicitacoesData.length, solicitacoesData[0].length)
      .setValues(solicitacoesData);
  }

  if (errosData.length > 0) {
    const lastRow = errosSheetData.getLastRow();
    errosSheetData.getRange(lastRow + 1, 1, errosData.length, errosData[0].length)
      .setValues(errosData);
  }

  if (respostasData.length > 0) {
    const lastRow = respostasSheetData.getLastRow();
    respostasSheetData.getRange(lastRow + 1, 1, respostasData.length, respostasData[0].length)
      .setValues(respostasData);
  }

  // Garantir que o usuário atual tenha acesso ADMIN (usa email real)
  configurarUsuarioAdmin_(adminEmail, setores);

  return {
    solicitacoes: solicitacoesData.length,
    erros: errosData.length,
    respostas: respostasData.length
  };
}

// Configura o usuário atual com perfil ADMIN e acesso a todos os setores
function configurarUsuarioAdmin_(email, setores) {
  const ss = getSpreadsheet_();

  // Garantir que os setores usados existam na lista
  const setoresSheet = ss.getSheetByName(CONFIG.SHEETS.SETORES);
  if (setoresSheet) {
    const existentes = setoresSheet.getDataRange().getValues().slice(1).map(r => r[0]);
    setores.forEach(setor => {
      if (setor && !existentes.includes(setor)) {
        setoresSheet.appendRow([setor]);
        existentes.push(setor);
      }
    });
  }

  // Configurar usuário ADMIN
  const usuariosSheet = ss.getSheetByName(CONFIG.SHEETS.USUARIOS);
  if (!usuariosSheet) return;

  const values = usuariosSheet.getDataRange().getValues();

  // Verificar se usuário já existe
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === email) {
      // Atualizar para ADMIN com todos os setores
      usuariosSheet.getRange(i + 1, 3).setValue('ADMIN');
      usuariosSheet.getRange(i + 1, 4).setValue('*');
      return;
    }
  }

  // Adicionar novo usuário ADMIN
  const nome = email.split('@')[0].replace(/[._]/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
  usuariosSheet.appendRow([email, nome, 'ADMIN', '*']);
}

function limparDadosTeste_() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Limpar Dados de Teste',
    '⚠️ ATENÇÃO: Isso irá APAGAR TODOS os dados das abas:\n\n' +
    '- Solicitações\n' +
    '- Erros\n' +
    '- Respostas\n' +
    '- Auditoria\n' +
    '- Logs_Debug\n\n' +
    'Esta ação NÃO pode ser desfeita!\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  // Confirmação extra
  const confirm = ui.alert(
    'Confirmação Final',
    'Tem CERTEZA que deseja apagar todos os dados?\n\nDigite SIM para confirmar.',
    ui.ButtonSet.OK_CANCEL
  );

  if (confirm !== ui.Button.OK) return;

  const ss = getSpreadsheet_();
  const sheetsToClean = [
    CONFIG.SHEETS.SOLICITACOES,
    CONFIG.SHEETS.ERROS,
    CONFIG.SHEETS.RESPOSTAS,
    CONFIG.SHEETS.AUDITORIA,
    CONFIG.SHEETS.LOGS_DEBUG
  ];

  let totalLimpo = 0;
  sheetsToClean.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // Usar clear ao invés de deleteRows para evitar erros
        const numRows = lastRow - 1;
        sheet.getRange(2, 1, numRows, sheet.getLastColumn()).clearContent();
        totalLimpo += numRows;

        // Tentar deletar linhas extras (deixando pelo menos 1 linha de dados vazia)
        try {
          if (lastRow > 2) {
            sheet.deleteRows(3, lastRow - 2);
          }
        } catch (e) {
          // Ignorar erro se não conseguir deletar linhas
        }
      }
    }
  });

  ui.alert('Dados limpos com sucesso!\n\nLinhas limpas: ' + totalLimpo);
}

function initDataStore_() {
  ensureSheets_();
  SpreadsheetApp.getUi().alert("Planilha inicializada com as abas necessarias.");
}

// ========== FORMATAÇÃO DAS PLANILHAS ==========

function formatarPlanilhas_() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Formatando planilhas... Aguarde.');

  try {
    const ss = getSpreadsheet_();
    let formatadas = 0;

    // Formatar cada sheet de dados
    formatadas += formatarSheet_(ss, CONFIG.SHEETS.SOLICITACOES, {
      headerColor: '#1e88e5',
      alternateColor: '#e3f2fd',
      columns: [
        { width: 280, format: 'TEXT' },      // [0] id_solicitacao
        { width: 100, format: 'TEXT' },      // [1] requisicao
        { width: 180, format: 'TEXT' },      // [2] solicitante
        { width: 150, format: 'DATETIME' },  // [3] data_hora_pedido
        { width: 100, format: 'TEXT' },      // [4] status
        { width: 200, format: 'TEXT' },      // [5] criado_por_email
        { width: 150, format: 'DATETIME' }   // [6] criado_em
      ]
    });

    formatadas += formatarSheet_(ss, CONFIG.SHEETS.ERROS, {
      headerColor: '#e53935',
      alternateColor: '#ffebee',
      columns: [
        { width: 280, format: 'TEXT' },      // [0] id_solicitacao
        { width: 80,  format: 'NUMBER' },    // [1] sequencia_erro
        { width: 250, format: 'TEXT' },      // [2] erro
        { width: 300, format: 'TEXT' },      // [3] detalhamento
        { width: 120, format: 'TEXT' },      // [4] setor_local
        { width: 100, format: 'TEXT' },      // [5] diferenca_valor
        { width: 150, format: 'DATETIME' },  // [6] criado_em
        { width: 130, format: 'TEXT' }       // [7] confirmacao_medica
      ]
    });

    formatadas += formatarSheet_(ss, CONFIG.SHEETS.RESPOSTAS, {
      headerColor: '#43a047',
      alternateColor: '#e8f5e9',
      columns: [
        { width: 280, format: 'TEXT' },      // [0]  id_resposta
        { width: 280, format: 'TEXT' },      // [1]  id_solicitacao
        { width: 80,  format: 'NUMBER' },    // [2]  sequencia_erro
        { width: 200, format: 'TEXT' },      // [3]  nome_responsavel
        { width: 200, format: 'TEXT' },      // [4]  email_responsavel
        { width: 80,  format: 'TEXT' },      // [5]  erro_corrigido
        { width: 80,  format: 'TEXT' },      // [6]  houve_diferenca_valor
        { width: 120, format: 'CURRENCY' },  // [7]  diferenca_valor_resposta
        { width: 250, format: 'TEXT' },      // [8]  observacoes
        { width: 150, format: 'DATETIME' },  // [9]  data_hora_correcao
        { width: 150, format: 'DATETIME' }   // [10] criado_em
      ]
    });

    formatadas += formatarSheet_(ss, CONFIG.SHEETS.AUDITORIA, {
      headerColor: '#7b1fa2',
      alternateColor: '#f3e5f5',
      columns: [
        { width: 280, format: 'TEXT' },      // id_auditoria
        { width: 200, format: 'TEXT' },      // email_usuario
        { width: 100, format: 'TEXT' },      // acao
        { width: 120, format: 'TEXT' },      // tabela
        { width: 280, format: 'TEXT' },      // chave_registro
        { width: 150, format: 'TEXT' },      // campo
        { width: 200, format: 'TEXT' },      // valor_anterior
        { width: 200, format: 'TEXT' },      // valor_novo
        { width: 120, format: 'TEXT' },      // ip
        { width: 150, format: 'DATETIME' }   // data_hora
      ]
    });

    formatadas += formatarSheet_(ss, CONFIG.SHEETS.USUARIOS, {
      headerColor: '#ff6f00',
      alternateColor: '#fff3e0',
      columns: [
        { width: 250, format: 'TEXT' },      // [0] email
        { width: 200, format: 'TEXT' },      // [1] nome
        { width: 100, format: 'TEXT' },      // [2] perfil
        { width: 200, format: 'TEXT' },      // [3] setores
        { width: 400, format: 'TEXT' }       // [4] permissoes (JSON)
      ]
    });

    // Formatar Logs_Debug
    formatadas += formatarSheet_(ss, CONFIG.SHEETS.LOGS_DEBUG, {
      headerColor: '#455a64',
      alternateColor: '#eceff1',
      columns: [
        { width: 150, format: 'DATETIME' },  // data_hora
        { width: 200, format: 'TEXT' },      // funcao
        { width: 200, format: 'TEXT' },      // email
        { width: 100, format: 'TEXT' },      // mensagem
        { width: 500, format: 'TEXT' }       // payload
      ]
    });

    // Formatar Limiares_SLA
    formatadas += formatarSheet_(ss, CONFIG.SHEETS.LIMIARES, {
      headerColor: '#00838f',
      alternateColor: '#e0f7fa',
      columns: [
        { width: 100, format: 'NUMBER' },    // horas_alerta
        { width: 100, format: 'NUMBER' },    // horas_critico
        { width: 150, format: 'TEXT' },      // fuso_padrao
        { width: 200, format: 'TEXT' },      // atualizado_por
        { width: 150, format: 'DATETIME' }   // atualizado_em
      ]
    });

    // Formatar Config_Geral
    formatadas += formatarSheet_(ss, CONFIG.SHEETS.CONFIG_GERAL, {
      headerColor: '#6a1b9a',
      alternateColor: '#f3e5f5',
      columns: [
        { width: 200, format: 'TEXT' },      // chave
        { width: 400, format: 'TEXT' }       // valor
      ]
    });

    // Formatar sheets de listas (simples)
    formatarSheetLista_(ss, CONFIG.SHEETS.SOLICITANTES, '#00897b');
    formatarSheetLista_(ss, CONFIG.SHEETS.SETORES, '#5e35b1');
    formatarSheetLista_(ss, CONFIG.SHEETS.ERROS_CADASTRO, '#d32f2f');
    formatarSheetLista_(ss, CONFIG.SHEETS.LISTAS, '#1565c0');

    // Remover linhas de grade de todas as abas
    ss.getSheets().forEach(sheet => {
      sheet.setHiddenGridlines(true);
    });

    ui.alert('Formatação concluída!\n\nPlanilhas formatadas: ' + formatadas);
  } catch (e) {
    ui.alert('Erro ao formatar: ' + e.toString());
  }
}

function formatarSheet_(ss, sheetName, config) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), config.columns.length);

  // Formatar cabeçalho
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground(config.headerColor);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');

  // Altura do cabeçalho
  sheet.setRowHeight(1, 30);

  // Formatar dados (sem cores alternadas para evitar timeout com muitas linhas)
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    dataRange.setFontSize(9);
    dataRange.setVerticalAlignment('middle');
    dataRange.setBackground('#ffffff');
    dataRange.setWrap(true);

    // Usar Banding para cores alternadas (muito mais eficiente que loop)
    const existingBandings = sheet.getBandings();
    existingBandings.forEach(b => b.remove());

    const fullRange = sheet.getRange(1, 1, lastRow, lastCol);
    fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

    // Ajustar cores do banding para as cores configuradas
    const bandings = sheet.getBandings();
    if (bandings.length > 0) {
      bandings[0].setHeaderRowColor(config.headerColor)
                 .setFirstRowColor('#ffffff')
                 .setSecondRowColor(config.alternateColor);
    }
  }

  // Aplicar formato de número/data e largura para cada coluna
  config.columns.forEach((col, index) => {
    if (index < lastCol) {
      // Definir largura da coluna
      sheet.setColumnWidth(index + 1, col.width || 100);

      // Aplicar formato de número/data se necessário
      if (lastRow > 1) {
        const colRange = sheet.getRange(2, index + 1, lastRow - 1, 1);
        if (col.format === 'DATETIME') {
          colRange.setNumberFormat('dd/MM/yyyy HH:mm:ss');
        } else if (col.format === 'DATE') {
          colRange.setNumberFormat('dd/MM/yyyy');
        } else if (col.format === 'CURRENCY') {
          colRange.setNumberFormat('R$ #,##0.00');
        } else if (col.format === 'NUMBER') {
          colRange.setNumberFormat('0');
          colRange.setHorizontalAlignment('center');
        }
      }
    }
  });

  // Congelar cabeçalho
  sheet.setFrozenRows(1);

  // Adicionar filtro
  if (lastRow > 1 && lastCol > 0) {
    const fullRange = sheet.getRange(1, 1, lastRow, lastCol);
    if (!sheet.getFilter()) {
      fullRange.createFilter();
    }
  }

  // Bordas
  if (lastRow > 0 && lastCol > 0) {
    const allRange = sheet.getRange(1, 1, lastRow, lastCol);
    allRange.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  }

  return 1;
}

function formatarSheetLista_(ss, sheetName, headerColor) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), 1);

  // Formatar cabeçalho
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground(headerColor);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');

  // Largura da coluna
  sheet.setColumnWidth(1, 300);

  // Dados
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    dataRange.setFontSize(10);
    dataRange.setWrap(true);

    // Cores alternadas — batch setBackgrounds() em vez de N chamadas individuais
    const bgMatrix = [];
    for (let i = 2; i <= lastRow; i++) {
      const cor = i % 2 === 0 ? '#f5f5f5' : '#ffffff';
      bgMatrix.push(Array(lastCol).fill(cor));
    }
    sheet.getRange(2, 1, lastRow - 1, lastCol).setBackgrounds(bgMatrix);
  }

  // Congelar e filtro
  sheet.setFrozenRows(1);
  if (lastRow > 1 && !sheet.getFilter()) {
    sheet.getRange(1, 1, lastRow, lastCol).createFilter();
  }

  // Bordas
  if (lastRow > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  }
}
