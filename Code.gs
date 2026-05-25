// ============================================================
//  SISTEMA DE DEMANDAS — Google Apps Script
//  Cole este código no Apps Script do seu Google Sheets
//  e publique como Web App (ver SETUP.md)
// ============================================================

// ID da pasta no Google Drive onde os anexos serão salvos
// Deixe vazio ('') para salvar na raiz do Drive
const PASTA_DRIVE_ID = '';

// Nome da aba onde as demandas serão registradas
const NOME_ABA = 'Demandas';

// Colunas da planilha (ordem fixa)
const COLUNAS = [
  'ID',
  'Data de envio',
  'Solicitante',
  'Empresa',
  'Briefing',
  'Prazo',
  'Status',
  'Responsável',
  'Observações',
  'Link do Anexo',
  'Horas',          // col 11 — preenchida manualmente na planilha
];

// ─── Função de teste (rode pelo editor do Apps Script) ────────
//
//  Selecione "testar" no menu de funções e clique em Executar.
//  Isso cria uma demanda fictícia na planilha para confirmar
//  que tudo está funcionando antes de usar o formulário.
//
function testar() {
  const dadosFicticios = {
    solicitante: 'Teste Manual',
    empresa: 'Rhama Analysis',
    briefing: 'Demanda de teste criada pelo editor do Apps Script.',
    prazo: '2099-12-31',
    fileName: '',
    fileType: '',
    fileBase64: '',
  };

  const sheet = obterOuCriarAba();
  const id = gerarId(sheet);
  const agora = new Date();

  const linha = [
    id,
    Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    dadosFicticios.solicitante,
    dadosFicticios.empresa,
    dadosFicticios.briefing,
    dadosFicticios.prazo,
    'Pendente',
    '',
    '',
    '',
    '',  // Horas (preenchido manualmente)
  ];

  sheet.appendRow(linha);
  aplicarFormatacao(sheet);

  Logger.log('✓ Demanda de teste criada com ID: ' + id);
  Logger.log('Verifique a aba "Demandas" na sua planilha.');
}

// ─── Requisição POST (recebe demanda do formulário) ───────────

function doPost(e) {
  try {
    // Guarda: só funciona via HTTP POST, não pelo editor do Apps Script
    if (!e || !e.postData || !e.postData.contents) {
      return resposta({
        status: 'erro',
        message: 'Nenhum dado recebido. Use o formulário HTML para enviar demandas — não rode doPost() diretamente pelo editor.',
      });
    }

    const dados = JSON.parse(e.postData.contents);

    const sheet = obterOuCriarAba();
    const id = gerarId(sheet);

    let linkAnexo = '';

    // Upload do arquivo, se houver
    if (dados.fileBase64 && dados.fileName) {
      linkAnexo = salvarAnexo(dados.fileBase64, dados.fileName, dados.fileType, id);
    }

    // Linha a inserir
    const agora = new Date();
    const linha = [
      id,
      Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      dados.solicitante || '',
      dados.empresa || '',
      dados.briefing || '',
      dados.prazo || '',
      'Pendente',   // status inicial
      '',           // responsável (preenchido manualmente)
      '',           // observações (preenchido manualmente)
      linkAnexo,
      '',           // horas (preenchido manualmente)
    ];

    sheet.appendRow(linha);
    aplicarFormatacao(sheet);

    return resposta({ status: 'ok', id });

  } catch (err) {
    console.error('Erro no doPost:', err);
    return resposta({ status: 'erro', message: err.message });
  }
}

// ─── Requisição GET (fornece dados para o dashboard) ─────────

function doGet(e) {
  try {
    const sheet = obterOuCriarAba();
    const dados = lerDemandas(sheet);
    return resposta({ status: 'ok', dados });
  } catch (err) {
    return resposta({ status: 'erro', message: err.message });
  }
}

// ─── Funções auxiliares ───────────────────────────────────────

function obterOuCriarAba() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NOME_ABA);

  if (!sheet) {
    sheet = ss.insertSheet(NOME_ABA);
    // Cabeçalho
    sheet.appendRow(COLUNAS);

    // Estilo do cabeçalho
    const headerRange = sheet.getRange(1, 1, 1, COLUNAS.length);
    headerRange.setBackground('#1a56db');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    sheet.setFrozenRows(1);

    // Larguras das colunas
    sheet.setColumnWidth(1, 80);   // ID
    sheet.setColumnWidth(2, 140);  // Data
    sheet.setColumnWidth(3, 160);  // Solicitante
    sheet.setColumnWidth(4, 140);  // Empresa
    sheet.setColumnWidth(5, 300);  // Briefing
    sheet.setColumnWidth(6, 110);  // Prazo
    sheet.setColumnWidth(7, 110);  // Status
    sheet.setColumnWidth(8, 150);  // Responsável
    sheet.setColumnWidth(9, 200);  // Observações
    sheet.setColumnWidth(10, 200); // Link Anexo
    sheet.setColumnWidth(11, 80);  // Horas

    // Validação de status na coluna G (a partir da linha 2)
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendente', 'Em andamento', 'Concluído', 'Cancelado'], true)
      .build();
    sheet.getRange('G2:G1000').setDataValidation(statusRule);
  }

  return sheet;
}

function gerarId(sheet) {
  const ultimaLinha = sheet.getLastRow();
  const num = ultimaLinha; // linha 1 = cabeçalho, linha 2 = demanda #1
  return 'DEM-' + String(num).padStart(4, '0');
}

function salvarAnexo(base64, nome, tipo, id) {
  try {
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, tipo || 'application/octet-stream', id + '_' + nome);

    let pasta;
    if (PASTA_DRIVE_ID) {
      pasta = DriveApp.getFolderById(PASTA_DRIVE_ID);
    } else {
      // Cria (ou reutiliza) pasta "Anexos Demandas" na raiz
      const pastas = DriveApp.getFoldersByName('Anexos Demandas');
      pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder('Anexos Demandas');
    }

    const arquivo = pasta.createFile(blob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return arquivo.getUrl();
  } catch (err) {
    console.error('Erro ao salvar anexo:', err);
    return 'Erro ao salvar anexo';
  }
}

function lerDemandas(sheet) {
  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return [];

  const range = sheet.getRange(2, 1, ultimaLinha - 1, COLUNAS.length);
  const valores = range.getValues();

  return valores.map(row => ({
    id:          row[0],
    dataEnvio:   row[1],
    solicitante: row[2],
    empresa:     row[3],
    briefing:    row[4],
    prazo:       row[5],
    status:      row[6],
    responsavel: row[7],
    observacoes: row[8],
    linkAnexo:   row[9],
    horas:       parseFloat(row[10]) || 0,  // número de horas trabalhadas (0 se vazio)
  }));
}

function aplicarFormatacao(sheet) {
  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return;

  // Colorir célula de status de acordo com o valor
  const statusRange = sheet.getRange(ultimaLinha, 7);
  const status = statusRange.getValue();

  const cores = {
    'Pendente':      { bg: '#fff3cd', fg: '#856404' },
    'Em andamento':  { bg: '#cce5ff', fg: '#004085' },
    'Concluído':     { bg: '#d4edda', fg: '#155724' },
    'Cancelado':     { bg: '#f8d7da', fg: '#721c24' },
  };

  if (cores[status]) {
    statusRange.setBackground(cores[status].bg);
    statusRange.setFontColor(cores[status].fg);
    statusRange.setFontWeight('bold');
  }

  // Zebra nas linhas
  const rowRange = sheet.getRange(ultimaLinha, 1, 1, COLUNAS.length);
  if (ultimaLinha % 2 === 0) {
    rowRange.setBackground('#f8faff');
  }
}

// Resposta JSON com CORS habilitado
function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
