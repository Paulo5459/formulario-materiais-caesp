/**
 * MATERIAIS CAESP/2026 - recebe os dados do formulário web e grava uma linha na planilha.
 *
 * Implantação: veja o passo a passo no chat / LEIA-ME.md.
 * Depois de implantar, copie a URL do Web App (termina em /exec) para ENDPOINT_URL no index.html.
 */

// ID da planilha "Levantamento CAESP 2026" (trecho entre /d/ e /edit na URL da planilha).
var SHEET_ID = '1TX0c0C5trW-dFZ8B3Eke_euca_CeBMBZj_Zc5Ul3Oo0';

var MAX_TEXTO = 200;

var HEADERS = [
  'Data/Hora',
  'Nome Completo',
  'Instituição',
  'Quantidade de Moeda',
  'Quantidade de Brevê Metálico',
  'Quantidade de Totem'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'error', message: 'Corpo da requisição vazio.' });
    }

    var data = JSON.parse(e.postData.contents);

    var nome = cleanText(data.nomeCompleto);
    var instituicao = cleanText(data.instituicao);
    var moeda = toNonNegativeInt(data.quantidadeMoeda);
    var breve = toNonNegativeInt(data.quantidadeBreveMetalico);
    var totem = toNonNegativeInt(data.quantidadeTotem);

    if (!nome || !instituicao || moeda === null || breve === null || totem === null) {
      return jsonResponse({ status: 'error', message: 'Dados inválidos ou incompletos.' });
    }

    getSheet().appendRow([new Date(), nome, instituicao, moeda, breve, totem]);

    return jsonResponse({ status: 'ok' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  } finally {
    lock.releaseLock();
  }
}

// Permite conferir pelo navegador se a implantação está ativa.
function doGet() {
  return jsonResponse({ status: 'ok', message: 'Endpoint de Materiais CAESP/2026 ativo.' });
}

function getSheet() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

  var firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (firstRow.join('') === '') {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

// Limita o tamanho e neutraliza fórmulas (=, +, -, @) para que o texto não seja executado na planilha.
function cleanText(value) {
  var text = (value === null || value === undefined ? '' : String(value)).trim().slice(0, MAX_TEXTO);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function toNonNegativeInt(value) {
  var n = Number(value);
  if (value === '' || value === null || isNaN(n) || !isFinite(n)) return null;
  if (Math.floor(n) !== n || n < 0) return null;
  return n;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
