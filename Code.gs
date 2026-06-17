var RESERVED_SHEETS = ['Pedidos', 'Unidades', 'Precos'];

function getCatalog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var catalog = [];

  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName().trim();
    if (RESERVED_SHEETS.indexOf(name) >= 0) return;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var nomeIdx = headers.indexOf('NOME');
    if (nomeIdx === -1) return;

    var codIdx = headers.indexOf('CÓD');
    if (codIdx === -1) codIdx = headers.indexOf('COD');
    var undIdx = headers.indexOf('UND');

    for (var i = 1; i < data.length; i++) {
      var row  = data[i];
      var nome = String(row[nomeIdx] || '').trim();
      if (!nome) continue;
      catalog.push({
        cod:  codIdx >= 0 ? String(row[codIdx]  || '').trim() : '',
        nome: nome,
        und:  undIdx >= 0 && row[undIdx] ? String(row[undIdx]).trim() : 'UND',
        cat:  name
      });
    }
  });

  return catalog;
}

function getUnidades() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Unidades');

  if (!sheet) {
    sheet = ss.insertSheet('Unidades');
    sheet.appendRow(['NOME', 'VALOR_MAXIMO']);
    sheet.appendRow(['Exemplo de Filial', 5000]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
  var nomeIdx = headers.indexOf('NOME');
  var maxIdx  = headers.indexOf('VALOR_MAXIMO');

  if (nomeIdx === -1) return [];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var nome = String(data[i][nomeIdx] || '').trim();
    if (!nome) continue;
    var valorMaximo = maxIdx >= 0 ? parseFloat(data[i][maxIdx]) || 0 : 0;
    result.push({ nome: nome, valorMaximo: valorMaximo });
  }
  return result;
}

function getPrecos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Precos');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(function(h) { return String(h).trim().toUpperCase().replace('Ó','O'); });
  var codIdx   = headers.indexOf('COD');
  var nomeIdx  = headers.indexOf('NOME');
  var precoIdx = headers.indexOf('PRECO');

  if (nomeIdx === -1 || precoIdx === -1) return [];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var preco = parseFloat(String(data[i][precoIdx]).replace(',', '.'));
    if (!preco || isNaN(preco) || preco <= 0) continue;
    var cod  = codIdx >= 0 ? String(data[i][codIdx] || '').trim() : '';
    var nome = String(data[i][nomeIdx] || '').trim();
    if (!nome) continue;
    result.push({ codigo: cod, nome: nome, precoMedio: preco });
  }
  return result;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Pedido de Suprimentos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserEmail() {
  try { return Session.getActiveUser().getEmail() || ''; }
  catch(e) { return ''; }
}

function processForm(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Pedidos';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow([
      'TIMESTAMP','UNIDADE','DATA PEDIDO','DATA ENTREGA',
      'CATEGORIA','CÓD','PRODUTO','UND','QTD','SUBTOTAL','TOTAL PEDIDO','OBSERVAÇÕES','USUARIO'
    ]);
    sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  var rows = [];

  data.itens.forEach(function(item) {
    rows.push([
      ts,
      data.unidade,
      data.dataPedido,
      data.dataEntrega,
      item.cat,
      item.cod,
      item.nome,
      item.und,
      item.qtd,
      item.subtotal || '',
      data.totalPedido || '',
      data.observacoes || '',
      data.email || ''
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 13).setValues(rows);
  }

  return { ok: true, ts: ts, total: rows.length };
}
