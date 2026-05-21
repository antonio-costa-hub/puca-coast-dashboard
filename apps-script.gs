var SPREADSHEET_ID = '1CGBTVDP3tKekeMBFAgbN3bYugOmUQ7w2vJrhWb-CEuY';

// ── Roteador ──────────────────────────────────────────────────────────────────
function doGet(e) {
  var action = 'getAll';
  if (e && e.parameter && e.parameter.action) {
    action = e.parameter.action;
  }

  var result;
  try {
    if      (action === 'getAll')    result = getAll();
    else if (action === 'updateQty') result = updateQty(parseInt(e.parameter.sheetRow, 10), parseInt(e.parameter.qty, 10));
    else if (action === 'getSales')  result = getSales();
    else if (action === 'addSale')   result = addSale(e.parameter);
    else                             result = { error: 'Acao desconhecida: ' + action };
  } catch (err) {
    result = { error: String(err) };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Estoque + financeiro ──────────────────────────────────────────────────────
function getAll() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var targetSheet = null;
  var headerRow = -1;
  var allVals = null;

  for (var s = 0; s < sheets.length; s++) {
    var vals = sheets[s].getDataRange().getValues();
    for (var r = 0; r < vals.length; r++) {
      if (String(vals[r][0]).trim() === 'Modelo' && String(vals[r][1]).trim() === 'Cor') {
        targetSheet = sheets[s];
        headerRow = r;
        allVals = vals;
        break;
      }
    }
    if (targetSheet) break;
  }

  if (!targetSheet) return { error: 'Tabela de estoque nao encontrada.' };

  var items = [];
  for (var i = headerRow + 1; i < allVals.length; i++) {
    var row = allVals[i];
    if (!row[0] || String(row[0]).trim() === '') break;
    items.push({
      sheetRow: i + 1,
      model: String(row[0]).trim(),
      color: String(row[1]).trim(),
      size:  String(row[2]).trim(),
      qty:   Number(row[3]) || 0,
      cost:  Number(row[4]) || 0,
      sale:  Number(row[5]) || 0
    });
  }

  return { ok: true, items: items, financial: readFinancial(ss), updatedAt: new Date().toISOString() };
}

function updateQty(sheetRow, qty) {
  if (isNaN(sheetRow) || isNaN(qty) || qty < 0) return { error: 'Parametros invalidos.' };
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var info = findStockSheet(ss);
  if (!info) return { error: 'Aba de estoque nao encontrada.' };
  info.sheet.getRange(sheetRow, 4).setValue(qty);
  SpreadsheetApp.flush();
  return { ok: true, sheetRow: sheetRow, qty: qty };
}

function findStockSheet(ss) {
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var vals = sheets[s].getDataRange().getValues();
    for (var r = 0; r < vals.length; r++) {
      if (String(vals[r][0]).trim() === 'Modelo' && String(vals[r][1]).trim() === 'Cor') {
        return { sheet: sheets[s], headerRow: r, vals: vals };
      }
    }
  }
  return null;
}

function readFinancial(ss) {
  var labels = {
    'Caixa Disponivel': 'caixa', 'Caixa Disponível': 'caixa',
    'A Receber - Parcela': 'aReceberParcela',
    'A Receber - Outra Entrada': 'aReceberOutra',
    'TOTAL DISPONIVEL': 'totalDisponivel', 'TOTAL DISPONÍVEL': 'totalDisponivel',
    'Dividas a Pagar': 'dividas', 'Dívidas a Pagar': 'dividas',
    'SALDO REAL': 'saldoReal',
    'Faturamento Total': 'faturamento',
    'Custo de Producao': 'custoProducao', 'Custo de Produção': 'custoProducao',
    'Total Produzido (pecas)': 'totalProduzido', 'Total Produzido (peças)': 'totalProduzido',
    'Total Vendido (pecas)': 'totalVendido', 'Total Vendido (peças)': 'totalVendido',
    'Taxa de Venda': 'taxaVenda'
  };
  var result = {};
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var vals = sheets[s].getDataRange().getValues();
    for (var r = 0; r < vals.length; r++) {
      var key = String(vals[r][0]).trim();
      if (labels[key] !== undefined && result[labels[key]] === undefined) {
        result[labels[key]] = vals[r][1];
      }
    }
  }
  return result;
}

// ── Vendas ────────────────────────────────────────────────────────────────────
function getSales() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSalesSheet(ss);
  var vals = sheet.getDataRange().getValues();
  if (vals.length <= 1) return { ok: true, sales: [] };

  var sales = [];
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    if (!row[0] && !row[1]) continue;
    sales.push({
      id:      i + 1,
      date:    row[0] ? String(row[0]) : '',
      buyer:   String(row[1] || ''),
      product: String(row[2] || ''),
      value:   Number(row[3]) || 0,
      payment: String(row[4] || 'Pix'),
      notes:   String(row[5] || '')
    });
  }
  return { ok: true, sales: sales };
}

function addSale(params) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSalesSheet(ss);

  var date    = params.date    ? decodeURIComponent(params.date)    : new Date().toLocaleDateString('pt-BR');
  var buyer   = params.buyer   ? decodeURIComponent(params.buyer)   : '';
  var product = params.product ? decodeURIComponent(params.product) : '';
  var value   = parseFloat(params.value)   || 0;
  var payment = params.payment ? decodeURIComponent(params.payment) : 'Pix';
  var notes   = params.notes   ? decodeURIComponent(params.notes)   : '';

  sheet.appendRow([date, buyer, product, value, payment, notes]);

  // Baixa o estoque automaticamente se informado
  if (params.sheetRow && params.qtyToDeduct) {
    var info = findStockSheet(ss);
    if (info) {
      var row   = parseInt(params.sheetRow, 10);
      var deduct = parseInt(params.qtyToDeduct, 10);
      var cur   = Number(info.sheet.getRange(row, 4).getValue()) || 0;
      info.sheet.getRange(row, 4).setValue(Math.max(0, cur - deduct));
    }
  }

  SpreadsheetApp.flush();
  return { ok: true };
}

function getOrCreateSalesSheet(ss) {
  var sheet = ss.getSheetByName('Vendas');
  if (!sheet) {
    sheet = ss.insertSheet('Vendas');
    var headers = ['Data', 'Comprador', 'Produto', 'Valor (R$)', 'Pagamento', 'Observação'];
    sheet.appendRow(headers);
    var range = sheet.getRange(1, 1, 1, headers.length);
    range.setFontWeight('bold');
    range.setBackground('#0e1628');
    range.setFontColor('#00c2ff');
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 260);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 220);
  }
  return sheet;
}

// ── RODE ESTA FUNÇÃO UMA VEZ para importar o histórico de vendas ──────────────
// Apps Script > selecione "initHistoricalSales" no menu > clique ▶ Executar
function initHistoricalSales() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSalesSheet(ss);

  if (sheet.getLastRow() > 1) {
    Logger.log('A aba Vendas ja tem dados. Apague as linhas e tente novamente.');
    return;
  }

  // [Data, Comprador, Produto, Valor, Pagamento, Observacao]
  var vendas = [
    ['01/01/2026', 'Marco',         'PDV Azul G',                             100, 'Pix', ''],
    ['01/01/2026', 'Antonio',       'Manifesto Preta GG',                     100, 'Pix', ''],
    ['01/01/2026', 'Toinho',        'Short GG',                               210, 'Pix', ''],
    ['01/01/2026', 'Clara Lage',    'PDV Branca GG + Manifesto Preta G + Short GG', 290, 'Pix', ''],
    ['01/01/2026', 'Yan',           'Manifesto Preta G',                      100, 'Pix', ''],
    ['01/01/2026', 'Stefano',       'Short GG',                               120, 'Pix', ''],
    ['01/01/2026', 'Davi Loiro',    'PDV Azul G',                             100, 'Pix', ''],
    ['01/01/2026', 'Davi Loiro',    'PDV Branca GG + Manifesto Preta G + Short GG', 320, 'Pix', 'Segundo pedido'],
    ['01/01/2026', 'Heitor',        'PDV Azul G + Short M',                   220, 'Pix', ''],
    ['01/01/2026', 'LaRRso',        'PDV Azul G + PDV Branca GG',             200, 'Pix', ''],
    ['01/01/2026', 'Gabi Prima',    'PDV Branca M + 2x Bonés',                270, 'Pix', ''],
    ['01/01/2026', 'Taissa',        'PDV Azul M + Boné',                      185, 'Pix', ''],
    ['01/01/2026', 'Tio Ricardo',   'Manifesto Verde GG',                     100, 'Pix', ''],
    ['01/01/2026', 'Jailton',       'Manifesto Preta G',                      100, 'Pix', ''],
    ['01/01/2026', 'Victor',        'Manifesto Preta G',                      105, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Cadinho',       'Manifesto Verde G',                      105, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Camila',        'Boné',                                    85, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Guga',          'PDV Azul M',                             110, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Tom',           'Boné + Manifesto Preta G + PDV Azul G',  300, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Antonio (Site)','PDV Azul G',                             110, 'Pix', 'Pré-Venda | Brinde: Boné'],
    ['01/01/2026', 'Ian (Site)',     'PDV Branca M + Boné',                   195, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Alex',          'Manifesto Preta G',                      105, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Kael',          'Manifesto Preta GG + Verde GG + PDV Azul GG + Boné Combo', 300, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Icaro',         'PDV Branca M + Azul M + Verde M + Short G + PDV Azul GG', 300, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Cadu',          'Short G + PDV Azul GG',                  230, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'JG',            'Boné',                                    85, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Diogo',         'Boné',                                    85, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Alan',          'Manifesto Preta M',                      115, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Jr',            'PDV Azul G',                             125, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Jorge',         'Manifesto Verde M',                      115, 'Pix', 'Pré-Venda'],
    ['01/01/2026', 'Ismael',        'PDV Branca P + Boné + Short',            300, 'Pix', 'Pré-Venda'],
  ];

  for (var i = 0; i < vendas.length; i++) {
    sheet.appendRow(vendas[i]);
  }

  // Formata valores como moeda
  sheet.getRange(2, 4, vendas.length, 1).setNumberFormat('R$ #,##0.00');

  SpreadsheetApp.flush();
  Logger.log('Pronto! ' + vendas.length + ' vendas importadas.');
}
