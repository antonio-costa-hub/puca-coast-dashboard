/**
 * PUCA COAST · DROP 4 — Backend (Google Apps Script)
 * ----------------------------------------------------
 * Conecta o dashboard à planilha. Lê o estoque (tabela de SKUs) e o registro
 * de vendas, e grava de volta quando você edita no dashboard.
 *
 * COMO PUBLICAR:
 *  1. Abra a planilha → Extensões > Apps Script
 *  2. Apague tudo, cole este arquivo inteiro e salve (Ctrl+S)
 *  3. Implantar > Nova implantação > Tipo: App da Web
 *  4. Executar como: Eu  ·  Acesso: Qualquer pessoa
 *  5. Implantar > Autorizar > copie a URL (.../exec) e cole no dashboard
 */

var SPREADSHEET_ID = '1t4dkMjd5ByYFV1hAScm3PkShlj66iSU-ZIFlWGbpcFY';

// Cabeçalhos que identificam cada tabela dentro da planilha
var STOCK_KEYS = ['sku', 'produto'];          // linha com "SKU" + "Produto"
var SALES_HEADER = ['data', 'item'];          // linha com "Data" + "Item (SKU)"

// ── Roteador ────────────────────────────────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'getAll';
  var result;
  try {
    if      (action === 'getAll')      result = getAll();
    else if (action === 'getSales')    result = getSales();
    else if (action === 'setProduced') result = setProduced(int(e.parameter.sheetRow), int(e.parameter.qty));
    else if (action === 'setSold')     result = setSold(int(e.parameter.sheetRow), int(e.parameter.qty));
    else if (action === 'addSaleLine') result = addSaleLine(e.parameter);
    else                               result = { error: 'Ação desconhecida: ' + action };
  } catch (err) {
    result = { error: String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function int(v) { return parseInt(v, 10); }
function norm(v) { return String(v == null ? '' : v).trim().toLowerCase(); }

// ── Localiza a tabela de estoque (SKU / Produto / Cor / Tamanho / ...) ───────
function findStockSheet(ss) {
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var vals = sheets[s].getDataRange().getValues();
    for (var r = 0; r < vals.length; r++) {
      if (norm(vals[r][0]) === STOCK_KEYS[0] && norm(vals[r][1]) === STOCK_KEYS[1]) {
        return { sheet: sheets[s], headerRow: r, vals: vals };
      }
    }
  }
  return null;
}

// Colunas (1-indexadas) na tabela de SKUs
// A SKU · B Produto · C Cor · D Tamanho · E Custo · F Preço · G Produzida · H Vendida · I Estoque
var COL = { sku:1, product:2, color:3, size:4, cost:5, sale:6, produced:7, sold:8, stock:9 };

// ── Lê tudo: itens de estoque + vendas ──────────────────────────────────────
function getAll() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var info = findStockSheet(ss);
  if (!info) return { error: 'Tabela de estoque (SKU/Produto) não encontrada.' };

  var items = [];
  for (var i = info.headerRow + 1; i < info.vals.length; i++) {
    var row = info.vals[i];
    var sku = String(row[0]).trim();
    if (!sku) break;                                  // primeira linha vazia = fim da tabela
    if (/^total/i.test(sku)) break;
    items.push({
      sheetRow: i + 1,
      sku:      sku,
      product:  String(row[1]).trim(),
      color:    String(row[2]).trim(),
      size:     String(row[3]).trim(),
      cost:     Number(row[4]) || 0,
      sale:     Number(row[5]) || 0,
      produced: Number(row[6]) || 0,
      sold:     Number(row[7]) || 0
    });
  }

  return {
    ok: true,
    items: items,
    sales: readSales(ss).sales,
    updatedAt: new Date().toISOString()
  };
}

// ── Estoque: grava quantidade produzida ─────────────────────────────────────
function setProduced(sheetRow, qty) {
  if (isNaN(sheetRow) || isNaN(qty) || qty < 0) return { error: 'Parâmetros inválidos.' };
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var info = findStockSheet(ss);
  if (!info) return { error: 'Aba de estoque não encontrada.' };
  info.sheet.getRange(sheetRow, COL.produced).setValue(qty);
  syncStockCell(info.sheet, sheetRow);
  SpreadsheetApp.flush();
  return { ok: true, sheetRow: sheetRow, produced: qty };
}

// ── Estoque: grava quantidade vendida ───────────────────────────────────────
function setSold(sheetRow, qty) {
  if (isNaN(sheetRow) || isNaN(qty) || qty < 0) return { error: 'Parâmetros inválidos.' };
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var info = findStockSheet(ss);
  if (!info) return { error: 'Aba de estoque não encontrada.' };
  var cell = info.sheet.getRange(sheetRow, COL.sold);
  // Se "Qtd vendida" for fórmula (ex.: SOMASE do registro), não sobrescreve.
  if (cell.getFormula()) {
    return { ok: true, sheetRow: sheetRow, sold: Number(cell.getValue()) || 0, formula: true };
  }
  cell.setValue(qty);
  syncStockCell(info.sheet, sheetRow);
  SpreadsheetApp.flush();
  return { ok: true, sheetRow: sheetRow, sold: qty };
}

// Se "Estoque atual" NÃO for fórmula, mantém = produzida − vendida
function syncStockCell(sheet, sheetRow) {
  var cell = sheet.getRange(sheetRow, COL.stock);
  if (cell.getFormula()) return;                      // é fórmula → não toca
  var prod = Number(sheet.getRange(sheetRow, COL.produced).getValue()) || 0;
  var sold = Number(sheet.getRange(sheetRow, COL.sold).getValue()) || 0;
  cell.setValue(Math.max(0, prod - sold));
}

// ── Vendas ──────────────────────────────────────────────────────────────────
function getSales() {
  return readSales(SpreadsheetApp.openById(SPREADSHEET_ID));
}

function findSalesSheet(ss) {
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var vals = sheets[s].getDataRange().getValues();
    for (var r = 0; r < vals.length; r++) {
      if (norm(vals[r][0]) === 'data' && norm(vals[r][1]).indexOf('item') === 0) {
        return { sheet: sheets[s], headerRow: r, vals: vals };
      }
    }
  }
  // Não achou: cria uma aba dedicada
  var sheet = ss.getSheetByName('Registro de Vendas');
  if (!sheet) {
    sheet = ss.insertSheet('Registro de Vendas');
    sheet.appendRow(['Data', 'Item (SKU)', 'Qtd', 'Preço unit.', 'Total', 'Canal / Cliente', 'Obs.']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return { sheet: sheet, headerRow: 0, vals: sheet.getDataRange().getValues() };
}

function readSales(ss) {
  var info = findSalesSheet(ss);
  var vals = info.sheet.getDataRange().getValues();
  var sales = [];
  for (var i = info.headerRow + 1; i < vals.length; i++) {
    var row = vals[i];
    if (!row[0] && !row[1] && !row[5]) continue;       // linha vazia
    sales.push({
      id:      i + 1,
      date:    fmtDate(row[0]),
      sku:     String(row[1] || ''),
      qty:     Number(row[2]) || 0,
      unit:    Number(row[3]) || 0,
      total:   Number(row[4]) || 0,
      client:  String(row[5] || ''),
      notes:   String(row[6] || '')
    });
  }
  return { ok: true, sales: sales };
}

// Grava UMA linha de venda (o dashboard chama uma vez por produto do pedido)
function addSaleLine(p) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var info = findSalesSheet(ss);

  var date   = p.date   ? decodeURIComponent(p.date)   : Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy');
  var sku    = p.sku    ? decodeURIComponent(p.sku)    : '';
  var qty    = Number(p.qty)  || 1;
  var unit   = Number(p.unit) || 0;
  var total  = qty * unit;
  var client = p.client ? decodeURIComponent(p.client) : '';
  var notes  = p.notes  ? decodeURIComponent(p.notes)  : '';

  // acha a primeira linha vazia abaixo do cabeçalho do registro
  var sheet = info.sheet;
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 7).setValues([[date, sku, qty, unit, total, client, notes]]);

  // baixa o estoque: soma à quantidade vendida do SKU
  // (se "Qtd vendida" for fórmula, ela já se atualiza sozinha com a linha acima)
  if (p.skuRow) {
    var sr = int(p.skuRow);
    var st = findStockSheet(ss);
    if (st && sr > 0) {
      var soldCell = st.sheet.getRange(sr, COL.sold);
      if (!soldCell.getFormula()) {
        var cur = Number(soldCell.getValue()) || 0;
        soldCell.setValue(cur + qty);
        syncStockCell(st.sheet, sr);
      }
    }
  }

  SpreadsheetApp.flush();
  return { ok: true };
}

function fmtDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'GMT-3', 'dd/MM/yyyy');
  return String(v || '');
}
