var SPREADSHEET_ID = '1CGBTVDP3tKekeMBFAgbN3bYugOmUQ7w2vJrhWb-CEuY';

// ── Roteador ──────────────────────────────────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getAll';
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

// ── Utilitário ────────────────────────────────────────────────────────────────
function findSheet(ss, keyword) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf(keyword) > -1) return sheets[i];
  }
  return null;
}

// ── getAll: estoque + financeiro + KPIs + projeções + combos ─────────────────
function getAll() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    ok:        true,
    items:     readEstoque(ss),
    financial: readFinancial(ss),
    kpis:      readKPIs(ss),
    projecoes: readProjecoes(ss),
    combos:    readCombos(ss),
    updatedAt: new Date().toISOString()
  };
}

// ── Estoque Detalhado ─────────────────────────────────────────────────────────
// Colunas: A=Modelo B=Cor C=Tamanho D=Quantidade E=Custo Unit. F=Preço Venda
function readEstoque(ss) {
  var sheet = findSheet(ss, 'Estoque');
  if (!sheet) return [];
  var vals = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    var model = String(row[0] || '').trim();
    if (!model || model.toUpperCase().indexOf('TOTAL') > -1) continue;
    items.push({
      sheetRow: i + 1,
      model:    model,
      color:    String(row[1] || '').trim(),
      size:     String(row[2] || '').trim(),
      qty:      Number(row[3]) || 0,
      cost:     Number(row[4]) || 0,
      sale:     Number(row[5]) || 0
    });
  }
  return items;
}

function updateQty(sheetRow, qty) {
  if (isNaN(sheetRow) || isNaN(qty) || qty < 0) return { error: 'Parametros invalidos.' };
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheet(ss, 'Estoque');
  if (!sheet) return { error: 'Aba Estoque Detalhado nao encontrada.' };
  sheet.getRange(sheetRow, 4).setValue(qty); // coluna D = Quantidade
  SpreadsheetApp.flush();
  return { ok: true, sheetRow: sheetRow, qty: qty };
}

// ── Resumo Financeiro ─────────────────────────────────────────────────────────
function readFinancial(ss) {
  var sheet = findSheet(ss, 'Resumo');
  if (!sheet) return {};
  var vals = sheet.getDataRange().getValues();
  var labels = {
    'Caixa Disponivel': 'caixa', 'Caixa Disponível': 'caixa',
    'A Receber - Parcela':       'aReceberParcela',
    'A Receber - Outra Entrada': 'aReceberOutra',
    'TOTAL DISPONIVEL': 'totalDisponivel', 'TOTAL DISPONÍVEL': 'totalDisponivel',
    'Dividas a Pagar': 'dividas', 'Dívidas a Pagar': 'dividas',
    'SALDO REAL':                'saldoReal',
    'Faturamento Total':         'faturamento',
    'Custo de Producao': 'custoProducao', 'Custo de Produção': 'custoProducao',
    'Total Produzido (pecas)': 'totalProduzido', 'Total Produzido (peças)': 'totalProduzido',
    'Total Vendido (pecas)': 'totalVendido', 'Total Vendido (peças)': 'totalVendido',
    'Estoque Atual (pecas)': 'estoqueAtual', 'Estoque Atual (peças)': 'estoqueAtual',
    'Taxa de Venda':             'taxaVenda',
    'Valor de Venda (preco cheio)': 'estoqueVenda', 'Valor de Venda (preço cheio)': 'estoqueVenda',
    'Valor de Custo':            'estoqueCusto',
    'Margem Potencial':          'margemPotencial'
  };
  var result = {};
  for (var r = 0; r < vals.length; r++) {
    // Remove emojis e normaliza para fazer match nos labels
    var key = String(vals[r][0] || '').replace(/[^\x00-\xFF]/g, '').trim();
    if (labels[key] !== undefined && result[labels[key]] === undefined) {
      result[labels[key]] = vals[r][1];
    }
  }
  return result;
}

// ── KPIs Dashboard ────────────────────────────────────────────────────────────
function readKPIs(ss) {
  var sheet = findSheet(ss, 'KPI');
  if (!sheet) return {};
  var vals = sheet.getDataRange().getValues();
  var labels = {
    'Taxa de Venda Atual':       'taxaVendaAtual',
    'Margem Media Produtos': 'margemMedia', 'Margem Média Produtos': 'margemMedia',
    'Ticket Medio': 'ticketMedio', 'Ticket Médio': 'ticketMedio',
    'Giro de Estoque (dias)':    'giroEstoque',
    'ROI da Producao': 'roi', 'ROI da Produção': 'roi'
  };
  var result = {};
  for (var r = 0; r < vals.length; r++) {
    var key = String(vals[r][0] || '').replace(/[^\x00-\xFF]/g, '').trim();
    if (labels[key]) result[labels[key]] = vals[r][1];
  }
  return result;
}

// ── Projeções ─────────────────────────────────────────────────────────────────
function readProjecoes(ss) {
  var sheet = findSheet(ss, 'Proj');
  if (!sheet) return { cenarioA: {}, cenarioB: {} };
  var vals = sheet.getDataRange().getValues();
  var cA = {}, cB = {}, scenario = '';
  var ml = {
    'Budget Ads':           'budgetAds',
    'ROAS Esperado':        'roas',
    'Faturamento Projetado':'faturamento',
    'Margem Bruta (55%)':   'margemBruta',
    'Lucro Liquido': 'lucroLiquido', 'Lucro Líquido': 'lucroLiquido'
  };
  for (var r = 0; r < vals.length; r++) {
    var c = String(vals[r][0] || '');
    var cu = c.toUpperCase();
    if (cu.indexOf('CENA') > -1 && cu.indexOf(' A') > -1 && cu.indexOf(' B') < 0) { scenario = 'A'; continue; }
    if (cu.indexOf('CENA') > -1 && cu.indexOf(' B') > -1)                         { scenario = 'B'; continue; }
    var lbl = ml[c.trim()];
    if (lbl) {
      if (scenario === 'A') cA[lbl] = vals[r][1];
      if (scenario === 'B') cB[lbl] = vals[r][1];
    }
  }
  return { cenarioA: cA, cenarioB: cB };
}

// ── Custos e Margens + Combos ─────────────────────────────────────────────────
function readCombos(ss) {
  var sheet = findSheet(ss, 'Custo');
  if (!sheet) return { produtos: [], combos: [] };
  var vals = sheet.getDataRange().getValues();
  var produtos = [], combos = [], section = '';
  for (var r = 0; r < vals.length; r++) {
    var cell = String(vals[r][0] || '').trim();
    if (cell === 'TABELA DE CUSTOS E PREÇOS') { section = 'prod';  continue; }
    if (cell === 'COMBOS SUGERIDOS')           { section = 'combo'; continue; }
    if (!cell || cell === 'Produto' || cell === 'Combo') continue;
    if (section === 'prod' && vals[r][1] !== '') {
      produtos.push({
        nome:      cell,
        custo:     Number(vals[r][1]) || 0,
        venda:     Number(vals[r][2]) || 0,
        margemR:   Number(vals[r][3]) || 0,
        margemPct: parseFloat(String(vals[r][4]).replace('%', '')) || 0,
        markup:    parseFloat(String(vals[r][5]).replace('x', '')) || 0
      });
    } else if (section === 'combo' && vals[r][2] !== '') {
      combos.push({
        nome:          cell,
        composicao:    String(vals[r][1] || '').trim(),
        custo:         Number(vals[r][2]) || 0,
        precoSugerido: Number(vals[r][3]) || 0,
        desconto:      Number(vals[r][4]) || 0,
        margem:        Number(vals[r][5]) || 0
      });
    }
  }
  return { produtos: produtos, combos: combos };
}

// ── Vendas ────────────────────────────────────────────────────────────────────
// Colunas: Data | Cliente | Produto | Valor | Pagamento | Status | Observações
function getSales() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheet(ss, 'Vendas');
  if (!sheet) return { ok: true, sales: [] };
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return { ok: true, sales: [] };

  // Detecta colunas pelo cabeçalho
  var h = vals[0].map(function(c){ return String(c).trim().toLowerCase(); });
  var ci = {
    date:    h.indexOf('data'),
    buyer:   Math.max(h.indexOf('cliente'), h.indexOf('nome')),
    product: h.indexOf('produto'),
    value:   h.indexOf('valor'),
    payment: h.indexOf('pagamento'),
    status:  h.indexOf('status'),
    notes:   Math.max(h.indexOf('observacao'), h.indexOf('observações'), h.indexOf('obs'))
  };
  if (ci.date    < 0) ci.date    = 0;
  if (ci.buyer   < 0) ci.buyer   = 1;
  if (ci.product < 0) ci.product = 2;
  if (ci.value   < 0) ci.value   = 3;
  if (ci.payment < 0) ci.payment = 4;
  if (ci.status  < 0) ci.status  = 5;

  var sales = [];
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    if (!row[ci.buyer] && !row[ci.product]) continue;
    var d = row[ci.date];
    var dateStr = d instanceof Date
      ? Utilities.formatDate(d, 'America/Sao_Paulo', 'dd/MM/yyyy')
      : String(d || '');
    sales.push({
      date:    dateStr,
      buyer:   String(row[ci.buyer]   || ''),
      product: String(row[ci.product] || ''),
      value:   Number(row[ci.value])  || 0,
      status:  String(row[ci.status]  || 'Pago'),
      payment: String(row[ci.payment] || 'Pix'),
      notes:   ci.notes >= 0 ? String(row[ci.notes] || '') : ''
    });
  }
  return { ok: true, sales: sales };
}

function addSale(params) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheet(ss, 'Vendas');
  if (!sheet) return { error: 'Aba de vendas nao encontrada.' };

  var date    = params.date    ? decodeURIComponent(params.date)    : new Date().toLocaleDateString('pt-BR');
  var buyer   = params.buyer   ? decodeURIComponent(params.buyer)   : '';
  var product = params.product ? decodeURIComponent(params.product) : '';
  var value   = parseFloat(params.value)   || 0;
  var status  = params.status  ? decodeURIComponent(params.status)  : 'Pago';
  var payment = params.payment ? decodeURIComponent(params.payment) : 'Pix';
  var notes   = params.notes   ? decodeURIComponent(params.notes)   : '';

  // Formato: Data, Cliente, Produto, Valor, Pagamento, Status, Observações
  sheet.appendRow([date, buyer, product, value, payment, status, notes]);

  // Baixa estoque automaticamente se informado
  if (params.sheetRow && params.qtyToDeduct) {
    var estSheet = findSheet(ss, 'Estoque');
    if (estSheet) {
      var row    = parseInt(params.sheetRow, 10);
      var deduct = parseInt(params.qtyToDeduct, 10);
      var cur = Number(estSheet.getRange(row, 4).getValue()) || 0;
      estSheet.getRange(row, 4).setValue(Math.max(0, cur - deduct));
    }
  }

  SpreadsheetApp.flush();
  return { ok: true };
}
