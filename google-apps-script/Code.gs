/**
 * Planilha · API de sincronização com Google Planilhas (versão 3: clientes, pagamentos e produtos)
 *
 * Como usar:
 * 1. Crie uma planilha no Google Planilhas.
 * 2. Menu Extensões → Apps Script. Apague TODO o conteúdo e cole este arquivo inteiro.
 * 3. Troque o valor de CHAVE abaixo por uma senha só sua.
 * 4. Implantar → Nova implantação → tipo "App da Web"
 *    - Executar como: Eu
 *    - Quem pode acessar: Qualquer pessoa
 * 5. Copie a URL que termina em /exec e cole no site (aba Financeiro → Google Planilhas),
 *    junto com a mesma CHAVE.
 *
 * Para atualizar um script que já existe: cole este arquivo no lugar do antigo, salve e vá em
 * Implantar → Gerenciar implantações → lápis → Versão: "Nova versão" → Implantar. A URL não muda.
 */

const CHAVE = 'COLE_AQUI_SUA_CHAVE';

const VERSAO = 3;
const ABA_VENDAS = 'Vendas';
const ABA_CLIENTES = 'Clientes';
const ABA_PRODUTOS = 'Produtos';
const ABA_CONFIG = 'Config';
const COLUNAS = [
  'ID', 'Data', 'Peça', 'Material', 'Peso (g)', 'Tempo (h)', 'Quantidade',
  'Custo unit.', 'Preço unit.', 'Taxas unit.', 'Lucro unit.',
  'Custo total', 'Faturamento', 'Lucro líquido', 'Criado em',
  'Cliente', 'Status', 'Pago em', 'Cliente ID', 'Taxa %', 'Taxa fixa', 'Produto ID',
];
const COLUNAS_CLIENTES = ['ID', 'Nome', 'WhatsApp', 'Criado em'];
const COLUNAS_PRODUTOS = [
  'ID', 'Produto', 'Material', 'Peso total (g)', 'Tempo total (h)', 'Peças por impressão', 'Seu tempo (min)',
  'Margem %', 'Taxa %', 'Taxa fixa', 'Custo por peça', 'Preço por peça', 'Criado em', 'Atualizado em',
];

/* ------------------------------------------------------------------ */

function doGet() {
  return saida({ ok: true, app: 'Planilha', versao: VERSAO, mensagem: 'API funcionando. Use o site para acessar os dados.' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return saida({ ok: false, erro: 'Requisição inválida.' });
  }
  if (!req || req.key !== CHAVE) {
    return saida({ ok: false, erro: 'Chave de acesso inválida.' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (req.action) {
      case 'ping':
        return saida({ ok: true, versao: VERSAO });
      case 'load':
        return saida({ ok: true, versao: VERSAO, sales: lerVendas(), clients: lerClientes(), products: lerProdutos(), investimento: lerInvestimento() });
      case 'add':
        adicionarVenda(req.sale);
        return saida({ ok: true });
      case 'update':
        atualizarVenda(req.sale);
        return saida({ ok: true });
      case 'delete':
        excluirVenda(req.id);
        return saida({ ok: true });
      case 'clear':
        limparVendas();
        return saida({ ok: true });
      case 'setInvest':
        salvarInvestimento(req.value);
        return saida({ ok: true });
      case 'saveClient':
        salvarCliente(req.client);
        return saida({ ok: true });
      case 'deleteClient':
        excluirCliente(req.id);
        return saida({ ok: true });
      case 'saveProduct':
        salvarProduto(req.product);
        return saida({ ok: true });
      case 'deleteProduct':
        excluirProduto(req.id);
        return saida({ ok: true });
      case 'replaceAll':
        substituirTudo(req.sales || [], req.investimento, req.clients, req.products);
        return saida({ ok: true });
      default:
        return saida({ ok: false, erro: 'Ação desconhecida.' });
    }
  } catch (err) {
    return saida({ ok: false, erro: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */

function saida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function cabecalho(sh, colunas, cor) {
  sh.getRange(1, 1, 1, colunas.length).setValues([colunas])
    .setFontWeight('bold').setBackground(cor).setFontColor('#000000');
  sh.setFrozenRows(1);
}

function abaVendas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_VENDAS);
  if (!sh) {
    sh = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0
      ? ss.getSheets()[0].setName(ABA_VENDAS)
      : ss.insertSheet(ABA_VENDAS);
  }
  if (sh.getLastRow() === 0 || sh.getLastColumn() < COLUNAS.length) {
    // planilha nova, ou criada pela versão 1 (sem as colunas de cliente/pagamento)
    cabecalho(sh, COLUNAS, '#c6f432');
    sh.getRange('A2:A').setNumberFormat('@');
    sh.getRange('B2:B').setNumberFormat('dd/mm/yyyy');
    sh.getRange('E2:E').setNumberFormat('0');
    sh.getRange('F2:F').setNumberFormat('0.0');
    sh.getRange('H2:N').setNumberFormat('"R$" #,##0.00');
    sh.getRange('O2:O').setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange('R2:R').setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange('S2:S').setNumberFormat('@');
    sh.getRange('V2:V').setNumberFormat('@');
    sh.setColumnWidth(3, 220);
    sh.setColumnWidth(16, 160);
  }
  return sh;
}

function abaClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_CLIENTES);
  if (!sh) {
    sh = ss.insertSheet(ABA_CLIENTES);
    cabecalho(sh, COLUNAS_CLIENTES, '#86efac');
    sh.getRange('A2:A').setNumberFormat('@');
    sh.getRange('C2:C').setNumberFormat('@');
    sh.getRange('D2:D').setNumberFormat('dd/mm/yyyy hh:mm');
    sh.setColumnWidth(2, 220);
    sh.setColumnWidth(3, 150);
  }
  return sh;
}

function abaConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_CONFIG);
  if (!sh) {
    sh = ss.insertSheet(ABA_CONFIG);
    sh.getRange('A1').setValue('Investimento na impressora').setFontWeight('bold');
    sh.getRange('B1').setNumberFormat('"R$" #,##0.00');
    sh.setColumnWidth(1, 220);
  }
  return sh;
}

function num(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function paraData(iso) {
  if (!iso) return '';
  const p = String(iso).split('-').map(Number);
  if (p.length !== 3 || p.some(isNaN)) return '';
  return new Date(p[0], p[1] - 1, p[2]);
}

function pagoDe(s) {
  return !(s.pago === false || s.pago === 'false' || s.pago === 0);
}

function linhaDaVenda(s) {
  const qtd = Math.max(1, num(s.qtd));
  const pago = pagoDe(s);
  return [
    String(s.id),
    paraData(s.data),
    String(s.nome || ''),
    String(s.material || ''),
    num(s.peso),
    num(s.horas),
    qtd,
    num(s.custoUnit),
    num(s.precoUnit),
    num(s.taxasUnit),
    num(s.lucroUnit),
    num(s.custoUnit) * qtd,
    num(s.precoUnit) * qtd,
    num(s.lucroUnit) * qtd,
    s.criadoEm ? new Date(num(s.criadoEm)) : new Date(),
    String(s.clienteNome || ''),
    pago ? 'Pago' : 'Pendente',
    pago && s.pagoEm ? new Date(num(s.pagoEm)) : '',
    String(s.clienteId || ''),
    s.taxaPct === undefined || s.taxaPct === null ? '' : num(s.taxaPct),
    s.taxaFixa === undefined || s.taxaFixa === null ? '' : num(s.taxaFixa),
    String(s.produtoId || ''),
  ];
}

function lerVendas() {
  const sh = abaVendas();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const tz = Session.getScriptTimeZone();
  const valores = sh.getRange(2, 1, last - 1, COLUNAS.length).getValues();
  return valores
    .filter((r) => r[0] !== '' && r[0] !== null)
    .map((r) => {
      const v = {
        id: String(r[0]),
        data: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1] || ''),
        nome: String(r[2] || ''),
        material: String(r[3] || ''),
        peso: num(r[4]),
        horas: num(r[5]),
        qtd: Math.max(1, num(r[6])),
        custoUnit: num(r[7]),
        precoUnit: num(r[8]),
        taxasUnit: num(r[9]),
        lucroUnit: num(r[10]),
        criadoEm: r[14] instanceof Date ? r[14].getTime() : 0,
        clienteNome: String(r[15] || ''),
        pago: String(r[16] || '').toLowerCase() !== 'pendente', // linhas antigas (vazias) contam como pagas
        pagoEm: r[17] instanceof Date ? r[17].getTime() : 0,
        clienteId: String(r[18] || ''),
      };
      if (r[19] !== '' && r[19] !== null) v.taxaPct = num(r[19]);
      if (r[20] !== '' && r[20] !== null) v.taxaFixa = num(r[20]);
      v.produtoId = String(r[21] || '');
      return v;
    });
}

function linhaDoId(sh, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function adicionarVenda(s) {
  if (!s || !s.id) throw new Error('Venda sem ID.');
  const sh = abaVendas();
  if (linhaDoId(sh, s.id) !== -1) return atualizarVenda(s); // já existe (envio repetido)
  sh.appendRow(linhaDaVenda(s));
}

function atualizarVenda(s) {
  if (!s || !s.id) throw new Error('Venda sem ID.');
  const sh = abaVendas();
  const row = linhaDoId(sh, s.id);
  if (row === -1) { sh.appendRow(linhaDaVenda(s)); return; }
  sh.getRange(row, 1, 1, COLUNAS.length).setValues([linhaDaVenda(s)]);
}

function excluirVenda(id) {
  const sh = abaVendas();
  const row = linhaDoId(sh, id);
  if (row !== -1) sh.deleteRow(row);
}

function limparVendas() {
  const sh = abaVendas();
  const last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
}

/* ---------------------------- Clientes ----------------------------- */

function lerClientes() {
  const sh = abaClientes();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, COLUNAS_CLIENTES.length).getValues()
    .filter((r) => r[0] !== '' && r[0] !== null)
    .map((r) => ({
      id: String(r[0]),
      nome: String(r[1] || ''),
      telefone: String(r[2] || '').replace(/\D/g, ''),
      criadoEm: r[3] instanceof Date ? r[3].getTime() : 0,
    }));
}

function linhaDoCliente(c) {
  return [String(c.id), String(c.nome || ''), String(c.telefone || ''), c.criadoEm ? new Date(num(c.criadoEm)) : new Date()];
}

function salvarCliente(c) {
  if (!c || !c.id) throw new Error('Cliente sem ID.');
  const sh = abaClientes();
  const row = linhaDoId(sh, c.id);
  if (row === -1) sh.appendRow(linhaDoCliente(c));
  else sh.getRange(row, 1, 1, COLUNAS_CLIENTES.length).setValues([linhaDoCliente(c)]);
}

function excluirCliente(id) {
  const sh = abaClientes();
  const row = linhaDoId(sh, id);
  if (row !== -1) sh.deleteRow(row);
}

/* ---------------------------- Produtos ----------------------------- */

function abaProdutos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_PRODUTOS);
  if (!sh) {
    sh = ss.insertSheet(ABA_PRODUTOS);
    cabecalho(sh, COLUNAS_PRODUTOS, '#fde68a');
    sh.getRange('A2:A').setNumberFormat('@');
    sh.getRange('K2:L').setNumberFormat('"R$" #,##0.00');
    sh.getRange('M2:N').setNumberFormat('dd/mm/yyyy hh:mm');
    sh.setColumnWidth(2, 220);
  }
  return sh;
}

function linhaDoProduto(p) {
  return [
    String(p.id), String(p.nome || ''), String(p.material || ''),
    num(p.pesoTotal), num(p.horasTotal), Math.max(1, num(p.pecas)), num(p.minutosMO),
    num(p.margem), num(p.taxaPct), num(p.taxaFixa), num(p.custoUnit), num(p.precoUnit),
    p.criadoEm ? new Date(num(p.criadoEm)) : new Date(),
    p.atualizadoEm ? new Date(num(p.atualizadoEm)) : new Date(),
  ];
}

function lerProdutos() {
  const sh = abaProdutos();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, COLUNAS_PRODUTOS.length).getValues()
    .filter((r) => r[0] !== '' && r[0] !== null)
    .map((r) => ({
      id: String(r[0]), nome: String(r[1] || ''), material: String(r[2] || ''),
      pesoTotal: num(r[3]), horasTotal: num(r[4]), pecas: Math.max(1, num(r[5])), minutosMO: num(r[6]),
      margem: num(r[7]), taxaPct: num(r[8]), taxaFixa: num(r[9]), custoUnit: num(r[10]), precoUnit: num(r[11]),
      criadoEm: r[12] instanceof Date ? r[12].getTime() : 0,
      atualizadoEm: r[13] instanceof Date ? r[13].getTime() : 0,
    }));
}

function salvarProduto(p) {
  if (!p || !p.id) throw new Error('Produto sem ID.');
  const sh = abaProdutos();
  const row = linhaDoId(sh, p.id);
  if (row === -1) sh.appendRow(linhaDoProduto(p));
  else sh.getRange(row, 1, 1, COLUNAS_PRODUTOS.length).setValues([linhaDoProduto(p)]);
}

function excluirProduto(id) {
  const sh = abaProdutos();
  const row = linhaDoId(sh, id);
  if (row !== -1) sh.deleteRow(row);
}

/* ------------------------------------------------------------------ */

function lerInvestimento() {
  const v = abaConfig().getRange('B1').getValue();
  return v === '' || v === null ? null : num(v);
}

function salvarInvestimento(v) {
  abaConfig().getRange('B1').setValue(num(v));
}

function substituirTudo(lista, investimento, clientes, produtos) {
  limparVendas();
  const sh = abaVendas();
  const linhas = lista.filter((s) => s && s.id).map(linhaDaVenda);
  if (linhas.length) sh.getRange(2, 1, linhas.length, COLUNAS.length).setValues(linhas);
  if (investimento !== undefined && investimento !== null) salvarInvestimento(investimento);
  if (Array.isArray(clientes)) {
    const shc = abaClientes();
    const last = shc.getLastRow();
    if (last > 1) shc.deleteRows(2, last - 1);
    const lc = clientes.filter((c) => c && c.id).map(linhaDoCliente);
    if (lc.length) shc.getRange(2, 1, lc.length, COLUNAS_CLIENTES.length).setValues(lc);
  }
  if (Array.isArray(produtos)) {
    const shp = abaProdutos();
    const last = shp.getLastRow();
    if (last > 1) shp.deleteRows(2, last - 1);
    const lp = produtos.filter((p) => p && p.id).map(linhaDoProduto);
    if (lp.length) shp.getRange(2, 1, lp.length, COLUNAS_PRODUTOS.length).setValues(lp);
  }
}
