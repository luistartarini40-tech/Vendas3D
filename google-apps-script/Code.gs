/**
 * Planilha · API de sincronização com Google Planilhas
 *
 * Como usar:
 * 1. Crie uma planilha no Google Planilhas.
 * 2. Menu Extensões → Apps Script. Apague o conteúdo e cole este arquivo inteiro.
 * 3. Troque o valor de CHAVE abaixo por uma senha só sua.
 * 4. Implantar → Nova implantação → tipo "App da Web"
 *    - Executar como: Eu
 *    - Quem pode acessar: Qualquer pessoa
 * 5. Copie a URL que termina em /exec e cole no site (aba Financeiro → Nuvem),
 *    junto com a mesma CHAVE.
 */

const CHAVE = 'COLE_AQUI_SUA_CHAVE';

const ABA_VENDAS = 'Vendas';
const ABA_CONFIG = 'Config';
const COLUNAS = [
  'ID', 'Data', 'Peça', 'Material', 'Peso (g)', 'Tempo (h)', 'Quantidade',
  'Custo unit.', 'Preço unit.', 'Taxas unit.', 'Lucro unit.',
  'Custo total', 'Faturamento', 'Lucro líquido', 'Criado em',
];

/* ------------------------------------------------------------------ */

function doGet() {
  return saida({ ok: true, app: 'Planilha', mensagem: 'API funcionando. Use o site para acessar os dados.' });
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
        return saida({ ok: true });
      case 'load':
        return saida({ ok: true, sales: lerVendas(), investimento: lerInvestimento() });
      case 'add':
        adicionarVenda(req.sale);
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
      case 'replaceAll':
        substituirTudo(req.sales || [], req.investimento);
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

function abaVendas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_VENDAS);
  if (!sh) {
    sh = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0
      ? ss.getSheets()[0].setName(ABA_VENDAS)
      : ss.insertSheet(ABA_VENDAS);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, COLUNAS.length).setValues([COLUNAS])
      .setFontWeight('bold').setBackground('#c6f432').setFontColor('#000000');
    sh.setFrozenRows(1);
    sh.getRange('A2:A').setNumberFormat('@');
    sh.getRange('B2:B').setNumberFormat('dd/mm/yyyy');
    sh.getRange('E2:E').setNumberFormat('0');
    sh.getRange('F2:F').setNumberFormat('0.0');
    sh.getRange('H2:N').setNumberFormat('"R$" #,##0.00');
    sh.getRange('O2:O').setNumberFormat('dd/mm/yyyy hh:mm');
    sh.setColumnWidth(3, 220);
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

function linhaDaVenda(s) {
  const qtd = Math.max(1, num(s.qtd));
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
    .map((r) => ({
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
    }));
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
  if (linhaDoId(sh, s.id) !== -1) return; // já existe (envio repetido)
  sh.appendRow(linhaDaVenda(s));
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

function lerInvestimento() {
  const v = abaConfig().getRange('B1').getValue();
  return v === '' || v === null ? null : num(v);
}

function salvarInvestimento(v) {
  abaConfig().getRange('B1').setValue(num(v));
}

function substituirTudo(lista, investimento) {
  limparVendas();
  const sh = abaVendas();
  const linhas = lista.filter((s) => s && s.id).map(linhaDaVenda);
  if (linhas.length) sh.getRange(2, 1, linhas.length, COLUNAS.length).setValues(linhas);
  if (investimento !== undefined && investimento !== null) salvarInvestimento(investimento);
}
