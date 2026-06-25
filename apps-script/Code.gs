/**
 * ROTAS PETIKO — Backend em Apps Script (v2)
 * Compatível com a planilha "Petiko - Oficinas e Costureiras.xlsx"
 *
 * COMO CONFIGURAR:
 * 1. Suba a planilha "Petiko - Oficinas e Costureiras.xlsx" para o Google Sheets
 *    (Arquivo > Importar, ou abra direto no Drive). Mantenha os nomes das abas:
 *    "Oficinas", "Motoristas", "Movimentações".
 * 2. Extensões > Apps Script. Cole este código substituindo o conteúdo padrão.
 * 3. Implantar > Nova implantação > tipo "App da Web".
 *    - Executar como: Eu
 *    - Quem tem acesso: Qualquer pessoa
 * 4. Copie a URL gerada (termina em /exec) e cole no campo "URL do Apps Script"
 *    dentro da ferramenta de rotas.
 *
 * O Cleiton preenche a aba "Movimentações" com a rota do dia (Data, Oficina,
 * Tipo, Observação, etc). A ferramenta de rotas busca essas linhas pela data
 * de hoje e já vem com os endereços preenchidos, prontos para revisar e enviar.
 */

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = e.parameter.action;

  if (action === 'clientes') return jsonOut(lerOficinas(ss));
  if (action === 'motoristas') return jsonOut(lerMotoristas(ss));
  if (action === 'movimentacoes') return jsonOut(lerMovimentacoesDoDia(ss, e.parameter.data));
  if (action === 'salvar') return jsonOut(salvarRota(ss, e.parameter.data));

  return jsonOut({ ok: false, erro: 'ação inválida: ' + action });
}

// ---------- leitura por nome de cabeçalho (robusto a reordenar colunas) ----------
function indexarCabecalho(headerRow) {
  var idx = {};
  headerRow.forEach(function (h, i) {
    if (h) idx[String(h).trim()] = i;
  });
  return idx;
}

function lerOficinas(ss) {
  var sheet = ss.getSheetByName('Oficinas');
  if (!sheet) return { ok: false, erro: 'Aba "Oficinas" não encontrada' };
  var rows = sheet.getDataRange().getValues();
  var headerRowIdx = 3; // linha 4 da planilha (0-indexed = 3)
  var idx = indexarCabecalho(rows[headerRowIdx]);
  var data = [];
  for (var i = headerRowIdx + 1; i < rows.length; i++) {
    var nome = rows[i][idx['Nome']];
    if (!nome) continue;
    data.push({
      nome: String(nome).trim(),
      endereco: String(rows[i][idx['Endereço']] || '').trim(),
      status: String(rows[i][idx['Status']] || '').trim()
    });
  }
  return { ok: true, data: data };
}

function lerMotoristas(ss) {
  var sheet = ss.getSheetByName('Motoristas');
  if (!sheet) return { ok: false, erro: 'Aba "Motoristas" não encontrada' };
  var rows = sheet.getDataRange().getValues();
  var idx = indexarCabecalho(rows[3]);
  var data = [];
  for (var i = 4; i < rows.length; i++) {
    var nome = rows[i][idx['Nome']];
    if (!nome) continue;
    data.push({ nome: String(nome).trim(), telefone: String(rows[i][idx['Telefone (DDI+DDD+número)']] || '').trim() });
  }
  return { ok: true, data: data };
}

function lerMovimentacoesDoDia(ss, dataStr) {
  var sheet = ss.getSheetByName('Movimentações');
  if (!sheet) return { ok: false, erro: 'Aba "Movimentações" não encontrada' };
  var rows = sheet.getDataRange().getValues();
  var idx = indexarCabecalho(rows[3]);
  var alvo = dataStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var data = [];
  for (var i = 4; i < rows.length; i++) {
    var dataCel = rows[i][idx['Data']];
    if (!dataCel) continue;
    var dataFmt = (dataCel instanceof Date)
      ? Utilities.formatDate(dataCel, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(dataCel).trim();
    if (dataFmt !== alvo) continue;

    data.push({
      oficina: String(rows[i][idx['Oficina']] || '').trim(),
      tipo: String(rows[i][idx['Tipo']] || '').trim(),
      produto: String(rows[i][idx['Linha/Produto']] || '').trim(),
      quantidade: String(rows[i][idx['Quantidade']] || '').trim(),
      observacao: String(rows[i][idx['Observação']] || '').trim(),
      endereco: String(rows[i][idx['Endereço']] || '').trim(),
      motorista: String(rows[i][idx['Motorista']] || '').trim()
    });
  }
  return { ok: true, data: data };
}

function salvarRota(ss, dataStr) {
  var sheet = ss.getSheetByName('Rotas');
  if (!sheet) {
    sheet = ss.insertSheet('Rotas');
    sheet.appendRow(['Data/Hora', 'Motorista', 'Telefone', 'Resumo da rota', 'Detalhes (JSON)', 'Link Maps']);
  }
  var obj;
  try {
    obj = JSON.parse(dataStr);
  } catch (err) {
    return { ok: false, erro: 'JSON inválido recebido' };
  }
  sheet.appendRow([
    new Date(),
    obj.motorista || '',
    obj.telefone || '',
    obj.resumo || '',
    JSON.stringify(obj.paradas || []),
    obj.link || ''
  ]);
  return { ok: true };
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────
// DROPDOWN DE OFICINAS + PREENCHIMENTO AUTOMÁTICO DE ENDEREÇO
// ─────────────────────────────────────────────────────────────

/**
 * Trigger automático: quando o usuário seleciona uma oficina na coluna "Oficina"
 * da aba Movimentações, preenche automaticamente a coluna "Endereço".
 * Funciona sem configuração extra — o Google Sheets chama onEdit() sozinho.
 */
function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== 'Movimentações') return;

  var headerRow = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idx = indexarCabecalho(headerRow);

  var oficinaCol = (idx['Oficina'] !== undefined ? idx['Oficina'] : -1) + 1;
  var enderecoCol = (idx['Endereço'] !== undefined ? idx['Endereço'] : -1) + 1;

  if (oficinaCol <= 0 || enderecoCol <= 0) return;
  if (e.range.getColumn() !== oficinaCol) return;
  if (e.range.getRow() <= 4) return; // linha 4 é cabeçalho

  var oficinaNome = (e.value || '').trim();
  if (!oficinaNome) {
    sheet.getRange(e.range.getRow(), enderecoCol).clearContent();
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var oficinasSheet = ss.getSheetByName('Oficinas');
  if (!oficinasSheet) return;

  var rows = oficinasSheet.getDataRange().getValues();
  var oIdx = indexarCabecalho(rows[3]); // cabeçalho na linha 4 (índice 3)

  for (var i = 4; i < rows.length; i++) {
    var nome = String(rows[i][oIdx['Nome']] || '').trim();
    if (nome.toLowerCase() === oficinaNome.toLowerCase()) {
      var endereco = String(rows[i][oIdx['Endereço']] || '').trim();
      sheet.getRange(e.range.getRow(), enderecoCol).setValue(endereco);
      return;
    }
  }
}

/**
 * Execute UMA VEZ manualmente (Executar > setupDropdownOficinas) para criar
 * o dropdown de seleção na coluna "Oficina" da aba Movimentações.
 * Depois disso o onEdit cuida do resto automaticamente.
 */
function setupDropdownOficinas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var oficinasSheet = ss.getSheetByName('Oficinas');
  var movSheet = ss.getSheetByName('Movimentações');

  if (!oficinasSheet) { SpreadsheetApp.getUi().alert('Aba "Oficinas" não encontrada!'); return; }
  if (!movSheet)       { SpreadsheetApp.getUi().alert('Aba "Movimentações" não encontrada!'); return; }

  // Coleta nomes da aba Oficinas
  var rows = oficinasSheet.getDataRange().getValues();
  var oIdx = indexarCabecalho(rows[3]);
  var nomes = [];
  for (var i = 4; i < rows.length; i++) {
    var nome = String(rows[i][oIdx['Nome']] || '').trim();
    if (nome) nomes.push(nome);
  }

  if (!nomes.length) { SpreadsheetApp.getUi().alert('Nenhuma oficina encontrada na aba "Oficinas".'); return; }

  // Descobre em qual coluna fica "Oficina" na aba Movimentações
  var header = movSheet.getRange(4, 1, 1, movSheet.getLastColumn()).getValues()[0];
  var mIdx = indexarCabecalho(header);
  var oficinaCol = (mIdx['Oficina'] !== undefined ? mIdx['Oficina'] : -1) + 1;

  if (oficinaCol <= 0) { SpreadsheetApp.getUi().alert('Coluna "Oficina" não encontrada na linha 4 da aba Movimentações.'); return; }

  // Aplica validação (dropdown) nas linhas 5 até 1000
  var range = movSheet.getRange(5, oficinaCol, 996, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(nomes, true)
    .setAllowInvalid(true)
    .build();
  range.setDataValidation(rule);

  SpreadsheetApp.getUi().alert('✅ Dropdown configurado com ' + nomes.length + ' oficinas!\nAgora ao selecionar uma oficina o endereço será preenchido automaticamente.');
}

