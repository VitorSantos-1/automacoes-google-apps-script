/*
 * SISTEMA DE AUDITORIA E VENCIMENTO
 * log técnico de correções e mantendo como objeto JS para ficar dentro do arquivo de código.
 * ======================================================================================== */
const DICIONARIO_DE_ETAPAS = {
  "01_config_unificado": "Removido duplicações e chaves repetidas. HEADER_ROWS=5, DATA_START_ROW=6. STATUS_COL_LETTER='H'. Ajustado BLOCK.REQUIRED_COLS e PROTECTION.",
  "02_utilitarios_indices_vazios": "Criadas colIndex_ (1-based), colLetter_ (1-based->letra), isEmpty_ (placeholders). Mantidos colLetterToIndex (0-based) e getColIndex (0-based).",
  "03_bloqueio_onedit_v2": "Novo onEdit(e) + onEditHandler: limpa escrita fora da célula esperada e conduz o cursor, respeitando filtros/visualizações.",
  "04_respeita_filtros": "Detecção de linhas visíveis com isRowVisible_ compatível com Filtro normal e Visualização de filtro.",
  "05_varredura_rapida": "findFirstIncompleteRowFast_ lê bloco único de dados (1 chamada getValues) — rápido e estável.",
  "06_gestao_vencidos": "Fluxo moverProdutosVencidos() -> analisarVencimentos() -> executarMovimentacao(); copia linhas integrais, limpa colunas editáveis na origem e mantém fórmulas.",
  "07_analises_relatorios": "mostrarAnaliseStatus(), mostrarAnaliseOrdenada(), gerarRelatorioAnalise(), filtrarEExportarAnalise() e encontrarDuplicidades() ordenam de forma crescente e usam datas flexíveis.",
  "08_1000_linhas_formulas": "Módulo robusto: verificarEAplicarLimite(), garantirFormulasCompletas(), verificarIntegridadeFormulas(), onChangeHandler(), monitoramentoContinuo(), resetTotalDoSistema(), etc.",
  "09_ajuste_formulas": "ajustarFormulaParaNovaLinha() protege referências com $ na linha e ajusta apenas quando necessário.",
  "10_protecao_dinamica": "Aplicação/atualização automática de proteções de colunas (não ordenáveis) com proprietário autorizado; monitora onChangeProtecao.",
  "11_triggers_unicos": "onOpen() único cria menu e garante (sem duplicar) triggers: onEdit, onChangeHandler, monitoramentoContinuo, onChangeProtecao.",
  "12_compatibilidade_indices": "Concilia índices 0-based (arrays) e 1-based (Ranges). STATUS_COL_INDEX é 0-based e usado com +1 em getRange.",
  "13_bugs_corregidos": "Removidas funções duplicadas (toIndexesKeepOrder_, isRowVisible_), adicionadas ausentes (colIndex_, colLetter_, isEmpty_, onChangeProtecao), corrigidas referências inconsistentes.",
  "14_performance": "Uso de LockService nas rotinas críticas, flush pontual e mensagens de UI enxutas.",
  "15_segurança": "Regras valem para qualquer usuário (onEdit instalável) e sem depender de permissões especiais nas ações simples."
};

/*
 * Configurações
 * ============================================================ */
const CONFIG = {
  // ——— Gestão de vencidos ———
  EDITABLE_COLS: ['A', 'B', 'C', 'E', 'G', 'M'],
  SRC_SHEET: 'Matriz',
  DST_SHEET: 'Vencidos',
  HEADER_ROWS: 5,        // cabeçalho ocupa as linhas 1..5
  DATA_START_ROW: 6,     // dados começam na linha 6
  STATUS_COL_LETTER: 'M',
  BACKUP_ENABLED: false,
  BATCH_SIZE: 100,

  EXPIRY_DATE_COL_LETTER: 'G',
  PROD_NAME_COL_LETTER: 'D',
  QUANTITY_COL_LETTER: 'E',
  PRICE_COL_LETTER: '',

  ANALYSIS_SHEET: 'ANALISE_VENCIDOS',
  FILTER_SHEET: 'FILTRO_ANALISE',

  STATUS_ACCEPTED: ['vencido', 'vencidos'],
  KEY_COLS: ['B', 'C'],

  // ——— Regra de bloqueio ao pular linha (onEdit instalável) ———
  BLOCK: {
    SHEET: 'Matriz',                     // Aba vigiada
    REQUIRED_COLS: ['A', 'B', 'C', 'E', 'G'] // Colunas obrigatórias na ORDEM
  },

  // ——— Proteção de colunas ———
  PROTECTION: {
    COLS: ['D', 'F', 'H', 'I', 'J', 'K', 'L'],
    START_ROW: 7
  },

  // ——— Gatilhos ———
  TRIGGERS: {
    MONITOR_EVERY_MINUTES: 1,             // monitoramento contínuo
    AUTO_MOVE_DEFAULT_HOUR: 6             // rotina diária (06:00 por padrão)
  },

  // ——— Timezone p/ formatações ———
  TIMEZONE: Session.getScriptTimeZone() || 'America/Sao_Paulo'
};

// Valores que devem ser tratados como "vazios" (placeholders, hífens, etc.)
const EMPTY_PLACEHOLDERS = ['selecionar', 'selecione', '--', '—', '-', '(selecione)', '(selecionar)'];


/*
 * Utilitários
 * ============================================================ */
function colLetterToIndex(letter) { // 0-based
  if (!letter) return -1;
  const s = String(letter).trim().toUpperCase();
  if (!s) return -1;
  let sum = 0;
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i) - 64;
    if (code < 1 || code > 26) return -1;
    sum = sum * 26 + code;
  }
  return sum - 1; // zero-based
}
function getColIndex(letter) { // 0-based (para acessar arrays .getValues())
  return letter && String(letter).trim() ? colLetterToIndex(letter) : -1;
}
function colIndex_(col) { // 1-based (para Ranges)
  if (typeof col === 'number') return col;
  const z = colLetterToIndex(col);
  return z >= 0 ? z + 1 : -1;
}
function colLetter_(idx1) { // 1-based -> letra
  let n = Math.max(1, idx1 | 0), s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function isEmpty_(v) {
  if (v === null || v === '') return true;
  if (typeof v === 'string') {
    const s = v.trim().toLowerCase();
    if (!s) return true;
    if (EMPTY_PLACEHOLDERS.indexOf(s) >= 0) return true;
  }
  return false;
}

const STATUS_COL_INDEX = colLetterToIndex(CONFIG.STATUS_COL_LETTER); // 0-based

function parseDateFlexible(val) {
  if (!val) return null;

  if (val instanceof Date) {
    // TÉCNICA DO DESLOCAMENTO: Soma 12 horas (em milissegundos) para corrigir o fuso
    const dataSegura = new Date(val.getTime() + (12 * 60 * 60 * 1000));

    const dia = dataSegura.getDate();
    const mes = dataSegura.getMonth();
    const ano = dataSegura.getFullYear();

    // Retorna a data cravada ao meio-dia
    return new Date(ano, mes, dia, 12, 0, 0);
  }

  if (typeof val === 'string') {
    const parts = val.trim().split('/');
    if (parts.length >= 3) {
      // Garante que pega apenas os números, ignorando qualquer hora grudada na string
      const dia = parseInt(parts[0], 10);
      const mes = parseInt(parts[1], 10);
      const ano = parseInt(parts[2].substring(0, 4), 10); // Pega só os 4 dígitos do ano

      return new Date(ano, mes - 1, dia, 12, 0, 0);
    }
  }

  return null;
}

function diffInDays(a, b) {
  const d1 = new Date(a.getFullYear(), a.getMonth(), a.getDate());
  const d2 = new Date(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.floor((d1 - d2) / 86400000);
}
function ensureSheet(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function setTable(sheet, startRow, startCol, data, headerBold) {
  if (!data || !data.length) return;
  sheet.getRange(startRow, startCol, data.length, data[0].length).setValues(data);
  if (headerBold) sheet.getRange(startRow, startCol, 1, data[0].length).setFontWeight('bold');
}
function formatDateBR(d) {
  if (!d) return '—';

  // 1. Se for um objeto de data verdadeiro
  if (d instanceof Date) {
    const dia = String(d.getDate()).padStart(2, '0');
    const mes = String(d.getMonth() + 1).padStart(2, '0');
    const ano = d.getFullYear();
    return `${dia}/${mes}/${ano}`;
  }

  // 2. Se for texto vindo da planilha
  const dStr = String(d).trim();

  // Se tiver barras (ex: 03/18/2026 ou 18/03/2026)
  if (dStr.includes('/')) {
    const partes = dStr.split('/');
    if (partes.length === 3) {
      // O pulo do gato: Se o número do meio for maior que 12, o Google inverteu (MM/DD/YYYY)
      if (parseInt(partes[1]) > 12) {
        return `${partes[1].padStart(2, '0')}/${partes[0].padStart(2, '0')}/${partes[2]}`;
      }
      // Se já estiver no padrão correto (DD/MM/YYYY), só garante que tenha 2 dígitos
      return `${partes[0].padStart(2, '0')}/${partes[1].padStart(2, '0')}/${partes[2]}`;
    }
  }

  // Se vier no padrão de banco de dados (ex: 2026-03-18)
  if (dStr.includes('-')) {
    const partes = dStr.split('T')[0].split('-');
    if (partes.length === 3) {
      return `${partes[2]}/${partes[1]}/${partes[0]}`;
    }
  }

  return dStr;
}

/** Mantém a ordem definida em REQUIRED_COLS (sem ordenar) — 1-based */
function toIndexesKeepOrder_(cols) {
  if (!cols) return [];
  var out = [];
  for (var i = 0; i < cols.length; i++) {
    var c = cols[i];
    out.push(typeof c === 'number' ? c : colIndex_(String(c)));
  }
  return out; // sem sort
}

/** Linha visível na interface? (suporta Filtro e Visualização de filtro) */
function isRowVisible_(sheet, row) {
  try { if (sheet.isRowHiddenByFilter && sheet.isRowHiddenByFilter(row)) return false; } catch (_) { }
  try { if (sheet.isRowHiddenByUser && sheet.isRowHiddenByUser(row)) return false; } catch (_) { }
  return true;
}

/** Foca com segurança na célula alvo */
function safeActivate_(sheet, row, col) {
  try {
    var rg = sheet.getRange(row, col);
    rg.activate();                         // ativa a célula
    sheet.setActiveSelection(rg);          // reforça seleção
    SpreadsheetApp.flush();
  } catch (_) { }
}

/** Varrida rápida: lê os valores UMA VEZ e checa obrigatórias na ordem informada */
function findFirstIncompleteRowFast_(sheet, dataStartRow, reqColsIdxOrdered) {
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return null;
  if (!reqColsIdxOrdered || !reqColsIdxOrdered.length) return null;

  var maxCol = Math.max.apply(null, reqColsIdxOrdered);
  var total = lastRow - dataStartRow + 1;
  if (total <= 0) return null;

  // Lê todo o bloco apenas 1x (rápido)
  var block = sheet.getRange(dataStartRow, 1, total, maxCol).getValues();

  for (var off = 0; off < block.length; off++) {
    var rowNum = dataStartRow + off;
    if (!isRowVisible_(sheet, rowNum)) continue;              // respeita filtros e views

    var row = block[off];
    var missing = [];
    for (var k = 0; k < reqColsIdxOrdered.length; k++) {
      var cIdx = reqColsIdxOrdered[k];       // 1-based
      var v = row[cIdx - 1];                 // array 0-based
      if (isEmpty_(v)) missing.push(colLetter_(cIdx));
    }
    if (missing.length) return { rowNum: rowNum, missing: missing };
  }
  return null;
}


/*
 * Move a linha inteira e limpa a origem
 * ============================================================ */
function colLettersTo1BasedIndexes(letters) {
  return (letters || []).map(L => colLetterToIndex(L) + 1).filter(n => n > 0);
}
function copiarLinhasInteirasParaDestino(srcSheet, dstSheet, rowIndexList) {
  if (!rowIndexList || !rowIndexList.length) return;
  const lastCol = srcSheet.getLastColumn();
  const matriz = rowIndexList.map(r1 => srcSheet.getRange(r1, 1, 1, lastCol).getValues()[0]);
  const dstLastRow = dstSheet.getLastRow();
  const targetRow = dstLastRow < CONFIG.DATA_START_ROW ? CONFIG.DATA_START_ROW : (dstLastRow + 1);
  dstSheet.getRange(targetRow, 1, matriz.length, lastCol).setValues(matriz);
}
function limparSomenteEditaveisNaOrigem(srcSheet, rowIndexList) {
  const editableCols = colLettersTo1BasedIndexes(CONFIG.EDITABLE_COLS);
  if (!editableCols.length || !rowIndexList || !rowIndexList.length) return;
  rowIndexList.forEach(r1 => {
    editableCols.forEach(c1 => srcSheet.getRange(r1, c1, 1, 1).clearContent());
  });
}


/*
 * Gestão de vencidos
 * ============================================================ */
function criarBackup(ss) {
  const src = ss.getSheetByName(CONFIG.SRC_SHEET);
  if (!src) return;
  const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
  const backupName = `BACKUP_${CONFIG.SRC_SHEET}_${timestamp}`;
  const backup = src.copyTo(ss);
  backup.setName(backupName);
  backup.hideSheet();
  console.log(`Backup criado: ${backupName}`);
}

function analisarVencimentos(sheet) {
  if (!sheet) throw new Error('Sheet não fornecida para análise');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  if (numRows <= 0) return { vencidos: [], linhasParaRemover: [], totalAnalisado: 0, semStatus: 0 };

  const data = sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, lastCol).getValues();
  const r = { vencidos: [], linhasParaRemover: [], totalAnalisado: 0, semStatus: 0 };

  const COLUNA_H_INDEX = 7; // Coluna H
  const expIdx = getColIndex(CONFIG.EXPIRY_DATE_COL_LETTER); // Puxa a coluna de Validade
  const hoje = new Date();

  data.forEach((row, idx) => {
    const temConteudo = row.some(c => c !== '' && c !== null);
    if (!temConteudo) return;
    r.totalAnalisado++;

    const valorColunaH = row[COLUNA_H_INDEX];
    let isVencido = false;

    // 1. Verifica se está escrito "vencido" na Coluna H (Mantém a sua regra original)
    if (valorColunaH && typeof valorColunaH === 'string') {
      const s = valorColunaH.toLowerCase().trim();
      if (s.includes('vencido')) {
        isVencido = true;
      }
    }

    // 2. CORREÇÃO: Verifica pela Data de Validade se é Hoje (0 dias) ou menos
    if (!isVencido && expIdx >= 0) {
      const validade = parseDateFlexible(row[expIdx]);
      if (validade) {
        const diasRestantes = diffInDays(validade, hoje);
        if (diasRestantes <= 0) { // Se for 0 dias (vence hoje) ou negativo
          isVencido = true;
        }
      }
    }

    // Se passou em qualquer um dos testes, adiciona na lista para mover
    if (isVencido) {
      r.vencidos.push(row);
      r.linhasParaRemover.push(CONFIG.DATA_START_ROW + idx);
    }
  });

  return r;
}

function executarMovimentacao(ss, linhas, linhasParaRemover) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const src = ss.getSheetByName(CONFIG.SRC_SHEET);
    let dst = ss.getSheetByName(CONFIG.DST_SHEET);

    if (!dst) {
      dst = ss.insertSheet(CONFIG.DST_SHEET);
      const header = src.getRange(1, 1, CONFIG.HEADER_ROWS, src.getLastColumn());
      header.copyTo(dst.getRange(1, 1), { formatOnly: false, contentsOnly: false });
    }

    if (linhasParaRemover && linhasParaRemover.length) {
      copiarLinhasInteirasParaDestino(src, dst, linhasParaRemover);
      console.log(`✅ ${linhasParaRemover.length} linha(s) copiadas p/ "${CONFIG.DST_SHEET}"`);
    }

    limparSomenteEditaveisNaOrigem(src, linhasParaRemover || []);
    console.log('✅ Linhas de origem limpas nas colunas editáveis.');

    if (typeof garantirFormulasCompletas === 'function') {
      try {
        Utilities.sleep(200);
        garantirFormulasCompletas();
        console.log('✅ Sistema 1000 linhas verificado/aplicado.');
      } catch (e) {
        console.log('ℹ️ Sistema 1000 linhas não aplicado: ' + e);
      }
    }

    SpreadsheetApp.flush();
    console.log('✅ Movimentação concluída com sucesso!');
  } finally {
    lock.releaseLock();
  }
}

function moverProdutosVencidos() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  try {
    if (CONFIG.BACKUP_ENABLED) criarBackup(ss);
    const src = ss.getSheetByName(CONFIG.SRC_SHEET);
    if (!src) throw new Error(`Aba "${CONFIG.SRC_SHEET}" não encontrada!`);

    const lastRow = src.getLastRow();
    if (lastRow < CONFIG.DATA_START_ROW) {
      ui.alert('Aviso', 'Não há dados para processar na aba de origem.', ui.ButtonSet.OK);
      return;
    }

    const resultado = analisarVencimentos(src);
    if (resultado.vencidos.length === 0) {
      ui.alert('ℹ️ Informação',
        `Nenhum produto com status "vencido" encontrado.\n\n` +
        `Total de analisados: ${resultado.totalAnalisado}\n` +
        `Sem status: ${resultado.semStatus}`, ui.ButtonSet.OK);
      return;
    }

    const resp = ui.alert(
      '⚠️ Confirmação Necessária',
      `Encontrados ${resultado.vencidos.length} produto(s) marcado(s) como "vencido".\n\n` +
      `Mover para "${CONFIG.DST_SHEET}"? (copia valores, mantém fórmulas na origem, cria backup)`,
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) {
      ui.alert('Operação cancelada.');
      return;
    }

    executarMovimentacao(ss, resultado.vencidos, resultado.linhasParaRemover);

    ui.alert(
      '✅ Concluído',
      `📦 ${resultado.vencidos.length} produto(s) movido(s)\n` +
      `   ${formatDateBR(new Date())}\n` +
      `💾 Backup: ${CONFIG.BACKUP_ENABLED ? 'Sim' : 'Não'}\n` +
      `🔧 Fórmulas: Mantidas\n`, ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('❌ Erro', `Ocorreu um erro:\n${err}`, ui.ButtonSet.OK);
    console.error('Erro em moverProdutosVencidos:', err);
  } finally {
    lock.releaseLock();
  }
}

function executarMovimentacaoAutomatica() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    criarBackup(ss);
    const src = ss.getSheetByName(CONFIG.SRC_SHEET);
    if (!src) return;
    const r = analisarVencimentos(src);
    if (r.vencidos.length > 0) {
      executarMovimentacao(ss, r.vencidos, r.linhasParaRemover);
      console.log(`Auto: ${r.vencidos.length} produto(s) movido(s) em ${new Date()}`);
    }
  } catch (err) {
    console.error('Erro na execução automática:', err);
  } finally {
    lock.releaseLock();
  }
}


/*
 * Análises e Relatórios
 * ============================================================ */
function mostrarAnaliseStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SRC_SHEET);
  const ui = SpreadsheetApp.getUi();

  if (!src) return ui.alert('Erro', `Aba "${CONFIG.SRC_SHEET}" não encontrada.`, ui.ButtonSet.OK);

  const lastRow = src.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return ui.alert('Não há dados para analisar.');

  const data = src.getRange(CONFIG.DATA_START_ROW, STATUS_COL_INDEX + 1, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
  let vencidos = 0, semStatus = 0, totalProdutos = 0;
  const outrosStatus = {};

  data.forEach(row => {
    const status = row[0];
    if (status !== '' && status !== null) {
      totalProdutos++;
      const statusStr = String(status).toLowerCase().trim();
      if (CONFIG.STATUS_ACCEPTED.includes(statusStr)) {
        vencidos++;
      } else {
        const statusOriginal = String(status).trim();
        outrosStatus[statusOriginal] = (outrosStatus[statusOriginal] || 0) + 1;
      }
    } else {
      semStatus++;
    }
  });

  let outrosStatusMsg = '';
  Object.keys(outrosStatus).sort((a, b) => a.localeCompare(b, 'pt-BR')).forEach(st => {
    outrosStatusMsg += `• ${st}: ${outrosStatus[st]} produto(s)\n`;
  });

  ui.alert(
    ' Análise de Status',
    ` Data: ${formatDateBR(new Date())}\n\n` +
    `📦 Total de produtos: ${totalProdutos + semStatus}\n\n` +
    `🔴 "Vencido": ${vencidos}\n` +
    `⚪ Sem status: ${semStatus}\n\n` +
    ` Outros status (A→Z):\n` +
    (outrosStatusMsg || '• Nenhum\n') +
    `\n💡 Use "▶️ Mover Produtos Vencidos" para transferir os itens marcados.`,
    ui.ButtonSet.OK
  );
}

function mostrarAnaliseOrdenada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SRC_SHEET);
  const ui = SpreadsheetApp.getUi();

  if (!src) return ui.alert(`Aba "${CONFIG.SRC_SHEET}" não encontrada.`);

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW) return ui.alert('Não há dados para analisar.');

  const statusIdx = STATUS_COL_INDEX;
  const expIdx = getColIndex(CONFIG.EXPIRY_DATE_COL_LETTER);
  const nameIdx = getColIndex(CONFIG.PROD_NAME_COL_LETTER);
  const qtyIdx = getColIndex(CONFIG.QUANTITY_COL_LETTER);

  const headerValues = src
    .getRange(CONFIG.HEADER_ROWS, 1, 1, lastCol)
    .getValues()[0]
    .map(h => (h || '').toString().trim().toUpperCase());

  function findCol(candidatos) {
    for (const n of candidatos) {
      const i = headerValues.indexOf(n.toUpperCase());
      if (i >= 0) return i;
    }
    return -1;
  }

  const mercadoIdx = findCol(['MERCADOLÓGICO', 'MERCADOLOGICO', 'MERCADO']);
  const codeIdx = findCol(['CÓDIGO EM BARRA', 'CODIGO EM BARRA', 'CÓDIGO DE BARRAS', 'CODIGO DE BARRAS', 'EAN', 'CÓDIGO', 'CODIGO']);
  const diasRestIdx = findCol(['DE HOJE ATÉ O VENCIMENTO', 'DE HOJE ATE O VENCIMENTO', 'DIAS ATÉ O VENCIMENTO', 'DE HOJE ATÉ O VENCIMENTO (DIAS)']);
  const vendaMediaIdx = findCol(['VD MD DIA', 'VENDA MEDIA DIA', 'VENDA MÉDIA DIA', 'VD MD', 'VENDA MEDIA DIARIA']);
  const DIAS_CRITICO = 15;

  const values = src
    .getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, lastCol)
    .getValues();

  const hoje = new Date();
  let total = 0, semStatus = 0;
  const vencidos = [];
  const criticos = [];

  values.forEach((row, i) => {
    const temConteudo = row.some(c => c !== '' && c !== null);
    if (!temConteudo) return;

    total++;
    const s = row[statusIdx] != null ? String(row[statusIdx]).trim() : '';
    const sLower = s.toLowerCase();
    if (!s) { semStatus++; }

    const rowNum = CONFIG.DATA_START_ROW + i;
    const validade = expIdx >= 0 ? parseDateFlexible(row[expIdx]) : null;

    let diasRestantes = null;
    if (validade) {
      diasRestantes = diffInDays(validade, hoje);
    } else if (diasRestIdx >= 0) {
      const celVal = row[diasRestIdx];
      if (celVal !== '' && celVal != null) {
        const numerico = parseFloat(String(celVal).replace(/[^\d\-\.]/g, ''));
        if (!isNaN(numerico)) diasRestantes = numerico;
      }
    }
    const obj = {
      rowNum,
      mercado: mercadoIdx >= 0 ? (row[mercadoIdx] || '').toString().trim() : '',
      codigo: codeIdx >= 0 ? (row[codeIdx] || '').toString().trim() : '',
      descricao: nameIdx >= 0 ? (row[nameIdx] || '').toString().trim() : '',
      quantidade: qtyIdx >= 0 ? row[qtyIdx] : '',
      vendaMediaDia: vendaMediaIdx >= 0 ? (parseFloat(row[vendaMediaIdx]) || 0) : 0,
      validade,
      diasRestantes,
      status: s
    };

    // Se o status for aceito OU os dias restantes forem 0 ou menos, vai para Vencidos
    if (CONFIG.STATUS_ACCEPTED.includes(sLower) || (diasRestantes !== null && diasRestantes <= 0)) {
      vencidos.push(obj);
    } else if (diasRestantes !== null && diasRestantes <= DIAS_CRITICO) {
      criticos.push(obj);
    }
  });

  const sortDias = (a, b) => {
    const av = a.diasRestantes != null ? a.diasRestantes : Infinity;
    const bv = b.diasRestantes != null ? b.diasRestantes : Infinity;
    return av !== bv ? av - bv : a.rowNum - b.rowNum;
  };
  vencidos.sort(sortDias);
  criticos.sort(sortDias);

  // Agrupamento por status para as abas
  const critPendentes = criticos.filter(v => !v.status || v.status.trim() === '');
  const critAguardar = criticos.filter(v => (v.status || '').toUpperCase() === 'AGUARDAR');
  const critTrocas = criticos.filter(v => (v.status || '').toUpperCase() === 'TROCA');
  const critPercas = criticos.filter(v => (v.status || '').toUpperCase() === 'PERDA');
  const critRebaixas = criticos.filter(v => (v.status || '').toUpperCase() === 'REBAIXADO');

  const pctCritico = total > 0 ? ((criticos.length / total) * 100).toFixed(1) : '0.0';
  const pctVencido = total > 0 ? ((vencidos.length / total) * 100).toFixed(1) : '0.0';

  const mercadosCriticos = [...new Set(criticos.map(v => v.mercado || '—'))].sort();

  function esc(v) { return (v || '').toString().replace(/"/g, '&quot;').replace(/</g, '<').replace(/>/g, '>'); }
  function san(v) { return (v === '' || v == null) ? '—' : String(v); }

  function buildRowsCriticos(lista) {
    if (!lista.length) return `<tr class="empty-row"><td colspan="9">Nenhum produto crítico no momento</td></tr>`;
    return lista.map(v => {
      const valTxt = v.validade ? formatDateBR(v.validade) : '—';
      const d = v.diasRestantes;
      const cls = d != null && d <= 3 ? 'r-urgent' : d != null && d <= 7 ? 'r-warning' : 'r-caution';

      const badge = d === null ? '<span class="badge b-gray">—</span>' :
        d <= 0 ? '<span class="badge b-red">Vencido</span>' : // Novo caso para vencido
          d <= 3 ? '<span class="badge b-red">' + d + 'd restantes</span>' :
            d <= 7 ? '<span class="badge b-orange">' + d + 'd restantes</span>' :
              '<span class="badge b-yellow">' + d + 'd restantes</span>';

      // ── PROBABILIDADE DE VENDA ──────────────────────────────────
      let probBadge = '<span class="badge b-gray">—</span>';
      if (d !== null) {
        const prob = (d !== null && v.quantidade > 0 && v.vendaMediaDia > 0)
          ? Math.min(100, Math.max(0, Math.round((d * v.vendaMediaDia / v.quantidade) * 100)))
          : (v.vendaMediaDia === 0 ? 0 : null);
        const probCls = prob >= 70 ? 'b-green'
          : prob >= 40 ? 'b-yellow'
            : 'b-red';
        probBadge = `<span class="badge ${probCls}">${prob}%</span>`;
      }

      return `
        <tr class="${cls}" data-merc="${esc(v.mercado)}">
    	  <td class="col-linha">${v.rowNum}</td>
    	  <td class="col-merc"  title="${esc(v.mercado)}">${san(v.mercado)}</td>
    	  <td class="col-cod"   title="${esc(v.codigo)}">${san(v.codigo)}</td>
    	  <td class="col-prod"  title="${esc(v.descricao)}">${san(v.descricao)}</td>
    	  <td class="col-qtd">${san(v.quantidade)}</td>
    	  <td class="col-venc">${valTxt}</td>
    	  <td class="col-dias">${badge}</td>
    	  <td class="col-prob">${probBadge}</td>
          <td class="col-status">
            <select class="status-select" data-linha="${v.rowNum}" data-original="${v.status}" onchange="atualizarStatus(this)">
              <option value="" ${!v.status ? 'selected' : ''}>Selecionar</option>
              <option value="AGUARDAR" ${v.status.toUpperCase() === 'AGUARDAR' ? 'selected' : ''}>AGUARDAR</option>
              <option value="REBAIXADO" ${v.status.toUpperCase() === 'REBAIXADO' ? 'selected' : ''}>REBAIXADO</option>
              <option value="TROCA" ${v.status.toUpperCase() === 'TROCA' ? 'selected' : ''}>TROCA</option>
              <option value="PERDA" ${v.status.toUpperCase() === 'PERDA' ? 'selected' : ''}>PERDA</option>
              <option value="VENCIDO" ${v.status.toUpperCase() === 'VENCIDO' ? 'selected' : ''}>VENCIDO</option>
            </select>
          </td>
        </tr>`;
    }).join('');
  }

  function buildRowsVencidos(lista) {
    if (!lista.length) return `<tr class="empty-row"><td colspan="8">Nenhum produto vencido no momento</td></tr>`;
    return lista.map(v => {
      const valTxt = v.validade ? formatDateBR(v.validade) : '—';
      const d = v.diasRestantes;
      const badge = d === null ? '<span class="badge b-gray">—</span>'
        : `<span class="badge b-red">Vencido</span>`;
      const prob = (d !== null && v.quantidade > 0 && v.vendaMediaDia > 0)
        ? Math.min(100, Math.max(0, Math.round((d * v.vendaMediaDia / v.quantidade) * 100)))
        : (v.vendaMediaDia === 0 ? 0 : null);
      const probBadge = prob === null ? '<span class="badge b-gray">—</span>'
        : `<span class="badge ${prob >= 70 ? 'b-green' : prob >= 40 ? 'b-yellow' : 'b-red'}">${prob}%</span>`;
      return `
        <tr class="r-expired" data-rownum="${v.rowNum}">
          <td class="col-linha">${v.rowNum}</td>
          <td class="col-merc"  title="${esc(v.mercado)}">${san(v.mercado)}</td>
          <td class="col-cod"   title="${esc(v.codigo)}">${san(v.codigo)}</td>
          <td class="col-prod"  title="${esc(v.descricao)}">${san(v.descricao)}</td>
          <td class="col-qtd">${san(v.quantidade)}</td>
          <td class="col-venc">${valTxt}</td>
          <td class="col-dias">${badge}</td>
          <td class="col-prob">${probBadge}</td>
        </tr>`;
    }).join('');
  }

  const html = `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0d1117;--sf:#161b22;--sf2:#21262d;--br:rgba(255,255,255,.1);
  --txt:#e6edf3;--mut:#8b949e;--sub:#484f58;
  --blue:#58a6ff;--blue-d:rgba(88,166,255,.15);
  --grn:#3fb950;--grn-d:rgba(63,185,80,.15);
  --amb:#e3b341;--amb-d:rgba(227,179,65,.15);
  --red:#f85149;--red-d:rgba(248,81,73,.15);
  --pur:#a371f7;
}
html,body{height:100%;font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--txt);font-size:12px;overflow:hidden}
body{display:flex;flex-direction:column}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-thumb{background:var(--sub);border-radius:3px}
.header{background:linear-gradient(135deg,#0d1117,#1a2234);border-bottom:1px solid var(--br);padding:0 20px;height:46px;min-height:46px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.h-left{display:flex;align-items:center;gap:12px}
.h-logo{width:28px;height:28px;border-radius:8px;background:linear-gradient(135deg,var(--blue),var(--pur));display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:800;color:#fff;flex-shrink:0}
.h-title{font-size:13px;font-weight:700}
.h-sep{width:1px;height:16px;background:var(--br)}
.h-sub{font-size:11px;color:var(--mut)}
.h-date{font-size:10.5px;color:var(--mut);background:var(--sf2);border:1px solid var(--br);border-radius:6px;padding:3px 11px}
.cards{background:var(--sf);border-bottom:1px solid var(--br);padding:9px 20px;display:flex;gap:8px;flex-shrink:0}
.card{flex:1;border:1px solid var(--br);border-radius:10px;padding:9px 13px;background:var(--sf2);display:flex;flex-direction:column;gap:2px;cursor:default;transition:transform .15s,box-shadow .15s;position:relative;overflow:hidden}
.card:hover{transform:translateY(-1px);box-shadow:0 4px 14px rgba(0,0,0,.3)}
.card-acc{position:absolute;top:0;left:0;right:0;height:2px}
.c-total .card-acc{background:linear-gradient(90deg,var(--blue),var(--pur))}
.c-sem  .card-acc{background:var(--sub)}
.c-crit .card-acc{background:var(--amb)}
.c-venc .card-acc{background:var(--red)}
.c-troc .card-acc{background:var(--grn)}
.c-perc .card-acc{background:var(--red)}
.c-reb  .card-acc{background:var(--amb)}
.card-n{font-size:22px;font-weight:800;line-height:1;margin-top:3px}
.c-total .card-n{background:linear-gradient(135deg,var(--blue),var(--pur));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.c-sem  .card-n{color:var(--mut)}
.c-crit .card-n{color:var(--amb)}
.c-venc .card-n{color:var(--red)}
.c-troc .card-n{color:var(--grn)}
.c-perc .card-n{color:var(--red)}
.c-reb  .card-n{color:var(--amb)}
.card-l{font-size:9px;text-transform:uppercase;letter-spacing:.7px;font-weight:600;color:var(--mut)}
.card-s{font-size:9px;color:var(--sub)}
.body{flex:1;min-height:0;padding:9px 14px;display:flex;flex-direction:column;gap:9px;overflow:hidden}
.panel{background:var(--sf);border:1px solid var(--br);border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.2);display:flex;flex-direction:column;position:relative}
.panel-head{height:40px;min-height:40px;padding:0 14px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--br);background:var(--sf2);border-radius:10px 10px 0 0;flex-shrink:0}
.ph-left{display:flex;align-items:center;gap:8px}
.ph-bar{width:3px;height:16px;border-radius:3px;flex-shrink:0}
.bar-a{background:linear-gradient(180deg,var(--amb),#d97706)}
.bar-r{background:linear-gradient(180deg,var(--red),#dc2626)}
.ph-title{font-size:12px;font-weight:700}
.ph-pill{font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px}
.pill-a{background:var(--amb-d);color:var(--amb);border:1px solid rgba(227,179,65,.3)}
.pill-r{background:var(--red-d);color:var(--red);border:1px solid rgba(248,81,73,.3)}
.ph-right{display:flex;align-items:center;gap:8px}
.tabs{display:flex;border-bottom:1px solid var(--br);background:var(--sf);flex-shrink:0;padding:0 14px}
.tab-btn{padding:7px 15px;font-size:11px;font-weight:600;color:var(--mut);background:none;border:none;border-bottom:2px solid transparent;cursor:pointer;display:flex;align-items:center;gap:6px;transition:color .15s,border-color .15s;margin-bottom:-1px;font-family:inherit}
.tab-btn:hover{color:var(--txt)}
.tab-btn.active{color:var(--txt)}
.tab-btn[data-tab="criticos"].active{border-bottom-color:var(--blue)}
.tab-btn[data-tab="trocas"].active{border-bottom-color:var(--grn)}
.tab-btn[data-tab="percas"].active{border-bottom-color:var(--red)}
.tab-btn[data-tab="rebaixas"].active{border-bottom-color:var(--amb)}
.tab-cnt{font-size:9px;font-weight:700;padding:1px 6px;border-radius:10px;background:var(--sf2);color:var(--mut);min-width:18px;text-align:center;transition:background .15s,color .15s}
.tab-btn[data-tab="criticos"].active .tab-cnt{background:var(--blue-d);color:var(--blue)}
.tab-btn[data-tab="trocas"].active   .tab-cnt{background:var(--grn-d);color:var(--grn)}
.tab-btn[data-tab="percas"].active   .tab-cnt{background:var(--red-d);color:var(--red)}
.tab-btn[data-tab="rebaixas"].active .tab-cnt{background:var(--amb-d);color:var(--amb)}
.tab-content{display:none}
.tab-content.active{display:block}
.filter-wrap{display:flex;align-items:center;gap:8px}
.filter-count{font-size:10px;color:var(--mut);white-space:nowrap}
.filter-select{height:26px;padding:0 9px;background:var(--sf2);color:var(--txt);border:1px solid var(--br);border-radius:6px;font-size:11px;font-family:inherit;cursor:pointer;outline:none;min-width:200px;transition:border-color .15s}
.filter-select:focus{border-color:var(--amb)}
.filter-select option{background:var(--sf2)}
.btn-move{height:27px;padding:0 14px;background:linear-gradient(135deg,#dc2626,#b91c1c);color:#fff;border:none;border-radius:7px;font-size:11px;font-weight:600;cursor:pointer;display:inline-flex;align-items:center;gap:5px;transition:opacity .15s,transform .1s;flex-shrink:0;font-family:inherit}
.btn-move:hover{opacity:.9;transform:translateY(-1px)}
.btn-report{height:27px;padding:0 13px;background:linear-gradient(135deg,#1e40af,#1e3a8a);color:#fff;border:none;border-radius:7px;font-size:11px;font-weight:600;cursor:pointer;display:inline-flex;align-items:center;gap:5px;transition:opacity .15s,transform .1s;flex-shrink:0;position:relative;font-family:inherit}
.btn-report:hover{opacity:.9;transform:translateY(-1px)}
.btn-report:disabled{opacity:.5;cursor:not-allowed;transform:none}
.tbl-wrap{overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:11.5px}
thead th{position:sticky;top:0;z-index:2;background:#0d1117;color:var(--mut);font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;padding:6px 10px;text-align:left;white-space:nowrap;border-bottom:1px solid var(--br)}
th.tc{text-align:center}
tbody td{padding:6px 10px;border-bottom:1px solid rgba(255,255,255,.04);vertical-align:middle;color:var(--txt);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;transition:background .1s}
tbody tr:hover td{background:rgba(255,255,255,.05)!important}
tbody tr.hidden{display:none}
.col-linha{width:46px;text-align:center;color:var(--sub);font-size:10.5px}
.col-merc{width:150px;color:var(--mut);font-size:11px}
.col-cod{width:125px;font-family:'Courier New',monospace;font-size:10.5px;color:var(--mut)}
.col-prod{color:var(--txt);font-weight:500}
.col-qtd{width:50px;text-align:center;font-weight:700}
.col-venc{width:88px;text-align:center;color:var(--mut);font-size:10.5px}
.col-dias{width:120px;text-align:center}
.col-prob{width:105px;text-align:center}
.col-status{width:120px;text-align:center}
.status-select{height:23px;padding:0 7px;background:var(--sf2);color:var(--txt);border:1px solid var(--br);border-radius:6px;font-size:10.5px;font-family:inherit;cursor:pointer;outline:none;transition:border-color .15s}
.status-select:focus{border-color:var(--blue)}
.status-select option{background:var(--sf2)}
.r-expired td{background:rgba(248,81,73,.06)!important}
.r-urgent  td{background:rgba(248,81,73,.06)!important}
.r-warning td{background:rgba(227,179,65,.06)!important}
.r-caution td{background:rgba(227,179,65,.03)!important}
.badge{display:inline-flex;align-items:center;padding:2px 9px;border-radius:20px;font-size:10px;font-weight:600;white-space:nowrap}
.b-green{background:var(--grn-d);color:var(--grn)}
.b-yellow{background:var(--amb-d);color:var(--amb)}
.b-red{background:var(--red-d);color:var(--red)}
.b-gray{background:rgba(255,255,255,.07);color:var(--mut)}
.b-orange{background:rgba(249,115,22,.15);color:#fb923c}
.empty-row td{text-align:center;padding:18px;color:var(--sub);font-size:11px;background:var(--sf)!important;font-style:italic}
.footer{min-height:26px;padding:0 18px;background:var(--sf);border-top:1px solid var(--br);display:flex;align-items:center;justify-content:space-between;font-size:10px;color:var(--sub);flex-shrink:0}
.f-dot{width:6px;height:6px;border-radius:50%;background:var(--grn);display:inline-block;margin-right:5px;animation:pulse 2s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
.export-menu{display:none;position:fixed;background:var(--sf2);border:1px solid var(--br);border-radius:10px;box-shadow:0 12px 32px rgba(0,0,0,.5);z-index:99999;min-width:190px;overflow:hidden}
.export-menu.open{display:block}
.export-item{display:flex;align-items:center;gap:10px;padding:10px 14px;font-size:11.5px;color:var(--txt);cursor:pointer;text-decoration:none;transition:background .1s}
.export-item:hover{background:rgba(255,255,255,.06)}
.export-item .ext{font-size:9px;font-weight:700;padding:1px 5px;border-radius:4px;letter-spacing:.5px}
.ext-xlsx{background:var(--grn-d);color:var(--grn)}
.ext-pdf{background:var(--red-d);color:var(--red)}
.spinner{display:none;width:12px;height:12px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:spin .6s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
#toast{position:fixed;bottom:38px;right:18px;background:var(--sf2);border:1px solid var(--br);border-radius:8px;padding:9px 15px;font-size:11.5px;color:var(--txt);box-shadow:0 8px 24px rgba(0,0,0,.4);opacity:0;transform:translateY(8px);transition:opacity .2s,transform .2s;pointer-events:none;z-index:9999}
#toast.show{opacity:1;transform:translateY(0)}
</style></head>
<body>
<div class="header">
  <div class="h-left">
    <span class="h-title">Auditoria de Vencimento</span>
    <div class="h-sep"></div>
    <span class="h-sub">Análise mercadológica</span>
  </div>
  <span class="h-date"> ${formatDateBR(new Date())}</span>
</div>
<div class="cards">
  <div class="card c-total"><div class="card-acc"></div><span class="card-n">${total}</span><span class="card-l">Total</span></div>
  <div class="card c-sem"><div class="card-acc"></div><span class="card-n">${semStatus}</span><span class="card-l">Sem status</span></div>
  <div class="card c-crit"><div class="card-acc"></div><span class="card-n" id="cn-crit">${criticos.length}</span><span class="card-l">Críticos</span><span class="card-s">${pctCritico}%</span></div>
  <div class="card c-venc"><div class="card-acc"></div><span class="card-n">${vencidos.length}</span><span class="card-l">Vencidos</span><span class="card-s">${pctVencido}%</span></div>
  <div class="card c-agua"><div class="card-acc"></div><span class="card-n" id="cn-aguardar">${critAguardar.length}</span><span class="card-l">Aguardar</span></div>
  <div class="card c-troc"><div class="card-acc"></div><span class="card-n" id="cn-trocas">${critTrocas.length}</span><span class="card-l">Trocas</span></div>
  <div class="card c-perc"><div class="card-acc"></div><span class="card-n" id="cn-percas">${critPercas.length}</span><span class="card-l">Perca</span></div>
  <div class="card c-reb"><div class="card-acc"></div><span class="card-n" id="cn-rebaixas">${critRebaixas.length}</span><span class="card-l">Rebaixa</span></div>
</div>
<div class="body" id="body-wrap">
  <div class="panel" id="panel-crit">
    <div class="panel-head">
      <div class="ph-left">
        <div class="ph-bar bar-a"></div>
        <span class="ph-title">Produtos em atenção</span>
        <span class="ph-pill pill-a" id="pill-crit">${criticos.length} produto(s)</span>
      </div>
      <div class="ph-right">
        <div class="filter-wrap">
          <span class="filter-count" id="filter-count"></span>
          <select class="filter-select" id="filtro-mercado" onchange="filtrarMercado()">
            <option value="">Todos os mercadológicos</option>
            ${mercadosCriticos.map(m => `<option value="${esc(m)}">${esc(m)}</option>`).join('')}
          </select>
        </div>
        <button class="btn-report" id="btn-report" onclick="gerarRelatorio()">
          <div class="spinner" id="spinner-report"></div>
          <span id="label-report">⬇ Exportar relatório</span>
        </button>
      </div>
    </div>
    <div class="tabs">
      <button class="tab-btn active" data-tab="criticos" onclick="switchTab('criticos')">Críticos<span class="tab-cnt" id="badge-criticos">${critPendentes.length}</span></button>
      <button class="tab-btn" data-tab="aguardar" onclick="switchTab('aguardar')">Aguardar<span class="tab-cnt" id="badge-aguardar">${critAguardar.length}</span></button>
      <button class="tab-btn" data-tab="trocas" onclick="switchTab('trocas')">Trocas<span class="tab-cnt" id="badge-trocas">${critTrocas.length}</span></button>
      <button class="tab-btn" data-tab="percas" onclick="switchTab('percas')">Perca<span class="tab-cnt" id="badge-percas">${critPercas.length}</span></button>
      <button class="tab-btn" data-tab="rebaixas" onclick="switchTab('rebaixas')">Rebaixa<span class="tab-cnt" id="badge-rebaixas">${critRebaixas.length}</span></button>
    </div>
    <div class="tbl-wrap" id="wrap-crit">
      <div class="tab-content active" id="tab-criticos">
        <table><colgroup><col style="width:46px"><col style="width:150px"><col style="width:125px"><col><col style="width:50px"><col style="width:88px"><col style="width:120px"><col style="width:105px"><col style="width:120px"></colgroup>
        <thead><tr><th class="tc">Linha</th><th>Mercadológico</th><th>Cód. Barra</th><th>Produto</th><th class="tc">Qtd</th><th class="tc">Vencimento</th><th class="tc">Até vencer</th><th class="tc">Prob. Venda</th><th class="tc">Status</th></tr></thead>
        <tbody id="tbody-criticos">${buildRowsCriticos(critPendentes)}</tbody></table>
      </div>
      <div class="tab-content" id="tab-aguardar">
        <table><colgroup><col style="width:46px"><col style="width:150px"><col style="width:125px"><col><col style="width:50px"><col style="width:88px"><col style="width:120px"><col style="width:105px"><col style="width:120px"></colgroup>
        <thead><tr><th class="tc">Linha</th><th>Mercadológico</th><th>Cód. Barra</th><th>Produto</th><th class="tc">Qtd</th><th class="tc">Vencimento</th><th class="tc">Até vencer</th><th class="tc">Prob. Venda</th><th class="tc">Status</th></tr></thead>
        <tbody id="tbody-aguardar">${buildRowsCriticos(critAguardar)}</tbody></table>
      </div>
      <div class="tab-content" id="tab-trocas">
        <table><colgroup><col style="width:46px"><col style="width:150px"><col style="width:125px"><col><col style="width:50px"><col style="width:88px"><col style="width:120px"><col style="width:105px"><col style="width:120px"></colgroup>
        <thead><tr><th class="tc">Linha</th><th>Mercadológico</th><th>Cód. Barra</th><th>Produto</th><th class="tc">Qtd</th><th class="tc">Vencimento</th><th class="tc">Até vencer</th><th class="tc">Prob. Venda</th><th class="tc">Status</th></tr></thead>
        <tbody id="tbody-trocas">${buildRowsCriticos(critTrocas)}</tbody></table>
      </div>
      <div class="tab-content" id="tab-percas">
        <table><colgroup><col style="width:46px"><col style="width:150px"><col style="width:125px"><col><col style="width:50px"><col style="width:88px"><col style="width:120px"><col style="width:105px"><col style="width:120px"></colgroup>
        <thead><tr><th class="tc">Linha</th><th>Mercadológico</th><th>Cód. Barra</th><th>Produto</th><th class="tc">Qtd</th><th class="tc">Vencimento</th><th class="tc">Até vencer</th><th class="tc">Prob. Venda</th><th class="tc">Status</th></tr></thead>
        <tbody id="tbody-percas">${buildRowsCriticos(critPercas)}</tbody></table>
      </div>
      <div class="tab-content" id="tab-rebaixas">
        <table><colgroup><col style="width:46px"><col style="width:150px"><col style="width:125px"><col><col style="width:50px"><col style="width:88px"><col style="width:120px"><col style="width:105px"><col style="width:120px"></colgroup>
        <thead><tr><th class="tc">Linha</th><th>Mercadológico</th><th>Cód. Barra</th><th>Produto</th><th class="tc">Qtd</th><th class="tc">Vencimento</th><th class="tc">Até vencer</th><th class="tc">Prob. Venda</th><th class="tc">Status</th></tr></thead>
        <tbody id="tbody-rebaixas">${buildRowsCriticos(critRebaixas)}</tbody></table>
      </div>
    </div>
  </div>
  <div class="panel" id="panel-venc">
    <div class="panel-head">
      <div class="ph-left">
        <div class="ph-bar bar-r"></div>
        <span class="ph-title">Vencidos</span>
        <span class="ph-pill pill-r" id="pill-venc">${vencidos.length} produto(s)</span>
      </div>
      <button class="btn-move" onclick="moverVencidos()">&#9654; Mover vencidos</button>
    </div>
    <div class="tbl-wrap" id="wrap-venc">
      <table><colgroup><col style="width:46px"><col style="width:150px"><col style="width:125px"><col><col style="width:50px"><col style="width:88px"><col style="width:120px"><col style="width:105px"></colgroup>
      <thead><tr><th class="tc">Linha</th><th>Mercadológico</th><th>Cód. Barra</th><th>Produto</th><th class="tc">Qtd</th><th class="tc">Vencimento</th><th class="tc">Situação</th><th class="tc">Prob. Venda</th></tr></thead>
      <tbody id="tbody-venc">${buildRowsVencidos(vencidos)}</tbody></table>
    </div>
  </div>
</div>
<div class="footer">
  <span><span class="f-dot"></span>Sistema em tempo real · dados da aba Matriz</span>
  <span>Use "Exportar relatório" para baixar em Excel ou PDF</span>
</div>
<div id="export-menu" class="export-menu">
  <a class="export-item" onclick="baixar('xlsx')"><span class="ext ext-xlsx">XLSX</span>Baixar como Excel</a>
  <a class="export-item" onclick="baixar('pdf')"><span class="ext ext-pdf">PDF</span>Baixar como PDF</a>
</div>
<div id="toast"></div>
<script>
  var ROW_H=32,THEAD_H=34;
  function ajustarAlturas(){
    var body=document.getElementById('body-wrap'),wc=document.getElementById('wrap-crit'),wv=document.getElementById('wrap-venc');
    if(!body||!wc||!wv)return;
    var hc=document.getElementById('panel-crit').querySelector('.panel-head').offsetHeight;
    var tabsEl=document.querySelector('.tabs');var tabsH=tabsEl?tabsEl.offsetHeight:0;
    var hv=document.getElementById('panel-venc').querySelector('.panel-head').offsetHeight;
    var gap=9,bodyH=body.clientHeight;
    var act=document.querySelector('.tab-content.active');
    var tb=act?act.querySelector('tbody'):null;
    var vis=tb?tb.querySelectorAll('tr:not(.hidden):not(.empty-row)').length:0;
    var cH=Math.min(THEAD_H+Math.max(vis,1)*ROW_H,Math.floor(bodyH*.58));
    cH=Math.max(cH,THEAD_H+ROW_H);
    var vH=Math.max(bodyH-hc-tabsH-cH-gap-hv,THEAD_H+ROW_H);
    wc.style.height=cH+'px';wv.style.height=vH+'px';
  }
  function switchTab(t){
    document.querySelectorAll('.tab-btn').forEach(function(b){b.classList.remove('active')});
    document.querySelectorAll('.tab-content').forEach(function(c){c.classList.remove('active')});
    document.querySelector('[data-tab="'+t+'"]').classList.add('active');
    document.getElementById('tab-'+t).classList.add('active');
    ajustarAlturas();
  }
  function updateCounts(){
    var map={criticos:'cn-crit',aguardar:'cn-aguardar',trocas:'cn-trocas',percas:'cn-percas',rebaixas:'cn-rebaixas'};
    var totalCrit=0;
    Object.keys(map).forEach(function(t){
      var tb=document.getElementById('tbody-'+t);
      var n=tb?tb.querySelectorAll('tr:not(.empty-row)').length:0;
      totalCrit+=n;
      var b=document.getElementById('badge-'+t);if(b)b.textContent=n;
      var c=document.getElementById(map[t]);if(c)c.textContent=n;
    });
    var p=document.getElementById('pill-crit');if(p)p.textContent=totalCrit+' produto(s)';
    var cTotCrit=document.querySelector('.c-crit .card-n');if(cTotCrit)cTotCrit.textContent=totalCrit;
    
    var tbCrit=document.getElementById('tbody-criticos'); 
    var semStatusNum = tbCrit ? tbCrit.querySelectorAll('tr:not(.empty-row)').length : 0;
    var cSem=document.querySelector('.c-sem .card-n');if(cSem)cSem.textContent=semStatusNum;

    var tbVenc=document.getElementById('tbody-venc');
    var vencidosNum = tbVenc ? tbVenc.querySelectorAll('tr:not(.empty-row)').length : 0;
    var cVenc=document.querySelector('.c-venc .card-n');if(cVenc)cVenc.textContent=vencidosNum;
    var pVenc=document.getElementById('pill-venc');if(pVenc)pVenc.textContent=vencidosNum+' produto(s)';
    
    var totalGeral=totalCrit+vencidosNum;
    var cTotal=document.querySelector('.c-total .card-n');if(cTotal)cTotal.textContent=totalGeral;
  }
  function fixEmpty(tbodyId){
    var tb=document.getElementById(tbodyId);if(!tb)return;
    var real=tb.querySelectorAll('tr:not(.empty-row)');
    var emp=tb.querySelector('.empty-row');
    if(real.length===0&&!emp){var tr=document.createElement('tr');tr.className='empty-row';var td=document.createElement('td');td.colSpan=9;td.textContent='Nenhum produto nessa categoria';tr.appendChild(td);tb.appendChild(tr);}
    else if(real.length>0&&emp)emp.parentNode.removeChild(emp);
  }
  function atualizarStatus(sel){
    var linha=sel.getAttribute('data-linha'),novoStatus=(sel.value||'').trim();
    var orig=sel.getAttribute('data-original')||'',tr=sel.closest('tr');
    var tabMap={AGUARDAR:'aguardar',TROCA:'trocas',PERDA:'percas',REBAIXADO:'rebaixas'};
    var dest=tabMap[novoStatus.toUpperCase()]||'criticos';
    var destTb=document.getElementById('tbody-'+dest),srcTb=tr.parentNode;
    google.script.run
      .withSuccessHandler(function(){
        var emp=destTb.querySelector('.empty-row');if(emp)destTb.removeChild(emp);
        srcTb.removeChild(tr);destTb.appendChild(tr);
        sel.setAttribute('data-original',novoStatus);
        fixEmpty(srcTb.id);updateCounts();ajustarAlturas();
        showToast('✓ Status: '+(novoStatus||'limpo'));switchTab(dest);
      })
      .withFailureHandler(function(err){sel.value=orig;alert('Erro: '+(err&&err.message?err.message:err));})
      .atualizarStatusNaPlanilha(linha,novoStatus);
  }
  function moverVencidos(){
    if(!confirm('Mover vencidos para a aba "Vencidos"?'))return;
    google.script.run
      .withSuccessHandler(function(){
        var tb=document.getElementById('tbody-venc');
        if(tb)tb.innerHTML='<tr class="empty-row"><td colspan="8">Nenhum produto vencido no momento</td></tr>';
        updateCounts();
        ajustarAlturas();showToast('✓ Produtos vencidos movidos!');
      })
      .withFailureHandler(function(err){alert('Erro: '+(err&&err.message?err.message:err));})
      .moverProdutosVencidos();
  }
  function filtrarMercado(){
    var filtro=document.getElementById('filtro-mercado').value.trim().toLowerCase();
    var act=document.querySelector('.tab-content.active');
    var rows=act?act.querySelectorAll('tbody tr'):[];var vis=0;
    rows.forEach(function(r){
      if(r.classList.contains('empty-row'))return;
      var ok=!filtro||(r.getAttribute('data-merc')||'').toLowerCase()===filtro;
      r.classList.toggle('hidden',!ok);if(ok)vis++;
    });
    var ce=document.getElementById('filter-count');
    if(ce)ce.textContent=filtro?(vis+' visíveis'):'';
    ajustarAlturas();
  }
  function showToast(msg){var t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');setTimeout(function(){t.classList.remove('show');},2500);}
  function gerarRelatorio(){
    var menu=document.getElementById('export-menu'),btn=document.getElementById('btn-report');
    if(menu.classList.contains('open')){menu.classList.remove('open');return;}
    var r=btn.getBoundingClientRect();menu.style.top=(r.bottom+4)+'px';menu.style.left=(r.right-190)+'px';menu.classList.add('open');
  }
  document.addEventListener('click',function(e){
    var btn=document.getElementById('btn-report'),menu=document.getElementById('export-menu');
    if(btn&&menu&&!btn.contains(e.target)&&!menu.contains(e.target))menu.classList.remove('open');
  });
  function baixar(fmt){
    var menu=document.getElementById('export-menu'),btn=document.getElementById('btn-report');
    var sp=document.getElementById('spinner-report'),lb=document.getElementById('label-report');
    var filtro=(document.getElementById('filtro-mercado')||{}).value||'';
    var abaAtiva=document.querySelector('.tab-btn.active').getAttribute('data-tab');
    menu.classList.remove('open');btn.disabled=true;sp.style.display='block';lb.textContent='Gerando...';
    google.script.run
      .withSuccessHandler(function(r){btn.disabled=false;sp.style.display='none';lb.textContent='⬇ Exportar relatório';window.open(fmt==='xlsx'?r.xlsx:r.pdf,'_blank');})
      .withFailureHandler(function(err){btn.disabled=false;sp.style.display='none';lb.textContent='⬇ Exportar relatório';alert('Erro: '+err.message);})
      .gerarRelatorioCriticosExcel(filtro, abaAtiva);
  }
  window.addEventListener('load',ajustarAlturas);
  window.addEventListener('resize',ajustarAlturas);
</script>
</body></html>`;

  const htmlOutput = HtmlService
    .createHtmlOutput(html)
    .setTitle('Análise Detalhada')
    .setWidth(1920)
    .setHeight(1080);

  ui.showModalDialog(htmlOutput, 'Auditoria De Vencimento');
}

function gerarRelatorioCriticosExcel(filtroMercado, filtroAba) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SRC_SHEET);

  if (!src) throw new Error(`Aba "${CONFIG.SRC_SHEET}" não encontrada.`);

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW) throw new Error('Não há dados para analisar.');

  const statusIdx = STATUS_COL_INDEX;
  const expIdx = getColIndex(CONFIG.EXPIRY_DATE_COL_LETTER);
  const nameIdx = getColIndex(CONFIG.PROD_NAME_COL_LETTER);
  const qtyIdx = getColIndex(CONFIG.QUANTITY_COL_LETTER);

  const headerValues = src
    .getRange(CONFIG.HEADER_ROWS, 1, 1, lastCol)
    .getValues()[0]
    .map(h => (h || '').toString().trim().toUpperCase());

  function findCol(candidatos) {
    for (const n of candidatos) {
      const i = headerValues.indexOf(n.toUpperCase());
      if (i >= 0) return i;
    }
    return -1;
  }

  const mercadoIdx = findCol(['MERCADOLÓGICO', 'MERCADOLOGICO', 'MERCADO']);
  const codeIdx = findCol(['CÓDIGO EM BARRA', 'CODIGO EM BARRA', 'CÓDIGO DE BARRAS', 'CODIGO DE BARRAS', 'EAN', 'CÓDIGO', 'CODIGO']);
  const diasRestIdx = findCol(['DE HOJE ATÉ O VENCIMENTO', 'DE HOJE ATE O VENCIMENTO', 'DIAS ATÉ O VENCIMENTO']);
  const vendaMediaIdx = findCol(['VD MD DIA', 'VENDA MEDIA DIA', 'VENDA MÉDIA DIA', 'VD MD', 'VENDA MEDIA DIARIA']);
  const DIAS_CRITICO = 15;
  const hoje = new Date();

  const values = src
    .getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, lastCol)
    .getValues();

  const criticos = [];

  values.forEach((row, i) => {
    const temConteudo = row.some(c => c !== '' && c !== null);
    if (!temConteudo) return;

    const s = row[statusIdx] != null ? String(row[statusIdx]).trim() : '';
    const sLower = s.toLowerCase();
    if (CONFIG.STATUS_ACCEPTED.includes(sLower)) return;

    const rowNum = CONFIG.DATA_START_ROW + i;
    const validade = expIdx >= 0 ? parseDateFlexible(row[expIdx]) : null;

    // --- CORREÇÃO: IGUALAR LÓGICA DE CÁLCULO DE DIAS COM O DASHBOARD ---
    let diasRestantes = null;
    if (validade) {
      diasRestantes = diffInDays(validade, hoje); // Prioriza o cálculo pela data exata
    } else if (diasRestIdx >= 0) {
      const celVal = row[diasRestIdx];
      if (celVal !== '' && celVal != null) {
        const num = parseFloat(String(celVal).replace(/[^\d\-\.]/g, ''));
        if (!isNaN(num)) diasRestantes = num;
      }
    }

    if (diasRestantes === null || diasRestantes > DIAS_CRITICO || diasRestantes <= 0) return;

    criticos.push({
      rowNum,
      mercado: mercadoIdx >= 0 ? (row[mercadoIdx] || '').toString().trim() : '',
      codigo: codeIdx >= 0 ? (row[codeIdx] || '').toString().trim() : '',
      descricao: nameIdx >= 0 ? (row[nameIdx] || '').toString().trim() : '',
      quantidade: qtyIdx >= 0 ? row[qtyIdx] : '',
      vendaMediaDia: vendaMediaIdx >= 0 ? (parseFloat(row[vendaMediaIdx]) || 0) : 0,
      diasRestantes,
      validade,
      status: s
    });
  });

  // ── Aplica o Filtro de Mercado ─────────────────────────────────
  let listaFinal = criticos;
  if (filtroMercado && filtroMercado.trim() !== '') {
    const f = filtroMercado.trim().toLowerCase();
    listaFinal = criticos.filter(v => (v.mercado || '').toLowerCase() === f);
  }

  // ── Aplica o Filtro de Aba ─────────────────────────────────────
  let nomeAba = 'CRÍTICOS';
  if (filtroAba) {
    if (filtroAba === 'aguardar') {
      listaFinal = listaFinal.filter(v => (v.status || '').toUpperCase() === 'AGUARDAR');
      nomeAba = 'AGUARDAR';
    } else if (filtroAba === 'trocas') {
      listaFinal = listaFinal.filter(v => (v.status || '').toUpperCase() === 'TROCA');
      nomeAba = 'TROCAS';
    } else if (filtroAba === 'percas') {
      listaFinal = listaFinal.filter(v => (v.status || '').toUpperCase() === 'PERDA');
      nomeAba = 'PERCA';
    } else if (filtroAba === 'rebaixas') {
      listaFinal = listaFinal.filter(v => (v.status || '').toUpperCase() === 'REBAIXADO');
      nomeAba = 'REBAIXA';
    } else {
      listaFinal = listaFinal.filter(v => !v.status || v.status.trim() === '');
      nomeAba = 'CRÍTICOS';
    }
  }

  // ── Ordenação ──────────────────────────────────────────────────
  listaFinal.sort((a, b) => {
    const av = a.diasRestantes != null ? a.diasRestantes : Infinity;
    const bv = b.diasRestantes != null ? b.diasRestantes : Infinity;
    return av !== bv ? av - bv : a.rowNum - b.rowNum;
  });

  // ── Cria planilha temporária ───────────────────────────────────
  const prefixoFile = (nomeAba || 'Criticos');
  const nomeArquivo = filtroMercado && filtroMercado.trim() !== ''
    ? `${prefixoFile}_${filtroMercado.trim()}_${Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd-MM-yyyy')}`
    : `${prefixoFile}_Todos_${Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd-MM-yyyy')}`;

  const tmpSS = SpreadsheetApp.create(nomeArquivo);
  const sheet = tmpSS.getActiveSheet();
  sheet.setName(prefixoFile);

  // Fonte Padrão
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily('Arial');

  // ── Cabeçalho Principal (Estilo Limpo da Primeira Foto) ────────
  sheet.getRange('A1:I5').setBackground('#ffffff'); // Aumentamos o fundo branco até a linha 5

  // Pega o nome do mercado (ou define 'TODOS' caso não tenha filtro)
  const nomeSetor = filtroMercado && filtroMercado.trim() !== '' ? filtroMercado.trim().toUpperCase() : 'TODOS OS SETORES';

  // 1. Linha A1: Título Principal
  sheet.getRange('A1:I1').merge();
  sheet.getRange('A1')
    .setValue(`RELATÓRIO DE ${nomeAba}`)
    .setFontColor('#003366') // Azul escuro
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // 2. Linha A2: Nome do Mercadológico
  sheet.getRange('A2:I2').merge();
  sheet.getRange('A2')
    .setValue(`ツ  ${nomeSetor}`)
    .setFontColor('#003366') // Mesma cor do título, mas um pouco menor
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // 3. Linhas A3, A4 e A5: Textos Auxiliares
  sheet.getRange('A3:I3').merge();
  sheet.getRange('A4:I4').merge();
  sheet.getRange('A5:I5').merge(); // Nova linha mesclada para o total

  sheet.getRange('A3').setValue(`✦ Gerado em: ${Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}`);
  sheet.getRange('A4').setValue(`✦ Planilha origem: Auditoria de vencimento do Mercadinho`);
  sheet.getRange('A5').setValue(`✦ Total de itens: ${listaFinal.length}`);

  // Formatação dos textos auxiliares
  sheet.getRange('A3:A5')
    .setFontColor('#64748b')
    .setFontSize(10)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // Alinha as linhas 3 e 4 no meio normalmente
  sheet.getRange('A3:A4').setVerticalAlignment('middle');

  // ── O SEGREDO DO ESPAÇAMENTO ESTÁ AQUI 👇 ──
  // Aumentamos a altura da linha 5 e alinhamos o texto no topo.
  // Isso cria um "respiro" em branco antes da tabela começar!
  sheet.setRowHeight(5, 35);
  sheet.getRange('A5').setVerticalAlignment('top');

  // ── Cabeçalho da tabela ────────────────────────────────────────
  const cabecalho = ['Linha', 'Mercadológico', 'Cód. Barra', 'Produto', 'Quantidade', 'Vencimento', 'Dias até vencer', 'Urgência', 'Prob. de Venda', 'Status'];
  const headerRange = sheet.getRange(6, 1, 1, cabecalho.length);

  headerRange.setValues([cabecalho])
    .setBackground('#1e293b') // Fundo escuro igual da foto
    .setFontColor('#ffffff')  // Texto branco
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Borda branca fina no cabeçalho para separar as colunas (Exatamente como na foto)
  headerRange.setBorder(true, true, true, true, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(6, 22);

  // ── Dados ──────────────────────────────────────────────────────
  if (listaFinal.length > 0) {
    const dadosParaPlanilha = listaFinal.map(v => {
      const p = (v.diasRestantes !== null && v.quantidade > 0 && v.vendaMediaDia > 0)
        ? Math.min(100, Math.max(0, Math.round((v.diasRestantes * v.vendaMediaDia / v.quantidade) * 100)))
        : (v.vendaMediaDia === 0 ? 0 : null);

      const probTxt = p !== null ? (p + '%') : '—';
      const urgenciaTxt = v.diasRestantes <= 3 ? 'URGENTE' : v.diasRestantes <= 7 ? 'ALTA' : 'MÉDIA';
      const valTxt = v.validade ? Utilities.formatDate(v.validade, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '—';

      return [
        v.rowNum,
        v.mercado || '—',
        v.codigo || '—',
        v.descricao || '—',
        v.quantidade !== '' ? v.quantidade : '—',
        valTxt,
        v.diasRestantes !== null ? (v.diasRestantes <= 0 ? 'VENCIDO' : v.diasRestantes) : '—',
        urgenciaTxt,
        probTxt,
        v.status || '—'
      ];
    });

    const dataRange = sheet.getRange(7, 1, dadosParaPlanilha.length, cabecalho.length);
    dataRange.setValues(dadosParaPlanilha);
    dataRange.setFontSize(9).setFontColor('#1e293b').setVerticalAlignment('middle');

    listaFinal.forEach((v, idx) => {
      const linha = 7 + idx;
      const d = v.diasRestantes;
      const corTxt = d <= 3 ? '#b91c1c' : d <= 7 ? '#c2410c' : '#92400e';

      // Cor de fundo alternada bem discreta
      const bgCor = (idx % 2 === 0) ? '#ffffff' : '#f8fafc';
      sheet.getRange(linha, 1, 1, cabecalho.length).setBackground(bgCor);

      // Cores específicas (Urgência e Dias)
      sheet.getRange(linha, 8).setFontWeight('bold').setFontColor(corTxt);
      sheet.getRange(linha, 7).setFontWeight('bold').setFontColor(corTxt);

      // Alinhamentos
      sheet.getRange(linha, 1).setHorizontalAlignment('center'); // Linha
      sheet.getRange(linha, 2).setHorizontalAlignment('left');   // Mercado
      sheet.getRange(linha, 3).setHorizontalAlignment('center'); // Cód Barra
      sheet.getRange(linha, 4).setHorizontalAlignment('left');   // Produto
      sheet.getRange(linha, 5).setHorizontalAlignment('center'); // Qtd
      sheet.getRange(linha, 6).setHorizontalAlignment('center'); // Venc
      sheet.getRange(linha, 7).setHorizontalAlignment('center'); // Dias
      sheet.getRange(linha, 8).setHorizontalAlignment('center'); // Urgência

      const prob = (d !== null && v.quantidade > 0 && v.vendaMediaDia > 0)
        ? Math.min(100, Math.max(0, Math.round((d * v.vendaMediaDia / v.quantidade) * 100)))
        : (v.vendaMediaDia === 0 ? 0 : null);

      // --- CORREÇÃO: LÓGICA DE CORES EXCEL ---
      if (prob !== null) {
        const corProb = prob >= 70 ? '#15803d' // Verde (Alta probabilidade)
          : prob >= 40 ? '#ca8a04' // Amarelo (Média)
            : '#b91c1c';             // Vermelho (Baixa)

        sheet.getRange(linha, 9)
          .setFontWeight('bold')
          .setFontColor(corProb)
          .setHorizontalAlignment('center');
      }

      sheet.setRowHeight(linha, 18);
    });

    // Borda geral da tabela (Cinza claro, separando tudo certinho)
    dataRange.setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Dimensionamento das colunas
  sheet.setColumnWidth(1, 40);   // Linha
  sheet.setColumnWidth(2, 125);  // Mercadológico
  sheet.setColumnWidth(3, 105);  // Cód. Barra
  sheet.setColumnWidth(4, 350);  // Produto (COMPACTADO PARA CABER)
  sheet.setColumnWidth(5, 75);   // Quantidade
  sheet.setColumnWidth(6, 85);   // Vencimento
  sheet.setColumnWidth(7, 105);   // Dias
  sheet.setColumnWidth(8, 80);   // Urgência
  sheet.setColumnWidth(9, 100);  // Prob. de Venda
  sheet.setColumnWidth(10, 85);  // Status

  // Congela cabeçalho
  sheet.setFrozenRows(6);

  // Garante que a planilha seja salva no drive antes de gerar o download
  SpreadsheetApp.flush();

  // ── Gera URLs de download ──────────────────────────────────────
  const fileId = tmpSS.getId();
  const sheetId = sheet.getSheetId();
  const urlXlsx = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx&id=${fileId}`;

  // Exporta removendo linhas de grade padrão (nossas bordas cuidarão do visual)
  const urlPdf = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=pdf&size=A4&portrait=false&fitw=true&gridlines=false&gid=${sheetId}`;

  try {
    DriveApp.getFileById(fileId).setTrashed(false);
  } catch (e) { }
  return { xlsx: urlXlsx, pdf: urlPdf, nome: nomeArquivo };
}

function gerarRelatorioAnalise() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SRC_SHEET);
  const ui = SpreadsheetApp.getUi();

  if (!src) return ui.alert(`Aba "${CONFIG.SRC_SHEET}" não encontrada.`);

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW) return ui.alert('Não há dados para analisar.');

  // ── Índices por letra configurada ──────────────────────────────
  const statusIdx = STATUS_COL_INDEX;
  const expIdx = getColIndex(CONFIG.EXPIRY_DATE_COL_LETTER);
  const nameIdx = getColIndex(CONFIG.PROD_NAME_COL_LETTER);
  const qtyIdx = getColIndex(CONFIG.QUANTITY_COL_LETTER);
  const loteIdx = getColIndex(CONFIG.LOTE_COL_LETTER);
  const priceIdx = getColIndex(CONFIG.PRICE_COL_LETTER);

  // ── Descobrir colunas pelo cabeçalho ───────────────────────────
  const headerValues = src
    .getRange(CONFIG.HEADER_ROWS, 1, 1, lastCol)
    .getValues()[0]
    .map(h => (h || '').toString().trim().toUpperCase());

  function findCol(nome) {
    const i = headerValues.indexOf(nome.toUpperCase());
    return i >= 0 ? i : -1; // 0-based
  }

  const mercadoIdx = findCol('MERCADOLÓGICO');

  const diasRestCandidatos = [
    'DE HOJE ATÉ O VENCIMENTO',
    'DIAS ATÉ O VENCIMENTO',
    'DE HOJE ATÉ O VENCIMENTO (DIAS)'
  ];
  let diasRestIdx = -1;
  for (const n of diasRestCandidatos) {
    diasRestIdx = findCol(n);
    if (diasRestIdx >= 0) break;
  }

  // ── Limiar crítico ─────────────────────────────────────────────
  const DIAS_CRITICO = 15;

  // ── Leitura dos dados ──────────────────────────────────────────
  const values = src
    .getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, lastCol)
    .getValues();

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0); // Normalizar para comparar apenas datas
  let total = 0, semStatus = 0;
  const vencidos = [];
  const criticos = [];
  let minVal = null, maxVal = null;
  let somaDiasVenc = 0, contaDiasVenc = 0;
  let somaDiasCrit = 0, contaDiasCrit = 0;
  const agingBuckets = { '0-7': 0, '8-30': 0, '31-90': 0, '90+': 0 };
  const mercadoVencidos = {};  // { cat: { qtd, itens } }
  const mercadoCriticos = {};  // { cat: { qtd, itens } }

  values.forEach((row, i) => {
    const temConteudo = row.some(c => c !== '' && c !== null);
    if (!temConteudo) return;

    total++;
    const s = row[statusIdx] != null ? String(row[statusIdx]).trim() : '';
    const sLower = s.toLowerCase();
    if (!s) { semStatus++; }

    const rowNum = CONFIG.DATA_START_ROW + i;
    const validade = expIdx >= 0 ? parseDateFlexible(row[expIdx]) : null;

    // calcular dias restantes (de hoje até o vencimento)
    let diasRestantes = null;
    if (validade) {
      diasRestantes = diffInDays(validade, hoje);
    } else if (diasRestIdx >= 0) {
      const celVal = row[diasRestIdx];
      if (celVal !== '' && celVal != null) {
        const numerico = parseFloat(String(celVal).replace(/[^\d\-\.]/g, ''));
        if (!isNaN(numerico)) diasRestantes = numerico;
      }
    }

    const diasVencidos = validade ? diffInDays(hoje, validade) : null;
    const quantidade = qtyIdx >= 0 ? Number(row[qtyIdx]) || 0 : 0;
    const mercado = mercadoIdx >= 0 ? (row[mercadoIdx] || '').toString().trim() : '';
    const cat = mercado || 'SEM CATEGORIA';

    const obj = {
      rowNum,
      mercado,
      descricao: nameIdx >= 0 ? (row[nameIdx] || '').toString().trim() : '',
      lote: loteIdx >= 0 ? (row[loteIdx] || '').toString().trim() : '',
      quantidade,
      validade,
      diasRestantes,
      diasVencidos,
      status: s,
      preco: priceIdx >= 0 ? row[priceIdx] : ''
    };

    if (CONFIG.STATUS_ACCEPTED.includes(sLower) || (diasRestantes !== null && diasRestantes <= 0)) {
      // ── bloco vencidos ──
      vencidos.push(obj);

      if (validade) {
        if (!minVal || validade < minVal) minVal = validade;
        if (!maxVal || validade > maxVal) maxVal = validade;
      }
      if (diasVencidos != null) {
        somaDiasVenc += diasVencidos;
        contaDiasVenc++;
        if (diasVencidos <= 7) agingBuckets['0-7']++;
        else if (diasVencidos <= 30) agingBuckets['8-30']++;
        else if (diasVencidos <= 90) agingBuckets['31-90']++;
        else agingBuckets['90+']++;
      }

      if (!mercadoVencidos[cat]) mercadoVencidos[cat] = { qtd: 0, itens: 0 };
      mercadoVencidos[cat].qtd += quantidade;
      mercadoVencidos[cat].itens += 1;

    } else if (diasRestantes !== null && diasRestantes <= DIAS_CRITICO) {
      // ── bloco críticos ──
      criticos.push(obj);

      if (diasRestantes != null) {
        somaDiasCrit += diasRestantes;
        contaDiasCrit++;
      }

      if (!mercadoCriticos[cat]) mercadoCriticos[cat] = { qtd: 0, itens: 0 };
      mercadoCriticos[cat].qtd += quantidade;
      mercadoCriticos[cat].itens += 1;
    }
  });

  // ── Ordenar por dias crescente ─────────────────────────────────
  const sortDias = (a, b) => {
    const av = a.diasRestantes != null ? a.diasRestantes : Infinity;
    const bv = b.diasRestantes != null ? b.diasRestantes : Infinity;
    return av !== bv ? av - bv : a.rowNum - b.rowNum;
  };
  vencidos.sort(sortDias);
  criticos.sort(sortDias);

  // ── Montar aba ANALISE_VENCIDOS ────────────────────────────────
  const analise = ensureSheet(ss, CONFIG.ANALYSIS_SHEET);
  analise.clear();

  let linha = 1;

  // ── BLOCO 1: RESUMO GERAL ──────────────────────────────────────
  const mediaDiasVenc = contaDiasVenc ? Math.round(somaDiasVenc / contaDiasVenc) : '-';
  const mediaDiasCrit = contaDiasCrit ? Math.round(somaDiasCrit / contaDiasCrit) : '-';

  const resumo = [
    ['Resumo Geral', 'Valor'],
    ['Data da análise', formatDateBR(new Date())],
    ['Total analisado', total],
    ['Sem status', semStatus],
    [`⚠️ Críticos `, criticos.length],
    ['🔴 Vencidos', vencidos.length],
    ['Menor validade (vencidos)', minVal ? formatDateBR(minVal) : '-'],
    ['Maior validade (vencidos)', maxVal ? formatDateBR(maxVal) : '-'],
    ['Média de dias vencidos', mediaDiasVenc],
    ['Média de dias restantes (críticos)', mediaDiasCrit]
  ];
  setTable(analise, linha, 1, resumo, true);
  linha += resumo.length + 2;

  // ── BLOCO 2: FAIXAS DE VENCIMENTO ─────────────────────────────
  if (contaDiasVenc > 0) {
    const agingTable = [
      ['Faixa (dias vencidos)', 'Quantidade de itens'],
      ['0-7', agingBuckets['0-7']],
      ['8-30', agingBuckets['8-30']],
      ['31-90', agingBuckets['31-90']],
      ['90+', agingBuckets['90+']]
    ];
    setTable(analise, linha, 1, [['Faixas de vencimento (crescente)']], true);
    setTable(analise, linha + 1, 1, agingTable, true);
    linha += agingTable.length + 3;
  }

  // ── BLOCO 3: VISÃO MERCADOLÓGICA – CRÍTICOS ───────────────────
  setTable(analise, linha, 1, [[`⚠️ Visão Mercadológica – Críticas `]], true);
  setTable(analise, linha + 1, 1,
    [['Mercadológico', 'Quantidade total', 'Nº de itens']], true);
  const mercadoCritTabela = Object.keys(mercadoCriticos)
    .sort((a, b) => a.localeCompare(b, 'pt-BR'))
    .map(cat => [cat, mercadoCriticos[cat].qtd, mercadoCriticos[cat].itens]);
  if (mercadoCritTabela.length) {
    setTable(analise, linha + 2, 1, mercadoCritTabela, false);
    linha += mercadoCritTabela.length + 4;
  } else {
    setTable(analise, linha + 2, 1, [['Nenhum produto crítico']], false);
    linha += 5;
  }

  // ── BLOCO 4: VISÃO MERCADOLÓGICA – VENCIDOS ───────────────────
  setTable(analise, linha, 1, [['🔴 Visão Mercadológica – Vencidos']], true);
  setTable(analise, linha + 1, 1,
    [['Mercadológico', 'Quantidade total', 'Nº de itens']], true);
  const mercadoVencTabela = Object.keys(mercadoVencidos)
    .sort((a, b) => a.localeCompare(b, 'pt-BR'))
    .map(cat => [cat, mercadoVencidos[cat].qtd, mercadoVencidos[cat].itens]);
  if (mercadoVencTabela.length) {
    setTable(analise, linha + 2, 1, mercadoVencTabela, false);
    linha += mercadoVencTabela.length + 4;
  } else {
    setTable(analise, linha + 2, 1, [['Nenhum vencido']], false);
    linha += 5;
  }

  // ── BLOCO 5: DETALHE CRÍTICOS ─────────────────────────────────
  const cabCab = [
    'Linha (Matriz)',
    'Mercadológico',
    'Produto',
    'Quantidade',
    'Vencimento',
    'De hoje até o vencimento (dias)',
    'Lote',
    'Status'
  ];

  const detalheCriticos = [cabCab].concat(
    criticos.map(v => [
      v.rowNum,
      v.mercado || '',
      v.descricao || '',
      v.quantidade || '',
      v.validade ? formatDateBR(v.validade) : '',
      v.diasRestantes != null ? v.diasRestantes : '',
      v.lote || '',
      v.status || ''
    ])
  );

  setTable(analise, linha, 1,
    [[`⚠️ Críticos – vencem em até ${DIAS_CRITICO} dias (crescente)`]], true);
  setTable(analise, linha + 1, 1, detalheCriticos, true);
  linha += detalheCriticos.length + 3;

  // ── BLOCO 6: DETALHE VENCIDOS ─────────────────────────────────
  const detalheVencidos = [cabCab].concat(
    vencidos.map(v => [
      v.rowNum,
      v.mercado || '',
      v.descricao || '',
      v.quantidade || '',
      v.validade ? formatDateBR(v.validade) : '',
      v.diasRestantes != null ? v.diasRestantes : '',
      v.lote || '',
      v.status || ''
    ])
  );

  setTable(analise, linha, 1,
    [['🔴 Vencidos – detalhe (crescente por validade)']], true);
  setTable(analise, linha + 1, 1, detalheVencidos, true);

  // ── Finalizar ──────────────────────────────────────────────────
  analise.setFrozenRows(1);
  SpreadsheetApp.flush();
  ui.alert(`Relatório gerado na aba "${CONFIG.ANALYSIS_SHEET}".`);
}

function filtrarEExportarAnalise() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SRC_SHEET);

  if (!src) return ui.alert(`Aba "${CONFIG.SRC_SHEET}" não encontrada.`);

  const resp = ui.prompt(
    'Filtro de Análise',
    'Informe um status (ex.: vencido) e opcionalmente um intervalo (dd/mm/aaaa - dd/mm/aaaa).\n' +
    'Ex.:\n' +
    'vencido\n' +
    'vencido 01/01/2024 - 31/03/2024',
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const texto = (resp.getResponseText() || '').trim();
  if (!texto) return ui.alert('Entrada vazia.');

  const m = texto.match(/^([^\d]+)\s*(\d{1,2}\/\d{1,2}\/\d{4})?\s*-\s*(\d{1,2}\/\d{1,2}\/\d{4})?$/);
  let statusFiltro = texto, dIni = null, dFim = null;

  if (m) {
    statusFiltro = m[1].trim();
    dIni = m[2] ? parseDateFlexible(m[2]) : null;
    dFim = m[3] ? parseDateFlexible(m[3]) : null;
  }

  const statusFiltroLower = statusFiltro.toLowerCase();

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW) return ui.alert('Não há dados.');

  const statusIdx = STATUS_COL_INDEX;
  const expIdx = getColIndex(CONFIG.EXPIRY_DATE_COL_LETTER);
  const values = src.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, lastCol).getValues();

  const linhas = [];

  values.forEach((row, i) => {
    const temConteudo = row.some(c => c !== '' && c !== null);
    if (!temConteudo) return;

    const s = row[statusIdx] != null ? String(row[statusIdx]).trim() : '';
    if (!s) return;
    if (s.toLowerCase().indexOf(statusFiltroLower) === -1) return;

    let okData = true;
    let validade = null;

    if (expIdx >= 0) {
      validade = parseDateFlexible(row[expIdx]);
      if (dIni && validade && validade < dIni) okData = false;
      if (dFim && validade && validade > dFim) okData = false;
    }

    if (!okData) return;

    linhas.push({ rowNum: CONFIG.DATA_START_ROW + i, validade, row });
  });

  if (expIdx >= 0) {
    linhas.sort((a, b) => {
      const ad = a.validade ? a.validade.getTime() : Infinity;
      const bd = b.validade ? b.validade.getTime() : Infinity;
      return ad !== bd ? ad - bd : a.rowNum - b.rowNum;
    });
  } else {
    linhas.sort((a, b) => a.rowNum - b.rowNum);
  }

  const outSheet = ensureSheet(ss, CONFIG.FILTER_SHEET);
  outSheet.clear();

  const headerRange = src.getRange(1, 1, CONFIG.HEADER_ROWS, lastCol);
  headerRange.copyTo(outSheet.getRange(1, 1), { formatOnly: false, contentsOnly: false });

  const dados = linhas.map(o => o.row);
  if (dados.length) {
    outSheet.getRange(CONFIG.DATA_START_ROW, 1, dados.length, lastCol).setValues(dados);
  }

  outSheet.setFrozenRows(CONFIG.HEADER_ROWS);

  ui.alert(`Filtrado ${dados.length} registro(s) para "${CONFIG.FILTER_SHEET}" em ordem crescente.`);
}

function encontrarDuplicidades() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SRC_SHEET);

  if (!src) return ui.alert(`Aba "${CONFIG.SRC_SHEET}" não encontrada.`);

  const keyCols = CONFIG.KEY_COLS || [];
  if (!keyCols.length) return ui.alert('Defina CONFIG.KEY_COLS com as letras das colunas (ex.: ["B","C"]).');

  const idxs = keyCols.map(getColIndex).filter(i => i >= 0);
  if (!idxs.length) return ui.alert('CONFIG.KEY_COLS inválida.');

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW) return ui.alert('Não há dados.');

  const values = src.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, lastCol).getValues();
  const mapa = new Map();

  values.forEach((row, i) => {
    const temConteudo = row.some(c => c !== '' && c !== null);
    if (!temConteudo) return;

    const chave = idxs.map(ix => row[ix] != null ? String(row[ix]).trim() : '').join(' | ');
    if (!chave) return;

    const info = mapa.get(chave) || { count: 0, linhas: [] };
    info.count++;
    info.linhas.push(CONFIG.DATA_START_ROW + i);
    mapa.set(chave, info);
  });

  const dups = Array.from(mapa.entries())
    .filter(([, v]) => v.count > 1)
    .map(([k, v]) => ({ k, count: v.count, linhas: v.linhas.sort((a, b) => a - b) }))
    .sort((a, b) => a.k.localeCompare(b, 'pt-BR'));

  if (!dups.length) return ui.alert('Nenhuma duplicidade encontrada.');

  const analise = ensureSheet(ss, CONFIG.ANALYSIS_SHEET);
  const startRow = Math.max(analise.getLastRow() + 2, 1);

  setTable(analise, startRow, 1, [['Duplicidades (ordem crescente)']], true);

  const dados = [['Chave', 'Ocorrências', 'Linhas (Matriz)']].concat(
    dups.map(d => [d.k, d.count, d.linhas.join(', ')])
  );

  setTable(analise, startRow + 1, 1, dados, true);

  ui.alert(`Encontradas ${dups.length} chaves duplicadas. Detalhes em "${CONFIG.ANALYSIS_SHEET}".`);
}


/*
 * Sistema 1000 linhas
 * ============================================================ */

// Configurando o sistema de 1000 linhas
// ------------------------------------------------------------
var NOME_DA_ABA_ALVO = 'Matriz'; // <--- COLOQUE O NOME DA ABA AQUI
var PROPRIEDADE_LINHAS = 'numeroLinhasAnterior_' + NOME_DA_ABA_ALVO.replace(/\s/g, '');
var LIMITE_MAXIMO_LINHAS = 1000;
var LINHA_MODELO_PADRAO = 6;
var CONFIG_LINHAS = {
  TRIGGERS: {
    MONITOR_EVERY_MINUTES: 5 // Minutos
  }
};


// Função para  auxiliar
// ------------------------------------------------------------
function pegarAbaAlvo() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName(NOME_DA_ABA_ALVO);

  if (!aba) {
    console.warn('⚠️ A aba configurada "' + NOME_DA_ABA_ALVO + '" não foi encontrada.');
    return null;
  }
  return aba;
}


// A lógica do sistema
// ------------------------------------------------------------
function verificarEAplicarLimite() {
  try {
    var aba = pegarAbaAlvo();
    if (!aba) return false;

    var linhasAtuais = aba.getLastRow();

    if (linhasAtuais > LIMITE_MAXIMO_LINHAS) {
      var linhasParaRemover = linhasAtuais - LIMITE_MAXIMO_LINHAS;
      console.log('🚨 LIMITE EXCEDIDO na aba "' + NOME_DA_ABA_ALVO + '"! Removendo ' + linhasParaRemover + ' linha(s).');
      aba.deleteRows(LIMITE_MAXIMO_LINHAS + 1, linhasParaRemover);

      PropertiesService.getScriptProperties().setProperty(PROPRIEDADE_LINHAS, LIMITE_MAXIMO_LINHAS.toString());
      console.log('✅ Limite restaurado: ' + aba.getLastRow() + ' linhas');
      return true;
    }
    return false;
  } catch (erro) {
    console.error('❌ ERRO CRÍTICO na verificação de limite: ' + erro.toString());
    return false;
  }
}

function garantirFormulasCompletas() {
  try {
    var aba = pegarAbaAlvo();
    if (!aba) return false;

    var linhasAtuais = aba.getLastRow();
    var ultimaColuna = aba.getLastColumn();

    console.log('🔍 INICIANDO VERIFICAÇÃO DE FÓRMULAS NA ABA: ' + NOME_DA_ABA_ALVO);

    if (linhasAtuais < LINHA_MODELO_PADRAO) {
      console.log('⚠️ Não há linha modelo suficiente para copiar fórmulas');
      return false;
    }

    var intervaloModelo = aba.getRange(LINHA_MODELO_PADRAO, 1, 1, ultimaColuna);
    var formulasModelo = intervaloModelo.getFormulas()[0];

    var linhasParaCompletar = LIMITE_MAXIMO_LINHAS - linhasAtuais;

    if (linhasParaCompletar > 0) {
      console.log('📝 Completando ' + linhasParaCompletar + ' linha(s) com fórmulas...');
      aba.insertRowsAfter(linhasAtuais, linhasParaCompletar);

      for (var i = linhasAtuais + 1; i <= LIMITE_MAXIMO_LINHAS; i++) {
        intervaloModelo.copyFormatToRange(aba, 1, ultimaColuna, i, i);

        for (var col = 1; col <= ultimaColuna; col++) {
          var formula = formulasModelo[col - 1];
          if (formula && formula !== '') {
            var formulaAjustada = ajustarFormulaParaNovaLinha(formula, LINHA_MODELO_PADRAO, i);
            aba.getRange(i, col).setFormula(formulaAjustada);
          } else {
            aba.getRange(i, col).setValue('');
          }

          var validacao = aba.getRange(LINHA_MODELO_PADRAO, col).getDataValidation();
          if (validacao) aba.getRange(i, col).setDataValidation(validacao);

          var formatoNumero = aba.getRange(LINHA_MODELO_PADRAO, col).getNumberFormat();
          if (formatoNumero) aba.getRange(i, col).setNumberFormat(formatoNumero);
        }

        if (i % 50 === 0) {
          console.log('  ⚡️ Processadas ' + i + ' de ' + LIMITE_MAXIMO_LINHAS + ' linhas...');
        }
      }

      console.log('✅ TODAS AS 1000 LINHAS AGORA TÊM FÓRMULAS!');
      return true;
    } else {
      console.log('✓ Já existem ' + linhasAtuais + ' linhas; verificando integridade...');
      verificarIntegridadeFormulas();
      return false;
    }
  } catch (erro) {
    console.error('❌ Erro ao garantir fórmulas: ' + erro.toString());
    return false;
  }
}

function verificarIntegridadeFormulas() {
  try {
    var aba = pegarAbaAlvo();
    if (!aba) return;

    var linhasAtuais = Math.min(aba.getLastRow(), LIMITE_MAXIMO_LINHAS);
    var ultimaColuna = aba.getLastColumn();

    console.log('🔍 Verificando integridade das fórmulas...');

    var formulasModelo = aba.getRange(LINHA_MODELO_PADRAO, 1, 1, ultimaColuna).getFormulas()[0];
    var colunasComFormula = [];

    for (var c = 0; c < formulasModelo.length; c++) {
      if (formulasModelo[c] && formulasModelo[c] !== '') {
        colunasComFormula.push(c + 1);
      }
    }

    if (!colunasComFormula.length) {
      console.log('ℹ️ Nenhuma fórmula na linha modelo');
      return;
    }

    var linhasCorrigidas = 0;

    for (var linha = LINHA_MODELO_PADRAO; linha <= linhasAtuais; linha++) {
      var necessita = false;

      for (var j = 0; j < colunasComFormula.length; j++) {
        var col = colunasComFormula[j];
        var celula = aba.getRange(linha, col);
        var formulaAtual = celula.getFormula();

        if (!formulaAtual || formulaAtual === '') {
          necessita = true;
          var formulaOriginal = formulasModelo[col - 1];
          var formulaAjustada = ajustarFormulaParaNovaLinha(formulaOriginal, LINHA_MODELO_PADRAO, linha);
          celula.setFormula(formulaAjustada);
        }
      }

      if (necessita) {
        linhasCorrigidas++;
        console.log('  🔧 Linha ' + linha + ' corrigida');
      }
    }

    if (linhasCorrigidas > 0) {
      console.log('✅ Integridade restaurada: ' + linhasCorrigidas + ' linha(s) corrigida(s)');
    } else {
      console.log('✓ Todas as linhas já possuem fórmulas corretas');
    }
  } catch (erro) {
    console.error('❌ Erro na verificação de integridade: ' + erro.toString());
  }
}

function onChangeHandler(e) {
  // Nota: O evento onChange dispara para qualquer mudança na planilha.
  // Nós forçamos a verificação APENAS na aba alvo, independente de onde a edição ocorreu.
  try {
    var aba = pegarAbaAlvo();
    if (!aba) return; // Se a aba alvo não existe, não faz nada.

    var houveLimpeza = verificarEAplicarLimite();
    if (houveLimpeza) {
      console.log('Limpeza realizada. Rechecar fórmulas...');
      garantirFormulasCompletas();
      return;
    }

    var linhasAtuais = aba.getLastRow();

    if (linhasAtuais < LIMITE_MAXIMO_LINHAS) {
      console.log('⚠️ Menos de ' + LIMITE_MAXIMO_LINHAS + ' linhas na aba alvo. Completando...');
      garantirFormulasCompletas();
      PropertiesService.getScriptProperties().setProperty(PROPRIEDADE_LINHAS, LIMITE_MAXIMO_LINHAS.toString());
      return;
    }

    var propriedades = PropertiesService.getScriptProperties();
    var linhasAnteriores = propriedades.getProperty(PROPRIEDADE_LINHAS);

    if (!linhasAnteriores) {
      garantirFormulasCompletas();
      propriedades.setProperty(PROPRIEDADE_LINHAS, LIMITE_MAXIMO_LINHAS.toString());
      console.log('Primeira execução – garantindo 1000 linhas com fórmulas');
      return;
    }

    linhasAnteriores = parseInt(linhasAnteriores, 10);

    if (linhasAtuais < linhasAnteriores) {
      var linhasExcluidas = linhasAnteriores - linhasAtuais;
      console.log('📉 Exclusão detectada: ' + linhasExcluidas + ' linha(s). Recompondo...');
      garantirFormulasCompletas();
    }

    verificarIntegridadeFormulas();
    propriedades.setProperty(PROPRIEDADE_LINHAS, LIMITE_MAXIMO_LINHAS.toString());
    console.log('✅ Mantendo ' + LIMITE_MAXIMO_LINHAS + ' linhas com fórmulas');
  } catch (erro) {
    console.error('❌ Erro no onChangeHandler: ' + erro.toString());
    verificarEAplicarLimite();
    garantirFormulasCompletas();
  }
}

function ajustarFormulaParaNovaLinha(formula, linhaOrigem, linhaDestino) {
  try {
    if (!formula || formula === '') return '';
    var padraoReferencia = /(\$?)([A-Z]+)(\$?)(\d+)/g;
    var formulaAjustada = formula.replace(padraoReferencia, function (match, dolarColuna, coluna, dolarLinha, linha) {
      var numeroLinha = parseInt(linha, 10);
      if (dolarLinha === '$') return match; // não altera se a linha estiver absoluta
      if (numeroLinha === linhaOrigem) {
        return (dolarColuna || '') + coluna + (dolarLinha || '') + linhaDestino;
      }
      return match;
    });
    return formulaAjustada;
  } catch (erro) {
    console.error('❌ Erro ao ajustar fórmula: ' + erro.toString());
    return formula;
  }
}

function monitoramentoContinuo() {
  try {
    var aba = pegarAbaAlvo();
    if (!aba) return;

    var linhasAtuais = aba.getLastRow();

    console.log('⏰ MONITORAMENTO CONTÍNUO NA ABA: ' + NOME_DA_ABA_ALVO);

    if (linhasAtuais > LIMITE_MAXIMO_LINHAS) {
      console.log('🚨 Limite excedido (' + linhasAtuais + ' linhas)');
      verificarEAplicarLimite();
      console.log('🔧 Limite corrigido');
    }

    if (linhasAtuais < LIMITE_MAXIMO_LINHAS) {
      console.log('📝 Completando para 1000 linhas');
      garantirFormulasCompletas();
    }

    verificarIntegridadeFormulas();
    console.log('✅ Verificação concluída – 1000 linhas garantidas');
  } catch (erro) {
    console.error('❌ Erro no monitoramento contínuo: ' + erro.toString());
  }
}

function criarTriggerDeMonitoramento() {
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === 'monitoramentoContinuo') {
        ScriptApp.deleteTrigger(t);
      }
    });

    ScriptApp.newTrigger('monitoramentoContinuo')
      .timeBased()
      .everyMinutes(CONFIG.TRIGGERS.MONITOR_EVERY_MINUTES)
      .create();

    console.log('✅ Trigger de monitoramento contínuo criado');
  } catch (erro) {
    console.error('❌ Erro ao criar trigger: ' + erro.toString());
  }
}


// Funções da interface
// ------------------------------------------------------------

function configurarTriggerAutomatico() {
  try {
    var ui = SpreadsheetApp.getUi();
    var aba = pegarAbaAlvo();

    if (!aba) {
      ui.alert('❌ Erro', 'A aba configurada no script ("' + NOME_DA_ABA_ALVO + '") não existe!', ui.ButtonSet.OK);
      return;
    }

    if (aba.getLastRow() < LINHA_MODELO_PADRAO) {
      ui.alert(
        '⚠️ Atenção',
        'A aba "' + NOME_DA_ABA_ALVO + '" precisa ter pelo menos ' + LINHA_MODELO_PADRAO + ' linhas.\n' +
        'A linha ' + LINHA_MODELO_PADRAO + ' será usada como modelo.',
        ui.ButtonSet.OK
      );
      return;
    }

    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function (t) {
      if (t.getHandlerFunction() === 'onChangeHandler') {
        ScriptApp.deleteTrigger(t);
      }
    });

    ScriptApp.newTrigger('onChangeHandler')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onChange()
      .create();

    criarTriggerDeMonitoramento();

    console.log('📋 Iniciando configuração do sistema...');
    var houveLimpeza = verificarEAplicarLimite();
    var formulasAplicadas = garantirFormulasCompletas();

    PropertiesService.getScriptProperties().setProperty(PROPRIEDADE_LINHAS, LIMITE_MAXIMO_LINHAS.toString());

    var msg = '✅ SISTEMA CONFIGURADO COM SUCESSO!\n\n' +
      '📌 ABA ALVO: ' + NOME_DA_ABA_ALVO + '\n' +
      '🛡️ PROTEÇÕES ATIVAS\n' +
      '• Trigger onChange: ✓ Ativo\n' +
      '• Monitoramento contínuo: ✓ A cada ' + CONFIG.TRIGGERS.MONITOR_EVERY_MINUTES + ' min\n' +
      '• Limite máximo: 🔒 ' + LIMITE_MAXIMO_LINHAS + ' linhas\n' +
      '• Linha modelo: Linha ' + LINHA_MODELO_PADRAO;

    ui.alert('🛡️ Sistema 1000 Linhas (Aba Única)', msg, ui.ButtonSet.OK);
  } catch (erro) {
    SpreadsheetApp.getUi().alert(
      '❌ Erro',
      'Erro ao configurar sistema: ' + erro.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function forcarMilLinhasComFormulas() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    '🔧 Forçar 1000 Linhas na Aba: ' + NOME_DA_ABA_ALVO,
    'Esta função irá processar APENAS a aba "' + NOME_DA_ABA_ALVO + '".\n\n' +
    '• Remover linhas extras (> 1000)\n' +
    '• Adicionar linhas se houver menos\n' +
    '• Garantir fórmulas (modelo linha ' + LINHA_MODELO_PADRAO + ')\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) return;

  try {
    var aba = pegarAbaAlvo();
    if (!aba) {
      ui.alert('Erro', 'Aba "' + NOME_DA_ABA_ALVO + '" não encontrada.', ui.ButtonSet.OK);
      return;
    }

    var linhasAntes = aba.getLastRow();

    verificarEAplicarLimite();
    garantirFormulasCompletas();
    verificarIntegridadeFormulas();

    var linhasDepois = aba.getLastRow();
    PropertiesService.getScriptProperties().setProperty(PROPRIEDADE_LINHAS, LIMITE_MAXIMO_LINHAS.toString());

    ui.alert(
      '✅ Sucesso!',
      'PROCESSO CONCLUÍDO NA ABA: ' + NOME_DA_ABA_ALVO + '\n\n' +
      '📊 Linhas antes: ' + linhasAntes + '\n' +
      '📊 Linhas depois: ' + linhasDepois,
      ui.ButtonSet.OK
    );
  } catch (erro) {
    ui.alert('❌ Erro', 'Erro durante o processo: ' + erro.toString(), ui.ButtonSet.OK);
  }
}

function verificarStatus() {
  var aba = pegarAbaAlvo();
  var ui = SpreadsheetApp.getUi();

  if (!aba) {
    ui.alert('Erro', 'Aba "' + NOME_DA_ABA_ALVO + '" não encontrada.', ui.ButtonSet.OK);
    return;
  }

  verificarEAplicarLimite();

  var propriedades = PropertiesService.getScriptProperties();
  var linhasArmazenadas = propriedades.getProperty(PROPRIEDADE_LINHAS);
  var linhasAtuais = aba.getLastRow();
  var ultimaColuna = aba.getLastColumn();

  var linhasComFormulas = 0;
  var linhasSemFormulas = [];

  if (linhasAtuais >= LINHA_MODELO_PADRAO && ultimaColuna > 0) {
    var formulasModelo = aba.getRange(LINHA_MODELO_PADRAO, 1, 1, ultimaColuna).getFormulas()[0];
    var colunasQueDevemTerFormula = 0;

    for (var c = 0; c < formulasModelo.length; c++) {
      if (formulasModelo[c] && formulasModelo[c] !== '') colunasQueDevemTerFormula++;
    }

    if (colunasQueDevemTerFormula > 0) {
      for (var i = LINHA_MODELO_PADRAO; i <= Math.min(linhasAtuais, LIMITE_MAXIMO_LINHAS); i++) {
        var formulas = aba.getRange(i, 1, 1, ultimaColuna).getFormulas()[0];
        var temFormula = formulas.some(function (f) { return f && f !== ''; });

        if (temFormula) {
          linhasComFormulas++;
        } else {
          linhasSemFormulas.push(i);
        }
      }
    }
  }

  var triggers = ScriptApp.getProjectTriggers();
  var triggerOnChange = false, triggerMonitoramento = false;

  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'onChangeHandler') triggerOnChange = true;
    if (t.getHandlerFunction() === 'monitoramentoContinuo') triggerMonitoramento = true;
  });

  var mensagem =
    '🛡️ STATUS DO SISTEMA (ABA: ' + NOME_DA_ABA_ALVO + ')\n\n' +
    '📊 INFORMAÇÕES\n' +
    '• Linhas totais: ' + linhasAtuais + ' / ' + LIMITE_MAXIMO_LINHAS + '\n' +
    '• Linhas com fórmulas: ' + linhasComFormulas + '\n' +
    '• Trigger onChange: ' + (triggerOnChange ? '✓ Ativo' : '✗ Inativo') + '\n' +
    '• Monitoramento: ' + (triggerMonitoramento ? '✓ Ativo' : '✗ Inativo') + '\n';

  ui.alert('📊 Status do Sistema', mensagem, ui.ButtonSet.OK);
}

function resetTotalDoSistema() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    '🔨 RESET TOTAL (ABA: ' + NOME_DA_ABA_ALVO + ')',
    'Deseja resetar o sistema e focar apenas na aba configurada?',
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) return;

  try {
    var aba = pegarAbaAlvo();
    if (!aba) {
      ui.alert('Erro', 'Aba "' + NOME_DA_ABA_ALVO + '" não existe.', ui.ButtonSet.OK);
      return;
    }

    var linhasAntes = aba.getLastRow();

    verificarEAplicarLimite();
    garantirFormulasCompletas();
    verificarIntegridadeFormulas();

    var linhasDepois = aba.getLastRow();

    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function (t) {
      if (t.getHandlerFunction() === 'onChangeHandler' ||
        t.getHandlerFunction() === 'monitoramentoContinuo') {
        ScriptApp.deleteTrigger(t);
      }
    });

    ScriptApp.newTrigger('onChangeHandler')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onChange()
      .create();

    criarTriggerDeMonitoramento();

    PropertiesService.getScriptProperties().setProperty(PROPRIEDADE_LINHAS, LIMITE_MAXIMO_LINHAS.toString());

    ui.alert('✅ Reset Total', 'Sistema resetado e configurado para a aba: ' + NOME_DA_ABA_ALVO, ui.ButtonSet.OK);
  } catch (erro) {
    ui.alert('❌ Erro', 'Erro durante reset: ' + erro.toString(), ui.ButtonSet.OK);
  }
}


/*
 * Bloqueio rápido + filtro / view
 * ============================================================ */
function onEdit(e) { // wrapper simples/instalável
  try { onEditHandler(e); } catch (err) { console.error('[onEdit wrapper] ', err); }
}


/**
 * Bloqueio ao pular linha, permitindo somente editar na liha que começou.
 * se não ele bloqueia a ação.
 * ===================================================================== */
function onEditHandler(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (!sheet) return;

    // Restringe à aba configurada
    if (String(sheet.getName()).toLowerCase() !== String(CONFIG.BLOCK.SHEET).toLowerCase()) return;

    var r = e.range;

    // Colunas obrigatórias NA ORDEM definida
    var reqColsIdxOrdered = toIndexesKeepOrder_(CONFIG.BLOCK.REQUIRED_COLS);

    // Primeira linha VISÍVEL com pendências (respeita Filtro e Visualização de filtro)
    var firstInc = findFirstIncompleteRowFast_(sheet, CONFIG.DATA_START_ROW, reqColsIdxOrdered);

    // Sem pendências visíveis → libera qualquer edição
    if (!firstInc) return;

    // 1) Dentro da própria linha pendente → permitir livremente
    if (r.getRow() === firstInc.rowNum) return;

    // 2) Linhas ACIMA da pendente → permitir (edições retroativas)
    if (r.getRow() < firstInc.rowNum) return;

    // 3) Tentou começar em linha ABAIXO da pendente → bloquear, limpar e redirecionar
    r.clearContent(); // cobre célula única ou blocos colados

    SpreadsheetApp.getActive().toast(
      'Complete primeiro a linha ' + firstInc.rowNum + '. Faltando: ' + firstInc.missing.join(', '),
      'Preenchimento obrigatório',
      6
    );

    var targetColIdx = colIndex_(firstInc.missing[0]); // 1-based
    safeActivate_(sheet, firstInc.rowNum, targetColIdx);
  } catch (err) {
    console.error('[onEditHandler v2.1] ', (err && err.stack) || err);
  }
}


/* Proteção das colunas
 * ============================================================ */
var NOME_PROPRIEDADE_PROTECAO = 'protecaoAtiva';
function protegerColunasComAviso() {
  var ui = SpreadsheetApp.getUi();
  var resposta = ui.alert(
    '⚠️ AVISO DE PROTEÇÃO DE COLUNAS',
    'Este script protegerá as seguintes colunas contra edição:\n\n' +
    '🔒 Colunas: ' + CONFIG.PROTECTION.COLS.join(', ') + '\n\n' +
    'IMPORTANTE:\n' +
    '• Não editáveis por outros usuários\n' +
    '• O proprietário permanece autorizado\n' +
    '• Atualiza dinamicamente ao mudar o tamanho da planilha\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resposta === ui.Button.NO) { ui.alert('Operação cancelada.'); return; }

  try {
    var resultado = aplicarProtecaoDinamica();
    if (resultado.sucesso) {
      ui.alert(
        '✅ PROTEÇÃO APLICADA',
        '📋 Planilha: ' + resultado.nomePlanilha + '\n' +
        '🔒 Colunas: ' + CONFIG.PROTECTION.COLS.join(', ') + '\n' +
        '📏 Linhas protegidas: ' + CONFIG.PROTECTION.START_ROW + ' até ' + resultado.ultimaLinha + '\n' +
        '👤 Editor autorizado: ' + (resultado.emailProprietario || 'Indisponível'),
        ui.ButtonSet.OK
      );
      PropertiesService.getScriptProperties().setProperty(NOME_PROPRIEDADE_PROTECAO, 'true');
    } else {
      throw new Error(resultado.erro);
    }
  } catch (erro) {
    ui.alert('❌ Erro ao Aplicar Proteção', String(erro), ui.ButtonSet.OK);
  }
}

function aplicarProtecaoDinamica() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getActiveSheet();
    var nomePlanilha = aba.getName();
    var emailProprietario = (Session.getActiveUser() && Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail && Session.getEffectiveUser().getEmail()) ||
      null;

    removerProtecoesAnteriores(aba);
    var ultimaLinha = aba.getMaxRows();

    CONFIG.PROTECTION.COLS.forEach(function (coluna) {
      protegerColunaCompleta(aba, coluna, ultimaLinha, emailProprietario);
    });

    configurarMonitoramentoAutomaticoProtecao();
    return { sucesso: true, nomePlanilha: nomePlanilha, ultimaLinha: ultimaLinha, emailProprietario: emailProprietario };
  } catch (erro) {
    console.error('Erro ao aplicar proteção: ' + erro.toString());
    return { sucesso: false, erro: erro.toString() };
  } finally {
    lock.releaseLock();
  }
}

function protegerColunaCompleta(aba, coluna, ultimaLinha, emailProprietario) {
  try {
    var intervalo = aba.getRange(coluna + CONFIG.PROTECTION.START_ROW + ':' + coluna + ultimaLinha);
    var protecao = intervalo.protect().setDescription('Coluna ' + coluna + ' – Proteção Automática');
    var editores = protecao.getEditors();
    if (editores && editores.length) protecao.removeEditors(editores);
    if (emailProprietario) protecao.addEditor(emailProprietario);
    protecao.setWarningOnly(false);
    console.log('Coluna ' + coluna + ' protegida.');
  } catch (e) {
    console.error('Falha ao proteger coluna ' + coluna + ': ' + e);
  }
}

function removerProtecoesAnteriores(aba) {
  try {
    var protecoes = aba.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    protecoes.forEach(function (protecao) {
      try { protecao.remove(); } catch (e) { /* ignora se não for dono */ }
    });
  } catch (e) {
    console.error('Erro ao remover proteções: ' + e);
  }
}

function atualizarProtecoes() {
  var ativa = PropertiesService.getScriptProperties().getProperty(NOME_PROPRIEDADE_PROTECAO);
  if (ativa !== 'true') return;
  try {
    var aba = SpreadsheetApp.getActiveSheet();
    var ultimaLinha = aba.getMaxRows();
    var emailProprietario = (Session.getActiveUser() && Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail && Session.getEffectiveUser().getEmail()) ||
      null;
    console.log('Atualizando proteções p/ ' + ultimaLinha + ' linhas');
    removerProtecoesAnteriores(aba);
    CONFIG.PROTECTION.COLS.forEach(function (c) { protegerColunaCompleta(aba, c, ultimaLinha, emailProprietario); });
    console.log('Proteções atualizadas.');
  } catch (erro) {
    console.error('Erro ao atualizar proteções: ' + erro.toString());
  }
}

function configurarMonitoramentoAutomaticoProtecao() {
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === 'onChangeProtecao') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('onChangeProtecao').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
    console.log('Monitoramento de proteção configurado.');
  } catch (erro) { console.error('Erro ao configurar monitoramento da proteção: ' + erro); }
}

function onChangeProtecao(e) {
  try { atualizarProtecoes(); } catch (err) { console.error('onChangeProtecao error: ', err); }
}


/* Menus e seus gatilhos
 * ============================================================ */
function ensureTrigger_(handler, kind) {
  var exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === handler && (kind ? t.getEventType && t.getEventType() === kind : true);
  });
  return exists;
}

function instalarTriggersBasicos_() {
  // onEdit (instalável, chama onEdit)
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onEdit')) {
    ScriptApp.newTrigger('onEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  }
  // onChangeHandler
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onChangeHandler')) {
    ScriptApp.newTrigger('onChangeHandler').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
  }
  // monitoramento contínuo
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'monitoramentoContinuo')) {
    ScriptApp.newTrigger('monitoramentoContinuo').timeBased().everyMinutes(CONFIG.TRIGGERS.MONITOR_EVERY_MINUTES).create();
  }
  // proteção onChange
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onChangeProtecao')) {
    ScriptApp.newTrigger('onChangeProtecao').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
  }
}

function onOpen() {
  try {
    // Menu único
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('🛡️ Sistema')
      .addSubMenu(
        ui.createMenu('▶️ Vencidos')
          .addItem('Analisar status', 'mostrarAnaliseStatus')
          .addItem('Análise detalhada', 'mostrarAnaliseOrdenada')
          .addItem('Gerar relatório de análise', 'gerarRelatorioAnalise')
          .addItem('Filtrar e Exportar', 'filtrarEExportarAnalise')
          .addItem('Mover produtos vencidos', 'moverProdutosVencidos')
          .addItem('Duplicidades [A → Z]', 'encontrarDuplicidades')
      )
      .addSubMenu(
        ui.createMenu('📐 - 1000 Linhas > Somente Vitor')
          .addItem('Configurar triggers automáticos', 'configurarTriggerAutomatico')
          .addItem('Forçar 1k de linhas com as Fórmulas', 'forcarMilLinhasComFormulas')
          .addItem('Verificar status', 'verificarStatus')
          .addItem('Reparar fórmulas quebradas', 'repararFormulasQuebradas')
          .addItem('Reset total do sistema', 'resetTotalDoSistema')
          .addItem('Configurar linha modelo...', 'configurarLinhaModelo')
          .addItem('Teste de segurança completo', 'testeDeSegurancaCompleto')
      )
      .addSubMenu(
        ui.createMenu('🔒 - Proteção > Somente Vitor')
          .addItem('Aplicar proteção de colunas', 'protegerColunasComAviso')
          .addItem('Atualizar proteções', 'atualizarProtecoes')
      )
      .addToUi();

    // Gatilhos base sem duplicar
    instalarTriggersBasicos_();
  } catch (e) {
    console.error('Erro no onOpen:', e);
  }
}

function atualizarStatusNaPlanilha(linha, novoStatus) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SRC_SHEET);
    if (!sheet) throw new Error('Aba de origem não encontrada.');

    const statusColIdx = colLetterToIndex(CONFIG.STATUS_COL_LETTER) + 1;
    sheet.getRange(linha, statusColIdx).setValue(novoStatus);

    return true;
  } catch (err) {
    throw new Error('Falha ao atualizar: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}