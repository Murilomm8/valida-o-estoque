const STORAGE_KEY = 'valida-estoque-state-v1';

const state = {
  operator: '',
  startedAt: null,
  locations: [],
  currentIndex: 0,
  records: {}
};

const el = {
  operatorName: document.getElementById('operator-name'),
  fileInput: document.getElementById('file-input'),
  startSession: document.getElementById('start-session'),
  importSection: document.getElementById('import-section'),
  conferenceSection: document.getElementById('conference-section'),
  reportSection: document.getElementById('report-section'),
  currentLocation: document.getElementById('current-location'),
  progressText: document.getElementById('progress-text'),
  itemsBody: document.getElementById('items-body'),
  confirmNext: document.getElementById('confirm-next'),
  reportSummary: document.getElementById('report-summary'),
  divergenceBody: document.getElementById('divergence-body'),
  exportCsv: document.getElementById('export-csv'),
  exportXlsx: document.getElementById('export-xlsx'),
  newSession: document.getElementById('new-session')
};

function normalizeHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '');
}

function parseLocation(rawValue) {
  const normalized = String(rawValue || '')
    .replace(/\*/g, '')
    .replace(/\\/g, '/')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();

  const compact = normalized.replace(/\s+/g, '');
  const directMatch = compact.match(/([A-Z]+)(\d+)\.(\d+)/);
  const spacedMatch = normalized.match(/([A-Z]+)\s+(\d+)\.(\d+)/);
  const match = directMatch || spacedMatch;

  if (!match) {
    return null;
  }

  const longarina = match[1];
  const altura = Number(match[2]);
  const posicao = Number(match[3]);

  return {
    raw: `${longarina} ${altura}.${posicao}`,
    longarina,
    altura,
    posicao
  };
}


function compareLocation(a, b) {
  const l = a.longarina.localeCompare(b.longarina);
  if (l !== 0) return l;
  if (a.altura !== b.altura) return a.altura - b.altura;
  return a.posicao - b.posicao;
}

function loadState() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) return;
  Object.assign(state, JSON.parse(saved));
  if (state.locations.length > 0 && state.currentIndex < state.locations.length) {
    openConference();
    renderCurrentLocation();
  }
}

function persistState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function groupRows(rows) {
  const locationMap = new Map();
  let ignoredRows = 0;

  rows.forEach((row) => {
    const location = parseLocation(row.localizacao);
    if (!location) {
      ignoredRows += 1;
      return;
    }

    const key = location.raw;
    if (!locationMap.has(key)) {
      locationMap.set(key, { ...location, items: [] });
    }
    locationMap.get(key).items.push({
      sku: String(row.sku || '').trim(),
      produto: String(row.produto || '').trim(),
      qtdEsperada: Number(row.quantidade || 0)
    });
  });

  const grouped = [...locationMap.values()].sort(compareLocation);
  if (!grouped.length) {
    throw new Error('Nenhuma localização válida encontrada na planilha importada.');
  }
  if (ignoredRows > 0) {
    console.warn(`Linhas ignoradas por localização inválida: ${ignoredRows}`);
  }

  return grouped;
}


function inferColumns(row) {
  const entries = Object.entries(row);
  const find = (names) => entries.find(([k]) => names.includes(normalizeHeader(k)))?.[0];

  const colLocalizacao = find(['localizacao', 'localização']);
  const colSku = find(['sku', 'codigo', 'codigo sku']);
  const colProduto = find(['produto', 'descricao', 'descrição']);
  const colQuantidade = find(['quantidade', 'qtd', 'qtd esperada', 'quantidade esperada']);

  if (!colLocalizacao || !colSku || !colProduto || !colQuantidade) {
    throw new Error('Não foi possível identificar as colunas necessárias na planilha.');
  }
  return { colLocalizacao, colSku, colProduto, colQuantidade };
}

function mapRows(rawRows) {
  if (!rawRows.length) throw new Error('Arquivo sem dados.');
  const cols = inferColumns(rawRows[0]);

  return rawRows
    .filter((r) => String(r[cols.colLocalizacao] || '').trim())
    .map((r) => ({
      localizacao: r[cols.colLocalizacao],
      sku: r[cols.colSku],
      produto: r[cols.colProduto],
      quantidade: r[cols.colQuantidade]
    }));
}
function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) throw new Error('CSV sem linhas suficientes.');

  const firstLine = lines[0];
  const commaCount = (firstLine.match(/,/g) || []).length;
  const semicolonCount = (firstLine.match(/;/g) || []).length;
  const delimiter = semicolonCount > commaCount ? ';' : ',';

  const headers = lines[0].split(delimiter).map((h) => h.trim());
  return lines.slice(1).map((line) => {
    const cols = line.split(delimiter);
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = cols[i] ?? '';
    });
    return obj;
  });
}



function readUInt16LE(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readUInt32LE(bytes, offset) {
  return (
    bytes[offset] |
    (bytes[offset + 1] << 8) |
    (bytes[offset + 2] << 16) |
    (bytes[offset + 3] << 24)
  ) >>> 0;
}

async function inflateRaw(data) {
  if (!window.DecompressionStream) {
    throw new Error('Navegador sem suporte a descompressão local de XLSX.');
  }
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([data]).stream().pipeThrough(ds);
  const out = await new Response(stream).arrayBuffer();
  return new Uint8Array(out);
}

async function unzipEntries(buffer) {
  const bytes = new Uint8Array(buffer);
  const decoder = new TextDecoder('utf-8');

  let eocdOffset = -1;
  const min = Math.max(0, bytes.length - 66000);
  for (let i = bytes.length - 22; i >= min; i -= 1) {
    if (
      bytes[i] === 0x50 &&
      bytes[i + 1] === 0x4b &&
      bytes[i + 2] === 0x05 &&
      bytes[i + 3] === 0x06
    ) {
      eocdOffset = i;
      break;
    }
  }

  if (eocdOffset < 0) throw new Error('ZIP inválido (EOCD não encontrado).');

  const centralDirectoryOffset = readUInt32LE(bytes, eocdOffset + 16);
  const totalEntries = readUInt16LE(bytes, eocdOffset + 10);

  const files = new Map();
  let ptr = centralDirectoryOffset;

  for (let i = 0; i < totalEntries; i += 1) {
    if (
      bytes[ptr] !== 0x50 ||
      bytes[ptr + 1] !== 0x4b ||
      bytes[ptr + 2] !== 0x01 ||
      bytes[ptr + 3] !== 0x02
    ) {
      throw new Error('ZIP inválido (cabeçalho de diretório central).');
    }

    const compressionMethod = readUInt16LE(bytes, ptr + 10);
    const compressedSize = readUInt32LE(bytes, ptr + 20);
    const fileNameLength = readUInt16LE(bytes, ptr + 28);
    const extraLength = readUInt16LE(bytes, ptr + 30);
    const commentLength = readUInt16LE(bytes, ptr + 32);
    const localHeaderOffset = readUInt32LE(bytes, ptr + 42);

    const fileNameStart = ptr + 46;
    const fileNameEnd = fileNameStart + fileNameLength;
    const fileName = decoder.decode(bytes.slice(fileNameStart, fileNameEnd));

    const lh = localHeaderOffset;
    if (
      bytes[lh] !== 0x50 ||
      bytes[lh + 1] !== 0x4b ||
      bytes[lh + 2] !== 0x03 ||
      bytes[lh + 3] !== 0x04
    ) {
      throw new Error('ZIP inválido (local header).');
    }

    const localNameLen = readUInt16LE(bytes, lh + 26);
    const localExtraLen = readUInt16LE(bytes, lh + 28);
    const dataStart = lh + 30 + localNameLen + localExtraLen;
    const compressed = bytes.slice(dataStart, dataStart + compressedSize);

    let content;
    if (compressionMethod === 0) {
      content = compressed;
    } else if (compressionMethod === 8) {
      content = await inflateRaw(compressed);
    } else {
      throw new Error(`Método de compressão ZIP não suportado: ${compressionMethod}`);
    }

    files.set(fileName, content);
    ptr = fileNameEnd + extraLength + commentLength;
  }

  return files;
}

function xmlDocFromBytes(bytes) {
  const text = new TextDecoder('utf-8').decode(bytes);
  return new DOMParser().parseFromString(text, 'application/xml');
}

function getCellText(cell, sharedStrings) {
  const type = cell.getAttribute('t');
  if (type === 's') {
    const v = cell.querySelector('v');
    if (!v) return '';
    const idx = Number(v.textContent || 0);
    return sharedStrings[idx] ?? '';
  }
  if (type === 'inlineStr') {
    return cell.querySelector('is t')?.textContent ?? '';
  }
  return cell.querySelector('v')?.textContent ?? '';
}

function colIndexFromRef(ref) {
  const col = (ref.match(/[A-Z]+/i)?.[0] || 'A').toUpperCase();
  let idx = 0;
  for (let i = 0; i < col.length; i += 1) {
    idx = idx * 26 + (col.charCodeAt(i) - 64);
  }
  return Math.max(0, idx - 1);
}

function parseSheetRows(sheetDoc, sharedStrings) {
  const rowNodes = [...sheetDoc.querySelectorAll('sheetData > row')];
  return rowNodes.map((row) => {
    const values = {};
    [...row.querySelectorAll('c')].forEach((cell) => {
      const ref = cell.getAttribute('r') || 'A1';
      values[colIndexFromRef(ref)] = getCellText(cell, sharedStrings);
    });
    const maxIndex = Math.max(-1, ...Object.keys(values).map(Number));
    return Array.from({ length: maxIndex + 1 }, (_, i) => values[i] ?? '');
  });
}

async function parseXlsxInBrowser(file) {
  const buffer = await file.arrayBuffer();
  const files = await unzipEntries(buffer);

  const workbookBytes = files.get('xl/workbook.xml');
  const relsBytes = files.get('xl/_rels/workbook.xml.rels');
  if (!workbookBytes || !relsBytes) {
    throw new Error('Estrutura XLSX inválida.');
  }

  const sharedBytes = files.get('xl/sharedStrings.xml');
  const sharedStrings = [];
  if (sharedBytes) {
    const sharedDoc = xmlDocFromBytes(sharedBytes);
    sharedDoc.querySelectorAll('si').forEach((si) => {
      const text = [...si.querySelectorAll('t')].map((t) => t.textContent || '').join('');
      sharedStrings.push(text);
    });
  }

  const workbookDoc = xmlDocFromBytes(workbookBytes);
  const firstSheet = workbookDoc.querySelector('sheets > sheet');
  if (!firstSheet) throw new Error('XLSX sem abas.');

  const relId = firstSheet.getAttribute('r:id');
  if (!relId) throw new Error('XLSX sem relacionamento da primeira aba.');

  const relsDoc = xmlDocFromBytes(relsBytes);
  const relNode = [...relsDoc.querySelectorAll('Relationship')].find((r) => r.getAttribute('Id') === relId);
  if (!relNode) throw new Error('Relacionamento da aba não encontrado.');

  const target = relNode.getAttribute('Target') || '';
  const normalizedTarget = target.replace(/^\/+/, '');
  const sheetPath = normalizedTarget.startsWith('worksheets/') ? `xl/${normalizedTarget}` : `xl/worksheets/${normalizedTarget.split('/').pop()}`;
  const sheetBytes = files.get(sheetPath);
  if (!sheetBytes) throw new Error('Arquivo da primeira aba não encontrado.');

  const sheetDoc = xmlDocFromBytes(sheetBytes);
  const rows = parseSheetRows(sheetDoc, sharedStrings).filter((r) => r.length > 0);
  if (rows.length < 2) throw new Error('Planilha XLSX sem dados suficientes.');

  const headers = rows[0].map((h, i) => String(h || '').trim() || `COL_${i + 1}`);
  const jsonRows = rows.slice(1).map((row) => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i] ?? '';
    });
    return obj;
  });

  return mapRows(jsonRows);
}

function openConference() {
  el.importSection.classList.add('hidden');
  el.reportSection.classList.add('hidden');
  el.conferenceSection.classList.remove('hidden');
}

function renderCurrentLocation() {
  const location = state.locations[state.currentIndex];
  if (!location) {
    finishSession();
    return;
  }

  el.currentLocation.textContent = location.raw;
  el.progressText.textContent = `Localização ${state.currentIndex + 1} de ${state.locations.length}`;
  const existing = state.records[location.raw]?.items || [];

  el.itemsBody.innerHTML = '';
  location.items.forEach((item, idx) => {
    const restored = existing[idx] || { qtdReal: item.qtdEsperada, notFound: false };
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${item.sku}</td>
      <td>${item.produto}</td>
      <td>${item.qtdEsperada}</td>
      <td><input type="number" min="0" step="1" value="${restored.qtdReal}" data-idx="${idx}" class="qtd-real"></td>
      <td><input type="checkbox" ${restored.notFound ? 'checked' : ''} data-idx="${idx}" class="not-found"></td>
      <td class="status-cell"></td>
    `;
    el.itemsBody.appendChild(tr);
  });
  refreshStatuses();
}

function refreshStatuses() {
  [...el.itemsBody.querySelectorAll('tr')].forEach((tr) => {
    const expected = Number(tr.children[2].textContent);
    const qtdInput = tr.querySelector('.qtd-real');
    const notFound = tr.querySelector('.not-found').checked;
    const real = notFound ? 0 : Number(qtdInput.value || 0);
    const statusCell = tr.querySelector('.status-cell');

    if (!notFound && real === expected) {
      statusCell.textContent = 'Correto';
      statusCell.className = 'status-cell status-ok';
    } else {
      statusCell.textContent = notFound ? 'Não encontrado' : 'Divergente';
      statusCell.className = 'status-cell status-div';
    }
  });
}

function confirmAndNext() {
  const location = state.locations[state.currentIndex];
  const now = new Date().toISOString();
  const rowRecords = [...el.itemsBody.querySelectorAll('tr')].map((tr, idx) => {
    const expected = Number(tr.children[2].textContent);
    const realInput = tr.querySelector('.qtd-real');
    const notFound = tr.querySelector('.not-found').checked;
    const real = notFound ? 0 : Number(realInput.value || 0);
    return {
      ...location.items[idx],
      qtdReal: real,
      notFound,
      difference: real - expected,
      status: !notFound && real === expected ? 'Correto' : notFound ? 'Não encontrado' : 'Divergente'
    };
  });

  const allOk = rowRecords.every((r) => r.status === 'Correto');
  state.records[location.raw] = {
    timestamp: now,
    status: allOk ? 'Correto' : 'Divergente',
    items: rowRecords
  };
  state.currentIndex += 1;
  persistState();
  renderCurrentLocation();
}

function buildReport() {
  const locationRecords = Object.entries(state.records);
  const total = state.locations.length;
  const correct = locationRecords.filter(([, rec]) => rec.status === 'Correto').length;
  const diverging = total - correct;

  el.reportSummary.innerHTML = `
    <div><strong>Operador:</strong><br>${state.operator}</div>
    <div><strong>Início:</strong><br>${new Date(state.startedAt).toLocaleString('pt-BR')}</div>
    <div><strong>Total localizações:</strong><br>${total}</div>
    <div><strong>Corretas:</strong><br>${correct}</div>
    <div><strong>Com divergência:</strong><br>${diverging}</div>
  `;

  const divergences = [];
  locationRecords.forEach(([loc, rec]) => {
    rec.items.forEach((item) => {
      if (item.status !== 'Correto') {
        divergences.push({
          localizacao: loc,
          sku: item.sku,
          esperado: item.qtdEsperada,
          encontrado: item.qtdReal,
          diferenca: item.difference
        });
      }
    });
  });

  el.divergenceBody.innerHTML = '';
  divergences.forEach((d) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${d.localizacao}</td><td>${d.sku}</td><td>${d.esperado}</td><td>${d.encontrado}</td><td>${d.diferenca}</td>`;
    el.divergenceBody.appendChild(tr);
  });

  return divergences;
}

function finishSession() {
  el.conferenceSection.classList.add('hidden');
  el.reportSection.classList.remove('hidden');
  buildReport();
  persistState();
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function exportCsvReport() {
  const divergences = buildReport();
  const headers = ['Localização', 'SKU', 'Esperado', 'Encontrado', 'Diferença'];
  const rows = divergences.map((d) => [d.localizacao, d.sku, d.esperado, d.encontrado, d.diferenca]);
  const csv = [headers, ...rows].map((r) => r.join(',')).join('\n');
  downloadBlob('relatorio-divergencias.csv', new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
}

function exportXlsxReport() {
  if (!window.XLSX) {
    alert('Biblioteca XLSX não disponível no modo offline atual. Use exportação CSV.');
    return;
  }
  const divergences = buildReport();
  const ws = XLSX.utils.json_to_sheet(divergences);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Divergencias');
  XLSX.writeFile(wb, 'relatorio-divergencias.xlsx');
}

async function importFile(file) {
  const ext = file.name.toLowerCase().split('.').pop();
  if (ext === 'csv') {
    const text = await file.text();
    return mapRows(parseCsv(text));
  }
  if (ext === 'xlsx') {
    if (window.XLSX) {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array' });
      const firstSheet = wb.SheetNames[0];
      const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[firstSheet], { defval: '' });
      return mapRows(rawRows);
    }

    try {
      return await parseXlsxInBrowser(file);
    } catch (xlsxError) {
      throw new Error(`Falha ao importar XLSX neste navegador: ${xlsxError.message}`);
    }
  }
  throw new Error('Formato não suportado. Use CSV ou XLSX.');
}

el.itemsBody.addEventListener('input', refreshStatuses);
el.itemsBody.addEventListener('change', refreshStatuses);
el.confirmNext.addEventListener('click', confirmAndNext);
document.addEventListener('keydown', (event) => {
  if (event.key === 'Enter' && !el.conferenceSection.classList.contains('hidden')) {
    event.preventDefault();
    confirmAndNext();
  }
});
el.exportCsv.addEventListener('click', exportCsvReport);
el.exportXlsx.addEventListener('click', exportXlsxReport);

el.newSession.addEventListener('click', () => {
  localStorage.removeItem(STORAGE_KEY);
  location.reload();
});

el.startSession.addEventListener('click', async () => {
  try {
    const operator = el.operatorName.value.trim();
    const file = el.fileInput.files[0];
    if (!operator) throw new Error('Informe o nome do operador.');
    if (!file) throw new Error('Selecione um arquivo para importar.');

    const rows = await importFile(file);
    const grouped = groupRows(rows);

    state.operator = operator;
    state.startedAt = new Date().toISOString();
    state.locations = grouped;
    state.currentIndex = 0;
    state.records = {};
    persistState();

    openConference();
    renderCurrentLocation();
  } catch (error) {
    alert(error.message);
    console.error(error);
  }
});

loadState();
