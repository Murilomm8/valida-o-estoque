const STORAGE_KEY = 'valida-estoque-state-v2';
const HISTORY_KEY = 'valida-estoque-history-v1';

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
  newSession: document.getElementById('new-session'),
  historySection: document.getElementById('history-section'),
  historyBody: document.getElementById('history-body')
};

function normalizeHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseLocation(rawValue) {
  const normalized = String(rawValue || '')
    .replace(/\*/g, '')
    .replace(/\\/g, '/')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();

  const match = normalized.match(/^\/?\s*([A-Z]+)\s*(\d+)\.(\d+)$/);
  if (!match) return null;

  const longarina = match[1];
  const altura = Number(match[2]);
  const posicao = Number(match[3]);
  return { raw: `${longarina} ${altura}.${posicao}`, longarina, altura, posicao };
}

function compareLocation(a, b) {
  const l = a.longarina.localeCompare(b.longarina);
  if (l !== 0) return l;
  if (a.altura !== b.altura) return a.altura - b.altura;
  return a.posicao - b.posicao;
}

function loadState() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    Object.assign(state, JSON.parse(saved));
  }
  if (state.locations.length > 0 && state.currentIndex < state.locations.length) {
    openConference();
    renderCurrentLocation();
  }
  renderHistory();
}

function persistState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function getHistory() {
  return JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]');
}

function saveHistory(history) {
  localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
}

function inferColumns(row) {
  const entries = Object.entries(row);
  const find = (names) => entries.find(([k]) => names.includes(normalizeHeader(k)))?.[0];

  const colLocalizacao = find(['localizacao', 'localização']);
  const colProduto = find(['produto', 'descricao', 'descrição']);
  const colUnidadeCx = find([
    'unidade por caixa',
    'unidade p caixa',
    'unidade caixa',
    'unidade f',
    'und cx',
    'un cx',
    'unidade cx',
    'unidade',
    'un'
  ]);
  const colVolume = find(['volume', 'vol']);
  const colValidade = find(['validade', 'vencimento', 'dt validade']);

  if (!colLocalizacao || !colProduto) {
    throw new Error('Não foi possível identificar as colunas obrigatórias: Localização e Produto.');
  }

  return { colLocalizacao, colProduto, colUnidadeCx, colVolume, colValidade };
}

function mapRows(rawRows) {
  if (!rawRows.length) throw new Error('Arquivo sem dados.');
  const cols = inferColumns(rawRows[0]);

  return rawRows
    .filter((r) => String(r[cols.colLocalizacao] || '').trim())
    .map((r) => ({
      localizacao: r[cols.colLocalizacao],
      produto: r[cols.colProduto],
      unidadeCaixa: cols.colUnidadeCx ? r[cols.colUnidadeCx] : '',
      volume: cols.colVolume ? r[cols.colVolume] : '',
      validade: cols.colValidade ? r[cols.colValidade] : ''
    }));
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
      produto: String(row.produto || '').trim(),
      unidadeCaixa: String(row.unidadeCaixa || '').trim(),
      volume: String(row.volume || '').trim(),
      validade: String(row.validade || '').trim()
    });
  });

  const grouped = [...locationMap.values()].sort(compareLocation);
  if (!grouped.length) throw new Error('Nenhuma localização válida encontrada na planilha.');
  if (ignoredRows > 0) console.warn(`Linhas ignoradas sem localização válida: ${ignoredRows}`);
  return grouped;
}

function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter(Boolean);

  if (lines.length < 2) throw new Error('CSV sem linhas suficientes.');
  const comma = (lines[0].match(/,/g) || []).length;
  const semi = (lines[0].match(/;/g) || []).length;
  const delimiter = semi > comma ? ';' : ',';
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
  return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
}

async function inflateRaw(data) {
  if (!window.DecompressionStream) throw new Error('Navegador sem suporte local para XLSX.');
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([data]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function unzipEntries(buffer) {
  const bytes = new Uint8Array(buffer);
  const decoder = new TextDecoder('utf-8');
  let eocdOffset = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 66000); i -= 1) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) {
      eocdOffset = i;
      break;
    }
  }
  if (eocdOffset < 0) throw new Error('ZIP inválido');

  const entries = readUInt16LE(bytes, eocdOffset + 10);
  let ptr = readUInt32LE(bytes, eocdOffset + 16);
  const files = new Map();

  for (let i = 0; i < entries; i += 1) {
    const compressionMethod = readUInt16LE(bytes, ptr + 10);
    const compressedSize = readUInt32LE(bytes, ptr + 20);
    const fileNameLength = readUInt16LE(bytes, ptr + 28);
    const extraLength = readUInt16LE(bytes, ptr + 30);
    const commentLength = readUInt16LE(bytes, ptr + 32);
    const localHeaderOffset = readUInt32LE(bytes, ptr + 42);

    const fileName = decoder.decode(bytes.slice(ptr + 46, ptr + 46 + fileNameLength));
    const lhNameLen = readUInt16LE(bytes, localHeaderOffset + 26);
    const lhExtraLen = readUInt16LE(bytes, localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + lhNameLen + lhExtraLen;
    const compressed = bytes.slice(dataStart, dataStart + compressedSize);

    let content;
    if (compressionMethod === 0) content = compressed;
    else if (compressionMethod === 8) content = await inflateRaw(compressed);
    else throw new Error('Compressão não suportada');

    files.set(fileName, content);
    ptr += 46 + fileNameLength + extraLength + commentLength;
  }
  return files;
}

function xmlDoc(bytes) {
  return new DOMParser().parseFromString(new TextDecoder('utf-8').decode(bytes), 'application/xml');
}

function getCellText(cell, sharedStrings) {
  const type = cell.getAttribute('t');
  if (type === 's') {
    const idx = Number(cell.querySelector('v')?.textContent || 0);
    return sharedStrings[idx] ?? '';
  }
  if (type === 'inlineStr') return cell.querySelector('is t')?.textContent ?? '';
  return cell.querySelector('v')?.textContent ?? '';
}

function colIndexFromRef(ref) {
  const letters = (ref.match(/[A-Z]+/i)?.[0] || 'A').toUpperCase();
  let idx = 0;
  for (let i = 0; i < letters.length; i += 1) idx = idx * 26 + (letters.charCodeAt(i) - 64);
  return idx - 1;
}

async function parseXlsxInBrowser(file) {
  const files = await unzipEntries(await file.arrayBuffer());
  const workbook = xmlDoc(files.get('xl/workbook.xml'));
  const rels = xmlDoc(files.get('xl/_rels/workbook.xml.rels'));
  const firstSheet = workbook.querySelector('sheets > sheet');
  if (!firstSheet) throw new Error('Planilha sem aba');

  const relId = firstSheet.getAttribute('r:id');
  const relNode = [...rels.querySelectorAll('Relationship')].find((n) => n.getAttribute('Id') === relId);
  const target = (relNode?.getAttribute('Target') || '').replace(/^\/+/, '');
  const sheetPath = target.startsWith('worksheets/') ? `xl/${target}` : `xl/worksheets/${target.split('/').pop()}`;
  const sheet = xmlDoc(files.get(sheetPath));

  const sharedStrings = [];
  if (files.get('xl/sharedStrings.xml')) {
    xmlDoc(files.get('xl/sharedStrings.xml')).querySelectorAll('si').forEach((si) => {
      sharedStrings.push([...si.querySelectorAll('t')].map((t) => t.textContent || '').join(''));
    });
  }

  const rows = [...sheet.querySelectorAll('sheetData > row')].map((row) => {
    const values = {};
    row.querySelectorAll('c').forEach((cell) => {
      values[colIndexFromRef(cell.getAttribute('r') || 'A1')] = getCellText(cell, sharedStrings);
    });
    const max = Math.max(-1, ...Object.keys(values).map(Number));
    return Array.from({ length: max + 1 }, (_, i) => values[i] ?? '');
  });

  if (rows.length < 2) throw new Error('XLSX sem dados');
  const headers = rows[0].map((h, i) => String(h || '').trim() || `COL_${i + 1}`);
  const jsonRows = rows.slice(1).map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? ''])));
  return mapRows(jsonRows);
}

async function importFile(file) {
  const ext = file.name.toLowerCase().split('.').pop();
  if (ext === 'csv') return mapRows(parseCsv(await file.text()));
  if (ext === 'xlsx') {
    if (window.XLSX?.read) {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
      return mapRows(rawRows);
    }
    return parseXlsxInBrowser(file);
  }
  throw new Error('Formato não suportado. Use CSV ou XLSX.');
}

function openConference() {
  el.importSection.classList.add('hidden');
  el.reportSection.classList.add('hidden');
  el.conferenceSection.classList.remove('hidden');
}

function renderCurrentLocation() {
  const location = state.locations[state.currentIndex];
  if (!location) return finishSession();

  el.currentLocation.textContent = location.raw;
  el.progressText.textContent = `Localização ${state.currentIndex + 1} de ${state.locations.length}`;

  const existing = state.records[location.raw]?.items || [];
  el.itemsBody.innerHTML = '';

  location.items.forEach((item, idx) => {
    const restored = existing[idx] || {
      produtoReal: item.produto,
      unidadeCaixaReal: item.unidadeCaixa,
      volumeReal: item.volume,
      validadeReal: item.validade,
      notFound: false
    };

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${item.produto}</td>
      <td>${item.unidadeCaixa}</td>
      <td>${item.volume}</td>
      <td>${item.validade}</td>
      <td><input type="text" value="${restored.produtoReal}"></td>
      <td><input type="text" value="${restored.unidadeCaixaReal}"></td>
      <td><input type="text" value="${restored.volumeReal}"></td>
      <td><input type="text" value="${restored.validadeReal}"></td>
      <td><input type="checkbox" ${restored.notFound ? 'checked' : ''} class="not-found"></td>
      <td class="status-cell"></td>
    `;
    el.itemsBody.appendChild(tr);
  });

  refreshStatuses();
}

function normalizeCompare(value) {
  return String(value || '').trim().toUpperCase();
}

function refreshStatuses() {
  [...el.itemsBody.querySelectorAll('tr')].forEach((tr) => {
    const expected = {
      produto: tr.children[0].textContent,
      unidadeCaixa: tr.children[1].textContent,
      volume: tr.children[2].textContent,
      validade: tr.children[3].textContent,
    };

    const real = {
      produto: tr.children[4].querySelector('input').value,
      unidadeCaixa: tr.children[5].querySelector('input').value,
      volume: tr.children[6].querySelector('input').value,
      validade: tr.children[7].querySelector('input').value,
      notFound: tr.children[8].querySelector('input').checked
    };

    const ok = !real.notFound
      && normalizeCompare(real.produto) === normalizeCompare(expected.produto)
      && normalizeCompare(real.unidadeCaixa) === normalizeCompare(expected.unidadeCaixa)
      && normalizeCompare(real.volume) === normalizeCompare(expected.volume)
      && normalizeCompare(real.validade) === normalizeCompare(expected.validade)
      ;

    const status = tr.children[9];
    status.textContent = ok ? 'Correto' : real.notFound ? 'Não encontrado' : 'Divergente';
    status.className = `status-cell ${ok ? 'status-ok' : 'status-div'}`;
  });
}

function captureRows(location) {
  return [...el.itemsBody.querySelectorAll('tr')].map((tr, idx) => {
    const item = location.items[idx];
    const produtoReal = tr.children[4].querySelector('input').value;
    const unidadeCaixaReal = tr.children[5].querySelector('input').value;
    const volumeReal = tr.children[6].querySelector('input').value;
    const validadeReal = tr.children[7].querySelector('input').value;
    const notFound = tr.children[8].querySelector('input').checked;

    const status = tr.children[9].textContent;

    return {
      ...item,
      produtoReal,
      unidadeCaixaReal,
      volumeReal,
      validadeReal,
      notFound,
      status
    };
  });
}

function confirmAndNext() {
  const location = state.locations[state.currentIndex];
  const items = captureRows(location);
  state.records[location.raw] = {
    timestamp: new Date().toISOString(),
    status: items.every((i) => i.status === 'Correto') ? 'Correto' : 'Divergente',
    items
  };
  state.currentIndex += 1;
  persistState();
  renderCurrentLocation();
}

function buildReportData() {
  const records = Object.entries(state.records);
  const total = state.locations.length;
  const corretas = records.filter(([, r]) => r.status === 'Correto').length;
  const divergentes = total - corretas;

  const divergences = [];
  records.forEach(([localizacao, rec]) => {
    rec.items.forEach((item) => {
      if (item.status === 'Correto') return;
      divergences.push({
        localizacao,
        esperado_produto: item.produto,
        encontrado_produto: item.produtoReal,
        esperado_unidade_caixa: item.unidadeCaixa,
        encontrado_unidade_caixa: item.unidadeCaixaReal,
        esperado_volume: item.volume,
        encontrado_volume: item.volumeReal,
        esperado_validade: item.validade,
        encontrado_validade: item.validadeReal,
        status: item.status
      });
    });
  });

  return { total, corretas, divergentes, divergences };
}

function renderReport() {
  const data = buildReportData();
  el.reportSummary.innerHTML = `
    <div><strong>Operador:</strong><br>${state.operator}</div>
    <div><strong>Início:</strong><br>${new Date(state.startedAt).toLocaleString('pt-BR')}</div>
    <div><strong>Total localizações:</strong><br>${data.total}</div>
    <div><strong>Corretas:</strong><br>${data.corretas}</div>
    <div><strong>Com divergência:</strong><br>${data.divergentes}</div>
  `;

  el.divergenceBody.innerHTML = '';
  data.divergences.forEach((d) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${d.localizacao}</td><td>${d.esperado_produto}</td><td>${d.encontrado_produto}</td><td>${d.status}</td>`;
    el.divergenceBody.appendChild(tr);
  });

  return data;
}

function makeCsv(rows) {
  if (!rows.length) return 'localizacao,esperado_produto,encontrado_produto,status\n';
  const headers = Object.keys(rows[0]);
  const esc = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
  return [headers.join(','), ...rows.map((r) => headers.map((h) => esc(r[h])).join(','))].join('\n');
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function crc32(bytes) {
  let crc = ~0;
  for (let i = 0; i < bytes.length; i += 1) {
    crc ^= bytes[i];
    for (let j = 0; j < 8; j += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (~crc) >>> 0;
}

function u16(n) { return [n & 255, (n >> 8) & 255]; }
function u32(n) { return [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >> 24) & 255]; }

function zipStore(files) {
  const enc = new TextEncoder();
  let offset = 0;
  const locals = [];
  const centrals = [];

  files.forEach((f) => {
    const name = enc.encode(f.name);
    const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
    const crc = crc32(data);

    const local = new Uint8Array([
      80, 75, 3, 4, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length), 0, 0,
      ...name, ...data
    ]);

    const central = new Uint8Array([
      80, 75, 1, 2, 20, 0, 20, 0, 0, 0, 0, 0, 0, 0,
      ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length), 0, 0, 0, 0, 0, 0, 0, 0,
      ...u32(offset), ...name
    ]);

    locals.push(local);
    centrals.push(central);
    offset += local.length;
  });

  const centralSize = centrals.reduce((s, c) => s + c.length, 0);
  const centralOffset = locals.reduce((s, l) => s + l.length, 0);
  const eocd = new Uint8Array([80, 75, 5, 6, 0, 0, 0, 0, ...u16(files.length), ...u16(files.length), ...u32(centralSize), ...u32(centralOffset), 0, 0]);

  return new Blob([...locals, ...centrals, eocd], { type: 'application/zip' });
}

function rowsToXlsxBlob(rows) {
  const headers = rows.length ? Object.keys(rows[0]) : ['Localização'];
  const esc = (v) => String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const sheetRows = [
    `<row r="1">${headers.map((h, i) => `<c r="${String.fromCharCode(65 + i)}1" t="inlineStr"><is><t>${esc(h)}</t></is></c>`).join('')}</row>`,
    ...rows.map((r, rowIndex) => {
      const rowNum = rowIndex + 2;
      return `<row r="${rowNum}">${headers.map((h, i) => `<c r="${String.fromCharCode(65 + i)}${rowNum}" t="inlineStr"><is><t>${esc(r[h])}</t></is></c>`).join('')}</row>`;
    })
  ].join('');

  const files = [
    {
      name: '[Content_Types].xml',
      data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`
    },
    {
      name: '_rels/.rels',
      data: `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
    },
    {
      name: 'xl/workbook.xml',
      data: `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Divergencias" sheetId="1" r:id="rId1"/></sheets></workbook>`
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`
    },
    {
      name: 'xl/worksheets/sheet1.xml',
      data: `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetRows}</sheetData></worksheet>`
    }
  ];

  return zipStore(files);
}

function exportRowsCsv(rows, filename) {
  downloadBlob(filename, new Blob([makeCsv(rows)], { type: 'text/csv;charset=utf-8;' }));
}

function exportRowsXlsx(rows, filename) {
  downloadBlob(filename, rowsToXlsxBlob(rows));
}

function exportCurrentCsv() {
  const data = buildReportData();
  exportRowsCsv(data.divergences, 'relatorio-divergencias.csv');
}

function exportCurrentXlsx() {
  const data = buildReportData();
  exportRowsXlsx(data.divergences, 'relatorio-divergencias.xlsx');
}

function finishSession() {
  el.conferenceSection.classList.add('hidden');
  el.reportSection.classList.remove('hidden');
  const data = renderReport();

  const history = getHistory();
  history.unshift({
    id: Date.now(),
    operator: state.operator,
    startedAt: state.startedAt,
    finishedAt: new Date().toISOString(),
    totalLocations: data.total,
    correctLocations: data.corretas,
    divergentLocations: data.divergentes,
    divergences: data.divergences
  });
  saveHistory(history.slice(0, 100));
  renderHistory();
  persistState();
}

function renderHistory() {
  const history = getHistory();
  if (!el.historyBody) return;

  el.historyBody.innerHTML = '';
  history.forEach((h) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${new Date(h.finishedAt).toLocaleString('pt-BR')}</td>
      <td>${h.operator}</td>
      <td>${h.totalLocations}</td>
      <td>${h.correctLocations}</td>
      <td>${h.divergentLocations}</td>
      <td>
        <button data-id="${h.id}" data-type="csv">CSV</button>
        <button data-id="${h.id}" data-type="xlsx">XLSX</button>
      </td>
    `;
    el.historyBody.appendChild(tr);
  });

  el.historySection?.classList.toggle('hidden', history.length === 0);
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

el.exportCsv.addEventListener('click', exportCurrentCsv);
el.exportXlsx.addEventListener('click', exportCurrentXlsx);

el.newSession.addEventListener('click', () => {
  localStorage.removeItem(STORAGE_KEY);
  location.reload();
});

el.historyBody?.addEventListener('click', (event) => {
  const target = event.target.closest('button[data-id]');
  if (!target) return;
  const id = Number(target.dataset.id);
  const type = target.dataset.type;
  const entry = getHistory().find((h) => h.id === id);
  if (!entry) return;
  if (type === 'csv') exportRowsCsv(entry.divergences, `historico-${id}.csv`);
  else exportRowsXlsx(entry.divergences, `historico-${id}.xlsx`);
});

el.startSession.addEventListener('click', async () => {
  try {
    const operator = el.operatorName.value.trim();
    const file = el.fileInput.files[0];
    if (!operator) throw new Error('Informe o nome do operador.');
    if (!file) throw new Error('Selecione um arquivo para importar.');

    const grouped = groupRows(await importFile(file));
    state.operator = operator;
    state.startedAt = new Date().toISOString();
    state.locations = grouped;
    state.currentIndex = 0;
    state.records = {};
    persistState();

    openConference();
    renderCurrentLocation();
  } catch (error) {
    alert(error.message || String(error));
    console.error(error);
  }
});

loadState();
