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
  const cleaned = String(rawValue || '')
    .replace(/\*/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
  const match = cleaned.match(/^([A-Z]+)\s+(\d+)\.(\d+)$/);
  if (!match) {
    throw new Error(`Localização inválida: "${rawValue}"`);
  }
  return {
    raw: cleaned,
    longarina: match[1],
    altura: Number(match[2]),
    posicao: Number(match[3])
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
  rows.forEach((row) => {
    const location = parseLocation(row.localizacao);
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

  return [...locationMap.values()].sort(compareLocation);
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
  const headers = lines[0].split(',').map((h) => h.trim());
  return lines.slice(1).map((line) => {
    const cols = line.split(',');
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = cols[i] ?? '';
    });
    return obj;
  });
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
    if (!window.XLSX) {
      throw new Error('Importação XLSX indisponível sem biblioteca XLSX local.');
    }
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const firstSheet = wb.SheetNames[0];
    const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[firstSheet], { defval: '' });
    return mapRows(rawRows);
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
