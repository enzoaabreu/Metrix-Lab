/* ========================================================================
   Painel de Consumo de Materiais — lógica 100% client-side
   ======================================================================== */

const state = {
  rows: [],        // dados normalizados: {produto, setor, data(Date), qte, unidade, valor, total}
  filtered: [],
  charts: {}
};

const COLORS = {
  teal: '#2FD4B8',
  amber: '#F0AE4E',
  coral: '#F0685E',
  violet: '#8B8CF0',
  dim: '#8A95A6',
  grid: 'rgba(255,255,255,0.06)'
};

/* ---------------------------------------------------------------------
   1. Dicionário de reconhecimento de colunas (aceita variações comuns)
   --------------------------------------------------------------------- */
const COLUMN_MAP = {
  produto: ['material', 'produto', 'item', 'descricao', 'descrição'],
  codigo: ['codigo', 'código', 'cod', 'cod.', 'sku'],
  setor: ['destino', 'setor', 'departamento', 'centro de custo', 'centro_custo'],
  origem: ['origem'],
  data: ['data', 'd. lancamento', 'd. lançamento', 'data movimento', 'data lancamento'],
  qte: ['qte', 'quantidade', 'qtd', 'qtde'],
  unidade: ['unidade', 'un', 'un.', 'unid'],
  valor: ['valor'],
  total: ['total', 'valor total'],
  responsavel: ['responsavel', 'responsável'],
};

function normalizeHeader(h) {
  return h.toString().trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // remove acentos
}

function detectColumns(headers) {
  const normalized = headers.map(normalizeHeader);
  const found = {};
  for (const [field, aliases] of Object.entries(COLUMN_MAP)) {
    const normAliases = aliases.map(normalizeHeader);
    const idx = normalized.findIndex(h => normAliases.includes(h));
    if (idx !== -1) found[field] = headers[idx];
  }
  return found;
}

/* ---------------------------------------------------------------------
   2. Parsing de números e datas em formato brasileiro
   --------------------------------------------------------------------- */
function parseBRNumber(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  let s = String(v).trim().replace(/"/g, '');
  if (s === '') return null;
  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function parseBRDate(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') {
    // serial de data do Excel
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  const s = String(v).trim().replace(/"/g, '');
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    let [, dd, mm, yyyy] = m;
    if (yyyy.length === 2) yyyy = '20' + yyyy;
    return new Date(+yyyy, +mm - 1, +dd);
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    const [, yyyy, mm, dd] = m;
    return new Date(+yyyy, +mm - 1, +dd);
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

/* ---------------------------------------------------------------------
   3. Parser de CSV robusto (delimita ; ou , , respeita aspas,
      e tolera linhas com campos extras/faltantes em relação ao header)
   --------------------------------------------------------------------- */
function detectDelimiter(headerLine) {
  const semi = (headerLine.match(/;/g) || []).length;
  const comma = (headerLine.match(/,/g) || []).length;
  return semi >= comma ? ';' : ',';
}

function parseCSVLine(line, delim) {
  const fields = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"') {
        if (line[i + 1] === '"') { cur += '"'; i++; }
        else { inQuotes = false; }
      } else cur += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === delim) { fields.push(cur); cur = ''; }
      else cur += c;
    }
  }
  fields.push(cur);
  return fields.map(f => f.trim());
}

function decodeBuffer(buffer) {
  // tenta UTF-8; se detectar caracteres de substituição, usa windows-1252 (compatível com Latin-1)
  let text = new TextDecoder('utf-8', { fatal: false }).decode(buffer);
  if (text.includes('\uFFFD')) {
    text = new TextDecoder('windows-1252').decode(buffer);
  }
  return text;
}

function parseCSV(buffer) {
  const text = decodeBuffer(buffer);
  const lines = text.split(/\r\n|\n|\r/).filter(l => l.trim() !== '');
  if (lines.length < 2) throw new Error('O arquivo não contém linhas de dados suficientes.');
  const delim = detectDelimiter(lines[0]);
  const headers = parseCSVLine(lines[0], delim);
  const records = [];
  for (let i = 1; i < lines.length; i++) {
    let fields = parseCSVLine(lines[i], delim);
    // tolera campos extras (delimitador sobrando no fim) ou faltantes
    if (fields.length > headers.length) fields = fields.slice(0, headers.length);
    while (fields.length < headers.length) fields.push('');
    const rec = {};
    headers.forEach((h, idx) => { rec[h] = fields[idx]; });
    records.push(rec);
  }
  return { headers, records };
}

function parseExcel(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });
  const headers = rows.length ? Object.keys(rows[0]) : [];
  return { headers, records: rows };
}

/* ---------------------------------------------------------------------
   4. Normalização para o modelo interno
   --------------------------------------------------------------------- */
function normalizeRecords(headers, records) {
  const cols = detectColumns(headers);
  const missing = [];
  if (!cols.produto) missing.push('Produto');
  if (!cols.qte) missing.push('Quantidade');
  if (!cols.data) missing.push('Data');
  if (missing.length) {
    throw new Error(`Não foi possível identificar as colunas: ${missing.join(', ')}. Verifique se a planilha contém essas informações.`);
  }

  const rows = [];
  for (const rec of records) {
    const produto = (rec[cols.produto] ?? '').toString().trim();
    const qte = parseBRNumber(rec[cols.qte]);
    const data = parseBRDate(rec[cols.data]);
    if (!produto || qte === null || !data) continue;
    rows.push({
      produto,
      codigo: cols.codigo ? String(rec[cols.codigo] ?? '').trim() : '',
      setor: cols.setor ? String(rec[cols.setor] ?? '').trim() || 'Não informado' : 'Não informado',
      origem: cols.origem ? String(rec[cols.origem] ?? '').trim() : '',
      data,
      qte,
      unidade: cols.unidade ? String(rec[cols.unidade] ?? '').trim() : '',
      valor: cols.valor ? parseBRNumber(rec[cols.valor]) : null,
      total: cols.total ? parseBRNumber(rec[cols.total]) : null,
      responsavel: cols.responsavel ? String(rec[cols.responsavel] ?? '').trim() : ''
    });
  }
  if (!rows.length) {
    throw new Error('Nenhuma linha válida foi encontrada após a leitura. Verifique o conteúdo da planilha.');
  }
  return rows;
}

/* ---------------------------------------------------------------------
   5. Upload handler
   --------------------------------------------------------------------- */
const fileInput = document.getElementById('fileInput');
const fileStatus = document.getElementById('fileStatus');
const emptyState = document.getElementById('emptyState');
const errorState = document.getElementById('errorState');
const errorMessage = document.getElementById('errorMessage');
const dashboard = document.getElementById('dashboard');

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  fileStatus.textContent = 'Processando…';
  try {
    const buffer = await file.arrayBuffer();
    let parsed;
    if (/\.csv$/i.test(file.name)) {
      parsed = parseCSV(buffer);
    } else {
      parsed = parseExcel(buffer);
    }
    const rows = normalizeRecords(parsed.headers, parsed.records);
    state.rows = rows;
    fileStatus.textContent = `${file.name} · ${rows.length} linhas`;
    errorState.style.display = 'none';
    emptyState.style.display = 'none';
    dashboard.style.display = 'block';
    populateFilters(rows);
    applyFilters();
  } catch (err) {
    console.error(err);
    fileStatus.textContent = '';
    dashboard.style.display = 'none';
    emptyState.style.display = 'none';
    errorState.style.display = 'block';
    errorMessage.textContent = err.message || 'Erro desconhecido ao ler o arquivo.';
  }
});

/* ---------------------------------------------------------------------
   6. Filtros
   --------------------------------------------------------------------- */
const elProduto = document.getElementById('filterProduto');
const elSetor = document.getElementById('filterSetor');
const elDataIni = document.getElementById('filterDataIni');
const elDataFim = document.getElementById('filterDataFim');
const elBusca = document.getElementById('filterBusca');

function populateFilters(rows) {
  const produtos = [...new Set(rows.map(r => r.produto))].sort();
  const setores = [...new Set(rows.map(r => r.setor))].sort();
  elProduto.innerHTML = '<option value="">Todos</option>' + produtos.map(p => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join('');
  elSetor.innerHTML = '<option value="">Todos</option>' + setores.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
}

[elProduto, elSetor, elDataIni, elDataFim, elBusca].forEach(el => {
  el.addEventListener('input', applyFilters);
});
document.getElementById('clearFilters').addEventListener('click', () => {
  elProduto.value = ''; elSetor.value = ''; elDataIni.value = ''; elDataFim.value = ''; elBusca.value = '';
  applyFilters();
});

function applyFilters() {
  const produto = elProduto.value;
  const setor = elSetor.value;
  const dIni = elDataIni.value ? new Date(elDataIni.value) : null;
  const dFim = elDataFim.value ? new Date(elDataFim.value) : null;
  const busca = elBusca.value.trim().toLowerCase();

  state.filtered = state.rows.filter(r => {
    if (produto && r.produto !== produto) return false;
    if (setor && r.setor !== setor) return false;
    if (dIni && r.data < dIni) return false;
    if (dFim && r.data > dFim) return false;
    if (busca) {
      const hay = `${r.produto} ${r.setor} ${r.codigo}`.toLowerCase();
      if (!hay.includes(busca)) return false;
    }
    return true;
  });

  render();
}

/* ---------------------------------------------------------------------
   7. Agregações
   --------------------------------------------------------------------- */
function monthKey(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; }
function monthLabel(key) {
  const [y, m] = key.split('-');
  const nomes = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  return `${nomes[+m - 1]}/${y.slice(2)}`;
}

function aggregateByMonth(rows) {
  const map = new Map();
  for (const r of rows) {
    const k = monthKey(r.data);
    map.set(k, (map.get(k) || 0) + r.qte);
  }
  return [...map.entries()].sort((a, b) => a[0].localeCompare(b[0]));
}

function aggregateByProduct(rows) {
  const map = new Map();
  for (const r of rows) map.set(r.produto, (map.get(r.produto) || 0) + r.qte);
  return [...map.entries()].sort((a, b) => b[1] - a[1]);
}

function aggregateBySector(rows) {
  const map = new Map();
  for (const r of rows) map.set(r.setor, (map.get(r.setor) || 0) + r.qte);
  return [...map.entries()].sort((a, b) => b[1] - a[1]);
}

function productMonthlyChange(rows) {
  // para cada produto, consumo por mês → identifica maior variação entre meses consecutivos
  const byProduct = new Map();
  for (const r of rows) {
    if (!byProduct.has(r.produto)) byProduct.set(r.produto, new Map());
    const m = byProduct.get(r.produto);
    const k = monthKey(r.data);
    m.set(k, (m.get(k) || 0) + r.qte);
  }
  const alerts = [];
  for (const [produto, months] of byProduct.entries()) {
    const sorted = [...months.entries()].sort((a, b) => a[0].localeCompare(b[0]));
    for (let i = 1; i < sorted.length; i++) {
      const [prevKey, prevVal] = sorted[i - 1];
      const [curKey, curVal] = sorted[i];
      if (prevVal === 0) continue;
      const pct = ((curVal - prevVal) / prevVal) * 100;
      if (Math.abs(pct) >= 15) {
        alerts.push({ produto, pct, prevKey, curKey, prevVal, curVal });
      }
    }
  }
  return alerts.sort((a, b) => Math.abs(b.pct) - Math.abs(a.pct));
}

/* ---------------------------------------------------------------------
   8. Render: KPIs, gráficos, alertas, tabela
   --------------------------------------------------------------------- */
function fmtNum(n) {
  return new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 1 }).format(n);
}
function fmtDate(d) {
  return d.toLocaleDateString('pt-BR');
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function render() {
  const rows = state.filtered;
  renderKPIs(rows);
  renderCharts(rows);
  renderAlerts(rows);
  renderTable(rows);
}

function renderKPIs(rows) {
  const total = rows.reduce((s, r) => s + r.qte, 0);
  const produtos = new Set(rows.map(r => r.produto)).size;
  const setores = new Set(rows.map(r => r.setor)).size;
  const movimentacoes = rows.length;
  const meses = aggregateByMonth(rows);
  const mediaMensal = meses.length ? total / meses.length : 0;
  const dias = rows.length ? (Math.max(...rows.map(r => r.data)) - Math.min(...rows.map(r => r.data))) / 86400000 + 1 : 1;
  const mediaDiaria = total / Math.max(dias, 1);
  const ranking = aggregateByProduct(rows);
  const maiorConsumidor = ranking.length ? ranking[0][0] : '—';

  const kpis = [
    { label: 'Consumo total', value: fmtNum(total), sub: `${movimentacoes} movimentações`, accent: 'teal' },
    { label: 'Produtos', value: produtos, sub: 'itens distintos', accent: 'violet' },
    { label: 'Setores', value: setores, sub: 'setores envolvidos', accent: 'violet' },
    { label: 'Média mensal', value: fmtNum(mediaMensal), sub: `${meses.length} mês(es) no período`, accent: 'amber' },
    { label: 'Média diária', value: fmtNum(mediaDiaria), sub: 'unidades/dia', accent: 'amber' },
    { label: 'Maior consumidor', value: maiorConsumidor, sub: 'produto com maior volume', accent: 'coral', isText: true },
  ];

  const grid = document.getElementById('kpiGrid');
  grid.innerHTML = kpis.map(k => `
    <div class="kpi" style="--kpi-accent: var(--accent-${k.accent})">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value" style="${k.isText ? 'font-size:16px; font-family:var(--font-body);' : ''}">${escapeHtml(String(k.value))}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>
  `).join('');
}

function destroyChart(id) { if (state.charts[id]) { state.charts[id].destroy(); delete state.charts[id]; } }

function baseOptions(extra = {}) {
  return Object.assign({
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { labels: { color: COLORS.dim, font: { family: 'Inter', size: 11 } } }
    },
    scales: {
      x: { ticks: { color: COLORS.dim, font: { size: 11 } }, grid: { color: COLORS.grid } },
      y: { ticks: { color: COLORS.dim, font: { size: 11 } }, grid: { color: COLORS.grid } }
    }
  }, extra);
}

function renderCharts(rows) {
  // Evolução mensal
  const meses = aggregateByMonth(rows);
  destroyChart('evolucao');
  state.charts.evolucao = new Chart(document.getElementById('chartEvolucao'), {
    type: 'line',
    data: {
      labels: meses.map(m => monthLabel(m[0])),
      datasets: [{
        label: 'Consumo',
        data: meses.map(m => m[1]),
        borderColor: COLORS.teal,
        backgroundColor: 'rgba(47,212,184,0.12)',
        tension: 0.35,
        fill: true,
        pointRadius: 3,
        pointBackgroundColor: COLORS.teal
      }]
    },
    options: baseOptions({ plugins: { legend: { display: false } } })
  });

  // Participação por setor (pizza)
  const setores = aggregateBySector(rows).slice(0, 8);
  destroyChart('setor');
  state.charts.setor = new Chart(document.getElementById('chartSetor'), {
    type: 'doughnut',
    data: {
      labels: setores.map(s => s[0]),
      datasets: [{
        data: setores.map(s => s[1]),
        backgroundColor: [COLORS.teal, COLORS.amber, COLORS.coral, COLORS.violet, '#4C8DF0', '#F0D24E', '#7BF04E', '#B0B7C3'],
        borderColor: '#161D27',
        borderWidth: 2
      }]
    },
    options: baseOptions({
      scales: {},
      plugins: { legend: { position: 'bottom', labels: { color: COLORS.dim, font: { size: 10 }, boxWidth: 10 } } }
    })
  });

  // Ranking de produtos (top 10, barra horizontal)
  const ranking = aggregateByProduct(rows).slice(0, 10);
  destroyChart('ranking');
  state.charts.ranking = new Chart(document.getElementById('chartRankingProdutos'), {
    type: 'bar',
    data: {
      labels: ranking.map(r => r[0].length > 22 ? r[0].slice(0, 22) + '…' : r[0]),
      datasets: [{ data: ranking.map(r => r[1]), backgroundColor: COLORS.teal, borderRadius: 4 }]
    },
    options: baseOptions({
      indexAxis: 'y',
      plugins: { legend: { display: false } }
    })
  });

  // Comparação entre setores (barras)
  const comp = aggregateBySector(rows).slice(0, 10);
  destroyChart('comparaSetor');
  state.charts.comparaSetor = new Chart(document.getElementById('chartComparaSetor'), {
    type: 'bar',
    data: {
      labels: comp.map(s => s[0].length > 16 ? s[0].slice(0, 16) + '…' : s[0]),
      datasets: [{ data: comp.map(s => s[1]), backgroundColor: COLORS.violet, borderRadius: 4 }]
    },
    options: baseOptions({ plugins: { legend: { display: false } } })
  });
}

function renderAlerts(rows) {
  const alerts = productMonthlyChange(rows).slice(0, 8);
  const list = document.getElementById('alertsList');
  if (!alerts.length) {
    list.innerHTML = '<div class="no-alerts">Nenhuma variação relevante (≥15%) entre meses consecutivos no período filtrado.</div>';
    return;
  }
  list.innerHTML = alerts.map(a => `
    <div class="alert-row">
      <span><span class="alert-name">${escapeHtml(a.produto)}</span><span class="alert-meta">${monthLabel(a.prevKey)} → ${monthLabel(a.curKey)} · ${fmtNum(a.prevVal)} → ${fmtNum(a.curVal)}</span></span>
      <span class="badge ${a.pct >= 0 ? 'up' : 'down'}">${a.pct >= 0 ? '+' : ''}${a.pct.toFixed(0)}%</span>
    </div>
  `).join('');
}

function renderTable(rows) {
  const body = document.getElementById('dataTableBody');
  const sample = rows.slice(0, 300); // limita render para performance
  document.getElementById('tableCount').textContent = `${rows.length} registros${rows.length > 300 ? ' (exibindo 300)' : ''}`;
  body.innerHTML = sample.map(r => `
    <tr>
      <td>${fmtDate(r.data)}</td>
      <td>${escapeHtml(r.produto)}</td>
      <td>${escapeHtml(r.setor)}</td>
      <td>${fmtNum(r.qte)}</td>
      <td>${escapeHtml(r.unidade)}</td>
      <td>${r.total !== null ? fmtNum(r.total) : '—'}</td>
    </tr>
  `).join('');
}
