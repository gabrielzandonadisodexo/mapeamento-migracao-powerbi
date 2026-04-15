const numberFormatter = new Intl.NumberFormat('pt-BR');
const percentFormatter = new Intl.NumberFormat('pt-BR', {
  maximumFractionDigits: 1,
  minimumFractionDigits: 1
});

const navLinks = Array.from(document.querySelectorAll('.nav-item'));
const bodyPage = document.body.dataset.page || 'summary';
const similarityBands = [
  { id: '0-20', label: '0-20%', min: 0, max: 20, tone: 'tone-red', description: 'Baixa sobreposicao' },
  { id: '20-40', label: '20-40%', min: 20, max: 40, tone: 'tone-orange', description: 'Sobreposicao inicial' },
  { id: '40-60', label: '40-60%', min: 40, max: 60, tone: 'tone-yellow', description: 'Sobreposicao media' },
  { id: '60-80', label: '60-80%', min: 60, max: 80, tone: 'tone-lime', description: 'Alta sobreposicao' },
  { id: '80-100', label: '80-100%', min: 80, max: 100, tone: 'tone-green', description: 'Maior reaproveitamento' }
];
let similarityState = null;
const sectionLinks = navLinks
  .map((link) => ({ link, section: getSectionForLink(link) }))
  .filter((item) => item.section);

function getSectionForLink(link) {
  const href = link.getAttribute('href') || '';
  const url = new URL(href, window.location.href);
  const samePath = normalizePath(url.pathname) === normalizePath(window.location.pathname);

  if (samePath && url.hash) {
    return document.querySelector(url.hash);
  }

  if (samePath && !url.hash && bodyPage === 'summary') {
    return document.querySelector('#resumo');
  }

  return null;
}

function normalizePath(pathname) {
  return pathname.replace(/\/index\.html$/, '/');
}

function updateActiveLink() {
  if (!sectionLinks.length) return;

  const current = sectionLinks
    .map(({ section }) => ({
      id: section.id,
      top: Math.abs(section.getBoundingClientRect().top)
    }))
    .sort((a, b) => a.top - b.top)[0];

  if (!current) return;

  for (const link of navLinks) {
    const section = getSectionForLink(link);
    link.classList.toggle('active', section?.id === current.id);
  }
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return '';
  const units = ['B', 'KB', 'MB', 'GB'];
  let size = bytes;
  let unit = 0;
  while (size >= 1024 && unit < units.length - 1) {
    size /= 1024;
    unit += 1;
  }
  return `${size.toLocaleString('pt-BR', { maximumFractionDigits: 1 })} ${units[unit]}`;
}

function formatNumber(value) {
  return numberFormatter.format(Number(value) || 0);
}

function formatPercent(value) {
  return `${percentFormatter.format(Number(value) || 0)}%`;
}

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

async function fetchJson(path) {
  const response = await fetch(path, { cache: 'no-store' });
  if (!response.ok) {
    throw new Error(`Falha ao carregar ${path}`);
  }
  return response.json();
}

async function loadManifest() {
  try {
    const manifest = await fetchJson('./manifest.json');

    for (const [key, file] of Object.entries(manifest.files || {})) {
      const node = document.querySelector(`[data-file="${key}"]`);
      if (!node) continue;
      const size = formatBytes(file.size);
      node.textContent = [size, file.updatedAt].filter(Boolean).join(' | ');
    }

    for (const [key, value] of Object.entries(manifest.metrics || {})) {
      const node = document.querySelector(`[data-metric="${key}"]`);
      if (node) node.textContent = formatNumber(value);
    }
  } catch {
    // Metadata is optional; downloads stay available without it.
  }
}

function setText(selector, value) {
  const node = document.querySelector(selector);
  if (node) node.textContent = value;
}

function setMetric(scope, key, value) {
  const node = document.querySelector(`[${scope}="${key}"]`);
  if (node) node.textContent = value;
}

function clearNode(node) {
  while (node.firstChild) node.removeChild(node.firstChild);
}

function createEmptyMessage(message, colspan = 1) {
  const row = document.createElement('tr');
  const cell = document.createElement('td');
  cell.colSpan = colspan;
  cell.className = 'empty-cell';
  cell.textContent = message;
  row.appendChild(cell);
  return row;
}

function createStatusPill(status) {
  const pill = document.createElement('span');
  pill.className = `status-pill ${statusClass(status)}`;
  pill.textContent = status || 'Nao informado';
  return pill;
}

function statusClass(status) {
  const normalized = normalizeText(status);
  if (normalized.includes('mapeado')) return 'is-ok';
  if (normalized.includes('erro')) return 'is-error';
  if (normalized.includes('datasource')) return 'is-warning';
  return 'is-review';
}

function renderBarList(containerId, items, labelSelector, valueSelector) {
  const container = document.getElementById(containerId);
  if (!container) return;
  clearNode(container);

  if (!items?.length) {
    const empty = document.createElement('p');
    empty.className = 'empty-state';
    empty.textContent = 'Sem dados para exibir.';
    container.appendChild(empty);
    return;
  }

  const max = Math.max(...items.map((item) => Number(valueSelector(item)) || 0), 1);

  for (const item of items) {
    const value = Number(valueSelector(item)) || 0;
    const row = document.createElement('div');
    row.className = 'bar-row';

    const meta = document.createElement('div');
    meta.className = 'bar-meta';

    const label = document.createElement('span');
    label.className = 'bar-label';
    label.textContent = labelSelector(item);

    const valueNode = document.createElement('strong');
    valueNode.className = 'bar-value';
    valueNode.textContent = formatNumber(value);

    const track = document.createElement('div');
    track.className = 'bar-track';

    const fill = document.createElement('span');
    fill.className = `bar-fill ${statusClass(item.name || item.status || '')}`;
    fill.style.width = `${Math.max((value / max) * 100, 3)}%`;

    meta.append(label, valueNode);
    track.appendChild(fill);
    row.append(meta, track);
    container.appendChild(row);
  }
}

async function renderAnalyticalPage() {
  let data;
  try {
    data = await fetchJson('./data/analytics.json');
  } catch {
    setText('#statusBars', 'Nao foi possivel carregar a base analitica.');
    return;
  }

  const summary = data.summary || {};
  setMetric('data-analytics', 'coberturaTabelaPct', formatPercent(summary.coberturaTabelaPct));
  setMetric('data-analytics', 'relatoriosComPendencia', formatNumber(summary.relatoriosComPendencia));
  setMetric('data-analytics', 'fontesComPendencia', formatNumber(summary.fontesComPendencia));
  setMetric('data-analytics', 'linhasSemTabela', formatNumber(summary.linhasSemTabela));

  const charts = data.charts || {};
  renderBarList('statusBars', charts.status, (item) => item.name, (item) => item.value);
  renderBarList('schemaBars', charts.topSchemas, (item) => item.name, (item) => item.value);
  renderBarList('tableBars', charts.topTables, (item) => item.name, (item) => item.value);
  renderBarList(
    'reportBars',
    charts.topReportsByLines,
    (item) => `${item.relatorio} | ${formatNumber(item.fontes)} fontes`,
    (item) => item.linhas
  );

  renderIssueTable(charts.reportsWithIssues || []);
}

function renderIssueTable(items) {
  const table = document.getElementById('issueTable');
  if (!table) return;
  clearNode(table);

  if (!items.length) {
    table.appendChild(createEmptyMessage('Nao ha pontos pendentes na base carregada.', 4));
    return;
  }

  for (const item of items) {
    const row = document.createElement('tr');

    const reportCell = document.createElement('td');
    if (item.link_relatorio?.startsWith('http')) {
      const link = document.createElement('a');
      link.href = item.link_relatorio;
      link.target = '_blank';
      link.rel = 'noreferrer';
      link.textContent = item.relatorio || 'Relatorio';
      reportCell.appendChild(link);
    } else {
      reportCell.textContent = item.relatorio || 'Relatorio';
    }

    const pendingCell = document.createElement('td');
    pendingCell.appendChild(createStatusPill(`${formatNumber(item.pendencias)} pendencias`));

    const sourceCell = document.createElement('td');
    sourceCell.textContent = formatNumber(item.fontes);

    const tableCell = document.createElement('td');
    tableCell.textContent = formatNumber(item.tabelas);

    row.append(reportCell, pendingCell, sourceCell, tableCell);
    table.appendChild(row);
  }
}

async function renderDetailedPage() {
  let data;
  try {
    data = await fetchJson('./data/details.json');
  } catch {
    setText('#detailCount', 'Nao foi possivel carregar a base detalhada.');
    return;
  }

  const summary = data.summary || {};
  for (const [key, value] of Object.entries(summary)) {
    setMetric('data-detail-metric', key, formatNumber(value));
  }

  const rows = Array.isArray(data.rows) ? data.rows : [];
  const state = {
    rows,
    filteredRows: rows,
    page: 1,
    pageSize: 50
  };

  renderSimilarityCard(rows);
  populateSelect('statusFilter', data.filters?.statuses || []);
  populateSelect('schemaFilter', data.filters?.schemas || []);
  populateSelect('reportFilter', data.filters?.reports || []);

  const applyAndRender = () => {
    state.page = 1;
    state.filteredRows = filterDetailRows(state.rows);
    renderDetailRows(state);
  };

  for (const id of ['searchInput', 'statusFilter', 'schemaFilter', 'reportFilter']) {
    document.getElementById(id)?.addEventListener('input', applyAndRender);
  }

  document.getElementById('prevPage')?.addEventListener('click', () => {
    state.page = Math.max(state.page - 1, 1);
    renderDetailRows(state);
  });

  document.getElementById('nextPage')?.addEventListener('click', () => {
    const pageCount = getPageCount(state);
    state.page = Math.min(state.page + 1, pageCount);
    renderDetailRows(state);
  });

  renderDetailRows(state);
}

function renderSimilarityCard(rows) {
  const model = buildSimilarityModel(rows);
  similarityState = {
    model,
    activeBandId: null,
    activePairId: null
  };

  renderSimilarityBands();
  renderSimilarityPanel();
}

function buildSimilarityModel(rows) {
  const reports = buildReportTableSets(rows);
  const pairs = [];

  for (let indexA = 0; indexA < reports.length; indexA += 1) {
    for (let indexB = indexA + 1; indexB < reports.length; indexB += 1) {
      const reportA = reports[indexA];
      const reportB = reports[indexB];
      const smaller = reportA.tableKeys.length <= reportB.tableKeys.length ? reportA : reportB;
      const larger = smaller === reportA ? reportB : reportA;
      const commonKeys = smaller.tableKeys.filter((key) => larger.tables.has(key));
      if (!commonKeys.length) continue;

      const denominator = Math.max(Math.min(reportA.tableKeys.length, reportB.tableKeys.length), 1);
      const similarityPct = Number(((commonKeys.length / denominator) * 100).toFixed(1));
      const band = getSimilarityBand(similarityPct);
      const commonTables = commonKeys
        .map((key) => reportA.tables.get(key) || reportB.tables.get(key))
        .sort((a, b) => a.localeCompare(b, 'pt-BR'));
      const onlyA = reportA.tableKeys
        .filter((key) => !reportB.tables.has(key))
        .map((key) => reportA.tables.get(key))
        .sort((a, b) => a.localeCompare(b, 'pt-BR'));
      const onlyB = reportB.tableKeys
        .filter((key) => !reportA.tables.has(key))
        .map((key) => reportB.tables.get(key))
        .sort((a, b) => a.localeCompare(b, 'pt-BR'));

      pairs.push({
        id: `${indexA}-${indexB}`,
        bandId: band.id,
        reportA,
        reportB,
        similarityPct,
        commonCount: commonTables.length,
        commonTables,
        onlyA,
        onlyB
      });
    }
  }

  pairs.sort((a, b) => (
    b.similarityPct - a.similarityPct ||
    b.commonCount - a.commonCount ||
    a.reportA.name.localeCompare(b.reportA.name, 'pt-BR')
  ));

  const pairsByBand = new Map(similarityBands.map((band) => [band.id, []]));
  for (const pair of pairs) {
    pairsByBand.get(pair.bandId)?.push(pair);
  }

  return {
    reports,
    pairs,
    pairsByBand
  };
}

function buildReportTableSets(rows) {
  const reports = new Map();

  for (const row of rows) {
    if (!row.tabela) continue;
    const reportName = row.relatorio || 'Relatorio sem nome';
    const table = getTableIdentity(row);
    if (!table.key) continue;

    if (!reports.has(reportName)) {
      reports.set(reportName, {
        name: reportName,
        link: row.link_relatorio || '',
        tables: new Map()
      });
    }

    const report = reports.get(reportName);
    if (!report.tables.has(table.key)) {
      report.tables.set(table.key, table.label);
    }
  }

  return [...reports.values()]
    .map((report) => ({
      ...report,
      tableKeys: [...report.tables.keys()].sort((a, b) => a.localeCompare(b, 'pt-BR'))
    }))
    .filter((report) => report.tableKeys.length > 0)
    .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
}

function getTableIdentity(row) {
  const schema = String(row.schema || '').trim();
  const table = String(row.tabela || '').trim();
  if (!table) return { key: '', label: '' };
  const label = schema ? `${schema}.${table}` : table;
  return {
    key: normalizeText(label),
    label
  };
}

function getSimilarityBand(percent) {
  return similarityBands.find((band) => (
    percent >= band.min && (percent < band.max || (band.max === 100 && percent <= 100))
  )) || similarityBands[0];
}

function renderSimilarityBands() {
  const container = document.getElementById('similarityBands');
  if (!container || !similarityState) return;
  clearNode(container);

  for (const band of similarityBands) {
    const pairs = similarityState.model.pairsByBand.get(band.id) || [];
    const maxCommon = Math.max(...pairs.map((pair) => pair.commonCount), 0);
    const button = document.createElement('button');
    button.className = `similarity-band ${band.tone}`;
    button.type = 'button';
    button.dataset.bandId = band.id;
    button.setAttribute('aria-expanded', String(similarityState.activeBandId === band.id));

    const label = document.createElement('strong');
    label.textContent = band.label;

    const description = document.createElement('span');
    description.textContent = band.description;

    const count = document.createElement('small');
    count.textContent = `${formatNumber(pairs.length)} pares | ate ${formatNumber(maxCommon)} tabelas em comum`;

    button.append(label, description, count);
    button.addEventListener('click', () => {
      similarityState.activeBandId = similarityState.activeBandId === band.id ? null : band.id;
      similarityState.activePairId = null;
      renderSimilarityBands();
      renderSimilarityPanel();
    });
    container.appendChild(button);
  }
}

function renderSimilarityPanel() {
  const panel = document.getElementById('similarityPanel');
  if (!panel || !similarityState) return;
  clearNode(panel);

  if (!similarityState.activeBandId) {
    const hint = document.createElement('p');
    hint.className = 'similarity-hint';
    hint.textContent = 'Clique em uma faixa para ver quais relatorios compartilham tabelas. Depois clique em um par para ver as tabelas em comum e as exclusivas de cada relatorio.';
    panel.appendChild(hint);
    return;
  }

  const band = similarityBands.find((item) => item.id === similarityState.activeBandId);
  const pairs = similarityState.model.pairsByBand.get(similarityState.activeBandId) || [];

  const heading = document.createElement('div');
  heading.className = 'similarity-panel-heading';
  const title = document.createElement('h3');
  title.textContent = `Faixa ${band.label}`;
  const subtitle = document.createElement('p');
  subtitle.textContent = `${formatNumber(pairs.length)} pares de relatorios com pelo menos uma tabela em comum.`;
  heading.append(title, subtitle);
  panel.appendChild(heading);

  if (!pairs.length) {
    const empty = document.createElement('p');
    empty.className = 'empty-state';
    empty.textContent = 'Nao foram encontrados relatorios nessa faixa.';
    panel.appendChild(empty);
    return;
  }

  const list = document.createElement('div');
  list.className = 'similarity-pair-list';

  for (const pair of pairs) {
    list.appendChild(createSimilarityPairNode(pair));
  }

  panel.appendChild(list);
}

function createSimilarityPairNode(pair) {
  const item = document.createElement('article');
  item.className = 'similarity-pair';

  const toggle = document.createElement('button');
  toggle.className = 'similarity-pair-toggle';
  toggle.type = 'button';
  toggle.dataset.pairId = pair.id;
  toggle.setAttribute('aria-expanded', String(similarityState.activePairId === pair.id));

  const reports = document.createElement('span');
  reports.className = 'similarity-reports';
  reports.textContent = `${pair.reportA.name} | ${pair.reportB.name}`;

  const stats = document.createElement('span');
  stats.className = 'similarity-stats';
  stats.textContent = `${formatPercent(pair.similarityPct)} similares | ${formatNumber(pair.commonCount)} tabelas em comum`;

  toggle.append(reports, stats);
  toggle.addEventListener('click', () => {
    similarityState.activePairId = similarityState.activePairId === pair.id ? null : pair.id;
    renderSimilarityPanel();
  });

  item.appendChild(toggle);

  if (similarityState.activePairId === pair.id) {
    item.appendChild(createSimilarityDetailsNode(pair));
  }

  return item;
}

function createSimilarityDetailsNode(pair) {
  const details = document.createElement('div');
  details.className = 'similarity-details';

  details.appendChild(createTableGroup('Tabelas em comum', pair.commonTables, 'is-common'));
  details.appendChild(createTableGroup(`Somente em ${pair.reportA.name}`, pair.onlyA));
  details.appendChild(createTableGroup(`Somente em ${pair.reportB.name}`, pair.onlyB));

  return details;
}

function createTableGroup(title, tables, className = '') {
  const group = document.createElement('div');
  group.className = `similarity-table-group ${className}`.trim();

  const heading = document.createElement('h4');
  heading.textContent = `${title} (${formatNumber(tables.length)})`;
  group.appendChild(heading);

  if (!tables.length) {
    const empty = document.createElement('p');
    empty.className = 'empty-state';
    empty.textContent = 'Nenhuma tabela exclusiva.';
    group.appendChild(empty);
    return group;
  }

  const list = document.createElement('ul');
  for (const table of tables) {
    const item = document.createElement('li');
    item.textContent = table;
    list.appendChild(item);
  }
  group.appendChild(list);
  return group;
}

function populateSelect(id, values) {
  const select = document.getElementById(id);
  if (!select) return;
  clearNode(select);

  const all = document.createElement('option');
  all.value = '';
  all.textContent = 'Todos';
  select.appendChild(all);

  for (const value of values) {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  }
}

function filterDetailRows(rows) {
  const term = normalizeText(document.getElementById('searchInput')?.value);
  const status = document.getElementById('statusFilter')?.value || '';
  const schema = document.getElementById('schemaFilter')?.value || '';
  const report = document.getElementById('reportFilter')?.value || '';

  return rows.filter((row) => {
    const rowSchema = row.schema || 'Sem schema';
    const matchesStatus = !status || row.status === status;
    const matchesSchema = !schema || rowSchema === schema;
    const matchesReport = !report || row.relatorio === report;
    const haystack = normalizeText([
      row.relatorio,
      row.datasource,
      row.schema,
      row.tabela,
      row.status,
      row.link_datasource
    ].join(' '));
    return matchesStatus && matchesSchema && matchesReport && (!term || haystack.includes(term));
  });
}

function getPageCount(state) {
  return Math.max(Math.ceil(state.filteredRows.length / state.pageSize), 1);
}

function renderDetailRows(state) {
  const table = document.getElementById('detailRows');
  if (!table) return;
  clearNode(table);

  const pageCount = getPageCount(state);
  state.page = Math.min(state.page, pageCount);
  const start = (state.page - 1) * state.pageSize;
  const visibleRows = state.filteredRows.slice(start, start + state.pageSize);

  if (!visibleRows.length) {
    table.appendChild(createEmptyMessage('Nenhuma linha encontrada para os filtros selecionados.', 6));
  }

  for (const row of visibleRows) {
    const tableRow = document.createElement('tr');
    appendCell(tableRow, row.relatorio || 'Sem relatorio', 'wide-cell');
    appendCell(tableRow, row.datasource || 'Sem datasource');
    appendCell(tableRow, row.schema || 'Sem schema');
    appendCell(tableRow, row.tabela || 'Sem tabela');

    const statusCell = document.createElement('td');
    statusCell.appendChild(createStatusPill(row.status));
    tableRow.appendChild(statusCell);

    const linkCell = document.createElement('td');
    linkCell.appendChild(createLinkNode(row));
    tableRow.appendChild(linkCell);

    table.appendChild(tableRow);
  }

  setText('#detailCount', `${formatNumber(state.filteredRows.length)} linhas encontradas`);
  setText('#pageInfo', `Pagina ${formatNumber(state.page)} de ${formatNumber(pageCount)}`);

  const prev = document.getElementById('prevPage');
  const next = document.getElementById('nextPage');
  if (prev) prev.disabled = state.page <= 1;
  if (next) next.disabled = state.page >= pageCount;
}

function appendCell(row, text, className = '') {
  const cell = document.createElement('td');
  if (className) cell.className = className;
  cell.textContent = text;
  row.appendChild(cell);
}

function createLinkNode(row) {
  const link = row.link_datasource || '';
  if (link.startsWith('http')) {
    const anchor = document.createElement('a');
    anchor.href = link;
    anchor.target = '_blank';
    anchor.rel = 'noreferrer';
    anchor.textContent = 'Abrir fonte';
    return anchor;
  }

  const note = document.createElement('span');
  note.className = 'muted-note';
  note.textContent = link.startsWith('ERRO:') ? 'Erro no workbook' : 'Sem link direto';
  note.title = link || 'Sem link de datasource';
  return note;
}

window.addEventListener('scroll', updateActiveLink, { passive: true });
window.addEventListener('resize', updateActiveLink);
updateActiveLink();
loadManifest();

if (bodyPage === 'analytical') {
  renderAnalyticalPage();
}

if (bodyPage === 'detailed') {
  renderDetailedPage();
}
