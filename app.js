const numberFormatter = new Intl.NumberFormat('pt-BR');
const percentFormatter = new Intl.NumberFormat('pt-BR', {
  maximumFractionDigits: 1,
  minimumFractionDigits: 1
});

const navLinks = Array.from(document.querySelectorAll('.nav-item'));
const bodyPage = document.body.dataset.page || 'summary';
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
