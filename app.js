const numberFormatter = new Intl.NumberFormat('pt-BR');
const percentFormatter = new Intl.NumberFormat('pt-BR', {
  maximumFractionDigits: 1,
  minimumFractionDigits: 1
});

const SEGMENTS = {
  finance: {
    uiLabel: 'Finance',
    brand: 'Finanças',
    summaryEyebrow: 'Tableau | Finanças | Inventário técnico',
    summaryTitle: 'Mapeamento de fontes e tabelas dos dashboards de Finanças',
    summaryLead: 'Inventário consolidado para apoiar análise de origem dos dados, dependências entre relatórios e priorização do mapeamento de Finanças.',
    analyticalEyebrow: 'Analytical | Finanças',
    analyticalTitle: 'Visão analítica das fontes e tabelas mapeadas',
    analyticalLead: 'Indicadores e gráficos calculados a partir da base consolidada do segmento de Finanças.',
    detailedEyebrow: 'Detailed | Finanças',
    detailedTitle: 'Consulta detalhada do inventário',
    detailedLead: 'Explore as linhas do inventário de Finanças com filtros por relatório, schema, status e busca livre.',
    excelDescription: 'Base consolidada com relatório, datasource, link, schema e tabela para o segmento de Finanças.',
    pdfDescription: 'Resumo executivo, método usado, números finais e pontos de revisão do mapeamento de Finanças.',
    reviewLabels: {
      erros: 'Painéis com falha ao localizar workbook ou abrir o fluxo padrão',
      semTabela: 'Fontes sem tabela identificável por SQL ou lineage',
      semDatasource: 'Painéis sem datasource encontrado no workbook ou lineage'
    }
  },
  hr: {
    uiLabel: 'HR',
    brand: 'HR',
    summaryEyebrow: 'Tableau | HR | Inventário técnico',
    summaryTitle: 'Mapeamento de fontes e tabelas dos dashboards de HR',
    summaryLead: 'Inventário consolidado para apoiar leitura da origem dos dados, dependências entre relatórios e priorização do mapeamento de HR.',
    analyticalEyebrow: 'Analytical | HR',
    analyticalTitle: 'Visão analítica das fontes e tabelas mapeadas',
    analyticalLead: 'Indicadores e gráficos calculados a partir da base consolidada do segmento de HR.',
    detailedEyebrow: 'Detailed | HR',
    detailedTitle: 'Consulta detalhada do inventário',
    detailedLead: 'Explore as linhas do inventário de HR com filtros por relatório, schema, status e busca livre.',
    excelDescription: 'Base consolidada com relatório, datasource, link, schema e tabela para o segmento de HR.',
    pdfDescription: 'Resumo executivo, método usado, números finais e pontos de revisão do mapeamento de HR.',
    reviewLabels: {
      erros: 'Painéis com falha ao localizar workbook ou abrir o fluxo padrão',
      semTabela: 'Fontes sem tabela identificável por SQL ou lineage',
      semDatasource: 'Painéis sem datasource encontrado no workbook ou lineage'
    }
  }
};

const PAGE_LABELS = {
  summary: 'Resumo',
  analytical: 'Analytical',
  detailed: 'Detailed'
};

const similarityBands = [
  { id: '0-20', label: '0-20%', min: 0, max: 20, tone: 'tone-red', description: 'Baixa sobreposição' },
  { id: '20-40', label: '20-40%', min: 20, max: 40, tone: 'tone-orange', description: 'Sobreposição inicial' },
  { id: '40-60', label: '40-60%', min: 40, max: 60, tone: 'tone-yellow', description: 'Sobreposição média' },
  { id: '60-80', label: '60-80%', min: 60, max: 80, tone: 'tone-lime', description: 'Alta sobreposição' },
  { id: '80-100', label: '80-100%', min: 80, max: 100, tone: 'tone-green', description: 'Maior reaproveitamento' }
];

const bodyPage = document.body.dataset.page || 'summary';
const navLinks = Array.from(document.querySelectorAll('.nav-item'));
const segmentLinks = Array.from(document.querySelectorAll('[data-segment-link]'));
const currentUrl = new URL(window.location.href);
const segmentKey = currentUrl.searchParams.get('segment') === 'hr' ? 'hr' : 'finance';
const segment = SEGMENTS[segmentKey];
const dataBase = `./data/${segmentKey}`;
let similarityState = null;

function normalizePath(pathname) {
  return pathname.replace(/\/index\.html$/, '/');
}

function withSegment(href, targetSegment = segmentKey) {
  const url = new URL(href, window.location.href);
  if (url.origin !== window.location.origin) return href;
  url.searchParams.set('segment', targetSegment);
  return `${url.pathname}${url.search}${url.hash}`;
}

function decorateSegmentAwareLinks() {
  for (const link of navLinks) {
    link.setAttribute('href', withSegment(link.getAttribute('href') || ''));
  }

  for (const link of segmentLinks) {
    const targetSegment = link.dataset.segmentLink || segmentKey;
    link.setAttribute('href', withSegment(link.getAttribute('href') || window.location.pathname, targetSegment));
    link.classList.toggle('active', targetSegment === segmentKey);
  }
}

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

function updateActiveLink() {
  const sectionLinks = navLinks
    .map((link) => ({ link, section: getSectionForLink(link) }))
    .filter((item) => item.section);

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
    if (section) {
      link.classList.toggle('active', section.id === current.id);
    }
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

async function fetchJson(filePath) {
  const response = await fetch(filePath, { cache: 'no-store' });
  if (!response.ok) {
    throw new Error(`Falha ao carregar ${filePath}`);
  }
  return response.json();
}

async function loadSegmentsCatalog() {
  try {
    const payload = await fetchJson('./data/segments.json');
    for (const link of segmentLinks) {
      const targetSegment = link.dataset.segmentLink;
      const entry = payload.segments.find((item) => item.segment === targetSegment);
      if (!entry?.available) {
        link.classList.add('is-disabled');
        link.setAttribute('aria-disabled', 'true');
        link.title = entry?.error || 'Segmento ainda não disponível';
      }
    }
  } catch {
    // Optional. Segment links still work when the catalog is unavailable.
  }
}

function applySegmentCopy() {
  document.title = `${PAGE_LABELS[bodyPage] || 'Mapa'} | ${segment.brand}`;
  setText('#brandName', segment.brand);
  setText('#summaryEyebrow', segment.summaryEyebrow);
  setText('#summaryTitle', segment.summaryTitle);
  setText('#summaryLead', segment.summaryLead);
  setText('#analyticalEyebrow', segment.analyticalEyebrow);
  setText('#analyticalTitle', segment.analyticalTitle);
  setText('#analyticalLead', segment.analyticalLead);
  setText('#detailedEyebrow', segment.detailedEyebrow);
  setText('#detailedTitle', segment.detailedTitle);
  setText('#detailedLead', segment.detailedLead);
  setText('#excelDescription', segment.excelDescription);
  setText('#pdfDescription', segment.pdfDescription);
  setText('#reviewErroLabel', segment.reviewLabels.erros);
  setText('#reviewSemTabelaLabel', segment.reviewLabels.semTabela);
  setText('#reviewSemDatasourceLabel', segment.reviewLabels.semDatasource);
}

function applyDownloadLinks(manifest) {
  const excel = manifest?.files?.excel;
  const pdf = manifest?.files?.pdf;

  if (excel) {
    setText('#excelName', excel.name);
    const summaryDownload = document.getElementById('excelDownload');
    const detailDownload = document.getElementById('detailExcelDownload');
    if (summaryDownload) summaryDownload.href = excel.href;
    if (detailDownload) detailDownload.href = excel.href;
  }

  if (pdf) {
    setText('#pdfName', pdf.name);
    const pdfDownload = document.getElementById('pdfDownload');
    if (pdfDownload) pdfDownload.href = pdf.href;
  } else {
    const pdfDownload = document.getElementById('pdfDownload');
    if (pdfDownload) {
      pdfDownload.classList.add('is-disabled');
      pdfDownload.removeAttribute('href');
      pdfDownload.textContent = 'PDF em preparação';
    }
  }
}

function applySummaryManifest(manifest) {
  for (const [key, value] of Object.entries(manifest.metrics || {})) {
    const node = document.querySelector(`[data-metric="${key}"]`);
    if (node) node.textContent = formatNumber(value);
  }

  setMetric('data-review', 'erros', formatNumber(manifest.metrics?.erros));
  setMetric('data-review', 'semTabela', formatNumber(manifest.metrics?.semTabela));
  setMetric('data-review', 'semDatasource', formatNumber(manifest.metrics?.semDatasource));

  for (const [key, file] of Object.entries(manifest.files || {})) {
    if (!file) continue;
    const node = document.querySelector(`[data-file="${key}"]`);
    if (!node) continue;
    const size = formatBytes(file.size);
    node.textContent = [size, file.updatedAt].filter(Boolean).join(' | ');
  }

  applyDownloadLinks(manifest);
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

async function loadManifest() {
  try {
    const manifest = await fetchJson(`${dataBase}/manifest.json`);
    applySummaryManifest(manifest);
    return manifest;
  } catch {
    return null;
  }
}

async function renderAnalyticalPage() {
  let data;
  try {
    data = await fetchJson(`${dataBase}/analytics.json`);
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

async function renderDetailedPage() {
  let data;
  try {
    data = await fetchJson(`${dataBase}/details.json`);
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
    document.getElementById(id)?.addEventListener('change', applyAndRender);
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
  const defaultBand = [...similarityBands]
    .reverse()
    .find((band) => (model.pairsByBand.get(band.id) || []).length > 0);
  const defaultPair = defaultBand ? model.pairsByBand.get(defaultBand.id)?.[0] : null;

  similarityState = {
    model,
    activeBandId: defaultBand?.id || null,
    activePairId: defaultPair?.id || null
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

  return { reports, pairs, pairsByBand };
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
      const nextBandId = similarityState.activeBandId === band.id ? null : band.id;
      const nextPairs = nextBandId ? similarityState.model.pairsByBand.get(nextBandId) || [] : [];
      similarityState.activeBandId = nextBandId;
      similarityState.activePairId = nextPairs[0]?.id || null;
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
    hint.textContent = 'Selecione uma faixa para priorizar relatórios com tabelas parecidas e avaliar onde um mesmo mapeamento pode atender mais de um report.';
    panel.appendChild(hint);
    return;
  }

  const band = similarityBands.find((item) => item.id === similarityState.activeBandId);
  const pairs = similarityState.model.pairsByBand.get(similarityState.activeBandId) || [];
  if (!pairs.length) {
    const empty = document.createElement('p');
    empty.className = 'empty-state';
    empty.textContent = 'Nao foram encontrados relatórios nessa faixa.';
    panel.appendChild(empty);
    return;
  }

  if (!pairs.some((pair) => pair.id === similarityState.activePairId)) {
    similarityState.activePairId = pairs[0]?.id || null;
  }

  const displayedPairs = pairs.slice(0, 12);
  const board = document.createElement('div');
  board.className = 'similarity-board';

  const header = document.createElement('div');
  header.className = 'similarity-board-header';
  const title = document.createElement('h3');
  title.textContent = `Melhores oportunidades em ${band.label}`;
  const subtitle = document.createElement('p');
  subtitle.textContent = 'Abra um par de relatórios para ver quais tabelas podem ser mapeadas uma vez e reaproveitadas nos dois.';
  header.append(title, subtitle);
  board.appendChild(header);

  const summary = document.createElement('div');
  summary.className = 'similarity-summary-grid';
  summary.appendChild(createSummaryTile('Pares nessa faixa', pairs.length));
  summary.appendChild(createSummaryTile('Maior reaproveitamento', Math.max(...pairs.map((pair) => pair.commonCount), 0), 'tabelas'));
  summary.appendChild(createSummaryTile('Lista priorizada', displayedPairs.length, 'pares'));
  board.appendChild(summary);

  const accordion = document.createElement('div');
  accordion.className = 'similarity-accordion';
  for (const [index, pair] of displayedPairs.entries()) {
    accordion.appendChild(createSimilarityAccordionNode(pair, index + 1));
  }
  board.appendChild(accordion);
  panel.appendChild(board);
}

function createSummaryTile(label, value, suffix = '') {
  const tile = document.createElement('article');
  tile.className = 'similarity-summary-tile';
  const number = document.createElement('strong');
  number.textContent = formatNumber(value);
  const text = document.createElement('span');
  text.textContent = suffix ? `${label} (${suffix})` : label;
  tile.append(number, text);
  return tile;
}

function createSimilarityAccordionNode(pair, rank) {
  const item = document.createElement('article');
  item.className = 'similarity-accordion-item';
  const isExpanded = similarityState.activePairId === pair.id;

  const trigger = document.createElement('button');
  trigger.className = 'similarity-accordion-trigger';
  trigger.type = 'button';
  trigger.setAttribute('aria-expanded', String(isExpanded));

  const rankNode = document.createElement('span');
  rankNode.className = 'similarity-rank';
  rankNode.textContent = `#${String(rank).padStart(2, '0')}`;

  const names = document.createElement('span');
  names.className = 'similarity-pair-summary';
  names.appendChild(createReportSummary('Relatório A', pair.reportA.name));
  names.appendChild(createReportSummary('Relatório B', pair.reportB.name));

  const metrics = document.createElement('span');
  metrics.className = 'similarity-trigger-metrics';
  metrics.appendChild(createMetricBadge(formatPercent(pair.similarityPct), 'similaridade'));
  metrics.appendChild(createMetricBadge(formatNumber(pair.commonCount), 'tabelas comuns', 'is-strong'));

  trigger.append(rankNode, names, metrics);
  trigger.addEventListener('click', () => {
    similarityState.activePairId = isExpanded ? null : pair.id;
    renderSimilarityPanel();
  });

  item.appendChild(trigger);
  if (isExpanded) {
    item.appendChild(createSimilarityAccordionBody(pair));
  }
  return item;
}

function createReportSummary(label, name) {
  const wrapper = document.createElement('span');
  wrapper.className = 'similarity-report-summary';
  const small = document.createElement('small');
  small.textContent = label;
  const strong = document.createElement('strong');
  strong.textContent = name;
  wrapper.append(small, strong);
  return wrapper;
}

function createMetricBadge(value, label, className = '') {
  const badge = document.createElement('span');
  badge.className = `similarity-metric-badge ${className}`.trim();
  const strong = document.createElement('strong');
  strong.textContent = value;
  const small = document.createElement('small');
  small.textContent = label;
  badge.append(strong, small);
  return badge;
}

function createSimilarityAccordionBody(pair) {
  const body = document.createElement('div');
  body.className = 'similarity-accordion-body';

  const recommendation = document.createElement('p');
  recommendation.className = 'similarity-recommendation';
  recommendation.textContent = `Priorize estas ${formatNumber(pair.commonCount)} tabelas quando o objetivo for atender os dois relatórios com o mesmo trabalho de mapeamento.`;
  body.appendChild(recommendation);

  const common = document.createElement('section');
  common.className = 'similarity-common-block';
  const commonTitle = document.createElement('h4');
  commonTitle.textContent = `Tabelas em comum (${formatNumber(pair.commonTables.length)})`;
  common.append(commonTitle, createChipList(pair.commonTables, 'Nenhuma tabela em comum.'));
  body.appendChild(common);

  if (!pair.onlyA.length && !pair.onlyB.length) {
    const fullOverlap = document.createElement('p');
    fullOverlap.className = 'similarity-full-overlap';
    fullOverlap.textContent = 'Os dois relatórios usam o mesmo conjunto de tabelas mapeadas nesta base.';
    body.appendChild(fullOverlap);
    return body;
  }

  const exclusive = document.createElement('div');
  exclusive.className = 'similarity-exclusive';
  exclusive.appendChild(createCompactTableGroup(`Só no Relatório A (${pair.reportA.name})`, pair.onlyA));
  exclusive.appendChild(createCompactTableGroup(`Só no Relatório B (${pair.reportB.name})`, pair.onlyB));
  body.appendChild(exclusive);
  return body;
}

function createChipList(tables, emptyText) {
  if (!tables.length) {
    const empty = document.createElement('p');
    empty.className = 'empty-state';
    empty.textContent = emptyText;
    return empty;
  }

  const list = document.createElement('div');
  list.className = 'table-chip-list';
  for (const table of tables) {
    const chip = document.createElement('span');
    chip.className = 'table-chip';
    chip.textContent = table;
    list.appendChild(chip);
  }
  return list;
}

function createCompactTableGroup(title, tables) {
  const group = document.createElement('div');
  group.className = 'similarity-exclusive-group';
  const heading = document.createElement('h4');
  heading.textContent = title;
  group.appendChild(heading);
  group.appendChild(createChipList(tables, 'Nenhuma tabela exclusiva.'));
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
  setText('#pageInfo', `Página ${formatNumber(state.page)} de ${formatNumber(pageCount)}`);

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

decorateSegmentAwareLinks();
applySegmentCopy();
loadSegmentsCatalog();
window.addEventListener('scroll', updateActiveLink, { passive: true });
window.addEventListener('resize', updateActiveLink);
updateActiveLink();

loadManifest().then(() => {
  if (bodyPage === 'analytical') {
    renderAnalyticalPage();
  }

  if (bodyPage === 'detailed') {
    renderDetailedPage();
  }
});
