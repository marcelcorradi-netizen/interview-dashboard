const state = {
  stakeholders: [],
  insights: [],
  stakeholderMap: new Map(),
  stakeholderDetails: new Map(),
  filters: {
    team: 'all',
    insightType: 'all',
  },
  currentView: 'overview',
  selections: {
    category: null,
    expectationCategory: null,
  },
  detailFilters: {
    type: 'all',
    category: 'all',
    team: 'all',
    stakeholder: 'all',
  },
  selectedDetailId: null,
};

const elements = {
  teamFilter: document.getElementById('team-filter'),
  navItems: Array.from(document.querySelectorAll('.nav-item')),
  pages: Array.from(document.querySelectorAll('[data-view]')),
  metrics: {
    stakeholders: document.getElementById('metric-stakeholders'),
    insights: document.getElementById('metric-insights'),
    tenure: document.getElementById('metric-tenure'),
  },
  charts: {
    teams: document.getElementById('chart-teams'),
    insights: document.getElementById('chart-insights'),
  },
  highlight: {
    type: document.getElementById('top-insight-type'),
    percent: document.getElementById('top-insight-percent'),
  },
  dimension: {
    categoryChart: document.getElementById('dimension-category-chart'),
    categoryCaption: document.getElementById('dimension-category-caption'),
    clearFilter: document.getElementById('dimension-clear-filter'),
    teamDonut: document.getElementById('dimension-team-donut'),
    teamLegend: document.getElementById('dimension-team-legend'),
    mentionsBody: document.getElementById('dimension-mentions-body'),
    tagsBody: document.getElementById('dimension-tags-body'),
    categoryTitle: document.getElementById('dimension-category-title'),
    teamTitle: document.getElementById('dimension-team-title'),
    tagsTitle: document.getElementById('dimension-tags-title'),
  },
  typeFilter: document.getElementById('type-filter'),
  typeFilterGroup: document.getElementById('type-filter-group'),
  teamFilterGroup: document.getElementById('team-filter-group'),
  expectations: {
    metrics: {
      expectationStakeholders: document.getElementById('expectation-stakeholders'),
      expectationStakeholdersCaption: document.getElementById('expectation-stakeholders-caption'),
      expectationCount: document.getElementById('expectation-count'),
      engagementStakeholders: document.getElementById('engagement-stakeholders'),
      engagementStakeholdersCaption: document.getElementById('engagement-stakeholders-caption'),
      engagementCount: document.getElementById('engagement-count'),
    },
    categoryChart: document.getElementById('expectation-category-chart'),
    categoryCaption: document.getElementById('expectation-category-caption'),
    clearFilter: document.getElementById('expectation-clear-filter'),
    highlightsBody: document.getElementById('expectation-highlights-body'),
    tagsBody: document.getElementById('expectation-tags-body'),
    teamDonut: document.getElementById('engagement-team-donut'),
    teamLegend: document.getElementById('engagement-team-legend'),
    rankingList: document.getElementById('engagement-ranking-list'),
  },
  detail: {
    typeFilter: document.getElementById('detail-type-filter'),
    categoryFilter: document.getElementById('detail-category-filter'),
    teamFilter: document.getElementById('detail-team-filter'),
    stakeholderFilter: document.getElementById('detail-stakeholder-filter'),
    tableBody: document.getElementById('detail-table-body'),
    profileContent: document.getElementById('detail-profile-content'),
    timelineChart: document.getElementById('detail-timeline-chart'),
    tagsBody: document.getElementById('detail-tags-body'),
  },
  final: {
    totalStakeholders: document.getElementById('final-total-stakeholders'),
    totalInsights: document.getElementById('final-total-insights'),
    expectationRate: document.getElementById('final-expectation-rate'),
    expectationCaption: document.getElementById('final-expectation-caption'),
    engagementRate: document.getElementById('final-engagement-rate'),
    engagementCaption: document.getElementById('final-engagement-caption'),
    themeGrid: document.getElementById('final-theme-grid'),
    actionsBody: document.getElementById('final-actions-body'),
    influenceBody: document.getElementById('final-influence-body'),
  },
};

let navOverflowController = null;
let tooltipController = null;

setupNavigation();
navOverflowController = setupNavOverflow();
if (navOverflowController) {
  window.addEventListener('load', () => navOverflowController.update());
}
tooltipController = initTooltip();
updateNavigation();
updateViewVisibility();

if (elements.teamFilter) {
  elements.teamFilter.addEventListener('change', (event) => {
    state.filters.team = event.target.value;
    state.selections.category = null;
    state.selections.expectationCategory = null;
    state.detailFilters.team = 'all';
    state.detailFilters.stakeholder = 'all';
    state.selectedDetailId = null;
    renderDashboard();
  });
}

if (elements.typeFilter) {
  elements.typeFilter.addEventListener('change', (event) => {
    state.filters.insightType = event.target.value;
    state.selections.category = null;
    renderDashboard();
  });
}

if (elements.dimension.clearFilter) {
  elements.dimension.clearFilter.addEventListener('click', () => {
    state.selections.category = null;
    renderDashboard();
  });
}

if (elements.expectations.clearFilter) {
  elements.expectations.clearFilter.addEventListener('click', () => {
    state.selections.expectationCategory = null;
    renderDashboard();
  });
}

const detailFiltersElements = elements.detail;
if (detailFiltersElements.typeFilter) {
  detailFiltersElements.typeFilter.addEventListener('change', (event) => {
    state.detailFilters.type = event.target.value;
    state.selectedDetailId = null;
    renderDashboard();
  });
}

if (detailFiltersElements.categoryFilter) {
  detailFiltersElements.categoryFilter.addEventListener('change', (event) => {
    state.detailFilters.category = event.target.value;
    state.selectedDetailId = null;
    renderDashboard();
  });
}

if (detailFiltersElements.teamFilter) {
  detailFiltersElements.teamFilter.addEventListener('change', (event) => {
    state.detailFilters.team = event.target.value;
    state.selectedDetailId = null;
    renderDashboard();
  });
}

if (detailFiltersElements.stakeholderFilter) {
  detailFiltersElements.stakeholderFilter.addEventListener('change', (event) => {
    state.detailFilters.stakeholder = event.target.value;
    state.selectedDetailId = null;
    renderDashboard();
  });
}

loadData();

function setupNavigation() {
  elements.navItems.forEach((item) => {
    const targetView = item.dataset.page;
    const sectionExists = elements.pages.some((section) => section.dataset.view === targetView);
    if (!sectionExists) {
      return;
    }
    item.addEventListener('click', () => setView(targetView));
  });
}

function setView(view) {
  if (state.currentView === view) {
    return;
  }

  if (view === 'detail') {
    if (state.filters.team !== 'all') {
      state.filters.team = 'all';
      if (elements.teamFilter) {
        elements.teamFilter.value = 'all';
      }
    }
    state.detailFilters = {
      type: 'all',
      category: 'all',
      team: 'all',
      stakeholder: 'all',
    };
    state.selectedDetailId = null;
    populateDetailFilters();
  }

  state.currentView = view;
  updateNavigation();
  updateViewVisibility();
  renderDashboard();
  if (navOverflowController) {
    navOverflowController.close();
    navOverflowController.update();
  }
}

async function loadData() {
  try {
    const [stakeholdersResponse, insightsResponse] = await Promise.all([
      fetch('data/stakeholders.json'),
      fetch('data/insights.json'),
    ]);

    if (!stakeholdersResponse.ok || !insightsResponse.ok) {
      throw new Error('Não foi possível carregar os dados.');
    }

    state.stakeholders = await stakeholdersResponse.json();
    const rawInsights = await insightsResponse.json();
    state.insights = rawInsights.map((insight, index) => ({ ...insight, _id: index }));
    state.stakeholderMap = new Map(
      state.stakeholders.map((stakeholder) => [stakeholder.id, stakeholder.time || 'Sem time']),
    );
    state.stakeholderDetails = new Map(
      state.stakeholders.map((stakeholder) => [stakeholder.id, stakeholder]),
    );

    populateTeamFilter();
    populateTypeFilter();
    populateDetailFilters();
    renderDashboard();
  } catch (error) {
    console.error(error);
    showErrorMessage(error.message);
  }
}

function populateTeamFilter() {
  if (!elements.teamFilter) return;
  const teams = Array.from(
    new Set(state.stakeholders.map((stakeholder) => stakeholder.time).filter(Boolean)),
  ).sort((a, b) => a.localeCompare(b, 'pt-BR'));

  if (!teams.includes(state.filters.team)) {
    state.filters.team = 'all';
  }

  const select = elements.teamFilter;
  select.innerHTML = '';
  const allOption = document.createElement('option');
  allOption.value = 'all';
  allOption.textContent = 'Todos';
  select.appendChild(allOption);

  teams.forEach((team) => {
    const option = document.createElement('option');
    option.value = team;
    option.textContent = team;
    select.appendChild(option);
  });

  select.value = state.filters.team;
}

function populateTypeFilter() {
  if (!elements.typeFilter) return;
  const types = Array.from(new Set(state.insights.map((insight) => insight.tipo).filter(Boolean))).sort(
    (a, b) => a.localeCompare(b, 'pt-BR'),
  );

  elements.typeFilter.innerHTML = '';
  const allOption = document.createElement('option');
  allOption.value = 'all';
  allOption.textContent = 'Todos';
  elements.typeFilter.appendChild(allOption);

  types.forEach((type) => {
    const option = document.createElement('option');
    option.value = type;
    option.textContent = type;
    elements.typeFilter.appendChild(option);
  });

  if (state.filters.insightType === 'all') {
    const preferred = types.find((type) => type.toLowerCase() === 'dor');
    state.filters.insightType = preferred || 'all';
  }

  if (!types.includes(state.filters.insightType) && state.filters.insightType !== 'all') {
    state.filters.insightType = 'all';
  }

  elements.typeFilter.value = state.filters.insightType;
}

function renderDashboard() {
  const { stakeholders, insights } = getFilteredData();
  renderOverview(stakeholders, insights);
  renderDimension(stakeholders, insights);
  renderExpectations(stakeholders, insights);
  renderDetail(insights);
  renderFinalView();
}

function getFilteredData() {
  const { team } = state.filters;
  const filteredStakeholders =
    team === 'all'
      ? [...state.stakeholders]
      : state.stakeholders.filter((stakeholder) => stakeholder.time === team);

  const stakeholderIds = new Set(filteredStakeholders.map((stakeholder) => stakeholder.id));
  const filteredInsights =
    stakeholderIds.size === 0
      ? []
      : state.insights.filter((insight) => stakeholderIds.has(insight.stakeholder_id));

  return { stakeholders: filteredStakeholders, insights: filteredInsights };
}

function renderOverview(stakeholders, insights) {
  renderMetrics(stakeholders, insights);
  renderTeamChart(stakeholders);
  renderInsightTypeChart(insights);
  renderHighlight(insights);
}

function renderMetrics(stakeholders, insights) {
  elements.metrics.stakeholders.textContent = stakeholders.length.toString();
  elements.metrics.insights.textContent = insights.length.toString();

  const averageTenure =
    stakeholders.length === 0
      ? 0
      : stakeholders.reduce((sum, stakeholder) => sum + (stakeholder.tempo_meses || 0), 0) /
        stakeholders.length;

  const roundedTenure = Math.round(averageTenure || 0);
  elements.metrics.tenure.textContent = roundedTenure.toString();
}

function renderTeamChart(stakeholders) {
  const container = elements.charts.teams;
  if (!container) return;

  const counts = stakeholders.reduce((acc, stakeholder) => {
    const key = stakeholder.time || 'Sem time';
    acc.set(key, (acc.get(key) || 0) + 1);
    return acc;
  }, new Map());

  const data = Array.from(counts.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([team, count]) => ({
      key: team,
      label: team,
      value: count,
    }));

  renderHorizontalBarChart(container, data, {
    emptyMessage: 'Nenhum stakeholder para exibir.',
    minPercent: 12,
    getOpacity: (_item, index) => Math.max(0.5, 0.95 - index * 0.05),
  });
}

function renderInsightTypeChart(insights) {
  const container = elements.charts.insights;
  if (!container) return;

  const counts = insights.reduce((acc, insight) => {
    const key = insight.tipo || 'Não classificado';
    acc.set(key, (acc.get(key) || 0) + 1);
    return acc;
  }, new Map());

  const data = Array.from(counts.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([type, count]) => ({
      key: type,
      label: type,
      value: count,
    }));

  renderHorizontalBarChart(container, data, {
    emptyMessage: 'Nenhum insight disponível.',
    minPercent: 8,
    getOpacity: (_item, index) => Math.max(0.55, 0.92 - index * 0.04),
  });
}

function renderHighlight(insights) {
  const total = insights.length;

  if (total === 0) {
    elements.highlight.type.textContent = 'Sem dados';
    elements.highlight.percent.textContent = '0% dos insights';
    return;
  }

  const counts = insights.reduce((acc, insight) => {
    const key = insight.tipo || 'Não classificado';
    acc.set(key, (acc.get(key) || 0) + 1);
    return acc;
  }, new Map());

  const [topType, topCount] = Array.from(counts.entries()).sort((a, b) => b[1] - a[1])[0];
  const percent = Math.round((topCount / total) * 100);

  elements.highlight.type.textContent = topType;
  elements.highlight.percent.textContent = `${percent}% dos insights`;
}

function renderDimension(_stakeholders, insights) {
  const categoryContainer = elements.dimension.categoryChart;
  if (!categoryContainer) return;

  const { team, insightType } = state.filters;

  const typeFilteredInsights =
    insightType === 'all'
      ? state.insights
      : state.insights.filter((insight) => (insight.tipo || '') === insightType);

  const eligibleStakeholders =
    team === 'all'
      ? state.stakeholders
      : state.stakeholders.filter((stakeholder) => stakeholder.time === team);

  const allowedIds = new Set(eligibleStakeholders.map((stakeholder) => stakeholder.id));

  const filteredInsights = typeFilteredInsights.filter((insight) =>
    allowedIds.has(insight.stakeholder_id),
  );

  updateDimensionTitles(insightType);

  renderDimensionCategoryChart(filteredInsights);

  const selectedCategory = state.selections.category;
  const categoryInsights =
    selectedCategory === null
      ? filteredInsights
      : filteredInsights.filter((insight) => (insight.categoria || '') === selectedCategory);

  renderDimensionTeamDistribution(categoryInsights);
  renderDimensionMentionsTable(categoryInsights);
  renderDimensionTagsTable(categoryInsights);
  updateDimensionFilterUI(selectedCategory, filteredInsights.length, categoryInsights.length, insightType);
}

function renderExpectations(_stakeholders, insights) {
  const expectations = elements.expectations;
  if (!expectations.categoryChart) return;

  const totalStakeholders = _stakeholders.length;
  const expectationTypes = ['Impacto_Esperado', 'Necessidade', 'Sugestao'];
  const expectationInsights = insights.filter((insight) =>
    expectationTypes.includes(insight.tipo || ''),
  );
  const engagementInsights = insights.filter(
    (insight) => (insight.tipo || '') === 'Engajamento',
  );

  const expectationStakeholders = new Set(
    expectationInsights.map((insight) => insight.stakeholder_id).filter(Boolean),
  );
  const engagementStakeholders = new Set(
    engagementInsights.map((insight) => insight.stakeholder_id).filter(Boolean),
  );

  if (expectations.metrics.expectationStakeholders) {
    expectations.metrics.expectationStakeholders.textContent = formatPercentage(
      expectationStakeholders.size,
      totalStakeholders,
    );
  }
  if (expectations.metrics.expectationStakeholdersCaption) {
    expectations.metrics.expectationStakeholdersCaption.textContent =
      totalStakeholders > 0
        ? `${expectationStakeholders.size} de ${totalStakeholders} stakeholders`
        : 'Nenhum stakeholder disponível';
  }
  if (expectations.metrics.expectationCount) {
    expectations.metrics.expectationCount.textContent = expectationInsights.length.toString();
  }
  if (expectations.metrics.engagementStakeholders) {
    expectations.metrics.engagementStakeholders.textContent = formatPercentage(
      engagementStakeholders.size,
      totalStakeholders,
    );
  }
  if (expectations.metrics.engagementStakeholdersCaption) {
    expectations.metrics.engagementStakeholdersCaption.textContent =
      totalStakeholders > 0
        ? `${engagementStakeholders.size} de ${totalStakeholders} stakeholders`
        : 'Nenhum stakeholder disponível';
  }
  if (expectations.metrics.engagementCount) {
    expectations.metrics.engagementCount.textContent = engagementInsights.length.toString();
  }

  renderExpectationsCategoryChart(expectationInsights);

  const selectedCategory = state.selections.expectationCategory;
  const filteredExpectations =
    selectedCategory === null
      ? expectationInsights
      : expectationInsights.filter(
          (insight) => (insight.categoria || 'Não classificado') === selectedCategory,
        );

  renderExpectationHighlights(filteredExpectations);
  renderExpectationTags(filteredExpectations);
  renderEngagementTeamDonut(engagementInsights);
  renderEngagementRanking(engagementInsights);
  updateExpectationCaption(selectedCategory, expectationInsights.length, filteredExpectations.length);
}

function renderExpectationsCategoryChart(insights) {
  const container = elements.expectations.categoryChart;
  container.innerHTML = '';

  const counts = insights.reduce((acc, insight) => {
    const category = insight.categoria || 'Não classificado';
    acc.set(category, (acc.get(category) || 0) + 1);
    return acc;
  }, new Map());

  if (state.selections.expectationCategory && !counts.has(state.selections.expectationCategory)) {
    state.selections.expectationCategory = null;
  }

  const data = Array.from(counts.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([category, count]) => ({
      key: category,
      label: category,
      value: count,
    }));

  const selectedKey = state.selections.expectationCategory;

  renderHorizontalBarChart(container, data, {
    emptyMessage: 'Nenhuma expectativa encontrada.',
    interactive: true,
    selectedKey,
    onSelect: handleExpectationCategorySelect,
    minPercent: 12,
    getOpacity: (_item, index, isSelected) => {
      if (selectedKey && !isSelected) {
        return 0.35;
      }
      return Math.max(0.6, 0.95 - index * 0.05);
    },
  });
}

function handleExpectationCategorySelect(category) {
  state.selections.expectationCategory =
    state.selections.expectationCategory === category ? null : category;
  renderDashboard();
}

function renderExpectationHighlights(insights) {
  const tbody = elements.expectations.highlightsBody;
  if (!tbody) return;
  tbody.innerHTML = '';

  if (!insights.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 2;
    cell.textContent = 'Nenhuma expectativa disponível para os filtros atuais.';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  const sorted = insights
    .slice()
    .sort(
      (a, b) =>
        new Date(b.data_entrevista || '1970-01-01').getTime() -
        new Date(a.data_entrevista || '1970-01-01').getTime(),
    );

  sorted.forEach((insight) => {
    const row = document.createElement('tr');
    const descriptionCell = document.createElement('td');
    setCellText(descriptionCell, insight.descricao || '(Sem descrição registrada)');

    const stakeholderCell = document.createElement('td');
    const details = state.stakeholderDetails.get(insight.stakeholder_id);
    const name = details?.nome || `Stakeholder ${insight.stakeholder_id || '-'}`;
    const team = details?.time || 'Sem time';
    const role = details?.cargo;
    const metaParts = [team];
    if (role) metaParts.push(role);
    const metaLabel = metaParts.join(' · ') || 'Sem informações adicionais';
    stakeholderCell.innerHTML = `<strong>${name}</strong><br><small>${metaLabel}</small>`;
    applyTooltip(stakeholderCell, `${name} · ${metaLabel}`);

    row.append(descriptionCell, stakeholderCell);
    tbody.appendChild(row);
  });
}

function renderExpectationTags(insights) {
  const tbody = elements.expectations.tagsBody;
  if (!tbody) return;
  tbody.innerHTML = '';

  if (!insights.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 2;
    cell.textContent = 'Nenhuma tag relacionada.';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  const counts = insights.reduce((acc, insight) => {
    (insight.tags || []).forEach((tag) => {
      const normalized = tag?.trim();
      if (!normalized) return;
      acc.set(normalized, (acc.get(normalized) || 0) + 1);
    });
    return acc;
  }, new Map());

  const entries = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);

  entries.forEach(([tag, count]) => {
    const row = document.createElement('tr');
    const nameCell = document.createElement('td');
    setCellText(nameCell, tag);
    const valueCell = document.createElement('td');
    setCellText(valueCell, count.toString());
    row.append(nameCell, valueCell);
    tbody.appendChild(row);
  });
}

function renderEngagementTeamDonut(insights) {
  const container = elements.expectations.teamDonut;
  const legend = elements.expectations.teamLegend;
  if (!container || !legend) return;
  container.innerHTML = '';
  legend.innerHTML = '';

  if (!insights.length) {
    container.appendChild(createPlaceholder('Nenhum registro de engajamento.'));
    return;
  }

  const counts = insights.reduce((acc, insight) => {
    const team = state.stakeholderMap.get(insight.stakeholder_id) || 'Sem time';
    acc.set(team, (acc.get(team) || 0) + 1);
    return acc;
  }, new Map());

  const entries = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);
  const total = entries.reduce((sum, [, count]) => sum + count, 0);

  const palette = ['#0f6bd6', '#1a88f2', '#60b4ff', '#2e4063', '#8a9ab8', '#60c5f1'];
  let currentAngle = 0;

  const segments = entries.map(([team, count], index) => {
    const color = palette[index % palette.length];
    const start = currentAngle;
    const angle = (count / total) * 360;
    currentAngle += angle;
    return {
      label: team,
      count,
      color,
      start,
      end: currentAngle,
    };
  });

  const gradient = segments
    .map(({ color, start, end }) => `${color} ${start}deg ${end}deg`)
    .join(', ');

  const donut = document.createElement('div');
  donut.className = 'donut-chart';
  donut.style.background = `conic-gradient(${gradient})`;

  const center = document.createElement('div');
  center.className = 'donut-center';
  center.innerHTML = `<strong>${total}</strong><span>menções</span>`;

  donut.appendChild(center);
  container.appendChild(donut);

  segments.forEach(({ label, count, color }) => {
    const li = document.createElement('li');
    const percent = Math.round((count / total) * 100);
    li.innerHTML = `
      <span class="legend-color" style="background:${color}"></span>
      <span class="legend-label">${label}</span>
      <span class="legend-value">${count} (${percent}%)</span>
    `;
    legend.appendChild(li);
  });
}

function renderEngagementRanking(insights) {
  const list = elements.expectations.rankingList;
  if (!list) return;
  list.innerHTML = '';

  if (!insights.length) {
    const item = document.createElement('li');
    item.className = 'chart-placeholder';
    item.textContent = 'Nenhuma menção de engajamento para exibir.';
    list.appendChild(item);
    return;
  }

  const counts = insights.reduce((acc, insight) => {
    const key = insight.stakeholder_id;
    if (!key) return acc;
    const record = acc.get(key) || { count: 0, tags: new Map() };
    record.count += 1;
    (insight.tags || []).forEach((tag) => {
      const label = tag?.trim();
      if (!label) return;
      record.tags.set(label, (record.tags.get(label) || 0) + 1);
    });
    acc.set(key, record);
    return acc;
  }, new Map());

  const ranking = Array.from(counts.entries()).sort((a, b) => b[1].count - a[1].count);

  const maxCount = ranking[0]?.[1].count || 1;

  ranking.forEach(([stakeholderId, data], index) => {
    const { count, tags } = data;
    const li = document.createElement('li');
    const rankIndex = document.createElement('span');
    rankIndex.className = 'rank-index';
    rankIndex.textContent = String(index + 1);

    const content = document.createElement('div');
    content.className = 'rank-content';

    const details = state.stakeholderDetails.get(stakeholderId);
    const name = details?.nome || `Stakeholder ${stakeholderId}`;
    const team = details?.time || 'Sem time';
    const role = details?.cargo ? ` · ${details.cargo}` : '';
    const topTag =
      Array.from(tags.entries()).sort((a, b) => b[1] - a[1])[0]?.[0] || 'Sem tag destaque';

    const nameEl = document.createElement('div');
    nameEl.className = 'rank-name';
    nameEl.textContent = name;

    const metaEl = document.createElement('div');
    metaEl.className = 'rank-meta';
    metaEl.textContent = `${count} menção(ões) · ${team}${role} · destaque: ${topTag}`;

    const bar = document.createElement('div');
    bar.className = 'rank-bar';
    const fill = document.createElement('span');
    fill.style.width = `${Math.max(10, (count / maxCount) * 100)}%`;
    bar.appendChild(fill);

    content.append(nameEl, metaEl, bar);
    li.append(rankIndex, content);
    list.appendChild(li);
  });
}

function updateExpectationCaption(selectedCategory, total, filtered) {
  const caption = elements.expectations.categoryCaption;
  const clearButton = elements.expectations.clearFilter;
  if (!caption || !clearButton) return;

  if (!total) {
    caption.textContent = 'Nenhuma expectativa encontrada para os filtros atuais.';
    clearButton.classList.add('is-hidden');
    return;
  }

  if (selectedCategory) {
    caption.textContent = `${filtered} expectativa(s) encontradas em “${selectedCategory}”.`;
    clearButton.classList.remove('is-hidden');
  } else {
    caption.textContent = 'Clique para filtrar os detalhes por categoria.';
    clearButton.classList.add('is-hidden');
  }
}

function populateDetailFilters() {
  const detail = elements.detail;
  if (!detail.typeFilter) return;

  const types = Array.from(
    new Set(state.insights.map((insight) => insight.tipo || 'Não classificado')),
  ).sort((a, b) => a.localeCompare(b, 'pt-BR'));
  setSelectOptions(detail.typeFilter, types, 'Todos', (value) => value, (value) => value);
  if (detail.typeFilter) {
    detail.typeFilter.value = types.includes(state.detailFilters.type) ? state.detailFilters.type : 'all';
    state.detailFilters.type = detail.typeFilter.value;
  }

  const categories = Array.from(
    new Set(state.insights.map((insight) => insight.categoria || 'Não classificado')),
  ).sort((a, b) => a.localeCompare(b, 'pt-BR'));
  setSelectOptions(
    detail.categoryFilter,
    categories,
    'Todas',
    (value) => value,
    (value) => value,
  );
  if (detail.categoryFilter) {
    detail.categoryFilter.value = categories.includes(state.detailFilters.category)
      ? state.detailFilters.category
      : 'all';
    state.detailFilters.category = detail.categoryFilter.value;
  }

  const teams = Array.from(new Set(state.stakeholders.map((stakeholder) => stakeholder.time || 'Sem time'))).sort(
    (a, b) => a.localeCompare(b, 'pt-BR'),
  );
  setSelectOptions(detail.teamFilter, teams, 'Todos', (value) => value, (value) => value);
  if (detail.teamFilter) {
    detail.teamFilter.value = teams.includes(state.detailFilters.team) ? state.detailFilters.team : 'all';
    state.detailFilters.team = detail.teamFilter.value;
  }

  const stakeholders = state.stakeholders.slice().sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
  setSelectOptions(
    detail.stakeholderFilter,
    stakeholders,
    'Todos',
    (stakeholder) => String(stakeholder.id),
    (stakeholder) => stakeholder.nome,
  );
  if (detail.stakeholderFilter) {
    const ids = stakeholders.map((s) => String(s.id));
    detail.stakeholderFilter.value = ids.includes(state.detailFilters.stakeholder)
      ? state.detailFilters.stakeholder
      : 'all';
    state.detailFilters.stakeholder = detail.stakeholderFilter.value;
  }
}

function renderDetail(insights) {
  const detailElements = elements.detail;
  if (!detailElements.tableBody) return;

  const filtered = applyDetailFilters(insights);
  if (filtered.length > 0) {
    const exists = filtered.some((insight) => insight._id === state.selectedDetailId);
    if (!exists) {
      state.selectedDetailId = filtered[0]._id;
    }
  } else {
    state.selectedDetailId = null;
  }

  renderDetailTable(filtered);
  renderDetailProfile(filtered);
  renderDetailTimeline(filtered);
  renderDetailTags(filtered);
}

function applyDetailFilters(insights) {
  const { type, category, team, stakeholder } = state.detailFilters;
  return insights.filter((insight) => {
    const insightType = insight.tipo || 'Não classificado';
    const insightCategory = insight.categoria || 'Não classificado';
    const insightTeam = state.stakeholderMap.get(insight.stakeholder_id) || 'Sem time';
    const stakeholderId = insight.stakeholder_id ? String(insight.stakeholder_id) : null;

    if (type !== 'all' && insightType !== type) return false;
    if (category !== 'all' && insightCategory !== category) return false;
    if (team !== 'all' && insightTeam !== team) return false;
    if (stakeholder !== 'all' && stakeholderId !== stakeholder) return false;
    return true;
  });
}

function renderDetailTable(insights) {
  const tbody = elements.detail.tableBody;
  if (!tbody) return;
  tbody.innerHTML = '';

  if (!insights.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 6;
    cell.textContent = 'Nenhum insight encontrado com os filtros atuais.';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  insights.forEach((insight) => {
    const row = document.createElement('tr');
    row.dataset.id = String(insight._id);
    if (insight._id === state.selectedDetailId) {
      row.classList.add('selected');
    }

    const dateCell = document.createElement('td');
    setCellText(dateCell, formatDate(insight.data_entrevista));

    const typeCell = document.createElement('td');
    setCellText(typeCell, insight.tipo || 'Não classificado');

    const categoryCell = document.createElement('td');
    setCellText(categoryCell, insight.categoria || 'Não classificado');

    const stakeholderCell = document.createElement('td');
    const details = state.stakeholderDetails.get(insight.stakeholder_id);
    if (details) {
      const metaParts = [details.time || 'Sem time'];
      if (details.cargo) metaParts.push(details.cargo);
      const metaLabel = metaParts.join(' · ');
      stakeholderCell.innerHTML = `<strong>${details.nome}</strong><br><small>${metaLabel}</small>`;
      applyTooltip(stakeholderCell, `${details.nome} · ${metaLabel}`);
    } else {
      setCellText(stakeholderCell, `Stakeholder ${insight.stakeholder_id || '-'}`);
    }

    const descriptionCell = document.createElement('td');
    setCellText(descriptionCell, insight.descricao || '(Sem descrição registrada)');

    const tagsCell = document.createElement('td');
    const tagsText = (insight.tags || []).join(', ') || '—';
    setCellText(tagsCell, tagsText);

    row.append(dateCell, typeCell, categoryCell, stakeholderCell, descriptionCell, tagsCell);
    row.addEventListener('click', () => {
      state.selectedDetailId = insight._id;
      Array.from(tbody.querySelectorAll('tr')).forEach((tr) => {
        tr.classList.toggle('selected', tr.dataset.id === String(insight._id));
      });
      renderDetailProfile(insights);
    });

    tbody.appendChild(row);
  });
}

function renderDetailProfile(filteredInsights) {
  const container = elements.detail.profileContent;
  if (!container) return;

  if (state.selectedDetailId === null || state.selectedDetailId === undefined) {
    container.innerHTML = '<p class="chart-placeholder">Selecione um insight para ver detalhes do stakeholder.</p>';
    return;
  }

  const selected =
    filteredInsights.find((insight) => insight._id === state.selectedDetailId) ||
    state.insights.find((insight) => insight._id === state.selectedDetailId);

  if (!selected) {
    container.innerHTML = '<p class="chart-placeholder">Seleção indisponível para os filtros atuais.</p>';
    return;
  }

  const details = state.stakeholderDetails.get(selected.stakeholder_id);
  if (!details) {
    container.innerHTML = '<p class="chart-placeholder">Dados do stakeholder não encontrados.</p>';
    return;
  }

  const stakeholderInsights = state.insights.filter(
    (insight) => insight.stakeholder_id === selected.stakeholder_id,
  );
  const countsByType = stakeholderInsights.reduce((acc, insight) => {
    const type = insight.tipo || 'Não classificado';
    acc.set(type, (acc.get(type) || 0) + 1);
    return acc;
  }, new Map());

  const topTags = stakeholderInsights.reduce((acc, insight) => {
    (insight.tags || []).forEach((tag) => {
      const label = tag?.trim();
      if (!label) return;
      acc.set(label, (acc.get(label) || 0) + 1);
    });
    return acc;
  }, new Map());

  const topTagEntries = Array.from(topTags.entries()).sort((a, b) => b[1] - a[1]).slice(0, 4);

  const topTypes = Array.from(countsByType.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([type, count]) => `${type} (${count})`);

  container.innerHTML = `
    <div class="profile-header">
      <div class="profile-avatar">${getInitials(details.nome)}</div>
      <div class="profile-info">
        <h3>${details.nome}</h3>
        <span>${details.cargo || 'Cargo não informado'}</span>
        <span>${details.time || 'Sem time'} · ${details.area || 'Sem área'}</span>
      </div>
    </div>
    <div class="profile-stats">
      <div class="stat">
        <strong>${stakeholderInsights.length}</strong>
        <span>Total de insights</span>
      </div>
      <div class="stat">
        <strong>${countsByType.get('Dor') || 0}</strong>
        <span>Dores reportadas</span>
      </div>
      <div class="stat">
        <strong>${countsByType.get('Engajamento') || 0}</strong>
        <span>Menções de engajamento</span>
      </div>
      <div class="stat">
        <strong>${Math.round(details.tempo_meses || 0)}</strong>
        <span>Meses na empresa</span>
      </div>
    </div>
    <ul class="profile-summary">
      <li><strong>Categoria atual:</strong> ${selected.categoria || 'Não classificado'}</li>
      <li><strong>Última menção:</strong> ${formatDate(selected.data_entrevista)}</li>
      <li><strong>Principais tipos:</strong> ${
        topTypes.length ? topTypes.join(', ') : 'Sem histórico registrado'
      }</li>
      <li><strong>Tags frequentes:</strong> ${
        topTagEntries.length ? topTagEntries.map(([tag]) => tag).join(', ') : 'Sem tags registradas'
      }</li>
    </ul>
  `;
}

function renderDetailTimeline(insights) {
  const container = elements.detail.timelineChart;
  if (!container) return;
  container.innerHTML = '';

  if (!insights.length) {
    container.appendChild(createPlaceholder('Nenhum insight para compor a timeline.'));
    return;
  }

  const counts = insights.reduce((acc, insight) => {
    const key = insight.data_entrevista ? insight.data_entrevista.slice(0, 7) : 'Sem data';
    acc.set(key, (acc.get(key) || 0) + 1);
    return acc;
  }, new Map());

  const entries = Array.from(counts.entries()).sort((a, b) => a[0].localeCompare(b[0]));
  const maxValue = Math.max(...entries.map(([, value]) => value));

  entries.forEach(([month, value]) => {
    const bar = document.createElement('div');
    bar.className = 'timeline-bar';

    const label = document.createElement('span');
    label.className = 'timeline-label';
    label.textContent = month === 'Sem data' ? 'Sem data' : formatMonthLabel(month);

    const track = document.createElement('div');
    track.className = 'timeline-track';

    const fill = document.createElement('div');
    fill.className = 'timeline-fill';
    fill.style.width = `${Math.max(8, (value / maxValue) * 100)}%`;

    const valueEl = document.createElement('span');
    valueEl.className = 'timeline-value';
    valueEl.textContent = value.toString();

    track.appendChild(fill);
    bar.append(label, track, valueEl);
    container.appendChild(bar);
  });
}

function renderDetailTags(insights) {
  const tbody = elements.detail.tagsBody;
  if (!tbody) return;
  tbody.innerHTML = '';

  if (!insights.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 2;
    cell.textContent = 'Nenhuma tag encontrada para os filtros atuais.';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  const counts = insights.reduce((acc, insight) => {
    (insight.tags || []).forEach((tag) => {
      const label = tag?.trim();
      if (!label) return;
      acc.set(label, (acc.get(label) || 0) + 1);
    });
    return acc;
  }, new Map());

  const entries = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);

  if (!entries.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 2;
    cell.textContent = 'Nenhuma tag encontrada para os filtros atuais.';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  entries.forEach(([tag, count]) => {
    const row = document.createElement('tr');
    const tagCell = document.createElement('td');
    setCellText(tagCell, tag);
    const valueCell = document.createElement('td');
    setCellText(valueCell, count.toString());
    row.append(tagCell, valueCell);
    tbody.appendChild(row);
  });
}

function renderFinalView() {
  const final = elements.final;
  if (!final.totalStakeholders) return;

  const totalStakeholders = state.stakeholders.length;
  const totalInsights = state.insights.length;
  const expectationTypes = ['Impacto_Esperado', 'Necessidade', 'Sugestao'];

  const expectationStakeholders = new Set(
    state.insights
      .filter((insight) => expectationTypes.includes(insight.tipo || ''))
      .map((insight) => insight.stakeholder_id)
      .filter(Boolean),
  );

  const engagementStakeholders = new Set(
    state.insights
      .filter((insight) => (insight.tipo || '') === 'Engajamento')
      .map((insight) => insight.stakeholder_id)
      .filter(Boolean),
  );

  final.totalStakeholders.textContent = totalStakeholders.toString();
  final.totalInsights.textContent = totalInsights.toString();
  final.expectationRate.textContent = formatPercentage(expectationStakeholders.size, totalStakeholders);
  final.expectationCaption.textContent = `${expectationStakeholders.size} de ${totalStakeholders} stakeholders`;
  final.engagementRate.textContent = formatPercentage(engagementStakeholders.size, totalStakeholders);
  final.engagementCaption.textContent = `${engagementStakeholders.size} de ${totalStakeholders} stakeholders`;

  renderFinalThemes();
  renderFinalActions();
  renderFinalInfluence();
}

function renderFinalThemes() {
  const container = elements.final.themeGrid;
  if (!container) return;
  container.innerHTML = '';

  const expectationTypes = ['Impacto_Esperado', 'Necessidade', 'Sugestao'];
  const normalizeLabel = (text = '') =>
    text
      .split(/[_\s]+/g)
      .filter(Boolean)
      .map((segment) => segment.charAt(0).toUpperCase() + segment.slice(1).toLowerCase())
      .join(' ');

  const getGroupLabel = (type) => {
    if (expectationTypes.includes(type || '')) return 'Expectativas';
    const normalized = (type || '').trim();
    if (!normalized) return 'Não classificado';
    if (normalized.toLowerCase() === 'dor') return 'Dores';
    if (normalized === 'Engajamento') return 'Engajamento';
    return normalizeLabel(normalized);
  };

  const groupedInsights = state.insights.reduce((acc, insight) => {
    const label = getGroupLabel(insight.tipo);
    const bucket = acc.get(label) || [];
    bucket.push(insight);
    acc.set(label, bucket);
    return acc;
  }, new Map());

  const blocks = Array.from(groupedInsights.entries())
    .map(([label, insights]) => {
      if (!insights.length) {
        return {
          label,
          description: 'Nenhuma menção registrada.',
          total: 0,
        };
      }
      const counts = insights.reduce((acc, insight) => {
        const category = insight.categoria || 'Não classificado';
        acc.set(category, (acc.get(category) || 0) + 1);
        return acc;
      }, new Map());
      const [primary, secondary] = Array.from(counts.entries())
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'))
        .slice(0, 2);

      const primaryText = primary
        ? `${primary[0]} (${primary[1]} menções)`
        : 'Nenhuma categoria destacada';
      const secondaryText = secondary ? `${secondary[0]} (${secondary[1]} menções)` : null;

      return {
        label,
        primaryText,
        secondaryText,
        total: insights.length,
      };
    })
    .sort((a, b) => b.total - a.total || a.label.localeCompare(b.label, 'pt-BR'));

  if (!blocks.length) {
    container.appendChild(createPlaceholder('Nenhum dado disponível para a síntese.'));
    return;
  }

  blocks.forEach((block) => {
    const themeBlock = document.createElement('div');
    themeBlock.className = 'theme-block';
    themeBlock.innerHTML = `
      <strong>${block.label}</strong>
      <span>${block.primaryText || block.description}</span>
      ${block.secondaryText ? `<span>${block.secondaryText}</span>` : ''}
    `;
    container.appendChild(themeBlock);
  });
}

function renderFinalActions() {
  const tbody = elements.final.actionsBody;
  if (!tbody) return;
  tbody.innerHTML = '';

  const counts = state.insights.reduce((acc, insight) => {
    (insight.tags || []).forEach((tag) => {
      const label = tag?.trim();
      if (!label) return;
      acc.set(label, (acc.get(label) || 0) + 1);
    });
    return acc;
  }, new Map());

  const entries = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]).slice(0, 6);

  if (!entries.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 2;
    cell.textContent = 'Nenhuma recomendação automática disponível.';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  entries.forEach(([tag, count]) => {
    const row = document.createElement('tr');
    const actionCell = document.createElement('td');
    setCellText(actionCell, `Priorizar iniciativas relacionadas a ${tag}.`);
    const baseCell = document.createElement('td');
    setCellText(baseCell, `Tag mencionada ${count} vez(es).`);
    row.append(actionCell, baseCell);
    tbody.appendChild(row);
  });
}

function renderFinalInfluence() {
  const tbody = elements.final.influenceBody;
  if (!tbody) return;
  tbody.innerHTML = '';

  const expectationTypes = ['Impacto_Esperado', 'Necessidade', 'Sugestao'];
  const influenceMap = new Map();

  state.insights.forEach((insight) => {
    const id = insight.stakeholder_id;
    if (!id) return;
    const entry =
      influenceMap.get(id) || {
        dores: 0,
        expectativas: 0,
        engajamento: 0,
        tags: new Map(),
      };
    const tipo = insight.tipo || '';
    if (tipo.toLowerCase() === 'dor') entry.dores += 1;
    if (expectationTypes.includes(tipo)) entry.expectativas += 1;
    if (tipo === 'Engajamento') entry.engajamento += 1;
    (insight.tags || []).forEach((tag) => {
      const label = tag?.trim();
      if (!label) return;
      entry.tags.set(label, (entry.tags.get(label) || 0) + 1);
    });
    influenceMap.set(id, entry);
  });

  const ranking = Array.from(influenceMap.entries())
    .map(([stakeholderId, data]) => {
      const total = data.dores + data.expectativas + data.engajamento;
      return { stakeholderId, data, total };
    })
    .filter((entry) => entry.total > 0)
    .sort((a, b) => b.total - a.total);

  if (!ranking.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 3;
    cell.textContent = 'Nenhum stakeholder com dados suficientes para análise de influência.';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  ranking.forEach(({ stakeholderId, data }) => {
    const details = state.stakeholderDetails.get(stakeholderId);
    const row = document.createElement('tr');

    const nameCell = document.createElement('td');
    if (details) {
      const metaParts = [details.time || 'Sem time', details.cargo || 'Sem cargo'];
      const metaLabel = metaParts.join(' · ');
      nameCell.innerHTML = `<strong>${details.nome}</strong><br><small>${metaLabel}</small>`;
      applyTooltip(nameCell, `${details.nome} · ${metaLabel}`);
    } else {
      setCellText(nameCell, `Stakeholder ${stakeholderId}`);
    }

    const insightsCell = document.createElement('td');
    const insightsSummary = `Dores: ${data.dores} · Expectativas: ${data.expectativas} · Engajamento: ${data.engajamento}`;
    setCellText(insightsCell, insightsSummary);

    const tags = Array.from(data.tags.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([tag]) => tag)
      .join(', ');
    const observationsCell = document.createElement('td');
    const observationText = tags ? `Tags recorrentes: ${tags}` : 'Sem tags destacadas.';
    setCellText(observationsCell, observationText);

    row.append(nameCell, insightsCell, observationsCell);
    tbody.appendChild(row);
  });
}

function renderDimensionCategoryChart(insights) {
  const container = elements.dimension.categoryChart;
  container.innerHTML = '';

  const counts = insights.reduce((acc, insight) => {
    const key = insight.categoria || 'Não classificado';
    acc.set(key, (acc.get(key) || 0) + 1);
    return acc;
  }, new Map());

  const data = Array.from(counts.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([category, count]) => ({
      key: category,
      label: category,
      value: count,
    }));

  if (state.selections.category && !counts.has(state.selections.category)) {
    state.selections.category = null;
  }

  const selectedKey = state.selections.category;

  renderHorizontalBarChart(container, data, {
    emptyMessage: 'Nenhuma dor encontrada.',
    interactive: true,
    selectedKey,
    onSelect: handleDimensionCategorySelect,
    minPercent: 10,
    getOpacity: (_item, index, isSelected) => {
      if (selectedKey && !isSelected) {
        return 0.35;
      }
      return Math.max(0.6, 0.95 - index * 0.05);
    },
  });
}

function handleDimensionCategorySelect(category) {
  state.selections.category =
    state.selections.category === category ? null : category;
  renderDashboard();
}

function renderDimensionTeamDistribution(insights) {
  const container = elements.dimension.teamDonut;
  const legend = elements.dimension.teamLegend;
  container.innerHTML = '';
  legend.innerHTML = '';

  if (!insights.length) {
    container.appendChild(createPlaceholder('Nenhum registro para exibir.'));
    return;
  }

  const counts = insights.reduce((acc, insight) => {
    const team = state.stakeholderMap.get(insight.stakeholder_id) || 'Sem time';
    acc.set(team, (acc.get(team) || 0) + 1);
    return acc;
  }, new Map());

  const entries = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);
  const total = entries.reduce((sum, [, count]) => sum + count, 0);

  const palette = ['#0f6bd6', '#1a88f2', '#60b4ff', '#2e4063', '#8a9ab8', '#60c5f1'];
  let currentAngle = 0;

  const segments = entries.map(([team, count], index) => {
    const color = palette[index % palette.length];
    const start = currentAngle;
    const angle = (count / total) * 360;
    currentAngle += angle;
    return {
      label: team,
      count,
      color,
      start,
      end: currentAngle,
    };
  });

  const gradient = segments
    .map(({ color, start, end }) => `${color} ${start}deg ${end}deg`)
    .join(', ');

  const donut = document.createElement('div');
  donut.className = 'donut-chart';
  donut.style.background = `conic-gradient(${gradient})`;

  const center = document.createElement('div');
  center.className = 'donut-center';
  center.innerHTML = `<strong>${total}</strong><span>registros</span>`;

  donut.appendChild(center);
  container.appendChild(donut);

  segments.forEach(({ label, count, color }) => {
    const li = document.createElement('li');
    const percent = Math.round((count / total) * 100);
    li.innerHTML = `
      <span class="legend-color" style="background:${color}"></span>
      <span class="legend-label">${label}</span>
      <span class="legend-value">${count} (${percent}%)</span>
    `;
    legend.appendChild(li);
  });
}

function renderDimensionMentionsTable(insights) {
  const tbody = elements.dimension.mentionsBody;
  tbody.innerHTML = '';

  if (!insights.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 1;
    cell.textContent = 'Nenhuma menção disponível.';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  insights.forEach((insight) => {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    setCellText(cell, insight.descricao || '(Sem descrição registrada)');
    row.appendChild(cell);
    tbody.appendChild(row);
  });
}

function renderDimensionTagsTable(insights) {
  const tbody = elements.dimension.tagsBody;
  tbody.innerHTML = '';

  if (!insights.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 2;
    cell.textContent = 'Nenhuma tag disponível.';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  const counts = insights.reduce((acc, insight) => {
    (insight.tags || []).forEach((tag) => {
      const normalized = tag?.trim();
      if (!normalized) return;
      acc.set(normalized, (acc.get(normalized) || 0) + 1);
    });
    return acc;
  }, new Map());

  const entries = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);

  entries.forEach(([tag, count]) => {
    const row = document.createElement('tr');
    const nameCell = document.createElement('td');
    setCellText(nameCell, tag);
    const valueCell = document.createElement('td');
    setCellText(valueCell, count.toString());
    row.append(nameCell, valueCell);
    tbody.appendChild(row);
  });
}

function updateDimensionFilterUI(selectedCategory, total, filtered, insightType) {
  const caption = elements.dimension.categoryCaption;
  const clearButton = elements.dimension.clearFilter;
  if (!caption || !clearButton) return;

  const typeText = insightType === 'all' ? 'todos os tipos de insight' : `o tipo ${insightType}`;

  if (!total) {
    caption.textContent = `Nenhum registro relacionado a ${typeText} com os filtros atuais.`;
    clearButton.classList.add('is-hidden');
    return;
  }

  if (selectedCategory) {
    caption.textContent = `${filtered} menções encontradas em “${selectedCategory}” relacionadas a ${typeText}.`;
    clearButton.classList.remove('is-hidden');
  } else {
    caption.textContent = `Selecione uma categoria para explorar detalhes relacionados a ${typeText}.`;
    clearButton.classList.add('is-hidden');
  }
}

function updateDimensionTitles(insightType) {
  const { categoryTitle, teamTitle, tagsTitle } = elements.dimension;

  if (categoryTitle) {
    categoryTitle.textContent =
      insightType === 'all'
        ? 'Categorias por tipo de insight'
        : `Categorias do tipo ${insightType}`;
  }

  if (teamTitle) {
    teamTitle.textContent =
      insightType === 'all'
        ? 'Distribuição por time'
        : `Distribuição do tipo ${insightType} por time`;
  }

  if (tagsTitle) {
    tagsTitle.textContent =
      insightType === 'all'
        ? 'Tags relacionadas'
        : `Tags relacionadas ao tipo ${insightType}`;
  }
}

function renderHorizontalBarChart(container, items, options = {}) {
  const {
    emptyMessage = 'Nenhum dado disponível.',
    interactive = false,
    selectedKey = null,
    onSelect,
    minPercent = 8,
    getOpacity,
    valueFormatter,
  } = options;

  container.innerHTML = '';

  if (!items.length) {
    container.appendChild(createPlaceholder(emptyMessage));
    return;
  }

  const maxValue = Math.max(...items.map((item) => item.value));

  items.forEach((item, index) => {
    const bar = document.createElement('div');
    bar.className = 'bar';
    applyTooltip(bar, `${item.label} — ${valueFormatter ? valueFormatter(item.value) : item.value}`);

    if (interactive) {
      bar.classList.add('interactive');
      if (selectedKey && item.key === selectedKey) {
        bar.classList.add('selected');
      }
      bar.addEventListener('click', () => {
        if (typeof onSelect === 'function') {
          onSelect(item.key);
        }
      });
    }

    const label = document.createElement('span');
    label.className = 'bar-label';
    label.textContent = item.label;
    applyTooltip(label, item.label);

    const track = document.createElement('div');
    track.className = 'bar-track';

    const fill = document.createElement('div');
    fill.className = 'bar-fill';
    const width = maxValue === 0 ? 0 : Math.max(minPercent, (item.value / maxValue) * 100);
    fill.style.width = `${width}%`;
    if (item.color) {
      fill.style.background = item.color;
    }

    if (typeof getOpacity === 'function') {
      const opacity = getOpacity(item, index, item.key === selectedKey);
      fill.style.opacity = opacity;
    } else if (interactive && selectedKey && item.key === selectedKey) {
      fill.style.opacity = 1;
    }

    const value = document.createElement('span');
    value.className = 'bar-value';
    value.textContent = valueFormatter ? valueFormatter(item.value) : item.value.toString();
    applyTooltip(value, value.textContent);

    track.appendChild(fill);
    bar.append(label, track, value);
    container.appendChild(bar);
  });
}

function setupNavOverflow() {
  const navList = document.querySelector('.sidebar-nav ul');
  const moreItem = navList?.querySelector('.nav-more');
  const moreMenu = moreItem?.querySelector('.nav-more-menu');

  if (!navList || !moreItem || !moreMenu) {
    return null;
  }

  let overflowItems = [];
  let resizeTimeout;

  const closeMenu = () => {
    moreItem.classList.remove('open');
    moreItem.setAttribute('aria-expanded', 'false');
    moreMenu.hidden = true;
  };

  const restoreItems = () => {
    overflowItems.forEach((item) => {
      item.classList.remove('nav-item-overflow');
      navList.insertBefore(item, moreItem);
    });
    overflowItems = [];
  };

  const updateOverflow = () => {
    restoreItems();
    closeMenu();
    moreMenu.innerHTML = '';

    const isCompact = window.matchMedia('(max-width: 1200px)').matches;
    if (!isCompact) {
      moreItem.classList.add('is-hidden');
      return;
    }

    moreItem.classList.remove('is-hidden');

    const navItems = Array.from(navList.querySelectorAll('.nav-item:not(.nav-more)'));
    const navWidth = navList.clientWidth;
    const styles = window.getComputedStyle(navList);
    const gapValue = parseFloat(styles.columnGap || styles.gap || '0') || 0;

    const totalItemsWidth =
      navItems.reduce((sum, item) => sum + item.offsetWidth, 0) +
      Math.max(0, navItems.length - 1) * gapValue;

    if (totalItemsWidth <= navWidth) {
      moreItem.classList.add('is-hidden');
      return;
    }

    let remainingWidth = Math.max(0, navWidth - moreItem.offsetWidth);
    let keptCount = 0;

    navItems.forEach((item) => {
      const gapBefore = keptCount > 0 ? gapValue : 0;
      const requiredWidth = gapBefore + item.offsetWidth;
      const requiredWithMore = requiredWidth + gapValue;
      if (requiredWithMore <= remainingWidth) {
        remainingWidth -= requiredWidth;
        keptCount += 1;
      } else {
        overflowItems.push(item);
      }
    });

    if (overflowItems.length) {
      overflowItems.forEach((item) => {
        item.classList.add('nav-item-overflow');
        moreMenu.appendChild(item);
      });
      moreItem.classList.remove('is-hidden');
      moreMenu.hidden = true;
    } else {
      moreItem.classList.add('is-hidden');
    }
  };

  const toggleMenu = () => {
    if (!overflowItems.length) return;
    const isOpen = moreItem.classList.toggle('open');
    moreItem.setAttribute('aria-expanded', String(isOpen));
    moreMenu.hidden = !isOpen;
  };

  moreItem.addEventListener('click', (event) => {
    if (event.target.closest('.nav-more-menu')) {
      closeMenu();
      return;
    }
    toggleMenu();
  });

  moreItem.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      toggleMenu();
    } else if (event.key === 'Escape') {
      closeMenu();
    }
  });

  document.addEventListener('click', (event) => {
    if (!moreItem.contains(event.target)) {
      closeMenu();
    }
  });

  const handleResize = () => {
    clearTimeout(resizeTimeout);
    resizeTimeout = window.setTimeout(updateOverflow, 150);
  };

  window.addEventListener('resize', handleResize);
  updateOverflow();

  return {
    update: updateOverflow,
    close: closeMenu,
  };
}

function initTooltip() {
  const tooltip = document.createElement('div');
  tooltip.className = 'tooltip-bubble';
  tooltip.hidden = true;
  document.body.appendChild(tooltip);

  let activeTarget = null;

  const showTooltip = (target, x, y) => {
    const text = target.dataset.tooltip;
    if (!text) return;
    tooltip.textContent = text;
    tooltip.style.top = `${Math.max(12, y - 28)}px`;
    tooltip.style.left = `${x + 16}px`;
    tooltip.hidden = false;
    tooltip.classList.add('visible');
  };

  const hideTooltip = () => {
    tooltip.hidden = true;
    tooltip.classList.remove('visible');
    activeTarget = null;
  };

  document.addEventListener('pointerover', (event) => {
    const target = event.target.closest('[data-tooltip]');
    if (!target || !target.dataset.tooltip) {
      return;
    }
    activeTarget = target;
    showTooltip(target, event.clientX, event.clientY);
  });

  document.addEventListener('pointermove', (event) => {
    if (!activeTarget) return;
    if (!event.target.closest('[data-tooltip]')) {
      hideTooltip();
      return;
    }
    showTooltip(activeTarget, event.clientX, event.clientY);
  });

  document.addEventListener('pointerout', (event) => {
    if (!activeTarget) return;
    const nextTarget = event.relatedTarget?.closest('[data-tooltip]');
    if (nextTarget === activeTarget) return;
    hideTooltip();
  });

  document.addEventListener('scroll', hideTooltip, true);

  return {
    hide: hideTooltip,
  };
}

function updateNavigation() {
  elements.navItems.forEach((item) => {
    if (item.dataset.page === state.currentView) {
      item.classList.add('active');
    } else {
      item.classList.remove('active');
    }
  });
}

function updateViewVisibility() {
  elements.pages.forEach((section) => {
    if (section.dataset.view === state.currentView) {
      section.classList.remove('is-hidden');
    } else {
      section.classList.add('is-hidden');
    }
  });
  updateFilterVisibility();
}

function showErrorMessage(message) {
  const cardsGrid = document.querySelector('.cards-grid');
  if (cardsGrid) {
    cardsGrid.innerHTML = '';
    const errorCard = document.createElement('article');
    errorCard.className = 'card';
    errorCard.textContent = message;
    cardsGrid.appendChild(errorCard);
  }
}

function createPlaceholder(text) {
  const placeholder = document.createElement('p');
  placeholder.className = 'chart-placeholder';
  placeholder.textContent = text;
  return placeholder;
}

function updateFilterVisibility() {
  if (elements.typeFilterGroup) {
    if (state.currentView === 'dimension') {
      elements.typeFilterGroup.classList.remove('is-hidden');
    } else {
      elements.typeFilterGroup.classList.add('is-hidden');
    }
  }

  if (elements.teamFilterGroup) {
    if (state.currentView === 'detail' || state.currentView === 'final') {
      elements.teamFilterGroup.classList.add('is-hidden');
    } else {
      elements.teamFilterGroup.classList.remove('is-hidden');
    }
  }
}

function setSelectOptions(select, values, allLabel, getValue, getLabel) {
  if (!select) return;
  const currentValue = select.value;
  select.innerHTML = '';
  const allOption = document.createElement('option');
  allOption.value = 'all';
  allOption.textContent = allLabel;
  select.appendChild(allOption);
  values.forEach((item) => {
    const option = document.createElement('option');
    option.value = getValue(item);
    option.textContent = getLabel(item);
    select.appendChild(option);
  });
  if (currentValue && Array.from(select.options).some((option) => option.value === currentValue)) {
    select.value = currentValue;
  } else {
    select.value = 'all';
  }
}

function applyTooltip(element, text) {
  if (!element) return;
  const value = text ?? '';
  const normalized = typeof value === 'string' ? value.trim() : String(value).trim();
  if (normalized && normalized !== '—') {
    element.removeAttribute('title');
    element.dataset.tooltip = normalized;
  } else {
    element.removeAttribute('title');
    delete element.dataset.tooltip;
  }
}

function setCellText(cell, text, fallback = '—') {
  const value = text ?? fallback;
  cell.textContent = value;
  applyTooltip(cell, value);
}

function formatDate(isoString) {
  if (!isoString) return 'Sem data';
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) return 'Sem data';
  return new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' }).format(date);
}

function formatMonthLabel(monthKey) {
  const [year, month] = monthKey.split('-');
  const date = new Date(Number(year), Number(month) - 1, 1);
  return new Intl.DateTimeFormat('pt-BR', { month: 'short', year: 'numeric' }).format(date);
}

function formatPercentage(part, total) {
  if (!total || total === 0) return '0%';
  return `${Math.round((part / total) * 100)}%`;
}

function getInitials(name = '') {
  return name
    .split(' ')
    .filter(Boolean)
    .slice(0, 2)
    .map((segment) => segment[0]?.toUpperCase() || '')
    .join('');
}
