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
  themePreference: 'system',
  appliedTheme: 'light',
  selections: {
    category: null,
    expectationCategory: null,
    tag: null,
    expectationTag: null,
  },
};

const elements = {
  brandLogo: document.getElementById('brand-logo'),
  themeToggle: document.getElementById('theme-toggle'),
  teamFilter: document.getElementById('team-filter'),
  navItems: Array.from(document.querySelectorAll('.nav-item')),
  pages: Array.from(document.querySelectorAll('[data-view]')),
  topBar: {
    title: document.getElementById('page-title'),
  },
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
    tagsCaption: document.getElementById('dimension-tags-caption'),
    clearTag: document.getElementById('dimension-tag-clear'),
  },
  typeFilter: document.getElementById('type-filter'),
  typeFilterGroup: document.getElementById('type-filter-group'),
  teamFilterGroup: document.getElementById('team-filter-group'),
  exportButton: document.getElementById('export-button'),
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
    tagsCaption: document.getElementById('expectation-tags-caption'),
    clearTag: document.getElementById('expectation-tag-clear'),
    teamDonut: document.getElementById('engagement-team-donut'),
    teamLegend: document.getElementById('engagement-team-legend'),
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

const THEME_STORAGE_KEY = 'stakeholder-dashboard-theme';
const themeMediaQuery =
  typeof window !== 'undefined' && typeof window.matchMedia === 'function'
    ? window.matchMedia('(prefers-color-scheme: dark)')
    : null;

const VIEW_TITLES = {
  overview: 'Visão geral',
  dimension: 'Dimensão',
  expectations: 'Expectativas & Engajamento',
  final: 'Panorama final',
};

let navOverflowController = null;
let tooltipController = null;

initTheme();
setupNavigation();
navOverflowController = setupNavOverflow();
if (navOverflowController) {
  window.addEventListener('load', () => navOverflowController.update());
}
tooltipController = initTooltip();
updateNavigation();
updateViewVisibility();
updatePageHeader();

if (elements.teamFilter) {
  elements.teamFilter.addEventListener('change', (event) => {
    state.filters.team = event.target.value;
    state.selections.category = null;
    state.selections.expectationCategory = null;
    state.selections.tag = null;
    state.selections.expectationTag = null;
    renderDashboard();
  });
}

if (elements.typeFilter) {
  elements.typeFilter.addEventListener('change', (event) => {
    state.filters.insightType = event.target.value;
    state.selections.category = null;
    state.selections.tag = null;
    state.selections.expectationTag = null;
    renderDashboard();
  });
}

if (elements.exportButton) {
  elements.exportButton.addEventListener('click', handleExportSnapshot);
}

if (elements.dimension.clearFilter) {
  elements.dimension.clearFilter.addEventListener('click', () => {
    state.selections.category = null;
    state.selections.tag = null;
    renderDashboard();
  });
}

if (elements.dimension.clearTag) {
  elements.dimension.clearTag.addEventListener('click', () => {
    state.selections.tag = null;
    renderDashboard();
  });
}

if (elements.expectations.clearFilter) {
  elements.expectations.clearFilter.addEventListener('click', () => {
    state.selections.expectationCategory = null;
    state.selections.expectationTag = null;
    renderDashboard();
  });
}

if (elements.expectations.clearTag) {
  elements.expectations.clearTag.addEventListener('click', () => {
    state.selections.expectationTag = null;
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

  state.currentView = view;
  updateNavigation();
  updateViewVisibility();
  updatePageHeader();
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

  const selectedCategory = state.selections.category;
  const categoryBaseInsights =
    selectedCategory === null
      ? filteredInsights
      : filteredInsights.filter((insight) => (insight.categoria || '') === selectedCategory);

  let selectedTag = state.selections.tag;
  if (
    selectedTag &&
    !categoryBaseInsights.some((insight) =>
      (insight.tags || []).some((tag) => tag && tag.trim() === selectedTag),
    )
  ) {
    selectedTag = null;
    state.selections.tag = null;
  }

  const categoryChartInsights =
    selectedTag === null
      ? filteredInsights
      : filteredInsights.filter((insight) =>
          (insight.tags || []).some((tag) => tag && tag.trim() === selectedTag),
        );

  renderDimensionCategoryChart(categoryChartInsights);

  const tagInsights =
    selectedTag === null
      ? categoryBaseInsights
      : categoryBaseInsights.filter((insight) =>
          (insight.tags || []).some((tag) => tag && tag.trim() === selectedTag),
        );

  renderDimensionTeamDistribution(tagInsights);
  renderDimensionMentionsTable(tagInsights);
  renderDimensionTagsTable(categoryBaseInsights, selectedTag);
  updateDimensionFilterUI(
    selectedCategory,
    selectedTag,
    filteredInsights.length,
    categoryBaseInsights.length,
    tagInsights.length,
    insightType,
  );
  updateDimensionTagUI(selectedTag, categoryBaseInsights.length, tagInsights.length);
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

  const selectedCategory = state.selections.expectationCategory;
  const categoryExpectations =
    selectedCategory === null
      ? expectationInsights
      : expectationInsights.filter(
          (insight) => (insight.categoria || 'Não classificado') === selectedCategory,
        );

  let selectedTag = state.selections.expectationTag;
  if (
    selectedTag &&
    !categoryExpectations.some((insight) =>
      (insight.tags || []).some((tag) => tag && tag.trim() === selectedTag),
    )
  ) {
    selectedTag = null;
    state.selections.expectationTag = null;
  }

  const expectationsForChart =
    selectedTag === null
      ? expectationInsights
      : expectationInsights.filter((insight) =>
          (insight.tags || []).some((tag) => tag && tag.trim() === selectedTag),
        );

  renderExpectationsCategoryChart(expectationsForChart);

  const tagFilteredExpectations =
    selectedTag === null
      ? categoryExpectations
      : categoryExpectations.filter((insight) =>
          (insight.tags || []).some((tag) => tag && tag.trim() === selectedTag),
        );

  renderExpectationHighlights(tagFilteredExpectations);
  renderExpectationTags(categoryExpectations, selectedTag);
  const engagementForDonut =
    selectedTag === null
      ? engagementInsights
      : engagementInsights.filter((insight) =>
          (insight.tags || []).some((tag) => tag && tag.trim() === selectedTag),
        );
  renderEngagementTeamDonut(engagementForDonut);
  updateExpectationCaption(
    selectedCategory,
    selectedTag,
    expectationInsights.length,
    categoryExpectations.length,
    tagFilteredExpectations.length,
  );
  updateExpectationTagUI(selectedTag, categoryExpectations.length, tagFilteredExpectations.length);
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
  state.selections.expectationTag = null;
  renderDashboard();
}

function renderExpectationHighlights(insights) {
  const tbody = elements.expectations.highlightsBody;
  if (!tbody) return;
  tbody.innerHTML = '';

  if (!insights.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 1;
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
    row.append(descriptionCell);
    tbody.appendChild(row);
  });
}

function renderExpectationTags(insights, selectedTag) {
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
  const activeTag = selectedTag && counts.has(selectedTag) ? selectedTag : null;

  entries.forEach(([tag, count]) => {
    const row = document.createElement('tr');
    const nameCell = document.createElement('td');
    setCellText(nameCell, tag);
    const valueCell = document.createElement('td');
    setCellText(valueCell, count.toString());
    row.append(nameCell, valueCell);
    row.addEventListener('click', () => handleExpectationTagSelect(tag));
    if (activeTag === tag) {
      row.classList.add('selected');
    }
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

function updateExpectationCaption(
  selectedCategory,
  selectedTag,
  totalAvailable,
  categoryCount,
  tagCount,
) {
  const caption = elements.expectations.categoryCaption;
  const clearButton = elements.expectations.clearFilter;
  if (!caption || !clearButton) return;

  if (!totalAvailable) {
    caption.textContent = 'Nenhuma expectativa encontrada para os filtros atuais.';
    clearButton.classList.add('is-hidden');
    return;
  }

  const hasCategory = Boolean(selectedCategory);
  const hasTag = Boolean(selectedTag);

  if (hasCategory && hasTag) {
    caption.textContent = `${tagCount} expectativa(s) em “${selectedCategory}” contendo a tag “${selectedTag}”.`;
    clearButton.classList.remove('is-hidden');
    return;
  }

  if (hasCategory) {
    caption.textContent = `${categoryCount} expectativa(s) encontradas em “${selectedCategory}”.`;
    clearButton.classList.remove('is-hidden');
    return;
  }

  if (hasTag) {
    caption.textContent = `${tagCount} expectativa(s) contendo a tag “${selectedTag}”.`;
    clearButton.classList.remove('is-hidden');
    return;
  }

  caption.textContent = 'Clique para filtrar os detalhes por categoria.';
  clearButton.classList.add('is-hidden');
}

function updateExpectationTagUI(selectedTag, totalAvailable, filteredCount) {
  const caption = elements.expectations.tagsCaption;
  const clearButton = elements.expectations.clearTag;
  if (!caption || !clearButton) return;

  if (!totalAvailable) {
    caption.textContent = 'Nenhuma tag disponível para os filtros atuais.';
    clearButton.classList.add('is-hidden');
    return;
  }

  if (selectedTag) {
    caption.textContent = `${filteredCount} expectativa(s) contendo a tag “${selectedTag}”.`;
    clearButton.classList.remove('is-hidden');
  } else {
    caption.textContent = 'Clique em uma tag para filtrar as expectativas.';
    clearButton.classList.add('is-hidden');
  }
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
  const keyTagCounts = new Map();

  state.insights.forEach((insight) => {
    const tipo = insight.tipo || '';
    const normalized = tipo.toLowerCase();
    const isKeyInsight =
      normalized === 'dor' || tipo === 'Engajamento' || expectationTypes.includes(tipo);
    if (!isKeyInsight) return;

    (insight.tags || []).forEach((rawTag) => {
      const tag = rawTag?.trim();
      if (!tag) return;
      const entry =
        keyTagCounts.get(tag) || {
          dores: 0,
          expectativas: 0,
          engajamento: 0,
        };

      if (normalized === 'dor') entry.dores += 1;
      if (expectationTypes.includes(tipo)) entry.expectativas += 1;
      if (tipo === 'Engajamento') entry.engajamento += 1;

      keyTagCounts.set(tag, entry);
    });
  });

  const entries = Array.from(keyTagCounts.entries())
    .map(([tag, totals]) => {
      const total = totals.dores + totals.expectativas + totals.engajamento;
      if (total === 0) return null;
      return {
        tag,
        totals,
        total,
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.total - a.total || a.tag.localeCompare(b.tag, 'pt-BR'));

  if (!entries.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 2;
    cell.textContent = 'Nenhuma tag com insights chave disponível.';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  entries.forEach(({ tag, totals, total }) => {
    const row = document.createElement('tr');

    const tagCell = document.createElement('td');
    tagCell.innerHTML = `<strong>${tag}</strong>`;
    applyTooltip(tagCell, tag);

    const summaryCell = document.createElement('td');
    const parts = [];
    if (totals.dores) parts.push(`Dores: ${totals.dores}`);
    if (totals.expectativas) parts.push(`Expectativas: ${totals.expectativas}`);
    if (totals.engajamento) parts.push(`Engajamento: ${totals.engajamento}`);
    const detail = parts.length ? ` (${parts.join(' · ')})` : '';
    const summaryText = `${total} insight${total === 1 ? '' : 's'} chave${detail}`;
    setCellText(summaryCell, summaryText);

    row.append(tagCell, summaryCell);
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
  state.selections.tag = null;
  renderDashboard();
}

function handleDimensionTagSelect(tag) {
  state.selections.tag = state.selections.tag === tag ? null : tag;
  renderDashboard();
}

function handleExpectationTagSelect(tag) {
  state.selections.expectationTag =
    state.selections.expectationTag === tag ? null : tag;
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

function renderDimensionTagsTable(insights, selectedTag) {
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
  const activeTag = selectedTag && counts.has(selectedTag) ? selectedTag : null;

  entries.forEach(([tag, count]) => {
    const row = document.createElement('tr');
    const nameCell = document.createElement('td');
    setCellText(nameCell, tag);
    const valueCell = document.createElement('td');
    setCellText(valueCell, count.toString());
    row.append(nameCell, valueCell);
    row.addEventListener('click', () => handleDimensionTagSelect(tag));
    if (activeTag === tag) {
      row.classList.add('selected');
    }
    tbody.appendChild(row);
  });
}

function updateDimensionFilterUI(
  selectedCategory,
  selectedTag,
  totalAvailable,
  categoryCount,
  tagCount,
  insightType,
) {
  const caption = elements.dimension.categoryCaption;
  const clearButton = elements.dimension.clearFilter;
  if (!caption || !clearButton) return;

  const typeText = insightType === 'all' ? 'todos os tipos de insight' : `o tipo ${insightType}`;

  if (!totalAvailable) {
    caption.textContent = `Nenhum registro relacionado a ${typeText} com os filtros atuais.`;
    clearButton.classList.add('is-hidden');
    return;
  }

  const hasCategory = Boolean(selectedCategory);
  const hasTag = Boolean(selectedTag);

  if (hasCategory && hasTag) {
    caption.textContent = `${categoryCount} menções em “${selectedCategory}”, sendo ${tagCount} com a tag “${selectedTag}”.`;
  } else if (hasCategory) {
    caption.textContent = `${categoryCount} menções encontradas em “${selectedCategory}” relacionadas a ${typeText}.`;
  } else if (hasTag) {
    caption.textContent = `${tagCount} menções contendo a tag “${selectedTag}” relacionadas a ${typeText}.`;
  } else {
    caption.textContent = `Selecione uma categoria ou tag para explorar detalhes relacionados a ${typeText}.`;
  }

  if (hasCategory || hasTag) {
    clearButton.classList.remove('is-hidden');
  } else {
    clearButton.classList.add('is-hidden');
  }
}

function updateDimensionTagUI(selectedTag, totalInsights, filteredInsights) {
  const caption = elements.dimension.tagsCaption;
  const clearButton = elements.dimension.clearTag;
  if (!caption || !clearButton) return;

  if (!totalInsights) {
    caption.textContent = 'Nenhuma tag disponível para os filtros atuais.';
    clearButton.classList.add('is-hidden');
    return;
  }

  if (selectedTag) {
    caption.textContent = `${filteredInsights} menção(ões) contendo a tag “${selectedTag}”.`;
    clearButton.classList.remove('is-hidden');
  } else {
    caption.textContent = 'Clique em uma tag para filtrar as menções.';
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
    if (state.currentView === 'final') {
      elements.teamFilterGroup.classList.add('is-hidden');
    } else {
      elements.teamFilterGroup.classList.remove('is-hidden');
    }
  }
}

function getStoredThemePreference() {
  try {
    const storedPreference = window.localStorage?.getItem(THEME_STORAGE_KEY);
    if (storedPreference === 'light' || storedPreference === 'dark' || storedPreference === 'system') {
      return storedPreference;
    }
  } catch (error) {
    // Ignore storage errors and fall back to system
  }
  return 'system';
}

function storeThemePreference(preference) {
  try {
    if (preference === 'system') {
      window.localStorage?.removeItem(THEME_STORAGE_KEY);
    } else {
      window.localStorage?.setItem(THEME_STORAGE_KEY, preference);
    }
  } catch (error) {
    // Ignore storage errors
  }
}

function resolveTheme(preference) {
  if (preference === 'dark' || preference === 'light') {
    return preference;
  }
  if (themeMediaQuery && themeMediaQuery.matches) {
    return 'dark';
  }
  return 'light';
}

function updateBrandLogo(resolvedTheme) {
  const { brandLogo } = elements;
  if (!brandLogo) {
    return;
  }
  const targetSrc = resolvedTheme === 'dark' ? brandLogo.dataset.logoDark : brandLogo.dataset.logoLight;
  if (targetSrc && brandLogo.getAttribute('src') !== targetSrc) {
    brandLogo.setAttribute('src', targetSrc);
  }
}

function updateThemeToggle() {
  const { themeToggle } = elements;
  if (!themeToggle) {
    return;
  }
  const isDark = state.appliedTheme === 'dark';
  themeToggle.setAttribute('aria-pressed', String(isDark));
  const toggleTitle = isDark ? 'Tema atual: escuro. Clique para mudar para claro.' : 'Tema atual: claro. Clique para mudar para escuro.';
  themeToggle.setAttribute('title', toggleTitle);
  themeToggle.setAttribute('aria-label', toggleTitle);
}

function applyTheme(preference) {
  const resolvedTheme = resolveTheme(preference);
  state.appliedTheme = resolvedTheme;
  if (resolvedTheme === 'dark') {
    document.body.dataset.theme = 'dark';
  } else {
    delete document.body.dataset.theme;
  }
  document.documentElement.style.colorScheme = resolvedTheme === 'dark' ? 'dark' : 'light';
  updateBrandLogo(resolvedTheme);
  updateThemeToggle();
}

function setThemePreference(preference) {
  const allowed = new Set(['system', 'light', 'dark']);
  const nextPreference = allowed.has(preference) ? preference : 'system';
  state.themePreference = nextPreference;
  storeThemePreference(nextPreference);
  applyTheme(nextPreference);
}

function handleSystemThemeChange() {
  if (state.themePreference === 'system') {
    applyTheme('system');
  }
}

function initTheme() {
  state.themePreference = getStoredThemePreference();
  applyTheme(state.themePreference);
  if (elements.themeToggle) {
    elements.themeToggle.addEventListener('click', () => {
      const targetTheme = state.appliedTheme === 'dark' ? 'light' : 'dark';
      if (themeMediaQuery) {
        const systemPrefersDark = themeMediaQuery.matches;
        const systemTheme = systemPrefersDark ? 'dark' : 'light';
        if (targetTheme === systemTheme) {
          setThemePreference('system');
          return;
        }
      }
      setThemePreference(targetTheme);
    });
  }
  if (themeMediaQuery) {
    if (typeof themeMediaQuery.addEventListener === 'function') {
      themeMediaQuery.addEventListener('change', handleSystemThemeChange);
    } else if (typeof themeMediaQuery.addListener === 'function') {
      themeMediaQuery.addListener(handleSystemThemeChange);
    }
  }
}

function updatePageHeader() {
  if (!elements.topBar || !elements.topBar.title) {
    return;
  }
  const nextTitle = VIEW_TITLES[state.currentView] || 'Pesquisa de Stakeholders';
  elements.topBar.title.textContent = nextTitle;
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

function handleExportSnapshot() {
  try {
    const snapshot = collectDashboardSnapshot();
    const contents = JSON.stringify(snapshot, null, 2);
    const blob = new Blob([contents], { type: 'application/json' });
    const timestamp = new Date().toISOString().replace(/[:]/g, '-');
    const filename = `dashboard-stakeholders-${timestamp}.json`;
    triggerDownload(blob, filename);
  } catch (error) {
    console.error('Erro ao exportar os dados do dashboard:', error);
    if (typeof window !== 'undefined' && typeof window.alert === 'function') {
      window.alert('Não foi possível exportar os dados. Verifique o console para mais detalhes.');
    }
  }
}

function collectDashboardSnapshot() {
  const { stakeholders, insights } = getFilteredData();
  return {
    generatedAt: new Date().toISOString(),
    filters: {
      global: { ...state.filters },
      selections: { ...state.selections },
      currentView: state.currentView,
    },
    overview: buildOverviewSection(stakeholders, insights),
    dimension: buildDimensionSection(),
    expectations: buildExpectationsSection(stakeholders, insights),
    final: buildFinalSection(),
    rawData: {
      stakeholders: stakeholders.map((stakeholder) => ({ ...stakeholder })),
      insights: insights.map((insight) => ({ ...insight })),
    },
  };
}

function buildOverviewSection(stakeholders, insights) {
  const totalStakeholders = stakeholders.length;
  const totalInsights = insights.length;

  const averageTenure =
    totalStakeholders === 0
      ? 0
      : stakeholders.reduce((sum, stakeholder) => sum + (stakeholder.tempo_meses || 0), 0) /
        totalStakeholders;

  const teamCounts = stakeholders.reduce((acc, stakeholder) => {
    const team = stakeholder.time || 'Sem time';
    acc.set(team, (acc.get(team) || 0) + 1);
    return acc;
  }, new Map());

  const insightTypeCounts = insights.reduce((acc, insight) => {
    const type = insight.tipo || 'Não classificado';
    acc.set(type, (acc.get(type) || 0) + 1);
    return acc;
  }, new Map());

  const teamDistribution = mapCountsToArray(teamCounts, totalStakeholders, 'team');
  const insightDistribution = mapCountsToArray(insightTypeCounts, totalInsights, 'type');

  const topInsight = insightDistribution[0] || null;
  const topInsightType = topInsight
    ? {
        type: topInsight.type,
        count: topInsight.count,
        percentage: topInsight.percentage,
      }
    : null;

  return {
    metrics: {
      totalStakeholders,
      totalInsights,
      averageTenureMonths: Math.round(averageTenure),
    },
    teamDistribution,
    insightDistribution,
    topInsightType,
  };
}

function buildDimensionSection() {
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

  const categoryCounts = filteredInsights.reduce((acc, insight) => {
    const category = insight.categoria || 'Não classificado';
    acc.set(category, (acc.get(category) || 0) + 1);
    return acc;
  }, new Map());

  const categories = mapCountsToArray(categoryCounts, filteredInsights.length, 'category');

  const selectedCategory = state.selections.category;
  const categoryInsights =
    selectedCategory === null
      ? filteredInsights
      : filteredInsights.filter(
          (insight) => (insight.categoria || 'Não classificado') === selectedCategory,
        );

  const teamCounts = categoryInsights.reduce((acc, insight) => {
    const label = state.stakeholderMap.get(insight.stakeholder_id) || 'Sem time';
    acc.set(label, (acc.get(label) || 0) + 1);
    return acc;
  }, new Map());
  const teamDistribution = mapCountsToArray(teamCounts, categoryInsights.length, 'team');

  const tagsCounts = categoryInsights.reduce((acc, insight) => {
    (insight.tags || []).forEach((tag) => {
      const normalized = tag?.trim();
      if (!normalized) return;
      acc.set(normalized, (acc.get(normalized) || 0) + 1);
    });
    return acc;
  }, new Map());
  const tagTotal = Array.from(tagsCounts.values()).reduce((sum, value) => sum + value, 0);
  const tags = mapCountsToArray(tagsCounts, tagTotal, 'tag');

  const mentions = categoryInsights.map((insight) => {
    const stakeholder = state.stakeholderDetails.get(insight.stakeholder_id);
    return {
      id: insight._id,
      stakeholderId: insight.stakeholder_id ?? null,
      stakeholder: stakeholder ? buildStakeholderReference(stakeholder) : null,
      descricao: insight.descricao || null,
      categoria: insight.categoria || 'Não classificado',
      tipo: insight.tipo || 'Não classificado',
      dataEntrevista: insight.data_entrevista || null,
      tags: (insight.tags || []).filter(Boolean),
    };
  });

  return {
    filters: { team, insightType, selectedCategory },
    totals: {
      availableInsights: filteredInsights.length,
      selectedCategoryInsights: categoryInsights.length,
    },
    categories,
    teamDistribution,
    tags,
    mentions,
  };
}

function buildExpectationsSection(stakeholders, insights) {
  const totalStakeholders = stakeholders.length;
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

  const categoriesCounts = expectationInsights.reduce((acc, insight) => {
    const category = insight.categoria || 'Não classificado';
    acc.set(category, (acc.get(category) || 0) + 1);
    return acc;
  }, new Map());
  const categories = mapCountsToArray(
    categoriesCounts,
    expectationInsights.length,
    'category',
  );

  const selectedCategory = state.selections.expectationCategory;
  const filteredExpectations =
    selectedCategory === null
      ? expectationInsights
      : expectationInsights.filter(
          (insight) => (insight.categoria || 'Não classificado') === selectedCategory,
        );

  const highlights = filteredExpectations
    .slice()
    .sort(
      (a, b) =>
        new Date(b.data_entrevista || '1970-01-01').getTime() -
        new Date(a.data_entrevista || '1970-01-01').getTime(),
    )
    .map((insight) => {
      const stakeholder = state.stakeholderDetails.get(insight.stakeholder_id);
      return {
        id: insight._id,
        stakeholderId: insight.stakeholder_id ?? null,
        stakeholder: stakeholder ? buildStakeholderReference(stakeholder) : null,
        descricao: insight.descricao || null,
        categoria: insight.categoria || 'Não classificado',
        tipo: insight.tipo || 'Não classificado',
        dataEntrevista: insight.data_entrevista || null,
        tags: (insight.tags || []).filter(Boolean),
      };
    });

  const tagsCounts = filteredExpectations.reduce((acc, insight) => {
    (insight.tags || []).forEach((tag) => {
      const normalized = tag?.trim();
      if (!normalized) return;
      acc.set(normalized, (acc.get(normalized) || 0) + 1);
    });
    return acc;
  }, new Map());
  const tagTotal = Array.from(tagsCounts.values()).reduce((sum, value) => sum + value, 0);
  const tags = mapCountsToArray(tagsCounts, tagTotal, 'tag');

  const engagementTeamCounts = engagementInsights.reduce((acc, insight) => {
    const teamLabel = state.stakeholderMap.get(insight.stakeholder_id) || 'Sem time';
    acc.set(teamLabel, (acc.get(teamLabel) || 0) + 1);
    return acc;
  }, new Map());
  const engagementTeam = mapCountsToArray(
    engagementTeamCounts,
    engagementInsights.length,
    'team',
  );

  const rankingMap = engagementInsights.reduce((acc, insight) => {
    const id = insight.stakeholder_id;
    if (!id) return acc;
    const record = acc.get(id) || { count: 0, tags: new Map() };
    record.count += 1;
    (insight.tags || []).forEach((tag) => {
      const normalized = tag?.trim();
      if (!normalized) return;
      record.tags.set(normalized, (record.tags.get(normalized) || 0) + 1);
    });
    acc.set(id, record);
    return acc;
  }, new Map());

  const rankingEntries = Array.from(rankingMap.entries()).sort(
    (a, b) => b[1].count - a[1].count,
  );
  const maxMentions = rankingEntries[0]?.[1].count || 0;

  const ranking = rankingEntries.map(([stakeholderId, data]) => {
    const stakeholder = state.stakeholderDetails.get(stakeholderId);
    const topTags = Array.from(data.tags.entries())
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'))
      .slice(0, 5)
      .map(([tag, count]) => ({ tag, count }));
    return {
      stakeholderId,
      stakeholder: stakeholder ? buildStakeholderReference(stakeholder) : null,
      mentions: data.count,
      relativeScore: maxMentions ? Number(((data.count / maxMentions) * 100).toFixed(2)) : 0,
      topTags,
    };
  });

  return {
    filters: {
      selectedCategory,
    },
    metrics: {
      totalStakeholders,
      expectationStakeholders: {
        count: expectationStakeholders.size,
        percentage: percentageOf(expectationStakeholders.size, totalStakeholders),
      },
      expectationInsights: expectationInsights.length,
      engagementStakeholders: {
        count: engagementStakeholders.size,
        percentage: percentageOf(engagementStakeholders.size, totalStakeholders),
      },
      engagementInsights: engagementInsights.length,
    },
    categories,
    highlights,
    tags,
    engagementTeam,
    ranking,
  };
}

function buildFinalSection() {
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

  const themes = Array.from(groupedInsights.entries())
    .map(([label, insights]) => {
      if (!insights.length) {
        return { label, total: 0, destaque: null, secundario: null };
      }
      const counts = insights.reduce((acc, insight) => {
        const category = insight.categoria || 'Não classificado';
        acc.set(category, (acc.get(category) || 0) + 1);
        return acc;
      }, new Map());
      const ordered = Array.from(counts.entries()).sort(
        (a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'),
      );
      const primary = ordered[0] || null;
      const secondary = ordered[1] || null;
      return {
        label,
        total: insights.length,
        destaque: primary
          ? { category: primary[0], count: primary[1] }
          : null,
        secundario: secondary ? { category: secondary[0], count: secondary[1] } : null,
      };
    })
    .sort((a, b) => b.total - a.total || a.label.localeCompare(b.label, 'pt-BR'));

  const actionTags = state.insights.reduce((acc, insight) => {
    (insight.tags || []).forEach((tag) => {
      const normalized = tag?.trim();
      if (!normalized) return;
      acc.set(normalized, (acc.get(normalized) || 0) + 1);
    });
    return acc;
  }, new Map());

  const actions = Array.from(actionTags.entries())
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'))
    .slice(0, 6)
    .map(([tag, count]) => ({
      tag,
      recommendation: `Priorizar iniciativas relacionadas a ${tag}.`,
      supportCount: count,
    }));

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
    if ((tipo || '').toLowerCase() === 'dor') entry.dores += 1;
    if (expectationTypes.includes(tipo)) entry.expectativas += 1;
    if (tipo === 'Engajamento') entry.engajamento += 1;
    (insight.tags || []).forEach((tag) => {
      const label = tag?.trim();
      if (!label) return;
      entry.tags.set(label, (entry.tags.get(label) || 0) + 1);
    });
    influenceMap.set(id, entry);
  });

  const influence = Array.from(influenceMap.entries())
    .map(([stakeholderId, data]) => {
      const total = data.dores + data.expectativas + data.engajamento;
      if (total === 0) return null;
      const stakeholder = state.stakeholderDetails.get(stakeholderId);
      const topTags = Array.from(data.tags.entries())
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'))
        .slice(0, 5)
        .map(([tag, count]) => ({ tag, count }));
      return {
        stakeholderId,
        stakeholder: stakeholder ? buildStakeholderReference(stakeholder) : null,
        totals: {
          dores: data.dores,
          expectativas: data.expectativas,
          engajamento: data.engajamento,
        },
        totalMentions: total,
        topTags,
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.totalMentions - a.totalMentions);

  return {
    metrics: {
      totalStakeholders,
      totalInsights,
      expectationStakeholders: {
        count: expectationStakeholders.size,
        percentage: percentageOf(expectationStakeholders.size, totalStakeholders),
      },
      engagementStakeholders: {
        count: engagementStakeholders.size,
        percentage: percentageOf(engagementStakeholders.size, totalStakeholders),
      },
    },
    themes,
    actions,
    influence,
  };
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function mapCountsToArray(map, total, labelKey) {
  return Array.from(map.entries())
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'))
    .map(([label, count]) => {
      const entry = {
        count,
        percentage: total ? Number(((count / total) * 100).toFixed(2)) : 0,
      };
      entry[labelKey] = label;
      return entry;
    });
}

function buildStakeholderReference(stakeholder) {
  if (!stakeholder) return null;
  return {
    id: stakeholder.id ?? null,
    nome: stakeholder.nome || null,
    cargo: stakeholder.cargo || null,
    time: stakeholder.time || null,
    area: stakeholder.area || null,
    senioridade: stakeholder.senioridade || null,
    tempo_meses: stakeholder.tempo_meses ?? null,
  };
}

function percentageOf(part, total) {
  if (!total || total === 0) return 0;
  return Number(((part / total) * 100).toFixed(2));
}
