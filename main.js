(function () {
  const App = window.OperationalAnalytics;
  const state = App.state;

  const dom = {};

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    cacheDom();
    checkDependencies();
    bindEvents();
    renderAll();
  }

  function cacheDom() {
    dom.dependencyStatus = document.getElementById("dependencyStatus");
    dom.periodList = document.getElementById("periodList");
    dom.addPeriodButton = document.getElementById("addPeriodButton");
    dom.comparisonModeSelect = document.getElementById("comparisonModeSelect");
    dom.previewList = document.getElementById("previewList");
    dom.mapperControls = document.getElementById("mapperControls");
    dom.addMetricButton = document.getElementById("addMetricButton");
    dom.metricList = document.getElementById("metricList");
    dom.analyzeButton = document.getElementById("analyzeButton");
    dom.exportCsvButton = document.getElementById("exportCsvButton");
    dom.exportExcelButton = document.getElementById("exportExcelButton");
    dom.warningsPanel = document.getElementById("warningsPanel");
    dom.summaryCards = document.getElementById("summaryCards");
    dom.chartMetricSelect = document.getElementById("chartMetricSelect");
    dom.deltaChart = document.getElementById("deltaChart");
    dom.moversPanel = document.getElementById("moversPanel");
    dom.resultsTable = document.getElementById("resultsTable");
  }

  function checkDependencies() {
    const missing = [];

    if (!window.XLSX) {
      missing.push("SheetJS");
    }

    if (!window.Chart) {
      missing.push("Chart.js");
    }

    if (!window.ExcelJS) {
      missing.push("ExcelJS");
    }

    if (missing.length) {
      dom.dependencyStatus.className = "header-status error";
      dom.dependencyStatus.textContent = "Не загружено: " + missing.join(", ");
      return;
    }

    dom.dependencyStatus.className = "header-status ok";
    dom.dependencyStatus.textContent = "Локальные библиотеки готовы";
  }

  function bindEvents() {
    dom.addPeriodButton.addEventListener("click", addPeriod);
    dom.comparisonModeSelect.addEventListener("change", function () {
      state.comparisonMode = dom.comparisonModeSelect.value;
      clearAnalysis();
      renderAll();
    });
    dom.periodList.addEventListener("change", handlePeriodChange);
    dom.periodList.addEventListener("input", handlePeriodInput);
    dom.periodList.addEventListener("click", handlePeriodClick);
    dom.mapperControls.addEventListener("change", handleMappingChange);
    dom.addMetricButton.addEventListener("click", addMetric);
    dom.metricList.addEventListener("change", handleMetricInput);
    dom.metricList.addEventListener("click", handleMetricClick);
    dom.analyzeButton.addEventListener("click", analyze);
    dom.exportCsvButton.addEventListener("click", exportCsv);
    dom.exportExcelButton.addEventListener("click", exportExcel);

    dom.chartMetricSelect.addEventListener("change", function () {
      state.selectedChartMetricId = dom.chartMetricSelect.value;
      renderChart();
    });
  }

  function addPeriod() {
    state.periods.push(App.createPeriod(state.periods.length));
    clearAnalysis();
    renderAll();
  }

  function handlePeriodInput(event) {
    if (event.target.name !== "periodLabel") {
      return;
    }

    const period = findPeriod(event.target.dataset.periodId);
    if (!period) {
      return;
    }

    period.label = event.target.value.trim() || getPeriodFallbackLabel(period);
    clearAnalysis();
    renderAnalysis();
  }

  function handlePeriodChange(event) {
    if (event.target.name === "periodFile") {
      handleFileChange(event.target.dataset.periodId, event.target.files[0]);
      return;
    }

    if (event.target.name === "periodLabel") {
      renderAll();
    }
  }

  function handlePeriodClick(event) {
    if (event.target.dataset.action !== "remove-period") {
      return;
    }

    const periodId = event.target.dataset.periodId;

    if (state.periods.length <= 2) {
      return;
    }

    state.periods = state.periods.filter(function (period) {
      return period.id !== periodId;
    });

    state.mapping.metrics.forEach(function (metric) {
      delete metric.columns[periodId];
    });

    clearAnalysis();
    renderAll();
  }

  async function handleFileChange(periodId, file) {
    const period = findPeriod(periodId);

    if (!period || !file) {
      return;
    }

    period.loading = true;
    renderPeriodUploads();

    try {
      state.messages = [];
      const table = await App.ExcelReader.readExcelFile(file);
      period.file = file;
      period.table = table;
      period.idColumn = "";
      period.loading = false;

      state.mapping.metrics.forEach(function (metric) {
        metric.columns[period.id] = "";
      });

      clearAnalysis();
      renderAll();
    } catch (error) {
      period.loading = false;
      period.file = file;
      period.table = null;
      period.idColumn = "";
      state.messages.push({
        type: "error",
        message: error.message,
      });
      clearAnalysis();
      renderAll();
    }
  }

  function clearAnalysis() {
    state.comparison = null;
    state.analytics = null;
    state.selectedChartMetricId = "";
  }

  function renderAll() {
    dom.comparisonModeSelect.value = state.comparisonMode;
    renderPeriodUploads();
    renderPreviews();
    renderColumnMapping();
    renderMetrics();
    renderAnalysis();
    App.UI.renderWarnings(dom.warningsPanel, collectWarnings());
  }

  function renderPeriodUploads() {
    dom.periodList.innerHTML = state.periods.map(renderPeriodCard).join("");
  }

  function renderPeriodCard(period, index) {
    const fileName = period.loading
      ? "Чтение файла..."
      : period.table
        ? period.file.name
        : period.file
          ? "Ошибка чтения"
          : "Файл не выбран";
    const removeButton =
      state.periods.length > 2
        ? '<button class="period-remove" type="button" data-action="remove-period" data-period-id="' +
          period.id +
          '" title="Удалить период">×</button>'
        : "";

    return (
      '<div class="period-card" data-period-id="' +
      period.id +
      '">' +
      '<div class="period-card__top">' +
      '<span class="period-badge">' +
      String(index + 1).padStart(2, "0") +
      "</span>" +
      removeButton +
      "</div>" +
      '<label class="field period-name-field"><span>Название периода</span><input type="text" name="periodLabel" data-period-id="' +
      period.id +
      '" value="' +
      App.UI.escapeHtml(period.label) +
      '" /></label>' +
      '<label class="file-drop file-drop--compact">' +
      '<span class="file-drop__label">' +
      App.UI.escapeHtml(period.label) +
      "</span>" +
      '<span class="file-drop__hint">Excel, CSV или TSV</span>' +
      '<input name="periodFile" data-period-id="' +
      period.id +
      '" type="file" accept=".xlsx,.xls,.csv,.tsv" />' +
      "<strong>" +
      App.UI.escapeHtml(fileName) +
      "</strong>" +
      "</label>" +
      "</div>"
    );
  }

  function renderPreviews() {
    dom.previewList.innerHTML = "";

    state.periods.forEach(function (period) {
      const panel = document.createElement("div");
      panel.className = "preview-panel empty-state";
      dom.previewList.appendChild(panel);
      App.UI.renderPreview(panel, period.table, "Загрузите файл: " + period.label);
    });
  }

  function renderColumnMapping() {
    if (!state.periods.length) {
      dom.mapperControls.className = "mapper-controls empty-state";
      dom.mapperControls.textContent = "Добавьте период";
      dom.addMetricButton.disabled = true;
      return;
    }

    dom.mapperControls.className = "mapper-controls";
    dom.mapperControls.innerHTML = state.periods.map(renderIdColumnSelect).join("");
    dom.addMetricButton.disabled = !areIdsReady();
  }

  function renderIdColumnSelect(period) {
    return (
      '<label class="field"><span>Юнит: ' +
      App.UI.escapeHtml(period.label) +
      '</span><select name="idColumn" data-period-id="' +
      period.id +
      '"' +
      (period.table ? "" : " disabled") +
      ">" +
      buildColumnOptions(period.table, period.idColumn, "Выберите колонку с юнитом") +
      "</select></label>"
    );
  }

  function handleMappingChange(event) {
    if (event.target.name !== "idColumn") {
      return;
    }

    const period = findPeriod(event.target.dataset.periodId);

    if (!period) {
      return;
    }

    period.idColumn = event.target.value;
    clearAnalysis();
    renderAll();
  }

  function renderMetrics() {
    syncMetricLabels();

    if (!areAllPeriodsLoaded()) {
      dom.metricList.className = "metric-list empty-state";
      dom.metricList.textContent = "Загрузите минимум два периода, чтобы добавить показатели";
      dom.analyzeButton.disabled = true;
      return;
    }

    if (!areIdsReady()) {
      dom.metricList.className = "metric-list empty-state";
      dom.metricList.textContent = "Выберите колонки с юнитами для всех периодов";
      dom.analyzeButton.disabled = true;
      return;
    }

    if (!state.mapping.metrics.length) {
      dom.metricList.className = "metric-list empty-state";
      dom.metricList.textContent = "Добавьте хотя бы один показатель";
      dom.analyzeButton.disabled = true;
      return;
    }

    dom.metricList.className = "metric-list";
    dom.metricList.innerHTML = state.mapping.metrics.map(renderMetricRow).join("");
    dom.analyzeButton.disabled = !isReadyToAnalyze();
  }

  function renderMetricRow(metric) {
    return (
      '<div class="metric-row metric-row--multi" data-metric-id="' +
      metric.id +
      '">' +
      '<div class="metric-name"><span>Показатель</span><strong>' +
      App.UI.escapeHtml(metric.label || "Выберите столбцы") +
      "</strong></div>" +
      '<div class="metric-columns">' +
      state.periods
        .map(function (period) {
          return renderMetricPeriodSelect(metric, period);
        })
        .join("") +
      "</div>" +
      '<button class="icon-button" type="button" data-action="remove-metric" title="Удалить показатель">×</button>' +
      "</div>"
    );
  }

  function renderMetricPeriodSelect(metric, period) {
    return (
      '<label class="field"><span>' +
      App.UI.escapeHtml(period.label) +
      '</span><select name="metricColumn" data-period-id="' +
      period.id +
      '"' +
      (period.table ? "" : " disabled") +
      ">" +
      buildColumnOptions(period.table, metric.columns[period.id], "Выберите показатель") +
      "</select></label>"
    );
  }

  function buildColumnOptions(table, selectedValue, placeholder) {
    const options = ['<option value="">' + App.UI.escapeHtml(placeholder) + "</option>"];

    if (table) {
      table.headers.forEach(function (header) {
        options.push(
          '<option value="' +
            header.id +
            '" ' +
            selected(selectedValue, header.id) +
            ">" +
            App.UI.escapeHtml(header.name) +
            "</option>"
        );
      });
    }

    return options.join("");
  }

  function selected(value, expected) {
    return value === expected ? "selected" : "";
  }

  function addMetric() {
    const columns = {};

    state.periods.forEach(function (period) {
      columns[period.id] = getDefaultMetricColumn(period);
    });

    state.mapping.metrics.push(App.createMetric(state.mapping.metrics.length, columns));
    clearAnalysis();
    renderAll();
  }

  function getDefaultMetricColumn(period) {
    if (!period.table) {
      return "";
    }

    const header = period.table.headers.find(function (item) {
      return item.id !== period.idColumn;
    });

    return header ? header.id : "";
  }

  function handleMetricInput(event) {
    if (event.target.name !== "metricColumn") {
      return;
    }

    const row = event.target.closest("[data-metric-id]");
    if (!row) {
      return;
    }

    const metric = findMetric(row.dataset.metricId);
    if (!metric) {
      return;
    }

    metric.columns[event.target.dataset.periodId] = event.target.value;
    syncMetricLabels();
    clearAnalysis();
    renderMetrics();
    renderAnalysis();
    App.UI.renderWarnings(dom.warningsPanel, collectWarnings());
  }

  function handleMetricClick(event) {
    if (event.target.dataset.action !== "remove-metric") {
      return;
    }

    const row = event.target.closest("[data-metric-id]");
    state.mapping.metrics = state.mapping.metrics.filter(function (metric) {
      return metric.id !== row.dataset.metricId;
    });
    clearAnalysis();
    renderAll();
  }

  function findMetric(metricId) {
    return state.mapping.metrics.find(function (metric) {
      return metric.id === metricId;
    });
  }

  function findPeriod(periodId) {
    return state.periods.find(function (period) {
      return period.id === periodId;
    });
  }

  function areAllPeriodsLoaded() {
    return (
      state.periods.length >= 2 &&
      state.periods.every(function (period) {
        return Boolean(period.table);
      })
    );
  }

  function areIdsReady() {
    return (
      areAllPeriodsLoaded() &&
      state.periods.every(function (period) {
        return Boolean(period.idColumn);
      })
    );
  }

  function isReadyToAnalyze() {
    return Boolean(
      areIdsReady() &&
        state.mapping.metrics.length &&
        state.mapping.metrics.every(function (metric) {
          return state.periods.every(function (period) {
            return Boolean(metric.columns[period.id]);
          });
        })
    );
  }

  function analyze() {
    state.messages = [];
    syncMetricLabels();

    if (!isReadyToAnalyze()) {
      state.messages.push({
        type: "error",
        message: "Для расчета нужны минимум два загруженных периода, колонки с юнитами и хотя бы один показатель",
      });
      renderAll();
      return;
    }

    state.comparison = App.Comparator.comparePeriods({
      periods: state.periods,
      metrics: state.mapping.metrics,
      comparisonMode: state.comparisonMode,
    });

    state.analytics = App.Analytics.buildAnalytics(state.comparison, state.mapping.metrics);
    state.selectedChartMetricId = state.mapping.metrics[0].id;
    renderAll();
  }

  function renderAnalysis() {
    App.UI.renderSummary(dom.summaryCards, state.analytics);
    App.UI.renderChartMetricSelect(dom.chartMetricSelect, state.mapping.metrics, state.selectedChartMetricId);
    dom.chartMetricSelect.disabled = !state.comparison || !state.mapping.metrics.length;
    App.UI.renderMovers(dom.moversPanel, state.analytics);
    App.UI.renderResultsTable(dom.resultsTable, state.comparison, state.mapping.metrics);
    renderChart();

    const hasComparison = Boolean(state.comparison);
    dom.exportCsvButton.disabled = !hasComparison;
    dom.exportExcelButton.disabled = !hasComparison;
  }

  function renderChart() {
    const metric = findMetric(state.selectedChartMetricId) || state.mapping.metrics[0];
    App.Charts.renderDeltaChart(dom.deltaChart, state.comparison, metric);
  }

  function syncMetricLabels() {
    state.mapping.metrics.forEach(function (metric) {
      metric.label = getMetricLabel(metric);
    });
  }

  function getMetricLabel(metric) {
    const lastPeriod = state.periods[state.periods.length - 1];
    const firstPeriod = state.periods[0];

    return (
      getColumnName(lastPeriod, metric.columns[lastPeriod.id]) ||
      getColumnName(firstPeriod, metric.columns[firstPeriod.id]) ||
      ""
    );
  }

  function getColumnName(period, columnId) {
    if (!period || !period.table || !columnId) {
      return "";
    }

    const header = period.table.headers.find(function (item) {
      return item.id === columnId;
    });

    return header ? header.name : "";
  }

  function collectWarnings() {
    const warnings = state.messages.slice();

    state.periods.forEach(function (period) {
      if (period.table) {
        warnings.push.apply(warnings, period.table.warnings);
      }
    });

    warnings.push.apply(warnings, compareHeadersWarning());

    if (state.comparison) {
      appendComparisonWarnings(warnings);
    }

    return warnings;
  }

  function compareHeadersWarning() {
    const loadedPeriods = state.periods.filter(function (period) {
      return period.table;
    });

    if (loadedPeriods.length < 2) {
      return [];
    }

    const base = loadedPeriods[0];
    const baseNames = new Set(
      base.table.headers.map(function (header) {
        return App.Normalizers.normalizeKey(header.name);
      })
    );
    const messages = [];

    loadedPeriods.slice(1).forEach(function (period) {
      const names = new Set(
        period.table.headers.map(function (header) {
          return App.Normalizers.normalizeKey(header.name);
        })
      );
      const onlyBase = Array.from(baseNames).filter(function (name) {
        return name && !names.has(name);
      });
      const onlyCurrent = Array.from(names).filter(function (name) {
        return name && !baseNames.has(name);
      });

      if (onlyBase.length || onlyCurrent.length) {
        messages.push({
          type: "warn",
          message:
            "Структуры «" +
            base.label +
            "» и «" +
            period.label +
            "» отличаются. Только в первом: " +
            previewList(onlyBase) +
            ". Только во втором: " +
            previewList(onlyCurrent) +
            ".",
        });
      }
    });

    return messages;
  }

  function previewList(items) {
    if (!items.length) {
      return "нет";
    }

    const visible = items.slice(0, 6).join(", ");
    return items.length > 6 ? visible + " +" + (items.length - 6) : visible;
  }

  function appendComparisonWarnings(warnings) {
    const comparison = state.comparison;

    comparison.emptyIdsByPeriod.forEach(function (item) {
      if (item.items.length) {
        warnings.push({
          type: "warn",
          message: "В периоде «" + item.periodLabel + "» есть строки без юнита: " + item.items.length,
        });
      }
    });

    comparison.duplicatesByPeriod.forEach(function (item) {
      if (item.items.length) {
        warnings.push({
          type: "warn",
          message: "В периоде «" + item.periodLabel + "» есть дубли юнитов: " + item.items.length,
        });
      }
    });

    comparison.missingByPeriod.forEach(function (item) {
      if (item.items.length) {
        warnings.push({
          type: "warn",
          message: "В периоде «" + item.periodLabel + "» отсутствуют юниты: " + item.items.length,
        });
      }
    });

    if (comparison.invalidValues.length) {
      warnings.push({
        type: "warn",
        message: "Найдены нечисловые значения в выбранных показателях: " + comparison.invalidValues.length,
      });
    }
  }

  function exportCsv() {
    if (!state.comparison) {
      return;
    }

    App.Exporters.exportCsv(state.comparison, state.mapping.metrics);
  }

  async function exportExcel() {
    if (!state.comparison) {
      return;
    }

    try {
      await App.Exporters.exportExcel(state.comparison, state.mapping.metrics, state.analytics);
    } catch (error) {
      state.messages.push({
        type: "error",
        message: "Не удалось сформировать Excel: " + error.message,
      });
      renderAll();
    }
  }

  function getPeriodFallbackLabel(period) {
    const index = state.periods.findIndex(function (item) {
      return item.id === period.id;
    });

    return "Период " + (index + 1);
  }
})();
