(function () {
  const App = window.OperationalAnalytics;
  const state = App.state;
  const SINGLE_FILE_METRIC_COLUMN_KEY = "__singleFileMetricColumn";

  const dom = {};
  let historyRecords = [];
  let hasUnsavedAnalysis = false;
  let historySearchQuery = "";

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    cacheDom();
    checkDependencies();
    historyRecords = App.HistoryStore ? App.HistoryStore.load() : [];
    bindEvents();
    initMotion();
    renderAll();
  }

  function cacheDom() {
    dom.dependencyStatus = document.getElementById("dependencyStatus");
    dom.periodSourceModeSelect = document.getElementById("periodSourceModeSelect");
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
    dom.saveAnalysisButton = document.getElementById("saveAnalysisButton");
    dom.warningsPanel = document.getElementById("warningsPanel");
    dom.globalImpactFilter = document.getElementById("globalImpactFilter");
    dom.globalDeltaMinFilter = document.getElementById("globalDeltaMinFilter");
    dom.globalDeltaMaxFilter = document.getElementById("globalDeltaMaxFilter");
    dom.globalObjectSearch = document.getElementById("globalObjectSearch");
    dom.globalDepartmentSearch = document.getElementById("globalDepartmentSearch");
    dom.globalFilterStatus = document.getElementById("globalFilterStatus");
    dom.resetGlobalFiltersButton = document.getElementById("resetGlobalFiltersButton");
    dom.clearHistoryButton = document.getElementById("clearHistoryButton");
    dom.historyList = document.getElementById("historyList");
    dom.historyTitle = document.getElementById("historyTitle");
    dom.historySearchInput = document.getElementById("historySearchInput");
    dom.summaryCards = document.getElementById("summaryCards");
    dom.chartMetricSelect = document.getElementById("chartMetricSelect");
    dom.deltaChart = document.getElementById("deltaChart");
    dom.moversPanel = document.getElementById("moversPanel");
    dom.resultsTable = document.getElementById("resultsTable");
    dom.saveAnalysisModal = document.getElementById("saveAnalysisModal");
    dom.saveAnalysisForm = document.getElementById("saveAnalysisForm");
    dom.saveAnalysisTitle = document.getElementById("saveAnalysisTitle");
    dom.saveAnalysisDescription = document.getElementById("saveAnalysisDescription");
    dom.saveAnalysisNameInput = document.getElementById("saveAnalysisNameInput");
    dom.saveAnalysisCancelButton = document.getElementById("saveAnalysisCancelButton");
    dom.saveAnalysisCloseButton = document.getElementById("saveAnalysisCloseButton");
    dom.saveAnalysisSubmitButton = document.getElementById("saveAnalysisSubmitButton");
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

    if (!dom.dependencyStatus) {
      if (missing.length) {
        console.error("Не загружены локальные библиотеки: " + missing.join(", "));
      }

      return;
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
    dom.periodSourceModeSelect.addEventListener("change", function () {
      state.periodSourceMode = dom.periodSourceModeSelect.value;
      state.messages = [];
      clearAnalysis();
      renderAll();
    });
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
    dom.saveAnalysisButton.addEventListener("click", openSaveAnalysisModal);
    dom.globalImpactFilter.addEventListener("change", handleGlobalFilterChange);
    dom.globalDeltaMinFilter.addEventListener("input", handleGlobalFilterChange);
    dom.globalDeltaMaxFilter.addEventListener("input", handleGlobalFilterChange);
    dom.globalObjectSearch.addEventListener("input", handleGlobalFilterChange);
    dom.globalDepartmentSearch.addEventListener("input", handleGlobalFilterChange);
    dom.resetGlobalFiltersButton.addEventListener("click", resetGlobalFilters);
    dom.clearHistoryButton.addEventListener("click", clearHistory);
    dom.historyList.addEventListener("click", handleHistoryClick);
    dom.historySearchInput.addEventListener("input", function () {
      historySearchQuery = dom.historySearchInput.value.trim().toLowerCase();
      renderHistoryPanel();
      refreshMotion();
    });
    dom.saveAnalysisForm.addEventListener("submit", saveAnalysisFromModal);
    dom.saveAnalysisCancelButton.addEventListener("click", closeSaveAnalysisModal);
    dom.saveAnalysisCloseButton.addEventListener("click", closeSaveAnalysisModal);
    dom.saveAnalysisModal.addEventListener("click", function (event) {
      if (event.target === dom.saveAnalysisModal) {
        closeSaveAnalysisModal();
      }
    });
    document.addEventListener("keydown", function (event) {
      if (event.key === "Escape" && !dom.saveAnalysisModal.hidden) {
        closeSaveAnalysisModal();
      }
    });
    window.addEventListener("beforeunload", handleBeforeUnload);
    window.addEventListener("operationalAnalytics:themechange", function () {
      renderChart();
    });

    dom.chartMetricSelect.addEventListener("change", function () {
      state.selectedChartMetricId = dom.chartMetricSelect.value;
      renderAnalysis();
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
    if (event.target.name === "singleFile") {
      handleSingleFileChange(event.target.files[0]);
      return;
    }

    if (event.target.name === "periodFile") {
      handleFileChange(event.target.dataset.periodId, event.target.files[0]);
      return;
    }

    if (event.target.name === "periodLabel") {
      renderAll();
    }
  }

  function handlePeriodClick(event) {
    const actionButton = event.target.closest("[data-action]");

    if (!actionButton) {
      return;
    }

    const periodId = actionButton.dataset.periodId;

    if (actionButton.dataset.action === "clear-single-file") {
      clearSingleFile();
      return;
    }

    if (actionButton.dataset.action === "clear-period-file") {
      clearPeriodFile(periodId);
      return;
    }

    if (actionButton.dataset.action !== "remove-period") {
      return;
    }

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

  function clearSingleFile() {
    state.singleFile = App.createSingleFileSource ? App.createSingleFileSource() : {
      file: null,
      table: null,
      periodColumn: "",
      idColumn: "",
      virtualPeriods: [],
      warnings: [],
      loading: false,
    };
    state.messages = [];
    clearAnalysis();
    renderAll();
  }

  function clearPeriodFile(periodId) {
    const period = findPeriod(periodId);

    if (!period) {
      return;
    }

    period.file = null;
    period.table = null;
    period.idColumn = "";
    period.loading = false;
    state.messages = [];

    state.mapping.metrics.forEach(function (metric) {
      metric.columns[periodId] = "";
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

  async function handleSingleFileChange(file) {
    const singleFile = ensureSingleFileSource();

    if (!file) {
      return;
    }

    singleFile.loading = true;
    renderPeriodUploads();

    try {
      state.messages = [];
      const table = await App.ExcelReader.readExcelFile(file);
      singleFile.file = file;
      singleFile.table = table;
      singleFile.periodColumn = "";
      singleFile.idColumn = "";
      singleFile.virtualPeriods = [];
      singleFile.warnings = [];
      singleFile.loading = false;
      clearAnalysis();
      renderAll();
    } catch (error) {
      singleFile.loading = false;
      singleFile.file = file;
      singleFile.table = null;
      singleFile.periodColumn = "";
      singleFile.idColumn = "";
      singleFile.virtualPeriods = [];
      singleFile.warnings = [];
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
    resetGlobalFiltersState();
    state.restoredHistoryMeta = null;
    state.restoredHistorySettings = null;
    hasUnsavedAnalysis = false;
  }

  function renderAll() {
    dom.periodSourceModeSelect.value = state.periodSourceMode || "multiFile";
    dom.comparisonModeSelect.value = state.comparisonMode;
    dom.addPeriodButton.disabled = isSingleFileSourceMode();
    renderPeriodUploads();
    renderPreviews();
    renderColumnMapping();
    renderMetrics();
    renderAnalysis();
    renderHistoryPanel();
    renderWarningsPanel();
    refreshMotion();
  }

  function renderPeriodUploads() {
    if (isSingleFileSourceMode()) {
      dom.periodList.className = "period-list sidebar-period-list";
      dom.periodList.innerHTML = renderSingleFileCard();
      return;
    }

    dom.periodList.className = "period-list sidebar-period-list";
    dom.periodList.innerHTML = state.periods.map(renderPeriodCard).join("");
  }

  function renderSingleFileCard() {
    const singleFile = ensureSingleFileSource();
    const fileName = singleFile.loading
      ? "Чтение файла..."
      : singleFile.table
        ? singleFile.file.name
        : singleFile.file
          ? "Ошибка чтения"
          : "Файл не выбран";
    const clearFileDisabled = singleFile.loading || (!singleFile.file && !singleFile.table) ? " disabled" : "";

    return (
      '<div class="period-card single-file-card">' +
      '<div class="single-file-mode-note">' +
      "<strong>Один файл</strong>" +
      "<span>Загрузите файл, где периоды находятся в одной из колонок.</span>" +
      "</div>" +
      '<label class="file-drop file-drop--compact">' +
      '<span class="file-drop__label">Один файл</span>' +
      '<span class="file-drop__hint">Excel, CSV или TSV</span>' +
      '<span class="file-drop__button">Загрузить</span>' +
      '<input name="singleFile" type="file" accept=".xlsx,.xls,.csv,.tsv" />' +
      "<strong>" +
      App.UI.escapeHtml(fileName) +
      "</strong>" +
      "</label>" +
      '<div class="period-file-actions">' +
      '<button class="button button-secondary file-clear-button" type="button" data-action="clear-single-file" title="Удалить файл" aria-label="Удалить файл"' +
      clearFileDisabled +
      ">×</button>" +
      "</div>" +
      "</div>"
    );
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
          '" title="Удалить период" aria-label="Удалить период">×</button>'
        : "";
    const periodActions = removeButton ? '<div class="period-card__top">' + removeButton + "</div>" : "";
    const clearFileDisabled = period.loading || (!period.file && !period.table) ? " disabled" : "";
    const fileActionLabel = "Загрузить";

    return (
      '<div class="period-card" data-period-id="' +
      period.id +
      '">' +
      periodActions +
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
      '<span class="file-drop__button">' +
      fileActionLabel +
      "</span>" +
      '<input name="periodFile" data-period-id="' +
      period.id +
      '" type="file" accept=".xlsx,.xls,.csv,.tsv" />' +
      "<strong>" +
      App.UI.escapeHtml(fileName) +
      "</strong>" +
      "</label>" +
      '<div class="period-file-actions">' +
      '<button class="button button-secondary file-clear-button" type="button" data-action="clear-period-file" data-period-id="' +
      period.id +
      '" title="Удалить файл" aria-label="Удалить файл"' +
      clearFileDisabled +
      ">×</button>" +
      "</div>" +
      "</div>"
    );
  }

  function renderPreviews() {
    dom.previewList.innerHTML = "";

    if (isSingleFileSourceMode()) {
      const singleFile = ensureSingleFileSource();
      const virtualPeriods = getSingleFileVirtualPeriods();

      if (virtualPeriods.length) {
        virtualPeriods.forEach(function (period) {
          const panel = document.createElement("div");
          panel.className = "preview-panel empty-state";
          dom.previewList.appendChild(panel);
          App.UI.renderPreview(panel, period.table, "Нет строк: " + period.label);
        });
        return;
      }

      const panel = document.createElement("div");
      panel.className = "preview-panel empty-state";
      dom.previewList.appendChild(panel);
      App.UI.renderPreview(panel, singleFile.table, "Загрузите один файл");
      return;
    }

    state.periods.forEach(function (period) {
      const panel = document.createElement("div");
      panel.className = "preview-panel empty-state";
      dom.previewList.appendChild(panel);
      App.UI.renderPreview(panel, period.table, "Загрузите файл: " + period.label);
    });
  }

  function renderColumnMapping() {
    if (isSingleFileSourceMode()) {
      renderSingleFileColumnMapping();
      dom.addMetricButton.disabled = !isSingleFileReadyForMetrics();
      return;
    }

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

  function renderSingleFileColumnMapping() {
    const singleFile = ensureSingleFileSource();

    if (!singleFile.table) {
      dom.mapperControls.className = "mapper-controls empty-state";
      dom.mapperControls.textContent = "Загрузите один файл для выбора объекта и колонки периода";
      return;
    }

    dom.mapperControls.className = "mapper-controls";
    dom.mapperControls.innerHTML =
      '<label class="field"><span>Объект</span><select name="singleIdColumn">' +
      buildColumnOptions(singleFile.table, singleFile.idColumn, "Выберите объект") +
      "</select></label>" +
      '<label class="field"><span>Колонка периода</span><select name="singlePeriodColumn">' +
      buildColumnOptions(singleFile.table, singleFile.periodColumn, "Выберите колонку периода") +
      "</select></label>" +
      renderSingleFilePeriodSummary();
  }

  function renderSingleFilePeriodSummary() {
    const singleFile = ensureSingleFileSource();

    if (!singleFile.table || !singleFile.periodColumn) {
      return '<div class="single-file-mode-note"><strong>Периоды не выбраны</strong><span>Выберите колонку с датой, месяцем, неделей или годом.</span></div>';
    }

    const values = getSingleFilePeriodValues();
    const virtualPeriods = getSingleFileVirtualPeriods();
    const rowCount = virtualPeriods.reduce(function (sum, period) {
      return sum + period.table.rows.length;
    }, 0);
    const preview = values.slice(0, 4).join(", ");
    const rest = values.length > 4 ? " +" + (values.length - 4) : "";

    return (
      '<div class="single-file-mode-note"><strong>Найдено периодов: ' +
      values.length +
      "</strong><span>" +
      App.UI.escapeHtml("Строк: " + rowCount + ". " + preview + rest) +
      "</span></div>"
    );
  }

  function renderIdColumnSelect(period) {
    return (
      '<label class="field"><span>Объект: ' +
      App.UI.escapeHtml(period.label) +
      '</span><select name="idColumn" data-period-id="' +
      period.id +
      '"' +
      (period.table ? "" : " disabled") +
      ">" +
      buildColumnOptions(period.table, period.idColumn, "Выберите объект") +
      "</select></label>"
    );
  }

  function handleMappingChange(event) {
    if (event.target.name === "singleIdColumn") {
      ensureSingleFileSource().idColumn = event.target.value;
      clearAnalysis();
      renderAll();
      return;
    }

    if (event.target.name === "singlePeriodColumn") {
      ensureSingleFileSource().periodColumn = event.target.value;
      rebuildSingleFileVirtualPeriods();
      clearAnalysis();
      renderAll();
      return;
    }

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
    if (isSingleFileSourceMode()) {
      renderSingleFileMetrics();
      return;
    }

    if (areAllPeriodsLoaded()) {
      syncMetricLabels();
    }

    if (!areAllPeriodsLoaded()) {
      dom.metricList.className = "metric-list empty-state";
      dom.metricList.textContent = "Загрузите минимум два периода";
      dom.analyzeButton.disabled = true;
      return;
    }

    if (!areIdsReady()) {
      dom.metricList.className = "metric-list empty-state";
      dom.metricList.textContent = "Выберите объект для всех периодов";
      dom.analyzeButton.disabled = true;
      return;
    }

    if (!state.mapping.metrics.length) {
      dom.metricList.className = "metric-list empty-state";
      dom.metricList.textContent = "Добавьте показатель";
      dom.analyzeButton.disabled = true;
      return;
    }

    dom.metricList.className = "metric-list";
    dom.metricList.innerHTML = state.mapping.metrics.map(renderMetricRow).join("");
    dom.analyzeButton.disabled = !isReadyToAnalyze();
  }

  function renderSingleFileMetrics() {
    const setupMessage = getSingleFileMetricSetupMessage();

    if (setupMessage) {
      dom.metricList.className = "metric-list empty-state";
      dom.metricList.textContent = setupMessage;
      dom.analyzeButton.disabled = true;
      return;
    }

    syncMetricLabels();

    if (!state.mapping.metrics.length) {
      dom.metricList.className = "metric-list empty-state";
      dom.metricList.textContent = "Добавьте показатель";
      dom.analyzeButton.disabled = true;
      return;
    }

    dom.metricList.className = "metric-list";
    dom.metricList.innerHTML = state.mapping.metrics.map(renderSingleFileMetricRow).join("");
    dom.analyzeButton.disabled = !isReadyToAnalyze();
  }

  function getSingleFileMetricSetupMessage() {
    const singleFile = ensureSingleFileSource();

    if (!singleFile.table) {
      return "Загрузите один файл";
    }

    if (!singleFile.idColumn) {
      return "Выберите объект для сравнения";
    }

    if (!singleFile.periodColumn) {
      return "Выберите колонку периода";
    }

    if (singleFile.idColumn === singleFile.periodColumn) {
      return "Колонка объекта и колонка периода должны отличаться";
    }

    if (getSingleFileVirtualPeriods().length < 2) {
      return "В файле должно быть минимум два периода";
    }

    return "";
  }

  function renderSingleFileMetricRow(metric) {
    const hasColumn = Boolean(getSingleFileMetricColumn(metric));

    return (
      '<div class="metric-row' +
      (hasColumn ? "" : " metric-row--invalid") +
      '" data-metric-id="' +
      metric.id +
      '">' +
      '<div class="metric-name"><span>Показатель</span><strong>' +
      App.UI.escapeHtml(metric.label || "Выберите столбец") +
      "</strong></div>" +
      '<label class="field"><span>Колонка в файле</span><select name="singleMetricColumn" data-metric-id="' +
      metric.id +
      '">' +
      buildColumnOptions(ensureSingleFileSource().table, getSingleFileMetricColumn(metric), "Выберите показатель") +
      "</select></label>" +
      '<button class="icon-button" type="button" data-action="remove-metric" title="Удалить показатель">×</button>' +
      "</div>"
    );
  }

  function renderMetricRow(metric) {
    const hasMismatch = Boolean(getMetricSelectionIssue(metric, 0));

    return (
      '<div class="metric-row metric-row--multi' +
      (hasMismatch ? " metric-row--invalid" : "") +
      '" data-metric-id="' +
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

    if (isSingleFileSourceMode()) {
      columns[SINGLE_FILE_METRIC_COLUMN_KEY] = getDefaultSingleFileMetricColumn();
    } else {
      state.periods.forEach(function (period) {
        columns[period.id] = getDefaultMetricColumn(period);
      });
    }

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

  function getDefaultSingleFileMetricColumn() {
    const singleFile = ensureSingleFileSource();

    if (!singleFile.table) {
      return "";
    }

    const header = singleFile.table.headers.find(function (item) {
      return item.id !== singleFile.idColumn && item.id !== singleFile.periodColumn;
    });

    return header ? header.id : "";
  }

  function handleMetricInput(event) {
    if (event.target.name === "singleMetricColumn") {
      const row = event.target.closest("[data-metric-id]");
      const metric = row ? findMetric(row.dataset.metricId) : null;

      if (!metric) {
        return;
      }

      metric.columns[SINGLE_FILE_METRIC_COLUMN_KEY] = event.target.value;
      syncMetricLabels();
      clearAnalysis();
      renderMetrics();
      renderAnalysis();
      renderWarningsPanel();
      return;
    }

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
    renderWarningsPanel();
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

  function getSingleFileMetricColumn(metric) {
    return metric && metric.columns ? metric.columns[SINGLE_FILE_METRIC_COLUMN_KEY] || "" : "";
  }

  function getComparisonPeriods() {
    if (!isSingleFileSourceMode()) {
      return state.periods;
    }

    const singleFile = ensureSingleFileSource();

    return getSingleFileVirtualPeriods().map(function (period) {
      return Object.assign({}, period, {
        idColumn: singleFile.idColumn,
      });
    });
  }

  function getComparisonMetrics(periods) {
    if (!isSingleFileSourceMode()) {
      return state.mapping.metrics;
    }

    const comparisonPeriods = periods || getComparisonPeriods();

    return state.mapping.metrics.map(function (metric) {
      const columnId = getSingleFileMetricColumn(metric);
      const columns = {};

      comparisonPeriods.forEach(function (period) {
        columns[period.id] = columnId;
      });

      return Object.assign({}, metric, {
        label: getSingleFileMetricLabel(metric) || metric.label,
        columns: columns,
      });
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

  function isSingleFileReadyForMetrics() {
    const singleFile = ensureSingleFileSource();

    return Boolean(
      singleFile.table &&
        singleFile.idColumn &&
        singleFile.periodColumn &&
        singleFile.idColumn !== singleFile.periodColumn &&
        getSingleFileVirtualPeriods().length >= 2
    );
  }

  function areSingleFileMetricsReady() {
    return (
      isSingleFileReadyForMetrics() &&
      state.mapping.metrics.length > 0 &&
      state.mapping.metrics.every(function (metric) {
        return Boolean(getSingleFileMetricColumn(metric));
      })
    );
  }

  function isReadyToAnalyze() {
    if (isSingleFileSourceMode()) {
      return areSingleFileMetricsReady();
    }

    return Boolean(
      areIdsReady() &&
        state.mapping.metrics.length &&
        state.mapping.metrics.every(function (metric) {
          return state.periods.every(function (period) {
            return Boolean(metric.columns[period.id]);
          });
        }) &&
        !getMetricSelectionIssues().length
    );
  }

  function analyze() {
    state.messages = [];
    syncMetricLabels();

    if (!isReadyToAnalyze()) {
      const metricIssues = getMetricSelectionIssues();
      if (!metricIssues.length) {
        state.messages.push({
          type: "error",
          message: isSingleFileSourceMode()
            ? "Для расчета выберите объект, колонку периода и хотя бы один показатель в загруженном файле"
            : "Для расчета нужны минимум два загруженных периода, колонки с объектами сравнения и хотя бы один показатель",
        });
      }
      renderAll();
      return;
    }

    const comparisonPeriods = getComparisonPeriods();
    const comparisonMetrics = getComparisonMetrics(comparisonPeriods);

    state.comparison = App.Comparator.comparePeriods({
      periods: comparisonPeriods,
      metrics: comparisonMetrics,
      comparisonMode: state.comparisonMode,
    });

    state.analytics = App.Analytics.buildAnalytics(state.comparison, state.mapping.metrics);
    state.selectedChartMetricId = state.mapping.metrics[0].id;
    state.restoredHistoryMeta = null;
    state.restoredHistorySettings = null;
    hasUnsavedAnalysis = true;
    renderAll();
  }

  function renderAnalysis() {
    const view = buildFilteredAnalysisView();

    renderGlobalFilters(view);
    App.UI.renderSummary(dom.summaryCards, view.analytics);
    App.UI.renderChartMetricSelect(dom.chartMetricSelect, state.mapping.metrics, state.selectedChartMetricId);
    dom.chartMetricSelect.disabled = !state.comparison || !state.mapping.metrics.length;
    App.UI.renderMovers(dom.moversPanel, view.analytics);
    App.UI.renderResultsTable(dom.resultsTable, view.comparison, state.mapping.metrics);
    renderChart(view.comparison);

    const hasComparison = Boolean(state.comparison);
    dom.exportCsvButton.disabled = !hasComparison;
    dom.exportExcelButton.disabled = !hasComparison;
    dom.saveAnalysisButton.disabled = !hasComparison;
    refreshMotion();
  }

  function renderHistoryPanel() {
    const filteredRecords = filterHistoryRecords(historyRecords, historySearchQuery);
    dom.historyTitle.textContent = "История анализов (" + historyRecords.length + ")";
    dom.historySearchInput.disabled = !historyRecords.length;
    App.UI.renderHistory(dom.historyList, filteredRecords, {
      hasRecords: Boolean(historyRecords.length),
      query: historySearchQuery,
    });
    dom.clearHistoryButton.disabled = !historyRecords.length;
  }

  function filterHistoryRecords(records, query) {
    if (!query) {
      return records;
    }

    return records.filter(function (record) {
      return String(record.title || "").toLowerCase().includes(query);
    });
  }

  function openSaveAnalysisModal() {
    if (!state.comparison || !App.HistoryStore) {
      return;
    }

    const createdAt = new Date().toISOString();
    dom.saveAnalysisForm.dataset.mode = "save";
    dom.saveAnalysisForm.dataset.historyId = "";
    dom.saveAnalysisTitle.textContent = "Сохранить анализ";
    dom.saveAnalysisDescription.textContent = "Название поможет быстрее найти результат в истории.";
    dom.saveAnalysisSubmitButton.textContent = "Сохранить";
    dom.saveAnalysisNameInput.value = App.HistoryStore.defaultTitle(createdAt);
    dom.saveAnalysisNameInput.dataset.createdAt = createdAt;
    dom.saveAnalysisModal.hidden = false;
    dom.saveAnalysisNameInput.focus();
    dom.saveAnalysisNameInput.select();
  }

  function closeSaveAnalysisModal() {
    dom.saveAnalysisModal.hidden = true;
    dom.saveAnalysisForm.dataset.mode = "";
    dom.saveAnalysisForm.dataset.historyId = "";
    dom.saveAnalysisNameInput.value = "";
    dom.saveAnalysisNameInput.dataset.createdAt = "";
  }

  function saveAnalysisFromModal(event) {
    event.preventDefault();

    if (dom.saveAnalysisForm.dataset.mode === "rename") {
      renameHistoryFromModal();
      return;
    }

    if (!state.comparison || !App.HistoryStore) {
      closeSaveAnalysisModal();
      return;
    }

    const createdAt = dom.saveAnalysisNameInput.dataset.createdAt || new Date().toISOString();
    const title = dom.saveAnalysisNameInput.value.trim() || App.HistoryStore.defaultTitle(createdAt);
    const result = App.HistoryStore.save(buildHistoryRecord(title, createdAt));

    historyRecords = result.records;
    if (result.ok) {
      historySearchQuery = "";
      dom.historySearchInput.value = "";
    }
    renderHistoryPanel();
    refreshMotion();

    if (!result.ok) {
      state.messages.push({
        type: "error",
        message: "Не удалось сохранить анализ. Возможно, локальное хранилище переполнено.",
      });
      renderWarningsPanel();
      return;
    }

    closeSaveAnalysisModal();
    hasUnsavedAnalysis = false;
  }

  function buildHistoryRecord(title, createdAt) {
    const metrics = cloneJson(state.mapping.metrics);
    const comparison = cloneJson(state.comparison);
    const analytics = cloneJson(state.analytics);
    const idColumns = getIdColumnMetadata();
    const restoredMeta = state.restoredHistoryMeta || {};
    const restoredSettings = state.restoredHistorySettings || {};
    const identifierLabel = getIdentifierLabel(idColumns) || restoredMeta.identifierLabel || "";

    return {
      createdAt: createdAt,
      title: title,
      pinned: false,
      meta: {
        periodSourceMode: state.periodSourceMode || "multiFile",
        comparisonMode: state.comparisonMode,
        periodCount: comparison && comparison.periods ? comparison.periods.length : state.periods.length,
        metricCount: metrics.length,
        totalUnits: analytics ? analytics.totalUnits : 0,
        totalCompared: analytics ? analytics.totalCompared : 0,
        rowCount: comparison && comparison.rows ? comparison.rows.length : 0,
        changedCount: getChangedCount(analytics),
        identifierLabel: identifierLabel,
      },
      metrics: metrics,
      settings: {
        periodSourceMode: state.periodSourceMode || "multiFile",
        comparisonMode: state.comparisonMode,
        selectedChartMetricId: state.selectedChartMetricId,
        singleFile:
          state.periodSourceMode === "singleFile"
            ? {
                periodColumn: ensureSingleFileSource().periodColumn || "",
                periodColumnName: getTableColumnName(ensureSingleFileSource().table, ensureSingleFileSource().periodColumn) || "",
                idColumn: ensureSingleFileSource().idColumn || "",
                idColumnName: getTableColumnName(ensureSingleFileSource().table, ensureSingleFileSource().idColumn) || "",
              }
            : null,
        idColumns: idColumns.length ? idColumns : restoredSettings.idColumns || [],
      },
      comparison: comparison,
      analytics: analytics,
    };
  }

  function getIdColumnMetadata() {
    if (isSingleFileSourceMode()) {
      const singleFile = ensureSingleFileSource();
      const periods = state.comparison && Array.isArray(state.comparison.periods) ? state.comparison.periods : getSingleFileVirtualPeriods();
      const columnName = getTableColumnName(singleFile.table, singleFile.idColumn) || "";

      if (!singleFile.idColumn) {
        return [];
      }

      return periods.map(function (period) {
        return {
          periodId: period.id,
          periodLabel: period.label,
          columnId: singleFile.idColumn,
          columnName: columnName,
        };
      });
    }

    return state.periods.map(function (period) {
      return {
        periodId: period.id,
        periodLabel: period.label,
        columnId: period.idColumn || "",
        columnName: getColumnName(period, period.idColumn) || "",
      };
    }).filter(function (item) {
      return item.columnId || item.columnName;
    });
  }

  function getIdentifierLabel(idColumns) {
    const names = idColumns
      .map(function (item) {
        return item.columnName || item.columnId;
      })
      .filter(Boolean);
    const uniqueNames = Array.from(new Set(names));

    return uniqueNames.join(", ");
  }

  function handleHistoryClick(event) {
    const actionButton = event.target.closest("[data-action]");

    if (!actionButton) {
      return;
    }

    const id = actionButton.dataset.historyId;

    if (actionButton.dataset.action === "open-history") {
      openHistoryRecord(id);
      return;
    }

    if (actionButton.dataset.action === "toggle-history-pin") {
      toggleHistoryPin(id);
      return;
    }

    if (actionButton.dataset.action === "rename-history") {
      openRenameAnalysisModal(id);
      return;
    }

    if (actionButton.dataset.action === "delete-history") {
      deleteHistoryRecord(id);
    }
  }

  function openRenameAnalysisModal(id) {
    const record = historyRecords.find(function (item) {
      return item.id === id;
    });

    if (!record || !App.HistoryStore) {
      return;
    }

    dom.saveAnalysisForm.dataset.mode = "rename";
    dom.saveAnalysisForm.dataset.historyId = id;
    dom.saveAnalysisTitle.textContent = "Переименовать анализ";
    dom.saveAnalysisDescription.textContent = "Дата создания и результаты останутся без изменений.";
    dom.saveAnalysisSubmitButton.textContent = "Сохранить";
    dom.saveAnalysisNameInput.value = record.title;
    dom.saveAnalysisNameInput.dataset.createdAt = record.createdAt || "";
    dom.saveAnalysisModal.hidden = false;
    dom.saveAnalysisNameInput.focus();
    dom.saveAnalysisNameInput.select();
  }

  function renameHistoryFromModal() {
    const id = dom.saveAnalysisForm.dataset.historyId;
    const record = historyRecords.find(function (item) {
      return item.id === id;
    });

    if (!record || !App.HistoryStore) {
      closeSaveAnalysisModal();
      return;
    }

    const fallbackTitle = App.HistoryStore.defaultTitle(record.createdAt);
    const title = dom.saveAnalysisNameInput.value.trim() || fallbackTitle;
    const result = App.HistoryStore.save(Object.assign({}, record, { title: title }));

    historyRecords = result.records;
    renderHistoryPanel();
    refreshMotion();

    if (!result.ok) {
      state.messages.push({
        type: "error",
        message: "Не удалось переименовать анализ. Возможно, локальное хранилище переполнено.",
      });
      renderWarningsPanel();
      return;
    }

    closeSaveAnalysisModal();
  }

  function openHistoryRecord(id) {
    const record = historyRecords.find(function (item) {
      return item.id === id;
    });

    if (!record || !record.comparison || !record.analytics) {
      return;
    }

    state.messages = [];
    state.periodSourceMode = record.settings.periodSourceMode || record.meta.periodSourceMode || "multiFile";
    state.comparisonMode = record.settings.comparisonMode || record.meta.comparisonMode || record.comparison.comparisonMode || "endpoint";
    state.comparison = cloneJson(record.comparison);
    state.analytics = cloneJson(record.analytics);
    state.mapping.metrics = cloneJson(record.metrics || []);
    state.selectedChartMetricId =
      record.settings.selectedChartMetricId ||
      (state.mapping.metrics[0] ? state.mapping.metrics[0].id : "");
    resetGlobalFiltersState();
    state.periods = restorePeriods(record);
    state.singleFile = App.createSingleFileSource ? App.createSingleFileSource() : null;
    state.restoredHistoryMeta = cloneJson(record.meta || {});
    state.restoredHistorySettings = cloneJson(record.settings || {});
    hasUnsavedAnalysis = false;
    renderAll();
  }

  function restorePeriods(record) {
    const periods = record.comparison && Array.isArray(record.comparison.periods) ? record.comparison.periods : [];

    if (periods.length >= 2) {
      return periods.map(function (period) {
        return {
          id: period.id,
          label: period.label,
          file: null,
          table: null,
          idColumn: getRestoredIdColumn(record, period.id),
        };
      });
    }

    return [App.createPeriod(0), App.createPeriod(1)];
  }

  function getRestoredIdColumn(record, periodId) {
    const idColumns = record.settings && Array.isArray(record.settings.idColumns) ? record.settings.idColumns : [];
    const item = idColumns.find(function (column) {
      return column.periodId === periodId;
    });

    return item ? item.columnId : "";
  }

  function toggleHistoryPin(id) {
    if (!App.HistoryStore) {
      return;
    }

    const result = App.HistoryStore.togglePinned(id);
    historyRecords = result.records;
    renderHistoryPanel();
    refreshMotion();
  }

  function deleteHistoryRecord(id) {
    if (!App.HistoryStore || !window.confirm("Удалить сохраненный анализ?")) {
      return;
    }

    animateHistoryRemoval(id, function () {
      const result = App.HistoryStore.remove(id);
      historyRecords = result.records;
      renderHistoryPanel();
      refreshMotion();
    });
  }

  function clearHistory() {
    if (!historyRecords.length || !App.HistoryStore) {
      return;
    }

    if (!window.confirm("Вы действительно хотите удалить всю историю анализов?")) {
      return;
    }

    const result = App.HistoryStore.clear();
    historyRecords = result.records;
    renderHistoryPanel();
    refreshMotion();
  }

  function cloneJson(value) {
    if (value === null || value === undefined) {
      return value;
    }

    return JSON.parse(JSON.stringify(value));
  }

  function getChangedCount(analytics) {
    if (!analytics || !Array.isArray(analytics.metricSummaries)) {
      return 0;
    }

    return analytics.metricSummaries.reduce(function (sum, summary) {
      return sum + (Number(summary.improvedCount) || 0) + (Number(summary.declinedCount) || 0);
    }, 0);
  }

  function animateHistoryRemoval(id, callback) {
    const card = Array.from(dom.historyList.querySelectorAll("[data-history-id]")).find(function (item) {
      return item.dataset.historyId === id;
    });

    if (!card) {
      callback();
      return;
    }

    card.classList.add("is-removing");
    window.setTimeout(callback, 190);
  }

  function handleBeforeUnload(event) {
    if (!hasUnsavedAnalysis || !state.comparison) {
      return;
    }

    event.preventDefault();
    event.returnValue = "";

    window.setTimeout(function () {
      if (hasUnsavedAnalysis && state.comparison && dom.saveAnalysisModal.hidden) {
        openSaveAnalysisModal();
      }
    }, 0);

    return "";
  }

  function initMotion() {
    if (App.Motion) {
      App.Motion.init();
    }
  }

  function refreshMotion() {
    if (!App.Motion) {
      return;
    }

    App.Motion.refresh();
  }

  function renderWarningsPanel() {
    dom.warningsPanel.dataset.hasFiles = isSingleFileSourceMode() || hasAnyLoadedPeriod() ? "true" : "false";
    App.UI.renderWarnings(dom.warningsPanel, collectWarnings());
  }

  function isSingleFileSourceMode() {
    return state.periodSourceMode === "singleFile";
  }

  function ensureSingleFileSource() {
    if (!state.singleFile) {
      state.singleFile = App.createSingleFileSource ? App.createSingleFileSource() : {
        file: null,
        table: null,
        periodColumn: "",
        idColumn: "",
        virtualPeriods: [],
        warnings: [],
        loading: false,
      };
    }

    return state.singleFile;
  }

  function rebuildSingleFileVirtualPeriods() {
    const singleFile = ensureSingleFileSource();

    if (!singleFile.table || !singleFile.periodColumn || !App.PeriodBuilder) {
      singleFile.virtualPeriods = [];
      singleFile.warnings = [];
      return [];
    }

    const result = App.PeriodBuilder.buildVirtualPeriods({
      table: singleFile.table,
      periodColumn: singleFile.periodColumn,
    });

    singleFile.virtualPeriods = result.periods;
    singleFile.warnings = result.warnings;

    return singleFile.virtualPeriods;
  }

  function getSingleFileVirtualPeriods() {
    const singleFile = ensureSingleFileSource();

    if (!singleFile.table || !singleFile.periodColumn) {
      return [];
    }

    if (!Array.isArray(singleFile.virtualPeriods) || !singleFile.virtualPeriods.length) {
      return rebuildSingleFileVirtualPeriods();
    }

    return singleFile.virtualPeriods;
  }

  function getSingleFilePeriodValues() {
    return getSingleFileVirtualPeriods().map(function (period) {
      return period.label;
    });
  }

  function hasAnyLoadedPeriod() {
    return state.periods.some(function (period) {
      return Boolean(period.table);
    });
  }

  function handleGlobalFilterChange() {
    ensureGlobalFilters();
    state.globalFilters.impact = dom.globalImpactFilter.value;
    state.globalFilters.deltaMin = dom.globalDeltaMinFilter.value;
    state.globalFilters.deltaMax = dom.globalDeltaMaxFilter.value;
    state.globalFilters.objectQuery = dom.globalObjectSearch.value;
    state.globalFilters.departmentQuery = dom.globalDepartmentSearch.value;
    renderAnalysis();
  }

  function resetGlobalFilters() {
    resetGlobalFiltersState();
    renderAnalysis();
  }

  function resetGlobalFiltersState() {
    state.globalFilters = App.createGlobalFilters ? App.createGlobalFilters() : {
      impact: "all",
      deltaMin: "",
      deltaMax: "",
      objectQuery: "",
      departmentQuery: "",
    };
  }

  function ensureGlobalFilters() {
    if (!state.globalFilters) {
      resetGlobalFiltersState();
    }
  }

  function buildFilteredAnalysisView() {
    ensureGlobalFilters();

    if (!state.comparison) {
      return {
        comparison: null,
        analytics: null,
        totalRows: 0,
        visibleRows: 0,
      };
    }

    const filteredComparison = buildFilteredComparison(state.comparison);
    const analytics = App.Analytics.buildAnalytics(filteredComparison, state.mapping.metrics);

    return {
      comparison: filteredComparison,
      analytics: analytics,
      totalRows: state.comparison.rows.length,
      visibleRows: filteredComparison.rows.length,
    };
  }

  function buildFilteredComparison(comparison) {
    if (!hasActiveGlobalFilters()) {
      return comparison;
    }

    return Object.assign({}, comparison, {
      rows: comparison.rows.filter(rowMatchesGlobalFilters),
    });
  }

  function rowMatchesGlobalFilters(row) {
    const filters = state.globalFilters;
    const objectQuery = normalizeSearch(filters.objectQuery);
    const departmentQuery = normalizeSearch(filters.departmentQuery);

    if (objectQuery && !matchesText([row.label, row.key], objectQuery)) {
      return false;
    }

    if (departmentQuery && !matchesText(getRowSearchValues(row), departmentQuery)) {
      return false;
    }

    if (!hasMetricGlobalFilters()) {
      return true;
    }

    const comparisons = getRowFilterComparisons(row);

    if (!comparisons.length) {
      return false;
    }

    return comparisons.some(matchGlobalMetricFilters);
  }

  function matchGlobalMetricFilters(item) {
    const filters = state.globalFilters;
    const min = parseFilterNumber(filters.deltaMin);
    const max = parseFilterNumber(filters.deltaMax);

    if (filters.impact === "good" && item.impact !== "good") {
      return false;
    }

    if (filters.impact === "bad" && item.impact !== "bad") {
      return false;
    }

    if (Number.isFinite(min) && item.delta < min) {
      return false;
    }

    if (Number.isFinite(max) && item.delta > max) {
      return false;
    }

    return true;
  }

  function getRowFilterComparisons(row) {
    const metric = findMetric(state.selectedChartMetricId) || state.mapping.metrics[0];
    const metricResults = metric
      ? row.metrics.filter(function (item) {
          return item.metricId === metric.id;
        })
      : row.metrics;

    return metricResults
      .flatMap(function (result) {
        return result && Array.isArray(result.comparisons) ? result.comparisons : [];
      })
      .filter(function (item) {
        return Number.isFinite(item.delta);
      });
  }

  function hasMetricGlobalFilters() {
    const filters = state.globalFilters;

    return (
      filters.impact !== "all" ||
      Number.isFinite(parseFilterNumber(filters.deltaMin)) ||
      Number.isFinite(parseFilterNumber(filters.deltaMax))
    );
  }

  function hasActiveGlobalFilters() {
    const filters = state.globalFilters;

    return Boolean(
      filters &&
        (filters.impact !== "all" ||
          String(filters.deltaMin || "").trim() ||
          String(filters.deltaMax || "").trim() ||
          String(filters.objectQuery || "").trim() ||
          String(filters.departmentQuery || "").trim())
    );
  }

  function renderGlobalFilters(view) {
    ensureGlobalFilters();

    const hasComparison = Boolean(state.comparison);
    const hasActiveFilters = hasActiveGlobalFilters();
    const controls = [
      dom.globalImpactFilter,
      dom.globalDeltaMinFilter,
      dom.globalDeltaMaxFilter,
      dom.globalObjectSearch,
      dom.globalDepartmentSearch,
    ];

    dom.globalImpactFilter.value = state.globalFilters.impact;
    dom.globalDeltaMinFilter.value = state.globalFilters.deltaMin;
    dom.globalDeltaMaxFilter.value = state.globalFilters.deltaMax;
    dom.globalObjectSearch.value = state.globalFilters.objectQuery;
    dom.globalDepartmentSearch.value = state.globalFilters.departmentQuery;

    controls.forEach(function (control) {
      control.disabled = !hasComparison;
    });

    dom.resetGlobalFiltersButton.disabled = !hasComparison || !hasActiveFilters;

    if (!hasComparison) {
      dom.globalFilterStatus.textContent = "Фильтры появятся после расчета";
      return;
    }

    if (!hasActiveFilters) {
      dom.globalFilterStatus.textContent = "Показаны все строки: " + view.totalRows;
      return;
    }

    dom.globalFilterStatus.textContent = view.visibleRows
      ? "Показано " + view.visibleRows + " из " + view.totalRows
      : "Ничего не найдено";
  }

  function getRowSearchValues(row) {
    const values = [row.label, row.key];

    if (!Array.isArray(row.records)) {
      return values;
    }

    row.records.forEach(function (record) {
      if (!record || !record.row || !record.row.values) {
        return;
      }

      Object.keys(record.row.values).forEach(function (key) {
        values.push(record.row.values[key]);
      });
    });

    return values;
  }

  function normalizeSearch(value) {
    return String(value || "").trim().toLowerCase();
  }

  function matchesText(values, query) {
    return values.some(function (value) {
      return String(value || "").toLowerCase().includes(query);
    });
  }

  function parseFilterNumber(value) {
    const text = String(value || "").replace(",", ".").trim();

    if (!text) {
      return NaN;
    }

    return Number(text);
  }

  function renderChart(comparison) {
    const metric = findMetric(state.selectedChartMetricId) || state.mapping.metrics[0];
    App.Charts.renderDeltaChart(dom.deltaChart, comparison || state.comparison, metric);
  }

  function syncMetricLabels() {
    state.mapping.metrics.forEach(function (metric) {
      metric.label = getMetricLabel(metric);
    });
  }

  function getMetricLabel(metric) {
    if (isSingleFileSourceMode()) {
      return getSingleFileMetricLabel(metric);
    }

    const lastPeriod = state.periods[state.periods.length - 1];
    const firstPeriod = state.periods[0];

    return (
      getColumnName(lastPeriod, metric.columns[lastPeriod.id]) ||
      getColumnName(firstPeriod, metric.columns[firstPeriod.id]) ||
      ""
    );
  }

  function getSingleFileMetricLabel(metric) {
    return getTableColumnName(ensureSingleFileSource().table, getSingleFileMetricColumn(metric));
  }

  function getColumnName(period, columnId) {
    return getTableColumnName(period ? period.table : null, columnId);
  }

  function getTableColumnName(table, columnId) {
    if (!table || !columnId) {
      return "";
    }

    const header = table.headers.find(function (item) {
      return item.id === columnId;
    });

    return header ? header.name : "";
  }

  function collectWarnings() {
    const warnings = state.messages.slice();

    if (isSingleFileSourceMode()) {
      const singleFile = ensureSingleFileSource();

      if (singleFile.table) {
        warnings.push.apply(warnings, singleFile.table.warnings);
      }

      if (!singleFile.table) {
        warnings.push({
          type: "neutral",
          message: "Загрузите один файл, чтобы выбрать объект и колонку периода.",
        });
      } else if (!singleFile.idColumn) {
        warnings.push({
          type: "neutral",
          message: "Файл загружен. Выберите колонку с объектом сравнения: сотрудником, командой, логином или другим ключом.",
        });
      } else if (!singleFile.periodColumn) {
        warnings.push({
          type: "neutral",
          message: "Файл загружен. Выберите колонку, в которой указаны даты, месяцы, недели или годы.",
        });
      } else if (singleFile.idColumn === singleFile.periodColumn) {
        warnings.push({
          type: "error",
          message: "Колонка объекта и колонка периода должны отличаться.",
        });
      } else {
        const periodValues = getSingleFilePeriodValues();
        warnings.push.apply(warnings, singleFile.warnings || []);
        warnings.push({
          type: periodValues.length >= 2 ? "neutral" : "warn",
          message:
            periodValues.length >= 2
              ? "Колонка периода выбрана. Найдено периодов: " + periodValues.length + "."
              : "В выбранной колонке периода найдено меньше двух уникальных значений.",
        });

        if (periodValues.length >= 2 && !state.mapping.metrics.length) {
          warnings.push({
            type: "neutral",
            message: "Добавьте показатель для сравнения виртуальных периодов.",
          });
        }
      }

      if (state.comparison) {
        appendComparisonWarnings(warnings);
      }

      return warnings;
    }

    state.periods.forEach(function (period) {
      if (period.table) {
        warnings.push.apply(warnings, period.table.warnings);
      }
    });

    warnings.push.apply(warnings, compareHeadersWarning());
    warnings.push.apply(warnings, getMetricSelectionWarnings());

    if (state.comparison) {
      appendComparisonWarnings(warnings);
    }

    return warnings;
  }

  function getMetricSelectionWarnings() {
    return getMetricSelectionIssues().map(function (issue) {
      return {
        type: "error",
        message: formatMetricSelectionIssue(issue),
      };
    });
  }

  function getMetricSelectionIssues() {
    if (isSingleFileSourceMode()) {
      return [];
    }

    if (!areIdsReady() || !state.mapping.metrics.length) {
      return [];
    }

    return state.mapping.metrics
      .map(function (metric, index) {
        return getMetricSelectionIssue(metric, index);
      })
      .filter(Boolean);
  }

  function getMetricSelectionIssue(metric, index) {
    if (!areAllPeriodsLoaded()) {
      return null;
    }

    const selections = state.periods.map(function (period) {
      const columnName = getColumnName(period, metric.columns[period.id]);

      return {
        periodLabel: period.label,
        columnName: columnName,
        normalizedName: normalizeMetricColumnName(columnName),
      };
    });

    if (
      selections.some(function (item) {
        return !item.normalizedName;
      })
    ) {
      return null;
    }

    const uniqueNames = new Set(
      selections.map(function (item) {
        return item.normalizedName;
      })
    );

    if (uniqueNames.size <= 1) {
      return null;
    }

    return {
      index: index,
      metric: metric,
      selections: selections,
    };
  }

  function normalizeMetricColumnName(value) {
    return App.Normalizers.normalizeKey(value)
      .replace(/[%№#]/g, " ")
      .replace(/\b(процент|проц|percent|pct)\b/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function formatMetricSelectionIssue(issue) {
    const details = issue.selections
      .map(function (item) {
        return "«" + item.periodLabel + "» — «" + item.columnName + "»";
      })
      .join("; ");

    return (
      "Показатель " +
      (issue.index + 1) +
      " выбран из разных столбцов: " +
      details +
      ". Выберите одинаковый показатель во всех периодах."
    );
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
          message: "В периоде «" + item.periodLabel + "» есть строки без объекта сравнения: " + item.items.length,
        });
      }
    });

    comparison.duplicatesByPeriod.forEach(function (item) {
      if (item.items.length) {
        warnings.push({
          type: "warn",
          message: "В периоде «" + item.periodLabel + "» есть дубли объектов сравнения: " + item.items.length,
        });
      }
    });

    comparison.missingByPeriod.forEach(function (item) {
      if (item.items.length) {
        warnings.push({
          type: "warn",
          message: "В периоде «" + item.periodLabel + "» отсутствуют объекты сравнения: " + item.items.length,
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
      await App.Exporters.exportExcel(state.comparison, state.mapping.metrics, state.analytics, {
        chartMetric: findMetric(state.selectedChartMetricId) || state.mapping.metrics[0],
      });
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
