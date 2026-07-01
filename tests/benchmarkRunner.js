(function (global) {
  const App = global.Metricum || {};
  const BENCHMARK_PERIOD_HEADER = "Период";
  const BENCHMARK_ID_HEADER = "Сотрудник";
  const BASE_HEADERS = [
    { id: "period", name: BENCHMARK_PERIOD_HEADER },
    { id: "employee", name: BENCHMARK_ID_HEADER },
    { id: "department", name: "Подразделение" },
    { id: "sales", name: "Продажи" },
    { id: "quality", name: "Качество %" },
    { id: "errors", name: "Ошибки" },
    { id: "aht", name: "AHT" },
  ];
  const dom = {};
  let isRunning = false;
  let hasResults = false;

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    dom.columnCountSelect = document.getElementById("columnCountSelect");
    dom.metricCountSelect = document.getElementById("metricCountSelect");
    dom.comparisonModeSelect = document.getElementById("comparisonModeSelect");
    dom.repeatCountSelect = document.getElementById("repeatCountSelect");
    dom.status = document.getElementById("benchmarkStatus");
    dom.results = document.getElementById("benchmarkResults");
    dom.runAllButton = document.getElementById("runAllButton");
    dom.clearButton = document.getElementById("clearButton");
    dom.rowButtons = Array.from(document.querySelectorAll("[data-rows]"));

    dom.rowButtons.forEach(function (button) {
      button.addEventListener("click", function () {
        runScenario(Number(button.dataset.rows));
      });
    });

    dom.runAllButton.addEventListener("click", function () {
      runAllScenarios();
    });

    dom.clearButton.addEventListener("click", clearResults);
  }

  async function runAllScenarios() {
    if (isRunning) {
      return;
    }

    for (const rowCount of [10000, 50000, 100000]) {
      await runScenario(rowCount);
    }
  }

  async function runScenario(rowCount) {
    if (isRunning) {
      return;
    }

    const options = getOptions(rowCount);
    const results = [];
    isRunning = true;
    setControlsDisabled(true);

    try {
      for (let index = 0; index < options.repeatCount; index += 1) {
        await setStatus(
          "Benchmark " + formatNumber(options.rowCount) + " строк",
          "Прогон " + (index + 1) + " из " + options.repeatCount
        );
        results.push(await runOnce(options));
      }

      appendResult(averageResults(results, options));
      await setStatus(
        "Benchmark завершен",
        "Последний сценарий: " + formatNumber(options.rowCount) + " строк, " + options.columnCount + " колонок.",
        false
      );
    } catch (error) {
      await setStatus("Benchmark остановлен", error.message || String(error), false);
    } finally {
      isRunning = false;
      setControlsDisabled(false);
    }
  }

  async function runOnce(options) {
    const startTotal = now();

    await setStatus("Генерация CSV", formatScenario(options));
    const startGenerate = now();
    const csv = await generateCsv(options.rowCount, options.columnCount);
    const generateMs = elapsed(startGenerate);
    const file = createCsvFile(csv, options);
    const fileSize = file.size || byteLength(csv);

    await setStatus("Чтение файла", "Передаем CSV в текущий ExcelReader");
    const startRead = now();
    const table = await App.ExcelReader.readExcelFile(file, {
      onProgress: function (status) {
        return setStatus(status.title, status.detail);
      },
    });
    const readMs = elapsed(startRead);
    const columnIds = getBenchmarkColumnIds(table);

    await setStatus("Сбор периодов", "Группируем строки по колонке периода");
    const startPeriods = now();
    const periodResult = App.PeriodBuilder.buildVirtualPeriods({
      table: table,
      periodColumn: columnIds.period,
    });
    const periods = periodResult.periods.map(function (period) {
      return Object.assign({}, period, {
        idColumn: columnIds.employee,
      });
    });
    const periodMs = elapsed(startPeriods);

    await setStatus("Сравнение", "Сопоставляем объекты и показатели");
    const metrics = buildMetrics(periods, options.metricCount, columnIds);
    const startCompare = now();
    const comparison = App.Comparator.comparePeriods({
      periods: periods,
      metrics: metrics,
      comparisonMode: options.comparisonMode,
    });
    const compareMs = elapsed(startCompare);

    await setStatus("Аналитика", "Готовим dashboard-агрегаты");
    const startAnalytics = now();
    const analytics = App.Analytics.buildAnalytics(comparison, metrics);
    const analyticsMs = elapsed(startAnalytics);

    return {
      rowCount: options.rowCount,
      columnCount: options.columnCount,
      cellCount: table.workload ? table.workload.cellCount : table.rows.length * table.headers.length,
      fileSize: fileSize,
      generateMs: generateMs,
      readMs: readMs,
      periodMs: periodMs,
      compareMs: compareMs,
      analyticsMs: analyticsMs,
      totalMs: elapsed(startTotal),
      objectCount: comparison.rows.length,
      warningCount:
        (table.warnings ? table.warnings.length : 0) +
        (periodResult.warnings ? periodResult.warnings.length : 0) +
        comparison.missingByPeriod.reduce(function (sum, group) {
          return sum + group.items.length;
        }, 0) +
        comparison.duplicatesByPeriod.reduce(function (sum, group) {
          return sum + group.items.length;
        }, 0) +
        comparison.invalidValues.length,
      analytics: analytics,
    };
  }

  function getOptions(rowCount) {
    return {
      rowCount: rowCount,
      columnCount: Number(dom.columnCountSelect.value) || 100,
      metricCount: Number(dom.metricCountSelect.value) || 3,
      comparisonMode: dom.comparisonModeSelect.value || "endpoint",
      repeatCount: Number(dom.repeatCountSelect.value) || 1,
    };
  }

  async function generateCsv(rowCount, columnCount) {
    const headers = buildHeaders(columnCount);
    const chunks = [headers.map(function (header) { return header.name; }).join(",") + "\n"];
    const chunkSize = 1000;

    for (let start = 0; start < rowCount; start += chunkSize) {
      const end = Math.min(start + chunkSize, rowCount);
      const lines = [];

      for (let index = start; index < end; index += 1) {
        lines.push(buildCsvRow(index, headers).join(","));
      }

      chunks.push(lines.join("\n") + "\n");
      await yieldToBrowser();
    }

    return chunks.join("");
  }

  function buildHeaders(columnCount) {
    const headers = BASE_HEADERS.slice(0, Math.max(0, Math.min(BASE_HEADERS.length, columnCount)));

    while (headers.length < columnCount) {
      const number = headers.length - BASE_HEADERS.length + 1;
      headers.push({
        id: "extra_" + number,
        name: "Доп поле " + number,
      });
    }

    return headers;
  }

  function buildCsvRow(index, headers) {
    const objectIndex = Math.floor(index / 2) + 1;
    const isCurrentPeriod = index % 2 === 1;
    const departmentNumber = (objectIndex % 12) + 1;
    const salesBase = 80 + (objectIndex % 170);
    const qualityBase = 78 + (objectIndex % 18);
    const errorsBase = objectIndex % 9;
    const ahtBase = 210 + (objectIndex % 80);
    const baseValues = {
      period: isCurrentPeriod ? "Февраль" : "Январь",
      employee: "Сотрудник " + objectIndex,
      department: "Отдел " + departmentNumber,
      sales: salesBase + (isCurrentPeriod ? objectIndex % 21 : 0),
      quality: qualityBase + (isCurrentPeriod ? (objectIndex % 5) - 2 : 0),
      errors: Math.max(0, errorsBase + (isCurrentPeriod ? (objectIndex % 3) - 1 : 0)),
      aht: ahtBase + (isCurrentPeriod ? (objectIndex % 11) - 5 : 0),
    };

    return headers.map(function (header, columnIndex) {
      if (Object.prototype.hasOwnProperty.call(baseValues, header.id)) {
        return baseValues[header.id];
      }

      return (objectIndex + columnIndex + (isCurrentPeriod ? 3 : 0)) % 1000;
    });
  }

  function createCsvFile(csv, options) {
    const fileName = "metricum-benchmark-" + options.rowCount + "x" + options.columnCount + ".csv";

    if (typeof File === "function") {
      return new File([csv], fileName, { type: "text/csv;charset=utf-8" });
    }

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    blob.name = fileName;
    return blob;
  }

  function getBenchmarkColumnIds(table) {
    return {
      period: findColumnId(table, BENCHMARK_PERIOD_HEADER),
      employee: findColumnId(table, BENCHMARK_ID_HEADER),
      sales: findColumnId(table, "Продажи"),
      quality: findColumnId(table, "Качество %"),
      errors: findColumnId(table, "Ошибки"),
      aht: findColumnId(table, "AHT"),
      extra_1: findColumnId(table, "Доп поле 1"),
    };
  }

  function findColumnId(table, name) {
    const header = table.headers.find(function (item) {
      return item.name === name;
    });

    if (!header) {
      throw new Error("Не найдена benchmark-колонка: " + name);
    }

    return header.id;
  }

  function buildMetrics(periods, metricCount, columnIds) {
    const metricColumns = ["sales", "quality", "errors", "aht", "extra_1"].slice(0, metricCount);

    return metricColumns.map(function (columnId, index) {
      const columns = {};

      periods.forEach(function (period) {
        columns[period.id] = columnIds[columnId];
      });

      return {
        id: "metric_" + columnId,
        label: metricLabel(columnId),
        aggregation: columnId === "quality" ? "avg" : "auto",
        columns: columns,
      };
    });
  }

  function metricLabel(columnId) {
    const labels = {
      sales: "Продажи",
      quality: "Качество",
      errors: "Ошибки",
      aht: "AHT",
      extra_1: "Доп. показатель",
    };

    return labels[columnId] || columnId;
  }

  function averageResults(results, options) {
    if (results.length === 1) {
      return Object.assign({}, results[0], { repeatCount: 1 });
    }

    const first = results[0];
    const numericKeys = ["generateMs", "readMs", "periodMs", "compareMs", "analyticsMs", "totalMs"];
    const output = Object.assign({}, first, {
      repeatCount: options.repeatCount,
      warningCount: Math.max.apply(null, results.map(function (result) { return result.warningCount; })),
    });

    numericKeys.forEach(function (key) {
      output[key] = results.reduce(function (sum, result) {
        return sum + result[key];
      }, 0) / results.length;
    });

    return output;
  }

  function appendResult(result) {
    if (!hasResults) {
      dom.results.innerHTML = "";
      hasResults = true;
    }

    const row = document.createElement("tr");
    row.innerHTML =
      cell(formatNumber(result.rowCount), "numeric") +
      cell(formatNumber(result.columnCount), "numeric") +
      cell(formatNumber(result.cellCount), "numeric") +
      cell(formatBytes(result.fileSize), "numeric") +
      cell(formatMs(result.generateMs), "numeric") +
      cell(formatMs(result.readMs), "numeric") +
      cell(formatMs(result.periodMs), "numeric") +
      cell(formatMs(result.compareMs), "numeric") +
      cell(formatMs(result.analyticsMs), "numeric") +
      cell(formatMs(result.totalMs), "numeric") +
      cell(formatNumber(result.objectCount), "numeric") +
      cell(formatWarnings(result));

    dom.results.appendChild(row);
  }

  function clearResults() {
    if (isRunning) {
      return;
    }

    hasResults = false;
    dom.results.innerHTML = '<tr><td colspan="12" class="empty">Результаты появятся после запуска benchmark.</td></tr>';
    setStatus("Готов к запуску", "Выберите размер таблицы. Во время чтения крупного файла вкладка может временно задуматься.", false);
  }

  function setControlsDisabled(disabled) {
    dom.rowButtons.forEach(function (button) {
      button.disabled = disabled;
    });
    dom.runAllButton.disabled = disabled;
    dom.clearButton.disabled = disabled;
    dom.columnCountSelect.disabled = disabled;
    dom.metricCountSelect.disabled = disabled;
    dom.comparisonModeSelect.disabled = disabled;
    dom.repeatCountSelect.disabled = disabled;
  }

  async function setStatus(title, detail, active) {
    dom.status.classList.toggle("active", active !== false);
    dom.status.innerHTML =
      '<span class="spinner" aria-hidden="true"></span>' +
      "<span><strong>" +
      escapeHtml(title || "Обработка") +
      "</strong>" +
      escapeHtml(detail || "Подождите немного") +
      "</span>";

    await yieldToBrowser();
  }

  function cell(value, className) {
    return '<td class="' + escapeHtml(className || "") + '">' + escapeHtml(value) + "</td>";
  }

  function formatScenario(options) {
    return formatNumber(options.rowCount) + " строк, " + options.columnCount + " колонок, " + formatNumber(options.rowCount * options.columnCount) + " ячеек";
  }

  function formatWarnings(result) {
    const prefix = result.repeatCount > 1 ? "среднее за " + result.repeatCount + " прогона; " : "";
    return prefix + (result.warningCount ? formatNumber(result.warningCount) : "нет");
  }

  function formatMs(value) {
    return formatNumber(Math.round(value)) + " мс";
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("ru-RU", {
      maximumFractionDigits: 0,
    }).format(Number(value) || 0);
  }

  function formatBytes(value) {
    const safeValue = Number(value) || 0;
    const units = ["Б", "КБ", "МБ", "ГБ"];
    let size = safeValue;
    let unitIndex = 0;

    while (size >= 1024 && unitIndex < units.length - 1) {
      size /= 1024;
      unitIndex += 1;
    }

    return new Intl.NumberFormat("ru-RU", {
      maximumFractionDigits: unitIndex === 0 ? 0 : 1,
    }).format(size) + " " + units[unitIndex];
  }

  function byteLength(text) {
    if (global.TextEncoder) {
      return new TextEncoder().encode(text).length;
    }

    return text.length;
  }

  function now() {
    return global.performance && typeof global.performance.now === "function"
      ? global.performance.now()
      : Date.now();
  }

  function elapsed(start) {
    return now() - start;
  }

  function yieldToBrowser() {
    return new Promise(function (resolve) {
      if (typeof global.requestAnimationFrame === "function") {
        global.requestAnimationFrame(function () {
          resolve();
        });
        return;
      }

      global.setTimeout(resolve, 0);
    });
  }

  function escapeHtml(value) {
    return String(value === null || value === undefined ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
})(window);
