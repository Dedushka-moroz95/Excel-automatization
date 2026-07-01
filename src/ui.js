(function (global) {
  const App = (global.Metricum = global.Metricum || {});
  const Normalizers = App.Normalizers;
  const Exporters = App.Exporters;

  function escapeHtml(value) {
    return String(value === null || value === undefined ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function renderPreview(container, table, emptyText) {
    if (!table) {
      container.className = "preview-panel empty-state";
      container.textContent = emptyText;
      return;
    }

    container.className = "preview-panel";
    const cellCount = table.workload && Number.isFinite(table.workload.cellCount)
      ? table.workload.cellCount
      : table.rows.length * table.headers.length;
    const headersHtml = table.headers
      .map(function (header) {
        return "<th>" + escapeHtml(header.name) + "</th>";
      })
      .join("");

    const rowsHtml = table.previewRows
      .map(function (row) {
        const cells = table.headers
          .map(function (header) {
            const value = row.values[header.id];
            const displayValue = value === null || value === undefined ? "" : value;
            return "<td>" + escapeHtml(displayValue) + "</td>";
          })
          .join("");

        return "<tr>" + cells + "</tr>";
      })
      .join("");

    container.innerHTML =
      '<div class="table-meta">' +
      "<span>" +
      escapeHtml(table.fileName) +
      "</span>" +
      "<span>Лист: " +
      escapeHtml(table.sheetName) +
      "</span>" +
      "<span>Строк: " +
      table.rows.length +
      "</span>" +
      "<span>Колонок: " +
      table.headers.length +
      "</span>" +
      "<span>Ячеек: " +
      escapeHtml(Normalizers.formatNumber(cellCount, 0)) +
      "</span>" +
      "</div>" +
      '<div class="table-scroll"><table class="preview-table"><thead><tr>' +
      headersHtml +
      "</tr></thead><tbody>" +
      rowsHtml +
      "</tbody></table></div>";
  }

  function fillColumnSelect(select, table, selectedValue, placeholder) {
    select.innerHTML = "";
    select.disabled = !table;

    const placeholderOption = document.createElement("option");
    placeholderOption.value = "";
    placeholderOption.textContent = placeholder || "Выберите колонку";
    select.appendChild(placeholderOption);

    if (!table) {
      return;
    }

    table.headers.forEach(function (header) {
      const option = document.createElement("option");
      option.value = header.id;
      option.textContent = header.name;
      select.appendChild(option);
    });

    select.value = selectedValue || "";
  }

  function renderWarnings(container, warnings) {
    if (container.dataset.variant === "inline-quality") {
      renderInlineQuality(container, warnings);
      return;
    }

    if (!warnings.length) {
      setClassName(container, "warnings-panel");
      container.innerHTML = '<div class="warning-item ok">Критичных проблем не найдено</div>';
      return;
    }

    setClassName(container, "warnings-panel");
    container.innerHTML = warnings
      .map(function (warning) {
        return (
          '<div class="warning-item ' +
          escapeHtml(warning.type || "warn") +
          '">' +
          escapeHtml(warning.message) +
          "</div>"
        );
      })
      .join("");
  }

  function renderInlineQuality(container, warnings) {
    const hasFiles = container.dataset.hasFiles === "true";
    const errorCount = warnings.filter(function (warning) {
      return warning.type === "error";
    }).length;
    const actionableWarnings = warnings.filter(function (warning) {
      return warning.type !== "neutral";
    }).length;
    const warnCount = actionableWarnings - errorCount;
    const hasOnlyNeutralMessages = warnings.length && !actionableWarnings;
    const tone = !hasFiles ? "neutral" : errorCount ? "error" : actionableWarnings ? "warn" : hasOnlyNeutralMessages ? "neutral" : "ok";
    const title = !hasFiles
      ? "Проверок пока нет"
      : errorCount
        ? "Нужна проверка"
        : actionableWarnings
          ? "Есть предупреждения"
          : hasOnlyNeutralMessages
            ? "Режим выбран"
            : "Файлы готовы";
    const summary = !hasFiles
      ? "Загрузите файлы, чтобы увидеть статус"
      : errorCount
        ? "Ошибок: " + errorCount + ", предупреждений: " + warnCount
        : actionableWarnings
          ? "Предупреждений: " + actionableWarnings
          : hasOnlyNeutralMessages
            ? "Подробности в деталях"
          : "Критичных проблем не найдено";

    setClassName(container, "data-quality-inline " + tone);
    container.dataset.variant = "inline-quality";
    container.innerHTML =
      '<div class="quality-main">' +
      '<span class="quality-dot"></span>' +
      "<div>" +
      "<strong>Качество данных</strong>" +
      "<span>" +
      escapeHtml(title + ". " + summary) +
      "</span>" +
      "</div></div>" +
      qualityDetails(warnings);
  }

  function qualityDetails(warnings) {
    if (!warnings.length) {
      return "";
    }

    const visibleWarnings = warnings.slice(0, 5)
      .map(function (warning) {
        return (
          '<li class="quality-item ' +
          escapeHtml(warning.type || "warn") +
          '">' +
          escapeHtml(warning.message) +
          "</li>"
        );
      })
      .join("");
    const restCount = warnings.length > 5 ? '<li class="quality-more">Еще ' + (warnings.length - 5) + "</li>" : "";

    return (
      '<details class="quality-details">' +
      "<summary>Детали</summary>" +
      '<ul class="quality-list">' +
      visibleWarnings +
      restCount +
      "</ul></details>"
    );
  }

  function setClassName(container, className) {
    const motionClasses = ["reveal-item", "is-visible"].filter(function (item) {
      return container.classList.contains(item);
    });

    container.className = className + (motionClasses.length ? " " + motionClasses.join(" ") : "");
  }

  function renderSummary(container, analytics) {
    if (!analytics) {
      container.className = "summary-grid empty-state";
      container.textContent = "Результаты появятся после расчета";
      return;
    }

    container.className = "summary-grid";
    container.innerHTML =
      summaryCard("Полных рядов", analytics.totalCompared) +
      summaryCard("Всего объектов", analytics.totalUnits) +
      summaryCard("Периодов", analytics.periodCount) +
      summaryCard("Проблем", analytics.missingTotal + analytics.duplicateIds + analytics.invalidValues);
  }

  function summaryCard(label, value) {
    const numericValue = Number(value) || 0;
    return (
      '<div class="summary-card"><span>' +
      escapeHtml(label) +
      '</span><strong data-count-to="' +
      numericValue +
      '">' +
      numericValue +
      "</strong></div>"
    );
  }

  function renderMovers(container, analytics) {
    if (!analytics || !analytics.metricSummaries.length) {
      container.innerHTML = "";
      return;
    }

    container.innerHTML = analytics.metricSummaries
      .map(function (summary) {
        return (
          '<div class="mover-card">' +
          "<h3>" +
          escapeHtml(summary.label) +
          "</h3>" +
          '<div class="table-meta">' +
          "<span>Рост: " +
          summary.improvedCount +
          "</span><span>Снижение: " +
          summary.declinedCount +
          "</span><span>Без изменений: " +
          summary.unchangedCount +
          "</span></div>" +
          '<div class="movers-grid">' +
          moverList("Макс. рост", summary.best, "good") +
          moverList("Макс. снижение", summary.worst, "bad") +
          "</div></div>"
        );
      })
      .join("");
  }

  function moverList(title, items, type) {
    if (!items.length) {
      return '<div><h3>' + escapeHtml(title) + '</h3><p class="muted">Нет данных</p></div>';
    }

    const rows = items
      .map(function (item) {
        const deltaClass = type === "good" ? "good-text" : "bad-text";
        return (
          "<li><span>" +
          escapeHtml(item.label) +
          (item.comparisonLabel ? '<small class="mover-period">' + escapeHtml(item.comparisonLabel) + "</small>" : "") +
          '</span><strong class="' +
          deltaClass +
          '">' +
          Normalizers.formatMetricDelta(item.delta, item.valueFormat, 2) +
          "</strong></li>"
        );
      })
      .join("");

    return "<div><h3>" + escapeHtml(title) + '</h3><ul class="mover-list">' + rows + "</ul></div>";
  }

  function renderResultsTable(container, comparison, metrics) {
    if (!comparison || !comparison.rows.length) {
      container.className = "results-panel empty-state";
      container.textContent = "Нет рассчитанных данных";
      return;
    }

    container.className = "results-panel";
    const visibleRows = comparison.rows.slice(0, 300);
    const isPairwise = comparison.comparisonMode === "sequential" || comparison.comparisonMode === "manual";
    const showTimeline = comparison.periods.length > 2 && !isPairwise;
    const metricHeaders = metrics.map(function (metric) {
      return buildMetricHeaders(metric, comparison.periods, showTimeline, comparison);
    }).join("");

    const rowsHtml = visibleRows
      .map(function (row) {
        const metricCells = metrics
          .map(function (metric) {
            const result = row.metrics.find(function (item) {
              return item.metricId === metric.id;
            });

            if (!result) {
              return emptyMetricCells(comparison.periods.length, showTimeline, comparison);
            }

            return buildMetricCells(result, comparison.periods, showTimeline, comparison);
          })
          .join("");

        return "<tr><td>" + escapeHtml(row.label) + "</td>" + metricCells + "</tr>";
      })
      .join("");

    const limitNote =
      comparison.rows.length > visibleRows.length
        ? '<div class="table-meta"><span>Показаны первые ' + visibleRows.length + " из " + comparison.rows.length + "</span></div>"
        : "";

    container.innerHTML =
      limitNote +
      '<div class="table-scroll"><table class="results-table"><thead><tr><th>Объект</th>' +
      metricHeaders +
      "</tr></thead><tbody>" +
      rowsHtml +
      "</tbody></table></div>";
  }

  function emptyMetricCells(periodCount, showTimeline, comparison) {
    const isPairwise = comparison.comparisonMode === "sequential" || comparison.comparisonMode === "manual";
    const count = isPairwise ? comparison.comparisonPairs.length : showTimeline ? periodCount + 1 : 1;
    return Array.from({ length: count })
      .map(function () {
        return "<td>—</td>";
      })
      .join("");
  }

  function buildMetricHeaders(metric, periods, showTimeline, comparison) {
    if (comparison.comparisonMode === "sequential" || comparison.comparisonMode === "manual") {
      return comparison.comparisonPairs
        .map(function (pair) {
          return "<th>" + escapeHtml(metric.label) + '<span class="th-subtitle">' + escapeHtml(pair.label) + "</span></th>";
        })
        .join("");
    }

    if (!showTimeline) {
      return "<th>" + escapeHtml(metric.label) + "</th>";
    }

    const periodHeaders = periods
      .map(function (period) {
        return "<th>" + escapeHtml(metric.label) + '<span class="th-subtitle">' + escapeHtml(period.label) + "</span></th>";
      })
      .join("");

    return (
      periodHeaders +
      "<th>" +
      escapeHtml(metric.label) +
      '<span class="th-subtitle">Итоговая динамика</span></th>'
    );
  }

  function buildMetricCells(result, periods, showTimeline, comparison) {
    if (comparison.comparisonMode === "sequential" || comparison.comparisonMode === "manual") {
      return result.comparisons
        .map(function (item) {
          return resultCell(item);
        })
        .join("");
    }

    if (!showTimeline) {
      return resultCell(result);
    }

    const valueCells = periods
      .map(function (period) {
        const item = result.periodValues.find(function (value) {
          return value.periodId === period.id;
        });

        return periodValueCell(item);
      })
      .join("");

    return valueCells + resultCell(result);
  }

  function periodValueCell(item) {
    if (!item || !item.isNumeric) {
      return '<td class="warn-text">нет данных</td>';
    }

    return '<td class="number">' + Normalizers.formatMetricValue(item.value, item.valueFormat, 2) + "</td>";
  }

  function resultCell(result) {
    if (!Number.isFinite(result.delta)) {
      return '<td class="warn-text">нет данных</td>';
    }

    const tone = result.impact === "good" ? "good-text" : result.impact === "bad" ? "bad-text" : "muted";
    const percent = Number.isFinite(result.deltaPercent) ? " · " + signedPercent(result.deltaPercent) : "";

    return (
      '<td class="number"><div class="result-cell"><strong class="' +
      tone +
      '">' +
      Normalizers.formatMetricDelta(result.delta, result.valueFormat, 2) +
      "</strong><span>" +
      impactLabel(result.impact) +
      percent +
      "</span></div></td>"
    );
  }

  function signedPercent(value) {
    if (!Number.isFinite(value)) {
      return "—";
    }

    const sign = value > 0 ? "+" : "";
    return sign + Normalizers.formatPercent(value);
  }

  function impactPill(impact) {
    const labels = {
      good: "рост",
      bad: "снижение",
      neutral: "без изменений",
      unknown: "нет данных",
    };
    const classes = {
      good: "good",
      bad: "bad",
      neutral: "neutral",
      unknown: "warn",
    };

    return '<span class="status-pill ' + classes[impact] + '">' + labels[impact] + "</span>";
  }

  function impactLabel(impact) {
    const labels = {
      good: "рост",
      bad: "снижение",
      neutral: "без изменений",
      unknown: "нет данных",
    };

    return labels[impact] || impact;
  }

  function renderChartMetricSelect(select, metrics, selectedValue) {
    select.innerHTML = "";
    select.disabled = !metrics.length;

    metrics.forEach(function (metric) {
      const option = document.createElement("option");
      option.value = metric.id;
      option.textContent = metric.label || "Показатель";
      select.appendChild(option);
    });

    select.value = selectedValue || (metrics[0] ? metrics[0].id : "");
  }

  function renderChartTypeSelect(select, selectedValue) {
    if (!select) {
      return;
    }

    const types = App.Charts && App.Charts.getChartTypes ? App.Charts.getChartTypes() : [];
    const fallbackTypes = types.length
      ? types
      : [
          { id: "bar-horizontal", label: "Динамика по объектам" },
          { id: "bar-vertical", label: "Вертикальные столбцы" },
          { id: "line-trend", label: "Тренд по периодам" },
          { id: "doughnut-impact", label: "Структура изменений" },
        ];

    select.innerHTML = fallbackTypes
      .map(function (type) {
        return '<option value="' + escapeHtml(type.id) + '">' + escapeHtml(type.label) + "</option>";
      })
      .join("");

    select.value = selectedValue || "bar-horizontal";

    if (!select.value) {
      select.value = "bar-horizontal";
    }
  }

  function renderHistory(container, records, options) {
    const renderOptions = options || {};

    if (!records.length) {
      container.className = "history-list";
      container.innerHTML = renderHistoryEmpty(renderOptions);
      return;
    }

    container.className = "history-list" + (renderOptions.query ? " history-list--filtered" : "");
    container.innerHTML = records.map(renderHistoryCard).join("");
  }

  function renderHistoryEmpty(options) {
    if (options.hasRecords && options.query) {
      return (
        '<div class="history-empty-card">' +
        "<h3>Ничего не найдено</h3>" +
        "<p>Попробуйте изменить запрос или очистить поле поиска.</p>" +
        "</div>"
      );
    }

    return (
      '<div class="history-empty-card">' +
      "<h3>История пока пуста</h3>" +
      "<p>Сохраните первый анализ после выполнения сравнения, чтобы быстро возвращаться к результатам без повторной загрузки файлов.</p>" +
      "</div>"
    );
  }

  function renderHistoryCard(record) {
    const meta = record.meta || {};
    const metricCount = Number(meta.metricCount) || (Array.isArray(record.metrics) ? record.metrics.length : 0);
    const periodCount = Number(meta.periodCount) || (record.comparison && record.comparison.periods ? record.comparison.periods.length : 0);
    const totalUnits = Number(meta.totalUnits) || (record.analytics ? record.analytics.totalUnits : 0) || 0;
    const totalCompared = Number(meta.totalCompared) || (record.analytics ? record.analytics.totalCompared : 0) || 0;
    const rowCount = Number(meta.rowCount) || (record.comparison && Array.isArray(record.comparison.rows) ? record.comparison.rows.length : 0);
    const changedCount = Number(meta.changedCount) || getHistoryChangedCount(record);
    const modeLabel =
      meta.comparisonMode === "manual" ? "ручной" : meta.comparisonMode === "sequential" ? "последовательный" : "итоговый";
    const pinnedLabel = record.pinned ? '<span class="history-pin">★ Закреплено</span>' : "";
    const pinnedClass = record.pinned ? " history-card--pinned" : "";
    const pinAction = record.pinned ? "Открепить" : "Закрепить";
    const identifierLabel = meta.identifierLabel ? '<span>Объект: ' + escapeHtml(meta.identifierLabel) + "</span>" : "";

    return (
      '<article class="history-card' +
      pinnedClass +
      '" data-history-id="' +
      escapeHtml(record.id) +
      '">' +
      '<div class="history-card__body">' +
      '<div class="history-card__title-row">' +
      "<h3>" +
      escapeHtml(record.title) +
      "</h3>" +
      pinnedLabel +
      "</div>" +
      '<div class="history-date">' +
      escapeHtml(formatHistoryDate(record.createdAt)) +
      "</div>" +
      '<div class="history-meta">' +
      "<span>" +
      periodCount +
      " периодов</span>" +
      "<span>" +
      metricCount +
      " показателей</span>" +
      "<span>Режим: " +
      escapeHtml(modeLabel) +
      "</span>" +
      identifierLabel +
      "</div></div>" +
      '<div class="history-stats">' +
      historyStat("Объектов", totalUnits) +
      historyStat("Строк", rowCount) +
      historyStat("Полных строк", totalCompared) +
      historyStat("Изменений", changedCount) +
      "</div>" +
      '<div class="history-actions">' +
      '<button class="button button-primary" type="button" data-action="open-history" data-history-id="' +
      escapeHtml(record.id) +
      '">Открыть</button>' +
      '<button class="button button-secondary" type="button" data-action="toggle-history-pin" data-history-id="' +
      escapeHtml(record.id) +
      '">' +
      pinAction +
      "</button>" +
      '<button class="button button-secondary" type="button" data-action="rename-history" data-history-id="' +
      escapeHtml(record.id) +
      '">Переименовать</button>' +
      '<button class="button button-secondary history-delete-button" type="button" data-action="delete-history" data-history-id="' +
      escapeHtml(record.id) +
      '">Удалить</button>' +
      "</div></article>"
    );
  }

  function historyStat(label, value) {
    return (
      '<div class="history-stat"><span>' +
      escapeHtml(label) +
      "</span><strong>" +
      escapeHtml(formatCompactNumber(value)) +
      "</strong></div>"
    );
  }

  function getHistoryChangedCount(record) {
    const analytics = record.analytics;

    if (!analytics || !Array.isArray(analytics.metricSummaries)) {
      return 0;
    }

    return analytics.metricSummaries.reduce(function (sum, summary) {
      return sum + (Number(summary.improvedCount) || 0) + (Number(summary.declinedCount) || 0);
    }, 0);
  }

  function formatCompactNumber(value) {
    return new Intl.NumberFormat("ru-RU", {
      maximumFractionDigits: 0,
    }).format(Number(value) || 0);
  }

  function formatHistoryDate(value) {
    if (App.HistoryStore && typeof App.HistoryStore.formatDateTime === "function") {
      return App.HistoryStore.formatDateTime(value);
    }

    const date = value ? new Date(value) : new Date();

    if (Number.isNaN(date.getTime())) {
      return "";
    }

    return new Intl.DateTimeFormat("ru-RU", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    }).format(date);
  }

  App.UI = {
    escapeHtml,
    renderPreview,
    fillColumnSelect,
    renderWarnings,
    renderSummary,
    renderMovers,
    renderResultsTable,
    renderChartMetricSelect,
    renderChartTypeSelect,
    renderHistory,
    impactPill,
    impactLabel,
    translateImpact: Exporters.translateImpact,
  };
})(window);
