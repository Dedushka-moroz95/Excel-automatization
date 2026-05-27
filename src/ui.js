(function (global) {
  const App = (global.OperationalAnalytics = global.OperationalAnalytics || {});
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
    const warnCount = warnings.length - errorCount;
    const tone = !hasFiles ? "neutral" : errorCount ? "error" : warnings.length ? "warn" : "ok";
    const title = !hasFiles ? "Проверок пока нет" : errorCount ? "Нужна проверка" : warnings.length ? "Есть предупреждения" : "Файлы готовы";
    const summary = !hasFiles
      ? "Загрузите файлы, чтобы увидеть статус"
      : errorCount
        ? "Ошибок: " + errorCount + ", предупреждений: " + warnCount
        : warnings.length
          ? "Предупреждений: " + warnings.length
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
      summaryCard("Всего юнитов", analytics.totalUnits) +
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
          Normalizers.formatNumber(item.delta, 2) +
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
    const isSequential = comparison.comparisonMode === "sequential";
    const showTimeline = comparison.periods.length > 2 && !isSequential;
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
      '<div class="table-scroll"><table class="results-table"><thead><tr><th>Юнит</th>' +
      metricHeaders +
      "</tr></thead><tbody>" +
      rowsHtml +
      "</tbody></table></div>";
  }

  function emptyMetricCells(periodCount, showTimeline, comparison) {
    const count = comparison.comparisonMode === "sequential" ? comparison.comparisonPairs.length : showTimeline ? periodCount + 1 : 1;
    return Array.from({ length: count })
      .map(function () {
        return "<td>—</td>";
      })
      .join("");
  }

  function buildMetricHeaders(metric, periods, showTimeline, comparison) {
    if (comparison.comparisonMode === "sequential") {
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
    if (comparison.comparisonMode === "sequential") {
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

    return '<td class="number">' + Normalizers.formatNumber(item.value, 2) + "</td>";
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
      signedNumber(result.delta) +
      "</strong><span>" +
      impactLabel(result.impact) +
      percent +
      "</span></div></td>"
    );
  }

  function signedNumber(value) {
    if (!Number.isFinite(value)) {
      return "—";
    }

    const sign = value > 0 ? "+" : "";
    return sign + Normalizers.formatNumber(value, 2);
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

  App.UI = {
    escapeHtml,
    renderPreview,
    fillColumnSelect,
    renderWarnings,
    renderSummary,
    renderMovers,
    renderResultsTable,
    renderChartMetricSelect,
    impactPill,
    impactLabel,
    translateImpact: Exporters.translateImpact,
  };
})(window);
