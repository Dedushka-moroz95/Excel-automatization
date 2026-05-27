(function (global) {
  const App = (global.OperationalAnalytics = global.OperationalAnalytics || {});
  const Normalizers = App.Normalizers;

  function exportCsv(comparison, metrics) {
    const rows = buildFlatRows(comparison, metrics);
    const csv = rows
      .map(function (row) {
        return row.map(toCsvCell).join(";");
      })
      .join("\r\n");

    downloadBlob("\ufeff" + csv, "comparison-report.csv", "text/csv;charset=utf-8");
  }

  async function exportExcel(comparison, metrics, analytics) {
    if (!global.ExcelJS) {
      throw new Error("Библиотека ExcelJS не загружена");
    }

    const workbook = new global.ExcelJS.Workbook();
    workbook.creator = "Operational Analytics";
    workbook.created = new Date();

    fillSummarySheet(workbook.addWorksheet("Summary"), analytics);
    fillComparisonSheet(workbook.addWorksheet("Comparison"), comparison, metrics);
    fillMissingSheet(workbook.addWorksheet("Missing"), comparison);
    fillDuplicateSheet(workbook.addWorksheet("Duplicates"), comparison);

    workbook.worksheets.forEach(function (sheet) {
      sheet.views = [{ state: "frozen", ySplit: 1 }];
      sheet.columns.forEach(function (column) {
        column.width = Math.min(Math.max(column.width || 14, 14), 32);
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    downloadBlob(buffer, "comparison-report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }

  function buildFlatRows(comparison, metrics) {
    const header = ["Юнит"];
    const isSequential = comparison.comparisonMode === "sequential";
    const showTimeline = comparison.periods.length > 2 && !isSequential;

    metrics.forEach(function (metric) {
      if (isSequential) {
        comparison.comparisonPairs.forEach(function (pair) {
          header.push(metric.label + " - " + pair.label);
        });
        return;
      }

      if (showTimeline) {
        comparison.periods.forEach(function (period) {
          header.push(metric.label + " - " + period.label);
        });
      }

      header.push(metric.label + " - итоговая динамика");
      header.push(metric.label + " - динамика %");
      header.push(metric.label + " - статус");
    });

    const rows = [header];

    comparison.rows.forEach(function (row) {
      const output = [row.label];

      metrics.forEach(function (metric) {
        const result = row.metrics.find(function (item) {
          return item.metricId === metric.id;
        });

        if (isSequential) {
          comparison.comparisonPairs.forEach(function (pair) {
            const item = result
              ? result.comparisons.find(function (comparisonItem) {
                  return comparisonItem.fromPeriodId === pair.fromPeriodId && comparisonItem.toPeriodId === pair.toPeriodId;
                })
              : null;
            output.push(formatResultForExport(item));
          });
          return;
        }

        if (showTimeline) {
          comparison.periods.forEach(function (period) {
            const periodValue = result
              ? result.periodValues.find(function (item) {
                  return item.periodId === period.id;
                })
              : null;
            output.push(periodValue && periodValue.isNumeric ? Normalizers.formatMetricValue(periodValue.value, periodValue.valueFormat, 2) : "");
          });
        }

        output.push(result && Number.isFinite(result.delta) ? Normalizers.formatMetricDelta(result.delta, result.valueFormat, 2) : "");
        output.push(result && Number.isFinite(result.deltaPercent) ? result.deltaPercent : "");
        output.push(result ? translateImpact(result.impact) : "");
      });

      rows.push(output);
    });

    return rows;
  }

  function fillSummarySheet(sheet, analytics) {
    sheet.columns = [
      { header: "Показатель", key: "name", width: 28 },
      { header: "Значение", key: "value", width: 16 },
    ];

    sheet.addRows([
      { name: "Периодов", value: analytics.periodCount },
      { name: "Всего юнитов", value: analytics.totalUnits },
      { name: "Полных рядов", value: analytics.totalCompared },
      { name: "Отсутствующие юниты", value: analytics.missingTotal },
      { name: "Дубликаты юнитов", value: analytics.duplicateIds },
      { name: "Нечисловые значения", value: analytics.invalidValues },
    ]);

    styleHeaderRow(sheet);
  }

  function fillComparisonSheet(sheet, comparison, metrics) {
    sheet.addRows(buildFlatRows(comparison, metrics));
    styleHeaderRow(sheet);
  }

  function fillMissingSheet(sheet, comparison) {
    sheet.addRow(["Период", "Юнит"]);

    comparison.missingByPeriod.forEach(function (group) {
      group.items.forEach(function (item) {
        sheet.addRow([group.periodLabel, item.label]);
      });
    });

    styleHeaderRow(sheet);
  }

  function fillDuplicateSheet(sheet, comparison) {
    sheet.addRow(["Период", "Юнит", "Строки"]);

    comparison.duplicatesByPeriod.forEach(function (group) {
      group.items.forEach(function (item) {
        sheet.addRow([group.periodLabel, item.label, item.rowNumbers.join(", ")]);
      });
    });

    styleHeaderRow(sheet);
  }

  function styleHeaderRow(sheet) {
    const header = sheet.getRow(1);
    header.font = { bold: true, color: { argb: "FF1F1F24" } };
    header.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFDD2D" },
    };
    header.alignment = { vertical: "middle" };
  }

  function toCsvCell(value) {
    if (value === null || value === undefined) {
      return "";
    }

    let text = String(value);

    if (/^[=+\-@]/.test(text)) {
      text = "'" + text;
    }

    if (/[;"\n\r]/.test(text)) {
      return '"' + text.replace(/"/g, '""') + '"';
    }

    return text;
  }

  function translateImpact(value) {
    const labels = {
      good: "рост",
      bad: "снижение",
      neutral: "без изменений",
      unknown: "нет данных",
    };

    return labels[value] || value;
  }

  function formatResultForExport(result) {
    if (!result || !Number.isFinite(result.delta)) {
      return "";
    }

    const percent =
      Number.isFinite(result.deltaPercent)
        ? " (" + (result.deltaPercent > 0 ? "+" : "") + Normalizers.formatPercent(result.deltaPercent) + ")"
        : "";

    return Normalizers.formatMetricDelta(result.delta, result.valueFormat, 2) + percent + " · " + translateImpact(result.impact);
  }

  function downloadBlob(content, fileName, mimeType) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  App.Exporters = {
    exportCsv,
    exportExcel,
    buildFlatRows,
    translateImpact,
    formatResultForExport,
  };
})(window);
