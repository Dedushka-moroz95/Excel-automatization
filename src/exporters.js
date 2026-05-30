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

  async function exportExcel(comparison, metrics, analytics, options) {
    if (!global.ExcelJS) {
      throw new Error("Библиотека ExcelJS не загружена");
    }

    const exportOptions = options || {};
    const workbook = new global.ExcelJS.Workbook();
    workbook.creator = "Operational Analytics";
    workbook.created = new Date();

    fillDashboardSheet(workbook, comparison, analytics, exportOptions.chartMetric || metrics[0]);
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

  function fillDashboardSheet(workbook, comparison, analytics, chartMetric) {
    const sheet = workbook.addWorksheet("Dashboard");
    const chartDataUrl = buildDashboardChartImage(comparison, chartMetric);
    const metricLabel = chartMetric ? chartMetric.label : "Показатель";

    sheet.columns = [
      { key: "a", width: 24 },
      { key: "b", width: 18 },
      { key: "c", width: 18 },
      { key: "d", width: 18 },
      { key: "e", width: 18 },
      { key: "f", width: 18 },
    ];

    sheet.mergeCells("A1:F1");
    sheet.getCell("A1").value = "Dashboard";
    sheet.getCell("A1").font = { bold: true, size: 22, color: { argb: "FF18212F" } };
    sheet.getCell("A1").alignment = { vertical: "middle" };
    sheet.getRow(1).height = 30;

    sheet.mergeCells("A2:F2");
    sheet.getCell("A2").value = "График динамики: " + metricLabel;
    sheet.getCell("A2").font = { bold: true, size: 12, color: { argb: "FF667085" } };

    sheet.addRow([]);
    sheet.addRow(["Показатель", "Значение"]);
    sheet.addRow(["Периодов", analytics.periodCount]);
    sheet.addRow(["Всего юнитов", analytics.totalUnits]);
    sheet.addRow(["Полных рядов", analytics.totalCompared]);
    sheet.addRow(["Проблем", analytics.missingTotal + analytics.duplicateIds + analytics.invalidValues]);

    styleHeaderRow(sheet, 4);

    if (chartDataUrl && typeof workbook.addImage === "function") {
      const imageId = workbook.addImage({
        base64: chartDataUrl,
        extension: "png",
      });

      sheet.addImage(imageId, {
        tl: { col: 0, row: 9 },
        ext: { width: 980, height: 500 },
      });
    } else {
      sheet.getCell("A11").value = "График недоступен для экспорта";
    }
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

  function buildDashboardChartImage(comparison, metric) {
    if (!metric || !global.document) {
      return "";
    }

    const rows = buildChartRows(comparison, metric);

    if (!rows.length) {
      return "";
    }

    const canvas = global.document.createElement("canvas");
    const width = 1200;
    const height = 620;
    const scale = global.devicePixelRatio || 1;
    canvas.width = width * scale;
    canvas.height = height * scale;
    canvas.style.width = width + "px";
    canvas.style.height = height + "px";

    const context = canvas.getContext("2d");
    if (!context) {
      return "";
    }

    context.scale(scale, scale);
    drawChart(context, width, height, rows, metric);

    return canvas.toDataURL("image/png");
  }

  function buildChartRows(comparison, metric) {
    const isSequential = comparison.comparisonMode === "sequential";

    return comparison.rows
      .flatMap(function (row) {
        const result = row.metrics.find(function (item) {
          return item.metricId === metric.id;
        });

        if (!result) {
          return [];
        }

        if (isSequential) {
          return result.comparisons
            .filter(function (item) {
              return Number.isFinite(item.delta);
            })
            .map(function (item) {
              return {
                label: row.label + " · " + item.label,
                delta: item.delta,
                impact: item.impact,
                valueFormat: item.valueFormat || result.valueFormat || "number",
              };
            });
        }

        if (!Number.isFinite(result.delta)) {
          return [];
        }

        return [{
          label: row.label,
          delta: result.delta,
          impact: result.impact,
          valueFormat: result.valueFormat || "number",
        }];
      })
      .sort(function (left, right) {
        return Math.abs(right.delta) - Math.abs(left.delta);
      })
      .slice(0, 15)
      .reverse();
  }

  function drawChart(context, width, height, rows, metric) {
    const padding = {
      top: 86,
      right: 170,
      bottom: 62,
      left: 260,
    };
    const chartWidth = width - padding.left - padding.right;
    const chartHeight = height - padding.top - padding.bottom;
    const rowGap = 10;
    const barHeight = Math.max(16, Math.min(26, (chartHeight - rowGap * (rows.length - 1)) / rows.length));
    const maxAbs = Math.max.apply(
      null,
      rows.map(function (row) {
        return Math.abs(row.delta);
      })
    ) || 1;
    const minValue = -maxAbs;
    const maxValue = maxAbs;
    const range = maxValue - minValue || 1;
    const zeroX = padding.left + ((0 - minValue) / range) * chartWidth;
    const valueFormat = rows.some(function (row) {
      return row.valueFormat === "percent";
    }) ? "percent" : "number";

    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);

    context.fillStyle = "#18212F";
    context.font = "700 28px Inter, Segoe UI, Arial, sans-serif";
    context.fillText("Динамика показателя", 32, 42);

    context.fillStyle = "#667085";
    context.font = "600 16px Inter, Segoe UI, Arial, sans-serif";
    context.fillText(metric.label || "Показатель", 32, 68);

    drawRoundedRect(context, padding.left, padding.top - 20, chartWidth, chartHeight + 40, 22, "#F8FAFC");

    context.strokeStyle = "#CBD5E1";
    context.lineWidth = 1;
    context.beginPath();
    context.moveTo(zeroX, padding.top - 10);
    context.lineTo(zeroX, padding.top + chartHeight + 10);
    context.stroke();

    rows.forEach(function (row, index) {
      const y = padding.top + index * (barHeight + rowGap);
      const valueX = padding.left + ((row.delta - minValue) / range) * chartWidth;
      const x = Math.min(zeroX, valueX);
      const barWidth = Math.max(4, Math.abs(valueX - zeroX));
      const barColor = getExportBarColor(row.impact);

      context.fillStyle = "#667085";
      context.font = "700 13px Inter, Segoe UI, Arial, sans-serif";
      context.textAlign = "right";
      context.textBaseline = "middle";
      context.fillText(trimText(context, row.label, 230), padding.left - 18, y + barHeight / 2);

      drawRoundedRect(context, x, y, barWidth, barHeight, 8, barColor);

      context.font = "800 13px Inter, Segoe UI, Arial, sans-serif";
      drawValueLabel(context, {
        text: Normalizers.formatMetricDelta(row.delta, valueFormat, 2),
        impact: row.impact,
        delta: row.delta,
        x: x,
        y: y + barHeight / 2,
        barWidth: barWidth,
        chartLeft: padding.left + 12,
        chartRight: padding.left + chartWidth - 12,
      });
    });

    context.fillStyle = "#94A3B8";
    context.font = "600 12px Inter, Segoe UI, Arial, sans-serif";
    context.textAlign = "left";
    context.fillText("Топ изменений по модулю значения. График сформирован локально в браузере.", 32, height - 26);
  }

  function drawValueLabel(context, options) {
    const textWidth = context.measureText(options.text).width;
    const defaultColor = options.impact === "bad" ? "#E11D48" : options.impact === "neutral" ? "#667085" : "#18212F";
    let textX;
    let textAlign;
    let textColor = defaultColor;

    if (options.delta < 0) {
      const outsideX = options.x - 8;

      if (outsideX - textWidth >= options.chartLeft) {
        textX = outsideX;
        textAlign = "right";
      } else if (options.barWidth >= textWidth + 18) {
        textX = options.x + 9;
        textAlign = "left";
        textColor = "#FFFFFF";
      } else {
        textX = options.chartLeft;
        textAlign = "left";
      }
    } else {
      const outsideX = options.x + options.barWidth + 8;

      if (outsideX + textWidth <= options.chartRight) {
        textX = outsideX;
        textAlign = "left";
      } else if (options.barWidth >= textWidth + 18) {
        textX = options.x + options.barWidth - 9;
        textAlign = "right";
      } else {
        textX = options.chartRight;
        textAlign = "right";
      }
    }

    context.fillStyle = textColor;
    context.textAlign = textAlign;
    context.fillText(options.text, textX, options.y);
  }

  function drawRoundedRect(context, x, y, width, height, radius, color) {
    const safeRadius = Math.min(radius, Math.abs(width) / 2, Math.abs(height) / 2);

    context.fillStyle = color;
    context.beginPath();
    context.moveTo(x + safeRadius, y);
    context.lineTo(x + width - safeRadius, y);
    context.quadraticCurveTo(x + width, y, x + width, y + safeRadius);
    context.lineTo(x + width, y + height - safeRadius);
    context.quadraticCurveTo(x + width, y + height, x + width - safeRadius, y + height);
    context.lineTo(x + safeRadius, y + height);
    context.quadraticCurveTo(x, y + height, x, y + height - safeRadius);
    context.lineTo(x, y + safeRadius);
    context.quadraticCurveTo(x, y, x + safeRadius, y);
    context.fill();
  }

  function trimText(context, text, maxWidth) {
    const value = String(text || "");

    if (context.measureText(value).width <= maxWidth) {
      return value;
    }

    let output = value;
    while (output.length > 1 && context.measureText(output + "...").width > maxWidth) {
      output = output.slice(0, -1);
    }

    return output + "...";
  }

  function getExportBarColor(impact) {
    if (impact === "bad") {
      return "#E11D48";
    }

    if (impact === "neutral") {
      return "#94A3B8";
    }

    return "#16A34A";
  }

  function styleHeaderRow(sheet, rowNumber) {
    const header = sheet.getRow(rowNumber || 1);
    header.font = { bold: true, color: { argb: "FF18212F" } };
    header.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFCCFBF1" },
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
