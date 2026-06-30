(function (global) {
  const App = (global.Metricum = global.Metricum || {});
  const Normalizers = App.Normalizers;
  let deltaChart = null;
  let chartObserver = null;
  const CHART_TYPES = [
    { id: "bar-horizontal", label: "Динамика по объектам", title: "Динамика по объектам" },
    { id: "bar-vertical", label: "Вертикальные столбцы", title: "Топ изменений" },
    { id: "line-trend", label: "Тренд по периодам", title: "Тренд по периодам" },
    { id: "doughnut-impact", label: "Структура изменений", title: "Структура изменений" },
  ];

  function renderDeltaChart(canvas, comparison, metric, chartType) {
    if (!canvas || !global.Chart) {
      return;
    }

    ensureTooltipPositioner();
    destroyChart();

    if (!comparison || !metric) {
      clearCanvas(canvas);
      return;
    }

    const normalizedChartType = normalizeChartType(chartType);

    if (normalizedChartType === "line-trend") {
      renderTrendChart(canvas, comparison, metric);
      return;
    }

    if (normalizedChartType === "doughnut-impact") {
      renderImpactChart(canvas, comparison, metric);
      return;
    }

    renderBarChart(canvas, comparison, metric, normalizedChartType);
  }

  function renderBarChart(canvas, comparison, metric, chartType) {
    const firstPeriod = comparison.periods[0];
    const lastPeriod = comparison.periods[comparison.periods.length - 1];
    const isPairwise = comparison.comparisonMode === "sequential" || comparison.comparisonMode === "manual";
    const isVertical = chartType === "bar-vertical";
    const rows = buildDeltaRows(comparison, metric)
      .sort(function (left, right) {
        return Math.abs(right.delta) - Math.abs(left.delta);
      })
      .slice(0, isVertical ? 10 : 15);
    const displayRows = isVertical ? rows : rows.slice().reverse();

    if (!displayRows.length) {
      clearCanvas(canvas);
      return;
    }

    const displayData = displayRows.map(function (row) {
      return row.delta;
    });
    const chartValueFormat = getChartValueFormat(displayRows);
    const waitForViewport = shouldWaitForViewport(canvas, displayData);
    const themeColors = getThemeColors();

    deltaChart = new global.Chart(canvas, {
      type: "bar",
      data: {
        labels: displayRows.map(function (row) {
          return row.label;
        }),
        datasets: [
          {
            label: metric.label + " " + (isPairwise ? "динамика по выбранным парам" : lastPeriod.label + " - " + firstPeriod.label),
            data: waitForViewport ? displayData.map(function () { return 0; }) : displayData,
            backgroundColor: function (context) {
              const chart = context.chart;
              const chartArea = chart.chartArea;
              const row = displayRows[context.dataIndex];

              if (!chartArea || !row) {
                return "rgba(20, 184, 166, 0.86)";
              }

              if (row.impact === "bad") {
                return createHorizontalGradient(chart.ctx, chartArea, "#FDA4AF", "#E11D48");
              }

              if (row.impact === "neutral") {
                return createHorizontalGradient(chart.ctx, chartArea, "#E2E8F0", "#94A3B8");
              }

              return createHorizontalGradient(chart.ctx, chartArea, "#86EFAC", "#16A34A");
            },
            borderRadius: 14,
            borderSkipped: false,
            barThickness: 18,
            maxBarThickness: 22,
            borderWidth: 0,
          },
        ],
      },
      options: {
        indexAxis: isVertical ? "x" : "y",
        interaction: {
          mode: isVertical ? "nearest" : "barHitbox",
          axis: "xy",
          intersect: !isVertical,
        },
        hover: {
          mode: isVertical ? "nearest" : "barHitbox",
          axis: "xy",
          intersect: !isVertical,
        },
        responsive: true,
        maintainAspectRatio: false,
        layout: {
          padding: {
            top: 18,
            right: 24,
            bottom: 8,
            left: 8,
          },
        },
        plugins: {
          legend: {
            display: false,
          },
          tooltip: {
            position: isVertical ? "nearest" : "cursorWithinBar",
            backgroundColor: themeColors.tooltipBackground,
            borderColor: themeColors.tooltipBorder,
            borderWidth: 1,
            caretPadding: 10,
            cornerRadius: 14,
            displayColors: false,
            titleColor: themeColors.tooltipText,
            bodyColor: themeColors.tooltipText,
            padding: 12,
            titleFont: {
              family: "Inter, system-ui, sans-serif",
              size: 13,
              weight: "700",
            },
            bodyFont: {
              family: "Inter, system-ui, sans-serif",
              size: 14,
              weight: "800",
            },
            callbacks: {
              title: function (items) {
                const row = items[0] ? displayRows[items[0].dataIndex] : null;
                return row ? row.unitLabel : "";
              },
              label: function (context) {
                const parsedValue = isVertical ? context.parsed.y : context.parsed.x;
                return "Изменение: " + Normalizers.formatMetricDelta(parsedValue, chartValueFormat, 2);
              },
            },
          },
        },
        scales: {
          x: {
            border: {
              display: false,
            },
            grid: {
              display: false,
            },
            ticks: {
              color: themeColors.axisMuted,
              padding: 8,
              callback: function (value) {
                if (isVertical) {
                  const label = this.getLabelForValue ? this.getLabelForValue(value) : value;
                  return trimLabel(label, 14);
                }

                return Normalizers.formatMetricDelta(Number(value), chartValueFormat, 1);
              },
              font: {
                family: "Inter, system-ui, sans-serif",
                size: isVertical ? 10 : 12,
                weight: "700",
              },
              maxRotation: isVertical ? 0 : 50,
              minRotation: 0,
            },
          },
          y: {
            border: {
              display: false,
            },
            grid: {
              display: false,
            },
            ticks: {
              autoSkip: isVertical,
              color: themeColors.axisText,
              padding: 10,
              callback: function (value) {
                if (isVertical) {
                  return Normalizers.formatMetricDelta(Number(value), chartValueFormat, 1);
                }

                return this.getLabelForValue ? this.getLabelForValue(value) : value;
              },
              font: {
                family: "Inter, system-ui, sans-serif",
                size: 12,
                weight: "700",
              },
            },
          },
        },
        animation: {
          duration: waitForViewport ? 0 : 460,
          easing: "easeOutQuart",
        },
      },
    });

    if (waitForViewport) {
      animateWhenVisible(canvas, displayData);
    }
  }

  function renderTrendChart(canvas, comparison, metric) {
    const trendRows = buildTrendRows(comparison, metric);

    if (!trendRows.length) {
      clearCanvas(canvas);
      return;
    }

    const themeColors = getThemeColors();
    const chartValueFormat = trendRows[0].valueFormat || "number";
    const waitForViewport = shouldWaitForViewport(canvas, trendRows.map(function (row) { return row.value; }));

    deltaChart = new global.Chart(canvas, {
      type: "line",
      data: {
        labels: trendRows.map(function (row) {
          return row.label;
        }),
        datasets: [
          {
            label: metric.label,
            data: waitForViewport
              ? trendRows.map(function () { return null; })
              : trendRows.map(function (row) { return row.value; }),
            borderColor: "#14B8A6",
            backgroundColor: "rgba(20, 184, 166, 0.14)",
            pointBackgroundColor: "#14B8A6",
            pointBorderColor: "#FFFFFF",
            pointBorderWidth: 2,
            pointRadius: 4,
            pointHoverRadius: 6,
            borderWidth: 3,
            fill: true,
            tension: 0.36,
            spanGaps: true,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
          mode: "nearest",
          intersect: false,
        },
        layout: {
          padding: {
            top: 18,
            right: 24,
            bottom: 8,
            left: 8,
          },
        },
        plugins: {
          legend: {
            display: false,
          },
          tooltip: {
            backgroundColor: themeColors.tooltipBackground,
            borderColor: themeColors.tooltipBorder,
            borderWidth: 1,
            cornerRadius: 14,
            displayColors: false,
            titleColor: themeColors.tooltipText,
            bodyColor: themeColors.tooltipText,
            padding: 12,
            callbacks: {
              label: function (context) {
                return metric.label + ": " + Normalizers.formatMetricValue(context.parsed.y, chartValueFormat, 2);
              },
            },
          },
        },
        scales: {
          x: {
            border: { display: false },
            grid: { display: false },
            ticks: {
              color: themeColors.axisText,
              font: {
                family: "Inter, system-ui, sans-serif",
                size: 12,
                weight: "700",
              },
            },
          },
          y: {
            border: { display: false },
            grid: {
              color: themeColors.grid,
            },
            ticks: {
              color: themeColors.axisMuted,
              callback: function (value) {
                return Normalizers.formatMetricValue(Number(value), chartValueFormat, 1);
              },
              font: {
                family: "Inter, system-ui, sans-serif",
                size: 12,
                weight: "700",
              },
            },
          },
        },
        animation: {
          duration: waitForViewport ? 0 : 460,
          easing: "easeOutQuart",
        },
      },
    });

    if (waitForViewport) {
      animateWhenVisible(canvas, trendRows.map(function (row) { return row.value; }));
    }
  }

  function renderImpactChart(canvas, comparison, metric) {
    const structure = buildImpactStructure(comparison, metric);
    const total = structure.reduce(function (sum, item) {
      return sum + item.value;
    }, 0);

    if (!total) {
      clearCanvas(canvas);
      return;
    }

    const themeColors = getThemeColors();

    deltaChart = new global.Chart(canvas, {
      type: "doughnut",
      data: {
        labels: structure.map(function (item) {
          return item.label;
        }),
        datasets: [
          {
            data: structure.map(function (item) {
              return item.value;
            }),
            backgroundColor: structure.map(function (item) {
              return item.color;
            }),
            borderColor: themeColors.surface,
            borderWidth: 5,
            hoverOffset: 6,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: "64%",
        layout: {
          padding: {
            top: 16,
            right: 24,
            bottom: 16,
            left: 24,
          },
        },
        plugins: {
          legend: {
            position: "bottom",
            labels: {
              color: themeColors.axisText,
              boxWidth: 10,
              boxHeight: 10,
              useBorderRadius: true,
              borderRadius: 999,
              font: {
                family: "Inter, system-ui, sans-serif",
                size: 12,
                weight: "800",
              },
            },
          },
          tooltip: {
            backgroundColor: themeColors.tooltipBackground,
            borderColor: themeColors.tooltipBorder,
            borderWidth: 1,
            cornerRadius: 14,
            displayColors: false,
            titleColor: themeColors.tooltipText,
            bodyColor: themeColors.tooltipText,
            padding: 12,
            callbacks: {
              label: function (context) {
                const value = Number(context.parsed) || 0;
                const percent = total ? Math.round((value / total) * 100) : 0;
                return value + " (" + percent + "%)";
              },
            },
          },
        },
        animation: {
          duration: 460,
          easing: "easeOutQuart",
        },
      },
    });
  }

  function buildDeltaRows(comparison, metric) {
    const firstPeriod = comparison.periods[0];
    const lastPeriod = comparison.periods[comparison.periods.length - 1];
    const isPairwise = comparison.comparisonMode === "sequential" || comparison.comparisonMode === "manual";

    return comparison.rows.flatMap(function (row) {
      const result = findMetricResult(row, metric);

      if (!result) {
        return [];
      }

      if (isPairwise) {
        return result.comparisons
          .filter(function (item) {
            return Number.isFinite(item.delta);
          })
          .map(function (item) {
            return {
              label: row.label + " · " + item.label,
              unitLabel: row.label,
              delta: item.delta,
              impact: item.impact,
              valueFormat: item.valueFormat || result.valueFormat || "number",
              comparisonLabel: item.label,
            };
          });
      }

      if (!Number.isFinite(result.delta)) {
        return [];
      }

      return [{
        label: row.label,
        unitLabel: row.label,
        delta: result.delta,
        impact: result.impact,
        valueFormat: result.valueFormat || "number",
        comparisonLabel: lastPeriod.label + " - " + firstPeriod.label,
      }];
    });
  }

  function buildTrendRows(comparison, metric) {
    return comparison.periods
      .map(function (period) {
        const values = comparison.rows
          .map(function (row) {
            const result = findMetricResult(row, metric);
            const periodValue = result && result.periodValues
              ? result.periodValues.find(function (item) {
                  return item.periodId === period.id;
                })
              : null;

            return periodValue && Number.isFinite(periodValue.value) ? periodValue.value : null;
          })
          .filter(function (value) {
            return Number.isFinite(value);
          });
        const valueFormat = getMetricValueFormat(comparison, metric);

        if (!values.length) {
          return null;
        }

        return {
          label: period.label,
          value: aggregateTrendValues(values, metric, valueFormat),
          valueFormat: valueFormat,
        };
      })
      .filter(Boolean);
  }

  function aggregateTrendValues(values, metric, valueFormat) {
    const method = metric && metric.aggregation ? metric.aggregation : "auto";

    if (method === "avg" || method === "first" || valueFormat === "percent") {
      return values.reduce(sumValues, 0) / values.length;
    }

    if (method === "min") {
      return Math.min.apply(null, values);
    }

    if (method === "max") {
      return Math.max.apply(null, values);
    }

    return values.reduce(sumValues, 0);
  }

  function buildImpactStructure(comparison, metric) {
    const rows = buildDeltaRows(comparison, metric);
    const improved = rows.filter(function (row) {
      return row.impact === "good";
    }).length;
    const declined = rows.filter(function (row) {
      return row.impact === "bad";
    }).length;
    const unchanged = rows.filter(function (row) {
      return row.impact === "neutral";
    }).length;

    return [
      { label: "Улучшения", value: improved, color: "#16A34A" },
      { label: "Ухудшения", value: declined, color: "#E11D48" },
      { label: "Без изменений", value: unchanged, color: "#94A3B8" },
    ].filter(function (item) {
      return item.value > 0;
    });
  }

  function findMetricResult(row, metric) {
    if (!row || !metric) {
      return null;
    }

    return row.metrics.find(function (item) {
      return item.metricId === metric.id;
    }) || null;
  }

  function getMetricValueFormat(comparison, metric) {
    const resultRow = comparison.rows.find(function (row) {
      return findMetricResult(row, metric);
    });
    const result = resultRow ? findMetricResult(resultRow, metric) : null;

    return result ? result.valueFormat || "number" : "number";
  }

  function sumValues(sum, value) {
    return sum + value;
  }

  function normalizeChartType(chartType) {
    return CHART_TYPES.some(function (type) {
      return type.id === chartType;
    })
      ? chartType
      : "bar-horizontal";
  }

  function getChartTypes() {
    return CHART_TYPES.map(function (type) {
      return {
        id: type.id,
        label: type.label,
      };
    });
  }

  function getChartTitle(chartType) {
    const type = CHART_TYPES.find(function (item) {
      return item.id === chartType;
    });

    return type ? type.title : CHART_TYPES[0].title;
  }

  function trimLabel(value, maxLength) {
    const text = String(value === null || value === undefined ? "" : value);

    if (text.length <= maxLength) {
      return text;
    }

    return text.slice(0, Math.max(1, maxLength - 1)) + "…";
  }

  function createHorizontalGradient(context, chartArea, fromColor, toColor) {
    const gradient = context.createLinearGradient(chartArea.left, 0, chartArea.right, 0);
    gradient.addColorStop(0, fromColor);
    gradient.addColorStop(1, toColor);
    return gradient;
  }

  function ensureTooltipPositioner() {
    registerBarHitboxMode();

    const Tooltip = global.Chart && global.Chart.Tooltip;

    if (!Tooltip || !Tooltip.positioners || Tooltip.positioners.cursorWithinBar) {
      return;
    }

    Tooltip.positioners.cursorWithinBar = function (elements, eventPosition) {
      if (!elements.length) {
        return false;
      }

      const activeElement = elements[0];
      const element = activeElement.element || activeElement;
      const bounds = getBarHitbox(element, this.chart ? this.chart.chartArea : null, true);
      const pointer = eventPosition || { x: bounds.left, y: bounds.centerY };

      return {
        x: clamp(pointer.x, bounds.left, bounds.right),
        y: bounds.centerY,
      };
    };
  }

  function registerBarHitboxMode() {
    const Interaction = global.Chart && global.Chart.Interaction;

    if (!Interaction || !Interaction.modes || Interaction.modes.barHitbox) {
      return;
    }

    Interaction.modes.barHitbox = function (chart, event) {
      const pointer = getEventPosition(chart, event);
      const chartArea = chart.chartArea;
      const activeItems = [];
      const horizontalPadding = 4;
      const verticalPadding = 6;

      if (!pointer || !chartArea) {
        return activeItems;
      }

      (chart.data.datasets || []).forEach(function (_dataset, datasetIndex) {
        const meta = chart.getDatasetMeta(datasetIndex);

        if (!meta || meta.hidden || !meta.data) {
          return;
        }

        meta.data.forEach(function (element, index) {
          if (!element || (element.hasValue && !element.hasValue())) {
            return;
          }

          const bounds = getBarHitbox(element, chartArea, true);
          const isInside =
            pointer.x >= bounds.left - horizontalPadding &&
            pointer.x <= bounds.right + horizontalPadding &&
            pointer.y >= bounds.top - verticalPadding &&
            pointer.y <= bounds.bottom + verticalPadding;

          if (isInside) {
            activeItems.push({
              element,
              datasetIndex,
              index,
            });
          }
        });
      });

      return activeItems;
    };
  }

  function getEventPosition(chart, event) {
    const nativeEvent = event && (event.native || event);

    if (chart && chart.canvas && nativeEvent) {
      const pointer = getNativePointer(nativeEvent);

      if (pointer) {
        const rect = chart.canvas.getBoundingClientRect();
        const scaleX = rect.width ? chart.width / rect.width : 1;
        const scaleY = rect.height ? chart.height / rect.height : 1;

        return {
          x: (pointer.clientX - rect.left) * scaleX,
          y: (pointer.clientY - rect.top) * scaleY,
        };
      }
    }

    if (event && Number.isFinite(event.x) && Number.isFinite(event.y)) {
      return {
        x: event.x,
        y: event.y,
      };
    }

    return null;
  }

  function getNativePointer(nativeEvent) {
    if (Number.isFinite(nativeEvent.clientX) && Number.isFinite(nativeEvent.clientY)) {
      return nativeEvent;
    }

    const touch = nativeEvent.touches && nativeEvent.touches[0];

    if (touch && Number.isFinite(touch.clientX) && Number.isFinite(touch.clientY)) {
      return touch;
    }

    return null;
  }

  function getBarHitbox(element, chartArea, useFinalPosition) {
    const props = element.getProps ? element.getProps(["x", "y", "base", "height"], useFinalPosition) : element;
    const pointX = Number.isFinite(props.x) ? props.x : 0;
    const pointY = Number.isFinite(props.y) ? props.y : 0;
    const baseX = Number.isFinite(props.base) ? props.base : pointX;
    const rawLeft = Math.min(pointX, baseX);
    const rawRight = Math.max(pointX, baseX);
    const barCenter = (rawLeft + rawRight) / 2;
    const minSpan = 18;
    const span = Math.max(rawRight - rawLeft, minSpan);
    const barLeft = barCenter - span / 2;
    const barRight = barCenter + span / 2;
    const safeLeft = chartArea ? Math.max(barLeft, chartArea.left) : barLeft;
    const safeRight = chartArea ? Math.min(barRight, chartArea.right) : barRight;
    const halfHeight = Math.max((props.height || 18) / 2, 9);

    return {
      left: safeLeft,
      right: safeRight,
      top: pointY - halfHeight,
      bottom: pointY + halfHeight,
      centerY: pointY,
    };
  }

  function clamp(value, min, max) {
    if (!Number.isFinite(value)) {
      return min;
    }

    if (min > max) {
      return (min + max) / 2;
    }

    return Math.min(Math.max(value, min), max);
  }

  function clearCanvas(canvas) {
    const context = canvas.getContext("2d");
    context.clearRect(0, 0, canvas.width, canvas.height);
  }

  function destroyChart() {
    disconnectChartObserver();

    if (deltaChart) {
      deltaChart.destroy();
      deltaChart = null;
    }
  }

  function shouldWaitForViewport(canvas, targetData) {
    return Boolean(
      targetData.length &&
        !prefersReducedMotion() &&
        "IntersectionObserver" in global &&
        !isNearViewport(canvas)
    );
  }

  function animateWhenVisible(canvas, targetData) {
    const target = canvas.closest(".chart-panel") || canvas;

    chartObserver = new IntersectionObserver(function (entries) {
      const isVisible = entries.some(function (entry) {
        return entry.isIntersecting;
      });

      if (!isVisible || !deltaChart || deltaChart.canvas !== canvas) {
        return;
      }

      deltaChart.data.datasets[0].data = targetData;
      deltaChart.options.animation.duration = 460;
      deltaChart.options.animation.easing = "easeOutQuart";
      deltaChart.update();
      disconnectChartObserver();
    }, {
      root: null,
      rootMargin: "0px 0px -12% 0px",
      threshold: 0.22,
    });

    chartObserver.observe(target);
  }

  function disconnectChartObserver() {
    if (chartObserver) {
      chartObserver.disconnect();
      chartObserver = null;
    }
  }

  function isNearViewport(element) {
    const rect = element.getBoundingClientRect();
    const viewportHeight = global.innerHeight || document.documentElement.clientHeight;

    return rect.top < viewportHeight * 0.86 && rect.bottom > viewportHeight * 0.08;
  }

  function prefersReducedMotion() {
    return Boolean(global.matchMedia && global.matchMedia("(prefers-reduced-motion: reduce)").matches);
  }

  function getChartValueFormat(rows) {
    const row = rows.find(function (item) {
      return item.valueFormat === "percent";
    });

    return row ? "percent" : "number";
  }

  function getThemeColors() {
    const styles = global.getComputedStyle(document.documentElement);
    const text = styles.getPropertyValue("--text").trim() || "#18212F";
    const secondary = styles.getPropertyValue("--text-secondary").trim() || "#667085";
    const neutral = styles.getPropertyValue("--neutral").trim() || "#94A3B8";
    const isDark = document.documentElement.classList.contains("theme-dark");

    return {
      axisText: secondary,
      axisMuted: neutral,
      grid: isDark ? "rgba(148, 163, 184, 0.16)" : "rgba(226, 232, 240, 0.74)",
      surface: isDark ? "#171D26" : "#FFFFFF",
      tooltipBackground: isDark ? "#F8FAFC" : text,
      tooltipBorder: isDark ? "rgba(15, 23, 42, 0.12)" : "rgba(255, 255, 255, 0.08)",
      tooltipText: isDark ? "#18212F" : "#FFFFFF",
    };
  }

  App.Charts = {
    renderDeltaChart,
    getChartTypes,
    getChartTitle,
  };
})(window);
