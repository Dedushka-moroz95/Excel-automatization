(function (global) {
  const App = (global.OperationalAnalytics = global.OperationalAnalytics || {});
  const Normalizers = App.Normalizers;
  let deltaChart = null;
  let chartObserver = null;

  function renderDeltaChart(canvas, comparison, metric) {
    if (!canvas || !global.Chart) {
      return;
    }

    ensureTooltipPositioner();
    destroyChart();

    if (!comparison || !metric) {
      clearCanvas(canvas);
      return;
    }

    const firstPeriod = comparison.periods[0];
    const lastPeriod = comparison.periods[comparison.periods.length - 1];
    const isSequential = comparison.comparisonMode === "sequential";
    const rows = comparison.rows
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
      })
      .sort(function (left, right) {
        return Math.abs(right.delta) - Math.abs(left.delta);
      })
      .slice(0, 15)
      .reverse();
    const targetData = rows.map(function (row) {
      return row.delta;
    });
    const chartValueFormat = getChartValueFormat(rows);
    const waitForViewport = shouldWaitForViewport(canvas, targetData);

    deltaChart = new global.Chart(canvas, {
      type: "bar",
      data: {
        labels: rows.map(function (row) {
          return row.label;
        }),
        datasets: [
          {
            label: metric.label + " " + (isSequential ? "последовательная динамика" : lastPeriod.label + " - " + firstPeriod.label),
            data: waitForViewport ? targetData.map(function () { return 0; }) : targetData,
            backgroundColor: function (context) {
              const chart = context.chart;
              const chartArea = chart.chartArea;
              const row = rows[context.dataIndex];

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
        indexAxis: "y",
        interaction: {
          mode: "nearest",
          axis: "xy",
          intersect: true,
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
            position: "cursorWithinBar",
            backgroundColor: "#18212F",
            borderColor: "rgba(255, 255, 255, 0.08)",
            borderWidth: 1,
            caretPadding: 10,
            cornerRadius: 14,
            displayColors: false,
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
                const row = items[0] ? rows[items[0].dataIndex] : null;
                return row ? row.unitLabel : "";
              },
              label: function (context) {
                return "Изменение: " + Normalizers.formatMetricDelta(context.parsed.x, chartValueFormat, 2);
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
              color: "#94A3B8",
              padding: 8,
              callback: function (value) {
                return Normalizers.formatMetricDelta(Number(value), chartValueFormat, 1);
              },
              font: {
                family: "Inter, system-ui, sans-serif",
                size: 12,
                weight: "700",
              },
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
              autoSkip: false,
              color: "#667085",
              padding: 10,
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
      animateWhenVisible(canvas, targetData);
    }
  }

  function createHorizontalGradient(context, chartArea, fromColor, toColor) {
    const gradient = context.createLinearGradient(chartArea.left, 0, chartArea.right, 0);
    gradient.addColorStop(0, fromColor);
    gradient.addColorStop(1, toColor);
    return gradient;
  }

  function ensureTooltipPositioner() {
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
      const props = element.getProps ? element.getProps(["x", "y", "base", "height"], true) : element;
      const chartArea = this.chart ? this.chart.chartArea : null;
      const pointX = Number.isFinite(props.x) ? props.x : 0;
      const pointY = Number.isFinite(props.y) ? props.y : 0;
      const baseX = Number.isFinite(props.base) ? props.base : pointX;
      const pointer = eventPosition || { x: pointX, y: pointY };
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
        x: clamp(pointer.x, safeLeft, safeRight),
        y: clamp(pointer.y, pointY - halfHeight, pointY + halfHeight),
      };
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

  App.Charts = {
    renderDeltaChart,
  };
})(window);
