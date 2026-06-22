(function (global) {
  const App = (global.Metricum = global.Metricum || {});
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
    const isPairwise = comparison.comparisonMode === "sequential" || comparison.comparisonMode === "manual";
    const rows = comparison.rows
      .flatMap(function (row) {
        const result = row.metrics.find(function (item) {
          return item.metricId === metric.id;
        });

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
    const themeColors = getThemeColors();

    deltaChart = new global.Chart(canvas, {
      type: "bar",
      data: {
        labels: rows.map(function (row) {
          return row.label;
        }),
        datasets: [
          {
            label: metric.label + " " + (isPairwise ? "динамика по выбранным парам" : lastPeriod.label + " - " + firstPeriod.label),
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
          mode: "barHitbox",
          axis: "xy",
          intersect: true,
        },
        hover: {
          mode: "barHitbox",
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
              color: themeColors.axisMuted,
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
              color: themeColors.axisText,
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
      tooltipBackground: isDark ? "#F8FAFC" : text,
      tooltipBorder: isDark ? "rgba(15, 23, 42, 0.12)" : "rgba(255, 255, 255, 0.08)",
      tooltipText: isDark ? "#18212F" : "#FFFFFF",
    };
  }

  App.Charts = {
    renderDeltaChart,
  };
})(window);
