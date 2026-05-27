(function (global) {
  const App = (global.OperationalAnalytics = global.OperationalAnalytics || {});
  const Normalizers = App.Normalizers;
  let deltaChart = null;

  function renderDeltaChart(canvas, comparison, metric) {
    if (!canvas || !global.Chart) {
      return;
    }

    if (deltaChart) {
      deltaChart.destroy();
      deltaChart = null;
    }

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
                delta: item.delta,
                impact: item.impact,
                comparisonLabel: item.label,
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
          comparisonLabel: lastPeriod.label + " - " + firstPeriod.label,
        }];
      })
      .sort(function (left, right) {
        return Math.abs(right.delta) - Math.abs(left.delta);
      })
      .slice(0, 15)
      .reverse();

    deltaChart = new global.Chart(canvas, {
      type: "bar",
      data: {
        labels: rows.map(function (row) {
          return row.label;
        }),
        datasets: [
          {
            label: metric.label + " " + (isSequential ? "последовательная динамика" : lastPeriod.label + " - " + firstPeriod.label),
            data: rows.map(function (row) {
              return row.delta;
            }),
            backgroundColor: function (context) {
              const chart = context.chart;
              const chartArea = chart.chartArea;
              const row = rows[context.dataIndex];

              if (!chartArea || !row) {
                return "rgba(255, 221, 45, 0.85)";
              }

              if (row.impact === "bad") {
                return createHorizontalGradient(chart.ctx, chartArea, "#FCA5A5", "#DC2626");
              }

              if (row.impact === "neutral") {
                return createHorizontalGradient(chart.ctx, chartArea, "#E5E7EB", "#9CA3AF");
              }

              return createHorizontalGradient(chart.ctx, chartArea, "#FFE870", "#FFDD2D");
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
            backgroundColor: "#1F1F24",
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
              label: function (context) {
                const row = rows[context.dataIndex];
                return " " + (row ? row.comparisonLabel : "") + ": " + Normalizers.formatNumber(context.parsed.x, 2);
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
              color: "#9CA3AF",
              padding: 8,
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
              color: "#6B7280",
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
          duration: 650,
          easing: "easeOutQuart",
        },
      },
    });
  }

  function createHorizontalGradient(context, chartArea, fromColor, toColor) {
    const gradient = context.createLinearGradient(chartArea.left, 0, chartArea.right, 0);
    gradient.addColorStop(0, fromColor);
    gradient.addColorStop(1, toColor);
    return gradient;
  }

  function clearCanvas(canvas) {
    const context = canvas.getContext("2d");
    context.clearRect(0, 0, canvas.width, canvas.height);
  }

  App.Charts = {
    renderDeltaChart,
  };
})(window);
