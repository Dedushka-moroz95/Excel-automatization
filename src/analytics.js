(function (global) {
  const App = (global.Metricum = global.Metricum || {});

  function buildAnalytics(comparison, metrics) {
    const metricSummaries = metrics.map(function (metric) {
      return buildMetricSummary(comparison.rows, metric, comparison.comparisonMode);
    });

    return {
      periodCount: comparison.periods.length,
      totalUnits: comparison.rows.length,
      totalCompared: comparison.rows.filter(function (row) {
        return row.isComplete;
      }).length,
      missingTotal: sumNestedItems(comparison.missingByPeriod),
      duplicateIds: sumNestedItems(comparison.duplicatesByPeriod),
      invalidValues: comparison.invalidValues.length,
      metricSummaries: metricSummaries,
    };
  }

  function sumNestedItems(groups) {
    return groups.reduce(function (sum, group) {
      return sum + group.items.length;
    }, 0);
  }

  function buildMetricSummary(rows, metric, comparisonMode) {
    const items = rows
      .flatMap(function (row) {
        const result = row.metrics.find(function (item) {
          return item.metricId === metric.id;
        });

        if (!result || !result.comparisons) {
          return [];
        }

        return result.comparisons
          .filter(function (comparison) {
            return Number.isFinite(comparison.delta);
          })
          .map(function (comparison) {
            return {
              key: row.key,
              label: row.label,
              comparisonLabel: comparisonMode === "sequential" || comparisonMode === "manual" ? comparison.label : "",
              valueA: comparison.valueA,
              valueB: comparison.valueB,
              delta: comparison.delta,
              deltaPercent: comparison.deltaPercent,
              valueFormat: comparison.valueFormat || result.valueFormat || "number",
              impact: comparison.impact,
              score: comparison.delta,
            };
          });
      })
      .filter(Boolean);

    const improved = items.filter(function (item) {
      return item.impact === "good";
    });
    const declined = items.filter(function (item) {
      return item.impact === "bad";
    });
    const unchanged = items.filter(function (item) {
      return item.impact === "neutral";
    });

    const best = items
      .filter(function (item) {
        return item.delta > 0;
      })
      .slice()
      .sort(function (left, right) {
        return right.score - left.score;
      })
      .slice(0, 5);

    const worst = items
      .filter(function (item) {
        return item.delta < 0;
      })
      .slice()
      .sort(function (left, right) {
        return left.score - right.score;
      })
      .slice(0, 5);

    return {
      metricId: metric.id,
      label: metric.label,
      valueFormat: items[0] ? items[0].valueFormat : "number",
      improvedCount: improved.length,
      declinedCount: declined.length,
      unchangedCount: unchanged.length,
      validCount: items.length,
      best: best,
      worst: worst,
    };
  }

  App.Analytics = {
    buildAnalytics,
  };
})(window);
