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
    const best = [];
    const worst = [];
    const showComparisonLabel = comparisonMode === "sequential" || comparisonMode === "manual";
    let improvedCount = 0;
    let declinedCount = 0;
    let unchangedCount = 0;
    let validCount = 0;
    let valueFormat = "number";

    rows.forEach(function (row) {
      const result = row.metrics.find(function (item) {
        return item.metricId === metric.id;
      });

      if (!result || !result.comparisons) {
        return;
      }

      result.comparisons.forEach(function (comparison) {
        if (!Number.isFinite(comparison.delta)) {
          return;
        }

        const item = {
          key: row.key,
          label: row.label,
          comparisonLabel: showComparisonLabel ? comparison.label : "",
          valueA: comparison.valueA,
          valueB: comparison.valueB,
          delta: comparison.delta,
          deltaPercent: comparison.deltaPercent,
          valueFormat: comparison.valueFormat || result.valueFormat || "number",
          impact: comparison.impact,
          score: comparison.delta,
        };

        validCount += 1;

        if (validCount === 1) {
          valueFormat = item.valueFormat;
        }

        if (item.impact === "good") {
          improvedCount += 1;
        } else if (item.impact === "bad") {
          declinedCount += 1;
        } else if (item.impact === "neutral") {
          unchangedCount += 1;
        }

        if (item.delta > 0) {
          pushTopItem(best, item, function (left, right) {
            return right.score - left.score;
          });
        } else if (item.delta < 0) {
          pushTopItem(worst, item, function (left, right) {
            return left.score - right.score;
          });
        }
      });
    });

    return {
      metricId: metric.id,
      label: metric.label,
      valueFormat: valueFormat,
      improvedCount: improvedCount,
      declinedCount: declinedCount,
      unchangedCount: unchangedCount,
      validCount: validCount,
      best: best,
      worst: worst,
    };
  }

  function pushTopItem(items, item, compare) {
    items.push(item);
    items.sort(compare);

    if (items.length > 5) {
      items.length = 5;
    }
  }

  App.Analytics = {
    buildAnalytics,
  };
})(window);
