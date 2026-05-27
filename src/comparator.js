(function (global) {
  const App = (global.OperationalAnalytics = global.OperationalAnalytics || {});
  const Normalizers = App.Normalizers;

  function comparePeriods(options) {
    const periods = options.periods;
    const metrics = options.metrics;
    const comparisonMode = options.comparisonMode || "endpoint";
    const comparisonPairs = buildComparisonPairs(periods, comparisonMode);
    const metricProfiles = buildMetricProfiles(periods, metrics);
    const periodIndexes = periods.map(function (period) {
      return {
        period: period,
        index: buildIndex(period.table, period.idColumn),
      };
    });
    const keys = collectKeys(periodIndexes);
    const rows = [];
    const invalidValues = [];
    const missingByPeriod = periodIndexes.map(function (periodIndex) {
      return {
        periodId: periodIndex.period.id,
        periodLabel: periodIndex.period.label,
        items: [],
      };
    });

    keys.forEach(function (key) {
      const records = periodIndexes.map(function (periodIndex) {
        return periodIndex.index.records.get(key) || null;
      });
      const label = firstPresentText.apply(
        null,
        records
          .filter(Boolean)
          .map(function (record) {
            return record.displayValue;
          })
          .concat([key])
      );

      records.forEach(function (record, periodIndex) {
        if (!record) {
          missingByPeriod[periodIndex].items.push({
            key: key,
            label: label,
          });
        }
      });

      const metricResults = metrics.map(function (metric) {
        return buildMetricResult(metric, periods, records, key, label, invalidValues, comparisonMode, metricProfiles.get(metric.id));
      });

      rows.push({
        key: key,
        label: label,
        isComplete: records.every(Boolean),
        records: records,
        metrics: metricResults,
      });
    });

    rows.sort(function (left, right) {
      return left.label.localeCompare(right.label, "ru");
    });

    return {
      periods: periods.map(function (period) {
        return {
          id: period.id,
          label: period.label,
          fileName: period.file ? period.file.name : "",
        };
      }),
      comparisonMode: comparisonMode,
      comparisonPairs: comparisonPairs,
      rows: rows,
      missingByPeriod: missingByPeriod,
      duplicatesByPeriod: periodIndexes.map(function (periodIndex) {
        return {
          periodId: periodIndex.period.id,
          periodLabel: periodIndex.period.label,
          items: periodIndex.index.duplicates,
        };
      }),
      emptyIdsByPeriod: periodIndexes.map(function (periodIndex) {
        return {
          periodId: periodIndex.period.id,
          periodLabel: periodIndex.period.label,
          items: periodIndex.index.emptyIds,
        };
      }),
      invalidValues: invalidValues,
    };
  }

  function buildComparisonPairs(periods, comparisonMode) {
    if (comparisonMode === "sequential") {
      return periods.slice(1).map(function (period, index) {
        const previous = periods[index];
        return {
          fromPeriodId: previous.id,
          fromPeriodLabel: previous.label,
          toPeriodId: period.id,
          toPeriodLabel: period.label,
          label: period.label + " - " + previous.label,
        };
      });
    }

    return [
      {
        fromPeriodId: periods[0].id,
        fromPeriodLabel: periods[0].label,
        toPeriodId: periods[periods.length - 1].id,
        toPeriodLabel: periods[periods.length - 1].label,
        label: periods[periods.length - 1].label + " - " + periods[0].label,
      },
    ];
  }

  function collectKeys(periodIndexes) {
    const keys = new Set();

    periodIndexes.forEach(function (periodIndex) {
      periodIndex.index.records.forEach(function (_record, key) {
        keys.add(key);
      });
    });

    return keys;
  }

  function buildIndex(table, idColumn) {
    const records = new Map();
    const duplicates = [];
    const emptyIds = [];

    table.rows.forEach(function (row) {
      const displayValue = row.values[idColumn];
      const key = Normalizers.normalizeKey(displayValue);

      if (!key) {
        emptyIds.push({
          rowNumber: row.rowNumber,
          value: displayValue,
        });
        return;
      }

      if (records.has(key)) {
        const existing = records.get(key);
        duplicates.push({
          key: key,
          label: firstPresentText(displayValue, existing.displayValue, key),
          rowNumbers: [existing.row.rowNumber, row.rowNumber],
        });
        return;
      }

      records.set(key, {
        key: key,
        displayValue: displayValue,
        row: row,
      });
    });

    return {
      records: records,
      duplicates: duplicates,
      emptyIds: emptyIds,
    };
  }

  function buildMetricResult(metric, periods, records, key, label, invalidValues, comparisonMode, metricProfile) {
    const profile = metricProfile || { valueFormat: "number", scale: 1 };
    const periodValues = periods.map(function (period, index) {
      const record = records[index];
      const columnId = metric.columns[period.id];

      if (!record) {
        return {
          periodId: period.id,
          periodLabel: period.label,
          value: null,
          raw: "",
          isMissing: true,
          isNumeric: false,
          valueFormat: profile.valueFormat,
        };
      }

      const raw = record.row.values[columnId];
      const number = Normalizers.normalizeNumber(raw);
      const normalizedValue = number.isNumeric ? number.value * profile.scale : number.value;

      if (!number.isNumeric) {
        invalidValues.push({
          key: key,
          label: label,
          periodId: period.id,
          periodLabel: period.label,
          metric: metric.label,
          value: raw,
        });
      }

      return {
        periodId: period.id,
        periodLabel: period.label,
        value: normalizedValue,
        raw: raw,
        isMissing: false,
        isNumeric: number.isNumeric,
        valueFormat: profile.valueFormat,
      };
    });
    const comparisons = buildMetricComparisons(periodValues, comparisonMode, profile.valueFormat);
    const primaryComparison = comparisons[comparisons.length - 1] || null;

    return {
      metricId: metric.id,
      label: metric.label,
      valueFormat: profile.valueFormat,
      periodValues: periodValues,
      comparisons: comparisons,
      valueA: primaryComparison ? primaryComparison.valueA : null,
      valueB: primaryComparison ? primaryComparison.valueB : null,
      rawA: primaryComparison ? primaryComparison.rawA : "",
      rawB: primaryComparison ? primaryComparison.rawB : "",
      delta: primaryComparison ? primaryComparison.delta : null,
      deltaPercent: primaryComparison ? primaryComparison.deltaPercent : null,
      impact: primaryComparison ? primaryComparison.impact : "unknown",
    };
  }

  function buildMetricComparisons(periodValues, comparisonMode, valueFormat) {
    if (comparisonMode === "sequential") {
      return periodValues.slice(1).map(function (current, index) {
        return buildSingleComparison(periodValues[index], current, valueFormat);
      });
    }

    return [buildSingleComparison(periodValues[0], periodValues[periodValues.length - 1], valueFormat)];
  }

  function buildSingleComparison(fromValue, toValue, valueFormat) {
    const hasNumbers = fromValue && toValue && fromValue.isNumeric && toValue.isNumeric;
    const delta = hasNumbers ? toValue.value - fromValue.value : null;
    const deltaPercent = hasNumbers ? Normalizers.calculatePercentChange(fromValue.value, toValue.value) : null;

    return {
      fromPeriodId: fromValue ? fromValue.periodId : "",
      fromPeriodLabel: fromValue ? fromValue.periodLabel : "",
      toPeriodId: toValue ? toValue.periodId : "",
      toPeriodLabel: toValue ? toValue.periodLabel : "",
      label: (toValue ? toValue.periodLabel : "") + " - " + (fromValue ? fromValue.periodLabel : ""),
      valueA: fromValue ? fromValue.value : null,
      valueB: toValue ? toValue.value : null,
      rawA: fromValue ? fromValue.raw : "",
      rawB: toValue ? toValue.raw : "",
      delta: delta,
      deltaPercent: deltaPercent,
      valueFormat: valueFormat || "number",
      impact: getImpact(delta, hasNumbers),
    };
  }

  function buildMetricProfiles(periods, metrics) {
    const profiles = new Map();

    metrics.forEach(function (metric) {
      profiles.set(metric.id, detectMetricProfile(periods, metric));
    });

    return profiles;
  }

  function detectMetricProfile(periods, metric) {
    let headerHints = 0;
    let rawPercentHints = 0;
    let numericCount = 0;
    let decimalRatioCount = 0;

    periods.forEach(function (period) {
      const columnId = metric.columns[period.id];
      const header = period.table.headers.find(function (item) {
        return item.id === columnId;
      });

      if (header && Normalizers.headerLooksPercent(header.name)) {
        headerHints += 1;
      }

      period.table.rows.forEach(function (row) {
        const raw = row.values[columnId];

        if (Normalizers.isEmptyValue(raw)) {
          return;
        }

        if (Normalizers.valueLooksPercent(raw)) {
          rawPercentHints += 1;
        }

        const number = Normalizers.normalizeNumber(raw);
        if (!number.isNumeric) {
          return;
        }

        numericCount += 1;

        if (!Normalizers.valueLooksPercent(raw) && Math.abs(number.value) > 0 && Math.abs(number.value) <= 1) {
          decimalRatioCount += 1;
        }
      });
    });

    const hasPercentHint = headerHints > 0 || rawPercentHints > 0;
    const decimalRatioShare = numericCount ? decimalRatioCount / numericCount : 0;
    const shouldScaleRatio = headerHints > 0 && rawPercentHints === 0 && decimalRatioShare >= 0.8;

    return {
      valueFormat: hasPercentHint ? "percent" : "number",
      scale: shouldScaleRatio ? 100 : 1,
    };
  }

  function firstPresentText() {
    for (let index = 0; index < arguments.length; index += 1) {
      const value = arguments[index];
      if (value !== null && value !== undefined && value !== "") {
        return value;
      }
    }

    return "";
  }

  function getImpact(delta, hasNumbers) {
    if (!hasNumbers || !Number.isFinite(delta)) {
      return "unknown";
    }

    if (delta === 0) {
      return "neutral";
    }

    return delta > 0 ? "good" : "bad";
  }

  App.Comparator = {
    comparePeriods,
  };
})(window);
