(function (global) {
  const App = (global.Metricum = global.Metricum || {});
  const Normalizers = App.Normalizers;
  const METRIC_PROFILE_SAMPLE_LIMIT = 5000;

  function comparePeriods(options) {
    const periods = options.periods;
    const metrics = options.metrics;
    const comparisonMode = options.comparisonMode || "endpoint";
    const comparisonPairs = buildComparisonPairs(periods, comparisonMode, options.comparisonPairs);
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
        return buildMetricResult(metric, periods, records, key, label, invalidValues, comparisonMode, comparisonPairs, metricProfiles.get(metric.id));
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

  function buildComparisonPairs(periods, comparisonMode, customPairs) {
    if (comparisonMode === "manual" && Array.isArray(customPairs) && customPairs.length) {
      return customPairs.map(function (pair) {
        const fromPeriod = periods.find(function (period) {
          return period.id === pair.fromPeriodId;
        });
        const toPeriod = periods.find(function (period) {
          return period.id === pair.toPeriodId;
        });

        return {
          fromPeriodId: pair.fromPeriodId,
          fromPeriodLabel: pair.fromPeriodLabel || (fromPeriod ? fromPeriod.label : ""),
          toPeriodId: pair.toPeriodId,
          toPeriodLabel: pair.toPeriodLabel || (toPeriod ? toPeriod.label : ""),
          label:
            pair.label ||
            ((pair.toPeriodLabel || (toPeriod ? toPeriod.label : "")) +
              " - " +
              (pair.fromPeriodLabel || (fromPeriod ? fromPeriod.label : ""))),
        };
      });
    }

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
    const duplicateGroups = new Map();
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
        let duplicateGroup = duplicateGroups.get(key);

        if (!duplicateGroup) {
          duplicateGroup = {
            key: key,
            label: firstPresentText(displayValue, existing.displayValue, key),
            rowNumbers: [existing.row.rowNumber],
          };
          duplicateGroups.set(key, duplicateGroup);
        }

        existing.rows.push(row);
        duplicateGroup.label = firstPresentText(displayValue, duplicateGroup.label, existing.displayValue, key);
        duplicateGroup.rowNumbers.push(row.rowNumber);
        return;
      }

      records.set(key, {
        key: key,
        displayValue: displayValue,
        row: row,
        rows: [row],
      });
    });

    return {
      records: records,
      duplicates: Array.from(duplicateGroups.values()).map(function (group) {
        return Object.assign({}, group, {
          count: group.rowNumbers.length,
        });
      }),
      emptyIds: emptyIds,
    };
  }

  function buildMetricResult(metric, periods, records, key, label, invalidValues, comparisonMode, comparisonPairs, metricProfile) {
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

      const aggregatedValue = aggregateMetricRows({
        rows: record.rows || [record.row],
        columnId: columnId,
        metric: metric,
        period: period,
        key: key,
        label: label,
        profile: profile,
        invalidValues: invalidValues,
      });

      return {
        periodId: period.id,
        periodLabel: period.label,
        value: aggregatedValue.value,
        raw: aggregatedValue.raw,
        isMissing: false,
        isNumeric: aggregatedValue.isNumeric,
        aggregation: aggregatedValue.aggregation,
        sourceCount: aggregatedValue.sourceCount,
        valueFormat: profile.valueFormat,
      };
    });
    const comparisons = buildMetricComparisons(periodValues, comparisonMode, profile.valueFormat, comparisonPairs);
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

  function aggregateMetricRows(options) {
    const rows = options.rows || [];
    const profile = options.profile;
    const method = resolveAggregation(options.metric.aggregation, profile.valueFormat);

    if (method === "first") {
      return aggregateFirstMetricValue(options);
    }

    let firstRaw = "";
    let hasFirstRaw = false;
    let numericCount = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;

    rows.forEach(function (row) {
      const raw = row.values[options.columnId];
      const number = Normalizers.normalizeNumber(raw);

      if (!hasFirstRaw && !Normalizers.isEmptyValue(raw)) {
        firstRaw = raw;
        hasFirstRaw = true;
      }

      if (number.isNumeric) {
        const value = number.value * profile.scale;
        numericCount += 1;
        sum += value;

        if (value < min) {
          min = value;
        }

        if (value > max) {
          max = value;
        }
        return;
      }

      if (!number.isEmpty) {
        pushInvalidMetricValue(options, raw);
      }
    });

    if (!numericCount) {
      return {
        value: null,
        raw: firstRaw,
        isNumeric: false,
        aggregation: method,
        sourceCount: rows.length,
      };
    }

    const value = aggregatePreparedNumbers({
      method: method,
      count: numericCount,
      sum: sum,
      min: min,
      max: max,
    });

    return {
      value: value,
      raw: Normalizers.formatMetricValue(value, profile.valueFormat, 2),
      isNumeric: Number.isFinite(value),
      aggregation: method,
      sourceCount: rows.length,
    };
  }

  function aggregateFirstMetricValue(options) {
    const row = options.rows[0];

    if (!row) {
      return {
        value: null,
        raw: "",
        isNumeric: false,
        aggregation: "first",
        sourceCount: 0,
      };
    }

    const raw = row.values[options.columnId];
    const number = Normalizers.normalizeNumber(raw);

    if (!number.isNumeric && !number.isEmpty) {
      pushInvalidMetricValue(options, raw);
    }

    return {
      value: number.isNumeric ? number.value * options.profile.scale : number.value,
      raw: raw,
      isNumeric: number.isNumeric,
      aggregation: "first",
      sourceCount: options.rows.length,
    };
  }

  function resolveAggregation(aggregation, valueFormat) {
    if (aggregation === "sum" || aggregation === "avg" || aggregation === "min" || aggregation === "max" || aggregation === "first") {
      return aggregation;
    }

    return valueFormat === "percent" ? "avg" : "sum";
  }

  function aggregatePreparedNumbers(options) {
    if (options.method === "avg") {
      return options.sum / options.count;
    }

    if (options.method === "min") {
      return options.min;
    }

    if (options.method === "max") {
      return options.max;
    }

    return options.sum;
  }

  function pushInvalidMetricValue(options, raw) {
    options.invalidValues.push({
      key: options.key,
      label: options.label,
      periodId: options.period.id,
      periodLabel: options.period.label,
      metric: options.metric.label,
      value: raw,
    });
  }

  function buildMetricComparisons(periodValues, comparisonMode, valueFormat, comparisonPairs) {
    if (comparisonMode === "manual") {
      const valuesByPeriod = new Map();

      periodValues.forEach(function (item) {
        valuesByPeriod.set(item.periodId, item);
      });

      return (comparisonPairs || [])
        .map(function (pair) {
          const fromValue = valuesByPeriod.get(pair.fromPeriodId);
          const toValue = valuesByPeriod.get(pair.toPeriodId);

          return buildSingleComparison(fromValue, toValue, valueFormat, pair.label);
        })
        .filter(Boolean);
    }

    if (comparisonMode === "sequential") {
      return periodValues.slice(1).map(function (current, index) {
        return buildSingleComparison(periodValues[index], current, valueFormat);
      });
    }

    return [buildSingleComparison(periodValues[0], periodValues[periodValues.length - 1], valueFormat)];
  }

  function buildSingleComparison(fromValue, toValue, valueFormat, label) {
    const hasNumbers = fromValue && toValue && fromValue.isNumeric && toValue.isNumeric;
    const delta = hasNumbers ? toValue.value - fromValue.value : null;
    const deltaPercent = hasNumbers ? Normalizers.calculatePercentChange(fromValue.value, toValue.value) : null;

    return {
      fromPeriodId: fromValue ? fromValue.periodId : "",
      fromPeriodLabel: fromValue ? fromValue.periodLabel : "",
      toPeriodId: toValue ? toValue.periodId : "",
      toPeriodLabel: toValue ? toValue.periodLabel : "",
      label: label || (toValue ? toValue.periodLabel : "") + " - " + (fromValue ? fromValue.periodLabel : ""),
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
    let scannedValues = 0;

    for (let periodIndex = 0; periodIndex < periods.length; periodIndex += 1) {
      const period = periods[periodIndex];
      const columnId = metric.columns[period.id];
      const header = findHeader(period.table.headers, columnId);

      if (header && Normalizers.headerLooksPercent(header.name)) {
        headerHints += 1;
      }
    }

    for (let periodIndex = 0; periodIndex < periods.length; periodIndex += 1) {
      const period = periods[periodIndex];
      const columnId = metric.columns[period.id];

      for (let rowIndex = 0; rowIndex < period.table.rows.length; rowIndex += 1) {
        if (scannedValues >= METRIC_PROFILE_SAMPLE_LIMIT) {
          break;
        }

        const row = period.table.rows[rowIndex];
        const raw = row.values[columnId];

        if (Normalizers.isEmptyValue(raw)) {
          continue;
        }

        scannedValues += 1;

        if (Normalizers.valueLooksPercent(raw)) {
          rawPercentHints += 1;
        }

        const number = Normalizers.normalizeNumber(raw);
        if (!number.isNumeric) {
          continue;
        }

        numericCount += 1;

        if (!Normalizers.valueLooksPercent(raw) && Math.abs(number.value) > 0 && Math.abs(number.value) <= 1) {
          decimalRatioCount += 1;
        }
      }

      if (scannedValues >= METRIC_PROFILE_SAMPLE_LIMIT) {
        break;
      }
    }

    const hasPercentHint = headerHints > 0 || rawPercentHints > 0;
    const decimalRatioShare = numericCount ? decimalRatioCount / numericCount : 0;
    const shouldScaleRatio = headerHints > 0 && rawPercentHints === 0 && decimalRatioShare >= 0.8;

    return {
      valueFormat: hasPercentHint ? "percent" : "number",
      scale: shouldScaleRatio ? 100 : 1,
    };
  }

  function findHeader(headers, columnId) {
    for (let index = 0; index < headers.length; index += 1) {
      if (headers[index].id === columnId) {
        return headers[index];
      }
    }

    return null;
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
