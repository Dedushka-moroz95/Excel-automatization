(function (global) {
  const App = (global.Metricum = global.Metricum || {});

  function createInitialState() {
    return {
      periodSourceMode: "multiFile",
      singleFile: createSingleFileSource(),
      periods: [createPeriod(0), createPeriod(1)],
      comparisonMode: "endpoint",
      manualComparisonPairs: [],
      mapping: {
        metrics: [],
      },
      comparison: null,
      analytics: null,
      selectedChartMetricId: "",
      selectedChartType: "bar-horizontal",
      globalFilters: createGlobalFilters(),
      messages: [],
    };
  }

  function createGlobalFilters() {
    return {
      impact: "all",
      deltaMin: "",
      deltaMax: "",
      objectQuery: "",
      departmentQuery: "",
    };
  }

  function createSingleFileSource() {
    return {
      file: null,
      table: null,
      periodColumn: "",
      idColumn: "",
      virtualPeriods: [],
      warnings: [],
      loading: false,
    };
  }

  function createPeriod(index) {
    return {
      id: "period_" + Date.now() + "_" + index + "_" + Math.random().toString(16).slice(2),
      label: "Период " + (index + 1),
      file: null,
      table: null,
      idColumn: "",
    };
  }

  function createMetric(index, columns) {
    const number = index + 1;

    return {
      id: "metric_" + Date.now() + "_" + number,
      label: "",
      aggregation: "auto",
      columns: columns || {},
    };
  }

  App.state = createInitialState();
  App.createInitialState = createInitialState;
  App.createSingleFileSource = createSingleFileSource;
  App.createGlobalFilters = createGlobalFilters;
  App.createPeriod = createPeriod;
  App.createMetric = createMetric;
})(window);
