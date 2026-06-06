(function (global) {
  const App = (global.OperationalAnalytics = global.OperationalAnalytics || {});

  function createInitialState() {
    return {
      periods: [createPeriod(0), createPeriod(1)],
      comparisonMode: "endpoint",
      mapping: {
        metrics: [],
      },
      comparison: null,
      analytics: null,
      selectedChartMetricId: "",
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
      columns: columns || {},
    };
  }

  App.state = createInitialState();
  App.createInitialState = createInitialState;
  App.createGlobalFilters = createGlobalFilters;
  App.createPeriod = createPeriod;
  App.createMetric = createMetric;
})(window);
