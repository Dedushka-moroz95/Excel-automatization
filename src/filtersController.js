(function (global) {
  const App = (global.Metricum = global.Metricum || {});

  function create(options) {
    const state = options.state;
    const dom = options.dom;
    const createGlobalFilters = options.createGlobalFilters;
    const getSelectedMetric = options.getSelectedMetric;
    const buildAnalytics = options.buildAnalytics;
    const onChange = options.onChange;
    const escapeHtml = options.escapeHtml;
    let objectSuggestionIndex = -1;

    function init() {
      dom.globalObjectSearch.addEventListener("input", renderObjectSuggestions);
      dom.globalObjectSearch.addEventListener("focus", renderObjectSuggestions);
      dom.globalObjectSearch.addEventListener("keydown", handleObjectSuggestionKeydown);
      dom.globalObjectSuggestions.addEventListener("mousedown", handleObjectSuggestionSelect);
      dom.applyGlobalFiltersButton.addEventListener("click", apply);
      dom.resetGlobalFiltersButton.addEventListener("click", reset);

      document.addEventListener("click", function (event) {
        if (!event.target.closest(".autocomplete-field")) {
          closeSuggestions();
        }
      });
    }

    function apply() {
      ensureState();
      state.globalFilters.impact = dom.globalImpactFilter.value;
      state.globalFilters.deltaMin = dom.globalDeltaMinFilter.value;
      state.globalFilters.deltaMax = dom.globalDeltaMaxFilter.value;
      state.globalFilters.objectQuery = dom.globalObjectSearch.value;
      state.globalFilters.departmentQuery = dom.globalDepartmentSearch.value;
      closeSuggestions();
      onChange();
    }

    function reset() {
      resetState();
      closeSuggestions();
      onChange();
    }

    function resetState() {
      state.globalFilters = createGlobalFilters
        ? createGlobalFilters()
        : {
            impact: "all",
            deltaMin: "",
            deltaMax: "",
            objectQuery: "",
            departmentQuery: "",
          };
    }

    function ensureState() {
      if (!state.globalFilters) {
        resetState();
      }
    }

    function buildView() {
      ensureState();

      if (!state.comparison) {
        return {
          comparison: null,
          analytics: null,
          totalRows: 0,
          visibleRows: 0,
        };
      }

      const filteredComparison = buildFilteredComparison(state.comparison);

      return {
        comparison: filteredComparison,
        analytics: buildAnalytics(filteredComparison),
        totalRows: state.comparison.rows.length,
        visibleRows: filteredComparison.rows.length,
      };
    }

    function buildFilteredComparison(comparison) {
      if (!hasActive()) {
        return comparison;
      }

      return Object.assign({}, comparison, {
        rows: comparison.rows.filter(rowMatchesFilters),
      });
    }

    function rowMatchesFilters(row) {
      const filters = state.globalFilters;
      const objectQuery = normalizeSearch(filters.objectQuery);
      const departmentQuery = normalizeSearch(filters.departmentQuery);

      if (objectQuery && !matchesText([row.label, row.key], objectQuery)) {
        return false;
      }

      if (departmentQuery && !matchesText(getRowSearchValues(row), departmentQuery)) {
        return false;
      }

      if (!hasMetricFilters()) {
        return true;
      }

      const comparisons = getRowFilterComparisons(row);

      if (!comparisons.length) {
        return false;
      }

      return comparisons.some(matchMetricFilters);
    }

    function matchMetricFilters(item) {
      const filters = state.globalFilters;
      const min = parseFilterNumber(filters.deltaMin);
      const max = parseFilterNumber(filters.deltaMax);

      if (filters.impact === "good" && item.impact !== "good") {
        return false;
      }

      if (filters.impact === "bad" && item.impact !== "bad") {
        return false;
      }

      if (Number.isFinite(min) && item.delta < min) {
        return false;
      }

      if (Number.isFinite(max) && item.delta > max) {
        return false;
      }

      return true;
    }

    function getRowFilterComparisons(row) {
      const metric = getSelectedMetric();
      const metricResults = metric
        ? row.metrics.filter(function (item) {
            return item.metricId === metric.id;
          })
        : row.metrics;

      return metricResults
        .flatMap(function (result) {
          return result && Array.isArray(result.comparisons) ? result.comparisons : [];
        })
        .filter(function (item) {
          return Number.isFinite(item.delta);
        });
    }

    function hasMetricFilters() {
      const filters = state.globalFilters;

      return (
        filters.impact !== "all" ||
        Number.isFinite(parseFilterNumber(filters.deltaMin)) ||
        Number.isFinite(parseFilterNumber(filters.deltaMax))
      );
    }

    function hasActive() {
      const filters = state.globalFilters;

      return Boolean(
        filters &&
          (filters.impact !== "all" ||
            String(filters.deltaMin || "").trim() ||
            String(filters.deltaMax || "").trim() ||
            String(filters.objectQuery || "").trim() ||
            String(filters.departmentQuery || "").trim())
      );
    }

    function render(view) {
      ensureState();

      const hasComparison = Boolean(state.comparison);
      const hasActiveFilters = hasActive();
      const controls = [
        dom.globalImpactFilter,
        dom.globalDeltaMinFilter,
        dom.globalDeltaMaxFilter,
        dom.globalObjectSearch,
        dom.globalDepartmentSearch,
      ];

      dom.globalImpactFilter.value = state.globalFilters.impact;
      dom.globalDeltaMinFilter.value = state.globalFilters.deltaMin;
      dom.globalDeltaMaxFilter.value = state.globalFilters.deltaMax;
      dom.globalObjectSearch.value = state.globalFilters.objectQuery;
      dom.globalDepartmentSearch.value = state.globalFilters.departmentQuery;

      controls.forEach(function (control) {
        control.disabled = !hasComparison;
      });

      dom.applyGlobalFiltersButton.disabled = !hasComparison;
      dom.resetGlobalFiltersButton.disabled = !hasComparison || !hasActiveFilters;

      if (!hasComparison) {
        closeSuggestions();
        dom.globalFilterStatus.textContent = "Фильтры появятся после расчета";
        return;
      }

      if (!hasActiveFilters) {
        dom.globalFilterStatus.textContent = "Показаны все строки: " + view.totalRows;
        return;
      }

      dom.globalFilterStatus.textContent = view.visibleRows
        ? "Показано " + view.visibleRows + " из " + view.totalRows
        : "Ничего не найдено";
    }

    function renderObjectSuggestions() {
      if (!dom.globalObjectSuggestions || !state.comparison || dom.globalObjectSearch.disabled) {
        closeSuggestions();
        return;
      }

      const query = dom.globalObjectSearch.value.trim();

      if (!query) {
        closeSuggestions();
        return;
      }

      const suggestions = getObjectSuggestions(query, 8);

      if (!suggestions.length) {
        objectSuggestionIndex = -1;
        dom.globalObjectSuggestions.hidden = false;
        dom.globalObjectSuggestions.innerHTML = '<div class="autocomplete-empty">Ничего не найдено</div>';
        dom.globalObjectSearch.setAttribute("aria-expanded", "true");
        return;
      }

      objectSuggestionIndex = -1;
      dom.globalObjectSuggestions.hidden = false;
      dom.globalObjectSuggestions.innerHTML = suggestions.map(renderObjectSuggestion).join("");
      dom.globalObjectSearch.setAttribute("aria-expanded", "true");
    }

    function getObjectSuggestions(query, limit) {
      const normalizedQuery = normalizeSearch(query);
      const rows = state.comparison && Array.isArray(state.comparison.rows) ? state.comparison.rows : [];
      const seen = new Set();
      const suggestions = [];

      rows.some(function (row) {
        const label = String(row.label || "").trim();
        const key = String(row.key || "").trim();
        const values = Array.from(new Set(getRowSearchValues(row).map(function (value) {
          return String(value || "").trim();
        }).filter(Boolean)));
        const matchedValue = values.find(function (value) {
          return normalizeSearch(value).includes(normalizedQuery);
        });

        if (!matchedValue) {
          return false;
        }

        const uniqueKey = normalizeSearch(label + "|" + key);

        if (!seen.has(uniqueKey)) {
          seen.add(uniqueKey);
          suggestions.push({
            label: label || key,
            key: key,
            hint:
              matchedValue &&
              normalizeSearch(matchedValue) !== normalizeSearch(label) &&
              normalizeSearch(matchedValue) !== normalizeSearch(key)
                ? matchedValue
                : key,
          });
        }

        return suggestions.length >= limit;
      });

      return suggestions;
    }

    function renderObjectSuggestion(suggestion, index) {
      const keyHint =
        suggestion.hint && normalizeSearch(suggestion.hint) !== normalizeSearch(suggestion.label)
          ? '<span class="autocomplete-item__hint">' + escapeHtml(suggestion.hint) + "</span>"
          : "";

      return (
        '<button class="autocomplete-item" type="button" role="option" data-object-suggestion="' +
        escapeHtml(suggestion.label) +
        '" data-suggestion-index="' +
        index +
        '">' +
        '<span class="autocomplete-item__label">' +
        escapeHtml(suggestion.label) +
        "</span>" +
        keyHint +
        "</button>"
      );
    }

    function handleObjectSuggestionSelect(event) {
      const option = event.target.closest("[data-object-suggestion]");

      if (!option) {
        return;
      }

      event.preventDefault();
      selectObjectSuggestion(option.dataset.objectSuggestion);
    }

    function handleObjectSuggestionKeydown(event) {
      if (dom.globalObjectSuggestions.hidden) {
        return;
      }

      const options = Array.from(dom.globalObjectSuggestions.querySelectorAll("[data-object-suggestion]"));

      if (!options.length) {
        if (event.key === "Escape") {
          closeSuggestions();
        }
        return;
      }

      if (event.key === "ArrowDown") {
        event.preventDefault();
        setActiveObjectSuggestion(Math.min(objectSuggestionIndex + 1, options.length - 1), options);
        return;
      }

      if (event.key === "ArrowUp") {
        event.preventDefault();
        setActiveObjectSuggestion(Math.max(objectSuggestionIndex - 1, 0), options);
        return;
      }

      if (event.key === "Enter" && objectSuggestionIndex >= 0) {
        event.preventDefault();
        selectObjectSuggestion(options[objectSuggestionIndex].dataset.objectSuggestion);
        return;
      }

      if (event.key === "Escape") {
        closeSuggestions();
      }
    }

    function setActiveObjectSuggestion(index, options) {
      objectSuggestionIndex = index;

      options.forEach(function (option, optionIndex) {
        option.classList.toggle("is-active", optionIndex === objectSuggestionIndex);
      });
    }

    function selectObjectSuggestion(value) {
      dom.globalObjectSearch.value = value || "";
      closeSuggestions();
      dom.globalObjectSearch.focus();
    }

    function closeSuggestions() {
      objectSuggestionIndex = -1;

      if (!dom.globalObjectSuggestions) {
        return;
      }

      dom.globalObjectSuggestions.hidden = true;
      dom.globalObjectSuggestions.innerHTML = "";

      if (dom.globalObjectSearch) {
        dom.globalObjectSearch.setAttribute("aria-expanded", "false");
      }
    }

    function getRowSearchValues(row) {
      const values = [row.label, row.key];

      if (!Array.isArray(row.records)) {
        return values;
      }

      row.records.forEach(function (record) {
        if (!record || !record.row || !record.row.values) {
          return;
        }

        Object.keys(record.row.values).forEach(function (key) {
          values.push(record.row.values[key]);
        });
      });

      return values;
    }

    function normalizeSearch(value) {
      return String(value || "").trim().toLowerCase();
    }

    function matchesText(values, query) {
      return values.some(function (value) {
        return String(value || "").toLowerCase().includes(query);
      });
    }

    function parseFilterNumber(value) {
      const text = String(value || "").replace(",", ".").trim();

      if (!text) {
        return NaN;
      }

      return Number(text);
    }

    return {
      init: init,
      resetState: resetState,
      buildView: buildView,
      render: render,
      closeSuggestions: closeSuggestions,
      hasActive: hasActive,
    };
  }

  App.FiltersController = {
    create: create,
  };
})(window);
