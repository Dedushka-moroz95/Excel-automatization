(function (global) {
  const App = (global.Metricum = global.Metricum || {});
  const Normalizers = App.Normalizers;

  function buildVirtualPeriods(options) {
    const table = options.table;
    const periodColumn = options.periodColumn;
    const warnings = [];

    if (!table || !periodColumn) {
      return {
        periods: [],
        warnings: warnings,
      };
    }

    const groups = groupRowsByPeriod(table.rows, periodColumn);

    if (groups.emptyRows.length) {
      warnings.push({
        type: "warn",
        message: "В колонке периода есть пустые значения: строк " + groups.emptyRows.length + ".",
      });
    }

    if (groups.items.length < 2) {
      warnings.push({
        type: "warn",
        message: "Для сравнения нужны минимум два периода внутри файла.",
      });
    }

    return {
      periods: groups.items.map(function (group, index) {
        return buildVirtualPeriod(table, periodColumn, group, index);
      }),
      warnings: warnings,
    };
  }

  function groupRowsByPeriod(rows, periodColumn) {
    const groupsByKey = new Map();
    const items = [];
    const emptyRows = [];

    rows.forEach(function (row) {
      const rawValue = row.values[periodColumn];
      const label = Normalizers.toText(rawValue).trim().replace(/\s+/g, " ");
      const key = Normalizers.normalizeKey(label);

      if (isEmptyPeriodKey(key)) {
        emptyRows.push(row.rowNumber);
        return;
      }

      if (!groupsByKey.has(key)) {
        const group = {
          key: key,
          label: label,
          rows: [],
        };
        groupsByKey.set(key, group);
        items.push(group);
      }

      groupsByKey.get(key).rows.push(row);
    });

    return {
      items: items,
      emptyRows: emptyRows,
    };
  }

  function isEmptyPeriodKey(key) {
    return key === "" || key === "-" || key === "—" || key === "n/a" || key === "нет";
  }

  function buildVirtualPeriod(table, periodColumn, group, index) {
    const id = "virtual_period_" + index + "_" + safeId(group.key);

    return {
      id: id,
      label: group.label || "Период " + (index + 1),
      file: null,
      table: {
        fileName: "Период: " + (group.label || index + 1),
        sheetName: table.sheetName,
        headerRowIndex: table.headerRowIndex,
        headers: table.headers,
        rows: group.rows,
        previewRows: group.rows.slice(0, 8),
        warnings: [],
      },
      idColumn: "",
      source: {
        type: "singleFile",
        fileName: table.fileName,
        periodColumn: periodColumn,
        periodValue: group.label,
        rowCount: group.rows.length,
      },
    };
  }

  function safeId(value) {
    return Normalizers.normalizeKey(value)
      .replace(/[^a-zа-яё0-9]+/gi, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, 40) || "period";
  }

  App.PeriodBuilder = {
    buildVirtualPeriods: buildVirtualPeriods,
  };
})(window);
