(function (global) {
  const App = (global.Metricum = global.Metricum || {});
  const Normalizers = App.Normalizers;
  const LARGE_ROW_THRESHOLD = 50000;
  const VERY_LARGE_ROW_THRESHOLD = 100000;
  const LARGE_CELL_THRESHOLD = 3000000;
  const VERY_LARGE_CELL_THRESHOLD = 5000000;

  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      const reader = new FileReader();

      reader.onload = function (event) {
        resolve(event.target.result);
      };

      reader.onerror = function () {
        reject(new Error("Не удалось прочитать файл: " + file.name));
      };

      reader.readAsArrayBuffer(file);
    });
  }

  async function readExcelFile(file, options) {
    if (!global.XLSX) {
      throw new Error("Библиотека SheetJS не загружена");
    }

    await notifyProgress(options, "Чтение файла", file.name);
    const buffer = await readFileAsArrayBuffer(file);
    await notifyProgress(options, "Анализ структуры", file.name);
    const workbook = global.XLSX.read(buffer, {
      type: "array",
      cellDates: true,
    });

    if (!workbook.SheetNames.length) {
      throw new Error("В файле нет листов: " + file.name);
    }

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    await notifyProgress(options, "Проверка таблицы", sheetName);
    const matrix = global.XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: "",
      blankrows: false,
    });

    if (!matrix.length) {
      throw new Error("Первый лист пустой: " + file.name);
    }

    const headerRowIndex = detectHeaderRow(matrix);
    const maxColumns = getMaxColumns(matrix, headerRowIndex);
    const headers = buildHeaders(matrix[headerRowIndex] || [], maxColumns);
    const rows = buildRows(matrix, headers, headerRowIndex);
    const workload = buildWorkloadInfo(headers, rows);
    const warnings = analyzeStructure(headers, rows, file.name, headerRowIndex).concat(
      analyzeWorkload(workload, file.name)
    );

    return {
      fileName: file.name,
      sheetName: sheetName,
      headerRowIndex: headerRowIndex,
      headers: headers,
      rows: rows,
      previewRows: rows.slice(0, 8),
      workload: workload,
      warnings: warnings,
    };
  }

  function detectHeaderRow(matrix) {
    let bestIndex = 0;
    let bestScore = -1;
    const limit = Math.min(matrix.length, 20);

    for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
      const row = matrix[rowIndex] || [];
      let filledCells = 0;
      let textCells = 0;

      for (let cellIndex = 0; cellIndex < row.length; cellIndex += 1) {
        const cell = row[cellIndex];

        if (Normalizers.isEmptyValue(cell)) {
          continue;
        }

        filledCells += 1;

        if (typeof cell === "string") {
          textCells += 1;
        }
      }

      const score = filledCells + textCells * 0.5;

      if (score > bestScore) {
        bestScore = score;
        bestIndex = rowIndex;
      }
    }

    return bestIndex;
  }

  function getMaxColumns(matrix, headerRowIndex) {
    const limit = Math.min(matrix.length, headerRowIndex + 50);
    let maxColumns = 0;

    for (let rowIndex = headerRowIndex; rowIndex < limit; rowIndex += 1) {
      const row = matrix[rowIndex] || [];

      if (row.length > maxColumns) {
        maxColumns = row.length;
      }
    }

    return maxColumns;
  }

  function buildHeaders(rawHeaders, maxColumns) {
    const headers = [];

    for (let index = 0; index < maxColumns; index += 1) {
      headers.push({
        id: "col_" + index,
        index: index,
        name: Normalizers.normalizeHeader(rawHeaders[index], "Колонка " + (index + 1)),
        originalName: rawHeaders[index] || "",
      });
    }

    return headers;
  }

  function buildRows(matrix, headers, headerRowIndex) {
    const rows = [];

    for (let matrixIndex = headerRowIndex + 1; matrixIndex < matrix.length; matrixIndex += 1) {
      const rawRow = matrix[matrixIndex] || [];

      if (!hasUsefulCells(rawRow)) {
        continue;
      }

      const values = {};

      for (let headerIndex = 0; headerIndex < headers.length; headerIndex += 1) {
        const header = headers[headerIndex];
        values[header.id] = rawRow[header.index] === undefined ? "" : rawRow[header.index];
      }

      rows.push({
        rowNumber: matrixIndex + 1,
        values: values,
      });
    }

    return rows;
  }

  function hasUsefulCells(row) {
    for (let index = 0; index < row.length; index += 1) {
      if (!Normalizers.isEmptyValue(row[index])) {
        return true;
      }
    }

    return false;
  }

  function analyzeStructure(headers, rows, fileName, headerRowIndex) {
    const warnings = [];
    const seen = new Map();
    const emptyHeaders = headers.filter(function (header) {
      return !header.originalName;
    });

    headers.forEach(function (header) {
      const key = Normalizers.normalizeKey(header.name);
      if (!seen.has(key)) {
        seen.set(key, []);
      }
      seen.get(key).push(header.name);
    });

    seen.forEach(function (items, key) {
      if (key && items.length > 1) {
        warnings.push({
          type: "warn",
          message: fileName + ": повторяется заголовок «" + items[0] + "»",
        });
      }
    });

    if (emptyHeaders.length > 0) {
      warnings.push({
        type: "warn",
        message: fileName + ": есть пустые заголовки, им назначены технические названия",
      });
    }

    if (headerRowIndex > 0) {
      warnings.push({
        type: "warn",
        message: fileName + ": заголовки найдены не в первой строке, а в строке " + (headerRowIndex + 1),
      });
    }

    if (!rows.length) {
      warnings.push({
        type: "error",
        message: fileName + ": после строки заголовков нет строк с данными",
      });
    }

    return warnings;
  }

  function buildWorkloadInfo(headers, rows) {
    const rowCount = rows.length;
    const columnCount = headers.length;
    const cellCount = rowCount * columnCount;
    const level =
      rowCount >= VERY_LARGE_ROW_THRESHOLD || cellCount >= VERY_LARGE_CELL_THRESHOLD
        ? "very-large"
        : rowCount >= LARGE_ROW_THRESHOLD || cellCount >= LARGE_CELL_THRESHOLD
          ? "large"
          : "normal";

    return {
      rowCount: rowCount,
      columnCount: columnCount,
      cellCount: cellCount,
      level: level,
    };
  }

  function analyzeWorkload(workload, fileName) {
    if (!workload || workload.level === "normal") {
      return [];
    }

    const prefix = workload.level === "very-large" ? "очень крупный файл" : "крупный файл";
    const message =
      fileName +
      ": " +
      prefix +
      " - " +
      formatCount(workload.rowCount) +
      " строк, " +
      formatCount(workload.columnCount) +
      " колонок, примерно " +
      formatCount(workload.cellCount) +
      " ячеек. Загрузка, расчет и экспорт могут занять больше времени.";

    return [{
      type: "warn",
      message: message,
    }];
  }

  function formatCount(value) {
    return Normalizers.formatNumber(Number(value) || 0, 0);
  }

  async function notifyProgress(options, title, detail) {
    if (options && typeof options.onProgress === "function") {
      await options.onProgress({
        title: title,
        detail: detail,
      });
      return;
    }

    await yieldToBrowser();
  }

  function yieldToBrowser() {
    return new Promise(function (resolve) {
      if (typeof global.requestAnimationFrame === "function") {
        global.requestAnimationFrame(function () {
          resolve();
        });
        return;
      }

      global.setTimeout(resolve, 0);
    });
  }

  App.ExcelReader = {
    readExcelFile,
  };
})(window);
