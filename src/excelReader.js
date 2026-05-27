(function (global) {
  const App = (global.OperationalAnalytics = global.OperationalAnalytics || {});
  const Normalizers = App.Normalizers;

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

  async function readExcelFile(file) {
    if (!global.XLSX) {
      throw new Error("Библиотека SheetJS не загружена");
    }

    const buffer = await readFileAsArrayBuffer(file);
    const workbook = global.XLSX.read(buffer, {
      type: "array",
      cellDates: true,
    });

    if (!workbook.SheetNames.length) {
      throw new Error("В файле нет листов: " + file.name);
    }

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
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
    const warnings = analyzeStructure(headers, rows, file.name, headerRowIndex);

    return {
      fileName: file.name,
      sheetName: sheetName,
      headerRowIndex: headerRowIndex,
      headers: headers,
      rows: rows,
      previewRows: rows.slice(0, 8),
      warnings: warnings,
    };
  }

  function detectHeaderRow(matrix) {
    let bestIndex = 0;
    let bestScore = -1;
    const limit = Math.min(matrix.length, 20);

    for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
      const row = matrix[rowIndex] || [];
      const filledCells = row.filter(function (cell) {
        return !Normalizers.isEmptyValue(cell);
      }).length;

      const textCells = row.filter(function (cell) {
        return typeof cell === "string" && !Normalizers.isEmptyValue(cell);
      }).length;

      const score = filledCells + textCells * 0.5;

      if (score > bestScore) {
        bestScore = score;
        bestIndex = rowIndex;
      }
    }

    return bestIndex;
  }

  function getMaxColumns(matrix, headerRowIndex) {
    const rowsForScan = matrix.slice(headerRowIndex, headerRowIndex + 50);
    return Math.max.apply(
      null,
      rowsForScan.map(function (row) {
        return row.length;
      })
    );
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
    return matrix
      .slice(headerRowIndex + 1)
      .filter(function (row) {
        return row.some(function (cell) {
          return !Normalizers.isEmptyValue(cell);
        });
      })
      .map(function (row, rowIndex) {
        const values = {};

        headers.forEach(function (header) {
          values[header.id] = row[header.index] === undefined ? "" : row[header.index];
        });

        return {
          rowNumber: headerRowIndex + 2 + rowIndex,
          values: values,
        };
      });
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

  App.ExcelReader = {
    readExcelFile,
  };
})(window);
