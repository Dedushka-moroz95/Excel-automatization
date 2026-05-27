(function (global) {
  const App = (global.OperationalAnalytics = global.OperationalAnalytics || {});

  function toText(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value).replace(/\u00a0/g, " ");
  }

  function normalizeKey(value) {
    return toText(value).trim().replace(/\s+/g, " ").toLowerCase();
  }

  function isEmptyValue(value) {
    const text = normalizeKey(value);
    return text === "" || text === "-" || text === "—" || text === "n/a" || text === "нет";
  }

  function normalizeHeader(value, fallback) {
    const text = toText(value).trim().replace(/\s+/g, " ");
    return text || fallback;
  }

  function normalizeNumber(value) {
    if (value === null || value === undefined || value === "") {
      return { value: null, isEmpty: true, isNumeric: false, raw: value };
    }

    if (typeof value === "number") {
      if (Number.isFinite(value)) {
        return { value: value, isEmpty: false, isNumeric: true, raw: value };
      }

      return { value: null, isEmpty: false, isNumeric: false, raw: value };
    }

    const originalText = toText(value).trim();

    if (isEmptyValue(originalText)) {
      return { value: null, isEmpty: true, isNumeric: false, raw: value };
    }

    let text = originalText
      .replace(/\s/g, "")
      .replace(/%/g, "")
      .replace(/[₽$€£]/g, "")
      .replace(/[^\d.,+\-()]/g, "");

    if (!text) {
      return { value: null, isEmpty: false, isNumeric: false, raw: value };
    }

    let sign = 1;
    if (text.startsWith("(") && text.endsWith(")")) {
      sign = -1;
      text = text.slice(1, -1);
    }

    text = normalizeDecimalSeparators(text);
    const number = Number(text);

    if (!Number.isFinite(number)) {
      return { value: null, isEmpty: false, isNumeric: false, raw: value };
    }

    return { value: number * sign, isEmpty: false, isNumeric: true, raw: value };
  }

  function normalizeDecimalSeparators(text) {
    const lastComma = text.lastIndexOf(",");
    const lastDot = text.lastIndexOf(".");

    if (lastComma !== -1 && lastDot !== -1) {
      const decimalSeparator = lastComma > lastDot ? "," : ".";
      const thousandSeparator = decimalSeparator === "," ? "." : ",";
      return text.split(thousandSeparator).join("").replace(decimalSeparator, ".");
    }

    if (lastComma !== -1) {
      const parts = text.split(",");
      if (parts.length === 2 && parts[1].length <= 3) {
        return parts[0].replace(/\./g, "") + "." + parts[1];
      }
      return parts.join("");
    }

    if (lastDot !== -1) {
      const parts = text.split(".");
      if (parts.length === 2 && parts[1].length <= 3) {
        return text;
      }
      return parts.join("");
    }

    return text;
  }

  function calculatePercentChange(previous, current) {
    if (!Number.isFinite(previous) || !Number.isFinite(current) || previous === 0) {
      return null;
    }

    return ((current - previous) / Math.abs(previous)) * 100;
  }

  function formatNumber(value, digits) {
    if (!Number.isFinite(value)) {
      return "—";
    }

    const safeDigits = digits === undefined ? 2 : digits;
    return new Intl.NumberFormat("ru-RU", {
      maximumFractionDigits: safeDigits,
    }).format(value);
  }

  function formatPercent(value) {
    if (!Number.isFinite(value)) {
      return "—";
    }

    return formatNumber(value, 1) + "%";
  }

  App.Normalizers = {
    toText,
    normalizeKey,
    normalizeHeader,
    normalizeNumber,
    calculatePercentChange,
    formatNumber,
    formatPercent,
    isEmptyValue,
  };
})(window);
