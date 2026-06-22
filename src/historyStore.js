(function (global) {
  const App = (global.Metricum = global.Metricum || {});
  const STORAGE_KEY = "metricum.history.v1";
  const LEGACY_STORAGE_KEY = "operational" + "Analytics.history.v1";
  const MAX_RECORDS = 50;

  function load() {
    try {
      const raw = readStoredHistory();
      const parsed = raw ? JSON.parse(raw) : [];

      return normalizeRecords(Array.isArray(parsed) ? parsed : []);
    } catch (_error) {
      return [];
    }
  }

  function save(record) {
    const records = load();
    const storedRecord = normalizeRecord(record);
    const withoutCurrent = records.filter(function (item) {
      return item.id !== storedRecord.id;
    });
    const nextRecords = trim(sortRecords([storedRecord].concat(withoutCurrent)));

    return persist(nextRecords);
  }

  function remove(id) {
    const records = load().filter(function (item) {
      return item.id !== id;
    });

    return persist(records);
  }

  function clear() {
    try {
      global.localStorage.removeItem(STORAGE_KEY);
      global.localStorage.removeItem(LEGACY_STORAGE_KEY);
      return { ok: true, records: [] };
    } catch (error) {
      return { ok: false, records: load(), error: error };
    }
  }

  function togglePinned(id) {
    const records = load().map(function (item) {
      if (item.id !== id) {
        return item;
      }

      return Object.assign({}, item, {
        pinned: !item.pinned,
      });
    });

    return persist(sortRecords(records));
  }

  function persist(records) {
    const normalized = trim(sortRecords(records));

    try {
      global.localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized));
      global.localStorage.removeItem(LEGACY_STORAGE_KEY);
      return { ok: true, records: normalized };
    } catch (error) {
      return { ok: false, records: load(), error: error };
    }
  }

  function readStoredHistory() {
    const current = global.localStorage.getItem(STORAGE_KEY);

    if (current) {
      return current;
    }

    const legacy = global.localStorage.getItem(LEGACY_STORAGE_KEY);

    if (legacy) {
      global.localStorage.setItem(STORAGE_KEY, legacy);
      global.localStorage.removeItem(LEGACY_STORAGE_KEY);
    }

    return legacy;
  }

  function normalizeRecords(records) {
    return trim(sortRecords(records.map(normalizeRecord)));
  }

  function normalizeRecord(record) {
    const createdAt = record && record.createdAt ? record.createdAt : new Date().toISOString();

    return {
      id: record && record.id ? String(record.id) : createId(),
      createdAt: createdAt,
      title: record && record.title ? String(record.title) : defaultTitle(createdAt),
      pinned: Boolean(record && record.pinned),
      meta: record && record.meta ? record.meta : {},
      metrics: Array.isArray(record && record.metrics) ? record.metrics : [],
      settings: record && record.settings ? record.settings : {},
      comparison: record ? record.comparison : null,
      analytics: record ? record.analytics : null,
    };
  }

  function sortRecords(records) {
    return records.slice().sort(function (left, right) {
      if (left.pinned !== right.pinned) {
        return left.pinned ? -1 : 1;
      }

      return new Date(right.createdAt).getTime() - new Date(left.createdAt).getTime();
    });
  }

  function trim(records) {
    return records.slice(0, MAX_RECORDS);
  }

  function createId() {
    if (global.crypto && typeof global.crypto.randomUUID === "function") {
      return global.crypto.randomUUID();
    }

    return "analysis_" + Date.now() + "_" + Math.random().toString(16).slice(2);
  }

  function defaultTitle(createdAt) {
    return "Анализ от " + formatDateTime(createdAt);
  }

  function formatDateTime(value) {
    const date = value ? new Date(value) : new Date();

    if (Number.isNaN(date.getTime())) {
      return "";
    }

    return new Intl.DateTimeFormat("ru-RU", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    }).format(date);
  }

  App.HistoryStore = {
    load: load,
    save: save,
    remove: remove,
    clear: clear,
    togglePinned: togglePinned,
    defaultTitle: defaultTitle,
    formatDateTime: formatDateTime,
  };
})(window);
