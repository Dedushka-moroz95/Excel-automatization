(function (global) {
  const STORAGE_KEY = "metricum.theme";
  const LEGACY_STORAGE_KEY = "operational" + "Analytics.theme";
  const DARK_CLASS = "theme-dark";

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    const button = document.getElementById("themeToggle");
    const label = document.getElementById("themeToggleLabel");

    if (!button || !label) {
      return;
    }

    applyTheme(readStoredTheme());
    syncButton(button, label);

    button.addEventListener("click", function () {
      const nextTheme = isDarkTheme() ? "light" : "dark";
      applyTheme(nextTheme);
      storeTheme(nextTheme);
      syncButton(button, label);
      notifyThemeChange(nextTheme);
    });
  }

  function applyTheme(theme) {
    document.documentElement.classList.toggle(DARK_CLASS, theme === "dark");
  }

  function syncButton(button, label) {
    const darkTheme = isDarkTheme();

    button.setAttribute("aria-pressed", String(darkTheme));
    button.dataset.theme = darkTheme ? "dark" : "light";
    label.textContent = darkTheme ? "Светлая тема" : "Темная тема";
  }

  function isDarkTheme() {
    return document.documentElement.classList.contains(DARK_CLASS);
  }

  function readStoredTheme() {
    try {
      const theme = readStoredValue();
      return theme === "dark" ? "dark" : "light";
    } catch (_error) {
      return "light";
    }
  }

  function storeTheme(theme) {
    try {
      global.localStorage.setItem(STORAGE_KEY, theme);
      global.localStorage.removeItem(LEGACY_STORAGE_KEY);
    } catch (_error) {
      // Theme persistence is optional; the toggle should keep working without storage.
    }
  }

  function notifyThemeChange(theme) {
    global.dispatchEvent(new CustomEvent("metricum:themechange", {
      detail: {
        theme: theme,
      },
    }));
  }

  function readStoredValue() {
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
})(window);
