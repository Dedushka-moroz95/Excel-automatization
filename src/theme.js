(function (global) {
  const STORAGE_KEY = "operationalAnalytics.theme";
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
      return global.localStorage.getItem(STORAGE_KEY) === "dark" ? "dark" : "light";
    } catch (_error) {
      return "light";
    }
  }

  function storeTheme(theme) {
    try {
      global.localStorage.setItem(STORAGE_KEY, theme);
    } catch (_error) {
      // Theme persistence is optional; the toggle should keep working without storage.
    }
  }

  function notifyThemeChange(theme) {
    global.dispatchEvent(new CustomEvent("operationalAnalytics:themechange", {
      detail: {
        theme: theme,
      },
    }));
  }
})(window);
