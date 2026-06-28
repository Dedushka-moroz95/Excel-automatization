(function (global) {
  const STORAGE_KEY = "metricum.theme";
  const DENSITY_STORAGE_KEY = "metricum.density";
  const LEGACY_STORAGE_KEY = "operational" + "Analytics.theme";
  const DARK_CLASS = "theme-dark";
  const DENSITY_CLASSES = ["density-compact", "density-dense"];
  const DENSITY_LABELS = {
    comfortable: "Комфортно",
    compact: "Компактно",
    dense: "Плотно",
  };

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    initTheme();
    initDensity();
  }

  function initTheme() {
    const button = document.getElementById("themeToggle");

    if (!button) {
      return;
    }

    applyTheme(readStoredTheme());
    syncButton(button);

    button.addEventListener("click", function () {
      const nextTheme = isDarkTheme() ? "light" : "dark";
      applyTheme(nextTheme);
      storeTheme(nextTheme);
      syncButton(button);
      notifyThemeChange(nextTheme);
    });
  }

  function initDensity() {
    const button = document.getElementById("densityToggle");
    const menu = document.getElementById("densityMenu");

    if (!button || !menu) {
      return;
    }

    const options = Array.from(menu.querySelectorAll("[data-density-option]"));

    applyDensity(readStoredDensity());
    syncDensityControls(button, options);

    button.addEventListener("click", function (event) {
      event.stopPropagation();
      toggleDensityMenu(button, menu);
    });

    options.forEach(function (option) {
      option.addEventListener("click", function () {
        const density = option.dataset.densityOption || "comfortable";

        applyDensity(density);
        storeDensity(density);
        syncDensityControls(button, options);
        closeDensityMenu(button, menu);
        notifyDensityChange(density);
      });
    });

    document.addEventListener("click", function (event) {
      if (!menu.hidden && !menu.contains(event.target) && event.target !== button) {
        closeDensityMenu(button, menu);
      }
    });

    document.addEventListener("keydown", function (event) {
      if (event.key === "Escape" && !menu.hidden) {
        closeDensityMenu(button, menu);
        button.focus();
      }
    });
  }

  function applyTheme(theme) {
    document.documentElement.classList.toggle(DARK_CLASS, theme === "dark");
  }

  function syncButton(button) {
    const darkTheme = isDarkTheme();
    const actionLabel = darkTheme ? "Включить светлую тему" : "Включить темную тему";

    button.setAttribute("aria-pressed", String(darkTheme));
    button.setAttribute("aria-label", actionLabel);
    button.setAttribute("title", actionLabel);
    button.dataset.theme = darkTheme ? "dark" : "light";
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

  function applyDensity(density) {
    const normalizedDensity = normalizeDensity(density);

    DENSITY_CLASSES.forEach(function (className) {
      document.documentElement.classList.remove(className);
    });

    if (normalizedDensity !== "comfortable") {
      document.documentElement.classList.add("density-" + normalizedDensity);
    }
  }

  function syncDensityControls(button, options) {
    const density = getCurrentDensity();
    const label = DENSITY_LABELS[density] || DENSITY_LABELS.comfortable;

    button.dataset.density = density;
    button.setAttribute("aria-label", "Плотность интерфейса: " + label);
    button.setAttribute("title", "Плотность интерфейса: " + label);

    options.forEach(function (option) {
      option.setAttribute("aria-checked", String(option.dataset.densityOption === density));
    });
  }

  function toggleDensityMenu(button, menu) {
    if (menu.hidden) {
      openDensityMenu(button, menu);
      return;
    }

    closeDensityMenu(button, menu);
  }

  function openDensityMenu(button, menu) {
    menu.hidden = false;
    button.setAttribute("aria-expanded", "true");
  }

  function closeDensityMenu(button, menu) {
    menu.hidden = true;
    button.setAttribute("aria-expanded", "false");
  }

  function getCurrentDensity() {
    if (document.documentElement.classList.contains("density-dense")) {
      return "dense";
    }

    if (document.documentElement.classList.contains("density-compact")) {
      return "compact";
    }

    return "comfortable";
  }

  function readStoredDensity() {
    try {
      return normalizeDensity(global.localStorage.getItem(DENSITY_STORAGE_KEY));
    } catch (_error) {
      return "comfortable";
    }
  }

  function storeDensity(density) {
    try {
      global.localStorage.setItem(DENSITY_STORAGE_KEY, normalizeDensity(density));
    } catch (_error) {
      // Density persistence is optional; the controls should keep working without storage.
    }
  }

  function normalizeDensity(density) {
    return density === "compact" || density === "dense" ? density : "comfortable";
  }

  function notifyDensityChange(density) {
    global.dispatchEvent(new CustomEvent("metricum:densitychange", {
      detail: {
        density: density,
      },
    }));

    global.requestAnimationFrame(function () {
      global.dispatchEvent(new Event("resize"));
    });
  }
})(window);
