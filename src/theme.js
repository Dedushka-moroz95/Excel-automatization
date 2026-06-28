(function (global) {
  const STORAGE_KEY = "metricum.theme";
  const ZOOM_STORAGE_KEY = "metricum.zoom";
  const ZOOM_CONFIRMED_KEY = "metricum.zoomConfirmed";
  const PREFERENCES_POSITION_KEY = "metricum.preferencesPosition";
  const LEGACY_STORAGE_KEY = "operational" + "Analytics.theme";
  const DARK_CLASS = "theme-dark";
  const DEFAULT_ZOOM = 100;
  const MIN_ZOOM = 25;
  const MAX_ZOOM = 500;
  const DRAG_MARGIN = 8;
  const DRAG_THRESHOLD = 4;

  let dragState = null;

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    initTheme();
    initZoom();
    initPreferencesPanel();
  }

  function initTheme() {
    const button = document.getElementById("themeToggle");

    if (!button) {
      return;
    }

    applyTheme(readStoredTheme());
    syncThemeButton(button);

    button.addEventListener("click", function () {
      const nextTheme = isDarkTheme() ? "light" : "dark";
      applyTheme(nextTheme);
      storeTheme(nextTheme);
      syncThemeButton(button);
      notifyThemeChange(nextTheme);
    });
  }

  function initZoom() {
    const button = document.getElementById("zoomToggle");
    const menu = document.getElementById("zoomMenu");

    if (!button || !menu) {
      return;
    }

    const options = Array.from(menu.querySelectorAll("[data-zoom-option]"));

    applyZoom(readStoredZoom());
    syncZoomControls(button, options);

    button.addEventListener("click", function (event) {
      event.stopPropagation();
      toggleZoomMenu(button, menu);
    });

    options.forEach(function (option) {
      option.addEventListener("click", function () {
        const zoom = normalizeZoom(option.dataset.zoomOption);

        applyZoom(zoom);
        storeZoom(zoom);
        syncZoomControls(button, options);
        closeZoomMenu(button, menu);
        notifyZoomChange(zoom);
      });
    });

    document.addEventListener("click", function (event) {
      if (!menu.hidden && !menu.contains(event.target) && event.target !== button) {
        closeZoomMenu(button, menu);
      }
    });

    document.addEventListener("keydown", function (event) {
      if (event.key === "Escape" && !menu.hidden) {
        closeZoomMenu(button, menu);
        button.focus();
      }
    });
  }

  function initPreferencesPanel() {
    const panel = document.querySelector(".app-preferences");

    if (!panel) {
      return;
    }

    applyStoredPreferencesPosition(panel);

    panel.addEventListener("pointerdown", function (event) {
      startPreferencesDrag(event, panel);
    });

    global.addEventListener("resize", function () {
      if (panel.dataset.positioned === "true") {
        clampPreferencesPosition(panel);
      }
    });
  }

  function applyTheme(theme) {
    document.documentElement.classList.toggle(DARK_CLASS, theme === "dark");
  }

  function syncThemeButton(button) {
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

  function applyZoom(zoom) {
    const normalizedZoom = normalizeZoom(zoom);
    const zoomValue = trimNumber(normalizedZoom / 100);

    document.documentElement.style.setProperty("--ui-zoom", zoomValue);
  }

  function syncZoomControls(button, options) {
    const zoom = getCurrentZoom();
    const label = "Масштаб интерфейса: " + zoom + "%";

    button.dataset.zoom = String(zoom);
    button.setAttribute("aria-label", label);
    button.setAttribute("title", label);

    options.forEach(function (option) {
      option.setAttribute("aria-checked", String(normalizeZoom(option.dataset.zoomOption) === zoom));
    });
  }

  function toggleZoomMenu(button, menu) {
    if (menu.hidden) {
      openZoomMenu(button, menu);
      return;
    }

    closeZoomMenu(button, menu);
  }

  function openZoomMenu(button, menu) {
    menu.hidden = false;
    button.setAttribute("aria-expanded", "true");
  }

  function closeZoomMenu(button, menu) {
    menu.hidden = true;
    button.setAttribute("aria-expanded", "false");
  }

  function getCurrentZoom() {
    const value = Number.parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--ui-zoom"));

    if (!Number.isFinite(value)) {
      return DEFAULT_ZOOM;
    }

    return normalizeZoom(Math.round(value * 100));
  }

  function readStoredZoom() {
    try {
      if (global.localStorage.getItem(ZOOM_CONFIRMED_KEY) !== "true") {
        return DEFAULT_ZOOM;
      }

      return normalizeZoom(global.localStorage.getItem(ZOOM_STORAGE_KEY));
    } catch (_error) {
      return DEFAULT_ZOOM;
    }
  }

  function storeZoom(zoom) {
    try {
      global.localStorage.setItem(ZOOM_STORAGE_KEY, String(normalizeZoom(zoom)));
      global.localStorage.setItem(ZOOM_CONFIRMED_KEY, "true");
      global.localStorage.removeItem("metricum.density");
    } catch (_error) {
      // Zoom persistence is optional; the controls should keep working without storage.
    }
  }

  function normalizeZoom(zoom) {
    const value = Number(zoom);

    if (!Number.isFinite(value)) {
      return DEFAULT_ZOOM;
    }

    return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, Math.round(value)));
  }

  function notifyZoomChange(zoom) {
    global.dispatchEvent(new CustomEvent("metricum:zoomchange", {
      detail: {
        zoom: zoom,
      },
    }));

    global.requestAnimationFrame(function () {
      global.dispatchEvent(new Event("resize"));
    });
  }

  function trimNumber(value) {
    return String(Number(value.toFixed(3)));
  }

  function startPreferencesDrag(event, panel) {
    if (
      event.button !== 0 ||
      event.target.closest(".preference-button") ||
      event.target.closest(".zoom-menu")
    ) {
      return;
    }

    const rect = panel.getBoundingClientRect();

    dragState = {
      panel: panel,
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      startLeft: rect.left,
      startTop: rect.top,
      moved: false,
    };

    panel.setPointerCapture(event.pointerId);
    panel.addEventListener("pointermove", movePreferencesPanel);
    panel.addEventListener("pointerup", stopPreferencesDrag, { once: true });
    panel.addEventListener("pointercancel", stopPreferencesDrag, { once: true });
  }

  function movePreferencesPanel(event) {
    if (!dragState || event.pointerId !== dragState.pointerId) {
      return;
    }

    const deltaX = event.clientX - dragState.startX;
    const deltaY = event.clientY - dragState.startY;

    if (!dragState.moved && Math.hypot(deltaX, deltaY) < DRAG_THRESHOLD) {
      return;
    }

    dragState.moved = true;
    dragState.panel.classList.add("is-dragging");
    setPreferencesPosition(dragState.panel, dragState.startLeft + deltaX, dragState.startTop + deltaY);
    event.preventDefault();
  }

  function stopPreferencesDrag(event) {
    if (!dragState || event.pointerId !== dragState.pointerId) {
      dragState = null;
      return;
    }

    const panel = dragState.panel;
    const moved = dragState.moved;

    panel.classList.remove("is-dragging");
    panel.removeEventListener("pointermove", movePreferencesPanel);
    panel.removeEventListener("pointerup", stopPreferencesDrag);
    panel.removeEventListener("pointercancel", stopPreferencesDrag);

    if (panel.hasPointerCapture(event.pointerId)) {
      panel.releasePointerCapture(event.pointerId);
    }

    if (moved) {
      storePreferencesPosition(panel);
      suppressNextClick(panel);
    }

    dragState = null;
  }

  function setPreferencesPosition(panel, left, top) {
    const rect = panel.getBoundingClientRect();
    const maxLeft = Math.max(DRAG_MARGIN, global.innerWidth - rect.width - DRAG_MARGIN);
    const maxTop = Math.max(DRAG_MARGIN, global.innerHeight - rect.height - DRAG_MARGIN);
    const nextLeft = Math.min(maxLeft, Math.max(DRAG_MARGIN, left));
    const nextTop = Math.min(maxTop, Math.max(DRAG_MARGIN, top));

    panel.style.left = nextLeft + "px";
    panel.style.top = nextTop + "px";
    panel.style.right = "auto";
    panel.style.bottom = "auto";
    panel.dataset.positioned = "true";
  }

  function clampPreferencesPosition(panel) {
    const rect = panel.getBoundingClientRect();

    setPreferencesPosition(panel, rect.left, rect.top);
    storePreferencesPosition(panel);
  }

  function applyStoredPreferencesPosition(panel) {
    const position = readPreferencesPosition();

    if (!position) {
      return;
    }

    setPreferencesPosition(panel, position.left, position.top);
  }

  function readPreferencesPosition() {
    try {
      const value = JSON.parse(global.localStorage.getItem(PREFERENCES_POSITION_KEY) || "null");

      if (!value || !Number.isFinite(value.left) || !Number.isFinite(value.top)) {
        return null;
      }

      return value;
    } catch (_error) {
      return null;
    }
  }

  function storePreferencesPosition(panel) {
    try {
      const rect = panel.getBoundingClientRect();
      global.localStorage.setItem(PREFERENCES_POSITION_KEY, JSON.stringify({
        left: Math.round(rect.left),
        top: Math.round(rect.top),
      }));
    } catch (_error) {
      // Floating panel position is optional; dragging should keep working without storage.
    }
  }

  function suppressNextClick(panel) {
    panel.addEventListener("click", function preventDragClick(event) {
      event.preventDefault();
      event.stopPropagation();
      panel.removeEventListener("click", preventDragClick, true);
    }, true);
  }
})(window);
