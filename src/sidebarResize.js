(function (global) {
  const STORAGE_KEY = "operationalAnalytics.sidebarWidth";
  const DESKTOP_MIN_WIDTH = 1100;
  const MIN_WIDTH = 250;
  const MAX_WIDTH = 420;
  const MAX_VIEWPORT_RATIO = 0.42;

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    const handle = document.querySelector(".sidebar-resize-handle");

    if (!handle) {
      return;
    }

    applyStoredWidth();

    if (global.PointerEvent) {
      handle.addEventListener("pointerdown", function (event) {
        if (!canStartResize(event)) {
          return;
        }

        event.preventDefault();
        document.body.classList.add("is-sidebar-resizing");
        let lastClientX = event.clientX;

        try {
          handle.setPointerCapture(event.pointerId);
        } catch (_error) {
          // Drag still works through document-level listeners.
        }

        function handleMove(moveEvent) {
          if (moveEvent.pointerId !== event.pointerId) {
            return;
          }

          moveEvent.preventDefault();
          lastClientX = moveEvent.clientX;
          setSidebarWidth(moveEvent.clientX, false);
        }

        function handleUp(upEvent) {
          if (upEvent.pointerId !== event.pointerId) {
            return;
          }

          try {
            handle.releasePointerCapture(upEvent.pointerId);
          } catch (_error) {
            // Pointer capture may already be released by the browser.
          }

          stopResize();
          const shouldUseLastClientX = upEvent.type === "pointercancel" || !Number.isFinite(upEvent.clientX);
          const finalClientX = shouldUseLastClientX ? lastClientX : upEvent.clientX;
          setSidebarWidth(finalClientX, true);
          document.removeEventListener("pointermove", handleMove);
          document.removeEventListener("pointerup", handleUp);
          document.removeEventListener("pointercancel", handleUp);
        }

        document.addEventListener("pointermove", handleMove, { passive: false });
        document.addEventListener("pointerup", handleUp);
        document.addEventListener("pointercancel", handleUp);
      });
    } else {
      handle.addEventListener("mousedown", function (event) {
        if (!canStartResize(event)) {
          return;
        }

        event.preventDefault();
        document.body.classList.add("is-sidebar-resizing");
        let lastClientX = event.clientX;

        function handleMove(moveEvent) {
          moveEvent.preventDefault();
          lastClientX = moveEvent.clientX;
          setSidebarWidth(moveEvent.clientX, false);
        }

        function handleUp(upEvent) {
          stopResize();
          setSidebarWidth(Number.isFinite(upEvent.clientX) ? upEvent.clientX : lastClientX, true);
          document.removeEventListener("mousemove", handleMove);
          document.removeEventListener("mouseup", handleUp);
        }

        document.addEventListener("mousemove", handleMove);
        document.addEventListener("mouseup", handleUp);
      });
    }

    global.addEventListener("resize", applyStoredWidth);
  }

  function canStartResize(event) {
    return isDesktop() && (event.button === undefined || event.button === 0);
  }

  function stopResize() {
    document.body.classList.remove("is-sidebar-resizing");
  }

  function applyStoredWidth() {
    if (!isDesktop()) {
      return;
    }

    const storedWidth = readStoredWidth();

    if (storedWidth) {
      setSidebarWidth(storedWidth, false);
    }
  }

  function setSidebarWidth(width, shouldPersist) {
    const numericWidth = Number(width);

    if (!Number.isFinite(numericWidth)) {
      return;
    }

    const nextWidth = clampWidth(numericWidth);

    document.documentElement.style.setProperty("--sidebar-width", nextWidth + "px");

    if (shouldPersist) {
      try {
        global.localStorage.setItem(STORAGE_KEY, String(nextWidth));
      } catch (_error) {
        // Width persistence is a convenience; resizing should keep working without storage.
      }
    }
  }

  function readStoredWidth() {
    try {
      const value = Number(global.localStorage.getItem(STORAGE_KEY));
      return Number.isFinite(value) ? value : 0;
    } catch (_error) {
      return 0;
    }
  }

  function clampWidth(width) {
    const viewportMax = Math.max(MIN_WIDTH, Math.min(MAX_WIDTH, Math.round(global.innerWidth * MAX_VIEWPORT_RATIO)));
    return Math.min(Math.max(Math.round(width), MIN_WIDTH), viewportMax);
  }

  function isDesktop() {
    return global.innerWidth >= DESKTOP_MIN_WIDTH;
  }
})(window);
