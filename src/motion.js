(function (global) {
  const App = (global.OperationalAnalytics = global.OperationalAnalytics || {});
  const REVEAL_SELECTOR = [
    ".app-header",
    ".workspace-section",
    ".section-title",
    ".period-toolbar",
    ".preview-panel",
    ".mapper-controls > .field",
    ".metric-toolbar",
    ".action-row",
    ".data-quality-inline",
    ".empty-state",
    ".warning-item",
    ".summary-card",
    ".chart-panel",
    ".mover-card",
    ".results-panel",
  ].join(",");

  const animatedCounters = new WeakSet();
  let revealObserver = null;
  let initialized = false;

  function init() {
    if (initialized) {
      return;
    }

    initialized = true;
    document.body.classList.add("motion-ready");

    if (!prefersReducedMotion() && "IntersectionObserver" in global) {
      revealObserver = new IntersectionObserver(handleReveal, {
        root: null,
        rootMargin: "0px 0px -8% 0px",
        threshold: 0.12,
      });
    }

    refresh();
  }

  function refresh(root) {
    const scope = root || document;
    const elements = Array.from(scope.querySelectorAll(REVEAL_SELECTOR));

    elements.forEach(function (element, index) {
      element.classList.add("reveal-item");
      element.style.setProperty("--motion-order", String(Math.min(index % 8, 7)));

      if (prefersReducedMotion() || !revealObserver) {
        revealNow(element);
        return;
      }

      if (element.classList.contains("is-visible")) {
        animateCounters(element);
        return;
      }

      if (isNearViewport(element)) {
        revealNow(element);
        return;
      }

      revealObserver.observe(element);
    });
  }

  function handleReveal(entries) {
    entries.forEach(function (entry) {
      if (!entry.isIntersecting) {
        return;
      }

      revealNow(entry.target);
      revealObserver.unobserve(entry.target);
    });
  }

  function revealNow(element) {
    element.classList.add("is-visible");
    animateCounters(element);
  }

  function animateCounters(root) {
    const counters = root.matches("[data-count-to]")
      ? [root]
      : Array.from(root.querySelectorAll("[data-count-to]"));

    counters.forEach(function (counter) {
      if (animatedCounters.has(counter)) {
        return;
      }

      animatedCounters.add(counter);
      animateCounter(counter);
    });
  }

  function animateCounter(counter) {
    const target = Number(counter.dataset.countTo);

    if (!Number.isFinite(target)) {
      return;
    }

    if (prefersReducedMotion()) {
      counter.textContent = formatNumber(target);
      return;
    }

    const duration = 460;
    const startedAt = performance.now();

    counter.textContent = "0";

    function tick(now) {
      const progress = Math.min((now - startedAt) / duration, 1);
      const eased = 1 - Math.pow(1 - progress, 3);
      counter.textContent = formatNumber(target * eased);

      if (progress < 1) {
        requestAnimationFrame(tick);
        return;
      }

      counter.textContent = formatNumber(target);
    }

    requestAnimationFrame(tick);
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("ru-RU", {
      maximumFractionDigits: 0,
    }).format(Math.round(value));
  }

  function prefersReducedMotion() {
    return Boolean(global.matchMedia && global.matchMedia("(prefers-reduced-motion: reduce)").matches);
  }

  function isNearViewport(element) {
    const rect = element.getBoundingClientRect();
    const viewportHeight = global.innerHeight || document.documentElement.clientHeight;

    return rect.top < viewportHeight * 0.94 && rect.bottom > viewportHeight * 0.02;
  }

  App.Motion = {
    init,
    refresh,
  };
})(window);
