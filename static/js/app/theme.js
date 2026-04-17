(function () {
  var STORAGE_KEY = "ppt-report-theme";

  function getTheme() {
    var t = document.documentElement.getAttribute("data-theme");
    return t === "light" ? "light" : "dark";
  }

  function applyTheme(theme) {
    if (theme !== "light" && theme !== "dark") return;
    document.documentElement.setAttribute("data-theme", theme);
    try {
      localStorage.setItem(STORAGE_KEY, theme);
    } catch (e) {}
    var meta = document.getElementById("meta-theme-color");
    if (meta) meta.setAttribute("content", theme === "light" ? "#ffffff" : "#0f2744");
    updateToggleUi();
  }

  function updateToggleUi() {
    var btn = document.getElementById("theme-toggle");
    if (!btn) return;
    var isDark = getTheme() === "dark";
    btn.setAttribute("aria-label", isDark ? "切换为浅色主题" : "切换为深色主题");
    btn.setAttribute("title", btn.getAttribute("aria-label"));
  }

  window.pptReportTheme = {
    apply: applyTheme,
    toggle: function () {
      applyTheme(getTheme() === "dark" ? "light" : "dark");
    },
    get: getTheme,
  };

  document.addEventListener("DOMContentLoaded", function () {
    var meta = document.getElementById("meta-theme-color");
    if (meta) meta.setAttribute("content", getTheme() === "light" ? "#ffffff" : "#0f2744");
    var btn = document.getElementById("theme-toggle");
    if (btn) {
      btn.addEventListener("click", function () {
        window.pptReportTheme.toggle();
      });
    }
    updateToggleUi();
  });
})();
