(function () {
  var STORAGE_KEY = "xlsxgrep-theme";
  var root = document.documentElement;
  var saved = localStorage.getItem(STORAGE_KEY) || "default";
  root.setAttribute("data-theme", saved);

  document.addEventListener("DOMContentLoaded", function () {
    var select = document.getElementById("theme-select");
    if (!select) return;
    select.value = saved;
    select.addEventListener("change", function () {
      var theme = select.value;
      root.setAttribute("data-theme", theme);
      localStorage.setItem(STORAGE_KEY, theme);
    });
  });
})();
