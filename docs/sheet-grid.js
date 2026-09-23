(function () {
  var COL_WIDTH = 96;
  var ROW_HEIGHT = 24;

  function toColumnLabel(n) {
    var label = "";
    while (n >= 0) {
      label = String.fromCharCode((n % 26) + 65) + label;
      n = Math.floor(n / 26) - 1;
    }
    return label;
  }

  function render() {
    var colContainer = document.getElementById("sheet-col-headers");
    var rowContainer = document.getElementById("sheet-row-headers");
    if (!colContainer || !rowContainer) return;

    var cols = Math.ceil(window.innerWidth / COL_WIDTH) + 1;
    var rows = Math.ceil(window.innerHeight / ROW_HEIGHT) + 1;

    var colHtml = "";
    for (var c = 0; c < cols; c++) {
      colHtml += "<span>" + toColumnLabel(c) + "</span>";
    }
    colContainer.innerHTML = colHtml;

    var rowHtml = "";
    for (var r = 1; r <= rows; r++) {
      rowHtml += "<span>" + r + "</span>";
    }
    rowContainer.innerHTML = rowHtml;
  }

  document.addEventListener("DOMContentLoaded", render);
  window.addEventListener("resize", render);
})();
