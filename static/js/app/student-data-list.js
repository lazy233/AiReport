/**
 * 数据管理列表：查询、表格、批量上传
 */
(function () {
  function esc(s) {
    var d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  function formatTime(iso) {
    if (!iso) return "—";
    try {
      var dt = new Date(iso);
      if (isNaN(dt.getTime())) return iso;
      return dt.toLocaleString("zh-CN", { hour12: false });
    } catch (e) {
      return iso;
    }
  }

  function getQueryInput() {
    return document.getElementById("sd-search-input");
  }

  function currentQuery() {
    var el = getQueryInput();
    return el ? el.value.trim() : "";
  }

  async function render() {
    var mount = document.getElementById("sd-list-mount");
    if (!mount || !window.StudentDataStore) return;

    var filtered = [];
    var q = currentQuery();
    try {
      filtered = await StudentDataStore.list(q);
    } catch (e) {
      mount.innerHTML = '<p class="error">加载失败：' + esc(e.message || String(e)) + "</p>";
      return;
    }

    if (!filtered.length) {
      mount.innerHTML =
        '<p class="muted">' + (q ? "无匹配结果。" : "暂无数据。") + "</p>";
      return;
    }

    var rows = filtered
      .map(function (it) {
        return (
          "<tr>" +
          "<td>" +
          esc(it.name || "—") +
          "</td>" +
          "<td>" +
          esc(it.studentId || "—") +
          "</td>" +
          "<td>" +
          esc(it.className || "—") +
          "</td>" +
          "<td>" +
          esc(formatTime(it.updatedAt)) +
          '</td><td class="col-actions">' +
          '<a href="/student-data/' +
          encodeURIComponent(it.id) +
          '">详情</a> · ' +
          '<a href="/student-data/' +
          encodeURIComponent(it.id) +
          '/edit">编辑</a>' +
          "</td></tr>"
        );
      })
      .join("");

    mount.innerHTML =
      '<table class="presentation-table sd-table"><thead><tr><th>姓名</th><th>学号</th><th>班级</th><th>更新时间</th><th>操作</th></tr></thead><tbody>' +
      rows +
      "</tbody></table>";
  }

  function showBulkMsg(html, isError) {
    var el = document.getElementById("sd-bulk-message");
    if (!el) return;
    el.hidden = false;
    el.className = isError ? "error" : "muted";
    el.style.marginTop = "12px";
    el.innerHTML = html;
  }

  function init() {
    var search = getQueryInput();
    var doSearch = document.getElementById("sd-search-btn");
    var refresh = document.getElementById("sd-refresh-btn");

    function runSearch() {
      render();
    }

    if (search) {
      search.addEventListener("keydown", function (e) {
        if (e.key === "Enter") runSearch();
      });
    }
    if (doSearch) doSearch.addEventListener("click", runSearch);
    if (refresh)
      refresh.addEventListener("click", function () {
        if (search) search.value = "";
        runSearch();
      });

    var fileInput = document.getElementById("sd-import-ai-file");
    if (fileInput && window.StudentDataStore && typeof StudentDataStore.importAiFile === "function") {
      fileInput.addEventListener("change", async function () {
        var f = fileInput.files && fileInput.files[0];
        if (!f) return;
        showBulkMsg("正在上传并由大模型解析，请稍候…", false);
        try {
          var res = await StudentDataStore.importAiFile(f);
          var parts = [
            "已写入 <strong>" +
              esc(String(res.imported != null ? res.imported : "0")) +
              "</strong> 条（有效数据行约 " +
              esc(String(res.totalRows != null ? res.totalRows : "—")) +
              " 行，分 " +
              esc(String(res.batches != null ? res.batches : "—")) +
              " 批，每批 " +
              esc(String(res.batchSize != null ? res.batchSize : "10")) +
              " 条）。"
          ];
          if (res.errors && res.errors.length) {
            parts.push(
              '<span class="error">提示：' + esc(res.errors.slice(0, 8).join("；")) + "</span>"
            );
          }
          showBulkMsg(parts.join(" "), !!(res.errors && res.errors.length));
          fileInput.value = "";
          await render();
        } catch (e) {
          showBulkMsg("智能导入失败：" + esc(e.message || String(e)), true);
          fileInput.value = "";
        }
      });
    }

    render();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
