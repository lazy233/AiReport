/**
 * 章节模板列表页：查（列表+搜索）/ 增（新建）/ 改（编辑）/ 删
 */
(function () {
  function esc(s) {
    var d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  function escapeAttr(s) {
    return String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;")
      .replace(/</g, "&lt;");
  }

  function formatTime(iso) {
    if (!iso) return "—";
    try {
      var d = new Date(iso);
      if (isNaN(d.getTime())) return iso;
      return d.toLocaleString("zh-CN", { hour12: false });
    } catch (e) {
      return iso;
    }
  }

  function getSearchQuery() {
    var el = document.getElementById("ct-list-q");
    return el ? (el.value || "").trim() : "";
  }

  async function deleteTemplate(it) {
    var name = (it && it.name) || "该模板";
    if (
      !window.confirm(
        "确定删除章节模板「" + name + "」吗？删除后不可恢复，且不影响已生成的历史文档。",
      )
    ) {
      return;
    }
    try {
      await ChapterTemplatesStore.remove(it.id);
      await render();
    } catch (e) {
      window.alert(e.message || String(e));
    }
  }

  async function render() {
    var mount = document.getElementById("ct-list-mount");
    if (!mount || !window.ChapterTemplatesStore) return;

    var q = getSearchQuery();
    var items = [];
    try {
      items = await ChapterTemplatesStore.list(q);
    } catch (e) {
      mount.innerHTML = '<p class="error">加载失败：' + esc(e.message || String(e)) + "</p>";
      return;
    }
    if (!items.length) {
      mount.innerHTML =
        '<p class="muted">' +
        (q ? "没有符合搜索条件的模板，可清空搜索后重试。" : "暂无章节模板，点击右上角「新建模板」创建第一个。") +
        "</p>";
      return;
    }

    var rows = items
      .map(function (it) {
        var desc = (it.description || "").trim();
        return (
          "<tr>" +
          "<td>" +
          esc(it.name || "（未命名）") +
          "</td>" +
          "<td class=\"muted\" style=\"max-width: 280px;\">" +
          (desc ? esc(desc) : "—") +
          "</td>" +
          "<td>" +
          esc(it.chapterCount) +
          "</td>" +
          "<td>" +
          esc(formatTime(it.updatedAt)) +
          '</td><td class="col-actions">' +
          '<a href="/chapter-templates/' +
          encodeURIComponent(it.id) +
          '">详情</a> · ' +
          '<a href="/chapter-templates/' +
          encodeURIComponent(it.id) +
          '/edit">编辑</a> · ' +
          '<button type="button" class="link-button ct-list-delete-btn" data-ct-id="' +
          escapeAttr(it.id || "") +
          '" data-ct-name="' +
          escapeAttr(it.name || "") +
          '">删除</button>' +
          "</td></tr>"
        );
      })
      .join("");

    mount.innerHTML =
      '<table class="presentation-table"><thead><tr><th>模板名称</th><th>说明</th><th>章节数</th><th>更新时间</th><th>操作</th></tr></thead><tbody>' +
      rows +
      "</tbody></table>";

    mount.querySelectorAll(".ct-list-delete-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var id = (btn.getAttribute("data-ct-id") || "").trim();
        var nm = btn.getAttribute("data-ct-name") || "";
        if (!id) return;
        deleteTemplate({ id: id, name: nm });
      });
    });
  }

  function bindToolbar() {
    var searchBtn = document.getElementById("ct-list-search");
    var refreshBtn = document.getElementById("ct-list-refresh");
    var qInput = document.getElementById("ct-list-q");
    if (searchBtn) searchBtn.addEventListener("click", function () { render(); });
    if (refreshBtn) refreshBtn.addEventListener("click", function () { render(); });
    if (qInput) {
      qInput.addEventListener("keydown", function (e) {
        if (e.key === "Enter") {
          e.preventDefault();
          render();
        }
      });
    }
  }

  function init() {
    bindToolbar();
    render();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
