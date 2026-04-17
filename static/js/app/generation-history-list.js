/**
 * 生成历史列表（服务端）
 */
(function () {
  var _currentQuery = "";
  var _currentPage = 1;

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

  function shortId(tid) {
    if (!tid) return "—";
    return tid.length > 10 ? tid.slice(0, 6) + "…" : tid;
  }

  async function render() {
    var mount = document.getElementById("gh-list-mount");
    if (!mount || !window.GenerationHistoryStore) return;

    mount.innerHTML = '<p class="muted">加载中…</p>';

    try {
      var pack = await GenerationHistoryStore.fetchList(_currentQuery, _currentPage);
      if (!pack.dbEnabled) {
        mount.innerHTML =
          '<p class="muted">数据库未启用，无法加载服务端历史。请配置 DATABASE_URL 并重启应用。</p>';
        return;
      }
      _currentPage = pack.page || _currentPage;
      var items = pack.items || [];
      var total = pack.total != null ? pack.total : 0;
      var totalPages = pack.totalPages != null ? pack.totalPages : 0;
      if (!items.length) {
        mount.innerHTML =
          '<p class="muted">' +
            (_currentQuery ? "未找到匹配记录。" : total === 0 ? "暂无生成记录。" : "暂无本页数据。") +
          "</p>";
        return;
      }

      var rows = items
        .map(function (it) {
          return (
            "<tr>" +
            "<td>" +
            esc(formatTime(it.createdAt)) +
            "</td><td>" +
            esc(it.topic || "—") +
            "</td><td><code style=\"font-size:12px\">" +
            esc(shortId(it.taskId)) +
            "</code></td><td>" +
            esc(it.slideCount) +
            "</td><td>" +
            esc((it.outputKind || "pptx").toUpperCase()) +
            '</td><td class="col-actions">' +
            '<a href="/history/' +
            encodeURIComponent(it.id) +
            '">查看</a> · ' +
            '<button type="button" class="buttonlink js-gh-dl" data-id="' +
            esc(it.id) +
            '">下载</button> · ' +
            '<button type="button" class="buttonlink danger js-gh-del" data-id="' +
            esc(it.id) +
            '">删除</button>' +
            "</td></tr>"
          );
        })
        .join("");

      var pg = _currentPage;
      var prevDis = pg <= 1 ? " disabled" : "";
      var nextDis = totalPages <= 0 || pg >= totalPages ? " disabled" : "";
      var pager =
        '<div class="gh-pagination">' +
        '<span class="gh-pagination-meta muted">共 ' +
        esc(String(total)) +
        " 条 · 每页 " +
        esc(String(pack.perPage || 10)) +
        " 条</span>" +
        '<div class="gh-pagination-actions">' +
        '<button type="button" class="button secondary js-gh-page-prev"' +
        prevDis +
        ">上一页</button>" +
        '<span class="gh-pagination-page muted">第 ' +
        esc(String(pg)) +
        " / " +
        esc(String(Math.max(totalPages, 1))) +
        " 页</span>" +
        '<button type="button" class="button secondary js-gh-page-next"' +
        nextDis +
        ">下一页</button>" +
        "</div></div>";

      mount.innerHTML =
        '<table class="presentation-table"><thead><tr><th>时间</th><th>主题</th><th>模板任务</th><th>生成页数</th><th>产物</th><th>操作</th></tr></thead><tbody>' +
        rows +
        "</tbody></table>" +
        pager;
    } catch (e) {
      mount.innerHTML =
        '<p class="error">' + esc(e.message || String(e)) + "</p>";
    }
  }

  async function downloadById(id) {
    if (!window.GenerationHistoryStore) return;
    try {
      await GenerationHistoryStore.downloadHistoryById(id);
    } catch (e) {
      alert(e.message || String(e));
    }
  }

  function init() {
    var wrap = document.getElementById("gh-list-wrap");
    if (!wrap) return;
    var searchInput = document.getElementById("gh-search-input");
    var searchBtn = document.getElementById("gh-search-btn");
    var cleanupSelect = document.getElementById("gh-cleanup-days");
    var cleanupBtn = document.getElementById("gh-cleanup-btn");
    function doSearch() {
      _currentQuery = (searchInput && searchInput.value ? searchInput.value : "").trim();
      _currentPage = 1;
      render();
    }
    if (searchBtn) {
      searchBtn.addEventListener("click", doSearch);
    }
    if (searchInput) {
      searchInput.addEventListener("keydown", function (ev) {
        if (ev.key === "Enter") {
          ev.preventDefault();
          doSearch();
        }
      });
    }
    if (cleanupBtn) {
      cleanupBtn.addEventListener("click", function () {
        var days = Number((cleanupSelect && cleanupSelect.value) || 0);
        if (!(days === 3 || days === 7 || days === 15)) return;
        if (!confirm("确认清理超过 " + days + " 天的历史记录吗？")) return;
        GenerationHistoryStore.cleanupByDays(days)
          .then(function (ret) {
            alert("清理完成，共删除 " + String(ret.removed || 0) + " 条记录。");
            return render();
          })
          .catch(function (e) {
            alert(e.message || String(e));
          });
      });
    }
    wrap.addEventListener("click", function (ev) {
      var prev = ev.target.closest(".js-gh-page-prev");
      if (prev && !prev.disabled) {
        _currentPage = Math.max(1, _currentPage - 1);
        render();
        return;
      }
      var next = ev.target.closest(".js-gh-page-next");
      if (next && !next.disabled) {
        _currentPage = _currentPage + 1;
        render();
        return;
      }
      var del = ev.target.closest(".js-gh-del");
      if (del) {
        var id = del.getAttribute("data-id");
        if (id && confirm("确定删除该条历史？")) {
          GenerationHistoryStore.remove(id)
            .then(function () {
              return render();
            })
            .catch(function (e) {
              alert(e.message || String(e));
            });
        }
        return;
      }
      var dl = ev.target.closest(".js-gh-dl");
      if (dl) {
        downloadById(dl.getAttribute("data-id"));
      }
    });
    render();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
