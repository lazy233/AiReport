/**
 * 上传页：已解析列表（数据库）— 查 / 改名 / 替换 / 删
 */
(function () {
  function getQueryParam(name) {
    var m = new RegExp("[?&]" + name + "=([^&#]*)").exec(window.location.search);
    return m ? decodeURIComponent(m[1].replace(/\+/g, " ")) : "";
  }

  function setQueryParam(name, value) {
    var url = new URL(window.location.href);
    if (value === "" || value === null || typeof value === "undefined") {
      url.searchParams.delete(name);
    } else {
      url.searchParams.set(name, value);
    }
    window.history.replaceState({}, "", url.pathname + url.search + url.hash);
  }

  function escapeAttr(s) {
    return String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;")
      .replace(/</g, "&lt;");
  }

  function sleep(ms) {
    return window.PptApp && PptApp.sleep
      ? PptApp.sleep(ms)
      : new Promise(function (r) {
          setTimeout(r, ms);
        });
  }

  async function pollParseJobUntilDone(jobId) {
    for (;;) {
      var r = await fetch("/api/parse_status/" + encodeURIComponent(jobId));
      var s = await r.json();
      if (!r.ok || !s.ok) throw new Error(s.error || "状态查询失败");
      if (s.status === "done" && s.task_id) return s.task_id;
      if (s.status === "error") throw new Error(s.error || s.message || "解析失败");
      await sleep(400);
    }
  }

  async function renamePresentation(taskId, currentName) {
    var msg = window.prompt(
      "修改为显示名称（仅保存到数据库，用于列表与导出提示；不会重命名服务器上的 .pptx 文件）：",
      currentName || "",
    );
    if (msg === null) return;
    var name = (msg || "").trim();
    if (!name) {
      window.alert("名称不能为空。");
      return;
    }
    try {
      var res = await fetch("/api/presentations/" + encodeURIComponent(taskId), {
        method: "PATCH",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify({ file_name: name }),
      });
      var data = await res.json().catch(function () {
        return {};
      });
      if (!res.ok || !data.ok) throw new Error(data.error || "保存失败");
      loadPresentationList();
    } catch (e) {
      window.alert(e.message || String(e));
    }
  }

  async function deletePresentation(taskId, displayName) {
    var label = displayName || taskId;
    if (!window.confirm("确定删除模板「" + label + "」吗？将同时删除解析记录、服务器上的 .pptx 及该任务的参考截图，且不可恢复。")) {
      return;
    }
    try {
      var res = await fetch("/api/presentations/" + encodeURIComponent(taskId), { method: "DELETE" });
      var data = await res.json().catch(function () {
        return {};
      });
      if (!res.ok || !data.ok) throw new Error(data.error || "删除失败");
      loadPresentationList();
    } catch (e) {
      window.alert(e.message || String(e));
    }
  }

  function bindReplaceFileInput() {
    var input = document.getElementById("presentation-replace-file");
    if (!input || input.dataset.boundReplace === "1") return;
    input.dataset.boundReplace = "1";
    input.addEventListener("change", async function () {
      var taskId = (input.dataset.replaceTaskId || "").trim();
      input.dataset.replaceTaskId = "";
      var f = input.files && input.files[0];
      input.value = "";
      if (!f || !taskId) return;
      var lower = (f.name || "").toLowerCase();
      if (!lower.endsWith(".pptx")) {
        window.alert("请仅选择 .pptx 文件。");
        return;
      }
      var loading = document.getElementById("presentation-list-loading");
      var wrap = document.getElementById("presentation-list-table");
      var empty = document.getElementById("presentation-list-empty");
      if (loading) {
        loading.hidden = false;
        loading.textContent = "正在上传并重新解析（替换模板文件）…";
      }
      if (wrap) wrap.hidden = true;
      if (empty) empty.hidden = true;
      try {
        var fd = new FormData();
        fd.append("ppt_file", f);
        fd.append("replace_task_id", taskId);
        var res = await fetch("/api/parse_start", { method: "POST", body: fd });
        var start = await res.json().catch(function () {
          return {};
        });
        if (!res.ok || !start.ok) throw new Error(start.error || "启动替换解析失败");
        await pollParseJobUntilDone(start.job_id);
        await loadPresentationList();
        var flash = document.getElementById("parse-flash");
        if (flash) {
          flash.hidden = false;
          flash.textContent = "已用新文件替换并重新解析：「" + (f.name || "") + "」。";
        }
      } catch (e) {
        window.alert(e.message || String(e));
        await loadPresentationList();
      }
    });
  }

  function startReplacePresentation(taskId) {
    bindReplaceFileInput();
    var input = document.getElementById("presentation-replace-file");
    if (!input) return;
    input.dataset.replaceTaskId = taskId;
    input.click();
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

  function renderTable(items) {
    var wrap = document.getElementById("presentation-list-table");
    var loading = document.getElementById("presentation-list-loading");
    var empty = document.getElementById("presentation-list-empty");
    if (!wrap) return;
    if (loading) loading.hidden = true;
    if (!items || !items.length) {
      wrap.hidden = true;
      if (empty) empty.hidden = false;
      return;
    }
    if (empty) empty.hidden = true;
    wrap.hidden = false;
    var rows = items
      .map(function (it) {
        var name = (it.file_name || "(未命名)").replace(/</g, "&lt;");
        var rawName = it.file_name || "";
        var impl = (it.parse_impl || "—").replace(/</g, "&lt;");
        var tid = it.task_id || "";
        return (
          "<tr>" +
          "<td>" +
          name +
          "</td><td>" +
          (it.slide_count != null ? it.slide_count : "—") +
          "</td><td>" +
          impl +
          "</td><td>" +
          formatTime(it.created_at) +
          '</td><td class="col-actions">' +
          '<a href="/presentations/' +
          encodeURIComponent(tid) +
          '">详情</a> · ' +
          '<button type="button" class="link-button presentation-rename-btn" data-presentation-rename="' +
          escapeAttr(tid) +
          '" data-current-name="' +
          escapeAttr(rawName) +
          '">改名</button> · ' +
          '<button type="button" class="link-button presentation-replace-btn" data-presentation-replace="' +
          escapeAttr(tid) +
          '">替换</button> · ' +
          '<button type="button" class="link-button presentation-delete-btn" data-presentation-delete="' +
          escapeAttr(tid) +
          '" data-display-name="' +
          escapeAttr(rawName) +
          '">删除</button>' +
          "</td></tr>"
        );
      })
      .join("");
    wrap.innerHTML =
      '<table class="presentation-table"><thead><tr><th>文件名</th><th>页数</th><th>解析实现</th><th>解析时间</th><th>操作</th></tr></thead><tbody>' +
      rows +
      "</tbody></table>";

    wrap.querySelectorAll(".presentation-rename-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var tid = (btn.getAttribute("data-presentation-rename") || "").trim();
        var cur = btn.getAttribute("data-current-name") || "";
        renamePresentation(tid, cur);
      });
    });
    wrap.querySelectorAll(".presentation-replace-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var tid = (btn.getAttribute("data-presentation-replace") || "").trim();
        if (tid) startReplacePresentation(tid);
      });
    });
    wrap.querySelectorAll(".presentation-delete-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var tid = (btn.getAttribute("data-presentation-delete") || "").trim();
        var dn = btn.getAttribute("data-display-name") || "";
        if (tid) deletePresentation(tid, dn);
      });
    });
  }

  async function loadPresentationList() {
    var loading = document.getElementById("presentation-list-loading");
    var wrap = document.getElementById("presentation-list-table");
    var empty = document.getElementById("presentation-list-empty");
    if (loading) {
      loading.hidden = false;
      loading.textContent = "加载列表中…";
    }
    if (wrap) wrap.hidden = true;
    if (empty) empty.hidden = true;
    try {
      var res = await fetch("/api/presentations");
      var data = await res.json();
      if (!data.ok) throw new Error(data.error || "列表加载失败");
      renderTable(data.items || []);
      var doneId = getQueryParam("parsed");
      var doneName = getQueryParam("parsed_name");
      if (doneId) {
        var flash = document.getElementById("parse-flash");
        if (flash) {
          flash.hidden = false;
          flash.textContent = doneName
            ? "解析已完成并保存：「" + doneName + "」。可在下表打开「详情」或前往文档生成页选用。"
            : "解析已完成并保存。可在下表打开「详情」查看完整结构。";
        }
        setQueryParam("parsed", "");
        setQueryParam("parsed_name", "");
      }
    } catch (e) {
      if (loading) {
        loading.hidden = false;
        loading.textContent = "加载失败：" + (e.message || String(e));
      }
    }
  }

  function init() {
    bindReplaceFileInput();
    var refresh = document.getElementById("presentation-list-refresh");
    if (refresh) refresh.addEventListener("click", loadPresentationList);
    loadPresentationList();
  }

  window.PptApp = window.PptApp || {};
  window.PptApp.loadPresentationList = loadPresentationList;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
