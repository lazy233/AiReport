/**
 * 文档生成页：从数据库列表选择模板并挂载生成表单
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

  function populateGenerateSelect(items, sel) {
    if (!sel) return;
    var current = sel.value;
    sel.innerHTML = '<option value="">— 请选择 —</option>';
    (items || []).forEach(function (it) {
      if (!it.task_id) return;
      var o = document.createElement("option");
      o.value = it.task_id;
      var label =
        (it.file_name || it.task_id) + " · " + (it.slide_count != null ? it.slide_count + " 页" : "");
      o.textContent = label.length > 80 ? label.slice(0, 77) + "…" : label;
      sel.appendChild(o);
    });
    if (current && Array.from(sel.options).some(function (o) { return o.value === current; })) {
      sel.value = current;
    }
  }

  function setGenerateRefTabPlaceholders(show) {
    var ch = document.getElementById("gen-ref-chapter-placeholder");
    var st = document.getElementById("gen-ref-student-placeholder");
    if (ch) ch.hidden = !show;
    if (st) st.hidden = !show;
  }

  function clearGenerateRefMounts() {
    ["gen-ref-mount-chapter", "gen-ref-mount-student", "gen-ref-footer-mount"].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) {
        while (el.firstChild) el.removeChild(el.firstChild);
      }
    });
    setGenerateRefTabPlaceholders(true);
  }

  function clearGenerateTopicFoldAnchor() {
    var anchor = document.getElementById("generate-topic-fold-anchor");
    if (!anchor) return;
    anchor.innerHTML = "";
  }

  function mountTopicFoldToAnchor(formRoot) {
    var anchor = document.getElementById("generate-topic-fold-anchor");
    var fold = formRoot && formRoot.querySelector("details.gen-topic-fold");
    if (!anchor || !fold) return;
    anchor.appendChild(fold);
  }

  function mountRefPartsToAnchor(formRoot) {
    var chM = document.getElementById("gen-ref-mount-chapter");
    var stM = document.getElementById("gen-ref-mount-student");
    var ftM = document.getElementById("gen-ref-footer-mount");
    if (!formRoot) return;
    var scope = formRoot.querySelector("#ai-generate-form") || formRoot;
    var tch = scope.querySelector("#gen-ref-toolbar-chapter");
    var tst = scope.querySelector("#gen-ref-toolbar-student");
    var chipCh = scope.querySelector("#gen-ref-pick-chapter-wrap");
    var chipSt = scope.querySelector("#gen-ref-pick-student-wrap");
    var picks = scope.querySelector("#gen-ref-picks");
    if (chM && tch) chM.appendChild(tch);
    if (chM && chipCh) chM.appendChild(chipCh);
    if (stM && tst) stM.appendChild(tst);
    if (stM && chipSt) stM.appendChild(chipSt);
    if (ftM && picks) ftM.appendChild(picks);
    setGenerateRefTabPlaceholders(false);
  }

  async function loadGenerateForm(taskId) {
    var root = document.getElementById("generate-panel-root");
    var errEl = document.getElementById("generate-context-error");
    var hint = document.getElementById("generate-context-hint");
    var mountResult = document.getElementById("ai-result-mount");
    if (!root) return;
    clearGenerateRefMounts();
    clearGenerateTopicFoldAnchor();
    if (errEl) {
      errEl.hidden = true;
      errEl.textContent = "";
    }
    if (hint) hint.hidden = true;
    if (mountResult) mountResult.innerHTML = "";
    var tid = (taskId || "").trim();
    root.innerHTML = '<p class="muted">正在加载生成表单…</p>';
    try {
      var res = await fetch("/partials/generate-form?task_id=" + encodeURIComponent(tid));
      var html = await res.text();
      if (!res.ok) {
        root.innerHTML = html;
        return;
      }
      root.innerHTML = html;
      if (tid) {
        var ctx = await fetch("/api/presentations/" + encodeURIComponent(tid));
        var meta = await ctx.json();
        if (hint && meta.ok) {
          hint.hidden = false;
          var msg =
            "已选择：" +
            (meta.file_name || tid) +
            "（" +
            (meta.slide_count != null ? meta.slide_count + " 页" : "? 页") +
            "）";
          if (meta.has_template === false) {
            msg += "。警告：服务器上未找到对应 .pptx 模板文件，生成后可能无法导出。";
          }
          hint.textContent = msg;
        }
      }
      var form = root.querySelector("#ai-generate-form");
      if (window.PptApp && window.PptApp.bindGenerateForm) {
        window.PptApp.bindGenerateForm(form);
      }
      mountTopicFoldToAnchor(root);
      if (window.PptApp && window.PptApp.syncGenerateRefPicksUi) {
        window.PptApp.syncGenerateRefPicksUi();
      }
      if (window.PptApp && window.PptApp.initChapterReferenceUi) {
        window.PptApp.initChapterReferenceUi(form);
      }
      mountRefPartsToAnchor(root);
      if (
        tid &&
        window.PptApp &&
        typeof window.PptApp.applyDefaultChapterTemplate === "function"
      ) {
        await window.PptApp.applyDefaultChapterTemplate(form);
      }
    } catch (e) {
      root.innerHTML = '<p class="error">加载失败：' + (e.message || String(e)) + "</p>";
      setGenerateRefTabPlaceholders(true);
    }
  }

  async function init() {
    var sel = document.getElementById("generate-task-select");
    if (!sel) return;

    try {
      var res = await fetch("/api/presentations");
      var data = await res.json();
      if (!data.ok) throw new Error(data.error || "列表加载失败");
      populateGenerateSelect(data.items || [], sel);
    } catch (e) {
      var errEl = document.getElementById("generate-context-error");
      if (errEl) {
        errEl.hidden = false;
        errEl.textContent = "加载模板列表失败：" + (e.message || String(e));
      }
    }

    var preId = getQueryParam("task_id");
    if (preId && Array.from(sel.options).some(function (o) { return o.value === preId; })) {
      sel.value = preId;
    }
    loadGenerateForm((sel.value || "").trim());

    sel.addEventListener("change", function () {
      var v = (sel.value || "").trim();
      setQueryParam("task_id", v);
      loadGenerateForm(v);
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
