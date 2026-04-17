/**
 * 文档生成页：从数据库列表选择模板并挂载生成表单
 */
(function () {
  var _allPresentationItems = [];

  function readWorkspaceTabFromUrl() {
    var m = new RegExp("[?&]gen=([^&#]*)").exec(window.location.search);
    var v = m ? decodeURIComponent(m[1].replace(/\+/g, " ")) : "";
    return v === "word" ? "word" : "ppt";
  }

  if (typeof window.__generateWorkspaceTab === "undefined") {
    window.__generateWorkspaceTab = readWorkspaceTabFromUrl();
  }

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

  function isWordTemplateItem(it) {
    var kind = it && it.template_kind ? String(it.template_kind) : "";
    return kind.indexOf("word_") === 0;
  }

  /** PPT / Word 顶部 Tab 决定模板下拉与空表单的类型，与表单内 data-gen-mode 独立。 */
  function currentGenerateMode() {
    return window.__generateWorkspaceTab === "word" ? "word" : "ppt";
  }

  function syncWorkspaceTabUI() {
    var tab = currentGenerateMode();
    document.querySelectorAll("[data-gen-workspace]").forEach(function (btn) {
      var isWord = btn.getAttribute("data-gen-workspace") === "word";
      var on = (tab === "word") === isWord;
      btn.classList.toggle("is-active", on);
      btn.setAttribute("aria-selected", on ? "true" : "false");
    });
    document.querySelectorAll("[data-gen-workspace-panel]").forEach(function (el) {
      var forTab = el.getAttribute("data-gen-workspace-panel");
      el.hidden = forTab !== tab;
    });
  }

  function populateGenerateSelect(items, sel, mode) {
    if (!sel) return;
    var current = sel.value;
    sel.innerHTML = '<option value="">— 请选择 —</option>';
    var filtered = (items || []).filter(function (it) {
      if (mode === "word") return isWordTemplateItem(it);
      return !isWordTemplateItem(it);
    });
    filtered.forEach(function (it) {
      if (!it.task_id) return;
      var o = document.createElement("option");
      o.value = it.task_id;
      var base = it.file_name || it.task_id;
      var label =
        mode === "word"
          ? base
          : base + " · " + (it.slide_count != null ? it.slide_count + " 页" : "");
      o.textContent = label.length > 80 ? label.slice(0, 77) + "…" : label;
      sel.appendChild(o);
    });
    if (current && Array.from(sel.options).some(function (o) { return o.value === current; })) {
      sel.value = current;
    }
  }

  function refreshGenerateTaskOptions() {
    var sel = document.getElementById("generate-task-select");
    if (!sel) return;
    populateGenerateSelect(_allPresentationItems, sel, currentGenerateMode());
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

  /** 在替换生成表单前读取当前报告类型 / 学生数据，供切换 PPT 后写回。 */
  function readGenerateRefPickSnapshot() {
    var root = document.getElementById("generate-panel-root");
    var form = root && root.querySelector("#ai-generate-form");
    if (!form) return null;
    var ch = form.querySelector("#gen-chapter-template-id");
    var st = form.querySelector("#gen-student-data-id");
    var chapterId = (ch && ch.value) ? String(ch.value).trim() : "";
    var studentId = (st && st.value) ? String(st.value).trim() : "";
    if (!chapterId && !studentId) return null;
    return {
      chapterId: chapterId,
      chapterLabel: (ch && ch.getAttribute("data-label")) || "",
      studentId: studentId,
      studentLabel: (st && st.getAttribute("data-label")) || "",
    };
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
    var refPickSnap = readGenerateRefPickSnapshot();
    clearGenerateRefMounts();
    clearGenerateTopicFoldAnchor();
    if (errEl) {
      errEl.hidden = true;
      errEl.textContent = "";
    }
    if (hint) hint.hidden = true;
    if (mountResult) mountResult.innerHTML = "";
    var tid = (taskId || "").trim();
    var genParam = currentGenerateMode() === "word" ? "word" : "ppt";
    root.innerHTML = '<p class="muted">正在加载生成表单…</p>';
    try {
      var res = await fetch(
        "/partials/generate-form?task_id=" +
          encodeURIComponent(tid) +
          "&gen=" +
          encodeURIComponent(genParam),
      );
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
          var msg = "已选择：" + (meta.file_name || tid);
          if (genParam !== "word") {
            msg +=
              "（" +
              (meta.slide_count != null ? meta.slide_count + " 页" : "? 页") +
              "）";
          }
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
      var skipDefaultChapter = !!(refPickSnap && refPickSnap.chapterId);
      var isWordWorkspace = currentGenerateMode() === "word";
      if (
        tid &&
        !isWordWorkspace &&
        window.PptApp &&
        typeof window.PptApp.applyDefaultChapterTemplate === "function" &&
        !skipDefaultChapter
      ) {
        await window.PptApp.applyDefaultChapterTemplate(form);
      }
      if (
        form &&
        window.PptApp &&
        typeof window.PptApp.setGenerateMode === "function"
      ) {
        var chEl = form.querySelector("#gen-chapter-template-id");
        var hasPick = chEl && String(chEl.value || "").trim();
        if (!hasPick) {
          window.PptApp.setGenerateMode(isWordWorkspace ? "word" : "ppt");
        }
      }
      if (window.PptApp && typeof window.PptApp.restoreGenerateRefPicks === "function") {
        await window.PptApp.restoreGenerateRefPicks(form, refPickSnap);
      }
    } catch (e) {
      root.innerHTML = '<p class="error">加载失败：' + (e.message || String(e)) + "</p>";
      setGenerateRefTabPlaceholders(true);
    }
  }

  function onWorkspaceTabChange(nextTab) {
    var next = nextTab === "word" ? "word" : "ppt";
    if (window.__generateWorkspaceTab === next) return;
    window.__generateWorkspaceTab = next;
    setQueryParam("gen", next === "word" ? "word" : "ppt");
    syncWorkspaceTabUI();
    if (window.PptApp && typeof window.PptApp.clearChapterTemplatePick === "function") {
      window.PptApp.clearChapterTemplatePick();
    }
    refreshGenerateTaskOptions();
    var sel = document.getElementById("generate-task-select");
    var v = sel ? (sel.value || "").trim() : "";
    if (!v) setQueryParam("task_id", "");
    loadGenerateForm(v);
  }

  async function init() {
    window.__generateWorkspaceTab = readWorkspaceTabFromUrl();
    syncWorkspaceTabUI();

    var sel = document.getElementById("generate-task-select");
    if (!sel) return;

    document.querySelectorAll("[data-gen-workspace]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var w = btn.getAttribute("data-gen-workspace") === "word" ? "word" : "ppt";
        onWorkspaceTabChange(w);
      });
    });

    try {
      var res = await fetch("/api/presentations");
      var data = await res.json();
      if (!data.ok) throw new Error(data.error || "列表加载失败");
      _allPresentationItems = data.items || [];
      populateGenerateSelect(_allPresentationItems, sel, currentGenerateMode());
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
  window.PptApp = window.PptApp || {};
  window.PptApp.refreshGenerateTaskOptions = refreshGenerateTaskOptions;
  window.PptApp.syncGenerateWorkspaceTabUI = syncWorkspaceTabUI;
  window.PptApp.getWorkspaceTab = currentGenerateMode;
})();
