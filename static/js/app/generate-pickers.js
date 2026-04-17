/**
 * 文档生成：报告类型 / 学生数据 列表选择（调用现有 /api/* 列表接口）
 */
(function () {
  var backdrop = null;
  var currentMode = null;
  var itemsCache = [];
  var WORD_REPORT_TEMPLATE_CODE = "word_table_fill";

  function esc(s) {
    var d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  function getForm() {
    return document.getElementById("ai-generate-form");
  }

  function getGenerateMode(form) {
    var mode = form && form.getAttribute("data-gen-mode");
    return mode === "word" ? "word" : "ppt";
  }

  function setGenerateMode(form, mode) {
    if (!form) return;
    var next = mode === "word" ? "word" : "ppt";
    form.setAttribute("data-gen-mode", next);
    var chapterFold = form.querySelector(".chapter-tabs-fold");
    if (chapterFold) chapterFold.hidden = next === "word";
    var wordNote = form.querySelector("[data-word-mode-note]");
    if (wordNote) wordNote.hidden = next !== "word";
    var selectedSlidesHidden = form.querySelector("#selected-slides-hidden");
    if (selectedSlidesHidden && next === "word") selectedSlidesHidden.innerHTML = "";
    if (window.PptApp && typeof window.PptApp.refreshGenerateTaskOptions === "function") {
      window.PptApp.refreshGenerateTaskOptions();
    }
  }

  /** PPT 解析结果中「章」块数量（kind=chapter，不含首页/目录/其他） */
  function countPptChapterGroups(form) {
    if (!form) return 0;
    return form.querySelectorAll('.group-select[data-group-kind="chapter"]').length;
  }

  /**
   * 校验模板章数不大于 PPT「章」块数；通过后始终勾选全部块（生成范围固定为全部分组）。
   * @returns {boolean}
   */
  function applyChapterTemplateAutoSelect(form, tplCount) {
    if (!form) return false;
    var n = parseInt(tplCount, 10);
    if (!Number.isFinite(n) || n <= 0) {
      window.alert("该报告类型没有有效章节定义。");
      return false;
    }
    var pptCh = countPptChapterGroups(form);
    if (n > pptCh) {
      window.alert(
        "报告类型包含 " +
          n +
          " 个章节，当前 PPT 仅解析出 " +
          pptCh +
          " 个「章」块，数量不匹配。请更换报告类型或更换 PPT。",
      );
      return false;
    }
    form.querySelectorAll(".group-select").forEach(function (cb) {
      cb.checked = true;
    });
    return true;
  }

  var DEFAULT_CHAPTER_TEMPLATE_NAME = "学生学期规划报告";

  async function applyDefaultChapterTemplate(form) {
    if (!form) return;
    if (window.__generateWorkspaceTab === "word") return;
    var el = form.querySelector("#gen-chapter-template-id");
    if (!el) return;
    if ((el.value || "").trim()) return;
    try {
      var res = await fetch("/api/chapter-templates");
      var data = await res.json();
      if (!res.ok || !data.ok || !Array.isArray(data.items)) return;
      var found = data.items.find(function (it) {
        var n = (it.name || "").trim();
        return n === DEFAULT_CHAPTER_TEMPLATE_NAME || n.indexOf(DEFAULT_CHAPTER_TEMPLATE_NAME) !== -1;
      });
      if (!found || !found.id) return;
      await pickChapterTemplate(
        String(found.id),
        (found.name || "").trim() || DEFAULT_CHAPTER_TEMPLATE_NAME,
        String(found.templateCode || "").trim(),
      );
    } catch (e) {
      /* 静默失败，用户可手动选择 */
    }
  }

  /**
   * Word 生成页目前仅支持一种报告类型（word_table_fill）；若库中仅此一条，进入页面时自动选中。
   */
  async function applyDefaultWordReportTemplate(form) {
    if (!form) return;
    if (window.__generateWorkspaceTab !== "word") return;
    var el = form.querySelector("#gen-chapter-template-id");
    if (!el) return;
    if ((el.value || "").trim()) return;
    try {
      var res = await fetch("/api/chapter-templates");
      var data = await res.json();
      if (!res.ok || !data.ok || !Array.isArray(data.items)) return;
      var wordItems = data.items.filter(function (it) {
        return String(it.templateCode || "").trim() === WORD_REPORT_TEMPLATE_CODE;
      });
      if (wordItems.length !== 1) return;
      var found = wordItems[0];
      if (!found.id) return;
      await pickChapterTemplate(
        String(found.id),
        (found.name || "").trim() || "Word 表格回填",
        WORD_REPORT_TEMPLATE_CODE,
      );
    } catch (e) {
      /* 静默失败，用户可手动选择 */
    }
  }

  async function pickChapterTemplate(id, primary, forcedTemplateCode) {
    var form = getForm();
    if (!form || !id) return false;
    try {
      var res = await fetch("/api/chapter-templates/" + encodeURIComponent(id));
      var data = await res.json();
      if (!res.ok || !data.ok) {
        window.alert(data.error || "加载报告类型详情失败");
        return false;
      }
      var item = data.item || {};
      var templateCode = String(forcedTemplateCode || item.templateCode || "").trim();
      var ws = window.__generateWorkspaceTab === "word" ? "word" : "ppt";
      if (ws === "word" && templateCode !== WORD_REPORT_TEMPLATE_CODE) {
        window.alert("当前为 Word 生成：请仅选择「Word 表格回填」报告类型。");
        return false;
      }
      if (ws === "ppt" && templateCode === WORD_REPORT_TEMPLATE_CODE) {
        window.alert("当前为 PPT 生成：请选择含章节的报告类型（非 Word 专用）。");
        return false;
      }
      var isWordMode = templateCode === WORD_REPORT_TEMPLATE_CODE;
      var tplCount = item.chapterCount;
      if (tplCount == null && Array.isArray(item.chapters)) {
        tplCount = item.chapters.length;
      }
      if (!isWordMode && !applyChapterTemplateAutoSelect(form, tplCount)) return false;
      setGenerateMode(form, isWordMode ? "word" : "ppt");
      setChapterTemplate(id, primary);
      if (typeof form._resyncChapters === "function") form._resyncChapters();
      return true;
    } catch (e) {
      window.alert(e.message || String(e));
      return false;
    }
  }

  function syncRefPicksUi() {
    var form = getForm();
    var strip = document.getElementById("gen-ref-picks");
    var wCh = document.getElementById("gen-ref-pick-chapter-wrap");
    var wSt = document.getElementById("gen-ref-pick-student-wrap");
    var lCh = document.getElementById("gen-ref-pick-chapter-label");
    var lSt = document.getElementById("gen-ref-pick-student-label");
    if (!form) return;
    var tid = (form.querySelector("#gen-chapter-template-id") || {}).value || "";
    var sid = (form.querySelector("#gen-student-data-id") || {}).value || "";
    var chName = (form.querySelector("#gen-chapter-template-id") || {}).getAttribute("data-label") || "";
    var stName = (form.querySelector("#gen-student-data-id") || {}).getAttribute("data-label") || "";
    if (lCh) lCh.textContent = tid ? chName || tid : "";
    if (lSt) lSt.textContent = sid ? stName || sid : "";
    if (wCh) wCh.hidden = !tid;
    if (wSt) wSt.hidden = !sid;
    if (strip) strip.hidden = !(tid && sid);
    if (window.PptApp && window.PptApp.syncGenResolveRow) {
      window.PptApp.syncGenResolveRow();
    }
    if (window.PptApp && window.PptApp.syncGenerateSubmitEnabled) {
      window.PptApp.syncGenerateSubmitEnabled();
    }
  }

  function setChapterTemplate(id, label) {
    var form = getForm();
    if (!form) return;
    var el = form.querySelector("#gen-chapter-template-id");
    if (!el) return;
    el.value = id || "";
    if (id) el.setAttribute("data-label", label || "");
    else {
      el.removeAttribute("data-label");
      var ws = window.__generateWorkspaceTab === "word" ? "word" : "ppt";
      setGenerateMode(form, ws);
    }
    syncRefPicksUi();
    if (window.PptApp && window.PptApp.resetChapterReferenceUi) {
      window.PptApp.resetChapterReferenceUi(getForm());
    }
    if (
      id &&
      window.PptApp &&
      typeof window.PptApp.applyDefaultCoverTitleFromTemplateName === "function"
    ) {
      window.PptApp.applyDefaultCoverTitleFromTemplateName(getForm(), label || "");
    }
  }

  function setStudentData(id, label) {
    var form = getForm();
    if (!form) return;
    var el = form.querySelector("#gen-student-data-id");
    if (!el) return;
    el.value = id || "";
    if (id) el.setAttribute("data-label", label || "");
    else el.removeAttribute("data-label");
    syncRefPicksUi();
    if (window.PptApp && window.PptApp.resetChapterReferenceUi) {
      window.PptApp.resetChapterReferenceUi(form);
    }
    var tplEl = form.querySelector("#gen-chapter-template-id");
    var tplLabel = tplEl ? tplEl.getAttribute("data-label") || "" : "";
    if (
      tplLabel.trim() &&
      window.PptApp &&
      typeof window.PptApp.applyDefaultCoverTitleFromTemplateName === "function"
    ) {
      window.PptApp.applyDefaultCoverTitleFromTemplateName(form, tplLabel);
    }
  }

  function ensureModal() {
    if (backdrop) return;
    backdrop = document.createElement("div");
    backdrop.className = "gen-picker-backdrop";
    backdrop.setAttribute("hidden", "");
    backdrop.innerHTML =
      '<div class="gen-picker-modal" role="dialog" aria-modal="true" aria-labelledby="gen-picker-title">' +
      '<div class="gen-picker-head">' +
      '<h3 id="gen-picker-title" class="gen-picker-title"></h3>' +
      '<button type="button" class="button secondary gen-picker-close" aria-label="关闭">关闭</button>' +
      "</div>" +
      '<div class="gen-picker-body">' +
      '<input type="search" class="gen-picker-search" placeholder="筛选列表…" autocomplete="off" />' +
      '<div class="gen-picker-list-wrap">' +
      '<ul class="gen-picker-list" id="gen-picker-list"></ul>' +
      '<p class="muted gen-picker-empty" id="gen-picker-empty" hidden>暂无数据</p>' +
      '<p class="error gen-picker-error" id="gen-picker-error" hidden></p>' +
      '<p class="muted gen-picker-loading" id="gen-picker-loading">加载中…</p>' +
      "</div>" +
      "</div>" +
      "</div>";
    document.body.appendChild(backdrop);

    backdrop.addEventListener("click", function (ev) {
      if (ev.target === backdrop) closeModal();
    });
    backdrop.querySelector(".gen-picker-close").addEventListener("click", closeModal);
    var search = backdrop.querySelector(".gen-picker-search");
    search.addEventListener("input", function () {
      renderList(itemsCache, search.value.trim().toLowerCase());
    });

    document.addEventListener("keydown", function (ev) {
      if (ev.key === "Escape" && backdrop && !backdrop.hidden) {
        closeModal();
      }
    });
  }

  function closeModal() {
    if (!backdrop) return;
    backdrop.setAttribute("hidden", "");
    currentMode = null;
    itemsCache = [];
  }

  function openModal() {
    ensureModal();
    backdrop.removeAttribute("hidden");
    var search = backdrop.querySelector(".gen-picker-search");
    if (search) {
      search.value = "";
      search.focus();
    }
  }

  function renderList(items, q) {
    var list = backdrop.querySelector("#gen-picker-list");
    var empty = backdrop.querySelector("#gen-picker-empty");
    var loading = backdrop.querySelector("#gen-picker-loading");
    var errEl = backdrop.querySelector("#gen-picker-error");
    if (loading) loading.hidden = true;
    if (errEl) errEl.hidden = true;
    if (!list || !empty) return;

    var filtered = items;
    if (q) {
      filtered = items.filter(function (it) {
        var blob = JSON.stringify(it).toLowerCase();
        return blob.indexOf(q) !== -1;
      });
    }

    list.innerHTML = "";
    if (!filtered.length) {
      empty.hidden = false;
      return;
    }
    empty.hidden = true;

    filtered.forEach(function (it) {
      var li = document.createElement("li");
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "gen-picker-row";
      var id = it.id != null ? String(it.id) : "";
      var primary = "";
      var sub = "";
      if (currentMode === "chapter-template") {
        primary = it.name || id;
        var tCode = String(it.templateCode || "").trim();
        if (tCode === WORD_REPORT_TEMPLATE_CODE && window.__generateWorkspaceTab !== "word") {
          primary = "【Word】" + primary;
        }
        sub =
          (it.chapterCount != null ? it.chapterCount + " 章 · " : "") +
          (it.updatedAt ? String(it.updatedAt).slice(0, 10) : "");
      } else {
        primary = it.name || it.studentId || id;
        sub = [it.studentId, it.className].filter(Boolean).join(" · ");
      }
      btn.setAttribute("data-pick-id", id);
      btn.setAttribute("data-pick-label", primary);
      if (currentMode === "chapter-template") {
        btn.setAttribute("data-pick-template-code", String(it.templateCode || "").trim());
      }
      btn.innerHTML =
        '<span class="gen-picker-row-main">' + esc(primary) + "</span>" +
        (sub ? '<span class="gen-picker-row-sub muted">' + esc(sub) + "</span>" : "");
      btn.addEventListener("click", async function () {
        if (currentMode === "chapter-template") {
          var form = getForm();
          if (!form) return;
          btn.disabled = true;
          try {
            var templateCode = btn.getAttribute("data-pick-template-code") || "";
            var ok = await pickChapterTemplate(id, primary, templateCode);
            if (ok) closeModal();
          } finally {
            btn.disabled = false;
          }
        } else {
          setStudentData(id, primary);
          closeModal();
        }
      });
      li.appendChild(btn);
      list.appendChild(li);
    });
  }

  async function loadAndShow(mode) {
    ensureModal();
    currentMode = mode;
    var titleEl = backdrop.querySelector("#gen-picker-title");
    var loading = backdrop.querySelector("#gen-picker-loading");
    var errEl = backdrop.querySelector("#gen-picker-error");
    var list = backdrop.querySelector("#gen-picker-list");
    var empty = backdrop.querySelector("#gen-picker-empty");
    var search = backdrop.querySelector(".gen-picker-search");

    if (titleEl) {
      titleEl.textContent = mode === "chapter-template" ? "选择报告类型" : "选择学生数据";
    }
    if (search) search.value = "";
    if (list) list.innerHTML = "";
    if (empty) empty.hidden = true;
    if (errEl) {
      errEl.hidden = true;
      errEl.textContent = "";
    }
    if (loading) loading.hidden = false;
    openModal();

    var url = mode === "chapter-template" ? "/api/chapter-templates" : "/api/student-data";
    try {
      var res = await fetch(url);
      var data = await res.json();
      if (!res.ok || !data.ok) {
        throw new Error(data.error || "加载失败");
      }
      itemsCache = data.items || [];
      if (mode === "chapter-template") {
        itemsCache = itemsCache.filter(function (it) {
          var code = String(it.templateCode || "").trim();
          var wspace = window.__generateWorkspaceTab === "word" ? "word" : "ppt";
          if (wspace === "word") return code === WORD_REPORT_TEMPLATE_CODE;
          return code !== WORD_REPORT_TEMPLATE_CODE;
        });
      }
      if (loading) loading.hidden = true;
      renderList(itemsCache, "");
    } catch (e) {
      if (loading) loading.hidden = true;
      if (errEl) {
        errEl.hidden = false;
        errEl.textContent = e.message || String(e);
      }
    }
  }

  document.body.addEventListener("click", function (ev) {
    var pick = ev.target.closest("[data-gen-picker]");
    if (pick) {
      if (!getForm()) return;
      var mode = pick.getAttribute("data-gen-picker");
      if (mode === "chapter-template" || mode === "student-data") {
        loadAndShow(mode);
      }
      return;
    }
    var clr = ev.target.closest("[data-gen-clear]");
    if (clr) {
      var f = getForm();
      var refPicks = document.getElementById("gen-ref-picks");
      var wrapCh = document.getElementById("gen-ref-pick-chapter-wrap");
      var wrapSt = document.getElementById("gen-ref-pick-student-wrap");
      var inChip =
        (wrapCh && wrapCh.contains(clr)) || (wrapSt && wrapSt.contains(clr));
      if (!f || (!f.contains(clr) && (!refPicks || !refPicks.contains(clr)) && !inChip)) return;
      var which = clr.getAttribute("data-gen-clear");
      if (which === "chapter-template") setChapterTemplate("", "");
      else if (which === "student-data") setStudentData("", "");
    }
  });

  /**
   * 切换 PPT 模板后恢复「报告类型 / 学生数据」（与 readGenerateRefPickSnapshot 配对使用）。
   * @param {HTMLFormElement|null} form
   * @param {{ chapterId?: string, chapterLabel?: string, studentId?: string, studentLabel?: string }|null} snap
   */
  async function restoreGenerateRefPicks(form, snap) {
    if (!form || !snap) return;
    var wantedChapter = !!(snap.chapterId && String(snap.chapterId).trim());
    if (snap.chapterId) {
      var ok = await pickChapterTemplate(String(snap.chapterId), String(snap.chapterLabel || snap.chapterId || "").trim());
      if (!ok) setChapterTemplate("", "");
    }
    if (snap.studentId) {
      setStudentData(String(snap.studentId), String(snap.studentLabel || "").trim());
    }
    var taskInp = form.querySelector('input[name="task_id"]');
    var taskId = taskInp ? String(taskInp.value || "").trim() : "";
    var chEl = form.querySelector("#gen-chapter-template-id");
    var hasChapter = chEl && String(chEl.value || "").trim();
    if (wantedChapter && taskId && !hasChapter) {
      await applyDefaultChapterTemplate(form);
    }
  }

  window.PptApp = window.PptApp || {};
  window.PptApp.syncGenerateRefPicksUi = syncRefPicksUi;
  window.PptApp.applyDefaultChapterTemplate = applyDefaultChapterTemplate;
  window.PptApp.applyDefaultWordReportTemplate = applyDefaultWordReportTemplate;
  window.PptApp.pickChapterTemplate = pickChapterTemplate;
  window.PptApp.restoreGenerateRefPicks = restoreGenerateRefPicks;
  window.PptApp.getGenerateMode = function () {
    return getGenerateMode(getForm());
  };
  window.PptApp.setGenerateMode = function (mode) {
    setGenerateMode(getForm(), mode);
    syncRefPicksUi();
  };
  window.PptApp.clearChapterTemplatePick = function () {
    setChapterTemplate("", "");
  };
})();
