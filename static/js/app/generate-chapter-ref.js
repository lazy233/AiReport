/**
 * 文档生成：解析到 PPT（前端）
 * - 调用 POST /api/resolve-chapter-reference：首页（若有）+ 各章；章节名按模板顺序写入章块；学生字段由服务端大模型分配。
 */
(function () {
  /** @type {Array<{key: string, label: string, value: string}> | null} */
  var cachedStudentFields = null;

  var _autoResolveDebounce = null;

  function esc(s) {
    var d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  function getForm() {
    return document.getElementById("ai-generate-form");
  }

  /** @returns {string[]} 不满足生成条件时的提示列表（空数组表示可生成） */
  function collectGenerateValidationMessages(form) {
    var msgs = [];
    if (!form) return ["表单未就绪。"];
    var taskId = getTaskId(form);
    if (!taskId) {
      msgs.push("请先选择已解析的 PPT。");
    }
    if (!form.querySelector(".chapter-tabs")) {
      msgs.push("章节结构未加载，请稍后重试或刷新页面。");
    }
    var tplId = (form.querySelector("#gen-chapter-template-id") || {}).value || "";
    if (!String(tplId).trim()) {
      msgs.push("请选择章节模板。");
    }
    var sid = (form.querySelector("#gen-student-data-id") || {}).value || "";
    if (!String(sid).trim()) {
      msgs.push("请选择学生数据。");
    }
    if (!form._chapterRefResolved) {
      msgs.push(
        "请等待「解析到 PPT」完成（字段分配成功后再生成）；若失败可点「重新解析」重试。",
      );
    }
    if (!form.querySelector(".group-select:checked")) {
      msgs.push("请至少勾选一个章节（纳入本次生成）。");
    }
    return msgs;
  }

  /** 章节引用解析 API 已成功执行过（更换模板/学生数据后会重置） */
  function syncGenerateSubmitEnabled() {
    var form = getForm();
    var btn = document.getElementById("ai-generate-submit");
    if (!form || !btn) return;
    var progressWrap = document.getElementById("generate-progress");
    if (progressWrap && progressWrap.classList.contains("is-running")) {
      btn.disabled = true;
      btn.setAttribute("title", "正在生成中，请稍候…");
      btn.classList.remove("gen-submit--invalid");
      return;
    }
    btn.disabled = false;
    var msgs = collectGenerateValidationMessages(form);
    if (msgs.length) {
      btn.setAttribute("title", msgs.join(" "));
      btn.classList.add("gen-submit--invalid");
    } else {
      btn.removeAttribute("title");
      btn.classList.remove("gen-submit--invalid");
    }
  }

  function isRefSlotKind(kind) {
    return kind === "chapter" || kind === "cover";
  }

  /** 与服务端 chapter_ref slots 顺序一致：首页（若有）+ 各章 */
  function getChapterPanels(form) {
    if (!form) return [];
    var out = [];
    var cov = form.querySelector('.chapter-tab-panel[data-panel-kind="cover"]');
    if (cov) out.push(cov);
    form.querySelectorAll('.chapter-tab-panel[data-panel-kind="chapter"]').forEach(function (p) {
      out.push(p);
    });
    return out;
  }

  function getTaskId(form) {
    if (!form) return "";
    var taskInp = form.querySelector('input[name="task_id"]');
    return ((taskInp && taskInp.value) || "").trim();
  }

  function setScreenshotStatus(panel, msg, isError) {
    var el = panel.querySelector("[data-chapter-ref-screenshot-status]");
    if (!el) return;
    if (!msg) {
      el.hidden = true;
      el.textContent = "";
      el.classList.remove("is-error");
      return;
    }
    el.hidden = false;
    el.textContent = msg;
    el.classList.toggle("is-error", !!isError);
  }

  function renderScreenshotList(panel) {
    var wrap = panel.querySelector("[data-chapter-ref-screenshot-list]");
    if (!wrap) return;
    var shots = panel._screenshots || [];
    wrap.innerHTML = "";
    shots.forEach(function (s, idx) {
      var div = document.createElement("div");
      div.className = "chapter-ref-screenshot-item";
      var img = document.createElement("img");
      img.alt = "";
      img.loading = "lazy";
      img.src = s.url || "";
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "chapter-ref-screenshot-remove";
      btn.setAttribute("data-shot-index", String(idx));
      btn.setAttribute("aria-label", "删除截图");
      btn.appendChild(document.createTextNode("×"));
      var meta = document.createElement("div");
      meta.className = "chapter-ref-screenshot-item-meta";
      meta.textContent = (s.originalName || "").slice(0, 80);
      div.appendChild(img);
      div.appendChild(btn);
      div.appendChild(meta);
      wrap.appendChild(div);
    });
  }

  function uploadScreenshotFiles(panel, form, fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return;
    var taskId = getTaskId(form);
    if (!taskId) {
      setScreenshotStatus(panel, "缺少 PPT 任务 ID，请重新选择模板。", true);
      return;
    }
    setScreenshotStatus(panel, "上传中…", false);
    (async function () {
      try {
        for (var i = 0; i < files.length; i++) {
          var fd = new FormData();
          fd.append("task_id", taskId);
          fd.append("file", files[i]);
          var res = await fetch("/api/chapter-ref-screenshot", { method: "POST", body: fd });
          var j = await res.json().catch(function () {
            return {};
          });
          if (!res.ok || !j.ok) throw new Error(j.error || "上传失败");
          panel._screenshots = panel._screenshots || [];
          panel._screenshots.push({
            url: j.item.url,
            storedFilename: j.item.storedFilename,
            originalName: j.item.originalName || files[i].name || "",
          });
          renderScreenshotList(panel);
          updateHiddenJson(form);
        }
        setScreenshotStatus(panel, "", false);
      } catch (e) {
        setScreenshotStatus(panel, e.message || String(e), true);
      }
    })();
  }

  function bindScreenshotPanel(form, panel) {
    if (panel._chRefShotUiBound) return;
    var uploadBtn = panel.querySelector("[data-chapter-ref-upload-btn]");
    var fileInput = panel.querySelector("[data-chapter-ref-file-input]");
    if (!uploadBtn || !fileInput) return;
    panel._chRefShotUiBound = true;
    uploadBtn.addEventListener("click", function () {
      fileInput.click();
    });
    fileInput.addEventListener("change", function () {
      uploadScreenshotFiles(panel, form, fileInput.files);
      fileInput.value = "";
    });
  }

  var _resolveHintIdleText = "";

  function ensureResolveHintIdleCaptured() {
    var el = document.getElementById("gen-resolve-hint");
    if (el && !_resolveHintIdleText) {
      _resolveHintIdleText = el.textContent.trim();
    }
  }

  function setResolveLoading(isLoading) {
    ensureResolveHintIdleCaptured();
    var btn = document.getElementById("gen-resolve-to-ppt");
    var spin = document.getElementById("gen-resolve-spinner");
    var hint = document.getElementById("gen-resolve-hint");
    var row = document.getElementById("gen-resolve-actions");
    if (!btn) return;
    if (isLoading) {
      if (!btn.getAttribute("data-label-idle")) {
        btn.setAttribute("data-label-idle", btn.textContent.trim());
      }
      btn.textContent = "解析中…";
      btn.disabled = true;
      btn.setAttribute("aria-busy", "true");
      if (spin) spin.removeAttribute("hidden");
      if (hint) {
        hint.textContent =
          "正在请求服务端，大模型正在为首页与各章分配学生字段，请稍候（通常数十秒内完成）。界面未卡死，可继续滚动浏览。";
      }
      if (row) row.classList.add("is-busy");
    } else {
      btn.textContent = btn.getAttribute("data-label-idle") || "重新解析";
      btn.disabled = false;
      btn.removeAttribute("aria-busy");
      if (spin) spin.setAttribute("hidden", "");
      if (hint && _resolveHintIdleText) hint.textContent = _resolveHintIdleText;
      if (row) row.classList.remove("is-busy");
    }
  }

  function yieldToPaint() {
    return new Promise(function (resolve) {
      if (typeof requestAnimationFrame === "function") {
        requestAnimationFrame(function () {
          setTimeout(resolve, 0);
        });
      } else {
        setTimeout(resolve, 0);
      }
    });
  }

  function scheduleAutoResolve() {
    if (_autoResolveDebounce) clearTimeout(_autoResolveDebounce);
    _autoResolveDebounce = setTimeout(function () {
      _autoResolveDebounce = null;
      var form = getForm();
      if (!form) return;
      var tid = (form.querySelector("#gen-chapter-template-id") || {}).value || "";
      var sid = (form.querySelector("#gen-student-data-id") || {}).value || "";
      if (!tid || !sid) return;
      var btn = document.getElementById("gen-resolve-to-ppt");
      if (btn && btn.getAttribute("aria-busy") === "true") return;
      onResolveClick();
    }, 450);
  }

  function syncResolveRowVisibility() {
    var form = getForm();
    var row = document.getElementById("gen-resolve-actions");
    if (!form || !row) return;
    var tid = (form.querySelector("#gen-chapter-template-id") || {}).value || "";
    var sid = (form.querySelector("#gen-student-data-id") || {}).value || "";
    row.hidden = !(tid && sid);
    if (tid && sid) {
      scheduleAutoResolve();
    } else if (_autoResolveDebounce) {
      clearTimeout(_autoResolveDebounce);
      _autoResolveDebounce = null;
    }
  }

  function measureTagsOneRowHeight(wrap) {
    var ch = wrap.children;
    if (!ch.length) return 0;
    var t0 = ch[0].offsetTop;
    var h = 0;
    var i;
    for (i = 0; i < ch.length; i++) {
      if (ch[i].offsetTop > t0 + 3) break;
      h = Math.max(h, ch[i].offsetHeight);
    }
    return h + 8;
  }

  function syncTagsCollapsedState(panel) {
    var wrap = panel.querySelector("[data-chapter-ref-tags]");
    var shell = panel.querySelector("[data-chapter-ref-tags-shell]");
    var btn = panel.querySelector("[data-chapter-ref-tags-toggle]");
    if (!wrap || !shell || !btn) return;
    wrap.style.maxHeight = "";
    wrap.style.overflow = "";
    shell.classList.remove("has-more");
    btn.hidden = true;
    if (!wrap.children.length) {
      panel._tagsExpanded = false;
      btn.setAttribute("aria-expanded", "false");
      return;
    }
    requestAnimationFrame(function () {
      requestAnimationFrame(function () {
        var hRow = measureTagsOneRowHeight(wrap);
        var full = wrap.scrollHeight;
        if (hRow <= 0 || full <= hRow + 2) {
          panel._tagsExpanded = false;
          btn.hidden = true;
          btn.setAttribute("aria-expanded", "false");
          return;
        }
        shell.classList.add("has-more");
        btn.hidden = false;
        var expanded = !!panel._tagsExpanded;
        btn.setAttribute("aria-expanded", expanded ? "true" : "false");
        btn.textContent = expanded
          ? "收起标签"
          : "展开全部标签（" + wrap.children.length + "）";
        if (!expanded) {
          shell.classList.remove("chapter-ref-tags-shell--open");
          wrap.style.maxHeight = hRow + "px";
          wrap.style.overflow = "hidden";
        } else {
          shell.classList.add("chapter-ref-tags-shell--open");
          wrap.style.maxHeight = "none";
          wrap.style.overflow = "visible";
        }
      });
    });
  }

  function renderTags(panel) {
    var wrap = panel.querySelector("[data-chapter-ref-tags]");
    if (!wrap) return;
    var list = panel._attachedFields || [];
    wrap.innerHTML = "";
    list.forEach(function (f, idx) {
      var tag = document.createElement("span");
      tag.className = "chapter-ref-tag";
      tag.innerHTML =
        '<span class="chapter-ref-tag-main">' +
        "<strong>" +
        esc(f.label) +
        "</strong>" +
        '<span class="chapter-ref-tag-value muted">' +
        esc(f.value.length > 48 ? f.value.slice(0, 48) + "…" : f.value) +
        "</span></span>" +
        '<button type="button" class="chapter-ref-tag-remove" data-tag-index="' +
        idx +
        '" aria-label="移除此字段">×</button>';
      wrap.appendChild(tag);
    });
    syncTagsCollapsedState(panel);
  }

  function fillFieldSelect(panel, form) {
    var sel = panel.querySelector("[data-chapter-ref-select]");
    if (!sel) return;
    var all = form.__studentFieldsAll || [];
    var attached = panel._attachedFields || [];
    var used = new Set(attached.map(function (x) {
      return x.key;
    }));
    sel.innerHTML = '<option value="">选择学生字段并添加…</option>';
    all.forEach(function (f) {
      if (used.has(f.key)) return;
      var opt = document.createElement("option");
      opt.value = f.key;
      opt.textContent = f.label + " · " + (f.value.length > 36 ? f.value.slice(0, 36) + "…" : f.value);
      sel.appendChild(opt);
    });
  }

  function refreshAllSelects(form) {
    getChapterPanels(form).forEach(function (p) {
      fillFieldSelect(p, form);
    });
  }

  function updateHiddenJson(form) {
    var el = document.getElementById("gen-chapter-ref-json");
    if (!el) return;
    var panels = getChapterPanels(form);
    var payload = { version: 2, slots: [] };
    panels.forEach(function (panel, idx) {
      var cb = panel.querySelector(".group-select");
      var slides = cb ? cb.getAttribute("data-slides") || "" : "";
      var titleInp = panel.querySelector("[data-chapter-ref-template-title]");
      var title = titleInp ? titleInp.value.trim() : "";
      var fields = panel._attachedFields || [];
      var shots = panel._screenshots || [];
      payload.slots.push({
        slotIndex: idx,
        slides: slides,
        templateTitle: title,
        fields: fields.map(function (f) {
          return { key: f.key, label: f.label, value: f.value };
        }),
        screenshots: shots.map(function (s) {
          return {
            url: s.url,
            storedFilename: s.storedFilename,
            originalName: s.originalName || "",
          };
        }),
      });
    });
    el.value = JSON.stringify(payload);
  }

  function resetChapterReferenceUi(form) {
    if (!form) return;
    form._chapterRefResolved = false;
    setResolveLoading(false);
    cachedStudentFields = null;
    form.__studentFieldsAll = null;
    var el = document.getElementById("gen-chapter-ref-json");
    if (el) el.value = "";
    getChapterPanels(form).forEach(function (panel) {
      panel._attachedFields = [];
      var row = panel.querySelector(".chapter-ref-template-row");
      var titleInp = panel.querySelector("[data-chapter-ref-template-title]");
      if (titleInp) titleInp.value = "";
      var kindReset = panel.getAttribute("data-panel-kind") || "";
      if (row) row.hidden = !isRefSlotKind(kindReset);
      panel._tagsExpanded = false;
      renderTags(panel);
      var sel = panel.querySelector("[data-chapter-ref-select]");
      if (sel) sel.innerHTML = '<option value="">选择学生字段并添加…</option>';
      panel._screenshots = [];
      renderScreenshotList(panel);
      setScreenshotStatus(panel, "", false);
    });
    syncResolveRowVisibility();
    syncGenerateSubmitEnabled();
  }

  function initPanelLayout(form) {
    form.querySelectorAll(".chapter-tab-panel").forEach(function (panel) {
      var kind = panel.getAttribute("data-panel-kind") || "";
      var note = panel.querySelector("[data-non-chapter-note]");
      var body = panel.querySelector("[data-chapter-ref-body]");
      var trow = panel.querySelector(".chapter-ref-template-row");
      if (isRefSlotKind(kind)) {
        if (note) note.hidden = true;
        if (body) body.hidden = false;
        if (trow) trow.hidden = false;
      } else {
        if (note) note.hidden = false;
        if (body) body.hidden = true;
        if (trow) trow.hidden = true;
      }
    });
  }

  async function onResolveClick() {
    var form = getForm();
    var btn = document.getElementById("gen-resolve-to-ppt");
    if (!form || !btn) return;
    var tid = (form.querySelector("#gen-chapter-template-id") || {}).value || "";
    var sid = (form.querySelector("#gen-student-data-id") || {}).value || "";
    var taskInp = form.querySelector('input[name="task_id"]');
    var taskId = (taskInp && taskInp.value) || "";
    if (!tid || !sid) return;
    if (!taskId.trim()) {
      return;
    }
    var panels = getChapterPanels(form);
    var hasChapter = !!form.querySelector('.chapter-tab-panel[data-panel-kind="chapter"]');
    if (!hasChapter) {
      window.alert("当前 PPT 未识别到「章」块，无法对齐章节模板。");
      return;
    }
    if (btn.getAttribute("aria-busy") === "true") return;
    setResolveLoading(true);
    await yieldToPaint();
    try {
      var res = await fetch("/api/resolve-chapter-reference", {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify({
          task_id: taskId.trim(),
          chapter_template_id: tid,
          student_data_id: sid,
        }),
      });
      var j = await res.json();
      if (!res.ok || !j.ok) throw new Error(j.error || "解析失败");
      var data = j.data || {};
      var slots = Array.isArray(data.slots) ? data.slots : [];
      var allFields = Array.isArray(data.allStudentFields) ? data.allStudentFields : [];
      cachedStudentFields = allFields;
      form.__studentFieldsAll = allFields;
      panels.forEach(function (panel, idx) {
        var slot = slots[idx] || {};
        var title = String(slot.templateTitle || "").trim();
        var row = panel.querySelector(".chapter-ref-template-row");
        var titleInp = panel.querySelector("[data-chapter-ref-template-title]");
        var pkind = (panel.getAttribute("data-panel-kind") || "").trim();
        if (titleInp) titleInp.value = title;
        if (row) row.hidden = !isRefSlotKind(pkind);
        var flist = Array.isArray(slot.fields) ? slot.fields : [];
        panel._attachedFields = flist.map(function (f) {
          return {
            key: f.key,
            label: f.label || f.key,
            value: f.value != null ? String(f.value) : "",
          };
        });
        panel._tagsExpanded = false;
        renderTags(panel);
        var remoteShots = Array.isArray(slot.screenshots) ? slot.screenshots : [];
        if (remoteShots.length) {
          panel._screenshots = remoteShots
            .map(function (s) {
              return {
                url: s.url || "",
                storedFilename: s.storedFilename || s.stored_filename || "",
                originalName: s.originalName || s.original_name || "",
              };
            })
            .filter(function (s) {
              return s.url && s.storedFilename;
            });
          renderScreenshotList(panel);
        }
      });
      refreshAllSelects(form);
      updateHiddenJson(form);
      form._chapterRefResolved = true;
    } catch (e) {
      form._chapterRefResolved = false;
      window.alert(e.message || String(e));
    } finally {
      setResolveLoading(false);
      syncGenerateSubmitEnabled();
    }
  }

  function bindForm(form) {
    if (!form || form.id !== "ai-generate-form") return;
    initPanelLayout(form);
    getChapterPanels(form).forEach(function (p) {
      if (!Array.isArray(p._attachedFields)) p._attachedFields = [];
      if (!Array.isArray(p._screenshots)) p._screenshots = [];
      bindScreenshotPanel(form, p);
    });
    if (!form._chRefTitleInputBound) {
      form._chRefTitleInputBound = true;
      form.addEventListener("input", function (ev) {
        if (ev.target && ev.target.matches && ev.target.matches("[data-chapter-ref-template-title]")) {
          updateHiddenJson(form);
        }
      });
    }
    var btn = document.getElementById("gen-resolve-to-ppt");
    if (btn && !btn.getAttribute("data-ch-ref-bound")) {
      btn.setAttribute("data-ch-ref-bound", "1");
      btn.addEventListener("click", onResolveClick);
    }
    getChapterPanels(form).forEach(function (panel) {
      if (panel._chRefAddBound) return;
      panel._chRefAddBound = true;
      var addBtn = panel.querySelector(".chapter-ref-add-btn");
      var sel = panel.querySelector("[data-chapter-ref-select]");
      if (!addBtn || !sel) return;
      addBtn.addEventListener("click", function () {
        var key = (sel.value || "").trim();
        if (!key) return;
        var all = form.__studentFieldsAll || [];
        var found = all.find(function (x) {
          return x.key === key;
        });
        if (!found) return;
        panel._attachedFields = panel._attachedFields || [];
        panel._attachedFields.push({
          key: found.key,
          label: found.label,
          value: found.value,
        });
        sel.value = "";
        renderTags(panel);
        refreshAllSelects(form);
        updateHiddenJson(form);
      });
    });
    if (!form._chRefTagDelegate) {
      form._chRefTagDelegate = true;
      form.addEventListener("click", function (ev) {
        var rmShot = ev.target.closest(".chapter-ref-screenshot-remove");
        if (rmShot && form.contains(rmShot)) {
          var spanel = rmShot.closest(".chapter-tab-panel");
          if (!spanel || !isRefSlotKind(spanel.getAttribute("data-panel-kind") || "")) return;
          var six = parseInt(rmShot.getAttribute("data-shot-index") || "-1", 10);
          if (!Number.isFinite(six) || six < 0) return;
          var slist = spanel._screenshots || [];
          var shot = slist[six];
          var tid = getTaskId(form);
          if (shot && shot.storedFilename && tid) {
            fetch(
              "/api/chapter-ref-screenshot/" +
                encodeURIComponent(tid) +
                "/" +
                encodeURIComponent(shot.storedFilename),
              { method: "DELETE" },
            ).catch(function () {});
          }
          slist.splice(six, 1);
          spanel._screenshots = slist;
          renderScreenshotList(spanel);
          setScreenshotStatus(spanel, "", false);
          updateHiddenJson(form);
          return;
        }
        var tagToggle = ev.target.closest("[data-chapter-ref-tags-toggle]");
        if (tagToggle && form.contains(tagToggle)) {
          var tpan = tagToggle.closest(".chapter-tab-panel");
          if (!tpan || !isRefSlotKind(tpan.getAttribute("data-panel-kind") || "")) return;
          tpan._tagsExpanded = !tpan._tagsExpanded;
          syncTagsCollapsedState(tpan);
          return;
        }
        var rm = ev.target.closest(".chapter-ref-tag-remove");
        if (!rm || !form.contains(rm)) return;
        var panel = rm.closest(".chapter-tab-panel");
        if (!panel || !isRefSlotKind(panel.getAttribute("data-panel-kind") || "")) return;
        var idx = parseInt(rm.getAttribute("data-tag-index") || "-1", 10);
        if (!Number.isFinite(idx) || idx < 0) return;
        panel._attachedFields = panel._attachedFields || [];
        panel._attachedFields.splice(idx, 1);
        renderTags(panel);
        refreshAllSelects(form);
        updateHiddenJson(form);
      });
    }
    syncResolveRowVisibility();
    syncGenerateSubmitEnabled();
  }

  window.PptApp = window.PptApp || {};
  window.PptApp.initChapterReferenceUi = bindForm;
  window.PptApp.syncGenResolveRow = syncResolveRowVisibility;
  window.PptApp.resetChapterReferenceUi = resetChapterReferenceUi;
  window.PptApp.triggerResolveChapterReference = onResolveClick;
  window.PptApp.syncGenerateSubmitEnabled = syncGenerateSubmitEnabled;
  window.PptApp.collectGenerateValidationMessages = collectGenerateValidationMessages;
})();
