/**
 * 文档生成：解析到 PPT（前端）
 * - 调用 POST /api/resolve-chapter-reference：章节名按模板顺序；学生字段由服务端大模型分配。
 */
(function () {
  /** @type {Array<{key: string, label: string, value: string}> | null} */
  var cachedStudentFields = null;

  function esc(s) {
    var d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  function getForm() {
    return document.getElementById("ai-generate-form");
  }

  function getChapterPanels(form) {
    if (!form) return [];
    return Array.prototype.slice.call(
      form.querySelectorAll('.chapter-tab-panel[data-panel-kind="chapter"]'),
    );
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
          "正在请求服务端，大模型正在为各章分配学生字段，请稍候（通常数十秒内完成）。界面未卡死，可继续滚动浏览。";
      }
      if (row) row.classList.add("is-busy");
    } else {
      btn.textContent = btn.getAttribute("data-label-idle") || "解析到 PPT";
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

  function syncResolveRowVisibility() {
    var form = getForm();
    var row = document.getElementById("gen-resolve-actions");
    if (!form || !row) return;
    var tid = (form.querySelector("#gen-chapter-template-id") || {}).value || "";
    var sid = (form.querySelector("#gen-student-data-id") || {}).value || "";
    row.hidden = !(tid && sid);
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
      var titleEl = panel.querySelector(".chapter-ref-template-title");
      var title = titleEl ? titleEl.textContent.trim() : "";
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
    setResolveLoading(false);
    cachedStudentFields = null;
    form.__studentFieldsAll = null;
    var el = document.getElementById("gen-chapter-ref-json");
    if (el) el.value = "";
    form.querySelectorAll('.chapter-tab-panel[data-panel-kind="chapter"]').forEach(function (panel) {
      panel._attachedFields = [];
      var row = panel.querySelector(".chapter-ref-template-row");
      var titleEl = panel.querySelector(".chapter-ref-template-title");
      if (titleEl) titleEl.textContent = "";
      if (row) row.hidden = true;
      panel._tagsExpanded = false;
      renderTags(panel);
      var sel = panel.querySelector("[data-chapter-ref-select]");
      if (sel) sel.innerHTML = '<option value="">选择学生字段并添加…</option>';
      panel._screenshots = [];
      renderScreenshotList(panel);
      setScreenshotStatus(panel, "", false);
    });
    syncResolveRowVisibility();
  }

  function initPanelLayout(form) {
    form.querySelectorAll(".chapter-tab-panel").forEach(function (panel) {
      var kind = panel.getAttribute("data-panel-kind") || "";
      var note = panel.querySelector("[data-non-chapter-note]");
      var body = panel.querySelector("[data-chapter-ref-body]");
      if (kind === "chapter") {
        if (note) note.hidden = true;
        if (body) body.hidden = false;
      } else {
        if (note) note.hidden = false;
        if (body) body.hidden = true;
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
      window.alert("缺少当前 PPT 任务 ID，请重新选择模板。");
      return;
    }
    var panels = getChapterPanels(form);
    if (!panels.length) {
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
        var titleEl = panel.querySelector(".chapter-ref-template-title");
        if (titleEl) titleEl.textContent = title;
        if (row) row.hidden = !title;
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
    } catch (e) {
      window.alert(e.message || String(e));
    } finally {
      setResolveLoading(false);
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
    var btn = document.getElementById("gen-resolve-to-ppt");
    if (btn && !btn.getAttribute("data-ch-ref-bound")) {
      btn.setAttribute("data-ch-ref-bound", "1");
      btn.addEventListener("click", onResolveClick);
    }
    form.querySelectorAll('.chapter-tab-panel[data-panel-kind="chapter"]').forEach(function (panel) {
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
          var spanel = rmShot.closest('.chapter-tab-panel[data-panel-kind="chapter"]');
          if (!spanel) return;
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
          var tpan = tagToggle.closest('.chapter-tab-panel[data-panel-kind="chapter"]');
          if (!tpan) return;
          tpan._tagsExpanded = !tpan._tagsExpanded;
          syncTagsCollapsedState(tpan);
          return;
        }
        var rm = ev.target.closest(".chapter-ref-tag-remove");
        if (!rm || !form.contains(rm)) return;
        var panel = rm.closest('.chapter-tab-panel[data-panel-kind="chapter"]');
        if (!panel) return;
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
  }

  window.PptApp = window.PptApp || {};
  window.PptApp.initChapterReferenceUi = bindForm;
  window.PptApp.syncGenResolveRow = syncResolveRowVisibility;
  window.PptApp.resetChapterReferenceUi = resetChapterReferenceUi;
})();
