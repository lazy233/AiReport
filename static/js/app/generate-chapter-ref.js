/**
 * 文档生成：解析到 PPT（前端）
 * - 调用 POST /api/resolve-chapter-reference：首页（若有）+ 各章；章节名按模板顺序写入章块；学生字段由服务端大模型分配。
 */
(function () {
  /** @type {Array<{key: string, label: string, value: string}> | null} */
  var cachedStudentFields = null;

  var _autoResolveDebounce = null;

  /** @type {HTMLElement | null} */
  var _fieldPickBackdrop = null;
  /** @type {HTMLFormElement | null} */
  var _fieldPickForm = null;
  /** @type {HTMLElement | null} */
  var _fieldPickPanel = null;

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
      msgs.push("请选择报告类型。");
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
      var taskId = getTaskId(form);
      if (!taskId) return;
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
    var taskId = getTaskId(form);
    var tid = (form.querySelector("#gen-chapter-template-id") || {}).value || "";
    var sid = (form.querySelector("#gen-student-data-id") || {}).value || "";
    row.hidden = !(taskId && tid && sid);
    if (taskId && tid && sid) {
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

  /** 首页「模板章节名」默认填章节模板名称（解析前后均可再改） */
  function applyDefaultCoverTitleFromTemplateName(form, templateName) {
    var name = String(templateName || "").trim();
    if (!form || !name) return;
    var cov = form.querySelector('.chapter-tab-panel[data-panel-kind="cover"]');
    if (!cov) return;
    var inp = cov.querySelector("[data-chapter-ref-template-title]");
    if (inp) inp.value = name;
    updateHiddenJson(form);
  }

  function closeChapterRefFieldModal() {
    if (!_fieldPickBackdrop) return;
    _fieldPickBackdrop.setAttribute("hidden", "");
    _fieldPickForm = null;
    _fieldPickPanel = null;
    var msg = _fieldPickBackdrop.querySelector("#ch-ref-field-custom-msg");
    if (msg) {
      msg.hidden = true;
      msg.textContent = "";
      msg.classList.remove("is-error");
    }
  }

  function getAvailableFieldsForPanel(form, panel) {
    var all = form.__studentFieldsAll || [];
    var used = new Set((panel._attachedFields || []).map(function (x) {
      return x.key;
    }));
    return all.filter(function (f) {
      return f && f.key && !used.has(f.key);
    });
  }

  function renderChapterRefFieldModalList() {
    if (!_fieldPickBackdrop || !_fieldPickForm || !_fieldPickPanel) return;
    var listEl = _fieldPickBackdrop.querySelector("#ch-ref-field-modal-list");
    var emptyEl = _fieldPickBackdrop.querySelector("#ch-ref-field-modal-empty");
    var confirmBtn = _fieldPickBackdrop.querySelector("#ch-ref-field-modal-confirm");
    var searchInp = _fieldPickBackdrop.querySelector("#ch-ref-field-modal-search");
    var q = (searchInp && searchInp.value.trim().toLowerCase()) || "";
    if (!listEl || !emptyEl) return;
    var available = getAvailableFieldsForPanel(_fieldPickForm, _fieldPickPanel);
    listEl.innerHTML = "";
    var shown = 0;
    available.forEach(function (f) {
      var blob = (f.label || "") + " " + (f.key || "") + " " + (f.value || "");
      if (q && blob.toLowerCase().indexOf(q) === -1) return;
      shown += 1;
      var li = document.createElement("li");
      li.className = "ch-ref-field-modal-item";
      var id = "ch-ref-modal-f-" + shown + "-" + String(f.key).replace(/[^a-zA-Z0-9_-]/g, "_");
      var label = document.createElement("label");
      label.className = "ch-ref-field-modal-label";
      label.setAttribute("for", id);
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.id = id;
      cb.value = f.key;
      cb.className = "ch-ref-field-modal-cb";
      var main = document.createElement("span");
      main.className = "ch-ref-field-modal-main";
      var strong = document.createElement("strong");
      strong.textContent = f.label || f.key;
      var keySpan = document.createElement("span");
      keySpan.className = "muted ch-ref-field-modal-key";
      keySpan.textContent = f.key;
      main.appendChild(strong);
      main.appendChild(keySpan);
      var sub = document.createElement("span");
      sub.className = "muted ch-ref-field-modal-value";
      sub.textContent = f.value.length > 80 ? f.value.slice(0, 77) + "…" : f.value;
      var col = document.createElement("div");
      col.className = "ch-ref-field-modal-textcol";
      col.appendChild(main);
      col.appendChild(sub);
      label.appendChild(cb);
      label.appendChild(col);
      li.appendChild(label);
      listEl.appendChild(li);
    });
    emptyEl.hidden = shown > 0;
    if (confirmBtn) {
      confirmBtn.disabled = !listEl.querySelector(".ch-ref-field-modal-cb:checked");
    }
  }

  function openChapterRefFieldModal(form, panel) {
    if (!form || !panel) return;
    ensureChapterRefFieldModal();
    _fieldPickForm = form;
    _fieldPickPanel = panel;
    var searchInp = _fieldPickBackdrop.querySelector("#ch-ref-field-modal-search");
    var nameInp = _fieldPickBackdrop.querySelector("#ch-ref-field-custom-name");
    var valInp = _fieldPickBackdrop.querySelector("#ch-ref-field-custom-value");
    var msg = _fieldPickBackdrop.querySelector("#ch-ref-field-custom-msg");
    if (searchInp) searchInp.value = "";
    if (nameInp) nameInp.value = "";
    if (valInp) valInp.value = "";
    if (msg) {
      msg.hidden = true;
      msg.textContent = "";
      msg.classList.remove("is-error");
    }
    renderChapterRefFieldModalList();
    _fieldPickBackdrop.removeAttribute("hidden");
    if (nameInp) {
      nameInp.focus();
    } else if (searchInp) {
      searchInp.focus();
    }
  }

  function confirmChapterRefFieldModal() {
    if (!_fieldPickForm || !_fieldPickPanel || !_fieldPickBackdrop) return;
    var listEl = _fieldPickBackdrop.querySelector("#ch-ref-field-modal-list");
    if (!listEl) return;
    var all = _fieldPickForm.__studentFieldsAll || [];
    var keys = Array.prototype.map.call(listEl.querySelectorAll(".ch-ref-field-modal-cb:checked"), function (c) {
      return c.value;
    });
    if (!keys.length) return;
    var panel = _fieldPickPanel;
    panel._attachedFields = panel._attachedFields || [];
    var existing = new Set(panel._attachedFields.map(function (x) {
      return x.key;
    }));
    keys.forEach(function (key) {
      if (existing.has(key)) return;
      var found = all.find(function (x) {
        return x.key === key;
      });
      if (!found) return;
      panel._attachedFields.push({
        key: found.key,
        label: found.label,
        value: found.value,
      });
      existing.add(key);
    });
    renderTags(panel);
    updateHiddenJson(_fieldPickForm);
    closeChapterRefFieldModal();
  }

  function makeCustomFieldKey(panel, seed) {
    var base = String(seed || "").trim().toLowerCase().replace(/[^a-z0-9_-]+/g, "_").replace(/^_+|_+$/g, "");
    if (!base) base = "custom";
    var key = "custom_" + base;
    var used = new Set((panel._attachedFields || []).map(function (x) { return x && x.key; }));
    if (!used.has(key)) return key;
    var i = 2;
    while (used.has(key + "_" + i)) i += 1;
    return key + "_" + i;
  }

  function setCustomFieldMsg(msg, isError) {
    if (!_fieldPickBackdrop) return;
    var el = _fieldPickBackdrop.querySelector("#ch-ref-field-custom-msg");
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

  function addCustomFieldFromModal() {
    if (!_fieldPickForm || !_fieldPickPanel || !_fieldPickBackdrop) return;
    var nameInp = _fieldPickBackdrop.querySelector("#ch-ref-field-custom-name");
    var valInp = _fieldPickBackdrop.querySelector("#ch-ref-field-custom-value");
    if (!nameInp || !valInp) return;
    var label = String(nameInp.value || "").trim();
    var value = String(valInp.value || "").trim();
    if (!label) {
      setCustomFieldMsg("请填写自定义字段名。", true);
      nameInp.focus();
      return;
    }
    if (!value) {
      setCustomFieldMsg("请填写自定义字段值。", true);
      valInp.focus();
      return;
    }
    var panel = _fieldPickPanel;
    panel._attachedFields = panel._attachedFields || [];
    panel._attachedFields.push({
      key: makeCustomFieldKey(panel, label),
      label: label,
      value: value,
    });
    renderTags(panel);
    updateHiddenJson(_fieldPickForm);
    setCustomFieldMsg("", false);
    closeChapterRefFieldModal();
  }

  function ensureChapterRefFieldModal() {
    if (_fieldPickBackdrop) return;
    _fieldPickBackdrop = document.createElement("div");
    _fieldPickBackdrop.className = "ch-ref-field-modal-backdrop";
    _fieldPickBackdrop.setAttribute("hidden", "");
    _fieldPickBackdrop.innerHTML =
      '<div class="ch-ref-field-modal" role="dialog" aria-modal="true" aria-labelledby="ch-ref-field-modal-title">' +
      '<div class="ch-ref-field-modal-head">' +
      '<h3 id="ch-ref-field-modal-title" class="ch-ref-field-modal-title">添加学生字段</h3>' +
      '<button type="button" class="button secondary ch-ref-field-modal-close" aria-label="关闭">关闭</button>' +
      "</div>" +
      '<div class="ch-ref-field-modal-body">' +
      '<div class="ch-ref-field-custom-box">' +
      '<p class="muted ch-ref-field-custom-kicker">自定义字段（仅本次生成有效）</p>' +
      '<div class="ch-ref-field-custom-row">' +
      '<input type="text" class="ch-ref-field-custom-input" id="ch-ref-field-custom-name" placeholder="字段名，例如：家庭住址" autocomplete="off" />' +
      '<input type="text" class="ch-ref-field-custom-input" id="ch-ref-field-custom-value" placeholder="字段值，例如：杭州市西湖区..." autocomplete="off" />' +
      '<button type="button" class="button secondary ch-ref-field-custom-add" id="ch-ref-field-custom-add">添加自定义字段</button>' +
      "</div>" +
      '<p class="muted ch-ref-field-custom-msg" id="ch-ref-field-custom-msg" hidden></p>' +
      "</div>" +
      '<input type="search" class="ch-ref-field-modal-search" id="ch-ref-field-modal-search" placeholder="搜索字段名或内容…" autocomplete="off" />' +
      '<div class="ch-ref-field-modal-list-wrap">' +
      '<p class="muted ch-ref-field-modal-empty" id="ch-ref-field-modal-empty" hidden>没有匹配的字段</p>' +
      '<ul class="ch-ref-field-modal-list" id="ch-ref-field-modal-list"></ul>' +
      "</div>" +
      '<div class="ch-ref-field-modal-foot">' +
      '<button type="button" class="button secondary ch-ref-field-modal-cancel">取消</button>' +
      '<button type="button" class="button ch-ref-field-modal-confirm" id="ch-ref-field-modal-confirm" disabled>添加所选</button>' +
      "</div>" +
      "</div>" +
      "</div>";
    document.body.appendChild(_fieldPickBackdrop);

    _fieldPickBackdrop.addEventListener("click", function (ev) {
      if (ev.target === _fieldPickBackdrop) closeChapterRefFieldModal();
    });
    _fieldPickBackdrop.querySelector(".ch-ref-field-modal-close").addEventListener("click", closeChapterRefFieldModal);
    _fieldPickBackdrop.querySelector(".ch-ref-field-modal-cancel").addEventListener("click", closeChapterRefFieldModal);
    _fieldPickBackdrop.querySelector(".ch-ref-field-modal-confirm").addEventListener("click", confirmChapterRefFieldModal);
    var search = _fieldPickBackdrop.querySelector("#ch-ref-field-modal-search");
    if (search) {
      search.addEventListener("input", function () {
        renderChapterRefFieldModalList();
      });
    }
    var listRoot = _fieldPickBackdrop.querySelector("#ch-ref-field-modal-list");
    if (listRoot) {
      listRoot.addEventListener("change", function () {
        var cbtn = _fieldPickBackdrop.querySelector("#ch-ref-field-modal-confirm");
        if (cbtn) cbtn.disabled = !listRoot.querySelector(".ch-ref-field-modal-cb:checked");
      });
    }
    var customAdd = _fieldPickBackdrop.querySelector("#ch-ref-field-custom-add");
    if (customAdd) customAdd.addEventListener("click", addCustomFieldFromModal);
    var customVal = _fieldPickBackdrop.querySelector("#ch-ref-field-custom-value");
    if (customVal) {
      customVal.addEventListener("keydown", function (ev) {
        if (ev.key === "Enter") {
          ev.preventDefault();
          addCustomFieldFromModal();
        }
      });
    }
    document.addEventListener("keydown", function (ev) {
      if (ev.key !== "Escape" || !_fieldPickBackdrop || _fieldPickBackdrop.hidden) return;
      closeChapterRefFieldModal();
    });
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
      window.alert("当前 PPT 未识别到「章」块，无法对齐报告类型。");
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
    if (!form._chRefFieldPickDelegate) {
      form._chRefFieldPickDelegate = true;
      form.addEventListener("click", function (ev) {
        var btn = ev.target.closest("[data-chapter-ref-add-fields]");
        if (!btn || !form.contains(btn)) return;
        var pan = btn.closest(".chapter-tab-panel");
        if (!pan || !isRefSlotKind(pan.getAttribute("data-panel-kind") || "")) return;
        openChapterRefFieldModal(form, pan);
      });
    }
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
  window.PptApp.applyDefaultCoverTitleFromTemplateName = applyDefaultCoverTitleFromTemplateName;
})();
