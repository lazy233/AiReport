/**
 * 章节勾选、异步生成、结果渲染、复制（表单可动态挂载，按需 bindGenerateForm）
 */
(function () {
  var sleep =
    window.PptApp && PptApp.sleep
      ? PptApp.sleep
      : function (ms) {
          return new Promise(function (r) {
            setTimeout(r, ms);
          });
        };

  function renderGenerationResult(data, taskId, historyId) {
    var mount = document.getElementById("ai-result-mount");
    if (!mount || !data) return;
    if (data.output_kind === "docx") {
      mount.innerHTML = "";
      var wordCard = document.createElement("div");
      wordCard.className = "card result-card";
      wordCard.id = "ai-result";
      var head = document.createElement("div");
      head.className = "section-head";
      head.innerHTML =
        '<div><h2>Word 生成结果</h2><p class="muted">已按学生数据回填 Word 表格内容。</p></div>';
      wordCard.appendChild(head);
      var sum = data.word_fill_summary || {};
      var stats = document.createElement("p");
      stats.className = "muted";
      stats.textContent =
        "表格数：" +
        (sum.table_count || 0) +
        "，变更单元格：" +
        (sum.touched_cells || 0) +
        "，占位符命中：" +
        (sum.placeholder_hits || 0);
      wordCard.appendChild(stats);
      if (historyId) {
        var dl = document.createElement("p");
        dl.style.marginTop = "10px";
        dl.innerHTML =
          '<a class="button" href="/api/generation_history/' +
          encodeURIComponent(historyId) +
          '/download">下载回填后的 Word</a>';
        wordCard.appendChild(dl);
      }
      mount.appendChild(wordCard);
      return;
    }
    if (!data.slides) return;
    mount.innerHTML = "";
    var card = document.createElement("div");
    card.className = "card result-card";
    card.id = "ai-result";
    var head = document.createElement("div");
    head.className = "section-head";
    head.innerHTML =
      '<div><h2>报告生成结果</h2><p class="muted">各页可填内容默认折叠；展开后可逐条查看、复制，或直接下载回填后的 PPT。</p></div>';
    card.appendChild(head);
    if (taskId) {
      var row = document.createElement("div");
      row.className = "download-row";
      row.innerHTML =
        '<a class="button" href="/export/' +
        encodeURIComponent(taskId) +
        '">下载填充后的 PPT</a><span class="muted">仅写入本次所选章节范围内的生成文本，其余页保持原样。</span>';
      card.appendChild(row);
    }
    var outer = document.createElement("details");
    outer.className = "result-slides-wrap";
    outer.open = false;
    var outerSum = document.createElement("summary");
    outerSum.innerHTML =
      '<span>各页可填内容</span><span class="summary-meta">共 ' +
      data.slides.length +
      " 页 · 点击展开</span>";
    outer.appendChild(outerSum);
    var outerBody = document.createElement("div");
    outerBody.className = "details-body";
    var stack = document.createElement("div");
    stack.className = "stack";
    data.slides.forEach(function (slide) {
      var comps = slide.components || [];
      var det = document.createElement("details");
      det.className = "result-slide-details";
      det.open = false;
      var sum = document.createElement("summary");
      sum.innerHTML =
        "<span>第 " +
        slide.slide_index +
        " 页可填内容</span><span class=\"summary-meta\">" +
        comps.length +
        " 条</span>";
      det.appendChild(sum);
      var body = document.createElement("div");
      body.className = "details-body";
      var list = document.createElement("div");
      list.className = "component-list";
      comps.forEach(function (comp, j) {
        var idx = j + 1;
        var sid = "comp-" + slide.slide_index + "-" + idx;
        var item = document.createElement("div");
        item.className = "component-item";
        var ch = document.createElement("div");
        ch.className = "component-head";
        ch.innerHTML =
          '<span class="component-meta">可填内容 ' +
          idx +
          '</span><button type="button" class="button secondary copy-btn" data-target="' +
          sid +
          '">复制</button>';
        var pre2 = document.createElement("pre");
        pre2.className = "component-text";
        pre2.id = sid;
        pre2.textContent = comp.generated_text || "";
        item.appendChild(ch);
        item.appendChild(pre2);
        list.appendChild(item);
      });
      body.appendChild(list);
      det.appendChild(body);
      stack.appendChild(det);
    });
    outerBody.appendChild(stack);
    outer.appendChild(outerBody);
    card.appendChild(outer);
    var raw = document.createElement("details");
    raw.className = "json-details";
    raw.style.marginTop = "16px";
    raw.innerHTML =
      '<summary><span>高级：查看 AI 结果 JSON（原始）</span><span class="summary-meta">默认折叠</span></summary><div class="details-body"><pre></pre></div>';
    raw.querySelector("pre").textContent = JSON.stringify(data, null, 2);
    card.appendChild(raw);
    mount.appendChild(card);
  }

  function setupCopyDelegation() {
    if (document.body.getAttribute("data-ppt-copy-delegation") === "1") return;
    document.body.setAttribute("data-ppt-copy-delegation", "1");
    document.body.addEventListener("click", async function (ev) {
      var btn = ev.target.closest(".copy-btn[data-target]");
      if (!btn) return;
      var targetId = btn.getAttribute("data-target");
      var pre = document.getElementById(targetId);
      var text = pre ? pre.innerText : "";
      if (!text) return;
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(text);
        } else {
          var ta = document.createElement("textarea");
          ta.value = text;
          document.body.appendChild(ta);
          ta.select();
          document.execCommand("copy");
          document.body.removeChild(ta);
        }
        var old = btn.textContent;
        btn.textContent = "已复制";
        setTimeout(function () {
          btn.textContent = old;
        }, 1200);
      } catch (e) {
        /* 忽略 */
      }
    });
  }

  /**
   * Tab 上标记「已纳入本次生成」
   * @param {HTMLFormElement} genForm
   */
  function syncChapterTabIncludedMarkers(genForm) {
    var root = genForm.querySelector("[data-chapter-tabs]");
    if (!root) return;
    var tabs = root.querySelectorAll('[role="tab"]');
    var panels = root.querySelectorAll('[role="tabpanel"]');
    panels.forEach(function (panel, i) {
      var cb = panel.querySelector(".group-select");
      var btn = tabs[i];
      if (!cb || !btn) return;
      btn.classList.toggle("chapter-tab-btn--included", cb.checked);
    });
  }

  /**
   * 章节块 Tab 切换（键盘 ← →）
   * @param {HTMLFormElement} genForm
   */
  function initChapterTabs(genForm) {
    var root = genForm.querySelector("[data-chapter-tabs]");
    if (!root) return;
    var tabs = Array.prototype.slice.call(root.querySelectorAll('[role="tab"]'));
    var panels = Array.prototype.slice.call(root.querySelectorAll('[role="tabpanel"]'));
    if (!tabs.length || tabs.length !== panels.length) return;

    function activate(index) {
      var i = Math.max(0, Math.min(index, tabs.length - 1));
      tabs.forEach(function (tab, j) {
        var on = j === i;
        tab.setAttribute("aria-selected", on ? "true" : "false");
        tab.tabIndex = on ? 0 : -1;
      });
      panels.forEach(function (panel, j) {
        if (j === i) panel.removeAttribute("hidden");
        else panel.setAttribute("hidden", "");
      });
    }

    tabs.forEach(function (tab, i) {
      tab.addEventListener("click", function () {
        activate(i);
      });
      tab.addEventListener("keydown", function (ev) {
        if (ev.key !== "ArrowRight" && ev.key !== "ArrowLeft") return;
        ev.preventDefault();
        var next = ev.key === "ArrowRight" ? i + 1 : i - 1;
        if (next < 0) next = tabs.length - 1;
        if (next >= tabs.length) next = 0;
        activate(next);
        tabs[next].focus();
      });
    });
  }

  /**
   * @param {HTMLFormElement} genForm
   */
  function bindGenerateForm(genForm) {
    if (!genForm) return;
    var hiddenMount = genForm.querySelector("#selected-slides-hidden");
    if (!hiddenMount) return;

    function groupSelectEls() {
      return Array.from(genForm.querySelectorAll(".group-select"));
    }

    function ensureAllGroupsSelected() {
      groupSelectEls().forEach(function (cb) {
        cb.checked = true;
      });
    }

    function syncChapterSelectionToForm() {
      if (!hiddenMount) return;
      hiddenMount.innerHTML = "";
      var seen = new Set();
      groupSelectEls().forEach(function (cb) {
        if (!cb.checked) return;
        var raw = cb.getAttribute("data-slides") || "";
        raw.split(",").forEach(function (piece) {
          var n = parseInt(piece.trim(), 10);
          if (!Number.isFinite(n) || seen.has(n)) return;
          seen.add(n);
          var inp = document.createElement("input");
          inp.type = "hidden";
          inp.name = "selected_slides";
          inp.value = String(n);
          hiddenMount.appendChild(inp);
        });
      });
    }

    ensureAllGroupsSelected();
    syncChapterSelectionToForm();
    initChapterTabs(genForm);
    syncChapterTabIncludedMarkers(genForm);

    genForm._resyncChapters = function () {
      ensureAllGroupsSelected();
      syncChapterSelectionToForm();
      syncChapterTabIncludedMarkers(genForm);
    };

    var progressWrap = genForm.querySelector("#generate-progress");
    var progressMsg = genForm.querySelector("#generate-progress-message");
    var progressBar = genForm.querySelector("#generate-progress-bar");
    var progressErr = genForm.querySelector("#generate-progress-error");
    var genSubmit = genForm.querySelector("#ai-generate-submit");

    function setGenUiRunning(isRun) {
      if (genSubmit) {
        if (isRun) {
          genSubmit.disabled = true;
          genSubmit.classList.remove("gen-submit--invalid");
        } else if (window.PptApp && window.PptApp.syncGenerateSubmitEnabled) {
          window.PptApp.syncGenerateSubmitEnabled();
        } else {
          genSubmit.disabled = false;
        }
      }
      if (!progressWrap) return;
      progressWrap.classList.toggle("is-visible", isRun || !!(progressErr && !progressErr.hidden));
      progressWrap.classList.toggle("is-running", isRun);
      if (!isRun && progressErr && progressErr.hidden) {
        progressWrap.classList.remove("is-visible");
      }
    }

    function applyStatusToBar(s) {
      var t = s.batch_total || 1;
      var i = s.batch_index || 0;
      var pct = 100;
      if (s.status === "running") {
        pct = Math.min(95, Math.max(2, Math.round(((Math.max(i, 1) - 1) / Math.max(t, 1)) * 100)));
      } else if (s.status === "error") {
        pct = 0;
      } else if (s.status === "done") {
        pct = 100;
      }
      if (progressBar) progressBar.style.width = pct + "%";
      if (progressMsg) progressMsg.textContent = s.message || "";
    }

    genForm.addEventListener("submit", async function (e) {
      e.preventDefault();
      if (genSubmit && genSubmit.disabled) {
        return;
      }
      ensureAllGroupsSelected();
      syncChapterSelectionToForm();
      var preMsgs =
        window.PptApp && window.PptApp.collectGenerateValidationMessages
          ? window.PptApp.collectGenerateValidationMessages(genForm)
          : [];
      if (preMsgs.length) {
        window.alert(preMsgs.join("\n"));
        return;
      }
      if (progressErr) {
        progressErr.hidden = true;
        progressErr.textContent = "";
      }
      var fd = new FormData(genForm);
      setGenUiRunning(true);
      if (progressBar) progressBar.style.width = "2%";
      if (progressMsg) progressMsg.textContent = "正在提交任务…";
      var jobId;
      try {
        var res = await fetch("/api/generate_start", { method: "POST", body: fd });
        var start = await res.json();
        if (!start.ok) throw new Error(start.error || "启动失败");
        jobId = start.job_id;
      } catch (err) {
        setGenUiRunning(false);
        if (progressErr) {
          progressErr.hidden = false;
          progressErr.textContent = err.message || String(err);
        }
        if (progressWrap) progressWrap.classList.add("is-visible");
        return;
      }
      for (;;) {
        var s;
        try {
          var r = await fetch("/api/generate_status/" + encodeURIComponent(jobId));
          s = await r.json();
          if (!r.ok || !s.ok) throw new Error(s.error || "状态查询失败");
        } catch (err) {
          setGenUiRunning(false);
          if (progressErr) {
            progressErr.hidden = false;
            progressErr.textContent = err.message || String(err);
          }
          if (progressWrap) progressWrap.classList.add("is-visible");
          break;
        }
        applyStatusToBar(s);
        if (s.status === "done") {
          setGenUiRunning(false);
          renderGenerationResult(s.result, s.task_id, s.history_id || "");
          var target = document.getElementById("ai-result");
          if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
          break;
        }
        if (s.status === "error") {
          setGenUiRunning(false);
          if (progressErr) {
            progressErr.hidden = false;
            progressErr.textContent = s.error || s.message || "生成失败";
          }
          if (progressWrap) progressWrap.classList.add("is-visible");
          break;
        }
        await sleep(450);
      }
    });
  }

  window.PptApp = window.PptApp || {};
  window.PptApp.bindGenerateForm = bindGenerateForm;
  window.PptApp.renderGenerationResult = renderGenerationResult;

  setupCopyDelegation();
})();
