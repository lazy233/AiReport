/**
 * 章节 Tab 条 + 内容区（详情 / 编辑 共用）
 */
(function (global) {
  function esc(s) {
    var d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  /**
   * @param {HTMLElement} root — 包含 #ct-tab-bar、#ct-tab-panel
   * @param {{ mode: string, initialDraft: object }} options
   */
  function mount(root, options) {
    if (!root || !window.ChapterTemplatesStore) return;

    var mode = options.mode || "detail";
    var isEditor = mode === "editor";
    var draft = options.initialDraft;
    var activeIdx = 0;

    var bar = root.querySelector("#ct-tab-bar");
    var panel = root.querySelector("#ct-tab-panel");

    if (!draft || !draft.chapters || !draft.chapters.length) {
      if (panel) {
        panel.innerHTML = '<p class="error">没有可用的章节数据。</p>';
      }
      return;
    }

    async function removeChapterAt(index) {
      if (draft.chapters.length <= 1) {
        alert("至少保留一个章节。");
        return;
      }
      var ch = draft.chapters[index];
      if (!ch) return;
      var msg = isEditor
        ? "确定删除章节「" + (ch.title || "") + "」吗？"
        : "确定从模板中删除章节「" + (ch.title || "") + "」吗？删除后将立即保存。";
      if (!confirm(msg)) return;
      draft.chapters.splice(index, 1);
      draft.chapters.forEach(function (c, i) { c.sort = i; });
      if (activeIdx >= draft.chapters.length) activeIdx = draft.chapters.length - 1;
      if (!isEditor) {
        try {
          await ChapterTemplatesStore.save(JSON.parse(JSON.stringify(draft)));
        } catch (e) {
          alert(e.message || String(e));
          return;
        }
      }
      renderTabs();
      renderPanel();
    }

    function renderTabs() {
      if (!bar) return;
      bar.innerHTML = "";
      draft.chapters.forEach(function (ch, i) {
        var tab = document.createElement("div");
        tab.className = "ct-tab" + (i === activeIdx ? " is-active" : "");
        tab.setAttribute("role", "tab");

        var label = document.createElement("button");
        label.type = "button";
        label.className = "ct-tab-label";
        label.textContent = ch.title || "未命名章节";
        label.addEventListener("click", function (e) {
          e.stopPropagation();
          activeIdx = i;
          renderTabs();
          renderPanel();
        });

        tab.appendChild(label);
        if (isEditor) {
          var close = document.createElement("button");
          close.type = "button";
          close.className = "ct-tab-close";
          close.setAttribute("aria-label", "删除章节");
          close.innerHTML = "&times;";
          close.addEventListener("click", function (e) {
            e.preventDefault();
            e.stopPropagation();
            removeChapterAt(i).catch(function (err) {
              alert(err && err.message ? err.message : String(err));
            });
          });
          tab.appendChild(close);
        }
        bar.appendChild(tab);
      });

      if (isEditor) {
        var addBtn = document.createElement("button");
        addBtn.type = "button";
        addBtn.className = "button secondary ct-tab-add";
        addBtn.textContent = "+ 添加章节";
        addBtn.addEventListener("click", function () {
          draft.chapters.push({
            id: ChapterTemplatesStore.generateId("ch"),
            title: "新章节",
            hint: "",
            sort: draft.chapters.length,
          });
          activeIdx = draft.chapters.length - 1;
          renderTabs();
          renderPanel();
        });
        bar.appendChild(addBtn);
      }
    }

    function updateActiveTabLabel(text) {
      if (!bar) return;
        var tabs = bar.querySelectorAll(".ct-tab");
        var tab = tabs[activeIdx];
        if (tab) {
          var lb = tab.querySelector(".ct-tab-label");
          if (lb) lb.textContent = text || "未命名章节";
        }
    }

    function renderPanel() {
      if (!panel) return;
      var ch = draft.chapters[activeIdx];
      if (!ch) return;

      if (isEditor) {
        panel.innerHTML = "";
        var f1 = document.createElement("div");
        f1.className = "ct-field";
        var l1 = document.createElement("label");
        l1.htmlFor = "ct-ch-title";
        l1.textContent = "章节标题";
        var ti = document.createElement("input");
        ti.type = "text";
        ti.id = "ct-ch-title";
        ti.className = "ct-input";
        ti.value = ch.title || "";
        ti.addEventListener("input", function () {
          ch.title = ti.value;
          updateActiveTabLabel(ch.title);
        });
        f1.appendChild(l1);
        f1.appendChild(ti);

        var f2 = document.createElement("div");
        f2.className = "ct-field";
        var l2 = document.createElement("label");
        l2.htmlFor = "ct-ch-hint";
        l2.textContent = "约束说明（可选）";
        var ta = document.createElement("textarea");
        ta.id = "ct-ch-hint";
        ta.className = "ct-textarea";
        ta.rows = 4;
        ta.placeholder = "描述本章节应包含哪些内容、版式建议等";
        ta.value = ch.hint || "";
        ta.addEventListener("input", function () {
          ch.hint = ta.value;
        });
        f2.appendChild(l2);
        f2.appendChild(ta);

        panel.appendChild(f1);
        panel.appendChild(f2);
      } else {
        panel.innerHTML =
          '<div class="ct-readonly-block">' +
          "<h3>" +
          esc(ch.title || "未命名章节") +
          "</h3>" +
          (ch.hint
            ? '<p class="muted" style="white-space:pre-wrap;">' + esc(ch.hint) + "</p>"
            : '<p class="muted">（未填写约束说明）</p>') +
          "</div>";
      }
    }

    renderTabs();
    renderPanel();

    if (isEditor) {
      var nameEl = document.getElementById("ct-meta-name");
      var descEl = document.getElementById("ct-meta-desc");
      if (nameEl) nameEl.value = draft.name || "";
      if (descEl) descEl.value = draft.description || "";

      var saveBtn = document.getElementById("ct-editor-save");
      if (saveBtn) {
        saveBtn.addEventListener("click", async function () {
          var n = document.getElementById("ct-meta-name");
          var d = document.getElementById("ct-meta-desc");
          draft.name = n ? n.value.trim() : "";
          draft.description = d ? d.value.trim() : "";
          try {
            var id = await ChapterTemplatesStore.save(JSON.parse(JSON.stringify(draft)));
            window.location.href = "/chapter-templates/" + encodeURIComponent(id);
          } catch (e) {
            alert(e.message || String(e));
          }
        });
      }
    }
  }

  global.ChapterTemplateTabs = { mount: mount };
})(window);
