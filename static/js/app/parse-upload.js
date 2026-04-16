/**
 * 上传 PPT → /api/parse_start 异步解析与轮询（拖放 + 自定义选择区）
 */
(function () {
  function lowerName(f) {
    return ((f && f.name) || "").toLowerCase();
  }

  function isPptxName(name) {
    return name.endsWith(".pptx");
  }

  function isDocxName(name) {
    return name.endsWith(".docx");
  }

  function isPptOnlyName(name) {
    return name.endsWith(".ppt") && !name.endsWith(".pptx");
  }

  function initParseUpload() {
    var form = document.getElementById("parse-upload-form");
    if (!form) return;

    var btn = document.getElementById("parse-submit-btn");
    var fileInput = document.getElementById("ppt-file-input");
    var dropzone = document.getElementById("ppt-dropzone");
    var browseBtn = document.getElementById("ppt-browse-btn");
    var fileNameEl = document.getElementById("ppt-file-name");
    var rejectEl = document.getElementById("ppt-file-reject");
    var wrap = document.getElementById("parse-progress");
    var msg = document.getElementById("parse-progress-message");
    var bar = document.getElementById("parse-progress-bar");
    var errEl = document.getElementById("parse-progress-error");

    if (!fileInput) return;

    function showReject(text) {
      if (!rejectEl) return;
      rejectEl.textContent = text || "";
      rejectEl.hidden = !text;
    }

    function hideReject() {
      showReject("");
    }

    function clearFileUi() {
      fileInput.value = "";
      if (fileNameEl) {
        fileNameEl.textContent = "";
        fileNameEl.hidden = true;
      }
    }

    /**
     * @param {File} file
     * @returns {boolean}
     */
    function assignPptFile(file) {
      if (!file) return false;
      var name = lowerName(file);
      if (!isPptxName(name) && !isDocxName(name) && !isPptOnlyName(name)) {
        showReject("请上传 .pptx、.docx 或 .ppt 文件。");
        return false;
      }
      if (isPptOnlyName(name)) {
        showReject("当前仅支持 .pptx 解析。请用 PowerPoint / WPS 将 .ppt 另存为 .pptx 后再上传。");
        return false;
      }
      hideReject();
      try {
        var dt = new DataTransfer();
        dt.items.add(file);
        fileInput.files = dt.files;
      } catch (e) {
        showReject("无法使用该文件，请换一份重试。");
        return false;
      }
      if (fileNameEl) {
        fileNameEl.textContent = "已选择：" + (file.name || "未命名");
        fileNameEl.hidden = false;
      }
      return true;
    }

    function openFilePicker() {
      fileInput.click();
    }

    if (browseBtn) {
      browseBtn.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        openFilePicker();
      });
    }

    if (dropzone) {
      dropzone.addEventListener("click", function (e) {
        if (e.target.closest(".ppt-browse-btn")) return;
        openFilePicker();
      });

      dropzone.addEventListener("keydown", function (e) {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          openFilePicker();
        }
      });

      var dragDepth = 0;
      dropzone.addEventListener("dragenter", function (e) {
        e.preventDefault();
        e.stopPropagation();
        dragDepth++;
        dropzone.classList.add("ppt-dropzone--active");
      });
      dropzone.addEventListener("dragover", function (e) {
        e.preventDefault();
        e.stopPropagation();
        if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
      });
      dropzone.addEventListener("dragleave", function (e) {
        e.preventDefault();
        e.stopPropagation();
        dragDepth = Math.max(0, dragDepth - 1);
        if (dragDepth === 0) dropzone.classList.remove("ppt-dropzone--active");
      });
      dropzone.addEventListener("drop", function (e) {
        e.preventDefault();
        e.stopPropagation();
        dragDepth = 0;
        dropzone.classList.remove("ppt-dropzone--active");
        var files = e.dataTransfer && e.dataTransfer.files;
        if (!files || !files.length) return;
        assignPptFile(files[0]);
      });
    }

    fileInput.addEventListener("change", function () {
      hideReject();
      var f = fileInput.files && fileInput.files[0];
      if (!f) {
        if (fileNameEl) {
          fileNameEl.textContent = "";
          fileNameEl.hidden = true;
        }
        return;
      }
      if (!assignPptFile(f)) {
        clearFileUi();
      }
    });

    function setParseRunning(on) {
      if (btn) btn.disabled = on;
      if (!wrap) return;
      wrap.classList.toggle("is-visible", on || !!(errEl && !errEl.hidden));
      wrap.classList.toggle("is-running", on);
      if (!on && errEl && errEl.hidden) wrap.classList.remove("is-visible");
    }

    function applyParsePhase(s) {
      var ph = s.phase || "";
      var pct = 15;
      if (s.status === "done") pct = 100;
      else if (s.status === "error") pct = 0;
      else if (ph === "queued") pct = 8;
      else if (ph === "storing") pct = 72;
      else if (ph === "parsing") pct = 45;
      else if (ph === "classifying") pct = 78;
      if (bar) bar.style.width = pct + "%";
      if (msg) msg.textContent = s.message || "";
    }

    form.addEventListener("submit", async function (e) {
      e.preventDefault();
      if (errEl) {
        errEl.hidden = true;
        errEl.textContent = "";
      }
      if (!fileInput.files || !fileInput.files.length) {
        showReject("请先选择或拖入一个 .pptx 或 .docx 文件。");
        return;
      }
      var fn = lowerName(fileInput.files[0]);
      if (!isPptxName(fn) && !isDocxName(fn)) {
        showReject("请上传 .pptx 或 .docx 后再提交。");
        return;
      }

      setParseRunning(true);
      if (bar) bar.style.width = "5%";
      if (msg)
        msg.textContent = isDocxName(fn) ? "正在上传 Word 文档…" : "正在上传文件…";

      var fd = new FormData(form);
      var jobId;
      try {
        var res = await fetch("/api/parse_start", { method: "POST", body: fd });
        var start = await res.json();
        if (!start.ok) throw new Error(start.error || "启动解析失败");
        jobId = start.job_id;
      } catch (ex) {
        setParseRunning(false);
        if (errEl) {
          errEl.hidden = false;
          errEl.textContent = ex.message || String(ex);
        }
        if (wrap) wrap.classList.add("is-visible");
        return;
      }

      var sleep =
        window.PptApp && PptApp.sleep
          ? PptApp.sleep
          : function (ms) {
              return new Promise(function (r) {
                setTimeout(r, ms);
              });
            };

      for (;;) {
        var s;
        try {
          var r = await fetch("/api/parse_status/" + encodeURIComponent(jobId));
          s = await r.json();
          if (!r.ok || !s.ok) throw new Error(s.error || "状态查询失败");
        } catch (ex) {
          setParseRunning(false);
          if (errEl) {
            errEl.hidden = false;
            errEl.textContent = ex.message || String(ex);
          }
          if (wrap) wrap.classList.add("is-visible");
          break;
        }
        applyParsePhase(s);
        if (s.status === "done" && s.task_id) {
          var doneFn = fileInput.files && fileInput.files[0] ? lowerName(fileInput.files[0].name) : "";
          if (msg)
            msg.textContent = isDocxName(doneFn)
              ? "Word 解析完成，正在刷新列表…"
              : "解析完成，已保存到数据库，正在刷新列表…";
          if (bar) bar.style.width = "100%";
          var orig =
            fileInput.files && fileInput.files[0] && fileInput.files[0].name
              ? fileInput.files[0].name
              : "";
          var q =
            "/upload?parsed=" +
            encodeURIComponent(s.task_id) +
            (orig ? "&parsed_name=" + encodeURIComponent(orig) : "");
          window.location.href = q;
          break;
        }
        if (s.status === "error") {
          setParseRunning(false);
          if (errEl) {
            errEl.hidden = false;
            errEl.textContent = s.error || s.message || "解析失败";
          }
          if (wrap) wrap.classList.add("is-visible");
          break;
        }
        await sleep(400);
      }
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initParseUpload);
  } else {
    initParseUpload();
  }
})();
