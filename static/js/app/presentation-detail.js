/**
 * 解析详情页：修改数据库中的模板显示名称
 */
(function () {
  function setStatus(el, msg, isError) {
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

  function init() {
    var btn = document.getElementById("presentation-display-name-save");
    var inp = document.getElementById("presentation-display-name-input");
    var status = document.getElementById("presentation-rename-status");
    var label = document.getElementById("presentation-display-name-label");
    if (!btn || !inp) return;
    var taskId = (btn.getAttribute("data-task-id") || "").trim();
    if (!taskId) return;

    btn.addEventListener("click", async function () {
      var name = (inp.value || "").trim();
      if (!name) {
        setStatus(status, "名称不能为空。", true);
        return;
      }
      setStatus(status, "保存中…", false);
      btn.disabled = true;
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
        var saved = data.file_name != null ? String(data.file_name) : name;
        inp.value = saved;
        if (label) label.textContent = saved;
        document.title = "解析详情 · " + saved;
        setStatus(status, "已保存。", false);
        setTimeout(function () {
          setStatus(status, "", false);
        }, 2500);
      } catch (e) {
        setStatus(status, e.message || String(e), true);
      } finally {
        btn.disabled = false;
      }
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
