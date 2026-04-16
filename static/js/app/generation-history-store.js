/**
 * 文案生成历史（服务端 /api/generation_history）
 */
(function (global) {
  /**
   * @returns {Promise<{ dbEnabled: boolean, items: Array<{id: string, taskId: string, topic: string, createdAt: string, slideCount: number}> }>}
   */
  async function fetchList(query) {
    var q = (query || "").trim();
    var url = "/api/generation_history";
    if (q) url += "?q=" + encodeURIComponent(q);
    var res = await fetch(url);
    var j = await res.json().catch(function () {
      return {};
    });
    if (!res.ok || !j.ok) {
      throw new Error(j.error || "加载历史列表失败（" + res.status + "）");
    }
    return {
      dbEnabled: !!j.db_enabled,
      items: Array.isArray(j.items) ? j.items : [],
    };
  }

  /**
   * @param {string} id
   * @returns {Promise<{ id: string, taskId: string, topic: string, createdAt: string, selectedSlides: number[], result: object } | null>}
   */
  async function fetchOne(id) {
    if (!id) return null;
    var res = await fetch("/api/generation_history/" + encodeURIComponent(id));
    var j = await res.json().catch(function () {
      return {};
    });
    if (res.status === 404) return null;
    if (!res.ok || !j.ok || !j.record) {
      throw new Error(j.error || "加载记录失败（" + res.status + "）");
    }
    return j.record;
  }

  /**
   * @param {string} id
   */
  async function removeRemote(id) {
    var res = await fetch("/api/generation_history/" + encodeURIComponent(id), { method: "DELETE" });
    var j = await res.json().catch(function () {
      return {};
    });
    if (!res.ok || !j.ok) {
      throw new Error(j.error || "删除失败（" + res.status + "）");
    }
  }

  async function cleanupRemote(days) {
    var d = Number(days);
    var res = await fetch("/api/generation_history/cleanup", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ days: d }),
    });
    var j = await res.json().catch(function () {
      return {};
    });
    if (!res.ok || !j.ok) {
      throw new Error(j.error || "清理失败（" + res.status + "）");
    }
    return {
      removed: Number(j.removed || 0),
      days: Number(j.days || d),
    };
  }

  var Store = {
    fetchList: fetchList,
    get: fetchOne,
    remove: removeRemote,
    cleanupByDays: cleanupRemote,

    /**
     * 历史下载：优先缓存成品，不存在则服务端回退导出。
     * 使用同源直链触发浏览器原生下载（可走系统/Chrome 下载进度），不再 fetch 整包进内存。
     * 若服务端返回 JSON 错误体，浏览器可能下载为小文件；列表/详情页在点击前已能确认记录存在。
     * @param {string} recordId
     * @returns {Promise<void>}
     */
    downloadHistoryById: function (recordId) {
      if (!recordId) {
        return Promise.reject(new Error("缺少历史记录 ID。"));
      }
      var url = "/api/generation_history/" + encodeURIComponent(recordId) + "/download";
      var a = document.createElement("a");
      a.href = url;
      a.rel = "noopener";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      return Promise.resolve();
    },
    // backward compatibility
    downloadHistoryPptxById: async function (recordId) {
      return Store.downloadHistoryById(recordId);
    },

    /**
     * 按任务与内存结果导出（POST /api/export），仍需 fetch + Blob；若需原生下载需后端提供 GET 导出或表单 POST。
     * @param {{ taskId: string, result: object }} rec
     * @returns {Promise<void>}
     */
    downloadPptx: async function (rec) {
      if (!rec || !rec.taskId || !rec.result) throw new Error("记录不完整。");
      var res = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ task_id: rec.taskId, data: rec.result }),
      });
      if (!res.ok) {
        var errText = "";
        try {
          var j = await res.json();
          errText = j.error || "";
        } catch (e) {
          errText = await res.text();
        }
        throw new Error(errText || "导出失败（" + res.status + "）");
      }
      var blob = await res.blob();
      var cd = res.headers.get("Content-Disposition") || "";
      var name = "presentation_filled.pptx";
      var m = /filename\*?=(?:UTF-8'')?["']?([^\"';]+)/i.exec(cd);
      if (m) {
        try {
          name = decodeURIComponent(m[1].trim());
        } catch (e) {
          name = m[1].trim();
        }
      } else {
        var stem = (rec.topic || "export").replace(/[\\/:*?"<>|]+/g, "_").slice(0, 80);
        if (stem) name = stem + "_filled.pptx";
      }
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    },
  };

  global.GenerationHistoryStore = Store;
})(window);
