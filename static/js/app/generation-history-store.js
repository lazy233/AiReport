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
     * 历史下载：优先缓存成品，不存在则服务端回退导出
     * @param {string} recordId
     * @returns {Promise<void>}
     */
    downloadHistoryPptxById: async function (recordId) {
      if (!recordId) throw new Error("缺少历史记录 ID。");
      var res = await fetch("/api/generation_history/" + encodeURIComponent(recordId) + "/download");
      if (!res.ok) {
        var errText = "";
        try {
          var j = await res.json();
          errText = j.error || "";
        } catch (e) {
          errText = await res.text();
        }
        throw new Error(errText || "历史下载失败（" + res.status + "）");
      }
      var blob = await res.blob();
      var cd = res.headers.get("Content-Disposition") || "";
      var name = "history_filled.pptx";
      var m = /filename\*?=(?:UTF-8'')?["']?([^\"';]+)/i.exec(cd);
      if (m) {
        try {
          name = decodeURIComponent(m[1].trim());
        } catch (e) {
          name = m[1].trim();
        }
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

    /**
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
