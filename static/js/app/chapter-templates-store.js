/**
 * 章节模板存储层（后端 API 版）
 */
(function (global) {
  var API_BASE = "/api/chapter-templates";

  function generateId(prefix) {
    return (prefix || "id") + "_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 9);
  }

  function norm(v) {
    return v == null ? "" : String(v).trim();
  }

  function normalizeTemplate(tpl) {
    var t = tpl && typeof tpl === "object" ? tpl : {};
    var rows = Array.isArray(t.chapters) ? t.chapters : [];
    rows = rows
      .map(function (ch, i) {
        return {
          id: norm(ch && ch.id) || generateId("ch"),
          title: norm(ch && ch.title) || "未命名章节",
          hint: norm(ch && ch.hint),
          sort: typeof (ch && ch.sort) === "number" ? ch.sort : i,
        };
      })
      .sort(function (a, b) { return (a.sort || 0) - (b.sort || 0); })
      .map(function (ch, i) {
        ch.sort = i;
        return ch;
      });
    return {
      id: norm(t.id),
      name: norm(t.name),
      description: norm(t.description),
      chapters: rows,
      chapterCount: rows.length,
      createdAt: norm(t.createdAt),
      updatedAt: norm(t.updatedAt),
    };
  }

  async function requestJson(url, options) {
    var res = await fetch(url, options || {});
    var body = null;
    try {
      body = await res.json();
    } catch (_e) {
      body = null;
    }
    if (!res.ok || !body || body.ok === false) {
      throw new Error((body && body.error) || ("请求失败（" + res.status + "）"));
    }
    return body;
  }

  var Store = {
    list: async function (query) {
      var q = norm(query);
      var data = await requestJson(API_BASE + (q ? "?q=" + encodeURIComponent(q) : ""), { method: "GET" });
      return (data.items || []).map(function (it) {
        return {
          id: norm(it.id),
          name: norm(it.name),
          description: norm(it.description),
          chapterCount: Number(it.chapterCount || 0),
          updatedAt: norm(it.updatedAt),
        };
      });
    },

    get: async function (id) {
      var tid = norm(id);
      if (!tid) return null;
      try {
        var data = await requestJson(API_BASE + "/" + encodeURIComponent(tid), { method: "GET" });
        return normalizeTemplate(data.item || {});
      } catch (e) {
        if (String(e && e.message || "").indexOf("未找到") >= 0) return null;
        throw e;
      }
    },

    save: async function (tpl) {
      var t = normalizeTemplate(tpl);
      if (!t.name) throw new Error("请填写模板名称。");
      if (!t.chapters.length) throw new Error("至少保留一个章节。");
      var isUpdate = !!t.id;
      var data = await requestJson(
        isUpdate ? API_BASE + "/" + encodeURIComponent(t.id) : API_BASE,
        {
          method: isUpdate ? "PUT" : "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(t),
        },
      );
      return norm(data.id || (data.item && data.item.id));
    },

    remove: async function (id) {
      var tid = norm(id);
      if (!tid) return;
      await requestJson(API_BASE + "/" + encodeURIComponent(tid), { method: "DELETE" });
    },

    removeChapter: async function (templateId, chapterId) {
      var tpl = await this.get(templateId);
      if (!tpl) throw new Error("模板不存在。");
      tpl.chapters = (tpl.chapters || []).filter(function (ch) { return ch.id !== chapterId; });
      if (!tpl.chapters.length) throw new Error("至少保留一个章节。");
      tpl.chapters.forEach(function (ch, i) { ch.sort = i; });
      return await this.save(tpl);
    },

    newEmpty: function () {
      var now = new Date().toISOString();
      return normalizeTemplate(
        {
          id: "",
          name: "",
          description: "",
          chapters: [{ id: generateId("ch"), title: "新章节", hint: "", sort: 0 }],
          createdAt: now,
          updatedAt: now,
        },
      );
    },

    generateId: generateId,
  };

  global.ChapterTemplatesStore = Store;
})(window);
