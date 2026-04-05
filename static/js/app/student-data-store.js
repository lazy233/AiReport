/**
 * 学生数据存储层（后端 API 版）
 */
(function (global) {
  var API_BASE = "/api/student-data";

  var PROFILE_SPEC = [
    {
      id: "basic",
      title: "基础信息",
      fields: [
        { key: "studentName", label: "学生姓名" },
        { key: "nicknameEn", label: "学生昵称 / 英文名" },
        { key: "school", label: "就读院校" },
        { key: "major", label: "专业方向" },
        { key: "gradeLevel", label: "年级" },
        { key: "currentTerm", label: "当前学期" },
        { key: "reportSubtitle", label: "报告副标题" },
        { key: "serviceStart", label: "服务开始日期" },
        { key: "plannerTeacher", label: "规划老师" },
        { key: "studentId", label: "学号", detailHidden: true },
        { key: "className", label: "班级", detailHidden: true },
        { key: "email", label: "邮箱", detailHidden: true },
        { key: "phone", label: "手机", detailHidden: true },
        { key: "remark", label: "备注", detailHidden: true }
      ]
    },
    {
      id: "learning",
      title: "学习画像",
      fields: [
        { key: "strongSubjects", label: "优势科目" },
        { key: "intlScores", label: "国际成绩" },
        { key: "studyIntent", label: "升学意向" },
        { key: "careerIntent", label: "就业意向" },
        { key: "interestSubjects", label: "兴趣科目" },
        { key: "longTermPlan", label: "长期规划" },
        { key: "learningStyle", label: "擅长学习方式" },
        { key: "weakAreas", label: "薄弱环节" }
      ]
    },
    {
      id: "hours",
      title: "课时数据",
      fields: [
        { key: "totalHours", label: "总课时" },
        { key: "usedHours", label: "已用课时" },
        { key: "remainingHours", label: "剩余课时" },
        { key: "tutorSubjects", label: "辅导科目" },
        { key: "previewSubjects", label: "预习科目" },
        { key: "skillDirection", label: "技能提升方向" },
        { key: "skillDescription", label: "技能提升描述" }
      ]
    },
    {
      id: "guidance",
      title: "成长指导数据",
      fields: [
        { key: "termSummary", label: "学期表现概述" },
        { key: "courseFeedback", label: "课程反馈与建议" },
        { key: "shortTermAdvice", label: "短期学习建议" },
        { key: "longTermDevelopment", label: "长期发展规划" }
      ]
    }
  ];

  function _normStr(v) {
    return v == null ? "" : String(v).trim();
  }

  function emptyProfileDims() {
    var out = {};
    PROFILE_SPEC.forEach(function (dim) {
      out[dim.id] = {};
      dim.fields.forEach(function (f) {
        out[dim.id][f.key] = "";
      });
    });
    return out;
  }

  function normalizeRecord(r) {
    var raw = r && typeof r === "object" ? r : {};
    var rawProfile = raw.profile && typeof raw.profile === "object" ? raw.profile : {};
    var dims = emptyProfileDims();
    PROFILE_SPEC.forEach(function (dim) {
      var src = rawProfile[dim.id];
      if (!src || typeof src !== "object") return;
      dim.fields.forEach(function (f) {
        if (src[f.key] != null) dims[dim.id][f.key] = _normStr(src[f.key]);
      });
    });

    if (!dims.basic.studentName) dims.basic.studentName = _normStr(raw.name);
    if (!dims.basic.studentId) dims.basic.studentId = _normStr(raw.studentId);
    if (!dims.basic.className) dims.basic.className = _normStr(raw.className);
    if (!dims.basic.email) dims.basic.email = _normStr(raw.email);
    if (!dims.basic.phone) dims.basic.phone = _normStr(raw.phone);
    if (!dims.basic.remark) dims.basic.remark = _normStr(raw.remark);

    return {
      id: _normStr(raw.id),
      name: dims.basic.studentName,
      studentId: dims.basic.studentId,
      className: dims.basic.className,
      email: dims.basic.email,
      phone: dims.basic.phone,
      remark: dims.basic.remark,
      content: _normStr(raw.content),
      profile: dims,
      createdAt: _normStr(raw.createdAt),
      updatedAt: _normStr(raw.updatedAt)
    };
  }

  function mergeDisplayProfile(rec) {
    return normalizeRecord(rec || {}).profile;
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
      var errMsg = (body && body.error) || ("请求失败（" + res.status + "）");
      throw new Error(errMsg);
    }
    return body;
  }

  function normKey(h) {
    return String(h || "")
      .replace(/\ufeff/g, "")
      .replace(/\s/g, "")
      .toLowerCase();
  }

  function cellValue(row, aliases) {
    for (var rk in row) {
      if (!Object.prototype.hasOwnProperty.call(row, rk)) continue;
      var nk = normKey(rk);
      for (var i = 0; i < aliases.length; i++) {
        if (normKey(aliases[i]) === nk) return row[rk];
      }
    }
    return "";
  }

  function rowToRecord(row) {
    var rec = {
      name: cellValue(row, ["姓名", "name", "学生姓名"]),
      studentId: cellValue(row, ["学号", "studentid", "student_id", "number"]),
      className: cellValue(row, ["班级", "classname", "class", "班级名称"]),
      email: cellValue(row, ["邮箱", "email", "e-mail"]),
      phone: cellValue(row, ["手机", "电话", "phone", "tel", "mobile"]),
      remark: cellValue(row, ["备注", "remark", "note", "说明"]),
      content: cellValue(row, ["数据内容", "content", "data", "正文", "生成数据", "ppt数据", "材料"])
    };
    if (row && typeof row.profile === "object" && row.profile !== null) rec.profile = row.profile;
    return normalizeRecord(rec);
  }

  function splitCSVLine(line) {
    var result = [];
    var cur = "";
    var inQ = false;
    for (var i = 0; i < line.length; i++) {
      var c = line[i];
      if (c === '"') {
        if (inQ && line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQ = !inQ;
        }
      } else if ((c === "," && !inQ) || (c === "\t" && !inQ)) {
        result.push(cur);
        cur = "";
      } else {
        cur += c;
      }
    }
    result.push(cur);
    return result;
  }

  function parseCSV(text) {
    var lines = String(text || "")
      .split(/\r?\n/)
      .map(function (l) { return l.replace(/\s+$/, ""); })
      .filter(function (l) { return l.length > 0; });
    if (!lines.length) return [];
    var headers = splitCSVLine(lines[0]);
    var rows = [];
    for (var i = 1; i < lines.length; i++) {
      var cols = splitCSVLine(lines[i]);
      var row = {};
      headers.forEach(function (h, j) {
        row[String(h).trim()] = cols[j] != null ? String(cols[j]).trim() : "";
      });
      rows.push(row);
    }
    return rows;
  }

  function parseJSONImport(text) {
    var data;
    try {
      data = JSON.parse(String(text || ""));
    } catch (_e) {
      data = null;
    }
    if (data == null) return [];
    if (Array.isArray(data)) return data;
    if (typeof data === "object") return [data];
    return [];
  }

  var Store = {
    list: async function (query) {
      var q = _normStr(query);
      var url = API_BASE + (q ? "?q=" + encodeURIComponent(q) : "");
      var data = await requestJson(url, { method: "GET" });
      return (data.items || []).map(normalizeRecord);
    },

    get: async function (id) {
      var rid = _normStr(id);
      if (!rid) return null;
      try {
        var data = await requestJson(API_BASE + "/" + encodeURIComponent(rid), { method: "GET" });
        return normalizeRecord(data.item || {});
      } catch (e) {
        if (String(e && e.message || "").indexOf("未找到") >= 0) return null;
        throw e;
      }
    },

    save: async function (rec) {
      var r = normalizeRecord(rec || {});
      if (!r.name && !r.studentId && !r.content) {
        throw new Error("请至少填写姓名、学号或数据内容中的一项。");
      }
      var isUpdate = !!_normStr(r.id);
      var data = await requestJson(
        isUpdate ? API_BASE + "/" + encodeURIComponent(r.id) : API_BASE,
        {
          method: isUpdate ? "PUT" : "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(r)
        }
      );
      return _normStr(data.id || (data.item && data.item.id));
    },

    remove: async function (id) {
      var rid = _normStr(id);
      if (!rid) return;
      await requestJson(API_BASE + "/" + encodeURIComponent(rid), { method: "DELETE" });
    },

    importFromText: async function (text, kind) {
      var rows = kind === "json" ? parseJSONImport(text) : parseCSV(text).map(rowToRecord);
      var imported = 0;
      var skipped = 0;
      var errors = [];
      for (var i = 0; i < rows.length; i++) {
        var raw = kind === "json" ? rowToRecord(rows[i]) : normalizeRecord(rows[i]);
        if (!raw.name && !raw.studentId && !raw.content) {
          skipped++;
          continue;
        }
        raw.id = "";
        try {
          await Store.save(raw);
          imported++;
        } catch (e) {
          errors.push("第 " + (i + 1) + " 行：" + (e.message || String(e)));
        }
      }
      return { imported: imported, skipped: skipped, errors: errors };
    },

    filterByQuery: function (records, q) {
      var s = _normStr(q).toLowerCase();
      if (!s) return records || [];
      return (records || []).filter(function (r) {
        var rr = normalizeRecord(r || {});
        var parts = [rr.name, rr.studentId, rr.className, rr.email, rr.phone, rr.remark, rr.content];
        PROFILE_SPEC.forEach(function (dim) {
          dim.fields.forEach(function (f) {
            parts.push(_normStr(rr.profile && rr.profile[dim.id] && rr.profile[dim.id][f.key]));
          });
        });
        return parts.join("\n").toLowerCase().indexOf(s) >= 0;
      });
    },

    profileSpec: PROFILE_SPEC,

    profileSpecForDetail: function () {
      return PROFILE_SPEC.map(function (dim) {
        return {
          id: dim.id,
          title: dim.title,
          fields: dim.fields.filter(function (f) { return !f.detailHidden; })
        };
      });
    },

    mergeDisplayProfile: mergeDisplayProfile
  };

  global.StudentDataStore = Store;
})(window);
