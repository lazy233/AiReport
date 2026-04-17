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
        { key: "gradeIntake", label: "年级/入学时间" },
        { key: "currentTerm", label: "当前学期" },
        { key: "product", label: "产品" },
        { key: "serviceStart", label: "服务开始日期" },
        { key: "plannerTeacher", label: "规划老师" },
        { key: "advisorTeacher", label: "教服老师" },
        { key: "studentId", label: "学号", detailHidden: true },
        { key: "className", label: "班级", detailHidden: true },
        { key: "email", label: "邮箱", detailHidden: true },
        { key: "phone", label: "手机", detailHidden: true },
        { key: "remark", label: "备注", detailHidden: true, fullWidth: true }
      ]
    },
    {
      id: "learning",
      title: "学习画像",
      fields: [
        { key: "strength_subjects", label: "擅长科目" },
        { key: "scores", label: "语言/国际成绩" },
        { key: "learning_good", label: "擅长学习形式" },
        { key: "learning_weak", label: "不擅长学习形式", fullWidth: true },
        { key: "interests", label: "兴趣方向" },
        { key: "study_goal", label: "升学意向" },
        { key: "career_goal", label: "就业意向" },
        { key: "long_goal", label: "长远目标", fullWidth: true },
        { key: "degree", label: "学位" },
        { key: "duration", label: "学制" },
        { key: "credits", label: "学分要求", fullWidth: true },
        { key: "course_rule", label: "课程要求", fullWidth: true },
        { key: "gpa_rule", label: "GPA要求", fullWidth: true },
        { key: "selection_rule", label: "选课规则", fullWidth: true },
        { key: "recommended_courses", label: "推荐课程", fullWidth: true },
        { key: "course_notes", label: "课程说明", fullWidth: true },
        { key: "term_plan", label: "学期规划", fullWidth: true },
        { key: "future_plan", label: "后续规划", fullWidth: true }
      ]
    },
    {
      id: "hours",
      title: "课时数据",
      fields: [
        { key: "totalHours", label: "总课时" },
        { key: "usedHours", label: "已用课时" },
        { key: "remainingHours", label: "剩余课时" },
        { key: "prep_courses", label: "预习课程" },
        { key: "tutoring_courses", label: "同步辅导课程" },
        { key: "skillDirection", label: "技能提升方向" },
        { key: "skillDescription", label: "技能提升描述", fullWidth: true }
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
    },
    {
      id: "term_summary",
      title: "学期总结/结单",
      fields: [
        { key: "student_summary", label: "学生情况", fullWidth: true },
        { key: "school_ddl", label: "校方DDL" },
        { key: "first_class_time", label: "首课时间" },
        { key: "first_class_note", label: "首课记录", fullWidth: true },
        { key: "summer_work", label: "暑期辅导" },
        { key: "term_work", label: "学期辅导" },
        { key: "recorded_courses", label: "录播完成" },
        { key: "grades", label: "成绩明细", fullWidth: true },
        { key: "gpa", label: "GPA" },
        { key: "target_gpa", label: "目标GPA" },
        { key: "final_score", label: "最终成绩" },
        { key: "services", label: "服务内容", fullWidth: true },
        { key: "service_count", label: "服务次数" },
        { key: "class_count", label: "课程次数" },
        { key: "total_duration", label: "总时长" },
        { key: "avg_duration", label: "平均课时" },
        { key: "communication", label: "沟通频次" },
        { key: "next_goal", label: "下阶段目标" },
        { key: "risk_courses", label: "风险科目" },
        { key: "suggestions", label: "下阶段建议", fullWidth: true },
        { key: "remarks", label: "备注", fullWidth: true }
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

    generateGuidanceWithAi: async function (profile, content) {
      var data = await requestJson(API_BASE + "/ai-guidance", {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify({
          profile: profile && typeof profile === "object" ? profile : {},
          content: content != null ? String(content) : ""
        })
      });
      return data;
    },

    importAiFile: async function (file) {
      var fd = new FormData();
      fd.append("file", file);
      var res = await fetch(API_BASE + "/import-ai", {
        method: "POST",
        body: fd
      });
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
