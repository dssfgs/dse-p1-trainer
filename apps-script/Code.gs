/**
 * DSE Maths Paper 1 Trainer — server side (Google Apps Script)
 * Consensus doc v1.4
 *
 * Deployment (required, see SETUP.md):
 *   Execute as        : User accessing the web app
 *   Who has access    : Anyone with Google Account
 */

var SS_ID = ''; // leave '' when the script is bound to the spreadsheet
var SHEETS = {
  Q: 'Questions',
  A: 'Attempts',
  S: 'Scores',
  ST: 'Students',
  T: 'Topics',
  SET: 'Settings'
};
var SECTIONS = ['A1', 'A2', 'B'];
var SECTION_MAX = 35;

var HEADERS = {
  Questions: ['Year', 'Section', 'QNo', 'Part', 'MaxMark', 'TopicID', 'SubTopic', 'PaperURL', 'QID'],
  Attempts: ['AttemptID', 'StudentID', 'Year', 'Section', 'AttemptNo', 'CreatedTime', 'SubmitTime',
             'TotalScore', 'MaxTotal', 'Percent', 'Status', 'Note'],
  Scores: ['AttemptID', 'QID', 'Year', 'Section', 'QNo', 'Part', 'Score', 'MaxMark', 'Rate',
           'TopicID', 'SubmitTime'],
  Students: ['Email', 'DisplayName', 'Role', 'Active', 'JoinDate', 'Class'],
  Topics: ['TopicID', 'Strand', 'NameEN', 'NameZH', 'Level', 'Active'],
  Settings: ['Key', 'Value', 'Remark']
};

var DEFAULT_SETTINGS = {
  WEAK_TOPIC_THRESHOLD: 0.6,
  MIN_PARTS_FOR_TOPIC: 3,
  DRAFT_TTL_DAYS: 30,
  DEFAULT_LANG: 'zh',
  ALLOW_FALLBACK_LOGIN: true,
  ALLOW_TEACHER_EDIT_SUBMITTED: false
};

/* ------------------------------------------------------------------ web app */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('DSE Paper 1 Trainer')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ---------------------------------------------------------------- utilities */

function ss_() {
  return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name + ' — run initSheets() first.');
  return sh;
}

function rows_(name) {
  var values = sheet_(name).getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0].map(function (h) { return String(h).trim(); });
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i].join('') === '') continue;
    var obj = {};
    for (var c = 0; c < head.length; c++) obj[head[c]] = values[i][c];
    out.push(obj);
  }
  return out;
}

function truthy_(v) {
  if (v === true) return true;
  var s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === '1' || s === 'Y';
}

function nowIso_() {
  return new Date().toISOString();
}

function settings_() {
  var out = {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { out[k] = DEFAULT_SETTINGS[k]; });
  rows_(SHEETS.SET).forEach(function (r) {
    var k = String(r.Key).trim();
    if (!k) return;
    var v = r.Value;
    if (typeof DEFAULT_SETTINGS[k] === 'boolean') v = truthy_(v);
    else if (typeof DEFAULT_SETTINGS[k] === 'number') v = Number(v);
    out[k] = v;
  });
  return out;
}

function qid_(year, section, qno, part) {
  return [year, section, 'Q' + qno + (part ? String(part).trim() : '')].join('-');
}

/* ------------------------------------------------------------------ identity */

function currentUser_() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) { email = ''; }
  if (!email) return { email: '', displayName: '', role: 'anonymous', klass: '' };
  var rec = null;
  rows_(SHEETS.ST).forEach(function (r) {
    if (String(r.Email).trim().toLowerCase() === email.toLowerCase() && truthy_(r.Active)) rec = r;
  });
  if (!rec) return { email: email, displayName: email, role: 'unregistered', klass: '' };
  return {
    email: email,
    displayName: String(rec.DisplayName || email),
    role: String(rec.Role || 'student').toLowerCase(),
    klass: String(rec.Class || '')
  };
}

function activeStudents_() {
  return rows_(SHEETS.ST)
    .filter(function (r) { return truthy_(r.Active); })
    .map(function (r) {
      return {
        email: String(r.Email).trim(),
        displayName: String(r.DisplayName || r.Email),
        klass: String(r.Class || ''),
        role: String(r.Role || 'student').toLowerCase()
      };
    });
}

/**
 * Resolves which StudentID this call may write/read.
 * Fallback login is only accepted when Session gives no email
 * AND Settings.ALLOW_FALLBACK_LOGIN is TRUE AND the id is an active student.
 */
function resolveStudent_(override) {
  var u = currentUser_();
  if (u.role === 'student' || u.role === 'teacher') return u.email;
  if (u.role === 'unregistered') {
    throw new Error('此 Google 帳戶未在名單內，請聯絡老師加入 Students 表。');
  }
  var st = settings_();
  if (!st.ALLOW_FALLBACK_LOGIN) throw new Error('無法識別帳戶，且未啟用備用登入。');
  var id = String(override || '').trim().toLowerCase();
  var ok = activeStudents_().some(function (s) { return s.email.toLowerCase() === id; });
  if (!ok) throw new Error('請先選擇你的名字。');
  return id;
}

function requireTeacher_() {
  var u = currentUser_();
  if (u.role !== 'teacher') throw new Error('只有老師可以查看班級總覽。');
  return u;
}

/* ----------------------------------------------------------------- bootstrap */

function getBootstrap(override) {
  var u = currentUser_();
  var st = settings_();
  var years = {};
  rows_(SHEETS.Q).forEach(function (q) {
    if (String(q.Year).trim() !== '') years[String(q.Year).trim()] = true;
  });
  var needPicker = (u.role === 'anonymous' && st.ALLOW_FALLBACK_LOGIN);
  return {
    user: u,
    resolvedStudent: needPicker ? String(override || '') : u.email,
    needPicker: needPicker,
    studentOptions: needPicker ? activeStudents_() : [],
    years: Object.keys(years).sort().reverse(),
    sections: SECTIONS,
    topics: rows_(SHEETS.T).filter(function (t) { return truthy_(t.Active); }).map(function (t) {
      return {
        id: String(t.TopicID).trim(),
        strand: String(t.Strand || ''),
        en: String(t.NameEN || ''),
        zh: String(t.NameZH || ''),
        level: String(t.Level || '')
      };
    }),
    settings: st,
    serverTime: nowIso_()
  };
}

/* ------------------------------------------------------------------ question */

function topicIndex_() {
  var map = {};
  rows_(SHEETS.T).forEach(function (t) {
    map[String(t.TopicID).trim()] = {
      en: String(t.NameEN || ''),
      zh: String(t.NameZH || ''),
      strand: String(t.Strand || ''),
      level: String(t.Level || '')
    };
  });
  return map;
}

function sectionQuestions_(year, section) {
  var y = String(year).trim();
  var s = String(section).trim();
  return rows_(SHEETS.Q).filter(function (q) {
    return String(q.Year).trim() === y && String(q.Section).trim() === s;
  }).sort(function (a, b) {
    var d = Number(a.QNo) - Number(b.QNo);
    return d !== 0 ? d : String(a.Part).localeCompare(String(b.Part));
  });
}

function getSectionQuestions(year, section) {
  var rowsQ = sectionQuestions_(year, section);
  if (!rowsQ.length) throw new Error('題庫中找不到 ' + year + ' ' + section + ' 的資料。');
  var topics = topicIndex_();
  var groups = [];
  var byQ = {};
  var maxTotal = 0;
  var paperUrl = '';
  rowsQ.forEach(function (q) {
    var qno = String(q.QNo).trim();
    var part = String(q.Part || '').trim();
    var max = Number(q.MaxMark) || 0;
    maxTotal += max;
    if (q.PaperURL && !paperUrl) paperUrl = String(q.PaperURL);
    if (!byQ[qno]) {
      byQ[qno] = { qno: qno, maxMark: 0, parts: [] };
      groups.push(byQ[qno]);
    }
    byQ[qno].maxMark += max;
    var tid = String(q.TopicID || '').trim();
    byQ[qno].parts.push({
      qid: String(q.QID || qid_(q.Year, q.Section, qno, part)).trim(),
      part: part,
      maxMark: max,
      topicId: tid,
      topicZh: topics[tid] ? topics[tid].zh : tid,
      topicEn: topics[tid] ? topics[tid].en : tid
    });
  });
  return {
    year: String(year).trim(),
    section: String(section).trim(),
    maxTotal: maxTotal,
    sectionMaxExpected: SECTION_MAX,
    paperUrl: paperUrl,
    questions: groups
  };
}

/* -------------------------------------------------------------------- submit */

function submitAttempt(payload) {
  payload = payload || {};
  var studentId = resolveStudent_(payload.studentIdOverride);
  var attemptId = String(payload.attemptId || '').trim();
  if (!attemptId) throw new Error('缺少 AttemptID。');

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('系統繁忙，請稍後再試。');
  try {
    // idempotency: same AttemptID must never create a second record
    var existing = rows_(SHEETS.A).filter(function (a) {
      return String(a.AttemptID).trim() === attemptId;
    })[0];
    if (existing) {
      return {
        duplicate: true,
        attemptId: attemptId,
        totalScore: Number(existing.TotalScore),
        maxTotal: Number(existing.MaxTotal),
        percent: Number(existing.Percent),
        attemptNo: Number(existing.AttemptNo)
      };
    }

    var year = String(payload.year).trim();
    var section = String(payload.section).trim();
    if (SECTIONS.indexOf(section) < 0) throw new Error('Section 不正確。');

    var defs = sectionQuestions_(year, section);
    if (!defs.length) throw new Error('題庫中找不到 ' + year + ' ' + section + '。');

    var defMap = {};
    defs.forEach(function (q) {
      var id = String(q.QID || qid_(q.Year, q.Section, q.QNo, q.Part)).trim();
      defMap[id] = q;
    });

    var given = {};
    (payload.scores || []).forEach(function (s) {
      given[String(s.qid).trim()] = s.score;
    });

    // every part must be present, integer, within 0..MaxMark
    var total = 0;
    var maxTotal = 0;
    var scoreRows = [];
    var submitTime = nowIso_();
    var topics = topicIndex_();

    Object.keys(defMap).forEach(function (id) {
      var def = defMap[id];
      var max = Number(def.MaxMark) || 0;
      maxTotal += max;
      if (!(id in given)) throw new Error('未填分數：' + id);
      var v = Number(given[id]);
      if (isNaN(v) || v !== Math.floor(v)) throw new Error('分數必須為整數：' + id);
      if (v < 0 || v > max) throw new Error('分數超出範圍（0-' + max + '）：' + id);
      total += v;
      var tid = String(def.TopicID || '').trim();
      scoreRows.push([
        attemptId, id, year, section, Number(def.QNo), String(def.Part || ''),
        v, max, max ? v / max : 0, tid, submitTime
      ]);
    });

    Object.keys(given).forEach(function (id) {
      if (!(id in defMap)) throw new Error('題庫沒有此小題：' + id);
    });

    var prior = rows_(SHEETS.A).filter(function (a) {
      return String(a.StudentID).trim().toLowerCase() === studentId
        && String(a.Year).trim() === year
        && String(a.Section).trim() === section
        && String(a.Status).trim() === 'submitted';
    }).length;
    var attemptNo = prior + 1;
    var percent = maxTotal ? Math.round((total / maxTotal) * 1000) / 10 : 0;

    sheet_(SHEETS.A).appendRow([
      attemptId, studentId, year, section, attemptNo,
      String(payload.createdTime || submitTime), submitTime,
      total, maxTotal, percent, 'submitted', String(payload.note || '')
    ]);

    if (scoreRows.length) {
      var shS = sheet_(SHEETS.S);
      shS.getRange(shS.getLastRow() + 1, 1, scoreRows.length, scoreRows[0].length)
        .setValues(scoreRows);
    }
    SpreadsheetApp.flush();

    return {
      duplicate: false,
      attemptId: attemptId,
      attemptNo: attemptNo,
      totalScore: total,
      maxTotal: maxTotal,
      percent: percent,
      submitTime: submitTime
    };
  } finally {
    lock.releaseLock();
  }
}

/* ----------------------------------------------------------------- dashboard */

function getStudentDashboard(opts) {
  opts = opts || {};
  var sid;
  if (opts.targetStudent) {
    requireTeacher_();
    sid = String(opts.targetStudent).trim().toLowerCase();
  } else {
    sid = resolveStudent_(opts.studentIdOverride);
  }
  var st = settings_();
  var topics = topicIndex_();

  var attempts = rows_(SHEETS.A).filter(function (a) {
    return String(a.StudentID).trim().toLowerCase() === sid
      && String(a.Status).trim() === 'submitted';
  });
  var attemptIds = {};
  attempts.forEach(function (a) { attemptIds[String(a.AttemptID).trim()] = true; });

  var scores = rows_(SHEETS.S).filter(function (s) {
    return attemptIds[String(s.AttemptID).trim()];
  });

  // per year+section: latest / best / count / series
  var cells = {};
  attempts.sort(function (a, b) {
    return String(a.SubmitTime).localeCompare(String(b.SubmitTime));
  }).forEach(function (a) {
    var key = String(a.Year).trim() + '|' + String(a.Section).trim();
    if (!cells[key]) {
      cells[key] = {
        year: String(a.Year).trim(), section: String(a.Section).trim(),
        count: 0, latest: null, best: null, series: []
      };
    }
    var c = cells[key];
    var pt = {
      attemptNo: Number(a.AttemptNo),
      score: Number(a.TotalScore),
      maxTotal: Number(a.MaxTotal),
      percent: Number(a.Percent),
      submitTime: String(a.SubmitTime)
    };
    c.count++;
    c.series.push(pt);
    c.latest = pt;
    if (!c.best || pt.score > c.best.score) c.best = pt;
  });

  // topic aggregation (single tag per part, by consensus D9)
  var tAgg = {};
  scores.forEach(function (s) {
    var tid = String(s.TopicID || '').trim();
    if (!tid) return;
    if (!tAgg[tid]) tAgg[tid] = { topicId: tid, parts: 0, score: 0, max: 0 };
    tAgg[tid].parts++;
    tAgg[tid].score += Number(s.Score) || 0;
    tAgg[tid].max += Number(s.MaxMark) || 0;
  });
  var topicRows = Object.keys(tAgg).map(function (tid) {
    var t = tAgg[tid];
    var info = topics[tid] || {};
    return {
      topicId: tid,
      nameZh: info.zh || tid,
      nameEn: info.en || tid,
      strand: info.strand || '',
      level: info.level || '',
      parts: t.parts,
      score: t.score,
      max: t.max,
      rate: t.max ? t.score / t.max : 0,
      enoughSample: t.parts >= Number(st.MIN_PARTS_FOR_TOPIC)
    };
  }).sort(function (a, b) { return a.rate - b.rate; });

  var strandAgg = {};
  topicRows.forEach(function (t) {
    var k = t.strand || 'Unknown';
    if (!strandAgg[k]) strandAgg[k] = { strand: k, score: 0, max: 0, parts: 0 };
    strandAgg[k].score += t.score;
    strandAgg[k].max += t.max;
    strandAgg[k].parts += t.parts;
  });

  var recent = attempts.slice().sort(function (a, b) {
    return String(b.SubmitTime).localeCompare(String(a.SubmitTime));
  }).slice(0, 10).map(function (a) {
    return {
      attemptId: String(a.AttemptID),
      year: String(a.Year).trim(),
      section: String(a.Section).trim(),
      attemptNo: Number(a.AttemptNo),
      score: Number(a.TotalScore),
      maxTotal: Number(a.MaxTotal),
      percent: Number(a.Percent),
      submitTime: String(a.SubmitTime),
      note: String(a.Note || '')
    };
  });

  return {
    studentId: sid,
    cells: Object.keys(cells).map(function (k) { return cells[k]; }),
    topics: topicRows,
    strands: Object.keys(strandAgg).map(function (k) {
      var s = strandAgg[k];
      s.rate = s.max ? s.score / s.max : 0;
      return s;
    }).sort(function (a, b) { return a.rate - b.rate; }),
    recent: recent,
    weakThreshold: Number(st.WEAK_TOPIC_THRESHOLD),
    minParts: Number(st.MIN_PARTS_FOR_TOPIC)
  };
}

function getAttemptDetail(attemptId) {
  var id = String(attemptId || '').trim();
  var head = rows_(SHEETS.A).filter(function (a) {
    return String(a.AttemptID).trim() === id;
  })[0];
  if (!head) throw new Error('找不到此記錄。');
  var u = currentUser_();
  var owner = String(head.StudentID).trim().toLowerCase();
  if (u.role !== 'teacher') {
    var me = resolveStudent_();
    if (me !== owner) throw new Error('無權查看其他學生的記錄。');
  }
  var topics = topicIndex_();
  var parts = rows_(SHEETS.S).filter(function (s) {
    return String(s.AttemptID).trim() === id;
  }).map(function (s) {
    var tid = String(s.TopicID || '').trim();
    return {
      qid: String(s.QID),
      qno: Number(s.QNo),
      part: String(s.Part || ''),
      score: Number(s.Score),
      maxMark: Number(s.MaxMark),
      rate: Number(s.Rate),
      topicId: tid,
      topicZh: topics[tid] ? topics[tid].zh : tid,
      topicEn: topics[tid] ? topics[tid].en : tid
    };
  }).sort(function (a, b) {
    return a.qno - b.qno || String(a.part).localeCompare(String(b.part));
  });
  return {
    attemptId: id,
    studentId: owner,
    year: String(head.Year).trim(),
    section: String(head.Section).trim(),
    attemptNo: Number(head.AttemptNo),
    totalScore: Number(head.TotalScore),
    maxTotal: Number(head.MaxTotal),
    percent: Number(head.Percent),
    submitTime: String(head.SubmitTime),
    note: String(head.Note || ''),
    parts: parts
  };
}

/* --------------------------------------------------------- teacher overview */

function getTeacherOverview(filter) {
  requireTeacher_();
  filter = filter || {};
  var fKlass = String(filter.klass || '').trim();
  var fYear = String(filter.year || '').trim();
  var fSection = String(filter.section || '').trim();
  var st = settings_();
  var topics = topicIndex_();

  var students = activeStudents_().filter(function (s) {
    return s.role === 'student' && (!fKlass || s.klass === fKlass);
  });
  var allow = {};
  students.forEach(function (s) { allow[s.email.toLowerCase()] = s; });

  var attempts = rows_(SHEETS.A).filter(function (a) {
    if (String(a.Status).trim() !== 'submitted') return false;
    if (!allow[String(a.StudentID).trim().toLowerCase()]) return false;
    if (fYear && String(a.Year).trim() !== fYear) return false;
    if (fSection && String(a.Section).trim() !== fSection) return false;
    return true;
  });

  var keep = {};
  attempts.forEach(function (a) { keep[String(a.AttemptID).trim()] = true; });
  var scores = rows_(SHEETS.S).filter(function (s) { return keep[String(s.AttemptID).trim()]; });

  var perStudent = {};
  students.forEach(function (s) {
    perStudent[s.email.toLowerCase()] = {
      email: s.email, displayName: s.displayName, klass: s.klass,
      attempts: 0, sections: {}, lastActivity: '', score: 0, max: 0,
      latest: null, best: null
    };
  });
  attempts.sort(function (a, b) {
    return String(a.SubmitTime).localeCompare(String(b.SubmitTime));
  }).forEach(function (a) {
    var r = perStudent[String(a.StudentID).trim().toLowerCase()];
    if (!r) return;
    r.attempts++;
    r.sections[String(a.Year).trim() + '|' + String(a.Section).trim()] = true;
    r.lastActivity = String(a.SubmitTime);
    r.score += Number(a.TotalScore) || 0;
    r.max += Number(a.MaxTotal) || 0;
    r.latest = { score: Number(a.TotalScore), percent: Number(a.Percent) };
    if (!r.best || Number(a.TotalScore) > r.best.score) {
      r.best = { score: Number(a.TotalScore), percent: Number(a.Percent) };
    }
  });

  var studentTopic = {};
  scores.forEach(function (s) {
    var id = String(s.AttemptID).trim();
    var owner = attempts.filter(function (a) { return String(a.AttemptID).trim() === id; })[0];
    if (!owner) return;
    var key = String(owner.StudentID).trim().toLowerCase();
    var tid = String(s.TopicID || '').trim();
    if (!studentTopic[key]) studentTopic[key] = {};
    if (!studentTopic[key][tid]) studentTopic[key][tid] = { score: 0, max: 0, parts: 0 };
    studentTopic[key][tid].score += Number(s.Score) || 0;
    studentTopic[key][tid].max += Number(s.MaxMark) || 0;
    studentTopic[key][tid].parts++;
  });

  var studentRows = Object.keys(perStudent).map(function (k) {
    var r = perStudent[k];
    var weak = '';
    var weakRate = 2;
    var tmap = studentTopic[k] || {};
    Object.keys(tmap).forEach(function (tid) {
      var t = tmap[tid];
      if (t.parts < Number(st.MIN_PARTS_FOR_TOPIC) || !t.max) return;
      var rate = t.score / t.max;
      if (rate < weakRate) { weakRate = rate; weak = tid; }
    });
    return {
      email: r.email,
      displayName: r.displayName,
      klass: r.klass,
      attempts: r.attempts,
      sectionsDone: Object.keys(r.sections).length,
      lastActivity: r.lastActivity,
      avgRate: r.max ? r.score / r.max : 0,
      latestScore: r.latest ? r.latest.score : null,
      bestScore: r.best ? r.best.score : null,
      weakTopicId: weak,
      weakTopicZh: weak && topics[weak] ? topics[weak].zh : weak,
      weakTopicEn: weak && topics[weak] ? topics[weak].en : weak,
      weakRate: weak ? weakRate : null
    };
  }).sort(function (a, b) {
    return String(b.lastActivity).localeCompare(String(a.lastActivity));
  });

  var classTopic = {};
  scores.forEach(function (s) {
    var tid = String(s.TopicID || '').trim();
    if (!tid) return;
    if (!classTopic[tid]) classTopic[tid] = { topicId: tid, score: 0, max: 0, parts: 0 };
    classTopic[tid].score += Number(s.Score) || 0;
    classTopic[tid].max += Number(s.MaxMark) || 0;
    classTopic[tid].parts++;
  });
  var weakest = Object.keys(classTopic).map(function (tid) {
    var t = classTopic[tid];
    var info = topics[tid] || {};
    return {
      topicId: tid, nameZh: info.zh || tid, nameEn: info.en || tid,
      parts: t.parts, rate: t.max ? t.score / t.max : 0
    };
  }).filter(function (t) {
    return t.parts >= Number(st.MIN_PARTS_FOR_TOPIC);
  }).sort(function (a, b) { return a.rate - b.rate; }).slice(0, 5);

  var perSection = {};
  attempts.forEach(function (a) {
    var s = String(a.Section).trim();
    if (!perSection[s]) perSection[s] = { section: s, score: 0, max: 0, count: 0 };
    perSection[s].score += Number(a.TotalScore) || 0;
    perSection[s].max += Number(a.MaxTotal) || 0;
    perSection[s].count++;
  });

  var totScore = 0, totMax = 0;
  attempts.forEach(function (a) {
    totScore += Number(a.TotalScore) || 0;
    totMax += Number(a.MaxTotal) || 0;
  });

  return {
    classes: activeStudents_().map(function (s) { return s.klass; })
      .filter(function (v, i, arr) { return v && arr.indexOf(v) === i; }).sort(),
    students: studentRows,
    summary: {
      activeStudents: studentRows.filter(function (s) { return s.attempts > 0; }).length,
      totalStudents: studentRows.length,
      submissions: attempts.length,
      avgRate: totMax ? totScore / totMax : 0,
      perSection: SECTIONS.map(function (s) {
        var r = perSection[s];
        return {
          section: s,
          count: r ? r.count : 0,
          avgRate: r && r.max ? r.score / r.max : 0
        };
      })
    },
    weakestTopics: weakest,
    filter: { klass: fKlass, year: fYear, section: fSection }
  };
}

/* --------------------------------------------------------------- maintenance */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DSE Trainer')
    .addItem('初始化工作表 initSheets', 'initSheets')
    .addItem('檢查題庫 validateQuestionBank', 'showValidation')
    .addToUi();
}

function initSheets() {
  var s = ss_();
  Object.keys(HEADERS).forEach(function (name) {
    var sh = s.getSheetByName(name) || s.insertSheet(name);
    var head = HEADERS[name];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  var setSh = s.getSheetByName(SHEETS.SET);
  if (setSh.getLastRow() < 2) {
    var rows = Object.keys(DEFAULT_SETTINGS).map(function (k) {
      return [k, DEFAULT_SETTINGS[k], ''];
    });
    setSh.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  applyTopicValidation();
  return '完成：已建立／更新 6 張工作表。';
}

/** dropdown for Questions.TopicID sourced from the Topics sheet */
function applyTopicValidation() {
  var qSh = sheet_(SHEETS.Q);
  var tSh = sheet_(SHEETS.T);
  var last = Math.max(tSh.getLastRow(), 2);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(tSh.getRange(2, 1, last - 1, 1), true)
    .setAllowInvalidData(false)
    .build();
  qSh.getRange(2, 6, Math.max(qSh.getMaxRows() - 1, 1), 1).setDataValidation(rule);

  var secRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SECTIONS, true)
    .setAllowInvalidData(false)
    .build();
  qSh.getRange(2, 2, Math.max(qSh.getMaxRows() - 1, 1), 1).setDataValidation(secRule);
}

function validateQuestionBank() {
  var issues = [];
  var qs = rows_(SHEETS.Q);
  var topicIds = {};
  rows_(SHEETS.T).forEach(function (t) { topicIds[String(t.TopicID).trim()] = true; });

  var totals = {};
  var seen = {};
  qs.forEach(function (q, i) {
    var line = i + 2;
    var year = String(q.Year).trim();
    var section = String(q.Section).trim();
    var qno = String(q.QNo).trim();
    var part = String(q.Part || '').trim();
    var max = Number(q.MaxMark);
    var tid = String(q.TopicID || '').trim();
    var id = String(q.QID || '').trim() || qid_(year, section, qno, part);

    if (!year) issues.push('第 ' + line + ' 行：缺 Year');
    if (SECTIONS.indexOf(section) < 0) issues.push('第 ' + line + ' 行：Section 無效（' + section + '）');
    if (!qno) issues.push('第 ' + line + ' 行：缺 QNo');
    if (!max || max <= 0 || max !== Math.floor(max)) {
      issues.push('第 ' + line + ' 行：MaxMark 必須為正整數（' + q.MaxMark + '）');
    }
    if (!tid) issues.push('第 ' + line + ' 行：缺 TopicID');
    else if (!topicIds[tid]) issues.push('第 ' + line + ' 行：TopicID 不存在（' + tid + '）');

    var expected = qid_(year, section, qno, part);
    if (String(q.QID || '').trim() && String(q.QID).trim() !== expected) {
      issues.push('第 ' + line + ' 行：QID 應為 ' + expected + '，現為 ' + q.QID);
    }
    if (seen[id]) issues.push('QID 重複：' + id);
    seen[id] = true;

    var key = year + ' ' + section;
    totals[key] = (totals[key] || 0) + (Number(max) || 0);
  });

  Object.keys(totals).sort().forEach(function (k) {
    if (totals[k] !== SECTION_MAX) {
      issues.push(k + ' 總分為 ' + totals[k] + '，應為 ' + SECTION_MAX);
    }
  });

  return {
    ok: issues.length === 0,
    sections: Object.keys(totals).sort().map(function (k) {
      return { key: k, total: totals[k] };
    }),
    issues: issues
  };
}

function showValidation() {
  var r = validateQuestionBank();
  var msg = r.ok
    ? '題庫檢查通過。\n\n' + r.sections.map(function (s) {
        return s.key + '：' + s.total + '/35';
      }).join('\n')
    : '發現 ' + r.issues.length + ' 個問題：\n\n' + r.issues.slice(0, 30).join('\n');
  SpreadsheetApp.getUi().alert(msg);
}
