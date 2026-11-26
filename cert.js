/* 한우리 안산단원지부 수료증 만들기 — 공통 로직 (브라우저·node 함께 사용) */
(function (root) {
  "use strict";

  // ---------- 글자 정리 ----------
  function clean(v) {
    if (v === null || v === undefined) return "";
    return String(v).normalize("NFKC").replace(/\s+/g, " ").trim();
  }

  // ---------- 학교 ----------
  var SCHOOL_FIX = { "세이트폴": "세인트폴" };
  function normSchool(v) {
    var s = clean(v);
    if (!s) return "";
    if (SCHOOL_FIX[s]) s = SCHOOL_FIX[s];
    if (/[^학]초$/.test(s) || s === "초") return s + "등학교";
    if (/초등$/.test(s)) return s + "학교";
    if (/[^학]중$/.test(s)) return s + "학교";
    if (/[^학]고$/.test(s)) return s + "등학교";
    return s;
  }

  // ---------- 학년 ----------
  function normGrade(v) {
    var g = clean(v).replace(/\s/g, "");
    if (!g) return "";
    var m;
    if ((m = g.match(/^(?:초|중|고)?(\d{1,2})$/))) return m[1] + "학년";
    if ((m = g.match(/^(?:초|중|고)(\d{1,2})학년$/))) return m[1] + "학년";
    if ((m = g.match(/^(\d{1,2})살$/))) return m[1] + "세";
    return g;
  }

  // ---------- 과정(단계) ----------
  var COURSE_RULES = [
    [/^(?:꿈꾸는)?새싹([12])$/, "꿈꾸는 새싹$1"],
    [/^어휘(?:리더)?([12])$/, "어휘리더$1"],
    [/^독(?:서)?(?:리더)?([12])$/, "독서리더$1"],
    [/^논(?:리)?(?:리더)?([12])$/, "논리리더$1"],
    [/^과탐$|^과학탐탐$/, "과학탐탐"],
    [/^세계문화[,·.]?지리$/, "세계문화지리"],
    [/^소(?:크|클|크라테스)?([123])$/, "소크라테스$1"],
    [/^(?:몰독|몰입|몰입독서)([PABCDSpabcds])$/, function (m, a) { return "몰입독서" + a.toUpperCase(); }],
    [/^초[·.,ㆍ・]?중등(?:논술|수료|논술수료)?$/, "초·중등 논술"],
    [/^중등(?:논술|수료|논술수료)$/, "중등 논술"],
  ];
  function normCourse(v) {
    var c = clean(v).replace(/\s*과정$/, "");
    if (!c) return "";
    var key = c.replace(/\s/g, "");
    for (var i = 0; i < COURSE_RULES.length; i++) {
      var r = COURSE_RULES[i];
      if (r[0].test(key)) return key.replace(r[0], r[1]);
    }
    return c;
  }
  function isKnownCourse(c) {
    var key = clean(c).replace(/\s*과정$/, "").replace(/\s/g, "");
    return COURSE_RULES.some(function (r) { return r[0].test(key); });
  }

  // ---------- 엑셀 표 → 수료증 목록 ----------
  var HEAD = {
    name: /^(이름|성명|학생|학생명)$/,
    school: /^(학교|학교명|소속)$/,
    grade: /^(학년|나이)$/,
    course: /^(수료단계|단계|과정|수료과정|과정명)$/,
    note: /^(비고|메모)$/,
  };
  function findHeader(rows) {
    for (var r = 0; r < Math.min(rows.length, 15); r++) {
      var map = {};
      (rows[r] || []).forEach(function (cell, ci) {
        var t = clean(cell).replace(/\s/g, "");
        Object.keys(HEAD).forEach(function (k) {
          if (map[k] === undefined && HEAD[k].test(t)) map[k] = ci;
        });
      });
      if (map.name !== undefined && map.course !== undefined) return { row: r, map: map };
    }
    return null;
  }

  // rows: 2차원 배열 (SheetJS sheet_to_json header:1)
  function rowsToCerts(rows, sheetName) {
    var h = findHeader(rows);
    if (!h) return null;
    var out = [];
    for (var r = h.row + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var get = function (k) { return h.map[k] === undefined ? "" : row[h.map[k]]; };
      var name = clean(get("name")).replace(/\s/g, "");
      var rawCourse = clean(get("course"));
      if (!name && !rawCourse) continue;
      var main = normCourse(rawCourse);
      var school = normSchool(get("school"));
      var grade = normGrade(get("grade"));
      var note = clean(get("note"));
      var courses = [];
      if (note) {
        var parts = note.split(/[+,/;]|그리고|및/).map(function (s) { return s.trim(); }).filter(Boolean);
        if (parts.length && parts.every(function (p) { return isKnownCourse(p) || normCourse(p) === main; })) {
          courses = parts.map(normCourse);
        }
      }
      if (main && courses.indexOf(main) < 0) courses.push(main);
      if (!courses.length) courses = [""];
      var seen = {};
      courses.forEach(function (c) {
        if (seen[c]) return; seen[c] = 1;
        out.push({ name: name, school: school, grade: grade, course: c, sheet: sheetName || "", src: r + 1 });
      });
    }
    return out;
  }

  // ---------- 확인할 점 ----------
  var SURNAMES = "김이박최정강조윤장임한오서신권황안송류유전홍고문양손배백허남심노하곽성차주우구민진나지엄채원천방공현함변염여추도소석선설마길연위표명기반왕금옥육인맹제모탁국어은편용예봉경사부가복태목형피두감음빈동온호범좌팽승간상갈옹견단담화계십묵".split("");
  function checkCerts(list) {
    var key = {}, byName = {};
    list.forEach(function (c) {
      var k = c.name + "|" + c.course;
      key[k] = (key[k] || 0) + 1;
      (byName[c.name + "|" + c.school] = byName[c.name + "|" + c.school] || []).push(c);
    });
    list.forEach(function (c) {
      var w = [];
      if (!c.name) w.push({ lv: "err", t: "이름이 비어 있어요" });
      if (!c.course) w.push({ lv: "err", t: "수료 단계가 비어 있어요" });
      if (c.name && (c.name.length < 2 || (c.name.length === 2 && SURNAMES.indexOf(c.name[0]) < 0)))
        w.push({ lv: "warn", t: "성이 빠졌는지 확인해 주세요" });
      if (c.name && c.course && key[c.name + "|" + c.course] > 1) w.push({ lv: "warn", t: "같은 이름·단계가 두 번 있어요" });
      if (!c.school) w.push({ lv: "info", t: "학교가 비어 있어요" });
      var other = (byName[c.name + "|" + c.school] || []).filter(function (x) { return x.course !== c.course && (x.src !== c.src || x.sheet !== c.sheet); });
      if (c.name && other.length)
        w.push({ lv: "info", t: "같은 학생이 다른 줄에도 있어요 (" + other.map(function (x) { return x.course; }).join(", ") + ")" });
      c.warn = w;
    });
    return list;
  }

  // ---------- PDF 그리기 ----------
  var W = 593, H = 840, L = 72.8, R = 502.9, BODY_R = 496.3;

  function makeDrawer(PDFLib, pdf, fonts, wmImage, stampImage) {
    var rgb = PDFLib.rgb(0, 0, 0);
    var fs = fonts.body, fb = fonts.title;
    function tw(t, size, f) { return (f || fs).widthOfTextAtSize(t, size); }
    function bw(t, size) {
      var n1 = (t.match(/「/g) || []).length, n2 = (t.match(/」/g) || []).length;
      return tw(t, size) - 0.3 * size * n1 - 0.5 * size * n2;
    }
    return function draw(c, year) {
      var p = pdf.addPage([W, H]);
      function put(x, yTop, t, size, f) { if (t) p.drawText(t, { x: x, y: H - yTop, size: size, font: f || fs, color: rgb }); }
      function bput(x, yTop, t, size) {
        var buf = "";
        for (var i = 0; i < t.length; i++) {
          var ch = t[i];
          if (ch === "「" || ch === "」") {
            if (buf) { put(x, yTop, buf, size); x += tw(buf, size); buf = ""; }
            if (ch === "「") { put(x - 0.3 * size, yTop, ch, size); x += 0.7 * size; }
            else { put(x, yTop, ch, size); x += 0.5 * size; }
          } else buf += ch;
        }
        if (buf) put(x, yTop, buf, size);
      }
      p.drawImage(wmImage, { x: 188.6, y: H - 506.0, width: 199.5, height: 186.3 });
      var t = "수 료 증";
      put((W - tw(t, 40, fb)) / 2, 116, t, 40, fb);
      put(102.4, 231, c.course + " 과정", 18);
      var line = (c.school + " " + c.grade).trim();
      put(R - tw(line, 18), 288, line, 18);
      var nm = c.name.split("").join(" ");
      put(R - tw(nm, 18), 316.5, nm, 18);

      var body = "위 학생은 (주)한우리열린교육에서 제공하는 한우리독서토론논술의「" + c.course + "」 과정을 수료하였으므로 이 수료증을 드립니다.";
      var size = 20.5, maxw = BODY_R - L, sp = tw(" ", size);
      var words = body.split(" "), lines = [], cur = [];
      words.forEach(function (w) {
        var trial = cur.concat([w]).join(" ");
        if (cur.length && bw(trial, size) > maxw - 10) { lines.push(cur); cur = [w]; } else cur.push(w);
      });
      lines.push(cur);
      var y = 397;
      lines.forEach(function (ws, i) {
        var x = L + (i === 0 ? 10 : 0), avail = maxw - (i === 0 ? 10 : 0);
        var content = ws.reduce(function (a, w) { return a + bw(w, size); }, 0);
        var gap = sp;
        if (i < lines.length - 1 && ws.length > 1) {
          var g = (avail - content) / (ws.length - 1);
          if (g < sp * 2.2) gap = g;
        }
        ws.forEach(function (w) { bput(x, y, w, size); x += bw(w, size) + gap; });
        y += 42;
      });
      [[year + "년 12월 1일", 607], ["한우리독서토론논술 안산단원지부", 656.5], ["지 부 장 이 미 진", 706]].forEach(function (a) {
        put((W - tw(a[0], 19)) / 2, a[1], a[0], 19);
      });
      if (stampImage) {
        // 도장: 이름 끝에 살짝 걸치게, 장마다 각도·진하기를 조금씩 다르게(찍은 느낌)
        var seed = 0, key = c.name + c.course;
        for (var k = 0; k < key.length; k++) seed = (seed * 31 + key.charCodeAt(k)) >>> 0;
        var rnd = function () { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; };
        var sz = 50 + rnd() * 3, ang = (rnd() - 0.5) * 7, op = 0.86 + rnd() * 0.1;
        var endX = (W + tw("지 부 장 이 미 진", 19)) / 2;
        var cx = endX + 18 + (rnd() - 0.5) * 4, cy = 706 - 7 + (rnd() - 0.5) * 4;
        var rad = ang * Math.PI / 180;
        // 가운데 기준으로 돌리기 위해 왼쪽 아래 꼭짓점 계산
        var hx = sz / 2, x0 = cx - (hx * Math.cos(rad) - hx * Math.sin(rad)), y0 = (H - cy) - (hx * Math.sin(rad) + hx * Math.cos(rad));
        p.drawImage(stampImage, { x: x0, y: y0, width: sz, height: sz, rotate: PDFLib.degrees(ang), opacity: op });
      }
    };
  }

  // assets: {body: bytes, title: bytes, wm: bytes}
  async function buildPdf(PDFLib, fontkit, assets, certs, year, subset) {
    var SUBSET = subset === true; // fontkit 부분 글꼴은 한글이 깨져서 쓰지 않음
    var pdf = await PDFLib.PDFDocument.create();
    pdf.registerFontkit(fontkit);
    pdf.setTitle(year + " 수료증");
    var fonts = {
      body: await pdf.embedFont(assets.body, { subset: SUBSET }),
      title: await pdf.embedFont(assets.title, { subset: SUBSET }),
    };
    var wm = await pdf.embedPng(assets.wm);
    var st = assets.stamp ? await pdf.embedPng(assets.stamp) : null;
    var draw = makeDrawer(PDFLib, pdf, fonts, wm, st);
    certs.forEach(function (c) { draw(c, year); });
    return await pdf.save();
  }

  function certYear(d) {
    d = d || new Date();
    return d.getMonth() <= 1 ? d.getFullYear() - 1 : d.getFullYear(); // 1~2월에 만들면 지난해 12월 1일
  }

  var SUBSET = true;
  var api = { clean: clean, normSchool: normSchool, normGrade: normGrade, normCourse: normCourse,
    rowsToCerts: rowsToCerts, checkCerts: checkCerts, buildPdf: buildPdf, certYear: certYear };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Cert = api;
})(this);
