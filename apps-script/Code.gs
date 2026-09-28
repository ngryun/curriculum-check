/** @OnlyCurrentDoc */
// 졸업이수요건 점검 프로그램 — 학생 기초조사 템플릿 코드 (이 시트의 사본을 만들어 쓰세요)
// 사용법: ① 시트 위 메뉴 [📋 기초조사 → 설정 붙여넣기]에 프로그램에서 복사한 설정 코드를 넣기
//        ② 저장한 창에서 [학생용 주소 만들기] (Apps Script API가 켜져 있어야 함: https://script.google.com/home/usersettings)
//        ③ 나온 주소를 프로그램에 붙여넣기. 자동 배포가 안 되면 [배포 → 새 배포 → 웹 앱] 실행: 나 · 액세스: 모든 사용자로 직접 배포
// 매니페스트(appsscript.json)는 apps-script/appsscript.json과 같아야 합니다 (웹 앱 설정과 권한 목록).
// 이 코드는 이 구글 시트 하나에만 접근하며(@OnlyCurrentDoc), 학생이 낸 내용은 이 시트에만 저장됩니다.

var BUNDLED_SURVEY = null;
var AUTO_DEPLOY = true;
var PAGE_HTML = "<!DOCTYPE html><html lang=\"ko\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"><title>기초조사</title><style>\n*{box-sizing:border-box}\nhtml,body{margin:0;padding:0;background:#F6F4EF;color:#1C2333;font-family:-apple-system,\"Apple SD Gothic Neo\",\"Malgun Gothic\",\"Noto Sans KR\",sans-serif;font-size:16px;line-height:1.5;-webkit-text-size-adjust:100%}\n#app{max-width:560px;margin:0 auto;padding:14px 14px 90px}\n.pv{background:#1C2333;color:#fff;font-size:13px;text-align:center;padding:6px 10px;border-radius:8px;margin-bottom:10px}\nh1{font-size:21px;margin:6px 0 4px;line-height:1.35}\n.sub{color:#6B7280;font-size:14px;margin-bottom:12px}\n.card{background:#fff;border:1px solid #DDD8CC;border-radius:14px;padding:14px;margin-bottom:12px}\n.card h2{font-size:17px;margin:0 0 10px}\n.card h2 small{font-weight:400;color:#6B7280;font-size:14px}\n.note{background:#EAF0F6;color:#2C5A8A;border-radius:10px;padding:10px 12px;font-size:14px;margin-bottom:12px;white-space:pre-wrap}\n.warn{background:#F7E9E3;color:#A2452C;border-radius:10px;padding:10px 12px;font-size:14px;margin:10px 0}\n.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(64px,1fr));gap:8px}\n.pill{border:1.5px solid #DDD8CC;background:#fff;border-radius:10px;padding:11px 0;font-size:16px;font-weight:700;color:#1C2333;text-align:center;cursor:pointer;-webkit-tap-highlight-color:transparent}\n.pill.on{background:#2C5A8A;border-color:#2C5A8A;color:#fff}\n.pill.skip{background:#EEEAE2;border-style:dashed;color:#B8B4A9;text-decoration:line-through;cursor:default}\ninput[type=text],textarea{width:100%;font-size:17px;padding:12px;border:1.5px solid #DDD8CC;border-radius:10px;font-family:inherit;background:#fff;color:#1C2333}\ntextarea{min-height:84px;resize:vertical;font-size:15px}\n.agree{display:flex;gap:10px;align-items:flex-start;font-size:14px;color:#1C2333;margin-top:10px;cursor:pointer}\n.agree input{width:22px;height:22px;flex-shrink:0;margin:0}\n.privacy{font-size:13.5px;color:#6B7280}\n.btn{display:block;width:100%;border:none;border-radius:12px;padding:15px;font-size:17px;font-weight:800;cursor:pointer;background:#2C5A8A;color:#fff;font-family:inherit}\n.btn.gray{background:#fff;color:#1C2333;border:1.5px solid #DDD8CC}\n.btn:disabled{background:#B8B4A9;color:#fff}\n.row{display:flex;gap:8px}\n.row .btn{flex:1}\n.bar{position:sticky;top:0;z-index:5;background:#F6F4EF;padding:8px 0 10px;margin-bottom:4px}\n.bar .who{font-weight:800;font-size:15px}\n.bar .prog{font-size:13.5px;color:#6B7280}\n.sem h2{display:flex;justify-content:space-between;align-items:center;gap:8px}\n.badge{font-size:12.5px;font-weight:800;border-radius:999px;padding:3px 10px;white-space:nowrap;background:#F7E9E3;color:#A2452C}\n.badge.ok{background:#E7F1EA;color:#2F6D4F}\n.away{display:flex;gap:10px;align-items:center;font-size:14px;color:#6B7280;background:#F6F4EF;border-radius:10px;padding:9px 10px;margin-bottom:10px;cursor:pointer}\n.away input{width:20px;height:20px;margin:0;flex-shrink:0}\n.grp{margin-top:12px}\n.grp-h{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:8px}\n.grp-t{font-weight:800;font-size:15.5px}\n.grp-t small{font-weight:400;color:#6B7280;font-size:13.5px}\n.cnt{font-size:14px;font-weight:800;border-radius:999px;padding:2px 10px;background:#F7E9E3;color:#A2452C;white-space:nowrap}\n.cnt.ok{background:#E7F1EA;color:#2F6D4F}\n.opt{display:flex;align-items:center;gap:10px;width:100%;text-align:left;border:1.5px solid #DDD8CC;background:#fff;border-radius:10px;padding:11px 12px;margin-bottom:7px;font-size:16px;color:#1C2333;cursor:pointer;font-family:inherit;-webkit-tap-highlight-color:transparent}\n.opt .box{width:22px;height:22px;border:2px solid #B8B4A9;border-radius:6px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:900;color:#fff}\n.opt .nm{flex:1}\n.opt .meta{font-size:12.5px;color:#6B7280;white-space:nowrap}\n.opt.on{border-color:#2C5A8A;background:#EAF0F6}\n.opt.on .box{background:#2C5A8A;border-color:#2C5A8A}\n.opt.dim{opacity:.45}\n.sumg{margin:6px 0 10px}\n.sumg b{display:block;font-size:14px;color:#6B7280;font-weight:600}\n.sumg span{font-size:15.5px}\n.done{text-align:center;padding:26px 14px}\n.done .big{font-size:52px;line-height:1;color:#2F6D4F}\n.toast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);max-width:92%;background:#1C2333;color:#fff;font-size:14.5px;padding:11px 16px;border-radius:12px;z-index:20;box-shadow:0 6px 20px rgba(0,0,0,.25)}\n.err{color:#A2452C;font-size:14.5px;margin:8px 0}\n</style></head><body><div id=\"app\"></div><script>\n(function () {\n  \"use strict\";\n  var BOOT = /*__BOOT__*/null;\n  var root = document.getElementById(\"app\");\n  if (!BOOT || !BOOT.survey) { root.textContent = \"조사 정보를 불러오지 못했어요. 담임 선생님께 알려 주세요.\"; return; }\n  var S = BOOT.survey;\n  var STATUS_TEXT = { done: \"들은 과목\", now: \"지금 듣는 과목\", next: \"신청한 과목\" };\n  var st = { step: \"info\", cls: null, num: null, name: \"\", agree: false, picks: {}, away: {}, memo: \"\", sending: false, result: null, error: \"\" };\n  S.classes.forEach(function (k) { if (k.c === BOOT.ban) st.cls = k.c; });\n\n  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }\n  function add(parent, child) { parent.appendChild(child); return child; }\n  function norm(s) { return String(s == null ? \"\" : s).replace(/\\s+/g, \" \").replace(/^\\s+|\\s+$/g, \"\"); }\n  function key(s) { return String(s || \"\").replace(/[\\s·‧•・･ᐧ⋅∙ㆍ]/g, \"\").toLowerCase(); }\n  function classInfo(c) { for (var i = 0; i < S.classes.length; i++) if (S.classes[i].c === c) return S.classes[i]; return null; }\n  function nameProblem(n) {\n    if (!n) return \"이름을 적어 주세요.\";\n    if (n.length > 20) return \"이름은 20자까지만 적을 수 있어요.\";\n    if (/[0-9]/.test(n)) return \"번호는 빼고 이름만 적어 주세요.\";\n    if (/^[=+\\-@'\"]/.test(n) || /[<>\\u0000-\\u001f\\u007f]/.test(n)) return \"이름에 쓸 수 없는 글자가 있어요.\";\n    return \"\";\n  }\n  var toastTimer = null;\n  function toast(msg) {\n    var old = document.getElementById(\"toast\");\n    if (old) old.parentNode.removeChild(old);\n    var t = el(\"div\", \"toast\", msg);\n    t.id = \"toast\";\n    document.body.appendChild(t);\n    clearTimeout(toastTimer);\n    toastTimer = setTimeout(function () { if (t.parentNode) t.parentNode.removeChild(t); }, 2600);\n  }\n  function go(step) { st.step = step; render(); try { window.scrollTo(0, 0); } catch (e) {} }\n  function header(parent) {\n    if (BOOT.preview) add(parent, el(\"div\", \"pv\", \"미리보기 화면입니다 — 여기서 누르는 [제출하기]는 실제로 제출되지 않아요.\"));\n    add(parent, el(\"h1\", null, S.title));\n    add(parent, el(\"div\", \"sub\", S.grade + \"학년 학생이 직접 적는 선택과목 조사예요.\"));\n  }\n\n  // ---- selections ----\n  function picked(gid) { return st.picks[gid] || []; }\n  function activeSems() { return S.semesters.filter(function (s) { return !st.away[s.code]; }); }\n  function semPicks(sem) {\n    var out = [];\n    sem.groups.forEach(function (g) { picked(g.id).forEach(function (n) { out.push({ name: n, gid: g.id }); }); });\n    return out;\n  }\n  function blockReason(sem, g, name) {\n    var mine = picked(g.id), i, j;\n    if (mine.indexOf(name) >= 0) return \"\";\n    var list = semPicks(sem);\n    for (i = 0; i < list.length; i++) if (key(list[i].name) === key(name)) return \"같은 학기의 다른 묶음에서 이미 골랐어요.\";\n    for (i = 0; i < S.semesters.length; i++) {\n      var o = S.semesters[i];\n      if (o.code === sem.code || o.code.charAt(0) !== sem.code.charAt(0) || st.away[o.code]) continue;\n      var ol = semPicks(o);\n      for (j = 0; j < ol.length; j++) if (key(ol[j].name) === key(name)) return o.label + \"에 이미 골랐어요. 같은 과목은 한 번만 들을 수 있어요.\";\n    }\n    if (g.n != null && mine.length >= g.n) return \"full\";\n    return \"\";\n  }\n  function groupDone(g) { return g.n == null || picked(g.id).length === g.n; }\n  function semDone(sem) { return !!st.away[sem.code] || sem.groups.every(groupDone); }\n  function problems() {\n    var out = [];\n    if (activeSems().length === 0) out.push({ text: \"모든 학기를 '다른 학교'로 표시할 수는 없어요.\" });\n    S.semesters.forEach(function (sem) {\n      if (st.away[sem.code]) return;\n      sem.groups.forEach(function (g) {\n        if (!groupDone(g)) out.push({ gid: g.id, text: sem.label + \" \" + g.label + \": \" + g.n + \"개 중 \" + picked(g.id).length + \"개 골랐어요.\" });\n      });\n    });\n    return out;\n  }\n\n  // ---- step 1: 반 · 번호 · 이름 ----\n  function renderInfo() {\n    header(root);\n    if (S.deadline) add(root, el(\"div\", \"note\", \"마감: \" + S.deadline));\n    if (S.notice) add(root, el(\"div\", \"note\", S.notice));\n\n    var c1 = add(root, el(\"div\", \"card\"));\n    add(c1, el(\"h2\", null, \"① 우리 반을 눌러 주세요\"));\n    var g1 = add(c1, el(\"div\", \"grid\"));\n    var c2 = add(root, el(\"div\", \"card\"));\n    var c3 = add(root, el(\"div\", \"card\"));\n    function drawNumbers() {\n      while (c2.firstChild) c2.removeChild(c2.firstChild);\n      add(c2, el(\"h2\", null, \"② 내 번호를 눌러 주세요\"));\n      var info = classInfo(st.cls);\n      if (!info) { add(c2, el(\"div\", \"sub\", \"먼저 위에서 반을 골라 주세요.\")); return; }\n      var skips = info.skip || [];\n      if (st.num != null && (st.num > info.size || skips.indexOf(st.num) >= 0)) st.num = null;\n      if (skips.length) add(c2, el(\"div\", \"sub\", \"흐리게 보이는 번호는 결번이라 누를 수 없어요.\"));\n      var g2 = add(c2, el(\"div\", \"grid\"));\n      for (var n = 1; n <= info.size; n++) (function (n) {\n        if (skips.indexOf(n) >= 0) {\n          var s = add(g2, el(\"button\", \"pill skip\", n + \"번\"));\n          s.type = \"button\";\n          s.disabled = true;\n          return;\n        }\n        var b = add(g2, el(\"button\", \"pill\" + (st.num === n ? \" on\" : \"\"), n + \"번\"));\n        b.type = \"button\";\n        b.onclick = function () { st.num = n; drawNumbers(); };\n      })(n);\n    }\n    S.classes.forEach(function (k) {\n      var b = add(g1, el(\"button\", \"pill\" + (st.cls === k.c ? \" on\" : \"\"), k.c + \"반\"));\n      b.type = \"button\";\n      b.onclick = function () {\n        st.cls = k.c;\n        Array.prototype.forEach.call(g1.children, function (x) { x.className = \"pill\"; });\n        b.className = \"pill on\";\n        drawNumbers();\n      };\n    });\n    drawNumbers();\n\n    add(c3, el(\"h2\", null, \"③ 이름을 적어 주세요\"));\n    var inp = add(c3, el(\"input\"));\n    inp.type = \"text\";\n    inp.value = st.name;\n    inp.placeholder = \"예: 홍길동\";\n    inp.setAttribute(\"autocomplete\", \"off\");\n    inp.oninput = function () { st.name = inp.value; };\n\n    var c4 = add(root, el(\"div\", \"card\"));\n    add(c4, el(\"h2\", null, \"개인정보 안내\"));\n    add(c4, el(\"div\", \"privacy\", \"적은 반·번호·이름과 고른 과목은 졸업에 필요한 학점을 채웠는지 확인하는 데에만 쓰고, 확인이 끝나면 지웁니다. 다른 친구의 번호나 이름으로 내지 마세요.\"));\n    var lab = add(c4, el(\"label\", \"agree\"));\n    var cb = add(lab, el(\"input\"));\n    cb.type = \"checkbox\";\n    cb.checked = st.agree;\n    cb.onchange = function () { st.agree = cb.checked; };\n    add(lab, el(\"span\", null, \"안내를 읽었어요.\"));\n\n    var err = add(root, el(\"div\", \"err\"));\n    var next = add(root, el(\"button\", \"btn\", \"다음: 과목 고르기\"));\n    next.type = \"button\";\n    next.onclick = function () {\n      st.name = norm(inp.value);\n      var msg = !st.cls ? \"반을 골라 주세요.\" : !st.num ? \"번호를 골라 주세요.\" : nameProblem(st.name) || (!st.agree ? \"개인정보 안내를 읽고 체크해 주세요.\" : \"\");\n      err.textContent = msg;\n      if (msg) { toast(msg); return; }\n      go(\"pick\");\n    };\n  }\n\n  // ---- step 2: 학기별 과목 고르기 ----\n  function renderPick() {\n    header(root);\n    var bar = add(root, el(\"div\", \"bar\"));\n    add(bar, el(\"div\", \"who\", S.grade + \"학년 \" + st.cls + \"반 \" + st.num + \"번 \" + st.name));\n    var prog = add(bar, el(\"div\", \"prog\"));\n    var views = [];\n\n    S.semesters.forEach(function (sem) {\n      var card = add(root, el(\"div\", \"card sem\"));\n      var h2 = add(card, el(\"h2\"));\n      var ttl = add(h2, el(\"span\", null, sem.label + \" \"));\n      add(ttl, el(\"small\", null, STATUS_TEXT[sem.status] || \"\"));\n      var badge = add(h2, el(\"span\", \"badge\"));\n      var lab = add(card, el(\"label\", \"away\"));\n      var cb = add(lab, el(\"input\"));\n      cb.type = \"checkbox\";\n      cb.checked = !!st.away[sem.code];\n      add(lab, el(\"span\", null, \"이 학기에는 다른 학교에 다녔어요 (전학 온 학생만 체크)\"));\n      var body = add(card, el(\"div\"));\n      var awayNote = add(card, el(\"div\", \"sub\", \"이 학기는 고르지 않아도 돼요. 전에 다닌 학교의 과목은 담임 선생님이 따로 확인해요.\"));\n      cb.onchange = function () { st.away[sem.code] = cb.checked; refresh(); };\n      sem.groups.forEach(function (g) {\n        var box = add(body, el(\"div\", \"grp\"));\n        box.id = \"g-\" + g.id;\n        var gh = add(box, el(\"div\", \"grp-h\"));\n        var gt = add(gh, el(\"div\", \"grp-t\", g.label + \" \"));\n        add(gt, el(\"small\", null, g.n != null ? g.n + \"개 고르기\" : \"들은 과목을 모두 고르기\"));\n        var cnt = add(gh, el(\"span\", \"cnt\"));\n        var opts = [];\n        g.options.forEach(function (o) {\n          var b = add(box, el(\"button\", \"opt\"));\n          b.type = \"button\";\n          add(b, el(\"span\", \"box\"));\n          add(b, el(\"span\", \"nm\", o.name));\n          add(b, el(\"span\", \"meta\", (o.group ? o.group + \" · \" : \"\") + o.credit + \"학점\"));\n          b.onclick = function () {\n            var list = picked(g.id).slice();\n            var at = list.indexOf(o.name);\n            if (at >= 0) list.splice(at, 1);\n            else {\n              var why = blockReason(sem, g, o.name);\n              if (why === \"full\") { toast(g.n + \"개까지만 고를 수 있어요. 바꾸려면 고른 과목을 먼저 눌러서 빼 주세요.\"); return; }\n              if (why) { toast(o.name + \" — \" + why); return; }\n              list.push(o.name);\n            }\n            st.picks[g.id] = list;\n            refresh();\n          };\n          opts.push({ node: b, name: o.name });\n        });\n        views.push({ sem: sem, g: g, cnt: cnt, opts: opts });\n      });\n      views.push({ sem: sem, badge: badge, body: body, awayNote: awayNote });\n    });\n\n    var err = add(root, el(\"div\", \"err\"));\n    var row = add(root, el(\"div\", \"row\"));\n    var back = add(row, el(\"button\", \"btn gray\", \"이전\"));\n    back.type = \"button\";\n    back.onclick = function () { go(\"info\"); };\n    var next = add(row, el(\"button\", \"btn\", \"다음: 확인하기\"));\n    next.type = \"button\";\n    next.onclick = function () {\n      var ps = problems();\n      if (ps.length) {\n        err.textContent = ps.map(function (p) { return \"· \" + p.text; }).join(\"\\n\");\n        err.style.whiteSpace = \"pre-wrap\";\n        toast(\"아직 다 고르지 않은 곳이 있어요.\");\n        var first = ps[0].gid && document.getElementById(\"g-\" + ps[0].gid);\n        if (first && first.scrollIntoView) first.scrollIntoView({ behavior: \"smooth\", block: \"center\" });\n        return;\n      }\n      go(\"confirm\");\n    };\n\n    function refresh() {\n      var doneCount = 0;\n      views.forEach(function (v) {\n        if (v.badge) {\n          var ok = semDone(v.sem);\n          if (ok) doneCount++;\n          v.badge.className = \"badge\" + (ok ? \" ok\" : \"\");\n          v.badge.textContent = st.away[v.sem.code] ? \"다른 학교\" : ok ? \"다 골랐어요\" : \"고르는 중\";\n          v.body.style.display = st.away[v.sem.code] ? \"none\" : \"\";\n          v.awayNote.style.display = st.away[v.sem.code] ? \"\" : \"none\";\n          return;\n        }\n        var mine = picked(v.g.id);\n        v.cnt.textContent = mine.length + (v.g.n != null ? \" / \" + v.g.n : \"개\");\n        v.cnt.className = \"cnt\" + (groupDone(v.g) ? \" ok\" : \"\");\n        v.opts.forEach(function (o) {\n          var on = mine.indexOf(o.name) >= 0;\n          o.node.className = \"opt\" + (on ? \" on\" : !on && blockReason(v.sem, v.g, o.name) ? \" dim\" : \"\");\n          o.node.firstChild.textContent = on ? \"✓\" : \"\";\n        });\n      });\n      prog.textContent = \"다 고른 학기 \" + doneCount + \" / \" + S.semesters.length;\n      if (!problems().length) err.textContent = \"\";\n    }\n    refresh();\n  }\n\n  // ---- step 3: 확인 · 제출 ----\n  function payload() {\n    var picks = [];\n    S.semesters.forEach(function (sem) {\n      if (st.away[sem.code]) return;\n      sem.groups.forEach(function (g) { picks.push({ gid: g.id, names: picked(g.id).slice() }); });\n    });\n    return {\n      v: 1, sid: S.id, grade: S.grade, cls: st.cls, num: st.num, name: st.name, agree: st.agree,\n      away: S.semesters.filter(function (s) { return st.away[s.code]; }).map(function (s) { return s.code; }),\n      picks: picks, memo: norm(st.memo).slice(0, 300)\n    };\n  }\n  function summary(parent) {\n    S.semesters.forEach(function (sem) {\n      var c = add(parent, el(\"div\", \"card\"));\n      add(c, el(\"h2\", null, sem.label));\n      if (st.away[sem.code]) { add(c, el(\"div\", \"sub\", \"다른 학교에 다녔어요\")); return; }\n      sem.groups.forEach(function (g) {\n        var d = add(c, el(\"div\", \"sumg\"));\n        add(d, el(\"b\", null, g.label));\n        add(d, el(\"span\", null, picked(g.id).join(\", \") || \"(고른 과목 없음)\"));\n      });\n    });\n  }\n  function renderConfirm() {\n    header(root);\n    var who = add(root, el(\"div\", \"card\"));\n    add(who, el(\"h2\", null, \"마지막으로 확인해 주세요\"));\n    add(who, el(\"div\", null, S.grade + \"학년 \" + st.cls + \"반 \" + st.num + \"번 \" + st.name));\n    summary(root);\n    var m = add(root, el(\"div\", \"card\"));\n    add(m, el(\"h2\", null, \"선생님께 남길 말 \"));\n    m.firstChild.appendChild(el(\"small\", null, \"(없으면 비워 두세요)\"));\n    var ta = add(m, el(\"textarea\"));\n    ta.maxLength = 300;\n    ta.value = st.memo;\n    ta.placeholder = \"예: 2학년 1학기에 과목을 바꿨어요.\";\n    ta.oninput = function () { st.memo = ta.value; };\n    var err = add(root, el(\"div\", \"warn\"));\n    err.style.display = \"none\";\n    var row = add(root, el(\"div\", \"row\"));\n    var back = add(row, el(\"button\", \"btn gray\", \"고치기\"));\n    back.type = \"button\";\n    back.onclick = function () { if (!st.sending) go(\"pick\"); };\n    var send = add(row, el(\"button\", \"btn\", \"제출하기\"));\n    send.type = \"button\";\n    send.onclick = function () {\n      if (st.sending) return;\n      st.sending = true;\n      send.disabled = true;\n      back.disabled = true;\n      send.textContent = \"보내는 중…\";\n      err.style.display = \"none\";\n      var p = payload();\n      function ok(res) {\n        st.sending = false;\n        if (!res || !res.ok) { fail({ message: (res && res.message) || \"알 수 없는 오류\" }); return; }\n        st.result = res;\n        go(\"done\");\n      }\n      function fail(e) {\n        st.sending = false;\n        send.disabled = false;\n        back.disabled = false;\n        send.textContent = \"제출하기\";\n        var msg = String((e && e.message) || e || \"\").replace(/^(Error|Exception|오류):\\s*/i, \"\");\n        err.textContent = \"제출하지 못했어요. \" + msg + \" (고른 내용은 그대로 있으니 잠시 뒤 [제출하기]를 다시 눌러 주세요.)\";\n        err.style.display = \"\";\n      }\n      if (typeof google !== \"undefined\" && google.script && google.script.run) {\n        google.script.run.withSuccessHandler(ok).withFailureHandler(fail).submitSurvey(p);\n      } else {\n        window.__lastPayload = p;\n        setTimeout(function () { ok({ ok: true, no: 0, at: \"미리보기\", preview: true }); }, 300);\n      }\n    };\n  }\n\n  function renderDone() {\n    header(root);\n    var c = add(root, el(\"div\", \"card done\"));\n    add(c, el(\"div\", \"big\", \"✓\"));\n    add(c, el(\"h2\", null, st.result && st.result.preview ? \"미리보기: 여기까지가 학생 화면이에요\" : \"제출되었어요. 고마워요!\"));\n    if (st.result && !st.result.preview) add(c, el(\"div\", \"sub\", \"제출 번호 \" + st.result.no + \" · \" + st.result.at));\n    add(c, el(\"div\", \"sub\", S.grade + \"학년 \" + st.cls + \"반 \" + st.num + \"번 \" + st.name));\n    summary(root);\n    add(root, el(\"div\", \"note\", \"잘못 냈다면 아래 버튼으로 처음부터 다시 내면 돼요. 마지막에 낸 내용으로 확인해요.\"));\n    var again = add(root, el(\"button\", \"btn gray\", \"처음부터 다시 하기\"));\n    again.type = \"button\";\n    again.onclick = function () {\n      var cls = st.cls;\n      st = { step: \"info\", cls: cls, num: null, name: \"\", agree: false, picks: {}, away: {}, memo: \"\", sending: false, result: null, error: \"\" };\n      go(\"info\");\n    };\n  }\n\n  function renderClosed() {\n    header(root);\n    var c = add(root, el(\"div\", \"card done\"));\n    add(c, el(\"h2\", null, \"조사가 마감되었어요\"));\n    add(c, el(\"div\", \"sub\", \"더 이상 제출할 수 없어요. 고칠 내용이 있으면 담임 선생님께 말씀드려 주세요.\"));\n  }\n\n  function render() {\n    while (root.firstChild) root.removeChild(root.firstChild);\n    if (BOOT.closed) return renderClosed();\n    if (st.step === \"pick\") return renderPick();\n    if (st.step === \"confirm\") return renderConfirm();\n    if (st.step === \"done\") return renderDone();\n    return renderInfo();\n  }\n  render();\n})();\n</script></body></html>";
var CONFIG_DIALOG_HTML = "<!DOCTYPE html><html><head><base target=\"_top\"><meta charset=\"utf-8\">\n<style>\nbody{font-family:\"Malgun Gothic\",\"Apple SD Gothic Neo\",sans-serif;font-size:13px;line-height:1.6;color:#1C2333;margin:0;padding:2px}\nol{margin:0 0 8px;padding-left:20px}\ntextarea{width:100%;height:170px;box-sizing:border-box;font-family:Consolas,Menlo,monospace;font-size:11px;border:1px solid #DDD8CC;border-radius:6px;padding:6px}\n.row{margin-top:10px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}\nbutton{padding:7px 16px;font-size:13px;border-radius:6px;border:1px solid #DDD8CC;background:#fff;cursor:pointer}\nbutton.primary{background:#2C5A8A;border-color:#2C5A8A;color:#fff;font-weight:bold}\nbutton:disabled{opacity:.5;cursor:default}\n#msg{margin-top:10px;white-space:pre-wrap;border-radius:6px}\n.ok{background:#E7F1EA;color:#2F6D4F;padding:8px 10px}\n.err{background:#F7E9E3;color:#A2452C;padding:8px 10px}\n</style></head><body>\n<ol>\n<li>졸업이수요건 점검 프로그램 → <b>② 수강신청 · 학생 확인 → 학생 기초조사</b>의 4번에서 <b>[설정 코드 복사]</b>를 누릅니다.</li>\n<li>아래 칸을 누르고 <b>Ctrl+V</b>로 붙여넣은 뒤 <b>[저장]</b>을 누릅니다.</li>\n</ol>\n<textarea id=\"t\" placeholder=\"여기에 붙여넣기 (Ctrl+V)\"></textarea>\n<div class=\"row\"><button id=\"save\" class=\"primary\">저장</button><button id=\"close\">닫기</button><span id=\"more\"></span></div>\n<div id=\"msg\"></div>\n<div id=\"url\" style=\"margin-top:10px\"></div>\n<script>\n\nvar API_SETTINGS_URL = \"https://script.google.com/home/usersettings\";\nfunction clean(e) { return String((e && e.message) || e).replace(/^(Error|Exception|오류):\\s*/i, \"\"); }\nfunction copyText(input, note) {\n  input.focus();\n  input.select();\n  var ok = false;\n  try { ok = document.execCommand(\"copy\"); } catch (e) { ok = false; }\n  note.textContent = ok ? \"✓ 복사했습니다\" : \"Ctrl+C로 복사하세요\";\n}\nfunction showUrl(box, r) {\n  box.innerHTML = \"\";\n  var head = document.createElement(\"div\");\n  head.style.cssText = \"background:#E7F1EA;color:#2F6D4F;padding:8px 10px;border-radius:6px;font-weight:bold\";\n  head.textContent = r.existed ? \"✓ 이 시트의 학생용 주소입니다.\" : \"✓ 학생용 주소를 만들었습니다.\";\n  box.appendChild(head);\n  var row = document.createElement(\"div\");\n  row.style.cssText = \"display:flex;gap:6px;margin-top:8px;align-items:center\";\n  var input = document.createElement(\"input\");\n  input.readOnly = true;\n  input.value = r.url;\n  input.style.cssText = \"flex:1;min-width:0;font-size:12px;padding:6px;border:1px solid #DDD8CC;border-radius:6px\";\n  input.onfocus = function () { input.select(); };\n  var btn = document.createElement(\"button\");\n  btn.className = \"primary\";\n  btn.textContent = \"주소 복사\";\n  var note = document.createElement(\"span\");\n  note.style.cssText = \"font-size:12px;color:#2F6D4F;white-space:nowrap\";\n  btn.onclick = function () { copyText(input, note); };\n  row.appendChild(input);\n  row.appendChild(btn);\n  row.appendChild(note);\n  box.appendChild(row);\n  var tip = document.createElement(\"div\");\n  tip.style.cssText = \"margin-top:8px;color:#6B7280\";\n  tip.textContent = \"이 주소를 졸업이수요건 점검 프로그램의 5번 칸(QR 코드 나눠주기)에 붙여넣으세요. 설정을 바꿔도 주소는 그대로입니다.\";\n  box.appendChild(tip);\n  if (r.domainOnly) {\n    var warn = document.createElement(\"div\");\n    warn.style.cssText = \"margin-top:6px;color:#A2452C\";\n    warn.textContent = \"학교 계정이라 같은 학교 계정으로 로그인한 학생만 열 수 있습니다. 학생들이 학교 구글 계정으로 로그인한 휴대폰에서 열게 안내하세요.\";\n    box.appendChild(warn);\n  }\n}\nfunction showAuth(box, r) {\n  box.innerHTML = \"\";\n  var msg = document.createElement(\"div\");\n  msg.style.cssText = \"background:#F7E9E3;color:#A2452C;padding:8px 10px;border-radius:6px;line-height:1.6\";\n  msg.textContent = \"학생용 주소를 만드는 데 필요한 구글 권한 중 일부가 허용되지 않았습니다. 아래 [권한 다시 허용하기]를 눌러 새 창에서 ‘모두 선택’에 체크하고 [계속]을 누른 뒤, 이 창으로 돌아와 [다시 시도]를 누르세요.\";\n  box.appendChild(msg);\n  var row = document.createElement(\"div\");\n  row.style.cssText = \"display:flex;gap:8px;margin-top:8px;align-items:center;flex-wrap:wrap\";\n  var a = document.createElement(\"a\");\n  a.href = r.authUrl;\n  a.target = \"_blank\";\n  a.textContent = \"권한 다시 허용하기 ↗\";\n  a.style.cssText = \"font-weight:bold\";\n  var again = document.createElement(\"button\");\n  again.textContent = \"다시 시도\";\n  again.onclick = function () { makeUrl(box); };\n  row.appendChild(a);\n  row.appendChild(again);\n  box.appendChild(row);\n}\nfunction makeUrl(box) {\n  box.innerHTML = \"\";\n  var wait = document.createElement(\"div\");\n  wait.textContent = \"학생용 주소를 만드는 중… (10초쯤 걸릴 수 있어요)\";\n  box.appendChild(wait);\n  google.script.run\n    .withSuccessHandler(function (r) {\n      if (r && r.needAuth) showAuth(box, r);\n      else showUrl(box, r);\n    })\n    .withFailureHandler(function (e) {\n      box.innerHTML = \"\";\n      var err = document.createElement(\"div\");\n      err.style.cssText = \"background:#F7E9E3;color:#A2452C;padding:8px 10px;border-radius:6px;white-space:pre-wrap\";\n      err.textContent = clean(e);\n      box.appendChild(err);\n      var row = document.createElement(\"div\");\n      row.style.cssText = \"display:flex;gap:8px;margin-top:8px;align-items:center;flex-wrap:wrap\";\n      if (clean(e).indexOf(\"usersettings\") >= 0) {\n        var a = document.createElement(\"a\");\n        a.href = API_SETTINGS_URL;\n        a.target = \"_blank\";\n        a.textContent = \"Apps Script API 설정 열기 ↗\";\n        row.appendChild(a);\n      }\n      var again = document.createElement(\"button\");\n      again.textContent = \"다시 시도\";\n      again.onclick = function () { makeUrl(box); };\n      row.appendChild(again);\n      box.appendChild(row);\n    })\n    .createStudentUrl();\n}\n\nvar t = document.getElementById(\"t\");\nvar btn = document.getElementById(\"save\");\nvar msg = document.getElementById(\"msg\");\nvar more = document.getElementById(\"more\");\nfunction show(kind, text) { msg.className = kind; msg.textContent = text; }\nfunction save(confirmed) {\n  btn.disabled = true;\n  more.innerHTML = \"\";\n  show(\"\", \"저장하는 중…\");\n  google.script.run\n    .withSuccessHandler(function (r) {\n      btn.disabled = false;\n      if (r && r.confirm) {\n        show(\"err\", r.confirm);\n        var yes = document.createElement(\"button\");\n        yes.textContent = \"그래도 바꾸기\";\n        yes.onclick = function () { save(true); };\n        more.appendChild(yes);\n        return;\n      }\n      show(\"ok\", (r.same ? \"✓ 이미 같은 설정이 들어 있습니다.\" : \"✓ 저장했습니다. 다시 배포하지 않아도 학생 화면에 바로 반영됩니다.\") + \"\\n\\n\" + r.summary);\n      if (r.autoDeploy) {\n        var box = document.getElementById(\"url\");\n        if (r.url) showUrl(box, { url: r.url, existed: true });\n        else {\n          box.innerHTML = \"\";\n          var mk = document.createElement(\"button\");\n          mk.className = \"primary\";\n          mk.textContent = \"다음: 학생용 주소 만들기\";\n          mk.onclick = function () { makeUrl(box); };\n          box.appendChild(mk);\n        }\n      }\n    })\n    .withFailureHandler(function (e) {\n      btn.disabled = false;\n      show(\"err\", String((e && e.message) || e).replace(/^(Error|Exception|오류):\\s*/i, \"\"));\n    })\n    .saveSurveyConfig(t.value, !!confirmed);\n}\nbtn.onclick = function () { save(false); };\ndocument.getElementById(\"close\").onclick = function () { google.script.host.close(); };\nt.focus();\n</script>\n</body></html>";
var URL_DIALOG_HTML = "<!DOCTYPE html><html><head><base target=\"_top\"><meta charset=\"utf-8\">\n<style>\nbody{font-family:\"Malgun Gothic\",\"Apple SD Gothic Neo\",sans-serif;font-size:13px;line-height:1.6;color:#1C2333;margin:0;padding:2px}\nbutton{padding:7px 16px;font-size:13px;border-radius:6px;border:1px solid #DDD8CC;background:#fff;cursor:pointer}\nbutton.primary{background:#2C5A8A;border-color:#2C5A8A;color:#fff;font-weight:bold}\nbutton:disabled{opacity:.5;cursor:default}\n</style></head><body>\n<div id=\"url\"></div>\n<div style=\"margin-top:12px\"><button id=\"close\">닫기</button></div>\n<script>\n\nvar API_SETTINGS_URL = \"https://script.google.com/home/usersettings\";\nfunction clean(e) { return String((e && e.message) || e).replace(/^(Error|Exception|오류):\\s*/i, \"\"); }\nfunction copyText(input, note) {\n  input.focus();\n  input.select();\n  var ok = false;\n  try { ok = document.execCommand(\"copy\"); } catch (e) { ok = false; }\n  note.textContent = ok ? \"✓ 복사했습니다\" : \"Ctrl+C로 복사하세요\";\n}\nfunction showUrl(box, r) {\n  box.innerHTML = \"\";\n  var head = document.createElement(\"div\");\n  head.style.cssText = \"background:#E7F1EA;color:#2F6D4F;padding:8px 10px;border-radius:6px;font-weight:bold\";\n  head.textContent = r.existed ? \"✓ 이 시트의 학생용 주소입니다.\" : \"✓ 학생용 주소를 만들었습니다.\";\n  box.appendChild(head);\n  var row = document.createElement(\"div\");\n  row.style.cssText = \"display:flex;gap:6px;margin-top:8px;align-items:center\";\n  var input = document.createElement(\"input\");\n  input.readOnly = true;\n  input.value = r.url;\n  input.style.cssText = \"flex:1;min-width:0;font-size:12px;padding:6px;border:1px solid #DDD8CC;border-radius:6px\";\n  input.onfocus = function () { input.select(); };\n  var btn = document.createElement(\"button\");\n  btn.className = \"primary\";\n  btn.textContent = \"주소 복사\";\n  var note = document.createElement(\"span\");\n  note.style.cssText = \"font-size:12px;color:#2F6D4F;white-space:nowrap\";\n  btn.onclick = function () { copyText(input, note); };\n  row.appendChild(input);\n  row.appendChild(btn);\n  row.appendChild(note);\n  box.appendChild(row);\n  var tip = document.createElement(\"div\");\n  tip.style.cssText = \"margin-top:8px;color:#6B7280\";\n  tip.textContent = \"이 주소를 졸업이수요건 점검 프로그램의 5번 칸(QR 코드 나눠주기)에 붙여넣으세요. 설정을 바꿔도 주소는 그대로입니다.\";\n  box.appendChild(tip);\n  if (r.domainOnly) {\n    var warn = document.createElement(\"div\");\n    warn.style.cssText = \"margin-top:6px;color:#A2452C\";\n    warn.textContent = \"학교 계정이라 같은 학교 계정으로 로그인한 학생만 열 수 있습니다. 학생들이 학교 구글 계정으로 로그인한 휴대폰에서 열게 안내하세요.\";\n    box.appendChild(warn);\n  }\n}\nfunction showAuth(box, r) {\n  box.innerHTML = \"\";\n  var msg = document.createElement(\"div\");\n  msg.style.cssText = \"background:#F7E9E3;color:#A2452C;padding:8px 10px;border-radius:6px;line-height:1.6\";\n  msg.textContent = \"학생용 주소를 만드는 데 필요한 구글 권한 중 일부가 허용되지 않았습니다. 아래 [권한 다시 허용하기]를 눌러 새 창에서 ‘모두 선택’에 체크하고 [계속]을 누른 뒤, 이 창으로 돌아와 [다시 시도]를 누르세요.\";\n  box.appendChild(msg);\n  var row = document.createElement(\"div\");\n  row.style.cssText = \"display:flex;gap:8px;margin-top:8px;align-items:center;flex-wrap:wrap\";\n  var a = document.createElement(\"a\");\n  a.href = r.authUrl;\n  a.target = \"_blank\";\n  a.textContent = \"권한 다시 허용하기 ↗\";\n  a.style.cssText = \"font-weight:bold\";\n  var again = document.createElement(\"button\");\n  again.textContent = \"다시 시도\";\n  again.onclick = function () { makeUrl(box); };\n  row.appendChild(a);\n  row.appendChild(again);\n  box.appendChild(row);\n}\nfunction makeUrl(box) {\n  box.innerHTML = \"\";\n  var wait = document.createElement(\"div\");\n  wait.textContent = \"학생용 주소를 만드는 중… (10초쯤 걸릴 수 있어요)\";\n  box.appendChild(wait);\n  google.script.run\n    .withSuccessHandler(function (r) {\n      if (r && r.needAuth) showAuth(box, r);\n      else showUrl(box, r);\n    })\n    .withFailureHandler(function (e) {\n      box.innerHTML = \"\";\n      var err = document.createElement(\"div\");\n      err.style.cssText = \"background:#F7E9E3;color:#A2452C;padding:8px 10px;border-radius:6px;white-space:pre-wrap\";\n      err.textContent = clean(e);\n      box.appendChild(err);\n      var row = document.createElement(\"div\");\n      row.style.cssText = \"display:flex;gap:8px;margin-top:8px;align-items:center;flex-wrap:wrap\";\n      if (clean(e).indexOf(\"usersettings\") >= 0) {\n        var a = document.createElement(\"a\");\n        a.href = API_SETTINGS_URL;\n        a.target = \"_blank\";\n        a.textContent = \"Apps Script API 설정 열기 ↗\";\n        row.appendChild(a);\n      }\n      var again = document.createElement(\"button\");\n      again.textContent = \"다시 시도\";\n      again.onclick = function () { makeUrl(box); };\n      row.appendChild(again);\n      box.appendChild(row);\n    })\n    .createStudentUrl();\n}\n\ndocument.getElementById(\"close\").onclick = function () { google.script.host.close(); };\nmakeUrl(document.getElementById(\"url\"));\n</script>\n</body></html>";
var DEPLOY_SCOPES = ["https://www.googleapis.com/auth/spreadsheets.currentonly","https://www.googleapis.com/auth/script.container.ui","https://www.googleapis.com/auth/userinfo.email","https://www.googleapis.com/auth/script.external_request","https://www.googleapis.com/auth/script.projects","https://www.googleapis.com/auth/script.deployments"];
var NOT_READY_HTML = "<!DOCTYPE html><html lang=\"ko\"><head><meta charset=\"utf-8\">\n<style>body{font-family:\"Malgun Gothic\",\"Apple SD Gothic Neo\",sans-serif;background:#F6F4EF;color:#1C2333;margin:0;padding:32px 18px;line-height:1.7}\n.card{max-width:460px;margin:0 auto;background:#fff;border:1px solid #DDD8CC;border-radius:12px;padding:22px}\nh1{font-size:18px;margin:0 0 8px}.t{font-size:13px;color:#6B7280;margin-top:14px;border-top:1px solid #DDD8CC;padding-top:12px}</style></head>\n<body><div class=\"card\"><h1>아직 조사 준비 중이에요</h1>\n<div>선생님이 조사를 준비하고 있어요. 조금 뒤에 다시 열어 주세요.</div>\n<div class=\"t\"><b>선생님께:</b> 이 시트에 아직 조사 설정이 없습니다. 구글 시트를 열고 위쪽 메뉴 <b>[📋 기초조사 → 설정 붙여넣기]</b>로 프로그램에서 복사한 설정 코드를 넣어 주세요. 다시 배포할 필요는 없습니다.</div>\n</div></body></html>";

var SHEET_SETTINGS = "설정";
var SHEET_STATUS = "제출현황";
var SHEET_LOG = "제출기록";
var SHEET_PICKS = "선택과목";
var SHEET_CONFIG = "조사설정";
var SHEET_START = "시작하기"; // 템플릿에 넣어 두는 안내 시트 — 조사 준비가 끝나면 지웁니다
var DEFAULT_SHEET_NAMES = ["시트1", "Sheet1"]; // 새 시트에 처음부터 있는 빈 시트
var LOG_HEADER = ["제출번호", "제출시각", "학년", "반", "번호", "이름", "다른 학교에 다닌 학기", "남긴 말", "고른 과목"];
var PICK_HEADER = ["제출번호", "제출시각", "학년", "반", "번호", "이름", "학기", "선택그룹", "과목"];
var AWAY_MARK = "(다른 학교에 다님)";
var META_LABEL = "조사 정보(수정하지 마세요)";
var MENU_NAME = "📋 기초조사";
var CONFIG_CHUNK = 40000; // 시트 한 칸에는 5만 글자까지만 들어가서 나눠 적습니다
var SEEN_PROP = "BUNDLED_SEEN";
var DEPLOYMENT_PROP = "DEPLOYMENT_ID";
var URL_PROP = "WEB_APP_URL";
var API_SETTINGS_URL = "https://script.google.com/home/usersettings";

var SURVEY = null; // load_()가 '조사설정' 시트에서 읽어 채웁니다

// 시트를 열면 위쪽 메뉴에 [📋 기초조사]가 생깁니다.
function onOpen() {
  var menu = SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem("설정 붙여넣기 / 바꾸기", "openConfigDialog");
  if (AUTO_DEPLOY) menu.addItem("학생용 주소 만들기 / 보기", "openUrlDialog");
  menu.addItem("지금 설정 보기", "showConfigSummary").addToUi();
}

function openUrlDialog() {
  requireOwner_();
  var html = HtmlService.createHtmlOutput(URL_DIALOG_HTML).setWidth(560).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, "학생용 주소");
}

// ---- 학생용 주소 만들기 (자동 배포) ----
// 템플릿 시트의 매니페스트(appsscript.json)에 웹 앱 설정(나 · 모든 사용자)이 들어 있어서,
// Apps Script API로 버전을 만들고 배포하면 편집기를 열지 않아도 웹 앱 주소가 생깁니다.
// 선생님 계정에서 'Google Apps Script API'가 켜져 있어야 합니다 (API_SETTINGS_URL, 계정마다 처음 한 번).
function createStudentUrl() {
  requireOwner_();
  if (!AUTO_DEPLOY) fail_("이 시트는 자동 배포를 쓸 수 없습니다. Apps Script 편집기에서 [배포 → 새 배포 → 웹 앱]으로 직접 배포해 주세요.");
  var ss = book_();
  load_(ss);
  if (!SURVEY) fail_("먼저 [설정 붙여넣기]로 조사 설정을 넣어 주세요.");
  // 구글 권한 화면은 항목마다 체크박스가 있어서, ‘모두 선택’을 하지 않으면 배포 권한이 빠진 채로 허용될 수 있습니다
  var auth = missingScopes_();
  if (auth) return { ok: false, needAuth: true, authUrl: auth.url };
  var props = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) fail_("다른 작업 중입니다. 잠시 뒤 다시 눌러 주세요.");
  try {
    // 이미 만든 주소가 살아 있으면 그대로 씁니다 (설정은 시트에서 읽으므로 다시 배포할 필요 없음)
    var existing = props.getProperty(DEPLOYMENT_PROP);
    if (existing) {
      var got = api_("get", "/deployments/" + encodeURIComponent(existing));
      var old = got.ok ? webApp_(got.body) : null;
      if (old) return saveUrl_(ss, old, true);
      if (!got.ok && got.code !== 404) apiFail_(got);
    }
    var v = api_("post", "/versions", { description: "학생 기초조사" });
    if (!v.ok) apiFail_(v);
    var dep = api_("post", "/deployments", {
      versionNumber: v.body.versionNumber,
      manifestFileName: "appsscript",
      description: "학생 기초조사 (자동 배포)"
    });
    if (!dep.ok) apiFail_(dep);
    var app = webApp_(dep.body);
    if (!app) fail_("배포는 됐지만 웹 앱 주소가 없습니다. 템플릿의 appsscript.json에 webapp 설정이 빠졌습니다. 제작팀에 알려 주세요.");
    props.setProperty(DEPLOYMENT_PROP, dep.body.deploymentId);
    return saveUrl_(ss, app, false);
  } finally {
    lock.releaseLock();
  }
}

// 매니페스트에 적힌 권한 중 선생님이 아직 허용하지 않은 것이 있으면 다시 허용하는 주소를 돌려줍니다
function missingScopes_() {
  try {
    var info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL, DEPLOY_SCOPES);
    if (info.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) return { url: info.getAuthorizationUrl() };
  } catch (e) {
    // 권한을 확인하는 기능이 없는 환경이면 그냥 배포를 시도합니다 (실패하면 apiFail_이 안내)
  }
  return null;
}

function saveUrl_(ss, app, existed) {
  PropertiesService.getScriptProperties().setProperty(URL_PROP, app.url);
  var settings = ss.getSheetByName(SHEET_SETTINGS);
  if (settings) settings.getRange(10, 1, 1, 2).setValues([["학생용 주소", app.url]]);
  return { ok: true, url: app.url, existed: existed, domainOnly: app.access !== "ANYONE_ANONYMOUS" };
}

function webApp_(deployment) {
  var eps = (deployment && deployment.entryPoints) || [];
  for (var i = 0; i < eps.length; i++) {
    var w = eps[i].entryPointType === "WEB_APP" && eps[i].webApp;
    if (w && w.url) return { url: w.url, access: (w.entryPointConfig && w.entryPointConfig.access) || "" };
  }
  return null;
}

function api_(method, path, body) {
  var opts = {
    method: method,
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  if (body) {
    opts.contentType = "application/json";
    opts.payload = JSON.stringify(body);
  }
  var res = UrlFetchApp.fetch("https://script.googleapis.com/v1/projects/" + ScriptApp.getScriptId() + path, opts);
  var code = res.getResponseCode();
  var json = {};
  try {
    json = JSON.parse(res.getContentText() || "{}");
  } catch (e) {
    json = {};
  }
  return { ok: code >= 200 && code < 300, code: code, body: json, message: (json.error && json.error.message) || "" };
}

function apiFail_(r) {
  var m = r.message || "";
  var raw = "\n\n[구글이 보낸 원래 문구] " + r.code + " " + m;
  // 선생님 계정의 'Google Apps Script API' 스위치가 꺼져 있는 경우
  if (/usersettings|User has not enabled the Apps Script API/i.test(m)) {
    fail_(
      "Apps Script API가 꺼져 있습니다. " + API_SETTINGS_URL + " 에서 ‘Google Apps Script API’를 ‘사용’으로 바꾼 뒤, 몇 분 지나 다시 눌러 주세요." +
        " 켰는데도 이 문구가 나오면, 스위치를 켠 구글 계정과 이 시트를 연 구글 계정(" + Session.getEffectiveUser().getEmail() + ")이 같은지 확인하세요." +
        raw
    );
  }
  // 스위치와는 별개로, 이 스크립트가 붙은 구글 클라우드 프로젝트 쪽에서 Apps Script API를 쓸 수 없는 경우
  if (/has not been used in project|SERVICE_DISABLED|console\.(developers|cloud)\.google\.com/i.test(m)) {
    fail_(
      "이 시트의 Apps Script가 연결된 구글 클라우드 프로젝트에서 Apps Script API를 쓸 수 없습니다. (계정 설정 스위치와는 다른 곳입니다) Apps Script 편집기에서 [배포 → 새 배포 → 웹 앱]으로 직접 배포해 주세요." +
        raw
    );
  }
  // 매니페스트(appsscript.json)에 배포 권한(script.projects · script.deployments)이 없어서 받은 토큰으로는 배포할 수 없는 경우
  if (/insufficient authentication scopes|ACCESS_TOKEN_SCOPE_INSUFFICIENT/i.test(m)) {
    fail_("배포 권한이 허용되지 않았습니다. 권한 허용 화면에서 ‘모두 선택’에 체크하지 않았거나, 이 시트의 권한 목록(appsscript.json)이 예전 것일 수 있습니다. 창을 닫고 [학생용 주소 만들기]를 다시 눌러 권한을 모두 허용하거나, Apps Script 편집기에서 [배포 → 새 배포 → 웹 앱]으로 직접 배포해 주세요." + raw);
  }
  if (r.code === 403) {
    fail_("자동 배포 권한이 없습니다. 학교(교육청) 계정이라면 관리자가 막아 두었을 수 있습니다. Apps Script 편집기에서 [배포 → 새 배포 → 웹 앱]으로 직접 배포해 주세요. (" + m + ")");
  }
  fail_("학생용 주소를 만들지 못했습니다 (" + r.code + "). 잠시 뒤 다시 눌러 주세요. " + m);
}

function openConfigDialog() {
  requireOwner_();
  var html = HtmlService.createHtmlOutput(CONFIG_DIALOG_HTML).setWidth(560).setHeight(470);
  SpreadsheetApp.getUi().showModalDialog(html, "기초조사 설정 붙여넣기");
}

function showConfigSummary() {
  requireOwner_();
  load_(book_());
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "지금 기초조사 설정",
    SURVEY ? summary_(SURVEY) : "아직 설정이 없습니다. 메뉴 [" + MENU_NAME + " → 설정 붙여넣기]로 프로그램에서 복사한 설정 코드를 넣어 주세요.",
    ui.ButtonSet.OK
  );
}

// 설정 붙여넣기 창의 [저장]이 부릅니다. 시트 주인만 쓸 수 있습니다 (학생 화면에서는 부를 수 없음).
function saveSurveyConfig(text, confirmed) {
  requireOwner_();
  var ss = book_();
  var s = parseConfig_(text);
  load_(ss);
  if (SURVEY && SURVEY.fp === s.fp) return withUrlInfo_({ ok: true, same: true, summary: summary_(s) });
  if (SURVEY && SURVEY.id !== s.id && !confirmed) {
    return {
      ok: false,
      confirm: "지금 시트에 있는 조사(" + SURVEY.title + ")와 다른 조사의 설정입니다. 이미 받은 응답은 그대로 남고, 앞으로 받는 응답만 새 설정을 따릅니다. 그래도 바꿀까요?"
    };
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) fail_("학생 제출을 처리하는 중입니다. 잠시 뒤 다시 눌러 주세요.");
  try {
    writeConfig_(ss, s);
    SURVEY = s;
    prepare_(ss, false);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return withUrlInfo_({ ok: true, summary: summary_(s) });
}

// 설정 창이 저장 뒤에 [학생용 주소 만들기] 버튼을 보여줄지, 이미 만든 주소를 보여줄지 정하는 데 씁니다
function withUrlInfo_(r) {
  r.autoDeploy = !!AUTO_DEPLOY;
  r.url = PropertiesService.getScriptProperties().getProperty(URL_PROP) || "";
  return r;
}

// (선택) 편집기에서 시트 탭을 미리 만들어 보고 싶을 때 실행합니다. 안 해도 학생이 처음 들어올 때 자동으로 만들어집니다.
function setup() {
  requireOwner_();
  var ss = book_();
  load_(ss);
  if (!SURVEY) {
    Logger.log("아직 조사 설정이 없습니다. 시트를 새로고침한 뒤 메뉴 [" + MENU_NAME + " → 설정 붙여넣기]로 설정을 넣어 주세요.");
    return;
  }
  prepare_(ss, true);
  SpreadsheetApp.flush();
  Logger.log("준비 완료: 시트 아래쪽에 설정 · 제출현황 · 제출기록 · 선택과목 탭이 생겼습니다.");
}

// 학생이 QR 코드로 들어오면 실행됩니다 (?ban=3 이면 3반이 미리 골라져 있음).
function doGet(e) {
  var ss = book_();
  load_(ss);
  if (!SURVEY) {
    return HtmlService.createHtmlOutput(NOT_READY_HTML)
      .setTitle("기초조사 준비 중")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }
  // 설정을 바꾼 직후 여러 학생이 동시에 열어도 시트를 한 번만 고치도록 잠급니다
  if (needsPrepare_(ss)) {
    var lock = LockService.getScriptLock();
    if (lock.tryLock(20000)) {
      try {
        prepare_(ss, false);
      } finally {
        lock.releaseLock();
      }
    }
  }
  var p = (e && e.parameter) || {};
  var ban = parseInt(p.ban, 10);
  var boot = { survey: SURVEY, ban: classOf_(ban) ? ban : null, closed: closed_(ss) };
  var json = JSON.stringify(boot).replace(/</g, "\\u003c");
  var html = PAGE_HTML.replace("/*__BOOT__*/null", function () { return json; });
  return HtmlService.createHtmlOutput(html)
    .setTitle(SURVEY.title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// 학생 화면의 [제출하기]가 부릅니다. 받은 내용을 조사 설정과 다시 대조한 뒤에만 시트에 적습니다.
function submitSurvey(p) {
  load_(book_());
  if (!SURVEY) fail_("아직 조사가 준비되지 않았어요. 선생님께 알려 주세요.");
  var v = validate_(p);
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) fail_("지금 제출하는 친구들이 많아요. 10초쯤 뒤에 다시 눌러 주세요.");
  try {
    var ss = book_();
    prepare_(ss, false);
    if (closed_(ss)) fail_("조사가 마감되어 제출할 수 없어요.");
    var log = ss.getSheetByName(SHEET_LOG);
    var picks = ss.getSheetByName(SHEET_PICKS);
    var no = nextNo_(log);
    var at = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    var base = [no, at, SURVEY.grade, v.cls, v.num, safe_(v.name)];

    var logRow = log.getLastRow() + 1;
    ensureRows_(log, logRow);
    log.getRange(logRow, 2, 1, 1).setNumberFormat("@");
    log.getRange(logRow, 6, 1, 4).setNumberFormat("@");
    log.getRange(logRow, 1, 1, 9).setValues([base.concat([v.awayLabels, safe_(v.memo), safe_(v.summary)])]);

    var out = v.rows.map(function (r) { return base.concat([r.semLabel, safe_(r.group), safe_(r.subject)]); });
    if (out.length > 0) {
      var start = picks.getLastRow() + 1;
      ensureRows_(picks, start + out.length - 1);
      picks.getRange(start, 2, out.length, 1).setNumberFormat("@");
      picks.getRange(start, 6, out.length, 4).setNumberFormat("@");
      picks.getRange(start, 1, out.length, 9).setValues(out);
    }
    SpreadsheetApp.flush();
    return { ok: true, no: no, at: at };
  } finally {
    lock.releaseLock();
  }
}

function book_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) fail_("이 코드는 구글 시트의 [확장 프로그램 → Apps Script]에 붙여넣어야 합니다.");
  return ss;
}

function fail_(msg) {
  throw new Error(msg);
}

// 학생(웹 앱 방문자)은 로그인하지 않았거나 다른 계정이라 여기서 막힙니다.
function requireOwner_() {
  var me = Session.getEffectiveUser().getEmail();
  var who = Session.getActiveUser().getEmail();
  if (!me || who !== me) fail_("이 시트를 만든 선생님 계정에서만 할 수 있습니다.");
}

// ---- 조사 설정 읽기 · 쓰기 ----
function load_(ss) {
  if (SURVEY) return SURVEY;
  if (bundledPending_()) {
    var lock = LockService.getScriptLock();
    if (lock.tryLock(20000)) {
      try {
        if (bundledPending_()) {
          writeConfig_(ss, checkSurvey_(BUNDLED_SURVEY));
          var seen = seenBundles_();
          seen.push(BUNDLED_SURVEY.fp);
          PropertiesService.getScriptProperties().setProperty(SEEN_PROP, JSON.stringify(seen.slice(-50)));
        }
      } finally {
        lock.releaseLock();
      }
    }
  }
  SURVEY = readConfig_(ss);
  return SURVEY;
}

// 설치 코드에 함께 붙어 온 설정은 처음 한 번만 시트로 옮깁니다. 그 뒤로는 메뉴에서 붙여넣은 설정이 우선입니다.
function bundledPending_() {
  if (typeof BUNDLED_SURVEY === "undefined" || !BUNDLED_SURVEY) return false;
  return seenBundles_().indexOf(BUNDLED_SURVEY.fp) < 0;
}

function seenBundles_() {
  try {
    var v = JSON.parse(PropertiesService.getScriptProperties().getProperty(SEEN_PROP) || "[]");
    return Array.isArray(v) ? v : [];
  } catch (e) {
    return [];
  }
}

function readConfig_(ss) {
  var sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh || sh.getLastRow() < 2) return null;
  var text = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); }).join("");
  try {
    return checkSurvey_(JSON.parse(text));
  } catch (e) {
    return null;
  }
}

function writeConfig_(ss, s) {
  var sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) sh = ss.insertSheet(SHEET_CONFIG);
  sh.clear();
  var text = JSON.stringify(s);
  var rows = [["기초조사 설정입니다. 고치거나 지우지 마세요. 설정을 바꾸려면 시트 위 메뉴 [" + MENU_NAME + " → 설정 붙여넣기]를 쓰세요."]];
  for (var i = 0; i < text.length; i += CONFIG_CHUNK) rows.push([text.slice(i, i + CONFIG_CHUNK)]);
  ensureRows_(sh, rows.length);
  sh.getRange(1, 1, rows.length, 1).setNumberFormat("@").setValues(rows);
  try {
    sh.hideSheet();
  } catch (e) {
    // 보이는 시트가 이것 하나뿐이면 숨길 수 없습니다 — 그대로 둡니다
  }
}

function parseConfig_(text) {
  var s = String(text == null ? "" : text).trim();
  if (!s) fail_("붙여넣은 내용이 없습니다. 프로그램에서 [설정 코드 복사]를 누른 뒤 이 칸에 붙여넣어(Ctrl+V) 주세요.");
  var obj;
  try {
    obj = JSON.parse(s);
  } catch (e) {
    fail_("설정 코드가 아니거나 일부가 빠졌습니다. 프로그램에서 [설정 코드 복사]를 다시 눌러 전부 붙여넣어 주세요.");
  }
  return checkSurvey_(obj);
}

// 프로그램이 만든 설정 그대로인지(잘리거나 손으로 고치지 않았는지) 설정 버전(fp)으로 확인합니다.
function checkSurvey_(s) {
  if (!s || typeof s !== "object" || s.v !== 1 || typeof s.id !== "string" || typeof s.fp !== "string") {
    fail_("기초조사 설정 코드가 아닙니다. 프로그램의 [설정 코드 복사]로 복사한 글을 그대로 붙여넣어 주세요.");
  }
  var content = { id: s.id, title: s.title, grade: s.grade, term: s.term, notice: s.notice, deadline: s.deadline, classes: s.classes, semesters: s.semesters };
  if ("v-" + fnv_(JSON.stringify(content)) !== s.fp) {
    fail_("설정 코드 일부가 빠졌거나 바뀌었습니다. 프로그램에서 [설정 코드 복사]를 다시 눌러 전부 붙여넣어 주세요.");
  }
  if (!(s.grade >= 1 && s.grade <= 3) || !Array.isArray(s.classes) || !s.classes.length || !Array.isArray(s.semesters) || !s.semesters.length) {
    fail_("설정에 학년·반·학기 정보가 비어 있습니다. 프로그램에서 조사 설정을 마친 뒤 다시 복사해 주세요.");
  }
  return s;
}

function fnv_(str) {
  var h = 0x811c9dc5;
  for (var i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  var hex = h.toString(16);
  while (hex.length < 8) hex = "0" + hex;
  return hex;
}

function summary_(s) {
  var total = 0;
  s.classes.forEach(function (k) { total += k.size - (k.skip || []).length; });
  return [
    "조사 이름: " + s.title,
    "대상: " + s.grade + "학년 " + s.classes.length + "개 반 (" + total + "명)",
    "학기: " + s.semesters.map(function (x) { return x.label; }).join(", "),
    "설정 버전: " + s.fp + (s.createdAt ? " (만든 날 " + s.createdAt + ")" : "")
  ].join("\n");
}

function needsPrepare_(ss) {
  var settings = ss.getSheetByName(SHEET_SETTINGS);
  if (!settings || !ss.getSheetByName(SHEET_LOG) || !ss.getSheetByName(SHEET_PICKS) || !ss.getSheetByName(SHEET_STATUS)) return true;
  return String(settings.getRange("B5").getValue()) !== SURVEY.fp;
}

function prepare_(ss, force) {
  var settings = ss.getSheetByName(SHEET_SETTINGS);
  var fresh = !settings;
  if (!settings) settings = ss.insertSheet(SHEET_SETTINGS);
  ensureHeader_(ss, SHEET_LOG, LOG_HEADER);
  ensureHeader_(ss, SHEET_PICKS, PICK_HEADER);
  var current = fresh ? "" : String(settings.getRange("B5").getValue());
  if (!force && !fresh && current === SURVEY.fp) {
    if (!ss.getSheetByName(SHEET_STATUS)) buildStatus_(ss);
    return;
  }
  var status = fresh ? "" : String(settings.getRange("B1").getValue());
  status = status.indexOf("마감") >= 0 ? "마감" : "받는 중";
  var total = 0;
  SURVEY.classes.forEach(function (k) { total += k.size - (k.skip || []).length; });
  var meta = { kind: "graduation-survey", v: 1, id: SURVEY.id, fp: SURVEY.fp, title: SURVEY.title, grade: SURVEY.grade, classes: SURVEY.classes, semesters: SURVEY.semesters.map(function (s) { return s.code; }) };
  settings.getRange(1, 2, 9, 1).setNumberFormat("@");
  settings.getRange(1, 1, 9, 2).setValues([
    ["조사 상태", status],
    ["조사 이름", safe_(SURVEY.title)],
    ["대상", SURVEY.grade + "학년 " + SURVEY.classes.length + "개 반 (" + total + "명)"],
    ["설정을 만든 날", SURVEY.createdAt],
    ["설정 버전", SURVEY.fp],
    ["", ""],
    ["마감하는 법", "B1 칸을 '마감'으로 바꾸면 학생 화면에 '조사가 마감되었어요'가 나오고 더 이상 제출되지 않습니다. 다시 받으려면 '받는 중'으로 바꾸세요."],
    ["설정 바꾸는 법", "프로그램에서 [설정 코드 복사] → 이 시트 위 메뉴 [" + MENU_NAME + " → 설정 붙여넣기]. 다시 배포하지 않아도 학생 화면에 바로 반영되고 QR 코드도 그대로입니다."],
    [META_LABEL, JSON.stringify(meta)]
  ]);
  settings.getRange("B1").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["받는 중", "마감"], true).setAllowInvalid(false).build()
  );
  settings.getRange(1, 1, 9, 1).setFontWeight("bold");
  settings.setColumnWidth(1, 170);
  settings.setColumnWidth(2, 560);
  buildStatus_(ss);
  removeLeftoverSheets_(ss);
}

// 조사에 쓰지 않는 시트를 지웁니다: 템플릿의 안내 시트(시작하기)와, 아무것도 적지 않은 기본 시트(시트1).
// 선생님이 시트1에 뭔가 적어 두었다면 지우지 않습니다.
function removeLeftoverSheets_(ss) {
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    var blank = sh.getLastRow() === 0 && sh.getLastColumn() === 0;
    if (name !== SHEET_START && !(DEFAULT_SHEET_NAMES.indexOf(name) >= 0 && blank)) return;
    if (ss.getSheets().length <= 1) return;
    try {
      ss.deleteSheet(sh);
    } catch (e) {
      // 지우지 못해도 조사에는 지장이 없습니다
    }
  });
  var status = ss.getSheetByName(SHEET_STATUS);
  if (status) {
    // 시트를 열면 제출현황이 맨 앞에 보이게
    try {
      ss.setActiveSheet(status);
      ss.moveActiveSheet(1);
    } catch (e) {
      // 순서를 못 바꿔도 괜찮습니다
    }
  }
}

function ensureHeader_(ss, name, header) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold").setBackground("#EAF0F6");
    sh.setFrozenRows(1);
  }
  return sh;
}

function buildStatus_(ss) {
  var sh = ss.getSheetByName(SHEET_STATUS);
  if (!sh) sh = ss.insertSheet(SHEET_STATUS);
  sh.clear();
  var students = 0;
  SURVEY.classes.forEach(function (k) { students += k.size; });
  ensureRows_(sh, 3 + Math.max(students, SURVEY.classes.length) + 2);
  sh.getRange(1, 1).setValue("제출 현황 — 학생이 제출할 때마다 자동으로 계산됩니다. 이 시트는 고치지 마세요.");
  sh.getRange(1, 1).setFontWeight("bold");
  sh.getRange(3, 1, 1, 4).setValues([["반", "인원", "제출", "미제출"]]).setFontWeight("bold").setBackground("#EAF0F6");
  sh.getRange(3, 6, 1, 4).setValues([["반", "번호", "제출 횟수", "상태"]]).setFontWeight("bold").setBackground("#EAF0F6");
  var classRows = [];
  var studentRows = [];
  var r = 4;
  SURVEY.classes.forEach(function (k, i) {
    var row = 4 + i;
    var skip = k.skip || [];
    classRows.push([k.c, k.size - skip.length, '=COUNTIFS(F:F,A' + row + ',I:I,"제출")', "=B" + row + "-C" + row]);
    for (var n = 1; n <= k.size; n++) {
      if (skip.indexOf(n) >= 0) studentRows.push([k.c, n, "", "결번"]);
      else studentRows.push([k.c, n, "=COUNTIFS('" + SHEET_LOG + "'!D:D,F" + r + ",'" + SHEET_LOG + "'!E:E,G" + r + ")", '=IF(H' + r + '>0,"제출","미제출")']);
      r++;
    }
  });
  if (classRows.length) sh.getRange(4, 1, classRows.length, 4).setValues(classRows);
  if (studentRows.length) sh.getRange(4, 6, studentRows.length, 4).setValues(studentRows);
  sh.setFrozenRows(3);
}

function ensureRows_(sh, lastRowNeeded) {
  var max = sh.getMaxRows();
  if (lastRowNeeded > max) sh.insertRowsAfter(max, lastRowNeeded - max + 200);
}

function closed_(ss) {
  var sh = ss.getSheetByName(SHEET_SETTINGS);
  return !!sh && String(sh.getRange("B1").getValue()).indexOf("마감") >= 0;
}

function classOf_(c) {
  for (var i = 0; i < SURVEY.classes.length; i++) if (SURVEY.classes[i].c === c) return SURVEY.classes[i];
  return null;
}

function nextNo_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var mx = 0;
  sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (row) {
    var n = Number(row[0]);
    if (n > mx) mx = n;
  });
  return mx + 1;
}

// 시트가 수식으로 읽지 않도록 (=, +, -, @ 로 시작하는 글자)
function safe_(v) {
  var s = String(v == null ? "" : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function cleanText_(v, max) {
  var s = String(v == null ? "" : v);
  var out = "";
  for (var i = 0; i < s.length; i++) {
    var code = s.charCodeAt(i);
    out += code < 32 || code === 127 ? " " : s.charAt(i);
  }
  return out.replace(/\s+/g, " ").trim().slice(0, max);
}

function key_(s) {
  return String(s || "").replace(/[\s·‧•・･ᐧ⋅∙ㆍ]/g, "").toLowerCase();
}

function validate_(p) {
  if (!p || typeof p !== "object") fail_("보낸 내용이 비어 있어요.");
  if (p.sid !== SURVEY.id) fail_("조사 화면이 바뀌었어요. 창을 닫고 QR 코드로 다시 들어와 주세요.");
  if (Number(p.grade) !== SURVEY.grade) fail_("학년 정보가 맞지 않아요. QR 코드로 다시 들어와 주세요.");
  var k = classOf_(Number(p.cls));
  if (!k) fail_("반을 다시 골라 주세요.");
  var num = Number(p.num);
  if (!(num >= 1 && num <= k.size && Math.floor(num) === num)) fail_("번호를 다시 골라 주세요.");
  if ((k.skip || []).indexOf(num) >= 0) fail_("결번으로 정해진 번호예요. 내 번호를 다시 확인해 주세요.");
  var name = cleanText_(p.name, 40);
  if (!name || name.length > 20 || /[0-9]/.test(name) || /^[=+\-@'"]/.test(name) || /[<>]/.test(name)) fail_("이름을 다시 확인해 주세요.");
  if (p.agree !== true) fail_("개인정보 안내를 읽고 체크해 주세요.");

  var away = {};
  (Array.isArray(p.away) ? p.away : []).forEach(function (code) {
    SURVEY.semesters.forEach(function (s) { if (s.code === code) away[code] = true; });
  });
  var byGid = {};
  (Array.isArray(p.picks) ? p.picks : []).forEach(function (x) {
    if (x && typeof x.gid === "string") byGid[x.gid] = Array.isArray(x.names) ? x.names : [];
  });

  var rows = [];
  var summary = [];
  var awayLabels = [];
  var active = 0;
  var gradeSeen = {};
  SURVEY.semesters.forEach(function (sem) {
    if (away[sem.code]) {
      awayLabels.push(sem.label);
      rows.push({ semLabel: sem.label, group: "-", subject: AWAY_MARK });
      return;
    }
    active++;
    var semSeen = {};
    sem.groups.forEach(function (g) {
      var allowed = {};
      g.options.forEach(function (o) { allowed[o.name] = true; });
      var uniq = [];
      (byGid[g.id] || []).forEach(function (n) {
        n = String(n);
        if (!allowed[n]) fail_(sem.label + " " + g.label + "에 없는 과목이 들어 있어요: " + n);
        if (uniq.indexOf(n) < 0) uniq.push(n);
      });
      if (g.n != null && uniq.length !== g.n) fail_(sem.label + " " + g.label + ": " + g.n + "개를 골라야 해요 (지금 " + uniq.length + "개).");
      uniq.forEach(function (n) {
        var kk = key_(n);
        if (semSeen[kk]) fail_(sem.label + "에 같은 과목을 두 번 골랐어요: " + n);
        semSeen[kk] = true;
        var gk = sem.code.charAt(0) + "|" + kk;
        if (gradeSeen[gk] && gradeSeen[gk] !== sem.code) fail_("같은 과목을 한 학년에 두 번 고를 수 없어요: " + n);
        gradeSeen[gk] = sem.code;
        rows.push({ semLabel: sem.label, group: g.label, subject: n });
      });
      summary.push(sem.label + " " + g.label + ": " + (uniq.length ? uniq.join(", ") : "없음"));
    });
  });
  if (!active) fail_("모든 학기를 '다른 학교'로 표시할 수는 없어요.");
  return { cls: k.c, num: num, name: name, awayLabels: awayLabels.join(", "), memo: cleanText_(p.memo, 300), rows: rows, summary: summary.join(" / ") };
}

