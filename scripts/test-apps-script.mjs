// 기초조사 Apps Script 코드를 구글 없이 확인하는 테스트 (SpreadsheetApp 등을 흉내 낸 가짜 환경에서 실행)
// 실행: npm run test:survey
import vm from "node:vm";
import assert from "node:assert/strict";
import { buildAppsScriptCode, buildTemplateAppsScriptCode, buildSurveyConfigCode, buildTemplateManifest } from "../src/survey/appsScript.js";

// ---- 프로그램(App.jsx)의 buildSurveyPayload와 같은 모양의 설정 ----
function fnv1aHex(s) {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}
function makePayload(over = {}) {
  const content = {
    id: "sv-test1",
    title: "2025 입학생 기초조사",
    grade: 2,
    term: 1,
    notice: "안내 <b>문구</b>   줄",
    deadline: "3월 20일",
    classes: [{ c: 1, size: 3 }, { c: 2, size: 2, skip: [2] }],
    semesters: [
      {
        code: "1-2",
        label: "1학년 2학기",
        status: "done",
        groups: [
          { id: "1-2#0", label: "선택 A", n: 1, options: [{ name: "물리학", group: "과학", credit: 4 }, { name: "화학", group: "과학", credit: 4 }] },
        ],
      },
      {
        code: "2-1",
        label: "2학년 1학기",
        status: "now",
        groups: [
          { id: "2-1#0", label: "선택 B", n: 2, options: [{ name: "경제", group: "사회", credit: 4 }, { name: "윤리와 사상", group: "사회", credit: 4 }, { name: "세계사", group: "사회", credit: 4 }] },
        ],
      },
    ],
    ...over,
  };
  return { v: 1, ...content, fp: "v-" + fnv1aHex(JSON.stringify(content)), createdAt: "2026-09-28", key: "kEyKEYkeyKEY1234567890ab" };
}

// ---- 가짜 구글 환경 ----
function colNum(a) {
  let n = 0;
  for (const ch of a) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
class Sheet {
  constructor(name) {
    this.name = name;
    this.cells = [];
    this.maxRows = 1000;
    this.hidden = false;
  }
  getLastRow() {
    for (let r = this.cells.length; r >= 1; r--) if ((this.cells[r - 1] || []).some((v) => v !== "" && v != null)) return r;
    return 0;
  }
  getMaxRows() {
    return this.maxRows;
  }
  getName() {
    return this.name;
  }
  setName(n) {
    this.name = n;
    return this;
  }
  setFontSize() {
    return this;
  }
  getLastColumn() {
    return Math.max(0, ...this.cells.map((r) => (r || []).reduce((mx, v, j) => (v !== "" && v != null ? j + 1 : mx), 0)));
  }
  insertRowsAfter(after, n) {
    this.maxRows += n;
  }
  clear() {
    this.cells = [];
  }
  setFrozenRows() {}
  setColumnWidth() {}
  hideSheet() {
    this.hidden = true;
  }
  getRange(r, c, nr = 1, nc = 1) {
    if (typeof r === "string") {
      const m = /^([A-Z]+)(\d+)$/.exec(r);
      return new Range(this, Number(m[2]), colNum(m[1]), 1, 1);
    }
    return new Range(this, r, c, nr, nc);
  }
}
class Range {
  constructor(sh, r, c, nr, nc) {
    Object.assign(this, { sh, r, c, nr, nc });
  }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(((this.sh.cells[this.r - 1 + i] || [])[this.c - 1 + j]) ?? "");
      out.push(row);
    }
    return out;
  }
  getValue() {
    return this.getValues()[0][0];
  }
  setValues(v) {
    assert.equal(v.length, this.nr, "setValues 행 수");
    v.forEach((row, i) => {
      assert.equal(row.length, this.nc, "setValues 열 수");
      if (this.r + i > this.sh.maxRows) throw new Error("시트 행 수를 넘음: " + this.sh.name);
      row.forEach((x, j) => {
        if (typeof x === "string" && x.length > 50000) throw new Error("한 칸 50,000자 초과");
        const R = this.r - 1 + i;
        this.sh.cells[R] = this.sh.cells[R] || [];
        this.sh.cells[R][this.c - 1 + j] = x;
      });
    });
    return this;
  }
  setValue(x) {
    return this.setValues([[x]]);
  }
  setNumberFormat() {
    return this;
  }
  setFontWeight() {
    return this;
  }
  setBackground() {
    return this;
  }
  setDataValidation() {
    return this;
  }
}
function makeEnv({ owner = "teacher@example.com", active = "" } = {}, store) {
  const ss = store.ss;
  const html = [];
  const chain = () => {
    const o = {};
    for (const k of ["requireValueInList", "setAllowInvalid"]) o[k] = () => o;
    o.build = () => ({});
    return o;
  };
  const out = (h) => {
    const o = { html: h, title: null };
    o.setTitle = (t) => ((o.title = t), o);
    o.addMetaTag = () => o;
    o.setWidth = () => o;
    o.setHeight = () => o;
    html.push(o);
    return o;
  };
  return {
    html,
    ctx: {
      SpreadsheetApp: {
        getActiveSpreadsheet: () => ss,
        flush() {},
        newDataValidation: chain,
        getUi: () => ({
          createMenu: () => {
            const m = { addItem: (label, fn) => (store.menu.push([label, fn]), m), addToUi: () => m };
            return m;
          },
          showModalDialog: (o) => store.dialogs.push(o),
          alert: (t, msg) => store.alerts.push(msg),
          ButtonSet: { OK: "OK" },
        }),
      },
      LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
      PropertiesService: {
        getScriptProperties: () => ({
          getProperty: (k) => store.props[k] ?? null,
          setProperty: (k, v) => (store.props[k] = v),
        }),
      },
      Session: {
        getEffectiveUser: () => ({ getEmail: () => owner }),
        getActiveUser: () => ({ getEmail: () => active }),
      },
      HtmlService: { createHtmlOutput: out },
      Utilities: { formatDate: () => "2026-09-28 10:00:00" },
      ScriptApp: { getScriptId: () => "SCRIPT123" },
      ContentService: {
        MimeType: { JSON: "json" },
        createTextOutput: (text) => {
          const o = { text, mime: null };
          o.setMimeType = (mt) => ((o.mime = mt), o);
          return o;
        },
      },
      Logger: { log: (m) => store.logs.push(m) },
    },
  };
}
function newStore(first = "시트1") {
  const sheets = new Map([[first, new Sheet(first)]]);
  const store = {
    props: {},
    menu: [],
    dialogs: [],
    alerts: [],
    logs: [],
    sheets,
    ss: {
      getSheetByName: (n) => [...sheets.values()].find((s) => s.name === n) || null,
      insertSheet: (n) => {
        const s = new Sheet(n);
        sheets.set(n, s);
        return s;
      },
      getSheets: () => [...sheets.values()],
      sheetNames: () => [...sheets.values()].map((s) => s.name),
      deleteSheet: (sh) => {
        if (sheets.size <= 1) throw new Error("마지막 시트는 지울 수 없음");
        for (const [k, v] of sheets) if (v === sh) sheets.delete(k);
      },
      setActiveSheet: (sh) => (store.active = sh.name),
      moveActiveSheet: (pos) => {
        const order = [...sheets.entries()];
        const i = order.findIndex(([n]) => n === store.active);
        const [item] = order.splice(i, 1);
        order.splice(pos - 1, 0, item);
        sheets.clear();
        order.forEach(([n, v]) => sheets.set(n, v));
      },
    },
  };
  return store;
}
// 한 번의 실행(학생 한 명의 접속, 선생님의 메뉴 클릭 등)은 전역 변수가 새로 시작하므로 매번 새 컨텍스트에서 코드를 돌립니다
function run(code, store, who, fn, ...args) {
  const env = makeEnv(who, store);
  const context = vm.createContext({ ...env.ctx, JSON, Math, Array, Object, String, Number, Error });
  vm.runInContext(code, context);
  const result = context[fn](...args);
  return { result, html: env.html };
}
const STUDENT = { active: "" };
const OWNER = { active: "teacher@example.com" };

function validPick(payload, cls = 1, num = 1, name = "김학생") {
  return {
    sid: payload.id,
    grade: payload.grade,
    cls,
    num,
    name,
    agree: true,
    away: [],
    memo: "=메모",
    picks: payload.semesters.flatMap((s) => s.groups.map((g) => ({ gid: g.id, names: g.options.slice(0, g.n).map((o) => o.name) }))),
  };
}
let passed = 0;
function test(name, f) {
  f();
  passed++;
  console.log("✓", name);
}

// ================= 설치 코드(설정이 함께 들어 있는 코드) =================
const A = makePayload();
const codeA = buildAppsScriptCode(A);

test("설치 코드: 붙여넣고 배포만 하면 학생 화면이 열린다 (setup 실행 없이)", () => {
  const st = newStore();
  const { result } = run(codeA, st, STUDENT, "doGet", { parameter: { ban: "2" } });
  assert.equal(result.title, A.title);
  assert.ok(result.html.includes('"ban":2'), "반이 미리 골라짐");
  assert.ok(!result.html.includes("/*__BOOT__*/null"), "BOOT 자리 채워짐");
  for (const n of ["설정", "제출현황", "제출기록", "선택과목", "조사설정"]) assert.ok(st.sheets.get(n), n + " 시트 생성");
  assert.ok(st.sheets.get("조사설정").hidden, "조사설정 시트는 숨김");
  assert.equal(st.sheets.get("설정").getRange("B5").getValue(), A.fp);
  assert.equal(JSON.parse(st.props.BUNDLED_SEEN)[0], A.fp);
});

test("학생 제출이 시트에 적힌다", () => {
  const st = newStore();
  run(codeA, st, STUDENT, "doGet", {});
  const { result } = run(codeA, st, STUDENT, "submitSurvey", validPick(A));
  assert.equal(result.ok, true);
  assert.equal(result.no, 1);
  const log = st.sheets.get("제출기록");
  assert.equal(log.getLastRow(), 2);
  assert.equal(log.getRange(2, 8).getValue(), "'=메모", "수식으로 읽히지 않게");
  assert.equal(st.sheets.get("선택과목").getLastRow(), 1 + 3, "과목 3개");
  const { result: r2 } = run(codeA, st, STUDENT, "submitSurvey", validPick(A, 1, 2, "이학생"));
  assert.equal(r2.no, 2);
});

test("잘못된 제출은 거절된다 (결번, 다른 조사)", () => {
  const st = newStore();
  assert.throws(() => run(codeA, st, STUDENT, "submitSurvey", validPick(A, 2, 2)), /결번/);
  assert.throws(() => run(codeA, st, STUDENT, "submitSurvey", { ...validPick(A), sid: "other" }), /조사 화면이 바뀌었어요/);
});

test("학생(웹 앱 방문자)은 설정을 바꿀 수 없다", () => {
  const st = newStore();
  run(codeA, st, STUDENT, "doGet", {});
  const B = makePayload({ title: "바꿔치기" });
  assert.throws(() => run(codeA, st, STUDENT, "saveSurveyConfig", buildSurveyConfigCode(B), true), /선생님 계정에서만/);
  assert.throws(() => run(codeA, st, { active: "student@school.kr" }, "saveSurveyConfig", buildSurveyConfigCode(B), true), /선생님 계정에서만/);
  assert.throws(() => run(codeA, st, STUDENT, "openConfigDialog"), /선생님 계정에서만/);
});

test("메뉴로 설정을 바꾸면 다시 배포하지 않아도 학생 화면에 반영된다", () => {
  const st = newStore();
  run(codeA, st, STUDENT, "doGet", {});
  const B = makePayload({ classes: [{ c: 1, size: 3 }, { c: 2, size: 2, skip: [2] }, { c: 3, size: 4 }], title: "바뀐 조사" });
  const { result } = run(codeA, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(B), false);
  assert.equal(result.ok, true);
  assert.match(result.summary, /3개 반 \(8명\)/);
  // 배포된 코드(설정 A가 들어 있음)가 다시 실행돼도 A로 되돌아가지 않아야 함
  const { result: page } = run(codeA, st, STUDENT, "doGet", { parameter: { ban: "3" } });
  assert.equal(page.title, "바뀐 조사");
  assert.ok(page.html.includes('"ban":3'));
  assert.equal(st.sheets.get("설정").getRange("B5").getValue(), B.fp);
  assert.equal(st.sheets.get("제출현황").getRange(6, 1).getValue(), 3, "제출현황에 3반 추가");
  const { result: same } = run(codeA, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(B), false);
  assert.equal(same.same, true);
});

test("편집기에 새 설치 코드를 붙여넣기만 하고 배포하지 않아도 번갈아 바뀌지 않는다", () => {
  const st = newStore();
  run(codeA, st, STUDENT, "doGet", {});
  const C = makePayload({ title: "C 조사" });
  const codeC = buildAppsScriptCode(C);
  run(codeC, st, OWNER, "showConfigSummary"); // 편집기의 새 코드(C)가 메뉴에서 실행됨 → C를 가져옴
  assert.match(st.alerts.at(-1), /C 조사/);
  const { result } = run(codeA, st, STUDENT, "doGet", {}); // 배포본은 여전히 A 코드
  assert.equal(result.title, "C 조사", "A를 다시 가져오지 않음");
});

test("잘리거나 손으로 고친 설정 코드는 거절된다", () => {
  const st = newStore();
  const text = buildSurveyConfigCode(A);
  assert.throws(() => run(codeA, st, OWNER, "saveSurveyConfig", text.slice(0, -40), false), /일부가 빠졌습니다|일부가 빠졌거나/);
  assert.throws(() => run(codeA, st, OWNER, "saveSurveyConfig", text.replace("2025 입학생", "2026 입학생"), false), /바뀌었습니다/);
  assert.throws(() => run(codeA, st, OWNER, "saveSurveyConfig", "  ", false), /붙여넣은 내용이 없습니다/);
  assert.throws(() => run(codeA, st, OWNER, "saveSurveyConfig", '{"hello":1}', false), /설정 코드가 아닙니다/);
});

test("다른 조사의 설정은 한 번 더 확인한다", () => {
  const st = newStore();
  run(codeA, st, STUDENT, "doGet", {});
  const D = makePayload({ id: "sv-other", title: "다른 조사" });
  const { result } = run(codeA, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(D), false);
  assert.equal(result.ok, false);
  assert.match(result.confirm, /다른 조사의 설정/);
  const { result: r2 } = run(codeA, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(D), true);
  assert.equal(r2.ok, true);
});

test("큰 설정(한 칸 5만 자 초과)도 나눠서 저장하고 다시 읽는다", () => {
  const st = newStore();
  const options = Array.from({ length: 900 }, (_, i) => ({ name: "과목" + i + "번 아주 긴 이름을 가진 선택 과목", group: "사회", credit: 4 }));
  const big = makePayload({ semesters: [{ code: "2-1", label: "2학년 1학기", status: "now", groups: [{ id: "2-1#0", label: "큰 묶음", n: 1, options }] }] });
  assert.ok(buildSurveyConfigCode(big).length > 50000);
  run(codeA, st, STUDENT, "doGet", {});
  run(codeA, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(big), false);
  assert.ok(st.sheets.get("조사설정").getLastRow() >= 3, "여러 칸에 나눠 적음");
  const { result } = run(codeA, st, STUDENT, "submitSurvey", validPick(big));
  assert.equal(result.ok, true);
});

// ================= 템플릿 코드(설정 없이 메뉴로 받는 코드) =================
const codeT = buildTemplateAppsScriptCode();

test("템플릿: 시트를 열면 [📋 기초조사] 메뉴가 생긴다", () => {
  const st = newStore();
  run(codeT, st, OWNER, "onOpen");
  assert.deepEqual(st.menu.map((m) => m[1]), ["openConfigDialog", "openDeployDialog", "showConfigSummary"]);
  run(codeT, st, OWNER, "openConfigDialog");
  assert.ok(st.dialogs[0].html.includes("saveSurveyConfig"));
  assert.ok(st.dialogs[0].html.includes("</script>"), "설정 창의 스크립트 태그가 제대로 닫힘");
});

test("템플릿: 설정 전에는 '준비 중' 화면, 설정을 넣으면 바로 열린다", () => {
  const st = newStore();
  const { result } = run(codeT, st, STUDENT, "doGet", {});
  assert.match(result.html, /아직 조사 준비 중이에요/);
  assert.throws(() => run(codeT, st, STUDENT, "submitSurvey", validPick(A)), /아직 조사가 준비되지 않았어요/);
  const { result: saved } = run(codeT, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(A), false);
  assert.equal(saved.ok, true);
  const { result: page } = run(codeT, st, STUDENT, "doGet", {});
  assert.equal(page.title, A.title);
  assert.equal(run(codeT, st, STUDENT, "submitSurvey", validPick(A)).result.ok, true);
});

test("학생 화면에 넣는 설정은 HTML을 깨뜨리지 않는다", () => {
  const st = newStore();
  const { result } = run(codeA, st, STUDENT, "doGet", {});
  const inline = result.html.slice(result.html.indexOf("<script>"));
  assert.ok(!/<b>문구/.test(inline), "< 가 \\u003c 로 바뀜");
});

test("조사를 준비하면 빈 시트1과 안내 시트(시작하기)를 지우고 제출현황을 맨 앞에 둔다", () => {
  const st = newStore();
  run(codeT, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(A), false);
  const names = [...st.sheets.keys()];
  assert.ok(!names.includes("시트1"), "빈 시트1 삭제");
  assert.equal(names[0], "제출현황");
  const st2 = newStore("시작하기");
  st2.sheets.get("시작하기").getRange(1, 1).setValue("위쪽 메뉴 [기초조사 → 설정 붙여넣기]를 쓰세요");
  run(codeT, st2, OWNER, "saveSurveyConfig", buildSurveyConfigCode(A), false);
  assert.ok(!st2.sheets.has("시작하기"), "안내 시트는 내용이 있어도 삭제");
  const st3 = newStore();
  run(codeA, st3, STUDENT, "doGet", {});
  assert.ok(!st3.sheets.has("시트1"), "설치 코드 방식: 학생 첫 접속 때 삭제");
});

test("선생님이 무언가 적어 둔 시트1은 지우지 않는다", () => {
  const st = newStore();
  st.sheets.get("시트1").getRange(3, 2).setValue("메모");
  run(codeT, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(A), false);
  assert.ok(st.sheets.has("시트1"));
  assert.ok(st.sheets.get("조사설정").hidden, "조사설정은 여전히 숨김");
});


test("템플릿: 사본을 처음 열면 빈 시트1이 ‘시작하기’ 안내 시트가 되고, 조사 준비 뒤엔 사라진다", () => {
  const st = newStore();
  run(codeT, st, OWNER, "onOpen");
  assert.deepEqual(st.ss.sheetNames(), ["시작하기"]);
  const start = st.ss.getSheetByName("시작하기");
  assert.match(String(start.getRange(2, 1).getValue()), /설정 코드 복사/);
  run(codeT, st, OWNER, "onOpen"); // 두 번 열어도 그대로
  assert.equal(start.getLastRow(), 5);
  run(codeT, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(A), false);
  assert.ok(!st.ss.getSheetByName("시작하기"), "조사 준비 뒤 삭제");
  run(codeT, st, OWNER, "onOpen"); // 준비된 뒤에는 다시 만들지 않음
  assert.ok(!st.ss.getSheetByName("시작하기"));
  assert.deepEqual(st.menu.slice(-3).map((x) => x[1]), ["openConfigDialog", "openDeployDialog", "showConfigSummary"]);
});

test("설정을 저장하면 같은 창에서 배포 안내와 편집기 주소를 이어서 준다", () => {
  const st = newStore();
  const { result } = run(codeT, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(A), false);
  assert.equal(result.editorUrl, "https://script.google.com/d/SCRIPT123/edit");
  run(codeT, st, OWNER, "openConfigDialog");
  const html = st.dialogs[0].html;
  assert.ok(html.includes("var DG_HTML = ") && html.includes("Apps Script 편집기 열기") && html.includes("모든 사용자"));
  assert.ok(!html.includes("__DG_"), "자리표시자가 남지 않음");
  run(codeT, st, OWNER, "openDeployDialog");
  const dep = st.dialogs[1].html;
  assert.ok(dep.includes('href = "https://script.google.com/d/SCRIPT123/edit"') && dep.includes('class="dg"'));
  assert.ok(!dep.includes("__EDITOR__") && !dep.includes("__DG_"));
  assert.throws(() => run(codeT, st, STUDENT, "openDeployDialog"), /선생님 계정에서만/);
});

test("템플릿 매니페스트: 배포 창 기본값만 있고 권한 목록은 없다", () => {
  const m = JSON.parse(buildTemplateManifest());
  assert.deepEqual(m.webapp, { executeAs: "USER_DEPLOYING", access: "ANYONE_ANONYMOUS" });
  assert.equal(m.oauthScopes, undefined);
});


// ================= 응답 내보내기 (프로그램이 시트에서 바로 가져오기) =================
test("내보내기: 열쇠가 맞을 때만 시트 내용을 엑셀과 같은 모양으로 돌려준다", () => {
  const st = newStore();
  run(codeA, st, STUDENT, "doGet", {});
  run(codeA, st, STUDENT, "submitSurvey", validPick(A));
  const { result } = run(codeA, st, STUDENT, "doGet", { parameter: { export: A.key } });
  assert.equal(result.mime, "json");
  const json = JSON.parse(result.text);
  assert.equal(json.kind, "graduation-survey-export");
  assert.deepEqual(Object.keys(json.sheets), ["선택과목", "제출기록", "설정"]);
  assert.deepEqual(json.sheets["선택과목"][0], ["제출번호", "제출시각", "학년", "반", "번호", "이름", "학기", "선택그룹", "과목"]);
  assert.equal(json.sheets["선택과목"].length, 1 + 3, "머리글 + 과목 3개");
  assert.equal(json.sheets["제출기록"][1][5], "김학생");
  assert.equal(json.sheets["설정"][4][1], A.fp, "설정 시트의 설정 버전");
  assert.equal(json.closed, false);
});

test("내보내기: 열쇠가 틀리거나 없으면 응답을 주지 않는다", () => {
  const st = newStore();
  run(codeA, st, STUDENT, "doGet", {});
  for (const bad of ["", "wrong", A.key.slice(0, -1)]) {
    const json = JSON.parse(run(codeA, st, STUDENT, "doGet", { parameter: { export: bad } }).result.text);
    assert.ok(json.error && !json.sheets, "거절: " + JSON.stringify(bad));
  }
  // 열쇠 없는 예전 설정 코드로 설치한 시트도 거절
  const noKey = { ...A };
  delete noKey.key;
  const st2 = newStore();
  run(codeT, st2, OWNER, "saveSurveyConfig", JSON.stringify(noKey), false);
  const json2 = JSON.parse(run(codeT, st2, STUDENT, "doGet", { parameter: { export: A.key } }).result.text);
  assert.ok(json2.error);
});

test("내보내기: 학생 화면과 설정 시트에는 열쇠가 나가지 않는다", () => {
  const st = newStore();
  const { result } = run(codeA, st, STUDENT, "doGet", { parameter: { ban: "1" } });
  assert.ok(!result.html.includes(A.key), "학생 화면 HTML");
  assert.ok(result.html.includes('"id":"' + A.id + '"'), "나머지 설정은 그대로");
  const settings = st.sheets.get("설정").getRange(9, 2).getValue();
  assert.ok(!String(settings).includes(A.key), "설정 시트의 조사 정보");
  assert.ok(String(st.sheets.get("조사설정").getRange(2, 1).getValue()).includes(A.key), "숨긴 조사설정 시트에는 있음 (선생님 시트)");
});

test("내보내기: 모양이 이상한 열쇠는 없는 것으로 친다", () => {
  const st = newStore();
  run(codeT, st, OWNER, "saveSurveyConfig", JSON.stringify({ ...A, key: "short" }), false);
  const json = JSON.parse(run(codeT, st, STUDENT, "doGet", { parameter: { export: "short" } }).result.text);
  assert.ok(json.error);
});


test("내보내기: 같은 설정에 열쇠만 새로 붙여넣어도 열쇠가 저장된다", () => {
  const noKey = { ...A };
  delete noKey.key;
  const st = newStore();
  run(codeT, st, OWNER, "saveSurveyConfig", JSON.stringify(noKey), false);
  assert.ok(JSON.parse(run(codeT, st, STUDENT, "doGet", { parameter: { export: A.key } }).result.text).error, "아직 열쇠 없음");
  const { result } = run(codeT, st, OWNER, "saveSurveyConfig", buildSurveyConfigCode(A), false);
  assert.equal(result.same, true);
  assert.ok(JSON.parse(run(codeT, st, STUDENT, "doGet", { parameter: { export: A.key } }).result.text).sheets, "열쇠 저장됨");
});

console.log(`\n${passed}개 통과`);
