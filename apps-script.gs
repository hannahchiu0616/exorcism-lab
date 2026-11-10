// 驅魔實驗室後端
// 資料存在部署者自己 Google 雲端硬碟裡的一份試算表。打卡等內容公開；留言和修改規則要用姓名和密碼登入。
// 安裝：
//   1. 到 script.google.com 建立新專案，貼上這整段，存檔
//   2. 上方函式選 setup → 執行 → 授權
//   3. 部署 → 新增部署作業 → 網頁應用程式 → 執行身分「我」、存取權「所有人」→ 部署 → 複製網址

const ADMIN = "邱子芸";
// 實驗者名單存在設定裡，邱子芸可以從網頁的「規則」分頁新增；這是還沒存過名單時的預設值
const DEFAULT_PEOPLE = [{ name: "邱子芸", joined: null }, { name: "江宇熙", joined: null }, { name: "邱家瑋", joined: "2026-11-09" }];
let PEOPLE = DEFAULT_PEOPLE.map(p => p.name); // 每次請求開始時由 book() 從設定讀入
const DEFAULT_OBSERVERS = ["邱美琪", "張克強"];
const NEVER = "2000-01-01"; // 定課的 by[人].to 設成這天，代表這個人從沒做過

const TABS = {
  "設定": ["項目", "值"],
  "打卡": ["日期", "完成項目", "實驗者"],
  "歷程": ["編號", "日期", "能量", "內容", "建立時間", "修改時間", "實驗者"],
  "留言": ["編號", "內容", "建立時間", "身分", "實驗者"],
};

const props = () => PropertiesService.getScriptProperties();

/* ---------- 登入 ---------- */
const sign = (value, key) => Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(value, key, Utilities.Charset.UTF_8));
function secret() {
  let s = props().getProperty("secret");
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props().setProperty("secret", s); }
  return s;
}
function hashPassword(password, salt) {
  let h = password;
  for (let i = 0; i < 1000; i++) h = sign(h, salt);
  return h;
}
function setPassword(name, password) {
  if (String(password).length < 4) throw appError("invalid", "密碼至少 4 個字");
  const salt = Utilities.getUuid();
  props().setProperty("pw:" + name, salt + ":" + hashPassword(String(password), salt));
}
// 權杖綁定目前的密碼，改密碼後舊的登入就失效
function makeToken(name, exp) {
  return exp + "." + sign(name + "|" + exp, secret() + props().getProperty("pw:" + name));
}
function observers(ss) {
  const row = rows(ss, "設定").filter(r => r[0] === "實驗")[0];
  const cfg = row ? parse(row[1], null) : null;
  return cfg && cfg.observers && cfg.observers.length ? cfg.observers : DEFAULT_OBSERVERS;
}
const users = ss => PEOPLE.concat(observers(ss));
function whoIs(ss, auth) {
  if (!auth || !auth.name || !auth.token) return "";
  const name = String(auth.name), exp = Number(String(auth.token).split(".")[0]);
  if (users(ss).indexOf(name) < 0 || !props().getProperty("pw:" + name) || !(exp > Date.now())) return "";
  return makeToken(name, exp) === auth.token ? name : "";
}
function makeSession(name) {
  const exp = Date.now() + 120 * 86400000;
  return { name: name, token: makeToken(name, exp) };
}
function login(ss, p) {
  const name = String(p.name || "").trim().slice(0, 40), password = String(p.password || "");
  const cache = CacheService.getScriptCache(), key = "fail:" + name, fails = Number(cache.get(key)) || 0;
  if (fails >= 5) throw appError("invalid", "試太多次了，請 10 分鐘後再試");
  const bad = () => { cache.put(key, String(fails + 1), 600); return appError("invalid", "姓名或密碼不正確"); };
  if (users(ss).indexOf(name) < 0) throw bad();
  const saved = props().getProperty("pw:" + name);
  if (saved) {
    const [salt, hash] = saved.split(":");
    if (hashPassword(password, salt) !== hash) throw bad();
  } else {
    setPassword(name, password); // 第一次登入時自己設定密碼
  }
  cache.remove(key);
  return makeSession(name);
}
// 每則留言只有寫的人看得到，被留言的驅魔勇士看不到
const canSeeNote = (n, me) => !!me && n.role === me;

/* ---------- 安裝 ---------- */
function setup() {
  const p = props();
  if (!p.getProperty("sheetId")) {
    const ss = SpreadsheetApp.create("驅魔實驗室紀錄");
    p.setProperty("sheetId", ss.getId());
    Object.keys(TABS).forEach(name => tab(ss, name));
    const first = ss.getSheetByName("工作表1") || ss.getSheetByName("Sheet1");
    if (first) ss.deleteSheet(first);
  }
  Logger.log("試算表：" + SpreadsheetApp.openById(p.getProperty("sheetId")).getUrl());
}

/* ---------- 試算表 ---------- */
function book() {
  const id = props().getProperty("sheetId");
  if (!id) throw appError("not_setup", "後端還沒安裝，請先執行 setup");
  const ss = SpreadsheetApp.openById(id);
  PEOPLE = cleanPeople((readCfg(ss) || {}).people).map(p => p.name);
  return ss;
}
function tab(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), TABS[name].length).setNumberFormat("@");
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  const head = TABS[name];
  // 舊版試算表少了後來新增的欄位，補上標題
  if (String(sh.getRange(1, head.length).getValue()) !== head[head.length - 1]) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
  }
  return sh;
}
function rows(ss, name) {
  const sh = tab(ss, name), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, TABS[name].length).getValues()
    .map(r => r.map(x => String(x)))
    .filter(r => r[0].trim());
}
function findRow(sh, key) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const col = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) if (String(col[i][0]) === key) return i + 2;
  return 0;
}
// 文字格式寫入；內容開頭的 = 由網頁端加上單引號，避免被當成公式
function put(sh, row, values) {
  sh.getRange(row, 1, 1, values.length).setNumberFormat("@").setValues([values.map(String)]);
}
function findCheckin(sh, date, who) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const col = sh.getRange(2, 1, n - 1, 3).getValues();
  for (let i = 0; i < col.length; i++) {
    if (String(col[i][0]) === date && (String(col[i][2]) || ADMIN) === who) return i + 2;
  }
  return 0;
}
function upsert(sh, key, values) {
  put(sh, findRow(sh, key) || sh.getLastRow() + 1, values);
}

function snapshot(ss, me) {
  const settings = {};
  rows(ss, "設定").forEach(r => { settings[r[0]] = r[1]; });
  const by = {};
  PEOPLE.forEach(name => { by[name] = {}; });
  rows(ss, "打卡").forEach(r => {
    const who = r[2] || ADMIN;
    (by[who] = by[who] || {})[r[0]] = { done: parse(r[1], {}) };
  });
  const cfg = parse(settings["實驗"], null);
  if (cfg) cfg.observers = observers(ss);
  return {
    cfg: cfg,
    checkins: by[ADMIN],
    checkinsBy: by,
    journal: rows(ss, "歷程").map(r => ({
      id: r[0], date: r[1], energy: r[2] ? Number(r[2]) : null, body: r[3],
      createdAt: Number(r[4]) || 0, editedAt: Number(r[5]) || 0, person: r[6] || ADMIN,
    })).filter(e => e.person === me),
    notes: rows(ss, "留言").map(r => ({ id: r[0], body: r[1], createdAt: Number(r[2]) || 0, role: r[3], person: r[4] || ADMIN }))
      .filter(n => canSeeNote(n, me)),
    me: me || null,
    hasPassword: me === ADMIN ? users(ss).filter(n => props().getProperty("pw:" + n)) : [],
  };
}
function parse(s, fallback) { try { return JSON.parse(s); } catch (e) { return fallback; } }

/* ---------- 檢查 ---------- */
function appError(code, message) { const e = new Error(message); e.appCode = code; return e; }
const isDate = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ""));
const person = s => (PEOPLE.indexOf(s) >= 0 ? s : ADMIN);
// 打卡期限：隔天中午 12 點（台灣時間，等於隔天的 UTC 04:00）
function pastDeadline(date) {
  const [y, m, d] = date.split("-").map(Number);
  return Date.now() >= Date.UTC(y, m - 1, d + 1, 4);
}
function text(s, max) {
  const t = String(s == null ? "" : s).trim().slice(0, max);
  if (!t) throw appError("invalid", "內容不能是空的");
  return t;
}
function cleanRest(r) {
  return {
    restDays: (Array.isArray(r.restDays) ? r.restDays : []).filter(isDate).slice(0, 366),
    restReasons: Object.keys(r.restReasons || {}).filter(isDate).slice(0, 366).reduce((o, d) => { o[d] = String(r.restReasons[d]).slice(0, 20); return o; }, {}),
  };
}
// 邱子芸固定排第一位；名字不能重複
function cleanPeople(list) {
  const src = Array.isArray(list) && list.length ? list : DEFAULT_PEOPLE;
  const seen = {}, out = [{ name: ADMIN, joined: null }];
  seen[ADMIN] = true;
  src.forEach(p => {
    const n = String((p && p.name) || "").trim().slice(0, 20);
    if (n && !seen[n]) { seen[n] = true; out.push({ name: n, joined: isDate(p.joined) ? p.joined : null }); }
  });
  return out.slice(0, 20);
}
function cleanBy(b, names) {
  return Object.keys(b || {}).filter(n => names.indexOf(n) >= 0 && b[n]).reduce((o, n) => {
    o[n] = { from: isDate(b[n].from) ? b[n].from : null, to: isDate(b[n].to) ? b[n].to : null };
    return o;
  }, {});
}
function cleanCfg(c) {
  if (!c || !isDate(c.startDate)) throw appError("invalid", "開始日期不正確");
  const days = Math.round(Number(c.days));
  if (!(days >= 1 && days <= 365)) throw appError("invalid", "實驗天數請填 1 到 365");
  if (!Array.isArray(c.practices) || c.practices.length > 100) throw appError("invalid", "定課清單不正確");
  const r = c.rules || {}, people = cleanPeople(c.people), names = people.map(p => p.name);
  return {
    startDate: c.startDate, days: days, people: people,
    observers: (Array.isArray(c.observers) ? c.observers : []).map(n => String(n).trim().slice(0, 20)).filter(n => n && names.indexOf(n) < 0).slice(0, 20),
    rules: {
      weekdays: (Array.isArray(r.weekdays) ? r.weekdays : [1, 2, 3, 4, 5]).map(Number).filter(d => d >= 0 && d <= 6),
      ...cleanRest(r),
      restBy: Object.keys(r.restBy || {}).filter(n => names.indexOf(n) >= 0).reduce((o, n) => { o[n] = cleanRest(r.restBy[n]); return o; }, {}),
      text: String(r.text == null ? "" : r.text).slice(0, 500),
    },
    practices: c.practices.map(p => ({
      id: text(p.id, 40), name: text(p.name, 60), note: String(p.note || "").slice(0, 120),
      addedOn: isDate(p.addedOn) ? p.addedOn : null, removedOn: isDate(p.removedOn) ? p.removedOn : null,
      by: cleanBy(p.by, names),
    })),
  };
}
function cleanDone(d) {
  const o = {};
  Object.keys(d || {}).slice(0, 100).forEach(k => { if (d[k]) o[String(k).slice(0, 40)] = true; });
  return o;
}

/* ---------- 網址入口 ---------- */
function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

const VERSION = "2026-11-10 勇士留言"; // 打開網頁應用程式網址就看得到，用來確認部署的是哪一版
function doGet() { return out({ ok: true, service: "exorcism-lab", version: VERSION }); }

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.action === "getAll") { const ss = book(); return out({ ok: true, data: snapshot(ss, whoIs(ss, p.auth)) }); }
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = book();
      if (p.action === "login") {
        const session = login(ss, p);
        return out({ ok: true, session: session, data: snapshot(ss, session.name) });
      }
      if (p.action === "changePassword") {
        const name = whoIs(ss, p.auth);
        if (!name) throw appError("invalid", "請先登入");
        setPassword(name, p.password);
        return out({ ok: true, session: makeSession(name), data: snapshot(ss, name) });
      }
      applyWrite(ss, p, whoIs(ss, p.auth));
      SpreadsheetApp.flush();
      return out({ ok: true, data: snapshot(ss, whoIs(ss, p.auth)) });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out({ ok: false, code: err.appCode || "server", error: err.message });
  }
}

// 心路歷程只有實驗者本人能寫和看
function ownJournalRow(sh, id, me) {
  if (PEOPLE.indexOf(me) < 0) throw appError("invalid", "請用驅魔勇士本人的帳號登入");
  const row = findRow(sh, String(id));
  if (row && (String(sh.getRange(row, 7).getValue()) || ADMIN) !== me) throw appError("invalid", "只能修改自己的紀錄");
  return row;
}

function readCfg(ss) {
  const row = rows(ss, "設定").filter(r => r[0] === "實驗")[0];
  return row ? parse(row[1], null) : null;
}
// 每位實驗者只改自己的定課；邱子芸也可以代改
function saveMenu(ss, p, me) {
  const who = p.person;
  if (PEOPLE.indexOf(who) < 0) throw appError("invalid", "找不到這位驅魔勇士");
  if (me !== who && me !== ADMIN) throw appError("invalid", "只能修改自己的定課");
  const cfg = readCfg(ss);
  if (!cfg) throw appError("invalid", "請先設定實驗期間");
  if (p.add) {
    const by = {};
    PEOPLE.filter(n => n !== who).forEach(n => { by[n] = { from: null, to: NEVER }; });
    cfg.practices.push({ id: p.add.id, name: p.add.name, note: p.add.note, addedOn: p.add.addedOn, removedOn: null, by: by });
  } else {
    const item = cfg.practices.filter(x => x.id === p.id)[0];
    if (!item) throw appError("invalid", "找不到這項定課");
    item.by = item.by || {};
    if (p.remove) {
      if (PEOPLE.some(n => !(item.by[n] && item.by[n].to))) throw appError("invalid", "還有人在做這項定課");
      item.removedOn = new Date().toISOString().slice(0, 10);
    } else {
      item.by[who] = p.entry;
    }
  }
  upsert(tab(ss, "設定"), "實驗", ["實驗", JSON.stringify(cleanCfg(cfg))]);
}

function applyWrite(ss, p, me) {
  switch (p.action) {
    case "saveConfig":
      if (me !== ADMIN) throw appError("invalid", "只有邱子芸可以修改規則");
      {
        // 定課由 saveMenu 管，這裡保留試算表裡的版本，才不會蓋掉別人剛改的菜單
        const cur = readCfg(ss), next = cleanCfg({ ...p.cfg, practices: cur ? cur.practices : p.cfg.practices });
        const names = next.people.map(x => x.name), used = {};
        rows(ss, "打卡").forEach(r => { used[r[2] || ADMIN] = true; });
        const gone = PEOPLE.filter(n => names.indexOf(n) < 0 && used[n]);
        if (gone.length) throw appError("invalid", gone.join("、") + " 已經有打卡紀錄，不能移除");
        // 新加入的實驗者從空白菜單開始，自己選要做的定課
        names.filter(n => PEOPLE.indexOf(n) < 0).forEach(n => next.practices.forEach(x => { x.by[n] = { from: null, to: NEVER }; }));
        upsert(tab(ss, "設定"), "實驗", ["實驗", JSON.stringify(next)]);
        PEOPLE = names;
      }
      return;
    case "saveMenu":
      saveMenu(ss, p, me);
      return;
    case "setPassword":
      if (me !== ADMIN) throw appError("invalid", "只有邱子芸可以設定密碼");
      if (users(ss).indexOf(p.name) < 0) throw appError("invalid", "找不到這個人");
      setPassword(p.name, p.password);
      return;
    case "clearPassword":
      if (me !== ADMIN) throw appError("invalid", "只有邱子芸可以清除密碼");
      if (p.name === ADMIN) throw appError("invalid", "不能清除邱子芸自己的密碼");
      props().deleteProperty("pw:" + p.name);
      return;
    case "setCheckin":
      if (!isDate(p.date)) throw appError("invalid", "日期不正確");
      if (pastDeadline(p.date)) throw appError("locked", "這天已鎖定，只能從後台修改。");
      {
        const sh = tab(ss, "打卡"), who = person(p.person);
        put(sh, findCheckin(sh, p.date, who) || sh.getLastRow() + 1, [p.date, JSON.stringify(cleanDone(p.done)), who]);
      }
      return;
    case "saveJournal": {
      if (!isDate(p.date)) throw appError("invalid", "日期不正確");
      const sh = tab(ss, "歷程"), row = ownJournalRow(sh, p.id, me), body = text(p.body, 5000), now = Date.now();
      if (p.id) {
        if (!row) throw appError("invalid", "找不到這筆紀錄");
        const created = String(sh.getRange(row, 5).getValue());
        put(sh, row, [p.id, p.date, "", body, created, now, me]);
      } else {
        put(sh, sh.getLastRow() + 1, [Utilities.getUuid(), p.date, "", body, now, "", me]);
      }
      return;
    }
    case "deleteJournal": {
      const sh = tab(ss, "歷程"), row = ownJournalRow(sh, p.id, me);
      if (row) sh.deleteRow(row);
      return;
    }
    case "addNote": {
      if (!me) throw appError("invalid", "請先登入再留言");
      if (person(p.person) === me) throw appError("invalid", "不能留言給自己");
      const sh = tab(ss, "留言");
      put(sh, sh.getLastRow() + 1, [Utilities.getUuid(), text(p.body, 3000), Date.now(), me, person(p.person)]);
      return;
    }
    case "deleteNote": {
      const sh = tab(ss, "留言"), row = findRow(sh, String(p.id));
      if (!row) return;
      const r = sh.getRange(row, 1, 1, 5).getValues()[0];
      if (!canSeeNote({ role: String(r[3]), person: String(r[4]) || ADMIN }, me)) throw appError("invalid", "只能刪除自己的留言");
      sh.deleteRow(row);
      return;
    }
    default:
      throw appError("invalid", "不支援的動作");
  }
}
