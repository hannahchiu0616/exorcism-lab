// 驅魔實驗室後端
// 資料存在部署者自己 Google 雲端硬碟裡的一份試算表；通行碼存在指令碼屬性，不在這份程式碼裡。
// 安裝：
//   1. 到 script.google.com 建立新專案，貼上這整段，存檔
//   2. 上方函式選 setup → 執行 → 授權 → 從執行記錄抄下兩組通行碼
//   3. 部署 → 新增部署作業 → 網頁應用程式 → 執行身分「我」、存取權「所有人」→ 部署 → 複製網址
// 忘記通行碼：執行 setup 會再顯示一次。要換通行碼：執行 resetCodes。

const TABS = {
  "設定": ["項目", "值"],
  "打卡": ["日期", "完成項目"],
  "歷程": ["編號", "日期", "能量", "內容", "建立時間", "修改時間"],
  "留言": ["編號", "內容", "建立時間", "身分"],
};
const ALPHABET = "ABCDEFGHJKMNPQRSTUVWXYZ23456789";
const MAX_FAILS = 20;

const props = () => PropertiesService.getScriptProperties();

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
  if (!p.getProperty("codeExperimenter")) resetCodes();
  else printCodes();
}
function resetCodes() {
  props().setProperty("codeExperimenter", newCode());
  props().setProperty("codeObserver", newCode());
  printCodes();
}
function printCodes() {
  Logger.log("實驗者通行碼：" + props().getProperty("codeExperimenter"));
  Logger.log("觀察者通行碼：" + props().getProperty("codeObserver"));
}
function newCode() {
  let s = "";
  for (let i = 0; i < 10; i++) s += ALPHABET[Math.floor(Math.random() * ALPHABET.length)];
  return s;
}

/* ---------- 試算表 ---------- */
function book() {
  const id = props().getProperty("sheetId");
  if (!id) throw appError("not_setup", "後端還沒安裝，請先執行 setup");
  return SpreadsheetApp.openById(id);
}
function tab(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), TABS[name].length).setNumberFormat("@");
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight("bold");
    sh.setFrozenRows(1);
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
// 文字格式寫入：內容開頭是 = 也不會被當成公式
function put(sh, row, values) {
  sh.getRange(row, 1, 1, values.length).setNumberFormat("@").setValues([values.map(String)]);
}
function upsert(sh, key, values) {
  put(sh, findRow(sh, key) || sh.getLastRow() + 1, values);
}
function removeRow(sh, key) {
  const row = findRow(sh, key);
  if (row) sh.deleteRow(row);
}

function snapshot(ss) {
  const settings = {};
  rows(ss, "設定").forEach(r => { settings[r[0]] = r[1]; });
  const checkins = {};
  rows(ss, "打卡").forEach(r => { checkins[r[0]] = { done: parse(r[1], {}) }; });
  return {
    cfg: parse(settings["實驗"], null),
    checkins: checkins,
    journal: rows(ss, "歷程").map(r => ({
      id: r[0], date: r[1], energy: r[2] ? Number(r[2]) : null, body: r[3],
      createdAt: Number(r[4]) || 0, editedAt: Number(r[5]) || 0,
    })),
    notes: rows(ss, "留言").map(r => ({ id: r[0], body: r[1], createdAt: Number(r[2]) || 0, role: r[3] })),
  };
}
function parse(s, fallback) { try { return JSON.parse(s); } catch (e) { return fallback; } }

/* ---------- 檢查 ---------- */
function appError(code, message) { const e = new Error(message); e.appCode = code; return e; }
const isDate = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ""));
function text(s, max) {
  const t = String(s == null ? "" : s).trim().slice(0, max);
  if (!t) throw appError("invalid", "內容不能是空的");
  return t;
}
function cleanCfg(c) {
  if (!c || !isDate(c.startDate)) throw appError("invalid", "開始日期不正確");
  const days = Math.round(Number(c.days));
  if (!(days >= 1 && days <= 365)) throw appError("invalid", "實驗天數請填 1 到 365");
  if (!Array.isArray(c.practices) || c.practices.length > 100) throw appError("invalid", "訂課清單不正確");
  return {
    startDate: c.startDate, days: days,
    practices: c.practices.map(p => ({
      id: text(p.id, 40), name: text(p.name, 60), note: String(p.note || "").slice(0, 120),
      addedOn: isDate(p.addedOn) ? p.addedOn : null, removedOn: isDate(p.removedOn) ? p.removedOn : null,
    })),
  };
}
function cleanDone(d) {
  const o = {};
  Object.keys(d || {}).slice(0, 100).forEach(k => { if (d[k]) o[String(k).slice(0, 40)] = true; });
  return o;
}
function roleOf(code) {
  const p = props(), c = String(code || "").trim().toUpperCase();
  if (c && c === p.getProperty("codeExperimenter")) return "experimenter";
  if (c && c === p.getProperty("codeObserver")) return "observer";
  return "";
}
// 擋亂試通行碼：全站 10 分鐘內錯太多次就暫停
function checkRole(code) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("fails") || 0);
  if (fails >= MAX_FAILS) throw appError("too_many", "嘗試太多次，請 10 分鐘後再試");
  const role = roleOf(code);
  if (!role) {
    cache.put("fails", String(fails + 1), 600);
    throw appError("bad_code", "通行碼不正確");
  }
  return role;
}

/* ---------- 網址入口 ---------- */
function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function doGet() { return out({ ok: true, service: "exorcism-lab" }); } // 只回報服務在線，不含紀錄

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const role = checkRole(p.code);
    if (p.action === "getAll") return out({ ok: true, role: role, data: snapshot(book()) });
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = book();
      applyWrite(ss, p, role);
      SpreadsheetApp.flush();
      return out({ ok: true, role: role, data: snapshot(ss) });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out({ ok: false, code: err.appCode || "server", error: err.message });
  }
}

function applyWrite(ss, p, role) {
  const exp = role === "experimenter";
  const expOnly = () => { if (!exp) throw appError("forbidden", "這個動作只有實驗者能做"); };
  switch (p.action) {
    case "saveConfig":
      expOnly();
      upsert(tab(ss, "設定"), "實驗", ["實驗", JSON.stringify(cleanCfg(p.cfg))]);
      return;
    case "setCheckin":
      expOnly();
      if (!isDate(p.date)) throw appError("invalid", "日期不正確");
      upsert(tab(ss, "打卡"), p.date, [p.date, JSON.stringify(cleanDone(p.done))]);
      return;
    case "saveJournal": {
      expOnly();
      if (!isDate(p.date)) throw appError("invalid", "日期不正確");
      const sh = tab(ss, "歷程"), body = text(p.body, 5000), now = Date.now();
      const energy = [1, 2, 3, 4, 5].indexOf(Number(p.energy)) >= 0 ? Number(p.energy) : "";
      if (p.id) {
        const row = findRow(sh, String(p.id));
        if (!row) throw appError("invalid", "找不到這筆紀錄");
        const created = String(sh.getRange(row, 5).getValue());
        put(sh, row, [p.id, p.date, energy, body, created, now]);
      } else {
        put(sh, sh.getLastRow() + 1, [Utilities.getUuid(), p.date, energy, body, now, ""]);
      }
      return;
    }
    case "deleteJournal":
      expOnly();
      removeRow(tab(ss, "歷程"), String(p.id));
      return;
    case "addNote":
      put(tab(ss, "留言"), tab(ss, "留言").getLastRow() + 1, [Utilities.getUuid(), text(p.body, 3000), Date.now(), role]);
      return;
    case "deleteNote": {
      const sh = tab(ss, "留言"), row = findRow(sh, String(p.id));
      if (!row) return;
      if (!exp && String(sh.getRange(row, 4).getValue()) !== role) throw appError("forbidden", "只能刪除自己的留言");
      sh.deleteRow(row);
      return;
    }
    default:
      throw appError("invalid", "不支援的動作");
  }
}
