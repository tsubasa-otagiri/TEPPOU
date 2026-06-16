import { useState, useEffect, useRef, useCallback, useMemo, memo } from "react";
import * as XLSX from "xlsx";

// ── Constants ──────────────────────────────────────────────────────────────────
const PASSWORD     = "0227";
const STORAGE_KEY  = "teppou_records_v3";   // legacy migration source
const SETTINGS_KEY = "teppou_settings_v1";
const PAGE_SIZE    = 100;
const API_BASE        = (import.meta.env.VITE_API_URL ?? "").replace(/\/$/, "");
const PAST_DEALS_KEY  = "teppou_past_deals_v1";   // 過去商談プル照合用
const PAST_MGMT_KEY   = "teppou_past_mgmt_v1";    // 過去商談管理（再アプローチ）
const UI_KEY          = "teppou_ui_v1";           // 列設定など UI 状態
const PAST_UI_KEY     = "teppou_past_ui_v1";      // 過去商談の UI 状態
const CLICK_REFRESH_COOLDOWN = 30_000; // クリック更新のクールダウン（30秒）

// ── API クライアント ────────────────────────────────────────────────────────────
async function apiGet(resource) {
  const r = await fetch(`${API_BASE}/api/${resource}`);
  if (r.status === 403) throw Object.assign(new Error("Forbidden"), { status: 403 });
  if (!r.ok) throw new Error(`API GET ${resource} failed: ${r.status}`);
  return r.json();
}
async function apiSet(resource, data) {
  if (!API_BASE) return; // ローカル開発時はスキップ
  const r = await fetch(`${API_BASE}/api/${resource}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  if (r.status === 403) throw Object.assign(new Error("Forbidden"), { status: 403 });
  if (!r.ok) throw new Error(`API SET ${resource} failed: ${r.status}`);
}

// ── IP アクセス制限 ────────────────────────────────────────────────────────────
const ALLOWED_IPS = new Set([
  "210.172.143.39",   // 主回線
  "210.172.143.37",   // Mac用VPN回線
  "210.172.130.69",   // 追加回線
]);

async function fetchMyIP() {
  const res = await fetch("https://api.ipify.org?format=json");
  const { ip } = await res.json();
  return ip;
}

function AccessDenied({ ip }) {
  return (
    <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl p-10 w-full max-w-sm text-center">
        <div className="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
          <svg className="w-8 h-8 text-red-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
              d="M12 9v2m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" />
          </svg>
        </div>
        <h1 className="text-xl font-bold text-slate-800 mb-2">アクセス拒否</h1>
        <p className="text-slate-500 text-sm mb-4">このIPアドレスからのアクセスは許可されていません。</p>
        <div className="bg-slate-100 rounded-lg px-4 py-2 text-xs text-slate-600 font-mono">{ip}</div>
        <p className="text-xs text-slate-400 mt-4">許可されたネットワークから接続してください。</p>
      </div>
    </div>
  );
}

function IPCheckScreen() {
  return (
    <div className="min-h-screen bg-slate-900 flex items-center justify-center">
      <div className="text-center">
        <div className="w-10 h-10 border-4 border-blue-500 border-t-transparent rounded-full animate-spin mx-auto mb-4" />
        <p className="text-slate-300 text-sm">接続確認中...</p>
      </div>
    </div>
  );
}
const IDB_NAME      = "teppou_idb";
const IDB_VER       = 3;                    // v3: records / past_mgmt / kv ストア
const IDB_STORE     = "records";

// ── IndexedDB helpers ──────────────────────────────────────────────────────────
function idbOpen() {
  return new Promise((res, rej) => {
    const req = indexedDB.open(IDB_NAME, IDB_VER);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(IDB_STORE))     db.createObjectStore(IDB_STORE,   { keyPath: "id" });
      if (!db.objectStoreNames.contains("past_mgmt"))   db.createObjectStore("past_mgmt", { keyPath: "id" });
      if (!db.objectStoreNames.contains("kv"))          db.createObjectStore("kv",        { keyPath: "k" });
    };
    req.onsuccess = e => res(e.target.result);
    req.onerror   = e => rej(e.target.error);
  });
}
async function idbGetAll(storeName = IDB_STORE) {
  const db = await idbOpen();
  return new Promise((res, rej) => {
    try {
      const req = db.transaction(storeName, "readonly").objectStore(storeName).getAll();
      req.onsuccess = () => res(req.result || []);
      req.onerror   = e => rej(e.target.error);
    } catch(e) { res([]); }
  });
}
async function idbPutAll(records, storeName = IDB_STORE) {
  const db = await idbOpen();
  return new Promise((res, rej) => {
    try {
      const tx    = db.transaction(storeName, "readwrite");
      const store = tx.objectStore(storeName);
      store.clear();
      records.forEach(r => store.put(r));
      tx.oncomplete = res;
      tx.onerror    = e => rej(e.target.error);
    } catch(e) { rej(e); }
  });
}
// 過去商談管理（past_mgmt ストア）
async function idbPastGetAll() { return idbGetAll("past_mgmt"); }
async function idbPastPutAll(records) { return idbPutAll(records, "past_mgmt"); }
// 旧・分離DB（teppou_past_idb）からの復旧用
async function idbLegacyPastGetAll() {
  return new Promise((res) => {
    try {
      const req = indexedDB.open("teppou_past_idb", 1);
      req.onupgradeneeded = e => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains("past_mgmt")) db.createObjectStore("past_mgmt", { keyPath: "id" });
      };
      req.onsuccess = e => {
        const db = e.target.result;
        try {
          const r = db.transaction("past_mgmt", "readonly").objectStore("past_mgmt").getAll();
          r.onsuccess = () => res(r.result || []);
          r.onerror   = () => res([]);
        } catch { res([]); }
      };
      req.onerror = () => res([]);
    } catch { res([]); }
  });
}
// 汎用 KV（pastDeals などの配列ブロブ用）
async function idbKvGet(key) {
  const db = await idbOpen();
  return new Promise((res) => {
    try {
      const req = db.transaction("kv", "readonly").objectStore("kv").get(key);
      req.onsuccess = () => res(req.result ? req.result.v : null);
      req.onerror   = () => res(null);
    } catch(e) { res(null); }
  });
}
async function idbKvSet(key, value) {
  const db = await idbOpen();
  return new Promise((res, rej) => {
    try {
      const tx = db.transaction("kv", "readwrite");
      tx.objectStore("kv").put({ k: key, v: value });
      tx.oncomplete = res;
      tx.onerror    = e => rej(e.target.error);
    } catch(e) { rej(e); }
  });
}

const IS_MEMBERS = [
  "櫻井　肇","上浦　諒大","井上　妃音","太田　小百合","十文字　菜月",
  "中　翔吾","早坂　直樹","小田切　翼","横井　優一","青木　大輔",
  "山田",
];

const STATUS_CFG = {
  "未架電":            { row:"",              bg:"bg-slate-100",  text:"text-slate-500",  border:"border-slate-300",  dot:"bg-slate-400"  },
  "0.日程調整":        { row:"bg-rose-50",    bg:"bg-rose-200",   text:"text-rose-800",   border:"border-rose-400",   dot:"bg-rose-600"   },
  "1.高確度":          { row:"bg-pink-50",    bg:"bg-pink-100",   text:"text-pink-700",   border:"border-pink-300",   dot:"bg-pink-500"   },
  "2.優先":            { row:"bg-orange-50",  bg:"bg-orange-100", text:"text-orange-700", border:"border-orange-300", dot:"bg-orange-400" },
  "3.並":              { row:"bg-slate-50",   bg:"bg-slate-200",  text:"text-slate-500",  border:"border-slate-400",  dot:"bg-slate-400"  },
  "4.受付カット":      { row:"bg-amber-50",   bg:"bg-amber-100",  text:"text-amber-800",  border:"border-amber-400",  dot:"bg-amber-500"  },
  "4.別担当架電":      { row:"bg-sky-50",     bg:"bg-sky-100",    text:"text-sky-700",    border:"border-sky-300",    dot:"bg-sky-500"    },
  "4.商談中":          { row:"bg-violet-50",  bg:"bg-violet-100", text:"text-violet-700", border:"border-violet-300", dot:"bg-violet-500" },
  "5.メール送付":      { row:"bg-yellow-50",  bg:"bg-yellow-100", text:"text-yellow-700", border:"border-yellow-300", dot:"bg-amber-400"  },
  "6.コネクト（改）":  { row:"bg-indigo-50",  bg:"bg-indigo-100", text:"text-indigo-600", border:"border-indigo-300", dot:"bg-indigo-400" },
  "7.コネクト（無）":  { row:"bg-blue-50",    bg:"bg-blue-100",   text:"text-blue-900",   border:"border-blue-600",   dot:"bg-blue-800"   },
  "8.不要":            { row:"bg-purple-50",  bg:"bg-purple-100", text:"text-purple-900", border:"border-purple-600", dot:"bg-purple-800" },
  "8.当社契約":        { row:"bg-gray-100",   bg:"bg-gray-200",   text:"text-gray-800",   border:"border-gray-500",   dot:"bg-gray-700"   },
  "9.アポ獲得":        { row:"bg-red-50",     bg:"bg-red-100",    text:"text-red-900",    border:"border-red-600",    dot:"bg-red-700"    },
  "リードなし":        { row:"bg-stone-100",  bg:"bg-stone-200",  text:"text-stone-600",  border:"border-stone-400",  dot:"bg-stone-500"  },
};

// ステータス別集計の表示順（冒頭の数字順 / 1.高確度を先頭・リードなしを最後尾）
const STATUS_ORDER = [
  "1.高確度","2.優先","3.並","4.受付カット","4.別担当架電","4.商談中",
  "5.メール送付","6.コネクト（改）","7.コネクト（無）","8.不要","8.当社契約",
  "9.アポ獲得","0.日程調整","未架電","リードなし",
];
const statusOrderIdx = s => { const i = STATUS_ORDER.indexOf(s); return i === -1 ? 998 : i; };

// 「架電」プリセット（ステータス絞り込み）: コネクト（無）・並・優先・高確度
const CALL_PRESET_STATUSES = ["7.コネクト（無）","3.並","2.優先","1.高確度"];

// 当日帰社の可能性がある不在理由 / 完了扱いステータス（アラート判定用）
const RECALL_REASONS = ["席外","午前","お昼","午後","夕方"];
const DONE_STATUSES  = ["8.不要","8.当社契約"];

// ── 再アプローチステータス ─────────────────────────────────────────────────────
const REAPPROACH_STATUS = {
  "未アプローチ": { bg:"bg-slate-100",   text:"text-slate-500",  dot:"bg-slate-400"  },
  "アプローチ中": { bg:"bg-blue-100",    text:"text-blue-700",   dot:"bg-blue-500"   },
  "連絡済み":     { bg:"bg-yellow-100",  text:"text-yellow-700", dot:"bg-amber-500"  },
  "商談中":       { bg:"bg-purple-100",  text:"text-purple-700", dot:"bg-purple-500" },
  "完了":         { bg:"bg-green-100",   text:"text-green-700",  dot:"bg-green-500"  },
  "見送り":       { bg:"bg-slate-200",   text:"text-slate-500",  dot:"bg-slate-400"  },
};

// 過去商談管理CSVヘッダーマッピング
function mapPastMgmtHeaders(headers) {
  const m = {};
  headers.forEach((h, i) => {
    const n = String(h).replace(/[\s　]/g,"").toLowerCase();
    if (/取引先名|会社名|企業名|ブランド名/.test(n))     m.companyName  = i;
    else if (/完了予定日|予定日/.test(n))      m.targetDate   = i;
    else if (/確度/.test(n))                   m.probability  = i;
    else if (/作成者|担当者/.test(n))          m.creator      = i;
    else if (/状況|フェーズ|進捗|ステータス|状態/.test(n)) m.progress = i; // 進捗（プルダウン）
  });
  return m;
}

// ── リード先（ソース）設定 ─────────────────────────────────────────────────────
const LEAD_SOURCE_CFG = {
  "エキテン":       { bg:"bg-rose-200",    text:"text-rose-900",    dot:"bg-rose-500"    },
  "更地リード":     { bg:"bg-indigo-800",  text:"text-white",       dot:"bg-indigo-900"  },
  "新規リード":     { bg:"bg-blue-200",    text:"text-blue-900",    dot:"bg-blue-500"    },
  "過去商談(他)":   { bg:"bg-orange-100",  text:"text-orange-800",  dot:"bg-orange-400"  },
  "過去商談(自)":   { bg:"bg-pink-100",    text:"text-pink-800",    dot:"bg-pink-400"    },
  "過去商談(DM)":   { bg:"bg-purple-100",  text:"text-purple-700",  dot:"bg-purple-400"  },
  "Google":         { bg:"bg-green-100",   text:"text-green-800",   dot:"bg-green-500"   },
  "過去プル":       { bg:"bg-teal-100",    text:"text-teal-800",    dot:"bg-teal-500"    },
  "トライハッチ":   { bg:"bg-purple-700",  text:"text-white",       dot:"bg-purple-900"  },
  "リスト":         { bg:"bg-white",       text:"text-slate-600",   dot:"bg-slate-300"   },
  "展示会":         { bg:"bg-slate-100",   text:"text-slate-600",   dot:"bg-slate-400"   },
  "セミナー":       { bg:"bg-slate-100",   text:"text-slate-600",   dot:"bg-slate-400"   },
  "名刺交換":       { bg:"bg-slate-50",    text:"text-slate-500",   dot:"bg-slate-300"   },
  "プル":           { bg:"bg-rose-100",    text:"text-rose-700",    dot:"bg-rose-400"    },
};

// ── 不在理由設定 ───────────────────────────────────────────────────────────────
const ABSENCE_REASON_CFG = {
  "席外":    { bg:"bg-slate-100",  text:"text-slate-600",  dot:"bg-slate-400"  },
  "午前":    { bg:"bg-sky-100",    text:"text-sky-700",    dot:"bg-sky-400"    },
  "お昼":    { bg:"bg-amber-100",  text:"text-amber-700",  dot:"bg-amber-400"  },
  "午後":    { bg:"bg-orange-100", text:"text-orange-700", dot:"bg-orange-400" },
  "夕方":    { bg:"bg-orange-200", text:"text-orange-800", dot:"bg-orange-500" },
  "出張":    { bg:"bg-red-700",    text:"text-white",      dot:"bg-red-900"    },
  "折返":    { bg:"bg-indigo-100", text:"text-indigo-700", dot:"bg-indigo-400" },
  "留守":    { bg:"bg-purple-100", text:"text-purple-700", dot:"bg-purple-400" },
  "休み":    { bg:"bg-teal-100",   text:"text-teal-700",   dot:"bg-teal-400"   },
  "不通":    { bg:"bg-gray-100",   text:"text-gray-600",   dot:"bg-gray-400"   },
  "着拒？":  { bg:"bg-rose-200",   text:"text-rose-800",   dot:"bg-rose-500"   },
};

// 列幅は1画面に収まるようタイトに最適化（table-fixed + truncate でガタつきゼロ）
const ALL_COLUMNS = [
  { key:"companyName",   label:"企業名",                              required:true,  w:"w-[184px]" },
  { key:"lastCallDate",  label:"架電日",                              required:false, w:"w-[96px]"  },
  { key:"nextCallDate",  label:"次回架電日",                          required:false, w:"w-[104px]" },
  { key:"status",        label:"状況",                                required:false, w:"w-[140px]" },
  { key:"industry",      label:"業種",                                required:false, w:"w-[88px]"  },
  { key:"leadSource",    label:"ソース",                              required:false, w:"w-[120px]" },
  { key:"leadAddedDate", label:"リード追加日",                        required:false, w:"w-[100px]" },
  { key:"hpSite",        label:"HPサイト",                            required:false, w:"w-[120px]" },
  { key:"gbp",           label:"GBP",                                 required:false, w:"w-[68px]"  },
  { key:"phone",         label:"電話番号",                            required:false, w:"w-[110px]" },
  { key:"assignee",      label:"担当者",                              required:false, w:"w-[88px]"  },
  { key:"createdBy",     label:"追加者",                              required:false, w:"w-[88px]"  },
  { key:"importMonth",   label:"取込月",                              required:false, w:"w-[84px]"  },
  { key:"department",    label:"部署",                                required:false, w:"w-[84px]"  },
  { key:"absenceReason", label:"不在理由",                            required:false, w:"w-[84px]"  },
  { key:"gbpManagement", label:"GBPの管理",                           required:false, w:"w-[92px]"  },
  { key:"memo",          label:"メモ",                                required:false, w:"w-[188px]" },
  { key:"storeCount",    label:"店舗数",                              required:false, w:"w-[76px]"  },
  { key:"refusalReason", label:"断り理由",                            required:false, w:"w-[96px]"  },
  { key:"posting",       label:"投稿",                                required:false, w:"w-[64px]"  },
  { key:"review",        label:"口コミ",                              required:false, w:"w-[68px]"  },
  { key:"sns",           label:"SNS",                                 required:false, w:"w-[60px]"  },
  { key:"instagram",     label:"Insta",                               required:false, w:"w-[64px]"  },
  { key:"line",          label:"Line",                                required:false, w:"w-[60px]"  },
  { key:"facebook",      label:"FB",                                  required:false, w:"w-[56px]"  },
  { key:"twitter",       label:"Twitter",                             required:false, w:"w-[72px]"  },
  { key:"os",            label:"OS",                                  required:false, w:"w-[56px]"  },
  { key:"mailFlag",      label:"メール",                              required:false, w:"w-[72px]"  },
  { key:"email",         label:"メアド",                              required:false, w:"w-[150px]" },
  { key:"gbpSiteUrl",    label:"GBPサイトURL",                        required:false, w:"w-[150px]" },
];

const DEFAULT_VISIBLE_COLS = [
  "companyName","lastCallDate","nextCallDate","status","absenceReason","storeCount","phone","assignee","memo",
];

// ── Helpers ────────────────────────────────────────────────────────────────────
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  // タブが含まれていればTSV（スプレッドシートのコピペ）として処理
  const isTab = lines[0].includes("\t");
  return lines.map(line => {
    if (isTab) return line.split("\t").map(c => c.trim());
    const cells = []; let inQ = false, cell = "";
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { if (inQ && line[i+1]==='"') { cell+='"'; i++; } else inQ=!inQ; }
      else if (ch===',' && !inQ) { cells.push(cell.trim()); cell=""; }
      else cell += ch;
    }
    cells.push(cell.trim()); return cells;
  });
}

function isAggregateRow(row) {
  const ne = row.filter(c => c.trim() !== "");
  if (!ne.length) return false;
  return ne.filter(c => /^\d[\d,.]*$/.test(c.trim())).length / ne.length >= 0.4;
}

function mapSalesHeaders(headers) {
  const m = {};
  headers.forEach((h, i) => {
    const n = h.replace(/[\s　]/g,"").toLowerCase();
    if      (/企業名|会社名|法人名|取引先名?/.test(n))                      m.companyName   = i;
    else if (/電話|tel|phone/.test(n))                            m.phone         = i;
    else if (/メアド/.test(n))                                    m.email         = i;
    else if (/^メール$/.test(n))                                  m.mailFlag      = i;
    else if (/gbpサイト/.test(n))                                 m.gbpSiteUrl    = i;
    else if (/gbpの管理|gbp管理/.test(n))                         m.gbpManagement = i;
    else if (/^gbp$/.test(n))                                     m.gbp           = i;
    else if (/^hp$|hpサイト|^hp[^a-z]/.test(n))                  m.hpSite        = i;
    else if (/簡易|次回架電時の情報|メモ|備考|note|コメント/.test(n)) m.memo       = i;
    else if (/担当者?名?|営業担当|作成者/.test(n))                 m.assignee      = i;
    else if (/ステータス|状態|状況/.test(n))                      m.status        = i;
    else if (/次回架電|次架電|コールバック/.test(n))               m.nextCallDate  = i;
    else if (/最終架電|^架電日/.test(n))                          m.lastCallDate  = i;
    else if (/業種/.test(n))                                      m.industry      = i;
    else if (/ソース|リスト出所/.test(n))                         m.leadSource    = i;
    else if (/部署|部門/.test(n))                                 m.department    = i;
    else if (/不在理由/.test(n))                                  m.absenceReason = i;
    else if (/店舗数|店舗/.test(n))                               m.storeCount    = i;
    else if (/断り理由|断り/.test(n))                             m.refusalReason = i;
    else if (/投稿/.test(n))                                      m.posting       = i;
    else if (/口コミ|レビュー/.test(n))                           m.review        = i;
    else if (/^sns$/.test(n))                                     m.sns           = i;
    else if (/insta|インスタ/.test(n))                            m.instagram     = i;
    else if (/^line$|ライン/.test(n))                             m.line          = i;
    else if (/^fb$|facebook/.test(n))                             m.facebook      = i;
    else if (/^twitter$|^x$|ツイッター/.test(n))                  m.twitter       = i;
    else if (/^os$/.test(n))                                      m.os            = i;
  });
  return m;
}

// 過去商談CSVヘッダーマッピング
function mapPastDealsHeaders(headers) {
  const m = {};
  headers.forEach((h, i) => {
    const n = h.replace(/[\s　]/g, "").toLowerCase();
    if (/企業名|会社名|法人名|取引先名?/.test(n))                   m.companyName  = i;
    else if (/過去.*状況|過去.*ステータス|状況|ステータス|状態|フェーズ/.test(n)) m.pastStatus = i;
    else if (/最終架電|架電日|過去.*架電|完了予定日/.test(n))         m.lastCallDate = i;
    // 自社側の架電スタッフ/担当者/作成者 → 商談所有者へ
    else if (/担当者?|営業担当|作成者|オペレーター|架電者|ユーザー名/.test(n)) m.dealOwner = i;
    else if (/メモ|備考|note|コメント|商談メモ|経緯/.test(n))         m.memo         = i;
  });
  return m;
}

function convertMitelStatus(callType, tags) {
  const t = String(tags || "");
  if (/アポ獲得[（(]新規[）)]/.test(t))           return "アポイント獲得商談";
  if (t.includes("コネクト"))                     return "担当コネクト";
  if (t.includes("受付カット"))                   return "受付断り";
  if (t.includes("担当者不在"))                   return "不在";
  const ct = String(callType || "");
  if (ct.includes("発信") && ct.includes("不在")) return "不通";
  return null;
}

function normName(s)    { return String(s||"").replace(/[\s　]/g,"").toLowerCase(); }
// 店舗数を数値化（カンマ除去）。未記入や0は null
function parseStoreCount(v) {
  const n = parseInt(String(v ?? "").replace(/,/g, ""), 10);
  return Number.isFinite(n) && n > 0 ? n : null;
}

// ── 店舗数の自動分析（未記入対策・高速インデックス方式） ─────────────────────────
// 全レコードから「企業名→確定店舗数」「企業名→重複件数」のMapを一度だけ構築する
function buildStoreIndex(allRecords, pastDeals = []) {
  const storeByName = new Map(); // normName → 最大確定店舗数
  const dupByName   = new Map(); // normName → 件数（営業リストのみ）
  for (const r of allRecords) {
    const n = normName(r.companyName);
    if (!n) continue;
    dupByName.set(n, (dupByName.get(n) || 0) + 1);
    const sc = parseStoreCount(r.storeCount);
    if (sc !== null) { const cur = storeByName.get(n); if (cur === undefined || sc > cur) storeByName.set(n, sc); }
  }
  for (const d of pastDeals) {
    const n = normName(d.companyName);
    if (!n) continue;
    const sc = parseStoreCount(d.storeCount);
    if (sc !== null) { const cur = storeByName.get(n); if (cur === undefined || sc > cur) storeByName.set(n, sc); }
  }
  return { storeByName, dupByName };
}

// O(1)：構築済みインデックスを使い、確定値 or 推測値（仮）を返す
function analyzeStoreCount(target, index) {
  const own = parseStoreCount(target.storeCount);
  if (own !== null) return { value: own, estimated: false };
  const tn = normName(target.companyName);
  if (!tn || !index) return { value: null, estimated: false };
  // ① 同名の確定店舗数
  const byName = index.storeByName.get(tn);
  if (byName !== undefined) return { value: byName, estimated: true };
  // ② 同名重複件数（多店舗チェーン）
  const dup = index.dupByName.get(tn) || 0;
  if (dup > 1) return { value: dup, estimated: true };
  return { value: null, estimated: false };
}

// ── CSV バックアップ生成 ────────────────────────────────────────────────────────
function generateBackupCSV(records) {
  const headers = ALL_COLUMNS.map(c => c.label);
  const rows = records.map(r => ALL_COLUMNS.map(c => {
    const v = r[c.key] ?? "";
    if (c.key === "lastCallDate" || c.key === "nextCallDate") return fmtDate(normDate(String(v)));
    return String(v);
  }));
  const bom = "﻿";
  return bom + [headers, ...rows].map(row =>
    row.map(c => (c.includes(",") || c.includes('"') || c.includes("\n"))
      ? `"${c.replace(/"/g,'""')}"` : c).join(",")
  ).join("\r\n");
}
function triggerCSVDownload(csv, filename) {
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}
function genId()        { return Date.now().toString(36)+Math.random().toString(36).slice(2); }
function getToday()     { return new Date().toISOString().slice(0,10); }
function nowIso()       { return new Date().toISOString(); }
// 表示用: YYYY-MM-DD → YYYY/MM/DD
function fmtDate(d) { return d ? d.replace(/-/g, "/") : "—"; }
// 正規化: 様々な日付文字列 → YYYY-MM-DD
function normDate(d) {
  if (!d) return "";
  const s = String(d).trim();
  if (!s) return "";
  // YYYY-MM-DD / YYYY/MM/DD（正常系）
  if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(s)) return s.replace(/\//g, "-");
  // YYYY/M/D（単桁あり）
  const ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (ymd) return `${ymd[1]}-${ymd[2].padStart(2,"0")}-${ymd[3].padStart(2,"0")}`;
  // M/D/YY または M/D/YYYY（US形式）
  const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (mdy) {
    const yr = mdy[3].length === 2 ? (parseInt(mdy[3]) <= 50 ? "20" : "19") + mdy[3] : mdy[3];
    return `${yr}-${mdy[1].padStart(2,"0")}-${mdy[2].padStart(2,"0")}`;
  }
  return s.slice(0, 10).replace(/\//g, "-");
}

// 旧ステータス名（10.リードなし）→ 新名（リードなし）へ移行
function migrateRecords(recs) {
  if (!Array.isArray(recs)) return recs;
  return recs.map(r => (r && r.status === "10.リードなし") ? { ...r, status: "リードなし" } : r);
}

// 次回架電日が架電日より前なら架電日に合わせる（架電済みなので次回は当日以降）
function clampNextCall(r) {
  const lc = normDate(r.lastCallDate), nc = normDate(r.nextCallDate);
  if (lc && nc && nc < lc) return { ...r, nextCallDate: lc };
  return r;
}

// URL → ドメイン（www除去）。ファビコン取得用
function faviconDomain(url) {
  // 文字列以外（数値・null・undefined等）は解析せず即フォールバック（バグ落ち防止）
  if (typeof url !== "string") return "";
  const s0 = url.trim();
  if (!s0) return "";
  try {
    const s = /^https?:\/\//i.test(s0) ? s0 : "https://" + s0;
    const host = new URL(s).hostname.replace(/^www\./, "");
    // 末尾がアルファベットTLD（.com/.jp等）の正当なドメインのみ採用。
    // スペース・不正文字・数値のみ・IPアドレス（例:0.0.48.57）は除外 → アバター表示。
    return /^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9-]+)*\.[a-z]{2,}$/i.test(host) ? host : "";
  } catch { return ""; }
}

// Google高画質ファビコンAPI（sz=64）。国内企業（日本郵便の〒等）も確実に取得できる。
function googleFavicon(domain) {
  return `https://www.google.com/s2/favicons?domain=${encodeURIComponent(domain)}&sz=64`;
}
// 旧Clearbit自動ロゴURLからドメインを取り出す（移行用）
function clearbitDomain(logoUrl) {
  const m = typeof logoUrl === "string" && logoUrl.match(/^https:\/\/logo\.clearbit\.com\/(.+)$/i);
  return m ? m[1] : "";
}
// 自動生成（Google/Clearbitファビコン）URLか判定
function isAutoFaviconUrl(u) {
  return typeof u === "string" && /^https:\/\/(www\.google\.com\/s2\/favicons|logo\.clearbit\.com)/i.test(u);
}

// 大手企業のロゴ・キュレーション（店舗数の多い主要企業に公式ファビコンのリンクを付与）。
// ハイフン・空白の表記揺れを吸収して企業名で照合。手動設定ロゴは尊重して上書きしない。
const KNOWN_LOGOS = [
  { kw: ["東京海上"],                         domain: "tokiomarine-nichido.co.jp" },
  { kw: ["日本郵便"],                         domain: "post.japanpost.jp" },
  { kw: ["セブンイレブン"],                   domain: "sej.co.jp" },
  { kw: ["タイムズ"],                         domain: "times24.co.jp" },
  { kw: ["ファミリーマート"],                 domain: "family.co.jp" },
  { kw: ["公文"],                             domain: "kumon.ne.jp" },
  { kw: ["ローソン"],                         domain: "lawson.co.jp" },
  { kw: ["eMobilityPower", "モビリティパワー"], domain: "e-mobipower.co.jp" },
];
const stripLogoKey = s => String(s || "").replace(/[\s　\-‐‑‒–—―－]/g, "");
function matchKnownLogo(name) {
  const n = stripLogoKey(name);
  for (const e of KNOWN_LOGOS) {
    if (e.kw.some(k => n.includes(stripLogoKey(k)))) return googleFavicon(e.domain);
  }
  return null;
}

// 頭文字アバターの配色パレット（フルクラス文字列＝TailwindのJITに確実に含める）
const AVATAR_COLORS = [
  "bg-slate-100 text-slate-600",
  "bg-blue-100 text-blue-700",
  "bg-emerald-100 text-emerald-700",
  "bg-amber-100 text-amber-700",
  "bg-rose-100 text-rose-700",
  "bg-violet-100 text-violet-700",
  "bg-cyan-100 text-cyan-700",
  "bg-pink-100 text-pink-700",
];
function avatarColor(name) {
  const s = String(name || "");
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
  return AVATAR_COLORS[h % AVATAR_COLORS.length];
}

// ── CompanyLogo ────────────────────────────────────────────────────────────────
// 企業ロゴは手動登録のみ（URL／画像アップロード／Ctrl+V貼り付けのBase64）。自動取得なし。
// 未登録・読み込み失敗時は中身なしの「白紙（無地枠）」を表示（頭文字アバターは表示しない）。
// 枠サイズ完全固定（w-6 h-6）でガタつきゼロ・遅延読み込み。
const CompanyLogo = memo(function CompanyLogo({ logoUrl }) {
  // 失敗したsrcを記録（行の使い回しで別レコードに変わっても誤表示しない）
  const [errSrc, setErrSrc] = useState(null);
  const showImg = logoUrl && errSrc !== logoUrl;
  return (
    <span className="w-6 h-6 rounded bg-white border border-slate-200 flex items-center justify-center flex-shrink-0 overflow-hidden shadow-sm p-0.5">
      {showImg && (
        <img
          src={logoUrl}
          alt=""
          loading="lazy"
          onError={() => setErrSrc(logoUrl)}
          className="w-full h-full object-contain"
        />
      )}
    </span>
  );
});

// ── StatusBadge ────────────────────────────────────────────────────────────────
function StatusBadge({ status }) {
  const c = STATUS_CFG[status] ?? { bg:"bg-gray-100", text:"text-gray-600", border:"border-gray-300" };
  return (
    <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium border whitespace-nowrap ${c.bg} ${c.text} ${c.border}`}>
      {status||"—"}
    </span>
  );
}

// ── LeadSourceBadge ────────────────────────────────────────────────────────────
function LeadSourceBadge({ source }) {
  if (!source) return null;
  const c = LEAD_SOURCE_CFG[source] ?? { bg:"bg-slate-100", text:"text-slate-600" };
  return (
    <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-semibold whitespace-nowrap border border-black/10 ${c.bg} ${c.text}`}>
      {source}
    </span>
  );
}

// ── AbsenceReasonBadge ─────────────────────────────────────────────────────────
function AbsenceReasonBadge({ reason }) {
  if (!reason) return null;
  const c = ABSENCE_REASON_CFG[reason] ?? { bg:"bg-slate-100", text:"text-slate-600" };
  return (
    <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-semibold whitespace-nowrap border border-black/10 ${c.bg} ${c.text}`}>
      {reason}
    </span>
  );
}

// ── StoreCountCell（確定値 / 推測値「（仮）」表示） ─────────────────────────────
function StoreCountCell({ analysis }) {
  if (!analysis || analysis.value === null)
    return <span className="text-slate-300 text-xs">—</span>;
  if (analysis.estimated) {
    return (
      <span className="inline-flex items-baseline gap-0.5">
        <span className="text-xs text-slate-500">{analysis.value.toLocaleString()}</span>
        <span className="text-[10px] text-slate-400 font-medium">（仮）</span>
      </span>
    );
  }
  return <span className="text-slate-700 text-xs font-medium">{analysis.value.toLocaleString()}</span>;
}

// ── White background removal (BFS flood fill from edges) ──────────────────────
function removeWhiteBg(sourceImg, threshold = 238) {
  const canvas = document.createElement("canvas");
  canvas.width  = sourceImg.naturalWidth;
  canvas.height = sourceImg.naturalHeight;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(sourceImg, 0, 0);
  const { width, height } = canvas;
  const imgData = ctx.getImageData(0, 0, width, height);
  const d = imgData.data;
  const visited = new Uint8Array(width * height);
  const queue = [];

  // Seed from all 4 edges
  for (let x = 0; x < width;  x++) { queue.push(x); queue.push((height - 1) * width + x); }
  for (let y = 1; y < height - 1; y++) { queue.push(y * width); queue.push(y * width + width - 1); }

  let qi = 0;
  while (qi < queue.length) {
    const p = queue[qi++];
    if (visited[p]) continue;
    visited[p] = 1;
    const i = p * 4;
    const x = p % width, y = (p / width) | 0;
    const flood =
      d[i + 3] === 0 ||                                          // already transparent
      (d[i] >= threshold && d[i + 1] >= threshold && d[i + 2] >= threshold); // white/near-white
    if (!flood) continue;
    d[i] = d[i + 1] = d[i + 2] = d[i + 3] = 0; // make transparent
    if (x > 0)          queue.push(p - 1);
    if (x < width - 1)  queue.push(p + 1);
    if (y > 0)          queue.push(p - width);
    if (y < height - 1) queue.push(p + width);
  }
  ctx.putImageData(imgData, 0, 0);
  return canvas;
}

// ── Trim transparent edges and scale content to fill canvas ──────────────────
function trimAndFill(canvas) {
  const ctx = canvas.getContext("2d");
  const { width, height } = canvas;
  const d = ctx.getImageData(0, 0, width, height).data;
  let x0 = width, y0 = height, x1 = 0, y1 = 0;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (d[(y * width + x) * 4 + 3] > 8) {
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        if (y > y1) y1 = y;
      }
    }
  }
  if (x0 >= x1 || y0 >= y1) return; // nothing visible
  // 5% padding around content
  const px = Math.max(4, Math.round((x1 - x0) * 0.05));
  const py = Math.max(4, Math.round((y1 - y0) * 0.05));
  x0 = Math.max(0, x0 - px); y0 = Math.max(0, y0 - py);
  x1 = Math.min(width,  x1 + px); y1 = Math.min(height, y1 + py);
  // Copy cropped region to a temp canvas
  const tmp = document.createElement("canvas");
  tmp.width = x1 - x0; tmp.height = y1 - y0;
  tmp.getContext("2d").drawImage(canvas, x0, y0, x1 - x0, y1 - y0, 0, 0, tmp.width, tmp.height);
  // Scale back to fill original canvas
  ctx.clearRect(0, 0, width, height);
  ctx.drawImage(tmp, 0, 0, width, height);
}

// ── CropModal ──────────────────────────────────────────────────────────────────
function CropModal({ src, onCrop, onClose }) {
  const [loaded,   setLoaded]   = useState(false);
  const [natSize,  setNatSize]  = useState({ w: 1, h: 1 });
  const [dispSize, setDispSize] = useState({ w: 360, h: 240 });
  const [box,      setBox]      = useState({ x: 0, y: 0, w: 150, h: 150 });
  const [drag,     setDrag]     = useState(null);
  const canvasRef = useRef();
  const imgRef    = useRef(null);
  const DMAX = 360;

  // Load image → remove white BG → init crop box
  useEffect(() => {
    const img = new Image();
    img.onload = () => {
      // Apply white background removal at full resolution
      const cleanCanvas = removeWhiteBg(img);
      const cleanImg = new Image();
      cleanImg.onload = () => {
        const scale = Math.min(DMAX / img.naturalWidth, DMAX / img.naturalHeight, 1);
        const w = Math.round(img.naturalWidth  * scale);
        const h = Math.round(img.naturalHeight * scale);
        imgRef.current = cleanImg;
        setNatSize({ w: img.naturalWidth, h: img.naturalHeight });
        setDispSize({ w, h });
        const sz = Math.round(Math.min(w, h) * 0.78);
        setBox({ x: Math.round((w - sz) / 2), y: Math.round((h - sz) / 2), w: sz, h: sz });
        setLoaded(true);
      };
      cleanImg.src = cleanCanvas.toDataURL("image/png");
    };
    img.src = src;
  }, [src]);

  // Draw canvas overlay
  useEffect(() => {
    if (!loaded || !canvasRef.current || !imgRef.current) return;
    const canvas = canvasRef.current;
    canvas.width  = dispSize.w;
    canvas.height = dispSize.h;
    const ctx = canvas.getContext("2d");
    const img = imgRef.current;
    const sx  = natSize.w / dispSize.w;
    const sy  = natSize.h / dispSize.h;

    // Dimmed full image
    ctx.globalAlpha = 0.32;
    ctx.drawImage(img, 0, 0, dispSize.w, dispSize.h);
    ctx.globalAlpha = 1;

    // Bright crop region
    ctx.drawImage(img,
      box.x * sx, box.y * sy, box.w * sx, box.h * sy,
      box.x, box.y, box.w, box.h
    );

    // Border
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 2;
    ctx.strokeRect(box.x + 1, box.y + 1, box.w - 2, box.h - 2);

    // Rule-of-thirds grid
    ctx.strokeStyle = "rgba(255,255,255,0.3)";
    ctx.lineWidth = 0.5;
    for (let i = 1; i < 3; i++) {
      const gx = box.x + (box.w * i) / 3;
      const gy = box.y + (box.h * i) / 3;
      ctx.beginPath(); ctx.moveTo(gx, box.y);    ctx.lineTo(gx, box.y + box.h); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(box.x, gy);    ctx.lineTo(box.x + box.w, gy); ctx.stroke();
    }

    // Corner handles
    ctx.shadowColor = "rgba(0,0,0,0.6)";
    ctx.shadowBlur  = 4;
    ctx.fillStyle   = "#ffffff";
    const hs = 8;
    [
      [box.x,             box.y            ],
      [box.x + box.w - hs, box.y           ],
      [box.x,              box.y + box.h - hs],
      [box.x + box.w - hs, box.y + box.h - hs],
    ].forEach(([cx, cy]) => ctx.fillRect(cx, cy, hs, hs));
    ctx.shadowBlur = 0;

    // Resize indicator (SE corner)
    ctx.strokeStyle = "#60a5fa";
    ctx.lineWidth = 2;
    const r = 12;
    ctx.beginPath();
    ctx.moveTo(box.x + box.w - r,     box.y + box.h - 2);
    ctx.lineTo(box.x + box.w - 2,     box.y + box.h - 2);
    ctx.lineTo(box.x + box.w - 2,     box.y + box.h - r);
    ctx.stroke();
  }, [loaded, box, dispSize, natSize]);

  // Mouse down: determine action
  const onMouseDown = (e) => {
    e.preventDefault();
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const px = (e.clientX - rect.left) * (canvas.width  / rect.width);
    const py = (e.clientY - rect.top)  * (canvas.height / rect.height);
    const hit = 16; // resize hit area
    if (px >= box.x + box.w - hit && py >= box.y + box.h - hit) {
      setDrag({ type:"resize", sx:px, sy:py, sb:{ ...box } });
    } else if (px >= box.x && px <= box.x + box.w && py >= box.y && py <= box.y + box.h) {
      setDrag({ type:"move",   sx:px, sy:py, sb:{ ...box } });
    }
  };

  // Global mouse move / up
  useEffect(() => {
    if (!drag) return;
    const getPos = (e) => {
      const canvas = canvasRef.current;
      if (!canvas) return { x: 0, y: 0 };
      const rect = canvas.getBoundingClientRect();
      return {
        x: (e.clientX - rect.left) * (canvas.width  / rect.width),
        y: (e.clientY - rect.top)  * (canvas.height / rect.height),
      };
    };
    const onMove = (e) => {
      const p = getPos(e);
      const dx = p.x - drag.sx;
      const dy = p.y - drag.sy;
      if (drag.type === "move") {
        setBox({
          ...drag.sb,
          x: Math.max(0, Math.min(dispSize.w - drag.sb.w, drag.sb.x + dx)),
          y: Math.max(0, Math.min(dispSize.h - drag.sb.h, drag.sb.y + dy)),
        });
      } else {
        const delta  = Math.max(dx, dy);
        const maxSz  = Math.min(dispSize.w - drag.sb.x, dispSize.h - drag.sb.y);
        const newSz  = Math.max(40, Math.min(maxSz, drag.sb.w + delta));
        setBox({ ...drag.sb, w: newSz, h: newSz });
      }
    };
    const onUp = () => setDrag(null);
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup",   onUp);
    return () => { window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
  }, [drag, dispSize]);

  // Apply crop via offscreen canvas
  const confirm = () => {
    const OUT = 512;
    const c   = document.createElement("canvas");
    c.width = OUT; c.height = OUT;
    const ctx = c.getContext("2d");
    const sx  = natSize.w / dispSize.w;
    const sy  = natSize.h / dispSize.h;
    ctx.drawImage(imgRef.current, box.x*sx, box.y*sy, box.w*sx, box.h*sy, 0, 0, OUT, OUT);
    trimAndFill(c); // 透明余白を自動トリミングしてコンテンツを拡大
    onCrop(c.toDataURL("image/png"));
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-black/75 flex items-center justify-center z-[60] p-4">
      <div className="bg-white rounded-2xl shadow-2xl p-6 w-full max-w-md">
        <div className="flex items-center justify-between mb-2">
          <h3 className="text-sm font-bold text-slate-800">✂️ トリミング</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
        </div>
        <p className="text-xs text-slate-500 mb-3">
          枠内ドラッグで移動 ／ 右下角（青線）ドラッグでサイズ変更
        </p>
        <div className="flex justify-center mb-4 rounded-xl overflow-hidden bg-slate-900 min-h-24">
          {loaded ? (
            <canvas
              ref={canvasRef}
              onMouseDown={onMouseDown}
              className="max-w-full cursor-crosshair"
            />
          ) : (
            <div className="h-24 flex items-center justify-center text-slate-400 text-sm w-full">
              読み込み中...
            </div>
          )}
        </div>
        <div className="flex gap-2 justify-end">
          <button onClick={onClose}
            className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">
            キャンセル
          </button>
          <button onClick={confirm} disabled={!loaded}
            className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg disabled:opacity-40 transition-colors">
            適用する
          </button>
        </div>
      </div>
    </div>
  );
}

// ── AppIcon ────────────────────────────────────────────────────────────────────
function AppIcon({ logo, size="md" }) {
  const wh = size==="lg" ? "w-16 h-16" : size==="sm" ? "w-7 h-7" : "w-9 h-9";
  const iw = size==="lg" ? "w-8 h-8"  : size==="sm" ? "w-4 h-4" : "w-5 h-5";
  if (logo) return <img src={logo} alt="logo" className={`${wh} rounded-xl object-contain`} />;
  return (
    <div className={`${wh} bg-blue-600 rounded-xl flex items-center justify-center shrink-0`}>
      <svg className={`${iw} text-white`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
          d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
      </svg>
    </div>
  );
}

// ── LoginScreen ────────────────────────────────────────────────────────────────
const ALLOWED_DOMAIN = "@gmotech.jp";
function LoginScreen({ onLogin, logo }) {
  const [memberId, setMemberId] = useState("");
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  const submit = e => {
    e.preventDefault();
    const id = memberId.trim().toLowerCase();
    // 層②：メールドメイン確認（@gmotech.jp 限定）
    if (!id.endsWith(ALLOWED_DOMAIN)) {
      setErr(`メンバーIDは ${ALLOWED_DOMAIN} のアドレスを入力してください`); return;
    }
    // 層⑥：個別ログイン認証（パスワード照合）
    if (pw !== PASSWORD) { setErr("パスワードが正しくありません"); setPw(""); return; }
    sessionStorage.setItem("teppou_auth", "1");
    sessionStorage.setItem("teppou_member", id);
    // 利用履歴（誰が・いつ）をローカルに記録
    try {
      const logs = JSON.parse(localStorage.getItem("teppou_access_log") || "[]");
      logs.unshift({ member: id, at: new Date().toISOString() });
      localStorage.setItem("teppou_access_log", JSON.stringify(logs.slice(0, 200)));
    } catch {}
    onLogin();
  };
  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-xl border border-slate-200 p-10 w-full max-w-sm">
        <div className="text-center mb-7">
          <div className="flex justify-center mb-4"><AppIcon logo={logo} size="lg" /></div>
          <h1 className="text-2xl font-bold text-slate-800">TEPPOU</h1>
          <p className="text-slate-500 text-sm mt-1">営業管理システム - ログイン</p>
          <p className="text-slate-400 text-xs">TEPPOU Login ／ GMOテック IS部門</p>
        </div>
        <form onSubmit={submit} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">メンバーID（@gmotech.jp）</label>
            <input type="email" value={memberId} autoFocus autoComplete="username"
              onChange={e => { setMemberId(e.target.value); setErr(""); }}
              className="w-full border border-slate-300 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="name@gmotech.jp" />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">パスワード</label>
            <input type="password" value={pw} autoComplete="current-password"
              onChange={e => { setPw(e.target.value); setErr(""); }}
              className="w-full border border-slate-300 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="パスワードを入力" />
            {err && <p className="text-rose-600 text-xs mt-1">{err}</p>}
          </div>
          <button type="submit"
            className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold rounded-lg py-2.5 text-sm transition-colors">
            ログイン
          </button>
        </form>
        <div className="mt-6 pt-4 border-t border-slate-100">
          <p className="text-[10px] text-slate-400 leading-relaxed">
            🔒 オフィスIP制限 ／ @gmotech.jp ドメイン確認 ／ 個別認証 ／ 利用履歴の記録により保護されています。オフィス外からはアクセスできません。
          </p>
        </div>
      </div>
    </div>
  );
}

// ── SettingsModal ──────────────────────────────────────────────────────────────
function SettingsModal({ settings, onSave, onClose, onExportBackup, onImportBackup, onRollback, hasAutoBackup, dataCounts }) {
  const backupFileRef = useRef();
  const [logo,         setLogo]         = useState(settings.logo    || null);
  const [favicon,      setFavicon]      = useState(settings.favicon || null);
  const [backupTimes,  setBackupTimes]  = useState((settings.backupTimes ?? ["10:00","14:00","18:00"]).join(", "));
  const [cropSrc,      setCropSrc]      = useState(null);
  const [cropTarget,   setCropTarget]   = useState(null); // "logo" | "favicon"
  const logoRef    = useRef();
  const faviconRef = useRef();

  const openCrop = (file, target) => {
    const r = new FileReader();
    r.onload = e => { setCropSrc(e.target.result); setCropTarget(target); };
    r.readAsDataURL(file);
  };

  const onCropDone = (dataUrl) => {
    // ロゴとファビコンは常に同期
    setLogo(dataUrl);
    setFavicon(dataUrl);
    setCropSrc(null); setCropTarget(null);
  };

  const UploadBtn = ({ inputRef, target, accept }) => (
    <label className="inline-flex items-center gap-2 bg-blue-50 hover:bg-blue-100 border border-blue-300 text-blue-700 px-3 py-1.5 rounded-lg text-xs font-medium cursor-pointer transition-colors">
      <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
          d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
      </svg>
      アップロード &amp; トリミング
      <input ref={inputRef} type="file" accept={accept || "image/*"} className="hidden"
        onChange={e => e.target.files[0] && openCrop(e.target.files[0], target)} />
    </label>
  );

  return (
    <>
      <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
        <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-6">
          <div className="flex items-center justify-between mb-6">
            <h2 className="text-base font-bold text-slate-800">⚙️ システム設定</h2>
            <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
          </div>

          {/* ロゴ / ファビコン（共通） */}
          <div className="mb-6">
            <p className="text-sm font-semibold text-slate-700 mb-3">アプリロゴ / ファビコン</p>
            <div className="flex items-center gap-4">
              <div className="w-16 h-16 border-2 border-dashed border-slate-300 rounded-xl flex items-center justify-center overflow-hidden bg-slate-50 shrink-0">
                {logo
                  ? <img src={logo} alt="logo" className="w-full h-full object-contain" />
                  : <span className="text-xs text-slate-400 text-center px-1">未設定</span>}
              </div>
              <div className="flex flex-col gap-2">
                <UploadBtn inputRef={logoRef} target="logo" />
                {logo && (
                  <button onClick={() => { setLogo(null); setFavicon(null); }}
                    className="text-xs text-rose-500 hover:text-rose-700 text-left">
                    削除
                  </button>
                )}
              </div>
            </div>
            <p className="text-xs text-slate-400 mt-2">
              PNG / JPEG 対応。アップロード後にトリミングできます。ロゴ・ファビコン（タブアイコン）に同時反映されます。
            </p>
          </div>

          {/* バックアップ時刻設定 */}
          <div className="mb-6">
            <p className="text-sm font-semibold text-slate-700 mb-1">自動バックアップ時刻</p>
            <input
              type="text"
              value={backupTimes}
              onChange={e => setBackupTimes(e.target.value)}
              placeholder="10:00, 14:00, 18:00"
              className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
            <p className="text-xs text-slate-400 mt-1">
              カンマ区切りで時刻を入力（例: 10:00, 14:00, 18:00）。アプリを開いているときに通知が表示されます。
            </p>
          </div>

          {/* データのバックアップと復元 */}
          <div className="mb-6 border border-slate-200 rounded-xl p-4 bg-slate-50">
            <p className="text-sm font-semibold text-slate-700 mb-1">💾 データのバックアップと復元</p>
            <p className="text-xs text-slate-400 mb-3">
              営業リスト（{(dataCounts?.records??0).toLocaleString()}件）と過去商談リスト（{(dataCounts?.pastMgmt??0).toLocaleString()}件）をPCに保存・復元できます。
            </p>

            {/* PCへ保存 */}
            <button onClick={onExportBackup}
              className="flex items-center gap-2 w-full justify-center bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-medium transition-colors mb-2">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/>
              </svg>
              PCへバックアップファイルを保存
            </button>

            {/* 復元 */}
            <label className="flex items-center gap-2 w-full justify-center bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 px-3 py-2 rounded-lg text-xs font-medium cursor-pointer transition-colors mb-2">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/>
              </svg>
              バックアップファイルからデータを復元
              <input ref={backupFileRef} type="file" accept=".json,application/json" className="hidden"
                onChange={e => { if (e.target.files[0]) onImportBackup(e.target.files[0]); e.target.value=""; }} />
            </label>

            {/* 1世代ロールバック */}
            <button onClick={onRollback} disabled={!hasAutoBackup}
              className="flex items-center gap-2 w-full justify-center bg-amber-50 hover:bg-amber-100 border border-amber-300 text-amber-700 px-3 py-2 rounded-lg text-xs font-medium transition-colors disabled:opacity-40 disabled:cursor-not-allowed">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 7v6h6M3 13a9 9 0 109-9"/>
              </svg>
              直前のインポート前の状態に戻す（1世代ロールバック）
            </button>
            <p className="text-xs text-slate-400 mt-1">
              {hasAutoBackup ? "インポート直前の状態が自動退避されています。" : "自動退避データはまだありません（CSV取込時に自動作成）。"}
            </p>
          </div>

          <div className="flex gap-2 justify-end pt-4 border-t border-slate-100">
            <button onClick={onClose}
              className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">
              キャンセル
            </button>
            <button onClick={() => {
              const times = backupTimes.split(",").map(t => t.trim()).filter(t => /^\d{2}:\d{2}$/.test(t));
              onSave({ logo, favicon, backupTimes: times });
              onClose();
            }}
              className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg transition-colors">
              保存する
            </button>
          </div>
        </div>
      </div>

      {cropSrc && (
        <CropModal
          src={cropSrc}
          onCrop={onCropDone}
          onClose={() => { setCropSrc(null); setCropTarget(null); }}
        />
      )}
    </>
  );
}

// ── CSV template definitions ───────────────────────────────────────────────────
const CSV_TEMPLATES = {
  sales: {
    filename: "TEPPOU_営業リスト_フォーマット.csv",
    rows: [
      ["架電日","次回架電日","状況","業種","ソース","企業名","HPサイト","GBP","電話番号","担当者","部署","不在理由","GBPの管理","メモ","店舗数","断り理由","投稿","口コミ","SNS","Insta","Line","FB","Twitter","OS","メール","メアド","GBPサイトURL"],
      ["2026-06-01","2026-06-08","未架電","IT","自社リスト","株式会社サンプル","https://example.com","あり","03-0000-0000","山田太郎","営業部","","○","サンプルメモ","1","","○","★3","○","○","○","○","○","iOS","","sample@example.com","https://maps.google.com/..."],
    ],
  },
  metel: {
    filename: "TEPPOU_MiiTel架電ログ_フォーマット.csv",
    rows: [
      ["ユーザー名","取引先会社名","電話番号","通話種別","タグ","架電日時","メモ"],
      ["櫻井　肇","株式会社サンプル","03-0000-0000","発信（不在）","担当者不在","2026-05-01 10:00:00","サンプルメモ"],
    ],
  },
};

function downloadTemplate(mode) {
  const tmpl = CSV_TEMPLATES[mode];
  const bom  = "﻿"; // UTF-8 BOM（Excelで文字化けしないように）
  const csv  = bom + tmpl.rows.map(row =>
    row.map(c => (c.includes(",") || c.includes('"') || c.includes("\n")) ? `"${c.replace(/"/g,'""')}"` : c).join(",")
  ).join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = tmpl.filename; a.click();
  URL.revokeObjectURL(url);
}

// ── ImportModal ────────────────────────────────────────────────────────────────
function ImportModal({ onImport, onImportPastDeals, onImportMetel, onImportOrders, onClose }) {
  const [mode,      setMode]      = useState("sales");
  const [inputMode, setInputMode] = useState("file");   // "file" | "paste"
  const [pasteText, setPasteText] = useState("");
  const [log,       setLog]       = useState(null);
  const [loading,   setLoading]   = useState(false);
  const [progress,  setProgress]  = useState("");
  const [callMonth, setCallMonth] = useState(() => new Date().toISOString().slice(0,7)); // "YYYY-MM"
  const fileRef = useRef();

  const processRows = (rows) => {
    setLoading(true);
    setProgress(`${rows.length.toLocaleString()} 行を解析中...`);
    setTimeout(() => {
      try {
        if (mode === "order") {
          const result = doImportOrders(rows);
          setLog(result);
          if (result.orders?.length > 0) onImportOrders(result.orders);
        } else if (mode === "past") {
          const result = doImportPastDeals(rows);
          setLog(result);
          if (result.deals?.length > 0) onImportPastDeals(result.deals);
        } else if (mode === "metel") {
          const result = doImportMetel(rows);
          if (result.error) { setLog(result); }
          else {
            const merge = onImportMetel(result.parsed); // {added, updated}
            setLog({ success:true, metel:true, ...merge, filtered:result.filtered, skipped:result.skipped });
          }
        } else {
          const result = doImportSales(rows);
          setLog(result);
          if (result.records?.length > 0) onImport(result.records);
        }
      } catch (ex) {
        setLog({ error: `インポートエラー: ${ex.message}`, records: [] });
      } finally {
        setLoading(false);
        setProgress("");
      }
    }, 60);
  };

  function doImportPastDeals(rows) {
    if (rows.length < 2) return { error:"データ行が不足しています", deals:[] };
    const headers = rows[0];
    const map     = mapPastDealsHeaders(headers);
    if (map.companyName === undefined) {
      return { error:`「企業名」列が見つかりませんでした。ヘッダー: ${headers.filter(h=>h).slice(0,6).join("、")}`, deals:[] };
    }
    const deals = []; let skipped = 0;
    for (const row of rows.slice(1)) {
      if (row.every(c => !c.trim())) continue;
      const company = (row[map.companyName] ?? "").trim();
      if (!company) { skipped++; continue; }
      const g = k => map[k] !== undefined ? (row[map[k]] || "").trim() : "";
      // 自社スタッフ名は dealOwner（商談所有者）へ。相手企業担当者は空白のまま
      deals.push({ companyName: company, pastStatus: g("pastStatus"), lastCallDate: normDate(g("lastCallDate")), dealOwner: g("dealOwner"), memo: g("memo"), importedAt: nowIso() });
    }
    return { success:true, deals, skipped, added: deals.length };
  }

  function doImportOrders(rows) {
    if (rows.length < 2) return { error:"データ行が不足しています", orders:[] };
    const headerIdx = isAggregateRow(rows[0]) ? 1 : 0;
    const headers = rows[headerIdx];
    const map = mapOrderHeaders(headers);
    if (map.companyName === undefined) {
      return { error:`「企業名」列が見つかりませんでした。ヘッダー: ${headers.filter(h=>h).slice(0,6).join("、")}`, orders:[] };
    }
    const orders = []; let skipped = 0;
    for (const row of rows.slice(headerIdx+1)) {
      if (row.every(c => !c.trim())) continue;
      const company = (row[map.companyName]??"").trim();
      if (!company) { skipped++; continue; }
      const g = k => map[k]!==undefined ? (row[map[k]]||"").trim() : "";
      orders.push({
        id: genId(), companyName: company,
        orderDate: normDate(g("orderDate")) || "",
        plan: g("plan"), amount: g("amount"),
        assignee: g("assignee"),
        payment: g("payment") || "未入金",
        startDate: normDate(g("startDate")) || "",
        memo: g("memo"), updatedAt: nowIso(),
      });
    }
    return { success:true, orders, skipped, added: orders.length };
  }

  const process = text => processRows(parseCSV(text));

  const handleFile = e => {
    const file = e.target.files[0];
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    setLoading(true);
    setProgress("ファイル読み込み中...");
    if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = ev => {
        setProgress("Excelを解析中...");
        setTimeout(() => {
          try {
            const wb = XLSX.read(ev.target.result, { type: "array", cellDates: true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, {
              header: 1, raw: true, defval: "",
            }).map(row => row.map(c => {
              if (c instanceof Date) return c.toISOString().slice(0, 10); // Date型→YYYY-MM-DD
              return String(c ?? "").trim();
            }));
            processRows(rows);
          } catch (ex) {
            setLog({ error: `ファイル読み込みエラー: ${ex.message}`, records: [] });
            setLoading(false); setProgress("");
          }
        }, 60);
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = ev => process(ev.target.result);
      reader.readAsText(file, "UTF-8");
    }
    e.target.value = "";
  };

  const handlePaste = () => {
    if (!pasteText.trim()) return;
    process(pasteText);
  };

  function doImportSales(rows) {
    if (rows.length < 2) return { error:"データ行が不足しています", records:[] };
    const headerIdx = isAggregateRow(rows[0]) ? 1 : 0;
    const headers   = rows[headerIdx];
    const map       = mapSalesHeaders(headers);

    // 企業名列が見つからない場合は早期エラー
    if (map.companyName === undefined) {
      const detected = headers.filter(h => h.trim()).slice(0, 8).join("、");
      return {
        error: `「企業名」列が見つかりませんでした。\n検出されたヘッダー: ${detected}\n\nヒント: 列名を「企業名」「会社名」「法人名」「取引先名」のいずれかにしてください。`,
        records: [],
      };
    }

    const records   = [];
    let skipped = 0;
    for (const row of rows.slice(headerIdx + 1)) {
      if (row.every(c => !c.trim())) continue;
      const company = row[map.companyName] ?? "";
      if (!company.trim()) { skipped++; continue; }
      const g = (k) => map[k] !== undefined ? (row[map[k]] || "").trim() : "";
      records.push({
        id:            genId(),
        companyName:   company.trim(),
        phone:         g("phone"),
        email:         g("email"),
        mailFlag:      g("mailFlag"),
        hpSite:        g("hpSite"),
        gbp:           g("gbp"),
        gbpSiteUrl:    g("gbpSiteUrl"),
        gbpManagement: g("gbpManagement"),
        status:        g("status") || "未架電",
        assignee:      g("assignee"),
        department:    g("department"),
        industry:      g("industry"),
        leadSource:    g("leadSource"),
        absenceReason: g("absenceReason"),
        memo:          g("memo"),
        storeCount:    g("storeCount"),
        refusalReason: g("refusalReason"),
        posting:       g("posting"),
        review:        g("review"),
        sns:           g("sns"),
        instagram:     g("instagram"),
        line:          g("line"),
        facebook:      g("facebook"),
        twitter:       g("twitter"),
        os:            g("os"),
        nextCallDate:  normDate(g("nextCallDate")),
        lastCallDate:  normDate(g("lastCallDate")),
        importedAt: nowIso(), updatedAt: nowIso(), source:"csv", leadAddedDate: getToday(),
      });
    }
    return { success:true, records, skipped, autoSkip: headerIdx === 1 };
  }

  function doImportMetel(rows) {
    if (rows.length < 2) return { error:"データ行が不足しています", parsed:[] };
    const headers = rows[0];
    const col = patterns => {
      for (let i = 0; i < headers.length; i++) {
        const h = headers[i].replace(/[\s　]/g,"");
        if (patterns.some(p => h.includes(p))) return i;
      }
      return -1;
    };
    const cAssignee = col(["ユーザー名","担当者","オペレーター","エージェント","架電者"]);
    const cCompany  = col(["取引先会社名","企業名","会社名","顧客名","取引先名"]);
    const cPhone    = col(["電話番号","TEL","電話"]);
    const cCallType = col(["通話種別","架電種別","種別"]);
    const cTags     = col(["タグ","ラベル"]);
    const cMemo     = col(["メモ","備考","コメント"]);
    const parsed = []; let filtered = 0, skipped = 0;

    for (const row of rows.slice(1)) {
      if (row.every(c => !c.trim())) continue;
      const operator = cAssignee >= 0 ? (row[cAssignee]||"").trim() : "";
      const isMember = IS_MEMBERS.some(m => {
        const mn = normName(m), an = normName(operator);
        return an === mn || an.includes(mn) || mn.includes(an);
      });
      if (!isMember) { filtered++; continue; }
      const company = cCompany >= 0 ? (row[cCompany]||"").trim() : "";
      if (!company) { skipped++; continue; }
      const callType = cCallType >= 0 ? (row[cCallType]||"").trim() : "";
      const tags     = cTags     >= 0 ? (row[cTags]    ||"").trim() : "";
      const status   = convertMitelStatus(callType, tags); // null可（情報不足時は更新しない）
      const baseMemo = cMemo >= 0 ? (row[cMemo]||"").trim() : "";
      const memo     = tags ? `【MiiTelタグ】${tags}${baseMemo ? "\n"+baseMemo : ""}` : baseMemo;
      parsed.push({ company, operator, status, importMonth: callMonth, memo, phone: cPhone>=0?(row[cPhone]||"").trim():"" });
    }
    return { success:true, parsed, filtered, skipped };
  }

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg p-6 relative">

        {/* Loading overlay */}
        {loading && (
          <div className="absolute inset-0 bg-white/90 flex flex-col items-center justify-center rounded-2xl z-10 gap-3">
            <div className="w-10 h-10 border-4 border-blue-600 border-t-transparent rounded-full animate-spin" />
            <p className="text-sm font-semibold text-slate-700">{progress || "処理中..."}</p>
            <p className="text-xs text-slate-400">大量データは数秒かかる場合があります</p>
          </div>
        )}

        <div className="flex items-center justify-between mb-5">
          <h2 className="text-base font-bold text-slate-800">インポート</h2>
          <button onClick={onClose} disabled={loading} className="text-slate-400 hover:text-slate-700 text-xl leading-none disabled:opacity-30">×</button>
        </div>

        {/* Mode selection */}
        <div className="space-y-3 mb-5">
          {[
            { value:"sales", icon:"📁", title:"自分の営業リストを取り込む",
              desc:"企業名・取引先名・会社名を自動マッピング。作成者→担当者、完了予定日→架電日として取り込み。" },
            { value:"metel", icon:"📞", title:"MiiTel架電ログを取り込む",
              desc:"ISメンバー10名に自動絞り込み。未登録企業は新規追加（追加者記録）、既登録は別担当者・最新架電日を更新。" },
            { value:"past",  icon:"📜", title:"過去商談リスト（プル照合用）を取り込む",
              desc:"企業名・過去の状況・担当者・メモを読込。メインリストと自動照合してバッジ表示します。" },
            { value:"order", icon:"🛍️", title:"受注案件リストを取り込む",
              desc:"企業名・受注日・商材・金額・担当者を自動マッピング。受注案件管理へ保存（同名は上書き）。" },
          ].map(opt => (
            <label key={opt.value}
              className={`flex items-start gap-3 p-3 rounded-xl border-2 cursor-pointer transition-colors
                ${mode===opt.value ? "border-blue-400 bg-blue-50" : "border-slate-200 hover:border-slate-300"}`}
              onClick={() => { setMode(opt.value); setLog(null); }}>
              <div className={`w-4 h-4 mt-0.5 rounded-full border-2 flex items-center justify-center shrink-0
                ${mode===opt.value ? "border-blue-600 bg-blue-600" : "border-slate-400"}`}>
                {mode===opt.value && <div className="w-2 h-2 rounded-full bg-white" />}
              </div>
              <div>
                <div className="text-sm font-semibold text-slate-800">{opt.icon} {opt.title}</div>
                <div className="text-xs text-slate-500 mt-0.5">{opt.desc}</div>
              </div>
            </label>
          ))}
        </div>

        {/* MiiTel: 架電月の選択 */}
        {mode === "metel" && (
          <div className="flex items-center gap-3 bg-blue-50 border border-blue-200 rounded-xl px-4 py-3 mb-4">
            <span className="text-xs font-semibold text-blue-700 shrink-0">📅 取込月</span>
            <input type="month" value={callMonth} onChange={e => setCallMonth(e.target.value)}
              className="border border-blue-300 rounded-lg px-3 py-1.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500"/>
            <span className="text-xs text-blue-600">取り込むログを「<strong>{callMonth.replace("-","/")}</strong>」の取込月として記録します</span>
          </div>
        )}

        {/* Template download */}
        <div className="flex items-center justify-between bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 mb-4">
          <div>
            <p className="text-xs font-semibold text-slate-700">📄 入力フォーマット（テンプレート）</p>
            <p className="text-xs text-slate-400 mt-0.5">ダウンロードして記入後、そのままインポートできます</p>
          </div>
          <button onClick={() => downloadTemplate(mode)}
            className="flex items-center gap-1.5 bg-white hover:bg-blue-50 border border-slate-300 hover:border-blue-400 text-slate-700 hover:text-blue-700 px-3 py-1.5 rounded-lg text-xs font-medium transition-colors shrink-0 ml-3">
            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
            </svg>
            CSV DL
          </button>
        </div>

        {/* Input mode tabs */}
        <div className="flex rounded-lg border border-slate-200 overflow-hidden mb-4 text-xs font-medium">
          {[
            { id:"file",  label:"📂 ファイル選択" },
            { id:"paste", label:"📋 テキスト貼り付け" },
          ].map(tab => (
            <button key={tab.id}
              onClick={() => { setInputMode(tab.id); setLog(null); }}
              className={`flex-1 py-2 transition-colors
                ${inputMode===tab.id
                  ? "bg-blue-600 text-white"
                  : "text-slate-600 hover:bg-slate-50"}`}>
              {tab.label}
            </button>
          ))}
        </div>

        {/* File drop zone */}
        {inputMode === "file" && (
          <label className="flex flex-col items-center justify-center gap-3 border-2 border-dashed border-slate-300 hover:border-blue-400 rounded-xl p-8 cursor-pointer transition-colors mb-4 bg-slate-50 hover:bg-blue-50">
            <svg className="w-9 h-9 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5}
                d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
            </svg>
            <span className="text-sm text-slate-600 font-medium">ファイルを選択</span>
            <span className="text-xs text-slate-400">Excel (.xlsx) / CSV (.csv) 対応</span>
            <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleFile} />
          </label>
        )}

        {/* Paste zone */}
        {inputMode === "paste" && (
          <div className="mb-4">
            <textarea
              value={pasteText}
              onChange={e => { setPasteText(e.target.value); setLog(null); }}
              placeholder={"ここにCSVテキストをペースト\n例: 企業名,電話番号,メモ\n株式会社〇〇,03-1234-5678,メモ内容"}
              rows={7}
              className="w-full border border-slate-300 rounded-xl px-3 py-2.5 text-xs font-mono text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none bg-slate-50 placeholder:text-slate-300"
            />
            <div className="flex justify-between items-center mt-2">
              <span className="text-xs text-slate-400">
                {pasteText.trim() ? `${pasteText.trim().split(/\r?\n/).length} 行` : ""}
              </span>
              <div className="flex gap-2">
                <button onClick={() => { setPasteText(""); setLog(null); }}
                  disabled={!pasteText}
                  className="px-3 py-1.5 text-xs text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50 disabled:opacity-30 transition-colors">
                  クリア
                </button>
                <button onClick={handlePaste}
                  disabled={!pasteText.trim()}
                  className="px-4 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg disabled:opacity-40 transition-colors">
                  取り込む
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Log */}
        {log && (
          <div className={`text-xs rounded-lg px-4 py-3 mb-4
            ${log.error ? "bg-red-50 border border-red-200 text-red-700" : "bg-green-50 border border-green-200 text-green-700"}`}>
            {log.error ? log.error : (
              <span>
                {log.metel
                  ? <>✅ MiiTel取込完了（取込月: {callMonth.replace("-","/")}）: 新規追加 <strong>{(log.added||0).toLocaleString()}件</strong> ／ 既存更新 <strong>{(log.updated||0).toLocaleString()}件</strong></>
                  : log.orders
                  ? <>✅ 受注案件インポート完了: <strong>{log.orders.length}件</strong>（同一企業名は上書き更新）</>
                  : log.deals
                  ? <>✅ 過去商談インポート完了: <strong>{log.deals.length}件</strong>（同一企業名は上書き更新）</>
                  : <>✅ インポート完了: <strong>{log.records.length}件</strong>追加</>}
                {log.filtered  > 0 && ` ／ ISメンバー以外: ${log.filtered}件除外`}
                {log.skipped   > 0 && ` ／ スキップ: ${log.skipped}件`}
                {log.autoSkip      && " ／ 1行目を集計行として自動スキップ"}
              </span>
            )}
          </div>
        )}

        <div className="flex justify-end">
          <button onClick={onClose}
            className="px-4 py-2 text-sm text-slate-600 border border-slate-200 rounded-lg hover:bg-slate-50">
            閉じる
          </button>
        </div>
      </div>
    </div>
  );
}

// ── StatusBulkUpdateModal ──────────────────────────────────────────────────────
// フェーズ（Excel）→ 状況（TEPPOU）変換。null=更新しない
function phaseToStatus(phase) {
  const p = String(phase || "").trim();
  if (/^0?[1-5]\./.test(p)) return "4.商談中";     // 01〜05 → 商談中
  if (/^0?6\./.test(p))      return "8.当社契約";   // 06 → 当社契約
  return null;                                       // 受注後取消など → 更新しない
}

function StatusBulkUpdateModal({ records, onUpdate, onClose }) {
  const [targetStatus, setTargetStatus] = useState("8.当社契約");
  const [matches, setMatches]           = useState(null); // { names, matched, plan, phaseMode }
  const [log, setLog]                   = useState(null);
  const [loading, setLoading]           = useState(false);
  const fileRef = useRef();

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true); setLog(null); setMatches(null);
    const readRows = (rows) => {
      const headers = rows[0] || [];
      const norm = h => String(h).replace(/[\s　]/g,"");
      const nameIdx  = headers.findIndex(h => /取引先名?|企業名|会社名|法人名/.test(norm(h)));
      const phaseIdx = headers.findIndex(h => /フェーズ|phase|状況|ステータス/.test(norm(h)));
      const nIdx = nameIdx >= 0 ? nameIdx : 1;
      const phaseMode = phaseIdx >= 0;

      // 企業名→フェーズ のマップ（フェーズ列がある場合）
      const phaseByName = new Map();
      const names = [];
      rows.slice(1).forEach(r => {
        const nm = String(r[nIdx]??"").trim();
        if (!nm) return;
        names.push(nm);
        if (phaseMode) phaseByName.set(normName(nm), String(r[phaseIdx]??"").trim());
      });
      const uniqNames = [...new Set(names)];

      // リスト内の一致企業を抽出し、フェーズ→新ステータスを計算
      const plan = []; // { record, newStatus }
      records.forEach(rec => {
        const rn = normName(rec.companyName);
        const hit = uniqNames.find(n => { const nn = normName(n); return nn === rn || nn.includes(rn) || rn.includes(nn); });
        if (!hit) return;
        let newStatus;
        if (phaseMode) {
          const ph = phaseByName.get(normName(hit));
          newStatus = phaseToStatus(ph);
          if (!newStatus) return; // 受注後取消など → 対象外
        } else {
          newStatus = targetStatus;
        }
        plan.push({ record: rec, newStatus });
      });

      setMatches({ uniqCount: uniqNames.length, plan, phaseMode });
      setLoading(false);
    };
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = ev => {
        const wb = XLSX.read(ev.target.result, { type:"array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        readRows(XLSX.utils.sheet_to_json(ws, { header:1, defval:"" }).map(r => r.map(c => String(c??'').trim())));
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = ev => readRows(parseCSV(ev.target.result));
      reader.readAsText(file, "UTF-8");
    }
    e.target.value = "";
  };

  const apply = () => {
    if (!matches?.plan.length) return;
    // id → 新ステータス のMap
    const map = new Map(matches.plan.map(p => [p.record.id, p.newStatus]));
    onUpdate(map);
    setLog({ updated: map.size });
  };

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg p-6">
        <div className="flex items-center justify-between mb-5">
          <h2 className="text-base font-bold text-slate-800">📝 ステータス一括更新</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
        </div>

        <div className="mb-4">
          <p className="text-xs text-slate-500 mb-2">企業名＋フェーズ列のあるExcel/CSVなら、フェーズに応じて状況を自動変換して一括更新します（フェーズ列がなければ下で選んだ状況に統一）。</p>
          <label className="flex items-center gap-2 bg-blue-50 hover:bg-blue-100 border border-blue-300 text-blue-700 px-3 py-2 rounded-lg text-xs font-medium cursor-pointer w-fit transition-colors">
            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
            </svg>
            ファイルを選択
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile} disabled={loading}/>
          </label>
          <p className="text-[10px] text-slate-400 mt-2">フェーズ変換: 01〜05.◯◯ → 4.商談中 ／ 06.決裁完了・契約合意 → 8.当社契約 ／ 受注後取消 → 対象外</p>
        </div>

        {loading && <p className="text-xs text-slate-500 mb-4 flex items-center gap-2"><span className="w-4 h-4 border-2 border-blue-500 border-t-transparent rounded-full animate-spin inline-block"/>照合中...</p>}

        {matches && !log && (
          <>
            <div className={`rounded-xl px-4 py-3 mb-4 text-sm ${matches.plan.length>0?"bg-teal-50 border border-teal-200 text-teal-800":"bg-slate-50 border border-slate-200 text-slate-500"}`}>
              ファイル内 <strong>{matches.uniqCount}</strong> 社 → 更新対象 <strong>{matches.plan.length}</strong> 社
              {matches.phaseMode && <span className="text-xs ml-1">（フェーズ自動変換モード）</span>}
              {matches.plan.length === 0 && <span className="text-xs ml-2">（一致または変換対象なし）</span>}
            </div>

            {matches.plan.length > 0 && (
              <>
                <div className="max-h-40 overflow-y-auto mb-4 border border-slate-200 rounded-xl">
                  <table className="w-full text-xs">
                    <thead className="bg-slate-50 sticky top-0">
                      <tr>
                        <th className="px-3 py-1.5 text-left text-slate-500">企業名</th>
                        <th className="px-3 py-1.5 text-left text-slate-500">現在</th>
                        <th className="px-3 py-1.5 text-left text-slate-500">→ 変更後</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {matches.plan.slice(0,12).map(({record:r, newStatus}) => (
                        <tr key={r.id}>
                          <td className="px-3 py-1.5 text-slate-700 truncate max-w-[140px]">{r.companyName}</td>
                          <td className="px-3 py-1.5"><StatusBadge status={r.status}/></td>
                          <td className="px-3 py-1.5"><StatusBadge status={newStatus}/></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {matches.plan.length > 12 && <p className="text-xs text-slate-400 text-center py-1">他 {matches.plan.length-12} 社…</p>}
                </div>

                {!matches.phaseMode && (
                  <div className="mb-4">
                    <label className="block text-xs text-slate-500 mb-1">変更後のステータス（フェーズ列なし時）</label>
                    <select value={targetStatus} onChange={e=>setTargetStatus(e.target.value)}
                      className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
                      {Object.keys(STATUS_CFG).map(s=><option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                )}
              </>
            )}
          </>
        )}

        {log && (
          <div className="bg-green-50 border border-green-200 text-green-800 rounded-xl px-4 py-3 mb-4 text-sm">
            ✅ <strong>{log.updated}社</strong>のステータスを更新しました
          </div>
        )}

        <div className="flex gap-2 justify-end">
          <button onClick={onClose} className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">
            {log ? "閉じる" : "キャンセル"}
          </button>
          {!log && matches?.plan.length > 0 && (
            <button onClick={apply}
              className="px-4 py-2 bg-teal-600 hover:bg-teal-700 text-white text-sm font-semibold rounded-lg transition-colors">
              {matches.plan.length}社を更新
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// ── DuplicateModal ─────────────────────────────────────────────────────────────
// ステータス優先度（低い数字ほど価値が高い → 残す候補）
const DUPE_PRIORITY = {
  "9.アポ獲得":1, "0.日程調整":2, "1.高確度":3, "4.商談中":4,
  "5.メール送付":5, "6.コネクト（改）":6, "7.コネクト（無）":7,
  "2.優先":8, "4.別担当架電":9, "4.受付カット":10,
  "8.当社契約":11, "8.不要":12, "3.並":13,
  "リードなし":98, "未架電":99, "":100,
};
const dupePrio = s => DUPE_PRIORITY[s] ?? 50;

// 優先度→更新日の順でソート（先頭＝残す）
function sortGroupForDupe(rs) {
  return [...rs].sort((a, b) => {
    // ① 当社契約（受注）は最優先で残す
    const aWon = a.status === "8.当社契約";
    const bWon = b.status === "8.当社契約";
    if (aWon !== bWon) return aWon ? -1 : 1;

    // ② 未架電は最後（削除候補）
    const aUncalled = a.status === "未架電";
    const bUncalled = b.status === "未架電";
    if (aUncalled !== bUncalled) return aUncalled ? 1 : -1;

    // ③ それ以外は架電日（lastCallDate）が新しい順で残す
    const aDate = normDate(a.lastCallDate) || normDate(a.updatedAt) || (a.importedAt || "");
    const bDate = normDate(b.lastCallDate) || normDate(b.updatedAt) || (b.importedAt || "");
    return bDate.localeCompare(aDate);
  });
}

function DuplicateModal({ records, onClean, onClose, sortFn, renderExtra }) {
  const sortGroup = sortFn || sortGroupForDupe;
  const [step,      setStep]      = useState("select"); // "select" | "confirm"
  const [deleteSet, setDeleteSet] = useState(() => {
    const s = new Set();
    const g = {};
    records.forEach(r => {
      const k = normName(r.companyName); if (!k) return;
      (g[k] = g[k]||[]).push(r);
    });
    Object.values(g).filter(rs => rs.length > 1).forEach(rs => {
      sortGroup(rs).slice(1).forEach(r => s.add(r.id));
    });
    return s;
  });

  const groups = (() => {
    const g = {};
    records.forEach(r => {
      const k = normName(r.companyName); if (!k) return;
      (g[k] = g[k]||[]).push(r);
    });
    return Object.values(g)
      .filter(rs => rs.length > 1)
      .map(rs => sortGroup(rs));
  })();

  const toggle = id => setDeleteSet(p => { const n = new Set(p); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const toDelete = records.filter(r => deleteSet.has(r.id));

  // ── Confirm screen ──
  if (step === "confirm") return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg p-6 flex flex-col max-h-[80vh]">
        <div className="flex items-center justify-between mb-4">
          <h2 className="text-base font-bold text-slate-800">🗑️ 削除確認</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
        </div>
        <p className="text-sm text-slate-600 mb-4">
          以下 <strong className="text-red-600">{toDelete.length}件</strong> を削除します。この操作は元に戻せません。
        </p>
        <div className="overflow-y-auto flex-1 mb-4 space-y-1 pr-1">
          {toDelete.map(r => (
            <div key={r.id} className="flex items-center gap-2 px-3 py-2 bg-red-50 rounded-lg border border-red-200 text-xs">
              <span className="font-semibold text-red-800 flex-1">{r.companyName}</span>
              <span className="text-slate-500">{(r.updatedAt||r.importedAt||"").slice(0,10)}</span>
              <StatusBadge status={r.status} />
            </div>
          ))}
        </div>
        <div className="flex gap-2 justify-end pt-4 border-t border-slate-100">
          <button onClick={() => setStep("select")}
            className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">
            ← 戻る
          </button>
          <button onClick={() => { onClean(Array.from(deleteSet)); onClose(); }}
            className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white text-sm font-semibold rounded-lg transition-colors">
            {toDelete.length}件を削除する
          </button>
        </div>
      </div>
    </div>
  );

  // ── Select screen ──
  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl p-6 flex flex-col max-h-[85vh]">
        <div className="flex items-center justify-between mb-2">
          <h2 className="text-base font-bold text-slate-800">重複クレンジング</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
        </div>
        <p className="text-sm text-slate-500 mb-4">
          {groups.length === 0
            ? "✅ 重複データは見つかりませんでした。"
            : <>チェックをつけたレコードを削除します。<strong className="text-red-600">{deleteSet.size}件</strong>が選択中。</>}
        </p>
        {groups.length > 0 && (
          <div className="overflow-y-auto flex-1 space-y-3 mb-4 pr-1">
            {groups.map((g, gi) => (
              <div key={gi} className="border border-slate-200 rounded-xl overflow-hidden">
                <div className="bg-slate-50 px-3 py-2 font-semibold text-sm text-slate-700 border-b border-slate-200">
                  {g[0].companyName}
                  <span className="ml-2 text-xs font-normal text-slate-400">{g.length}件の重複</span>
                </div>
                <div className="divide-y divide-slate-100">
                  {g.map((r, i) => (
                    <label key={r.id}
                      className={`flex items-center gap-3 px-3 py-2.5 cursor-pointer hover:bg-slate-50 transition-colors
                        ${deleteSet.has(r.id) ? "bg-red-50" : ""}`}>
                      <input type="checkbox" checked={deleteSet.has(r.id)} onChange={() => toggle(r.id)}
                        className="rounded border-slate-300 text-red-600 shrink-0" />
                      <span className={`text-xs font-medium shrink-0 w-12 ${deleteSet.has(r.id) ? "text-red-500" : "text-green-600"}`}>
                        {deleteSet.has(r.id) ? "× 削除" : "✓ 残す"}
                      </span>
                      <span className="text-xs text-slate-500 shrink-0">{(r.updatedAt||r.importedAt||"").slice(0,10)}</span>
                      <span className="text-xs text-slate-600 shrink-0">{r.phone||"—"}</span>
                      <StatusBadge status={r.status} />
                      {r.assignee && <span className="text-xs text-slate-500">{r.assignee}</span>}
                      {renderExtra && renderExtra(r)}
                    </label>
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}
        <div className="flex gap-2 justify-end pt-4 border-t border-slate-100">
          <button onClick={onClose}
            className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">
            キャンセル
          </button>
          {groups.length > 0 && (
            <button onClick={() => setStep("confirm")} disabled={deleteSet.size === 0}
              className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white text-sm font-semibold rounded-lg disabled:opacity-40 transition-colors">
              選択した {deleteSet.size}件 を削除する →
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// ── RecordFormModal (shared by New & Edit) ─────────────────────────────────────
function RecordFormModal({ initial, title, onSave, onClose, onDelete, pastDeal, storeEstimate, ordered, onCopyToOrder }) {
  const [form, setForm] = useState({
    companyName:"", phone:"", email:"", mailFlag:"",
    hpSite:"", gbp:"", gbpSiteUrl:"", gbpManagement:"",
    status:"未架電", assignee:"", department:"",
    lastCallDate:"", nextCallDate:"",
    industry:"", leadSource:"", absenceReason:"", refusalReason:"",
    storeCount:"", posting:"", review:"",
    sns:"", instagram:"", line:"", facebook:"", twitter:"", os:"",
    memo:"",
    ...initial,
  });
  const upd = (k,v) => setForm(f => ({ ...f, [k]:v }));
  const logoFileRef = useRef();
  // 取り込んだ画像をCanvasで24×24pxに圧縮しBase64でlogo_urlへ固定保存（容量極小）
  const onImagePicked = useCallback((file) => {
    if (!file || !file.type || !file.type.startsWith("image/")) return;
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const c = document.createElement("canvas");
        c.width = 24; c.height = 24;
        const ctx = c.getContext("2d");
        ctx.clearRect(0, 0, 24, 24);
        ctx.drawImage(img, 0, 0, 24, 24);
        try { setForm(f => ({ ...f, logoUrl: c.toDataURL("image/png") })); } catch {}
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  }, []);
  // モーダル内で Ctrl+V → クリップボードの画像を取り込み（画像が無ければ通常の貼り付けを許可）
  const onModalPaste = useCallback((e) => {
    const items = e.clipboardData && e.clipboardData.items;
    if (!items) return;
    for (const it of items) {
      if (it.type && it.type.startsWith("image/")) {
        const file = it.getAsFile();
        if (file) { e.preventDefault(); onImagePicked(file); return; }
      }
    }
  }, [onImagePicked]);
  const txt = (key, label, colSpan=1, type="text") => (
    <div className={colSpan===2 ? "col-span-2" : ""}>
      <label className="block text-xs text-slate-500 mb-1">{label}</label>
      <input type={type} value={form[key]||""}
        onChange={e => upd(key, e.target.value)}
        className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
    </div>
  );
  const SectionLabel = ({children}) => (
    <div className="col-span-2 pt-2 pb-0.5 border-b border-slate-100">
      <span className="text-xs font-semibold text-slate-400 uppercase tracking-wide">{children}</span>
    </div>
  );

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4" onPaste={onModalPaste}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl flex flex-col max-h-[92vh]">
        <div className="flex items-center justify-between px-6 pt-5 pb-4 border-b border-slate-100 shrink-0">
          <h2 className="text-base font-bold text-slate-800">{title}</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
        </div>
        <div className="overflow-y-auto flex-1 px-6 py-4">
          <div className="grid grid-cols-2 gap-3">

            {/* 企業ロゴ（URL / 画像アップロード / Ctrl+V貼り付け）— 最上段 */}
            <div className="col-span-2">
              <label className="block text-xs text-slate-500 mb-1">🌐 企業ロゴ（URL / 画像アップロード / Ctrl+V貼り付け）</label>
              <div className="flex items-center gap-2">
                <CompanyLogo logoUrl={form.logoUrl} />
                <input type="text" value={form.logoUrl || ""}
                  onChange={e => upd("logoUrl", e.target.value)}
                  placeholder="https://... または画像を貼り付け／アップロード（未登録は白紙）"
                  className="flex-1 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                <input ref={logoFileRef} type="file" accept="image/*" className="hidden"
                  onChange={e => { onImagePicked(e.target.files && e.target.files[0]); e.target.value = ""; }} />
                <button type="button" onClick={() => logoFileRef.current && logoFileRef.current.click()}
                  className="px-2.5 py-2 text-xs text-blue-600 border border-blue-300 rounded-lg hover:bg-blue-50 whitespace-nowrap">
                  🖼️ 画像
                </button>
                {form.logoUrl && (
                  <button type="button" onClick={() => upd("logoUrl", "")}
                    className="px-2.5 py-2 text-xs text-slate-400 hover:text-slate-700 border border-slate-200 rounded-lg hover:bg-slate-50 whitespace-nowrap">
                    クリア
                  </button>
                )}
              </div>
              <p className="text-[11px] text-slate-400 mt-1">画像は <strong>Ctrl+V</strong> でこのモーダルに直接貼り付け、または「🖼️ 画像」から選択できます。<strong>24×24px</strong> に圧縮して保存します（URL手入力も可）。未登録の場合は白紙の枠を表示します。</p>
            </div>

            {/* 企業情報 */}
            <SectionLabel>企業情報</SectionLabel>
            <div className="col-span-2">
              <label className="flex items-center gap-1.5 text-xs text-slate-500 mb-1">
                <CompanyLogo logoUrl={form.logoUrl} url={form.hpSite} name={form.companyName} />
                企業名 <span className="text-rose-500">*</span>
              </label>
              <input type="text" value={form.companyName} autoFocus
                onChange={e => upd("companyName", e.target.value)}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
            </div>
            {txt("phone",      "電話番号")}
            {txt("department", "部署")}
            {txt("industry",   "業種")}
            <div>
              <label className="block text-xs text-slate-500 mb-1">ソース</label>
              <select value={form.leadSource||""} onChange={e => upd("leadSource", e.target.value)}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
                <option value="">— 未選択 —</option>
                {Object.keys(LEAD_SOURCE_CFG).map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              {form.leadSource && (
                <div className="mt-1.5"><LeadSourceBadge source={form.leadSource} /></div>
              )}
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">
                店舗数
                {storeEstimate && storeEstimate.estimated && storeEstimate.value !== null && (
                  <span className="ml-2 text-[10px] text-slate-400 font-normal">自動分析: {storeEstimate.value.toLocaleString()}（仮）</span>
                )}
              </label>
              <input type="text" value={form.storeCount||""}
                onChange={e => upd("storeCount", e.target.value)}
                placeholder={storeEstimate && storeEstimate.estimated && storeEstimate.value !== null ? `${storeEstimate.value}（仮・未確定）` : ""}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">リード追加日</label>
              <input type="date" value={normDate(form.leadAddedDate) || getToday()}
                onChange={e => upd("leadAddedDate", e.target.value)}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
            </div>

            {/* 架電管理 */}
            <SectionLabel>架電管理</SectionLabel>
            <div>
              <label className="block text-xs text-slate-500 mb-1">架電日</label>
              <input type="date" value={normDate(form.lastCallDate)||""}
                onChange={e => upd("lastCallDate", e.target.value)}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">次回架電日</label>
              <input type="date" value={normDate(form.nextCallDate)||""}
                onChange={e => upd("nextCallDate", e.target.value)}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">状況</label>
              <select value={form.status||"未架電"} onChange={e => upd("status", e.target.value)}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
                {Object.keys(STATUS_CFG).map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            {txt("assignee",      "担当者")}
            <div>
              <label className="block text-xs text-slate-500 mb-1">不在理由</label>
              <select value={form.absenceReason||""} onChange={e => upd("absenceReason", e.target.value)}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
                <option value="">— 未選択 —</option>
                {Object.keys(ABSENCE_REASON_CFG).map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              {form.absenceReason && <div className="mt-1.5"><AbsenceReasonBadge reason={form.absenceReason} /></div>}
            </div>
            {txt("refusalReason", "断り理由")}

            {/* Web / GBP */}
            <SectionLabel>Web / GBP</SectionLabel>
            {txt("hpSite",        "HPサイト",   2)}
            {txt("gbp",           "GBP")}
            {txt("gbpManagement", "GBPの管理")}
            {txt("gbpSiteUrl",    "GBPサイトURL", 2)}
            {txt("posting",       "投稿")}
            {txt("review",        "口コミ")}

            {/* SNS */}
            <SectionLabel>SNS</SectionLabel>
            {txt("sns",       "SNS")}
            {txt("instagram", "Insta")}
            {txt("line",      "Line")}
            {txt("facebook",  "FB")}
            {txt("twitter",   "Twitter")}
            {txt("os",        "OS")}

            {/* メール */}
            <SectionLabel>メール</SectionLabel>
            {txt("mailFlag", "メール")}
            {txt("email",    "メアド", 1, "email")}

            {/* メモ */}
            <SectionLabel>メモ</SectionLabel>
            <div className="col-span-2">
              <label className="block text-xs text-slate-500 mb-1">メモ</label>
              <textarea value={form.memo||""} onChange={e => upd("memo", e.target.value)} rows={4}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none" />
            </div>

            {/* 過去商談プル表示 */}
            {pastDeal && (
              <div className="col-span-2 mt-2">
                <div className="bg-purple-50 border border-purple-200 rounded-xl p-4">
                  <p className="text-xs font-bold text-purple-700 mb-3">📜【プル照合】過去の商談・架電履歴</p>
                  <div className="grid grid-cols-2 gap-3 text-xs">
                    {pastDeal.pastStatus && (
                      <div>
                        <span className="text-slate-400 block mb-0.5">過去の状況</span>
                        <span className="font-semibold text-purple-800 bg-purple-100 px-2 py-0.5 rounded">{pastDeal.pastStatus}</span>
                      </div>
                    )}
                    {pastDeal.lastCallDate && (
                      <div>
                        <span className="text-slate-400 block mb-0.5">過去の最終架電日</span>
                        <span className="font-semibold text-slate-700">{fmtDate(pastDeal.lastCallDate)}</span>
                      </div>
                    )}
                    {pastDeal.dealOwner && (
                      <div>
                        <span className="text-slate-400 block mb-0.5">当時の自社「担当者」</span>
                        <span className="inline-flex items-center gap-1 font-semibold text-indigo-700 bg-indigo-100 border border-indigo-200 px-2 py-0.5 rounded">
                          👤 {pastDeal.dealOwner}
                        </span>
                      </div>
                    )}
                    {pastDeal.memo && (
                      <div className="col-span-2">
                        <span className="text-slate-400 block mb-0.5">当時の商談メモ・経緯</span>
                        <p className="text-slate-700 whitespace-pre-wrap bg-white rounded-lg p-2 border border-purple-100">{pastDeal.memo}</p>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
        <div className="flex items-center justify-between px-6 py-4 border-t border-slate-100 shrink-0">
          <div className="flex items-center gap-2">
            {onDelete && (
              <button onClick={onDelete}
                className="px-4 py-2 text-sm text-rose-500 hover:text-rose-700 border border-rose-200 hover:border-rose-400 hover:bg-rose-50 rounded-lg transition-colors">
                🗑️ 削除する
              </button>
            )}
            {onCopyToOrder && (
              <button onClick={onCopyToOrder}
                className="px-4 py-2 text-sm text-emerald-700 hover:text-emerald-800 border border-emerald-300 hover:bg-emerald-50 rounded-lg transition-colors">
                🛍️ 受注案件へコピー
              </button>
            )}
          </div>
          <div className="flex gap-2">
            <button onClick={onClose}
              className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">
              キャンセル
            </button>
            <button onClick={() => { if (!form.companyName.trim()) return; onSave(form); onClose(); }}
              disabled={!form.companyName.trim()}
              className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg disabled:opacity-40 transition-colors">
              保存する
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── レポート自動生成 ───────────────────────────────────────────────────────────
function generateReport(records, today, soon) {
  const total = records.length;
  if (total === 0) return null;

  const todayCalls   = records.filter(r => normDate(r.lastCallDate) === today).length;
  const apo          = records.filter(r => r.status === "9.アポ獲得").length;
  const apoRate      = (apo / total * 100).toFixed(1);
  const overdue      = records.filter(r => { const nd = normDate(r.nextCallDate); return nd && nd < today && !["8.不要","8.当社契約"].includes(r.status); });
  const upcoming     = records.filter(r => { const nd = normDate(r.nextCallDate); return nd && nd >= today && nd <= soon; });
  const mikaiden     = records.filter(r => !r.lastCallDate || r.lastCallDate === "").length;
  const statusMap    = {};
  records.forEach(r => { statusMap[r.status] = (statusMap[r.status]||0)+1; });

  // 店舗数別アポ率
  const STORE_RANGES = [
    ["1000店舗以上",1000,Infinity],["500〜999店舗",500,999],["300〜499店舗",300,499],
    ["100〜299店舗",100,299],["25〜99店舗",25,99],["10〜24店舗",10,24],["10店舗以下",0,9],
  ];
  const storeStats = STORE_RANGES.map(([label,lo,hi]) => {
    const inRange = records.filter(r => { const n=parseInt(String(r.storeCount||"").replace(/,/g,""))||0; return n>=lo&&n<=hi; });
    const a = inRange.filter(r => r.status==="9.アポ獲得").length;
    return { label, count:inRange.length, apo:a, rate: inRange.length ? (a/inRange.length*100) : 0 };
  }).filter(d=>d.count>0).sort((a,b)=>b.rate-a.rate);
  const bestSegment = storeStats[0];

  // 担当者別アポ数
  const assigneeApo = {};
  records.filter(r=>r.status==="9.アポ獲得").forEach(r=>{ const a=r.assignee||"未設定"; assigneeApo[a]=(assigneeApo[a]||0)+1; });
  const topAssignee = Object.entries(assigneeApo).sort((a,b)=>b[1]-a[1])[0];

  // コメント生成
  const insights = [];
  const actions  = [];
  const strategy = [];

  // 架電サマリー
  if (todayCalls > 0) insights.push(`本日 **${todayCalls}社** に架電済みです。`);
  else insights.push(`本日の架電はまだありません。企業名をクリックして架電を記録しましょう。`);

  // アポ率評価
  const apoRateNum = parseFloat(apoRate);
  if (apoRateNum >= 8) insights.push(`アポ獲得率 **${apoRate}%** は非常に高い水準です。現在のアプローチを継続してください。`);
  else if (apoRateNum >= 4) insights.push(`アポ獲得率 **${apoRate}%** は良好なペースです。`);
  else if (apoRateNum >= 1) insights.push(`アポ獲得率 **${apoRate}%** です。アプローチ方法の見直しで改善が期待できます。`);
  else if (total > 100) insights.push(`アポ獲得率 **${apoRate}%** と低めです。ターゲット選定とトーク改善を検討してください。`);

  // 期限超過
  if (overdue.length > 0) {
    actions.push(`🔴 **期限超過** ${overdue.length}社 — 本日優先的に架電してください。`);
  }
  if (upcoming.length > 0) {
    actions.push(`⚠️ **3日以内の架電予定** ${upcoming.length}社 — 事前準備をしておきましょう。`);
  }
  if (mikaiden > 0) {
    actions.push(`📞 **未架電** ${mikaiden.toLocaleString()}社 — 優先度の高い企業からアプローチしてください。`);
  }

  // 戦略提案
  if (bestSegment && bestSegment.rate > 0) {
    strategy.push(`🎯 **「${bestSegment.label}」が最高アポ率 ${bestSegment.rate.toFixed(1)}%** — このセグメントへの集中アプローチを推奨します（${bestSegment.count}社中${bestSegment.apo}社成功）。`);
  }
  if (topAssignee) {
    strategy.push(`🏆 **${topAssignee[0]}** がアポ獲得数トップ（${topAssignee[1]}件）。成功事例をチームで共有しましょう。`);
  }
  if (storeStats.length >= 2) {
    const worst = [...storeStats].sort((a,b)=>a.rate-b.rate).find(d=>d.count>=10);
    if (worst && worst.rate < 1) strategy.push(`📉 「${worst.label}」のアポ率が ${worst.rate.toFixed(1)}%と低い傾向があります。このセグメントのアプローチ方法の見直しを検討してください。`);
  }
  const connectCount = (statusMap["6.コネクト（改）"]||0) + (statusMap["7.コネクト（無）"]||0);
  if (connectCount > total * 0.2) {
    strategy.push(`💬 コネクト状態の企業が全体の ${(connectCount/total*100).toFixed(0)}%（${connectCount}社）あります。担当者へのコネクトをアポに転換する粘り強いフォローが重要です。`);
  }

  return { todayCalls, apo, apoRate, total, overdue: overdue.length, upcoming: upcoming.length, insights, actions, strategy, storeStats, topAssignee };
}

function ReportView({ records, pastDeals = [], orders = [] }) {
  const today = getToday();
  const soon  = (() => { const d=new Date(); d.setDate(d.getDate()+3); return d.toISOString().slice(0,10); })();
  const report = generateReport(records, today, soon);

  // AnalysisView のグラフ部分も残す
  const statusMap = {};
  records.forEach(r => { statusMap[r.status]=(statusMap[r.status]||0)+1; });
  const maxSt = Math.max(...Object.values(statusMap),1);
  const statusData = Object.keys(STATUS_CFG)
    .map(s => ({ status:s, count:statusMap[s]||0, cfg:STATUS_CFG[s] }))
    .filter(d=>d.count>0).sort((a,b)=>b.count-a.count);

  const Bar = ({count,max,colorClass}) => (
    <div className="flex-1 bg-slate-100 rounded-full h-4 overflow-hidden">
      <div className={`h-full rounded-full ${colorClass}`} style={{width:`${Math.max(2,count/max*100)}%`}}/>
    </div>
  );

  if (!report) return (
    <div className="bg-white rounded-xl border border-slate-200 py-16 text-center text-slate-400 text-sm">
      データがありません。CSVをインポートしてください。
    </div>
  );

  const Section = ({title, children}) => (
    <div className="bg-white rounded-xl border border-slate-200 p-5">
      <h3 className="text-sm font-bold text-slate-700 mb-3">{title}</h3>
      {children}
    </div>
  );

  const renderMd = (text) => {
    const parts = text.split(/\*\*(.*?)\*\*/g);
    return parts.map((p, i) => i % 2 === 1 ? <strong key={i}>{p}</strong> : p);
  };

  // ── 分析①：IS担当者別マトリクス ──────────────────────────────────────────────
  const [memberSortKey, setMemberSortKey] = useState("apoRate");
  const recOwner = (r) => `${r.assignee||""} ${r.createdBy||""}`; // 担当者+追加者で照合
  const memberStats = IS_MEMBERS.map(m => {
    const mn = normName(m);
    const rs = records.filter(r => {
      const on = normName(recOwner(r));
      return on.includes(mn) || mn.includes(normName(r.assignee||"")) && r.assignee;
    });
    const cnt = (...st) => rs.filter(r => st.includes(r.status)).length;
    const apo     = cnt("9.アポ獲得","0.日程調整");
    const connect = cnt("6.コネクト（改）","7.コネクト（無）");
    const refused = cnt("4.受付カット");
    const absent  = rs.filter(r => ["不在","不通","2.優先","3.並"].includes(r.status)).length;
    const apoRate = connect > 0 ? (apo/connect*100) : 0;
    return { name:m, total:rs.length, apo, connect, refused, absent, apoRate };
  }).filter(s => s.total > 0)
    .sort((a,b) => memberSortKey==="apoRate" ? b.apoRate-a.apoRate
                 : memberSortKey==="apo"     ? b.apo-a.apo
                 : b.total-a.total);

  // ── 分析③：業種別 成果ランキング（トップ5） ─────────────────────────────────
  const indMap = {};
  records.forEach(r => {
    if (!r.industry) return;
    const k = r.industry;
    indMap[k] = indMap[k] || { industry:k, total:0, apo:0, connect:0 };
    indMap[k].total++;
    if (["9.アポ獲得","0.日程調整"].includes(r.status)) indMap[k].apo++;
    if (["6.コネクト（改）","7.コネクト（無）"].includes(r.status)) indMap[k].connect++;
  });
  const industryRank = Object.values(indMap)
    .sort((a,b) => (b.apo+b.connect) - (a.apo+a.connect)).slice(0,5);

  // ── 分析③：店舗規模別セグメント ─────────────────────────────────────────────
  const storeIdx = useMemo(() => buildStoreIndex(records, pastDeals), [records, pastDeals]);
  const sizeSegs = [
    { label:"1〜9店舗",   lo:1,   hi:9 },
    { label:"10〜99店舗", lo:10,  hi:99 },
    { label:"100店舗以上", lo:100, hi:Infinity },
  ].map(seg => {
    const rs = records.filter(r => {
      const n = analyzeStoreCount(r, storeIdx).value || 0;
      return n >= seg.lo && n <= seg.hi;
    });
    const apo = rs.filter(r => ["9.アポ獲得","0.日程調整"].includes(r.status)).length;
    return { ...seg, total:rs.length, apo, rate: rs.length ? (apo/rs.length*100) : 0 };
  });

  // ── 受注集計 ─────────────────────────────────────────────────────────────────
  const orderAmount = (o) => { const n = parseInt(String(o.amount??"").replace(/[,，円\s]/g,""),10); return Number.isFinite(n)?n:0; };
  const thisMonth = today.slice(0,7); // YYYY-MM
  const monthOrders = orders.filter(o => normDate(o.orderDate).slice(0,7) === thisMonth);
  const monthTotal  = monthOrders.reduce((s,o)=>s+orderAmount(o),0);
  const orderTotalAll = orders.reduce((s,o)=>s+orderAmount(o),0);
  // IS担当者別 受注貢献
  const orderByMember = IS_MEMBERS.map(m => {
    const mn = normName(m);
    const rs = orders.filter(o => { const an=normName(o.assignee||""); return an===mn || an.includes(mn) || mn.includes(an); });
    return { name:m, count:rs.length, amount:rs.reduce((s,o)=>s+orderAmount(o),0) };
  }).filter(s => s.count>0).sort((a,b)=>b.amount-a.amount);

  return (
    <div className="space-y-4">
      {/* KPI */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        {[
          { label:"総件数",     value:report.total.toLocaleString(),  color:"text-blue-600",   bg:"bg-blue-50"   },
          { label:"本日架電数",  value:report.todayCalls,              color:"text-amber-600",  bg:"bg-amber-50"  },
          { label:"アポ数",     value:report.apo,                     color:"text-teal-600",   bg:"bg-teal-50"   },
          { label:"アポ率",     value:`${report.apoRate}%`,           color:"text-green-600",  bg:"bg-green-50"  },
        ].map(k=>(
          <div key={k.label} className={`${k.bg} rounded-xl border border-slate-200 p-4 text-center`}>
            <div className={`text-3xl font-bold ${k.color}`}>{k.value}</div>
            <div className="text-xs text-slate-500 mt-1">{k.label}</div>
          </div>
        ))}
      </div>

      {/* 受注集計 */}
      <Section title="🛍️ 受注集計">
        <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
          {[
            { label:`当月受注総額（${thisMonth.replace("-","/")}）`, value:`¥${monthTotal.toLocaleString()}`, color:"text-emerald-700", bg:"bg-emerald-50" },
            { label:"当月受注件数", value:`${monthOrders.length}件`, color:"text-emerald-700", bg:"bg-emerald-50" },
            { label:"累計受注総額", value:`¥${orderTotalAll.toLocaleString()}`, color:"text-teal-700",  bg:"bg-teal-50" },
          ].map(k=>(
            <div key={k.label} className={`${k.bg} rounded-xl border border-slate-200 p-4 text-center`}>
              <div className={`text-2xl font-bold ${k.color}`}>{k.value}</div>
              <div className="text-xs text-slate-500 mt-1">{k.label}</div>
            </div>
          ))}
        </div>
        <p className="text-xs font-semibold text-slate-500 mb-2">IS担当者別 受注貢献ランキング</p>
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead><tr className="border-b-2 border-slate-200 text-slate-500 text-left">
              <th className="pb-2 font-semibold">担当者</th>
              <th className="pb-2 font-semibold text-right">受注件数</th>
              <th className="pb-2 font-semibold text-right">合計受注金額</th>
            </tr></thead>
            <tbody>
              {orderByMember.length===0 ? (
                <tr><td colSpan={3} className="py-6 text-center text-slate-400">受注データがありません</td></tr>
              ) : orderByMember.map((s,i)=>(
                <tr key={s.name} className="border-b border-slate-50 hover:bg-slate-50">
                  <td className="py-2 text-slate-700 font-medium whitespace-nowrap">{i+1}. {s.name}</td>
                  <td className="py-2 text-right text-slate-600">{s.count}件</td>
                  <td className="py-2 text-right font-bold text-emerald-700">¥{s.amount.toLocaleString()}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Section>

      {/* 今日のアクション */}
      {report.actions.length > 0 && (
        <Section title="📌 今日のアクション">
          <ul className="space-y-2">
            {report.actions.map((a,i) => (
              <li key={i} className="text-sm text-slate-700 bg-slate-50 rounded-lg px-4 py-2.5">{renderMd(a)}</li>
            ))}
          </ul>
        </Section>
      )}

      {/* インサイト */}
      <Section title="💡 データインサイト">
        <ul className="space-y-2">
          {report.insights.map((ins,i) => (
            <li key={i} className="text-sm text-slate-700 leading-relaxed">{renderMd(ins)}</li>
          ))}
        </ul>
      </Section>

      {/* 推奨戦略 */}
      {report.strategy.length > 0 && (
        <Section title="🎯 推奨戦略">
          <ul className="space-y-3">
            {report.strategy.map((s,i) => (
              <li key={i} className="text-sm text-slate-700 bg-indigo-50 border border-indigo-100 rounded-lg px-4 py-3 leading-relaxed">{renderMd(s)}</li>
            ))}
          </ul>
        </Section>
      )}

      {/* 分析①：IS担当者別マトリクス */}
      <Section title="👥 IS担当者別 活動・成果マトリクス">
        <div className="flex gap-1.5 mb-3">
          {[["apoRate","アポ獲得率順"],["apo","アポ数順"],["total","件数順"]].map(([k,l])=>(
            <button key={k} onClick={()=>setMemberSortKey(k)}
              className={`px-2.5 py-1 text-xs rounded-lg border transition-colors ${memberSortKey===k ? "bg-blue-600 text-white border-blue-600" : "border-slate-200 text-slate-500 hover:bg-slate-50"}`}>
              {l}
            </button>
          ))}
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr className="border-b-2 border-slate-200 text-slate-500 text-left">
                <th className="pb-2 font-semibold">担当者</th>
                <th className="pb-2 font-semibold text-right">総件数</th>
                <th className="pb-2 font-semibold text-right">アポ</th>
                <th className="pb-2 font-semibold text-right">コネクト</th>
                <th className="pb-2 font-semibold text-right">受付断り</th>
                <th className="pb-2 font-semibold text-right">不在/不通</th>
                <th className="pb-2 font-semibold text-right">アポ獲得率</th>
              </tr>
            </thead>
            <tbody>
              {memberStats.length === 0 ? (
                <tr><td colSpan={7} className="py-6 text-center text-slate-400">担当者データがありません</td></tr>
              ) : memberStats.map(s => (
                <tr key={s.name} className="border-b border-slate-50 hover:bg-slate-50">
                  <td className="py-2 text-slate-700 font-medium whitespace-nowrap">{s.name}</td>
                  <td className="py-2 text-right text-slate-600">{s.total.toLocaleString()}</td>
                  <td className="py-2 text-right font-bold text-red-700">{s.apo}</td>
                  <td className="py-2 text-right text-blue-700">{s.connect}</td>
                  <td className="py-2 text-right text-amber-700">{s.refused}</td>
                  <td className="py-2 text-right text-slate-500">{s.absent}</td>
                  <td className={`py-2 text-right font-bold ${s.apoRate>=20?"text-teal-600":s.apoRate>=10?"text-blue-600":"text-slate-500"}`}>
                    {s.apoRate.toFixed(1)}%
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p className="text-[10px] text-slate-400 mt-2">※ アポ獲得率 = アポ獲得商談数 ÷ 担当コネクト数（分母0は0%）</p>
      </Section>

      {/* 分析②：ステータス別パイプライン */}
      <Section title="📊 ステータス別パイプライン（ファネル）">
        <div className="space-y-2">
          {statusData.map(d => (
            <div key={d.status} className="flex items-center gap-2">
              <span className="text-xs text-slate-600 w-32 shrink-0 truncate">{d.status}</span>
              <div className="flex-1 bg-slate-200 rounded-full h-4 overflow-hidden">
                <div className={`h-full rounded-full ${d.cfg?.dot??"bg-slate-400"}`} style={{width:`${Math.max(2,d.count/maxSt*100)}%`}} />
              </div>
              <span className="text-xs font-bold text-slate-700 w-14 text-right shrink-0">{d.count.toLocaleString()}</span>
              <span className="text-xs text-slate-400 w-9 text-right shrink-0">{Math.round(d.count/records.length*100)}%</span>
            </div>
          ))}
        </div>
      </Section>

      {/* 分析③：業種別 成果ランキング */}
      {industryRank.length > 0 && (
        <Section title="🏭 業種別 成果ランキング（トップ5）">
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="border-b-2 border-slate-200 text-slate-500 text-left">
                  <th className="pb-2 font-semibold">業種</th>
                  <th className="pb-2 font-semibold text-right">件数</th>
                  <th className="pb-2 font-semibold text-right">アポ</th>
                  <th className="pb-2 font-semibold text-right">コネクト</th>
                </tr>
              </thead>
              <tbody>
                {industryRank.map((d,i) => (
                  <tr key={d.industry} className="border-b border-slate-50 hover:bg-slate-50">
                    <td className="py-2 text-slate-700 font-medium">{i+1}. {d.industry}</td>
                    <td className="py-2 text-right text-slate-600">{d.total.toLocaleString()}</td>
                    <td className="py-2 text-right font-bold text-red-700">{d.apo}</td>
                    <td className="py-2 text-right text-blue-700">{d.connect}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Section>
      )}

      {/* 分析③：店舗規模別セグメント */}
      <Section title="🏬 店舗規模別セグメント（仮含む）">
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr className="border-b-2 border-slate-200 text-slate-500 text-left">
                <th className="pb-2 font-semibold">規模</th>
                <th className="pb-2 font-semibold text-right">アプローチ</th>
                <th className="pb-2 font-semibold text-right">アポ獲得</th>
                <th className="pb-2 font-semibold text-right">率</th>
              </tr>
            </thead>
            <tbody>
              {sizeSegs.map(d => (
                <tr key={d.label} className="border-b border-slate-50 hover:bg-slate-50">
                  <td className="py-2 text-slate-700 font-medium">{d.label}</td>
                  <td className="py-2 text-right text-slate-600">{d.total.toLocaleString()}</td>
                  <td className="py-2 text-right font-bold text-red-700">{d.apo}</td>
                  <td className={`py-2 text-right font-semibold ${d.rate>=8?"text-teal-600":d.rate>=4?"text-blue-600":"text-slate-500"}`}>
                    {d.rate.toFixed(1)}%
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Section>

      {/* 店舗数別アポ率 */}
      {report.storeStats.length > 0 && (
        <Section title="🏪 店舗数別アポ率">
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="border-b-2 border-slate-200 text-slate-500 text-left">
                  <th className="pb-2 font-semibold">店舗数</th>
                  <th className="pb-2 font-semibold text-right">社数</th>
                  <th className="pb-2 font-semibold text-right">アポ</th>
                  <th className="pb-2 font-semibold text-right">率</th>
                </tr>
              </thead>
              <tbody>
                {report.storeStats.map(d => (
                  <tr key={d.label} className="border-b border-slate-50 hover:bg-slate-50">
                    <td className="py-2 text-slate-700 font-medium">{d.label}</td>
                    <td className="py-2 text-right text-slate-600">{d.count.toLocaleString()}</td>
                    <td className="py-2 text-right font-bold text-blue-700">{d.apo}</td>
                    <td className={`py-2 text-right font-semibold ${d.rate>=8?"text-teal-600":d.rate>=4?"text-blue-600":"text-slate-500"}`}>
                      {d.rate.toFixed(1)}%
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Section>
      )}
    </div>
  );
}

// ── AnalysisView ───────────────────────────────────────────────────────────────
function AnalysisView({ records }) {
  const today = getToday();
  const soon  = (() => { const d=new Date(); d.setDate(d.getDate()+3); return d.toISOString().slice(0,10); })();

  const kpis = [
    { label:"総件数",    value: records.length,                                                         color:"text-blue-600",  bg:"bg-blue-50"  },
    { label:"本日架電数", value: records.filter(r=>normDate(r.lastCallDate)===today).length,              color:"text-amber-600", bg:"bg-amber-50" },
    { label:"アポ数",    value: records.filter(r=>r.status==="9.アポ獲得").length,                        color:"text-teal-600",  bg:"bg-teal-50"  },
  ];

  const statusMap = {};
  records.forEach(r => { statusMap[r.status]=(statusMap[r.status]||0)+1; });
  const maxSt = Math.max(...Object.values(statusMap),1);
  const statusData = Object.keys(STATUS_CFG)
    .map(s => ({ status:s, count:statusMap[s]||0, cfg:STATUS_CFG[s] }))
    .filter(d=>d.count>0).sort((a,b)=>b.count-a.count);

  // 担当者 あり/なし
  const assigneeAri   = records.filter(r => r.assignee?.trim()).length;
  const assigneeNashi = records.length - assigneeAri;
  const assigneeAriPct = records.length ? Math.round(assigneeAri/records.length*100) : 0;

  const indMap={};
  records.forEach(r=>{ if(r.industry) indMap[r.industry]=(indMap[r.industry]||0)+1; });
  const indData=Object.entries(indMap).sort((a,b)=>b[1]-a[1]).slice(0,10);

  // 店舗数別分析
  const STORE_RANGES = [
    ["1000店舗以上", 1000, Infinity],
    ["500〜999店舗",  500,  999],
    ["300〜499店舗",  300,  499],
    ["100〜299店舗",  100,  299],
    ["80〜99店舗",     80,   99],
    ["50〜79店舗",     50,   79],
    ["25〜49店舗",     25,   49],
    ["11〜24店舗",     11,   24],
    ["10店舗以下",      0,   10],
  ];
  const APPO_STATUSES = ["9.アポ獲得"];
  const storeAnalysis = STORE_RANGES.map(([label,lo,hi]) => {
    const inRange = records.filter(r => { const n=parseInt(String(r.storeCount||"").replace(/,/g,""))||0; return n>=lo && n<=hi; });
    const appo    = inRange.filter(r => APPO_STATUSES.includes(r.status)).length;
    const rate    = inRange.length ? (appo/inRange.length*100).toFixed(1) : null;
    return { label, count:inRange.length, appo, rate };
  });
  const maxStoreCount = Math.max(...storeAnalysis.map(d=>d.count), 1);

  const Card=({label,value,color,bg})=>(
    <div className={`${bg} rounded-xl border border-slate-200 p-4 text-center`}>
      <div className={`text-3xl font-bold ${color}`}>{value.toLocaleString()}</div>
      <div className="text-xs text-slate-500 mt-1">{label}</div>
    </div>
  );
  const Bar=({count,max,colorClass})=>(
    <div className="flex-1 bg-slate-100 rounded-full h-4 overflow-hidden">
      <div className={`h-full rounded-full transition-all ${colorClass}`}
        style={{width:`${Math.max(2,count/max*100)}%`}} />
    </div>
  );

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        {kpis.map(k=><Card key={k.label} {...k}/>)}
      </div>

      {/* ステータス別 */}
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <h3 className="text-sm font-semibold text-slate-700 mb-4">ステータス別件数</h3>
        <div className="space-y-2">
          {statusData.map(d=>(
            <div key={d.status} className="flex items-center gap-2">
              <span className="text-xs text-slate-600 w-32 shrink-0 truncate">{d.status}</span>
              <Bar count={d.count} max={maxSt} colorClass={d.cfg?.dot??"bg-slate-400"} />
              <span className="text-xs font-bold text-slate-700 w-14 text-right shrink-0">{d.count.toLocaleString()}</span>
              <span className="text-xs text-slate-400 w-9 text-right shrink-0">{Math.round(d.count/records.length*100)}%</span>
            </div>
          ))}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {/* 担当者 あり/なし */}
        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <h3 className="text-sm font-semibold text-slate-700 mb-4">担当者 割合</h3>
          <div className="flex h-6 rounded-full overflow-hidden mb-3">
            <div className="bg-blue-500 transition-all" style={{width:`${assigneeAriPct}%`}} />
            <div className="bg-slate-200 flex-1" />
          </div>
          <div className="flex gap-4 text-xs">
            <span className="flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-full bg-blue-500 shrink-0" />
              <span className="text-slate-700">担当者あり</span>
              <strong className="text-blue-700 ml-1">{assigneeAriPct}%</strong>
              <span className="text-slate-400">（{assigneeAri.toLocaleString()}件）</span>
            </span>
            <span className="flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-full bg-slate-300 shrink-0" />
              <span className="text-slate-700">担当者なし</span>
              <strong className="text-slate-500 ml-1">{100-assigneeAriPct}%</strong>
              <span className="text-slate-400">（{assigneeNashi.toLocaleString()}件）</span>
            </span>
          </div>
        </div>

        {/* 業種別 */}
        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <h3 className="text-sm font-semibold text-slate-700 mb-3">業種別件数（上位10）</h3>
          {indData.length===0 ? <p className="text-xs text-slate-400">業種データなし</p> : (
            <div className="space-y-2">
              {indData.map(([name,count])=>(
                <div key={name} className="flex items-center gap-2">
                  <span className="text-xs text-slate-600 flex-1 truncate">{name}</span>
                  <Bar count={count} max={indData[0][1]} colorClass="bg-blue-400" />
                  <span className="text-xs font-semibold text-slate-700 w-8 text-right shrink-0">{count}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* 店舗数別分析 */}
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <h3 className="text-sm font-semibold text-slate-700 mb-3">店舗数別分析</h3>
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr className="border-b-2 border-slate-200 text-slate-500 text-left">
                <th className="pb-2 font-semibold">店舗数</th>
                <th className="pb-2 font-semibold text-right">社数</th>
                <th className="pb-2 font-semibold text-right">アポ</th>
                <th className="pb-2 font-semibold text-right">率</th>
                <th className="pb-2 font-semibold pl-4 w-32">社数バー</th>
              </tr>
            </thead>
            <tbody>
              {storeAnalysis.map(d => (
                <tr key={d.label} className="border-b border-slate-50 hover:bg-slate-50">
                  <td className="py-2 text-slate-700 font-medium">{d.label}</td>
                  <td className="py-2 text-right text-slate-600">{d.count.toLocaleString()}</td>
                  <td className="py-2 text-right font-bold text-blue-700">{d.appo}</td>
                  <td className={`py-2 text-right font-semibold ${
                    d.rate === null ? "text-slate-300"
                    : parseFloat(d.rate) >= 10 ? "text-teal-600"
                    : parseFloat(d.rate) >= 5  ? "text-blue-600"
                    : "text-slate-500"
                  }`}>
                    {d.rate === null ? "—" : `${d.rate}%`}
                  </td>
                  <td className="py-2 pl-4">
                    <div className="w-full bg-slate-100 rounded-full h-3 overflow-hidden">
                      <div className="h-full bg-emerald-400 rounded-full"
                        style={{width:`${Math.max(2, d.count/maxStoreCount*100)}%`}} />
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="border-t-2 border-slate-200 font-semibold">
                <td className="pt-2 text-slate-700">合計</td>
                <td className="pt-2 text-right text-slate-700">
                  {storeAnalysis.reduce((s,d)=>s+d.count,0).toLocaleString()}
                </td>
                <td className="pt-2 text-right text-blue-700">
                  {storeAnalysis.reduce((s,d)=>s+d.appo,0)}
                </td>
                <td className="pt-2 text-right text-slate-500">
                  {(() => {
                    const tot = storeAnalysis.reduce((s,d)=>s+d.count,0);
                    const apo = storeAnalysis.reduce((s,d)=>s+d.appo,0);
                    return tot ? `${(apo/tot*100).toFixed(1)}%` : "—";
                  })()}
                </td>
                <td />
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── PullView ────────────────────────────────────────────────────────────────────
function PullView({ records }) {
  const [inputText, setInputText] = useState("");
  const [results,   setResults]   = useState(null);
  const today = getToday();
  const soon  = (() => { const d=new Date(); d.setDate(d.getDate()+3); return d.toISOString().slice(0,10); })();

  const getAction = rec => {
    if (!rec) return { label:"❓ リスト未登録",   color:"text-slate-400",  bg:"" };
    if (["8.不要","8.当社契約"].includes(rec.status))
                   return { label:"✅ 対応不要",     color:"text-slate-400",  bg:"bg-slate-50" };
    if (["9.アポ獲得","0.日程調整","1.高確度"].includes(rec.status))
                   return { label:"🎯 商談フォロー", color:"text-teal-700",   bg:"bg-teal-50" };
    const nd = normDate(rec.nextCallDate);
    if (nd && nd < today)  return { label:"🔴 至急架電",    color:"text-red-700",   bg:"bg-red-50" };
    if (nd && nd <= soon)  return { label:"⚠️ 近日架電",    color:"text-amber-700", bg:"bg-amber-50" };
    if (nd)                return { label:"📅 架電予定あり", color:"text-blue-700",  bg:"bg-blue-50" };
    return                        { label:"📞 架電日未設定", color:"text-orange-600",bg:"bg-orange-50" };
  };

  const doMatch = () => {
    const names = inputText.split(/[\r\n]+/)
      .flatMap(line => line.split(/[\t,]/).map(s=>s.trim()))
      .filter(Boolean);
    setResults(names.map(name => {
      const norm = normName(name);
      const matched = records.filter(r => {
        const rn = normName(r.companyName);
        return rn === norm || rn.includes(norm) || norm.includes(rn);
      });
      return { name, matched };
    }));
  };

  const matched   = results ? results.filter(r=>r.matched.length>0).length : 0;
  const unmatched = results ? results.filter(r=>r.matched.length===0).length : 0;

  const ACTION_ORDER = ["🔴 至急架電","⚠️ 近日架電","📞 架電日未設定","📅 架電予定あり","🎯 商談フォロー","✅ 対応不要","❓ リスト未登録"];

  return (
    <div className="space-y-4">
      {/* 入力パネル */}
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <h3 className="text-sm font-semibold text-slate-700 mb-1">過去商談プルリスト 照合</h3>
        <p className="text-xs text-slate-400 mb-3">企業名を1行ずつ貼り付けると、現在のステータス・アクションを自動判定します。</p>
        <textarea
          value={inputText}
          onChange={e=>{ setInputText(e.target.value); setResults(null); }}
          placeholder={"株式会社〇〇\n株式会社△△\n（Excelからのコピペも可）"}
          rows={6}
          className="w-full border border-slate-300 rounded-xl px-3 py-2.5 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none font-mono placeholder:text-slate-300"
        />
        <div className="flex justify-between items-center mt-2">
          <span className="text-xs text-slate-400">
            {inputText.trim() ? `${inputText.trim().split(/[\r\n]+/).filter(Boolean).length} 行入力中` : ""}
          </span>
          <button onClick={doMatch} disabled={!inputText.trim()}
            className="px-5 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold rounded-lg disabled:opacity-40 transition-colors">
            照合する
          </button>
        </div>
      </div>

      {/* 結果サマリー */}
      {results && (
        <div className="grid grid-cols-3 gap-3 text-center">
          {[
            { label:"照合対象",  value:results.length,  color:"text-slate-700",  bg:"bg-white" },
            { label:"マッチ",    value:matched,          color:"text-blue-700",   bg:"bg-blue-50" },
            { label:"未登録",    value:unmatched,        color:"text-slate-400",  bg:"bg-slate-50" },
          ].map(k=>(
            <div key={k.label} className={`${k.bg} rounded-xl border border-slate-200 py-3`}>
              <div className={`text-2xl font-bold ${k.color}`}>{k.value}</div>
              <div className="text-xs text-slate-500 mt-0.5">{k.label}</div>
            </div>
          ))}
        </div>
      )}

      {/* 結果テーブル */}
      {results && (
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead className="bg-slate-50 border-b border-slate-200">
                <tr>
                  {["企業名（入力）","マッチ企業名","状況","架電日","次回架電日","担当者","メモ","アクション"].map(h=>(
                    <th key={h} className="px-3 py-2.5 text-left font-semibold text-slate-500 whitespace-nowrap">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {results
                  .slice()
                  .sort((a,b)=>{
                    const la = a.matched.length ? ACTION_ORDER.indexOf(getAction(a.matched[0]).label) : ACTION_ORDER.length-1;
                    const lb = b.matched.length ? ACTION_ORDER.indexOf(getAction(b.matched[0]).label) : ACTION_ORDER.length-1;
                    return la - lb;
                  })
                  .flatMap((res, i) => {
                    if (res.matched.length === 0) {
                      const act = getAction(null);
                      return [(
                        <tr key={`u${i}`} className="hover:bg-slate-50/60">
                          <td className="px-3 py-2.5 text-slate-600 font-medium">{res.name}</td>
                          <td className="px-3 py-2.5 text-slate-300 italic">未登録</td>
                          <td colSpan={5} className="px-3 py-2.5 text-slate-300">—</td>
                          <td className={`px-3 py-2.5 font-semibold ${act.color}`}>{act.label}</td>
                        </tr>
                      )];
                    }
                    return res.matched.map((rec, j) => {
                      const act = getAction(rec);
                      return (
                        <tr key={`m${i}-${j}`} className={`hover:brightness-95 ${act.bg}`}>
                          <td className="px-3 py-2.5 text-slate-500">{j===0 ? res.name : ""}</td>
                          <td className="px-3 py-2.5 font-semibold text-slate-800">{rec.companyName}</td>
                          <td className="px-3 py-2.5"><StatusBadge status={rec.status}/></td>
                          <td className="px-3 py-2.5 text-slate-600">{fmtDate(normDate(rec.lastCallDate))||"—"}</td>
                          <td className="px-3 py-2.5 text-slate-600">{fmtDate(normDate(rec.nextCallDate))||"—"}</td>
                          <td className="px-3 py-2.5 text-slate-600">{rec.assignee||"—"}</td>
                          <td className="px-3 py-2.5 text-slate-500 max-w-48 truncate" title={rec.memo||""}>{rec.memo||"—"}</td>
                          <td className={`px-3 py-2.5 font-bold whitespace-nowrap ${act.color}`}>{act.label}</td>
                        </tr>
                      );
                    });
                  })
                }
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// ── 過去商談専用追加列（ALL_COLUMNSに含まれない） ──────────────────────────────
const PAST_EXTRA_COLS = [
  { key:"progress",   label:"進捗",       required:false, w:"w-[150px]" },
  { key:"targetDate", label:"完了予定日", required:false, w:"w-[110px]" },
];

// ── 受注案件管理 ───────────────────────────────────────────────────────────────
const ORDER_PLANS = ["ライトプラン","スタンダードプラン","プレミアムプラン","その他"];
const PAYMENT_STATUS_CFG = {
  "未入金":   { bg:"bg-rose-100",   text:"text-rose-700",   dot:"bg-rose-500"   },
  "一部入金": { bg:"bg-amber-100",  text:"text-amber-700",  dot:"bg-amber-500"  },
  "入金済":   { bg:"bg-green-100",  text:"text-green-700",  dot:"bg-green-500"  },
  "要確認":   { bg:"bg-slate-200",  text:"text-slate-600",  dot:"bg-slate-500"  },
};
const ORDER_COLS = [
  { key:"orderDate",    label:"受注日",          required:false, w:"w-[110px]" },
  { key:"companyName",  label:"企業名",          required:true,  w:"w-[220px]" },
  { key:"plan",         label:"プラン/商材",     required:false, w:"w-[150px]" },
  { key:"amount",       label:"受注金額",        required:false, w:"w-[120px]" },
  { key:"assignee",     label:"担当者",          required:false, w:"w-[110px]" },
  { key:"payment",      label:"入金ステータス",  required:false, w:"w-[130px]" },
  { key:"startDate",    label:"稼働開始日",      required:false, w:"w-[110px]" },
  { key:"memo",         label:"受注メモ/案件詳細", required:false, w:"w-[240px]" },
];
const ORDER_DEFAULT_VISIBLE = ORDER_COLS.map(c => c.key);

// 受注CSVヘッダーマッピング
function mapOrderHeaders(headers) {
  const m = {};
  headers.forEach((h, i) => {
    const n = String(h).replace(/[\s　]/g,"");
    if (/企業名|会社名|取引先名?|法人名/.test(n))      m.companyName = i;
    else if (/受注日|契約日|成約日/.test(n))            m.orderDate   = i;
    else if (/プラン|商材|商品|サービス/.test(n))        m.plan        = i;
    else if (/金額|受注額|契約金額|売上/.test(n))        m.amount      = i;
    else if (/担当者?|営業担当/.test(n))                m.assignee    = i;
    else if (/入金|支払/.test(n))                       m.payment     = i;
    else if (/稼働開始|開始日/.test(n))                 m.startDate   = i;
    else if (/メモ|備考|詳細|コメント/.test(n))          m.memo        = i;
  });
  return m;
}
function parseAmount(v) {
  const n = parseInt(String(v ?? "").replace(/[,，円\s]/g,""), 10);
  return Number.isFinite(n) ? n : 0;
}
const ALL_PAST_COLS = [...ALL_COLUMNS, ...PAST_EXTRA_COLS];
const DEFAULT_PAST_VISIBLE = [
  "companyName","progress","lastCallDate","nextCallDate","status","storeCount","phone",
  "createdBy","assignee","leadSource","memo","targetDate",
];

// ── エンタープライズ管理：列はメインリストと同一＋大手専用列 ───────────────────────
const ENT_EXTRA_COLS = [
  { key:"internalOwner", label:"社内担当者",         w:"w-[100px]" },
  { key:"corpName",      label:"法人名",             w:"w-[200px]" },
  { key:"dealName",      label:"商談名",             w:"w-[140px]" },
  { key:"gmoPhase",      label:"GMO営業フェーズ",    w:"w-[200px]" },
  { key:"tryhatchStatus",label:"トライハッチ営業状況", w:"w-[160px]" },
];
const ENT_COLS = [...ALL_COLUMNS, ...ENT_EXTRA_COLS];
const ENT_DEFAULT_VISIBLE = [
  "companyName","lastCallDate","nextCallDate","status","internalOwner",
  "industry","storeCount","phone","gmoPhase","memo",
];

// 大手「状況」→ TEPPOU状況へのマッピング
function entStatusMap(s) {
  const v = String(s||"").trim();
  if (/契約|受注/.test(v))     return "8.当社契約";
  if (/商談/.test(v))          return "4.商談中";
  if (/アプローチ/.test(v))    return "2.優先";
  if (/失注|断り|不要/.test(v)) return "8.不要";
  return v && STATUS_CFG[v] ? v : "未架電";
}

// 大手シートのヘッダー→キー マッピング
function mapEntHeaders(headers) {
  const m = {};
  headers.forEach((h, i) => {
    const n = String(h).replace(/[\s　]/g,"");
    if (/企業・ブランド名|企業名|ブランド名/.test(n))      m.companyName  = i;
    else if (/^状況$/.test(n))                            m.status       = i;
    else if (/社内担当者/.test(n))                        m.internalOwner= i;
    else if (/^状態$/.test(n))                            m.state        = i;
    else if (/法人名/.test(n))                            m.corpName     = i;
    else if (/業種/.test(n))                              m.industry     = i;
    else if (/店舗数/.test(n))                            m.storeCount   = i;
    else if (/氏名/.test(n))                              m.personName   = i;
    else if (/商談名/.test(n))                            m.dealName     = i;
    else if (/GMO営業フェーズ|営業フェーズ|フェーズ/.test(n)) m.gmoPhase  = i;
    else if (/最終更新日|更新日/.test(n))                 m.lastUpdate   = i;
    else if (/トライハッチ/.test(n))                      m.tryhatchStatus = i;
    else if (/担当者/.test(n) && m.contactName===undefined) m.contactName = i;
  });
  return m;
}

// ── EnterpriseView ─────────────────────────────────────────────────────────────
function EnterpriseView({ enterprise, setEnterprise, records = [] }) {
  const savedUI = (() => { try { return JSON.parse(localStorage.getItem("teppou_ent_ui_v1") || "{}"); } catch { return {}; } })();
  const [search,      setSearch]      = useState("");
  const [stFilter,    setStFilter]    = useState("all");
  const [editCell,    setEditCell]    = useState(null);
  const [visibleCols, setVisibleCols] = useState(Array.isArray(savedUI.visibleCols)&&savedUI.visibleCols.length ? savedUI.visibleCols : ENT_DEFAULT_VISIBLE);
  const [showColDrop, setShowColDrop] = useState(false);
  const [sortKey,     setSortKey]     = useState(savedUI.sortKey ?? null);
  const [sortDir,     setSortDir]     = useState(savedUI.sortDir || "asc");
  const [log,         setLog]         = useState(null);
  const [loading,     setLoading]     = useState(false);
  const [page,        setPage]        = useState(1);
  const colDropRef = useRef();
  const fileRef    = useRef();
  const PAGE = 100;

  useEffect(() => { try { localStorage.setItem("teppou_ent_ui_v1", JSON.stringify({ visibleCols, sortKey, sortDir })); } catch {} }, [visibleCols, sortKey, sortDir]);
  useEffect(() => {
    if (!showColDrop) return;
    const fn = e => { if (colDropRef.current && !colDropRef.current.contains(e.target)) setShowColDrop(false); };
    document.addEventListener("mousedown", fn);
    return () => document.removeEventListener("mousedown", fn);
  }, [showColDrop]);

  const visibleDefs = visibleCols.map(k => ENT_COLS.find(c => c.key===k)).filter(Boolean);
  const today = getToday();
  const storeIndex = useMemo(() => buildStoreIndex(enterprise, []), [enterprise]);
  const [listModal, setListModal] = useState(null); // { name, matched[] }
  const currentNames = useMemo(() => new Set(records.map(r => normName(r.companyName))), [records]);
  const isInCurrent = (name) => { const n = normName(name); return [...currentNames].some(cn => cn===n || cn.includes(n) || n.includes(cn)); };

  // ステータス集計
  const stCounts = useMemo(() => {
    const m = {};
    enterprise.forEach(r => { const s = r.status||"未設定"; m[s]=(m[s]||0)+1; });
    return m;
  }, [enterprise]);

  const filtered = useMemo(() => {
    let rs = enterprise.filter(r => {
      if (stFilter !== "all" && (r.status||"未設定") !== stFilter) return false;
      if (search) { const q=search.toLowerCase(); return ENT_COLS.some(c => String(r[c.key]||"").toLowerCase().includes(q)); }
      return true;
    });
    if (sortKey) rs = [...rs].sort((a,b) => {
      const va=String(a[sortKey]||""), vb=String(b[sortKey]||"");
      const cmp = va.localeCompare(vb,"ja"); return sortDir==="asc"?cmp:-cmp;
    });
    return rs;
  }, [enterprise, search, stFilter, sortKey, sortDir]);

  const totalPages = Math.max(1, Math.ceil(filtered.length/PAGE));
  const paginated  = filtered.slice((page-1)*PAGE, page*PAGE);

  // インポート
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true); setLog(null);
    const reader = new FileReader();
    reader.onload = ev => {
      setTimeout(() => {
        try {
          const wb = XLSX.read(ev.target.result, { type:"array", cellDates:true });
          // 「大手」シート優先、なければ先頭
          const sheetName = wb.SheetNames.find(n => /大手/.test(n)) || wb.SheetNames[0];
          const ws = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:"" })
            .map(r => r.map(c => c instanceof Date ? c.toISOString().slice(0,10) : String(c??"").trim()));
          const map = mapEntHeaders(rows[0]||[]);
          if (map.companyName === undefined) { setLog({ error:`「企業・ブランド名」列が見つかりません（シート: ${sheetName}）` }); setLoading(false); return; }
          const g = (k,row) => map[k]!==undefined ? String(row[map[k]]??"").trim() : "";
          const items = [];
          for (const row of rows.slice(1)) {
            if (row.every(c => !String(c).trim())) continue;
            const name = g("companyName",row);
            if (!name) continue;
            const rawDate = map.lastUpdate!==undefined ? row[map.lastUpdate] : "";
            const lu = /^\d{4}-\d{2}-\d{2}/.test(String(rawDate)) ? String(rawDate).slice(0,10) : (rawDate ? normDate(String(rawDate)) : "");
            const memoParts = [];
            if (g("personName",row))     memoParts.push(`氏名:${g("personName",row)}`);
            if (g("state",row))          memoParts.push(`状態:${g("state",row)}`);
            items.push({
              id: genId(),
              // メインリストと同じ標準フィールド
              companyName:  name,
              status:       entStatusMap(g("status",row)),
              assignee:     "",
              industry:     g("industry",row),
              storeCount:   g("storeCount",row),
              phone:        "",
              lastCallDate: lu,        // 最終更新日→架電日
              nextCallDate: "",
              memo:         memoParts.join(" / "),
              leadAddedDate: getToday(),
              // 大手専用フィールド
              internalOwner: g("internalOwner",row),
              corpName:      g("corpName",row),
              dealName:      g("dealName",row),
              gmoPhase:      g("gmoPhase",row),
              tryhatchStatus:g("tryhatchStatus",row),
              importedAt: nowIso(), updatedAt: nowIso(),
            });
          }
          setEnterprise(items); // 全置換（最新シートで上書き）
          setLog({ success:true, added: items.length, sheet: sheetName });
          setPage(1);
        } catch(ex) { setLog({ error:`読み込みエラー: ${ex.message}` }); }
        finally { setLoading(false); }
      }, 60);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const saveCell = (id, key, val) => {
    setEditCell(null);
    setEnterprise(prev => prev.map(r => r.id===id ? { ...r, [key]: val, updatedAt: nowIso() } : r));
  };

  const inputCls = "border border-blue-400 rounded px-1 py-0.5 text-xs focus:outline-none bg-white";

  return (
    <div className="space-y-4">
      {/* ツールバー */}
      <div className="bg-white rounded-xl border border-slate-200 p-4 space-y-3">
        <div className="flex flex-wrap items-center gap-2">
          <label className="flex items-center gap-1.5 bg-blue-50 hover:bg-blue-100 border border-blue-300 text-blue-700 px-3 py-2 rounded-lg text-xs font-medium cursor-pointer transition-colors">
            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
            </svg>
            Excelインポート（大手シート）
            <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFile} disabled={loading}/>
          </label>
          {/* 列設定 */}
          <div className="relative" ref={colDropRef}>
            <button onClick={() => setShowColDrop(v=>!v)}
              className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-300 text-slate-600 px-3 py-2 rounded-lg text-xs font-medium transition-colors">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 10h18M3 14h18M10 3v18M14 3v18"/></svg>
              列設定
            </button>
            {showColDrop && (
              <div className="absolute left-0 top-full mt-1 bg-white rounded-xl border border-slate-200 shadow-xl z-20 p-3 w-64 max-h-96 overflow-y-auto">
                <p className="text-xs font-semibold text-slate-400 mb-1 px-1">表示中の列（↑↓で並べ替え）</p>
                {visibleCols.map((key, idx) => {
                  const col = ENT_COLS.find(c=>c.key===key); if (!col) return null;
                  return (
                    <div key={key} className="flex items-center gap-1 px-2 py-1 hover:bg-slate-50 rounded-lg">
                      <div className="flex flex-col">
                        <button disabled={idx===0} onClick={()=>setVisibleCols(p=>{const n=[...p];[n[idx-1],n[idx]]=[n[idx],n[idx-1]];return n;})} className="text-slate-400 hover:text-blue-600 disabled:opacity-20 leading-none text-[10px]">▲</button>
                        <button disabled={idx===visibleCols.length-1} onClick={()=>setVisibleCols(p=>{const n=[...p];[n[idx+1],n[idx]]=[n[idx],n[idx+1]];return n;})} className="text-slate-400 hover:text-blue-600 disabled:opacity-20 leading-none text-[10px]">▼</button>
                      </div>
                      <span className="text-xs text-slate-700 flex-1 truncate">{col.label}</span>
                      {key!=="companyName" && <button onClick={()=>setVisibleCols(p=>p.filter(k=>k!==key))} className="text-rose-400 hover:text-rose-600 text-xs">×</button>}
                    </div>
                  );
                })}
                {ENT_COLS.some(c=>!visibleCols.includes(c.key)) && (<>
                  <p className="text-xs font-semibold text-slate-400 mt-2 mb-1 px-1 border-t border-slate-100 pt-2">非表示の列（＋で追加）</p>
                  {ENT_COLS.filter(c=>!visibleCols.includes(c.key)).map(col => (
                    <button key={col.key} onClick={()=>setVisibleCols(p=>[...p,col.key])} className="flex items-center gap-2 px-2 py-1 hover:bg-slate-50 rounded-lg w-full text-left">
                      <span className="text-blue-500 text-xs">＋</span><span className="text-xs text-slate-500">{col.label}</span>
                    </button>
                  ))}
                </>)}
              </div>
            )}
          </div>
          <span className="text-xs text-slate-400 ml-auto">{enterprise.length.toLocaleString()}件 / 表示: {filtered.length.toLocaleString()}件</span>
        </div>
        {/* ステータスフィルター */}
        {Object.keys(stCounts).length > 0 && (
          <div className="flex flex-wrap gap-1.5">
            <button onClick={()=>{setStFilter("all");setPage(1);}} className={`px-2.5 py-1 text-xs rounded-lg border transition-colors ${stFilter==="all"?"bg-blue-600 text-white border-blue-600":"border-slate-200 text-slate-500 hover:bg-slate-50"}`}>全表示</button>
            {Object.entries(stCounts).sort((a,b)=>b[1]-a[1]).map(([s,c])=>(
              <button key={s} onClick={()=>{setStFilter(stFilter===s?"all":s);setPage(1);}}
                className={`px-2.5 py-1 text-xs rounded-lg border transition-colors ${stFilter===s?"bg-slate-700 text-white border-slate-700":"bg-slate-100 text-slate-600 border-slate-200 hover:bg-slate-50"}`}>
                {s} {c}
              </button>
            ))}
          </div>
        )}
        {/* 検索 */}
        <div className="relative">
          <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>
          <input value={search} onChange={e=>{setSearch(e.target.value);setPage(1);}} placeholder="企業名・法人名・業種・商談名などで検索..."
            className="w-full pl-9 pr-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"/>
        </div>
      </div>

      {log && <div className={`text-xs rounded-lg px-4 py-3 ${log.error?"bg-red-50 border border-red-200 text-red-700":"bg-green-50 border border-green-200 text-green-700"}`}>
        {log.error || `✅ ${log.added.toLocaleString()}件を取り込みました（シート: ${log.sheet}）`}
      </div>}
      {loading && <div className="flex items-center gap-2 text-sm text-slate-500"><span className="w-4 h-4 border-2 border-blue-500 border-t-transparent rounded-full animate-spin"/>インポート処理中...</div>}

      {/* テーブル */}
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-auto max-h-[calc(100vh-300px)]">
          <table className="text-xs border-collapse table-fixed">
            <thead className="bg-slate-50 border-b border-slate-200 sticky top-0 z-10">
              <tr>
                {visibleDefs.map(col => (
                  <th key={col.key} onClick={()=>{ if(sortKey===col.key) setSortDir(d=>d==="asc"?"desc":"asc"); else {setSortKey(col.key);setSortDir("asc");} setPage(1); }}
                    className={`${col.w} px-3 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap cursor-pointer hover:bg-slate-100 bg-slate-50`}>
                    <span className="flex items-center gap-1"><span className="truncate">{col.label}</span>
                      {sortKey===col.key ? <span className="text-blue-500 shrink-0">{sortDir==="asc"?"▲":"▼"}</span> : <span className="text-slate-300 shrink-0">⇅</span>}
                    </span>
                  </th>
                ))}
                <th className="w-[96px] px-3 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap bg-slate-50">現在リスト</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {paginated.length === 0 ? (
                <tr><td colSpan={visibleDefs.length+1} className="text-center py-14 text-slate-400 text-sm">
                  {enterprise.length===0 ? "「大手」シートを含むExcelをインポートしてください。" : "条件に一致するデータがありません。"}
                </td></tr>
              ) : paginated.map(rec => (
                <tr key={rec.id} className="hover:bg-slate-50/60 transition-colors">
                  {visibleDefs.map(col => {
                    const isEd = editCell?.id===rec.id && editCell?.key===col.key && col.key!=="companyName";
                    const open = () => col.key!=="companyName" && setEditCell({ id:rec.id, key:col.key });
                    const save = (v) => saveCell(rec.id, col.key, (col.key==="lastCallDate"||col.key==="nextCallDate")? normDate(v) : v);
                    const cancel = () => setEditCell(null);
                    const val = rec[col.key];
                    const td = (inner) => <td key={col.key} className={`${col.w} px-3 py-2 overflow-hidden whitespace-nowrap align-middle`}>{inner}</td>;

                    // 編集モード
                    if (isEd) {
                      if (col.key==="status") return td(
                        <select autoFocus defaultValue={val||"未架電"} className={`${inputCls} w-full`} onChange={e=>save(e.target.value)} onBlur={cancel} onKeyDown={e=>e.key==="Escape"&&cancel()}>
                          {Object.keys(STATUS_CFG).map(s=><option key={s} value={s}>{s}</option>)}
                        </select>);
                      if (col.key==="leadSource") return td(
                        <select autoFocus defaultValue={val||""} className={`${inputCls} w-full`} onChange={e=>save(e.target.value)} onBlur={cancel} onKeyDown={e=>e.key==="Escape"&&cancel()}>
                          <option value="">—</option>{Object.keys(LEAD_SOURCE_CFG).map(s=><option key={s} value={s}>{s}</option>)}
                        </select>);
                      if (col.key==="absenceReason") return td(
                        <select autoFocus defaultValue={val||""} className={`${inputCls} w-full`} onChange={e=>save(e.target.value)} onBlur={cancel} onKeyDown={e=>e.key==="Escape"&&cancel()}>
                          <option value="">—</option>{Object.keys(ABSENCE_REASON_CFG).map(s=><option key={s} value={s}>{s}</option>)}
                        </select>);
                      if (col.key==="lastCallDate"||col.key==="nextCallDate"||col.key==="leadAddedDate") {
                        const dd = col.key==="lastCallDate" ? (normDate(val)||today) : (normDate(val)||"");
                        return td(<input type="date" autoFocus defaultValue={dd} className={`${inputCls} w-full`}
                          onBlur={e=>save(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")save(e.target.value);if(e.key==="Escape")cancel();}}/>);
                      }
                      if (col.key==="memo") return td(
                        <textarea autoFocus defaultValue={val||""} rows={2} className={`${inputCls} w-full resize-none`}
                          onBlur={e=>save(e.target.value)} onKeyDown={e=>{if(e.key==="Escape")cancel();if(e.key==="Enter"&&e.ctrlKey)save(e.target.value);}}/>);
                      return td(<input type="text" autoFocus defaultValue={val||""} className={`${inputCls} w-full`}
                        onBlur={e=>save(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")save(e.target.value);if(e.key==="Escape")cancel();}}/>);
                    }

                    // 表示モード
                    if (col.key==="companyName") return td(<div className="whitespace-nowrap truncate font-medium text-slate-800" title={String(val||"")}>{val||"—"}</div>);
                    if (col.key==="status") return td(<span onClick={open} className="cursor-pointer"><StatusBadge status={val}/></span>);
                    if (col.key==="leadSource") return td(val ? <span onClick={open} className="cursor-pointer"><LeadSourceBadge source={val}/></span> : <span onClick={open} className="text-slate-300 text-xs cursor-pointer">— 設定</span>);
                    if (col.key==="absenceReason") return td(val ? <span onClick={open} className="cursor-pointer"><AbsenceReasonBadge reason={val}/></span> : <span onClick={open} className="text-slate-300 text-xs cursor-pointer">— 設定</span>);
                    if (col.key==="lastCallDate"||col.key==="nextCallDate"||col.key==="leadAddedDate") {
                      const nd = normDate(val);
                      return td(<span onClick={open} className={`cursor-pointer text-xs hover:bg-slate-100 rounded px-1 ${nd&&nd<today&&col.key==="nextCallDate"?"text-red-600 font-bold":"text-slate-600"}`}>{nd?fmtDate(nd):<span className="text-slate-300">— 設定</span>}</span>);
                    }
                    if (col.key==="storeCount") return td(<span onClick={open} className="cursor-pointer hover:bg-slate-50 rounded px-0.5"><StoreCountCell analysis={analyzeStoreCount(rec, storeIndex)}/></span>);
                    if (col.key==="memo") return td(<div onClick={open} className="truncate text-slate-600 text-xs cursor-pointer hover:bg-slate-50 rounded" title={String(val||"")}>{val||<span className="text-slate-300">—</span>}</div>);
                    return td(<div onClick={open} className="truncate text-slate-600 text-xs cursor-pointer hover:bg-slate-50 rounded" title={String(val||"")}>{val||<span className="text-slate-300">—</span>}</div>);
                  })}
                  <td className="w-[96px] px-3 py-2 whitespace-nowrap align-middle">
                    {isInCurrent(rec.companyName)
                      ? <button onClick={() => {
                          const n = normName(rec.companyName);
                          const matched = records.filter(r => { const rn=normName(r.companyName); return rn===n||rn.includes(n)||n.includes(rn); });
                          setListModal({ name: rec.companyName, matched });
                        }}
                          className="text-teal-700 bg-teal-50 border border-teal-200 hover:bg-teal-100 px-1.5 py-0.5 rounded text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap">
                          ✓ 表示
                        </button>
                      : <span className="text-slate-300 text-xs">—</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {totalPages > 1 && (
          <div className="px-4 py-3 border-t border-slate-100 flex items-center justify-between gap-4 flex-wrap bg-white rounded-b-xl">
            <span className="text-xs text-slate-400">{filtered.length.toLocaleString()}件中 {(page-1)*PAGE+1}–{Math.min(page*PAGE,filtered.length)}件</span>
            <div className="flex gap-1">
              <button onClick={()=>setPage(1)} disabled={page===1} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">«</button>
              <button onClick={()=>setPage(p=>p-1)} disabled={page===1} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">‹</button>
              <span className="px-3 py-1 text-xs text-slate-600">{page} / {totalPages}</span>
              <button onClick={()=>setPage(p=>p+1)} disabled={page===totalPages} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">›</button>
              <button onClick={()=>setPage(totalPages)} disabled={page===totalPages} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">»</button>
            </div>
          </div>
        )}
      </div>

      {/* 現在リスト表示モーダル */}
      {listModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl flex flex-col max-h-[85vh]">
            <div className="flex items-center justify-between px-6 pt-5 pb-4 border-b border-slate-100 shrink-0">
              <div>
                <h2 className="text-base font-bold text-slate-800">📋 現在のリスト — 一致レコード</h2>
                <p className="text-xs text-slate-400 mt-0.5">{listModal.name}</p>
              </div>
              <button onClick={() => setListModal(null)} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
            </div>
            <div className="overflow-y-auto flex-1 px-6 py-4">
              {listModal.matched.length === 0 ? (
                <p className="text-slate-400 text-sm text-center py-8">一致するレコードが見つかりませんでした。</p>
              ) : (
                <div className="space-y-3">
                  {listModal.matched.map(r => {
                    const sc = STATUS_CFG[r.status] ?? {};
                    return (
                      <div key={r.id} className={`rounded-xl border p-4 ${sc.row || "bg-white"} ${sc.border || "border-slate-200"}`}>
                        <div className="flex flex-wrap items-center gap-2 mb-2">
                          <span className="font-semibold text-slate-800 text-sm">{r.companyName}</span>
                          <StatusBadge status={r.status}/>
                          {r.leadSource && <LeadSourceBadge source={r.leadSource}/>}
                        </div>
                        <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-1 text-xs text-slate-600">
                          {r.phone        && <span>📞 {r.phone}</span>}
                          {r.createdBy    && <span>👤 担当: {r.createdBy}</span>}
                          {r.lastCallDate && <span>📅 架電日: {fmtDate(normDate(r.lastCallDate))}</span>}
                          {r.nextCallDate && <span>🔔 次回: {fmtDate(normDate(r.nextCallDate))}</span>}
                          {r.storeCount   && <span>🏪 {r.storeCount}店舗</span>}
                          {r.industry     && <span>🏭 {r.industry}</span>}
                        </div>
                        {r.memo && <p className="mt-2 text-xs text-slate-500 bg-slate-50 rounded-lg px-3 py-2 whitespace-pre-wrap">{r.memo}</p>}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
            <div className="px-6 py-4 border-t border-slate-100 flex justify-end shrink-0">
              <button onClick={() => setListModal(null)} className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">閉じる</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── OrderView（受注案件管理） ──────────────────────────────────────────────────
function OrderView({ orders, setOrders, members }) {
  const savedUI = (() => { try { return JSON.parse(localStorage.getItem("teppou_order_ui_v1") || "{}"); } catch { return {}; } })();
  const [search,      setSearch]      = useState("");
  const [payFilter,   setPayFilter]   = useState("all");
  const [editCell,    setEditCell]    = useState(null);
  const [visibleCols, setVisibleCols] = useState(Array.isArray(savedUI.visibleCols)&&savedUI.visibleCols.length ? savedUI.visibleCols : ORDER_DEFAULT_VISIBLE);
  const [showColDrop, setShowColDrop] = useState(false);
  const [sortKey,     setSortKey]     = useState(savedUI.sortKey ?? "orderDate");
  const [sortDir,     setSortDir]     = useState(savedUI.sortDir || "desc");
  const [page,        setPage]        = useState(1);
  const colDropRef = useRef();
  const PAGE = 100;

  useEffect(() => { try { localStorage.setItem("teppou_order_ui_v1", JSON.stringify({ visibleCols, sortKey, sortDir })); } catch {} }, [visibleCols, sortKey, sortDir]);
  useEffect(() => {
    if (!showColDrop) return;
    const fn = e => { if (colDropRef.current && !colDropRef.current.contains(e.target)) setShowColDrop(false); };
    document.addEventListener("mousedown", fn);
    return () => document.removeEventListener("mousedown", fn);
  }, [showColDrop]);

  const visibleDefs = visibleCols.map(k => ORDER_COLS.find(c=>c.key===k)).filter(Boolean);
  const payCounts = useMemo(() => { const m={}; orders.forEach(o=>{const s=o.payment||"未入金"; m[s]=(m[s]||0)+1;}); return m; }, [orders]);

  const filtered = useMemo(() => {
    let rs = orders.filter(o => {
      if (payFilter!=="all" && (o.payment||"未入金")!==payFilter) return false;
      if (search) { const q=search.toLowerCase(); return ORDER_COLS.some(c=>String(o[c.key]||"").toLowerCase().includes(q)); }
      return true;
    });
    if (sortKey) rs=[...rs].sort((a,b)=>{
      if (sortKey==="amount") { const d=parseAmount(a.amount)-parseAmount(b.amount); return sortDir==="asc"?d:-d; }
      const va=String(a[sortKey]||""), vb=String(b[sortKey]||""); const cmp=va.localeCompare(vb,"ja"); return sortDir==="asc"?cmp:-cmp;
    });
    return rs;
  }, [orders, search, payFilter, sortKey, sortDir]);

  const totalAmount = filtered.reduce((s,o)=>s+parseAmount(o.amount),0);
  const totalPages = Math.max(1, Math.ceil(filtered.length/PAGE));
  const paginated  = filtered.slice((page-1)*PAGE, page*PAGE);

  const addRow = () => {
    setOrders(p => [{ id:genId(), orderDate:getToday(), companyName:"", plan:"スタンダードプラン", amount:"", assignee:"", payment:"未入金", startDate:"", memo:"", updatedAt:nowIso() }, ...p]);
    setPage(1);
  };
  const saveCell = (id,key,val) => { setEditCell(null); setOrders(p=>p.map(o=>o.id===id?{...o,[key]:val,updatedAt:nowIso()}:o)); };
  const delRow = (id) => { if(window.confirm("この受注案件を削除しますか？")) setOrders(p=>p.filter(o=>o.id!==id)); };

  const inputCls = "border border-blue-400 rounded px-1 py-0.5 text-xs focus:outline-none bg-white";

  return (
    <div className="space-y-4">
      <div className="bg-white rounded-xl border border-slate-200 p-4 space-y-3">
        <div className="flex flex-wrap items-center gap-2">
          <button onClick={addRow} className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-semibold transition-colors">
            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4"/></svg>
            受注を追加
          </button>
          <div className="relative" ref={colDropRef}>
            <button onClick={()=>setShowColDrop(v=>!v)} className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-300 text-slate-600 px-3 py-2 rounded-lg text-xs font-medium transition-colors">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 10h18M3 14h18M10 3v18M14 3v18"/></svg>列設定
            </button>
            {showColDrop && (
              <div className="absolute left-0 top-full mt-1 bg-white rounded-xl border border-slate-200 shadow-xl z-20 p-3 w-64 max-h-96 overflow-y-auto">
                <p className="text-xs font-semibold text-slate-400 mb-1 px-1">表示中の列（↑↓で並べ替え）</p>
                {visibleCols.map((key,idx)=>{ const col=ORDER_COLS.find(c=>c.key===key); if(!col) return null; return (
                  <div key={key} className="flex items-center gap-1 px-2 py-1 hover:bg-slate-50 rounded-lg">
                    <div className="flex flex-col">
                      <button disabled={idx===0} onClick={()=>setVisibleCols(p=>{const n=[...p];[n[idx-1],n[idx]]=[n[idx],n[idx-1]];return n;})} className="text-slate-400 hover:text-blue-600 disabled:opacity-20 leading-none text-[10px]">▲</button>
                      <button disabled={idx===visibleCols.length-1} onClick={()=>setVisibleCols(p=>{const n=[...p];[n[idx+1],n[idx]]=[n[idx],n[idx+1]];return n;})} className="text-slate-400 hover:text-blue-600 disabled:opacity-20 leading-none text-[10px]">▼</button>
                    </div>
                    <span className="text-xs text-slate-700 flex-1 truncate">{col.label}</span>
                    {key!=="companyName" && <button onClick={()=>setVisibleCols(p=>p.filter(k=>k!==key))} className="text-rose-400 hover:text-rose-600 text-xs">×</button>}
                  </div>);
                })}
                {ORDER_COLS.some(c=>!visibleCols.includes(c.key)) && (<>
                  <p className="text-xs font-semibold text-slate-400 mt-2 mb-1 px-1 border-t border-slate-100 pt-2">非表示の列</p>
                  {ORDER_COLS.filter(c=>!visibleCols.includes(c.key)).map(col=>(
                    <button key={col.key} onClick={()=>setVisibleCols(p=>[...p,col.key])} className="flex items-center gap-2 px-2 py-1 hover:bg-slate-50 rounded-lg w-full text-left"><span className="text-blue-500 text-xs">＋</span><span className="text-xs text-slate-500">{col.label}</span></button>
                  ))}
                </>)}
              </div>
            )}
          </div>
          <span className="text-xs text-slate-400 ml-auto">{orders.length.toLocaleString()}件 ／ 表示合計: <strong className="text-emerald-700">¥{totalAmount.toLocaleString()}</strong></span>
        </div>
        {Object.keys(payCounts).length>0 && (
          <div className="flex flex-wrap gap-1.5">
            <button onClick={()=>{setPayFilter("all");setPage(1);}} className={`px-2.5 py-1 text-xs rounded-lg border transition-colors ${payFilter==="all"?"bg-blue-600 text-white border-blue-600":"border-slate-200 text-slate-500 hover:bg-slate-50"}`}>全表示</button>
            {Object.keys(PAYMENT_STATUS_CFG).map(s=> payCounts[s]?(
              <button key={s} onClick={()=>{setPayFilter(payFilter===s?"all":s);setPage(1);}}
                className={`flex items-center gap-1 px-2.5 py-1 text-xs rounded-lg border transition-colors ${payFilter===s?`${PAYMENT_STATUS_CFG[s].bg} ${PAYMENT_STATUS_CFG[s].text} border-current ring-2 ring-blue-400 ring-offset-1`:`${PAYMENT_STATUS_CFG[s].bg} ${PAYMENT_STATUS_CFG[s].text} border-slate-200`}`}>
                <span className={`w-1.5 h-1.5 rounded-full ${PAYMENT_STATUS_CFG[s].dot}`}/>{s} {payCounts[s]}
              </button>
            ):null)}
          </div>
        )}
        <div className="relative">
          <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>
          <input value={search} onChange={e=>{setSearch(e.target.value);setPage(1);}} placeholder="企業名・プラン・担当者・メモで検索..." className="w-full pl-9 pr-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"/>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-auto max-h-[calc(100vh-300px)]">
          <table className="text-xs border-collapse table-fixed">
            <thead className="bg-slate-50 border-b border-slate-200 sticky top-0 z-10">
              <tr>
                {visibleDefs.map(col=>(
                  <th key={col.key} onClick={()=>{ if(sortKey===col.key) setSortDir(d=>d==="asc"?"desc":"asc"); else {setSortKey(col.key);setSortDir("asc");} setPage(1); }}
                    className={`${col.w} px-3 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap cursor-pointer hover:bg-slate-100 bg-slate-50`}>
                    <span className="flex items-center gap-1"><span className="truncate">{col.label}</span>{sortKey===col.key?<span className="text-blue-500 shrink-0">{sortDir==="asc"?"▲":"▼"}</span>:<span className="text-slate-300 shrink-0">⇅</span>}</span>
                  </th>
                ))}
                <th className="w-[64px] px-3 py-2.5 text-left text-xs font-semibold text-slate-500 bg-slate-50">操作</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {paginated.length===0 ? (
                <tr><td colSpan={visibleDefs.length+1} className="text-center py-14 text-slate-400 text-sm">受注案件がありません。「受注を追加」またはCSVインポートで登録してください。</td></tr>
              ) : paginated.map(o=>(
                <tr key={o.id} className="hover:bg-slate-50/60 transition-colors">
                  {visibleDefs.map(col=>{
                    const isEd = editCell?.id===o.id && editCell?.key===col.key;
                    const open = () => setEditCell({ id:o.id, key:col.key });
                    const save = (v) => saveCell(o.id, col.key, (col.key==="orderDate"||col.key==="startDate")?normDate(v):v);
                    const cancel = () => setEditCell(null);
                    const val = o[col.key];
                    const td = (inner) => <td key={col.key} className={`${col.w} px-3 py-2 overflow-hidden whitespace-nowrap align-middle`}>{inner}</td>;
                    if (isEd) {
                      if (col.key==="plan") return td(<select autoFocus defaultValue={val||""} className={`${inputCls} w-full`} onChange={e=>save(e.target.value)} onBlur={cancel} onKeyDown={e=>e.key==="Escape"&&cancel()}><option value="">—</option>{ORDER_PLANS.map(s=><option key={s} value={s}>{s}</option>)}</select>);
                      if (col.key==="payment") return td(<select autoFocus defaultValue={val||"未入金"} className={`${inputCls} w-full`} onChange={e=>save(e.target.value)} onBlur={cancel} onKeyDown={e=>e.key==="Escape"&&cancel()}>{Object.keys(PAYMENT_STATUS_CFG).map(s=><option key={s} value={s}>{s}</option>)}</select>);
                      if (col.key==="assignee") return td(<><input list="ent-members" autoFocus defaultValue={val||""} className={`${inputCls} w-full`} onBlur={e=>save(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")save(e.target.value);if(e.key==="Escape")cancel();}}/><datalist id="ent-members">{members.map(m=><option key={m} value={m}/>)}</datalist></>);
                      if (col.key==="orderDate"||col.key==="startDate") return td(<input type="date" autoFocus defaultValue={normDate(val)||(col.key==="orderDate"?getToday():"")} className={`${inputCls} w-full`} onBlur={e=>save(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")save(e.target.value);if(e.key==="Escape")cancel();}}/>);
                      if (col.key==="amount") return td(<input type="text" autoFocus defaultValue={val||""} className={`${inputCls} w-full`} onBlur={e=>save(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")save(e.target.value);if(e.key==="Escape")cancel();}}/>);
                      if (col.key==="memo") return td(<textarea autoFocus defaultValue={val||""} rows={2} className={`${inputCls} w-full resize-none`} onBlur={e=>save(e.target.value)} onKeyDown={e=>{if(e.key==="Escape")cancel();if(e.key==="Enter"&&e.ctrlKey)save(e.target.value);}}/>);
                      return td(<input type="text" autoFocus defaultValue={val||""} className={`${inputCls} w-full`} onBlur={e=>save(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")save(e.target.value);if(e.key==="Escape")cancel();}}/>);
                    }
                    if (col.key==="companyName") return td(<div onClick={open} className="flex items-center gap-1 cursor-pointer hover:text-blue-600" title={String(val||"")}><CompanyLogo logoUrl={o.logoUrl} url={o.hpSite} name={val} /><span className="whitespace-nowrap truncate font-medium text-slate-800">{val||<span className="text-slate-300">— 入力</span>}</span></div>);
                    if (col.key==="plan") return td(val?<span onClick={open} className="cursor-pointer inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-indigo-100 text-indigo-700 border border-indigo-200 truncate max-w-full">{val}</span>:<span onClick={open} className="text-slate-300 text-xs cursor-pointer">— 設定</span>);
                    if (col.key==="payment") { const c=PAYMENT_STATUS_CFG[val||"未入金"]??{}; return td(<span onClick={open} className={`cursor-pointer inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold border border-black/10 ${c.bg} ${c.text}`}><span className={`w-1.5 h-1.5 rounded-full ${c.dot}`}/>{val||"未入金"}</span>); }
                    if (col.key==="amount") return td(<span onClick={open} className="cursor-pointer text-slate-700 font-semibold hover:bg-slate-50 rounded px-1">{val?`¥${parseAmount(val).toLocaleString()}`:<span className="text-slate-300 font-normal">—</span>}</span>);
                    if (col.key==="orderDate"||col.key==="startDate") { const nd=normDate(val); return td(<span onClick={open} className="cursor-pointer text-slate-600 hover:bg-slate-100 rounded px-1">{nd?fmtDate(nd):<span className="text-slate-300">— 設定</span>}</span>); }
                    if (col.key==="memo") return td(<div onClick={open} className="truncate text-slate-600 cursor-pointer hover:bg-slate-50 rounded" title={String(val||"")}>{val||<span className="text-slate-300">—</span>}</div>);
                    return td(<div onClick={open} className="truncate text-slate-600 cursor-pointer hover:bg-slate-50 rounded" title={String(val||"")}>{val||<span className="text-slate-300">—</span>}</div>);
                  })}
                  <td className="w-[64px] px-3 py-2 whitespace-nowrap align-middle">
                    <button onClick={()=>delRow(o.id)} className="text-xs text-rose-500 hover:text-rose-700 font-medium">削除</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {totalPages>1 && (
          <div className="px-4 py-3 border-t border-slate-100 flex items-center justify-between gap-4 flex-wrap bg-white rounded-b-xl">
            <span className="text-xs text-slate-400">{filtered.length.toLocaleString()}件中 {(page-1)*PAGE+1}–{Math.min(page*PAGE,filtered.length)}件</span>
            <div className="flex gap-1">
              <button onClick={()=>setPage(1)} disabled={page===1} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">«</button>
              <button onClick={()=>setPage(p=>p-1)} disabled={page===1} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">‹</button>
              <span className="px-3 py-1 text-xs text-slate-600">{page} / {totalPages}</span>
              <button onClick={()=>setPage(p=>p+1)} disabled={page===totalPages} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">›</button>
              <button onClick={()=>setPage(totalPages)} disabled={page===totalPages} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">»</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── PastMgmtView ───────────────────────────────────────────────────────────────
function PastMgmtView({ pastMgmt, setPastMgmt, records, onGoToList, onAddToList, onBeforeImport }) {
  const [search,       setSearch]       = useState("");
  const [editCell,     setEditCell]     = useState(null);
  const [log,          setLog]          = useState(null);
  const [loading,      setLoading]      = useState(false);
  const [notInListOnly, setNotInListOnly] = useState(false); // 現在リストにない案件のみ
  const [excludeTodayCalled, setExcludeTodayCalled] = useState(false); // 今日架電を除外
  const [addedIds,     setAddedIds]     = useState(new Set()); // 追加済み表示用
  const savedPastUI = (() => { try { return JSON.parse(localStorage.getItem(PAST_UI_KEY) || "{}"); } catch { return {}; } })();
  const [visibleCols,  setVisibleCols]  = useState(Array.isArray(savedPastUI.visibleCols) && savedPastUI.visibleCols.length ? savedPastUI.visibleCols : DEFAULT_PAST_VISIBLE);
  const [showColDrop,  setShowColDrop]  = useState(false);
  const [sortKey,      setSortKey]      = useState(savedPastUI.sortKey ?? null);
  const [sortDir,      setSortDir]      = useState(savedPastUI.sortDir || "asc");
  const [showDupe,     setShowDupe]     = useState(false);
  const [listModal,    setListModal]    = useState(null); // { name, matched[] }
  const [copiedId,     setCopiedId]     = useState(null);
  const colDropRef = useRef();
  const fileRef    = useRef();
  const today = getToday();
  const soon  = (() => { const d=new Date(); d.setDate(d.getDate()+3); return d.toISOString().slice(0,10); })();

  // UI 状態（列設定・ソート）を保存
  useEffect(() => {
    try { localStorage.setItem(PAST_UI_KEY, JSON.stringify({ visibleCols, sortKey, sortDir })); } catch {}
  }, [visibleCols, sortKey, sortDir]);

  // 現在リストの企業名セット（照合用）
  const currentNames = useMemo(() =>
    new Set(records.map(r => normName(r.companyName))), [records]);
  // 店舗数分析インデックス（過去商談＋営業リスト、一度だけ構築）
  const storeIndex = useMemo(() => buildStoreIndex(pastMgmt, records), [pastMgmt, records]);
  // 進捗プルダウンの選択肢（データ内の状況値から動的生成）
  const progressOptions = useMemo(() => {
    const s = new Set();
    pastMgmt.forEach(r => { if (r.progress && String(r.progress).trim()) s.add(String(r.progress).trim()); });
    return [...s].sort();
  }, [pastMgmt]);

  const isInCurrent = (name) => {
    const n = normName(name);
    return [...currentNames].some(cn => cn === n || cn.includes(n) || n.includes(cn));
  };

  // フィルタ・ソート
  const filtered = useMemo(() => {
    let rs = pastMgmt.filter(r => {
      if (notInListOnly && isInCurrent(r.companyName)) return false;
      if (excludeTodayCalled && normDate(r.lastCallDate) === today) return false;
      if (search) {
        const q = search.toLowerCase();
        return ALL_PAST_COLS.some(c => String(r[c.key]||"").toLowerCase().includes(q));
      }
      return true;
    });
    if (sortKey) {
      rs = [...rs].sort((a, b) => {
        const va = String(a[sortKey]||""), vb = String(b[sortKey]||"");
        const cmp = va.localeCompare(vb, "ja");
        return sortDir === "asc" ? cmp : -cmp;
      });
    }
    return rs;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pastMgmt, search, sortKey, sortDir, notInListOnly, excludeTodayCalled, currentNames]);

  const [page, setPage] = useState(1);
  const PAGE = 100;
  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE));
  const paginated  = filtered.slice((page-1)*PAGE, page*PAGE);

  // 表示列定義
  const visibleDefs = visibleCols.map(k => ALL_PAST_COLS.find(c => c.key === k)).filter(Boolean);

  // カラムドロップ外側クリックで閉じる
  useEffect(() => {
    if (!showColDrop) return;
    const fn = e => { if (colDropRef.current && !colDropRef.current.contains(e.target)) setShowColDrop(false); };
    document.addEventListener("mousedown", fn);
    return () => document.removeEventListener("mousedown", fn);
  }, [showColDrop]);

  // インポート
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    const doProcess = (rows) => {
      setTimeout(() => {
        try {
          const headers = rows[0];
          const map = mapPastMgmtHeaders(headers);
          // 通常のsalesヘッダーも試みる
          const salesMap = mapSalesHeaders(headers);
          const compIdx = map.companyName ?? salesMap.companyName;
          if (compIdx === undefined) {
            setLog({ error:`「取引先名」「企業名」列が見つかりません。検出: ${headers.filter(h=>h).slice(0,6).join("、")}` });
            setLoading(false); return;
          }
          const items = []; let skipped = 0;
          const g = (idx, row) => idx !== undefined ? String(row[idx]??'').trim() : "";
          for (const row of rows.slice(1)) {
            if (row.every(c => !String(c).trim())) continue;
            const name = String(row[compIdx]??"").trim();
            if (!name) { skipped++; continue; }
            const rawDate = map.targetDate !== undefined ? row[map.targetDate] : "";
            const td = rawDate instanceof Date ? rawDate.toISOString().slice(0,10)
                     : rawDate ? normDate(String(rawDate)) : "";
            items.push({
              id: genId(),
              // ALL_COLUMNS フィールド
              companyName:   name,
              phone:         g(salesMap.phone, row),
              email:         g(salesMap.email, row),
              hpSite:        g(salesMap.hpSite, row),
              gbp:           g(salesMap.gbp, row),
              gbpSiteUrl:    g(salesMap.gbpSiteUrl, row),
              gbpManagement: g(salesMap.gbpManagement, row),
              status:        g(salesMap.status, row) || "未架電",
              assignee:      g(salesMap.assignee ?? map.creator, row), // 作成者/担当者→商談所有者
              createdBy:     "",
              department:    g(salesMap.department, row),
              industry:      g(salesMap.industry, row),
              leadSource:    g(salesMap.leadSource, row),
              absenceReason: g(salesMap.absenceReason, row),
              memo:          g(salesMap.memo, row),
              storeCount:    g(salesMap.storeCount, row),
              refusalReason: g(salesMap.refusalReason, row),
              posting:       g(salesMap.posting, row),
              review:        g(salesMap.review, row),
              sns:           g(salesMap.sns, row),
              instagram:     g(salesMap.instagram, row),
              line:          g(salesMap.line, row),
              facebook:      g(salesMap.facebook, row),
              twitter:       g(salesMap.twitter, row),
              os:            g(salesMap.os, row),
              mailFlag:      g(salesMap.mailFlag, row),
              lastCallDate:  normDate(g(salesMap.lastCallDate, row)),
              nextCallDate:  normDate(g(salesMap.nextCallDate, row)),
              leadAddedDate: getToday(),
              // 過去商談専用
              targetDate:       td,
              progress:         map.progress !== undefined ? String(row[map.progress]??"").trim() : "",
              importedAt: nowIso(), updatedAt: nowIso(),
            });
          }
          onBeforeImport?.(); // インポート前に自動退避
          // 重複チェックなし: すべて追加
          setPastMgmt(prev => {
            const next = [...prev, ...items];
            setLog({ success:true, added: items.length, skipped });
            return next;
          });
          setPage(1);
        } catch(ex) {
          setLog({ error: `エラー: ${ex.message}` });
        } finally { setLoading(false); }
      }, 60);
    };
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = ev => {
        const wb = XLSX.read(ev.target.result, { type:"array", cellDates:true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        doProcess(XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:"" }));
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = ev => doProcess(parseCSV(ev.target.result));
      reader.readAsText(file, "UTF-8");
    }
    e.target.value = "";
  };

  const saveCell = (id, key, val) => {
    setEditCell(null);
    const normalized = (key === "lastCallDate" || key === "nextCallDate" || key === "targetDate")
      ? normDate(val) : val;
    setPastMgmt(prev => prev.map(r => {
      if (r.id !== id) return r;
      const updated = { ...r, [key]: normalized, updatedAt: nowIso() };
      return updated;
    }));
  };


  const inputCls = "border border-blue-400 rounded px-1 py-0.5 text-xs focus:outline-none bg-white";

  return (
    <div className="space-y-4">
      {/* ツールバー */}
      <div className="bg-white rounded-xl border border-slate-200 p-4 space-y-3">
        <div className="flex flex-wrap items-center gap-2">
          <label className="flex items-center gap-1.5 bg-blue-50 hover:bg-blue-100 border border-blue-300 text-blue-700 px-3 py-2 rounded-lg text-xs font-medium cursor-pointer transition-colors">
            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
            </svg>
            Excelインポート
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile} disabled={loading}/>
          </label>
          {/* 列設定 */}
          <div className="relative" ref={colDropRef}>
            <button onClick={() => setShowColDrop(v=>!v)}
              className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-300 text-slate-600 px-3 py-2 rounded-lg text-xs font-medium transition-colors">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 10h18M3 14h18M10 3v18M14 3v18"/>
              </svg>
              列設定
              <svg className={`w-3 h-3 transition-transform ${showColDrop?"rotate-180":""}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7"/>
              </svg>
            </button>
            {showColDrop && (
              <div className="absolute left-0 top-full mt-1 bg-white rounded-xl border border-slate-200 shadow-xl z-20 p-3 w-64 max-h-96 overflow-y-auto">
                <p className="text-xs font-semibold text-slate-400 mb-1 px-1">表示中の列（↑↓で並べ替え）</p>
                {visibleCols.map((key, idx) => {
                  const col = ALL_PAST_COLS.find(c => c.key === key);
                  if (!col) return null;
                  return (
                    <div key={key} className="flex items-center gap-1 px-2 py-1 hover:bg-slate-50 rounded-lg">
                      <div className="flex flex-col">
                        <button disabled={idx===0}
                          onClick={() => setVisibleCols(p => { const n=[...p]; [n[idx-1],n[idx]]=[n[idx],n[idx-1]]; return n; })}
                          className="text-slate-400 hover:text-blue-600 disabled:opacity-20 leading-none text-[10px]">▲</button>
                        <button disabled={idx===visibleCols.length-1}
                          onClick={() => setVisibleCols(p => { const n=[...p]; [n[idx+1],n[idx]]=[n[idx],n[idx+1]]; return n; })}
                          className="text-slate-400 hover:text-blue-600 disabled:opacity-20 leading-none text-[10px]">▼</button>
                      </div>
                      <span className="text-xs text-slate-700 flex-1 truncate">{col.label}</span>
                      {col.required
                        ? <span className="text-xs text-slate-300">必須</span>
                        : <button onClick={() => setVisibleCols(p => p.filter(k => k!==key))}
                            className="text-rose-400 hover:text-rose-600 text-xs">×</button>}
                    </div>
                  );
                })}
                {ALL_PAST_COLS.some(c => !visibleCols.includes(c.key)) && (
                  <>
                    <p className="text-xs font-semibold text-slate-400 mt-2 mb-1 px-1 border-t border-slate-100 pt-2">非表示の列（＋で追加）</p>
                    {ALL_PAST_COLS.filter(c => !visibleCols.includes(c.key)).map(col => (
                      <button key={col.key}
                        onClick={() => setVisibleCols(p => [...p, col.key])}
                        className="flex items-center gap-2 px-2 py-1 hover:bg-slate-50 rounded-lg w-full text-left">
                        <span className="text-blue-500 text-xs">＋</span>
                        <span className="text-xs text-slate-500">{col.label}</span>
                      </button>
                    ))}
                  </>
                )}
              </div>
            )}
          </div>
          {pastMgmt.length > 0 && (
            <>
              <button onClick={() => setShowDupe(true)}
                className="flex items-center gap-1.5 bg-amber-50 hover:bg-amber-100 border border-amber-300 text-amber-700 px-3 py-2 rounded-lg text-xs font-medium transition-colors">
                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z"/>
                </svg>
                重複クレンジング
              </button>
            </>
          )}
          <span className="text-xs text-slate-400 ml-auto">{pastMgmt.length.toLocaleString()}件 / 表示: {filtered.length.toLocaleString()}件</span>
        </div>
        {/* 検索・フィルター */}
        <div className="flex flex-wrap gap-2">
          <div className="relative flex-1 min-w-48">
            <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/>
            </svg>
            <input value={search} onChange={e => { setSearch(e.target.value); setPage(1); }}
              placeholder="企業名・担当者・メモなどで検索..."
              className="w-full pl-9 pr-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"/>
          </div>
          <button onClick={() => { setNotInListOnly(v=>!v); setPage(1); }}
            className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-medium border transition-colors whitespace-nowrap
              ${notInListOnly ? "bg-orange-600 text-white border-orange-600" : "bg-white text-slate-600 border-slate-300 hover:bg-slate-50"}`}>
            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M18.364 18.364A9 9 0 005.636 5.636m12.728 12.728A9 9 0 015.636 5.636m12.728 12.728L5.636 5.636"/>
            </svg>
            リスト未登録のみ{notInListOnly && ` (${filtered.length})`}
          </button>
          <button onClick={() => { setExcludeTodayCalled(v=>!v); setPage(1); }}
            className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-medium border transition-colors whitespace-nowrap
              ${excludeTodayCalled ? "bg-amber-600 text-white border-amber-600" : "bg-white text-slate-600 border-slate-300 hover:bg-slate-50"}`}>
            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M18.364 18.364A9 9 0 005.636 5.636m12.728 12.728A9 9 0 015.636 5.636m12.728 12.728L5.636 5.636"/>
            </svg>
            今日架電を除外
          </button>
        </div>
      </div>

      {/* ログ */}
      {log && <div className={`text-xs rounded-lg px-4 py-3 ${log.error ? "bg-red-50 border border-red-200 text-red-700" : "bg-green-50 border border-green-200 text-green-700"}`}>
        {log.error || `✅ ${log.added.toLocaleString()}件追加（空行スキップ: ${log.skipped}件）`}
      </div>}
      {loading && <div className="flex items-center gap-2 text-sm text-slate-500">
        <div className="w-4 h-4 border-2 border-blue-500 border-t-transparent rounded-full animate-spin"/> インポート処理中...
      </div>}

      {/* テーブル */}
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-auto max-h-[calc(100vh-300px)]">
        <table className="text-xs border-collapse table-fixed">
            <thead className="bg-slate-50 border-b border-slate-200 sticky top-0 z-10">
              <tr>
                {visibleDefs.map(col => (
                  <th key={col.key}
                    onClick={() => { if(sortKey===col.key) setSortDir(d=>d==="asc"?"desc":"asc"); else{setSortKey(col.key);setSortDir("asc");} setPage(1); }}
                    className={`${col.w||"w-[120px]"} px-3 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap cursor-pointer hover:bg-slate-100 transition-colors bg-slate-50`}>
                    <span className="flex items-center gap-1"><span className="truncate">{col.label}</span>
                      {sortKey===col.key ? <span className="text-blue-500 shrink-0">{sortDir==="asc"?"▲":"▼"}</span> : <span className="text-slate-300 shrink-0">⇅</span>}
                    </span>
                  </th>
                ))}
                <th className="w-[88px] px-3 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap bg-slate-50">現在リスト</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {paginated.length === 0 ? (
                <tr><td colSpan={visibleDefs.length+1} className="text-center py-14 text-slate-400 text-sm">
                  {pastMgmt.length === 0 ? "Excelファイルをインポートしてください。" : "条件に一致するデータがありません。"}
                </td></tr>
              ) : paginated.map(rec => {
                const inCurrent = isInCurrent(rec.companyName);
                const isToday   = normDate(rec.lastCallDate) === today;
                const rowColor  = isToday
                  ? "bg-yellow-100 border-l-4 border-yellow-400"
                  : (STATUS_CFG[rec.status]?.row || (inCurrent ? "bg-teal-50/30" : ""));
                return (
                  <tr key={rec.id} className={`transition-colors hover:brightness-95 ${rowColor}`}>
                    {visibleDefs.map(col => {
                      const isEd  = editCell?.id===rec.id && editCell?.key===col.key && col.key !== "companyName";
                      const open  = () => setEditCell({ id:rec.id, key:col.key });
                      const save  = (v) => saveCell(rec.id, col.key, v);
                      const cancel= () => setEditCell(null);
                      const val   = rec[col.key];

                      if (isEd) {
                        if (col.key === "status") return <td key={col.key} className="px-3 py-2">
                          <select autoFocus defaultValue={val||"未架電"} className={`${inputCls} w-36`}
                            onChange={e=>save(e.target.value)} onBlur={cancel} onKeyDown={e=>e.key==="Escape"&&cancel()}>
                            {Object.keys(STATUS_CFG).map(s=><option key={s} value={s}>{s}</option>)}
                          </select></td>;
                        if (col.key === "progress") return <td key={col.key} className="px-3 py-2">
                          <select autoFocus defaultValue={val||""} className={`${inputCls} w-40`}
                            onChange={e=>save(e.target.value)} onBlur={cancel} onKeyDown={e=>e.key==="Escape"&&cancel()}>
                            <option value="">—</option>
                            {progressOptions.map(s=><option key={s} value={s}>{s}</option>)}
                          </select></td>;
                        if (col.key === "leadSource") return <td key={col.key} className="px-3 py-2">
                          <select autoFocus defaultValue={val||""} className={`${inputCls} w-32`}
                            onChange={e=>save(e.target.value)} onBlur={cancel} onKeyDown={e=>e.key==="Escape"&&cancel()}>
                            <option value="">—</option>
                            {Object.keys(LEAD_SOURCE_CFG).map(s=><option key={s} value={s}>{s}</option>)}
                          </select></td>;
                        if (col.key === "absenceReason") return <td key={col.key} className="px-3 py-2">
                          <select autoFocus defaultValue={val||""} className={`${inputCls} w-28`}
                            onChange={e=>save(e.target.value)} onBlur={cancel} onKeyDown={e=>e.key==="Escape"&&cancel()}>
                            <option value="">—</option>
                            {Object.keys(ABSENCE_REASON_CFG).map(s=><option key={s} value={s}>{s}</option>)}
                          </select></td>;
                        if (col.key==="lastCallDate"||col.key==="nextCallDate"||col.key==="targetDate"||col.key==="leadAddedDate") {
                          const dateDefault = col.key === "lastCallDate"
                            ? (normDate(val) || today)   // 架電日は未設定なら本日
                            : (normDate(val) || "");
                          return <td key={col.key} className="px-3 py-2">
                            <input type="date" autoFocus defaultValue={dateDefault}
                              className={`${inputCls} w-32`}
                              onBlur={e=>save(e.target.value)}
                              onKeyDown={e=>{if(e.key==="Enter")save(e.target.value);if(e.key==="Escape")cancel();}}/></td>;
                        }
                        if (col.key==="memo") return <td key={col.key} className="px-3 py-2">
                          <textarea autoFocus defaultValue={val||""} rows={2} className={`${inputCls} w-48 resize-none`}
                            onBlur={e=>save(e.target.value)} onKeyDown={e=>{if(e.key==="Escape")cancel();if(e.key==="Enter"&&e.ctrlKey)save(e.target.value);}}/></td>;
                        return <td key={col.key} className="px-3 py-2">
                          <input type="text" autoFocus defaultValue={val||""}
                            className={`${inputCls} w-32`}
                            onBlur={e=>save(e.target.value)}
                            onKeyDown={e=>{if(e.key==="Enter")save(e.target.value);if(e.key==="Escape")cancel();}}/></td>;
                      }

                      // VIEW モード（企業名はコピーのみ・編集不可）
                      if (col.key==="companyName") return (
                        <td key={col.key} className="px-3 py-2 max-w-44">
                          <button
                            onClick={() => {
                              navigator.clipboard.writeText(val||"").then(() => {
                                setCopiedId(rec.id);
                                setTimeout(() => setCopiedId(null), 1500);
                              });
                              // 架電日を本日に自動更新
                              setPastMgmt(prev => prev.map(r => r.id===rec.id ? { ...r, lastCallDate: getToday(), updatedAt: nowIso() } : r));
                            }}
                            title="クリックでコピー（架電日を本日に更新）"
                            className={`group flex items-center gap-1 text-left w-full min-w-0 transition-colors ${copiedId===rec.id ? "text-green-600" : "text-slate-800 hover:text-blue-600"}`}>
                            <span className="font-medium text-xs whitespace-nowrap truncate max-w-40">{val||"—"}</span>
                            {copiedId === rec.id
                              ? <span className="text-green-500 text-xs shrink-0">✓</span>
                              : <svg className="w-3 h-3 shrink-0 text-slate-300 group-hover:text-blue-400 opacity-0 group-hover:opacity-100 transition-opacity"
                                  fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                                    d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z"/>
                                </svg>}
                          </button>
                        </td>
                      );
                      if (col.key==="status") return <td key={col.key} className="px-3 py-2">
                        <span onClick={open} className="cursor-pointer"><StatusBadge status={val}/></span></td>;
                      if (col.key==="progress") return <td key={col.key} className="px-3 py-2 overflow-hidden">
                        {val ? <span onClick={open} className="cursor-pointer inline-flex items-center px-2 py-0.5 rounded-full text-xs font-semibold bg-violet-100 text-violet-700 border border-violet-200 truncate max-w-full">{val}</span>
                             : <span onClick={open} className="text-slate-300 text-xs cursor-pointer hover:text-slate-500">— 設定</span>}</td>;
                      if (col.key==="leadSource") return <td key={col.key} className="px-3 py-2">
                        <span onClick={open} className="cursor-pointer hover:opacity-80"><LeadSourceBadge source={val}/></span></td>;
                      if (col.key==="absenceReason") return <td key={col.key} className="px-3 py-2">
                        {val ? <span onClick={open} className="cursor-pointer hover:opacity-80"><AbsenceReasonBadge reason={val}/></span>
                             : <span onClick={open} className="text-slate-300 text-xs cursor-pointer hover:text-slate-500">— 設定</span>}</td>;
                      if (col.key==="lastCallDate"||col.key==="nextCallDate"||col.key==="targetDate"||col.key==="leadAddedDate") {
                        const nd = normDate(val) || (col.key==="leadAddedDate" ? today : "");
                        return <td key={col.key} className="px-3 py-2 whitespace-nowrap">
                          <span onClick={open} className={`cursor-pointer text-xs hover:bg-slate-100 rounded px-1 transition-colors ${nd&&nd<today&&col.key==="nextCallDate"?"text-red-600 font-bold":"text-slate-600"}`}>
                            {nd ? fmtDate(nd) : <span className="text-slate-300">— 設定</span>}
                          </span></td>;
                      }
                      if (col.key==="memo") return <td key={col.key} className="px-3 py-2 max-w-48">
                        <span onClick={open} className="cursor-pointer text-slate-600 text-xs block max-w-44 truncate hover:bg-slate-50 rounded transition-colors" title={val||""}>
                          {val||<span className="text-slate-300">—</span>}</span></td>;
                      if (col.key==="storeCount") return <td key={col.key} className="px-3 py-2">
                        <span onClick={open} className="cursor-pointer hover:bg-slate-50 rounded px-0.5 transition-colors">
                          <StoreCountCell analysis={analyzeStoreCount(rec, storeIndex)} />
                        </span></td>;
                      return <td key={col.key} className="px-3 py-2">
                        <span onClick={open} className="text-slate-700 text-xs cursor-pointer hover:bg-slate-50 rounded px-0.5 transition-colors">
                          {val||<span className="text-slate-300">—</span>}</span></td>;
                    })}
                    <td className="px-3 py-2 whitespace-nowrap min-w-[72px]">
                      {inCurrent
                        ? <button onClick={() => {
                            const n = normName(rec.companyName);
                            const matched = records.filter(r => {
                              const rn = normName(r.companyName);
                              return rn === n || rn.includes(n) || n.includes(rn);
                            });
                            setListModal({ name: rec.companyName, matched });
                          }}
                            className="text-teal-700 bg-teal-50 border border-teal-200 hover:bg-teal-100 px-1.5 py-0.5 rounded text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap">
                            ✓ 表示
                          </button>
                        : addedIds.has(rec.id)
                        ? <span className="text-green-600 text-xs font-semibold whitespace-nowrap">✓ 追加済</span>
                        : <button onClick={() => {
                            onAddToList?.(rec);
                            setAddedIds(p => { const s = new Set(p); s.add(rec.id); return s; });
                          }}
                            className="flex items-center gap-1 text-blue-700 bg-blue-50 border border-blue-300 hover:bg-blue-100 px-1.5 py-0.5 rounded text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap">
                            <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4"/>
                            </svg>
                            追加
                          </button>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>{/* スクロール枠ここまで */}
        {totalPages > 1 && (
          <div className="px-4 py-3 border-t border-slate-100 flex items-center justify-between gap-4 flex-wrap bg-white rounded-b-xl">
            <span className="text-xs text-slate-400">{filtered.length.toLocaleString()}件中 {(page-1)*PAGE+1}–{Math.min(page*PAGE,filtered.length)}件</span>
            <div className="flex gap-1">
              <button onClick={()=>setPage(1)} disabled={page===1} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">«</button>
              <button onClick={()=>setPage(p=>p-1)} disabled={page===1} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">‹</button>
              <span className="px-3 py-1 text-xs text-slate-600">{page} / {totalPages}</span>
              <button onClick={()=>setPage(p=>p+1)} disabled={page===totalPages} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">›</button>
              <button onClick={()=>setPage(totalPages)} disabled={page===totalPages} className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">»</button>
            </div>
          </div>
        )}
      </div>

      {/* 現在リスト表示モーダル */}
      {listModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl flex flex-col max-h-[85vh]">
            <div className="flex items-center justify-between px-6 pt-5 pb-4 border-b border-slate-100 shrink-0">
              <div>
                <h2 className="text-base font-bold text-slate-800">📋 現在のリスト — 一致レコード</h2>
                <p className="text-xs text-slate-400 mt-0.5">{listModal.name}</p>
              </div>
              <button onClick={() => setListModal(null)} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
            </div>
            <div className="overflow-y-auto flex-1 px-6 py-4">
              {listModal.matched.length === 0 ? (
                <p className="text-slate-400 text-sm text-center py-8">一致するレコードが見つかりませんでした。</p>
              ) : (
                <div className="space-y-3">
                  {listModal.matched.map(r => {
                    const sc = STATUS_CFG[r.status] ?? {};
                    return (
                      <div key={r.id} className={`rounded-xl border p-4 ${sc.row || "bg-white"} ${sc.border || "border-slate-200"}`}>
                        <div className="flex flex-wrap items-center gap-2 mb-2">
                          <span className="font-semibold text-slate-800 text-sm">{r.companyName}</span>
                          <StatusBadge status={r.status}/>
                          {r.leadSource && <LeadSourceBadge source={r.leadSource}/>}
                        </div>
                        <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-1 text-xs text-slate-600">
                          {r.phone       && <span>📞 {r.phone}</span>}
                          {r.assignee    && <span>👤 {r.assignee}</span>}
                          {r.lastCallDate && <span>📅 架電日: {fmtDate(normDate(r.lastCallDate))}</span>}
                          {r.nextCallDate && <span>🔔 次回: {fmtDate(normDate(r.nextCallDate))}</span>}
                          {r.storeCount   && <span>🏪 {r.storeCount}店舗</span>}
                          {r.industry     && <span>🏭 {r.industry}</span>}
                        </div>
                        {r.memo && (
                          <p className="mt-2 text-xs text-slate-500 bg-slate-50 rounded-lg px-3 py-2 whitespace-pre-wrap">{r.memo}</p>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
            <div className="px-6 py-4 border-t border-slate-100 flex justify-end shrink-0">
              <button onClick={() => setListModal(null)}
                className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">
                閉じる
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 重複クレンジング */}
      {showDupe && (
        <DuplicateModal
          records={pastMgmt}
          onClean={ids => { const del = new Set(ids); setPastMgmt(p => p.filter(r => !del.has(r.id))); }}
          onClose={() => setShowDupe(false)}
          sortFn={rs => [...rs].sort((a, b) => {
            const ad = normDate(a.targetDate) || "";
            const bd = normDate(b.targetDate) || "";
            return bd.localeCompare(ad);
          })}
          renderExtra={r => (
            <>
              {r.leadAddedDate && <span className="text-xs text-slate-400 shrink-0">追加日: {fmtDate(normDate(r.leadAddedDate))}</span>}
              {r.targetDate    && <span className="text-xs text-blue-500 shrink-0 font-medium">完了予定日: {fmtDate(normDate(r.targetDate))}</span>}
            </>
          )}
        />
      )}
    </div>
  );
}

// ── HelpModal ──────────────────────────────────────────────────────────────────
function HelpModal({ onClose }) {
  const sections = [
    { icon:"📥", title:"CSVインポート", body:`・「自分の営業リスト」：企業名・電話・状況などを自動マッピング。Excel(.xlsx)にも対応。\n・「MiiTel架電ログ」：ISメンバー10名に自動絞り込み。\n・「過去商談リスト」：過去の商談データをインポートし、現在のリストと自動照合します。` },
    { icon:"📞", title:"MiiTel架電ログの最新の取り込み仕様", body:`MiiTel（ユーザー名・取引先会社名）のログを取り込むと、企業データベースを自動整理します。\n・未登録の企業 → 新規リードとして自動追加し、架電したオペレーター名を「追加者」列に記録。\n・既登録の企業 → 今回架電したオペレーター名を「別担当者」列に記録し、通話日付を「最新架電日（架電日）」に更新。\n・ISメンバー10名以外のログは自動で除外されます。\n・「追加者」「別担当者」列は列設定メニューから表示切り替えできます。` },
    { icon:"📜", title:"担当者の定義と各種取り込みマッピング", body:`【担当者】列は社内メンバー（ISスタッフ）の名前を記録します（旧「商談所有者」から名称統一。相手企業の担当者項目は廃止）。\n・MiiTel取込：既登録企業は今回の架電オペレーターを「担当者」に更新、未登録企業は新規追加して「追加者」に記録します。\n・過去商談取込：過去履歴の自社スタッフ名を「担当者」にマッピングして保存します。\n・編集モーダル最下部の「過去の商談・架電履歴」では、当時の自社スタッフを「当時の自社『担当者』」として表示します。` },
    { icon:"📊", title:"レポート分析画面の集計定義", body:`「📊 レポート」タブで活動成果を可視化します。\n・IS担当者別マトリクス：担当者ごとの総件数・アポ・コネクト・受付断り・不在/不通を集計。アポ獲得率＝アポ獲得商談数÷担当コネクト数（分母0は0%）。獲得率/アポ数/件数で並べ替え可能。\n・ステータス別パイプライン：全データの状況別件数と割合を横棒グラフで表示し、リードの滞留箇所を可視化。\n・業種別ランキング：アポ＋コネクト数の多い業種トップ5を自動抽出。\n・店舗規模別：1〜9 / 10〜99 / 100店舗以上（仮分析含む）の規模別にアプローチ件数とアポ獲得率を集計。` },
    { icon:"🔄", title:"重複クレンジング", body:`同一企業名のレコードをグループ化し、削除対象をチェックボックスで選択して一括削除できます。ステータス優先度順にソートされ、未架電・並が削除候補に自動選択されます。` },
    { icon:"⚙️", title:"設定（ロゴ・ファビコン・バックアップ）", body:`・ロゴ・ファビコン：PNG/JPEG をアップロードすると白背景を自動透過し、トリミングできます。\n・自動バックアップ：指定時刻になると通知バナーが表示され、CSV をダウンロードできます。` },
    { icon:"💾", title:"データのバックアップと復元", body:`⚙️設定モーダル内の「💾 データのバックアップと復元」から操作します。\n・「PCへバックアップファイルを保存」：営業リストと過去商談リストを1つのJSONファイルとしてPCに保存（ブラウザ容量を消費しません）。\n・「バックアップファイルからデータを復元」：保存したJSONを読み込んで現在のデータを上書き復元します。\n・「直前のインポート前の状態に戻す（1世代ロールバック）」：CSV取込でデータがおかしくなった時、ワンクリックで取込直前の状態に戻せます（取込のたびに自動退避）。` },
    { icon:"🏢", title:"店舗数未記入データの自動分析と（仮）表示", body:`店舗数が未記入の企業は、企業名をもとに自動で推測値を表示します。\n① 同名・系列社名で店舗数が入力済みのレコードがあれば、その数値を採用。\n② データベース内に全く同じ企業名が複数登録されている場合（多店舗チェーン等）、その重複件数を推測値とします。\n・手動入力の確定値はそのまま数値表示、自動推測値は「5（仮）」のようにグレーの「（仮）」付きで表示されます。\n・編集モーダルで正しい数値を入力して保存すると、確定値として上書きされ「（仮）」が外れます。` },
    { icon:"✏️", title:"インライン編集", body:`テーブルのセルをクリックすると直接編集できます。状況はセレクト、メモはテキストエリア、架電日は本日をデフォルトで表示します。Enterで保存、Escapeでキャンセル。` },
    { icon:"🏢", title:"エンタープライズ管理タブ", body:`大手・多店舗企業の商談を管理する専用タブです。\n・「Excelインポート（大手シート）」で、大手シート（状況・社内担当者・企業ブランド名・法人名・業種・店舗数・商談名・GMO営業フェーズ・最終更新日など）を取り込みます（インポートで全置換）。\n・通常リストと同様に、列設定（表示/並べ替え）・ソート・検索・状況フィルター・セルのインライン編集に対応。\n・データは営業リストとは独立して保存されます。` },
    { icon:"🛍️", title:"受注案件管理の使い方", body:`受注した案件を管理する専用タブ（営業リストとは独立保存）です。\n・項目：受注日・企業名・プラン/商材・受注金額・担当者・入金ステータス（未入金/一部入金/入金済/要確認）・稼働開始日・受注メモ。\n・「受注を追加」で手動登録、またはCSVインポートの「🛍️ 受注案件リストを取り込む」で一括取込（企業名・受注日・商材・金額・担当者を自動マッピング、同名は上書き）。\n・営業リストで状況が「成約/当社契約」または受注登録済みの企業には、リスト・編集画面に「🛍️受注済案件」バッジを表示。\n・編集モーダルの「🛍️ 受注案件へコピー」で、企業名・担当者を引き継いで受注案件に雛形登録できます。\n・レポート分析に当月受注総額・件数、IS担当者別の受注貢献ランキングを表示します。` },
  ];
  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl flex flex-col max-h-[88vh]">
        <div className="flex items-center justify-between px-6 pt-5 pb-4 border-b border-slate-100 shrink-0">
          <h2 className="text-base font-bold text-slate-800">📖 TEPPOU 使い方マニュアル</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
        </div>
        <div className="overflow-y-auto flex-1 px-6 py-4 space-y-4">
          {sections.map(s => (
            <div key={s.title} className="bg-slate-50 rounded-xl p-4">
              <p className="text-sm font-bold text-slate-800 mb-2">{s.icon} {s.title}</p>
              <p className="text-xs text-slate-600 whitespace-pre-line leading-relaxed">{s.body}</p>
            </div>
          ))}
        </div>
        <div className="px-6 py-4 border-t border-slate-100 flex justify-end shrink-0">
          <button onClick={onClose}
            className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-medium rounded-lg transition-colors">
            閉じる
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Main App ───────────────────────────────────────────────────────────────────
export default function App() {
  // ── API 同期用 refs ──────────────────────────────────────────────────────────
  const apiLoadedRef      = useRef(false);   // 初回 API 取得完了フラグ
  const lastWriteRef      = useRef(0);       // 最終書き込み時刻 (ms)
  const [syncing, setSyncing] = useState(false); // 手動更新中スピナー

  // ── IP チェック: null=確認中, true=許可, false=拒否 ──────────────────────────
  const [ipStatus,       setIpStatus]       = useState(null);
  const [myIp,           setMyIp]           = useState("");
  useEffect(() => {
    fetchMyIP()
      .then(ip => { setMyIp(ip); setIpStatus(ALLOWED_IPS.has(ip)); })
      .catch(() => setIpStatus(true)); // IP取得失敗時は通過（オフライン環境等）
  }, []);

  const [loggedIn,       setLoggedIn]       = useState(() => sessionStorage.getItem("teppou_auth")==="1");
  const [records,        setRecords]        = useState([]);
  const [pastDeals,      setPastDeals]      = useState([]);   // 過去商談プル照合用
  const [pastMgmt,       setPastMgmt]       = useState([]);   // 過去商談管理（再アプローチ）
  const [enterprise,     setEnterprise]     = useState([]);   // エンタープライズ管理（大手）
  const [orders,         setOrders]         = useState([]);   // 受注案件管理
  const [settings,       setSettings]       = useState({ logo:null, favicon:null, backupTimes:["10:00","14:00","18:00"] });
  const [showHelp,       setShowHelp]       = useState(false);
  const [search,         setSearch]         = useState("");
  const [statusFilterSet, setStatusFilterSet] = useState(() => new Set(Object.keys(STATUS_CFG)));
  const [assigneeFilter, setAssigneeFilter] = useState("all");
  const [excludeTodayCalled, setExcludeTodayCalled] = useState(false); // 今日架電したリストを除外
  const [leadSourceFilter, setLeadSourceFilter] = useState("all"); // ソース絞り込み
  const [alertFilter,    setAlertFilter]    = useState(null); // null | "recall" | "next"（アラート絞り込み）
  // UI 状態をまとめて localStorage から復元
  const savedUI = (() => { try { return JSON.parse(localStorage.getItem(UI_KEY) || "{}"); } catch { return {}; } })();
  const [visibleCols,    setVisibleCols]    = useState(Array.isArray(savedUI.visibleCols) && savedUI.visibleCols.length ? savedUI.visibleCols : DEFAULT_VISIBLE_COLS);
  const [showColDrop,    setShowColDrop]    = useState(false);
  const [page,           setPage]           = useState(1);
  const [showSettings,   setShowSettings]   = useState(false);
  const [showImport,     setShowImport]     = useState(false);
  const [showDupe,       setShowDupe]       = useState(false);
  const [showBulkStatus, setShowBulkStatus] = useState(false);
  const [hasAutoBackup,  setHasAutoBackup]  = useState(false);
  const [showNew,        setShowNew]        = useState(false);
  const [editRec,        setEditRec]        = useState(null);
  const [selected,       setSelected]       = useState(new Set());
  const [view,           setView]           = useState(["list","analysis","pastmgmt","enterprise","orders"].includes(savedUI.view) ? savedUI.view : "list");
  const [showPullList,   setShowPullList]   = useState(false);
  const [copiedId,       setCopiedId]       = useState(null);
  const [editingCell,    setEditingCell]    = useState(null); // { id, key }
  const [sortKey,        setSortKey]        = useState(savedUI.sortKey ?? null);
  const [sortDir,        setSortDir]        = useState(savedUI.sortDir || "asc");
  const ALL_STATUS_KEYS = Object.keys(STATUS_CFG);

  // STATUS_CFG に追加されたキー（未架電など）が filterSet に入っていない場合に自動補完
  useEffect(() => {
    setStatusFilterSet(prev => {
      const missing = ALL_STATUS_KEYS.filter(k => !prev.has(k));
      if (missing.length === 0) return prev;
      return new Set([...prev, ...missing]);
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const toggleStatus = key => setStatusFilterSet(prev => {
    const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n;
  });
  // アラートチップ: 押すと該当案件のみ表示。他の絞り込みが該当リードを隠さないよう解除する
  const focusAlert = kind => {
    const turningOn = alertFilter !== kind;
    setAlertFilter(turningOn ? kind : null);
    if (turningOn) {
      setStatusFilterSet(new Set(ALL_STATUS_KEYS)); // 全表示に戻す
      setExcludeTodayCalled(false);                 // 当日再架電は本日架電なので除外を解除
      setAssigneeFilter("all");
      setLeadSourceFilter("all");
      setSearch("");
    }
    setPage(1);
  };
  const [storageWarning, setStorageWarning] = useState(false);
  const colDropRef = useRef();

  // ── Persistence ──────────────────────────────────────────────────────────────

  // ローカルキャッシュ（IndexedDB）から読み込むヘルパー
  const loadFromLocal = useCallback(async () => {
    try {
      const recs = await idbGetAll();
      if (recs.length > 0) return migrateRecords(recs);
    } catch {}
    try {
      const s = localStorage.getItem(STORAGE_KEY);
      if (s) return migrateRecords(JSON.parse(s));
    } catch {}
    return null;
  }, []);

  // ローカルキャッシュに書き込む
  const saveToLocal = useCallback((recs) => {
    idbPutAll(recs).catch(() => {
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(recs)); } catch {}
    });
  }, []);

  // API からデータを取得してstateを更新
  const fetchAllFromAPI = useCallback(async ({ manual = false } = {}) => {
    if (!API_BASE) return;
    if (manual) setSyncing(true);
    const fetchStart = Date.now();
    try {
      const recs = await apiGet("records");
      // 巻き戻し防止: 書き込みから60秒以内はポーリング上書きしない
      if (!manual && fetchStart < lastWriteRef.current + 60_000) return;
      if (Array.isArray(recs)) {
        const migrated = migrateRecords(recs);
        setRecords(migrated);
        saveToLocal(migrated);
        setStorageWarning(false);
      }
    } catch (e) {
      if (e.status === 403) { setIpStatus(false); }
    } finally {
      if (manual) setSyncing(false);
    }
  }, [saveToLocal]);

  // API同期ヘルパー（mutations から呼ぶ）
  const syncToAPI = useCallback((newRecs) => {
    // IndexedDB/localStorage には常に保存（API未設定でも動く）
    saveToLocal(newRecs);
    // API同期はAPIが設定済み・初回ロード完了後のみ
    if (!API_BASE || !apiLoadedRef.current) return;
    lastWriteRef.current = Date.now();
    apiSet("records", newRecs).catch(e => {
      if (e.status === 403) setIpStatus(false);
    });
  }, [saveToLocal]);

  // 企業ロゴURLを「初回のみ」レコードに確定保存（以降は再解析・再取得なしの超軽量化）。
  // logoUrl が既にあるレコードは触らないので、確定後はこの処理は何もしない（無限ループなし）。
  useEffect(() => {
    if (!records.length) return;
    // ロゴ運用: 大手企業（KNOWN_LOGOS）には公式ファビコンのリンクを付与。
    // それ以外の自動生成ファビコンは白紙化。手動ロゴ（data:や任意URL）は尊重して維持。
    let changed = false;
    const next = records.map(r => {
      const known = matchKnownLogo(r.companyName);
      if (known) {
        if (r.logoUrl === known) return r;                          // 付与済み → 変更なし
        if (!r.logoUrl || isAutoFaviconUrl(r.logoUrl)) {            // 未設定 or 旧自動 → 付与
          changed = true; return { ...r, logoUrl: known };
        }
        return r;                                                   // 手動ロゴあり → 尊重
      }
      if (isAutoFaviconUrl(r.logoUrl)) { changed = true; return { ...r, logoUrl: "" }; } // 大手以外の自動は白紙化
      return r;
    });
    if (changed) { setRecords(next); syncToAPI(next); }
  }, [records, syncToAPI]);

  // 初回ロード
  useEffect(() => {
    // settings はローカルのみ
    try { const s = localStorage.getItem(SETTINGS_KEY);   if (s) setSettings(JSON.parse(s));  } catch {}
    // pastDeals: IndexedDB(kv) → なければ localStorage から移行
    idbKvGet("pastDeals").then(d => {
      if (Array.isArray(d) && d.length) setPastDeals(d);
      else {
        try { const s = localStorage.getItem(PAST_DEALS_KEY); if (s) { const p = JSON.parse(s); setPastDeals(p); idbKvSet("pastDeals", p).catch(()=>{}); localStorage.removeItem(PAST_DEALS_KEY); } } catch {}
      }
    }).catch(() => { try { const s = localStorage.getItem(PAST_DEALS_KEY); if (s) setPastDeals(JSON.parse(s)); } catch {} });
    // エンタープライズ管理（大手）
    idbKvGet("enterprise").then(d => { if (Array.isArray(d) && d.length) setEnterprise(d); }).catch(()=>{});
    // 受注案件管理
    idbKvGet("orders").then(d => { if (Array.isArray(d) && d.length) setOrders(d); }).catch(()=>{
      try { const s = localStorage.getItem("sales_mgr_orders"); if (s) setOrders(JSON.parse(s)); } catch {}
    });
    // 過去商談管理: 現IDB → 旧分離IDB → localStorage の順で復旧
    (async () => {
      try {
        let d = await idbPastGetAll();
        if (!d || d.length === 0) {
          // 旧・分離DB から復旧
          const legacy = await idbLegacyPastGetAll();
          if (legacy && legacy.length > 0) { d = legacy; idbPastPutAll(legacy).catch(()=>{}); }
        }
        if (!d || d.length === 0) {
          // localStorage から復旧
          try { const s = localStorage.getItem(PAST_MGMT_KEY); if (s) { d = JSON.parse(s); idbPastPutAll(d).catch(()=>{}); } } catch {}
        }
        if (d && d.length > 0) setPastMgmt(d);
      } catch {
        try { const s = localStorage.getItem(PAST_MGMT_KEY); if (s) setPastMgmt(JSON.parse(s)); } catch {}
      }
    })();

    if (API_BASE) {
      // API から取得 → なければローカルキャッシュ → なければ API に移行
      apiGet("records").then(async recs => {
        if (Array.isArray(recs) && recs.length > 0) {
          const migrated = migrateRecords(recs);
          setRecords(migrated);
          saveToLocal(migrated);
        } else {
          // API が空: ローカルキャッシュを API にマイグレーション
          const local = await loadFromLocal();
          if (local && local.length > 0) {
            setRecords(local);
            apiSet("records", local).catch(console.error);
          }
        }
        apiLoadedRef.current = true;
      }).catch(async (e) => {
        if (e.status === 403) { setIpStatus(false); return; }
        // API エラー: ローカルキャッシュにフォールバック
        const local = await loadFromLocal();
        if (local) setRecords(local);
        apiLoadedRef.current = true;
      });
    } else {
      // API_BASE 未設定（ローカル開発）: IndexedDB のみ
      loadFromLocal().then(local => { if (local) setRecords(local); });
      apiLoadedRef.current = true;
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // クリック時に更新（30秒クールダウン）
  const lastFetchTimeRef = useRef(0);
  const handleAppClick = useCallback(() => {
    if (!API_BASE) return;
    const now = Date.now();
    if (now - lastFetchTimeRef.current < CLICK_REFRESH_COOLDOWN) return;
    lastFetchTimeRef.current = now;
    fetchAllFromAPI();
  }, [fetchAllFromAPI]);

  useEffect(() => { try { localStorage.setItem(SETTINGS_KEY,   JSON.stringify(settings));  } catch {} }, [settings]);
  useEffect(() => { idbKvSet("pastDeals", pastDeals).catch(() => { try { localStorage.setItem(PAST_DEALS_KEY, JSON.stringify(pastDeals)); } catch {} }); }, [pastDeals]);
  useEffect(() => { idbKvSet("enterprise", enterprise).catch(()=>{}); }, [enterprise]);
  useEffect(() => { idbKvSet("orders", orders).catch(() => { try { localStorage.setItem("sales_mgr_orders", JSON.stringify(orders)); } catch {} }); }, [orders]);
  useEffect(() => { idbPastPutAll(pastMgmt).catch(() => { try { localStorage.setItem(PAST_MGMT_KEY, JSON.stringify(pastMgmt)); } catch {} }); }, [pastMgmt]);
  // UI 状態（列設定・ビュー・ソート）を保存
  useEffect(() => {
    try { localStorage.setItem(UI_KEY, JSON.stringify({ visibleCols, view, sortKey, sortDir })); } catch {}
  }, [visibleCols, view, sortKey, sortDir]);

  // ── Favicon ───────────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!settings.favicon) return;
    let link = document.querySelector("link[rel~='icon']");
    if (!link) { link = document.createElement("link"); link.rel = "icon"; document.head.appendChild(link); }
    link.type = "image/png";
    link.setAttribute("sizes", "512x512");
    link.href = settings.favicon;
  }, [settings.favicon]);

  // ── 自動バックアップ ───────────────────────────────────────────────────────────
  const [backupNotify,    setBackupNotify]    = useState(null); // "HH:MM" or null
  const lastBackupDoneRef = useRef({});  // { "YYYY-MM-DD_HH:MM": true }

  useEffect(() => {
    const times = settings.backupTimes ?? [];
    if (!times.length) return;
    const check = () => {
      const now  = new Date();
      const hhmm = `${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;
      const key  = `${getToday()}_${hhmm}`;
      if (times.includes(hhmm) && !lastBackupDoneRef.current[key]) {
        lastBackupDoneRef.current[key] = true;
        setBackupNotify(hhmm);
      }
    };
    const id = setInterval(check, 60_000);
    check();
    return () => clearInterval(id);
  }, [settings.backupTimes]);

  const doBackupDownload = (label = "") => {
    const ts  = getToday().replace(/-/g,"");
    const hm  = label.replace(":","") || new Date().toTimeString().slice(0,5).replace(":","");
    const csv = generateBackupCSV(records);
    triggerCSVDownload(csv, `TEPPOU_backup_${ts}_${hm}.csv`);
    setBackupNotify(null);
  };

  // ── Close column dropdown on outside click ────────────────────────────────────
  useEffect(() => {
    if (!showColDrop) return;
    const fn = e => { if (colDropRef.current && !colDropRef.current.contains(e.target)) setShowColDrop(false); };
    document.addEventListener("mousedown", fn);
    return () => document.removeEventListener("mousedown", fn);
  }, [showColDrop]);

  const today = getToday();
  const soon = (() => { const d = new Date(); d.setDate(d.getDate()+3); return d.toISOString().slice(0,10); })();
  // アラート判定（フィルターと表示の両方で使用）
  const isRecallAlert = r => normDate(r.lastCallDate) === today && RECALL_REASONS.includes(r.absenceReason||"");
  const isNextAlert   = r => r.nextCallDate && normDate(r.nextCallDate) <= soon && !DONE_STATUSES.includes(r.status);

  // ── Derived data ──────────────────────────────────────────────────────────────
  const filtered = records.filter(r => {
    // statusFilterSet が空 = フィルターなし（全表示）。一部選択時のみ絞り込む
    if (statusFilterSet.size > 0 && statusFilterSet.size < ALL_STATUS_KEYS.length && !statusFilterSet.has(r.status)) return false;
    if (assigneeFilter !== "all" && r.assignee !== assigneeFilter) return false;
    if (excludeTodayCalled && normDate(r.lastCallDate) === today) return false;
    if (leadSourceFilter !== "all" && (r.leadSource||"") !== leadSourceFilter) return false;
    if (alertFilter === "recall" && !isRecallAlert(r)) return false;
    if (alertFilter === "next"   && !isNextAlert(r))   return false;
    if (search) {
      const q = search;
      if (!(r.companyName||"").includes(q) && !(r.phone||"").includes(q) &&
          !(r.assignee||"").includes(q)     && !(r.memo||"").includes(q)  &&
          !(r.email||"").includes(q)) return false;
    }
    return true;
  });

  const sortedFiltered = sortKey
    ? [...filtered].sort((a, b) => {
        let va = a[sortKey] ?? "", vb = b[sortKey] ?? "";
        let cmp;
        if (sortKey === "storeCount") {
          cmp = (parseInt(String(va).replace(/,/g,""))||0) - (parseInt(String(vb).replace(/,/g,""))||0);
        } else {
          if (sortKey === "lastCallDate" || sortKey === "nextCallDate") { va = normDate(va); vb = normDate(vb); }
          cmp = String(va).localeCompare(String(vb), "ja");
        }
        const dirCmp = sortDir === "asc" ? cmp : -cmp;
        // 値が同じ場合は id で固定（行が入れ替わらないように）
        return dirCmp !== 0 ? dirCmp : String(a.id).localeCompare(String(b.id));
      })
    : filtered;

  // 表示順を凍結: 編集/コピーで値が変わっても並びは即時更新せず、
  // タブ切替・フィルター/ソート/検索変更・件数増減のタイミングで再計算する
  const [frozenIds, setFrozenIds] = useState(null);
  useEffect(() => {
    setFrozenIds(sortedFiltered.map(r => r.id));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [view, sortKey, sortDir, search, assigneeFilter, excludeTodayCalled, leadSourceFilter, alertFilter, statusFilterSet, records.length]);

  const recById = new Map(records.map(r => [r.id, r]));
  // 店舗数分析インデックス（一度だけ構築）
  const storeIndex = useMemo(() => buildStoreIndex(records, [...pastDeals, ...pastMgmt]), [records, pastDeals, pastMgmt]);
  const displayList = frozenIds
    ? frozenIds.map(id => recById.get(id)).filter(Boolean)
    : sortedFiltered;

  const totalPages = Math.max(1, Math.ceil(displayList.length / PAGE_SIZE));
  const paginated  = displayList.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);

  const statsMap = {};
  records.forEach(r => { statsMap[r.status] = (statsMap[r.status]||0) + 1; });
  const stats = Object.entries(statsMap)
    .map(([s,c]) => ({ status:s, count:c, ...(STATUS_CFG[s]??{}) }))
    .sort((a,b) => statusOrderIdx(a.status) - statusOrderIdx(b.status));

  const alerts       = records.filter(isNextAlert);
  // 当日帰社の可能性がある不在（本日架電 × 該当不在理由）→ 再架電アラート
  const recallAlerts = records.filter(isRecallAlert);
  const assignees = [...new Set(records.map(r => r.assignee).filter(Boolean))];
  // visibleCols の並び順で列を表示（未知キーは除外）
  const visibleDefs = visibleCols.map(k => ALL_COLUMNS.find(c => c.key === k)).filter(Boolean);

  // ── Mutations ─────────────────────────────────────────────────────────────────
  // 過去商談の追加（同一企業名は上書き）
  // ── バックアップ・復元 ──────────────────────────────────────────────────────
  // 最新stateをrefにミラー（スナップショット用）
  const stateRef = useRef({ records, pastDeals, pastMgmt });
  useEffect(() => { stateRef.current = { records, pastDeals, pastMgmt }; }, [records, pastDeals, pastMgmt]);

  // 初回: 自動退避の有無を確認
  useEffect(() => { idbKvGet("sales_mgr_auto_backup").then(b => setHasAutoBackup(!!b)); }, []);

  // インポート直前に1世代自動退避（IndexedDBへ。localStorage圧迫なし）
  const snapshotForRollback = useCallback(() => {
    const snap = { ...stateRef.current, savedAt: nowIso() };
    idbKvSet("sales_mgr_auto_backup", snap).then(() => setHasAutoBackup(true)).catch(()=>{});
  }, []);

  // PCへバックアップファイル保存（ブラウザ容量を消費しない）
  const exportBackup = useCallback(() => {
    const { records:r, pastDeals:pd, pastMgmt:pm } = stateRef.current;
    const payload = { app:"TEPPOU", version:1, exportedAt:nowIso(), records:r, pastDeals:pd, pastMgmt:pm };
    const blob = new Blob([JSON.stringify(payload)], { type:"application/json" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url;
    a.download = `teppou_backup_${getToday().replace(/-/g,"")}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }, []);

  // バックアップファイルから復元
  const importBackup = useCallback((file) => {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const data = JSON.parse(ev.target.result);
        const rCnt  = Array.isArray(data.records)   ? data.records.length   : 0;
        const pmCnt = Array.isArray(data.pastMgmt)  ? data.pastMgmt.length  : 0;
        if (!window.confirm(`バックアップから復元します。\n\n営業リスト: ${rCnt.toLocaleString()}件\n過去商談: ${pmCnt.toLocaleString()}件\n\n現在のデータは上書きされます。よろしいですか？`)) return;
        snapshotForRollback(); // 復元前に現状を自動退避
        if (Array.isArray(data.records))  { setRecords(data.records);   syncToAPI(data.records); }
        if (Array.isArray(data.pastDeals)) setPastDeals(data.pastDeals);
        if (Array.isArray(data.pastMgmt))  setPastMgmt(data.pastMgmt);
        window.alert("✅ 復元が完了しました。");
      } catch (e) {
        window.alert("❌ ファイルの読み込みに失敗しました。正しいバックアップファイルを選択してください。");
      }
    };
    reader.readAsText(file, "UTF-8");
  }, [snapshotForRollback, syncToAPI]);

  // 1世代ロールバック
  const rollbackAutoBackup = useCallback(async () => {
    const snap = await idbKvGet("sales_mgr_auto_backup");
    if (!snap) { window.alert("自動退避データがありません。"); return; }
    const when = snap.savedAt ? new Date(snap.savedAt).toLocaleString("ja-JP") : "";
    if (!window.confirm(`インポート直前の状態に戻します。\n（退避時刻: ${when}）\n\n現在のデータは破棄されます。よろしいですか？`)) return;
    if (Array.isArray(snap.records))  { setRecords(snap.records);   syncToAPI(snap.records); }
    if (Array.isArray(snap.pastDeals)) setPastDeals(snap.pastDeals);
    if (Array.isArray(snap.pastMgmt))  setPastMgmt(snap.pastMgmt);
    window.alert("✅ 直前の状態に戻しました。");
  }, [syncToAPI]);

  const addPastDeals = useCallback(newDeals => {
    snapshotForRollback(); // インポート前に自動退避
    setPastDeals(prev => {
      const map = {};
      prev.forEach(d => { map[normName(d.companyName)] = d; });
      newDeals.forEach(d => { map[normName(d.companyName)] = { ...map[normName(d.companyName)], ...d }; });
      return Object.values(map);
    });
  }, [snapshotForRollback]);

  // 受注CSVインポート（企業名で上書き更新）
  const addOrders = useCallback((newOrders) => {
    snapshotForRollback();
    setOrders(prev => {
      const map = {};
      prev.forEach(o => { map[normName(o.companyName)] = o; });
      newOrders.forEach(o => { const k = normName(o.companyName); map[k] = { ...map[k], ...o, id: map[k]?.id || o.id }; });
      return Object.values(map);
    });
  }, [snapshotForRollback]);

  // 営業リスト→受注案件へワンクリックコピー
  const addOrderFromRecord = useCallback((rec) => {
    setOrders(prev => [{
      id: genId(),
      orderDate:   getToday(),
      companyName: rec.companyName,
      plan:        "スタンダードプラン",
      amount:      "",
      assignee:    rec.createdBy || rec.assignee || "",
      payment:     "未入金",
      startDate:   "",
      memo:        rec.memo || "",
      updatedAt:   nowIso(),
    }, ...prev]);
  }, []);

  // 受注済み企業名セット（バッジ用）
  const orderedNames = useMemo(() => new Set(orders.map(o => normName(o.companyName))), [orders]);
  const isOrdered = useCallback((rec) => {
    if (rec?.status === "成約" || rec?.status === "8.当社契約") return true;
    return orderedNames.has(normName(rec?.companyName || ""));
  }, [orderedNames]);

  const addPastDealToList = useCallback((deal) => {
    setRecords(p => {
      // 過去商談の状況を引き継ぐ（STATUS_CFGに存在する値のみ。なければ未架電）
      const carriedStatus = (deal.status && STATUS_CFG[deal.status]) ? deal.status
                          : (deal.pastStatus && STATUS_CFG[deal.pastStatus]) ? deal.pastStatus
                          : "未架電";
      const next = [...p, {
        id: genId(),
        companyName:   deal.companyName,
        phone:         deal.phone || "",
        status:        carriedStatus,
        assignee:      "",   // 担当者は引き継がない
        createdBy:     "",
        storeCount:    deal.storeCount || "",
        lastCallDate:  normDate(deal.lastCallDate) || "",
        nextCallDate:  "",
        memo:          deal.memo || "",
        leadSource:    "過去商談(他)",
        leadAddedDate: getToday(),
        importedAt: nowIso(), updatedAt: nowIso(), source:"past-add",
      }];
      syncToAPI(next);
      return next;
    });
  }, [syncToAPI]);

  const importMetelMerge = useCallback((parsed) => {
    snapshotForRollback(); // インポート前に自動退避
    let added = 0, updated = 0;
    setRecords(prev => {
      const byName = new Map(prev.map(r => [normName(r.companyName), r]));
      parsed.forEach(p => {
        const key = normName(p.company);
        const existing = byName.get(key);
        if (existing) {
          // 既登録: 担当者(assignee)・取込月を更新（架電日は維持）
          byName.set(key, {
            ...existing,
            assignee: p.operator,       // 担当者＝今回の架電オペレーター
            importMonth: p.importMonth || existing.importMonth,
            status: p.status || existing.status,
            updatedAt: nowIso(),
          });
        } else {
          // 未登録: 新規リードとして追加。追加者=オペレーター
          byName.set(key, {
            id: genId(),
            companyName: p.company,
            phone: p.phone || "",
            status: p.status || "未架電",
            assignee: "",
            createdBy: p.operator,      // 追加者
            importMonth: p.importMonth || "",
            lastCallDate: "",
            memo: p.memo || "",
            leadAddedDate: getToday(),
            source: "metel",
            importedAt: nowIso(), updatedAt: nowIso(),
          });
        }
      });
      const next = [...byName.values()];
      syncToAPI(next);
      return next;
    });
    // 件数集計（現在の records を基準）
    const existingNames = new Set(records.map(r => normName(r.companyName)));
    const seen = new Set();
    parsed.forEach(p => {
      const key = normName(p.company);
      if (existingNames.has(key)) { if (!seen.has(key)) updated++; }
      else if (!seen.has(key)) added++;
      seen.add(key);
    });
    setStatusFilterSet(new Set(Object.keys(STATUS_CFG)));
    setView("list");
    setPage(1);
    return { added, updated };
  }, [records, syncToAPI, snapshotForRollback]);

  // 過去商談の企業名照合ヘルパー
  const findPastDeal = useCallback((companyName) => {
    const n = normName(companyName);
    return pastDeals.find(d => normName(d.companyName) === n || normName(d.companyName).includes(n) || n.includes(normName(d.companyName)));
  }, [pastDeals]);

  const addRecords = useCallback(recs => {
    snapshotForRollback(); // インポート前に自動退避
    setRecords(p => { const next = [...p, ...recs]; syncToAPI(next); return next; });
    // インポート後: フィルター・検索・ページをすべてリセットして確実に表示
    setStatusFilterSet(new Set(Object.keys(STATUS_CFG)));
    setSearch("");
    setAssigneeFilter("all");
    setSortKey(null);
    setPage(1);
    setView("list");
  }, [syncToAPI, snapshotForRollback]);

  const saveRecord = useCallback(form => {
    const isEdit = records.some(r => r.id === form.id);
    if (isEdit) {
      setRecords(p => { const next = p.map(r => r.id===form.id ? clampNextCall({ ...r, ...form, updatedAt:nowIso() }) : r); syncToAPI(next); return next; });
    } else {
      setRecords(p => { const next = [...p, clampNextCall({ ...form, id:genId(), callCount:form.callCount||0, importedAt:nowIso(), updatedAt:nowIso(), source:"manual", leadAddedDate: form.leadAddedDate || getToday() })]; syncToAPI(next); return next; });
    }
  }, [records, syncToAPI]);

  const deleteRecord = useCallback(id => {
    if (!window.confirm("このレコードを削除しますか？")) return;
    setRecords(p => { const next = p.filter(r => r.id !== id); syncToAPI(next); return next; });
    setSelected(p => { const n = new Set(p); n.delete(id); return n; });
  }, [syncToAPI]);

  const deleteSelected = useCallback(() => {
    if (!selected.size) return;
    if (!window.confirm(`選択した ${selected.size} 件を削除しますか？`)) return;
    setRecords(p => { const next = p.filter(r => !selected.has(r.id)); syncToAPI(next); return next; });
    setSelected(new Set());
  }, [selected, syncToAPI]);

  const saveInlineValue = useCallback((id, key, raw) => {
    setEditingCell(null);
    const value = (key === "lastCallDate" || key === "nextCallDate") ? normDate(raw) : raw;
    setRecords(p => {
      const next = p.map(r => r.id === id ? clampNextCall({ ...r, [key]: value, updatedAt: nowIso() }) : r);
      syncToAPI(next);
      return next;
    });
  }, [syncToAPI]);

  const copyCompanyName = useCallback((text, id) => {
    navigator.clipboard.writeText(text).then(() => {
      setCopiedId(id);
      setTimeout(() => setCopiedId(null), 1500);
    });
    // 架電日を本日に自動更新
    setRecords(p => {
      const next = p.map(r => r.id === id ? { ...r, lastCallDate: getToday(), updatedAt: nowIso() } : r);
      syncToAPI(next);
      return next;
    });
  }, [syncToAPI]);

  // statusMap: Map(id → 新ステータス)。または (Set(ids), status) の旧形式も許容
  const bulkUpdateStatus = useCallback((statusMap, statusArg) => {
    setRecords(p => {
      const next = p.map(r => {
        if (statusMap instanceof Map) {
          return statusMap.has(r.id) ? { ...r, status: statusMap.get(r.id), updatedAt: nowIso() } : r;
        }
        return statusMap.has(r.id) ? { ...r, status: statusArg, updatedAt: nowIso() } : r;
      });
      syncToAPI(next);
      return next;
    });
  }, [syncToAPI]);

  const cleanDuplicates = useCallback(ids => {
    const del = new Set(ids);
    setRecords(p => { const next = p.filter(r => !del.has(r.id)); syncToAPI(next); return next; });
  }, [syncToAPI]);

  const toggleSelect = useCallback((id, checked) => {
    setSelected(p => { const n = new Set(p); checked ? n.add(id) : n.delete(id); return n; });
  }, []);

  const togglePageSelect = useCallback(checked => {
    setSelected(p => {
      const n = new Set(p);
      paginated.forEach(r => checked ? n.add(r.id) : n.delete(r.id));
      return n;
    });
  }, [paginated]);

  // ── Pagination helpers ────────────────────────────────────────────────────────
  const goPage = p => { setPage(Math.max(1, Math.min(totalPages, p))); };
  const pageNums = (() => {
    const result = [];
    for (let i = 1; i <= totalPages; i++) {
      if (i===1 || i===totalPages || Math.abs(i-page)<=2) result.push(i);
      else if (result[result.length-1] !== "…") result.push("…");
    }
    return result;
  })();

  // ─────────────────────────────────────────────────────────────────────────────
  if (ipStatus === null) return <IPCheckScreen />;
  if (ipStatus === false) return <AccessDenied ip={myIp} />;
  if (!loggedIn) return <LoginScreen onLogin={() => setLoggedIn(true)} logo={settings.logo} />;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800" onClick={handleAppClick}>

      {/* ── Header ── */}
      <header className="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-30">
        <div className="max-w-screen-2xl mx-auto px-4 py-3 flex items-center justify-between gap-3 flex-wrap">
          <div className="flex items-center gap-3">
            <AppIcon logo={settings.logo} size="sm" />
            <div>
              <div className="font-bold text-slate-800 leading-tight text-sm">TEPPOU</div>
              <div className="text-xs text-slate-400">GMOテック IS部門</div>
            </div>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-xs text-slate-500 hidden sm:block">
              総件数: <span className="font-bold text-blue-600">{records.length}</span>件
            </span>
            {API_BASE && (
              <button onClick={() => fetchAllFromAPI({ manual: true })}
                disabled={syncing}
                className="flex items-center gap-1 text-xs text-slate-600 border border-slate-200 px-3 py-1.5 rounded-lg hover:bg-slate-50 transition-colors disabled:opacity-50">
                <span className={syncing ? "animate-spin inline-block" : ""}>🔄</span>
                更新
              </button>
            )}
            <button onClick={() => setShowHelp(true)}
              className="flex items-center gap-1 text-xs text-slate-600 border border-slate-200 px-3 py-1.5 rounded-lg hover:bg-slate-50 transition-colors">
              📖 ヘルプ
            </button>
            <button onClick={() => setShowSettings(true)}
              className="flex items-center gap-1 text-xs text-slate-600 border border-slate-200 px-3 py-1.5 rounded-lg hover:bg-slate-50 transition-colors">
              ⚙️ 設定
            </button>
            <button onClick={() => { sessionStorage.removeItem("teppou_auth"); setLoggedIn(false); }}
              className="text-xs text-slate-400 hover:text-slate-700 border border-slate-200 px-3 py-1.5 rounded-lg hover:bg-slate-50 transition-colors">
              ログアウト
            </button>
          </div>
        </div>
      </header>

      <div className="max-w-screen-2xl mx-auto px-4 pt-3 pb-5 space-y-4">

        {/* ── Page tabs ── */}
        <div className="flex gap-1 bg-slate-100 rounded-xl p-1 w-fit">
          {[["list","📋 リスト"],["analysis","📊 レポート"],["pastmgmt","📂 過去商談"],["enterprise","🏢 エンプラ"],["orders","🛍️ 受注案件"]].map(([v,label])=>(
            <button key={v} onClick={()=>setView(v)}
              className={`px-4 py-1.5 text-xs font-semibold rounded-lg transition-colors
                ${view===v ? "bg-white text-blue-700 shadow-sm" : "text-slate-500 hover:text-slate-700"}`}>
              {label}
            </button>
          ))}
        </div>

        {/* ── Analysis view ── */}
        {view==="analysis" && <ReportView records={records} pastDeals={pastDeals} orders={orders} />}

        {/* ── 過去商談管理 ── */}
        {view==="pastmgmt" && <PastMgmtView pastMgmt={pastMgmt} setPastMgmt={setPastMgmt} records={records}
          onGoToList={name => { setView("list"); setSearch(name); setPage(1); }}
          onAddToList={addPastDealToList} onBeforeImport={snapshotForRollback} />}

        {/* ── エンタープライズ管理 ── */}
        {view==="enterprise" && <PastMgmtView pastMgmt={enterprise} setPastMgmt={setEnterprise} records={records}
          onGoToList={name => { setView("list"); setSearch(name); setPage(1); }}
          onAddToList={addPastDealToList} onBeforeImport={snapshotForRollback} />}

        {/* ── 受注案件管理 ── */}
        {view==="orders" && <OrderView orders={orders} setOrders={setOrders} members={IS_MEMBERS} />}

        {/* ── List view ── */}
        {view==="list" && <>

        {/* ── バックアップ通知バナー ── */}
        {backupNotify && (
          <div className="bg-blue-50 border border-blue-300 rounded-xl px-4 py-3 flex flex-wrap items-center gap-3">
            <span className="text-sm font-semibold text-blue-700 shrink-0">💾 {backupNotify} バックアップ</span>
            <span className="text-xs text-blue-600 flex-1">データのバックアップ時刻になりました（{records.length.toLocaleString()}件）</span>
            <button onClick={() => doBackupDownload(backupNotify)}
              className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg transition-colors shrink-0">
              CSVダウンロード
            </button>
            <button onClick={() => setBackupNotify(null)}
              className="px-3 py-1.5 text-xs text-blue-500 hover:text-blue-700 shrink-0">
              後で
            </button>
          </div>
        )}

        {/* ── Storage warning ── */}
        {storageWarning && (
          <div className="bg-orange-50 border border-orange-300 rounded-xl px-4 py-3 flex flex-wrap items-center gap-2">
            <span className="text-sm font-semibold text-orange-700 shrink-0">⚠️ 保存エラー</span>
            <span className="text-xs text-orange-600">
              ブラウザへのデータ保存に失敗しました。
              <strong>ページを閉じるとデータが消える可能性があります。</strong>
              プライベートブラウズモードを使用している場合はオフにしてください。
            </span>
          </div>
        )}

        {/* ── アラート（小さく件数表示・押すと該当案件をリスト表示） ── */}
        {(recallAlerts.length > 0 || alerts.length > 0) && (
          <div className="flex flex-wrap items-center gap-2">
            {recallAlerts.length > 0 && (
              <button onClick={() => focusAlert("recall")}
                title="押すと当日再架電の案件のみ表示"
                className={`flex items-center gap-1.5 text-xs font-semibold px-3 py-1.5 rounded-full border transition-colors
                  ${alertFilter==="recall" ? "bg-sky-600 text-white border-sky-600" : "bg-sky-50 text-sky-700 border-sky-300 hover:bg-sky-100"}`}>
                📞 当日再架電 {recallAlerts.length}件
              </button>
            )}
            {alerts.length > 0 && (
              <button onClick={() => focusAlert("next")}
                title="押すと次回架電日の案件のみ表示"
                className={`flex items-center gap-1.5 text-xs font-semibold px-3 py-1.5 rounded-full border transition-colors
                  ${alertFilter==="next" ? "bg-amber-600 text-white border-amber-600" : "bg-amber-50 text-amber-700 border-amber-300 hover:bg-amber-100"}`}>
                📅 次回架電 {alerts.length}件
              </button>
            )}
            {alertFilter && (
              <button onClick={() => { setAlertFilter(null); setPage(1); }}
                className="text-xs text-slate-400 hover:text-slate-700 underline">
                絞り込み解除
              </button>
            )}
          </div>
        )}

        {/* ── Stats bar (multi-select) ── */}
        {stats.length > 0 && (
          <div className="bg-white rounded-xl border border-slate-200 p-4">
            <div className="flex items-center justify-between mb-2">
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">
                ステータス別集計
                <span className="ml-2 font-normal text-slate-300">
                  （クリックで絞り込み）
                </span>
              </p>
              <div className="flex gap-1.5">
                <button
                  onClick={() => { setStatusFilterSet(new Set(ALL_STATUS_KEYS)); setPage(1); }}
                  className="px-2.5 py-1 text-xs border border-blue-300 rounded-lg text-blue-600 hover:bg-blue-50 transition-colors">
                  全表示
                </button>
                <button
                  onClick={() => { setStatusFilterSet(new Set()); setPage(1); }}
                  className="px-2.5 py-1 text-xs border border-slate-300 rounded-lg text-slate-400 hover:bg-slate-50 transition-colors">
                  全非表示
                </button>
                <button
                  onClick={() => { setStatusFilterSet(new Set(CALL_PRESET_STATUSES)); setPage(1); }}
                  title="コネクト（無）・並・優先・高確度に絞る"
                  className="px-2.5 py-1 text-xs border border-emerald-300 rounded-lg text-emerald-600 hover:bg-emerald-50 transition-colors font-semibold">
                  架電
                </button>
              </div>
            </div>
            <div className="flex flex-wrap gap-2">
              {stats.map(s => {
                const active = statusFilterSet.has(s.status);
                return (
                  <button key={s.status}
                    onClick={() => { toggleStatus(s.status); setPage(1); }}
                    className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg border text-xs font-medium transition-all
                      ${active
                        ? `${s.bg??"bg-gray-100"} ${s.text??"text-gray-600"} ${s.border??"border-gray-300"}`
                        : "bg-white text-slate-300 border-slate-200"}`}>
                    <span className={`w-2 h-2 rounded-full transition-colors ${active ? (s.dot??"bg-gray-400") : "bg-slate-200"}`} />
                    <span className={active ? "" : "line-through"}>{s.status}</span>
                    <span className={`font-bold ${active ? "" : "text-slate-200"}`}>{s.count}</span>
                  </button>
                );
              })}
            </div>
          </div>
        )}

        {/* ── Toolbar ── */}
        <div className="bg-white rounded-xl border border-slate-200 p-4 space-y-3">
          <div className="flex flex-wrap items-center gap-2">

            {/* CSV Import */}
            <button onClick={() => setShowImport(true)}
              className="flex items-center gap-1.5 bg-blue-50 hover:bg-blue-100 border border-blue-300 text-blue-700 px-3 py-2 rounded-lg text-xs font-medium transition-colors">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                  d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
              </svg>
              CSVインポート
            </button>

            {/* ステータス一括更新 */}
            <button onClick={() => setShowBulkStatus(true)}
              className="flex items-center gap-1.5 bg-teal-50 hover:bg-teal-100 border border-teal-300 text-teal-700 px-3 py-2 rounded-lg text-xs font-medium transition-colors">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/>
              </svg>
              ステータス一括更新
            </button>

            {/* Duplicate cleanse */}
            <button onClick={() => setShowDupe(true)}
              className="flex items-center gap-1.5 bg-amber-50 hover:bg-amber-100 border border-amber-300 text-amber-700 px-3 py-2 rounded-lg text-xs font-medium transition-colors">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                  d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" />
              </svg>
              重複クレンジング
            </button>

            {/* 手動バックアップ */}
            <button onClick={() => doBackupDownload()}
              disabled={records.length === 0}
              className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-300 text-slate-600 px-3 py-2 rounded-lg text-xs font-medium transition-colors disabled:opacity-40">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                  d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/>
              </svg>
              バックアップ
            </button>

            {/* Column settings dropdown */}
            <div className="relative" ref={colDropRef}>
              <button onClick={() => setShowColDrop(v => !v)}
                className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-300 text-slate-600 px-3 py-2 rounded-lg text-xs font-medium transition-colors">
                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 10h18M3 14h18M10 3v18M14 3v18" />
                </svg>
                列設定
                <svg className={`w-3 h-3 transition-transform ${showColDrop ? "rotate-180" : ""}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                </svg>
              </button>
              {showColDrop && (
                <div className="absolute left-0 top-full mt-1 bg-white rounded-xl border border-slate-200 shadow-xl z-20 p-3 w-64 max-h-96 overflow-y-auto">
                  <p className="text-xs font-semibold text-slate-400 mb-1 px-1">表示中の列（↑↓で並べ替え）</p>
                  {visibleCols.map((key, idx) => {
                    const col = ALL_COLUMNS.find(c => c.key === key);
                    if (!col) return null;
                    return (
                      <div key={key} className="flex items-center gap-1 px-2 py-1 hover:bg-slate-50 rounded-lg">
                        <div className="flex flex-col">
                          <button disabled={idx===0}
                            onClick={() => setVisibleCols(p => { const n=[...p]; [n[idx-1],n[idx]]=[n[idx],n[idx-1]]; return n; })}
                            className="text-slate-400 hover:text-blue-600 disabled:opacity-20 leading-none text-[10px]">▲</button>
                          <button disabled={idx===visibleCols.length-1}
                            onClick={() => setVisibleCols(p => { const n=[...p]; [n[idx+1],n[idx]]=[n[idx],n[idx+1]]; return n; })}
                            className="text-slate-400 hover:text-blue-600 disabled:opacity-20 leading-none text-[10px]">▼</button>
                        </div>
                        <span className="text-xs text-slate-700 flex-1 truncate">{col.label}</span>
                        {col.required
                          ? <span className="text-xs text-slate-300">必須</span>
                          : <button onClick={() => setVisibleCols(p => p.filter(k => k!==key))}
                              className="text-rose-400 hover:text-rose-600 text-xs">×</button>}
                      </div>
                    );
                  })}
                  {ALL_COLUMNS.some(c => !visibleCols.includes(c.key)) && (
                    <>
                      <p className="text-xs font-semibold text-slate-400 mt-2 mb-1 px-1 border-t border-slate-100 pt-2">非表示の列（＋で追加）</p>
                      {ALL_COLUMNS.filter(c => !visibleCols.includes(c.key)).map(col => (
                        <button key={col.key}
                          onClick={() => setVisibleCols(p => [...p, col.key])}
                          className="flex items-center gap-2 px-2 py-1 hover:bg-slate-50 rounded-lg w-full text-left">
                          <span className="text-blue-500 text-xs">＋</span>
                          <span className="text-xs text-slate-500">{col.label}</span>
                        </button>
                      ))}
                    </>
                  )}
                </div>
              )}
            </div>

            {/* New record */}
            <button onClick={() => setShowNew(true)}
              className="flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-xs font-semibold transition-colors">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
              </svg>
              新規追加
            </button>

            {/* Batch delete */}
            {selected.size > 0 && (
              <button onClick={deleteSelected}
                className="flex items-center gap-1.5 bg-red-50 hover:bg-red-100 border border-red-300 text-red-700 px-3 py-2 rounded-lg text-xs font-medium transition-colors">
                選択削除（{selected.size}件）
              </button>
            )}
          </div>

          {/* Search & filters */}
          <div className="flex flex-wrap gap-2">
            <div className="relative flex-1 min-w-48">
              <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
              </svg>
              <input type="text" value={search}
                onChange={e => { setSearch(e.target.value); setPage(1); }}
                placeholder="企業名・電話番号・担当者・メモで検索..."
                className="w-full pl-9 pr-9 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" />
              {search && (
                <button onClick={() => { setSearch(""); setPage(1); }} title="クリア"
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 w-5 h-5 flex items-center justify-center rounded-full bg-slate-200 hover:bg-slate-300 text-slate-600 text-xs transition-colors">
                  ×
                </button>
              )}
            </div>
            {/* ソース絞り込み */}
            <select value={leadSourceFilter} onChange={e => { setLeadSourceFilter(e.target.value); setPage(1); }}
              className="border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
              <option value="all">全ソース</option>
              {Object.keys(LEAD_SOURCE_CFG).map(s => <option key={s} value={s}>{s}</option>)}
            </select>
            {/* 今日架電を除外フィルター */}
            <button onClick={() => { setExcludeTodayCalled(v => !v); setPage(1); }}
              className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-medium border transition-colors whitespace-nowrap
                ${excludeTodayCalled ? "bg-amber-600 text-white border-amber-600" : "bg-white text-slate-600 border-slate-300 hover:bg-slate-50"}`}>
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M18.364 18.364A9 9 0 005.636 5.636m12.728 12.728A9 9 0 015.636 5.636m12.728 12.728L5.636 5.636"/>
              </svg>
              今日架電を除外
            </button>
          </div>
        </div>

        {/* ── Table ── */}
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="overflow-auto max-h-[calc(100vh-280px)]">
          <table className="w-full text-sm border-collapse table-fixed">
              <thead className="bg-slate-50 border-b border-slate-200 sticky top-0 z-10">
                <tr>
                  <th className="w-[44px] px-3 py-3 text-left bg-slate-50">
                    <input type="checkbox"
                      checked={paginated.length > 0 && paginated.every(r => selected.has(r.id))}
                      onChange={e => togglePageSelect(e.target.checked)}
                      className="rounded border-slate-300 text-blue-600" />
                  </th>
                  {visibleDefs.map(col => (
                    <th key={col.key}
                      onClick={() => {
                        if (sortKey === col.key) setSortDir(d => d==="asc" ? "desc" : "asc");
                        else { setSortKey(col.key); setSortDir("asc"); }
                        setPage(1);
                      }}
                      className={`${col.w||"w-[120px]"} px-3 py-3 text-left text-xs font-semibold text-slate-500 whitespace-nowrap cursor-pointer hover:bg-slate-100 transition-colors bg-slate-50`}>
                      <span className="flex items-center gap-1">
                        <span className="truncate">{col.label}</span>
                        {sortKey === col.key
                          ? <span className="text-blue-500 shrink-0">{sortDir==="asc" ? "▲" : "▼"}</span>
                          : <span className="text-slate-300 shrink-0">⇅</span>}
                      </span>
                    </th>
                  ))}
                  <th className="w-[80px] px-3 py-3 text-left text-xs font-semibold text-slate-500 bg-slate-50">操作</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {paginated.length === 0 ? (
                  <tr>
                    <td colSpan={visibleDefs.length + 2}
                      className="text-center py-16 text-slate-400 text-sm">
                      データがありません。CSVをインポートするか、新規追加してください。
                    </td>
                  </tr>
                ) : paginated.map(rec => {
                  const isToday  = normDate(rec.lastCallDate) === today;
                  const rowColor = selected.has(rec.id)
                    ? "bg-blue-100"
                    : isToday
                      ? "bg-yellow-100 border-l-4 border-yellow-400"
                      : (STATUS_CFG[rec.status]?.row ?? "");
                  return (
                  <tr key={rec.id} className={`transition-colors hover:brightness-95 ${rowColor}`}>
                    <td className="px-3 py-2.5">
                      <input type="checkbox" checked={selected.has(rec.id)}
                        onChange={e => toggleSelect(rec.id, e.target.checked)}
                        className="rounded border-slate-300 text-blue-600" />
                    </td>
                    {visibleDefs.map(col => {
                      const isEditing = editingCell?.id === rec.id && editingCell?.key === col.key && col.key !== "companyName";
                      const openEdit  = () => setEditingCell({ id: rec.id, key: col.key });
                      const save      = (val) => saveInlineValue(rec.id, col.key, val);
                      const cancel    = () => setEditingCell(null);
                      const inputCls  = "border border-blue-400 rounded px-1 py-0.5 text-xs focus:outline-none focus:ring-1 focus:ring-blue-400 bg-white";
                      const mkInput   = (type, extra = {}) => (
                        <input type={type} autoFocus defaultValue={
                            (type === "date") ? normDate(rec[col.key]) || "" : rec[col.key] ?? ""
                          }
                          className={`${inputCls} ${type === "date" ? "w-32" : type === "number" ? "w-20" : "w-36"}`}
                          onBlur={e    => save(e.target.value)}
                          onKeyDown={e => { if (e.key === "Enter") save(e.target.value); if (e.key === "Escape") cancel(); }}
                          {...extra}
                        />
                      );

                      // ── EDIT MODE ───────────────────────────────────────────
                      let editEl = null;
                      if (isEditing) {
                        if (col.key === "status") {
                          editEl = (
                            <select autoFocus defaultValue={rec.status}
                              className={`${inputCls} w-36`}
                              onChange={e => save(e.target.value)}
                              onBlur={cancel}
                              onKeyDown={e => e.key === "Escape" && cancel()}>
                              {Object.keys(STATUS_CFG).map(s => <option key={s} value={s}>{s}</option>)}
                            </select>
                          );
                        } else if (col.key === "leadSource") {
                          editEl = (
                            <select autoFocus defaultValue={rec.leadSource || ""}
                              className={`${inputCls} w-36`}
                              onChange={e => save(e.target.value)}
                              onBlur={cancel}
                              onKeyDown={e => e.key === "Escape" && cancel()}>
                              <option value="">— 未選択 —</option>
                              {Object.keys(LEAD_SOURCE_CFG).map(s => <option key={s} value={s}>{s}</option>)}
                            </select>
                          );
                        } else if (col.key === "absenceReason") {
                          editEl = (
                            <select autoFocus defaultValue={rec.absenceReason || ""}
                              className={`${inputCls} w-28`}
                              onChange={e => save(e.target.value)}
                              onBlur={cancel}
                              onKeyDown={e => e.key === "Escape" && cancel()}>
                              <option value="">— 未選択 —</option>
                              {Object.keys(ABSENCE_REASON_CFG).map(s => <option key={s} value={s}>{s}</option>)}
                            </select>
                          );
                        } else if (col.key === "importMonth") {
                          editEl = (
                            <input type="month" autoFocus defaultValue={rec.importMonth || ""}
                              className={`${inputCls} w-32`}
                              onBlur={e    => save(e.target.value)}
                              onKeyDown={e => { if (e.key === "Enter") save(e.target.value); if (e.key === "Escape") cancel(); }}
                            />
                          );
                        } else if (col.key === "lastCallDate" || col.key === "nextCallDate" || col.key === "leadAddedDate") {
                          // 架電日はデフォルトを「本日」、次回架電日・リード追加日は既存値
                          const dateDefault = col.key === "lastCallDate"
                            ? (normDate(rec[col.key]) || today)
                            : (normDate(rec[col.key]) || "");
                          editEl = (
                            <input type="date" autoFocus defaultValue={dateDefault}
                              className={`${inputCls} w-32`}
                              onBlur={e    => save(e.target.value)}
                              onKeyDown={e => { if (e.key === "Enter") save(e.target.value); if (e.key === "Escape") cancel(); }}
                            />
                          );
                        } else if (col.key === "storeCount") {
                          editEl = mkInput("number");
                        } else if (col.key === "memo") {
                          editEl = (
                            <textarea autoFocus defaultValue={rec.memo || ""}
                              rows={3}
                              className={`${inputCls} w-56 resize-none`}
                              onBlur={e    => save(e.target.value)}
                              onKeyDown={e => { if (e.key === "Escape") cancel(); if (e.key === "Enter" && e.ctrlKey) save(e.target.value); }}
                            />
                          );
                        } else {
                          editEl = mkInput("text");
                        }
                      }

                      // ── VIEW MODE ───────────────────────────────────────────
                      let viewEl = null;
                      if (!isEditing) {
                        const val = rec[col.key];
                        const empty = <span className="text-slate-300 text-xs">—</span>;
                        if (col.key === "companyName") {
                          const pd = findPastDeal(val);
                          viewEl = (
                            <button
                              onClick={() => copyCompanyName(val, rec.id)}
                              title="クリックでコピー"
                              className={`group flex items-center gap-1 w-full min-w-0 text-left transition-colors ${copiedId===rec.id?"text-green-600":"text-slate-800 hover:text-blue-600"}`}>
                              <CompanyLogo logoUrl={rec.logoUrl} url={rec.hpSite} name={val} />
                              <span className="font-medium text-xs whitespace-nowrap truncate max-w-40">{val || "—"}</span>
                              {isOrdered(rec) && (
                                <span className="shrink-0 text-xs px-1.5 py-0.5 rounded-full font-semibold bg-emerald-100 text-emerald-700 border border-emerald-300 whitespace-nowrap">
                                  🛍️受注済案件
                                </span>
                              )}
                              {pd && (
                                <span className="shrink-0 text-xs px-1.5 py-0.5 rounded-full font-semibold bg-purple-100 text-purple-700 border border-purple-300 whitespace-nowrap">
                                  📜{pd.pastStatus || "過去商談あり"}
                                </span>
                              )}
                              {copiedId === rec.id
                                ? <span className="text-green-500 text-xs shrink-0 ml-auto">✓</span>
                                : <svg className="w-3 h-3 shrink-0 text-slate-300 group-hover:text-blue-400 opacity-0 group-hover:opacity-100 transition-opacity ml-auto"
                                    fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                                      d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z"/>
                                  </svg>}
                            </button>
                          );
                        } else if (col.key === "status") {
                          viewEl = <span onClick={openEdit} className="cursor-pointer"><StatusBadge status={val}/></span>;
                        } else if (col.key === "leadSource") {
                          viewEl = val
                            ? <span onClick={openEdit} className="cursor-pointer hover:opacity-80 transition-opacity"><LeadSourceBadge source={val}/></span>
                            : <span onClick={openEdit} className="text-slate-300 text-xs cursor-pointer hover:text-slate-500">— 設定</span>;
                        } else if (col.key === "absenceReason") {
                          viewEl = val
                            ? <span onClick={openEdit} className="cursor-pointer hover:opacity-80 transition-opacity"><AbsenceReasonBadge reason={val}/></span>
                            : <span onClick={openEdit} className="text-slate-300 text-xs cursor-pointer hover:text-slate-500">— 設定</span>;
                        } else if (col.key === "importMonth") {
                          viewEl = val
                            ? <span onClick={openEdit} className="inline-flex items-center gap-1 bg-blue-50 text-blue-700 text-xs font-semibold px-1.5 py-0.5 rounded cursor-pointer hover:bg-blue-100 transition-colors">📅 {String(val).replace("-","/")}</span>
                            : <span onClick={openEdit} className="text-slate-300 text-xs cursor-pointer hover:text-slate-500">— 設定</span>;
                        } else if (col.key === "storeCount") {
                          viewEl = <span onClick={openEdit} className="cursor-pointer hover:bg-slate-50 rounded px-0.5 transition-colors">
                            <StoreCountCell analysis={analyzeStoreCount(rec, storeIndex)} />
                          </span>;
                        } else if ((col.key === "hpSite" || col.key === "gbpSiteUrl") && val) {
                          viewEl = (
                            <span className="flex items-center gap-1">
                              <a href={val} target="_blank" rel="noreferrer"
                                onClick={e => e.stopPropagation()}
                                className="text-blue-600 hover:underline text-xs block max-w-32 truncate">{val}</a>
                              <button onClick={openEdit} className="shrink-0 text-slate-300 hover:text-slate-500">
                                <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/>
                                </svg>
                              </button>
                            </span>
                          );
                        } else if (col.key === "lastCallDate" || col.key === "nextCallDate" || col.key === "leadAddedDate") {
                          const nd = normDate(val) || (col.key === "leadAddedDate" ? today : "");
                          if (!nd) {
                            viewEl = <span onClick={openEdit} className="text-slate-300 text-xs cursor-pointer hover:text-slate-500 hover:bg-slate-50 rounded px-1 transition-colors">— 設定</span>;
                          } else if (col.key === "nextCallDate") {
                            const badge = nd < today
                              ? <span className="inline-flex items-center gap-1 bg-red-100 text-red-700 text-xs font-bold px-1.5 py-0.5 rounded">{fmtDate(nd)} 🔴</span>
                              : nd === today
                                ? <span className="inline-flex items-center gap-1 bg-amber-100 text-amber-700 text-xs font-bold px-1.5 py-0.5 rounded">{fmtDate(nd)} ⚠️</span>
                                : nd <= soon
                                  ? <span className="inline-flex items-center gap-1 bg-yellow-50 text-yellow-700 text-xs font-semibold px-1.5 py-0.5 rounded">{fmtDate(nd)} ⏰</span>
                                  : <span className="text-slate-700 text-xs">{fmtDate(nd)}</span>;
                            viewEl = <span onClick={openEdit} className="cursor-pointer hover:opacity-70 transition-opacity">{badge}</span>;
                          } else {
                            viewEl = <span onClick={openEdit} className="text-slate-700 text-xs cursor-pointer hover:bg-slate-100 rounded px-1 transition-colors">{fmtDate(nd)}</span>;
                          }
                        } else if (col.key === "memo") {
                          viewEl = (
                            <span onClick={openEdit}
                              className="text-slate-600 text-xs block max-w-56 truncate cursor-pointer hover:bg-slate-50 rounded transition-colors"
                              title={val || ""}>
                              {val || empty}
                            </span>
                          );
                        } else {
                          viewEl = (
                            <span onClick={openEdit}
                              className="text-slate-700 text-xs cursor-pointer hover:bg-slate-50 rounded px-0.5 transition-colors">
                              {val || empty}
                            </span>
                          );
                        }
                      }

                      return (
                        <td key={col.key} className={`${col.w||"w-[120px]"} px-3 py-2 whitespace-nowrap align-middle ${isEditing ? "relative" : "overflow-hidden"}`}>
                          {isEditing ? (
                            // 編集時はクリップせず、フロート表示で入力欄を最大表示（日付ピッカーも表示）
                            <div className="absolute z-30 left-1 top-1/2 -translate-y-1/2 bg-white rounded-lg shadow-xl ring-1 ring-blue-300 p-1.5 min-w-[200px] w-max max-w-[360px]">
                              {editEl}
                            </div>
                          ) : (
                            <div className="truncate">{viewEl}</div>
                          )}
                        </td>
                      );
                    })}
                    <td className="w-[80px] px-3 py-2.5 whitespace-nowrap align-middle">
                      <button onClick={() => setEditRec(rec)}
                        className="text-xs text-blue-600 hover:text-blue-800 font-medium">
                        編集
                      </button>
                    </td>
                  </tr>
                  );
                })}
              </tbody>
            </table>
          </div>{/* スクロール枠ここまで */}

          {/* ── Pagination（枠外固定） ── */}
          <div className="px-4 py-3 border-t border-slate-100 flex items-center justify-between gap-4 flex-wrap bg-white rounded-b-xl">
            <span className="text-xs text-slate-400">
              {displayList.length > 0
                ? `${displayList.length}件中 ${(page-1)*PAGE_SIZE+1}–${Math.min(page*PAGE_SIZE, displayList.length)} 件表示（全${records.length}件）`
                : `全${records.length}件`}
            </span>
            {totalPages > 1 && (
              <div className="flex items-center gap-1">
                <button onClick={() => goPage(1)} disabled={page===1}
                  className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">«</button>
                <button onClick={() => goPage(page-1)} disabled={page===1}
                  className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">‹</button>
                {pageNums.map((p, i) =>
                  p === "…" ? (
                    <span key={`ellipsis-${i}`} className="px-1.5 text-slate-400 text-xs">…</span>
                  ) : (
                    <button key={p} onClick={() => goPage(p)}
                      className={`w-7 h-7 text-xs rounded border font-medium transition-colors
                        ${page===p ? "bg-blue-600 text-white border-blue-600" : "border-slate-200 text-slate-600 hover:bg-slate-50"}`}>
                      {p}
                    </button>
                  )
                )}
                <button onClick={() => goPage(page+1)} disabled={page===totalPages}
                  className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">›</button>
                <button onClick={() => goPage(totalPages)} disabled={page===totalPages}
                  className="px-2 py-1 text-xs rounded border border-slate-200 text-slate-500 disabled:opacity-30 hover:bg-slate-50">»</button>
              </div>
            )}
          </div>
        </div>

        </> /* end list view */}
      </div>

      {/* ── Modals ── */}
      {showSettings && (
        <SettingsModal settings={settings} onSave={s => setSettings(s)} onClose={() => setShowSettings(false)}
          onExportBackup={exportBackup} onImportBackup={importBackup} onRollback={rollbackAutoBackup}
          hasAutoBackup={hasAutoBackup} dataCounts={{ records: records.length, pastMgmt: pastMgmt.length }} />
      )}
      {showImport && (
        <ImportModal onImport={addRecords} onImportPastDeals={addPastDeals} onImportMetel={importMetelMerge} onImportOrders={addOrders} onClose={() => setShowImport(false)} />
      )}
      {showDupe && (
        <DuplicateModal records={records} onClean={cleanDuplicates} onClose={() => setShowDupe(false)} />
      )}
      {showBulkStatus && (
        <StatusBulkUpdateModal records={records} onUpdate={bulkUpdateStatus} onClose={() => setShowBulkStatus(false)} />
      )}
      {showNew && (
        <RecordFormModal
          initial={{}}
          title="新規レコード追加"
          onSave={form => saveRecord(form)}
          onClose={() => setShowNew(false)}
        />
      )}
      {editRec && (
        <RecordFormModal
          initial={editRec}
          title="レコード編集"
          onSave={form => saveRecord({ ...editRec, ...form })}
          onClose={() => setEditRec(null)}
          onDelete={() => { deleteRecord(editRec.id); setEditRec(null); }}
          pastDeal={findPastDeal(editRec.companyName)}
          storeEstimate={analyzeStoreCount(editRec, storeIndex)}
          ordered={isOrdered(editRec)}
          onCopyToOrder={() => { addOrderFromRecord(editRec); window.alert("🛍️ 受注案件管理に登録しました。受注案件タブで詳細を入力してください。"); }}
        />
      )}
      {showHelp && <HelpModal onClose={() => setShowHelp(false)} />}
    </div>
  );
}
