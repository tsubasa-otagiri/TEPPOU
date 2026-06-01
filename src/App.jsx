import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ── Constants ──────────────────────────────────────────────────────────────────
const PASSWORD     = "1111";
const STORAGE_KEY  = "teppou_records_v3";   // legacy migration source
const SETTINGS_KEY = "teppou_settings_v1";
const PAGE_SIZE    = 100;

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
const IDB_NAME     = "teppou_idb";
const IDB_VER      = 1;
const IDB_STORE    = "records";

// ── IndexedDB helpers ──────────────────────────────────────────────────────────
function idbOpen() {
  return new Promise((res, rej) => {
    const req = indexedDB.open(IDB_NAME, IDB_VER);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(IDB_STORE))
        db.createObjectStore(IDB_STORE, { keyPath: "id" });
    };
    req.onsuccess = e => res(e.target.result);
    req.onerror   = e => rej(e.target.error);
  });
}
async function idbGetAll() {
  const db = await idbOpen();
  return new Promise((res, rej) => {
    const req = db.transaction(IDB_STORE, "readonly").objectStore(IDB_STORE).getAll();
    req.onsuccess = () => res(req.result || []);
    req.onerror   = e => rej(e.target.error);
  });
}
async function idbPutAll(records) {
  const db = await idbOpen();
  return new Promise((res, rej) => {
    const tx    = db.transaction(IDB_STORE, "readwrite");
    const store = tx.objectStore(IDB_STORE);
    store.clear();
    records.forEach(r => store.put(r));
    tx.oncomplete = res;
    tx.onerror    = e => rej(e.target.error);
  });
}

const IS_MEMBERS = [
  "櫻井　肇","上浦　諒大","井上　妃音","太田　小百合","十文字　菜月",
  "中　翔吾","早坂　直樹","小田切　翼","横井　優一","青木　大輔",
];

const STATUS_CFG = {
  "0.日程調整":        { row:"bg-rose-50",    bg:"bg-rose-200",   text:"text-rose-800",   border:"border-rose-400",   dot:"bg-rose-600"   },
  "1.高確度":          { row:"bg-pink-50",    bg:"bg-pink-100",   text:"text-pink-700",   border:"border-pink-300",   dot:"bg-pink-500"   },
  "2.優先":            { row:"bg-orange-50",  bg:"bg-orange-100", text:"text-orange-700", border:"border-orange-300", dot:"bg-orange-400" },
  "3.並":              { row:"",              bg:"bg-gray-100",   text:"text-gray-500",   border:"border-gray-300",   dot:"bg-gray-400"   },
  "4.受付カット":      { row:"bg-slate-50",   bg:"bg-slate-100",  text:"text-slate-600",  border:"border-slate-300",  dot:"bg-slate-500"  },
  "4.別担当架電":      { row:"bg-sky-50",     bg:"bg-sky-100",    text:"text-sky-700",    border:"border-sky-300",    dot:"bg-sky-500"    },
  "4.商談中":          { row:"bg-violet-50",  bg:"bg-violet-100", text:"text-violet-700", border:"border-violet-300", dot:"bg-violet-500" },
  "5.メール送付":      { row:"bg-yellow-50",  bg:"bg-yellow-100", text:"text-yellow-700", border:"border-yellow-300", dot:"bg-amber-400"  },
  "6.コネクト（改）":  { row:"bg-indigo-50",  bg:"bg-indigo-100", text:"text-indigo-600", border:"border-indigo-300", dot:"bg-indigo-400" },
  "7.コネクト（無）":  { row:"bg-blue-50",    bg:"bg-blue-100",   text:"text-blue-900",   border:"border-blue-600",   dot:"bg-blue-800"   },
  "8.不要":            { row:"bg-purple-50",  bg:"bg-purple-100", text:"text-purple-900", border:"border-purple-600", dot:"bg-purple-800" },
  "8.当社契約":        { row:"bg-gray-100",   bg:"bg-gray-200",   text:"text-gray-800",   border:"border-gray-500",   dot:"bg-gray-700"   },
  "9.アポ獲得":        { row:"bg-red-50",     bg:"bg-red-100",    text:"text-red-900",    border:"border-red-600",    dot:"bg-red-700"    },
};

const ALL_COLUMNS = [
  { key:"companyName",   label:"企業名",                              required:true  },
  { key:"lastCallDate",  label:"架電日",                              required:false },
  { key:"nextCallDate",  label:"次回架電日",                          required:false },
  { key:"status",        label:"状況",                                required:false },
  { key:"industry",      label:"業種",                                required:false },
  { key:"leadSource",    label:"ソース",                              required:false },
  { key:"hpSite",        label:"HPサイト",                            required:false },
  { key:"gbp",           label:"GBP",                                 required:false },
  { key:"phone",         label:"電話番号",                            required:false },
  { key:"assignee",      label:"担当者",                              required:false },
  { key:"department",    label:"部署",                                required:false },
  { key:"absenceReason", label:"不在理由",                            required:false },
  { key:"gbpManagement", label:"GBPの管理",                           required:false },
  { key:"memo",          label:"簡易的な情報/次回架電時の情報メモ",    required:false },
  { key:"storeCount",    label:"店舗数",                              required:false },
  { key:"refusalReason", label:"断り理由",                            required:false },
  { key:"posting",       label:"投稿",                                required:false },
  { key:"review",        label:"口コミ",                              required:false },
  { key:"sns",           label:"SNS",                                 required:false },
  { key:"instagram",     label:"Insta",                               required:false },
  { key:"line",          label:"Line",                                required:false },
  { key:"facebook",      label:"FB",                                  required:false },
  { key:"twitter",       label:"Twitter",                             required:false },
  { key:"os",            label:"OS",                                  required:false },
  { key:"mailFlag",      label:"メール",                              required:false },
  { key:"email",         label:"メアド",                              required:false },
  { key:"gbpSiteUrl",    label:"GBPサイトURL",                        required:false },
];

const DEFAULT_VISIBLE_COLS = [
  "companyName","lastCallDate","nextCallDate","status","storeCount","phone","assignee","memo",
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
    if      (/企業名|会社名|法人名/.test(n))                      m.companyName   = i;
    else if (/電話|tel|phone/.test(n))                            m.phone         = i;
    else if (/メアド/.test(n))                                    m.email         = i;
    else if (/^メール$/.test(n))                                  m.mailFlag      = i;
    else if (/gbpサイト/.test(n))                                 m.gbpSiteUrl    = i;
    else if (/gbpの管理|gbp管理/.test(n))                         m.gbpManagement = i;
    else if (/^gbp$/.test(n))                                     m.gbp           = i;
    else if (/^hp$|hpサイト|^hp[^a-z]/.test(n))                  m.hpSite        = i;
    else if (/簡易|次回架電時の情報|メモ|備考|note|コメント/.test(n)) m.memo       = i;
    else if (/担当者?名?|営業担当/.test(n))                       m.assignee      = i;
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

// ── StatusBadge ────────────────────────────────────────────────────────────────
function StatusBadge({ status }) {
  const c = STATUS_CFG[status] ?? { bg:"bg-gray-100", text:"text-gray-600", border:"border-gray-300" };
  return (
    <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium border whitespace-nowrap ${c.bg} ${c.text} ${c.border}`}>
      {status||"—"}
    </span>
  );
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
              className="max-w-full cursor-crosshair select-none"
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
function LoginScreen({ onLogin, logo }) {
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  const submit = e => {
    e.preventDefault();
    if (pw === PASSWORD) { sessionStorage.setItem("teppou_auth","1"); onLogin(); }
    else { setErr("パスワードが正しくありません"); setPw(""); }
  };
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-800 to-slate-900 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl p-10 w-full max-w-sm">
        <div className="text-center mb-8">
          <div className="flex justify-center mb-4"><AppIcon logo={logo} size="lg" /></div>
          <h1 className="text-2xl font-bold text-slate-800">TEPPOU</h1>
          <p className="text-slate-500 text-sm mt-1">GMOテック IS部門</p>
        </div>
        <form onSubmit={submit} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">パスワード</label>
            <input type="password" value={pw} autoFocus
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
      </div>
    </div>
  );
}

// ── SettingsModal ──────────────────────────────────────────────────────────────
function SettingsModal({ settings, onSave, onClose }) {
  const [logo,       setLogo]       = useState(settings.logo    || null);
  const [favicon,    setFavicon]    = useState(settings.favicon || null);
  const [cropSrc,    setCropSrc]    = useState(null);
  const [cropTarget, setCropTarget] = useState(null); // "logo" | "favicon"
  const logoRef    = useRef();
  const faviconRef = useRef();

  const openCrop = (file, target) => {
    const r = new FileReader();
    r.onload = e => { setCropSrc(e.target.result); setCropTarget(target); };
    r.readAsDataURL(file);
  };

  const onCropDone = (dataUrl) => {
    if (cropTarget === "logo")    setLogo(dataUrl);
    if (cropTarget === "favicon") setFavicon(dataUrl);
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

          {/* Logo */}
          <div className="mb-6">
            <p className="text-sm font-semibold text-slate-700 mb-3">アプリロゴ</p>
            <div className="flex items-center gap-4">
              <div className="w-16 h-16 border-2 border-dashed border-slate-300 rounded-xl flex items-center justify-center overflow-hidden bg-slate-50 shrink-0">
                {logo
                  ? <img src={logo} alt="logo" className="w-full h-full object-contain" />
                  : <span className="text-xs text-slate-400 text-center px-1">未設定</span>}
              </div>
              <div className="flex flex-col gap-2">
                <UploadBtn inputRef={logoRef} target="logo" />
                {logo && (
                  <button onClick={() => setLogo(null)}
                    className="text-xs text-rose-500 hover:text-rose-700 text-left">
                    ロゴを削除
                  </button>
                )}
              </div>
            </div>
            <p className="text-xs text-slate-400 mt-2">
              PNG / JPEG 対応。アップロード後にトリミングできます。ログイン画面・ヘッダーに反映されます。
            </p>
          </div>

          {/* Favicon */}
          <div className="mb-6">
            <p className="text-sm font-semibold text-slate-700 mb-3">ファビコン（ブラウザタブアイコン）</p>
            <div className="flex items-center gap-4">
              <div className="w-10 h-10 border-2 border-dashed border-slate-300 rounded-lg flex items-center justify-center overflow-hidden bg-slate-50 shrink-0">
                {favicon
                  ? <img src={favicon} alt="favicon" className="w-8 h-8 object-contain" />
                  : <span className="text-xs text-slate-400">—</span>}
              </div>
              <div className="flex flex-col gap-2">
                <UploadBtn inputRef={faviconRef} target="favicon" />
                {favicon && (
                  <button onClick={() => setFavicon(null)}
                    className="text-xs text-rose-500 hover:text-rose-700 text-left">
                    ファビコンを削除
                  </button>
                )}
              </div>
            </div>
            <p className="text-xs text-slate-400 mt-2">
              ICO / PNG 推奨。保存するとブラウザタブのアイコンに即時反映されます。
            </p>
          </div>

          <div className="flex gap-2 justify-end pt-4 border-t border-slate-100">
            <button onClick={onClose}
              className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">
              キャンセル
            </button>
            <button onClick={() => { onSave({ logo, favicon }); onClose(); }}
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
      ["架電日","次回架電日","状況","業種","ソース","企業名","HPサイト","GBP","電話番号","担当者","部署","不在理由","GBPの管理","簡易的な情報/次回架電時の情報メモ","店舗数","断り理由","投稿","口コミ","SNS","Insta","Line","FB","Twitter","OS","メール","メアド","GBPサイトURL"],
      ["2026-06-01","2026-06-08","未架電","IT","自社リスト","株式会社サンプル","https://example.com","あり","03-0000-0000","山田太郎","営業部","","○","サンプルメモ","1","","○","★3","○","○","○","○","○","iOS","","sample@example.com","https://maps.google.com/..."],
    ],
  },
  metel: {
    filename: "TEPPOU_ミーてるログ_フォーマット.csv",
    rows: [
      ["担当者","企業名","電話番号","通話種別","タグ","架電日時","メモ"],
      ["櫻井　肇","株式会社サンプル","03-0000-0000","発信（不在）","担当者不在","2026-06-01 10:00:00","サンプルメモ"],
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
function ImportModal({ onImport, onClose }) {
  const [mode,      setMode]      = useState("sales");
  const [inputMode, setInputMode] = useState("file");   // "file" | "paste"
  const [pasteText, setPasteText] = useState("");
  const [log,       setLog]       = useState(null);
  const [loading,   setLoading]   = useState(false);
  const [progress,  setProgress]  = useState("");
  const fileRef = useRef();

  const processRows = (rows) => {
    setLoading(true);
    setProgress(`${rows.length.toLocaleString()} 行を解析中...`);
    // setTimeout で React にスピナー描画の猶予を与えてからブロック処理
    setTimeout(() => {
      try {
        const result = mode === "sales" ? doImportSales(rows) : doImportMetel(rows);
        setLog(result);
        if (result.records?.length > 0) onImport(result.records);
      } catch (ex) {
        setLog({ error: `インポートエラー: ${ex.message}`, records: [] });
      } finally {
        setLoading(false);
        setProgress("");
      }
    }, 60);
  };

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
    const records   = [];
    let skipped = 0;
    for (const row of rows.slice(headerIdx + 1)) {
      if (row.every(c => !c.trim())) continue;
      const company = map.companyName !== undefined ? row[map.companyName] : "";
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
        importedAt: nowIso(), updatedAt: nowIso(), source:"csv",
      });
    }
    return { success:true, records, skipped, autoSkip: headerIdx === 1 };
  }

  function doImportMetel(rows) {
    if (rows.length < 2) return { error:"データ行が不足しています", records:[] };
    const headers = rows[0];
    const col = patterns => {
      for (let i = 0; i < headers.length; i++) {
        const h = headers[i].replace(/[\s　]/g,"");
        if (patterns.some(p => h.includes(p))) return i;
      }
      return -1;
    };
    const cAssignee = col(["担当者","オペレーター","エージェント","架電者"]);
    const cCompany  = col(["企業名","会社名","顧客名","取引先名"]);
    const cPhone    = col(["電話番号","TEL","電話"]);
    const cCallType = col(["通話種別","架電種別","種別"]);
    const cTags     = col(["タグ","ラベル"]);
    const cDate     = col(["架電日時","通話日時","日時","日付","架電日"]);
    const cMemo     = col(["メモ","備考","コメント"]);
    const records = []; let filtered = 0, skipped = 0;

    for (const row of rows.slice(1)) {
      if (row.every(c => !c.trim())) continue;
      const assignee = cAssignee >= 0 ? (row[cAssignee]||"").trim() : "";
      const isMember = IS_MEMBERS.some(m => {
        const mn = normName(m), an = normName(assignee);
        return an === mn || an.includes(mn) || mn.includes(an);
      });
      if (!isMember) { filtered++; continue; }
      const company = cCompany >= 0 ? (row[cCompany]||"").trim() : "";
      if (!company) { skipped++; continue; }
      const callType = cCallType >= 0 ? (row[cCallType]||"").trim() : "";
      const tags     = cTags     >= 0 ? (row[cTags]    ||"").trim() : "";
      const status   = convertMitelStatus(callType, tags) ?? "不通";
      const rawDate  = cDate >= 0 ? (row[cDate]||"").trim() : "";
      const dateStr  = rawDate ? rawDate.slice(0,10).replace(/\//g,"-") : "";
      const baseMemo = cMemo >= 0 ? (row[cMemo]||"").trim() : "";
      const memo     = tags ? `【ミーてるタグ】${tags}${baseMemo ? "\n"+baseMemo : ""}` : baseMemo;
      records.push({
        id:genId(), companyName:company, phone: cPhone>=0?(row[cPhone]||"").trim():"",
        email:"", url:"", status, assignee, lastCallDate:dateStr, nextCallDate:"",
        callCount:1, memo, importedAt:nowIso(), updatedAt:nowIso(), source:"metel",
      });
    }
    return { success:true, records, filtered, skipped };
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
              desc:"1行目が集計行のCSVも自動補正。企業名・電話・メモ等を自動マッピング。" },
            { value:"metel", icon:"📞", title:"ミーてるの架電ログを取り込む",
              desc:"ISメンバー10名に自動絞り込み。タグ → ステータス自動変換 & メモへ記録。" },
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
                ✅ インポート完了: <strong>{log.records.length}件</strong>追加
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

// ── DuplicateModal ─────────────────────────────────────────────────────────────
// ステータス優先度（低い数字ほど価値が高い → 残す候補）
const DUPE_PRIORITY = {
  "9.アポ獲得":1, "0.日程調整":2, "1.高確度":3, "4.商談中":4,
  "5.メール送付":5, "6.コネクト（改）":6, "7.コネクト（無）":7,
  "2.優先":8, "4.別担当架電":9, "4.受付カット":10,
  "8.当社契約":11, "8.不要":12, "3.並":13,
  "未架電":99, "":100,
};
const dupePrio = s => DUPE_PRIORITY[s] ?? 50;

// 優先度→更新日の順でソート（先頭＝残す）
function sortGroupForDupe(rs) {
  return [...rs].sort((a,b) => {
    const dp = dupePrio(a.status) - dupePrio(b.status);
    if (dp !== 0) return dp;
    return new Date(b.updatedAt||b.importedAt) - new Date(a.updatedAt||a.importedAt);
  });
}

function DuplicateModal({ records, onClean, onClose }) {
  const [step,      setStep]      = useState("select"); // "select" | "confirm"
  const [deleteSet, setDeleteSet] = useState(() => {
    const s = new Set();
    const g = {};
    records.forEach(r => {
      const k = normName(r.companyName); if (!k) return;
      (g[k] = g[k]||[]).push(r);
    });
    Object.values(g).filter(rs => rs.length > 1).forEach(rs => {
      // 先頭（最優先）を残し、残りをデフォルト削除対象に
      sortGroupForDupe(rs).slice(1).forEach(r => s.add(r.id));
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
      .map(rs => sortGroupForDupe(rs));
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
function RecordFormModal({ initial, title, onSave, onClose, onDelete }) {
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
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl flex flex-col max-h-[92vh]">
        <div className="flex items-center justify-between px-6 pt-5 pb-4 border-b border-slate-100 shrink-0">
          <h2 className="text-base font-bold text-slate-800">{title}</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
        </div>
        <div className="overflow-y-auto flex-1 px-6 py-4">
          <div className="grid grid-cols-2 gap-3">

            {/* 企業情報 */}
            <SectionLabel>企業情報</SectionLabel>
            <div className="col-span-2">
              <label className="block text-xs text-slate-500 mb-1">企業名 <span className="text-rose-500">*</span></label>
              <input type="text" value={form.companyName} autoFocus
                onChange={e => upd("companyName", e.target.value)}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
            </div>
            {txt("phone",      "電話番号")}
            {txt("department", "部署")}
            {txt("industry",   "業種")}
            {txt("leadSource", "ソース")}
            {txt("storeCount", "店舗数")}

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
            {txt("absenceReason", "不在理由")}
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
              <label className="block text-xs text-slate-500 mb-1">簡易的な情報/次回架電時の情報メモ</label>
              <textarea value={form.memo||""} onChange={e => upd("memo", e.target.value)} rows={4}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none" />
            </div>
          </div>
        </div>
        <div className="flex items-center justify-between px-6 py-4 border-t border-slate-100 shrink-0">
          <div>
            {onDelete && (
              <button onClick={onDelete}
                className="px-4 py-2 text-sm text-rose-500 hover:text-rose-700 border border-rose-200 hover:border-rose-400 hover:bg-rose-50 rounded-lg transition-colors">
                🗑️ 削除する
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

// ── AnalysisView ───────────────────────────────────────────────────────────────
function AnalysisView({ records }) {
  const today = getToday();
  const soon  = (() => { const d=new Date(); d.setDate(d.getDate()+3); return d.toISOString().slice(0,10); })();

  const kpis = [
    { label:"総件数",      value: records.length,                                                                               color:"text-blue-600",   bg:"bg-blue-50"   },
    { label:"本日架電予定", value: records.filter(r=>normDate(r.nextCallDate)===today).length,                                    color:"text-amber-600",  bg:"bg-amber-50"  },
    { label:"3日以内期限",  value: records.filter(r=>{ const nd=normDate(r.nextCallDate); return nd>today&&nd<=soon; }).length,   color:"text-yellow-600", bg:"bg-yellow-50" },
    { label:"商談/アポ",    value: records.filter(r=>["0.日程調整","1.高確度","4.商談中","9.アポ獲得"].includes(r.status)).length, color:"text-teal-600",   bg:"bg-teal-50"   },
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

// ── Main App ───────────────────────────────────────────────────────────────────
export default function App() {
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
  const [settings,       setSettings]       = useState({ logo:null, favicon:null });
  const [search,         setSearch]         = useState("");
  const [statusFilter,   setStatusFilter]   = useState("all");
  const [assigneeFilter, setAssigneeFilter] = useState("all");
  const [visibleCols,    setVisibleCols]    = useState(DEFAULT_VISIBLE_COLS);
  const [showColDrop,    setShowColDrop]    = useState(false);
  const [page,           setPage]           = useState(1);
  const [showSettings,   setShowSettings]   = useState(false);
  const [showImport,     setShowImport]     = useState(false);
  const [showDupe,       setShowDupe]       = useState(false);
  const [showNew,        setShowNew]        = useState(false);
  const [editRec,        setEditRec]        = useState(null);
  const [selected,       setSelected]       = useState(new Set());
  const [view,           setView]           = useState("list");   // "list" | "analysis"
  const [showPullList,   setShowPullList]   = useState(false);
  const [copiedId,       setCopiedId]       = useState(null);
  const [sortKey,        setSortKey]        = useState(null);
  const [sortDir,        setSortDir]        = useState("asc");
  const [storageWarning, setStorageWarning] = useState(false);
  const colDropRef = useRef();

  // ── Persistence ──────────────────────────────────────────────────────────────
  // 初回ロード: IndexedDB → なければ localStorage から移行
  useEffect(() => {
    idbGetAll().then(recs => {
      if (recs.length > 0) {
        setRecords(recs);
      } else {
        // localStorage からの移行
        try {
          const s = localStorage.getItem(STORAGE_KEY);
          if (s) {
            const parsed = JSON.parse(s);
            setRecords(parsed);
            idbPutAll(parsed).then(() => localStorage.removeItem(STORAGE_KEY));
          }
        } catch {}
      }
    }).catch(() => {
      // IndexedDB 使用不可の場合は localStorage にフォールバック
      try { const s = localStorage.getItem(STORAGE_KEY); if (s) setRecords(JSON.parse(s)); } catch {}
    });
    try { const s = localStorage.getItem(SETTINGS_KEY); if (s) setSettings(JSON.parse(s)); } catch {}
  }, []);

  // records 変更時: IndexedDB に保存（500ms デバウンス）
  const idbSaveTimer = useRef(null);
  useEffect(() => {
    if (idbSaveTimer.current) clearTimeout(idbSaveTimer.current);
    idbSaveTimer.current = setTimeout(() => {
      idbPutAll(records).then(() => setStorageWarning(false)).catch(() => {
        // IndexedDB 失敗時は localStorage にフォールバック
        try { localStorage.setItem(STORAGE_KEY, JSON.stringify(records)); setStorageWarning(false); }
        catch { setStorageWarning(true); }
      });
    }, 500);
    return () => clearTimeout(idbSaveTimer.current);
  }, [records]);

  useEffect(() => { try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); } catch {} }, [settings]);

  // ── Favicon ───────────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!settings.favicon) return;
    let link = document.querySelector("link[rel~='icon']");
    if (!link) { link = document.createElement("link"); link.rel = "icon"; document.head.appendChild(link); }
    link.type = "image/png";
    link.setAttribute("sizes", "512x512");
    link.href = settings.favicon;
  }, [settings.favicon]);

  // ── Close column dropdown on outside click ────────────────────────────────────
  useEffect(() => {
    if (!showColDrop) return;
    const fn = e => { if (colDropRef.current && !colDropRef.current.contains(e.target)) setShowColDrop(false); };
    document.addEventListener("mousedown", fn);
    return () => document.removeEventListener("mousedown", fn);
  }, [showColDrop]);

  const today = getToday();

  // ── Derived data ──────────────────────────────────────────────────────────────
  const filtered = records.filter(r => {
    if (statusFilter !== "all" && r.status !== statusFilter) return false;
    if (assigneeFilter !== "all" && r.assignee !== assigneeFilter) return false;
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
        if (sortKey === "lastCallDate" || sortKey === "nextCallDate") { va = normDate(va); vb = normDate(vb); }
        if (sortKey === "storeCount") { va = parseInt(String(va).replace(/,/g,""))||0; vb = parseInt(String(vb).replace(/,/g,""))||0; return sortDir==="asc" ? va-vb : vb-va; }
        const cmp = String(va).localeCompare(String(vb), "ja");
        return sortDir === "asc" ? cmp : -cmp;
      })
    : filtered;

  const totalPages = Math.max(1, Math.ceil(sortedFiltered.length / PAGE_SIZE));
  const paginated  = sortedFiltered.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);

  const statsMap = {};
  records.forEach(r => { statsMap[r.status] = (statsMap[r.status]||0) + 1; });
  const stats = Object.entries(statsMap).map(([s,c]) => ({ status:s, count:c, ...(STATUS_CFG[s]??{}) }));

  const soon = (() => { const d = new Date(); d.setDate(d.getDate()+3); return d.toISOString().slice(0,10); })();
  const doneStatuses = ["8.不要","8.当社契約"];
  const alerts   = records.filter(r => r.nextCallDate && normDate(r.nextCallDate) <= soon && !doneStatuses.includes(r.status));
  const assignees = [...new Set(records.map(r => r.assignee).filter(Boolean))];
  const visibleDefs = ALL_COLUMNS.filter(c => visibleCols.includes(c.key));

  // ── Mutations ─────────────────────────────────────────────────────────────────
  const addRecords = useCallback(recs => { setRecords(p => [...p, ...recs]); setPage(1); }, []);

  const saveRecord = useCallback(form => {
    const isEdit = records.some(r => r.id === form.id);
    if (isEdit) {
      setRecords(p => p.map(r => r.id===form.id ? { ...r, ...form, updatedAt:nowIso() } : r));
    } else {
      setRecords(p => [...p, { ...form, id:genId(), callCount:form.callCount||0, importedAt:nowIso(), updatedAt:nowIso(), source:"manual" }]);
    }
  }, [records]);

  const deleteRecord = useCallback(id => {
    if (!window.confirm("このレコードを削除しますか？")) return;
    setRecords(p => p.filter(r => r.id !== id));
    setSelected(p => { const n = new Set(p); n.delete(id); return n; });
  }, []);

  const deleteSelected = useCallback(() => {
    if (!selected.size) return;
    if (!window.confirm(`選択した ${selected.size} 件を削除しますか？`)) return;
    setRecords(p => p.filter(r => !selected.has(r.id)));
    setSelected(new Set());
  }, [selected]);

  const copyCompanyName = useCallback((text, id) => {
    navigator.clipboard.writeText(text).then(() => {
      setCopiedId(id);
      setTimeout(() => setCopiedId(null), 1500);
    });
  }, []);

  const cleanDuplicates = useCallback(ids => {
    const del = new Set(ids);
    setRecords(p => p.filter(r => !del.has(r.id)));
  }, []);

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
    <div className="min-h-screen bg-slate-50 text-slate-800">

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
            {alerts.length > 0 && (
              <span className="bg-amber-100 text-amber-700 text-xs font-semibold px-2.5 py-1 rounded-full border border-amber-300">
                ⏰ {alerts.length}件アラート
              </span>
            )}
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
          {[["list","📋 リスト"],["analysis","📊 分析"],["pull","🔍 プル照合"]].map(([v,label])=>(
            <button key={v} onClick={()=>setView(v)}
              className={`px-4 py-1.5 text-xs font-semibold rounded-lg transition-colors
                ${view===v ? "bg-white text-blue-700 shadow-sm" : "text-slate-500 hover:text-slate-700"}`}>
              {label}
            </button>
          ))}
        </div>

        {/* ── Analysis view ── */}
        {view==="analysis" && <AnalysisView records={records} />}

        {/* ── Pull view ── */}
        {view==="pull" && <PullView records={records} />}

        {/* ── List view ── */}
        {view==="list" && <>

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

        {/* ── Alert bar ── */}
        {alerts.length > 0 && (
          <div className="bg-amber-50 border border-amber-300 rounded-xl px-4 py-3 flex flex-wrap items-center gap-2">
            <span className="text-sm font-semibold text-amber-700 shrink-0">📅 次回架電日アラート</span>
            {alerts.slice(0,5).map(r => (
              <span key={r.id} className="bg-amber-100 border border-amber-300 text-amber-800 text-xs px-2 py-0.5 rounded-full">
                {r.companyName}（{fmtDate(normDate(r.nextCallDate))}）
              </span>
            ))}
            {alerts.length > 5 && <span className="text-xs text-amber-600">他 {alerts.length-5} 件</span>}
          </div>
        )}

        {/* ── Stats bar ── */}
        {stats.length > 0 && (
          <div className="bg-white rounded-xl border border-slate-200 p-4">
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-3">ステータス別集計</p>
            <div className="flex flex-wrap gap-2">
              {stats.map(s => (
                <button key={s.status}
                  onClick={() => { setStatusFilter(statusFilter===s.status ? "all" : s.status); setPage(1); }}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg border text-xs font-medium transition-all
                    ${statusFilter===s.status
                      ? `${s.bg??"bg-gray-100"} ${s.text??"text-gray-600"} ${s.border??"border-gray-300"} ring-2 ring-offset-1 ring-blue-400`
                      : `${s.bg??"bg-gray-100"} ${s.text??"text-gray-600"} ${s.border??"border-gray-300"} hover:opacity-80`}`}>
                  <span className={`w-2 h-2 rounded-full ${s.dot??"bg-gray-400"}`} />
                  {s.status}
                  <span className="font-bold">{s.count}</span>
                </button>
              ))}
              {statusFilter !== "all" && (
                <button onClick={() => { setStatusFilter("all"); setPage(1); }}
                  className="px-3 py-1.5 text-xs border border-slate-300 rounded-lg text-slate-500 hover:bg-slate-50">
                  全表示
                </button>
              )}
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

            {/* Duplicate cleanse */}
            <button onClick={() => setShowDupe(true)}
              className="flex items-center gap-1.5 bg-amber-50 hover:bg-amber-100 border border-amber-300 text-amber-700 px-3 py-2 rounded-lg text-xs font-medium transition-colors">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                  d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" />
              </svg>
              重複クレンジング
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
                <div className="absolute left-0 top-full mt-1 bg-white rounded-xl border border-slate-200 shadow-xl z-20 p-3 min-w-44">
                  <p className="text-xs font-semibold text-slate-400 mb-2 px-1">表示する列</p>
                  {ALL_COLUMNS.map(col => (
                    <label key={col.key}
                      className="flex items-center gap-2.5 px-2 py-1.5 hover:bg-slate-50 rounded-lg cursor-pointer">
                      <input type="checkbox" checked={visibleCols.includes(col.key)} disabled={col.required}
                        onChange={e => setVisibleCols(p => e.target.checked ? [...p, col.key] : p.filter(k => k!==col.key))}
                        className="rounded border-slate-300 text-blue-600" />
                      <span className="text-xs text-slate-700">{col.label}</span>
                      {col.required && <span className="ml-auto text-xs text-slate-300">必須</span>}
                    </label>
                  ))}
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
                className="w-full pl-9 pr-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" />
            </div>
            <select value={statusFilter} onChange={e => { setStatusFilter(e.target.value); setPage(1); }}
              className="border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
              <option value="all">全ステータス</option>
              {Object.keys(STATUS_CFG).map(s => <option key={s} value={s}>{s}</option>)}
            </select>
            <select value={assigneeFilter} onChange={e => { setAssigneeFilter(e.target.value); setPage(1); }}
              className="border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
              <option value="all">全担当者</option>
              {assignees.map(a => <option key={a} value={a}>{a}</option>)}
            </select>
          </div>
        </div>

        {/* ── Table ── */}
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm border-collapse">
              <thead className="bg-slate-50 border-b border-slate-200">
                <tr>
                  <th className="w-10 px-3 py-3 text-left">
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
                      className="px-3 py-3 text-left text-xs font-semibold text-slate-500 whitespace-nowrap cursor-pointer hover:bg-slate-100 select-none transition-colors">
                      <span className="inline-flex items-center gap-1">
                        {col.label}
                        {sortKey === col.key
                          ? <span className="text-blue-500">{sortDir==="asc" ? "▲" : "▼"}</span>
                          : <span className="text-slate-300">⇅</span>}
                      </span>
                    </th>
                  ))}
                  <th className="px-3 py-3 text-left text-xs font-semibold text-slate-500">操作</th>
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
                  const rowColor = selected.has(rec.id) ? "bg-blue-100" : (STATUS_CFG[rec.status]?.row ?? "");
                  return (
                  <tr key={rec.id} className={`transition-colors hover:brightness-95 ${rowColor}`}>
                    <td className="px-3 py-2.5">
                      <input type="checkbox" checked={selected.has(rec.id)}
                        onChange={e => toggleSelect(rec.id, e.target.checked)}
                        className="rounded border-slate-300 text-blue-600" />
                    </td>
                    {visibleDefs.map(col => (
                      <td key={col.key} className="px-3 py-2.5 whitespace-nowrap max-w-xs">
                        {col.key === "companyName" ? (
                          <button onClick={() => copyCompanyName(rec.companyName, rec.id)}
                            title="クリックでコピー"
                            className={`group flex items-center gap-1 text-xs text-left w-full transition-colors
                              ${copiedId===rec.id ? "text-green-600" : "text-slate-800 hover:text-blue-600"}`}>
                            <span className="font-medium truncate max-w-48">{rec.companyName||"—"}</span>
                            {copiedId===rec.id
                              ? <span className="text-green-500 text-xs shrink-0">✓</span>
                              : <svg className="w-3 h-3 shrink-0 text-slate-300 group-hover:text-blue-400 opacity-0 group-hover:opacity-100 transition-opacity"
                                  fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                                    d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z"/>
                                </svg>
                            }
                          </button>
                        ) : col.key === "status" ? (
                          <StatusBadge status={rec.status} />
                        ) : (col.key === "hpSite" || col.key === "gbpSiteUrl") && rec[col.key] ? (
                          <a href={rec[col.key]} target="_blank" rel="noreferrer"
                            className="text-blue-600 hover:underline text-xs block max-w-36 truncate">
                            {rec[col.key]}
                          </a>
                        ) : col.key === "nextCallDate" && rec.nextCallDate ? (
                          (() => { const nd = normDate(rec.nextCallDate); return (
                            nd < today
                              ? <span className="inline-flex items-center gap-1 bg-red-100 text-red-700 text-xs font-bold px-1.5 py-0.5 rounded">{fmtDate(nd)} 🔴</span>
                              : nd === today
                                ? <span className="inline-flex items-center gap-1 bg-amber-100 text-amber-700 text-xs font-bold px-1.5 py-0.5 rounded">{fmtDate(nd)} ⚠️</span>
                                : nd <= soon
                                  ? <span className="inline-flex items-center gap-1 bg-yellow-50 text-yellow-700 text-xs font-semibold px-1.5 py-0.5 rounded">{fmtDate(nd)} ⏰</span>
                                  : <span className="text-slate-700 text-xs">{fmtDate(nd)}</span>
                          ); })()
                        ) : col.key === "lastCallDate" && rec.lastCallDate ? (
                          <span className="text-slate-700 text-xs">{fmtDate(normDate(rec.lastCallDate))}</span>
                        ) : col.key === "memo" ? (
                          <span className="text-slate-600 text-xs block max-w-56 truncate" title={rec.memo||""}>
                            {rec.memo || "—"}
                          </span>
                        ) : (
                          <span className="text-slate-700 text-xs">{rec[col.key] || "—"}</span>
                        )}
                      </td>
                    ))}
                    <td className="px-3 py-2.5 whitespace-nowrap">
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
          </div>

          {/* ── Pagination ── */}
          <div className="px-4 py-3 border-t border-slate-100 flex items-center justify-between gap-4 flex-wrap">
            <span className="text-xs text-slate-400">
              {sortedFiltered.length > 0
                ? `${sortedFiltered.length}件中 ${(page-1)*PAGE_SIZE+1}–${Math.min(page*PAGE_SIZE, sortedFiltered.length)} 件表示（全${records.length}件）`
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
        <SettingsModal settings={settings} onSave={s => setSettings(s)} onClose={() => setShowSettings(false)} />
      )}
      {showImport && (
        <ImportModal onImport={addRecords} onClose={() => setShowImport(false)} />
      )}
      {showDupe && (
        <DuplicateModal records={records} onClean={cleanDuplicates} onClose={() => setShowDupe(false)} />
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
        />
      )}
    </div>
  );
}
