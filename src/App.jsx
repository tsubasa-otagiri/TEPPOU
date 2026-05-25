import { useState, useEffect, useRef, useCallback } from "react";

// ── Constants ──────────────────────────────────────────────────────────────────
const PASSWORD   = "1111";
const STORAGE_KEY  = "teppou_records_v2";
const SETTINGS_KEY = "teppou_settings_v1";
const PAGE_SIZE  = 100;

const IS_MEMBERS = [
  "櫻井　肇","上浦　諒大","井上　妃音","太田　小百合","十文字　菜月",
  "中　翔吾","早坂　直樹","小田切　翼","横井　優一","青木　大輔",
];

const STATUS_CFG = {
  "未架電":            { bg:"bg-slate-100",  text:"text-slate-600",  border:"border-slate-300",  dot:"bg-slate-400"  },
  "不通":              { bg:"bg-slate-200",  text:"text-slate-700",  border:"border-slate-400",  dot:"bg-slate-500"  },
  "不在":              { bg:"bg-orange-100", text:"text-orange-700", border:"border-orange-300", dot:"bg-orange-400" },
  "受付断り":          { bg:"bg-rose-100",   text:"text-rose-700",   border:"border-rose-300",   dot:"bg-rose-500"   },
  "担当コネクト":      { bg:"bg-sky-100",    text:"text-sky-700",    border:"border-sky-300",    dot:"bg-sky-500"    },
  "アポイント獲得商談":{ bg:"bg-teal-100",   text:"text-teal-700",   border:"border-teal-300",   dot:"bg-teal-500"   },
  "商談中":            { bg:"bg-blue-100",   text:"text-blue-700",   border:"border-blue-300",   dot:"bg-blue-500"   },
  "成約":              { bg:"bg-green-100",  text:"text-green-700",  border:"border-green-300",  dot:"bg-green-500"  },
  "失注":              { bg:"bg-red-100",    text:"text-red-700",    border:"border-red-300",    dot:"bg-red-500"    },
  "保留":              { bg:"bg-yellow-100", text:"text-yellow-700", border:"border-yellow-300", dot:"bg-yellow-500" },
  "折り返し待ち":      { bg:"bg-purple-100", text:"text-purple-700", border:"border-purple-300", dot:"bg-purple-500" },
};

const ALL_COLUMNS = [
  { key:"companyName",  label:"企業名",        required:true  },
  { key:"phone",        label:"電話番号",       required:false },
  { key:"email",        label:"メールアドレス", required:false },
  { key:"url",          label:"GBP/URL",        required:false },
  { key:"status",       label:"ステータス",     required:false },
  { key:"assignee",     label:"担当者",         required:false },
  { key:"lastCallDate", label:"最終架電日",     required:false },
  { key:"nextCallDate", label:"次回架電日",     required:false },
  { key:"callCount",    label:"架電回数",       required:false },
  { key:"memo",         label:"メモ/情報",      required:false },
];

// ── Helpers ────────────────────────────────────────────────────────────────────
function parseCSV(text) {
  return text.trim().split(/\r?\n/).map(line => {
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
    if      (/企業名|会社名|法人名/.test(n))              m.companyName  = i;
    else if (/電話|tel|phone/.test(n))                    m.phone        = i;
    else if (/メール|mail|email/.test(n))                 m.email        = i;
    else if (/url|サイト|gbp|ホームページ/.test(n))        m.url          = i;
    else if (/メモ|備考|note|コメント|情報/.test(n))        m.memo         = i;
    else if (/担当者?名?|営業担当/.test(n))                m.assignee     = i;
    else if (/ステータス|状態/.test(n))                    m.status       = i;
    else if (/次回架電|次架電|コールバック/.test(n))        m.nextCallDate = i;
    else if (/最終架電|架電日/.test(n))                    m.lastCallDate = i;
    else if (/架電回数|コール回数/.test(n))                m.callCount    = i;
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

function normName(s) { return String(s||"").replace(/[\s　]/g,"").toLowerCase(); }
function genId()     { return Date.now().toString(36)+Math.random().toString(36).slice(2); }
function getToday()  { return new Date().toISOString().slice(0,10); }
function nowIso()    { return new Date().toISOString(); }

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

// ── ImportModal ────────────────────────────────────────────────────────────────
function ImportModal({ onImport, onClose }) {
  const [mode, setMode] = useState("sales");
  const [log,  setLog]  = useState(null);
  const fileRef = useRef();

  const handleFile = e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const rows = parseCSV(ev.target.result);
        const result = mode === "sales" ? doImportSales(rows) : doImportMetel(rows);
        setLog(result);
        if (result.records && result.records.length > 0) onImport(result.records);
      } catch (ex) {
        setLog({ error: `インポートエラー: ${ex.message}`, records: [] });
      }
    };
    reader.readAsText(file, "UTF-8");
    e.target.value = "";
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
      records.push({
        id:           genId(),
        companyName:  company.trim(),
        phone:        map.phone        !== undefined ? (row[map.phone]        ||"").trim() : "",
        email:        map.email        !== undefined ? (row[map.email]        ||"").trim() : "",
        url:          map.url          !== undefined ? (row[map.url]          ||"").trim() : "",
        status:      (map.status       !== undefined ? (row[map.status]       ||"")        : "") || "未架電",
        assignee:     map.assignee     !== undefined ? (row[map.assignee]     ||"").trim() : "",
        memo:         map.memo         !== undefined ? (row[map.memo]         ||"").trim() : "",
        nextCallDate: map.nextCallDate  !== undefined ? (row[map.nextCallDate] ||"").trim() : "",
        lastCallDate: map.lastCallDate  !== undefined ? (row[map.lastCallDate] ||"").trim() : "",
        callCount:    map.callCount     !== undefined ? parseInt(row[map.callCount])||0      : 0,
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
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg p-6">
        <div className="flex items-center justify-between mb-5">
          <h2 className="text-base font-bold text-slate-800">CSVインポート</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
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
              onClick={() => setMode(opt.value)}>
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

        {/* Drop zone */}
        <label className="flex flex-col items-center justify-center gap-3 border-2 border-dashed border-slate-300 hover:border-blue-400 rounded-xl p-8 cursor-pointer transition-colors mb-4 bg-slate-50 hover:bg-blue-50">
          <svg className="w-9 h-9 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5}
              d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
          </svg>
          <span className="text-sm text-slate-600 font-medium">CSVファイルを選択</span>
          <span className="text-xs text-slate-400">文字コード: UTF-8 推奨</span>
          <input ref={fileRef} type="file" accept=".csv" className="hidden" onChange={handleFile} />
        </label>

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
function DuplicateModal({ records, onClean, onClose }) {
  const groups = (() => {
    const g = {};
    records.forEach(r => {
      const key = normName(r.companyName);
      if (!key) return;
      (g[key] = g[key]||[]).push(r);
    });
    return Object.values(g)
      .filter(rs => rs.length > 1)
      .map(rs => rs.sort((a,b) => new Date(b.updatedAt||b.importedAt) - new Date(a.updatedAt||a.importedAt)));
  })();

  const totalDel = groups.reduce((s,g) => s+g.length-1, 0);

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl p-6 flex flex-col max-h-[80vh]">
        <div className="flex items-center justify-between mb-4">
          <h2 className="text-base font-bold text-slate-800">重複クレンジング</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
        </div>
        <p className="text-sm text-slate-500 mb-4">
          {groups.length === 0
            ? "✅ 重複データは見つかりませんでした。"
            : `${groups.length}グループ、${totalDel}件の古いデータを削除します（最新1件を残します）。`}
        </p>
        {groups.length > 0 && (
          <div className="overflow-y-auto flex-1 space-y-2 mb-4 pr-1">
            {groups.map((g, gi) => (
              <div key={gi} className="border border-slate-200 rounded-xl p-3">
                <div className="font-semibold text-sm text-slate-700 mb-2">{g[0].companyName}</div>
                <div className="space-y-1">
                  {g.map((r, i) => (
                    <div key={r.id} className={`flex flex-wrap items-center gap-2 text-xs
                      ${i===0 ? "text-green-700" : "text-slate-400 line-through"}`}>
                      <span className="font-medium shrink-0">{i===0 ? "✓ 残す" : "× 削除"}</span>
                      <span>{(r.updatedAt||r.importedAt||"").slice(0,10)}</span>
                      <span>{r.phone||"—"}</span>
                      <StatusBadge status={r.status} />
                      {r.assignee && <span>{r.assignee}</span>}
                    </div>
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
            <button onClick={() => { onClean(groups); onClose(); }}
              className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white text-sm font-semibold rounded-lg transition-colors">
              {totalDel}件の重複を削除する
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// ── RecordFormModal (shared by New & Edit) ─────────────────────────────────────
function RecordFormModal({ initial, title, onSave, onClose }) {
  const [form, setForm] = useState({
    companyName:"", phone:"", email:"", url:"",
    status:"未架電", assignee:"", nextCallDate:"", callCount:0, memo:"",
    ...initial,
  });
  const upd = (k,v) => setForm(f => ({ ...f, [k]:v }));

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg p-6">
        <div className="flex items-center justify-between mb-5">
          <h2 className="text-base font-bold text-slate-800">{title}</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 text-xl leading-none">×</button>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div className="col-span-2">
            <label className="block text-xs text-slate-500 mb-1">企業名 <span className="text-rose-500">*</span></label>
            <input type="text" value={form.companyName}
              onChange={e => upd("companyName", e.target.value)}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <div>
            <label className="block text-xs text-slate-500 mb-1">電話番号</label>
            <input type="text" value={form.phone||""} onChange={e => upd("phone", e.target.value)}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <div>
            <label className="block text-xs text-slate-500 mb-1">メールアドレス</label>
            <input type="email" value={form.email||""} onChange={e => upd("email", e.target.value)}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <div className="col-span-2">
            <label className="block text-xs text-slate-500 mb-1">GBP/URL</label>
            <input type="text" value={form.url||""} onChange={e => upd("url", e.target.value)}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <div>
            <label className="block text-xs text-slate-500 mb-1">ステータス</label>
            <select value={form.status||"未架電"} onChange={e => upd("status", e.target.value)}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
              {Object.keys(STATUS_CFG).map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-xs text-slate-500 mb-1">担当者</label>
            <input type="text" value={form.assignee||""} onChange={e => upd("assignee", e.target.value)}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <div>
            <label className="block text-xs text-slate-500 mb-1">次回架電日</label>
            <input type="date" value={form.nextCallDate||""} onChange={e => upd("nextCallDate", e.target.value)}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <div>
            <label className="block text-xs text-slate-500 mb-1">架電回数</label>
            <input type="number" min="0" value={form.callCount||0}
              onChange={e => upd("callCount", parseInt(e.target.value)||0)}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <div className="col-span-2">
            <label className="block text-xs text-slate-500 mb-1">メモ/情報</label>
            <textarea value={form.memo||""} onChange={e => upd("memo", e.target.value)} rows={3}
              className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none" />
          </div>
        </div>
        <div className="flex gap-2 justify-end mt-4 pt-4 border-t border-slate-100">
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
  );
}

// ── Main App ───────────────────────────────────────────────────────────────────
export default function App() {
  const [loggedIn,       setLoggedIn]       = useState(() => sessionStorage.getItem("teppou_auth")==="1");
  const [records,        setRecords]        = useState([]);
  const [settings,       setSettings]       = useState({ logo:null, favicon:null });
  const [search,         setSearch]         = useState("");
  const [statusFilter,   setStatusFilter]   = useState("all");
  const [assigneeFilter, setAssigneeFilter] = useState("all");
  const [visibleCols,    setVisibleCols]    = useState(ALL_COLUMNS.map(c => c.key));
  const [showColDrop,    setShowColDrop]    = useState(false);
  const [page,           setPage]           = useState(1);
  const [showSettings,   setShowSettings]   = useState(false);
  const [showImport,     setShowImport]     = useState(false);
  const [showDupe,       setShowDupe]       = useState(false);
  const [showNew,        setShowNew]        = useState(false);
  const [editRec,        setEditRec]        = useState(null);
  const [selected,       setSelected]       = useState(new Set());
  const colDropRef = useRef();

  // ── Persistence ──────────────────────────────────────────────────────────────
  useEffect(() => {
    try { const s = localStorage.getItem(STORAGE_KEY);  if (s) setRecords(JSON.parse(s));  } catch {}
    try { const s = localStorage.getItem(SETTINGS_KEY); if (s) setSettings(JSON.parse(s)); } catch {}
  }, []);

  useEffect(() => { localStorage.setItem(STORAGE_KEY,  JSON.stringify(records));  }, [records]);
  useEffect(() => { localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); }, [settings]);

  // ── Favicon ───────────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!settings.favicon) return;
    let link = document.querySelector("link[rel~='icon']");
    if (!link) { link = document.createElement("link"); link.rel = "icon"; document.head.appendChild(link); }
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

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const paginated  = filtered.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);

  const statsMap = {};
  records.forEach(r => { statsMap[r.status] = (statsMap[r.status]||0) + 1; });
  const stats = Object.entries(statsMap).map(([s,c]) => ({ status:s, count:c, ...(STATUS_CFG[s]??{}) }));

  const alerts   = records.filter(r => r.nextCallDate && r.nextCallDate <= today && r.status !== "成約" && r.status !== "失注");
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

  const cleanDuplicates = useCallback(groups => {
    const del = new Set();
    groups.forEach(g => g.slice(1).forEach(r => del.add(r.id)));
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

      <div className="max-w-screen-2xl mx-auto px-4 py-5 space-y-4">

        {/* ── Alert bar ── */}
        {alerts.length > 0 && (
          <div className="bg-amber-50 border border-amber-300 rounded-xl px-4 py-3 flex flex-wrap items-center gap-2">
            <span className="text-sm font-semibold text-amber-700 shrink-0">📅 次回架電日アラート</span>
            {alerts.slice(0,5).map(r => (
              <span key={r.id} className="bg-amber-100 border border-amber-300 text-amber-800 text-xs px-2 py-0.5 rounded-full">
                {r.companyName}（{r.nextCallDate}）
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
                    <th key={col.key} className="px-3 py-3 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">
                      {col.label}
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
                ) : paginated.map(rec => (
                  <tr key={rec.id}
                    className={`hover:bg-slate-50/60 transition-colors ${selected.has(rec.id) ? "bg-blue-50" : ""}`}>
                    <td className="px-3 py-2.5">
                      <input type="checkbox" checked={selected.has(rec.id)}
                        onChange={e => toggleSelect(rec.id, e.target.checked)}
                        className="rounded border-slate-300 text-blue-600" />
                    </td>
                    {visibleDefs.map(col => (
                      <td key={col.key} className="px-3 py-2.5 whitespace-nowrap max-w-xs">
                        {col.key === "status" ? (
                          <StatusBadge status={rec.status} />
                        ) : col.key === "url" && rec.url ? (
                          <a href={rec.url} target="_blank" rel="noreferrer"
                            className="text-blue-600 hover:underline text-xs block max-w-36 truncate">
                            {rec.url}
                          </a>
                        ) : col.key === "nextCallDate" && rec.nextCallDate && rec.nextCallDate <= today ? (
                          <span className="text-amber-600 font-semibold text-xs">{rec.nextCallDate} ⚠</span>
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
                        className="text-xs text-blue-600 hover:text-blue-800 font-medium mr-2">
                        編集
                      </button>
                      <button onClick={() => deleteRecord(rec.id)}
                        className="text-xs text-rose-500 hover:text-rose-700 font-medium">
                        削除
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* ── Pagination ── */}
          <div className="px-4 py-3 border-t border-slate-100 flex items-center justify-between gap-4 flex-wrap">
            <span className="text-xs text-slate-400">
              {filtered.length > 0
                ? `${filtered.length}件中 ${(page-1)*PAGE_SIZE+1}–${Math.min(page*PAGE_SIZE, filtered.length)} 件表示（全${records.length}件）`
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
        />
      )}
    </div>
  );
}
