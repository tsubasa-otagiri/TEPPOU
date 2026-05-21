import { useState, useEffect, useRef } from "react";

// ─── 定数 ─────────────────────────────────────────────────────────────────────
const PASSWORD = "1111";
const STORAGE_KEY = "sales_mgmt_v1";

const IS_MEMBERS = [
  "櫻井　肇", "上浦　諒大", "井上　妃音", "太田　小百合", "十文字　菜月",
  "中　翔吾", "早坂　直樹", "小田切　翼", "横井　優一", "青木　大輔",
];

const STATUS_CFG = {
  "未架電":            { bg: "bg-slate-100",  text: "text-slate-600",  border: "border-slate-300",  dot: "bg-slate-400"  },
  "不通":              { bg: "bg-gray-200",   text: "text-gray-600",   border: "border-gray-400",   dot: "bg-gray-500"   },
  "不在":              { bg: "bg-orange-100", text: "text-orange-700", border: "border-orange-300", dot: "bg-orange-400" },
  "受付断り":          { bg: "bg-rose-100",   text: "text-rose-700",   border: "border-rose-300",   dot: "bg-rose-500"   },
  "担当コネクト":      { bg: "bg-sky-100",    text: "text-sky-700",    border: "border-sky-300",    dot: "bg-sky-500"    },
  "アポイント獲得商談":{ bg: "bg-teal-100",   text: "text-teal-700",   border: "border-teal-300",   dot: "bg-teal-500"   },
  "商談中":            { bg: "bg-blue-100",   text: "text-blue-700",   border: "border-blue-300",   dot: "bg-blue-500"   },
  "成約":              { bg: "bg-green-100",  text: "text-green-700",  border: "border-green-300",  dot: "bg-green-500"  },
  "失注":              { bg: "bg-red-100",    text: "text-red-700",    border: "border-red-300",    dot: "bg-red-500"    },
  "保留":              { bg: "bg-yellow-100", text: "text-yellow-700", border: "border-yellow-300", dot: "bg-yellow-500" },
  "折り返し待ち":      { bg: "bg-purple-100", text: "text-purple-700", border: "border-purple-300", dot: "bg-purple-500" },
};

const ALL_COLUMNS = [
  { key: "companyName",  label: "企業名",        required: true  },
  { key: "phone",        label: "電話番号",       required: false },
  { key: "email",        label: "メールアドレス", required: false },
  { key: "url",          label: "GBP/URL",        required: false },
  { key: "status",       label: "ステータス",     required: false },
  { key: "assignee",     label: "担当者",         required: false },
  { key: "lastCallDate", label: "最終架電日",     required: false },
  { key: "nextCallDate", label: "次回架電日",     required: false },
  { key: "callCount",    label: "架電回数",       required: false },
  { key: "memo",         label: "メモ",           required: false },
];

// ─── CSV ユーティリティ ────────────────────────────────────────────────────────
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  return lines.map((line) => {
    const cells = [];
    let inQ = false, cell = "";
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQ && line[i + 1] === '"') { cell += '"'; i++; }
        else inQ = !inQ;
      } else if (ch === "," && !inQ) {
        cells.push(cell.trim()); cell = "";
      } else cell += ch;
    }
    cells.push(cell.trim());
    return cells;
  });
}

// 1行目が集計行（数字が40%以上）かどうか判定
function isAggregateRow(row) {
  const nonEmpty = row.filter((c) => c.trim() !== "");
  if (nonEmpty.length === 0) return false;
  const numCount = nonEmpty.filter((c) => /^\d[\d,.]*$/.test(c.trim())).length;
  return numCount / nonEmpty.length >= 0.4;
}

// ヘッダー → フィールドキーのマッピング
function mapSalesHeaders(headers) {
  const m = {};
  headers.forEach((h, i) => {
    const n = h.replace(/[\s　]/g, "").toLowerCase();
    if      (/企業名|会社名|法人名|companyname/.test(n))        m.companyName  = i;
    else if (/電話|tel|phone/.test(n))                          m.phone        = i;
    else if (/メール|mail|email/.test(n))                       m.email        = i;
    else if (/url|サイト|gbp|ホームページ|website/.test(n))      m.url          = i;
    else if (/メモ|備考|note|コメント/.test(n))                  m.memo         = i;
    else if (/担当者?名?|営業担当/.test(n))                      m.assignee     = i;
    else if (/ステータス|状態|status/.test(n))                   m.status       = i;
    else if (/次回架電|次架電|コールバック/.test(n))              m.nextCallDate = i;
    else if (/最終架電|架電日|lastcall/.test(n))                 m.lastCallDate = i;
    else if (/架電回数|コール回数/.test(n))                      m.callCount    = i;
  });
  return m;
}

// ミーてる: タグ+通話種別 → ステータス変換（優先順位順）
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

function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2);
}

function normName(s) {
  return String(s || "").replace(/[\s　]/g, "").toLowerCase();
}

// ─── StatusBadge ─────────────────────────────────────────────────────────────
function StatusBadge({ status }) {
  const c = STATUS_CFG[status] ?? { bg: "bg-gray-100", text: "text-gray-600", border: "border-gray-300" };
  return (
    <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium border whitespace-nowrap ${c.bg} ${c.text} ${c.border}`}>
      {status || "—"}
    </span>
  );
}

// ─── LoginScreen ──────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  const submit = (e) => {
    e.preventDefault();
    if (pw === PASSWORD) { sessionStorage.setItem("sales_auth", "1"); onLogin(); }
    else { setErr("パスワードが正しくありません"); setPw(""); }
  };
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-800 to-slate-900 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl p-10 w-full max-w-sm">
        <div className="text-center mb-8">
          <div className="w-16 h-16 bg-blue-600 rounded-2xl flex items-center justify-center mx-auto mb-4">
            <svg className="w-8 h-8 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4" />
            </svg>
          </div>
          <h1 className="text-2xl font-bold text-slate-800">営業管理システム</h1>
          <p className="text-slate-500 text-sm mt-1">GMOテック IS部門</p>
        </div>
        <form onSubmit={submit} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">パスワード</label>
            <input type="password" value={pw} autoFocus
              onChange={(e) => { setPw(e.target.value); setErr(""); }}
              className="w-full border border-slate-300 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="パスワードを入力"
            />
            {err && <p className="text-rose-600 text-xs mt-1">{err}</p>}
          </div>
          <button type="submit" className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold rounded-lg py-2.5 text-sm transition-colors">
            ログイン
          </button>
        </form>
      </div>
    </div>
  );
}

// ─── App (main) ───────────────────────────────────────────────────────────────
export default function App() {
  const [loggedIn, setLoggedIn]           = useState(() => sessionStorage.getItem("sales_auth") === "1");
  const [records, setRecords]             = useState([]);
  const [search, setSearch]               = useState("");
  const [statusFilter, setStatusFilter]   = useState("all");
  const [assigneeFilter, setAssigneeFilter] = useState("all");
  const [visibleCols, setVisibleCols]     = useState(ALL_COLUMNS.map((c) => c.key));
  const [showColModal, setShowColModal]   = useState(false);
  const [importMode, setImportMode]       = useState("sales");
  const [importLog, setImportLog]         = useState(null);
  const [editingId, setEditingId]         = useState(null);
  const [editForm, setEditForm]           = useState({});
  const [addingNew, setAddingNew]         = useState(false);
  const [newForm, setNewForm]             = useState(emptyForm());
  const [showDupeModal, setShowDupeModal] = useState(false);
  const [dupeGroups, setDupeGroups]       = useState([]);
  const [selected, setSelected]           = useState(new Set());
  const fileRef = useRef();

  function emptyForm() {
    return { companyName: "", phone: "", email: "", url: "", status: "未架電", assignee: "", nextCallDate: "", memo: "" };
  }

  // localStorage 永続化
  useEffect(() => {
    try { const s = localStorage.getItem(STORAGE_KEY); if (s) setRecords(JSON.parse(s)); } catch {}
  }, []);
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
  }, [records]);

  const todayStr = new Date().toISOString().slice(0, 10);

  // フィルタリング
  const filtered = records.filter((r) => {
    if (statusFilter !== "all" && r.status !== statusFilter) return false;
    if (assigneeFilter !== "all" && r.assignee !== assigneeFilter) return false;
    if (search) {
      const q = search;
      if (
        !(r.companyName || "").includes(q) &&
        !(r.phone       || "").includes(q) &&
        !(r.assignee    || "").includes(q) &&
        !(r.memo        || "").includes(q) &&
        !(r.email       || "").includes(q)
      ) return false;
    }
    return true;
  });

  // 架電アラート
  const alerts = records.filter(
    (r) => r.nextCallDate && r.nextCallDate <= todayStr && r.status !== "成約" && r.status !== "失注"
  );

  // 担当者一覧
  const assignees = [...new Set(records.map((r) => r.assignee).filter(Boolean))];

  // ステータス集計
  const statsMap = {};
  records.forEach((r) => { statsMap[r.status] = (statsMap[r.status] || 0) + 1; });
  const stats = Object.entries(statsMap).map(([s, c]) => ({ status: s, count: c, ...(STATUS_CFG[s] ?? {}) }));

  // ── CSVインポート ──────────────────────────────────────────────────────────
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const rows = parseCSV(ev.target.result);
        importMode === "sales" ? importSales(rows) : importMetel(rows);
      } catch (ex) {
        setImportLog({ error: `インポートエラー: ${ex.message}` });
      }
    };
    reader.readAsText(file, "UTF-8");
    e.target.value = "";
  };

  function importSales(rows) {
    if (rows.length < 2) { setImportLog({ error: "データが不足しています" }); return; }

    // 1行目が集計行なら2行目をヘッダーとして使う
    let headerIdx = 0;
    if (isAggregateRow(rows[0])) headerIdx = 1;

    const headers = rows[headerIdx];
    const map = mapSalesHeaders(headers);
    const dataRows = rows.slice(headerIdx + 1);

    const added = [];
    let skipped = 0;
    for (const row of dataRows) {
      if (row.every((c) => !c.trim())) continue;
      const company = map.companyName !== undefined ? row[map.companyName] : "";
      if (!company.trim()) { skipped++; continue; }
      added.push({
        id: generateId(),
        companyName:  company.trim(),
        phone:        map.phone        !== undefined ? row[map.phone]        : "",
        email:        map.email        !== undefined ? row[map.email]        : "",
        url:          map.url          !== undefined ? row[map.url]          : "",
        status:      (map.status       !== undefined ? row[map.status]       : "") || "未架電",
        assignee:     map.assignee     !== undefined ? row[map.assignee]     : "",
        memo:         map.memo         !== undefined ? row[map.memo]         : "",
        nextCallDate: map.nextCallDate  !== undefined ? row[map.nextCallDate] : "",
        lastCallDate: map.lastCallDate  !== undefined ? row[map.lastCallDate] : "",
        callCount:    map.callCount     !== undefined ? parseInt(row[map.callCount]) || 0 : 0,
        importedAt: new Date().toISOString(),
        source: "csv",
      });
    }
    setRecords((prev) => [...prev, ...added]);
    setImportLog({ success: true, count: added.length, skipped, autoSkip: headerIdx === 1 });
  }

  function importMetel(rows) {
    if (rows.length < 2) { setImportLog({ error: "データが不足しています" }); return; }

    const headers = rows[0];
    const col = (patterns) => {
      for (let i = 0; i < headers.length; i++) {
        const h = headers[i].replace(/[\s　]/g, "");
        if (patterns.some((p) => h.includes(p))) return i;
      }
      return -1;
    };

    const cAssignee = col(["担当者", "オペレーター", "エージェント", "架電者"]);
    const cCompany  = col(["企業名", "会社名", "顧客名", "取引先名"]);
    const cPhone    = col(["電話番号", "TEL", "電話"]);
    const cCallType = col(["通話種別", "架電種別", "種別", "通話タイプ"]);
    const cTags     = col(["タグ", "ラベル"]);
    const cDate     = col(["架電日時", "通話日時", "日時", "日付", "架電日"]);
    const cMemo     = col(["メモ", "備考", "コメント", "会話内容"]);

    let added = 0, filtered = 0, skipped = 0;
    const newRecs = [];

    for (const row of rows.slice(1)) {
      if (row.every((c) => !c.trim())) continue;

      // ISメンバーフィルタ（全角半角スペース正規化）
      const assignee     = cAssignee >= 0 ? (row[cAssignee] ?? "").trim() : "";
      const assigneeNorm = normName(assignee);
      const isMember     = IS_MEMBERS.some((m) => {
        const mn = normName(m);
        return assigneeNorm === mn || assigneeNorm.includes(mn) || mn.includes(assigneeNorm);
      });
      if (!isMember) { filtered++; continue; }

      const company = cCompany >= 0 ? (row[cCompany] ?? "").trim() : "";
      if (!company) { skipped++; continue; }

      const callType = cCallType >= 0 ? (row[cCallType] ?? "").trim() : "";
      const tags     = cTags     >= 0 ? (row[cTags]     ?? "").trim() : "";
      const status   = convertMitelStatus(callType, tags) ?? "不通";
      const rawDate  = cDate     >= 0 ? (row[cDate]     ?? "").trim() : "";
      const dateStr  = rawDate ? rawDate.slice(0, 10).replace(/\//g, "-") : "";

      newRecs.push({
        id: generateId(),
        companyName:  company,
        phone:        cPhone >= 0 ? (row[cPhone] ?? "").trim() : "",
        email: "", url: "",
        status, assignee,
        lastCallDate: dateStr,
        nextCallDate: "",
        callCount: 1,
        memo: cMemo >= 0 ? (row[cMemo] ?? "").trim() : "",
        importedAt: new Date().toISOString(),
        source: "metel",
      });
      added++;
    }
    setRecords((prev) => [...prev, ...newRecs]);
    setImportLog({ success: true, count: added, filtered, skipped });
  }

  // ── 重複クレンジング ───────────────────────────────────────────────────────
  const detectDupes = () => {
    const groups = {};
    records.forEach((r) => {
      const key = normName(r.companyName);
      if (!key) return;
      (groups[key] = groups[key] || []).push(r);
    });
    const dupes = Object.entries(groups)
      .filter(([, rs]) => rs.length > 1)
      .map(([, rs]) => ({ records: rs.sort((a, b) => new Date(b.importedAt) - new Date(a.importedAt)) }));
    setDupeGroups(dupes);
    setShowDupeModal(true);
  };

  const cleanDupes = () => {
    const toDelete = new Set();
    dupeGroups.forEach(({ records: rs }) => rs.slice(1).forEach((r) => toDelete.add(r.id)));
    setRecords((prev) => prev.filter((r) => !toDelete.has(r.id)));
    setShowDupeModal(false);
    setDupeGroups([]);
  };

  // ── CRUD ──────────────────────────────────────────────────────────────────
  const startEdit = (rec) => { setEditingId(rec.id); setEditForm({ ...rec }); };
  const saveEdit  = () => {
    setRecords((prev) => prev.map((r) => r.id === editingId ? { ...r, ...editForm } : r));
    setEditingId(null);
  };
  const deleteOne = (id) => {
    if (window.confirm("このレコードを削除しますか？"))
      setRecords((prev) => prev.filter((r) => r.id !== id));
  };
  const deleteSelected = () => {
    if (!selected.size) return;
    if (window.confirm(`選択した ${selected.size} 件を削除しますか？`)) {
      setRecords((prev) => prev.filter((r) => !selected.has(r.id)));
      setSelected(new Set());
    }
  };
  const saveNew = () => {
    if (!newForm.companyName.trim()) return;
    setRecords((prev) => [...prev, { ...newForm, id: generateId(), importedAt: new Date().toISOString(), source: "manual", callCount: 0 }]);
    setAddingNew(false);
    setNewForm(emptyForm());
  };

  // ─────────────────────────────────────────────────────────────────────────
  if (!loggedIn) return <LoginScreen onLogin={() => setLoggedIn(true)} />;

  const visibleDefs = ALL_COLUMNS.filter((c) => visibleCols.includes(c.key));

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800">

      {/* ヘッダー */}
      <header className="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-30">
        <div className="max-w-screen-2xl mx-auto px-4 py-3 flex items-center justify-between gap-4 flex-wrap">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 bg-blue-600 rounded-xl flex items-center justify-center shrink-0">
              <svg className="w-5 h-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                  d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
              </svg>
            </div>
            <div>
              <div className="font-bold text-slate-800 text-base leading-tight">営業管理システム</div>
              <div className="text-xs text-slate-400">GMOテック IS部門</div>
            </div>
          </div>
          <div className="flex items-center gap-3 flex-wrap">
            <span className="text-sm text-slate-500">
              総件数: <span className="font-bold text-blue-600">{records.length}</span>件
            </span>
            {alerts.length > 0 && (
              <span className="bg-amber-100 text-amber-700 text-xs font-semibold px-3 py-1 rounded-full border border-amber-300">
                ⏰ 架電アラート {alerts.length}件
              </span>
            )}
            <button
              onClick={() => { sessionStorage.removeItem("sales_auth"); setLoggedIn(false); }}
              className="text-xs text-slate-400 hover:text-slate-700 border border-slate-200 px-3 py-1.5 rounded-lg hover:bg-slate-50 transition-colors"
            >
              ログアウト
            </button>
          </div>
        </div>
      </header>

      <div className="max-w-screen-2xl mx-auto px-4 py-5 space-y-4">

        {/* 架電アラートバー */}
        {alerts.length > 0 && (
          <div className="bg-amber-50 border border-amber-300 rounded-xl p-3 flex flex-wrap items-center gap-2">
            <span className="text-sm font-semibold text-amber-700 shrink-0">次回架電日アラート:</span>
            {alerts.slice(0, 6).map((r) => (
              <span key={r.id} className="bg-amber-100 border border-amber-300 text-amber-800 text-xs px-2 py-0.5 rounded-full">
                {r.companyName}（{r.nextCallDate}）
              </span>
            ))}
            {alerts.length > 6 && <span className="text-xs text-amber-600">他 {alerts.length - 6} 件</span>}
          </div>
        )}

        {/* ステータス集計バー */}
        {stats.length > 0 && (
          <div className="bg-white rounded-xl border border-slate-200 p-4">
            <p className="text-xs font-semibold text-slate-500 mb-3 uppercase tracking-wide">ステータス別集計</p>
            <div className="flex flex-wrap gap-2">
              {stats.map((s) => (
                <button key={s.status}
                  onClick={() => setStatusFilter(statusFilter === s.status ? "all" : s.status)}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg border text-xs font-medium transition-all
                    ${statusFilter === s.status
                      ? `${s.bg ?? "bg-gray-100"} ${s.text ?? "text-gray-600"} ${s.border ?? "border-gray-300"} ring-2 ring-offset-1 ring-blue-400`
                      : `${s.bg ?? "bg-gray-100"} ${s.text ?? "text-gray-600"} ${s.border ?? "border-gray-300"} hover:opacity-80`}`}
                >
                  <span className={`w-2 h-2 rounded-full ${s.dot ?? "bg-gray-400"}`} />
                  {s.status}
                  <span className="font-bold">{s.count}</span>
                </button>
              ))}
              {statusFilter !== "all" && (
                <button onClick={() => setStatusFilter("all")}
                  className="px-3 py-1.5 text-xs border border-slate-300 rounded-lg text-slate-500 hover:bg-slate-50">
                  全表示
                </button>
              )}
            </div>
          </div>
        )}

        {/* ツールバー */}
        <div className="bg-white rounded-xl border border-slate-200 p-4 space-y-3">

          {/* アクション群 */}
          <div className="flex flex-wrap items-center gap-2">

            {/* インポートモードトグル */}
            <div className="flex rounded-lg border border-slate-300 overflow-hidden text-xs font-medium shrink-0">
              <button onClick={() => setImportMode("sales")}
                className={`px-3 py-2 transition-colors ${importMode === "sales" ? "bg-blue-600 text-white" : "bg-white text-slate-600 hover:bg-slate-50"}`}>
                自分の営業リスト
              </button>
              <button onClick={() => setImportMode("metel")}
                className={`px-3 py-2 border-l border-slate-300 transition-colors ${importMode === "metel" ? "bg-blue-600 text-white" : "bg-white text-slate-600 hover:bg-slate-50"}`}>
                ミーてる架電ログ
              </button>
            </div>

            {/* CSVインポート */}
            <label className="flex items-center gap-1.5 bg-blue-50 hover:bg-blue-100 border border-blue-300 text-blue-700 px-3 py-2 rounded-lg text-xs font-medium cursor-pointer transition-colors shrink-0">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                  d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
              </svg>
              CSVインポート（{importMode === "sales" ? "営業リスト" : "ミーてる"}）
              <input ref={fileRef} type="file" accept=".csv" className="hidden" onChange={handleFile} />
            </label>

            {/* 重複クレンジング */}
            <button onClick={detectDupes}
              className="flex items-center gap-1.5 bg-amber-50 hover:bg-amber-100 border border-amber-300 text-amber-700 px-3 py-2 rounded-lg text-xs font-medium transition-colors shrink-0">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                  d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" />
              </svg>
              重複クレンジング
            </button>

            {/* 列設定 */}
            <button onClick={() => setShowColModal(true)}
              className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-300 text-slate-600 px-3 py-2 rounded-lg text-xs font-medium transition-colors shrink-0">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 10h18M3 14h18M10 3v18M14 3v18" />
              </svg>
              列設定
            </button>

            {/* 新規追加 */}
            <button onClick={() => setAddingNew(true)}
              className="flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-xs font-semibold transition-colors shrink-0">
              <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
              </svg>
              新規追加
            </button>

            {/* 選択削除 */}
            {selected.size > 0 && (
              <button onClick={deleteSelected}
                className="flex items-center gap-1.5 bg-red-50 hover:bg-red-100 border border-red-300 text-red-700 px-3 py-2 rounded-lg text-xs font-medium transition-colors shrink-0">
                選択削除（{selected.size}件）
              </button>
            )}
          </div>

          {/* インポートログ */}
          {importLog && (
            <div className={`flex items-start justify-between text-xs rounded-lg px-3 py-2 ${
              importLog.error ? "bg-red-50 border border-red-200 text-red-700" : "bg-green-50 border border-green-200 text-green-700"
            }`}>
              <span>
                {importLog.error ? importLog.error : (
                  <>
                    インポート完了: <strong>{importLog.count}件</strong> 追加
                    {importLog.filtered  > 0 && `　／　ISメンバー以外: ${importLog.filtered}件 除外`}
                    {importLog.skipped   > 0 && `　／　スキップ: ${importLog.skipped}件`}
                    {importLog.autoSkip      && "　／　1行目を集計行として自動スキップ"}
                  </>
                )}
              </span>
              <button onClick={() => setImportLog(null)} className="ml-3 opacity-60 hover:opacity-100 text-base leading-none">×</button>
            </div>
          )}

          {/* 検索 & フィルタ */}
          <div className="flex flex-wrap gap-2">
            <div className="relative flex-1 min-w-48">
              <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
              </svg>
              <input type="text" value={search} onChange={(e) => setSearch(e.target.value)}
                placeholder="企業名・電話番号・担当者・メモで検索..."
                className="w-full pl-9 pr-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>
            <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}
              className="border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
              <option value="all">全ステータス</option>
              {Object.keys(STATUS_CFG).map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
            <select value={assigneeFilter} onChange={(e) => setAssigneeFilter(e.target.value)}
              className="border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
              <option value="all">全担当者</option>
              {assignees.map((a) => <option key={a} value={a}>{a}</option>)}
            </select>
          </div>
        </div>

        {/* データテーブル */}
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm border-collapse">
              <thead className="bg-slate-50 border-b border-slate-200">
                <tr>
                  <th className="w-10 px-3 py-3">
                    <input type="checkbox"
                      checked={selected.size === filtered.length && filtered.length > 0}
                      onChange={(e) => setSelected(e.target.checked ? new Set(filtered.map((r) => r.id)) : new Set())}
                      className="rounded border-slate-300 text-blue-600"
                    />
                  </th>
                  {visibleDefs.map((col) => (
                    <th key={col.key} className="px-3 py-3 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">
                      {col.label}
                    </th>
                  ))}
                  <th className="px-3 py-3 text-left text-xs font-semibold text-slate-500">操作</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filtered.length === 0 ? (
                  <tr>
                    <td colSpan={visibleDefs.length + 2} className="text-center py-16 text-slate-400 text-sm">
                      データがありません。CSVをインポートするか、新規追加してください。
                    </td>
                  </tr>
                ) : filtered.map((rec) => (
                  <tr key={rec.id} className={`hover:bg-slate-50/60 transition-colors ${selected.has(rec.id) ? "bg-blue-50" : ""}`}>
                    <td className="px-3 py-2.5">
                      <input type="checkbox" checked={selected.has(rec.id)}
                        onChange={(e) => {
                          const nx = new Set(selected);
                          e.target.checked ? nx.add(rec.id) : nx.delete(rec.id);
                          setSelected(nx);
                        }}
                        className="rounded border-slate-300 text-blue-600"
                      />
                    </td>

                    {editingId === rec.id ? (
                      <>
                        {visibleDefs.map((col) => (
                          <td key={col.key} className="px-2 py-1.5">
                            {col.key === "status" ? (
                              <select value={editForm.status || ""} onChange={(e) => setEditForm((f) => ({ ...f, status: e.target.value }))}
                                className="border border-slate-300 rounded px-2 py-1 text-xs w-full bg-white">
                                {Object.keys(STATUS_CFG).map((s) => <option key={s} value={s}>{s}</option>)}
                              </select>
                            ) : (
                              <input
                                type={col.key.includes("Date") ? "date" : "text"}
                                value={editForm[col.key] || ""}
                                onChange={(e) => setEditForm((f) => ({ ...f, [col.key]: e.target.value }))}
                                className="border border-slate-300 rounded px-2 py-1 text-xs w-full min-w-24"
                              />
                            )}
                          </td>
                        ))}
                        <td className="px-2 py-1.5 whitespace-nowrap">
                          <button onClick={saveEdit} className="text-xs bg-blue-600 hover:bg-blue-700 text-white px-2.5 py-1 rounded mr-1 transition-colors">保存</button>
                          <button onClick={() => setEditingId(null)} className="text-xs text-slate-400 hover:text-slate-700">キャンセル</button>
                        </td>
                      </>
                    ) : (
                      <>
                        {visibleDefs.map((col) => (
                          <td key={col.key} className="px-3 py-2.5 whitespace-nowrap max-w-xs">
                            {col.key === "status" ? (
                              <StatusBadge status={rec.status} />
                            ) : col.key === "url" && rec.url ? (
                              <a href={rec.url} target="_blank" rel="noreferrer"
                                className="text-blue-600 hover:underline text-xs block max-w-36 truncate">
                                {rec.url}
                              </a>
                            ) : col.key === "nextCallDate" && rec.nextCallDate && rec.nextCallDate <= todayStr ? (
                              <span className="text-amber-600 font-semibold text-xs">{rec.nextCallDate} ⚠</span>
                            ) : col.key === "memo" ? (
                              <span className="text-slate-600 text-xs block max-w-48 truncate" title={rec[col.key]}>
                                {rec[col.key] || "—"}
                              </span>
                            ) : (
                              <span className="text-slate-700 text-xs">{rec[col.key] || "—"}</span>
                            )}
                          </td>
                        ))}
                        <td className="px-3 py-2.5 whitespace-nowrap">
                          <button onClick={() => startEdit(rec)} className="text-xs text-blue-600 hover:text-blue-800 mr-2 font-medium">編集</button>
                          <button onClick={() => deleteOne(rec.id)} className="text-xs text-rose-500 hover:text-rose-700 font-medium">削除</button>
                        </td>
                      </>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="px-4 py-2 border-t border-slate-100 text-xs text-slate-400 flex justify-between">
            <span>{filtered.length}件表示 ／ 全{records.length}件</span>
            <span>更新: {new Date().toLocaleDateString("ja-JP")}</span>
          </div>
        </div>
      </div>

      {/* ── 新規追加モーダル ── */}
      {addingNew && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg p-6">
            <h2 className="text-base font-bold text-slate-800 mb-4">新規レコード追加</h2>
            <div className="grid grid-cols-2 gap-3">
              {[
                { key: "companyName",  label: "企業名 *",      span: 2 },
                { key: "phone",        label: "電話番号" },
                { key: "email",        label: "メールアドレス" },
                { key: "url",          label: "GBP/URL",        span: 2 },
                { key: "assignee",     label: "担当者" },
                { key: "nextCallDate", label: "次回架電日",     type: "date" },
              ].map((f) => (
                <div key={f.key} className={f.span === 2 ? "col-span-2" : ""}>
                  <label className="block text-xs text-slate-500 mb-1">{f.label}</label>
                  <input type={f.type || "text"} value={newForm[f.key] || ""}
                    onChange={(e) => setNewForm((p) => ({ ...p, [f.key]: e.target.value }))}
                    className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                  />
                </div>
              ))}
              <div>
                <label className="block text-xs text-slate-500 mb-1">ステータス</label>
                <select value={newForm.status} onChange={(e) => setNewForm((p) => ({ ...p, status: e.target.value }))}
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">
                  {Object.keys(STATUS_CFG).map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
              <div className="col-span-2">
                <label className="block text-xs text-slate-500 mb-1">メモ</label>
                <textarea value={newForm.memo || ""} onChange={(e) => setNewForm((p) => ({ ...p, memo: e.target.value }))}
                  rows={3}
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none" />
              </div>
            </div>
            <div className="flex gap-2 mt-4 justify-end">
              <button onClick={() => { setAddingNew(false); setNewForm(emptyForm()); }}
                className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">キャンセル</button>
              <button onClick={saveNew} disabled={!newForm.companyName.trim()}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg disabled:opacity-40 transition-colors">
                追加する
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── 列設定モーダル ── */}
      {showColModal && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-xs p-6">
            <h2 className="text-base font-bold text-slate-800 mb-4">表示列の設定</h2>
            <div className="space-y-1">
              {ALL_COLUMNS.map((col) => (
                <label key={col.key} className="flex items-center gap-3 px-2 py-2 hover:bg-slate-50 rounded-lg cursor-pointer">
                  <input type="checkbox" checked={visibleCols.includes(col.key)} disabled={col.required}
                    onChange={(e) => setVisibleCols((p) => e.target.checked ? [...p, col.key] : p.filter((k) => k !== col.key))}
                    className="rounded border-slate-300 text-blue-600"
                  />
                  <span className="text-sm text-slate-700">{col.label}</span>
                  {col.required && <span className="ml-auto text-xs text-slate-400">必須</span>}
                </label>
              ))}
            </div>
            <button onClick={() => setShowColModal(false)}
              className="mt-4 w-full bg-blue-600 hover:bg-blue-700 text-white rounded-lg py-2 text-sm font-medium transition-colors">
              閉じる
            </button>
          </div>
        </div>
      )}

      {/* ── 重複クレンジングモーダル ── */}
      {showDupeModal && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl p-6 flex flex-col max-h-[80vh]">
            <h2 className="text-base font-bold text-slate-800 mb-1">重複クレンジング</h2>
            <p className="text-sm text-slate-500 mb-4">
              {dupeGroups.length === 0
                ? "重複データは見つかりませんでした。"
                : `${dupeGroups.length}グループ / ${dupeGroups.reduce((s, g) => s + g.records.length - 1, 0)}件の重複を検出。最新を残して削除します。`}
            </p>
            {dupeGroups.length > 0 && (
              <div className="overflow-y-auto flex-1 space-y-2 mb-4 pr-1">
                {dupeGroups.map((g, gi) => (
                  <div key={gi} className="border border-slate-200 rounded-xl p-3">
                    <div className="font-semibold text-sm text-slate-700 mb-2">{g.records[0].companyName}</div>
                    <div className="space-y-1">
                      {g.records.map((r, i) => (
                        <div key={r.id} className={`flex flex-wrap items-center gap-2 text-xs ${i === 0 ? "text-green-700" : "text-slate-400 line-through"}`}>
                          <span className="shrink-0 font-medium">{i === 0 ? "✓ 残す" : "× 削除"}</span>
                          <span>{r.importedAt?.slice(0, 10)}</span>
                          <span>{r.phone || "—"}</span>
                          <StatusBadge status={r.status} />
                          {r.assignee && <span>{r.assignee}</span>}
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}
            <div className="flex gap-2 justify-end">
              <button onClick={() => { setShowDupeModal(false); setDupeGroups([]); }}
                className="px-4 py-2 text-sm text-slate-500 border border-slate-200 rounded-lg hover:bg-slate-50">キャンセル</button>
              {dupeGroups.length > 0 && (
                <button onClick={cleanDupes}
                  className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white text-sm font-semibold rounded-lg transition-colors">
                  重複を削除する
                </button>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
