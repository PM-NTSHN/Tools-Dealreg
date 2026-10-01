/* ==========================================================================
   Tool tổng hợp Dealreg — NTS Hanoi Corp.
   Toàn bộ xử lý diễn ra trong trình duyệt; không có request mạng nào.
   ========================================================================== */
(function () {
'use strict';

const VERSION = '__VERSION__';
const DAY = 864e5;

/* ------------------------------------------------------------------ utils */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const nfc = s => (s == null ? '' : String(s)).normalize('NFC');
// Bỏ dấu từng ký tự, giữ nguyên độ dài chuỗi để highlight chính xác vị trí
const foldMap = new Map();
function foldChar(c) {
  let f = foldMap.get(c);
  if (f === undefined) {
    if (c === 'đ' || c === 'Đ') f = 'd';
    else { f = c.normalize('NFD')[0].toLowerCase(); if (f.length !== 1) f = c.toLowerCase(); if (f.length !== 1) f = c; }
    foldMap.set(c, f);
  }
  return f;
}
function fold(s) { s = nfc(s); let o = ''; for (let i = 0; i < s.length; i++) o += foldChar(s[i]); return o; }
const slug = s => fold(s).replace(/[^a-z0-9]+/g, '');
const pct = (a, b, d = 1) => (b ? (a / b * 100).toFixed(d).replace('.', ',') + '%' : '—');
const fmtN = n => (n == null ? '—' : Number(n).toLocaleString('vi-VN'));
const pad = n => String(n).padStart(2, '0');
const fmtD = d => (d ? `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}` : '');
const fmtDT = d => (d ? `${pad(d.getHours())}:${pad(d.getMinutes())} ${fmtD(d)}` : '');
const isoD = d => (d ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` : '');
const sod = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const addMonths = (d, m) => { const x = new Date(d.getFullYear(), d.getMonth() + m, 1); const last = new Date(x.getFullYear(), x.getMonth() + 1, 0).getDate(); x.setDate(Math.min(d.getDate(), last)); return x; };
const quarterKey = d => `${d.getFullYear()}-Q${Math.floor(d.getMonth() / 3) + 1}`;
const quarterLabel = k => { const [y, q] = k.split('-'); return `${q}/${y}`; };
const monthKey = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
const monthLabel = k => { const [y, m] = k.split('-'); return `T${+m}/${y}`; };
function weekKey(d) { const x = sod(d); const wd = (x.getDay() + 6) % 7; x.setDate(x.getDate() - wd); return isoD(x); }
const median = a => { if (!a.length) return null; const s = [...a].sort((x, y) => x - y); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
function countBy(rows, fn) { const m = new Map(); for (const r of rows) for (const k of [].concat(fn(r))) m.set(k, (m.get(k) || 0) + 1); return m; }
const sortedEntries = m => [...m.entries()].sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0]), 'vi'));
function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }
function toast(msg) { const t = document.createElement('div'); t.className = 'toast'; t.textContent = msg; document.body.appendChild(t); setTimeout(() => t.remove(), 2600); }

function parseDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === 'number') { // Excel serial
    if (v < 1 || v > 2958465) return null;
    const ms = Math.round((v - 25569) * DAY); const u = new Date(ms);
    return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate(), u.getUTCHours(), u.getUTCMinutes(), u.getUTCSeconds());
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s+(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (m) return new Date(+m[6], m[5] - 1, +m[4], +m[1], +m[2], +(m[3] || 0));
  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) return new Date(+m[3], m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return new Date(+m[1], m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  m = s.match(/^(\d{1,2})[\/.-](\d{4})$/); // mm/yyyy
  if (m) return new Date(+m[2], +m[1], 0);
  m = s.match(/^Q([1-4])[\/ -]?(\d{4})$/i);
  if (m) return new Date(+m[2], m[1] * 3, 0);
  return null;
}

/* --------------------------------------------------------------- columns */
// Map tiêu đề cột (không dấu, chữ thường) → trường nội bộ. Nhận diện theo tên nên không phụ thuộc thứ tự cột.
const COLMAP = {
  no: ['no.', 'no', 'stt'],
  id: ['id', 'ma phieu'],
  creator: ['nguoi tao'],
  service: ['dich vu'],
  name: ['ten', 'ten phieu'],
  category: ['danh muc phieu'],
  desc: ['mo ta phieu', 'mo ta'],
  attach: ['tep dinh kem phieu'],
  pm: ['nguoi thuc hien'],
  followers: ['nguoi theo doi'],
  created: ['tao luc'],
  updated: ['cap nhat lan cuoi luc'],
  sale: ['sale'],
  eu: ['ten end user (bang tieng anh)', 'ten end user', 'end user'],
  euEmail: ['email cua end user'],
  custInfo: ['thong tin khach hang'],
  project: ['ten du an'],
  reseller: ['reseller'],
  vendor: ['ten hang'],
  bom: ['thong tin bom (model, so luong, nam support...)', 'thong tin bom'],
  timeline: ['timeline du an'],
  website: ['website khach hang'],
  note: ['note', 'ghi chu'],
  crm: ['link deal crm'],
  origin: ['link to origin object'],
  blockId: ['id khoi cua phieu'],
  block: ['ten khoi'],
  blockType: ['loai khoi'],
  status: ['trang thai'],
  blockCreated: ['tao khoi luc'],
  blockUpdated: ['cap nhat khoi lan cuoi luc'],
  sla: ['sla/thoi han'],
  overdue: ['qua han'],
};
const REQUIRED = ['id', 'name', 'status', 'block'];
const LABEL_OF = { id: 'ID', name: 'Tên', status: 'Trạng thái', block: 'Tên khối' };

/* ------------------------------------------------------------ dictionaries */
// Chuẩn hoá tên hãng: viết hoa/thường và lỗi chính tả thường gặp
const VENDOR_ALIAS = { optswat: 'opswat', opswatt: 'opswat', kaspesky: 'kaspersky', karpersky: 'kaspersky', infoexpres: 'infoexpress', proofpont: 'proofpoint' };
const VENDOR_CANON = { opswat: 'OPSWAT', gtb: 'GTB', hcl: 'HCL', arcserve: 'Arcserve', kaspersky: 'Kaspersky', barracuda: 'Barracuda', safetica: 'Safetica', sophos: 'Sophos', delinea: 'Delinea', progress: 'Progress', infoexpress: 'InfoExpress', netgear: 'Netgear', netscout: 'NETSCOUT', acronis: 'Acronis', opentext: 'OpenText', proofpoint: 'Proofpoint', zecurion: 'Zecurion', stellar: 'Stellar', radware: 'Radware', trellix: 'Trellix', forcepoint: 'Forcepoint', secpod: 'SecPod', thales: 'Thales', cloudflare: 'Cloudflare', qualys: 'Qualys', fortinet: 'Fortinet', paloalto: 'Palo Alto', checkpoint: 'Check Point', veeam: 'Veeam', manageengine: 'ManageEngine', solarwinds: 'SolarWinds', gigamon: 'Gigamon', imperva: 'Imperva', tenable: 'Tenable', rapid7: 'Rapid7', trendmicro: 'Trend Micro', eset: 'ESET', bitdefender: 'Bitdefender' };

const FLOW = ['Tạo yêu cầu đăng ký dự án', 'Tự động tìm PM hãng', 'PM tiếp nhận', 'PM gửi DR đến hãng', 'Hãng check', 'Approved', 'Rejected'];
const BLOCK_COLORS = ['#94a3b8', '#7c8fb5', '#2b7fd8', '#6b5bd6', '#f59e0b', '#1fa55b', '#e5484d', '#10b7de', '#c2410c', '#64748b'];
const RESULT = {
  'Approved': { c: '#1fa55b', cls: 'b-ok' },
  'Đang xử lý': { c: '#f59e0b', cls: 'b-warn' },
  'Rejected': { c: '#e5484d', cls: 'b-bad' },
  'Hoàn thành khác': { c: '#94a3b8', cls: 'b-mute' },
};
const FLAGS = {
  '0': { t: 'Bình thường (≤ 1 tháng)', s: 'Bình thường', c: '#1fa55b' },
  '1': { t: 'Mức 1 · > 1 tháng', s: 'Mức 1', c: '#f2b705' },
  '2': { t: 'Mức 2 · > 3 tháng', s: 'Mức 2', c: '#f07c1b' },
  '3': { t: 'Mức 3 · > 6 tháng', s: 'Mức 3', c: '#e03a3e' },
  '4': { t: 'Mức 4 · > 1 năm', s: 'Mức 4', c: '#8b1538' },
  'na': { t: 'Không áp dụng', s: '—', c: '#9aa6b2' },
};
const TLS = {
  past: { t: 'Đã qua timeline', c: '#e5484d' },
  d30: { t: 'Trong 30 ngày tới', c: '#f59e0b' },
  d90: { t: '31 – 90 ngày tới', c: '#2b7fd8' },
  later: { t: 'Sau 90 ngày', c: '#1fa55b' },
  none: { t: 'Chưa có timeline', c: '#9aa6b2' },
};
const ISSUES = {
  conflict: 'Trùng EU + Hãng, khác Sale',
  dup: 'Trùng EU + Hãng',
  vendorMismatch: 'Hãng ở tên phiếu ≠ cột Tên hãng',
  tlPastOpen: 'Đã qua timeline, chưa xong',
  slaOver: 'Quá hạn SLA',
  noSale: 'Thiếu Sale',
  noVendor: 'Thiếu Tên hãng',
  noTimeline: 'Thiếu Timeline dự án',
  noEU: 'Thiếu End user',
  __ok: 'Không có cảnh báo',
};

/* ------------------------------------------------------------------ state */
const S = {
  files: [],          // {name, rows, added, dup, sheet}
  rows: [],           // phiếu đã chuẩn hoá
  byId: new Map(),
  ref: sod(new Date()),
  scope: 'all',
  view: [],
  sort: { k: 'lastUpdate', d: -1 },
  page: 1, pageSize: 50,
  f: {},              // facet → Set
  q: '', dateField: 'created', from: null, to: null,
  trendGran: 'month',
  charts: {},
  people: new Map(),  // key → {full, user}
  vendorLabel: new Map(),
  resellerLabel: new Map(),
  hiddenCols: new Set(['creator', 'reseller', 'project', 'sla', 'note', 'created', 'followers']),
  failing: null,
};

const personLabel = k => { if (!k || k === '__none') return '(Trống)'; const p = S.people.get(k); return p ? (p.full || '@' + p.user) : k; };
const personSub = k => { const p = S.people.get(k); return p && p.full && p.user ? '@' + p.user : ''; };

/* ----------------------------------------------------------------- facets */
const FACETS = {
  sale: { label: 'Sale', ph: 'Tất cả — gõ tên sale…', vals: r => [r.sale || '__none'], lab: personLabel, sub: personSub },
  creator: { label: 'Người tạo', ph: 'Tất cả — gõ tên…', vals: r => [r.creator || '__none'], lab: personLabel, sub: personSub },
  pm: { label: 'PM phụ trách', ph: 'Tất cả — gõ tên PM…', vals: r => (r.pms.length ? r.pms : ['__none']), lab: personLabel, sub: personSub },
  vendor: { label: 'Hãng', ph: 'Tất cả — gõ tên hãng…', vals: r => (r.vendors.length ? r.vendors : ['__none']), lab: k => (k === '__none' ? '(Trống)' : S.vendorLabel.get(k) || k) },
  block: { label: 'Tiến độ xử lý', ph: 'Tất cả', vals: r => [r.block || '__none'], order: () => blockOrder() },
  status: { label: 'Trạng thái', ph: 'Tất cả', vals: r => [r.status || '__none'] },
  flag: { label: 'Cờ nhắc update', ph: 'Tất cả', vals: r => [r.flag], lab: k => FLAGS[k].t, order: () => ['4', '3', '2', '1', '0', 'na'], dot: k => FLAGS[k].c },
  result: { label: 'Kết quả DR', ph: 'Tất cả', vals: r => [r.result], order: () => Object.keys(RESULT), dot: k => RESULT[k].c },
  tl: { label: 'Tình trạng Timeline DA', ph: 'Tất cả', vals: r => [r.tl], lab: k => TLS[k].t, order: () => Object.keys(TLS), dot: k => TLS[k].c },
  reseller: { label: 'Reseller / Partner', ph: 'Tất cả — gõ tên…', vals: r => (r.resellers.length ? r.resellers : ['__none']), lab: k => (k === '__none' ? '(Trống)' : S.resellerLabel.get(k) || k) },
  overdue: { label: 'Quá hạn SLA', ph: 'Tất cả', vals: r => [r.overdue || '__none'] },
  issue: { label: 'Cảnh báo dữ liệu', ph: 'Tất cả', vals: r => (r.issues.length ? r.issues : ['__ok']), lab: k => ISSUES[k] || k, order: () => Object.keys(ISSUES) },
  follower: { label: 'Người theo dõi', ph: 'Tất cả — gõ tên…', vals: r => (r.followers.length ? r.followers : ['__none']), lab: personLabel, sub: personSub },
  month: { label: 'Tháng tạo phiếu', ph: 'Tất cả', vals: r => [r.month || '__none'], lab: k => (k === '__none' ? '(Trống)' : 'Tháng ' + monthLabel(k).slice(1)), order: all => [...all].sort().reverse() },
};
Object.keys(FACETS).forEach(k => { S.f[k] = new Set(); });
const facetLabel = (fk, v) => { const f = FACETS[fk]; if (v === '__none') return '(Trống)'; return f.lab ? f.lab(v) : v; };

function blockOrder() {
  const seen = new Set(S.rows.map(r => r.block).filter(Boolean));
  const out = FLOW.filter(b => seen.has(b));
  const extra = [...seen].filter(b => !FLOW.includes(b));
  // khối lạ chèn trước Approved/Rejected
  const endIdx = out.findIndex(b => b === 'Approved' || b === 'Rejected');
  if (endIdx < 0) return out.concat(extra);
  return out.slice(0, endIdx).concat(extra, out.slice(endIdx));
}
const blockColor = b => { const i = blockOrder().indexOf(b); if (b === 'Approved') return '#1fa55b'; if (b === 'Rejected') return '#e5484d'; return BLOCK_COLORS[i >= 0 ? i % BLOCK_COLORS.length : 9]; };

/* ------------------------------------------------------------- file load */
async function loadFiles(fileList) {
  const files = [...fileList].filter(f => /\.(xlsx|xlsm|xls|csv)$/i.test(f.name));
  if (!files.length) { toast('Vui lòng chọn file Excel (.xlsx / .xls)'); return; }
  const warns = [];
  for (const file of files) {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: true, dense: false });
      const res = pickSheet(wb);
      if (!res) { warns.push(`<b>${esc(file.name)}</b>: không tìm thấy bảng dữ liệu có cột ID / Tên / Trạng thái / Tên khối.`); continue; }
      const missing = REQUIRED.filter(k => res.map[k] == null);
      if (missing.length) warns.push(`<b>${esc(file.name)}</b>: thiếu cột ${missing.map(k => LABEL_OF[k]).join(', ')}.`);
      const optMissing = ['sale', 'vendor', 'pm', 'creator', 'updated', 'timeline'].filter(k => res.map[k] == null);
      if (optMissing.length) warns.push(`<b>${esc(file.name)}</b>: không thấy cột ${optMissing.join(', ')} — các bộ lọc liên quan sẽ trống.`);
      const raw = res.data.map(row => rawRecord(row, res.map)).filter(x => x.id);
      let added = 0, dup = 0;
      for (const rr of raw) {
        rr._file = file.name;
        const ex = S.byId.get(rr.id);
        if (ex) {
          dup++;
          const a = lastOf(ex), b = lastOf(rr);
          if ((b || 0) >= (a || 0)) { S.byId.set(rr.id, rr); }
        } else { S.byId.set(rr.id, rr); added++; }
      }
      S.files.push({ name: file.name, rows: raw.length, added, dup, sheet: res.sheet });
    } catch (e) {
      console.error(e);
      warns.push(`<b>${esc(file.name)}</b>: lỗi đọc file (${esc(e.message)}).`);
    }
  }
  rebuild();
  $('#load-warn').innerHTML = warns.length ? `<div class="warnbox">⚠ ${warns.join('<br>⚠ ')}</div>` : '';
  if (S.rows.length) toast(`Đã nạp ${fmtN(S.rows.length)} phiếu từ ${S.files.length} file`);
}
const lastOf = r => Math.max(+(parseDate(r.updated) || 0), +(parseDate(r.blockUpdated) || 0));

function pickSheet(wb) {
  let best = null;
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) continue;
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    for (let i = 0; i < Math.min(25, data.length); i++) {
      const map = headerMap(data[i] || []);
      const score = Object.keys(map).length;
      if (map.id != null && score >= 4 && (!best || score > best.score)) best = { sheet: name, score, map, data: data.slice(i + 1) };
    }
  }
  return best;
}
function headerMap(row) {
  const map = {};
  const heads = row.map(h => fold(h).trim().replace(/\s+/g, ' '));
  for (const [k, alts] of Object.entries(COLMAP)) {
    let idx = heads.findIndex(h => alts.includes(h));
    if (idx < 0 && k !== 'name' && k !== 'id' && k !== 'no') idx = heads.findIndex(h => h && alts.some(a => a.length > 5 && h.startsWith(a)));
    if (idx >= 0) map[k] = idx;
  }
  return map;
}
function rawRecord(row, map) {
  const o = {};
  for (const [k, i] of Object.entries(map)) {
    let v = row[i];
    if (v instanceof Date) { o[k] = v; continue; }
    if (v != null && typeof v !== 'number') v = nfc(String(v)).replace(/ /g, ' ').trim();
    o[k] = v === '' ? null : v;
  }
  if (o.id != null) o.id = String(o.id).trim();
  return o;
}

/* ------------------------------------------------------------- normalise */
function nameKey(full) {
  const p = fold(full).replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter(Boolean);
  if (!p.length) return '';
  if (p.length === 1) return p[0];
  return p[p.length - 1] + p[0]; // "Vy Công Quý" → "quyvy" (quy ước username Base)
}
const userTokens = s => (s ? String(s).match(/@[\w.\-]+/g) || [] : []).map(t => t.slice(1));

function rebuild() {
  const raws = [...S.byId.values()];
  // 1) Danh bạ người: username ↔ họ tên
  const users = new Map();   // lower → original username
  const fulls = new Map();   // nameKey → Map(fold(full) → full)
  const noteUser = u => { const l = u.toLowerCase(); if (!users.has(l)) users.set(l, u); };
  for (const r of raws) {
    if (r.creator) String(r.creator).split(/[\s,;]+/).filter(Boolean).forEach(u => noteUser(u.replace(/^@/, '')));
    userTokens(r.pm).forEach(noteUser); userTokens(r.followers).forEach(noteUser);
    if (r.sale) {
      const toks = userTokens(r.sale);
      if (toks.length) toks.forEach(noteUser);
      else { const k = nameKey(r.sale); if (k) { if (!fulls.has(k)) fulls.set(k, new Map()); const m = fulls.get(k); const fk = fold(r.sale); if (!m.has(fk)) m.set(fk, titleName(r.sale)); } }
    }
  }
  S.people = new Map();
  for (const [l, u] of users) S.people.set(l, { user: u, full: null });
  const saleKeyOfFull = new Map();
  for (const [k, m] of fulls) {
    if (m.size === 1) {
      const full = [...m.values()][0];
      const p = S.people.get(k);
      if (p) p.full = full; else S.people.set(k, { user: null, full });
      saleKeyOfFull.set([...m.keys()][0], k);
    } else {
      // Trùng khoá (2 người khác tên cùng ra 1 username) → giữ riêng theo họ tên, không gộp
      for (const [fk, full] of m) { const kk = 'n:' + fk; S.people.set(kk, { user: null, full }); saleKeyOfFull.set(fk, kk); }
    }
  }
  // 2) Hãng & reseller
  const vCount = new Map(), rCount = new Map();
  const vKey = s => { let k = slug(s); return VENDOR_ALIAS[k] || k; };
  const splitList = s => (s ? String(s).split(/\s*[,;\/\n]\s*|\s+&\s+/).map(x => x.trim()).filter(Boolean) : []);
  const tally = (m, k, lab) => { if (!m.has(k)) m.set(k, new Map()); const x = m.get(k); x.set(lab, (x.get(lab) || 0) + 1); };
  // 3) Dựng bản ghi
  S.rows = raws.map(o => {
    const r = {
      id: o.id, no: o.no, name: o.name || '', service: o.service || '',
      desc: o.desc || '', note: o.note || '', bom: o.bom || '',
      eu: o.eu || '', euEmail: o.euEmail || '', custInfo: o.custInfo || '', project: o.project || '',
      website: o.website || '', crm: o.crm || '', origin: o.origin || '',
      block: o.block || '', blockType: o.blockType || '', blockId: o.blockId ? String(o.blockId) : '',
      status: o.status || '', overdue: o.overdue || '',
      created: parseDate(o.created), updated: parseDate(o.updated),
      blockCreated: parseDate(o.blockCreated), blockUpdated: parseDate(o.blockUpdated),
      sla: parseDate(o.sla), timeline: parseDate(o.timeline), timelineRaw: o.timeline instanceof Date ? fmtD(o.timeline) : (o.timeline || ''),
      file: o._file,
      saleRaw: o.sale || '', vendorRaw: o.vendor || '', resellerRaw: o.reseller || '',
      creatorRaw: o.creator || '', pmRaw: o.pm || '', followersRaw: o.followers || '',
    };
    r.creator = o.creator ? String(o.creator).replace(/^@/, '').split(/[\s,;]+/)[0].toLowerCase() : '';
    r.pms = [...new Set(userTokens(o.pm).map(u => u.toLowerCase()))];
    if (!r.pms.length && o.pm) r.pms = [String(o.pm).replace(/^@/, '').trim().toLowerCase()];
    r.followers = [...new Set(userTokens(o.followers).map(u => u.toLowerCase()))];
    if (o.sale) {
      const toks = userTokens(o.sale);
      r.sale = toks.length ? toks[0].toLowerCase() : (saleKeyOfFull.get(fold(o.sale)) || nameKey(o.sale));
    } else r.sale = '';
    r.vendors = [...new Set(splitList(o.vendor).map(v => { const k = vKey(v); tally(vCount, k, v); return k; }).filter(Boolean))];
    r.resellers = [...new Set(splitList(o.reseller).map(v => { const k = slug(v); tally(rCount, k, v); return k; }).filter(Boolean))];
    r.attachments = parseAttach(o.attach);
    r.lastUpdate = [r.updated, r.blockUpdated].filter(Boolean).sort((a, b) => b - a)[0] || r.created || null;
    const fb = fold(r.block);
    r.result = fb.includes('approv') || fb.includes('duyet') ? 'Approved' : fb.includes('reject') || fb.includes('tu choi') ? 'Rejected'
      : fold(r.status).includes('hoan thanh') ? 'Hoàn thành khác' : 'Đang xử lý';
    r.open = r.result === 'Đang xử lý';
    r.month = r.created ? monthKey(r.created) : '';
    r.euKey = slug(r.eu);
    return r;
  });
  const pickLabel = (m, canon) => { const out = new Map(); for (const [k, labs] of m) { let lab = canon && canon[k]; if (!lab) { lab = sortedEntries(labs)[0][0]; if (lab === lab.toLowerCase()) lab = lab.charAt(0).toUpperCase() + lab.slice(1); } out.set(k, lab); } return out; };
  S.vendorLabel = pickLabel(vCount, VENDOR_CANON);
  S.resellerLabel = pickLabel(rCount);
  // 4) Quan hệ: trùng EU + hãng, phiếu cùng EU
  const byEU = new Map();
  for (const r of S.rows) if (r.euKey) { if (!byEU.has(r.euKey)) byEU.set(r.euKey, []); byEU.get(r.euKey).push(r); }
  const knownVendors = new Set([...S.vendorLabel.keys(), ...Object.keys(VENDOR_CANON)]);
  for (const r of S.rows) {
    r.related = (byEU.get(r.euKey) || []).filter(x => x !== r);
    r.staticIssues = [];
    const dupPeers = r.result === 'Rejected' ? [] : r.related.filter(x => x.result !== 'Rejected' && x.vendors.some(v => r.vendors.includes(v)));
    r.dupPeers = dupPeers;
    if (dupPeers.length) { r.staticIssues.push('dup'); if (dupPeers.some(x => x.sale && r.sale && x.sale !== r.sale)) r.staticIssues.unshift('conflict'); }
    const segs = r.name.split(/[_|–]+/).map(s => s.trim()).filter(Boolean);
    const last = segs.length > 1 ? segs[segs.length - 1] : '';
    if (last && r.vendors.length) {
      const words = [slug(last), ...fold(last).split(/[^a-z0-9]+/)].map(w => VENDOR_ALIAS[w] || w).filter(w => w.length >= 3);
      const hit = words.find(w => knownVendors.has(w));
      if (hit && !r.vendors.includes(hit)) { r.staticIssues.push('vendorMismatch'); r.nameVendor = VENDOR_CANON[hit] || S.vendorLabel.get(hit) || last; }
    }
    if (!r.sale) r.staticIssues.push('noSale');
    if (!r.vendors.length) r.staticIssues.push('noVendor');
    if (!r.timeline) r.staticIssues.push('noTimeline');
    if (!r.eu) r.staticIssues.push('noEU');
    r.hay = fold([r.id, r.blockId, r.name, r.eu, r.euEmail, r.project, r.custInfo, r.note, r.desc, r.bom, r.website, r.resellerRaw, r.vendorRaw,
      r.vendors.map(v => S.vendorLabel.get(v)).join(' '), r.saleRaw, personLabel(r.sale), r.creatorRaw, personLabel(r.creator),
      r.pmRaw, r.pms.map(personLabel).join(' '), r.block, r.status, r.crm].join(' \u0001 '));
  }
  // Mặc định ngày mốc = hôm nay
  derive();
  // Xoá lựa chọn không còn tồn tại
  S.page = 1;
  renderFileBar();
  const has = S.rows.length > 0;
  $('#app').classList.toggle('hidden', !has);
  $('#upload-section').classList.toggle('hidden', has);
  $('#filebar-section').classList.toggle('hidden', !S.files.length);
  $('#btn-export').disabled = !has; $('#btn-export-more').disabled = !has;
  if (has) update();
  renderHeroSub();
}
function titleName(s) { return nfc(s).trim().replace(/\s+/g, ' ').split(' ').map(w => w.charAt(0).toLocaleUpperCase('vi') + w.slice(1)).join(' '); }
function parseAttach(s) {
  if (!s) return [];
  const out = []; const re = /([^\n]*?)\s*-\s*\((https?:\/\/[^\s)]+)\)/g; let m;
  while ((m = re.exec(s))) out.push({ name: m[1].replace(/^[,;\s]*\d+\.\s*/, '').trim() || 'Tệp đính kèm', url: m[2] });
  if (!out.length) out.push({ name: String(s).slice(0, 120), url: '' });
  return out;
}

// Các trường phụ thuộc ngày mốc / phạm vi cờ
function derive() {
  const ref = S.ref, refEnd = new Date(+ref + DAY - 1);
  const b1 = addMonths(ref, -1), b3 = addMonths(ref, -3), b6 = addMonths(ref, -6), b12 = addMonths(ref, -12);
  for (const r of S.rows) {
    const lu = r.lastUpdate;
    r.days = lu ? Math.max(0, Math.floor((refEnd - lu) / DAY)) : null;
    r.age = r.created ? Math.max(0, Math.floor((refEnd - r.created) / DAY)) : null;
    const inScope = S.scope === 'all' || (S.scope === 'open' ? r.open : r.result !== 'Rejected');
    if (!inScope || !lu) r.flag = 'na';
    else r.flag = lu < b12 ? '4' : lu < b6 ? '3' : lu < b3 ? '2' : lu < b1 ? '1' : '0';
    r.flagN = r.flag === 'na' ? -1 : +r.flag;
    if (!r.timeline) r.tl = 'none';
    else { const d = Math.floor((sod(r.timeline) - ref) / DAY); r.tl = d < 0 ? 'past' : d <= 30 ? 'd30' : d <= 90 ? 'd90' : 'later'; }
    r.issues = [...r.staticIssues];
    if (r.open && r.tl === 'past') r.issues.push('tlPastOpen');
    if (r.open && fold(r.overdue) === 'co') r.issues.push('slaOver');
    r.procDays = !r.open && r.created && r.lastUpdate ? Math.max(0, (r.lastUpdate - r.created) / DAY) : null;
  }
}

/* --------------------------------------------------------------- filter */
function qTokens() { return fold(S.q).split(/\s+/).filter(Boolean); }
function computeView() {
  const facetKeys = Object.keys(FACETS).filter(k => S.f[k].size);
  const toks = qTokens();
  const df = S.dateField, from = S.from, to = S.to ? new Date(+S.to + DAY - 1) : null;
  const view = []; const failing = new Map(); // row → single failing key
  for (const r of S.rows) {
    let fails = 0, which = null;
    for (const k of facetKeys) {
      const set = S.f[k];
      if (!FACETS[k].vals(r).some(v => set.has(v))) { fails++; which = k; if (fails > 1) break; }
    }
    if (fails <= 1 && toks.length && !toks.every(t => r.hay.includes(t))) { fails++; which = 'q'; }
    if (fails <= 1 && (from || to)) { const d = r[df]; if (!d || (from && d < from) || (to && d > to)) { fails++; which = 'date'; } }
    if (!fails) view.push(r); else if (fails === 1) failing.set(r, which);
  }
  S.failing = failing;
  const col = COLS.find(c => c.k === S.sort.k) || COLS[0];
  const dir = S.sort.d;
  view.sort((a, b) => { const x = col.sv(a), y = col.sv(b); if (x == null && y == null) return 0; if (x == null) return 1; if (y == null) return -1; return (x < y ? -1 : x > y ? 1 : 0) * dir || (+b.id - +a.id); });
  S.view = view;
}
function facetCounts(fk) {
  const m = new Map();
  const add = r => { for (const v of FACETS[fk].vals(r)) m.set(v, (m.get(v) || 0) + 1); };
  S.view.forEach(add);
  for (const [r, k] of S.failing) if (k === fk) add(r);
  return m;
}
function facetOptions(fk) {
  const f = FACETS[fk];
  const all = new Map(); for (const r of S.rows) for (const v of f.vals(r)) all.set(v, (all.get(v) || 0) + 1);
  const cnt = facetCounts(fk);
  let keys = f.order ? f.order([...all.keys()]).filter(k => all.has(k)) : [...all.keys()].sort((a, b) => (cnt.get(b) || 0) - (cnt.get(a) || 0) || all.get(b) - all.get(a) || facetLabel(fk, a).localeCompare(facetLabel(fk, b), 'vi'));
  if (f.order) keys = keys.concat([...all.keys()].filter(k => !keys.includes(k)));
  if (keys.includes('__none')) keys = keys.filter(k => k !== '__none').concat('__none');
  return keys.map(v => ({ v, label: facetLabel(fk, v), sub: f.sub ? f.sub(v) : '', count: cnt.get(v) || 0, dot: f.dot ? f.dot(v) : null }));
}
function toggleFacet(fk, v, only) {
  const s = S.f[fk];
  if (only) { const was = s.size === 1 && s.has(v); s.clear(); if (!was) s.add(v); }
  else if (s.has(v)) s.delete(v); else s.add(v);
  S.page = 1; update();
}
function setFacet(fk, vals) { S.f[fk] = new Set(vals); }
function resetFilters() {
  Object.keys(FACETS).forEach(k => S.f[k].clear());
  S.q = ''; $('#q').value = ''; $('#searchbox').classList.remove('has-val');
  S.from = S.to = null; $('#date-from').value = ''; $('#date-to').value = ''; $('#date-preset').value = '';
  S.page = 1; update();
}

/* ---------------------------------------------------------- multiselect */
const MS = {};
class MultiSelect {
  constructor(host, fk) {
    this.fk = fk; this.f = FACETS[fk]; this.host = host; this.act = 0; this.opts = [];
    host.innerHTML = `<div class="lbl">${esc(this.f.label)}<span class="cnt"></span></div>
      <div class="ms"><div class="ms-box"><span class="tags" style="display:contents"></span><input class="ms-input" placeholder="${esc(this.f.ph)}" autocomplete="off"></div>
      <div class="ms-pop hidden"><div class="ms-tools"><button data-a="all">✓ Chọn các mục đang hiện</button><button data-a="none">✕ Bỏ chọn</button></div><div class="ms-list"></div></div></div>`;
    this.el = $('.ms', host); this.box = $('.ms-box', host); this.tags = $('.tags', host); this.inp = $('.ms-input', host);
    this.pop = $('.ms-pop', host); this.list = $('.ms-list', host); this.cnt = $('.cnt', host);
    this.box.addEventListener('mousedown', e => { if (e.target.closest('button')) return; if (e.target !== this.inp) { e.preventDefault(); this.inp.focus(); } this.open(); });
    this.inp.addEventListener('focus', () => this.open());
    this.inp.addEventListener('input', () => { this.act = 0; this.open(); this.renderList(); });
    this.inp.addEventListener('keydown', e => this.key(e));
    this.tags.addEventListener('click', e => { const b = e.target.closest('button[data-v]'); if (b) { e.stopPropagation(); toggleFacet(fk, b.dataset.v); } });
    this.list.addEventListener('mousedown', e => e.preventDefault());
    this.list.addEventListener('click', e => { const o = e.target.closest('.ms-opt'); if (o) toggleFacet(fk, o.dataset.v); });
    this.pop.addEventListener('mousedown', e => { if (e.target.closest('.ms-tools')) e.preventDefault(); });
    this.pop.querySelector('.ms-tools').addEventListener('click', e => {
      const a = e.target.closest('button'); if (!a) return;
      if (a.dataset.a === 'none') S.f[fk].clear(); else this.filtered().forEach(o => S.f[fk].add(o.v));
      S.page = 1; update();
    });
    document.addEventListener('mousedown', e => { if (!this.el.contains(e.target)) this.close(); });
  }
  filtered() {
    const t = fold(this.inp.value).trim();
    if (!t) return this.opts;
    const toks = t.split(/\s+/);
    return this.opts.filter(o => { const h = fold(o.label + ' ' + o.sub); return toks.every(x => h.includes(x)); })
      .sort((a, b) => (fold(b.label).startsWith(t) - fold(a.label).startsWith(t)) || b.count - a.count);
  }
  open() { if (!this.pop.classList.contains('hidden')) return; this.pop.classList.remove('hidden'); this.el.classList.add('open'); this.renderList(); }
  close() { this.pop.classList.add('hidden'); this.el.classList.remove('open'); this.inp.value = ''; }
  key(e) {
    const items = this.filtered();
    if (e.key === 'ArrowDown') { e.preventDefault(); this.open(); this.act = Math.min(items.length - 1, this.act + 1); this.renderList(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); this.act = Math.max(0, this.act - 1); this.renderList(); }
    else if (e.key === 'Enter') { e.preventDefault(); const o = items[this.act]; if (o) { this.inp.value = ''; toggleFacet(this.fk, o.v); } }
    else if (e.key === 'Escape') { this.close(); this.inp.blur(); }
    else if (e.key === 'Backspace' && !this.inp.value && S.f[this.fk].size) { const last = [...S.f[this.fk]].pop(); toggleFacet(this.fk, last); }
  }
  setOptions(opts) {
    this.opts = opts;
    const sel = S.f[this.fk];
    this.el.classList.toggle('active', sel.size > 0);
    const arr = [...sel];
    this.tags.innerHTML = arr.slice(0, 2).map(v => `<span class="ms-tag" title="${esc(facetLabel(this.fk, v))}"><span>${esc(facetLabel(this.fk, v))}</span><button data-v="${esc(v)}" title="Bỏ">✕</button></span>`).join('')
      + (arr.length > 2 ? `<span class="ms-more">+${arr.length - 2}</span>` : '');
    this.inp.placeholder = sel.size ? '' : this.f.ph;
    const nonzero = opts.filter(o => o.count).length;
    this.cnt.textContent = sel.size ? `${sel.size} đã chọn` : `${nonzero}`;
    if (!this.pop.classList.contains('hidden')) this.renderList();
  }
  renderList() {
    const items = this.filtered(); const sel = S.f[this.fk];
    const t = this.inp.value.trim();
    if (this.act >= items.length) this.act = Math.max(0, items.length - 1);
    const st = this.list.scrollTop;
    this.list.innerHTML = items.length ? items.map((o, i) => `<div class="ms-opt ${sel.has(o.v) ? 'sel' : ''} ${i === this.act ? 'act' : ''} ${o.count ? '' : 'zero'}" data-v="${esc(o.v)}">
      <span class="cb"></span>${o.dot ? `<span class="dot" style="background:${o.dot}"></span>` : ''}
      <span class="lb">${hl(o.label, t)}${o.sub ? `<small>${hl(o.sub, t)}</small>` : ''}</span><span class="c">${o.count}</span></div>`).join('')
      : `<div class="ms-empty">Không có kết quả khớp “${esc(t)}”</div>`;
    this.list.scrollTop = st;
    const a = this.list.querySelector('.act'); if (a) { const lt = this.list.scrollTop, lh = this.list.clientHeight; if (a.offsetTop < lt) this.list.scrollTop = a.offsetTop; else if (a.offsetTop + a.offsetHeight > lt + lh) this.list.scrollTop = a.offsetTop + a.offsetHeight - lh; }
  }
}
// Highlight không phân biệt dấu
function hl(text, q) {
  text = nfc(text);
  const toks = (Array.isArray(q) ? q : fold(q || '').split(/\s+/)).filter(t => t.length >= 1);
  if (!toks.length) return esc(text);
  const f = fold(text); const mark = new Uint8Array(text.length);
  for (const t of toks) { let i = f.indexOf(t); while (i >= 0) { mark.fill(1, i, i + t.length); i = f.indexOf(t, i + t.length); } }
  let out = '', on = false;
  for (let i = 0; i < text.length; i++) { if (mark[i] && !on) { out += '<mark>'; on = true; } if (!mark[i] && on) { out += '</mark>'; on = false; } out += esc(text[i]); }
  return out + (on ? '</mark>' : '');
}

/* ------------------------------------------------------------- quick chips */
const QUICK = [
  { t: 'Đang xử lý', set: { result: ['Đang xử lý'] }, c: '#f59e0b' },
  { t: 'Chờ hãng check', set: { block: ['Hãng check'] }, c: '#6b5bd6' },
  { t: 'Có cờ nhắc', set: { flag: ['1', '2', '3', '4'] }, c: '#f07c1b' },
  { t: 'Cờ ≥ Mức 2', set: { flag: ['2', '3', '4'] }, c: '#e03a3e' },
  { t: 'Timeline ≤ 30 ngày', set: { tl: ['d30'] }, c: '#2b7fd8' },
  { t: 'Đã qua timeline – chưa xong', set: { issue: ['tlPastOpen'] }, c: '#e5484d' },
  { t: 'Quá hạn SLA', set: { issue: ['slaOver'] }, c: '#8b1538' },
  { t: 'Trùng EU + Hãng', set: { issue: ['dup'] }, c: '#10b7de' },
  { t: 'Approved', set: { result: ['Approved'] }, c: '#1fa55b' },
  { t: 'Rejected', set: { result: ['Rejected'] }, c: '#e5484d' },
];
const quickOn = q => Object.entries(q.set).every(([k, vs]) => S.f[k].size === vs.length && vs.every(v => S.f[k].has(v)));
function quickCount(q) { return S.rows.filter(r => Object.entries(q.set).every(([k, vs]) => FACETS[k].vals(r).some(v => vs.includes(v)))).length; }
function renderQuick() {
  $('#quick').innerHTML = `<span class="lbl-inline">Lọc nhanh</span>` + QUICK.map((q, i) =>
    `<button class="pchip ${quickOn(q) ? 'on' : ''}" data-i="${i}"><span class="dot" style="background:${q.c}"></span>${esc(q.t)}<span class="n">${quickCount(q)}</span></button>`).join('');
}

/* -------------------------------------------------------------- render */
function update() {
  computeView();
  for (const fk of Object.keys(MS)) MS[fk].setOptions(facetOptions(fk));
  renderQuick(); renderActive(); renderAlert(); renderPanels(); renderKPIs(); renderInsights(); renderCharts(); renderTable();
}
function renderHeroSub() {
  if (!S.rows.length) { $('#hero-sub').textContent = 'Nạp file Excel xuất từ Base Service để bắt đầu. Dữ liệu chỉ xử lý trên máy của bạn.'; return; }
  const cs = S.rows.map(r => r.created).filter(Boolean).sort((a, b) => a - b);
  const lu = S.rows.map(r => r.lastUpdate).filter(Boolean).sort((a, b) => b - a)[0];
  $('#hero-sub').textContent = `${S.files.map(f => f.name).join(', ')} — ${fmtN(S.rows.length)} phiếu · tạo từ ${fmtD(cs[0])} đến ${fmtD(cs[cs.length - 1])} · cập nhật mới nhất ${fmtDT(lu)} · ngày mốc ${fmtD(S.ref)}`;
}
function renderFileBar() {
  $('#filebar').innerHTML = `<b style="color:var(--brand-900)">Dữ liệu nguồn:</b>` + S.files.map((f, i) =>
    `<span class="filechip" title="Sheet: ${esc(f.sheet)} · ${f.rows} dòng · ${f.added} phiếu mới · ${f.dup} trùng ID">📄 <b>${esc(f.name)}</b> · ${fmtN(f.rows)} dòng${f.dup ? ` · ${f.dup} trùng` : ''}</span>`).join('')
    + `<span class="muted">→ ${fmtN(S.rows.length)} phiếu duy nhất</span><span style="flex:1"></span>
    <button class="btn btn-sm" id="btn-add">＋ Nạp thêm file</button><button class="btn btn-sm btn-ghost" id="btn-clear">Xoá dữ liệu</button>`;
}
function renderActive() {
  const chips = [];
  if (S.q) chips.push(`<span class="fchip"><b>Từ khoá:</b> “${esc(S.q)}”<button data-clr="q">✕</button></span>`);
  if (S.from || S.to) { const lab = $('#date-field').selectedOptions[0].text; chips.push(`<span class="fchip"><b>${esc(lab)}:</b> ${S.from ? fmtD(S.from) : '…'} → ${S.to ? fmtD(S.to) : '…'}<button data-clr="date">✕</button></span>`); }
  for (const fk of Object.keys(FACETS)) for (const v of S.f[fk]) chips.push(`<span class="fchip"><b>${esc(FACETS[fk].label)}:</b> ${esc(facetLabel(fk, v))}<button data-fk="${fk}" data-v="${esc(v)}">✕</button></span>`);
  $('#activebar').innerHTML = `<span class="count">Đang hiển thị <span class="n">${fmtN(S.view.length)}</span> / ${fmtN(S.rows.length)} phiếu</span>`
    + (chips.length ? chips.join('') + `<button class="btn btn-sm btn-ghost" data-clr="all">Xoá tất cả</button>` : `<span class="muted">Chưa áp dụng bộ lọc nào</span>`);
  $('#sticky-count').textContent = `· ${fmtN(S.view.length)}/${fmtN(S.rows.length)} phiếu${chips.length ? ` · ${chips.length} bộ lọc` : ''}`;
  $('#tbl-hint').textContent = `${fmtN(S.view.length)} phiếu · bấm vào dòng để xem chi tiết · bấm tiêu đề cột để sắp xếp`;
}

const ICON_FLAG = '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M5 21V4h11l-1.5 4L16 12H7v9z"/></svg>';
const flagHtml = r => r.flag === 'na' ? '<span class="flag lna">—</span>' : r.flag === '0' ? '<span class="flag l0">● OK</span>' : `<span class="flag l${r.flag}">${ICON_FLAG}${FLAGS[r.flag].s}</span>`;
const resultBadge = r => `<span class="badge ${RESULT[r.result].cls}">${esc(r.result)}</span>`;
const blockBadge = b => b ? `<span class="badge" style="background:${blockColor(b)}1a;color:${blockColor(b)};border-color:${blockColor(b)}40">${esc(b)}</span>` : '';
const statusBadge = s => s ? `<span class="badge ${fold(s).includes('hoan thanh') ? 'b-teal' : 'b-warn'}">${esc(s)}</span>` : '';

function renderAlert() {
  const V = S.view; const fl = V.filter(r => r.flagN >= 1);
  const by = countBy(fl, r => r.flag);
  const el = $('#alert');
  if (!fl.length) {
    el.innerHTML = `<div class="alert ok"><div class="a-ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg></div>
      <div class="a-txt"><b>Tất cả phiếu trong phạm vi đang lọc đều được cập nhật trong vòng 1 tháng</b> (tính đến ${fmtD(S.ref)}).</div></div>`;
    return;
  }
  const oldest = fl.reduce((a, b) => (b.days > a.days ? b : a));
  el.innerHTML = `<div class="alert"><div class="a-ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l2 2"/><path d="M5 3 2 6"/><path d="m22 6-3-3"/></svg></div>
    <div class="a-txt"><b class="r">${fmtN(fl.length)}</b> phiếu đã <b>hơn 1 tháng</b> chưa được cập nhật (tính đến ${fmtD(S.ref)}). Lâu nhất: <b>#${esc(oldest.id)}</b> — ${fmtN(oldest.days)} ngày.
      <div class="lvls">${['4', '3', '2', '1'].filter(k => by.get(k)).map(k => `<button class="pchip ${S.f.flag.size === 1 && S.f.flag.has(k) ? 'on' : ''}" data-flag="${k}"><span class="dot" style="background:${FLAGS[k].c}"></span>${FLAGS[k].t}<span class="n">${by.get(k)}</span></button>`).join('')}</div></div>
    <button class="btn btn-primary" id="a-list">Xem danh sách</button>
    <button class="btn btn-accent" id="a-remind">✉ Soạn nội dung nhắc</button></div>`;
}

function renderPanels() {
  const V = S.view;
  $('#dc-total').textContent = fmtN(V.length);
  const res = countBy(V, r => r.result);
  const rk = Object.keys(RESULT).filter(k => res.get(k));
  chart('c-result', donutCfg(rk.map(k => RESULT[k].c), rk.map(k => res.get(k)), rk, (i) => toggleFacet('result', rk[i], true)));
  $('#lg-result').innerHTML = rk.map(k => `<li data-fk="result" data-v="${esc(k)}" class="${S.f.result.has(k) ? 'on' : ''}"><span class="sw" style="background:${RESULT[k].c}"></span>${esc(k)}<span class="v">${fmtN(res.get(k))} <small>(${pct(res.get(k), V.length)})</small></span></li>`).join('') || '<li>—</li>';
  const open = V.filter(r => r.open); const st = countBy(open, r => r.block || '(Trống)');
  $('#lg-steps').innerHTML = blockOrder().filter(b => st.get(b)).map(b => `<li data-fk="block" data-v="${esc(b)}" class="${S.f.block.has(b) ? 'on' : ''}"><span class="sw" style="background:${blockColor(b)}"></span>${esc(b)}<span class="v">${st.get(b)}</span></li>`).join('') || '<li style="opacity:.8">Không có phiếu đang xử lý</li>';
  const fl = countBy(V, r => r.flag);
  const fk = ['0', '1', '2', '3', '4', 'na'].filter(k => fl.get(k));
  $('#dc-flag').textContent = fmtN(V.filter(r => r.flagN >= 1).length);
  chart('c-flag', donutCfg(fk.map(k => FLAGS[k].c), fk.map(k => fl.get(k)), fk.map(k => FLAGS[k].t), i => toggleFacet('flag', fk[i], true)));
  $('#lg-flag').innerHTML = fk.map(k => `<li data-fk="flag" data-v="${k}" class="${S.f.flag.has(k) ? 'on' : ''}"><span class="sw" style="background:${FLAGS[k].c}"></span>${FLAGS[k].t}<span class="v">${fmtN(fl.get(k))} <small>(${pct(fl.get(k), V.length)})</small></span></li>`).join('');
  const tl = countBy(V, r => r.tl);
  $('#lg-tl').innerHTML = Object.keys(TLS).filter(k => tl.get(k)).map(k => `<li data-fk="tl" data-v="${k}" class="${S.f.tl.has(k) ? 'on' : ''}"><span class="sw" style="background:${TLS[k].c}"></span>${TLS[k].t}<span class="v">${tl.get(k)}</span></li>`).join('');
}

function renderKPIs() {
  const V = S.view, n = V.length;
  const res = countBy(V, r => r.result);
  const A = res.get('Approved') || 0, R = res.get('Rejected') || 0, O = res.get('Đang xử lý') || 0;
  const flg = V.filter(r => r.flagN >= 1).length;
  const sla = V.filter(r => r.issues.includes('slaOver')).length;
  const appr = V.filter(r => r.result === 'Approved' && r.procDays != null).map(r => r.procDays);
  const med = median(appr);
  const nSale = new Set(V.map(r => r.sale).filter(Boolean)).size, nV = new Set(V.flatMap(r => r.vendors)).size;
  const K = [
    { l: 'Tổng số phiếu', v: fmtN(n), s: `${nSale} sale · ${nV} hãng`, c: 'var(--brand-600)' },
    { l: 'Đang xử lý', v: fmtN(O), s: `${pct(O, n)} tổng số`, c: '#f59e0b', q: { result: ['Đang xử lý'] } },
    { l: 'Approved', v: fmtN(A), s: `${pct(A, n)} tổng số`, c: '#1fa55b', q: { result: ['Approved'] } },
    { l: 'Rejected', v: fmtN(R), s: `${pct(R, n)} tổng số`, c: '#e5484d', q: { result: ['Rejected'] } },
    { l: 'Tỷ lệ duyệt', v: pct(A, A + R), s: `trên ${fmtN(A + R)} phiếu đã có kết quả`, c: 'var(--navy-500)' },
    { l: 'Cần nhắc update', v: fmtN(flg), s: '> 1 tháng không cập nhật', c: '#f07c1b', q: { flag: ['1', '2', '3', '4'] } },
    { l: 'Quá hạn SLA', v: fmtN(sla), s: 'phiếu đang xử lý quá hạn', c: '#8b1538', q: { issue: ['slaOver'] } },
    { l: 'TG duyệt (trung vị)', v: med == null ? '—' : (med < 1 ? '< 1' : fmtN(Math.round(med))) , s: 'ngày, từ tạo → Approved', c: 'var(--brand-700)' },
  ];
  $('#kpis').innerHTML = K.map((k, i) => `<div class="kpi ${k.q ? 'click' : ''}" data-k="${i}" style="--c:${k.c}" ${k.q ? 'title="Bấm để lọc"' : ''}><div class="k-l">${k.l}</div><div class="k-v">${k.v}</div><div class="k-s">${k.s}</div></div>`).join('');
  S._kpi = K;
}

function renderInsights() {
  const V = S.view, n = V.length, out = [];
  if (!n) { $('#insights').innerHTML = '<li>Không có phiếu nào khớp bộ lọc.</li>'; return; }
  const res = countBy(V, r => r.result); const A = res.get('Approved') || 0, R = res.get('Rejected') || 0, O = res.get('Đang xử lý') || 0;
  out.push(`Có <b>${fmtN(n)}</b> phiếu: <b>${fmtN(A)}</b> Approved (${pct(A, n)}), <b class="r">${fmtN(R)}</b> Rejected, <b class="o">${fmtN(O)}</b> đang xử lý. Tỷ lệ duyệt <b>${pct(A, A + R)}</b>.`);
  const top = (m, lab, k = 3) => sortedEntries(m).filter(([x]) => x && x !== '__none').slice(0, k).map(([x, c]) => `<b>${esc(lab(x))}</b> (${c} · ${pct(c, n)})`).join(', ');
  out.push(`Hãng nhiều phiếu nhất: ${top(countBy(V, r => r.vendors), k => S.vendorLabel.get(k) || k)}.`);
  out.push(`Sale đăng ký nhiều nhất: ${top(countBy(V, r => r.sale), personLabel)}.`);
  out.push(`PM phụ trách nhiều nhất: ${top(countBy(V, r => r.pms), personLabel)}.`);
  const openBy = sortedEntries(countBy(V.filter(r => r.open), r => r.pms));
  if (openBy.length) out.push(`Tồn đọng nhiều nhất: <b>${esc(personLabel(openBy[0][0]))}</b> đang giữ <b class="o">${openBy[0][1]}</b> phiếu chưa hoàn thành.`);
  const fl = V.filter(r => r.flagN >= 1);
  if (fl.length) { const o = fl.reduce((a, b) => (b.days > a.days ? b : a)); out.push(`<b class="r">${fmtN(fl.length)}</b> phiếu hơn 1 tháng chưa cập nhật; lâu nhất <b>#${esc(o.id)}</b> ${esc(o.name)} — <b class="r">${o.days}</b> ngày (PM ${esc(o.pms.map(personLabel).join(', ') || '—')}).`); }
  const dec = new Map(); V.forEach(r => { if (r.result === 'Approved' || r.result === 'Rejected') r.vendors.forEach(v => { const x = dec.get(v) || { a: 0, r: 0 }; x[r.result === 'Approved' ? 'a' : 'r']++; dec.set(v, x); }); });
  const rej = [...dec.entries()].filter(([, x]) => x.a + x.r >= 5 && x.r).sort((a, b) => b[1].r / (b[1].a + b[1].r) - a[1].r / (a[1].a + a[1].r))[0];
  if (rej) out.push(`Hãng có tỷ lệ Rejected cao nhất: <b>${esc(S.vendorLabel.get(rej[0]))}</b> — ${rej[1].r}/${rej[1].a + rej[1].r} phiếu (${pct(rej[1].r, rej[1].a + rej[1].r)}).`);
  const past = V.filter(r => r.issues.includes('tlPastOpen')).length, soon = V.filter(r => r.tl === 'd30' && r.result !== 'Rejected').length;
  if (past || soon) out.push(`${past ? `<b class="r">${past}</b> phiếu đang xử lý đã qua timeline dự án` : ''}${past && soon ? '; ' : ''}${soon ? `<b class="o">${soon}</b> dự án có timeline trong 30 ngày tới` : ''}.`);
  const dupG = new Set(V.filter(r => r.issues.includes('dup')).map(r => r.euKey + '|' + r.vendors.join(',')));
  const conf = V.filter(r => r.issues.includes('conflict')).length;
  if (dupG.size) out.push(`<b>${dupG.size}</b> nhóm End user + Hãng được đăng ký nhiều lần${conf ? `, trong đó <b class="r">${conf}</b> phiếu do <b>sale khác nhau</b> đăng ký — cần kiểm tra xung đột` : ''}.`);
  const mm = sortedEntries(countBy(V.filter(r => r.month), r => r.month));
  if (mm.length) out.push(`Tháng có nhiều phiếu tạo mới nhất: <b>${monthLabel(mm[0][0])}</b> (${mm[0][1]} phiếu).`);
  const pd = median(V.filter(r => r.procDays != null).map(r => r.procDays));
  if (pd != null) out.push(`Thời gian xử lý trung vị của phiếu đã hoàn thành: <b>${pd < 1 ? '< 1' : Math.round(pd)}</b> ngày.`);
  const mis = V.filter(r => r.issues.includes('vendorMismatch')).length;
  if (mis) out.push(`<b class="o">${mis}</b> phiếu có tên hãng trong tiêu đề khác cột “Tên hãng” — nên rà soát lại dữ liệu.`);
  $('#insights').innerHTML = out.map(x => `<li>${x}</li>`).join('');
}

/* --------------------------------------------------------------- charts */
function chart(id, cfg) {
  const c = S.charts[id];
  if (c) { c.data = cfg.data; c.options = cfg.options; c.update(); return c; }
  return (S.charts[id] = new Chart(document.getElementById(id), cfg));
}
function donutCfg(colors, data, labels, onPick) {
  return {
    type: 'doughnut',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: '#fff', borderWidth: 2, hoverOffset: 6 }] },
    options: {
      cutout: '64%', responsive: true, maintainAspectRatio: false, animation: { duration: 500 },
      plugins: { legend: { display: false }, datalabels: { display: false }, tooltip: { callbacks: { label: c => ` ${c.label}: ${c.raw} (${pct(c.raw, c.dataset.data.reduce((a, b) => a + b, 0))})` } } },
      onClick: (e, els) => { if (els.length) onPick(els[0].index); },
    },
  };
}
function grad(color, horizontal) {
  return ctx => {
    const { chart } = ctx; const a = chart.chartArea; if (!a) return color;
    const g = horizontal ? chart.ctx.createLinearGradient(a.left, 0, a.right, 0) : chart.ctx.createLinearGradient(0, a.bottom, 0, a.top);
    g.addColorStop(0, color + 'bb'); g.addColorStop(1, color); return g;
  };
}
const shortName = s => { s = String(s); const p = s.split(' '); return s.length > 18 && p.length > 2 ? p.slice(-2).join(' ') : s; };
function barCfg({ labels, datasets, horizontal, stacked, onPick, fullLabels, legend = true, totals = true }) {
  const valueAxis = { stacked, beginAtZero: true, grid: { color: '#edf1f4' }, ticks: { precision: 0 } };
  const catAxis = { stacked, grid: { display: false }, ticks: { autoSkip: false, maxRotation: 50, minRotation: 0, font: { size: 11.5 }, callback: function (v) { const l = this.getLabelForValue(v); return l.length > 26 ? l.slice(0, 25) + '…' : l; } } };
  return {
    type: 'bar',
    data: { labels, datasets: datasets.map(d => ({ borderRadius: 4, borderSkipped: false, maxBarThickness: horizontal ? 20 : 42, ...d })) },
    options: {
      indexAxis: horizontal ? 'y' : 'x', responsive: true, maintainAspectRatio: false, animation: { duration: 450 },
      scales: horizontal ? { x: valueAxis, y: catAxis } : { x: catAxis, y: valueAxis },
      layout: { padding: { top: horizontal ? 0 : 18, right: horizontal ? 30 : 6 } },
      plugins: {
        legend: { display: legend && datasets.length > 1, position: 'bottom', labels: { usePointStyle: true, pointStyle: 'rectRounded', boxWidth: 10, padding: 14 } },
        tooltip: { mode: 'index', intersect: false, callbacks: { title: it => (fullLabels ? fullLabels[it[0].dataIndex] : it[0].label), footer: it => (stacked && it.length > 1 ? 'Tổng: ' + it.reduce((a, b) => a + b.raw, 0) : '') } },
        datalabels: {
          display: c => { const v = c.dataset.data[c.dataIndex]; if (!v) return false; if (stacked) { const tot = c.chart.data.datasets.reduce((a, d, i) => a + (c.chart.isDatasetVisible(i) ? (d.data[c.dataIndex] || 0) : 0), 0); return v / (tot || 1) > .12 && tot >= 4; } return true; },
          color: stacked ? '#fff' : '#334155', font: { weight: 700, size: 10.5 },
          anchor: stacked ? 'center' : 'end', align: stacked ? 'center' : (horizontal ? 'right' : 'top'), offset: 2, clamp: true,
        },
      },
      onHover: (e, els) => { e.native.target.style.cursor = els.length && onPick ? 'pointer' : 'default'; },
      onClick: (e, els) => { if (els.length && onPick) onPick(els[0].index); },
    },
    plugins: totals && stacked ? [stackTotals] : [],
  };
}
// Ghi tổng ở đầu cột xếp chồng
const stackTotals = {
  id: 'stackTotals',
  afterDatasetsDraw(ch) {
    const { ctx } = ch; const horiz = ch.options.indexAxis === 'y';
    const metas = ch.data.datasets.map((d, i) => ch.getDatasetMeta(i)).filter((m, i) => ch.isDatasetVisible(i));
    if (!metas.length) return;
    ctx.save(); ctx.font = '700 11px ' + Chart.defaults.font.family; ctx.fillStyle = '#1c2b36';
    ch.data.labels.forEach((l, i) => {
      const tot = ch.data.datasets.reduce((a, d, j) => a + (ch.isDatasetVisible(j) ? (d.data[i] || 0) : 0), 0); if (!tot) return;
      let edge = null; metas.forEach(m => { const b = m.data[i]; if (!b || !(m._parsed[i] && (horiz ? m._parsed[i].x : m._parsed[i].y))) return; edge = edge == null ? (horiz ? b.x : b.y) : horiz ? Math.max(edge, b.x) : Math.min(edge, b.y); });
      if (edge == null) return; const b0 = metas[0].data[i];
      if (horiz) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(tot, edge + 5, b0.y); }
      else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(tot, b0.x, edge - 4); }
    });
    ctx.restore();
  },
};
function topN(m, n) { const e = sortedEntries(m).filter(([k]) => k !== '__none'); return e.slice(0, n); }

function renderCharts() {
  const V = S.view;
  const resKeys = Object.keys(RESULT).filter(k => V.some(r => r.result === k));
  const stackByResult = (keyFn, n, labFn, fk) => {
    const tot = countBy(V, keyFn); const top = topN(tot, n).map(([k]) => k);
    const ds = resKeys.map(rk => ({ label: rk, data: top.map(k => V.filter(r => r.result === rk && [].concat(keyFn(r)).includes(k)).length), backgroundColor: grad(RESULT[rk].c) }));
    const labels = top.map(labFn);
    return barCfg({ labels: labels.map(shortName), fullLabels: labels, datasets: ds, stacked: true, onPick: i => toggleFacet(fk, top[i]) });
  };
  chart('c-sale', stackByResult(r => r.sale || '__none', 15, personLabel, 'sale'));
  // Hãng
  { const top = topN(countBy(V, r => r.vendors), 15); const labs = top.map(([k]) => S.vendorLabel.get(k) || k);
    chart('c-vendor', barCfg({ labels: labs, datasets: [{ label: 'Số phiếu', data: top.map(x => x[1]), backgroundColor: grad('#0b5cb5', true) }], horizontal: true, onPick: i => toggleFacet('vendor', top[i][0]) })); }
  // PM theo cờ
  { const tot = countBy(V, r => r.pms); const top = topN(tot, 14).map(([k]) => k);
    const fk = ['0', '1', '2', '3', '4', 'na'].filter(f => V.some(r => r.flag === f));
    const ds = fk.map(f => ({ label: FLAGS[f].t, data: top.map(k => V.filter(r => r.flag === f && r.pms.includes(k)).length), backgroundColor: FLAGS[f].c }));
    const labels = top.map(personLabel);
    chart('c-pm', barCfg({ labels: labels.map(shortName), fullLabels: labels, datasets: ds, stacked: true, onPick: i => toggleFacet('pm', top[i]) })); }
  // Khối
  { const cnt = countBy(V, r => r.block); const bo = blockOrder().filter(b => cnt.get(b));
    chart('c-block', barCfg({ labels: bo, datasets: [{ label: 'Số phiếu', data: bo.map(b => cnt.get(b)), backgroundColor: bo.map(blockColor) }], horizontal: true, onPick: i => toggleFacet('block', bo[i]) })); }
  // Xu hướng
  { const g = S.trendGran; const kf = g === 'week' ? weekKey : g === 'quarter' ? quarterKey : monthKey;
    const keys = [...new Set(V.filter(r => r.created).map(r => kf(r.created)))].sort();
    const lab = k => (g === 'week' ? 'Tuần ' + fmtD(new Date(k)).slice(0, 5) : g === 'quarter' ? quarterLabel(k) : monthLabel(k));
    const cnt = f => keys.map(k => V.filter(r => r.created && kf(r.created) === k && f(r)).length);
    const ln = (label, data, c, fill) => ({ type: 'line', label, data, borderColor: c, backgroundColor: fill ? (ctx => { const a = ctx.chart.chartArea; if (!a) return c + '22'; const gr = ctx.chart.ctx.createLinearGradient(0, a.top, 0, a.bottom); gr.addColorStop(0, c + '55'); gr.addColorStop(1, c + '00'); return gr; }) : c, fill, tension: .35, pointRadius: 3.5, pointHoverRadius: 6, borderWidth: 2.5, pointBackgroundColor: '#fff', pointBorderWidth: 2 });
    const cfg = barCfg({ labels: keys.map(lab), datasets: [], stacked: false, totals: false });
    cfg.type = 'line';
    cfg.data.datasets = [ln('Tạo mới', cnt(() => true), '#0b5cb5', true), ln('Approved', cnt(r => r.result === 'Approved'), '#1fa55b'), ln('Rejected', cnt(r => r.result === 'Rejected'), '#e5484d'), ln('Đang xử lý', cnt(r => r.open), '#f59e0b')];
    cfg.options.plugins.legend.display = true;
    cfg.options.plugins.datalabels = { display: c => c.datasetIndex === 0 && keys.length <= 24, align: 'top', color: '#0b5cb5', font: { weight: 700, size: 10.5 } };
    cfg.options.interaction = { mode: 'index', intersect: false };
    chart('c-trend', cfg); }
  // Người tạo
  { const top = topN(countBy(V, r => r.creator || '__none'), 15); const labs = top.map(([k]) => personLabel(k));
    chart('c-creator', barCfg({ labels: labs.map(shortName), fullLabels: labs, datasets: [{ label: 'Số phiếu', data: top.map(x => x[1]), backgroundColor: grad('#10b7de', true) }], horizontal: true, onPick: i => toggleFacet('creator', top[i][0]) })); }
  // Pipeline theo quý timeline
  { const P = V.filter(r => r.timeline && r.result !== 'Rejected');
    const keys = [...new Set(P.map(r => quarterKey(r.timeline)))].sort();
    const ds = ['Approved', 'Đang xử lý', 'Hoàn thành khác'].filter(k => P.some(r => r.result === k)).map(k => ({ label: k, data: keys.map(q => P.filter(r => r.result === k && quarterKey(r.timeline) === q).length), backgroundColor: grad(RESULT[k].c) }));
    const refQ = quarterKey(S.ref);
    const cfg = barCfg({ labels: keys.map(k => quarterLabel(k) + (k === refQ ? ' ◂' : '')), datasets: ds, stacked: true, onPick: i => { const [y, q] = keys[i].split('-Q'); S.dateField = 'timeline'; $('#date-field').value = 'timeline'; S.from = new Date(+y, (q - 1) * 3, 1); S.to = new Date(+y, q * 3, 0); $('#date-from').value = isoD(S.from); $('#date-to').value = isoD(S.to); S.page = 1; update(); } });
    chart('c-pipe', cfg); }
  // Reseller
  { const top = topN(countBy(V, r => r.resellers), 12); const labs = top.map(([k]) => S.resellerLabel.get(k) || k);
    chart('c-reseller', barCfg({ labels: labs, datasets: [{ label: 'Số phiếu', data: top.map(x => x[1]), backgroundColor: grad('#f26522', true) }], horizontal: true, onPick: i => toggleFacet('reseller', top[i][0]) })); }
}

/* ---------------------------------------------------------------- table */
const persons = ks => ks.map(personLabel).join(', ');
const COLS = [
  { k: 'flag', t: 'Cờ nhắc', sv: r => r.flagN * 1e6 + (r.days || 0), td: r => `<td class="nowrap">${flagHtml(r)}<div class="days">${r.days == null ? '' : r.days + ' ngày'}</div></td>` },
  { k: 'id', t: 'ID', sv: r => +r.id || 0, td: (r, q) => `<td class="id">${hl(r.id, q)}</td>` },
  { k: 'name', t: 'Tên phiếu', sv: r => fold(r.name), td: (r, q) => `<td class="name">${hl(r.name, q)}${r.project ? `<span class="sub">${hl(r.project, q)}</span>` : ''}</td>` },
  { k: 'eu', t: 'End user', sv: r => fold(r.eu), td: (r, q) => `<td class="eu">${hl(r.eu, q)}</td>` },
  { k: 'vendor', t: 'Hãng', sv: r => fold(r.vendors.map(v => S.vendorLabel.get(v)).join()), td: r => `<td class="nowrap"><b>${esc(r.vendors.map(v => S.vendorLabel.get(v)).join(', '))}</b>${r.issues.includes('vendorMismatch') ? ` <span title="Tên phiếu ghi hãng ${esc(r.nameVendor)}">⚠</span>` : ''}</td>` },
  { k: 'sale', t: 'Sale', sv: r => fold(personLabel(r.sale)), td: r => `<td class="nowrap">${esc(r.sale ? personLabel(r.sale) : '')}</td>` },
  { k: 'pm', t: 'PM phụ trách', sv: r => fold(persons(r.pms)), td: r => `<td class="nowrap">${esc(persons(r.pms))}</td>` },
  { k: 'creator', t: 'Người tạo', sv: r => fold(personLabel(r.creator)), td: r => `<td class="nowrap">${esc(personLabel(r.creator))}</td>` },
  { k: 'block', t: 'Tiến độ', sv: r => blockOrder().indexOf(r.block), td: r => `<td>${blockBadge(r.block)}</td>` },
  { k: 'status', t: 'Trạng thái', sv: r => r.status, td: r => `<td>${statusBadge(r.status)}${r.issues.includes('slaOver') ? '<div class="days" style="color:var(--bad)">Quá hạn SLA</div>' : ''}</td>` },
  { k: 'reseller', t: 'Reseller', sv: r => fold(r.resellerRaw), td: r => `<td>${esc(r.resellerRaw)}</td>` },
  { k: 'project', t: 'Tên dự án', sv: r => fold(r.project), td: (r, q) => `<td>${hl(r.project, q)}</td>` },
  { k: 'timeline', t: 'Timeline DA', sv: r => r.timeline ? +r.timeline : null, td: r => `<td class="nowrap">${r.timeline ? fmtD(r.timeline) : `<span class="muted">${esc(r.timelineRaw)}</span>`}${r.timeline ? `<div class="days" style="color:${TLS[r.tl].c}">${TLS[r.tl].t}</div>` : ''}</td>` },
  { k: 'created', t: 'Ngày tạo', sv: r => r.created ? +r.created : null, td: r => `<td class="nowrap num">${fmtDT(r.created)}</td>` },
  { k: 'lastUpdate', t: 'Cập nhật cuối', sv: r => r.lastUpdate ? +r.lastUpdate : null, td: r => `<td class="nowrap num">${fmtDT(r.lastUpdate)}</td>` },
  { k: 'sla', t: 'Hạn SLA', sv: r => r.sla ? +r.sla : null, td: r => `<td class="nowrap num">${fmtDT(r.sla)}</td>` },
  { k: 'followers', t: 'Người theo dõi', sv: r => r.followers.length, td: r => `<td>${esc(persons(r.followers))}</td>` },
  { k: 'note', t: 'Note', sv: r => fold(r.note), td: (r, q) => `<td style="min-width:200px">${hl(r.note, q)}</td>` },
];
function renderTable() {
  const cols = COLS.filter(c => !S.hiddenCols.has(c.k));
  $('#grid thead').innerHTML = '<tr>' + cols.map(c => `<th data-k="${c.k}" class="${S.sort.k === c.k ? 'sorted' : ''}">${esc(c.t)}<span class="ar">${S.sort.k === c.k ? (S.sort.d > 0 ? '▲' : '▼') : '↕'}</span></th>`).join('') + '</tr>';
  const n = S.view.length, ps = S.pageSize, pages = Math.max(1, Math.ceil(n / ps));
  if (S.page > pages) S.page = pages;
  const start = (S.page - 1) * ps, slice = S.view.slice(start, start + ps);
  const q = qTokens();
  $('#grid tbody').innerHTML = slice.length ? slice.map(r => `<tr data-id="${esc(r.id)}">${cols.map(c => c.td(r, q)).join('')}</tr>`).join('')
    : `<tr><td colspan="${cols.length}"><div class="empty-state">Không có phiếu nào khớp bộ lọc. <button class="btn btn-sm" data-clr="all">Xoá bộ lọc</button></div></td></tr>`;
  $('#page-info').textContent = n ? `${fmtN(start + 1)}–${fmtN(Math.min(n, start + ps))} / ${fmtN(n)} phiếu` : '';
  const btn = (p, l, dis, on) => `<button data-p="${p}" ${dis ? 'disabled' : ''} class="${on ? 'on' : ''}">${l}</button>`;
  let h = btn(S.page - 1, '‹', S.page === 1);
  const win = new Set([1, pages, S.page - 1, S.page, S.page + 1, S.page - 2, S.page + 2].filter(p => p >= 1 && p <= pages));
  let prev = 0; [...win].sort((a, b) => a - b).forEach(p => { if (p - prev > 1) h += '<span class="muted">…</span>'; h += btn(p, p, false, p === S.page); prev = p; });
  h += btn(S.page + 1, '›', S.page === pages);
  $('#pager').innerHTML = h;
  $('#cols-menu').innerHTML = COLS.map(c => `<label><input type="checkbox" data-col="${c.k}" ${S.hiddenCols.has(c.k) ? '' : 'checked'}> ${esc(c.t)}</label>`).join('');
}

/* --------------------------------------------------------------- drawer */
function openDetail(id) {
  const r = S.byIdRow.get(id); if (!r) return;
  const idx = S.view.indexOf(r);
  const bo = blockOrder().filter(b => b !== 'Approved' && b !== 'Rejected');
  const cur = r.block; const curIdx = bo.indexOf(cur);
  const steps = bo.map((b, i) => {
    let cls = '';
    if (r.result === 'Approved' || r.result === 'Rejected' || r.result === 'Hoàn thành khác') cls = 'done'; else if (curIdx >= 0) cls = i < curIdx ? 'done' : i === curIdx ? 'cur' : '';
    return `<div class="step ${cls}"><div class="c">${cls === 'done' ? '✓' : i + 1}</div>${esc(b)}</div>`;
  }).join('') + (r.result === 'Rejected' ? `<div class="step badend"><div class="c">✕</div>Rejected</div>`
    : r.result === 'Approved' ? `<div class="step okend"><div class="c">✓</div>Approved</div>`
    : r.result === 'Hoàn thành khác' ? `<div class="step okend"><div class="c">✓</div>${esc(cur || 'Hoàn thành')}</div>`
    : `<div class="step"><div class="c">★</div>Approved / Rejected</div>`);
  const kv = (k, v, full) => `<div class="${full ? 'full' : ''}"><div class="k">${k}</div><div class="v">${v || '<span class="muted">—</span>'}</div></div>`;
  const link = u => u ? `<a href="${esc(/^https?:/i.test(u) ? u : 'https://' + u)}" target="_blank" rel="noopener noreferrer">${esc(u)}</a>` : '';
  const people = ks => ks.map(k => `${esc(personLabel(k))}${personSub(k) ? ` <span class="muted">${esc(personSub(k))}</span>` : ''}`).join('<br>');
  const html = `<div class="overlay" id="ov"><div class="drawer" role="dialog" aria-modal="true">
    <div class="drawer-h">
      <div class="top"><span class="idtag">#${esc(r.id)}</span>${r.vendors.map(v => `<span class="badge" style="background:#fff;color:var(--navy-700)">${esc(S.vendorLabel.get(v))}</span>`).join('')}
        <div class="x"><button id="d-prev" title="Phiếu trước (←)" ${idx <= 0 ? 'disabled style="opacity:.4"' : ''}>‹</button><button id="d-next" title="Phiếu sau (→)" ${idx < 0 || idx >= S.view.length - 1 ? 'disabled style="opacity:.4"' : ''}>›</button><button id="d-copy" title="Sao chép tóm tắt">⧉</button><button id="d-close" title="Đóng (Esc)">✕</button></div></div>
      <h2>${esc(r.name)}</h2>
      <div class="meta">${resultBadge(r)} ${r.status && r.status !== r.result ? statusBadge(r.status) : ''} ${blockBadge(r.block)} ${flagHtml(r).replace('flag l0', 'flag l0" style="color:#fff')} ${r.days != null ? `<span class="badge" style="background:rgba(255,255,255,.18);color:#fff;border-color:rgba(255,255,255,.3)">${r.days} ngày chưa cập nhật</span>` : ''}</div>
    </div>
    <div class="drawer-b">
      <div class="dsec"><h4>Tiến độ xử lý</h4><div class="in"><div class="stepper">${steps}</div></div></div>
      <div class="dsec"><h4>Mốc thời gian</h4><div class="in"><div class="timeline-mini">
        <div class="tm"><div class="k">Tạo phiếu</div><div class="v">${fmtD(r.created) || '—'}</div><div class="s">${r.created ? fmtDT(r.created).slice(0, 5) + ` · ${r.age} ngày trước` : ''}</div></div>
        <div class="tm"><div class="k">Cập nhật cuối</div><div class="v">${fmtD(r.lastUpdate) || '—'}</div><div class="s">${r.days != null ? r.days + ' ngày trước' : ''}</div></div>
        <div class="tm"><div class="k">Timeline dự án</div><div class="v">${r.timeline ? fmtD(r.timeline) : esc(r.timelineRaw) || '—'}</div><div class="s" style="color:${TLS[r.tl].c}">${TLS[r.tl].t}</div></div>
        <div class="tm"><div class="k">Hạn SLA</div><div class="v">${fmtD(r.sla) || '—'}</div><div class="s" ${r.issues.includes('slaOver') ? 'style="color:var(--bad)"' : ''}>${r.overdue ? 'Quá hạn: ' + esc(r.overdue) : ''}</div></div>
      </div></div></div>
      ${r.issues.length ? `<div class="dsec"><h4>⚠ Cảnh báo</h4><div class="in issues">${r.issues.map(i => `<span class="badge ${i === 'conflict' || i === 'slaOver' || i === 'tlPastOpen' ? 'b-bad' : 'b-warn'}">${esc(ISSUES[i])}${i === 'vendorMismatch' ? ': ' + esc(r.nameVendor) : ''}</span>`).join('')}</div></div>` : ''}
      <div class="dsec"><h4>Khách hàng &amp; dự án</h4><div class="in kv">
        ${kv('End user', `<b>${esc(r.eu)}</b>`)}${kv('Email End user', r.euEmail ? `<a href="mailto:${esc(r.euEmail)}">${esc(r.euEmail)}</a>` : '')}
        ${kv('Tên dự án', esc(r.project))}${kv('Website', link(r.website))}
        ${kv('Reseller / Partner', esc(r.resellerRaw))}${kv('Hãng', esc(r.vendors.map(v => S.vendorLabel.get(v)).join(', ')))}
        ${kv('Thông tin khách hàng', esc(r.custInfo), true)}
        ${kv('Thông tin BOM', r.bom ? `<pre class="pre">${esc(r.bom)}</pre>` : '', true)}
      </div></div>
      <div class="dsec"><h4>Phụ trách</h4><div class="in kv">
        ${kv('Sale', r.sale ? people([r.sale]) : '')}${kv('Người tạo', people(r.creator ? [r.creator] : []))}
        ${kv('PM phụ trách (người thực hiện)', people(r.pms))}${kv('Người theo dõi', people(r.followers))}
      </div></div>
      ${r.note || r.crm ? `<div class="dsec"><h4>Note &amp; CRM</h4><div class="in kv">${kv('Note', r.note ? `<pre class="pre">${esc(r.note)}</pre>` : '', true)}${kv('Link Deal CRM', link(r.crm), true)}</div></div>` : ''}
      ${r.desc ? `<div class="dsec"><h4>Mô tả phiếu</h4><div class="in"><pre class="pre">${esc(r.desc)}</pre></div></div>` : ''}
      ${r.attachments.length ? `<div class="dsec"><h4>Tệp đính kèm</h4><div class="in att">${r.attachments.map(a => a.url ? `<a href="${esc(a.url)}" target="_blank" rel="noopener noreferrer">📎 ${esc(a.name)}</a>` : `<span>📎 ${esc(a.name)}</span>`).join('')}</div></div>` : ''}
      ${r.related.length ? `<div class="dsec"><h4>Phiếu khác cùng End user (${r.related.length})</h4><div class="in"><ul class="rel">${r.related.sort((a, b) => (b.created || 0) - (a.created || 0)).map(x => `<li data-open="${esc(x.id)}"><span class="id">#${esc(x.id)}</span><span style="flex:1">${esc(x.name)}<br><span class="muted" style="font-size:12px">${esc(personLabel(x.sale))} · ${fmtD(x.created)}</span></span>${resultBadge(x)}</li>`).join('')}</ul></div></div>` : ''}
      <div class="dsec"><h4>Thông tin hệ thống</h4><div class="in kv">
        ${kv('Dịch vụ', esc(r.service))}${kv('ID khối / Loại khối', esc([r.blockId, r.blockType].filter(Boolean).join(' · ')))}
        ${kv('Tạo khối lúc', fmtDT(r.blockCreated))}${kv('Cập nhật khối lần cuối', fmtDT(r.blockUpdated))}
        ${kv('Cập nhật phiếu lần cuối', fmtDT(r.updated))}${kv('Nguồn', esc(r.file))}
      </div></div>
    </div></div></div>`;
  $('#layer').innerHTML = html;
  document.body.style.overflow = 'hidden';
  const close = () => { $('#layer').innerHTML = ''; document.body.style.overflow = ''; document.removeEventListener('keydown', onKey); };
  const go = d => { const j = idx + d; if (idx >= 0 && j >= 0 && j < S.view.length) { document.removeEventListener('keydown', onKey); openDetail(S.view[j].id); } };
  const onKey = e => { if (e.key === 'Escape') close(); else if (e.key === 'ArrowLeft') go(-1); else if (e.key === 'ArrowRight') go(1); };
  document.addEventListener('keydown', onKey);
  $('#ov').addEventListener('mousedown', e => { if (e.target.id === 'ov') close(); });
  $('#d-close').onclick = close; $('#d-prev').onclick = () => go(-1); $('#d-next').onclick = () => go(1);
  $('#d-copy').onclick = () => copyText(summaryText(r)).then(() => toast('Đã sao chép tóm tắt phiếu'));
  $$('#ov [data-open]').forEach(li => li.onclick = () => { document.removeEventListener('keydown', onKey); openDetail(li.dataset.open); });
}
function summaryText(r) {
  return [`#${r.id} — ${r.name}`, `Hãng: ${r.vendors.map(v => S.vendorLabel.get(v)).join(', ')} | End user: ${r.eu}`, r.project ? `Dự án: ${r.project}` : '', `Sale: ${personLabel(r.sale)} | PM: ${persons(r.pms)} | Reseller: ${r.resellerRaw || '—'}`,
    `Tiến độ: ${r.block} | Trạng thái: ${r.status} | Kết quả: ${r.result}`, `Timeline: ${r.timeline ? fmtD(r.timeline) : r.timelineRaw || '—'} | Tạo: ${fmtD(r.created)} | Cập nhật cuối: ${fmtD(r.lastUpdate)} (${r.days} ngày)`,
    r.bom ? `BOM: ${r.bom.replace(/\s*\n\s*/g, '; ')}` : '', r.note ? `Note: ${r.note}` : '', r.crm ? `CRM: ${r.crm}` : ''].filter(Boolean).join('\n');
}
async function copyText(t) {
  try { await navigator.clipboard.writeText(t); }
  catch (e) { const ta = document.createElement('textarea'); ta.value = t; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove(); }
}

/* ------------------------------------------------------------- reminder */
function openReminder() {
  $('#layer').innerHTML = `<div class="overlay" id="ov" style="justify-content:center;align-items:center"><div class="modal">
    <div class="drawer-h"><div class="top"><h2>✉ Soạn nội dung nhắc cập nhật tiến độ</h2><div class="x"><button id="r-close">✕</button></div></div>
      <div style="opacity:.9;font-size:13px">Tạo sẵn nội dung để gửi qua Base / email / chat. Dựa trên các phiếu đang lọc.</div></div>
    <div class="mb">
      <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px">
        <div class="field"><label>Nhóm theo</label><select class="select" id="r-group"><option value="pm">PM phụ trách</option><option value="sale">Sale</option><option value="vendor">Hãng</option></select></div>
        <div class="field"><label>Từ mức cờ</label><select class="select" id="r-min"><option value="1">Mức 1 (> 1 tháng)</option><option value="2">Mức 2 (> 3 tháng)</option><option value="3">Mức 3 (> 6 tháng)</option><option value="4">Mức 4 (> 1 năm)</option></select></div>
        <div class="field"><label>Phiếu</label><select class="select" id="r-open"><option value="open">Chỉ phiếu đang xử lý</option><option value="all">Tất cả phiếu có cờ</option></select></div>
      </div>
      <textarea id="r-text" spellcheck="false"></textarea>
    </div>
    <div class="mf"><span class="muted" id="r-count"></span><span style="flex:1"></span><button class="btn" id="r-dl">⬇ Tải .txt</button><button class="btn btn-primary" id="r-copy">⧉ Sao chép</button></div>
  </div></div>`;
  const gen = () => {
    const g = $('#r-group').value, min = +$('#r-min').value, onlyOpen = $('#r-open').value === 'open';
    const rows = S.view.filter(r => r.flagN >= min && (!onlyOpen || r.open)).sort((a, b) => b.days - a.days);
    const key = r => (g === 'pm' ? (r.pms.length ? r.pms : ['__none']) : g === 'sale' ? [r.sale || '__none'] : (r.vendors.length ? r.vendors : ['__none']));
    const lab = k => (k === '__none' ? '(Chưa xác định)' : g === 'vendor' ? S.vendorLabel.get(k) : personLabel(k) + (personSub(k) ? ` (${personSub(k)})` : ''));
    const groups = new Map(); rows.forEach(r => key(r).forEach(k => { if (!groups.has(k)) groups.set(k, []); groups.get(k).push(r); }));
    const lines = [`[NHẮC CẬP NHẬT TIẾN ĐỘ DEALREG] — tính đến ${fmtD(S.ref)}`, `Có ${rows.length} phiếu ${onlyOpen ? 'đang xử lý ' : ''}chưa được cập nhật ${FLAGS[min].t.split('·')[1].trim()}. Nhờ anh/chị cập nhật tiến độ giúp:`, ''];
    [...groups.entries()].sort((a, b) => b[1].length - a[1].length).forEach(([k, rs]) => {
      lines.push(`▶ ${lab(k)} — ${rs.length} phiếu`);
      rs.forEach(r => lines.push(`   • #${r.id} ${r.name} | Hãng: ${r.vendors.map(v => S.vendorLabel.get(v)).join(', ') || '—'} | Sale: ${personLabel(r.sale)} | Tiến độ: ${r.block || '—'} | ${r.days} ngày chưa cập nhật (${FLAGS[r.flag].s})`));
      lines.push('');
    });
    lines.push('Trân trọng,', 'Phòng Quản lý sản phẩm — NTS Hanoi Corp.');
    $('#r-text').value = lines.join('\n');
    $('#r-count').textContent = `${rows.length} phiếu · ${groups.size} nhóm`;
  };
  ['#r-group', '#r-min', '#r-open'].forEach(s => $(s).onchange = gen); gen();
  const close = () => { $('#layer').innerHTML = ''; };
  $('#r-close').onclick = close; $('#ov').addEventListener('mousedown', e => { if (e.target.id === 'ov') close(); });
  $('#r-copy').onclick = () => copyText($('#r-text').value).then(() => toast('Đã sao chép nội dung nhắc'));
  $('#r-dl').onclick = () => download(new Blob(['﻿' + $('#r-text').value], { type: 'text/plain;charset=utf-8' }), `Nhac cap nhat Dealreg ${fmtD(S.ref).replace(/\//g, '.')}.txt`);
}
function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500); }

/* ---------------------------------------------------------------- search */
function renderSuggest() {
  const box = $('#suggest'); const q = S.q.trim();
  if (!q || document.activeElement !== $('#q')) { box.classList.add('hidden'); return; }
  const toks = qTokens(); const fq = fold(q).trim();
  const score = r => (r.id === q ? 1000 : r.id.startsWith(q) ? 800 : 0) + (fold(r.name).includes(fq) ? 300 : 0) + (fold(r.eu).includes(fq) ? 200 : 0) + (fold(r.project).includes(fq) ? 150 : 0) + (r.created ? +r.created / 1e13 : 0);
  const tickets = S.view.map(r => [r, score(r)]).sort((a, b) => b[1] - a[1]).slice(0, 8).map(x => x[0]);
  const ent = [];
  const addEnt = (fk, k, lab, sub) => { const h = fold(lab + ' ' + (sub || '')); if (toks.every(t => h.includes(t)) && !S.f[fk].has(k)) ent.push({ fk, k, lab, sub }); };
  new Set(S.rows.map(r => r.sale).filter(Boolean)).forEach(k => addEnt('sale', k, personLabel(k), personSub(k)));
  new Set(S.rows.flatMap(r => r.pms)).forEach(k => addEnt('pm', k, personLabel(k), personSub(k)));
  new Set(S.rows.map(r => r.creator).filter(Boolean)).forEach(k => addEnt('creator', k, personLabel(k), personSub(k)));
  S.vendorLabel.forEach((lab, k) => addEnt('vendor', k, lab));
  S.resellerLabel.forEach((lab, k) => addEnt('reseller', k, lab));
  const items = [];
  let h = '';
  if (ent.length) { h += '<div class="grp">Lọc theo người / hãng / đối tác</div>'; ent.slice(0, 6).forEach(e => { items.push(e); h += `<div class="it" data-i="${items.length - 1}"><span class="t">${hl(e.lab, toks)}${e.sub ? ` <small style="display:inline">${hl(e.sub, toks)}</small>` : ''}</span><span class="tag">${esc(FACETS[e.fk].label)}</span></div>`; }); }
  h += `<div class="grp">Phiếu khớp (${fmtN(S.view.length)}) — Enter để xem trong bảng</div>`;
  if (!tickets.length) h += '<div class="empty">Không tìm thấy phiếu nào.</div>';
  tickets.forEach(r => { items.push({ r }); h += `<div class="it" data-i="${items.length - 1}"><span class="id">${hl(r.id, toks)}</span><span class="t">${hl(r.name, toks)}<small>${hl(r.eu, toks)} · ${esc(personLabel(r.sale))} · ${esc(r.block)}</small></span>${resultBadge(r)}</div>`; });
  box.innerHTML = h; box.classList.remove('hidden'); S._sug = items; S._sugAct = -1;
}
function pickSuggest(i) {
  const it = S._sug && S._sug[i]; if (!it) return;
  $('#suggest').classList.add('hidden');
  if (it.r) { openDetail(it.r.id); return; }
  S.q = ''; $('#q').value = ''; $('#searchbox').classList.remove('has-val');
  S.f[it.fk].add(it.k); S.page = 1; update();
}

/* ---------------------------------------------------------------- export */
const XS = {
  border: { top: { style: 'thin', color: { rgb: 'CFD9E0' } }, bottom: { style: 'thin', color: { rgb: 'CFD9E0' } }, left: { style: 'thin', color: { rgb: 'CFD9E0' } }, right: { style: 'thin', color: { rgb: 'CFD9E0' } } },
};
const F = (o = {}) => ({ name: 'Calibri', sz: 11, ...o });
const ST = {
  head: { font: F({ bold: true, color: { rgb: 'FFFFFF' } }), fill: { patternType: 'solid', fgColor: { rgb: '0B5CB5' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: XS.border },
  cell: { font: F(), alignment: { vertical: 'top', wrapText: false }, border: XS.border },
  wrap: { font: F(), alignment: { vertical: 'top', wrapText: true }, border: XS.border },
  center: { font: F(), alignment: { vertical: 'top', horizontal: 'center' }, border: XS.border },
  bold: { font: F({ bold: true }), alignment: { vertical: 'top' }, border: XS.border },
  link: { font: F({ color: { rgb: '2B63C6' }, underline: true }), alignment: { vertical: 'top' }, border: XS.border },
  title: { font: F({ bold: true, sz: 18, color: { rgb: '0A4F9E' } }) },
  sub: { font: F({ italic: true, color: { rgb: '62727F' } }) },
  sect: { font: F({ bold: true, sz: 12, color: { rgb: 'FFFFFF' } }), fill: { patternType: 'solid', fgColor: { rgb: 'F26522' } }, alignment: { vertical: 'center' } },
  kl: { font: F({ color: { rgb: '1C2B36' } }), fill: { patternType: 'solid', fgColor: { rgb: 'EEF6FE' } }, border: XS.border, alignment: { vertical: 'center' } },
  kv: { font: F({ bold: true, sz: 12, color: { rgb: '0B5CB5' } }), border: XS.border, alignment: { horizontal: 'right', vertical: 'center' } },
  total: { font: F({ bold: true }), fill: { patternType: 'solid', fgColor: { rgb: 'D5E8FB' } }, border: XS.border },
  foot: { font: F({ italic: true, color: { rgb: '0B5CB5' } }) },
};
const FLAG_FILL = { '1': ['F2B705', '3D2E00'], '2': ['F07C1B', 'FFFFFF'], '3': ['E03A3E', 'FFFFFF'], '4': ['8B1538', 'FFFFFF'], '0': ['E3F6EC', '137A41'], na: ['EEF2F5', '62727F'] };
const RES_FILL = { 'Approved': ['E3F6EC', '137A41'], 'Rejected': ['FDE8E8', 'B3262B'], 'Đang xử lý': ['FFF4DC', '92590A'], 'Hoàn thành khác': ['EEF2F5', '62727F'] };
const fillStyle = ([bg, fg], bold = true) => ({ font: F({ bold, color: { rgb: fg } }), fill: { patternType: 'solid', fgColor: { rgb: bg } }, alignment: { vertical: 'top', horizontal: 'center' }, border: XS.border });
const serial = d => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - Date.UTC(1899, 11, 30)) / DAY;

class SheetB {
  constructor() { this.ws = {}; this.R = 0; this.C = 0; this.merges = []; this.rowsH = []; }
  set(r, c, v, s, opt = {}) {
    let cell;
    if (v instanceof Date) { const z = opt.z || 'dd/mm/yyyy hh:mm'; cell = { t: 'n', v: serial(v), z, s: { ...s, numFmt: z } }; }
    else if (typeof v === 'number' && isFinite(v)) { cell = { t: 'n', v, s: opt.z ? { ...s, numFmt: opt.z } : s }; if (opt.z) cell.z = opt.z; }
    else { cell = { t: 's', v: v == null ? '' : String(v).slice(0, 32000), s }; }
    if (opt.link) cell.l = { Target: opt.link, Tooltip: 'Mở liên kết' };
    this.ws[XLSX.utils.encode_cell({ r, c })] = cell;
    this.R = Math.max(this.R, r); this.C = Math.max(this.C, c);
    return this;
  }
  merge(r1, c1, r2, c2) { this.merges.push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } }); }
  table(r0, c0, headers, rows, colSpec = []) {
    headers.forEach((h, j) => this.set(r0, c0 + j, h, ST.head));
    rows.forEach((row, i) => row.forEach((v, j) => {
      const sp = colSpec[j] || {}; let s = sp.s || ST.cell; let opt = { z: sp.z };
      if (typeof sp.style === 'function') s = sp.style(v, i) || s;
      if (sp.link && v) opt.link = String(v);
      this.set(r0 + 1 + i, c0 + j, v, s, opt);
    }));
    return r0 + rows.length + 1;
  }
  build(widths, { autofilter } = {}) {
    const ws = this.ws;
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(this.R, 0), c: Math.max(this.C, 0) } });
    if (widths) ws['!cols'] = widths.map(w => ({ wch: w }));
    if (this.merges.length) ws['!merges'] = this.merges;
    if (this.rowsH.length) ws['!rows'] = this.rowsH;
    if (autofilter) ws['!autofilter'] = { ref: autofilter };
    return ws;
  }
}

function pivot(rows, keyFn, labFn) {
  const m = new Map();
  for (const r of rows) for (const k of [].concat(keyFn(r))) {
    if (!m.has(k)) m.set(k, { k, n: 0, open: 0, a: 0, rj: 0, f: [0, 0, 0, 0, 0], past: 0, last: null });
    const x = m.get(k); x.n++; if (r.open) x.open++; if (r.result === 'Approved') x.a++; if (r.result === 'Rejected') x.rj++;
    if (r.flagN >= 1) x.f[r.flagN]++; if (r.issues.includes('tlPastOpen')) x.past++; if (r.lastUpdate && (!x.last || r.lastUpdate > x.last)) x.last = r.lastUpdate;
  }
  return [...m.values()].sort((a, b) => b.n - a.n).map(x => [labFn(x.k), x.n, x.open, x.a, x.rj, x.a + x.rj ? x.a / (x.a + x.rj) : '', x.f[1] + x.f[2] + x.f[3] + x.f[4], x.f[1], x.f[2], x.f[3], x.f[4], x.past, x.last]);
}
const PIVOT_H = ['Tổng phiếu', 'Đang xử lý', 'Approved', 'Rejected', 'Tỷ lệ duyệt', 'Có cờ nhắc', 'Mức 1 (>1 tháng)', 'Mức 2 (>3 tháng)', 'Mức 3 (>6 tháng)', 'Mức 4 (>1 năm)', 'Qua timeline, chưa xong', 'Cập nhật gần nhất'];
function pivotSheet(title, firstCol, rows) {
  const b = new SheetB();
  b.set(0, 0, title, ST.title); b.set(1, 0, `Ngày mốc: ${fmtD(S.ref)} · ${rows.length} dòng`, ST.sub);
  const num = { s: ST.center }; const spec = [{ s: ST.bold }, num, num, num, num, { s: ST.center, z: '0.0%' }, num, num, num, num, num, num, { s: ST.center, z: 'dd/mm/yyyy' }];
  spec[6] = { s: ST.center, style: v => (v ? fillStyle(['FDE8E8', 'B3262B']) : null) };
  const end = b.table(3, 0, [firstCol, ...PIVOT_H], rows, spec);
  const tot = rows.reduce((a, r) => { for (let j = 1; j <= 11; j++) if (j !== 5) a[j] = (a[j] || 0) + (r[j] || 0); return a; }, []);
  b.set(end, 0, 'TỔNG', ST.total); for (let j = 1; j <= 12; j++) b.set(end, j, j === 5 ? (tot[3] + tot[4] ? tot[3] / (tot[3] + tot[4]) : '') : j === 12 ? '' : tot[j], { ...ST.total, alignment: { horizontal: 'center' } }, { z: j === 5 ? '0.0%' : undefined });
  return b.build([30, 11, 11, 11, 11, 11, 12, 12, 12, 12, 12, 14, 16], { autofilter: XLSX.utils.encode_range({ s: { r: 3, c: 0 }, e: { r: 3 + rows.length, c: 12 } }) });
}

function exportExcel(mode) {
  const rows = mode === 'all' ? [...S.rows].sort((a, b) => (b.lastUpdate || 0) - (a.lastUpdate || 0)) : S.view;
  if (!rows.length) { toast('Không có phiếu nào để xuất'); return; }
  const m = S.ref.getMonth() + 1, y = S.ref.getFullYear();
  const fname = `Tổng hợp Dealreg Tháng ${m}.${y}.xlsx`;
  const wb = XLSX.utils.book_new();
  const vend = r => r.vendors.map(v => S.vendorLabel.get(v)).join(', ');
  const n = rows.length; const res = countBy(rows, r => r.result);
  const A = res.get('Approved') || 0, R = res.get('Rejected') || 0, O = res.get('Đang xử lý') || 0;
  const flagged = rows.filter(r => r.flagN >= 1);

  // 1. Tổng quan
  {
    const b = new SheetB(); let r = 0;
    b.set(r, 0, `TỔNG HỢP ĐĂNG KÝ DỰ ÁN (DEALREG) — THÁNG ${m}/${y}`, ST.title); b.merge(r, 0, r, 5); r++;
    b.set(r++, 0, 'Công ty Nam Trường Sơn Hà Nội (NTS Hanoi Corp.) — Phòng Quản lý sản phẩm', ST.sub);
    b.set(r++, 0, `Ngày mốc tính: ${fmtD(S.ref)} · Xuất lúc: ${fmtDT(new Date())} · Nguồn: ${S.files.map(f => f.name).join(', ')}`, ST.sub);
    const fl = [];
    if (mode !== 'all') {
      if (S.q) fl.push(`Từ khoá “${S.q}”`);
      if (S.from || S.to) fl.push(`${$('#date-field').selectedOptions[0].text}: ${S.from ? fmtD(S.from) : '…'} → ${S.to ? fmtD(S.to) : '…'}`);
      for (const fk of Object.keys(FACETS)) if (S.f[fk].size) fl.push(`${FACETS[fk].label}: ${[...S.f[fk]].map(v => facetLabel(fk, v)).join(', ')}`);
    }
    b.set(r++, 0, `Phạm vi dữ liệu: ${mode === 'all' ? 'Toàn bộ dữ liệu' : fl.length ? 'Theo bộ lọc — ' + fl.join(' | ') : 'Toàn bộ (không lọc)'}`, { font: F({ bold: true, color: { rgb: '0B5CB5' } }), alignment: { wrapText: true, vertical: 'top' } }); b.merge(r - 1, 0, r - 1, 5); b.rowsH[r - 1] = { hpt: fl.length > 2 ? 45 : 18 };
    r++;
    b.set(r, 0, 'CHỈ TIÊU CHÍNH', ST.sect); for (let c = 1; c <= 2; c++) b.set(r, c, '', ST.sect); b.merge(r, 0, r, 2); r++;
    const appr = rows.filter(x => x.result === 'Approved' && x.procDays != null).map(x => x.procDays);
    const kpis = [
      ['Tổng số phiếu', n, ''], ['Đang xử lý', O, n ? O / n : ''], ['Approved', A, n ? A / n : ''], ['Rejected', R, n ? R / n : ''],
      ['Tỷ lệ duyệt (Approved / đã có kết quả)', A + R ? A / (A + R) : '', ''], ['Số Sale', new Set(rows.map(x => x.sale).filter(Boolean)).size, ''], ['Số Hãng', new Set(rows.flatMap(x => x.vendors)).size, ''],
      ['Số End user', new Set(rows.map(x => x.euKey).filter(Boolean)).size, ''],
      ['Cần nhắc update (> 1 tháng)', flagged.length, n ? flagged.length / n : ''], ['Phiếu đang xử lý quá hạn SLA', rows.filter(x => x.issues.includes('slaOver')).length, ''],
      ['Phiếu đang xử lý đã qua timeline dự án', rows.filter(x => x.issues.includes('tlPastOpen')).length, ''], ['Phiếu trùng End user + Hãng', rows.filter(x => x.issues.includes('dup')).length, ''],
      ['Thời gian từ tạo → Approved (trung vị, ngày)', appr.length ? Math.round(median(appr) * 10) / 10 : '', ''],
    ];
    kpis.forEach(([k, v, p]) => { b.set(r, 0, k, ST.kl); b.set(r, 1, v, ST.kv, { z: /Tỷ lệ/.test(k) ? '0.0%' : undefined }); b.set(r, 2, p, { ...ST.cell, font: F({ color: { rgb: '62727F' } }), alignment: { horizontal: 'right' } }, { z: '0.0%' }); r++; });
    r++;
    const block = (title, head, data, spec) => { b.set(r, 0, title, ST.sect); for (let c = 1; c < head.length; c++) b.set(r, c, '', ST.sect); b.merge(r, 0, r, head.length - 1); r++; r = b.table(r, 0, head, data, spec) + 1; };
    const cb = countBy(rows, x => x.block || '(Trống)');
    block('THEO TIẾN ĐỘ XỬ LÝ (TÊN KHỐI)', ['Tiến độ', 'Số phiếu', 'Tỷ lệ'], blockOrder().concat(cb.has('(Trống)') ? ['(Trống)'] : []).filter(k => cb.get(k)).map(k => [k, cb.get(k), cb.get(k) / n]), [{ s: ST.bold }, { s: ST.center }, { s: ST.center, z: '0.0%' }]);
    const cf = countBy(rows, x => x.flag);
    block('THEO CỜ NHẮC CẬP NHẬT', ['Cờ nhắc', 'Số phiếu', 'Tỷ lệ'], ['4', '3', '2', '1', '0', 'na'].filter(k => cf.get(k)).map(k => [FLAGS[k].t, cf.get(k), cf.get(k) / n]), [{ style: (v, i) => fillStyle(FLAG_FILL[['4', '3', '2', '1', '0', 'na'].filter(k => cf.get(k))[i]]) }, { s: ST.center }, { s: ST.center, z: '0.0%' }]);
    const ctl = countBy(rows, x => x.tl);
    block('THEO TIMELINE DỰ ÁN', ['Tình trạng', 'Số phiếu', 'Tỷ lệ'], Object.keys(TLS).filter(k => ctl.get(k)).map(k => [TLS[k].t, ctl.get(k), ctl.get(k) / n]), [{ s: ST.bold }, { s: ST.center }, { s: ST.center, z: '0.0%' }]);
    const top = (fn, lab, k) => topN(countBy(rows, fn), k).map(([x, c]) => [lab(x), c, c / n]);
    block('TOP 10 HÃNG', ['Hãng', 'Số phiếu', 'Tỷ lệ'], top(x => x.vendors, x => S.vendorLabel.get(x) || x, 10), [{ s: ST.bold }, { s: ST.center }, { s: ST.center, z: '0.0%' }]);
    block('TOP 10 SALE', ['Sale', 'Số phiếu', 'Tỷ lệ'], top(x => x.sale || '__none', personLabel, 10), [{ s: ST.bold }, { s: ST.center }, { s: ST.center, z: '0.0%' }]);
    b.set(r + 1, 0, 'Dữ liệu được tổng hợp nội bộ từ phòng Quản lý sản phẩm - Công ty Nam Trường Sơn Hà Nội', ST.foot);
    XLSX.utils.book_append_sheet(wb, b.build([46, 16, 12, 12, 12, 12]), 'Tổng quan');
  }
  // 2. Danh sách chi tiết
  {
    const H = ['STT', 'ID', 'Cờ nhắc', 'Số ngày chưa cập nhật', 'Tên phiếu', 'Hãng', 'End user', 'Tên dự án', 'Reseller / Partner', 'Thông tin BOM', 'Timeline dự án', 'Quý timeline', 'Tình trạng timeline',
      'Sale', 'Người tạo', 'PM phụ trách', 'Tiến độ (Tên khối)', 'Trạng thái', 'Kết quả DR', 'Ngày tạo', 'Cập nhật cuối', 'Tuổi phiếu (ngày)', 'Hạn SLA', 'Quá hạn SLA',
      'Cảnh báo dữ liệu', 'Note', 'Link Deal CRM', 'Email End user', 'Website khách hàng', 'Thông tin khách hàng', 'Người theo dõi', 'Mô tả phiếu', 'Tệp đính kèm'];
    const data = rows.map((r, i) => [i + 1, +r.id || r.id, FLAGS[r.flag].t, r.days, r.name, vend(r), r.eu, r.project, r.resellerRaw, r.bom, r.timeline || r.timelineRaw, r.timeline ? quarterLabel(quarterKey(r.timeline)) : '', TLS[r.tl].t,
      personLabel(r.sale || '__none'), personLabel(r.creator), persons(r.pms), r.block, r.status, r.result, r.created, r.lastUpdate, r.age, r.sla, r.overdue,
      r.issues.map(k => ISSUES[k]).join('; '), r.note, r.crm, r.euEmail, r.website, r.custInfo, persons(r.followers), r.desc, r.attachments.map(a => a.name).join('; ')]);
    const b = new SheetB();
    const C = ST.center, W = ST.wrap;
    const spec = [{ s: C }, { s: ST.bold }, { style: (v, i) => fillStyle(FLAG_FILL[rows[i].flag]) }, { s: C }, { s: { ...W, font: F({ bold: true }) } }, { s: ST.bold }, { s: W }, { s: W }, { s: W }, { s: W }, { s: C, z: 'dd/mm/yyyy' }, { s: C }, { style: (v, i) => ({ ...ST.center, font: F({ color: { rgb: TLS[rows[i].tl].c.slice(1).toUpperCase() } }) }) },
      {}, {}, {}, {}, {}, { style: (v) => fillStyle(RES_FILL[v] || RES_FILL['Hoàn thành khác']) }, { s: C, z: 'dd/mm/yyyy hh:mm' }, { s: C, z: 'dd/mm/yyyy hh:mm' }, { s: C }, { s: C, z: 'dd/mm/yyyy hh:mm' }, { s: C },
      { s: { ...W, font: F({ color: { rgb: 'B3262B' } }) } }, { s: W }, { s: ST.link, link: true }, {}, {}, { s: W }, { s: W }, { s: W }, { s: W }];
    b.table(0, 0, H, data, spec);
    b.rowsH[0] = { hpt: 32 };
    XLSX.utils.book_append_sheet(wb, b.build([5, 9, 22, 10, 46, 14, 32, 30, 18, 44, 12, 9, 17, 22, 20, 20, 22, 12, 13, 16, 16, 9, 16, 9, 30, 32, 34, 26, 24, 30, 30, 60, 30], { autofilter: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: data.length, c: H.length - 1 } }) }), 'Danh sách Dealreg');
  }
  // 3. Cần nhắc update
  {
    const fr = flagged.slice().sort((a, b) => b.flagN - a.flagN || b.days - a.days);
    const H = ['STT', 'Cờ nhắc', 'Số ngày chưa cập nhật', 'Cập nhật cuối', 'ID', 'Tên phiếu', 'Hãng', 'End user', 'Sale', 'PM phụ trách', 'Tiến độ (Tên khối)', 'Trạng thái', 'Timeline dự án', 'Link Deal CRM'];
    const data = fr.map((r, i) => [i + 1, FLAGS[r.flag].t, r.days, r.lastUpdate, +r.id || r.id, r.name, vend(r), r.eu, personLabel(r.sale || '__none'), persons(r.pms), r.block, r.status, r.timeline || r.timelineRaw, r.crm]);
    const b = new SheetB();
    b.table(0, 0, H, data, [{ s: ST.center }, { style: (v, i) => fillStyle(FLAG_FILL[fr[i].flag]) }, { s: ST.center }, { s: ST.center, z: 'dd/mm/yyyy' }, { s: ST.bold }, { s: ST.wrap }, { s: ST.bold }, { s: ST.wrap }, {}, {}, {}, {}, { s: ST.center, z: 'dd/mm/yyyy' }, { s: ST.link, link: true }]);
    b.rowsH[0] = { hpt: 30 };
    XLSX.utils.book_append_sheet(wb, b.build([5, 20, 11, 13, 9, 46, 14, 32, 22, 20, 22, 12, 13, 34], { autofilter: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(1, data.length), c: H.length - 1 } }) }), 'Cần nhắc update');
  }
  // 4-6. Pivot
  XLSX.utils.book_append_sheet(wb, pivotSheet('TỔNG HỢP THEO SALE', 'Sale', pivot(rows, r => r.sale || '__none', personLabel)), 'Theo Sale');
  XLSX.utils.book_append_sheet(wb, pivotSheet('TỔNG HỢP THEO HÃNG', 'Hãng', pivot(rows, r => (r.vendors.length ? r.vendors : ['__none']), k => (k === '__none' ? '(Trống)' : S.vendorLabel.get(k)))), 'Theo Hãng');
  XLSX.utils.book_append_sheet(wb, pivotSheet('TỔNG HỢP THEO PM PHỤ TRÁCH', 'PM phụ trách', pivot(rows, r => (r.pms.length ? r.pms : ['__none']), personLabel)), 'Theo PM');
  // 7. Theo tháng
  {
    const keys = [...new Set(rows.filter(r => r.created).map(r => monthKey(r.created)))].sort();
    const data = keys.map(k => { const rs = rows.filter(r => r.created && monthKey(r.created) === k); const a = rs.filter(r => r.result === 'Approved').length, rj = rs.filter(r => r.result === 'Rejected').length; return [monthLabel(k), rs.length, rs.filter(r => r.open).length, a, rj, a + rj ? a / (a + rj) : '', new Set(rs.map(r => r.sale).filter(Boolean)).size, new Set(rs.flatMap(r => r.vendors)).size]; });
    const b = new SheetB(); b.set(0, 0, 'XU HƯỚNG THEO THÁNG TẠO PHIẾU', ST.title);
    const C = { s: ST.center };
    b.table(2, 0, ['Tháng', 'Tạo mới', 'Đang xử lý', 'Approved', 'Rejected', 'Tỷ lệ duyệt', 'Số Sale', 'Số Hãng'], data, [{ s: ST.bold }, C, C, C, C, { s: ST.center, z: '0.0%' }, C, C]);
    // Pipeline theo quý timeline
    const P = rows.filter(r => r.timeline && r.result !== 'Rejected'); const qk = [...new Set(P.map(r => quarterKey(r.timeline)))].sort();
    const r0 = data.length + 5; b.set(r0 - 1, 0, 'PIPELINE THEO QUÝ TIMELINE DỰ ÁN (không tính Rejected)', ST.title);
    b.table(r0, 0, ['Quý', 'Số phiếu', 'Approved', 'Đang xử lý', 'Số Hãng', 'Số End user'], qk.map(q => { const rs = P.filter(r => quarterKey(r.timeline) === q); return [quarterLabel(q), rs.length, rs.filter(r => r.result === 'Approved').length, rs.filter(r => r.open).length, new Set(rs.flatMap(r => r.vendors)).size, new Set(rs.map(r => r.euKey).filter(Boolean)).size]; }), [{ s: ST.bold }, C, C, C, C, C]);
    XLSX.utils.book_append_sheet(wb, b.build([16, 12, 12, 12, 12, 12, 10, 10]), 'Theo tháng & quý');
  }
  // 8. Kiểm tra dữ liệu
  {
    const iss = rows.filter(r => r.issues.length).sort((a, b) => b.issues.includes('conflict') - a.issues.includes('conflict') || (a.euKey > b.euKey ? 1 : -1));
    const H = ['STT', 'ID', 'Cảnh báo', 'Tên phiếu', 'Hãng (cột)', 'Hãng ở tên phiếu', 'End user', 'Sale', 'PM phụ trách', 'Kết quả DR', 'Phiếu trùng EU + Hãng', 'Timeline dự án', 'Ngày tạo'];
    const data = iss.map((r, i) => [i + 1, +r.id || r.id, r.issues.map(k => ISSUES[k]).join('; '), r.name, vend(r), r.nameVendor || '', r.eu, personLabel(r.sale || '__none'), persons(r.pms), r.result, r.dupPeers.map(x => `#${x.id} (${personLabel(x.sale)})`).join(', '), r.timeline || r.timelineRaw, r.created]);
    const b = new SheetB();
    b.table(0, 0, H, data, [{ s: ST.center }, { s: ST.bold }, { s: { ...ST.wrap, font: F({ bold: true, color: { rgb: 'B3262B' } }) } }, { s: ST.wrap }, {}, {}, { s: ST.wrap }, {}, {}, { style: v => fillStyle(RES_FILL[v] || RES_FILL['Hoàn thành khác']) }, { s: ST.wrap }, { s: ST.center, z: 'dd/mm/yyyy' }, { s: ST.center, z: 'dd/mm/yyyy' }]);
    XLSX.utils.book_append_sheet(wb, b.build([5, 9, 34, 44, 14, 14, 30, 22, 20, 13, 30, 13, 12], { autofilter: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(1, data.length), c: H.length - 1 } }) }), 'Kiểm tra dữ liệu');
  }
  wb.Props = { Title: fname.replace('.xlsx', ''), Author: 'Phòng Quản lý sản phẩm - NTS Hanoi Corp.', Company: 'Công ty Nam Trường Sơn Hà Nội' };
  let out = XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: true });
  try {
    out = patchViews(out, [{ grid: false }, { r: 1, c: 5 }, { r: 1, c: 6 }, { r: 4, c: 1 }, { r: 4, c: 1 }, { r: 4, c: 1 }, { r: 3, c: 1 }, { r: 1, c: 4 }]);
  } catch (e) { console.warn('Không cố định được dòng tiêu đề', e); }
  download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fname);
  toast(`Đã xuất “${fname}” — ${fmtN(rows.length)} phiếu`);
}
// Cố định dòng/cột tiêu đề (freeze panes) & ẩn lưới — chèn trực tiếp vào XML của sheet
function patchViews(buf, views) {
  const cfb = XLSX.CFB.read(new Uint8Array(buf), { type: 'array' });
  views.forEach((v, i) => {
    const idx = cfb.FullPaths.findIndex(p => p.endsWith(`xl/worksheets/sheet${i + 1}.xml`));
    if (idx < 0 || !v) return;
    const e = cfb.FileIndex[idx]; let x = new TextDecoder().decode(e.content);
    let inner = '';
    if (v.r || v.c) {
      const tl = XLSX.utils.encode_cell({ r: v.r || 0, c: v.c || 0 });
      const pane = v.r && v.c ? 'bottomRight' : v.r ? 'bottomLeft' : 'topRight';
      inner = `<pane ${v.c ? `xSplit="${v.c}" ` : ''}${v.r ? `ySplit="${v.r}" ` : ''}topLeftCell="${tl}" activePane="${pane}" state="frozen"/><selection pane="${pane}" activeCell="${tl}" sqref="${tl}"/>`;
    }
    const open = `<sheetView${v.grid === false ? ' showGridLines="0"' : ''} workbookViewId="0"`;
    x = x.replace(/<sheetView workbookViewId="0"\/>/, inner ? `${open}>${inner}</sheetView>` : `${open}/>`);
    e.content = new TextEncoder().encode(x); e.size = e.content.length;
  });
  return XLSX.CFB.write(cfb, { fileType: 'zip', type: 'array', compression: true });
}

/* ---------------------------------------------------------------- events */
function bind() {
  const fi = $('#file-input');
  const pick = () => { fi.value = ''; fi.click(); };
  fi.onchange = () => { if (fi.files.length) loadFiles(fi.files); };
  $('#btn-open').onclick = pick; $('#drop').onclick = pick;
  // kéo-thả toàn trang
  let dragN = 0;
  window.addEventListener('dragenter', e => { e.preventDefault(); dragN++; $('#drop').classList.add('drag'); });
  window.addEventListener('dragleave', () => { if (--dragN <= 0) { dragN = 0; $('#drop').classList.remove('drag'); } });
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('drop', e => { e.preventDefault(); dragN = 0; $('#drop').classList.remove('drag'); if (e.dataTransfer.files.length) loadFiles(e.dataTransfer.files); });
  $('#filebar').addEventListener('click', e => {
    if (e.target.id === 'btn-add') pick();
    if (e.target.id === 'btn-clear' && confirm('Xoá toàn bộ dữ liệu đã nạp?')) { S.files = []; S.byId = new Map(); resetFiltersSilent(); rebuild(); $('#load-warn').innerHTML = ''; }
  });
  // export
  $('#btn-export').onclick = () => exportExcel('filtered');
  $('#sticky-export').onclick = () => exportExcel('filtered');
  $('#btn-export-more').onclick = e => { e.stopPropagation(); $('#export-menu').classList.toggle('hidden'); };
  $('#export-menu').onclick = e => { const b = e.target.closest('[data-exp]'); if (b) { $('#export-menu').classList.add('hidden'); exportExcel(b.dataset.exp); } };
  document.addEventListener('click', e => { if (!e.target.closest('#export-split')) $('#export-menu').classList.add('hidden'); if (!e.target.closest('.colpick')) $('#cols-menu').classList.add('hidden'); });
  // facets
  $$('[data-facet]').forEach(h => { MS[h.dataset.facet] = new MultiSelect(h, h.dataset.facet); });
  // search
  const q = $('#q');
  const doSearch = debounce(() => { S.q = q.value.trim(); S.page = 1; update(); renderSuggest(); }, 160);
  q.addEventListener('input', () => { $('#searchbox').classList.toggle('has-val', !!q.value); doSearch(); });
  q.addEventListener('focus', renderSuggest);
  q.addEventListener('blur', () => setTimeout(() => $('#suggest').classList.add('hidden'), 150));
  q.addEventListener('keydown', e => {
    const its = $$('#suggest .it');
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); if (!its.length) return; S._sugAct = (S._sugAct + (e.key === 'ArrowDown' ? 1 : -1) + its.length) % its.length; its.forEach((x, i) => x.classList.toggle('act', i === S._sugAct)); its[S._sugAct].scrollIntoView({ block: 'nearest' }); }
    else if (e.key === 'Enter') { e.preventDefault(); if (S._sugAct >= 0) pickSuggest(+its[S._sugAct].dataset.i); else { $('#suggest').classList.add('hidden'); S.q = q.value.trim(); update(); $('#results').scrollIntoView({ behavior: 'smooth' }); } }
    else if (e.key === 'Escape') { $('#suggest').classList.add('hidden'); }
  });
  $('#suggest').addEventListener('mousedown', e => { e.preventDefault(); const it = e.target.closest('.it'); if (it) pickSuggest(+it.dataset.i); });
  $('#q-clear').onclick = () => { q.value = ''; S.q = ''; $('#searchbox').classList.remove('has-val'); update(); q.focus(); };
  // date
  const dchg = () => { S.dateField = $('#date-field').value; S.from = $('#date-from').value ? new Date($('#date-from').value + 'T00:00:00') : null; S.to = $('#date-to').value ? new Date($('#date-to').value + 'T00:00:00') : null; S.page = 1; update(); };
  ['#date-field', '#date-from', '#date-to'].forEach(s => $(s).addEventListener('change', () => { if (s !== '#date-field') $('#date-preset').value = ''; dchg(); }));
  $('#date-preset').onchange = () => {
    const p = $('#date-preset').value; if (!p) return; const t = S.ref; const Y = t.getFullYear(), M = t.getMonth(), Q = Math.floor(M / 3);
    const R = { '7d': [new Date(+t - 6 * DAY), t], '30d': [new Date(+t - 29 * DAY), t], thisM: [new Date(Y, M, 1), new Date(Y, M + 1, 0)], lastM: [new Date(Y, M - 1, 1), new Date(Y, M, 0)],
      thisQ: [new Date(Y, Q * 3, 1), new Date(Y, Q * 3 + 3, 0)], lastQ: [new Date(Y, Q * 3 - 3, 1), new Date(Y, Q * 3, 0)], nextQ: [new Date(Y, Q * 3 + 3, 1), new Date(Y, Q * 3 + 6, 0)],
      thisY: [new Date(Y, 0, 1), new Date(Y, 11, 31)], nextY: [new Date(Y + 1, 0, 1), new Date(Y + 1, 11, 31)] }[p];
    $('#date-from').value = isoD(R[0]); $('#date-to').value = isoD(R[1]); dchg();
  };
  // quick, active chips, legends, kpi, alert
  $('#quick').addEventListener('click', e => { const b = e.target.closest('[data-i]'); if (!b) return; const qd = QUICK[+b.dataset.i]; const on = quickOn(qd); Object.entries(qd.set).forEach(([k, vs]) => setFacet(k, on ? [] : vs)); S.page = 1; update(); });
  document.addEventListener('click', e => {
    const c = e.target.closest('[data-clr]');
    if (c) { const k = c.dataset.clr; if (k === 'all') return resetFilters(); if (k === 'q') { S.q = ''; $('#q').value = ''; $('#searchbox').classList.remove('has-val'); } if (k === 'date') { S.from = S.to = null; $('#date-from').value = $('#date-to').value = ''; $('#date-preset').value = ''; } S.page = 1; return update(); }
    const x = e.target.closest('#activebar [data-fk]'); if (x) return toggleFacet(x.dataset.fk, x.dataset.v);
    const lg = e.target.closest('.legend [data-fk]'); if (lg) return toggleFacet(lg.dataset.fk, lg.dataset.v);
    const kp = e.target.closest('.kpi.click'); if (kp) { const K = S._kpi[+kp.dataset.k]; const on = Object.entries(K.q).every(([k, vs]) => S.f[k].size === vs.length && vs.every(v => S.f[k].has(v))); Object.entries(K.q).forEach(([k, vs]) => setFacet(k, on ? [] : vs)); S.page = 1; return update(); }
    const af = e.target.closest('[data-flag]'); if (af) return toggleFacet('flag', af.dataset.flag, true);
    if (e.target.closest('#a-list')) { setFacet('flag', ['1', '2', '3', '4']); S.sort = { k: 'flag', d: -1 }; S.page = 1; update(); $('#results').scrollIntoView({ behavior: 'smooth' }); return; }
    if (e.target.closest('#a-remind')) return openReminder();
  });
  $('#btn-reset').onclick = resetFilters;
  $('#adv-toggle').onclick = () => { $('#adv').classList.toggle('hidden'); $('#adv-toggle').classList.toggle('open'); };
  $('#ref-date').value = isoD(S.ref);
  $('#ref-date').onchange = () => { const v = $('#ref-date').value; S.ref = v ? new Date(v + 'T00:00:00') : sod(new Date()); if (!v) $('#ref-date').value = isoD(S.ref); derive(); renderHeroSub(); update(); };
  $('#flag-scope').onchange = () => { S.scope = $('#flag-scope').value; derive(); update(); };
  // trend granularity
  $('#trend-gran').onclick = e => { const b = e.target.closest('[data-g]'); if (!b) return; S.trendGran = b.dataset.g; $$('#trend-gran button').forEach(x => x.classList.toggle('on', x === b)); renderCharts(); };
  // table
  $('#grid thead').addEventListener('click', e => { const th = e.target.closest('th'); if (!th) return; const k = th.dataset.k; S.sort = S.sort.k === k ? { k, d: -S.sort.d } : { k, d: ['name', 'eu', 'vendor', 'sale', 'pm', 'creator', 'block', 'status', 'reseller', 'project', 'note'].includes(k) ? 1 : -1 }; update(); });
  $('#grid tbody').addEventListener('click', e => { const tr = e.target.closest('tr[data-id]'); if (tr) openDetail(tr.dataset.id); });
  $('#pager').addEventListener('click', e => { const b = e.target.closest('[data-p]'); if (!b || b.disabled) return; S.page = +b.dataset.p; renderTable(); $('#results').scrollIntoView({ behavior: 'smooth' }); });
  $('#page-size').onchange = () => { S.pageSize = +$('#page-size').value; S.page = 1; renderTable(); };
  $('#btn-cols').onclick = e => { e.stopPropagation(); $('#cols-menu').classList.toggle('hidden'); };
  $('#cols-menu').addEventListener('change', e => { const k = e.target.dataset.col; if (e.target.checked) S.hiddenCols.delete(k); else S.hiddenCols.add(k); renderTable(); });
  // sticky bar
  const sb = $('#stickybar');
  window.addEventListener('scroll', () => { const f = $('#filters'); sb.classList.toggle('show', S.rows.length > 0 && f.getBoundingClientRect().bottom < 0); }, { passive: true });
  $('#sticky-filter').onclick = () => $('#filters').scrollIntoView({ behavior: 'smooth' });
}
function resetFiltersSilent() { Object.keys(FACETS).forEach(k => S.f[k].clear()); S.q = ''; $('#q').value = ''; S.from = S.to = null; }

/* ------------------------------------------------------------------ init */
Object.defineProperty(S, 'byIdRow', { get() { if (this._rowsRef !== this.rows) { this._rowsRef = this.rows; this._byIdRow = new Map(this.rows.map(r => [r.id, r])); } return this._byIdRow; } });
function init() {
  Chart.register(ChartDataLabels);
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
  Chart.defaults.font.size = 12;
  Chart.defaults.color = '#51606c';
  Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(15,42,51,.94)';
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.cornerRadius = 8;
  Chart.defaults.plugins.datalabels = Object.assign(Chart.defaults.plugins.datalabels || {}, { display: false });
  bind();
}
init();
window.DealregTool = { S, loadFiles, exportExcel }; // hỗ trợ kiểm thử
})();
