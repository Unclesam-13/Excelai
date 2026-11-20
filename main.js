'use strict';
/* AI 表格助手 · 表格工作台 —— Obsidian 插件（无需构建，直接使用本文件） */
const {
  Plugin, ItemView, Notice, Modal, FuzzySuggestModal, PluginSettingTab,
  Setting, setIcon, normalizePath, TFile,
} = require('obsidian');

const VIEW_TYPE = 'ai-table-assistant-view';
const MAX_FILE = 200 * 1024 * 1024; // 导入 / 附件大小上限 200MB

/* 表格文件夹里由插件生成、给 AI 看的说明文件（不是表格） */
const GUIDE_FILES = ['AI使用说明.md', '表格目录.md', 'CLAUDE.md', 'AGENTS.md'];
/* 旧版本内置 AI 用的设置，已经不用了 */
const OBSOLETE_SETTINGS = ['provider', 'apiKey', 'baseUrl', 'model', 'maxRows', 'preview', 'autoArchive'];

const DEFAULT_SETTINGS = {
  folder: 'AI表格助手',
  lastTable: '',
  master: '',
  pinned: [],
  recent: [],
  archiveMeta: {},
  openAfterExport: true,
  meta: {},
};

/* ====================================================================
 * 一、表格工具（纯函数）
 * 表格模型：{ headers: string[], rows: string[][] }
 * ==================================================================== */
const cellStr = (c) => (c == null ? '' : String(c));
const isBlank = (c) => c == null || String(c).trim() === '';

function normalizeTable(t) {
  const headers = (t.headers || []).map((h) => cellStr(h).trim());
  let width = headers.length;
  for (const r of t.rows || []) if (r.length > width) width = r.length;
  while (headers.length < width) headers.push('列' + (headers.length + 1));
  const rows = (t.rows || []).map((r) => {
    const out = r.map(cellStr);
    while (out.length < width) out.push('');
    return out;
  });
  return { headers, rows };
}

/**
 * 把二维数组整理成表格：去掉空行、去掉没有任何内容的列（比如 Excel 里格式延伸到 XFD 列的那种），
 * detectHeader 时自动跳过表头上方的标题行（返回在 title 里）。
 */
function cleanGrid(grid, opts) {
  opts = opts || {};
  const ids0 = opts.rowIds || (grid || []).map((_, i) => i);
  const ids = [];
  grid = (grid || []).filter((r, i) => { const ok = r && r.some((c) => !isBlank(c)); if (ok) ids.push(ids0[i]); return ok; });
  if (!grid.length) return { table: { headers: [], rows: [] }, title: '', src: null };

  let headerIdx = 0;
  if (opts.detectHeader) {
    const counts = grid.slice(0, 30).map((r) => r.reduce((n, c) => n + (isBlank(c) || /^列\d+$/.test(String(c).trim()) ? 0 : 1), 0));
    const max = Math.max(...counts);
    if (max >= 2) {
      const idx = counts.findIndex((n) => n >= Math.max(2, Math.ceil(max * 0.6)));
      if (idx > 0 && idx < 15) headerIdx = idx;
    }
  }
  const title = grid.slice(0, headerIdx)
    .map((r) => r.filter((c) => !isBlank(c) && !/^列\d+$/.test(String(c).trim())).map((c) => String(c).trim()).join(' '))
    .filter(Boolean).join(' / ');

  const headerRow = (grid[headerIdx] || []).map((c) => cellStr(c).trim());
  const rows = grid.slice(headerIdx + 1);
  let width = headerRow.length;
  for (const r of rows) if (r.length > width) width = r.length;

  const keep = [];
  for (let c = 0; c < width; c++) {
    const h = headerRow[c] || '';
    let used = h !== '' && !/^列\d+$/.test(h);
    if (!used) for (const r of rows) { if (!isBlank(r[c])) { used = true; break; } }
    if (used) keep.push(c);
  }
  const seen = new Map();
  const headers = keep.map((c, i) => {
    let h = headerRow[c] || '列' + (i + 1);
    const n = seen.get(h) || 0;
    seen.set(h, n + 1);
    if (n) h = `${h}_${n + 1}`;
    return h;
  });
  const out = new Array(rows.length);
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i], line = new Array(keep.length);
    for (let k = 0; k < keep.length; k++) line[k] = cellStr(r[keep[k]]);
    out[i] = line;
  }
  return { table: { headers, rows: out }, title, src: { headerRow: ids[headerIdx], rows: ids.slice(headerIdx + 1), cols: keep } };
}

function gridToTable(grid) {
  return cleanGrid(grid).table;
}

function cloneTable(t) {
  const rows = t.rows.map((r) => {
    const c = r.slice();
    if (r.__src !== undefined) c.__src = r.__src;
    return c;
  });
  const o = { headers: t.headers.slice(), rows };
  if (t.colSrc) o.colSrc = t.colSrc.slice();
  return o;
}

/* ---------- 文本编码（Excel 另存的中文 CSV 常是 GBK） ---------- */
function decodeText(buf) {
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  if (u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) return new TextDecoder('utf-8').decode(u8.subarray(3));
  if (u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder('utf-16le').decode(u8.subarray(2));
  if (u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder('utf-16be').decode(u8.subarray(2));
  try { return new TextDecoder('utf-8', { fatal: true }).decode(u8); } catch (e) { /* 不是 UTF-8 */ }
  try { return new TextDecoder('gb18030').decode(u8); } catch (e) { return new TextDecoder('utf-8').decode(u8); }
}

/* ---------- CSV（按块切片，比逐字拼接快很多） ---------- */
function parseCSV(text) {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const nl = text.indexOf('\n');
  const first = nl < 0 ? text : text.slice(0, nl);
  const count = (ch) => first.split(ch).length - 1;
  let D = 44; // ,
  if (count('\t') > count(',')) D = 9;
  else if (count(';') > count(',')) D = 59;
  const n = text.length;
  const rows = [];
  let row = [];
  let i = 0;
  const scan = (k) => {
    while (k < n) {
      const ch = text.charCodeAt(k);
      if (ch === D || ch === 10 || ch === 13) break;
      k++;
    }
    return k;
  };
  while (i < n) {
    let val;
    if (text.charCodeAt(i) === 34) {
      let j = i + 1;
      val = '';
      for (;;) {
        const q = text.indexOf('"', j);
        if (q < 0) { val += text.slice(j); i = n; break; }
        if (text.charCodeAt(q + 1) === 34) { val += text.slice(j, q + 1); j = q + 2; }
        else { val += text.slice(j, q); i = q + 1; break; }
      }
      const k = scan(i);
      if (k > i) val += text.slice(i, k);
      i = k;
    } else {
      const k = scan(i);
      val = text.slice(i, k);
      i = k;
    }
    row.push(val);
    if (i >= n) break;
    const ch = text.charCodeAt(i);
    if (ch === D) {
      i++;
      if (i >= n) row.push('');
      continue;
    }
    i += ch === 13 && text.charCodeAt(i + 1) === 10 ? 2 : 1;
    rows.push(row);
    row = [];
  }
  if (row.length) rows.push(row);
  return rows.filter((r) => r.some((x) => x.trim() !== ''));
}

function csvCell(v) {
  v = cellStr(v);
  return /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
}

function toCSV(t) {
  const out = new Array(t.rows.length + 1);
  out[0] = t.headers.map(csvCell).join(',');
  for (let i = 0; i < t.rows.length; i++) out[i + 1] = t.rows[i].map(csvCell).join(',');
  return out.join('\n') + '\n';
}

/* ---------- Markdown 表格 ---------- */
function splitMdRow(line) {
  let s = line.trim();
  if (s.startsWith('|')) s = s.slice(1);
  if (s.endsWith('|') && !s.endsWith('\\|')) s = s.slice(0, -1);
  const cells = [];
  let cur = '';
  for (let i = 0; i < s.length; i++) {
    if (s[i] === '\\' && s[i + 1] === '|') { cur += '|'; i++; }
    else if (s[i] === '|') { cells.push(cur); cur = ''; }
    else cur += s[i];
  }
  cells.push(cur);
  return cells.map((c) => c.trim().replace(/<br\s*\/?>/gi, '\n'));
}

/** 找到文本里的第一个 Markdown 表格，返回 {start, end, table}（行号，end 不含）或 null */
function findMdTable(text) {
  const lines = text.split('\n');
  const isRow = (l) => /^\s*\|.*\|\s*$/.test(l || '');
  const isSep = (l) => /^\s*\|?\s*:?-{1,}:?\s*(\|\s*:?-{1,}:?\s*)*\|?\s*$/.test(l || '');
  for (let i = 0; i < lines.length - 1; i++) {
    if (isRow(lines[i]) && isSep(lines[i + 1])) {
      let j = i + 2;
      while (j < lines.length && isRow(lines[j])) j++;
      const headers = splitMdRow(lines[i]);
      const rows = lines.slice(i + 2, j).map(splitMdRow);
      return { start: i, end: j, table: normalizeTable({ headers, rows }) };
    }
  }
  return null;
}

function mdCell(v) {
  return cellStr(v).replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>');
}

function toMarkdown(t) {
  const line = (r) => '| ' + r.map(mdCell).join(' | ') + ' |';
  return [line(t.headers), '| ' + t.headers.map(() => '---').join(' | ') + ' |', ...t.rows.map(line)].join('\n');
}

/* ---------- XLSX 读取（流式解压 + 逐行解析，200MB 级文件也不会一次性生成超大字符串） ---------- */
const xmlUnescape = (s) =>
  s.indexOf('&') < 0 ? s :
    s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
      .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
      .replace(/&amp;/g, '&');
const xmlEscape = (s) =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');

async function unzip(arrayBuffer) {
  const u8 = new Uint8Array(arrayBuffer);
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let e = u8.length - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error('不是有效的 xlsx 文件（如果是旧版 .xls 或加密文件，请先在 Excel 里另存为 .xlsx）');
  const n = dv.getUint16(e + 10, true);
  let p = dv.getUint32(e + 16, true);
  const dec = new TextDecoder();
  const entries = {};
  for (let i = 0; i < n; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true);
    const crc = dv.getUint32(p + 16, true);
    const usize = dv.getUint32(p + 24, true);
    const csize = dv.getUint32(p + 20, true);
    const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    const lho = dv.getUint32(p + 42, true);
    entries[dec.decode(u8.subarray(p + 46, p + 46 + nl))] = { method, csize, usize, crc, lho };
    p += 46 + nl + el + cl;
  }
  const dataOf = (en) => {
    const start = en.lho + 30 + dv.getUint16(en.lho + 26, true) + dv.getUint16(en.lho + 28, true);
    return u8.subarray(start, start + en.csize);
  };
  return {
    names: Object.keys(entries),
    has: (name) => !!entries[name],
    /** 原样取出压缩数据（重新打包时不用解压） */
    raw: (name) => { const en = entries[name]; return en && { method: en.method, crc: en.crc, usize: en.usize, data: dataOf(en) }; },
    /** 逐块产出解压后的文本 */
    async *chunks(name) {
      const en = entries[name];
      if (!en) return;
      const data = dataOf(en);
      const td = new TextDecoder();
      if (en.method === 0) {
        for (let o = 0; o < data.length; o += 4 << 20) yield td.decode(data.subarray(o, o + (4 << 20)), { stream: true });
      } else {
        const reader = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          yield td.decode(value, { stream: true });
        }
      }
      const tail = td.decode();
      if (tail) yield tail;
    },
    async text(name) {
      let s = '';
      for await (const c of this.chunks(name)) s += c;
      return s;
    },
  };
}

/** 按闭合标签切块处理大 XML：cb 收到的片段里只包含完整的 <tag>…</tag> */
async function eachXmlBlock(zip, name, tag, cb) {
  const close = '</' + tag + '>';
  let buf = '';
  for await (const chunk of zip.chunks(name)) {
    buf += chunk;
    const end = buf.lastIndexOf(close);
    if (end < 0) continue;
    const cut = end + close.length;
    await cb(buf.slice(0, cut));
    buf = buf.slice(cut);
  }
  if (buf) await cb(buf);
}

/* Excel 日期格式 */
const BUILTIN_DATE = new Set([14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 57, 58]);
const BUILTIN_TIME = new Set([18, 19, 20, 21, 32, 33, 45, 46, 47, 55, 56]);
function fmtKind(id, code) {
  if (BUILTIN_DATE.has(id)) return id === 22 ? 'datetime' : 'date';
  if (BUILTIN_TIME.has(id)) return 'time';
  if (!code) return '';
  const c = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '').toLowerCase();
  if (!/[ymdhs]/.test(c) || /^[#0.,%\s-]*$/.test(c)) return '';
  const hasDate = /[yd]/.test(c) || (/m/.test(c) && !/[hs]/.test(c));
  const hasTime = /[hs]/.test(c);
  return hasDate && hasTime ? 'datetime' : hasDate ? 'date' : hasTime ? 'time' : '';
}
function excelDate(serial, kind) {
  const ms = Math.round((serial - 25569) * 86400000);
  const d = new Date(ms);
  if (isNaN(d)) return String(serial);
  const z = (x) => String(x).padStart(2, '0');
  const date = `${d.getUTCFullYear()}-${z(d.getUTCMonth() + 1)}-${z(d.getUTCDate())}`;
  const time = `${z(d.getUTCHours())}:${z(d.getUTCMinutes())}` + (d.getUTCSeconds() ? ':' + z(d.getUTCSeconds()) : '');
  if (kind === 'time') return time;
  if (kind === 'datetime') return time === '00:00' ? date : date + ' ' + time;
  return date;
}
function tidyNumber(v) {
  if (/^-?\d+\.\d{10,}$/.test(v) || /e-?\d+$/i.test(v)) {
    const x = Number(v);
    if (isFinite(x)) return String(parseFloat(x.toPrecision(15)));
  }
  return v;
}

/** 打开工作簿：读出工作表列表、共享字符串、日期样式（只解压需要的部分，图片不会被解压） */
async function openWorkbook(arrayBuffer) {
  const zip = await unzip(arrayBuffer);
  const sheets = [];
  try {
    const wb = await zip.text('xl/workbook.xml');
    const rels = await zip.text('xl/_rels/workbook.xml.rels');
    const relMap = {};
    rels.replace(/<Relationship\b[^>]*>/g, (tag) => {
      const id = (/\bId="([^"]+)"/.exec(tag) || [])[1];
      const target = (/\bTarget="([^"]+)"/.exec(tag) || [])[1];
      if (id && target) relMap[id] = target;
      return '';
    });
    wb.replace(/<sheet\b[^>]*>/g, (tag) => {
      const name = xmlUnescape((/\bname="([^"]*)"/.exec(tag) || [])[1] || '');
      const rid = (/\br:id="([^"]+)"/.exec(tag) || /\bid="([^"]+)"/.exec(tag) || [])[1];
      const state = (/\bstate="([^"]+)"/.exec(tag) || [])[1] || 'visible';
      let target = relMap[rid];
      if (target) {
        target = target.replace(/^\//, '');
        if (!target.startsWith('xl/')) target = 'xl/' + target;
        target = target.replace(/\/\.\//g, '/');
      }
      if (target && zip.has(target)) sheets.push({ name: name || '工作表' + (sheets.length + 1), path: target, hidden: state !== 'visible' });
      return '';
    });
  } catch (e) { /* 用后备规则 */ }
  if (!sheets.length) {
    zip.names.filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))
      .sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]))
      .forEach((n, i) => sheets.push({ name: 'Sheet' + (i + 1), path: n, hidden: false }));
  }
  if (!sheets.length) throw new Error('xlsx 中没有找到工作表');

  /* 样式：哪些单元格是日期 */
  const dateStyle = [];
  if (zip.has('xl/styles.xml')) {
    const st = await zip.text('xl/styles.xml');
    const custom = {};
    st.replace(/<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g, (_, id, code) => { custom[+id] = xmlUnescape(code); return ''; });
    const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(st);
    if (xfs) {
      const re = /<xf\b([^>]*)/g;
      let m;
      while ((m = re.exec(xfs[1]))) {
        const id = +((/numFmtId="(\d+)"/.exec(m[1]) || [])[1] || 0);
        dateStyle.push(fmtKind(id, custom[id]));
      }
    }
  }
  const shared = [];
  if (zip.has('xl/sharedStrings.xml')) {
    await eachXmlBlock(zip, 'xl/sharedStrings.xml', 'si', (seg) => {
      const re = /<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g;
      let m;
      while ((m = re.exec(seg))) shared.push(m[1] ? xmlTextOf(m[1]) : '');
    });
  }
  let visual = null;
  try { visual = await readVisualStyles(zip); } catch (e) { /* 读不出颜色就不显示 */ }
  return { zip, sheets, dateStyle, shared, visual };
}

/* ---------- Excel 颜色（底色 / 字体颜色 / 加粗）---------- */
const INDEXED_COLORS = ['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','800000','008000','000080','808000','800080','008080','C0C0C0','808080','9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF','000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF','00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99','3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696','003366','339966','003300','333300','993300','993366','333399','333333'];
function applyTint(hex, tint) {
  if (!tint) return hex;
  let r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, sat = 0, l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    sat = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  if (sat === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + sat) : l + sat - l * sat, p = 2 * l - q; r = hue(p, q, h + 1 / 3); g = hue(p, q, h); b = hue(p, q, h - 1 / 3); }
  return [r, g, b].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}
function resolveColor(tag, theme) {
  if (!tag) return null;
  const rgb = (/\brgb="([0-9A-Fa-f]{6,8})"/.exec(tag) || [])[1];
  const tint = +((/\btint="(-?[\d.eE-]+)"/.exec(tag) || [])[1] || 0);
  let hex = null;
  if (rgb) hex = rgb.slice(-6).toUpperCase();
  else {
    const th = (/\btheme="(\d+)"/.exec(tag) || [])[1];
    const ix = (/\bindexed="(\d+)"/.exec(tag) || [])[1];
    if (th != null && theme[+th]) hex = theme[+th];
    else if (ix != null && INDEXED_COLORS[+ix]) hex = INDEXED_COLORS[+ix];
  }
  return hex ? applyTint(hex, tint) : null;
}
async function readVisualStyles(zip) {
  if (!zip.has('xl/styles.xml')) return null;
  const st = await zip.text('xl/styles.xml');
  /* 主题色：0 lt1, 1 dk1, 2 lt2, 3 dk2, 4-9 accent1-6, 10 hlink, 11 folHlink */
  const theme = [];
  const themePath = zip.names.find((n) => /^xl\/theme\/theme\d*\.xml$/.test(n));
  if (themePath) {
    const tx = await zip.text(themePath);
    const pick = (name) => {
      const m = new RegExp('<a:' + name + '>([\\s\\S]*?)</a:' + name + '>').exec(tx);
      if (!m) return null;
      return ((/srgbClr val="([0-9A-Fa-f]{6})"/.exec(m[1]) || /lastClr="([0-9A-Fa-f]{6})"/.exec(m[1]) || [])[1] || '').toUpperCase() || null;
    };
    ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].forEach((n) => theme.push(pick(n)));
  }
  const section = (tag) => (new RegExp('<' + tag + '\\b[^>]*>([\\s\\S]*?)</' + tag + '>').exec(st) || [])[1] || '';
  const fills = [];
  section('fills').replace(/<fill\b[^>]*>([\s\S]*?)<\/fill>|<fill\b[^>]*\/>/g, (all, inner) => {
    inner = inner || '';
    const pt = (/patternType="([^"]+)"/.exec(inner) || [])[1];
    fills.push(pt === 'solid' ? resolveColor((/<fgColor\b[^>]*>/.exec(inner) || [])[0], theme) : null);
    return '';
  });
  const fonts = [];
  section('fonts').replace(/<font\b[^>]*>([\s\S]*?)<\/font>|<font\b[^>]*\/>/g, (all, inner) => {
    inner = inner || '';
    fonts.push({
      color: resolveColor((/<color\b[^>]*>/.exec(inner) || [])[0], theme),
      b: /<b\s*\/>|<b\s+val="(1|true)"/.test(inner),
      i: /<i\s*\/>|<i\s+val="(1|true)"/.test(inner),
    });
    return '';
  });
  const visual = [];
  section('cellXfs').replace(/<xf\b([^>]*)(?:\/>|>[\s\S]*?<\/xf>)/g, (all, attrs) => {
    const fill = fills[+((/fillId="(\d+)"/.exec(attrs) || [])[1] || 0)] || null;
    const font = fonts[+((/fontId="(\d+)"/.exec(attrs) || [])[1] || 0)] || {};
    const v = {};
    if (fill && fill !== 'FFFFFF') v.bg = '#' + fill;
    if (font.color && font.color !== '000000') v.color = '#' + font.color;
    if (font.b) v.b = 1;
    if (font.i) v.i = 1;
    visual.push(Object.keys(v).length ? v : null);
    return '';
  });
  return visual;
}

function xmlTextOf(frag) {
  let out = '';
  frag = frag.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  const re = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
  let m;
  while ((m = re.exec(frag))) out += xmlUnescape(m[1]);
  return out;
}

/** 读取一个工作表，返回二维数组（只保留有内容的单元格） */
async function readSheet(book, sheet, onProgress, wantStyles) {
  const { zip, dateStyle, shared } = book;
  const grid = [];
  const styles = wantStyles ? {} : null;
  const merges = [];
  let rowsSeen = 0;
  const cellRef = (ref) => {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    if (!m) return null;
    let col = 0;
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    return [+m[2] - 1, col - 1];
  };
  await eachXmlBlock(zip, sheet.path, 'row', (seg) => {
    if (seg.indexOf('<mergeCell') >= 0) {
      seg.replace(/<mergeCell\b[^>]*\bref="([A-Z]+\d+):([A-Z]+\d+)"/g, (_, a, b) => { const x = cellRef(a), y = cellRef(b); if (x && y) merges.push([x, y]); return ''; });
    }
    const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
    let rm;
    while ((rm = rowRe.exec(seg))) {
      rowsSeen++;
      if (!rm[2]) continue;
      const rAttr = /\br="(\d+)"/.exec(rm[1]);
      const rowIdx = rAttr ? +rAttr[1] - 1 : rowsSeen - 1;
      const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let cm, auto = 0, line = null;
      while ((cm = cellRe.exec(rm[2]))) {
        const attrs = cm[1];
        const ref = /\br="([A-Z]+)\d*"/.exec(attrs);
        let col = auto;
        if (ref) { col = 0; for (const ch of ref[1]) col = col * 26 + (ch.charCodeAt(0) - 64); col -= 1; }
        auto = col + 1;
        if (styles) {
          const sm = /\bs="(\d+)"/.exec(attrs);
          if (sm && book.visual && book.visual[+sm[1]] && col < 2000) styles[rowIdx + ':' + col] = +sm[1];
        }
        const body = cm[2];
        if (!body) continue; // 只有格式、没有值
        const type = (/\bt="([^"]+)"/.exec(attrs) || [])[1];
        const v = (/<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body) || [])[1];
        let val = '';
        if (type === 's') val = shared[+v] || '';
        else if (type === 'inlineStr') val = xmlTextOf(body);
        else if (type === 'b') val = v === '1' ? 'TRUE' : 'FALSE';
        else if (type === 'e') val = v || '';
        else if (v != null) {
          val = xmlUnescape(v);
          if (type !== 'str') {
            const s = +((/\bs="(\d+)"/.exec(attrs) || [])[1] || 0);
            const kind = dateStyle[s];
            if (kind && val !== '' && !isNaN(+val)) val = excelDate(+val, kind);
            else val = tidyNumber(val);
          }
        }
        if (val === '') continue;
        if (!line) line = [];
        line[col] = val;
      }
      if (line) grid[rowIdx] = line;
    }
    if (onProgress) onProgress(rowsSeen);
  });
  /* 合并单元格：把左上角的值向下填充（例如表头里上下合并的「下单日期」、按客户合并的多行） */
  for (const [[r1, c1], [r2]] of merges) {
    if (r2 <= r1 || r2 - r1 > 500) continue;
    const v = grid[r1] && grid[r1][c1];
    if (v == null || v === '') continue;
    for (let r = r1 + 1; r <= r2; r++) {
      if (!grid[r]) grid[r] = [];
      if (grid[r][c1] == null || grid[r][c1] === '') grid[r][c1] = v;
    }
  }
  const out = [];
  const rowIds = [];
  for (let r = 0; r < grid.length; r++) {
    const src = grid[r];
    if (!src) continue;
    const line = new Array(src.length);
    for (let c = 0; c < src.length; c++) line[c] = src[c] == null ? '' : src[c];
    out.push(line);
    rowIds.push(r);
  }
  out.rowIds = rowIds;
  out.styles = styles;
  out.raw = grid;
  return out;
}

/** 兼容旧接口：只读第一个工作表 */
async function parseXLSX(arrayBuffer, onProgress) {
  const book = await openWorkbook(arrayBuffer);
  return readSheet(book, book.sheets[0], onProgress);
}

/** 读取所有工作表 → [{ name, hidden, table, title }]（空表会被去掉） */
async function readWorkbook(arrayBuffer, onProgress) {
  const book = await openWorkbook(arrayBuffer);
  const out = [];
  book.sheets.forEach((sh, i) => { sh.index = i; });
  for (const sh of book.sheets) {
    const grid = await readSheet(book, sh, onProgress, true);
    const res = cleanGrid(grid, { detectHeader: true, rowIds: grid.rowIds });
    if (!res.table.headers.length) continue;
    /* 只保留表格范围内的样式 */
    const styles = {};
    if (grid.styles && res.src) {
      const rows = [res.src.headerRow, ...res.src.rows];
      for (const r of rows) for (const c of res.src.cols) { const k = r + ':' + c; if (grid.styles[k] != null) styles[k] = grid.styles[k]; }
    }
    out.push({ name: sh.name, index: sh.index, hidden: sh.hidden, xml: sh.path, table: res.table, title: res.title, src: res.src, styles });
  }
  out.visual = book.visual;
  return out;
}

/* ---------- 把修改写回原 Excel（保留颜色、合并单元格、图片、公式、列宽…） ---------- */
async function deflateRaw(u8) {
  const stream = new Blob([u8]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** 重新打包：未改动的条目直接拷贝压缩数据；replaced: Map(名称 → 新文本)；removed: Set(名称) */
async function rezip(zip, replaced, removed) {
  const enc = new TextEncoder();
  const parts = [], central = [];
  let offset = 0;
  for (const name of zip.names) {
    if (removed && removed.has(name)) continue;
    let method, crc, usize, data;
    if (replaced && replaced.has(name)) {
      const plain = enc.encode(replaced.get(name));
      crc = crc32(plain);
      usize = plain.length;
      data = await deflateRaw(plain);
      method = 8;
    } else {
      const r = zip.raw(name);
      ({ method, crc, usize, data } = r);
    }
    const nm = enc.encode(name);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
    lh.setUint16(8, method, true); lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true);
    lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, usize, true);
    lh.setUint16(26, nm.length, true); lh.setUint16(28, 0, true);
    parts.push(new Uint8Array(lh.buffer), nm, data);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
    ch.setUint16(8, 0x0800, true); ch.setUint16(10, method, true); ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true);
    ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, usize, true);
    ch.setUint16(28, nm.length, true); ch.setUint32(42, offset, true);
    central.push(new Uint8Array(ch.buffer), nm);
    offset += 30 + nm.length + data.length;
  }
  const csize = central.reduce((a, b) => a + b.length, 0);
  const count = central.length / 2;
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, count, true); end.setUint16(10, count, true);
  end.setUint32(12, csize, true); end.setUint32(16, offset, true);
  return new Blob([...parts, ...central, new Uint8Array(end.buffer)]).arrayBuffer();
}

const colIndexOf = (letters) => { let c = 0; for (const ch of letters) c = c * 26 + (ch.charCodeAt(0) - 64); return c - 1; };

/**
 * 算出要写回原工作表的单元格：
 * table 当前表格（行带 __src，表带 colSrc），map 导入时的对应关系，orig 原工作表解析结果（含合并单元格填充后的值）
 * 返回 Map(原始行号 0 起 → Map(原始列号 → 新值，'' 表示清空))，以及新行/新列用的样式模板
 */
function sheetChanges(table, map, orig) {
  const origAt = (r, c) => { const row = orig.raw[r]; return row && row[c] != null ? String(row[c]) : ''; };
  let maxRow = orig.raw.length - 1, maxCol = 0;
  for (const row of orig.raw) if (row && row.length - 1 > maxCol) maxCol = row.length - 1;
  for (const c of map.cols0 || map.cols) if (c > maxCol) maxCol = c;
  /* 列对应：优先用 colSrc，没有就按列名找 */
  let colSrc = table.colSrc && table.colSrc.length === table.headers.length ? table.colSrc.slice() : null;
  if (!colSrc) {
    colSrc = table.headers.map((h) => {
      const c = (map.cols0 || map.cols).find((cc) => origAt(map.headerRow, cc).trim() === String(h).trim());
      return c == null ? -1 : c;
    });
  }
  const target = colSrc.map((c) => (c >= 0 ? c : ++maxCol));
  const changes = new Map();
  const set = (r, c, v) => { if (!changes.has(r)) changes.set(r, new Map()); changes.get(r).set(c, v); };
  const H = map.headerRow;
  /* 表头 */
  const cols0 = map.cols0 || map.cols || [];
  table.headers.forEach((h, k) => {
    const o = origAt(H, target[k]);
    const i0 = cols0.indexOf(target[k]);
    if (i0 >= 0 && map.headers0 && map.headers0[i0] === String(h)) return; // 表头和导入时一样，没改过
    if (/^(列|新列)\d+$/.test(String(h)) && o === '') return; // 原表头本来就是空的，自动起的名字不写回
    if (String(h) !== o && !(colSrc[k] < 0 && !h)) set(H, target[k], String(h));
  });
  /* 删掉的列：清空表头和原数据行里的值 */
  const liveCols = new Set(colSrc.filter((c) => c >= 0));
  const deadCols = (map.cols0 || map.cols).filter((c) => !liveCols.has(c));
  const rows0 = map.rows0 || map.rows || [];
  for (const c of deadCols) { if (origAt(H, c)) set(H, c, ''); for (const r of rows0) if (origAt(r, c)) set(r, c, ''); }
  /* 删掉的行：清空它在表格范围内的值（格式保留） */
  const liveRows = new Set(table.rows.map((r) => r.__src).filter((x) => x != null && x >= 0));
  for (const r of rows0) if (!liveRows.has(r)) for (const c of [...liveCols, ...deadCols]) if (origAt(r, c)) set(r, c, '');
  /* 现有行 / 新行 */
  const tmplRow = rows0.length ? rows0[rows0.length - 1] : H;
  const newRows = [];
  for (const row of table.rows) {
    let R = row.__src;
    const isNew = R == null || R < 0;
    if (isNew) { R = ++maxRow; newRows.push(R); }
    row.forEach((v, k) => {
      v = cellStr(v);
      if (isNew) { if (v !== '') set(R, target[k], v); }
      else if (v !== origAt(R, target[k])) set(R, target[k], v);
    });
  }
  const tmplCol = (map.cols0 || map.cols).length ? (map.cols0 || map.cols)[(map.cols0 || map.cols).length - 1] : 0;
  return { changes, newRows: new Set(newRows), newCols: new Set(target.filter((c, k) => colSrc[k] < 0)), tmplRow, tmplCol, maxRow, maxCol, headerRow: H, tableCols: target };
}

/** 公式整体平移（复制到新行时用）：$ 锁定的部分不动，引号里的文字不动 */
function shiftFormula(f, dr, dc) {
  return f.split(/("(?:[^"]|"")*"|&quot;[\s\S]*?&quot;)/).map((part, i) => (i % 2 ? part : part.replace(/(^|[^A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![0-9A-Za-z_(])/g, (m, pre, d1, col, d2, row) => {
    const c = d1 || !dc ? col : colLetter(Math.max(0, colIndexOf(col) + dc));
    const r = d2 ? row : String(Math.max(1, +row + dr));
    return pre + d1 + c + d2 + r;
  }))).join('');
}

/** 把改动打进工作表 XML */
function patchSheetXml(xml, ch, dateStyles) {
  const { changes, newCols, tmplCol } = ch;
  let { newRows, tmplRow } = ch;
  if (!changes.size) return { xml, formulasDropped: 0 };
  const sdStart = xml.indexOf('<sheetData');
  const sdOpenEnd = xml.indexOf('>', sdStart) + 1;
  const selfClosing = xml[sdOpenEnd - 2] === '/';
  const sdEnd = selfClosing ? sdOpenEnd : xml.indexOf('</sheetData>', sdOpenEnd);
  const body = selfClosing ? '' : xml.slice(sdOpenEnd, sdEnd);

  /* 拆出所有行 */
  const rows = new Map(); // 行号(1 起) → { attrs, cells: Map(col → xml) }
  const order = [];
  body.replace(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, (all, attrs, inner) => {
    const rn = +((/\br="(\d+)"/.exec(attrs) || [])[1] || (order.length ? order[order.length - 1] + 1 : 1));
    const cells = new Map();
    if (inner) inner.replace(/<c\b([^>]*?)(?:\/>|>[\s\S]*?<\/c>)/g, (cx, cattrs) => {
      const ref = /\br="([A-Z]+)\d+"/.exec(cattrs);
      if (ref) cells.set(colIndexOf(ref[1]), cx);
      return '';
    });
    rows.set(rn, { attrs, cells });
    order.push(rn);
    return '';
  });
  const styleOf = (cx) => (cx && (/\bs="(\d+)"/.exec(cx) || [])[1]) || null;
  const hasContent = (cx) => !/<f\b/.test(cx) && /<v>[^<]|<is>/.test(cx);

  /* 新行放在哪：很多表在数据下面预先拉好了几十行公式 / 格式的空行，
     新行应该接着最后一行真实数据往下填进这些空行，而不是排到它们后面 */
  const H1 = (ch.headerRow || 0) + 1;
  const newSorted = [...newRows].sort((a, b) => a - b);
  let base = H1;
  if (newSorted.length) {
    const exist = [...rows.keys()].sort((a, b) => b - a);
    for (const rn of exist) {
      if (rn <= H1) break;
      const row = rows.get(rn);
      const chg = changes.get(rn - 1);
      const edited = chg && !newRows.has(rn - 1) && [...chg.values()].some((v) => v !== '');
      const deleted = chg && !newRows.has(rn - 1) && [...chg.keys()].some((c) => row.cells.get(c) && hasContent(row.cells.get(c)));
      const busy = [...row.cells.entries()].some(([c, cx]) => hasContent(cx) && !(deleted && chg.has(c)));
      if (edited || busy) { base = rn; break; }
    }
    /* 把临时行号换成实际行号 */
    const moved = new Map();
    newSorted.forEach((r0, k) => { const m = changes.get(r0); changes.delete(r0); if (m) moved.set(base + k, m); });
    for (const [r0, m] of moved) {
      const cur = changes.get(r0);
      if (cur) for (const [c, v] of m) cur.set(c, v); else changes.set(r0, m);
    }
    newRows = new Set(newSorted.map((_, k) => base + k));
    if (base > H1) tmplRow = base - 1;
  }

  /* 每一列的样式：从最后一行真实数据往上找最近一个有样式的格子 */
  const colStyleCache = new Map();
  const nearStyle = (c) => {
    if (colStyleCache.has(c)) return colStyleCache.get(c);
    let s = null;
    for (let rn = tmplRow + 1, n = 0; rn > H1 && n < 60; rn--, n++) {
      const r = rows.get(rn);
      const st = r && styleOf(r.cells.get(c));
      if (st) { s = st; break; }
    }
    colStyleCache.set(c, s);
    return s;
  };
  const tmplStyle = (rn, c) => {
    if (newCols.has(c)) { const r = rows.get(rn); const s = r && styleOf(r.cells.get(tmplCol)); if (s) return s; return nearStyle(tmplCol); }
    return nearStyle(c);
  };

  /* 删除行 / 列时要清空的格子如果是公式，就保留公式不动（整列共享公式不受影响）；
     新行落在预先拉好公式的行里时，公式列也保持公式 */
  for (const [r0, cols] of changes) {
    const row = rows.get(r0 + 1);
    if (!row) continue;
    for (const [c, v] of [...cols]) {
      if (v !== '' && !newRows.has(r0)) continue;
      const cx = row.cells.get(c);
      if (cx && /<f\b/.test(cx)) cols.delete(c);
    }
    if (!cols.size) changes.delete(r0);
  }
  if (!changes.size) return { xml, formulasDropped: 0 };

  /* 被改动单元格如果是共享公式的主格，整组共享公式转成数值，避免文件损坏 */
  const killSi = new Set();
  for (const [r0, cols] of changes) {
    const row = rows.get(r0 + 1);
    if (!row) continue;
    for (const c of cols.keys()) {
      const cx = row.cells.get(c);
      const f = cx && /<f\b([^>]*)>/.exec(cx);
      if (f && /t="shared"/.test(f[1]) && /\bref="/.test(f[1])) { const si = (/\bsi="(\d+)"/.exec(f[1]) || [])[1]; if (si != null) killSi.add(si); }
    }
  }
  let formulasDropped = 0;
  if (killSi.size) {
    for (const row of rows.values()) {
      for (const [c, cx] of row.cells) {
        const f = /<f\b([^>]*?)(?:\/>|>[\s\S]*?<\/f>)/.exec(cx);
        if (f && /t="shared"/.test(f[1])) {
          const si = (/\bsi="(\d+)"/.exec(f[1]) || [])[1];
          if (killSi.has(si)) { row.cells.set(c, cx.replace(f[0], '')); formulasDropped++; }
        }
      }
    }
  }

  const isNum = (v) => /^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(v) && v.length <= 16;
  const anyDate = dateStyles ? dateStyles.findIndex((k) => k === 'date') : -1;
  const cellXml = (ref, s, v) => {
    if (s == null && anyDate > 0 && /^\d{4}-\d{2}-\d{2}$/.test(v)) s = String(anyDate);
    const sa = s != null ? ` s="${s}"` : '';
    if (v === '') return `<c r="${ref}"${sa}/>`;
    if (isNum(v)) return `<c r="${ref}"${sa}><v>${v}</v></c>`;
    if (/^\d{4}-\d{2}-\d{2}$/.test(v) && s != null && dateStyles && dateStyles[+s]) {
      const serial = Date.UTC(+v.slice(0, 4), +v.slice(5, 7) - 1, +v.slice(8, 10)) / 86400000 + 25569;
      if (isFinite(serial)) return `<c r="${ref}"${sa}><v>${serial}</v></c>`;
    }
    return `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(v)}</t></is></c>`;
  };

  let maxR = 0, maxC = 0;
  for (const [r0, cols] of changes) {
    const rn = r0 + 1;
    let row = rows.get(rn);
    if (!row) { row = { attrs: ` r="${rn}"`, cells: new Map() }; rows.set(rn, row); order.push(rn); }
    for (const [c, v] of cols) {
      const ref = colLetter(c) + rn;
      const old = row.cells.get(c);
      const s = styleOf(old) || (newRows.has(r0) || newCols.has(c) ? tmplStyle(rn, c) : null);
      row.cells.set(c, cellXml(ref, s, v));
      if (c > maxC) maxC = c;
    }
    if (rn > maxR) maxR = rn;
  }
  /* 新行：沿用上一行真实数据的行高、每列格式（空格子也带边框/底色）、公式 */
  const tmpl = rows.get(tmplRow + 1);
  const masters = new Map();
  for (const [rn, row] of rows) for (const [c, cx] of row.cells) {
    const f = /<f\b([^>]*)>([\s\S]*?)<\/f>/.exec(cx);
    if (f && /t="shared"/.test(f[1]) && /\bref="/.test(f[1])) { const si = (/\bsi="(\d+)"/.exec(f[1]) || [])[1]; if (si != null && !masters.has(si)) masters.set(si, { text: f[2], rn, c }); }
  }
  const formulaFor = (cx, c, rn) => {
    const f = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(cx || '');
    if (!f || /t="array"|t="dataTable"/.test(f[1])) return null;
    if (/t="shared"/.test(f[1])) {
      const si = (/\bsi="(\d+)"/.exec(f[1]) || [])[1];
      const m = masters.get(si);
      return m ? shiftFormula(m.text, rn - m.rn, c - m.c) : null;
    }
    return f[2] ? shiftFormula(f[2], rn - (tmplRow + 1), 0) : null;
  };
  const tmplCols = new Set([...(tmpl ? tmpl.cells.keys() : []), ...(ch.tableCols || [])]);
  const mergeAdd = [];
  const tmplMerges = [];
  const mergeRows = new Set();
  xml.replace(/<mergeCell\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g, (all, c1, r1, c2, r2) => {
    for (let r = +r1; r <= +r2 && r - +r1 < 500; r++) mergeRows.add(r);
    if (+r1 === tmplRow + 1 && +r2 === +r1) tmplMerges.push([c1, c2]);
    return '';
  });
  for (const r0 of newRows) {
    const rn = r0 + 1;
    let row = rows.get(rn);
    if (!row) { row = { attrs: ` r="${rn}"`, cells: new Map() }; rows.set(rn, row); order.push(rn); }
    if (tmpl && row.attrs.replace(/\s*\bspans="[^"]*"/, '') === ` r="${rn}"`) {
      row.attrs = ` r="${rn}"` + tmpl.attrs.replace(/\s*\br="\d+"/, '').replace(/\s*\bspans="[^"]*"/, '');
    }
    for (const c of tmplCols) {
      const curCx = row.cells.get(c);
      if (curCx && /<f\b/.test(curCx)) continue; // 本来就有公式
      const fx = tmpl && formulaFor(tmpl.cells.get(c), c, rn); // 公式列：新行也用公式，结果由 Excel 算
      if (!fx && curCx && /<v>|<is>/.test(curCx)) continue;
      const ref = colLetter(c) + rn;
      const s = styleOf(curCx) || tmplStyle(rn, c);
      if (fx) row.cells.set(c, `<c r="${ref}"${s ? ` s="${s}"` : ''}><f>${fx}</f></c>`);
      else if (s && !row.cells.has(c)) row.cells.set(c, `<c r="${ref}" s="${s}"/>`);
      if (c > maxC) maxC = c;
    }
    if (rn > maxR) maxR = rn;
    if (!mergeRows.has(rn)) tmplMerges.forEach(([c1, c2]) => mergeAdd.push(`<mergeCell ref="${c1}${rn}:${c2}${rn}"/>`));
  }
  order.sort((a, b) => a - b);
  const out = [...new Set(order)].map((rn) => {
    const row = rows.get(rn);
    const attrs = row.attrs.replace(/\s*\bspans="[^"]*"/, '');
    if (!row.cells.size) return `<row${attrs}/>`;
    const cells = [...row.cells.entries()].sort((a, b) => a[0] - b[0]).map((x) => x[1]).join('');
    return `<row${attrs}>${cells}</row>`;
  }).join('');
  let res = xml.slice(0, sdStart) + '<sheetData>' + out + '</sheetData>' + xml.slice(selfClosing ? sdOpenEnd : sdEnd + '</sheetData>'.length);
  if (mergeAdd.length) {
    if (/<mergeCells\b/.test(res)) res = res.replace(/<mergeCells\b[^>]*>/, (m) => m.replace(/\s*count="\d+"/, '')).replace('</mergeCells>', mergeAdd.join('') + '</mergeCells>');
    else res = res.replace('</sheetData>', '</sheetData><mergeCells>' + mergeAdd.join('') + '</mergeCells>');
  }
  /* 更新使用范围 */
  res = res.replace(/<dimension\b[^>]*ref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"[^>]*\/>/, (all, c1, r1, c2, r2) => {
    const endC = Math.max(c2 ? colIndexOf(c2) : colIndexOf(c1), maxC);
    const endR = Math.max(+(r2 || r1), maxR);
    return `<dimension ref="${c1}${r1}:${colLetter(endC)}${endR}"/>`;
  });
  return { xml: res, formulasDropped };
}

/** 生成「写回原文件」的新 Excel：sheets = [{ index, table, map }] */
async function patchWorkbook(srcBuf, sheets) {
  const book = await openWorkbook(srcBuf);
  const replaced = new Map();
  let changed = 0, formulasDropped = 0;
  for (const sh of sheets) {
    const meta = book.sheets[sh.index];
    if (!meta) continue;
    const orig = await readSheet(book, meta);
    const ch = sheetChanges(sh.table, sh.map, orig);
    if (!ch.changes.size) continue;
    const xml = await book.zip.text(meta.path);
    const r = patchSheetXml(xml, ch, book.dateStyle);
    replaced.set(meta.path, r.xml);
    for (const m of ch.changes.values()) changed += m.size;
    formulasDropped += r.formulasDropped;
  }
  const removed = new Set();
  if (replaced.size) {
    /* 让 Excel 打开时重新计算公式；删掉计算链（Excel 会自动重建） */
    if (book.zip.has('xl/calcChain.xml')) {
      removed.add('xl/calcChain.xml');
      const ct = await book.zip.text('[Content_Types].xml');
      replaced.set('[Content_Types].xml', ct.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, ''));
      const rels = await book.zip.text('xl/_rels/workbook.xml.rels');
      replaced.set('xl/_rels/workbook.xml.rels', rels.replace(/<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/, ''));
    }
    let wb = await book.zip.text('xl/workbook.xml');
    if (/<calcPr\b/.test(wb)) wb = wb.replace(/<calcPr\b([^>]*?)(\/?)>/, (all, attrs, sl) => `<calcPr${attrs.replace(/\s*fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${sl}>`);
    else if (/<\/definedNames>/.test(wb)) wb = wb.replace('</definedNames>', '</definedNames><calcPr fullCalcOnLoad="1"/>');
    else wb = wb.replace('</sheets>', '</sheets><calcPr fullCalcOnLoad="1"/>');
    replaced.set('xl/workbook.xml', wb);
  }
  const out = await rezip(book.zip, replaced, removed);
  return { buffer: out, changed, formulasDropped, sheetsChanged: [...replaced.keys()].filter((k) => k.startsWith('xl/worksheets/')).length };
}

/* ---------- XLSX 写出（不压缩的 zip，无第三方依赖） ---------- */
let CRC_TABLE = null;
function crc32(u8) {
  if (!CRC_TABLE) {
    CRC_TABLE = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      CRC_TABLE[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (let i = 0; i < u8.length; i++) crc = CRC_TABLE[(crc ^ u8[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function zipStore(files) {
  const enc = new TextEncoder();
  const parts = [], central = [];
  let offset = 0;
  for (const f of files) {
    const name = enc.encode(f.name), data = enc.encode(f.content), crc = crc32(data);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
    lh.setUint16(8, 0, true); lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true);
    lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true);
    lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
    parts.push(new Uint8Array(lh.buffer), name, data);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
    ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true); ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true);
    ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true);
    ch.setUint16(28, name.length, true); ch.setUint32(42, offset, true);
    central.push(new Uint8Array(ch.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const csize = central.reduce((a, b) => a + b.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
  end.setUint32(12, csize, true); end.setUint32(16, offset, true);
  const all = [...parts, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(all.reduce((a, b) => a + b.length, 0));
  let p = 0;
  for (const a of all) { out.set(a, p); p += a.length; }
  return out.buffer;
}

function colLetter(i) {
  let s = '';
  for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}

function toXLSX(t) {
  return toXLSXBook([{ name: 'Sheet1', table: t }]);
}

function sheetXml(t) {
  const isNum = (v) => /^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(v) && v.length <= 15;
  const isDate = (v) => /^\d{4}-\d{2}-\d{2}$/.test(v);
  const rowsXml = [t.headers, ...t.rows].map((r, ri) => {
    const cells = r.map((v, ci) => {
      v = cellStr(v);
      if (v === '') return '';
      const ref = colLetter(ci) + (ri + 1);
      if (ri > 0 && isNum(v)) return `<c r="${ref}"><v>${v}</v></c>`;
      if (ri > 0 && isDate(v)) {
        const serial = Date.UTC(+v.slice(0, 4), +v.slice(5, 7) - 1, +v.slice(8, 10)) / 86400000 + 25569;
        if (isFinite(serial)) return `<c r="${ref}" s="1"><v>${serial}</v></c>`;
      }
      return `<c r="${ref}" t="inlineStr"${ri === 0 ? ' s="2"' : ''}><is><t xml:space="preserve">${xmlEscape(v)}</t></is></c>`;
    }).join('');
    return `<row r="${ri + 1}">${cells}</row>`;
  }).join('');
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const cols = t.headers.map((h, i) => {
    let w = String(h).length * 2;
    for (let r = 0; r < Math.min(t.rows.length, 200); r++) {
      const v = cellStr(t.rows[r][i]);
      let len = 0;
      for (const ch of v) len += ch.charCodeAt(0) > 255 ? 2 : 1;
      if (len > w) w = len;
    }
    return `<col min="${i + 1}" max="${i + 1}" width="${Math.min(60, Math.max(8, w + 2))}" customWidth="1"/>`;
  }).join('');
  return head + `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${rowsXml}</sheetData></worksheet>`;
}

const XLSX_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const XLSX_STYLES = XLSX_HEAD + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="2"><font><sz val="11"/><name val="等线"/></font><font><b/><sz val="11"/><name val="等线"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf/><xf numFmtId="164" applyNumberFormat="1"/><xf fontId="1" applyFont="1"/></cellXfs></styleSheet>';

/** 多工作表 Excel：sheets = [{ name, table }] */
function toXLSXBook(sheets) {
  const head = XLSX_HEAD;
  const used = new Set();
  const names = sheets.map((sh, i) => {
    let n = String(sh.name || 'Sheet' + (i + 1)).replace(/[\[\]:*?\/\\]/g, '_').replace(/^'|'$/g, '').trim().slice(0, 31) || 'Sheet' + (i + 1);
    let k = 2;
    while (used.has(n.toLowerCase())) n = n.slice(0, 28) + '_' + k++;
    used.add(n.toLowerCase());
    return n;
  });
  const files = [
    { name: '[Content_Types].xml', content: head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
    { name: '_rels/.rels', content: head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
    { name: 'xl/workbook.xml', content: head + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + names.map((n, i) => `<sheet name="${xmlEscape(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>' },
    { name: 'xl/_rels/workbook.xml.rels', content: head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') + `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
    { name: 'xl/styles.xml', content: XLSX_STYLES },
  ];
  sheets.forEach((sh, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, content: sheetXml(sh.table) }));
  return zipStore(files);
}

/* ---------- AI 返回的操作 → 应用到表格 ---------- */
function findCol(t, name) {
  if (typeof name === 'number') return name >= 1 && name <= t.headers.length ? name - 1 : -1;
  const s = cellStr(name).trim();
  let i = t.headers.indexOf(s);
  if (i < 0) i = t.headers.findIndex((h) => h.trim().toLowerCase() === s.toLowerCase());
  if (i < 0) {
    const ns = s.replace(/\s+/g, '').toLowerCase();
    if (ns) i = t.headers.findIndex((h) => h.replace(/\s+/g, '').toLowerCase() === ns);
  }
  return i;
}

function addColumn(t, name, values, def) {
  t.headers.push(String(name));
  if (t.colSrc) t.colSrc.push(-1);
  t.rows.forEach((r, i) => {
    const v = values && values[i] != null ? values[i] : def != null ? def : '';
    r.push(String(v));
  });
  return t.headers.length - 1;
}

const keyNorm = (v) => cellStr(v).trim().replace(/\s+/g, '').toLowerCase();





/**
 * 宽松修复 AI 返回的 JSON：字符串里的英文双引号、直接换行、多余逗号、被截断没写完的结尾。
 */





const hKey = (h) => cellStr(h).replace(/[\s（）()【】\[\]：:、\/\-_*#.。，,]/g, '').toLowerCase();
/** 两组列名的相似度 0~1（Dice 系数，允许包含关系） */
function headerScore(a, b) {
  const A = [...new Set(a.map(hKey).filter((k) => k && !/^(列|新列)\d+$/.test(k)))];
  const B = [...new Set(b.map(hKey).filter((k) => k && !/^(列|新列)\d+$/.test(k)))];
  if (!A.length || !B.length) return 0;
  let hit = 0;
  for (const x of A) if (B.some((y) => y === x || (x.length >= 2 && y.length >= 2 && (y.includes(x) || x.includes(y))))) hit++;
  return (2 * hit) / (A.length + B.length);
}


/* ====================================================================
 * 二、AI 接口
 * ==================================================================== */

/* ---------- 附件读取 ---------- */


const nodeObsidian = () => require('obsidian');

/* ====================================================================
 * 外部修改对比：找出两版表格之间新增 / 删除 / 修改了什么
 * ==================================================================== */
function hashStr(s) {
  let h1 = 0x811c9dc5, h2 = 5381;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    h1 = Math.imul(h1 ^ c, 16777619);
    h2 = (Math.imul(h2, 33) + c) | 0;
  }
  return (h1 >>> 0).toString(36) + (h2 >>> 0).toString(36) + s.length.toString(36);
}

function diffTables(a, b) {
  const used = new Set();
  const colMap = b.headers.map((h) => {
    const i = a.headers.findIndex((x, k) => !used.has(k) && String(x) === String(h));
    if (i >= 0) used.add(i);
    return i;
  });
  const renamed = [];
  b.headers.forEach((h, j) => {
    if (colMap[j] < 0 && j < a.headers.length && !used.has(j)) { colMap[j] = j; used.add(j); renamed.push([String(a.headers[j]), String(h)]); }
  });
  const addedCols = b.headers.map((h, j) => (colMap[j] < 0 ? j : -1)).filter((j) => j >= 0);
  const removedCols = a.headers.map((h, i) => (used.has(i) ? null : String(h))).filter((x) => x != null);
  const common = colMap.map((i, j) => [i, j]).filter(([i]) => i >= 0);
  const keyA = a.rows.map((r) => common.map(([i]) => cellStr(r[i])).join('\u0001'));
  const keyB = b.rows.map((r) => common.map(([, j]) => cellStr(r[j])).join('\u0001'));
  const n = keyA.length, m = keyB.length;
  const rowMap = new Array(m).fill(-1);
  let p = 0;
  while (p < n && p < m && keyA[p] === keyB[p]) { rowMap[p] = p; p++; }
  let s = 0;
  while (s < n - p && s < m - p && keyA[n - 1 - s] === keyB[m - 1 - s]) { rowMap[m - 1 - s] = n - 1 - s; s++; }
  const A0 = p, A1 = n - s, B0 = p, B1 = m - s, na = A1 - A0, nb = B1 - B0;
  const pairs = [];
  if (na && nb) {
    if (na * nb <= 4e6 && na < 65535 && nb < 65535) {
      const W = nb + 1, dp = new Uint16Array((na + 1) * W);
      for (let i = na - 1; i >= 0; i--) {
        for (let j = nb - 1; j >= 0; j--) {
          dp[i * W + j] = keyA[A0 + i] === keyB[B0 + j] ? dp[(i + 1) * W + j + 1] + 1 : Math.max(dp[(i + 1) * W + j], dp[i * W + j + 1]);
        }
      }
      let i = 0, j = 0;
      while (i < na && j < nb) {
        if (keyA[A0 + i] === keyB[B0 + j]) { pairs.push([A0 + i, B0 + j]); i++; j++; }
        else if (dp[(i + 1) * W + j] >= dp[i * W + j + 1]) i++;
        else j++;
      }
    } else {
      const pos = new Map();
      for (let i = A0; i < A1; i++) { const k = keyA[i]; if (!pos.has(k)) pos.set(k, []); pos.get(k).push(i); }
      let last = A0 - 1;
      for (let j = B0; j < B1; j++) {
        const q = pos.get(keyB[j]);
        if (!q) continue;
        while (q.length && q[0] <= last) q.shift();
        if (q.length) { last = q.shift(); pairs.push([last, j]); }
      }
    }
  }
  pairs.forEach(([i, j]) => { rowMap[j] = i; });
  const cells = [], added = [], removed = [];
  const cmpRow = (i, j) => {
    for (const [ci, cj] of common) {
      const o = cellStr(a.rows[i][ci]), v = cellStr(b.rows[j][cj]);
      if (o !== v) cells.push({ r: j, c: cj, o, n: v });
    }
  };
  const sim = (i, j) => {
    let eq = 0, tot = 0;
    for (const [ci, cj] of common) {
      const o = cellStr(a.rows[i][ci]), v = cellStr(b.rows[j][cj]);
      if (o || v) { tot++; if (o === v) eq++; }
    }
    return tot ? eq / tot : 0;
  };
  const anchors = [[A0 - 1, B0 - 1], ...pairs, [A1, B1]];
  for (let k = 0; k + 1 < anchors.length; k++) {
    const ga = [], gb = [];
    for (let i = anchors[k][0] + 1; i < anchors[k + 1][0]; i++) ga.push(i);
    for (let j = anchors[k][1] + 1; j < anchors[k + 1][1]; j++) gb.push(j);
    if (!ga.length) { added.push(...gb); continue; }
    if (!gb.length) { removed.push(...ga); continue; }
    const taken = new Set();
    if (ga.length * gb.length <= 40000) {
      for (const i of ga) {
        let best = -1, bs = 0.34;
        for (const j of gb) { if (taken.has(j)) continue; const sc = sim(i, j); if (sc > bs) { bs = sc; best = j; } }
        if (best >= 0) { taken.add(best); rowMap[best] = i; cmpRow(i, best); } else removed.push(i);
      }
    } else {
      const c = Math.min(ga.length, gb.length);
      for (let x = 0; x < c; x++) { taken.add(gb[x]); rowMap[gb[x]] = ga[x]; cmpRow(ga[x], gb[x]); }
      removed.push(...ga.slice(c));
    }
    gb.forEach((j) => { if (!taken.has(j)) added.push(j); });
  }
  cells.sort((x, y) => x.r - y.r || x.c - y.c);
  added.sort((x, y) => x - y);
  removed.sort((x, y) => x - y);
  const filled = addedCols.map((j) => b.rows.filter((r) => !isBlank(r[j])).length);
  return { colMap, renamed, addedCols, removedCols, rowMap, cells, added, removed, filled };
}

/* 交给外部 AI 时的常用要求 */
const AI_TEMPLATES = [
  ['录入单据', '把附件里的单据内容录入这张表：按表格现有的列整理，每条记录一行，加在表格末尾；已经有的记录（同一单号，或同一品名 + 规格 + 日期）就更新那一行，不要重复添加。'],
  ['按价格表更新', '用附件里的价格表更新这张表的价格：按品名 + 规格匹配，只改价格相关的列，不要新增或删除行；表里找不到的品名请在回复里列出来。'],
  ['核对数据', '检查这张表有没有重复行、空缺、算错（例如 数量 × 单价 ≠ 金额）或格式不统一的地方。先在回复里列出来，不要修改文件。'],
  ['整理分类', '根据每一行的内容判断类别，在表格最后新增一列「分类」并填好。不要改动其它列，也不要调整行的顺序。'],
];

const ARCHIVE_KINDS = [['全部', null], ['图片', 'image'], ['PDF', 'pdf'], ['表格', 'table'], ['文档', 'doc'], ['其他', 'other']];
function archiveKind(f) {
  const e = (f.extension || (f.name || '').split('.').pop() || '').toLowerCase();
  if (/^(png|jpe?g|webp|gif|bmp|heic|svg)$/.test(e)) return 'image';
  if (e === 'pdf') return 'pdf';
  if (/^(csv|tsv|xlsx|xlsm|xls|ods|numbers)$/.test(e)) return 'table';
  if (/^(docx?|wps|txt|md|pptx?|rtf|odt)$/.test(e)) return 'doc';
  return 'other';
}
/** 把附件区的图片复制到剪贴板（可以直接粘贴到微信、Word、Excel 等） */
async function copyImageFile(app, tf) {
  const buf = await app.vault.readBinary(tf);
  try {
    const { clipboard, nativeImage } = require('electron');
    const img = nativeImage.createFromBuffer(Buffer.from(buf));
    if (img && !img.isEmpty()) { clipboard.writeImage(img); return true; }
  } catch (e) { /* 不是桌面版，换浏览器的方法 */ }
  const ext = (tf.extension || '').toLowerCase();
  let blob = new Blob([buf], { type: MIME[ext] || 'image/png' });
  if (blob.type !== 'image/png') {
    const bmp = await createImageBitmap(blob);
    const cv = document.createElement('canvas');
    cv.width = bmp.width; cv.height = bmp.height;
    cv.getContext('2d').drawImage(bmp, 0, 0);
    blob = await new Promise((r) => cv.toBlob(r, 'image/png'));
  }
  await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
  return true;
}
async function copyImageNotice(app, tf) {
  try { await copyImageFile(app, tf); new Notice(`已复制图片「${tf.name}」，可以直接粘贴。`); }
  catch (e) { new Notice('复制图片失败：' + (e.message || e)); }
}

const fmtSize = (n) => (n >= 1048576 ? (n / 1048576).toFixed(1) + ' MB' : n >= 1024 ? Math.round(n / 1024) + ' KB' : n + ' B');

/** 用 Obsidian 自带的 PDF.js 把 PDF 渲染成图片（附件区预览用） */
async function renderPdfPages(app, tf, width, maxPages) {
  const lib = await nodeObsidian().loadPdfJs();
  const data = new Uint8Array(await app.vault.readBinary(tf));
  const doc = await lib.getDocument({ data }).promise;
  const urls = [];
  for (let i = 1; i <= Math.min(doc.numPages, maxPages || 1); i++) {
    const page = await doc.getPage(i);
    const vp0 = page.getViewport({ scale: 1 });
    const vp = page.getViewport({ scale: width / vp0.width });
    const cv = document.createElement('canvas');
    cv.width = Math.round(vp.width);
    cv.height = Math.round(vp.height);
    const ctx = cv.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, cv.width, cv.height);
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    urls.push(cv.toDataURL('image/jpeg', 0.82));
  }
  const total = doc.numPages;
  try { doc.destroy(); } catch (e) { /* 忽略 */ }
  return { urls, total };
}


/** 读取表格文件 → { table, title } */
async function fileToTable(file, onProgress) {
  if (file.size > MAX_FILE) throw new Error(`文件超过 200MB（${(file.size / 1048576).toFixed(0)}MB）`);
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (ext === 'xlsx' || ext === 'xlsm') return cleanGrid(await parseXLSX(await file.arrayBuffer(), onProgress), { detectHeader: true });
  if (ext === 'xls') throw new Error('不支持旧版 .xls，请先在 Excel 里另存为 .xlsx 或 .csv。');
  const text = decodeText(await file.arrayBuffer());
  if (ext === 'md' || ext === 'markdown') {
    const f = findMdTable(text);
    if (!f) throw new Error(`「${file.name}」里没有找到 Markdown 表格。`);
    return cleanGrid([f.table.headers, ...f.table.rows]);
  }
  return cleanGrid(parseCSV(text), { detectHeader: true });
}

/** 读取表格文件 → [{ name, index, table, title }]；xlsx 会读出所有工作表 */
async function fileToSheets(file, onProgress) {
  if (file.size > MAX_FILE) throw new Error(`文件超过 200MB（${(file.size / 1048576).toFixed(0)}MB）`);
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (ext === 'xlsx' || ext === 'xlsm') {
    const sheets = await readWorkbook(await file.arrayBuffer(), onProgress);
    const visible = sheets.filter((s) => !s.hidden);
    const out = visible.length ? visible : sheets;
    out.visual = sheets.visual;
    return out;
  }
  const r = await fileToTable(file, onProgress);
  return [{ name: '', index: 0, table: r.table, title: r.title }];
}

const baseName = (name) => name.replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|#^[\]]/g, '_').trim() || '表格';
const TABLE_EXT = /\.(csv|tsv|xlsx|xlsm|md|markdown)$/i;
const safeSeg = (s) => String(s).replace(/[\\/:*?"<>|#^[\]]/g, '_').replace(/^\.+/, '').trim();

/** 在 drop 事件里同步调用：把拖入的文件和文件夹（递归）展开成 [{file, rel}] */
function collectDropped(dt) {
  const items = dt && dt.items ? Array.from(dt.items).filter((i) => i.kind === 'file') : [];
  // DataTransfer 只在事件同步阶段有效，先把 entry 和 file 全部取出来
  const pairs = items.map((i) => ({ entry: i.webkitGetAsEntry ? i.webkitGetAsEntry() : null, file: i.getAsFile() }));
  const plain = pairs.length ? [] : Array.from((dt && dt.files) || []);
  return (async () => {
    const out = [];
    const walk = async (entry, prefix) => {
      if (entry.name.startsWith('.') || entry.name.startsWith('~$')) return;
      if (entry.isFile) {
        const f = await new Promise((res, rej) => entry.file(res, rej));
        out.push({ file: f, rel: prefix + entry.name });
      } else if (entry.isDirectory) {
        const reader = entry.createReader();
        for (;;) {
          const batch = await new Promise((res, rej) => reader.readEntries(res, rej));
          if (!batch.length) break;
          for (const e of batch) await walk(e, prefix + entry.name + '/');
        }
      }
    };
    for (const p of pairs) {
      if (p.entry) await walk(p.entry, '');
      else if (p.file) out.push({ file: p.file, rel: p.file.name });
    }
    for (const f of plain) out.push({ file: f, rel: f.name });
    return out;
  })();
}

/* ---------- 按路径读取本机文件（桌面版）：用于微信等拖出来只给路径、或粘贴路径的情况 ---------- */
const nodeReq = (m) => { try { return require(m); } catch (e) { return null; } };
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', webp: 'image/webp', gif: 'image/gif', bmp: 'image/bmp', pdf: 'application/pdf', csv: 'text/csv', tsv: 'text/tab-separated-values', md: 'text/markdown', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12' };
const USEFUL_EXT = /\.(csv|tsv|xlsx|xlsm|md|markdown|png|jpe?g|webp|gif|bmp|pdf)$/i;

function parsePaths(text) {
  return String(text || '').split(/[\r\n]+/)
    .map((s) => s.trim().replace(/^["'“”]+|["'“”]+$/g, '').trim())
    .map((s) => {
      if (/^file:\/\/\//i.test(s)) { try { s = decodeURIComponent(s.slice(8)); } catch (e) { s = s.slice(8); } return /^[a-zA-Z]:/.test(s) ? s : '/' + s; }
      if (/^file:\/\//i.test(s)) { try { return decodeURIComponent(s.slice(7)); } catch (e) { return s.slice(7); } }
      return s;
    })
    .filter((s) => /^([a-zA-Z]:[\\/]|\\\\|\/)/.test(s));
}

/** 把本机路径（文件或文件夹）读成 [{file, rel}] */
async function filesFromPaths(paths) {
  const fs = nodeReq('fs'), path = nodeReq('path');
  if (!fs || !path) throw new Error('只有电脑版 Obsidian 才能按路径读取文件。');
  const found = [], missing = [];
  for (const p of paths) {
    let st = null;
    try { st = await fs.promises.stat(p); } catch (e) { missing.push(p); continue; }
    if (st.isDirectory()) {
      const base = path.dirname(p);
      const walk = async (dir, depth) => {
        if (depth > 8 || found.length > 5000) return;
        let names = [];
        try { names = await fs.promises.readdir(dir); } catch (e) { return; }
        for (const name of names) {
          if (name.startsWith('.') || name.startsWith('~$')) continue;
          const full = path.join(dir, name);
          const s = await fs.promises.stat(full).catch(() => null);
          if (!s) continue;
          if (s.isDirectory()) await walk(full, depth + 1);
          else found.push({ full, rel: path.relative(base, full).split(path.sep).join('/'), size: s.size, mtime: s.mtimeMs });
        }
      };
      await walk(p, 0);
    } else found.push({ full: p, rel: path.basename(p), size: st.size, mtime: st.mtimeMs });
  }
  const out = [];
  for (const x of found) {
    const name = path.basename(x.full);
    if (!USEFUL_EXT.test(name)) { out.push({ file: new File([], name), rel: x.rel }); continue; }
    if (x.size > MAX_FILE) { missing.push(`${x.full}（超过 200MB）`); continue; }
    const buf = await fs.promises.readFile(x.full);
    const ext = (name.split('.').pop() || '').toLowerCase();
    out.push({ file: new File([buf], name, { type: MIME[ext] || '', lastModified: x.mtime }), rel: x.rel });
  }
  if (missing.length) new Notice(`有 ${missing.length} 个路径读不到：${missing.slice(0, 2).join('；')}${missing.length > 2 ? '…' : ''}`, 10000);
  return out;
}

/** 读系统剪贴板里「复制的文件」（资源管理器 Ctrl+C） */
function clipboardFilePaths() {
  const el = nodeReq('electron');
  const cb = el && el.clipboard;
  if (!cb) return [];
  try {
    const buf = cb.readBuffer && cb.readBuffer('FileNameW');
    if (buf && buf.length) {
      const p = buf.toString('ucs2').replace(/\u0000+$/, '');
      if (p) return [p];
    }
  } catch (e) { /* 不支持 */ }
  try {
    const t = cb.read && cb.read('public.file-url');
    if (t) return parsePaths(t);
  } catch (e) { /* 非 mac */ }
  return [];
}


/* ====================================================================
 * 三、弹窗
 * ==================================================================== */
class TablePickModal extends FuzzySuggestModal {
  constructor(app, files, onPick, meta, placeholder) {
    super(app);
    this.files = files;
    this.onPick = onPick;
    this.meta = meta || {};
    this.setPlaceholder(placeholder || '选择表格（.csv 文件，或含表格的 .md 笔记）');
  }
  getItems() {
    if (this.files) return this.files;
    return this.app.vault.getFiles()
      .filter((f) => f.extension === 'csv' || f.extension === 'md')
      .sort((a, b) => (a.extension === b.extension ? b.stat.mtime - a.stat.mtime : a.extension === 'csv' ? -1 : 1));
  }
  getItemText(f) {
    const m = this.meta[f.path];
    return m && m.sheet ? `${m.sheet}  ·  ${m.workbook}` : f.path;
  }
  onChooseItem(f) { this.onPick(f); }
}

class NewTableModal extends Modal {
  constructor(app, onSubmit) { super(app); this.onSubmit = onSubmit; }
  onOpen() {
    let name = '新表格', cols = '', fmt = 'csv';
    this.titleEl.setText('新建表格');
    new Setting(this.contentEl).setName('名称').addText((t) => t.setValue(name).onChange((v) => (name = v)));
    new Setting(this.contentEl).setName('列名').setDesc('用逗号分隔，例如：日期, 摘要, 科目, 借方, 贷方。可以留空，让 AI 根据内容自动建列。')
      .addText((t) => t.setPlaceholder('日期, 摘要, 金额').onChange((v) => (cols = v)));
    new Setting(this.contentEl).setName('保存为')
      .addDropdown((d) => d.addOption('csv', 'CSV 文件').addOption('md', 'Markdown 笔记').onChange((v) => (fmt = v)));
    new Setting(this.contentEl).addButton((b) => b.setButtonText('创建').setCta().onClick(() => {
      this.close();
      this.onSubmit(name.trim() || '新表格', cols.split(/[,，、]/).map((s) => s.trim()).filter(Boolean), fmt);
    }));
  }
  onClose() { this.contentEl.empty(); }
}

class ExportModal extends Modal {
  constructor(app, onPick, book, linked) { super(app); this.onPick = onPick; this.book = book; this.linked = linked; }
  onOpen() {
    this.titleEl.setText('导出表格');
    if (this.linked) {
      new Setting(this.contentEl)
        .setName(this.book ? `整个工作簿，保留原 Excel 格式（推荐）` : 'Excel，保留原格式（推荐）')
        .setDesc(this.book
          ? `把这 ${this.book.count} 个工作表里的修改写回原来的 Excel：颜色、合并单元格、公式、图片、列宽、其它工作表都保留。`
          : '把修改写回原来的 Excel：颜色、合并单元格、公式、图片、列宽都保留。')
        .addButton((b) => b.setButtonText('导出').setCta().onClick(() => { this.close(); this.onPick('linked'); }));
    }
    if (this.book) {
      new Setting(this.contentEl).setName(this.linked ? '整个工作簿，干净格式' : '整个工作簿 → 一个 Excel').setDesc(`把「${this.book.name}」的 ${this.book.count} 个工作表导出成一个新的多工作表 .xlsx（只有数据，不带原来的颜色）`)
        .addButton((b) => { b.setButtonText('导出'); if (!this.linked) b.setCta(); b.onClick(() => { this.close(); this.onPick('book'); }); });
    }
    const opts = [
      ['xlsx', this.book ? '只导出当前工作表 (.xlsx)' : 'Excel (.xlsx)', '用 Excel / WPS / Numbers 打开'],
      ['csv', 'CSV (.csv)', '通用格式，带 BOM，Excel 打开不乱码'],
      ['md', 'Markdown (.md)', '导出为一篇含表格的笔记'],
    ];
    for (const [fmt, name, desc] of opts) {
      new Setting(this.contentEl).setName(name).setDesc(desc)
        .addButton((b) => b.setButtonText('导出').onClick(() => { this.close(); this.onPick(fmt); }));
    }
  }
  onClose() { this.contentEl.empty(); }
}

class CategoryModal extends Modal {
  constructor(app, current, existing, onSubmit) {
    super(app);
    this.current = current || '';
    this.existing = existing;
    this.onSubmit = onSubmit;
  }
  onOpen() {
    this.titleEl.setText('设置分类');
    let val = this.current;
    let input;
    new Setting(this.contentEl).setName('分类名称').setDesc('留空表示「未分类」。')
      .addText((t) => { input = t; t.setValue(val).onChange((v) => (val = v)); });
    if (this.existing.length) {
      const wrap = this.contentEl.createDiv('ata-catbar');
      this.existing.forEach((c) => wrap.createEl('button', { text: c, cls: 'ata-catchip' })
        .addEventListener('click', () => { val = c; input.setValue(c); }));
    }
    new Setting(this.contentEl).addButton((b) => b.setButtonText('保存').setCta().onClick(() => {
      this.close();
      this.onSubmit(val.trim());
    }));
  }
  onClose() { this.contentEl.empty(); }
}


/** 附件区：大图 / PDF 预览 */
class LightboxModal extends Modal {
  constructor(app, view, tf) { super(app); this.view = view; this.tf = tf; }
  async onOpen() {
    const v = this.view, tf = this.tf;
    this.modalEl.addClass('ata-lightbox');
    this.titleEl.setText(tf.name);
    const bar = this.contentEl.createDiv('ata-lb-bar');
    const btn = (text, fn, cta) => { const b = bar.createEl('button', { text, cls: cta ? 'mod-cta' : '' }); b.addEventListener('click', fn); };
    btn('交给 AI 填表', () => { this.close(); v.handoff([tf]); }, true);
    if (v.file) btn(`加到「${(v.settings.meta[v.file.path] || {}).sheet || v.file.basename}」的附件`, () => { this.close(); v.addAiAtts(v.file.path, [tf]); v.openTable(v.file); });
    btn('复制文件路径', () => v.copyText(v.absPath(tf.path), '文件路径'));
    if (archiveKind(tf) === 'image') btn('复制图片', () => copyImageNotice(this.app, tf));
    btn('在 Obsidian 中打开', () => { this.close(); this.app.workspace.getLeaf('tab').openFile(tf); });
    if (this.app.openWithDefaultApp) btn('用默认程序打开', () => this.app.openWithDefaultApp(tf.path));
    const body = this.contentEl.createDiv('ata-lb-body');
    const kind = archiveKind(tf);
    if (kind === 'image') {
      body.createEl('img', { attr: { src: this.app.vault.getResourcePath(tf) } });
    } else if (kind === 'pdf') {
      body.createDiv({ cls: 'ata-muted', text: '正在渲染 PDF…' });
      try {
        const pages = await renderPdfPages(this.app, tf, 900, 5);
        body.empty();
        pages.urls.forEach((u, i) => { body.createEl('img', { attr: { src: u } }); body.createDiv({ cls: 'ata-lb-page', text: `第 ${i + 1} / ${pages.total} 页` }); });
        if (pages.total > pages.urls.length) body.createDiv({ cls: 'ata-muted', text: `只预览了前 ${pages.urls.length} 页，完整内容请「在 Obsidian 中打开」。` });
      } catch (e) {
        body.empty();
        body.createDiv({ cls: 'ata-muted', text: 'PDF 预览失败，可以「在 Obsidian 中打开」查看。' });
      }
    } else {
      body.createDiv({ cls: 'ata-muted', text: '这类文件不能直接预览，可以用默认程序打开。' });
    }
    const m = (v.settings.archiveMeta || {})[tf.path] || {};
    if (m.used && m.used.length) this.contentEl.createDiv({ cls: 'ata-muted', text: '已提取到：' + m.used.map((p) => p.split('/').pop().replace(/\.csv$/, '')).join('、') });
  }
  onClose() { this.contentEl.empty(); }
}

class TextPromptModal extends Modal {
  constructor(app, title, value, onSubmit, copyOnly) { super(app); this.t = title; this.v = value || ''; this.onSubmit = onSubmit; this.copyOnly = copyOnly; }
  onOpen() {
    this.titleEl.setText(this.t);
    const ta = this.contentEl.createEl('textarea', { cls: 'ata-path-input', attr: { rows: this.copyOnly ? '14' : '3' } });
    ta.value = this.v;
    setTimeout(() => ta.focus(), 30);
    if (this.copyOnly) {
      new Setting(this.contentEl).addButton((b) => b.setButtonText('复制').setCta().onClick(async () => {
        try { await navigator.clipboard.writeText(ta.value); new Notice('已复制。'); } catch (e) { ta.select(); document.execCommand('copy'); new Notice('已复制。'); }
        this.close();
      }));
      return;
    }
    new Setting(this.contentEl).addButton((b) => b.setButtonText('保存').setCta().onClick(() => { this.close(); this.onSubmit(ta.value.trim()); }));
  }
  onClose() { this.contentEl.empty(); }
}

/** 工作区里「从附件区选」 */
class ArchivePickModal extends Modal {
  constructor(app, view, onPick) { super(app); this.view = view; this.onPick = onPick; this.picked = new Set(); this.q = ''; }
  onOpen() {
    const v = this.view;
    this.modalEl.addClass('ata-route-modal');
    this.titleEl.setText('从附件区选择文件');
    const search = this.contentEl.createEl('input', { cls: 'ata-route-search', attr: { type: 'text', placeholder: '搜索文件名、备注…' } });
    search.addEventListener('input', () => { this.q = search.value.trim().toLowerCase(); this.draw(); });
    this.grid = this.contentEl.createDiv('ata-apick');
    const bar = this.contentEl.createDiv('ata-route-bar');
    this.countEl = bar.createSpan({ cls: 'ata-muted', text: '已选 0 个' });
    const ok = bar.createEl('button', { cls: 'mod-cta', text: '加到附件' });
    ok.addEventListener('click', async () => {
      const files = v.archiveFiles().filter((f) => this.picked.has(f.path));
      if (!files.length) { new Notice('还没有选择文件。'); return; }
      this.close();
      if (this.onPick) this.onPick(files);
      else if (v.file) { v.addAiAtts(v.file.path, files); v.renderAiPanel(); }
    });
    this.draw();
    search.focus();
  }
  draw() {
    const v = this.view;
    this.grid.empty();
    const meta = v.settings.archiveMeta || {};
    const files = v.archiveFiles().filter((f) => !this.q || (f.name + ' ' + ((meta[f.path] || {}).note || '')).toLowerCase().includes(this.q)).slice(0, 120);
    if (!files.length) this.grid.createDiv({ cls: 'ata-muted', text: '附件区还没有文件。' });
    files.forEach((f) => {
      const c = this.grid.createDiv('ata-apick-item' + (this.picked.has(f.path) ? ' is-selected' : ''));
      v.renderThumb(c.createDiv('ata-athumb'), f);
      c.createDiv({ cls: 'ata-apick-name', text: f.name });
      c.addEventListener('click', () => {
        if (this.picked.has(f.path)) this.picked.delete(f.path); else this.picked.add(f.path);
        c.toggleClass('is-selected', this.picked.has(f.path));
        this.countEl.setText(`已选 ${this.picked.size} 个`);
      });
    });
  }
  onClose() { this.contentEl.empty(); }
}

class ConfirmModal extends Modal {
  /** buttons: [{ text, cta, warning, onClick }] */
  constructor(app, title, message, buttons) { super(app); this.t = title; this.msg = message; this.buttons = buttons; }
  onOpen() {
    this.titleEl.setText(this.t);
    String(this.msg).split('\n').forEach((line) => this.contentEl.createEl('p', { text: line }));
    const bar = this.contentEl.createDiv('ata-confirm-bar');
    for (const b of this.buttons) {
      const el = bar.createEl('button', { text: b.text, cls: b.warning ? 'mod-warning' : b.cta ? 'mod-cta' : '' });
      el.addEventListener('click', () => { this.close(); b.onClick && b.onClick(); });
    }
    bar.createEl('button', { text: '取消' }).addEventListener('click', () => this.close());
  }
  onClose() { this.contentEl.empty(); }
}


class PathModal extends Modal {
  constructor(app, view) { super(app); this.view = view; }
  onOpen() {
    this.titleEl.setText('按路径导入');
    this.contentEl.createEl('p', { cls: 'setting-item-description', text: '把文件或文件夹的完整路径粘贴到下面，一行一个。在资源管理器里选中文件后按 Ctrl+Shift+C（复制文件地址）就能得到路径；微信里的文件可以右键「在文件夹中显示」后复制路径。' });
    const ta = this.contentEl.createEl('textarea', { cls: 'ata-path-input', attr: { rows: '5', placeholder: '"D:\\文档\\订单明细.xlsx"\nD:\\资料\\价格表文件夹' } });
    setTimeout(() => ta.focus(), 30);
    new Setting(this.contentEl).addButton((b) => b.setButtonText('读取').setCta().onClick(async () => {
      const paths = parsePaths(ta.value);
      if (!paths.length) { new Notice('没有识别到路径，请粘贴以盘符开头的完整路径，例如 D:\\文档\\表格.xlsx'); return; }
      this.close();
      await this.view.receive(await filesFromPaths(paths));
    }));
  }
  onClose() { this.contentEl.empty(); }
}

class HeaderRowModal extends Modal {
  constructor(app, table, onSubmit) { super(app); this.table = table; this.onSubmit = onSubmit; }
  onOpen() {
    this.titleEl.setText('设置表头行');
    const t = this.table;
    this.contentEl.createEl('p', { cls: 'setting-item-description', text: 'Excel 里表头上方常有「XX 年 XX 月对账单」之类的标题行。选中真正的表头所在的那一行，它上面的内容会保存为表格说明（AI 处理时也能看到）。' });
    const list = this.contentEl.createDiv('ata-hr-list');
    t.rows.slice(0, 12).forEach((r, i) => {
      const row = list.createDiv('ata-hr-row');
      row.createSpan({ cls: 'ata-hr-num', text: `第 ${i + 1} 行` });
      row.createSpan({ cls: 'ata-hr-text', text: r.filter((c) => !isBlank(c)).slice(0, 8).join(' | ') || '（空）' });
      row.addEventListener('click', () => { this.close(); this.onSubmit(i + 1); });
    });
  }
  onClose() { this.contentEl.empty(); }
}


/* ====================================================================
 * 四、工作台面板
 * ==================================================================== */
const fmtTime = (ms) => {
  const d = new Date(ms), now = Date.now(), diff = (now - ms) / 1000;
  if (diff < 60) return '刚刚';
  if (diff < 3600) return Math.floor(diff / 60) + ' 分钟前';
  if (diff < 86400) return Math.floor(diff / 3600) + ' 小时前';
  if (diff < 86400 * 7) return Math.floor(diff / 86400) + ' 天前';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

const CAT_COLORS = ['#10b981', '#3b82f6', '#f97316', '#8b5cf6', '#ec4899', '#14b8a6', '#eab308', '#ef4444', '#6366f1', '#84cc16'];

class TableWorkbenchView extends ItemView {
  constructor(leaf, plugin) {
    super(leaf);
    this.plugin = plugin;
    this.tab = 'home';
    this.libFilter = '全部';
    this.libSort = 'recent';
    this.selectMode = false;
    this.selected = new Set();
    this.query = '';
    this.stats = new Map(); // path -> {mtime, headers, rows}
    this.table = { headers: [], rows: [] };
    this.file = null;
    this.undoStack = [];
    this.undoMap = new Map();   // 每张表各自的撤销记录
    this.importing = null;
  }
  getViewType() { return VIEW_TYPE; }
  getDisplayText() { return '表格工作台'; }
  getIcon() { return 'table'; }

  get settings() { return this.plugin.settings; }

  async onOpen() {
    this.root = this.contentEl;
    this.root.empty();
    this.root.addClass('ata-root');

    /* 整个面板都可以拖入文件 / 文件夹 */
    let depth = 0;
    const zoneOf = (e) => {
      const el = e.target && e.target.closest ? e.target : null;
      if (!el) return null;
      if (el.closest('.ata-archive')) return 'archive';
      if (el.closest('.ata-aipanel')) return 'ai';
      const card = el.closest('[data-drop-table]');
      if (card) return 'table:' + card.getAttribute('data-drop-table');
      if (el.closest('.ata-inbox')) return 'inbox';
      if (el.closest('.ata-dropzone')) return 'import';
      return null;
    };
    const on = (type, fn) => {
      this.root.addEventListener(type, fn, true);
      this.register(() => this.root.removeEventListener(type, fn, true));
    };
    on('dragenter', (e) => {
      if (!this.wantsDrag(e)) return;
      e.preventDefault();
      depth++;
      this.root.addClass('ata-file-over');
    });
    on('dragleave', (e) => {
      if (!this.wantsDrag(e)) return;
      depth = Math.max(0, depth - 1);
      if (!depth) {
        this.root.removeClass('ata-file-over');
        if (this.overCard) { this.overCard.removeClass('is-drop-over'); this.overCard = null; }
      }
    });
    on('dragover', (e) => {
      if (!this.wantsDrag(e)) return;
      e.preventDefault();
      e.stopPropagation();
      e.dataTransfer.dropEffect = 'copy';
      const z = zoneOf(e);
      this.root.toggleClass('ata-over-master', z === 'inbox');
      this.root.toggleClass('ata-over-chat', z === 'ai');
      const card = e.target && e.target.closest ? e.target.closest('[data-drop-table]') : null;
      if (card !== this.overCard) {
        if (this.overCard) this.overCard.removeClass('is-drop-over');
        if (card) card.addClass('is-drop-over');
        this.overCard = card;
      }
    });
    on('drop', (e) => {
      if (!this.wantsDrag(e)) return;
      const dt = e.dataTransfer;
      const types = Array.from(dt.types || []);
      let text = '';
      try { text = dt.getData('text/uri-list') || dt.getData('text/plain') || ''; } catch (err) { /* 读不到 */ }
      const paths = parsePaths(text);
      if (!types.includes('Files') && !paths.length) return; // 普通文字，交给默认行为
      e.preventDefault();
      e.stopPropagation();
      depth = 0;
      this.root.removeClass('ata-file-over');
      this.root.removeClass('ata-over-master');
      this.root.removeClass('ata-over-chat');
      if (this.overCard) { this.overCard.removeClass('is-drop-over'); this.overCard = null; }
      const zone = zoneOf(e);
      console.log('[AI表格助手] drop', { types, files: dt.files ? dt.files.length : 0, text: text.slice(0, 300) });
      collectDropped(dt)
        .then(async (list) => {
          const real = list.filter((x) => x.file.size > 0 || TABLE_EXT.test(x.file.name));
          if (!real.length && paths.length) list = await filesFromPaths(paths);
          if (!list.length) {
            new Notice('没有收到文件内容。如果是从微信拖的：在微信里右键文件 →「在文件夹中显示」，再从文件夹拖进来；或者在这里按 Ctrl+V 粘贴文件 / 路径。', 12000);
            return;
          }
          await this.handleDrop(list, zone);
        })
        .catch((err) => new Notice('读取拖入的文件失败：' + (err.message || err), 8000));
    });

    /* Ctrl+V：粘贴复制的文件或文件路径 */
    this.registerDomEvent(document, 'paste', (e) => {
      if (!this.root.isConnected || this.root.offsetParent === null) return;
      const ws = this.app.workspace;
      const active = (ws.getActiveViewOfType && ws.getActiveViewOfType(ItemView)) || (ws.activeLeaf && ws.activeLeaf.view);
      if (active !== this && !this.root.contains(document.activeElement)) return;
      const tgt = e.target;
      const editable = tgt && (tgt.closest && (tgt.closest('textarea, input, [contenteditable="true"]')));
      const cd = e.clipboardData;
      const files = Array.from((cd && cd.files) || []);
      let paths = [];
      if (!files.length) {
        paths = parsePaths(cd ? cd.getData('text/plain') : '');
        if (!paths.length && !editable) paths = clipboardFilePaths();
      }
      if (!files.length && !paths.length) return;
      if (editable && !files.length && parsePaths(cd.getData('text/plain')).length === 0) return;
      e.preventDefault();
      e.stopPropagation();
      if (this.tab === 'archive') {
        if (files.length) this.archiveSave(files, '粘贴');
        else filesFromPaths(paths).then((list) => this.archiveSave(list.map((x) => x.file).filter((f) => f.size > 0), '粘贴')).catch((err) => new Notice(err.message, 8000));
        return;
      }
      if (files.length) this.receive(files.map((f) => ({ file: f, rel: f.name })));
      else filesFromPaths(paths).then((list) => this.receive(list)).catch((err) => new Notice(err.message, 8000));
    }, true);

    /* 附件区上传（任何文件） */
    this.archiveInput = this.makeInput({ multiple: 'true' }, (files) => this.archiveSave(files, '上传'));

    /* 隐藏的文件选择器 */
    this.fileInput = this.makeInput({ accept: '.csv,.tsv,.xlsx,.xlsm,.md', multiple: 'true' }, (files) =>
      this.importBatch(files.map((f) => ({ file: f, rel: f.name }))));
    this.folderInput = this.makeInput({ webkitdirectory: 'true', multiple: 'true' }, (files) =>
      this.importBatch(files.map((f) => ({ file: f, rel: f.webkitRelativePath || f.name }))));
    /* 首页「收到单据」：任何文件都先存进附件区，再选要填的表 */
    this.attachInput = this.makeInput({ multiple: 'true' }, (files) => this.receiveAtts(files.map((f) => ({ file: f, rel: f.name })), null));
    /* 工作区右侧「交给外部 AI」上传附件 */
    this.aiInput = this.makeInput({ multiple: 'true' }, async (files) => {
      const path = this.aiUploadFor || (this.file && this.file.path);
      const saved = await this.archiveSave(files, '上传', true);
      if (path && saved.length) { this.addAiAtts(path, saved); new Notice(`已存入附件区并加到附件：${saved.map((f) => f.name).join('、')}`); }
      this.renderAiPanel();
    });

    /* 监视表格文件：外部（AI 或其它程序）改了，就对比出改动，标出来让用户确认或撤销 */
    this.registerEvent(this.app.vault.on('modify', (f) => {
      if (this.snapReady && this.isTableFile(f)) this.checkFile(f, 'modify');
      this.onVaultChange(f);
    }));
    this.registerEvent(this.app.vault.on('create', (f) => {
      if (this.snapReady && !this.importRun && this.isTableFile(f)) setTimeout(() => this.checkFile(f, 'create'), 300);
      this.onVaultChange(f);
    }));
    this.registerEvent(this.app.vault.on('delete', (f) => {
      if (this.file && f.path === this.file.path) { this.file = null; this.table = { headers: [], rows: [] }; }
      this.onExtDelete(f);
      this.onVaultChange(f);
    }));
    this.registerEvent(this.app.vault.on('rename', (f, old) => { this.onExtRename(f, old); this.scheduleRender(); }));

    const last = this.settings.lastTable && this.app.vault.getAbstractFileByPath(this.settings.lastTable);
    if (last instanceof TFile) await this.loadFile(last, { silent: true, switchTab: false });
    await this.render();
    this.register(() => this.endTour(false));
    if (!this.settings.onboarded) setTimeout(() => this.startTour(), 600);
    /* 找出工作台关着时被外部改过的表，并更新给 AI 看的说明文件 */
    setTimeout(async () => {
      try { await this.scanExternal(); } catch (e) { console.warn('[AI表格助手] 检查外部修改失败', e); this.snapReady = true; }
      try { await this.writeAiGuide(); } catch (e) { console.warn('[AI表格助手] 写说明文件失败', e); }
    }, 300);
  }

  /* ================= 新手引导 ================= */
  tourSteps() {
    const toTab = (t) => async () => { if (this.tab !== t) { this.tab = t; await this.render(); } };
    const ws = async () => { if (this.file) await toTab('workspace')(); };
    return [
      { kind: 'welcome' },
      { sel: '.ata-tabs', title: '四个页面', before: toTab('home'),
        text: '「首页」看最近发生了什么，包括外部 AI 改了哪些表；「表格库」放所有表格（多工作表的 Excel 会归在一起）；「工作区」打开一张表查看和修改；「附件区」存收到的单据、截图、PDF。' },
      { sel: '.ata-top-actions', title: '第一步：导入你的表格', before: toTab('home'),
        text: '点「导入文件」或「导入文件夹」，把平时要填的 Excel / CSV 导进来。导入后每张表都是库里的一个 CSV 文件，外部 AI 可以直接读写。从微信拖不进来时，复制文件后在这里按 Ctrl+V，或点「按路径」。' },
      { sel: '.ata-homestrip .ata-inbox', title: '收到单据就丢这里', before: toTab('home'),
        text: '订单、送货单、价格表的截图或 PDF 拖到这里（或点一下选择）。文件会存进附件区，然后选要填哪张表，打开后附件已经放好。' },
      { sel: '.ata-qstrip', title: '常用表一键切换', before: toTab('home'), optional: true,
        text: '最近用过和设为常用的表排在这里，点一下就切过去。也可以把单据直接拖到某个表名上。' },
      { sel: '.ata-oplog', title: '操作记录', before: toTab('home'),
        text: '外部 AI 改了什么、你手动改了什么、导入导出和删除，都按时间记在这里。外部修改可以点「查看明细」看每个格子原来是什么、改成了什么，再确认或撤销。' },
      { sel: '.ata-aipanel', title: '交给外部 AI', before: ws, needs: () => !!this.file,
        alt: { sel: '.ata-tabs', text: '打开任意一张表后，「工作区」右边是「交给外部 AI」面板：写要求、加附件，点「复制给 AI 的指令」，粘贴给能读写电脑文件的 AI（例如 Claude 桌面版、Claude Code、Codex）。' },
        text: '写下要求（或点上面的常用要求），加上单据附件，点「复制给 AI 的指令」，粘贴给能读写电脑文件的 AI，例如 Claude 桌面版连上这个文件夹、Claude Code、Codex。指令里已经写好了表格和附件的完整路径，以及修改规则。' },
      { sel: '.ata-tablepane', title: 'AI 改完会自动显示', before: ws, needs: () => !!this.file,
        text: 'AI 改完文件，这里会立刻刷新：绿色是新增的行，黄色是改过的格子，上方出现提示条，可以「确认」或「撤销」。你也可以直接在表格里手动改。' },
      { sel: '.ata-export-btn', title: '导出成 Excel', before: ws, needs: () => !!this.file,
        text: '改好后点「导出」。从 Excel 导入的表会按原文件的格式导出：颜色、公式、合并单元格、图片都保留，多个工作表一起导出成一个文件。' },
      { sel: '.ata-archive-drop', title: '附件区', before: toTab('archive'),
        text: '收到的文件都存在这里，图片和 PDF 有预览，图片能一键复制。点卡片上的「交给 AI」，选一张表就能把它作为附件交给外部 AI。' },
      { kind: 'done' },
    ];
  }

  async startTour(i) {
    this.endTour(false);
    this.tourIdx = i || 0;
    const ov = this.tourEl = document.body.createDiv('ata-tour');
    this.tourHole = ov.createDiv('ata-tour-hole');
    this.tourCard = ov.createDiv('ata-tour-card');
    ov.addEventListener('click', (e) => { if (e.target === ov) e.stopPropagation(); });
    this.tourKey = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); this.endTour(true); }
      else if (e.key === 'ArrowRight' && !/INPUT|SELECT|TEXTAREA/.test((document.activeElement || {}).tagName)) this.tourGo(1);
      else if (e.key === 'ArrowLeft' && !/INPUT|SELECT|TEXTAREA/.test((document.activeElement || {}).tagName)) this.tourGo(-1);
    };
    this.tourResize = () => this.tourPlace();
    document.addEventListener('keydown', this.tourKey, true);
    window.addEventListener('resize', this.tourResize);
    await this.tourShow();
  }

  async tourGo(d) {
    const steps = this.tourSteps();
    let i = this.tourIdx + d;
    while (i > 0 && i < steps.length - 1 && steps[i].optional && !this.root.querySelector(steps[i].sel)) i += d;
    if (i < 0) i = 0;
    if (i >= steps.length) { this.endTour(true); return; }
    this.tourIdx = i;
    await this.tourShow();
  }

  async tourShow() {
    if (!this.tourEl) return;
    const steps = this.tourSteps();
    let st = steps[this.tourIdx];
    if (st.before) await st.before();
    if (st.needs && !st.needs()) st = Object.assign({}, st, st.alt || {});
    const card = this.tourCard;
    card.empty();
    card.removeClass('is-center');
    card.removeClass('is-wide');
    const n = steps.length - 2;
    const close = card.createEl('button', { cls: 'ata-tour-x', attr: { 'aria-label': '关闭引导' } });
    setIcon(close, 'x');
    close.addEventListener('click', () => this.endTour(true));

    if (st.kind === 'welcome') this.tourWelcome(card);
    else if (st.kind === 'done') this.tourDone(card);
    else {
      card.createDiv({ cls: 'ata-tour-step', text: `${this.tourIdx} / ${n}` });
      card.createDiv({ cls: 'ata-tour-title', text: st.title });
      card.createDiv({ cls: 'ata-tour-text', text: st.text });
      const bar = card.createDiv('ata-tour-bar');
      const skip = bar.createEl('button', { cls: 'ata-tour-skip', text: '跳过引导' });
      skip.addEventListener('click', () => this.endTour(true));
      const sp = bar.createDiv('ata-tour-btns');
      const prev = sp.createEl('button', { text: '上一步' });
      prev.addEventListener('click', () => this.tourGo(-1));
      const next = sp.createEl('button', { cls: 'mod-cta', text: '下一步' });
      next.addEventListener('click', () => this.tourGo(1));
      setTimeout(() => next.focus(), 30);
    }
    this.tourTarget = st.sel ? this.root.querySelector(st.sel) : null;
    if (this.tourTarget) {
      this.tourTarget.scrollIntoView({ block: 'nearest' });
      await new Promise((r) => setTimeout(r, 120));
    }
    this.tourPlace();
  }

  tourPlace() {
    if (!this.tourEl) return;
    const card = this.tourCard, hole = this.tourHole, t = this.tourTarget;
    const vw = window.innerWidth, vh = window.innerHeight;
    if (!t || !t.isConnected) {
      hole.style.display = 'none';
      this.tourEl.addClass('is-dim');
      card.addClass('is-center');
      card.style.left = card.style.top = '';
      return;
    }
    this.tourEl.removeClass('is-dim');
    const r = t.getBoundingClientRect(), pad = 6;
    const top = Math.max(4, r.top - pad), left = Math.max(4, r.left - pad);
    const h = Math.min(r.bottom + pad, vh - 4) - top, w = Math.min(r.right + pad, vw - 4) - left;
    Object.assign(hole.style, { display: 'block', top: top + 'px', left: left + 'px', width: w + 'px', height: Math.max(h, 20) + 'px' });
    const cw = card.offsetWidth || 360, chh = card.offsetHeight || 200;
    if (h > vh * 0.5 && (left - cw - 16 > 8 || left + w + cw + 16 < vw)) { // 很高的目标：卡片放在它旁边
      card.style.left = (left - cw - 16 > 8 ? left - cw - 16 : left + w + 16) + 'px';
      card.style.top = Math.max(8, Math.min(top + 40, vh - chh - 8)) + 'px';
      return;
    }
    if (h > vh * 0.5) { // 旁边放不下：盖在目标中间偏下，不挡住开头
      card.style.left = Math.max(8, left + w / 2 - cw / 2) + 'px';
      card.style.top = Math.max(8, Math.min(top + h * 0.45, vh - chh - 8)) + 'px';
      return;
    }
    let cy = top + h + 12;
    if (cy + chh > vh - 8) cy = top - chh - 12;
    if (cy < 8) cy = Math.min(vh - chh - 8, Math.max(8, top + 16)); // 目标太高时盖在目标里面
    let cx = Math.min(Math.max(8, left + Math.min(w, 400) / 2 - cw / 2), vw - cw - 8);
    if (cy > top && cy < top + h && left + w + cw + 20 < vw) { cx = left + w + 12; cy = Math.max(8, Math.min(top, vh - chh - 8)); }
    card.style.left = cx + 'px';
    card.style.top = cy + 'px';
  }

  tourWelcome(card) {
    card.addClass('is-wide');
    card.createDiv({ cls: 'ata-tour-title big', text: '欢迎使用表格工作台 👋' });
    card.createDiv({ cls: 'ata-tour-text', text: 'Obsidian 负责存放和展示表格；识别单据、填表这些活交给你自己用的 AI（Claude、ChatGPT 等）。AI 直接修改库里的表格文件，改了哪里这里会标出来，你确认或撤销，最后按原格式导出 Excel。' });
    const flow = card.createDiv('ata-tour-flow');
    [['folder-down', '导入表格'], ['inbox', '存进单据'], ['clipboard-copy', '复制指令给 AI'], ['bot', 'AI 改文件'], ['check', '确认 / 撤销'], ['download', '导出 Excel']].forEach(([ic, tx], i) => {
      if (i) flow.createSpan({ cls: 'ata-tour-arrow', text: '→' });
      const c = flow.createDiv('ata-tour-flowitem');
      setIcon(c.createDiv('ata-tour-flowicon'), ic);
      c.createSpan({ text: tx });
    });
    const box = card.createDiv('ata-tour-setup');
    box.createDiv({ cls: 'ata-tour-sub', text: '需要什么样的 AI？' });
    box.createDiv({ cls: 'ata-tour-hint', text: '能直接读写你电脑上文件的 AI 工具，例如：Claude 桌面版（把下面这个文件夹连给它）、Claude Code、Codex 等。网页版聊天读不到你电脑上的文件。' });
    const row = box.createDiv('ata-tour-row');
    row.createSpan({ text: '表格文件夹' });
    const path = this.absPath(normalizePath(this.settings.folder));
    const pv = row.createEl('code', { cls: 'ata-tour-path', text: path });
    pv.addEventListener('click', () => this.copyText(path, '文件夹路径'));
    box.createDiv({ cls: 'ata-tour-hint', text: '插件会在这个文件夹里放一份「AI使用说明.md」和「表格目录.md」，AI 读了就知道表格在哪、怎么改。点上面的路径可以复制。' });
    const bar = card.createDiv('ata-tour-bar');
    const skip = bar.createEl('button', { cls: 'ata-tour-skip', text: '跳过，我自己摸索' });
    skip.addEventListener('click', () => this.endTour(true));
    const go = bar.createDiv('ata-tour-btns').createEl('button', { cls: 'mod-cta', text: '开始 1 分钟导览 →' });
    go.addEventListener('click', () => this.tourGo(1));
  }

  tourDone(card) {
    card.addClass('is-wide');
    card.createDiv({ cls: 'ata-tour-title big', text: '可以开始了 🎉' });
    card.createDiv({ cls: 'ata-tour-text', text: '建议先导入一张你平时要填的表，再把一张单据交给 AI 试试。也可以先用示例表练手：' });
    const ul = card.createEl('ul', { cls: 'ata-tour-list' });
    ['示例表里有两条送货记录，右边已经写好一句示例要求', '点「复制给 AI 的指令」，粘贴给你的 AI', 'AI 改完文件，新增的行会标成绿色，点「确认」或「撤销」'].forEach((x) => ul.createEl('li', { text: x }));
    card.createDiv({ cls: 'ata-tour-hint', text: '以后想再看这个引导：点顶部的「引导」按钮，或在插件设置里点「重新查看使用引导」。' });
    const bar = card.createDiv('ata-tour-bar');
    const sp = bar.createDiv('ata-tour-btns');
    const demo = sp.createEl('button', { text: '用示例表试试' });
    demo.addEventListener('click', async () => { this.endTour(true); await this.createDemoTable(); });
    const done = sp.createEl('button', { cls: 'mod-cta', text: '完成' });
    done.addEventListener('click', () => this.endTour(true));
    setTimeout(() => done.focus(), 30);
  }

  async createDemoTable() {
    const d = new Date(), z = (n) => String(n).padStart(2, '0');
    const day = (k) => { const x = new Date(d.getTime() - k * 86400000); return `${x.getFullYear()}-${z(x.getMonth() + 1)}-${z(x.getDate())}`; };
    const t = { headers: ['日期', '客户', '品名', '规格', '数量(米)', '单价', '金额'], rows: [
      [day(2), '华盛布行', '棉布', '280克', '120', '8.5', '1020'],
      [day(1), '东方纺织', '涤纶布', '200克', '300', '6.2', '1860'],
    ] };
    try {
      await this.ensureFolder(this.settings.folder);
      const path = await this.uniquePath(this.settings.folder, '示例-送货记录', 'csv');
      await this.snapPut(path, toCSV(t));
      const file = await this.app.vault.create(path, toCSV(t));
      this.logOp({ type: 'create', path: file.path, text: '新建示例表（新手引导）' });
      this.aiState(file.path).ask = `帮我加一条：${day(0)} 华盛布行 又送了棉布 280克 200米，单价 8.5，金额自己算好。`;
      await this.openTable(file);
      new Notice('示例表已经建好，右边点「复制给 AI 的指令」试试。');
    } catch (e) { new Notice('创建示例表失败：' + (e.message || e)); }
  }

  endTour(done) {
    if (this.tourKey) document.removeEventListener('keydown', this.tourKey, true);
    if (this.tourResize) window.removeEventListener('resize', this.tourResize);
    this.tourKey = this.tourResize = null;
    if (this.tourEl) { this.tourEl.remove(); this.tourEl = null; }
    if (done && !this.settings.onboarded) { this.settings.onboarded = true; this.plugin.saveSettings(); }
    if (done && this.tab === 'archive' && this.tourIdx > 0) { this.tab = 'home'; this.render(); }
  }

  hasFiles(e) {
    return e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
  }

  /** 接受文件拖入；也接受只带路径文字的拖入（微信等），但不抢输入框里的普通文字拖放 */
  wantsDrag(e) {
    const types = Array.from((e.dataTransfer && e.dataTransfer.types) || []);
    if (types.includes('Files') || types.includes('text/uri-list')) return true;
    if (types.includes('text/plain')) {
      const t = e.target;
      return !(t && t.closest && t.closest('textarea, input, [contenteditable="true"]'));
    }
    return false;
  }

  /** 粘贴 / 按路径收到的文件：全是表格或整个文件夹就导入；有单据、截图就存进附件区交给 AI */
  async receive(list) {
    list = (list || []).filter(Boolean);
    if (!list.length) { new Notice('没有读到文件。'); return; }
    const isFolder = list.some((x) => x.rel.includes('/'));
    if (isFolder || list.every((x) => TABLE_EXT.test(x.file.name))) { await this.importBatch(list); return; }
    await this.receiveAtts(list, this.tab === 'workspace' && this.file ? this.file : null);
  }

  /** 存进附件区，加到某张表的附件里（没指定表就让用户选） */
  async receiveAtts(list, target) {
    const saved = await this.archiveSave(list.map((x) => x.file || x).filter((f) => f && f.size > 0), '收到', true);
    if (!saved.length) { new Notice('没有读到文件内容。'); return; }
    if (target && this.file && target.path === this.file.path && this.tab === 'workspace') {
      this.addAiAtts(target.path, saved);
      this.renderAiPanel();
      new Notice(`已存入附件区，并加到这张表的附件：${saved.map((f) => f.name).join('、')}`);
      return;
    }
    await this.handoff(saved, target);
  }

  makeInput(attrs, onFiles) {
    const el = this.root.createEl('input', { type: 'file', cls: 'ata-hidden', attr: attrs });
    el.addEventListener('change', () => {
      const files = Array.from(el.files || []);
      el.value = '';
      if (files.length) onFiles(files);
    });
    return el;
  }

  onVaultChange(f) {
    if (this.importRun) return; // 导入期间不重绘，结束后统一刷新
    if (f && f.path && f.path.startsWith(this.archiveDir() + '/')) { if (this.tab === 'archive' && !this.archiving) this.scheduleRender(); return; }
    if (!(f instanceof TFile) || !/^(csv|md)$/.test(f.extension)) return;
    if (this.selfWrite && this.file && f.path === this.file.path) return;
    if (this.tab !== 'workspace') this.scheduleRender();
  }

  scheduleRender() {
    clearTimeout(this.renderTimer);
    this.renderTimer = setTimeout(() => this.render(), 400);
  }

  /* ---------- 表格库数据 ---------- */
  tableFiles() {
    return this.app.vault.getFiles().filter((f) => this.isTableFile(f));
  }

  async statOf(f) {
    const c = this.stats.get(f.path);
    if (c && c.mtime === f.stat.mtime) return c;
    let t = { headers: [], rows: [] }, rows = 0;
    try {
      if (f.extension === 'md') {
        const m = findMdTable(await this.app.vault.cachedRead(f));
        if (m) t = m.table;
        rows = t.rows.length;
      } else if (f.stat.size > 2e6) {
        /* 大文件：只解析开头一段拿表头，行数按换行符估算 */
        const text = await this.app.vault.read(f);
        let n = 0, p = -1;
        while ((p = text.indexOf('\n', p + 1)) >= 0) n++;
        if (!text.endsWith('\n')) n++;
        const head = text.slice(0, 65536);
        const cut = head.lastIndexOf('\n');
        t = gridToTable(parseCSV(cut > 0 ? head.slice(0, cut) : head));
        rows = Math.max(0, n - 1);
      } else {
        t = gridToTable(parseCSV(await this.app.vault.cachedRead(f)));
        rows = t.rows.length;
      }
    } catch (e) { /* 读取失败当空表 */ }
    const s = { mtime: f.stat.mtime, headers: t.headers, rows, sample: t.rows.slice(0, 3) };
    this.stats.set(f.path, s);
    return s;
  }

  async library() {
    const files = this.tableFiles();
    const meta = this.settings.meta || {};
    const out = [];
    for (const f of files) {
      const s = await this.statOf(f);
      out.push({ file: f, ...s, category: (meta[f.path] && meta[f.path].category) || '' });
    }
    return out;
  }

  categoryColor(cat) {
    if (!cat) return 'var(--text-faint)';
    const all = Object.values(this.settings.meta || {}).map((m) => m.category).filter(Boolean);
    const uniq = [...new Set(all)].sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'));
    const i = uniq.indexOf(cat);
    return CAT_COLORS[(i < 0 ? 0 : i) % CAT_COLORS.length];
  }

  async setCategory(path, cat) {
    const meta = this.settings.meta || (this.settings.meta = {});
    if (cat) meta[path] = Object.assign({}, meta[path], { category: cat });
    else if (meta[path]) delete meta[path].category;
    if (meta[path] && !Object.keys(meta[path]).length) delete meta[path];
    await this.plugin.saveSettings();
  }

  /* ---------- 渲染 ---------- */
  async render() {
    if (!this.root) return;
    const scroll = this.root.scrollTop;
    const lib = await this.library();
    this.lib = lib;
    this.root.querySelectorAll(':scope > :not(input):not(.ata-importbar)').forEach((el) => el.remove());

    const banner = this.root.createDiv('ata-drop-banner');
    setIcon(banner.createSpan('ata-icon'), 'folder-down');
    banner.createSpan({ text: '松开即可导入' });
    banner.createSpan({ cls: 'ata-muted', text: this.tab === 'workspace' ? '拖到右侧对话区 = 作为当前表的附件；拖到上方表名标签 = 填进那张表；其它地方 = 帮你选表' : '拖到某张表的卡片上 = 直接填那张表；其它地方 = 帮你找该填的表；文件夹 = 导入表格库' });

    this.renderTopbar(lib);
    this.renderTabs(lib);
    const body = this.root.createDiv('ata-body');
    this.renderedTab = this.tab;
    if (this.tab === 'library') this.renderLibrary(body, lib);
    else if (this.tab === 'workspace') this.renderWorkspace(body);
    else if (this.tab === 'archive') this.renderArchive(body);
    else this.renderHome(body, lib);
    if (this.tab !== 'workspace') this.root.scrollTop = scroll;
  }

  ico(parent, name, cls) {
    const s = parent.createSpan(cls || 'ata-icon');
    setIcon(s, name);
    return s;
  }

  btn(parent, text, icon, onClick, cls) {
    const b = parent.createEl('button', { cls: 'ata-btn ' + (cls || '') });
    if (icon) this.ico(b, icon);
    if (text) b.createSpan({ text });
    b.addEventListener('click', onClick);
    return b;
  }

  iconBtn(parent, icon, label, onClick) {
    const b = parent.createEl('button', { cls: 'ata-iconbtn', attr: { 'aria-label': label } });
    setIcon(b, icon);
    b.addEventListener('click', (e) => { e.stopPropagation(); onClick(e); });
    return b;
  }

  renderTopbar(lib) {
    const bar = this.root.createDiv('ata-topbar');
    const search = bar.createDiv('ata-search');
    this.ico(search, 'search', 'ata-search-icon');
    const input = search.createEl('input', { cls: 'ata-search-input', attr: { type: 'text', placeholder: '搜索表格名称、列名、分类…（回车打开第一个）' } });
    input.value = this.query;
    const drop = search.createDiv('ata-search-drop');
    const update = () => {
      this.query = input.value.trim();
      drop.empty();
      if (!this.query) { drop.removeClass('is-open'); return; }
      const hits = this.searchLib(lib, this.query).slice(0, 8);
      if (!hits.length) drop.createDiv({ cls: 'ata-search-empty', text: '没有匹配的表格' });
      hits.forEach((h) => {
        const row = drop.createDiv('ata-search-item');
        this.ico(row, 'table');
        const mid = row.createDiv('ata-row-mid');
        mid.createDiv({ cls: 'ata-row-title', text: h.file.basename });
        mid.createDiv({ cls: 'ata-row-sub', text: `${h.rows} 行 · ${h.headers.slice(0, 5).join('、')}` });
        if (h.category) row.createSpan({ cls: 'ata-type-chip', text: h.category });
        row.addEventListener('mousedown', (e) => { e.preventDefault(); this.openTable(h.file); });
      });
      drop.addClass('is-open');
    };
    input.addEventListener('input', update);
    input.addEventListener('focus', update);
    input.addEventListener('blur', () => setTimeout(() => drop.removeClass('is-open'), 150));
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        const h = this.searchLib(lib, input.value.trim())[0];
        if (h) this.openTable(h.file);
      } else if (e.key === 'Escape') { input.value = ''; update(); input.blur(); }
    });

    const acts = bar.createDiv('ata-top-actions');
    const top = (icon, text, fn) => {
      const b = acts.createEl('button', { cls: 'ata-top-btn' });
      this.ico(b, icon);
      b.createSpan({ text, cls: 'ata-top-label' });
      b.addEventListener('click', fn);
    };
    top('file-up', '导入文件', () => this.fileInput.click());
    top('folder-up', '导入文件夹', () => this.folderInput.click());
    top('clipboard-paste', '按路径', () => new PathModal(this.app, this).open());
    top('file-plus', '新建', () => new NewTableModal(this.app, (n, c, f) => this.createTable(n, c, f)).open());
    top('help-circle', '引导', () => this.startTour());
    top('settings', '设置', () => {
      this.app.setting.open();
      this.app.setting.openTabById(this.plugin.manifest.id);
    });
  }

  searchLib(lib, q) {
    if (!q) return [];
    q = q.toLowerCase();
    return lib.filter((t) => [t.file.path, t.category, ...t.headers].join(' ').toLowerCase().includes(q))
      .sort((a, b) => (a.file.basename.toLowerCase().includes(q) ? 0 : 1) - (b.file.basename.toLowerCase().includes(q) ? 0 : 1) || b.mtime - a.mtime);
  }

  renderTabs(lib) {
    const tabs = this.root.createDiv('ata-tabs');
    const tab = (id, icon, text, count) => {
      const b = tabs.createEl('button', { cls: 'ata-tab' + (this.tab === id ? ' is-active' : '') });
      this.ico(b, icon);
      b.createSpan({ text });
      if (count != null) b.createSpan({ cls: 'ata-tab-count', text: String(count) });
      b.addEventListener('click', () => { this.tab = id; this.render(); });
    };
    tab('home', 'layout-dashboard', '首页');
    tab('library', 'library', '表格库', lib.length);
    tab('workspace', 'table-2', this.file ? '工作区 · ' + this.file.basename : '工作区');
    tab('archive', 'archive', '附件区', this.archiveFiles().length);
    if (this.importing) tabs.createDiv({ cls: 'ata-importing', text: this.importing });
  }

  /* ---------- 首页 ---------- */
  renderHome(body, lib) {
    const head = body.createDiv('ata-header');
    const left = head.createDiv();
    const title = left.createDiv('ata-title');
    this.ico(title, 'table');
    title.createSpan({ text: '表格工作台' });
    const sp = this.splitLibrary(lib);
    const weekAgo = Date.now() - 7 * 86400000;
    left.createDiv({ cls: 'ata-sub', text: `${sp.singles.length + sp.books.length} 个表格${sp.books.length ? `（${sp.books.length} 个 Excel 工作簿）` : ''} · ${lib.reduce((a, t) => a + t.rows, 0).toLocaleString()} 行数据 · 本周改过 ${lib.filter((t) => t.mtime > weekAgo).length} 张 · 附件 ${this.archiveFiles().length} 个` });
    const right = head.createDiv('ata-header-right');
    this.btn(right, '新建表格', 'plus', () => new NewTableModal(this.app, (n, c, f) => this.createTable(n, c, f)).open(), 'mod-cta');

    /* 两个小入口：收到单据 / 导入表格 */
    const strip = body.createDiv('ata-homestrip');
    const ib = strip.createDiv('ata-inbox ata-mini');
    this.ico(ib, 'inbox', 'ata-dropzone-icon');
    const t1 = ib.createDiv('ata-dropzone-text');
    t1.createEl('b', { text: '收到单据 / 截图 / 价格表' });
    t1.createSpan({ text: '拖到这里或点击：存进附件区，选好要填的表，复制指令交给 AI' });
    ib.addEventListener('click', () => this.attachInput.click());
    const dz = strip.createDiv('ata-dropzone ata-mini');
    this.ico(dz, 'folder-down', 'ata-dropzone-icon');
    const t2 = dz.createDiv('ata-dropzone-text');
    t2.createEl('b', { text: '导入表格到表格库' });
    t2.createSpan({ text: '拖入文件或文件夹，或点击选择' });
    dz.addEventListener('click', () => this.fileInput.click());

    /* 常用和最近：一行小卡片，也可以把单据拖上去 */
    const files = this.recentFiles(10);
    if (files.length) {
      const row = body.createDiv('ata-qstrip');
      row.createSpan({ cls: 'ata-muted', text: '常用 / 最近：' });
      files.forEach((f) => {
        const m = (this.settings.meta || {})[f.path] || {};
        const c = row.createEl('button', { cls: 'ata-qchip' + (this.file && this.file.path === f.path ? ' is-current' : ''), attr: { title: (m.workbook ? m.workbook + ' · ' : '') + f.path + '\n可以把单据拖到这里直接填这张表' } });
        c.setAttribute('data-drop-table', f.path);
        if (this.isPinned(f.path)) c.createSpan({ cls: 'ata-qchip-star', text: '★' });
        c.createSpan({ text: m.sheet || f.basename });
        c.addEventListener('click', () => this.openTable(f));
      });
    }

    this.renderOpLog(body);
  }




  /** 工作区顶部：常用 + 最近，像浏览器标签一样切换 */
  renderSwitchBar(body) {
    const files = this.recentFiles(10);
    if (files.length < 2) return;
    const meta = this.settings.meta || {};
    const bar = body.createDiv('ata-switchbar');
    files.forEach((f) => {
      const m = meta[f.path] || {};
      const b = bar.createEl('button', { cls: 'ata-switch' + (this.file && f.path === this.file.path ? ' is-active' : '') });
      b.setAttribute('data-drop-table', f.path);
      if (this.isPinned(f.path)) b.createSpan({ cls: 'ata-star', text: '★' });
      b.createSpan({ text: m.sheet ? `${m.sheet}` : f.basename });
      if (m.workbook) b.title = `${m.workbook} · ${m.sheet}`;
      const pending = this.pendingFor(f.path).length;
      if (pending) b.createSpan({ cls: 'ata-tab-count is-ext', text: String(pending), attr: { title: '有待确认的外部修改' } });
      b.addEventListener('click', () => { if (!this.file || f.path !== this.file.path) this.openTable(f); });
    });
    const more = bar.createEl('button', { cls: 'ata-switch ata-switch-more', text: '更多…' });
    more.addEventListener('click', () => this.quickSwitch());
  }

  quickSwitch() {
    const recent = this.settings.recent || [];
    const files = this.tableFiles().sort((a, b) => {
      const ia = recent.indexOf(a.path), ib = recent.indexOf(b.path);
      return (ia < 0 ? 1e9 : ia) - (ib < 0 ? 1e9 : ib) || b.stat.mtime - a.stat.mtime;
    });
    new TablePickModal(this.app, files, (f) => this.openTable(f), this.settings.meta).open();
  }

  section(parent, icon, title, right) {
    const sec = parent.createDiv('ata-section');
    const h = sec.createDiv('ata-section-head');
    const t = h.createDiv('ata-section-title');
    this.ico(t, icon);
    t.createSpan({ text: title });
    if (right) h.createDiv({ cls: 'ata-muted', text: right });
    return sec.createDiv('ata-section-body');
  }

  row(parent, t) {
    const r = parent.createDiv('ata-row');
    const ic = r.createSpan('ata-ficon');
    ic.style.setProperty('--fc', this.categoryColor(t.category));
    setIcon(ic, t.file.extension === 'md' ? 'file-text' : 'table');
    const mid = r.createDiv('ata-row-mid');
    const rm = (this.settings.meta || {})[t.file.path] || {};
    mid.createDiv({ cls: 'ata-row-title', text: rm.sheet ? `${rm.sheet}  ·  📗 ${rm.workbook}` : t.file.basename });
    const sub = mid.createDiv('ata-row-sub');
    if (t.category) sub.createSpan({ text: t.category });
    sub.createSpan({ text: `${t.rows} 行 × ${t.headers.length} 列` });
    sub.createSpan({ text: t.file.parent && t.file.parent.path !== '/' ? t.file.parent.path : '' });
    r.createDiv({ cls: 'ata-row-right', text: fmtTime(t.mtime) });
    r.addEventListener('click', () => this.openTable(t.file));
  }

  renderDropzone(parent) {
    const dz = parent.createDiv('ata-dropzone');
    this.ico(dz, 'folder-down', 'ata-dropzone-icon');
    const txt = dz.createDiv('ata-dropzone-text');
    txt.createEl('b', { text: '把表格文件或整个文件夹导入表格库' });
    txt.createSpan({ text: `支持 csv / tsv / xlsx / md，文件夹会连同子文件夹一起导入，保存在「${this.settings.folder}」下并保留原目录结构。也可以点这里选择文件夹，或在工作台里按 Ctrl+V 粘贴复制的文件 / 文件路径。` });
    dz.addEventListener('click', () => this.folderInput.click());
  }

  /* ---------- 表格库 ---------- */
  renderLibrary(body, lib) {
    const head = body.createDiv('ata-header');
    const left = head.createDiv();
    const title = left.createDiv('ata-title');
    this.ico(title, 'library');
    title.createSpan({ text: '全部表格' });
    left.createDiv({ cls: 'ata-sub', text: '点击卡片进入工作区编辑；点分类标签可手动修改分类。' });
    const right = head.createDiv('ata-header-right');
    const seg = right.createDiv('ata-seg');
    [['recent', '最近'], ['name', '名称'], ['rows', '行数']].forEach(([k, label]) => {
      const b = seg.createEl('button', { cls: 'ata-seg-btn' + (this.libSort === k ? ' is-active' : ''), text: label });
      b.addEventListener('click', () => { this.libSort = k; this.render(); });
    });
    this.btn(right, this.selectMode ? '完成' : '批量管理', this.selectMode ? 'check' : 'list-checks', () => {
      this.selectMode = !this.selectMode;
      this.selected = new Set();
      this.render();
    });
    this.btn(right, '导入文件夹', 'folder-up', () => this.folderInput.click(), 'mod-cta');

    if (!lib.length) {
      this.renderDropzone(body);
      return;
    }

    /* 多工作表的 Excel 单独成组，不和普通表格混在一起 */
    const { books, singles } = this.splitLibrary(lib);
    const allLib = lib;
    lib = singles;

    const counts = new Map();
    lib.forEach((t) => counts.set(t.category || '未分类', (counts.get(t.category || '未分类') || 0) + 1));
    const catbar = body.createDiv('ata-catbar');
    const chip = (name, n) => {
      const b = catbar.createEl('button', { cls: 'ata-catchip' + (this.libFilter === name ? ' is-active' : '') });
      if (name !== '全部' && name !== '未分类') {
        const d = b.createSpan('ata-dot');
        d.style.setProperty('--pc', this.categoryColor(name));
      }
      b.createSpan({ text: name });
      b.createSpan({ cls: 'ata-tab-count', text: String(n) });
      b.addEventListener('click', () => { this.libFilter = name; this.render(); });
    };
    chip('全部', lib.length);
    [...counts.entries()].filter(([k]) => k !== '未分类').sort((a, b) => b[1] - a[1]).forEach(([k, n]) => chip(k, n));
    if (counts.has('未分类')) chip('未分类', counts.get('未分类'));
    if (this.libFilter !== '全部' && !counts.has(this.libFilter)) this.libFilter = '全部';

    let list = lib.filter((t) => this.libFilter === '全部' || (t.category || '未分类') === this.libFilter);
    const bookList = books.filter((b) => this.libFilter === '全部' || b.sheets.some((t) => (t.category || '未分类') === this.libFilter));
    if (this.selectMode) {
      const sb = body.createDiv('ata-selbar');
      this.selCountEl = sb.createSpan({ cls: 'ata-selcount', text: `已选 ${this.selected.size} 个` });
      this.btn(sb, '全选当前列表', 'check-square', () => { list.forEach((t) => this.selected.add(t.file.path)); bookList.forEach((b) => b.sheets.forEach((t) => this.selected.add(t.file.path))); this.render(); }, 'ata-btn-sm');
      this.btn(sb, '清空选择', 'square', () => { this.selected.clear(); this.render(); }, 'ata-btn-sm');
      this.btn(sb, '设为常用', 'star', async () => {
        const pin = this.settings.pinned || (this.settings.pinned = []);
        this.selected.forEach((p) => { if (!pin.includes(p)) pin.push(p); });
        await this.plugin.saveSettings();
        this.render();
      }, 'ata-btn-sm');
      this.btn(sb, '删除所选', 'trash-2', () => {
        const files = [...this.selected].map((p) => this.app.vault.getAbstractFileByPath(p)).filter((f) => f instanceof TFile);
        if (!files.length) { new Notice('还没有选择表格。'); return; }
        new ConfirmModal(this.app, `删除选中的 ${files.length} 个表格？`, files.slice(0, 8).map((f) => '• ' + f.path).join('\n') + (files.length > 8 ? `\n……共 ${files.length} 个` : '') + '\n会移到回收站，需要时可以恢复。',
          [{ text: `删除 ${files.length} 个`, warning: true, onClick: () => this.deleteFiles(files) }]).open();
      }, 'ata-btn-sm is-danger');
    }
    const sorters = {
      recent: (a, b) => b.mtime - a.mtime,
      name: (a, b) => a.file.basename.localeCompare(b.file.basename, 'zh-Hans-CN', { numeric: true }),
      rows: (a, b) => b.rows - a.rows,
    };
    list = list.sort(sorters[this.libSort]);

    if (bookList.length) {
      const bs = body.createDiv('ata-section');
      const bh = bs.createDiv('ata-section-head');
      const bt = bh.createDiv('ata-section-title');
      this.ico(bt, 'book-open');
      bt.createSpan({ text: `Excel 工作簿（${bookList.length}）` });
      bh.createDiv({ cls: 'ata-muted', text: '一个 Excel 里有多个工作表；点开后在表格下方切换工作表，改完可以按原格式整本导出' });
      const bg = bs.createDiv('ata-grid');
      const bsort = { recent: (a, b) => b.mtime - a.mtime, name: (a, b) => a.name.localeCompare(b.name, 'zh-Hans-CN', { numeric: true }), rows: (a, b) => b.rows - a.rows };
      bookList.sort(bsort[this.libSort]).forEach((b) => this.bookCard(bg, b));
    }
    const ss = body.createDiv('ata-section');
    if (bookList.length) {
      const sh = ss.createDiv('ata-section-head');
      const st = sh.createDiv('ata-section-title');
      this.ico(st, 'table');
      st.createSpan({ text: `单独的表格（${list.length}）` });
    }
    const grid = ss.createDiv('ata-grid');
    if (!list.length) grid.createDiv({ cls: 'ata-empty-sm', text: '没有单独的表格。' });
    list.forEach((t) => this.card(grid, t));
    void allLib;
  }

  /** 把表格库分成「多工作表的 Excel」和「单独的表格」 */
  splitLibrary(lib) {
    const meta = this.settings.meta || {};
    const groups = new Map();
    const singles = [];
    for (const t of lib) {
      const m = meta[t.file.path];
      if (m && m.workbook) {
        const key = (m.link && m.link.id) || ((t.file.parent ? t.file.parent.path : '') + '|' + m.workbook);
        if (!groups.has(key)) groups.set(key, { key, name: m.workbook, sheets: [], linkId: m.link && m.link.id });
        groups.get(key).sheets.push(t);
      } else singles.push(t);
    }
    const books = [];
    for (const g of groups.values()) {
      if (g.sheets.length < 2) { singles.push(...g.sheets); continue; }
      g.sheets.sort((a, b) => ((meta[a.file.path] || {}).sheetIndex || 0) - ((meta[b.file.path] || {}).sheetIndex || 0));
      g.rows = g.sheets.reduce((a, t) => a + t.rows, 0);
      g.mtime = Math.max(...g.sheets.map((t) => t.mtime));
      g.folder = g.sheets[0].file.parent ? g.sheets[0].file.parent.path : '';
      books.push(g);
    }
    return { books, singles };
  }

  bookCard(grid, b) {
    const meta = this.settings.meta || {};
    const recent = this.settings.recent || [];
    const c = grid.createDiv('ata-card ata-pcard ata-bookcard' + (this.file && b.sheets.some((t) => t.file.path === this.file.path) ? ' is-current' : ''));
    c.style.setProperty('--pc', '#16a34a');
    const top = c.createDiv('ata-card-top');
    const ic = top.createSpan('ata-ficon');
    ic.style.setProperty('--fc', '#16a34a');
    setIcon(ic, 'file-spreadsheet');
    top.createDiv({ cls: 'ata-card-title', text: b.name });
    top.createSpan({ cls: 'ata-pchip', text: `${b.sheets.length} 个工作表` }).style.setProperty('--pc', '#16a34a');
    c.createDiv({ cls: 'ata-card-desc', text: b.linkId ? '已保存原 Excel，可按原格式（颜色、公式、图片）导出' : b.folder });
    const chips = c.createDiv('ata-chips');
    b.sheets.slice(0, 10).forEach((t) => {
      const ch = chips.createEl('button', { cls: 'ata-colchip ata-sheetchip', text: (meta[t.file.path] || {}).sheet || t.file.basename });
      ch.setAttribute('data-drop-table', t.file.path);
      ch.addEventListener('click', (e) => { e.stopPropagation(); this.openTable(t.file); });
    });
    if (b.sheets.length > 10) chips.createSpan({ cls: 'ata-colchip is-more', text: `+${b.sheets.length - 10}` });
    const metaEl = c.createDiv('ata-card-meta');
    const m = (icon, text) => { const sp = metaEl.createSpan('ata-meta'); this.ico(sp, icon); sp.createSpan({ text }); };
    m('rows-3', `${b.rows.toLocaleString()} 行`);
    const acts = metaEl.createDiv('ata-card-acts');
    this.iconBtn(acts, 'download', '导出整个工作簿', () => this.exportDialog(b.sheets[0].file));
    this.iconBtn(acts, 'trash-2', '删除整个工作簿', () => new ConfirmModal(this.app, `删除「${b.name}」？`, `会删除这个 Excel 拆出来的全部 ${b.sheets.length} 个工作表（移到回收站，可以恢复）。`,
      [{ text: `删除 ${b.sheets.length} 个工作表`, warning: true, onClick: () => this.deleteFiles(b.sheets.map((t) => t.file)) }]).open()).addClass('is-danger');
    metaEl.createSpan({ cls: 'ata-meta-time', text: fmtTime(b.mtime) });
    const open = () => {
      const last = recent.find((p) => b.sheets.some((t) => t.file.path === p));
      const t = (last && b.sheets.find((x) => x.file.path === last)) || b.sheets[0];
      this.openTable(t.file);
    };
    if (this.selectMode) {
      c.addClass('is-selectable');
      const all = b.sheets.every((t) => this.selected.has(t.file.path));
      if (all) c.addClass('is-selected');
      const box = c.createEl('input', { type: 'checkbox', cls: 'ata-card-check' });
      box.checked = all;
      const toggle = (on) => { b.sheets.forEach((t) => { if (on) this.selected.add(t.file.path); else this.selected.delete(t.file.path); }); c.toggleClass('is-selected', on); box.checked = on; if (this.selCountEl) this.selCountEl.setText(`已选 ${this.selected.size} 个`); };
      box.addEventListener('click', (e) => e.stopPropagation());
      box.addEventListener('change', () => toggle(box.checked));
      c.addEventListener('click', () => toggle(!b.sheets.every((t) => this.selected.has(t.file.path))));
      return;
    }
    c.addEventListener('click', open);
  }

  card(grid, t) {
    const c = grid.createDiv('ata-card ata-pcard' + (this.file && this.file.path === t.file.path ? ' is-current' : ''));
    c.setAttribute('data-drop-table', t.file.path);
    c.style.setProperty('--pc', this.categoryColor(t.category));
    const top = c.createDiv('ata-card-top');
    const ic = top.createSpan('ata-ficon');
    ic.style.setProperty('--fc', this.categoryColor(t.category));
    setIcon(ic, t.file.extension === 'md' ? 'file-text' : 'table');
    const tt = top.createDiv({ cls: 'ata-card-title', text: t.file.basename });
    if (this.isPinned(t.file.path)) tt.prepend(createSpan({ cls: 'ata-star', text: '★ ' }));
    const cat = top.createSpan({ cls: 'ata-pchip is-clickable', text: t.category || '未分类' });
    if (!t.category) cat.addClass('is-empty');
    cat.addEventListener('click', (e) => {
      e.stopPropagation();
      const existing = [...new Set(this.lib.map((x) => x.category).filter(Boolean))];
      new CategoryModal(this.app, t.category, existing, async (v) => { await this.setCategory(t.file.path, v); this.render(); }).open();
    });

    const tm = (this.settings.meta || {})[t.file.path];
    c.createDiv({ cls: 'ata-card-desc', text: tm && tm.workbook ? `📗 ${tm.workbook} · 工作表「${tm.sheet}」` : t.file.parent && t.file.parent.path !== '/' ? t.file.parent.path : '库根目录' });
    const chips = c.createDiv('ata-chips');
    t.headers.slice(0, 6).forEach((h) => chips.createSpan({ cls: 'ata-colchip', text: h }));
    if (t.headers.length > 6) chips.createSpan({ cls: 'ata-colchip is-more', text: `+${t.headers.length - 6}` });
    if (!t.headers.length) chips.createSpan({ cls: 'ata-colchip is-more', text: '空表' });

    const meta = c.createDiv('ata-card-meta');
    const m = (icon, text) => { const s = meta.createSpan('ata-meta'); this.ico(s, icon); s.createSpan({ text }); };
    m('rows-3', `${t.rows} 行`);
    m('columns-3', `${t.headers.length} 列`);
    const acts = meta.createDiv('ata-card-acts');
    this.iconBtn(acts, 'star', this.isPinned(t.file.path) ? '取消常用' : '设为常用', () => this.togglePin(t.file));
    this.iconBtn(acts, 'download', '导出', () => this.exportDialog(t.file));
    this.iconBtn(acts, 'trash-2', '删除', () => this.confirmDelete(t.file)).addClass('is-danger');
    this.iconBtn(acts, 'file-input', '在 Obsidian 中打开原文件', () => this.app.workspace.getLeaf('tab').openFile(t.file));
    meta.createSpan({ cls: 'ata-meta-time', text: fmtTime(t.mtime) });
    if (this.selectMode) {
      c.addClass('is-selectable');
      if (this.selected.has(t.file.path)) c.addClass('is-selected');
      const box = c.createEl('input', { type: 'checkbox', cls: 'ata-card-check' });
      box.checked = this.selected.has(t.file.path);
      box.addEventListener('click', (e) => e.stopPropagation());
      box.addEventListener('change', () => this.toggleSelect(t.file.path, c, box.checked));
      c.addEventListener('click', () => this.toggleSelect(t.file.path, c, !this.selected.has(t.file.path), box));
      return;
    }
    c.addEventListener('click', () => this.openTable(t.file));
  }

  toggleSelect(path, card, on, box) {
    if (on) this.selected.add(path); else this.selected.delete(path);
    card.toggleClass('is-selected', on);
    if (box) box.checked = on;
    if (this.selCountEl) this.selCountEl.setText(`已选 ${this.selected.size} 个`);
  }

  /* ---------- 工作区 ---------- */
  renderWorkspace(body) {
    if (!this.file) {
      const e = body.createDiv('ata-empty');
      this.ico(e, 'table-2', 'ata-dropzone-icon');
      e.createDiv({ text: '还没有打开表格。从表格库里选一个，或者新建 / 导入一个。' });
      const r = e.createDiv('ata-header-right');
      this.btn(r, '去表格库', 'library', () => { this.tab = 'library'; this.render(); }, 'mod-cta');
      this.btn(r, '新建表格', 'plus', () => new NewTableModal(this.app, (n, c, f) => this.createTable(n, c, f)).open());
      return;
    }
    const meta = (this.settings.meta || {})[this.file.path];
    const cat = meta && meta.category;

    this.renderSwitchBar(body);
    const head = body.createDiv('ata-header');
    const left = head.createDiv();
    const title = left.createDiv('ata-title');
    this.ico(title, 'table-2');
    title.createSpan({ text: this.file.basename });
    const chip = title.createSpan({ cls: 'ata-pchip is-clickable', text: cat || '未分类' });
    chip.style.setProperty('--pc', this.categoryColor(cat));
    chip.addEventListener('click', () => {
      const existing = [...new Set((this.lib || []).map((x) => x.category).filter(Boolean))];
      new CategoryModal(this.app, cat, existing, async (v) => { await this.setCategory(this.file.path, v); this.render(); }).open();
    });
    left.createDiv({ cls: 'ata-sub', text: this.file.path });
    const right = head.createDiv('ata-header-right');
    const pinned = this.isPinned(this.file.path);
    const mb = this.btn(right, pinned ? '常用' : '设为常用', 'star', () => this.togglePin(this.file));
    if (pinned) mb.addClass('is-master');
    this.btn(right, '切换表格', 'arrow-left-right', () => this.quickSwitch());
    this.undoBtn = this.btn(right, '撤销', 'undo-2', () => this.undo());
    this.undoBtn.disabled = !this.undoStack.length;
    this.btn(right, '原文件', 'file-input', () => this.app.workspace.getLeaf('tab').openFile(this.file));
    this.btn(right, '删除', 'trash-2', () => this.confirmDelete(this.file), 'is-danger');
    this.btn(right, '导出', 'download', () => this.exportDialog(), 'mod-cta ata-export-btn');

    const ws = this.wsEl = body.createDiv('ata-ws');
    requestAnimationFrame(() => this.fitWorkspace());

    /* 左：表格（工作表标签在表格下方，和 Excel 一样） */
    const tablePane = ws.createDiv('ata-tablepane');
    this.reviewEl = tablePane.createDiv('ata-review');
    this.renderReview();
    this.previewEl = tablePane.createDiv('ata-tablewrap');
    this.renderSheetTabs(tablePane);
    const foot = tablePane.createDiv('ata-tablefoot');
    this.btn(foot, '新增行', 'plus', () => this.addRow(), 'ata-btn-sm');
    this.btn(foot, '新增列', 'columns-3', () => this.addCol(), 'ata-btn-sm');
    this.btn(foot, '设表头行', 'heading', () => this.pickHeaderRow(), 'ata-btn-sm').title = '表头不对时（上面有标题行），把某一行设为表头';
    this.statEl = foot.createDiv({ cls: 'ata-muted ata-spacer-left' });
    this.renderTable();

    /* 右：交给外部 AI */
    this.aiPanelEl = ws.createDiv('ata-chatpanel ata-aipanel');
    this.renderAiPanel();
  }

  /** 工作区高度按窗口实际剩余空间计算，保证表格底部的工作表标签和按钮栏都看得到 */
  fitWorkspace() {
    const ws = this.wsEl;
    if (!ws || !ws.isConnected || this.tab !== 'workspace') return;
    if (window.getComputedStyle(ws).gridTemplateColumns.split(' ').length < 2) { ws.style.height = ''; return; } // 窄屏上下排列时不限高
    const top = ws.getBoundingClientRect().top - this.root.getBoundingClientRect().top + this.root.scrollTop;
    const pad = parseFloat(window.getComputedStyle(this.root).paddingBottom) || 0;
    ws.style.height = Math.max(380, this.root.clientHeight - top - Math.min(pad, 24) - 4) + 'px';
  }

  onResize() { this.fitWorkspace(); }

  /** 同一个 Excel 拆出来的工作表（按原顺序） */
  workbookSiblings(file) {
    const meta = this.settings.meta || {};
    const m = file && meta[file.path];
    if (!m || !m.workbook) return [];
    const id = m.link && m.link.id;
    const folder = file.parent ? file.parent.path : '';
    return this.app.vault.getFiles()
      .filter((f) => {
        const x = meta[f.path];
        if (f.extension !== 'csv' || !x) return false;
        return id ? x.link && x.link.id === id : f.parent && f.parent.path === folder && x.workbook === m.workbook;
      })
      .sort((a, b) => (meta[a.path].sheetIndex || 0) - (meta[b.path].sheetIndex || 0));
  }

  /** 像 Excel 一样放在表格下方的工作表标签 */
  renderSheetTabs(parent) {
    const sibs = this.workbookSiblings(this.file);
    if (sibs.length < 2) return;
    const meta = this.settings.meta || {};
    const bar = parent.createDiv('ata-sheettabs');
    const lab = bar.createSpan({ cls: 'ata-sheettabs-label', attr: { title: meta[this.file.path].workbook } });
    this.ico(lab, 'file-spreadsheet');
    lab.createSpan({ text: meta[this.file.path].workbook });
    const tabs = bar.createDiv('ata-sheettabs-list');
    sibs.forEach((f) => {
      const b = tabs.createEl('button', { cls: 'ata-sheettab' + (f.path === this.file.path ? ' is-active' : ''), text: meta[f.path].sheet || f.basename });
      b.setAttribute('data-drop-table', f.path);
      if (this.isPinned(f.path)) b.prepend(createSpan({ cls: 'ata-star', text: '★' }));
      b.addEventListener('click', () => { if (f.path !== this.file.path) this.openTable(f); });
    });
    setTimeout(() => { const a = tabs.querySelector('.is-active'); if (a) a.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }, 0);
  }

  /* ---------- 和原 Excel 的对应关系（保存在表格库的隐藏文件夹 .ata 里） ---------- */
  ataDir() { return normalizePath(this.settings.folder) + '/.ata'; }

  /* ================= 外部修改：快照、对比、确认、撤销 ================= */
  isTableFile(f) {
    if (!(f instanceof TFile)) return false;
    const folder = normalizePath(this.settings.folder) + '/';
    if (f.path.startsWith(this.archiveDir() + '/') || f.path.startsWith(folder + '导出/')) return false;
    if (f.path.startsWith(folder) && GUIDE_FILES.includes(f.path.slice(folder.length))) return false;
    if (f.extension === 'csv') return true;
    if (f.extension === 'md') {
      const meta = this.settings.meta || {};
      if (!(f.path.startsWith(folder) || meta[f.path])) return false;
      const cache = this.app.metadataCache.getFileCache(f);
      return !cache || !cache.sections || cache.sections.some((s) => s.type === 'table');
    }
    return false;
  }

  snapDir() { return this.ataDir() + '/snap'; }

  async snapIndex() {
    if (this.snapIdx) return this.snapIdx;
    try { this.snapIdx = JSON.parse(await this.app.vault.adapter.read(this.snapDir() + '/index.json')); this.snapFresh = false; }
    catch (e) { this.snapIdx = {}; this.snapFresh = true; }
    if (!this.snapText) this.snapText = new Map();
    return this.snapIdx;
  }

  saveSnapIndex() {
    clearTimeout(this.snapIdxTimer);
    this.snapIdxTimer = setTimeout(async () => {
      try { await this.ensureAdapterDir(this.snapDir()); await this.app.vault.adapter.write(this.snapDir() + '/index.json', JSON.stringify(this.snapIdx)); }
      catch (e) { console.warn('[AI表格助手] 保存快照索引失败', e); }
    }, 500);
  }

  /** 记下「插件自己写进去的内容」：之后文件变成别的样子，就是外部改的 */
  async snapPut(path, text, stat) {
    const idx = await this.snapIndex();
    const e = idx[path] || (idx[path] = { k: hashStr(path + '|' + Date.now()).slice(0, 12) });
    e.h = hashStr(text);
    e.n = this.tableLabel(path);
    if (stat) { e.m = stat.mtime; e.s = stat.size; } else { e.m = 0; e.s = -1; }
    this.snapText.set(path, text);
    try {
      await this.ensureAdapterDir(this.snapDir());
      await this.app.vault.adapter.write(`${this.snapDir()}/${e.k}.txt`, text);
    } catch (err) { console.warn('[AI表格助手] 保存快照失败', err); }
    this.saveSnapIndex();
  }

  async snapGet(path) {
    const idx = await this.snapIndex();
    if (this.snapText.has(path)) return this.snapText.get(path);
    const e = idx[path];
    if (!e) return null;
    try { const t = await this.app.vault.adapter.read(`${this.snapDir()}/${e.k}.txt`); this.snapText.set(path, t); return t; }
    catch (err) { return null; }
  }

  async snapDrop(path) {
    const idx = await this.snapIndex();
    const e = idx[path];
    if (!e) return;
    delete idx[path];
    this.snapText.delete(path);
    try { await this.app.vault.adapter.remove(`${this.snapDir()}/${e.k}.txt`); } catch (err) { /* 已经没有了 */ }
    this.saveSnapIndex();
  }

  parseTableText(f, text) {
    if (f.extension === 'md') { const m = findMdTable(text); return m ? m.table : { headers: [], rows: [] }; }
    return gridToTable(parseCSV(text));
  }

  /** 每个文件的检查排队进行，避免同一个文件连着改两次时乱序 */
  queueCheck(path, fn) {
    if (!this.checkQ) this.checkQ = new Map();
    const prev = this.checkQ.get(path) || Promise.resolve();
    const next = prev.then(fn).catch((e) => console.warn('[AI表格助手] 检查外部修改失败', e));
    this.checkQ.set(path, next);
    next.then(() => { if (this.checkQ.get(path) === next) this.checkQ.delete(path); });
    return next;
  }

  async checkFile(f, why) {
    if (!this.isTableFile(f)) return;
    await this.queueCheck(f.path, async () => {
      const idx = await this.snapIndex();
      const text = await this.app.vault.read(f);
      const e = idx[f.path];
      const old = e ? await this.snapGet(f.path) : null;
      if (old == null) {
        await this.snapPut(f.path, text, f.stat);
        if (why === 'create' || why === 'scan-new') {
          const t = this.parseTableText(f, text);
          this.logOp({ type: 'ext-new', path: f.path, text: `外部新建了表格（${t.rows.length} 行，列：${t.headers.slice(0, 8).join('、')}${t.headers.length > 8 ? '…' : ''}）`, review: 'ok' });
          this.scheduleIndex();
        }
        return;
      }
      if (old === text) { e.m = f.stat.mtime; e.s = f.stat.size; this.saveSnapIndex(); return; }
      await this.handleExternal(f, old, text);
    });
  }

  async handleExternal(f, oldText, newText) {
    const oldT = this.parseTableText(f, oldText), newT = this.parseTableText(f, newText);
    const d = diffTables(oldT, newT);
    const nothing = !d.cells.length && !d.added.length && !d.removed.length && !d.addedCols.length && !d.removedCols.length && !d.renamed.length;
    const meta = (this.settings.meta || {})[f.path];
    let link = null, oldRows = null, oldCols = null;
    if (meta && meta.link) {
      link = await this.readLink(meta.link.id, meta.link.sheet);
      if (link) {
        oldRows = (link.map.rows || []).slice();
        oldCols = link.map.cols ? link.map.cols.slice() : null;
        link.map.rows = d.rowMap.map((i) => (i >= 0 && i < oldRows.length ? oldRows[i] : -1));
        if (oldCols && oldCols.length === oldT.headers.length) link.map.cols = d.colMap.map((i) => (i >= 0 ? oldCols[i] : -1));
        await this.writeLinkMap(link);
      }
    }
    let hist = '';
    if (!nothing) {
      try {
        const k = ((await this.snapIndex())[f.path] || {}).k || hashStr(f.path).slice(0, 12);
        const dir = `${this.ataDir()}/history/${k}`;
        await this.ensureAdapterDir(dir);
        hist = `${dir}/${Date.now()}.json`;
        await this.app.vault.adapter.write(hist, JSON.stringify({ path: f.path, text: oldText, rows: oldRows, cols: oldCols }));
        const all = (await this.app.vault.adapter.list(dir)).files.sort();
        for (const old of all.slice(0, Math.max(0, all.length - 40))) { try { await this.app.vault.adapter.remove(old); } catch (e) { /* 忽略 */ } }
      } catch (e) { console.warn('[AI表格助手] 保存修改前版本失败', e); hist = ''; }
    }
    await this.snapPut(f.path, newText, f.stat);

    /* 正在看的就是这张表：重新读进来，标出改动 */
    if (this.file && this.file.path === f.path) {
      clearTimeout(this.saveTimer);
      this.saveTimer = null;
      this.table = newT;
      if (link) { this.link = link; this.attachProvenance(this.table, link.map); }
      this.tableVersion = (this.tableVersion || 0) + 1;
    }
    if (nothing) { if (this.file && this.file.path === f.path && this.tab === 'workspace') this.renderTable(); return; }

    const hdr = newT.headers.map(String);
    const hl = { added: new Set(d.added), cells: new Map() };
    d.cells.forEach((x) => { if (!hl.cells.has(x.r)) hl.cells.set(x.r, new Set()); hl.cells.get(x.r).add(x.c); });
    d.addedCols.forEach((j) => newT.rows.forEach((r, ri) => { if (!isBlank(r[j]) && !hl.added.has(ri)) { if (!hl.cells.has(ri)) hl.cells.set(ri, new Set()); hl.cells.get(ri).add(j); } }));
    if (!this.hlMap) this.hlMap = new Map();
    this.hlMap.set(f.path, hl);

    const parts = [];
    if (d.added.length) parts.push(`新增 ${d.added.length} 行`);
    if (d.cells.length) parts.push(`修改 ${d.cells.length} 个单元格`);
    if (d.removed.length) parts.push(`删除 ${d.removed.length} 行`);
    const items = [
      ...d.renamed.map(([o, n]) => `列「${o}」改名为「${n}」`),
      ...d.addedCols.map((j, k) => `新增列「${hdr[j]}」（${d.filled[k]} 行有内容）`),
      ...d.removedCols.map((h) => `删除了列「${h}」`),
    ];
    if (!parts.length && items.length) parts.push(items[0]);
    this.logOp({
      type: 'ext', path: f.path, text: parts.join('，'), review: 'pending', hist, after: hashStr(newText),
      cells: d.cells.slice(0, 300).map((x) => ({ r: x.r + 1, h: hdr[x.c], o: x.o, n: x.n })), cellsN: d.cells.length,
      hdr, added: d.added.slice(0, 30).map((r) => newT.rows[r].map(cellStr)), addedN: d.added.length,
      oldHdr: oldT.headers.map(String), rows: d.removed.slice(0, 30).map((r) => oldT.rows[r].map(cellStr)), rowsN: d.removed.length,
      items,
    });
    this.scheduleIndex();
    new Notice(`「${this.tableLabel(f.path)}」被外部修改：${parts.join('，')}。在工作台里可以查看、确认或撤销。`, 8000);
    if (this.file && this.file.path === f.path && this.tab === 'workspace') { this.renderTable(); this.renderReview(); this.renderAiPanel(); }
    else this.scheduleRender();
  }

  pendingFor(path) {
    return (this.plugin.opLog || []).filter((e) => e.type === 'ext' && e.path === path && e.review === 'pending');
  }

  async confirmExternal(path) {
    this.pendingFor(path).forEach((e) => { e.review = 'ok'; });
    if (this.hlMap) this.hlMap.delete(path);
    this.saveOpLog();
    if (this.file && this.file.path === path && this.tab === 'workspace') { this.renderTable(); this.renderReview(); this.renderAiPanel(); }
    else this.render();
  }

  async undoExternal(e) {
    const f = this.app.vault.getAbstractFileByPath(e.path);
    if (!(f instanceof TFile)) { new Notice('这张表已经不在了，没法撤销。'); return; }
    if (!e.hist) { new Notice('没有保存修改前的版本，没法撤销。'); return; }
    let h;
    try { h = JSON.parse(await this.app.vault.adapter.read(e.hist)); } catch (err) { new Notice('找不到修改前的版本，没法撤销。'); return; }
    const cur = await this.app.vault.read(f);
    const go = async () => {
      await this.flushSave();
      await this.snapPut(f.path, h.text);
      await this.app.vault.modify(f, h.text);
      await this.snapPut(f.path, h.text, f.stat);
      const meta = (this.settings.meta || {})[f.path];
      let link = null;
      if (meta && meta.link && h.rows) {
        link = await this.readLink(meta.link.id, meta.link.sheet);
        if (link) { link.map.rows = h.rows; if (h.cols) link.map.cols = h.cols; await this.writeLinkMap(link); }
      }
      (this.plugin.opLog || []).forEach((x) => { if (x.type === 'ext' && x.path === f.path && x.review === 'pending' && x.t >= e.t) x.review = 'undone'; });
      e.review = 'undone';
      if (this.hlMap) this.hlMap.delete(f.path);
      this.logOp({ type: 'undo', path: f.path, text: `撤销了外部修改（${e.text}）` });
      if (this.file && this.file.path === f.path) {
        this.table = this.parseTableText(f, h.text);
        if (link) { this.link = link; this.attachProvenance(this.table, link.map); }
        this.tableVersion = (this.tableVersion || 0) + 1;
      }
      new Notice('已撤销这次外部修改，表格恢复到修改前。');
      await this.render();
    };
    if (hashStr(cur) !== e.after) {
      new ConfirmModal(this.app, '这张表之后又被改过', '撤销会把表格恢复到这次修改之前，之后的修改（包括你在工作台里的手动修改）也会一起退回。确定要撤销吗？', [{ text: '仍然撤销', warning: true, onClick: go }]).open();
    } else await go();
  }

  /** 打开工作台时：把关掉期间别人改过的表找出来 */
  async scanExternal() {
    if (this.scanning) return;
    this.scanning = true;
    try {
      const idx = await this.snapIndex();
      const fresh = this.snapFresh;
      const files = this.tableFiles();
      const seen = new Set();
      let n = 0;
      for (const f of files) {
        seen.add(f.path);
        const e = idx[f.path];
        if (e && e.m === f.stat.mtime && e.s === f.stat.size) continue;
        await this.checkFile(f, e || fresh ? 'scan' : 'scan-new');
        if (++n % 20 === 0) await new Promise((r) => setTimeout(r, 0));
      }
      for (const p of Object.keys(idx)) {
        if (seen.has(p) || this.app.vault.getAbstractFileByPath(p)) continue;
        this.logOp({ type: 'ext-del', path: p, name: idx[p].n, text: '表格文件被删除了（不是在工作台里删的）', review: 'ok' });
        await this.snapDrop(p);
      }
      this.snapFresh = false;
      this.saveSnapIndex();
    } finally {
      this.scanning = false;
      this.snapReady = true;
    }
  }

  onExtDelete(f) {
    if (!this.snapIdx || !this.snapIdx[f.path]) return;
    const name = this.snapIdx[f.path].n;
    if (!this.selfDeleting) this.logOp({ type: 'ext-del', path: f.path, name, text: '表格文件被删除了（不是在工作台里删的）', review: 'ok' });
    this.snapDrop(f.path);
    this.scheduleIndex();
  }

  onExtRename(f, old) {
    if (this.snapIdx && this.snapIdx[old]) {
      this.snapIdx[f.path] = this.snapIdx[old];
      delete this.snapIdx[old];
      if (this.snapText && this.snapText.has(old)) { this.snapText.set(f.path, this.snapText.get(old)); this.snapText.delete(old); }
      this.saveSnapIndex();
    }
    if (this.hlMap && this.hlMap.has(old)) { this.hlMap.set(f.path, this.hlMap.get(old)); this.hlMap.delete(old); }
    (this.plugin.opLog || []).forEach((e) => { if (e.path === old) e.path = f.path; });
    this.saveOpLog();
    this.scheduleIndex();
  }

  /** 工作区表格上方的提示条：外部刚改过这张表 */
  renderReview() {
    const el = this.reviewEl;
    if (!el || !this.file) return;
    el.empty();
    const pend = this.pendingFor(this.file.path);
    el.toggleClass('is-hidden', !pend.length);
    if (!pend.length) return;
    const last = pend[pend.length - 1];
    this.ico(el, 'bot', 'ata-review-icon');
    const tx = el.createDiv('ata-review-text');
    const tm = new Date(last.t);
    tx.createEl('b', { text: `外部修改了这张表${pend.length > 1 ? `（${pend.length} 次）` : ''}` });
    tx.createSpan({ text: ` · ${String(tm.getHours()).padStart(2, '0')}:${String(tm.getMinutes()).padStart(2, '0')} · ${last.text}。绿色是新增的行，黄色是改过的格子。` });
    const bar = el.createDiv('ata-review-btns');
    this.btn(bar, '查看明细', 'list', () => { this.tab = 'home'; this.opKind = '外部修改'; if (!this.opOpen) this.opOpen = new Set(); this.opOpen.add(last.t + '|' + last.type + '|' + (last.path || '')); this.render(); }, 'ata-btn-sm');
    this.btn(bar, '撤销', 'undo-2', () => this.undoExternal(last), 'ata-btn-sm');
    this.btn(bar, '确认', 'check', () => this.confirmExternal(this.file.path), 'ata-btn-sm mod-cta');
  }

  /* ================= 交给外部 AI ================= */
  basePath() {
    const a = this.app.vault.adapter;
    return a && typeof a.getBasePath === 'function' ? a.getBasePath() : '';
  }

  absPath(p) {
    const base = this.basePath();
    if (!base) return p;
    const sep = base.includes('\\') ? '\\' : '/';
    return base.replace(/[\\/]+$/, '') + sep + p.split('/').join(sep);
  }

  guidePath() { return normalizePath(this.settings.folder + '/AI使用说明.md'); }

  aiState(path) {
    if (!this.aiStates) this.aiStates = new Map();
    if (!this.aiStates.has(path)) this.aiStates.set(path, { ask: '', atts: [] });
    return this.aiStates.get(path);
  }

  addAiAtts(path, tfs) {
    const st = this.aiState(path);
    tfs.forEach((tf) => { if (tf && !st.atts.includes(tf.path)) st.atts.push(tf.path); });
  }

  buildInstruction(file, st) {
    const meta = (this.settings.meta || {})[file.path] || {};
    const stat = this.stats.get(file.path);
    const cols = stat ? stat.headers.slice(0, 30).join('、') : '';
    const lines = [
      '请帮我修改 Obsidian「表格工作台」里的一张表格，直接修改我电脑上的文件。',
      '',
      `1. 先阅读说明文件（里面有文件格式和修改规则）：${this.absPath(this.guidePath())}`,
      `2. 要修改的表格：${this.absPath(file.path)}`,
      `   （${meta.workbook ? `Excel「${meta.workbook}」里的工作表「${meta.sheet}」` : file.basename}${cols ? `；列：${cols}` : ''}${stat ? `；现在 ${stat.rows} 行` : ''}）`,
    ];
    const atts = st.atts.filter((p) => this.app.vault.getAbstractFileByPath(p));
    if (atts.length) {
      lines.push('3. 参考的附件：');
      atts.forEach((p) => lines.push(`   - ${this.absPath(p)}`));
    }
    lines.push(`${atts.length ? 4 : 3}. 要求：${st.ask.trim() || '（请按上面的说明处理附件里的内容）'}`);
    lines.push('', '改完后请简要告诉我新增、修改了哪些行。我会在 Obsidian 里查看改动，确认或撤销。');
    return lines.join('\n');
  }

  async copyText(text, what) {
    try { await navigator.clipboard.writeText(text); new Notice(`已复制${what || ''}，粘贴给你的 AI 就行。`); }
    catch (e) {
      try { require('electron').clipboard.writeText(text); new Notice(`已复制${what || ''}，粘贴给你的 AI 就行。`); }
      catch (err) { new Notice('复制失败：' + (e.message || e)); }
    }
  }

  /** 工作区右侧：把这张表交给外部 AI */
  renderAiPanel() {
    const el = this.aiPanelEl;
    if (!el || !el.isConnected || !this.file) return;
    el.empty();
    const file = this.file;
    const st = this.aiState(file.path);
    const head = el.createDiv('ata-chat-head');
    this.ico(head, 'bot');
    head.createSpan({ text: '交给外部 AI' });

    const body = el.createDiv('ata-aip');
    /* 1. 待确认的外部修改 */
    const pend = this.pendingFor(file.path);
    if (pend.length) {
      const box = body.createDiv('ata-aip-pending');
      box.createDiv({ cls: 'ata-aip-label', text: `待确认的外部修改（${pend.length}）` });
      pend.slice(-3).reverse().forEach((e) => {
        const tm = new Date(e.t);
        box.createDiv({ cls: 'ata-aip-pitem', text: `${String(tm.getHours()).padStart(2, '0')}:${String(tm.getMinutes()).padStart(2, '0')}  ${e.text}` });
      });
      const b = box.createDiv('ata-aip-row');
      this.btn(b, '撤销最近一次', 'undo-2', () => this.undoExternal(pend[pend.length - 1]), 'ata-btn-sm');
      this.btn(b, '全部确认', 'check', () => this.confirmExternal(file.path), 'ata-btn-sm mod-cta');
    }

    /* 2. 要 AI 做什么 */
    body.createDiv({ cls: 'ata-aip-label', text: '要 AI 做什么' });
    const chips = body.createDiv('ata-chips');
    const ta = body.createEl('textarea', { cls: 'ata-aip-ask', attr: { rows: '5', placeholder: '用中文写要求，比如：把附件里的送货单录入这张表，金额 = 数量 × 单价' } });
    ta.value = st.ask;
    ta.addEventListener('input', () => { st.ask = ta.value; });
    AI_TEMPLATES.forEach(([label, text]) => {
      chips.createEl('button', { cls: 'ata-catchip', text: label }).addEventListener('click', () => { st.ask = ta.value = text; ta.focus(); });
    });

    /* 3. 附件 */
    const ah = body.createDiv('ata-aip-label');
    ah.createSpan({ text: '附件（单据、截图、价格表）' });
    const atts = body.createDiv('ata-aip-atts');
    st.atts = st.atts.filter((p) => this.app.vault.getAbstractFileByPath(p));
    st.atts.forEach((p) => {
      const tf = this.app.vault.getAbstractFileByPath(p);
      const c = atts.createDiv('ata-aip-att');
      const th = c.createDiv('ata-athumb');
      this.renderThumb(th, tf);
      th.addEventListener('click', () => new LightboxModal(this.app, this, tf).open());
      c.createDiv({ cls: 'ata-aip-attname', text: tf.name, attr: { title: tf.path } });
      const x = c.createEl('button', { cls: 'ata-aip-x', text: '×', attr: { 'aria-label': '移除' } });
      x.addEventListener('click', () => { st.atts = st.atts.filter((q) => q !== p); this.renderAiPanel(); });
    });
    const ab = body.createDiv('ata-aip-row');
    this.btn(ab, '上传 / 拖入文件', 'upload', () => { this.aiUploadFor = file.path; this.aiInput.click(); }, 'ata-btn-sm');
    this.btn(ab, '从附件区选', 'archive', () => new ArchivePickModal(this.app, this, (tfs) => { this.addAiAtts(file.path, tfs); this.renderAiPanel(); }).open(), 'ata-btn-sm');

    /* 4. 复制指令 */
    const go = body.createDiv('ata-aip-go');
    this.btn(go, '复制给 AI 的指令', 'clipboard-copy', () => this.copyText(this.buildInstruction(file, st), '指令'), 'mod-cta');
    const more = body.createDiv('ata-aip-row');
    this.btn(more, '只复制表格路径', 'link', () => this.copyText(this.absPath(file.path), '表格路径'), 'ata-btn-sm');
    this.btn(more, '预览指令', 'eye', () => new TextPromptModal(this.app, '给 AI 的指令（可以复制）', this.buildInstruction(file, st), () => {}, true).open(), 'ata-btn-sm');

    const tip = body.createDiv('ata-aip-tip');
    tip.createDiv({ text: '把指令粘贴给能直接读写电脑文件的 AI（例如 Claude 桌面版连上这个文件夹、Claude Code、Codex 等），它会直接改这张表的文件。' });
    tip.createDiv({ text: `AI 需要能访问的文件夹：${this.absPath(normalizePath(this.settings.folder))}` });
    tip.createDiv({ text: '改动会自动出现在左边：绿色是新增，黄色是修改，可以确认或撤销。网页版聊天读不到你电脑上的文件。' });
  }

  /** 收到单据：先存进附件区，再选要填的表，打开后附件已经放好 */
  async handoff(tfs, target) {
    tfs = (tfs || []).filter(Boolean);
    const go = async (file) => {
      this.addAiAtts(file.path, tfs);
      const st = this.aiState(file.path);
      if (!st.ask.trim()) st.ask = AI_TEMPLATES[0][1];
      await this.openTable(file);
    };
    if (target instanceof TFile) { await go(target); return; }
    const files = this.recentFiles(200);
    const all = this.tableFiles();
    const rest = all.filter((f) => !files.includes(f));
    new TablePickModal(this.app, [...files, ...rest], (f) => go(f), this.settings.meta, tfs.length ? `已存入附件区：${tfs.map((x) => x.name).join('、')}。选择要填的表…` : null).open();
  }

  /* ================= 给 AI 看的说明文件和表格目录 ================= */
  async writeAiGuide() {
    const folder = normalizePath(this.settings.folder);
    await this.ensureFolder(folder);
    const abs = this.absPath(folder);
    const guide = `<!-- 由「AI 表格助手」插件自动生成，每次打开工作台都会更新，手动修改会被覆盖 -->
# 表格工作台 · 给 AI 的使用说明

你正在帮用户修改 Obsidian 里「表格工作台」管理的表格。表格就是这个文件夹里的普通文件，直接读写文件即可。用户会在 Obsidian 里看到你改了哪些地方（新增的行标绿色，改过的格子标黄色），并决定确认还是撤销。

## 文件夹结构

表格文件夹：\`${abs}\`

- \`表格目录.md\`：所有表格的清单（文件路径、来源、列名、行数、一行样例）。**不知道该改哪张表时先看它。**
- 各个 \`.csv\` 文件：表格本身。从多工作表 Excel 导入的表放在以 Excel 文件名命名的子文件夹里，每个工作表一个 csv。
- \`附件区/年-月/\`：用户存的单据、截图、PDF、价格表等参考资料。
- \`导出/\`：导出的 Excel，不要修改。
- \`.ata/\`：插件内部数据（原 Excel 副本、格式对应关系、修改记录），**不要读写这个文件夹**。

## 修改表格的规则

1. **CSV 格式**：UTF-8 编码；第一行是表头；字段里有逗号、引号或换行时，用英文双引号包起来，内部的双引号写成两个 \`""\`。
2. **不要改表头**（列名和列的顺序），除非用户明确要求。需要新列时加在最后。
3. **新增的行加在表格末尾**，不要插到中间，也不要调整已有行的顺序；用户要求删除时才删除行。
4. **只改需要改的行和格子**，其它内容保持原样（包括空格子、原来的写法）。
5. **格式和原表保持一致**：日期写成 \`YYYY-MM-DD\`（或照搬这一列已有的写法）；数字不要加千分位逗号和货币符号；金额等计算列请自己算好再填。
6. 从 Excel 导入的表，用户导出时会把你的修改按原 Excel 的颜色、公式、格式写回去，所以请保持行列结构，不要整体重写或重新排序。
7. Markdown 笔记里的表格：只修改其中的表格部分，表格外的文字不要动。
8. **不要删除、改名或移动表格文件。** 需要新表时，在表格文件夹里新建一个 csv 文件（第一行写表头）。
9. 用户只让你「检查」或「列出」时，不要修改文件。

## 改完以后

在回复里简要说明：改了哪张表、新增了几行、改了哪些格子、有哪些拿不准需要用户核对的地方。
`;
    const ptr = `<!-- 由「AI 表格助手」插件自动生成 -->\n# 说明\n\n这个文件夹是 Obsidian「表格工作台」的表格库。修改任何表格之前，请先阅读同一文件夹里的 \`AI使用说明.md\`，并按其中的规则操作；要找表格请看 \`表格目录.md\`。\n`;
    const put = async (name, text) => {
      const p = `${folder}/${name}`;
      const f = this.app.vault.getAbstractFileByPath(p);
      try {
        if (f instanceof TFile) { if ((await this.app.vault.read(f)) !== text) await this.app.vault.modify(f, text); }
        else await this.app.vault.create(p, text);
      } catch (e) { console.warn('[AI表格助手] 写说明文件失败', name, e); }
    };
    await put('AI使用说明.md', guide);
    await put('CLAUDE.md', ptr);
    await put('AGENTS.md', ptr);
    await this.writeTableIndex();
  }

  scheduleIndex() {
    clearTimeout(this.indexTimer);
    this.indexTimer = setTimeout(() => this.writeTableIndex(), 4000);
  }

  async writeTableIndex() {
    const folder = normalizePath(this.settings.folder);
    const lib = await this.library();
    const meta = this.settings.meta || {};
    const lines = ['<!-- 由「AI 表格助手」插件自动生成，表格有变化时自动更新，手动修改会被覆盖 -->', '# 表格目录', '', `共 ${lib.length} 张表格。路径相对于 Obsidian 库：\`${this.basePath() || '（库文件夹）'}\``, ''];
    const books = new Map();
    lib.forEach((t) => {
      const m = meta[t.file.path] || {};
      const key = m.workbook ? '工作簿：' + m.workbook : '单独的表格';
      if (!books.has(key)) books.set(key, []);
      books.get(key).push(t);
    });
    for (const [key, list] of books) {
      lines.push(`## ${key}`, '');
      list.sort((a, b) => ((meta[a.file.path] || {}).sheetIndex || 0) - ((meta[b.file.path] || {}).sheetIndex || 0) || a.file.path.localeCompare(b.file.path));
      for (const t of list) {
        const m = meta[t.file.path] || {};
        const d = new Date(t.mtime);
        lines.push(`### ${m.sheet || t.file.basename}`);
        lines.push(`- 文件：\`${t.file.path}\``);
        if (m.category) lines.push(`- 分类：${m.category}`);
        lines.push(`- 列（${t.headers.length}）：${t.headers.join('、')}`);
        lines.push(`- 行数：${t.rows}，最后修改：${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`);
        if (t.sample && t.sample.length) lines.push(`- 样例：${t.headers.map((h, i) => `${h}=${cellStr(t.sample[0][i])}`).filter((x) => !/=$/.test(x)).slice(0, 12).join('；')}`);
        lines.push('');
      }
    }
    const text = lines.join('\n');
    const p = `${folder}/表格目录.md`;
    const f = this.app.vault.getAbstractFileByPath(p);
    try {
      if (f instanceof TFile) { if ((await this.app.vault.read(f)).replace(/\r/g, '') !== text) await this.app.vault.modify(f, text); }
      else { await this.ensureFolder(folder); await this.app.vault.create(p, text); }
    } catch (e) { console.warn('[AI表格助手] 写表格目录失败', e); }
  }

  /* ================= 操作记录 ================= */
  async loadOpLog() {
    const pl = this.plugin;
    if (pl.opLog) return pl.opLog;
    try { pl.opLog = JSON.parse(await this.app.vault.adapter.read(this.ataDir() + '/log.json')); } catch (e) { pl.opLog = []; }
    if (!Array.isArray(pl.opLog)) pl.opLog = [];
    return pl.opLog;
  }

  saveOpLog() {
    const pl = this.plugin;
    clearTimeout(pl.opLogTimer);
    pl.opLogTimer = setTimeout(async () => {
      try {
        const a = this.app.vault.adapter;
        if (!(await a.exists(this.ataDir()))) await a.mkdir(this.ataDir());
        await a.write(this.ataDir() + '/log.json', JSON.stringify(pl.opLog.slice(-3000)));
      } catch (e) { console.warn('[AI表格助手] 保存操作记录失败', e); }
    }, 800);
  }

  tableLabel(path) {
    const m = (this.settings.meta || {})[path] || {};
    return m.sheet ? `${m.sheet} · ${m.workbook || ''}` : String(path || '').split('/').pop().replace(/\.(csv|md)$/, '');
  }

  /** 记一条操作；type: ai / edit / row / col / undo / import / create / export / delete / archive */
  logOp(e) {
    const list = this.plugin.opLog;
    if (!list) { this.loadOpLog().then(() => this.logOp(e)); return; }
    e.t = Date.now();
    if (e.path && !e.name) e.name = this.tableLabel(e.path);
    /* 连续手动改同一张表的格子（3 分钟内）合成一条 */
    const last = list[list.length - 1];
    if (e.type === 'edit' && last && last.type === 'edit' && last.path === e.path && e.t - last.t < 180000) {
      for (const c of e.cells || []) {
        const same = last.cells.find((x) => x.r === c.r && x.h === c.h);
        if (same) same.n = c.n; else last.cells.push(c);
      }
      last.cells = last.cells.filter((x) => x.o !== x.n).slice(-200);
      last.t = e.t;
      last.text = `手动修改了 ${last.cells.length} 个单元格`;
      if (!last.cells.length) list.pop();
    } else {
      if (e.type === 'edit' && !e.text) e.text = `手动修改了 ${(e.cells || []).length} 个单元格`;
      list.push(e);
    }
    if (list.length > 3000) list.splice(0, list.length - 3000);
    this.saveOpLog();
    if (this.tab === 'home' && !this.importRun) this.scheduleRender();
  }


  renderOpLog(body) {
    const sec = body.createDiv('ata-section ata-oplog');
    const h = sec.createDiv('ata-section-head');
    const t = h.createDiv('ata-section-title');
    this.ico(t, 'history');
    t.createSpan({ text: '操作记录' });
    const hr = h.createDiv('ata-oplog-tools');
    const box = sec.createDiv();
    if (!this.plugin.opLog) {
      box.createDiv({ cls: 'ata-empty-sm', text: '正在读取…' });
      this.loadOpLog().then(() => { if (this.tab === 'home') this.render(); });
      return;
    }
    const all = this.plugin.opLog;
    const KINDS = [['全部', null], ['外部修改', ['ext', 'ext-new', 'ext-del']], ['AI 修改（旧）', ['ai']], ['手动修改', ['edit', 'row', 'col', 'undo']], ['导入导出', ['import', 'export', 'create']], ['附件', ['archive']], ['删除', ['delete']]];
    const kindKey = this.opKind || '全部';
    const types = (KINDS.find((k) => k[0] === kindKey) || KINDS[0])[1];
    const chips = hr.createDiv('ata-catbar');
    KINDS.forEach(([label, ts]) => {
      const n = ts ? all.filter((e) => ts.includes(e.type)).length : all.length;
      if (ts && !n) return;
      const b = chips.createEl('button', { cls: 'ata-catchip' + (kindKey === label ? ' is-active' : '') });
      b.createSpan({ text: label });
      b.createSpan({ cls: 'ata-tab-count', text: String(n) });
      b.addEventListener('click', () => { this.opKind = label; this.opShow = 60; this.render(); });
    });
    const q = hr.createEl('input', { cls: 'ata-archive-search', attr: { type: 'text', placeholder: '搜表名、内容…' } });
    q.value = this.opQuery || '';
    q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { this.opQuery = q.value.trim(); this.opShow = 60; this.render(); } });
    q.addEventListener('change', () => { this.opQuery = q.value.trim(); this.opShow = 60; this.render(); });
    if (all.length) {
      const clr = hr.createEl('button', { cls: 'ata-iconbtn', attr: { 'aria-label': '清空操作记录' } });
      setIcon(clr, 'trash-2');
      clr.addEventListener('click', () => new ConfirmModal(this.app, '清空所有操作记录？', '只清空这里的记录，表格本身不受影响。', [{ text: '清空', warning: true, onClick: () => { this.plugin.opLog.length = 0; this.saveOpLog(); this.render(); } }]).open());
    }
    const ql = (this.opQuery || '').toLowerCase();
    const list = all.filter((e) => (!types || types.includes(e.type)) &&
      (!ql || [e.name, e.text, e.ask, ...(e.atts || [])].join(' ').toLowerCase().includes(ql) || (e.cells || []).some((c) => (c.o + ' ' + c.n).toLowerCase().includes(ql))));
    if (!list.length) {
      box.createDiv({ cls: 'ata-empty-sm', text: all.length ? '没有符合条件的记录。' : '还没有操作记录。之后导入、外部 AI 的修改、手动修改、导出、删除都会记在这里，可以随时回来查改了什么。' });
      return;
    }
    const TYPE = { ext: ['外部修改', 'bot'], 'ext-new': ['外部新建', 'file-plus'], 'ext-del': ['外部删除', 'file-x'], ai: ['AI', 'sparkles'], edit: ['手动', 'pencil'], row: ['行', 'rows-3'], col: ['列', 'columns-3'], undo: ['撤销', 'undo-2'], import: ['导入', 'folder-down'], export: ['导出', 'download'], create: ['新建', 'file-plus'], delete: ['删除', 'trash-2'], archive: ['附件', 'archive'] };
    const show = this.opShow || 60;
    const items = list.slice(-show).reverse();
    const dayOf = (ts) => { const d = new Date(ts); return `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`; };
    const today = dayOf(Date.now()), yest = dayOf(Date.now() - 86400000);
    if (!this.opOpen) this.opOpen = new Set();
    let curDay = '';
    const tl = box.createDiv('ata-tl');
    for (const e of items) {
      const d = dayOf(e.t);
      if (d !== curDay) {
        curDay = d;
        const dt = new Date(e.t);
        tl.createDiv({ cls: 'ata-tl-day', text: d === today ? '今天' : d === yest ? '昨天' : `${dt.getMonth() + 1} 月 ${dt.getDate()} 日 · ${['周日', '周一', '周二', '周三', '周四', '周五', '周六'][dt.getDay()]}` });
      }
      const it = tl.createDiv('ata-tl-item t-' + e.type);
      const tm = new Date(e.t);
      it.createDiv({ cls: 'ata-tl-time', text: `${String(tm.getHours()).padStart(2, '0')}:${String(tm.getMinutes()).padStart(2, '0')}` });
      const dot = it.createDiv('ata-tl-dot');
      setIcon(dot, (TYPE[e.type] || ['', 'dot'])[1]);
      const main = it.createDiv('ata-tl-main');
      const line = main.createDiv('ata-tl-line');
      line.createSpan({ cls: 'ata-tl-badge', text: (TYPE[e.type] || [e.type])[0] });
      if (e.path) {
        const tf = this.app.vault.getAbstractFileByPath(e.path);
        const chip = line.createEl(tf instanceof TFile ? 'button' : 'span', { cls: 'ata-tl-table' + (tf instanceof TFile ? '' : ' is-gone'), text: e.name || this.tableLabel(e.path) });
        if (tf instanceof TFile) chip.addEventListener('click', () => this.openTable(tf));
        else chip.title = '这张表已经不在了';
      }
      line.createSpan({ cls: 'ata-tl-text', text: e.text || '' });
      if (e.type === 'ext' && e.review) {
        line.createSpan({ cls: 'ata-tl-review r-' + e.review, text: { pending: '待确认', ok: '已确认', undone: '已撤销' }[e.review] || '' });
        if (e.review === 'pending') {
          const rb = main.createDiv('ata-tl-rbtns');
          const latest = this.pendingFor(e.path).slice(-1)[0] === e;
          const ok = rb.createEl('button', { cls: 'ata-tl-rbtn is-ok', text: '确认' });
          ok.addEventListener('click', () => this.confirmExternal(e.path));
          if (latest && e.hist) {
            const un = rb.createEl('button', { cls: 'ata-tl-rbtn', text: '撤销这次修改' });
            un.addEventListener('click', () => this.undoExternal(e));
          }
        }
      }
      if (e.ask) main.createDiv({ cls: 'ata-tl-ask', text: '“' + e.ask + '”' + (e.atts && e.atts.length ? `  📎 ${e.atts.join('、')}` : '') });
      else if (e.atts && e.atts.length) main.createDiv({ cls: 'ata-tl-ask', text: '📎 ' + e.atts.join('、') });
      if (e.file) {
        const ef = this.app.vault.getAbstractFileByPath(e.file);
        if (ef instanceof TFile) {
          const a = main.createEl('button', { cls: 'ata-tl-link', text: '打开 ' + ef.name });
          a.addEventListener('click', () => (this.app.openWithDefaultApp ? this.app.openWithDefaultApp(ef.path) : this.app.workspace.getLeaf('tab').openFile(ef)));
        }
      }
      const nCells = e.cellsN || (e.cells || []).length, nAdd = (e.added || []).length, nRow = e.rowsN || (e.rows || []).length;
      if (nCells || nAdd || nRow || (e.items && e.items.length)) {
        const key = e.t + '|' + e.type + '|' + (e.path || '');
        const open = this.opOpen.has(key);
        const tg = main.createEl('button', { cls: 'ata-tl-toggle', text: (open ? '收起' : '查看明细') + ` (${nCells + (e.addedN || nAdd) + nRow + (e.items ? e.items.length : 0)})` });
        tg.addEventListener('click', () => { if (this.opOpen.has(key)) this.opOpen.delete(key); else this.opOpen.add(key); this.render(); });
        if (open) this.renderOpDetail(main.createDiv('ata-tl-detail'), e);
      }
    }
    if (list.length > show) {
      const more = box.createEl('button', { cls: 'ata-tl-more', text: `显示更早的记录（还有 ${list.length - show} 条）` });
      more.addEventListener('click', () => { this.opShow = show + 100; this.render(); });
    }
  }

  renderOpDetail(el, e) {
    if (e.cells && e.cells.length) {
      el.createDiv({ cls: 'ata-tl-sub', text: `修改的单元格（${e.cellsN || e.cells.length}）` });
      const tb = el.createEl('table', { cls: 'ata-tl-tbl' });
      const hr = tb.createEl('tr');
      ['行', '列', '原来', '改成'].forEach((x) => hr.createEl('th', { text: x }));
      e.cells.slice(0, 100).forEach((c) => {
        const tr = tb.createEl('tr');
        tr.createEl('td', { text: c.r === 0 ? '表头' : String(c.r) });
        tr.createEl('td', { text: c.h });
        tr.createEl('td', { cls: 'is-old', text: c.o || '（空）' });
        tr.createEl('td', { cls: 'is-new', text: c.n || '（空）' });
      });
      if ((e.cellsN || e.cells.length) > 100) el.createDiv({ cls: 'ata-muted', text: `……共 ${e.cellsN || e.cells.length} 个` });
    }
    const rowsTable = (title, hdr, rows, more, cls) => {
      el.createDiv({ cls: 'ata-tl-sub', text: title });
      const used = hdr.map((_, i) => i).filter((i) => rows.some((r) => r[i] !== '' && r[i] != null)).slice(0, 12);
      const wrap = el.createDiv('ata-tl-wrap');
      const tb = wrap.createEl('table', { cls: 'ata-tl-tbl ' + cls });
      const hr = tb.createEl('tr');
      used.forEach((i) => hr.createEl('th', { text: hdr[i] }));
      rows.forEach((r) => { const tr = tb.createEl('tr'); used.forEach((i) => tr.createEl('td', { text: r[i] == null ? '' : String(r[i]) })); });
      if (more > rows.length) el.createDiv({ cls: 'ata-muted', text: `……共 ${more} 行` });
    };
    if (e.added && e.added.length) rowsTable(`新增的行（${e.addedN || e.added.length}）`, e.hdr || [], e.added, e.addedN || 0, 'is-added');
    if (e.rows && e.rows.length) rowsTable(`删除的行（${e.rowsN || e.rows.length}）`, e.oldHdr || e.hdr || [], e.rows, e.rowsN || 0, 'is-removed');
    if (e.items && e.items.length) {
      const ul = el.createEl('ul', { cls: 'ata-tl-list' });
      e.items.slice(0, 50).forEach((x) => ul.createEl('li', { text: x }));
      if (e.items.length > 50) el.createDiv({ cls: 'ata-muted', text: `……共 ${e.items.length} 项` });
    }
    if (e.notes && e.notes.length && e.type === 'ai' && !(e.cells && e.cells.length) && !(e.added && e.added.length)) {
      const ul = el.createEl('ul', { cls: 'ata-tl-list' });
      e.notes.forEach((x) => ul.createEl('li', { text: x }));
    }
  }

  async ensureAdapterDir(dir) {
    const a = this.app.vault.adapter;
    let cur = '';
    for (const seg of normalizePath(dir).split('/')) {
      cur = cur ? cur + '/' + seg : seg;
      if (!(await a.exists(cur))) { try { await a.mkdir(cur); } catch (e) { /* 并发时可能已存在 */ } }
    }
  }

  async readLink(id, idx) {
    const a = this.app.vault.adapter;
    const base = `${this.ataDir()}/maps/${id}`;
    try {
      const map = JSON.parse(await a.read(`${base}/${idx}.json`));
      let styles = null;
      if (map.styles) {
        /* 旧格式：颜色信息挪到单独的文件里，之后每次保存就不用重写这几万条 */
        styles = map.styles;
        delete map.styles;
        try { await a.write(`${base}/${idx}.styles.json`, JSON.stringify(styles)); await a.write(`${base}/${idx}.json`, JSON.stringify(map)); } catch (e) { /* 下次再迁移 */ }
      } else {
        try { styles = JSON.parse(await a.read(`${base}/${idx}.styles.json`)); } catch (e) { styles = null; }
      }
      if (!this.xfCache) this.xfCache = new Map();
      if (!this.xfCache.has(id)) {
        let xf = null;
        try { xf = JSON.parse(await a.read(`${base}/xf.json`)); } catch (e) { /* 没有颜色信息 */ }
        this.xfCache.set(id, xf);
      }
      return { id, sheet: idx, map, styles, xf: this.xfCache.get(id) };
    } catch (e) {
      return null;
    }
  }

  async writeLinkMap(link) {
    try {
      await this.app.vault.adapter.write(`${this.ataDir()}/maps/${link.id}/${link.sheet}.json`, JSON.stringify(link.map));
    } catch (e) { console.warn('[AI表格助手] 保存对应关系失败', e); }
  }

  /** 给表格行/列贴上「原来在 Excel 哪一行、哪一列」 */
  attachProvenance(t, map) {
    const rows = map.rows || [];
    t.rows.forEach((r, i) => { r.__src = i < rows.length ? rows[i] : -1; });
    t.colSrc = map.cols && map.cols.length === t.headers.length ? map.cols.slice() : null;
  }

  async loadTableFor(f) {
    const text = await this.app.vault.read(f);
    const t = gridToTable(parseCSV(text));
    const m = (this.settings.meta || {})[f.path];
    let link = null;
    if (m && m.link) {
      link = await this.readLink(m.link.id, m.link.sheet);
      if (link) this.attachProvenance(t, link.map);
    }
    return { table: t, link };
  }

  /** 按内容估算列宽（只看表头和抽样行），之后固定不变，编辑时不用重排整张表 */
  colWidths(t) {
    const units = (v) => { let n = 0; for (const ch of cellStr(v)) n += ch.charCodeAt(0) > 255 ? 2 : 1; return n; };
    const sample = t.rows.length > 400 ? [...t.rows.slice(0, 300), ...t.rows.slice(-100)] : t.rows;
    return t.headers.map((h, ci) => {
      let u = units(h) + 2;
      for (const r of sample) { const x = units(r[ci]); if (x > u) u = x; if (u > 40) break; }
      return Math.round(Math.max(56, Math.min(300, u * 7.2 + 22)));
    });
  }

  /**
   * 虚拟滚动：页面上只放看得见的几十行，再多的行数也不卡。
   * 行号、表头固定；列宽固定（table-layout: fixed），编辑单元格时不触发整表重排。
   */
  renderTable() {
    if (!this.previewEl) return;
    const key = this.file && this.file.path;
    const keep = this.vtFor === key ? { top: this.previewEl.scrollTop, left: this.previewEl.scrollLeft } : null;
    this.vtFor = key;
    this.previewEl.empty();
    this.previewEl.onscroll = null;
    const t = this.table;
    if (!t.headers.length) {
      this.previewEl.createDiv({ cls: 'ata-empty-sm ata-pad', text: '这是一个空表。可以点下方「新增列」，或者在右边让外部 AI 按单据建好列。' });
      if (this.statEl) this.statEl.setText('');
      return;
    }
    const ROW_H = 30, NUM_W = 54, OVER = 12;
    const vis = this.link && this.link.xf ? this.link.xf : null;
    const sty = vis && this.link.styles ? this.link.styles : null;
    const paint = (el, R, C) => {
      if (!sty || R == null || R < 0 || C == null || C < 0) return;
      const sIdx = sty[R + ':' + C];
      const v = sIdx != null ? vis[sIdx] : null;
      if (!v) return;
      if (v.bg) { el.style.background = v.bg; if (!v.color) el.style.color = '#1a1a1a'; }
      if (v.color) el.style.color = v.color;
      if (v.b) el.style.fontWeight = '700';
      if (v.i) el.style.fontStyle = 'italic';
    };
    const colSrc = t.colSrc || [];
    const widths = this.colWidths(t);

    const table = document.createElement('table');
    table.className = 'ata-table ata-vtable';
    table.style.width = (NUM_W + widths.reduce((a, b) => a + b, 0)) + 'px';
    const cg = document.createElement('colgroup');
    [NUM_W, ...widths].forEach((w) => { const c = document.createElement('col'); c.style.width = w + 'px'; cg.appendChild(c); });
    table.appendChild(cg);

    const thead = table.createTHead().insertRow();
    const th0 = document.createElement('th');
    th0.className = 'ata-num';
    th0.textContent = '#';
    thead.appendChild(th0);
    t.headers.forEach((h, ci) => {
      const th = document.createElement('th');
      const label = document.createElement('span');
      label.className = 'ata-th-label';
      label.textContent = h;
      label.contentEditable = 'true';
      label.dataset.c = ci;
      label.title = h + '（点击修改列名）';
      const del = document.createElement('button');
      del.className = 'ata-coldel';
      del.textContent = '×';
      del.title = `删除「${h}」这一列`;
      del.dataset.col = ci;
      th.appendChild(label);
      th.appendChild(del);
      if (this.link) paint(th, this.link.map.headerRow, colSrc[ci]);
      thead.appendChild(th);
    });

    const tbody = table.createTBody();
    const pad = () => { const tr = document.createElement('tr'); tr.className = 'ata-vpad'; const td = document.createElement('td'); td.colSpan = t.headers.length + 1; tr.appendChild(td); return tr; };
    const topPad = pad(), botPad = pad();
    tbody.appendChild(topPad);
    tbody.appendChild(botPad);
    const vt = { start: -1, end: -1 };

    const hl = this.hlMap && key ? this.hlMap.get(key) : null;
    const buildRow = (ri) => {
      const r = t.rows[ri];
      const tr = document.createElement('tr');
      tr.dataset.r = ri;
      if (hl && hl.added.has(ri)) tr.className = 'is-ext-add';
      const hlc = hl && hl.cells.get(ri);
      const num = document.createElement('td');
      num.className = 'ata-num';
      num.dataset.r = ri;
      const nt = document.createElement('span');
      nt.className = 'ata-num-text';
      nt.textContent = String(ri + 1);
      const del = document.createElement('button');
      del.className = 'ata-rowdel';
      del.setAttribute('aria-label', '删除这一行');
      del.textContent = '×';
      num.appendChild(nt);
      num.appendChild(del);
      tr.appendChild(num);
      for (let ci = 0; ci < t.headers.length; ci++) {
        const td = document.createElement('td');
        const v = r[ci] == null ? '' : String(r[ci]);
        td.textContent = v;
        if (v.length > 10) td.title = v;
        td.contentEditable = 'true';
        td.dataset.r = ri;
        td.dataset.c = ci;
        if (sty) paint(td, r.__src, colSrc[ci]);
        if (hlc && hlc.has(ci)) td.classList.add('is-ext-mod');
        tr.appendChild(td);
      }
      return tr;
    };

    const draw = (force) => {
      const n = t.rows.length;
      const wrap = this.previewEl;
      const vh = wrap.clientHeight || 600;
      const top = Math.max(0, wrap.scrollTop - ROW_H);
      const first = Math.floor(top / ROW_H);
      const start = Math.max(0, Math.floor(first / 4) * 4 - OVER);
      const end = Math.min(n, start + Math.ceil(vh / ROW_H) + OVER * 2 + 4);
      if (!force && start === vt.start && end === vt.end) return;
      /* 正在编辑的格子如果还在可见范围内，先不重画，免得打断输入 */
      const act = document.activeElement;
      if (!force && act && tbody.contains(act) && act.dataset && act.dataset.r != null) {
        const ar = +act.dataset.r;
        if (ar >= start && ar < end) return;
        act.blur();
      }
      const overlap = !force && vt.start >= 0 && start < vt.end && end > vt.start;
      if (overlap) {
        /* 只增删滚进来 / 滚出去的行 */
        for (let ri = vt.start; ri < Math.min(start, vt.end); ri++) { const tr = topPad.nextSibling; if (tr !== botPad) tbody.removeChild(tr); }
        for (let ri = Math.max(end, vt.start); ri < vt.end; ri++) { const tr = botPad.previousSibling; if (tr !== topPad) tbody.removeChild(tr); }
        if (start < vt.start) {
          const frag = document.createDocumentFragment();
          for (let ri = start; ri < vt.start; ri++) frag.appendChild(buildRow(ri));
          tbody.insertBefore(frag, topPad.nextSibling);
        }
        if (end > vt.end) {
          const frag = document.createDocumentFragment();
          for (let ri = vt.end; ri < end; ri++) frag.appendChild(buildRow(ri));
          tbody.insertBefore(frag, botPad);
        }
      } else {
        while (topPad.nextSibling && topPad.nextSibling !== botPad) tbody.removeChild(topPad.nextSibling);
        const frag = document.createDocumentFragment();
        for (let ri = start; ri < end; ri++) frag.appendChild(buildRow(ri));
        tbody.insertBefore(frag, botPad);
      }
      vt.start = start;
      vt.end = end;
      topPad.firstChild.style.height = (start * ROW_H) + 'px';
      botPad.firstChild.style.height = ((n - end) * ROW_H) + 'px';
      topPad.style.display = start ? '' : 'none';
      botPad.style.display = n - end ? '' : 'none';
    };
    this.redrawRows = () => draw(true);
    let raf = 0;
    this.previewEl.onscroll = () => { if (!raf) raf = requestAnimationFrame(() => { raf = 0; draw(false); }); };

    table.addEventListener('focusout', (e) => {
      const el = e.target;
      if (!el || el.dataset.c == null) return;
      const ci = +el.dataset.c;
      if (el.classList.contains('ata-th-label')) {
        const o = String(t.headers[ci]);
        this.commitEdit(el, () => t.headers[ci], (v) => { t.headers[ci] = v; this.logOp({ type: 'edit', path: this.file.path, text: '', cells: [{ r: 0, h: '（列名）', o, n: v }] }); });
      } else {
        const ri = +el.dataset.r;
        if (!t.rows[ri]) return;
        const o = cellStr(t.rows[ri][ci]);
        this.commitEdit(el, () => cellStr(t.rows[ri][ci]), (v) => { t.rows[ri][ci] = v; if (v.length > 10) el.title = v; else el.removeAttribute('title'); this.logOp({ type: 'edit', path: this.file.path, text: '', cells: [{ r: ri + 1, h: String(t.headers[ci]), o, n: v }] }); });
      }
    });
    table.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey && e.target.dataset && e.target.dataset.c != null) { e.preventDefault(); e.target.blur(); }
    });
    table.addEventListener('click', async (e) => {
      const cd = e.target.closest && e.target.closest('.ata-coldel');
      if (cd) {
        const ci = +cd.dataset.col;
        this.pushUndo();
        this.logOp({ type: 'col', path: this.file.path, text: `删除了列「${t.headers[ci]}」（${t.rows.filter((r) => !isBlank(r[ci])).length} 个有内容的格子）` });
        t.headers.splice(ci, 1);
        t.rows.forEach((r) => r.splice(ci, 1));
        if (t.colSrc) t.colSrc.splice(ci, 1);
        await this.saveTable();
        this.renderTable();
        if (this.undoBtn) this.undoBtn.disabled = false;
        new Notice('已删除一列，可以点「撤销」恢复。');
        return;
      }
      const del = e.target.closest && e.target.closest('.ata-rowdel');
      if (!del) return;
      const ri = +del.parentElement.dataset.r;
      this.pushUndo();
      this.logOp({ type: 'row', path: this.file.path, text: `删除了第 ${ri + 1} 行`, hdr: t.headers.map(String), rows: [(t.rows[ri] || []).map(cellStr)] });
      t.rows.splice(ri, 1);
      if (this.hlMap) this.hlMap.delete(this.file.path);
      await this.saveTable();
      this.renderTable();
      if (this.undoBtn) this.undoBtn.disabled = false;
    });
    this.previewEl.appendChild(table);
    if (keep) { this.previewEl.scrollTop = keep.top; this.previewEl.scrollLeft = keep.left; }
    draw(true);
    if (this.statEl) this.statEl.setText(`${t.rows.length.toLocaleString()} 行 × ${t.headers.length} 列 · 单元格可直接编辑`);
  }

  async commitEdit(el, get, set) {
    const nv = el.innerText.replace(/\n$/, '');
    if (nv === get()) return;
    this.pushUndo();
    set(nv);
    this.scheduleSave();
    if (this.undoBtn) this.undoBtn.disabled = false;
  }

  scheduleSave() {
    clearTimeout(this.saveTimer);
    this.saveTimer = setTimeout(() => { this.saveTimer = null; this.saveTable(); }, 600);
  }

  async flushSave() {
    if (this.saveTimer) { clearTimeout(this.saveTimer); this.saveTimer = null; await this.saveTable(); }
  }

  async addRow() {
    if (!this.table.headers.length) { new Notice('请先新增列。'); return; }
    this.pushUndo();
    this.logOp({ type: 'row', path: this.file.path, text: '新增了一个空行' });
    this.table.rows.push(this.table.headers.map(() => ''));
    await this.saveTable();
    this.renderTable();
    this.previewEl.scrollTop = this.previewEl.scrollHeight;
    if (this.redrawRows) this.redrawRows();
    const last = this.previewEl.querySelector(`td[data-r="${this.table.rows.length - 1}"][data-c="0"]`);
    if (last) last.focus();
  }

  async addCol() {
    this.pushUndo();
    this.logOp({ type: 'col', path: this.file.path, text: `新增了列「新列${this.table.headers.length + 1}」` });
    addColumn(this.table, '新列' + (this.table.headers.length + 1));
    await this.saveTable();
    this.renderTable();
    this.previewEl.scrollLeft = this.previewEl.scrollWidth;
  }

  pickHeaderRow() {
    const t = this.table;
    if (!t.rows.length) return;
    new HeaderRowModal(this.app, t, async (n) => {
      this.pushUndo();
      this.logOp({ type: 'col', path: this.file.path, text: `把第 ${n} 行设为表头` });
      const above = [t.headers, ...t.rows.slice(0, n - 1)]
        .map((r) => r.filter((c) => !isBlank(c) && !/^(列|新列)\d+$/.test(c)).join(' ')).filter(Boolean).join(' / ');
      const passed = [t.rows[n - 1], ...t.rows.slice(n)];
      const res = cleanGrid(passed);
      if (res.src) {
        res.table.rows.forEach((r, i) => { const old = passed[res.src.rows[i]]; r.__src = old && old.__src != null ? old.__src : -1; });
        if (t.colSrc) res.table.colSrc = res.src.cols.map((k) => (t.colSrc[k] != null ? t.colSrc[k] : -1));
        if (this.link && t.rows[n - 1].__src >= 0) this.link.map.headerRow = t.rows[n - 1].__src;
      }
      this.table = res.table;
      if (above) {
        const meta = this.settings.meta || (this.settings.meta = {});
        meta[this.file.path] = Object.assign({}, meta[this.file.path], { title: above });
        await this.plugin.saveSettings();
      }
      await this.saveTable();
      this.render();
      new Notice(`已把第 ${n} 行设为表头${above ? '，上方标题已存为表格说明' : ''}。可点「撤销」恢复。`);
    }).open();
  }












  /* ---------- 文件读写 ---------- */
  async ensureFolder(path) {
    let cur = '';
    for (const seg of normalizePath(path).split('/')) {
      cur = cur ? cur + '/' + seg : seg;
      if (!this.app.vault.getAbstractFileByPath(cur)) {
        try { await this.app.vault.createFolder(cur); } catch (e) { /* 已存在 */ }
      }
    }
  }

  async uniquePath(folder, base, ext) {
    await this.ensureFolder(folder);
    let p = normalizePath(`${folder}/${base}.${ext}`), i = 1;
    while (this.app.vault.getAbstractFileByPath(p)) p = normalizePath(`${folder}/${base}-${i++}.${ext}`);
    return p;
  }

  async openTable(file) {
    await this.loadFile(file, { switchTab: true });
  }

  async loadFile(file, opts) {
    opts = opts || {};
    await this.flushSave();
    const big = file.stat.size > 8e6 ? new Notice(`正在打开大表「${file.basename}」（${(file.stat.size / 1048576).toFixed(0)}MB）…`, 0) : null;
    try {
      const text = await this.app.vault.read(file);
      if (file.extension === 'md') {
        const f = findMdTable(text);
        this.table = f ? f.table : { headers: [], rows: [] };
        if (!f && !opts.silent) new Notice('这篇笔记里还没有表格。');
      } else {
        this.table = gridToTable(parseCSV(text));
      }
      const lm = (this.settings.meta || {})[file.path];
      this.link = lm && lm.link ? await this.readLink(lm.link.id, lm.link.sheet) : null;
      if (this.link) this.attachProvenance(this.table, this.link.map);
      if (!this.file || this.file.path !== file.path) {
        if (this.file) {
          this.undoMap.set(this.file.path, this.undoStack);
          if (this.undoMap.size > 8) this.undoMap.delete(this.undoMap.keys().next().value);
        }
        this.undoStack = this.undoMap.get(file.path) || [];
        this.showRows = 0;
      }
      const rec = (this.settings.recent || []).filter((p) => p !== file.path);
      rec.unshift(file.path);
      this.settings.recent = rec.slice(0, 20);
      this.tableVersion = (this.tableVersion || 0) + 1;
      this.file = file;
      const meta = this.settings.meta || (this.settings.meta = {});
      if (file.extension === 'md' && !file.path.startsWith(normalizePath(this.settings.folder) + '/') && !meta[file.path]) meta[file.path] = { pinned: true };
      this.settings.lastTable = file.path;
      await this.plugin.saveSettings();
      if (big) big.hide();
      if (opts.switchTab) { this.tab = 'workspace'; await this.render(); }
    } catch (e) {
      if (big) big.hide();
      new Notice('打开表格失败：' + e.message);
    }
  }

  async saveTable() {
    if (!this.file) return;
    this.selfWrite = true;
    try {
      if (this.file.extension === 'md') {
        const text = await this.app.vault.read(this.file);
        const f = findMdTable(text);
        const md = toMarkdown(this.table);
        let out;
        if (f) {
          const lines = text.split('\n');
          lines.splice(f.start, f.end - f.start, md);
          out = lines.join('\n');
        } else {
          out = text.replace(/\s*$/, '') + (text.trim() ? '\n\n' : '') + md + '\n';
        }
        await this.snapPut(this.file.path, out);
        await this.app.vault.modify(this.file, out);
      } else {
        const csv = toCSV(this.table);
        await this.snapPut(this.file.path, csv);
        await this.app.vault.modify(this.file, csv);
        if (this.link) {
          this.link.map.rows = this.table.rows.map((r) => (r.__src == null ? -1 : r.__src));
          if (this.table.colSrc && this.table.colSrc.length === this.table.headers.length) this.link.map.cols = this.table.colSrc.slice();
          await this.writeLinkMap(this.link);
        }
      }
    } finally {
      setTimeout(() => (this.selfWrite = false), 300);
    }
  }

  async handleDrop(list, zone) {
    if (!list.length) return;
    if (zone === 'archive' || this.tab === 'archive') { await this.archiveSave(list.map((x) => x.file).filter((f) => f.size > 0), '拖入'); return; }
    const isFolder = list.some((x) => x.rel.includes('/'));
    if (zone === 'import' || isFolder) { await this.importBatch(list); return; }
    if (zone && zone.startsWith('table:')) {
      const f = this.app.vault.getAbstractFileByPath(zone.slice(6));
      if (f instanceof TFile) { await this.receiveAtts(list, f); return; }
    }
    if (zone === 'ai' && this.file) { await this.receiveAtts(list, this.file); return; }
    if (zone === 'inbox') { await this.receiveAtts(list, null); return; }
    await this.receive(list);
  }

  /** 删除表格文件（进回收站），并清理分类、常用、最近等记录 */
  async deleteFiles(files) {
    files = files.filter((f) => f instanceof TFile);
    if (!files.length) return;
    const names = files.map((f) => this.tableLabel(f.path));
    this.logOp({ type: 'delete', text: files.length === 1 ? `删除了表格「${names[0]}」（在回收站）` : `删除了 ${files.length} 个表格（在回收站）`, items: names });
    this.selfDeleting = (this.selfDeleting || 0) + 1;
    setTimeout(() => { this.selfDeleting = Math.max(0, (this.selfDeleting || 1) - 1); }, 4000 + files.length * 50);
    const folders = new Set();
    let n = 0;
    for (const f of files) {
      try {
        if (this.file && f.path === this.file.path) {
          clearTimeout(this.saveTimer);
          this.saveTimer = null;
          this.file = null;
          this.table = { headers: [], rows: [] };
          this.undoStack = [];
        }
        if (f.parent && f.parent.path && f.parent.path !== '/') folders.add(f.parent);
        if (this.app.fileManager && this.app.fileManager.trashFile) await this.app.fileManager.trashFile(f);
        else await this.app.vault.trash(f, true);
        const st = this.settings;
        if (st.meta) delete st.meta[f.path];
        st.pinned = (st.pinned || []).filter((x) => x !== f.path);
        st.recent = (st.recent || []).filter((x) => x !== f.path);
        if (st.lastTable === f.path) st.lastTable = '';
        if (st.master === f.path) st.master = '';
        this.undoMap.delete(f.path);
        if (this.aiStates) this.aiStates.delete(f.path);
        this.stats.delete(f.path);
        if (this.selected) this.selected.delete(f.path);
        n++;
      } catch (e) {
        new Notice(`删除「${f.name}」失败：${e.message || e}`, 8000);
      }
    }
    /* 删空了的文件夹（例如整个 Excel 的工作表都删了）也一并移走，但不动表格库根目录 */
    const root = normalizePath(this.settings.folder);
    for (const folder of folders) {
      try {
        if (folder.path !== root && folder.children && folder.children.length === 0) {
          if (this.app.fileManager && this.app.fileManager.trashFile) await this.app.fileManager.trashFile(folder);
          else await this.app.vault.trash(folder, true);
        }
      } catch (e) { /* 不是空文件夹就留着 */ }
    }
    /* 某个 Excel 的工作表全删了，就把保存的原文件和对应关系也删掉 */
    try {
      const alive = new Set(Object.values(this.settings.meta || {}).map((m) => m && m.link && m.link.id).filter(Boolean));
      const a = this.app.vault.adapter;
      const srcDir = `${this.ataDir()}/sources`;
      if (await a.exists(srcDir)) {
        const ls = await a.list(srcDir);
        for (const fp of ls.files || []) {
          const id = fp.split('/').pop().replace(/\.xlsx$/, '');
          if (!alive.has(id)) {
            await a.remove(fp);
            const md = `${this.ataDir()}/maps/${id}`;
            if (await a.exists(md)) await a.rmdir(md, true);
          }
        }
      }
    } catch (e) { /* 清理失败不影响 */ }
    await this.plugin.saveSettings();
    new Notice(`已删除 ${n} 个表格（在回收站里可以恢复）。`);
    if (!this.file && this.tab === 'workspace') this.tab = 'library';
    await this.render();
  }

  confirmDelete(file) {
    const meta = (this.settings.meta || {})[file.path] || {};
    const sibs = this.workbookSiblings(file);
    const name = meta.sheet || file.basename;
    const buttons = [{ text: sibs.length > 1 ? '只删这个工作表' : '删除', warning: true, onClick: () => this.deleteFiles([file]) }];
    if (sibs.length > 1) buttons.push({ text: `删除整个「${meta.workbook}」（${sibs.length} 个工作表）`, warning: true, onClick: () => this.deleteFiles(sibs) });
    new ConfirmModal(this.app, `删除「${name}」？`, `文件：${file.path}\n会移到回收站（系统回收站或库里的 .trash，取决于 Obsidian 的设置），需要时可以恢复。`, buttons).open();
  }

  /* ================= 附件区 ================= */
  archiveDir() { return normalizePath(this.settings.folder + '/附件区'); }

  archiveFiles() {
    const dir = this.archiveDir() + '/';
    return this.app.vault.getFiles().filter((f) => f.path.startsWith(dir)).sort((a, b) => b.stat.ctime - a.stat.ctime || b.stat.mtime - a.stat.mtime);
  }

  /** 存进附件区（按月份分文件夹；同名同大小的文件不重复存） */
  async archiveSave(files, from, quiet) {
    files = (files || []).filter(Boolean);
    if (!files.length) return [];
    const d = new Date(), z = (n) => String(n).padStart(2, '0');
    const folder = `${this.archiveDir()}/${d.getFullYear()}-${z(d.getMonth() + 1)}`;
    const existing = this.archiveFiles();
    const meta = this.settings.archiveMeta || (this.settings.archiveMeta = {});
    const saved = [];
    let dup = 0, tooBig = 0;
    this.archiving = true;
    try {
      for (const f of files) {
        if (f.size > MAX_FILE) { tooBig++; continue; }
        const same = existing.find((x) => x.name === f.name && x.stat.size === f.size);
        if (same) { dup++; saved.push(same); continue; }
        const dot = f.name.lastIndexOf('.');
        const base = safeSeg(dot > 0 ? f.name.slice(0, dot) : f.name) || '文件';
        const ext = dot > 0 ? f.name.slice(dot + 1) : 'bin';
        const path = await this.uniquePath(folder, base, ext);
        const tf = await this.app.vault.createBinary(path, await f.arrayBuffer());
        meta[tf.path] = { added: Date.now(), from: from || '' };
        saved.push(tf);
        existing.push(tf);
      }
    } finally {
      this.archiving = false;
    }
    await this.plugin.saveSettings();
    if (!quiet) {
      const msg = [`已存入附件区 ${saved.length - dup} 个文件`];
      if (dup) msg.push(`${dup} 个已经存过`);
      if (tooBig) msg.push(`${tooBig} 个超过 200MB 未存`);
      new Notice(msg.join('，'));
    }
    const fresh = saved.length - dup;
    if (fresh > 0) this.logOp({ type: 'archive', text: `存入附件区 ${fresh} 个文件${from ? `（${from}）` : ''}`, items: saved.slice(saved.length - fresh).map((f) => f.name) });
    if (this.tab === 'archive') await this.render();
    return saved;
  }




  /** 缩略图：图片直接显示，PDF 渲染第一页（看得到的时候才渲染），其它显示图标 */
  renderThumb(box, tf) {
    const kind = archiveKind(tf);
    box.addClass('k-' + kind);
    if (kind === 'image') {
      box.createEl('img', { attr: { src: this.app.vault.getResourcePath(tf), loading: 'lazy', alt: tf.name } });
      return;
    }
    const icon = box.createDiv('ata-athumb-icon');
    setIcon(icon, kind === 'pdf' ? 'file-text' : kind === 'table' ? 'file-spreadsheet' : kind === 'doc' ? 'file-type' : 'file');
    box.createDiv({ cls: 'ata-athumb-ext', text: (tf.extension || '').toUpperCase() });
    if (kind !== 'pdf') return;
    if (!this.thumbCache) this.thumbCache = new Map();
    const key = tf.path + '@' + tf.stat.mtime;
    const put = (url) => { if (!url || !box.isConnected) return; box.empty(); box.addClass('has-img'); box.createEl('img', { attr: { src: url } }); };
    if (this.thumbCache.has(key)) { put(this.thumbCache.get(key)); return; }
    const io = new IntersectionObserver((ents) => {
      if (!ents.some((x) => x.isIntersecting)) return;
      io.disconnect();
      this.thumbQueue = (this.thumbQueue || Promise.resolve()).then(async () => {
        try {
          const r = await renderPdfPages(this.app, tf, 360, 1);
          this.thumbCache.set(key, r.urls[0]);
          put(r.urls[0]);
        } catch (e) { this.thumbCache.set(key, null); }
      });
    });
    setTimeout(() => io.observe(box), 0);
  }

  renderArchive(body) {
    const wrap = body.createDiv('ata-archive');
    const all = this.archiveFiles();
    const meta = this.settings.archiveMeta || {};
    if (!this.archSel) this.archSel = new Set();
    const head = wrap.createDiv('ata-header');
    const left = head.createDiv();
    const title = left.createDiv('ata-title');
    this.ico(title, 'archive');
    title.createSpan({ text: '附件区' });
    left.createDiv({ cls: 'ata-sub', text: `${all.length} 个文件 · 共 ${fmtSize(all.reduce((a, f) => a + f.stat.size, 0))} · 收到的单据、截图、PDF、Excel 都可以先存在这里，要用时点「交给 AI」` });
    const right = head.createDiv('ata-header-right');
    this.btn(right, this.archSelect ? '完成' : '批量管理', this.archSelect ? 'check' : 'list-checks', () => { this.archSelect = !this.archSelect; this.archSel.clear(); this.render(); });
    if (this.app.showInFolder) this.btn(right, '打开文件夹', 'folder-open', async () => { await this.ensureFolder(this.archiveDir()); this.app.showInFolder(this.archiveDir()); });
    this.btn(right, '上传文件', 'upload', () => this.archiveInput.click(), 'mod-cta');

    const dz = wrap.createDiv('ata-dropzone ata-archive-drop');
    this.ico(dz, 'archive', 'ata-dropzone-icon');
    const txt = dz.createDiv('ata-dropzone-text');
    txt.createEl('b', { text: '把文件拖到这里存档（任何类型都可以）' });
    txt.createSpan({ text: `也可以点这里选择文件，或在本页按 Ctrl+V 粘贴复制的文件 / 截图 / 文件路径。保存在「${this.archiveDir()}」，按月份分文件夹。外部 AI 可以直接读这些文件。` });
    dz.addEventListener('click', () => this.archiveInput.click());

    if (!all.length) return;
    const counts = {};
    all.forEach((f) => { const k = archiveKind(f); counts[k] = (counts[k] || 0) + 1; });
    const fbar = wrap.createDiv('ata-filterbar');
    const catbar = fbar.createDiv('ata-catbar');
    if (!this.archKind) this.archKind = null;
    ARCHIVE_KINDS.forEach(([label, k]) => {
      const n = k ? counts[k] || 0 : all.length;
      if (k && !n) return;
      const b = catbar.createEl('button', { cls: 'ata-catchip' + (this.archKind === k ? ' is-active' : '') });
      b.createSpan({ text: label });
      b.createSpan({ cls: 'ata-tab-count', text: String(n) });
      b.addEventListener('click', () => { this.archKind = k; this.render(); });
    });
    const q = fbar.createEl('input', { cls: 'ata-archive-search', attr: { type: 'text', placeholder: '搜索文件名、备注…' } });
    q.value = this.archQuery || '';
    let list = all.filter((f) => (!this.archKind || archiveKind(f) === this.archKind));
    const applyQ = () => {
      const s = (this.archQuery || '').toLowerCase();
      grid.querySelectorAll('.ata-acard').forEach((c) => { c.style.display = !s || c.dataset.q.includes(s) ? '' : 'none'; });
    };
    q.addEventListener('input', () => { this.archQuery = q.value.trim(); applyQ(); });

    if (this.archSelect) {
      const sb = wrap.createDiv('ata-selbar');
      const cnt = sb.createSpan({ cls: 'ata-selcount', text: `已选 ${this.archSel.size} 个` });
      this.archCountEl = cnt;
      const sel = () => all.filter((f) => this.archSel.has(f.path));
      this.btn(sb, '全选', 'check-square', () => { list.forEach((f) => this.archSel.add(f.path)); this.render(); }, 'ata-btn-sm');
      this.btn(sb, '交给 AI 填表', 'bot', () => { const s = sel(); if (s.length) this.handoff(s); }, 'ata-btn-sm mod-cta');
      if (this.file) this.btn(sb, '加到当前表的附件', 'paperclip', () => { const s = sel(); if (s.length) { this.addAiAtts(this.file.path, s); this.openTable(this.file); } }, 'ata-btn-sm');
      this.btn(sb, '删除所选', 'trash-2', () => {
        const s = sel();
        if (!s.length) return;
        new ConfirmModal(this.app, `删除 ${s.length} 个附件？`, s.slice(0, 8).map((f) => '• ' + f.name).join('\n') + '\n会移到回收站，需要时可以恢复。', [{ text: `删除 ${s.length} 个`, warning: true, onClick: () => this.archiveDelete(s) }]).open();
      }, 'ata-btn-sm is-danger');
    }

    const grid = wrap.createDiv('ata-agrid');
    list.slice(0, 600).forEach((f) => {
      const m = meta[f.path] || {};
      const c = grid.createDiv('ata-acard' + (this.archSel.has(f.path) ? ' is-selected' : ''));
      c.dataset.q = (f.name + ' ' + (m.note || '')).toLowerCase();
      const th = c.createDiv('ata-athumb');
      this.renderThumb(th, f);
      th.addEventListener('click', () => { if (!this.archSelect) new LightboxModal(this.app, this, f).open(); });
      const info = c.createDiv('ata-ainfo');
      info.createDiv({ cls: 'ata-aname', text: f.name, attr: { title: f.path } });
      info.createDiv({ cls: 'ata-asub', text: `${fmtSize(f.stat.size)} · ${fmtTime(m.added || f.stat.ctime || f.stat.mtime)}${m.from ? ' · ' + m.from : ''}` });
      if (m.used && m.used.length) {
        const u = info.createDiv('ata-aused');
        u.createSpan({ text: '已提取到 ' });
        m.used.slice(0, 2).forEach((p) => {
          const tf = this.app.vault.getAbstractFileByPath(p);
          const name = ((this.settings.meta || {})[p] || {}).sheet || p.split('/').pop().replace(/\.csv$/, '');
          const chip = u.createEl('button', { cls: 'ata-colchip ata-sheetchip', text: name });
          chip.addEventListener('click', (e) => { e.stopPropagation(); if (tf instanceof TFile) this.openTable(tf); });
        });
        if (m.used.length > 2) u.createSpan({ cls: 'ata-muted', text: ` 等 ${m.used.length} 张表` });
      }
      const note = info.createDiv({ cls: 'ata-anote' + (m.note ? '' : ' is-empty'), text: m.note || '＋ 备注' });
      note.addEventListener('click', (e) => {
        e.stopPropagation();
        new TextPromptModal(this.app, `「${f.name}」的备注`, m.note, async (v) => {
          const mm = this.settings.archiveMeta[f.path] || (this.settings.archiveMeta[f.path] = { added: f.stat.ctime });
          mm.note = v;
          await this.plugin.saveSettings();
          this.render();
        }).open();
      });
      const acts = c.createDiv('ata-aacts');
      const go = acts.createEl('button', { cls: 'ata-abtn is-cta', text: '交给 AI', attr: { title: '选一张表，把这个文件作为附件，复制指令给外部 AI' } });
      go.addEventListener('click', (e) => { e.stopPropagation(); this.handoff([f]); });
      if (this.file) {
        const cur = acts.createEl('button', { cls: 'ata-abtn', text: '加到当前表', attr: { title: '作为附件加到「' + this.file.basename + '」' } });
        cur.addEventListener('click', (e) => { e.stopPropagation(); this.addAiAtts(this.file.path, [f]); this.openTable(this.file); });
      }
      const sp = c.createDiv('ata-atools');
      if (archiveKind(f) === 'image') this.iconBtn(sp, 'copy', '复制图片', () => copyImageNotice(this.app, f));
      this.iconBtn(sp, 'eye', '预览', () => new LightboxModal(this.app, this, f).open());
      if (this.app.openWithDefaultApp) this.iconBtn(sp, 'external-link', '用默认程序打开', () => this.app.openWithDefaultApp(f.path));
      this.iconBtn(sp, 'trash-2', '删除', () => new ConfirmModal(this.app, `删除「${f.name}」？`, '会移到回收站，需要时可以恢复。', [{ text: '删除', warning: true, onClick: () => this.archiveDelete([f]) }]).open()).addClass('is-danger');
      if (this.archSelect) {
        const box = c.createEl('input', { type: 'checkbox', cls: 'ata-card-check' });
        box.checked = this.archSel.has(f.path);
        const toggle = () => { if (this.archSel.has(f.path)) this.archSel.delete(f.path); else this.archSel.add(f.path); box.checked = this.archSel.has(f.path); c.toggleClass('is-selected', box.checked); if (this.archCountEl) this.archCountEl.setText(`已选 ${this.archSel.size} 个`); };
        box.addEventListener('click', (e) => { e.stopPropagation(); });
        box.addEventListener('change', () => { if (box.checked) this.archSel.add(f.path); else this.archSel.delete(f.path); c.toggleClass('is-selected', box.checked); if (this.archCountEl) this.archCountEl.setText(`已选 ${this.archSel.size} 个`); });
        c.addEventListener('click', toggle);
      }
    });
    if (list.length > 600) wrap.createDiv({ cls: 'ata-muted', text: `只显示最近 600 个，共 ${list.length} 个；可以用搜索找更早的文件。` });
    applyQ();
  }

  async archiveDelete(tfs) {
    this.archiving = true;
    try {
      for (const tf of tfs) {
        try {
          if (this.app.fileManager && this.app.fileManager.trashFile) await this.app.fileManager.trashFile(tf);
          else await this.app.vault.trash(tf, true);
          if (this.settings.archiveMeta) delete this.settings.archiveMeta[tf.path];
          if (this.archSel) this.archSel.delete(tf.path);
        } catch (e) { new Notice(`删除「${tf.name}」失败：${e.message || e}`); }
      }
    } finally { this.archiving = false; }
    await this.plugin.saveSettings();
    new Notice(`已删除 ${tfs.length} 个附件（在回收站里可以恢复）。`);
    this.logOp({ type: 'delete', text: `删除了 ${tfs.length} 个附件（在回收站）`, items: tfs.map((f) => f.name) });
    await this.render();
  }

  isPinned(path) { return (this.settings.pinned || []).includes(path); }

  async togglePin(file) {
    const list = this.settings.pinned || (this.settings.pinned = []);
    const i = list.indexOf(file.path);
    if (i >= 0) list.splice(i, 1); else list.push(file.path);
    await this.plugin.saveSettings();
    this.render();
  }

  recentFiles(limit) {
    const seen = new Set(), out = [];
    for (const p of [...(this.settings.pinned || []), ...(this.settings.recent || [])]) {
      if (seen.has(p)) continue;
      seen.add(p);
      const f = this.app.vault.getAbstractFileByPath(p);
      if (f instanceof TFile) out.push(f);
      if (out.length >= (limit || 12)) break;
    }
    return out;
  }

  uniquePathSync(folder, base, ext, taken) {
    let p = normalizePath(`${folder}/${base}.${ext}`), i = 1;
    while (taken.has(p.toLowerCase()) || this.app.vault.getAbstractFileByPath(p)) p = normalizePath(`${folder}/${base}-${i++}.${ext}`);
    taken.add(p.toLowerCase());
    return p;
  }

  async importBatch(list) {
    if (this.importRun) { new Notice('正在导入中，请等这一批完成或先取消。'); return; }
    const tables = list.filter((x) => TABLE_EXT.test(x.file.name) && !x.file.name.startsWith('~$'));
    const skipped = list.length - tables.length;
    if (!tables.length) { new Notice('没有找到可导入的表格文件（支持 csv / tsv / xlsx / md）。'); return; }
    const root = normalizePath(this.settings.folder);
    const run = this.importRun = { cancel: false, done: 0, total: tables.length, current: '', bytes: 0 };
    const totalBytes = tables.reduce((a, x) => a + (x.file.size || 0), 0);
    this.showImportBar(totalBytes);
    let ok = 0, empty = 0, lastFile = null;
    const failed = [], taken = new Set(), folders = new Map();
    const meta = this.settings.meta || (this.settings.meta = {});
    const ensure = (folder) => {
      if (!folders.has(folder)) folders.set(folder, this.ensureFolder(folder));
      return folders.get(folder);
    };
    let next = 0;
    const worker = async () => {
      while (!run.cancel) {
        const i = next++;
        if (i >= tables.length) return;
        const { file, rel } = tables[i];
        run.current = file.name;
        this.updateImportBar();
        try {
          const sheets = await fileToSheets(file);
          if (run.cancel) return;
          if (!sheets.length) empty++;
          else {
            const dirs = rel.split('/').slice(0, -1).map(safeSeg).filter(Boolean);
            const multi = sheets.length > 1;
            const book = baseName(file.name);
            const folder = [root, ...dirs, ...(multi ? [book] : [])].join('/');
            await ensure(folder);
            /* Excel：保存一份原文件和对应关系，导出时把改动写回原文件，颜色、公式、图片都保留 */
            let linkId = null;
            if (sheets[0].src) {
              try {
                linkId = Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
                const dir = `${this.ataDir()}/maps/${linkId}`;
                await this.ensureAdapterDir(dir);
                await this.ensureAdapterDir(`${this.ataDir()}/sources`);
                await this.app.vault.adapter.writeBinary(`${this.ataDir()}/sources/${linkId}.xlsx`, await file.arrayBuffer());
                const used = new Set();
                sheets.forEach((sh) => Object.values(sh.styles || {}).forEach((x) => used.add(x)));
                const xf = {};
                used.forEach((i) => { if (sheets.visual && sheets.visual[i]) xf[i] = sheets.visual[i]; });
                await this.app.vault.adapter.write(`${dir}/xf.json`, JSON.stringify(xf));
                await this.app.vault.adapter.write(`${dir}/book.json`, JSON.stringify({ name: file.name, imported: Date.now(), sheets: sheets.map((x) => ({ name: x.name, index: x.index })) }));
                for (const sh of sheets) {
                  await this.app.vault.adapter.write(`${dir}/${sh.index}.json`, JSON.stringify({
                    sheet: sh.name, headerRow: sh.src.headerRow, rows0: sh.src.rows, cols0: sh.src.cols, rows: sh.src.rows, cols: sh.src.cols, headers0: sh.table.headers,
                  }));
                  await this.app.vault.adapter.write(`${dir}/${sh.index}.styles.json`, JSON.stringify(sh.styles || {}));
                }
              } catch (e) {
                console.warn('[AI表格助手] 保存原 Excel 失败，导出时将不能保留原格式', e);
                linkId = null;
              }
            }
            for (const sh of sheets) {
              const path = this.uniquePathSync(folder, multi ? baseName(sh.name) : book, 'csv', taken);
              const csv = toCSV(sh.table);
              await this.snapPut(path, csv);
              lastFile = await this.app.vault.create(path, csv);
              const m = {};
              if (sh.title) m.title = sh.title;
              if (linkId) m.link = { id: linkId, sheet: sh.index };
              if (multi) Object.assign(m, { workbook: book, sheet: sh.name, sheetIndex: sh.index });
              if (Object.keys(m).length) meta[path] = Object.assign({}, meta[path], m);
              ok++;
              (run.names || (run.names = [])).push(`${multi ? book + ' · ' + sh.name : book}（${sh.table.rows.length} 行）`);
              if (multi) run.sheets = (run.sheets || 0) + 1;
            }
            if (multi) run.books = (run.books || 0) + 1;
          }
        } catch (e) {
          if (/没有找到 Markdown 表格/.test(e.message)) empty++;
          else failed.push(`${file.name}（${e.message}）`);
        }
        run.done++;
        run.bytes += file.size || 0;
        this.updateImportBar();
        await new Promise((r) => setTimeout(r, 0)); // 让界面喘口气
      }
    };
    try {
      await Promise.all(Array.from({ length: Math.min(4, tables.length) }, worker));
    } finally {
      this.importRun = null;
      this.hideImportBar();
      await this.plugin.saveSettings();
    }
    const msg = [`${run.cancel ? '已取消，' : ''}导入完成：${ok} 个表格${run.books ? `（其中 ${run.books} 个 Excel 共 ${run.sheets} 个工作表，每个工作表一张表）` : ''}`];
    if (empty) msg.push(`${empty} 个没有表格内容`);
    if (skipped) msg.push(`${skipped} 个非表格文件已跳过`);
    if (failed.length) msg.push(`${failed.length} 个失败：${failed.slice(0, 3).join('；')}${failed.length > 3 ? '…' : ''}`);
    new Notice(msg.join('，'), failed.length ? 12000 : 6000);
    if (ok || failed.length) this.logOp({ type: 'import', text: msg.join('，'), items: (run.names || []).slice(0, 300) });
    if (ok === 1 && lastFile) await this.openTable(lastFile);
    else {
      if (ok > 1) { this.tab = 'library'; this.libFilter = '未分类'; }
      await this.render();
    }
  }

  showImportBar(totalBytes) {
    this.hideImportBar();
    const bar = this.importBar = createDiv('ata-importbar');
    this.root.prepend(bar);
    this.ico(bar, 'loader', 'ata-spin');
    const mid = bar.createDiv('ata-importbar-mid');
    this.importText = mid.createDiv({ cls: 'ata-importbar-text' });
    const track = mid.createDiv('ata-progress');
    this.importFill = track.createDiv('ata-progress-fill');
    this.importTotalBytes = totalBytes || 1;
    this.btn(bar, '取消', 'x', () => { if (this.importRun) this.importRun.cancel = true; }, 'ata-btn-sm');
    this.updateImportBar();
  }

  updateImportBar() {
    const run = this.importRun;
    if (!run || !this.importText) return;
    this.importText.setText(`正在导入 ${run.done} / ${run.total}${run.current ? `：${run.current}` : ''}`);
    const pct = Math.max(run.done / run.total, run.bytes / this.importTotalBytes);
    this.importFill.style.width = Math.round(pct * 100) + '%';
  }

  hideImportBar() {
    if (this.importBar) this.importBar.remove();
    this.importBar = this.importText = this.importFill = null;
  }

  async createTable(name, cols, fmt, quiet) {
    try {
      const t = { headers: cols, rows: [] };
      const path = await this.uniquePath(this.settings.folder, baseName(name), fmt);
      const content = fmt === 'md' ? (cols.length ? toMarkdown(t) + '\n' : '') : cols.length ? toCSV(t) : '';
      await this.snapPut(path, content);
      const file = await this.app.vault.create(path, content);
      this.logOp({ type: 'create', path: file.path, text: `新建表格${cols.length ? `（${cols.slice(0, 6).join('、')}${cols.length > 6 ? '…' : ''}）` : ''}` });
      if (!quiet) await this.openTable(file);
      return file;
    } catch (e) {
      new Notice('新建失败：' + e.message);
      return null;
    }
  }

  async exportDialog(file) {
    const target = file || this.file;
    if (!target) { new Notice('没有可导出的表格。'); return; }
    let table = this.table;
    if (!this.file || target.path !== this.file.path) {
      const text = await this.app.vault.read(target);
      const m = target.extension === 'md' ? findMdTable(text) : null;
      table = target.extension === 'md' ? (m ? m.table : { headers: [], rows: [] }) : gridToTable(parseCSV(text));
    }
    if (!table.headers.length) { new Notice('这个表格是空的。'); return; }
    const sibs = this.workbookSiblings(target);
    const tm = (this.settings.meta || {})[target.path] || {};
    const book = sibs.length > 1 ? { name: tm.workbook, count: sibs.length } : null;
    const linked = tm.link && (await this.app.vault.adapter.exists(`${this.ataDir()}/sources/${tm.link.id}.xlsx`)) ? tm.link.id : null;
    new ExportModal(this.app, (fmt) => {
      if (fmt === 'book') this.exportWorkbook(target, sibs);
      else if (fmt === 'linked') this.exportLinked(target, book ? sibs : [target]);
      else this.exportTable(target, table, fmt);
    }, book, linked).open();
  }

  /** 把改动写回原 Excel：颜色、合并单元格、图片、公式、列宽、没改的工作表都保留 */
  async exportLinked(file, files) {
    const notice = new Notice('正在按原 Excel 格式导出…', 0);
    try {
      await this.flushSave();
      const meta = this.settings.meta || {};
      const id = meta[file.path].link.id;
      const src = await this.app.vault.adapter.readBinary(`${this.ataDir()}/sources/${id}.xlsx`);
      const sheets = [];
      for (const f of files) {
        const m = meta[f.path];
        if (!m || !m.link || m.link.id !== id) continue;
        let table, link;
        if (this.file && f.path === this.file.path) { table = this.table; link = this.link; }
        else ({ table, link } = await this.loadTableFor(f));
        if (link) sheets.push({ index: m.link.sheet, table, map: link.map });
      }
      const res = await patchWorkbook(src, sheets);
      let bookName = meta[file.path].workbook || file.basename;
      try { bookName = baseName(JSON.parse(await this.app.vault.adapter.read(`${this.ataDir()}/maps/${id}/book.json`)).name); } catch (e) { /* 用工作簿名 */ }
      const d = new Date(), z = (n) => String(n).padStart(2, '0');
      const stamp = `${d.getFullYear()}${z(d.getMonth() + 1)}${z(d.getDate())}-${z(d.getHours())}${z(d.getMinutes())}`;
      const path = await this.uniquePath(this.settings.folder + '/导出', `${bookName}-${stamp}`, 'xlsx');
      await this.app.vault.createBinary(path, res.buffer);
      notice.hide();
      const msg = `已按原格式导出到 ${path}：${res.sheetsChanged ? `${res.sheetsChanged} 个工作表有改动，共 ${res.changed} 个单元格` : '没有改动，和原文件一致'}` +
        (res.formulasDropped ? `；有 ${res.formulasDropped} 个共享公式因所在单元格被改动而转成了数值` : '');
      new Notice(msg, 10000);
      this.logOp({ type: 'export', path: file.path, name: bookName, text: `按原格式导出 Excel${res.sheetsChanged ? `（${res.sheetsChanged} 个工作表有改动，${res.changed} 个单元格）` : '（没有改动）'}`, file: path });
      if (this.settings.openAfterExport && this.app.openWithDefaultApp) {
        try { await this.app.openWithDefaultApp(path); } catch (e) { /* 打不开就只保存 */ }
      }
    } catch (e) {
      notice.hide();
      new Notice('按原格式导出失败：' + (e.message || e) + '。可以改用普通 Excel 导出。', 10000);
    }
  }

  async exportWorkbook(file, sibs) {
    try {
      await this.flushSave();
      const meta = this.settings.meta || {};
      const sheets = [];
      for (const f of sibs) {
        const t = this.file && f.path === this.file.path ? this.table : gridToTable(parseCSV(await this.app.vault.read(f)));
        sheets.push({ name: meta[f.path].sheet || f.basename, table: t });
      }
      const d = new Date(), z = (n) => String(n).padStart(2, '0');
      const stamp = `${d.getFullYear()}${z(d.getMonth() + 1)}${z(d.getDate())}-${z(d.getHours())}${z(d.getMinutes())}`;
      const path = await this.uniquePath(this.settings.folder + '/导出', `${meta[file.path].workbook}-${stamp}`, 'xlsx');
      await this.app.vault.createBinary(path, toXLSXBook(sheets));
      new Notice(`已导出整个工作簿（${sheets.length} 个工作表）到 ${path}`, 8000);
      this.logOp({ type: 'export', path: file.path, name: meta[file.path].workbook, text: `导出整个工作簿（${sheets.length} 个工作表）`, file: path });
      if (this.settings.openAfterExport && this.app.openWithDefaultApp) {
        try { await this.app.openWithDefaultApp(path); } catch (e) { /* 打不开就只保存 */ }
      }
    } catch (e) {
      new Notice('导出失败：' + e.message, 8000);
    }
  }

  async exportTable(file, table, fmt) {
    try {
      const d = new Date(), z = (n) => String(n).padStart(2, '0');
      const stamp = `${d.getFullYear()}${z(d.getMonth() + 1)}${z(d.getDate())}-${z(d.getHours())}${z(d.getMinutes())}`;
      const path = await this.uniquePath(this.settings.folder + '/导出', file.basename + '-' + stamp, fmt);
      if (fmt === 'xlsx') await this.app.vault.createBinary(path, toXLSX(table));
      else if (fmt === 'csv') await this.app.vault.create(path, '﻿' + toCSV(table));
      else await this.app.vault.create(path, toMarkdown(table) + '\n');
      new Notice('已导出到 ' + path, 8000);
      this.logOp({ type: 'export', path: file.path, text: `导出为 ${fmt.toUpperCase()}`, file: path });
      if (fmt === 'xlsx' && this.settings.openAfterExport && this.app.openWithDefaultApp) {
        try { await this.app.openWithDefaultApp(path); } catch (e) { /* 打不开就只保存 */ }
      }
    } catch (e) {
      new Notice('导出失败：' + e.message, 8000);
    }
  }

  /* ---------- 撤销 ---------- */
  pushUndo() {
    this.tableVersion = (this.tableVersion || 0) + 1;
    this.undoStack.push(cloneTable(this.table));
    if (this.undoStack.length > 30) this.undoStack.shift();
  }

  async undo() {
    if (!this.undoStack.length) { new Notice('没有可撤销的修改。'); return; }
    this.table = this.undoStack.pop();
    this.tableVersion = (this.tableVersion || 0) + 1;
    if (this.file) this.logOp({ type: 'undo', path: this.file.path, text: '撤销了上一步修改' });
    await this.saveTable();
    this.renderTable();
    if (this.undoBtn) this.undoBtn.disabled = !this.undoStack.length;
    new Notice('已撤销。');
  }
}

/* ====================================================================
 * 五、设置页 & 插件入口
 * ==================================================================== */
class AITableSettingTab extends PluginSettingTab {
  constructor(app, plugin) { super(app, plugin); this.plugin = plugin; }
  view() { const l = this.app.workspace.getLeavesOfType(VIEW_TYPE)[0]; return l && l.view instanceof TableWorkbenchView ? l.view : null; }
  display() {
    const { containerEl } = this;
    const s = this.plugin.settings;
    const save = () => this.plugin.saveSettings();
    containerEl.empty();

    new Setting(containerEl).setName('使用引导').setDesc('一步步介绍每个功能怎么用。')
      .addButton((b) => b.setButtonText('重新查看使用引导').onClick(async () => {
        this.app.setting.close();
        const v = await this.plugin.activateView();
        if (v) v.startTour();
      }));
    new Setting(containerEl).setName('表格保存文件夹').setDesc('导入、新建的表格放在这里；导出文件在其中的「导出」子文件夹，附件在「附件区」。外部 AI 需要能访问这个文件夹。')
      .addText((t) => t.setValue(s.folder).onChange(async (v) => { s.folder = v.trim() || DEFAULT_SETTINGS.folder; await save(); }));
    new Setting(containerEl).setName('给 AI 的说明文件').setDesc('插件会在表格文件夹里生成「AI使用说明.md」「表格目录.md」以及指向它们的 CLAUDE.md、AGENTS.md，告诉外部 AI 表格在哪、怎么改。每次打开工作台都会自动更新。')
      .addButton((b) => b.setButtonText('现在重新生成').onClick(async () => {
        const v = this.view() || (await this.plugin.activateView());
        if (v) { await v.writeAiGuide(); new Notice('已更新 AI 使用说明和表格目录。'); }
      }));
    new Setting(containerEl).setName('导出 Excel 后自动打开').setDesc('导出 .xlsx 后用系统默认程序（Excel / WPS）打开。')
      .addToggle((t) => t.setValue(s.openAfterExport !== false).onChange(async (v) => { s.openAfterExport = v; await save(); }));
  }
}

class AITableAssistantPlugin extends Plugin {
  async onload() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    if (!this.settings.meta) this.settings.meta = {};
    if (!Array.isArray(this.settings.pinned)) this.settings.pinned = [];
    if (!Array.isArray(this.settings.recent)) this.settings.recent = [];
    if (!this.settings.archiveMeta || typeof this.settings.archiveMeta !== 'object') this.settings.archiveMeta = {};
    /* 旧版本内置 AI 的设置（包括 API Key）已经不用了，从插件数据里清掉 */
    if (OBSOLETE_SETTINGS.some((k) => k in this.settings)) { OBSOLETE_SETTINGS.forEach((k) => delete this.settings[k]); await this.saveSettings(); }
    if (this.settings.master && !this.settings.pinned.includes(this.settings.master)) this.settings.pinned.push(this.settings.master);
    this.registerView(VIEW_TYPE, (leaf) => new TableWorkbenchView(leaf, this));
    this.addRibbonIcon('table', '表格工作台', () => this.activateView());
    this.addCommand({ id: 'open-panel', name: '打开表格工作台', callback: () => this.activateView() });
    this.addCommand({ id: 'tour', name: '查看使用引导', callback: () => this.activateView().then((v) => v && v.startTour()) });
    this.addCommand({ id: 'quick-switch', name: '快速切换表格', callback: () => this.activateView().then((v) => v && v.quickSwitch()) });
    this.addCommand({
      id: 'use-active-file',
      name: '在工作台中打开当前表格',
      checkCallback: (checking) => {
        const f = this.app.workspace.getActiveFile();
        if (!f || (f.extension !== 'md' && f.extension !== 'csv')) return false;
        if (!checking) this.activateView().then((v) => v && v.openTable(f));
        return true;
      },
    });
    /* 文件改名 / 删除时同步分类信息 */
    this.registerEvent(this.app.vault.on('rename', async (f, old) => {
      const m = this.settings.meta;
      if (m[old]) { m[f.path] = m[old]; delete m[old]; await this.saveSettings(); }
      const am = this.settings.archiveMeta || {};
      if (am[old]) { am[f.path] = am[old]; delete am[old]; await this.saveSettings(); }
      if (this.settings.lastTable === old) { this.settings.lastTable = f.path; await this.saveSettings(); }
    }));
    this.registerEvent(this.app.vault.on('delete', async (f) => {
      if (this.settings.meta[f.path]) { delete this.settings.meta[f.path]; await this.saveSettings(); }
    }));
    this.addSettingTab(new AITableSettingTab(this.app, this));
  }

  async saveSettings() { await this.saveData(this.settings); }

  async activateView() {
    const ws = this.app.workspace;
    let leaf = ws.getLeavesOfType(VIEW_TYPE)[0];
    /* 旧版本放在侧边栏，迁到主编辑区 */
    if (leaf && leaf.getRoot() !== ws.rootSplit) { leaf.detach(); leaf = null; }
    if (!leaf) {
      leaf = ws.getLeaf('tab');
      await leaf.setViewState({ type: VIEW_TYPE, active: true });
    }
    ws.revealLeaf(leaf);
    return leaf.view instanceof TableWorkbenchView ? leaf.view : null;
  }
}

AITableAssistantPlugin._test = {
  parseCSV, toCSV, gridToTable, findMdTable, toMarkdown, parseXLSX, toXLSX,
  normalizeTable, collectDropped, cleanGrid, decodeText, readWorkbook, toXLSXBook, fileToSheets, patchWorkbook, cloneTable, diffTables, hashStr,
};
module.exports = AITableAssistantPlugin;
