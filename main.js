'use strict';
/* 表格工作台 —— Obsidian 插件（无需构建，直接使用本文件） */
const {
  Plugin, ItemView, Notice, Modal, FuzzySuggestModal, PluginSettingTab,
  Setting, setIcon, normalizePath, TFile, requestUrl,
} = require('obsidian');

const VIEW_TYPE = 'ai-table-assistant-view';
const MAX_FILE = 200 * 1024 * 1024; // 导入 / 附件大小上限 200MB

/* 旧版本在表格文件夹里生成的说明文件（不是表格） */
const GUIDE_FILES = ['AI使用说明.md', '表格目录.md', 'CLAUDE.md', 'AGENTS.md'];
/* 旧版本用过、现在不用的设置 */
const OBSOLETE_SETTINGS = ['provider', 'apiKey', 'baseUrl', 'model', 'maxRows', 'preview', 'autoArchive'];

const OLD_FOLDER = 'AI表格助手';
const AUTO_DIR = '自动处理';
const DATA_DIR = '数据';
const OUT_DIR = '表格';
/* 插件在表格文件夹根目录生成的说明文件（不是表格） */
const DOC_FILES = ['操作守则.md', '数据目录.md', '操作日志.md', '新手教程.md'];
const DEFAULT_FOLDER = '表格工作台';
const DEFAULT_SETTINGS = {
  folder: DEFAULT_FOLDER,
  lastTable: '',
  master: '',
  pinned: [],
  recent: [],
  archiveMeta: {},
  openAfterExport: true,
  meta: {},
  scan: null,
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
  for (const c of deadCols) if (origAt(H, c)) set(H, c, '');
  const liveRows = new Set(table.rows.map((r) => r.__src).filter((x) => x != null && x >= 0));
  /* 删掉的列：清空原数据行里的值（非结构变化时在这里做） */
  const clearDeadCols = () => { for (const c of deadCols) for (const r of rows0) if (liveRows.has(r) && origAt(r, c)) set(r, c, ''); };
  const tmplCol = (map.cols0 || map.cols).length ? (map.cols0 || map.cols)[(map.cols0 || map.cols).length - 1] : 0;
  const newCols = new Set(target.filter((c, k) => colSrc[k] < 0));
  /* 插在中间的新行、删掉的行：要像 Excel 一样整行插入 / 删除，下面的行跟着移动 */
  const srcSeq = table.rows.map((r) => (r.__src != null && r.__src >= 0 ? r.__src : -1));
  const laterSrc = new Array(srcSeq.length).fill(false);
  for (let i = srcSeq.length - 2; i >= 0; i--) laterSrc[i] = laterSrc[i + 1] || srcSeq[i + 1] >= 0;
  const ins = new Map(); // 原表第几行（0 起）之后 → 插入的表格行
  let lastSrc = H;
  srcSeq.forEach((sr, i) => {
    if (sr >= 0) lastSrc = sr;
    else if (laterSrc[i]) { if (!ins.has(lastSrc)) ins.set(lastSrc, []); ins.get(lastSrc).push(i); }
  });
  const deleted = rows0.filter((r) => !liveRows.has(r)).sort((a, b) => a - b);
  if (ins.size || deleted.length) {
    const insKeys = [...ins.keys()].sort((a, b) => a - b);
    const insCum = [];
    insKeys.reduce((acc, k) => { acc += ins.get(k).length; insCum.push(acc); return acc; }, 0);
    const delSet = new Set(deleted);
    const countBelow = (arr, x) => { let lo = 0, hi = arr.length; while (lo < hi) { const m = (lo + hi) >> 1; if (arr[m] < x) lo = m + 1; else hi = m; } return lo; };
    /* r0（0 起）→ 新位置（0 起），被删 → -1 */
    const nr0 = (r0) => {
      if (delSet.has(r0)) return -1;
      const k = countBelow(insKeys, r0);
      return r0 - countBelow(deleted, r0) + (k ? insCum[k - 1] : 0);
    };
    const rowMap = (n1) => { const x = nr0(n1 - 1); return x < 0 ? 0 : x + 1; };
    const insRows = new Map(); // 新位置（0 起）→ 用来抄格式的行（新位置，0 起）
    const firstLive = srcSeq.find((x) => x >= 0);
    for (const [after, idxs] of ins) {
      const baseNew = nr0(after);
      const tmplNew = after === H && firstLive != null ? nr0(firstLive) + idxs.length : baseNew;
      idxs.forEach((ti, j) => {
        const R = baseNew + 1 + j;
        insRows.set(R, tmplNew > baseNew ? tmplNew : baseNew);
        table.rows[ti].forEach((v, k) => { v = cellStr(v); if (v !== '') set(R, target[k], v); });
      });
    }
    /* 原有行：按新位置比较 */
    for (const row of table.rows) {
      const R = row.__src;
      if (R == null || R < 0) continue;
      const N = nr0(R);
      row.forEach((v, k) => { v = cellStr(v); if (v !== origAt(R, target[k])) set(N, target[k], v); });
      for (const c of deadCols) if (origAt(R, c)) set(N, c, '');
    }
    /* 末尾的新行照旧接在最后（填进预先拉好的空行） */
    let maxNew = maxRow + insCum.reduce((a, b) => Math.max(a, b), 0) + 1;
    const newRows = [];
    srcSeq.forEach((sr, i) => {
      if (sr >= 0 || laterSrc[i]) return;
      const R = ++maxNew;
      newRows.push(R);
      table.rows[i].forEach((v, k) => { v = cellStr(v); if (v !== '') set(R, target[k], v); });
    });
    const lastLive = [...liveRows].sort((a, b) => b - a)[0];
    const tmplRow = lastLive != null ? nr0(lastLive) : H;
    const insAfter = new Map([...ins].map(([k, v]) => [k + 1, v.length]));
    const dataRows = new Set(rows0.map((r) => r + 1));
    return { changes, newRows: new Set(newRows), newCols, tmplRow, tmplCol, maxRow: maxNew, maxCol, headerRow: H, tableCols: target, rowMap, insRows, insAfter, dataRows, inserted: [...ins.values()].reduce((a, x) => a + x.length, 0), deleted: deleted.length };
  }
  clearDeadCols();
  /* 删掉的行：只在末尾时清空它的值（格式保留） */
  for (const r of deleted) for (const c of [...liveCols, ...deadCols]) if (origAt(r, c)) set(r, c, '');
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
  return { changes, newRows: new Set(newRows), newCols, tmplRow, tmplCol, maxRow, maxCol, headerRow: H, tableCols: target };
}

/* ====================================================================
 * 插入 / 删除行：像 Excel 一样把下面的行整体移动，公式、合并单元格、
 * 图片、批注、条件格式、数据验证、其它工作表的引用都跟着调整
 * ==================================================================== */
/** map(行号1起) → 新行号；0 表示这一行被删掉了 */
function makeRowShifter(map, insAfter, dataRows) {
  const near = (n, lim, step) => {
    for (let k = n, i = 0; step > 0 ? k <= lim : k >= lim; k += step, i++) {
      const m = map(k);
      if (m) return m;
      if (i > 5000) return map(lim) || 0;
    }
    return 0;
  };
  const range = (a, b) => {
    if (a > b) [a, b] = [b, a];
    let s = near(a, b, 1), e = near(b, a, -1);
    if (!(s && e && s <= e)) return null;
    /* 紧贴着数据区域的首尾插入新行时，把区域扩大到新行（合计、条件格式等会包含新数据） */
    if (insAfter && dataRows && b > a) {
      const kEnd = insAfter.get(b);
      if (kEnd && dataRows.has(b) && map(b)) e = map(b) + kEnd;
      const kStart = insAfter.get(a - 1);
      if (kStart && dataRows.has(a) && map(a)) s = map(a) - kStart;
    }
    return [s, e];
  };
  return { map, range };
}

const SHEET_PREFIX = "(?:'(?:[^']|'')+'|[^\\s!'\"(),:;=+\\-*/&<>^%{}\\[\\]]+)!";
const REF_PRE = '(^|[\\s(),;=+\\-*/&<>^%{}\\[\\]"])';
const REF_RE = new RegExp(REF_PRE + '(' + SHEET_PREFIX + ')?(\\$?)([A-Z]{1,3})(\\$?)(\\d+)(?::(\\$?)([A-Z]{1,3})(\\$?)(\\d+))?(?![0-9A-Za-z_(!])', 'g');
const ROWREF_RE = new RegExp(REF_PRE + '(' + SHEET_PREFIX + ')?(\\$?)(\\d+):(\\$?)(\\d+)(?![0-9A-Za-z_(!:])', 'g');
const unq = (p) => (p ? p.slice(0, -1).replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : '');

/** 改公式里的行号。own=true：没写工作表名的引用算这张表的；sheet：要改的那张表的名字 */
function remapFormula(text, sh, sheet, own) {
  if (!text) return text;
  const hit = (prefix) => (prefix ? unq(prefix) === sheet : own);
  return text.split(/("(?:[^"]|"")*"|&quot;[\s\S]*?&quot;)/).map((part, i) => {
    if (i % 2) return part;
    part = part.replace(REF_RE, (m, pre, px, d1, c1, d2, r1, d3, c2, d4, r2) => {
      if (!hit(px)) return m;
      px = px || '';
      if (r2 == null) { const n = sh.map(+r1); return n ? `${pre}${px}${d1}${c1}${d2}${n}` : `${pre}${px}#REF!`; }
      const rg = sh.range(+r1, +r2);
      return rg ? `${pre}${px}${d1}${c1}${d2}${rg[0]}:${d3}${c2}${d4}${rg[1]}` : `${pre}${px}#REF!`;
    });
    return part.replace(ROWREF_RE, (m, pre, px, d1, r1, d2, r2) => {
      if (!hit(px)) return m;
      const rg = sh.range(+r1, +r2);
      return rg ? `${pre}${px || ''}${d1}${rg[0]}:${d2}${rg[1]}` : `${pre}${px || ''}#REF!`;
    });
  }).join('');
}

/** XML 元素里的公式：先还原转义字符（中文表名常写成 &#35746; 这种），改完再转义回去 */
function remapXmlFormula(text, sh, sheet, own) {
  if (!text) return text;
  const raw = text.indexOf('&') >= 0 ? xmlUnescape(text.replace(/&amp;/g, '\u0000AMP\u0000')).replace(/\u0000AMP\u0000/g, '&') : text;
  const out = remapFormula(raw, sh, sheet, own);
  return out === raw ? text : xmlEscape(out);
}

/** 改 "A1:B5 C7" 这样的区域列表；返回 '' 表示全部被删掉 */
function remapSqref(sq, sh) {
  return String(sq).split(/\s+/).filter(Boolean).map((one) => {
    const m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(one.replace(/\$/g, ''));
    if (!m) return one;
    if (!m[3]) { const n = sh.map(+m[2]); return n ? `${m[1]}${n}` : ''; }
    const rg = sh.range(+m[2], +m[4]);
    return rg ? `${m[1]}${rg[0]}:${m[3]}${rg[1]}` : '';
  }).filter(Boolean).join(' ');
}

/** 工作表 XML 里 sheetData 以外的部分：合并单元格、条件格式、数据验证、超链接、筛选、选区、分页符… */
function remapSheetExtras(xml, sh) {
  const sdStart = xml.indexOf('<sheetData');
  const sdEndTag = xml.indexOf('</sheetData>');
  const sdEnd = sdEndTag >= 0 ? sdEndTag + 12 : xml.indexOf('>', sdStart) + 1;
  const fix = (s) => s
    .replace(/<mergeCell\b[^>]*\bref="([^"]+)"[^>]*\/>/g, (all, ref) => { const r = remapSqref(ref, sh); return r && r.includes(':') && !/^([A-Z]+)(\d+):\1\2$/.test(r) ? all.replace(ref, r) : ''; })
    .replace(/(<(?:conditionalFormatting|dataValidation|selection|ignoredError|protectedRange)\b[^>]*\bsqref=")([^"]+)(")/g, (all, a, ref, b) => a + (remapSqref(ref, sh) || 'A1') + b)
    .replace(/(<(?:hyperlink|autoFilter|dimension)\b[^>]*\bref=")([^"]+)(")/g, (all, a, ref, b) => a + (remapSqref(ref, sh) || ref) + b)
    .replace(/(\b(?:activeCell|topLeftCell)=")([A-Z]+)(\d+)(")/g, (all, a, c, r, b) => a + c + (sh.map(+r) || r) + b)
    .replace(/(<(?:formula|formula1|formula2)\b[^>]*>)([\s\S]*?)(<\/(?:formula|formula1|formula2)>)/g, (all, a, f, b) => a + remapXmlFormula(f, sh, '', true) + b)
    .replace(/<rowBreaks\b[\s\S]*?<\/rowBreaks>/g, (blk) => blk.replace(/(<brk\b[^>]*\bid=")(\d+)(")/g, (all, a, id, b) => a + (sh.map(+id) || id) + b));
  let out = fix(xml.slice(0, sdStart)) + xml.slice(sdStart, sdEnd) + fix(xml.slice(sdEnd));
  out = out.replace(/<mergeCells\b([^>]*)>([\s\S]*?)<\/mergeCells>/, (all, attrs, inner) => {
    const n = (inner.match(/<mergeCell\b/g) || []).length;
    return n ? `<mergeCells count="${n}">${inner}</mergeCells>` : '';
  });
  return out;
}

/** 图片 / 图表的位置（drawing XML，行号从 0 起） */
function remapDrawing(xml, sh) {
  return xml.replace(/<((?:\w+:)?)(twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/\1\2>/g, (anchor, ns) => {
    let fromRow = null, newFrom = null;
    const rowTag = (part) => new RegExp('(<' + ns + part + '>[\\s\\S]*?<' + ns + 'row>)(\\d+)(</' + ns + 'row>)');
    anchor = anchor.replace(rowTag('from'), (all, a, r, b) => {
      fromRow = +r;
      let n = sh.map(fromRow + 1);
      if (!n) { const rg = sh.range(fromRow + 1, fromRow + 200000); n = rg ? rg[0] : fromRow + 1; }
      newFrom = n - 1;
      return a + newFrom + b;
    });
    anchor = anchor.replace(rowTag('to'), (all, a, r, b) => {
      let n = sh.map(+r + 1);
      if (!n) n = (newFrom != null ? newFrom + (+r - fromRow) : +r) + 1;
      return a + Math.max(n - 1, newFrom || 0) + b;
    });
    return anchor;
  });
}

/** 旧式批注的显示框位置（VML，行号从 0 起） */
function remapVml(xml, sh) {
  return xml
    .replace(/(<x:Row>)(\d+)(<\/x:Row>)/g, (all, a, r, b) => a + ((sh.map(+r + 1) || +r + 1) - 1) + b)
    .replace(/(<x:Anchor>\s*)([^<]+)(<\/x:Anchor>)/g, (all, a, v, b) => {
      const p = v.split(',').map((x) => x.trim());
      if (p.length === 8) { p[2] = String((sh.map(+p[2] + 1) || +p[2] + 1) - 1); p[6] = String((sh.map(+p[6] + 1) || +p[6] + 1) - 1); }
      return a + p.join(', ') + b;
    });
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
  if (!changes.size && !ch.rowMap) return { xml, formulasDropped: 0 };
  const sh = ch.rowMap ? makeRowShifter(ch.rowMap, ch.insAfter, ch.dataRows) : null;
  if (sh) xml = remapSheetExtras(xml, sh);
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
  /* 整行插入 / 删除：行号、单元格地址、公式里的引用都换成新的 */
  const preKill = new Set();
  if (sh) {
    const moved = new Map(), mOrder = [];
    for (const rn of order) {
      const nr = sh.map(rn);
      const row = rows.get(rn);
      if (!nr) {
        for (const cx of row.cells.values()) { const f = /<f\b([^>]*)>/.exec(cx); if (f && /t="shared"/.test(f[1]) && /\bref="/.test(f[1])) { const si = (/\bsi="(\d+)"/.exec(f[1]) || [])[1]; if (si != null) preKill.add(si); } }
        continue;
      }
      const attrs = row.attrs.replace(/\br="\d+"/, `r="${nr}"`);
      const cells = new Map();
      for (const [c, cx] of row.cells) {
        let nx = cx.replace(/(<c\b[^>]*\br=")([A-Z]+)\d+(")/, `$1$2${nr}$3`);
        nx = nx.replace(/(<f\b)([^>]*?)(\/>|>([\s\S]*?)<\/f>)/, (all, a, fa, rest, text) => {
          fa = fa.replace(/\bref="([^"]+)"/, (m, ref) => `ref="${remapSqref(ref, sh) || ref}"`);
          if (text == null) return a + fa + rest;
          return `${a}${fa}>${remapXmlFormula(text, sh, ch.sheetName || '', true)}</f>`;
        });
        cells.set(c, nx);
      }
      moved.set(nr, { attrs, cells });
      mOrder.push(nr);
    }
    rows.clear();
    for (const [k, v] of moved) rows.set(k, v);
    order.length = 0;
    order.push(...mOrder);
  }
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
  if (!changes.size && !sh) return { xml, formulasDropped: 0 };

  /* 被改动单元格如果是共享公式的主格，整组共享公式转成数值，避免文件损坏 */
  const killSi = new Set(preKill);
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
      const insT = ch.insRows && ch.insRows.get(r0);
      const s = styleOf(old) || (insT != null ? styleOf((rows.get(insT + 1) || { cells: new Map() }).cells.get(c)) || tmplStyle(rn, c) : newRows.has(r0) || newCols.has(c) ? tmplStyle(rn, c) : null);
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
  const formulaFor = (cx, c, rn, tpRn) => {
    const f = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(cx || '');
    if (!f || /t="array"|t="dataTable"/.test(f[1])) return null;
    if (/t="shared"/.test(f[1])) {
      const si = (/\bsi="(\d+)"/.exec(f[1]) || [])[1];
      const m = masters.get(si);
      return m ? shiftFormula(m.text, rn - m.rn, c - m.c) : null;
    }
    return f[2] ? shiftFormula(f[2], rn - (tpRn || tmplRow + 1), 0) : null;
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
  const fillRow = (rn, tp, tpRn, merges1) => {
    let row = rows.get(rn);
    if (!row) { row = { attrs: ` r="${rn}"`, cells: new Map() }; rows.set(rn, row); order.push(rn); }
    if (tp && row.attrs.replace(/\s*\bspans="[^"]*"/, '') === ` r="${rn}"`) {
      row.attrs = ` r="${rn}"` + tp.attrs.replace(/\s*\br="\d+"/, '').replace(/\s*\bspans="[^"]*"/, '');
    }
    const cols = new Set([...(tp ? tp.cells.keys() : []), ...(ch.tableCols || [])]);
    for (const c of cols) {
      const curCx = row.cells.get(c);
      if (curCx && /<f\b/.test(curCx)) continue; // 本来就有公式
      const fx = tp && formulaFor(tp.cells.get(c), c, rn, tpRn); // 公式列：新行也用公式，结果由 Excel 算
      if (!fx && curCx && /<v>|<is>/.test(curCx)) continue;
      const ref = colLetter(c) + rn;
      const st = styleOf(curCx) || (tp && styleOf(tp.cells.get(c))) || tmplStyle(rn, c);
      if (fx) row.cells.set(c, `<c r="${ref}"${st ? ` s="${st}"` : ''}><f>${fx}</f></c>`);
      else if (st && !row.cells.has(c)) row.cells.set(c, `<c r="${ref}" s="${st}"/>`);
      if (c > maxC) maxC = c;
    }
    if (rn > maxR) maxR = rn;
    if (!mergeRows.has(rn)) merges1.forEach(([c1, c2]) => mergeAdd.push(`<mergeCell ref="${c1}${rn}:${c2}${rn}"/>`));
  };
  for (const r0 of newRows) fillRow(r0 + 1, tmpl, tmplRow + 1, tmplMerges);
  /* 插在中间的新行：照它上面那一行（插在表头下面时照下面一行）的格式和公式 */
  if (ch.insRows) {
    for (const [r0, t0] of ch.insRows) {
      const tp = rows.get(t0 + 1);
      const m1 = [];
      xml.replace(/<mergeCell\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g, (all, c1, r1, c2, r2) => { if (+r1 === t0 + 1 && +r2 === +r1) m1.push([c1, c2]); return ''; });
      fillRow(r0 + 1, tp, t0 + 1, m1);
    }
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
  let changed = 0, formulasDropped = 0, inserted = 0, deletedRows = 0;
  const shifts = [];
  for (const sh of sheets) {
    const meta = book.sheets[sh.index];
    if (!meta) continue;
    const orig = await readSheet(book, meta);
    const ch = sheetChanges(sh.table, sh.map, orig);
    if (!ch.changes.size && !ch.rowMap) continue;
    ch.sheetName = meta.name;
    const xml = await book.zip.text(meta.path);
    const r = patchSheetXml(xml, ch, book.dateStyle);
    replaced.set(meta.path, r.xml);
    for (const m of ch.changes.values()) changed += m.size;
    formulasDropped += r.formulasDropped;
    if (ch.rowMap) { shifts.push({ name: meta.name, path: meta.path, sh: makeRowShifter(ch.rowMap, ch.insAfter, ch.dataRows) }); inserted += ch.inserted || 0; deletedRows += ch.deleted || 0; }
  }
  /* 整行插入 / 删除后：图片、批注、表格对象、其它工作表和图表里引用这张表的地方也要跟着移动 */
  if (shifts.length) {
    const zip = book.zip;
    const textOf = async (p) => (replaced.has(p) ? replaced.get(p) : zip.has(p) ? zip.text(p) : null);
    const relsOf = (p) => p.replace(/([^/]+)$/, '_rels/$1.rels');
    const resolve = (base, target) => {
      if (target.startsWith('/')) return target.slice(1);
      const parts = base.split('/').slice(0, -1);
      for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); }
      return parts.join('/');
    };
    for (const sf of shifts) {
      const rels = await textOf(relsOf(sf.path));
      if (!rels) continue;
      const targets = [];
      rels.replace(/<Relationship\b[^>]*>/g, (tag) => {
        const type = (/\bType="([^"]+)"/.exec(tag) || [])[1] || '';
        const t = (/\bTarget="([^"]+)"/.exec(tag) || [])[1];
        if (t && !/TargetMode="External"/.test(tag)) targets.push({ type: type.split('/').pop(), path: resolve(sf.path, t) });
        return '';
      });
      for (const tg of targets) {
        const x = await textOf(tg.path);
        if (x == null) continue;
        if (tg.type === 'drawing') replaced.set(tg.path, remapDrawing(x, sf.sh));
        else if (tg.type === 'vmlDrawing') replaced.set(tg.path, remapVml(x, sf.sh));
        else if (tg.type === 'comments') replaced.set(tg.path, x.replace(/<comment\b[^>]*\bref="([A-Z]+)(\d+)"[^>]*>[\s\S]*?<\/comment>/g, (all, c, r) => { const n = sf.sh.map(+r); return n ? all.replace(`ref="${c}${r}"`, `ref="${c}${n}"`) : ''; }));
        else if (tg.type === 'table') replaced.set(tg.path, x.replace(/(<(?:table|autoFilter|sortState|sortCondition)\b[^>]*\bref=")([^"]+)(")/g, (all, a, ref, b) => a + (remapSqref(ref, sf.sh) || ref) + b).replace(/(<calculatedColumnFormula\b[^>]*>)([\s\S]*?)(<\/calculatedColumnFormula>)/g, (all, a, f, b) => a + remapXmlFormula(f, sf.sh, sf.name, false) + b));
      }
    }
    const mentions = (text) => {
      const dx = text.indexOf('&#') >= 0 ? text.replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&apos;/g, "'") : text.replace(/&apos;/g, "'");
      return shifts.filter((sf) => dx.includes(sf.name + '!') || dx.includes(sf.name.replace(/'/g, "''") + "'!"));
    };
    const remapQualified = (text, re) => {
      for (const sf of mentions(text)) text = text.replace(re, (all, a, f, b) => a + remapXmlFormula(f, sf.sh, sf.name, false) + b);
      return text;
    };
    for (const meta of book.sheets) {
      const x = await textOf(meta.path);
      if (x == null || !x.includes('!') || !mentions(x).length) continue;
      const y = remapQualified(x, /(<f\b[^>]*>)([\s\S]*?)(<\/f>)/g);
      if (y !== x) replaced.set(meta.path, y);
    }
    for (const n of zip.names.filter((n) => /^xl\/charts\/chart\d*\.xml$/.test(n))) {
      const x = await textOf(n);
      if (!x || !mentions(x).length) continue;
      const y = remapQualified(x, /(<c:f>)([\s\S]*?)(<\/c:f>)/g);
      if (y !== x) replaced.set(n, y);
    }
    for (const n of zip.names.filter((n) => /^xl\/pivotCache\/pivotCacheDefinition\d*\.xml$/.test(n))) {
      const x = await textOf(n);
      if (!x) continue;
      const y = x.replace(/<worksheetSource\b([^>]*)\/>/, (all, attrs) => {
        const sheet = xmlUnescape((/\bsheet="([^"]*)"/.exec(attrs) || [])[1] || '');
        const sf = shifts.find((z) => z.name === sheet);
        return sf ? all.replace(/\bref="([^"]+)"/, (m, ref) => `ref="${remapSqref(ref, sf.sh) || ref}"`) : all;
      });
      if (y !== x) replaced.set(n, y);
    }
    const wb0 = await textOf('xl/workbook.xml');
    if (wb0) { const y = remapQualified(wb0, /(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g); if (y !== wb0) replaced.set('xl/workbook.xml', y); }
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
    let wb = replaced.has('xl/workbook.xml') ? replaced.get('xl/workbook.xml') : await book.zip.text('xl/workbook.xml');
    if (/<calcPr\b/.test(wb)) wb = wb.replace(/<calcPr\b([^>]*?)(\/?)>/, (all, attrs, sl) => `<calcPr${attrs.replace(/\s*fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${sl}>`);
    else if (/<\/definedNames>/.test(wb)) wb = wb.replace('</definedNames>', '</definedNames><calcPr fullCalcOnLoad="1"/>');
    else wb = wb.replace('</sheets>', '</sheets><calcPr fullCalcOnLoad="1"/>');
    replaced.set('xl/workbook.xml', wb);
  }
  const out = await rezip(book.zip, replaced, removed);
  return { buffer: out, changed, formulasDropped, inserted, deletedRows, sheetsChanged: [...replaced.keys()].filter((k) => /^xl\/worksheets\/[^/]+\.xml$/.test(k)).length };
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

/* ====================================================================
 * 原表样式视图：按原 Excel 的列宽、行高、合并单元格、边框、字体、对齐显示
 * ==================================================================== */
const XL_LAYOUT_VER = 1;
const XL_BORDER = { thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', mediumDashed: '2px dashed', dotted: '1px dotted', hair: '1px dotted', double: '3px double', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dotted', mediumDashDotDot: '2px dotted', slantDashDot: '2px dashed' };
const BUILTIN_FMT = { 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 14: 'yyyy/m/d', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy/m/d h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 49: '@', 55: 'yyyy"年"m"月"', 56: 'm"月"d"日"', 57: 'yyyy"年"m"月"', 58: 'm"月"d"日"' };

/** 读出完整的单元格样式：底色、字体（颜色/粗斜体/下划线/字号/字体名）、边框、对齐、数字格式 */
async function readFullStyles(zip) {
  if (!zip.has('xl/styles.xml')) return [];
  const st = await zip.text('xl/styles.xml');
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
  const fmts = {};
  st.replace(/<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g, (_, id, code) => { fmts[+id] = xmlUnescape(code); return ''; });
  const fills = [];
  section('fills').replace(/<fill\b[^>]*>([\s\S]*?)<\/fill>|<fill\b[^>]*\/>/g, (all, inner) => {
    inner = inner || '';
    const pt = (/patternType="([^"]+)"/.exec(inner) || [])[1];
    fills.push(pt && pt !== 'none' ? resolveColor((/<fgColor\b[^>]*>/.exec(inner) || [])[0], theme) : null);
    return '';
  });
  const fonts = [];
  section('fonts').replace(/<font\b[^>]*>([\s\S]*?)<\/font>|<font\b[^>]*\/>/g, (all, inner) => {
    inner = inner || '';
    fonts.push({
      color: resolveColor((/<color\b[^>]*>/.exec(inner) || [])[0], theme),
      b: /<b\s*\/>|<b\s+val="(1|true)"/.test(inner),
      i: /<i\s*\/>|<i\s+val="(1|true)"/.test(inner),
      u: /<u\s*\/>|<u\s+val="(single|double)/.test(inner),
      s: /<strike\s*\/>|<strike\s+val="(1|true)"/.test(inner),
      sz: +((/<sz\s+val="([\d.]+)"/.exec(inner) || [])[1] || 11),
      name: xmlUnescape((/<name\s+val="([^"]*)"/.exec(inner) || [])[1] || ''),
    });
    return '';
  });
  const borders = [];
  section('borders').replace(/<border\b[^>]*>([\s\S]*?)<\/border>|<border\b[^>]*\/>/g, (all, inner) => {
    inner = inner || '';
    const side = (name) => {
      const m = new RegExp('<' + name + '\\b([^>]*?)(?:/>|>([\\s\\S]*?)</' + name + '>)').exec(inner);
      if (!m) return null;
      const style = (/style="([^"]+)"/.exec(m[1]) || [])[1];
      if (!style || !XL_BORDER[style]) return null;
      const col = resolveColor((/<color\b[^>]*>/.exec(m[2] || '') || [])[0], theme) || '000000';
      return XL_BORDER[style] + ' #' + col;
    };
    borders.push({ l: side('left') || side('start'), r: side('right') || side('end'), t: side('top'), b: side('bottom') });
    return '';
  });
  const out = [];
  section('cellXfs').replace(/<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g, (all, attrs, inner) => {
    const fill = fills[+((/fillId="(\d+)"/.exec(attrs) || [])[1] || 0)] || null;
    const font = fonts[+((/fontId="(\d+)"/.exec(attrs) || [])[1] || 0)] || {};
    const bd = borders[+((/borderId="(\d+)"/.exec(attrs) || [])[1] || 0)] || {};
    const fid = +((/numFmtId="(\d+)"/.exec(attrs) || [])[1] || 0);
    const al = (/<alignment\b([^>]*)/.exec(inner || '') || [])[1] || '';
    const x = {};
    if (fill) x.bg = '#' + fill;
    if (font.color && font.color !== '000000') x.fc = '#' + font.color;
    if (font.b) x.b = 1;
    if (font.i) x.i = 1;
    if (font.u) x.u = 1;
    if (font.s) x.s = 1;
    if (font.sz && font.sz !== 11) x.sz = font.sz;
    if (font.name) x.fn = font.name;
    ['l', 'r', 't', 'b'].forEach((k) => { if (bd[k]) x['b' + k] = bd[k]; });
    const ha = (/horizontal="([^"]+)"/.exec(al) || [])[1];
    const va = (/vertical="([^"]+)"/.exec(al) || [])[1];
    if (ha && ha !== 'general') x.ha = ha;
    x.va = va || 'bottom';
    if (/wrapText="(1|true)"/.test(al)) x.w = 1;
    const ind = +((/indent="(\d+)"/.exec(al) || [])[1] || 0);
    if (ind) x.ind = ind;
    const code = fmts[fid] != null ? fmts[fid] : BUILTIN_FMT[fid];
    if (code && code !== 'General') x.f = code;
    const kind = fmtKind(fid, fmts[fid]);
    if (kind) x.k = kind;
    out.push(x);
    return '';
  });
  return out;
}

/** 读一个工作表的版面：列宽、行高、合并、冻结、所有单元格的值和样式编号 */
async function readSheetLayout(book, sheet) {
  const { zip, dateStyle, shared } = book;
  const L = { defW: 64, defH: 20, colW: {}, colHide: [], rowH: {}, rowHide: [], merges: [], fr: 0, fc: 0, grid: true, v: {}, s: {}, maxRow: 0, maxCol: 0 };
  let headDone = false, lastVal = -1, lastValCol = -1, lastStyled = -1, lastStyledCol = -1, rowsSeen = 0;
  const colOf = (letters) => { let c = 0; for (const ch of letters) c = c * 26 + (ch.charCodeAt(0) - 64); return c - 1; };
  await eachXmlBlock(zip, sheet.path, 'row', (seg) => {
    if (!headDone) {
      const head = seg.slice(0, Math.max(0, seg.indexOf('<sheetData')) || seg.length);
      if (/showGridLines="(0|false)"/.test(head)) L.grid = false;
      const pane = /<pane\b([^>]*)/.exec(head);
      if (pane && /state="frozen/.test(pane[1])) {
        L.fc = Math.round(+((/xSplit="([\d.]+)"/.exec(pane[1]) || [])[1] || 0));
        L.fr = Math.round(+((/ySplit="([\d.]+)"/.exec(pane[1]) || [])[1] || 0));
      }
      const fmt = /<sheetFormatPr\b([^>]*)/.exec(head);
      if (fmt) {
        const dh = +((/defaultRowHeight="([\d.]+)"/.exec(fmt[1]) || [])[1] || 0);
        if (dh) L.defH = Math.round(dh * 4 / 3);
        const dw = +((/defaultColWidth="([\d.]+)"/.exec(fmt[1]) || [])[1] || 0);
        const bw = +((/baseColWidth="([\d.]+)"/.exec(fmt[1]) || [])[1] || 0);
        if (dw) L.defW = Math.round(dw * 7);
        else if (bw) L.defW = Math.ceil((bw * 7 + 5) / 8) * 8;
      }
      head.replace(/<col\b([^>]*)/g, (all, a) => {
        const min = +((/min="(\d+)"/.exec(a) || [])[1] || 0), max = +((/max="(\d+)"/.exec(a) || [])[1] || 0);
        const w = +((/width="([\d.]+)"/.exec(a) || [])[1] || 0);
        const hid = /hidden="(1|true)"/.test(a);
        for (let c = min - 1; c < Math.min(max, 1024); c++) {
          if (w) L.colW[c] = Math.round(w * 7);
          if (hid) L.colHide.push(c);
        }
        return '';
      });
      headDone = true;
    }
    if (seg.indexOf('<mergeCell') >= 0) {
      seg.replace(/<mergeCell\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g, (_, a, b, c, d) => { L.merges.push([+b - 1, colOf(a), +d - 1, colOf(c)]); return ''; });
    }
    const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
    let rm;
    while ((rm = rowRe.exec(seg))) {
      rowsSeen++;
      const ra = rm[1];
      const r = /\br="(\d+)"/.exec(ra) ? +/\br="(\d+)"/.exec(ra)[1] - 1 : rowsSeen - 1;
      const ht = /\bht="([\d.]+)"/.exec(ra);
      if (ht && /customHeight="(1|true)"/.test(ra)) L.rowH[r] = Math.round(+ht[1] * 4 / 3);
      else if (ht && Math.abs(+ht[1] * 4 / 3 - L.defH) > 2) L.rowH[r] = Math.round(+ht[1] * 4 / 3);
      if (/hidden="(1|true)"/.test(ra)) L.rowHide.push(r);
      if (!rm[2]) continue;
      const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let cm, auto = 0;
      while ((cm = cellRe.exec(rm[2]))) {
        const attrs = cm[1];
        const ref = /\br="([A-Z]+)\d*"/.exec(attrs);
        const c = ref ? colOf(ref[1]) : auto;
        auto = c + 1;
        if (c > 1023) continue;
        const sm = /\bs="(\d+)"/.exec(attrs);
        if (sm && +sm[1]) { L.s[r + ':' + c] = +sm[1]; if (r > lastStyled) lastStyled = r; if (c > lastStyledCol) lastStyledCol = c; }
        const body = cm[2];
        if (!body) continue;
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
            const kind = dateStyle[sm ? +sm[1] : 0];
            if (kind && val !== '' && !isNaN(+val)) val = excelDate(+val, kind);
            else val = tidyNumber(val);
          }
        }
        if (val === '') continue;
        L.v[r + ':' + c] = val;
        if (r > lastVal) lastVal = r;
        if (c > lastValCol) lastValCol = c;
      }
    }
  });
  L.maxRow = Math.min(Math.max(lastVal + 1, Math.min(lastStyled + 1, lastVal + 200), 1), 200000);
  L.maxCol = Math.min(Math.max(lastValCol + 1, Math.min(lastStyledCol + 1, lastValCol + 30), 1), 1024);
  L.merges = L.merges.filter(([r1, c1]) => r1 < L.maxRow && c1 < L.maxCol);
  return L;
}

/** 一次读完整本 Excel 所有工作表的版面 */
async function buildBookLayouts(buf) {
  const book = await openWorkbook(buf);
  const xf = await readFullStyles(book.zip);
  const sheets = [];
  for (const sh of book.sheets) sheets.push(await readSheetLayout(book, sh));
  return { xf, sheets };
}

/** 按 Excel 数字格式显示数值 */
function xlFormat(v, x) {
  if (v == null || v === '' || !x || !x.f) return v == null ? '' : String(v);
  const code = x.f;
  if (code === '@' || /^general$/i.test(code)) return String(v);
  if (x.k) return xlFormatDate(String(v), code);
  const n = Number(v);
  if (!/^\s*-?[\d.]+(e[-+]?\d+)?\s*$/i.test(String(v)) || !isFinite(n)) return String(v);
  const secs = code.split(';');
  let sec = secs[0], neg = false;
  if (n < 0 && secs[1] != null && secs[1] !== '') { sec = secs[1]; neg = true; }
  else if (n === 0 && secs[2] != null && secs[2] !== '') sec = secs[2];
  sec = sec.replace(/\[\$([^\]-]*)[^\]]*\]/g, '"$1"').replace(/\[[^\]]*\]/g, '');
  const lits = [];
  sec = sec.replace(/"([^"]*)"|\\(.)|_(.)|\*(.)/g, (m, q, esc, sp) => { lits.push(q != null ? q : esc != null ? esc : sp != null ? ' ' : ''); return String.fromCharCode(0xE000 + lits.length - 1); });
  const unlit = (t) => t.replace(/[\uE000-\uEFFF]/g, (ch) => lits[ch.charCodeAt(0) - 0xE000]);
  if (/E\+/i.test(sec)) return String(v);
  const num = /[#0?][#0?,]*(\.[#0?]*)?|\.[#0?]+/.exec(sec);
  if (!num) return unlit(sec);
  const pct = (sec.match(/%/g) || []).length;
  let x2 = neg ? Math.abs(n) : n;
  for (let i = 0; i < pct; i++) x2 *= 100;
  const dec = num[1] || '';
  const intPart = num[0].split('.')[0];
  const scaled = (intPart.match(/,+$/) || [''])[0].length;
  if (scaled) x2 /= Math.pow(1000, scaled);
  const minD = (dec.match(/0/g) || []).length, maxD = (dec.match(/[0#?]/g) || []).length;
  let str = Math.abs(x2).toFixed(maxD);
  if (maxD > minD) str = str.replace(new RegExp('0{1,' + (maxD - minD) + '}$'), '').replace(/\.$/, minD ? '.' : '');
  const ip0 = intPart.replace(/,+$/, '');
  if (/,/.test(ip0)) { const [ip, fp] = str.split('.'); str = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (fp != null ? '.' + fp : ''); }
  if (!/0/.test(ip0) && /^0\./.test(str)) str = str.slice(1);
  const sign = x2 < 0 && !neg ? '-' : '';
  return unlit(sec.slice(0, num.index) + sign + str + sec.slice(num.index + num[0].length));
}

function xlFormatDate(v, code) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(v.trim());
  const t = !m && /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(v.trim());
  if (!m && !t) return v;
  const Y = m ? +m[1] : 1900, M = m ? +m[2] : 1, D = m ? +m[3] : 1;
  const z = (x) => String(x).padStart(2, '0');
  const h = m ? +(m[4] || 0) : +t[1], mi = m ? +(m[5] || 0) : +t[2], se = m ? +(m[6] || 0) : +(t[3] || 0);
  if (/\[\$-F800\]/i.test(code)) return `${Y}年${M}月${D}日`;
  if (/\[\$-F400\]/i.test(code)) return `${z(h)}:${z(mi)}:${z(se)}`;
  const sec = code.split(';')[0].replace(/\[\$-[^\]]*\]/g, '').replace(/\[[^\]]*\]/g, (x) => (/^\[h+\]$/i.test(x) ? x.slice(1, -1) : ''));
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const WEEK = ['日', '一', '二', '三', '四', '五', '六'];
  const ampm = /AM\/PM|上午\/下午/i.test(sec);
  const toks = [];
  const re = /"([^"]*)"|\\(.)|(yyyy|yy|e|mmmm|mmm|mm|m|dddd|ddd|dd|d|aaaa|aaa|hh|h|ss|s|AM\/PM|上午\/下午|0+)|([\s\S])/gi;
  let k;
  while ((k = re.exec(sec))) toks.push(k[1] != null ? { lit: k[1] } : k[2] != null ? { lit: k[2] } : k[3] ? { tk: k[3] } : { lit: k[4] });
  const out = [];
  toks.forEach((tk, i) => {
    if (tk.lit != null) { out.push(tk.lit === '_' ? ' ' : tk.lit); return; }
    const t0 = tk.tk, lo = t0.toLowerCase();
    const prev = toks.slice(0, i).reverse().find((x) => x.tk), next = toks.slice(i + 1).find((x) => x.tk);
    const isMin = /^m{1,2}$/.test(lo) && ((prev && /^h/i.test(prev.tk)) || (next && /^s/i.test(next.tk)));
    if (lo === 'yyyy' || lo === 'e') out.push(String(Y));
    else if (lo === 'yy') out.push(z(Y % 100));
    else if (isMin) out.push(lo === 'mm' ? z(mi) : String(mi));
    else if (lo === 'mmmm' || lo === 'mmm') out.push(MON[M - 1]);
    else if (lo === 'mm') out.push(z(M));
    else if (lo === 'm') out.push(String(M));
    else if (lo === 'dddd' || lo === 'aaaa') out.push('星期' + WEEK[new Date(Date.UTC(Y, M - 1, D)).getUTCDay()]);
    else if (lo === 'ddd' || lo === 'aaa') out.push(WEEK[new Date(Date.UTC(Y, M - 1, D)).getUTCDay()]);
    else if (lo === 'dd') out.push(z(D));
    else if (lo === 'd') out.push(String(D));
    else if (lo === 'hh') out.push(z(ampm ? (h % 12 || 12) : h));
    else if (lo === 'h') out.push(String(ampm ? (h % 12 || 12) : h));
    else if (lo === 'ss') out.push(z(se));
    else if (lo === 's') out.push(String(se));
    else if (lo === 'am/pm') out.push(h < 12 ? 'AM' : 'PM');
    else if (t0 === '上午/下午') out.push(h < 12 ? '上午' : '下午');
    else out.push('');
  });
  return out.join('');
}

/* ---------- 表格小工具 ---------- */
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


/* ====================================================================
 * 附件 AI 扫描：识别单据上的日期，按时间分类
 * ==================================================================== */
const SCAN_PRESETS = {
  deepseek: { label: 'DeepSeek', format: 'openai', url: 'https://api.deepseek.com/v1', model: 'deepseek-flash' },
  qwen: { label: '通义千问（阿里云百炼）', format: 'openai', url: 'https://dashscope.aliyuncs.com/compatible-mode/v1', model: 'qwen-vl-max-latest' },
  doubao: { label: '豆包（火山方舟）', format: 'openai', url: 'https://ark.cn-beijing.volces.com/api/v3', model: '' },
  zhipu: { label: '智谱 GLM', format: 'openai', url: 'https://open.bigmodel.cn/api/paas/v4', model: 'glm-4v-flash' },
  openai: { label: 'OpenAI', format: 'openai', url: 'https://api.openai.com/v1', model: 'gpt-4o-mini' },
  claude: { label: 'Claude（Anthropic）', format: 'anthropic', url: 'https://api.anthropic.com/v1', model: 'claude-haiku-5-5' },
  custom: { label: '其他（OpenAI 兼容接口）', format: 'openai', url: '', model: '' },
};
const SCAN_DEFAULTS = { provider: 'qwen', url: SCAN_PRESETS.qwen.url, key: '', model: SCAN_PRESETS.qwen.model, by: 'month', auto: true, rename: true, maxTokens: 4000, nameTpl: '{日期}_{类型}_{对方}_{单号}' };
const scanCfg = (s) => Object.assign({}, SCAN_DEFAULTS, s.scan || {});
const scanReady = (s) => { const c = scanCfg(s); return !!(c.key && c.url && c.model); };
const scannable = (f) => { const k = archiveKind(f); return (k === 'image' && !/^(svg|heic)$/i.test(f.extension || '')) || k === 'pdf'; };
/** 年月日文件夹（2026-11、2026-11/11-15）——AI 整理时会移动；自己建的分类文件夹不动 */
const isDateFolder = (fd) => fd === '' || /^\d{4}-\d{2}(\/\d{2}-\d{2})?$/.test(fd);

/** 把各种写法的日期统一成 YYYY-MM-DD 或 YYYY-MM；认不出返回 '' */
function normDate(v) {
  if (v == null) return '';
  const t = String(v).trim().replace(/[年月.\/]/g, '-').replace(/[日号]/g, '').replace(/\s+/g, '');
  const m = t.match(/^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?/);
  if (!m) return '';
  const y = +m[1], mo = +m[2], d = m[3] ? +m[3] : 0;
  if (y < 2000 || y > new Date().getFullYear() + 1 || mo < 1 || mo > 12 || d > 31) return '';
  const z = (n) => String(n).padStart(2, '0');
  return d ? `${y}-${z(mo)}-${z(d)}` : `${y}-${z(mo)}`;
}

function parseScanReply(text) {
  const t = String(text || '').replace(/```(?:json)?/gi, '');
  const a = t.indexOf('{'), b = t.lastIndexOf('}');
  if (a < 0 || b < a) throw new Error('AI 没有按格式回答（回复可能被截断，可以在设置里调高回复长度上限）：' + t.slice(0, 60));
  const o = JSON.parse(t.slice(a, b + 1));
  const str = (x) => (x == null || x === 'null' ? '' : String(x).trim()).slice(0, 60);
  return { date: normDate(o.date), type: str(o.type), party: str(o.party), no: str(o.no), summary: str(o.summary) };
}

/** 图片缩小到长边 1600px 再发送（省流量、省钱，识别日期足够清楚） */
async function imageToDataUrl(app, tf) {
  if (archiveKind(tf) === 'pdf') return (await renderPdfPages(app, tf, 1400, 1)).urls[0];
  const buf = await app.vault.readBinary(tf);
  const ext = (tf.extension || '').toLowerCase();
  const blob = new Blob([buf], { type: ext === 'png' ? 'image/png' : ext === 'webp' ? 'image/webp' : ext === 'gif' ? 'image/gif' : ext === 'bmp' ? 'image/bmp' : 'image/jpeg' });
  const url = URL.createObjectURL(blob);
  try {
    const img = await new Promise((ok, bad) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => bad(new Error('图片打不开（格式不支持或文件损坏）')); i.src = url; });
    const k = Math.min(1, 1600 / Math.max(img.naturalWidth || 1, img.naturalHeight || 1));
    const cv = document.createElement('canvas');
    cv.width = Math.max(1, Math.round(img.naturalWidth * k));
    cv.height = Math.max(1, Math.round(img.naturalHeight * k));
    const ctx = cv.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, cv.width, cv.height);
    ctx.drawImage(img, 0, 0, cv.width, cv.height);
    return cv.toDataURL('image/jpeg', 0.85);
  } finally { URL.revokeObjectURL(url); }
}

/** 调用视觉模型；image 为空时只发文字（测试连接用） */
async function callVision(cfg, prompt, image) {
  const base = String(cfg.url || '').trim().replace(/\/+$/, '').replace(/\/(chat\/completions|messages)$/, '');
  const anth = (SCAN_PRESETS[cfg.provider] || {}).format === 'anthropic';
  let url, headers, body;
  if (anth) {
    const content = [];
    if (image) content.push({ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: image.split(',')[1] } });
    content.push({ type: 'text', text: prompt });
    url = base + '/messages';
    headers = { 'x-api-key': cfg.key, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' };
    body = { model: cfg.model, max_tokens: cfg.maxTokens || 4000, messages: [{ role: 'user', content }] };
  } else {
    const content = [{ type: 'text', text: prompt }];
    if (image) content.push({ type: 'image_url', image_url: { url: image } });
    url = base + '/chat/completions';
    headers = { Authorization: 'Bearer ' + cfg.key, 'Content-Type': 'application/json' };
    body = { model: cfg.model, temperature: 0, max_tokens: cfg.maxTokens || 4000, messages: [{ role: 'user', content }] };
  }
  for (let attempt = 0; ; attempt++) {
    let r;
    try { r = await requestUrl({ url, method: 'POST', headers, body: JSON.stringify(body), throw: false }); } catch (e) {
      if (attempt < 2) { await new Promise((ok) => setTimeout(ok, 1500 * (attempt + 1))); continue; }
      throw new Error('连不上 API：' + (e.message || e));
    }
    let j = null;
    try { j = JSON.parse(r.text); } catch (e) { /* 不是 JSON */ }
    if ((r.status === 429 || r.status >= 500) && attempt < 3) { await new Promise((ok) => setTimeout(ok, 2500 * (attempt + 1))); continue; }
    if (r.status >= 400) {
      const msg = j && j.error ? (j.error.message || j.error.code || JSON.stringify(j.error)) : String(r.text || '').slice(0, 120);
      throw new Error(`API 返回 ${r.status}：${msg}${r.status === 401 || r.status === 403 ? '（检查 API Key）' : r.status === 404 ? '（检查 API 地址和模型名）' : ''}`);
    }
    let text = '';
    if (anth) text = ((j && j.content) || []).filter((c) => c.type === 'text').map((c) => c.text).join('');
    else { const c = j && j.choices && j.choices[0] && j.choices[0].message ? j.choices[0].message.content : ''; text = Array.isArray(c) ? c.map((x) => x.text || '').join('') : String(c || ''); }
    return text;
  }
}

function scanPrompt(tf, received) {
  return `这是收到的一份文件（文件名：${tf.name}，收到日期：${received}）。请找出文件上的业务日期——开单日期、送货日期、发票日期、合同签订日期等，以单据本身写的日期为准，不是拍照或打印时间；手写的日期也要认。
只写了月和日、没写年份的，按收到日期推断年份（不会晚于收到日期太多）。完全看不出日期就写 null。
只输出一个 JSON，不要其它文字：
{"date":"YYYY-MM-DD（只看得出月份就写 YYYY-MM）","type":"文件类型，例如 送货单、发票、收据、对账单、合同、聊天截图、其他","party":"对方单位或客户名称，没有写空字符串","no":"单号，没有写空字符串","summary":"不超过 20 字的说明"}`;
}

/** 按模板生成新文件名：{日期} {类型} {对方} {单号} {说明} {原名}；空的字段连同多余的分隔符一起去掉 */
function scanFileName(tpl, r, orig) {
  const map = { 日期: r.date, 类型: r.type, 对方: r.party, 单号: r.no, 说明: r.summary, 原名: orig };
  let out = String(tpl || SCAN_DEFAULTS.nameTpl).replace(/\{(日期|类型|对方|单号|说明|原名)\}/g, (_, k) => safeSeg(map[k] || '').replace(/[_\s]+/g, ' ').trim());
  out = out.replace(/([_\-\s·]){2,}/g, '$1').replace(/^[_\-\s·]+|[_\-\s·]+$/g, '');
  return safeSeg(out).slice(0, 80);
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
    new Setting(this.contentEl).setName('列名').setDesc('用逗号分隔，例如：日期, 摘要, 科目, 借方, 贷方。可以留空，之后再加列。')
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
    this.contentEl.createEl('p', { cls: 'setting-item-description', text: 'Excel 里表头上方常有「XX 年 XX 月对账单」之类的标题行。选中真正的表头所在的那一行，它上面的内容会保存为表格说明。' });
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
      const fdz = el.closest('[data-drop-folder]');
      if (fdz) return 'folder:' + fdz.getAttribute('data-drop-folder');
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
      const card = e.target && e.target.closest ? e.target.closest('[data-drop-table],[data-drop-folder]') : null;
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
      console.log('[表格工作台] drop', { types, files: dt.files ? dt.files.length : 0, text: text.slice(0, 300) });
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
    /* 首页「收到单据」：任何文件都存进附件区 */
    this.attachInput = this.makeInput({ multiple: 'true' }, (files) => this.receiveAtts(files.map((f) => ({ file: f, rel: f.name }))));

    /* 监视表格文件：被其它程序改了，就对比出改动，标出来让用户确认或撤销 */
    this.registerEvent(this.app.vault.on('modify', (f) => {
      if (this.snapReady && this.isTableFile(f)) this.checkFile(f, 'modify');
      this.onVaultChange(f);
    }));
    this.registerEvent(this.app.vault.on('create', (f) => {
      const ak = this.snapReady && this.autoKind(f);
      if (ak) this.queueAuto(f, ak);
      if (this.snapReady && !this.importRun && this.isTableFile(f)) setTimeout(() => this.checkFile(f, 'create'), 300);
      this.onVaultChange(f);
    }));
    this.registerEvent(this.app.vault.on('delete', (f) => {
      if (this.file && f.path === this.file.path) { this.file = null; this.table = { headers: [], rows: [] }; }
      this.onExtDelete(f);
      this.onVaultChange(f);
    }));
    this.registerEvent(this.app.vault.on('rename', (f, old) => {
      this.onExtRename(f, old);
      const ak = this.snapReady && this.autoKind(f);
      if (ak) this.queueAuto(f, ak);
      this.scheduleRender();
    }));

    /* 旧版本的表格文件夹叫「AI表格助手」，整体改名为「表格工作台」（连同隐藏的 .ata 数据一起搬） */
    try {
      const v = this.app.vault;
      if (!v.getAbstractFileByPath(normalizePath(this.settings.folder)) && v.getAbstractFileByPath(OLD_FOLDER)) this.settings.folder = OLD_FOLDER;
      if (normalizePath(this.settings.folder) === OLD_FOLDER && !v.getAbstractFileByPath(DEFAULT_FOLDER)) await this.migrateFolder(DEFAULT_FOLDER);
    } catch (e) { console.warn('[表格工作台] 文件夹改名失败', e); new Notice('表格文件夹改名失败，继续使用原来的文件夹：' + (e.message || e)); }
    try { await this.rebuildFromMaps(); } catch (e) { console.warn('[表格工作台] 恢复 Excel 对应关系失败', e); }
    if (this.plugin.restoredFromBackup) new Notice(`插件设置丢失了，已从表格文件夹里的备份恢复 ${this.plugin.restoredFromBackup} 张表的分组和格式信息。`, 10000);
    const last = this.settings.lastTable && this.app.vault.getAbstractFileByPath(this.settings.lastTable);
    if (last instanceof TFile) await this.loadFile(last, { silent: true, switchTab: false });
    await this.render();
    this.register(() => this.endTour(false));
    if (!this.settings.onboarded) setTimeout(() => this.startTour(), 600);
    /* 找出工作台关着时被外部改过的表；清掉旧版本生成的说明文件 */
    setTimeout(async () => {
      try { await this.scanExternal(); } catch (e) { console.warn('[表格工作台] 检查外部修改失败', e); this.snapReady = true; }
      try { await this.removeOldGuides(); } catch (e) { /* 不影响使用 */ }
      try { await this.moveIntoData(); } catch (e) { console.warn('[表格工作台] 搬进数据文件夹失败', e); }
      if (this.settings.autoProcess !== false) {
        try { await this.ensureAutoDirs(); this.scanAuto(); } catch (e) { console.warn('[表格工作台] 准备自动处理文件夹失败', e); }
      }
      try { await this.writeRules(); await this.writeTutorial(); await this.writeDataIndex(); await this.writeLogFile(); } catch (e) { console.warn('[表格工作台] 写说明文件失败', e); }
      this.checkOutputs();
    }, 300);
  }

  /* ================= 新手引导 ================= */
  tourSteps() {
    const toTab = (t) => async () => { if (this.tab !== t) { this.tab = t; await this.render(); } };
    return [
      { kind: 'welcome' },
      { sel: '.ata-tabs', title: '三个页面', before: toTab('home'),
        text: '「日志」记录每一次操作；「文件管理」管理所有表格文件；「附件区」存单据、截图、PDF。修改表格交给外部 AI，这里负责存放、展示和记录。' },
      { sel: '.ata-top-actions', title: '导入表格', before: toTab('home'),
        text: '点「导入文件」或「导入文件夹」把 Excel / CSV 导进来，也可以直接拖进面板。导入时自动转换成 AI 能快速读写的底层数据（「数据」文件夹），同时生成一份原表 Excel（「表格」文件夹）。' },
      { sel: '.ata-oplog', title: '日志', before: toTab('home'),
        text: 'AI 改了哪张表、哪些格子从什么改成什么、加了哪些行、原表是否已更新，都按时间记在这里；改错了可以点「撤销这次修改」。AI 写的处理说明也会出现在这里。' },
      { sel: '.ata-header-right', title: '操作守则', before: toTab('home'),
        text: '「操作守则」是写给 AI 看的规则：文件放在哪、怎么改、改完怎么记录。让 AI 干活时，点「复制 AI 开场白」粘贴给它，后面接上你的需求就行。完整用法看「新手教程」。' },
      { sel: '.ata-fmlist', title: '文件管理', before: toTab('library'), optional: true,
        alt: { sel: '.ata-tabs', text: '导入表格后，「文件管理」里会列出每个表格文件：勾选后可以复制原表地址、复制数据地址、删除，也能看到原表 Excel 是不是最新的。' },
        text: '每个表格文件一行：勾选后可以批量复制原表地址、复制底层数据地址、删除；右边能看到原表 Excel 是不是已经更新到最新。点多工作表的 Excel 可以展开看每张表。' },
      { sel: '.ata-folderbar', title: '附件区', before: toTab('archive'), optional: true,
        alt: { sel: '.ata-tabs', text: '「附件区」存收到的文件，可以新建文件夹分类，多选后批量复制文件或路径。' },
        text: '收到的文件存在这里。可以新建文件夹分类，点「多选」后批量复制文件（直接粘贴到微信）、复制路径、移动到文件夹。' },
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
    card.createDiv({ cls: 'ata-tour-text', text: '这里负责存放、展示和记录；改表格、算数交给你自己的 AI（能读写电脑文件的那种）。导入的 Excel 会转换成 AI 能快速读写的底层数据，AI 改完后自动生成新的原表 Excel，覆盖上一版。' });
    const flow = card.createDiv('ata-tour-flow');
    [['folder-down', '导入 Excel'], ['database', '转成底层数据'], ['bot', 'AI 读写数据'], ['file-spreadsheet', '自动生成原表'], ['scroll-text', '记录日志']].forEach(([ic, tx], i) => {
      if (i) flow.createSpan({ cls: 'ata-tour-arrow', text: '→' });
      const c = flow.createDiv('ata-tour-flowitem');
      setIcon(c.createDiv('ata-tour-flowicon'), ic);
      c.createSpan({ text: tx });
    });
    const box = card.createDiv('ata-tour-setup');
    box.createDiv({ cls: 'ata-tour-sub', text: '给 AI 的第一句话' });
    const guide = this.absPath(normalizePath(`${this.settings.folder}/操作守则.md`));
    const say = `请先阅读 ${guide}，以后按里面的规则帮我处理表格。`;
    const pv = box.createEl('code', { cls: 'ata-tour-path', text: say });
    pv.addEventListener('click', () => this.copyText(say, '这句话'));
    box.createDiv({ cls: 'ata-tour-hint', text: '点上面这句话可以复制，发给你的 AI。之后直接跟它说要做什么就行，比如「把附件区里今天的送货单加进订单表」。' });
    const bar = card.createDiv('ata-tour-bar');
    const skip = bar.createEl('button', { cls: 'ata-tour-skip', text: '跳过' });
    skip.addEventListener('click', () => this.endTour(true));
    const go = bar.createDiv('ata-tour-btns').createEl('button', { cls: 'mod-cta', text: '看看三个页面 →' });
    go.addEventListener('click', () => this.tourGo(1));
  }

  tourDone(card) {
    card.addClass('is-wide');
    card.createDiv({ cls: 'ata-tour-title big', text: '可以开始了 🎉' });
    card.createDiv({ cls: 'ata-tour-text', text: '先导入你平时用的 Excel，然后在「日志」页点「复制 AI 开场白」发给外部 AI，让它试着加一行数据，回来看日志和原表。怎么跟 AI 说、哪些 AI 能用，都写在「新手教程」里。' });
    card.createDiv({ cls: 'ata-tour-hint', text: '以后想再看：顶部的「教程」打开新手教程，「引导」重看这个界面介绍。' });
    const bar = card.createDiv('ata-tour-bar');
    const bs = bar.createDiv('ata-tour-btns');
    const tut = bs.createEl('button', { text: '打开新手教程' });
    tut.addEventListener('click', () => { this.endTour(true); this.openTutorial(); });
    const done = bs.createEl('button', { cls: 'mod-cta', text: '完成' });
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
      await this.openTable(file);
      new Notice('示例表已经建好，可以直接点格子修改。');
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

  /** 粘贴 / 按路径收到的文件：全是表格或整个文件夹就导入，其它文件存进附件区 */
  async receive(list) {
    list = (list || []).filter(Boolean);
    if (!list.length) { new Notice('没有读到文件。'); return; }
    const isFolder = list.some((x) => x.rel.includes('/'));
    if (isFolder || list.every((x) => TABLE_EXT.test(x.file.name))) { await this.importBatch(list); return; }
    await this.receiveAtts(list);
  }

  /** 存进附件区 */
  async receiveAtts(list) {
    const saved = await this.archiveSave(list.map((x) => x.file || x).filter((f) => f && f.size > 0), '收到');
    if (!saved.length) new Notice('没有读到文件内容。');
  }

  /** 旧版本在表格文件夹里生成的说明文件，现在不需要了（只删插件自己生成的，移到回收站） */
  async removeOldGuides() {
    const folder = normalizePath(this.settings.folder);
    for (const name of GUIDE_FILES) {
      const f = this.app.vault.getAbstractFileByPath(`${folder}/${name}`);
      if (!(f instanceof TFile)) continue;
      const text = await this.app.vault.read(f);
      if (!text.startsWith('<!-- 由「AI 表格助手」插件自动生成')) continue;
      if (this.app.fileManager && this.app.fileManager.trashFile) await this.app.fileManager.trashFile(f);
      else await this.app.vault.trash(f, true);
    }
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
    banner.createSpan({ cls: 'ata-muted', text: this.tab === 'archive' ? '存进附件区' + (this.archFolder ? `「${this.archFolder}」` : '') : 'Excel / CSV = 导入表格；文件夹 = 整个导入；其它文件 = 存进附件区' });

    this.renderTopbar(lib);
    this.renderTabs(lib);
    const body = this.root.createDiv('ata-body');
    this.renderedTab = this.tab;
    if (this.tab === 'workspace') this.tab = 'library';
    if (this.tab === 'library') this.renderLibrary(body, lib);
    else if (this.tab === 'archive') this.renderArchive(body);
    else this.renderHome(body, lib);
    this.root.scrollTop = scroll;
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
    const input = search.createEl('input', { cls: 'ata-search-input', attr: { type: 'text', placeholder: '搜索表格名称、工作表、列名…（回车在文件管理里定位）' } });
    input.value = this.query;
    const drop = search.createDiv('ata-search-drop');
    const update = () => {
      this.query = input.value.trim();
      const ql = this.query.toLowerCase();
      this.root.querySelectorAll('.ata-fmrow[data-q]').forEach((r) => { r.style.display = !ql || r.dataset.q.includes(ql) ? '' : 'none'; });
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
    top('graduation-cap', '教程', () => this.openTutorial());
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
    tab('home', 'scroll-text', '日志');
    tab('library', 'folder-tree', '文件管理', this.fmGroups(lib).length);
    tab('archive', 'archive', '附件区', this.archiveFiles().length);
  }


  /* ---------- 首页 ---------- */





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
    this.footEl = tablePane.createDiv('ata-tablefoot');
    this.renderWorkspaceFoot();
    this.renderTable();

  }

  renderWorkspaceFoot() {
    const foot = this.footEl;
    if (!foot || !foot.isConnected) return;
    foot.empty();
    if (this.link && this.file && this.file.extension === 'csv') {
      const seg = foot.createDiv('ata-seg');
      const ex = this.isExcelMode();
      [['excel', '原表样式', 'sheet'], ['data', '编辑数据', 'pencil']].forEach(([k, label, ic]) => {
        const b = seg.createEl('button', { cls: 'ata-seg-btn' + ((k === 'excel') === ex ? ' is-active' : '') });
        setIcon(b.createSpan('ata-seg-ico'), ic);
        b.createSpan({ text: label });
        b.addEventListener('click', () => { if ((k === 'excel') !== this.isExcelMode()) this.setViewMode(k); });
      });
    }
    if (this.isExcelMode()) {
      const zm = foot.createDiv('ata-zoom');
      const set = (v) => { this.xlZoom = Math.min(2, Math.max(0.5, Math.round(v * 10) / 10)); this.renderWorkspaceFoot(); this.renderTable(); };
      this.iconBtn(zm, 'zoom-out', '缩小', () => set((this.xlZoom || 1) - 0.1));
      const lab = zm.createEl('button', { cls: 'ata-zoom-lab', text: Math.round((this.xlZoom || 1) * 100) + '%', attr: { title: '恢复 100%' } });
      lab.addEventListener('click', () => set(1));
      this.iconBtn(zm, 'zoom-in', '放大', () => set((this.xlZoom || 1) + 0.1));
    } else {
      this.btn(foot, '新增行', 'plus', () => this.addRow(), 'ata-btn-sm');
      this.btn(foot, '新增列', 'columns-3', () => this.addCol(), 'ata-btn-sm');
      this.btn(foot, '设表头行', 'heading', () => this.pickHeaderRow(), 'ata-btn-sm').title = '表头不对时（上面有标题行），把某一行设为表头';
    }
    this.statEl = foot.createDiv({ cls: 'ata-muted ata-spacer-left' });
  }

  /** 工作区高度按窗口实际剩余空间计算，保证表格底部的工作表标签和按钮栏都看得到 */
  fitWorkspace() {
    const ws = this.wsEl;
    if (!ws || !ws.isConnected || this.tab !== 'workspace') return;
    if (this.root.clientWidth < 560) { ws.style.height = ''; return; } // 很窄时不限高
    const top = ws.getBoundingClientRect().top - this.root.getBoundingClientRect().top + this.root.scrollTop;
    const pad = parseFloat(window.getComputedStyle(this.root).paddingBottom) || 0;
    ws.style.height = Math.max(380, this.root.clientHeight - top - Math.min(pad, 24) - 4) + 'px';
  }

  onResize() { this.fitWorkspace(); if (this.isExcelMode() && this.xlRedraw) this.xlRedraw(); }

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

  /** 把整个表格文件夹改名，所有记录（分组、常用、附件备注、快照、操作记录）跟着改 */
  async migrateFolder(newName) {
    const s = this.settings, oldF = normalizePath(s.folder), newF = normalizePath(newName);
    if (oldF === newF) return false;
    const fo = this.app.vault.getAbstractFileByPath(oldF);
    if (!fo || fo instanceof TFile) return false;
    if (this.app.vault.getAbstractFileByPath(newF) || (await this.app.vault.adapter.exists(newF))) return false;
    await this.snapIndex();
    await this.loadOpLog();
    await this.flushSave();
    try { await (this.outQ || Promise.resolve()); } catch (e) { /* 继续 */ }
    const mv = (p) => (typeof p === 'string' && (p === oldF || p.startsWith(oldF + '/')) ? newF + p.slice(oldF.length) : p);
    const remap = (o) => { const out = {}; Object.keys(o || {}).forEach((k) => { out[mv(k)] = o[k]; }); return out; };
    this.migrating = true;
    try {
      await this.app.fileManager.renameFile(fo, newF);
    } finally { this.migrating = false; }
    s.meta = remap(s.meta);
    s.archiveMeta = remap(s.archiveMeta);
    s.pinned = (s.pinned || []).map(mv);
    s.recent = (s.recent || []).map(mv);
    s.lastTable = mv(s.lastTable);
    s.master = mv(s.master);
    s.folder = newF;
    this.snapIdx = remap(this.snapIdx);
    if (this.snapText) this.snapText = new Map([...this.snapText].map(([k, v]) => [mv(k), v]));
    if (this.hlMap) this.hlMap = new Map([...this.hlMap].map(([k, v]) => [mv(k), v]));
    (this.plugin.opLog || []).forEach((e) => { e.path = mv(e.path); if (e.file) e.file = mv(e.file); if (e.hist) e.hist = mv(e.hist); });
    this.stats = new Map();
    if (this.file) { const f = this.app.vault.getAbstractFileByPath(mv(this.file.path)); this.file = f instanceof TFile ? f : null; }
    await this.plugin.saveSettings();
    this.saveSnapIndex();
    this.saveOpLog();
    new Notice(`表格文件夹已改名为「${newF}」，所有表格和记录都跟着搬过去了。`, 8000);
    return true;
  }

  /** 插件设置里缺了 Excel 工作簿的对应关系（比如换了电脑），按 .ata 里保存的原 Excel 信息补回来 */
  async rebuildFromMaps() {
    const a = this.app.vault.adapter, dir = this.ataDir() + '/maps';
    if (!(await a.exists(dir))) return;
    const ids = (await a.list(dir)).folders || [];
    const books = [];
    for (const d of ids) {
      try { const b = JSON.parse(await a.read(d + '/book.json')); books.push({ id: d.split('/').pop(), ...b }); } catch (e) { /* 不完整，跳过 */ }
    }
    books.sort((x, y) => (x.imported || 0) - (y.imported || 0));
    const meta = this.settings.meta || (this.settings.meta = {});
    const root = normalizePath(this.settings.folder);
    const csvs = this.app.vault.getFiles().filter((f) => f.extension === 'csv' && f.path.startsWith(root + '/'));
    let fixed = 0;
    for (const b of books) {
      if (!(await a.exists(`${this.ataDir()}/sources/${b.id}.xlsx`))) continue;
      const book = baseName(b.name || '');
      const sheets = b.sheets || [];
      const multi = sheets.length > 1;
      for (const sh of sheets) {
        const want = multi ? `/${book}/${baseName(sh.name)}.csv` : `/${book}.csv`;
        const f = csvs.find((x) => ('/' + x.path).endsWith(want));
        if (!f) continue;
        const m = meta[f.path] || (meta[f.path] = {});
        if (m.link && m.link.id === b.id) continue;
        m.link = { id: b.id, sheet: sh.index };
        if (multi) Object.assign(m, { workbook: book, sheet: sh.name, sheetIndex: sh.index });
        fixed++;
      }
    }
    if (fixed) { await this.plugin.saveSettings(); new Notice(`已恢复 ${fixed} 张表和原 Excel 的对应关系。`); }
  }

  /* ================= 数据（底层 CSV）与原表（自动生成的 Excel） ================= */
  dataDir() { return normalizePath(this.settings.folder + '/' + DATA_DIR); }
  outDir() { return normalizePath(this.settings.folder + '/' + OUT_DIR); }

  /** 一张表所属的「文件」：Excel 工作簿（多张表）或单独一张表 */
  groupFiles(file) {
    const sibs = this.workbookSiblings(file);
    return sibs.length ? sibs : [file];
  }

  /** 这张表（或它所在工作簿）对应的原表 Excel 放在哪：表格/ 下按数据/ 的目录结构 */
  outPathOf(file) {
    const data = this.dataDir() + '/';
    const m = (this.settings.meta || {})[file.path] || {};
    const book = m.workbook && this.workbookSiblings(file).length;
    let rel = book ? (file.parent ? file.parent.path : '') : file.path.replace(/\.[^./]+$/, '');
    if (rel.startsWith(data)) rel = rel.slice(data.length);
    else rel = rel.split('/').pop() || file.basename;
    return normalizePath(`${this.outDir()}/${rel}.xlsx`);
  }

  /** 改动后自动重新生成原表（防抖，排队，一次一个） */
  scheduleOutput(file, delay) {
    if (!(file instanceof TFile) || !this.isTableFile(file)) return;
    if (!this.outTimers) this.outTimers = new Map();
    if (!this.outState) this.outState = new Map();
    const dest = this.outPathOf(file);
    clearTimeout(this.outTimers.get(dest));
    this.outState.set(dest, { status: 'pending', time: Date.now() });
    this.outTimers.set(dest, setTimeout(() => {
      this.outTimers.delete(dest);
      this.outQ = (this.outQ || Promise.resolve()).then(() => this.writeOutput(file, dest, 0)).catch((e) => console.warn('[表格工作台] 生成原表失败', e));
    }, delay == null ? 2500 : delay));
    if (this.tab === 'library') this.scheduleRender();
  }

  async writeOutput(file, dest, attempt) {
    const cur = this.app.vault.getAbstractFileByPath(file.path);
    if (!(cur instanceof TFile)) { this.outState.delete(dest); return; }
    file = cur;
    const realDest = this.outPathOf(file);
    if (realDest !== dest) { this.outState.delete(dest); dest = realDest; }
    const meta = this.settings.meta || {};
    const tm = meta[file.path] || {};
    const sibs = this.workbookSiblings(file);
    const opts = { quiet: true, silent: true, dest, nolog: true };
    try {
      await this.ensureFolder(dest.replace(/\/[^/]+$/, ''));
      const linked = tm.link && (await this.app.vault.adapter.exists(`${this.ataDir()}/sources/${tm.link.id}.xlsx`));
      let res;
      if (linked) res = await this.exportLinked(file, sibs.length > 1 ? sibs : [file], opts);
      else if (sibs.length > 1) res = await this.exportWorkbook(file, sibs, opts);
      else {
        const text = await this.app.vault.read(file);
        const t = file.extension === 'md' ? ((findMdTable(text) || {}).table || { headers: [], rows: [] }) : gridToTable(parseCSV(text));
        res = await this.exportTable(file, t, 'xlsx', opts);
      }
      this.outState.set(dest, { status: 'ok', time: Date.now() });
      /* 不单独记一条，标在引起这次更新的那条记录上 */
      const paths = new Set(this.groupFiles(file).map((x) => x.path));
      const log = this.plugin.opLog || [];
      for (let i = log.length - 1; i >= 0 && i >= log.length - 50; i--) {
        const e = log[i];
        if (Date.now() - e.t > 30 * 60000) break;
        if (paths.has(e.path) && /^(ext|ext-new|undo|edit|row|col)$/.test(e.type)) { e.output = Date.now(); e.outFile = dest; break; }
      }
      this.saveOpLog();
      this.scheduleDocs();
    } catch (e) {
      const msg = e.message || String(e);
      const busy = /EBUSY|EPERM|locked|被占用|resource busy/i.test(msg);
      if (attempt < 6) {
        this.outState.set(dest, { status: busy ? 'busy' : 'pending', time: Date.now(), msg });
        setTimeout(() => { this.outQ = this.outQ.then(() => this.writeOutput(file, dest, attempt + 1)); }, busy ? 20000 : 5000);
      } else {
        this.outState.set(dest, { status: 'error', time: Date.now(), msg });
        this.logOp({ type: 'output', path: file.path, text: `原表更新失败：${busy ? '文件被 Excel 或其它程序打开着，关掉后点「立即更新原表」' : msg}`, error: true });
      }
    }
    if (this.tab === 'library') this.scheduleRender();
  }

  /** 打开时检查：原表还没生成、或者比数据旧的，都重新生成 */
  checkOutputs() {
    const seen = new Set();
    for (const f of this.tableFiles()) {
      const dest = this.outPathOf(f);
      if (seen.has(dest)) continue;
      seen.add(dest);
      const out = this.app.vault.getAbstractFileByPath(dest);
      const newest = Math.max(...this.groupFiles(f).map((x) => x.stat.mtime));
      if (!(out instanceof TFile) || out.stat.mtime + 1000 < newest) this.scheduleOutput(f, 1500 + seen.size * 200);
    }
  }

  /** 把旧版本放在表格文件夹根目录的表搬进「数据」文件夹 */
  async moveIntoData() {
    const root = normalizePath(this.settings.folder), data = this.dataDir();
    const skip = new Set([DATA_DIR, OUT_DIR, '附件区', '导出', AUTO_DIR]);
    const tops = new Set();
    this.app.vault.getFiles().forEach((f) => {
      if (!f.path.startsWith(root + '/')) return;
      const rest = f.path.slice(root.length + 1), seg = rest.split('/')[0];
      if (skip.has(seg) || seg.startsWith('.')) return;
      if (!rest.includes('/')) { if (f.extension === 'csv' || (f.extension === 'md' && this.isTableFile(f) && !DOC_FILES.includes(seg))) tops.add(seg); }
      else if (f.extension === 'csv') tops.add(seg);
    });
    if (!tops.size) return 0;
    await this.ensureFolder(data);
    await this.snapIndex();
    await this.loadOpLog();
    await this.flushSave();
    try { await (this.outQ || Promise.resolve()); } catch (e) { /* 继续 */ }
    const pairs = [];
    this.migrating = true;
    try {
      for (const seg of tops) {
        const from = `${root}/${seg}`, to = `${data}/${seg}`;
        const fo = this.app.vault.getAbstractFileByPath(from);
        if (!fo || this.app.vault.getAbstractFileByPath(to)) continue;
        await this.app.fileManager.renameFile(fo, to);
        pairs.push([from, to]);
      }
    } finally { this.migrating = false; }
    const mv = (p) => { if (typeof p !== 'string') return p; for (const [a, b] of pairs) if (p === a || p.startsWith(a + '/')) return b + p.slice(a.length); return p; };
    await this.remapAll(mv);
    if (pairs.length) new Notice(`已把 ${pairs.length} 个表格 / 工作簿搬进「${DATA_DIR}」文件夹，原表 Excel 会自动生成在「${OUT_DIR}」文件夹。`, 10000);
    return pairs.length;
  }

  /** 文件搬家后，把所有记录里的路径跟着改 */
  async remapAll(mv) {
    const s = this.settings;
    const remap = (o) => { const out = {}; Object.keys(o || {}).forEach((k) => { out[mv(k)] = o[k]; }); return out; };
    s.meta = remap(s.meta);
    s.archiveMeta = remap(s.archiveMeta);
    s.pinned = (s.pinned || []).map(mv);
    s.recent = (s.recent || []).map(mv);
    s.lastTable = mv(s.lastTable);
    s.master = mv(s.master);
    if (this.snapIdx) this.snapIdx = remap(this.snapIdx);
    if (this.snapText) this.snapText = new Map([...this.snapText].map(([k, v]) => [mv(k), v]));
    if (this.hlMap) this.hlMap = new Map([...this.hlMap].map(([k, v]) => [mv(k), v]));
    (this.plugin.opLog || []).forEach((e) => { e.path = mv(e.path); if (e.file) e.file = mv(e.file); if (e.hist) e.hist = mv(e.hist); });
    this.stats = new Map();
    if (this.file) { const f = this.app.vault.getAbstractFileByPath(mv(this.file.path)); this.file = f instanceof TFile ? f : null; }
    await this.plugin.saveSettings();
    this.saveSnapIndex();
    this.saveOpLog();
  }

  /* ================= 给 AI 看的守则、数据目录、操作日志 ================= */
  async writeDoc(name, text) {
    const p = normalizePath(`${this.settings.folder}/${name}`);
    const f = this.app.vault.getAbstractFileByPath(p);
    try {
      if (f instanceof TFile) { if ((await this.app.vault.read(f)).replace(/\r/g, '') !== text) await this.app.vault.modify(f, text); }
      else { await this.ensureFolder(this.settings.folder); await this.app.vault.create(p, text); }
    } catch (e) { console.warn('[表格工作台] 写文件失败', name, e); }
  }

  /** 发给外部 AI 的开场白：告诉它文件夹在哪、先读守则 */
  aiOpener() {
    const root = this.absPath(normalizePath(this.settings.folder));
    const sep = root.includes('\\') ? '\\' : '/';
    return `你可以读写这个文件夹：${root}
开始前先完整阅读 ${root}${sep}操作守则.md，严格按里面的规则做：表格只改「数据」文件夹里的 CSV，不要动「表格」文件夹里的 Excel；做完在「自动处理${sep}记录」里写一条记录。
我的需求是：`;
  }

  async writeTutorial() {
    const root = this.absPath(normalizePath(this.settings.folder));
    const sep = root.includes('\\') ? '\\' : '/';
    const opener = this.aiOpener().split('\n').map((l) => '> ' + l).join('\n');
    const text = `<!-- 表格工作台自动生成，插件更新时会刷新，手动修改会被覆盖 -->
# 表格工作台 · 新手教程

**一句话：Obsidian 负责存放、展示和记录，外部 AI 负责干活。**

你把 Excel 导进来，插件会把它拆成 AI 读得快的底层数据（CSV）。你让 AI（Claude、ChatGPT、DeepSeek……）去改这些数据，插件几秒内自动生成和原表一模一样的 Excel（颜色、公式、合并单元格、图片都保留），并把每一次改动记进日志。

\`\`\`
你的 Excel ──导入──▶ 数据/（CSV，AI 读写这里）──自动生成──▶ 表格/（原生 Excel，拿去用）
                                    │
                                    └──▶ 日志：改了哪些格子、加了哪些行，可撤销
\`\`\`

你的表格文件夹：\`${root}\`

---

## 第一步：导入你的表格（1 分钟）

1. 点左边栏的表格图标，打开「表格工作台」。
2. 点顶部「导入文件」，选你平时用的 Excel（.xlsx / .xlsm / .csv）。也可以直接把文件拖进面板，或点「导入文件夹」一次导入一整个文件夹。
3. 导入完成后去「文件管理」看看：每个工作簿一行，点开能看到每张工作表；右边能看到原表 Excel 是不是最新的。

> 老版 .xls 需要先在 Excel 里「另存为」.xlsx。

导入后会多出这些东西：

| 位置 | 是什么 | 谁来用 |
|---|---|---|
| \`数据${sep}\` | 底层数据，每张工作表一个 CSV | **AI 改这里** |
| \`表格${sep}\` | 自动生成的原生 Excel，和原表格式一样 | **你拿去用、发给别人** |
| \`附件区${sep}\` | 你存的单据照片、截图、PDF | 你和 AI 都可以看 |
| \`操作守则.md\` | 写给 AI 的规则 | AI 开工前先读 |
| \`数据目录.md\` | 所有表格的清单（列名、行数、样例） | AI 用来找表 |
| \`操作日志.md\` | 所有操作记录 | 你核对用 |

---

## 第二步：让外部 AI 来干活

### 哪些 AI 可以用

**最好用：能直接读写电脑文件的 AI。** 例如 Claude 桌面版（把文件夹添加给它）、Claude Code、Codex、Cursor、Trae 等。把上面的表格文件夹交给它，它就能直接改，改完 Excel 自动更新。

**只能聊天的网页 AI 也能用，但要手动搬文件**（见本节最后）。

### 给 AI 的开场白（每次新对话先发这段）

在「日志」页点 **「复制 AI 开场白」**，粘贴给 AI，后面接上你的需求：

${opener}

开场白告诉 AI 三件事：文件在哪、先读操作守则、做完要记录。AI 读完守则就知道该怎么改、改哪里、不能动什么。

### 可以这样跟 AI 说（直接照抄改一改）

**录入数据**
- 「看附件区 2026-11 文件夹里今天的送货单照片，把每张单子的日期、客户、品名、规格、数量、单价录进『送货记录』表，金额自己算好。看不清的地方列出来让我核对。」
- 「在『！！客户』表里 11 月 15 日那几行后面插入一行：2026-11-15，华盛布行，棉布 280 克，200 米，单价 8.5。」（插在中间也可以，原表会像 Excel 一样整行插入）

**修改**
- 「把『价格表』里所有华盛布行的单价改成 8.5，金额重新算。」
- 「删除『送货记录』里单号 0011047 那一行。」

**只查不改**
- 「统计 10 月每个客户的金额合计，只告诉我结果，不要改任何文件。」
- 「找出『！！客户』表里金额和数量×单价对不上的行。」

**导入新表**
- 「把 D:\\下载\\11月对账单.xlsx 复制到 自动处理${sep}导入 里导入，然后告诉我生成了哪些表。」

### AI 改完之后你看哪里

1. **日志页**：最上面就是刚才的改动——哪张表、哪几个格子从什么改成什么、加了哪几行、原表 Excel 有没有更新，还有 AI 写的处理说明。
2. **原表 Excel**：在「文件管理」里勾选这张表 →「复制原表地址」，或点那一行的文件夹图标「在文件夹中显示」，打开就是最新版。
3. **改错了**：在日志里那一条上点 **「撤销这次修改」**，数据和原表 Excel 都会恢复。

### 只能聊天的网页 AI 怎么用

1. 在「文件管理」里勾选表格 →「复制数据地址」，找到对应的 CSV 文件，上传给 AI（大表可以只发需要的部分）。
2. 让 AI 按「操作守则」的 CSV 规则修改，**给你返回完整的 CSV 文件**。
3. 把 AI 给的 CSV 放回 \`数据${sep}\` 里原来的位置覆盖原文件，插件会自动生成 Excel、记录日志。

---

## 第三步：附件区和 AI 扫描

- 收到的单据照片、截图、PDF 拖进「附件区」，或者在附件区按 Ctrl+V 粘贴。
- 点「多选」可以批量复制文件（直接粘贴到微信）、复制路径、移动到文件夹、删除。
- **AI 扫描**：在 设置 → 表格工作台 →「附件 AI 扫描」里选服务商（DeepSeek、通义千问、豆包、智谱、OpenAI、Claude）、填 API Key、点「测试」。之后：
  - 新存进来的图片 / PDF 会自动识别单据上的日期，放进按月（或按日）分的文件夹；
  - 自动改名成「2026-11-15_送货单_对方名称_单号」这样的名字；
  - 附件区按日期分组显示，一眼能看出每天收到哪些单子；
  - 之前存的旧附件，点附件区右上角「AI 按时间整理」一次处理完。

整理好日期以后，跟外部 AI 说「把附件区 2026-11 里 11-15 那天的送货单录进表」就很准确了。

---

## 常见问题

**原表 Excel 没更新？**
插件只在 Obsidian 开着的时候工作，所以平时让 Obsidian 在后台开着。没开时改的，下次打开会自动补上。如果那个 Excel 正在 Excel / WPS 里打开，文件被占用，插件会等你关掉后再写——关掉 Excel 就会更新。

**可以让 AI 排序或整表重写吗？**
不要。排序会让原表里每一行的颜色、格式对不上号。插入、删除、修改单行都没问题。

**可以直接改「表格」文件夹里的 Excel 吗？**
不要，它会被自动覆盖。要改就改「数据」里的 CSV（或者让 AI 改）。原始 Excel 有了新版本，用「自动处理${sep}导入${sep}替换」重新导入。

**换电脑怎么办？**
把整个 Obsidian 库文件夹复制过去，**包括里面隐藏的 \`.ata\` 文件夹**（存着原表格式和修改历史），再装上插件即可。

**API Key 安全吗？**
只保存在这台电脑这个库的插件设置里，图片只发给你选的服务商。表格内容不会经过插件发给任何 AI——表格是外部 AI 自己读的。

**想再看一遍界面介绍？**
点工作台顶部的「引导」。
`;
    await this.writeDoc('新手教程.md', text);
  }

  async openTutorial() {
    await this.writeTutorial();
    const f = this.app.vault.getAbstractFileByPath(normalizePath(`${this.settings.folder}/新手教程.md`));
    if (f instanceof TFile) await this.app.workspace.getLeaf('tab').openFile(f);
  }

  async writeRules() {
    const root = normalizePath(this.settings.folder);
    const abs = this.absPath(root);
    const sep = abs.includes('\\') ? '\\' : '/';
    const text = `<!-- 表格工作台自动生成，插件更新时会刷新，手动修改会被覆盖 -->
# 操作守则（给处理这些表格的 AI 和程序看）

这个文件夹由 Obsidian 的「表格工作台」插件管理。你（AI）负责按用户的指令读取、修改表格；插件负责把你改好的数据自动转换成原生 Excel，并记录每一次操作。**开始任何任务前，先读完这份守则。**

## 1. 文件夹结构

根目录：\`${abs}\`

| 位置 | 是什么 | 你能不能改 |
|---|---|---|
| \`数据目录.md\` | 所有表格的清单：文件位置、列名、行数、样例 | 只读，先看它找表 |
| \`数据${sep}\` | **表格的底层数据**，每个工作表一个 CSV 文件；多工作表的 Excel 是一个子文件夹 | **读写这里** |
| \`表格${sep}\` | 自动生成的原生 Excel（和数据一一对应，保留原表的颜色、公式、合并单元格、图片） | 只读，不要改，会被覆盖 |
| \`附件区${sep}\` | 用户存的单据、截图、PDF、价格表。开了 AI 扫描时按单据日期放在 \`年-月\`（或 \`年-月${sep}月-日\`）文件夹里，文件名是「日期_类型_对方_单号」 | 只读 |
| \`自动处理${sep}\` | 导入新 Excel、请求导出、写操作记录（见第 4、5 节） | 按说明使用 |
| \`操作日志.md\` | 插件自动记录的所有操作，最新的在最上面 | 只读，可用来了解之前做过什么 |
| \`.ata${sep}\` | 插件内部数据 | **不要读写** |

## 2. 修改表格的方法

1. 先看 \`数据目录.md\` 找到要改的表，确认列名。
2. 直接修改 \`数据${sep}\` 里对应的 CSV 文件。**不要去改 \`表格${sep}\` 里的 Excel，也不要整本打开 Excel 文件**——那样又慢又容易丢格式。
3. 大表不要整个读入：先读表头和最后几十行，需要查找时按关键字搜索。
4. 修改保存后，插件会在几秒内自动把这个工作簿重新生成为 \`表格${sep}\` 里的原生 Excel，覆盖上一版，并写进 \`操作日志.md\`。

## 3. CSV 的规则

- UTF-8 编码；第一行是表头；字段里有逗号、引号或换行时用英文双引号包起来，内部的双引号写成两个 \`""\`。
- **不要改表头**（列名、列的顺序），除非用户明确要求；需要新列时加在最后一列后面。
- **新增的行可以放在任意位置**：按日期、客户等顺序插在该在的地方，或者加在末尾都行。生成原表时会像 Excel 一样整行插入，下面的行、公式、合并单元格、图片、批注都跟着下移；紧挨着数据首尾插入时，合计公式的范围会自动包含新行。
- 用户要求删除时才删行；删除也会整行删掉，下面的行往上移。
- **不要调整已有行的先后顺序**（不要排序、不要整表重写），否则原表里的颜色和格式会对不上。
- 只改需要改的格子，其它内容保持原样（包括空格子和原来的写法）。
- 日期写成 \`YYYY-MM-DD\`（或照搬这一列已有的写法）；数字不要加千分位逗号和货币符号。
- 需要计算的（金额 = 数量 × 单价、合计等）请自己算好再写进去：CSV 里没有公式，导出时原表里已有的公式会保留，新增行会沿用上一行的公式。
- Markdown 笔记里的表格：只改表格部分，表格外的文字不要动。
- **不要删除、改名或移动 \`数据${sep}\` 里的文件。** 需要新表时，在 \`数据${sep}\` 里新建一个 CSV（第一行写表头），插件会自动为它生成 Excel。

## 4. 导入新的原始 Excel

- 新表格：复制到 \`自动处理${sep}导入${sep}\`。
- 拿到某个已导入表格的新版本、想用它替换旧的：复制到 \`自动处理${sep}导入${sep}替换${sep}\`（同名的旧表会进回收站）。
- 等 \`自动处理${sep}导入${sep}已完成${sep}\` 里出现 \`原文件名.结果.txt\`，里面写着生成了哪些 CSV。
- 支持 .xlsx / .xlsm / .csv / .tsv。老版 .xls 需要先另存为 .xlsx。

## 5. 每次任务结束要写记录

在 \`自动处理${sep}记录${sep}\` 新建一个 .txt 文件（文件名随意），第一行写一句话总结，后面写细节，例如：

\`\`\`
按 11 月 20 日送货单给「！！客户」加了 3 行
- 新增：2026-11-20 华盛布行 棉布 280克 200米 单价 8.5 金额 1700
- 依据：附件区/2026-11/送货单0001.jpg
- 需要用户核对：第 2 行单价看不清，按 8.5 填写
\`\`\`

插件会把它收进 \`操作日志.md\` 和首页的日志里。表格的具体改动（哪一格从什么改成什么）插件会自动对比记录，不用你写。

## 6. 需要马上拿到 Excel 时

原表一般几秒内就会自动更新。如果要另存一份到指定位置，在 \`自动处理${sep}导出${sep}\` 新建 .txt，写：

\`\`\`
表格: 工作簿名或表名
输出: 要另存的完整路径.xlsx
\`\`\`

结果写在 \`自动处理${sep}导出${sep}已完成${sep}请求文件名.结果.txt\`。

## 7. 其它

- 用户只让你「查」「看」「算一下告诉我」时，不要修改文件。
- 改错了不用慌：用户可以在 Obsidian 里一键撤销任何一次修改。
- 回复用户时说清楚：改了哪个表、新增 / 修改了哪几行、原表 Excel 在哪（\`表格${sep}…\`）、有哪些需要用户核对。
`;
    await this.writeDoc('操作守则.md', text);
  }

  scheduleDocs() {
    clearTimeout(this.docTimer);
    this.docTimer = setTimeout(() => { this.writeDataIndex().catch(() => {}); this.writeLogFile().catch(() => {}); }, 5000);
  }

  async writeDataIndex() {
    const lib = await this.library();
    const meta = this.settings.meta || {};
    const sp = this.splitLibrary(lib);
    const rel = (p) => p.slice(normalizePath(this.settings.folder).length + 1);
    const lines = ['<!-- 表格工作台自动生成，表格有变化时自动更新，手动修改会被覆盖 -->', '# 数据目录', '',
      `共 ${sp.books.length + sp.singles.length} 个表格文件（${sp.books.length} 个多工作表 Excel，${lib.length} 张表）。路径都相对于本文件所在的文件夹。修改请看「操作守则.md」。`, ''];
    const tableLines = (t) => {
      const m = meta[t.file.path] || {};
      const d = new Date(t.mtime), z = (n) => String(n).padStart(2, '0');
      lines.push(`- 底层数据：\`${rel(t.file.path)}\``);
      lines.push(`- 列（${t.headers.length}）：${t.headers.join('、')}`);
      lines.push(`- 行数：${t.rows}　最后修改：${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())} ${z(d.getHours())}:${z(d.getMinutes())}`);
      if (m.title) lines.push(`- 表头上方的标题：${String(m.title).replace(/\n/g, ' ').slice(0, 80)}`);
      if (t.sample && t.sample.length) lines.push(`- 样例：${t.headers.map((h, i) => `${h}=${cellStr(t.sample[0][i])}`).filter((x) => !/=$/.test(x)).slice(0, 12).join('；')}`);
      lines.push('');
    };
    for (const b of sp.books) {
      lines.push(`## ${b.name}（Excel 工作簿，${b.sheets.length} 张表）`, '');
      lines.push(`原表：\`${rel(this.outPathOf(b.sheets[0].file))}\``, '');
      for (const t of b.sheets) { lines.push(`### ${(meta[t.file.path] || {}).sheet || t.file.basename}`); tableLines(t); }
    }
    if (sp.singles.length) {
      lines.push('## 单独的表格', '');
      for (const t of sp.singles) { lines.push(`### ${t.file.basename}`); lines.push(`- 原表：\`${rel(this.outPathOf(t.file))}\``); tableLines(t); }
    }
    await this.writeDoc('数据目录.md', lines.join('\n'));
  }

  async writeLogFile() {
    const list = (await this.loadOpLog()).slice(-400).reverse();
    const NAME = { ext: '修改', 'ext-new': '新建表格', 'ext-del': '删除表格', output: '原表更新', note: '记录', import: '导入', export: '导出', create: '新建', delete: '删除', archive: '附件', undo: '撤销', edit: '手动修改', row: '行', col: '列', ai: '旧版' };
    const z = (n) => String(n).padStart(2, '0');
    const lines = ['<!-- 表格工作台自动生成，最新的在最上面，手动修改会被覆盖 -->', '# 操作日志', ''];
    let day = '';
    for (const e of list) {
      const d = new Date(e.t);
      const ds = `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
      if (ds !== day) { day = ds; lines.push(`## ${ds}`, ''); }
      lines.push(`- ${z(d.getHours())}:${z(d.getMinutes())} 【${NAME[e.type] || e.type}】${e.name ? e.name + '：' : ''}${e.text || ''}${e.review === 'undone' ? '（已撤销）' : ''}${e.output && e.outFile ? `（原表已更新：${e.outFile.slice(normalizePath(this.settings.folder).length + 1)}）` : ''}`);
      if (e.type === 'ext') {
        (e.cells || []).slice(0, 8).forEach((c) => lines.push(`  - 第 ${c.r} 行「${c.h}」：${c.o || '（空）'} → ${c.n || '（空）'}`));
        if ((e.cellsN || 0) > 8) lines.push(`  - ……共 ${e.cellsN} 个格子`);
        (e.added || []).slice(0, 5).forEach((r) => lines.push(`  - 新增：${(e.hdr || []).map((h, i) => (r[i] ? `${h}=${r[i]}` : '')).filter(Boolean).slice(0, 10).join('；')}`));
        if ((e.addedN || 0) > 5) lines.push(`  - ……共新增 ${e.addedN} 行`);
      }
      if (e.type === 'note' && e.items) e.items.slice(0, 20).forEach((x) => lines.push(`  ${x.startsWith('-') ? x : '- ' + x}`));
    }
    await this.writeDoc('操作日志.md', lines.join('\n'));
  }

  /** 「自动处理/记录」里的说明文件，收进日志 */
  async autoNote(tf) {
    const text = (await this.app.vault.read(tf)).replace(/\r/g, '');
    const ls = text.split('\n').map((x) => x.trimEnd()).filter((x) => x.trim());
    if (ls.length) this.logOp({ type: 'note', text: ls[0].replace(/^#+\s*/, ''), items: ls.slice(1, 60) });
    await this.moveDone(tf, this.autoDir() + '/记录/已记录');
  }

  /* ================= 自动处理：放文件即导入、放请求即导出、链接调用 ================= */
  autoDir() { return normalizePath(this.settings.folder + '/' + AUTO_DIR); }

  async ensureAutoDirs() {
    const d = this.autoDir();
    for (const sub of ['', '/导入', '/导入/替换', '/导入/已完成', '/导出', '/导出/已完成', '/记录', '/记录/已记录']) await this.ensureFolder(d + sub);
    const help = `${d}/使用说明.md`;
    const f = this.app.vault.getAbstractFileByPath(help);
    const text = this.autoHelpText();
    if (!(f instanceof TFile)) { try { await this.app.vault.create(help, text); } catch (e) { /* 已存在 */ } }
    else if ((await this.app.vault.read(f)).startsWith('<!-- 表格工作台自动生成') ) { const old = await this.app.vault.read(f); if (old !== text) await this.app.vault.modify(f, text); }
  }

  autoHelpText() {
    const abs = this.absPath(this.autoDir());
    const vault = this.app.vault.getName ? this.app.vault.getName() : '';
    const enc = encodeURIComponent;
    return `<!-- 表格工作台自动生成，会随插件更新；要写自己的笔记请另建文件 -->
# 自动处理

Obsidian 打开着的时候，往这里放文件就会自动处理，其它程序（脚本、AI 工具等）也能用。

文件夹位置：\`${abs}\`

## 自动导入 Excel

- 把 .xlsx / .xlsm / .csv / .tsv 放进 \`导入/\`：自动转换成表格工作台里的表（和点「导入文件」一样），同名的会另存一份，不覆盖。
- 放进 \`导入/替换/\`：如果已经导入过同名的 Excel，先把旧的表移到回收站，再导入新的（适合「每次都拿最新版原表来改」）。
- 处理完，原文件移到 \`导入/已完成/\`，旁边生成 \`原文件名.结果.txt\`，写着生成了哪些表（每张表是一个 CSV 文件，可以直接改）。

## 自动导出 Excel

在 \`导出/\` 里新建一个 .txt 文件（文件名随意），内容写要导出的表，每项一行：

\`\`\`
表格: 史2026-客户订单待发明细！！（26-2.27(2)
输出: D:\\桌面\\订单-最新.xlsx
\`\`\`

- 「表格」可以写 Excel 工作簿名、表名或 CSV 文件路径。从 Excel 导入的表会按原格式（颜色、公式、合并单元格、图片）整本导出。
- 「输出」可以不写，不写就只保存在表格文件夹的「导出」里；写了会再复制一份到这个位置。
- 处理完，请求文件移到 \`导出/已完成/\`，旁边生成 \`请求文件名.结果.txt\`，写着导出文件的位置或出错原因。

## 用链接调用

也可以打开下面这样的链接（Windows 命令行用 \`start "" "链接"\`，Mac 用 \`open "链接"\`）：

- 导入：\`obsidian://table-workbench?vault=${enc(vault)}&action=import&path=${enc('D:\\\\订单.xlsx')}\`（加 \`&mode=replace\` 表示替换同名的）
- 导出：\`obsidian://table-workbench?vault=${enc(vault)}&action=export&table=${enc('工作簿名')}&out=${enc('D:\\\\订单-最新.xlsx')}\`
- 打开某张表：\`obsidian://table-workbench?vault=${enc(vault)}&action=open&table=${enc('表名')}\`

链接调用的结果写在 \`${AUTO_DIR}/最近一次结果.txt\`。
`;
  }

  /** 按名字找表：CSV 路径、工作簿名、工作表名、文件名都可以 */
  resolveTable(name) {
    name = String(name || '').trim().replace(/^["'「]|["'」]$/g, '');
    if (!name) return null;
    const meta = this.settings.meta || {};
    const files = this.tableFiles();
    const norm = (x) => normalizePath(String(x)).toLowerCase();
    const base = this.basePath();
    let rel = name;
    if (base && name.replace(/\\/g, '/').toLowerCase().startsWith(base.replace(/\\/g, '/').toLowerCase() + '/')) rel = name.replace(/\\/g, '/').slice(base.length + 1);
    let f = files.find((x) => norm(x.path) === norm(rel)) || files.find((x) => norm(x.path) === norm(rel + '.csv'));
    if (f) return f;
    const book = name.replace(/\.(xlsx|xlsm|csv)$/i, '');
    const inBook = files.filter((x) => meta[x.path] && meta[x.path].workbook === book).sort((a, b) => (meta[a.path].sheetIndex || 0) - (meta[b.path].sheetIndex || 0));
    if (inBook.length) return inBook[0];
    return files.find((x) => x.basename === book) || files.find((x) => meta[x.path] && meta[x.path].sheet === book) || null;
  }

  /** 导出一张表或它所在的整本工作簿；out 是另外再存一份的位置（可以是库外的绝对路径） */
  async runExport(target, out, from) {
    const meta = this.settings.meta || {};
    const tm = meta[target.path] || {};
    const sibs = this.workbookSiblings(target);
    const linked = tm.link && (await this.app.vault.adapter.exists(`${this.ataDir()}/sources/${tm.link.id}.xlsx`));
    let res;
    if (linked) res = await this.exportLinked(target, sibs.length > 1 ? sibs : [target], { quiet: true, from });
    else if (sibs.length > 1) res = await this.exportWorkbook(target, sibs, { quiet: true, from });
    else {
      const t = target.extension === 'md' ? ((findMdTable(await this.app.vault.read(target)) || {}).table || { headers: [], rows: [] }) : gridToTable(parseCSV(await this.app.vault.read(target)));
      res = await this.exportTable(target, t, 'xlsx', { quiet: true, from });
    }
    if (!res || !res.path) throw new Error('导出没有生成文件');
    let copied = '', copyError = '';
    if (out) try {
      const fs = nodeReq('fs'), path = nodeReq('path');
      if (!fs || !path) throw new Error('只有电脑版 Obsidian 才能复制到库外的位置');
      let dest = String(out).trim().replace(/^["']|["']$/g, '');
      if (!path.isAbsolute(dest)) dest = path.join(this.basePath(), dest);
      if (!/\.xlsx$/i.test(dest)) {
        let isDir = false;
        try { isDir = (await fs.promises.stat(dest)).isDirectory(); } catch (e) { /* 不存在 */ }
        dest = isDir ? path.join(dest, res.path.split('/').pop()) : dest + '.xlsx';
      }
      await fs.promises.mkdir(path.dirname(dest), { recursive: true });
      await fs.promises.writeFile(dest, Buffer.from(await this.app.vault.adapter.readBinary(res.path)));
      copied = dest;
    } catch (e) { copyError = e.message || String(e); }
    return { path: res.path, abs: this.absPath(res.path), copied, copyError, changed: res.changed, sheetsChanged: res.sheetsChanged, linked: !!linked };
  }

  /** 替换模式：先把同名工作簿已经导入的表移到回收站 */
  async removeImported(fileName) {
    const book = baseName(fileName);
    const meta = this.settings.meta || {};
    const root = this.dataDir();
    const old = this.tableFiles().filter((f) => (meta[f.path] && meta[f.path].workbook === book) || f.path === `${root}/${book}.csv`);
    if (old.length) {
      await this.deleteFiles(old);
      await new Promise((r) => setTimeout(r, 300));
    }
    return old.length;
  }

  async runImport(list, replace, from) {
    let removed = 0;
    if (replace) for (const x of list) removed += await this.removeImported(x.file.name);
    const res = (await this.importBatch(list, { quiet: true, wait: true, from })) || { ok: 0, failed: [], paths: [] };
    res.removed = removed;
    return res;
  }

  async writeResult(path, lines) {
    try {
      const f = this.app.vault.getAbstractFileByPath(path);
      const text = lines.join('\n') + '\n';
      if (f instanceof TFile) await this.app.vault.modify(f, text); else await this.app.vault.create(path, text);
    } catch (e) { console.warn('[表格工作台] 写结果失败', e); }
  }

  stamp() {
    const d = new Date(), z = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())} ${z(d.getHours())}:${z(d.getMinutes())}:${z(d.getSeconds())}`;
  }

  async moveDone(tf, dir) {
    await this.ensureFolder(dir);
    const dot = tf.name.lastIndexOf('.');
    const dest = await this.uniquePath(dir, dot > 0 ? tf.name.slice(0, dot) : tf.name, dot > 0 ? tf.name.slice(dot + 1) : 'bin');
    await this.app.fileManager.renameFile(tf, dest);
    return dest;
  }

  /** 「自动处理/导入」里出现了表格文件 */
  async autoImport(tf) {
    const d = this.autoDir();
    const replace = tf.path.startsWith(d + '/导入/替换/');
    const file = new File([await this.app.vault.readBinary(tf)], tf.name, { type: MIME[(tf.extension || '').toLowerCase()] || '' });
    const lines = [`时间：${this.stamp()}`, `原文件：${tf.name}`, `方式：${replace ? '替换同名' : '新导入'}`];
    try {
      const res = await this.runImport([{ file, rel: tf.name }], replace, '自动处理');
      if (res.removed) lines.push(`已把旧版本的 ${res.removed} 张表移到回收站`);
      lines.push(`结果：${res.ok ? `成功，生成 ${res.ok} 张表` : '没有导入任何表格'}`);
      if (res.failed && res.failed.length) lines.push('失败：' + res.failed.join('；'));
      if (res.paths && res.paths.length) { lines.push('', '生成的表格（每张表一个 CSV 文件）：'); res.paths.forEach((p) => lines.push('- ' + this.absPath(p))); }
    } catch (e) {
      lines.push('结果：失败，' + (e.message || e));
    }
    const moved = await this.moveDone(tf, d + '/导入/已完成');
    await this.writeResult(moved.replace(/\.[^.\/]+$/, '') + '.结果.txt', lines);
  }

  /** 「自动处理/导出」里出现了请求文件 */
  async autoExport(tf) {
    const d = this.autoDir();
    const text = await this.app.vault.read(tf);
    let req = {};
    try { req = JSON.parse(text); } catch (e) {
      text.split(/\r?\n/).forEach((line) => {
        const m = /^\s*(表格|表|工作簿|table|输出|out|output)\s*[:：=]\s*(.+?)\s*$/i.exec(line);
        if (m) req[/^(输出|out|output)$/i.test(m[1]) ? 'out' : 'table'] = m[2];
      });
      if (!req.table) { const first = text.split(/\r?\n/).map((x) => x.trim()).find(Boolean); if (first) req.table = first; }
    }
    const name = req.table || req['表格'] || req['工作簿'];
    const out = req.out || req['输出'];
    const lines = [`时间：${this.stamp()}`, `要导出：${name || '（没写）'}`];
    try {
      const target = this.resolveTable(name);
      if (!target) throw new Error(`找不到「${name}」这张表或工作簿`);
      const res = await this.runExport(target, out, '自动处理');
      lines.push('结果：成功', `导出文件：${res.abs}`);
      if (res.copied) lines.push(`另存到：${res.copied}`);
      if (res.copyError) lines.push(`另存失败：${res.copyError}（导出文件本身已经生成）`);
      lines.push(`方式：${res.linked ? `按原 Excel 格式导出${res.sheetsChanged ? `（${res.sheetsChanged} 个工作表有改动，共 ${res.changed} 个单元格）` : '（和原文件一致）'}` : '普通 Excel'}`);
    } catch (e) {
      lines.push('结果：失败，' + (e.message || e));
    }
    const moved = await this.moveDone(tf, d + '/导出/已完成');
    await this.writeResult(moved.replace(/\.[^.\/]+$/, '') + '.结果.txt', lines);
  }

  /** 排队处理，一次只做一件，避免同时导入导出 */
  queueAuto(tf, kind) {
    if (!this.autoSeen) this.autoSeen = new Set();
    if (this.autoSeen.has(tf.path)) return;
    this.autoSeen.add(tf.path);
    this.autoQ = (this.autoQ || Promise.resolve()).then(async () => {
      /* 等文件写完（大小不再变化） */
      let last = -1;
      for (let i = 0; i < 40; i++) {
        const cur = this.app.vault.getAbstractFileByPath(tf.path);
        if (!(cur instanceof TFile)) return;
        if (cur.stat.size === last && last > 0) break;
        last = cur.stat.size;
        await new Promise((r) => setTimeout(r, 700));
      }
      try { if (kind === 'import') await this.autoImport(tf); else if (kind === 'note') await this.autoNote(tf); else await this.autoExport(tf); }
      catch (e) { console.warn('[表格工作台] 自动处理失败', e); new Notice('自动处理失败：' + (e.message || e)); }
    }).finally(() => this.autoSeen.delete(tf.path));
  }

  autoKind(f) {
    if (!(f instanceof TFile) || this.settings.autoProcess === false) return null;
    const d = this.autoDir();
    if (/\.结果\.txt$/.test(f.name) || f.name === '使用说明.md' || f.name === '最近一次结果.txt') return null;
    if (f.path.startsWith(d + '/导入/') && !f.path.startsWith(d + '/导入/已完成/') && TABLE_EXT.test(f.name) && !f.name.startsWith('~$')) return 'import';
    if (f.path.startsWith(d + '/导出/') && !f.path.startsWith(d + '/导出/已完成/') && /\.(txt|json|md)$/i.test(f.name)) return 'export';
    if (f.path.startsWith(d + '/记录/') && !f.path.startsWith(d + '/记录/已记录/') && /\.(txt|md)$/i.test(f.name)) return 'note';
    return null;
  }

  scanAuto() {
    this.app.vault.getFiles().forEach((f) => { const k = this.autoKind(f); if (k) this.queueAuto(f, k); });
  }

  /** obsidian://table-workbench?action=…  */
  async handleProtocol(p) {
    const lines = [`时间：${this.stamp()}`, `操作：${p.action || ''}`];
    try {
      if (p.action === 'import') {
        if (!p.path) throw new Error('缺少 path 参数（要导入的文件的完整路径）');
        const list = await filesFromPaths(String(p.path).split('|').filter(Boolean));
        if (!list.length) throw new Error('读不到这个文件');
        const res = await this.runImport(list, p.mode === 'replace', '链接调用');
        lines.push(`结果：${res.ok ? `成功，生成 ${res.ok} 张表` : '没有导入任何表格'}`);
        if (res.failed && res.failed.length) lines.push('失败：' + res.failed.join('；'));
        (res.paths || []).forEach((x) => lines.push('- ' + this.absPath(x)));
      } else if (p.action === 'export') {
        const target = this.resolveTable(p.table);
        if (!target) throw new Error(`找不到「${p.table || ''}」这张表或工作簿`);
        const res = await this.runExport(target, p.out, '链接调用');
        lines.push('结果：成功', `导出文件：${res.abs}`);
        if (res.copied) lines.push(`另存到：${res.copied}`);
        if (res.copyError) lines.push(`另存失败：${res.copyError}（导出文件本身已经生成）`);
      } else if (p.action === 'open') {
        const target = this.resolveTable(p.table);
        if (!target) throw new Error(`找不到「${p.table || ''}」这张表或工作簿`);
        await this.openTable(target);
        lines.push('结果：已打开 ' + target.path);
      } else throw new Error('不认识的操作，可以用 import / export / open');
    } catch (e) {
      lines.push('结果：失败，' + (e.message || e));
      new Notice('表格工作台：' + (e.message || e), 8000);
    }
    await this.ensureFolder(this.autoDir());
    await this.writeResult(`${this.autoDir()}/最近一次结果.txt`, lines);
  }

  /* ================= 文件管理 ================= */
  fmGroups(lib) {
    const sp = this.splitLibrary(lib);
    const meta = this.settings.meta || {};
    const groups = [];
    sp.books.forEach((b) => groups.push({ key: b.key, name: b.name, book: true, files: b.sheets.map((t) => t.file), stats: b.sheets, rows: b.rows, mtime: b.mtime, folder: b.folder }));
    sp.singles.forEach((t) => groups.push({ key: t.file.path, name: (meta[t.file.path] || {}).sheet || t.file.basename, book: false, files: [t.file], stats: [t], rows: t.rows, mtime: t.mtime }));
    groups.forEach((g) => {
      g.out = this.outPathOf(g.files[0]);
      g.linked = !!((meta[g.files[0].path] || {}).link);
      g.dataPath = g.book ? g.folder : g.files[0].path;
    });
    return groups;
  }

  async copyPaths(paths, what) {
    const text = paths.map((p) => this.absPath(p)).join('\n');
    await this.copyText(text, `${paths.length > 1 ? ` ${paths.length} 个` : ''}${what || '地址'}`);
  }

  /** 把文件本身复制到剪贴板（可以直接粘贴到微信、资源管理器）。Windows 支持多个文件，其它系统复制路径 */
  async copyFiles(paths) {
    const abs = paths.map((p) => this.absPath(p));
    const cp = nodeReq('child_process');
    const isWin = typeof process !== 'undefined' && process.platform === 'win32';
    const isMac = typeof process !== 'undefined' && process.platform === 'darwin';
    try {
      if (cp && isWin) {
        const list = abs.map((p) => `'${p.replace(/'/g, "''")}'`).join(',');
        await new Promise((res, rej) => cp.execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', `Set-Clipboard -Path ${list}`], { windowsHide: true }, (err) => (err ? rej(err) : res())));
        new Notice(`已复制 ${abs.length} 个文件，可以直接粘贴到微信或文件夹里。`);
        return;
      }
      if (cp && isMac && abs.length === 1) {
        await new Promise((res, rej) => cp.execFile('osascript', ['-e', `set the clipboard to (POSIX file "${abs[0].replace(/"/g, '\\"')}")`], (err) => (err ? rej(err) : res())));
        new Notice('已复制文件，可以直接粘贴。');
        return;
      }
    } catch (e) { console.warn('[表格工作台] 复制文件失败', e); }
    if (abs.length === 1 && /\.(png|jpe?g|webp|gif|bmp)$/i.test(abs[0])) {
      const tf = this.app.vault.getAbstractFileByPath(paths[0]);
      if (tf instanceof TFile) { await copyImageNotice(this.app, tf); return; }
    }
    await this.copyText(abs.join('\n'), `${abs.length} 个文件的路径（这台电脑不支持直接复制多个文件）`);
  }

  showInSystem(path) {
    const abs = this.absPath(path);
    try {
      const { shell } = require('electron');
      if (shell && shell.showItemInFolder) { shell.showItemInFolder(abs); return; }
    } catch (e) { /* 不是桌面版 */ }
    if (this.app.showInFolder) this.app.showInFolder(path);
  }

  async deleteGroups(groups) {
    const files = groups.flatMap((g) => g.files);
    const outs = groups.map((g) => this.app.vault.getAbstractFileByPath(g.out)).filter((f) => f instanceof TFile);
    await this.deleteFiles(files);
    for (const f of outs) {
      try { if (this.app.fileManager && this.app.fileManager.trashFile) await this.app.fileManager.trashFile(f); else await this.app.vault.trash(f, true); } catch (e) { /* 已经没了 */ }
    }
    if (this.fmSel) groups.forEach((g) => this.fmSel.delete(g.key));
    this.scheduleDocs();
    await this.render();
  }

  renderLibrary(body, lib) {
    const groups = this.fmGroups(lib);
    if (!this.fmSel) this.fmSel = new Set();
    if (!this.fmOpen) this.fmOpen = new Set();
    const head = body.createDiv('ata-header');
    const left = head.createDiv();
    const title = left.createDiv('ata-title');
    this.ico(title, 'folder-tree');
    title.createSpan({ text: '文件管理' });
    left.createDiv({ cls: 'ata-sub', text: `${groups.length} 个表格文件（${groups.filter((g) => g.book).length} 个多工作表 Excel，共 ${lib.length} 张表）· 原表 Excel 在「${OUT_DIR}」文件夹，底层数据在「${DATA_DIR}」文件夹，改了数据会自动更新原表` });
    const right = head.createDiv('ata-header-right');
    this.btn(right, '原表文件夹', 'folder-open', async () => { await this.ensureFolder(this.outDir()); this.showInSystem(this.outDir()); });
    this.btn(right, '全部更新原表', 'refresh-cw', () => { const seen = new Set(); this.tableFiles().forEach((f, i) => { const d = this.outPathOf(f); if (!seen.has(d)) { seen.add(d); this.scheduleOutput(f, 300 + i * 50); } }); new Notice('已开始重新生成全部原表。'); });

    if (!groups.length) { this.renderDropzone(body); return; }
    const shown = groups.slice().sort((a, b) => b.mtime - a.mtime);
    const selected = () => groups.filter((g) => this.fmSel.has(g.key));

    const bar = body.createDiv('ata-fmbar');
    const all = bar.createEl('input', { type: 'checkbox', cls: 'ata-fm-check' });
    all.checked = shown.length > 0 && shown.every((g) => this.fmSel.has(g.key));
    all.addEventListener('change', () => { shown.forEach((g) => (all.checked ? this.fmSel.add(g.key) : this.fmSel.delete(g.key))); this.render(); });
    const n = selected().length;
    bar.createSpan({ cls: 'ata-muted', text: n ? `已选 ${n} 个` : `${shown.length} 个 · 勾选后可以批量复制地址、删除` });
    const acts = bar.createDiv('ata-fmbar-acts');
    if (n) {
      this.btn(acts, '复制原表地址', 'copy', () => this.copyPaths(selected().map((g) => g.out), '原表地址'), 'ata-btn-sm');
      this.btn(acts, '复制数据地址', 'database', () => this.copyPaths(selected().map((g) => g.dataPath), '数据地址'), 'ata-btn-sm');
      this.btn(acts, '复制原表文件', 'files', () => this.copyFiles(selected().map((g) => g.out).filter((p) => this.app.vault.getAbstractFileByPath(p))), 'ata-btn-sm');
      this.btn(acts, '立即更新原表', 'refresh-cw', () => { selected().forEach((g, i) => this.scheduleOutput(g.files[0], 200 + i * 50)); }, 'ata-btn-sm');
      this.btn(acts, '删除', 'trash-2', () => {
        const s = selected();
        new ConfirmModal(this.app, `删除 ${s.length} 个表格文件？`, s.slice(0, 10).map((g) => '• ' + g.name).join('\n') + '\n底层数据和原表 Excel 都会移到回收站，需要时可以恢复。', [{ text: `删除 ${s.length} 个`, warning: true, onClick: () => this.deleteGroups(s) }]).open();
      }, 'ata-btn-sm is-danger');
      this.btn(acts, '取消选择', 'x', () => { this.fmSel.clear(); this.render(); }, 'ata-btn-sm');
    }

    const list = body.createDiv('ata-fmlist');
    const hdr = list.createDiv('ata-fmrow is-head');
    ['', '名称', '类型', '行数', '数据更新', '原表 Excel', ''].forEach((x) => hdr.createDiv({ text: x }));
    const now = Date.now();
    for (const g of shown) {
      const row = list.createDiv('ata-fmrow' + (this.fmSel.has(g.key) ? ' is-selected' : '') + (this.revealKey === g.key ? ' is-flash' : ''));
      row.dataset.key = g.key;
      row.dataset.q = [g.name, g.dataPath, ...g.stats.flatMap((t) => [t.file.basename, ...(t.headers || [])])].join(' ').toLowerCase();
      const cb = row.createDiv().createEl('input', { type: 'checkbox', cls: 'ata-fm-check' });
      cb.checked = this.fmSel.has(g.key);
      cb.addEventListener('click', (e) => e.stopPropagation());
      cb.addEventListener('change', () => { if (cb.checked) this.fmSel.add(g.key); else this.fmSel.delete(g.key); this.render(); });
      const nm = row.createDiv('ata-fm-name');
      const tw = nm.createSpan('ata-fm-twist');
      setIcon(tw, g.book ? (this.fmOpen.has(g.key) ? 'chevron-down' : 'chevron-right') : 'dot');
      this.ico(nm, g.book ? 'book-open' : 'sheet');
      const nt = nm.createDiv('ata-fm-nt');
      nt.createDiv({ cls: 'ata-fm-title', text: g.name, attr: { title: g.name } });
      nt.createDiv({ cls: 'ata-fm-path', text: g.dataPath, attr: { title: this.absPath(g.dataPath) } });
      row.createDiv({ cls: 'ata-fm-type', text: g.book ? `Excel · ${g.files.length} 张表` : g.linked ? 'Excel' : 'CSV 表' });
      row.createDiv({ cls: 'ata-fm-num', text: g.rows.toLocaleString() });
      row.createDiv({ cls: 'ata-fm-time', text: fmtTime(g.mtime) });
      const oc = row.createDiv('ata-fm-out');
      const of = this.app.vault.getAbstractFileByPath(g.out);
      const st = this.outState && this.outState.get(g.out);
      if (st && st.status === 'pending') oc.createSpan({ cls: 'ata-fm-st is-pending', text: '更新中…' });
      else if (st && st.status === 'busy') oc.createSpan({ cls: 'ata-fm-st is-warn', text: '被占用，稍后重试', attr: { title: '原表 Excel 正被 Excel / WPS 打开，关掉后会自动更新' } });
      else if (st && st.status === 'error') oc.createSpan({ cls: 'ata-fm-st is-err', text: '更新失败', attr: { title: st.msg || '' } });
      else if (of instanceof TFile) oc.createSpan({ cls: 'ata-fm-st is-ok', text: '✓ ' + fmtTime(of.stat.mtime) + (of.stat.mtime + 1000 < g.mtime ? '（比数据旧）' : ''), attr: { title: this.absPath(g.out) } });
      else oc.createSpan({ cls: 'ata-fm-st is-pending', text: '还没生成' });
      const ac = row.createDiv('ata-fm-acts');
      this.iconBtn(ac, 'copy', '复制原表地址', () => this.copyPaths([g.out], '原表地址'));
      this.iconBtn(ac, 'database', '复制数据地址', () => this.copyPaths([g.dataPath], '数据地址'));
      if (of instanceof TFile) {
        this.iconBtn(ac, 'external-link', '用 Excel 打开原表', () => this.app.openWithDefaultApp && this.app.openWithDefaultApp(g.out));
        this.iconBtn(ac, 'folder-open', '在文件夹中显示', () => this.showInSystem(g.out));
      }
      this.iconBtn(ac, 'trash-2', '删除', () => new ConfirmModal(this.app, `删除「${g.name}」？`, '底层数据和原表 Excel 都会移到回收站，需要时可以恢复。', [{ text: '删除', warning: true, onClick: () => this.deleteGroups([g]) }]).open()).addClass('is-danger');
      row.addEventListener('click', (e) => {
        if (e.target.closest('button, input')) return;
        if (g.book) { if (this.fmOpen.has(g.key)) this.fmOpen.delete(g.key); else this.fmOpen.add(g.key); this.render(); }
        else { if (this.fmSel.has(g.key)) this.fmSel.delete(g.key); else this.fmSel.add(g.key); this.render(); }
      });
      if (g.book && this.fmOpen.has(g.key)) {
        const sub = list.createDiv('ata-fmsub');
        g.stats.forEach((t) => {
          const m = (this.settings.meta || {})[t.file.path] || {};
          const r = sub.createDiv('ata-fmsubrow');
          this.ico(r, 'sheet');
          r.createSpan({ cls: 'ata-fm-sheet', text: m.sheet || t.file.basename });
          r.createSpan({ cls: 'ata-muted', text: `${t.rows.toLocaleString()} 行 · ${t.headers.length} 列 · ${t.headers.slice(0, 6).join('、')}${t.headers.length > 6 ? '…' : ''}` });
          const a2 = r.createDiv('ata-fm-acts');
          this.iconBtn(a2, 'database', '复制这张表的数据地址', () => this.copyPaths([t.file.path], '数据地址'));
          this.iconBtn(a2, 'file-text', '在 Obsidian 里查看 CSV', () => this.app.workspace.getLeaf('tab').openFile(t.file));
        });
      }
    }
    if (this.query) { const ql = this.query.toLowerCase(); list.querySelectorAll('.ata-fmrow[data-q]').forEach((r) => { r.style.display = r.dataset.q.includes(ql) ? '' : 'none'; }); }
    if (this.revealKey) {
      const el = list.querySelector(`.ata-fmrow[data-key="${CSS.escape(this.revealKey)}"]`);
      if (el) setTimeout(() => el.scrollIntoView({ block: 'center' }), 50);
      setTimeout(() => { this.revealKey = null; }, 1500);
    }
  }

  /** 以前「打开表格」的地方：现在跳到文件管理里定位这个文件 */
  async openTable(file) {
    if (!(file instanceof TFile)) return;
    const g = this.fmGroups(await this.library()).find((x) => x.files.some((f) => f.path === file.path));
    this.revealKey = g ? g.key : null;
    if (g && g.book) { if (!this.fmOpen) this.fmOpen = new Set(); this.fmOpen.add(g.key); }
    this.tab = 'library';
    await this.render();
  }

  /* ================= 首页：日志 ================= */
  renderHome(body) {
    const head = body.createDiv('ata-header');
    const left = head.createDiv();
    const title = left.createDiv('ata-title');
    this.ico(title, 'scroll-text');
    title.createSpan({ text: '操作日志' });
    left.createDiv({ cls: 'ata-sub', text: '每一次修改（改了哪些格子、加了哪些行）、原表 Excel 的自动更新、导入、删除、处理说明都记在这里，同时写进表格文件夹里的「操作日志.md」。' });
    const right = head.createDiv('ata-header-right');
    const openDoc = (name) => { const f = this.app.vault.getAbstractFileByPath(normalizePath(`${this.settings.folder}/${name}`)); if (f instanceof TFile) this.app.workspace.getLeaf('tab').openFile(f); else new Notice('还没生成，稍等几秒。'); };
    this.btn(right, '数据目录', 'list-tree', () => openDoc('数据目录.md'));
    this.btn(right, '新手教程', 'graduation-cap', () => this.openTutorial());
    this.btn(right, '操作守则', 'book-open-check', () => openDoc('操作守则.md'));
    this.btn(right, '复制 AI 开场白', 'clipboard-copy', () => this.copyText(this.aiOpener(), 'AI 开场白（粘贴给 AI，后面接上你的需求）'), 'mod-cta');
    this.renderOpLog(body);
  }

  /* ================= 附件区文件夹 ================= */
  archiveFolders() {
    const dir = this.archiveDir() + '/';
    const set = new Set((this.settings.archiveFolders || []).filter(Boolean));
    this.archiveFiles().forEach((f) => { const rel = f.path.slice(dir.length); if (rel.includes('/')) set.add(rel.slice(0, rel.lastIndexOf('/'))); });
    const dt = (x) => /^\d{4}-\d{2}(\/|$)/.test(x);
    return [...set].sort((a, b) => (dt(a) === dt(b) ? (dt(a) ? b.localeCompare(a) : a.localeCompare(b, 'zh-Hans-CN')) : dt(a) ? 1 : -1));
  }

  folderOf(tf) {
    const rel = tf.path.slice(this.archiveDir().length + 1);
    return rel.includes('/') ? rel.slice(0, rel.lastIndexOf('/')) : '';
  }

  async newArchiveFolder(then) {
    new TextPromptModal(this.app, '新建附件文件夹（可以用 / 建子文件夹，例如：客户/华盛）', '', async (name) => {
      name = normalizePath(String(name || '').split('/').map(safeSeg).filter(Boolean).join('/'));
      if (!name) return;
      await this.ensureFolder(`${this.archiveDir()}/${name}`);
      const s = this.settings;
      s.archiveFolders = [...new Set([...(s.archiveFolders || []), name])];
      await this.plugin.saveSettings();
      if (then) await then(name); else { this.archFolder = name; await this.render(); }
    }).open();
  }

  async moveArchive(tfs, folder) {
    const dest = folder ? `${this.archiveDir()}/${folder}` : this.archiveDir();
    await this.ensureFolder(dest);
    const am = this.settings.archiveMeta || {};
    this.archiving = true;
    let n = 0;
    try {
      for (const tf of tfs) {
        if (this.folderOf(tf) === folder) continue;
        const dot = tf.name.lastIndexOf('.');
        const np = await this.uniquePath(dest, dot > 0 ? tf.name.slice(0, dot) : tf.name, dot > 0 ? tf.name.slice(dot + 1) : 'bin');
        const old = tf.path;
        await this.app.fileManager.renameFile(tf, np);
        if (am[old]) { am[np] = am[old]; delete am[old]; }
        n++;
      }
    } finally { this.archiving = false; }
    await this.plugin.saveSettings();
    this.logOp({ type: 'archive', text: `把 ${n} 个附件移到「${folder || '附件区'}」`, items: tfs.map((f) => f.name) });
    if (this.archSel) this.archSel.clear();
    new Notice(`已移动 ${n} 个文件到「${folder || '附件区'}」。`);
    await this.render();
  }

  pickArchiveFolder(onPick) {
    const v = this;
    class FolderPick extends FuzzySuggestModal {
      getItems() { return ['＋ 新建文件夹…', ...v.archiveFolders()]; }
      getItemText(x) { return x; }
      onChooseItem(x) { if (x.startsWith('＋')) v.newArchiveFolder(onPick); else onPick(x); }
    }
    const m = new FolderPick(this.app);
    m.setPlaceholder('移动到哪个文件夹？');
    m.open();
  }

  /* ================= 原表样式视图 ================= */
  isExcelMode() {
    if (!this.file || !this.link || this.file.extension !== 'csv') return false;
    if (!this.viewModes) this.viewModes = new Map();
    return (this.viewModes.get(this.file.path) || 'excel') === 'excel';
  }

  setViewMode(mode) {
    if (!this.file) return;
    if (!this.viewModes) this.viewModes = new Map();
    this.viewModes.set(this.file.path, mode);
    this.renderWorkspaceFoot();
    this.renderTable();
  }

  /** 读取（第一次时生成）原 Excel 的版面信息，缓存在 .ata/maps/<id>/ 里 */
  async getLayout(link) {
    if (!this.xlCache) this.xlCache = new Map();
    const key = link.id + ':' + link.sheet;
    if (this.xlCache.has(key)) return this.xlCache.get(key);
    const a = this.app.vault.adapter, dir = `${this.ataDir()}/maps/${link.id}`;
    let xf = this.xlCache.get(link.id + ':xf'), L = null;
    try {
      if (!xf) xf = JSON.parse(await a.read(`${dir}/layout.xf.json`));
      L = JSON.parse(await a.read(`${dir}/${link.sheet}.layout.json`));
      if (L.ver !== XL_LAYOUT_VER) L = null;
    } catch (e) { L = null; }
    if (!L) {
      if (!this.xlBuilding) this.xlBuilding = new Map();
      if (!this.xlBuilding.has(link.id)) {
        this.xlBuilding.set(link.id, (async () => {
          const buf = await a.readBinary(`${this.ataDir()}/sources/${link.id}.xlsx`);
          const res = await buildBookLayouts(buf);
          await a.write(`${dir}/layout.xf.json`, JSON.stringify(res.xf));
          for (let i = 0; i < res.sheets.length; i++) {
            res.sheets[i].ver = XL_LAYOUT_VER;
            await a.write(`${dir}/${i}.layout.json`, JSON.stringify(res.sheets[i]));
            this.xlCache.set(link.id + ':' + i, { L: res.sheets[i], xf: res.xf });
          }
          this.xlCache.set(link.id + ':xf', res.xf);
        })().finally(() => this.xlBuilding.delete(link.id)));
      }
      await this.xlBuilding.get(link.id);
      return this.xlCache.get(key) || null;
    }
    this.xlCache.set(link.id + ':xf', xf);
    const out = { L, xf };
    this.xlCache.set(key, out);
    return out;
  }

  async renderExcel() {
    const wrap = this.previewEl, file = this.file, link = this.link;
    if (!wrap || !file || !link) return;
    const token = (this.xlToken = (this.xlToken || 0) + 1);
    wrap.onscroll = null;
    wrap.addClass('ata-xlwrap');
    let lay = this.xlCache && this.xlCache.get(link.id + ':' + link.sheet);
    if (!lay) {
      wrap.empty();
      wrap.createDiv({ cls: 'ata-empty-sm ata-pad', text: '正在读取原表格式…（每个工作簿只在第一次打开时需要几秒）' });
      try { lay = await this.getLayout(link); }
      catch (e) { lay = null; console.warn('[表格工作台] 读取原表格式失败', e); }
      if (token !== this.xlToken || this.file !== file) return;
      if (!lay) {
        wrap.empty();
        const box = wrap.createDiv({ cls: 'ata-empty-sm ata-pad', text: '读取原 Excel 的格式失败（原文件副本可能不在了），先用「编辑数据」视图查看。' });
        this.btn(box, '切到编辑数据', 'table-2', () => this.setViewMode('data'), 'ata-btn-sm');
        return;
      }
    }
    this.drawExcel(lay);
  }

  drawExcel(lay) {
    const wrap = this.previewEl, t = this.table, map = this.link.map, L = lay.L, XF = lay.xf;
    const z = this.xlZoom || 1;
    const GW = 46, HH = 22;
    const H = map.headerRow;
    const rows0 = new Set(map.rows0 || []);
    /* 表格行 → 原表行 */
    const srcToTr = new Map(), newTr = [];
    t.rows.forEach((r, i) => { if (r.__src != null && r.__src >= 0 && !srcToTr.has(r.__src)) srcToTr.set(r.__src, i); else newTr.push(i); });
    let lastData = H;
    rows0.forEach((r) => { if (r > lastData) lastData = r; });
    srcToTr.forEach((i, r) => { if (r > lastData) lastData = r; });
    const base = lastData + 1, nNew = newTr.length;
    const colSrc = t.colSrc && t.colSrc.length === t.headers.length ? t.colSrc : (map.cols || []).slice(0, t.headers.length);
    let maxCol = L.maxCol;
    const scToTc = new Map();
    let tmplC = 0;
    t.headers.forEach((h, k) => {
      const c = colSrc[k];
      if (c != null && c >= 0) { scToTc.set(c, k); tmplC = c; if (c >= maxCol) maxCol = c + 1; }
    });
    t.headers.forEach((h, k) => { const c = colSrc[k]; if (!(c != null && c >= 0)) scToTc.set(maxCol++, k); });
    const tcToSc = new Map();
    scToTc.forEach((k, c) => tcToSc.set(k, c));
    const totalRows = Math.max(L.maxRow, base) + nNew;
    const sheetRow = (d) => (d < base ? d : d < base + nNew ? -(d - base) - 1 : d - nNew);
    const dispRow = (r) => (r < base ? r : r + nNew);
    const hideR = new Set(L.rowHide), hideC = new Set(L.colHide);
    const rowH = (d) => { const r = sheetRow(d); const rr = r < 0 ? lastData : r; return hideR.has(rr) && r >= 0 ? 0 : (L.rowH[rr] || L.defH); };
    const colW = (c) => (c >= L.maxCol && !(c in L.colW) ? (L.colW[tmplC] || L.defW) : hideC.has(c) ? 0 : (L.colW[c] || L.defW));
    const rowY = new Float64Array(totalRows + 1);
    for (let d = 0; d < totalRows; d++) rowY[d + 1] = rowY[d] + rowH(d) * z;
    const colX = new Float64Array(maxCol + 1);
    for (let c = 0; c < maxCol; c++) colX[c + 1] = colX[c] + colW(c) * z;
    const styleOf = (d, c) => {
      const r = sheetRow(d), sr = r < 0 ? lastData : r;
      let i = L.s[sr + ':' + c];
      if (i == null && c >= L.maxCol) i = L.s[sr + ':' + tmplC];
      return XF[i || 0] || {};
    };
    const hl = this.hlMap && this.hlMap.get(this.file.path);
    const cellRef = (d, c) => {
      const tc = scToTc.get(c);
      const r = sheetRow(d);
      if (tc == null) return null;
      if (r < 0) return [newTr[-r - 1], tc];
      if (r === H) return [-1, tc];
      const tr = srcToTr.get(r);
      return tr != null ? [tr, tc] : null;
    };
    const valueOf = (d, c) => {
      const r = sheetRow(d), tc = scToTc.get(c);
      if (r < 0) return tc != null ? cellStr(t.rows[newTr[-r - 1]][tc]) : '';
      if (r === H && tc != null) {
        const h = String(t.headers[tc]), raw = L.v[r + ':' + c] || '';
        return /^(列|新列)\d+$/.test(h) && !raw ? '' : h;
      }
      if (tc != null) {
        const tr = srcToTr.get(r);
        if (tr != null) return cellStr(t.rows[tr][tc]);
        if (rows0.has(r)) return '';
      }
      return L.v[r + ':' + c] || '';
    };
    /* 合并单元格（换成显示坐标） */
    const merges = L.merges.map(([r1, c1, r2, c2]) => [dispRow(r1), c1, dispRow(r2) + (r1 < base && r2 >= base ? nNew : 0), c2]);
    const cover = new Map();
    merges.forEach((m, i) => {
      if ((m[2] - m[0] + 1) * (m[3] - m[1] + 1) > 20000) return;
      for (let d = m[0]; d <= m[2]; d++) for (let c = m[1]; c <= m[3]; c++) cover.set(d * 4096 + c, i);
    });
    const frD = L.fr > 0 ? dispRow(Math.min(L.fr, totalRows)) : 0, fc = Math.min(L.fc || 0, maxCol);
    const FX = colX[fc], FY = rowY[frD];
    const totalW = colX[maxCol], totalH = rowY[totalRows];

    wrap.empty();
    const vp = wrap.createDiv('ata-xlvp');
    const spacer = wrap.createDiv('ata-xlspacer');
    spacer.style.width = (GW + totalW + 40) + 'px';
    spacer.style.height = (HH + totalH + 40) + 'px';
    const mk = (cls) => vp.createDiv('ata-xlpane ' + cls);
    const corner = mk('xl-corner'), colHdr = mk('xl-colhdr'), rowHdr = mk('xl-rowhdr');
    const pTL = mk('xl-cells'), pTR = mk('xl-cells'), pBL = mk('xl-cells'), pBR = mk('xl-cells');
    corner.setAttribute('style', `left:0;top:0;width:${GW}px;height:${HH}px`);
    const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const lowerBound = (arr, n, y) => { let lo = 0, hi = n; while (lo < hi) { const mid = (lo + hi) >> 1; if (arr[mid + 1] <= y) lo = mid + 1; else hi = mid; } return lo; };
    const sel = this.xlSel && this.xlSel.file === this.file.path ? this.xlSel : null;

    const cellHtml = (d, c, x0, y0, d2, c2) => {
      const left = colX[c] - x0, top = rowY[d] - y0;
      const w = colX[(c2 == null ? c : c2) + 1] - colX[c], h = rowY[(d2 == null ? d : d2) + 1] - rowY[d];
      if (w <= 0 || h <= 0) return '';
      const x = styleOf(d, c);
      const raw = valueOf(d, c);
      const txt = raw === '' ? '' : xlFormat(raw, x);
      const st = [`left:${left}px`, `top:${top}px`, `width:${w}px`, `height:${h}px`];
      if (x.bg) st.push(`background:${x.bg}`);
      const right = x.br || styleOf(d, (c2 == null ? c : c2) + 1).bl || (L.grid && !x.bg ? '1px solid #dadce0' : '');
      const bottom = x.bb || styleOf((d2 == null ? d : d2) + 1, c).bt || (L.grid && !x.bg ? '1px solid #dadce0' : '');
      if (right) st.push(`border-right:${right}`);
      if (bottom) st.push(`border-bottom:${bottom}`);
      if (c === 0 && x.bl) st.push(`border-left:${x.bl}`);
      if (d === 0 && x.bt) st.push(`border-top:${x.bt}`);
      if (x.fc) st.push(`color:${x.fc}`);
      if (x.b) st.push('font-weight:700');
      if (x.i) st.push('font-style:italic');
      if (x.u || x.s) st.push(`text-decoration:${[x.u ? 'underline' : '', x.s ? 'line-through' : ''].join(' ')}`);
      st.push(`font-size:${((x.sz || 11) * z).toFixed(1)}pt`);
      if (x.fn) st.push(`font-family:'${x.fn.replace(/['"<>&]/g, '')}','等线','DengXian','Microsoft YaHei',sans-serif`);
      const isNum = !x.ha && txt !== '' && (x.k || /^[-+(]?[¥$€]?[\d,]+(\.\d+)?%?\)?$/.test(txt.trim()));
      const ha = x.ha || (isNum ? 'right' : 'left');
      st.push(`justify-content:${ha === 'center' || ha === 'centerContinuous' || ha === 'distributed' ? 'center' : ha === 'right' ? 'flex-end' : 'flex-start'}`);
      st.push(`align-items:${x.va === 'top' ? 'flex-start' : x.va === 'center' || x.va === 'distributed' || x.va === 'justify' ? 'center' : 'flex-end'}`);
      if (x.ind) st.push(`padding-left:${x.ind * 9 * z + 2}px`);
      let cls = 'xc';
      let inner = esc(txt);
      if (x.w) cls += ' xw';
      else if (txt && ha === 'left' && d2 == null) {
        /* 文字超出格子时，像 Excel 一样延伸到右边的空格子里 */
        let ext = w, k = c + 1;
        while (k < maxCol && k < c + 8 && !cover.has(d * 4096 + k) && valueOf(d, k) === '') { ext += colX[k + 1] - colX[k]; k++; }
        if (ext > w) { cls += ' xo'; inner = `<span style="max-width:${ext - 4}px">${inner}</span>`; }
      }
      const ref = cellRef(d, c);
      if (hl && ref && ref[0] >= 0 && (hl.added.has(ref[0]) || (hl.cells.get(ref[0]) && hl.cells.get(ref[0]).has(ref[1])))) cls += hl.added.has(ref[0]) ? ' xadd' : ' xmod';
      if (sel && sel.d === d && sel.c === c) cls += ' xsel';
      return `<div class="${cls}" data-d="${d}" data-c="${c}" style="${st.join(';')}">${inner}</div>`;
    };

    const drawPane = (el, rA, rB, cA, cB, sx, sy, px, py, pw, ph) => {
      el.style.cssText = `left:${px}px;top:${py}px;width:${Math.max(0, pw)}px;height:${Math.max(0, ph)}px`;
      if (pw <= 0 || ph <= 0 || rA >= rB || cA >= cB) { el.innerHTML = ''; return; }
      const x0 = colX[cA] + sx, y0 = rowY[rA] + sy;
      const d0 = Math.max(rA, lowerBound(rowY, totalRows, y0)), c0 = Math.max(cA, lowerBound(colX, maxCol, x0));
      let d1 = d0; while (d1 < rB && rowY[d1] < y0 + ph) d1++;
      let c1 = c0; while (c1 < cB && colX[c1] < x0 + pw) c1++;
      const out = [], done = new Set();
      for (let d = d0; d < d1; d++) {
        if (rowY[d + 1] === rowY[d]) continue;
        for (let c = c0; c < c1; c++) {
          if (colX[c + 1] === colX[c]) continue;
          const mi = cover.get(d * 4096 + c);
          if (mi != null) {
            if (done.has(mi)) continue;
            done.add(mi);
            const m = merges[mi];
            out.push(cellHtml(m[0], m[1], x0, y0, m[2], m[3]));
            continue;
          }
          out.push(cellHtml(d, c, x0, y0));
        }
      }
      el.innerHTML = out.join('');
    };

    const draw = () => {
      if (!wrap.isConnected) return;
      const W = wrap.clientWidth, Hh = wrap.clientHeight;
      vp.style.width = W + 'px';
      vp.style.height = Hh + 'px';
      vp.style.marginBottom = -Hh + 'px';
      const sx = wrap.scrollLeft, sy = wrap.scrollTop;
      const aw = W - GW, ah = Hh - HH;
      drawPane(pTL, 0, frD, 0, fc, 0, 0, GW, HH, Math.min(FX, aw), Math.min(FY, ah));
      drawPane(pTR, 0, frD, fc, maxCol, sx, 0, GW + FX, HH, aw - FX, Math.min(FY, ah));
      drawPane(pBL, frD, totalRows, 0, fc, 0, sy, GW, HH + FY, Math.min(FX, aw), ah - FY);
      drawPane(pBR, frD, totalRows, fc, maxCol, sx, sy, GW + FX, HH + FY, aw - FX, ah - FY);
      /* 列标 A B C… */
      const ch = [];
      const colCell = (c, left) => `<div class="xh${sel && sel.c === c ? ' xhs' : ''}" style="left:${left}px;width:${colX[c + 1] - colX[c]}px">${colLetter(c)}</div>`;
      for (let c = 0; c < fc; c++) if (colX[c + 1] > colX[c]) ch.push(colCell(c, colX[c]));
      { const x0 = colX[fc] + sx; let c = Math.max(fc, lowerBound(colX, maxCol, x0)); for (; c < maxCol && colX[c] < x0 + aw - FX; c++) if (colX[c + 1] > colX[c]) ch.push(colCell(c, FX + colX[c] - x0)); }
      colHdr.style.cssText = `left:${GW}px;top:0;width:${aw}px;height:${HH}px`;
      colHdr.innerHTML = ch.join('');
      /* 行号 */
      const rh = [];
      const rowCell = (d, top) => { const r = sheetRow(d); return `<div class="xr${r < 0 ? ' xrn' : ''}${sel && sel.d === d ? ' xhs' : ''}" style="top:${top}px;height:${rowY[d + 1] - rowY[d]}px">${r < 0 ? '新' : r + 1}</div>`; };
      for (let d = 0; d < frD; d++) if (rowY[d + 1] > rowY[d]) rh.push(rowCell(d, rowY[d]));
      { const y0 = rowY[frD] + sy; let d = Math.max(frD, lowerBound(rowY, totalRows, y0)); for (; d < totalRows && rowY[d] < y0 + ah - FY; d++) if (rowY[d + 1] > rowY[d]) rh.push(rowCell(d, FY + rowY[d] - y0)); }
      rowHdr.style.cssText = `left:0;top:${HH}px;width:${GW}px;height:${ah}px`;
      rowHdr.innerHTML = rh.join('');
    };
    let raf = 0;
    wrap.onscroll = () => { if (!raf) raf = requestAnimationFrame(() => { raf = 0; draw(); }); };
    this.xlRedraw = draw;
    this.redrawRows = draw;
    vp.onclick = (e) => {
      const el = e.target.closest && e.target.closest('.xc');
      if (!el) return;
      const d = +el.dataset.d, c = +el.dataset.c;
      this.xlSel = { file: this.file.path, d, c };
      const r = sheetRow(d);
      const raw = valueOf(d, c);
      const ref = cellRef(d, c);
      if (this.statEl) this.statEl.setText(`${colLetter(c)}${r < 0 ? '（新增行）' : r + 1}：${raw === '' ? '（空）' : raw}${ref && ref[0] >= 0 ? `　· 表格第 ${ref[0] + 1} 行「${t.headers[ref[1]]}」，双击可修改` : ref ? '　· 列名，双击可修改' : ''}`);
      vp.querySelectorAll('.xsel').forEach((x) => x.classList.remove('xsel'));
      el.classList.add('xsel');
    };
    vp.ondblclick = (e) => {
      const el = e.target.closest && e.target.closest('.xc');
      if (!el) return;
      const ref = cellRef(+el.dataset.d, +el.dataset.c);
      if (!ref) { new Notice('这个格子不在表格数据范围里（例如标题、备注、合计），只能在 Excel 里修改。'); return; }
      this.focusAfter = ref;
      this.setViewMode('data');
    };
    requestAnimationFrame(draw);
    const shown = t.rows.length;
    if (this.statEl) this.statEl.setText(`原表样式 · ${shown.toLocaleString()} 行数据${nNew ? `（其中 ${nNew} 行是新增的，排在原表数据后面）` : ''} · 点格子查看，双击修改`);
  }

  /* ================= 外部修改：快照、对比、确认、撤销 ================= */
  isTableFile(f) {
    if (!(f instanceof TFile)) return false;
    const folder = normalizePath(this.settings.folder) + '/';
    if (f.path.startsWith(this.archiveDir() + '/') || f.path.startsWith(folder + '导出/') || f.path.startsWith(folder + AUTO_DIR + '/')) return false;
    if (f.path.startsWith(folder) && (GUIDE_FILES.includes(f.path.slice(folder.length)) || DOC_FILES.includes(f.path.slice(folder.length)))) return false;
    if (f.path.startsWith(folder + OUT_DIR + '/')) return false;
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
      catch (e) { console.warn('[表格工作台] 保存快照索引失败', e); }
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
    } catch (err) { console.warn('[表格工作台] 保存快照失败', err); }
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
    const next = prev.then(fn).catch((e) => console.warn('[表格工作台] 检查外部修改失败', e));
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
          this.logOp({ type: 'ext-new', path: f.path, text: `新建了表格（${t.rows.length} 行，列：${t.headers.slice(0, 8).join('、')}${t.headers.length > 8 ? '…' : ''}）`, review: 'ok' });
          this.scheduleOutput(f);
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
      } catch (e) { console.warn('[表格工作台] 保存修改前版本失败', e); hist = ''; }
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
    this.scheduleOutput(f);
    this.scheduleDocs();
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
      type: 'ext', path: f.path, text: parts.join('，'), review: 'ok', hist, after: hashStr(newText),
      cells: d.cells.slice(0, 300).map((x) => ({ r: x.r + 1, h: hdr[x.c], o: x.o, n: x.n })), cellsN: d.cells.length,
      hdr, added: d.added.slice(0, 30).map((r) => newT.rows[r].map(cellStr)), addedN: d.added.length,
      oldHdr: oldT.headers.map(String), rows: d.removed.slice(0, 30).map((r) => oldT.rows[r].map(cellStr)), rowsN: d.removed.length,
      items,
    });
    new Notice(`「${this.tableLabel(f.path)}」${parts.join('，')}，原表 Excel 稍后自动更新。`, 6000);
    if (this.file && this.file.path === f.path && this.tab === 'workspace') { this.renderTable(); this.renderReview(); }
    else this.scheduleRender();
  }

  pendingFor(path) {
    return (this.plugin.opLog || []).filter((e) => e.type === 'ext' && e.path === path && e.review === 'pending');
  }

  async confirmExternal(path) {
    this.pendingFor(path).forEach((e) => { e.review = 'ok'; });
    if (this.hlMap) this.hlMap.delete(path);
    this.saveOpLog();
    if (this.file && this.file.path === path && this.tab === 'workspace') { this.renderTable(); this.renderReview(); }
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
      (this.plugin.opLog || []).forEach((x) => { if (x.type === 'ext' && x.path === f.path && x.review !== 'undone' && x.t >= e.t) x.review = 'undone'; });
      e.review = 'undone';
      if (this.hlMap) this.hlMap.delete(f.path);
      this.logOp({ type: 'undo', path: f.path, text: `撤销了外部修改（${e.text}）` });
      if (this.file && this.file.path === f.path) {
        this.table = this.parseTableText(f, h.text);
        if (link) { this.link = link; this.attachProvenance(this.table, link.map); }
        this.tableVersion = (this.tableVersion || 0) + 1;
      }
      new Notice('已撤销这次修改，表格恢复到修改前，原表 Excel 也会跟着更新。');
      this.scheduleOutput(f);
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
    this.ico(el, 'file-diff', 'ata-review-icon');
    const tx = el.createDiv('ata-review-text');
    const tm = new Date(last.t);
    tx.createEl('b', { text: `外部修改了这张表${pend.length > 1 ? `（${pend.length} 次）` : ''}` });
    tx.createSpan({ text: ` · ${String(tm.getHours()).padStart(2, '0')}:${String(tm.getMinutes()).padStart(2, '0')} · ${last.text}。绿色是新增的行，黄色是改过的格子。` });
    const bar = el.createDiv('ata-review-btns');
    this.btn(bar, '查看明细', 'list', () => { this.tab = 'home'; this.opKind = '外部修改'; if (!this.opOpen) this.opOpen = new Set(); this.opOpen.add(last.t + '|' + last.type + '|' + (last.path || '')); this.render(); }, 'ata-btn-sm');
    this.btn(bar, '撤销', 'undo-2', () => this.undoExternal(last), 'ata-btn-sm');
    this.btn(bar, '确认', 'check', () => this.confirmExternal(this.file.path), 'ata-btn-sm mod-cta');
  }

  /* ================= 复制路径 ================= */
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





  async copyText(text, what) {
    try { await navigator.clipboard.writeText(text); new Notice(`已复制${what || ''}。`); }
    catch (e) {
      try { require('electron').clipboard.writeText(text); new Notice(`已复制${what || ''}。`); }
      catch (err) { new Notice('复制失败：' + (e.message || e)); }
    }
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
      } catch (e) { console.warn('[表格工作台] 保存操作记录失败', e); }
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
    this.scheduleDocs();
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
    const KINDS = [['全部', null], ['表格修改', ['ext', 'ext-new', 'ext-del', 'edit', 'row', 'col']], ['处理说明', ['note']], ['原表更新失败', ['output']], ['旧版记录', ['ai']], ['撤销', ['undo']], ['导入导出', ['import', 'export', 'create']], ['附件', ['archive']], ['删除', ['delete']]];
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
      box.createDiv({ cls: 'ata-empty-sm', text: all.length ? '没有符合条件的记录。' : '还没有操作记录。之后导入、外部修改、手动修改、导出、删除都会记在这里，可以随时回来查改了什么。' });
      return;
    }
    const TYPE = { note: ['说明', 'message-square-text'], output: ['原表', 'file-spreadsheet'], ext: ['修改', 'file-diff'], 'ext-new': ['新建', 'file-plus'], 'ext-del': ['删除', 'file-x'], ai: ['旧版', 'history'], edit: ['手动', 'pencil'], row: ['行', 'rows-3'], col: ['列', 'columns-3'], undo: ['撤销', 'undo-2'], import: ['导入', 'folder-down'], export: ['导出', 'download'], create: ['新建', 'file-plus'], delete: ['删除', 'trash-2'], archive: ['附件', 'archive'] };
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
      if (e.review === 'undone') line.createSpan({ cls: 'ata-tl-review r-undone', text: '已撤销' });
      if (e.output) {
        const oc = line.createSpan({ cls: 'ata-tl-review r-ok', text: '原表已更新', attr: { title: e.outFile ? this.absPath(e.outFile) : '' } });
        if (e.outFile) { oc.addClass('is-link'); oc.addEventListener('click', () => this.app.openWithDefaultApp && this.app.openWithDefaultApp(e.outFile)); }
      }
      if (e.error) line.createSpan({ cls: 'ata-tl-review r-err', text: '失败' });
      if (e.type === 'ext' && e.hist && e.review !== 'undone') {
        const latest = (this.plugin.opLog || []).filter((x) => x.type === 'ext' && x.path === e.path && x.review !== 'undone').slice(-1)[0] === e;
        if (latest) {
          const rb = main.createDiv('ata-tl-rbtns');
          const un = rb.createEl('button', { cls: 'ata-tl-rbtn', text: '撤销这次修改' });
          un.addEventListener('click', () => this.undoExternal(e));
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
    } catch (e) { console.warn('[表格工作台] 保存对应关系失败', e); }
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
    if (this.isExcelMode()) { this.vtFor = null; this.renderExcel(); return; }
    this.xlToken = (this.xlToken || 0) + 1;
    this.previewEl.removeClass('ata-xlwrap');
    const key = this.file && this.file.path;
    const keep = this.vtFor === key ? { top: this.previewEl.scrollTop, left: this.previewEl.scrollLeft } : null;
    this.vtFor = key;
    this.previewEl.empty();
    this.previewEl.onscroll = null;
    const t = this.table;
    if (!t.headers.length) {
      this.previewEl.createDiv({ cls: 'ata-empty-sm ata-pad', text: '这是一个空表。可以点下方「新增列」开始建表。' });
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
    /* 从原表样式双击过来：跳到那个格子开始编辑 */
    if (this.focusAfter) {
      const [ri, ci] = this.focusAfter;
      this.focusAfter = null;
      if (ri >= 0) this.previewEl.scrollTop = Math.max(0, ri * ROW_H - this.previewEl.clientHeight / 3);
      draw(true);
      setTimeout(() => {
        const el = ri < 0 ? this.previewEl.querySelector(`.ata-th-label[data-c="${ci}"]`) : this.previewEl.querySelector(`td[data-r="${ri}"][data-c="${ci}"]`);
        if (el) { el.scrollIntoView({ block: 'nearest', inline: 'nearest' }); el.focus(); }
      }, 30);
    }
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
    this.scheduleOutput(this.file);
  }

  async handleDrop(list, zone) {
    if (!list.length) return;
    if (zone && zone.startsWith('folder:')) { const keep = this.archFolder; this.archFolder = zone.slice(7); await this.archiveSave(list.map((x) => x.file).filter((f) => f.size > 0), '拖入'); if (keep !== undefined) this.archFolder = keep; return; }
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
    const folder = this.tab === 'archive' && this.archFolder ? `${this.archiveDir()}/${this.archFolder}` : `${this.archiveDir()}/${d.getFullYear()}-${z(d.getMonth() + 1)}`;
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
    const newOnes = saved.filter((f) => meta[f.path] && !meta[f.path].scan && Date.now() - meta[f.path].added < 600000 && scannable(f));
    if (newOnes.length && scanCfg(this.settings).auto && scanReady(this.settings)) this.scanArchive(newOnes, true);
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

  /* ================= 附件 AI 扫描 ================= */
  dateFolderFor(date) {
    if (!date) return '';
    const ym = date.slice(0, 7);
    if (scanCfg(this.settings).by === 'day' && date.length === 10) return `${ym}/${date.slice(5)}`;
    return ym;
  }

  openScanSettings() {
    new Notice('先在设置里填好「附件 AI 扫描」（选服务商、填 API Key）。');
    try { this.app.setting.open(); this.app.setting.openTabById(this.plugin.manifest.id); } catch (e) { /* 旧版本 */ }
  }

  /** 确认后扫描；force=true 时已经识别过的也重新识别 */
  askScan(files, force) {
    if (!scanReady(this.settings)) { this.openScanSettings(); return; }
    const meta = this.settings.archiveMeta || {};
    const ok = files.filter(scannable);
    const todo = ok.filter((f) => force || !(meta[f.path] && meta[f.path].scan && meta[f.path].scan.date));
    const skipped = files.length - ok.length;
    if (!todo.length) { new Notice(skipped && !ok.length ? `选中的 ${skipped} 个文件不是图片或 PDF，AI 识别不了。` : '这里的图片 / PDF 都已经识别过了。要重新识别，点「多选」选中后再点「AI 识别日期」。'); return; }
    const c = scanCfg(this.settings);
    const how = c.by === 'day' ? '按「年-月 / 月-日」文件夹' : '按「年-月」文件夹';
    new ConfirmModal(this.app, `AI 识别 ${todo.length} 个文件的日期并按时间归类？`,
      `会把图片发给「${(SCAN_PRESETS[c.provider] || {}).label || c.url}」（模型 ${c.model}）识别单据上的日期，然后${how}整理${c.rename ? '，并按识别内容改名（格式：' + c.nameTpl + '）' : ''}。\n在你自己建的分类文件夹（不是日期文件夹）里的文件只记录日期，不移动。${skipped ? `\n${skipped} 个不是图片 / PDF 的文件跳过。` : ''}`,
      [{ text: `开始识别 ${todo.length} 个`, cta: true, onClick: () => this.scanArchive(todo) }]).open();
  }

  /** 识别日期 → 记在附件信息里 → 移到日期文件夹（可选重命名） */
  async scanArchive(files, quiet) {
    const cfg = scanCfg(this.settings);
    if (!scanReady(this.settings) || !files.length) return;
    if (this.scanRunning) { this.scanQueue = (this.scanQueue || []).concat(files); if (!quiet) new Notice('正在识别上一批，这批排在后面。'); return; }
    this.scanRunning = true;
    const meta = this.settings.archiveMeta || (this.settings.archiveMeta = {});
    const notice = new Notice(`AI 正在识别附件日期 0/${files.length}…`, 0);
    const res = [];
    let done = 0, next = 0;
    const z = (n) => String(n).padStart(2, '0');
    const worker = async () => {
      while (next < files.length) {
        const tf = files[next++];
        const m = meta[tf.path] || (meta[tf.path] = { added: tf.stat.ctime || Date.now() });
        try {
          const rc = new Date(m.added || tf.stat.ctime || tf.stat.mtime || Date.now());
          const img = await imageToDataUrl(this.app, tf);
          const r = parseScanReply(await callVision(cfg, scanPrompt(tf, `${rc.getFullYear()}-${z(rc.getMonth() + 1)}-${z(rc.getDate())}`), img));
          m.scan = Object.assign(r, { at: Date.now(), model: cfg.model });
          if (!m.note) m.note = [r.type, r.party, r.no].filter(Boolean).join(' ');
          res.push({ tf, r, name0: tf.name });
        } catch (e) {
          m.scan = { error: String(e.message || e).slice(0, 160), at: Date.now() };
          res.push({ tf, err: m.scan.error, name0: tf.name });
          if (/检查 API Key|连不上/.test(m.scan.error)) next = files.length; // Key 错了或断网就停下
        }
        done++;
        try { notice.setMessage(`AI 正在识别附件日期 ${done}/${files.length}…`); } catch (e) { /* 忽略 */ }
      }
    };
    try {
      await Promise.all([worker(), worker(), worker()]);
      const moved = {};
      let renamed = 0;
      this.archiving = true;
      try {
        for (const { tf, r } of res) {
          if (!r || !this.app.vault.getAbstractFileByPath(tf.path)) continue;
          if (!r.date && !(cfg.rename && (r.party || r.no))) continue;
          const cur = this.folderOf(tf);
          const target = r.date && isDateFolder(cur) ? this.dateFolderFor(r.date) : cur;
          const base0 = tf.basename || tf.name.replace(/\.[^.]+$/, '');
          let base = base0;
          if (cfg.rename) base = scanFileName(cfg.nameTpl, r, (meta[tf.path] && meta[tf.path].origName) || base0) || base0;
          if (cfg.rename && meta[tf.path] && !meta[tf.path].origName && base !== base0) meta[tf.path].origName = base0;
          if (target === cur && base === base0) continue;
          const dest = target ? `${this.archiveDir()}/${target}` : this.archiveDir();
          await this.ensureFolder(dest);
          const np = await this.uniquePath(dest, base, tf.extension || 'bin');
          const old = tf.path;
          await this.app.fileManager.renameFile(tf, np);
          if (meta[old]) { meta[np] = meta[old]; delete meta[old]; }
          if (this.archSel && this.archSel.has(old)) { this.archSel.delete(old); this.archSel.add(np); }
          if (target !== cur) moved[target || '附件区'] = (moved[target || '附件区'] || 0) + 1;
          if (base !== base0) renamed++;
        }
      } finally { this.archiving = false; }
      await this.plugin.saveSettings();
      const ok = res.filter((x) => x.r && x.r.date).length;
      const noDate = res.filter((x) => x.r && !x.r.date).length;
      const bad = res.filter((x) => x.err);
      const where = Object.entries(moved).sort().map(([k, n]) => `${k}（${n}）`).join('、');
      const parts = [`识别出 ${ok} 个日期`];
      if (where) parts.push(`归入 ${where}`);
      if (renamed) parts.push(`改名 ${renamed} 个`);
      if (noDate) parts.push(`${noDate} 个看不出日期`);
      if (bad.length) parts.push(`${bad.length} 个失败`);
      if (files.length > res.length) parts.push(`${files.length - res.length} 个没识别（已停止）`);
      try { notice.hide(); } catch (e) { /* 忽略 */ }
      new Notice('AI 扫描完成：' + parts.join('，') + (bad.length ? `\n失败原因：${bad[0].err}` : ''), bad.length ? 12000 : 6000);
      this.logOp({ type: 'archive', text: `AI 扫描 ${res.length} 个附件：${parts.join('，')}`, items: res.map((x) => `${x.name0}${x.r ? `：${x.r.date || '无日期'}${x.r.type ? ' ' + x.r.type : ''}${x.tf.name !== x.name0 ? ' → ' + (this.folderOf(x.tf) ? this.folderOf(x.tf) + '/' : '') + x.tf.name : ''}` : '：失败（' + x.err + '）'}`) });
    } finally {
      this.scanRunning = false;
      try { notice.hide(); } catch (e) { /* 忽略 */ }
    }
    if (this.tab === 'archive') await this.render();
    const q = this.scanQueue;
    this.scanQueue = null;
    if (q && q.length) await this.scanArchive(q.filter((f) => this.app.vault.getAbstractFileByPath(f.path)), true);
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
    left.createDiv({ cls: 'ata-sub', text: `${all.length} 个文件 · 共 ${fmtSize(all.reduce((a, f) => a + f.stat.size, 0))} · 收到的单据、截图、PDF、Excel 都可以存在这里，可以新建文件夹分类` });
    const right = head.createDiv('ata-header-right');
    this.btn(right, this.archSelect ? '完成' : '多选', this.archSelect ? 'check' : 'list-checks', () => { this.archSelect = !this.archSelect; this.archSel.clear(); this.render(); });
    this.btn(right, 'AI 按时间整理', 'sparkles', () => this.askScan(this.archSelect && this.archSel.size ? all.filter((f) => this.archSel.has(f.path)) : this.archList || all, this.archSelect && this.archSel.size > 0));
    this.btn(right, '打开文件夹', 'folder-open', async () => { const d = this.archFolder ? `${this.archiveDir()}/${this.archFolder}` : this.archiveDir(); await this.ensureFolder(d); this.showInSystem(d); });
    this.btn(right, '上传文件', 'upload', () => this.archiveInput.click(), 'mod-cta');

    const dz = wrap.createDiv('ata-dropzone ata-archive-drop');
    this.ico(dz, 'archive', 'ata-dropzone-icon');
    const txt = dz.createDiv('ata-dropzone-text');
    txt.createEl('b', { text: '把文件拖到这里存档（任何类型都可以）' });
    txt.createSpan({ text: `也可以点这里选择文件，或在本页按 Ctrl+V 粘贴复制的文件 / 截图 / 文件路径。${this.archFolder ? `存进当前文件夹「${this.archFolder}」。` : '没选文件夹时按月份归档。'}` });
    dz.addEventListener('click', () => this.archiveInput.click());

    /* 文件夹 */
    const folders = this.archiveFolders();
    if (this.archFolder && !folders.includes(this.archFolder)) this.archFolder = '';
    const inFolder = (f, fd) => { const x = this.folderOf(f); return x === fd || x.startsWith(fd + '/'); };
    const fdbar = wrap.createDiv('ata-folderbar');
    const fchip = (label, fd, n) => {
      const b = fdbar.createEl('button', { cls: 'ata-folderchip' + ((this.archFolder || '') === fd ? ' is-active' : '') });
      this.ico(b, fd ? 'folder' : 'folders');
      b.createSpan({ text: label });
      b.createSpan({ cls: 'ata-tab-count', text: String(n) });
      b.addEventListener('click', () => { this.archFolder = fd; this.render(); });
      if (fd) {
        b.setAttribute('data-drop-folder', fd);
        b.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          const cnt = all.filter((f) => inFolder(f, fd)).length;
          if (cnt) { new Notice(`「${fd}」里还有 ${cnt} 个文件，先移走或删掉才能删除文件夹。`); return; }
          new ConfirmModal(this.app, `删除空文件夹「${fd}」？`, '', [{ text: '删除', warning: true, onClick: async () => {
            const fo = this.app.vault.getAbstractFileByPath(`${this.archiveDir()}/${fd}`);
            try { if (fo) await this.app.vault.delete(fo, true); } catch (err) { /* 已经没有了 */ }
            this.settings.archiveFolders = (this.settings.archiveFolders || []).filter((x) => x !== fd);
            await this.plugin.saveSettings();
            this.archFolder = '';
            this.render();
          } }]).open();
        });
      }
    };
    fchip('全部', '', all.length);
    folders.forEach((fd) => fchip(fd, fd, all.filter((f) => inFolder(f, fd)).length));
    const nb = fdbar.createEl('button', { cls: 'ata-folderchip is-new' });
    this.ico(nb, 'folder-plus');
    nb.createSpan({ text: '新建文件夹' });
    nb.addEventListener('click', () => this.newArchiveFolder());
    if (folders.length) fdbar.createSpan({ cls: 'ata-muted ata-folderhint', text: '右键空文件夹可以删除' });

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
    let list = all.filter((f) => (!this.archKind || archiveKind(f) === this.archKind) && (!this.archFolder || inFolder(f, this.archFolder)));
    const docDate = (f) => { const x = meta[f.path] && meta[f.path].scan; return (x && x.date) || ''; };
    const z2 = (n) => String(n).padStart(2, '0');
    const recvDay = (f) => { const d = new Date((meta[f.path] && meta[f.path].added) || f.stat.ctime || f.stat.mtime); return `${d.getFullYear()}-${z2(d.getMonth() + 1)}-${z2(d.getDate())}`; };
    const hasDoc = all.some((f) => docDate(f));
    const mode = this.archSort || (hasDoc ? 'doc' : 'recv');
    /* 按日期分组：同一天的放在一起，中间用日期标题隔开 */
    const groupKey = mode === 'doc' ? (f) => docDate(f) || '~' : recvDay;
    list = list.slice().sort((a, b) => { const ka = groupKey(a), kb = groupKey(b); return ka === kb ? b.stat.ctime - a.stat.ctime : ka === '~' ? 1 : kb === '~' ? -1 : kb.localeCompare(ka); });
    this.archList = list;
    const sortbar = fbar.createDiv('ata-catbar ata-sortbar');
    [['按单据日期分组', 'doc'], ['按收到日期分组', 'recv']].forEach(([label, k]) => {
      if (k === 'doc' && !hasDoc) return;
      const b = sortbar.createEl('button', { cls: 'ata-catchip' + (mode === k ? ' is-active' : ''), text: label });
      b.addEventListener('click', () => { this.archSort = k; this.render(); });
    });
    const applyQ = () => {
      const s = (this.archQuery || '').toLowerCase();
      const vis = new Set();
      grid.querySelectorAll('.ata-acard').forEach((c) => { const on = !s || c.dataset.q.includes(s); c.style.display = on ? '' : 'none'; if (on) vis.add(c.dataset.g); });
      grid.querySelectorAll('.ata-agroup').forEach((h) => { h.style.display = vis.has(h.dataset.g) ? '' : 'none'; });
    };
    q.addEventListener('input', () => { this.archQuery = q.value.trim(); applyQ(); });

    if (this.archSelect) {
      const sb = wrap.createDiv('ata-selbar');
      const cnt = sb.createSpan({ cls: 'ata-selcount', text: `已选 ${this.archSel.size} 个` });
      this.archCountEl = cnt;
      const sel = () => all.filter((f) => this.archSel.has(f.path));
      this.btn(sb, '全选', 'check-square', () => { list.forEach((f) => this.archSel.add(f.path)); this.render(); }, 'ata-btn-sm');
      this.btn(sb, '复制文件', 'files', () => { const s = sel(); if (s.length) this.copyFiles(s.map((f) => f.path)); }, 'ata-btn-sm mod-cta');
      this.btn(sb, '复制路径', 'copy', () => { const s = sel(); if (s.length) this.copyPaths(s.map((f) => f.path), '路径'); }, 'ata-btn-sm');
      this.btn(sb, '移动到…', 'folder-input', () => { const s = sel(); if (s.length) this.pickArchiveFolder((fd) => this.moveArchive(s, fd)); }, 'ata-btn-sm');
      this.btn(sb, 'AI 识别日期', 'sparkles', () => { const s = sel(); if (s.length) this.askScan(s, true); }, 'ata-btn-sm');
      this.btn(sb, '删除所选', 'trash-2', () => {
        const s = sel();
        if (!s.length) return;
        new ConfirmModal(this.app, `删除 ${s.length} 个附件？`, s.slice(0, 8).map((f) => '• ' + f.name).join('\n') + '\n会移到回收站，需要时可以恢复。', [{ text: `删除 ${s.length} 个`, warning: true, onClick: () => this.archiveDelete(s) }]).open();
      }, 'ata-btn-sm is-danger');
    }

    const grid = wrap.createDiv('ata-agrid');
    const WEEK = '日一二三四五六';
    const shown = list.slice(0, 600);
    const groupCount = {};
    shown.forEach((f) => { const k = groupKey(f); groupCount[k] = (groupCount[k] || 0) + 1; });
    let lastKey = null, gi = 0;
    shown.forEach((f) => {
      const gk = groupKey(f);
      if (gk !== lastKey) {
        lastKey = gk; gi++;
        const gh = grid.createDiv({ cls: 'ata-agroup' + (gk === '~' ? ' is-none' : ''), attr: { 'data-g': String(gi) } });
        this.ico(gh, gk === '~' ? 'help-circle' : 'calendar-days');
        let label = gk;
        if (gk === '~') label = '没有单据日期（还没识别或没认出）';
        else if (gk.length === 10) { const d = new Date(gk + 'T00:00:00'); label = `${gk}  周${WEEK[d.getDay()]}`; }
        else label = `${gk}（只认出月份）`;
        gh.createSpan({ cls: 'ata-agroup-date', text: label });
        gh.createSpan({ cls: 'ata-tab-count', text: String(groupCount[gk]) });
        if (this.archSelect) {
          const keyNow = gk;
          const b = gh.createEl('button', { cls: 'ata-agroup-sel', text: '选中这一组' });
          b.addEventListener('click', () => { list.filter((x) => groupKey(x) === keyNow).forEach((x) => this.archSel.add(x.path)); this.render(); });
        }
      }
      const m = meta[f.path] || {};
      const c = grid.createDiv('ata-acard' + (this.archSel.has(f.path) ? ' is-selected' : ''));
      c.dataset.g = String(gi);
      c.dataset.q = (f.name + ' ' + (m.note || '') + ' ' + (m.scan ? [m.scan.date, m.scan.type, m.scan.party, m.scan.no, m.scan.summary].join(' ') : '')).toLowerCase();
      const th = c.createDiv('ata-athumb');
      this.renderThumb(th, f);
      th.addEventListener('click', () => { if (!this.archSelect) new LightboxModal(this.app, this, f).open(); });
      const info = c.createDiv('ata-ainfo');
      info.createDiv({ cls: 'ata-aname', text: f.name, attr: { title: f.path } });
      info.createDiv({ cls: 'ata-asub', text: `${fmtSize(f.stat.size)} · ${fmtTime(m.added || f.stat.ctime || f.stat.mtime)}${m.from ? ' · ' + m.from : ''}` });
      if (m.scan) {
        const sd = info.createDiv({ cls: 'ata-ascan' + (m.scan.error ? ' is-error' : m.scan.date ? '' : ' is-none') });
        this.ico(sd, m.scan.error ? 'alert-circle' : 'calendar');
        sd.createSpan({ text: m.scan.error ? '识别失败' : m.scan.date ? `单据日期 ${m.scan.date}${m.scan.type ? ' · ' + m.scan.type : ''}` : '没认出日期' });
        sd.setAttribute('title', m.scan.error || [m.scan.party, m.scan.no, m.scan.summary].filter(Boolean).join(' · '));
      }
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
      const op = acts.createEl('button', { cls: 'ata-abtn', text: '打开' });
      op.addEventListener('click', (e) => { e.stopPropagation(); if (this.app.openWithDefaultApp) this.app.openWithDefaultApp(f.path); else this.app.workspace.getLeaf('tab').openFile(f); });
      const cp = acts.createEl('button', { cls: 'ata-abtn', text: '复制路径' });
      cp.addEventListener('click', (e) => { e.stopPropagation(); this.copyText(this.absPath(f.path), '文件路径'); });
      const sp = c.createDiv('ata-atools');
      if (archiveKind(f) === 'image') this.iconBtn(sp, 'copy', '复制图片', () => copyImageNotice(this.app, f));
      this.iconBtn(sp, 'folder-input', '移动到文件夹', () => this.pickArchiveFolder((fd) => this.moveArchive([f], fd)));
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

  async importBatch(list, opts) {
    opts = opts || {};
    if (this.importRun) {
      if (!opts.wait) { new Notice('正在导入中，请等这一批完成或先取消。'); return; }
      while (this.importRun) await new Promise((r) => setTimeout(r, 300));
    }
    const tables = list.filter((x) => TABLE_EXT.test(x.file.name) && !x.file.name.startsWith('~$'));
    const skipped = list.length - tables.length;
    if (!tables.length) { new Notice('没有找到可导入的表格文件（支持 csv / tsv / xlsx / md）。'); return; }
    const root = this.dataDir();
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
                console.warn('[表格工作台] 保存原 Excel 失败，导出时将不能保留原格式', e);
                linkId = null;
              }
            }
            for (const sh of sheets) {
              const path = this.uniquePathSync(folder, multi ? baseName(sh.name) : book, 'csv', taken);
              const csv = toCSV(sh.table);
              await this.snapPut(path, csv);
              lastFile = await this.app.vault.create(path, csv);
              (run.paths || (run.paths = [])).push(path);
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
    /* 每个导入的表格马上生成一份原表 Excel */
    const doneGroups = new Set();
    (run.paths || []).forEach((p) => {
      const f = this.app.vault.getAbstractFileByPath(p);
      if (!(f instanceof TFile)) return;
      const dest = this.outPathOf(f);
      if (doneGroups.has(dest)) return;
      doneGroups.add(dest);
      this.scheduleOutput(f, 500);
    });
    this.scheduleDocs();
    if (ok || failed.length) this.logOp({ type: 'import', text: msg.join('，') + (opts.from ? `（${opts.from}）` : ''), items: (run.names || []).slice(0, 300) });
    if (opts.quiet) { await this.render(); return { ok, failed, paths: run.paths || [], names: run.names || [] }; }
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
      await this.ensureFolder(this.dataDir());
      const path = await this.uniquePath(this.dataDir(), baseName(name), fmt);
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
  async exportLinked(file, files, opts) {
    opts = opts || {};
    const notice = opts.silent ? { hide() {} } : new Notice('正在按原 Excel 格式导出…', 0);
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
      const path = opts.dest || await this.uniquePath(this.settings.folder + '/导出', `${bookName}-${stamp}`, 'xlsx');
      await this.writeBinaryTo(path, res.buffer);
      notice.hide();
      const msg = `已按原格式导出到 ${path}：${res.sheetsChanged ? `${res.sheetsChanged} 个工作表有改动，共 ${res.changed} 个单元格` : '没有改动，和原文件一致'}` +
        (res.formulasDropped ? `；有 ${res.formulasDropped} 个共享公式因所在单元格被改动而转成了数值` : '');
      if (!opts.silent) new Notice(msg, 10000);
      if (!opts.nolog) this.logOp({ type: 'export', path: file.path, name: bookName, text: `按原格式导出 Excel${res.sheetsChanged ? `（${res.sheetsChanged} 个工作表有改动，${res.changed} 个单元格）` : '（没有改动）'}${opts.from ? `（${opts.from}）` : ''}`, file: path });
      if (!opts.quiet && this.settings.openAfterExport && this.app.openWithDefaultApp) {
        try { await this.app.openWithDefaultApp(path); } catch (e) { /* 打不开就只保存 */ }
      }
      return { path, changed: res.changed, sheetsChanged: res.sheetsChanged };
    } catch (e) {
      notice.hide();
      if (opts.quiet) throw e;
      new Notice('按原格式导出失败：' + (e.message || e) + '。可以改用普通 Excel 导出。', 10000);
    }
  }

  async exportWorkbook(file, sibs, opts) {
    opts = opts || {};
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
      const path = opts.dest || await this.uniquePath(this.settings.folder + '/导出', `${meta[file.path].workbook}-${stamp}`, 'xlsx');
      await this.writeBinaryTo(path, toXLSXBook(sheets));
      if (!opts.silent) new Notice(`已导出整个工作簿（${sheets.length} 个工作表）到 ${path}`, 8000);
      if (!opts.nolog) this.logOp({ type: 'export', path: file.path, name: meta[file.path].workbook, text: `导出整个工作簿（${sheets.length} 个工作表）${opts.from ? `（${opts.from}）` : ''}`, file: path });
      if (!opts.quiet && this.settings.openAfterExport && this.app.openWithDefaultApp) {
        try { await this.app.openWithDefaultApp(path); } catch (e) { /* 打不开就只保存 */ }
      }
      return { path };
    } catch (e) {
      if (opts.quiet) throw e;
      new Notice('导出失败：' + e.message, 8000);
    }
  }

  async exportTable(file, table, fmt, opts) {
    opts = opts || {};
    try {
      const d = new Date(), z = (n) => String(n).padStart(2, '0');
      const stamp = `${d.getFullYear()}${z(d.getMonth() + 1)}${z(d.getDate())}-${z(d.getHours())}${z(d.getMinutes())}`;
      const path = opts.dest || await this.uniquePath(this.settings.folder + '/导出', file.basename + '-' + stamp, fmt);
      if (fmt === 'xlsx') await this.writeBinaryTo(path, toXLSX(table));
      else if (fmt === 'csv') await this.app.vault.create(path, '﻿' + toCSV(table));
      else await this.app.vault.create(path, toMarkdown(table) + '\n');
      if (!opts.silent) new Notice('已导出到 ' + path, 8000);
      if (!opts.nolog) this.logOp({ type: 'export', path: file.path, text: `导出为 ${fmt.toUpperCase()}${opts.from ? `（${opts.from}）` : ''}`, file: path });
      if (!opts.quiet && fmt === 'xlsx' && this.settings.openAfterExport && this.app.openWithDefaultApp) {
        try { await this.app.openWithDefaultApp(path); } catch (e) { /* 打不开就只保存 */ }
      }
      return { path };
    } catch (e) {
      if (opts.quiet) throw e;
      new Notice('导出失败：' + e.message, 8000);
    }
  }

  /** 写二进制文件：已存在就覆盖 */
  async writeBinaryTo(path, buf) {
    const f = this.app.vault.getAbstractFileByPath(path);
    if (f instanceof TFile) await this.app.vault.modifyBinary(f, buf);
    else await this.app.vault.createBinary(path, buf);
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
    new Setting(containerEl).setName('表格保存文件夹').setDesc('导入、新建的表格放在这里；导出文件在其中的「导出」子文件夹，附件在「附件区」。')
      .addText((t) => t.setValue(s.folder).onChange(async (v) => { s.folder = v.trim() || DEFAULT_SETTINGS.folder; await save(); }));
    new Setting(containerEl).setName('自动处理文件夹').setDesc(`开启后，往「${s.folder}/${AUTO_DIR}/导入」放 Excel 会自动导入，往「${AUTO_DIR}/导出」放请求文件会自动导出；也可以用 obsidian://table-workbench 链接调用。用法见该文件夹里的「使用说明」。Obsidian 启动时会在后台打开工作台来处理。`)
      .addToggle((t) => t.setValue(s.autoProcess !== false).onChange(async (v) => {
        s.autoProcess = v; await save();
        if (v) { const view = this.view() || (await this.plugin.activateView(true, true)); if (view) { await view.ensureAutoDirs(); view.scanAuto(); } }
      }))
      .addButton((b) => b.setButtonText('打开使用说明').onClick(async () => {
        const view = this.view() || (await this.plugin.activateView(true, true));
        if (!view) return;
        await view.ensureAutoDirs();
        const f = this.app.vault.getAbstractFileByPath(`${view.autoDir()}/使用说明.md`);
        if (f instanceof TFile) { this.app.setting.close(); this.app.workspace.getLeaf('tab').openFile(f); }
      }));
    containerEl.createEl('h3', { text: '附件 AI 扫描（按时间分类）' });
    containerEl.createEl('div', { cls: 'setting-item-description', text: '用能看图的 AI 识别附件区图片 / PDF 上的单据日期，自动放进按时间分的文件夹，并记下单据类型、对方名称、单号。API Key 只保存在本机这个库的插件数据里，图片只发给你选的服务商。' });
    const sc = scanCfg(s);
    const saveScan = async () => { s.scan = Object.assign({}, sc); await save(); };
    new Setting(containerEl).setName('服务商').setDesc('选好后自动填 API 地址和模型，也可以自己改。')
      .addDropdown((d) => { Object.entries(SCAN_PRESETS).forEach(([k, v]) => d.addOption(k, v.label)); d.setValue(sc.provider).onChange(async (v) => { sc.provider = v; const p = SCAN_PRESETS[v]; if (p.url) sc.url = p.url; sc.model = p.model; await saveScan(); this.display(); }); });
    new Setting(containerEl).setName('API 地址').setDesc((SCAN_PRESETS[sc.provider] || {}).format === 'anthropic' ? 'Anthropic 格式，填到 /v1。' : 'OpenAI 兼容格式，填到 /v1 这一级（不用加 /chat/completions）。')
      .addText((t) => t.setPlaceholder('https://…/v1').setValue(sc.url).onChange(async (v) => { sc.url = v.trim(); await saveScan(); }));
    new Setting(containerEl).setName('API Key')
      .addText((t) => { t.inputEl.type = 'password'; t.setPlaceholder('sk-…').setValue(sc.key).onChange(async (v) => { sc.key = v.trim(); await saveScan(); }); });
    new Setting(containerEl).setName('模型').setDesc(sc.provider === 'doubao' ? '豆包要填控制台里支持图片的视觉模型 ID 或推理接入点 ID（ep-…）。' : sc.provider === 'deepseek' ? '用 deepseek-flash（V4.1 Flash，能看图）；deepseek-v4-pro 等纯文字模型看不了图片。' : '要用能看图片的模型（视觉 / VL 模型）。')
      .addText((t) => t.setPlaceholder('模型名').setValue(sc.model).onChange(async (v) => { sc.model = v.trim(); await saveScan(); }));
    new Setting(containerEl).setName('回复长度上限（tokens）').setDesc('每次识别 AI 最多能回复多少 token。带「思考」的模型会先想再答，上限太低会被截断导致识别失败；一般 4000 足够，失败多可以调到 8000。')
      .addText((t) => { t.inputEl.type = 'number'; t.setPlaceholder('4000').setValue(String(sc.maxTokens)).onChange(async (v) => { const n = parseInt(v, 10); sc.maxTokens = n >= 200 ? Math.min(n, 32000) : 4000; await saveScan(); }); });
    new Setting(containerEl).setName('测试连接').setDesc('发一句话给 AI，看 API 地址、Key、模型是否填对。')
      .addButton((b) => b.setButtonText('测试').onClick(async () => {
        if (!sc.key || !sc.url || !sc.model) { new Notice('先填好 API 地址、API Key 和模型。'); return; }
        new Notice('正在测试…');
        try { const r = await callVision(sc, '只回复两个字：可以'); new Notice('连接成功，AI 回复：' + String(r).slice(0, 30)); } catch (e) { new Notice('连接失败：' + (e.message || e), 10000); }
      }));
    new Setting(containerEl).setName('分类方式').setDesc('识别出日期后放进哪种文件夹。只整理附件区根目录和日期文件夹里的附件；你自己建的分类文件夹（例如 客户/华盛）里的不会被移走，只记录日期。')
      .addDropdown((d) => d.addOption('month', '按月：2026-11').addOption('day', '按日：2026-11 / 11-15').setValue(sc.by).onChange(async (v) => { sc.by = v; await saveScan(); }));
    new Setting(containerEl).setName('新附件自动识别').setDesc('存进附件区的图片 / PDF 自动识别日期并归类（填好 API 后生效）。关掉后在附件区点「AI 按时间整理」手动识别。')
      .addToggle((t) => t.setValue(sc.auto).onChange(async (v) => { sc.auto = v; await saveScan(); }));
    new Setting(containerEl).setName('按识别内容自动改名').setDesc('识别后把「微信图片_2026…」这类文件名改成有意义的名字，例如「2026-11-15_送货单_浙江盈亿纺织_0011047」。')
      .addToggle((t) => t.setValue(sc.rename).onChange(async (v) => { sc.rename = v; await saveScan(); }));
    new Setting(containerEl).setName('文件名格式').setDesc('可用：{日期} {类型} {对方} {单号} {说明} {原名}。识别不出的字段会自动去掉。')
      .addText((t) => t.setPlaceholder(SCAN_DEFAULTS.nameTpl).setValue(sc.nameTpl).onChange(async (v) => { sc.nameTpl = v.trim() || SCAN_DEFAULTS.nameTpl; await saveScan(); }));
    containerEl.createEl('h3', { text: '其他' });
    new Setting(containerEl).setName('导出 Excel 后自动打开').setDesc('导出 .xlsx 后用系统默认程序（Excel / WPS）打开。')
      .addToggle((t) => t.setValue(s.openAfterExport !== false).onChange(async (v) => { s.openAfterExport = v; await save(); }));
  }
}

class AITableAssistantPlugin extends Plugin {
  async onload() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    if (!this.settings.meta) this.settings.meta = {};
    await this.restoreFromBackup();
    if (!Array.isArray(this.settings.pinned)) this.settings.pinned = [];
    if (!Array.isArray(this.settings.recent)) this.settings.recent = [];
    if (!this.settings.archiveMeta || typeof this.settings.archiveMeta !== 'object') this.settings.archiveMeta = {};
    /* 旧版本用过的设置（包括 API Key）已经不用了，从插件数据里清掉 */
    if (OBSOLETE_SETTINGS.some((k) => k in this.settings)) { OBSOLETE_SETTINGS.forEach((k) => delete this.settings[k]); await this.saveSettings(); }
    if (this.settings.master && !this.settings.pinned.includes(this.settings.master)) this.settings.pinned.push(this.settings.master);
    this.registerView(VIEW_TYPE, (leaf) => new TableWorkbenchView(leaf, this));
    this.addRibbonIcon('table', '表格工作台', () => this.activateView());
    this.addCommand({ id: 'open-panel', name: '打开表格工作台', callback: () => this.activateView() });
    this.addCommand({ id: 'refresh-outputs', name: '立即更新全部原表 Excel', callback: () => this.activateView(true, true).then((v) => { if (!v) return; const seen = new Set(); v.tableFiles().forEach((f, i) => { const d = v.outPathOf(f); if (!seen.has(d)) { seen.add(d); v.scheduleOutput(f, 300 + i * 50); } }); }) });
    this.addCommand({ id: 'tour', name: '查看使用引导', callback: () => this.activateView().then((v) => v && v.startTour()) });
    this.addCommand({ id: 'quick-switch', name: '在文件管理中查找表格', callback: () => this.activateView().then((v) => v && v.quickSwitch()) });
    this.addCommand({
      id: 'use-active-file',
      name: '在文件管理中定位当前表格',
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
    /* obsidian://table-workbench?action=import|export|open */
    this.registerObsidianProtocolHandler('table-workbench', async (params) => {
      const v = await this.activateView(true);
      if (v) await v.handleProtocol(params);
    });
    /* 自动处理要在后台运行：Obsidian 启动后把工作台在后台打开（不抢焦点） */
    this.app.workspace.onLayoutReady(() => {
      if (this.settings.autoProcess === false) return;
      if (!this.app.workspace.getLeavesOfType(VIEW_TYPE).length) this.activateView(true, true);
    });
  }

  async saveSettings() {
    await this.saveData(this.settings);
    clearTimeout(this.backupTimer);
    this.backupTimer = setTimeout(() => this.writeBackup(), 1500);
  }

  backupPath(folder) { return normalizePath((folder || this.settings.folder) + '/.ata/settings-backup.json'); }

  /** 插件设置（Excel 分组、原格式对应、分类、常用）另存一份在表格文件夹里：
      换电脑、同步出错或插件被重装时，data.json 丢了也能恢复 */
  async writeBackup() {
    const s = this.settings, a = this.app.vault.adapter;
    const n = Object.keys(s.meta || {}).length;
    try {
      if (!n) {
        try { const old = JSON.parse(await a.read(this.backupPath())); if (Object.keys(old.meta || {}).length) return; } catch (e) { /* 没有旧备份 */ }
      }
      const dir = this.backupPath().replace(/\/[^/]+$/, '');
      if (!(await a.exists(dir))) await a.mkdir(dir);
      await a.write(this.backupPath(), JSON.stringify({ savedAt: Date.now(), meta: s.meta, pinned: s.pinned, recent: s.recent, archiveMeta: s.archiveMeta, onboarded: s.onboarded, lastTable: s.lastTable }));
    } catch (e) { console.warn('[表格工作台] 备份设置失败', e); }
  }

  async restoreFromBackup() {
    const s = this.settings;
    if (Object.keys(s.meta || {}).length) return;
    let b = null;
    for (const f of [s.folder, DEFAULT_FOLDER, OLD_FOLDER]) {
      try { b = JSON.parse(await this.app.vault.adapter.read(this.backupPath(f))); if (b) { s.folder = f; break; } } catch (e) { /* 试下一个 */ }
    }
    if (!b) return;
    if (!b || !Object.keys(b.meta || {}).length) return;
    s.meta = b.meta;
    if (!(s.pinned || []).length && Array.isArray(b.pinned)) s.pinned = b.pinned;
    if ((s.recent || []).length < 2 && Array.isArray(b.recent)) s.recent = [...new Set([...(s.recent || []), ...b.recent])].slice(0, 20);
    if (!Object.keys(s.archiveMeta || {}).length && b.archiveMeta) s.archiveMeta = b.archiveMeta;
    if (b.onboarded) s.onboarded = true;
    await this.saveData(s);
    this.restoredFromBackup = Object.keys(b.meta).length;
  }

  async activateView(quiet, background) {
    const ws = this.app.workspace;
    let leaf = ws.getLeavesOfType(VIEW_TYPE)[0];
    /* 旧版本放在侧边栏，迁到主编辑区 */
    if (leaf && leaf.getRoot() !== ws.rootSplit) { leaf.detach(); leaf = null; }
    if (!leaf) {
      leaf = ws.getLeaf('tab');
      await leaf.setViewState({ type: VIEW_TYPE, active: !background });
    }
    if (leaf.loadIfDeferred) await leaf.loadIfDeferred();
    if (!background) ws.revealLeaf(leaf);
    return leaf.view instanceof TableWorkbenchView ? leaf.view : null;
  }
}

AITableAssistantPlugin._test = {
  parseCSV, toCSV, gridToTable, findMdTable, toMarkdown, parseXLSX, toXLSX,
  normalizeTable, collectDropped, cleanGrid, decodeText, readWorkbook, toXLSXBook, fileToSheets, patchWorkbook, cloneTable, diffTables, hashStr,
};
module.exports = AITableAssistantPlugin;
