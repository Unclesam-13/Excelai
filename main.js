'use strict';
/* AI 表格助手 · 表格工作台 —— Obsidian 插件（无需构建，直接使用本文件） */
const {
  Plugin, ItemView, Notice, Modal, FuzzySuggestModal, PluginSettingTab,
  Setting, requestUrl, setIcon, normalizePath, Platform, TFile, MarkdownRenderer,
} = require('obsidian');

const VIEW_TYPE = 'ai-table-assistant-view';
const MAX_FILE = 200 * 1024 * 1024; // 导入 / 附件大小上限 200MB
const MAX_PDF = 30 * 1024 * 1024;   // PDF 发给 AI 的上限

const PROVIDER_DEFAULTS = {
  anthropic: { baseUrl: 'https://api.anthropic.com', model: 'claude-sonnet-5-5', label: 'Claude (Anthropic)' },
  deepseek: { baseUrl: 'https://api.deepseek.com', model: 'deepseek-flash', visionModel: 'deepseek-flash', label: 'DeepSeek（深度求索）' },
  openai: { baseUrl: 'https://api.openai.com/v1', model: 'gpt-4o', label: 'OpenAI 兼容（通义、Kimi、豆包等）' },
};
const providerOf = (settings) => (PROVIDER_DEFAULTS[settings.provider] ? settings.provider : 'anthropic');

const DEFAULT_SETTINGS = {
  provider: 'anthropic',
  apiKey: '',
  baseUrl: '',
  model: '',
  folder: 'AI表格助手',
  maxRows: 300,
  lastTable: '',
  master: '',
  pinned: [],
  recent: [],
  preview: true,
  autoArchive: true,
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

function newTrack() { return { cells: new Map(), added: new Set(), structural: false }; }
function markCell(track, r, c) {
  if (!track) return;
  if (!track.cells.has(r)) track.cells.set(r, new Set());
  track.cells.get(r).add(c);
}

/** 按关键列合并：匹配到就更新非空字段，匹配不到就新增 */
function upsertRows(t, keys, objs, track) {
  const str = (v) => (v == null ? '' : typeof v === 'object' ? JSON.stringify(v) : String(v));
  const keyCols = (keys || []).map((k) => findCol(t, k)).filter((c) => c >= 0);
  const index = new Map();
  if (keyCols.length) {
    t.rows.forEach((r, i) => {
      const k = keyCols.map((c) => keyNorm(r[c])).join('\u0001');
      if (k.replace(/\u0001/g, '') && !index.has(k)) index.set(k, i);
    });
  }
  let upd = 0, cells = 0, add = 0;
  for (const obj of objs || []) {
    if (!obj || typeof obj !== 'object' || Array.isArray(obj)) continue;
    const entries = [];
    for (const k of Object.keys(obj)) {
      let c = findCol(t, k);
      if (c < 0) { if (!String(k).trim()) continue; c = addColumn(t, k); }
      entries.push([c, str(obj[k]).trim()]);
    }
    if (!entries.some(([, v]) => v.trim() !== '')) continue;
    let hit;
    let keyVal = '';
    if (keyCols.length) {
      keyVal = keyCols.map((c) => { const e = entries.find(([cc]) => cc === c); return keyNorm(e ? e[1] : ''); }).join('\u0001');
      if (keyVal.replace(/\u0001/g, '')) hit = index.get(keyVal);
    }
    if (hit !== undefined) {
      let changed = 0;
      for (const [c, v] of entries) {
        if (v === '' || keyCols.includes(c)) continue;
        while (t.rows[hit].length < t.headers.length) t.rows[hit].push('');
        if (t.rows[hit][c] !== v) { t.rows[hit][c] = v; changed++; markCell(track, hit, c); }
      }
      if (changed) { upd++; cells += changed; }
    } else {
      const line = t.headers.map(() => '');
      for (const [c, v] of entries) line[c] = v;
      t.rows.push(line);
      const idx = t.rows.length - 1;
      if (keyVal && keyVal.replace(/\u0001/g, '')) index.set(keyVal, idx);
      if (track) track.added.add(idx);
      add++;
    }
  }
  return { upd, cells, add };
}

/** 就地修改表格，返回变更说明数组。ctx: { attachments: {名称: table}, track } */
function applyActions(t, actions, ctx) {
  ctx = ctx || {};
  const track = ctx.track;
  const log = [];
  const str = (v) => (v == null ? '' : typeof v === 'object' ? JSON.stringify(v) : String(v));
  const keyList = (k) => (Array.isArray(k) ? k : k ? [k] : []).map(String).filter(Boolean);
  for (const a of actions || []) {
    if (!a || typeof a !== 'object') continue;
    switch (a.type) {
      case 'add_rows': {
        let n = 0;
        for (const r of a.rows || []) {
          const line = t.headers.map(() => '');
          if (Array.isArray(r)) {
            r.forEach((v, i) => {
              if (i >= t.headers.length) { addColumn(t, '列' + (i + 1)); line.push(''); }
              line[i] = str(v);
            });
          } else if (r && typeof r === 'object') {
            for (const k of Object.keys(r)) {
              let c = findCol(t, k);
              if (c < 0) { c = addColumn(t, k); line.push(''); }
              line[c] = str(r[k]);
            }
          } else continue;
          while (line.length < t.headers.length) line.push('');
          t.rows.push(line);
          if (track) track.added.add(t.rows.length - 1);
          n++;
        }
        if (n) log.push(`新增 ${n} 行`);
        break;
      }
      case 'upsert_rows': {
        const keys = keyList(a.key || a.keys);
        const r = upsertRows(t, keys, a.rows, track);
        if (r.upd || r.add) log.push(`${keys.length ? `按「${keys.join(' + ')}」匹配：` : ''}更新 ${r.upd} 行（${r.cells} 处），新增 ${r.add} 行`);
        break;
      }
      case 'merge_attachment': {
        const atts = ctx.attachments || {};
        const names = Object.keys(atts);
        const name = a.attachment && atts[a.attachment] ? a.attachment
          : names.find((n) => a.attachment && (n.includes(a.attachment) || String(a.attachment).includes(n))) || (names.length === 1 ? names[0] : null);
        const src = name && atts[name];
        if (!src) { log.push(`⚠ 找不到附件表格「${a.attachment || ''}」`); break; }
        const mapping = a.mapping || {};
        const pairs = Object.keys(mapping).map((from) => [findCol(src, from), mapping[from]]).filter(([c, to]) => c >= 0 && to);
        if (!pairs.length) { log.push(`⚠ 附件「${name}」的列对应关系无效`); break; }
        const set = a.set && typeof a.set === 'object' ? a.set : {};
        const objs = src.rows.map((r) => {
          const o = {};
          for (const [c, to] of pairs) o[to] = r[c] == null ? '' : r[c];
          return o;
        }).filter((o) => Object.values(o).some((v) => String(v).trim() !== ''))
          .map((o) => Object.assign(o, set));
        const keys = keyList(a.key || a.keys);
        const r = upsertRows(t, keys, objs, track);
        log.push(`从「${name}」合并 ${objs.length} 行：更新 ${r.upd} 行（${r.cells} 处），新增 ${r.add} 行`);
        break;
      }
      case 'update_cells': {
        let n = 0;
        for (const u of a.updates || []) {
          const ri = +u.row - 1;
          const r = t.rows[ri];
          if (!r) continue;
          let c = findCol(t, u.column);
          if (c < 0 && typeof u.column === 'string' && u.column.trim()) c = addColumn(t, u.column);
          if (c < 0) continue;
          while (r.length < t.headers.length) r.push('');
          r[c] = str(u.value);
          markCell(track, ri, c);
          n++;
        }
        if (n) log.push(`修改 ${n} 个单元格`);
        break;
      }
      case 'add_column': {
        if (!a.name) break;
        let c = findCol(t, a.name);
        const isNew = c < 0;
        if (isNew) c = addColumn(t, a.name, null, a.default);
        t.rows.forEach((r, i) => {
          if (a.values && a.values[i] != null) { r[c] = str(a.values[i]); markCell(track, i, c); }
          else if (isNew && a.default != null) markCell(track, i, c);
        });
        log.push(isNew ? `新增列「${a.name}」` : `更新列「${a.name}」`);
        break;
      }
      case 'rename_column': {
        const c = findCol(t, a.from);
        if (c >= 0 && a.to) { t.headers[c] = String(a.to); log.push(`列「${a.from}」改名为「${a.to}」`); }
        break;
      }
      case 'delete_column': {
        const c = findCol(t, a.name);
        if (c >= 0) {
          t.headers.splice(c, 1);
          t.rows.forEach((r) => r.splice(c, 1));
          if (t.colSrc) t.colSrc.splice(c, 1);
          if (track) track.structural = true;
          log.push(`删除列「${a.name}」`);
        }
        break;
      }
      case 'delete_rows': {
        const idx = [...new Set((a.rows || []).map((x) => +x - 1))]
          .filter((i) => i >= 0 && i < t.rows.length).sort((x, y) => y - x);
        idx.forEach((i) => t.rows.splice(i, 1));
        if (idx.length) { if (track) track.structural = true; log.push(`删除 ${idx.length} 行`); }
        break;
      }
      case 'sort': {
        const c = findCol(t, a.column);
        if (c < 0) break;
        const dir = String(a.order || 'asc').toLowerCase().startsWith('d') ? -1 : 1;
        const num = (v) => /^-?[\d,]*\.?\d+$/.test(String(v).trim()) ? parseFloat(String(v).replace(/,/g, '')) : NaN;
        t.rows.sort((x, y) => {
          const nx = num(x[c]), ny = num(y[c]);
          return dir * (!isNaN(nx) && !isNaN(ny) ? nx - ny : String(x[c]).localeCompare(String(y[c]), 'zh-Hans-CN', { numeric: true }));
        });
        if (track) track.structural = true;
        log.push(`按「${t.headers[c]}」${dir > 0 ? '升序' : '降序'}排序`);
        break;
      }
      case 'replace_table': {
        if (!Array.isArray(a.headers)) break;
        const nt = normalizeTable({ headers: a.headers.map(str), rows: (a.rows || []).map((r) => (Array.isArray(r) ? r.map(str) : [])) });
        t.headers = nt.headers;
        t.rows = nt.rows;
        t.colSrc = null;
        if (track) track.structural = true;
        log.push(`重建表格（${t.rows.length} 行 × ${t.headers.length} 列）`);
        break;
      }
    }
  }
  return log;
}

function tableToPrompt(t, maxRows) {
  if (!t.headers.length) return '（空表，还没有任何列）';
  const shown = t.rows.slice(0, maxRows);
  const lines = [['#', ...t.headers].map(csvCell).join(',')];
  shown.forEach((r, i) => lines.push([String(i + 1), ...r].map(csvCell).join(',')));
  let s = lines.join('\n');
  if (t.rows.length > shown.length) s += `\n（共 ${t.rows.length} 行，这里只显示前 ${shown.length} 行）`;
  return s;
}

/**
 * 宽松修复 AI 返回的 JSON：字符串里的英文双引号、直接换行、多余逗号、被截断没写完的结尾。
 */
function repairJSON(src) {
  let out = '';
  const stack = [];
  let inStr = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (inStr) {
      if (ch === '\\') {
        const nx = src[i + 1];
        if (nx === undefined) break;
        if ('"\\/bfnrtu'.includes(nx)) { out += ch + nx; i++; }
        else out += '\\\\';
        continue;
      }
      if (ch === '"') {
        let j = i + 1;
        while (j < src.length && /\s/.test(src[j])) j++;
        const nx = src[j];
        if (nx === undefined || nx === ',' || nx === '}' || nx === ']' || nx === ':') { inStr = false; out += ch; }
        else out += '\\"'; // 字符串中间的英文引号
        continue;
      }
      if (ch === '\n') { out += '\\n'; continue; }
      if (ch === '\r') continue;
      if (ch === '\t') { out += '\\t'; continue; }
      out += ch;
      continue;
    }
    if (ch === '"') { inStr = true; out += ch; continue; }
    if (ch === '{' || ch === '[') { stack.push(ch === '{' ? '}' : ']'); out += ch; continue; }
    if (ch === '}' || ch === ']') {
      out = out.replace(/,\s*$/, '');
      if (stack.length && stack[stack.length - 1] === ch) stack.pop();
      out += ch;
      if (!stack.length) break; // 最外层结束，后面的说明文字不要
      continue;
    }
    out += ch;
  }
  repairJSON.truncated = inStr || stack.length > 0;
  if (inStr) out += '"';
  out = out.replace(/,\s*$/, '').replace(/:\s*$/, ': null');
  while (stack.length) { out = out.replace(/,\s*$/, ''); out += stack.pop(); }
  return out;
}

function extractJSON(text) {
  if (!text) return null;
  const s = String(text).trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  const a = s.indexOf('{');
  if (a < 0) return null;
  const b = s.lastIndexOf('}');
  if (b > a) { try { return JSON.parse(s.slice(a, b + 1)); } catch (e) { /* 试着修复 */ } }
  try {
    const r = JSON.parse(repairJSON(s.slice(a)));
    if (r && typeof r === 'object' && repairJSON.truncated) Object.defineProperty(r, '__truncated', { value: true });
    return r;
  } catch (e) { /* 还是不行 */ }
  return null;
}

/** 实在解析不了时，至少把 reply 的文字取出来 */
function salvageReply(text) {
  const m = /"reply"\s*:\s*"([\s\S]*?)(?:"\s*,\s*"actions"|"\s*}\s*$|$)/.exec(String(text || ''));
  if (!m) return null;
  return m[1].replace(/\\n/g, '\n').replace(/\\t/g, '  ').replace(/\\"/g, '"').replace(/\\\\/g, '\\');
}

const SYSTEM_PROMPT = `你是 Obsidian 里的「AI 表格助手」，服务对象是做会计 / 采购 / 对账工作的用户。每次用户发言时，你会看到当前「目标表格」（CSV 形式，第一列 # 是行号，不属于表格数据），以及用户可能附带的图片、PDF（单据、价格表、对账单、截图）或其它表格文件。你的工作是按用户要求修改目标表格，或回答关于表格的问题。

【输出格式】只输出一个 JSON 对象，不要输出任何其它文字，不要使用代码块：
{"reply":"给用户的简短中文说明","actions":[ ... ]}

【可用的 action】
- {"type":"upsert_rows","key":["关键列", ...],"rows":[{"列名":"值", ...}]}
  按关键列合并：关键列的值和已有行相同就更新该行（只更新你给出的非空字段），否则新增一行。插件会在【整张表】里匹配，即使你只看到了部分行。录入单据、更新价格表时优先用它。
- {"type":"merge_attachment","attachment":"附件文件名","mapping":{"附件列名":"目标表列名", ...},"key":["目标表关键列", ...],"set":{"目标表列名":"固定值"}}
  把附件表格的【全部行】按 mapping 换成目标表的列，再按 key 合并（规则同 upsert_rows）。附件是表格且行数较多时必须用它，不要把附件内容逐行抄进 upsert_rows。set 用来给所有行统一填值，例如来源、更新日期。不需要匹配时 key 用空数组（全部新增）。
- {"type":"add_rows","rows":[{"列名":"值", ...}]}  在末尾追加行（确定都是新记录时用）
- {"type":"update_cells","updates":[{"row":行号,"column":"列名","value":"值"}]}
- {"type":"add_column","name":"列名","values":["第1行的值","第2行的值", ...]}  新增一列（列已存在则覆盖），values 按行号顺序
- {"type":"rename_column","from":"旧列名","to":"新列名"}
- {"type":"delete_column","name":"列名"}
- {"type":"delete_rows","rows":[行号, ...]}  放在 actions 最后
- {"type":"sort","column":"列名","order":"asc 或 desc"}
- {"type":"replace_table","headers":["列1", ...],"rows":[["值", ...], ...]}  整表重建，仅在需要大幅重组或目标表为空需要建表时使用

【规则】
1. 行号从 1 开始，不含表头。所有值都用字符串。日期统一写成 YYYY-MM-DD，金额、单价、数量只写数字（不带单位和千分位逗号）。
2. 关键列要选能唯一识别一条记录的列，例如：品名+规格、物料编码、单号、发票号码。目标表里没有合适关键列时，用 add_rows。
3. 如果目标表里有「来源」「来源单据」「更新日期」「录入日期」之类的列，填上附件名称和单据日期（单据上没有日期就用今天）。
4. 图片 / PDF：仔细识别其中的文字、表格、票据，按目标表的列含义对应填入；一张单据有多条明细就输出多行。看不到的信息留空，绝不编造；看不清或不确定的值在 reply 里点名提醒用户核对。
5. 分类任务：用 add_column 新增一列（默认叫「分类」），为每一行给出类别。
6. 如果用户只是提问或要求检查，actions 用空数组，在 reply 里回答。
7. 表格被截断显示时，按行号的修改（update_cells / delete_rows / add_column）只能针对你看到的行，并在 reply 里说明。
8. reply 简短，说明做了什么、哪些需要人工核对。`;

const ROUTE_PROMPT = `你是表格归档助手。用户手上有很多张要填写的表格，现在收到了一些新资料（单据、送货单、价格表、对账单、订单截图、PDF 或表格文件），需要判断这些资料应该填进哪张表。你会看到候选表格列表（编号、名称、所属文件、列名、一行样例）和新资料的内容。

只输出一个 JSON 对象，不要输出其它文字：
{"reply":"一句话说明资料是什么","candidates":[{"id":编号,"score":0到100,"reason":"12字以内的理由"}],"new_table":null 或 {"name":"建议表名","headers":["列1","列2", ...]}}

规则：candidates 最多 5 个，按匹配程度从高到低；资料中的字段与表格列越吻合、业务越相关（同一客户、同类单据）分数越高。如果没有一张表合适（最高低于 40 分），给出 new_table 建议，列名用资料里的字段；否则 new_table 为 null。`;

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

const MASTER_PROMPT = '请把附件里的信息整理进当前表格：按关键列（如品名+规格、编码或单号）匹配已有记录，有变化的字段更新，没有的记录新增；如果表格里有「来源」「更新日期」之类的列，填上附件名和单据日期。附件是表格时用 merge_attachment。最后说明改了什么、哪些需要我核对。';

/* ====================================================================
 * 二、AI 接口
 * ==================================================================== */
async function callAI(settings, history, userText, files, system) {
  system = system || SYSTEM_PROMPT;
  files = files || [];
  const provider = providerOf(settings);
  const def = PROVIDER_DEFAULTS[provider];
  const base = (settings.baseUrl || def.baseUrl).trim().replace(/\/+$/, '');
  let model = (settings.model || def.model).trim();
  if (!settings.apiKey) throw new Error('还没有填写 API Key，请到 设置 → AI 表格助手 中填写。');
  /* DeepSeek / OpenAI 兼容接口不能直接读 PDF：把 PDF 前几页转成图片 */
  if (provider !== 'anthropic' && files.some((f) => f.kind === 'pdf')) {
    const out = [];
    for (const f of files) {
      if (f.kind !== 'pdf') { out.push(f); continue; }
      const pages = await pdfToImages(f);
      if (!pages.length) throw new Error(`当前 AI 服务不能直接读 PDF，而且「${f.name}」转图片失败。请截图后作为图片发送，或改用 Claude。`);
      out.push(...pages);
    }
    files = out;
  }
  /* DeepSeek 只有 Flash 能看图片：带图片时自动换成能看图的模型 */
  if (provider === 'deepseek' && files.length && !/flash/i.test(model)) model = def.visionModel;

  let url, headers, body;
  if (provider === 'anthropic') {
    url = base + '/v1/messages';
    headers = { 'content-type': 'application/json', 'x-api-key': settings.apiKey, 'anthropic-version': '2023-06-01' };
    const content = files.map((f) => f.kind === 'pdf'
      ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: f.data }, title: f.name }
      : { type: 'image', source: { type: 'base64', media_type: f.mediaType, data: f.data } });
    content.push({ type: 'text', text: userText });
    body = { model, max_tokens: 16000, system, messages: [...history, { role: 'user', content }] };
  } else {
    url = base + '/chat/completions';
    headers = { 'content-type': 'application/json', authorization: 'Bearer ' + settings.apiKey };
    const content = files.length
      ? [{ type: 'text', text: userText }, ...files.map((im) => ({ type: 'image_url', image_url: { url: `data:${im.mediaType};base64,${im.data}` } }))]
      : userText;
    body = { model, messages: [{ role: 'system', content: system }, ...history, { role: 'user', content }] };
    if (provider === 'deepseek') {
      body.response_format = { type: 'json_object' }; // 提示词里要求输出 JSON，DeepSeek 支持 JSON 模式
      body.max_tokens = 16000;
    }
  }

  const res = await requestUrl({ url, method: 'POST', headers, body: JSON.stringify(body), throw: false });
  let json = null;
  try { json = JSON.parse(res.text); } catch (e) { /* 非 JSON */ }
  if (res.status < 200 || res.status >= 300) {
    const msg = (json && json.error && (json.error.message || json.error)) || (res.text || '').slice(0, 300);
    throw new Error(`AI 接口返回错误（${res.status}）：${typeof msg === 'string' ? msg : JSON.stringify(msg)}`);
  }
  if (!json) throw new Error('AI 接口返回的内容无法解析。');
  if (provider === 'anthropic') return (json.content || []).filter((c) => c.type === 'text').map((c) => c.text).join('');
  if (json.choices && json.choices[0] && json.choices[0].finish_reason === 'length') console.warn('[AI表格助手] 回复被截断');
  const c = json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content;
  return Array.isArray(c) ? c.map((x) => x.text || '').join('') : c || '';
}

/* ---------- 附件读取 ---------- */
function readAsDataURL(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(new Error('读取文件失败'));
    r.readAsDataURL(file);
  });
}

/** 图片 → base64（过大时缩小并转 JPEG） */
async function imageToPayload(file) {
  let dataUrl = await readAsDataURL(file);
  let mediaType = file.type || 'image/png';
  const supported = /^image\/(jpeg|png|webp|gif)$/.test(mediaType);
  try {
    const img = await new Promise((res, rej) => {
      const i = new Image();
      i.onload = () => res(i);
      i.onerror = rej;
      i.src = dataUrl;
    });
    const scale = Math.min(1, 2000 / Math.max(img.width, img.height));
    if (scale < 1 || file.size > 3e6 || !supported) {
      const cv = document.createElement('canvas');
      cv.width = Math.max(1, Math.round(img.width * scale));
      cv.height = Math.max(1, Math.round(img.height * scale));
      const ctx = cv.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, cv.width, cv.height);
      ctx.drawImage(img, 0, 0, cv.width, cv.height);
      dataUrl = cv.toDataURL('image/jpeg', 0.9);
      mediaType = 'image/jpeg';
    }
  } catch (e) {
    if (!supported) throw new Error(`无法读取图片「${file.name}」，请换成 JPG / PNG 格式。`);
  }
  return { kind: 'image', name: file.name || '图片', mediaType, data: dataUrl.slice(dataUrl.indexOf(',') + 1), dataUrl };
}

/** 用 Obsidian 自带的 PDF.js 把 PDF 前几页渲染成图片（给不能直接读 PDF 的 AI 用） */
async function pdfToImages(pdf, maxPages) {
  maxPages = maxPages || 6;
  let lib = null;
  try { lib = await nodeObsidian().loadPdfJs(); } catch (e) { lib = window.pdfjsLib || null; }
  if (!lib) return [];
  const bytes = Uint8Array.from(atob(pdf.data), (c) => c.charCodeAt(0));
  const doc = await lib.getDocument({ data: bytes }).promise;
  const out = [];
  for (let i = 1; i <= Math.min(doc.numPages, maxPages); i++) {
    const page = await doc.getPage(i);
    const vp0 = page.getViewport({ scale: 1 });
    const scale = Math.min(2.5, 1800 / Math.max(vp0.width, vp0.height));
    const vp = page.getViewport({ scale });
    const cv = document.createElement('canvas');
    cv.width = Math.round(vp.width);
    cv.height = Math.round(vp.height);
    const ctx = cv.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, cv.width, cv.height);
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    const url = cv.toDataURL('image/jpeg', 0.88);
    out.push({ kind: 'image', name: `${pdf.name} 第${i}页`, mediaType: 'image/jpeg', data: url.slice(url.indexOf(',') + 1) });
  }
  return out;
}
const nodeObsidian = () => require('obsidian');

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

async function pdfToPayload(file) {
  if (file.size > MAX_PDF) throw new Error(`PDF「${file.name}」超过 30MB，AI 读不了这么大的 PDF，请拆分或截图后发送。`);
  const dataUrl = await readAsDataURL(file);
  return { kind: 'pdf', name: file.name, data: dataUrl.slice(dataUrl.indexOf(',') + 1) };
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

const isImageFile = (f) => /^image\//.test(f.type) || /\.(png|jpe?g|webp|gif|heic|bmp)$/i.test(f.name || '');
const isPdfFile = (f) => f.type === 'application/pdf' || /\.pdf$/i.test(f.name || '');
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

const CLASSIFY_PROMPT = `你是表格文件归类助手。用户会给你若干个表格文件的路径、列名和前几行样例。请根据内容（不要只看文件名）为每个文件归入一个简洁的类别（2~6 个字，例如：收入、费用报销、银行流水、发票、工资、固定资产、往来款、预算）。用户指定了类别体系时只能使用用户给的类别。类别之间要互斥，数量尽量少。

只输出一个 JSON 对象，不要输出其它文字：
{"reply":"一句话说明归类结果","categories":[{"path":"文件路径（原样返回）","category":"类别"}]}`;

/* ====================================================================
 * 三、弹窗
 * ==================================================================== */
class TablePickModal extends FuzzySuggestModal {
  constructor(app, files, onPick, meta) {
    super(app);
    this.files = files;
    this.onPick = onPick;
    this.meta = meta || {};
    this.setPlaceholder('选择表格（.csv 文件，或含表格的 .md 笔记）');
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

class RouteModal extends Modal {
  /** 收到新资料：推荐应该填进哪张表 */
  constructor(app, view, list) {
    super(app);
    this.view = view;
    this.list = list;
    this.scores = new Map(); // path -> { local, ai, reason }
    this.query = '';
  }
  async onOpen() {
    const v = this.view, list = this.list;
    this.modalEl.addClass('ata-route-modal');
    this.tables = list.filter((x) => TABLE_EXT.test(x.file.name));
    this.docs = list.filter((x) => isImageFile(x.file) || isPdfFile(x.file));
    const other = list.length - this.tables.length - this.docs.length;
    this.titleEl.setText('要填到哪张表？');
    const info = this.contentEl.createDiv('ata-route-info');
    info.createDiv({ text: `收到 ${list.length} 个文件：${list.slice(0, 3).map((x) => x.file.name).join('、')}${list.length > 3 ? ' 等' : ''}` });
    if (other) info.createDiv({ cls: 'ata-muted', text: `${other} 个不是表格 / 图片 / PDF，会跳过。` });

    const search = this.contentEl.createEl('input', { cls: 'ata-route-search', attr: { type: 'text', placeholder: '搜索表格名称、列名…' } });
    search.addEventListener('input', () => { this.query = search.value.trim().toLowerCase(); this.draw(); });
    this.statusEl = this.contentEl.createDiv('ata-route-status');
    this.listEl = this.contentEl.createDiv('ata-route-list');
    this.extraEl = this.contentEl.createDiv('ata-route-extra');

    const bar = this.contentEl.createDiv('ata-route-bar');
    if (this.tables.length) {
      const b = bar.createEl('button', { text: `不填现有表，导入为新表格（${this.tables.length} 个文件）` });
      b.addEventListener('click', () => { this.close(); v.importBatch(list); });
    }
    const nb = bar.createEl('button', { text: '新建一张空表再填' });
    nb.addEventListener('click', () => {
      this.close();
      new NewTableModal(this.app, async (n, c, f) => { const file = await v.createTable(n, c, f, true); if (file) await v.fillTable(file, list); }).open();
    });

    this.lib = v.lib && v.lib.length ? v.lib : await v.library();
    /* 表格文件：先在本地按列名比对 */
    this.incoming = [];
    for (const x of this.tables) {
      try {
        for (const sh of await fileToSheets(x.file)) this.incoming.push({ name: x.file.name + (sh.name ? '·' + sh.name : ''), headers: sh.table.headers, sample: sh.table.rows[0] || [] });
      } catch (e) { /* 读不了就不比 */ }
    }
    for (const t of this.lib) {
      let best = 0;
      for (const inc of this.incoming) best = Math.max(best, headerScore(inc.headers, t.headers));
      if (best) this.scores.set(t.file.path, { local: best });
    }
    this.draw();
    search.focus();

    const top = Math.max(0, ...[...this.scores.values()].map((s) => s.local || 0));
    if (this.lib.length > 1 && v.settings.apiKey && (this.docs.length || top < 0.6)) this.askAI();
    else if (!v.settings.apiKey && this.docs.length) this.statusEl.setText('填写 API Key 后，AI 可以看图片 / PDF 帮你判断该填哪张表。');
  }

  rank() {
    const recent = this.view.settings.recent || [];
    const pinned = this.view.settings.pinned || [];
    const cur = this.view.file && this.view.file.path;
    return this.lib.map((t) => {
      const s = this.scores.get(t.file.path) || {};
      const score = s.ai != null ? s.ai / 100 * 0.7 + (s.local || 0) * 0.3 : (s.local || 0);
      const ri = recent.indexOf(t.file.path);
      const bonus = (pinned.includes(t.file.path) ? 0.003 : 0) + (ri >= 0 ? 0.004 - ri * 0.0001 : 0) + (t.file.path === cur ? 0.001 : 0);
      return { t, s, score, sort: score + bonus };
    }).filter((x) => {
      if (!this.query) return true;
      const m = (this.view.settings.meta || {})[x.t.file.path] || {};
      return [x.t.file.path, m.sheet, m.workbook, ...x.t.headers].join(' ').toLowerCase().includes(this.query);
    }).sort((a, b) => b.sort - a.sort);
  }

  draw() {
    const v = this.view;
    this.listEl.empty();
    const rows = this.rank().slice(0, 40);
    if (!rows.length) this.listEl.createDiv({ cls: 'ata-muted', text: '没有匹配的表格。' });
    const meta = v.settings.meta || {};
    rows.forEach(({ t, s, score }, i) => {
      const r = this.listEl.createDiv('ata-route-row' + (i === 0 && score >= 0.4 ? ' is-best' : ''));
      const left = r.createDiv('ata-route-main');
      const name = left.createDiv('ata-route-name');
      if ((v.settings.pinned || []).includes(t.file.path)) name.createSpan({ cls: 'ata-star', text: '★' });
      const m = meta[t.file.path] || {};
      name.createSpan({ text: m.sheet ? `${m.sheet}` : t.file.basename });
      if (m.workbook) name.createSpan({ cls: 'ata-route-book', text: m.workbook });
      if (v.file && v.file.path === t.file.path) name.createSpan({ cls: 'ata-route-book', text: '当前' });
      left.createDiv({ cls: 'ata-route-sub', text: `${t.rows} 行 · ${t.headers.slice(0, 8).join('、')}` });
      if (s.reason) left.createDiv({ cls: 'ata-route-reason', text: '🤖 ' + s.reason });
      const pct = Math.round(score * 100);
      const badge = r.createDiv({ cls: 'ata-route-score' + (pct >= 60 ? ' is-high' : pct >= 35 ? ' is-mid' : ''), text: pct ? `${pct}%` : '' });
      if (!pct) badge.addClass('is-none');
      r.addEventListener('click', () => { this.close(); v.fillTable(t.file, this.list); });
    });
  }

  async askAI() {
    const v = this.view;
    this.statusEl.empty();
    this.statusEl.addClass('is-loading');
    this.statusEl.setText('🤖 AI 正在看资料，判断该填哪张表…（可以直接点选，不用等）');
    try {
      const files = [];
      for (const x of this.docs.slice(0, 4)) {
        files.push(isPdfFile(x.file) ? await pdfToPayload(x.file) : await imageToPayload(x.file));
      }
      const pool = this.rank().slice(0, 80).map((x) => x.t);
      const meta = v.settings.meta || {};
      let prompt = '【候选表格】\n';
      pool.forEach((t, i) => {
        const m = meta[t.file.path] || {};
        const sample = (t.sample && t.sample[0] ? t.sample[0] : []).slice(0, 10).map((c) => String(c).slice(0, 16)).join(' | ');
        prompt += `${i + 1}. ${m.sheet || t.file.basename}${m.workbook ? `（文件：${m.workbook}）` : ''}${m.title ? `（标题：${m.title.slice(0, 40)}）` : ''}\n   列：${t.headers.slice(0, 25).join('、')}\n   样例：${sample}\n`;
      });
      prompt += '\n【新资料】\n';
      this.incoming.slice(0, 6).forEach((inc) => { prompt += `表格「${inc.name}」列：${inc.headers.slice(0, 25).join('、')}；样例：${inc.sample.slice(0, 10).join(' | ')}\n`; });
      files.forEach((f) => { prompt += `${f.kind === 'pdf' ? 'PDF' : '图片'}「${f.name}」（见上方）\n`; });
      const raw = await callAI(v.settings, [], prompt, files, ROUTE_PROMPT);
      const res = extractJSON(raw);
      if (!res || !Array.isArray(res.candidates)) throw new Error('AI 返回格式不对');
      for (const c of res.candidates) {
        const t = pool[+c.id - 1];
        if (!t) continue;
        const s = this.scores.get(t.file.path) || {};
        this.scores.set(t.file.path, Object.assign(s, { ai: Math.max(0, Math.min(100, +c.score || 0)), reason: c.reason || '' }));
      }
      this.statusEl.removeClass('is-loading');
      this.statusEl.setText('🤖 ' + (res.reply || 'AI 已给出推荐'));
      this.extraEl.empty();
      if (res.new_table && res.new_table.name && Array.isArray(res.new_table.headers)) {
        const nt = res.new_table;
        const b = this.extraEl.createEl('button', { cls: 'mod-cta', text: `没有合适的表？按 AI 建议新建「${nt.name}」（${nt.headers.slice(0, 6).join('、')}${nt.headers.length > 6 ? '…' : ''}）` });
        b.addEventListener('click', async () => {
          this.close();
          const file = await v.createTable(nt.name, nt.headers.map(String), 'csv', true);
          if (file) await v.fillTable(file, this.list);
        });
      }
      this.draw();
    } catch (e) {
      this.statusEl.removeClass('is-loading');
      this.statusEl.setText('AI 判断失败：' + (e.message || e) + '（可以直接手动点选）');
    }
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
    btn('送去提取', () => { this.close(); v.archiveExtract([tf], 'route'); }, true);
    if (v.file) btn(`加到「${(v.settings.meta[v.file.path] || {}).sheet || v.file.basename}」`, () => { this.close(); v.archiveExtract([tf], 'current'); });
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
  constructor(app, title, value, onSubmit) { super(app); this.t = title; this.v = value || ''; this.onSubmit = onSubmit; }
  onOpen() {
    this.titleEl.setText(this.t);
    const ta = this.contentEl.createEl('textarea', { cls: 'ata-path-input', attr: { rows: '3' } });
    ta.value = this.v;
    setTimeout(() => ta.focus(), 30);
    new Setting(this.contentEl).addButton((b) => b.setButtonText('保存').setCta().onClick(() => { this.close(); this.onSubmit(ta.value.trim()); }));
  }
  onClose() { this.contentEl.empty(); }
}

/** 工作区里「从附件区选」 */
class ArchivePickModal extends Modal {
  constructor(app, view) { super(app); this.view = view; this.picked = new Set(); this.q = ''; }
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
      await v.archiveExtract(files, 'current');
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

class SheetPickModal extends Modal {
  constructor(app, fileName, sheets, onSubmit) { super(app); this.fileName = fileName; this.sheets = sheets; this.onSubmit = onSubmit; }
  onOpen() {
    this.titleEl.setText(`「${this.fileName}」有 ${this.sheets.length} 个工作表`);
    this.contentEl.createEl('p', { cls: 'setting-item-description', text: '选择要作为附件发给 AI 的工作表（选得越少越省钱，也越准确）：' });
    const picked = new Set();
    const list = this.contentEl.createDiv('ata-hr-list');
    this.sheets.forEach((sh, i) => {
      const row = list.createEl('label', { cls: 'ata-hr-row' });
      const cb = row.createEl('input', { type: 'checkbox' });
      row.createSpan({ cls: 'ata-hr-num', text: sh.name });
      row.createSpan({ cls: 'ata-hr-text', text: `${sh.table.rows.length} 行 · ${sh.table.headers.slice(0, 6).join('、')}` });
      cb.addEventListener('change', () => { if (cb.checked) picked.add(i); else picked.delete(i); });
    });
    new Setting(this.contentEl)
      .addButton((b) => b.setButtonText('全选').onClick(() => { list.querySelectorAll('input').forEach((x, i) => { x.checked = true; picked.add(i); }); }))
      .addButton((b) => b.setButtonText('添加所选').setCta().onClick(() => {
        if (!picked.size) { new Notice('请至少选一个工作表。'); return; }
        this.onSubmit([...picked].sort((a, b) => a - b).map((i) => this.sheets[i]));
        this.close();
      }));
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

class ClassifyModal extends Modal {
  constructor(app, count, uncategorized, onSubmit) {
    super(app);
    this.count = count;
    this.uncategorized = uncategorized;
    this.onSubmit = onSubmit;
  }
  onOpen() {
    this.titleEl.setText('AI 归类表格文件');
    let scope = this.uncategorized ? 'uncat' : 'all', cats = '';
    this.contentEl.createEl('p', { cls: 'setting-item-description', text: 'AI 会读取每个表格的列名和前几行，给文件打上类别标签（不会移动或修改文件）。' });
    new Setting(this.contentEl).setName('范围')
      .addDropdown((d) => d.addOption('uncat', `仅未分类（${this.uncategorized} 个）`).addOption('all', `全部（${this.count} 个）`)
        .setValue(scope).onChange((v) => (scope = v)));
    new Setting(this.contentEl).setName('指定类别（可选）').setDesc('用逗号分隔；留空则由 AI 自行归纳。')
      .addText((t) => t.setPlaceholder('收入, 费用, 银行流水, 发票').onChange((v) => (cats = v)));
    new Setting(this.contentEl).addButton((b) => b.setButtonText('开始归类').setCta().onClick(() => {
      this.close();
      this.onSubmit(scope, cats.split(/[,，、]/).map((s) => s.trim()).filter(Boolean));
    }));
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
    this.histories = new Map();
    this.chatLogs = new Map();
    this.attachments = [];
    this.undoMap = new Map();   // 每张表各自的撤销记录
    this.attachMap = new Map(); // 每张表各自的待发附件
    this.busy = false;
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
      if (el.closest('.ata-chatpanel')) return 'chat';
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
      this.root.toggleClass('ata-over-chat', z === 'chat');
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
      if (editable && editable.tagName === 'TEXTAREA' && editable.closest('.ata-chatpanel')) return; // 对话框自己处理
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
    this.attachInput = this.makeInput({ accept: 'image/*,.pdf,.csv,.tsv,.xlsx,.xlsm,.md', multiple: 'true' }, (files) => {
      if (this.pendingRoute) { this.pendingRoute = false; this.receive(files.map((f) => ({ file: f, rel: f.name }))); }
      else this.addAttachments(files);
    });

    this.registerEvent(this.app.vault.on('modify', (f) => this.onVaultChange(f)));
    this.registerEvent(this.app.vault.on('create', (f) => this.onVaultChange(f)));
    this.registerEvent(this.app.vault.on('delete', (f) => {
      if (this.file && f.path === this.file.path) { this.file = null; this.table = { headers: [], rows: [] }; }
      this.onVaultChange(f);
    }));
    this.registerEvent(this.app.vault.on('rename', () => this.scheduleRender()));

    const last = this.settings.lastTable && this.app.vault.getAbstractFileByPath(this.settings.lastTable);
    if (last instanceof TFile) await this.loadFile(last, { silent: true, switchTab: false });
    await this.render();
    this.register(() => this.endTour(false));
    if (!this.settings.onboarded) setTimeout(() => this.startTour(), 600);
  }

  /* ================= 新手引导 ================= */
  tourSteps() {
    const toTab = (t) => async () => { if (this.tab !== t) { this.tab = t; await this.render(); } };
    return [
      { kind: 'welcome' },
      { sel: '.ata-tabs', title: '四个页面', before: toTab('home'),
        text: '「首页」看最近做了什么；「表格库」放所有表格（多工作表的 Excel 会归在一起）；「工作区」是打开一张表和 AI 一起填；「附件区」存收到的单据、截图、PDF。' },
      { sel: '.ata-top-actions', title: '第一步：导入你的表格', before: toTab('home'),
        text: '点「导入文件」或「导入文件夹」，把平时要填的 Excel / CSV 导进来，一次可以导很多个。从微信里拖不进来时，复制文件后在这里按 Ctrl+V，或点「按路径」粘贴文件路径。' },
      { sel: '.ata-homestrip .ata-inbox', title: '收到单据就丢这里', before: toTab('home'),
        text: '订单、送货单、价格表的截图或 PDF 拖到这里（或点一下选择），插件会推荐该填哪张表，选好后 AI 自动识别内容填进去。' },
      { sel: '.ata-qstrip', title: '常用表一键切换', before: toTab('home'), optional: true,
        text: '最近用过和设为常用的表会排在这里，点一下就切过去。也可以把单据直接拖到某个表名上，直接填那张表。' },
      { sel: '.ata-oplog', title: '操作记录', before: toTab('home'),
        text: 'AI 改了什么、你手动改了什么、导入导出和删除，都会按时间记在这里。点「查看明细」能看到每个格子原来是什么、改成了什么。' },
      { sel: '.ata-chatpanel', title: '和 AI 一起填表', before: async () => { if (this.file) await toTab('workspace')(); }, needs: () => !!this.file,
        alt: { sel: '.ata-tabs', text: '打开任意一张表后，「工作区」右边是 AI 对话框：用中文说要改什么，或者点回形针加图片 / PDF，AI 会给出修改。打开一张表后可以在设置里重新看这段引导。' },
        text: '在这里用中文告诉 AI 要做什么，比如「把这张送货单加进去」「把单价都改成含税价」。点回形针可以加图片、PDF、表格。AI 改完会先标出来（黄色是修改，绿色是新增），确认后点「应用」，不满意可以「撤销」。' },
      { sel: '.ata-export-btn', title: '导出成 Excel', before: async () => { if (this.file) await toTab('workspace')(); }, needs: () => !!this.file,
        text: '改好后点「导出」。从 Excel 导入的表会按原文件的格式导出：颜色、公式、合并单元格、图片都保留，多个工作表一起导出成一个文件。' },
      { sel: '.ata-archive-drop', title: '附件区', before: toTab('archive'),
        text: '收到的文件可以先存在这里，图片和 PDF 有预览，图片能一键复制。要用时点卡片上的「送去提取」，交给 AI 填到表里。发给 AI 的附件也会自动存一份。' },
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
    if (h > vh * 0.5 && left - cw - 16 > 8) { // 很高的目标（比如对话框）：卡片放在它左边
      card.style.left = (left - cw - 16) + 'px';
      card.style.top = Math.max(8, Math.min(top + 40, vh - chh - 8)) + 'px';
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
    const s = this.settings;
    card.addClass('is-wide');
    card.createDiv({ cls: 'ata-tour-title big', text: '欢迎使用 AI 表格助手 👋' });
    card.createDiv({ cls: 'ata-tour-text', text: '收到订单、送货单、价格表时，不用再自己翻表格一格格填：拖进来，AI 识别内容、找到对应的表、填好，改完按原格式导出 Excel。' });
    const flow = card.createDiv('ata-tour-flow');
    [['folder-down', '导入表格'], ['inbox', '丢进单据'], ['sparkles', 'AI 填表'], ['check', '确认修改'], ['download', '导出 Excel']].forEach(([ic, tx], i) => {
      if (i) flow.createSpan({ cls: 'ata-tour-arrow', text: '→' });
      const c = flow.createDiv('ata-tour-flowitem');
      setIcon(c.createDiv('ata-tour-flowicon'), ic);
      c.createSpan({ text: tx });
    });
    const box = card.createDiv('ata-tour-setup');
    box.createDiv({ cls: 'ata-tour-sub', text: '先连上 AI（只需要一次）' });
    const row1 = box.createDiv('ata-tour-row');
    row1.createSpan({ text: 'AI 服务' });
    const sel = row1.createEl('select', { cls: 'dropdown' });
    Object.keys(PROVIDER_DEFAULTS).forEach((k) => sel.createEl('option', { attr: { value: k }, text: PROVIDER_DEFAULTS[k].label }));
    sel.value = providerOf(s);
    const row2 = box.createDiv('ata-tour-row');
    row2.createSpan({ text: 'API Key' });
    const key = row2.createEl('input', { attr: { type: 'password', placeholder: s.apiKey ? '已填写（不改就留空）' : '粘贴你的 API Key' } });
    const hint = box.createDiv({ cls: 'ata-tour-hint' });
    const setHint = () => hint.setText(sel.value === 'deepseek' ? 'DeepSeek 在 platform.deepseek.com 注册后创建 API Key，价格便宜，能识别图片。'
      : sel.value === 'anthropic' ? 'Claude 在 console.anthropic.com 创建 API Key，识别图片和 PDF 效果最好。'
        : '通义千问、Kimi、豆包等兼容 OpenAI 接口的服务，接口地址和模型可以之后在设置里填。');
    setHint();
    sel.addEventListener('change', setHint);
    if (s.apiKey) box.createDiv({ cls: 'ata-tour-ok', text: `✓ 已经配置好 ${PROVIDER_DEFAULTS[providerOf(s)].label}，可以直接开始` });
    box.createDiv({ cls: 'ata-tour-hint', text: 'Key 只保存在你自己的库里。现在不填也可以，之后在「设置」里填。' });
    const saveAi = async () => {
      if (sel.value !== providerOf(s)) { s.provider = sel.value; s.baseUrl = ''; s.model = ''; }
      if (key.value.trim()) s.apiKey = key.value.trim();
      await this.plugin.saveSettings();
    };
    const bar = card.createDiv('ata-tour-bar');
    const skip = bar.createEl('button', { cls: 'ata-tour-skip', text: '跳过，我自己摸索' });
    skip.addEventListener('click', async () => { await saveAi(); this.endTour(true); });
    const go = bar.createDiv('ata-tour-btns').createEl('button', { cls: 'mod-cta', text: '开始 1 分钟导览 →' });
    go.addEventListener('click', async () => { await saveAi(); this.tourGo(1); });
  }

  tourDone(card) {
    card.addClass('is-wide');
    card.createDiv({ cls: 'ata-tour-title big', text: '可以开始了 🎉' });
    card.createDiv({ cls: 'ata-tour-text', text: '建议先导入一张你平时要填的表，然后把一张单据截图拖进来试试。也可以先用示例表练练手：' });
    const ul = card.createEl('ul', { cls: 'ata-tour-list' });
    ['示例表里已经有两条送货记录', '对话框里填好了一句示例指令，直接点发送', 'AI 会新增一行并标成绿色，点「应用」就写进表里'].forEach((x) => ul.createEl('li', { text: x }));
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
      const file = await this.app.vault.create(path, toCSV(t));
      this.logOp({ type: 'create', path: file.path, text: '新建示例表（新手引导）' });
      const ask = `帮我加一条：${day(0)} 华盛布行 又送了棉布 280克 200米，单价 8.5，金额算一下`;
      this.pendingPrompt = ask;
      await this.openTable(file);
      setTimeout(() => { /* 新建文件会触发一次重绘，等它画完再把示例指令填进去 */
        if (this.inputEl && this.inputEl.isConnected && !this.inputEl.value.trim()) { this.inputEl.value = ask; this.inputEl.focus(); }
        this.pendingPrompt = null;
      }, 600);
      new Notice('示例表已经建好，直接点发送试试。');
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

  /** 粘贴 / 按路径收到的文件：只有一种明显去向时直接处理，否则让用户选 */
  async receive(list) {
    list = (list || []).filter(Boolean);
    if (!list.length) { new Notice('没有读到文件。'); return; }
    if (list.some((x) => x.rel.includes('/'))) { await this.importBatch(list); return; }
    if (!list.some((x) => TABLE_EXT.test(x.file.name) || isImageFile(x.file) || isPdfFile(x.file))) { new Notice('没有可以处理的表格、图片或 PDF。'); return; }
    new RouteModal(this.app, this, list).open();
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
    const folder = normalizePath(this.settings.folder) + '/';
    const meta = this.settings.meta || {};
    const arch = this.archiveDir() + '/';
    return this.app.vault.getFiles().filter((f) => {
      if (f.path.startsWith(arch)) return false;
      if (f.extension === 'csv') return !f.path.startsWith(folder + '导出/');
      if (f.extension === 'md') {
        if (f.path.startsWith(folder + '导出/')) return false;
        if (!(f.path.startsWith(folder) || meta[f.path])) return false;
        const cache = this.app.metadataCache.getFileCache(f);
        return !cache || !cache.sections || cache.sections.some((s) => s.type === 'table');
      }
      return false;
    });
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
    this.btn(right, 'AI 归类', 'sparkles', () => this.classifyDialog());
    this.btn(right, '新建表格', 'plus', () => new NewTableModal(this.app, (n, c, f) => this.createTable(n, c, f)).open(), 'mod-cta');

    /* 两个小入口：收到单据 / 导入表格 */
    const strip = body.createDiv('ata-homestrip');
    const ib = strip.createDiv('ata-inbox ata-mini');
    this.ico(ib, 'inbox', 'ata-dropzone-icon');
    const t1 = ib.createDiv('ata-dropzone-text');
    t1.createEl('b', { text: '收到单据 / 截图 / 价格表' });
    t1.createSpan({ text: '拖到这里或点击，自动找到要填的表' });
    ib.addEventListener('click', () => { this.pendingRoute = true; this.attachInput.click(); });
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

  masterFile() { return null; }

  renderInbox(body) {
    const box = body.createDiv('ata-inbox');
    this.ico(box, 'inbox', 'ata-dropzone-icon');
    const txt = box.createDiv('ata-dropzone-text');
    txt.createEl('b', { text: '拿到新的单据 / 送货单 / 价格表 / 截图？拖到这里 → 自动找到要填的表' });
    txt.createSpan({ text: '支持图片、PDF、Excel、CSV，可一次拖多个。表格文件按列名比对，图片和 PDF 由 AI 看一眼判断；推荐结果点一下就跳过去，附件和指令都带好。已经知道填哪张表，就直接拖到下面那张表的卡片上。' });
    box.addEventListener('click', () => { this.pendingRoute = true; this.attachInput.click(); });
  }

  /** 常用 + 最近：可点击切换、可拖文件进去 */
  renderQuickTables(body, lib) {
    const files = this.recentFiles(12);
    const sec = body.createDiv('ata-section');
    const h = sec.createDiv('ata-section-head');
    const t = h.createDiv('ata-section-title');
    this.ico(t, 'layers');
    t.createSpan({ text: '常用和最近的表' });
    h.createDiv({ cls: 'ata-muted', text: '点击切换；把单据拖到卡片上就直接填这张表；☆ 设为常用' });
    if (!files.length) {
      sec.createDiv({ cls: 'ata-empty-sm', text: '还没有打开过表格。去「表格库」点开一张，它就会出现在这里。' });
      return;
    }
    const grid = sec.createDiv('ata-qgrid');
    const byPath = new Map(lib.map((x) => [x.file.path, x]));
    const meta = this.settings.meta || {};
    files.forEach((f) => {
      const st = byPath.get(f.path);
      const m = meta[f.path] || {};
      const c = grid.createDiv('ata-qcard' + (this.file && this.file.path === f.path ? ' is-current' : ''));
      c.setAttribute('data-drop-table', f.path);
      const top = c.createDiv('ata-qcard-top');
      const star = top.createEl('button', { cls: 'ata-qstar' + (this.isPinned(f.path) ? ' is-on' : ''), text: this.isPinned(f.path) ? '★' : '☆', attr: { 'aria-label': this.isPinned(f.path) ? '取消常用' : '设为常用' } });
      star.addEventListener('click', (e) => { e.stopPropagation(); this.togglePin(f); });
      top.createDiv({ cls: 'ata-qcard-title', text: m.sheet || f.basename });
      c.createDiv({ cls: 'ata-qcard-sub', text: m.workbook ? `📗 ${m.workbook}` : (f.parent && f.parent.path !== '/' ? f.parent.path : '') });
      c.createDiv({ cls: 'ata-qcard-cols', text: st ? `${st.rows} 行 · ${st.headers.slice(0, 5).join('、')}` : '' });
      c.createDiv({ cls: 'ata-qcard-drop', text: '松开 → 填进这张表' });
      c.addEventListener('click', () => this.openTable(f));
    });
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
      const pending = this.attachMap.get(f.path);
      if (pending && pending.length) b.createSpan({ cls: 'ata-tab-count', text: String(pending.length) });
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
    this.btn(right, 'AI 归类', 'sparkles', () => this.classifyDialog());
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
    this.previewEl = tablePane.createDiv('ata-tablewrap');
    this.renderSheetTabs(tablePane);
    const foot = tablePane.createDiv('ata-tablefoot');
    this.btn(foot, '新增行', 'plus', () => this.addRow(), 'ata-btn-sm');
    this.btn(foot, '新增列', 'columns-3', () => this.addCol(), 'ata-btn-sm');
    this.btn(foot, '设表头行', 'heading', () => this.pickHeaderRow(), 'ata-btn-sm').title = '表头不对时（上面有标题行），把某一行设为表头';
    this.statEl = foot.createDiv({ cls: 'ata-muted ata-spacer-left' });
    this.renderTable();

    /* 右：AI 对话 */
    const chat = ws.createDiv('ata-chatpanel');
    const ch = chat.createDiv('ata-chat-head');
    this.ico(ch, 'bot');
    ch.createSpan({ text: 'AI 助手' });
    ch.createSpan({ cls: 'ata-muted', text: (this.settings.model || PROVIDER_DEFAULTS[providerOf(this.settings)].model) });
    this.iconBtn(ch, 'eraser', '清空对话', () => {
      this.histories.delete(this.file.path);
      this.chatLogs.delete(this.file.path);
      this.renderChatLog();
    }).addClass('ata-ml-auto');

    this.chatEl = chat.createDiv('ata-chat');
    this.renderChatLog();

    const chips = chat.createDiv('ata-chips ata-prompts');
    const p = (label, text) => chips.createEl('button', { text: label, cls: 'ata-catchip' })
      .addEventListener('click', () => { this.inputEl.value = text; this.inputEl.focus(); });
    p('合并附件到本表', MASTER_PROMPT);
    p('识别单据录入', '请识别图片 / PDF 里的单据内容，按本表的列整理后录入（按关键列匹配：已有的更新，没有的新增）。');
    p('按内容分类', '请根据每一行的内容进行分类，新增一列「分类」，并把同类排在一起。');
    p('检查数据', '请检查表格里有没有明显的错误、重复行、空缺或格式不统一的地方，先列出来，不要改表。');

    this.attachEl = chat.createDiv('ata-attach');
    this.renderAttachments();

    const inputRow = chat.createDiv('ata-input');
    this.iconBtn(inputRow, 'paperclip', '添加图片、PDF 或表格附件', () => { this.pendingRoute = false; this.attachInput.click(); });
    this.iconBtn(inputRow, 'archive', '从附件区选', () => new ArchivePickModal(this.app, this).open());
    this.inputEl = inputRow.createEl('textarea', { attr: { rows: '2', placeholder: '输入要求；图片 / PDF / 表格可粘贴或拖到这里…' } });
    if (this.pendingPrompt) { this.inputEl.value = this.pendingPrompt; this.pendingPrompt = null; setTimeout(() => this.inputEl.focus(), 50); }
    this.inputEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey && !e.isComposing && !Platform.isMobile) { e.preventDefault(); this.send(); }
    });
    this.inputEl.addEventListener('paste', async (e) => {
      const files = Array.from((e.clipboardData && e.clipboardData.files) || []);
      if (files.length) { e.preventDefault(); await this.addAttachments(files); }
    });
    this.sendBtn = inputRow.createEl('button', { cls: 'ata-send', attr: { 'aria-label': '发送（Enter）' } });
    setIcon(this.sendBtn, 'send');
    this.sendBtn.disabled = this.busy;
    this.sendBtn.addEventListener('click', () => this.send());
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
    } else list.push(e);
    if (list.length > 3000) list.splice(0, list.length - 3000);
    this.saveOpLog();
    if (this.tab === 'home' && !this.importRun) this.scheduleRender();
  }

  /** 把 AI 的修改整理成记录 */
  logAI(path, before, draft, track, changes, ask, atts) {
    const cells = [];
    for (const [r, cs] of track.cells) {
      if (track.added.has(r)) continue;
      for (const c of cs) {
        if (cells.length >= 200) break;
        cells.push({ r: r + 1, h: String(draft.headers[c]), o: before.rows[r] ? cellStr(before.rows[r][c]) : '', n: cellStr(draft.rows[r] && draft.rows[r][c]) });
      }
    }
    const added = [...track.added].sort((a, b) => a - b);
    const parts = [];
    if (added.length) parts.push(`新增 ${added.length} 行`);
    if (cells.length) parts.push(`修改 ${cells.length} 个单元格`);
    if (track.structural && changes && changes.length) parts.push(...changes.filter((x) => !/新增|添加|修改了/.test(x)).slice(0, 3));
    this.logOp({
      type: 'ai', path, text: parts.join('，') || (changes || []).join('；') || 'AI 修改',
      ask: ask ? String(ask).slice(0, 300) : '', atts: atts || [], cells,
      hdr: draft.headers.map(String), added: added.slice(0, 30).map((r) => (draft.rows[r] || []).map(cellStr)), addedN: added.length,
      notes: (changes || []).slice(0, 8),
    });
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
    const KINDS = [['全部', null], ['AI 修改', ['ai']], ['手动修改', ['edit', 'row', 'col', 'undo']], ['导入导出', ['import', 'export', 'create']], ['附件', ['archive']], ['删除', ['delete']]];
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
      box.createDiv({ cls: 'ata-empty-sm', text: all.length ? '没有符合条件的记录。' : '还没有操作记录。之后导入、AI 填表、手动修改、导出、删除都会记在这里，可以随时回来查改了什么。' });
      return;
    }
    const TYPE = { ai: ['AI', 'sparkles'], edit: ['手动', 'pencil'], row: ['行', 'rows-3'], col: ['列', 'columns-3'], undo: ['撤销', 'undo-2'], import: ['导入', 'folder-down'], export: ['导出', 'download'], create: ['新建', 'file-plus'], delete: ['删除', 'trash-2'], archive: ['附件', 'archive'] };
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
      if (e.ask) main.createDiv({ cls: 'ata-tl-ask', text: '“' + e.ask + '”' + (e.atts && e.atts.length ? `  📎 ${e.atts.join('、')}` : '') });
      else if (e.atts && e.atts.length) main.createDiv({ cls: 'ata-tl-ask', text: '📎 ' + e.atts.join('、') });
      if (e.file) {
        const ef = this.app.vault.getAbstractFileByPath(e.file);
        if (ef instanceof TFile) {
          const a = main.createEl('button', { cls: 'ata-tl-link', text: '打开 ' + ef.name });
          a.addEventListener('click', () => (this.app.openWithDefaultApp ? this.app.openWithDefaultApp(ef.path) : this.app.workspace.getLeaf('tab').openFile(ef)));
        }
      }
      const nCells = (e.cells || []).length, nAdd = (e.added || []).length, nRow = (e.rows || []).length;
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
      el.createDiv({ cls: 'ata-tl-sub', text: `修改的单元格（${e.cells.length}）` });
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
      if (e.cells.length > 100) el.createDiv({ cls: 'ata-muted', text: `……共 ${e.cells.length} 个` });
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
    if (e.rows && e.rows.length) rowsTable(`删除的行（${e.rows.length}）`, e.hdr || [], e.rows, e.rowsN || 0, 'is-removed');
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
      this.previewEl.createDiv({ cls: 'ata-empty-sm ata-pad', text: '这是一个空表。可以直接告诉 AI 需要哪些列，或者发图片让它自动建表；也可以点下方「新增列」。' });
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

    const buildRow = (ri) => {
      const r = t.rows[ri];
      const tr = document.createElement('tr');
      tr.dataset.r = ri;
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

  /* ---------- 对话 ---------- */
  chatLog() {
    const k = this.file ? this.file.path : '';
    if (!this.chatLogs.has(k)) this.chatLogs.set(k, []);
    return this.chatLogs.get(k);
  }

  renderChatLog() {
    if (!this.chatEl) return;
    this.chatEl.empty();
    const log = this.chatLog();
    if (!log.length) {
      const hello = this.chatEl.createDiv('ata-msg ata-assistant');
      hello.createDiv({ cls: 'ata-text', text: '我可以帮你：\n• 给每一行分类、排序、增删改\n• 识别图片（票据、截图、名片…）并录入\n• 把附件表格的信息合并进来\n• 检查数据问题\n\n修改会直接保存到文件，可随时撤销。' });
    }
    log.forEach((m) => this.drawMsg(m));
  }

  drawMsg(m, replaceEl) {
    const el = createDiv('ata-msg ata-' + m.role);
    if (replaceEl) replaceEl.replaceWith(el); else this.chatEl.appendChild(el);
    const textEl = el.createDiv({ cls: 'ata-text' });
    if ((m.role === 'assistant' || m.role === 'error') && MarkdownRenderer) {
      textEl.addClass('ata-md');
      const md = String(m.text || '').replace(/^\s*[•·]\s?/gm, '- ');
      try {
        const r = MarkdownRenderer.render ? MarkdownRenderer.render(this.app, md, textEl, (this.file && this.file.path) || '', this)
          : MarkdownRenderer.renderMarkdown(md, textEl, (this.file && this.file.path) || '', this);
        if (r && r.catch) r.catch(() => textEl.setText(m.text));
      } catch (e) { textEl.setText(m.text); }
    } else textEl.setText(m.text);
    if (m.atts && m.atts.length) {
      const strip = el.createDiv('ata-attach');
      m.atts.forEach((a) => {
        const item = strip.createDiv('ata-att');
        if (a.dataUrl) item.createEl('img', { attr: { src: a.dataUrl } });
        else item.createSpan({ text: (a.kind === 'pdf' ? '📄 ' : '▦ ') + a.name });
      });
    }
    if (m.changes && m.changes.length) {
      const ul = el.createEl('ul', { cls: 'ata-changes' });
      m.changes.forEach((c) => ul.createEl('li', { text: c }));
    }
    if (m.preview) this.renderPreview(el, m);
    if (!replaceEl) this.chatEl.scrollTop = this.chatEl.scrollHeight;
    return el;
  }

  renderPreview(el, m) {
    el.addClass('ata-msg-wide');
    const p = m.preview;
    const box = el.createDiv('ata-pv');
    if (m.done) {
      box.createDiv({ cls: 'ata-pv-done', text: m.done === 'applied' ? '✓ 已应用到表格' : m.done === 'stale' ? '表格在这期间被改过，这次修改已作废，请重新发送' : '已放弃这次修改' });
      return;
    }
    const { draft, track, before } = p;
    const rows = [...new Set([...track.cells.keys(), ...track.added])].sort((a, b) => a - b);
    if (track.structural) {
      box.createDiv({ cls: 'ata-muted', text: '这次修改包含排序 / 删除 / 重建，应用后可在左侧表格查看，不满意可以撤销。' });
    } else if (rows.length) {
      const colSet = new Set([0, 1].filter((c) => c < draft.headers.length));
      track.cells.forEach((set) => set.forEach((c) => colSet.add(c)));
      track.added.forEach((r) => draft.rows[r].forEach((v, c) => { if (!isBlank(v)) colSet.add(c); }));
      const cols = [...colSet].sort((a, b) => a - b).slice(0, 10);
      const wrap = box.createDiv('ata-pv-wrap');
      const tb = wrap.createEl('table', { cls: 'ata-pv-table' });
      const hr = tb.createEl('tr');
      hr.createEl('th', { text: '' });
      cols.forEach((c) => hr.createEl('th', { text: draft.headers[c] }));
      rows.slice(0, 40).forEach((r) => {
        const isNew = track.added.has(r);
        const tr = tb.createEl('tr', { cls: isNew ? 'is-added' : '' });
        tr.createEl('td', { cls: 'ata-pv-tag', text: isNew ? '新增' : `第 ${r + 1} 行` });
        const changed = track.cells.get(r) || new Set();
        cols.forEach((c) => {
          const v = cellStr(draft.rows[r][c]);
          const td = tr.createEl('td', { text: v });
          if (!isNew && changed.has(c)) {
            td.addClass('is-changed');
            const old = before.rows[r] ? cellStr(before.rows[r][c]) : '';
            td.title = `原来：${old || '（空）'}`;
            if (old) td.createDiv({ cls: 'ata-pv-old', text: old });
          }
        });
      });
      if (rows.length > 40) box.createDiv({ cls: 'ata-muted', text: `……还有 ${rows.length - 40} 行，应用后在表格里查看。` });
      box.createDiv({ cls: 'ata-muted', text: '黄色 = 修改（灰色小字是原值），绿色 = 新增。' });
    }
    const bar = box.createDiv('ata-pv-bar');
    this.btn(bar, '应用修改', 'check', () => this.applyPreview(m, el), 'mod-cta ata-btn-sm');
    this.btn(bar, '放弃', 'x', () => { m.done = 'discarded'; m.preview = Object.assign({}, m.preview, { draft: null }); this.drawMsg(m, el); }, 'ata-btn-sm');
  }

  async applyPreview(m, el) {
    const p = m.preview;
    if (!this.file || this.file.path !== p.path || this.tableVersion !== p.version) {
      m.done = 'stale';
      this.drawMsg(m, el);
      return;
    }
    this.pushUndo();
    this.logAI(p.path, p.before, p.draft, p.track, m.changes, p.ask, p.atts);
    this.table = p.draft;
    await this.saveTable();
    this.showRows = Math.max(this.showRows || 0, Math.min(this.table.rows.length, (this.showRows || 0) + p.track.added.size));
    this.renderTable();
    if (this.undoBtn) this.undoBtn.disabled = false;
    m.done = 'applied';
    this.drawMsg(m, el);
    if (p.track.added.size && this.previewEl) { this.previewEl.scrollTop = this.previewEl.scrollHeight; if (this.redrawRows) this.redrawRows(); }
  }

  addMsg(role, text, extra) {
    const m = Object.assign({ role, text }, extra || {});
    this.chatLog().push(m);
    if (this.chatEl && this.tab === 'workspace') {
      const hello = this.chatEl.querySelector('.ata-msg.ata-assistant:only-child');
      if (hello && this.chatLog().length === 1) hello.remove();
      return this.drawMsg(m);
    }
    return null;
  }

  renderAttachments() {
    if (!this.attachEl) return;
    this.attachEl.empty();
    this.attachments.forEach((a, i) => {
      const item = this.attachEl.createDiv('ata-att');
      if (a.kind === 'image') item.createEl('img', { attr: { src: a.dataUrl } });
      else if (a.kind === 'pdf') item.createSpan({ text: '📄 ' + a.name });
      else item.createSpan({ text: `▦ ${a.name}（${a.table.rows.length} 行）` });
      const x = item.createEl('button', { text: '×', cls: 'ata-att-x', attr: { 'aria-label': '移除' } });
      x.addEventListener('click', () => { this.attachments.splice(i, 1); this.renderAttachments(); });
    });
  }

  async addAttachments(files) {
    const notice = files.length > 1 || files.some((f) => f.size > 5e6) ? new Notice('正在读取附件…', 0) : null;
    for (const f of files) {
      try {
        const before = this.attachments.length;
        if (isImageFile(f)) this.attachments.push(await imageToPayload(f));
        else if (isPdfFile(f)) this.attachments.push(await pdfToPayload(f));
        else if (TABLE_EXT.test(f.name)) {
          let sheets = await fileToSheets(f);
          if (sheets.length > 1) {
            if (notice) notice.hide();
            sheets = await new Promise((res) => {
              const m = new SheetPickModal(this.app, f.name, sheets, res);
              const close = m.onClose.bind(m);
              m.onClose = () => { close(); res(null); };
              m.open();
            }) || [];
          }
          sheets.forEach((sh) => this.attachments.push({ kind: 'table', name: sh.name ? `${f.name}·${sh.name}` : f.name, table: sh.table, title: sh.title }));
        } else new Notice(`「${f.name}」不是图片、PDF 或表格，已跳过。`);
        /* 记住原文件：发送后自动存进附件区，或者记下是从附件区哪个文件来的 */
        for (let k = before; k < this.attachments.length; k++) {
          Object.defineProperty(this.attachments[k], '_file', { value: f, enumerable: false });
          if (f.__archivePath) Object.defineProperty(this.attachments[k], '_src', { value: f.__archivePath, enumerable: false });
        }
      } catch (e) {
        new Notice(e.message, 8000);
      }
    }
    if (notice) notice.hide();
    if (this.tab !== 'workspace' || this.renderedTab !== 'workspace') { this.tab = 'workspace'; await this.render(); }
    else {
      this.renderAttachments();
      if (this.pendingPrompt && this.inputEl && !this.inputEl.value.trim()) { this.inputEl.value = this.pendingPrompt; this.inputEl.focus(); }
      this.pendingPrompt = null;
    }
  }

  async send() {
    if (this.busy) return;
    const text = this.inputEl.value.trim();
    const atts = this.attachments;
    if (!text && !atts.length) return;
    if (!this.file) { new Notice('请先打开一个表格。'); return; }
    await this.flushSave();
    const file = this.file;

    const files = atts.filter((a) => a.kind === 'image' || a.kind === 'pdf');
    const tables = atts.filter((a) => a.kind === 'table');
    const ask = text || (tables.length || files.length ? MASTER_PROMPT : '');

    this.addMsg('user', ask, { atts: atts.map((a) => ({ name: a.name, kind: a.kind, dataUrl: a.dataUrl })) });
    this.inputEl.value = '';
    this.attachments = [];
    this.renderAttachments();

    const max = Math.max(1, +this.settings.maxRows || 300);
    const d = new Date(), z = (n) => String(n).padStart(2, '0');
    const meta = (this.settings.meta || {})[file.path];
    let prompt = `今天是 ${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}。\n\n【目标表格：${file.name}${(this.settings.meta || {})[file.path] && this.settings.meta[file.path].sheet ? `（Excel「${this.settings.meta[file.path].workbook}」的工作表「${this.settings.meta[file.path].sheet}」）` : ''}】${meta && meta.title ? `\n表格标题 / 说明：${meta.title}` : ''}\n${tableToPrompt(this.table, max)}\n`;
    tables.forEach((a) => {
      const n = a.table.rows.length;
      const show = n <= 60 ? n : 20;
      prompt += `\n【附件表格：${a.name}】共 ${n} 行${a.title ? `，标题：${a.title}` : ''}${show < n ? `；这里只给出前 ${show} 行，请用 merge_attachment 让插件处理全部行` : ''}\n${tableToPrompt(a.table, show)}\n`;
    });
    files.forEach((f) => { prompt += `\n【附件${f.kind === 'pdf' ? ' PDF' : '图片'}：${f.name}】（见上方内容）`; });
    prompt += `\n\n【用户要求】\n${ask}`;

    this.busy = true;
    if (this.sendBtn) this.sendBtn.disabled = true;
    const waiting = this.chatEl && this.chatEl.createDiv({ cls: 'ata-msg ata-assistant ata-wait', text: files.length ? '正在识别附件…' : '思考中…' });
    if (this.chatEl) this.chatEl.scrollTop = this.chatEl.scrollHeight;
    const history = this.histories.get(file.path) || [];
    try {
      const raw = await callAI(this.settings, history.slice(-6), prompt, files);
      if (waiting) waiting.remove();
      const parsed = extractJSON(raw);
      if (!this.file || this.file.path !== file.path) throw new Error('回复到达前已切换了表格，本次修改没有应用。');
      if (!parsed) {
        const said = salvageReply(raw);
        if (!said && raw && String(raw).indexOf('{') < 0) this.addMsg('assistant', String(raw).trim() + '\n\n> 这次没有修改表格。');
        else this.addMsg('assistant', said
          ? said + '\n\n> ⚠ AI 返回的格式不完整，**这次没有修改表格**。可以再发一次，或者让它「按 JSON 格式重新输出修改」。'
          : '⚠ AI 返回的内容无法识别，**这次没有修改表格**。原始回复：\n\n' + String(raw || '（空）').slice(0, 2000), { raw: true });
      } else {
        const attMap = {};
        tables.forEach((a) => { attMap[a.name] = a.table; });
        let changes = [];
        if (Array.isArray(parsed.actions) && parsed.actions.length) {
          const draft = cloneTable(this.table);
          const track = newTrack();
          changes = applyActions(draft, parsed.actions, { attachments: attMap, track });
          const real = track.structural || track.added.size || track.cells.size;
          if (parsed.__truncated) parsed.reply = (parsed.reply || '') + '\n\n> ⚠ AI 的回复被截断了，**最后一条记录可能不完整**，请在下面的预览里核对后再应用。';
          if (real && (this.settings.preview || parsed.__truncated)) {
            this.addMsg('assistant', parsed.reply || '请确认下面的修改：', {
              changes,
              preview: { draft, track, before: this.table, path: file.path, version: this.tableVersion, ask: text, atts: atts.map((a) => a.name) },
            });
          } else if (real) {
            this.pushUndo();
            this.logAI(file.path, this.table, draft, track, changes, text, atts.map((a) => a.name));
            this.table = draft;
            await this.saveTable();
            this.renderTable();
            if (this.undoBtn) this.undoBtn.disabled = false;
            this.addMsg('assistant', parsed.reply || '已完成。', { changes });
          } else {
            this.addMsg('assistant', parsed.reply || '（没有需要修改的地方）', { changes });
          }
        } else {
          this.addMsg('assistant', parsed.reply || '（没有修改）');
        }
      }
      this.archiveAfterSend(atts, file).catch((e) => console.warn('[AI表格助手] 存档附件失败', e));
      const attNote = atts.length ? `\n（附带 ${atts.length} 个附件：${atts.map((a) => a.name).join('、')}）` : '';
      history.push({ role: 'user', content: ask + attNote });
      history.push({ role: 'assistant', content: raw || '{"reply":"","actions":[]}' });
      this.histories.set(file.path, history);
    } catch (e) {
      if (waiting) waiting.remove();
      this.addMsg('error', e.message || String(e));
    } finally {
      this.busy = false;
      if (this.sendBtn) this.sendBtn.disabled = false;
    }
  }

  /* ---------- AI 归类表格文件 ---------- */
  classifyDialog() {
    const lib = this.lib || [];
    if (!lib.length) { new Notice('表格库是空的，先导入一些表格。'); return; }
    const uncat = lib.filter((t) => !t.category).length;
    new ClassifyModal(this.app, lib.length, uncat, (scope, cats) => this.classify(scope, cats)).open();
  }

  async classify(scope, cats) {
    const lib = (this.lib || []).filter((t) => scope === 'all' || !t.category);
    if (!lib.length) { new Notice('没有需要归类的表格。'); return; }
    const notice = new Notice(`AI 正在归类 ${lib.length} 个表格…`, 0);
    try {
      const BATCH = 40;
      let done = 0;
      const existing = [...new Set((this.lib || []).map((t) => t.category).filter(Boolean))];
      for (let i = 0; i < lib.length; i += BATCH) {
        const part = lib.slice(i, i + BATCH);
        let prompt = cats.length ? `【指定类别】${cats.join('、')}\n\n` : existing.length ? `【已有类别，可沿用】${existing.join('、')}\n\n` : '';
        prompt += '【表格列表】\n';
        part.forEach((t, k) => {
          prompt += `\n${k + 1}. 路径：${t.file.path}\n   列名：${t.headers.join('、') || '（空表）'}\n`;
          t.sample.forEach((r) => { prompt += '   样例：' + r.map((c) => String(c).slice(0, 30)).join(' | ') + '\n'; });
        });
        const raw = await callAI(this.settings, [], prompt, [], CLASSIFY_PROMPT);
        const parsed = extractJSON(raw);
        if (!parsed || !Array.isArray(parsed.categories)) throw new Error('AI 返回的格式不对：' + String(raw).slice(0, 200));
        const paths = new Set(part.map((t) => t.file.path));
        for (const c of parsed.categories) {
          if (c && paths.has(c.path) && c.category) {
            await this.setCategory(c.path, String(c.category).trim());
            done++;
            if (!existing.includes(c.category)) existing.push(c.category);
          }
        }
        notice.setMessage(`AI 正在归类… ${Math.min(i + BATCH, lib.length)} / ${lib.length}`);
      }
      notice.hide();
      new Notice(`归类完成：${done} 个表格已打上分类。`);
      this.libFilter = '全部';
      this.tab = 'library';
      await this.render();
    } catch (e) {
      notice.hide();
      new Notice('归类失败：' + (e.message || e), 10000);
    }
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
        if (!f && !opts.silent) new Notice('这篇笔记里还没有表格，AI 写入时会追加到文末。');
      } else {
        this.table = gridToTable(parseCSV(text));
      }
      const lm = (this.settings.meta || {})[file.path];
      this.link = lm && lm.link ? await this.readLink(lm.link.id, lm.link.sheet) : null;
      if (this.link) this.attachProvenance(this.table, this.link.map);
      if (!this.file || this.file.path !== file.path) {
        if (this.file) {
          this.undoMap.set(this.file.path, this.undoStack);
          this.attachMap.set(this.file.path, this.attachments);
          if (this.undoMap.size > 8) this.undoMap.delete(this.undoMap.keys().next().value);
        }
        this.undoStack = this.undoMap.get(file.path) || [];
        this.attachments = this.attachMap.get(file.path) || [];
        this.attachMap.delete(file.path);
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
        await this.app.vault.modify(this.file, out);
      } else {
        await this.app.vault.modify(this.file, toCSV(this.table));
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
    const docs = list.filter((x) => isImageFile(x.file) || isPdfFile(x.file));
    const tables = list.filter((x) => TABLE_EXT.test(x.file.name));
    const isFolder = list.some((x) => x.rel.includes('/'));
    if (zone && zone.startsWith('table:')) {
      const f = this.app.vault.getAbstractFileByPath(zone.slice(6));
      if (f instanceof TFile) { await this.fillTable(f, list); return; }
    }
    if (zone === 'chat' && this.file) { await this.addAttachments([...docs, ...tables].map((x) => x.file)); return; }
    if (zone === 'import' || isFolder) { await this.importBatch(list); return; }
    if (!docs.length && !tables.length) { new Notice('没有可以处理的表格、图片或 PDF。'); return; }
    if (this.tab === 'workspace' && this.file && zone !== 'inbox' && !tables.length) {
      await this.addAttachments(docs.map((x) => x.file));
      return;
    }
    new RouteModal(this.app, this, list).open();
  }

  /** 打开某张表，把资料作为附件放进去，并填好默认指令 */
  async fillTable(file, list) {
    if (!(file instanceof TFile)) return;
    const files = (list || []).map((x) => x.file || x).filter((f) => TABLE_EXT.test(f.name) || isImageFile(f) || isPdfFile(f));
    if (!this.file || this.file.path !== file.path) await this.loadFile(file);
    this.pendingPrompt = MASTER_PROMPT;
    this.tab = 'workspace';
    this.renderedTab = null;
    await this.addAttachments(files);
  }

  /** 删除表格文件（进回收站），并清理分类、常用、最近等记录 */
  async deleteFiles(files) {
    files = files.filter((f) => f instanceof TFile);
    if (!files.length) return;
    const names = files.map((f) => this.tableLabel(f.path));
    this.logOp({ type: 'delete', text: files.length === 1 ? `删除了表格「${names[0]}」（在回收站）` : `删除了 ${files.length} 个表格（在回收站）`, items: names });
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
          this.attachments = [];
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
        this.attachMap.delete(f.path);
        this.chatLogs.delete(f.path);
        this.histories.delete(f.path);
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

  async archiveToFile(tf) {
    const ext = (tf.extension || '').toLowerCase();
    const f = new File([await this.app.vault.readBinary(tf)], tf.name, { type: MIME[ext] || '', lastModified: tf.stat.mtime });
    Object.defineProperty(f, '__archivePath', { value: tf.path });
    return f;
  }

  /** 从附件区送去提取：route = 推荐要填的表；current = 加到当前表 */
  async archiveExtract(tfs, mode) {
    const usable = tfs.filter((tf) => ['image', 'pdf', 'table'].includes(archiveKind(tf)));
    if (!usable.length) { new Notice('只有图片、PDF 和表格文件可以让 AI 提取内容。'); return; }
    if (usable.length < tfs.length) new Notice(`${tfs.length - usable.length} 个文件不是图片 / PDF / 表格，已跳过。`);
    const files = [];
    for (const tf of usable) files.push(await this.archiveToFile(tf));
    const list = files.map((f) => ({ file: f, rel: f.name }));
    if (mode === 'current' && this.file) await this.fillTable(this.file, list);
    else new RouteModal(this.app, this, list).open();
  }

  /** 发送成功后：自动存档新附件、记下附件被提取到了哪张表 */
  async archiveAfterSend(atts, file) {
    const meta = this.settings.archiveMeta || (this.settings.archiveMeta = {});
    const seen = new Set();
    for (const a of atts) {
      let src = a._src;
      if (!src && a._file && this.settings.autoArchive !== false && !seen.has(a._file)) {
        const saved = await this.archiveSave([a._file], '发送给 AI', true);
        if (saved[0]) { src = saved[0].path; Object.defineProperty(a._file, '__archivePath', { value: src, configurable: true }); }
      }
      seen.add(a._file);
      if (src && meta[src]) {
        const used = meta[src].used || (meta[src].used = []);
        if (!used.includes(file.path)) used.push(file.path);
      }
    }
    await this.plugin.saveSettings();
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
    left.createDiv({ cls: 'ata-sub', text: `${all.length} 个文件 · 共 ${fmtSize(all.reduce((a, f) => a + f.stat.size, 0))} · 收到的单据、截图、PDF、Excel 都可以先存在这里，要用时点「送去提取」` });
    const right = head.createDiv('ata-header-right');
    this.btn(right, this.archSelect ? '完成' : '批量管理', this.archSelect ? 'check' : 'list-checks', () => { this.archSelect = !this.archSelect; this.archSel.clear(); this.render(); });
    if (this.app.showInFolder) this.btn(right, '打开文件夹', 'folder-open', async () => { await this.ensureFolder(this.archiveDir()); this.app.showInFolder(this.archiveDir()); });
    this.btn(right, '上传文件', 'upload', () => this.archiveInput.click(), 'mod-cta');

    const dz = wrap.createDiv('ata-dropzone ata-archive-drop');
    this.ico(dz, 'archive', 'ata-dropzone-icon');
    const txt = dz.createDiv('ata-dropzone-text');
    txt.createEl('b', { text: '把文件拖到这里存档（任何类型都可以）' });
    txt.createSpan({ text: `也可以点这里选择文件，或在本页按 Ctrl+V 粘贴复制的文件 / 截图 / 文件路径。保存在「${this.archiveDir()}」，按月份分文件夹。${this.settings.autoArchive !== false ? '发给 AI 的附件也会自动存一份。' : ''}` });
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
      this.btn(sb, '送去提取', 'wand-2', () => { const s = sel(); if (s.length) this.archiveExtract(s, 'route'); }, 'ata-btn-sm mod-cta');
      if (this.file) this.btn(sb, '加到当前表', 'paperclip', () => { const s = sel(); if (s.length) this.archiveExtract(s, 'current'); }, 'ata-btn-sm');
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
      const usable = ['image', 'pdf', 'table'].includes(archiveKind(f));
      if (usable) {
        const go = acts.createEl('button', { cls: 'ata-abtn is-cta', text: '送去提取' });
        go.addEventListener('click', (e) => { e.stopPropagation(); this.archiveExtract([f], 'route'); });
        if (this.file) {
          const cur = acts.createEl('button', { cls: 'ata-abtn', text: '加到当前表', attr: { title: '作为附件加到「' + this.file.basename + '」' } });
          cur.addEventListener('click', (e) => { e.stopPropagation(); this.archiveExtract([f], 'current'); });
        }
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
              lastFile = await this.app.vault.create(path, toCSV(sh.table));
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
      if (this.file && files.some((f) => f.path === this.file.path)) this.addMsg('system', msg);
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
      if (this.file && file.path === this.file.path) this.addMsg('system', `已导出到：${path}`);
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
  display() {
    const { containerEl } = this;
    const s = this.plugin.settings;
    const save = () => this.plugin.saveSettings();
    containerEl.empty();
    const def = PROVIDER_DEFAULTS[providerOf(s)];

    new Setting(containerEl).setName('使用引导').setDesc('一步步介绍每个功能怎么用。')
      .addButton((b) => b.setButtonText('重新查看使用引导').onClick(async () => {
        this.app.setting.close();
        const v = await this.plugin.activateView();
        if (v) v.startTour();
      }));
    new Setting(containerEl).setName('AI 服务').setDesc('DeepSeek：在 platform.deepseek.com 申请 API Key；deepseek-flash 能看图片，带图片 / PDF 时插件会自动用它。「OpenAI 兼容」可接通义千问、Kimi、豆包等提供兼容接口的服务。')
      .addDropdown((d) => {
        Object.keys(PROVIDER_DEFAULTS).forEach((k) => d.addOption(k, PROVIDER_DEFAULTS[k].label));
        d.setValue(providerOf(s)).onChange(async (v) => { s.provider = v; s.baseUrl = ''; s.model = ''; await save(); this.display(); });
      });
    new Setting(containerEl).setName('API Key').setDesc('只保存在本 vault 的插件数据里（data.json）。')
      .addText((t) => { t.inputEl.type = 'password'; t.setValue(s.apiKey).onChange(async (v) => { s.apiKey = v.trim(); await save(); }); });
    new Setting(containerEl).setName('接口地址').setDesc('留空使用默认：' + def.baseUrl)
      .addText((t) => t.setPlaceholder(def.baseUrl).setValue(s.baseUrl).onChange(async (v) => { s.baseUrl = v.trim(); await save(); }));
    new Setting(containerEl).setName('模型').setDesc('留空使用默认：' + def.model + (providerOf(s) === 'deepseek' ? '（也可以填 deepseek-v4-pro：更强但只能处理文字，带图片时会自动换成 deepseek-flash）' : ''))
      .addText((t) => t.setPlaceholder(def.model).setValue(s.model).onChange(async (v) => { s.model = v.trim(); await save(); }));
    new Setting(containerEl).setName('测试连接').setDesc('发一条很短的消息，确认 API Key 和模型可用。')
      .addButton((b) => b.setButtonText('测试').onClick(async () => {
        b.setButtonText('测试中…');
        try {
          const raw = await callAI(s, [], '请只回复 JSON：{"reply":"ok","actions":[]}', [], '你是测试助手，只输出 JSON。');
          new Notice('连接成功 ✓ ' + String(raw).slice(0, 60), 6000);
        } catch (e) {
          new Notice('连接失败：' + (e.message || e), 10000);
        }
        b.setButtonText('测试');
      }));
    new Setting(containerEl).setName('表格保存文件夹').setDesc('导入、新建的表格放在这里，导出文件放在其中的「导出」子文件夹。')
      .addText((t) => t.setValue(s.folder).onChange(async (v) => { s.folder = v.trim() || DEFAULT_SETTINGS.folder; await save(); }));
    new Setting(containerEl).setName('发给 AI 的附件自动存进附件区').setDesc('图片、PDF、表格发给 AI 后，原文件自动存一份到「附件区」，方便以后查找或再次提取。')
      .addToggle((t) => t.setValue(s.autoArchive !== false).onChange(async (v) => { s.autoArchive = v; await save(); }));
    new Setting(containerEl).setName('AI 修改前先预览').setDesc('开启后，AI 的修改会先以「黄色 = 修改、绿色 = 新增」的形式给你看，点「应用」才写入表格。')
      .addToggle((t) => t.setValue(s.preview !== false).onChange(async (v) => { s.preview = v; await save(); }));
    new Setting(containerEl).setName('导出 Excel 后自动打开').setDesc('导出 .xlsx 后用系统默认程序（Excel / WPS）打开。')
      .addToggle((t) => t.setValue(s.openAfterExport !== false).onChange(async (v) => { s.openAfterExport = v; await save(); }));
    new Setting(containerEl).setName('发送给 AI 的最大行数').setDesc('表格超过这个行数时，只把前面的行发给 AI（控制费用）。')
      .addText((t) => t.setValue(String(s.maxRows)).onChange(async (v) => { s.maxRows = Math.max(1, parseInt(v) || 300); await save(); }));
  }
}

class AITableAssistantPlugin extends Plugin {
  async onload() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    if (!this.settings.meta) this.settings.meta = {};
    if (!Array.isArray(this.settings.pinned)) this.settings.pinned = [];
    if (!Array.isArray(this.settings.recent)) this.settings.recent = [];
    if (!this.settings.archiveMeta || typeof this.settings.archiveMeta !== 'object') this.settings.archiveMeta = {};
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
  applyActions, tableToPrompt, extractJSON, repairJSON, salvageReply, normalizeTable, collectDropped, cleanGrid, decodeText, newTrack, upsertRows, readWorkbook, toXLSXBook, fileToSheets, patchWorkbook, cloneTable,
};
module.exports = AITableAssistantPlugin;
