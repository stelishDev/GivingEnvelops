// Minimal, dependency-free spreadsheet reader for the Giving Envelopes app.
// Reads the first worksheet of an .xlsx file (a zip of XML files) into an
// array of rows (arrays of strings). Also reads simple .csv files.
// Uses the browser's built-in DecompressionStream, so it works offline.

const td = new TextDecoder('utf-8');

const u16 = (v, o) => v.getUint16(o, true);
const u32 = (v, o) => v.getUint32(o, true);

async function inflateRaw(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function openZip(buf) {
  const bytes = new Uint8Array(buf);
  const v = new DataView(buf);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (u32(v, i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('This does not look like an .xlsx file. Please choose an Excel (.xlsx) or CSV file.');
  const count = u16(v, eocd + 10);
  let p = u32(v, eocd + 16);
  const entries = new Map();
  for (let n = 0; n < count; n++) {
    if (u32(v, p) !== 0x02014b50) throw new Error('The Excel file appears to be damaged.');
    const method = u16(v, p + 10);
    const csize = u32(v, p + 20);
    const nameLen = u16(v, p + 28);
    const extraLen = u16(v, p + 30);
    const commentLen = u16(v, p + 32);
    const lho = u32(v, p + 42);
    const name = td.decode(bytes.subarray(p + 46, p + 46 + nameLen));
    entries.set(name, { method, csize, lho });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return { bytes, v, entries };
}

async function readEntry(zip, name) {
  const e = zip.entries.get(name);
  if (!e) return null;
  const { v, bytes } = zip;
  const nameLen = u16(v, e.lho + 26);
  const extraLen = u16(v, e.lho + 28);
  const start = e.lho + 30 + nameLen + extraLen;
  const data = bytes.subarray(start, start + e.csize);
  if (e.method === 0) return td.decode(data);
  if (e.method === 8) return td.decode(await inflateRaw(data));
  throw new Error('Unsupported compression method in Excel file.');
}

function decodeEntities(s) {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, g) => {
    switch (g) {
      case 'amp': return '&';
      case 'lt': return '<';
      case 'gt': return '>';
      case 'quot': return '"';
      case 'apos': return "'";
      default: {
        const code = g[1] === 'x' ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10);
        return String.fromCodePoint(code);
      }
    }
  });
}

function attr(tag, name) {
  const m = new RegExp('(?:^|\\s)' + name + '="([^"]*)"').exec(tag);
  return m ? decodeEntities(m[1]) : null;
}

function textOf(xml) {
  // Join all <t> runs (rich text) and ignore phonetic runs.
  const clean = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  let out = '';
  const re = /<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g;
  let m;
  while ((m = re.exec(clean))) out += m[1] ? decodeEntities(m[1]) : '';
  return out;
}

function parseSharedStrings(xml) {
  const list = [];
  if (!xml) return list;
  const re = /<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g;
  let m;
  while ((m = re.exec(xml))) list.push(m[1] ? textOf(m[1]) : '');
  return list;
}

function colIndex(ref) {
  const letters = /^[A-Za-z]+/.exec(ref)[0].toUpperCase();
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function parseSheet(xml, shared) {
  const rows = [];
  let maxCol = 0;
  const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
  let rm;
  let autoRow = 0;
  while ((rm = rowRe.exec(xml))) {
    const rAttr = attr(rm[1], 'r');
    const rIdx = rAttr ? parseInt(rAttr, 10) - 1 : autoRow;
    autoRow = rIdx + 1;
    const cells = [];
    if (rm[2]) {
      const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let cm;
      let autoCol = 0;
      while ((cm = cellRe.exec(rm[2]))) {
        const ref = attr(cm[1], 'r');
        const cIdx = ref ? colIndex(ref) : autoCol;
        autoCol = cIdx + 1;
        const type = attr(cm[1], 't');
        const body = cm[2] || '';
        let val = '';
        if (type === 'inlineStr') {
          val = textOf(body);
        } else {
          const vm = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body);
          const raw = vm ? decodeEntities(vm[1]) : '';
          val = type === 's' ? (shared[parseInt(raw, 10)] ?? '') : raw;
        }
        cells[cIdx] = String(val).trim();
        if (cIdx + 1 > maxCol) maxCol = cIdx + 1;
      }
    }
    rows[rIdx] = cells;
  }
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    const full = [];
    for (let c = 0; c < maxCol; c++) full.push(r[c] ?? '');
    out.push(full);
  }
  return out;
}

async function firstSheetPath(zip) {
  const wb = await readEntry(zip, 'xl/workbook.xml');
  const rels = await readEntry(zip, 'xl/_rels/workbook.xml.rels');
  if (wb && rels) {
    const sheetTag = /<sheet\b[^>]*>/.exec(wb);
    const rid = sheetTag ? attr(sheetTag[0], 'r:id') : null;
    if (rid) {
      const relRe = /<Relationship\b[^>]*>/g;
      let m;
      while ((m = relRe.exec(rels))) {
        if (attr(m[0], 'Id') === rid) {
          let target = attr(m[0], 'Target') || '';
          target = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
          if (zip.entries.has(target)) return target;
        }
      }
    }
  }
  return zip.entries.has('xl/worksheets/sheet1.xml') ? 'xl/worksheets/sheet1.xml' : null;
}

export async function readXlsx(buf) {
  const zip = openZip(buf);
  const path = await firstSheetPath(zip);
  if (!path) throw new Error('No worksheet found in this Excel file.');
  const shared = parseSharedStrings(await readEntry(zip, 'xl/sharedStrings.xml'));
  const xml = await readEntry(zip, path);
  return parseSheet(xml, shared);
}

export function parseCsv(text) {
  text = text.replace(/^﻿/, '');
  const rows = [];
  let row = [];
  let cur = '';
  let quoted = false;
  const delim = (text.split('\n')[0].match(/\t/g) || []).length > (text.split('\n')[0].match(/,/g) || []).length ? '\t' : ',';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; }
      else if (ch === '"') quoted = false;
      else cur += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === delim) { row.push(cur.trim()); cur = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cur.trim()); cur = '';
      rows.push(row); row = [];
    } else cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur.trim()); rows.push(row); }
  const width = Math.max(0, ...rows.map(r => r.length));
  return rows.map(r => { const f = r.slice(); while (f.length < width) f.push(''); return f; });
}

export async function readSpreadsheetFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith('.csv') || name.endsWith('.txt')) return parseCsv(await file.text());
  if (name.endsWith('.xls')) throw new Error('Old .xls files are not supported. In Excel, use Save As → .xlsx and try again.');
  return readXlsx(await file.arrayBuffer());
}
