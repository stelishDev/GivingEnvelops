// Run: node tests/test-reader.mjs
import { readFile } from 'node:fs/promises';
import assert from 'node:assert/strict';
import { readXlsx, parseCsv } from '../xlsx-reader.js';

const buf = await readFile(new URL('./sample-envelopes.xlsx', import.meta.url));
const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
const rows = await readXlsx(ab);

assert.deepEqual(rows[0], ['Christian Name:', 'Surname:', 'CONG']);
assert.deepEqual(rows[1], ['Alivia', 'Afamasaga', '7']);
assert.deepEqual(rows.find((r) => r[0] === "D'Neah"), ["D'Neah", 'Cordtz', '9']);
assert.deepEqual(rows.find((r) => r[0] === 'Esté'), ['Esté', 'Ngatai', '1']);
assert.deepEqual(rows.find((r) => r[0] === 'Maleko & Chloe'), ['Maleko & Chloe', 'Afamasaga', '7']);
assert.equal(rows.length, 215);
assert.ok(rows.every((r) => r.length === 3));

const csv = parseCsv('Christian Name:,Surname:,CONG\r\n"Smith, Jr",Jones,3\r\nAnn,Lee,4\r\n');
assert.deepEqual(csv[1], ['Smith, Jr', 'Jones', '3']);
assert.equal(csv.length, 3);

console.log(`OK — read ${rows.length - 1} people from the sample spreadsheet.`);
