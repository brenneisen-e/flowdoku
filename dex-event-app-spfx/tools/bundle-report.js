#!/usr/bin/env node
'use strict';
/**
 * v31.97: Größen-Bericht zum Bundle.
 *
 *   DEX_BUNDLE_STATS=1 ./node_modules/.bin/gulp bundle --ship
 *   node tools/bundle-report.js [Suchbegriff …]
 *
 * Liest temp/bundle-stats.json (s. build/bundle-stats.js) und die fertigen
 * Dateien aus release/assets. webpack kennt Modulgrößen nur VOR der
 * Minifikation; die minifizierte und die gezippte Größe je Modul wird deshalb
 * anteilig aus seinem Chunk umgelegt — für „wie viel macht X aus" genau
 * genug, für einzelne Bytes nicht.
 *
 * Ausgabe: Chunks (Haupt-Bundle = lädt jede Seite), die größten Module, und
 * je Suchbegriff (Regex auf den Modulpfad) die Summe, z.B. „hotel".
 */
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const root = path.join(__dirname, '..');
const stats = JSON.parse(fs.readFileSync(path.join(root, 'temp', 'bundle-stats.json'), 'utf8'));
const assetsDir = path.join(root, 'release', 'assets');

const kb = n => (n / 1024).toFixed(0).padStart(6) + ' KB';
const rows = [];
const chunkRows = [];
for (const c of stats.chunks) {
  const file = (c.files || []).find(f => f.endsWith('.js'));
  if (!file || !fs.existsSync(path.join(assetsDir, file))) continue;
  const buf = fs.readFileSync(path.join(assetsDir, file));
  const min = buf.length;
  const zip = zlib.deflateRawSync(buf, { level: 9 }).length;
  const raw = c.modules.reduce((a, m) => a + m.size, 0) || 1;
  chunkRows.push({ file, initial: c.initial, min, zip, top: c.modules.slice().sort((a, b) => b.size - a.size)[0] });
  for (const m of c.modules) {
    rows.push({ name: m.name, file, initial: c.initial, min: min * m.size / raw, zip: zip * m.size / raw });
  }
}

const totalMin = chunkRows.reduce((a, c) => a + c.min, 0);
const totalZip = chunkRows.reduce((a, c) => a + c.zip, 0);
console.log(`\nAlle Chunks: ${kb(totalMin)} minifiziert, ${kb(totalZip)} gezippt (≈ Anteil im .sppkg)\n`);
console.log('Chunks (★ = Haupt-Bundle, lädt jede Seite):');
chunkRows.sort((a, b) => b.min - a.min).slice(0, 15).forEach(c => {
  console.log(`  ${c.initial ? '★' : ' '} ${kb(c.min)} ${kb(c.zip)}  ${c.file.slice(0, 44).padEnd(44)} ${c.top ? c.top.name.replace(/^.*\/(lib|node_modules)\//, '$1/').slice(0, 60) : ''}`);
});

console.log('\nGrößte Module (minifiziert / gezippt, anteilig):');
rows.slice().sort((a, b) => b.min - a.min).slice(0, 25).forEach(r => {
  console.log(`  ${r.initial ? '★' : ' '} ${kb(r.min)} ${kb(r.zip)}  ${r.name.replace(/^\.\/(lib\/webparts\/dexEventPlatform\/)?/, '').slice(0, 90)}`);
});

for (const q of process.argv.slice(2)) {
  const re = new RegExp(q, 'i');
  const hit = rows.filter(r => re.test(r.name));
  const min = hit.reduce((a, r) => a + r.min, 0);
  const zip = hit.reduce((a, r) => a + r.zip, 0);
  const initMin = hit.filter(r => r.initial).reduce((a, r) => a + r.min, 0);
  console.log(`\n„${q}": ${hit.length} Module, ${kb(min)} minifiziert, ${kb(zip)} gezippt, davon im Haupt-Bundle ${kb(initMin)}`);
  hit.sort((a, b) => b.min - a.min).slice(0, 12).forEach(r => console.log(`  ${r.initial ? '★' : ' '} ${kb(r.min)} ${kb(r.zip)}  ${r.name.replace(/^\.\/(lib\/webparts\/dexEventPlatform\/)?/, '').slice(0, 90)}`));
}
