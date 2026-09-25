'use strict';
/**
 * v31.97: Welche Module stecken in welchem Chunk — und wie groß sind sie?
 *
 * Schreibt nach dem webpack-Lauf eine schlanke JSON-Datei: je Chunk Name,
 * Dateien und die enthaltenen Module mit ihrer Größe VOR der Minifikation
 * (webpack kennt nur diese). Die Minifikation schrumpft alle Module etwa
 * gleich stark, für die Anteile reicht das. tools/bundle-report.js rechnet
 * daraus die Tabelle, inkl. „was davon lädt jede Seite" (Einstiegs-Chunk).
 */
const fs = require('fs');
const path = require('path');

function attachBundleStats(config, outFile) {
  config.plugins = config.plugins || [];
  config.plugins.push({
    apply(compiler) {
      compiler.hooks.done.tap('DexBundleStats', stats => {
        // Direkt aus dem Chunk-Graphen: toJson gruppiert Module nach Pfad
        // und verschachtelt die zusammengefassten (ConcatenatedModule).
        const comp = stats.compilation;
        const cg = comp.chunkGraph;
        const flat = m => (m.modules ? Array.from(m.modules).flatMap(flat) : [m]);
        const chunks = Array.from(comp.chunks).map(c => ({
          id: c.id,
          initial: c.canBeInitial(),
          files: Array.from(c.files),
          modules: cg.getChunkModules(c).flatMap(flat).map(m => ({
            name: m.readableIdentifier(comp.requestShortener),
            size: m.size(),
          })),
        }));
        const assets = Object.keys(comp.assets).map(n => ({ name: n, size: comp.assets[n].size() }));
        fs.mkdirSync(path.dirname(outFile), { recursive: true });
        fs.writeFileSync(outFile, JSON.stringify({ chunks, assets }));
        console.log('[bundle-stats] geschrieben: ' + outFile);
      });
    },
  });
}

module.exports = { attachBundleStats };
