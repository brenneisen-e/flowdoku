'use strict';

const path = require('path');
const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);
  result.set('serve', result.get('serve-deprecated'));
  return result;
};

// v30.40: Persistenter Terser-Cache. Rund 40 der 56 webpack-Sekunden sind reine
// Minifikation; der Cache holt sie zurück (gemessen: 52 s → 14 s, Ausgabe
// byte-identisch über alle 78 Dateien). Details und die Fallstricke stehen in
// build/minify-cache.js — insbesondere, warum webpacks eigener
// `cache: { type: 'filesystem' }` hier NICHT geht.
//
// Kein Cache im `serve`-Lauf: Dort wird nicht minifiziert, der Wrapper wäre
// wirkungslos und stünde nur im Weg.
build.configureWebpack.mergeConfig({
  additionalConfiguration: (config) => {
    try {
      const { attachMinifyCache } = require('./build/minify-cache');
      const ok = attachMinifyCache(config, path.join(__dirname, '.minifycache'));
      if (!ok) console.warn('[minify-cache] Minifier nicht gefunden — Build läuft unverändert, nur langsamer.');
    } catch (e) {
      // Der Cache ist Kür. Fehlt er, baut das Projekt unverändert — nur langsamer.
      console.warn('[minify-cache] übersprungen: ' + e.message);
    }
    return config;
  },
});

build.initialize(require('gulp'));
