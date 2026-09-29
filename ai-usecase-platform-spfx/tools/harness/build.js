/* Baut den Screenshot-Harness: out/bundle.js, out/app.css, out/index.html.
 *
 * Einmalig:  cd tools/harness && npm i --no-audit --no-fund
 * Bauen:     node build.js
 *
 * Nicht Teil des Produkt-Builds — kein gulp, kein Eintrag in tsconfig.json
 * (dessen `include` nennt nur `src/**`, `tools/` bleibt außen vor).
 * Kein Typ-Check: esbuild entfernt Typen nur. Den Typ-Check macht der
 * Produkt-Build.
 */
const esbuild = require('esbuild');
const path = require('path');
const fs = require('fs');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..', '..');
const OUT = path.join(__dirname, 'out');
fs.mkdirSync(OUT, { recursive: true });

// SCSS-Modul: im Bundle ein Proxy, der jeden Klassennamen unverändert
// zurückgibt (styles.dexApp → 'dexApp'). Das flach kompilierte app.css trägt
// dieselben Namen.
const stubStyles = {
  name: 'stub-scss-modules',
  setup(build) {
    build.onResolve({ filter: /\.module\.scss$/ }, () => ({ path: 'aiuc-styles-stub', namespace: 'stub' }));
    build.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({
      contents: `const p = new Proxy({}, { get: (_t, k) => (typeof k === 'string' ? k : '') }); module.exports = { __esModule: true, default: p };`,
      loader: 'js',
    }));
  },
};

// SPFx-Laufzeitpakete (@microsoft/sp-*) gibt es ausserhalb des SPFx-Bundlers
// nicht. Zwei Wege:
//
//  - `@microsoft/sp-http` bekommt echte Exporte. Die App liest
//    `SPHttpClient.configurations.v1` als WERT; ein Proxy trägt das NICHT,
//    weil esbuild benannte Importe über die eigenen Schlüssel des Moduls
//    auflöst und ein Proxy über eine Funktion nur `length`/`name`/`prototype`
//    hat — `SPHttpClient` käme als `undefined` an (erster Rauchtest, das
//    Anlegen der Listen stürzte mit „reading 'configurations'").
//  - alles andere unter `@microsoft/` ist ein rekursiver Proxy: jede
//    Eigenschaft ist wieder ein aufrufbarer Proxy. Das genügt für Typen, die
//    esbuild ohnehin entfernt.
const stubMicrosoft = {
  name: 'stub-microsoft',
  setup(build) {
    build.onResolve({ filter: /^@microsoft\// }, (args) => ({ path: args.path, namespace: args.path === '@microsoft/sp-http' ? 'ms-http' : 'ms-stub' }));
    build.onLoad({ filter: /.*/, namespace: 'ms-http' }, () => ({
      contents: `
        class SPHttpClient {}
        SPHttpClient.configurations = { v1: { flags: 'harness-v1' } };
        module.exports = { SPHttpClient, SPHttpClientResponse: class SPHttpClientResponse {}, SPHttpClientConfiguration: class SPHttpClientConfiguration {} };
      `,
      loader: 'js',
    }));
    build.onLoad({ filter: /.*/, namespace: 'ms-stub' }, () => ({
      contents: `
        const handler = { get: (t, k) => (k === '__esModule' ? true : k === 'then' ? undefined : (k === Symbol.toPrimitive ? () => '' : mk())), apply: () => mk(), construct: () => mk() };
        function mk() { return new Proxy(function () {}, handler); }
        module.exports = mk();
      `,
      loader: 'js',
    }));
  },
};

async function main() {
  await esbuild.build({
    entryPoints: [path.join(__dirname, 'entry.tsx')],
    bundle: true,
    outfile: path.join(OUT, 'bundle.js'),
    platform: 'browser',
    format: 'iife',
    target: ['es2018'],
    tsconfig: path.join(ROOT, 'tsconfig.json'),
    jsx: 'transform',
    nodePaths: [path.join(ROOT, 'node_modules')],
    loader: { '.css': 'css', '.png': 'dataurl', '.svg': 'dataurl', '.woff': 'dataurl', '.woff2': 'dataurl', '.ttf': 'dataurl', '.gif': 'dataurl', '.jpg': 'dataurl' },
    define: { 'process.env.NODE_ENV': '"development"', 'global': 'window' },
    plugins: [stubStyles, stubMicrosoft],
    logLevel: 'warning',
    sourcemap: false,
    minify: false,
  });

  // SCSS flach kompilieren. `sass` kennt :global nicht (das ist CSS-Module-
  // Syntax) und würde es als Selektor ausgeben — deshalb danach entfernen.
  const scss = path.join(ROOT, 'src', 'webparts', 'aiUseCasePlatform', 'components', 'AiUseCasePlatform.module.scss');
  const cssOut = path.join(OUT, 'app.css');
  execFileSync(path.join(ROOT, 'node_modules', '.bin', 'sass'), ['--no-source-map', '--quiet', scss, cssOut], { stdio: 'inherit' });
  let css = fs.readFileSync(cssOut, 'utf8');
  css = css.replace(/:global\(([^)]*)\)/g, '$1').replace(/ ?:global ?/g, ' ');
  fs.writeFileSync(cssOut, css);

  fs.copyFileSync(path.join(__dirname, 'index.html'), path.join(OUT, 'index.html'));
  console.log('BUILD_OK');
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
