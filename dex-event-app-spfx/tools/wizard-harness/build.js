/* v31.2: Render-Harness für den Event-Wizard — ausserhalb von SharePoint.
 *
 * Warum: SPFx hat seit 1.13 keinen lokalen Workbench mehr, und der gehostete
 * braucht den Tenant. Wer sehen will, wie ein Wizard-Schritt aussieht, musste
 * bauen, hochladen, klicken. Der Harness bündelt EventCreationPage mit
 * gemockten Providern (Events, Rollen, Navigation, User) und echten
 * Language-/Dialog-Providern, kompiliert das SCSS-Modul flach und lässt
 * Chromium je Schritt ein PNG machen (shot.js).
 *
 * Einmalig:  cd tools/wizard-harness && npm i --no-audit --no-fund esbuild@0.21.5 playwright@1.55.0
 * Bauen:     node build.js          → out/bundle.js, out/bundle.css, out/app.css, out/index.html
 * Schiessen: node shot.js edit      → out/shots/edit-step-1.png … (Edit-Modus mit Beispiel-Event)
 *            node shot.js create    → out/shots/create-initial.png (Anlege-Modus: Nutzungsbedingungen)
 *
 * Nicht Teil des Produkt-Builds. Verhalten mit Dienst-Aufrufen (Speichern,
 * Suche) ist im Harness leer — es geht um Darstellung, Reihenfolge, Texte.
 */
const esbuild = require('esbuild');
const path = require('path');
const fs = require('fs');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..', '..');
const OUT = path.join(__dirname, 'out');
fs.mkdirSync(OUT, { recursive: true });

// SCSS-Module: im Bundle durch einen Proxy ersetzen, der jeden Klassennamen
// unverändert zurückgibt (styles.dexApp → 'dexApp'). Das flach kompilierte
// app.css trägt dieselben Namen.
const stubStyles = {
  name: 'stub-scss-modules',
  setup(build) {
    build.onResolve({ filter: /\.module\.scss$/ }, () => ({ path: 'dex-styles-stub', namespace: 'stub' }));
    build.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({
      contents: `const p = new Proxy({}, { get: (_t, k) => (typeof k === 'string' ? k : '') }); module.exports = { __esModule: true, default: p };`,
      loader: 'js',
    }));
  },
};

// SPFx-Laufzeitpakete (@microsoft/sp-*) ziehen interne Pakete, die es
// ausserhalb des SPFx-Bundlers nicht gibt. Rekursiver Proxy: jede Eigenschaft
// ist wieder ein aufrufbarer Proxy (Klasse, Funktion, Konstante).
const stubMicrosoft = {
  name: 'stub-microsoft',
  setup(build) {
    build.onResolve({ filter: /^@microsoft\// }, (args) => ({ path: args.path, namespace: 'ms-stub' }));
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
    loader: { '.css': 'css', '.png': 'dataurl', '.svg': 'dataurl', '.woff': 'dataurl', '.woff2': 'dataurl', '.ttf': 'dataurl', '.eot': 'dataurl', '.gif': 'dataurl', '.jpg': 'dataurl' },
    define: { 'process.env.NODE_ENV': '"development"', 'global': 'window' },
    plugins: [stubStyles, stubMicrosoft],
    logLevel: 'warning',
    sourcemap: false,
    minify: false,
  });

  // SCSS flach kompilieren. `sass` kennt :global nicht (das ist CSS-Module-
  // Syntax) und würde es als Selektor ausgeben — deshalb danach entfernen.
  const scss = path.join(ROOT, 'src', 'webparts', 'dexEventPlatform', 'components', 'DexEventPlatform.module.scss');
  const cssOut = path.join(OUT, 'app.css');
  execFileSync(path.join(ROOT, 'node_modules', '.bin', 'sass'), ['--no-source-map', scss, cssOut], { stdio: 'inherit' });
  let css = fs.readFileSync(cssOut, 'utf8');
  css = css.replace(/:global\(([^)]*)\)/g, '$1').replace(/ ?:global ?/g, ' ');
  fs.writeFileSync(cssOut, css);

  fs.copyFileSync(path.join(__dirname, 'index.html'), path.join(OUT, 'index.html'));
  console.log('BUILD_OK');
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
