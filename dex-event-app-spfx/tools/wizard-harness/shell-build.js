/* Baut das Bundle für die App-HÜLLE (shell-entry.tsx) nach out/shell/.
 *
 * Getrennt von build.js, damit das bestehende out/bundle.js unberührt bleibt:
 * `node build.js` und `node shell-build.js` dürfen sich nicht überschreiben.
 * Plugins und Loader sind bewusst identisch zu build.js — wer dort etwas
 * ändert, ändert es hier mit.
 */
const esbuild = require('esbuild');
const path = require('path');
const fs = require('fs');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..', '..');
const OUT = path.join(__dirname, 'out', 'shell');
fs.mkdirSync(OUT, { recursive: true });

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
    entryPoints: [path.join(__dirname, 'shell-entry.tsx')],
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

  const scss = path.join(ROOT, 'src', 'webparts', 'dexEventPlatform', 'components', 'DexEventPlatform.module.scss');
  const cssOut = path.join(OUT, 'app.css');
  execFileSync(path.join(ROOT, 'node_modules', '.bin', 'sass'), ['--no-source-map', scss, cssOut], { stdio: 'inherit' });
  let css = fs.readFileSync(cssOut, 'utf8');
  css = css.replace(/:global\(([^)]*)\)/g, '$1').replace(/ ?:global ?/g, ' ');
  fs.writeFileSync(cssOut, css);

  // Eigenes index.html: `#root { min-height: 100vh }` aus dem Wizard-Harness
  // waere hier falsch — die Huelle setzt ihre Hoehe selbst und `html, body`
  // stehen auf `overflow: hidden`.
  fs.writeFileSync(path.join(OUT, 'index.html'), `<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DEX Shell Harness</title>
<link rel="stylesheet" href="app.css">
<link rel="stylesheet" href="bundle.css">
<style>
  html, body { margin: 0; background: #f5f5f5; font-family: Aptos, "Open Sans", "Segoe UI", Arial, Helvetica, sans-serif; }
</style>
</head>
<body>
<div id="root"></div>
<script src="bundle.js"></script>
</body>
</html>
`);
  console.log('SHELL_BUILD_OK');
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
