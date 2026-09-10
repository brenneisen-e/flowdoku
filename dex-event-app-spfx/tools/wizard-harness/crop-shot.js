/* v31.9.8: Baut crop-entry.tsx und schiesst ein PNG vom Zuschnitt-Dialog.
 * Dieselben Kniffe wie build.js/shot.js (SCSS-Stub, @microsoft-Stub, flach
 * kompiliertes app.css) — nur mit einem anderen Einstiegspunkt.
 * Nicht Teil des Produkt-Builds.  Aufruf: node crop-shot.js
 */
const esbuild = require('esbuild');
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');

const ROOT = path.resolve(__dirname, '..', '..');
const OUT = path.join(__dirname, 'out');
fs.mkdirSync(path.join(OUT, 'shots'), { recursive: true });

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

(async () => {
  await esbuild.build({
    entryPoints: [path.join(__dirname, 'crop-entry.tsx')],
    bundle: true,
    outfile: path.join(OUT, 'crop.js'),
    platform: 'browser', format: 'iife', target: ['es2018'],
    tsconfig: path.join(ROOT, 'tsconfig.json'), jsx: 'transform',
    nodePaths: [path.join(ROOT, 'node_modules')],
    loader: { '.css': 'css', '.png': 'dataurl', '.svg': 'dataurl', '.woff': 'dataurl', '.woff2': 'dataurl', '.ttf': 'dataurl', '.eot': 'dataurl', '.gif': 'dataurl', '.jpg': 'dataurl' },
    define: { 'process.env.NODE_ENV': '"development"', 'global': 'window' },
    plugins: [stubStyles, stubMicrosoft],
    logLevel: 'warning', sourcemap: false, minify: false,
  });

  const html = `<!doctype html><html lang="de"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="app.css"></head>
<body style="margin:0;background:#f3f3f1"><div id="root"></div><script src="crop.js"></script></body></html>`;
  fs.writeFileSync(path.join(OUT, 'crop.html'), html);

  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  for (const [name, w, h] of [['desktop', 1280, 1000], ['mobile', 400, 900]]) {
    const page = await browser.newPage({ viewport: { width: w, height: h }, deviceScaleFactor: 2 });
    await page.goto('file://' + path.join(OUT, 'crop.html'));
    await page.waitForTimeout(1200);
    const file = path.join(OUT, 'shots', `crop-${name}.png`);
    await page.screenshot({ path: file });
    console.log('SHOT ' + file);
    await page.close();
  }
  await browser.close();
})();
