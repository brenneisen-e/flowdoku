/* Screenshots der App-HÜLLE (Header + .app-layout + .main-content).
 *
 *   node shell-build.js && node shell-shot.js
 *
 * Zwei Breiten je Ansicht: 1280 (Desktop) und 400 (Handy). Es wird bewusst
 * NICHT `fullPage` geschossen — `.app-layout` bekommt seine Höhe per JS und
 * `html, body` stehen auf `overflow: hidden`, genau wie live. Ein
 * fullPage-Bild würde eine Seite zeigen, die es in SharePoint nicht gibt.
 *
 * Zusätzlich zum PNG schreibt jeder Lauf eine `*-dom.json` neben das Bild:
 * die gemessene Schachtelung (Klasse, Kasten, die tragenden CSS-Werte). Ein
 * Screenshot zeigt, wie es aussieht — die JSON sagt, WORAN es liegt.
 */
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const OUT = path.join(__dirname, 'out', 'shots');
fs.mkdirSync(OUT, { recursive: true });
const candidates = ['/opt/pw-browsers/chromium-1194/chrome-linux/chrome', '/opt/pw-browsers/chromium/chrome-linux/chrome', '/opt/pw-browsers/chromium'];
const exe = candidates.find(p => { try { return fs.statSync(p).isFile(); } catch { return false; } });
const INDEX = 'file://' + path.join(__dirname, 'out', 'shell', 'index.html');

const SHOTS = [
  { name: 'shell-landing-admin', view: 'landing', role: 'admin', ready: '.landing' },
  { name: 'shell-landing-user', view: 'landing', role: 'user', ready: '.landing' },
  { name: 'shell-start-admin', view: 'start', role: 'admin', ready: '.page-container' },
];

/** Misst die Kette vom App-Wurzel-Element bis zur Seite. Genau die Werte, an
 *  denen eine Übertragung auf eine andere App scheitert. */
const PROBE = () => {
  const sel = [
    '.dexApp', '.app-layout', 'header.header', '.header-left', '.header-right',
    '.main-content', '.landing', '.landing__hero', '.landing__card', '.landing__orb',
    '.landing__text', '.page-container', '.dex-cluster-grid', '.dex-cluster-tiles',
    '.start-card', '.start-card__icon', '.start-grid',
  ];
  const out = {};
  for (const s of sel) {
    const el = document.querySelector(s);
    if (!el) { out[s] = null; continue; }
    const r = el.getBoundingClientRect();
    const c = getComputedStyle(el);
    out[s] = {
      box: { x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height) },
      display: c.display, position: c.position, flex: c.flex, height: c.height,
      minHeight: c.minHeight, maxWidth: c.maxWidth, overflowY: c.overflowY,
      padding: c.padding, margin: c.margin, gap: c.gap,
      gridTemplateColumns: c.gridTemplateColumns,
      background: c.backgroundColor, borderBottom: c.borderBottomWidth + ' ' + c.borderBottomColor,
      scrollHeight: el.scrollHeight, clientHeight: el.clientHeight,
      childCount: el.children.length,
    };
  }
  out['__body'] = { overflow: getComputedStyle(document.body).overflow, scrollHeight: document.body.scrollHeight };
  out['__startCards'] = document.querySelectorAll('.start-card').length;
  return out;
};

async function shoot(browser, spec, width) {
  const mobile = width <= 480;
  const context = await browser.newContext({
    viewport: { width, height: mobile ? 860 : 900 },
    deviceScaleFactor: 1,
    isMobile: mobile,
    hasTouch: mobile,
  });
  const page = await context.newPage();
  const logs = [];
  page.on('console', m => { if (m.type() === 'error' || m.type() === 'warning') logs.push(`[${m.type()}] ${m.text().slice(0, 300)}`); });
  page.on('pageerror', e => logs.push(`[pageerror] ${String(e).slice(0, 400)}`));
  page.on('requestfailed', r => logs.push(`[request] ${r.failure()?.errorText || 'failed'} ${r.url().slice(0, 140)}`));
  await page.goto(`${INDEX}?view=${spec.view}&role=${spec.role}${mobile ? '&mobile=1' : ''}`);
  const ok = await page.waitForSelector(spec.ready, { timeout: 20000 }).then(() => true).catch(() => false);
  if (!ok) logs.push(`[harness] Selektor "${spec.ready}" kam nicht — nichts gerendert.`);
  await page.waitForTimeout(1800);
  const file = path.join(OUT, `${spec.name}-${width}.png`);
  await page.screenshot({ path: file });
  const dom = await page.evaluate(PROBE);
  fs.writeFileSync(path.join(OUT, `${spec.name}-${width}-dom.json`), JSON.stringify(dom, null, 2));
  fs.writeFileSync(path.join(OUT, `${spec.name}-${width}-console.log`), logs.join('\n'));
  const errs = logs.filter(l => l.startsWith('[pageerror]')).length;
  console.log(`${spec.name} @${width}: ${ok ? 'gerendert' : 'LEER'} · console ${logs.length} · pageerror ${errs}`);
  await context.close();
  return errs;
}

(async () => {
  const browser = await chromium.launch({ executablePath: exe, headless: true });
  let bad = 0;
  for (const s of SHOTS) {
    bad += await shoot(browser, s, 1280);
    bad += await shoot(browser, s, 400);
  }
  console.log(bad === 0 ? 'ALLE OHNE PAGEERROR' : `PAGEERRORS: ${bad}`);
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
