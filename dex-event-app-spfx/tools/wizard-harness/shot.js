/* Screenshots aus dem Harness (Playwright + Chromium).
 *
 *   node shot.js edit    → alle Wizard-Schritte im Bearbeiten-Modus
 *   node shot.js create  → Wizard im Anlege-Modus
 *   node shot.js pages   → die Teilnehmer-Seiten, je einmal Desktop und Handy
 */
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const OUT = path.join(__dirname, 'out', 'shots');
fs.mkdirSync(OUT, { recursive: true });
const candidates = ['/opt/pw-browsers/chromium-1194/chrome-linux/chrome', '/opt/pw-browsers/chromium/chrome-linux/chrome', '/opt/pw-browsers/chromium'];
const exe = candidates.find(p => { try { return fs.statSync(p).isFile(); } catch { return false; } });
const mode = process.argv[2] || 'edit';
const INDEX = 'file://' + path.join(__dirname, 'out', 'index.html');

/* Die Teilnehmer-Seiten. `ready` ist ein Selektor, der erst steht, wenn die
 * Seite wirklich gerendert hat — ein fester Timeout würde bei einer kaputten
 * Seite ein weißes Bild als Erfolg durchwinken. */
const PAGES = [
  { name: 'landing',   page: 'landing',  ready: '.landing' },
  { name: 'start',     page: 'start',    ready: '.page-container' },
  { name: 'list',      page: 'list',     ready: '.page-container' },
  // Die Listen-Ansicht ist die einzige Stelle der Event-Übersicht mit einem
  // eigenen isMobile-Zweig — ohne sie bliebe der Handy-Auftrag halb sichtbar.
  { name: 'list-liste', page: 'list',    ready: '.page-container', query: '&view=list' },
  { name: 'register',  page: 'register', ready: '.page-container' },
  // Dieselbe Seite für die Klammer: nur hier gibt es die Termin-Auswahl,
  // die Wartelisten-Kachel und die schon gebuchten Termine.
  { name: 'register-klammer', page: 'register', ready: '.page-container', query: '&event=ev-umb' },
  { name: 'myevents',  page: 'myevents', ready: '.page-container' },
];

/** Eine Seite in einem eigenen Kontext öffnen, damit localStorage je Bild leer
 *  startet und ein Fehler nicht in die nächste Seite hineinblutet. */
async function shootPage(browser, spec, mobile) {
  const { name, page: key, ready, query } = spec;
  const suffix = mobile ? 'mobile' : 'desktop';
  const context = await browser.newContext({
    viewport: mobile ? { width: 390, height: 844 } : { width: 1360, height: 900 },
    deviceScaleFactor: mobile ? 2 : 1,
    isMobile: mobile,
    hasTouch: mobile,
  });
  const page = await context.newPage();
  const logs = [];
  page.on('console', m => { if (m.type() === 'error' || m.type() === 'warning') logs.push(`[${m.type()}] ${m.text().slice(0, 300)}`); });
  page.on('pageerror', e => logs.push(`[pageerror] ${String(e).slice(0, 400)}`));
  // Ohne die URL steht in der Console nur „Failed to load resource" — und
  // damit weiß niemand, ob eine Schrift aus dem Netz fehlt (erwartbar) oder
  // ein Bild der Seite (nicht erwartbar).
  page.on('requestfailed', r => logs.push(`[request] ${r.failure()?.errorText || 'failed'} ${r.url().slice(0, 160)}`));
  const url = `${INDEX}?page=${key}${mobile ? '&mobile=1' : ''}${query || ''}`;
  await page.goto(url);
  const ok = await page.waitForSelector(ready, { timeout: 20000 }).then(() => true).catch(() => false);
  if (!ok) logs.push(`[harness] Selektor "${ready}" kam nicht — die Seite hat vermutlich nichts gerendert.`);
  // Nachlauf für Mount-Effekte (Anmeldungen lesen, Bilder, Zähler).
  await page.waitForTimeout(1800);
  const file = path.join(OUT, `page-${name}-${suffix}.png`);
  await page.screenshot({ path: file, fullPage: true });
  fs.writeFileSync(path.join(OUT, `page-${name}-${suffix}-console.log`), logs.join('\n'));
  const errs = logs.filter(l => l.startsWith('[pageerror]')).length;
  console.log(`${name} ${suffix}: ${ok ? 'gerendert' : 'LEER'} · console ${logs.length} · pageerror ${errs}`);
  await context.close();
  return errs;
}

(async () => {
  const browser = await chromium.launch({ executablePath: exe, headless: true });

  if (mode === 'pages') {
    let bad = 0;
    for (const p of PAGES) {
      bad += await shootPage(browser, p, false);
      bad += await shootPage(browser, p, true);
    }
    console.log(bad === 0 ? 'ALLE SEITEN OHNE PAGEERROR' : `PAGEERRORS: ${bad}`);
    await browser.close();
    return;
  }

  const page = await browser.newPage({ viewport: { width: 1360, height: 900 }, deviceScaleFactor: 1 });
  const logs = [];
  page.on('console', m => { if (m.type() === 'error' || m.type() === 'warning') logs.push(`[${m.type()}] ${m.text().slice(0, 300)}`); });
  page.on('pageerror', e => logs.push(`[pageerror] ${String(e).slice(0, 400)}`));
  await page.goto(`${INDEX}?mode=${mode}`);
  await page.waitForSelector('.dex-wizard-step', { timeout: 20000 }).catch(() => {});
  await page.waitForTimeout(800);
  await page.screenshot({ path: path.join(OUT, `${mode}-initial.png`), fullPage: true });
  const n = await page.locator('.dex-wizard-step').count();
  console.log('steps:', n);
  for (let i = 0; i < n; i++) {
    await page.locator(`[data-tour="wizard-step-${i}"]`).click({ timeout: 5000 }).catch(e => logs.push(`click step ${i}: ${e.message.slice(0, 120)}`));
    await page.waitForTimeout(500);
    await page.screenshot({ path: path.join(OUT, `${mode}-step-${i + 1}.png`), fullPage: true });
  }
  fs.writeFileSync(path.join(OUT, `${mode}-console.log`), logs.join('\n'));
  console.log('console lines:', logs.length);
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
