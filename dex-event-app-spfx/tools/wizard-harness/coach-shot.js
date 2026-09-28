/* v32.1.0: Das Mitmach-Tutorial einmal wie ein Mensch durchspielen.
 *
 *   node build.js && node coach-shot.js   → out/shots/coach-NN-*.png
 *
 * Tippt, klickt und wählt im echten Assistenten (Harness-Mocks, nichts wird
 * gespeichert) und fotografiert jede Station. Bricht NICHT beim ersten Fehler
 * ab: Jede Station meldet OK/FEHLT, damit man sieht, wo der Coach hängt.
 */
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const OUT = path.join(__dirname, 'out', 'shots');
fs.mkdirSync(OUT, { recursive: true });
const candidates = ['/opt/pw-browsers/chromium', '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', '/opt/pw-browsers/chromium/chrome-linux/chrome'];
const exe = candidates.find(p => { try { return fs.statSync(p).isFile(); } catch { return false; } });
const INDEX = 'file://' + path.join(__dirname, 'out', 'index.html');

(async () => {
  const browser = await chromium.launch(exe ? { executablePath: exe } : {});
  const page = await browser.newPage({ viewport: { width: 1360, height: 900 } });
  page.setDefaultTimeout(6000);
  const errors = [];
  page.on('pageerror', e => errors.push(String(e && e.message || e)));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await page.goto(INDEX + '?mode=create&coach=1', { waitUntil: 'domcontentloaded' });
  let n = 0;
  const card = page.locator('[data-coach] [role="dialog"]');
  const titel = async () => (await card.locator('h3').textContent().catch(() => '')) || '';
  const shot = async (name) => {
    n += 1;
    const f = path.join(OUT, `coach-${String(n).padStart(2, '0')}-${name}.png`);
    await page.waitForTimeout(450);
    await page.screenshot({ path: f });
    console.log(`${String(n).padStart(2, '0')} ${name.padEnd(14)} Karte: ${await titel()}`);
  };
  const coachWeiter = async () => { await card.getByRole('button', { name: /Weiter|Los geht/ }).click(); };
  const sichtbar = async (sel) => {
    const all = page.locator(sel);
    const c = await all.count();
    for (let i = 0; i < c; i++) { if (await all.nth(i).isVisible()) return all.nth(i); }
    return all.first();
  };
  const schritt = async (label, fn) => { try { await fn(); } catch (e) { console.log(`   FEHLT bei ${label}: ${String(e.message).split('\n')[0]}`); } };

  await page.waitForSelector('[data-coach]', { timeout: 15000 });
  await shot('intro');
  await schritt('intro', coachWeiter);

  // Nutzungsbedingungen (nur wenn der Dialog kommt)
  await page.waitForTimeout(600);
  await shot('terms');
  await schritt('terms', async () => {
    const dlg = page.locator('[role="dialog"]:not([data-coach] [role="dialog"])').last();
    const boxes = dlg.locator('input[type="checkbox"]');
    const c = await boxes.count();
    for (let i = 0; i < c; i++) await boxes.nth(i).check({ force: true });
    await dlg.getByRole('button', { name: /Akzeptieren/ }).click();
  });
  // Admins bekommen danach die Abrechnungsfrage — beim Test-Event „Nein".
  await page.waitForTimeout(600);
  await shot('abrechnung');
  await schritt('abrechnung', async () => { await page.getByRole('button', { name: /^Nein$/ }).click(); });

  await page.waitForTimeout(700);
  await shot('titel-leer');
  await schritt('titel', async () => { await (await sichtbar('[data-tour="wizard-title"]')).fill('Team-Frühstück'); });
  await shot('titel-getippt');
  await schritt('titel-weiter', coachWeiter);

  await shot('datum-leer');
  await schritt('datum', async () => {
    const inputs = page.locator('[data-tour="wizard-dates"] input');
    const inTagen = (t, h) => { const d = new Date(Date.now() + t * 86400000); return `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}, ${String(h).padStart(2, '0')}:00`; };
    await inputs.nth(0).click(); await inputs.nth(0).fill(inTagen(7, 9)); await page.keyboard.press('Enter');
    await inputs.nth(1).click(); await inputs.nth(1).fill(inTagen(7, 11)); await page.keyboard.press('Enter');
    await page.keyboard.press('Escape');
  });
  await shot('datum-gesetzt');
  await schritt('datum-weiter', coachWeiter);
  await shot('entwurf');
  await schritt('entwurf-weiter', coachWeiter);
  await shot('weiter-1');
  await schritt('weiter-1', async () => { await (await sichtbar('[data-tour="wizard-next"]')).click(); });
  await page.waitForTimeout(700);
  await shot('organizer');
  await schritt('weiter-2', async () => { await (await sichtbar('[data-tour="wizard-next"]')).click(); });
  await page.waitForTimeout(700);
  await shot('ort');
  await schritt('ort', async () => { await (await sichtbar('[data-tour="wizard-location"]')).fill('Büro Köln, Raum 3'); });
  await schritt('ort-weiter', coachWeiter);
  await shot('weiter-3');
  await schritt('weiter-3', async () => { await (await sichtbar('[data-tour="wizard-next"]')).click(); });
  await page.waitForTimeout(700);
  await shot('plaetze');
  await schritt('plaetze-begrenzt', async () => { await (await sichtbar('[data-tour="wizard-capacity-mode"] button:nth-child(2)')).click(); });
  await page.waitForTimeout(400);
  await shot('plaetze-feld');
  await schritt('plaetze', async () => { await (await sichtbar('[data-tour="wizard-capacity"]')).fill('10'); });
  await schritt('plaetze-weiter', coachWeiter);
  await schritt('weiter-4', async () => { await (await sichtbar('[data-tour="wizard-next"]')).click(); });
  await page.waitForTimeout(700);
  await shot('frage');
  await schritt('frage-add', async () => { await (await sichtbar('[data-tour="wizard-add-question"]')).click(); });
  await page.waitForTimeout(500);
  await shot('frage-neu');
  await schritt('frage-label', async () => { await (await sichtbar('[data-tour="wizard-question-label"]')).fill('Kaffee oder Tee?'); });
  await shot('frage-getippt');
  await schritt('frage-weiter', coachWeiter);
  await schritt('weiter-5', async () => { await (await sichtbar('[data-tour="wizard-next"]')).click(); });
  await page.waitForTimeout(700);
  await shot('kanal');
  await schritt('kanal-weiter', coachWeiter);
  await shot('anlegen');
  await schritt('anlegen', async () => { await (await sichtbar('[data-tour="wizard-create-early"]')).click(); });
  await page.waitForTimeout(900);
  await shot('countdown');

  console.log(errors.length ? `\nKonsolen-Fehler (${errors.length}):\n  ` + Array.from(new Set(errors)).slice(0, 12).join('\n  ') : '\nKeine Konsolen-Fehler.');
  await browser.close();
})();
