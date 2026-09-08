/* Screenshots aller Wizard-Schritte aus dem Harness (Playwright + Chromium). */
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const OUT = path.join(__dirname, 'out', 'shots');
fs.mkdirSync(OUT, { recursive: true });
const candidates = ['/opt/pw-browsers/chromium-1194/chrome-linux/chrome', '/opt/pw-browsers/chromium/chrome-linux/chrome', '/opt/pw-browsers/chromium'];
const exe = candidates.find(p => { try { return fs.statSync(p).isFile(); } catch { return false; } });
const mode = process.argv[2] || 'edit';

(async () => {
  const browser = await chromium.launch({ executablePath: exe, headless: true });
  const page = await browser.newPage({ viewport: { width: 1360, height: 900 }, deviceScaleFactor: 1 });
  const logs = [];
  page.on('console', m => { if (m.type() === 'error' || m.type() === 'warning') logs.push(`[${m.type()}] ${m.text().slice(0, 300)}`); });
  page.on('pageerror', e => logs.push(`[pageerror] ${String(e).slice(0, 400)}`));
  await page.goto('file://' + path.join(__dirname, 'out', 'index.html') + `?mode=${mode}`);
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
