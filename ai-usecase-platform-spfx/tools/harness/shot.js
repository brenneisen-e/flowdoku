/* Screenshot-Lauf: klickt sich durch die ECHTE App und macht je Ansicht ein Bild.
 *
 *   node shot.js                          alles (3 Rollen × Rechner/Handy, dazu Fehler- und Leerzustände)
 *   node shot.js role=user device=mobile  nur diese Rollen/Geräte (Komma-Liste erlaubt)
 *   node shot.js lang=en                  die Reisen in englischer Oberfläche
 *   node shot.js states=0                 ohne Fehler-/Leerzustände
 *   node shot.js only=admin-desktop       nur Läufe, deren Name das enthält (räumt out/shots nicht auf)
 *
 * Bilder:  out/shots/<rolle>-<seite>-<desktop|mobile>.png   (…-voll.png = ganze Seite, wenn sie länger als das Fenster ist)
 * Bericht: out/report.json + out/shots/<lauf>-console.log
 *
 * „Wie ein Mensch": Jede Seite wird über Klicks erreicht (Landing → Start →
 * Kachel → Karte → Zurück …), nicht per Adresse gesetzt. Ausnahme ist der
 * Deep-Link (`?uc=3`), der ist ein eigener Test der App.
 *
 * Ein Schritt, der scheitert, bricht seine Reise ab und steht als Befund im
 * Bericht — mit einem Bild vom Stand, an dem es hing (`…-FEHLER-…`). Ein
 * weißer Bildschirm (React-Baum abgestürzt) ist der schlimmste Fund und wird
 * nach jedem Bild geprüft.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const { start } = require('./serve');

const OUT = path.join(__dirname, 'out');
const SHOTS = path.join(OUT, 'shots');

/* ------------------------------------------------------------------ CLI --- */

const cli = {};
process.argv.slice(2).forEach(a => { const i = a.indexOf('='); if (i > 0) cli[a.slice(0, i)] = a.slice(i + 1); });
const liste = (k, alle) => (cli[k] ? cli[k].split(',').map(s => s.trim()).filter(Boolean) : alle);
const ROLLEN = liste('role', ['admin', 'organizer', 'user']);
const GERAETE = liste('device', ['desktop', 'mobile']);
const SPRACHE = cli.lang === 'en' ? 'en' : 'de';
const MIT_ZUSTAENDEN = cli.states !== '0';
const NUR = cli.only || '';

const exe = [process.env.PLAYWRIGHT_CHROMIUM, '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', '/opt/pw-browsers/chromium']
  .filter(Boolean).find(p => { try { return fs.statSync(p).isFile(); } catch { return false; } });

/* --------------------------------------------------------------- Texte ---- */

/* Beschriftungen der Oberfläche, die die Reise anklickt — in beiden Sprachen. */
const TEXTE = {
  de: {
    kUseCases: 'Use Cases', kStudio: 'Use Case Studio', kProtokoll: 'Protokoll', kRollen: 'Rollenverwaltung',
    zurueck: 'Zurück', nurLive: 'Nur aufrufbare', neu: 'Neuer Use Case', speichern: 'Speichern', abbrechen: 'Abbrechen',
    bearbeiten: 'Bearbeiten', loeschen: 'Löschen', verlauf: 'Verlauf', mehr: 'Betreuer, Bewertung, Aufruf und Reihenfolge',
    vergeben: 'Rolle vergeben', allesKlar: 'Alles klar', menue: 'Menü', fragen: 'Hast du Fragen?', ueber: 'Über die App',
    userAnsicht: 'User-Ansicht', userZurueck: 'User-Ansicht · zurück', organizerWerden: 'Organizer werden?',
    suche: 'Use Cases durchsuchen', entfernen: 'Entfernen', rechtePruefen: /^Rechte prüfen$/,
    neuLaden: 'Neu laden', bereich: 'Backoffice-Prozesse',
    tabPersonen: 'Personen', tabRechte: 'Rechte je Rolle', eintragen: /^Als .* eintragen$/, geprueft: /Einträge geprüft/,
    nachsetzen: /Fehlende Rechte nachsetzen/, jetztNachsetzen: 'Jetzt nachsetzen', entziehen: /Überzählige Rechte entziehen/, jetztEntziehen: 'Jetzt entziehen',
  },
  en: {
    kUseCases: 'Use Cases', kStudio: 'Use Case Studio', kProtokoll: 'Log', kRollen: 'Role management',
    zurueck: 'Back', nurLive: 'Only callable', neu: 'New use case', speichern: 'Save', abbrechen: 'Cancel',
    bearbeiten: 'Edit', loeschen: 'Delete', verlauf: 'History', mehr: 'Maintainers, assessment, launch and order',
    vergeben: 'Assign role', allesKlar: 'Got it', menue: 'Menu', fragen: 'Any questions?', ueber: 'About the app',
    userAnsicht: 'User view', userZurueck: 'User view · back', organizerWerden: 'Become an organizer?',
    suche: 'Search use cases', entfernen: 'Remove', rechtePruefen: /^Check rights$/,
    neuLaden: 'Reload', bereich: 'Backoffice-Prozesse',
    tabPersonen: 'People', tabRechte: 'Rights per role', eintragen: /^Add as /, geprueft: /entries checked/,
    nachsetzen: /Re-grant missing rights/, jetztNachsetzen: 'Re-grant now', entziehen: /Revoke surplus rights/, jetztEntziehen: 'Revoke now',
  },
};

const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/* Ein Profilbild, wie SharePoint es liefert (userphoto.aspx) — sonst zeigt jeder Lauf nur die Initialen. */
const AVATAR = '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 96 96"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#86bc25"/><stop offset="1" stop-color="#0076a8"/></linearGradient></defs><rect width="96" height="96" fill="url(#g)"/><circle cx="48" cy="38" r="17" fill="#fff" opacity=".92"/><path d="M14 96c2-22 16-32 34-32s32 10 34 32z" fill="#fff" opacity=".92"/></svg>';

/* ------------------------------------------------------------------ Lauf --- */

class Lauf {
  constructor(browser, base, spec, bericht) {
    this.browser = browser;
    this.base = base;
    this.spec = spec;
    this.t = TEXTE[spec.lang];
    this.rec = {
      id: spec.id, rolle: spec.role, geraet: spec.device, sprache: spec.lang, zustand: spec.state,
      shots: [], befunde: [], schritte: [], uebersprungen: [],
      konsole: { fehler: [], warnungen: [], pageerrors: [], anfragenGescheitert: [], harnessUnbehandelt: [] },
      requests: 0, dauerMs: 0, breiten: {},
    };
    bericht.laeufe.push(this.rec);
    this.abbruch = false;
    this.aktSchritt = '';
    this.ueberlaufGemeldet = new Set();
    this.optikGemeldet = new Set();
  }

  get mobil() { return this.spec.device === 'mobile'; }

  async oeffneKontext() {
    const m = this.mobil;
    this.ctx = await this.browser.newContext({
      viewport: m ? { width: 390, height: 844 } : { width: 1280, height: 900 },
      deviceScaleFactor: m ? 2 : 1,
      isMobile: m,
      hasTouch: m,
      locale: this.spec.lang === 'en' ? 'en-GB' : 'de-DE',
      timezoneId: 'Europe/Berlin',
      // „Link kopieren" schreibt in die Zwischenablage; ohne Freigabe zeigt die Seite den Ausweichweg.
      permissions: ['clipboard-read', 'clipboard-write'],
    });
    // Ohne SharePoint gibt es kein userphoto.aspx — der Harness antwortet selbst.
    await this.ctx.route('https://harness.local/**', route => {
      const url = route.request().url();
      if (!this.spec.ohneFoto && /userphoto\.aspx/.test(url)) return route.fulfill({ status: 200, contentType: 'image/svg+xml', body: AVATAR });
      return route.fulfill({ status: 404, body: '' });
    });
    this.page = await this.ctx.newPage();
    // Feste Uhrzeit: Die Begrüßung („Guten Morgen/Tag/Abend") hängt an der Stunde,
    // und die Bilder sollen an jedem Tag gleich aussehen. Timer laufen weiter.
    if (this.spec.uhr !== false) {
      try { await this.page.clock.setFixedTime(new Date('2026-09-29T13:00:00+02:00')); } catch { /* ältere Playwright-Version */ }
    }
    const k = this.rec.konsole;
    this.page.on('console', msg => {
      const txt = `[${this.aktSchritt || '-'}] ${msg.text().slice(0, 400)}`;
      if (msg.type() === 'error') k.fehler.push(txt);
      else if (msg.type() === 'warning') k.warnungen.push(txt);
    });
    this.page.on('pageerror', e => k.pageerrors.push(`[${this.aktSchritt || '-'}] ${String(e).slice(0, 600)}`));
    this.page.on('requestfailed', r => k.anfragenGescheitert.push(`[${this.aktSchritt || '-'}] ${r.failure() ? r.failure().errorText : 'failed'} ${r.url().slice(0, 160)}`));
  }

  url(extra) {
    const q = new URLSearchParams({ role: this.spec.role, lang: this.spec.lang, state: this.spec.state || 'ok' });
    if (this.mobil) q.set('mobile', '1');
    Object.keys(extra || {}).forEach(k => q.set(k, String(extra[k])));
    return `${this.base}/?${q.toString()}`;
  }

  async gehe(extra) {
    if (!this.ctx) await this.oeffneKontext();
    await this.sammleHarness();
    await this.page.goto(this.url(extra));
    await this.page.locator('.app-layout .header').waitFor({ timeout: 30000 });
  }

  /** Die Anfragen, die die Attrappe nicht kannte, vor jedem Neuladen einsammeln — danach sind sie weg. */
  async sammleHarness() {
    if (!this.page || this.page.url() === 'about:blank') return;
    try {
      const h = await this.page.evaluate(() => window.__harness ? { u: window.__harness.unhandled.slice(), n: window.__harness.requests.length } : null);
      if (h) {
        h.u.forEach(x => { if (this.rec.konsole.harnessUnbehandelt.indexOf(x) < 0) this.rec.konsole.harnessUnbehandelt.push(x); });
        this.rec.requests += h.n;
      }
    } catch { /* Seite gerade im Wechsel */ }
  }

  befund(seite, problem, evidenz) {
    this.rec.befunde.push({ seite: `${this.spec.role}/${this.spec.device}${this.spec.state && this.spec.state !== 'ok' ? '/' + this.spec.state : ''}/${seite}`, problem, evidenz: String(evidenz).slice(0, 700) });
  }

  /** Ein Schritt der Reise. Scheitert er, steht es im Bericht, und die Reise endet dort. */
  async schritt(name, fn, opt) {
    if (this.abbruch && !(opt && opt.trotzdem)) { this.rec.uebersprungen.push(name); return; }
    this.aktSchritt = name;
    try {
      await fn();
      this.rec.schritte.push({ name, ok: true });
    } catch (e) {
      const msg = String(e && e.message ? e.message : e).split('\n').slice(0, 3).join(' ').slice(0, 500);
      this.rec.schritte.push({ name, ok: false, fehler: msg });
      this.befund(name, `Schritt „${name}" ist gescheitert — die Reise bricht hier ab.`, msg);
      try {
        const datei = `${this.spec.role}-FEHLER-${name}-${this.spec.device}.png`;
        await this.page.screenshot({ path: path.join(SHOTS, datei) });
        this.rec.shots.push({ pfad: path.join(SHOTS, datei), was: `Stand beim Scheitern des Schritts „${name}".` });
      } catch { /* */ }
      this.abbruch = true;
    }
  }

  /** Nur ein Schritt, der eine Seite/Kachel voraussetzt, die es in dieser Rolle nicht gibt: still überspringen. */
  async vorhanden(locator) { try { return (await locator.count()) > 0; } catch { return false; } }

  /* ---- Prüfungen nach jedem Bild ---- */

  async pruefe(seite) {
    const s = await this.page.evaluate(() => {
      const lay = document.querySelector('.app-layout');
      const mc = document.querySelector('.main-content');
      return { layout: !!lay, header: !!document.querySelector('.app-layout .header'), text: mc ? (mc.innerText || '').trim().length : 0 };
    });
    if (!s.layout || !s.header) {
      this.befund(seite, 'WEISSER BILDSCHIRM: Die App-Hülle (.app-layout mit Kopfzeile) ist nicht mehr im DOM — der React-Baum ist abgestürzt.', `.app-layout ${s.layout ? 'da' : 'fehlt'}, Kopfzeile ${s.header ? 'da' : 'fehlt'}`);
      throw new Error('weißer Bildschirm');
    }
    if (s.text < 15) this.befund(seite, 'Der Seiteninhalt (.main-content) ist praktisch leer.', `${s.text} Zeichen sichtbarer Text`);

    // Zwei Aussehens-Prüfungen, die sich aus Bildern ergaben und sonst niemand nachzählt:
    const optik = await this.page.evaluate(() => {
      const sichtbar = el => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
      // 1) Zeilentitel und -untertitel als <span> ohne display:block: Die Texte laufen ineinander („…entfernterika.muster@…").
      const inline = [...document.querySelectorAll('.dex-ui-row-title, .dex-ui-row-sub')].filter(el => sichtbar(el) && getComputedStyle(el).display === 'inline');
      // 2) Ein <button> im Browser-Standard (2px outset) ist ein Knopf, dem seine Klasse fehlt.
      const nackt = [...document.querySelectorAll('button')].filter(b => sichtbar(b) && getComputedStyle(b).borderTopStyle === 'outset');
      const page = document.querySelector('.page-container');
      return {
        inline: inline.slice(0, 2).map(e => `${e.className}: „${(e.textContent || '').trim().slice(0, 40)}"`), inlineN: inline.length,
        nackt: nackt.slice(0, 2).map(b => `button.${b.className || '(ohne Klasse)'}: „${(b.textContent || '').trim().slice(0, 40)}"`), nacktN: nackt.length,
        breite: page ? Math.round(page.getBoundingClientRect().width) : 0,
      };
    });
    // Je Lauf einmal je Beispiel: dieselbe Zeile steht auf jedem Bild derselben Seite.
    if (optik.inlineN > 0 && !this.optikGemeldet.has('inline:' + optik.inline[0])) {
      this.optikGemeldet.add('inline:' + optik.inline[0]);
      this.befund(seite, `Zeilentitel/-untertitel (.dex-ui-row-title/-sub) sind inline (${optik.inlineN} Stück): Text und Nachbartext laufen ineinander.`, optik.inline.join(' | '));
    }
    if (optik.nacktN > 0 && !this.optikGemeldet.has('nackt:' + optik.nackt[0])) {
      this.optikGemeldet.add('nackt:' + optik.nackt[0]);
      this.befund(seite, `Ein Knopf erscheint im Browser-Standard-Aussehen (ohne Stil, ${optik.nacktN} Stück).`, optik.nackt.join(' | '));
    }
    this.rec.breiten[seite] = optik.breite;

    // Horizontaler Überlauf: ragt etwas rechts aus dem Fenster, ohne dass ein Vorfahr es beschneidet/scrollt?
    const ueber = await this.page.evaluate(() => {
      const w = window.innerWidth;
      const raus = [];
      const beschnitten = el => {
        for (let p = el.parentElement; p && p !== document.body; p = p.parentElement) {
          const cs = getComputedStyle(p);
          if (/(hidden|auto|scroll|clip)/.test(cs.overflowX) && p.getBoundingClientRect().right <= w + 1) return true;
        }
        return false;
      };
      document.querySelectorAll('.app-layout *').forEach(el => {
        if (raus.length >= 3) return;
        const cs = getComputedStyle(el);
        if (cs.position === 'fixed' || cs.visibility === 'hidden' || cs.display === 'none') return;
        const r = el.getBoundingClientRect();
        if (r.width > 0 && r.height > 0 && r.right > w + 1 && !beschnitten(el)) {
          raus.push(`${el.tagName.toLowerCase()}${el.className && typeof el.className === 'string' ? '.' + el.className.split(/\s+/).slice(0, 2).join('.') : ''} rechts ${Math.round(r.right)} > Fenster ${w}`);
        }
      });
      return raus;
    });
    if (ueber.length && !this.ueberlaufGemeldet.has(ueber[0])) {
      this.ueberlaufGemeldet.add(ueber[0]);
      this.befund(seite, 'Inhalt ragt rechts aus dem Fenster (horizontaler Überlauf).', ueber.join(' | '));
    }
  }

  /* ---- Bilder ---- */

  dateiname(seite, suffix) {
    return `${this.spec.role}-${seite}-${this.spec.device}${suffix || ''}${this.spec.lang === 'en' ? '-en' : ''}.png`;
  }

  async shot(seite, was, opt) {
    const o = opt || {};
    await this.page.waitForTimeout(o.warte === undefined ? 350 : o.warte);
    await this.pruefe(seite);
    const datei = this.dateiname(seite);
    await this.page.screenshot({ path: path.join(SHOTS, datei) });
    this.rec.shots.push({ pfad: path.join(SHOTS, datei), was });
    if (o.voll) await this.vollbild(seite, was);
  }

  /** Ist die Seite länger als das Fenster, zusätzlich ein Bild mit hohem Fenster: die App-Hülle folgt der Fensterhöhe (useShellHeight), so steht die ganze Seite im Bild. */
  async vollbild(seite, was) {
    const m = await this.page.evaluate(() => {
      const mc = document.querySelector('.main-content'); const lay = document.querySelector('.app-layout');
      return mc && lay ? { scroll: mc.scrollHeight, client: mc.clientHeight, layout: lay.clientHeight } : null;
    });
    if (!m || m.scroll <= m.client + 40) return;
    const vp = this.page.viewportSize();
    const hoehe = Math.min(m.scroll + (m.layout - m.client) + 8, this.mobil ? 4400 : 5600);
    await this.page.setViewportSize({ width: vp.width, height: hoehe });
    await this.page.waitForTimeout(500);
    const datei = this.dateiname(seite, '-voll');
    await this.page.screenshot({ path: path.join(SHOTS, datei) });
    this.rec.shots.push({ pfad: path.join(SHOTS, datei), was: `${was} (ganze Seite, ${m.scroll} px lang)` });
    await this.page.setViewportSize(vp);
    await this.page.waitForTimeout(350);
  }

  /* ---- Orientierung ---- */

  async wartLanding() {
    await this.page.locator('.landing h1').waitFor({ timeout: 20000 });
    await this.page.waitForFunction(() => { const l = document.querySelector('.landing'); return l && !/werden geladen|Loading the use cases/.test(l.innerText || ''); }, null, { timeout: 25000 });
  }

  async wartWand() {
    await this.page.locator('.page-container h1.dex-ui-page-head-title', { hasText: /^Use Cases$/ }).waitFor({ timeout: 20000 });
    await this.page.waitForFunction(() => !/werden geladen|Loading use cases/.test((document.querySelector('.page-container') || {}).innerText || ''), null, { timeout: 25000 });
  }

  async wartSeite() {
    await this.page.locator('.page-container h1').first().waitFor({ timeout: 20000 });
    await this.page.waitForFunction(() => !/wird geladen|is loading/.test((document.querySelector('.page-container') || {}).innerText || ''), null, { timeout: 25000 });
  }

  kachel(titel) {
    const rx = new RegExp('^' + esc(titel) + '$');
    return this.page.locator('.start-card, .dex-start-rows button').filter({ has: this.page.locator('h2, .dex-ui-row-title').filter({ hasText: rx }) });
  }

  karte(titel) { return this.page.getByRole('button', { name: new RegExp('^' + esc(titel) + ' — ') }); }

  async start() {
    await this.page.getByRole('button', { name: 'Start', exact: true }).click();
    await this.page.locator('.dex-cluster-grid, .dex-start-rows').waitFor({ timeout: 15000 });
  }

  async zurueckHeader() { await this.page.locator('.header .back-btn').click(); await this.page.waitForTimeout(300); }
  async zurueckSeite() { await this.page.locator('.page-container .dex-ui-textbtn', { hasText: this.t.zurueck }).first().click(); await this.page.waitForTimeout(300); }

  dialog() { return this.page.locator('[role="dialog"]').last(); }
  async wartDialog() { await this.dialog().waitFor({ timeout: 10000 }); await this.page.waitForTimeout(250); }

  /** Im offenen Dialog bis ganz nach unten blättern, damit man sieht, was hinter dem Falz steht. */
  async dialogNachUnten() {
    await this.dialog().evaluate(root => {
      [root, ...root.querySelectorAll('*')].forEach(el => {
        const cs = getComputedStyle(el);
        if (/(auto|scroll)/.test(cs.overflowY) && el.scrollHeight > el.clientHeight) el.scrollTop = el.scrollHeight;
      });
    });
    await this.page.waitForTimeout(250);
  }

  /** Die Use Cases der Attrappe — damit die Reise keine Titel festverdrahtet (die Startliste ändert sich) und die Erwartungen aus den Daten kommen. */
  async daten() {
    if (this._d) return this._d;
    await this.page.waitForFunction(() => window.__harness && window.__harness.lists.useCases.items.length > 0, null, { timeout: 30000 });
    this._d = await this.page.evaluate(() => {
      const it = window.__harness.lists.useCases.items.slice().sort((a, b) => a.Id - b.Id);
      const finde = f => { const z = it.filter(f)[0]; return z ? z.Id : 0; };
      const titel = {}; it.forEach(z => { titel[z.Id] = z.Title; });
      return {
        titel, total: it.length,
        archiviert: it.filter(z => z.UcStatus === 'Archiviert').length,
        erste: it[0].Id, drei: (it[2] || it[0]).Id, such: (it[3] || it[0]).Title, bereich: (it[9] || it[0]).Bereich || '',
        inArbeit: finde(z => z.UcStatus === 'InArbeit'),
        ohneLink: finde(z => z.UcStatus === 'Live' && !z.LinkDeployment),
        eingebettet: finde(z => z.AufrufArt === 'eingebettet'),
        archiv: finde(z => z.UcStatus === 'Archiviert'),
        ohneBild: finde(z => !z.BildUrl && z.UcStatus === 'Live'),
      };
    });
    return this._d;
  }

  /** Wartet, bis die Seite Text zeigt und nichts mehr lädt — ein leerer Ladekasten zählt nicht als Inhalt. */
  async wartInhalt() {
    await this.page.waitForFunction(() => {
      const t = ((document.querySelector('.page-container') || {}).innerText || '').trim();
      return t.length > 20 && !/wird geladen|werden geladen|is loading|Loading/.test(t);
    }, null, { timeout: 25000 });
  }

  /** Was die Attrappe nach einer Aktion tatsächlich hält — der Beleg, dass Schreiben WIRKLICH durchlief. */
  async harness(fn) { return this.page.evaluate(fn); }

  async ende() {
    await this.sammleHarness();
    const k = this.rec.konsole;
    k.pageerrors.forEach(e => this.befund('console', 'pageerror — der React-Baum oder ein Handler ist gestorben.', e));
    // Im Normalzustand ist jeder console.error unerwartet; in Fehlerzuständen melden die Seiten ihre Lesefehler selbst.
    if ((this.spec.state === 'ok' || !this.spec.state) && !this.spec.ohneFoto) k.fehler.forEach(e => this.befund('console', 'console.error im Normalzustand.', e));
    // Breite der Inhaltsspalte je Seite (nur am Rechner aussagekräftig): Sie soll nicht vom Inhalt abhängen.
    const br = this.rec.breiten;
    const namen = Object.keys(br).filter(n => /^(usecases|detail|studio|protokoll|rollen)$/.test(n) && br[n] > 0);
    if (!this.mobil && namen.length >= 3) {
      const w = namen.map(n => br[n]);
      if (Math.max.apply(null, w) - Math.min.apply(null, w) > 150) this.befund('layout', 'Die Breite der Inhaltsspalte hängt vom Inhalt ab: Die Seiten sind unterschiedlich breit.', namen.map(n => `${n} ${br[n]} px`).join(', '));
    }
    k.harnessUnbehandelt.forEach(u => this.befund('rest', 'Die App ruft SharePoint-REST auf, den die Attrappe nicht kennt (entweder fehlt er im Harness oder die App ruft etwas Falsches).', u));
    const log = [];
    Object.keys(k).forEach(key => { if (k[key].length) { log.push(`## ${key}`); k[key].forEach(x => log.push(x)); } });
    fs.writeFileSync(path.join(SHOTS, `${this.spec.id}-console.log`), log.join('\n'));
    try { if (this.ctx) await this.ctx.close(); } catch { /* */ }
  }
}

/* -------------------------------------------------------------- Reisen ---- */

const ROLLEN_TEXT = { admin: 'Admin', organizer: 'Use Case Organizer', user: 'User' };
const STARTBEZUG = {
  admin: 'Start-Übersicht mit vier Abschnitten (Entdecken, Organisation mit Studio und Protokoll, Support, Verwaltung mit Rollenverwaltung).',
  organizer: 'Start-Übersicht mit drei Abschnitten (Entdecken, Organisation mit Studio und Protokoll, Support) — ohne Rollenverwaltung.',
  user: 'Start-Übersicht mit zwei Abschnitten (Entdecken, Support); die Kachel „Use Case Studio" ist ausgegraut mit „Organizer werden?".',
};

/** Die lange Reise: Landing → Start → Wand → Detail → … je nach Rolle Studio, Protokoll, Rollen. */
async function reise(l) {
  const t = l.t; const R = l.spec.role; const istOrg = R === 'admin' || R === 'organizer';
  let D = null;
  const T = id => (D && D.titel[id]) || '';
  await l.gehe();

  await l.schritt('landing', async () => {
    await l.wartLanding();
    D = await l.daten();
    await l.shot('landing', `Landing Page als ${ROLLEN_TEXT[R]}: Orb, Begrüßung mit Vorname, Hinweiskarte mit Zahl der Use Cases, Start-Knopf${R === 'user' && !l.mobil ? ', Einladung „Eigenen Use Case einstellen"' : ''}.`);
  });

  await l.schritt('start', async () => {
    await l.start();
    await l.shot('start', `${STARTBEZUG[R]} Rechner: große Kacheln, Handy: Zeilen.`, { voll: true });
  });

  await l.schritt('usecases', async () => {
    await l.kachel(t.kUseCases).click();
    await l.wartWand();
    const n = await l.page.getByRole('button', { name: / — / }).count();
    const soll = R === 'user' ? D.total - D.archiviert : D.total;
    if (n !== soll) l.befund('usecases', `Die Wand zeigt ${n} Kacheln, erwartet ${soll} (${R === 'user' ? 'archivierte sind für User unsichtbar' : 'Organizer/Admin sehen auch archivierte'}).`, `${n} Buttons mit „ — " im Namen`);
    await l.shot('usecases', `Kachelwand als ${ROLLEN_TEXT[R]}: ${soll} Use Cases als Kacheln mit Bild oder Kürzel, Status, Bewertung; Filterleiste (Bereich, „${t.nurLive}").`, { voll: true });
  });

  await l.schritt('filter', async () => {
    await l.page.getByRole('button', { name: t.nurLive }).click();
    await l.page.waitForTimeout(250);
    await l.shot('usecases-aufrufbar', 'Kachelwand mit gesetztem Filter „Nur aufrufbare": nur Use Cases mit Status Live.');
    await l.page.getByRole('button', { name: t.nurLive }).click();
    await l.page.locator('#uc-bereich').selectOption({ label: D.bereich });
    await l.page.waitForTimeout(250);
    await l.shot('usecases-bereich', `Kachelwand gefiltert auf den Bereich „${D.bereich}".`);
    await l.page.locator('#uc-bereich').selectOption({ value: '' });
  });

  await l.schritt('detail', async () => {
    await l.karte(T(D.erste)).click();
    await l.page.getByRole('heading', { level: 1, name: T(D.erste) }).waitFor({ timeout: 10000 });
    await l.shot('detail', 'Detailseite eines Live-Use-Cases mit Kachelbild: Kopf, Start-Knopf, Beschreibung, Bewertung, Ressourcen.' + (istOrg ? ' Organizer sehen zusätzlich „Verlauf" und „Bearbeiten".' : ''), { voll: true });
    // „Link kopieren": Die Zwischenablage ist freigegeben — der Text muss der Deep-Link dieses Use Cases sein.
    await l.page.getByRole('button', { name: /Link kopieren|Copy link/ }).click();
    await l.page.getByText(/Link kopiert|Link copied/).waitFor({ timeout: 5000 });
    const kopiert = await l.page.evaluate(() => navigator.clipboard.readText());
    if (!new RegExp('[?&]uc=' + D.erste + '$').test(kopiert)) l.befund('detail', 'Der kopierte Link führt nicht auf diesen Use Case.', `Zwischenablage: ${kopiert}`);
    await l.shot('detail-link-kopiert', 'Detailseite nach „Link kopieren": Knopf zeigt „Link kopiert" (der Link steht in der Zwischenablage).');
    await l.zurueckSeite();
    await l.wartWand();
  });

  await l.schritt('detail-varianten', async () => {
    const faelle = [
      [D.inArbeit, 'detail-inarbeit', 'Detailseite eines Use Cases „In Arbeit": statt Start-Knopf der Hinweis, dass die Demo gebaut wird.'],
      [D.ohneLink, 'detail-ohne-link', 'Detailseite eines Use Cases auf „Live" ohne Deployment-Link: Hinweis auf den Pflegefehler statt Start-Knopf.'],
      [D.eingebettet, 'detail-eingebettet', 'Detailseite eines Use Cases mit Aufruf „eingebettet": Knopf „Demo hier öffnen" (nicht geklickt — das iframe würde ins Netz laden).'],
    ];
    for (const [id, seite, was] of faelle) {
      if (!id) { l.rec.uebersprungen.push(`${seite} (die Daten enthalten diesen Fall nicht)`); continue; }
      await l.karte(T(id)).click();
      await l.page.getByRole('heading', { level: 1, name: T(id) }).waitFor({ timeout: 10000 });
      await l.shot(seite, was);
      await l.zurueckSeite();
      await l.wartWand();
    }
    // Archivierte Use Cases sehen nur Organizer und Admins.
    if (D.archiv) {
      const archiv = l.karte(T(D.archiv));
      if (istOrg) {
        await archiv.click();
        await l.page.getByRole('heading', { level: 1, name: T(D.archiv) }).waitFor({ timeout: 10000 });
        await l.shot('detail-archiviert', 'Detailseite eines archivierten Use Cases (nur für Organizer/Admin erreichbar): Hinweis „archiviert".');
        await l.zurueckSeite();
        await l.wartWand();
      } else if ((await archiv.count()) > 0) {
        l.befund('usecases', `Ein archivierter Use Case („${T(D.archiv)}") steht für die Rolle User auf der Wand.`, 'Karte vorhanden');
      }
    }
  });

  await l.schritt('zurueck-start', async () => {
    await l.zurueckHeader();
    await l.page.locator('.dex-cluster-grid, .dex-start-rows').waitFor({ timeout: 10000 });
  });

  if (istOrg) {
    await l.schritt('studio', async () => {
      await l.kachel(t.kStudio).click();
      await l.page.locator('.page-container h1', { hasText: 'Use Case Studio' }).waitFor({ timeout: 10000 });
      await l.shot('studio', 'Use Case Studio: Liste aller Use Cases mit Status, Link-Hinweis, Bearbeiten/Löschen; oben „Protokoll" und „Neuer Use Case".', { voll: true });
    });

    await l.schritt('studio-neu', async () => {
      await l.page.getByRole('button', { name: t.neu }).first().click();
      await l.wartDialog();
      await l.shot('studio-dialog-neu', 'Use Case Studio, Dialog „Neuer Use Case" offen: Pflichtteil (Titel, Kurzbeschreibung, Bereich, Status, Kachelbild) und Ressourcen.');
      await l.dialog().getByRole('button', { name: t.mehr }).click();
      await l.dialogNachUnten();
      await l.shot('studio-dialog-neu-mehr', 'Dialog „Neuer Use Case" bis ganz unten geblättert, Aufklapper „Betreuer, Bewertung, Aufruf, Reihenfolge" offen.');
      await l.dialog().evaluate(root => { [root, ...root.querySelectorAll('*')].forEach(el => { el.scrollTop = 0; }); });
    });

    await l.schritt('studio-speichern', async () => {
      const d = l.dialog();
      await d.locator('#uc-titel').fill('Harness-Testdemo');
      await d.locator('#uc-kurz').fill('Nur im Harness angelegt, um den Schreibweg zu prüfen.');
      await d.locator('#uc-bereich').fill('Harness');
      await d.getByRole('button', { name: /^Live/ }).click();
      await d.locator('#uc-deploy').fill('https://example.com/harness-demo');
      await l.shot('studio-dialog-ausgefuellt', 'Dialog „Neuer Use Case" ausgefüllt (Titel, Kurzbeschreibung, Bereich, Status Live, Deployment-Link), vor dem Speichern.');
      await d.getByRole('button', { name: t.speichern }).click();
      await l.page.locator('[role="dialog"]').waitFor({ state: 'detached', timeout: 15000 });
      await l.page.locator('.dex-ui-row', { hasText: 'Harness-Testdemo' }).waitFor({ timeout: 10000 });
      const h = await l.harness(() => ({
        uc: window.__harness.lists.useCases.items.filter(z => z.Title === 'Harness-Testdemo').length,
        log: window.__harness.lists.log.items.filter(z => z.Aktion === 'angelegt' && z.Detail === 'Harness-Testdemo').length,
      }));
      if (h.uc !== 1) l.befund('studio-speichern', 'Nach „Speichern" liegt die Zeile nicht genau einmal in der Liste.', JSON.stringify(h));
      if (h.log !== 1) l.befund('studio-speichern', 'Nach dem Anlegen steht kein Protokolleintrag „angelegt".', JSON.stringify(h));
      await l.shot('studio-nach-speichern', 'Use Case Studio nach dem Anlegen: „Harness-Testdemo" steht in der Liste.', { voll: true });
    });

    await l.schritt('studio-bearbeiten', async () => {
      await l.page.locator('.dex-ui-row', { hasText: T(D.erste) }).getByRole('button', { name: t.bearbeiten }).click();
      await l.wartDialog();
      await l.shot('studio-dialog-bearbeiten', 'Use Case Studio, Dialog „Use Case bearbeiten" für den ersten Use Case: alle Felder mit den gespeicherten Werten, Kachelbild sichtbar.');
      await l.dialog().getByRole('button', { name: t.abbrechen }).click();
      await l.page.locator('[role="dialog"]').waitFor({ state: 'detached', timeout: 10000 });
    });

    await l.schritt('studio-loeschen', async () => {
      await l.page.locator('.dex-ui-row', { hasText: 'Harness-Testdemo' }).getByRole('button', { name: t.loeschen }).click();
      await l.wartDialog();
      await l.shot('studio-loeschen-bestaetigen', 'Bestätigungsdialog beim Löschen eines Use Cases (Hinweis auf den Papierkorb).');
      await l.dialog().getByRole('button', { name: t.loeschen }).click();
      await l.page.locator('.dex-ui-row', { hasText: 'Harness-Testdemo' }).waitFor({ state: 'detached', timeout: 10000 });
      const n = await l.harness(() => window.__harness.lists.useCases.items.filter(z => z.Title === 'Harness-Testdemo').length);
      if (n !== 0) l.befund('studio-loeschen', 'Der Use Case steht nach dem Löschen noch in der Liste.', `${n} Zeilen`);
    });

    await l.schritt('protokoll', async () => {
      await l.page.locator('.page-container .dex-ui-textbtn', { hasText: t.kProtokoll }).first().click();
      await l.wartSeite();
      const protokoll = await l.page.locator('.page-container').innerText();
      if (!/Harness-Testdemo/.test(protokoll)) l.befund('protokoll', 'Das Protokoll zeigt die eben ausgeführten Aktionen („Harness-Testdemo" angelegt und gelöscht) nicht.', protokoll.slice(0, 200).replace(/\n/g, ' | '));
      await l.shot('protokoll', 'Protokoll: Wer hat wann was geändert — mit den Einträgen dieser Sitzung („angelegt", „gelöscht") ganz oben.', { voll: true });
    });

    await l.schritt('protokoll-usecase', async () => {
      await l.zurueckSeite();   // → Studio
      await l.zurueckHeader();  // → Start
      await l.kachel(t.kUseCases).click();
      await l.wartWand();
      await l.karte(T(D.erste)).click();
      await l.page.getByRole('heading', { level: 1, name: T(D.erste) }).waitFor({ timeout: 10000 });
      await l.shot('detail-organizer', 'Detailseite als Organizer/Admin: Kopfzeile der Seite mit „Link kopieren", „Verlauf" und „Bearbeiten".');
      await l.page.getByRole('button', { name: t.verlauf }).click();
      await l.wartSeite();
      await l.shot('protokoll-usecase', 'Protokoll, auf einen Use Case eingeengt (über „Verlauf" auf der Detailseite): Chip „Nur …" oben.', { voll: true });
      await l.zurueckSeite();  // → Detail
      await l.zurueckSeite();  // → Wand
      await l.zurueckHeader(); // → Start
      await l.page.locator('.dex-cluster-grid, .dex-start-rows').waitFor({ timeout: 10000 });
    });

    await l.schritt('user-ansicht', async () => {
      await l.page.getByRole('button', { name: t.menue }).click();
      await l.page.getByRole('menuitemradio', { name: new RegExp(esc(t.userAnsicht)) }).click();
      await l.page.locator('.header-right .dex-ui-chip').waitFor({ timeout: 5000 });
      await l.shot('user-ansicht', 'Start-Übersicht in der „User-Ansicht" (Vorschau als normaler Nutzer): Chip „User-Ansicht · zurück" in der Kopfzeile, Studio-Kachel ausgegraut.');
      await l.page.locator('.header-right .dex-ui-chip').click();
      await l.page.locator('.header-right .dex-ui-chip').waitFor({ state: 'detached', timeout: 5000 });
    });
  }

  if (R === 'admin') {
    await l.schritt('rollen', async () => {
      await l.kachel(t.kRollen).click();
      await l.wartSeite();
      await l.page.getByRole('tab', { name: t.tabPersonen }).waitFor({ timeout: 10000 });
      await l.shot('rollen', 'Rollenverwaltung, Reiter „Personen": Suche nach einer Person, darunter Kennzahlen, Filter und die Tabelle der vergebenen Rollen (Admin, Use Case Organizer, User) mit Spalte „SharePoint-Rechte".', { voll: true });
    });

    await l.schritt('rollen-suche', async () => {
      await l.page.locator('#rolle-suche').click();
      await l.page.keyboard.type('wagner', { delay: 40 });
      await l.page.locator('.dex-ui-row--framed.dex-ui-rowbtn').first().waitFor({ timeout: 15000 });
      await l.shot('rollen-suche', 'Rollenverwaltung: Personensuche mit „wagner" — Trefferliste (Name, Adresse, Position) aus dem People-Picker.');
    });

    await l.schritt('rollen-vergeben', async () => {
      await l.page.locator('.dex-ui-row--framed.dex-ui-rowbtn', { hasText: 'Wagner, Sophie' }).click();
      const eintragen = l.page.getByRole('button', { name: t.eintragen });
      await eintragen.waitFor({ timeout: 10000 });
      await l.shot('rollen-person-gewaehlt', 'Rollenverwaltung: Person aus der Suche gewählt — Rollenauswahl als Kacheln und Knopf „Als Use Case Organizer eintragen".');
      await eintragen.click();
      const banner = l.page.locator('.page-container .dex-ui-callout[role]').first();
      await banner.waitFor({ timeout: 45000 });
      const text = (await banner.innerText()).replace(/\n/g, ' ');
      const ton = await banner.getAttribute('class');
      if (!/callout--success/.test(ton || '')) l.befund('rollen-vergeben', 'Die Rollenvergabe meldet keinen Erfolg.', `${ton} — ${text.slice(0, 250)}`);
      const h = await l.harness(() => {
        const z = window.__harness.lists.roles.items.filter(r => r.Title === 'sophie.wagner@deloitte.de')[0];
        const pr = window.__harness.prinzipale.filter(p => p.Email === 'sophie.wagner@deloitte.de')[0];
        const rechte = pr ? ['roles', 'useCases', 'log'].map(k => `${k}:${Array.from(window.__harness.lists[k].assign.get(pr.Id) || []).join('+')}`).join(' ') : 'kein Konto';
        return { rolle: z ? z.Role : null, rechte };
      });
      if (h.rolle !== 'Kurator') l.befund('rollen-vergeben', 'Die vergebene Rolle „Organizer" steht nicht als „Kurator" in der Rollenliste.', JSON.stringify(h));
      if (!/roles:1073741826/.test(h.rechte) || !/useCases:1073741830/.test(h.rechte) || !/log:1073741827/.test(h.rechte)) l.befund('rollen-vergeben', 'Nach der Rollenvergabe fehlen Rechte auf den Listen (erwartet: Rollenliste Read, Use Cases Edit, Protokoll Contribute).', h.rechte);
      await l.shot('rollen-vergeben', 'Rollenverwaltung nach „Als Use Case Organizer eintragen": grüne Rückmeldung, die Person steht in der Tabelle.', { voll: true });
    });

    await l.schritt('rollen-entfernen', async () => {
      await l.page.locator('tr', { hasText: 'Testmann' }).getByRole('button', { name: /Rolle entfernen|Remove role/ }).click();
      await l.wartDialog();
      await l.shot('rollen-entfernen-bestaetigen', 'Rollenverwaltung: Bestätigung beim Entfernen einer Rolle (die Person verliert auch die Rechte auf den Listen).');
      await l.dialog().getByRole('button', { name: t.abbrechen }).click();
      await l.page.locator('[role="dialog"]').waitFor({ state: 'detached', timeout: 5000 });
    });

    await l.schritt('rollen-rechte', async () => {
      await l.page.getByRole('button', { name: t.rechtePruefen }).click();
      await l.page.getByText(t.geprueft).waitFor({ timeout: 60000 });
      await l.shot('rollen-rechte-pruefen', 'Rollenverwaltung nach „Rechte prüfen": Bericht über Lücken (Tim Testmann, ausgeschiedene Person), Überschuss (Sabine Schulz) und Alias-Schreibweise (j.klein@) — Spalte „SharePoint-Rechte" gefüllt.', { warte: 600, voll: true });
      const nachsetzen = l.page.getByRole('button', { name: t.nachsetzen });
      if (await l.vorhanden(nachsetzen)) {
        await nachsetzen.click();
        await l.wartDialog();
        await l.shot('rollen-nachsetzen-bestaetigen', 'Rollenverwaltung: Bestätigung „Fehlende Rechte nachsetzen".');
        await l.dialog().getByRole('button', { name: t.jetztNachsetzen }).click();
        await l.page.locator('[role="dialog"]').waitFor({ state: 'detached', timeout: 5000 });
        await l.page.locator('.page-container .dex-ui-callout--success, .page-container .dex-ui-callout--warn').filter({ hasText: /nachgesetzt|re-granted|Nachsetzen|Re-grant|Rechte/i }).last().waitFor({ timeout: 60000 });
        await l.shot('rollen-rechte-nachgesetzt', 'Rollenverwaltung nach „Fehlende Rechte nachsetzen": Ergebnis (behoben / weiterhin offen, z. B. ausgeschiedene Person).', { warte: 600, voll: true });
      } else {
        l.befund('rollen-rechte', 'Der Bericht meldet keine Lücken, obwohl die Attrappe eine Person mit fehlenden Rechten (Tim Testmann) führt.', 'kein Knopf „Fehlende Rechte nachsetzen"');
      }
      const entziehen = l.page.getByRole('button', { name: t.entziehen });
      if (await l.vorhanden(entziehen)) {
        await entziehen.click();
        await l.wartDialog();
        await l.dialog().getByRole('button', { name: t.jetztEntziehen }).click();
        await l.page.locator('[role="dialog"]').waitFor({ state: 'detached', timeout: 5000 });
        await l.page.waitForFunction(() => !document.querySelector('.dex-ui-progress'), null, { timeout: 60000 }).catch(() => undefined);
        await l.shot('rollen-rechte-entzogen', 'Rollenverwaltung nach „Überzählige Rechte entziehen": Ergebnis für Sabine Schulz (User mit Rest-Rechten).', { warte: 600, voll: true });
      }
    });

    await l.schritt('rollen-matrix', async () => {
      await l.page.getByRole('tab', { name: t.tabRechte }).click();
      await l.page.waitForTimeout(400);
      await l.shot('rollen-matrix', 'Rollenverwaltung, Reiter „Rechte je Rolle": Matrix, was Admin, Use Case Organizer und User dürfen (nur Ansicht), die eigene Rolle hervorgehoben.', { voll: true });
    });
  }

  if (R === 'user') {
    await l.schritt('organizer-werden', async () => {
      await l.kachel(t.kStudio).getByText(t.organizerWerden).click();
      await l.wartDialog();
      await l.shot('organizer-werden', 'Als User auf „Organizer werden?" geklickt: Dialog „Use Case Organizer werden" mit Vorlage für die Mail an die Plattform-Ansprechperson.');
      await l.dialog().getByRole('button', { name: t.abbrechen }).click();
      await l.page.locator('[role="dialog"]').waitFor({ state: 'detached', timeout: 5000 });
    });
  }
}

/** Kopfzeile und Dialoge: Suche, Menü, Profil, „Hast du Fragen?", „Über die App", Sprache. */
async function kopf(l) {
  const t = l.t; const R = l.spec.role;
  let D = null;
  await l.gehe();
  await l.schritt('start', async () => { await l.wartLanding(); D = await l.daten(); await l.start(); });
  // Ein Suchwort, das es sicher gibt: das letzte Wort eines Titels, sechs Buchstaben.
  const wort = () => D.such.split(/[\s-]+/).pop().slice(0, 6).toLowerCase();

  await l.schritt('suche', async () => {
    if (l.mobil) {
      // Auf dem Handy hat die Kopfzeile keinen Platz für die Suche — sie steht auf der Kachelwand.
      await l.kachel(t.kUseCases).click();
      await l.wartWand();
      await l.page.locator('#uc-suche').click();
      await l.page.keyboard.type(wort(), { delay: 40 });
      await l.page.waitForTimeout(300);
      await l.shot('suche', `Handy: Kachelwand mit Suchfeld (steht dort statt in der Kopfzeile), „${wort()}" eingetippt, Wand auf die Treffer gefiltert.`);
      await l.page.locator('#uc-suche').fill('');
      await l.zurueckHeader();
    } else {
      const feld = l.page.getByRole('textbox', { name: t.suche });
      await feld.click();
      await l.page.keyboard.type(wort(), { delay: 40 });
      await l.page.locator('[role="listbox"]').waitFor({ timeout: 5000 });
      await l.shot('suche', `Kopfzeilen-Suche: „${wort()}" eingetippt, Ergebnisfenster mit den besten Treffern (Titel, Bereich, Status).`);
      await l.page.keyboard.press('Enter');
      await l.wartWand();
      await l.shot('suche-wand', `Nach Enter: Kachelwand auf „${wort()}" gefiltert, Chip „Suche: …" zeigt, wonach gefiltert ist.`);
      await feld.fill('');
      await l.zurueckHeader();
    }
    await l.page.locator('.dex-cluster-grid, .dex-start-rows').waitFor({ timeout: 10000 });
  });

  await l.schritt('menue', async () => {
    await l.page.getByRole('button', { name: t.menue }).click();
    await l.page.getByRole('menu').waitFor({ timeout: 5000 });
    await l.shot('menue', `Burger-Menü offen: „Über die App"${R === 'user' ? '' : ' und der Abschnitt „Ansicht" (Organizer-/User-Ansicht)'}.`);
    await l.page.keyboard.press('Escape');
    await l.page.getByRole('menu').waitFor({ state: 'detached', timeout: 5000 });
  });

  await l.schritt('profil', async () => {
    await l.page.locator('.header-right img[alt], .header-right .header-avatar').last().click();
    await l.page.getByText('Page-ID:').waitFor({ timeout: 5000 });
    await l.shot('profil', `Profil-Popup: Foto, Name, Adresse, Rollen-Chip „${ROLLEN_TEXT[R]}" und die Page-ID.`);
    await l.page.keyboard.press('Escape');
    await l.page.getByText('Page-ID:').waitFor({ state: 'detached', timeout: 5000 });
  });

  await l.schritt('fragen', async () => {
    await l.page.getByRole('button', { name: t.fragen }).first().click();
    await l.wartDialog();
    await l.dialog().locator('textarea').fill(`Wie starte ich die Demo „${D.titel[D.erste]}" für einen Kundentermin?`);
    await l.shot('fragen', 'Dialog „Hast du Fragen?": Anlass zur Auswahl, Textfeld ausgefüllt, „E-Mail schreiben" und Ausweichadresse.');
    await l.dialog().getByRole('button', { name: t.abbrechen }).click();
    await l.page.locator('[role="dialog"]').waitFor({ state: 'detached', timeout: 5000 });
  });

  await l.schritt('ueber', async () => {
    await l.page.getByRole('button', { name: t.menue }).click();
    await l.page.getByRole('menuitem', { name: new RegExp(esc(t.ueber)) }).click();
    await l.wartDialog();
    await l.shot('ueber', `Dialog „Über die App": wofür die Plattform da ist, drei Schritte, wer was darf (Rolle „${ROLLEN_TEXT[R]}" hervorgehoben).`);
    await l.dialogNachUnten();
    await l.shot('ueber-unten', 'Dialog „Über die App" nach unten geblättert.');
    await l.page.keyboard.press('Escape');
    await l.page.locator('[role="dialog"]').waitFor({ state: 'detached', timeout: 5000 });
  });

  await l.schritt('sprache', async () => {
    await l.page.getByRole('group').getByRole('button', { name: 'EN', exact: true }).click();
    await l.page.waitForTimeout(400);
    await l.shot('start-en', 'Start-Übersicht nach Klick auf „EN" in der Kopfzeile: die ganze Oberfläche in Englisch.');
    await l.page.getByRole('group').getByRole('button', { name: 'DE', exact: true }).click();
  });
}

/** Kachelbild: hochladen, zuschneiden, speichern, wieder entfernen — der Schreibweg mit den meisten Schritten (Anhang an der Zeile, Reihenfolge „erst hochladen, dann speichern, dann altes Bild entfernen"). */
async function bild(l) {
  const t = l.t;
  let D = null;
  const zeile = () => l.page.locator('.dex-ui-row', { hasText: D.titel[D.ohneBild] });
  const oeffneBearbeiten = async () => {
    await zeile().getByRole('button', { name: t.bearbeiten }).click();
    await l.wartDialog();
  };
  await l.gehe();
  await l.schritt('bild-studio', async () => {
    await l.wartLanding();
    D = await l.daten();
    if (!D.ohneBild) throw new Error('kein Use Case ohne Bild in den Daten');
    await l.start();
    await l.kachel(t.kStudio).click();
    await l.page.locator('.page-container h1', { hasText: 'Use Case Studio' }).waitFor({ timeout: 10000 });
  });

  await l.schritt('bild-zuschnitt', async () => {
    await oeffneBearbeiten();
    // Ein Bild im Seitenverhältnis 4:3 — die Kachel ist 16:9, der Zuschnitt hat also etwas zu tun.
    const b64 = await l.page.evaluate(() => {
      const c = document.createElement('canvas'); c.width = 800; c.height = 600;
      const x = c.getContext('2d');
      const g = x.createLinearGradient(0, 0, 800, 600); g.addColorStop(0, '#0076a8'); g.addColorStop(1, '#86bc25');
      x.fillStyle = g; x.fillRect(0, 0, 800, 600);
      x.strokeStyle = 'rgba(255,255,255,.7)'; x.lineWidth = 6; x.strokeRect(20, 20, 760, 560);
      x.fillStyle = '#fff'; x.font = 'bold 96px Arial'; x.fillText('HARNESS', 150, 330);
      return c.toDataURL('image/png').split(',')[1];
    });
    await l.dialog().locator('input[type="file"]').setInputFiles({ name: 'kachel.png', mimeType: 'image/png', buffer: Buffer.from(b64, 'base64') });
    await l.page.locator('[role="dialog"]').last().getByRole('button', { name: /Übernehmen|Apply/ }).waitFor({ timeout: 10000 });
    await l.shot('studio-zuschnitt', 'Kachelbild gewählt: Dialog „Bild zuschneiden" über dem Bearbeiten-Dialog — Vorschau im Format der Kachel (16:9), Regler zum Zoomen.');
    await l.page.locator('[role="dialog"]').last().getByRole('button', { name: /Übernehmen|Apply/ }).click();
    await l.page.locator('[role="dialog"]').nth(1).waitFor({ state: 'detached', timeout: 10000 });
    await l.page.waitForTimeout(300);
    await l.shot('studio-bild-gewaehlt', 'Bearbeiten-Dialog nach „Übernehmen": die Vorschau zeigt den Zuschnitt, darunter der Hinweis „Das neue Bild wird beim Speichern hochgeladen".');
  });

  await l.schritt('bild-speichern', async () => {
    await l.dialog().getByRole('button', { name: t.speichern }).click();
    await l.page.locator('[role="dialog"]').waitFor({ state: 'detached', timeout: 20000 });
    await l.page.waitForTimeout(500);
    const h = await l.page.evaluate(id => {
      const z = window.__harness.lists.useCases.items.filter(x => x.Id === id)[0];
      return { url: z ? z.BildUrl : null, anhaenge: (window.__harness.lists.useCases.anhaenge[id] || []).map(a => a.FileName) };
    }, D.ohneBild);
    if (!h.url || h.anhaenge.length !== 1) l.befund('bild-speichern', 'Nach dem Speichern mit neuem Bild stimmt Zeile oder Anhang nicht (erwartet: BildUrl gesetzt, genau ein Anhang).', JSON.stringify(h));
    else {
      const geladen = await l.page.evaluate(async u => { try { const r = await fetch(u); return { ok: r.ok, typ: r.headers.get('content-type') }; } catch (e) { return { ok: false, typ: String(e) }; } }, h.url);
      if (!geladen.ok || !/^image\//.test(geladen.typ || '')) l.befund('bild-speichern', 'Die gespeicherte Bild-Adresse liefert kein Bild.', `${h.url} → ${JSON.stringify(geladen)}`);
    }
    await l.shot('studio-nach-bild', 'Use Case Studio nach dem Speichern mit neuem Kachelbild.');
    // Auf der Kachelwand muss die Kachel das Bild tragen.
    await l.zurueckHeader();
    await l.kachel(t.kUseCases).click();
    await l.wartWand();
    const karte = l.karte(D.titel[D.ohneBild]);
    await karte.scrollIntoViewIfNeeded();
    const hat = await karte.locator('span').first().evaluate(el => /url\(/.test(el.style.background || getComputedStyle(el).backgroundImage));
    if (!hat) l.befund('bild-speichern', 'Die Kachel trägt nach dem Speichern kein Hintergrundbild.', 'style.background ohne url(...)');
    await l.shot('usecases-neues-bild', 'Kachelwand, zur Kachel mit dem eben hochgeladenen Bild gescrollt.');
  });

  await l.schritt('bild-entfernen', async () => {
    await l.zurueckHeader();
    await l.kachel(t.kStudio).click();
    await l.page.locator('.page-container h1', { hasText: 'Use Case Studio' }).waitFor({ timeout: 10000 });
    await oeffneBearbeiten();
    await l.dialog().getByRole('button', { name: /Bild entfernen|Remove image/ }).click();
    await l.dialog().getByRole('button', { name: t.speichern }).click();
    await l.page.locator('[role="dialog"]').waitFor({ state: 'detached', timeout: 20000 });
    await l.page.waitForTimeout(500);
    const h = await l.page.evaluate(id => {
      const z = window.__harness.lists.useCases.items.filter(x => x.Id === id)[0];
      return { url: z ? z.BildUrl : null, anhaenge: (window.__harness.lists.useCases.anhaenge[id] || []).length };
    }, D.ohneBild);
    if (h.url || h.anhaenge !== 0) l.befund('bild-entfernen', 'Nach „Bild entfernen" und Speichern stehen Adresse oder Anhang noch.', JSON.stringify(h));
  });
}

/** Deep-Links — eine echte Prüfung der App: sie liest `?uc=` selbst beim Start. */
async function deeplink(l) {
  const t = l.t; const R = l.spec.role;
  let D = null;
  // Die Id 3 ist stabil (Id = Reihenfolge der Startliste); den Titel dazu liest der Lauf aus den Daten.
  await l.schritt('deeplink', async () => {
    await l.gehe({ uc: 3 });
    D = await l.daten();
    await l.page.getByRole('heading', { level: 1, name: D.titel[3] }).waitFor({ timeout: 15000 });
    await l.shot('deeplink', `Deep-Link ?uc=3: die App startet direkt auf der Detailseite von „${D.titel[3]}" (ohne Landing Page).`, { voll: true });
    // Ohne Verlauf gibt „Zurück" nichts her — die App muss dann eine Ebene höher gehen, statt tot zu sein.
    await l.zurueckSeite();
    await l.page.locator('.dex-cluster-grid, .dex-start-rows, .page-container h1.dex-ui-page-head-title').first().waitFor({ timeout: 10000 });
    const ziel1 = (await l.page.locator('.dex-cluster-grid, .dex-start-rows').count()) > 0 ? 'Start-Übersicht' : 'Kachelwand';
    await l.shot('deeplink-zurueck', `Deep-Link, dann „Zurück" IN der Seite: ${ziel1} (es gibt keinen Verlauf, die App geht eine Ebene höher).`);
    // Dasselbe mit dem „Zurück" der Kopfzeile.
    await l.gehe({ uc: 3 });
    await l.page.getByRole('heading', { level: 1, name: D.titel[3] }).waitFor({ timeout: 15000 });
    await l.zurueckHeader();
    await l.page.locator('.dex-cluster-grid, .dex-start-rows, .page-container h1.dex-ui-page-head-title').first().waitFor({ timeout: 10000 });
    const ziel2 = (await l.page.locator('.dex-cluster-grid, .dex-start-rows').count()) > 0 ? 'Start-Übersicht' : 'Kachelwand';
    if (ziel1 !== ziel2) l.befund('deeplink', 'Auf einer per Deep-Link geöffneten Detailseite führen die beiden „Zurück"-Knöpfe an verschiedene Orte.', `„Zurück" in der Seite → ${ziel1}; „Zurück" in der Kopfzeile → ${ziel2}`);
  }, { trotzdem: true });

  await l.schritt('deeplink-unbekannt', async () => {
    await l.gehe({ uc: 9999 });
    await l.page.locator('.page-container .dex-ui-empty-title').waitFor({ timeout: 15000 });
    await l.shot('deeplink-unbekannt', 'Deep-Link auf eine Nummer, die es nicht gibt (?uc=9999): Leerzustand „nicht (mehr) da" mit Weg zur Übersicht.');
  }, { trotzdem: true });

  await l.schritt('deeplink-archiviert', async () => {
    if (!D || !D.archiv) { l.rec.uebersprungen.push('deeplink-archiviert (kein archivierter Use Case in den Daten)'); return; }
    await l.gehe({ uc: D.archiv });
    if (R === 'user') {
      await l.page.locator('.page-container .dex-ui-empty-title').waitFor({ timeout: 15000 });
      await l.shot('deeplink-archiviert', 'Deep-Link auf einen archivierten Use Case als User: Leerzustand — archivierte sind nur für Organizer/Admin sichtbar.');
    } else {
      await l.page.getByRole('heading', { level: 1, name: D.titel[D.archiv] }).waitFor({ timeout: 15000 });
      await l.shot('deeplink-archiviert', 'Deep-Link auf einen archivierten Use Case als Organizer/Admin: die Detailseite ist erreichbar.');
    }
  }, { trotzdem: true });

  await l.schritt('deeplink-ungueltig', async () => {
    await l.gehe({ uc: 'abc' });
    await l.wartLanding();   // kein Deep-Link → Landing Page, keine halb gelesene Id
  }, { trotzdem: true });
}

/** Ein Fehler-/Leerzustand: Landing → Wand → Deep-Link → (je nach Rolle) Studio, Protokoll. */
async function zustand(l, opt) {
  const t = l.t; const R = l.spec.role; const istOrg = R === 'admin' || R === 'organizer';
  const o = opt || {};
  const z = o.tag || l.spec.state;
  // Ein Lesefehler ist keine Null: In diesen Zuständen darf die Seite nicht „leer", „0" oder „nicht gefunden" sagen.
  const pruefeText = async (seite, sel, regex) => {
    if (!regex) return;
    const txt = await l.page.locator(sel).first().innerText();
    const m = txt.match(regex);
    if (m) l.befund(seite, 'Die Seite stellt einen Lesefehler als leere Liste / „0" / „nicht gefunden" dar — eine Aussage über Daten, die niemand lesen konnte.', `Text enthält „${m[0]}" (Zustand ${l.spec.state}) — ${txt.replace(/\s+/g, ' ').slice(0, 160)}`);
  };
  await l.gehe(o.extra);
  await l.schritt('landing', async () => {
    if (o.sofort) await l.page.locator('.landing h1').waitFor({ timeout: 20000 });
    else await l.wartLanding();
    await l.shot(`landing-${z}`, o.was.landing, { warte: o.sofort ? 100 : 350 });
    await pruefeText(`landing-${z}`, '.landing', o.nichtSagen);
  });
  await l.schritt('start', async () => {
    await l.start();
    if (o.startBild) await l.shot(`start-${z}`, o.was.start, { voll: true });
  });
  await l.schritt('usecases', async () => {
    await l.kachel(t.kUseCases).click();
    await l.page.locator('.page-container h1.dex-ui-page-head-title').waitFor({ timeout: 10000 });
    if (!o.sofort) await l.page.waitForFunction(() => !/werden geladen|Loading use cases/.test((document.querySelector('.page-container') || {}).innerText || ''), null, { timeout: 25000 });
    await l.shot(`usecases-${z}`, o.was.usecases, { voll: true, warte: o.sofort ? 100 : 350 });
    await pruefeText(`usecases-${z}`, '.page-container', o.nichtSagen);
  });
  if (o.deeplink !== false) {
    await l.schritt('deeplink', async () => {
      await l.gehe({ ...(o.extra || {}), uc: 3 });
      if (o.sofort) await l.page.locator('.page-container').waitFor({ timeout: 15000 });
      else await l.wartInhalt();
      await l.shot(`deeplink-${z}`, o.was.deeplink, { warte: o.sofort ? 100 : 350 });
      await pruefeText(`deeplink-${z}`, '.page-container', o.nichtSagen);
    });
  }
  if (istOrg && o.studio) {
    await l.schritt('studio', async () => {
      await l.gehe(o.extra); await l.wartLanding(); await l.start();
      await l.kachel(t.kStudio).click();
      await l.page.locator('.page-container h1', { hasText: 'Use Case Studio' }).waitFor({ timeout: 10000 });
      await l.page.waitForTimeout(400);
      await l.shot(`studio-${z}`, o.was.studio, { voll: true });
      await pruefeText(`studio-${z}`, '.page-container', o.nichtSagen);
    });
  }
  if (istOrg && o.protokoll) {
    await l.schritt('protokoll', async () => {
      await l.gehe(o.extra); await l.wartLanding(); await l.start();
      await l.kachel(t.kProtokoll).click();
      await l.wartSeite();
      await l.shot(`protokoll-${z}`, o.was.protokoll, { voll: true });
      await pruefeText(`protokoll-${z}`, '.page-container', o.nichtSagenProtokoll);
    });
  }
  if (o.nachher) await o.nachher(l);
}

/* --------------------------------------------------------- Lauf-Planung ---- */

/* Was eine Seite bei einem LESEFEHLER nicht sagen darf: leer, „0", „nicht (mehr) da". */
const LESEFEHLER_TEXTE = /Noch keine Use Cases|No use cases yet|Noch nichts angelegt|Nothing here yet|\b0 Einträge|\b0 entries|\b0 Demos|nicht \(mehr\) da|not \(or no longer\) here/;

function plane() {
  const laeufe = [];
  const add = (id, spec, fn) => { if (!NUR || id.indexOf(NUR) >= 0) laeufe.push({ id, spec: { id, lang: SPRACHE, state: 'ok', ...spec }, fn }); };

  for (const role of ROLLEN) {
    for (const device of GERAETE) {
      add(`${role}-${device}-reise`, { role, device }, reise);
      add(`${role}-${device}-kopf`, { role, device }, kopf);
      add(`${role}-${device}-deeplink`, { role, device }, deeplink);
    }
  }
  // Kachelbild hochladen/zuschneiden/entfernen: Organizer-Weg, am Rechner und am Handy. Echte Uhr: Der Anhang-Name trägt einen Zeitstempel.
  if (ROLLEN.indexOf('admin') >= 0 && GERAETE.indexOf('desktop') >= 0) add('admin-desktop-bild', { role: 'admin', device: 'desktop', uhr: false }, bild);
  if (ROLLEN.indexOf('organizer') >= 0 && GERAETE.indexOf('mobile') >= 0) add('organizer-mobile-bild', { role: 'organizer', device: 'mobile', uhr: false }, bild);
  // Ohne Foto: die Initialen der Kopfzeile (der Zustand einer Person ohne Profilbild).
  if (ROLLEN.indexOf('user') >= 0 && GERAETE.indexOf('desktop') >= 0) {
    add('user-desktop-ohnefoto', { role: 'user', device: 'desktop', ohneFoto: true }, async l => {
      await l.gehe();
      await l.schritt('ohnefoto', async () => { await l.wartLanding(); await l.shot('landing-ohnefoto', 'Landing Page einer Person ohne Profilbild: die Kopfzeile zeigt die Initialen im grauen Kreis.'); });
    });
  }
  if (!MIT_ZUSTAENDEN) return laeufe;

  const dev = d => GERAETE.indexOf(d) >= 0;
  const gib = (role, device) => ROLLEN.indexOf(role) >= 0 && dev(device);

  if (gib('admin', 'desktop')) {
    add('zustand-forbidden-admin-desktop', { role: 'admin', device: 'desktop', state: 'forbidden' }, l => zustand(l, {
      studio: true, protokoll: true, nichtSagen: LESEFEHLER_TEXTE,
      was: {
        landing: 'Landing Page, wenn die Use-Case-Liste mit 403 antwortet: Hinweiskarte „konnte nicht gelesen werden" statt einer erfundenen Zahl.',
        usecases: 'Kachelwand bei 403 auf die Use-Case-Liste: roter Kasten „konnten nicht geladen werden" mit Ursache (Leserecht), HTTP-Status und Klartext von SharePoint — kein „Keine Use Cases".',
        deeplink: 'Deep-Link auf einen Use Case bei 403: Fehlerkasten „konnte nicht geladen werden", kein „nicht (mehr) da".',
        studio: 'Use Case Studio bei 403 auf die Use-Case-Liste (Admin).',
        protokoll: 'Protokoll (lesbar) bei 403 auf die Use-Case-Liste.',
      },
    }));
    add('zustand-error-admin-desktop', { role: 'admin', device: 'desktop', state: 'error' }, l => zustand(l, {
      studio: true, nichtSagen: LESEFEHLER_TEXTE,
      was: {
        landing: 'Landing Page, wenn die Use-Case-Liste mit 500 antwortet.',
        usecases: 'Kachelwand bei HTTP 500: Lade-Fehler statt leerer Plattform, mit „Erneut versuchen".',
        deeplink: 'Deep-Link auf einen Use Case bei HTTP 500.',
        studio: 'Use Case Studio bei HTTP 500 auf die Use-Case-Liste (Admin).',
      },
    }));
    add('zustand-empty-admin-desktop', { role: 'admin', device: 'desktop', state: 'empty' }, l => zustand(l, {
      studio: true, protokoll: true,
      was: {
        landing: 'Landing Page bei leerer Liste UND leerem Protokoll (Erstinstallation): die App befüllt sich beim Start selbst mit den Start-Use-Cases — Zahl auf der Hinweiskarte.',
        usecases: 'Kachelwand nach der automatischen Erstbefüllung (Liste und Protokoll waren leer): die Start-Use-Cases stehen da.',
        deeplink: 'Deep-Link ?uc=3 nach der Erstbefüllung.',
        studio: 'Use Case Studio nach der Erstbefüllung.',
        protokoll: 'Protokoll nach der Erstbefüllung: der Merker „Erstbefüllung" ist der einzige Eintrag.',
      },
    }));
    add('zustand-leer-admin-desktop', { role: 'admin', device: 'desktop', state: 'leer' }, l => zustand(l, {
      studio: true, protokoll: true,
      was: {
        landing: 'Landing Page, wenn alle Use Cases absichtlich gelöscht wurden (Protokoll trägt den Erstbefüllungs-Merker): „Noch keine Use Cases" mit Hinweis auf das Studio.',
        usecases: 'Kachelwand ohne Use Cases (Admin): Leerzustand mit „Die fünf Use Cases aus dem Konzept anlegen" und „Eigenen Use Case anlegen".',
        deeplink: 'Deep-Link ?uc=3 bei leerer Plattform: Leerzustand „nicht (mehr) da".',
        studio: 'Use Case Studio ohne Einträge: „Noch nichts angelegt" mit „Ersten Use Case anlegen" und „Die Start-Use-Cases anlegen".',
        protokoll: 'Protokoll mit nur dem Merker der Erstbefüllung.',
      },
    }));
    add('zustand-nolog-admin-desktop', { role: 'admin', device: 'desktop', state: 'nolog' }, l => zustand(l, {
      deeplink: false, protokoll: true,
      was: {
        landing: 'Landing Page bei gefüllter Liste, aber leerem Protokoll.',
        usecases: 'Kachelwand bei gefüllter Liste, aber leerem Protokoll (keine Erstbefüllung, weil die Liste nicht leer ist).',
        protokoll: 'Protokoll ohne Einträge: „Noch nichts protokolliert".',
      },
    }));
    add('zustand-logerror-admin-desktop', { role: 'admin', device: 'desktop', state: 'logerror' }, l => zustand(l, {
      deeplink: false, protokoll: true, nichtSagenProtokoll: /Noch nichts protokolliert|Nothing logged yet|\b0 Einträge|\b0 entries/,
      was: {
        landing: 'Landing Page, wenn nur das Protokoll mit 500 antwortet (Use Cases lesbar).',
        usecases: 'Kachelwand, wenn nur das Protokoll mit 500 antwortet.',
        protokoll: 'Protokoll bei HTTP 500: Fehlerkasten „konnte nicht gelesen werden" statt „Noch nichts protokolliert".',
      },
    }));
    add('zustand-fresh-admin-desktop', { role: 'admin', device: 'desktop', state: 'fresh' }, l => zustand(l, {
      deeplink: false, protokoll: true,
      nachher: async ll => {
        await ll.schritt('rollen', async () => {
          await ll.gehe({ state: 'fresh' }); await ll.wartLanding(); await ll.start();
          await ll.kachel(ll.t.kRollen).click();
          await ll.wartSeite();
          await ll.shot('rollen-fresh', 'Rollenverwaltung nach einer Erstinstallation: genau eine Zeile — die anlegende Person als Admin.');
        });
      },
      was: {
        landing: 'Erstinstallation ohne jede Liste: Die App legt Listen, Spalten und Rechte an, macht die anlegende Person zum Admin und befüllt die Start-Use-Cases.',
        usecases: 'Kachelwand nach der Erstinstallation (alle Listen frisch angelegt).',
        protokoll: 'Protokoll nach der Erstinstallation: der Merker der Erstbefüllung.',
      },
    }));
    add('zustand-laden-admin-desktop', { role: 'admin', device: 'desktop', state: 'ok' }, l => zustand(l, {
      extra: { delay: 250 }, sofort: true, tag: 'laedt',
      was: { landing: 'Landing Page WÄHREND die Use Cases laden (jede Antwort 250 ms verzögert, rund 35 Anfragen bis zur Wand): Hinweiskarte mit Fortschrittsbalken „Die Use Cases werden geladen …".', usecases: 'Kachelwand WÄHREND des Ladens: Fortschrittsbalken und „Use Cases werden geladen …".', deeplink: 'Deep-Link ?uc=3 WÄHREND die Use Cases laden: der Ladezustand der Detailseite (bei echten Antwortzeiten dauert er Sekunden).' },
    }));
    add('zustand-erstinstallation-first-desktop', { role: 'first', device: 'desktop', state: 'ok' }, async l => {
      await l.gehe();
      await l.schritt('erstinstallation', async () => {
        await l.wartLanding();
        await l.shot('landing-erstinstallation', 'Rollenliste leer und lesbar (Erstinstallation): die angemeldete Person wird Admin — die Fußzeile sagt „Deine Rolle: Admin".');
        const txt = await l.page.locator('footer').innerText();
        if (!/Admin/.test(txt)) l.befund('erstinstallation', 'Bei leerer Rollenliste wird die angemeldete Person nicht Admin.', txt);
      });
    });
  }
  if (gib('user', 'mobile')) {
    add('zustand-forbidden-user-mobile', { role: 'user', device: 'mobile', state: 'forbidden' }, l => zustand(l, {
      startBild: true, nichtSagen: LESEFEHLER_TEXTE,
      was: {
        landing: 'Handy, User: Landing Page bei 403 auf die Use-Case-Liste.',
        start: 'Handy, User: Start-Übersicht bei 403 auf die Use-Case-Liste.',
        usecases: 'Handy, User: Kachelwand bei 403 — Fehlerkasten mit Ursache und „Erneut versuchen".',
        deeplink: 'Handy, User: Deep-Link bei 403.',
      },
    }));
    add('zustand-empty-user-mobile', { role: 'user', device: 'mobile', state: 'empty' }, l => zustand(l, {
      was: {
        landing: 'Handy, User bei leerer Liste UND leerem Protokoll: Ein User darf nichts anlegen, die Erstbefüllung scheitert am Recht — Landing zeigt „Noch keine Use Cases".',
        usecases: 'Handy, User: leere Kachelwand („Noch keine Use Cases", ohne die Anlege-Knöpfe, die nur Organizer sehen).',
        deeplink: 'Handy, User: Deep-Link bei leerer Plattform.',
      },
    }));
  }
  if (gib('organizer', 'desktop')) {
    add('zustand-roles403-organizer-desktop', { role: 'organizer', device: 'desktop', state: 'roles403' }, l => zustand(l, {
      startBild: true, deeplink: false,
      was: {
        landing: 'Organizer, dem das Leserecht auf der Rollenliste fehlt (403): Die Rolle wirkt nicht, die Fußzeile sagt „Deine Rolle: User", oben steht die Warnung.',
        start: 'Start-Übersicht bei 403 auf die Rollenliste: Warnkasten „Deine Rolle konnte nicht geprüft werden", Studio-Kachel ausgegraut mit Erklärung.',
        usecases: 'Kachelwand bei 403 auf die Rollenliste: Use Cases lesbar, aber ohne Organizer-Rechte (kein Studio-Knopf).',
      },
    }));
  }
  if (gib('organizer', 'mobile')) {
    add('zustand-roles403-organizer-mobile', { role: 'organizer', device: 'mobile', state: 'roles403' }, l => zustand(l, {
      startBild: true, deeplink: false,
      was: {
        landing: 'Handy, Organizer bei 403 auf die Rollenliste: Landing Page mit Warnung.',
        start: 'Handy, Organizer bei 403 auf die Rollenliste: Start-Übersicht mit Warnkasten und ausgegrautem Studio.',
        usecases: 'Handy, Organizer bei 403 auf die Rollenliste: Kachelwand ohne Organizer-Rechte.',
      },
    }));
  }
  return laeufe;
}

/* ---------------------------------------------------------------- main ---- */

(async () => {
  if (!exe) console.warn('Kein Chromium unter /opt/pw-browsers gefunden — Playwright nimmt seinen eigenen.');
  fs.mkdirSync(SHOTS, { recursive: true });
  if (!NUR) fs.readdirSync(SHOTS).forEach(f => { if (/\.(png|log)$/.test(f)) fs.unlinkSync(path.join(SHOTS, f)); });
  if (!fs.existsSync(path.join(OUT, 'bundle.js'))) { console.error('out/bundle.js fehlt — erst `node build.js`.'); process.exit(2); }

  const server = await start(0);
  const base = `http://127.0.0.1:${server.address().port}`;
  const browser = await chromium.launch({ executablePath: exe, headless: true });
  const bericht = { erstellt: new Date().toISOString(), laeufe: [] };
  const plan = plane();
  console.log(`${plan.length} Läufe geplant.`);

  for (const p of plan) {
    const t0 = Date.now();
    const l = new Lauf(browser, base, p.spec, bericht);
    try {
      await p.fn(l);
    } catch (e) {
      l.befund('lauf', `Der Lauf „${p.id}" ist abgebrochen.`, String(e && e.message ? e.message : e));
    }
    await l.ende();
    l.rec.dauerMs = Date.now() - t0;
    const k = l.rec.konsole;
    console.log(`${p.id.padEnd(42)} ${String(l.rec.shots.length).padStart(3)} Bilder · ${String(l.rec.befunde.length).padStart(2)} Befunde · pageerror ${k.pageerrors.length} · console.error ${k.fehler.length} · ${(l.rec.dauerMs / 1000).toFixed(1)} s`);
  }

  await browser.close();
  server.close();
  fs.writeFileSync(path.join(OUT, 'report.json'), JSON.stringify(bericht, null, 2));

  const alle = bericht.laeufe;
  const bilder = alle.reduce((n, r) => n + r.shots.length, 0);
  const pe = alle.reduce((n, r) => n + r.konsole.pageerrors.length, 0);
  const bf = alle.reduce((n, r) => n + r.befunde.length, 0);
  console.log(`\n${alle.length} Läufe · ${bilder} Bilder · ${bf} Befunde · ${pe === 0 ? 'ALLE LÄUFE OHNE PAGEERROR' : 'PAGEERRORS: ' + pe}`);
  console.log('Bericht: out/report.json · Bilder: out/shots/');
})().catch(e => { console.error(e); process.exit(1); });
