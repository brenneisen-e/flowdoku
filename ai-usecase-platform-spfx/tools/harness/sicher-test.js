/**
 * Prüft `utils/sicher.ts` im echten Chromium mit Angriffsbeispielen.
 *
 *   node sicher-test.js        (nach `npm i` in diesem Ordner)
 *
 * Das Produkt hat keine Tests; die Bereinigung von Listenwerten ist aber die Stelle,
 * an der ein Fehler ein gespeichertes XSS wäre (Sicherheits-Review 29.09.2026). Der
 * Lauf bündelt die Datei mit esbuild, setzt jede Zeichenkette durch `bereinigeHtml`,
 * schreibt das Ergebnis in ein echtes DOM und prüft, dass kein Skript lief. Exit 1,
 * wenn ein Fall durchrutscht.
 */
const path = require('path');
const fs = require('fs');
const esbuild = require('esbuild');
const { chromium } = require('playwright');

const QUELLE = path.resolve(__dirname, '../../src/webparts/aiUseCasePlatform/utils/sicher.ts');

const CHROMIUM = [process.env.PLAYWRIGHT_CHROMIUM, '/opt/pw-browsers/chromium']
  .concat(fs.existsSync('/opt/pw-browsers') ? fs.readdirSync('/opt/pw-browsers').filter(d => /^chromium-/.test(d)).map(d => `/opt/pw-browsers/${d}/chrome-linux/chrome`) : [])
  .filter(p => p && fs.existsSync(p))[0];

/** [Eingabe, erwartete Ausgabe] — `null` = nur prüfen, dass nichts Gefährliches drinsteht. */
const HTML_FAELLE = [
  ['<p>Hallo <strong>Welt</strong></p>', '<p>Hallo <strong>Welt</strong></p>'],
  ['<img src=x onerror="window.__xss=1">', ''],
  ['<script>window.__xss=1</script>Text', 'Text'],
  ['<a href="javascript:alert(1)">klick</a>', 'klick'],
  ['<a href="  JaVaScRiPt:alert(1)">bös</a>', 'bös'],
  ['<a href="java\tscript:alert(1)">tab</a>', 'tab'],
  ['<a href="https://ok.example/x?a=1&b=2" onclick="evil()">gut</a>', '<a href="https://ok.example/x?a=1&amp;b=2" target="_blank" rel="noopener noreferrer">gut</a>'],
  ['<svg onload="window.__xss=1"><circle/></svg>Nach svg', 'Nach svg'],
  ['<iframe src="https://evil"></iframe>x', 'x'],
  ['<div style="background:url(javascript:1)">im div <em>kursiv</em></div>', 'im div <em>kursiv</em>'],
  ['<p onmouseover="x()">hover</p>', '<p>hover</p>'],
  ['<math><mi xlink:href="javascript:alert(1)">m</mi></math>', ''],
  ['<form action="https://evil"><input name=x><button>ok</button></form>Rest', 'Rest'],
  ['&lt;script&gt;nicht ausführen&lt;/script&gt;', '&lt;script&gt;nicht ausführen&lt;/script&gt;'],
];

const URL_FAELLE = [
  ['https://kiarbeitsplatz.pages.dev/', 'https://kiarbeitsplatz.pages.dev/'],
  ['  https://a.b/c?d=e#f  ', 'https://a.b/c?d=e#f'],
  ['javascript:alert(1)', ''],
  ['  JaVaScRiPt:alert(1)', ''],
  ['java\tscript:alert(1)', ''],
  ['data:text/html,x', ''],
  ['//evil.example', ''],
  ['github.com/x', ''],
  ['https://x/ y', ''],
  ['https:///nohost', ''],
  ['https://a.b/"x', ''],
];

const MAIL_FAELLE = [
  ['erika.muster@deloitte.de', 'erika.muster@deloitte.de'],
  ["o'brien@x.com", "o'brien@x.com"],
  ['a@b.de?bcc=x@evil.de', ''],
  ['x@y.de,z@w.de', ''],
  ['a b@c.de', ''],
  ['a@b', ''],
];

(async () => {
  const bundle = await esbuild.build({ entryPoints: [QUELLE], bundle: true, format: 'iife', globalName: 'S', write: false });
  const browser = await chromium.launch(CHROMIUM ? { executablePath: CHROMIUM } : {});
  const page = await browser.newPage();
  await page.setContent('<html><body></body></html>');
  await page.addScriptTag({ content: bundle.outputFiles[0].text });
  let fehler = 0;
  const pruefe = async (name, fn, faelle) => {
    for (const [ein, soll] of faelle) {
      const ist = await page.evaluate(([f, v]) => window.S[f](v), [fn, ein]);
      if (ist !== soll) { fehler++; console.log(`FEHLER ${name}: ${JSON.stringify(ein)}\n   erwartet ${JSON.stringify(soll)}\n   erhalten ${JSON.stringify(ist)}`); }
    }
    console.log(`${name}: ${faelle.length} Fälle`);
  };
  await pruefe('bereinigeHtml', 'bereinigeHtml', HTML_FAELLE);
  await pruefe('sichereUrl', 'sichereUrl', URL_FAELLE);
  await pruefe('sichereMail', 'sichereMail', MAIL_FAELLE);

  // Ausführen: Alles Gefährliche zusammen in ein echtes DOM setzen und einen Klick auslösen.
  const xss = await page.evaluate(async () => {
    const d = document.createElement('div');
    document.body.appendChild(d);
    d.innerHTML = window.S.bereinigeHtml('<img src=x onerror="window.__xss=1"><script>window.__xss=1</script><a href="javascript:window.__xss=1">x</a><svg onload="window.__xss=1"></svg>');
    const a = d.querySelector('a'); if (a) a.click();
    await new Promise(r => setTimeout(r, 300));
    return window.__xss;
  });
  if (xss !== undefined) { fehler++; console.log('FEHLER: Es lief ein Skript (window.__xss gesetzt).'); }
  console.log(`Ausführung im DOM: ${xss === undefined ? 'nichts gelaufen' : 'SKRIPT GELAUFEN'}`);

  await browser.close();
  console.log(fehler === 0 ? 'ALLE FÄLLE BESTANDEN' : `${fehler} FEHLER`);
  process.exit(fehler === 0 ? 0 : 1);
})().catch(e => { console.error(e); process.exit(1); });
