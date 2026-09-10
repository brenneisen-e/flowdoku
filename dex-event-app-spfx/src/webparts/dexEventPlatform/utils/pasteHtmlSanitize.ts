/* Eingefügtes HTML für den Mail-Editor säubern.
 *
 * Vorgeschichte: `handlePaste` in `HtmlEditorModal` hat seit v18.39 JEDES
 * eingefügte HTML verworfen und nur `text/plain` übernommen. Der Grund war
 * richtig — Outlook und Word bringen `<div>`/`<p>` mit eigenen Außenabständen,
 * `mso-`-Eigenschaften, Klassen und Schriftfamilien mit; einmal im Editor ließ
 * sich der „plötzlich größere Zeilenabstand" nicht mehr wegklicken. Nur war
 * der Preis zu hoch: Beim Einfügen einer fertigen Outlook-Mail gingen ALLE
 * Links verloren, Fett und Überschriften wurden flach, und die Aufzählungs-
 * punkte kamen als „•"-Zeichen im Fließtext an (so gemeldet am 10.09.2026 mit
 * einer Rundmail an 150 Empfänger).
 *
 * Die Lösung ist keine Rücknahme, sondern eine Positivliste: Wir behalten
 * genau die Auszeichnungen, die eine Mail braucht — Link, fett, kursiv,
 * unterstrichen, Aufzählung, Zeilenumbruch — und werfen ALLES andere weg,
 * insbesondere jedes `style`, jede Klasse und jede Breitenangabe. Damit kann
 * das eingefügte Fremd-HTML den Zeilenabstand nicht mehr anfassen, und die
 * Begründung von v18.39 bleibt erfüllt.
 *
 * Bewusst NICHT übernommen:
 *  - Bilder. Ein `<img>` aus einer fremden Mail zeigt entweder auf einen
 *    Server, den der Empfänger nicht erreicht, oder trägt ein Base64-Bild, das
 *    an der KB-Grenze der Mail-Warteschlange vorbeikäme. Bilder haben ihren
 *    eigenen, größengeprüften Weg (Knopf bzw. Strg+V ohne Text, v31.7).
 *  - Tabellen. Ohne die Breiten- und Rahmenangaben, die wir strippen, fällt
 *    eine Outlook-Tabelle unvorhersehbar zusammen. Der Inhalt bleibt, Zeile
 *    für Zeile — das ist die verlässlichere Zusage.
 *  - Farben und Schriftgrößen. Sie sind im Editor eigene Knöpfe; aus einer
 *    fremden Mail übernommen wären sie der Anfang genau des Durcheinanders,
 *    das v18.39 verhindern wollte.
 *
 * Fett/kursiv/unterstrichen erkennt der Filter AUCH, wenn sie als Inline-Style
 * ankommen (`<span style="font-weight:bold">`) — Outlook schreibt es oft so.
 * Ohne diese Umsetzung wäre die halbe Formatierung trotz Positivliste weg.
 */

/** Elemente, die restlos verschwinden — mitsamt Inhalt. */
const DROP = new Set([
  'SCRIPT', 'STYLE', 'HEAD', 'META', 'LINK', 'TITLE', 'NOSCRIPT', 'TEMPLATE',
  'IMG', 'PICTURE', 'SOURCE', 'SVG', 'CANVAS', 'VIDEO', 'AUDIO', 'IFRAME',
  'OBJECT', 'EMBED', 'FORM', 'INPUT', 'BUTTON', 'SELECT', 'TEXTAREA', 'OPTION',
]);

/** Elemente, die als Einheit erhalten bleiben (Tag → Ausgabe-Tag). */
const KEEP: { [tag: string]: string } = {
  B: 'b', STRONG: 'b', I: 'i', EM: 'i', U: 'u', S: 's', STRIKE: 's', DEL: 's',
  UL: 'ul', OL: 'ol', LI: 'li',
};

/** Elemente, deren Ende eine neue Zeile bedeutet. */
const BLOCK = new Set([
  'P', 'DIV', 'SECTION', 'ARTICLE', 'BLOCKQUOTE', 'ADDRESS', 'HEADER', 'FOOTER',
  'MAIN', 'NAV', 'ASIDE', 'FIGURE', 'FIGCAPTION', 'DL', 'DT', 'DD', 'PRE',
  'CENTER', 'HR', 'TR', 'TABLE', 'CAPTION',
]);

/** Nur diese Ziele sind in einer Mail sinnvoll — alles andere wird entlinkt. */
const SAFE_HREF = /^(https?:\/\/|mailto:|tel:)/i;

const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Fett/kursiv/unterstrichen aus einem Inline-Style herauslesen. */
function styleWrap(el: Element): { open: string; close: string } {
  const st = el.getAttribute('style') || '';
  if (!st) return { open: '', close: '' };
  let open = '';
  let close = '';
  const tag = el.tagName;
  if (tag !== 'B' && tag !== 'STRONG' && /font-weight\s*:\s*(bold|bolder|[6-9]00)/i.test(st)) {
    open += '<b>'; close = '</b>' + close;
  }
  if (tag !== 'I' && tag !== 'EM' && /font-style\s*:\s*italic/i.test(st)) {
    open += '<i>'; close = '</i>' + close;
  }
  if (tag !== 'U' && /text-decoration[^;]*underline/i.test(st)) {
    open += '<u>'; close = '</u>' + close;
  }
  return { open, close };
}

function renderNode(node: Node): string {
  // Textknoten: Der Quelltext einer Mail ist eingerückt und umbrochen; ohne
  // das Zusammenziehen stünden diese Umbrüche als Leerzeichen im Ergebnis.
  if (node.nodeType === 3) return esc((node.nodeValue || '').replace(/\s+/g, ' '));
  if (node.nodeType !== 1) return ''; // Kommentare (auch die von Word) fallen weg

  const el = node as Element;
  const tag = el.tagName;
  if (DROP.has(tag)) return '';
  if (tag === 'BR') return '<br>';

  let inner = '';
  for (let i = 0; i < el.childNodes.length; i++) inner += renderNode(el.childNodes[i]);

  const { open, close } = styleWrap(el);

  if (tag === 'A') {
    const href = (el.getAttribute('href') || '').trim();
    // Ein Link ohne brauchbares Ziel wird zu seinem Text — ein toter Link ist
    // schlechter als gar keiner.
    if (!SAFE_HREF.test(href)) return open + inner + close;
    return `${open}<a href="${esc(href)}">${inner}</a>${close}`;
  }
  if (/^H[1-6]$/.test(tag)) return `<b>${inner}</b><br>`;
  if (KEEP[tag]) {
    const t = KEEP[tag];
    return `${open}<${t}>${inner}</${t}>${close}`;
  }
  // Tabellenzellen werden zu Text mit Abstand, nicht zu einer neuen Zeile —
  // sonst zerfiele eine dreispaltige Zeile in drei Zeilen. Outlook packt in
  // jede Zelle ein `<p class="MsoNormal">`; dessen Umbruch am Zellenende muss
  // deshalb weg, sonst tut genau das, was hier verhindert werden soll.
  if (tag === 'TD' || tag === 'TH') return open + inner.replace(/(\s|<br>)+$/i, '') + close + ' ';
  if (BLOCK.has(tag)) return open + inner + close + '<br>';
  return open + inner + close;
}

/**
 * Macht eingefügtes Fremd-HTML für den Editor brauchbar.
 *
 * @returns Gesäubertes HTML, oder '' wenn nichts Sichtbares übrig bleibt —
 *          dann fällt der Aufrufer auf den reinen Text zurück.
 */
export function sanitizePastedHtml(html: string): string {
  if (!html || !html.trim()) return '';
  let doc: Document;
  try {
    // Ein per DOMParser erzeugtes Dokument ist inaktiv: kein Skript läuft,
    // keine Ressource wird geladen. Deshalb wird hier geparst und NICHT über
    // ein innerHTML im lebenden Baum.
    doc = new DOMParser().parseFromString(html, 'text/html');
  } catch { return ''; }
  if (!doc || !doc.body) return '';

  let out = renderNode(doc.body);

  // Aufräumen: leere Auszeichnungen, Umbruch-Kaskaden (eine Outlook-Mail
  // besteht aus verschachtelten <div>, jedes davon liefert ein <br>) und die
  // Umbrüche am Anfang und Ende der Einfügung.
  for (let i = 0; i < 3; i++) {
    out = out.replace(/<(b|i|u|s|a)(\s[^>]*)?>(\s|<br>)*<\/\1>/gi, '');
  }
  out = out.replace(/(<br>\s*){3,}/gi, '<br><br>');
  out = out.replace(/^(\s|<br>)+/i, '').replace(/(\s|<br>)+$/i, '');
  // Ein <br> direkt vor oder nach einer Liste erzeugt in Mail-Clients eine
  // zusätzliche Leerzeile — die Liste bringt ihren Abstand selbst mit.
  out = out.replace(/<br>\s*<(ul|ol)>/gi, '<$1>').replace(/<\/(ul|ol)>\s*<br>/gi, '</$1>');

  // Nichts Sichtbares übrig? Dann hat der Aufrufer mit dem Klartext die
  // bessere Karte.
  return /[^\s]/.test(out.replace(/<[^>]*>/g, '')) ? out : '';
}
