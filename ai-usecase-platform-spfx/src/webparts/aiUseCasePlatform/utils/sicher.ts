/**
 * Werte aus der Liste sind NICHT vertrauenswürdig.
 *
 * Die Pflegeseite prüft, was ein Organizer eintippt — aber die Liste
 * `AIUC_UseCases` ist in SharePoint auch direkt beschreibbar (jede Person mit
 * Schreibrecht auf der Liste), und die Beschreibung hat in der Pflegeseite
 * gar kein Feld. Was hier ankommt, kann also alles sein. Die Sichtprüfung des
 * Sicherheits-Reviews (29.09.2026) hat drei Stellen gefunden, an denen ein
 * Listenwert ungeprüft im Browser landete:
 *
 *  - `dangerouslySetInnerHTML` mit der Beschreibung (gespeichertes XSS: das
 *    Skript liefe im Origin der SharePoint-Seite, mit der Sitzung des
 *    Betrachters — auch eines Admins),
 *  - `<iframe src>` und `window.open` mit einem Link (`javascript:` in einem
 *    iframe erbt den Origin der einbettenden Seite),
 *  - `mailto:` mit einer Betreuer-Adresse (`a@b.de?bcc=…` hängt Empfänger an).
 *
 * Die drei Funktionen hier sind die EINE Stelle dafür. Wer einen neuen Wert
 * aus der Liste in href, src, window.open oder HTML setzt, geht hindurch.
 */

/**
 * Nur absolute http(s)-Adressen — alles andere wird zu `''`.
 *
 * Bewusst KEIN `new URL(v).protocol`: Browser entfernen Tabulatoren und
 * Zeilenumbrüche aus einer Adresse, bevor sie das Schema lesen, und ignorieren
 * Groß-/Kleinschreibung und Leerraum davor — `java\tscript:…` und
 * `  JaVaScRiPt:…` ergeben beide `javascript:`. Die Prüfung verlangt deshalb
 * `http(s)://` unmittelbar am Anfang und lässt Leerraum, Steuerzeichen,
 * Anführungszeichen, spitze Klammern und Rückwärtsschrägstriche nirgends zu.
 * Relative Werte (`/sites/…`, `//fremd.example`, `github.com/x`) fallen mit
 * heraus: Sie zeigten auf die SharePoint-Seite selbst bzw. auf einen anderen
 * Host als gemeint.
 */
export function sichereUrl(v: string | undefined | null): string {
  const s = (v || '').replace(/^\s+|\s+$/g, '');
  if (!s) return '';
  if (!/^https?:\/\/[^\s/?#]+/i.test(s)) return '';
  // eslint-disable-next-line no-control-regex
  if (/[\s\u0000-\u001f\u007f"<>\\]/.test(s)) return '';
  return s;
}

/**
 * Eine einzelne E-Mail-Adresse — sonst `''`.
 *
 * Streng, weil das Ergebnis in `mailto:` landet: `?`, `&`, `=`, `,` und `%`
 * hängen dort Empfänger, Betreff und Text an (`a@b.de?bcc=x@fremd.example`).
 * Der Apostroph bleibt erlaubt (O'Brien).
 */
export function sichereMail(v: string | undefined | null): string {
  const s = (v || '').replace(/^\s+|\s+$/g, '');
  return /^[A-Za-z0-9._+'-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$/.test(s) ? s : '';
}

/** Tags, die bleiben. Alles andere wird ausgepackt (Text bleibt) oder ganz entfernt. */
const ERLAUBT: { [tag: string]: boolean } = {
  p: true, br: true, strong: true, b: true, em: true, i: true, u: true,
  ul: true, ol: true, li: true, a: true, h3: true, h4: true, blockquote: true, code: true,
};

/** Tags, die samt Inhalt verschwinden — ihr Text ist kein Text für Leser. */
const VERWERFEN: { [tag: string]: boolean } = {
  script: true, style: true, iframe: true, object: true, embed: true, form: true,
  input: true, button: true, textarea: true, select: true, svg: true, math: true,
  link: true, meta: true, base: true, template: true, noscript: true, frame: true,
  frameset: true, applet: true, title: true, head: true,
};

/**
 * HTML auf eine kleine Liste von Auszeichnungen zurückführen.
 *
 * Das Ergebnis enthält nur `p, br, strong, b, em, i, u, ul, ol, li, a, h3, h4,
 * blockquote, code` — und KEIN Attribut außer `href` an `a` (nur `http(s)`,
 * öffnet in neuem Tab mit `noopener noreferrer`). Ereignis-Attribute, `style`,
 * `class`, Bilder und Skripte gibt es danach nicht mehr.
 *
 * Geparst wird mit `DOMParser` in ein Dokument OHNE Browserkontext: Dort läuft
 * weder ein Skript noch lädt ein Bild, und `onerror` feuert nie. Das
 * Bereinigte wird als Text zusammengesetzt (Text über `textContent` maskiert),
 * nicht aus dem Original kopiert.
 */
export function bereinigeHtml(html: string | undefined | null): string {
  const roh = html || '';
  if (!roh) return '';
  let doc: Document;
  try {
    doc = new DOMParser().parseFromString(roh, 'text/html');
  } catch {
    return '';
  }
  const maskiere = (t: string): string => t
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  const baue = (knoten: Node): string => {
    let out = '';
    for (let i = 0; i < knoten.childNodes.length; i++) {
      const k = knoten.childNodes[i];
      if (k.nodeType === 3) { // Text
        out += maskiere(k.nodeValue || '');
        continue;
      }
      if (k.nodeType !== 1) continue; // Kommentare, Verarbeitungsanweisungen: weg
      const el = k as Element;
      const tag = el.tagName.toLowerCase();
      if (VERWERFEN[tag]) continue;
      if (!ERLAUBT[tag]) { out += baue(el); continue; }
      if (tag === 'br') { out += '<br>'; continue; }
      if (tag === 'a') {
        const href = sichereUrl(el.getAttribute('href'));
        const inner = baue(el);
        out += href
          ? `<a href="${maskiere(href)}" target="_blank" rel="noopener noreferrer">${inner}</a>`
          : inner;
        continue;
      }
      out += `<${tag}>${baue(el)}</${tag}>`;
    }
    return out;
  };

  return doc.body ? baue(doc.body) : '';
}
