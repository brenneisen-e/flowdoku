/**
 * v31.7: Bilder INLINE im Mailtext — Kompression, Grenzen, HTML.
 *
 * Warum überhaupt inline und nicht als Anhang: Der Deloitte-Mailflow blockt
 * Power-Automate-Mails mit Anhang (NDR, v26.71). Der Weg, der trägt, ist
 * derselbe, den die QR-Mail, das Deloitte-Logo und das Kopfbild schon gehen —
 * das Bild steckt als `data:image/...;base64,…` mitten im HTML. Diese Datei
 * macht daraus einen Weg, den auch ein Organizer im Editor nehmen kann.
 *
 * **Die Grenzen sind nicht geraten, sie kommen aus der Ablage.** Der fertige
 * Mail-Body landet in `DEX_Emails.Body` — einer mehrzeiligen SharePoint-Spalte
 * (Note). Die nimmt rund 64.000 Zeichen; alles darüber lehnt SharePoint ab,
 * und die Mail steht dann gar nicht erst in der Warteschlange. Von diesen
 * 64.000 Zeichen gehört dem Organizer NICHT alles:
 *
 *  - `wrapTemplate` legt rund 3.500 Zeichen Gerüst drumherum (Kopf, grüne
 *    Linie, Fußzeile, Legal-Absatz),
 *  - das Deloitte-Logo im Kopf steckt als Base64 mit drin (`cachedLogoBase64`,
 *    Größenordnung 5.000–15.000 Zeichen),
 *  - bei „Event-Foto im Kopf" kommt zusätzlich das eingebackene Kopfbild dazu.
 *
 * Deshalb bekommt der Text im Editor ein Budget von 44 KB und nicht 62 —
 * der Rest ist die Reserve für das, was die App selbst noch anbaut. Wer die
 * Zahlen ändert: Es sind ZEICHEN der gespeicherten Zeichenkette, nicht Bytes
 * des Bildes (Base64 ist rund ein Drittel größer als die Datei).
 *
 * Was die Grenze NICHT ist: eine Frage der Empfängerzahl. Die Rundmail legt
 * EINE Zeile für alle an (`queueEmail` mit den Adressen als `;`-Liste in
 * `Recipient`, s. MassmailComposerModal) — das Bild steht also einmal in der
 * Warteschlange, nicht 300-mal. Ein Bild wird dadurch nicht teurer, je mehr
 * Leute es bekommen. Die Mails, die je Person entstehen (Bestätigung, QR),
 * schreiben je Person eine Zeile; dort gilt dieselbe Grenze je Zeile.
 *
 * Dieselben Grenzen gelten für die anderen Ziele des Editors, weil sie
 * dieselbe Art Spalte sind: `OutlookBody` und `Description` auf `DEX_Events`
 * sind ebenfalls Note-Spalten, und die Mail-Vorlagen des Wizards liegen im
 * JSON von `EmailTemplateOverrides` — dort teilen sie sich 2 MB mit dem
 * Mail-Logo, das dort schon steckt. Ein Budget für alle ist deshalb das
 * ehrlichere als vier.
 *
 * Der Roundtrip Speichern → Laden → Speichern trägt eine Data-URL unverändert
 * (geprüft): `replacePlaceholders` sucht `{{…}}` (in Base64 gibt es keine
 * geschweiften Klammern), `wrapTemplate` setzt den Body roh ein,
 * `stripOutlookWrapper` schneidet an `</td>`, `normalizeMadeWithLink` fasst nur
 * `<a>`-Tags an. Die EINE Stelle, die zugeschnappt hätte, war
 * `reinsertOrganizerPlaceholder` — sie suchte Organizer-Namen im ganzen String
 * und traf damit auch ins Base64; sie ersetzt seit v31.7 nur noch im Text
 * zwischen den Tags (s. components/wizard/wizardHelpers).
 */
import { compressImage } from './imageCompress';

const KB = 1024;

/**
 * Breite, auf die ein eingefügtes Bild gerechnet wird.
 *
 * Nicht 600: Die Mail-Tabelle ist 600 px breit, die Inhaltszelle hat aber
 * `padding:0 30px 30px 30px` (s. `wrapTemplate`) — nutzbar sind 540 px. Ein
 * 600-px-Bild würde in Outlook-Desktop über die Spalte hinauslaufen, weil die
 * Word-Engine `max-width` ignoriert (dieselbe Falle wie beim Kopfbild, v28.24).
 */
export const INLINE_IMG_MAX_WIDTH = 540;

/** Zielgröße eines Bildes im gespeicherten Text (KB Zeichen). Darunter hört
 *  die Kompressions-Leiter auf. */
export const INLINE_IMG_TARGET_KB = 28;
/** Absolute Obergrenze je Bild. Was auch nach der letzten Stufe darüber
 *  liegt, wird abgelehnt statt stillschweigend eingefügt. */
export const INLINE_IMG_MAX_KB = 40;
/** Budget für den ganzen Text im Editor (s. Kopf-Kommentar). */
export const MAIL_BODY_MAX_KB = 44;
/** Ab hier wird die Zeile unter dem Editor zur Warnung — noch erlaubt, aber
 *  ein weiteres Bild passt voraussichtlich nicht mehr. */
export const MAIL_BODY_WARN_KB = 30;

/**
 * Ergebnis eines Einfüge-Versuchs. Bewusst EIN Typ statt einer
 * unterschiedenen Union: Das Projekt steht auf `strict: false`, dort trägt
 * die Verengung über `ok` nicht zuverlässig. `chars` ist auch im Fehlerfall
 * gefüllt — die Meldung soll die tatsächlich erreichte Größe nennen, nicht
 * die des Originals.
 */
export interface InlineImageOutcome {
  ok: boolean;
  /** Nur bei `ok: false`. */
  reason?: 'unreadable' | 'too-big';
  dataUrl: string;
  width: number;
  height: number;
  chars: number;
}

/** Zeichen → KB, gerundet, für die Anzeige. */
export function charsToKb(chars: number): number {
  return Math.max(1, Math.round(chars / KB));
}

const fileToDataUrl = (file: File): Promise<string> =>
  new Promise<string>(resolve => {
    const r = new FileReader();
    r.onload = e => resolve((e.target?.result as string) || '');
    r.onerror = () => resolve('');
    r.readAsDataURL(file);
  });

const measure = (dataUrl: string): Promise<{ width: number; height: number }> =>
  new Promise(resolve => {
    const i = new Image();
    i.onload = () => resolve({ width: i.naturalWidth || 0, height: i.naturalHeight || 0 });
    i.onerror = () => resolve({ width: 0, height: 0 });
    i.src = dataUrl;
  });

/**
 * Kompressions-Leiter. Erst die Qualität, dann die Breite — ein Screenshot
 * bleibt so lange wie möglich lesbar, und erst wenn das nicht reicht, wird er
 * kleiner. `flattenToJpeg` ist überall an: Ein PNG-Screenshot wiegt als PNG
 * ein Vielfaches, und der Mailhintergrund ist ohnehin weiß (s. imageCompress).
 */
const LADDER: Array<{ w: number; q: number }> = [
  { w: INLINE_IMG_MAX_WIDTH, q: 0.82 },
  { w: INLINE_IMG_MAX_WIDTH, q: 0.7 },
  { w: INLINE_IMG_MAX_WIDTH, q: 0.58 },
  { w: 460, q: 0.6 },
  { w: 380, q: 0.6 },
  { w: 300, q: 0.55 },
];

/**
 * Datei → fertige Data-URL für den Mailtext.
 *
 * Liefert immer das KLEINSTE Ergebnis der Leiter, auch wenn keine Stufe das
 * Ziel erreicht — dann entscheidet die Obergrenze, ob es eingefügt werden darf.
 * So sieht der Nutzer im Fehlerfall die tatsächlich erreichte Größe und nicht
 * die des Originals.
 */
// v31.9.7: `maxWidth` optional — der Mail-KOPF darf 600 px breit sein (die
// Inhaltszelle der Mail ist schmaler, der Kopf nicht). Ohne Angabe bleibt
// es bei den 540 px fuer Bilder IM Text.
export async function buildInlineImage(file: File, maxWidth?: number): Promise<InlineImageOutcome> {
  const fail = (reason: 'unreadable' | 'too-big', chars: number): InlineImageOutcome =>
    ({ ok: false, reason, dataUrl: '', width: 0, height: 0, chars });
  const target = INLINE_IMG_TARGET_KB * KB;
  // Nur die drei VOLLBREITEN Stufen bekommen den Wunschwert; die drei
  // Reduktions-Stufen (460/380/300) bleiben, was sie sind — sie sind der
  // Notausgang fuer zu grosse Bilder, und den darf ein 600er-Wunsch nicht
  // aufblasen. Ein KLEINERER Wunsch deckelt auch sie.
  const w0 = Math.max(1, Math.round(maxWidth || INLINE_IMG_MAX_WIDTH));
  let best = '';
  for (const step of LADDER) {
    try {
      const out = await fileToDataUrl(await compressImage(file, step.w === INLINE_IMG_MAX_WIDTH ? w0 : Math.min(step.w, w0), step.q, true));
      if (!out || out.indexOf('data:image/') !== 0) continue;
      if (!best || out.length < best.length) best = out;
      if (out.length <= target) break;
    } catch {
      // Eine Stufe kann scheitern (Canvas ohne Kontext, Bild nicht dekodierbar)
      // — die nächste probieren wir trotzdem, sonst hinge alles an der ersten.
      continue;
    }
  }
  if (!best) return fail('unreadable', 0);
  const dim = await measure(best);
  // Breite 0 heißt: nichts Zeichenbares (z.B. ein SVG ohne feste Maße). Ohne
  // diese Prüfung landete ein 0x0-Bild im Text, das niemand sieht und niemand
  // wieder findet.
  if (!dim.width || !dim.height) return fail('unreadable', best.length);
  if (best.length > INLINE_IMG_MAX_KB * KB) return fail('too-big', best.length);
  return { ok: true, dataUrl: best, width: dim.width, height: dim.height, chars: best.length };
}

/**
 * Das `<img>`, das in den Text kommt.
 *
 * Feste `width` PLUS `max-width:100%` — genau in dieser Reihenfolge und nicht
 * `width:100%`: Outlook-Desktop rendert mit der Word-Engine und ignoriert
 * `max-width`, nimmt also die feste Breite; alle anderen Clients schrumpfen auf
 * schmalen Displays über `max-width` weiter mit (Begründung wie in
 * `buildHeroRow`, v28.24).
 */
export function inlineImageHtml(dataUrl: string, width: number): string {
  const w = Math.max(1, Math.min(Math.round(width), INLINE_IMG_MAX_WIDTH));
  return `<img src="${dataUrl}" width="${w}" alt="" style="display:block;max-width:100%;height:auto;margin:12px 0;border:0;" />`;
}

/** Wie viele Inline-Bilder stecken im Text? Nur für die Anzeige. */
export function countInlineImages(html: string): number {
  if (!html) return 0;
  const m = html.match(/<img\b[^>]*src\s*=\s*["']data:image\//gi);
  return m ? m.length : 0;
}

/**
 * Bild aus der Zwischenablage — oder `null`, wenn es ein Text-Einfügen ist.
 *
 * Die Reihenfolge ist der Punkt: Word, Excel und PowerPoint legen NEBEN dem
 * Text ein gerendertes Bild des kopierten Bereichs in die Zwischenablage. Ohne
 * die Textprüfung zuerst würde aus „Absatz aus Word kopieren" ein Screenshot,
 * und das v18.39-Verhalten (Einfügen ist reiner Text) wäre nebenbei
 * ausgehebelt. Nur wenn GAR KEIN Text mitkommt — der Normalfall bei
 * Snipping Tool, Druck-Taste, „Bild kopieren" im Browser — ist es ein Bild.
 */
export function imageFileFromClipboard(dt: DataTransfer | null): File | null {
  if (!dt) return null;
  let text = '';
  try { text = dt.getData('text/plain') || ''; } catch { text = ''; }
  if (text.trim()) return null;
  const items = dt.items;
  if (items) {
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      if (it && it.kind === 'file' && (it.type || '').indexOf('image/') === 0) {
        const f = it.getAsFile();
        if (f) return f;
      }
    }
  }
  const files = dt.files;
  if (files && files.length > 0 && (files[0].type || '').indexOf('image/') === 0) return files[0];
  return null;
}
