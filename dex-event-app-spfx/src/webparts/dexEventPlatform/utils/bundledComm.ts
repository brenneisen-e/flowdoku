/**
 * v30.61: Gebündelte Kommunikation — eine Mail und ein Kalendereintrag für das
 * ganze Event statt einer pro Termin.
 *
 * Nutzer-Ansage 01.09.2026: „Man soll einstellen können, ob jedes Sub-Event
 * seine eigene Kommunikation (Mails und Kalender separat) hat oder nur über die
 * Klammer läuft. Wenn man das entsprechend einstellt, kriegt man statt 10
 * Einzelmails für die Sub-Events eine Mail für das Klammer-Event."
 *
 * **Es ist eine Umkehrung, kein zweiter Versandweg.** Heute gilt im Code
 * `suppressParentNotifications = subEventsOnlyMode`: Die Klammer schweigt
 * bewusst (ihre Zeile ist eine Schattenzeile, v15.25), und jeder Termin
 * verschickt selbst. Gebündelt heißt genau andersherum — die Termine werden
 * still angelegt (`suppressMail`/`suppressOutlook`, beides gibt es schon), und
 * die Klammer verschickt einmal. Damit gibt es weiterhin EINEN Pfad, der
 * Mails auslöst; nur die Ebene wechselt.
 *
 * **Mail und Kalender sind getrennt schaltbar** (Nutzer-Entscheidung): Es gibt
 * Events, bei denen eine Sammelmail richtig ist, die einzelnen Sessions aber im
 * Kalender stehen sollen — und umgekehrt.
 *
 * **Der Kalendereintrag ist der der Klammer.** Ihr Zeitraum umspannt das ganze
 * Event; der Outlook-Flow liest Start und Ende ohnehin aus dem Event-Item, es
 * muss also nichts gerechnet werden. Die Folge, offen gesagt: Wer nur Tag 2 und
 * Tag 4 bucht, bekommt einen Eintrag über Tag 1 bis 5. Das ist die Bedeutung
 * von „ein Termin über alles" — welche Tage gebucht sind, steht in der
 * Beschreibung.
 */

/** Die drei Schalter. Alle default false = bisheriges Verhalten. */
export interface BundledComm {
  /** Eine Sammel-Bestätigung von der Klammer statt einer Mail je Termin. */
  mail: boolean;
  /** Ein Kalendereintrag (der der Klammer) statt eines je Termin. */
  outlook: boolean;
  /** Ein QR-Code fürs Gesamtevent statt eines je Termin. */
  qr: boolean;
}

export const BUNDLED_COMM_DEFAULT: BundledComm = { mail: false, outlook: false, qr: false };

/** Die Piggyback-Schlüssel — beim Laden der Overrides zu strippen (CLAUDE.md). */
export const BUNDLED_COMM_KEYS = ['_commBundledMail', '_commBundledOutlook', '_commBundledQr'] as const;

/**
 * Modus eines Events lesen. Erwartet die KLAMMER; auf einem Termin gibt es
 * nichts zu bündeln, deshalb liefert ein Kind-Event immer den Default.
 */
export function bundledCommOf(
  ev: { emailTemplateOverrides?: string; parentEventId?: string } | null | undefined
): BundledComm {
  if (!ev || ev.parentEventId) return BUNDLED_COMM_DEFAULT;
  try {
    const o = JSON.parse(ev.emailTemplateOverrides || '{}') || {};
    return {
      mail: !!o._commBundledMail,
      outlook: !!o._commBundledOutlook,
      qr: !!o._commBundledQr,
    };
  } catch { return BUNDLED_COMM_DEFAULT; }
}

/** Als Piggyback-Objekt zum Mergen in die Overrides. Aus-Schalter werden
 *  NICHT geschrieben, damit ein Event ohne Bündelung keinen Ballast trägt. */
export function bundledCommConfig(m: BundledComm): Record<string, boolean> {
  const out: Record<string, boolean> = {};
  if (m.mail) out._commBundledMail = true;
  if (m.outlook) out._commBundledOutlook = true;
  if (m.qr) out._commBundledQr = true;
  return out;
}

/**
 * v30.71: „Gemeinsam für alle Termine" — ein Schalter mit Gedächtnis statt des
 * einmaligen Kopier-Knopfs aus v30.60.
 *
 * Nutzer-Ansage 02.09.2026: „kein Button, sondern ein Wechselschalter —
 * entweder einzeln oder für alle Termine zusammen." Der Knopf kopierte einmal;
 * sobald danach jemand am Haupt-Event etwas änderte, liefen die Termine wieder
 * auseinander. Der Schalter bleibt an: Solange er steht, bekommt jeder Termin
 * beim Speichern die Kommunikations-Werte des Haupt-Events (persistSubEvents),
 * und die Termin-Reiter sind im Kommunikations-Schritt nur Anzeige.
 *
 * Etwas ANDERES als die Bündel-Schalter oben: Dort geht es darum, wie OFT eine
 * Mail kommt (eine fürs Gesamt-Event statt einer je Termin), hier darum, WAS
 * drinsteht. Beides lebt als Piggyback in den Overrides der Klammer.
 *
 * Voreinstellung (Nutzer-Entscheidung): neue Events GEMEINSAM, bestehende
 * Events EINZELN — damit sich an laufenden Events nichts ändert.
 */
export const COMM_SHARED_KEY = '_commShared';

export function commSharedOf(
  ev: { emailTemplateOverrides?: string; parentEventId?: string } | null | undefined
): boolean {
  if (!ev || ev.parentEventId) return false;
  try {
    const o = JSON.parse(ev.emailTemplateOverrides || '{}') || {};
    return o[COMM_SHARED_KEY] === true;
  } catch { return false; }
}

/** Nur „an" wird geschrieben — ein Event im Einzel-Modus trägt keinen Ballast. */
export function commSharedConfig(on: boolean): Record<string, boolean> {
  return on ? { [COMM_SHARED_KEY]: true } : {};
}

/** Ein gebuchter Termin, wie er in der Sammelmail steht. */
export interface BundledItem {
  title: string;
  startDate?: string;
  endDate?: string;
  location?: string;
}

const TD = 'padding:7px 12px 7px 0;border-bottom:1px solid #eee;font-size:14px;vertical-align:top;';

function esc(s: string): string {
  return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function fmtRange(startIso?: string, endIso?: string): string {
  const s = new Date(startIso || '');
  if (!isFinite(s.getTime())) return '';
  const d = s.toLocaleDateString('de-DE', { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' });
  const e = new Date(endIso || '');
  // Ganztägig wird als 00:00–23:59 gespeichert (v28.91) — die Uhrzeit dann
  // wegzulassen ist keine Auslassung, sondern die Aussage.
  const allDay = s.getHours() === 0 && s.getMinutes() === 0
    && isFinite(e.getTime()) && e.getHours() === 23 && e.getMinutes() >= 58;
  if (allDay) return d;
  const t1 = s.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit', hour12: false });
  if (!isFinite(e.getTime())) return `${d}, ${t1}`;
  const t2 = e.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit', hour12: false });
  const sameDay = s.toDateString() === e.toDateString();
  return sameDay ? `${d}, ${t1}–${t2} Uhr` : `${d}, ${t1} bis ${e.toLocaleDateString('de-DE')}, ${t2}`;
}

/**
 * Die Terminliste für die Sammelmail.
 *
 * Ohne sie wäre die Bündelung ein Rückschritt: Zehn Einzelmails sagen wenigstens,
 * WOFÜR man angemeldet ist. Eine Sammelmail ohne Liste sagt das nicht mehr.
 * Sortiert nach Startzeit, weil die Mail wie ein Programm gelesen wird.
 */
export function bundledItemsTableHtml(items: BundledItem[], isDe: boolean): string {
  const rows = [...(items || [])]
    .sort((a, b) => new Date(a.startDate || 0).getTime() - new Date(b.startDate || 0).getTime())
    .map(it => {
      const when = fmtRange(it.startDate, it.endDate);
      return `<tr>`
        + `<td style="${TD}font-weight:600;">${esc(it.title || '—')}</td>`
        + `<td style="${TD}white-space:nowrap;color:#444;">${esc(when)}</td>`
        + `<td style="${TD}color:#444;">${esc(it.location || '')}</td>`
        + `</tr>`;
    }).join('');
  if (!rows) return '';
  const head = isDe
    ? ['Termin', 'Wann', 'Wo']
    : ['Session', 'When', 'Where'];
  return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;margin:18px 0;">`
    + `<tr>${head.map(h => `<th style="text-align:left;padding:0 12px 6px 0;border-bottom:2px solid #86bc25;font-size:12px;text-transform:uppercase;letter-spacing:0.04em;color:#63666A;">${h}</th>`).join('')}</tr>`
    + rows
    + `</table>`;
}

/** Überschrift über der Liste — nennt die Anzahl, damit Unvollständigkeit auffällt. */
export function bundledItemsHeading(n: number, isDe: boolean, childTermPlural?: string): string {
  const term = (childTermPlural || (isDe ? 'Termine' : 'sessions')).trim();
  if (isDe) return n === 1 ? 'Dein gebuchter Termin:' : `Deine ${n} gebuchten ${term}:`;
  return n === 1 ? 'Your booked session:' : `Your ${n} booked ${term}:`;
}
