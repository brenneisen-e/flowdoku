/**
 * v30.5: F&A-Abrechnung — gemeinsame Logik für Wizard (Schritt 10),
 * Organizer Center (Aktion „Event-Abrechnung") und das F&A Center.
 *
 * Alles hängt am Piggyback `_billing` in `EmailTemplateOverrides` des
 * HAUPTEVENTS (Klammer/Einzel-Event — Sub-Events tragen keine eigene
 * Abrechnung). Der Wizard schreibt relevant/sendMode/fields; alle übrigen
 * Schlüssel (Log, Versand-Stempel, Snapshots, Abschluss) werden von den
 * F&A-Flows über `patchEventOverridesValue` gepflegt und vom Wizard beim
 * Speichern ERHALTEN (billingExtraRef in EventCreationPage) — sonst würde
 * jeder Wizard-Save die Historie löschen.
 */

import { DeloitteEvent } from '../types';
import { BILLING_FIELDS, BillingFieldDef } from '../data/billingFields';

/** Ein Eintrag der revisionssicheren Historie (Fachkonzept Abschnitt 13). */
export interface BillingLogEntry {
  ts: string;          // ISO-Zeitpunkt
  by: string;          // Anzeigename oder E-Mail des Auslösers ('System' bei Automatik)
  action: string;      // Klartext, z.B. 'Abrechnungsinformationen an F&A versendet'
  mailType?: 'info' | 'list';
  to?: string;
  cc?: string;
  subject?: string;
  /** Vollständiger HTML-Inhalt der Mail (Konzept: „Anzeige des vollständigen
   *  E-Mail-Inhalts"). Nur bei Versand-Einträgen gesetzt. */
  body?: string;
  old?: unknown;
  neu?: unknown;
}

export interface BillingData {
  relevant: boolean;
  sendMode: 'auto' | 'manual';
  fields: Record<string, string>;
  log?: BillingLogEntry[];
  /** Abschluss durch F&A („Als abgerechnet markieren") — bleibt dauerhaft. */
  settled?: { ts: string; by: string };
  /** Letzter Versand der Abrechnungsinfos / Teilnehmerliste (manuell ODER auto). */
  infoSentAt?: string;
  listSentAt?: string;
  /** Doppelversand-Schutz des Automatik-Jobs (zusätzlich zur DEX_Emails-Prüfung). */
  autoInfoSentAt?: string;
  autoListSentAt?: string;
  /** Einmal-Erinnerung an die Organizer, wenn der Auto-Versand an fehlenden Feldern scheitert. */
  autoInfoReminderAt?: string;
  /** Zuletzt an F&A übermittelter Stand (Detailansicht zeigt NUR Übermitteltes). */
  infoSnapshot?: Record<string, string>;
  listSnapshot?: Array<{ name: string; email: string; status: string }>;
}

/** F&A-Verteiler + Änderungsprotokoll (persistiert als eigene Zeile in DEX_EmailTemplates). */
export interface FAConfig {
  infoRecipients: string[];
  listRecipients: string[];
  log: Array<{ ts: string; by: string; action: string; old?: string; neu?: string }>;
}

export function parseBillingOf(ev: Pick<DeloitteEvent, 'emailTemplateOverrides'> | null | undefined): BillingData | null {
  try {
    const b = JSON.parse(ev?.emailTemplateOverrides || '{}')._billing;
    if (b && typeof b === 'object' && typeof b.relevant === 'boolean') return b as BillingData;
  } catch { /* kein Blob — keine Abrechnung */ }
  return null;
}

export function missingBillingFields(b: BillingData | null): BillingFieldDef[] {
  const fields = (b && b.fields) || {};
  return BILLING_FIELDS.filter(f => !(fields[f.id] || '').trim());
}

/**
 * Statusmodell — v30.47 auf das Fachkonzept („Statusmodell (Erweiterung)")
 * gebracht. Ausschließlich systemseitig abgeleitet: Es gibt keinen Ort, an
 * dem jemand einen Status von Hand setzt. Ableiten statt speichern ist hier
 * Absicht — ein gespeicherter Status wäre eine zweite Wahrheit neben den
 * Daten, und die beiden laufen irgendwann auseinander (Feld nachträglich
 * geleert, Versand wiederholt, Eventdatum verschoben).
 *
 * **Was sich gegenüber v30.5 ändert und warum es wichtig ist:** Der frühere
 * Zustand `upcoming` fasste ZWEI Dinge zusammen, die das Konzept ausdrücklich
 * trennt — „alle Pflichtfelder gepflegt, aber noch nichts an F&A geschickt"
 * und „an F&A geschickt, Event steht noch aus". Für den Organizer sind das
 * gegensätzliche Aussagen: Beim einen ist er dran, beim anderen wartet er.
 * Genau deshalb steht `readyToSend` jetzt eigenständig da; `infoSentAt` lag
 * im Datensatz längst vor, es wurde nur nie ausgewertet.
 *
 * Die fünf Stufen des Konzepts:
 *
 *   1 incomplete    Mindestens ein Pflichtfeld fehlt → Versand nicht möglich
 *   2 readyToSend   Alles gepflegt, noch nicht an F&A versendet
 *   3 infoSent      An F&A versendet, Event noch nicht stattgefunden
 *   4 listPending   Event vorbei, Teilnehmerliste noch nicht versendet
 *   5 settled       Abgerechnet
 *
 * Dazu EIN zusätzlicher Zwischenzustand, den das Konzept nicht benennt:
 * `sentAwaitSettle` = Teilnehmerliste übermittelt, Abschluss durch F&A noch
 * offen. Er bleibt bewusst erhalten, weil Stufe 5 laut Konzept BEIDES
 * verlangt („Die Teilnehmerliste wurde erfolgreich an F&A versendet" UND
 * „Der Abrechnungsprozess wurde durch F&A abgeschlossen"). Ohne ihn stünde
 * „Abgerechnet" an Events, die F&A nie bestätigt hat — die Tabelle behauptete
 * einen Abschluss, den es nicht gibt.
 */
export type FAStatus =
  | 'incomplete'
  | 'readyToSend'
  | 'infoSent'
  | 'listPending'
  | 'sentAwaitSettle'
  | 'settled';

/** Reihenfolge der Statusübergänge — für die Fortschritts-Anzeige. */
export const FA_STATUS_ORDER: FAStatus[] = [
  'incomplete', 'readyToSend', 'infoSent', 'listPending', 'sentAwaitSettle', 'settled',
];

export function faStatusOf(ev: DeloitteEvent, b?: BillingData | null): FAStatus {
  const billing = b === undefined ? parseBillingOf(ev) : b;
  if (!billing) return 'incomplete';
  // Der Abschluss durch F&A steht über allem: Ist er gesetzt, bleibt er —
  // auch wenn danach noch jemand ein Feld anfasst.
  if (billing.settled) return 'settled';
  if (missingBillingFields(billing).length > 0) return 'incomplete';
  // Konzept, Stufe 2: „Der Status bleibt bestehen, bis die Abrechnungs-
  // informationen ERSTMALIG an F&A versendet wurden." Das gilt auch dann,
  // wenn das Event schon vorbei ist — dann ist der Info-Versand der
  // blockierende Schritt und nicht die Teilnehmerliste.
  if (!billing.infoSentAt) return 'readyToSend';
  const end = new Date(ev.endDate || ev.startDate || '');
  const over = isFinite(end.getTime()) && end.getTime() < Date.now();
  if (!over) return 'infoSent';
  // Konzept, Stufe 4: bleibt bestehen, bis die Teilnehmerliste versendet ist —
  // manuell ODER automatisch. Beide Wege stempeln `listSentAt`.
  return billing.listSentAt ? 'sentAwaitSettle' : 'listPending';
}

/** Beschriftungen wörtlich aus dem Fachkonzept. */
export const FA_STATUS_LABELS: Record<FAStatus, string> = {
  incomplete: 'Abrechnungsrelevante Informationen unvollständig',
  readyToSend: 'Abrechnungsrelevante Informationen vollständig, Versendung ausstehend',
  infoSent: 'Abrechnungsrelevante Informationen versendet, Event ausstehend',
  listPending: 'Event stattgefunden, Teilnehmerlistenversand ausstehend',
  sentAwaitSettle: 'Teilnehmerliste versendet — Abschluss durch F&A offen',
  settled: 'Abgerechnet',
};

/** Kurzform für enge Stellen (Tabellen-Spalte, Kachel). */
export const FA_STATUS_SHORT: Record<FAStatus, string> = {
  incomplete: 'Angaben unvollständig',
  readyToSend: 'Versand ausstehend',
  infoSent: 'Event ausstehend',
  listPending: 'Teilnehmerliste ausstehend',
  sentAwaitSettle: 'Abschluss durch F&A offen',
  settled: 'Abgerechnet',
};

/** Was ist als NÄCHSTES zu tun? Leer, wenn nichts offen ist. */
export const FA_STATUS_NEXT: Record<FAStatus, string> = {
  incomplete: 'Fehlende Pflichtangaben ergänzen — erst danach ist ein Versand möglich.',
  readyToSend: 'Abrechnungsinformationen an F&A versenden.',
  infoSent: 'Nichts zu tun — die Teilnehmerliste geht nach dem Event.',
  listPending: 'Teilnehmerliste an F&A versenden.',
  sentAwaitSettle: 'Nichts zu tun — F&A schließt die Abrechnung ab.',
  settled: '',
};

export const FA_STATUS_COLORS: Record<FAStatus, { bg: string; fg: string }> = {
  incomplete: { bg: 'rgba(218,41,28,0.12)', fg: '#b02318' },
  readyToSend: { bg: 'rgba(237,139,0,0.15)', fg: '#b86700' },
  infoSent: { bg: 'rgba(0,118,168,0.12)', fg: '#0076a8' },
  listPending: { bg: 'rgba(237,139,0,0.15)', fg: '#b86700' },
  sentAwaitSettle: { bg: 'rgba(134,188,37,0.15)', fg: '#4a7c1f' },
  settled: { bg: 'rgba(134,188,37,0.25)', fg: '#2e7d32' },
};

const esc = (s: string): string =>
  String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const fmtDateTime = (iso: string): string => {
  const d = new Date(iso || '');
  return isFinite(d.getTime())
    ? d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ' ' +
      d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
    : '—';
};

const MAIL_WRAP_START = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333;line-height:1.5;">';
const TABLE_STYLE = 'border-collapse:collapse;width:100%;max-width:640px;';
const TH_STYLE = 'text-align:left;padding:6px 10px;border:1px solid #ddd;background:#f4f7ee;font-weight:600;';
const TD_STYLE = 'padding:6px 10px;border:1px solid #ddd;vertical-align:top;';

/** HTML-Body der Mail „Abrechnungsinformationen an F&A" — Eventkopf + alle elf Felder. */
export function renderBillingInfoMailBody(ev: DeloitteEvent, b: BillingData, byName: string, faCenterUrl?: string): string {
  const rows = BILLING_FIELDS.map(f =>
    `<tr><td style="${TD_STYLE}">${esc(f.label)}</td><td style="${TD_STYLE}">${esc((b.fields || {})[f.id] || '—')}</td></tr>`
  ).join('');
  return `${MAIL_WRAP_START}
<p>Guten Tag,</p>
<p>zur Veranstaltung <strong>${esc(ev.title)}</strong> (Event-ID ${esc(String(ev.eventNumber || ev.id))}) übermitteln wir die abrechnungsrelevanten Informationen:</p>
<table style="${TABLE_STYLE}">
<tr><th style="${TH_STYLE}" colspan="2">Eventinformationen</th></tr>
<tr><td style="${TD_STYLE}">Eventname</td><td style="${TD_STYLE}">${esc(ev.title)}</td></tr>
<tr><td style="${TD_STYLE}">Eventdatum</td><td style="${TD_STYLE}">${fmtDateTime(ev.startDate)}${ev.endDate && ev.endDate !== ev.startDate ? ' – ' + fmtDateTime(ev.endDate) : ''}</td></tr>
<tr><td style="${TD_STYLE}">Ort</td><td style="${TD_STYLE}">${esc(ev.location || '—')}</td></tr>
<tr><td style="${TD_STYLE}">Organizer</td><td style="${TD_STYLE}">${esc((ev.organizers || []).join('; ') || '—')}</td></tr>
<tr><th style="${TH_STYLE}" colspan="2">Abrechnungsrelevante Informationen</th></tr>
${rows}
</table>
${faCenterUrl ? `${mailButton(faCenterUrl, 'Event im F&A Center öffnen')}
<p style="color:#666;font-size:12px;margin-top:-8px;">Dort findest du alle übermittelten Stände zu diesem Event, die Teilnehmerliste zum Download und die Möglichkeit, das Event als abgerechnet zu markieren.</p>` : ''}
<p style="color:#666;font-size:12px;">Ausgelöst von ${esc(byName)} über die DEX Event Experience Platform.</p>
</div>`;
}

/** HTML-Body der Mail „Teilnehmerliste an F&A". */
/**
 * v30.24: Download-Schaltfläche für die Mail.
 *
 * Statt eines Datei-Anhangs (im Deloitte-Tenant unmöglich — jede Power-
 * Automate-Mail MIT Anhang wird per NDR abgewiesen, s. v26.71) trägt die
 * Mail einen Deep-Link ins F&A Center. Dort steht der versendete Stand als
 * Tabelle und lässt sich mit einem Klick als Excel herunterladen. Vorteil
 * gegenüber dem Anhang: Die personenbezogene Liste liegt nicht in
 * Postfächern herum, sondern bleibt in DEX hinter der F&A-Rolle.
 */
function mailButton(url: string, label: string): string {
  return `<p style="margin:18px 0;">
<a href="${esc(url)}" style="display:inline-block;background:#86bc25;color:#ffffff;padding:11px 22px;border-radius:6px;text-decoration:none;font-weight:600;font-size:14px;">${esc(label)}</a>
</p>`;
}

export function renderBillingListMailBody(
  ev: DeloitteEvent,
  participants: Array<{ name: string; email: string; status: string }>,
  byName: string,
  /** v30.24: Deep-Link ins F&A Center (Download der Liste als Excel). */
  faCenterUrl?: string
): string {
  const rows = participants.map((p, i) =>
    `<tr><td style="${TD_STYLE}">${i + 1}</td><td style="${TD_STYLE}">${esc(p.name)}</td><td style="${TD_STYLE}">${esc(p.email)}</td><td style="${TD_STYLE}">${esc(p.status)}</td></tr>`
  ).join('');
  const linkBlock = faCenterUrl
    ? `${mailButton(faCenterUrl, 'Teilnehmerliste als Excel herunterladen')}
<p style="color:#666;font-size:12px;margin-top:-8px;">Der Link öffnet dieses Event im F&amp;A Center der DEX-App — dort lädst du genau diesen Stand als Excel-Datei herunter. Dafür brauchst du die Rolle „F&amp;A" (oder Admin) in DEX; falls der Zugriff fehlt, melde dich bei dex.event@deloitte.de.</p>`
    : '';
  return `${MAIL_WRAP_START}
<p>Guten Tag,</p>
<p>anbei die Teilnehmerliste zur Veranstaltung <strong>${esc(ev.title)}</strong> (Event-ID ${esc(String(ev.eventNumber || ev.id))}, ${fmtDateTime(ev.startDate)}) — <strong>${participants.length}</strong> ${participants.length === 1 ? 'Person' : 'Personen'}:</p>
${linkBlock}
<table style="${TABLE_STYLE}">
<tr><th style="${TH_STYLE}">#</th><th style="${TH_STYLE}">Name</th><th style="${TH_STYLE}">E-Mail</th><th style="${TH_STYLE}">Status</th></tr>
${rows}
</table>
<p style="color:#666;font-size:12px;">Ausgelöst von ${esc(byName)} über die DEX Event Experience Platform.</p>
</div>`;
}

/** Historie kompakt halten: Mail-Bodys sind die größten Brocken — nur die
 *  letzten 15 Einträge behalten ihren Body, ältere nur die Metadaten. Das
 *  Feld EmailTemplateOverrides trägt sonst irgendwann das 2-MB-Limit. */
export function trimBillingLog(log: BillingLogEntry[]): BillingLogEntry[] {
  const capped = log.slice(-60);
  const withBody = capped.filter(e => !!e.body);
  const dropBodies = Math.max(0, withBody.length - 15);
  let dropped = 0;
  return capped.map(e => {
    if (e.body && dropped < dropBodies) { dropped++; return { ...e, body: undefined }; }
    return e;
  });
}
