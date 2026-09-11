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
  listSnapshot?: FAListRow[];
}

/**
 * Eine Zeile der an F&A übermittelten Teilnehmerliste.
 *
 * v30.50: `firstName`/`lastName`/`country`/`company` sind neu und bewusst
 * OPTIONAL — Snapshots aus früheren Versendungen tragen nur die ersten drei
 * Felder. Pflicht wäre hier ein Typfehler an jedem Altbestand und eine leere
 * Detailansicht für alles, was vor diesem Release gemeldet wurde.
 */
export interface FAListRow {
  name: string;
  email: string;
  status: string;
  firstName?: string;
  lastName?: string;
  country?: string;
  company?: string;
  /** v30.60: Von F&A im Center nachgetragen (siehe `activeEmployeesLookupUrl`).
   *  DEX kennt die Personalnummer nicht und kann sie auch nicht ermitteln —
   *  sie steht in einer Backoffice-Liste, auf die nur F&A Zugriff hat. */
  personalNr?: string;
  /** Kostenstelle des Mitarbeiters — dieselbe Herkunft wie `personalNr`. */
  costCenter?: string;
}

/**
 * v30.60: Deep-Link in die Backoffice-Liste „Active Employees", vorbelegt mit
 * dem Nachnamen der Person.
 *
 * Nutzer-Ansage 01.09.2026: Die F&A-Kolleginnen und -Kollegen haben Zugriff
 * auf diese SharePoint-Seite und tragen von dort die Personalnummer je
 * Teilnehmer nach. Der Parameter `k` ist der Suchbegriff der Seite, `ql=1031`
 * die deutsche Oberfläche — beides aus dem vom Nutzer gelieferten Link.
 *
 * Gesucht wird mit dem NACHNAMEN, nicht mit dem vollen Namen: Die Liste
 * durchsucht Mitarbeiterdatensätze, und „Max Mustermann" trifft dort nichts,
 * weil Vor- und Nachname getrennt stehen. Fehlt der Nachname (Alt-Snapshots
 * vor v30.50 tragen ihn nicht), nehmen wir das letzte Wort des Anzeigenamens —
 * eine Näherung, aber eine, die den Klick nicht wertlos macht.
 */
export const FA_ACTIVE_EMPLOYEES_URL =
  'https://teams.de.deloitte.com/sites/dtBackoffice/SitePages/Active%20Employees.aspx';

export function activeEmployeesLookupUrl(row: Pick<FAListRow, 'lastName' | 'name'>): string {
  const fromName = (row.name || '').trim().split(/\s+/).slice(-1)[0] || '';
  const key = (row.lastName || '').trim() || fromName;
  return `${FA_ACTIVE_EMPLOYEES_URL}?k=${encodeURIComponent(key)}&ql=1031`;
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
  participants: FAListRow[],
  byName: string,
  /** v30.24: Deep-Link ins F&A Center (Download der Liste als Excel). */
  faCenterUrl?: string,
  /** v30.50: Abrechnungsdaten — stehen laut Konzept ÜBER den Teilnehmerdaten. */
  b?: BillingData | null
): string {
  const rows = participants.map((p, i) =>
    `<tr><td style="${TD_STYLE}">${i + 1}</td><td style="${TD_STYLE}">${esc(p.name)}</td><td style="${TD_STYLE}">${esc(p.email)}</td><td style="${TD_STYLE}">${esc(p.status)}</td></tr>`
  ).join('');
  // v30.50: „Bei der Versendung der Teilnehmerliste sollen oberhalb der
  // eigentlichen Teilnehmerdaten zusätzlich alle abrechnungsrelevanten
  // Informationen des Events aufgeführt werden." Das gilt für die Datei UND
  // für die Mail: Wer die Mail liest, ohne die Excel zu öffnen, hätte sonst
  // eine Namensliste ohne jeden Bezug zu Ariba-Nummer und WBS-Code.
  const billingBlock = b
    ? `<table style="${TABLE_STYLE}">
<tr><th style="${TH_STYLE}" colspan="2">Abrechnungsrelevante Informationen</th></tr>
${BILLING_FIELDS.map(f =>
      `<tr><td style="${TD_STYLE}">${esc(f.label)}</td><td style="${TD_STYLE}">${esc((b.fields || {})[f.id] || '—')}</td></tr>`
    ).join('')}
</table>`
    : '';
  const linkBlock = faCenterUrl
    ? `${mailButton(faCenterUrl, 'Teilnehmerliste als Excel herunterladen')}
<p style="color:#666;font-size:12px;margin-top:-8px;">Der Link öffnet dieses Event im F&amp;A Center der DEX-App — dort lädst du genau diesen Stand als Excel-Datei herunter. Dafür brauchst du die Rolle „F&amp;A" (oder Admin) in DEX; falls der Zugriff fehlt, melde dich bei dex.event@deloitte.de.</p>`
    : '';
  return `${MAIL_WRAP_START}
<p>Guten Tag,</p>
<p>anbei die Teilnehmerliste zur Veranstaltung <strong>${esc(ev.title)}</strong> (Event-ID ${esc(String(ev.eventNumber || ev.id))}, ${fmtDateTime(ev.startDate)}) — <strong>${participants.length}</strong> ${participants.length === 1 ? 'Person' : 'Personen'}:</p>
${linkBlock}
${billingBlock}
<table style="${TABLE_STYLE}">
<tr><th style="${TH_STYLE}" colspan="4">Teilnehmer</th></tr>
<tr><th style="${TH_STYLE}">#</th><th style="${TH_STYLE}">Name</th><th style="${TH_STYLE}">E-Mail</th><th style="${TH_STYLE}">Status</th></tr>
${rows}
</table>
<p style="color:#666;font-size:12px;">Ausgelöst von ${esc(byName)} über die DEX Event Experience Platform.</p>
</div>`;
}

/**
 * v31.13 — die F&A-Liste als Mail, ausgelöst vom Organizer.
 *
 * Nutzer-Ansage 11.09.2026: Die dritte Export-Ansicht soll allen Organizern
 * offenstehen und „auch in die Mail gepackt werden, mit Danke, dass du das
 * Event mit DEX organisiert hast."
 *
 * Zwei Unterschiede zu `renderBillingListMailBody` (dem offiziellen Versand
 * an die F&A-Verteiler), und beide sind Absicht:
 *
 *  - **Die Liste steht IM Text, nicht als Anhang.** `queueEmail` kann seit
 *    v26.62 einen Anhang mitgeben, aber der DEX_SEND_MAIL-Flow hängt ihn
 *    nicht an — er würde stillschweigend fehlen (v26.71: der Tenant blockt
 *    Anhänge aus Power Automate ohnehin). Eine Tabelle, die ankommt, ist
 *    mehr wert als ein Anhang, der es nicht tut.
 *  - **Kein Stempel, kein Snapshot, kein `_billing`-Eintrag.** Das hier ist
 *    eine Arbeitskopie für den Organizer, keine Übermittlung an F&A. Würde
 *    sie `listSentAt` setzen, zeigte das F&A Center „Teilnehmerliste
 *    versendet", ohne dass F&A je etwas bekommen hat.
 */
export function renderOrganizerFAListMailBody(
  ev: DeloitteEvent,
  participants: FAListRow[],
  byName: string,
  nurEingecheckt: boolean,
  b?: BillingData | null
): string {
  const rows = participants.map((p, i) =>
    `<tr><td style="${TD_STYLE}">${i + 1}</td><td style="${TD_STYLE}">${esc(p.name)}</td><td style="${TD_STYLE}">${esc(p.email)}</td><td style="${TD_STYLE}">${esc(p.status)}</td></tr>`
  ).join('');
  // Welche Auswahl in der Liste steckt, gehört über die Tabelle. „37 Personen"
  // ohne diesen Satz ist für einen Bewirtungsbeleg keine belastbare Angabe.
  const auswahl = nurEingecheckt
    ? 'Enthalten sind <strong>nur eingecheckte Personen</strong> — also die, deren Anwesenheit erfasst wurde.'
    : 'Enthalten sind <strong>alle angemeldeten Personen</strong>, die sich nicht abgemeldet haben (Angemeldet, QR versendet, Eingecheckt). Wartende und Abgemeldete stehen nicht darauf.';
  const billingBlock = b
    ? `<table style="${TABLE_STYLE}">
<tr><th style="${TH_STYLE}" colspan="2">Abrechnungsrelevante Informationen</th></tr>
${BILLING_FIELDS.map(f =>
      `<tr><td style="${TD_STYLE}">${esc(f.label)}</td><td style="${TD_STYLE}">${esc((b.fields || {})[f.id] || '—')}</td></tr>`
    ).join('')}
</table>`
    : '';
  return `${MAIL_WRAP_START}
<p>Hallo,</p>
<p><strong>danke, dass du dein Event mit DEX organisiert hast.</strong> Hier ist die Teilnehmerliste zu
<strong>${esc(ev.title)}</strong> (Event-ID ${esc(String(ev.eventNumber || ev.id))}, ${fmtDateTime(ev.startDate)}) im Aufbau, den F&amp;A einliest —
<strong>${participants.length}</strong> ${participants.length === 1 ? 'Person' : 'Personen'}.</p>
<p style="color:#555;">${auswahl}</p>
${billingBlock}
<table style="${TABLE_STYLE}">
<tr><th style="${TH_STYLE}" colspan="4">Teilnehmer</th></tr>
<tr><th style="${TH_STYLE}">#</th><th style="${TH_STYLE}">Name</th><th style="${TH_STYLE}">E-Mail</th><th style="${TH_STYLE}">Status</th></tr>
${rows}
</table>
<p style="color:#666;font-size:12px;">Dieselbe Liste bekommst du im Organizer Center über &bdquo;Excel-Export&ldquo; als Datei &mdash; dort ist sie die dritte Ansicht &bdquo;F&amp;A-Liste&ldquo;. Ausgelöst von ${esc(byName)} über die DEX Event Experience Platform.</p>
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

/**
 * v30.50: `Name <email>` in seine zwei Teile zerlegen.
 *
 * Das ist das Format, das `UserFieldPicker` überall in der App schreibt —
 * auch in `_billing.fields.contact`. Bisher wurde der Wert nur als Ganzes
 * angezeigt; für die Profil-Darstellung im F&A Center braucht es beide
 * Teile getrennt. Enthält der Wert keine spitzen Klammern, ist er entweder
 * eine nackte Adresse (Alt-Bestand aus der Freitext-Zeit vor v30.45) oder
 * ein reiner Name — beides wird hier unterschieden, statt blind zu raten.
 */
export function parsePersonValue(raw: string | undefined | null): { name: string; email: string } {
  const v = (raw || '').trim();
  if (!v) return { name: '', email: '' };
  const m = v.match(/^(.*?)\s*<([^>]+)>\s*$/);
  if (m) return { name: (m[1] || '').trim(), email: (m[2] || '').trim() };
  if (v.indexOf('@') > 0) return { name: '', email: v };
  return { name: v, email: '' };
}

/**
 * v30.50: Land eines Teilnehmers.
 *
 * **Annahme, bewusst an EINER Stelle:** Es gibt in DEX kein Länderfeld. Das
 * Verzeichnis liefert den Standort teils als `DE - Koeln`, teils nur als
 * `Koeln`. Aus einem vorangestellten Zwei-Buchstaben-Code lesen wir das Land
 * ab; sonst gilt Deutschland — der Pilot läuft auf der deutschen Tenant-Site
 * mit deutschen Standorten. Wenn F&A hier eines Tages echte Länder braucht,
 * gehört ein Feld ins Profil, nicht mehr Rateregeln hierher.
 */
const COUNTRY_BY_CODE: Record<string, string> = {
  DE: 'Deutschland', AT: 'Österreich', CH: 'Schweiz',
  NL: 'Niederlande', BE: 'Belgien', FR: 'Frankreich',
  GB: 'Vereinigtes Königreich', UK: 'Vereinigtes Königreich',
  US: 'USA', PL: 'Polen', ES: 'Spanien', IT: 'Italien',
};
/**
 * v31.14: Der Zwei-Buchstaben-Code, den die F&A-Vorlage in der Spalte
 * `Country` erwartet („DE", „NL"). Bis v31.13 stand dort „Deutschland" —
 * der Abgleich mit der Originalvorlage (11.09.2026) hat es gezeigt.
 *
 * Nimmt entweder den Standort (`DE - Koeln`) oder einen bereits aufgelösten
 * Landesnamen entgegen; Snapshots aus früheren Versendungen tragen den Namen.
 */
export function countryCodeOfLocation(value: string | undefined | null): string {
  const v = (value || '').trim();
  if (!v) return 'DE';
  const m = v.match(/^([A-Za-z]{2})(\s*[-–]\s*|$)/);
  if (m && COUNTRY_BY_CODE[m[1].toUpperCase()]) return m[1].toUpperCase();
  for (const code of Object.keys(COUNTRY_BY_CODE)) {
    if (COUNTRY_BY_CODE[code].toLowerCase() === v.toLowerCase()) return code;
  }
  return 'DE';
}

export function countryOfLocation(location: string | undefined | null): string {
  const loc = (location || '').trim();
  const m = loc.match(/^([A-Za-z]{2})\s*[-–]\s*/);
  if (m) {
    const hit = COUNTRY_BY_CODE[m[1].toUpperCase()];
    if (hit) return hit;
  }
  return 'Deutschland';
}

/** Die Registrierungs-Felder, die eine F&A-Zeile braucht — strukturell getypt,
 *  damit dieses Modul nicht vom EventService abhängt. */
export interface FASourceRegistration {
  ParticipantName?: string;
  ParticipantEmail?: string;
  Vorname?: string;
  Nachname?: string;
  Status?: string;
  Location?: string;
  Company?: string;
}

/**
 * v30.50: EINE Abbildung Registrierung → F&A-Zeile.
 *
 * Es gibt zwei Wege, auf denen dieselbe Liste zu F&A geht (Versand-Snapshot
 * in `sendFAMail`, Direkt-Download im Abrechnungs-Dialog) und einen dritten,
 * der sie wieder anzeigt (F&A Center). Vor v30.50 hatte jeder seine eigene
 * `.map()` — genau die Konstruktion, aus der zwei Ansichten mit
 * unterschiedlichen Zahlen entstehen. Abgemeldete werden hier gefiltert,
 * nicht beim Aufrufer.
 */
export function faRowsFromRegistrations(regs: FASourceRegistration[] | null | undefined): FAListRow[] {
  return (regs || [])
    .filter(r => r.Status !== 'Abgemeldet')
    .map(r => {
      const first = (r.Vorname || '').trim();
      const last = (r.Nachname || '').trim();
      return {
        name: (r.ParticipantName || `${first} ${last}`.trim()) || r.ParticipantEmail || '—',
        email: r.ParticipantEmail || '',
        status: r.Status || '',
        firstName: first,
        lastName: last,
        country: countryOfLocation(r.Location),
        company: (r.Company || '').trim(),
      };
    });
}

/**
 * v30.50: Die Excel-Datei für F&A im geforderten Aufbau (Fachkonzept
 * „Versand der Teilnehmerliste an F&A", 01.09.2026).
 *
 * Zeilen 1–11 tragen die abrechnungsrelevanten Informationen (Label in
 * Spalte A, Wert in Spalte B) — dieselben elf Felder wie `BILLING_FIELDS`,
 * in derselben Reihenfolge wie die Vorlage von F&A. Zeile 12/13 bleiben
 * leer, Zeile 14 ist die Kopfzeile der Teilnehmerliste, ab Zeile 15 stehen
 * die Personen.
 *
 * **Die Beschriftungen sind bewusst die der F&A-Vorlage, nicht unsere
 * Feld-Labels.** F&A liest diese Datei seit Jahren in dieser Form; eine
 * schönere Beschriftung wäre für uns kosmetisch und für den Empfänger eine
 * unbekannte Datei. `Personalnummer` und `kostenstelle des Mitarbeiters`
 * standen DEX ursprünglich nicht zur Verfügung; seit v30.60 kann F&A sie im
 * Center nachtragen (`FAListRow.personalNr`) — nachgetragene Werte landen in
 * der Datei, nicht nachgetragene bleiben leer wie zuvor.
 */
/**
 * v31.14 — die Kopfzeilen EXAKT wie die Vorlage „Attendee_List_Upload_Template".
 *
 * Der Nutzer hat die Originaldatei am 11.09.2026 hochgeladen; der Abgleich mit
 * unserer Ausgabe zeigte vier Abweichungen, und jede davon ist eine, die beim
 * Empfänger auffällt:
 *
 *  1. Die Vorlage hat **neun** Kopfzeilen, unsere elf. `Mice Project Nr` und
 *     `Ariba Bestellnummer` gibt es dort nicht — sie standen bei uns MITTEN
 *     im Block und haben alles darunter um zwei Zeilen verschoben. Die
 *     Teilnehmer-Kopfzeile liegt in der Vorlage auf **Zeile 13**, bei uns lag
 *     sie auf 14. Beide Angaben bleiben erhalten, aber auf dem Infoblatt.
 *  2. Zeile 9 heißt „…kurze Info **zur** Bewirtung…", nicht „zum Anlass der
 *     Bewirtung" (letzteres war die Formulierung aus `BILLING_FIELDS`).
 *  3. Zeile 11 der Vorlage trägt den Satz, an WEN die Liste geht
 *     (attendancelist@deloitte.de). Ohne ihn ist die Datei für den, der sie
 *     bekommt, eine Datei ohne Adresse.
 *  4. `Country` ist in der Vorlage ein **Zwei-Buchstaben-Code** („DE", „NL"),
 *     bei uns stand „Deutschland".
 */
export const FA_TEMPLATE_FIELDS: Array<{ id: string; label: string }> = [
  { id: 'contact', label: 'Kontaktperson (für etwaige Rückfragen): ' },
  { id: 'docNo', label: 'Documenten Nr ( sh Swift Launchpad):' },
  { id: 'vendor', label: 'Lieferantenname:' },
  { id: 'company', label: 'Gesellschaft, die die Rechnung erhalten hat:' },
  { id: 'category', label: 'Arbeitsessen/Belohnungsessen/Sonstiges oder Geschenk' },
  { id: 'date', label: 'Veranstaltungs-bzw. Bewirtungsdatum:' },
  { id: 'place', label: 'Ort der Bewirtung/Veranstaltung:' },
  { id: 'wbs', label: 'WBS Code / Kostenstelle:' },
  { id: 'name', label: 'Name der Veranstaltung bzw. kurze Info zur Bewirtung oder zum Geschenk:' },
];

/** Die Adresse aus der Vorlage — Zeile 11 des Blattes. */
export const FA_SEND_TO = 'Attendancelist@deloitte.de';
export const FA_SEND_TO_LINE = `Versand der Liste bitte per e-mail an: ${FA_SEND_TO}`;

/**
 * Die fünf Werte der Spalte „Participent Type" samt ihrer Bedeutung — wörtlich
 * aus dem Blatt „Wichtige Info" der Vorlage.
 */
export const FA_PARTICIPANT_TYPES: Array<{ wert: string; bedeutung: string }> = [
  { wert: 'Employee', bedeutung: 'Arbeitnehmer der Gesellschaft die die Rechnung empfangen hat' },
  { wert: 'Customer', bedeutung: 'Geschäftspartner' },
  { wert: 'Intercompany', bedeutung: 'Arbeitnehmer anderer Deloitte Gesellschaften in Deutschland' },
  { wert: 'Interfirm', bedeutung: 'Arbeitnehmer von Deloitte Gesellschaften aus dem Ausland (z. B. USA, NL, GB, Italien)' },
  { wert: 'Official', bedeutung: 'Amtsträger' },
];

/**
 * v31.14: Spalte A wird GEFÜLLT — sie blieb bis v31.13 leer (gemeldet
 * 11.09.2026: „die erste Spalte wurde nicht sauber gefüllt").
 *
 * v30.67 hatte sie bewusst leer gelassen mit der Begründung, ein erfundener
 * Wert wäre für F&A ein Wert. Die Begründung war richtig, die Folge falsch:
 * Eine Pflichtspalte leer zu lassen heißt, dass jemand sie für 85 Zeilen von
 * Hand nachträgt — und dabei rät er genauso, nur ohne die Information, die
 * DEX hat. Entschieden wird jetzt an der Adresse, nicht geraten:
 *
 *  - `…@deloitte.de` → `Employee` (Arbeitnehmer). Der Normalfall bei einem
 *    internen Event auf einer Deloitte-Site.
 *  - jede andere Domain → `Customer` (Geschäftspartner) — genau das Wort,
 *    das die Vorlage für „extern" vorsieht.
 *
 * `Intercompany`, `Interfirm` und `Official` kann DEX NICHT unterscheiden:
 * Dafür müsste es die rechnungsempfangende Gesellschaft mit der Gesellschaft
 * der Person vergleichen, und die Profilwerte („Deloitte", „Deloitte
 * Consulting") passen auf keine Rechnungsgesellschaft („Deloitte GmbH").
 * Deshalb steht die vollständige Auswahl mit ihrer Bedeutung auf dem zweiten
 * Blatt der Datei — überschreiben ist ein Tippen, Raten wäre eine Recherche.
 */
export function participantTypeOf(email: string | undefined | null): string {
  const e = (email || '').trim().toLowerCase();
  if (!e) return '';
  return /@deloitte\.[a-z.]+$/.test(e) ? 'Employee' : 'Customer';
}

export const FA_SHEET_PARTICIPANT_HEADERS = [
  'Participent Type', 'Email', 'First Name', 'Last Name',
  'Country', 'Company Name', 'Personalnummer', 'kostenstelle des Mitarbeiters',
];

/**
 * Die neun Kopf-Werte, sinnvoll vorbelegt.
 *
 * v31.14 (Nutzer-Ansage 11.09.2026: „der Header soll auch abgefragt werden
 * bei F&A-Export und sinnvoll mit Vorschlägen vorbefüllt sein im Modal").
 * Was in `_billing` gepflegt ist, gewinnt — das hat jemand bewusst
 * eingetragen. Ist es leer, nimmt DEX, was es über das Event ohnehin weiß:
 * Datum, Ort und Titel standen bisher leer in der Datei, obwohl sie zwei
 * Felder weiter im selben Event stehen.
 */
export function faHeaderDefaults(
  ev: DeloitteEvent,
  b: BillingData | null | undefined,
  me?: { name?: string; email?: string },
): Record<string, string> {
  const f = (b && b.fields) || {};
  const wert = (id: string): string => (f[id] || '').trim();

  // Die Kontaktperson steht als `Name <email>` im Datensatz. In der Datei für
  // F&A gehört der Name nach vorn und die Adresse in Klammern dahinter — eine
  // rohe spitze Klammer liest dort niemand.
  let contact = wert('contact');
  if (contact) {
    const p = parsePersonValue(contact);
    contact = p.name && p.email ? `${p.name} (${p.email})` : (p.name || p.email);
  } else if (me && (me.name || me.email)) {
    contact = me.name && me.email ? `${me.name} (${me.email})` : (me.name || me.email || '');
  }

  const datum = wert('date') || (ev.startDate
    ? new Date(ev.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })
    : '');

  return {
    contact,
    docNo: wert('docNo'),
    vendor: wert('vendor'),
    company: wert('company'),
    category: wert('category'),
    date: datum,
    place: wert('place') || (ev.location || '').trim(),
    wbs: wert('wbs'),
    name: wert('name') || (ev.title || '').trim(),
    // Nicht im Vorlagen-Blatt, aber von DEX erfasst — landet auf dem Infoblatt.
    mice: wert('mice'),
    ariba: wert('ariba'),
  };
}

/**
 * Das Teilnehmer-Blatt im Aufbau der Vorlage.
 *
 * `header` überschreibt die Vorbelegung — so trägt die Datei genau das, was
 * im Dialog stand. Ohne `header` greift `faHeaderDefaults`.
 */
export function buildFASheetAoa(
  ev: DeloitteEvent,
  b: BillingData | null | undefined,
  rows: FAListRow[],
  header?: Record<string, string>,
): string[][] {
  const werte = header || faHeaderDefaults(ev, b);
  const aoa: string[][] = [];
  for (const def of FA_TEMPLATE_FIELDS) {
    // Fällt ein Feld leer aus, bleibt die Zelle leer statt „—": Die Datei
    // wird von F&A weiterverarbeitet, und ein Gedankenstrich ist dort ein
    // Wert, kein fehlender Wert.
    aoa.push([def.label, (werte[def.id] || '').trim()]);
  }
  aoa.push([]);                     // Zeile 10 — leer, wie in der Vorlage
  aoa.push([FA_SEND_TO_LINE]);      // Zeile 11
  aoa.push([]);                     // Zeile 12
  aoa.push(FA_SHEET_PARTICIPANT_HEADERS.slice()); // Zeile 13
  for (const r of rows) {
    // v30.67 (Review): Snapshots vor v30.50 tragen nur `name`, kein
    // firstName/lastName. Seit Spalte A leer bleibt, stand für diese Zeilen
    // GAR KEIN Name mehr in der Datei — nur die E-Mail. Deshalb den Namen
    // in First/Last Name zerlegen, wenn die getrennten Felder fehlen:
    // letztes Wort = Nachname, dieselbe Näherung wie activeEmployeesLookupUrl.
    const hasSplitName = !!((r.firstName || '').trim() || (r.lastName || '').trim());
    const nameParts = (r.name || '').trim().split(/\s+/).filter(Boolean);
    const firstName = hasSplitName ? (r.firstName || '') : nameParts.slice(0, -1).join(' ');
    const lastName = hasSplitName ? (r.lastName || '') : (nameParts.slice(-1)[0] || '');
    aoa.push([
      participantTypeOf(r.email),
      r.email || '',
      firstName,
      lastName,
      countryCodeOfLocation(r.country),
      r.company || '',
      // v30.60: Nicht mehr grundsätzlich leer. Trägt F&A die Nummer im Center
      // nach, steht sie hier — sonst bleibt die Zelle leer wie bisher und
      // wird in der Datei ergänzt. Ein „—" wäre für F&A ein Wert.
      r.personalNr || '',
      r.costCenter || '',
    ]);
  }
  void ev;
  return aoa;
}

/**
 * Das zweite Blatt: die Auswahlwerte der Spalte A und die Angaben, die die
 * Vorlage nicht kennt (MICE, Ariba).
 *
 * **Warum kein echtes Excel-Dropdown:** Die Datenprüfung („Data Validation")
 * ist in der eingesetzten Bibliothek (SheetJS Community Edition) nicht
 * schreibbar. Ein Dropdown lässt sich mit ihr schlicht nicht erzeugen; was
 * geht, ist die Liste beizulegen, damit niemand raten muss, was in Spalte A
 * erlaubt ist. Das steht hier, damit es beim nächsten Mal niemand erneut
 * ausprobiert.
 */
export function buildFAInfoSheetAoa(werte: Record<string, string>): string[][] {
  const aoa: string[][] = [
    [FA_SEND_TO_LINE],
    [],
    ['Spalte „Participent Type" — erlaubte Werte'],
    ['Wert', 'Bedeutung'],
  ];
  for (const t of FA_PARTICIPANT_TYPES) aoa.push([t.wert, t.bedeutung]);
  aoa.push([]);
  aoa.push(['DEX füllt die Spalte anhand der E-Mail-Adresse vor: @deloitte.de = Employee, alles andere = Customer.']);
  aoa.push(['Intercompany, Interfirm und Official kann DEX nicht unterscheiden — bitte dort von Hand überschreiben.']);
  aoa.push([]);
  aoa.push(['Weitere Angaben aus DEX (nicht Teil der Vorlage)']);
  aoa.push(['Mice Project Nr', (werte.mice || '').trim()]);
  aoa.push(['Ariba Bestellnummer', (werte.ariba || '').trim()]);
  return aoa;
}

/**
 * v30.50: Die F&A-Datei erzeugen und herunterladen — EINE Implementierung
 * für beide Einstiege (Abrechnungs-Dialog im Organizer Center, F&A Center).
 *
 * Der Anker-Download statt `saveAs` ist Absicht: Im SPFx-Iframe ist der
 * Datei-Dialog der Bibliothek häufig blockiert.
 */
export async function downloadFAParticipantXlsx(
  ev: DeloitteEvent,
  b: BillingData | null | undefined,
  rows: FAListRow[],
  sentAt?: string
): Promise<void> {
  const aoa = buildFASheetAoa(ev, b, rows);
  const safeName = (ev.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
  const stamp = (sentAt || new Date().toISOString()).slice(0, 10);
  // xlsx erst beim Klick nachladen (schwerste Dependency der App).
  const XLSX = await import('xlsx');
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (ws as any)['!cols'] = [
    { wch: 62 }, { wch: 34 }, { wch: 18 }, { wch: 18 },
    { wch: 16 }, { wch: 26 }, { wch: 16 }, { wch: 26 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Teilnehmende_Empfängerliste');
  // v31.14: Das zweite Blatt gehoert zur Datei, nicht zu einem der beiden
  // Wege — sonst haette der Download aus dem F&A Center die Auswahlwerte der
  // Spalte A und der Export aus dem Organizer Center nicht (oder umgekehrt).
  const wsInfo = XLSX.utils.aoa_to_sheet(buildFAInfoSheetAoa(faHeaderDefaults(ev, b)));
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (wsInfo as any)['!cols'] = [{ wch: 28 }, { wch: 80 }];
  XLSX.utils.book_append_sheet(wb, wsInfo, 'Wichtige Info');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Teilnehmerliste_FA_${safeName}_${stamp}.xlsx`;
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
}
