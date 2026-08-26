/**
 * v30.6 — Modularisierung Stufe 2, Tranche 2: Thema „DEX_Outlook-Queue +
 * Graph-Abfragen" (Einladungen, Declines, Konto-/Company-Checks).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs. Der
 * Session-Merker _outlookLocationBackfilled bleibt als public-Feld an der
 * Service-Instanz.
 */

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { buildOutlookLocation } from '../../utils/eventFormat';
import type { EventService, SPRegistration, DeclineCheckResult } from '../EventService';
import { REG_LIST_NAME } from '../EventService';


/**
 * Outlook-Termin-Queue-Liste erstellen falls nicht vorhanden.
 * Power Automate reagiert auf neue Einträge und lädt Teilnehmer
 * zum Outlook-Termin ein oder aus. Der Flow holt sich alle Event-Details
 * (Titel, Datum, Ort, CalendarLink) aus der DEX_Events-Liste via EventId.
 *
 * Spalten:
 * - Title: Kurzbeschreibung (z.B. "Einladung: B2Run")
 * - Attendee: E-Mail-Adresse des Teilnehmers
 * - EventId: ID des Events in DEX_Events (Referenz)
 * - ActionType: Einladen, Ausladen
 * - Status: Pending, Sent, Failed
 * - SentDate: Wann wurde die Aktion ausgeführt
 */
export async function ensureOutlookList(svc: EventService): Promise<void> {
  const listName = 'DEX_Outlook';
  const exists = await svc.listExists(listName);
  if (exists) return;

  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Outlook-Termin-Queue: Power Automate lädt Teilnehmer ein/aus',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });

  const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
    { title: 'Attendee', type: 2 },
    { title: 'EventId', type: 2 },
    // ActionType:
    //  - Einladen / Ausladen: einzelnen Attendee zum Outlook-Termin hinzufügen/entfernen
    //  - UpdateEvent: Titel/Start/Ende aktualisieren (kein Attendee)
    //  - DeleteEvent: kompletten Kalender-Termin löschen (wird beim Löschen eines Events
    //    aus der App abgesetzt, inkl. CalendarLink damit der Flow nicht auf DEX_Events
    //    angewiesen ist - das Event-Item wird direkt danach aus DEX_Events gelöscht).
    { title: 'ActionType', type: 6, choices: ['Einladen', 'Ausladen', 'UpdateEvent', 'DeleteEvent'], metaType: 'SP.FieldChoice' },
    { title: 'Status', type: 6, choices: ['Pending', 'Sent', 'Failed'], metaType: 'SP.FieldChoice' },
    { title: 'SentDate', type: 4 },
    // CalendarLink (iCalUId) - nur für DeleteEvent nötig, damit der Flow das Outlook-
    // Event auch dann noch finden kann, wenn das DEX_Events-Item schon gelöscht wurde.
    { title: 'CalendarLink', type: 3 },
  ];

  for (const f of fields) {
    const payload: Record<string, unknown> = {
      '__metadata': { 'type': f.metaType || 'SP.Field' },
      'Title': f.title,
      'FieldTypeKind': f.type,
      'Required': false,
    };
    if (f.choices) {
      payload['Choices'] = { 'results': f.choices };
    }
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
  }

  await svc.configureDefaultView(listName, [
    'Attendee', 'EventId', 'SubEventId', 'ActionType', 'Status', 'SentDate', 'CalendarLink',
  ]);

  await svc.setQueueListPermissions(listName);
}

/**
 * v18.34: Backfill für Bestands-Events. Stellt sicher, dass das DEX_Events-Item
 * eine gefüllte OutlookLocation hat, BEVOR der Flow den Termin anlegt/aktualisiert.
 * Neue Events bekommen OutlookLocation bereits beim Anlegen/Bearbeiten — alte
 * (vor v18.34 erstellte) Events hätten sonst eine leere Spalte.
 */
async function backfillOutlookLocation(svc: EventService, eventId: string): Promise<void> {
  if (!eventId || svc._outlookLocationBackfilled.has(eventId)) return;
  svc._outlookLocationBackfilled.add(eventId);
  try {
    const numId = Number(eventId);
    if (isNaN(numId)) return;
    const ev = await svc.getEvent(numId);
    if (!ev) return;
    if (ev.OutlookLocation && ev.OutlookLocation.trim() !== '') return; // schon gesetzt
    const loc = buildOutlookLocation(ev.Location, ev.LocationAddress);
    if (loc) {
      await svc.updateEvent(numId, { 'OutlookLocation': loc });
    }
  } catch { /* best effort — Backfill darf den Queue-Eintrag nie blockieren */ }
}

/**
 * v28.37: Wer hat für dieses Event schon eine Einladungsmail bekommen?
 *
 * Liest die DEX_Emails-Zeilen vom Typ `Einladung` zum Event und sammelt
 * Empfaenger aus `Recipient` und `Bcc` (Massenversand läuft in 450er-Chunks
 * über Bcc, im To steht dann nur der ausloesende Organizer). Adressen
 * lowercase, dedupliziert.
 *
 * WICHTIG für den Aufrufer: Alte DEX_Emails-Zeilen werden nach rund einem
 * Monat archiviert. Für länger zurückliegende Versaende ist die Liste
 * daher unvollstaendig — das Ergebnis taugt zum Nachfassen innerhalb einer
 * laufenden Einladungsrunde, nicht als lueckenlose Historie.
 */
export async function getInvitedRecipients(svc: EventService, eventId: string | number): Promise<string[]> {
  const id = String(eventId || '').trim();
  if (!id) return [];
  const out = new Set<string>();
  let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Emails')/items`
    + `?$select=Recipient,Bcc&$filter=EmailType eq 'Einladung' and EventId eq '${id.replace(/'/g, "''")}'&$top=500`;
  let guard = 0;
  while (url && guard < 20) {
    guard++;
    let resp: SPHttpClientResponse;
    try { resp = await svc._sp.get(url, SPHttpClient.configurations.v1); }
    catch { break; }
    if (!resp.ok) break;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let data: any;
    try { data = await resp.json(); } catch { break; }
    const items = data.value || data.d?.results || [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    for (const it of items as any[]) {
      const raw = `${it.Recipient || ''};${it.Bcc || ''}`;
      for (const part of raw.split(/[;,]/)) {
        const e = (part || '').trim().toLowerCase();
        if (e.indexOf('@') > 0) out.add(e);
      }
    }
    url = data['odata.nextLink'] || data['@odata.nextLink'] || (data.d && data.d.__next) || null;
  }
  return Array.from(out);
}

/**
 * Outlook-Termin-Einladung in die Queue eintragen.
 * Flow holt Event-Details (Datum, Ort, CalendarLink) aus DEX_Events via EventId.
 */
export async function queueOutlookEvent(
  svc: EventService,
  attendee: string,
  eventId: string,
  eventTitle: string,
  actionType: 'Einladen' | 'Ausladen' | 'UpdateEvent'
): Promise<boolean> {
  try {
    // v18.34: OutlookLocation für Bestands-Events nachziehen (einmal pro Event/Session).
    await backfillOutlookLocation(svc, eventId);
    const response = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`,
      {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_OutlookListItem' },
        'Title': `${actionType}: ${eventTitle}`,
        'Attendee': attendee,
        'EventId': eventId,
        'ActionType': actionType,
        'Status': 'Pending',
      }
    );
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * DeleteEvent in die DEX_Outlook-Queue eintragen. Wird vom deleteEvent-Flow
 * aufgerufen, BEVOR das DEX_Events-Item gelöscht wird. Der DEX_Outlook_Einladungen-
 * Flow findet den Outlook-Termin über CalendarLink (iCalUId) und löscht ihn.
 * Attendee bleibt leer - DeleteEvent wirkt event-weit.
 */
export async function queueOutlookDeleteEvent(
  svc: EventService,
  eventId: string,
  eventTitle: string,
  calendarLink: string
): Promise<boolean> {
  try {
    const response = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`,
      {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_OutlookListItem' },
        'Title': `DeleteEvent: ${eventTitle}`,
        'Attendee': '',
        'EventId': eventId,
        'ActionType': 'DeleteEvent',
        'Status': 'Pending',
        'CalendarLink': calendarLink,
      }
    );
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * Liefert alle Attendees, die den Outlook-Kalendertermin abgelehnt haben.
 *
 * Liest den Termin im Postfach der Shared Mailbox `no_reply.events@deloitte.de`
 * via Microsoft Graph. Der Admin-User braucht dafür delegate/shared access
 * auf das Postfach (Mailbox-Permission) + die SPFx-App muss `Calendars.Read.Shared`
 * im Admin Center genehmigt bekommen.
 *
 * Holt `OutlookEventId` + `CalendarLink` via `GET` auf DEX_Events/{id}. Primärer
 * Lookup des Outlook-Events über `OutlookEventId`. Wenn leer (alte Events):
 * Fallback über `iCalUId` per `$filter`.
 *
 * Rückgabe-Status:
 * - `ok: true`, `attendees: [...]` - Termin gefunden, Declines extrahiert
 * - `ok: false`, `reason: 'no-pointer'` - DEX_Events hat weder OutlookEventId noch CalendarLink
 * - `ok: false`, `reason: 'not-found'` - Outlook-Termin existiert nicht (mehr)
 * - `ok: false`, `reason: 'forbidden'` - Admin hat keine Mailbox-Permission oder Tenant-Admin hat Calendars.Read.Shared nicht genehmigt
 * - `ok: false`, `reason: 'error'` - unerwarteter Fehler
 */
export async function getDeclinedAttendees(
  svc: EventService,
  eventId: number | string
): Promise<DeclineCheckResult> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ctx = svc.context as any;
  if (!ctx.msGraphClientFactory) return { ok: false, attendees: [], reason: 'error', message: 'Graph-Client nicht verfügbar.' };

  // 1. OutlookEventId + CalendarLink aus DEX_Events holen. Nutzt den bewährten
  // `getEvent()`-Path (gleiche Abfrage wie der Rest der App). Direktes
  // $select=OutlookEventId,CalendarLink hatte in v5.18 zu leeren Strings
  // geführt obwohl die Spalten in SharePoint gefüllt waren.
  let outlookEventId = '';
  let calendarLink = '';
  let loadedEvent = false;
  try {
    const numericId = Number(eventId);
    const spEvent = await svc.getEvent(numericId);
    if (spEvent) {
      loadedEvent = true;
      outlookEventId = String(spEvent.OutlookEventId || '');
      calendarLink = String(spEvent.CalendarLink || '');
      console.warn('[DEX] getDeclinedAttendees: Event geladen', {
        id: numericId,
        outlookEventIdLen: outlookEventId.length,
        calendarLinkLen: calendarLink.length,
      });
    } else {
      console.warn('[DEX] getDeclinedAttendees: getEvent() lieferte null', { eventId });
    }
  } catch (err) {
    console.warn('[DEX] getDeclinedAttendees: getEvent() warf', err);
  }
  if (!outlookEventId && !calendarLink) {
    return {
      ok: false,
      attendees: [],
      reason: 'no-pointer',
      message: loadedEvent
        ? `Event-Item (Id=${eventId}) enthält weder OutlookEventId noch CalendarLink.`
        : `Event-Item (Id=${eventId}) konnte nicht aus DEX_Events geladen werden (403/404?). Details siehe Browser-Console.`,
    };
  }

  // 2. Outlook-Termin via Graph laden
  const mailbox = 'no_reply.events@deloitte.de';
  try {
    const client = await ctx.msGraphClientFactory.getClient('3');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let ev: any = null;
    if (outlookEventId) {
      ev = await client.api(`/users/${mailbox}/events/${outlookEventId}`)
        .select('id,subject,attendees')
        .get();
    } else {
      const escaped = calendarLink.replace(/'/g, "''");
      const resp = await client.api(`/users/${mailbox}/events`)
        .filter(`iCalUId eq '${escaped}'`)
        .select('id,subject,attendees')
        .top(1)
        .get();
      ev = (resp?.value || [])[0] || null;
    }
    if (!ev) return { ok: false, attendees: [], reason: 'not-found' };
    if (!Array.isArray(ev.attendees)) return { ok: true, attendees: [] };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const declined = ev.attendees
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .filter((a: any) => a?.status?.response === 'declined')
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .map((a: any) => ({
        email: String(a?.emailAddress?.address || '').toLowerCase(),
        name: String(a?.emailAddress?.name || ''),
      }))
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .filter((a: any) => a.email);
    return { ok: true, attendees: declined };
  } catch (err) {
    console.warn('[DEX] getDeclinedAttendees failed:', err);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const status = (err as any)?.statusCode || (err as any)?.status;
    if (status === 403 || status === 401) return { ok: false, attendees: [], reason: 'forbidden' };
    if (status === 404) return { ok: false, attendees: [], reason: 'not-found' };
    return { ok: false, attendees: [], reason: 'error', message: err instanceof Error ? err.message : String(err) };
  }
}

/**
 * v22.7: Prüft per Microsoft Graph, ob die übergebenen (Deloitte-)E-Mail-
 * Adressen noch zu einem aktiven Konto gehören. Liefert die Liste der
 * Adressen, zu denen KEIN aktives Konto gefunden wurde — die Person hat
 * womöglich das Unternehmen verlassen oder das Konto ist deaktiviert.
 *
 * - Nur @deloitte-Adressen werden geprüft; externe/Nicht-Deloitte-Adressen
 *   werden übersprungen (nicht zuverlässig prüfbar) und nie gemeldet.
 * - Batches von je 8 Adressen pro Graph-Request (mail/UPN-OR-Filter).
 * - Best-effort: nur Adressen aus ERFOLGREICH abgefragten Batches können
 *   als inaktiv gemeldet werden — fehlgeschlagene Batches erzeugen keinen
 *   Fehlalarm. `ok=false`, wenn gar nichts geprüft werden konnte.
 */
export async function checkAccountsActive(
  svc: EventService,
  emails: string[]
): Promise<{ ok: boolean; inactive: string[] }> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ctx = svc.context as any;
  if (!ctx.msGraphClientFactory) return { ok: false, inactive: [] };
  const candidates = Array.from(new Set(
    emails
      .map(e => (e || '').trim().toLowerCase())
      .filter(e => /@(.*\.)?deloitte\.(de|com)$/i.test(e))
  ));
  if (candidates.length === 0) return { ok: true, inactive: [] };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let client: any;
  try {
    client = await ctx.msGraphClientFactory.getClient('3');
  } catch {
    return { ok: false, inactive: [] };
  }
  const activeSet = new Set<string>();
  const checkedSet = new Set<string>();
  const esc = (s: string): string => s.replace(/'/g, "''");
  // v24.34 HOTFIX: Graph erlaubt max. 15 OR-Klauseln im $filter. Pro E-Mail
  // erzeugen wir 2 Klauseln (mail + userPrincipalName) → Batch 8 = 16 Klauseln
  // → JEDER Batch scheiterte mit HTTP 400. Batch 7 = 14 Klauseln.
  const BATCH = 7;
  for (let i = 0; i < candidates.length; i += BATCH) {
    const batch = candidates.slice(i, i + BATCH);
    const clauses = batch
      .map(e => `mail eq '${esc(e)}' or userPrincipalName eq '${esc(e)}'`)
      .join(' or ');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const resp: any = await client.api('/users')
        .filter(`(${clauses})`)
        .select('mail,userPrincipalName,accountEnabled')
        .top(999)
        .get();
      // Batch gilt als geprüft (egal ob jemand gefunden wurde).
      for (const e of batch) checkedSet.add(e);
      const found = (resp?.value || []) as Array<{ mail?: string; userPrincipalName?: string; accountEnabled?: boolean }>;
      for (const u of found) {
        if (u.accountEnabled === false) continue; // deaktiviert → zählt als inaktiv
        if (u.mail) activeSet.add(u.mail.toLowerCase());
        if (u.userPrincipalName) activeSet.add(u.userPrincipalName.toLowerCase());
      }
    } catch (err) {
      console.warn('[DEX] checkAccountsActive batch failed:', err);
    }
  }
  if (checkedSet.size === 0) return { ok: false, inactive: [] };
  // v26.42: ZWEITER Durchgang für nicht gefundene Adressen — KONTO-UMBENENNUNG
  // erkennen (z.B. Heirat: UPN + primäre Mail wechseln auf den neuen Nachnamen,
  // die ALTE Adresse bleibt als smtp:-Alias am selben, weiterhin AKTIVEN Konto).
  // Der mail/UPN-Filter oben findet solche Konten nicht → früher Fehlalarm
  // „hat Deloitte verlassen". proxyAddresses enthält die alte Adresse als Alias.
  // Best-effort: schlägt die Abfrage fehl (z.B. Berechtigung), bleibt das
  // bisherige Verhalten — der Durchgang kann Personen nur RETTEN, nie zusätzlich
  // belasten.
  const missing = candidates.filter(e => checkedSet.has(e) && !activeSet.has(e));
  // 2 Klauseln je Adresse (smtp:/SMTP: — der Präfix-Vergleich ist case-sensitiv,
  // Sekundär-Aliasse tragen 'smtp:', der Primär-Eintrag 'SMTP:') → 7×2 = 14 ≤ 15.
  for (let i = 0; i < missing.length; i += BATCH) {
    const batch = missing.slice(i, i + BATCH);
    const clauses = batch
      .map(e => `proxyAddresses/any(p: p eq 'smtp:${esc(e)}') or proxyAddresses/any(p: p eq 'SMTP:${esc(e)}')`)
      .join(' or ');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const resp: any = await client.api('/users')
        .header('ConsistencyLevel', 'eventual')
        .query({ '$count': 'true' })
        .filter(`(${clauses})`)
        .select('mail,userPrincipalName,accountEnabled,proxyAddresses')
        .top(999)
        .get();
      const found = (resp?.value || []) as Array<{ mail?: string; userPrincipalName?: string; accountEnabled?: boolean; proxyAddresses?: string[] }>;
      for (const u of found) {
        if (u.accountEnabled === false) continue; // wirklich deaktiviert
        // Alle Aliasse des Kontos als aktiv markieren — darunter die alte Adresse.
        for (const p of (u.proxyAddresses || [])) {
          const addr = String(p || '').replace(/^smtps?:/i, '').trim().toLowerCase();
          if (addr) activeSet.add(addr);
        }
        if (u.mail) activeSet.add(u.mail.toLowerCase());
        if (u.userPrincipalName) activeSet.add(u.userPrincipalName.toLowerCase());
        console.warn('[DEX] checkAccountsActive: Konto umbenannt (Alias-Treffer), NICHT inaktiv:', u.mail || u.userPrincipalName);
      }
    } catch (err) {
      // 403/400 (fehlende Graph-Berechtigung für proxyAddresses o.ä.) →
      // keine Rettung möglich, Verhalten wie vor v26.42.
      console.warn('[DEX] checkAccountsActive proxy-alias pass failed:', err);
    }
  }
  const inactive = candidates.filter(e => checkedSet.has(e) && !activeSet.has(e));
  return { ok: true, inactive };
}

/**
 * v24.33: Unternehmenszugehörigkeit („Company name" / Graph `companyName`)
 * des eingeloggten Users via Microsoft Graph. Die SP-UserProfile-Property
 * „Company" ist im Tenant nicht zuverlässig gefüllt — Graph `/me` schon.
 */
export async function getMyCompanyViaGraph(svc: EventService): Promise<string> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ctx = svc.context as any;
  if (!ctx.msGraphClientFactory) return '';
  try {
    const client = await ctx.msGraphClientFactory.getClient('3');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const resp: any = await client.api('/me').select('companyName').get();
    return (resp?.companyName || '').trim();
  } catch {
    return '';
  }
}

/**
 * v24.33: Unternehmenszugehörigkeit für mehrere E-Mails via Graph (Batch à 8,
 * gleiches Muster wie checkAccountsActive). Liefert eine Map
 * lowercased-E-Mail → companyName. Für den Backfill bestehender Teilnehmer.
 */
export async function getCompaniesByEmails(svc: EventService, emails: string[]): Promise<Record<string, string>> {
  const out: Record<string, string> = {};
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ctx = svc.context as any;
  if (!ctx.msGraphClientFactory) return out;
  const candidates = Array.from(new Set(
    emails.map(e => (e || '').trim().toLowerCase()).filter(e => /@(.*\.)?deloitte\.(de|com)$/i.test(e))
  ));
  if (candidates.length === 0) return out;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let client: any;
  try { client = await ctx.msGraphClientFactory.getClient('3'); } catch { return out; }
  const esc = (s: string): string => s.replace(/'/g, "''");
  // v24.34 HOTFIX: Graph erlaubt max. 15 OR-Klauseln im $filter. Pro E-Mail
  // erzeugen wir 2 Klauseln (mail + userPrincipalName) → Batch 8 = 16 Klauseln
  // → JEDER Batch scheiterte mit HTTP 400. Batch 7 = 14 Klauseln.
  const BATCH = 7;
  for (let i = 0; i < candidates.length; i += BATCH) {
    const batch = candidates.slice(i, i + BATCH);
    const clauses = batch.map(e => `mail eq '${esc(e)}' or userPrincipalName eq '${esc(e)}'`).join(' or ');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const resp: any = await client.api('/users').filter(`(${clauses})`).select('mail,userPrincipalName,companyName').top(999).get();
      const found = (resp?.value || []) as Array<{ mail?: string; userPrincipalName?: string; companyName?: string }>;
      for (const u of found) {
        const comp = (u.companyName || '').trim();
        if (!comp) continue;
        if (u.mail) out[u.mail.toLowerCase()] = comp;
        if (u.userPrincipalName) out[u.userPrincipalName.toLowerCase()] = comp;
      }
    } catch (err) {
      console.warn('[DEX] getCompaniesByEmails batch failed:', err);
    }
  }
  return out;
}

/**
 * v24.33: Trägt die Unternehmenszugehörigkeit für bestehende Teilnehmer einer
 * Liste nach (Backfill) — lädt alle Zeilen, holt `companyName` via Graph und
 * setzt `Company` per MERGE, aber nur dort, wo es noch leer ist. Best-effort:
 * fehlt die Spalte/das Recht, wird die Zeile übersprungen.
 */
export async function backfillCompanyForList(svc: EventService, subsiteUrl: string): Promise<{ updated: number; checked: number }> {
  let regs: SPRegistration[] = [];
  try { regs = await svc.getAllRegistrations(subsiteUrl); } catch { return { updated: 0, checked: 0 }; }
  const emails = Array.from(new Set(regs.map(r => (r.ParticipantEmail || '').toLowerCase()).filter(Boolean)));
  if (emails.length === 0) return { updated: 0, checked: regs.length };
  const compMap = await svc.getCompaniesByEmails(emails);
  let updated = 0;
  for (const r of regs) {
    const comp = compMap[(r.ParticipantEmail || '').toLowerCase()];
    if (!comp) continue;
    if ((r.Company || '').trim()) continue; // schon gesetzt
    if (!r.Id) continue;
    try {
      await svc._merge(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${r.Id})`, { Company: comp });
      updated++;
    } catch { /* Spalte evtl. nicht da / keine Rechte — überspringen */ }
  }
  return { updated, checked: regs.length };
}

