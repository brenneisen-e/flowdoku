/**
 * v30.66 — Modularisierung Stufe 2: Thema „Merk- und Protokoll-Listen rund um
 * Mails": DEX_InactiveNotices (wer wurde schon auf ein totes Konto
 * hingewiesen), DEX_PostEventMails (Nachbereitungs-Mail je Event nur einmal),
 * DEX_EventComms (Kommunikations-Log der Rundmails) und DEX_OutlookLocks
 * (Sperre gegen doppelte Outlook-Läufe).
 *
 * Alle vier Listen sind reine Nebenbuchhaltung: Sie halten fest, was schon
 * verschickt wurde, damit ein zweiter Lauf nicht dieselbe Mail erzeugt.
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService, EventCommRow } from '../EventService';

/**
 * v24.51: Liste `DEX_InactiveNotices` — Dedup-Marker für die „Organizer über
 * inaktives Konto informieren"-Mail. ReadSecurity=1 (alle Admins lesen
 * dieselben Marker → klickt ein zweiter Admin, wird NICHT erneut gesendet).
 * Spalten: EventId + ParticipantEmail (Schlüssel) + SentByEmail (informativ).
 */
export async function ensureInactiveNoticesList(svc: EventService): Promise<void> {
  const listName = 'DEX_InactiveNotices';
  const exists = await svc.listExists(listName);
  if (exists) return;
  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Dedup: Organizer wurde über ein inaktives Konto informiert (v24.51). Eine Mail pro Event+Person, egal welcher Admin klickt.',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });
  const fields: Array<{ title: string; type: number }> = [
    { title: 'EventId', type: 2 },
    { title: 'ParticipantEmail', type: 2 },
    { title: 'SentByEmail', type: 2 },
  ];
  for (const f of fields) {
    try {
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false,
      });
    } catch { /* */ }
  }
  try { await svc.configureDefaultView(listName, ['EventId', 'ParticipantEmail', 'SentByEmail', 'Created']); } catch { /* */ }
  try { await svc.setQueueListPermissions(listName); } catch { /* */ }
  // ReadSecurity bleibt 1 (Default) — alle Admins müssen dieselben Marker
  // sehen können (sonst keine Cross-Admin-Dedup).
}

/** v24.51: Bereits benachrichtigte (Event+Email)-Paare für ein Event lesen. */
export async function getSentInactiveNotices(svc: EventService, eventId: string): Promise<Set<string>> {
  const out = new Set<string>();
  try {
    const esc = (eventId || '').replace(/'/g, "''");
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_InactiveNotices')/items?$select=ParticipantEmail&$filter=EventId eq '${esc}'&$top=500`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return out;
    const data = await resp.json();
    for (const it of (data.value || data.d?.results || [])) {
      const e = (it.ParticipantEmail || '').toLowerCase().trim();
      if (e) out.add(e);
    }
  } catch { /* */ }
  return out;
}

/** v24.51: Benachrichtigungs-Marker anlegen (claim). */
export async function recordInactiveNotice(svc: EventService, eventId: string, participantEmail: string): Promise<void> {
  try {
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_InactiveNotices')/items`, {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_InactiveNoticesListItem' },
      'Title': `${eventId} | ${participantEmail}`.slice(0, 250),
      'EventId': eventId,
      'ParticipantEmail': participantEmail,
      'SentByEmail': (svc.context.pageContext.user.email || '').toLowerCase(),
    });
  } catch (err) { console.warn('[DEX] recordInactiveNotice failed:', err); }
}

/**
 * v26.39: `DEX_PostEventMails` — PERSISTENTER Dedup-Marker für die Post-Event-
 * „Danke & Hinweis zur Aufbewahrung"-Mail an die Organizer. Vorher hing die
 * Entdopplung an der TRANSIENTEN `DEX_Emails`-Queue (nach Versand + Archivierung
 * wieder leer → `hasQueuedEmail` fand nichts → Mail wurde beim nächsten
 * App-Öffnen ERNEUT verschickt) + Browser-`localStorage` (pro Gerät/User). Ein
 * dauerhafter Marker pro EventId behebt das geräte- und zeitübergreifend.
 * Rückgabe: true = Liste wurde JETZT NEU angelegt (Erststart → Altbestand seeden).
 */
export async function ensurePostEventMailsList(svc: EventService): Promise<boolean> {
  const listName = 'DEX_PostEventMails';
  try { if (await svc.listExists(listName)) return false; } catch { return false; }
  const cr = await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Dedup (v26.39): „Danke & Hinweis zur Aufbewahrung"-Mail wurde für dieses Event bereits an die Organizer verschickt. Ein Marker pro Event — verhindert Mehrfachversand.',
    'BaseTemplate': 100, 'AllowContentTypes': false,
  });
  if (!cr.ok) return false;
  for (const f of [{ title: 'EventId', type: 2 }, { title: 'SentByEmail', type: 2 }]) {
    try { await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, { '__metadata': { 'type': 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false }); } catch { /* */ }
  }
  try { await svc.configureDefaultView(listName, ['EventId', 'SentByEmail', 'Created']); } catch { /* */ }
  try { await svc.setQueueListPermissions(listName); } catch { /* */ }
  return true;
}

/** v26.39: Alle EventIds, für die die Post-Event-Mail schon raus ist. */
export async function getPostEventMailSentEventIds(svc: EventService): Promise<Set<string>> {
  const out = new Set<string>();
  try {
    if (!(await svc.listExists('DEX_PostEventMails'))) return out;
    let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_PostEventMails')/items?$select=EventId&$top=5000`;
    while (url) {
      const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) break;
      const data = await resp.json();
      for (const it of (data.value || data.d?.results || [])) { const e = String(it.EventId || '').trim(); if (e) out.add(e); }
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    }
  } catch { /* */ }
  return out;
}

/** v26.39: Marker setzen — Post-Event-Mail für dieses Event ist erledigt. */
export async function recordPostEventMail(svc: EventService, eventId: string | number): Promise<void> {
  try {
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_PostEventMails')/items`, {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_PostEventMailsListItem' },
      'Title': `${eventId}`.slice(0, 250),
      'EventId': String(eventId),
      'SentByEmail': (svc.context.pageContext.user.email || '').toLowerCase(),
    });
  } catch (err) { console.warn('[DEX] recordPostEventMail failed:', err); }
}

// =====================================================================
// v26.41: DEX_EventComms — dauerhaftes Kommunikations-Log der EVENT-RUNDMAILS
// (Einladung, Massenmail, Ankündigungen). Organizer sehen die Historie im
// Organizer Center; Teilnehmer die Rundmails unter „Meine Events" (Nachrücker/
// Spätanmelder können nachlesen). KEINE persönlichen Bestätigungsmails.
// Lesbar für alle Site-Nutzer (Rundmails sind nicht vertraulich); geschrieben
// nur bei Organizer-Aktionen (queueEmail-Begleitung).
// =====================================================================
export async function ensureEventCommsList(svc: EventService): Promise<void> {
  const listName = 'DEX_EventComms';
  try { if (await svc.listExists(listName)) return; } catch { return; }
  const cr = await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Kommunikations-Log (v26.41): Event-Rundmails (Einladung/Massenmail/Ankündigung) mit Zeitstempel — Organizer-Historie + Teilnehmer-Ansicht unter „Meine Events". Keine persönlichen Bestätigungsmails.',
    'BaseTemplate': 100, 'AllowContentTypes': false,
  });
  if (!cr.ok) { console.warn('[DEX] DEX_EventComms konnte nicht angelegt werden.'); return; }
  const fields: Array<{ title: string; type: number; note?: boolean }> = [
    { title: 'EventId', type: 2 }, { title: 'EventTitle', type: 2 },
    { title: 'Subject', type: 2 }, { title: 'BodyHtml', type: 3, note: true },
    { title: 'EmailType', type: 2 }, { title: 'SentByEmail', type: 2 }, { title: 'SentByName', type: 2 },
  ];
  for (const f of fields) {
    try {
      const payload: Record<string, unknown> = { '__metadata': { 'type': f.note ? 'SP.FieldMultiLineText' : 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false };
      if (f.note) { payload['RichText'] = false; payload['NumberOfLines'] = 12; }
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    } catch { /* einzelne Feld-Fehler ignorieren */ }
  }
  try { await svc.configureDefaultView(listName, ['EventId', 'EventTitle', 'Subject', 'EmailType', 'SentByName', 'Created']); } catch { /* */ }
  // KEIN Inheritance-Break: alle Site-Nutzer dürfen die Rundmails lesen.
}

/** Eine gesendete Event-Rundmail ins Log schreiben. */
export async function logEventComm(svc: EventService, meta: { eventId: string | number; eventTitle: string; subject: string; bodyHtml: string; emailType: string }): Promise<void> {
  try {
    await svc.ensureEventCommsList();
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EventComms')/items`, {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_EventCommsListItem' },
      'Title': `${meta.eventId}: ${(meta.subject || '').slice(0, 180)}`.slice(0, 250),
      'EventId': String(meta.eventId), 'EventTitle': meta.eventTitle || '',
      'Subject': meta.subject || '', 'BodyHtml': meta.bodyHtml || '',
      'EmailType': meta.emailType || '',
      'SentByEmail': (svc.context.pageContext.user.email || '').toLowerCase(),
      'SentByName': svc.context.pageContext.user.displayName || '',
    });
  } catch (err) { console.warn('[DEX] logEventComm failed:', err); }
}

/** Alle Rundmails eines Events (neueste zuerst). */
export async function getEventComms(svc: EventService, eventId: string | number): Promise<EventCommRow[]> {
  const out: EventCommRow[] = [];
  try {
    if (!(await svc.listExists('DEX_EventComms'))) return out;
    const esc = String(eventId).replace(/'/g, "''");
    const sel = 'Id,EventId,Subject,BodyHtml,EmailType,SentByName,SentByEmail,Created';
    let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EventComms')/items?$select=${sel}&$filter=EventId eq '${esc}'&$orderby=Created desc&$top=500`;
    while (url) {
      const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) break;
      const data = await resp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const it of ((data.value || data.d?.results || []) as any[])) {
        out.push({ id: Number(it.Id), eventId: String(it.EventId || ''), subject: it.Subject || '', bodyHtml: it.BodyHtml || '', emailType: it.EmailType || '', sentByName: it.SentByName || '', sentByEmail: it.SentByEmail || '', created: it.Created || '' });
      }
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    }
  } catch (err) { console.warn('[DEX] getEventComms failed:', err); }
  return out;
}

/** v26.69: Eine Log-Zeile aus dem Kommunikations-Log löschen — z. B. wenn ein
 *  Eintrag versehentlich protokolliert wurde und den „Bereits versendete Infos"-
 *  Hinweis fälschlich auslöst. Gibt true bei Erfolg zurück. */
export async function deleteEventComm(svc: EventService, id: number): Promise<boolean> {
  try {
    if (!(await svc.listExists('DEX_EventComms'))) return false;
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EventComms')/items(${Number(id)})`;
    const resp = await svc._delete(url);
    return resp.ok;
  } catch (err) { console.warn('[DEX] deleteEventComm failed:', err); return false; }
}

/** Gibt es überhaupt Rundmails zu diesem Event? (für den Anmeldemail-Hinweis) */
/**
 * @param excludeTypes v29.11: Rundmail-Arten, die NICHT zählen sollen.
 *
 * Hintergrund: Der Hinweis „Bereits versendete Infos zu diesem Event“ in der
 * Anmeldebestätigung soll die Person auf Kommunikation aufmerksam machen, die
 * sie verpasst haben könnte. Die Einladung ist genau das nicht — über sie ist
 * die Person meist überhaupt erst gekommen, und wer sich ohne Einladung
 * anmeldet, kann sie in der App ohnehin nachlesen. Stand nur eine Einladung
 * im Log, verwies der Hinweis also auf die Mail, die man gerade in der Hand
 * hatte. Lohnend ist er erst, wenn es DARÜBER HINAUS etwas gab.
 *
 * Die Auswertung läuft bewusst im Code und nicht als OData-Filter: `ne` würde
 * Zeilen mit leerem EmailType je nach Auslegung verschlucken. Alt-Zeilen ohne
 * Art zählen hier als „weitere Mail“ — im Zweifel lieber hinweisen als eine
 * echte Ankündigung verschweigen.
 */
export async function hasEventComms(
  svc: EventService,
  eventId: string | number,
  excludeTypes?: string[],
): Promise<boolean> {
  try {
    if (!(await svc.listExists('DEX_EventComms'))) return false;
    const esc = String(eventId).replace(/'/g, "''");
    const resp = await svc._sp.get(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EventComms')/items?$select=Id,EmailType&$filter=EventId eq '${esc}'&$top=200`, SPHttpClient.configurations.v1);
    if (!resp.ok) return false;
    const data = await resp.json();
    const items = (data.value || data.d?.results || []) as Array<{ EmailType?: string }>;
    if (!Array.isArray(items) || items.length === 0) return false;
    if (!excludeTypes || excludeTypes.length === 0) return true;
    const skip = excludeTypes.map(t => (t || '').trim().toLowerCase());
    return items.some(it => skip.indexOf((it.EmailType || '').trim().toLowerCase()) < 0);
  } catch { return false; }
}

export async function ensureOutlookLocksList(svc: EventService): Promise<void> {
  const listName = 'DEX_OutlookLocks';
  const exists = await svc.listExists(listName);
  if (exists) return;

  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Pro-Event-Sperre für den Outlook-Einladungs-Flow (v18.48) — verhindert gleichzeitige Läufe für dasselbe Event.',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });

  // EventId: der Lock-Schlüssel. Muss eindeutig + indiziert sein, damit das
  // Create-als-Lock-Erwerb-Muster funktioniert (zweiter gleichzeitiger
  // Create für dieselbe EventId schlägt fehl -> der Lauf wartet & retryt).
  await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
    '__metadata': { 'type': 'SP.Field' },
    'Title': 'EventId',
    'FieldTypeKind': 2,
    'Required': false,
  });
  // LockedAt: rein informativ (Debugging hängengebliebener Locks).
  await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
    '__metadata': { 'type': 'SP.Field' },
    'Title': 'LockedAt',
    'FieldTypeKind': 4,
    'Required': false,
  });

  // EventId indizieren und Eindeutigkeit erzwingen. Reihenfolge wichtig:
  // erst Indexed, dann EnforceUniqueValues (SP verlangt eine indizierte
  // Spalte für die Eindeutigkeits-Prüfung). Auf einer frischen, leeren
  // Liste ist das unkritisch (keine Duplikate vorhanden).
  const fieldUrl = `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('EventId')`;
  try {
    await svc._merge(fieldUrl, { 'Indexed': true });
    await svc._merge(fieldUrl, { 'EnforceUniqueValues': true });
  } catch (e) {
    console.warn('[DEX] ensureOutlookLocksList: EnforceUniqueValues konnte nicht gesetzt werden:', e);
  }

  await svc.configureDefaultView(listName, ['EventId', 'LockedAt']);

  // Schreibrechte wie bei den anderen Queue-Listen (DEX_Emails etc.) —
  // der Flow-Connection-Account muss Lock-Items anlegen/löschen können.
  try {
    await svc.setQueueListPermissions(listName);
  } catch { /* best-effort */ }
}
