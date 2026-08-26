/**
 * v30.11 — Modularisierung Stufe 2: Thema „DEX_Archive" (v21/v23.40).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 *
 * Drei Teile, eine Ablage: (1) die DEX_Archive-Liste selbst (ensure +
 * Permissions), (2) der Archiv-Lauf, der Zeilen abgelaufener Events aus den
 * Queue-/Log-Listen (ARCHIVE_SOURCES) atomar Insert→Delete ins Archiv
 * verschiebt, (3) das Löschkonzept, das Archiv-Zeilen nach Stichdatum
 * endgültig entfernt. `_delete` bleibt als allgemeiner HTTP-Helfer an der
 * Klasse (wird auch außerhalb des Archivs gebraucht).
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

// ==================== v21: Archivierung ====================
// Globale Queue-/Log-Listen, deren Zeilen abgelaufener Events ins
// DEX_Archive wandern (verschieben = nach Insert aus der Quelle löschen).
// DEX_Emails/Outlook/IDReorder/ChangeLog matchen über EventId, DEX_AccessFix
// über SubsiteUrl (hat keine EventId).
// v22.2: Pro Quelle eine SCHLANKE Feldauswahl für den Archiv-Lauf.
// WICHTIG: DEX_Emails OHNE `Body` — 1000+ komplette HTML-Mail-Bodies in
// den Browser zu laden hängte den Archiv-Lauf praktisch auf (zig MB).
// Der Mail-Body wird bewusst NICHT mitarchiviert (Metadaten reichen als
// Nachweis; die Mail ist längst versendet).
// hasStatus: Liste besitzt eine Status-Spalte → Zeilen mit 'Pending'
// (Flow hat sie noch nicht verarbeitet) werden NICHT archiviert, damit
// keine unversendete Mail / kein offener Auftrag aus der Queue verschwindet.
// v23.47: select MUSS den GESAMTEN Inhalt der Zeile abdecken — das Archiv
// ist die End-Ablage, der Originaldatensatz wird nach dem Insert gelöscht.
// Frühere Selects ließen Inhalte weg: DEX_Emails OHNE `Body` (kompletter
// Mailtext!) und DEX_Outlook OHNE `CalendarLink` — die fehlten damit im
// Payload und gingen beim Löschen verloren. Jetzt: vollständige Feldlisten.
const ARCHIVE_SOURCES: Array<{ list: string; matchBy: 'eventId' | 'subsiteUrl'; select: string; hasStatus: boolean }> = [
  { list: 'DEX_Emails', matchBy: 'eventId', select: 'Id,Title,Recipient,RecipientName,Body,EmailType,EventTitle,EventId,Status,Cc,Bcc,Importance,Created', hasStatus: true },
  { list: 'DEX_Outlook', matchBy: 'eventId', select: 'Id,Title,Attendee,EventId,ActionType,Status,CalendarLink,Created', hasStatus: true },
  { list: 'DEX_IDReorder', matchBy: 'eventId', select: 'Id,Title,EventId,EventNumber,SubsiteUrl,Status,CancelledName,CancelledEmail,Created', hasStatus: true },
  { list: 'DEX_ChangeLog', matchBy: 'eventId', select: 'Id,Title,Action,TargetType,TargetId,TargetName,EventId,EventTitle,ActorName,ActorEmail,Details,Created', hasStatus: false },
  { list: 'DEX_AccessFix', matchBy: 'subsiteUrl', select: 'Id,Title,SubsiteUrl,ItemId,ParticipantEmail,Status,Created', hasStatus: true },
];

/**
 * v21: DEX_Archive anlegen (Site-Collection-Root) — generisches Schema, das
 * Zeilen aus mehreren Quell-Listen aufnimmt: SourceList, EventId, EventTitle,
 * OriginalId, ArchivedAt + Payload (JSON der Originalzeile). NUR Admins
 * (Owners) bekommen Zugriff (setArchiveListPermissions, kein Visitors-Grant).
 */
export async function ensureArchiveList(svc: EventService): Promise<void> {
  const listName = 'DEX_Archive';
  const exists = await svc.listExists(listName);
  if (exists) return;
  const createResp = await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Archiv abgelaufener Event-Zeilen aus den Queue-/Log-Listen (v21). Nur für Admins lesbar.',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });
  if (!createResp.ok) {
    console.warn('[DEX] DEX_Archive konnte nicht angelegt werden — vermutlich fehlen Owner-Rechte.');
    return;
  }
  const fields: Array<{ title: string; type: number; metaType?: string }> = [
    { title: 'SourceList', type: 2 },
    { title: 'EventId', type: 2 },
    { title: 'EventTitle', type: 2 },
    { title: 'OriginalId', type: 9 },
    { title: 'ArchivedAt', type: 4 },
    { title: 'Payload', type: 3, metaType: 'SP.FieldMultiLineText' },
  ];
  for (const f of fields) {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title, 'FieldTypeKind': f.type, 'Required': false,
      };
      if (f.metaType === 'SP.FieldMultiLineText') { payload['RichText'] = false; payload['NumberOfLines'] = 6; }
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    } catch { /* einzelne Feld-Fehler ignorieren */ }
  }
  try { await svc.configureDefaultView(listName, ['SourceList', 'EventTitle', 'EventId', 'OriginalId', 'ArchivedAt']); } catch { /* */ }
  try { await setArchiveListPermissions(svc, listName); } catch { /* */ }
}

/** Admin-only: Vererbung brechen, NUR Owners (Site-Admins) Full Control —
 *  bewusst KEIN Visitors-/Members-Grant, damit das Archiv nicht von allen
 *  lesbar ist. */
async function setArchiveListPermissions(svc: EventService, listName: string): Promise<void> {
  try {
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`, {});
    const ownersResp = await svc._sp.get(`${svc.siteUrl}/_api/web/associatedownergroup?$select=Id`, SPHttpClient.configurations.v1);
    if (ownersResp.ok) {
      const d = await ownersResp.json();
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741829)`, {});
    }
  } catch (e) { console.warn('[DEX] setArchiveListPermissions failed:', e); }
}



/** Lädt alle Zeilen einer Liste (paged, nometadata). `select` schränkt die
 *  Felder ein (fürs Zählen leichtgewichtig); ohne select = alle Felder
 *  (für den Payload). */
async function loadAllListRows(svc: EventService, listName: string, select?: string): Promise<Array<Record<string, unknown>>> {
  const rows: Array<Record<string, unknown>> = [];
  let url = `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$top=500${select ? `&$select=${select}` : ''}`;
  let guard = 0;
  while (url && guard < 500) {
    guard++;
    const resp = await svc._sp.get(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' },
    });
    if (!resp.ok) break;
    const data = await resp.json();
    const arr: Array<Record<string, unknown>> = data.value || data.d?.results || [];
    rows.push(...arr);
    url = (data['odata.nextLink'] as string) || (data['@odata.nextLink'] as string) || '';
  }
  return rows;
}

function rowMatchesExpired(
  r: Record<string, unknown>, matchBy: 'eventId' | 'subsiteUrl',
  expiredEventIds: Set<string>, expiredSubsiteUrls: Set<string>,
  // v23.39: ALLE aktuell existierenden Event-IDs / Subsites. Eine Zeile, deren
  // Bezug NICHT mehr darin vorkommt, gehört zu einem gelöschten Event
  // (verwaist) und ist ebenfalls archivreif. Die size>0-Wächter verhindern,
  // dass bei (noch) nicht geladenen Events fälschlich ALLES als verwaist gilt.
  allEventIds: Set<string>, allSubsiteUrls: Set<string>
): boolean {
  // v22.2: 'Pending' = der Flow hat die Zeile noch nicht verarbeitet —
  // niemals archivieren (sonst verschwindet z.B. eine unversendete Mail
  // aus der Queue). Hängengebliebene Pendings bleiben so in der
  // Arbeitsliste sichtbar, wo der Admin sie sehen soll.
  if (String(r['Status'] || '') === 'Pending') return false;
  // v26.32: Generelle 1-Monats-Karenz — KEINE Zeile wird archiviert, solange
  // sie jünger als ~1 Monat ist, egal ob sie event-los ist (EventId leer/'0'),
  // an ein bereits ABGELAUFENES Event hängt oder zu einem gelöschten (verwaisten)
  // Event gehört. So bleiben frische Mails (Ticket-/Anfrage-Bestätigungen,
  // Organizer-Anträge, Wochenbericht — auch solche zu einer Frage über ein
  // schon vergangenes Event) mindestens einen Monat in der Queue sichtbar.
  // Vorher (v26.26) griff die Karenz NUR für event-lose/verwaiste Zeilen;
  // Mails abgelaufener Events verschwanden sofort — genau das war das Problem.
  const ORPHAN_GRACE_MS = 30 * 24 * 60 * 60 * 1000; // ~1 Monat
  const c = String(r['Created'] || '');
  const createdTs = c ? new Date(c).getTime() : NaN;
  // Ohne/mit ungültigem Erstellungsdatum konservativ NICHT archivieren.
  if (isNaN(createdTs)) return false;
  if ((Date.now() - createdTs) < ORPHAN_GRACE_MS) return false;
  // Ab hier ist die Zeile ≥ 1 Monat alt.
  if (matchBy === 'eventId') {
    const id = String(r['EventId'] || '').trim();
    // Keinem Event zugeordnet (leer/'0') → archivreif (alt genug).
    if (!id || id === '0') return true;
    // Event abgelaufen → archivreif.
    if (expiredEventIds.has(id)) return true;
    // Event existiert nicht mehr (gelöscht/verwaist) → archivreif.
    if (allEventIds.size > 0 && !allEventIds.has(id)) return true;
    // Event noch aktiv → nicht archivieren.
    return false;
  }
  const su = String(r['SubsiteUrl'] || '').toLowerCase().trim();
  if (!su) return true;
  if (expiredSubsiteUrls.has(su)) return true;
  if (allSubsiteUrls.size > 0 && !allSubsiteUrls.has(su)) return true;
  return false;
}

/** v21: Zählt die archivreifen Zeilen pro Quell-Liste (leichtgewichtig:
 *  nur Id+EventId bzw. Id+SubsiteUrl). */
export async function countArchivableRows(svc: EventService, 
  expiredEventIds: Set<string>, expiredSubsiteUrls: Set<string>,
  allEventIds: Set<string> = new Set(), allSubsiteUrls: Set<string> = new Set()
): Promise<{ total: number; perList: Record<string, number> }> {
  const perList: Record<string, number> = {};
  let total = 0;
  for (const src of ARCHIVE_SOURCES) {
    let c = 0;
    try {
      // v22.2: Status mitladen (wo vorhanden), damit der Pending-Ausschluss
      // schon beim Zählen greift und die Box-Zahl zum Lauf passt.
      // v26.26: Created mitladen — die Orphan-Regel (1-Monats-Karenz) braucht
      // das Erstellungsdatum, sonst zählt die Box anders als der echte Lauf.
      const base = src.matchBy === 'eventId' ? 'Id,EventId,Created' : 'Id,SubsiteUrl,Created';
      const select = src.hasStatus ? `${base},Status` : base;
      const rows = await loadAllListRows(svc, src.list, select);
      for (const r of rows) {
        if (rowMatchesExpired(r, src.matchBy, expiredEventIds, expiredSubsiteUrls, allEventIds, allSubsiteUrls)) c++;
      }
    } catch { /* Liste evtl. nicht vorhanden */ }
    perList[src.list] = c;
    total += c;
  }
  return { total, perList };
}

/** v21: Verschiebt alle archivreifen Zeilen ins DEX_Archive (Insert →
 *  Delete aus der Quelle). Sequentiell (SP-Throttling). onProgress meldet
 *  Listen- + Zeilen-Fortschritt fürs Modal. */
export async function archiveExpiredRows(svc: EventService, 
  expiredEventIds: Set<string>, expiredSubsiteUrls: Set<string>,
  eventTitleById: Record<string, string>,
  onProgress?: (listIdx: number, listTotal: number, listName: string, done: number, total: number) => void,
  // v22.2: Abbruch-Check (UI-Button). Sauber: bereits verschobene Zeilen
  // bleiben im Archiv (jede Zeile ist atomar Insert→Delete), der Rest
  // bleibt in der Quelle und kommt beim nächsten Lauf dran.
  shouldCancel?: () => boolean,
  // v23.39: alle aktuellen Event-IDs / Subsites (für die Verwaist-Erkennung).
  allEventIds: Set<string> = new Set(), allSubsiteUrls: Set<string> = new Set()
): Promise<{ archived: number; failed: number; cancelled: boolean; perList: Record<string, number> }> {
  const result = { archived: 0, failed: 0, cancelled: false, perList: {} as Record<string, number> };
  const sources = ARCHIVE_SOURCES;
  // v23.47: Payload soll den Inhalt vollständig festhalten (vorher bei 4000
  // Zeichen gekappt — ein kompletter HTML-Mailtext ist länger und wurde so
  // abgeschnitten). Großzügiger Sicherheits-Cap (60000), der praktisch jeden
  // Mailtext komplett aufnimmt, aber pathologische Riesenwerte begrenzt.
  const MAX_FIELD = 60000;
  const buildPayload = (r: Record<string, unknown>): string => {
    const out: Record<string, unknown> = {};
    for (const k of Object.keys(r)) {
      const v = r[k];
      out[k] = (typeof v === 'string' && v.length > MAX_FIELD) ? `${v.slice(0, MAX_FIELD)}… [gekürzt]` : v;
    }
    return JSON.stringify(out);
  };
  for (let si = 0; si < sources.length; si++) {
    const src = sources[si];
    if (shouldCancel && shouldCancel()) { result.cancelled = true; break; }
    let listArchived = 0;
    try {
      // v22.2: Fortschritt SOFORT melden (vorher kam der erste Callback erst
      // nach dem Komplett-Laden — das Modal wirkte eingefroren).
      if (onProgress) onProgress(si, sources.length, src.list, 0, 0);
      const rows = await loadAllListRows(svc, src.list, src.select);
      const matching = rows.filter(r => rowMatchesExpired(r, src.matchBy, expiredEventIds, expiredSubsiteUrls, allEventIds, allSubsiteUrls));
      if (onProgress) onProgress(si, sources.length, src.list, 0, matching.length);
      for (let i = 0; i < matching.length; i++) {
        if (shouldCancel && shouldCancel()) { result.cancelled = true; break; }
        const r = matching[i];
        const origId = Number(r['Id'] || 0);
        const eid = src.matchBy === 'eventId' ? String(r['EventId'] || '') : '';
        const title = eid ? (eventTitleById[eid] || '') : '';
        try {
          const ins = await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Archive')/items`, {
            '__metadata': { 'type': 'SP.Data.DEX_x005f_ArchiveListItem' },
            'Title': `${src.list}#${origId}`.slice(0, 250),
            'SourceList': src.list,
            'EventId': eid,
            'EventTitle': title,
            'OriginalId': origId,
            'ArchivedAt': new Date().toISOString(),
            'Payload': buildPayload(r),
          });
          if (ins.ok && origId > 0) {
            // Nur löschen, wenn der Archiv-Insert geklappt hat (kein Datenverlust).
            const del = await svc._delete(`${svc.siteUrl}/_api/web/lists/getbytitle('${src.list}')/items(${origId})`);
            if (del.ok) { listArchived++; result.archived++; } else { result.failed++; }
          } else {
            result.failed++;
          }
        } catch { result.failed++; }
        if (onProgress) onProgress(si, sources.length, src.list, i + 1, matching.length);
      }
    } catch { /* Liste nicht vorhanden — überspringen */ }
    result.perList[src.list] = listArchived;
    if (result.cancelled) break;
  }
  return result;
}


// ==================== Archiv-Löschkonzept (v23.40) ====================
// DEX_Archive-Einträge sind die End-Ablage. Damit die Liste nicht unendlich
// wächst, können Admins Einträge löschen, die älter als ein Stichdatum sind
// (v23.48: standardmäßig 1 Monat nach Ablauf). „ArchivedAt" ist der Ablage-Zeitpunkt.

/** Zählt DEX_Archive-Zeilen mit ArchivedAt älter als `olderThanIso`. */
export async function countDeletableArchiveRows(svc: EventService, olderThanIso: string): Promise<number> {
  try {
    const cutoff = new Date(olderThanIso).getTime();
    if (!isFinite(cutoff)) return 0;
    const rows = await loadAllListRows(svc, 'DEX_Archive', 'Id,ArchivedAt');
    let c = 0;
    for (const r of rows) {
      const a = r['ArchivedAt'] ? new Date(String(r['ArchivedAt'])).getTime() : 0;
      if (a > 0 && a < cutoff) c++;
    }
    return c;
  } catch { return 0; }
}

/** Löscht DEX_Archive-Zeilen älter als `olderThanIso` (sequentiell, mit
 *  Fortschritt + Abbruch). Hartes DELETE — bewusst (das Archiv ist die
 *  letzte Stufe; ältere Einträge braucht niemand mehr). */
export async function deleteOldArchiveRows(svc: EventService, 
  olderThanIso: string,
  onProgress?: (done: number, total: number) => void,
  shouldCancel?: () => boolean
): Promise<{ deleted: number; failed: number; cancelled: boolean }> {
  const result = { deleted: 0, failed: 0, cancelled: false };
  try {
    const cutoff = new Date(olderThanIso).getTime();
    if (!isFinite(cutoff)) return result;
    const rows = await loadAllListRows(svc, 'DEX_Archive', 'Id,ArchivedAt');
    const targets = rows.filter(r => {
      const a = r['ArchivedAt'] ? new Date(String(r['ArchivedAt'])).getTime() : 0;
      return a > 0 && a < cutoff;
    });
    if (onProgress) onProgress(0, targets.length);
    for (let i = 0; i < targets.length; i++) {
      if (shouldCancel && shouldCancel()) { result.cancelled = true; break; }
      const id = Number(targets[i]['Id'] || 0);
      if (id > 0) {
        try {
          const del = await svc._delete(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Archive')/items(${id})`);
          if (del.ok) result.deleted++; else result.failed++;
        } catch { result.failed++; }
      }
      if (onProgress) onProgress(i + 1, targets.length);
    }
  } catch { /* Liste evtl. nicht vorhanden */ }
  return result;
}

