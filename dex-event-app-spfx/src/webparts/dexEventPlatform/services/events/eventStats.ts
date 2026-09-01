/**
 * v30.66 — Modularisierung Stufe 2: Thema „Event-Statistik und Datenlöschung":
 * DEX_EventStats (die Kennzahlen eines Events überleben dort die Löschung der
 * Teilnehmerliste) plus die beiden Löschpfade — `deleteParticipantData`
 * (Subsite samt Anmeldungen weg, Statistik bleibt) und `deleteEventItemOnly`
 * (nur die Zeile in DEX_Events, für den nicht-destruktiven Recreate-Pfad).
 *
 * Reihenfolge in `deleteParticipantData` NICHT umstellen: erst die prüfbare
 * Nebenbuchhaltung (DEX_Participants) sequentiell aufräumen, dann erst
 * unwiderruflich löschen — siehe CLAUDE.md, v29.2.
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
// EventService als WERT-Import: nur für den deferred Zugriff auf die statische
// stripNoteWrapper — der Zyklus ist unkritisch, weil der Zugriff erst zur
// Laufzeit in Funktionskörpern passiert (wie REG_LIST_NAME).
import { EventService } from '../EventService';

// v30.66: waren `private static readonly` an der Klasse; beide Namen werden
// nur von diesem Thema gebraucht und sind deshalb mitgewandert.
const EVENTSTATS_LIST = 'DEX_EventStats';
const EVENTSTATS_ITEM_TYPE = 'SP.Data.DEX_x005f_EventStatsListItem';

/** Legt DEX_EventStats (Statistik-Archiv, KEINE PII) an, falls nicht vorhanden.
 *  Existiert die Liste bereits, werden nur fehlende (neue) Spalten nachgerüstet
 *  — z.B. `Organizer` (v26.33), das es in früheren Versionen noch nicht gab. */
export async function ensureEventStatsList(svc: EventService): Promise<void> {
  const listName = EVENTSTATS_LIST;
  const fields: Array<{ title: string; type: number }> = [
    { title: 'EventNumber', type: 9 }, { title: 'EventTitle', type: 2 },
    { title: 'EventType', type: 2 }, { title: 'EventLocation', type: 2 },
    { title: 'EventStart', type: 4 }, { title: 'EventEnd', type: 4 },
    { title: 'MaxParticipants', type: 9 }, { title: 'RegisteredCount', type: 9 },
    { title: 'QRSentCount', type: 9 }, { title: 'CheckedInCount', type: 9 },
    { title: 'NoShowCount', type: 9 }, { title: 'WaitlistCount', type: 9 },
    { title: 'DeregisteredCount', type: 9 }, { title: 'ArchivedByEmail', type: 2 },
    { title: 'ArchivedDate', type: 4 }, { title: 'Organizer', type: 3 },
  ];
  let exists = false;
  try { exists = await svc.listExists(listName); } catch { exists = false; }
  if (!exists) {
    const createResp = await svc._post(`${svc.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Statistik-Archiv (v26): KPIs eines Events, nachdem die Teilnehmerliste nach 3 Monaten gelöscht wurde. Enthält KEINE personenbezogenen Daten.',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });
    if (!createResp.ok) { console.warn('[DEX] DEX_EventStats konnte nicht angelegt werden — vermutlich fehlen Owner-Rechte.'); return; }
  }
  // Nur fehlende Felder anlegen (idempotent): bestehende Internal-Names holen.
  const have = new Set<string>();
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName&$top=200`,
      SPHttpClient.configurations.v1,
    );
    if (resp.ok) {
      const data = await resp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const f of ((data.value || data.d?.results || []) as any[])) { if (f.InternalName) have.add(String(f.InternalName)); }
    }
  } catch { /* dann versuchen wir einfach alle anzulegen */ }
  for (const f of fields) {
    if (have.has(f.title)) continue;
    try {
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' }, 'Title': f.title, 'FieldTypeKind': f.type, 'Required': false,
      });
    } catch { /* einzelne Feld-Fehler ignorieren */ }
  }
}

/** EventNumbers, für die bereits ein Statistik-Archiv-Eintrag existiert. */
export async function getArchivedStatsEventNumbers(svc: EventService): Promise<Set<number>> {
  const out = new Set<number>();
  try {
    if (!(await svc.listExists(EVENTSTATS_LIST))) return out;
    let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('${EVENTSTATS_LIST}')/items?$select=EventNumber&$top=5000`;
    while (url) {
      const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) break;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const it of (items as any[])) { const n = Number(it.EventNumber); if (!Number.isNaN(n)) out.add(n); }
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    }
  } catch { /* best-effort */ }
  return out;
}

/**
 * Liest alle Zeilen aus dem Statistik-Archiv (DEX_EventStats). Reine KPI-
 * Daten, KEINE PII. Für die Anzeige-Kachel „Statistik-Archiv" im Admin Center.
 */
export async function getEventStats(svc: EventService): Promise<Array<{
  id: number; eventNumber: number; eventTitle: string; eventType: string;
  location: string; startDate: string; endDate: string; maxParticipants: number | null;
  registeredCount: number; qrSentCount: number; checkedInCount: number;
  noShowCount: number; waitlistCount: number; deregisteredCount: number;
  organizer: string; archivedByEmail: string; archivedDate: string;
}>> {
  const out: Array<{
    id: number; eventNumber: number; eventTitle: string; eventType: string;
    location: string; startDate: string; endDate: string; maxParticipants: number | null;
    registeredCount: number; qrSentCount: number; checkedInCount: number;
    noShowCount: number; waitlistCount: number; deregisteredCount: number;
    organizer: string; archivedByEmail: string; archivedDate: string;
  }> = [];
  try {
    if (!(await svc.listExists(EVENTSTATS_LIST))) return out;
    const sel = 'Id,EventNumber,EventTitle,EventType,EventLocation,EventStart,EventEnd,MaxParticipants,RegisteredCount,QRSentCount,CheckedInCount,NoShowCount,WaitlistCount,DeregisteredCount,Organizer,ArchivedByEmail,ArchivedDate';
    let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('${EVENTSTATS_LIST}')/items?$select=${sel}&$orderby=ArchivedDate desc&$top=5000`;
    while (url) {
      const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) break;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const it of (items as any[])) {
        const num = (v: unknown): number => { const n = Number(v); return Number.isNaN(n) ? 0 : n; };
        out.push({
          id: Number(it.Id),
          eventNumber: num(it.EventNumber),
          eventTitle: it.EventTitle || '',
          eventType: it.EventType || '',
          location: it.EventLocation || '',
          startDate: it.EventStart || '',
          endDate: it.EventEnd || '',
          maxParticipants: (it.MaxParticipants === null || it.MaxParticipants === undefined) ? null : num(it.MaxParticipants),
          registeredCount: num(it.RegisteredCount),
          qrSentCount: num(it.QRSentCount),
          checkedInCount: num(it.CheckedInCount),
          noShowCount: num(it.NoShowCount),
          waitlistCount: num(it.WaitlistCount),
          deregisteredCount: num(it.DeregisteredCount),
          organizer: EventService.stripNoteWrapper(it.Organizer) || '',
          archivedByEmail: it.ArchivedByEmail || '',
          archivedDate: it.ArchivedDate || '',
        });
      }
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    }
  } catch (err) { console.warn('[DEX] getEventStats failed:', err); }
  return out;
}

/**
 * Berechnet die KPIs eines Events aus der Teilnehmerliste und schreibt EINE
 * Zeile ins DEX_EventStats (keine PII). Muss VOR dem Löschen der Subsite
 * laufen. Idempotent: existiert schon eine Zeile für die EventNumber, wird
 * nichts geschrieben (true).
 */
export async function archiveEventStats(svc: EventService, meta: {
  eventNumber: number; eventTitle: string; eventType?: string; location?: string;
  startDate?: string; endDate?: string; maxParticipants?: number; subsiteUrl?: string;
  organizer?: string;
}): Promise<boolean> {
  try {
    await svc.ensureEventStatsList();
    const existing = await svc.getArchivedStatsEventNumbers();
    if (existing.has(meta.eventNumber)) return true;
    // v30.66: `getAllRegistrations` WIRFT NICHT — bei einem HTTP-Fehler bricht die
    // Leseschleife ab und liefert das bis dahin Gelesene, bei 403/500 also `[]`.
    // Ohne geprueften Status waeren alle KPIs 0, und weil der Aufrufer genau auf
    // dieses `true` hin die Subsite UNWIDERRUFLICH loescht, waeren die echten
    // Zahlen danach nicht mehr rekonstruierbar. Ein Lesefehler heisst deshalb:
    // nicht archivieren, nicht loeschen, beim naechsten Lauf erneut versuchen.
    let readError = 0;
    const regs = meta.subsiteUrl
      ? await svc.getAllRegistrations(meta.subsiteUrl, (status: number) => { readError = status; })
      : [];
    if (readError) {
      console.warn('[DEX] archiveEventStats: Teilnehmerliste nicht lesbar (HTTP ' + readError + ') — Event ' + meta.eventNumber + ' wird NICHT archiviert.');
      return false;
    }
    const countBy = (pred: (s: string) => boolean): number => regs.filter(r => pred(r.Status || '')).length;
    const me = (svc.context.pageContext.user.email || '').toLowerCase();
    const resp = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${EVENTSTATS_LIST}')/items`,
      {
        '__metadata': { 'type': EVENTSTATS_ITEM_TYPE },
        'Title': `${meta.eventNumber}: ${(meta.eventTitle || '').slice(0, 200)}`,
        'EventNumber': meta.eventNumber, 'EventTitle': meta.eventTitle || '',
        'EventType': meta.eventType || '', 'EventLocation': meta.location || '',
        'EventStart': meta.startDate || null, 'EventEnd': meta.endDate || null,
        'MaxParticipants': typeof meta.maxParticipants === 'number' ? meta.maxParticipants : null,
        'RegisteredCount': countBy(s => s === 'Angemeldet' || s === 'QR versendet' || s === 'Eingecheckt'),
        'QRSentCount': countBy(s => s === 'QR versendet' || s === 'Eingecheckt'),
        'CheckedInCount': countBy(s => s === 'Eingecheckt'),
        'NoShowCount': countBy(s => s === 'No-Show'),
        'WaitlistCount': countBy(s => s === 'Warteliste'),
        'DeregisteredCount': countBy(s => s === 'Abgemeldet'),
        'ArchivedByEmail': me, 'ArchivedDate': new Date().toISOString(),
        'Organizer': meta.organizer || '',
      }
    );
    return resp.ok;
  } catch (err) { console.warn('[DEX] archiveEventStats failed:', err); return false; }
}

/** Rollback: entfernt die (soeben geschriebene) Statistik-Zeile(n) einer
 *  EventNumber — genutzt, wenn die anschließende Löschung fehlschlägt, damit
 *  das Event beim nächsten Lauf erneut verarbeitet wird. */
export async function deleteEventStatsRow(svc: EventService, eventNumber: number): Promise<boolean> {
  try {
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('${EVENTSTATS_LIST}')/items?$select=Id&$filter=EventNumber eq ${eventNumber}&$top=50`;
    const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (!resp.ok) return false;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    let ok = true;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    for (const it of (items as any[])) {
      try { await svc._delete(`${svc.siteUrl}/_api/web/lists/getbytitle('${EVENTSTATS_LIST}')/items(${it.Id})`); }
      catch { ok = false; }
    }
    return ok;
  } catch { return false; }
}

/**
 * Löscht die Teilnehmerliste (Subsite) eines Events, LÄSST das DEX_Events-Item
 * bestehen (Event bleibt sichtbar; KPIs stehen im DEX_EventStats). Bereinigt
 * zusätzlich DEX_Participants (EventNumber entfernen). Gegenstück zu
 * deleteEvent(), das auch das Event-Item entfernt.
 */
export async function deleteParticipantData(svc: EventService, eventId: number): Promise<boolean> {
  try {
    const event = await svc.getEvent(eventId);
    if (!event) return false;
    /**
     * v29.3: Reihenfolge umgedreht — erst das Register, dann die
     * unwiderrufliche Löschung.
     *
     * Bisher wurde zuerst die Subsite recycelt und danach das Register
     * aufgeräumt, und zwar so, dass ein Scheitern nicht auffiel:
     * `getAllParticipants()` liest NICHT strikt (eine unvollständige Seite
     * kam still als vollständige Liste zurück), alle Schreibvorgänge liefen
     * als EIN `Promise.all` gleichzeitig los (bei hunderten Teilnehmern die
     * Einladung zur SharePoint-Drosselung), und jeder Fehler wurde per
     * `.catch(() => null)` weggeworfen. Was danach im Register stand, wusste
     * niemand — die Teilnehmerliste war aber schon weg, also ließ sich der
     * Rest auch nicht mehr nachrechnen.
     *
     * Genau das ist der Rückstand, den die Register-Prüfung heute als
     * „Verweis ohne Zeile" meldet. Jetzt gilt: strikt lesen, sequentiell
     * schreiben, jeden Fehlschlag zählen — und bei Fehlern GAR NICHT
     * löschen. Der Aufrufer wertet `false` bereits als „später erneut
     * versuchen" und rollt die Statistik-Zeile zurück; nichts geht verloren,
     * weil die Subsite dann noch steht.
     */
    if (event.EventNumber) {
      let registryFailed = 0;
      try {
        const allParticipants = await svc.fetchAllParticipantsOrThrow();
        const en = String(event.EventNumber);
        const affected = allParticipants.filter(p =>
          (p.EventRegistered?.split(',').map(s => s.trim()).includes(en))
          || (p.EventOnWaitlist?.split(',').map(s => s.trim()).includes(en)));
        for (const p of affected) {
          // eslint-disable-next-line no-await-in-loop
          const ok = await svc.removeParticipantEvent(p.Email, event.EventNumber).catch(() => false);
          if (!ok) registryFailed += 1;
        }
      } catch (err) {
        console.warn('[DEX] deleteParticipantData: Register nicht vollständig lesbar — Löschung abgebrochen:', err);
        return false;
      }
      if (registryFailed > 0) {
        console.warn(`[DEX] deleteParticipantData: ${registryFailed} Register-Einträge nicht aktualisiert — Löschung abgebrochen (Event ${event.EventNumber}).`);
        return false;
      }
    }
    if (event.SubsiteUrl) {
      try { await svc._post(`${event.SubsiteUrl}/_api/web/recycle`, {}); }
      catch { console.warn('[DEX] Teilnehmer-Subsite konnte nicht recycelt werden:', event.SubsiteUrl); }
    }
    if (event.RegistrationListName && event.RegistrationListName !== 'Teilnehmer') {
      try { await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${event.RegistrationListName.replace(/'/g, "''")}')/recycle`, {}); }
      catch { /* */ }
    }
    try {
      await svc.writeChangeLog({
        action: 'ParticipantListDeleted', targetType: 'Event', targetId: String(eventId),
        targetName: event.Title || '', eventId: String(eventId), eventTitle: event.Title || '',
        details: { subsiteUrl: event.SubsiteUrl || '', eventNumber: event.EventNumber, note: 'Teilnehmerliste gelöscht (3-Monats-Löschkonzept); KPIs im DEX_EventStats.' },
      });
    } catch { /* */ }
    return true;
  } catch { return false; }
}

/**
 * v11.69: Löscht NUR das DEX_Events-Listenitem — KEIN Cascade auf Subsite,
 * KEIN Outlook-DeleteEvent in die Queue, KEIN EventImage-Recycle, KEIN
 * DEX_Participants-Cleanup. Gegenstück zu `deleteEvent()`, das alles
 * mit-aufräumt.
 *
 * Nutzungs-Szenario: Outlook-Termin nachträglich auf einem bereits
 * angelegten Sub-Event aktivieren. Der `DEX_CreateOutlookEvent`-Flow
 * triggert ausschliesslich auf NEUE DEX_Events-Items (GetOnNewItems) —
 * ein MERGE-Update reicht nicht aus. Statt das ganze Sub-Event komplett
 * delete+recreate zu machen (was kaskadierend Subsite + Teilnehmer-
 * anmeldungen mitlöschen würde), wird hier nur die DEX_Events-Zeile
 * entfernt und gleich darauf eine neue mit `createEvent({ existingSubsiteUrl,
 * existingRegistrationListName })` angelegt. Die alte Subsite mit allen
 * Anmeldungen bleibt unangetastet und wird einfach an die neue Zeile
 * gekoppelt.
 *
 * **Garantie:** Diese Methode ruft KEIN `recycle()` auf der Subsite, KEIN
 * `recycle()` auf der Teilnehmerliste und KEIN `removeParticipantEvent()`.
 * Nur das DEX_Events-Item wird per REST-DELETE entfernt — alles andere
 * bleibt 1:1 erhalten.
 */
export async function deleteEventItemOnly(svc: EventService, eventId: string | number): Promise<boolean> {
  try {
    const idNum = typeof eventId === 'string' ? parseInt(eventId, 10) : eventId;
    if (!idNum || Number.isNaN(idNum)) return false;
    const response = await svc._delete(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${idNum})`
    );
    return response.ok;
  } catch {
    return false;
  }
}
