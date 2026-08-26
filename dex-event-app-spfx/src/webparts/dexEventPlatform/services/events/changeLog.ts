/**
 * v30.6 — Modularisierung Stufe 2: Thema „DEX_ChangeLog" (Audit, v9.0).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 *
 * Audit-Liste für alle Änderungen an Events und Teilnehmern. Read-
 * Berechtigung für Organizer/Admin (gleiche Permission-Pattern wie
 * DEX_Roles), Schreibrechte für alle (damit User-Aktionen wie
 * Anmeldung/Abmeldung mitloggen können).
 */

import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

export interface ChangeLogEntryInput {
  action: string;
  targetType: 'Event' | 'Participant' | 'Subsite' | 'Other';
  targetId?: string;
  targetName?: string;
  eventId?: string;
  eventTitle?: string;
  actorName?: string;
  actorEmail?: string;
  details?: string | Record<string, unknown>;
}

export interface ChangeLogRow {
  Id: number; Created: string; Action: string; TargetType: string;
  TargetId: string; TargetName: string; EventId: string; EventTitle: string;
  ActorName: string; ActorEmail: string; Details: string;
}

export async function ensureChangeLogList(svc: EventService): Promise<void> {
  const listName = 'DEX_ChangeLog';
  try {
    const exists = await svc.listExists(listName);
    if (exists) {
      try { await ensureChangeLogPermissions(svc, listName); } catch { /* */ }
      return;
    }
    // Liste existiert nicht — versuchen sie anzulegen. Schlägt fehl
    // wenn der aktuelle User keine Owner-Permissions hat. Dann wird der
    // App-Start nicht blockiert (Audit-Log ist best-effort für User
    // ohne Schreibrechte auf der Liste-Erstellung).
    const createResp = await svc._post(`${svc.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Audit-Log für alle Änderungen an Events und Teilnehmern (v9.0)',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });
    if (!createResp.ok) {
      console.warn('[DEX] DEX_ChangeLog konnte nicht angelegt werden — vermutlich fehlen dem User Owner-Rechte. App läuft weiter, Audit-Einträge fehlen aber.');
      return;
    }
    const fields = [
      { title: 'Action', type: 6, choices: ['EventCreated', 'EventUpdated', 'EventArchived', 'EventRestored', 'EventDeletedPermanent', 'EventDeletedTest', 'ParticipantRegistered', 'ParticipantCancelled', 'ParticipantReactivated', 'ParticipantUpdated', 'ParticipantCheckedIn', 'ParticipantCheckedOut', 'IDReorder', 'Other'], metaType: 'SP.FieldChoice' },
      { title: 'TargetType', type: 6, choices: ['Event', 'Participant', 'Subsite', 'Other'], metaType: 'SP.FieldChoice' },
      { title: 'TargetId', type: 2 },
      { title: 'TargetName', type: 2 },
      { title: 'EventId', type: 2 },
      { title: 'EventTitle', type: 2 },
      { title: 'ActorName', type: 2 },
      { title: 'ActorEmail', type: 2 },
      { title: 'Details', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 6 },
    ];
    for (const f of fields) {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const payload: Record<string, any> = {
          '__metadata': { 'type': f.metaType || 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        };
        if (f.choices) payload['Choices'] = { 'results': f.choices };
        if (f.metaType === 'SP.FieldMultiLineText') {
          payload['RichText'] = !!f.richText;
          if (typeof f.numberOfLines === 'number') payload['NumberOfLines'] = f.numberOfLines;
        }
        await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
      } catch { /* einzelne Feld-Fehler ignorieren */ }
    }
    try {
      await svc.configureDefaultView(listName, [
        'Created', 'Action', 'TargetType', 'TargetName', 'EventTitle', 'ActorName', 'Details',
      ]);
    } catch { /* View-Setup ist optional */ }
    try { await ensureChangeLogPermissions(svc, listName); } catch { /* */ }
  } catch (err) {
    console.warn('[DEX] ensureChangeLogList failed (best-effort, App läuft weiter):', err);
  }
}

// Berechtigungen: Site-Members und alle authentifizierten User können
// Einträge HINZUFUEGEN (damit Self-Reg/Cancel mitschreibt), aber nur
// Organizer/Admin können LESEN. Setzt Item-Level-Read = "Only their own".
async function ensureChangeLogPermissions(svc: EventService, listName: string): Promise<void> {
  try {
    // v24.77: Permissions nur EINMAL setzen. Sind die Berechtigungen schon
    // eindeutig (Vererbung bereits gebrochen), läuft das komplette Setup
    // unten NICHT erneut. Vorher feuerte breakroleinheritance +
    // addroleassignment bei JEDEM Boot neu → wiederkehrendes 400-Rauschen
    // in der Konsole. Gleiches Muster wie ensureEventsList bei den
    // Event-Permissions.
    try {
      const listInfo = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=HasUniqueRoleAssignments`,
        SPHttpClient.configurations.v1
      );
      if (listInfo.ok) {
        const data = await listInfo.json();
        if (data.HasUniqueRoleAssignments) return; // schon eingerichtet
      }
    } catch { /* im Zweifel weiter unten einrichten */ }
    // 1. Inheritance brechen
    await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
      {}
    );
    // 2. Owners (Admin-Group) → Full Control
    const ownersResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/associatedownergroup?$select=Id`, SPHttpClient.configurations.v1
    );
    if (ownersResp.ok) {
      const d = await ownersResp.json();
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741829)`, {}
      );
    }
    // 3. Members (Organizer-Group typischerweise) → Contribute (sollen
    //    auch schreiben können, damit Organizer-Aktionen wie Event-
    //    Updates protokolliert werden).
    const membersResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/associatedmembergroup?$select=Id`, SPHttpClient.configurations.v1
    );
    if (membersResp.ok) {
      const d = await membersResp.json();
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741827)`, {} // 1073741827 = Contribute
      );
    }
    // 4. Visitors (DEALL / Authenticated Users) → Contribute (damit
    //    User-Aktionen wie Self-Anmeldung mitloggen können).
    const visitorsId = await svc.getVisitorsGroupId();
    if (visitorsId) {
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741827)`, {} // Contribute
      );
    }
    // 5. Item-Level-Read = "ReadAllItems" (1) — Organizer und Admin
    //    müssen ALLE Einträge sehen, nicht nur eigene. Eigene
    //    Lese-Beschränkung wäre für Audit-Log nutzlos.
    // v22.2 FIX: war vorher ein nackter POST (kein MERGE) → SharePoint
    // antwortete bei jedem Boot mit HTTP 400 (Console-Rauschen); der Wert
    // wurde nie gesetzt (Default ist ohnehin 1, daher ohne Folgen).
    // v26.87: MERGE zusätzlich auf nometadata umgestellt (verbose+__metadata
    // war weiterhin 400) — jetzt greift es tatsächlich.
    await svc._setListSecurity(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')`, { ReadSecurity: 1 });
  } catch (e) {
    console.warn('[DEX] ensureChangeLogPermissions failed:', e);
  }
}

/**
 * v9.0: Audit-Eintrag schreiben. Best-effort — Fehler werden nur
 * geloggt, blocken die aufrufende Aktion nie. Wird automatisch von
 * createEvent / updateEvent / deleteEvent / registerForEvent /
 * cancelRegistration / adminUpdateRegistration / etc. gerufen.
 */
export async function writeChangeLog(svc: EventService, entry: ChangeLogEntryInput): Promise<void> {
  try {
    const me = svc.context.pageContext.user;
    const actorName = entry.actorName || me.displayName || '';
    const actorEmail = (entry.actorEmail || me.email || '').toLowerCase();
    const detailsStr = typeof entry.details === 'string'
      ? entry.details
      : entry.details
        ? JSON.stringify(entry.details)
        : '';
    // ENTWICKLUNG.md-Hinweis: bei odata=nometadata KEIN __metadata im Body —
    // SP leitet den Typ aus der URL ab. Robust gegen List-Type-Encoding-
    // Quirks (Bug-Story v7.28 → v7.29). Nutzen wir hier statt verbose-POST
    // damit ein verschmierter Type-Name den ChangeLog-Insert nicht
    // stillschweigend killt.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: Record<string, any> = {
      'Title': `${entry.action}: ${entry.targetName || entry.targetId || '-'}`,
      'Action': entry.action,
      'TargetType': entry.targetType,
      'TargetId': entry.targetId || '',
      'TargetName': entry.targetName || '',
      'EventId': entry.eventId || '',
      'EventTitle': entry.eventTitle || '',
      'ActorName': actorName,
      'ActorEmail': actorEmail,
      'Details': detailsStr.substring(0, 30000),
    };
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_ChangeLog')/items`;
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
      },
      body: JSON.stringify(payload),
    };
    await svc._sp.post(url, SPHttpClient.configurations.v1, options);
  } catch (err) {
    console.warn('[DEX] writeChangeLog failed:', err);
  }
}

/**
 * Audit-Log lesen — Organizer/Admin only (durch SP-Permissions
 * geschützt). Liefert die letzten N Einträge, optional gefiltert
 * nach EventId.
 */
export async function readChangeLog(svc: EventService, opts?: { eventId?: string; top?: number }): Promise<ChangeLogRow[]> {
  const top = opts?.top || 200;
  const base = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_ChangeLog')/items`;
  const sel = `$select=Id,Created,Action,TargetType,TargetId,TargetName,EventId,EventTitle,ActorName,ActorEmail,Details`;
  const filter = opts?.eventId
    ? `&$filter=EventId eq '${String(opts.eventId).replace(/'/g, "''")}'`
    : '';
  // v19.13 BUG-FIX: Das Audit-Log lud mit HTTP 400 (→ „0 Einträge"). Zwei
  // typische Ursachen, gegen beide robust:
  //  (a) `$orderby=Created` auf einer großen Liste ohne Index auf `Created`
  //      → „List View Threshold"-Fehler (400). Deshalb jetzt nach `Id desc`
  //      sortieren — `Id` ist IMMER indiziert, und da auto-increment entspricht
  //      die Reihenfolge chronologisch absteigend.
  //  (b) Ein Feld im `$select` existiert auf einer Bestands-/Legacy-Liste nicht
  //      (Feld-Anlage best-effort) → 400. Deshalb Fallbacks ohne `$select` und
  //      notfalls ohne Server-Filter (dann client-seitig nach EventId filtern).
  const candidates = [
    `${base}?${sel}&$orderby=Id desc&$top=${top}${filter}`,
    `${base}?$orderby=Id desc&$top=${top}${filter}`,
    `${base}?$orderby=Id desc&$top=${top}`,
  ];
  for (let i = 0; i < candidates.length; i++) {
    try {
      const resp = await svc._sp.get(candidates[i], SPHttpClient.configurations.v1);
      if (!resp.ok) continue;
      const data = await resp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let items: any[] = data.value || data.d?.results || [];
      // Letzter Kandidat hat keinen Server-Filter (falls die EventId-Spalte
      // im Filter das 400 ausgelöst hat) → client-seitig nachfiltern.
      if (i === candidates.length - 1 && opts?.eventId) {
        items = items.filter(it => String((it && it.EventId) || '') === String(opts.eventId));
      }
      return items;
    } catch { /* nächsten Kandidaten versuchen */ }
  }
  console.warn('[DEX] readChangeLog: alle Abfrage-Varianten fehlgeschlagen.');
  return [];
}
