/**
 * v30.66 — Modularisierung Stufe 2: Thema „Team-Beitritts-Anfragen"
 * (DEX_TeamJoinRequests, v11.83).
 *
 * Die Liste liegt bewusst auf Site-Collection-Ebene und nicht pro Subsite:
 * So findet ein Team-Lead alle offenen Anfragen in EINER Abfrage. Sie ist
 * die einzige Queue-Liste OHNE Item-Level-Security — fremde Anfragen muss
 * der Lead lesen können (siehe CLAUDE.md).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

// ==================== DEX_TeamJoinRequests (v11.83) ====================

/**
 * v11.83: Globale Liste für Team-Beitritts-Anfragen (Approve-Queue).
 * Liegt auf der Site-Collection-Ebene (nicht pro Subsite), damit alle
 * Events darauf zugreifen können und der Team-Lead alle ausstehenden
 * Anfragen in einer einzigen Query findet.
 *
 * Spalten:
 * - Title: Anzeige-Zusammenfassung "RequesterName -> Event-Title"
 * - EventId: ID des Events in DEX_Events
 * - TeamId: UUID der Team-Anmeldung
 * - RequesterEmail: Email des Anfragenden
 * - RequesterDisplayName: Anzeigename des Anfragenden
 * - Status: Pending / Approved / Rejected
 * - DecidedDate: Wann hat der Team-Lead entschieden
 * - DecidedByEmail: Email des entscheidenden Leads
 */
export async function ensureTeamJoinRequestsList(svc: EventService): Promise<void> {
  const listName = 'DEX_TeamJoinRequests';
  const exists = await svc.listExists(listName);
  if (exists) {
    // v13.0: Backfill für ältere Installationen, die die Liste vor
    // v11.83 angelegt haben (DecidedDate/DecidedByEmail damals nicht
    // vorhanden). Ohne diesen Patch schlägt decideTeamJoinRequest
    // beim MERGE auf die fehlenden Felder mit HTTP 400 fehl.
    await ensureMissingTeamJoinRequestsFields(svc, listName);
    return;
  }

  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Approve-Queue für Team-Beitritts-Anfragen (v11.83+).',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });

  const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
    { title: 'EventId', type: 2 },
    { title: 'TeamId', type: 2 },
    { title: 'RequesterEmail', type: 2 },
    { title: 'RequesterDisplayName', type: 2 },
    { title: 'Status', type: 6, choices: ['Pending', 'Approved', 'Rejected'], metaType: 'SP.FieldChoice' },
    { title: 'DecidedDate', type: 4 },
    { title: 'DecidedByEmail', type: 2 },
    // v18.73: event-spezifische Antworten des Anfragenden als JSON — werden
    // beim Approve auf den neuen Member angewandt (Note = multi-line text).
    { title: 'CustomData', type: 3 },
  ];

  for (const f of fields) {
    const payload: Record<string, unknown> = {
      '__metadata': { 'type': f.metaType || 'SP.Field' },
      'Title': f.title,
      'FieldTypeKind': f.type,
      'Required': false,
    };
    if (f.choices) payload['Choices'] = { 'results': f.choices };
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
  }

  await svc.configureDefaultView(listName, [
    'EventId', 'TeamId', 'RequesterEmail', 'RequesterDisplayName',
    'Status', 'Created', 'DecidedDate', 'DecidedByEmail',
  ]);

  // Schreibrechte für alle Authentifizierten (analog zu DEX_Emails-Queue):
  // jeder darf eine Anfrage erstellen, aber Item-Level-Security greift
  // sowieso über den Lead-Check beim Approve-Pfad.
  try {
    await svc.setQueueListPermissions(listName);
  } catch { /* best-effort */ }
}

/**
 * v13.0: Backfill fehlender Felder in einer bestehenden DEX_TeamJoinRequests-
 * Liste. Greift bei Tenants die die Liste vor v11.83 angelegt haben.
 */
async function ensureMissingTeamJoinRequestsFields(svc: EventService, listName: string): Promise<void> {
  const wanted = [
    { title: 'DecidedDate', type: 4 },
    { title: 'DecidedByEmail', type: 2 },
    // v18.73: CustomData-Spalte auf Bestands-Listen nachziehen.
    { title: 'CustomData', type: 3 },
  ];
  for (const f of wanted) {
    try {
      const resp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${f.title}')?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) continue; // existiert
    } catch { /* anlegen */ }
    try {
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      });
    } catch (e) {
      console.warn(`[DEX] ensureMissingTeamJoinRequestsFields: failed to add '${f.title}':`, e);
    }
  }
}

/**
 * v11.83: Neue Team-Beitritts-Anfrage anlegen.
 */
export async function createTeamJoinRequest(svc: EventService, args: {
  eventId: string;
  eventTitle: string;
  teamId: string;
  requesterEmail: string;
  requesterDisplayName: string;
  // v18.73: event-spezifische Antworten als JSON (optional).
  customData?: string;
}): Promise<{ ok: boolean; itemId?: number }> {
  try {
    const payload = {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_TeamJoinRequestsListItem' },
      'Title': `${args.requesterDisplayName} -> ${args.eventTitle}`.slice(0, 250),
      'EventId': args.eventId,
      'TeamId': args.teamId,
      'RequesterEmail': args.requesterEmail,
      'RequesterDisplayName': args.requesterDisplayName,
      'Status': 'Pending',
      ...(args.customData ? { 'CustomData': args.customData } : {}),
    };
    const resp = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_TeamJoinRequests')/items`,
      payload
    );
    if (!resp.ok) return { ok: false };
    try {
      const j = await resp.json();
      const id: number = j?.d?.Id || j?.Id || 0;
      return { ok: true, itemId: id };
    } catch {
      return { ok: true };
    }
  } catch {
    return { ok: false };
  }
}

/**
 * v11.83: Alle Pending-Beitritts-Anfragen — optional gefiltert nach
 * Event und/oder Team. Wird für die "Beitritts-Anfragen"-Box im
 * Team-Lead-UI in MyEventsPage aufgerufen.
 */
export async function listTeamJoinRequests(svc: EventService, args: {
  eventId?: string;
  teamId?: string;
  status?: 'Pending' | 'Approved' | 'Rejected';
}): Promise<Array<{ Id: number; EventId: string; TeamId: string; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string; DecidedDate?: string; DecidedByEmail?: string; CustomData?: string }>> {
  try {
    const clauses: string[] = [];
    if (args.eventId) clauses.push(`EventId eq '${args.eventId.replace(/'/g, "''")}'`);
    if (args.teamId) clauses.push(`TeamId eq '${args.teamId.replace(/'/g, "''")}'`);
    clauses.push(`Status eq '${args.status || 'Pending'}'`);
    const filter = clauses.join(' and ');
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_TeamJoinRequests')/items?$filter=${encodeURIComponent(filter)}&$top=200&$orderby=Created asc`;
    const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (!resp.ok) return [];
    const data = await resp.json();
    return data.value || data.d?.results || [];
  } catch {
    return [];
  }
}

/**
 * v11.83: Approve/Reject einer Beitritts-Anfrage — schreibt Status,
 * DecidedDate und DecidedByEmail. Die Folge-Logik (Member-Insert,
 * Mails) liegt im EventContext, weil dort die Subsite-/Event-Lookups
 * verfügbar sind.
 */
export async function decideTeamJoinRequest(
  svc: EventService,
  requestId: number,
  decision: 'Approved' | 'Rejected',
  decidedByEmail: string
): Promise<boolean> {
  try {
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_TeamJoinRequests')/items(${requestId})`;
    const body = {
      'Status': decision,
      'DecidedDate': new Date().toISOString(),
      'DecidedByEmail': decidedByEmail || '',
    };
    const resp = await svc._merge(url, body);
    return !!resp.ok;
  } catch {
    return false;
  }
}
