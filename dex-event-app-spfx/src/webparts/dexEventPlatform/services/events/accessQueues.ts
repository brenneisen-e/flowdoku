/**
 * v30.66 — Modularisierung Stufe 2: Thema „Zugriffs-Queues": DEX_AccessFix
 * (der Flow trägt die Person nachträglich als Leser auf ihre eigene
 * Anmeldezeile ein) und DEX_AssistantAccess (Assistenz-Verknüpfungen samt
 * Änderungs-/Storno-Anträgen, Flow DEX_AssistantAccess_Grant, v24.41).
 *
 * Beides sind Warteschlangen für Power Automate: Die App darf die Rechte
 * nicht selbst setzen, sie legt nur einen Auftrag ab.
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService, AssistantLink } from '../EventService';

/**
 * v18.48: Sperr-Liste für den Outlook-Einladungs-Flow (DEX_Outlook_Einladungen).
 *
 * Hintergrund: der Einladungs-Flow patcht pro Anmeldung/Abmeldung die
 * KOMPLETTE Teilnehmerliste eines Outlook-Termins (bei Grossevents bis zu
 * 1500 Personen) an Microsoft Graph. Lief der Flow seriell (Concurrency 1),
 * standen Anmeldungen für völlig UNTERSCHIEDLICHE Events stundenlang in
 * der Warteschlange. Lösung „Option B": die Trigger-Concurrency wird hoch-
 * gesetzt (z.B. 25 parallele Läufe), und ein Pro-Event-Lock verhindert,
 * dass zwei Läufe für dasSELBE Event gleichzeitig die Teilnehmerliste
 * lesen-und-schreiben (Race -> verlorene Einträge).
 *
 * Die Liste hat eine eindeutige (Enforce-Unique) Spalte `EventId`. Der Flow
 * „erwirbt" den Lock per Create-Item: gelingt das Create, hat er den Lock;
 * schlägt es wegen der Eindeutigkeits-Prüfung fehl, hält gerade ein
 * anderer Lauf desselben Events den Lock -> kurz warten und erneut
 * versuchen. Am Ende löscht der Flow das Lock-Item wieder (Release).
 *
 * Die UI-Schritt-für-Schritt-Anleitung steht in `docs/flow-jsons.md` unter
 * „UI-Anleitung 2026-06-02 (v18.48) — Option B: Pro-Event-Lock für
 * parallele Outlook-Läufe".
 */
/**
 * v20.7: Queue-Liste `DEX_AccessFix` (Site-Collection-Root) für den
 * Assistenz-Fall der Fremd-Anmeldung. Wenn `trySetItemAuthor` mangels
 * „Listen verwalten"-Rechten scheitert (normaler Contribute-User, z.B.
 * Assistenz meldet einen Partner an), schreibt die App hier einen
 * Auftrag — der Power-Automate-Flow `DEX_AccessFix_Autor` (läuft mit
 * Service-Identität, hat Full Control) setzt dann den Zeilen-Autor auf
 * den Teilnehmer und markiert den Auftrag als Done/Failed.
 * Spalten: SubsiteUrl (Text), ItemId (Number), ParticipantEmail (Text),
 * Status (Text: Pending/Done/Failed). Schreibrechte für alle User via
 * setQueueListPermissions (analog DEX_Emails).
 */
export async function ensureAccessFixList(svc: EventService): Promise<void> {
  const listName = 'DEX_AccessFix';
  const exists = await svc.listExists(listName);
  if (exists) return;

  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Queue: Zeilen-Autor bei Fremd-Anmeldungen auf den Teilnehmer setzen (v20.7, Flow DEX_AccessFix_Autor).',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });
  const fields: Array<{ title: string; type: number }> = [
    { title: 'SubsiteUrl', type: 2 },
    { title: 'ItemId', type: 9 },
    { title: 'ParticipantEmail', type: 2 },
    { title: 'Status', type: 2 },
  ];
  for (const f of fields) {
    try {
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      });
    } catch { /* einzelne Feld-Fehler ignorieren */ }
  }
  try {
    await svc.configureDefaultView(listName, ['SubsiteUrl', 'ItemId', 'ParticipantEmail', 'Status', 'Created']);
  } catch { /* View optional */ }
  try {
    await svc.setQueueListPermissions(listName);
  } catch { /* best-effort */ }
  // Item-Level-Security: User sehen/ändern nur EIGENE Aufträge (sonst wäre
  // ablesbar, wer wen angemeldet hat). Der Flow läuft als Site-Owner und
  // sieht alle Items („Listen verwalten" hebelt die Item-Beschränkung aus).
  // v26.87: zuverlässiger nometadata-MERGE (verbose+__metadata → HTTP 400).
  await svc._setListSecurity(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')`, { ReadSecurity: 2, WriteSecurity: 2 });
}

/**
 * v20.7: Auftrag für den DEX_AccessFix_Autor-Flow einreihen (siehe
 * ensureAccessFixList). Best-effort — Fehler blocken die Anmeldung nie.
 */
export async function queueAccessFix(svc: EventService, subsiteUrl: string, itemId: number, participantEmail: string): Promise<void> {
  try {
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_AccessFix')/items`, {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_AccessFixListItem' },
      'Title': `${participantEmail} -> Item ${itemId}`.slice(0, 250),
      'SubsiteUrl': subsiteUrl,
      'ItemId': itemId,
      'ParticipantEmail': participantEmail,
      'Status': 'Pending',
    });
  } catch (err) {
    console.warn('[DEX] queueAccessFix fehlgeschlagen (best-effort):', err);
  }
}

/**
 * v24.41: Liste `DEX_AssistantAccess` — Delegations-/Zugriffs-Queue für die
 * „Meine Assistenz beauftragen"-Funktion. Wenn ein Admin/Director sich für
 * ein Event anmeldet und eine Assistenz angibt, wird hier ein Eintrag
 * angelegt. Zwei Zwecke:
 *  1. **Flow-Auftrag:** Der Flow `DEX_AssistantAccess_Grant` setzt den
 *     Zeilen-Autor (`Created By`) der Teilnehmer-Anmeldung auf die Assistenz,
 *     damit diese die Anmeldung in ihrer „Assistenz"-Kachel sieht/bearbeitet
 *     (unter „nur eigene Elemente" geht das NUR über den Autor — siehe
 *     Recherche v24.41). Für Admins setzt die App den Autor direkt; für
 *     normale Directoren erledigt es der Flow.
 *  2. **Info-Zeile für den Director:** Da der Director nach der Delegation
 *     nicht mehr Autor der Teilnehmer-Zeile ist (und sie unter ILS nicht mehr
 *     sieht), liest „Meine Events" hier (der Director ist Autor SEINES
 *     Delegations-Eintrags → sieht ihn) eine schreibgeschützte Zeile
 *     „Angemeldet für X — verwaltet von Assistenz Y".
 * ReadSecurity/WriteSecurity=2: jeder sieht nur die EIGENEN Delegations-
 * Einträge; der Flow (Site-Owner) sieht alle.
 */
export async function ensureAssistantAccessList(svc: EventService): Promise<void> {
  const listName = 'DEX_AssistantAccess';
  const exists = await svc.listExists(listName);
  if (exists) return;

  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Queue + Info: Anmeldung an eine Assistenz delegieren — Zeilen-Autor auf die Assistenz setzen (v24.41, Flow DEX_AssistantAccess_Grant).',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });
  const fields: Array<{ title: string; type: number }> = [
    { title: 'SubsiteUrl', type: 2 },
    { title: 'ItemId', type: 9 },
    { title: 'EventId', type: 2 },
    { title: 'EventTitle', type: 2 },
    { title: 'ParticipantEmail', type: 2 },   // die angemeldete Person
    { title: 'ParticipantName', type: 2 },
    { title: 'AssistantEmail', type: 2 },     // die verknüpfte Assistenz
    { title: 'AssistantName', type: 2 },
    { title: 'OwnerEmail', type: 2 },         // wer die Anmeldung VERWALTET
    { title: 'LinkType', type: 2 },           // 'delegation' (Selbst-Anmeldung+Assistenz) | 'proxy' (Assistenz meldet an)
    { title: 'Status', type: 2 },             // 'Active' | 'Cancelled'
    { title: 'RequestType', type: 2 },        // '' | 'change' | 'cancel'
    { title: 'RequestNote', type: 3 },        // Note
    { title: 'RequestedByEmail', type: 2 },
    { title: 'RequestedByName', type: 2 },
    { title: 'RequestStatus', type: 2 },      // '' | 'Open' | 'Done' | 'Rejected'
  ];
  for (const f of fields) {
    try {
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      });
    } catch { /* einzelne Feld-Fehler ignorieren */ }
  }
  try {
    await svc.configureDefaultView(listName, ['SubsiteUrl', 'ItemId', 'EventTitle', 'ParticipantEmail', 'AssistantEmail', 'OwnerEmail', 'LinkType', 'Status', 'RequestType', 'RequestStatus', 'Created']);
  } catch { /* View optional */ }
  try {
    await svc.setQueueListPermissions(listName);
  } catch { /* best-effort */ }
  // v24.41: ReadSecurity=1/WriteSecurity=1 — jeder darf lesen UND schreiben.
  // Die Liste enthält bewusst NUR Koordinations-Daten (Verknüpfung Person ↔
  // Assistenz + Anforderungs-Status), KEINE sensiblen Anmelde-Antworten (die
  // bleiben ILS-geschützt auf der Subsite). So sehen beide Seiten ihre
  // relevanten Einträge (App filtert) und können Anforderungen schreiben —
  // alles ohne Flow.
  // v26.87: zuverlässiger nometadata-MERGE (verbose+__metadata → HTTP 400).
  await svc._setListSecurity(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')`, { ReadSecurity: 1, WriteSecurity: 1 });
}

/**
 * v24.41: Verknüpfungs-/Koordinations-Eintrag in DEX_AssistantAccess anlegen.
 * `ownerEmail` = wer die Anmeldung verwaltet (Owner). `linkType`:
 * 'delegation' (Person meldet sich selbst an + benennt Assistenz) oder
 * 'proxy' (Assistenz meldet die Person an). Best-effort.
 */
export async function queueAssistantAccess(svc: EventService, args: {
  subsiteUrl: string; itemId: number; eventId: string; eventTitle: string;
  participantEmail: string; participantName: string; assistantEmail: string; assistantName: string;
  ownerEmail: string; linkType: 'delegation' | 'proxy';
}): Promise<void> {
  try {
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items`, {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_AssistantAccessListItem' },
      'Title': `${args.participantEmail} <-> ${args.assistantEmail} (${args.eventTitle || args.eventId})`.slice(0, 250),
      'SubsiteUrl': args.subsiteUrl,
      'ItemId': args.itemId,
      'EventId': args.eventId,
      'EventTitle': (args.eventTitle || '').slice(0, 250),
      'ParticipantEmail': args.participantEmail,
      'ParticipantName': (args.participantName || '').slice(0, 250),
      'AssistantEmail': args.assistantEmail,
      'AssistantName': (args.assistantName || '').slice(0, 250),
      'OwnerEmail': args.ownerEmail,
      'LinkType': args.linkType,
      'Status': 'Active',
      'RequestStatus': '',
    });
  } catch (err) {
    console.warn('[DEX] queueAssistantAccess fehlgeschlagen (best-effort):', err);
  }
}

/**
 * v24.41: Alle AKTIVEN Verknüpfungen lesen, die den eingeloggten User
 * betreffen (als angemeldete Person, als verknüpfte Assistenz ODER als Owner).
 * ReadSecurity=1 — der Server liefert alle; wir filtern serverseitig auf die
 * drei Email-Felder. Die App kategorisiert danach (Info-Ansichten + offene
 * Anforderungen an den Owner).
 */
export async function getAssistantLinksForUser(svc: EventService, myEmail: string): Promise<AssistantLink[]> {
  const out: AssistantLink[] = [];
  const me = (myEmail || '').toLowerCase().trim();
  if (!me) return out;
  const esc = me.replace(/'/g, "''");
  const filter = `Status eq 'Active' and (ParticipantEmail eq '${esc}' or AssistantEmail eq '${esc}' or OwnerEmail eq '${esc}')`;
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items?$select=Id,SubsiteUrl,ItemId,EventId,EventTitle,ParticipantEmail,ParticipantName,AssistantEmail,AssistantName,OwnerEmail,LinkType,Status,RequestType,RequestNote,RequestedByEmail,RequestedByName,RequestStatus,Created&$filter=${encodeURIComponent(filter)}&$orderby=Created desc&$top=500`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return out;
    const data = await resp.json();
    const items: Array<Record<string, string | number>> = data.value || data.d?.results || [];
    for (const it of items) {
      out.push({
        id: Number(it.Id) || 0,
        subsiteUrl: String(it.SubsiteUrl || ''),
        itemId: Number(it.ItemId) || 0,
        eventId: String(it.EventId || ''),
        eventTitle: String(it.EventTitle || ''),
        participantEmail: String(it.ParticipantEmail || ''),
        participantName: String(it.ParticipantName || ''),
        assistantEmail: String(it.AssistantEmail || ''),
        assistantName: String(it.AssistantName || ''),
        ownerEmail: String(it.OwnerEmail || ''),
        linkType: String(it.LinkType || ''),
        status: String(it.Status || ''),
        requestType: String(it.RequestType || ''),
        requestNote: String(it.RequestNote || ''),
        requestedByEmail: String(it.RequestedByEmail || ''),
        requestedByName: String(it.RequestedByName || ''),
        requestStatus: String(it.RequestStatus || ''),
        created: String(it.Created || ''),
      });
    }
  } catch { /* best-effort */ }
  return out;
}

/**
 * v24.42: Eine Änderungs-/Abmelde-Anforderung auf einem Link setzen
 * (RequestType/RequestNote/RequestStatus=Open + Anforderer). Schreibbar von
 * jedem (WriteSecurity=1) — der nicht-Owner stellt die Anforderung.
 */
export async function setAssistantLinkRequest(svc: EventService, linkId: number, args: {
  requestType: 'change' | 'cancel'; note: string; requestedByEmail: string; requestedByName: string;
}): Promise<boolean> {
  try {
    const resp = await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items(${linkId})`,
      {
        'RequestType': args.requestType,
        'RequestNote': (args.note || '').slice(0, 1000),
        'RequestedByEmail': args.requestedByEmail,
        'RequestedByName': (args.requestedByName || '').slice(0, 250),
        'RequestStatus': 'Open',
      }
    );
    return resp.ok;
  } catch { return false; }
}

/** v24.42: Anforderung als erledigt/abgelehnt markieren (Owner). */
export async function resolveAssistantLinkRequest(svc: EventService, linkId: number, decision: 'Done' | 'Rejected'): Promise<boolean> {
  try {
    const resp = await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items(${linkId})`,
      { 'RequestStatus': decision }
    );
    return resp.ok;
  } catch { return false; }
}

/** v24.41: Link beim Abmelden auf 'Cancelled' setzen (Info verschwindet). */
export async function setAssistantLinkStatusForRegistration(svc: EventService, itemId: number, subsiteUrl: string, status: 'Cancelled'): Promise<void> {
  try {
    const esc = (subsiteUrl || '').replace(/'/g, "''");
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items?$select=Id&$filter=ItemId eq ${itemId} and SubsiteUrl eq '${esc}' and Status eq 'Active'&$top=20`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return;
    const data = await resp.json();
    const items: Array<{ Id: number }> = data.value || data.d?.results || [];
    for (const it of items) {
      try { await svc._merge(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_AssistantAccess')/items(${it.Id})`, { 'Status': status }); } catch { /* */ }
    }
  } catch { /* best-effort */ }
}
