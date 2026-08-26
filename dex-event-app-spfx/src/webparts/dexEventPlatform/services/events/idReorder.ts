/**
 * v30.6 — Modularisierung Stufe 2: Thema „DEX_IDReorder-Queue".
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs. Der
 * Session-Merker für die CancelledName-Spalte bleibt als öffentliches
 * Feld an der Service-Instanz (svc._idReorderCancelledFieldEnsured).
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

/**
 * Queue-Liste für ID-Neuvergabe erstellen falls nicht vorhanden.
 * Power Automate reagiert auf neue Einträge und vergibt TeilnehmerIDs
 * auf der jeweiligen Subsite-Teilnehmerliste lückenlos neu.
 *
 * Spalten: Title, EventId, EventNumber, SubsiteUrl,
 * Status (Pending/Processing/Done/Failed), CancelledName.
 */
export async function ensureIDReorderList(svc: EventService): Promise<void> {
  const listName = 'DEX_IDReorder';
  const exists = await svc.listExists(listName);
  if (exists) return;

  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Queue für TeilnehmerID-Neuvergabe nach Abmeldungen',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });

  const fields = [
    { title: 'EventId', type: 2 },
    { title: 'EventNumber', type: 9 },
    { title: 'SubsiteUrl', type: 2 },
    { title: 'Status', type: 6, choices: ['Pending', 'Processing', 'Done', 'Failed'], metaType: 'SP.FieldChoice' },
    // v18.65: Name der abgemeldeten Person (für die Organizer-Nachrücker-Mail).
    { title: 'CancelledName', type: 2 },
  ];

  for (const f of fields) {
    const payload: Record<string, unknown> = {
      '__metadata': { 'type': f.metaType || 'SP.Field' },
      'Title': f.title,
      'FieldTypeKind': f.type,
      'Required': false,
    };
    if ((f as { choices?: string[] }).choices) {
      payload['Choices'] = { 'results': (f as { choices: string[] }).choices };
    }
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
  }

  await svc.configureDefaultView(listName, [
    'EventId', 'EventNumber', 'SubsiteUrl', 'Status',
  ]);

  await svc.setQueueListPermissions(listName);
}

// v18.65: einmal pro Session die CancelledName-Spalte auf DEX_IDReorder
// nachrüsten (Bestands-Listen). Selbstheilend, weil der zentrale
// initEvents-ensure-Pfad bei gesetztem ENSURE_FLAG übersprungen wird.
async function ensureIDReorderCancelledNameField(svc: EventService): Promise<void> {
  if (svc._idReorderCancelledFieldEnsured) return;
  svc._idReorderCancelledFieldEnsured = true;
  // v19.5: CancelledName UND CancelledEmail nachrüsten. CancelledEmail erlaubt
  // dem Nachrück-Flow, die abgemeldete Person eindeutig zu adressieren
  // (Replaced-Audit: ReplacedByParticipantEmail auf der abgemeldeten Person +
  // ReplacedParticipantEmail auf der nachrückenden Person).
  for (const fieldTitle of ['CancelledName', 'CancelledEmail']) {
    try {
      const resp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')/fields/getbytitle('${fieldTitle}')?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) continue; // existiert bereits
    } catch { /* anlegen */ }
    try {
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')/fields`, {
        '__metadata': { 'type': 'SP.Field' }, 'Title': fieldTitle, 'FieldTypeKind': 2, 'Required': false,
      });
    } catch { /* best-effort — Retry-ohne-Feld unten fängt es ab */ }
  }
}

/**
 * ID-Reorder in Queue eintragen (nach Abmeldung).
 */
export async function queueIDReorder(
  svc: EventService,
  eventId: string,
  eventNumber: number,
  subsiteUrl: string,
  eventTitle: string,
  // v18.65: Name der abgemeldeten Person — wird in die Queue geschrieben,
  // damit der DEX_IDReorder-Flow ihn direkt aus dem Trigger lesen kann (statt
  // die „jüngste Abmeldung" abzufragen, was bei gleichzeitigen Abmeldungen
  // während des Flow-Laufs falsch sein könnte). Genutzt für die
  // Organizer-Nachrücker-Mail (OrgNachruecker-Template).
  cancelledName?: string,
  // v19.5: E-Mail der abgemeldeten Person — der Nachrück-Flow nutzt sie für
  // das Replaced-Audit (Hat ersetzt / Wurde ersetzt durch).
  cancelledEmail?: string
): Promise<boolean> {
  try {
    // ListItemEntityTypeFullName dynamisch ermitteln
    let listItemType = 'SP.Data.DEX_x005f_IDReorderListItem';
    try {
      const typeResp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1
      );
      if (typeResp.ok) {
        const typeData = await typeResp.json();
        listItemType = typeData.ListItemEntityTypeFullName || typeData.d?.ListItemEntityTypeFullName || listItemType;
      }
    } catch { /* Fallback auf Standard-Name */ }

    if (cancelledName || cancelledEmail) { try { await ensureIDReorderCancelledNameField(svc); } catch { /* */ } }

    const baseBody: Record<string, unknown> = {
      '__metadata': { 'type': listItemType },
      'Title': `Reorder: ${eventTitle}`,
      'EventId': eventId,
      'EventNumber': eventNumber,
      'SubsiteUrl': subsiteUrl,
      'Status': 'Pending',
    };
    // v19.5: CancelledName + CancelledEmail als optionale Zusatzfelder.
    const extra: Record<string, unknown> = {};
    if (cancelledName) extra['CancelledName'] = cancelledName;
    if (cancelledEmail) extra['CancelledEmail'] = cancelledEmail;
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')/items`;
    let response = await svc._post(url, Object.keys(extra).length ? { ...baseBody, ...extra } : baseBody);
    // Falls die Zusatz-Spalten (noch) fehlen, schlägt der erste POST fehl —
    // dann ohne die Felder erneut posten, damit der Reorder NIEMALS verloren
    // geht (kritischer Pfad).
    if (!response.ok && Object.keys(extra).length) {
      response = await svc._post(url, baseBody);
    }
    return response.ok;
  } catch {
    return false;
  }
}
