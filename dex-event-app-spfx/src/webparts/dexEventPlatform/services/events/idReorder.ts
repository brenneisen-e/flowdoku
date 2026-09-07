/**
 * v30.6 — Modularisierung Stufe 2: Thema „DEX_IDReorder-Queue".
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs. Der
 * Session-Merker für die CancelledName-Spalte bleibt als öffentliches
 * Feld an der Service-Instanz (svc._idReorderCancelledFieldEnsured).
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';
import { addPendingReorder, readPendingReorders, removePendingReorder } from '../../utils/reorderHeal';

/** Ein Reorder-Auftrag, wie ihn der Flow DEX_IDReorder_TeilnehmerIDs liest. */
export interface ReorderJob {
  eventId: string;
  eventNumber: number;
  subsiteUrl: string;
  eventTitle: string;
  cancelledName?: string;
  cancelledEmail?: string;
}

export interface QueueReorderResult {
  ok: boolean;
  /** HTTP-Status des letzten Versuchs (0 = Ausnahme/Netz). */
  status: number;
  attempts: number;
  /** true = als Merker abgelegt, Nachzug beim nächsten Erfolg/App-Start. */
  pending: boolean;
}

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

// v30.80: Der ListItemEntityType ändert sich nie — einmal je Sitzung holen.
// Vorher lief der GET vor JEDEM Auftrag, also unter Drosselung ein zweiter
// Treffer für den 429, bevor der eigentliche POST überhaupt rausging.
let cachedListItemType: string | null = null;

/**
 * Ein einzelner Schreibversuch. Liefert den HTTP-Status mit, damit der
 * geprüfte Pfad unten 429/503 (warten und wiederholen) von 403 (Rechte —
 * Wiederholen ist sinnlos) unterscheiden kann. `retryAfterMs` kommt aus dem
 * Retry-After-Header, den SharePoint bei Drosselung mitschickt.
 */
async function postReorderOnce(svc: EventService, job: ReorderJob): Promise<{ ok: boolean; status: number; retryAfterMs: number }> {
  try {
    let listItemType = cachedListItemType || 'SP.Data.DEX_x005f_IDReorderListItem';
    if (!cachedListItemType) {
      try {
        const typeResp = await svc._sp.get(
          `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')?$select=ListItemEntityTypeFullName`,
          SPHttpClient.configurations.v1
        );
        if (typeResp.ok) {
          const typeData = await typeResp.json();
          listItemType = typeData.ListItemEntityTypeFullName || typeData.d?.ListItemEntityTypeFullName || listItemType;
          cachedListItemType = listItemType;
        }
      } catch { /* Fallback auf Standard-Name */ }
    }

    if (job.cancelledName || job.cancelledEmail) { try { await ensureIDReorderCancelledNameField(svc); } catch { /* */ } }

    const baseBody: Record<string, unknown> = {
      '__metadata': { 'type': listItemType },
      'Title': `Reorder: ${job.eventTitle}`,
      'EventId': job.eventId,
      'EventNumber': job.eventNumber,
      'SubsiteUrl': job.subsiteUrl,
      'Status': 'Pending',
    };
    // v19.5: CancelledName + CancelledEmail als optionale Zusatzfelder.
    const extra: Record<string, unknown> = {};
    if (job.cancelledName) extra['CancelledName'] = job.cancelledName;
    if (job.cancelledEmail) extra['CancelledEmail'] = job.cancelledEmail;
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')/items`;
    let response = await svc._post(url, Object.keys(extra).length ? { ...baseBody, ...extra } : baseBody);
    // Falls die Zusatz-Spalten (noch) fehlen, schlägt der erste POST mit 400
    // fehl — dann ohne die Felder erneut posten. NICHT bei 429/503: dort wäre
    // der zweite POST nur der nächste Drossel-Treffer.
    if (!response.ok && response.status !== 429 && response.status !== 503 && Object.keys(extra).length) {
      response = await svc._post(url, baseBody);
    }
    let retryAfterMs = 0;
    try {
      const ra = response.headers && response.headers.get ? response.headers.get('Retry-After') : null;
      const sec = ra ? parseInt(ra, 10) : NaN;
      if (isFinite(sec) && sec > 0) retryAfterMs = sec * 1000;
    } catch { /* kein Header */ }
    return { ok: response.ok, status: response.status, retryAfterMs };
  } catch {
    return { ok: false, status: 0, retryAfterMs: 0 };
  }
}

/**
 * ID-Reorder in Queue eintragen (nach Abmeldung) — EIN Versuch, ohne Log.
 * Für neue Aufrufer gilt `queueIDReorderChecked`; diese Form bleibt für den
 * Nachzug und für Stellen, die selbst wiederholen.
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
  const r = await postReorderOnce(svc, { eventId, eventNumber, subsiteUrl, eventTitle, cancelledName, cancelledEmail });
  return r.ok;
}

const sleep = (ms: number): Promise<void> => new Promise(resolve => setTimeout(resolve, ms));

/**
 * v30.80: DER Schreibpfad für Reorder-Aufträge — alle zehn Aufrufstellen
 * (Selbst-, Team-, Proxy-, Organizer-, Wartelisten-, Sammel-, Inaktiv-
 * Abmeldung, Gruppenwechsel, Reaktivierung, Nachzug) laufen hier durch.
 *
 * Bis zu vier Versuche mit wachsender Pause (1,5 s · 4 s · 8 s, bei
 * Retry-After entsprechend länger, gedeckelt auf 20 s). Bei 403 wird nicht
 * wiederholt — fehlende Rechte heilen nicht durch Warten. Scheitern alle
 * Versuche: Merker in localStorage (utils/reorderHeal) + Zeile
 * `IDReorderQueueFailed` im Event-Log, damit der Organizer es im Organizer
 * Center sieht. Vorher (v30.76) gab es das nur an drei Stellen mit einem
 * einzigen Wiederholungsversuch nach 1,5 s — in einer Abmelde-Welle mit 27
 * Terminen ist das genau die Pause, die die Drosselung nicht abwartet.
 *
 * Nach jedem Erfolg werden offene Merker anderer Subsites nachgezogen —
 * so heilt die Sitzung sich selbst, sobald SharePoint wieder antwortet.
 */
export async function queueIDReorderChecked(svc: EventService, job: ReorderJob, via: string): Promise<QueueReorderResult> {
  const delays = [1500, 4000, 8000];
  let last = { ok: false, status: 0, retryAfterMs: 0 };
  let attempts = 0;
  for (let i = 0; i <= delays.length; i++) {
    attempts += 1;
    last = await postReorderOnce(svc, job);
    if (last.ok) break;
    if (last.status === 403 || last.status === 401) break;
    if (i < delays.length) await sleep(Math.min(20000, Math.max(delays[i], last.retryAfterMs)));
  }
  if (last.ok) {
    removePendingReorder(job.subsiteUrl);
    // Offene Merker anderer Listen gleich mit nachziehen (SharePoint antwortet
    // gerade) — best-effort, sequentiell, ohne die Aufrufstelle zu bremsen.
    void replayPendingReorders(svc, job.subsiteUrl);
    return { ok: true, status: last.status, attempts, pending: false };
  }
  console.warn(`[DEX] queueIDReorder (${via}) failed after ${attempts} attempts, status ${last.status} — Merker gesetzt`);
  addPendingReorder({ ...job, via, ts: Date.now() });
  svc.writeChangeLog({
    action: 'IDReorderQueueFailed',
    targetType: 'Participant',
    targetId: job.cancelledEmail || '',
    targetName: job.cancelledName || job.cancelledEmail || '',
    eventId: job.eventId,
    eventTitle: job.eventTitle,
    details: {
      via, attempts, httpStatus: last.status,
      hint: last.status === 403
        ? 'Kein Schreibrecht auf DEX_IDReorder (403) — im Admin Center „Berechtigungen reparieren" ausführen, danach „Nachrücken & IDs für ALLE Events nachholen".'
        : 'Kein DEX_IDReorder-Auftrag geschrieben (Drosselung) — die App holt ihn beim nächsten App-Start nach; sonst im Admin Center „Nachrücken & IDs für ALLE Events nachholen" ausführen.',
    },
  }).catch(() => { /* Log ist best-effort */ });
  return { ok: false, status: last.status, attempts, pending: true };
}

/**
 * v30.80: Offene Merker abarbeiten — je Auftrag EIN Versuch (wer hier
 * scheitert, bleibt Merker; 14-Tage-Verfall). `skipSubsiteUrl` = die Liste,
 * für die gerade erfolgreich geschrieben wurde. Liefert die Zahl der
 * nachgeholten Aufträge; jeder bekommt eine Zeile `IDReorderQueueHealed`.
 */
export async function replayPendingReorders(svc: EventService, skipSubsiteUrl?: string): Promise<number> {
  let healed = 0;
  for (const p of readPendingReorders()) {
    if (skipSubsiteUrl && p.subsiteUrl.toLowerCase() === skipSubsiteUrl.toLowerCase()) continue;
    const r = await postReorderOnce(svc, p);
    if (!r.ok) continue;
    healed += 1;
    removePendingReorder(p.subsiteUrl);
    svc.writeChangeLog({
      action: 'IDReorderQueueHealed',
      targetType: 'Participant',
      targetId: p.cancelledEmail || '',
      targetName: p.cancelledName || p.cancelledEmail || '',
      eventId: p.eventId,
      eventTitle: p.eventTitle,
      details: { via: p.via, failedAt: new Date(p.ts).toISOString(), hint: 'Reorder-Auftrag nachgeholt — Neu-Nummerierung und Nachrücken laufen jetzt über den Flow.' },
    }).catch(() => { /* */ });
    await sleep(800);
  }
  return healed;
}
