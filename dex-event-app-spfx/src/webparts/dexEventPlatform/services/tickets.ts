/**
 * v28.95: Ticketsystem aus `EventService` herausgeloest — der erste Schnitt
 * der Klasse nach Thema (v26.0.0 eingefuehrt, 463 Zeilen).
 *
 * Das Rezept, das auch fuer die weiteren Themen gilt: Die Klasse behaelt ihre
 * oeffentlichen Methoden, sie delegieren nur noch. Hier stehen freie
 * Funktionen, die den Service als ersten Parameter bekommen — dadurch aendert
 * sich an keiner Aufrufstelle etwas, und der Compiler prueft den ganzen Weg.
 *
 * `TicketsHost` beschreibt genau, was diese Funktionen vom Service brauchen:
 * die SharePoint-Primitive und den Kontext. Mehr sehen sie nicht.
 */
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DexTicket, TicketAttachment, TicketFollowUp } from '../types';

export interface TicketsHost {
  siteUrl: string;
  context: WebPartContext;
  listExists(listName: string): Promise<boolean>;
  setQueueListPermissions(listName: string): Promise<void>;
  configureDefaultView(listName: string, fieldNames: string[], baseUrl?: string, opts?: { rebuild?: boolean }): Promise<void>;
  _post(url: string, body: object): Promise<SPHttpClientResponse>;
  _merge(url: string, body: object): Promise<SPHttpClientResponse>;
  _mergeIfMatch(url: string, body: object, etag: string): Promise<SPHttpClientResponse>;
}

// ==================== Ticketsystem (v26.0.0) ====================
// Globale Liste DEX_Tickets (Site-Collection-Root). Nutzer stellen über den
// grünen „Hast du Fragen?"-Header-Button Fragen; Power-User/Admins bzw. die
// Organizer des betroffenen Events beantworten sie. Queue-Schreibrechte wie
// DEX_Emails — bewusst KEINE Item-Level-Security (analog DEX_TeamJoinRequests:
// die Beantwortenden müssen fremde Tickets lesen können). Screenshots als
// Item-Attachments mit Namens-Präfix ask_ (Fragesteller) bzw. ans_ (Antwort).

const TICKETS_LIST = 'DEX_Tickets';
const TICKETS_ITEM_TYPE = 'SP.Data.DEX_x005f_TicketsListItem';

/** Ticket-Liste anlegen (idempotent), Felder + Queue-Schreibrechte. */
export async function ensureTicketsList(svc: TicketsHost): Promise<void> {
  const listName = TICKETS_LIST;
  const exists = await svc.listExists(listName);
  if (exists) { await ensureTicketExtraFields(svc); return; }
  const createResp = await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Ticketsystem (v26): Fragen der Nutzer + Antworten der Power-User/Organizer.',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });
  if (!createResp.ok) {
    console.warn('[DEX] DEX_Tickets konnte nicht angelegt werden — vermutlich fehlen Owner-Rechte.');
    return;
  }
  const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string; note?: boolean }> = [
    { title: 'Questions', type: 3, note: true },
    { title: 'Status', type: 6, choices: ['Open', 'InProgress', 'Closed'], metaType: 'SP.FieldChoice' },
    { title: 'AskerEmail', type: 2 },
    { title: 'AskerName', type: 2 },
    { title: 'AskerRole', type: 2 },
    { title: 'Audience', type: 2 },
    { title: 'TicketEventId', type: 2 },
    { title: 'TicketEventTitle', type: 2 },
    { title: 'AssignedOrganizers', type: 3, note: true },
    { title: 'PageContext', type: 2 },
    { title: 'AskWizardStep', type: 9 },
    { title: 'Category', type: 2 }, // v26.60: 'Question' | 'Bug'
    { title: 'AnswerWizardMarker', type: 2 }, // v26.52: Markierungsbox (JSON {x,y,w,h} in %) auf der Wizard-Vorschau
    { title: 'AnswerText', type: 3, note: true },
    { title: 'AnswerArticleIds', type: 3, note: true },
    { title: 'AnswerWizardStep', type: 9 },
    { title: 'AnsweredByEmail', type: 2 },
    { title: 'AnsweredByName', type: 2 },
    { title: 'AnsweredAt', type: 4 },
    { title: 'ClaimedByEmail', type: 2 },
    { title: 'ClaimedByName', type: 2 },
    { title: 'ClaimedAt', type: 4 },
    // v26.8: Standort/Position für die Foto-Kontaktkarte + Rückfragen-Verlauf.
    { title: 'AskerLocation', type: 2 },
    { title: 'AskerJobTitle', type: 2 },
    { title: 'AnsweredByLocation', type: 2 },
    { title: 'AnsweredByJobTitle', type: 2 },
    { title: 'FollowUps', type: 3, note: true },
  ];
  for (const f of fields) {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': f.metaType || (f.note ? 'SP.FieldMultiLineText' : 'SP.Field') },
        'Title': f.title, 'FieldTypeKind': f.type, 'Required': false,
      };
      if (f.choices) payload['Choices'] = { 'results': f.choices };
      if (f.note) { payload['RichText'] = false; payload['NumberOfLines'] = 8; }
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    } catch { /* einzelne Feld-Fehler ignorieren */ }
  }
  try {
    await svc.configureDefaultView(listName, [
      'Status', 'AskerName', 'AskerRole', 'Audience', 'TicketEventTitle', 'Created', 'AnsweredAt',
    ]);
  } catch { /* View optional */ }
  try { await svc.setQueueListPermissions(listName); } catch { /* best-effort */ }
}

/** v26.8: fehlende Felder auf einer bereits existierenden DEX_Tickets-Liste
 *  nachziehen.
 *  v26.61 BUG-FIX: Vorher diente AskWizardStep als Sentinel („vorhanden →
 *  nichts zu tun") — dadurch wurden JÜNGERE Spalten (AnswerWizardMarker
 *  v26.52, Category v26.60) auf Bestandslisten NIE angelegt; der Ticket-
 *  Select lief in HTTP 400 und die Tickets-Seite blieb komplett leer.
 *  Jetzt: echter Feld-Diff — vorhandene InternalNames laden und nur
 *  Fehlendes anlegen (Sentinel-Muster ist für wachsende Listen tabu). */
async function ensureTicketExtraFields(svc: TicketsHost): Promise<void> {
  const listName = TICKETS_LIST;
  const extra: Array<{ title: string; type: number; note?: boolean }> = [
    { title: 'AskerLocation', type: 2 },
    { title: 'AskerJobTitle', type: 2 },
    { title: 'AnsweredByLocation', type: 2 },
    { title: 'AnsweredByJobTitle', type: 2 },
    { title: 'FollowUps', type: 3, note: true },
    { title: 'AskWizardStep', type: 9 },
    { title: 'Category', type: 2 }, // v26.60: 'Question' | 'Bug'
    { title: 'AnswerWizardMarker', type: 2 }, // v26.52: Markierungsbox (JSON {x,y,w,h} in %) auf der Wizard-Vorschau
  ];
  let existing: Set<string> | null = null;
  try {
    const resp = await svc.context.spHttpClient.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName&$filter=Hidden eq false&$top=200`,
      SPHttpClient.configurations.v1);
    if (resp.ok) {
      const d = await resp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      existing = new Set(((d.value || d.d?.results || []) as any[]).map(f => String(f.InternalName)));
    }
  } catch { /* Diff nicht lesbar → alle versuchen (Duplikate schlagen einzeln fehl) */ }
  for (const f of extra) {
    if (existing && existing.has(f.title)) continue;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': f.note ? 'SP.FieldMultiLineText' : 'SP.Field' },
        'Title': f.title, 'FieldTypeKind': f.type, 'Required': false,
      };
      if (f.note) { payload['RichText'] = false; payload['NumberOfLines'] = 8; }
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    } catch { /* einzelne Feld-Fehler ignorieren */ }
  }
}

/** Neues Ticket anlegen. Liefert die Item-Id zurück (für Attachment-Upload). */
export async function createTicket(svc: TicketsHost, t: {
  questions: string[];
  askerEmail: string; askerName: string; askerRole: string;
  askerLocation?: string; askerJobTitle?: string;
  audience: string; eventId: string; eventTitle: string;
  assignedOrganizers: string[]; pageContext: string;
  askWizardStep?: number | null;
  /** v26.60: 'bug' = Bug-Report (Benachrichtigung an die DEX-Maintainer
   *  statt an alle Power-User); sonst inhaltliche Frage. */
  category?: 'question' | 'bug';
}): Promise<number | null> {
  try {
    const first = (t.questions[0] || 'Frage').replace(/\s+/g, ' ').trim().slice(0, 240) || 'Frage';
    const resp = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${TICKETS_LIST}')/items`,
      {
        '__metadata': { 'type': TICKETS_ITEM_TYPE },
        'Title': first,
        'Questions': JSON.stringify(t.questions || []),
        'Status': 'Open',
        'AskerEmail': t.askerEmail, 'AskerName': t.askerName, 'AskerRole': t.askerRole,
        'AskerLocation': t.askerLocation || '', 'AskerJobTitle': t.askerJobTitle || '',
        'Audience': t.audience,
        'TicketEventId': t.eventId || '', 'TicketEventTitle': t.eventTitle || '',
        'AssignedOrganizers': JSON.stringify(t.assignedOrganizers || []),
        'PageContext': t.pageContext || '',
        'AskWizardStep': (t.askWizardStep == null) ? null : t.askWizardStep,
        'Category': t.category === 'bug' ? 'Bug' : 'Question',
      }
    );
    if (!resp.ok) return null;
    const data = await resp.json();
    const id = data?.d?.Id ?? data?.Id;
    return typeof id === 'number' ? id : (id != null ? Number(id) : null);
  } catch (err) {
    console.warn('[DEX] createTicket failed:', err);
    return null;
  }
}

/** Screenshot als Item-Attachment anhängen (kind = ask_ / ans_). */
export async function addTicketAttachment(svc: TicketsHost, itemId: number, file: File, kind: 'ask' | 'ans'): Promise<boolean> {
  try {
    const buf = await file.arrayBuffer();
    const safeName = (file.name || 'screenshot.png').replace(/[^a-zA-Z0-9._-]+/g, '_');
    const ts = new Date().toISOString().replace(/[:.]/g, '-').replace(/T/, '_').slice(0, 19);
    const finalName = `${kind}_${ts}_${safeName}`;
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('${TICKETS_LIST}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(finalName)}')`;
    const resp = await svc.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' },
      body: buf,
    });
    return resp.ok;
  } catch (err) {
    console.warn('[DEX] addTicketAttachment failed:', err);
    return false;
  }
}

/** Voller Ticket-Select (inkl. neuerer Spalten Category/AnswerWizardMarker). */
const TICKETS_SEL_FULL = 'Id,Title,Questions,Status,AskerEmail,AskerName,AskerRole,AskerLocation,AskerJobTitle,Audience,TicketEventId,TicketEventTitle,AssignedOrganizers,PageContext,AskWizardStep,Category,AnswerText,AnswerArticleIds,AnswerWizardStep,AnswerWizardMarker,AnsweredByEmail,AnsweredByName,AnsweredByLocation,AnsweredByJobTitle,AnsweredAt,ClaimedByEmail,ClaimedByName,ClaimedAt,FollowUps,Created';
/** Legacy-Select ohne die jüngeren Spalten — Fallback, wenn die Live-Liste
 *  sie (noch) nicht hat und auch nicht angelegt werden können. */
const TICKETS_SEL_LEGACY = 'Id,Title,Questions,Status,AskerEmail,AskerName,AskerRole,AskerLocation,AskerJobTitle,Audience,TicketEventId,TicketEventTitle,AssignedOrganizers,PageContext,AskWizardStep,AnswerText,AnswerArticleIds,AnswerWizardStep,AnsweredByEmail,AnsweredByName,AnsweredByLocation,AnsweredByJobTitle,AnsweredAt,ClaimedByEmail,ClaimedByName,ClaimedAt,FollowUps,Created';

/** v26.61: Ticket-Items robust laden. HTTP 400 heißt fast immer: eine neu
 *  eingeführte Spalte fehlt auf der Live-Liste (der alte Sentinel-Bug in
 *  ensureTicketExtraFields hat Category/AnswerWizardMarker nie angelegt).
 *  Selbstheilung: Spalten nachziehen → EIN Retry mit vollem Select →
 *  andernfalls Legacy-Select ohne die neuen Spalten, damit die Tickets-Seite
 *  NIE leer bleibt (fehlende Werte fallen auf Defaults zurück).
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function _getTicketItems(svc: TicketsHost, suffix: string): Promise<any[] | null> {
  const base = `${svc.siteUrl}/_api/web/lists/getbytitle('${TICKETS_LIST}')/items?$select=`;
  let resp = await svc.context.spHttpClient.get(`${base}${TICKETS_SEL_FULL}${suffix}`, SPHttpClient.configurations.v1);
  if (resp.status === 400) {
    try { await ensureTicketExtraFields(svc); } catch { /* */ }
    resp = await svc.context.spHttpClient.get(`${base}${TICKETS_SEL_FULL}${suffix}`, SPHttpClient.configurations.v1);
    if (resp.status === 400) {
      console.warn('[DEX] Ticket-Select weiterhin 400 — Fallback auf Legacy-Spalten (ohne Category/AnswerWizardMarker).');
      resp = await svc.context.spHttpClient.get(`${base}${TICKETS_SEL_LEGACY}${suffix}`, SPHttpClient.configurations.v1);
    }
  }
  if (!resp.ok) return null;
  const data = await resp.json();
  return data.value || data.d?.results || [];
}

/** Alle Tickets laden (inkl. Anhänge per $expand). Neueste zuerst. */
export async function getTickets(svc: TicketsHost): Promise<DexTicket[]> {
  try {
    const items = await _getTicketItems(svc, '&$expand=AttachmentFiles&$orderby=Created desc&$top=500');
    if (!items) return [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (items as any[]).map((it) => _mapTicket(svc, it));
  } catch (err) {
    console.warn('[DEX] getTickets failed:', err);
    return [];
  }
}

/** Eigene Tickets eines Fragestellers (für die „Deine Fragen"-Ansicht im
 *  Ask-Modal — der Fragesteller sieht Status + Antwort in der App). */
export async function getMyTickets(svc: TicketsHost, email: string): Promise<DexTicket[]> {
  try {
    const safe = (email || '').replace(/'/g, "''");
    const items = await _getTicketItems(svc, `&$expand=AttachmentFiles&$filter=AskerEmail eq '${safe}'&$orderby=Created desc&$top=100`);
    if (!items) return [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (items as any[]).map((it) => _mapTicket(svc, it));
  } catch (err) {
    console.warn('[DEX] getMyTickets failed:', err);
    return [];
  }
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function _mapTicket(svc: TicketsHost, it: any): DexTicket {
  const parseArr = (s: unknown): string[] => {
    try { const a = JSON.parse((s as string) || '[]'); return Array.isArray(a) ? a.map((x) => String(x)) : []; } catch { return []; }
  };
  const attsRaw = (it.AttachmentFiles && (it.AttachmentFiles.results || it.AttachmentFiles)) || [];
  const attachments: TicketAttachment[] = (Array.isArray(attsRaw) ? attsRaw : []).map((a: { FileName?: string; ServerRelativeUrl?: string }) => {
    const fn = a.FileName || '';
    const kind: TicketAttachment['kind'] = fn.indexOf('ask_') === 0 ? 'ask' : fn.indexOf('ans_') === 0 ? 'ans' : 'other';
    return { fileName: fn, url: a.ServerRelativeUrl || '', kind };
  });
  const stepRaw = it.AnswerWizardStep;
  const askStepRaw = it.AskWizardStep;
  const parseFollowUps = (s: unknown): TicketFollowUp[] => {
    try {
      const a = JSON.parse((s as string) || '[]');
      if (!Array.isArray(a)) return [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return a.map((x: any) => ({
        byEmail: String(x.byEmail || ''),
        byName: String(x.byName || ''),
        byRole: (x.byRole === 'answerer' ? 'answerer' : 'asker') as TicketFollowUp['byRole'],
        text: String(x.text || ''),
        at: String(x.at || ''),
      }));
    } catch { return []; }
  };
  return {
    id: it.Id,
    title: it.Title || '',
    questions: parseArr(it.Questions),
    status: (it.Status || 'Open') as DexTicket['status'],
    askerEmail: it.AskerEmail || '',
    askerName: it.AskerName || '',
    askerRole: (it.AskerRole || 'User') as DexTicket['askerRole'],
    askerLocation: it.AskerLocation || '',
    askerJobTitle: it.AskerJobTitle || '',
    audience: (it.Audience || 'PowerUser') as DexTicket['audience'],
    eventId: it.TicketEventId || '',
    eventTitle: it.TicketEventTitle || '',
    assignedOrganizers: parseArr(it.AssignedOrganizers),
    pageContext: it.PageContext || '',
    // v26.60: Bug-Report vs. inhaltliche Frage (Bestand ohne Category = Frage).
    category: (it.Category === 'Bug' ? 'bug' : 'question') as DexTicket['category'],
    askWizardStep: (askStepRaw === 0 || (askStepRaw != null && askStepRaw !== '')) ? Number(askStepRaw) : null,
    answerText: it.AnswerText || '',
    answerArticleIds: parseArr(it.AnswerArticleIds),
    answerWizardStep: (stepRaw === 0 || (stepRaw != null && stepRaw !== '')) ? Number(stepRaw) : null,
    // v26.52: Markierungsbox (JSON {x,y,w,h} in Prozent) — defensiv parsen.
    answerWizardMarker: (() => {
      try {
        const m = JSON.parse(it.AnswerWizardMarker || 'null');
        if (m && typeof m.x === 'number' && typeof m.y === 'number' && typeof m.w === 'number' && typeof m.h === 'number') return m;
      } catch { /* */ }
      return null;
    })(),
    answeredByEmail: it.AnsweredByEmail || '',
    answeredByName: it.AnsweredByName || '',
    answeredByLocation: it.AnsweredByLocation || '',
    answeredByJobTitle: it.AnsweredByJobTitle || '',
    answeredAt: it.AnsweredAt || '',
    claimedByEmail: it.ClaimedByEmail || '',
    claimedByName: it.ClaimedByName || '',
    claimedAt: it.ClaimedAt || '',
    created: it.Created || '',
    attachments,
    followUps: parseFollowUps(it.FollowUps),
  };
}

/** Ticket „in Bearbeitung" nehmen (Claim). */
/**
 * v26.32: Ticket übernehmen mit OPTIMISTIC CONCURRENCY, damit nicht zwei
 * Power-User gleichzeitig dasselbe Ticket übernehmen. Ablauf:
 *   1. Aktuellen Stand (Status/Claim) + ETag lesen.
 *   2. `onlyIfOpen` (Standard beim Klick auf ein OFFENES Ticket): ist es
 *      inzwischen von jemand anderem übernommen/geschlossen → conflict.
 *   3. Bedingter MERGE mit IF-MATCH=<ETag>. Hat zwischen Schritt 1 und 3
 *      jemand geschrieben → HTTP 412 → wir lesen den aktuellen Claimer nach
 *      und melden conflict (kein stilles Überschreiben mehr).
 * Rückgabe: { ok, conflict?, claimedByName?, status? }.
 */
export async function claimTicket(svc: TicketsHost,
  itemId: number, email: string, name: string, opts?: { onlyIfOpen?: boolean }
): Promise<{ ok: boolean; conflict?: boolean; claimedByName?: string; status?: string }> {
  const itemUrl = `${svc.siteUrl}/_api/web/lists/getbytitle('${TICKETS_LIST}')/items(${itemId})`;
  const readCurrent = async (): Promise<{ status: string; claimEmail: string; claimName: string; etag: string } | null> => {
    try {
      const getResp = await svc.context.spHttpClient.get(
        `${itemUrl}?$select=Status,ClaimedByEmail,ClaimedByName`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );
      if (!getResp.ok) return null;
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      const data = await getResp.json();
      return {
        status: String(data.Status || 'Open'),
        claimEmail: String(data.ClaimedByEmail || ''),
        claimName: String(data.ClaimedByName || ''),
        etag,
      };
    } catch { return null; }
  };
  try {
    const cur = await readCurrent();
    if (!cur || !cur.etag) return { ok: false };
    const meLc = (email || '').toLowerCase();
    const takenByOther = cur.claimEmail && cur.claimEmail.toLowerCase() !== meLc;
    // Beim Klick auf ein OFFENES Ticket: schon vergeben/geschlossen? → Konflikt.
    if (opts?.onlyIfOpen && (cur.status !== 'Open') && takenByOther) {
      return { ok: false, conflict: true, claimedByName: cur.claimName || cur.claimEmail, status: cur.status };
    }
    // Bereits beantwortet/geschlossen → nie überschreiben.
    if (cur.status === 'Closed') {
      return { ok: false, conflict: true, claimedByName: cur.claimName || cur.claimEmail, status: cur.status };
    }
    const resp = await svc._mergeIfMatch(
      itemUrl,
      { 'Status': 'InProgress', 'ClaimedByEmail': email, 'ClaimedByName': name, 'ClaimedAt': new Date().toISOString() },
      cur.etag
    );
    if (resp.status === 412) {
      // Race verloren (jemand war zwischen Lesen und Schreiben schneller).
      const after = await readCurrent();
      return { ok: false, conflict: true, claimedByName: after ? (after.claimName || after.claimEmail) : undefined, status: after?.status };
    }
    return { ok: resp.ok };
  } catch (err) {
    console.warn('[DEX] claimTicket failed:', err);
    return { ok: false };
  }
}

/** Ticket wieder freigeben (zurück auf „Open"), damit ein anderer übernimmt. */
export async function releaseTicket(svc: TicketsHost, itemId: number): Promise<boolean> {
  try {
    const resp = await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${TICKETS_LIST}')/items(${itemId})`,
      { 'Status': 'Open', 'ClaimedByEmail': '', 'ClaimedByName': '', 'ClaimedAt': null }
    );
    return resp.ok;
  } catch (err) {
    console.warn('[DEX] releaseTicket failed:', err);
    return false;
  }
}

/** Ticket beantworten + schließen. */
export async function answerTicket(svc: TicketsHost, itemId: number, a: {
  answerText: string; articleIds: string[]; wizardStep: number | null;
  wizardMarker?: { x: number; y: number; w: number; h: number } | null;
  answeredByEmail: string; answeredByName: string;
  answeredByLocation?: string; answeredByJobTitle?: string;
}): Promise<boolean> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: Record<string, any> = {
      'Status': 'Closed',
      'AnswerText': a.answerText || '',
      'AnswerArticleIds': JSON.stringify(a.articleIds || []),
      'AnswerWizardStep': (a.wizardStep == null) ? null : a.wizardStep,
      'AnswerWizardMarker': a.wizardMarker ? JSON.stringify(a.wizardMarker) : '',
      'AnsweredByEmail': a.answeredByEmail,
      'AnsweredByName': a.answeredByName,
      'AnsweredByLocation': a.answeredByLocation || '',
      'AnsweredByJobTitle': a.answeredByJobTitle || '',
      'AnsweredAt': new Date().toISOString(),
    };
    const resp = await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${TICKETS_LIST}')/items(${itemId})`,
      body
    );
    return resp.ok;
  } catch (err) {
    console.warn('[DEX] answerTicket failed:', err);
    return false;
  }
}

/** v26.8: Rückfragen-Verlauf schreiben + Status setzen. Vom Fragesteller
 *  (Status zurück auf InProgress, der/dem Beantwortenden zugewiesen) ODER von
 *  der/dem Beantwortenden als Folge-Antwort (Status Closed). */
export async function setTicketFollowUps(svc: TicketsHost, itemId: number, followUps: TicketFollowUp[], extra: {
  status?: string; claimedByEmail?: string; claimedByName?: string; claimedAt?: string | null;
}): Promise<boolean> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: Record<string, any> = { 'FollowUps': JSON.stringify(followUps || []) };
    if (extra.status !== undefined) body['Status'] = extra.status;
    if (extra.claimedByEmail !== undefined) body['ClaimedByEmail'] = extra.claimedByEmail;
    if (extra.claimedByName !== undefined) body['ClaimedByName'] = extra.claimedByName;
    if (extra.claimedAt !== undefined) body['ClaimedAt'] = extra.claimedAt;
    const resp = await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${TICKETS_LIST}')/items(${itemId})`,
      body
    );
    return resp.ok;
  } catch (err) {
    console.warn('[DEX] setTicketFollowUps failed:', err);
    return false;
  }
}

/** v26.8: Ticket schließen, ohne eine Antwort zu senden („keine Antwort nötig",
 *  z.B. wenn die Rückfrage nur ein Dankeschön war). */
export async function closeTicketNoAnswer(svc: TicketsHost, itemId: number): Promise<boolean> {
  try {
    const resp = await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${TICKETS_LIST}')/items(${itemId})`,
      { 'Status': 'Closed' }
    );
    return resp.ok;
  } catch (err) {
    console.warn('[DEX] closeTicketNoAnswer failed:', err);
    return false;
  }
}

// ==================== DEX_OrganizerRequests (v23.37) ====================
// Anträge „Organizer werden". Jeder authentifizierte User darf einen Antrag
// anlegen (setQueueListPermissions); Admins sehen offene Anträge in der App
// und bestätigen sie (→ Organizer-Rolle wird vergeben).
