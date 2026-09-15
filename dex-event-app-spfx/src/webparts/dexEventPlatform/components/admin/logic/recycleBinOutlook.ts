/**
 * v31.49: „Gelöschte Termine im Papierkorb suchen".
 *
 * Anlass (15.09.2026): Der Recreate-Pfad des Wizards (bis v31.47, s.
 * persistSubEvents) legte für ein Sub-Event eine NEUE DEX_Events-Zeile an und
 * löschte die alte — samt ihrem `CalendarLink`, dem einzigen Verweis auf den
 * Outlook-Termin, auf dem hunderte Teilnehmer stehen. Die alte Zeile liegt
 * (wenn überhaupt) nur noch im Papierkorb der Site. Diese Aktion sucht sie
 * dort und holt sie zurück.
 *
 * Warum „zurückholen" und nicht „Link kopieren": Der Papierkorb liefert über
 * REST nur Metadaten (Titel, Löschdatum, wer), nicht die Feldwerte — den
 * CalendarLink sieht man erst nach dem Restore. Und die alte Zeile ist die
 * bessere Zeile: gleiche Item-Id und EventNumber wie in DEX_Participants,
 * in Queue-Zeilen und in alten Mails. Deshalb: alte Zeile wiederherstellen,
 * neue Zeile (nur die Zeile, nie die geteilte Subsite) entfernen, den
 * überzähligen neuen Termin über die Queue absagen.
 *
 * Reihenfolge nach CLAUDE.md „Der unumkehrbare Schritt kommt zuletzt": erst
 * wiederherstellen und prüfen (ParentEventId, Subsite, CalendarLink), dann
 * erst die neue Zeile löschen. Passt etwas nicht, wird der Restore
 * zurückgenommen und nichts verändert. Unbekannt sperrt: Ist der Papierkorb
 * nicht lesbar, gibt es kein „nichts gefunden", sondern „nicht lesbar".
 */
import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService, SPEvent } from '../../../services/EventService';
import type { DeloitteEvent } from '../../../types';

export interface PapierkorbTreffer {
  /** GUID des Papierkorb-Eintrags. */
  binId: string;
  /** Erste Stufe (Web) oder zweite Stufe (Site Collection). */
  stufe: 'web' | 'site';
  title: string;
  /** Item-Id der gelöschten Zeile (aus LeafName „123_.000"), null wenn nicht lesbar. */
  oldItemId: number | null;
  deletedDate: string;
  deletedBy: string;
  /** Aktuelles Sub-Event mit demselben Titel — Ziel des Austauschs. */
  ziel: { id: string; title: string } | null;
  /** Warum kein Ziel: kein oder mehrere Sub-Events mit diesem Titel. */
  zielGrund: 'ok' | 'kein-subevent' | 'mehrdeutig' | 'keine-id';
}

export interface PapierkorbSuche {
  status: 'ok' | 'nicht-lesbar';
  /** false = zweite Stufe (Site Collection) war nicht lesbar, Treffer nur aus der ersten. */
  zweiteStufeGelesen: boolean;
  treffer: PapierkorbTreffer[];
}

interface BinRow {
  Id: string;
  Title?: string;
  LeafName?: string;
  DirName?: string;
  DeletedDate?: string;
  DeletedByName?: string;
  ItemType?: number;
}

const norm = (s: string | undefined | null): string => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();

/** ItemType 3 = ListItem (SP.RecycleBinItemType). */
const RECYCLE_LIST_ITEM = 3;

async function leseBin(svc: EventService, stufe: 'web' | 'site'): Promise<BinRow[] | null> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/${stufe}/RecycleBin?$select=Id,Title,LeafName,DirName,DeletedDate,DeletedByName,ItemType&$top=2000`,
      SPHttpClient.configurations.v1,
    );
    if (!resp.ok) return null;
    const data = await resp.json();
    return (data.value || data.d?.results || []) as BinRow[];
  } catch {
    return null;
  }
}

/**
 * Sucht im Papierkorb nach gelöschten DEX_Events-Zeilen, deren Titel zu
 * einem Sub-Event dieses Events passt.
 */
export async function suchePapierkorb(svc: EventService, event: DeloitteEvent, kinder: DeloitteEvent[]): Promise<PapierkorbSuche> {
  const web = await leseBin(svc, 'web');
  if (!web) return { status: 'nicht-lesbar', zweiteStufeGelesen: false, treffer: [] };
  const site = await leseBin(svc, 'site');
  const rows: Array<BinRow & { stufe: 'web' | 'site' }> = [];
  const seen = new Set<string>();
  for (const r of web) { if (!seen.has(r.Id)) { seen.add(r.Id); rows.push({ ...r, stufe: 'web' }); } }
  for (const r of (site || [])) { if (!seen.has(r.Id)) { seen.add(r.Id); rows.push({ ...r, stufe: 'site' }); } }
  // Titel-Index der aktuellen Sub-Events. Die Klammer selbst kann nicht
  // ersetzt worden sein (ihre Id steht in ParentEventId aller Kinder) —
  // deshalb zählen nur Kinder.
  const byTitle: Record<string, DeloitteEvent[]> = {};
  for (const k of kinder) {
    const key = norm(k.title);
    if (!key) continue;
    (byTitle[key] = byTitle[key] || []).push(k);
  }
  const treffer: PapierkorbTreffer[] = [];
  for (const r of rows) {
    if (r.ItemType !== RECYCLE_LIST_ITEM) continue;
    if (norm(r.DirName).indexOf('lists/dex_events') < 0) continue;
    const key = norm(r.Title);
    const kandidaten = byTitle[key];
    if (!kandidaten) continue;
    const m = /^(\d+)_\./.exec(r.LeafName || '');
    const oldItemId = m ? parseInt(m[1], 10) : null;
    treffer.push({
      binId: r.Id,
      stufe: r.stufe,
      title: r.Title || '',
      oldItemId,
      deletedDate: r.DeletedDate || '',
      deletedBy: r.DeletedByName || '',
      ziel: kandidaten.length === 1 ? { id: kandidaten[0].id, title: kandidaten[0].title } : null,
      zielGrund: oldItemId === null ? 'keine-id' : kandidaten.length === 1 ? 'ok' : 'mehrdeutig',
    });
  }
  treffer.sort((a, b) => (b.deletedDate || '').localeCompare(a.deletedDate || ''));
  return { status: 'ok', zweiteStufeGelesen: !!site, treffer };
}

export type RueckholStatus =
  | 'fertig'
  | 'restore-fehlgeschlagen'
  | 'alte-zeile-nicht-lesbar'
  | 'falsches-event'       // ParentEventId ≠ dieses Event → Restore zurückgenommen
  | 'andere-subsite'       // Subsite der alten Zeile ≠ Subsite des Ziels → Restore zurückgenommen
  | 'kein-termin'          // alte Zeile ohne CalendarLink → Restore zurückgenommen
  | 'ziel-nicht-lesbar'    // aktuelle Zeile nicht lesbar → Restore zurückgenommen
  | 'neue-zeile-bleibt';   // alte Zeile steht, neue konnte nicht gelöscht werden → doppelt

export interface RueckholErgebnis {
  status: RueckholStatus;
  oldItemId: number;
  newItemId: string;
  calendarLink: string;
  /** Was mit dem Termin der neuen Zeile passiert ist: eingereiht = Absage in
   *  der Queue; behalten = auf Wunsch nicht abgesagt (er bleibt in den
   *  Kalendern derer, die darauf stehen, hängt aber an keiner DEX-Zeile mehr);
   *  keiner = die neue Zeile hatte keinen eigenen Termin; fehlgeschlagen =
   *  Absage konnte nicht eingereiht werden (von Hand absagen). */
  neuerTermin: 'eingereiht' | 'behalten' | 'keiner' | 'fehlgeschlagen';
  /** true = Restore wurde zurückgenommen, Zustand wie vorher. */
  zurueckgenommen: boolean;
}

/**
 * Alte Zeile wiederherstellen, prüfen, neue Zeile ersetzen.
 * Die Prüfungen laufen VOR dem einzigen unumkehrbaren Schritt (Löschen der
 * neuen Zeile); jeder Fehlschlag davor nimmt den Restore zurück.
 * `neuenTerminAbsagen` = Nutzer-Wahl (15.09.2026: „die Action soll die
 * Auswahl haben, ob man den neuen/doppelten Termin löschen möchte oder nicht").
 */
export async function holeZeileZurueck(svc: EventService, event: DeloitteEvent, t: PapierkorbTreffer, neuenTerminAbsagen: boolean): Promise<RueckholErgebnis> {
  const oldItemId = t.oldItemId as number;
  const newItemId = t.ziel ? t.ziel.id : '';
  const erg = (status: RueckholStatus, extra?: Partial<RueckholErgebnis>): RueckholErgebnis => ({
    status, oldItemId, newItemId, calendarLink: '', neuerTermin: 'keiner', zurueckgenommen: false, ...extra,
  });
  if (!t.ziel || oldItemId === null) return erg('ziel-nicht-lesbar');

  // 1. Wiederherstellen.
  try {
    const r = await svc._post(`${svc.siteUrl}/_api/${t.stufe}/RecycleBin('${t.binId}')/Restore()`, {});
    if (!r.ok) return erg('restore-fehlgeschlagen');
  } catch {
    return erg('restore-fehlgeschlagen');
  }
  // Rücknahme = die Zeile wieder löschen (nur die Zeile). Damit ist der
  // Zustand wie vor dem Klick.
  const zurueck = async (status: RueckholStatus): Promise<RueckholErgebnis> => {
    const ok = await svc.deleteEventItemOnly(oldItemId);
    return erg(status, { zurueckgenommen: ok });
  };

  // 2. Alte Zeile lesen und prüfen.
  const alt: SPEvent | null = await svc.getEvent(oldItemId);
  if (!alt) return erg('alte-zeile-nicht-lesbar');
  if (String(alt.ParentEventId || '') !== String(event.id)) return zurueck('falsches-event');
  const link = (alt.CalendarLink || '').trim();
  if (!link) return zurueck('kein-termin');
  const neu: SPEvent | null = await svc.getEvent(Number(newItemId));
  if (!neu) return zurueck('ziel-nicht-lesbar');
  if (norm(neu.SubsiteUrl) !== norm(alt.SubsiteUrl)) return zurueck('andere-subsite');

  // 3. Überzähligen neuen Termin absagen (Queue; trifft nur, wer auf dem
  //    neuen Termin steht — beim Recreate sind das die Organizer).
  let neuerTermin: RueckholErgebnis['neuerTermin'] = 'keiner';
  const neuerLink = (neu.CalendarLink || '').trim();
  if (neuerLink && neuerLink !== link) {
    if (neuenTerminAbsagen) {
      let ok = false;
      try { ok = await svc.queueOutlookDeleteEvent(String(newItemId), neu.Title || '', neuerLink); } catch { ok = false; }
      neuerTermin = ok ? 'eingereiht' : 'fehlgeschlagen';
    } else {
      neuerTermin = 'behalten';
    }
  }

  // 4. Der unumkehrbare Schritt: neue Zeile entfernen (NUR die Zeile — die
  //    Subsite trägt die Anmeldungen und gehört jetzt wieder der alten).
  const weg = await svc.deleteEventItemOnly(newItemId);
  svc.writeChangeLog({
    action: weg ? 'EventItemRestoredFromRecycleBin' : 'EventItemRestoredDuplicate',
    targetType: 'Event',
    targetId: String(oldItemId),
    targetName: alt.Title || '',
    eventId: String(event.id),
    eventTitle: event.title || '',
    details: { restoredId: oldItemId, replacedId: newItemId, calendarLink: link, newAppointment: neuerTermin },
  }).catch(() => { /* Log ist best-effort */ });
  return erg(weg ? 'fertig' : 'neue-zeile-bleibt', { calendarLink: link, neuerTermin });
}

/*
 * v31.50: Zweiter Weg — wenn die alte Zeile NICHT im Papierkorb liegt (der
 * alte Pfad löschte per REST-DELETE, das landet nicht im Papierkorb). Dann
 * existiert nur noch der Termin selbst, im Kalender der Shared Mailbox. Die
 * App darf ihn lesen: `Calendars.Read.Shared` ist seit v5.x genehmigt
 * (getDeclinedAttendees liest denselben Kalender). Also: Termine rund um den
 * Start des Sub-Events aus dem Kalender holen, die iCalUIds zeigen, und den
 * gewählten Termin in die aktuelle Zeile schreiben — der Weg, den der Nutzer
 * am 15.09.2026 „manuell austauschen" nannte, nur ohne Abtippen.
 */
export const NO_REPLY_MAILBOX = 'no_reply.events@deloitte.de';

export interface KalenderTermin {
  graphId: string;
  iCalUId: string;
  subject: string;
  /** ISO-UTC des Starts. */
  start: string;
  created: string;
  attendees: number;
  /** iCalUId ist der CalendarLink der aktuellen Zeile. */
  verknuepft: boolean;
  /** Betreff passt zu Titel oder Outlook-Betreff des Sub-Events. */
  passt: boolean;
}

export type KalenderSuche =
  | { status: 'ok'; termine: KalenderTermin[]; fenster: string }
  | { status: 'kein-graph' | 'kein-zugriff' | 'fehler'; message?: string };

export async function sucheKalender(svc: EventService, kind: DeloitteEvent): Promise<KalenderSuche> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ctx = svc.context as any;
  if (!ctx.msGraphClientFactory) return { status: 'kein-graph' };
  const anker = kind.outlookStart || kind.startDate || '';
  const d = anker ? new Date(anker) : null;
  const linkAktuell = (kind.calendarLink || '').trim();
  const titel = norm(kind.title);
  const betreff = norm(kind.outlookSubject);
  const select = 'id,subject,iCalUId,start,createdDateTime,attendees';
  try {
    const client = await ctx.msGraphClientFactory.getClient('3');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let rows: any[] = [];
    let fenster = '';
    if (d && !isNaN(d.getTime())) {
      // ±3 Tage um den Start: calendarView löst auch Serien auf und braucht
      // keinen Betreff — der kann sich seit dem Anlegen geändert haben.
      const von = new Date(d.getTime() - 3 * 864e5).toISOString();
      const bis = new Date(d.getTime() + 3 * 864e5).toISOString();
      fenster = `${von.slice(0, 10)} – ${bis.slice(0, 10)}`;
      const resp = await client.api(`/users/${NO_REPLY_MAILBOX}/calendarView`)
        .query({ startDateTime: von, endDateTime: bis })
        .select(select).top(100).get();
      rows = resp?.value || [];
    } else {
      const s = (kind.outlookSubject || kind.title || '').slice(0, 40).replace(/'/g, "''");
      const resp = await client.api(`/users/${NO_REPLY_MAILBOX}/events`)
        .filter(`startswith(subject,'${s}')`)
        .select(select).top(100).get();
      rows = resp?.value || [];
    }
    const termine: KalenderTermin[] = rows.map(r => {
      const subj = norm(r?.subject);
      const uid = String(r?.iCalUId || '');
      const startRaw = String(r?.start?.dateTime || '');
      return {
        graphId: String(r?.id || ''),
        iCalUId: uid,
        subject: String(r?.subject || ''),
        start: startRaw ? (startRaw.replace(/\.\d+$/, '') + (/[zZ]$/.test(startRaw) ? '' : 'Z')) : '',
        created: String(r?.createdDateTime || ''),
        attendees: Array.isArray(r?.attendees) ? r.attendees.length : 0,
        verknuepft: !!uid && uid === linkAktuell,
        passt: !!subj && ((!!titel && (subj === titel || subj.indexOf(titel) >= 0)) || (!!betreff && (subj === betreff || subj.indexOf(betreff) >= 0))),
      };
    }).filter(t => !!t.iCalUId);
    termine.sort((a, b) => (Number(b.passt) - Number(a.passt)) || (b.attendees - a.attendees));
    return { status: 'ok', termine, fenster };
  } catch (err) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const status = (err as any)?.statusCode || (err as any)?.status;
    if (status === 401 || status === 403) return { status: 'kein-zugriff' };
    return { status: 'fehler', message: err instanceof Error ? err.message : String(err) };
  }
}

export interface VerknuepfErgebnis {
  status: 'fertig' | 'schon-verknuepft' | 'zeile-nicht-lesbar' | 'schreiben-fehlgeschlagen';
  /** Was mit dem bisher verknüpften Termin passiert ist (s. RueckholErgebnis.neuerTermin). */
  bisheriger: 'eingereiht' | 'behalten' | 'keiner' | 'fehlgeschlagen';
}

/**
 * Schreibt die iCalUId des gewählten Termins als CalendarLink in die aktuelle
 * Zeile. Erst der umkehrbare Schritt (MERGE), dann — auf Wunsch — die Absage
 * des bisher verknüpften Termins über die Queue.
 */
export async function verknuepfeTermin(svc: EventService, event: DeloitteEvent, kind: DeloitteEvent, termin: KalenderTermin, bisherigenAbsagen: boolean): Promise<VerknuepfErgebnis> {
  const zeile = await svc.getEvent(Number(kind.id));
  if (!zeile) return { status: 'zeile-nicht-lesbar', bisheriger: 'keiner' };
  const bisher = (zeile.CalendarLink || '').trim();
  if (bisher === termin.iCalUId) return { status: 'schon-verknuepft', bisheriger: 'keiner' };
  try {
    // OutlookEventId mit leeren: Steht dort 'FAILED', hielte die App den
    // Termin weiter für gescheitert.
    const r = await svc._merge(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${Number(kind.id)})`, { CalendarLink: termin.iCalUId, OutlookEventId: '' });
    if (!r.ok) return { status: 'schreiben-fehlgeschlagen', bisheriger: 'keiner' };
  } catch {
    return { status: 'schreiben-fehlgeschlagen', bisheriger: 'keiner' };
  }
  let bisheriger: VerknuepfErgebnis['bisheriger'] = 'keiner';
  if (bisher) {
    if (bisherigenAbsagen) {
      let ok = false;
      try { ok = await svc.queueOutlookDeleteEvent(String(kind.id), zeile.Title || '', bisher); } catch { ok = false; }
      bisheriger = ok ? 'eingereiht' : 'fehlgeschlagen';
    } else {
      bisheriger = 'behalten';
    }
  }
  svc.writeChangeLog({
    action: 'EventCalendarLinkRelinked',
    targetType: 'Event',
    targetId: String(kind.id),
    targetName: zeile.Title || '',
    eventId: String(event.id),
    eventTitle: event.title || '',
    details: { calendarLink: termin.iCalUId, previousCalendarLink: bisher, subject: termin.subject, attendees: termin.attendees, previousAppointment: bisheriger },
  }).catch(() => { /* Log ist best-effort */ });
  return { status: 'fertig', bisheriger };
}
