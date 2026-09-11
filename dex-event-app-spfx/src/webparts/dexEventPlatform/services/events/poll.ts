/**
 * Umfrage nach dem Event — „Warum warst du nicht dabei?"
 *
 * Aus der Nutzer-Frage vom 10.09.2026: eine Abstimmung, bei der die Leute
 * „mit einer Checkbox in einer Mail direkt antworten" können.
 *
 * ## Warum es keine echte Checkbox in der Mail ist
 *
 * E-Mail-Clients führen keine Formulare aus; Outlook entfernt `<form>` beim
 * Rendern, und eine Checkbox ohne Formular hätte niemanden, an den sie etwas
 * schicken könnte. Was hier steht, ist der Weg, den jedes Umfrage-Werkzeug in
 * Wirklichkeit geht: Jede Antwortmöglichkeit ist ein LINK, der wie ein
 * Kästchen aussieht (`☐ Zu früher Startblock`). Ein Klick öffnet DEX, die
 * Antwort ist sofort gespeichert, und auf der Seite kann man weitere Gründe
 * anhaken oder etwas dazuschreiben.
 *
 * Der erste Klick zählt bereits — wer danach abbricht, hat trotzdem
 * geantwortet. Das ist der ganze Zweck: Rückläufe sterben an jedem
 * zusätzlichen Schritt.
 *
 * ## Wie weit „anonym" in SharePoint reicht — und wo es aufhört
 *
 * Nutzer-Entscheidung: anonym, aber mit einem Merker, wer schon geantwortet
 * hat (sonst kann man nicht erinnern).
 *
 * Der erste Entwurf dieser Datei löste das über Einmal-Token: Die Adresse
 * sollte beim ersten Klick aus der Einladungszeile gelöscht werden, sodass
 * Token und Person nicht mehr zusammenführbar sind. Das trägt nicht, und
 * zwar aus einem Grund, den kein Umbau wegkonstruiert: **SharePoint schreibt
 * an JEDE Zeile `Author` und `Editor`.** Wer eine Zeile anlegt oder ändert,
 * steht in ihr — auch wenn in keiner Spalte ein Name steht. Eine vom
 * Teilnehmer geschriebene Antwortzeile ist damit immer zuordenbar, und ein
 * Token hätte nur so getan, als wäre sie es nicht.
 *
 * Was hier deshalb gebaut ist, und was die Oberfläche und die Mail auch so
 * sagen:
 *
 *  - Die Antwort steht in `DEX_PollAnswers` mit Item-Level-Security (2/2):
 *    Teilnehmende sehen und ändern **nur die eigene** Zeile, nie die der
 *    anderen.
 *  - Ist die Umfrage als anonym angelegt, zeigt und exportiert DEX **an
 *    keiner Stelle**, wer was geantwortet hat — die Auswertung kennt nur
 *    Summen und die Freitexte ohne Absender.
 *  - Wer Vollzugriff auf die Website hat, kann in SharePoint selbst trotzdem
 *    sehen, wer eine Zeile geschrieben hat. Echte Anonymität bräuchte einen
 *    Power-Automate-Flow, der die Antwort unter einem Dienstkonto schreibt.
 *    Das steht so in der Auswertung, statt eine Zusage zu geben, die die
 *    Plattform nicht hält.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';


export const POLL_LIST = {
  config: 'DEX_Polls',
  answers: 'DEX_PollAnswers',
};

/**
 * Der Listen-Entitätstyp für `__metadata`. SharePoint kodiert den Unterstrich
 * im Listennamen als `_x005f_` — `SP.Data.DEX_PollsListItem` gibt es nicht,
 * `SP.Data.DEX_x005f_PollsListItem` schon. Genau so steht es in `queueEmail`.
 */
const itemType = (listName: string): string =>
  `SP.Data.${listName.replace(/_/g, '_x005f_')}ListItem`;

/**
 * v31.27: Den Typ NACHFRAGEN statt ihn zu raten.
 *
 * Der Nutzer-Befund vom 11.09.2026 („Die Umfrage konnte nicht gespeichert
 * werden — es wurde nichts in die Mail eingefügt") hat hier seine Ursache.
 * `itemType` leitet den Entitätstyp aus dem Listen-TITEL ab. SharePoint
 * bildet ihn aber aus dem internen Namen zum Zeitpunkt der Anlage — und der
 * weicht ab, sobald beim Anlegen etwas dazwischenkommt: Eine Liste, die
 * schon existierte und später umbenannt wurde, behält ihren alten internen
 * Namen, und `DEX_Polls` heißt dann intern womöglich `DEX_Polls1` oder
 * `Liste`. Der geratene Typ ist dann falsch, und weil `_post` mit
 * `odata=verbose` sendet, ist ein falscher Typ ein harter HTTP 400.
 *
 * `ListItemEntityTypeFullName` ist die Auskunft der Liste über sich selbst.
 * Sie wird je Liste einmal geholt und gemerkt — der geratene Wert bleibt
 * der Rückfall, wenn die Abfrage nicht durchgeht (dann ist er immer noch
 * besser als gar keiner).
 *
 * Dasselbe Muster steckt in der AI Use Case Platform (`entityType`), wo es
 * aus demselben Grund entstanden ist: HTTP 400 beim ersten Schreibversuch.
 */
const typCache: Record<string, string> = {};
async function echterItemType(svc: EventService, listName: string): Promise<string> {
  if (typCache[listName]) return typCache[listName];
  try {
    const r = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    if (r.ok) {
      const d = await r.json();
      const t = d.ListItemEntityTypeFullName || d.d?.ListItemEntityTypeFullName;
      if (t) { typCache[listName] = t; return t; }
    }
  } catch { /* Rückfall unten */ }
  return itemType(listName);
}

/** Vorschlag für einen Firmenlauf — der Fall aus der Nutzer-Frage. */
export const POLL_VORLAGEN_DE: Array<{ titel: string; frage: string; optionen: string[] }> = [
  {
    titel: 'Absagen verstehen',
    frage: 'Schade, dass du nicht dabei warst — woran lag es?',
    optionen: [
      'Zu früher Startblock',
      'Krank geworden',
      'Wichtiger Kundentermin',
      'Anreise zu aufwendig',
      'Kein Interesse mehr',
      'Anderer Grund',
    ],
  },
  {
    titel: 'Rückmeldung zum Event',
    frage: 'Wie hat dir das Event gefallen?',
    optionen: [
      'Sehr gut',
      'Gut',
      'Ging so',
      'Nicht gut',
    ],
  },
  {
    titel: 'Wunsch für das nächste Mal',
    frage: 'Was sollen wir beim nächsten Mal anders machen?',
    optionen: [
      'Anderer Termin',
      'Anderer Ort',
      'Mehr Zeit zum Netzwerken',
      'Kürzeres Programm',
      'Passt alles so',
    ],
  },
];

export interface PollConfig {
  /** Item-Id in DEX_Polls — 0, wenn es noch keine Umfrage gibt. */
  id: number;
  eventNumber: number;
  eventId: string;
  frage: string;
  optionen: string[];
  /** Mehrere Gründe erlaubt? Vorgabe: ja — „krank UND Kundentermin" ist ein echter Fall. */
  mehrfach: boolean;
  /** DEX zeigt dann nirgends, wer was geantwortet hat. */
  anonym: boolean;
  /** Ausgeschaltet = die Antwortseite nimmt nichts mehr an. */
  aktiv: boolean;
}

export interface PollAnswerRow {
  id: number;
  /** Leer, wenn die Umfrage anonym ist — dann liest die Auswertung ihn gar nicht erst. */
  email: string;
  optionen: string[];
  freitext: string;
  answeredAt: string;
}

export interface PollStats {
  /** Wie viele geantwortet haben. */
  geantwortet: number;
  /** Je Antwortmöglichkeit die Anzahl. Bei Mehrfachnennung ist die Summe größer. */
  zaehler: Record<string, number>;
  /** Die Freitexte — ohne Absender, auch bei nicht-anonymen Umfragen in dieser Reihenfolge gemischt. */
  freitexte: string[];
  /** Wer schon geantwortet hat. Leer bei anonymer Umfrage. */
  emails: string[];
  /**
   * `true` = die Liste war nicht lesbar. Die Zahlen sind dann KEINE Aussage
   * über die Umfrage, sondern über gar nichts (v30.66/30.67).
   */
  nichtLesbar: boolean;
}

const esc = (v: string): string => (v || '').replace(/'/g, "''");

const zeilen = (v: unknown): string[] =>
  String(v || '').split('\n').map(s => s.trim()).filter(Boolean);

// ===================================================================
// Listen
// ===================================================================

interface FeldDef { titel: string; typ: number; note?: boolean }

async function feldAnlegen(svc: EventService, liste: string, f: FeldDef): Promise<void> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const payload: Record<string, any> = {
    '__metadata': { 'type': f.note ? 'SP.FieldMultiLineText' : 'SP.Field' },
    'Title': f.titel,
    'FieldTypeKind': f.typ,
    'Required': false,
  };
  if (f.note) {
    // RichText=false ist hier nicht Geschmack: Mit RichText wickelt SharePoint
    // den Wert in <div class="ExternalClass…"> — und die Optionen werden
    // zeilenweise zurückgelesen.
    payload['RichText'] = false;
    payload['NumberOfLines'] = 6;
  }
  await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${liste}')/fields`, payload);
}

/** Ein Feld nachziehen, falls es fehlt (Bestandslisten aus einer älteren Version). */
async function feldSicherstellen(svc: EventService, liste: string, f: FeldDef): Promise<void> {
  try {
    const probe = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${liste}')/fields/getbytitle('${f.titel}')?$select=Id`,
      SPHttpClient.configurations.v1,
    );
    if (probe.ok) return;
    await feldAnlegen(svc, liste, f);
  } catch { /* keine Manage-Lists-Rechte → ein Admin legt es an */ }
}

const CONFIG_FELDER: FeldDef[] = [
  { titel: 'EventNumber', typ: 9 },
  { titel: 'EventId', typ: 2 },
  { titel: 'Frage', typ: 3, note: true },
  { titel: 'Optionen', typ: 3, note: true },
  { titel: 'Mehrfach', typ: 8 },
  { titel: 'Anonym', typ: 8 },
  { titel: 'Aktiv', typ: 8 },
];

const ANSWER_FELDER: FeldDef[] = [
  { titel: 'EventNumber', typ: 9 },
  { titel: 'ParticipantEmail', typ: 2 },
  { titel: 'Optionen', typ: 3, note: true },
  { titel: 'Freitext', typ: 3, note: true },
  { titel: 'AnsweredAt', typ: 4 },
];

async function listeAnlegen(
  svc: EventService,
  name: string,
  beschreibung: string,
  felder: FeldDef[],
): Promise<boolean> {
  if (await svc.listExists(name)) {
    for (const f of felder) {
      // eslint-disable-next-line no-await-in-loop
      await feldSicherstellen(svc, name, f);
    }
    return false;
  }
  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': name,
    'Description': beschreibung,
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });
  for (const f of felder) {
    // eslint-disable-next-line no-await-in-loop
    await feldAnlegen(svc, name, f);
  }
  return true;
}

/**
 * Die beiden Listen sicherstellen.
 *
 * Wird vor jedem Schreib- und Lesezugriff aus dem Organizer Center gerufen;
 * die Antwortseite ruft es NICHT — Teilnehmende haben keine Manage-Lists-
 * Rechte, dort wäre es nur eine Runde 403 in der Konsole.
 */
export async function ensurePollLists(svc: EventService): Promise<void> {
  await listeAnlegen(
    svc, POLL_LIST.config,
    'Umfragen nach dem Event: Frage und Antwortmöglichkeiten je Event.',
    CONFIG_FELDER,
  );
  const neu = await listeAnlegen(
    svc, POLL_LIST.answers,
    'Antworten der Event-Umfragen. Zeilenweise gesichert — wer antwortet, sieht nur die eigene Zeile.',
    ANSWER_FELDER,
  );

  await rechteSetzen(svc, POLL_LIST.config);
  await rechteSetzen(svc, POLL_LIST.answers);

  try {
    // 2/2 = jede Person sieht und ändert nur die eigene Antwort. Ohne das
    // könnte jede teilnehmende Person die Antworten aller anderen lesen —
    // Contribute allein reicht dafür aus (dieselbe Falle wie beim Check-in-
    // Team, v30.87, nur andersherum).
    await svc._setListSecurity(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${POLL_LIST.answers}')`,
      { ReadSecurity: 2, WriteSecurity: 2 },
    );
  } catch { /* best-effort */ }

  if (neu) {
    try {
      await svc.configureDefaultView(POLL_LIST.answers, ['EventNumber', 'Optionen', 'Freitext', 'AnsweredAt']);
      await svc.configureDefaultView(POLL_LIST.config, ['EventNumber', 'Frage', 'Aktiv']);
    } catch { /* best-effort */ }
  }
}

/**
 * Rechte auf einer Umfrage-Liste.
 *
 * Anders als bei den Queue-Listen (`setQueueListPermissions`) wird die
 * Vererbung mit `copyRoleAssignments=true` gebrochen. Der Grund ist die
 * Item-Level-Security: Mit 2/2 sieht nur noch fremde Zeilen, wer „Listen
 * verwalten" hat. Kopiert man die Zuweisungen NICHT, verliert jeder
 * Organizer, dessen Vollzugriff am Web hängt, den Blick auf die Antworten —
 * und die Auswertung zeigte ihm dann seine eigene Zeile als Gesamtergebnis.
 * Ein unvollständiges Ergebnis, das vollständig aussieht, ist schlimmer als
 * gar keins.
 *
 * Dazu kommt die aktuell angemeldete Person: Wer die Umfrage anlegt, muss
 * sie auch auswerten können, auch wenn der Vollzugriff erst nach dem Bruch
 * vergeben wurde. Ein einzelner POST ist dabei keine Vergabe (v30.85) —
 * es wird nachgelesen.
 */
async function rechteSetzen(svc: EventService, listName: string): Promise<void> {
  const base = `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')`;
  try {
    const info = await svc._sp.get(`${base}?$select=HasUniqueRoleAssignments`, SPHttpClient.configurations.v1);
    if (info.ok) {
      const d = await info.json();
      const eigene = d.HasUniqueRoleAssignments ?? d.d?.HasUniqueRoleAssignments;
      if (!eigene) {
        await svc._post(`${base}/breakroleinheritance(copyRoleAssignments=true, clearSubscopes=true)`, {});
      }
    }
  } catch { /* best-effort */ }

  try {
    // Teilnehmende müssen schreiben dürfen — sonst endet der Klick aus der
    // Mail in einem 403, das niemand sieht.
    const besucher = await svc.getVisitorsGroupId();
    if (besucher) {
      await svc._post(`${base}/roleassignments/addroleassignment(principalid=${besucher}, roledefid=1073741827)`, {});
    }
  } catch { /* best-effort */ }

  try {
    const me = await svc._sp.get(`${svc.siteUrl}/_api/web/currentuser?$select=Id`, SPHttpClient.configurations.v1);
    if (!me.ok) return;
    const d = await me.json();
    const userId = Number(d.Id ?? d.d?.Id ?? 0);
    if (!userId) return;
    const url = `${base}/roleassignments/getbyprincipalid(${userId})/roledefinitionbindings`;
    const vorher = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (vorher.ok) {
      const b = await vorher.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ids = ((b.value || b.d?.results || []) as any[]).map(x => Number(x.Id || 0));
      if (ids.indexOf(1073741829) >= 0 || ids.indexOf(1073741830) >= 0) return; // Full oder Edit reicht
    }
    await svc._post(`${base}/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`, {});
    // Nachlesen: Der POST-Status allein ist keine Zusage (v30.85).
    const nachher = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (!nachher.ok) {
      console.warn(`[DEX] poll/rechteSetzen: Vollzugriff auf ${listName} nicht bestätigt.`);
    }
  } catch { /* best-effort */ }
}

// ===================================================================
// Konfiguration je Event
// ===================================================================

/** `null` heißt „nicht lesbar ODER keine Umfrage" — die Aufrufer unterscheiden über `ok`. */
export async function getPoll(
  svc: EventService,
  eventNumber: number,
): Promise<{ ok: boolean; poll: PollConfig | null }> {
  try {
    const r = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${POLL_LIST.config}')/items`
      + `?$filter=EventNumber eq ${eventNumber}&$top=1`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    // 404 = Liste gibt es noch nicht. Das ist keine Störung, sondern „noch
    // keine Umfrage angelegt" — deshalb ok, aber ohne Umfrage.
    if (r.status === 404) return { ok: true, poll: null };
    if (!r.ok) return { ok: false, poll: null };
    const d = await r.json();
    const row = (d.value || [])[0];
    if (!row) return { ok: true, poll: null };
    return {
      ok: true,
      poll: {
        id: row.Id,
        eventNumber,
        eventId: row.EventId || '',
        frage: row.Frage || '',
        optionen: zeilen(row.Optionen),
        mehrfach: row.Mehrfach !== false,
        anonym: !!row.Anonym,
        aktiv: !!row.Aktiv,
      },
    };
  } catch {
    return { ok: false, poll: null };
  }
}

export async function savePoll(svc: EventService, poll: PollConfig): Promise<boolean> {
  const body = {
    '__metadata': { 'type': await echterItemType(svc, POLL_LIST.config) },
    'Title': `Event ${poll.eventNumber}`,
    'EventNumber': poll.eventNumber,
    'EventId': poll.eventId,
    'Frage': poll.frage,
    'Optionen': poll.optionen.join('\n'),
    'Mehrfach': poll.mehrfach,
    'Anonym': poll.anonym,
    'Aktiv': poll.aktiv,
  };
  try {
    const basis = `${svc.siteUrl}/_api/web/lists/getbytitle('${POLL_LIST.config}')/items`;
    const r = poll.id > 0
      ? await svc._merge(`${basis}(${poll.id})`, body)
      : await svc._post(basis, body);
    if (r.ok || r.status === 406) return true;
    /*
     * v31.27: Den Grund NENNEN. Bisher kam nur `false` zurueck, und die
     * Oberflaeche sagte „bitte versuch es noch einmal" — ein Rat, der bei
     * einem 403 oder 400 nie hilft, weil ein zweiter Versuch dasselbe
     * ergibt. Der Text der Antwort steht jetzt in der Konsole; er nennt bei
     * SharePoint fast immer die Spalte oder das Recht, an dem es scheitert.
     */
    let grund = '';
    try { grund = (await r.text()).substring(0, 400); } catch { /* */ }
    console.error(`[DEX] savePoll: HTTP ${r.status} — ${grund}`);
    return false;
  } catch (e) {
    console.error('[DEX] savePoll:', e);
    return false;
  }
}

// ===================================================================
// Antworten
// ===================================================================

/** Die eigene Antwort — damit die Seite sie angehakt zeigt statt sie zu überschreiben. */
export async function getMyPollAnswer(
  svc: EventService,
  eventNumber: number,
  email: string,
): Promise<{ ok: boolean; answer: PollAnswerRow | null }> {
  const lc = (email || '').trim().toLowerCase();
  if (!lc) return { ok: true, answer: null };
  try {
    const r = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${POLL_LIST.answers}')/items`
      + `?$filter=EventNumber eq ${eventNumber} and ParticipantEmail eq '${esc(lc)}'&$top=1`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    if (r.status === 404) return { ok: true, answer: null };
    if (!r.ok) return { ok: false, answer: null };
    const d = await r.json();
    const row = (d.value || [])[0];
    if (!row) return { ok: true, answer: null };
    return {
      ok: true,
      answer: {
        id: row.Id,
        email: row.ParticipantEmail || '',
        optionen: zeilen(row.Optionen),
        freitext: row.Freitext || '',
        answeredAt: row.AnsweredAt || '',
      },
    };
  } catch {
    return { ok: false, answer: null };
  }
}

/**
 * Antwort speichern — anlegen oder die eigene Zeile ergänzen.
 *
 * `existingId` kommt von `getMyPollAnswer`. Ohne ihn wird angelegt; eine
 * zweite Zeile derselben Person wäre eine zweite Stimme, deshalb liest die
 * Seite vorher.
 *
 * Die Adresse steht auch bei einer anonymen Umfrage in der Zeile — ohne sie
 * gäbe es keinen Teilnahme-Merker und kein „hat schon geantwortet". Was
 * anonym bedeutet, entscheidet die AUSWERTUNG: `getPollStats` liest die
 * Adressen dann gar nicht erst aus.
 */
export async function savePollAnswer(
  svc: EventService,
  eventNumber: number,
  email: string,
  existingId: number,
  optionen: string[],
  freitext: string,
): Promise<boolean> {
  const lc = (email || '').trim().toLowerCase();
  const body = {
    '__metadata': { 'type': await echterItemType(svc, POLL_LIST.answers) },
    'Title': `Event ${eventNumber}`,
    'EventNumber': eventNumber,
    'ParticipantEmail': lc,
    'Optionen': optionen.join('\n'),
    'Freitext': freitext || '',
    'AnsweredAt': new Date().toISOString(),
  };
  try {
    const basis = `${svc.siteUrl}/_api/web/lists/getbytitle('${POLL_LIST.answers}')/items`;
    const r = existingId > 0
      ? await svc._merge(`${basis}(${existingId})`, body)
      : await svc._post(basis, body);
    return r.ok || r.status === 406;
  } catch (e) {
    console.error('[DEX] savePollAnswer:', e);
    return false;
  }
}

/**
 * Die eigene Antwort wieder löschen.
 *
 * Gebaut für den Test: Die Testmail trägt dieselben Links wie die echte Mail,
 * ein Klick darin ist also eine echte Antwort. Ohne diesen Weg müsste der
 * Organizer seine Probe-Stimme in der Auswertung stehen lassen — bei einer
 * anonymen Umfrage könnte er sie nicht einmal wiederfinden.
 */
export async function deletePollAnswer(
  svc: EventService,
  eventNumber: number,
  email: string,
): Promise<boolean> {
  const mein = await getMyPollAnswer(svc, eventNumber, email);
  if (!mein.ok) return false;
  if (!mein.answer) return true; // nichts da = nichts zu tun, und das ist Erfolg
  try {
    const r = await svc._delete(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${POLL_LIST.answers}')/items(${mein.answer.id})`,
    );
    return r.ok || r.status === 204;
  } catch {
    return false;
  }
}

// ===================================================================
// Auswertung
// ===================================================================

/**
 * Die Zahlen zur Umfrage.
 *
 * `nichtLesbar` ist kein Beiwerk: War die Liste nicht lesbar, sind die Zahlen
 * keine Aussage über die Umfrage, sondern über gar nichts. Die Auswertung
 * sagt das, statt „0 Antworten" zu zeigen — die Lehre, die DEX rund sechzig
 * Audit-Befunde gekostet hat (v30.66/30.67).
 */
export async function getPollStats(
  svc: EventService,
  eventNumber: number,
  anonym: boolean,
): Promise<PollStats> {
  const stats: PollStats = { geantwortet: 0, zaehler: {}, freitexte: [], emails: [], nichtLesbar: false };
  try {
    const r = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${POLL_LIST.answers}')/items`
      + `?$filter=EventNumber eq ${eventNumber}&$top=2000`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    // Eine noch nicht angelegte Liste heißt „noch keine Antwort" — das ist
    // eine Aussage, und zwar eine richtige.
    if (r.status === 404) return stats;
    if (!r.ok) { stats.nichtLesbar = true; return stats; }
    const d = await r.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rows = (d.value || []) as any[];
    stats.geantwortet = rows.length;
    for (const row of rows) {
      for (const o of zeilen(row.Optionen)) {
        stats.zaehler[o] = (stats.zaehler[o] || 0) + 1;
      }
      const ft = String(row.Freitext || '').trim();
      if (ft) stats.freitexte.push(ft);
      // Bei einer anonymen Umfrage wird die Spalte nicht einmal angefasst —
      // so kann sie auch nicht versehentlich in einen Export geraten.
      if (!anonym) {
        const e = String(row.ParticipantEmail || '').trim().toLowerCase();
        if (e) stats.emails.push(e);
      }
    }
    // Die Freitexte in eine Reihenfolge bringen, die nichts über die
    // Antwortzeit verrät — sonst ist „der erste Freitext" bei einer anonymen
    // Umfrage die Person, die zuerst geklickt hat.
    if (anonym) stats.freitexte.sort((a, b) => a.localeCompare(b, 'de'));
  } catch {
    stats.nichtLesbar = true;
  }
  return stats;
}

/**
 * Wer noch NICHT geantwortet hat — Grundlage für die Erinnerung.
 *
 * Auch bei einer anonymen Umfrage: Wer geantwortet hat, steht in der Spalte;
 * die Auswertung liest sie nur nicht. Fürs Erinnern muss sie gelesen werden,
 * und genau dafür ist der Teilnahme-Merker da.
 *
 * `null` = die Liste war nicht lesbar. Ein leeres Array hieße sonst „alle
 * haben geantwortet" und die Erinnerung ginge an niemanden.
 */
export async function getPollNonResponders(
  svc: EventService,
  eventNumber: number,
  eingeladen: string[],
): Promise<string[] | null> {
  try {
    const r = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${POLL_LIST.answers}')/items`
      + `?$filter=EventNumber eq ${eventNumber}&$select=ParticipantEmail&$top=2000`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    let fertig = new Set<string>();
    if (r.status === 404) {
      fertig = new Set<string>();
    } else if (!r.ok) {
      return null;
    } else {
      const d = await r.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      fertig = new Set<string>(((d.value || []) as any[])
        .map(x => String(x.ParticipantEmail || '').trim().toLowerCase())
        .filter(Boolean));
    }
    const out: string[] = [];
    const gesehen = new Set<string>();
    for (const e of eingeladen) {
      const lc = (e || '').trim().toLowerCase();
      if (!lc || gesehen.has(lc) || fertig.has(lc)) continue;
      gesehen.add(lc);
      out.push(lc);
    }
    return out;
  } catch {
    return null;
  }
}
