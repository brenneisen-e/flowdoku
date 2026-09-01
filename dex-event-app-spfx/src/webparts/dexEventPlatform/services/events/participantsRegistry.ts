/**
 * v30.66 — Modularisierung Stufe 2: Thema „Personen-Register"
 * (DEX_Participants): eine Zeile je Person mit den EventNumbers, für die sie
 * angemeldet oder auf der Warteliste ist — plus die Prüf- und
 * Aufräum-Läufe dazu.
 *
 * Zwei Dinge, die hier wehtun (siehe CLAUDE.md):
 * Der Schlüssel ist die E-Mail-Adresse, und die ist NICHT eindeutig (SMTP vs.
 * UPN/Alias) — dieselbe Person kann zweimal im Register stehen. Und ein leeres
 * Leseergebnis ist keine Aussage: `getAllRegistrations` wirft nicht, deshalb
 * lesen die Prüf-Läufe hier strikt (`fetchAllParticipantsOrThrow`) und brechen
 * bei einem Fehler ab, statt „verwaist" zu melden.
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService, SPParticipant, SPRegistration } from '../EventService';

// ==================== DEX_Participants Liste ====================

/**
 * Zentrale Teilnehmer-Liste erstellen falls nicht vorhanden.
 * Speichert pro Person die EventNumbers für Registrierung und Warteliste.
 */
export async function ensureParticipantsList(svc: EventService): Promise<void> {
  const listName = 'DEX_Participants';
  const exists = await svc.listExists(listName);
  if (exists) {
    await ensureMissingParticipantsFields(svc, listName);
    await ensureParticipantsIndexes(svc, listName);
    await svc.configureDefaultView(listName, [
      'Vorname', 'Nachname', 'Email', 'EventRegistered', 'EventOnWaitlist',
    ]);
    return;
  }

  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'Zentrale Teilnehmerliste der DEX Event Experience Platform',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });

  const fields = [
    { title: 'Vorname', type: 2 },
    { title: 'Nachname', type: 2 },
    { title: 'Email', type: 2 },
    { title: 'EventRegistered', type: 3 }, // Note für beliebig viele EventNumbers
    { title: 'EventOnWaitlist', type: 3 }, // Note für beliebig viele EventNumbers
  ];

  for (const f of fields) {
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
      '__metadata': { 'type': 'SP.Field' },
      'Title': f.title,
      'FieldTypeKind': f.type,
      'Required': false,
    });
  }

  await svc.configureDefaultView(listName, [
    'Vorname', 'Nachname', 'Email', 'EventRegistered', 'EventOnWaitlist',
  ]);
  await ensureParticipantsIndexes(svc, listName);

  await svc.setEmailsListPermissions(listName);
}

/**
 * v28.25: Index auf der Spalte `Email` sicherstellen.
 *
 * DEX_Participants wächst mit jeder je angemeldeten Person. Überschreitet die
 * Liste die SharePoint-Schwelle von 5000 Elementen, scheitert JEDE Abfrage,
 * die auf einer NICHT indizierten Spalte filtert oder sortiert — und zwar mit
 * HTTP 500 („exceeds the list view threshold"), nicht mit einer sprechenden
 * Meldung. Genau das legte im Tenant `getParticipantByEmail` (Filter auf
 * Email) und damit die gesamte Schattenbuchhaltung lahm: „Meine Events" blieb
 * leer, jede An-/Abmeldung konnte das Register nicht mehr fortschreiben und
 * die Reparatur-Aktion lief in eine Fehlerwand.
 *
 * Ein Index auf `Email` hebt die Sperre für genau diese Filter-Abfragen auf.
 * Best-effort: Das Setzen braucht „Listen verwalten" (Admin/Organizer der
 * Site); fehlt das Recht, bleibt alles wie bisher.
 */
async function ensureParticipantsIndexes(svc: EventService, listName: string): Promise<void> {
  try {
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Email')`;
    const resp = await svc._sp.get(`${url}?$select=Indexed`, SPHttpClient.configurations.v1);
    if (!resp.ok) return;
    const data = await resp.json();
    if (data && data.Indexed === true) return;
    const m = await svc._merge(url, { 'Indexed': true });
    if (!m.ok) console.warn('[DEX] Index auf DEX_Participants.Email konnte nicht gesetzt werden (HTTP ' + m.status + ') — bei >5000 Einträgen scheitern gefilterte Abfragen.');
  } catch (err) {
    console.warn('[DEX] ensureParticipantsIndexes fehlgeschlagen (best-effort):', err);
  }
}

async function ensureMissingParticipantsFields(svc: EventService, listName: string): Promise<void> {
  const requiredFields = [
    { title: 'Vorname', type: 2 },
    { title: 'Nachname', type: 2 },
    { title: 'Email', type: 2 },
    { title: 'EventRegistered', type: 3 },
    { title: 'EventOnWaitlist', type: 3 },
  ];

  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName,Title&$filter=Hidden eq false&$top=200`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return;
    const data = await response.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const existingNames = new Set((data.value || []).flatMap((f: any) => [f.InternalName, f.Title]));

    for (const f of requiredFields) {
      if (!existingNames.has(f.title)) {
        await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          '__metadata': { 'type': 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        });
      }
    }
  } catch { /* optional */ }
}

/**
 * Teilnehmer-Eintrag per Email suchen
 */
export async function getParticipantByEmail(svc: EventService, email: string): Promise<SPParticipant | null> {
  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items?$filter=Email eq '${email.replace(/'/g, "''")}'&$select=Id,Title,Vorname,Nachname,Email,EventRegistered,EventOnWaitlist&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return null;
    const data = await response.json();
    if (data.value && data.value.length > 0) return data.value[0];
    return null;
  } catch {
    return null;
  }
}

/**
 * Teilnehmer anlegen oder aktualisieren bei Registrierung.
 * Fügt eventNumber zu EventRegistered oder EventOnWaitlist hinzu.
 */
export async function upsertParticipant(
  svc: EventService,
  vorname: string,
  nachname: string,
  email: string,
  eventNumber: number,
  status: string // 'Angemeldet' | 'Warteliste'
): Promise<boolean> {
  try {
    const existing = await svc.getParticipantByEmail(email);
    return applyParticipantEvent(svc, existing, vorname, nachname, email, eventNumber, status);
  } catch {
    return false;
  }
}

/**
 * v28.25: Schreib-Teil von `upsertParticipant`, aber mit BEREITS bekanntem
 * Register-Eintrag. Der Massen-Abgleich (backfillParticipantRegistry) lädt
 * das Register einmal komplett und spart sich damit die Einzelabfrage pro
 * Person — bei mehreren hundert Teilnehmern hunderte Requests weniger, und
 * es funktioniert auch dann, wenn die Einzelabfrage an der 5000-Element-
 * Schwelle scheitern würde.
 */
async function applyParticipantEvent(
  svc: EventService,
  existing: SPParticipant | null,
  vorname: string,
  nachname: string,
  email: string,
  eventNumber: number,
  status: string
): Promise<boolean> {
  try {
    if (existing) {
      // EventNumber zu richtigem Feld hinzufügen
      let registered = existing.EventRegistered ? existing.EventRegistered.split(',').map(s => s.trim()).filter(s => s) : [];
      let waitlist = existing.EventOnWaitlist ? existing.EventOnWaitlist.split(',').map(s => s.trim()).filter(s => s) : [];
      const en = eventNumber.toString();

      // Erst aus beiden entfernen
      registered = registered.filter(n => n !== en);
      waitlist = waitlist.filter(n => n !== en);

      if (status === 'Warteliste') {
        waitlist.push(en);
      } else {
        registered.push(en);
      }

      await svc._merge(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${existing.Id})`,
        {
          'Vorname': vorname,
          'Nachname': nachname,
          'EventRegistered': registered.join(','),
          'EventOnWaitlist': waitlist.join(','),
        }
      );
    } else {
      // Neuen Eintrag erstellen
      const isWaitlist = status === 'Warteliste';
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items`, {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_ParticipantsListItem' },
        'Title': email,
        'Vorname': vorname,
        'Nachname': nachname,
        'Email': email,
        'EventRegistered': isWaitlist ? '' : eventNumber.toString(),
        'EventOnWaitlist': isWaitlist ? eventNumber.toString() : '',
      });
    }
    return true;
  } catch {
    return false;
  }
}

/**
 * EventNumber aus den Feldern eines Teilnehmers entfernen (bei Abmeldung).
 */
/**
 * v29.0: Register GEGEN die Teilnehmerlisten prüfen.
 *
 * `analyzeParticipantRegistry` findet nur zwei Dinge: mehrere Einträge zur
 * selben E-Mail und Verweise auf GELÖSCHTE Events. Ob eine Event-Nummer, die
 * auf ein existierendes Event zeigt, dort auch eine Zeile hat, prüft sie
 * nicht — genau daraus entsteht der Fall „Meine Events sagt angemeldet, die
 * Teilnehmerliste kennt die Person nicht" (v28.99).
 *
 * Vorgehen: einmal das Register lesen, daraus je Event-Nummer die Menge der
 * E-Mails bilden, und dann JE EVENT dessen Teilnehmerliste EINMAL laden.
 * Verglichen wird gegen aktive Zeilen — „Abgemeldet" zählt nicht, denn eine
 * abgemeldete Person gehört nicht mehr ins Register.
 *
 * WICHTIG: Schlägt das Lesen einer Liste fehl, wird das Event ÜBERSPRUNGEN
 * und als solches gezählt. Aus einem Netzwerkfehler „keine Zeile gefunden"
 * abzuleiten, würde gültige Anmeldungen aus dem Register werfen.
 */
export async function analyzeRegistryAgainstLists(
  svc: EventService,
  events: Array<{ eventNumber?: number; title: string; subsiteUrl?: string }>,
  onProgress?: (_done: number, _total: number, _title: string) => void,
  onRead?: (_loaded: number) => void,
): Promise<{
  checkedEvents: number; skippedEvents: number;
  stale: Array<{ email: string; eventNumber: number; title: string }>;
  /** Je geprüftem Event: wie viele Verweise, wie viele davon ohne Zeile,
   *  wie viele aktive Zeilen die Liste überhaupt hat. Das ist die Grundlage
   *  für die Plausibilitäts-Prüfung unten — und für die Frage, WARUM etwas
   *  auseinanderläuft. */
  perEvent: Array<{ title: string; eventNumber: number; referenced: number; missing: number; rows: number; suspicious: boolean; listGone: boolean }>;
  /** Events, bei denen (fast) ALLE Verweise ins Leere zeigen, OBWOHL ihre
   *  Teilnehmerliste lesbar ist. Ihre Verweise stehen NICHT in `stale` —
   *  siehe Begründung unten. */
  suspiciousEvents: Array<{ title: string; eventNumber: number; referenced: number; missing: number; rows: number }>;
  /** v29.3: Events, deren Teilnehmerliste NICHT MEHR EXISTIERT (HTTP 404 —
   *  in aller Regel das 3-Monats-Löschkonzept, das die Subsite recycelt und
   *  das Event-Item stehen lässt). Ihre Verweise sind eindeutig Rückstand
   *  und stehen in `stale`. */
  deletedListEvents: Array<{ title: string; eventNumber: number; referenced: number }>;
}> {
  const all = await svc.fetchAllParticipantsOrThrow(onRead);
  // Event-Nummer → E-Mails, die laut Register dort angemeldet sind.
  const byNumber: Record<number, Set<string>> = {};
  for (const p of all) {
    const em = (p.Email || '').trim().toLowerCase();
    if (!em) continue;
    `${p.EventRegistered || ''},${p.EventOnWaitlist || ''}`
      .split(',').map(x => parseInt(x.trim(), 10)).filter(n => !isNaN(n) && n > 0)
      .forEach(n => { (byNumber[n] = byNumber[n] || new Set<string>()).add(em); });
  }
  const relevant = events.filter(e =>
    typeof e.eventNumber === 'number' && e.eventNumber > 0
    && !!e.subsiteUrl && !!byNumber[e.eventNumber]);
  const stale: Array<{ email: string; eventNumber: number; title: string }> = [];
  const perEvent: Array<{ title: string; eventNumber: number; referenced: number; missing: number; rows: number; suspicious: boolean; listGone: boolean }> = [];
  const suspiciousEvents: Array<{ title: string; eventNumber: number; referenced: number; missing: number; rows: number }> = [];
  const deletedListEvents: Array<{ title: string; eventNumber: number; referenced: number }> = [];
  let checkedEvents = 0;
  let skippedEvents = 0;
  for (let i = 0; i < relevant.length; i++) {
    const ev = relevant[i];
    const num = ev.eventNumber as number;
    if (onProgress) onProgress(i + 1, relevant.length, ev.title || '');
    let rows: SPRegistration[] | null = null;
    /**
     * v29.3: Der Lesefehler wird jetzt AUSGEWERTET statt verschluckt.
     * `getAllRegistrations` bricht bei einem HTTP-Fehler ab und liefert eine
     * leere Liste — bis v29.2 war deshalb „Subsite existiert nicht mehr"
     * (404) von „niemand ist angemeldet" nicht zu unterscheiden. Genau das
     * ist hier der Regelfall und nicht die Ausnahme: Das 3-Monats-
     * Löschkonzept (`deleteParticipantData`) recycelt die Subsite und LÄSST
     * das Event-Item stehen. Das Event ist danach weiter in der Liste, seine
     * Teilnehmerliste aber weg — und alle Register-Verweise darauf sind
     * Rückstand, den genau dieses Löschkonzept hinterlassen hat.
     */
    let httpStatus = 0;
    let hadHttpError = false;
    try {
      // eslint-disable-next-line no-await-in-loop
      rows = await svc.getAllRegistrations(ev.subsiteUrl as string, s => { hadHttpError = true; httpStatus = s; });
    } catch { rows = null; }
    // 404/410 = Liste/Subsite gibt es nicht mehr → eindeutige Aussage.
    const listGone = hadHttpError && (httpStatus === 404 || httpStatus === 410);
    // Jeder ANDERE Fehler (403 fehlende Rechte, 429 Drosselung, 500, Netz)
    // sagt nichts über den Inhalt aus — solche Events werden übersprungen,
    // nicht geraten.
    if (!rows || (hadHttpError && !listGone)) { skippedEvents += 1; continue; }
    checkedEvents += 1;
    if (listGone) {
      const refsGone = Array.from(byNumber[num]);
      perEvent.push({ title: ev.title || '', eventNumber: num, referenced: refsGone.length, missing: refsGone.length, rows: 0, suspicious: false, listGone: true });
      deletedListEvents.push({ title: ev.title || '', eventNumber: num, referenced: refsGone.length });
      refsGone.forEach(em => { stale.push({ email: em, eventNumber: num, title: ev.title || '' }); });
      continue;
    }
    const active = new Set<string>();
    for (const r of rows) {
      const st = (r.Status || '').trim();
      if (st === 'Abgemeldet') continue;
      const em = (r.ParticipantEmail || '').trim().toLowerCase();
      if (em) active.add(em);
    }
    const refs = Array.from(byNumber[num]);
    const missing = refs.filter(em => !active.has(em));
    /**
     * Plausibilitäts-Riegel — gilt seit v29.3 nur noch für Events mit
     * LESBARER Teilnehmerliste. Wenn dort praktisch alle Verweise ins Leere
     * zeigen, ist die naheliegende Erklärung nicht, dass hunderte
     * Abmeldungen einzeln schiefgingen — sondern dass Register und Liste bei
     * diesem Event gar nicht vergleichbar sind. Denkbare Gründe: eine andere
     * Teilnehmerliste als die Standard-Liste, eine neu angelegte/geleerte
     * Liste bei erhaltenem Register, oder ein Event, das seine Anmeldungen
     * woanders führt.
     *
     * In dem Fall wäre ein Entfernen der Verweise ein Datenverlust, kein
     * Aufräumen. Solche Events werden deshalb ausgewiesen, aber NICHT
     * bereinigt — die Entscheidung darüber braucht einen Blick in die
     * betroffene Liste, nicht einen Knopfdruck.
     */
    const suspicious = refs.length >= 5 && missing.length >= Math.ceil(refs.length * 0.9);
    perEvent.push({ title: ev.title || '', eventNumber: num, referenced: refs.length, missing: missing.length, rows: active.size, suspicious, listGone: false });
    if (suspicious) {
      suspiciousEvents.push({ title: ev.title || '', eventNumber: num, referenced: refs.length, missing: missing.length, rows: active.size });
      continue;
    }
    missing.forEach(em => { stale.push({ email: em, eventNumber: num, title: ev.title || '' }); });
  }
  perEvent.sort((a, b) => b.missing - a.missing);
  deletedListEvents.sort((a, b) => b.referenced - a.referenced);
  return { checkedEvents, skippedEvents, stale, perEvent, suspiciousEvents, deletedListEvents };
}

/**
 * v29.4: Alle vergebenen `EventNumber` direkt aus `DEX_Events` lesen —
 * strikt, also mit Fehler statt stiller Teilliste.
 *
 * Bewusst NICHT die im Client geladene Event-Liste: `loadEvents` lässt
 * einzelne Events aus, wenn ihr Mapping scheitert (v9.41, damit ein kaputtes
 * Event nicht die ganze Liste kippt). Für eine Anzeige ist das richtig — für
 * die Frage „gibt es dieses Event noch?" wäre es fatal, weil ein
 * ausgelassenes Event wie ein gelöschtes aussähe und seine Verweise entfernt
 * würden.
 *
 * Fensterung nach `Id` wie in `readAllParticipants`: schwellenfest und
 * unabhängig vom nextLink-Format.
 */
async function readAllEventNumbersOrThrow(svc: EventService): Promise<Set<number>> {
  const out = new Set<number>();
  const PAGE = 2000;
  const MAX_PAGES = 100;
  let lastId = 0;
  for (let page = 0; page < MAX_PAGES; page++) {
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`
      + `?$select=Id,EventNumber&$filter=Id gt ${lastId}&$orderby=Id asc&$top=${PAGE}`;
    // eslint-disable-next-line no-await-in-loop
    const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (!resp.ok) throw new Error(`DEX_Events nicht vollständig lesbar (HTTP ${resp.status}).`);
    // eslint-disable-next-line no-await-in-loop
    const data = await resp.json();
    const items = (data.value || data.d?.results || []) as Array<{ Id: number; EventNumber?: number }>;
    if (items.length === 0) break;
    items.forEach(it => {
      if (typeof it.EventNumber === 'number' && it.EventNumber > 0) out.add(it.EventNumber);
      if (typeof it.Id === 'number' && it.Id > lastId) lastId = it.Id;
    });
    if (items.length < PAGE) break;
  }
  return out;
}

/**
 * v29.4: Verweise im Register auf Event-Nummern, die es in `DEX_Events`
 * NICHT MEHR GIBT. Bis v29.3 wurden die nur gezählt („wirkungslos, aber
 * harmlos") — sie sind aber personenbezogener Rückstand gelöschter Events
 * und gehören weg.
 *
 * Zwei Riegel, weil ein Fehlurteil hier das ganze Register leeren würde:
 *  - Event-Nummern und Register werden BEIDE strikt gelesen; ein Lesefehler
 *    wirft, statt eine Teilmenge als „alles" zu behandeln.
 *  - Eine leere Nummern-Menge wird als Fehler gewertet, nicht als „es gibt
 *    keine Events mehr".
 */
export async function collectOrphanRegistryNumbers(
  svc: EventService,
  onRead?: (_loaded: number) => void,
): Promise<Array<{ email: string; eventNumber: number }>> {
  const valid = await readAllEventNumbersOrThrow(svc);
  if (valid.size === 0) throw new Error('Keine Event-Nummern gefunden — Abbruch, statt alle Verweise als verwaist zu werten.');
  const all = await svc.fetchAllParticipantsOrThrow(onRead);
  const out: Array<{ email: string; eventNumber: number }> = [];
  for (const p of all) {
    const em = (p.Email || '').trim().toLowerCase();
    if (!em) continue;
    const nums = `${p.EventRegistered || ''},${p.EventOnWaitlist || ''}`
      .split(',').map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n) && n > 0);
    const seen = new Set<number>();
    nums.forEach(n => { if (!valid.has(n) && !seen.has(n)) { seen.add(n); out.push({ email: em, eventNumber: n }); } });
  }
  return out;
}

/**
 * v29.0: Die von `analyzeRegistryAgainstLists` gefundenen Verweise aus dem
 * Register nehmen. Je E-Mail EIN Schreibvorgang, auch wenn mehrere Nummern
 * betroffen sind. Der Eintrag selbst bleibt stehen — er kann weitere,
 * gültige Events tragen.
 */
export async function pruneStaleRegistryNumbers(
  svc: EventService,
  stale: Array<{ email: string; eventNumber: number }>,
  onProgress?: (_done: number, _total: number) => void,
): Promise<{ updated: number; removed: number; failed: number }> {
  const byEmail: Record<string, number[]> = {};
  stale.forEach(s => { (byEmail[s.email] = byEmail[s.email] || []).push(s.eventNumber); });
  const emails = Object.keys(byEmail);
  let updated = 0; let removed = 0; let failed = 0;
  for (let i = 0; i < emails.length; i++) {
    const em = emails[i];
    if (onProgress) onProgress(i + 1, emails.length);
    try {
      // eslint-disable-next-line no-await-in-loop
      const rec = await svc.getParticipantByEmail(em);
      if (!rec) { failed += 1; continue; }
      const drop = new Set(byEmail[em].map(n => String(n)));
      const keep = (v?: string): string => (v || '').split(',').map(x => x.trim())
        .filter(x => x && !drop.has(x)).join(',');
      const nextReg = keep(rec.EventRegistered);
      const nextWait = keep(rec.EventOnWaitlist);
      // eslint-disable-next-line no-await-in-loop
      const resp = await svc._merge(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${rec.Id})`,
        { 'EventRegistered': nextReg, 'EventOnWaitlist': nextWait },
      );
      if (resp.ok) { updated += 1; removed += byEmail[em].length; }
      else failed += 1;
    } catch { failed += 1; }
  }
  return { updated, removed, failed };
}

export async function removeParticipantEvent(svc: EventService, email: string, eventNumber: number): Promise<boolean> {
  try {
    const existing = await svc.getParticipantByEmail(email);
    if (!existing) return false;

    const en = eventNumber.toString();
    const registered = existing.EventRegistered ? existing.EventRegistered.split(',').map(s => s.trim()).filter(s => s && s !== en) : [];
    const waitlist = existing.EventOnWaitlist ? existing.EventOnWaitlist.split(',').map(s => s.trim()).filter(s => s && s !== en) : [];

    await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${existing.Id})`,
      {
        'EventRegistered': registered.join(','),
        'EventOnWaitlist': waitlist.join(','),
      }
    );
    return true;
  } catch {
    return false;
  }
}

/**
 * Alle Teilnehmer laden (für Admin-Seite).
 */
/**
 * v28.27: Das Register vollständig lesen — per ID-Fenster statt über den
 * „nextLink".
 *
 * Zwei Fallen stecken hier drin, und v28.25/26 sind in beide getreten:
 *  1. **Schwelle:** `$orderby` über eine nicht indizierte Spalte scheitert ab
 *     5000 Listenelementen mit HTTP 500. Nach `Id` zu sortieren ist dagegen
 *     immer erlaubt — die ID-Spalte ist von Haus aus indiziert.
 *  2. **nextLink:** SharePoint benennt den Folgeseiten-Link je nach
 *     ausgehandeltem OData-Format unterschiedlich (`odata.nextLink`,
 *     `@odata.nextLink`, `d.__next`). Wer nur eine Variante prüft, hält nach
 *     der ERSTEN Seite an und meldet fröhlich Vollständigkeit — genau das
 *     ließ die Dubletten-Prüfung „2000 Einträge geprüft" melden, obwohl die
 *     Liste ein Vielfaches davon enthält.
 *
 * Deshalb hier gar kein nextLink mehr: Wir holen aufsteigend nach `Id` und
 * setzen als Fenster `Id gt <letzte gelesene Id>`. Das ist deterministisch,
 * schwellenfest und formatunabhängig.
 *
 * @param strict wirft bei einem HTTP-Fehler, statt still eine unvollständige
 *   Liste zu liefern (für Abläufe, die aus dem Ergebnis auf „unbekannt"
 *   schließen).
 */
async function readAllParticipants(svc: EventService, strict: boolean, onPage?: (loaded: number) => void): Promise<SPParticipant[]> {
  const out: SPParticipant[] = [];
  const PAGE = 2000;
  const MAX_PAGES = 100; // Reißleine (200k Einträge) gegen Endlosschleifen
  let lastId = 0;
  for (let page = 0; page < MAX_PAGES; page++) {
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items`
      + `?$select=Id,Title,Vorname,Nachname,Email,EventRegistered,EventOnWaitlist`
      + `&$orderby=Id&$top=${PAGE}&$filter=${encodeURIComponent(`Id gt ${lastId}`)}`;
    let items: SPParticipant[] = [];
    try {
      const response = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) {
        if (strict) {
          throw new Error(`DEX_Participants nicht lesbar (HTTP ${response.status}). Bei mehr als 5000 Einträgen braucht die Spalte „Email" einen Index — die App versucht ihn beim Start automatisch zu setzen (erfordert „Listen verwalten").`);
        }
        break;
      }
      const data = await response.json();
      items = (data.value || data.d?.results || []) as SPParticipant[];
    } catch (err) {
      if (strict) throw err;
      break;
    }
    out.push(...items);
    // v28.29: Nach jeder Seite melden — das Register hat inzwischen mehrere
    // tausend Zeilen, der Lesevorgang dauert spürbar. Ohne Rückmeldung sah
    // die Wartungs-Kachel aus, als würde nichts passieren.
    if (onPage) { try { onPage(out.length); } catch { /* UI-Fehler nie durchreichen */ } }
    if (items.length < PAGE) break;
    const last = items[items.length - 1];
    if (!last || typeof last.Id !== 'number' || last.Id <= lastId) break; // Schutz vor Stillstand
    lastId = last.Id;
  }
  return out;
}

export async function getAllParticipants(svc: EventService): Promise<SPParticipant[]> {
  const allItems = await readAllParticipants(svc, false);
  allItems.sort((a, b) =>
    (a.Nachname || '').localeCompare(b.Nachname || '', 'de')
    || (a.Vorname || '').localeCompare(b.Vorname || '', 'de'));
  return allItems;
}

/**
 * v28.23: Teilnehmer-Register (DEX_Participants) für EIN Event nachziehen.
 *
 * DEX_Participants ist die zentrale „Schattenbuchhaltung": Pro Person stehen
 * dort die Event-Nummern, für die sie angemeldet ist bzw. auf der Warteliste
 * steht. Sie liegt auf der Haupt-Site und unterliegt NICHT der
 * Item-Level-Security der Teilnehmerlisten — deshalb ist sie die einzige
 * Quelle, die auch stellvertretend angelegte Anmeldungen zuverlässig kennt.
 * „Meine Events" startet von hier, und seit v28.22 hängt auch die
 * Doppel-Anmelde-Vorwarnung daran.
 *
 * Der Dual-Write bei jeder Anmeldung ist best-effort (`.catch(warn)`) —
 * schlägt er fehl (Netzwerk, Rechte, Timeout), fehlt der Eintrag dauerhaft.
 * Diese Methode gleicht ihn für die übergebene Teilnehmerliste ab: Sie
 * ergänzt fehlende Event-Nummern und korrigiert Einträge, die im falschen
 * Feld stehen (Warteliste ↔ angemeldet). Es wird NICHTS entfernt — für
 * abgemeldete Personen räumt der normale Abmelde-Pfad auf.
 */
/**
 * v28.26: Zustand des zentralen Teilnehmer-Registers analysieren.
 *
 * Über die Jahre sammeln sich dort zwei Sorten Altlasten:
 *  - **Dubletten:** mehrere Einträge zur selben E-Mail. Sie entstehen, wenn
 *    der Lookup vor dem Schreiben scheitert (z.B. der HTTP-500-Fall aus
 *    v28.25) — dann legt die App einen zweiten Eintrag an, und ab da landen
 *    Anmeldungen mal im einen, mal im anderen. „Meine Events" zeigt dann je
 *    nach Treffer nur einen Teil der Events.
 *  - **Verwaiste Event-Nummern:** Anmeldungen zu Events, die es nicht mehr
 *    gibt. Beim Löschen eines Events räumt die App das Register NICHT mit
 *    auf. Harmlos (die Nummer läuft ins Leere), aber Ballast.
 */
export async function analyzeParticipantRegistry(
  svc: EventService,
  validEventNumbers: number[],
  onRead?: (loaded: number) => void,
): Promise<{
  total: number; duplicateGroups: number; surplusRecords: number; orphanNumbers: number; noEmail: number;
}> {
  const all = await svc.fetchAllParticipantsOrThrow(onRead);
  const valid = new Set(validEventNumbers.filter(n => typeof n === 'number' && n > 0));
  const byEmail: Record<string, SPParticipant[]> = {};
  let noEmail = 0;
  let orphanNumbers = 0;
  for (const p of all) {
    const em = (p.Email || '').trim().toLowerCase();
    if (!em) { noEmail += 1; continue; }
    (byEmail[em] = byEmail[em] || []).push(p);
    if (valid.size > 0) {
      const nums = `${p.EventRegistered || ''},${p.EventOnWaitlist || ''}`
        .split(',').map(s => parseInt(s.trim(), 10)).filter(n => !isNaN(n) && n > 0);
      orphanNumbers += nums.filter(n => !valid.has(n)).length;
    }
  }
  let duplicateGroups = 0;
  let surplusRecords = 0;
  Object.keys(byEmail).forEach(em => {
    const n = byEmail[em].length;
    if (n > 1) { duplicateGroups += 1; surplusRecords += n - 1; }
  });
  return { total: all.length, duplicateGroups, surplusRecords, orphanNumbers, noEmail };
}

/**
 * v28.26: Dubletten im Teilnehmer-Register zusammenführen.
 *
 * Je E-Mail bleibt der ÄLTESTE Eintrag (kleinste Id) stehen und erhält die
 * VEREINIGUNG aller Event-Nummern; steht dieselbe Nummer bei einem Eintrag
 * als „angemeldet" und beim anderen als „Warteliste", gewinnt „angemeldet".
 * Name-Felder werden aus dem ersten nicht-leeren Wert aufgefüllt. Die
 * überzähligen Einträge werden danach gelöscht. Es gehen also KEINE
 * Anmelde-Informationen verloren — im Gegenteil, die zusammengeführte Zeile
 * kennt danach alle Events der Person.
 */
export async function mergeDuplicateParticipants(
  svc: EventService,
  onProgress?: (done: number, total: number) => void,
  onRead?: (loaded: number) => void,
): Promise<{ groups: number; deleted: number; failed: number }> {
  const all = await svc.fetchAllParticipantsOrThrow(onRead);
  const byEmail: Record<string, SPParticipant[]> = {};
  for (const p of all) {
    const em = (p.Email || '').trim().toLowerCase();
    if (!em) continue;
    (byEmail[em] = byEmail[em] || []).push(p);
  }
  const groups = Object.keys(byEmail).filter(em => byEmail[em].length > 1);
  let deleted = 0;
  let failed = 0;
  let done = 0;
  const parseNums = (s?: string): string[] => (s || '').split(',').map(x => x.trim()).filter(Boolean);
  for (const em of groups) {
    const recs = byEmail[em].slice().sort((a, b) => a.Id - b.Id);
    const keeper = recs[0];
    const registered = new Set<string>();
    const waitlist = new Set<string>();
    let vorname = '';
    let nachname = '';
    for (const r of recs) {
      parseNums(r.EventRegistered).forEach(n => registered.add(n));
      parseNums(r.EventOnWaitlist).forEach(n => waitlist.add(n));
      if (!vorname && (r.Vorname || '').trim()) vorname = (r.Vorname || '').trim();
      if (!nachname && (r.Nachname || '').trim()) nachname = (r.Nachname || '').trim();
    }
    // „Angemeldet" sticht „Warteliste" — dieselbe Nummer nie in beiden Feldern.
    registered.forEach(n => waitlist.delete(n));
    try {
      const m = await svc._merge(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${keeper.Id})`,
        {
          'Vorname': vorname,
          'Nachname': nachname,
          'EventRegistered': Array.from(registered).join(','),
          'EventOnWaitlist': Array.from(waitlist).join(','),
        },
      );
      if (!m.ok) { failed += 1; done += 1; if (onProgress) onProgress(done, groups.length); continue; }
      for (const r of recs.slice(1)) {
        try {
          const resp = await svc._sp.post(
            `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${r.Id})`,
            SPHttpClient.configurations.v1,
            { headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', 'Accept': 'application/json;odata=nometadata' } },
          );
          if (resp.ok) deleted += 1; else failed += 1;
        } catch { failed += 1; }
      }
    } catch { failed += 1; }
    done += 1;
    if (onProgress) onProgress(done, groups.length);
  }
  return { groups: groups.length, deleted, failed };
}

/**
 * v28.25: Wie `getAllParticipants`, wirft aber bei einem HTTP-Fehler, statt
 * still eine unvollständige Liste zu liefern. Für Abläufe, die aus dem
 * Ergebnis auf „Person ist unbekannt" schließen (Register-Abgleich).
 */
export async function fetchAllParticipantsOrThrow(svc: EventService, onPage?: (loaded: number) => void): Promise<SPParticipant[]> {
  return readAllParticipants(svc, true, onPage);
}

export async function backfillParticipantRegistry(
  svc: EventService,
  subsiteUrl: string,
  eventNumber: number,
  onProgress?: (done: number, total: number) => void,
): Promise<{ active: number; fixed: number; failed: number }> {
  const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
  const regs = await svc.getAllRegistrations(subsiteUrl);
  const active = regs.filter(r => ACTIVE.indexOf(r.Status || '') >= 0
    && (r.ParticipantEmail || '').indexOf('@') > 0);
  if (active.length === 0 || !eventNumber) return { active: 0, fixed: 0, failed: 0 };
  // v28.25: Register EINMAL laden — und einen Lesefehler NICHT verschlucken.
  // Wäre die Liste nicht lesbar (z.B. 5000-Element-Schwelle) und wir liefen
  // trotzdem weiter, hielte der Abgleich jede Person für unbekannt und legte
  // reihenweise Dubletten an. Lieber sauber abbrechen.
  const all = await svc.fetchAllParticipantsOrThrow();
  const byEmail: Record<string, SPParticipant> = {};
  for (const p of all) {
    const e = (p.Email || '').trim().toLowerCase();
    if (e) byEmail[e] = p;
  }
  const en = String(eventNumber);
  const has = (field: string | undefined): boolean =>
    (field || '').split(',').map(s => s.trim()).indexOf(en) >= 0;
  let fixed = 0;
  let failed = 0;
  let done = 0;
  for (const r of active) {
    const em = (r.ParticipantEmail || '').trim().toLowerCase();
    const rec = byEmail[em];
    const wantWaitlist = r.Status === 'Warteliste';
    const alreadyRight = !!rec && (wantWaitlist ? has(rec.EventOnWaitlist) : has(rec.EventRegistered));
    if (!alreadyRight) {
      // Bekannten Register-Eintrag direkt mitgeben — spart die Einzelabfrage
      // pro Person (die an der Schwelle ohnehin scheitern könnte).
      const ok = await applyParticipantEvent(svc, 
        rec || null, r.Vorname || '', r.Nachname || '', r.ParticipantEmail,
        eventNumber, wantWaitlist ? 'Warteliste' : 'Angemeldet',
      );
      if (ok) fixed += 1; else failed += 1;
    }
    done += 1;
    if (onProgress && (done % 10 === 0 || done === active.length)) onProgress(done, active.length);
  }
  return { active: active.length, fixed, failed };
}
