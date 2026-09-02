/**
 * v30.66 — Modularisierung Stufe 2: Thema „Platz-Zähler" — die Felder
 * SeatsTaken/SeatsTakenDurch/SeatsTakenFun auf der Counter-Liste, ihre
 * atomare Reservierung (ETag), das Nachziehen an den Ist-Stand und die
 * Wartelisten-Reihenfolge.
 *
 * Die Warteliste hat keine Positionsspalte: Die Position ist der Rang nach
 * TeilnehmerID — deshalb sortiert `setWaitlistPosition` die IDs um, statt
 * eine Prioritätsspalte zu schreiben (der Flow DEX_IDReorder würde sie
 * ignorieren, siehe CLAUDE.md).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { subscribeListChanges } from '../../utils/spListRealtime';
import { dlog } from '../../utils/debugLog';
import { ACTIVE_STATI } from '../EventService';
import type { EventService } from '../EventService';
import { COUNTER_LIST_NAME, REG_LIST_NAME } from '../EventService';

/**
 * TeilnehmerIDs sequentiell neu vergeben (1, 2, 3, ...).
 *
 * Seit v6.5: Zwei-Pass-Reorder.
 * 1. Erst alle **Angemeldeten** (Status ∈ Angemeldet / QR versendet / Eingecheckt)
 *    in Reihenfolge der Registrierung (SP-ItemId asc) → bekommen IDs 1..N.
 * 2. Danach alle **Warteliste**-Teilnehmer in Reihenfolge der Registrierung
 *    → bekommen IDs N+1..N+M (Warteliste hängt lückenlos hinten an).
 * 3. Abgemeldete bekommen TeilnehmerID = null.
 *
 * Damit stehen im Teilnehmerlisten-Grid die Angemeldeten sauber oben, die
 * Warteliste sauber unten — kein Durchmischen mehr.
 */
/**
 * v28.70: Eine Person auf der Warteliste an eine bestimmte Position setzen
 * (z.B. „ganz nach vorn"), ohne die Reihenfolge der Angemeldeten anzufassen.
 *
 * Warum über die TeilnehmerID? Die Warteliste hat KEINE eigene
 * Positionsspalte — die Position ist der Rang innerhalb von
 * Status='Warteliste', sortiert nach `TeilnehmerID asc` (Gleichstand: Item-Id).
 * Genau so sortieren beide Stellen, die nachrücken: `promoteFirstWaitlistItem`
 * in der App UND der Flow `DEX_IDReorder_TeilnehmerIDs` (Order By
 * `TeilnehmerID asc`, s. docs/flow-jsons.md). Eine zusätzliche Prioritaets-
 * spalte würde der Flow ignorieren und weiter den Falschen nachrücken
 * lassen — deshalb wird hier die TeilnehmerID selbst umsortiert. Damit zieht
 * die neue Reihenfolge in beiden Pfaden.
 *
 * Die Angemeldeten behalten ihre IDs (ihre relative Reihenfolge ändert sich
 * nicht, sie werden auf dieselben 1..N abgebildet); es verschieben sich nur
 * die IDs innerhalb der Warteliste.
 *
 * @param targetPosition 1-basiert. Wird auf 1..(Anzahl Wartende) geklemmt.
 * @param group v30.67: Gruppe (`PreferredStarterType`) bei geteilten
 *   Kapazitäten mit getrennten Wartelisten. Dann sind `targetPosition`,
 *   `from` und `to` der Rang INNERHALB dieser Gruppe — genau die Zahl, die
 *   das Organizer Center anzeigt (waitlistTruePos rankt je Gruppe) und nach
 *   der `promoteFirstWaitlistItem` mit `onlyWithPreferredType` nachrückt.
 *   Bisher galt der Wert als Index in die GESAMT-Warteliste: „Platz 2" in
 *   der Funstarter-Tabelle schob die Person zwischen die ersten beiden
 *   Durchstarter-Wartenden und nummerierte deren ganze Liste um.
 *   `undefined` = eine gemeinsame Warteliste (Normalfall und
 *   `splitSharedWaitlist`); `''` = die Zeilen OHNE (oder mit unbekannter)
 *   Gruppe — die dritte Tabelle „Warteliste ohne Gruppe" im Organizer Center.
 */
export async function setWaitlistPosition(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  targetPosition: number,
  group?: string
): Promise<{ ok: boolean; from: number; to: number; changed: number; error?: string }> {
  const fail = (error: string): { ok: boolean; from: number; to: number; changed: number; error: string } =>
    ({ ok: false, from: 0, to: 0, changed: 0, error });
  try {
    type Row = { Id: number; Status: string; TeilnehmerID: number | null; PreferredStarterType?: string | null };
    const allItems: Row[] = [];
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,Status,TeilnehmerID,PreferredStarterType&$orderby=Id asc&$top=5000`;
    while (url) {
      const response = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) return fail(`Teilnehmerliste konnte nicht gelesen werden (HTTP ${response.status}).`);
      const data = await response.json();
      allItems.push(...(data.value || data.d?.results || []));
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    }
    const NO_TID = Number.MAX_SAFE_INTEGER;
    const byTidThenId = (a: { Id: number; TeilnehmerID: number | null }, b: { Id: number; TeilnehmerID: number | null }): number =>
      ((a.TeilnehmerID ?? NO_TID) - (b.TeilnehmerID ?? NO_TID)) || (a.Id - b.Id);
    const activeItems = allItems
      .filter(i => i.Status === 'Angemeldet' || i.Status === 'QR versendet' || i.Status === 'Eingecheckt')
      .sort(byTidThenId);
    // v30.67: Gesamt-Warteliste (bestimmt die globale ID-Folge) und die
    // Zielgruppe darin (bestimmt, was umsortiert wird). Ohne `group` sind
    // beide dieselbe Liste — das bisherige Verhalten.
    // v30.67 (Review): `undefined` (Gesamtliste) von `''` (Zeilen ohne
    // Gruppe) trennen — vorher fielen beide auf „alle", und ein Platzwechsel
    // in der Tabelle „Warteliste ohne Gruppe" sortierte die Person über die
    // komplette Warteliste, vor alle Durchstarter- und Funstarter-Wartenden.
    // normGroup ist dieselbe Definition wie waitlistUnassigned in AdminPage.
    const normGroup = (v: string | null | undefined): string => {
      const t = (v || '').trim();
      return (t === 'Durchstarter' || t === 'Funstarter') ? t : '';
    };
    const inGroup = (i: Row): boolean => group === undefined || normGroup(i.PreferredStarterType) === normGroup(group);
    const waitlistAll = allItems.filter(i => i.Status === 'Warteliste').sort(byTidThenId);
    const waitlist = waitlistAll.filter(inGroup);

    const fromIdx = waitlist.findIndex(i => i.Id === itemId);
    if (fromIdx < 0) return fail('Diese Person steht nicht (mehr) auf der Warteliste — bitte die Liste neu laden.');
    if (waitlist.length < 2) return fail('Auf der Warteliste steht nur diese eine Person — es gibt nichts umzusortieren.');
    const toIdx = Math.max(0, Math.min(waitlist.length - 1, Math.round(targetPosition) - 1));
    if (toIdx === fromIdx) {
      return { ok: true, from: fromIdx + 1, to: toIdx + 1, changed: 0 };
    }
    const moved = waitlist.splice(fromIdx, 1)[0];
    waitlist.splice(toIdx, 0, moved);

    // Ziel-IDs: Angemeldete 1..N (unverändert), danach die neue
    // Warteliste-Reihenfolge, Abgemeldete null.
    // v30.67: Die Plätze der Gesamt-Warteliste, die zur Zielgruppe gehören,
    // werden in der neuen Gruppen-Reihenfolge belegt; Zeilen fremder Gruppen
    // behalten Platz und ID — nur der Rang INNERHALB der Gruppe ändert sich.
    let gi = 0;
    const mergedWaitlist = waitlistAll.map(i => (inGroup(i) ? waitlist[gi++] : i));
    const targetIds = new Map<number, number | null>();
    let nextId = 1;
    for (const item of activeItems) targetIds.set(item.Id, nextId++);
    for (const item of mergedWaitlist) targetIds.set(item.Id, nextId++);
    for (const item of allItems) if (!targetIds.has(item.Id)) targetIds.set(item.Id, null);

    let changed = 0;
    for (const item of allItems) {
      const newId = targetIds.get(item.Id) ?? null;
      if (newId === item.TeilnehmerID) continue;
      const resp = await svc._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${item.Id})`,
        { 'TeilnehmerID': newId }
      );
      if (resp.ok || resp.status === 406) { changed++; }
      else { return fail(`Position konnte nicht vollständig gesetzt werden (HTTP ${resp.status}). Bitte „Teilnehmer-IDs neu vergeben" ausführen und erneut versuchen.`); }
    }
    try { await svc.syncCounterToMax(subsiteUrl); } catch { /* best-effort */ }
    return { ok: true, from: fromIdx + 1, to: toIdx + 1, changed };
  } catch (err) {
    return fail(`Unerwarteter Fehler: ${err instanceof Error ? err.message.slice(0, 160) : 'unbekannt'}`);
  }
}

export async function reorderParticipantIDs(
  svc: EventService,
  subsiteUrl: string,
  onProgress?: (pct: number) => void
): Promise<{ success: number; errors: number }> {
  // Alle Items laden, sortiert nach SP Id (Erstellungsreihenfolge = Reihenfolge der Registrierung)
  const allItems: Array<{ Id: number; Status: string; TeilnehmerID: number | null }> = [];
  let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,Status,TeilnehmerID&$orderby=Id asc&$top=5000`;

  // v30.66: Ein Lesefehler darf hier NICHT einfach die Schleife verlassen.
  // Danach wird ueber `allItems` renummeriert — auf einer halb gelesenen Liste
  // bekaemen die gelesenen Zeilen die IDs 1..N, waehrend die ungelesenen ihre
  // alten behalten. Das Ergebnis waeren DOPPELTE TeilnehmerIDs, und die Methode
  // meldete dazu „0 Fehler", weil nur die Schreibvorgaenge gezaehlt werden.
  // Die Regel aus CLAUDE.md gilt auch hier: bei Fehlern abbrechen, BEVOR
  // etwas geschrieben wird. Beide Aufrufer fangen die Ausnahme ab; nicht
  // renummeriert ist allemal besser als falsch renummeriert.
  while (url) {
    let response: SPHttpClientResponse;
    try {
      response = await svc._sp.get(url, SPHttpClient.configurations.v1);
    } catch {
      throw new Error('reorderParticipantIDs: Teilnehmerliste nicht vollstaendig lesbar — abgebrochen, es wurde nichts geaendert.');
    }
    if (!response.ok) {
      throw new Error('reorderParticipantIDs: Teilnehmerliste nicht vollstaendig lesbar (HTTP ' + response.status + ') — abgebrochen, es wurde nichts geaendert.');
    }
    const data = await response.json();
    allItems.push(...(data.value || data.d?.results || []));
    url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
  }

  // Ziel-IDs in einem ersten Durchlauf berechnen: erst Angemeldete, dann Warteliste.
  // v22.20: Innerhalb der Gruppen wird primär nach der VORHANDENEN TeilnehmerID
  // sortiert (aufsteigend, ohne ID ans Ende), Gleichstand nach Item-Id. Damit ist
  // die Client-Renummerierung deckungsgleich mit dem DEX_IDReorder-Flow (der
  // ebenfalls nach TeilnehmerID sortiert) — egal welcher Pfad zuletzt lief, das
  // Ergebnis ist identisch und die IDs „springen" nicht mehr zwischen
  // Admin-Button und Flow-Lauf hin und her.
  const targetIds = new Map<number, number | null>();
  let nextId = 1;
  const NO_TID = Number.MAX_SAFE_INTEGER;
  const byTidThenId = (a: { Id: number; TeilnehmerID: number | null }, b: { Id: number; TeilnehmerID: number | null }): number =>
    ((a.TeilnehmerID ?? NO_TID) - (b.TeilnehmerID ?? NO_TID)) || (a.Id - b.Id);
  // Pass 1: Angemeldete / QR versendet / Eingecheckt
  const activeItems = allItems
    .filter(item => item.Status === 'Angemeldet' || item.Status === 'QR versendet' || item.Status === 'Eingecheckt')
    .sort(byTidThenId);
  for (const item of activeItems) {
    targetIds.set(item.Id, nextId++);
  }
  // Pass 2: Warteliste
  const waitlistItems = allItems
    .filter(item => item.Status === 'Warteliste')
    .sort(byTidThenId);
  for (const item of waitlistItems) {
    targetIds.set(item.Id, nextId++);
  }
  // Pass 3: Abgemeldete (TeilnehmerID=null)
  for (const item of allItems) {
    if (!targetIds.has(item.Id)) {
      targetIds.set(item.Id, null);
    }
  }

  let success = 0;
  let errors = 0;
  const totalItems = allItems.length || 1;
  let processed = 0;
  if (onProgress) { try { onProgress(0); } catch { /* */ } }
  for (const item of allItems) {
    const newId = targetIds.get(item.Id) ?? null;
    if (newId === item.TeilnehmerID) {
      success++;
    } else {
      try {
        const resp = await svc._merge(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${item.Id})`,
          { 'TeilnehmerID': newId }
        );
        if (resp.ok || resp.status === 406) { success++; } else { errors++; }
      } catch { errors++; }
    }
    processed++;
    if (onProgress) {
      // 0..95 % während der Merges; die letzten 5 % für syncCounterToMax.
      try { onProgress(Math.min(95, Math.round((processed / totalItems) * 95))); } catch { /* */ }
    }
  }

  // v7.31 / v9.14: Counter konsistent halten — syncCounterToMax patcht
  // Counter (monotonic up-only). ensureCounterList wurde hier ursprünglich
  // (v9.13) ebenfalls gerufen, hat aber Race-Conditions ausgelöst. Die
  // Counter-Liste sollte zum Zeitpunkt eines Reorders ohnehin existieren —
  // sonst hat die App ein anderes Problem das ein expliziter Klick auf
  // "Counter zurücksetzen" löst.
  try { await svc.syncCounterToMax(subsiteUrl); } catch { /* best-effort */ }
  if (onProgress) { try { onProgress(100); } catch { /* */ } }

  return { success, errors };
}

/**
 * Ungefilterte Gesamt-Item-Zahl einer Liste (List-Property `ItemCount`,
 * NICHT security-getrimmt). Dient als Vollständigkeits-Check: liefert eine
 * Item-Abfrage weniger Zeilen als `ItemCount`, beschneidet die
 * Element-Sicherheit („nur eigene Elemente") die Sicht des Aufrufers.
 * -1 bei Lesefehler.
 */
async function getListItemCount(svc: EventService, subsiteUrl: string, listName: string): Promise<number> {
  try {
    const resp = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ItemCount`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return -1;
    const data = await resp.json();
    const raw = data?.ItemCount ?? data?.d?.ItemCount;
    const n = typeof raw === 'number' ? raw : parseInt(String(raw ?? ''), 10);
    return Number.isFinite(n) && n >= 0 ? n : -1;
  } catch { return -1; }
}

/**
 * Zählt die aktiven (= nicht Warteliste/Abgemeldet) Anmeldungen, gesamt und
 * pro Starter-Gruppe. Quelle ist die echte Teilnehmerliste — wird zum Seeden
 * und Reconcilen der Sitzplatz-Counter genutzt.
 *
 * v27.10 REGRESSIONS-FIX: Seit v26.87 greift die Element-Sicherheit („nur
 * eigene Elemente") auf den Teilnehmerlisten WIRKLICH — vorher schlug das
 * Setzen still mit HTTP 400 fehl. Für normale User liefert
 * getAllRegistrations seitdem nur noch die EIGENEN Zeilen. Eine darauf
 * basierende Zählung wäre katastrophal zu niedrig: real beobachtet hat ein
 * Self-Cancel über syncSeatsToActiveCount `SeatsTaken=0` geschrieben,
 * wonach Neu-Anmeldungen an der kompletten Warteliste vorbei direkt
 * „Angemeldet" wurden. Deshalb wird die gelesene Zeilenzahl gegen den
 * ungefilterten `ItemCount` der Liste geprüft — ist die Sicht unvollständig
 * (oder nicht verifizierbar), fliegt ein Fehler und die Aufrufer handeln
 * fail-safe (Sync schreibt nichts, reserveSeat nutzt den reinen Counter).
 */
async function getActiveCounts(svc: EventService, subsiteUrl: string): Promise<{ total: number; durch: number; fun: number; waitlist: number }> {
  const regs = await svc.getAllRegistrations(subsiteUrl);
  const itemCount = await getListItemCount(svc, subsiteUrl, REG_LIST_NAME);
  if (itemCount < 0 || regs.length < itemCount) {
    throw new Error(`[DEX] getActiveCounts: Sicht unvollständig (${regs.length} von ${itemCount} Items lesbar) — Zählung unbrauchbar (Item-Level-Security).`);
  }
  const active = regs.filter(r => ACTIVE_STATI.indexOf(r.Status) >= 0);
  return {
    total: active.length,
    durch: active.filter(r => r.StarterType === 'Durchstarter').length,
    fun: active.filter(r => r.StarterType === 'Funstarter').length,
    // v24.73: Warteliste-Zahl für den (privilegierten) Counter-Reconcile.
    waitlist: regs.filter(r => r.Status === 'Warteliste').length,
  };
}

export function seatFieldFor(svc: EventService, group: string): 'SeatsTaken' | 'SeatsTakenDurch' | 'SeatsTakenFun' {
  if (group === 'Durchstarter') return 'SeatsTakenDurch';
  if (group === 'Funstarter') return 'SeatsTakenFun';
  return 'SeatsTaken';
}

/**
 * v11.36: Atomare Sitzplatz-Reservierung pro Gruppe via ETag-CAS auf der
 * DEX_TeilnehmerCounter-Liste — exakt dasselbe bewährte Muster wie
 * getNextTeilnehmerId (IF-MATCH, 412-Retry mit Backoff).
 *
 * Verhindert die Überbuchung bei zeitgleichen Anmeldungen: zwei parallele
 * Anmeldungen können nicht beide den letzten Platz greifen — die CAS
 * serialisiert das Increment, der Verlierer liest neu und sieht „voll".
 *
 * Rückgabe:
 * - 'reserved' → Platz wurde atomar belegt, Aufrufer darf 'Angemeldet' setzen
 * - 'full'     → Gruppe/Event ist voll → Aufrufer setzt 'Warteliste'
 * - 'error'    → Counter nicht nutzbar (Liste fehlt, Permission, Retries
 *                 erschöpft). Aufrufer MUSS fail-closed handeln (Warteliste),
 *                 NICHT optimistisch 'Angemeldet'.
 *
 * Self-Seed: ist das Seat-Feld noch nie gesetzt (null), wird es einmalig aus
 * der echten aktiven Anzahl der Gruppe initialisiert, bevor entschieden wird.
 */
export async function reserveSeat(
  svc: EventService,
  subsiteUrl: string,
  group: '' | 'Durchstarter' | 'Funstarter',
  cap: number,
  count: number = 1
): Promise<'reserved' | 'full' | 'error'> {
  // cap <= 0 = unbegrenzt → kein Reservieren nötig.
  if (!cap || cap <= 0) return 'reserved';
  const inc = Math.max(1, Math.floor(count));
  const field = svc.seatFieldFor(group);
  const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
  // v18.8 (Überbuchungs-Fix): Der Counter allein ist NICHT verlässlich.
  // Der Power-Automate-Nachrück-Flow promotet Warteliste→Angemeldet, ohne
  // SeatsTaken zu erhöhen; läuft der app-seitige syncSeatsToActiveCount
  // zeitlich VOR dieser asynchronen Promotion, steht der Counter unter dem
  // echten Aktiv-Bestand. Folge (real beobachtet): trotz voller Warteliste
  // sah der nächste Registrant einen Phantom-Platz und überbuchte. Deshalb
  // lesen wir EINMAL pro Aufruf den echten Aktiv-Bestand der Gruppe und
  // floor-en den Counter-Wert dagegen (max). Das schließt die Drift-Lücke,
  // erhält die atomare CAS-Serialisierung paralleler Anmeldungen UND heilt
  // den Counter nach oben. Bei Lesefehler (Throttling): kein Floor (-1) →
  // Fallback auf reines Counter-Verhalten, nicht schlechter als vorher.
  let realActive = -1;
  try {
    const rc = await getActiveCounts(svc, subsiteUrl);
    realActive = group === 'Durchstarter' ? rc.durch : group === 'Funstarter' ? rc.fun : rc.total;
  } catch { realActive = -1; }
  const MAX_RETRIES = 40;
  let triedLazyCreate = false;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    let getResp: SPHttpClientResponse;
    try {
      getResp = await svc._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
    } catch {
      return 'error';
    }
    if (!getResp.ok) {
      if (getResp.status === 404 && !triedLazyCreate) {
        triedLazyCreate = true;
        try { await svc.ensureCounterList(subsiteUrl); continue; } catch { return 'error'; }
      }
      return 'error';
    }
    const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
    if (!etag) return 'error';
    let data;
    try { data = await getResp.json(); } catch { return 'error'; }
    const rawVal = data?.[field] ?? data?.d?.[field];
    let current: number;
    if (rawVal === null || rawVal === undefined) {
      // Feld noch nie initialisiert → aus echtem Bestand seeden. v18.8:
      // den bereits oben gelesenen realActive wiederverwenden (kein zweiter
      // getActiveCounts-Roundtrip); nur falls der Read fehlschlug (-1),
      // konservativ auf 0.
      current = realActive >= 0 ? realActive : 0;
    } else {
      current = typeof rawVal === 'number' ? rawVal : (parseInt(String(rawVal), 10) || 0);
    }
    // v18.8: gegen echten Aktiv-Bestand floor-en (siehe Kommentar oben) —
    // fängt eine durch den Nachrück-Flow nach unten gedriftete Zählung ab.
    if (realActive >= 0 && realActive > current) current = realActive;
    // v11.82: Team-Anmeldungen reservieren N Plätze atomar. Wenn nicht alle
    // N in dieselbe Gruppe passen, schlägt die Reservierung als „full" fehl —
    // der Aufrufer setzt das gesamte Team auf Warteliste (kein Teil-Team
    // aktivieren). Bei count=1 (Solo) ist das Verhalten identisch zu vorher.
    if (current + inc > cap) return 'full';
    const patchResp = await svc._mergeIfMatch(counterItemUrl, { [field]: current + inc }, etag);
    if (patchResp.ok) return 'reserved';
    if (patchResp.status !== 412) return 'error';
    const baseDelay = Math.min(500, 50 * Math.pow(1.4, attempt));
    await new Promise(res => setTimeout(res, Math.floor(baseDelay * (0.5 + Math.random()))));
  }
  // Retries erschöpft → fail-closed (Aufrufer setzt Warteliste).
  console.warn('[DEX] reserveSeat: 40 Retries erschöpft — fail-closed (Warteliste).');
  return 'error';
}

/**
 * v11.36: Sitzplatz-Counter mit dem echten aktiven Bestand abgleichen.
 * Nach Abmeldungen, Reorder und Überbuchungs-Bereinigung aufrufen. Die
 * Power-Automate-Nachrück-Promotion fasst den Counter nicht an — dieser
 * Reconcile (aktive Anzahl aus der Liste) hält ihn ehrlich. Best-effort,
 * ETag-CAS, blockiert nie den aufrufenden Flow.
 *
 * v27.10: Liefert zurück, ob tatsächlich synchronisiert wurde. `false`
 * heißt insbesondere: der Aufrufer sieht die Teilnehmerliste nur beschnitten
 * (Item-Level-Security, siehe getActiveCounts) — dann wird bewusst NICHTS
 * geschrieben und der Aufrufer muss ggf. additiv am Counter arbeiten
 * (releaseSeatAfterCancel).
 */
export async function syncSeatsToActiveCount(
  svc: EventService,
  subsiteUrl: string,
  opts: { isSplit: boolean }
): Promise<boolean> {
  let counts: { total: number; durch: number; fun: number; waitlist: number };
  try { counts = await getActiveCounts(svc, subsiteUrl); } catch (err) {
    console.warn('[DEX] syncSeatsToActiveCount übersprungen:', err);
    return false;
  }
  // v24.76: WaitlistTaken-Feld sicherstellen, sonst HTTP 400 beim MERGE auf
  // Bestands-Events (Feld noch nicht angelegt).
  await svc.ensureCounterFieldsOnce(subsiteUrl);
  const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
  // v24.73: WaitlistTaken (rein informativ, nicht überbuchungs-relevant) beim
  // Reconcile mitschreiben — so heilt eine durch Flow-Promotion gedriftete
  // Warteliste-Zahl. Läuft NUR in privilegierten Kontexten (getActiveCounts
  // braucht Vollzugriff), daher hier korrekt; in User-Self-Cancel-Pfaden wird
  // syncSeatsToActiveCount bewusst nicht aufgerufen.
  const desired = opts.isSplit
    ? { SeatsTakenDurch: counts.durch, SeatsTakenFun: counts.fun, SeatsTaken: counts.total, WaitlistTaken: counts.waitlist }
    : { SeatsTaken: counts.total, WaitlistTaken: counts.waitlist };
  for (let attempt = 0; attempt < 6; attempt++) {
    try {
      const getResp = await svc._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
      if (!getResp.ok) return false;
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return false;
      const patchResp = await svc._mergeIfMatch(counterItemUrl, desired, etag);
      if (patchResp.ok) return true;
      if (patchResp.status !== 412) return false;
      await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
    } catch { return false; }
  }
  return false;
}

/**
 * v27.10: Ein einzelnes Sitzplatz-Counter-Feld additiv anpassen (atomar per
 * ETag-CAS, floor bei 0) — Pendant zu adjustWaitlistCounter für SeatsTaken/
 * SeatsTakenDurch/SeatsTakenFun. Für Aufrufer OHNE Vollzugriff auf die
 * Teilnehmerliste, die keine Absolutwerte schreiben dürfen (siehe
 * getActiveCounts). Best-effort, blockiert nie.
 */
export async function adjustSeatCounterField(
  svc: EventService,
  subsiteUrl: string,
  field: 'SeatsTaken' | 'SeatsTakenDurch' | 'SeatsTakenFun',
  delta: number
): Promise<void> {
  if (!subsiteUrl || !delta) return;
  const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
  for (let attempt = 0; attempt < 8; attempt++) {
    try {
      const getResp = await svc._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
      if (!getResp.ok) return;
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return;
      const data = await getResp.json();
      const rawVal = data?.[field] ?? data?.d?.[field];
      // Feld nie initialisiert → additiv sinnlos; der nächste privilegierte
      // Reconcile seedet es korrekt.
      if (rawVal === null || rawVal === undefined) return;
      const current = typeof rawVal === 'number' ? rawVal : (parseInt(String(rawVal), 10) || 0);
      const next = Math.max(0, current + delta);
      const patchResp = await svc._mergeIfMatch(counterItemUrl, { [field]: next }, etag);
      if (patchResp.ok) return;
      if (patchResp.status !== 412) return;
      await new Promise(res => setTimeout(res, 40 + Math.floor(Math.random() * 80)));
    } catch { return; }
  }
}

/**
 * v27.10 REGRESSIONS-FIX: Counter-Pflege nach einer Abmeldung — ersetzt den
 * direkten syncSeatsToActiveCount-Aufruf in den Abmelde-Pfaden (Selbst-
 * Abmeldung, Team-Lead-Abmeldung, Assistenz-Abmeldung), die auch von
 * NORMALEN Usern ausgelöst werden.
 *
 * Hintergrund: Seit die Element-Sicherheit („nur eigene Elemente", v26.87)
 * wirklich greift, sah ein normaler User beim Reconcile nur die eigene —
 * gerade abgemeldete — Zeile und schrieb SeatsTaken=0. Folge (real, 15.07.):
 * Neu-Anmeldungen wurden trotz voller Warteliste direkt „Angemeldet" und
 * überholten alle Wartenden.
 *
 * Ablauf jetzt:
 * 1. Voll-Reconcile versuchen (greift nur bei Vollzugriff, z.B. Organizer/
 *    Admin — exakteste Variante, heilt auch alte Drift).
 * 2. Sonst additiv (ILS-sicher):
 *    - Wartelisten-Zeile abgemeldet → WaitlistTaken −1.
 *    - Aktive Zeile abgemeldet → Platz nur freigeben (SeatsTaken −1), wenn
 *      die Warteliste laut Counter LEER ist. Steht jemand auf der Warteliste,
 *      besetzt der IDReorder-/Nachrück-Flow den Platz sofort FIFO-fair —
 *      SeatsTaken bleibt dann unverändert korrekt, und eine parallele
 *      Neu-Anmeldung kann die Wartenden nicht überholen. Bei unbekanntem
 *      Wartelisten-Stand: fail-closed nichts tun (privilegierter Reconcile
 *      heilt spätestens beim nächsten Admin-/Organizer-Boot).
 */
export async function releaseSeatAfterCancel(
  svc: EventService,
  subsiteUrl: string,
  opts: { isSplit: boolean; previousStatus: string; starterType?: string; waitlistDisabled?: boolean }
): Promise<void> {
  // v30.63: Jeder Zweig sagt in der Konsole, WAS er getan hat und warum.
  // Der Zähler wird seit v30.62 auf der Anmeldeseite angezeigt; wenn eine
  // Zahl nicht stimmt, muss sich die Ursache ohne Rätselraten ablesen lassen.
  // eslint-disable-next-line no-console
  const log = (msg: string): void => dlog('seats', `[DEX][seats] Abmeldung — ${msg}`);
  try {
    const synced = await svc.syncSeatsToActiveCount(subsiteUrl, { isSplit: opts.isSplit });
    if (synced) { log('Voll-Abgleich gelaufen (Vollzugriff) — Zähler exakt neu gesetzt.'); return; }
    if (opts.previousStatus === 'Warteliste') {
      await svc.adjustWaitlistCounter(subsiteUrl, -1);
      log('Wartelisten-Zeile abgemeldet → WaitlistTaken −1.');
      return;
    }
    if (ACTIVE_STATI.indexOf(opts.previousStatus) < 0) {
      log(`Vorheriger Status „${opts.previousStatus}" war nicht aktiv — kein Platz freizugeben.`);
      return;
    }
    // v27.11: Warteliste vom Organizer abgeschaltet → es rückt NIEMAND nach
    // (App-Gates + Flow-Bedingung). Der Platz muss dann direkt freigegeben
    // werden — sonst blieben frei gewordene Plätze dauerhaft als belegt
    // gezählt (Deadlock, bis ein privilegierter Reconcile läuft).
    if (opts.waitlistDisabled) {
      log('Warteliste ist abgeschaltet → SeatsTaken −1 (Platz sofort frei).');
      await svc.adjustSeatCounterField(subsiteUrl, 'SeatsTaken', -1);
      if (opts.isSplit && opts.starterType === 'Durchstarter') {
        await svc.adjustSeatCounterField(subsiteUrl, 'SeatsTakenDurch', -1);
      } else if (opts.isSplit && opts.starterType === 'Funstarter') {
        await svc.adjustSeatCounterField(subsiteUrl, 'SeatsTakenFun', -1);
      }
      return;
    }
    const stats = await svc.getCounterStats(subsiteUrl, opts.isSplit);
    // stats.waitlist: -1 = unbekannt (Feld nie gepflegt) → fail-closed.
    if (!stats) { log('Zähler nicht lesbar → nichts geändert (heilt beim nächsten Abgleich).'); return; }
    if (stats.waitlist < 0) {
      log('Wartelisten-Stand UNBEKANNT (Feld nie geschrieben) → nichts geändert. Seit v30.63 wird es bei neuen Events mit 0 angelegt; Bestands-Events heilt der Abgleich beim nächsten Organizer-/Admin-Start.');
      return;
    }
    if (stats.waitlist !== 0) {
      log(`${stats.waitlist} Person(en) auf der Warteliste → SeatsTaken bleibt (der Nachrück-Flow besetzt den Platz sofort).`);
      return;
    }
    log('Warteliste leer → SeatsTaken −1.');
    await svc.adjustSeatCounterField(subsiteUrl, 'SeatsTaken', -1);
    if (opts.isSplit && opts.starterType === 'Durchstarter') {
      await svc.adjustSeatCounterField(subsiteUrl, 'SeatsTakenDurch', -1);
    } else if (opts.isSplit && opts.starterType === 'Funstarter') {
      await svc.adjustSeatCounterField(subsiteUrl, 'SeatsTakenFun', -1);
    }
  } catch (err) {
    console.warn('[DEX] releaseSeatAfterCancel failed (best-effort):', err);
  }
}

/**
 * v24.73: Warteliste-Zähler im Counter additiv anpassen (atomar per ETag-CAS).
 * REIN INFORMATIV — `WaitlistTaken` gatet keine Überbuchung; ein verlorener
 * Bump verfälscht nur kurz die angezeigte Warteliste-Zahl und wird vom
 * privilegierten `syncSeatsToActiveCount`-Reconcile wieder geheilt. Wird von
 * den Anmelde-/Abmelde-Pfaden mit delta +1/-1 aufgerufen (auch von normalen
 * Usern — der Counter ist für alle schreibbar). Best-effort, blockiert nie.
 */
export async function adjustWaitlistCounter(svc: EventService, subsiteUrl: string, delta: number): Promise<void> {
  if (!subsiteUrl || !delta) return;
  // v24.76: Feld sicherstellen (greift für privilegierte Aufrufer; normale
  // User dürfen kein Feld anlegen → dann wird der Write unten übersprungen).
  await svc.ensureCounterFieldsOnce(subsiteUrl);
  const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
  for (let attempt = 0; attempt < 8; attempt++) {
    try {
      const getResp = await svc._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
      if (!getResp.ok) return;
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return;
      const data = await getResp.json();
      const body = data?.d ?? data;
      // v24.76: Feld (noch) nicht vorhanden → MERGE würde 400 liefern → still
      // überspringen (kein Konsolen-Fehler). Sobald ein Admin reconcilet, ist
      // das Feld da und der Bump greift.
      if (body && typeof body === 'object' && !('WaitlistTaken' in body)) return;
      const rawVal = data?.WaitlistTaken ?? data?.d?.WaitlistTaken;
      const current = typeof rawVal === 'number' ? rawVal : (parseInt(String(rawVal), 10) || 0);
      const next = Math.max(0, current + delta);
      const patchResp = await svc._mergeIfMatch(counterItemUrl, { WaitlistTaken: next }, etag);
      if (patchResp.ok) return;
      if (patchResp.status !== 412) return; // anderer Fehler → aufgeben (best-effort)
      await new Promise(res => setTimeout(res, 40 + Math.floor(Math.random() * 80)));
    } catch { return; }
  }
}

/**
 * v24.73: Live-Plätze aus dem Counter lesen — für ALLE lesbar (auch normale
 * Teilnehmer, im Gegensatz zur item-level-gesicherten Teilnehmerliste). Quelle
 * der Anzeige-Zahlen (aktiv = SeatsTaken, Warteliste = WaitlistTaken). Liefert
 * `null`, wenn der Counter (noch) nicht existiert/lesbar ist → Aufrufer fällt
 * dann auf den bisherigen (item-level-gefilterten) Zählweg zurück.
 */
// v30.67: durch/fun/groupsKnown zusaetzlich — die Anmeldeseite braucht bei
// geteilten Kapazitaeten die Gruppenwerte aus dem Zaehler, weil die
// Teilnehmerliste zeilenweise gesichert ist und fuer Teilnehmer nur die eigene
// Zeile zeigt (jede Zahl daraus waere erfunden). reserveSeat pflegt bei Split
// NUR die Gruppenfelder, deshalb sind sie hier die verlaessliche Quelle.
export interface CounterStats { active: number; waitlist: number; seatsKnown: boolean; durch: number; fun: number; groupsKnown: boolean }
export async function getCounterStats(svc: EventService, subsiteUrl: string, isSplit: boolean): Promise<CounterStats | null> {
  if (!subsiteUrl) return null;
  const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
  try {
    const resp = await svc._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
    if (!resp.ok) return null;
    const data = await resp.json();
    const num = (v: unknown): number => typeof v === 'number' ? v : (parseInt(String(v ?? ''), 10) || 0);
    const total = num(data?.SeatsTaken ?? data?.d?.SeatsTaken);
    const durch = num(data?.SeatsTakenDurch ?? data?.d?.SeatsTakenDurch);
    const fun = num(data?.SeatsTakenFun ?? data?.d?.SeatsTakenFun);
    const wRaw = data?.WaitlistTaken ?? data?.d?.WaitlistTaken;
    // v30.62: Ist das Feld noch NIE gesetzt worden (null/undefined), liefert
    // `num()` eine 0 — und eine 0 liest sich wie „niemand angemeldet". Für die
    // Termin-Kacheln der Anmeldeseite ist das der Unterschied zwischen einer
    // Aussage und einer Erfindung, deshalb wird er hier durchgereicht.
    const sRaw = data?.SeatsTaken ?? data?.d?.SeatsTaken;
    const sDurchRaw = data?.SeatsTakenDurch ?? data?.d?.SeatsTakenDurch;
    const sFunRaw = data?.SeatsTakenFun ?? data?.d?.SeatsTakenFun;
    const known = (v: unknown): boolean => v !== null && v !== undefined && v !== '';
    const seatsKnown = known(sRaw) || (isSplit && (known(sDurchRaw) || known(sFunRaw)));
    // SeatsTaken ist der Gesamt-Aktiv-Wert; bei Split fällt er ggf. auf
    // Durch+Fun zurück, falls der Gesamtwert (noch) nicht gepflegt wurde.
    const active = total > 0 ? total : (isSplit ? durch + fun : total);
    const waitlist = (wRaw === null || wRaw === undefined) ? -1 : num(wRaw);
    const groupsKnown = known(sDurchRaw) || known(sFunRaw);
    return { active, waitlist, seatsKnown, durch, fun, groupsKnown };
  } catch { return null; }
}

/**
 * v24.75: Echtzeit-Push auf eine Liste der Event-Subsite abonnieren.
 * kind='counter' → DEX_TeilnehmerCounter (für alle lesbar; Anmeldeformular),
 * kind='participants' → Teilnehmerliste (Organizer-Vollzugriff). Liefert eine
 * Cleanup-Funktion. Best-effort (siehe utils/spListRealtime).
 */
export async function subscribeListRealtime(
  svc: EventService,
  subsiteUrl: string,
  kind: 'counter' | 'participants',
  onChange: () => void
): Promise<() => void> {
  if (!subsiteUrl) return () => { /* */ };
  const listTitle = kind === 'counter' ? COUNTER_LIST_NAME : REG_LIST_NAME;
  try { return await subscribeListChanges(svc.context.spHttpClient, subsiteUrl, listTitle, onChange); }
  catch { return () => { /* */ }; }
}
