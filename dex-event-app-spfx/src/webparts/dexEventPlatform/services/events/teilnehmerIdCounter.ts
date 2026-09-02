/**
 * v30.66 — Modularisierung Stufe 2: Thema „TeilnehmerID-Zähler".
 *
 * Pro Subsite eine Liste DEX_TeilnehmerCounter mit genau EINEM Item, dessen
 * NextValue per If-Match-Header (ETag) hochgezählt wird — deshalb können
 * mehrere Personen gleichzeitig anmelden, ohne dieselbe ID zu bekommen.
 * Die Warteliste hat keine Positionsspalte: Die Position IST der Rang nach
 * TeilnehmerID (siehe CLAUDE.md), weshalb an diesen Zahlen mehr hängt als
 * eine hübsche Nummer.
 *
 * Der Sitzungs-Merker `_counterFieldsEnsured` ist Instanz-Zustand und bleibt
 * an der Klasse. Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { EventService } from '../EventService';
import { COUNTER_LIST_ITEM_TYPE, COUNTER_LIST_NAME, REG_LIST_NAME } from '../EventService';

// v7.28: Atomare TeilnehmerID-Vergabe über die DEX_TeilnehmerCounter-Liste
// pro Event-Subsite. Verhindert Race-Conditions wenn viele User gleichzeitig
// anmelden — ohne das vorher passieren konnte, dass zwei User dieselbe ID
// bekommen (siehe Bug-Report v7.27 → v7.28).
//
// Ablauf:
//   1. Counter-Item GET'en, ETag aus Response-Header lesen.
//   2. NextValue + 1 mit IF-MATCH: <etag> via MERGE schreiben.
//   3. Bei 412 (ETag-Mismatch = jemand war schneller) → kurzes Jitter +
//      Retry, max 8x.
//   4. Bei Erfolg: NextValue+1 ist die neue TeilnehmerID des aufrufenden Users.
//
// Fallback: Wenn die Counter-Liste nicht existiert (z.B. legacy event ohne
// "Spalten fixen"-Lauf), kommt undefined zurück — der Aufrufer fällt dann
// auf das alte (race-anfällige) max+1-Verfahren zurück. Bestandsschutz.
// v7.28 / v9.10: Nächste TeilnehmerID atomar holen.
//   1. Counter-Item GET'en, ETag aus Response-Header lesen.
//   2. Counter-Item PATCH'en mit IF-MATCH=<ETag>, NextValue=current+1.
//   3. Bei 412 (ETag-Mismatch = jemand war schneller) → Exponential
//      Backoff mit Full Jitter, dann Retry. Bis zu 40 Versuche.
//   4. Bei Erfolg: NextValue+1 ist die neue TeilnehmerID des aufrufenden Users.
//
// v9.10: Counter-Liste wird ON-DEMAND angelegt+geseeded, falls sie fehlt
// (z.B. weil das Event vor v7.28 erstellt wurde). Vorher gab undefined
// zurück → Aufrufer fiel auf max+1 zurück → Race-Condition bei
// Massen-Anmeldungen. Jetzt: einmalig ensureCounterList() rufen, dann
// erneut versuchen. Damit ist der race-anfällige Fallback nur noch
// erreicht, wenn auch das Anlegen scheitert (Permission-Issue).
//
// v9.10: Retries 8 → 40, Backoff von festem Jitter auf Exponential
// Backoff mit Full Jitter (Cap 500ms). Bei 50+ parallelen Anmeldungen
// wahren 8 Retries praktisch garantiert ausgeschöpft.
export async function getNextTeilnehmerId(svc: EventService, subsiteUrl: string): Promise<number | undefined> {
  const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
  const MAX_RETRIES = 40;
  let triedLazyCreate = false;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    let getResp: SPHttpClientResponse;
    try {
      getResp = await svc._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
    } catch {
      return undefined;
    }
    if (!getResp.ok) {
      // 404 = Counter-Liste / Item existiert nicht.
      // v9.10: Statt direkt undefined zu liefern, einmalig versuchen die
      // Liste anzulegen + zu seeden (idempotent). Wenn das klappt, gleich
      // weiter — wenn nicht, geben wir auf.
      if (getResp.status === 404 && !triedLazyCreate) {
        triedLazyCreate = true;
        try {
          await svc.ensureCounterList(subsiteUrl);
          // Kein delay — direkt nächste Iteration, die das frische Item liest.
          continue;
        } catch {
          return undefined;
        }
      }
      return undefined;
    }
    const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
    if (!etag) return undefined;
    let data;
    try { data = await getResp.json(); } catch { return undefined; }
    // v9.13: NextValue defensiv parsen — handhabt sowohl number als auch
    // (in seltenen SP-Konfigurationen) string. Sorgt dafür dass current
    // bei einer korrekt gespeicherten 165 nie auf 0 fallback'ed.
    const rawNextValue = data?.NextValue ?? data?.d?.NextValue;
    const current = typeof rawNextValue === 'number'
      ? rawNextValue
      : (typeof rawNextValue === 'string' ? (parseInt(rawNextValue, 10) || 0) : 0);
    // v9.13: Counter ist die Source of Truth. Wir vertrauen ihm und
    // inkrementieren atomar via ETag-CAS. KEIN zusätzlicher Lesezugriff
    // auf die Teilnehmerliste mehr — das Counter-Pattern existiert genau,
    // damit wir hier NICHT max(TID) aus der Teilnehmerliste rechnen
    // müssen. Wenn der Counter korrupt sein sollte (z.B. von altem
    // syncCounterToMax-Bug auf 0 gepatcht), gibt's den expliziten
    // "Counter zurücksetzen"-Button im Admin Center für den Fix.
    const next = current + 1;
    const patchResp = await svc._mergeIfMatch(counterItemUrl, { 'NextValue': next }, etag);
    if (patchResp.ok) return next;
    if (patchResp.status !== 412) {
      // Anderer Fehler (z.B. 500) → kein Sinn weiter zu retry'n
      return undefined;
    }
    // 412 = ETag-Mismatch = jemand war schneller → Exponential Backoff
    // mit Full Jitter (Cap 500ms). Cluster bei Massen-Anmeldungen
    // werden so zuverlässig entzerrt — ohne Backoff laufen alle
    // Clients sekundengleich in den nächsten Conflict.
    const baseDelay = Math.min(500, 50 * Math.pow(1.4, attempt));
    const delay = Math.floor(baseDelay * (0.5 + Math.random()));
    await new Promise(res => setTimeout(res, delay));
  }
  // Nach 40 Retries aufgeben — Aufrufer kann die Anmeldung sauber
  // mit TeilnehmerID=null durchziehen lassen und der Admin lädt
  // anschliessend "IDs neu vergeben".
  console.warn('[DEX] getNextTeilnehmerId: 40 retries erschöpft — TeilnehmerID bleibt unset, Admin sollte IDs neu vergeben.');
  return undefined;
}

// v24.76: Counter-Felder EINMAL pro Subsite und Sitzung sicherstellen, bevor
// darauf geschrieben wird — sonst liefert der MERGE auf Bestands-Events ein
// HTTP 400 (Feld existiert noch nicht). Der Merker dazu liegt an der Instanz.
export async function ensureCounterFieldsOnce(svc: EventService, subsiteUrl: string): Promise<void> {
  if (!subsiteUrl || svc._counterFieldsEnsured.has(subsiteUrl)) return;
  try { await svc.ensureCounterListField(subsiteUrl); } catch { /* best-effort */ }
  svc._counterFieldsEnsured.add(subsiteUrl);
}

// Hilfsroutine: prüft ob auf der Counter-Liste die NextValue-Spalte
// existiert und legt sie an wenn sie fehlt. Idempotent.
export async function ensureCounterListField(svc: EventService, subsiteUrl: string): Promise<void> {
  const fieldsUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/fields`;
  // v7.28: NextValue (TeilnehmerID-Automat).
  // v11.36: SeatsTaken / SeatsTakenDurch / SeatsTakenFun — atomare
  // Sitzplatz-Reservierung pro Gruppe (gegen Überbuchung bei
  // zeitgleichen Anmeldungen). Alle Number-Felder, default 0/leer.
  const wanted = ['NextValue', 'SeatsTaken', 'SeatsTakenDurch', 'SeatsTakenFun', 'WaitlistTaken'];
  for (const name of wanted) {
    try {
      const probe = await svc._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/fields/getbytitle('${name}')`,
        SPHttpClient.configurations.v1
      );
      if (probe.ok) continue;
      await svc._post(
        fieldsUrl,
        { '__metadata': { 'type': 'SP.Field' }, 'Title': name, 'FieldTypeKind': 9, 'Required': false }
      );
    } catch { /* best-effort; reserveSeat fällt sonst sauber zurück */ }
  }
}

// v7.28: Counter-Liste für ein Event anlegen (1 Liste mit 1 Item) und
// direkt mit dem aktuellen Max-Wert seeden — damit bestehende Events ohne
// ID-Lückenproduktion umsteigen können.
// Idempotent: tut nichts wenn die Liste schon existiert.
// v7.29-Fix: Item-Inserts nutzen den korrekt _x005f_-encodeten Type-Namen
// (genauso wie wir das für DEX_Events machen). Vorher wurde der Listen-
// name 1:1 in den Type übernommen, was bei Unterstrich stillschweigend
// zu HTTP 400 führt → leere Counter-Liste.
export async function ensureCounterList(svc: EventService, subsiteUrl: string): Promise<{ created: boolean; seededValue?: number }> {
  const probe = await svc._sp.get(
    `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')`,
    SPHttpClient.configurations.v1
  );
  const itemsUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items`;
  const seedItem = async (): Promise<number> => {
    const maxId = await svc.getCurrentMaxTeilnehmerId(subsiteUrl);
    const resp = await svc._post(itemsUrl, {
      '__metadata': { 'type': COUNTER_LIST_ITEM_TYPE },
      'Title': 'TeilnehmerID',
      'NextValue': maxId,
      // v30.63: Sitzplatz- und Wartelisten-Zähler AUSDRÜCKLICH auf 0 setzen.
      //
      // Bisher blieben beide Felder null, bis die erste Anmeldung sie anfasst
      // — und `WaitlistTaken` fasst niemand an, solange nie jemand auf der
      // Warteliste steht. `getCounterStats` meldet null aber als „unbekannt"
      // (-1), und `releaseSeatAfterCancel` handelt bei unbekanntem
      // Wartelisten-Stand fail-closed: Es zählt NICHT herunter. Ergebnis war
      // ein Zähler, der bei jeder Selbst-Abmeldung zu hoch stehen blieb, bis
      // ein Admin die App öffnete.
      //
      // Auf einer frischen Subsite ist die 0 keine Annahme, sondern die
      // Tatsache: Es gibt noch keine Anmeldung. Genau deshalb darf sie hier
      // geschrieben werden — und nur hier.
      'SeatsTaken': 0,
      'WaitlistTaken': 0,
    });
    if (!resp.ok) {
      const errBody = await resp.text().catch(() => '');
      console.warn('[DEX] ensureCounterList: Seed-Item fehlgeschlagen, status=', resp.status, errBody.substring(0, 300));
    }
    return maxId;
  };

  if (probe.ok) {
    // Liste existiert — sicherstellen dass das Schema komplett ist und ein Item drin liegt.
    try { await svc.ensureCounterListField(subsiteUrl); } catch { /* */ }
    // v9.13/v9.14: setCounterListPermissions wurde hier ursprünglich
    // mitaufgerufen, hat aber bei laufender Event-Anlage Race-Conditions
    // ausgelöst (breakroleinheritance gegen frisch provisionierte Liste
    // in derselben Request-Welle). Permissions werden jetzt nur noch
    // explizit über den "Counter zurücksetzen"-Button gefixt — siehe
    // resetCounterToMax. Bestehende Events können damit per Admin-Klick
    // geheilt werden, neue Events bekommen ihre Permissions im
    // create-Branch unten gesetzt.
    const itemListResp = await svc._sp.get(
      `${itemsUrl}?$top=1`,
      SPHttpClient.configurations.v1
    );
    if (itemListResp.ok) {
      const data = await itemListResp.json();
      const list = data.value || data.d?.results || [];
      if (list.length > 0) return { created: false }; // alles ok
    }
    // Liste ohne Item → nachseeden
    const seededValue = await seedItem();
    return { created: false, seededValue };
  }

  // Liste neu anlegen
  await svc._post(
    `${subsiteUrl}/_api/web/lists`,
    {
      '__metadata': { 'type': 'SP.List' },
      'BaseTemplate': 100,
      'Title': COUNTER_LIST_NAME,
      'Description': 'Atomarer Counter für TeilnehmerID-Vergabe (ETag-basiert). Nicht manuell editieren.',
      'AllowContentTypes': false,
      'ContentTypesEnabled': false,
      'EnableVersioning': false,
      'EnableMinorVersions': false,
      'OnQuickLaunch': false,
    }
  );
  await svc.ensureCounterListField(subsiteUrl);
  // v9.13: Counter-Liste muss explizit Contribute-Rechte für Visitors
  // bekommen, damit normale User die ETag-CAS-Inkrementierung
  // durchführen können. Ohne das schlägt PATCH NextValue mit 401/403
  // fehl → getNextTeilnehmerId gibt undefined zurück → TID landet null
  // (oder im allerersten Lazy-Create-Pfad bei 1).
  try { await svc.setCounterListPermissions(subsiteUrl); } catch { /* */ }
  const seededValue = await seedItem();
  return { created: true, seededValue };
}

/**
 * Berechtigungen für DEX_TeilnehmerCounter setzen — analog zur
 * Teilnehmerliste:
 *   - Owners der Hauptsite: Full Control (1073741829)
 *   - Visitors (DEALL): Contribute (1073741827) → ETag-CAS-Inkrement
 *   - Organizer-Mail (falls bekannt): Full Control
 *
 * Idempotent: kann auf bestehenden Counter-Listen erneut aufgerufen
 * werden um v9.13-Permissions nachzupatchen. Die Funktion bricht
 * Rollen-Vererbung explizit (clearSubscopes=true), damit Read-Only-
 * Inheritance vom Subsite nicht versehentlich greift.
 */
export async function setCounterListPermissions(svc: EventService, subsiteUrl: string, organizerEmail?: string): Promise<void> {
  try {
    await svc._post(
      `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
      {}
    );
    const ownersResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/associatedownergroup?$select=Id`,
      SPHttpClient.configurations.v1
    );
    if (ownersResp.ok) {
      const ownersData = await ownersResp.json();
      const ownersId = ownersData.Id ?? ownersData.d?.Id;
      if (ownersId) {
        await svc._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/roleassignments/addroleassignment(principalid=${ownersId}, roledefid=1073741829)`,
          {}
        );
      }
    }
    // Visitors → Contribute. KRITISCH: damit normale User
    // den Counter atomar inkrementieren können.
    const visitorsId = await svc.getVisitorsGroupId();
    if (visitorsId) {
      await svc._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741827)`,
        {}
      );
    }
    // Organizer optional → Full Control
    if (organizerEmail) {
      try {
        const userResp = await svc._sp.get(
          `${svc.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(organizerEmail)}')?$select=Id`,
          SPHttpClient.configurations.v1
        );
        if (userResp.ok) {
          const userData = await userResp.json();
          const userId = userData.Id ?? userData.d?.Id;
          if (userId) {
            await svc._post(
              `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
              {}
            );
          }
        }
      } catch { /* Organizer-Permission ist optional */ }
    }
  } catch (err) {
    console.warn('[DEX] setCounterListPermissions fehlgeschlagen:', err);
  }
}

// v7.28 / v9.13: Aktuellen Max-Wert von TeilnehmerID in der Teilnehmerliste
// lesen. Wird beim **Seeden** des Counters und beim **Sync nach Reorder**
// genutzt — die Counter-Liste selbst ist im Normalbetrieb die Source of
// Truth (siehe getNextTeilnehmerId).
//
// **Bugfix v9.13:** Vorher hat $orderby=TeilnehmerID desc&$top=1 unter
// bestimmten Bedingungen das null-Item zuerst geliefert (SP sortiert NULL-
// Werte bei Number-Feldern oft als "größter Wert" in desc-Order).
// Sobald irgendjemand abgemeldet war (TID=null) lief die Funktion ins
// null-Branch und gab 0 zurück.
//
// Konsequenz im alten Code (vor v9.12): syncCounterToMax patcht den
// Counter auf liveMax=0 RUNTER → nächste Anmeldung kriegt TID=1 →
// Duplikat zu echten aktiven Teilnehmern. Genau der Fall den der User
// beim Go-Live live gesehen hat (Theresa #1 obwohl 165 aktive Anmeldungen).
//
// Fix: $filter=TeilnehmerID gt 0 schliesst NULL und 0 explizit aus —
// funktioniert unabhängig von SP-NULL-Sortier-Konventionen.
export async function getCurrentMaxTeilnehmerId(svc: EventService, subsiteUrl: string): Promise<number> {
  try {
    const resp = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=TeilnehmerID&$filter=TeilnehmerID gt 0&$orderby=TeilnehmerID desc&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return 0;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    if (items.length > 0 && items[0].TeilnehmerID != null) return items[0].TeilnehmerID;
  } catch { /* */ }
  return 0;
}

// v7.31 / v9.12: Counter mit aktuellem Max-Wert konsistent halten — wird
// nach Cancel und nach reorderParticipantIDs aufgerufen. Die Logik ist
// **monotonic up-only**: der Counter geht NIE runter (sonst werden
// cancelled IDs reused — exakter Duplikat-Bug aus dem Go-Live).
//
// - current >= liveMax: nichts zu tun (Counter steht bereits hoch genug).
// - current  < liveMax: Counter HOCH auf liveMax setzen (z.B. nach
//   reorderParticipantIDs, der die TIDs im Bereich [1..N_active] vergibt
//   wo N_active größer sein kann als der bisherige Counter-Stand).
//
// Vorher (Bug bis v9.11): exakt umgekehrt — Counter wurde auf liveMax
// RUNTER gesetzt wenn current > liveMax. Das produzierte sowohl bei
// Cancel als auch nach IDReorder Duplikate.
//
// ETag-CAS mit Retry, damit eine parallele Anmeldung den Counter nicht
// zwischen Read und Write wegrasselt.
// v9.13: Oeffentliche Recovery-Methode für den Admin-Button "Counter
// zurücksetzen". Liest den aktuellen Max-TID aus der Teilnehmerliste und
// setzt den Counter auf diesen Wert (per ETag-CAS, monotonic up-only via
// syncCounterToMax). Gibt den neuen Counter-Wert zurück damit der Admin
// direkt sehen kann auf was es gepatcht wurde.
export async function resetCounterToMax(svc: EventService, subsiteUrl: string): Promise<{ counter: number; max: number }> {
  // v11.27: bidirektionaler Reset. Vorher rief diese Methode nur
  // syncCounterToMax auf — das ist monotonic up-only und liess einen
  // zu hohen Counter unverändert. Genau das hat der Maintainer beobachtet:
  // Counter=11, Max-TID=4, Klick auf "Counter zurücksetzen" → keine
  // Änderung, weiterhin 11. Jetzt setzen wir den Counter explizit
  // auf max(TID) — egal ob er drunter (gefährlich, Doppel-IDs möglich)
  // oder drüber stand (harmlos, nur Lücken-Springen).
  try { await svc.ensureCounterList(subsiteUrl); } catch { /* */ }
  const max = await svc.getCurrentMaxTeilnehmerId(subsiteUrl);
  const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
  let finalCounter = 0;
  // ETag-CAS-Loop, falls jemand parallel inserted und den Counter inkrementiert.
  for (let attempt = 0; attempt < 8; attempt++) {
    try {
      const getResp = await svc._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
      if (!getResp.ok) break;
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) break;
      const data = await getResp.json();
      const rawCurrent = data?.NextValue ?? data?.d?.NextValue;
      const current = typeof rawCurrent === 'number' ? rawCurrent : (typeof rawCurrent === 'string' ? (parseInt(rawCurrent, 10) || 0) : 0);
      if (current === max) {
        finalCounter = current;
        break;
      }
      const patchResp = await svc._mergeIfMatch(counterItemUrl, { 'NextValue': max }, etag);
      if (patchResp.ok) {
        finalCounter = max;
        console.warn(`[DEX] resetCounterToMax: counter von ${current} auf ${max} gesetzt (Subsite: ${subsiteUrl}).`);
        break;
      }
      if (patchResp.status !== 412) {
        finalCounter = current;
        break;
      }
      // 412 = jemand war schneller, nochmal lesen+patchen
      await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
    } catch (err) {
      console.warn('[DEX] resetCounterToMax error:', err);
      break;
    }
  }
  return { counter: finalCounter, max };
}

export async function syncCounterToMax(svc: EventService, subsiteUrl: string): Promise<void> {
  const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
  for (let attempt = 0; attempt < 8; attempt++) {
    const liveMax = await svc.getCurrentMaxTeilnehmerId(subsiteUrl);
    let getResp: SPHttpClientResponse;
    try {
      getResp = await svc._sp.get(counterItemUrl, SPHttpClient.configurations.v1);
    } catch {
      return;
    }
    if (!getResp.ok) return;
    const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
    if (!etag) return;
    let data;
    try { data = await getResp.json(); } catch { return; }
    const current = typeof data?.NextValue === 'number' ? data.NextValue : 0;
    if (current >= liveMax) return; // bereits konsistent — Counter ist nicht "zu klein"
    const patchResp = await svc._mergeIfMatch(counterItemUrl, { 'NextValue': liveMax }, etag);
    if (patchResp.ok) {
      console.warn(`[DEX] syncCounterToMax: counter von ${current} auf ${liveMax} hochgezogen.`);
      return;
    }
    if (patchResp.status !== 412) return;
    // 412 = jemand war schneller, nochmal lesen+patchen
    await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
  }
  // Nach 8 Retries aufgeben — best-effort, blockiert keine andere Aktion
}
