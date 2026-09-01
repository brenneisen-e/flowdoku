/**
 * v30.66 — Modularisierung Stufe 2: Thema „Status einer Anmeldung":
 * abmelden, absagen, einchecken, No-Show, auschecken, QR-Versand vermerken,
 * Gruppe wechseln — plus die Zählungen, die daraus folgen.
 *
 * Abmelden ist der einzige Weg, an dem das Nachrücken hängt
 * (`promoteFirstWaitlistItem`): Eine Kapazitätsänderung erzeugt kein
 * Ereignis, an dem etwas hinge (siehe CLAUDE.md).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService, SPRegistration } from '../EventService';
import { ACTIVE_STATI, REG_LIST_ITEM_TYPE, REG_LIST_NAME } from '../EventService';

/**
 * v10.27: User wechselt seine Split-Capacity-Gruppe.
 *
 * Logik:
 * - Lädt aktuelle aktive Registrierungen der Subsite und zählt die Anzahl
 *   pro StarterType.
 * - Wenn die Ziel-Gruppe noch unter ihrer Kapazität liegt: User wird mit
 *   neuem StarterType direkt als 'Angemeldet' eingetragen.
 * - Wenn die Ziel-Gruppe bereits voll ist: User wandert auf die Warteliste
 *   mit PreferredStarterType=newType. StarterType bleibt leer (wie bei
 *   Erst-Anmeldung auf Warteliste). Nachgerückt wird er erst, wenn ein
 *   Platz in der Ziel-Gruppe frei wird (siehe Power-Automate-Flow).
 *
 * Liefert { ok, status, full } zurück — die App nutzt das, um die richtige
 * Mail (Anmeldung vs. Warteliste) zu queuen und dem User Feedback zu geben.
 */
/**
 * v11.24: Tauscht StarterType (und PreferredStarterType) bei ALLEN
 * Registrierungen einer Subsite: jeder 'Durchstarter' wird zu
 * 'Funstarter' und umgekehrt. Wird vom Admin-Center aufgerufen, wenn
 * der Organizer im Wizard die Reihenfolge der Gruppen-Labels +
 * -Kapazitäten getauscht hat — die existierenden Anmeldungen sind
 * dann technisch noch in der „alten" Slot-Bedeutung. Dieser Flip
 * synchronisiert sie mit der neuen Reihenfolge.
 *
 * Liefert die Anzahl erfolgreich aktualisierter Items zurück.
 */
export async function flipAllStarterTypes(svc: EventService, subsiteUrl: string): Promise<{ ok: boolean; updated: number; failed: number }> {
  try {
    const all = await svc.getAllRegistrations(subsiteUrl);
    let updated = 0;
    let failed = 0;
    for (const r of all) {
      const flip = (t: string | undefined): string => {
        if (t === 'Durchstarter') return 'Funstarter';
        if (t === 'Funstarter') return 'Durchstarter';
        return t || '';
      };
      const newStarter = flip(r.StarterType);
      const newPref = flip(r.PreferredStarterType);
      if (newStarter === (r.StarterType || '') && newPref === (r.PreferredStarterType || '')) continue;
      try {
        const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${r.Id})`;
        const body: Record<string, unknown> = {};
        if (newStarter !== (r.StarterType || '')) body['StarterType'] = newStarter;
        if (newPref !== (r.PreferredStarterType || '')) body['PreferredStarterType'] = newPref;
        const resp = await svc._merge(url, body);
        if (resp.ok) updated++;
        else failed++;
      } catch { failed++; }
    }
    return { ok: failed === 0, updated, failed };
  } catch (err) {
    console.warn('[DEX] flipAllStarterTypes error:', err);
    return { ok: false, updated: 0, failed: 0 };
  }
}

export async function switchSplitGroup(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  newType: 'Durchstarter' | 'Funstarter',
  durchstarterCapacity: number,
  funstarterCapacity: number,
): Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' | 'Failed'; full: boolean }> {
  try {
    // v27.10 REGRESSIONS-FIX: Die frühere Zählung über getAllRegistrations
    // ist für normale User unbrauchbar, seit die Element-Sicherheit („nur
    // eigene Elemente", v26.87) wirklich greift — sie sahen nur die eigene
    // Zeile, zählten die Ziel-Gruppe als leer und wechselten an einer
    // vollen Gruppe (inkl. deren Warteliste) vorbei direkt auf „Angemeldet".
    // Stattdessen jetzt die ATOMARE Sitzplatz-Reservierung über den für
    // alle lesbaren/schreibbaren Gruppen-Counter (gleicher Mechanismus wie
    // bei der Neu-Anmeldung, ETag-CAS): 'reserved' → Platz sicher belegt,
    // 'full'/'error' → fail-closed Warteliste. Der eigene Eintrag steckt
    // bei einem echten Wechsel nie im Ziel-Gruppen-Zähler (aktiv in der
    // ANDEREN Gruppe oder auf der Warteliste).
    // Vorherigen Zustand der eigenen Zeile lesen (unter Item-Level-Security
    // immer sichtbar) — für die additive Counter-Pflege unten.
    let prevStatus = '';
    try {
      const ownResp = await svc._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=Status`,
        SPHttpClient.configurations.v1
      );
      if (ownResp.ok) {
        const ownData = await ownResp.json();
        const own = ownData?.d ?? ownData;
        prevStatus = own?.Status || '';
      }
    } catch { /* best-effort — Counter-Pflege unten fällt dann konservativ aus */ }
    const wasActive = ACTIVE_STATI.indexOf(prevStatus) >= 0;
    const wasWaitlist = prevStatus === 'Warteliste';
    const targetCap = newType === 'Durchstarter' ? durchstarterCapacity : funstarterCapacity;
    const seat = targetCap > 0
      ? await svc.reserveSeat(subsiteUrl, newType, targetCap)
      : 'reserved'; // cap <= 0 = unbegrenzt
    const goWaitlist = seat !== 'reserved';
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: Record<string, any> = goWaitlist
      ? { 'Status': 'Warteliste', 'StarterType': '', 'PreferredStarterType': newType }
      : { 'Status': 'Angemeldet', 'StarterType': newType, 'PreferredStarterType': newType };
    // v22.20: FIFO-Fairness — wer auf die Warteliste der Zielgruppe wechselt,
    // verliert seine alte (niedrige) TeilnehmerID und reiht sich mit einer
    // frischen Counter-ID HINTEN ein. Sonst würde er beim typ-bewussten
    // Nachrücken (TeilnehmerID asc) alle überholen, die schon länger warten.
    // Best-effort: schlägt der Counter fehl, bleibt die alte ID stehen und der
    // anschließende Reorder-Lauf normalisiert wenigstens die Nummerierung.
    if (goWaitlist) {
      try {
        const freshId = await svc.getNextTeilnehmerId(subsiteUrl);
        if (typeof freshId === 'number' && freshId > 0) body['TeilnehmerID'] = freshId;
      } catch { /* alte ID behalten */ }
    }
    const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`;
    const resp = await svc._merge(url, body);
    if (!resp.ok) {
      console.warn('[DEX] switchSplitGroup MERGE failed:', resp.status);
      // Reservierten Ziel-Gruppen-Platz zurückgeben — die Zeile wurde ja
      // nicht umgestellt.
      if (!goWaitlist && targetCap > 0) {
        await svc.adjustSeatCounterField(subsiteUrl, svc.seatFieldFor(newType), -1);
      }
      return { ok: false, status: 'Failed', full: goWaitlist };
    }
    // v27.10: Additive Counter-Pflege (ILS-sicher, best-effort). Nur die
    // eindeutig sicheren Anpassungen — alles Unklare heilt der nächste
    // privilegierte Reconcile:
    // - Aktiv → Warteliste: WaitlistTaken +1 (der frei werdende Quell-Slot
    //   wird bewusst NICHT dekrementiert — fail-closed, Nachrücken/Reconcile
    //   übernimmt).
    // - Warteliste → Aktiv: WaitlistTaken −1 (Ziel-Gruppen-Zähler hat
    //   reserveSeat bereits atomar erhöht).
    if (goWaitlist && wasActive) {
      await svc.adjustWaitlistCounter(subsiteUrl, +1);
    } else if (!goWaitlist && wasWaitlist) {
      await svc.adjustWaitlistCounter(subsiteUrl, -1);
    }
    // Aktiv → Aktiv (Gruppenwechsel): Quell-Gruppen-Zähler bleibt bewusst
    // stehen (fail-closed) — der nächste privilegierte Reconcile setzt ihn
    // exakt.
    return { ok: true, status: goWaitlist ? 'Warteliste' : 'Angemeldet', full: goWaitlist };
  } catch (err) {
    console.warn('[DEX] switchSplitGroup error:', err);
    return { ok: false, status: 'Failed', full: false };
  }
}

/**
 * Registrierung stornieren
 */
export async function cancelRegistration(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  cancelledByName?: string,
  cancelledByEmail?: string
): Promise<boolean> {
  try {
    // Audit: wer hat die Abmeldung ausgelöst?
    // Bei Self-Cancel = der User selbst. Bei "Teilnehmer abmelden" durch den
    // Organizer/Admin im Admin Center = der eingeloggte Organizer/Admin.
    const auditName = cancelledByName || svc.context.pageContext.user.displayName || '';
    const auditEmail = (cancelledByEmail || svc.context.pageContext.user.email || '').toLowerCase();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const corePayload: Record<string, any> = {
      'Status': 'Abgemeldet',
      'CancellationDate': new Date().toISOString(),
      'TeilnehmerID': null,
    };
    // Audit-Felder optional dazu - aeltere Subsites haben die Spalten evtl. noch
    // nicht (kommt erst mit Commit a10a608). Ein 400 von SP würde dann die
    // ganze Abmeldung blocken. Strategie: erst mit Audit-Feldern versuchen,
    // bei Misserfolg ohne sie nochmal probieren.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fullPayload: Record<string, any> = { ...corePayload };
    if (auditName) fullPayload['CancelledByName'] = auditName;
    if (auditEmail) fullPayload['CancelledByEmail'] = auditEmail;
    const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`;
    let response = await svc._merge(url, fullPayload);
    if (!response.ok && (auditName || auditEmail)) {
      // Fallback ohne Audit-Felder (Subsite-Liste hat die Spalten noch nicht)
      console.warn('[DEX] cancelRegistration with audit failed (' + response.status + '), retrying without audit fields');
      response = await svc._merge(url, corePayload);
    }
    // v7.31: Counter mit aktuellem Max syncen, damit er nicht "davonrast"
    // wenn der höchste ID-Inhaber sich abmeldet. Best-effort, blockiert
    // die Abmeldung nicht wenn's fehlschlägt. syncCounterToMax liest den
    // Max-Wert intern frisch (race-frei gegen parallele Anmeldungen).
    if (response.ok) {
      try { await svc.syncCounterToMax(subsiteUrl); } catch { /* */ }
      // v24.41: Koordinations-Liste synchron halten — bei JEDER Abmeldung den
      // zugehörigen Assistenz-Link (falls vorhanden) auf 'Cancelled' setzen,
      // damit die Info bei der anderen Seite verschwindet. Deckt alle Cancel-
      // Pfade ab (Self / Proxy / Admin / Team), weil sie hier durchlaufen.
      try { await svc.setAssistantLinkStatusForRegistration(itemId, subsiteUrl, 'Cancelled'); } catch { /* */ }
    }
    return response.ok;
  } catch (err) {
    console.warn('[DEX] cancelRegistration error:', err);
    return false;
  }
}

/**
 * v18.11: Proaktive Absage durch einen Teilnehmer, der sich NICHT angemeldet
 * hat („Ich nehme nicht teil"). Legt eine Teilnehmer-Zeile direkt mit
 * Status='Abgemeldet' an — KEINE Sitzplatz-Reservierung, KEINE TeilnehmerID.
 * Profil-Daten (Vorname/Nachname/Location/JobTitle/Department) werden geladen,
 * damit die Abmeldungs-Liste im Admin-Center dieselben Spalten füllen kann
 * wie Teilnehmer-/Warteliste. Marker `_declined` in CustomData unterscheidet
 * die proaktive Absage von einer regulären Abmeldung (die nach vorheriger
 * Anmeldung erfolgte).
 */
export async function declineRegistration(
  svc: EventService,
  subsiteUrl: string,
  firstName: string,
  surname: string,
  participantEmail: string,
  actorName?: string,
  actorEmail?: string
): Promise<boolean> {
  try {
    const myEmail = (svc.context.pageContext.user.email || '').toLowerCase();
    const profile = (participantEmail || '').toLowerCase() === myEmail
      ? await svc.getCurrentUserProfile()
      : await svc.getUserProfileByEmail(participantEmail);
    const nowIso = new Date().toISOString();
    // v22.57: Claims-Token-Schutz. In manchen Kontexten liefert der Browser
    // als „displayName" das SharePoint-Claims-Login (z.B.
    // „i:0#.f|membership|user@deloitte.de" bzw. „0#.f|membership|…"). Das
    // landete bisher 1:1 als Vorname in der Absage-Zeile. Wir verwenden
    // deshalb bevorzugt den sauberen Namen aus dem Benutzerprofil und filtern
    // Claims-artige Werte raus.
    const looksLikeClaim = (s: string): boolean => /\|membership\||0#\.f\||^i:0#/i.test((s || '').trim());
    const cleanFirst = looksLikeClaim(firstName) ? '' : (firstName || '').trim();
    const cleanLast = looksLikeClaim(surname) ? '' : (surname || '').trim();
    const effFirst = (profile.firstName || '').trim() || cleanFirst;
    const effLast = (profile.lastName || '').trim() || cleanLast;
    // Anzeigename: bevorzugt Profil-PreferredName, sonst Vor-/Nachname,
    // sonst die E-Mail (nie das Claims-Token).
    const effName = (profile.displayName && !looksLikeClaim(profile.displayName) ? profile.displayName : '')
      || `${effFirst} ${effLast}`.trim()
      || participantEmail;
    const auditNameRaw = actorName || svc.context.pageContext.user.displayName || '';
    const auditName = looksLikeClaim(auditNameRaw) ? effName : auditNameRaw;
    const auditEmail = (actorEmail || svc.context.pageContext.user.email || '').toLowerCase();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: Record<string, any> = {
      '__metadata': { 'type': REG_LIST_ITEM_TYPE },
      'Title': participantEmail,
      'Vorname': effFirst,
      'Nachname': effLast,
      'ParticipantName': effName,
      'ParticipantEmail': participantEmail,
      'Department': profile.department,
      'Location': profile.location,
      'JobTitle': profile.jobTitle,
      'Phone': profile.phone,
      'Company': profile.company,
      'Status': 'Abgemeldet',
      'RegistrationDate': nowIso,
      'CancellationDate': nowIso,
      // Marker: proaktive Absage (nie angemeldet gewesen).
      'CustomData': JSON.stringify({ _declined: 'true' }),
    };
    if (auditName) { payload['RegisteredByName'] = auditName; payload['CancelledByName'] = auditName; }
    if (auditEmail) { payload['RegisteredByEmail'] = auditEmail; payload['CancelledByEmail'] = auditEmail; }
    const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`;
    let resp = await svc._post(url, payload);
    if (!resp.ok && (auditName || auditEmail || payload['Company'])) {
      // Fallback ohne Audit-Felder (alte Subsite-Liste ohne diese Spalten).
      // v24.32: zusätzlich Company strippen — fehlt die Spalte, würde der
      // Insert sonst auch im Fallback an Company scheitern.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const noAudit: Record<string, any> = { ...payload };
      delete noAudit['RegisteredByName']; delete noAudit['RegisteredByEmail'];
      delete noAudit['CancelledByName']; delete noAudit['CancelledByEmail'];
      delete noAudit['Company'];
      resp = await svc._post(url, noAudit);
    }
    return resp.ok;
  } catch (err) {
    console.warn('[DEX] declineRegistration error:', err);
    return false;
  }
}

/**
 * Teilnehmer einchecken (Status auf 'Eingecheckt' setzen).
 * v7.16: Erfasst zusätzlich, WANN und VON WEM der Check-In ausgelöst
 * wurde (CheckedInDate / CheckedInByName / CheckedInByEmail). Diese
 * Spalten werden bei neuen Events über createRegistrationList() automatisch
 * angelegt; für bestehende Events muss der Admin einmalig die Kachel
 * "Spalten fixen" im Admin-Center klicken, damit der Check-In nicht mit
 * HTTP 400 fehlschlägt.
 */
export async function checkInParticipant(
  svc: EventService,
  subsiteUrl: string,
  itemId: number
): Promise<boolean> {
  try {
    const me = svc.context.pageContext.user;
    const response = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      {
        'Status': 'Eingecheckt',
        'CheckedInDate': new Date().toISOString(),
        'CheckedInByName': me.displayName || '',
        'CheckedInByEmail': me.email || me.loginName || '',
      }
    );
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * v23.28/v23.29: Teilnehmer als „No-Show" markieren (war angemeldet, aber
 * nicht erschienen). Reuse der Check-in-Audit-Spalten (CheckedInBy*), damit
 * kein neues Schema nötig ist. **Nur für Events, deren Teilnehmerliste die
 * 'No-Show'-Choice kennt** (= ab v23.28 NEU angelegte Events). Bestehende
 * Events werden bewusst NICHT automatisch migriert — dort lehnt SharePoint
 * den Wert ab (HTTP 400) und die Methode liefert `false`.
 */
export async function markNoShowParticipant(
  svc: EventService,
  subsiteUrl: string,
  itemId: number
): Promise<boolean> {
  try {
    const me = svc.context.pageContext.user;
    const response = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      {
        'Status': 'No-Show',
        'CheckedInDate': new Date().toISOString(),
        'CheckedInByName': me.displayName || '',
        'CheckedInByEmail': me.email || me.loginName || '',
      }
    );
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * Teilnehmer auschecken (Status zurück auf 'Angemeldet' setzen)
 */
export async function checkOutParticipant(
  svc: EventService,
  subsiteUrl: string,
  itemId: number
): Promise<boolean> {
  try {
    const response = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      { 'Status': 'Angemeldet' }
    );
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * Status eines Teilnehmers auf 'QR versendet' setzen
 */
export async function setQRSentStatus(
  svc: EventService,
  subsiteUrl: string,
  itemId: number
): Promise<boolean> {
  try {
    const response = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      { 'Status': 'QR versendet' }
    );
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * Registrierung per Email auf einer Subsite finden
 */
export async function getRegistrationByEmail(
  svc: EventService,
  subsiteUrl: string,
  email: string
): Promise<SPRegistration | null> {
  try {
    const response = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=ParticipantEmail eq '${email.replace(/'/g, "''")}'&$select=Id,Title,Vorname,Nachname,ParticipantName,ParticipantEmail,Status,RegistrationDate,RegisteredByName,RegisteredByEmail,CancellationDate,CancelledByName,CancelledByEmail,CustomData,Department,JobTitle,Location,Company&$top=1`,
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
 * Aktuelle Teilnehmeranzahl ermitteln
 */
export async function getRegistrationCount(svc: EventService, subsiteUrl: string): Promise<{ registered: number; waitlist: number }> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const allItems: any[] = [];
  // $top=5000 (SP-REST-Maximum) statt 500 — sonst werden bei Events mit
  // ≥500 Einträgen die Counts auf den Event-Karten falsch berechnet,
  // weil SharePoint bei $orderby+$top mit ILS nicht zuverlässig nextLink
  // liefert wenn die Page exakt voll ist.
  let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Status&$top=5000`;

  while (url) {
    try {
      const response = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) break;
      const data = await response.json();
      allItems.push(...(data.value || data.d?.results || []));
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    } catch {
      break;
    }
  }

  const registered = allItems.filter((i: { Status: string }) => i.Status === 'Angemeldet' || i.Status === 'QR versendet' || i.Status === 'Eingecheckt').length;
  const waitlist = allItems.filter((i: { Status: string }) => i.Status === 'Warteliste').length;
  return { registered, waitlist };
}

/**
 * v22.74: Aktive + Warteliste-E-Mails einer Teilnehmerliste (lowercase) —
 * für die EINDEUTIGE Personenzählung einer Klammer über alle Sub-Events
 * (eine Person, die sich für mehrere Sub-Events anmeldet, zählt einmal).
 */
export async function getParticipantEmailsByStatus(svc: EventService, subsiteUrl: string): Promise<{ active: string[]; waitlist: string[] }> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const allItems: any[] = [];
  let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,ParticipantEmail,Status&$top=5000`;
  while (url) {
    try {
      const response = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) break;
      const data = await response.json();
      allItems.push(...(data.value || data.d?.results || []));
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    } catch {
      break;
    }
  }
  const active: string[] = [];
  const waitlist: string[] = [];
  for (const i of allItems) {
    // v23.3: Zeilen OHNE (gültige) ParticipantEmail wurden frueher komplett
    // übersprungen — dadurch zaehlte die entdoppelte Klammer-/Kachel-Zahl
    // weniger Koepfe als die Tabelle Zeilen hat (z.B. 188 statt 190). Eine
    // Anmeldung ohne E-Mail ist trotzdem ein realer Kopf (belegt einen Platz),
    // bekommt nur keine Mails. Deshalb als eigener Schlüssel (Zeilen-Id)
    // mitzaehlen, statt sie zu verschlucken.
    const email = (i.ParticipantEmail || '').toLowerCase().trim();
    const key = email || `__noemail#${i.Id}`;
    if (i.Status === 'Angemeldet' || i.Status === 'QR versendet' || i.Status === 'Eingecheckt') active.push(key);
    else if (i.Status === 'Warteliste') waitlist.push(key);
  }
  return { active, waitlist };
}

/**
 * Title-Feld (= Teilnehmer-ID) aktualisieren
 */
export async function updateRegistrationTitle(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  newTitle: string
): Promise<boolean> {
  try {
    const response = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      { 'Title': newTitle }
    );
    return response.ok;
  } catch {
    return false;
  }
}
