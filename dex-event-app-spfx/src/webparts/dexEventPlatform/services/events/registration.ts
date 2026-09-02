/**
 * v30.66 — Modularisierung Stufe 2: Thema „Anmelden" — eine Zeile in der
 * Teilnehmerliste anlegen, mit allem, was daran hängt: TeilnehmerID aus dem
 * Zähler, Platz reservieren oder auf die Warteliste, Team-Anmeldung und
 * Team-Lead-Wechsel.
 *
 * Seit v30.42 legt `registerForEvent` die Klammer-Zeile selbst mit an — kein
 * Aufrufer muss mehr daran denken. `skipShadowParent` setzt NUR die
 * Anmeldeseite (sie legt die Klammer zuletzt an, MIT den übergreifenden
 * Antworten), und `shadowEnsuredRef` verhindert 19 gleiche Prüfungen bei 19
 * Terminen (siehe CLAUDE.md).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
// EventService als WERT-Import: nur für den deferred Zugriff auf die statische
// stripNoteWrapper — der Zyklus ist unkritisch, weil der Zugriff erst zur
// Laufzeit in Funktionskörpern passiert (wie REG_LIST_NAME).
import { EventService } from '../EventService';
import type { SPRegistration } from '../EventService';
import { REG_LIST_ITEM_TYPE, REG_LIST_NAME } from '../EventService';
import { sessionIdentities } from '../../utils/sessionIdentities';

// ==================== Registrierungen ====================

/**
 * Registrierung für ein Event erstellen.
 * Operiert auf der Subsite des Events.
 */
export async function registerForEvent(
  svc: EventService,
  subsiteUrl: string,
  firstName: string,
  surname: string,
  participantEmail: string,
  customData: Record<string, string>,
  status: string = 'Angemeldet',
  customFieldMap?: Record<string, string>, // cf.id -> SP InternalName
  starterType?: string, // B2Run: effektiver Typ (nach Fallback)
  preferredStarterType?: string, // B2Run: Wunsch-Typ (was der User eigentlich wollte)
  registeredByName?: string, // Audit: Name des Users der die Anmeldung auslöst
  registeredByEmail?: string, // Audit: E-Mail des Users der die Anmeldung auslöst
  // v18.74: Nachweis der Zustimmung bei stellvertretender Anmeldung — wird in
  // die SP-Spalte ProxyConsent geschrieben (leer bei Selbst-Anmeldung).
  proxyConsent?: string,
  // v19.9: Der CLIENT hat bereits zuverlässig festgestellt, dass der/die
  // Anmeldende Haupt- oder Co-Organizer dieses Events ist (aus dem geladenen
  // Event-Objekt: organizerEmails/coOrganizerEmails ⊇ aktueller User). Diese
  // Information wird hier durchgereicht und hat Vorrang vor der fragilen
  // serverseitigen Ableitung (SubsiteUrl-Filter + Note-Feld-Parsing +
  // pageContext-Identität), die im Tenant gelegentlich fehlschlug und
  // legitime Organizer mit „bereits angemeldet" ablehnte.
  actorIsEventOrganizer: boolean = false,
  // v23.10: Der CLIENT hat bereits validiert, dass der/die Anmeldende eine
  // Assistenz ist UND das Ziel Partner/Director (RegistrationPage prüft das
  // beim Submit gegen den Picker-JobTitle). Diesem Flag vertrauen wir — die
  // serverseitige Ableitung in canRegisterForOthers hängt an
  // Profil-Lookups (Title/SPS-JobTitle), die im Tenant unzuverlässig leer
  // zurückkamen und legitime Assistenzen mit „nicht berechtigt" ablehnten.
  // Gilt NUR für Check A (Berechtigung) — NICHT für die Deadline (Assistenz
  // darf wie ein normaler User nicht nach Frist anmelden).
  clientAssistantAllowed: boolean = false
// v23.9: Statt nacktem boolean ein konkreter Grund bei Misserfolg, damit die
// UI nicht mehr pauschal „bereits registriert" anzeigt (irreführend, wenn der
// echte Grund Berechtigung/Deadline/Insert-Fehler war).
  // v30.58: `detail` trägt die Klartext-Antwort von SharePoint bei einem
  // abgelehnten Insert (z.B. „The field or property 'X' does not exist") —
  // der `reason` bleibt maschinenlesbar, die Ursache geht nicht verloren.
): Promise<{ ok: boolean; reason?: 'not-allowed' | 'deadline' | 'insert-failed' | 'error'; detail?: string }> {
  try {
    // ---- Permission-Checks (v3.9.2 / v3.9.3) ----
    // Serverseitige Prüfungen — nicht perfekt (SPFx läuft im Browser),
    // aber fangt naiven App-Bypass (F12, direkter Service-Aufruf) ab.
    const sessionEmail = (svc.context.pageContext.user.email || '').toLowerCase();
    const targetEmail = (participantEmail || '').toLowerCase();
    // v30.67: Alle Schreibweisen der angemeldeten Person (pageContext-E-Mail
    // UND die Adresse aus dem loginName) — dieselbe Menge wie in
    // `canRegisterForOthers`. Steht im Event die SMTP-Adresse und liefert der
    // pageContext den UPN/Alias, scheiterte die Frist-Ausnahme unten sonst
    // auch für den HAUPT-Organizer.
    const sessionIds = sessionIdentities(svc.context);

    // Event-Metadaten laden (Deadline + OrganizerEmail) über SubsiteUrl.
    // Beide Checks nutzen die gleiche Abfrage — einmal laden, mehrfach prüfen.
    let eventDeadline = '';
    let eventOrganizerEmails: string[] = [];
    try {
      const subsiteEsc = encodeURIComponent(subsiteUrl.replace(/'/g, "''"));
      const evResp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${subsiteEsc}'&$top=1&$select=RegistrationDeadline,OrganizerEmail,EmailTemplateOverrides`,
        SPHttpClient.configurations.v1
      );
      if (evResp.ok) {
        const evData = await evResp.json();
        const items = evData.value || evData.d?.results || [];
        if (items.length > 0) {
          eventDeadline = items[0].RegistrationDeadline || '';
          // v30.67: Derselbe Personenkreis wie in `canRegisterForOthers`
          // (services/events/profileData.ts, v19.6): HTML des Note-Felds
          // robust strippen, an `;`/`,`/Zeilenumbruch splitten UND die
          // Co-Organizer aus dem Piggyback `_coOrganizers` dazunehmen.
          // Vorher zählte hier nur `OrganizerEmail` mit `;`-Split — ein
          // Co-Organizer wurde nach Fristablauf abgewiesen, obwohl die
          // Rollenmatrix „Nach Anmeldefrist registrieren" für ihn zusagt
          // und Check A ihn wenige Zeilen darüber längst akzeptiert.
          const splitEmails = (raw: string | null | undefined): string[] =>
            (raw || '')
              .replace(/<br\s*\/?>/gi, ';')
              .replace(/<\/div>\s*<div[^>]*>/gi, ';')
              .replace(/<[^>]+>/g, '')
              .split(/[;,\n\r]+/)
              .map(s => s.trim().toLowerCase())
              .filter(Boolean);
          eventOrganizerEmails = splitEmails(items[0].OrganizerEmail);
          try {
            const ovRaw = EventService.stripNoteWrapper(items[0].EmailTemplateOverrides) || '{}';
            const ov = JSON.parse(ovRaw);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const list = (ov as any)._coOrganizers;
            if (Array.isArray(list)) {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              for (const x of list as any[]) {
                const e = String(x?.email || '').toLowerCase().trim();
                if (e) eventOrganizerEmails.push(e);
              }
            }
          } catch { /* kein/ungültiges Override-JSON → keine Co-Organizer */ }
        }
      }
    } catch { /* Bei Load-Fehler konservativ weitermachen — andere Checks greifen */ }

    // Check A: Darf der User für eine andere Person registrieren?
    // v19.9: Wenn der Client bereits bestätigt hat, dass der/die Anmeldende
    // Organizer/Co-Organizer dieses Events ist, vertrauen wir dem (gleiche
    // Datengrundlage wie die Button-Sichtbarkeit) und überspringen die
    // fragile serverseitige Ableitung. Sonst Fallback auf canRegisterForOthers
    // (deckt Admin-Rolle + Assistant-Ausnahme zuverlässig ab).
    if (targetEmail && targetEmail !== sessionEmail) {
      const allowed = actorIsEventOrganizer || clientAssistantAllowed || await svc.canRegisterForOthers(subsiteUrl, participantEmail);
      if (!allowed) {
        console.warn(`[DEX] registerForEvent DENIED: ${sessionEmail} versuchte ${targetEmail} zu registrieren — weder Organizer noch Admin noch erlaubter Assistant-Fall.`);
        return { ok: false, reason: 'not-allowed' };
      }
    }

    // Check B: Deadline abgelaufen? Nur Event-Organizer + Admin dürfen nach
    // Deadline registrieren (auch für sich selbst). Assistant NICHT — das ist
    // wie ein normaler User.
    if (eventDeadline) {
      const deadlineDate = new Date(eventDeadline);
      if (!isNaN(deadlineDate.getTime()) && deadlineDate < new Date()) {
        // v30.67: Haupt- UND Co-Organizer, gegen ALLE Session-Identitäten —
        // plus das Client-Flag, dem Check A (die schärfere Prüfung) seit
        // v19.9 bereits vertraut; es kommt aus denselben DEX_Events-Daten
        // wie die Sichtbarkeit des „Teilnehmer hinzufügen"-Dialogs. Die
        // Assistenz-Ausnahme (`clientAssistantAllowed`) bleibt hier bewusst
        // draußen: Sie darf die Frist nicht umgehen.
        const isEventOrganizer = actorIsEventOrganizer
          || eventOrganizerEmails.some(e => sessionIds.has(e));
        let isAdmin = false;
        try {
          const esc = sessionEmail.replace(/'/g, "''");
          const roleResp = await svc._sp.get(
            `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
            SPHttpClient.configurations.v1
          );
          if (roleResp.ok) {
            const rd = await roleResp.json();
            const rItems = rd.value || rd.d?.results || [];
            if (rItems.length > 0 && (rItems[0].Role === 'Admin' || rItems[0].Role === 'IT-Admin')) isAdmin = true;
          }
        } catch { /* ignore */ }

        if (!isEventOrganizer && !isAdmin) {
          console.warn(`[DEX] registerForEvent DENIED (deadline): ${sessionEmail} versuchte nach Deadline ${eventDeadline} zu registrieren — weder Event-Organizer noch Admin.`);
          return { ok: false, reason: 'deadline' };
        }
      }
    }
    // ---- Ende Permission-Checks ----

    // v7.28 / v9.10: Nächste TeilnehmerID atomar über den Subsite-Counter
    // holen (ETag-CAS, verhindert Race-Conditions bei parallelen Anmeldungen).
    // Counter wird bei Bedarf on-demand angelegt + geseeded.
    //
    // v9.10: Der alte race-anfällige Fallback "max+1" wurde entfernt — bei
    // Massen-Anmeldungen (Go-Live große Events) hat er Duplikate produziert,
    // weil zwei Clients gleichzeitig den gleichen Max-Wert lesen und beide
    // mit Max+1 schreiben. Wenn der atomare Counter ausnahmsweise gar nicht
    // erreichbar ist, lassen wir TeilnehmerID undefined und der Admin
    // lädt anschliessend "IDs neu vergeben" — Lückenfreiheit ist nicht
    // hart kritisch, Eindeutigkeit ist es.
    const nextId = await svc.getNextTeilnehmerId(subsiteUrl);

    // Profildaten laden - für den TATSAECHLICHEN Teilnehmer (nicht den eingeloggten User!)
    // Wenn jemand für eine andere Person registriert, muss deren Profil geladen werden,
    // sonst wird der eigene JobTitle/Department/Office in deren Teilnehmer-Eintrag geschrieben.
    const myEmail = (svc.context.pageContext.user.email || '').toLowerCase();
    const profile = participantEmail.toLowerCase() === myEmail
      ? await svc.getCurrentUserProfile()
      : await svc.getUserProfileByEmail(participantEmail);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: Record<string, any> = {
      '__metadata': { 'type': REG_LIST_ITEM_TYPE },
      'Title': participantEmail,
      // v9.10: TeilnehmerID nur setzen wenn der atomare Counter sie geliefert hat.
      // Bei Counter-Outage bleibt das Feld leer — Admin kann nachträglich
      // "IDs neu vergeben" laufen lassen, was sequentielle IDs setzt.
      ...(typeof nextId === 'number' ? { 'TeilnehmerID': nextId } : {}),
      'Anrede': customData.salutation || '',
      'Vorname': firstName,
      'Nachname': surname,
      'ParticipantName': `${firstName} ${surname}`,
      'ParticipantEmail': participantEmail,
      'Department': profile.department,
      'Location': profile.location,
      'JobTitle': profile.jobTitle,
      'Phone': profile.phone,
      // v24.29: Unternehmenszugehörigkeit / Rechtsträger mitschreiben.
      'Company': profile.company,
      'Status': status,
      'RegistrationDate': new Date().toISOString(),
      'CustomData': JSON.stringify(customData),
    };

    // Audit: wer hat die Anmeldung ausgelöst?
    // Bei Self-Registration = der User selbst. Bei "Für andere Person registrieren"
    // = der Organizer/Admin der geklickt hat. Fallback wenn nichts übergeben: aus pageContext.
    const auditName = registeredByName || svc.context.pageContext.user.displayName || '';
    const auditEmail = (registeredByEmail || svc.context.pageContext.user.email || '').toLowerCase();
    if (auditName) payload['RegisteredByName'] = auditName;
    if (auditEmail) payload['RegisteredByEmail'] = auditEmail;
    // v18.74: Zustimmungs-Nachweis bei stellvertretender Anmeldung.
    if (proxyConsent) payload['ProxyConsent'] = proxyConsent;

    // B2Run: Starter-Typ + Wunsch-Typ schreiben (bei normalen Events null)
    if (starterType) payload['StarterType'] = starterType;
    if (preferredStarterType) payload['PreferredStarterType'] = preferredStarterType;

    // Custom Field Werte in die echten SP-Spalten schreiben.
    // Wichtig: Wenn spInternalName fehlt (z.B. weil der Admin das Feld später
    // ergänzt hat ohne Spalte in der Teilnehmerliste), würde der Wert
    // SILENT VERLOREN GEHEN — deshalb ein console.warn damit der Admin im
    // Admin Center per "Custom Fields prüfen" das Mapping fixen kann.
    if (customFieldMap) {
      for (const cfId of Object.keys(customData)) {
        if (cfId === 'salutation') continue;
        if (!customData[cfId]) continue;
        const spFieldName = customFieldMap[cfId];
        if (spFieldName) {
          payload[spFieldName] = customData[cfId];
        } else {
          console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields prüfen' ausführen.`);
        }
      }
    }

    let response = await svc._post(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
      payload
    );
    // v19.9 BUG-FIX: Stellvertretende Anmeldung schlug auf älteren
    // Teilnehmerlisten fehl, weil die erst mit v18.74 eingeführte Spalte
    // `ProxyConsent` dort noch nicht existiert — der Insert mit
    // `ProxyConsent` im Body wird dann von SharePoint mit HTTP 400
    // ("property does not exist") abgewiesen. Da `ProxyConsent` NUR bei
    // Stellvertreter-Anmeldungen im Body steht, scheiterte ausschließlich der
    // Proxy-Pfad (Selbst-Anmeldung lief, weil dort kein ProxyConsent gesetzt
    // wird) — für den User sah es aus wie „Person bereits angemeldet", obwohl
    // gar nichts gespeichert wurde. Fix: Insert einmal OHNE das optionale
    // Audit-Feld wiederholen, damit die Anmeldung nicht an einer fehlenden
    // Spalte scheitert. Der Zustimmungs-Nachweis geht dann verloren (der
    // Admin kann die Spalte per „Spalten fixen" nachrüsten), die Anmeldung
    // selbst gelingt aber.
    // v24.32: Gleiches Schutz-Muster jetzt AUCH für die v24.29-Spalte
    // `Company` — auf Teilnehmerlisten, auf denen „Spalten fixen" noch nicht
    // lief, existiert sie nicht → der Insert mit `Company` im Body würde sonst
    // mit HTTP 400 scheitern und die GANZE Anmeldung kaputtmachen. Deshalb bei
    // einem fehlgeschlagenen Insert das optionale Feld strippen und erneut
    // versuchen. Folge ohne Spalte: Anmeldung gelingt, Unternehmens-Wert wird
    // nicht getrackt (Admin kann „Spalten fixen" nachziehen).
    if (!response.ok && (payload['ProxyConsent'] || payload['Company'])) {
      console.warn('[DEX] registerForEvent: Insert fehlgeschlagen — Retry OHNE ProxyConsent/Company (Spalte evtl. nicht vorhanden). Bitte im Admin Center "Spalten fixen" ausführen.');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const retryPayload: Record<string, any> = { ...payload };
      delete retryPayload['ProxyConsent'];
      delete retryPayload['Company'];
      response = await svc._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
        retryPayload
      );
    }
    if (!response.ok) {
      // v30.58: Den Grund AUSLESEN statt nur „insert-failed" zu melden.
      //
      // Bisher endete jeder abgelehnte Insert als anonymes „insert-failed" —
      // und genau das hat die Suche nach der fehlenden Klammer-Zeile so lange
      // aufgehalten: SharePoint sagt in der Antwort, WELCHE Spalte es nicht
      // gibt („The field or property 'X' does not exist"), nur hat es nie
      // jemand gelesen. Der Retry oben strippt bloß zwei fest verdrahtete
      // Spalten (`ProxyConsent`, `Company`); ein nachträglich angelegtes
      // Abfragefeld, dessen Spalte auf DIESER Liste fehlt, bringt den ganzen
      // Insert zu Fall — und zwar nur bei den Personen, die das Feld
      // ausfüllen. Das erklärt, warum es immer dieselben wenigen trifft und
      // nicht zufällig streut wie eine Drosselung.
      let detail = '';
      try {
        const txt = await response.text();
        const m = txt.match(/"value"\s*:\s*"([^"]{0,400})"/);
        detail = (m ? m[1] : txt).slice(0, 400);
      } catch { /* Body nicht lesbar */ }
      console.warn('[DEX] registerForEvent: Insert abgelehnt', {
        status: response.status, subsiteUrl, detail,
        felder: Object.keys(payload),
      });
      return { ok: false, reason: 'insert-failed', detail };
    }

    // Inserted-Item-Id EINMALIG aus der Response lesen (der Body lässt sich
    // nur einmal konsumieren) — wird sowohl für die Dedup-Prüfung als auch
    // für das Setzen des Autors (stellvertretende Anmeldung) gebraucht.
    let insertedId = 0;
    try {
      const respJson = await response.json();
      insertedId = respJson?.d?.Id || respJson?.Id || 0;
    } catch { /* Body nicht lesbar — Dedup/Autor-Set entfallen, Insert war ok */ }

    // v9.10: Post-Insert Safety Net — bei Massen-Anmeldungen (Go-Live)
    // gab es trotz ETag-Counter vereinzelt Duplikate. Ursache war der
    // alte max+1-Fallback (jetzt entfernt) und ggf. Edge-Cases im
    // Counter-Pfad. Als zusätzliche Versicherung: nach dem Insert
    // prüfen, ob jetzt zwei Einträge dieselbe TeilnehmerID haben.
    // Wenn ja: der mit der HOEHEREN SP-Item-Id verliert (= der spätere
    // Insert), holt sich frisch eine ID am Counter und patcht sich.
    // So bleiben die zuerst eingetroffenen Anmeldungen stabil.
    if (typeof nextId === 'number' && nextId > 0 && insertedId > 0) {
      try {
        const dupResp = await svc._sp.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,TeilnehmerID&$filter=TeilnehmerID eq ${nextId}&$top=10`,
          SPHttpClient.configurations.v1
        );
        if (dupResp.ok) {
          const dupData = await dupResp.json();
          const dupItems: Array<{ Id: number; TeilnehmerID: number }> = dupData.value || dupData.d?.results || [];
          if (dupItems.length > 1) {
            const minId = Math.min(...dupItems.map(d => d.Id));
            if (insertedId !== minId) {
              // Wir haben verloren — fresh ID holen + patchen
              const fresh = await svc.getNextTeilnehmerId(subsiteUrl);
              if (typeof fresh === 'number' && fresh > 0) {
                await svc._merge(
                  `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${insertedId})`,
                  { 'TeilnehmerID': fresh }
                );
                console.warn(`[DEX] Post-insert dedup: TeilnehmerID ${nextId} kollidierte, Item ${insertedId} hat jetzt #${fresh}.`);
              } else {
                console.warn(`[DEX] Post-insert dedup: kollidierende TeilnehmerID ${nextId} entdeckt, aber Counter lieferte keine fresh ID. Admin sollte "IDs neu vergeben" laufen lassen.`);
              }
            }
          }
        }
      } catch (err) {
        // Safety-Net-Fehler nicht kritisch — Insert war erfolgreich
        console.warn('[DEX] Post-insert dedup check fehlgeschlagen:', err);
      }
    }

    // v20.5: Stellvertretende Anmeldung (Akteur != Teilnehmer) → den
    // Teilnehmer zum Autor der Zeile machen, damit er seine eigene Anmeldung
    // in "Meine Events" sieht und sich selbst abmelden kann. auditEmail ist
    // bereits lowercased; bei Selbst-Anmeldung sind beide gleich → kein Set.
    if (insertedId > 0 && auditEmail && participantEmail && participantEmail.toLowerCase().trim() !== auditEmail) {
      await svc.trySetItemAuthor(subsiteUrl, REG_LIST_NAME, insertedId, participantEmail);
    }

    return { ok: true };
  } catch {
    return { ok: false, reason: 'error' };
  }
}

/**
 * v11.82: Ein einzelnes Teilnehmer-Item im Team-Modus anlegen.
 *
 * Unterschied zu `registerForEvent`: kein eigener Permission-Check (der
 * Aufrufer hat schon im Team-Submit alle Mitglieder validiert), kein
 * Post-Insert Dedup-Loop (der ist im Team-Pfad überflüssig — wenn ein
 * Member mit Kollision verliert, fixt es der Folge-IDReorder). Nimmt
 * Profil-Daten und Anzeige-Namen direkt entgegen, weil der Lead-Submit
 * pro Member ohnehin schon das Graph-Profil geladen hat.
 */
export async function registerTeamMember(
  svc: EventService,
  subsiteUrl: string,
  args: {
    firstName: string;
    lastName: string;
    email: string;
    profile: { department: string; location: string; jobTitle: string; phone: string; company?: string };
    status: 'Angemeldet' | 'Warteliste';
    teamId: string;
    teamLead: boolean;
    teamName?: string;
    customData?: Record<string, string>;
    customFieldMap?: Record<string, string>;
    starterType?: string;
    preferredStarterType?: string;
    registeredByName?: string;
    registeredByEmail?: string;
    salutation?: string;
  }
): Promise<{ ok: boolean; teilnehmerId?: number; itemId?: number }> {
  try {
    const nextId = await svc.getNextTeilnehmerId(subsiteUrl);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: Record<string, any> = {
      '__metadata': { 'type': REG_LIST_ITEM_TYPE },
      'Title': args.email,
      ...(typeof nextId === 'number' ? { 'TeilnehmerID': nextId } : {}),
      'Anrede': args.salutation || '',
      'Vorname': args.firstName,
      'Nachname': args.lastName,
      'ParticipantName': `${args.firstName} ${args.lastName}`.trim(),
      'ParticipantEmail': args.email,
      'Department': args.profile.department,
      'Location': args.profile.location,
      'JobTitle': args.profile.jobTitle,
      'Phone': args.profile.phone,
      'Company': args.profile.company || '',
      'Status': args.status,
      'RegistrationDate': new Date().toISOString(),
      'TeamId': args.teamId,
      'TeamLead': !!args.teamLead,
      'TeamName': args.teamName || '',
      'CustomData': JSON.stringify(args.customData || {}),
    };
    if (args.registeredByName) payload['RegisteredByName'] = args.registeredByName;
    if (args.registeredByEmail) payload['RegisteredByEmail'] = args.registeredByEmail;
    if (args.starterType) payload['StarterType'] = args.starterType;
    if (args.preferredStarterType) payload['PreferredStarterType'] = args.preferredStarterType;
    if (args.customFieldMap && args.customData) {
      for (const cfId of Object.keys(args.customData)) {
        if (cfId === 'salutation') continue;
        const v = args.customData[cfId];
        if (!v) continue;
        const spName = args.customFieldMap[cfId];
        if (spName) payload[spName] = v;
      }
    }
    let response = await svc._post(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
      payload
    );
    // v24.32: Retry OHNE Company, falls die Spalte auf der Liste fehlt (s.
    // registerForEvent) — sonst bräche die Team-Anmeldung auf Alt-Listen.
    if (!response.ok && payload['Company']) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const retryPayload: Record<string, any> = { ...payload };
      delete retryPayload['Company'];
      response = await svc._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
        retryPayload
      );
    }
    if (!response.ok) return { ok: false };
    try {
      const respJson = await response.json();
      const itemId: number = respJson?.d?.Id || respJson?.Id || 0;
      return { ok: true, teilnehmerId: typeof nextId === 'number' ? nextId : undefined, itemId };
    } catch {
      return { ok: true, teilnehmerId: typeof nextId === 'number' ? nextId : undefined };
    }
  } catch {
    return { ok: false };
  }
}

/**
 * v11.82: Alle Mitglieder eines Teams (per TeamId) zu einer Registrierung
 * laden — wird in „Meine Events" zum Rendern des Team-Badges genutzt.
 */
export async function getTeamMembers(svc: EventService, subsiteUrl: string, teamId: string): Promise<SPRegistration[]> {
  if (!teamId) return [];
  try {
    const tidEsc = teamId.replace(/'/g, "''");
    const response = await svc._sp.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=TeamId eq '${tidEsc}'&$top=100&$orderby=TeamLead desc,Id asc`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return [];
    const data = await response.json();
    return data.value || data.d?.results || [];
  } catch {
    return [];
  }
}

/**
 * v11.83: Auf einer existierenden Teilnehmer-Zeile das Feld TeamLead
 * auf true setzen (Auto-Promote nach Lead-Cancel). MERGE auf der
 * Teilnehmerliste — die Subsite kennt das Item über `itemId`.
 */
export async function promoteToTeamLead(svc: EventService, subsiteUrl: string, itemId: number): Promise<boolean> {
  try {
    const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`;
    const resp = await svc._merge(url, { TeamLead: true });
    return !!resp.ok;
  } catch {
    return false;
  }
}

/**
 * v11.84: Lead-Rolle innerhalb eines Teams von einer Person auf eine andere
 * übergeben. Wird im Admin Center per Dropdown im Teams-Block ausgelöst.
 * Best-effort transaktional: erst die neue Lead-Zeile auf TeamLead=true
 * setzen, danach die alte auf TeamLead=false. Schlägt der zweite MERGE
 * fehl, gibt es kurzfristig zwei Leads — der Aufrufer kann dann erneut
 * versuchen oder die Liste manuell reparieren. Keine echte Transaktion,
 * SharePoint bietet sowas auf Listen-Ebene nicht.
 */
export async function transferTeamLead(
  svc: EventService,
  subsiteUrl: string,
  fromLeadItemId: number,
  toNewLeadItemId: number
): Promise<boolean> {
  if (!subsiteUrl || !fromLeadItemId || !toNewLeadItemId || fromLeadItemId === toNewLeadItemId) {
    return false;
  }
  try {
    const newUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${toNewLeadItemId})`;
    const r1 = await svc._merge(newUrl, { TeamLead: true });
    if (!r1.ok) return false;
    const oldUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${fromLeadItemId})`;
    const r2 = await svc._merge(oldUrl, { TeamLead: false });
    return !!r2.ok;
  } catch {
    return false;
  }
}

/**
 * v11.83: Prüfen, ob eine bestimmte Email-Adresse bereits aktiv beim
 * Event angemeldet ist (Status in Angemeldet/QR versendet/Eingecheckt/
 * Warteliste). Wird vor jedem Team-Add (Initial, Add-Member, Beitritt)
 * benutzt, um Doppel-Anmeldungen sauber abzuweisen, bevor ein Sitzplatz
 * reserviert wird.
 *
 * Rückgabe: true = blockieren, false = frei (auch bei SP-Fehlern, weil
 * der Aufrufer dann auf die strikteren Stellen-internen Checks zurück-
 * fällt; ein lauter Throw würde den Pfad unnötig abbrechen).
 */
export async function isUserAlreadyOnEvent(svc: EventService, subsiteUrl: string, email: string): Promise<boolean> {
  if (!subsiteUrl || !email) return false;
  try {
    const emEsc = email.trim().replace(/'/g, "''");
    const blockingStatuses = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const statusClause = blockingStatuses.map(s => `Status eq '${s}'`).join(' or ');
    const filter = `(ParticipantEmail eq '${emEsc}') and (${statusClause})`;
    const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=${encodeURIComponent(filter)}&$top=1&$select=Id,Status,ParticipantEmail`;
    const response = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) return false;
    const data = await response.json();
    const items = data.value || data.d?.results || [];
    return items.length > 0;
  } catch {
    return false;
  }
}
