/**
 * v30.66 — Modularisierung Stufe 2: Thema „Event-Subsite anlegen und
 * berechtigen": Subsite je Event, Teilnehmerliste mit den Spalten der
 * Custom Fields, Organizer-Rechte, Item-Level-Security und die
 * Assistenz-/Proxy-Reparaturen.
 *
 * Berechtigungen gelten JE SUBSITE, und jedes Sub-Event hat eine eigene:
 * `ensureOrganizerPermissionsMulti` löst die User-Id einmal je Person auf und
 * setzt sie auf allen übergebenen Subsites — der Klammer-Pfad ist nie der
 * ganze Pfad (siehe CLAUDE.md, v30.37).
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService, CustomField } from '../EventService';
import { REG_LIST_NAME } from '../EventService';

// ==================== Subsites ====================

/**
 * URL-Suffix aus Event-Titel generieren.
 * "B2Run Frankfurt 2026" → "b2run-frankfurt-2026-k8f3a"
 */
function generateSubsiteUrl(svc: EventService, title: string): string {
  const slug = title
    .toLowerCase()
    .replace(/[äÄ]/g, 'ae').replace(/[öÖ]/g, 'oe').replace(/[üÜ]/g, 'ue').replace(/ß/g, 'ss')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .substring(0, 40);
  const suffix = Date.now().toString(36).slice(-5);
  return `${slug}-${suffix}`;
}

/**
 * Subsite für ein Event erstellen.
 * Versucht mehrere Templates falls eines fehlschlägt.
 * Gibt die absolute URL der neuen Subsite zurück.
 */
export async function createEventSubsite(svc: EventService, title: string, description: string): Promise<string | null> {
  const urlSuffix = generateSubsiteUrl(svc, title);
  const desc = description || `Event-Subsite: ${title}`;

  // Templates in Reihenfolge versuchen:
  // STS#3 = Modern ohne Group, STS#0 = Classic Team Site, STS = Blank
  const templates = ['STS#3', 'STS#0', 'STS'];

  for (const template of templates) {
    try {
      const payload = {
        'parameters': {
          '__metadata': { 'type': 'SP.WebCreationInformation' },
          'Title': title,
          'Url': urlSuffix,
          'Description': desc,
          'Language': 1031,
          'WebTemplate': template,
          'UseSamePermissionsAsParentSite': false,
        }
      };

      const response = await svc._post(`${svc.siteUrl}/_api/web/webs/add`, payload);
      if (response.ok) {
        const result = await response.json();
        const subsiteAbsoluteUrl = result.d?.Url || result.Url;
        // Subsite erfolgreich erstellt
        return subsiteAbsoluteUrl || `${svc.siteUrl}/${urlSuffix}`;
      }

      // Fehlerdetails loggen
      try {
        const err = await response.json();
        console.warn(`[DEX] Template ${template} fehlgeschlagen (${response.status}):`, err.error?.message?.value || err);
      } catch {
        console.warn(`[DEX] Template ${template} fehlgeschlagen: ${response.status}`);
      }
    } catch (e) {
      console.warn(`[DEX] Template ${template} Fehler:`, e);
    }
  }

  console.error('[DEX] Subsite konnte mit keinem Template erstellt werden');
  return null;
}

// ==================== Teilnehmerlisten (auf Subsites) ====================

/**
 * Teilnehmerliste auf einer Subsite erstellen.
 * Liste heißt immer "Teilnehmer".
 */
export async function createRegistrationList(
  svc: EventService,
  subsiteUrl: string,
  customFields: CustomField[],
  organizerEmail: string
): Promise<Record<string, string>> {
  // Liste erstellen
  await svc._post(`${subsiteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': REG_LIST_NAME,
    'Description': 'Teilnehmerliste für dieses Event',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });

  // Basis-Spalten
  const baseFields = [
    { title: 'TeilnehmerID', type: 9 }, // Number - fortlaufende ID
    { title: 'Anrede', type: 6, choices: ['Frau', 'Herr', 'Divers'], metaType: 'SP.FieldChoice' },
    { title: 'Vorname', type: 2 },
    { title: 'Nachname', type: 2 },
    { title: 'ParticipantName', type: 2 }, // Backward compat
    { title: 'ParticipantEmail', type: 2 },
    { title: 'Department', type: 2 },
    { title: 'Location', type: 2 },
    { title: 'JobTitle', type: 2 },
    { title: 'Phone', type: 2 },
    // v24.29: Unternehmenszugehörigkeit / Rechtsträger (aus dem Profil).
    { title: 'Company', type: 2 },
    { title: 'Status', type: 6, choices: ['Angemeldet', 'QR versendet', 'Warteliste', 'Eingecheckt', 'No-Show', 'Abgemeldet'], metaType: 'SP.FieldChoice' },
    { title: 'StarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' }, // B2Run: Typ-Auswahl
    { title: 'PreferredStarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' }, // B2Run: Wunsch-Typ (wenn Fallback oder Warteliste)
    // v10.13: B2Run-Leistungsnachweis-Bestätigung. Virtuelles Feld der
    // RegistrationPage, das nur durchläuft wenn durchstarterRequiresProof
    // aktiv ist — die SP-Spalte muss aber existieren sonst kippt die
    // Anmeldung mit HTTP 400. Wird auf jeder neuen Teilnehmerliste angelegt
    // damit B2Run-Events nicht später nochmal manuell repariert werden müssen.
    { title: 'b2run_leistungsnachweis', type: 2 },
    { title: 'QuizScore', type: 9 }, // Number - Anzahl richtiger Antworten
    { title: 'QuizAnswers', type: 3 }, // Note - JSON der Antworten (für Statistik)
    { title: 'QuizCompletedAt', type: 4 }, // DateTime
    { title: 'RegistrationDate', type: 4 },
    { title: 'RegisteredByName', type: 2 },  // Audit: Name des Users der die Anmeldung durchgeführt hat
    { title: 'RegisteredByEmail', type: 2 }, // Audit: E-Mail des Users der die Anmeldung durchgeführt hat
    { title: 'ProxyConsent', type: 3 },      // v18.74: Nachweis der Zustimmung bei stellvertretender Anmeldung (Note)
    { title: 'LastModifiedDate', type: 4 },
    { title: 'ChangeLog', type: 3 }, // Note (multiline) - Änderungshistorie
    { title: 'CancellationDate', type: 4 },
    { title: 'CancelledByName', type: 2 },   // Audit: Name des Users der die Abmeldung ausgelöst hat
    { title: 'CancelledByEmail', type: 2 },  // Audit: E-Mail des Users der die Abmeldung ausgelöst hat
    { title: 'CheckedInDate', type: 4 },     // v7.16: Check-In-Audit — Zeitpunkt
    { title: 'CheckedInByName', type: 2 },   // v7.16: Check-In-Audit — Helfer-Name
    { title: 'CheckedInByEmail', type: 2 },  // v7.16: Check-In-Audit — Helfer-E-Mail
    // v17.15: Nachrück-Audit (siehe SPRegistration-Interface):
    // - PromotedDate: gesetzt beim Promote auf die nachrückende Person.
    // - ReplacedParticipantEmail: E-Mail der Person, deren Cancel den
    //   Promote ausgelöst hat („Ersetzt wen") — auf der promoteten Person.
    // - ReplacedByParticipantEmail: E-Mail der nachrückenden Person
    //   („Ersetzt durch") — auf der cancelnden Person.
    { title: 'PromotedDate', type: 4 },
    { title: 'ReplacedParticipantEmail', type: 2 },
    { title: 'ReplacedByParticipantEmail', type: 2 },
    // v11.36: Überbuchungs-Review-Marker. '' = normal, 'Pending' = vom
    // „Überbuchung prüfen"-Lauf als über Kapazität erkannt; der Admin
    // entscheidet pro Person (auf Warteliste / Platz behalten).
    { title: 'OverbookReview', type: 2 },
    { title: 'ConsentReview', type: 2 }, // v26.47: Externe Anmeldung — 'Pending' = Datenschutz-Rückmeldung offen
    // v11.82: Team-Anmeldung — drei Spalten gruppieren Mitglieder eines
    // gemeinsam angemeldeten Teams. TeamId = UUID (gleicher Wert für alle
    // Mitglieder), TeamLead = true nur für die anmeldende Person, TeamName
    // = optionaler frei wählbarer Name (nur wenn das Event AskTeamName an
    // hat). Bei Nicht-Team-Anmeldungen bleiben alle drei Felder leer.
    { title: 'TeamId', type: 2 },
    { title: 'TeamLead', type: 8 },
    { title: 'TeamName', type: 2 },
    { title: 'CustomData', type: 3 },
  ];

  for (const f of baseFields) {
    const payload: Record<string, unknown> = {
      '__metadata': { 'type': f.metaType || 'SP.Field' },
      'Title': f.title,
      'FieldTypeKind': f.type,
      'Required': false,
    };
    if ((f as { choices?: string[] }).choices) {
      payload['Choices'] = { 'results': (f as { choices: string[] }).choices };
    }
    await svc._post(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, payload);
  }

  // Custom Fields als eigene Spalten anlegen + InternalName merken
  const customFieldViewNames: string[] = [];
  const fieldMap: Record<string, string> = {}; // cf.id -> SP InternalName
  for (const cf of customFields) {
    if (!cf.label) continue;
    // v19.0: Dokument-Felder bekommen KEINE Spalte — die Datei wird als
    // Attachment an die Teilnehmer-Zeile gehängt, nicht als Spaltenwert.
    if (cf.type === 'document') continue;
    let fieldPayload: Record<string, unknown>;

    if (cf.type === 'select' && cf.options && cf.options.length > 0) {
      fieldPayload = {
        '__metadata': { 'type': 'SP.FieldChoice' },
        'Title': cf.label,
        'FieldTypeKind': 6,
        'Required': false,
        'Choices': { 'results': cf.options },
      };
    } else if (cf.type === 'number') {
      fieldPayload = {
        '__metadata': { 'type': 'SP.Field' },
        'Title': cf.label,
        'FieldTypeKind': 9,
        'Required': false,
      };
    } else if (cf.type === 'checkbox') {
      fieldPayload = {
        '__metadata': { 'type': 'SP.Field' },
        'Title': cf.label,
        'FieldTypeKind': 8,
        'Required': false,
      };
    } else {
      fieldPayload = {
        '__metadata': { 'type': 'SP.Field' },
        'Title': cf.label,
        'FieldTypeKind': 2,
        'Required': false,
      };
    }

    try {
      const fieldResponse = await svc._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`,
        fieldPayload
      );
      if (fieldResponse.ok) {
        const fieldResult = await fieldResponse.json();
        const internalName = fieldResult.d?.InternalName || fieldResult.InternalName || cf.label;
        fieldMap[cf.id] = internalName;
        customFieldViewNames.push(internalName);
      }
    } catch {
      console.warn('[DEX] Custom Field konnte nicht angelegt werden:', cf.label);
    }
  }

  // FieldMap wird als Rückgabewert an den Caller zurückgegeben

  // Default View komplett neu aufbauen (Basis + Custom Fields). Mit rebuild:true
  // werden alle SP-Default-Spalten (Modified, Created, ID, Type, Compliance Asset,
  // App Created By, ...) aus der View rausgeworfen — nur funktionelle Felder.
  await svc.configureDefaultView(REG_LIST_NAME, [
    'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'ParticipantEmail', 'Department', 'Location', 'JobTitle', 'Company', 'Phone', 'StarterType', 'PreferredStarterType', 'Status', 'RegistrationDate', 'RegisteredByName', 'RegisteredByEmail', 'ProxyConsent', 'CancellationDate', 'CancelledByName', 'CancelledByEmail',
    ...customFieldViewNames,
    // v11.82: Team-Spalten am Ende der View (nach allen Custom Fields, vor
    // System-Spalten). So bleibt die View bei Nicht-Team-Events unauffällig
    // und bei Team-Events sieht der Organizer auf einen Blick, wer mit wem
    // angemeldet ist.
    'TeamId', 'TeamLead', 'TeamName',
  ], subsiteUrl, { rebuild: true });

  // Item-Level Permissions
  await setItemLevelPermissions(svc, subsiteUrl);

  // Berechtigungen
  await setRegistrationListPermissions(svc, subsiteUrl, organizerEmail);

  // v7.28: Counter-Liste für atomare TeilnehmerID-Vergabe anlegen
  // (Race-Condition-Schutz bei parallelen Anmeldungen).
  try {
    await svc.ensureCounterList(subsiteUrl);
  } catch {
    // Nicht kritisch — falls das schiefgeht, fallback auf max+1 in upsertParticipant.
  }

  return fieldMap;
}

/**
 * v9.35: Berechtigungs-Sync für nachträglich hinzugefügte Organizer/Co-Organizer.
 *
 * Wird im Wizard im Edit-Modus nach updateEvent aufgerufen. Geht über die
 * komma-/semikolon-separierte Liste aller Organizer-Mails und stellt sicher,
 * dass jede Person Full Control auf der Subsite + auf der Teilnehmerliste hat.
 *
 * Idempotent: Personen, die bereits Full Control haben, werden von SharePoints
 * `addroleassignment` einfach durchgereicht (kein Fehler, kein Doppel-Eintrag).
 * Existierende Item-Level-Permissions auf der Liste bleiben unangetastet — wir
 * brechen die Inheritance hier NICHT erneut, sondern fügen nur fehlende Principals
 * obendrauf hinzu.
 */
export async function ensureOrganizerPermissions(svc: EventService, subsiteUrl: string, organizerEmails: string): Promise<void> {
  await svc.ensureOrganizerPermissionsMulti([subsiteUrl], organizerEmails);
}

/**
 * v30.37: Derselbe Sync über MEHRERE Subsites — Klammer **und** alle
 * Sub-Events.
 *
 * Warum das nötig war: `ensureOrganizerPermissions` lief im Wizard nur über
 * `editEvent.subsiteUrl`, also die Klammer. Jedes Sub-Event hat aber eine
 * EIGENE Subsite mit eigener Teilnehmerliste. Wer nachträglich als
 * Co-Organizer dazukam, bekam Full Control auf der Klammer und auf keinem
 * einzigen Termin — und weil `getAllRegistrations` bei 403 nicht wirft,
 * sondern `[]` liefert, sah diese Person überall „0 Teilnehmer" statt einer
 * Fehlermeldung. Ein Klammer-Event mit 19 Office-Tagen war damit für sie
 * vollständig leer (Befund 31.08.2026, Carolin R.).
 *
 * Die User-Id wird EINMAL je Person aufgelöst, nicht je Subsite — sonst sind
 * es bei 19 Terminen × 3 Organizern 57 Lookups gegen dieselben drei Konten.
 *
 * @returns Zähler fürs UI: wie viele Zuweisungen BESTÄTIGT wurden, wie viele
 *   Personen gar nicht auflösbar waren und (v30.67) welche Zuweisungen
 *   SharePoint abgelehnt hat. Ein „schon vorhanden" zählt als Erfolg —
 *   SharePoints `addroleassignment` ist idempotent und antwortet dann 200.
 */
/**
 * v30.67: Ergebnis von `ensureOrganizerPermissionsMulti`. `failed` führt jede
 * Zuweisung auf, die SharePoint NICHT mit 2xx beantwortet hat (`status` 0 =
 * Netzwerkfehler). `grants` zählt nur bestätigte Zuweisungen — Web-Root UND
 * Teilnehmerliste je Person und Subsite, also bis zu zwei je Paar.
 */
export interface OrganizerPermissionFailure {
  site: string;
  userId: number;
  /** 'web' = Rechte auf der Subsite selbst, 'list' = auf der Teilnehmerliste. */
  scope: 'web' | 'list';
  status: number;
}
export interface OrganizerPermissionsResult {
  sites: number;
  users: number;
  grants: number;
  unresolved: string[];
  failed: OrganizerPermissionFailure[];
}
export async function ensureOrganizerPermissionsMulti(
  svc: EventService,
  subsiteUrls: string[],
  organizerEmails: string
): Promise<OrganizerPermissionsResult> {
  const sites = (subsiteUrls || []).map(s => (s || '').trim()).filter(Boolean);
  const emails = (organizerEmails || '').split(/[;,]/).map(s => s.trim()).filter(Boolean);
  const result: OrganizerPermissionsResult = { sites: sites.length, users: 0, grants: 0, unresolved: [], failed: [] };
  if (sites.length === 0 || emails.length === 0) return result;

  // 1) Personen auflösen — einmal, nicht je Subsite.
  const userIds: number[] = [];
  for (const em of emails) {
    try {
      const userResponse = await svc._sp.get(
        `${svc.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(em)}')?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (!userResponse.ok) { result.unresolved.push(em); continue; }
      const userData = await userResponse.json();
      const userId = userData.d?.Id || userData.Id;
      if (userId) userIds.push(userId); else result.unresolved.push(em);
    } catch { result.unresolved.push(em); }
  }
  result.users = userIds.length;

  // 2) Je Subsite: Web-Level + Teilnehmerliste. Beides einzeln gekapselt —
  //    eine recycelte Subsite darf die übrigen Termine nicht abbrechen.
  //
  // v30.67: Beide Antworten werden GEPRÜFT. `svc._post` wirft bei HTTP-
  // Fehlern nicht — der `try/catch` fing nur Netzwerkfehler, `grants++` lief
  // bei 403 (kein „Manage Permissions" auf dem Termin) und 404 (Subsite
  // recycelt) genauso, und die Listen-Zuweisung wurde gar nicht gezählt. Die
  // Kachel „Organizer-Berechtigungen reparieren" meldete deshalb grün
  // „3 Person(en) auf 20 Liste(n) berechtigt", während die Co-Organizerin
  // weiter überall 0 Teilnehmer sah. Ein bereits vorhandenes Recht
  // beantwortet SharePoint mit 200 — Erfolg und Scheitern sind also
  // unterscheidbar.
  const grant = async (site: string, userId: number, scope: 'web' | 'list', url: string): Promise<void> => {
    let status = 0;
    try {
      const r = await svc._post(url, {});
      status = r.status;
      if (r.ok) { result.grants++; return; }
    } catch { status = 0; }
    result.failed.push({ site, userId, scope, status });
  };
  for (const site of sites) {
    for (const userId of userIds) {
      await grant(site, userId, 'web', `${site}/_api/web/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`);
      await grant(site, userId, 'list', `${site}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`);
    }
  }
  if (result.failed.length > 0) {
    // Eine Sammel-Warnung statt einer je Zuweisung — bei 19 Terminen × 3
    // Personen wären es sonst über hundert Zeilen.
    const lines = result.failed.map(f => `${f.site} [${f.scope}] user ${f.userId} → HTTP ${f.status}`);
    console.warn(`[DEX] ensureOrganizerPermissionsMulti: ${result.failed.length} Zuweisung(en) NICHT gesetzt:\n${lines.join('\n')}`);
  }
  return result;
}

/**
 * Subsite-Berechtigungen: Owners Full Control, Members Read (damit User die Subsite betreten können).
 */
export async function setSubsitePermissions(svc: EventService, subsiteUrl: string, organizerEmail: string): Promise<void> {
  try {
    // Owners der Hauptsite: Full Control auf der Subsite
    const ownersResponse = await svc._sp.get(
      `${svc.siteUrl}/_api/web/associatedownergroup?$select=Id`,
      SPHttpClient.configurations.v1
    );
    if (ownersResponse.ok) {
      const ownersData = await ownersResponse.json();
      const ownersId = ownersData.d?.Id || ownersData.Id;
      await svc._post(
        `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${ownersId}, roledefid=1073741829)`,
        {}
      );
    }

    // Visitors der Hauptsite: Read auf der Subsite (damit User die Subsite betreten können)
    const visitorsId = await svc.getVisitorsGroupId();
    if (visitorsId) {
      await svc._post(
        `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741826)`,
        {}
      );
    }

    // Organizer: Full Control auf der Subsite. v9.18: organizerEmail kann
    // ";"-separiert mehrere Emails enthalten — Hauptorganizer + Co-Organizer
    // bekommen alle Full Control auf der Subsite.
    if (organizerEmail) {
      const emails = organizerEmail.split(/[;,]/).map(s => s.trim()).filter(Boolean);
      for (const em of emails) {
        try {
          const userResponse = await svc._sp.get(
            `${svc.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(em)}')?$select=Id`,
            SPHttpClient.configurations.v1
          );
          if (userResponse.ok) {
            const userData = await userResponse.json();
            const userId = userData.d?.Id || userData.Id;
            await svc._post(
              `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
              {}
            );
          }
        } catch { /* Organizer-Berechtigung optional */ }
      }
    }
  } catch {
    console.warn('[DEX] Subsite-Berechtigungen konnten nicht gesetzt werden');
  }
}

/**
 * Item-Level Permissions auf der Teilnehmerliste setzen.
 */
async function setItemLevelPermissions(svc: EventService, subsiteUrl: string): Promise<void> {
  // v26.87: zuverlässiger nometadata-MERGE (verbose+__metadata → HTTP 400
  // unter SPFx odata-version 3.0; die Teilnehmerlisten waren dadurch bislang
  // ungeschützt auf 1/1 statt „nur eigene Elemente" 2/2).
  await svc._setListSecurity(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')`, { ReadSecurity: 2, WriteSecurity: 2 });
}

/**
 * v20.5: Setzt nachträglich den Autor (Created By / SharePoint-Ersteller)
 * einer Teilnehmer-Zeile auf den TEILNEHMER selbst.
 *
 * Hintergrund: Die Teilnehmerlisten laufen mit Item-Level-Security
 * (ReadSecurity=2 / WriteSecurity=2) — ein User darf nur Items LESEN und
 * BEARBEITEN, die ER ERSTELLT hat (geprüft am Autor, NICHT am Feld
 * ParticipantEmail). Bei einer stellvertretenden Anmeldung (Organizer/Admin
 * meldet eine andere Person an) wäre sonst der Akteur der Autor — der
 * angemeldete Teilnehmer sähe seine eigene Anmeldung NICHT in "Meine
 * Events" und könnte sich nicht selbst abmelden. Indem der Teilnehmer zum
 * Autor wird, bekommt er Lese- + Abmelde-Zugriff auf SEINE Zeile.
 *
 * Best-effort: das Umsetzen von AuthorId erfordert "Listen verwalten" /
 * Full Control auf der Liste. Organizer (eigenes Event) und Admin haben das;
 * ein normaler Contribute-User (z.B. eine Assistenz) NICHT — dort schlägt
 * der MERGE mit 403 fehl und wird still ignoriert (die Zeile bleibt beim
 * Akteur als Autor). Der eigentliche Akteur ist ohnehin separat im Feld
 * RegisteredByEmail protokolliert, der Audit-Nachweis bleibt also erhalten.
 */
/**
 * v24.40: Eine Teilnehmer-Zeile einer **Assistenz** zuordnen — der Admin
 * übergibt damit die Verwaltung der (Fremd-)Anmeldung an eine bestimmte
 * Assistenz. Setzt ZWEI Dinge:
 *  1. `RegisteredByEmail`/`RegisteredByName` auf die Assistenz (Audit + Filter
 *     der „Assistenz"-Kachel).
 *  2. Den **Zeilen-Autor** (`Created By` / `AuthorId`) auf die Assistenz —
 *     unter Item-Level-Security („nur eigene Elemente") ist das die
 *     Voraussetzung, damit eine NORMALE Assistenz die Zeile überhaupt
 *     lesen/bearbeiten darf (sonst sieht sie sie in ihrer Kachel nicht).
 * Best-effort: Schlägt der Autor-Wechsel mangels Rechten fehl, landet er in
 * der `DEX_AccessFix`-Queue (Flow setzt ihn nach). Gibt zurück, ob der
 * RegisteredBy-Schreibvorgang gelang.
 */
export async function assignRegistrationToAssistant(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  assistantEmail: string,
  assistantName: string
): Promise<boolean> {
  try {
    const merge = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
      { 'RegisteredByEmail': assistantEmail, 'RegisteredByName': assistantName }
    );
    // Zeilen-Autor auf die Assistenz (Voraussetzung für ILS-Lesezugriff).
    await svc.trySetItemAuthor(subsiteUrl, REG_LIST_NAME, itemId, assistantEmail);
    return merge.ok;
  } catch (err) {
    console.warn('[DEX] assignRegistrationToAssistant error:', err);
    return false;
  }
}

export async function trySetItemAuthor(svc: EventService, subsiteUrl: string, listName: string, itemId: number, participantEmail: string): Promise<void> {
  try {
    // 1. Teilnehmer als SP-User der Subsite sicherstellen + dessen Id holen.
    const ensureResp = await svc._sp.post(
      `${subsiteUrl}/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
        },
        body: JSON.stringify({ logonName: participantEmail }),
      }
    );
    if (!ensureResp.ok) {
      // v20.7: z.B. Assistenz ohne ausreichende Rechte auf der Subsite →
      // Auftrag in die DEX_AccessFix-Queue, der Flow setzt den Autor.
      await svc.queueAccessFix(subsiteUrl, itemId, participantEmail);
      return;
    }
    const u = await ensureResp.json();
    const userId: number = u?.Id || u?.d?.Id || 0;
    if (!userId) {
      await svc.queueAccessFix(subsiteUrl, itemId, participantEmail);
      return;
    }
    // 2. AuthorId der Zeile auf den Teilnehmer setzen (nometadata-MERGE,
    //    daher KEIN __metadata im Body). 403 = fehlende Rechte → Queue.
    const m = await svc._merge(
      `${subsiteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`,
      { 'AuthorId': userId }
    );
    if (!m.ok) {
      // v20.7: typischer Assistenz-Fall — Contribute reicht nicht, um den
      // Autor zu setzen (braucht "Listen verwalten"). Flow übernimmt.
      await svc.queueAccessFix(subsiteUrl, itemId, participantEmail);
    }
  } catch {
    // Best-effort: auch hier den Flow-Auftrag versuchen — scheitert auch
    // der, bleibt die Zeile beim Akteur (Verhalten wie vor v20.5).
    try { await svc.queueAccessFix(subsiteUrl, itemId, participantEmail); } catch { /* */ }
  }
}

/**
 * v20.6: Reparatur-Werkzeug (Admin) — prüft EINE Teilnehmerliste und
 * repariert den Zugriff bei Fremd-Anmeldungen. Zwei Schritte:
 *
 * 1. **Item-Level-Security verifizieren:** liest ReadSecurity/WriteSecurity
 *    der Liste. Steht sie NICHT auf 2/2 („nur eigene Elemente" — z.B. weil
 *    der Set beim Anlegen still fehlschlug, siehe Security-Audit v20.x),
 *    wird sie neu gesetzt und per Read-back verifiziert.
 * 2. **Fremd-Anmeldungen (Anmeldung durch Dritte):** lädt alle Items mit
 *    Autor und setzt bei jedem Item, dessen `RegisteredByEmail` von der
 *    `ParticipantEmail` abweicht UND dessen Autor noch nicht der Teilnehmer
 *    ist, den Autor auf den Teilnehmer (`ensureuser` → `AuthorId`-MERGE,
 *    pro E-Mail gecacht). Damit sieht die angemeldete Person ihre eigene
 *    Zeile in „Meine Events" und kann sich selbst abmelden (v20.5-Logik,
 *    rückwirkend für Bestands-Anmeldungen).
 *
 * Läuft sequentiell (SP-Throttling-Schonung). `onProgress` meldet den
 * Item-Fortschritt für die UI. Externe Teilnehmer (kein Tenant-Login)
 * scheitern am `ensureuser` und landen in `authorFailed` — erwartbar.
 */
export async function repairProxyRegistrationAccess(
  svc: EventService,
  subsiteUrl: string,
  onProgress?: (done: number, total: number) => void
): Promise<{ ilsWasWrong: boolean; ilsFixed: boolean; itemsTotal: number; proxyFound: number; authorFixed: number; authorFailed: number }> {
  const result = { ilsWasWrong: false, ilsFixed: false, itemsTotal: 0, proxyFound: 0, authorFixed: 0, authorFailed: 0 };

  // ---- Schritt 1: Listen-Sicherheit „nur eigene Elemente" sicherstellen ----
  // v21 FIX: Der v20.6-Check meldete fälschlich ALLE Listen als unsicher
  // („27 falsch, 0 repariert"), obwohl die ILS nachweislich aktiv war
  // (Fremd-Zeilen unsichtbar). Ursache: Antwortformat/Typ der
  // ReadSecurity-Property nicht deterministisch behandelt. Jetzt: explizit
  // nometadata anfordern, Werte hart zu Zahlen koerzieren und bei
  // Unplausibilität die ROHE Antwort loggen statt „falsch" zu raten —
  // nur ein KLARER numerischer Wert ungleich 2 zählt als unsicher.
  const readSecurity = async (): Promise<{ rs: number; ws: number } | null> => {
    try {
      const resp = await svc._sp.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')?$select=ReadSecurity,WriteSecurity`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );
      if (!resp.ok) return null;
      const d = await resp.json();
      const rawRs = d.ReadSecurity ?? d.d?.ReadSecurity;
      const rawWs = d.WriteSecurity ?? d.d?.WriteSecurity;
      const rs = Number(rawRs);
      const ws = Number(rawWs);
      if (!Number.isFinite(rs) || !Number.isFinite(ws)) {
        console.warn('[DEX][ILS-Check] ReadSecurity/WriteSecurity nicht lesbar — rohe Antwort:', subsiteUrl, JSON.stringify(d).slice(0, 400));
        return null;
      }
      return { rs, ws };
    } catch (e) {
      console.warn('[DEX][ILS-Check] Lesen fehlgeschlagen:', subsiteUrl, e);
      return null;
    }
  };
  const before = await readSecurity();
  if (before && (before.rs !== 2 || before.ws !== 2)) {
    console.warn(`[DEX][ILS-Check] Liste meldet ReadSecurity=${before.rs}/WriteSecurity=${before.ws} (erwartet 2/2):`, subsiteUrl);
    result.ilsWasWrong = true;
    await setItemLevelPermissions(svc, subsiteUrl);
    const after = await readSecurity();
    result.ilsFixed = !!after && after.rs === 2 && after.ws === 2;
    if (!result.ilsFixed) {
      console.warn('[DEX][ILS-Check] Read-back nach Fix weiterhin abweichend:', subsiteUrl, after);
    }
  }

  // ---- Schritt 2: Items mit Autor laden (paged) ----
  type Row = { Id: number; ParticipantEmail?: string; RegisteredByEmail?: string; Author?: { EMail?: string } };
  const items: Row[] = [];
  let url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,ParticipantEmail,RegisteredByEmail,Author/EMail&$expand=Author&$top=500`;
  while (url) {
    const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (!resp.ok) break;
    const data = await resp.json();
    const arr: Row[] = data.value || data.d?.results || [];
    items.push(...arr);
    url = data['odata.nextLink'] || data['@odata.nextLink'] || data.d?.__next || '';
  }
  result.itemsTotal = items.length;

  // ---- Schritt 3: Fremd-Anmeldungen → Autor auf Teilnehmer setzen ----
  const userIdCache: Record<string, number> = {};
  let done = 0;
  for (const it of items) {
    done++;
    const pe = (it.ParticipantEmail || '').toLowerCase().trim();
    const rb = (it.RegisteredByEmail || '').toLowerCase().trim();
    const au = (it.Author?.EMail || '').toLowerCase().trim();
    // Nur Fremd-Anmeldungen: RegisteredByEmail vorhanden UND != Teilnehmer.
    // (Alt-Bestand ohne RegisteredByEmail = vor v3.x — dort ist der Autor
    // ohnehin der Teilnehmer selbst, weil es nur Selbst-Anmeldung gab.)
    if (!pe || !rb || pe === rb) { if (onProgress) onProgress(done, items.length); continue; }
    result.proxyFound++;
    // Autor stimmt schon (z.B. v20.5-Anmeldung oder früherer Lauf) → ok.
    if (au === pe) { if (onProgress) onProgress(done, items.length); continue; }
    try {
      let uid = userIdCache[pe] || 0;
      if (!uid) {
        const er = await svc._sp.post(
          `${subsiteUrl}/_api/web/ensureuser`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
            },
            body: JSON.stringify({ logonName: pe }),
          }
        );
        if (er.ok) {
          const u = await er.json();
          uid = u?.Id || u?.d?.Id || 0;
          if (uid) userIdCache[pe] = uid;
        }
      }
      if (!uid) {
        result.authorFailed++;
        if (onProgress) onProgress(done, items.length);
        continue;
      }
      const m = await svc._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${it.Id})`,
        { 'AuthorId': uid }
      );
      if (m.ok) result.authorFixed++; else result.authorFailed++;
    } catch {
      result.authorFailed++;
    }
    if (onProgress) onProgress(done, items.length);
  }
  return result;
}

/**
 * Berechtigungen für Teilnehmerliste auf der Subsite setzen.
 */
async function setRegistrationListPermissions(svc: EventService, subsiteUrl: string, organizerEmail: string): Promise<void> {
  try {
    await svc._post(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
      {}
    );

    // Site Owners der Hauptsite: Full Control
    const ownersResponse = await svc._sp.get(
      `${svc.siteUrl}/_api/web/associatedownergroup?$select=Id`,
      SPHttpClient.configurations.v1
    );
    if (ownersResponse.ok) {
      const ownersData = await ownersResponse.json();
      await svc._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${ownersData.Id}, roledefid=1073741829)`,
        {}
      );
    }

    // Visitors: Contribute (damit User sich registrieren können)
    const visitorsId = await svc.getVisitorsGroupId();
    if (visitorsId) {
      await svc._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741827)`,
        {}
      );
    }

    // Organizer: Full Control. v9.18: organizerEmail kann ";"-separiert
    // mehrere Emails enthalten (Hauptorganizer + Co-Organizer).
    if (organizerEmail) {
      const emails = organizerEmail.split(/[;,]/).map(s => s.trim()).filter(Boolean);
      for (const em of emails) {
        try {
          const userResponse = await svc._sp.get(
            `${svc.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(em)}')?$select=Id`,
            SPHttpClient.configurations.v1
          );
          if (userResponse.ok) {
            const userData = await userResponse.json();
            await svc._post(
              `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${userData.Id}, roledefid=1073741829)`,
              {}
            );
          }
        } catch { /* Organizer-Berechtigung optional */ }
      }
    }
  } catch {
    // Berechtigungen konnten nicht gesetzt werden
  }
}
