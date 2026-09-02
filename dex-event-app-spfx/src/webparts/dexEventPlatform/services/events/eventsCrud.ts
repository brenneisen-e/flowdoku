/**
 * v30.66 — Modularisierung Stufe 2: Thema „Event-CRUD" — die Zeilen in
 * DEX_Events lesen, anlegen (samt Subsite und Teilnehmerliste), ändern und
 * löschen.
 *
 * Zwei Dinge, die hier hängen (siehe CLAUDE.md): `EndDate` fällt zentral auf
 * `StartDate` zurück, weil der Outlook-Flow bei leerem Ende abbricht — und
 * der Flow triggert nur auf NEUE Listeneinträge, ein MERGE stösst ihn nie an.
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import { normalizeMadeWithLink } from '../EmailTemplates';
import { buildOutlookLocation } from '../../utils/eventFormat';
import { dlog } from '../../utils/debugLog';
import type { EventService, CustomField, SPEvent } from '../EventService';
import { REG_LIST_NAME } from '../EventService';

// v30.66: war `private static readonly` an der Klasse — die Spaltenliste wird
// nur von den Lese-Methoden dieses Themas gebraucht.
const EVENT_SELECT = 'Id,Title,EventStatus,EventNumber,Description,Location,LocationAddress,LocationFilter,Audience,AudienceResolvedEmails,FilterMode,StartDate,EndDate,RegistrationDeadline,LastDeregisterDate,MaxParticipants,CurrentParticipants,WaitlistEnabled,MandatoryRegistration,EventImageUrl,EmailImageBase64,Organizer,OrganizerEmail,ContactName,ContactEmail,ContactOrganizerEmail,ContactInfo,OutlookEventId,CalendarLink,OutlookBody,OutlookSubject,OutlookStart,OutlookEnd,OutlookLocation,AllDay,ShowAsFree,SkipOrganizerInvite,EmailLanguage,RegistrationLanguage,EmailTemplateOverrides,DisableEmails,DisableRegistrationEmail,DisableCancellationEmail,AutoDeregisterOnDecline,InactiveHandling,DisableOutlook,OutlookDirty,AutoSendQRCode,ActiveFrom,NotifyOrgRegisterMode,NotifyOrgRegisterFromDate,NotifyOrgCancelMode,ExcludedUsers,IsFictive,DurchstarterCapacity,FunstarterCapacity,SplitLabelA,SplitLabelB,SplitDescA,SplitDescB,SplitHelpText,SplitSectionTitle,SplitSharedWaitlist,AllowAttendeeUpload,AttendeeUploadHint,AttendeeUploadLabel,AskSalutation,ConfirmDialogEnabled,ConfirmDialogMode,ConfirmDialogText,SelfCheckInEnabled,SelfCheckInToken,SelfCheckInFrom,SelfCheckInTo,TeamRegistrationEnabled,TeamSize,AskTeamName,TeamPartialAllowed,TeamOpenSlotsVisible,TeamJoinRequiresApproval,BilingualFields,CustomFields,Agenda,Transfers,Documents,FunZone,QuizClusterSize,ParentEventId,RegistrationListName,SubsiteUrl,Modified,Created';

/**
 * Seed-Events anlegen falls sie nicht existieren (einmalig beim ersten Start).
 */
export async function seedEvents(svc: EventService): Promise<void> {
  try {
    // Prüfen ob "Assistenz Meeting 2026" schon existiert
    const check = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=Title eq 'Assistenz Meeting 2026'&$top=1&$select=Id`,
      SPHttpClient.configurations.v1
    );
    if (check.ok) {
      const data = await check.json();
      const items = data.value || data.d?.results || [];
      if (items.length > 0) return; // Existiert bereits
    }

    // Event anlegen
    await svc.createEvent({
      title: 'Assistenz Meeting 2026',
      type: 'Other',
      status: 'Active',
      description: 'Assistenz Meeting Mai 2026 - Frankfurt am Main',
      location: 'Frankfurt am Main',
      locationFilter: '',
      audience: 'All',
      filterMode: 'OR',
      startDate: '2026-05-07T11:00:00.000Z',
      endDate: '2026-05-08T15:00:00.000Z',
      registrationDeadline: '2026-04-09T00:00:00.000Z',
      lastDeregisterDate: '',
      maxParticipants: 130,
      waitlistEnabled: true,
      eventImageUrl: '',
      organizer: 'Maerzluft, Petra; Schwartz, Eva',
      organizerEmail: 'pmaerzluft@deloitte.de',
      outlookEventId: '',
      outlookBody: '',
      emailLanguage: 'EN',
      emailTemplateOverrides: '',
      customFields: [
        { id: 'travel', label: 'You will travel with?', type: 'select', required: false, visible: true, options: ['Train', 'Car', 'Public Transport'] },
        { id: 'deutschlandticket', label: 'Do you own a Deutschlandticket?', type: 'select', required: false, visible: true, options: ['Yes', 'No'] },
        { id: 'expenses', label: 'Please insert the total amount of your travel expenses!', type: 'text', required: false, visible: true },
      ],
      agenda: '[]',
      transfers: '[]',
      documents: '[]',
    });
  } catch { /* Seed fehlgeschlagen - nicht kritisch */ }
}

/**
 * Alle Events laden
 */
export async function getEvents(svc: EventService): Promise<SPEvent[]> {
  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EVENT_SELECT}&$orderby=StartDate desc&$top=100`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return [];
    // v29.51 (Messpunkt): Das ist die EINZIGE blockierende Datenabfrage des
    // Starts — und EVENT_SELECT holt 79 Spalten, darunter EmailImageBase64
    // und EmailTemplateOverrides mit eingebetteten Bildern. Ob das ein paar
    // Kilobyte oder mehrere Megabyte sind, entscheidet über den nächsten
    // Optimierungsschritt; bisher wurde darüber geraten. `.text()` +
    // JSON.parse ist genau das, was `.json()` intern auch tut — der Umweg
    // kostet nichts und liefert die exakte Byte-Zahl.
    const raw = await response.text();
    const t0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : 0;
    const data = JSON.parse(raw);
    const parseMs = t0 ? Math.round(performance.now() - t0) : -1;
    const rows = data.value || [];
    dlog('perf',
      `[DEX][perf][getEvents] ${rows.length} Events · ${Math.round(raw.length / 1024)} KB JSON · parse ${parseMs} ms`
    );
    return rows;
  } catch {
    return [];
  }
}

/**
 * Einzelnes Event laden
 */
export async function getEvent(svc: EventService, eventId: number): Promise<SPEvent | null> {
  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=${EVENT_SELECT}`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return null;
    return await response.json();
  } catch {
    return null;
  }
}

/**
 * v18.33: Event anhand des Self-Check-in-Tokens finden (für den statischen
 * Check-in-Link ?action=selfcheckin&token=…). Liefert das erste Event mit
 * passendem Token. Alle eingeloggten User dürfen DEX_Events lesen.
 */
export async function getEventBySelfCheckInToken(svc: EventService, token: string): Promise<SPEvent | null> {
  try {
    const safe = token.replace(/'/g, "''");
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EVENT_SELECT}&$filter=SelfCheckInToken eq '${safe}'&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return null;
    const data = await response.json();
    const arr = data.value || (data.d && data.d.results) || [];
    return arr.length > 0 ? arr[0] : null;
  } catch {
    return null;
  }
}

/**
 * v18.33: Event anhand der Event-Nummer finden (für den rotierenden Live-QR
 * ?action=selfcheckin&event=<Nr>&code=…&t=…).
 */
export async function getEventByEventNumber(svc: EventService, eventNumber: number): Promise<SPEvent | null> {
  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EVENT_SELECT}&$filter=EventNumber eq ${eventNumber}&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return null;
    const data = await response.json();
    const arr = data.value || (data.d && data.d.results) || [];
    return arr.length > 0 ? arr[0] : null;
  } catch {
    return null;
  }
}

/**
 * Neues Event erstellen + Subsite mit Teilnehmerliste anlegen
 */
export async function createEvent(svc: EventService, event: {
  title: string;
  status: string;
  type: string;
  description: string;
  location: string;
  locationAddress?: string; // JSON-String: { street, houseNo, zip, city }
  outlookSubject?: string; // v18.42: Betreff des Outlook-Termins (leer = Titel)
  outlookStart?: string; // v18.44: abweichende Start-Zeit (ISO, leer = Event-Start)
  outlookEnd?: string;   // v18.44: abweichende End-Zeit (ISO, leer = Event-Ende)
  outlookLocation?: string; // v18.40: manueller Outlook-Ort (leer = Auto aus Ort + Adresse)
  allDay?: boolean; // v29.52: ganztägiger Termin (Flow setzt daraus isAllDay)
  showAsFree?: boolean; // v29.54: Termin als „Frei" anzeigen (Flow: showAs)
  skipOrganizerInvite?: boolean; // v29.55: Organizer nicht einladen (Flow: requiredAttendees)
  outlookIsOnlineMeeting?: boolean; // v30.26: Termin als Teams-Besprechung anlegen (Flow: isOnlineMeeting)
  locationFilter: string;
  audience: string;
  /** v16.4: Vor-aufgelöste E-Mails der Audience-DLs, ';'-separiert, lowercase. */
  audienceResolvedEmails?: string;
  filterMode: string;
  startDate: string;
  endDate: string;
  registrationDeadline: string;
  lastDeregisterDate: string;
  /** v29.19: Auto-Aktivierungszeitpunkt (UTC-ISO). Der Wizard bot das Feld
   *  auch beim ANLEGEN an, persistiert wurde es aber nur im Edit-Pfad —
   *  ein als Entwurf angelegtes Event mit „Aktiv ab" ging nie von allein
   *  live. */
  activeFrom?: string;
  maxParticipants: number;
  waitlistEnabled: boolean;
  mandatoryRegistration?: boolean; // v24.64: Pflicht-Sub-Event

  eventImageUrl: string;
  organizer: string;
  organizerEmail: string;
  /** v10.16: optionaler Ansprechpartner (Anzeige-Feld). */
  contactName?: string;
  contactEmail?: string;
  contactOrganizerEmail?: string;
  contactInfo?: string;
  outlookEventId: string;
  outlookBody: string;
  agenda?: string; // JSON-Array mit Agenda-Einträgen
  transfers?: string; // JSON-Array mit Transferzeiten
  documents?: string; // JSON-Array mit Dokumenten
  funZone?: string; // JSON-Array mit Quiz-Fragen
  quizClusterSize?: number; // 1..4 - Fragen pro Quiz-Ansicht
  /** Seit v6.4: wenn gesetzt, wird dieses Event als Sub-Event angelegt und zeigt auf das angegebene Parent-Event. */
  parentEventId?: string;
  emailLanguage?: string;
  registrationLanguage?: 'de' | 'en';
  emailTemplateOverrides?: string;
  disableEmails?: boolean;
  disableRegistrationEmail?: boolean;
  disableCancellationEmail?: boolean;
  autoDeregisterOnDecline?: boolean;
  inactiveHandling?: string;
  disableOutlook?: boolean;
  notifyOrgRegisterMode?: 'never' | 'always' | 'fromDate';
  notifyOrgRegisterFromDate?: string;
  notifyOrgCancelMode?: 'never' | 'always' | 'afterDeadline';
  excludedUsers?: string[];
  isFictive?: boolean;
  durchstarterCapacity?: number;
  funstarterCapacity?: number;
  splitLabelA?: string;
  splitLabelB?: string;
  splitDescA?: string;
  splitDescB?: string;
  splitHelpText?: string;
  splitSectionTitle?: string;
  splitSharedWaitlist?: boolean;
  allowAttendeeUpload?: boolean;
  attendeeUploadHint?: string;
  attendeeUploadLabel?: string;
  /** v11.80: Anrede im Registrierungsformular abfragen (Default false). */
  askSalutation?: boolean;
  /** v18.75: Sicherheitshinweis vor dem Absenden der Anmeldung. */
  confirmDialogEnabled?: boolean;
  confirmDialogMode?: string; // 'summary' | 'freetext'
  confirmDialogText?: string;
  /** v18.33: Self-Check-in per QR-Code erlauben (Default false). */
  selfCheckInEnabled?: boolean;
  /** v18.33: Geheimer Token (statischer Link + HMAC-Schlüssel rotierender QR). */
  selfCheckInToken?: string;
  /** v18.33: optionaler Start des Check-in-Fensters (ISO). */
  selfCheckInFrom?: string;
  /** v18.33: optionales Ende des Check-in-Fensters (ISO). */
  selfCheckInTo?: string;
  /** v11.80: Team-Anmeldung erlauben (Default false). */
  teamRegistrationEnabled?: boolean;
  /** v11.80: Maximale Teamgröße (0 = nicht gesetzt). */
  teamSize?: number;
  /** v11.80: Team-Name abfragen (Default false). */
  askTeamName?: boolean;
  /** v11.81: Auch Teil-Teams zulassen (Default false = nur komplette Teams). */
  teamPartialAllowed?: boolean;
  /** v11.81: Offene Slots öffentlich für Beitritt sichtbar (Default false). */
  teamOpenSlotsVisible?: boolean;
  /** v11.81: Beitritt erfordert Bestätigung durch Team-Kapitän (Default false). */
  teamJoinRequiresApproval?: boolean;
  /** v17.20: Custom-Fields zweisprachig anbieten (DE+EN). */
  bilingualFields?: boolean;
  customFields: CustomField[];
  /** v11.69: Wenn `existingSubsiteUrl` UND `existingRegistrationListName`
   *  gesetzt sind, wird KEINE neue Subsite und KEINE neue Teilnehmer-
   *  liste angelegt — stattdessen werden die mitgegebenen Werte direkt in
   *  das neue DEX_Events-Item geschrieben. Hintergrund: Outlook-Termin
   *  nachträglich aktivieren ohne Verlust bestehender Anmeldungen — das
   *  Sub-Event wird mit `deleteEventItemOnly()` aus DEX_Events entfernt
   *  und hier neu angelegt, wobei die alte Subsite + Teilnehmerliste
   *  unangetastet bleiben und an die neue Event-Zeile angehängt werden.
   *  Damit triggert der `DEX_CreateOutlookEvent`-Flow (GetOnNewItems) auf
   *  dem neuen Item und legt den Outlook-Termin an. */
  existingSubsiteUrl?: string;
  existingRegistrationListName?: string;
  /** v11.87: Optionaler Progress-Callback. Wird zu Beginn jeder Teil-
   *  Operation aufgerufen — die UI kann darauf den Fortschrittsbalken
   *  und die Unter-Caption sichtbar bewegen, statt minutenlang auf
   *  „Event wird vorbereitet..." stehen zu bleiben. Stages decken
   *  die langsamen SP-Operationen ab (Subsite-Create, Listen-Create,
   *  Permissions, Counter, Views). */
  onProgress?: (stage:
    | 'start'
    | 'subsite-creating'
    | 'subsite-done'
    | 'permissions'
    | 'list-creating'
    | 'list-done'
    | 'item-insert'
    | 'done'
  ) => void;
}): Promise<number | null> {
  const reportProgress = (stage:
    | 'start'
    | 'subsite-creating'
    | 'subsite-done'
    | 'permissions'
    | 'list-creating'
    | 'list-done'
    | 'item-insert'
    | 'done'
  ): void => {
    try { event.onProgress?.(stage); } catch { /* */ }
  };
  try {
    reportProgress('start');
    // 0. Nächste EventNumber ermitteln
    // v30.67: Ein Lesefehler ist KEIN leeres Ergebnis. Bisher diente
    // `nextEventNumber = 1` gleichzeitig als Startwert für „Liste ist leer"
    // und als stiller Fallback für „Abfrage gescheitert" (429/500/Netz) — ab
    // da trugen zwei Events dieselbe Nummer, und DEX_Participants, „Meine
    // Events", der QR-Code und deleteEvent schlüsseln alle über diese Nummer.
    // Deshalb: bei Fehler abbrechen, BEVOR irgendetwas angelegt wird (der
    // Wizard zeigt Errors aus createEvent an). Unmittelbar vor dem Insert wird
    // die Nummer noch einmal nachgelesen und nach dem Insert auf Eindeutigkeit
    // geprüft — zwischen erstem Lesen und Schreiben liegt die Subsite-Anlage
    // (oft > 30 s), zwei parallel anlegende Organizer zogen dieselbe Nummer.
    const readNextEventNumber = async (): Promise<number> => {
      let enResp;
      try {
        enResp = await svc._sp.get(
          `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=EventNumber&$orderby=EventNumber desc&$top=1`,
          SPHttpClient.configurations.v1
        );
      } catch (err) {
        throw new Error(`Event-Nummer konnte nicht ermittelt werden (${err instanceof Error ? err.message : 'Netzwerkfehler'}) — bitte erneut versuchen.`);
      }
      if (!enResp.ok) throw new Error(`Event-Nummer konnte nicht ermittelt werden (HTTP ${enResp.status}) — bitte erneut versuchen.`);
      const enData = await enResp.json();
      const items = enData.value || enData.d?.results || [];
      const top = items.length > 0 ? Number(items[0].EventNumber) : 0;
      return (Number.isFinite(top) && top > 0) ? top + 1 : 1;
    };
    let nextEventNumber = await readNextEventNumber();

    // v11.69: Reuse-Pfad — wenn `existingSubsiteUrl` UND
    // `existingRegistrationListName` mitgegeben wurden, überspringen wir
    // 1) Subsite-Anlegen, 2) Subsite-Permissions, 3) Teilnehmerliste
    // anlegen. Die mitgegebene Subsite bleibt unangetastet inkl. aller
    // Teilnehmer-Anmeldungen. Custom-Fields werden ohne spInternalName-
    // Anreicherung übernommen — die Felder existieren bereits auf der
    // alten Teilnehmerliste mit den korrekten Internal-Names.
    const reuseSubsite = !!(event.existingSubsiteUrl && event.existingRegistrationListName);
    let subsiteUrl: string;
    let enrichedCustomFields: CustomField[];
    const coOrgEmailsForPerm: string[] = (() => {
      try {
        const o = JSON.parse(event.emailTemplateOverrides || '{}');
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const list = (o as any)._coOrganizers;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        if (Array.isArray(list)) return list.map((x: any) => String(x?.email || '')).filter(Boolean);
      } catch { /* */ }
      return [];
    })();
    const allOrgEmails = [event.organizerEmail || '', ...coOrgEmailsForPerm].filter(Boolean).join(';');
    const regListName = reuseSubsite ? (event.existingRegistrationListName as string) : REG_LIST_NAME;

    if (reuseSubsite) {
      // v11.69: bestehende Subsite + Teilnehmerliste wiederverwenden.
      subsiteUrl = event.existingSubsiteUrl as string;
      // Custom-Fields unverändert übernehmen — die Liste existiert
      // bereits, kein neues Schema nötig.
      enrichedCustomFields = event.customFields.map(cf => ({ ...cf }));
    } else {
      // 1. Subsite für das Event erstellen
      reportProgress('subsite-creating');
      const createdSubsite = await svc.createEventSubsite(event.title, event.description);
      if (!createdSubsite) {
        console.error('[DEX] Subsite konnte nicht erstellt werden');
        throw new Error('Subsite konnte nicht erstellt werden. Fehlende Berechtigung? Bitte wende dich an einen Site-Administrator.');
      }
      subsiteUrl = createdSubsite;
      reportProgress('subsite-done');

      // 2. Subsite-Berechtigungen: Members der Parent-Site auf der Subsite berechtigen.
      // v9.18: Co-Organizer-Emails aus emailTemplateOverrides._coOrganizers extrahieren
      // und mit dem Hauptorganizer zusammen Full Control erteilen.
      reportProgress('permissions');
      await svc.setSubsitePermissions(subsiteUrl, allOrgEmails);

      // 3. Teilnehmerliste auf der Subsite erstellen
      reportProgress('list-creating');
      const fieldMap: Record<string, string> = await svc.createRegistrationList(subsiteUrl, event.customFields, allOrgEmails);
      reportProgress('list-done');

      // Custom Fields mit SP InternalName anreichern
      enrichedCustomFields = event.customFields.map(cf => ({
        ...cf,
        spInternalName: fieldMap[cf.id] || '',
      }));
    }

    // 3. Event in DEX_Events eintragen
    const payload = {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
      'Title': event.title,
      'EventNumber': nextEventNumber,
      'EventStatus': event.status,
      'Description': event.description,
      'Location': event.location,
      'LocationAddress': event.locationAddress || '',
      // v18.42: Outlook-Betreff (leer = Flow fällt auf Titel zurück via coalesce).
      'OutlookSubject': (event.outlookSubject && event.outlookSubject.trim()) ? event.outlookSubject.trim() : '',
      // v18.44: abweichendes Outlook-Datum (leer = Flow nutzt StartDate/EndDate).
      'OutlookStart': event.outlookStart || null,
      'OutlookEnd': event.outlookEnd || null,
      // v29.52: ganztägig. Start/Ende bleiben bewusst wie gesetzt (00:00/23:59)
      // — die Umrechnung auf die Ganztags-Grenzen macht der Flow. Solange der
      // Flow das Feld noch nicht liest, verhält sich alles wie bisher.
      'AllDay': !!event.allDay,
      'ShowAsFree': !!event.showAsFree, // v29.54
      'OutlookIsOnlineMeeting': !!event.outlookIsOnlineMeeting, // v30.26
      'SkipOrganizerInvite': !!event.skipOrganizerInvite, // v29.55
      // v18.34/v18.40: Outlook-Ort = manuelle Überschreibung, sonst
      // automatisch aus Veranstaltungsort + Adresse. Flow mappt OutlookLocation 1:1.
      // v26.54: hart auf 255 kappen (einzeilige Text-Spalte — s. updateEvent).
      'OutlookLocation': ((event.outlookLocation && event.outlookLocation.trim())
        ? event.outlookLocation.trim()
        : buildOutlookLocation(event.location, event.locationAddress)).slice(0, 255),
      'LocationFilter': event.locationFilter,
      'Audience': event.audience,
      'AudienceResolvedEmails': event.audienceResolvedEmails || '',
      'FilterMode': event.filterMode || 'OR',
      'StartDate': event.startDate || null,
      // v22.17/v28.66: EndDate darf NIE leer in DEX_Events landen — der
      // DEX_CreateOutlookEvent-Flow rechnet convertFromUtc(coalesce(
      // OutlookEnd, EndDate)); bei null stürzt „Create event (V4)" ab und es
      // entsteht kein Outlook-Termin. Die Aufrufer setzen den Fallback zwar
      // schon, hier wird er zentral erzwungen (letzte Instanz vor dem
      // Schreiben — s. auch updateEvent).
      'EndDate': event.endDate || event.startDate || null,
      'RegistrationDeadline': event.registrationDeadline || null,
      // v29.19: s. Interface — vorher nur im Edit-Pfad geschrieben.
      'ActiveFrom': event.activeFrom || null,
      'LastDeregisterDate': event.lastDeregisterDate || null,
      'MaxParticipants': event.maxParticipants,
      'WaitlistEnabled': event.waitlistEnabled,
      'MandatoryRegistration': !!event.mandatoryRegistration,
      'EventImageUrl': event.eventImageUrl,
      // Custom-Event-Logo aus emailTemplateOverrides._eventLogo extrahieren (falls
      // vorhanden) und als EmailImageBase64 persistieren — damit der Power-Automate-Flow
      // es als {{ORB_URL}} in Mail + Outlook-Termin einsetzt.
      'EmailImageBase64': (() => {
        try {
          const o = JSON.parse(event.emailTemplateOverrides || '{}');
          return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
        } catch { return ''; }
      })(),
      'Organizer': event.organizer,
      'OrganizerEmail': event.organizerEmail,
      // v10.16: optionaler Ansprechpartner (Anzeige-Feld). Strings können
      // leer sein — leer = kein Ansprechpartner gepflegt.
      'ContactName': event.contactName || '',
      'ContactEmail': event.contactEmail || '',
      'ContactOrganizerEmail': event.contactOrganizerEmail || '',
      'ContactInfo': event.contactInfo || '',
      'OutlookEventId': event.outlookEventId,
      // outlookBody kommt bereits vollständig gewickelt + mit aufgelösten Variablen
      // aus EventCreationPage — hier nur durchreichen.
      // v29.42: auch im Termin-Text die Fußzeile auf die kanonische Adresse.
      'OutlookBody': normalizeMadeWithLink(event.outlookBody || ''),
      'EmailLanguage': event.emailLanguage || 'EN',
      'RegistrationLanguage': event.registrationLanguage || '',
      'EmailTemplateOverrides': event.emailTemplateOverrides || '',
      'DisableEmails': !!event.disableEmails,
      'DisableRegistrationEmail': !!event.disableRegistrationEmail,
      'DisableCancellationEmail': !!event.disableCancellationEmail,
      'AutoDeregisterOnDecline': !!event.autoDeregisterOnDecline,
      'InactiveHandling': event.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify',
      'DisableOutlook': !!event.disableOutlook,
      'NotifyOrgRegisterMode': (() => {
        const m = event.notifyOrgRegisterMode || 'never';
        return m === 'always' ? 'Always' : m === 'fromDate' ? 'FromDate' : 'Never';
      })(),
      'NotifyOrgRegisterFromDate': event.notifyOrgRegisterFromDate || null,
      'NotifyOrgCancelMode': (() => {
        const m = event.notifyOrgCancelMode || 'never';
        return m === 'always' ? 'Always' : m === 'afterDeadline' ? 'AfterDeadline' : 'Never';
      })(),
      'ExcludedUsers': (event.excludedUsers || []).filter(Boolean).join(';'),
      'IsFictive': !!event.isFictive,
      'DurchstarterCapacity': typeof event.durchstarterCapacity === 'number' ? event.durchstarterCapacity : null,
      'FunstarterCapacity': typeof event.funstarterCapacity === 'number' ? event.funstarterCapacity : null,
      'SplitLabelA': event.splitLabelA || '',
      'SplitLabelB': event.splitLabelB || '',
      'SplitDescA': event.splitDescA || '',
      'SplitDescB': event.splitDescB || '',
      'SplitHelpText': event.splitHelpText || '',
      'SplitSectionTitle': event.splitSectionTitle || '',
      'SplitSharedWaitlist': !!event.splitSharedWaitlist,
      'AllowAttendeeUpload': !!event.allowAttendeeUpload,
      'AttendeeUploadHint': event.attendeeUploadHint || '',
      'AttendeeUploadLabel': event.attendeeUploadLabel || '',
      'AskSalutation': !!event.askSalutation,
      'ConfirmDialogEnabled': !!event.confirmDialogEnabled,
      'ConfirmDialogMode': event.confirmDialogMode || '',
      'ConfirmDialogText': event.confirmDialogText || '',
      'SelfCheckInEnabled': !!event.selfCheckInEnabled,
      'SelfCheckInToken': event.selfCheckInToken || '',
      'SelfCheckInFrom': event.selfCheckInFrom || null,
      'SelfCheckInTo': event.selfCheckInTo || null,
      'TeamRegistrationEnabled': !!event.teamRegistrationEnabled,
      'TeamSize': typeof event.teamSize === 'number' && event.teamSize > 0 ? event.teamSize : null,
      'AskTeamName': !!event.askTeamName,
      'TeamPartialAllowed': !!event.teamPartialAllowed,
      'TeamOpenSlotsVisible': !!event.teamOpenSlotsVisible,
      'TeamJoinRequiresApproval': !!event.teamJoinRequiresApproval,
      'BilingualFields': !!event.bilingualFields,
      'CustomFields': JSON.stringify(enrichedCustomFields),
      'Agenda': event.agenda || '[]',
      'Transfers': event.transfers || '[]',
      'Documents': event.documents || '[]',
      'FunZone': event.funZone || '[]',
      'QuizClusterSize': typeof event.quizClusterSize === 'number' ? event.quizClusterSize : null,
      'ParentEventId': event.parentEventId || '',
      'RegistrationListName': regListName,
      'RegistrationListUrl': `${subsiteUrl}/Lists/${regListName}/AllItems.aspx`,
      'SubsiteUrl': subsiteUrl,
    };

    reportProgress('item-insert');
    // v30.67: Nummer direkt vor dem Insert nachlesen (s. Kommentar oben).
    // Best-effort: Scheitert das Nachlesen, bleibt die zu Beginn geprüfte
    // Nummer — die Subsite steht schon, ein Abbruch hier hinterließe sie
    // verwaist. Im Reuse-Pfad (Recreate ohne Subsite-Anlage) liegen zwischen
    // beiden Lesevorgängen nur Millisekunden, dort entfällt der Request.
    if (!reuseSubsite) {
      try {
        const fresh = await readNextEventNumber();
        if (fresh > nextEventNumber) { nextEventNumber = fresh; payload.EventNumber = fresh; }
      } catch { /* zu Beginn geprüfte Nummer behalten */ }
    }
    // v28.10: gleicher 2-MB-Schutz wie in updateEvent — zu große Payloads
    // (eingebettete Logos/Bilder) sauber abfangen statt kryptischem 400.
    if (JSON.stringify(payload).length > 1_900_000) {
      throw new Error('Die Event-Daten überschreiten das SharePoint-Limit von 2 MB. Ursache ist fast immer ein zu großes eingebettetes Bild (Mail-Logo, Outlook-Kopfbild oder ein Bild im Mail-/Termin-Text). Bitte das Bild entfernen oder neu (kleiner) hochladen.');
    }
    const response = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`,
      payload
    );

    if (!response.ok) return null;
    const result = await response.json();
    const newItemId: number = result.d?.Id || result.Id;
    // v30.67: Eindeutigkeit nach dem Insert prüfen — analog zum Post-Insert-
    // Dedup der TeilnehmerID in registerForEvent. Haben zwei Organizer im
    // selben Fenster dieselbe Nummer gezogen, behält die ältere Zeile
    // (kleinere Id) die Nummer, unsere bekommt die nächste freie. Best-effort:
    // Ein Fehler hier darf das angelegte Event nicht mehr kippen, er wird
    // nur gemeldet.
    try {
      const dupResp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=Id&$filter=EventNumber eq ${nextEventNumber}&$top=10`,
        SPHttpClient.configurations.v1
      );
      if (dupResp.ok) {
        const dupData = await dupResp.json();
        const dupItems: Array<{ Id: number }> = dupData.value || dupData.d?.results || [];
        if (dupItems.length > 1 && newItemId !== Math.min(...dupItems.map(d => d.Id))) {
          const fresh = await readNextEventNumber();
          const fix = await svc._merge(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${newItemId})`, { 'EventNumber': fresh });
          if (fix.ok) console.warn(`[DEX] createEvent: EventNumber ${nextEventNumber} war doppelt vergeben — Item ${newItemId} auf ${fresh} korrigiert.`);
          else console.warn(`[DEX] createEvent: EventNumber ${nextEventNumber} doppelt vergeben, Korrektur fehlgeschlagen (HTTP ${fix.status}) — bitte im Admin Center prüfen.`);
        }
      }
    } catch (err) { console.warn('[DEX] createEvent: Eindeutigkeits-Prüfung der EventNumber fehlgeschlagen:', err); }
    reportProgress('done');
    return newItemId;
  } catch (err) {
    if (err instanceof Error) throw err;
    return null;
  }
}

/**
 * Admin-Cleanup beim App-Start: alle Events mit EventStatus='Active' und EndDate < jetzt
 * werden automatisch auf 'Completed' gesetzt. Liefert die Anzahl der aktualisierten Events.
 */
export async function markExpiredEventsAsCompleted(svc: EventService): Promise<number> {
  try {
    // SharePoint OData Filter: Active + EndDate < jetzt
    const nowIso = new Date().toISOString();
    const filter = `EventStatus eq 'Active' and EndDate lt datetime'${nowIso}'`;
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=${encodeURIComponent(filter)}&$select=Id,Title,EndDate&$top=500`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return 0;
    const data = await resp.json();
    const items: Array<{ Id: number; Title: string }> = data.value || data.d?.results || [];
    if (items.length === 0) return 0;

    let updated = 0;
    for (const it of items) {
      try {
        const ok = await svc.updateEvent(it.Id, { 'EventStatus': 'Completed' });
        if (ok) updated += 1;
      } catch { /* einzelnes Update überspringen */ }
    }
    return updated;
  } catch (err) {
    console.warn('[DEX] markExpiredEventsAsCompleted failed:', err);
    return 0;
  }
}

/**
 * Event aktualisieren
 */
/**
 * v11.11: Versionsverlauf des Event-Items aus DEX_Events lesen, um
 * versehentlich gelöschte Custom-Fields (z.B. b2run_*-Felder nach
 * der zu aggressiven v11.9-Migration) wieder zurückzuholen.
 *
 * Liefert eine Liste der Versionen, jeweils mit dem geparsten
 * `CustomFields`-Array (sortiert: neueste zuerst). Werte ohne
 * CustomFields oder mit leerem Array fallen einfach mit raus, sind
 * aber nicht gefiltert — der Caller entscheidet, welche Version
 * relevant ist.
 */
export async function getEventCustomFieldsHistory(svc: EventService, eventId: number): Promise<Array<{
  versionLabel: string;
  modified: string;
  customFields: Array<Record<string, unknown>>;
}>> {
  try {
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/versions?$select=VersionLabel,Modified,CustomFields`;
    const response = await svc._sp.get(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' },
    });
    if (!response.ok) {
      console.warn('[DEX] getEventCustomFieldsHistory failed:', response.status);
      return [];
    }
    const data = await response.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const versions: any[] = data.value || [];
    return versions.map(v => {
      let parsed: Array<Record<string, unknown>> = [];
      try {
        const raw = (v.CustomFields || '').toString();
        if (raw.trim()) {
          const obj = JSON.parse(raw);
          if (Array.isArray(obj)) parsed = obj as Array<Record<string, unknown>>;
        }
      } catch { /* invalid JSON in old version → leer */ }
      return {
        versionLabel: String(v.VersionLabel || ''),
        modified: String(v.Modified || ''),
        customFields: parsed,
      };
    });
  } catch (err) {
    console.warn('[DEX] getEventCustomFieldsHistory error:', err);
    return [];
  }
}

export async function updateEvent(svc: EventService, eventId: number, updates: Record<string, unknown>, retried?: boolean): Promise<boolean> {
  svc.lastUpdateEventError = '';
  try {
    // v28.66: zentraler Schutz für EndDate — analog zu createEvent. Ein
    // leeres EndDate in DEX_Events lässt den DEX_CreateOutlookEvent-Flow in
    // „Create event (V4)" mit convertFromUtc(null) abstürzen. Deshalb hier,
    // am gemeinsamen Nadelöhr aller Update-Pfade, aufräumen:
    //  - leeres EndDate + StartDate im selben Update -> Start als Ende,
    //  - sonst das Feld weglassen, statt einen gespeicherten Wert mit null
    //    zu überschreiben (leer war ohnehin nie ein gültiger Zustand).
    const safeUpdates: Record<string, unknown> = { ...updates };
    if ('EndDate' in safeUpdates && !safeUpdates.EndDate) {
      if (safeUpdates.StartDate) {
        safeUpdates.EndDate = safeUpdates.StartDate;
      } else {
        delete safeUpdates.EndDate;
      }
    }
    const payload = {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
      ...safeUpdates,
    };

    // v28.10: SharePoint lehnt REST-Bodies > 2 MB mit einem kryptischen
    // HTTP 400 ab („The request message is too big"). Vorab prüfen und
    // eine verständliche Meldung liefern — Verursacher ist praktisch
    // immer ein zu großes eingebettetes Bild (Mail-/Outlook-Logo oder
    // ein ins Mail-/Termin-Template eingefügtes Bild).
    const LIMIT = 1_900_000;
    const payloadStr = JSON.stringify(payload);
    if (payloadStr.length > LIMIT) {
      // v28.31: Statt aufzugeben in MEHREREN Requests nacheinander schreiben.
      // Das 2-MB-Limit gilt pro REST-Aufruf, nicht pro Item — ein Event mit
      // eingebetteten Bildern (OutlookBody + EmailTemplateOverrides +
      // EmailImageBase64 tragen dasselbe Bild je einmal) passt problemlos,
      // wenn man die Felder auf mehrere MERGEs verteilt. Vorher brach der
      // Save hier still ab: In der Konsole stand nur eine Warnung, im Wizard
      // passierte auf „Speichern" schlicht nichts.
      const FIELD_OVERHEAD = 160; // __metadata + Klammern/Kommas
      const entries = Object.keys(safeUpdates)
        .map(k => ({ k, size: JSON.stringify({ [k]: safeUpdates[k] }).length }))
        .sort((a, b) => b.size - a.size);
      // Ein EINZELNES Feld über dem Limit lässt sich nicht aufteilen — hier
      // hilft nur ein kleineres Bild. Feldname mitgeben, damit der Organizer
      // weiß, wo er suchen muss.
      const tooBig = entries.filter(e => e.size + FIELD_OVERHEAD > LIMIT);
      if (tooBig.length > 0) {
        svc.lastUpdateEventError = `Ein einzelnes Feld ist zu groß für SharePoint (${tooBig.map(e => `${e.k}: ${Math.round(e.size / 1024)} KB`).join(', ')}). Ursache ist praktisch immer ein zu großes eingebettetes Bild (Mail-Logo, Outlook-Kopfbild oder ein Bild im Mail-/Termin-Text). Bitte das Bild entfernen oder kleiner erneut hochladen.`;
        console.warn('[DEX] updateEvent: einzelnes Feld über dem Limit —', tooBig);
        return false;
      }
      const groups: Array<Record<string, unknown>> = [];
      let cur: Record<string, unknown> = {};
      let curSize = FIELD_OVERHEAD;
      for (const e of entries) {
        if (curSize + e.size > LIMIT && Object.keys(cur).length > 0) {
          groups.push(cur); cur = {}; curSize = FIELD_OVERHEAD;
        }
        cur[e.k] = safeUpdates[e.k];
        curSize += e.size;
      }
      if (Object.keys(cur).length > 0) groups.push(cur);
      console.warn(`[DEX] updateEvent: Payload ${payloadStr.length} Bytes > Limit — wird in ${groups.length} aufeinanderfolgende Schreibvorgänge aufgeteilt.`);
      for (let i = 0; i < groups.length; i++) {
        const ok = await svc.updateEvent(eventId, groups[i], retried);
        if (!ok) {
          // lastUpdateEventError kommt aus dem fehlgeschlagenen Teil-Request.
          svc.lastUpdateEventError = `Teil ${i + 1} von ${groups.length} konnte nicht gespeichert werden. ${svc.lastUpdateEventError}`.trim();
          return false;
        }
      }
      return true;
    }

    const response = await svc._sp.post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE',
          'odata-version': '',
        },
        body: payloadStr,
      }
    );
    if (!response.ok) {
      const errText = await response.text().catch(() => '');
      console.warn('[DEX] updateEvent failed:', response.status, errText.substring(0, 400));
      // SharePoint-Fehlertext extrahieren (verbose: error.message.value).
      let spMsg = '';
      try {
        const parsed = JSON.parse(errText);
        spMsg = parsed?.error?.message?.value || parsed?.['odata.error']?.message?.value || '';
      } catch { /* kein JSON */ }
      const statusHint = response.status === 403
        ? 'Keine Berechtigung — du brauchst Schreibrechte auf der Event-Liste (Organizer/Admin).'
        : response.status === 404
          ? 'Das Event wurde in der Liste nicht gefunden — womöglich wurde es zwischenzeitlich gelöscht.'
          : response.status === 409 || response.status === 412
            ? 'Das Event wurde zeitgleich von jemand anderem geändert — bitte neu laden und erneut speichern.'
            : '';
      svc.lastUpdateEventError = [`HTTP ${response.status}`, statusHint, spMsg && spMsg !== statusHint ? spMsg.slice(0, 300) : '']
        .filter(Boolean).join(' — ');

      // v26.54: „Invalid text value" = ein String-Wert passt nicht in eine
      // EINZEILIGE Text-Spalte (255-Zeichen-Limit). SharePoint nennt das
      // betroffene Feld nicht — wir diagnostizieren selbst: Payload-Werte
      // gegen die Live-Feldtypen der Liste halten. Spalten, die laut
      // Schema-Definition ohnehin mehrzeilig (Note) sein sollten, werden
      // sofort migriert und der Save EINMAL automatisch wiederholt. Alle
      // anderen Treffer werden in der Fehlermeldung beim Namen genannt.
      if (/invalid text value|text field contains invalid data/i.test(spMsg)) {
        const offenders = await findInvalidTextFields(svc, 'DEX_Events', safeUpdates);
        if (offenders.length > 0) {
          console.warn('[DEX] updateEvent: Werte passen nicht in einzeilige Text-Spalten:',
            offenders.map((o) => `${o.internalName} (${o.length} Zeichen${o.intendedNote ? ', sollte Note sein' : ''})`).join(', '));
          const healable = offenders.filter((o) => o.intendedNote);
          if (!retried && healable.length > 0) {
            for (const o of healable) {
              await svc._upgradeTextFieldToNote('DEX_Events', o.title);
            }
            return svc.updateEvent(eventId, safeUpdates, true);
          }
          svc.lastUpdateEventError += ` | ${offenders
            .map((o) => `Betroffenes Feld: „${o.title}" — ${o.length} Zeichen, die Spalte ist einzeiliger Text (max. 255 Zeichen)`)
            .join('; ')}`;
        }
      }
    }
    return response.ok;
  } catch (err) {
    svc.lastUpdateEventError = `Netzwerkfehler — keine Verbindung zu SharePoint${err instanceof Error && err.message ? ` (${err.message.slice(0, 150)})` : ''}.`;
    return false;
  }
}

/**
 * v26.54: Diagnose-Helfer für „Invalid text value. A text field contains
 * invalid data." beim Event-Update. Findet alle String-Werte im Update-
 * Payload, die zu lang für eine einzeilige Text-Spalte sind (> 255 Zeichen
 * oder mit Zeilenumbrüchen), deren Ziel-Spalte auf der LIVE-Liste aber
 * tatsächlich als einzeiliger Text ('Text') liegt. `intendedNote` markiert
 * Spalten, die laut Schema-Definition (getEventsFieldDefinitions) eigentlich
 * mehrzeilig (Typ 3, Note) sein sollten — die dürfen automatisch per
 * _upgradeTextFieldToNote geheilt werden.
 */
async function findInvalidTextFields(
  svc: EventService,
  listName: string,
  updates: Record<string, unknown>
): Promise<Array<{ internalName: string; title: string; length: number; intendedNote: boolean }>> {
  const out: Array<{ internalName: string; title: string; length: number; intendedNote: boolean }> = [];
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName,Title,TypeAsString&$filter=Hidden eq false&$top=300`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return out;
    const data = await resp.json();
    const fields: Array<{ InternalName: string; Title: string; TypeAsString: string }> = data.value || [];
    const defs = svc.getEventsFieldDefinitions();
    for (const key of Object.keys(updates)) {
      const v = updates[key];
      if (typeof v !== 'string') continue;
      if (v.length <= 255 && v.indexOf('\n') < 0) continue;
      const f = fields.filter((x) => x.InternalName === key)[0];
      if (!f || f.TypeAsString !== 'Text') continue;
      const def = defs.filter((d) => d.title === key)[0];
      out.push({ internalName: key, title: f.Title, length: v.length, intendedNote: !!def && def.type === 3 });
    }
  } catch { /* Diagnose darf den Fehlerpfad nie zusätzlich brechen */ }
  return out;
}

/**
 * Event vollständig löschen:
 * 1. Subsite löschen (inkl. Teilnehmerliste) - für neue Events
 * 2. Alte Registrierungsliste löschen (DEX_Reg_*) - für alte Events
 * 3. Event-Eintrag aus DEX_Events löschen
 */
export async function deleteEvent(svc: EventService, eventId: number): Promise<boolean> {
  try {
    // Event-Daten laden um SubsiteUrl und RegistrationListName zu bekommen
    const event = await svc.getEvent(eventId);
    if (!event) return false;

    // v30.67: Register ZUERST — dieselbe Reihenfolge wie deleteParticipantData
    // seit v29.3 (CLAUDE.md „Löschungen zuerst im Register"). Bisher wurde die
    // Subsite recycelt und DANACH DEX_Participants aufgeräumt, und zwar so,
    // dass ein Scheitern nicht auffiel: nicht strikt gelesen (eine abgebrochene
    // Seite kam still als vollständige Liste), alle MERGEs gleichzeitig
    // (Promise.all — die Einladung zur Drosselung) und jeder Fehler per
    // `.catch(() => null)` verworfen. Was im Register stehen blieb, ließ sich
    // nicht mehr nachrechnen, weil die Teilnehmerliste schon weg war — genau
    // die „Verweis ohne Zeile"-Rückstände der Register-Prüfung. Jetzt: strikt
    // lesen, sequentiell schreiben, Fehler zählen, und bei Fehlern mit `false`
    // abbrechen, BEVOR irgendetwas recycelt oder in die Outlook-Queue gestellt
    // wird. Das Event steht dann noch; der Löschversuch lässt sich wiederholen.
    if (event.EventNumber) {
      let registryFailed = 0;
      try {
        const allParticipants = await svc.fetchAllParticipantsOrThrow();
        const en = String(event.EventNumber);
        const affected = allParticipants.filter(p =>
          (p.EventRegistered?.split(',').map(s => s.trim()).includes(en))
          || (p.EventOnWaitlist?.split(',').map(s => s.trim()).includes(en)));
        for (const p of affected) {
          // eslint-disable-next-line no-await-in-loop
          const ok = await svc.removeParticipantEvent(p.Email, event.EventNumber).catch(() => false);
          if (!ok) registryFailed += 1;
        }
      } catch (err) {
        console.warn('[DEX] deleteEvent: Register nicht vollständig lesbar — Löschung abgebrochen:', err);
        return false;
      }
      if (registryFailed > 0) {
        console.warn(`[DEX] deleteEvent: ${registryFailed} Register-Einträge nicht aktualisiert — Löschung abgebrochen (Event ${event.EventNumber}).`);
        return false;
      }
    }

    // 0. Outlook-Kalendereintrag per Queue löschen (VOR allem anderen, damit
    //    CalendarLink noch vorhanden ist). Der DEX_Outlook_Einladungen-Flow
    //    greift den DeleteEvent-Eintrag auf und löscht den Kalender-Termin
    //    im Shared Mailbox über den Flow-Service-Account.
    //    Fehler hier ignorieren - Event-Delete soll trotzdem durchlaufen.
    if (event.CalendarLink) {
      try {
        await svc.queueOutlookDeleteEvent(String(eventId), event.Title || '', event.CalendarLink);
      } catch { /* Queue-Fehler ignorieren */ }
    }
    // 1. Subsite RECYCEN (v9.0: nicht mehr per DELETE, sonst landet die
    //    Subsite permanent weg ohne Recycle-Bin-Eintrag. recycle() legt
    //    die Subsite mitsamt Teilnehmerliste 93 Tage in den Site
    //    Collection Recycle Bin → ein Tenant-Admin / Site Collection
    //    Admin kann sie dort wiederherstellen falls nötig.
    if (event.SubsiteUrl) {
      try {
        await svc._post(`${event.SubsiteUrl}/_api/web/recycle`, {});
      } catch {
        console.warn('[DEX] Subsite konnte nicht in den Recycle Bin verschoben werden:', event.SubsiteUrl);
      }
    }

    // 2. Event-Bild ebenfalls RECYCEN statt löschen.
    if (event.EventImageUrl) {
      try {
        const url = new URL(event.EventImageUrl);
        const serverRelUrl = url.pathname;
        if (serverRelUrl.indexOf('DEX_EventImages') >= 0) {
          await svc._post(
            `${svc.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelUrl}')/recycle`,
            {}
          );
        }
      } catch {
        console.warn('[DEX] Event-Bild konnte nicht in den Recycle Bin verschoben werden');
      }
    }

    // 3. Alte Registrierungsliste recyceln (legacy Events ohne Subsite).
    if (event.RegistrationListName && event.RegistrationListName !== 'Teilnehmer') {
      try {
        await svc._post(
          `${svc.siteUrl}/_api/web/lists/getbytitle('${event.RegistrationListName.replace(/'/g, "''")}')/recycle`,
          {}
        );
      } catch {
        console.warn('[DEX] Alte Registrierungsliste konnte nicht recycelt werden:', event.RegistrationListName);
      }
    }

    // (v30.67: Der DEX_Participants-Block stand hier — er läuft jetzt ganz
    // oben, VOR dem ersten Recycle, siehe Kommentar dort.)

    // 4. Event-Dokumente löschen (SiteAssets/DEX_EventDocs/Event_{number}_*)
    if (event.EventNumber) {
      try {
        const serverRelUrl = svc.context.pageContext.web.serverRelativeUrl;
        const safeName = (event.Title || '').replace(/[#%&*:<>?/\\|"']/g, '').replace(/\s+/g, '_').substring(0, 50);
        const folderName = safeName ? `Event_${event.EventNumber}_${safeName}` : `Event_${event.EventNumber}`;
        await svc._delete(`${svc.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelUrl}/SiteAssets/DEX_EventDocs/${folderName}')`);
      } catch {
        // Fallback: alten Ordnernamen ohne Titel probieren
        try {
          const serverRelUrl = svc.context.pageContext.web.serverRelativeUrl;
          await svc._delete(`${svc.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelUrl}/SiteAssets/DEX_EventDocs/Event_${event.EventNumber}')`);
        } catch { /* Ordner nicht gefunden */ }
      }
    }

    // 5. Event-Eintrag aus DEX_Events RECYCEN (v9.0: per recycle() statt
    //    delete(), damit ein Admin via SharePoint-Recycle-Bin das Item
    //    bei Bedarf 93 Tage lang wiederherstellen kann).
    const response = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/recycle`,
      {}
    );

    // 6. Audit-Eintrag in DEX_ChangeLog (v9.0). Best-effort, blockt
    //    den Lösch-Vorgang nicht falls Logging fehlschlägt.
    try {
      await svc.writeChangeLog({
        action: 'EventDeletedTest', // wird vom Aufrufer überschrieben
        targetType: 'Event',
        targetId: String(eventId),
        targetName: event.Title || '',
        eventId: String(eventId),
        eventTitle: event.Title || '',
        details: {
          subsiteUrl: event.SubsiteUrl || '',
          eventNumber: event.EventNumber,
          recycledTo: 'SharePoint Recycle Bin (93 Tage)',
        },
      });
    } catch { /* */ }

    return response.ok;
  } catch {
    return false;
  }
}

/** v26.13: Versions-Historie der CustomFields-Spalte eines DEX_Events-Items
 *  (neueste zuerst). Grundlage für die Wiederherstellung versehentlich
 *  überschriebener Custom-Field-Beschreibungen (helpText etc.) aus der
 *  SharePoint-Versionshistorie. */
export async function getEventCustomFieldsVersions(svc: EventService, itemId: number): Promise<Array<{ created: string; customFields: string }>> {
  // WICHTIG (v26.15): $select=Created,CustomFields ist PFLICHT — sonst liefert
  // der versions-Endpunkt ALLE Felder pro Version (inkl. der riesigen
  // OutlookBody-/EmailTemplateOverrides-Base64-Logos). Bei stark bearbeiteten
  // Events (z.B. 188 Versionen) sprengt das die Antwortgröße und SharePoint
  // bricht nach ~51 Versionen ab → die Version MIT der Beschreibung fehlte und
  // es kam fälschlich „helpText in Historie: false". KEIN $orderby (das löst
  // auf dem versions-Endpunkt 400 aus) — wir sortieren clientseitig nach
  // Created absteigend (neueste zuerst). Folgeseiten via nextLink einsammeln.
  const out: Array<{ created: string; customFields: string }> = [];
  let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${itemId})/versions?$select=Created,CustomFields&$top=500`;
  let guard = 0;
  try {
    while (url && guard < 25) {
      guard++;
      const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) { console.warn('[DEX restore] versions HTTP', resp.status, 'für Item', itemId); break; }
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const v of (items as any[])) {
        out.push({ created: v.Created || '', customFields: typeof v.CustomFields === 'string' ? v.CustomFields : '' });
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      url = (data as any)['@odata.nextLink'] || (data.d && (data.d as any).__next) || null;
    }
  } catch (e) {
    console.warn('[DEX restore] versions fetch failed für Item', itemId, e);
  }
  out.sort((a, b) => (b.created || '').localeCompare(a.created || ''));
  return out;
}
