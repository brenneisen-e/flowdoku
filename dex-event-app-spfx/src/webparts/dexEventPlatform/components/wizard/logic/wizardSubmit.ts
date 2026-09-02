/* wizardSubmit — aus EventCreationPage.tsx ausgelagert (Zeilen 4335-5973 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
import * as React from 'react';
import { setSaveInProgress } from '../../../utils/saveGuard';
import { buildOutlookLocation } from '../../../utils/eventFormat';
import { outlookLogoPiggyback, resolveAudienceMembersToCsv, serializeCustomFields } from '../../wizard/wizardHelpers';
import { formatOrganizerList } from '../../../context/EventContext';
import { buildOutlookBody, eventCreatedEmail, getCachedOrbBase64, replacePlaceholders } from '../../../services/EmailTemplates';
import { BundledComm, bundledCommConfig } from '../../../utils/bundledComm';
import { EventService } from '../../../services/EventService';
import { compressImage } from '../../../utils/imageCompress';
import { AgendaItem, EventType } from '../../../types';
import { CustomFieldInput } from '../../wizard/customFieldInput';
import { ImgView, SubEventDraft } from '../../wizard/wizardTypes';
import { EmailOverrideEntry } from '../../wizard/emailOverrideEntry';

export interface WizardSubmitCtx {
  activeFrom: string;
  addrCity: string;
  addrHouseNo: string;
  addrStreet: string;
  addrZip: string;
  agenda: AgendaItem[];
  allDay: boolean;
  allowAttendeeUpload: boolean;
  askSalutation: boolean;
  askTeamName: boolean;
  assistantsCanSee: boolean;
  attendeeUploadHint: string;
  attendeeUploadLabel: string;
  audience: string;
  berlinLocalToUtcIso: (localStr: string) => string;
  bilingualFields: boolean;
  billingPiggyback: () => Record<string, unknown>;
  bundledComm: BundledComm;
  childEventsOf: (parentEventId: string) => import("../../../types/index").DeloitteEvent[];
  childGender: "" | "m" | "f" | "n";
  childTermPlural: string;
  childTermSingular: string;
  computeFormSnapshot: () => string;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  confirmDialogEnabled: boolean;
  confirmDialogMode: string;
  confirmDialogText: string;
  contactEmail: string;
  contactInfo: string;
  contactName: string;
  contactOrganizerEmail: string;
  coOrganizerEmails: string[];
  coOrganizerNames: string[];
  createdEventIdRef: React.MutableRefObject<string>;
  createEvent: (event: import("../../../context/eventContextTypes").CreateEventInput) => Promise<number>;
  currentUser: import("../../../types/index").User;
  customFields: CustomFieldInput[];
  deadlineToEndOfDayIso: (dateStr: string) => string | null;
  description: string;
  documents: { name: string; file?: File; url: string; size: number; }[];
  DRAFT_KEY: string;
  durchstarterCapacity: string;
  durchstarterRequiresProof: boolean;
  durchstarterStartblock: string;
  editEvent: import("../../../types/index").DeloitteEvent;
  effTeamsLink: () => string;
  endDate: string;
  eventImageUrl: string;
  eventType: EventType;
  excludedUsers: string[];
  filterMode: "AND" | "OR";
  funstarterCapacity: string;
  funstarterStartblock: string;
  getGroupMembers: (groupEmail: string) => Promise<{ groupName: string; members: { email: string; displayName: string; firstName?: string; lastName?: string; jobTitle?: string; location?: string; }[]; }>;
  getLastEventUpdateError: () => string;
  headerImageLayoutConfig: { _headerImageLayout: { width: number; paddingV: number; paddingH: number; }; } | { _headerImageLayout?: undefined; };
  headerLayoutFor: (logoB64: string) => {    imageWidth: number;    imagePaddingV: number;    imagePaddingH: number;};
  hiddenOrganizerEmails: string[];
  hideOrganizer: boolean;
  hideOrganizerIndividualOnly: boolean;
  imageBanner: boolean;
  imageDisplay: { card?: ImgView; hero?: ImgView; };
  imageFile: File;
  imageOrigAspect: number;
  imageOrigFile: File;
  initialDocumentNames: string[];
  initialFormSnapshotRef: React.MutableRefObject<string>;
  initialOrgGetsSubInvitesRef: React.MutableRefObject<boolean>;
  initialSubEventDbIds: string[];
  isB2runTemplate: boolean;
  isDe: boolean;
  isEditMode: boolean;
  isFictive: boolean;
  klammerDeadline: string;
  lastDeregisterDate: string;
  lastDraftJsonRef: React.MutableRefObject<string>;
  location: string;
  locationFilter: string;
  mainCommDisabledAck: boolean;
  mainEventLabel: string;
  mainEventLabelMode: "none" | "default" | "custom";
  maxParticipants: string;
  noCancelAfterDeadline: boolean;
  noDescription: boolean;
  notifyAdminsExternalAudienceAccess: (eventTitle: string, persons: string[], requesterName: string) => Promise<void>;
  notifyNewCoOrganizers: (eventId: string, eventTitle: string, added: { name: string; email: string; }[], isDe: boolean, disableOutlook?: boolean) => Promise<void>;
  notifyOrgCancelMode: "never" | "always" | "afterDeadline";
  notifyOrgRegisterFromDate: string;
  notifyOrgRegisterMode: "never" | "always" | "fromDate";
  onlineMeetingMode: "none" | "own" | "auto";
  organizer: string;
  organizerDisplayLarge: boolean;
  organizerEmails: string[];
  orgGetsSubInvites: boolean;
  outlookEndOverride: string;
  outlookLocationOverride: string;
  outlookStartOverride: string;
  outlookTeamsLink: () => string;
  pendingOutlookDirtyWriteRef: React.MutableRefObject<boolean>;
  pendingOutlookDirtyWriteRefs: React.MutableRefObject<Record<string, boolean>>;
  pendingOutlookUpdateForSubEventsRef: React.MutableRefObject<string[]>;
  pendingOutlookUpdateForTopRef: React.MutableRefObject<boolean>;
  pendingSuccessDispatchRef: React.MutableRefObject<{ title: string; eventId: string; type: 'create' | 'update'; }>;
  persistSubEventsForParent: (parentEventId: string, onStep?: (done: number, total: number, title: string) => void) => Promise<void>;
  previewBeforeActive: boolean;
  qrScannerEmails: string[];
  qrScannerNames: string[];
  quiz: { id: string; question: string; options: string[]; correctIndices: number[]; imageBase64?: string; section?: string; }[];
  quizClusterSize: number;
  refreshEvents: () => Promise<void>;
  registrationDeadline: string;
  registrationLanguage: "" | "de" | "en";
  regRuleEnabled: boolean;
  requestCoOrganizerApprovals: (orgNames: string, orgEmails: string, eventTitle: string) => Promise<void>;
  requireSubEventSelection: boolean;
  resolveTopLevelCommState: () => { emailLanguage: string; emailLogoBase64: string; outlookLogoBase64: string; outlookBody: string; outlookHeading: string; outlookSubheading: string; outlookSubject: string; disableEmails: boolean; disableRegistrationEmail: boolean; disableCancellationEmail: boolean; autoDeregisterOnDecline: boolean; inactiveHandling?: 'notify' | 'autoderegister'; disableOutlook: boolean; emailTemplateOverrides: Record<string, EmailOverrideEntry>; };
  sanitizeOrganizerPairs: () => {    orgString: string;    orgEmailString: string;    droppedCount: number;};
  selectedEventId: string;
  setDraftSavedAt: React.Dispatch<React.SetStateAction<number>>;
  setError: React.Dispatch<React.SetStateAction<string>>;
  setImageUploadError: React.Dispatch<React.SetStateAction<string>>;
  setIsSubmitting: React.Dispatch<React.SetStateAction<boolean>>;
  setNavigationGuard: (guard: () => Promise<boolean>) => void;
  setPendingDraft: React.Dispatch<React.SetStateAction<{ savedAt: number; data: Record<string, unknown>; }>>;
  setPendingSuccessDispatch: React.Dispatch<React.SetStateAction<{ title: string; eventId: string; type: "create" | "update"; }>>;
  setProgress: React.Dispatch<React.SetStateAction<number>>;
  setProgressLabel: React.Dispatch<React.SetStateAction<string>>;
  setRemovedSavedSubs: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setShowSummaryModal: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  showAsFree: boolean;
  shrinkLogoB64: (b64: string) => Promise<string>;
  splitDescA: string;
  splitDescB: string;
  splitDisplayOrderReversed: boolean;
  splitHelpText: string;
  splitLabelA: string;
  splitLabelB: string;
  splitSectionTitle: string;
  splitSharedWaitlist: boolean;
  startDate: string;
  subDeadlineRulePiggyback: () => Record<string, unknown>;
  subEventCalendar: boolean;
  subEventOpenRulePiggyback: () => Record<string, unknown>;
  subEventSingleChoice: boolean;
  subEventsOnlyMode: boolean;
  subEventsOptIn: boolean;
  subEventsRef: React.MutableRefObject<SubEventDraft[]>;
  teamJoinRequiresApproval: boolean;
  teamMembersCannotCreate: boolean;
  teamOpenSlotsVisible: boolean;
  teamPartialAllowed: boolean;
  teamRegistrationEnabled: boolean;
  teamSize: number;
  teamTermPlural: string;
  teamTermSingular: string;
  testTeamEmails: string[];
  testTeamNames: string[];
  title: string;
  transferTimes: { id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string; }[];
  unlimitedParticipants: boolean;
  updateEvent: (eventId: string, updates: Record<string, unknown>, opts?: { skipReload?: boolean; }) => Promise<boolean>;
  userCancelAllowed: boolean;
  useSplitCapacities: boolean;
  visAllSubsPiggyback: () => Record<string, unknown>;
  waitlistEnabled: boolean;
  wizardImgAspect: number;
}

export async function runWizardSubmit(ctx: WizardSubmitCtx): Promise<void> {
  const { activeFrom, addrCity, addrHouseNo, addrStreet, addrZip, agenda, allDay, allowAttendeeUpload, askSalutation, askTeamName, assistantsCanSee, attendeeUploadHint, attendeeUploadLabel, audience, berlinLocalToUtcIso, bilingualFields, billingPiggyback, bundledComm, childEventsOf, childGender, childTermPlural, childTermSingular, computeFormSnapshot, confirmDialog, confirmDialogEnabled, confirmDialogMode, confirmDialogText, contactEmail, contactInfo, contactName, contactOrganizerEmail, coOrganizerEmails, coOrganizerNames, createdEventIdRef, createEvent, currentUser, customFields, deadlineToEndOfDayIso, description, documents, DRAFT_KEY, durchstarterCapacity, durchstarterRequiresProof, durchstarterStartblock, editEvent, effTeamsLink, endDate, eventImageUrl, eventType, excludedUsers, filterMode, funstarterCapacity, funstarterStartblock, getGroupMembers, getLastEventUpdateError, headerImageLayoutConfig, headerLayoutFor, hiddenOrganizerEmails, hideOrganizer, hideOrganizerIndividualOnly, imageBanner, imageDisplay, imageFile, imageOrigAspect, imageOrigFile, initialDocumentNames, initialFormSnapshotRef, initialOrgGetsSubInvitesRef, initialSubEventDbIds, isB2runTemplate, isDe, isEditMode, isFictive, klammerDeadline, lastDeregisterDate, lastDraftJsonRef, location, locationFilter, mainCommDisabledAck, mainEventLabel, mainEventLabelMode, maxParticipants, noCancelAfterDeadline, noDescription, notifyAdminsExternalAudienceAccess, notifyNewCoOrganizers, notifyOrgCancelMode, notifyOrgRegisterFromDate, notifyOrgRegisterMode, onlineMeetingMode, organizer, organizerDisplayLarge, organizerEmails, orgGetsSubInvites, outlookEndOverride, outlookLocationOverride, outlookStartOverride, outlookTeamsLink, pendingOutlookDirtyWriteRef, pendingOutlookDirtyWriteRefs, pendingOutlookUpdateForSubEventsRef, pendingOutlookUpdateForTopRef, pendingSuccessDispatchRef, persistSubEventsForParent, previewBeforeActive, qrScannerEmails, qrScannerNames, quiz, quizClusterSize, refreshEvents, registrationDeadline, registrationLanguage, regRuleEnabled, requestCoOrganizerApprovals, requireSubEventSelection, resolveTopLevelCommState, sanitizeOrganizerPairs, selectedEventId, setDraftSavedAt, setError, setImageUploadError, setIsSubmitting, setNavigationGuard, setPendingDraft, setPendingSuccessDispatch, setProgress, setProgressLabel, setRemovedSavedSubs, setShowSummaryModal, showAlert, showAsFree, shrinkLogoB64, splitDescA, splitDescB, splitDisplayOrderReversed, splitHelpText, splitLabelA, splitLabelB, splitSectionTitle, splitSharedWaitlist, startDate, subDeadlineRulePiggyback, subEventCalendar, subEventOpenRulePiggyback, subEventSingleChoice, subEventsOnlyMode, subEventsOptIn, subEventsRef, teamJoinRequiresApproval, teamMembersCannotCreate, teamOpenSlotsVisible, teamPartialAllowed, teamRegistrationEnabled, teamSize, teamTermPlural, teamTermSingular, testTeamEmails, testTeamNames, title, transferTimes, unlimitedParticipants, updateEvent, userCancelAllowed, useSplitCapacities, visAllSubsPiggyback, waitlistEnabled, wizardImgAspect } = ctx;
    // v9.14: Beschreibung ist jetzt optional. Nur Title bleibt Pflicht.
    if (!title) return;

    // v18.36: Harte Datums-Validierung als letzter Riegel — das Enddatum darf
    // NIE vor (oder gleich) dem Startdatum liegen. Outlook lehnt solche Termine
    // ab und der DEX_CreateOutlookEvent-Flow failt dann mit HTTP 400
    // („At least one property failed validation"). Gilt für das Hauptevent UND
    // jedes Sub-Event — Sub-Events liefen bisher ohne Datums-Prüfung durch.
    const dateProblems: string[] = [];
    if (startDate && endDate && new Date(endDate) <= new Date(startDate)) {
      dateProblems.push(isDe ? 'Hauptevent' : 'Main event');
    }
    subEventsRef.current.forEach(s => {
      if (s.title && s.title.trim() && s.startDate && s.endDate && new Date(s.endDate) <= new Date(s.startDate)) {
        dateProblems.push(isDe ? `Sub-Event „${s.title}"` : `Sub-event „${s.title}"`);
      }
    });
    if (dateProblems.length > 0) {
      // eslint-disable-next-line no-alert
      showAlert(isDe
        ? `Das Enddatum darf nicht vor dem Startdatum liegen. Bitte korrigiere das Datum bei: ${dateProblems.join(', ')}.`
        : `The end date must not be before the start date. Please fix the date for: ${dateProblems.join(', ')}.`);
      return;
    }

    // v11.93: Top-Level-Kommunikations-Werte sauber resolven (s. Helper-
    // Doku oben). Sonst würden, falls beim Speichern ein Sub-Event-Tab
    // aktiv ist, die Sub-Event-States (Logo, Outlook-Body, Headings,
    // etc.) fälschlich auf das Haupt-Event geschrieben.
    const topComm = resolveTopLevelCommState();
    const effEmailLanguage = topComm.emailLanguage;
    // v28.10: Logos beim Speichern hart verkleinern (>400 KB → 600px).
    // Rettet auch Events, bei denen früher ein unkomprimiertes Bild als
    // Logo übernommen wurde — die ließen sich sonst gar nicht mehr
    // speichern (SharePoint-2-MB-Limit, das Logo steckt bis zu 3× im Payload).
    const effEmailLogo = await shrinkLogoB64(topComm.emailLogoBase64);
    const effOutlookLogo = await shrinkLogoB64(topComm.outlookLogoBase64);
    const effOutlookBody = topComm.outlookBody;
    const effOutlookHeading = topComm.outlookHeading;
    const effOutlookSubheading = topComm.outlookSubheading;
    const effOutlookSubject = topComm.outlookSubject;
    const effDisableEmails = topComm.disableEmails;
    // v19.22: granulare An-/Abmelde-Mail-Schalter des Hauptevents top-level
    // auflösen (auf Sub-Tabs hält der State den Sub-Wert → resolveTopLevelCommState).
    const effDisableRegistrationEmail = topComm.disableRegistrationEmail;
    const effDisableCancellationEmail = topComm.disableCancellationEmail;
    // v19.24: Auto-Abmeldung jetzt per Sub-Event gespiegelt → top-level auflösen.
    const effAutoDeregisterOnDecline = topComm.autoDeregisterOnDecline;
    const effInactiveHandling = topComm.inactiveHandling || 'notify';
    const effDisableOutlook = topComm.disableOutlook;

    // v14.4 / v14.5: Wenn das Hauptevent Sub-Events hat UND die
    // Kommunikation auf Top-Level abgestellt ist, muss entweder der
    // „Sub-Event verpflichtend"-Toggle aktiv sein (erzwingt es im
    // Anmeldeformular) ODER der Organizer den Ack-Haken gesetzt haben.
    // Sonst landen Teilnehmer ohne Bestätigungs-Mail und ohne Kalender-
    // Termin in der Liste.
    const hasSubs = subEventsRef.current.some(s => s.title && s.title.trim());
    if (hasSubs && (effDisableEmails || effDisableOutlook) && !requireSubEventSelection && !mainCommDisabledAck) {
      // eslint-disable-next-line no-alert
      showAlert(isDe
        ? 'Du hast die Kommunikation für das Hauptevent deaktiviert. Bitte aktiviere in Schritt 6 (Kommunikation, Tab „Haupt-Event") entweder den Toggle „Anmeldung für mindestens ein Sub-Event verpflichtend" ODER bestätige den Ack-Haken — sonst landen Teilnehmer stumm in der Liste.'
        : 'You disabled communication for the main event. Please either enable the toggle „Require selecting at least one sub-event" in step 6 OR tick the acknowledgement — otherwise attendees land silently in the list.');
      return;
    }

    // v27.11: Destruktive Sub-Event-Löschungen VOR dem Speichern explizit
    // bestätigen lassen. persistSubEventsForParent löscht am Ende alle beim
    // Edit-Start vorhandenen Sub-Events, die nicht mehr im Formular sind —
    // kaskadierend inklusive Subsite/Teilnehmerliste. Das passierte bisher
    // still (z.B. nach Abschalten des Sub-Event-Toggles oder Entfernen einer
    // Karte) — Anmeldungen waren ohne Rückfrage weg.
    if (editEvent && initialSubEventDbIds.length > 0) {
      const keptIds = new Set(subEventsRef.current.map(s => s.dbId).filter(Boolean) as string[]);
      const toDelete = initialSubEventDbIds.filter(id => !keptIds.has(id));
      if (toDelete.length > 0) {
        const titles = toDelete
          .map(id => { const k = childEventsOf(editEvent.id).find(c => c.id === id); return k ? `„${k.title}"` : ''; })
          .filter(Boolean).join(', ');
        const ok = await confirmDialog(
          isDe
            ? `Achtung: Beim Speichern werden ${toDelete.length} Sub-Event(s) endgültig gelöscht${titles ? `: ${titles}` : ''} — inklusive Teilnehmerliste und aller Anmeldungen (93 Tage im Papierkorb). Fortfahren?`
            : `Warning: saving will permanently delete ${toDelete.length} sub-event(s)${titles ? `: ${titles}` : ''} — including their attendee lists and all registrations (recycled for 93 days). Continue?`,
          { danger: true, title: isDe ? 'Sub-Events löschen' : 'Delete sub-events', confirmLabel: isDe ? 'Speichern & löschen' : 'Save & delete' },
        );
        if (!ok) return;
      }
    }

    // v26.51: Fristen-Validierung — An-/Abmeldefrist darf NICHT nach dem
    // Event-Beginn liegen (ergibt keinen Sinn und führte zu inkonsistenten
    // Zuständen). Gilt für Anlegen UND Bearbeiten.
    if (startDate) {
      const startTs = new Date(startDate).getTime();
      const fmtDt = (v: string): string => {
        try { return new Date(v).toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); }
        catch { return v; }
      };
      if (isFinite(startTs)) {
        if (registrationDeadline && new Date(registrationDeadline).getTime() > startTs) {
          showAlert(isDe
            ? `Die Anmeldefrist (${fmtDt(registrationDeadline)}) liegt NACH dem Event-Beginn (${fmtDt(startDate)}). Bitte setze die Frist auf einen Zeitpunkt vor dem Event.`
            : `The registration deadline (${fmtDt(registrationDeadline)}) is AFTER the event start (${fmtDt(startDate)}). Please set the deadline to a time before the event.`, { variant: 'error' });
          return;
        }
        if (userCancelAllowed && lastDeregisterDate && new Date(lastDeregisterDate).getTime() > startTs) {
          showAlert(isDe
            ? `Die Abmeldefrist (${fmtDt(lastDeregisterDate)}) liegt NACH dem Event-Beginn (${fmtDt(startDate)}). Bitte setze die Frist auf einen Zeitpunkt vor dem Event.`
            : `The deregistration deadline (${fmtDt(lastDeregisterDate)}) is AFTER the event start (${fmtDt(startDate)}). Please set the deadline to a time before the event.`, { variant: 'error' });
          return;
        }
      }
    }

    setIsSubmitting(true);
    setError('');
    // v28.98: Ab hier laeuft der Speichervorgang — „Zurück" im Header ist
    // gesperrt, bis er (auch im Fehlerfall) beendet ist.
    setSaveInProgress(true);
    setProgress(0);

    // Schritt 1: Bild wird später (nach Event-Erstellung) als Item-Attachment hochgeladen.
    // Bestehende URL beibehalten (z.B. bei Edit ohne neues Bild).
    setProgress(5);
    setProgressLabel('Event wird vorbereitet...');
    // v29.18: Den Anzeige-Cache-Buster (`?v=<timestamp>`, seit v26.17 in
    // EventContext.buildDisplayImageUrl angehängt) VOR dem Zurückschreiben
    // strippen. `eventImageUrl` kommt aus `editEvent.imageUrl` und trägt ihn
    // mit; ohne Strip schrieb jeder Edit-Save die Anzeige-URL in die Spalte,
    // und der nächste Load hängte ein weiteres `&v=` an — die gespeicherte
    // URL wuchs mit jedem Speichern. Attachment-URLs haben nie einen eigenen
    // Query-Teil, das Kappen an `?` ist deshalb verlustfrei.
    const imageUrl = (eventImageUrl || '').split('?')[0];
    setProgress(15);

    if (isEditMode && selectedEventId) {
      setProgressLabel('Event wird aktualisiert...');
      // v29.77: v11.18-Debug-Trace („customFields state at save") entfernt —
      // der helpText-Roundtrip ist lange verifiziert, das Log war nur Laerm.
      // Sanitize: paart Organizer-Names + -Emails 1:1, droppt unvollständige
      // Pairs — verhindert Mismatch-State in DEX_Events.
      const sanitizedOrgPairEdit = sanitizeOrganizerPairs();
      // Event aktualisieren - nur bekannte Felder senden
      const updates: Record<string, unknown> = {
        'Title': title,
        'Description': description,
        'Location': location,
        'LocationAddress': (addrStreet || addrHouseNo || addrZip || addrCity)
          ? JSON.stringify({ street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity })
          : '',
        // v18.34/v18.40: Outlook-Ort = manuelle Überschreibung, sonst Auto aus
        // Veranstaltungsort + Adresse. Flow mappt OutlookLocation 1:1.
        // v26.54: hart auf 255 kappen — die Spalte ist einzeiliger Text, ein
        // überlanger Ort+Adresse-String ließ sonst den KOMPLETTEN Save mit
        // „Invalid text value" (HTTP 500) abbrechen.
        'OutlookLocation': (outlookLocationOverride.trim() || buildOutlookLocation(location, { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity })).slice(0, 255),
        'LocationFilter': locationFilter,
        'Audience': audience,
        // v16.4: Audience-DLs vor-aufgelöst mitschreiben.
        'AudienceResolvedEmails': await resolveAudienceMembersToCsv(audience, getGroupMembers),
        'FilterMode': filterMode,
        'StartDate': startDate ? berlinLocalToUtcIso(startDate) : null,
        // v22.17: EndDate nie leer lassen (Outlook-Flow-Crash, s.o.) — Fallback Start.
        'EndDate': endDate ? berlinLocalToUtcIso(endDate) : (startDate ? berlinLocalToUtcIso(startDate) : null),
        'RegistrationDeadline': deadlineToEndOfDayIso(registrationDeadline),
        // v29.21 (Audit): bei geteilter Kapazität IMMER 0 (dokumentierte
        // Invariante) — je nach Klickreihenfolge stand hier sonst mal 0, mal
        // die Gruppen-Summe, und alles, was MaxParticipants liest, sah zwei
        // verschiedene Stände für dieselbe Konfiguration.
        'MaxParticipants': useSplitCapacities ? 0 : (unlimitedParticipants ? 0 : (Number(maxParticipants) || 0)),
        // v20.0 BUG-FIX (Audit): WaitlistEnabled wurde beim Edit nie persistiert —
        // das Umschalten der Warteliste auf einem bestehenden Event ging still
        // verloren (nur der Create-Pfad schrieb das Feld).
        'WaitlistEnabled': !!waitlistEnabled,
        'EventImageUrl': imageUrl,
        'Organizer': sanitizedOrgPairEdit.orgString,
        'OrganizerEmail': sanitizedOrgPairEdit.orgEmailString,
        // v10.16: optionaler Ansprechpartner (Anzeige-Feld, kein Login)
        'ContactName': contactName.trim(),
        'ContactEmail': contactEmail.trim(),
        // v28.5: Rückfragen-Kontakt-Organizer persistieren.
        'ContactOrganizerEmail': (contactOrganizerEmail || '').trim(),
        'ContactInfo': contactInfo.trim(),
        // v17.22: zentraler serializeCustomFields-Helper (Options-Pairing +
        // EN-Varianten konsistent zu allen Pfaden).
        'CustomFields': JSON.stringify(serializeCustomFields(customFields, bilingualFields)),
      };

      // Optionale Felder - immer senden damit Löschungen wirken
      // v29.25: Ohne Selbst-Abmeldung gibt es keine Abmeldefrist — ein
      // gespeicherter Alt-Wert würde auf der Anmeldeseite weiter angezeigt.
      updates['LastDeregisterDate'] = userCancelAllowed ? deadlineToEndOfDayIso(lastDeregisterDate) : null;
      // Outlook-Body: Variablen werden bereits hier aufgelöst (gleicher Body für alle Teilnehmer).
      const outlookVars: Record<string, string> = {
        EventTitle: title,
        // v27.5: normalisierte Organizer-Namen für {{Organizer}} (siehe unten).
        Organizer: formatOrganizerList([organizer], effEmailLanguage) || organizer,
        ContactEmail: contactEmail.trim(),
        Location: location,
        Address: [addrStreet, addrHouseNo].filter(Boolean).join(' ') + ((addrZip || addrCity) ? ', ' + [addrZip, addrCity].filter(Boolean).join(' ') : ''),
        StartDate: startDate ? new Date(startDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
        EndDate: endDate ? new Date(endDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
      };
      // v7.4: Auch wenn der User keinen Outlook-Body eingegeben hat, IMMER
      // das Outlook-Mail-Layout (buildOutlookBody) verwenden + einen
      // Default-Body einsetzen, der den Empfänger an die Organizer
      // verweist. Sonst kommt der Termin ganz ohne Body — wirkt
      // unprofessionell und der Teilnehmer hat keinen Ansprechpartner
      // bei organisatorischen Rückfragen.
      // v24.60: Namen wie die Anmelde-Mail normalisieren („Schwartz, Eva" →
      // „Eva Schwartz") und mit „und"/„and" verbinden — nicht stumpf mit Komma.
      const orgNames = formatOrganizerList([organizer], effEmailLanguage);
      const escHtml = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      // v9.8: Default-Body enthält jetzt auch den Abmelde-Hinweis analog zur
      // Anmeldebestätigungs-Mail. Sonst weiß der Empfänger nicht, wie er
      // sich abmelden kann — die Outlook-Decline-Funktion triggert zwar einen
      // Reminder-Flow, aber der eigentliche App-Abmelde-Pfad ist sauberer.
      const APP_URL_OL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';
      const defaultOutlookBody = effEmailLanguage === 'EN'
        ? `<p>You are registered for the event <strong>${escHtml(title)}</strong>.</p>`
          + `<p>If you are unable to attend, please cancel your registration in time via the <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;My Events&ldquo;).</p>`
          + `<p>For organizational questions please contact <strong>${escHtml(orgNames || 'the organizer')}</strong>.</p>`
        : `<p>Ihr seid für das Event <strong>${escHtml(title)}</strong> angemeldet.</p>`
          + `<p>Falls ihr nicht teilnehmen könnt, meldet euch bitte rechtzeitig über die <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;Meine Events&ldquo;) ab.</p>`
          + `<p>Bei organisatorischen Fragen wendet euch bitte an <strong>${escHtml(orgNames || 'den Organizer')}</strong>.</p>`;
      const resolvedBody = effOutlookBody
        ? replacePlaceholders(effOutlookBody, outlookVars)
        : defaultOutlookBody;
      const resolvedOlHeading = effOutlookHeading ? replacePlaceholders(effOutlookHeading, outlookVars) : title;
      // v27.5: Default-Unter-Überschrift = Ort (nicht Datum).
      const resolvedOlSub = effOutlookSubheading ? replacePlaceholders(effOutlookSubheading, outlookVars) : (location || undefined);
      // v18.73: Header-Bild Größe + Innenabstand (event-weit) in den Outlook-Body.
      // v30.67: `effEmailLanguage` statt des rohen States — auf einem Sub-Reiter
      // trägt `emailLanguage` den Wert des Sub-Events, und der Teams-Beitreten-
      // Block im Hauptevent-Termin kam dann in der falschen Sprache.
      const wrappedOutlook = buildOutlookBody(resolvedOlHeading, resolvedBody, resolvedOlSub, headerLayoutFor(effOutlookLogo), outlookTeamsLink(), (effEmailLanguage || '').toUpperCase() !== 'EN');
      // v11.93: Top-Level-Logo aus dem Resolver — sonst würde beim Speichern
      // aus einem Sub-Tab das falsche Logo aufs Haupt-Event geschrieben.
      updates['OutlookBody'] = wrappedOutlook.replace(/\{\{ORB_URL\}\}/g, effOutlookLogo || getCachedOrbBase64() || '');
      // v18.42: Betreff des Outlook-Termins mit-persistieren (leer = Titel via Flow-Fallback).
      updates['OutlookSubject'] = effOutlookSubject.trim();
      // v18.44: abweichendes Outlook-Datum mit-persistieren (leer = Event-Datum via Flow-Fallback).
      updates['OutlookStart'] = outlookStartOverride || null;
      updates['OutlookEnd'] = outlookEndOverride || null;
      updates['AllDay'] = !!allDay; // v29.52
      updates['ShowAsFree'] = !!showAsFree; // v29.54
      // v30.26: Teams-Besprechung nur im Modus „DEX erzeugt den Link".
      updates['OutlookIsOnlineMeeting'] = onlineMeetingMode === 'auto';
      updates['SkipOrganizerInvite'] = !orgGetsSubInvites; // v29.55
      updates['Agenda'] = JSON.stringify(agenda);
      updates['Transfers'] = JSON.stringify(transferTimes);
      updates['FunZone'] = JSON.stringify(quiz);
      updates['QuizClusterSize'] = Math.min(Math.max(1, quizClusterSize || 1), 4);
      updates['EmailLanguage'] = effEmailLanguage;
      // v18.35: erzwungene Anmeldeseiten-Sprache mit-persistieren ('' = App-Sprache).
      updates['RegistrationLanguage'] = registrationLanguage || '';
      // v6.15: B2Run-Config (Starter-Typ → Startblock, Leistungsnachweis-Pflicht)
      // wird in EmailTemplateOverrides._b2run piggyback gespeichert, damit keine
      // neue SP-Spalte nötig ist.
      const b2runExtraConfig = (durchstarterStartblock || funstarterStartblock || durchstarterRequiresProof)
        ? { _b2run: {
            ...(durchstarterStartblock ? { durchstarterStartblock } : {}),
            ...(funstarterStartblock ? { funstarterStartblock } : {}),
            ...(durchstarterRequiresProof ? { durchstarterRequiresProof: true } : {}),
          } }
        : {};
      // v6.19: QR-Code-Scanner piggyback in EmailTemplateOverrides._qrScanners
      const qrScannerConfig = qrScannerEmails.length > 0
        ? { _qrScanners: qrScannerNames.map((n, i) => ({ name: n, email: qrScannerEmails[i] || '' })).filter(x => x.email) }
        : {};
      // v9.18: Co-Organizer-Liste piggyback in EmailTemplateOverrides._coOrganizers
      const coOrganizerConfig = coOrganizerEmails.length > 0
        ? { _coOrganizers: coOrganizerNames.map((n, i) => ({ name: n, email: coOrganizerEmails[i] || '' })).filter(x => x.email) }
        : {};
      // v9.21: Test-Team-Liste piggyback in EmailTemplateOverrides._testTeam
      const testTeamConfig = testTeamEmails.length > 0
        ? { _testTeam: testTeamNames.map((n, i) => ({ name: n, email: testTeamEmails[i] || '' })).filter(x => x.email) }
        : {};
      // v11.25: Display-Reihenfolge-Toggle als piggyback in
      // EmailTemplateOverrides._splitDisplayOrderReversed.
      const splitDispRevConfig = splitDisplayOrderReversed && useSplitCapacities
        ? { _splitDisplayOrderReversed: true }
        : {};
      // v11.93: Top-Level-Logos aus dem Resolver lesen, NICHT direkt aus
      // den State-Variablen — sonst wird beim Speichern aus einem Sub-Tab
      // das Sub-Logo aufs Haupt-Event geschrieben.
      // v14.4: Mail-Text-Overrides ebenfalls aus dem Resolver — vorher wurden
      // beim Speichern auf einem Sub-Tab die Sub-Overrides fälschlich aufs
      // Hauptevent gemerged.
      const topOverrides = topComm.emailTemplateOverrides || {};
      // v14.8: subEventsOnlyMode impliziert requireSubEventSelection — wenn die
      // Hauptevent-Anmeldung gar nicht mehr angeboten wird, MUSS jeder Teilnehmer
      // mindestens einen Sub-Event auswählen, sonst kommt keine Anmeldung zustande.
      const effRequireSubEventSelection = requireSubEventSelection || subEventsOnlyMode;
      const requireSubEventConfig = effRequireSubEventSelection
        ? { _requireSubEventSelection: true }
        : {};
      // v14.8: Sub-Events-Only-Modus + Custom-Bezeichnung als Piggyback.
      const subEventsOnlyConfig = subEventsOnlyMode
        ? { _subEventsOnlyMode: true }
        : {};
      // v28.2: Sub-Events SOFT-deaktiviert — Toggle aus, aber es existieren
      // Sub-Event-Drafts: Kinder bleiben gespeichert, werden aber auf der
      // Anmeldeseite nicht angeboten (kein Löschen mehr über den Toggle).
      const subEventsDisabledConfig = (!subEventsOptIn && subEventsRef.current.some(s => s.title && s.title.trim()))
        ? { _subEventsDisabled: true }
        : {};
      // v28.5: Bild-Banner-Layout (Piggyback).
      const imageBannerConfig = imageBanner ? { _imageBanner: true } : {};
      // v28.79: „Keine Beschreibung nutzen" persistieren. Bisher war das ein
      // reiner UI-Schalter — gespeichert wurde nur ein leeres Feld. Damit
      // konnte niemand mehr unterscheiden, ob der Organizer bewusst keine
      // Beschreibung wollte oder sie schlicht vergessen hat; die Box
      // „Nächste Schritte" meldete sie deshalb ewig als fehlend.
      const noDescriptionConfig = (noDescription && !description.trim()) ? { _noDescription: true } : {};
      // v28.91: Kalender-Modus nur setzen, wenn er aktiv ist — ein
      // abgewaehlter Schalter darf keinen Rest im Blob hinterlassen.
      const subEventCalendarConfig = (subEventCalendar && subEventsOptIn) ? { _subEventCalendar: true } : {};
      const subEventSingleChoiceConfig = (subEventSingleChoice && subEventsOptIn) ? { _subEventSingleChoice: true } : {};
      // v30.61: Nur gesetzte Schalter schreiben — ein Event ohne Bündelung
      // trägt keinen Ballast in den Overrides.
      const bundledCommConf = subEventsOptIn ? bundledCommConfig(bundledComm) : {};
      // v29.25: Abmelde-Sperren — nur setzen, wenn aktiv (ein abgewählter
      // Schalter darf keinen Rest im Blob hinterlassen). Die Nach-Frist-
      // Sperre nur, solange die Selbst-Abmeldung überhaupt erlaubt ist.
      const noSelfCancelConfig = !userCancelAllowed ? { _noSelfCancel: true } : {};
      const noCancelAfterDeadlineConfig = (userCancelAllowed && noCancelAfterDeadline) ? { _noCancelAfterDeadline: true } : {};
      // v29.38: Teams-Link nur speichern, wenn er wie ein Link aussieht — ein
      // halb eingefuegter Text würde sonst als toter Knopf im Termin landen.
      const teamsLinkConfig = /^https?:\/\//i.test(effTeamsLink()) ? { _teamsLink: effTeamsLink() } : {};
      // v28.11: Bestehende Original-Bild-URL beim Edit-Save WEITERTRAGEN —
      // sonst würde der frisch zusammengebaute Overrides-Blob sie wegwerfen.
      // v28.12: auch bei neuem Bild erstmal mitschreiben; der Post-Save-Code
      // patcht sie danach via patchEventOverridesKey auf den frischen Wert
      // bzw. entfernt sie, wenn das neue Bild kein Original braucht.
      const imageOrigUrlConfig = (editEvent && editEvent.imageOrigUrl)
        ? { _imageOrigUrl: editEvent.imageOrigUrl }
        : {};
      // v28.20: explizite Klammer-Frist — nur im Klammer-Modus; leeres Feld
      // entfernt den Key (Blob wird frisch zusammengebaut).
      // v30.6: Bei aktiver rollierender „Anmeldung bis"-Regel NIE mitschreiben
      // — die Fristen gelten dann je Termin, und eine stehengebliebene
      // Klammer-Frist sperrte sonst das gesamte Event (Soft-Opening-Fall).
      const klammerDeadlineConfig: Record<string, unknown> = (() => {
        if (subEventsOptIn && regRuleEnabled) return {};
        const iso = (subEventsOnlyMode && klammerDeadline) ? deadlineToEndOfDayIso(klammerDeadline) : null;
        return iso ? { _klammerDeadline: iso } : {};
      })();
      const childTermConfig = (childTermSingular.trim() || childTermPlural.trim())
        ? { _childEventTerm: { singular: childTermSingular.trim(), plural: childTermPlural.trim(), ...(childGender ? { gender: childGender } : {}) } }
        : {};
      // v18.9: Organizer-Anzeige ausblenden (Piggyback).
      const hideOrganizerConfig = hideOrganizer ? { _hideOrganizer: true } : {};
      const hiddenOrganizersConfig: Record<string, unknown> = hiddenOrganizerEmails.length > 0 ? { _hiddenOrganizers: hiddenOrganizerEmails } : {};
      const hideOrgIndividualConfig: Record<string, unknown> = hideOrganizerIndividualOnly ? { _hideOrgIndividual: true } : {};
      // v22.78: Team-Begriff + „keine neuen Teams"-Flag (Piggyback).
      const teamTermConfig = (teamTermSingular.trim() || teamTermPlural.trim())
        ? { _teamTerm: { singular: teamTermSingular.trim(), plural: teamTermPlural.trim() } }
        : {};
      const teamNoCreateConfig = teamMembersCannotCreate ? { _teamMembersCannotCreate: true } : {};
      // v24.58: Anzeige-Bezeichnung des Haupt-Events (Piggyback).
      const mainEventLabelConfig: Record<string, unknown> =
        mainEventLabelMode === 'none' ? { _mainEventLabel: { mode: 'none' } }
        : (mainEventLabelMode === 'custom' && mainEventLabel.trim()) ? { _mainEventLabel: { mode: 'custom', text: mainEventLabel.trim() } }
        : {};
      // v23.6: Assistenz-Sichtbarkeit (Piggyback).
      const assistantsCanSeeConfig = assistantsCanSee ? { _assistantsCanSee: true } : {};
      // v23.25: Organizer groß darstellen (Piggyback).
      const organizerDisplayLargeConfig = organizerDisplayLarge ? { _organizerDisplayLarge: true } : {};
      // v23.14: Vorschau vor Aktivierung (nur sinnvoll mit activeFrom).
      const previewBeforeActiveConfig = (previewBeforeActive && activeFrom) ? { _previewBeforeActive: true } : {};
      // v23.19: Pro-Ansicht-Bilddarstellung — nur Views speichern, die vom
      // Standard (zoom 1, posY 50) abweichen.
      const imageDisplayConfig = (() => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const out: any = {};
        (['card', 'hero'] as const).forEach(k => {
          const v = imageDisplay[k];
          if (!v) return;
          const hSet = typeof v.height === 'number' && v.height !== 340;
          if (v.zoom !== 1 || v.posY !== 50 || hSet) out[k] = { zoom: v.zoom, posY: v.posY, ...(hSet ? { height: v.height } : {}) };
        });
        return Object.keys(out).length ? { _imageDisplay: out } : {};
      })();
      // v28.2: Die frühere ||-Monsterkette überschritt mit dem neuen
      // _subEventsDisabled-Config TypeScripts Union-Komplexitätslimit
      // (TS2590) — jetzt als typisiertes Array + .some().
      // v29.19: Die Hotel-Planung WEITERTRAGEN — sie wird beim Laden bewusst
      // aus dem Wizard-State gestrippt (v28.39: „wird nur im Organizer Center
      // gepflegt"), der Blob hier wird aber komplett FRISCH zusammengebaut.
      // Ohne diesen Carry-Forward löschte JEDER Wizard-Save eines Events die
      // gesamte Hotel-Konfiguration (_hotels/_hotelStays/_hotelVisible/
      // _hotelRules) — Hotels, Zeiträume, Verteil-Regeln, Sichtbarkeit.
      // Gleiche Mechanik wie imageOrigUrlConfig (v28.11), nur aus dem rohen
      // Overrides-JSON des editEvent, weil die Werte im State nicht existieren.
      const hotelCarryConfig = ((): Record<string, unknown> => {
        try {
          const raw = JSON.parse(editEvent?.emailTemplateOverrides || '{}') as Record<string, unknown>;
          const out: Record<string, unknown> = {};
          for (const k of ['_hotels', '_hotelStays', '_hotelVisible', '_hotelRules']) {
            if (raw && raw[k] !== undefined) out[k] = raw[k];
          }
          return out;
        } catch { return {}; }
      })();
      const topPiggybackConfigs: Array<Record<string, unknown>> = [
        b2runExtraConfig, qrScannerConfig, coOrganizerConfig, testTeamConfig,
        splitDispRevConfig, requireSubEventConfig, subEventsOnlyConfig,
        subEventsDisabledConfig, imageBannerConfig, imageOrigUrlConfig, klammerDeadlineConfig, childTermConfig, teamTermConfig,
        teamNoCreateConfig, mainEventLabelConfig, assistantsCanSeeConfig,
        organizerDisplayLargeConfig, previewBeforeActiveConfig,
        imageDisplayConfig, hideOrganizerConfig, hiddenOrganizersConfig,
        hideOrgIndividualConfig, headerImageLayoutConfig, noDescriptionConfig,
        subEventCalendarConfig, subEventSingleChoiceConfig, bundledCommConf, noSelfCancelConfig, noCancelAfterDeadlineConfig, teamsLinkConfig, hotelCarryConfig,
        billingPiggyback(), // v29.66: F&A-Pilot
        subEventOpenRulePiggyback(), // v29.67
        visAllSubsPiggyback(), // v29.75
        subDeadlineRulePiggyback(), // v29.76
      ];
      updates['EmailTemplateOverrides'] = (Object.keys(topOverrides).length > 0 || !!effEmailLogo || !!effOutlookLogo || topPiggybackConfigs.some(o => Object.keys(o).length > 0))
        // v28.2: Object.assign statt Spread-Kette — die Literal-Spreads
        // überschritten TypeScripts Union-Komplexitätslimit (TS2590).
        // Piggyback-Keys sind disjunkt, die Reihenfolge im Array entspricht
        // der alten Spread-Reihenfolge; topOverrides bleibt LETZTER Merge.
        ? JSON.stringify(Object.assign(
            {},
            (effEmailLogo ? { _eventLogo: effEmailLogo } : {}),
            outlookLogoPiggyback(effEmailLogo, effOutlookLogo),
            ...topPiggybackConfigs,
            topOverrides,
          ))
        : '';
      // v9.21: ActiveFrom als SP-DateTime
      // v30.66: berlinLocalToUtcIso wie bei allen anderen Fristen (siehe activeFrom-State).
      updates['ActiveFrom'] = activeFrom ? (berlinLocalToUtcIso(activeFrom) || null) : null;
      // Custom-Mail-Logo in EmailImageBase64 (SP-Spalte) — der Flow ersetzt
      // {{ORB_URL}} in Mails damit. Wenn leer: Flow fällt auf _Config
      // DefaultImageBase64 (DEX-Orb) zurück.
      updates['EmailImageBase64'] = effEmailLogo || '';
      updates['DisableEmails'] = effDisableEmails;
      // v19.22: granulare An-/Abmelde-Mail-Schalter des Hauptevents (top-level
      // aufgelöst, damit ein Save von einem Sub-Tab nicht den Sub-Wert aufs
      // Hauptevent schreibt). Pro Sub-Event werden sie in persistSubEventsForParent
      // geschrieben.
      updates['DisableRegistrationEmail'] = effDisableRegistrationEmail;
      updates['DisableCancellationEmail'] = effDisableCancellationEmail;
      // v19.23: Outlook-Absage = Auto-Abmeldung (Top-Level-Event).
      updates['AutoDeregisterOnDecline'] = effAutoDeregisterOnDecline;
      updates['InactiveHandling'] = effInactiveHandling;
      updates['DisableOutlook'] = effDisableOutlook;
      // v11.57: OutlookDirty schreiben. Wenn Outlook-relevante Änderungen
      // anstehen und der Organizer im Update-Confirm-Modal die Checkbox
      // *nicht* gesetzt hat, bleibt der Flag dirty=true; bei Checkbox=true
      // (= UpdateEvent wird gequeued) wird dirty wieder auf false gesetzt.
      // Wenn keine Outlook-relevante Änderung vorlag (z.B. nur Beschreibung
      // angepasst), wird dirty NICHT angefasst — der Wert bleibt wie er war.
      // Den eigentlichen Wert setzen wir aus dem Modal-State unten (siehe
      // pendingOutlookDirtyWrite).
      if (pendingOutlookDirtyWriteRef.current !== null) {
        updates['OutlookDirty'] = pendingOutlookDirtyWriteRef.current;
      }
      updates['NotifyOrgRegisterMode'] = notifyOrgRegisterMode === 'always' ? 'Always' : notifyOrgRegisterMode === 'fromDate' ? 'FromDate' : 'Never';
      updates['NotifyOrgRegisterFromDate'] = notifyOrgRegisterMode === 'fromDate' && notifyOrgRegisterFromDate ? berlinLocalToUtcIso(notifyOrgRegisterFromDate) : null;
      updates['NotifyOrgCancelMode'] = notifyOrgCancelMode === 'always' ? 'Always' : notifyOrgCancelMode === 'afterDeadline' ? 'AfterDeadline' : 'Never';
      updates['ExcludedUsers'] = excludedUsers.filter(Boolean).join(';');
      updates['IsFictive'] = isFictive;
      // v22.15: Auto-Heilung — steht das Event auf „Abgeschlossen" (z.B. vom
      // Auto-Cleanup wegen eines alten Testdatums), das End-Datum liegt nach
      // dieser Bearbeitung aber in der Zukunft, zurück auf Aktiv setzen.
      // Sonst bleibt das Event unsichtbar, obwohl es noch bevorsteht.
      if (editEvent && editEvent.status === 'Completed') {
        const newEndIso = endDate ? berlinLocalToUtcIso(endDate) : (startDate ? berlinLocalToUtcIso(startDate) : '');
        if (newEndIso && new Date(newEndIso).getTime() > Date.now()) {
          updates['EventStatus'] = 'Active';
        }
      }
      if (useSplitCapacities) {
        updates['DurchstarterCapacity'] = parseInt(durchstarterCapacity, 10) || 0;
        updates['FunstarterCapacity'] = parseInt(funstarterCapacity, 10) || 0;
        // v10.20: frei wählbare Bezeichnungen mitschreiben — leer = Default-
        // Fallback in der Registration-UI ('Durchstarter' / 'Funstarter').
        updates['SplitLabelA'] = (splitLabelA || '').trim();
        updates['SplitLabelB'] = (splitLabelB || '').trim();
        updates['SplitDescA'] = (splitDescA || '').trim();
        updates['SplitDescB'] = (splitDescB || '').trim();
        updates['SplitHelpText'] = (splitHelpText || '').trim();
        updates['SplitSectionTitle'] = (splitSectionTitle || '').trim();
        updates['SplitSharedWaitlist'] = !!splitSharedWaitlist;
      } else {
        // Split deaktiviert: Kapazitäten nullen + Labels leer setzen, damit
        // die Registration-Logik nicht irrtümlich den Split-Pfad nimmt.
        updates['DurchstarterCapacity'] = null;
        updates['FunstarterCapacity'] = null;
        updates['SplitLabelA'] = '';
        updates['SplitLabelB'] = '';
        updates['SplitDescA'] = '';
        updates['SplitDescB'] = '';
        updates['SplitHelpText'] = '';
        updates['SplitSectionTitle'] = '';
        updates['SplitSharedWaitlist'] = false;
      }
      // v11.0: Teilnehmer-Upload-Setting
      updates['AllowAttendeeUpload'] = !!allowAttendeeUpload;
      updates['AttendeeUploadHint'] = (attendeeUploadHint || '').trim();
      updates['AttendeeUploadLabel'] = (attendeeUploadLabel || '').trim();
      // v11.80: Anrede-Toggle + Team-Anmeldung-Konfiguration mit-persistieren.
      updates['AskSalutation'] = !!askSalutation;
      // v18.75: Sicherheitshinweis vor dem Absenden mit-persistieren.
      updates['ConfirmDialogEnabled'] = !!confirmDialogEnabled;
      updates['ConfirmDialogMode'] = confirmDialogEnabled ? (confirmDialogMode || 'summary') : '';
      updates['ConfirmDialogText'] = confirmDialogEnabled && confirmDialogMode === 'freetext' ? confirmDialogText : '';
      // v18.33: Self-Check-in mit-persistieren. Token nur schreiben, wenn aktiv.
      // v20.2: SelfCheckIn*-Spalten werden vom Wizard-Edit bewusst NICHT mehr
      // geschrieben — sie gehören jetzt dem Admin Center (Auto-Aktivierung +
      // Kachel-Modal). Würde der Wizard sie weiter mitschreiben, würde jeder
      // Wizard-Save die dort gesetzten Werte zurücksetzen.
      updates['TeamRegistrationEnabled'] = !!teamRegistrationEnabled;
      updates['TeamSize'] = teamRegistrationEnabled && teamSize > 0 ? teamSize : null;
      updates['AskTeamName'] = !!askTeamName;
      // v11.81: Erweiterte Team-Konfiguration mit-persistieren.
      updates['TeamPartialAllowed'] = !!(teamRegistrationEnabled && teamPartialAllowed);
      updates['TeamOpenSlotsVisible'] = !!(teamRegistrationEnabled && teamOpenSlotsVisible);
      updates['TeamJoinRequiresApproval'] = !!(teamRegistrationEnabled && teamOpenSlotsVisible && teamJoinRequiresApproval);
      // v17.20: Bilingual-Toggle persistieren.
      updates['BilingualFields'] = !!bilingualFields;

      // v11.22: feinere Progress-Stufen während Edit-Save. Vorher
      // sprang es bei 50% sehr lange auf der Stelle, weil zwischen
      // setProgress(50) und setProgress(100) die Dokument-Sync,
      // updateEvent, Berechtigungs-Sync, Sub-Event-Persistierung,
      // Teilnehmer-Spalten-Sync, Bild-Upload und Outlook-Update
      // nacheinander liefen — alles ohne Zwischen-Tick.
      setProgress(40);
      setProgressLabel(isDe ? 'Dokumente werden synchronisiert...' : 'Syncing documents...');

      // v29.21 (Audit): Der Dokument-Sync (Attachments löschen/hochladen)
      // lief hier VOR dem updateEvent — schlug der Save danach fehl
      // („Invalid text value", 412), blieb der Wizard mit Fehlermeldung
      // offen, die Löschungen waren aber schon unwiderruflich draußen.
      // Der Sync steht jetzt IM Erfolgszweig (s.u., syncEventDocuments) —
      // dieselbe Regel wie beim Löschkonzept: Erst der prüfbare Schritt,
      // dann der unumkehrbare.
      const syncEventDocuments = async (): Promise<void> => {
        if (!selectedEventId) return;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ctx = (window as any).__dexSpfxContext;
        if (!ctx) return;
        const svc = new EventService(ctx);
        // Namen die weiterhin vorhanden sind (als bestehende Attachments, ohne file);
        // erst löschen, dann uploaden (SharePoint verbietet Duplikat-Namen).
        const keptOriginalNames = new Set(
          documents.filter(d => !d.file).map(d => d.name)
        );
        const toDelete = initialDocumentNames.filter(name => !keptOriginalNames.has(name));
        for (const fileName of toDelete) {
          try {
            await svc.deleteEventDocument(Number(selectedEventId), fileName);
          } catch { /* einzelner Delete-Fehler darf Save nicht blockieren */ }
        }
        for (const doc of documents) {
          if (doc.file) {
            try {
              await svc.uploadEventDocument(Number(selectedEventId), doc.file);
            } catch { /* einzelner Upload-Fehler darf Save nicht blockieren */ }
          }
        }
      };

      setProgress(55);
      setProgressLabel(isDe ? 'Event-Daten werden gespeichert...' : 'Saving event data...');

      // v11.20: Direkt vor dem updateEvent-Call loggen was am SP-Server
      // landet. Damit sehen wir im Browser-DevTools:
      //   1. ob updates['CustomFields'] als JSON-String den helpText
      //      enthält (= Save sendet's korrekt → SP-Persist OK).
      //   2. oder ob updates['CustomFields'] ohne helpText/onlyForGroup
      //      ankommt (= State zum Save-Zeitpunkt war schon kaputt).
      // v29.77: Debug-Log („updates.CustomFields about to POST") entfernt.
      const success = await updateEvent(selectedEventId, updates);
      if (success) {
        // v26.57: NEU zur Zielgruppe hinzugekommene Personen außerhalb von
        // @deloitte.de → Approve-Mail an die Admins (SharePoint ist im Default
        // nur für Deloitte DE ALL freigeschaltet; internationale Kolleg:innen
        // brauchen zusätzlich Site-Zugriff). Nur der Diff gegen den vorherigen
        // Stand, damit nicht bei jedem Save erneut gemailt wird. Fire-and-forget.
        try {
          const prevAudLc = new Set((editEvent?.audienceFilter || []).map(a => (a || '').trim().toLowerCase()));
          const addedNonDe = audience.split(',')
            .map(s => s.trim())
            .filter(a => a.indexOf('@') > 0 && !a.toLowerCase().endsWith('@deloitte.de') && !prevAudLc.has(a.toLowerCase()));
          if (addedNonDe.length > 0) {
            void notifyAdminsExternalAudienceAccess(title, addedNonDe, `${currentUser.firstName} ${currentUser.surname}`.trim()).catch(() => { /* */ });
          }
        } catch { /* darf den Save nie stören */ }

        try { await syncEventDocuments(); }
        catch (err) { console.warn('[DEX] Dokument-Sync fehlgeschlagen:', err); }
        // v29.18: Bild-Upload VOR den Sub-Events — wie im Create-Pfad (v29.17).
        // Der Block stand bei Fortschritt 90, HINTER persistSubEventsForParent:
        // Wer im Bearbeiten den Kalender aktiviert und 20+ Tage anklickt, löst
        // dort ebenso viele Subsite-Anlagen aus; der Upload lief danach in die
        // gedrosselte SharePoint-Instanz und scheiterte — die Meldung dazu war
        // eine kleine rote Zeile in Schritt 1, verdeckt vom Fortschritts-
        // Fenster bei 95 %. „Der Edit-Pfad war abgesichert" stimmte nur
        // strukturell (eigener try/catch), nicht praktisch: Er hing an
        // derselben Drossel und scheiterte genauso still.
        if (imageFile) {
          try {
            setProgressLabel(isDe ? 'Bild wird hochgeladen...' : 'Uploading image...');
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctx = (window as any).__dexSpfxContext;
            if (ctx) {
              const svc = new EventService(ctx);
              const compressed = await compressImage(imageFile);
              const uploadedUrl = await svc.uploadEventImageAsAttachment(Number(selectedEventId), compressed);
              if (uploadedUrl) {
                await svc.updateEventImageUrl(Number(selectedEventId), uploadedUrl);
                // v28.11: Querformat-Original als zweites Attachment sichern,
                // wenn der Zuschnitt es rund/quadratisch gemacht hat — die
                // Event-Liste zeigt dann das Original. Sonst evtl. vorhandenes
                // Alt-Original aufräumen. (patchEventOverridesKey liest den
                // frisch von updateEvent geschriebenen Overrides-Stand und
                // patcht nur den einen Schlüssel — Reihenfolge passt.)
                try {
                  if (imageOrigFile && (imageOrigAspect || 0) >= 1.2 && wizardImgAspect != null && wizardImgAspect < 1.2) {
                    const origCompressed = await compressImage(imageOrigFile, 1600, 0.85);
                    const origUrl = await svc.uploadEventOrigImageAsAttachment(Number(selectedEventId), origCompressed);
                    if (origUrl) await svc.patchEventOverridesKey(Number(selectedEventId), '_imageOrigUrl', origUrl);
                  } else if (imageOrigFile) {
                    // v28.12: Nur aufräumen, wenn wir die QUELLE des neuen
                    // Bilds kennen (frischer Upload/Capture) und sie kein
                    // Original braucht. Ohne imageOrigFile (Re-Crop eines
                    // bereits runden Bestands) bleibt ein gespeichertes
                    // Original unangetastet.
                    await svc.deleteEventOrigImageAttachment(Number(selectedEventId));
                    await svc.patchEventOverridesKey(Number(selectedEventId), '_imageOrigUrl', '');
                  }
                } catch (origErr) { console.warn('[DEX] Original-Bild speichern fehlgeschlagen:', origErr); }
                // Events neu laden, damit die UI das frische Bild ohne Hard-Refresh
                // anzeigt (updateEvent oben hat schon geladen, aber da war
                // EventImageUrl noch der alte Wert). v29.18: In einem EIGENEN
                // try — ein Refresh-Fehler ist kein Upload-Fehler; vorher
                // meldete er „Bild-Upload fehlgeschlagen", obwohl das Bild
                // längst gespeichert war.
                try { await refreshEvents(); }
                catch (rErr) { console.warn('[DEX] Refresh nach Bild-Upload fehlgeschlagen (Bild ist gespeichert):', rErr); }
              } else {
                setImageUploadError('Bild-Upload fehlgeschlagen.');
                showAlert(isDe
                  ? 'Das Event-Bild konnte nicht hochgeladen werden. Alle übrigen Änderungen werden gespeichert — bitte lade das Bild danach in Schritt 1 erneut hoch.'
                  : 'The event image could not be uploaded. All other changes are being saved — please upload the image again in step 1 afterwards.', { variant: 'error' });
              }
            }
          } catch (err) {
            console.warn('[DEX] Bild-Upload fehlgeschlagen', err);
            setImageUploadError('Bild-Upload fehlgeschlagen.');
            // v29.18: Sichtbar melden statt nur der roten Zeile in Schritt 1 —
            // die verdeckt das Fortschritts-Fenster, und der Save endet mit
            // „Änderungen gespeichert!", obwohl das Bild fehlt.
            showAlert(isDe
              ? 'Das Event-Bild konnte nicht hochgeladen werden. Alle übrigen Änderungen werden gespeichert — bitte lade das Bild danach in Schritt 1 erneut hoch.'
              : 'The event image could not be uploaded. All other changes are being saved — please upload the image again in step 1 afterwards.', { variant: 'error' });
          }
        }

        setProgress(65);
        setProgressLabel(isDe ? 'Berechtigungen werden gesetzt...' : 'Setting permissions...');
        // v9.35: Berechtigungs-Sync — beim Edit können neue Co-Organizer hinzugekommen
        // sein, die bisher nur in EmailTemplateOverrides._coOrganizers stehen, aber
        // noch keine SharePoint-Berechtigung auf Subsite + Teilnehmerliste haben.
        // ensureOrganizerPermissions ist idempotent: bestehende Rechte werden nicht
        // doppelt vergeben, neue kommen sauber dazu.
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ctxPerm = (window as any).__dexSpfxContext;
          if (ctxPerm && editEvent?.subsiteUrl) {
            const svcPerm = new EventService(ctxPerm);
            const allOrgEmailsForPerm = [
              organizerEmails.join(';'),
              coOrganizerEmails.join(';'),
            ].filter(Boolean).join(';');
            if (allOrgEmailsForPerm) {
              // v30.37: Klammer UND alle Sub-Events. Jeder Termin hat eine
              // eigene Subsite mit eigener Teilnehmerliste — bis v30.36 lief
              // der Sync nur über die Klammer. Ein nachträglich benannter
              // Co-Organizer konnte die Klammer sehen und KEINEN einzigen
              // Termin; weil getAllRegistrations bei 403 `[]` liefert, kam das
              // in der App als „0 Teilnehmer" an statt als Fehler.
              const permSites = [editEvent.subsiteUrl]
                .concat(childEventsOf(editEvent.id).map(k => k.subsiteUrl || ''))
                .filter(Boolean);
              await svcPerm.ensureOrganizerPermissionsMulti(permSites, allOrgEmailsForPerm);
            }
          }
        } catch (err) { console.warn('[DEX] Permission-Sync für Organizer fehlgeschlagen:', err); }

        // v26.24: Co-Organizer-Freigabe. Benannte Organizer, die noch KEIN
        // Organizer/Admin sind, brauchen Schreibrecht auf DEX_Events (kommt aus
        // der Members-Gruppe, nicht aus dem Organizer-Eintrag). Das kann ein
        // normaler Organizer nicht selbst vergeben → wir legen pro solcher Person
        // einen „Organizer werden"-Antrag an; die Admins bekommen die Mail mit
        // Deep-Link und geben frei. Best-effort, blockt den Save nicht.
        try {
          await requestCoOrganizerApprovals(sanitizedOrgPairEdit.orgString, sanitizedOrgPairEdit.orgEmailString, title);
        } catch (err) { console.warn('[DEX] Co-Organizer-Freigabe-Anträge fehlgeschlagen:', err); }

        // v26.34: Neu hinzugefügte (Co-)Organizer benachrichtigen + Outlook-
        // Kalendereinladung. „Neu" = im gespeicherten Organizer-Set, aber vorher
        // (editEvent.organizerEmails) NICHT enthalten. Best-effort, blockt nie.
        try {
          const prevSet = new Set((editEvent?.organizerEmails || []).map(e => (e || '').toLowerCase().trim()).filter(Boolean));
          const meLc = (currentUser.email || '').toLowerCase();
          const newEmails = (sanitizedOrgPairEdit.orgEmailString || '').split(';').map(s => s.trim());
          const newNames = (sanitizedOrgPairEdit.orgString || '').split(';').map(s => s.trim());
          const addedCoOrgs = newEmails
            .map((email, i) => ({ email, name: newNames[i] || email }))
            .filter(p => p.email && p.email.indexOf('@') > 0 && !prevSet.has(p.email.toLowerCase()) && p.email.toLowerCase() !== meLc);
          if (addedCoOrgs.length > 0) {
            const evLangDe = (editEvent?.emailLanguage || 'EN').toUpperCase() === 'DE';
            await notifyNewCoOrganizers(selectedEventId, title, addedCoOrgs, evLangDe, !!editEvent?.disableOutlook);
          }
        } catch (err) { console.warn('[DEX] Co-Organizer-Benachrichtigung fehlgeschlagen:', err); }

        setProgress(75);
        setProgressLabel(isDe ? 'Sub-Events werden gespeichert...' : 'Saving sub-events...');
        // Sub-Events persistieren (create/update/delete pro Draft). Seit v6.4.
        // v28.98: Der Abschnitt 75–82 % gehört den Sub-Events. Jeder gespeicherte
        // Termin schiebt den Balken ein Stück und sagt, welcher gerade dran war
        // („3 von 9 …") — sonst steht er bei neun Terminen minutenlang auf 75 %.
        try {
          await persistSubEventsForParent(selectedEventId, (done, total, subTitle) => {
            setProgress(75 + Math.round((done / Math.max(total, 1)) * 7));
            setProgressLabel(isDe
              ? `Sub-Events werden gespeichert… ${done} von ${total}${subTitle ? ` — ${subTitle}` : ''}`
              : `Saving sub-events… ${done} of ${total}${subTitle ? ` — ${subTitle}` : ''}`);
          });
        }
        catch (err) { console.warn('[DEX] Sub-Events persistieren fehlgeschlagen:', err); }
        // v29.22: Die orange Vormerkliste ist nach dem Save Geschichte — die
        // Termine wurden (nach Rückfrage) gelöscht; stehen gebliebene Marker
        // würden beim nächsten Öffnen Geister anzeigen.
        setRemovedSavedSubs([]);

        setProgress(82);
        setProgressLabel(isDe ? 'Teilnehmerlisten-Spalten werden geprüft...' : 'Verifying participant list columns...');
        // Custom-Fields-Columns auf der Teilnehmerliste auto-sync: falls
        // neue Custom-Fields ohne spInternalName hinzugekommen sind oder
        // SP-Spalten fehlen, jetzt anlegen + spInternalName ins Event
        // zurückschreiben.
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ctx = (window as any).__dexSpfxContext;
          if (ctx && editEvent?.subsiteUrl) {
            const svc = new EventService(ctx);
            const cfForFix = customFields
              .filter(f => f.label && f.label.trim().length > 0)
              .map(f => ({
                id: f.id, label: f.label.trim(), type: f.type, required: f.required, visible: f.visible,
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                spInternalName: (f as any).spInternalName || '',
                ...(f.type === 'select' ? { options: f.options.map(o => o.trim()).filter(Boolean), ...(f.multi ? { multi: true } : {}) } : {}),
                ...(f.onlyForGroup && f.onlyForGroup !== 'all' ? { onlyForGroup: f.onlyForGroup } : {}),
                // v11.15: externalLinks (AGB/Datenschutz-URLs etc.) beim Save
                // mit-persistieren — vorher haben alle drei Persist-Pfade
                // (Edit-Save, Create-Save, Sub-Event-Save) sie gedroppt.
                ...(f.externalLinks && f.externalLinks.length > 0
                  ? { externalLinks: f.externalLinks.map(x => ({ label: x.label, url: x.url })) }
                  : {}),
                // v11.21 KRITISCHER BUG-FIX: helpText UND showIf wurden hier
                // beim cfForFix-Mapping nicht mit-übernommen — der zweite
                // updateEvent-Call (siehe `merged`-JSON unten, überschreibt
                // CustomFields nochmal mit der spInternalName-Anreicherung)
                // hat die helpText- und showIf-Properties wieder vom SP-
                // Item entfernt. Folge: jede gespeicherte Beschreibung war
                // direkt nach dem Save wieder weg, obwohl der erste
                // updateEvent sie korrekt geschrieben hatte.
                ...(f.helpText && f.helpText.trim() ? { helpText: f.helpText.trim() } : {}),
                // v18.20 BUG-FIX (gleiches Muster wie v11.21): helpTextStyle
                // wurde hier nicht mit-übernommen — der zweite updateEvent-Call
                // (merged-JSON unten) hat die Property wieder vom SP-Item
                // entfernt. Folge: „Text unter dem Feld-Titel" war direkt nach
                // dem Speichern wieder weg und das Feld zeigte die „i"-Box.
                ...(f.helpTextStyle === 'inline' ? { helpTextStyle: 'inline' as const } : {}),
                ...(f.showIf && f.showIf.fieldId && f.showIf.values && f.showIf.values.length > 0
                  ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } }
                  : {}),
                // v11.94: confirmLabel auch im cfForFix-Mapping mit-übernehmen,
                // sonst überschreibt der zweite updateEvent-Call die Property weg.
                ...(f.type === 'checkbox' && f.confirmLabel && f.confirmLabel.trim()
                  ? { confirmLabel: f.confirmLabel.trim() }
                  : {}),
                // v18.41 (gleiches Muster wie v11.21/v18.20): ccOnEmails muss
                // auch im zweiten Write mit, sonst droppt der spInternalName-
                // Merge die Property direkt nach dem Speichern wieder.
                ...((f.type === 'user' || f.type === 'roommate') && f.ccOnEmails ? { ccOnEmails: true } : {}),
                ...(f.type === 'roommate' && f.notifyRoommate === false ? { notifyRoommate: false } : {}),
                ...((f.type === 'user' || f.type === 'roommate') && f.audienceOnly ? { audienceOnly: true } : {}),
              }));
            // v11.6 BUG-FIX: vorher wurde hier `isB2runTemplate` (= b2run_*-
            // Custom-Fields vorhanden) als Indikator genutzt. Das war falsch,
            // sobald die generische Split-Capacity ohne B2Run-Template
            // genutzt wird — dann hat das Event Split-Kapazitäten + StarterType-
            // Werte in der Teilnehmerliste, aber `isB2runTemplate=false`. Der
            // Fix-Lauf hat daraufhin StarterType + PreferredStarterType
            // gelöscht und die Teilnehmer-Daten weggeworfen. Korrekter Check:
            // entweder altes B2Run-Template ODER Split-Capacity aktiv.
            const splitActive = useSplitCapacities && ((parseInt(durchstarterCapacity, 10) || 0) > 0 || (parseInt(funstarterCapacity, 10) || 0) > 0);
            const fixResult = await svc.fixRegistrationListColumns(editEvent.subsiteUrl, {
              isB2Run: isB2runTemplate || splitActive,
              hasQuiz: quiz.length > 0,
              customFields: cfForFix,
            });
            if (fixResult.customFieldMap && Object.keys(fixResult.customFieldMap).length > 0) {
              // v19.20: ROBUSTER FIX für den wiederkehrenden „zweiter
              // CustomFields-Write droppt Properties"-Bug (siehe ENTWICKLUNG.md).
              // Früher wurde dieser zweite Write aus dem manuell gepflegten
              // cfForFix-Mapping gebaut — jede dort vergessene Property wurde
              // damit direkt nach dem Speichern wieder vom SP-Item entfernt
              // (zuletzt die EN-Varianten labelEn/helpTextEn/confirmLabelEn/
              // optionsEn; historisch multi/ccOnEmails/helpText/showIf/…).
              // Jetzt nehmen wir den KANONISCHEN serializeCustomFields-Output
              // (der ALLE Properties korrekt persistiert) und ergänzen pro
              // Feld nur noch spInternalName. So kann die Property-Liste nie
              // wieder veralten — cfForFix dient ab jetzt ausschließlich dem
              // Spalten-Fix-Aufruf oben, nicht mehr der Persistenz.
              const spById: Record<string, string> = {};
              for (const f of customFields) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                spById[f.id] = (f as any).spInternalName || '';
              }
              const merged = serializeCustomFields(customFields, bilingualFields).map(f => ({
                ...f,
                spInternalName: fixResult.customFieldMap![f.id] || spById[f.id] || '',
              }));
              await updateEvent(selectedEventId, { 'CustomFields': JSON.stringify(merged) });
            }
            // v30.60: DIESELBE Behandlung für die Sub-Events.
            //
            // Bisher lief der Spalten-Abgleich beim Speichern nur über die
            // Klammer. Sub-Event-Listen bekamen ihre Spalten einmal bei der
            // Anlage und danach nie wieder — wer später ein Abfragefeld
            // ergänzt, hat auf jeder Termin-Liste eine fehlende Spalte. Und
            // eine fehlende Spalte lässt SharePoint nicht etwa den Wert
            // weglassen, sondern lehnt den GANZEN Insert ab: Die Anmeldung
            // scheitert, sobald jemand dieses Feld ausfüllt. Der Befund aus
            // dem Bestand (12 fehlende Spalten über 100 Listen) ist genau das.
            for (const sub of childEventsOf(editEvent.id)) {
              if (!sub.subsiteUrl) continue;
              try {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const subCf = ((sub as any).eventSpecificFields || []).map((f: any) => ({
                  id: f.id, label: f.label, type: f.type, required: f.required, options: f.options,
                  visible: true, spInternalName: f.spInternalName || '',
                }));
                await svc.fixRegistrationListColumns(sub.subsiteUrl, {
                  isB2Run: !!(sub.durchstarterCapacity || sub.funstarterCapacity),
                  hasQuiz: !!(sub.quiz && sub.quiz.length > 0),
                  customFields: subCf,
                });
              } catch (e) { console.warn('[DEX] Spalten-Abgleich für Sub-Event fehlgeschlagen:', sub.id, e); }
            }

            // v30.60: NACHSEHEN und MELDEN, statt still weiterzugehen.
            //
            // Der ganze Block lief bisher in einem `catch`, das nur in die
            // Konsole schrieb. Scheiterte der Spalten-Abgleich, meldete das
            // Speichern trotzdem Erfolg — und das Problem fiel erst auf, wenn
            // Wochen später eine Anmeldung scheiterte. Wer speichert, muss
            // erfahren, dass etwas nicht sitzt; sonst wird aus einem
            // behebbaren Zustand ein unerklärlicher Ausfall.
            const stillBroken: string[] = [];
            for (const target of [editEvent, ...childEventsOf(editEvent.id)]) {
              if (!target.subsiteUrl) continue;
              try {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const flds = ((target as any).eventSpecificFields || []).map((f: any) => ({ id: f.id, label: f.label, spInternalName: f.spInternalName || '' }));
                const d = await svc.diagnoseRegistrationList(target.subsiteUrl, flds);
                if (d.missingColumns.length > 0) {
                  stillBroken.push(`${target.title}: ${d.missingColumns.map(m => m.label).join(', ')}`);
                }
              } catch { /* nicht lesbar — dann bleibt es beim Konsolen-Hinweis */ }
            }
            if (stillBroken.length > 0) {
              showAlert(
                (isDe
                  ? 'Achtung: Für diese Abfragefelder konnte die Spalte in der Teilnehmerliste NICHT angelegt werden:\n\n'
                  : 'Warning: the column could NOT be created for these fields:\n\n')
                + stillBroken.join('\n')
                + (isDe
                  ? '\n\nSolange das so ist, scheitert jede Anmeldung, bei der jemand eines dieser Felder ausfüllt. Bitte im Admin Center „Spalten fixen (alle Events)“ ausführen oder das Speichern wiederholen.'
                  : '\n\nUntil this is fixed, every registration that fills one of these fields will fail. Please run "Fix columns (all events)" in the Admin Center.'),
                { variant: 'error' }
              );
            }
          }
        } catch (err) {
          console.warn('[DEX] Auto-fix Teilnehmer-Columns fehlgeschlagen:', err);
          // v30.60: Auch der harte Fehlschlag darf nicht stumm bleiben.
          showAlert(isDe
            ? 'Der Abgleich der Teilnehmerlisten-Spalten ist fehlgeschlagen. Das Event ist gespeichert, aber neu angelegte Abfragefelder haben womöglich keine Spalte — dann scheitern Anmeldungen, bei denen sie ausgefüllt werden. Bitte im Admin Center „Spalten fixen (alle Events)“ ausführen.'
            : 'Syncing the participant list columns failed. The event is saved, but newly added fields may have no column.', { variant: 'error' });
        }

        // v29.77: Der Block hier war ein stummer Sammelposten — „Outlook wird
        // aktualisiert… 95%" stand minutenlang ohne Angabe, WELCHER Termin
        // gerade dran ist (und seit v29.74 wartet die Drossel-Logik ehrlich,
        // also auch mal lange). Jetzt: Balken startet bei 90, jeder Queue-
        // Eintrag rueckt ihn vor und der Label nennt Termin und Zaehler.
        setProgress(90);
        setProgressLabel(isDe ? 'Outlook wird aktualisiert...' : 'Updating Outlook...');
        // v30.67: `effDisableOutlook` statt des rohen States — auf einem
        // Sub-Reiter trägt `disableOutlook` den Wert DIESES Sub-Events.
        const outlookTotal =
          ((!effDisableOutlook && pendingOutlookUpdateForTopRef.current) ? 1 : 0)
          + pendingOutlookUpdateForSubEventsRef.current.length
          + (orgGetsSubInvites !== initialOrgGetsSubInvitesRef.current ? (1 + subEventsRef.current.filter(se => !!se.dbId && !se.disableOutlook).length) : 0);
        let outlookDone = 0;
        const tickOutlook = (label: string): void => {
          outlookDone++;
          setProgress(90 + Math.min(9, Math.round((outlookDone / Math.max(1, outlookTotal)) * 9)));
          setProgressLabel(label);
        };
        // v11.63: Outlook-Updates pro Event entscheiden — der Organizer hat
        // im Confirm-Modal pro betroffenem Event (Top + Sub) einzeln ent-
        // schieden. Top-Event bekommt UpdateEvent nur, wenn explizit
        // angehakt (pendingOutlookUpdateForTopRef.current). OutlookDirty
        // wurde für das Top-Event schon im updateEvent-Call oben mit
        // pendingOutlookDirtyWriteRef geschrieben.
        // v30.67: `effDisableOutlook` statt `disableOutlook` — genau die
        // v11.93-Falle: Speichert der Organizer von einem Sub-Reiter aus,
        // dessen Termin „Outlook deaktiviert" hat, stand hier der Sub-Wert,
        // der Zweig wurde übersprungen, KEIN UpdateEvent landete in der Queue
        // — und OutlookDirty war oben schon auf false geschrieben. Der
        // Kalender der Teilnehmer behielt den alten Stand, und der Wizard bot
        // die Aktualisierung nie wieder an. Die Erkennung (outlookChanges.ts)
        // löste längst über resolveTopLevelCommState auf; nur dieser Zweig
        // und die Zählung darüber lasen noch roh.
        // v30.67: Queue-Ergebnis PRÜFEN — queueOutlookEvent wirft nie, es
        // meldet Fehlschlag (429) als false. Dann den Dirty-Marker zurück auf
        // true setzen, damit der nächste Wizard-Lauf die offene Aktualisierung
        // wieder anbietet, und den Termin am Ende namentlich melden.
        const failedOutlookTitles: string[] = [];
        if (!effDisableOutlook && pendingOutlookUpdateForTopRef.current) {
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctx = (window as any).__dexSpfxContext;
            if (ctx && editEvent?.subsiteUrl) {
              const svc = new EventService(ctx);
              tickOutlook(isDe ? `Outlook wird aktualisiert… (${title || 'Hauptevent'})` : `Updating Outlook… (${title || 'main event'})`);
              const topQueued = await svc.queueOutlookEvent('', selectedEventId, title, 'UpdateEvent');
              if (!topQueued) {
                failedOutlookTitles.push(title || (isDe ? 'Hauptevent' : 'main event'));
                const restored = await updateEvent(selectedEventId, { 'OutlookDirty': true }, { skipReload: true });
                if (!restored) console.warn('[DEX][v30.67] OutlookDirty konnte nach gescheitertem UpdateEvent nicht zurückgesetzt werden:', selectedEventId);
              }
            }
          } catch (err) { console.warn('[DEX][v30.67] Outlook-Update des Hauptevents fehlgeschlagen:', err); failedOutlookTitles.push(title || (isDe ? 'Hauptevent' : 'main event')); }
        }
        // v11.63: Sub-Event-Outlook-Updates pro angehaktem Sub-Event.
        if (pendingOutlookUpdateForSubEventsRef.current.length > 0) {
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctx = (window as any).__dexSpfxContext;
            if (ctx) {
              const svc = new EventService(ctx);
              const subIds = pendingOutlookUpdateForSubEventsRef.current;
              for (let subIdx = 0; subIdx < subIds.length; subIdx++) {
                const subId = subIds[subIdx];
                const subDraft = subEventsRef.current.find(s => s.dbId === subId);
                const subTitle = subDraft?.title || '';
                tickOutlook(isDe
                  ? `Outlook wird aktualisiert… (${subIdx + 1}/${subIds.length}: ${subTitle || 'Termin'})`
                  : `Updating Outlook… (${subIdx + 1}/${subIds.length}: ${subTitle || 'date'})`);
                // v30.67: Rückgabewert prüfen — queueOutlookEvent WIRFT nicht,
                // es liefert false (429 nach ~20 Schreibzugriffen). Vorher
                // wurde OutlookDirty trotzdem gelöscht: kein Eintrag in
                // DEX_Outlook, Kalender alt, und der Wizard meldete beim
                // nächsten Öffnen keine offene Änderung mehr. Dieselbe
                // Klasse wie v29.21 (updateEvent) und v29.48 (deleteEvent).
                try {
                  const subQueued = await svc.queueOutlookEvent('', subId, subTitle, 'UpdateEvent');
                  if (subQueued) {
                    await updateEvent(subId, { 'OutlookDirty': false }, { skipReload: true });
                  } else {
                    failedOutlookTitles.push(subTitle || (isDe ? 'Termin' : 'date'));
                  }
                } catch (err) { console.warn('[DEX][v30.67] Outlook-Update eines Sub-Events fehlgeschlagen:', subId, err); failedOutlookTitles.push(subTitle || (isDe ? 'Termin' : 'date')); }
              }
            }
          } catch { /* Sub-Outlook-Updates optional */ }
        }
        // v29.56: Hat der Organizer die Einladungs-Entscheidung umgestellt,
        // reicht das neue Flag NICHT — es steuert nur `requiredAttendees` beim
        // ANLEGEN. Bestehende Termine behalten ihre Teilnehmerliste. Also die
        // Organizer über die normale Queue an- bzw. abmelden, Event für Event.
        if (orgGetsSubInvites !== initialOrgGetsSubInvitesRef.current) {
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctx = (window as any).__dexSpfxContext;
            // Dieselbe Bereinigung wie beim Speichern (sanitizeOrganizerPairs),
            // damit hier nicht eine Rohadresse aus dem State landet.
            const orgMails = (sanitizeOrganizerPairs().orgEmailString || '')
              .split(';').map(x => x.trim()).filter(Boolean);
            if (ctx && orgMails.length > 0) {
              const svc = new EventService(ctx);
              const action: 'Einladen' | 'Ausladen' = orgGetsSubInvites ? 'Einladen' : 'Ausladen';
              // Klammer + alle gespeicherten Sub-Events. Ohne Outlook-Termin
              // (disableOutlook) ist der Eintrag wirkungslos, aber harmlos.
              const targets: Array<{ id: string; title: string }> = [
                { id: String(selectedEventId), title },
                ...subEventsRef.current
                  .filter(se => !!se.dbId && !se.disableOutlook)
                  .map(se => ({ id: se.dbId as string, title: se.title || '' })),
              ];
              for (let tgtIdx = 0; tgtIdx < targets.length; tgtIdx++) {
                const tgt = targets[tgtIdx];
                tickOutlook(isDe
                  ? `Organizer werden ${action === 'Einladen' ? 'eingeladen' : 'ausgeladen'}… (${tgtIdx + 1}/${targets.length}: ${tgt.title || 'Termin'})`
                  : `${action === 'Einladen' ? 'Inviting' : 'Removing'} organizers… (${tgtIdx + 1}/${targets.length}: ${tgt.title || 'date'})`);
                for (const mail of orgMails) {
                  try { await svc.queueOutlookEvent(mail, tgt.id, tgt.title, action); }
                  catch { /* einzelne Queue-Fehler nicht eskalieren */ }
                }
              }
              initialOrgGetsSubInvitesRef.current = orgGetsSubInvites;
            }
          } catch { /* Organizer-Nachzug ist best-effort */ }
        }
        // v11.63: Sub-Events, die im Modal waren aber NICHT angehakt wurden,
        // bekommen OutlookDirty=true, damit beim nächsten Wizard-Lauf der
        // Hinweis erscheint. Aus pendingOutlookDirtyWriteRefs lesen — Top-
        // Level haben wir oben bereits über pendingOutlookDirtyWriteRef
        // erledigt, hier nur Sub-Events.
        try {
          const dirtyMap = pendingOutlookDirtyWriteRefs.current || {};
          const checkedSubIds = new Set(pendingOutlookUpdateForSubEventsRef.current);
          for (const subId of Object.keys(dirtyMap)) {
            if (subId === selectedEventId) continue; // Top-Level schon erledigt
            if (checkedSubIds.has(subId)) continue;  // bereits auf false gesetzt
            if (dirtyMap[subId] === true) {
              try { await updateEvent(subId, { 'OutlookDirty': true }, { skipReload: true }); }
              catch { /* */ }
            }
          }
        } catch { /* */ }
        // v30.67: gescheiterte Outlook-Aufträge benennen statt still
        // „gespeichert" zu melden (s. failedOutlookTitles oben).
        if (failedOutlookTitles.length > 0) {
          showAlert(isDe
            ? `Das Event ist gespeichert, aber für ${failedOutlookTitles.length} Termin${failedOutlookTitles.length === 1 ? '' : 'e'} konnte die Outlook-Aktualisierung nicht angestoßen werden: ${failedOutlookTitles.join(', ')}. Die Kalender der Teilnehmer zeigen dort noch den alten Stand — bitte öffne das Event später erneut und bestätige die Aktualisierung noch einmal.`
            : `The event is saved, but the Outlook update could not be queued for ${failedOutlookTitles.length} date${failedOutlookTitles.length === 1 ? '' : 's'}: ${failedOutlookTitles.join(', ')}. Attendees' calendars still show the old state there — please reopen the event later and confirm the update again.`, { variant: 'error' });
        }
        setProgress(100);
        setProgressLabel('Änderungen gespeichert!');
        // v9.45: Soft-Refresh analog zum Create-Pfad. Wizard verlassen via
        // CustomEvent, DexEventPlatform navigiert + zeigt Banner. Refresh wird
        // beim Update sofort getriggert (loadEvents in updateEvent hat schon
        // gefeuert) — kein delayed refresh nötig wie beim Create.
        try {
          // v17.4: Nach erfolgreichem Save den Initial-Snapshot auf den
          // aktuellen Stand setzen, damit die Navigation-Guard (Unsaved-
          // Changes-Confirm) anschliessend nicht falsch auslöst. Sonst
          // sieht der User bei jedem Zurück-Klick nach Save das Modal,
          // obwohl alles persistiert ist.
          initialFormSnapshotRef.current = computeFormSnapshot();
          setNavigationGuard(null);
          // v17.21: Statt sofort den Wizard zu verlassen, öffnet sich erst
          // das Summary-Export-Modal. Der eigentliche Success-Dispatch läuft
          // erst, wenn der User dort eine Auswahl getroffen hat (PDF / Word /
          // Nein, danke).
          pendingSuccessDispatchRef.current = { title, eventId: String(selectedEventId), type: 'update' };
          setPendingSuccessDispatch({ title, eventId: String(selectedEventId), type: 'update' });
          setShowSummaryModal(true);
        } catch { /* */ }
        setIsSubmitting(false);
      } else {
        setIsSubmitting(false);
        setProgress(0);
        // v26.51: Grund IMMER mit anzeigen (vorher stand er nur in der Konsole).
        const reason = getLastEventUpdateError();
        const msg = (isDe ? 'Event konnte nicht aktualisiert werden.' : 'The event could not be updated.')
          + (reason
            ? `${isDe ? ' Grund: ' : ' Reason: '}${reason}`
            : (isDe
              ? ' Grund unbekannt — bitte Details aus der Browser-Konsole (F12) an das DEX-Team melden.'
              : ' Unknown reason — please report the details from the browser console (F12) to the DEX team.'));
        setError(msg);
        // v28.31: Die Fehlerzeile steht ganz oben im Wizard — beim Speichern aus
        // Schritt 6/7 war sie ausserhalb des Sichtfelds, der Klick sah aus, als
        // würde schlicht nichts passieren. Zusätzlich als Dialog zeigen.
        showAlert(msg, { variant: 'error' });
      }
    } else {
      // Neues Event erstellen — v11.87: Progress wird per Callback-Stage vom
      // EventService getrieben, damit der Balken sich tatsächlich an die
      // laufende SP-Operation koppelt und nicht stumm bei 92 % stehen bleibt.
      try {
      // Sub-Event-Anzahl bestimmt die Aufteilung des Bereichs 30 % - 90 %.
      // Bei N Sub-Events haben wir (1 Top + N Sub) Anlagen, die diesen
      // Bereich gleichmäßig füllen. Ohne Sub-Events bleibt das Hauptevent
      // den ganzen Bereich für sich.
      const subEventDraftsCount = subEventsRef.current.filter(d => d.title && d.title.trim()).length;
      const totalAnlagen = 1 + subEventDraftsCount;
      const topStart = 30;
      const topEnd = topStart + (90 - topStart) / totalAnlagen;
      // Innerhalb des Top-Event-Slots (topStart..topEnd) werden die Stages
      // verteilt: subsite-creating, permissions, list-creating, list-done,
      // item-insert, done.
      const reportCreateStage = (stage: string): void => {
        const slot = topEnd - topStart;
        switch (stage) {
          case 'start':
            setProgress(Math.round(topStart));
            setProgressLabel('Event-Daten gespeichert — Teilnehmer-Subsite wird angelegt...');
            break;
          case 'subsite-creating':
            setProgress(Math.round(topStart + slot * 0.05));
            setProgressLabel('Teilnehmer-Subsite wird angelegt...');
            break;
          case 'subsite-done':
            setProgress(Math.round(topStart + slot * 0.35));
            setProgressLabel('Subsite angelegt — Berechtigungen werden gesetzt...');
            break;
          case 'permissions':
            setProgress(Math.round(topStart + slot * 0.45));
            setProgressLabel('Berechtigungen werden gesetzt...');
            break;
          case 'list-creating':
            setProgress(Math.round(topStart + slot * 0.55));
            setProgressLabel('Teilnehmerliste wird angelegt...');
            break;
          case 'list-done':
            setProgress(Math.round(topStart + slot * 0.80));
            setProgressLabel('Teilnehmerliste fertig — Views werden konfiguriert...');
            break;
          case 'item-insert':
            setProgress(Math.round(topStart + slot * 0.90));
            setProgressLabel('Event-Daten werden gespeichert...');
            break;
          case 'done':
            setProgress(Math.round(topEnd));
            setProgressLabel(subEventDraftsCount > 0
              ? `Haupt-Event angelegt — Sub-Event 1 wird angelegt...`
              : 'Haupt-Event angelegt — Berechtigungen und Aufräumarbeiten...');
            break;
          default:
            break;
        }
      };

      setProgress(10);
      setProgressLabel('Event-Daten werden vorbereitet...');

      // v16.4: Audience-DLs beim Save in Member-E-Mails auflösen, damit der
      // Runtime-Sichtbarkeits-Check sie ohne weitere Graph-Calls treffen kann.
      const audienceResolved = await resolveAudienceMembersToCsv(audience, getGroupMembers);

      const sanitizedOrgPairCreate = sanitizeOrganizerPairs();
      const eventId = await createEvent({
        title,
        type: eventType,
        status: 'Active',
        description,
        location,
        locationAddress: (addrStreet || addrHouseNo || addrZip || addrCity)
          ? JSON.stringify({ street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity })
          : '',
        // v18.40: manueller Outlook-Ort (leer = Auto in createEvent).
        outlookLocation: outlookLocationOverride.trim() || undefined,
        outlookSubject: effOutlookSubject.trim() || undefined,
        outlookStart: outlookStartOverride || undefined,
        outlookEnd: outlookEndOverride || undefined,
        allDay, // v29.52
        showAsFree, // v29.54
        outlookIsOnlineMeeting: onlineMeetingMode === 'auto', // v30.26
        skipOrganizerInvite: !orgGetsSubInvites, // v29.55
        locationFilter,
        audience,
        audienceResolvedEmails: audienceResolved,
        filterMode,
        startDate: startDate ? berlinLocalToUtcIso(startDate) : '',
        // v22.17: EndDate nie leer lassen (Outlook-Flow-Crash, s.o.) — Fallback Start.
        endDate: endDate ? berlinLocalToUtcIso(endDate) : (startDate ? berlinLocalToUtcIso(startDate) : ''),
        registrationDeadline: deadlineToEndOfDayIso(registrationDeadline) || '',
        // v29.25: Ohne Selbst-Abmeldung keine Abmeldefrist (s. Edit-Pfad).
        lastDeregisterDate: userCancelAllowed ? (deadlineToEndOfDayIso(lastDeregisterDate) || '') : '',
        // v29.19: „Aktiv ab" auch beim ANLEGEN persistieren — gleiche
        // Konvertierung wie der Edit-Pfad. Vorher wurde nur das abhängige
        // _previewBeforeActive-Flag geschrieben, das Datum selbst nicht.
        activeFrom: activeFrom ? (berlinLocalToUtcIso(activeFrom) || undefined) : undefined,
        // v29.21: Split-Invariante (s. Edit-Pfad).
        maxParticipants: useSplitCapacities ? 0 : (unlimitedParticipants ? 0 : (Number(maxParticipants) || 0)),
        waitlistEnabled,
        eventImageUrl: imageUrl,
        // Sanitize: paart Organizer-Names + -Emails 1:1, droppt unvollständige
        // Pairs (Name ohne Email oder umgekehrt) — verhindert Mismatch-State
        // in DEX_Events durch Drift während Edit/Closure-Bugs.
        organizer: sanitizedOrgPairCreate.orgString,
        organizerEmail: sanitizedOrgPairCreate.orgEmailString,
        contactName: contactName.trim(),
        contactEmail: contactEmail.trim(),
        contactOrganizerEmail: (contactOrganizerEmail || '').trim(),
        contactInfo: contactInfo.trim(),
        outlookEventId: '',
        outlookBody: (() => {
          // v7.4: Auch wenn der User keinen Outlook-Body geschrieben hat,
          // immer das Outlook-Layout mit einem Default-Body erzeugen,
          // der auf den Organizer für organisatorische Fragen verweist
          // (analog zur Anmeldebestätigungs-Mail).
          // v24.60: Namen wie die Anmelde-Mail normalisieren + mit „und"/„and"
          // verbinden (nicht stumpf Komma-getrennt).
          const orgNames = formatOrganizerList([organizer], effEmailLanguage);
          const vars = {
            EventTitle: title,
            // v27.5: {{Organizer}} = normalisierte Namen statt roher Join.
            Organizer: orgNames || organizer,
            ContactEmail: contactEmail.trim(),
            Location: location,
            Address: [addrStreet, addrHouseNo].filter(Boolean).join(' ') + ((addrZip || addrCity) ? ', ' + [addrZip, addrCity].filter(Boolean).join(' ') : ''),
            StartDate: startDate ? new Date(startDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
            EndDate: endDate ? new Date(endDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
          };
          const escHtml = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
          // v9.8: gleicher Default-Body wie im Update-Pfad — inkl. Abmelde-Hinweis
          // mit Link auf die App ("Meine Events"-Tab).
          const APP_URL_OL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';
          const defaultBody = effEmailLanguage === 'EN'
            ? `<p>You are registered for the event <strong>${escHtml(title)}</strong>.</p>`
              + `<p>If you are unable to attend, please cancel your registration in time via the <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;My Events&ldquo;).</p>`
              + `<p>For organizational questions please contact <strong>${escHtml(orgNames || 'the organizer')}</strong>.</p>`
            : `<p>Ihr seid für das Event <strong>${escHtml(title)}</strong> angemeldet.</p>`
              + `<p>Falls ihr nicht teilnehmen könnt, meldet euch bitte rechtzeitig über die <a href="${APP_URL_OL}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;Meine Events&ldquo;) ab.</p>`
              + `<p>Bei organisatorischen Fragen wendet euch bitte an <strong>${escHtml(orgNames || 'den Organizer')}</strong>.</p>`;
          const resolvedBody = effOutlookBody ? replacePlaceholders(effOutlookBody, vars) : defaultBody;
          const resolvedHeading = effOutlookHeading ? replacePlaceholders(effOutlookHeading, vars) : title;
          // v27.5: Default-Unter-Überschrift = Ort (nicht Datum).
          const resolvedSub = effOutlookSubheading ? replacePlaceholders(effOutlookSubheading, vars) : (location || undefined);
          // v18.73: Header-Bild Größe + Innenabstand (event-weit) in den Outlook-Body.
          // v30.67: effEmailLanguage — s. Update-Pfad.
          const wrapped = buildOutlookBody(resolvedHeading, resolvedBody, resolvedSub, headerLayoutFor(effOutlookLogo), outlookTeamsLink(), (effEmailLanguage || '').toUpperCase() !== 'EN');
          // v11.93: Logo aus Top-Level-Resolver, sonst landet beim Speichern
          // aus einem Sub-Tab das Sub-Logo aufs Haupt-Event.
          return wrapped.replace(/\{\{ORB_URL\}\}/g, effOutlookLogo || getCachedOrbBase64() || '');
        })(),
        agenda: JSON.stringify(agenda),
        transfers: JSON.stringify(transferTimes),
        documents: '[]', // Dokumente werden nach erfolgreichem Upload gespeichert
        funZone: JSON.stringify(quiz),
        quizClusterSize: Math.min(Math.max(1, quizClusterSize || 1), 4),
        emailLanguage: effEmailLanguage,
        registrationLanguage: registrationLanguage || undefined,
        emailTemplateOverrides: (() => {
          const b2runExtra = (durchstarterStartblock || funstarterStartblock || durchstarterRequiresProof)
            ? { _b2run: {
                ...(durchstarterStartblock ? { durchstarterStartblock } : {}),
                ...(funstarterStartblock ? { funstarterStartblock } : {}),
                ...(durchstarterRequiresProof ? { durchstarterRequiresProof: true } : {}),
              } }
            : {};
          const qrExtra = qrScannerEmails.length > 0
            ? { _qrScanners: qrScannerNames.map((n, i) => ({ name: n, email: qrScannerEmails[i] || '' })).filter(x => x.email) }
            : {};
          // v9.18: Co-Organizer ebenso in EmailTemplateOverrides piggybacken.
          const coExtra = coOrganizerEmails.length > 0
            ? { _coOrganizers: coOrganizerNames.map((n, i) => ({ name: n, email: coOrganizerEmails[i] || '' })).filter(x => x.email) }
            : {};
          // v9.21: Test-Team ebenso piggybacken.
          const ttExtra = testTeamEmails.length > 0
            ? { _testTeam: testTeamNames.map((n, i) => ({ name: n, email: testTeamEmails[i] || '' })).filter(x => x.email) }
            : {};
          // v11.25: Display-Reihenfolge-Toggle (siehe Edit-Pfad oben).
          const splitDispRevExtra = splitDisplayOrderReversed && useSplitCapacities
            ? { _splitDisplayOrderReversed: true }
            : {};
          // v14.8: subEventsOnlyMode + custom childEventTerm piggybacken;
          // subEventsOnlyMode impliziert requireSubEventSelection.
          const effRequireSubEventSelection = requireSubEventSelection || subEventsOnlyMode;
          const reqSubEvtExtra = effRequireSubEventSelection
            ? { _requireSubEventSelection: true }
            : {};
          const subEvtsOnlyExtra = subEventsOnlyMode
            ? { _subEventsOnlyMode: true }
            : {};
          // v28.2: Soft-Disable-Flag (s. Edit-Pfad).
          const subEvtsDisabledExtra = (!subEventsOptIn && subEventsRef.current.some(s => s.title && s.title.trim()))
            ? { _subEventsDisabled: true }
            : {};
          // v28.5: Bild-Banner-Layout (Piggyback).
          const imageBannerExtra = imageBanner ? { _imageBanner: true } : {};
          // v28.20: explizite Klammer-Frist (s. Edit-Pfad).
          // v30.6: bei aktiver rollierender Regel nie mitschreiben (s. Edit-Pfad).
          const klammerDeadlineExtra: Record<string, unknown> = (() => {
            if (subEventsOptIn && regRuleEnabled) return {};
            const iso = (subEventsOnlyMode && klammerDeadline) ? deadlineToEndOfDayIso(klammerDeadline) : null;
            return iso ? { _klammerDeadline: iso } : {};
          })();
          const childTermExtra = (childTermSingular.trim() || childTermPlural.trim())
            ? { _childEventTerm: { singular: childTermSingular.trim(), plural: childTermPlural.trim(), ...(childGender ? { gender: childGender } : {}) } }
            : {};
          // v18.9: Organizer-Anzeige ausblenden (Piggyback).
          const hideOrganizerExtra = hideOrganizer ? { _hideOrganizer: true } : {};
          const hiddenOrganizersExtra: Record<string, unknown> = hiddenOrganizerEmails.length > 0 ? { _hiddenOrganizers: hiddenOrganizerEmails } : {};
          const hideOrgIndividualExtra: Record<string, unknown> = hideOrganizerIndividualOnly ? { _hideOrgIndividual: true } : {};
          // v22.78: Team-Begriff + „keine neuen Teams"-Flag (Piggyback).
          const teamTermExtra = (teamTermSingular.trim() || teamTermPlural.trim())
            ? { _teamTerm: { singular: teamTermSingular.trim(), plural: teamTermPlural.trim() } }
            : {};
          const teamNoCreateExtra = teamMembersCannotCreate ? { _teamMembersCannotCreate: true } : {};
          // v24.58: Anzeige-Bezeichnung des Haupt-Events (Piggyback).
          const mainEventLabelExtra: Record<string, unknown> =
            mainEventLabelMode === 'none' ? { _mainEventLabel: { mode: 'none' } }
            : (mainEventLabelMode === 'custom' && mainEventLabel.trim()) ? { _mainEventLabel: { mode: 'custom', text: mainEventLabel.trim() } }
            : {};
          // v23.6: Assistenz-Sichtbarkeit (Piggyback).
          const assistantsCanSeeExtra = assistantsCanSee ? { _assistantsCanSee: true } : {};
          // v23.25: Organizer groß darstellen (Piggyback).
          const organizerDisplayLargeExtra = organizerDisplayLarge ? { _organizerDisplayLarge: true } : {};
          // v23.14: Vorschau vor Aktivierung (nur sinnvoll mit activeFrom).
          const previewBeforeActiveExtra = (previewBeforeActive && activeFrom) ? { _previewBeforeActive: true } : {};
          // v23.19: Pro-Ansicht-Bilddarstellung (nur abweichende Views).
          const imageDisplayExtra = (() => {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const out: any = {};
            (['card', 'hero'] as const).forEach(k => {
              const v = imageDisplay[k];
              if (!v) return;
              const hSet = typeof v.height === 'number' && v.height !== 340;
              if (v.zoom !== 1 || v.posY !== 50 || hSet) out[k] = { zoom: v.zoom, posY: v.posY, ...(hSet ? { height: v.height } : {}) };
            });
            return Object.keys(out).length ? { _imageDisplay: out } : {};
          })();
          // v11.93: Top-Level-Logos aus dem Resolver lesen.
          // v28.2: s. Edit-Pfad — Array + .some() statt ||-Kette (TS2590).
          const createPiggybackConfigs: Array<Record<string, unknown>> = [
            b2runExtra, qrExtra, coExtra, ttExtra, splitDispRevExtra,
            reqSubEvtExtra, subEvtsOnlyExtra, subEvtsDisabledExtra,
            imageBannerExtra, klammerDeadlineExtra, childTermExtra, teamTermExtra, teamNoCreateExtra,
            mainEventLabelExtra, assistantsCanSeeExtra,
            organizerDisplayLargeExtra, previewBeforeActiveExtra,
            imageDisplayExtra, hideOrganizerExtra, hiddenOrganizersExtra,
            hideOrgIndividualExtra, headerImageLayoutConfig,
            // v28.79: „Keine Beschreibung nutzen" auch beim Anlegen merken.
            ((noDescription && !description.trim()) ? { _noDescription: true } : {}),
            // v28.91: Kalender-Modus der Sub-Events.
            ((subEventCalendar && subEventsOptIn) ? { _subEventCalendar: true } : {}),
            ((subEventSingleChoice && subEventsOptIn) ? { _subEventSingleChoice: true } : {}),
            (subEventsOptIn ? bundledCommConfig(bundledComm) : {}),
            // v29.25: Abmelde-Sperren auch beim Anlegen.
            (!userCancelAllowed ? { _noSelfCancel: true } : {}),
            billingPiggyback(), // v29.66: F&A-Pilot
            subEventOpenRulePiggyback(), // v29.67
            visAllSubsPiggyback(), // v29.75
            subDeadlineRulePiggyback(), // v29.76
            ((userCancelAllowed && noCancelAfterDeadline) ? { _noCancelAfterDeadline: true } : {}),
            // v29.38: Teams-Link auch beim Anlegen.
            (/^https?:\/\//i.test(effTeamsLink()) ? { _teamsLink: effTeamsLink() } : {}),
          ];
          // v29.19: Overrides aus dem Top-Level-Resolver — wie im Edit-Pfad
          // (v14.4). Der rohe State hält beim Speichern von einem Sub-Reiter
          // aus die SUB-Werte (switchCommTab-Spiegelung); die landeten hier
          // als letzter Merge auf dem Hauptevent, während die zuvor gesetzten
          // Hauptevent-Overrides im Snapshot verloren gingen. Alle übrigen
          // Kommunikationsfelder gingen längst über topComm — nur diese nicht.
          const topOverridesCreate = topComm.emailTemplateOverrides || {};
          const hasAny = Object.keys(topOverridesCreate).length > 0 || !!effEmailLogo || !!effOutlookLogo || createPiggybackConfigs.some(o => Object.keys(o).length > 0);
          return hasAny
            // v28.2: Object.assign statt Spread-Kette (TS2590, s. Edit-Pfad).
            // Keys disjunkt; topOverridesCreate bleibt LETZTER Merge.
            ? JSON.stringify(Object.assign(
                {},
                (effEmailLogo ? { _eventLogo: effEmailLogo } : {}),
                outlookLogoPiggyback(effEmailLogo, effOutlookLogo),
                ...createPiggybackConfigs,
                topOverridesCreate,
              ))
            : '';
        })(),
        // v11.93: aus dem Top-Level-Resolver — Sub-Tab-Werte würden sonst
        // beim Save fälschlich aufs Haupt-Event übernommen.
        disableEmails: effDisableEmails,
        // v19.22: granulare An-/Abmelde-Mail-Schalter (Top-Level aufgelöst).
        disableRegistrationEmail: effDisableRegistrationEmail,
        disableCancellationEmail: effDisableCancellationEmail,
        // v19.23/v19.24: Outlook-Absage = Auto-Abmeldung (Top-Level aufgelöst).
        autoDeregisterOnDecline: effAutoDeregisterOnDecline,
        inactiveHandling: effInactiveHandling,
        disableOutlook: effDisableOutlook,
        notifyOrgRegisterMode,
        notifyOrgRegisterFromDate: notifyOrgRegisterMode === 'fromDate' && notifyOrgRegisterFromDate ? berlinLocalToUtcIso(notifyOrgRegisterFromDate) : '',
        notifyOrgCancelMode,
        excludedUsers,
        isFictive,
        durchstarterCapacity: useSplitCapacities ? (parseInt(durchstarterCapacity, 10) || 0) : undefined,
        funstarterCapacity: useSplitCapacities ? (parseInt(funstarterCapacity, 10) || 0) : undefined,
        splitLabelA: useSplitCapacities ? (splitLabelA || '').trim() : undefined,
        splitLabelB: useSplitCapacities ? (splitLabelB || '').trim() : undefined,
        splitDescA: useSplitCapacities ? (splitDescA || '').trim() : undefined,
        splitDescB: useSplitCapacities ? (splitDescB || '').trim() : undefined,
        splitHelpText: useSplitCapacities ? (splitHelpText || '').trim() : undefined,
        splitSectionTitle: useSplitCapacities ? (splitSectionTitle || '').trim() : undefined,
        splitSharedWaitlist: useSplitCapacities ? !!splitSharedWaitlist : undefined,
        allowAttendeeUpload: !!allowAttendeeUpload,
        attendeeUploadHint: (attendeeUploadHint || '').trim() || undefined,
        attendeeUploadLabel: (attendeeUploadLabel || '').trim() || undefined,
        // v18.33: Self-Check-in mit-durchreichen (Token + optionales Fenster).
        // v20.2: SelfCheckIn*-Felder beim Create nicht mehr gesetzt — neue
        // Events starten ohne Token; die Auto-Aktivierung beim ersten Klick
        // (Check-in-Seite / Admin Center / QR-Kachel) übernimmt das.
        // v11.80: Anrede-Toggle + Team-Anmelde-Konfiguration mit-durchreichen.
        askSalutation: !!askSalutation,
        // v18.75: Sicherheitshinweis vor dem Absenden.
        confirmDialogEnabled: !!confirmDialogEnabled,
        confirmDialogMode: confirmDialogEnabled ? (confirmDialogMode || 'summary') : '',
        confirmDialogText: confirmDialogEnabled && confirmDialogMode === 'freetext' ? confirmDialogText : '',
        teamRegistrationEnabled: !!teamRegistrationEnabled,
        teamSize: teamRegistrationEnabled && teamSize > 0 ? teamSize : undefined,
        askTeamName: !!askTeamName,
        // v11.81: Erweiterte Team-Konfiguration mit-durchreichen.
        teamPartialAllowed: !!(teamRegistrationEnabled && teamPartialAllowed),
        teamOpenSlotsVisible: !!(teamRegistrationEnabled && teamOpenSlotsVisible),
        teamJoinRequiresApproval: !!(teamRegistrationEnabled && teamOpenSlotsVisible && teamJoinRequiresApproval),
        // v17.20: Bilingual-Toggle durchreichen.
        bilingualFields: !!bilingualFields,
        // v17.22: zentraler serializeCustomFields-Helper.
        customFields: serializeCustomFields(customFields, bilingualFields),
        onProgress: reportCreateStage,
      });

      if (eventId) {
        // v26.24: Co-Organizer-Freigabe (siehe Edit-Pfad) — für benannte
        // Organizer, die noch kein Organizer/Admin sind, einen „Organizer
        // werden"-Antrag zur Admin-Freigabe anlegen. Best-effort.
        try {
          await requestCoOrganizerApprovals(sanitizedOrgPairCreate.orgString, sanitizedOrgPairCreate.orgEmailString, title);
        } catch (err) { console.warn('[DEX] Co-Organizer-Freigabe-Anträge (Create) fehlgeschlagen:', err); }

        // v26.57: Zielgruppen-Personen außerhalb von @deloitte.de → Approve-Mail
        // an die Admins für den Site-Zugriff (SharePoint-Default: nur Deloitte
        // DE ALL). Beim Neu-Anlegen zählt jede Nicht-DE-Adresse. Fire-and-forget.
        try {
          const nonDe = audience.split(',')
            .map(s => s.trim())
            .filter(a => a.indexOf('@') > 0 && !a.toLowerCase().endsWith('@deloitte.de'));
          if (nonDe.length > 0) {
            void notifyAdminsExternalAudienceAccess(title, nonDe, `${currentUser.firstName} ${currentUser.surname}`.trim()).catch(() => { /* */ });
          }
        } catch { /* darf den Save nie stören */ }

        // v29.17: Bild und Dokumente SOFORT nach dem Anlegen des Hauptevents
        // hochladen — VOR den Sub-Events. Bisher hingen beide am Ende des
        // Pfads, hinter persistSubEventsForParent und einem getEvents() im
        // selben try-Block (dessen catch alles schluckte). Bei einem
        // Kalender-Event mit 20+ Terminen legt persistSubEventsForParent
        // ebenso viele Subsites an; wenn SharePoint danach drosselt, flog
        // getEvents() — und das Klammer-Bild wurde still nie hochgeladen.
        // Das Event-Item existiert hier bereits, mehr braucht der Upload nicht.
        if (imageFile || documents.length > 0) {
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctxEarly = (window as any).__dexSpfxContext;
            if (ctxEarly) {
              const svcEarly = new EventService(ctxEarly);
              if (imageFile) {
                try {
                  setProgressLabel('Bild wird hochgeladen...');
                  const compressed = await compressImage(imageFile);
                  const uploadedUrl = await svcEarly.uploadEventImageAsAttachment(Number(eventId), compressed);
                  if (uploadedUrl) {
                    await svcEarly.updateEventImageUrl(Number(eventId), uploadedUrl);
                    // v28.11: Querformat-Original als zweites Attachment sichern
                    // (Zuschnitt rund/quadratisch → Event-Liste zeigt Original).
                    try {
                      if (imageOrigFile && (imageOrigAspect || 0) >= 1.2 && wizardImgAspect != null && wizardImgAspect < 1.2) {
                        const origCompressed = await compressImage(imageOrigFile, 1600, 0.85);
                        const origUrl = await svcEarly.uploadEventOrigImageAsAttachment(Number(eventId), origCompressed);
                        if (origUrl) await svcEarly.patchEventOverridesKey(Number(eventId), '_imageOrigUrl', origUrl);
                      }
                    } catch (origErr) { console.warn('[DEX] Original-Bild speichern fehlgeschlagen:', origErr); }
                  } else {
                    setImageUploadError('Bild-Upload fehlgeschlagen.');
                    // v29.19: Sichtbar melden — die rote Zeile in Schritt 1
                    // ist vom Fortschritts-Fenster verdeckt, und nach dem
                    // Success-Dispatch verlaesst der Wizard die Seite. Ohne
                    // Alert ging das Event still ohne Bild live (Paritaet zum
                    // Edit-Pfad, v29.18).
                    showAlert(isDe
                      ? 'Das Event-Bild konnte nicht hochgeladen werden. Das Event wird trotzdem angelegt — bitte lade das Bild danach über Bearbeiten in Schritt 1 erneut hoch.'
                      : 'The event image could not be uploaded. The event is still being created — please upload the image again via Edit in step 1 afterwards.', { variant: 'error' });
                  }
                } catch (err) {
                  console.warn('[DEX] Bild-Upload fehlgeschlagen', err);
                  setImageUploadError('Bild-Upload fehlgeschlagen.');
                  showAlert(isDe
                    ? 'Das Event-Bild konnte nicht hochgeladen werden. Das Event wird trotzdem angelegt — bitte lade das Bild danach über Bearbeiten in Schritt 1 erneut hoch.'
                    : 'The event image could not be uploaded. The event is still being created — please upload the image again via Edit in step 1 afterwards.', { variant: 'error' });
                }
              }
              // Dokumente einzeln best-effort — ein defektes Dokument darf
              // weder die übrigen noch den Rest des Saves mitreißen.
              for (const doc of documents) {
                if (!doc.file) continue;
                try { await svcEarly.uploadEventDocument(Number(eventId), doc.file); }
                catch (err) { console.warn('[DEX] Dokument-Upload fehlgeschlagen:', doc.file.name, err); }
              }
            }
          } catch (err) { console.warn('[DEX] Upload-Block (früh) fehlgeschlagen:', err); }
        }

        // v11.87: Sub-Events bekommen den Bereich (topEnd..90) gleichmäßig
        // aufgeteilt — pro Sub-Event ein eigener Stage-Slot. persistSubEventsForParent
        // erhält einen Sub-Progress-Callback über ein Window-Event-Bus-ähnliches
        // Setup ist hier nicht nötig, weil wir die Schleife per index zählen.
        if (subEventDraftsCount > 0) {
          const subSlotSize = (90 - topEnd) / subEventDraftsCount;
          // Wir setzen pro Sub-Event-Start manuell den Progress und übergeben
          // optional einen onProgress-Callback an persistSubEventsForParent, um
          // den Sub-Site-Anlage-Fortschritt fein abzubilden. Da
          // persistSubEventsForParent intern über sequenzielle createEvent
          // läuft, koppeln wir den Sub-Progress an einen externen Counter.
          let processedSubIdx = 0;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__dexSubEventProgress = (stage: string): void => {
            const base = topEnd + subSlotSize * processedSubIdx;
            const slot = subSlotSize;
            switch (stage) {
              case 'start':
                setProgress(Math.round(base));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} wird vorbereitet...`);
                break;
              case 'subsite-creating':
                setProgress(Math.round(base + slot * 0.10));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} — Subsite wird angelegt...`);
                break;
              case 'subsite-done':
                setProgress(Math.round(base + slot * 0.45));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} — Berechtigungen werden gesetzt...`);
                break;
              case 'list-creating':
                setProgress(Math.round(base + slot * 0.60));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} — Teilnehmerliste wird angelegt...`);
                break;
              case 'list-done':
                setProgress(Math.round(base + slot * 0.80));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} — Views werden konfiguriert...`);
                break;
              case 'item-insert':
                setProgress(Math.round(base + slot * 0.90));
                setProgressLabel(`Sub-Event ${processedSubIdx + 1} von ${subEventDraftsCount} — Event-Daten werden gespeichert...`);
                break;
              case 'done':
                processedSubIdx += 1;
                setProgress(Math.round(topEnd + subSlotSize * processedSubIdx));
                setProgressLabel(processedSubIdx >= subEventDraftsCount
                  ? 'Sub-Events fertig — Aufräumarbeiten...'
                  : `Sub-Event ${processedSubIdx} fertig — Sub-Event ${processedSubIdx + 1} wird angelegt...`);
                break;
              default:
                break;
            }
          };
          setProgress(Math.round(topEnd));
          setProgressLabel(`Sub-Event 1 von ${subEventDraftsCount} wird angelegt...`);
        } else {
          setProgress(Math.round(topEnd));
          setProgressLabel('Haupt-Event angelegt — Aufräumarbeiten...');
        }
        try { await persistSubEventsForParent(String(eventId)); }
        catch (err) { console.warn('[DEX] Sub-Events beim Create persistieren fehlgeschlagen:', err); }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        try { delete (window as any).__dexSubEventProgress; } catch { /* */ }
        setProgress(92);
        setProgressLabel('Letzte Schritte...');
        // E-Mail an Organisator senden. Bild und Dokumente sind seit v29.17
        // bereits VOR den Sub-Events hochgeladen (s.o.) — hier ist nur noch
        // die Mail übrig, und getEvents() dient allein der Subsite-URL darin.
        // (Zum refreshEvents gilt weiter v9.41: KEIN Refresh direkt nach
        // Create — die frische Subsite ist noch nicht API-konsistent.)
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ctx = (window as any).__dexSpfxContext;
          if (ctx) {
            const svc = new EventService(ctx);
            // Subsite-URL aus dem neu geladenen Event holen
            const allEvents = await svc.getEvents();
            const created = allEvents.find(e => String(e.Id) === String(eventId));
            const subsiteUrl = created?.SubsiteUrl || '';
            // Event-Created Mail an alle Organizer senden.
            // {{Name}} in der Anrede = nur Vorname (nicht voller Name), darum
            // den Organizer-String anhand von ";" in Namen splitten und pro
            // Name das erste Token als Vorname nehmen. Paarweise zu den
            // organizerEmails - bei Längen-Mismatch fällt wir auf den
            // ersten Namen zurück.
            const allOrgEmailsRaw = organizerEmails.length > 0 ? organizerEmails : [currentUser.email];
            // v28.71: pro Adresse nur EINE Mail. Steht dieselbe Person zweimal
            // in der Organizer-Liste (z.B. durch einen Doppel-Eintrag), kam die
            // „Event angelegt"-Mail bisher entsprechend oft an.
            const seenOrgMails = new Set<string>();
            const allOrgEmails = allOrgEmailsRaw.filter(e => {
              const lc = (e || '').trim().toLowerCase();
              if (!lc || seenOrgMails.has(lc)) return false;
              seenOrgMails.add(lc);
              return true;
            });
            const orgNames = organizer.split(';').map(s => s.trim()).filter(Boolean);
            for (let i = 0; i < allOrgEmails.length; i++) {
              const orgEmail = allOrgEmails[i];
              const orgFullName = orgNames[i] || orgNames[0] || `${currentUser.firstName} ${currentUser.surname}`;
              const orgFirstName = orgFullName.split(/\s+/)[0] || orgFullName;
              // v29.32: Kopfbild-Layout des Events mitgeben — sonst kam die
              // „Event angelegt"-Mail immer mit dem kleinen zentrierten Bild,
              // auch bei Events mit Vollbild-Kopf (seit v29.29 der Default).
              // headerLayoutFor deckelt ohne eigenes Bild weiterhin auf 180 px.
              const createdMailLayout = headerLayoutFor(effEmailLogo);
              const emailData = eventCreatedEmail(orgFirstName, title, subsiteUrl, {
                imageWidth: createdMailLayout.imageWidth,
                imagePaddingV: createdMailLayout.imagePaddingV,
                imagePaddingH: createdMailLayout.imagePaddingH,
              });
              svc.queueEmail(
                emailData.subject, orgEmail, orgFullName, emailData.body,
                'EventErstellt', title, String(eventId)
              ).catch(err => console.warn('[DEX]', err));
            }
          }
        } catch { /* E-Mail-Fehler ignorieren */ }
        // v9.45: Soft-Refresh statt Hard-Reload. Statt die Success-Page zu rendern
        // (wo zwischen Wizard und SuccessPage ein React #300 auftrat) ODER die
        // Page hart zu reloaden (was den User auf der Landing-Seite landete),
        // gehen wir den Mittelweg:
        //
        // 1. Wizard sofort verlassen via dispatchEvent('dex-event-submit-success',
        //    {title, eventId, type}) — DexEventPlatform hört darauf, navigiert zur
        //    Event-Liste und zeigt den grünen Erfolgs-Banner.
        // 2. setIsSubmitting(false) damit der Wizard unmounted (kein hängender
        //    Submit-State).
        // 3. Delayed refreshEvents (3 Sekunden später) lädt das frisch erstellte
        //    Event nach — SP hatte dann genug Zeit zum Propagieren und der Read
        //    auf die neue Subsite läuft sauber durch (gleicher Pfad wie der
        //    Aktualisieren-Button im Header — der hat nie Probleme).
        setProgress(100);
        setProgressLabel('Event erfolgreich erstellt!');
        try {
          // v17.4: gleicher Reset wie im Update-Pfad, damit der
          // Navigation-Guard nach erfolgreichem Create nicht stört.
          initialFormSnapshotRef.current = computeFormSnapshot();
          setNavigationGuard(null);
          // v17.21: Summary-Export-Modal vor dem Submit-Success-Dispatch.
          // v28.71: merken, dass dieser Wizard-Durchlauf sein Event bereits
          // angelegt hat — jeder weitere „Speichern"-Klick darf kein zweites
          // Event mehr erzeugen (s. handleSubmit).
          createdEventIdRef.current = String(eventId);
          // v30.4: Der Entwurf ist mit dem Anlegen erledigt — HIER löschen.
          // Der bisherige Aufräum-Effect hing an `submitted`, das der
          // Create-Pfad nie setzt (er verlässt den Wizard über den
          // Success-Dispatch) — deshalb tauchte der Entwurf nach dem
          // Erstellen wieder auf. Zusätzlich draftSavedAt/pendingDraft
          // zurücksetzen, damit weder Anzeige noch Autosave ihn wiederbeleben.
          try { localStorage.removeItem(DRAFT_KEY); } catch { /* */ }
          lastDraftJsonRef.current = '';
          setDraftSavedAt(null);
          setPendingDraft(null);
          pendingSuccessDispatchRef.current = { title, eventId: String(eventId), type: 'create' };
          setPendingSuccessDispatch({ title, eventId: String(eventId), type: 'create' });
          setShowSummaryModal(true);
        } catch { /* */ }
        setIsSubmitting(false);
        // Delayed Refresh — SP braucht typischerweise 2-5s bis frische Subsite-
        // Listen API-konsistent abrufbar sind. 3000ms ist ein guter Kompromiss.
        setTimeout(() => {
          refreshEvents().catch(err => console.warn('[DEX] post-create soft refresh fehlgeschlagen:', err));
        }, 3000);
      } else {
        setIsSubmitting(false);
        setProgress(0);
        setError('Event konnte nicht erstellt werden. Bitte versuche es erneut.');
      }
      } catch (err) {
        setIsSubmitting(false);
        setProgress(0);
        setError(err instanceof Error ? err.message : 'Event konnte nicht erstellt werden.');
      }
    }
}
