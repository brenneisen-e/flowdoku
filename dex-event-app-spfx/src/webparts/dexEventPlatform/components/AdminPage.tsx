/**
 * Admin / Organizer Seite
 *
 * Zeigt alle Events des Admins. Nach Auswahl eines Events:
 * - Event bearbeiten (Daten ändern)
 * - Teilnehmerliste anzeigen
 * - Neues Event erstellen
 * (v27.13: „Teilnehmerliste in SharePoint öffnen" entfernt — alle Aktionen
 *  laufen über die App.)
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { DeloitteEvent } from '../types';
import { SPRegistration } from '../services/EventService';
import { Users } from './Icons';
import B2RunBibImportModal from './admin/B2RunBibImportModal';
import B2RunTodoModal from './admin/B2RunTodoModal';
import ShirtSizeModal from './admin/ShirtSizeModal';
import { SHIRT_PATTERN } from '../utils/checkInExtras';
import { isEventOver } from '../utils/eventFormat';
import AddParticipantsModal from './admin/AddParticipantsModal';
import { accountCheckCacheKey } from '../utils/accountCheckCache';
// v20.1: Self-Check-in jederzeit aktivierbar (Token-Erzeugung beim Klick).
// v20.2: + statische Check-in-URL für die QR-Kachel im Event-Detail.
// v20.3: + Default-Zeitfenster (2 Std. vor Start bis Event-Ende) zur Vorbelegung.
import { useIsMobile } from '../utils/useIsMobile';
// v20.0 (Audit): xlsx + qrcode werden nicht mehr statisch importiert, sondern
// erst beim tatsächlichen Gebrauch (Export-Klick / QR-Vorschau) als eigener
// Chunk nachgeladen — spart ~1 MB im Haupt-Bundle.
import { EventService, EventCommRow } from '../services/EventService';
import { SharePointService } from '../services/SharePointService';
// v26.47: Externe Anmeldung — Einladung als .eml-Entwurf (X-Unsent) zum
// Selbst-Versenden durch die anmeldende Person (App kann keine externen
// Adressen anmailen).
import ImageCropModal from './ImageCropModal';
// v28.95: Platzhalter für Events ohne eigenes Foto. Zuerst das im Admin
// Center unter „Logo & Branding" hinterlegte DEX-Orb (DefaultImageBase64 im
// _Config-Eintrag von DEX_EmailTemplates) — das ist die Stelle, an der es
// ausgetauscht wird, und dann soll der Tausch überall greifen. Das
// gebuendelte PNG ist nur der Rueckfall, solange der Cache noch nicht
// geladen ist (frischer Tab, erster Render).
import TicketEventBox from './tickets/TicketEventBox';
// v20.4: moderne Confirm-/Alert-Modals statt window.confirm/alert.
import { useDialog } from '../context/DialogContext';
import { shortSubEventTitle } from '../utils/subEventTitle';
import { ActionsRegistryProvider } from './admin/ActionsMenu';
import BillingActionPanel from './admin/BillingActionPanel';
import { parseBillingOf } from '../utils/faBilling';
import { MailHeaderImage, MAIL_HEADER_IMAGE_DEFAULT } from '../utils/mailHeaderImage';










// v30.66: Modul-Ebene (Konstanten, Styles, Typen, Kopfbild-Helfer) liegt jetzt
// in components/admin/adminConstants|adminStyles|adminTypes — sie kennt den
// State nicht und war restlos abtrennbar.
import { DUP_ACTIVE_STATI, ACCESS_DENIED_MSG } from './admin/adminConstants';
import { ConsolidatedRow, AudiencePerson, MassmailAudience, AdminToastState, DeniedSubEventList } from './admin/adminTypes';
import JumpButtons from './admin/JumpButtons';
import { NameFixModal } from './admin/modals/NameFixModal';
import { AccessFixModal } from './admin/modals/AccessFixModal';
import { SelfCheckInModal } from './admin/modals/SelfCheckInModal';
import { CheckInHubModal } from './admin/modals/CheckInHubModal';
import { QrSendModal } from './admin/modals/QrSendModal';
import { QrEditModal } from './admin/modals/QrEditModal';
import { EditRegModal } from './admin/modals/EditRegModal';
import { ParticipantDetailModal } from './admin/modals/ParticipantDetailModal';
import { AssignAssistModal } from './admin/modals/AssignAssistModal';
import { MainFieldsEditModal } from './admin/modals/MainFieldsEditModal';
import { DeregModal } from './admin/modals/DeregModal';
import { QrPreviewModal } from './admin/modals/QrPreviewModal';
import { CommsLogModal } from './admin/modals/CommsLogModal';
import { MassmailPickModal } from './admin/modals/MassmailPickModal';
import { MassmailPasteModal } from './admin/modals/MassmailPasteModal';
import { ExcelTargetModal } from './admin/modals/ExcelTargetModal';
import { MassmailComposerModal } from './admin/modals/MassmailComposerModal';
import { InviteComposerModal } from './admin/modals/InviteComposerModal';
import { DeclineCheckModal } from './admin/modals/DeclineCheckModal';
import { AttachmentsModal } from './admin/modals/AttachmentsModal';
import { ReorderProgressOverlay } from './admin/modals/ReorderProgressOverlay';
import { DupCancelModal } from './admin/modals/DupCancelModal';
import { OverbookDecisionModal } from './admin/modals/OverbookDecisionModal';
import { WaitlistPositionModal } from './admin/modals/WaitlistPositionModal';
import { AdminAddMemberModal } from './admin/modals/AdminAddMemberModal';
import { PendingPeopleBox } from './admin/participants/PendingPeopleBox';
import { IdGapHintBox } from './admin/participants/IdGapHintBox';
import { DuplicateRegHintBox } from './admin/participants/DuplicateRegHintBox';
import { DuplicateInSubEventHintBox } from './admin/participants/DuplicateInSubEventHintBox';
import { MissingEmailHintBox } from './admin/participants/MissingEmailHintBox';
import { OverbookReviewBox } from './admin/participants/OverbookReviewBox';
import { TeamsSection } from './admin/participants/TeamsSection';
import { TeamMailModal } from './admin/participants/TeamMailModal';
import { ParticipantTable } from './admin/participants/ParticipantTable';
import { WaitlistTables } from './admin/participants/WaitlistTables';
import { CancelledList } from './admin/participants/CancelledList';
import { AdminToast } from './admin/sections/AdminToast';
import { InactiveAccountsBox } from './admin/sections/InactiveAccountsBox';
import { DuplicateEventsBox } from './admin/sections/DuplicateEventsBox';
import { EventDetailCard } from './admin/sections/EventDetailCard';
import { NextStepsBox } from './admin/sections/NextStepsBox';
import { BillingStatusStrip } from './admin/sections/BillingStatusStrip';
import { AdminActionsCard } from './admin/sections/AdminActionsCard';
import { KpiTiles } from './admin/sections/KpiTiles';
import { HotelPlanningSection } from './admin/sections/HotelPlanningSection';
import { QuizStatsSection } from './admin/sections/QuizStatsSection';
import { ActiveEventHintsBox } from './admin/sections/ActiveEventHintsBox';
import { AudienceVisibilityRow } from './admin/sections/AudienceVisibilityRow';
import { DangerZoneModal } from './admin/sections/DangerZoneModal';
import { ChangeLogModal } from './admin/sections/ChangeLogModal';
import { EventOverviewScreen } from './admin/sections/EventOverviewScreen';
import { ConsolidatedView } from './admin/sections/ConsolidatedView';
import { useCancelPipeline } from './admin/logic/useCancelPipeline';
import { useEditModalHandlers } from './admin/logic/useEditModalHandlers';
import { createKlammerActions } from './admin/logic/createKlammerActions';
import { useWaitlistActions } from './admin/logic/useWaitlistActions';
import { useTeamActions } from './admin/logic/useTeamActions';
import { useColumnConfig } from './admin/logic/useColumnConfig';
import { createExportActions } from './admin/logic/createExportActions';
import { useEventSelection } from './admin/logic/useEventSelection';
import { useMailComposers } from './admin/logic/useMailComposers';
import { createQrMailActions } from './admin/logic/createQrMailActions';

export default function AdminPage(): React.ReactElement {
  const isMobile = useIsMobile();
  const { navigate, selectedEventId } = useNavigation();
  // v14.11: zusätzlich `events` (alle Events inkl. Sub-Events) als `allEvents`
  // für die Parent-Lookup-Logik im konsolidierten View + im Sub-Event-Detail.
  const { events: allEvents, topLevelEvents: events, childEventsOf, isEventsLoading, getAllRegistrations, deleteEvent, countExternalRegistrations, getOrganizerArchivedEventIds, archiveEventForOrganizer, unarchiveEventForOrganizer, updateEvent, refreshEvents, addTeamMember, assignTeamlessToTeam, notifyExistingTeamMembers, transferTeamLead, registerForEvent, subscribeEventRealtime, sendCompleteRegistrationReminder } = useEvents();
  // v26.67: laufende „Erinnerung senden"-Aktion pro verwaister Anmeldung (Id).
  const [reminderBusyId, setReminderBusyId] = React.useState<number | null>(null);
  // v26.85: „Erinnerung senden" in der „Fehlende Klammer-Anmeldung"-Box (emailKey).
  const [missingReminderKey, setMissingReminderKey] = React.useState<string | null>(null);
  // v24.38: läuft gerade ein „Zur Klammer hinzufügen" für diese E-Mail?
  const [addingToKlammer, setAddingToKlammer] = React.useState<string | null>(null);
  // v30.14: Fortschritt des Sammel-Fixes „Alle fehlenden Klammer-Anmeldungen
  // still nachtragen" ('' = läuft nicht).
  const [bulkKlammerProgress, setBulkKlammerProgress] = React.useState<string>('');
  // v24.40: Modal „Assistenz zuordnen" — Person an eine gewählte Assistenz
  // übergeben (RegisteredBy + Zeilen-Autor auf Klammer + alle Sub-Events).
  const [assignAssistRow, setAssignAssistRow] = React.useState<ConsolidatedRow | null>(null);
  const [assignAssistValue, setAssignAssistValue] = React.useState('');
  const [assignAssistBusy, setAssignAssistBusy] = React.useState(false);
  const [isRefreshing, setIsRefreshing] = React.useState(false);
  const handleRefresh = async (): Promise<void> => {
    if (isRefreshing) return;
    setIsRefreshing(true);
    try {
      await refreshEvents();
      // Wenn ein Event gerade selektiert ist, auch dessen Registrations neu laden
      if (selectedEvent) {
        // v30.67: `getAllRegistrations` wirft nie — bei 429/403/500 kam bis
        // hierher `[]` zurück und ersetzte die vollständige Liste durch eine
        // leere, ohne Fehlertext. Wie im Auswahl-Pfad (v30.37) zählt der
        // Status: bei Fehler bleibt der geladene Stand stehen, und ein
        // Hinweis über der Tabelle sagt, dass er alt ist.
        // v30.67 (Review): über den gemeinsamen Helfer `reloadRegistrations`.
        // Bei Erfolg auch `regLoadError` löschen — der Auswahl-Pfad setzt ihn,
        // und nichts räumte ihn wieder weg: Nach einer 429 beim Öffnen stand
        // „konnte nicht gelesen werden" dauerhaft da, obwohl die Liste längst
        // wieder lesbar war. Und die Termin-Listen (Klammer-Modus) mit
        // anstoßen, sonst bleibt auch das Banner „N Teilnehmerliste(n) nicht
        // lesbar" nach einer Drosselung stehen.
        const regs = await reloadRegistrations();
        if (regs) {
          setRegLoadError('');
          setSubRegReloadTick(t => t + 1);
        }
      }
    } finally { setIsRefreshing(false); }
  };
  const { currentUser } = useCurrentUser();
  // v27.11: getGroupMembers für die Verteiler-Auflösung der Einladungs-Mail.
  const { isAdmin, siteUrl, currentUserRole, searchUser, searchUsers, getGroupMembers, isImpersonating, isFA } = useRoles();
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  // v30.66: Diese drei standen frueher mitten im Koerper (~Zeile 2600). Sie
  // haengen an nichts ausser `window` und werden von fast jeder ausgelagerten
  // Logik-Gruppe gebraucht — weiter unten waere jeder Auszug ein TDZ-Fehler.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const spfxContext = (window as any).__dexSpfxContext;
  const eventServiceRef = React.useMemo(() => spfxContext ? new EventService(spfxContext) : null, []);
  // v28.65: Für die Rollenliste (Namens-Reparatur).
  const spServiceRef = React.useMemo(() => spfxContext ? new SharePointService(spfxContext) : null, []);
  // Admin-Toast für Abmelde-/Nachrück-Feedback (seit v6.8):
  //  - 'cancelling': während die Abmeldung + Nachrück-Suche läuft (orange, Spinner)
  //  - 'promoted'  : erfolgreicher Nachrücker mit Namen + Typ (grün)
  //  - 'no-promote': Abmeldung ok, aber keiner auf der Warteliste (grau)
  // v30.66: Der Typ heisst jetzt `AdminToastState` und steht in admin/adminTypes.ts —
  // ausgelagerte Abmelde-Pipeline braucht ihn, ein Typ im Funktionskoerper
  // ist von aussen nicht referenzierbar.
  const [adminToast, setAdminToast] = React.useState<AdminToastState | null>(null);
  // v20.4: App-Modals statt nativer Browser-Dialoge.
  const { confirmDialog, showAlert } = useDialog();
  const [selectedEvent, setSelectedEvent] = React.useState<DeloitteEvent | null>(null);
  // v30.66: ebenfalls vorgezogen — die Wartelisten-Logik liest sie, und die
  // stand vorher weit vor der urspruenglichen Deklaration.
  // Seit v6.5: getrennte Wartelisten bei B2Run-Split-Kapazitäten (Durchstarter/Funstarter).
  // Die Split-Aktivierung erkennen wir daran, dass beide Kapazitäts-Felder gesetzt und > 0 sind.
  const isSplitCapacity = !!selectedEvent
    && typeof selectedEvent.durchstarterCapacity === 'number'
    && typeof selectedEvent.funstarterCapacity === 'number'
    && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
  // v26.18: Beim Wechsel zwischen Event-Detail und Event-Liste (das ist ein
  // interner State-Wechsel, KEINE Seiten-Navigation → der globale Scroll-Reset
  // greift nicht) immer nach oben scrollen. Betrifft besonders „Zurück" aus der
  // Detailansicht: vorher blieb die Liste weit unten gescrollt.
  React.useEffect(() => {
    const toTop = (el: Element | null): void => { if (el) { try { (el as HTMLElement).scrollTop = 0; } catch { /* */ } } };
    try { window.scrollTo(0, 0); } catch { /* */ }
    toTop(document.scrollingElement);
    toTop(document.documentElement);
    toTop(document.body);
    ['[data-automation-id="contentScrollRegion"]', '.SPPageChrome-content', '#spPageCanvasContent', '[class*="contentScrollRegion"]']
      .forEach(sel => { try { document.querySelectorAll(sel).forEach(toTop); } catch { /* */ } });
  }, [selectedEvent?.id]);
  const [registrations, setRegistrations] = React.useState<SPRegistration[]>([]);
  // v30.67 (Review): EIN Nachlade-Pfad für alle Schreibpfade des Organizer
  // Centers. `getAllRegistrations` wirft nie — jede Stelle, die nach einem
  // Schreibvorgang `setRegistrations(await getAllRegistrations(id))` rief,
  // ersetzte bei 429/403/500 die volle Liste still durch `[]` („Noch keine
  // Teilnehmer registriert", alle KPI-Kacheln 0, kein Hinweis). Der Refresh-
  // Knopf und der Echtzeit-Push hatten das seit fd6cd155 richtig, rund 25
  // weitere Stellen nicht — „Person überall löschen" traf mit dem Reload
  // gern in die Drosselung, die das Löschen selbst ausgelöst hatte.
  // Semantik wie dort: bei Fehler bleibt der alte Stand stehen, der Hinweis
  // über der Tabelle sagt, dass er alt ist. Rückgabe `null` heißt „nicht
  // lesbar", damit Aufrufer, die mit der Liste weiterrechnen, abbrechen
  // statt mit `[]` zu arbeiten.
  // Steht bewusst HIER: `useCancelPipeline` weiter unten bekommt es als
  // ctx-Feld — darunter deklariert wäre es ein TDZ-Fehler. `setRegStaleHint`
  // ist erst später deklariert; als Closure-Zugriff ist das unkritisch (läuft
  // erst nach dem Render), dasselbe Muster wie `handleRefresh` oben.
  const reloadRegistrations = async (): Promise<SPRegistration[] | null> => {
    if (!selectedEvent) return null;
    let failedStatus = -1;
    const regs = await getAllRegistrations(selectedEvent.id, st => { failedStatus = st; });
    if (failedStatus >= 0) {
      setRegStaleHint(failedStatus === 401 || failedStatus === 403 ? 'denied' : 'transient');
      return null;
    }
    setRegStaleHint('');
    setRegistrations(regs);
    return regs;
  };
  // v28.39: Hotel-Planung eingeklappt starten — Events ohne Übernachtung
  // sollen von dem Abschnitt nichts merken.
  const [hotelPanelOpen, setHotelPanelOpen] = React.useState(false);
  // v24.75: Echtzeit-Push auf die Teilnehmerliste des gewählten Events. Meldet
  // sich jemand an/ab, kommt eine Push-Benachrichtigung → die Tabelle (und die
  // Zähler) laden leise nach. Organizer/Admin haben Vollzugriff → Subscription
  // greift. Best-effort: ohne Socket bleibt der manuelle/Refresh-Knopf.
  React.useEffect(() => {
    const ev = selectedEvent;
    if (!ev || !ev.id || !(ev.subsiteUrl || '').trim()) return undefined;
    let cancelled = false;
    let cleanupSocket: (() => void) | null = null;
    const reload = (): void => {
      // v30.67: Ein Push während einer Anmeldewelle trifft gern auf 429 —
      // und `getAllRegistrations` liefert dann `[]` statt zu werfen. Das
      // `.catch` hier war toter Code; die Tabelle, alle KPI-Kacheln und die
      // Warteliste fielen still auf 0. Bei Fehler bleibt der alte Stand, mit
      // Hinweis — NICHT `regLoadError`, das würde die Tabelle ersetzen.
      let failedStatus = -1;
      getAllRegistrations(ev.id, st => { failedStatus = st; }).then(r => {
        if (cancelled) return;
        if (failedStatus >= 0) {
          setRegStaleHint(failedStatus === 401 || failedStatus === 403 ? 'denied' : 'transient');
          return;
        }
        setRegStaleHint('');
        setRegistrations(r);
      }).catch(() => { /* */ });
    };
    subscribeEventRealtime(ev.id, 'participants', reload)
      .then(c => { if (cancelled) c(); else cleanupSocket = c; })
      .catch(() => { /* best-effort */ });
    return () => { cancelled = true; if (cleanupSocket) cleanupSocket(); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, selectedEvent?.subsiteUrl]);
  // v22.7: Konten-Aktiv-Check — Adressen (lowercase) der Teilnehmer, deren
  // Deloitte-Konto nicht mehr aktiv ist (Person hat womöglich das Unternehmen
  // verlassen). Wird im Hintergrund max. 1×/Tag pro Event geprüft.
  const [inactiveAccounts, setInactiveAccounts] = React.useState<string[]>([]);
  // v23.2: Doppel-Anmeldungen erkennen. Eine E-Mail mit ≥2 NICHT-abgemeldeten
  // Zeilen in derselben Teilnehmerliste = Duplikat (z.B. dieselbe Person in
  // zwei Teams). Wird oben in einer Hinweis-Box surfaced + die Zeilen werden
  // rot markiert. duplicateEmails = Set der betroffenen Adressen (lowercase).
  const duplicateEmails = React.useMemo<Set<string>>(() => {
    const counts: Record<string, number> = {};
    for (const r of registrations) {
      // v28.21: nur wirklich AKTIVE Zeilen zählen (vorher „alles außer
      // Abgemeldet" — damit galt z.B. eine „No-Show"-Zeile als Duplikat).
      if (DUP_ACTIVE_STATI.indexOf(r.Status || '') < 0) continue;
      const em = (r.ParticipantEmail || '').trim().toLowerCase();
      if (!em) continue;
      counts[em] = (counts[em] || 0) + 1;
    }
    const dup = new Set<string>();
    Object.keys(counts).forEach(em => { if (counts[em] > 1) dup.add(em); });
    return dup;
  }, [registrations]);
  // v23.3: Aktive Zeilen OHNE gültige E-Mail. Solche Anmeldungen belegen einen
  // Platz, zaehlen aber in den entdoppelten Zahlen (Kachel/Klammer/KPI) frueher
  // nicht mit (E-Mail = Dedup-Schlüssel) → „188 statt 190". Ausserdem bekommen
  // sie KEINE Bestätigung/QR/Outlook. Oben als Hinweis-Box surfacen.
  const missingEmailRegs = React.useMemo<SPRegistration[]>(() => {
    return registrations.filter(r => {
      if ((r.Status || '') === 'Abgemeldet') return false;
      const em = (r.ParticipantEmail || '').trim();
      return !em || em.indexOf('@') < 0;
    });
  }, [registrations]);
  // v23.2: Duplikat-Abmelde-Modal — gezogene Zeile + Entscheidung still löschen
  // (Duplikat entfernen, keine Mail/Outlook/Nachrücken) vs. normal abmelden.
  const [dupCancelReg, setDupCancelReg] = React.useState<SPRegistration | null>(null);
  const [dupCancelBusy, setDupCancelBusy] = React.useState(false);

  // v23.2: Standard-Abmeldung einer Teilnehmer-Zeile (extrahiert aus dem
  // Abmelden-Button, damit das Duplikat-Modal denselben „normal abmelden"-Pfad
  // wiederverwenden kann). Enthält KEINEN Confirm — der Aufrufer bestätigt.
  // Spiegelt das bisherige Inline-Verhalten 1:1 (vergangenes Event → still,
  // sonst Abmelde-Mail + Outlook-Ausladen + Nachrücken + ID-Reorder).

  // v30.66: useCancelPipeline — Rumpf in logic/useCancelPipeline.ts.
  const {
    buildCancellationMail, cleanupShadowDuplicates, isSyncingRegistry,
    performSilentDuplicateDelete, performStandardCancel, setIsSyncingRegistry,
    setSyncRegistryResult, shadowDupBusy, syncRegistryResult,
  } = useCancelPipeline({
    allEvents, confirmDialog, currentUser, duplicateEmails, eventServiceRef,
    isDe, registrations, reloadRegistrations, selectedEvent, setAdminToast, showAlert,
  });
  // v22.16: „Hinweise"-Box für aktive Events — Busy-State für den 1-Klick-
  // Sprach-Fix + Tick, damit „Ausblenden" (localStorage) sofort re-rendert.
  const [hintLangBusy, setHintLangBusy] = React.useState(false);
  const [hintsDismissTick, setHintsDismissTick] = React.useState(0);
  // v24.50: Welche Hinweise sind aufgeklappt? (Default: alle eingeklappt —
  // nur die Überschriften zeigen, Klick klappt den Inhalt auf.)
  const [expandedHintIds, setExpandedHintIds] = React.useState<Set<string>>(new Set());
  // v11.97/v11.98: bei Events mit Split-Kapazität (zwei Gruppen) wird die
  // Aktiv-Teilnehmer-Tabelle standardmäßig nach Gruppe getrennt angezeigt
  // (kleinere Gruppe zuerst). Per Toggle umschaltbar auf zusammengeführte
  // Sicht. Default: 'split'. Bei Events ohne Split-Kapazität ohne Wirkung.
  const [splitParticipantsView, setSplitParticipantsView] = React.useState<'split' | 'merged'>('split');
  // v14.11: subEventsOnlyMode-Konsolidierung. Wenn das selektierte Hauptevent
  // im „Nur Sub-Events"-Modus ist und Sub-Events hat, hat das Hauptevent
  // selbst keine direkten Teilnehmer — stattdessen werden alle Sub-Event-
  // Teilnehmer pro Person zu einer Zeile aggregiert (Matrix-View mit einer
  // X-Spalte pro Sub-Event). Hier halten wir die rohen Registrierungen je
  // Sub-Event.
  const [subEventRegsByEventId, setSubEventRegsByEventId] = React.useState<Record<string, SPRegistration[]>>({});
  const [isLoadingSubEventRegs, setIsLoadingSubEventRegs] = React.useState(false);
  // v30.37: Titel der Termine, deren Teilnehmerliste nicht lesbar war
  // (Berechtigung/gelöschte Subsite). Leer = alles gelesen.
  // v30.67 (Review): mit HTTP-Status — das Banner unten nennt „erst
  // nachträglich als Organizer benannt" nur noch bei 401/403/404; eine 429
  // oder ein Netzfehler ist keine Rechtefrage und schickte Organizer auf die
  // Aktion „Berechtigungen reparieren", die daran nichts ändert.
  const [deniedSubEventLists, setDeniedSubEventLists] = React.useState<DeniedSubEventList[]>([]);
  // v22.59: manueller Reload-Trigger für die Sub-Event-Regs (z.B. nach dem
  // Löschen einer konsolidierten Abmeldung).
  const [subRegReloadTick, setSubRegReloadTick] = React.useState(0);
  const [expandedConsolidatedEmail, setExpandedConsolidatedEmail] = React.useState<string | null>(null);
  // v23.5: Personen-Spalten (Vorname/Nachname/Email/Job Title/Standort) in der
  // konsolidierten Matrix einklappbar — eingeklappt steht nur Foto + Name, damit
  // die Event-spezifischen Spalten mehr Platz bekommen.
  // v23.32: Standard = eingeklappt (Foto + zweizeilige Person: Name fett,
  // darunter „Position • Standort"). Gilt für konsolidierte UND normale Tabelle.
  const [personalColsCollapsed, setPersonalColsCollapsed] = React.useState(true);
  // v30.21: Hover-State für den Auf-/Zuklapp-Knopf der Personen-Spalten
  // (Inline-Styles können kein :hover).
  const [colToggleHover, setColToggleHover] = React.useState(false);
  // v24.31: Teilnehmer-Detailmodal — Klick auf eine Person zeigt Kontakt-
  // Detailinfos (Foto, E-Mail, MS-Teams-Chat, Position/Standort/Unternehmen/
  // Abteilung/Telefon/Status). Bewusst „Detailinfos", nicht die Fragen/Antworten.
  const [participantDetail, setParticipantDetail] = React.useState<{
    name: string; email: string; jobTitle: string; location: string; company: string;
    department: string; phone: string; status: string; tid: number | null;
  } | null>(null);
  // v14.11: eigene Sort-States für den Matrix-View. `consolidatedSort` kann
  // 'id' | 'vorname' | 'nachname' | 'email' | 'jobTitle' | 'location' |
  // 'child:<eventId>' sein. Default: 'nachname' aufsteigend.
  // v15.23: Default-Sort im konsolidierten View jetzt chronologisch
  // nach erster Anmeldung (früheste zuerst), nicht mehr alphabetisch
  // nach Nachname. Damit ist die # in der Liste die Reihenfolge der
  // Anmeldung, nicht die Alphabet-Position.
  const [consolidatedSort, setConsolidatedSort] = React.useState<string>('id');
  const [consolidatedSortAsc, setConsolidatedSortAsc] = React.useState<boolean>(true);
  // v11.0: Bei Events mit Teilnehmer-Upload alle Attachment-Listen
  // einmalig laden, sobald sich registrations oder das ausgewählte
  // Event ändern. Damit zeigt der „Anhang"-Button in der Action-Spalte
  // sofort die korrekte Anzahl.
  React.useEffect(() => {
    // v19.0: Attachments auch laden, wenn das Event ein Dokument-Custom-Feld hat
    // (nicht nur beim generischen Attendee-Upload).
    const hasDocField = (selectedEvent?.eventSpecificFields || []).some(f => f.type === 'document');
    if (!selectedEvent || (!selectedEvent.allowAttendeeUpload && !hasDocField) || !eventServiceRef || !selectedEvent.subsiteUrl) {
      setAttachmentsByReg({});
      return;
    }
    const subsiteUrl = selectedEvent.subsiteUrl;
    const ids = registrations.map(r => r.Id).filter(Boolean);
    let cancelled = false;
    (async () => {
      const map: Record<number, Array<{ fileName: string; serverRelativeUrl: string }>> = {};
      for (const id of ids) {
        try {
          const list = await eventServiceRef!.listRegistrationAttachments(subsiteUrl, id);
          if (list.length > 0) map[id] = list;
        } catch { /* */ }
      }
      if (!cancelled) setAttachmentsByReg(map);
    })().catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, selectedEvent?.allowAttendeeUpload, registrations.length]);
  // v9.29: Header-Refresh-Button triggert ein globales Event — wir hooken uns ein.
  React.useEffect(() => {
    const onRefresh = (): void => { void handleRefresh(); };
    window.addEventListener('dex-refresh-page', onRefresh);
    return () => window.removeEventListener('dex-refresh-page', onRefresh);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent]);
  // v14.11: subEventsOnlyMode — alle Sub-Event-Anmeldungen einsammeln, um
  // den konsolidierten Matrix-View pro Person zu rendern. Nur aktiv, wenn
  // das selektierte Event tatsächlich Hauptevent ohne eigene Anmeldungen
  // (subEventsOnlyMode) ist und es Sub-Events gibt.
  React.useEffect(() => {
    if (!selectedEvent || !selectedEvent.subEventsOnlyMode) {
      setSubEventRegsByEventId({});
      return;
    }
    const children = childEventsOf(selectedEvent.id);
    if (children.length === 0) {
      setSubEventRegsByEventId({});
      return;
    }
    let cancelled = false;
    setIsLoadingSubEventRegs(true);
    (async () => {
      const map: Record<string, SPRegistration[]> = {};
      // v30.37: Termine, deren Teilnehmerliste NICHT gelesen werden konnte.
      // Bis v30.36 wurde daraus stillschweigend `[]` — für einen Organizer
      // ohne Rechte auf den Sub-Event-Subsites sah ein volles Event dadurch
      // exakt so aus wie ein leeres (jeder Tag „0", KPI-Kacheln 0,
      // „Teilnehmer (0)"). Der Rückruf existiert seit v29.3, nur genutzt hat
      // ihn hier niemand.
      const denied: DeniedSubEventList[] = [];
      for (const ch of children) {
        try {
          // v30.67: JEDER Rückruf heißt „nicht gelesen" — nicht nur vier
          // Codes. Der Rückruf feuert per Definition bei jedem nicht-ok-Status;
          // 429 (Drosselung beim 12. von 19 Terminen) und 5xx fehlten in der
          // Liste und wurden zu „0 Teilnehmer" ohne Banner. Referenz:
          // `analyzeRegistryAgainstLists` wertet nur 404/410 als eindeutig,
          // alles andere als „sagt nichts über den Inhalt".
          const regs = await getAllRegistrations(ch.id, st => {
            denied.push({ title: ch.title || ch.id, status: st });
          });
          map[ch.id] = regs;
        } catch {
          map[ch.id] = [];
          // v30.67 (Review): -1 = Ausnahme statt HTTP-Status → „unbekannt".
          denied.push({ title: ch.title || ch.id, status: -1 });
        }
      }
      if (!cancelled) {
        setSubEventRegsByEventId(map);
        setDeniedSubEventLists(denied);
        setIsLoadingSubEventRegs(false);
      }
    })().catch(() => { if (!cancelled) setIsLoadingSubEventRegs(false); });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, selectedEvent?.subEventsOnlyMode, subRegReloadTick]);
  // v15.14: Wenn ein Sub-Event direkt selektiert wurde, laden wir die
  // Registrierungen des Parent-Events mit, damit die Pastel-A-Spalten
  // (Custom-Fields des Hauptevents) pro Teilnehmer-Zeile mit den
  // tatsächlichen Antworten aus der Parent-Registrierung gefüllt
  // werden können. Vorher waren diese Zellen leer („-"), weil die
  // Sub-Event-Registrierung diese Antworten nicht enthält.
  const [parentRegsByEmail, setParentRegsByEmail] = React.useState<Record<string, SPRegistration>>({});
  React.useEffect(() => {
    if (!selectedEvent || !selectedEvent.parentEventId) {
      setParentRegsByEmail({});
      return;
    }
    let cancelled = false;
    (async () => {
      try {
        const parentRegs = await getAllRegistrations(selectedEvent.parentEventId!);
        const map: Record<string, SPRegistration> = {};
        for (const r of parentRegs) {
          const key = (r.ParticipantEmail || '').toLowerCase().trim();
          if (key) map[key] = r;
        }
        if (!cancelled) setParentRegsByEmail(map);
      } catch {
        if (!cancelled) setParentRegsByEmail({});
      }
    })().catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, selectedEvent?.parentEventId]);
  const [isLoadingRegs, setIsLoadingRegs] = React.useState(false);
  const [regLoadError, setRegLoadError] = React.useState('');
  // v30.67 (Review): „Liste nicht lesbar" als Flag für die Zähler. Die
  // KPI-Kacheln und „Aktuell registriert" stehen außerhalb des
  // `regLoadError`-Zweigs und rechneten aus der leeren Liste eine „0" —
  // neben dem Text „kein Zugriff" stand also weiter „0 angemeldet". Mit dem
  // Flag rendern sie „—". Bewusst NUR `regLoadError`: bei `regStaleHint`
  // steht der zuletzt geladene Stand in der Tabelle, dann dürfen die Kacheln
  // dieselben (alten) Zahlen zeigen — Kachel und Tabelle bleiben deckungsgleich.
  const regsUnknown = !!regLoadError;
  // v30.67 (Review): Im Klammer-Modus sind Summen über die Termin-Listen nur
  // eine Untergrenze, sobald eine Liste nicht lesbar war → Kacheln zeigen „≥ N".
  const subListsIncomplete = deniedSubEventLists.length > 0;
  // v30.67: Der LETZTE Nachlade-Versuch (Refresh-Knopf, Echtzeit-Push) ist
  // fehlgeschlagen; die Tabelle zeigt deshalb den zuvor geladenen Stand.
  // Bewusst getrennt von `regLoadError`: Das ersetzt die Tabelle — hier gibt
  // es aber gültige Daten, nur ältere. Der Auswahl-Pfad setzt frisch.
  const [regStaleHint, setRegStaleHint] = React.useState<'' | 'denied' | 'transient'>('');
  React.useEffect(() => { setRegStaleHint(''); }, [selectedEvent?.id]);
  // v18.24: beim Event-/Tab-Wechsel die aktuelle Höhe der Detail-Card
  // „einfrieren", solange die Teilnehmer neu geladen werden — sonst klappt
  // die Card auf die „Lade..."-Zeile zusammen und springt danach wieder auf
  // (klein→groß-Flackern). null = keine Reservierung aktiv.
  const detailCardRef = React.useRef<HTMLDivElement>(null);
  const [reservedDetailHeight, setReservedDetailHeight] = React.useState<number | undefined>(undefined);
  // v23.6: Breiten-Reservierung gegen das „Springen" der Detail-Karte beim
  // Wechsel zwischen Klammer und Sub-Events (manche Tabs sind breiter, z.B.
  // konsolidierte Matrix vs. schmales Sub-Event). Die Karte wächst auf die
  // größte Inhaltsbreite INNERHALB derselben Event-Gruppe (Hauptevent + seine
  // Sub-Events) und schrumpft danach nicht mehr — das breiteste Event gibt die
  // Breite vor. Reset bei Wechsel auf eine andere Event-Gruppe.
  const [reservedDetailWidth, setReservedDetailWidth] = React.useState<number | undefined>(undefined);
  const widthGroupRef = React.useRef<string>('');
  // v23.6: Misst nach jedem relevanten Render die tatsächliche Inhaltsbreite
  // (scrollWidth inkl. überlaufender Tabelle) und hält das Maximum pro
  // Event-Gruppe (Hauptevent-ID = parentEventId || id) fest. Wird als minWidth
  // an die Karte gelegt → schmale Tabs bleiben so breit wie der breiteste,
  // und während des Nachladens schrumpft nichts (kein „erst klein, dann breit").
  React.useLayoutEffect(() => {
    if (!selectedEvent || !detailCardRef.current) return;
    const groupId = selectedEvent.parentEventId || selectedEvent.id;
    const w = detailCardRef.current.scrollWidth;
    if (widthGroupRef.current !== groupId) {
      // Neue Event-Gruppe → frisch mit der natürlichen Breite starten.
      widthGroupRef.current = groupId;
      setReservedDetailWidth(w || undefined);
    } else {
      setReservedDetailWidth(prev => (prev && prev >= w ? prev : w));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, isLoadingRegs, registrations]);
  const [deletingId, setDeletingId] = React.useState<string | null>(null);
  // v9.0: Danger-Zone-Modal — User muss den Event-Titel exakt (lowercase)
  // eintippen bevor der Lösch-Button aktiv wird. Schutz gegen versehentliche
  // Löschungen (früher: Click-to-Confirm-Pattern, war zu schwach).
  const [confirmDeleteEvent, setConfirmDeleteEvent] = React.useState<DeloitteEvent | null>(null);
  const [confirmDeleteText, setConfirmDeleteText] = React.useState('');
  // v24.0: Lösch-Berechtigungs-Prüfung. Beim Öffnen des Lösch-Dialogs wird
  // ermittelt, ob das Event Anmeldungen über das Organizer-Team hinaus hat
  // ("ehemals aktiv"). Wenn ja: nur Admins, und frühestens 1 Jahr nach
  // Event-Ende. Ohne Fremd-Anmeldungen (Entwurf/leer): einfaches Ja genügt.
  const [deletePolicy, setDeletePolicy] = React.useState<
    { loading: true } |
    { loading: false; allowed: boolean; requiresTitle: boolean; externalCount: number; reason?: string }
    | null
  >(null);
  React.useEffect(() => {
    if (!confirmDeleteEvent) { setDeletePolicy(null); return; }
    let cancelled = false;
    setDeletePolicy({ loading: true });
    (async () => {
      // v30.66: Ein Fehler beim Zaehlen ist KEINE Null. Vorher lief der
      // catch-Zweig auf `externalCount = 0` — und 0 ist genau der Wert, der
      // unten die Loeschung freigibt. Ein nicht lesbares Event wird jetzt
      // gesperrt statt freigegeben.
      let externalCount = 0;
      let countFailed = false;
      try { externalCount = await countExternalRegistrations(confirmDeleteEvent); } catch { countFailed = true; }
      if (cancelled) return;
      if (countFailed) {
        setDeletePolicy({ loading: false, allowed: false, requiresTitle: false, externalCount: 0,
          reason: isDe
            ? 'Die Teilnehmerliste dieses Events konnte nicht gelesen werden. Ob es Anmeldungen ausserhalb des Organizer-Teams gibt, ist damit unbekannt — geloescht wird deshalb nicht. Bitte spaeter erneut versuchen oder die Berechtigungen auf den Termin-Subsites pruefen.'
            : 'The attendee list of this event could not be read, so it is unknown whether there are registrations beyond the organizer team. Deletion is blocked. Please retry later or check the permissions on the sub-event subsites.' });
        return;
      }
      const endRaw = confirmDeleteEvent.endDate || confirmDeleteEvent.startDate;
      const endTs = endRaw ? new Date(endRaw).getTime() : 0;
      // v26.32: Aufbewahrung 3 Monate (statt vormals 1 Jahr) — konsistent mit dem
      // Teilnehmerlisten-Löschkonzept (Kalendermonate wie participantDeleteDueTs).
      const retentionCutoff = (() => { const d = new Date(); d.setMonth(d.getMonth() - 3); return d.getTime(); })();
      const overRetention = endTs > 0 && endTs < retentionCutoff;
      // v28.7: Demo-Events (Titel „Demo-…" aus den Demo-Vorlagen) sind reine
      // Vorführ-Daten — ein Admin darf sie IMMER sofort löschen, auch wenn
      // sich Teilnehmer außerhalb des Organizer-Teams angemeldet haben. Die
      // 3-Monats-Aufbewahrung schützt echte Teilnehmerlisten, keine Demos.
      const isDemoEvent = /^demo[-\s]/i.test((confirmDeleteEvent.title || '').trim());
      if (externalCount > 0 && isDemoEvent && isAdmin) {
        setDeletePolicy({ loading: false, allowed: true, requiresTitle: false, externalCount });
      } else if (externalCount > 0) {
        // Ehemals aktiv (echte Teilnehmer) → geschützt.
        if (!isAdmin) {
          setDeletePolicy({ loading: false, allowed: false, requiresTitle: false, externalCount,
            reason: isDe
              ? 'Dieses Event hatte Anmeldungen über das Organizer-Team hinaus. Es darf nur von einem Admin gelöscht werden — und das frühestens 3 Monate nach dem Event (Aufbewahrung der Teilnehmerliste). Du kannst das Event stattdessen archivieren (aus deiner Übersicht ausblenden).'
              : 'This event had registrations beyond the organizer team. Only an admin may delete it — and only three months after the event at the earliest. You can archive it instead (hide from your overview).' });
        } else if (!overRetention) {
          setDeletePolicy({ loading: false, allowed: false, requiresTitle: false, externalCount,
            reason: isDe
              ? 'Dieses Event hat Anmeldungen über das Organizer-Team hinaus. Die Teilnehmerliste wird 3 Monate aufbewahrt — das vollständige Löschen des Events ist erst 3 Monate nach dem Event-Ende möglich.'
              : 'This event has registrations beyond the organizer team. The attendee list is kept for three months — full deletion of the event is only possible three months after the event ends.' });
        } else {
          setDeletePolicy({ loading: false, allowed: true, requiresTitle: true, externalCount });
        }
      } else {
        // Keine Fremd-Anmeldungen (Entwurf/leer) → einfaches Ja genügt.
        setDeletePolicy({ loading: false, allowed: true, requiresTitle: false, externalCount });
      }
    })().catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [confirmDeleteEvent, isAdmin]);
  // v9.0: ChangeLog-Modal — Admin/Organizer sehen den Audit-Log aller
  // Event- und Teilnehmer-Änderungen (DEX_ChangeLog).
  const [showChangeLogModal, setShowChangeLogModal] = React.useState(false);
  const [changeLogEntries, setChangeLogEntries] = React.useState<Array<{
    Id: number; Created: string; Action: string; TargetType: string;
    TargetId: string; TargetName: string; EventId: string; EventTitle: string;
    ActorName: string; ActorEmail: string; Details: string;
  }>>([]);
  const [changeLogLoading, setChangeLogLoading] = React.useState(false);
  const [changeLogFilterAction, setChangeLogFilterAction] = React.useState('');
  const [changeLogFilterEvent, setChangeLogFilterEvent] = React.useState('');
  const [changeLogFilterActor, setChangeLogFilterActor] = React.useState('');
  // Self-Actions (User registriert/storniert sich selbst) sind Datenrauschen
  // im Audit-Log — Admins/Organizer wollen normalerweise nur Aktionen sehen,
  // die jemand AUF einen anderen User angewendet hat. Marker im Details-JSON:
  // `"asActor":"self"` bei selbst durchgeführten Registrierungen/Stornos.
  const [changeLogHideSelf, setChangeLogHideSelf] = React.useState(true);
  const openChangeLog = async (): Promise<void> => {
    if (!eventServiceRef) return;
    setShowChangeLogModal(true);
    setChangeLogLoading(true);
    try {
      const entries = await eventServiceRef.readChangeLog({ top: 500 });
      setChangeLogEntries(entries);
    } finally {
      setChangeLogLoading(false);
    }
  };
  const [isDeleting, setIsDeleting] = React.useState(false);
  const [copiedEmails, setCopiedEmails] = React.useState(false);
  const [copiedDeepLink, setCopiedDeepLink] = React.useState(false);
  const [isSendingQR, setIsSendingQR] = React.useState(false);
  const [qrSentCount, setQrSentCount] = React.useState(0);
  // v9.15: QR-Code-Versand-Modal mit Test-/Volldurchlauf. v20.7: der
  // Auto-Send-Toggle ist entfallen — Auto-Send ist immer aktiv.
  const [qrSendModalOpen, setQrSendModalOpen] = React.useState(false);
  // v30.36: Sammel-Einstieg „QR-Codes und Check-In". Zwei Schritte, damit pro
  // Bildschirm nur EINE Entscheidung ansteht: erst wofuer, dann wie.
  const [checkInHubOpen, setCheckInHubOpen] = React.useState(false);
  const [checkInHubStep, setCheckInHubStep] = React.useState<'choose' | 'checkin'>('choose');
  // v30.36: Aufklapper im QR-Versand-Modal. Erklaerendes soll auf Abruf da
  // sein, nicht beim Oeffnen den Blick auf die drei Schritte verstellen.
  const [qrHelpOpen, setQrHelpOpen] = React.useState(false);
  const [qrSubMailsOpen, setQrSubMailsOpen] = React.useState(false);
  // v30.5: Modal „Event-Abrechnung" (Versand an F&A + Historie).
  const [billingPanelOpen, setBillingPanelOpen] = React.useState(false);
  const [qrSendResult, setQrSendResult] = React.useState<string | null>(null);
  // v9.37: Vorschau der QR-Code-Mail (analog zur Live-Vorschau im Event-Wizard
  // unter „Kommunikation"). Der Organizer sieht damit vorab genau die Mail, die
  // beim Versand rausgeht — inklusive echtem QR-Code für ihn selbst als Empfänger.
  const [qrPreviewOpen, setQrPreviewOpen] = React.useState(false);
  const [qrPreviewHtml, setQrPreviewHtml] = React.useState('');
  const [qrPreviewSubject, setQrPreviewSubject] = React.useState('');
  const [qrPreviewLoading, setQrPreviewLoading] = React.useState(false);
  // v22.18: QR-Mail-Text pro Event anpassbar (HtmlEditorModal mit Live-
  // Vorschau, gespeichert im Event → gilt auch für den Auto-Versand).
  const [qrEditOpen, setQrEditOpen] = React.useState(false);
  // v29.26: Ziel des QR-Mail-Editors. null = das geöffnete Event selbst;
  // gesetzt = ein Sub-Event, dessen QR-Mail vom Hauptevent aus gestaltet
  // wird (der Override liegt IMMER auf der Zeile des Ziel-Events).
  const [qrEditTarget, setQrEditTarget] = React.useState<DeloitteEvent | null>(null);
  const [qrEditSubject, setQrEditSubject] = React.useState('');
  const [qrEditHeading, setQrEditHeading] = React.useState('');
  const [qrEditSubheading, setQrEditSubheading] = React.useState('');
  const [qrEditBody, setQrEditBody] = React.useState('');
  const [qrEditSaving, setQrEditSaving] = React.useState(false);
  const [qrEditSampleBlock, setQrEditSampleBlock] = React.useState('');
  // v30.60: Das Beispiel-QR-Bild getrennt halten — der Block daneben wird bei
  // jedem Wechsel der Block-Sprache neu gebaut, das Bild selbst nicht.
  const [qrEditSampleImg, setQrEditSampleImg] = React.useState('');
  // v30.52: Kopf-Bild der QR-Mail. Anders als bei Massen-/Einladungsmail wird
  // das hier GESPEICHERT (im QR-Override), weil die QR-Mail auch automatisch
  // bei neuen Anmeldungen rausgeht — dann gibt es kein Fenster, in dem man
  // die Breite noch einstellen könnte.
  const [qrHeaderImage, setQrHeaderImage] = React.useState<MailHeaderImage>(MAIL_HEADER_IMAGE_DEFAULT);
  // v30.60: Sprache des festen Blocks neben dem QR-Code. '' = der Mail-Sprache
  // des Events folgen (Normalfall); DE/EN nur, wenn Mailtext und Event-Sprache
  // bewusst auseinandergehen.
  const [qrBlockLang, setQrBlockLang] = React.useState<'' | 'DE' | 'EN'>('');
  // v30.61: Eigener Hinweistext unter der Teilnehmer-ID. '' = Standardsatz in
  // der Block-Sprache; '-' = gar kein Hinweis (s. buildQrBlockHtml).
  const [qrBlockNote, setQrBlockNote] = React.useState('');
  const [qrEventPhotoB64, setQrEventPhotoB64] = React.useState('');
  const [searchQuery, setSearchQuery] = React.useState('');
  // v29.26: „Teilnehmer hinzufügen"-Dialog (Organizer-Ausnahme-Weg).
  const [addParticipantsOpen, setAddParticipantsOpen] = React.useState(false);
  // v29.32: Sichtbarkeits-Zeile über der Teilnehmerliste — eingeklappt zeigt
  // sie, wie viele Personen das Event sehen können; aufgeklappt, woraus sich
  // das zusammensetzt. `visibilityResolved` hält das Ergebnis einer LIVE-
  // Auflösung der Verteiler (null = noch nicht ausgeführt; dann gilt die beim
  // Event-Speichern eingefrorene Liste `audienceResolvedEmails`).
  const [visibilityOpen, setVisibilityOpen] = React.useState(false);
  // v29.36: Nicht mehr nur Adressen — der Verteiler-Abruf liefert ohnehin Name,
  // Position und Standort mit. Die brauchte der Nachfass-Schritt, um Personen
  // als Personen zu zeigen statt als Adressliste.
  const [visibilityResolved, setVisibilityResolved] = React.useState<AudiencePerson[] | null>(null);
  const [visibilityBusy, setVisibilityBusy] = React.useState(false);
  const [pendingCheckBusy, setPendingCheckBusy] = React.useState(false);
  // v29.36: Erster Schritt des Nachfassens — WER fehlt noch. Die Mail kommt
  // erst danach; vorher sah der Organizer nur einen vorbefüllten Mail-Dialog
  // und musste den Adressen glauben.
  const [pendingPeople, setPendingPeople] = React.useState<{ people: AudiencePerson[]; reachable: number } | null>(null);
  // v29.36: Lange Adresslisten in der Sichtbarkeit erst auf Wunsch ganz zeigen.
  const [visibilityAllAddresses, setVisibilityAllAddresses] = React.useState(false);
  // v26.11: Sprung zur Person in der Teilnehmerliste (aus der „Konto inaktiv"-
  // Hinweisbox) — filtert die Liste auf die Adresse und scrollt sie in den Blick.
  const participantListRef = React.useRef<HTMLDivElement>(null);
  const jumpToParticipant = (email: string): void => {
    setSearchQuery(email);
    window.setTimeout(() => {
      try { participantListRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }); } catch { /* */ }
    }, 80);
  };
  // v23.33: string statt fixer Union — erlaubt Sortierung nach Custom-Field-
  // Spalten (Keys „cf-…"/„cfp-…") und Job Title / Standort.
  const [sortColumn, setSortColumn] = React.useState<string>('id');
  const [sortAsc, setSortAsc] = React.useState(true);
  // v17.8: Sortierung der Warteliste (analog Teilnehmerliste). Default 'pos'
  // = Reihenfolge nach TeilnehmerID asc (FIFO-Position der Warteliste).
  const [waitlistSortColumn, setWaitlistSortColumn] = React.useState<'pos' | 'vorname' | 'nachname' | 'email' | 'jobtitle' | 'location' | 'date'>('pos');
  const [waitlistSortAsc, setWaitlistSortAsc] = React.useState(true);
  // v18.11: Sortierung der Abmeldungs-Liste (gleiche Spalten wie Teilnehmer/Warteliste).
  const [cancelledSortColumn, setCancelledSortColumn] = React.useState<'vorname' | 'nachname' | 'email' | 'jobtitle' | 'location' | 'type' | 'date'>('date');
  const [cancelledSortAsc, setCancelledSortAsc] = React.useState(false);
  const [isReorderingIDs, setIsReorderingIDs] = React.useState(false);
  const [reorderResult, setReorderResult] = React.useState<string | null>(null);
  // v18.70: Manueller Nachrück-Button (freien Platz mit erstem Wartelistler füllen)
  const [isPromoting, setIsPromoting] = React.useState(false);
  const [promoteResult, setPromoteResult] = React.useState<string | null>(null);
  // v28.70: Wartelisten-Platz einer Person manuell setzen (z.B. auf 1).
  // v28.74: Hover-/Fokus-Reiter im Event-Details-Tabstrip (Inline-Styles
  // können kein :hover).
  const [evTabHover, setEvTabHover] = React.useState<string | null>(null);
  // v28.79: Sichtbarkeits-Liste in der „Naechste Schritte"-Box eingeklappt.
  const [visListOpen, setVisListOpen] = React.useState<boolean>(false);
  const [wlPosModal, setWlPosModal] = React.useState<{ reg: SPRegistration; currentPos: number; total: number } | null>(null);
  const [wlPosValue, setWlPosValue] = React.useState<string>('1');
  const [wlPosBusy, setWlPosBusy] = React.useState(false);
  const [isResettingCounter, setIsResettingCounter] = React.useState(false);
  const [resetCounterResult, setResetCounterResult] = React.useState<string | null>(null);
  const [isFixingColumns, setIsFixingColumns] = React.useState(false);
  const [fixColumnsResult, setFixColumnsResult] = React.useState<string | null>(null);
  // v11.36: Überbuchungs-Bereinigung
  const [isDetectingOverbook, setIsDetectingOverbook] = React.useState(false);
  const [detectOverbookResult, setDetectOverbookResult] = React.useState<string | null>(null);
  // Modal-State für „Bestätigen" (einzeln oder Sammel) bzw. „Platz behalten".
  // mode: 'confirm' = auf Warteliste; 'keep' = Platz behalten.
  // targets: betroffene Registrierungen (1 = einzeln, n = Sammel „Alle").
  const [overbookModal, setOverbookModal] = React.useState<{
    mode: 'confirm' | 'keep';
    targets: SPRegistration[];
  } | null>(null);
  const [obWithMail, setObWithMail] = React.useState(true);
  const [obMailSubject, setObMailSubject] = React.useState('');
  const [obMailBody, setObMailBody] = React.useState('');
  const [obMailLang, setObMailLang] = React.useState<'DE' | 'EN'>('DE');
  const [obRemoveCalendar, setObRemoveCalendar] = React.useState(true);
  const [obKeepVariant, setObKeepVariant] = React.useState<'active' | 'firstWaitlist'>('firstWaitlist');
  const [obBusy, setObBusy] = React.useState(false);
  // v11.36: Fortschritts-Overlay für die ID-Neuvergabe (0..100 %, null = aus).
  const [reorderProgress, setReorderProgress] = React.useState<number | null>(null);
  const [reorderProgressLabel, setReorderProgressLabel] = React.useState('');
  const [isFixingFields, setIsFixingFields] = React.useState(false);
  const [fixFieldsResult, setFixFieldsResult] = React.useState<string | null>(null);
  // v11.84: Teams-Section (Admin Center Team-Management).
  const [teamsCollapsed, setTeamsCollapsed] = React.useState<boolean>(false);
  // Add-Member-Modal pro Team — gleiche Mechanik wie MyEventsPage.
  const [adminAddMemberDialog, setAdminAddMemberDialog] = React.useState<{
    teamId: string;
    teamName: string;
    freeSlots: number;
    /** v17.1: true wenn dieser Dialog im „Neues Team anlegen"-Flow geöffnet
     *  wurde — dann zeigen wir ein optionales Team-Name-Eingabefeld
     *  und übernehmen den eingegebenen Namen beim Insert. */
    isNewTeam?: boolean;
  } | null>(null);
  const [adminAddMemberPick, setAdminAddMemberPick] = React.useState<{ email: string; displayName: string } | null>(null);
  const [adminAddMemberQuery, setAdminAddMemberQuery] = React.useState('');
  const [adminAddMemberResults, setAdminAddMemberResults] = React.useState<Array<{ email: string; displayName: string }>>([]);
  const [adminAddMemberSearching, setAdminAddMemberSearching] = React.useState(false);
  const [adminAddMemberConsent, setAdminAddMemberConsent] = React.useState(false);
  // v17.4: Multi-Select aus teamlosen Personen + Lead-Auswahl + Mail-Opt-in.
  const [adminAddTeamlessPicks, setAdminAddTeamlessPicks] = React.useState<Set<number>>(new Set());
  const [adminAddLeadRegId, setAdminAddLeadRegId] = React.useState<number | null>(null);
  const [adminAddSendMail, setAdminAddSendMail] = React.useState<boolean>(false);
  // v22.42: Organizer kann die Bestätigungs-/Info-Mail der Team-Zuordnung
  // optional als Kopie (CC) an sich selbst bekommen.
  const [adminAddCcOrganizer, setAdminAddCcOrganizer] = React.useState<boolean>(false);
  // v22.49: Kommunikation an die ÜBRIGEN Team-Mitglieder (optional) — Reichweite
  // alle vs. nur Lead. Plus: bei ganz neuer Person ist die Anmeldebestätigung
  // (+ Outlook) an die Person ebenfalls optional (Default an).
  const [adminAddNotifyOthers, setAdminAddNotifyOthers] = React.useState<boolean>(false);
  const [adminAddNotifyScope, setAdminAddNotifyScope] = React.useState<'all' | 'lead'>('all');
  const [adminAddNewPersonMail, setAdminAddNewPersonMail] = React.useState<boolean>(true);
  const [adminAddMemberBusy, setAdminAddMemberBusy] = React.useState(false);
  const [adminAddMemberError, setAdminAddMemberError] = React.useState('');
  const [adminAddMemberIncludeIntl, setAdminAddMemberIncludeIntl] = React.useState(false);
  const adminAddMemberQueryTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  React.useEffect(() => {
    const q = adminAddMemberQuery.trim();
    if (q.length >= 2 && !adminAddMemberPick) {
      (async () => {
        setAdminAddMemberSearching(true);
        try {
          const res = await searchUsers(q, adminAddMemberIncludeIntl);
          setAdminAddMemberResults(res.map(r => ({ email: r.email, displayName: r.displayName })));
        } catch { setAdminAddMemberResults([]); }
        setAdminAddMemberSearching(false);
      })().catch(() => {});
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [adminAddMemberIncludeIntl]);
  // Lead-Transfer-Dropdown: pro Team ein offener Dropdown-Index (TeamId-Key).
  const [leadTransferOpenFor, setLeadTransferOpenFor] = React.useState<string | null>(null);
  // v22.45: Pro-Team „Anpassen"-Modus — erst dann erscheinen die
  // „Entfernen"-Buttons pro Mitglied (nicht dauerhaft an jedem Namen).
  const [teamEditOpenFor, setTeamEditOpenFor] = React.useState<string | null>(null);
  const [leadTransferBusy, setLeadTransferBusy] = React.useState(false);
  // v23.0: Drag&Drop-Zuordnung in der Teams-Sektion. dragRegId = gezogene
  // Registrierung, dragOverTid = aktuelles Drop-Ziel ('' = „ohne Team").
  const [dragRegId, setDragRegId] = React.useState<number | null>(null);
  const [dragOverTid, setDragOverTid] = React.useState<string | null>(null);
  // v23.0: Per-Team-Info-Mail (z.B. Teams-Einwahllink je Break-Out-Session).
  const [teamMailOpen, setTeamMailOpen] = React.useState(false);
  const [teamMailSubject, setTeamMailSubject] = React.useState('');
  const [teamMailBody, setTeamMailBody] = React.useState('');
  const [teamMailInfoByTid, setTeamMailInfoByTid] = React.useState<Record<string, string>>({});
  const [teamMailSending, setTeamMailSending] = React.useState(false);
  // Toast nach erfolgreicher Aktion in der Teams-Section.
  const [teamsToast, setTeamsToast] = React.useState<string>('');
  const [isRefreshingProfiles, setIsRefreshingProfiles] = React.useState(false);
  const [refreshProfilesResult, setRefreshProfilesResult] = React.useState<string | null>(null);
  // v30.37: Berechtigungs-Reparatur über Klammer + alle Termine.
  const [isRepairingPerms, setIsRepairingPerms] = React.useState(false);
  const [repairPermsResult, setRepairPermsResult] = React.useState<string | null>(null);
  // Globale Reparatur: Organizer-Email-Mismatch über alle Events fixen
  const [isRepairingOrganizers, setIsRepairingOrganizers] = React.useState(false);
  const [repairOrganizersResult, setRepairOrganizersResult] = React.useState<string | null>(null);
  // v28.65: Reparatur der Claims-Login-Tokens („0#.f|membership|…") in Namen —
  // Teilnehmerzeilen, Organizer-Namen der Events und die Rollenliste.
  const [isRepairingNames, setIsRepairingNames] = React.useState(false);
  const [repairNamesResult, setRepairNamesResult] = React.useState<string | null>(null);
  const [nameFixModal, setNameFixModal] = React.useState<{
    running: boolean; step: string; evIdx: number; evTotal: number; summary: string[] | null;
  } | null>(null);
  // v20.6: Reparatur "Fremd-Anmeldungen: Zugriff" über alle aktiven Events —
  // prüft pro Teilnehmerliste die "nur eigene Elemente"-Sicherheit und setzt
  // bei Anmeldungen durch Dritte den Zeilen-Autor auf den Teilnehmer.
  const [isRepairingAccess, setIsRepairingAccess] = React.useState(false);
  const [repairAccessResult, setRepairAccessResult] = React.useState<string | null>(null);
  // v20.7: Fortschritts-Modal für die Zugriffs-Reparatur (Event i/N +
  // Eintrag x/y + Abschluss-Summary). running=false ⇒ Summary + Schließen.
  const [accessFixModal, setAccessFixModal] = React.useState<{
    running: boolean;
    evIdx: number; evTotal: number; evTitle: string;
    itemDone: number; itemTotal: number;
    summary: string[] | null;
  } | null>(null);
  // Email Compose Modal
  const [showEmailModal, setShowEmailModal] = React.useState(false);
  // v17.10: Massmail-Target-Picker. Erst Zielgruppe wählen, dann den
  // RichText-Editor öffnen. Mode = 'closed' | 'pick' | 'paste' | 'editor'.
  const [massmailMode, setMassmailMode] = React.useState<'closed' | 'pick' | 'paste' | 'editor'>('closed');
  const [massmailAudience, setMassmailAudience] = React.useState<MassmailAudience>('active');
  // v22.9: Eigene Status-Auswahl ('custom') — welche Status die Mail bekommen.
  const [massmailStatuses, setMassmailStatuses] = React.useState<Set<string>>(new Set(['Angemeldet', 'QR versendet', 'Eingecheckt']));
  // Für 'nachruecker': der eingefügte Rohtext + die nach Extraktion
  // verbleibenden Teilnehmer (= angemeldete Personen, die NICHT in der
  // eingefügten Liste stehen).
  const [massmailPasteRaw, setMassmailPasteRaw] = React.useState<string>('');
  const [emailSubject, setEmailSubject] = React.useState('');
  const [emailHeading, setEmailHeading] = React.useState('');
  const [emailBody, setEmailBody] = React.useState('');
  const [emailSending, setEmailSending] = React.useState(false);
  // v30.51: Frei wählbares CC für Massenmail und Einladungsmail. Die Organizer
  // stehen weiterhin AUTOMATISCH auf CC — das hier kommt zusätzlich dazu und
  // ersetzt es nicht: Wer eine Mail an sein Event schickt, soll sie immer auch
  // selbst im Postfach haben, unabhängig davon, was er hier einträgt.
  const [massmailCc, setMassmailCc] = React.useState<string[]>([]);
  const [inviteCc, setInviteCc] = React.useState<string[]>([]);
  // v22.9: Massenmail-Entwurf pro Event speichern (wie Einladungsmail) +
  // Testmail an die Organizer.
  const [massmailDraftSaved, setMassmailDraftSaved] = React.useState(false);
  const massmailHydratingRef = React.useRef(false);
  const [massmailTesting, setMassmailTesting] = React.useState(false);
  const [massmailTestMsg, setMassmailTestMsg] = React.useState<string | null>(null);
  // v22.11: Unter-Überschrift der Massenmail editierbar (Parität zur
  // Einladungsmail; vorher war das Feld sichtbar, aber nicht angebunden).
  const [massmailSubheading, setMassmailSubheading] = React.useState('');
  // v26.78: Bild im Mail-Kopf (Hero) der Massenmail wählbar — 'logo' = Standard
  // (DEX-Logo/Orb bzw. konfiguriertes Mail-Logo) oder 'event' = das Event-Foto.
  // Bei 'event' wird das Event-Bild client-seitig als Base64 (getCachedImage)
  // direkt in den {{ORB_URL}}-Platzhalter gebacken, damit es unabhängig vom
  // Flow-Default im Mail-Kopf erscheint. massmailEventPhotoB64 wird beim Öffnen
  // des Massenmail-Editors vorgeladen (für Vorschau + Versand).
  // v30.52: Auswahl UND Maße des Kopf-Bildes liegen jetzt in EINEM Objekt
  // (utils/mailHeaderImage) — vorher waren es je Mail-Dialog zwei getrennte
  // States, und die QR-Mail hatte gar keine.
  const [massmailHeaderImage, setMassmailHeaderImage] = React.useState<MailHeaderImage>(MAIL_HEADER_IMAGE_DEFAULT);
  const [massmailEventPhotoB64, setMassmailEventPhotoB64] = React.useState<string>('');
  // v26.88: dieselbe „Bild im Mail-Kopf"-Wahl für die EINLADUNGSMAIL.
  const [inviteHeaderImage, setInviteHeaderImage] = React.useState<MailHeaderImage>(MAIL_HEADER_IMAGE_DEFAULT);
  // v26.98: Event-Foto im Mail-Composer zuschneiden (invite/massmail).
  const [composerCrop, setComposerCrop] = React.useState<'invite' | 'massmail' | 'qr' | null>(null);
  const [inviteEventPhotoB64, setInviteEventPhotoB64] = React.useState<string>('');
  // v11.40: Einladungsmail-Modal — Mail mit Anmelde-Link an Organizer (zum
  // Weiterleiten) oder direkt an den hinterlegten Mailverteiler des Events.
  const [showInviteModal, setShowInviteModal] = React.useState(false);
  const [inviteSubject, setInviteSubject] = React.useState('');
  const [inviteHeading, setInviteHeading] = React.useState('');
  const [inviteBody, setInviteBody] = React.useState('');
  const [inviteTarget, setInviteTarget] = React.useState<'organizer' | 'audience' | 'pending' | 'uninvited'>('organizer');
  // v28.37: Der Mailverteiler kann mehrere hundert Adressen haben — im Dialog
  // stand die komplette Liste als Fliesstext und schob alles andere nach unten.
  // Jetzt eingeklappt mit Aufklapp-Knopf.
  const [inviteAudienceOpen, setInviteAudienceOpen] = React.useState(false);
  // v28.37: Adressen, die für dieses Event schon eine Einladungsmail bekommen
  // haben (aus DEX_Emails, Typ „Einladung"). Wird beim Öffnen des Dialogs
  // geladen; null = noch nicht geladen bzw. nicht ermittelbar.
  // v30.67 (Review): undefined = wird geladen, null = Lesen gescheitert
  // (Abgleich nicht möglich), Set = bekannt. Beide „unbekannt"-Zustände
  // sperren den Modus „Nur an noch nicht Eingeladene" — sonst ging die
  // Einladung bei einem 429 auf DEX_Emails ein zweites Mal an alle.
  const [invitedLc, setInvitedLc] = React.useState<Set<string> | null | undefined>(undefined);
  // v28.37: Vom Organizer von Hand angepasste Empfaengerliste. null = keine
  // Anpassung, es gilt die per Radio gewählte Liste. Sobald etwas entfernt
  // oder ergänzt wird, übernimmt diese Liste.
  const [inviteCustomEmails, setInviteCustomEmails] = React.useState<string[] | null>(null);
  const [inviteAddInput, setInviteAddInput] = React.useState('');
  const [inviteSending, setInviteSending] = React.useState(false);
  // v26.94: Header-Bild-Größe (Breite/Innenabstand) auch in Einladungs- und
  // Massenmail einstellbar — gleiche Steuerung wie im Wizard (inkl. „Volle
  // Breite"). Transient pro Session.
  // v30.52: Die beiden Layout-States sind in `massmailHeaderImage` /
  // `inviteHeaderImage` aufgegangen (oben) — Auswahl und Maße gehören
  // zusammen, sonst muss jede Aufrufstelle zwei Dinge einsammeln.
  // v22.5: Unter-Überschrift der Einladungsmail (vorher nicht erfasst) + Entwurf-
  // Speicherung pro Event in localStorage, damit ein angefangener Text beim
  // Schließen + erneuten Öffnen erhalten bleibt.
  const [inviteSubheading, setInviteSubheading] = React.useState('');
  // verhindert, dass das Auto-Speichern den gerade geladenen Entwurf sofort
  // wieder überschreibt, bevor der State gesetzt ist.
  const inviteHydratingRef = React.useRef(false);
  // v22.5: kurzes „Gespeichert"-Feedback nach Klick auf den Speichern-Button.
  const [inviteDraftSaved, setInviteDraftSaved] = React.useState(false);
  // Gesendete-Rundmails-Viewer: liest den durablen Kommunikations-Log
  // (DEX_EventComms) für das aktuell gewählte Event, neueste zuerst.
  const [showCommsModal, setShowCommsModal] = React.useState(false);
  const [commsLoading, setCommsLoading] = React.useState(false);
  const [commsRows, setCommsRows] = React.useState<EventCommRow[]>([]);
  const [commsExpandedId, setCommsExpandedId] = React.useState<number | null>(null);
  const openCommsModal = (): void => {
    if (!eventServiceRef || !selectedEvent) return;
    setShowCommsModal(true);
    setCommsExpandedId(null);
    setCommsLoading(true);
    setCommsRows([]);
    eventServiceRef.getEventComms(selectedEvent.id)
      .then(rows => { setCommsRows(rows); })
      .catch(() => { setCommsRows([]); })
      .then(() => { setCommsLoading(false); }, () => { setCommsLoading(false); });
  };
  // v26.69: laufendes Löschen einer Log-Zeile (Id) im Kommunikations-Log.
  const [commsDeletingId, setCommsDeletingId] = React.useState<number | null>(null);
  const deleteCommRow = async (row: EventCommRow): Promise<void> => {
    if (!eventServiceRef || !selectedEvent) return;
    setCommsDeletingId(row.id);
    const ok = await eventServiceRef.deleteEventComm(row.id).catch(() => false);
    setCommsDeletingId(null);
    if (!ok) {
      showAlert(isDe ? 'Eintrag konnte nicht gelöscht werden.' : 'The entry could not be deleted.', { variant: 'error' });
      return;
    }
    setCommsRows(prev => prev.filter(r => r.id !== row.id));
    showAlert(isDe ? 'Eintrag aus dem Kommunikations-Log entfernt.' : 'Entry removed from the communication log.', { variant: 'success' });
  };
  const [showExportMenu, setShowExportMenu] = React.useState(false);
  // v17.12: Zielgruppen-Picker für Excel-Export.
  // v27.9: chooseMode = die Format-Auswahl (Deloitte/B2Run) wird IM Modal
  // getroffen statt im Anker-Dropdown (der im „Aktion auswählen"-Menü mit
  // overflow:auto abgeschnitten wurde → Auswahl war unsichtbar).
  const [excelTargetModal, setExcelTargetModal] = React.useState<null | { mode: 'deloitte' | 'b2run'; chooseMode?: boolean }>(null);
  // v30.48: Rücklauf des Veranstalters mit den echten Startnummern einlesen.
  const [bibImportOpen, setBibImportOpen] = React.useState(false);
  // v30.54: Offene Ummeldungen beim Veranstalter — live aus der Liste.
  const [b2runTodoOpen, setB2runTodoOpen] = React.useState(false);
  // v30.60: Bestellliste der Trikots (s. components/admin/ShirtSizeModal).
  const [shirtSizeOpen, setShirtSizeOpen] = React.useState(false);
  // v30.60: Aufgeklappte Reiter-Gruppe („Day 1" …). null = die zuletzt
  // sinnvolle Gruppe wird beim Rendern bestimmt (die des gewählten Termins).
  const [openTabGroup, setOpenTabGroup] = React.useState<string | null>(null);
  const [excelAudience, setExcelAudience] = React.useState<'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled'>('active');
  // v20.4: Excel-Export im Klammer-Modus — konsolidierte Matrix (eine Zeile
  // pro Person, Spalten pro Sub-Event) und/oder einzelne Sub-Event-Blätter
  // sind im Export-Modal wählbar.
  const [excelIncludeMatrix, setExcelIncludeMatrix] = React.useState(true);
  const [excelSubIds, setExcelSubIds] = React.useState<Set<string>>(new Set());
  // Outlook-Decline-Check (Admin only): zeigt Teilnehmer, die in Outlook
  // abgesagt haben, aber in der Teilnehmerliste noch aktiv gelistet sind.
  const [isCheckingDeclines, setIsCheckingDeclines] = React.useState(false);
  const [declineResult, setDeclineResult] = React.useState<{
    declinedAndRegistered: Array<{ email: string; name: string; reg: SPRegistration }>;
    declinedTotal: number;
    error: string | null;
  } | null>(null);
  const [showDeclineModal, setShowDeclineModal] = React.useState(false);
  const [declineCopied, setDeclineCopied] = React.useState(false);


  // v8.0: In-App-Edit-Modal für Teilnehmer (Admin/Organizer kann jeden
  // Teilnehmer-Eintrag direkt aus der Liste editieren — Anrede, Name, Email,
  // Phone, Department, Location, JobTitle, Status, plus alle Custom-Felder).
  // Beim Save wird eine Audit-Zeile in ChangeLog geschrieben (wer/wann/was)
  // und LastModifiedDate gesetzt — kein direkter SP-Edit mehr nötig, was
  // gleichzeitig das deutsche Datumsformat-Problem in SP umgeht.
  const [editingReg, setEditingReg] = React.useState<SPRegistration | null>(null);
  const [editForm, setEditForm] = React.useState<Record<string, string>>({});
  const [isSavingEdit, setIsSavingEdit] = React.useState(false);
  const [editError, setEditError] = React.useState('');
  // v11.0: Attachment-Modal pro Teilnehmer-Reihe — nur wenn das Event
  // den Teilnehmer-Upload erlaubt (event.allowAttendeeUpload). Modal
  // listet alle Item-Attachments der jeweiligen Reg-Zeile + bietet
  // Download-Link + Lösch-Button (Admin/Organizer kann auch fremde
  // Uploads löschen). Map(regId → attachments) wird beim Laden der
  // Teilnehmerliste einmalig befüllt, damit der Anhang-Button die
  // Anzahl direkt anzeigen kann.
  const [attachmentsByReg, setAttachmentsByReg] = React.useState<Record<number, Array<{ fileName: string; serverRelativeUrl: string }>>>({});
  const [attachmentsModalReg, setAttachmentsModalReg] = React.useState<SPRegistration | null>(null);
  const [attachmentsBusy, setAttachmentsBusy] = React.useState(false);
  // v19.30 — Feature A: Im konsolidierten View (Hauptevent mit Sub-Events) die
  // Custom-Felder des Hauptevents („Felder des Hauptevents") pro Teilnehmer
  // editierbar machen. Die Antworten leben in der Registrierung der Person auf
  // der Hauptevent-Subsite (`selectedEvent.subsiteUrl`). `mainFieldsEditReg`
  // hält diese Parent-Registrierung; `mainFieldsEditForm` die editierten Werte
  // (Field-ID → String). `mainFieldsEditName` nur zur Anzeige im Modal-Titel.
  const [mainFieldsEditReg, setMainFieldsEditReg] = React.useState<SPRegistration | null>(null);
  const [mainFieldsEditName, setMainFieldsEditName] = React.useState('');
  const [mainFieldsEditForm, setMainFieldsEditForm] = React.useState<Record<string, string>>({});
  const [mainFieldsEditSaving, setMainFieldsEditSaving] = React.useState(false);
  const [mainFieldsEditError, setMainFieldsEditError] = React.useState('');
  // v24.37: Speicherziel der Klammer-/Hauptevent-Felder. Normalfall = die
  // Hauptevent-Teilnehmerliste (`selectedEvent.subsiteUrl`, isParent=true). Hat
  // die Person KEINE Hauptevent-Anmeldung (nur Sub-Events, „falsche"/
  // Schatten-Anmeldungen im Klammer-Modus), schreiben wir die Klammer-Antworten
  // in die CustomData ihrer frühesten aktiven Sub-Event-Zeile (isParent=false)
  // — der konsolidierte View liest Klammer-Felder als Fallback genau von dort.
  const [mainFieldsEditSubsite, setMainFieldsEditSubsite] = React.useState('');
  const [mainFieldsEditTargetIsParent, setMainFieldsEditTargetIsParent] = React.useState(true);
  // v19.30 — Feature B: Abmeldung eines Teilnehmers aus dem konsolidierten
  // View mit Sub-Event-Auswahl. Der Modal listet alle Sub-Events, für die die
  // Person aktiv angemeldet ist (Status angemeldet/QR/eingecheckt/Warteliste),
  // je mit Checkbox. `deregModal` hält die betroffene Person + die abmeldbaren
  // Sub-Event-Registrierungen; `deregSelected` die angehakten Sub-Event-IDs.
  // v29.29: `isParent` markiert die Klammer-Zeile — sie war bisher gar nicht
  // Teil der Auswahl, dadurch blieb die Person nach dem Abmelden aller
  // Sub-Events auf der Klammer angemeldet („Teilnehmer (1)").
  const [deregModal, setDeregModal] = React.useState<{
    emailKey: string;
    name: string;
    email: string;
    items: Array<{ child: DeloitteEvent; reg: SPRegistration; isParent?: boolean }>;
  } | null>(null);
  const [deregSelected, setDeregSelected] = React.useState<Set<string>>(new Set());
  const [deregBusy, setDeregBusy] = React.useState(false);
  // v29.29: Stille Abmeldung — keine Abmelde-Mail, keine Outlook-Absage.
  // Der Auslöser dafür ist der Regelfall „Person hat das Unternehmen
  // verlassen": Das Postfach existiert nicht mehr, die Mail liefe ins Leere
  // (oder als Bounce zurück an das Sammelpostfach). Vorausgewählt, wenn der
  // Konto-Check die Adresse als inaktiv gemeldet hat.
  const [deregSilent, setDeregSilent] = React.useState(false);
  // v30.66: useEditModalHandlers — Rumpf in logic/useEditModalHandlers.ts.
  const {
    closeEditModal, openEditModal, saveEdit,
  } = useEditModalHandlers({
    currentUser, editForm, editingReg, eventServiceRef, isDe, reloadRegistrations, searchUser,
    selectedEvent, setEditError, setEditForm, setEditingReg, setIsSavingEdit,
  });


  // v19.30 — Feature D: Audit-Log vorgefiltert auf das aktuell selektierte
  // Event öffnen (setzt den Event-/Ziel-Filter auf den Event-Titel).
  const openChangeLogForEvent = (): void => {
    setChangeLogFilterEvent(selectedEvent?.title || '');
    setChangeLogFilterAction('');
    setChangeLogFilterActor('');
    void openChangeLog();
  };

  // v30.66: useWaitlistActions — Rumpf in logic/useWaitlistActions.ts.
  const {
    idFixCheckedForRef, idRecheckBusy, recentCancellation, reloadRegistrationsForIdCheck,
    runIdReorder, runManualPromote, runOverbookResolution,
  } = useWaitlistActions({
    allEvents, confirmDialog, eventServiceRef, getAllRegistrations, isDe, isSplitCapacity,
    obKeepVariant, obMailBody, obMailLang, obMailSubject, obRemoveCalendar, obWithMail,
    overbookModal, registrations, reloadRegistrations, selectedEvent, setAdminToast, setIsPromoting, setIsReorderingIDs,
    setObBusy, setObMailBody, setObMailLang, setObMailSubject, setOverbookModal, setPromoteResult,
    setReorderProgress, setReorderProgressLabel, setReorderResult,
  });

  // v30.66: useTeamActions — Rumpf in logic/useTeamActions.tsx.
  const {
    columnOrder, getActiveTeams, hiddenColumns, moveRegToTeam, onTeamDrop, openTeamMailDialog,
    sendTeamMails, setColumnOrder, setHiddenColumns, setShowColumnPicker, setShowMatches,
    showColumnPicker, showMatches,
  } = useTeamActions({
    dragRegId, eventServiceRef, idFixCheckedForRef, isDe, isLoadingRegs,
    recentCancellation, registrations, reloadRegistrations, reloadRegistrationsForIdCheck, selectedEvent,
    setDragOverTid, setDragRegId, setTeamMailBody, setTeamMailInfoByTid,
    setTeamMailOpen, setTeamMailSending, setTeamMailSubject, showAlert, teamMailBody,
    teamMailInfoByTid, teamMailSubject,
  });


  // v22.7: Hintergrund-Check beim Öffnen eines Events — sind die E-Mail-Adressen
  // der Teilnehmer noch zu einem aktiven Deloitte-Konto? Ergebnis wird pro Event
  // max. 1×/Tag in localStorage gecacht (kein Graph-Call bei jedem Öffnen).
  React.useEffect(() => {
    setInactiveAccounts([]);
    if (!selectedEvent || !eventServiceRef || registrations.length === 0) return undefined;
    const emails = Array.from(new Set(registrations
      .filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt' || r.Status === 'Warteliste')
      .map(r => (r.ParticipantEmail || '').trim().toLowerCase())
      .filter(Boolean)));
    if (emails.length === 0) return undefined;
    // v26.42: _v2 — alte Caches enthielten Fehlalarme für umbenannte Konten (Heirat).
    // v29.31: Schlüssel zentral (utils/accountCheckCache) — die Startseiten-Box
    // und die Invalidierung nach einer Abmeldung müssen denselben treffen.
    const cacheKey = accountCheckCacheKey(selectedEvent.id);
    try {
      const raw = window.localStorage.getItem(cacheKey);
      if (raw) {
        const parsed = JSON.parse(raw) as { ts?: number; inactive?: string[]; checked?: string[] };
        const fresh = !!parsed && typeof parsed.ts === 'number' && (Date.now() - parsed.ts) < 24 * 60 * 60 * 1000 && Array.isArray(parsed.inactive);
        // v22.43: Cache nur nutzen, wenn ALLE aktuellen Adressen schon geprüft
        // wurden. Sind seit dem letzten Lauf neue Teilnehmer dazugekommen
        // (Adresse nicht in `checked`), wird frisch geprüft — sonst blieben
        // neu hinzugefügte Personen bis zu 24h ungeprüft (Bug v22.7–v22.42).
        const checked = Array.isArray(parsed?.checked) ? (parsed!.checked as string[]) : [];
        const coversAll = checked.length > 0 && emails.every(e => checked.indexOf(e) >= 0);
        if (fresh && coversAll) {
          setInactiveAccounts((parsed!.inactive || []).filter(e => emails.indexOf(e) >= 0));
          return undefined;
        }
      }
    } catch { /* localStorage evtl. blockiert */ }
    let cancelled = false;
    (async () => {
      try {
        const res = await eventServiceRef.checkAccountsActive(emails);
        if (cancelled || !res.ok) return;
        setInactiveAccounts(res.inactive);
        try { window.localStorage.setItem(cacheKey, JSON.stringify({ ts: Date.now(), inactive: res.inactive, checked: emails })); } catch { /* */ }
      } catch { /* best-effort */ }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, registrations.length, eventServiceRef]);

  // SuperAdmin sieht alle Events, EventAdmin nur seine + QR-Scanner-Events.
  // Zugriff wird strikt per E-Mail geprüft — NICHT per Namens-Substring
  // (hatte mehrere Jahre einen Match-per-Surname-Bug, der bei häufigen Nachnamen
  // zu False-Positives führte: z.B. eine Assistentin "Frau Müller" konnte Events
  // sehen, deren Organizer auf "Max Müller" hieß — weil "müller" in "max müller"
  // vorkommt. Seit v6.20 nur noch exakt per currentUser.email gegen
  // event.organizerEmails bzw. event.qrScannerEmails.)
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isQRScannerFor = (ev: DeloitteEvent): boolean =>
    !!currentEmailLc && !!ev.qrScannerEmails && ev.qrScannerEmails.some(e => e.toLowerCase() === currentEmailLc);
  // v9.18: Co-Organizer haben pro Event die gleichen Rechte wie der Hauptorganizer.
  // isOrganizerFor returned true sowohl für event.organizerEmails als auch
  // für event.coOrganizerEmails (per-Event-Rolle).
  const isOrganizerFor = (ev: DeloitteEvent): boolean => {
    // v18.3: Im Demo-Modus ist der (User-)Demo-Account „Organizer" des
    // synthetischen Demo-Events — so sieht er die Teilnehmer-Verwaltung
    // (read-only) im Admin-Center. Greift nur für das Demo-Event.
    if (isImpersonating && ev.isDemoShowcase) return true;
    if (!currentEmailLc) return false;
    if (ev.organizerEmails && ev.organizerEmails.some(e => e.toLowerCase() === currentEmailLc)) return true;
    if (ev.coOrganizerEmails && ev.coOrganizerEmails.some(e => (e || '').toLowerCase() === currentEmailLc)) return true;
    // v22.14: Organizer des HAUPTEVENTS gelten auch auf dessen Sub-Events als
    // Organizer. Vorher waren Sub-Event-Tabs für Parent-Organizer beschnitten
    // (Status-Badge nicht klickbar, Organizer-Aktionen ausgeblendet), wenn die
    // Organizer-Liste des Kindes nicht (mehr) identisch gepflegt war.
    if (ev.parentEventId) {
      const parent = allEvents.find(p => p.id === ev.parentEventId);
      if (parent) {
        if (parent.organizerEmails && parent.organizerEmails.some(e => (e || '').toLowerCase() === currentEmailLc)) return true;
        if (parent.coOrganizerEmails && parent.coOrganizerEmails.some(e => (e || '').toLowerCase() === currentEmailLc)) return true;
      }
    }
    return false;
  };
  const adminEvents = isAdmin
    ? events
    : events.filter(e => isOrganizerFor(e) || isQRScannerFor(e));
  // Wenn der User NUR QR-Scanner ist (nicht Organizer + nicht Admin), dann läuft die
  // Admin-Page im eingeschränkten Modus für das ausgewählte Event: nur KPI-Kacheln
  // + QR-Code-Scanner-Button sichtbar.
  const isQRScannerOnlyForSelected = !!selectedEvent && !isAdmin && !isOrganizerFor(selectedEvent) && isQRScannerFor(selectedEvent);

  // Für Admins: vergangene Events in eine einklappbare Sektion auslagern
  // (Organizer sehen nur ihre eigenen Events — dort bleiben auch abgelaufene
  // sichtbar, weil der Organizer sie für den Abschluss / CSV-Export etc.
  // evtl. direkt griffbereit braucht).
  const now = Date.now();
  const isPastEvent = (e: DeloitteEvent): boolean =>
    !!e.endDate && new Date(e.endDate).getTime() < now;
  // v24.8 (O): abgeschlossene/vergangene Events sind für Organizer gesperrt —
  // keine Abmeldung/Löschung/Feld-Bearbeitung mehr (Archivierungsschutz; nur
  // No-Show über den Check-in bleibt). Admins behalten vollen Zugriff.
  const orgPastLock = !isAdmin && !!selectedEvent && isEventOver(selectedEvent);
  // v26.18/v26.21: Duplikat-Warnung — versehentlich mehrfach angelegte Versionen.
  // Zwei Stufen, beide am GLEICHEN Tag:
  //  1) EXAKT gleicher (normalisierter) Name → immer Duplikat.
  //  2) ÄHNLICHER Name + gemeinsamer Organizer → Duplikat, wenn der Namens-Anfang
  //     ≥ 2 Wörter übereinstimmt ODER die Wort-Überlappung ≥ 50% ist. So werden
  //     z.B. „MD Academy 1st test" + „MD Academy 2026" erkannt, aber echte
  //     verschiedene Events („Frühlingsfest Berlin" vs „… Hamburg") NICHT.
  const duplicateEvents = React.useMemo<DeloitteEvent[]>(() => {
    if (!selectedEvent) return [];
    const norm = (s?: string): string => (s || '').trim().toLowerCase().replace(/\s+/g, ' ');
    const toks = (s?: string): string[] => norm(s).split(' ').filter(t => t.length >= 2);
    const dayKey = (d?: string): string => { if (!d) return ''; try { const dt = new Date(d); return `${dt.getFullYear()}-${dt.getMonth()}-${dt.getDate()}`; } catch { return ''; } };
    const orgSet = (e: DeloitteEvent): Set<string> => new Set([...(e.organizerEmails || []), ...(e.coOrganizerEmails || [])].map(x => (x || '').trim().toLowerCase()).filter(Boolean));
    const prefixLen = (a: string[], b: string[]): number => { let i = 0; while (i < a.length && i < b.length && a[i] === b[i]) i++; return i; };
    const jaccard = (a: string[], b: string[]): number => { if (!a.length || !b.length) return 0; const sb = new Set(b); let inter = 0; new Set(a).forEach(t => { if (sb.has(t)) inter++; }); const uni = new Set([...a, ...b]).size; return uni ? inter / uni : 0; };
    const selTitle = norm(selectedEvent.title);
    if (!selTitle) return [];
    const selToks = toks(selectedEvent.title);
    const selDay = dayKey(selectedEvent.startDate);
    const selOrg = orgSet(selectedEvent);
    const shareOrg = (e: DeloitteEvent): boolean => { const o = orgSet(e); for (const x of Array.from(selOrg)) { if (o.has(x)) return true; } return false; };
    return (events || []).filter(e => {
      if (e.id === selectedEvent.id || e.parentEventId) return false;
      // v26.66 BUG-FIX: Klammer-Beziehung ist KEIN Duplikat. Sub-Events (mit
      // parentEventId) sind oben schon raus — aber wenn das SELEKTIERTE Event
      // selbst ein Sub-Event ist, hat sein Hauptevent (Parent) keinen
      // parentEventId und wurde fälschlich als Duplikat gemeldet (z. B.
      // „P/D Meeting T&T+ 2026" ⇄ dessen Sub-Event „… | DINNER"). Parent des
      // selektierten Sub-Events explizit ausschließen.
      if (selectedEvent.parentEventId && String(e.id) === String(selectedEvent.parentEventId)) return false;
      const eDay = dayKey(e.startDate);
      // gleicher Tag (falls beide ein Datum haben).
      if (selDay && eDay && eDay !== selDay) return false;
      const eTitle = norm(e.title);
      if (eTitle === selTitle) return true; // Stufe 1: exakt
      // Stufe 2: ähnlich + gemeinsamer Organizer
      if (!shareOrg(e)) return false;
      const eToks = toks(e.title);
      return prefixLen(selToks, eToks) >= 2 || jaccard(selToks, eToks) >= 0.5;
    });
  }, [events, selectedEvent]);
  const currentEventsRaw = isAdmin ? adminEvents.filter(e => !isPastEvent(e)) : adminEvents;
  const pastEventsRaw = isAdmin ? adminEvents.filter(isPastEvent) : [];
  const [showPastEvents, setShowPastEvents] = React.useState(false);
  // v24.6: Organizer-Archiv — abgelaufene Events aus DER EIGENEN Übersicht
  // ausblenden (pro Person, reiner Anzeige-Filter; Event/Daten bleiben).
  const [archivedEventIds, setArchivedEventIds] = React.useState<Set<string>>(new Set());
  const [showArchivedEvents, setShowArchivedEvents] = React.useState(false);
  const [archiveBusyId, setArchiveBusyId] = React.useState<string | null>(null);
  const archivedLoadedRef = React.useRef(false);
  React.useEffect(() => {
    if (archivedLoadedRef.current) return;
    archivedLoadedRef.current = true;
    getOrganizerArchivedEventIds().then(setArchivedEventIds).catch(() => { /* best-effort */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const handleArchiveEvent = async (event: DeloitteEvent): Promise<void> => {
    setArchiveBusyId(event.id);
    try {
      const ok = await archiveEventForOrganizer(event.id);
      if (ok) setArchivedEventIds(prev => { const n = new Set(prev); n.add(event.id); return n; });
    } catch { /* */ } finally { setArchiveBusyId(null); }
  };
  const handleUnarchiveEvent = async (event: DeloitteEvent): Promise<void> => {
    setArchiveBusyId(event.id);
    try {
      const ok = await unarchiveEventForOrganizer(event.id);
      if (ok) setArchivedEventIds(prev => { const n = new Set(prev); n.delete(event.id); return n; });
    } catch { /* */ } finally { setArchiveBusyId(null); }
  };
  const archivedCount = adminEvents.filter(e => archivedEventIds.has(e.id)).length;
  // v18.2: Entwurf-Filter + Sortierung der Admin/Organizer-Event-Liste.
  // Default-Sortierung alphabetisch nach Titel; alternativ nach Startdatum
  // aufsteigend. „Entwürfe ausblenden" filtert isFictive-Events raus.
  const [hideDrafts, setHideDrafts] = React.useState(false);
  const [eventSortMode, setEventSortMode] = React.useState<'alpha' | 'date'>('alpha');
  const draftCount = adminEvents.filter(e => e.isFictive).length;
  const sortAndFilterEvents = React.useCallback((list: DeloitteEvent[]): DeloitteEvent[] => {
    let arr = list.slice();
    if (hideDrafts) arr = arr.filter(e => !e.isFictive);
    // v24.6: archivierte (für mich ausgeblendete) Events nur zeigen, wenn der
    // „Archivierte anzeigen"-Schalter an ist.
    if (!showArchivedEvents) arr = arr.filter(e => !archivedEventIds.has(e.id));
    arr.sort((a, b) => {
      if (eventSortMode === 'date') {
        const am = a.startDate ? new Date(a.startDate).getTime() : Number.POSITIVE_INFINITY;
        const bm = b.startDate ? new Date(b.startDate).getTime() : Number.POSITIVE_INFINITY;
        if (am !== bm) return am - bm;
        return (a.title || '').localeCompare(b.title || '', isDe ? 'de' : 'en');
      }
      return (a.title || '').localeCompare(b.title || '', isDe ? 'de' : 'en');
    });
    return arr;
  }, [hideDrafts, eventSortMode, isDe, showArchivedEvents, archivedEventIds]);
  const currentEvents = sortAndFilterEvents(currentEventsRaw);
  const pastEvents = sortAndFilterEvents(pastEventsRaw);

  // v30.66: useColumnConfig — Rumpf in logic/useColumnConfig.ts.
  const {
    availableColumns, hasRoommateColumn, hasWaitlistActivity, hideColumn, moveColumn, showColumn,
  } = useColumnConfig({
    allEvents, columnOrder, hiddenColumns, isDe, registrations, selectedEvent, setColumnOrder,
    setHiddenColumns,
  });


  // v30.66: useEventSelection — Rumpf in logic/useEventSelection.ts.
  const {
    handleSelectEvent, openSelfCheckInModal, sciBusy, sciFrom, sciModalOpen, sciModalQr,
    sciSaveMsg, sciTo, sciToken, setSciBusy, setSciFrom, setSciModalOpen, setSciSaveMsg, setSciTo,
    toggleDraftStatus,
  } = useEventSelection({
    adminEvents, childEventsOf, confirmDialog, detailCardRef, eventServiceRef, getAllRegistrations,
    isDe, navigate, refreshEvents, registrations, reloadRegistrations, selectedEvent, selectedEventId,
    setDeniedSubEventLists, setIsLoadingRegs, setRegLoadError, setRegistrations,
    setReservedDetailHeight, setSelectedEvent, showAlert, updateEvent,
  });

  // v30.66: useMailComposers — Rumpf in logic/useMailComposers.tsx.
  const {
    applyInviteHero, applyMassmailHero, inviteHeaderOpts, massmailHeaderOpts, openInviteModal,
    openMassmailPicker, openPendingReminder, resetInviteDraft, resetMassmailDraft,
    resolveAudienceEmails, saveInviteDraft, saveMassmailDraft, sendMassmailTestToOrganizers,
  } = useMailComposers({
    currentUser, emailBody, emailHeading, emailSubject, eventServiceRef, getGroupMembers,
    inviteBody, inviteEventPhotoB64, inviteHeaderImage, inviteHeading, inviteHydratingRef,
    inviteSubheading, inviteSubject, inviteTarget, isDe, massmailEventPhotoB64,
    massmailHeaderImage, massmailHydratingRef, massmailMode, massmailSubheading, pendingCheckBusy,
    registrations, selectedEvent, setEmailBody, setEmailHeading, setEmailSubject,
    setInviteAddInput, setInviteAudienceOpen, setInviteBody, setInviteCustomEmails,
    setInviteDraftSaved, setInviteEventPhotoB64, setInviteHeaderImage, setInviteHeading,
    setInviteSubheading, setInviteSubject, setInviteTarget, setInvitedLc, setMassmailAudience,
    setMassmailCc, setMassmailDraftSaved, setMassmailEventPhotoB64, setMassmailHeaderImage,
    setMassmailMode, setMassmailPasteRaw, setMassmailSubheading, setMassmailTestMsg,
    setMassmailTesting, setPendingCheckBusy, setPendingPeople, setShowInviteModal,
    setVisibilityResolved, showAlert, showEmailModal, showInviteModal, siteUrl,
    subEventRegsByEventId, visibilityResolved,
  });


  // v30.66: createQrMailActions — Rumpf in logic/createQrMailActions.tsx.
  const {
    closeQrMailEditor, getQrMailOverride, openQrMailEditor, qrFullSendAction, qrPreviewAction,
    qrTestSendAction, saveQrMailOverride, saveSelfCheckInWindow,
  } = createQrMailActions({
    confirmDialog, currentUser, eventServiceRef, isDe, qrBlockLang,
    qrBlockNote, qrEditBody, qrEditHeading, qrEditSaving, qrEditSubheading, qrEditSubject,
    qrEditTarget, qrHeaderImage, refreshEvents, registrations, reloadRegistrations, sciBusy, sciFrom, sciTo,
    selectedEvent, setIsSendingQR, setQrBlockLang, setQrBlockNote, setQrEditBody, setQrEditHeading,
    setQrEditOpen, setQrEditSampleBlock, setQrEditSampleImg, setQrEditSaving, setQrEditSubheading,
    setQrEditSubject, setQrEditTarget, setQrEventPhotoB64, setQrHeaderImage, setQrPreviewHtml,
    setQrPreviewLoading, setQrPreviewOpen, setQrPreviewSubject, setQrSendModalOpen,
    setQrSendResult, setQrSentCount, setSciBusy, setSciSaveMsg, setSelectedEvent,
    showAlert, updateEvent,
  });

  // Danger-Zone-Modal als gemeinsames Element — wird in BEIDEN Render-Branches
  // (Event-Liste und Event-Detail) eingehängt, sonst läuft der Löschen-Klick auf
  // der Event-Liste ins Leere (Bug v9.x: Modal war nur im Detail-Branch gerendert).
  const dangerZoneModal: React.ReactElement | null = confirmDeleteEvent ? <DangerZoneModal confirmDeleteEvent={confirmDeleteEvent} confirmDeleteText={confirmDeleteText} deleteEvent={deleteEvent} deletePolicy={deletePolicy} isDe={isDe} isDeleting={isDeleting} setConfirmDeleteEvent={setConfirmDeleteEvent} setConfirmDeleteText={setConfirmDeleteText} setDeletingId={setDeletingId} setIsDeleting={setIsDeleting} /> : null;

  // ChangeLog-/Audit-Log-Modal als gemeinsames Element — wie das Danger-Zone-
  // Modal muss auch dieses in BEIDEN Render-Branches verfügbar sein, sonst
  // öffnet sich der "Audit log"-Button auf der Event-Liste ins Leere.
  const changeLogModal: React.ReactElement | null = showChangeLogModal ? <ChangeLogModal changeLogEntries={changeLogEntries} changeLogFilterAction={changeLogFilterAction} changeLogFilterActor={changeLogFilterActor} changeLogFilterEvent={changeLogFilterEvent} changeLogHideSelf={changeLogHideSelf} changeLogLoading={changeLogLoading} isDe={isDe} setChangeLogFilterAction={setChangeLogFilterAction} setChangeLogFilterActor={setChangeLogFilterActor} setChangeLogFilterEvent={setChangeLogFilterEvent} setChangeLogHideSelf={setChangeLogHideSelf} setShowChangeLogModal={setShowChangeLogModal} /> : null;

  // v6.20: Access-Gate — wer weder Admin noch Organizer eines Events noch QR-Scanner
  // eines Events ist, darf die Admin-Seite gar nicht erst sehen. Zeigt eine klare
  // "Kein Zugriff"-Meldung statt einer leeren Event-Liste.
  if (!selectedEvent && !isAdmin && adminEvents.length === 0) {
    return (
      <div className="page-container" role="main">
        <h2 className="mb-16">{t('admin.title')}</h2>
        <div className="card" style={{ padding: 48, textAlign: 'center' }}>
          <p style={{ color: 'var(--dex-gray-700)', marginBottom: 8, fontWeight: 600 }}>
            {t('admin.noaccess.title') || 'Kein Zugriff'}
          </p>
          <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.88rem', maxWidth: 520, margin: '0 auto' }}>
            {t('admin.noaccess.msg')}
          </p>
          <div style={{ marginTop: 24, display: 'flex', gap: 12, justifyContent: 'center', flexWrap: 'wrap' }}>
            {/* v7.2: Organizer ohne eigenes Event sehen hier einen direkten
                Shortcut zum Event-Erstellen — sonst sitzen sie in dieser
                Sackgasse ohne sichtbaren nächsten Schritt. Admins sehen den
                Button nicht, weil sie bereits alle Events in adminEvents haben. */}
            {currentUserRole !== 'User' && (
              <button className="btn btn-primary" onClick={() => navigate('create-event')}>
                + {t('admin.newevent')}
              </button>
            )}
            <button className="btn btn-secondary" onClick={() => navigate('landing')}>
              {t('reg.backtoevents') || 'Zurück'}
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (!selectedEvent) {
    const eventOverviewScreenProps = {
      adminEvents, archiveBusyId, archivedCount, archivedEventIds, changeLogModal, currentEvents,
      dangerZoneModal, deletingId, draftCount, eventSortMode, handleArchiveEvent, handleSelectEvent,
      handleUnarchiveEvent, hideDrafts, isAdmin, isDe, isDeleting, isEventsLoading,
      isPastEvent, locale, navigate, pastEvents, setConfirmDeleteEvent, setConfirmDeleteText,
      setEventSortMode, setHideDrafts, setShowArchivedEvents, setShowPastEvents, showArchivedEvents, showPastEvents,
      t,
    };
    // Event-Auswahl
    return (
      <EventOverviewScreen {...eventOverviewScreenProps} />
    );
  }

  // Event ausgewählt - Detail-Ansicht
  const query = searchQuery.toLowerCase().trim();
  // v23.32: Standort ohne Länder-Präfix („DE - Köln" → „Köln").
  const stripLocPrefix = (loc: string): string => (loc || '').replace(/^[A-Za-z]{2}\s*[-–—]\s*/, '').trim();
  // v23.32: Suchtreffer im Text hervorheben (grün markiert). Gibt einen
  // React-Knoten zurück; ohne aktive Suche einfach den Text.
  const highlightMatch = (text: unknown): React.ReactNode => {
    const s = text == null ? '' : String(text);
    if (!query || !s) return s;
    const lower = s.toLowerCase();
    // v26.65: pro Such-Wort hervorheben (analog zur Token-Suche). Alle
    // Treffer-Bereiche sammeln, überlappungsfrei zusammenführen, dann rendern —
    // so wird bei „Alexander Knoth" jedes der beiden Wörter grün markiert.
    const tokens = query.split(/\s+/).filter(Boolean);
    const ranges: Array<[number, number]> = [];
    for (const t of tokens) {
      let idx = lower.indexOf(t);
      while (idx >= 0) { ranges.push([idx, idx + t.length]); idx = lower.indexOf(t, idx + t.length); }
    }
    if (ranges.length === 0) return s;
    ranges.sort((a, b) => a[0] - b[0]);
    const merged: Array<[number, number]> = [];
    for (const r of ranges) {
      const last = merged[merged.length - 1];
      if (last && r[0] <= last[1]) { last[1] = Math.max(last[1], r[1]); }
      else merged.push([r[0], r[1]]);
    }
    const parts: React.ReactNode[] = [];
    let i = 0; let k = 0;
    for (const [a, b] of merged) {
      if (a > i) parts.push(s.slice(i, a));
      parts.push(<mark key={k++} style={{ background: 'rgba(134,188,37,0.5)', color: 'inherit', padding: 0, borderRadius: 2 }}>{s.slice(a, b)}</mark>);
      i = b;
    }
    if (i < s.length) parts.push(s.slice(i));
    return <>{parts}</>;
  };
  // v23.32: durchsucht ALLE Spalten (Name, E-Mail, ID, Job Title, Standort,
  // Status und alle Custom-Field-Antworten).
  const matchesSearch = (reg: SPRegistration): boolean => {
    if (!query) return true;
    // v26.65: Token-basierte Suche über EINEN kombinierten Text (analog zur
    // konsolidierten Ansicht) — jedes getippte Wort muss irgendwo vorkommen,
    // Reihenfolge egal. Vorher war der volle Name nur in der Reihenfolge
    // „Vorname Nachname" findbar; „Knoth Alexander" oder feldübergreifende
    // Kombinationen („Alexander Berlin") gingen nicht.
    const tokens = query.split(/\s+/).filter(Boolean);
    if (tokens.length === 0) return true;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || '');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const anyReg = reg as any;
    let hay = `${name} ${reg.ParticipantEmail || ''} ${String(reg.TeilnehmerID || '')} ${String(anyReg.JobTitle || '')} ${String(anyReg.Location || '')} ${String(reg.Status || '')}`;
    if (reg.CustomData) {
      try {
        const cd = JSON.parse(reg.CustomData);
        for (const ck of Object.keys(cd)) { const v = cd[ck]; if (v != null) hay += ' ' + String(v); }
      } catch { /* */ }
    }
    hay = hay.toLowerCase();
    return tokens.every(t => hay.indexOf(t) >= 0);
  };

  const sortRegs = (a: SPRegistration, b: SPRegistration): number => {
    let cmp = 0;
    switch (sortColumn) {
      case 'id': cmp = (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0); break;
      case 'anrede': cmp = (a.Anrede || '').localeCompare(b.Anrede || '', 'de'); break;
      // v11.26: getrennte Vorname / Nachname Sortierung.
      case 'vorname': {
        const na = (a.Vorname || (a.ParticipantName || '').split(' ')[0] || '');
        const nb = (b.Vorname || (b.ParticipantName || '').split(' ')[0] || '');
        cmp = na.localeCompare(nb, 'de');
        break;
      }
      case 'nachname': {
        // Fallback für Alt-Daten ohne separates Vorname/Nachname:
        // Letztes Wort aus ParticipantName als Nachname.
        const lastWord = (s: string): string => {
          const parts = s.trim().split(/\s+/);
          return parts.length > 0 ? parts[parts.length - 1] : '';
        };
        const na = a.Nachname || lastWord(a.ParticipantName || '');
        const nb = b.Nachname || lastWord(b.ParticipantName || '');
        cmp = na.localeCompare(nb, 'de');
        break;
      }
      case 'email': cmp = (a.ParticipantEmail || '').localeCompare(b.ParticipantEmail || ''); break;
      case 'status': cmp = (a.Status || '').localeCompare(b.Status || ''); break;
      case 'date': cmp = new Date(a.RegistrationDate || 0).getTime() - new Date(b.RegistrationDate || 0).getTime(); break;
      // v23.33: Job Title / Standort sortierbar.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      case 'jobTitle': cmp = String((a as any).JobTitle || '').localeCompare(String((b as any).JobTitle || ''), 'de'); break;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      case 'location': cmp = String((a as any).Location || '').localeCompare(String((b as any).Location || ''), 'de'); break;
      default: {
        // v23.33: nach Custom-Field-Spalte sortieren (Key „cf-…" eigenes Event,
        // „cfp-…" Parent-Feld im Sub-Event-Detail).
        if (sortColumn.indexOf('cf-') === 0 || sortColumn.indexOf('cfp-') === 0) {
          const cfId = sortColumn.replace(/^cfp?-/, '');
          const valOf = (r: SPRegistration): string => {
            let v: unknown = undefined;
            if (r.CustomData) { try { v = JSON.parse(r.CustomData)[cfId]; } catch { /* */ } }
            // cfp-: Parent-Feld evtl. nur in der Parent-Registrierung.
            if ((v === undefined || v === null || v === '') && sortColumn.indexOf('cfp-') === 0) {
              const pk = (r.ParticipantEmail || '').toLowerCase().trim();
              const pr = pk ? parentRegsByEmail[pk] : undefined;
              if (pr && pr.CustomData) { try { v = JSON.parse(pr.CustomData)[cfId]; } catch { /* */ } }
            }
            return (v === undefined || v === null) ? '' : String(v);
          };
          cmp = valOf(a).localeCompare(valOf(b), 'de');
        }
        break;
      }
    }
    return sortAsc ? cmp : -cmp;
  };

  const handleSort = (col: string): void => {
    if (sortColumn === col) { setSortAsc(!sortAsc); }
    else { setSortColumn(col); setSortAsc(true); }
  };

  const sortIcon = (col: string): string => col === sortColumn ? (sortAsc ? ' \u25B2' : ' \u25BC') : '';

  const activeRegs = registrations
    .filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt')
    .filter(matchesSearch)
    .sort(sortRegs);
  const waitlistRegs = registrations.filter(r => r.Status === 'Warteliste').filter(matchesSearch)
    // v12.10: Warteliste nach TeilnehmerID asc sortieren statt
    // RegistrationDate. Damit ist die UI-Reihenfolge konsistent mit
    // der Nachrück-Logik in promoteFirstWaitlistItem (siehe EventService).
    .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
  // v14.11: konsolidierte Matrix-Zeilen für den „Nur Sub-Events"-Modus
  // (subEventsOnlyMode). Aggregation per ParticipantEmail (lowercase).
  // Standard-Felder werden aus der ersten gefundenen Sub-Event-
  // Registrierung kopiert. Pro Sub-Event wird ein Map-Eintrag mit der
  // jeweiligen aktiven Registrierung gehalten (oder undefined).
  const isConsolidatedMode = !!(selectedEvent
    && selectedEvent.subEventsOnlyMode
    && childEventsOf(selectedEvent.id).length > 0);
  const consolidatedChildren: DeloitteEvent[] = isConsolidatedMode
    ? childEventsOf(selectedEvent!.id)
    : [];
  // v22.59: Abmeldungen-Liste im Klammer-Modus über ALLE Sub-Events
  // konsolidieren (vorher nur die Klammer-Subsite → KPI „Abgemeldet" und Liste
  // klafften auseinander). Jede Zeile trägt ihre Subsite + Section-Titel mit,
  // damit das Löschen die richtige Liste trifft.
  const cancelledRegs: Array<SPRegistration & { _subsiteUrl?: string; _sectionTitle?: string; _sectionId?: string }> = isConsolidatedMode
    ? [
        // v22.63: Klammer-eigene Abmeldungen (z.B. „Ich nehme nicht teil"-Absagen
        // — declineEvent schreibt auf die Klammer-Subsite) MIT aufnehmen, sonst
        // verschwinden sie aus der Liste.
        ...registrations
          .filter(r => r.Status === 'Abgemeldet')
          .map(r => ({ ...r, _subsiteUrl: selectedEvent!.subsiteUrl, _sectionTitle: isDe ? 'Gesamt-Event' : 'Overall event', _sectionId: '__parent' })),
        ...consolidatedChildren.reduce<Array<SPRegistration & { _subsiteUrl?: string; _sectionTitle?: string; _sectionId?: string }>>((acc, ch) => {
          for (const r of (subEventRegsByEventId[ch.id] || [])) {
            if (r.Status !== 'Abgemeldet') continue;
            acc.push({ ...r, _subsiteUrl: ch.subsiteUrl, _sectionTitle: shortSubEventTitle(ch.title, selectedEvent!.title), _sectionId: ch.id });
          }
          return acc;
        }, []),
      ].filter(matchesSearch)
    : registrations.filter(r => r.Status === 'Abgemeldet').filter(matchesSearch);
  // v14.11: wenn ein Sub-Event direkt selektiert ist, das Parent-Event
  // ermitteln — der Sub-Event-Detail-View blendet dessen Custom-Fields
  // (Pastel A) zusätzlich neben den eigenen (Pastel B) ein.
  const parentEventForSelected: DeloitteEvent | null = (selectedEvent && selectedEvent.parentEventId)
    ? (allEvents.find(e => e.id === selectedEvent.parentEventId) || null)
    : null;
  // v15.2 HOTFIX: React.useMemo entfernt — der Hook stand NACH dem early
  // return `if (!selectedEvent)` weiter oben (~Zeile 1940) und feuerte
  // damit nur, wenn ein Event selektiert war. Das verletzte die Rules of
  // Hooks (React error #310 „Rendered more hooks than during the previous
  // render") und crashte die App, sobald der User vom Event-Picker auf
  // eine Detail-Ansicht wechselte. Berechnung läuft jetzt pro Render —
  // ist günstig genug, weil consolidatedFiltered weiter unten ohnehin
  // pro Render neu rechnet.
  const consolidatedRows: ConsolidatedRow[] = (() => {
    if (!isConsolidatedMode || !selectedEvent) return [];
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const byEmail: Record<string, ConsolidatedRow> = {};
    for (const ch of consolidatedChildren) {
      const regs = subEventRegsByEventId[ch.id] || [];
      for (const r of regs) {
        if (ACTIVE.indexOf(r.Status) < 0) continue;
        const emailKey = (r.ParticipantEmail || '').toLowerCase().trim();
        if (!emailKey) continue;
        let row = byEmail[emailKey];
        if (!row) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const anyR = r as any;
          row = {
            emailKey,
            email: r.ParticipantEmail || '',
            vorname: r.Vorname || ((r.ParticipantName || '').split(' ')[0] || ''),
            nachname: r.Nachname || (() => {
              const parts = (r.ParticipantName || '').trim().split(/\s+/);
              return parts.length > 1 ? parts.slice(1).join(' ') : '';
            })(),
            jobTitle: anyR.JobTitle || '',
            location: anyR.Location || '',
            company: anyR.Company || '',
            teilnehmerId: r.TeilnehmerID || null,
            earliestRegistrationTs: r.RegistrationDate ? new Date(r.RegistrationDate).getTime() : Number.POSITIVE_INFINITY,
            perChild: {},
            activeCount: 0,
          };
          byEmail[emailKey] = row;
        } else {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const anyR = r as any;
          if (!row.jobTitle && anyR.JobTitle) row.jobTitle = anyR.JobTitle;
          if (!row.location && anyR.Location) row.location = anyR.Location;
          // v24.35: Unternehmen aus späteren (Sub-)Registrierungen nachfüllen,
          // falls die erste Zeile noch keins hatte.
          if (!row.company && anyR.Company) row.company = anyR.Company;
          if (!row.vorname && r.Vorname) row.vorname = r.Vorname;
          if (!row.nachname && r.Nachname) row.nachname = r.Nachname;
          // Früheste RegistrationDate übernehmen (min).
          const ts = r.RegistrationDate ? new Date(r.RegistrationDate).getTime() : Number.POSITIVE_INFINITY;
          if (ts < row.earliestRegistrationTs) row.earliestRegistrationTs = ts;
        }
        row.perChild[ch.id] = r;
        row.activeCount += 1;
      }
    }
    return Object.values(byEmail);
  })();
  // v28.21: ECHTE Doppel-Anmeldungen im Klammer-Modus = zwei aktive Zeilen
  // derselben Person in DEMSELBEN Sub-Event (nur das belegt zwei Plätze).
  // Zwei Zeilen auf der KLAMMER sind dagegen nur doppelte Schatten-Zeilen
  // (v15.25: die Klammer-Zeile ist reine Datenvollständigkeit, ohne Platz,
  // Mail oder Outlook) — die meldet die rote Box nicht mehr als
  // „Doppel-Anmeldung", sondern separat als technischen Hinweis.
  const subEventDupGroups: Array<{ sectionTitle: string; email: string; name: string; count: number; rowsInfo: string }> = (() => {
    if (!isConsolidatedMode || !selectedEvent) return [];
    const out: Array<{ sectionTitle: string; email: string; name: string; count: number; rowsInfo: string }> = [];
    // v30.14: Status + Zeitpunkt je Zeile ausweisen — die Duplikate können auf
    // der WARTELISTE liegen und sind dann in der Teilnehmer-Tabelle unsichtbar
    // (Befund: „2× angemeldet", aber in der Liste nicht auffindbar; beide
    // Zeilen standen als Platz 1+2 auf der Warteliste desselben Tages).
    const fmtDupRow = (r: SPRegistration): string => {
      const st = r.Status || '?';
      const d = r.RegistrationDate ? new Date(r.RegistrationDate) : null;
      const ds = d && isFinite(d.getTime())
        ? d.toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' })
        : '';
      return ds ? `${st} ${ds}` : st;
    };
    for (const ch of consolidatedChildren) {
      const byEmail: Record<string, SPRegistration[]> = {};
      for (const r of (subEventRegsByEventId[ch.id] || [])) {
        if (DUP_ACTIVE_STATI.indexOf(r.Status || '') < 0) continue;
        const em = (r.ParticipantEmail || '').trim().toLowerCase();
        if (!em) continue;
        (byEmail[em] = byEmail[em] || []).push(r);
      }
      Object.keys(byEmail).forEach(em => {
        const rows = byEmail[em];
        if (rows.length < 2) return;
        const f = rows[0];
        out.push({
          sectionTitle: shortSubEventTitle(ch.title, selectedEvent.title),
          email: em,
          name: (f.Vorname && f.Nachname) ? `${f.Vorname} ${f.Nachname}` : (f.ParticipantName || em),
          count: rows.length,
          rowsInfo: rows.map(fmtDupRow).join(' + '),
        });
      });
    }
    return out;
  })();
  // v14.11: Such-Filter + Sort für die konsolidierten Zeilen.
  const consolidatedFiltered: ConsolidatedRow[] = (() => {
    const q = (searchQuery || '').toLowerCase().trim();
    // v23.32: CustomData-Wert pro Feld-ID lesen.
    const cdVal = (cdStr: string | undefined, fid: string): string => {
      if (!cdStr) return '';
      try { const v = JSON.parse(cdStr)[fid]; return (v === undefined || v === null) ? '' : String(v); } catch { return ''; }
    };
    const parentRegOf = (row: ConsolidatedRow): SPRegistration | undefined =>
      registrations.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey);
    // v26.65: CustomData-Werte als durchsuchbaren Text zusammenfügen (statt
    // pro Wert gegen den GANZEN Query zu matchen) — Basis für die Token-Suche.
    const cdHaystack = (cdStr?: string): string => {
      if (!cdStr) return '';
      try { const cd = JSON.parse(cdStr); return Object.keys(cd).map(ck => (cd[ck] == null ? '' : String(cd[ck]))).join(' '); } catch { return ''; }
    };
    // v26.65 BUG-FIX: Token-basierte Suche über EINEN kombinierten Text aus allen
    // durchsuchbaren Feldern. Vorher wurde der komplette Suchstring gegen jedes
    // Einzelfeld geprüft — „Alexander Knoth" (Vor- + Nachname) fand nichts, weil
    // der ganze String weder nur im Vor- noch nur im Nachnamen steht. Jetzt muss
    // JEDES getrennte Wort irgendwo im kombinierten Text vorkommen (Reihenfolge
    // egal → „Knoth Alexander" findet dieselbe Person, Job Title/Standort/Custom-
    // Felder werden weiter mitdurchsucht).
    const matches = (row: ConsolidatedRow): boolean => {
      if (!q) return true;
      const tokens = q.split(/\s+/).filter(Boolean);
      if (tokens.length === 0) return true;
      const pReg = parentRegOf(row);
      let hay = `${row.vorname} ${row.nachname} ${row.email} ${String(row.teilnehmerId || '')} ${row.jobTitle || ''} ${row.location || ''} ${row.company || ''}`;
      if (pReg) hay += ' ' + cdHaystack(pReg.CustomData);
      for (const cid of Object.keys(row.perChild)) { const r = row.perChild[cid]; if (r) hay += ' ' + cdHaystack(r.CustomData); }
      hay = hay.toLowerCase();
      return tokens.every(t => hay.indexOf(t) >= 0);
    };
    const filtered = consolidatedRows.filter(matches);
    const cs = consolidatedSort;
    const dir = consolidatedSortAsc ? 1 : -1;
    const cmp = (a: ConsolidatedRow, b: ConsolidatedRow): number => {
      // v15.23: #-Spalte sortiert nach Reihenfolge der ersten Anmeldung
      // (RegistrationDate min), nicht mehr nach TeilnehmerID — die TID ist
      // pro Sub-Event und bei konsolidierten Personen mehrdeutig.
      if (cs === 'id') return (a.earliestRegistrationTs - b.earliestRegistrationTs) * dir;
      if (cs === 'vorname') return a.vorname.localeCompare(b.vorname, 'de') * dir;
      if (cs === 'nachname') return a.nachname.localeCompare(b.nachname, 'de') * dir;
      if (cs === 'email') return a.email.localeCompare(b.email) * dir;
      if (cs === 'jobTitle') return a.jobTitle.localeCompare(b.jobTitle, 'de') * dir;
      if (cs === 'location') return a.location.localeCompare(b.location, 'de') * dir;
      if (cs && cs.indexOf('child:') === 0) {
        const cid = cs.substring(6);
        const ra = a.perChild[cid] ? 1 : 0;
        const rb = b.perChild[cid] ? 1 : 0;
        return (rb - ra) * dir;
      }
      // v23.32: nach Parent-Custom-Feld sortieren (Key „pf:<fieldId>").
      if (cs && cs.indexOf('pf:') === 0) {
        const fid = cs.substring(3);
        return cdVal(parentRegOf(a)?.CustomData, fid).localeCompare(cdVal(parentRegOf(b)?.CustomData, fid), 'de') * dir;
      }
      // v23.32: nach Sub-Event-Custom-Feld sortieren (Key „cf:<childId>|<fieldId>").
      if (cs && cs.indexOf('cf:') === 0) {
        const rest = cs.substring(3);
        const sep = rest.indexOf('|');
        const cid = sep >= 0 ? rest.slice(0, sep) : rest;
        const fid = sep >= 0 ? rest.slice(sep + 1) : '';
        return cdVal(a.perChild[cid]?.CustomData, fid).localeCompare(cdVal(b.perChild[cid]?.CustomData, fid), 'de') * dir;
      }
      return 0;
    };
    return filtered.sort(cmp);
  })();
  // v14.11: konsolidierter Matrix-View. Wird nur gerendert, wenn das
  // selektierte Hauptevent `subEventsOnlyMode === true` ist und mind.
  // ein Sub-Event hat. Standard-Spalten neutral, parent-event-level
  // Custom-Fields in Pastel A (hellblau), pro Sub-Event eine X-Spalte,
  // sub-event-spezifische Custom-Fields in Pastel B (hellgelb)
  // gruppiert pro Sub-Event-Header.
  // v30.60: Gibt es an diesem Event überhaupt ein Größen-Feld? Entscheidet, ob
  // die Aktion „Benötigte T-Shirts" erscheint. Geprüft wird über dieselbe Regel
  // wie am Check-in-Tisch (utils/checkInExtras) — auf der Klammer UND auf den
  // Terminen, weil das Feld auf beiden Ebenen stehen kann.
  //
  // BEWUSST OHNE useMemo: Diese Stelle liegt hinter einem frühen Return
  // (dieselbe Lage wie der v15.2-Hotfix oben) — ein Hook hier reißt die
  // Hook-Reihenfolge und ESLint bricht den Build ab. Die Prüfung läuft über
  // eine Handvoll Feld-Labels; das ist billiger als die Rettung wäre.
  const shirtFieldExists = !!selectedEvent
    && [selectedEvent, ...events.filter(e => e.parentEventId === selectedEvent.id)].some(ev =>
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (((ev as any).eventSpecificFields || []) as Array<{ label?: string }>)
        .some(f => SHIRT_PATTERN.test(f.label || '')));
  const waitlistDurch = isSplitCapacity
    ? waitlistRegs.filter(r => r.PreferredStarterType === 'Durchstarter')
    : [];
  const waitlistFun = isSplitCapacity
    ? waitlistRegs.filter(r => r.PreferredStarterType === 'Funstarter')
    : [];
  const waitlistUnassigned = isSplitCapacity
    ? waitlistRegs.filter(r => !r.PreferredStarterType || (r.PreferredStarterType !== 'Durchstarter' && r.PreferredStarterType !== 'Funstarter'))
    : [];

  // v26.31: Wahre FIFO-Position je Wartelisten-Person (Rang nach TeilnehmerID
  // innerhalb der jeweiligen Gruppe), berechnet aus der UNGEFILTERTEN Liste —
  // damit die „Platz"-Anzeige auch beim Suchen/Filtern korrekt bleibt (sonst
  // zeigte eine gefilterte Trefferliste fälschlich Platz 1, 2, …).
  // v30.67: Dazu die wahre GRUPPENGRÖSSE je Person (`waitlistTrueTotal`) —
  // der „Platz ändern"-Dialog bekam bisher `regs.length`, also die Zahl der
  // SUCHTREFFER: Bei aktiver Suche hieß es „Platz 7 von 1", `max=1`, und nur
  // die Vorbelegung „1" ließ sich speichern.
  const { waitlistTruePos, waitlistTrueTotal } = ((): { waitlistTruePos: Record<number, number>; waitlistTrueTotal: Record<number, number> } => {
    const map: Record<number, number> = {};
    const total: Record<number, number> = {};
    const rank = (arr: SPRegistration[]): void => {
      const s = arr.slice().sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
      s.forEach((r, idx) => { if (typeof r.Id === 'number') { map[r.Id] = idx + 1; total[r.Id] = s.length; } });
    };
    const all = registrations.filter(r => r.Status === 'Warteliste');
    // v30.67 (Review): je Gruppe NUR bei getrennten Wartelisten. Bei
    // `splitSharedWaitlist` ist es ein Topf (CLAUDE.md): der „Platz ändern"-
    // Dialog übergibt dann `group = undefined`, der Service sortiert über die
    // Gesamtliste, und das Nachrücken rechnet ebenso — hier stand trotzdem
    // der Gruppen-Rang („Platz 2 von 5" in der Tabelle, „Platz 1, vorher 4"
    // im Ergebnis). Die drei Gruppen-Tabellen bleiben getrennt gerendert —
    // sie zeigen den Wunsch-Typ, das ist Information —, die Platz-Zahl gilt
    // dann aber im gemeinsamen Topf, wie überall sonst.
    if (isSplitCapacity && !selectedEvent?.splitSharedWaitlist) {
      rank(all.filter(r => r.PreferredStarterType === 'Durchstarter'));
      rank(all.filter(r => r.PreferredStarterType === 'Funstarter'));
      rank(all.filter(r => !r.PreferredStarterType || (r.PreferredStarterType !== 'Durchstarter' && r.PreferredStarterType !== 'Funstarter')));
    } else {
      rank(all);
    }
    return { waitlistTruePos: map, waitlistTrueTotal: total };
  })();

  // Roommate-Matching: durchsucht CustomData nach roommate-Type Feldern, extrahiert
  // Email aus "Name <email>"-Format, baut Map Email -> Partner-Email. Match-Badge,
  // wenn beide sich gegenseitig ausgewählt haben.
  // v11.65: ausschliesslich `roommate`-Felder, nicht mehr `user`. Bei Assistant-
  // /generischen User-Pickern macht ein „Match"-Badge semantisch keinen Sinn —
  // der wurde fälschlich auch dort gezeigt, wenn Person A und B sich
  // gegenseitig als Assistant eingetragen haben.
  const userFieldIds = (selectedEvent?.eventSpecificFields || [])
    .filter(f => f.type === 'roommate')
    .map(f => f.id);

  // Render-Funktionen pro Spalte — als eine Map, damit der Header + die Body-Zeilen
  // die gleiche stabile ID benutzen. Wird bei jedem Registration-Render neu aufgebaut,
  // weil die renderCell-Lambdas auf aktuelle Closures (handleSort, sortIcon, …) angewiesen sind.
  // Das ist günstig, weil die Funktion nur Pointer speichert.
  const roommateChoice: Record<string, { partnerEmail: string; partnerName: string }> = {};
  if (userFieldIds.length > 0) {
    const allActiveAndWaitlist = registrations.filter(r =>
      r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt' || r.Status === 'Warteliste'
    );
    for (const r of allActiveAndWaitlist) {
      const email = (r.ParticipantEmail || '').toLowerCase();
      if (!email) continue;
      let cd: Record<string, string> = {};
      try { cd = JSON.parse(r.CustomData || '{}'); } catch { /* */ }
      for (const fid of userFieldIds) {
        const v = cd[fid];
        if (!v) continue;
        const m = v.match(/<([^>]+@[^>]+)>/);
        if (!m) continue;
        const pEmail = m[1].trim().toLowerCase();
        const pName = v.replace(/<[^>]*>/, '').trim();
        if (pEmail && pEmail !== email) {
          roommateChoice[email] = { partnerEmail: pEmail, partnerName: pName };
          break; // nur erstes user-Feld
        }
      }
    }
  }
  const getRoommateInfo = (reg: { ParticipantEmail?: string }): { partnerName: string; partnerEmail: string; mutual: boolean } | null => {
    const email = (reg.ParticipantEmail || '').toLowerCase();
    if (!email) return null;
    const choice = roommateChoice[email];
    if (!choice) return null;
    const reverse = roommateChoice[choice.partnerEmail];
    const mutual = !!reverse && reverse.partnerEmail === email;
    return { partnerName: choice.partnerName || choice.partnerEmail, partnerEmail: choice.partnerEmail, mutual };
  };

  // v26.44: berechnet die gegenseitigen Roommate-Paare INNERHALB einer
  // Zeilenmenge. Ein Paar zählt nur, wenn BEIDE Personen in `rows` enthalten
  // sind — ist der Partner rausgefiltert (Suche) oder gar nicht (mehr)
  // registriert, gilt die Person als „ohne Match". Dedupe über den Schlüssel
  // aus beiden lowercased E-Mails (sortiert + gejoint), damit jedes Paar genau
  // EINMAL erscheint. Wird sowohl von der „Matches anzeigen"-Gruppierung als
  // auch vom Excel-Export („Roommate-Match"-Spalte) benutzt.
  const computeRoommatePairs = (rows: SPRegistration[]): Array<[SPRegistration, SPRegistration]> => {
    const byEmail: Record<string, SPRegistration> = {};
    for (const r of rows) {
      const e = (r.ParticipantEmail || '').trim().toLowerCase();
      if (e && !byEmail[e]) byEmail[e] = r;
    }
    const seen = new Set<string>();
    const pairs: Array<[SPRegistration, SPRegistration]> = [];
    for (const r of rows) {
      const email = (r.ParticipantEmail || '').trim().toLowerCase();
      if (!email) continue;
      const info = getRoommateInfo(r);
      if (!info || !info.mutual) continue;
      const partnerEmail = (info.partnerEmail || '').trim().toLowerCase();
      const partner = byEmail[partnerEmail];
      if (!partner || partner === r) continue;
      const key = [email, partnerEmail].sort().join('|');
      if (seen.has(key)) continue;
      seen.add(key);
      pairs.push([r, partner]);
    }
    return pairs;
  };

  // v30.66: createExportActions — Rumpf in logic/createExportActions.ts.
  const {
    exportConsolidatedExcel, exportCsv,
  } = createExportActions({
    computeRoommatePairs, consolidatedChildren, hasRoommateColumn, isDe, registrations,
    selectedEvent, showAlert, subEventRegsByEventId,
  });
  // v30.66: createKlammerActions — Rumpf in logic/createKlammerActions.ts.
  const {
    addAllToKlammer, addToKlammer, closeDeregModal, closeMainFieldsEdit, openDeregModal,
    openMainFieldsEdit, runDeregModal, saveMainFieldsEdit, submitAssignAssistant,
  } = createKlammerActions({
    assignAssistRow, assignAssistValue, buildCancellationMail, bulkKlammerProgress, childEventsOf,
    confirmDialog, consolidatedChildren, consolidatedRows, currentUser, deregModal, deregSelected,
    deregSilent, eventServiceRef, inactiveAccounts, isDe, mainFieldsEditForm,
    mainFieldsEditName, mainFieldsEditReg, mainFieldsEditSubsite, mainFieldsEditTargetIsParent,
    registerForEvent, registrations, reloadRegistrations, selectedEvent, setAddingToKlammer, setAssignAssistBusy,
    setAssignAssistRow, setAssignAssistValue, setBulkKlammerProgress, setDeregBusy, setDeregModal,
    setDeregSelected, setDeregSilent, setMainFieldsEditError, setMainFieldsEditForm,
    setMainFieldsEditName, setMainFieldsEditReg, setMainFieldsEditSaving, setMainFieldsEditSubsite,
    setMainFieldsEditTargetIsParent,
    setSubRegReloadTick, showAlert,
  });
  // v30.66: Props-Buendel der ausgelagerten Teilansichten. Bewusst hier, direkt
  // vor dem return — davor stehen alle Deklarationen, weiter oben waere es ein
  // TDZ-Fehler auf die spaeter deklarierten Handler.
  const nameFixModalProps = {
    isDe, nameFixModal, setNameFixModal,
  };
  const accessFixModalProps = {
    accessFixModal, isDe, setAccessFixModal,
  };
  const selfCheckInModalProps = {
    isDe, navigate, saveSelfCheckInWindow, sciBusy, sciFrom, sciModalOpen,
    sciModalQr, sciSaveMsg, sciTo, sciToken, selectedEvent, setSciFrom,
    setSciModalOpen, setSciTo,
  };
  const checkInHubModalProps = {
    checkInHubOpen, checkInHubStep, isDe, navigate, openSelfCheckInModal, sciBusy,
    selectedEvent, setCheckInHubOpen, setCheckInHubStep, setQrSendModalOpen, setQrSendResult,
  };
  const qrSendModalProps = {
    childEventsOf, currentUser, getQrMailOverride, isDe, isSendingQR, openQrMailEditor,
    qrFullSendAction, qrHelpOpen, qrPreviewAction, qrPreviewLoading, qrSendModalOpen, qrSendResult,
    qrSentCount, qrSubMailsOpen, qrTestSendAction, registrations, selectedEvent, setQrHelpOpen,
    setQrSendModalOpen, setQrSubMailsOpen,
  };
  const qrEditModalProps = {
    closeQrMailEditor, currentUser, getQrMailOverride, isDe, isSendingQR, qrBlockLang,
    qrBlockNote, qrEditBody, qrEditHeading, qrEditOpen, qrEditSampleBlock, qrEditSampleImg,
    qrEditSaving, qrEditSubheading, qrEditSubject, qrEditTarget, qrEventPhotoB64, qrFullSendAction,
    qrHeaderImage, qrSendResult, qrSentCount, qrTestSendAction, registrations, saveQrMailOverride,
    selectedEvent, setComposerCrop, setQrBlockLang, setQrBlockNote, setQrEditBody, setQrEditHeading,
    setQrEditSampleBlock, setQrEditSubheading, setQrEditSubject, setQrHeaderImage,
  };
  const editRegModalProps = {
    closeEditModal, editError, editForm, isDe, isSavingEdit,
    saveEdit, selectedEvent, setEditForm,
  };
  const participantDetailModalProps = {
    isDe, participantDetail, setParticipantDetail,
  };
  const assignAssistModalProps = {
    assignAssistBusy, assignAssistRow, assignAssistValue, isDe, searchUser, searchUsers,
    setAssignAssistRow, setAssignAssistValue, submitAssignAssistant,
  };
  const mainFieldsEditModalProps = {
    closeMainFieldsEdit, isDe, mainFieldsEditError, mainFieldsEditForm, mainFieldsEditName, mainFieldsEditReg,
    mainFieldsEditSaving, saveMainFieldsEdit, selectedEvent, setMainFieldsEditForm,
  };
  const deregModalProps = {
    closeDeregModal, deregBusy, deregModal, deregSelected, deregSilent, inactiveAccounts,
    isDe, runDeregModal, selectedEvent, setDeregSelected, setDeregSilent,
  };
  const qrPreviewModalProps = {
    isDe, qrPreviewHtml, qrPreviewOpen, qrPreviewSubject, setQrPreviewOpen,
  };
  const commsLogModalProps = {
    commsDeletingId, commsExpandedId, commsLoading, commsRows, confirmDialog, deleteCommRow,
    isDe, selectedEvent, setCommsExpandedId, setShowCommsModal, showCommsModal,
  };
  const massmailPickModalProps = {
    massmailAudience, massmailStatuses, registrations, setMassmailAudience,
    setMassmailMode, setMassmailPasteRaw, setMassmailStatuses, setShowEmailModal,
  };
  const massmailPasteModalProps = {
    massmailPasteRaw, registrations, setMassmailMode, setMassmailPasteRaw,
    setShowEmailModal, showAlert,
  };
  const excelTargetModalProps = {
    consolidatedChildren, excelAudience, excelIncludeMatrix, excelSubIds, excelTargetModal, exportConsolidatedExcel,
    exportCsv, isConsolidatedMode, isDe, selectedEvent, setExcelAudience, setExcelIncludeMatrix,
    setExcelSubIds, setExcelTargetModal, subEventRegsByEventId,
  };
  const massmailComposerModalProps = {
    applyMassmailHero, confirmDialog, emailBody, emailHeading, emailSending, emailSubject,
    eventServiceRef, isDe, massmailAudience, massmailCc, massmailDraftSaved, massmailEventPhotoB64,
    massmailHeaderImage, massmailHeaderOpts, massmailPasteRaw, massmailStatuses, massmailSubheading, massmailTesting,
    massmailTestMsg, registrations, resetMassmailDraft, saveMassmailDraft, searchUser, searchUsers,
    selectedEvent, sendMassmailTestToOrganizers, setComposerCrop, setEmailBody, setEmailHeading, setEmailSending,
    setEmailSubject, setMassmailCc, setMassmailHeaderImage, setMassmailMode, setMassmailPasteRaw, setMassmailSubheading,
    setShowEmailModal, showAlert, showEmailModal,
  };
  const inviteComposerModalProps = {
    applyInviteHero, confirmDialog, currentUser, eventServiceRef, getGroupMembers, inviteAddInput,
    inviteAudienceOpen, inviteBody, inviteCc, inviteCustomEmails, invitedLc, inviteDraftSaved,
    inviteEventPhotoB64, inviteHeaderImage, inviteHeaderOpts, inviteHeading, inviteSending, inviteSubheading,
    inviteSubject, inviteTarget, isDe, refreshEvents, registrations, resetInviteDraft,
    saveInviteDraft, searchUser, searchUsers, selectedEvent, setComposerCrop, setInviteAddInput,
    setInviteAudienceOpen, setInviteBody, setInviteCc, setInviteCustomEmails, setInviteHeaderImage, setInviteHeading,
    setInviteSending, setInviteSubheading, setInviteSubject, setInviteTarget, setShowInviteModal, showAlert,
    showInviteModal, siteUrl, updateEvent,
  };
  const declineCheckModalProps = {
    declineCopied, declineResult, isDe, setDeclineCopied, setShowDeclineModal, showAlert,
    showDeclineModal,
  };
  const attachmentsModalProps = {
    attachmentsBusy, attachmentsByReg, attachmentsModalReg, confirmDialog, eventServiceRef, isDe,
    selectedEvent, setAttachmentsBusy, setAttachmentsByReg, setAttachmentsModalReg, showAlert,
  };
  const reorderProgressOverlayProps = {
    reorderProgress, reorderProgressLabel,
  };
  const dupCancelModalProps = {
    dupCancelBusy, dupCancelReg, isDe, performSilentDuplicateDelete, performStandardCancel,
    setDupCancelBusy, setDupCancelReg, showAlert,
  };
  const overbookDecisionModalProps = {
    obBusy, obKeepVariant, obMailBody, obMailLang, obMailSubject, obRemoveCalendar,
    obWithMail, overbookModal, runOverbookResolution, setObKeepVariant, setObMailBody,
    setObMailLang, setObMailSubject, setObRemoveCalendar, setObWithMail, setOverbookModal,
  };
  const waitlistPositionModalProps = {
    eventServiceRef, isDe, reloadRegistrations, selectedEvent, setWlPosBusy,
    setWlPosModal, setWlPosValue, showAlert, wlPosBusy, wlPosModal, wlPosValue,
  };
  const adminAddMemberModalProps = {
    addTeamMember, adminAddCcOrganizer, adminAddLeadRegId, adminAddMemberBusy, adminAddMemberConsent, adminAddMemberDialog,
    adminAddMemberError, adminAddMemberIncludeIntl, adminAddMemberPick, adminAddMemberQuery, adminAddMemberQueryTimer, adminAddMemberResults,
    adminAddMemberSearching, adminAddNewPersonMail, adminAddNotifyOthers, adminAddNotifyScope, adminAddSendMail, adminAddTeamlessPicks,
    assignTeamlessToTeam, currentUser, isDe, notifyExistingTeamMembers, registrations, reloadRegistrations,
    searchUsers, selectedEvent, setAdminAddCcOrganizer, setAdminAddLeadRegId, setAdminAddMemberBusy, setAdminAddMemberConsent,
    setAdminAddMemberDialog, setAdminAddMemberError, setAdminAddMemberIncludeIntl, setAdminAddMemberPick, setAdminAddMemberQuery, setAdminAddMemberResults,
    setAdminAddMemberSearching, setAdminAddNewPersonMail, setAdminAddNotifyOthers, setAdminAddNotifyScope, setAdminAddSendMail, setAdminAddTeamlessPicks,
    setTeamsToast,
  };
  // v30.66: Props-Buendel der ausgelagerten Teilansichten. Bewusst hier, direkt
  // vor dem return — davor stehen alle Deklarationen, weiter oben waere es ein
  // TDZ-Fehler auf die spaeter deklarierten Handler.
  const pendingPeopleBoxProps = {
    isDe, openInviteModal, pendingPeople, setInviteAudienceOpen, setInviteCustomEmails, setInviteTarget,
    setPendingPeople, showAlert,
  };
  const idGapHintBoxProps = {
    confirmDialog, eventServiceRef, idRecheckBusy, isAdmin, isDe, isOrganizerFor,
    isReorderingIDs, recentCancellation, registrations, reloadRegistrationsForIdCheck, runIdReorder, selectedEvent,
  };
  const duplicateRegHintBoxProps = {
    cleanupShadowDuplicates, duplicateEmails, isAdmin, isConsolidatedMode, isDe, isOrganizerFor,
    registrations, selectedEvent, shadowDupBusy,
  };
  const duplicateInSubEventHintBoxProps = {
    isDe, subEventDupGroups,
  };
  const missingEmailHintBoxProps = {
    isDe, missingEmailRegs, selectedEvent,
  };
  const overbookReviewBoxProps = {
    isDe, isSplitCapacity, registrations, selectedEvent, setObKeepVariant, setObRemoveCalendar,
    setObWithMail, setOverbookModal,
  };
  const teamsSectionProps = {
    confirmDialog, currentUser, dragOverTid, dragRegId, eventServiceRef, getActiveTeams,
    isAdmin, isDe, isLoadingRegs, isMobile, isOrganizerFor,
    leadTransferBusy, leadTransferOpenFor, moveRegToTeam, onTeamDrop, openTeamMailDialog, registrations,
    reloadRegistrations, selectedEvent, setAdminAddCcOrganizer, setAdminAddLeadRegId, setAdminAddMemberConsent, setAdminAddMemberDialog, setAdminAddMemberError,
    setAdminAddMemberPick, setAdminAddMemberQuery, setAdminAddMemberResults, setAdminAddNewPersonMail, setAdminAddNotifyOthers, setAdminAddNotifyScope,
    setAdminAddSendMail, setAdminAddTeamlessPicks, setDragOverTid, setDragRegId, setLeadTransferBusy, setLeadTransferOpenFor,
    setTeamEditOpenFor, setTeamsCollapsed, setTeamsToast, showAlert, teamEditOpenFor,
    teamsCollapsed, transferTeamLead,
  };
  const teamMailModalProps = {
    getActiveTeams, isDe, selectedEvent, sendTeamMails, setTeamMailBody, setTeamMailInfoByTid,
    setTeamMailOpen, setTeamMailSubject, teamMailBody, teamMailInfoByTid, teamMailSending,
    teamMailSubject,
  };
  const participantTableProps = {
    activeRegs, allEvents, attachmentsByReg, availableColumns, colToggleHover, columnOrder,
    computeRoommatePairs, confirmDialog, duplicateEmails, eventServiceRef, getRoommateInfo,
    handleSort, hasRoommateColumn, hiddenColumns, hideColumn, highlightMatch, inactiveAccounts,
    isDe, isSplitCapacity, moveColumn, openEditModal, orgPastLock, parentEventForSelected,
    parentRegsByEmail, performStandardCancel, personalColsCollapsed, query, registrations, reloadRegistrations, selectedEvent,
    setAttachmentsModalReg, setColToggleHover, setDupCancelReg, setParticipantDetail, setPersonalColsCollapsed,
    setShowColumnPicker, setSplitParticipantsView, showAlert, showColumn, showColumnPicker, showMatches,
    sortIcon, splitParticipantsView, stripLocPrefix,
  };
  const waitlistTablesProps = {
    buildCancellationMail, confirmDialog, currentUser, eventServiceRef, isDe,
    isSplitCapacity, query, reloadRegistrations, selectedEvent, setWaitlistSortAsc, setWaitlistSortColumn,
    setWlPosModal, setWlPosValue, showAlert, waitlistDurch, waitlistFun, waitlistRegs,
    waitlistSortAsc, waitlistSortColumn, waitlistTruePos, waitlistTrueTotal, waitlistUnassigned, wlPosBusy,
  };
  const cancelledListProps = {
    cancelledRegs, cancelledSortAsc, cancelledSortColumn, confirmDialog, consolidatedChildren, eventServiceRef,
    hasWaitlistActivity, isAdmin, isConsolidatedMode, isDe, isOrganizerFor,
    registrations, reloadRegistrations, selectedEvent, setCancelledSortAsc, setCancelledSortColumn, setSubRegReloadTick,
    showAlert, stripLocPrefix, subEventRegsByEventId,
  };
  // v30.66: Props-Buendel der ausgelagerten Teilansichten. Bewusst hier, direkt
  // vor dem return — davor stehen alle Deklarationen, weiter oben waere es ein
  // TDZ-Fehler auf die spaeter deklarierten Handler.
  const adminToastProps = {
    adminToast, isMobile, setAdminToast,
  };
  const inactiveAccountsBoxProps = {
    consolidatedRows, inactiveAccounts, isAdmin, isConsolidatedMode, isDe, isOrganizerFor,
    jumpToParticipant, openDeregModal, orgPastLock, registrations, selectedEvent,
  };
  const duplicateEventsBoxProps = {
    duplicateEvents, isDe, selectedEvent, setConfirmDeleteEvent,
  };
  const eventDetailCardProps = {
    activeRegs, childEventsOf, confirmDialog, detailCardRef, events,
    evTabHover, handleSelectEvent, isAdmin, isConsolidatedMode, isDe, isImpersonating,
    isLoadingRegs, isMobile, isOrganizerFor, navigate, openTabGroup, registrations, regsUnknown,
    reservedDetailHeight, reservedDetailWidth, selectedEvent, setCheckInHubOpen, setCheckInHubStep, setEvTabHover,
    setOpenTabGroup, subEventRegsByEventId, subListsIncomplete, t, toggleDraftStatus, waitlistRegs,
  };
  const nextStepsBoxProps = {
    childEventsOf, isDe, openInviteModal, selectedEvent,
    setVisListOpen, visListOpen,
  };
  const billingStatusStripProps = {
    isAdmin, isDe, isFA, isOrganizerFor, navigate,
    selectedEvent, setBillingPanelOpen,
  };
  const adminActionsCardProps = {
    adminEvents, allEvents, childEventsOf, confirmDialog, copiedDeepLink, copiedEmails,
    detectOverbookResult, eventServiceRef, fixColumnsResult, fixFieldsResult, isAdmin,
    isCheckingDeclines, isDe, isDetectingOverbook, isFixingColumns, isFixingFields, isOrganizerFor,
    isPromoting, isRefreshingProfiles, isReorderingIDs, isRepairingAccess, isRepairingNames,
    isRepairingOrganizers, isRepairingPerms, isResettingCounter, isSendingQR, isSplitCapacity, isSyncingRegistry,
    navigate, openChangeLogForEvent, openCommsModal, openInviteModal, openMassmailPicker, promoteResult,
    qrSentCount, refreshEvents, refreshProfilesResult, registrations, reloadRegistrations, reorderResult, repairAccessResult,
    repairNamesResult, repairOrganizersResult, repairPermsResult, resetCounterResult, runIdReorder, runManualPromote,
    searchUsers, selectedEvent, setAccessFixModal, setB2runTodoOpen, setBibImportOpen, setBillingPanelOpen,
    setCheckInHubOpen, setCheckInHubStep, setCopiedDeepLink, setCopiedEmails, setDeclineCopied, setDeclineResult,
    setDetectOverbookResult, setExcelAudience, setExcelTargetModal, setFixColumnsResult, setFixFieldsResult, setIsCheckingDeclines,
    setIsDetectingOverbook, setIsFixingColumns, setIsFixingFields, setIsRefreshingProfiles, setIsRepairingAccess, setIsRepairingNames,
    setIsRepairingOrganizers, setIsRepairingPerms, setIsResettingCounter, setIsSyncingRegistry, setNameFixModal, setRefreshProfilesResult,
    setRepairAccessResult, setRepairNamesResult, setRepairOrganizersResult, setRepairPermsResult, setResetCounterResult,
    setShirtSizeOpen, setShowDeclineModal, setShowExportMenu, setSubRegReloadTick, setSyncRegistryResult, shirtFieldExists,
    showAlert, showExportMenu, siteUrl, spServiceRef, syncRegistryResult, t,
    updateEvent,
  };
  const kpiTilesProps = {
    isConsolidatedMode, isDe, isSplitCapacity, registrations, regsUnknown, selectedEvent, subEventRegsByEventId, subListsIncomplete, t,
  };
  const hotelPlanningSectionProps = {
    childEventsOf, confirmDialog, hotelPanelOpen, isDe,
    refreshEvents, registrations, reloadRegistrations, selectedEvent, setHotelPanelOpen,
    showAlert, subEventRegsByEventId,
  };
  const quizStatsSectionProps = {
    registrations, selectedEvent,
  };
  const activeEventHintsBoxProps = {
    childEventsOf, expandedHintIds, hintLangBusy, hintsDismissTick, isDe,
    parentEventForSelected, refreshEvents, selectedEvent, setExpandedHintIds, setHintLangBusy,
    setHintsDismissTick, setQrSendModalOpen, setSelectedEvent, showAlert, updateEvent,
  };
  const audienceVisibilityRowProps = {
    isAdmin, isDe, isOrganizerFor, openPendingReminder, orgPastLock, pendingCheckBusy,
    resolveAudienceEmails, selectedEvent, setVisibilityAllAddresses, setVisibilityBusy, setVisibilityOpen, setVisibilityResolved,
    visibilityAllAddresses, visibilityBusy, visibilityOpen, visibilityResolved,
  };
  const consolidatedViewProps = {
    addAllToKlammer, addingToKlammer, addToKlammer, bulkKlammerProgress, colToggleHover, confirmDialog,
    consolidatedChildren, consolidatedFiltered, consolidatedRows, consolidatedSort, consolidatedSortAsc, deniedSubEventLists, expandedConsolidatedEmail,
    highlightMatch, inactiveAccounts, isAdmin, isConsolidatedMode, isDe, isLoadingSubEventRegs,
    isOrganizerFor, missingReminderKey, openDeregModal, openMainFieldsEdit, orgPastLock, performSilentDuplicateDelete,
    personalColsCollapsed, registrations, reminderBusyId, searchQuery, selectedEvent, sendCompleteRegistrationReminder,
    setAssignAssistRow, setAssignAssistValue, setColToggleHover, setConsolidatedSort, setConsolidatedSortAsc, setExpandedConsolidatedEmail,
    setMissingReminderKey, setParticipantDetail, setPersonalColsCollapsed, setReminderBusyId, setSelectedEvent, showAlert,
    stripLocPrefix, subEventRegsByEventId,
  };
  return (
    <div className="page-container" role="main">
      {/* Admin-Toast: drei Phasen beim Abmelden (seit v6.8).
          1. cancelling — orange, Spinner, läuft während der Abmeldung+Promote-Suche
          2. promoted   — grün, zeigt den Nachrücker
          3. no-promote — grau, Abmeldung ok, keiner auf der Warteliste */}
      {adminToast && <AdminToast {...adminToastProps} />}
      {/* Keyframes für Spinner */}
      <style>{`@keyframes dex-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
      {/* v9.29: Inline Zurück + Aktualisieren entfernt — beides liegt jetzt im Header.
          Eventauswahl-Reset („zurück zur Event-Liste") triggern wir über den Header-Back —
          siehe Listener weiter oben, der bei navigate-Wechsel selectedEvent zurücksetzt. */}

      {/* v22.7: Hinweisbox, wenn Teilnehmer-Konten nicht mehr aktiv sind
          (Person hat womöglich Deloitte verlassen). Hintergrund-Check beim
          Öffnen, max. 1×/Tag pro Event. */}
      {inactiveAccounts.length > 0 && <InactiveAccountsBox {...inactiveAccountsBoxProps} />}

      {/* v26.18/v26.22: Duplikat-Warnung — gleicher/ähnlicher Name + gleicher Tag.
          Es werden ALLE betroffenen Versionen gelistet (auch die gerade geöffnete),
          jeweils mit Erstell-Zeitstempel + Löschen-Knopf. */}
      {duplicateEvents.length > 0 && selectedEvent && <DuplicateEventsBox {...duplicateEventsBoxProps} />}

      {/* v12.7: Aktionen-Card aufgelöst — alle ActionTiles registrieren
          sich jetzt im ActionsRegistryProvider. Die Dropdown-Liste sitzt
          unten in der linken Event-Detail-Card. Daher 1-Spalten-Grid
          statt vorher 2-Spalten. */}
      <ActionsRegistryProvider>
      <div className="admin-event-info-grid" style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 24, marginBottom: 24 }}>
        {/* v22.5: Detail-Card + „Nächste Schritte"-Box rechts daneben (Desktop;
            stapelt auf Mobile via flex-wrap). Die Box erscheint nur für Entwürfe
            und nur für Admin/Organizer. */}
        <div style={{ display: 'flex', gap: 24, alignItems: 'flex-start', flexWrap: 'wrap' }}>
        <EventDetailCard {...eventDetailCardProps} />
        {/* v22.5: „Nächste Schritte"-Box rechts neben der Detail-Card — nur für
            Entwürfe (Admin/Organizer). Erklärt, was nach dem Anlegen noch zu tun
            ist: finalisieren, Test-An-/Abmeldung, live schalten (+ wer es sieht),
            Einladungsmail verschicken, Anmeldungen verfolgen.
            v28.47: flex-grow statt 0 — auf breiten Laptop-Screens schob die
            Detail-Card (flex 1 1 420px + gemessene minWidth) die Box in die
            nächste Zeile, wo sie mit 460px als schmale Saeule links stand und
            rechts daneben alles leer blieb. Mit grow füllt sie die Zeile. */}
        {(isAdmin || isOrganizerFor(selectedEvent)) && !!selectedEvent.isFictive && !selectedEvent.isDemoShowcase && <NextStepsBox {...nextStepsBoxProps} />}
        </div>

        {/* v7.6: Aktionen-Bereich als Kachel-Grid (auto-fit ab 220px, max 4
            pro Zeile auf Desktop). Default Grau, beim Hover Deloitte-Grün mit
            leichtem Schatten. Jede Kachel zeigt SVG-Icon + Titel + ausführliche
            Beschreibung + Rollen-Badge ("Organizer" oder "Nur Admin"). Die
            ehemals in der TN-Toolbar versteckten Wartungs-Aktionen (IDs neu
            vergeben, Spalten fixen, Felder reparieren, Profile neu laden) sind
            seit v7.6 hier integriert — der Organizer/Admin findet alle Event-
            relevanten Aktionen an einem Ort. QR-Scanner sehen den ganzen Block
            nicht. */}
        {/* v30.23 (F&A-Nachlieferung §1): Status der abrechnungsrelevanten
            Angaben direkt auf der Organizer-Seite — nicht erst im Wizard.
            Bewusst OHNE Status-Gate: Die Box (und die Aktionen darunter)
            müssen laut Fachkonzept auch im ENTWURF verfügbar sein, weil
            F&A die Eckdaten oft vor der Aktivierung braucht. */}
        {!isQRScannerOnlyForSelected && !selectedEvent.isDemoShowcase && parseBillingOf(selectedEvent)?.relevant === true && <BillingStatusStrip {...billingStatusStripProps} />}
        {!isQRScannerOnlyForSelected && !selectedEvent.isDemoShowcase && <AdminActionsCard {...adminActionsCardProps} />}
      </div>
      </ActionsRegistryProvider>

      {/* v30.44: Das Modal „Event-Abrechnung" stand bis hierher INNERHALB von
          `ActionsCollapsibleCard` — und die rendert ihre Kinder seit v12.7 in
          `display: none` (die ActionTiles melden sich nur noch per Context beim
          Dropdown an, gezeichnet wird dort nichts mehr). Das Modal wurde also
          erzeugt, der State kippte, der onClick lief — sichtbar wurde nie
          etwas. Genau das ist die Beobachtung aus dem Fachkonzept: „Ein Klick
          auf den Button führt aktuell jedoch zu keiner Aktion", und dasselbe
          bei der Aktion im Dropdown. Beide Einstiege waren nie kaputt, nur der
          Ort des Modals war falsch. Es gehört auf Seitenebene, zu den anderen
          Modals. */}
      {billingPanelOpen && selectedEvent && (
        <BillingActionPanel event={selectedEvent} onClose={() => setBillingPanelOpen(false)} />
      )}

      {/* Zähler + QR/Check-in Aktionen.
          v9.14: Warteliste-KPI wird nur gerendert wenn Event eine Warteliste hat.
          Sonst Grid auf 4 Spalten.
          v11.32: Bei Split-Capacity wird die separate Kapazitäts-Karten-Reihe
          unten in die „Angemeldet"-Kachel hochgezogen. Die Kachel bekommt
          dann doppelte Breite (2fr) damit Group-A/B-Breakdown sauber drin
          Platz hat — keine zwei breiten Vollbreite-Karten mehr. */}
      {/* v26: Box „Offene Fragen (User)" — zwischen Event-Infos und KPI-Kacheln.
          Zeigt die Fragen normaler User zu diesem Event (Ticketsystem). */}
      {selectedEvent && <TicketEventBox eventId={selectedEvent.id} />}
      <KpiTiles {...kpiTilesProps} />

      {/* v9.20: Check-In starten + QR-Codes versenden sind jetzt im Aktionen-Grid
          unten als ActionTile gerendert (nicht mehr als eigene Button-Reihe).
          Damit sind alle Quick-Actions an EINEM Ort zusammengefasst. Auch für
          Check-In-only-User (qrScanner-Mode) — die sehen weiterhin nur den
          Check-In-Tile, da das Aktionen-Grid für sie unten gefiltert ist. */}
      {!isQRScannerOnlyForSelected && (<>

      {/* v28.47: Hinweis-Box für Events, die im Anmeldeformular nach einer
          Unterkunft fragen, aber noch keine Hotels hinterlegt haben. Genau die
          Organizer pflegen die Zuordnung sonst weiter in Excel, weil sie nicht
          wissen, dass es das Tool gibt. Sobald das erste Hotel angelegt ist,
          verschwindet die Box wieder. */}
      {selectedEvent && selectedEvent.subsiteUrl && (isAdmin || isOrganizerFor(selectedEvent))
        && (selectedEvent.hotels || []).length === 0
        && (selectedEvent.eventSpecificFields || []).some(f =>
          /hotel|unterkunft|übernacht|übernacht|accommodation|lodging/i.test(`${f.label || ''} ${(f.options || []).join(' ')}`))
        && !hotelPanelOpen && (
        <div className="card" style={{ marginBottom: 16, borderLeft: '4px solid var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.06)' }}>
          <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start', flexWrap: 'wrap' }}>
            <div style={{ flex: 1, minWidth: 240 }}>
              <div style={{ fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', marginBottom: 4 }}>
                {isDe ? 'Dieses Event fragt nach einer Unterkunft — nutze die Hotel-Planung' : 'This event asks about accommodation — use the hotel planning'}
              </div>
              <div style={{ fontSize: '0.84rem', color: 'var(--dex-gray-700)', lineHeight: 1.55 }}>
                {isDe
                  ? <>Lege deine Hotels an, gib Zeiträume als Vorlage vor und ordne die Teilnehmer zu — mit Kontingent-Warnung, Belegung je Nacht, Rooming-Liste als Excel und personalisierter Hotel-Mail (Assistenz automatisch in Cc). Das ersetzt die Excel-Liste nebenher.</>
                  : <>Create your hotels, define stay templates and assign attendees — with capacity warnings, occupancy per night, a rooming list as Excel and a personalised hotel email (assistant auto-CC&apos;d). It replaces the spreadsheet on the side.</>}
              </div>
            </div>
            <button type="button" className="btn btn-primary" style={{ fontSize: '0.82rem', padding: '8px 16px', flexShrink: 0 }}
              onClick={() => setHotelPanelOpen(true)}>
              {isDe ? 'Hotel-Planung öffnen' : 'Open hotel planning'}
            </button>
          </div>
        </div>
      )}

      {/* ===== HOTEL-PLANUNG (v28.39, collapsible) =====
          Nur für Organizer/Admin und nur, wenn das Event eine eigene
          Teilnehmerliste hat. Standardmaessig eingeklappt — Events ohne
          Übernachtung sollen davon nichts merken. */}
      {/* v28.47: auch im Klammer-Modus. Die Übernachtung hängt an der PERSON,
          nicht am einzelnen Sub-Event — und genau eine Zeile je Person hat die
          Teilnehmerliste der Klammer (die Schattenzeilen aus v15.25). Das ist die
          richtige Ebene für die Hotel-Zuordnung; vorher war der Abschnitt bei
          Klammer-Events komplett ausgeblendet. */}
      {/* v28.90: …und nur, wenn es überhaupt eine Hotelfrage GIBT. Der
          eingeklappte Balken stand bisher unter jedem Event — auch unter einem
          zweistündigen Lunch, wo niemand übernachtet. Erkannt wird die Frage am
          Feldtyp „daterange" (Übernachtungs-Zeitraum, v28.63) oder an der
          Beschriftung eines Abfragefelds; geprüft wird das Event selbst UND
          seine Sub-Events, weil die Frage bei einer Klammer auf beiden Ebenen
          stehen kann. Ist die Planung schon im Gange (Hotels angelegt oder
          Personen zugeordnet), bleibt der Abschnitt in jedem Fall sichtbar —
          sonst verschwände eine bestehende Planung mitsamt ihrer Bedienung,
          wenn jemand das Abfragefeld nachträglich entfernt. */}
      {selectedEvent && selectedEvent.subsiteUrl && (isAdmin || isOrganizerFor(selectedEvent)) && <HotelPlanningSection {...hotelPlanningSectionProps} />}

      {/* ===== QUIZ-STATISTIK (collapsible, oberhalb Teilnehmerliste) ===== */}
      {selectedEvent && selectedEvent.quiz && selectedEvent.quiz.length > 0 && <QuizStatsSection {...quizStatsSectionProps} />}

        {/* v22.16: „Hinweise"-Box für AKTIVE Events — Pendant zur „Nächste
            Schritte"-Box bei Entwürfen. Zeigt smarte Empfehlungen (z.B.
            englischer Inhalt → Anmeldesprache fest auf Englisch stellen).
            Erscheint nur, wenn mindestens ein Hinweis zutrifft; jeder Hinweis
            ist pro Event ausblendbar (localStorage). */}
        {(isAdmin || isOrganizerFor(selectedEvent)) && !selectedEvent.isFictive && !selectedEvent.isDemoShowcase && <ActiveEventHintsBox {...activeEventHintsBoxProps} />}

      {/* v29.32: Sichtbarkeits-Zeile — wer kann das Event überhaupt sehen, und
          wer davon hat noch nicht geantwortet? Steht bewusst DIREKT über der
          Teilnehmerliste: Die Liste zeigt, wer zugesagt hat; die Frage
          „und die anderen?" stellt sich genau hier. Eingeklappt nur die Zahl,
          die Zusammensetzung erst auf Klick. */}
      {selectedEvent && <AudienceVisibilityRow {...audienceVisibilityRowProps} />}

      {/* Teilnehmerliste */}
      <div ref={participantListRef} className="card" style={{ padding: 24 }}>
        {/* v11.28: Suchfeld direkt neben dem „Teilnehmer (N)"-Header
            statt rechtsbündig — flüssiger Lese-Flow von links nach
            rechts, kein Sprung über die ganze Card-Breite mehr. */}
        <div className="mb-16" style={{ display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap' }}>
          <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
            <Users size={18} /> Teilnehmer ({isConsolidatedMode ? consolidatedFiltered.length : activeRegs.length})
            {isConsolidatedMode && (() => {
              const term = (selectedEvent && selectedEvent.childEventTermPlural) || (isDe ? 'Sub-Events' : 'sub-events');
              return (
                <span style={{ marginLeft: 8, fontSize: '0.72rem', fontWeight: 500, color: 'var(--dex-gray-500)' }}>
                  — {isDe ? 'konsolidiert über' : 'consolidated across'} {consolidatedChildren.length} {term}
                </span>
              );
            })()}
          </h3>
          <input
            type="text"
            className="form-input"
            placeholder="Teilnehmer suchen..."
            value={searchQuery}
            onChange={e => setSearchQuery(e.target.value)}
            style={{ maxWidth: 280, padding: '6px 12px', fontSize: '0.85rem' }}
          />
          {/* v29.26: Teilnehmer manuell anmelden — Ausnahme-Weg für
              Organizer. Der Hinweis auf die Selbst-Registrierung steht im
              Tooltip UND prominent im Dialog selbst. */}
          <button
            type="button"
            className="btn btn-secondary"
            style={{ fontSize: '0.8rem', padding: '6px 12px' }}
            onClick={() => setAddParticipantsOpen(true)}
            title={isDe
              ? 'Teilnehmer registrieren sich normalerweise selbst über die Anmeldeseite — dieser Weg ist für Ausnahmen (nachträgliche Zusagen, übernommene Listen).'
              : 'Attendees normally register themselves via the registration page — this path is for exceptions (late confirmations, imported lists).'}
          >
            + {isDe ? 'Teilnehmer hinzufügen' : 'Add attendees'}
          </button>
          {/* v26.44: „Matches anzeigen" — nur bei Events mit Roommate-Spalte.
              Gruppiert die Tabelle in gegenseitige Paare (Match 1, 2, …) +
              Rest-Cluster; wirkt auf die aktuell gefilterte Trefferliste. */}
          {!isConsolidatedMode && hasRoommateColumn && (
            <button
              type="button"
              className={showMatches ? 'btn btn-primary' : 'btn btn-secondary'}
              style={{ fontSize: '0.8rem', padding: '6px 12px' }}
              onClick={() => setShowMatches(v => !v)}
              title={isDe
                ? 'Gegenseitige Roommate-Auswahlen als Paare gruppiert anzeigen'
                : 'Group mutual roommate picks as pairs'}
            >
              {showMatches
                ? (isDe ? 'Matches ausblenden' : 'Hide matches')
                : (isDe ? 'Matches anzeigen' : 'Show matches')}
            </button>
          )}
        </div>
        {/* v29.26: Manuelles Anmelden — Ziel-Auswahl, Massenimport-Match,
            Mail/Outlook-Optionen, Feld-Abfrage pro Person. */}
        {selectedEvent && (
          <AddParticipantsModal
            open={addParticipantsOpen}
            onClose={() => setAddParticipantsOpen(false)}
            onDone={() => {
              // v30.67 (Review): gemeinsamer Helfer — nach einem Massen-
              // Hinzufügen ist die 429 auf dem Reload der Normalfall.
              void reloadRegistrations();
              // v30.14: Die Einzel-Anmeldungen laufen jetzt mit skipReload —
              // EIN Sammel-Refresh für Kacheln/Zähler, nicht awaiten.
              void refreshEvents().catch(() => { /* best-effort */ });
            }}
            mainEvent={parentEventForSelected || selectedEvent}
            childEvents={childEventsOf((parentEventForSelected || selectedEvent).id)}
            preselectedId={selectedEvent.id}
            searchUsers={searchUsers}
            registerForEvent={registerForEvent}
            isDe={isDe}
          />
        )}
        {/* v29.36: Schritt 1 des Nachfassens — WER fehlt noch. Personen mit Foto,
            Name und Position in Zeilen, damit man sieht, wen man anschreibt.
            Erst der Knopf unten öffnet den Mail-Dialog (Schritt 2). */}
        {pendingPeople && <PendingPeopleBox {...pendingPeopleBoxProps} />}
        {/* v15.14: Legende für die Pastel-Hintergründe — sowohl in der
            Sub-Event-Detail-Ansicht (Parent-CFs + eigene CFs) als auch im
            konsolidierten Hauptevent-View. Vorher war die Legende NUR im
            konsolidierten View sichtbar und der Organizer hat im Sub-
            Event-Tab nicht gewusst, was Pastel A vs Pastel B bedeutet. */}
        {parentEventForSelected && (
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, marginBottom: 10, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
            {((parentEventForSelected.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim()).length > 0) && (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                <span style={{ width: 14, height: 14, borderRadius: 3, background: 'rgba(0, 118, 168, 0.16)', border: '1px solid rgba(0, 118, 168, 0.3)' }} />
                {isDe ? 'Felder des Hauptevents' : 'Main-event fields'}
              </span>
            )}
            {((selectedEvent?.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim()).length > 0) && (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                <span style={{ width: 14, height: 14, borderRadius: 3, background: 'rgba(255, 191, 0, 0.22)', border: '1px solid rgba(255, 191, 0, 0.4)' }} />
                {isDe
                  ? `Felder dieses ${(selectedEvent && (selectedEvent as DeloitteEvent & { childEventTermSingular?: string }).childEventTermSingular) || 'Sub-Events'}`
                  : `Fields of this ${(selectedEvent && (selectedEvent as DeloitteEvent & { childEventTermSingular?: string }).childEventTermSingular) || 'sub-event'}`}
              </span>
            )}
          </div>
        )}

        {/* v11.70: Inline-Hinweisbox statt Modal — bei einer kürzlich
            erfolgten Abmeldung läuft die automatische Korrektur evtl. noch
            (Nachrücken + ID-Neuvergabe per Power-Automate-Batch). Solange
            sich die IDs evtl. noch verschieben, soll der Organizer nicht
            parallel manuell „IDs neu vergeben" anstoßen. */}
        <IdGapHintBox {...idGapHintBoxProps} />

        <DuplicateRegHintBox {...duplicateRegHintBoxProps} />

        <DuplicateInSubEventHintBox {...duplicateInSubEventHintBoxProps} />

        <MissingEmailHintBox {...missingEmailHintBoxProps} />

        <OverbookReviewBox {...overbookReviewBoxProps} />

        <TeamsSection {...teamsSectionProps} />

        {teamsToast && (
          <div style={{
            marginBottom: 14, padding: '10px 14px', borderRadius: 8,
            background: 'rgba(134,188,37,0.12)', border: '1px solid var(--dex-green, #86bc25)',
            color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.88rem',
          }}>
            {teamsToast}
          </div>
        )}

        {/* v23.0: Per-Team-Info-Mail. Jedes aktive Mitglied bekommt eine eigene
            Mail; pro Team trägt der Organizer eine team-spezifische Info ein
            (z.B. einen eigenen Teams-Einwahllink). */}
        {teamMailOpen && selectedEvent && <TeamMailModal {...teamMailModalProps} />}

        {/* v30.37: Fehlende Leserechte auf einzelnen Termin-Listen. Das
            gehört ÜBER die Tabelle und nicht anstelle davon — die Termine,
            die gelesen werden konnten, sind ja korrekt. Ohne diesen Hinweis
            sah ein Organizer ohne Rechte auf den Sub-Event-Subsites ein
            volles Event als leeres (jede Spalte „0"). */}
        {deniedSubEventLists.length > 0 && (() => {
          // v30.67 (Review): Die Ursache „erst nachträglich als Organizer
          // benannt" gilt nur für 401/403/404. Eine 429 oder ein Netzfehler
          // ist keine Rechtefrage — dafür hilft „Aktualisieren", nicht die
          // Reparatur-Aktion. Bei Mischung zählt die Rechte-Ursache, die
          // Statuscodes je Termin stehen dahinter.
          const permDenied = deniedSubEventLists.some(d => d.status === 401 || d.status === 403 || d.status === 404);
          const n = deniedSubEventLists.length;
          return (
          <div style={{
            border: '1px solid var(--dex-red)', background: '#fff5f5', borderRadius: 8,
            padding: '12px 14px', marginBottom: 12, fontSize: 13, lineHeight: 1.5,
          }}>
            <strong style={{ color: 'var(--dex-red)' }}>
              {permDenied
                ? (isDe ? `Kein Zugriff auf ${n} Teilnehmerliste(n)` : `No access to ${n} participant list(s)`)
                : (isDe ? `${n} Teilnehmerliste(n) konnten gerade nicht gelesen werden` : `${n} participant list(s) could not be read right now`)}
            </strong>
            <div style={{ marginTop: 6 }}>
              {permDenied
                ? (isDe
                  ? 'Die Zahlen unten sind deshalb unvollständig — betroffene Termine erscheinen mit 0 Teilnehmern, obwohl dort Anmeldungen liegen können. Grund ist fast immer, dass du erst nachträglich als Organizer benannt wurdest: Die Berechtigung wurde dann nur auf dem Haupt-Event gesetzt, nicht auf den einzelnen Terminen. Ein Admin oder der Haupt-Organizer behebt das über die Aktion „Organizer-Berechtigungen reparieren“.'
                  : 'The numbers below are therefore incomplete — affected dates show 0 participants even though registrations may exist. This almost always happens when you were named organizer after the event was created: permissions were then set on the main event only, not on the individual dates. An admin or the main organizer can fix this via the action „Repair organizer permissions“.')
                : (isDe
                  ? 'Die Zahlen unten sind deshalb unvollständig — betroffene Termine erscheinen mit 0 Teilnehmern, obwohl dort Anmeldungen liegen können. Ursache ist eine SharePoint-Drosselung oder ein Netzfehler, keine fehlende Berechtigung — bitte „Aktualisieren“ klicken.'
                  : 'The numbers below are therefore incomplete — affected dates show 0 participants even though registrations may exist. The cause is SharePoint throttling or a network error, not a missing permission — please click „Refresh“.')}
            </div>
            <div style={{ marginTop: 6, color: 'var(--dex-gray-500)' }}>
              {deniedSubEventLists.slice(0, 8).map(d => d.status > 0 ? `${d.title} (HTTP ${d.status})` : d.title).join(' · ')}
              {n > 8 ? ` … (+${n - 8})` : ''}
            </div>
          </div>
          );
        })()}
        {/* v30.67: Nachladen fehlgeschlagen — die Tabelle bleibt (alter Stand),
            der Hinweis kommt dazu. Vorher wurde die Liste bei jedem HTTP-Fehler
            eines Pushs still durch `[]` ersetzt. */}
        {regStaleHint && !regLoadError && (
          <p style={{ color: 'var(--dex-orange-dark, #b35a00)', background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 6, padding: '8px 10px', fontSize: 13, marginBottom: 12 }}>
            {regStaleHint === 'denied'
              ? (isDe
                ? 'Aktualisierung fehlgeschlagen: kein Zugriff auf die Teilnehmerliste. Angezeigt wird der zuletzt geladene Stand.'
                : 'Refresh failed: no access to the participant list. Showing the last loaded state.')
              : (isDe
                ? 'Aktualisierung fehlgeschlagen (Drosselung oder Netz). Angezeigt wird der zuletzt geladene Stand — bitte „Aktualisieren“ erneut versuchen.'
                : 'Refresh failed (throttling or network). Showing the last loaded state — please try „Refresh“ again.')}
          </p>
        )}
        {regLoadError ? (
          <p style={{ color: 'var(--dex-red)', fontStyle: 'italic' }}>
            {regLoadError === ACCESS_DENIED_MSG
              ? (isDe
                ? 'Du hast keinen Zugriff auf die Teilnehmerliste dieses Events. Das ist kein leeres Event — die Liste lässt sich mit deinem Konto nur nicht lesen. Ein Admin oder der Haupt-Organizer kann das über die Aktion „Organizer-Berechtigungen reparieren" beheben.'
                : 'You do not have access to this event’s participant list. This is not an empty event — the list simply cannot be read with your account. An admin or the main organizer can fix this via the action "Repair organizer permissions".')
              : regLoadError}
          </p>
        ) : isLoadingRegs ? (
          <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{isDe ? 'Lade Teilnehmer...' : 'Loading participants...'}</p>
        ) : isConsolidatedMode ? (
          // v14.11: konsolidierter Matrix-View für Events im
          // „Nur Sub-Events"-Modus. Eine Zeile pro eindeutigem Teilnehmer,
          // X-Spalten pro Sub-Event, plus Event-Level- (Pastel A) und
          // Sub-Event-Level- (Pastel B) Custom-Field-Spalten gruppiert.
          <ConsolidatedView {...consolidatedViewProps} />
        ) : activeRegs.length === 0 ? (
          <p style={{ color: 'var(--dex-gray-400)' }}>{isDe ? 'Noch keine Teilnehmer registriert.' : 'No participants registered yet.'}</p>
        ) : (
          <ParticipantTable {...participantTableProps} />
        )}

        {/* v17.8: Anker für Floating-Jump-Button „Zur Warteliste". */}
        <div id="admin-waitlist-anchor" style={{ scrollMarginTop: 80 }} />
        <WaitlistTables {...waitlistTablesProps} />

        {cancelledRegs.length > 0 && <CancelledList {...cancelledListProps} />}
      </div>

      {/* ===== TEILNEHMER-EDIT MODAL (v8.0) ===== */}
      {dangerZoneModal}

      {changeLogModal}

      {/* v20.7: Fortschritts-Modal der Zugriffs-Reparatur („Fremd-Anmeldungen:
          Zugriff reparieren") — Event i/N, Eintrag x/y, Balken, Abschluss-
          Summary. Während des Laufs nicht wegklickbar. */}
      {/* v28.65: Fortschritt der Namens-Reparatur. */}
      {nameFixModal && <NameFixModal {...nameFixModalProps} />}

      {accessFixModal && <AccessFixModal {...accessFixModalProps} />}

      {/* v20.2: Self-Check-in-Modal (von der QR-Kachel unter dem Event-Logo):
          großer QR, Erklärtext, PDF-/Live-Aktionen + editierbares Zeitfenster. */}
      {sciModalOpen && selectedEvent && <SelfCheckInModal {...selfCheckInModalProps} />}

      {/* v30.36: Auswahl-Modal hinter „QR-Codes und Check-In". */}
      {checkInHubOpen && selectedEvent && <CheckInHubModal {...checkInHubModalProps} />}

      {/* v9.15: QR-Code-Versand-Modal — Test (nur Organizer) / Volldurchlauf
          (alle Angemeldeten) / Auto-Send-Toggle für zukünftige Anmeldungen. */}
      {qrSendModalOpen && selectedEvent && <QrSendModal {...qrSendModalProps} />}

      {/* ===== v22.18: QR-MAIL-TEXT ANPASSEN (HtmlEditorModal mit Live-Vorschau,
          gespeichert im Event — gilt auch für den Auto-Versand) ===== */}
      {qrEditOpen && selectedEvent && <QrEditModal {...qrEditModalProps} />}

      {editingReg && selectedEvent && <EditRegModal {...editRegModalProps} />}

      {/* v19.30 (Feature A): Bearbeiten der Hauptevent-Custom-Felder einer
          konsolidierten Zeile. Schreibt in die Registrierung der Person auf
          der Hauptevent-Subsite — gleiche Persistenz wie das Teilnehmer-Edit. */}
      {/* v24.31: Teilnehmer-Detailmodal (Klick auf eine Person) — Kontakt
          (E-Mail, MS-Teams-Chat) + Detailinfos (Position/Abteilung/Unternehmen/
          Standort/Telefon/Status). Bewusst „Detailinfos", nicht die Antworten. */}
        {participantDetail && <ParticipantDetailModal {...participantDetailModalProps} />}
        {/* v24.40: „Assistenz zuordnen"-Modal. */}
        {assignAssistRow && <AssignAssistModal {...assignAssistModalProps} />}
        {mainFieldsEditReg && selectedEvent && <MainFieldsEditModal {...mainFieldsEditModalProps} />}

      {/* v19.30 (Feature B): Abmelde-Modal mit Sub-Event-Auswahl. Listet alle
          Sub-Events, für die die Person aktiv angemeldet ist, je mit Checkbox
          plus „Alle"-Schalter. Beim Bestätigen werden die gewählten Sub-Events
          abgemeldet (inkl. Mail/Outlook/Nachrücken/ID-Reorder). */}
      {deregModal && selectedEvent && <DeregModal {...deregModalProps} />}

      {/* v9.37: Vorschau-Modal für die QR-Code-Mail. Rendert das wirklich
          versendete Mail-HTML in einem sandboxed iframe — analog zur Live-
          Preview im Event-Wizard unter Kommunikation. Editieren ist hier
          NICHT vorgesehen, der Body wird zentral aus der QR-Code-Vorlage
          gebaut. */}
      {qrPreviewOpen && <QrPreviewModal {...qrPreviewModalProps} />}

      {/* Gesendete Rundmails — Kommunikations-Log (DEX_EventComms) des Events.
          Liste (neueste zuerst); Klick auf eine Zeile blendet den kompletten
          HTML-Body ein — gerendert wie die QR-Mail-Vorschau (iframe/srcDoc,
          isoliert per sandbox=""). */}
      {showCommsModal && selectedEvent && <CommsLogModal {...commsLogModalProps} />}

      {/* v17.10: Step 1 — Zielgruppen-Picker für Massenmail. Erscheint vor
          dem RichText-Editor. */}
      {massmailMode === 'pick' && selectedEvent && <MassmailPickModal {...massmailPickModalProps} />}

      {/* v17.10: Step 2 (nur für 'nachruecker') — Paste-Eingabe + Extraktion */}
      {massmailMode === 'paste' && selectedEvent && <MassmailPasteModal {...massmailPasteModalProps} />}

      {/* v30.54: Offene Aufgaben beim Veranstalter (B2Run Köln). */}
      {shirtSizeOpen && selectedEvent && (
        <ShirtSizeModal event={selectedEvent} onClose={() => setShirtSizeOpen(false)} />
      )}

      {b2runTodoOpen && selectedEvent && eventServiceRef && (
        <B2RunTodoModal
          event={selectedEvent}
          service={eventServiceRef}
          onClose={() => setB2runTodoOpen(false)}
        />
      )}

      {/* v30.48: Startnummern-Rücklauf (B2Run Köln). */}
      {bibImportOpen && selectedEvent && eventServiceRef && (
        <B2RunBibImportModal
          event={selectedEvent}
          service={eventServiceRef}
          onClose={() => setBibImportOpen(false)}
          onDone={() => { reloadRegistrationsForIdCheck().catch(() => { /* best-effort */ }); }}
        />
      )}

      {/* v17.12: Excel-Export-Zielgruppen-Picker. */}
      {excelTargetModal && selectedEvent && <ExcelTargetModal {...excelTargetModalProps} />}

      {/* ===== MASSENMAIL MODAL (HtmlEditorModal mit Toolbar, Variablen, Live-Preview) ===== */}
      {showEmailModal && selectedEvent && <MassmailComposerModal {...massmailComposerModalProps} />}

      {/* ===== EINLADUNGSMAIL MODAL (v11.40) ===== */}
      {showInviteModal && selectedEvent && <InviteComposerModal {...inviteComposerModalProps} />}

      {/* v26.98: Zuschneiden des Event-Fotos im Mail-Composer (invite/massmail). */}
      {composerCrop && (
        <ImageCropModal
          open={!!composerCrop}
          src={composerCrop === 'invite' ? inviteEventPhotoB64 : composerCrop === 'qr' ? qrEventPhotoB64 : massmailEventPhotoB64}
          isDe={isDe}
          allowAspect
          onClose={() => setComposerCrop(null)}
          onApply={(dataUrl) => {
            if (composerCrop === 'invite') setInviteEventPhotoB64(dataUrl);
            else if (composerCrop === 'qr') setQrEventPhotoB64(dataUrl);
            else setMassmailEventPhotoB64(dataUrl);
            setComposerCrop(null);
          }}
        />
      )}

      {/* ===== OUTLOOK-DECLINE-CHECK MODAL (Admin only) ===== */}
      {showDeclineModal && declineResult && <DeclineCheckModal {...declineCheckModalProps} />}
      </>)}

      {/* v11.0: Modal für Teilnehmer-Attachments. Liste der hochgeladenen
          Dateien mit Download-Link + Lösch-Button (Admin/Organizer kann
          fremde Uploads löschen). Plus optional eigener Upload-Button für
          den Admin (z.B. Bestätigungsbescheinigung im Namen des
          Teilnehmers anhängen). */}
      {attachmentsModalReg && <AttachmentsModal {...attachmentsModalProps} />}

      {/* v11.36: Fortschritts-Overlay für die ID-Neuvergabe (mit %). */}
      {reorderProgress !== null && <ReorderProgressOverlay {...reorderProgressOverlayProps} />}

      {/* v11.70: kein Modal mehr — der Hinweis wird inline über der
          Teilnehmerliste angezeigt (siehe Render-Block oberhalb der
          Teilnehmer-Tabelle). */}

      {/* v23.2: Duplikat-Abmelde-Modal — beim Abmelden einer doppelt
          angemeldeten Person fragt die App, ob die Zeile STILL entfernt werden
          soll (Duplikat löschen, ohne Mail/Outlook/Nachrücken — die Person
          bleibt über ihre zweite Zeile angemeldet) oder normal abgemeldet. */}
      {dupCancelReg && selectedEvent && <DupCancelModal {...dupCancelModalProps} />}

      {/* v11.36: Überbuchungs-Entscheidungs-Modal (Bestätigen / Platz behalten) */}
      {overbookModal && selectedEvent && <OverbookDecisionModal {...overbookDecisionModalProps} />}

      {/* v28.70: Wartelisten-Platz manuell setzen. */}
      {wlPosModal && selectedEvent && <WaitlistPositionModal {...waitlistPositionModalProps} />}

      {adminAddMemberDialog && selectedEvent && <AdminAddMemberModal {...adminAddMemberModalProps} />}
      {/* v17.8: Floating Jump-Buttons. Erscheinen sobald der User durch die
          Teilnehmer-Tabelle scrollt — sparen Zeit bei langen Listen. */}
      {/* v17.11.1: JumpButtons immer rendern wenn ein Event selektiert ist —
          das interne Show-Gating (scrollY > 300) reicht. Früher Threshold
          activeRegs>10 war zu hoch, für Test-Events mit wenig TN
          erschienen die Buttons nie. */}
      {selectedEvent && (
        <JumpButtons hasWaitlist={waitlistRegs.length > 0} />
      )}
    </div>
  );
}
