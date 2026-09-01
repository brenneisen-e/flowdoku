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
import HotelPlanningPanel from './HotelPlanningPanel';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { DeloitteEvent } from '../types';
import { buildHashDeepLink } from '../utils/deepLink';
import { SPRegistration, REG_LIST_NAME } from '../services/EventService';
import { B2RUN_KOELN_HEADERS, B2RUN_KOELN_ALTERSKLASSE, mapAnredeToB2Run, mapStarterTypeToStartblock, isB2RunKoelnTitle } from '../data/b2runKoeln';
import { Plus, Users, FileText, Trash2, Copy, Mail, Send, Download, Pencil, ExternalLink, AlertCircle, Hash, Shirt, Columns, Wrench, RefreshCw, X, Check, Link2, ChevronUp, ChevronDown, QrCode, Info, Calendar, Pin } from './Icons';
import B2RunBibImportModal from './admin/B2RunBibImportModal';
import B2RunTodoModal from './admin/B2RunTodoModal';
import ShirtSizeModal from './admin/ShirtSizeModal';
import { SHIRT_PATTERN } from '../utils/checkInExtras';
import { groupSubEventTabs, stripGroupPrefix } from '../utils/subEventGroups';
import RecipientPicker from './admin/RecipientPicker';
import OrganizerList from './OrganizerList';
import { PersonContactHover } from './PersonContactHover';
import { downloadSelfCheckInPdf } from '../utils/selfCheckInPdf';
import { isEventOver } from '../utils/eventFormat';
import { withParentTitleSubject } from '../utils/mailSubject';
import { selfCancelLocked } from '../utils/cancelPolicy';
import AddParticipantsModal from './admin/AddParticipantsModal';
import { TeamsJoinButton } from './TeamsJoinButton';
import { eventTeamsLink, locationWithoutTeamsUrl } from '../utils/teamsLink';
import { accountCheckCacheKey, invalidateInactiveAccountCache } from '../utils/accountCheckCache';
import { isDeloitteInternalEmail, isExternalEmail } from '../utils/deloitteDomain';
// v20.1: Self-Check-in jederzeit aktivierbar (Token-Erzeugung beim Klick).
// v20.2: + statische Check-in-URL für die QR-Kachel im Event-Detail.
// v20.3: + Default-Zeitfenster (2 Std. vor Start bis Event-Ende) zur Vorbelegung.
import { generateSelfCheckInToken, buildStaticCheckInUrl, defaultCheckInWindow } from '../utils/selfCheckIn';
import { useIsMobile } from '../utils/useIsMobile';
// v20.0 (Audit): xlsx + qrcode werden nicht mehr statisch importiert, sondern
// erst beim tatsächlichen Gebrauch (Export-Klick / QR-Vorschau) als eigener
// Chunk nachgeladen — spart ~1 MB im Haupt-Bundle.
import { EventService, EventCommRow } from '../services/EventService';
import { SharePointService } from '../services/SharePointService';
import { qrCodeEmail, qrEmailDefaults, buildQrBlockHtml, QrEmailOverride, cancellationEmail, promotionEmail, wrapTemplate, replacePlaceholders, buildEmailFromTemplate, getCachedLogoBase64, getCachedOrbBase64, injectIntoEmailContent, externalInvitationEmail } from '../services/EmailTemplates';
import { buildParticipantQrDataUrl } from '../utils/qrWithMark';
// v26.47: Externe Anmeldung — Einladung als .eml-Entwurf (X-Unsent) zum
// Selbst-Versenden durch die anmeldende Person (App kann keine externen
// Adressen anmailen).
import { buildUnsentEmlDraft, downloadEml } from '../utils/emlDraft';
import { getCachedImage } from '../utils/imageCache';
import { applyEventTemplateOverride, formatOrganizerList } from '../context/EventContext';
import { HtmlEditorModal } from './HtmlEditorModal';
import ImageCropModal from './ImageCropModal';
import { InfoTooltip } from './InfoTooltip';
import { MultiSelectDropdown } from './MultiSelectDropdown';
import Modal from './Modal';
// v28.95: Platzhalter für Events ohne eigenes Foto. Zuerst das im Admin
// Center unter „Logo & Branding" hinterlegte DEX-Orb (DefaultImageBase64 im
// _Config-Eintrag von DEX_EmailTemplates) — das ist die Stelle, an der es
// ausgetauscht wird, und dann soll der Tausch überall greifen. Das
// gebuendelte PNG ist nur der Rueckfall, solange der Cache noch nicht
// geladen ist (frischer Tab, erster Render).
import { DEX_ORB_PNG } from '../data/brandLogos';
import { Icon } from '@fluentui/react/lib/Icon';
import TicketEventBox from './tickets/TicketEventBox';
import InternationalSearchToggle from './InternationalSearchToggle';
import { UserFieldPicker } from './UserFieldPicker';
// v20.4: moderne Confirm-/Alert-Modals statt window.confirm/alert.
import { useDialog } from '../context/DialogContext';
import { shortSubEventTitle } from '../utils/subEventTitle';
import { formatDate, getStatusColor, localizeStatus, stripHtmlToText, looksEnglishText, translateStatus } from '../utils/eventStatus';
import { formatAllDayPeriod } from '../utils/eventFormat';
import { getBlockedInviteRecipients } from '../utils/inviteGuards';
import { ActionTile, SplitMergeToggle, ActionsCollapsibleCard, ActionsRegistryProvider, ActionsDropdown } from './admin/ActionsMenu';
import BillingActionPanel from './admin/BillingActionPanel';
import { parseBillingOf, missingBillingFields, faStatusOf, FA_STATUS_LABELS, FA_STATUS_COLORS, FA_STATUS_NEXT } from '../utils/faBilling';
import { BILLING_FIELDS, canEditBilling } from '../data/billingFields';
import MailHeaderImageChooser from './admin/MailHeaderImageChooser';
import { MailHeaderImage, MAIL_HEADER_IMAGE_DEFAULT, mailHeaderOpts, applyHeroImage, hasOwnHeaderImage, normalizeMailHeaderImage, isDefaultMailHeaderImage } from '../utils/mailHeaderImage';










// v30.66: Modul-Ebene (Konstanten, Styles, Typen, Kopfbild-Helfer) liegt jetzt
// in components/admin/adminConstants|adminStyles|adminTypes — sie kennt den
// State nicht und war restlos abtrennbar.
import { DUP_ACTIVE_STATI, SAMPLE_QR_ID, ACCESS_DENIED_MSG, eventHeaderImageLayout } from './admin/adminConstants';
import { qrDisclosureStyle, hubChoiceStyle } from './admin/adminStyles';
import { ConsolidatedRow, AudiencePerson, MassmailAudience } from './admin/adminTypes';
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
        try {
          const regs = await getAllRegistrations(selectedEvent.id);
          setRegistrations(regs);
        } catch { /* */ }
      }
    } finally { setIsRefreshing(false); }
  };
  const { currentUser } = useCurrentUser();
  // v27.11: getGroupMembers für die Verteiler-Auflösung der Einladungs-Mail.
  const { isAdmin, siteUrl, currentUserRole, searchUser, searchUsers, getGroupMembers, isImpersonating, isFA } = useRoles();
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  // v20.4: App-Modals statt nativer Browser-Dialoge.
  const { confirmDialog, showAlert } = useDialog();
  const [selectedEvent, setSelectedEvent] = React.useState<DeloitteEvent | null>(null);
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
      getAllRegistrations(ev.id).then(r => { if (!cancelled) setRegistrations(r); }).catch(() => { /* */ });
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

  /**
   * v29.44: Abmelde-Mail bauen — für ALLE Organizer-Wege gleich.
   *
   * Vorher nahm das Organizer Center überall `cancellationEmail(...)`, also den
   * fest eingebauten Standardtext. Der Selbst-Abmelde-Weg des Teilnehmers löst
   * dagegen seit jeher die gepflegte Vorlage auf (SharePoint-Template +
   * Event-Override). Dieselbe Abmeldung sah damit unterschiedlich aus, je
   * nachdem, WER sie ausgelöst hat — und ein Organizer, der den Text seines
   * Sub-Events angepasst hatte, bekam ihn nie zu sehen.
   */
  const buildCancellationMail = async (
    ev: DeloitteEvent,
    reg: SPRegistration,
    fullName: string,
  ): Promise<{ subject: string; body: string }> => {
    const fallback = cancellationEmail(fullName, ev.title);
    if (!eventServiceRef) return fallback;
    try {
      const lang = ev.emailLanguage || 'EN';
      const spTplRaw = await eventServiceRef.getEmailTemplate('Abmeldung', lang).catch(() => null);
      const spTpl = applyEventTemplateOverride(spTplRaw, ev.emailTemplateOverrides, 'Abmeldung');
      if (!spTpl) return fallback;
      return buildEmailFromTemplate(spTpl, {
        Name: (reg.Vorname || '').trim() || fullName,
        EventTitle: ev.title,
        AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      });
    } catch { return fallback; }
  };
  const performStandardCancel = async (reg: SPRegistration): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
    const eventWasOver = isEventOver(selectedEvent);
    setAdminToast({ kind: 'cancelling', name });
    const cancelledStarterType = reg.StarterType || '';
    await eventServiceRef.cancelRegistration(selectedEvent.subsiteUrl, reg.Id, `${currentUser.firstName} ${currentUser.surname}`.trim(), currentUser.email);
    // v29.31: Die „Konto inaktiv"-Prüfung ist 24 h gecacht. Ohne Verwerfen
    // meldete die Sammel-Box auf der Startseite die eben abgemeldete Person
    // bis zum nächsten Tag weiter als offenen Fall.
    invalidateInactiveAccountCache([selectedEvent.id, selectedEvent.parentEventId || '']);
    if (reg.ParticipantEmail && !eventWasOver) {
      if (!selectedEvent.disableEmails && !selectedEvent.disableCancellationEmail) {
        const emailData = await buildCancellationMail(selectedEvent, reg, name);
        eventServiceRef.queueEmail(
          emailData.subject, reg.ParticipantEmail, name, emailData.body,
          'Abmeldung', selectedEvent.title, selectedEvent.id
        ).catch(err => console.warn('[DEX]', err));
      }
      if (!selectedEvent.disableOutlook) {
        eventServiceRef.queueOutlookEvent(
          reg.ParticipantEmail, selectedEvent.id, selectedEvent.title, 'Ausladen'
        ).catch(err => console.warn('[DEX]', err));
      }
    }
    if (reg.ParticipantEmail && selectedEvent.eventNumber) {
      eventServiceRef.removeParticipantEvent(
        reg.ParticipantEmail, selectedEvent.eventNumber
      ).catch(err => console.warn('[DEX]', err));
    }
    const isSplitEvent = typeof selectedEvent.durchstarterCapacity === 'number'
      && typeof selectedEvent.funstarterCapacity === 'number'
      && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
    const useTypeFilter = isSplitEvent && !selectedEvent.splitSharedWaitlist;
    // v27.11: WaitlistEnabled=false ist jetzt ein echter Kill-Switch — kein
    // automatisches Nachrücken mehr, wenn der Organizer die Warteliste
    // abgeschaltet hat (manuelles Nachrücken über den Admin-Button bleibt
    // als bewusster Override möglich).
    if (!eventWasOver && selectedEvent.waitlistEnabled !== false) {
      try {
        const promoted = await eventServiceRef.promoteFirstWaitlistItem(
          selectedEvent.subsiteUrl,
          cancelledStarterType || undefined,
          selectedEvent.maxParticipants,
          (useTypeFilter && cancelledStarterType) ? cancelledStarterType : undefined,
          { itemId: reg.Id, participantEmail: reg.ParticipantEmail || '' },
        );
        if (promoted && promoted.success && promoted.email) {
          setAdminToast({ kind: 'promoted', name: promoted.name || promoted.email, email: promoted.email, type: cancelledStarterType || undefined });
          if (!selectedEvent.disableEmails) {
            try {
              const lang = selectedEvent.emailLanguage || 'EN';
              const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
              const promoteVars = {
                Name: promotedFirstName,
                EventTitle: selectedEvent.title,
                Organizer: formatOrganizerList(selectedEvent.organizers, lang),
                AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
                WaitlistPosition: '',
              };
              let emailData: { subject: string; body: string };
              const spTplRaw = await eventServiceRef.getEmailTemplate('Nachruecken', lang).catch(() => null);
              const spTpl = applyEventTemplateOverride(spTplRaw, selectedEvent.emailTemplateOverrides, 'Nachruecken');
              if (spTpl) { emailData = buildEmailFromTemplate(spTpl, promoteVars); }
              else { emailData = promotionEmail(promotedFirstName, selectedEvent.title); }
              await eventServiceRef.queueEmail(
                withParentTitleSubject(emailData.subject, selectedEvent.parentEventId ? allEvents.find(e => e.id === selectedEvent.parentEventId) : undefined),
                promoted.email, promoted.name || '', emailData.body,
                'Nachruecken', selectedEvent.title, selectedEvent.id
              );
            } catch (err) { console.warn('[DEX] promote-email failed:', err); }
          }
          if (!selectedEvent.disableOutlook) {
            try { await eventServiceRef.queueOutlookEvent(promoted.email, selectedEvent.id, selectedEvent.title, 'Einladen'); }
            catch (err) { console.warn('[DEX] promote-outlook failed:', err); }
          }
        } else {
          setAdminToast({ kind: 'no-promote', name });
        }
      } catch (err) {
        console.warn('[DEX] promoteFirstWaitlistItem failed:', err);
        setAdminToast({ kind: 'no-promote', name });
      }
    }
    if (selectedEvent.subsiteUrl && !eventWasOver) {
      try {
        const ok = await eventServiceRef.queueIDReorder(
          selectedEvent.id, selectedEvent.eventNumber || 0,
          selectedEvent.subsiteUrl, selectedEvent.title, name, reg.ParticipantEmail || undefined
        );
        if (!ok) {
          console.warn('[DEX] queueIDReorder returned false');
          showAlert(isDe ? 'Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.' : 'Cancellation successful, but the ID reorder entry could not be written to the queue. Please click "Reassign IDs" once.');
        }
      } catch (err) {
        console.warn('[DEX] queueIDReorder threw:', err);
        showAlert('Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.');
      }
    }
    const regs = await getAllRegistrations(selectedEvent.id);
    setRegistrations(regs);
  };

  // v23.2: Stilles Löschen einer doppelten Anmelde-Zeile. Anders als die
  // normale Abmeldung wird die Zeile HART gelöscht (kein „Abgemeldet"-Status,
  // der die Abmeldungs-Liste aufblähen würde) und es laufen KEINE Seiteneffekte
  // (keine Abmelde-Mail, kein Outlook-Ausladen, kein Nachrücken, kein
  // ID-Reorder, kein DEX_Participants-Cleanup) — die Person bleibt über ihre
  // andere Zeile regulär angemeldet. Sitzplatz-Counter wird nachgezogen.
  const performSilentDuplicateDelete = async (reg: SPRegistration): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
    setAdminToast({ kind: 'cancelling', name });
    try {
      await eventServiceRef.deleteRegistration(selectedEvent.subsiteUrl, reg.Id);
      try {
        await eventServiceRef.writeChangeLog({
          action: 'RegistrationDeleted',
          targetType: 'Participant',
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          targetId: ((reg as any).ParticipantEmail || '') + '#' + reg.Id,
          targetName: name,
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { note: 'Doppel-Anmeldung still entfernt (Duplikat). Person bleibt über die zweite Zeile angemeldet.' },
        });
      } catch (err) { console.warn('[DEX] writeChangeLog (dup delete) failed:', err); }
      try {
        const isSplit = typeof selectedEvent.durchstarterCapacity === 'number'
          && typeof selectedEvent.funstarterCapacity === 'number'
          && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
        await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit });
      } catch { /* best-effort */ }
    } catch (err) {
      console.warn('[DEX] performSilentDuplicateDelete failed:', err);
    }
    const regs = await getAllRegistrations(selectedEvent.id);
    setRegistrations(regs);
    setAdminToast(null);
  };
  // v28.23: Doppelte Klammer-Schatten-Zeilen in einem Rutsch bereinigen.
  // Diese Zeilen tauchen in der konsolidierten Klammer-Tabelle NICHT auf (dort
  // steht pro Person genau eine Zeile, aggregiert über die Sub-Events) — ohne
  // diesen Knopf käme der Organizer also gar nicht an sie heran. Behalten wird
  // je Person die Zeile mit den meisten ausgefüllten Hauptevent-Antworten
  // (Tie-Break: die älteste), alle weiteren werden still gelöscht: keine Mail,
  // kein Outlook, kein Nachrücken.
  const [shadowDupBusy, setShadowDupBusy] = React.useState(false);
  // v28.23: Abgleich des zentralen Teilnehmer-Registers (DEX_Participants).
  const [isSyncingRegistry, setIsSyncingRegistry] = React.useState(false);
  const [syncRegistryResult, setSyncRegistryResult] = React.useState<string | null>(null);
  const cleanupShadowDuplicates = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl || shadowDupBusy) return;
    const groups: SPRegistration[][] = [];
    duplicateEmails.forEach(em => {
      const rows = registrations.filter(r => DUP_ACTIVE_STATI.indexOf(r.Status || '') >= 0
        && (r.ParticipantEmail || '').trim().toLowerCase() === em);
      if (rows.length > 1) groups.push(rows);
    });
    if (groups.length === 0) return;
    const extra = groups.reduce((n, rows) => n + rows.length - 1, 0);
    const ok = await confirmDialog(
      isDe
        ? `${extra} doppelte Klammer-Zeile(n) bei ${groups.length} Person(en) entfernen?\n\nJe Person bleibt die Zeile mit den meisten ausgefüllten Hauptevent-Antworten erhalten. Die Entfernung läuft still — ohne Abmelde-Mail, ohne Outlook-Absage, ohne Nachrücken. Die Anmeldungen in den Sub-Events bleiben unberührt.`
        : `Remove ${extra} duplicate overall-event row(s) for ${groups.length} person(s)?\n\nFor each person the row with the most main-event answers is kept. Removal is silent — no cancellation email, no Outlook removal, no waitlist promotion. The sub-event registrations are untouched.`,
      { danger: true, confirmLabel: isDe ? 'Zeilen entfernen' : 'Remove rows' },
    );
    if (!ok) return;
    setShadowDupBusy(true);
    const answerScore = (r: SPRegistration): number => {
      try {
        const o = JSON.parse(r.CustomData || '{}') as Record<string, unknown>;
        return Object.keys(o).filter(k => String(o[k] === null || o[k] === undefined ? '' : o[k]).trim()).length;
      } catch { return 0; }
    };
    let removed = 0;
    for (const rows of groups) {
      const sorted = rows.slice().sort((a, b) => (answerScore(b) - answerScore(a)) || (a.Id - b.Id));
      for (const r of sorted.slice(1)) {
        try {
          await eventServiceRef.deleteRegistration(selectedEvent.subsiteUrl, r.Id);
          removed += 1;
          try {
            await eventServiceRef.writeChangeLog({
              action: 'RegistrationDeleted',
              targetType: 'Participant',
              targetId: `${r.ParticipantEmail || ''}#${r.Id}`,
              targetName: (r.Vorname && r.Nachname) ? `${r.Vorname} ${r.Nachname}` : (r.ParticipantName || ''),
              eventId: selectedEvent.id,
              eventTitle: selectedEvent.title,
              details: { note: 'Doppelte Klammer-Schatten-Zeile still entfernt (v28.23). Sub-Event-Anmeldungen unberührt.' },
            });
          } catch { /* best-effort */ }
        } catch (err) { console.warn('[DEX] cleanupShadowDuplicates failed for item', r.Id, err); }
      }
    }
    try {
      const isSplit = typeof selectedEvent.durchstarterCapacity === 'number'
        && typeof selectedEvent.funstarterCapacity === 'number'
        && ((selectedEvent.durchstarterCapacity || 0) > 0 || (selectedEvent.funstarterCapacity || 0) > 0);
      await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit });
    } catch { /* best-effort */ }
    try {
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch { /* */ }
    setShadowDupBusy(false);
    showAlert(
      isDe ? `${removed} doppelte Klammer-Zeile(n) entfernt.` : `${removed} duplicate overall-event row(s) removed.`,
      { variant: 'success' },
    );
  };
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
  const [deniedSubEventLists, setDeniedSubEventLists] = React.useState<string[]>([]);
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
      const denied: string[] = [];
      for (const ch of children) {
        try {
          const regs = await getAllRegistrations(ch.id, st => {
            if (st === 401 || st === 403 || st === 404 || st === 0) denied.push(ch.title || ch.id);
          });
          map[ch.id] = regs;
        } catch {
          map[ch.id] = [];
          denied.push(ch.title || ch.id);
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
      let externalCount = 0;
      try { externalCount = await countExternalRegistrations(confirmDeleteEvent); } catch { externalCount = 0; }
      if (cancelled) return;
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
  const [invitedLc, setInvitedLc] = React.useState<Set<string> | null>(null);
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

  // Admin-Toast für Abmelde-/Nachrück-Feedback (seit v6.8):
  //  - 'cancelling': während die Abmeldung + Nachrück-Suche läuft (orange, Spinner)
  //  - 'promoted'  : erfolgreicher Nachrücker mit Namen + Typ (grün)
  //  - 'no-promote': Abmeldung ok, aber keiner auf der Warteliste (grau)
  type AdminToast =
    | { kind: 'cancelling'; name: string }
    | { kind: 'promoted'; name: string; email: string; type?: string }
    | { kind: 'no-promote'; name: string };
  const [adminToast, setAdminToast] = React.useState<AdminToast | null>(null);

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
  const openEditModal = (reg: SPRegistration): void => {
    setEditError('');
    setEditingReg(reg);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const r = reg as any;
    const initial: Record<string, string> = {
      Anrede: r.Anrede || '',
      Vorname: r.Vorname || '',
      Nachname: r.Nachname || '',
      ParticipantEmail: r.ParticipantEmail || '',
      Phone: r.Phone || '',
      Department: r.Department || '',
      Location: r.Location || '',
      JobTitle: r.JobTitle || '',
      Status: r.Status || '',
      // v10.13+: B2Run-Felder ins Edit-Form aufnehmen damit das B2Run-Modul
      // im Edit-Modal die aktuellen Werte vorbefüllen kann. Strings (auch
      // wenn leer) — bei Nicht-B2Run-Events werden die Felder im Modal
      // sowieso nicht angezeigt.
      StarterType: r.StarterType || '',
      PreferredStarterType: r.PreferredStarterType || '',
    };
    // Custom-Field-Werte aus dem reg laden (sie sind als SP-Spalten gespeichert)
    if (selectedEvent?.eventSpecificFields) {
      for (const f of selectedEvent.eventSpecificFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = (f as any).spInternalName || '';
        if (sp) initial[sp] = (r[sp] !== undefined && r[sp] !== null) ? String(r[sp]) : '';
      }
    }
    setEditForm(initial);
  };
  const closeEditModal = (): void => {
    setEditingReg(null);
    setEditForm({});
    setEditError('');
  };
  const saveEdit = async (): Promise<void> => {
    if (!editingReg || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsSavingEdit(true);
    setEditError('');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const r = editingReg as any;

      // Stamm-Daten (Vorname, Nachname, E-Mail) werden seit v9.7 ebenfalls
      // editierbar gemacht — z.B. um Tippfehler nach manueller Anlage zu
      // korrigieren. Validierung:
      //   1. E-Mail muss eine Deloitte-Deutschland-Adresse sein (@deloitte.de).
      //      Die Plattform ist nur für DEALL freigeschaltet — auch @deloitte.com
      //      (US/Global) zählt als extern. Sonst Abbruch mit Fehler.
      //   2. Person muss in M365 existieren (searchUserByEmail). Sonst
      //      Abbruch mit "Tippfehler"-Hinweis.
      // Die uebrigen Profil-Felder (Phone, Department, Location, JobTitle)
      // bleiben read-only — sie kommen aus dem M365-Profil.
      const oldVorname = String(r.Vorname || '');
      const oldNachname = String(r.Nachname || '');
      const oldEmail = String(r.ParticipantEmail || '');
      const newVorname = (editForm.Vorname || '').trim();
      const newNachname = (editForm.Nachname || '').trim();
      const newEmail = (editForm.ParticipantEmail || '').trim();
      const stammChanged = newVorname !== oldVorname || newNachname !== oldNachname || newEmail !== oldEmail;

      const profileFields: { Department?: string; Location?: string; JobTitle?: string } = {};
      if (stammChanged) {
        // Plausibilität: nicht-leer
        if (!newVorname || !newNachname || !newEmail) {
          setEditError(isDe
            ? 'Vorname, Nachname und E-Mail dürfen nicht leer sein.'
            : 'First name, last name and email must not be empty.');
          return;
        }
        // Domain-Check: nur Deloitte-Adressen zulassen (v27.11: beliebige
        // Member-Firm-Domain, konsistent zur International-Suche v26.57).
        const lower = newEmail.toLowerCase();
        const isDeloitte = isDeloitteInternalEmail(lower);
        if (!isDeloitte) {
          setEditError(isDe
            ? `Externe E-Mail-Adresse — nicht erlaubt. Bitte eine Deloitte-Adresse verwenden (z.B. @deloitte.de oder eine andere Member-Firm-Domain).`
            : `External email address — not allowed. Please use a Deloitte mailbox (e.g. @deloitte.de or another member-firm domain).`);
          return;
        }
        // Existenz-Check via M365-Profile (UPN!=SMTP-aware). Wenn wir hier
        // nichts finden, ist es entweder ein Tippfehler oder ein Account
        // der gar nicht (mehr) im Tenant ist — beides nicht akzeptabel.
        if (newEmail.toLowerCase() !== oldEmail.toLowerCase()) {
          const profile = await searchUser(newEmail);
          if (!profile || !profile.displayName) {
            setEditError(isDe
              ? `Person mit E-Mail "${newEmail}" wurde im Deloitte-Tenant nicht gefunden. Bitte Adresse prüfen (Tippfehler?).`
              : `No person found in the Deloitte tenant for "${newEmail}". Please check the address (typo?).`);
            return;
          }
          // Profil-Daten gleich mit-übernehmen, damit der Eintrag konsistent
          // bleibt (Department / Location / JobTitle passen zum neuen User).
          profileFields.Department = ''; // searchUser liefert displayName/location/jobTitle
          profileFields.Location = profile.location || '';
          profileFields.JobTitle = profile.jobTitle || '';
        }
      }

      const oldValues: Record<string, unknown> = {};
      const patch: Record<string, unknown> = {};
      const fieldLabelMap: Record<string, string> = {};

      if (stammChanged) {
        if (newVorname !== oldVorname) {
          oldValues.Vorname = oldVorname; patch.Vorname = newVorname; fieldLabelMap.Vorname = isDe ? 'Vorname' : 'First name';
        }
        if (newNachname !== oldNachname) {
          oldValues.Nachname = oldNachname; patch.Nachname = newNachname; fieldLabelMap.Nachname = isDe ? 'Nachname' : 'Last name';
        }
        if (newEmail !== oldEmail) {
          oldValues.ParticipantEmail = oldEmail; patch.ParticipantEmail = newEmail; fieldLabelMap.ParticipantEmail = 'E-Mail';
          // Profil-Daten mit aktualisieren (nur wenn überhaupt was zurückkam)
          if (profileFields.Location) {
            oldValues.Location = String(r.Location || ''); patch.Location = profileFields.Location;
            fieldLabelMap.Location = isDe ? 'Standort' : 'Location';
          }
          if (profileFields.JobTitle) {
            oldValues.JobTitle = String(r.JobTitle || ''); patch.JobTitle = profileFields.JobTitle;
            fieldLabelMap.JobTitle = 'Job Title';
          }
        }
      }

      // Custom-Felder des Events. v10.15+: nur Felder ins Patch aufnehmen die
      // sich tatsächlich geändert haben — sonst sendet ein unverändertes
      // Choice-Feld ohne ausgewählten Wert einen leeren String an SP, der
      // mit HTTP 400 'Invalid choice' kippt und das ganze Update abbricht.
      if (selectedEvent?.eventSpecificFields) {
        for (const f of selectedEvent.eventSpecificFields) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const sp = (f as any).spInternalName || '';
          if (!sp) continue;
          const oldVal = r[sp] !== undefined && r[sp] !== null ? String(r[sp]) : '';
          const newVal = editForm[sp] || '';
          if (newVal === oldVal) continue;  // unverändert → skip
          fieldLabelMap[sp] = f.label;
          oldValues[sp] = oldVal;
          patch[sp] = newVal;
        }
      }

      // v10.13+: B2Run-Felder explizit ins Patch aufnehmen, wenn sich was
      // geändert hat. Sind keine regulären customFields, daher werden sie
      // oben in der eventSpecificFields-Loop nicht abgeholt. Nur bei
      // Split-Capacity-Events relevant.
      const isSplitEvent = !!selectedEvent
        && (selectedEvent.durchstarterCapacity || 0) > 0
        && (selectedEvent.funstarterCapacity || 0) > 0;
      if (isSplitEvent) {
        const oldStarter = String(r.StarterType || '');
        const newStarter = (editForm.StarterType || '').trim();
        if (newStarter !== oldStarter) {
          oldValues.StarterType = oldStarter;
          patch.StarterType = newStarter;
          fieldLabelMap.StarterType = isDe ? 'Starter-Typ' : 'Starter type';
        }
        const oldPref = String(r.PreferredStarterType || '');
        const newPref = (editForm.PreferredStarterType || '').trim();
        if (newPref !== oldPref) {
          oldValues.PreferredStarterType = oldPref;
          patch.PreferredStarterType = newPref;
          fieldLabelMap.PreferredStarterType = isDe ? 'Wunsch-Starter-Typ' : 'Preferred starter type';
        }
      }
      if (Object.keys(patch).length === 0) {
        // Keine Änderung — nichts zu tun.
        closeEditModal();
        return;
      }
      const actor = {
        name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email,
        email: currentUser.email,
      };
      const ok = await eventServiceRef.adminUpdateRegistration(
        selectedEvent.subsiteUrl, editingReg.Id, patch, actor, oldValues, fieldLabelMap
      );
      if (!ok) {
        // Häufigste 400-Ursache: eine SP-Spalte aus dem Patch existiert nicht
        // auf dieser Teilnehmerliste (z.B. StarterType auf einem v9-Event ohne
        // B2Run-Schema, oder ein neu hinzugefügtes Custom-Field ohne 'Spalten
        // fixen'-Run). Hilfreicher Hinweis auf den Repair-Button.
        setEditError(isDe
          ? 'Speichern fehlgeschlagen — vermutlich fehlt eine SP-Spalte in der Teilnehmerliste. Klicke einmal „Spalten fixen" im Toolbox-Bereich des Events, dann erneut versuchen.'
          : 'Save failed — likely a missing SP column on the participant list. Click „Fix columns" in the event toolbox once, then retry.');
        return;
      }
      // v9.0: Audit-Log mit Diff der geänderten Felder
      try {
        const changes: Record<string, { old: unknown; new: unknown }> = {};
        for (const k of Object.keys(patch)) {
          if (oldValues[k] !== patch[k]) changes[k] = { old: oldValues[k], new: patch[k] };
        }
        await eventServiceRef.writeChangeLog({
          action: 'ParticipantUpdated',
          targetType: 'Participant',
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          targetId: ((editingReg as any).ParticipantEmail || '') + '#' + editingReg.Id,
          targetName: `${editingReg.Vorname || ''} ${editingReg.Nachname || ''}`.trim(),
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { changes },
        });
      } catch { /* */ }
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
      closeEditModal();
    } catch (err) {
      console.warn('[DEX] saveEdit error:', err);
      setEditError(isDe
        ? 'Unerwarteter Fehler beim Speichern.'
        : 'Unexpected error while saving.');
    } finally {
      setIsSavingEdit(false);
    }
  };

  // v19.30 — Feature A: Edit-Modal für die Hauptevent-Custom-Felder einer
  // konsolidierten Zeile öffnen. Die Antworten stehen in der Registrierung der
  // Person auf der Hauptevent-Subsite. Wir suchen sie per E-Mail in
  // `registrations` (das ist die Teilnehmerliste des selektierten Hauptevents).
  // v24.38: Fehlende Klammer-/Hauptevent-Anmeldung nachtragen. Eine Person, die
  // nur in Sub-Events angemeldet ist, aber keine Schatten-Zeile auf der
  // Klammer-Teilnehmerliste hat (Daten-Anomalie), wird hier vom Organizer
  // händisch ergänzt. Das ist genau die „Schatten-Registrierung" des
  // subEventsOnlyMode — `registerForEvent` auf das Klammer-Event unterdrückt
  // bei subEventsOnlyMode automatisch Mail + Outlook (suppressParentNotifications).
  // v30.14: Kern ohne Dialog/Alerts/Reload — wird vom Einzel-Knopf UND vom
  // Sammel-Fix („Alle still nachtragen") benutzt. Liefert true bei Erfolg.
  const addToKlammerCore = async (row: ConsolidatedRow): Promise<boolean> => {
    if (!selectedEvent) return false;
    const name = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email;
    try {
      // v24.39: Realen Registranten aus den Sub-Event-Zeilen ableiten (wer die
      // Sub-Events angemeldet hat). Die Klammer-Schatten-Zeile wird DEMSELBEN
      // zugeschrieben — NICHT dem Admin, der nur die Datenkorrektur macht.
      // Dadurch (1) taucht die Zeile nicht fälschlich im „Assistenz" des Admins
      // auf und (2) sieht die echte Assistenz die Klammer-Anmeldung in IHRER
      // „Assistenz"-Kachel und kann die Klammer-Felder dort pflegen.
      let realByEmail = '';
      let realByName = '';
      const subRegs = (Object.values(row.perChild).filter(Boolean) as SPRegistration[])
        .slice()
        .sort((a, b) => {
          const ta = a.RegistrationDate ? new Date(a.RegistrationDate).getTime() : Number.POSITIVE_INFINITY;
          const tb = b.RegistrationDate ? new Date(b.RegistrationDate).getTime() : Number.POSITIVE_INFINITY;
          return ta - tb;
        });
      for (const r of subRegs) {
        if ((r.RegisteredByEmail || '').trim()) { realByEmail = (r.RegisteredByEmail || '').trim(); realByName = (r.RegisteredByName || '').trim(); break; }
      }
      // Fallback: keine Stellvertreter-Info auf den Subs → als Selbst-Anmeldung
      // behandeln (Schatten gehört dann der Person selbst, nicht dem Admin).
      if (!realByEmail) { realByEmail = row.email; realByName = name; }

      // v30.14: skipReload — der Kern läuft im Sammel-Fix in Serie; der eine
      // Refresh kommt vom Aufrufer, sonst zöge jede Person einen loadEvents.
      // v30.57: KEIN `proxyConsentConfirmed` mehr.
      //
      // Das Nachtragen einer fehlenden Klammer-Zeile ist eine Datenkorrektur —
      // niemand hat dabei jemanden um Zustimmung gefragt. Das Flag schrieb
      // aber genau das in die Spalte `ProxyConsent`: „Zustimmung der Person
      // zur stellvertretenden Anmeldung bestätigt durch <Admin> am <Datum>".
      // Zusammen mit dem `RegisteredBy`-Rückschreiben zwei Zeilen weiter unten
      // entstand ein Datensatz, der sich selbst widerspricht: angemeldet von
      // der Person selbst, Zustimmung bestätigt durch jemand anderen.
      //
      // Ein erfundener Zustimmungsnachweis ist die unangenehmste Sorte
      // falscher Daten — er sieht aus wie ein Beleg. Das Flag steuert
      // ausschließlich diesen Text (s. EventContext, `proxyConsentStr`) und
      // ist KEIN Rechte-Schalter; ohne es bleibt die Spalte leer, und wer die
      // Korrektur ausgelöst hat, steht ohnehin im ChangeLog-Eintrag unten.
      // `actorAllowedAsAssistant` bleibt — das ist die Rechte-Seite.
      const res = await registerForEvent(
        selectedEvent.id, {}, row.vorname || '', row.nachname || '', row.email, undefined,
        { suppressMail: true, suppressOutlook: true, actorAllowedAsAssistant: true, skipReload: true }
      );
      if (res && res.ok) {
        const regs = await getAllRegistrations(selectedEvent.id);
        // Schatten-Zeile dem realen Registranten zuschreiben (registerForEvent
        // hat den eingeloggten Admin als RegisteredBy gesetzt).
        const newParent = regs.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey);
        if (newParent && eventServiceRef && selectedEvent.subsiteUrl
          && realByEmail.toLowerCase() !== (currentUser.email || '').toLowerCase()) {
          try {
            await eventServiceRef.adminUpdateRegistration(
              selectedEvent.subsiteUrl, newParent.Id,
              { RegisteredByEmail: realByEmail, RegisteredByName: realByName },
              { name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email, email: currentUser.email }
            );
          } catch { /* best-effort */ }
        }
        if (eventServiceRef) {
          try {
            await eventServiceRef.writeChangeLog({
              action: 'ParticipantUpdated',
              targetType: 'Participant',
              targetId: (row.email || '') + '#klammer',
              targetName: name,
              eventId: selectedEvent.id,
              eventTitle: selectedEvent.title,
              details: { scope: 'addedMissingKlammerRegistration', actorEmail: currentUser.email, attributedTo: realByEmail },
            });
          } catch { /* */ }
        }
        return true;
      }
      return false;
    } catch (err) {
      console.warn('[DEX] addToKlammer error:', err);
      return false;
    }
  };

  const addToKlammer = async (row: ConsolidatedRow): Promise<void> => {
    if (!selectedEvent) return;
    const name = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email;
    if (!(await confirmDialog(
      isDe
        ? `Fehlende Hauptanmeldung: „${name}" ist nur in Sub-Events angemeldet, fehlt aber am Klammer-Event „${selectedEvent.title}".\n\nDie fehlende Klammer-Anmeldung jetzt ergänzen? (Es wird KEINE Mail und KEIN Outlook-Termin versendet — reine Datenkorrektur.)`
        : `Missing main registration: „${name}" is only in sub-events but missing on the umbrella event „${selectedEvent.title}".\n\nAdd the missing umbrella registration now? (No email and no Outlook invite are sent — data correction only.)`,
      { confirmLabel: isDe ? 'Hinzufügen' : 'Add' }
    ))) return;
    setAddingToKlammer(row.emailKey);
    try {
      const ok = await addToKlammerCore(row);
      if (ok) {
        const regs2 = await getAllRegistrations(selectedEvent.id);
        setRegistrations(regs2);
        showAlert(isDe ? `„${name}" wurde zum Klammer-Event hinzugefügt.` : `„${name}" was added to the umbrella event.`, { variant: 'success' });
      } else {
        showAlert(isDe ? 'Hinzufügen fehlgeschlagen — bitte erneut versuchen.' : 'Adding failed — please try again.', { variant: 'error' });
      }
    } finally {
      setAddingToKlammer(null);
    }
  };

  // v30.14: Sammel-Fix — ALLE fehlenden Klammer-Anmeldungen still nachtragen.
  // Sequentiell mit Fehlerzähler (CLAUDE.md-Regel: prüfbar, kein Promise.all-
  // Feuerwerk unter Drosselung); jede Zeile läuft über denselben Kern wie der
  // Einzel-Knopf (inkl. Zuschreibung an den realen Registranten + ChangeLog).
  const addAllToKlammer = async (rows: ConsolidatedRow[]): Promise<void> => {
    if (!selectedEvent || rows.length === 0 || bulkKlammerProgress) return;
    if (!(await confirmDialog(
      isDe
        ? `Alle ${rows.length} fehlenden Klammer-Anmeldungen jetzt still nachtragen? (Es wird KEINE Mail und KEIN Outlook-Termin versendet — reine Datenkorrektur. Die Zeilen werden der jeweils anmeldenden Person zugeschrieben.)`
        : `Add all ${rows.length} missing umbrella registrations silently now? (No email and no Outlook invite are sent — data correction only. Rows are attributed to whoever registered the sub-events.)`,
      { confirmLabel: isDe ? `Alle ${rows.length} nachtragen` : `Add all ${rows.length}` }
    ))) return;
    let okCount = 0;
    let failCount = 0;
    try {
      for (let i = 0; i < rows.length; i++) {
        setBulkKlammerProgress(`${i + 1}/${rows.length}`);
        if (await addToKlammerCore(rows[i])) okCount++; else failCount++;
      }
      const regs2 = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs2);
      showAlert(
        failCount === 0
          ? (isDe ? `${okCount} Klammer-Anmeldungen nachgetragen.` : `${okCount} umbrella registrations added.`)
          : (isDe
            ? `${okCount} Klammer-Anmeldungen nachgetragen, ${failCount} fehlgeschlagen (typisch: SharePoint-Drosselung) — bitte in ein paar Minuten erneut ausführen, bereits nachgetragene werden übersprungen.`
            : `${okCount} umbrella registrations added, ${failCount} failed (typically SharePoint throttling) — please run again in a few minutes; already-added ones are skipped.`),
        { variant: failCount === 0 ? 'success' : 'error' });
    } finally {
      setBulkKlammerProgress('');
    }
  };

  // v24.40: Eine Person (Klammer + alle aktiven Sub-Event-Anmeldungen) einer
  // gewählten Assistenz zuordnen, damit diese die Anmeldung in ihrer
  // „Assistenz"-Kachel vollständig verwalten kann. Setzt pro betroffener Zeile
  // RegisteredBy + Zeilen-Autor auf die Assistenz (eventService-Helfer).
  const submitAssignAssistant = async (): Promise<void> => {
    if (!assignAssistRow || !selectedEvent || !eventServiceRef) return;
    const m = (assignAssistValue || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
    if (!m) {
      showAlert(isDe ? 'Bitte eine Assistenz aus der Suche auswählen.' : 'Please select an assistant from the search.', { variant: 'error' });
      return;
    }
    const assistName = m[1].trim();
    const assistEmail = m[2].trim();
    const row = assignAssistRow;
    const personName = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email;
    if ((assistEmail || '').toLowerCase() === (row.emailKey || '')) {
      showAlert(isDe ? 'Die Assistenz darf nicht dieselbe Person wie die angemeldete Person sein.' : 'The assistant must not be the same person as the registered person.', { variant: 'error' });
      return;
    }
    setAssignAssistBusy(true);
    try {
      let done = 0;
      let failed = 0;
      // 1) Klammer-/Hauptevent-Zeile.
      const parentReg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey);
      if (parentReg && selectedEvent.subsiteUrl) {
        const ok = await eventServiceRef.assignRegistrationToAssistant(selectedEvent.subsiteUrl, parentReg.Id, assistEmail, assistName);
        if (ok) done += 1; else failed += 1;
      }
      // 2) Alle aktiven Sub-Event-Zeilen der Person.
      const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
      for (const ch of consolidatedChildren) {
        const r = row.perChild[ch.id];
        if (!r || ACTIVE.indexOf(r.Status) < 0 || !ch.subsiteUrl) continue;
        const ok = await eventServiceRef.assignRegistrationToAssistant(ch.subsiteUrl, r.Id, assistEmail, assistName);
        if (ok) done += 1; else failed += 1;
      }
      try {
        await eventServiceRef.writeChangeLog({
          action: 'ParticipantUpdated',
          targetType: 'Participant',
          targetId: (row.email || '') + '#assistant',
          targetName: personName,
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { scope: 'assignedAssistant', assistantEmail: assistEmail, actorEmail: currentUser.email, rowsUpdated: done },
        });
      } catch { /* */ }
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
      setSubRegReloadTick(t => t + 1);
      setAssignAssistRow(null);
      setAssignAssistValue('');
      if (failed === 0) {
        showAlert(isDe ? `„${personName}" wurde der Assistenz „${assistName}" zugeordnet (${done} Anmeldung(en)). Sie kann die Anmeldung jetzt in ihrer „Assistenz"-Kachel verwalten.` : `„${personName}" was assigned to assistant „${assistName}" (${done} registration(s)). They can now manage it in their „Assistant" tile.`, { variant: 'success' });
      } else {
        showAlert(isDe ? `Teilweise zugeordnet: ${done} erfolgreich, ${failed} fehlgeschlagen (evtl. fehlende Rechte). Bei externen/nicht auffindbaren Konten ist die Zuordnung nicht möglich.` : `Partially assigned: ${done} ok, ${failed} failed (possibly missing permissions).`, { variant: 'error' });
      }
    } catch (err) {
      console.warn('[DEX] submitAssignAssistant error:', err);
      showAlert(isDe ? 'Unerwarteter Fehler bei der Zuordnung.' : 'Unexpected error during assignment.', { variant: 'error' });
    } finally {
      setAssignAssistBusy(false);
    }
  };

  const openMainFieldsEdit = (emailKey: string, displayName: string): void => {
    if (!selectedEvent) return;
    setMainFieldsEditError('');
    // 1) Bevorzugt die Hauptevent-Anmeldung (Klammer-Subsite).
    const parentReg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === emailKey) || null;
    let targetReg: SPRegistration | null = parentReg;
    let targetSubsite = selectedEvent.subsiteUrl || '';
    let isParent = true;
    // 2) Fallback (Klammer-Modus ohne Hauptevent-Anmeldung): die früheste
    //    aktive Sub-Event-Zeile der Person als Speicherort nehmen.
    if (!parentReg) {
      const row = consolidatedRows.find(r => r.emailKey === emailKey);
      if (row) {
        const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
        let best: { reg: SPRegistration; sub: string; ts: number } | null = null;
        for (const ch of consolidatedChildren) {
          const r = row.perChild[ch.id];
          if (!r || ACTIVE.indexOf(r.Status) < 0) continue;
          const ts = r.RegistrationDate ? new Date(r.RegistrationDate).getTime() : Number.POSITIVE_INFINITY;
          if (!best || ts < best.ts) best = { reg: r, sub: ch.subsiteUrl || '', ts };
        }
        if (best && best.sub) { targetReg = best.reg; targetSubsite = best.sub; isParent = false; }
      }
    }
    setMainFieldsEditReg(targetReg);
    setMainFieldsEditSubsite(targetSubsite);
    setMainFieldsEditTargetIsParent(isParent);
    setMainFieldsEditName(displayName);
    const initial: Record<string, string> = {};
    if (targetReg) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyReg = targetReg as any;
      let cd: Record<string, unknown> = {};
      if (anyReg.CustomData) { try { cd = JSON.parse(anyReg.CustomData); } catch { /* */ } }
      const parentFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'document' && f.label && f.label.trim());
      for (const f of parentFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = (f as any).spInternalName || '';
        // Bei Sub-Event-Ziel NUR aus CustomData lesen (die Klammer-SP-Spalten
        // existieren auf der Sub-Event-Liste i.d.R. nicht).
        let v: unknown = (isParent && sp) ? anyReg[sp] : undefined;
        if (v === undefined || v === null || v === '') v = cd[f.id];
        initial[f.id] = (v === undefined || v === null) ? '' : String(v);
      }
    }
    setMainFieldsEditForm(initial);
  };
  const closeMainFieldsEdit = (): void => {
    setMainFieldsEditReg(null);
    setMainFieldsEditName('');
    setMainFieldsEditForm({});
    setMainFieldsEditError('');
    setMainFieldsEditSubsite('');
    setMainFieldsEditTargetIsParent(true);
  };
  // v19.30 — Feature A: Speichern der Hauptevent-Custom-Felder. Persistiert
  // über dasselbe `adminUpdateRegistration` wie das reguläre Teilnehmer-Edit
  // (schreibt die SP-Spalten der Hauptevent-Teilnehmerliste) und legt eine
  // Audit-Zeile 'ParticipantUpdated' mit dem Vorher/Nachher-Diff an. Es werden
  // nur geänderte Felder ins Patch aufgenommen — sonst kippt ein unverändertes
  // Choice-Feld den ganzen Save (HTTP 400 'Invalid choice').
  const saveMainFieldsEdit = async (): Promise<void> => {
    const targetSubsite = mainFieldsEditSubsite || selectedEvent?.subsiteUrl || '';
    if (!mainFieldsEditReg || !eventServiceRef || !targetSubsite || !selectedEvent) return;
    const isParentTarget = mainFieldsEditTargetIsParent;
    setMainFieldsEditSaving(true);
    setMainFieldsEditError('');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyReg = mainFieldsEditReg as any;
      let cd: Record<string, unknown> = {};
      if (anyReg.CustomData) { try { cd = JSON.parse(anyReg.CustomData); } catch { /* */ } }
      const parentFields = (selectedEvent?.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'document' && f.label && f.label.trim());
      const patch: Record<string, unknown> = {};
      const oldValues: Record<string, unknown> = {};
      const fieldLabelMap: Record<string, string> = {};
      const nextCd: Record<string, unknown> = { ...cd };
      let cdChanged = false;
      for (const f of parentFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = (f as any).spInternalName || '';
        const newVal = mainFieldsEditForm[f.id] || '';
        let oldVal = '';
        // Bei Sub-Event-Ziel den alten Wert NUR aus CustomData lesen (keine
        // Klammer-SP-Spalten auf der Sub-Event-Liste).
        let oldFromSp: unknown = (isParentTarget && sp) ? anyReg[sp] : undefined;
        if (oldFromSp === undefined || oldFromSp === null || oldFromSp === '') oldFromSp = cd[f.id];
        if (oldFromSp !== undefined && oldFromSp !== null) oldVal = String(oldFromSp);
        if (newVal === oldVal) continue; // unverändert → überspringen
        const keyForAudit = (isParentTarget && sp) ? sp : f.id;
        fieldLabelMap[keyForAudit] = f.label;
        oldValues[keyForAudit] = oldVal;
        // SP-Spalte NUR beim Hauptevent-Ziel patchen (Sub-Event-Liste hat die
        // Klammer-Spalten nicht → würde HTTP 400 werfen). Sonst CustomData-only.
        if (isParentTarget && sp) patch[sp] = newVal;
        nextCd[f.id] = newVal;
        cdChanged = true;
      }
      if (!cdChanged && Object.keys(patch).length === 0) {
        closeMainFieldsEdit();
        return;
      }
      // CustomData immer mitschreiben, damit der konsolidierte View (der bei
      // fehlender SP-Spalte auf CustomData zurückfällt) konsistent bleibt.
      if (cdChanged) patch['CustomData'] = JSON.stringify(nextCd);
      const actor = {
        name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email,
        email: currentUser.email,
      };
      const ok = await eventServiceRef.adminUpdateRegistration(
        targetSubsite, mainFieldsEditReg.Id, patch, actor, oldValues, fieldLabelMap
      );
      if (!ok) {
        setMainFieldsEditError(isDe
          ? 'Speichern fehlgeschlagen — vermutlich fehlt eine SP-Spalte in der Hauptevent-Teilnehmerliste. Klicke einmal „Spalten fixen" für das Hauptevent, dann erneut versuchen.'
          : 'Save failed — likely a missing SP column on the main-event participant list. Click „Fix columns" for the main event once, then retry.');
        return;
      }
      // Audit-Log mit Diff der geänderten Felder (analog saveEdit).
      try {
        const changes: Record<string, { old: unknown; new: unknown }> = {};
        for (const k of Object.keys(oldValues)) {
          changes[k] = { old: oldValues[k], new: (k in patch ? patch[k] : mainFieldsEditForm[k]) };
        }
        await eventServiceRef.writeChangeLog({
          action: 'ParticipantUpdated',
          targetType: 'Participant',
          targetId: (mainFieldsEditReg.ParticipantEmail || '') + '#' + mainFieldsEditReg.Id,
          targetName: `${mainFieldsEditReg.Vorname || ''} ${mainFieldsEditReg.Nachname || ''}`.trim() || mainFieldsEditName,
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { changes, scope: 'mainEventFields' },
        });
      } catch { /* */ }
      // Neu laden, damit die Klammer-Feld-Spalten die neuen Werte zeigen.
      // Hauptevent-Ziel → Hauptevent-Teilnehmerliste; Sub-Event-Ziel → die
      // konsolidierten Sub-Event-Registrierungen neu ziehen.
      if (isParentTarget) {
        const regs = await getAllRegistrations(selectedEvent.id);
        setRegistrations(regs);
      } else {
        setSubRegReloadTick(t => t + 1);
      }
      closeMainFieldsEdit();
    } catch (err) {
      console.warn('[DEX] saveMainFieldsEdit error:', err);
      setMainFieldsEditError(isDe
        ? 'Unerwarteter Fehler beim Speichern.'
        : 'Unexpected error while saving.');
    } finally {
      setMainFieldsEditSaving(false);
    }
  };

  // v19.30 — Feature B: Sub-Event-Registrierungen neu laden (nach einer
  // Abmeldung im konsolidierten View). Spiegelt den Lade-Effekt von oben.
  const reloadSubEventRegs = async (): Promise<void> => {
    if (!selectedEvent || !selectedEvent.subEventsOnlyMode) return;
    const children = childEventsOf(selectedEvent.id);
    const map: Record<string, SPRegistration[]> = {};
    for (const ch of children) {
      try { map[ch.id] = await getAllRegistrations(ch.id); }
      catch { map[ch.id] = []; }
    }
    setSubEventRegsByEventId(map);
  };
  // v19.30 — Feature B: Abmelde-Modal für eine konsolidierte Zeile öffnen.
  // Sammelt alle Sub-Events, in denen die Person aktiv angemeldet ist.
  const openDeregModal = (row: ConsolidatedRow): void => {
    if (!selectedEvent) return;
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const items: Array<{ child: DeloitteEvent; reg: SPRegistration; isParent?: boolean }> = [];
    // v29.29: Die Klammer-Zeile ZUERST — sie gehört zur Person genauso wie die
    // Sub-Events. Ohne sie blieb nach dem Abmelden aller Sub-Events eine
    // Schatten-Anmeldung auf dem Hauptevent stehen, die die Teilnehmerzahl
    // weiter mitzählte und in „Meine Events" der Person erschien.
    const parentReg = registrations.find(r =>
      (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey && ACTIVE.indexOf(r.Status) >= 0);
    if (parentReg) items.push({ child: selectedEvent, reg: parentReg, isParent: true });
    for (const ch of childEventsOf(selectedEvent.id)) {
      const r = row.perChild[ch.id];
      if (r && ACTIVE.indexOf(r.Status) >= 0) items.push({ child: ch, reg: r });
    }
    setDeregModal({
      emailKey: row.emailKey,
      name: `${row.vorname} ${row.nachname}`.trim() || row.email,
      email: row.email,
      items,
    });
    // Default: alles vorausgewählt — der häufigste Fall ist „ganz abmelden".
    // Der Organizer kann einzelne wieder abwählen.
    setDeregSelected(new Set(items.map(i => i.child.id)));
    // v29.29: Bei einem als inaktiv gemeldeten Konto (Person hat das
    // Unternehmen verlassen) ist die stille Abmeldung der Normalfall.
    setDeregSilent(inactiveAccounts.indexOf(row.emailKey) >= 0);
  };
  const closeDeregModal = (): void => {
    setDeregModal(null);
    setDeregSelected(new Set());
    setDeregSilent(false);
  };
  // v19.30 — Feature B: Abmeldung pro gewähltem Sub-Event durchführen. Spiegelt
  // exakt die Nebenwirkungen des Einzel-Event-Abmeldens (Abmelde-Mail +
  // Outlook 'Ausladen' + DEX_Participants-Cleanup + Nachrücken + ID-Reorder)
  // und schreibt zusätzlich pro Abmeldung eine 'RegistrationCancelled'-
  // Audit-Zeile (die der Einzel-Pfad nicht setzt).
  const runDeregModal = async (): Promise<void> => {
    if (!deregModal || !eventServiceRef) return;
    setDeregBusy(true);
    const actorName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
    const actorEmail = currentUser.email;
    const chosen = deregModal.items.filter(i => deregSelected.has(i.child.id));
    for (const { child, reg, isParent } of chosen) {
      const sub = child.subsiteUrl;
      if (!sub) continue;
      const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
      const cancelledStarterType = reg.StarterType || '';
      try {
        await eventServiceRef.cancelRegistration(sub, reg.Id, actorName, actorEmail);
        // Audit-Zeile (Feature D: Abmeldungen sollen im Event-Log auftauchen).
        try {
          await eventServiceRef.writeChangeLog({
            action: 'RegistrationCancelled',
            targetType: 'Participant',
            targetId: (reg.ParticipantEmail || '') + '#' + reg.Id,
            targetName: name,
            eventId: child.id,
            eventTitle: child.title,
            details: { asActor: 'organizer', via: 'consolidatedDeregister', ...(deregSilent ? { silent: true } : {}), ...(isParent ? { level: 'parent' } : {}) },
          });
        } catch { /* */ }
        // Abmelde-Mail + Outlook 'Ausladen' (event-weite Schalter respektieren).
        // v22.22: Vergangenes Sub-Event → stille Abmeldung (keine Mail, kein
        // Outlook, kein Nachrücken, kein ID-Reorder).
        const childWasOver = isEventOver(child);
        // v29.29: `deregSilent` unterdrückt NUR die Benachrichtigung der
        // ausscheidenden Person (erloschenes Postfach) — Nachrücken und
        // ID-Reorder laufen weiter, der frei gewordene Platz soll ja an die
        // Warteliste gehen und die nachrückende Person ihre Mail bekommen.
        const notifyLeaver = !childWasOver && !deregSilent;
        // v29.44: Auf der KLAMMER keine zweite Abmelde-Bestätigung, wenn im
        // selben Lauf auch Sub-Events abgemeldet werden — dafür ging deren
        // eigene Mail schon raus. Die Klammer-Zeile ist seit v29.29 Teil des
        // Dialogs; seither bekam der Teilnehmer zusätzlich eine Mail mit dem
        // Klammer-Titel, obwohl er sich von einem Termin abgemeldet hat. Im
        // Modus „nur Sub-Events" ist die Klammer ohnehin nur eine
        // Schattenzeile — dort nie eine eigene Mail.
        const skipParentMail = !!isParent && (!!child.subEventsOnlyMode || chosen.some(i => !i.isParent));
        if (reg.ParticipantEmail && notifyLeaver) {
          if (!child.disableEmails && !child.disableCancellationEmail && !skipParentMail) {
            try {
              // v29.44: die für das Event gepflegte Abmelde-Vorlage nehmen —
              // vorher IMMER der Code-Standardtext. Deshalb sah die vom
              // Organizer ausgelöste Abmeldung anders aus als die, die der
              // Teilnehmer beim Selbst-Abmelden bekommt (dort läuft es seit
              // jeher über Vorlage + Event-Override). Gleicher Weg wie beim
              // Nachrücken ein paar Zeilen weiter unten.
              const emailData = await buildCancellationMail(child, reg, name);
              await eventServiceRef.queueEmail(
                emailData.subject, reg.ParticipantEmail, name, emailData.body,
                'Abmeldung', child.title, child.id
              );
            } catch (err) { console.warn('[DEX]', err); }
          }
          if (!child.disableOutlook) {
            try {
              await eventServiceRef.queueOutlookEvent(
                reg.ParticipantEmail, child.id, child.title, 'Ausladen'
              );
            } catch (err) { console.warn('[DEX]', err); }
          }
        }
        // DEX_Participants aufräumen.
        if (reg.ParticipantEmail && child.eventNumber) {
          eventServiceRef.removeParticipantEvent(reg.ParticipantEmail, child.eventNumber)
            .catch(err => console.warn('[DEX]', err));
        }
        // Client-seitiges Nachrücken (typ-bewusst bei Split-Capacity, außer
        // splitSharedWaitlist) — identisch zum Einzel-Event-Abmelden.
        const isSplitEvent = typeof child.durchstarterCapacity === 'number'
          && typeof child.funstarterCapacity === 'number'
          && ((child.durchstarterCapacity || 0) > 0 || (child.funstarterCapacity || 0) > 0);
        const useTypeFilter = isSplitEvent && !child.splitSharedWaitlist;
        if (!childWasOver) {
        // v27.11: Kein automatisches Nachrücken, wenn die Warteliste des
        // Sub-Events abgeschaltet ist (Kill-Switch, s. Einzel-Event-Abmelden).
        // Der ID-Reorder unten läuft weiterhin.
        // v29.29: Auf der KLAMMER nie nachrücken — sie ist keine
        // Anmeldeeinheit (maxParticipants ist dort 0, die Zeile ist eine
        // Schattenzeile). Ein Nachrücken dort würde eine fremde Person auf
        // eine Ebene heben, die niemand bucht.
        if (!isParent && child.waitlistEnabled !== false) {
        try {
          const promoted = await eventServiceRef.promoteFirstWaitlistItem(
            sub,
            cancelledStarterType || undefined,
            child.maxParticipants,
            (useTypeFilter && cancelledStarterType) ? cancelledStarterType : undefined,
            { itemId: reg.Id, participantEmail: reg.ParticipantEmail || '' },
          );
          if (promoted && promoted.success && promoted.email) {
            if (!child.disableEmails) {
              try {
                const lang = child.emailLanguage || 'EN';
                const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
                const promoteVars = {
                  Name: promotedFirstName,
                  EventTitle: child.title,
                  Organizer: formatOrganizerList(child.organizers, lang),
                  AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
                  WaitlistPosition: '',
                };
                let emailData: { subject: string; body: string };
                const spTplRaw = await eventServiceRef.getEmailTemplate('Nachruecken', lang).catch(() => null);
                const spTpl = applyEventTemplateOverride(spTplRaw, child.emailTemplateOverrides, 'Nachruecken');
                if (spTpl) emailData = buildEmailFromTemplate(spTpl, promoteVars);
                else emailData = promotionEmail(promotedFirstName, child.title);
                await eventServiceRef.queueEmail(
                  withParentTitleSubject(emailData.subject, selectedEvent && selectedEvent.subEventCalendar ? selectedEvent : undefined),
                  promoted.email, promoted.name || '', emailData.body,
                  'Nachruecken', child.title, child.id
                );
              } catch (err) { console.warn('[DEX] promote-email failed:', err); }
            }
            if (!child.disableOutlook) {
              try {
                await eventServiceRef.queueOutlookEvent(promoted.email, child.id, child.title, 'Einladen');
              } catch (err) { console.warn('[DEX] promote-outlook failed:', err); }
            }
          }
        } catch (err) { console.warn('[DEX] promoteFirstWaitlistItem failed:', err); }
        }
        // ID-Reorder in die Queue (Flow macht nur noch Reorder).
        try {
          await eventServiceRef.queueIDReorder(
            child.id, child.eventNumber || 0, sub, child.title, name, reg.ParticipantEmail || undefined
          );
        } catch (err) { console.warn('[DEX] queueIDReorder threw:', err); }
        }
      } catch (err) {
        console.warn('[DEX] consolidated deregister failed for child', child.id, err);
      }
    }
    try { await reloadSubEventRegs(); } catch { /* */ }
    // v29.31: Gecachte „Konto inaktiv"-Ergebnisse dieser Event-Familie
    // verwerfen (s. performStandardCancel).
    invalidateInactiveAccountCache(chosen.map(c => c.child.id).concat(selectedEvent ? [selectedEvent.id] : []));
    // v29.29: Auch die Klammer-Teilnehmerliste neu laden — seit die
    // Hauptevent-Zeile mit abgemeldet werden kann, wäre die Kopfzeile
    // („Teilnehmer (N)") sonst bis zum nächsten Öffnen veraltet.
    if (selectedEvent) {
      try { setRegistrations(await getAllRegistrations(selectedEvent.id)); } catch { /* */ }
    }
    setDeregBusy(false);
    closeDeregModal();
  };

  // v19.30 — Feature D: Audit-Log vorgefiltert auf das aktuell selektierte
  // Event öffnen (setzt den Event-/Ziel-Filter auf den Event-Titel).
  const openChangeLogForEvent = (): void => {
    setChangeLogFilterEvent(selectedEvent?.title || '');
    setChangeLogFilterAction('');
    setChangeLogFilterActor('');
    void openChangeLog();
  };

  // v11.36: Fairer Wartelisten-Rang einer Person in ihrer Gruppe — gleiche
  // Logik wie die Review-Box. Genutzt für die "neue Warteliste-Position" im
  // Mailtext (Vorschlag + Sammel-Versand).
  const getFairWaitlistRank = (reg: SPRegistration): number => {
    if (!selectedEvent) return 0;
    const ACT = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const isSplit = isSplitCapacity;
    const keyOf = (r: SPRegistration): string => isSplit ? (r.StarterType || r.PreferredStarterType || '?') : 'all';
    const capOf = (k: string): number => !isSplit
      ? (selectedEvent.maxParticipants || 0)
      : (k === 'Durchstarter' ? (selectedEvent.durchstarterCapacity || 0) : k === 'Funstarter' ? (selectedEvent.funstarterCapacity || 0) : 0);
    const k = keyOf(reg);
    const activeSorted = registrations
      .filter(r => ACT.indexOf(r.Status) >= 0 && keyOf(r) === k)
      .slice().sort((a, b) => a.Id - b.Id);
    const cap = capOf(k);
    const overCap = cap > 0 ? activeSorted.slice(cap) : [];
    const existingWl = registrations.filter(r => r.Status === 'Warteliste' && keyOf(r) === k);
    const fairWl = [...overCap, ...existingWl].sort((a, b) =>
      new Date(a.RegistrationDate).getTime() - new Date(b.RegistrationDate).getTime());
    const idx = fairWl.findIndex(x => x.Id === reg.Id);
    return idx >= 0 ? idx + 1 : 0;
  };

  // v11.36: Beim Öffnen des „Bestätigen"-Dialogs die Mail-Sprache aus dem
  // Event vorbelegen (Default DE wenn nicht explizit EN) — umschaltbar.
  React.useEffect(() => {
    if (overbookModal?.mode === 'confirm') {
      setObMailLang((selectedEvent?.emailLanguage || '').toUpperCase() === 'EN' ? 'EN' : 'DE');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [overbookModal]);

  // v11.36: Mailtext vorbefüllen — reagiert auf Dialog-Öffnen UND Sprachwahl.
  // Enthält die neue Wartelisten-Position der ersten Zielperson.
  // v13.0: buildOverbookApologyEmail ist jetzt async (lädt Template aus
  // DEX_EmailTemplates). Effect wartet auf das Promise und setzt State
  // wenn der Modal noch offen ist.
  React.useEffect(() => {
    if (overbookModal?.mode === 'confirm' && eventServiceRef && selectedEvent) {
      const t = overbookModal.targets[0];
      const nm = t ? ((t.Vorname && t.Nachname) ? `${t.Vorname} ${t.Nachname}` : t.ParticipantName) : '';
      const pos = t ? getFairWaitlistRank(t) : 0;
      let cancelled = false;
      eventServiceRef.buildOverbookApologyEmail(nm, selectedEvent.title, obMailLang, pos).then(m => {
        if (cancelled) return;
        setObMailSubject(m.subject);
        setObMailBody(m.body);
      }).catch(() => { /* */ });
      return () => { cancelled = true; };
    }
    return undefined;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [overbookModal, obMailLang]);

  // v11.36: Überbuchungs-Entscheidung ausführen (einzeln oder Sammel) und
  // danach IDs neu vergeben + Counter/Seat-Sync + Liste neu laden.
  const runOverbookResolution = async (): Promise<void> => {
    if (!overbookModal || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setObBusy(true);
    const sub = selectedEvent.subsiteUrl;
    const isBulk = overbookModal.targets.length > 1;
    try {
      for (const reg of overbookModal.targets) {
        const grp = reg.StarterType || reg.PreferredStarterType || '';
        const nm = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
        if (overbookModal.mode === 'confirm') {
          await eventServiceRef.resolveOverbookToWaitlist(sub, reg.Id, grp);
          if (obWithMail && reg.ParticipantEmail && !selectedEvent.disableEmails) {
            // Einzeln: ggf. vom Admin editierter Text. Sammel: pro Person
            // frisch personalisiert aus dem Standard-Template.
            const mail = isBulk
              ? await eventServiceRef.buildOverbookApologyEmail(nm, selectedEvent.title, obMailLang, getFairWaitlistRank(reg))
              : { subject: obMailSubject, body: obMailBody };
            try {
              await eventServiceRef.queueEmail(
                mail.subject, reg.ParticipantEmail, nm, mail.body,
                'Info', selectedEvent.title, selectedEvent.id
              );
            } catch { /* Mail-Fehler darf Korrektur nicht blockieren */ }
          }
          if (obRemoveCalendar && reg.ParticipantEmail && !selectedEvent.disableOutlook) {
            try {
              await eventServiceRef.queueOutlookEvent(
                reg.ParticipantEmail, selectedEvent.id, selectedEvent.title, 'Ausladen'
              );
            } catch { /* Kalender-Abmeldung best-effort */ }
          }
        } else {
          // Platz behalten
          if (obKeepVariant === 'active') {
            await eventServiceRef.resolveOverbookKeepActive(sub, reg.Id);
          } else {
            await eventServiceRef.resolveOverbookKeepAsFirstWaitlist(sub, reg.Id, grp);
          }
        }
      }
      // IDs neu vergeben (Aktive 1..N, Warteliste N+1..) + Counter + Seat-Sync.
      // Mit Fortschritts-Overlay, damit man bei großen Listen sieht wie weit.
      setReorderProgressLabel('IDs werden neu vergeben…');
      setReorderProgress(0);
      try { await eventServiceRef.reorderParticipantIDs(sub, pct => setReorderProgress(pct)); } catch { /* */ }
      try { await eventServiceRef.syncSeatsToActiveCount(sub, { isSplit: isSplitCapacity }); } catch { /* */ }
      setReorderProgress(null);
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch { /* einzelne Fehler werden geschluckt; Liste wird trotzdem neu geladen */ }
    setObBusy(false);
    setOverbookModal(null);
  };

  // v11.36: TeilnehmerIDs neu vergeben — gemeinsam von der Toolbox-Kachel
  // UND dem Hinweis-Modal genutzt (mit %-Fortschritts-Overlay).
  const runIdReorder = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsReorderingIDs(true);
    setReorderResult(null);
    setReorderProgressLabel(isDe ? 'IDs werden neu vergeben…' : 'Reassigning IDs…');
    setReorderProgress(0);
    try {
      const result = await eventServiceRef.reorderParticipantIDs(
        selectedEvent.subsiteUrl,
        pct => setReorderProgress(pct)
      );
      setReorderResult(isDe
        ? `${result.success} aktualisiert, ${result.errors} Fehler`
        : `${result.success} updated, ${result.errors} errors`);
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch {
      setReorderResult(isDe ? 'Fehler beim Neuvergeben der IDs' : 'Error reassigning IDs');
    }
    setReorderProgress(null);
    setIsReorderingIDs(false);
  };

  /**
   * v29.16: Nachrück-Mail + Outlook-Einladung für EINE nachgerückte Person.
   * Aus `runManualPromote` herausgezogen, weil das Füllen mehrerer freier
   * Plätze dieselbe Benachrichtigung je Person braucht — zwei Kopien dieses
   * Mail-Aufbaus wären beim nächsten Template-Wechsel auseinandergelaufen.
   */
  const notifyPromoted = async (promoted: { email: string; name?: string }): Promise<void> => {
    if (!eventServiceRef || !selectedEvent) return;
    if (!selectedEvent.disableEmails) {
      try {
        const lang = selectedEvent.emailLanguage || 'EN';
        const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
        const promoteVars = {
          Name: promotedFirstName,
          EventTitle: selectedEvent.title,
          Organizer: formatOrganizerList(selectedEvent.organizers, lang),
          AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
          WaitlistPosition: '',
        };
        let emailData: { subject: string; body: string };
        const spTplRaw = await eventServiceRef.getEmailTemplate('Nachruecken', lang).catch(() => null);
        const spTpl = applyEventTemplateOverride(spTplRaw, selectedEvent.emailTemplateOverrides, 'Nachruecken');
        if (spTpl) {
          emailData = buildEmailFromTemplate(spTpl, promoteVars);
        } else {
          emailData = promotionEmail(promotedFirstName, selectedEvent.title);
        }
        await eventServiceRef.queueEmail(
          withParentTitleSubject(emailData.subject, selectedEvent.parentEventId ? allEvents.find(e => e.id === selectedEvent.parentEventId) : undefined),
          promoted.email, promoted.name || '', emailData.body,
          'Nachruecken', selectedEvent.title, selectedEvent.id
        );
      } catch (err) { console.warn('[DEX] promote-email failed:', err); }
    }
    if (!selectedEvent.disableOutlook) {
      try {
        await eventServiceRef.queueOutlookEvent(
          promoted.email, selectedEvent.id, selectedEvent.title, 'Einladen'
        );
      } catch (err) { console.warn('[DEX] promote-outlook failed:', err); }
    }
  };

  /**
   * v18.70 / v29.16: Freie Plätze mit der Warteliste füllen.
   *
   * v18.70 rückte GENAU EINE Person nach und übergab weder Gruppe noch die
   * Gruppen-Kapazität. Auf einem Event mit zwei Gruppen ging das doppelt
   * daneben: Als Obergrenze stand `maxParticipants` drin — das ist bei
   * geteilten Kapazitäten 0, also fand gar keine Prüfung statt — und ohne
   * Typfilter nahm die Abfrage den ersten Wartelistler nach TeilnehmerID,
   * egal aus welcher Gruppe. Eine noch volle Gruppe konnte damit überbucht
   * werden, während die Gruppe mit freien Plätzen leer ausging.
   *
   * Zweiter Punkt, der den Fall überhaupt erst sichtbar macht: Nachgerückt
   * wird sonst NUR beim Abmelden. Erhöht der Organizer eine Gruppengröße,
   * passiert von allein nichts — es gibt kein Ereignis, an dem etwas hinge.
   * Genau dafür ist diese Aktion da, und sie füllt deshalb ALLE freien
   * Plätze auf einmal statt einen pro Klick.
   *
   * Gezählt wird je Gruppe (bzw. einmal gesamt ohne geteilte Kapazität) mit
   * derselben Zuordnung wie überall sonst: StarterType, ersatzweise
   * PreferredStarterType. Bei `splitSharedWaitlist` gibt es nur eine
   * Warteliste — dann entscheidet die Reihenfolge, nicht die Gruppe.
   */
  const runManualPromote = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsPromoting(true);
    setPromoteResult(null);
    try {
      // Frisch lesen: `registrations` kann Minuten alt sein, und wir leiten
      // daraus ab, wie oft nachgerückt wird. Auf einem veralteten Stand
      // würde die Schleife über die Kapazität hinauslaufen.
      const fresh = await getAllRegistrations(selectedEvent.id);
      const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
      const lblA = (selectedEvent.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
      const lblB = (selectedEvent.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
      // Geteilte Kapazität MIT getrennten Wartelisten → je Gruppe rechnen.
      // Gemeinsame Warteliste (splitSharedWaitlist) verhält sich wie ein
      // einzelner Topf: Wer zuerst wartet, rückt nach.
      const perGroup = isSplitCapacity && !selectedEvent.splitSharedWaitlist;
      const groups: Array<{ key?: string; label: string; cap: number }> = perGroup
        ? [
          { key: 'Durchstarter', label: lblA, cap: selectedEvent.durchstarterCapacity || 0 },
          { key: 'Funstarter', label: lblB, cap: selectedEvent.funstarterCapacity || 0 },
        ]
        : [{
          key: undefined,
          label: isDe ? 'Plätze' : 'Seats',
          cap: isSplitCapacity
            // Gemeinsame Warteliste: Die Obergrenze ist die Summe beider
            // Gruppen — `maxParticipants` ist bei geteilten Kapazitäten 0.
            ? (selectedEvent.durchstarterCapacity || 0) + (selectedEvent.funstarterCapacity || 0)
            : (selectedEvent.maxParticipants || 0),
        }];
      const groupOf = (r: SPRegistration): string => r.StarterType || r.PreferredStarterType || '';
      const plan = groups.map(g => {
        const inGroup = (r: SPRegistration): boolean => !g.key || groupOf(r) === g.key;
        const active = fresh.filter(r => ACTIVE.indexOf(r.Status) >= 0 && inGroup(r)).length;
        const waiting = fresh.filter(r => r.Status === 'Warteliste' && inGroup(r)).length;
        // Kapazität 0 heißt „unbegrenzt" — dann rückt die ganze Warteliste nach.
        const free = g.cap > 0 ? Math.max(0, g.cap - active) : waiting;
        return { ...g, active, waiting, count: Math.min(free, waiting), free };
      });
      const total = plan.reduce((n, g) => n + g.count, 0);

      if (total === 0) {
        // Den tatsächlichen Grund nennen, nicht den erstbesten: „voll" und
        // „niemand wartet" führen zu ganz verschiedenen nächsten Schritten.
        const anyWaiting = plan.some(g => g.waiting > 0);
        setPromoteResult(!anyWaiting
          ? (isDe ? 'Niemand auf der Warteliste.' : 'Nobody on the waitlist.')
          : perGroup
            ? (isDe
              ? `Kein freier Platz in den Gruppen, in denen jemand wartet (${plan.filter(g => g.waiting > 0).map(g => `${g.label}: ${g.active}/${g.cap}`).join(', ')}).`
              : `No free seat in the groups where people are waiting (${plan.filter(g => g.waiting > 0).map(g => `${g.label}: ${g.active}/${g.cap}`).join(', ')}).`)
            : (isDe ? 'Kein freier Platz — Event ist voll.' : 'No free seat — event is full.'));
        setIsPromoting(false);
        return;
      }

      const lines = plan
        .filter(g => g.waiting > 0 || g.free > 0)
        .map(g => perGroup
          ? `• ${g.label}: ${g.active}/${g.cap || '∞'} belegt · ${g.waiting} auf der Warteliste → ${g.count} rücken nach`
          : `• ${g.active}/${g.cap || '∞'} belegt · ${g.waiting} auf der Warteliste → ${g.count} rücken nach`)
        .join('\n');
      const ok = await confirmDialog(
        isDe
          ? `${total} ${total === 1 ? 'Person' : 'Personen'} von der Warteliste nachrücken lassen?\n\n${lines}\n\nJede nachgerückte Person bekommt den Status „Angemeldet", eine Nachrück-Mail und eine Outlook-Einladung. Danach werden die TeilnehmerIDs neu vergeben.`
          : `Move ${total} ${total === 1 ? 'person' : 'people'} up from the waitlist?\n\n${lines}\n\nEach of them gets status “Registered”, a promotion email and an Outlook invite. Participant IDs are reassigned afterwards.`,
        { confirmLabel: isDe ? 'Nachrücken' : 'Promote' },
      );
      if (!ok) { setIsPromoting(false); return; }

      const promotedNames: string[] = [];
      let failed = 0;
      for (const g of plan) {
        for (let i = 0; i < g.count; i++) {
          // maxParticipants bewusst NICHT mitgeben: Die Überbuchungs-Sperre
          // im Service zählt über die GANZE Liste und kann eine Gruppe nicht
          // getrennt prüfen. Die Anzahl steht oben schon fest, frisch
          // gerechnet — hier wird nur genau so oft nachgerückt.
          const promoted = await eventServiceRef.promoteFirstWaitlistItem(
            selectedEvent.subsiteUrl,
            undefined,
            undefined,
            g.key,
          );
          if (promoted && promoted.success && promoted.email) {
            promotedNames.push(promoted.name || promoted.email);
            setAdminToast({ kind: 'promoted', name: promoted.name || promoted.email, email: promoted.email });
            await notifyPromoted({ email: promoted.email, name: promoted.name });
          } else {
            // Warteliste unerwartet leer (jemand hat sich zwischendurch
            // abgemeldet) — kein Fehler, nur nichts mehr zu tun.
            failed++;
            break;
          }
        }
      }

      // IDs neu vergeben + Counter/Seat-Sync + Liste neu laden — einmal am
      // Ende, nicht je Person: Der Reorder schreibt jede Zeile an.
      await runIdReorder();
      try {
        await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit: isSplitCapacity });
      } catch { /* */ }

      const n = promotedNames.length;
      setPromoteResult(n === 0
        ? (isDe ? 'Es konnte niemand nachrücken.' : 'Nobody could be promoted.')
        : n === 1
          ? (isDe ? `${promotedNames[0]} ist nachgerückt.` : `${promotedNames[0]} moved up.`)
          : (isDe
            ? `${n} Personen sind nachgerückt: ${promotedNames.join(', ')}.`
            : `${n} people moved up: ${promotedNames.join(', ')}.`));
      if (failed > 0 && n > 0) {
        console.warn('[DEX] runManualPromote: Warteliste war früher leer als erwartet.');
      }
    } catch (err) {
      console.warn('[DEX] runManualPromote failed:', err);
      setPromoteResult(isDe ? 'Fehler beim Nachrücken.' : 'Error promoting.');
    }
    setIsPromoting(false);
  };

  // v11.70 / v11.71: Hinweis-Box „IDs sind ggf. nicht korrekt" wird jetzt
  // an die tatsächliche TeilnehmerID-Sequenz gekoppelt — nicht mehr an
  // eine 10-Minuten-Zeit-Heuristik nach der letzten Abmeldung.
  //
  // Erwartet: alle nicht-abgemeldeten Einträge (Status in
  // Angemeldet/QR versendet/Eingecheckt/Warteliste) haben TeilnehmerIDs,
  // die nach Sortierung lückenlos 1..N durchlaufen. Sobald
  //   - eine ID fehlt (Lücke),
  //   - eine ID doppelt vorkommt,
  //   - ein nicht-abgemeldeter Eintrag keine (oder ≤0) ID hat,
  // ist der Zustand „IDs evtl. nicht korrekt". Typischer Trigger: gerade
  // erfolgte Abmeldung, der DEX_IDReorder-Flow ist noch nicht fertig.
  // Das gibt einen ehrlichen Status — die Box verschwindet automatisch,
  // sobald der Flow durch ist (statt nach willkürlichen 10 Minuten).
  const recentCancellation = (regs: SPRegistration[]): { recent: boolean; whenIso: string; detail: string } => {
    const active = regs.filter(r => r.Status !== 'Abgemeldet');
    if (active.length === 0) return { recent: false, whenIso: '', detail: '' };
    const ids: number[] = [];
    let noId = 0;
    for (const r of active) {
      const id = Number(r.TeilnehmerID);
      if (!isFinite(id) || id <= 0) { noId++; continue; }
      ids.push(id);
    }
    ids.sort((a, b) => a - b);
    // v22.12: konkrete Diagnose statt nur ja/nein — erste Lücke + Duplikate
    // zählen, damit die Box belegt, WAS in den geladenen Daten falsch ist.
    let dups = 0;
    let firstGapAt = 0;
    for (let i = 0; i < ids.length; i++) {
      if (i > 0 && ids[i] === ids[i - 1]) dups++;
      if (firstGapAt === 0 && ids[i] !== i + 1) firstGapAt = i + 1;
    }
    if (noId === 0 && dups === 0 && firstGapAt === 0) return { recent: false, whenIso: '', detail: '' };
    const parts: string[] = [];
    if (firstGapAt > 0) parts.push(`Nummern nicht durchgängig (erwartet Nr. ${firstGapAt})`);
    if (dups > 0) parts.push(`${dups} doppelte Nummer${dups === 1 ? '' : 'n'}`);
    if (noId > 0) parts.push(`${noId} Eintr${noId === 1 ? 'ag' : 'äge'} ohne Nummer`);
    return {
      recent: true,
      whenIso: latestCancelIso(regs),
      detail: `${active.length} aktive Einträge — ${parts.join(', ')}`,
    };
  };
  // Hilfsfunktion: jüngste CancellationDate aus der Liste (für den
  // optionalen Zeit-Hinweis in der Box).
  const latestCancelIso = (regs: SPRegistration[]): string => {
    let latest = 0;
    for (const r of regs) {
      if (r.Status !== 'Abgemeldet') continue;
      const t = new Date(r.CancellationDate || '').getTime();
      if (!isNaN(t) && t > latest) latest = t;
    }
    return latest > 0 ? new Date(latest).toISOString() : '';
  };
  const idFixCheckedForRef = React.useRef<string | null>(null);
  // v22.12: solange die geladenen Daten kaputte IDs zeigen, lädt die App die
  // Teilnehmerliste automatisch alle 30 Sekunden neu — sobald der
  // DEX_IDReorder-Flow durch ist, verschwindet die Warn-Box von selbst
  // (vorher musste man manuell „Aktualisieren" klicken und hielt den
  // durchgelaufenen Flow fälschlich für kaputt).
  const idRecheckBusyRef = React.useRef(false);
  const [idRecheckBusy, setIdRecheckBusy] = React.useState(false);
  const reloadRegistrationsForIdCheck = React.useCallback(async (): Promise<void> => {
    if (!selectedEvent || idRecheckBusyRef.current) return;
    idRecheckBusyRef.current = true;
    setIdRecheckBusy(true);
    try {
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch { /* best-effort */ }
    idRecheckBusyRef.current = false;
    setIdRecheckBusy(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id]);

  // v23.0: Eine Registrierung per Drag&Drop in ein Team/eine Break-Out-Session
  // (targetTid) oder zurück „ohne Team" ('') verschieben. War die Person Lead
  // ihres alten Teams und bleiben Mitglieder, rückt die früheste nach.
  const moveRegToTeam = async (reg: SPRegistration, targetTid: string, targetTeamName: string | undefined): Promise<void> => {
    if (!selectedEvent?.subsiteUrl || !eventServiceRef) return;
    const sub = selectedEvent.subsiteUrl;
    const curTid = reg.TeamId || '';
    if (curTid === (targetTid || '')) return;
    try {
      await eventServiceRef.assignRegistrationToTeam(sub, reg.Id, targetTid || '', targetTeamName || '', false);
      if (curTid && reg.TeamLead) {
        const rest = registrations.filter(x => x.Id !== reg.Id && (x.TeamId || '') === curTid && x.Status !== 'Abgemeldet');
        if (rest.length > 0) {
          rest.sort((a, b) => ((a.TeilnehmerID ?? 9_999_999) as number) - ((b.TeilnehmerID ?? 9_999_999) as number));
          const tn = rest.find(x => x.TeamName)?.TeamName || '';
          try { await eventServiceRef.assignRegistrationToTeam(sub, rest[0].Id, curTid, tn || '', true); } catch { /* */ }
        }
      }
      const nm = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantName || reg.ParticipantEmail;
      eventServiceRef.writeChangeLog({
        action: targetTid ? 'TeamMemberAssigned' : 'TeamMemberRemoved',
        targetType: 'Participant', targetId: reg.ParticipantEmail, targetName: nm,
        eventId: selectedEvent.id, eventTitle: selectedEvent.title,
        details: { fromTeam: curTid, toTeam: targetTid, via: 'dragdrop' },
      }).catch(() => { /* */ });
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch (err) { console.warn('[DEX] moveRegToTeam failed:', err); }
  };
  // Drop-Handler: gezogene Registrierung ermitteln + verschieben.
  const onTeamDrop = (targetTid: string, targetTeamName: string | undefined): void => {
    setDragOverTid(null);
    const id = dragRegId;
    setDragRegId(null);
    if (id === null) return;
    const reg = registrations.find(r => r.Id === id);
    if (reg) moveRegToTeam(reg, targetTid, targetTeamName).catch(() => { /* */ });
  };

  // v23.0: Aktive Teams aus den geladenen Registrierungen gruppieren
  // (für die Per-Team-Mail).
  const getActiveTeams = (): Array<{ tid: string; teamName: string; members: SPRegistration[] }> => {
    const map: Record<string, SPRegistration[]> = {};
    for (const r of registrations) {
      if (r.Status === 'Abgemeldet') continue;
      const tid = r.TeamId || '';
      if (!tid) continue;
      (map[tid] = map[tid] || []).push(r);
    }
    return Object.entries(map).map(([tid, members]) => ({ tid, members, teamName: members.find(m => m.TeamName)?.TeamName || '' }));
  };
  // Mail-Dialog mit vorausgefülltem Text öffnen.
  const openTeamMailDialog = (): void => {
    if (!selectedEvent) return;
    const termS = selectedEvent.teamTermSingular || 'Team';
    setTeamMailSubject(isDe ? `Deine ${termS}: ${selectedEvent.title}` : `Your ${termS}: ${selectedEvent.title}`);
    setTeamMailBody(isDe
      ? `<p>Hallo {{Vorname}},</p>\n<p>hier sind die Infos zu deiner <strong>${termS}</strong> beim Event <strong>{{EventTitle}}</strong>:</p>\n<p><strong>{{TeamName}}</strong></p>\n<p>{{TeamInfo}}</p>\n<p>Viele Grüße<br />Dein Event-Team</p>`
      : `<p>Hi {{Vorname}},</p>\n<p>here is the info for your <strong>${termS}</strong> at <strong>{{EventTitle}}</strong>:</p>\n<p><strong>{{TeamName}}</strong></p>\n<p>{{TeamInfo}}</p>\n<p>Best regards<br />Your event team</p>`);
    const init: Record<string, string> = {};
    for (const t of getActiveTeams()) init[t.tid] = teamMailInfoByTid[t.tid] || '';
    setTeamMailInfoByTid(init);
    setTeamMailOpen(true);
  };
  // Pro Team: jedes aktive Mitglied bekommt eine eigene Mail mit team-
  // spezifischer Info (z.B. Teams-Einwahllink). Im Deloitte-Layout gewrappt.
  const sendTeamMails = async (): Promise<void> => {
    if (!selectedEvent || !eventServiceRef) return;
    if (selectedEvent.disableEmails) {
      showAlert(isDe ? 'E-Mails sind für dieses Event deaktiviert (Schritt 6 „Kommunikation").' : 'Emails are disabled for this event (step 6 “Communication”).', { variant: 'error' });
      return;
    }
    setTeamMailSending(true);
    const escHtml = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    // URLs in der Team-Info klickbar machen + Zeilenumbrüche zu <br>.
    const linkify = (raw: string): string => escHtml(raw)
      .replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" style="color:#86bc25;font-weight:600;">$1</a>')
      .replace(/\n/g, '<br />');
    const termS = selectedEvent.teamTermSingular || 'Team';
    let sent = 0;
    for (const t of getActiveTeams()) {
      const infoHtml = linkify((teamMailInfoByTid[t.tid] || '').trim());
      const tName = t.teamName || termS;
      for (const m of t.members) {
        const first = (m.Vorname && m.Vorname.trim()) || (m.ParticipantName || '').split(/\s+/)[0] || '';
        const fullName = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantName || m.ParticipantEmail;
        const bodyFilled = teamMailBody
          .replace(/\{\{Vorname\}\}/g, escHtml(first))
          .replace(/\{\{Name\}\}/g, escHtml(fullName))
          .replace(/\{\{TeamName\}\}/g, escHtml(tName))
          .replace(/\{\{EventTitle\}\}/g, escHtml(selectedEvent.title))
          .replace(/\{\{TeamInfo\}\}/g, infoHtml || (isDe ? '<em>(keine zusätzlichen Infos)</em>' : '<em>(no additional info)</em>'));
        const subjectFilled = teamMailSubject
          .replace(/\{\{TeamName\}\}/g, tName)
          .replace(/\{\{EventTitle\}\}/g, selectedEvent.title);
        const wrapped = wrapTemplate('#86bc25', subjectFilled, tName, bodyFilled);
        try {
          const ok = await eventServiceRef.queueEmail(subjectFilled, m.ParticipantEmail, fullName, wrapped, 'TeamInfo', selectedEvent.title, selectedEvent.id);
          if (ok) sent += 1;
        } catch { /* best-effort pro Empfänger */ }
      }
    }
    setTeamMailSending(false);
    setTeamMailOpen(false);
    showAlert(isDe ? `${sent} Mail(s) in die Warteschlange gelegt — sie werden in Kürze versendet.` : `${sent} mail(s) queued — they will be sent shortly.`, { variant: 'success' });
  };

  // Max. 10 automatische Neu-Checks (≈5 Min) pro Event — wenn die Lücke dann
  // immer noch da ist, ist sie echt (Tail-Race, siehe Box-Text) und kein
  // weiteres Polling nötig.
  const idRecheckCountRef = React.useRef(0);
  React.useEffect(() => { idRecheckCountRef.current = 0; }, [selectedEvent?.id]);
  React.useEffect(() => {
    if (!selectedEvent) return undefined;
    // v22.67: kein ID-Durchgängigkeits-Polling im Klammer-Modus (Schatten-Zeilen
    // haben keine fortlaufenden Nummern — das war ein Fehlalarm).
    if (selectedEvent.subEventsOnlyMode) return undefined;
    if (!recentCancellation(registrations).recent) return undefined;
    if (idRecheckCountRef.current >= 10) return undefined;
    const timer = window.setInterval(() => {
      idRecheckCountRef.current++;
      if (idRecheckCountRef.current > 10) { window.clearInterval(timer); return; }
      reloadRegistrationsForIdCheck().catch(() => { /* */ });
    }, 30000);
    return () => window.clearInterval(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, registrations, reloadRegistrationsForIdCheck]);

  // v11.70: kein Modal mehr beim Event-Öffnen — der Hinweis steht ab
  // jetzt direkt als Box oben in der Teilnehmerliste, solange die
  // Bedingung erfüllt ist (siehe Render-Block unten). Der Ref bleibt
  // erhalten, um in Zukunft ein erneutes „Mount-Trigger"-Verhalten
  // einbauen zu können, ohne den Save-Pfad zu touchen.
  React.useEffect(() => {
    if (!selectedEvent || isLoadingRegs || registrations.length === 0) return;
    if (idFixCheckedForRef.current === selectedEvent.id) return;
    idFixCheckedForRef.current = selectedEvent.id;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id, registrations, isLoadingRegs]);

  // v6.17: Spaltenkonfiguration der Teilnehmertabelle (pro Event, lokal gespeichert).
  //  - columnOrder = geordnete Liste sichtbarer Spalten-IDs
  //  - hiddenColumns = ausgeblendete Spalten-IDs (können übers "+ Spalte"-Popover wieder zugeschaltet werden)
  // Die Spezialspalten 'id' / 'vorname' / 'nachname' / 'action' sind alwaysVisible und können nicht ausgeblendet werden.
  const [columnOrder, setColumnOrder] = React.useState<string[]>([]);
  const [hiddenColumns, setHiddenColumns] = React.useState<string[]>([]);
  const [showColumnPicker, setShowColumnPicker] = React.useState(false);
  // v26.44: „Matches anzeigen" — gruppiert die Teilnehmer-Tabelle in gegenseitige
  // Roommate-Paare (Match 1, Match 2, …) + Rest-Cluster. Nur relevant, wenn das
  // Event überhaupt eine Roommate-Spalte hat.
  const [showMatches, setShowMatches] = React.useState(false);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const spfxContext = (window as any).__dexSpfxContext;
  const eventServiceRef = React.useMemo(() => spfxContext ? new EventService(spfxContext) : null, []);
  // v28.65: Für die Rollenliste (Namens-Reparatur).
  const spServiceRef = React.useMemo(() => spfxContext ? new SharePointService(spfxContext) : null, []);

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

  // v6.17: Verfügbare Spalten der Teilnehmer-Tabelle aufbauen. MUSS vor dem
  // early return `if (!selectedEvent) return ...` stehen — sonst verletzen
  // die Hooks die Rules-of-Hooks (unterschiedliche Hook-Anzahl pro Render =
  // React Error #310).
  // v19.11: Hat dieses Event überhaupt Warteliste-/Nachrück-Aktivität? Nur dann
  // sind die Nachrück-Audit-Spalten („Nachgerückt am", „Hat ersetzt", „Wurde
  // ersetzt durch") sinnvoll. `waitlistEnabled` allein reicht NICHT, weil es per
  // Default `true` ist (e.WaitlistEnabled !== false) — Events ohne konfigurierte
  // Warteliste hätten sonst immer die leeren Audit-Spalten. „Aktiv" = jemand
  // steht auf der Warteliste ODER es gibt bereits Nachrück-Daten.
  const hasWaitlistActivity = React.useMemo(() => registrations.some(r =>
    r.Status === 'Warteliste'
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    || !!(r as any).PromotedDate || !!(r as any).ReplacedParticipantEmail || !!(r as any).ReplacedByParticipantEmail
  ), [registrations]);
  const availableColumns = React.useMemo(() => {
    const isSplit = !!selectedEvent
      && typeof selectedEvent.durchstarterCapacity === 'number'
      && typeof selectedEvent.funstarterCapacity === 'number'
      && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
    const userIds = (selectedEvent?.eventSpecificFields || [])
      .filter(f => f.type === 'user' || f.type === 'roommate')
      .map(f => f.id);
    // v14.8: Anrede-Spalte nur anbieten, wenn das Event die Anrede beim
    // Anmelden tatsächlich abfragt (askSalutation). Sonst landet eine leere
    // Spalte voller „-" im Admin-Center, die niemand braucht.
    const askSal = !!selectedEvent?.askSalutation;
    const cols: Array<{ id: string; label: string; alwaysVisible?: boolean }> = [
      { id: 'id', label: '#', alwaysVisible: true },
      ...(askSal ? [{ id: 'anrede', label: 'Anrede' }] : []),
      // v11.26: getrennte Vorname / Nachname Spalten statt der einen
      // kombinierten 'name'-Spalte. Alte localStorage-Einträge mit 'name'
      // werden im useEffect-Loader unten in 'vorname','nachname' migriert.
      { id: 'vorname', label: 'Vorname', alwaysVisible: true },
      { id: 'nachname', label: 'Nachname', alwaysVisible: true },
      { id: 'email', label: 'Email' },
      { id: 'jobTitle', label: 'Job Title' },
      { id: 'location', label: 'Standort' },
      // v24.33: Unternehmenszugehörigkeit / Rechtsträger als eigene Spalte.
      { id: 'company', label: isDe ? 'Unternehmen' : 'Company' },
    ];
    // v11.6: bei Split-Capacity die frei wählbaren Gruppen-Labels nutzen
    // (Fallback auf 'Starter-Typ' wenn keine Labels gesetzt sind).
    if (isSplit) {
      const lblA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || '';
      const lblB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || '';
      const colLabel = (lblA && lblB) ? `${lblA} / ${lblB}` : (isDe ? 'Gruppe' : 'Group');
      cols.push({ id: 'starterType', label: colLabel });
    }
    // v30.48: Startnummer nur anbieten, wenn tatsächlich eine importiert wurde.
    // Ohne Import wäre es eine Spalte voller „—" an jedem Event, das zufällig
    // die Spalte in der Liste hat.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (registrations.some(r => String((r as any).Startnummer || '').trim())) {
      cols.push({ id: 'startnummer', label: isDe ? 'Startnummer' : 'Bib number' });
    }
    cols.push({ id: 'status', label: 'Status' });
    cols.push({ id: 'date', label: 'Registriert am' });
    // v17.15/v17.17.1: Nachrück-Audit-Spalten — nur sichtbar wenn das
    // Event überhaupt eine Warteliste haben KANN (waitlistEnabled UND
    // maxParticipants > 0). Bei „Unbegrenzt"-Events kommt nie jemand auf
    // die Warteliste, deshalb sind die drei Audit-Spalten ohne Inhalt.
    // v19.11: Zusätzlich `hasWaitlistActivity` — Events OHNE echte Warteliste
    // (Default waitlistEnabled=true, aber niemand wartet/nachgerückt) zeigen die
    // leeren Audit-Spalten jetzt nicht mehr.
    if (selectedEvent?.waitlistEnabled && (selectedEvent?.maxParticipants || 0) > 0 && hasWaitlistActivity) {
      cols.push({ id: 'promotedDate', label: 'Nachgerückt am' });
      // v19.4: „Hat ersetzt" = die abgemeldete Person, deren Platz diese Person
      // übernommen hat. „Wurde ersetzt durch" wandert in die Abmeldungen-Tabelle
      // (gehört zur abgemeldeten Person, nicht zur aktiven).
      cols.push({ id: 'replaced', label: 'Hat ersetzt' });
    }
    cols.push({ id: 'registeredBy', label: 'Registriert von' });
    // v16.1: Team-Spalte — zeigt pro Teilnehmer den Team-Namen (falls Team-
    // Anmeldung aktiv und der TN in einem Team ist).
    if (selectedEvent?.teamRegistrationEnabled) {
      cols.push({ id: 'team', label: 'Team' });
    }
    if (userIds.length > 0) {
      // v11.56: Label aus dem ersten roommate-/user-Feld ableiten, statt hart
      // „Zimmerpartner" zu nennen. Wenn ein roommate-Feld existiert, nimm dessen
      // Label (User-Picker-Pairs); sonst das erste user-Feld; Fallback bleibt
      // der deutsche Default.
      const fields = selectedEvent?.eventSpecificFields || [];
      const firstRoommate = fields.filter(f => f.type === 'roommate' && f.label && f.label.trim())[0];
      const firstUser = fields.filter(f => f.type === 'user' && f.label && f.label.trim())[0];
      const roommateLabel = (firstRoommate?.label || firstUser?.label || 'Zimmerpartner').trim();
      cols.push({ id: 'roommate', label: roommateLabel });
    }
    // v14.11: Wenn ein Sub-Event selektiert ist, blenden wir die
    // Custom-Fields des Parent-Events (Pastel A) zusätzlich ein. Die
    // eigenen Sub-Event-Fields (Pastel B) folgen direkt danach. ID-
    // Präfix `cfp-` unterscheidet Parent- von Sub-Event-Feldern (`cf-`).
    const parentForCols: DeloitteEvent | null = (selectedEvent && selectedEvent.parentEventId)
      ? (allEvents.find(e => e.id === selectedEvent.parentEventId) || null)
      : null;
    if (parentForCols) {
      const ownIds = new Set((selectedEvent?.eventSpecificFields || []).map(f => f.id));
      // v19.10: 'roommate' (wie 'user') NICHT als generische Spalte ausgeben —
      // diese Felder werden bereits über die dedizierte „roommate"-Spalte (mit
      // Match-Badge) gerendert. Sonst erscheint das Feld DOPPELT (einmal mit
      // Match, einmal als roher „Name <email>"-Text).
      for (const f of (parentForCols.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'roommate' && f.label && f.label.trim())) {
        // Sub-Events erben Parent-Felder evtl. 1:1 (Wizard kopiert das beim
        // Anlegen). Nicht doppelt ausgeben, wenn das eigene Feld die
        // gleiche ID hat — in dem Fall reicht die Sub-Event-Spalte.
        if (ownIds.has(f.id)) continue;
        cols.push({ id: `cfp-${f.id}`, label: f.label });
      }
    }
    // v19.10: 'roommate'-Felder (wie 'user') hier ausschließen — sie haben
    // bereits die dedizierte „roommate"-Spalte mit Match-Badge. Vorher fehlte
    // `f.type !== 'roommate'`, deshalb erschien ein Zimmerpartner-Feld DOPPELT:
    // einmal als Match-Spalte, einmal als generische cf-Spalte mit rohem
    // „Nachname, Vorname <email>"-Text.
    for (const f of (selectedEvent?.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'roommate' && f.label && f.label.trim())) {
      cols.push({ id: `cf-${f.id}`, label: f.label });
    }
    cols.push({ id: 'action', label: 'Aktion', alwaysVisible: true });
    return cols;
  }, [
    selectedEvent?.id,
    selectedEvent?.parentEventId,
    selectedEvent?.durchstarterCapacity,
    selectedEvent?.funstarterCapacity,
    (selectedEvent?.eventSpecificFields || []).map(f => `${f.id}:${f.type}:${f.label}`).join(','),
    // v14.11: Parent-Custom-Fields als Dep
    (() => {
      if (!selectedEvent?.parentEventId) return '';
      const p = allEvents.find(e => e.id === selectedEvent.parentEventId);
      return (p?.eventSpecificFields || []).map(f => `${f.id}:${f.type}:${f.label}`).join(',');
    })(),
    // v19.11: Audit-Spalten-Sichtbarkeit hängt an der Warteliste-Aktivität.
    hasWaitlistActivity,
    // v30.48: Die Startnummern-Spalte erscheint erst nach dem Import — also
    // sobald irgendeine Zeile eine Nummer trägt.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    registrations.some(r => String((r as any).Startnummer || '').trim()),
  ]);

  // v26.44: gibt es überhaupt eine Roommate-Spalte? Steuert den
  // „Matches anzeigen"-Toggle, die Paar-Gruppierung der Teilnehmer-Tabelle
  // und die „Roommate-Match"-Spalte im Excel-Export.
  const hasRoommateColumn = availableColumns.some(c => c.id === 'roommate');

  const columnStorageKey = selectedEvent ? `dex_admin_columns_${selectedEvent.id}` : '';
  // localStorage-Load beim Event-Wechsel.
  React.useEffect(() => {
    if (!selectedEvent) { setColumnOrder([]); setHiddenColumns([]); return; }
    const allIds = availableColumns.map(c => c.id);
    try {
      const raw = localStorage.getItem(columnStorageKey);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.order) && Array.isArray(parsed.hidden)) {
          // v11.26: Migration alter Spaltenkonfigurationen — die zentrale
          // 'name'-Spalte wurde in 'vorname' + 'nachname' aufgeteilt. Wenn
          // ein gespeichertes Layout noch 'name' enthält, an gleicher
          // Position durch ['vorname','nachname'] ersetzen, damit der
          // User seine gewünschte Reihenfolge beibehält.
          const migratedOrder: string[] = [];
          for (const id of parsed.order as string[]) {
            if (id === 'name') {
              if (migratedOrder.indexOf('vorname') < 0) migratedOrder.push('vorname');
              if (migratedOrder.indexOf('nachname') < 0) migratedOrder.push('nachname');
            } else {
              migratedOrder.push(id);
            }
          }
          const knownOrder = migratedOrder.filter((id: string) => allIds.indexOf(id) >= 0);
          const missing = allIds.filter(id => knownOrder.indexOf(id) < 0);
          // v15.3: neu hinzugekommene Spalten (z.B. nach Custom-Field-Anlage
          // an einem bestehenden Event) VOR der „Aktion"-Spalte einreihen,
          // nicht hinten dran — sonst landen sie rechts neben den Buttons.
          const actionPos = knownOrder.indexOf('action');
          let mergedOrder = actionPos >= 0
            ? [...knownOrder.slice(0, actionPos), ...missing, ...knownOrder.slice(actionPos)]
            : [...knownOrder, ...missing];
          // v24.37: Die neu hinzugekommene 'company'-Spalte direkt HINTER
          // 'Standort' (location) einreihen statt ganz rechts vor den Aktionen
          // — gilt nur, solange der User sie noch nicht selbst positioniert hat
          // (also wenn sie frisch in `missing` steckt).
          if (missing.indexOf('company') >= 0) {
            mergedOrder = mergedOrder.filter(id => id !== 'company');
            const locIdx = mergedOrder.indexOf('location');
            if (locIdx >= 0) {
              mergedOrder.splice(locIdx + 1, 0, 'company');
            } else {
              const aPos = mergedOrder.indexOf('action');
              if (aPos >= 0) mergedOrder.splice(aPos, 0, 'company'); else mergedOrder.push('company');
            }
          }
          setColumnOrder(mergedOrder);
          // 'name' aus hidden auch herausfiltern (wenn jemals manuell hidden gesetzt wurde,
          // unwahrscheinlich da alwaysVisible — aber defensiv).
          setHiddenColumns(parsed.hidden.filter((id: string) => id !== 'name' && allIds.indexOf(id) >= 0));
          return;
        }
      }
    } catch { /* kaputte Config ignorieren */ }
    setColumnOrder(allIds);
    setHiddenColumns([]);
  }, [columnStorageKey, availableColumns.map(c => c.id).join(',')]);

  // Persistieren bei Änderungen.
  React.useEffect(() => {
    if (!columnStorageKey || columnOrder.length === 0) return;
    try {
      localStorage.setItem(columnStorageKey, JSON.stringify({ order: columnOrder, hidden: hiddenColumns }));
    } catch { /* quota exceeded oder private mode → ignorieren */ }
  }, [columnStorageKey, columnOrder.join(','), hiddenColumns.join(',')]);

  // Helper: Spalte ausblenden / wieder einblenden / verschieben.
  const hideColumn = (id: string): void => {
    const col = availableColumns.find(c => c.id === id);
    if (!col || col.alwaysVisible) return;
    if (hiddenColumns.indexOf(id) < 0) setHiddenColumns([...hiddenColumns, id]);
  };
  const showColumn = (id: string): void => {
    setHiddenColumns(hiddenColumns.filter(h => h !== id));
    if (columnOrder.indexOf(id) < 0) {
      const actionIdx = columnOrder.indexOf('action');
      const next = [...columnOrder];
      if (actionIdx >= 0) next.splice(actionIdx, 0, id); else next.push(id);
      setColumnOrder(next);
    }
  };
  const moveColumn = (id: string, direction: -1 | 1): void => {
    const idx = columnOrder.indexOf(id);
    if (idx < 0) return;
    const target = idx + direction;
    if (target < 0 || target >= columnOrder.length) return;
    if (columnOrder[target] === 'action') return;
    const next = [...columnOrder];
    [next[idx], next[target]] = [next[target], next[idx]];
    setColumnOrder(next);
  };

  /**
   * CSV Export für Teilnehmerlisten.
   * - 'deloitte': alle internen Felder (Anrede, Name, Email, Department, Location, JobTitle, Phone, Status, ...)
   * - 'b2run': Format exakt wie die offizielle B2Run-Köln-Meldedatei (16 Spalten laut B2RUN_KOELN_HEADERS: Nr., Anrede, Vorname, Nachname, E-Mail, Startblock, Zustimmung AGB & Datenschutzhinweise, Anonym, Gruppe, Straße/PLZ/Stadt (privat), Mobilnummer, Verwendung Infoservice, Altersklasse, Nordic Walker)
   */
  const exportCsv = (mode: 'deloitte' | 'b2run', audience: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled' = 'active'): void => {
    if (!selectedEvent) return;
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const audienceFilter = (r: SPRegistration): boolean => {
      if (audience === 'waitOnly') return r.Status === 'Warteliste';
      if (audience === 'activePlusWait') return ACTIVE.indexOf(r.Status) >= 0 || r.Status === 'Warteliste';
      // v20.4: alles inkl. Abgemeldete (Status-Spalte ist im Export enthalten).
      if (audience === 'withCancelled') return true;
      return ACTIVE.indexOf(r.Status) >= 0;
    };
    // v17.12: nach TeilnehmerID asc sortieren (vorher random / Status-Reihenfolge).
    const activeRegsForExport = registrations
      .filter(audienceFilter)
      .slice()
      .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
    if (activeRegsForExport.length === 0) { showAlert('Keine Teilnehmer zum Exportieren.'); return; }

    // v20.0 (Audit): toter CSV-Escaper `esc` entfernt — seit dem Umstieg auf
    // natives XLSX (v8.4) wurde er nie mehr aufgerufen.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parseCustom = (json: string): Record<string, any> => {
      try { return JSON.parse(json || '{}'); } catch { return {}; }
    };

    let headers: string[] = [];
    let rows: (string | number)[][] = [];

    if (mode === 'b2run') {
      // v26.48: Struktur exakt wie die OFFIZIELLE B2Run-Köln-Meldedatei
      // (Deloitte_Teilnehmer_-innen_b2run-koeln-<jahr>.xlsx) — 16 Spalten
      // inkl. „Straße" mit ß und der neuen Spalte „Nordic Walker".
      // Zentrale Spec in data/b2runKoeln.ts.
      headers = [...B2RUN_KOELN_HEADERS];
      rows = activeRegsForExport.map((r, idx) => {
        const cd = parseCustom(r.CustomData || '{}');
        const vorname = r.Vorname || (r.ParticipantName || '').split(' ').slice(0, -1).join(' ') || '';
        const nachname = r.Nachname || (r.ParticipantName || '').split(' ').slice(-1).join(' ') || '';
        return [
          idx + 1, // Nr. — laufende Nummer 1..n (die offizielle Datei nummeriert fortlaufend, NICHT TeilnehmerID)
          cd.b2run_geschlecht || mapAnredeToB2Run(r.Anrede), // 'männlich'/'weiblich'/'divers' (klein, wie Original)
          vorname,
          nachname,
          r.ParticipantEmail || '',
          cd.b2run_startblock || mapStarterTypeToStartblock(r.StarterType),
          cd.b2run_datenschutz ? 'Ja' : 'Nein',
          cd.b2run_anonym ? 'Ja' : 'Nein',
          cd.b2run_gruppe || '',
          '', // Straße und Hausnummer (privat) — nicht abgefragt, darf leer bleiben
          '', // PLZ (privat) — nicht abgefragt
          '', // Stadt (privat) — nicht abgefragt
          cd.b2run_mobilnummer || '',
          cd.b2run_infoservice ? 1 : 0, // Original-Datei nutzt 0/1 (Zahl), nicht Ja/Nein
          cd.b2run_altersklasse || B2RUN_KOELN_ALTERSKLASSE,
          cd.b2run_nordicwalker ? 'Ja' : 'Nein',
        ];
      });
    } else {
      // Deloitte View: alle internen Felder
      // v23.7: Team-Spalte nur bei Team-Events (oder wenn überhaupt eine
      // Team-Zuordnung existiert) — der frei benannte Begriff als Spaltenkopf.
      const includeTeam = !!selectedEvent.teamRegistrationEnabled || activeRegsForExport.some(r => !!r.TeamId);
      const teamHeader = selectedEvent.teamTermSingular || 'Team';
      // v30.48: Startnummer nur, wenn sie importiert wurde — sonst eine leere
      // Spalte in jedem Export.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const includeBib = activeRegsForExport.some(r => String((r as any).Startnummer || '').trim());
      headers = [
        'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'Email',
        'Department', 'Location', 'JobTitle', 'Phone',
        'Status', 'RegistrationDate',
        ...(includeBib ? ['Startnummer'] : []),
        ...(includeTeam ? [teamHeader, `${teamHeader}-Lead`] : []),
      ];
      // Dynamisch alle Custom Field Labels aus dem Event sammeln
      const customLabels: Array<{ id: string; label: string }> = (selectedEvent.eventSpecificFields || []).map(f => ({ id: f.id, label: f.label }));
      headers = headers.concat(customLabels.map(cf => cf.label));

      // v26.44: „Roommate-Match"-Spalte — nur wenn das Event eine Roommate-
      // Spalte hat. Paare werden über die VOLLE Export-Zeilenmenge berechnet
      // (nicht über den UI-Suchfilter), gleiche Dedupe-Logik wie die
      // „Matches anzeigen"-Gruppierung in der Teilnehmer-Tabelle.
      const roommatePairLabelByEmail: Record<string, string> = {};
      if (hasRoommateColumn) {
        const nameOf = (x: SPRegistration): string =>
          `${x.Vorname || ''} ${x.Nachname || ''}`.trim() || x.ParticipantName || x.ParticipantEmail || '';
        computeRoommatePairs(activeRegsForExport).forEach(([a, b], pi) => {
          const ea = (a.ParticipantEmail || '').trim().toLowerCase();
          const eb = (b.ParticipantEmail || '').trim().toLowerCase();
          roommatePairLabelByEmail[ea] = `Match ${pi + 1} (mit ${nameOf(b)})`;
          roommatePairLabelByEmail[eb] = `Match ${pi + 1} (mit ${nameOf(a)})`;
        });
        headers.push('Roommate-Match');
      }

      rows = activeRegsForExport.map(r => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const anyReg = r as any;
        const cd = parseCustom(r.CustomData || '{}');
        const base = [
          String(r.TeilnehmerID || ''),
          r.Anrede || '',
          r.Vorname || '',
          r.Nachname || '',
          r.ParticipantEmail || '',
          anyReg.Department || '',
          anyReg.Location || '',
          anyReg.JobTitle || '',
          anyReg.Phone || '',
          r.Status || '',
          r.RegistrationDate ? new Date(r.RegistrationDate).toLocaleString('de-DE') : '',
          ...(includeBib ? [String(anyReg.Startnummer || '')] : []),
          ...(includeTeam ? [r.TeamName || '', r.TeamLead ? 'Ja' : ''] : []),
        ];
        const customValues = customLabels.map(cf => {
          const v = cd[cf.id];
          if (v === undefined || v === null) return '';
          if (typeof v === 'boolean') return v ? 'Ja' : 'Nein';
          return String(v);
        });
        const row = base.concat(customValues);
        if (hasRoommateColumn) {
          row.push(roommatePairLabelByEmail[(r.ParticipantEmail || '').trim().toLowerCase()]
            || 'Ohne Preferred Roommate oder Match');
        }
        return row;
      });
    }

    const safeName = (selectedEvent.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
    // XLSX Export — natives Excel-Format, automatische Spalten-Breiten, keine
    // CSV-Escaping-Quirks. Gilt für beide Modi (Teilnehmerliste + B2Run).
    const aoa: (string | number)[][] = [headers, ...rows];
    // v26.48: Sheet-Name wie in der offiziellen Meldedatei („B2Run Köln <Jahr>",
    // ≤31 Zeichen — XLSX-Limit unkritisch). Jahr aus dem Event-Startdatum.
    const b2runYear = selectedEvent.startDate ? String(new Date(selectedEvent.startDate).getFullYear()) : '';
    const sheetName = mode === 'b2run' ? ('B2Run Köln ' + b2runYear).trim() : 'Teilnehmer';
    const filePrefix = mode === 'b2run' ? 'B2Run' : 'Teilnehmer';
    // v26.48: Bei B2Run-Köln-Events exakt der offizielle Dateiname des
    // Veranstalters (Deloitte_Teilnehmer_-innen_b2run-koeln-<jahr>.xlsx);
    // sonst bleibt das bisherige Namensschema.
    const fileName = mode === 'b2run' && isB2RunKoelnTitle(selectedEvent.title)
      ? `Deloitte_Teilnehmer_-innen_b2run-koeln${b2runYear ? '-' + b2runYear : ''}.xlsx`
      : `${filePrefix}_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`;

    // v20.0 (Audit): xlsx erst beim Export-Klick als Chunk nachladen — die
    // Bibliothek ist mit Abstand die schwerste Dependency und wird nur hier
    // gebraucht. Der .then/.catch-Pfad ersetzt das frühere try/catch.
    // v8.4: Manueller Blob-Download statt XLSX.writeFile. Im SPFx-Iframe-
    // Context ist saveAs/createObjectURL häufig blockiert (CORS / Sandbox-
    // Policies), wodurch der Download stillschweigend nicht startet. Mit
    // anchor.click() läuft das in jeder Browser-Umgebung zuverlässig.
    import('xlsx').then(XLSX => {
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const colWidths = headers.map((h, ci) => {
        const maxLen = Math.max(h.length, ...rows.map(r => String(r[ci] || '').length));
        return { wch: Math.min(40, Math.max(10, maxLen + 2)) };
      });
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = colWidths;
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 0);
    }).catch(err => {
      console.warn('[DEX] Excel-Export fehlgeschlagen:', err);
      showAlert(isDe
        ? 'Excel-Export fehlgeschlagen. Bitte Browser-Console prüfen.'
        : 'Excel export failed. Please check the browser console.');
    });
  };

  // v20.4: Excel-Export der konsolidierten Klammer-Ansicht. Baut EINE Datei
  // mit (wählbar) einem Matrix-Blatt — eine Zeile pro Person, Spalten =
  // Stammdaten + Klammer-Felder + pro Sub-Event der Status + dessen Feld-
  // Antworten — und/oder je einem eigenen Blatt pro gewähltem Sub-Event.
  // Datenquellen sind die bereits geladenen States (registrations = Klammer-
  // Zeilen, subEventRegsByEventId = Sub-Event-Listen) — kein Extra-Roundtrip.
  const exportConsolidatedExcel = (
    audience: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled',
    includeMatrix: boolean,
    subIds: string[]
  ): void => {
    if (!selectedEvent) return;
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const matches = (r: SPRegistration): boolean => {
      if (audience === 'waitOnly') return r.Status === 'Warteliste';
      if (audience === 'activePlusWait') return ACTIVE.indexOf(r.Status) >= 0 || r.Status === 'Warteliste';
      if (audience === 'withCancelled') return true;
      return ACTIVE.indexOf(r.Status) >= 0;
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parseCustom = (json: string): Record<string, any> => {
      try { return JSON.parse(json || '{}'); } catch { return {}; }
    };
    const fieldVal = (cd: Record<string, unknown>, id: string): string => {
      const v = cd[id];
      if (v === undefined || v === null) return '';
      if (typeof v === 'boolean') return v ? 'Ja' : 'Nein';
      return String(v);
    };
    const chosenChildren = consolidatedChildren.filter(c => subIds.indexOf(c.id) >= 0);
    const sheets: Array<{ name: string; headers: string[]; rows: string[][] }> = [];
    const sanitizeSheet = (s: string): string => (s || 'Blatt').replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31) || 'Blatt';

    if (includeMatrix) {
      const parentFields = (selectedEvent.eventSpecificFields || []).filter(f => f.label);
      type PersonRow = {
        vorname: string; nachname: string; email: string; jobTitle: string; location: string;
        teamName: string;
        parentCd: Record<string, unknown>;
        perChild: Record<string, SPRegistration | undefined>;
        hasMatch: boolean;
      };
      const persons: Record<string, PersonRow> = {};
      const ensurePerson = (r: SPRegistration): PersonRow => {
        const key = (r.ParticipantEmail || '').toLowerCase().trim();
        if (!persons[key]) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const anyReg = r as any;
          persons[key] = {
            vorname: r.Vorname || '', nachname: r.Nachname || '',
            email: r.ParticipantEmail || '',
            jobTitle: anyReg.JobTitle || '', location: anyReg.Location || '',
            teamName: '',
            parentCd: {}, perChild: {}, hasMatch: false,
          };
        }
        // v23.7: Team-Name aus der ersten Zeile übernehmen, die einen hat.
        if (r.TeamName && !persons[key].teamName) persons[key].teamName = r.TeamName;
        return persons[key];
      };
      for (const r of registrations) {
        const p = ensurePerson(r);
        p.parentCd = parseCustom(r.CustomData || '{}');
        if (matches(r)) p.hasMatch = true;
      }
      for (const child of consolidatedChildren) {
        const regs = subEventRegsByEventId[child.id] || [];
        for (const r of regs) {
          const p = ensurePerson(r);
          if (matches(r)) {
            p.perChild[child.id] = r;
            p.hasMatch = true;
          }
        }
      }
      // v23.7: Team-Spalte nur, wenn überhaupt Team-Zuordnungen existieren.
      const anyTeam = Object.keys(persons).some(k => !!persons[k].teamName);
      const teamHdr = selectedEvent.teamTermSingular || 'Team';
      const matrixHeaders: string[] = ['Vorname', 'Nachname', 'Email', 'JobTitle', 'Standort']
        .concat(anyTeam ? [teamHdr] : [])
        .concat(parentFields.map(f => f.label));
      for (const child of consolidatedChildren) {
        const short = shortSubEventTitle(child.title, selectedEvent.title) || child.title || '?';
        matrixHeaders.push(short);
        for (const f of (child.eventSpecificFields || []).filter(ff => ff.label)) {
          matrixHeaders.push(`${short}: ${f.label}`);
        }
      }
      const matrixRows: string[][] = Object.keys(persons)
        .map(k => persons[k])
        .filter(p => p.hasMatch)
        .sort((a, b) => (a.nachname || '').localeCompare(b.nachname || '', 'de') || (a.vorname || '').localeCompare(b.vorname || '', 'de'))
        .map(p => {
          const row: string[] = [p.vorname, p.nachname, p.email, p.jobTitle, p.location]
            .concat(anyTeam ? [p.teamName || ''] : [])
            .concat(parentFields.map(f => fieldVal(p.parentCd, f.id)));
          for (const child of consolidatedChildren) {
            const reg = p.perChild[child.id];
            row.push(reg ? (reg.Status || '') : '');
            const cd = reg ? parseCustom(reg.CustomData || '{}') : {};
            for (const f of (child.eventSpecificFields || []).filter(ff => ff.label)) {
              row.push(reg ? fieldVal(cd, f.id) : '');
            }
          }
          return row;
        });
      sheets.push({ name: 'Konsolidiert', headers: matrixHeaders, rows: matrixRows });
    }

    for (const child of chosenChildren) {
      const regs = (subEventRegsByEventId[child.id] || [])
        .filter(matches)
        .slice()
        .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
      const childFields = (child.eventSpecificFields || []).filter(f => f.label);
      // v23.7: Team-Spalte je Sub-Event-Blatt, wenn dort Team-Zuordnungen sind.
      const childAnyTeam = regs.some(r => !!r.TeamName);
      const childTeamHdr = selectedEvent.teamTermSingular || 'Team';
      const headers = ['TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'Email', 'Department', 'Location', 'JobTitle', 'Status', 'RegistrationDate']
        .concat(childAnyTeam ? [childTeamHdr] : [])
        .concat(childFields.map(f => f.label));
      const rows = regs.map(r => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const anyReg = r as any;
        const cd = parseCustom(r.CustomData || '{}');
        return [
          String(r.TeilnehmerID || ''), r.Anrede || '', r.Vorname || '', r.Nachname || '',
          r.ParticipantEmail || '', anyReg.Department || '', anyReg.Location || '', anyReg.JobTitle || '',
          r.Status || '', r.RegistrationDate ? new Date(r.RegistrationDate).toLocaleString('de-DE') : '',
        ].concat(childAnyTeam ? [r.TeamName || ''] : []).concat(childFields.map(f => fieldVal(cd, f.id)));
      });
      sheets.push({ name: sanitizeSheet(shortSubEventTitle(child.title, selectedEvent.title) || child.title || 'Sub-Event'), headers, rows });
    }

    if (sheets.length === 0) { showAlert(isDe ? 'Bitte mindestens die Matrix oder ein Sub-Event auswählen.' : 'Please select at least the matrix or one sub-event.'); return; }
    const safeName = (selectedEvent.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
    const fileName = `Konsolidiert_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    import('xlsx').then(XLSX => {
      const wb = XLSX.utils.book_new();
      const usedNames = new Set<string>();
      for (const sheet of sheets) {
        const aoa: (string | number)[][] = [sheet.headers, ...sheet.rows];
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const colWidths = sheet.headers.map((h, ci) => {
          const maxLen = Math.max(h.length, ...sheet.rows.map(r => String(r[ci] || '').length));
          return { wch: Math.min(40, Math.max(10, maxLen + 2)) };
        });
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (ws as any)['!cols'] = colWidths;
        // Doppelte Blattnamen entschärfen (xlsx verlangt eindeutige Namen).
        let name = sheet.name;
        let i = 2;
        while (usedNames.has(name)) { name = `${sheet.name.slice(0, 28)}_${i}`; i++; }
        usedNames.add(name);
        XLSX.utils.book_append_sheet(wb, ws, name);
      }
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 0);
    }).catch(err => {
      console.warn('[DEX] Konsolidierter Excel-Export fehlgeschlagen:', err);
      showAlert(isDe ? 'Excel-Export fehlgeschlagen. Bitte Browser-Console prüfen.' : 'Excel export failed. Please check the browser console.', { variant: 'error' });
    });
  };

  const handleSelectEvent = async (event: DeloitteEvent): Promise<void> => {
    // v18.24: aktuelle Card-Höhe einfrieren, BEVOR der State wechselt (DOM
    // zeigt noch den alten Stand) — verhindert das Zusammenklappen während
    // die Teilnehmer des neuen Events geladen werden.
    setReservedDetailHeight(detailCardRef.current?.offsetHeight);
    setSelectedEvent(event);
    // v10.19: NavigationContext.selectedEventId mitziehen, damit Header die
    // Page-ID granular ableiten kann (admin-center vs. admin-event) und der
    // Deep-Link-Kopier-Button immer die echte Item-ID des aktuell offenen
    // Events kennt. Skip falls bereits synchron — sonst doppelter History-
    // Eintrag beim Auto-Select via Deep-Link.
    if (selectedEventId !== event.id) {
      navigate('admin', event.id);
    }
    setIsLoadingRegs(true);
    setRegLoadError('');
    setDeniedSubEventLists([]);
    try {
      // v30.37: Auch hier zählt der HTTP-Status. Ein 403 kam bisher als leere
      // Liste an und wurde als „Noch keine Teilnehmer registriert." gerendert —
      // die freundlichste denkbare Lüge.
      let ownDenied = 0;
      const regs = await getAllRegistrations(event.id, st => {
        if (st === 401 || st === 403 || st === 404 || st === 0) ownDenied = st;
      });
      setRegistrations(regs);
      if (ownDenied) setRegLoadError(ACCESS_DENIED_MSG);
    } catch {
      setRegistrations([]);
      setRegLoadError('Teilnehmerliste konnte nicht geladen werden.');
    }
    setIsLoadingRegs(false);
    // Reservierung freigeben — der neue Inhalt steht jetzt, die Card nimmt
    // im selben Render die echte neue Höhe an (kein Zwischen-Kollaps).
    setReservedDetailHeight(undefined);
  };

  // v6.31: wenn navigation.selectedEventId gesetzt ist beim Mount (z.B. vom
  // Handbuch-Preview oder einem Deep-Link), direkt in die Detail-Ansicht
  // springen statt auf die Event-Auswahl-Liste.
  const didAutoSelectRef = React.useRef(false);
  React.useEffect(() => {
    if (didAutoSelectRef.current) return;
    if (!selectedEventId || selectedEvent) return;
    const match = adminEvents.find(e => e.id === selectedEventId);
    if (match) {
      didAutoSelectRef.current = true;
      handleSelectEvent(match).catch(() => { /* Fehler wird intern gesetzt */ });
    }
  }, [selectedEventId, adminEvents, selectedEvent]);

  // v22.40: Auto-Heilung stale Überbuchungs-Marker. Hat sich seit dem
  // „Überbuchung prüfen"-Lauf jemand abgemeldet, passt eine vorher als
  // überbucht markierte Person womöglich wieder in die Kapazität (oder ist
  // selbst nicht mehr aktiv). Solche `OverbookReview='Pending'`-Marker werden
  // hier still entfernt — sonst zeigt die Review-Box (und die orange Tabellen-
  // Markierung) jemanden als „über Kapazität", der längst regulär drinsteht.
  const overbookHealRef = React.useRef(false);
  React.useEffect(() => {
    if (overbookHealRef.current) { overbookHealRef.current = false; return; }
    if (!selectedEvent || !selectedEvent.subsiteUrl || !eventServiceRef) return;
    const flagged = registrations.filter(r => r.OverbookReview === 'Pending');
    if (flagged.length === 0) return;
    const isSplit = typeof selectedEvent.durchstarterCapacity === 'number'
      && typeof selectedEvent.funstarterCapacity === 'number'
      && ((selectedEvent.durchstarterCapacity || 0) > 0 || (selectedEvent.funstarterCapacity || 0) > 0);
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const groupOf = (r: SPRegistration): string => r.StarterType || r.PreferredStarterType || '';
    const keyOf = (r: SPRegistration): string => isSplit ? (groupOf(r) || '?') : 'all';
    const capOf = (key: string): number => {
      if (!isSplit) return selectedEvent.maxParticipants || 0;
      if (key === 'Durchstarter') return selectedEvent.durchstarterCapacity || 0;
      if (key === 'Funstarter') return selectedEvent.funstarterCapacity || 0;
      return 0;
    };
    const activeByGroup: Record<string, SPRegistration[]> = {};
    registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0).slice().sort((a, b) => a.Id - b.Id)
      .forEach(r => { const k = keyOf(r); (activeByGroup[k] = activeByGroup[k] || []).push(r); });
    // Stale = nicht (mehr) aktiv ODER Position passt wieder in die Kapazität.
    const stale = flagged.filter(r => {
      const k = keyOf(r); const cap = capOf(k); const bucket = activeByGroup[k] || [];
      const idx = bucket.findIndex(x => x.Id === r.Id);
      if (idx < 0) return true;
      return !(cap > 0 && (idx + 1) > cap);
    });
    if (stale.length === 0) return;
    let cancelled = false;
    (async () => {
      for (const r of stale) {
        try { await eventServiceRef.clearOverbookMark(selectedEvent.subsiteUrl, r.Id); }
        catch (err) { console.warn('[DEX] clearOverbookMark (auto-heal) failed:', err); }
      }
      if (cancelled) return;
      overbookHealRef.current = true; // nächsten Effekt-Lauf nach Reload überspringen
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    })().catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [registrations, selectedEvent?.id]);

  // Soft-Refresh-Sync: wenn `events` durch refreshEvents() aktualisiert wurde
  // (z.B. nach Event-deaktivieren, Event-aktivieren, Edit-save), den lokalen
  // `selectedEvent`-State aus der frischen Liste neu derivieren. Sonst bleibt
  // der Status-Badge (z.B. „Aktiv" vs „Entwurf") nach Toggle stale, weil
  // `selectedEvent` ein eigener useState ist und nicht aus `events` derived.
  React.useEffect(() => {
    if (!selectedEvent) return;
    const fresh = adminEvents.find(e => e.id === selectedEvent.id);
    if (fresh && fresh !== selectedEvent) {
      setSelectedEvent(fresh);
    }
  }, [adminEvents]);

  // v17.9: Map regId → Beitritts-Position (1-basiert, sortiert nach
  // v17.15: joinOrderById useMemo entfernt — wurde mit der Beitritts-#-
  // Spalte (v17.9) eingeführt, die der User in v17.10 wieder rausgeworfen
  // hat. Damit kein Hook mehr, der bei /joinOrder/ stale referenziert war.

  // v20.0 (Audit): ungenutzte Helper-Funktion getRegListUrl entfernt
  // (war seit Jahren nie aufgerufen, lieferte ohnehin nur `<base>/Lists`).

  // v20.1: Self-Check-in auto-aktivieren, falls das Event noch keinen aktiven
  // Token hat (Wizard-Toggle nie gesetzt): Token erzeugen + am Event
  // persistieren. Damit sind QR-PDF + Live-Anzeige grundsätzlich immer
  // verfügbar — der Klick auf die Aktion IST die Aktivierung.
  const ensureSelfCheckInReady = async (ev: DeloitteEvent): Promise<string | null> => {
    if (ev.selfCheckInEnabled && ev.selfCheckInToken) return ev.selfCheckInToken;
    const token = ev.selfCheckInToken || generateSelfCheckInToken();
    let ok = false;
    try {
      ok = await updateEvent(ev.id, { 'SelfCheckInEnabled': true, 'SelfCheckInToken': token });
    } catch { ok = false; }
    if (!ok) {
      showAlert(isDe
        ? 'Self-Check-in konnte nicht aktiviert werden (Speichern am Event fehlgeschlagen). Bitte erneut versuchen.'
        : 'Self check-in could not be activated (saving to the event failed). Please try again.');
      return null;
    }
    return token;
  };

  // v20.2: Self-Check-in-QR-Kachel unter dem Event-Logo + Erklär-/Einstell-Modal.
  // Die Kachel erscheint ab 5 Tagen vor Event-Start ODER sobald QR-Codes
  // versendet wurden; Klick öffnet das Modal mit großem QR, PDF-/Live-Aktionen
  // und dem editierbaren Check-in-Zeitfenster (Von/Bis).
  const [sciModalOpen, setSciModalOpen] = React.useState(false);
  const [sciModalQr, setSciModalQr] = React.useState('');
  const [sciToken, setSciToken] = React.useState('');
  const [sciFrom, setSciFrom] = React.useState('');
  const [sciTo, setSciTo] = React.useState('');
  const [sciBusy, setSciBusy] = React.useState(false);
  const [sciSaveMsg, setSciSaveMsg] = React.useState('');
  // v30.38: Der Mini-QR ist entfallen. Er wurde beim Öffnen JEDES Events
  // erzeugt (qrcode-Chunk + Canvas), nur um in der Kachel als 64-px-Vorschau zu
  // stehen — die Kachel führt jetzt in den Einstieg „QR-Codes und Check-In" und
  // zeigt ein Icon. Eine Vorschau des Self-Check-in-Codes an einer Stelle, an der
  // man sich noch gar nicht für Self-Check-in entschieden hat, nahm die Auswahl
  // vorweg; den großen QR gibt es im Modal weiterhin (`sciModalQr`).
  const isoToLocalInput = (iso?: string): string => {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    const p = (n: number): string => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
  };
  const openSelfCheckInModal = async (): Promise<void> => {
    if (!selectedEvent || sciBusy) return;
    setSciBusy(true);
    try {
      const token = await ensureSelfCheckInReady(selectedEvent);
      if (!token) return;
      setSciToken(token);
      // v20.3: Von/Bis immer vorbelegen — gespeicherte Werte ODER der
      // Standard (2 Stunden vor Event-Start bis Event-Ende). Der Standard
      // gilt auch zur Laufzeit, solange nichts anderes gespeichert ist
      // (isWithinCheckInWindow) — Anzeige und Verhalten sind damit deckungsgleich.
      const def = defaultCheckInWindow(selectedEvent.startDate, selectedEvent.endDate);
      setSciFrom(isoToLocalInput(selectedEvent.selfCheckInFrom) || (def.opensAt ? isoToLocalInput(def.opensAt.toISOString()) : ''));
      setSciTo(isoToLocalInput(selectedEvent.selfCheckInTo) || (def.closesAt ? isoToLocalInput(def.closesAt.toISOString()) : ''));
      setSciSaveMsg('');
      try {
        const QRCode = await import('qrcode');
        setSciModalQr(await QRCode.toDataURL(buildStaticCheckInUrl(token), { width: 560, margin: 1 }));
      } catch { setSciModalQr(''); }
      setSciModalOpen(true);
    } finally { setSciBusy(false); }
  };
  // v20.3: Der Status-Badge neben dem Event-Titel ist klickbar — Aktiv ⇄
  // Entwurf (ersetzt den früheren Eintrag im Aktionen-Menü). Gleiche Logik
  // wie der alte v11.89-Toggle: IsFictive flippen, beim Live-Schalten
  // Legacy-EventStatus auf 'Active' setzen.
  const toggleDraftStatus = async (): Promise<void> => {
    if (!selectedEvent) return;
    const isDraft = !!selectedEvent.isFictive;
    // v22.15: Abgeschlossen/Abgesagt → zurück auf Aktiv (Reaktivierung).
    // Vorher waren diese Zustände eine Sackgasse — auch für Admins.
    const isFinalState = !isDraft && (selectedEvent.status === 'Completed' || selectedEvent.status === 'Cancelled');
    if (isFinalState) {
      const fromLabel = isDe ? localizeStatus(selectedEvent.status) : selectedEvent.status;
      if (!(await confirmDialog(
        isDe
          ? `Event von „${fromLabel}" wieder auf Aktiv setzen? Danach ist es für die Berechtigten wieder sichtbar und buchbar. Hinweis: Liegt das End-Datum in der Vergangenheit, setzt der automatische Aufräum-Lauf das Event beim nächsten App-Start erneut auf „Abgeschlossen" — dann zuerst das Datum korrigieren.`
          : `Set event from "${fromLabel}" back to Active? It will be visible and bookable for eligible users again. Note: if the end date is in the past, the automatic cleanup will set it back to "Completed" on the next app start — fix the date first in that case.`,
        { title: isDe ? 'Event reaktivieren' : 'Reactivate event', confirmLabel: isDe ? 'Auf Aktiv setzen' : 'Set to Active' },
      ))) return;
      const ok = await updateEvent(selectedEvent.id, { 'EventStatus': 'Active' });
      if (ok) {
        setSelectedEvent(prev => prev ? { ...prev, status: 'Active' } : prev);
        await refreshEvents();
      } else {
        showAlert(isDe
          ? 'Der Status konnte nicht geändert werden. Vermutlich fehlen dir Schreibrechte auf der Event-Liste — bitte einen Haupt-Organizer oder Admin den Status umschalten lassen.'
          : 'The status could not be changed. You probably lack write permission on the event list — please ask a main organizer or admin to switch the status.', { variant: 'error' });
      }
      return;
    }
    const nextIsFictive = !isDraft;
    const confirmMsg = nextIsFictive
      ? (isDe ? 'Event auf "Entwurf" zurücksetzen? Reguläre User sehen das Event danach nicht mehr.' : 'Reset event to "draft"? Regular users will no longer see the event afterwards.')
      : (isDe ? 'Event live schalten? Alle Berechtigten können sich danach anmelden.' : 'Publish event? All eligible users can register afterwards.');
    if (!(await confirmDialog(confirmMsg, { title: isDe ? 'Event-Status ändern' : 'Change event status', confirmLabel: nextIsFictive ? (isDe ? 'Auf Entwurf setzen' : 'Set to draft') : (isDe ? 'Live schalten' : 'Publish') }))) return;
    const patch: Record<string, unknown> = { 'IsFictive': nextIsFictive };
    if (!nextIsFictive) patch['EventStatus'] = 'Active';
    const ok = await updateEvent(selectedEvent.id, patch);
    if (ok) {
      // Badge sofort umschalten — selectedEvent ist lokaler State und wird
      // durch refreshEvents nicht automatisch ersetzt.
      setSelectedEvent(prev => prev ? { ...prev, isFictive: nextIsFictive, ...(nextIsFictive ? {} : { status: 'Active' }) } : prev);
      // v22.67: Beim Live-Schalten eines Events mit Sub-Events werden die
      // Sub-Events automatisch mit live geschaltet (Entwurf → Aktiv) — sonst
      // bliebe das Event sichtbar, aber die Sub-Events wären für Teilnehmer
      // nicht buchbar.
      if (!nextIsFictive) {
        for (const c of childEventsOf(selectedEvent.id)) {
          if (c.isFictive) {
            try { await updateEvent(c.id, { 'IsFictive': false, 'EventStatus': 'Active' }); } catch { /* best-effort */ }
          }
        }
      }
      await refreshEvents();
    } else {
      // v22.14: vorher scheiterte der Klick STUMM — der Organizer dachte,
      // der Status lasse sich nicht ändern, ohne zu erfahren warum.
      showAlert(isDe
        ? 'Der Status konnte nicht geändert werden. Vermutlich fehlen dir Schreibrechte auf der Event-Liste (z.B. als Co-Organizer ohne Organizer-Rolle) — bitte einen Haupt-Organizer oder Admin den Status umschalten lassen.'
        : 'The status could not be changed. You probably lack write permission on the event list (e.g. co-organizer without the organizer role) — please ask a main organizer or admin to switch the status.', { variant: 'error' });
    }
  };

  // v22.5: Einladungsmail — Default-Texte bauen, Entwurf laden/speichern
  // (localStorage pro Event), Modal öffnen, zurücksetzen.
  const inviteDraftKey = (id: string): string => `dex_invite_draft_${id}`;
  // v26.89: B2Run-Köln-Events bekommen einen eigenen, dynamischen Einladungs-
  // text-Vorschlag (bilingual DE + EN) — mit Datum, Ort und Platzzahl aus dem
  // Event. Der Organizer kann ihn wie jeden anderen Entwurf frei überschreiben.
  const buildB2RunKoelnInviteDefaults = (ev: DeloitteEvent, appUrl: string, signatureNames: string): { subject: string; heading: string; subheading: string; body: string } => {
    const start = ev.startDate ? new Date(ev.startDate) : null;
    const validStart = start && !isNaN(start.getTime()) ? start : null;
    const dateDe = validStart ? validStart.toLocaleDateString('de-DE', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' }) : '';
    const dateEn = validStart ? validStart.toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' }) : '';
    const venue = (ev.location || '').trim() || 'RheinEnergieStadion';
    const plaetze = (ev.maxParticipants && ev.maxParticipants > 0) ? ev.maxParticipants : 100;
    const dateLineDe = dateDe || 'Datum folgt';
    const dateLineEn = dateEn || 'date to follow';
    // „DEX App" als grün gestylter Link (statt der langen URL im Klartext).
    const appLinkDe = `<a href="${appUrl}" style="color:#86bc25;font-weight:600;">DEX App</a>`;
    const appLinkEn = `<a href="${appUrl}" style="color:#86bc25;font-weight:600;">DEX app</a>`;
    const b2runSite = '<a href="https://www.b2run.de" style="color:#86bc25;font-weight:600;">B2RUN Website</a>';
    const de = `
<p>Liebes Team,</p>
<p>es ist so weit: wir haben unseren Standort wieder für den <strong>B2Run Firmenlauf</strong> angemeldet. Es werden vorerst <strong>${plaetze} Startplätze</strong> zur Verfügung stehen. Diese werden nach dem Motto „First come, first run" vergeben.</p>
<p>Die Startgebühren inkl. einer Spende an Menschen für Menschen, die Kosten für ein Laufshirt sowie ein Teamzelt mit einer kleinen Verpflegung nach dem Lauf werden dabei von Deloitte übernommen.</p>
<p style="text-align:center;"><strong>Die Anmeldung ist ab sofort über die ${appLinkDe} möglich.</strong></p>
<p><strong>Wichtige Hinweise:</strong></p>
<ul>
  <li>Falls ihr es nicht unter die ersten ${plaetze} schaffen solltet, meldet euch bitte trotzdem über die App an. Ihr werdet automatisch der Reihe nach auf eine Warteliste gesetzt.</li>
  <li>Falls ihr aus wichtigen Gründen nicht am Lauf teilnehmen könnt, sagt eure Teilnahme bitte frühzeitig über die App („My Events") wieder ab. Die Plätze werden automatisch von der Warteliste – der Reihe nach – vergeben und ihr erhaltet eine automatische E-Mail.</li>
  <li>Bitte meldet euch nur an, wenn ihr auch wirklich am B2RUN teilnehmen könnt und möchtet. <strong>Wir zahlen für jede Anmeldung eine Teilnehmergebühr, die wir im Falle eines No-Shows nicht erstattet bekommen.</strong></li>
</ul>
<p><strong>Infos zum Lauf/Event:</strong></p>
<ul>
  <li>${dateLineDe} am <strong>${venue}</strong></li>
  <li>Beginn der Veranstaltung: 15:00 Uhr</li>
  <li>Teamtreff Deloitte, Startnummernübergabe und Aufwärmen: 16:00 Uhr</li>
  <li>Startzeit Deloitte: 17:00 Uhr</li>
  <li>Distanz: 5,3 km</li>
  <li>Anschließend: Get-Together, Teamfotos und Catering sowie Afterparty im Stadioninnenraum (ab 20:00 Uhr)</li>
</ul>
<p><em>Genauere Infos zu den Zeiten und Treffpunkten werden wir euch Mitte / Ende August mitteilen.</em></p>
<p><strong>Startfelder</strong></p>
<p>Beim B2Run gibt es zwei Startfelder: „Funstarter" und „Durchstarter".</p>
<ul>
  <li>Das <strong>Durchstarter</strong>-Feld ist für schnelle und ambitionierte Läufer:innen gedacht, die „freie Bahn" haben möchten (Richtwerte Männer &lt;4 Min/km, Frauen &lt;5 Min/km).</li>
  <li>Für das <strong>Funstarter</strong>-Feld gibt es keine Richtwerte – hier steht der Laufspaß im Vordergrund.</li>
  <li>Sofern jemand am Durchstarter-Lauf teilnehmen möchte, wählt dies entsprechend bei der Anmeldung aus. Die Laufstrecke für beide Startfelder beträgt 5,3 Kilometer.</li>
</ul>
<p><strong>Laufshirts</strong></p>
<p>Für jede:n Läufer:in gibt es ein Deloitte-Laufshirt. Wählt bitte bei der Anmeldung eure Größe aus.</p>
<p>Weitere Informationen sind auf der ${b2runSite} zu finden.</p>
<p>Bei Fragen wendet euch bitte direkt an unser Gruppenpostfach.</p>
<p>Auf die Plätze, fertig, los.</p>
<p>Mit sportlichen Grüßen<br />${signatureNames}</p>`.trim();
    const en = `
<p>Dear team,</p>
<p>The time has come: we have registered our location for the <strong>B2Run company run</strong> again. For now <strong>${plaetze} starting places</strong> will be available, allocated on a „first come, first run" basis.</p>
<p>The registration fees (incl. a donation to Menschen für Menschen), the cost of a running shirt and a team tent with light refreshments after the run are covered by Deloitte.</p>
<p style="text-align:center;"><strong>Registration is now open via the ${appLinkEn}.</strong></p>
<p><strong>Important information:</strong></p>
<ul>
  <li>If you don't make it into the first ${plaetze}, please register via the app anyway. You will automatically be placed on a waiting list in order of registration.</li>
  <li>If you are unable to take part for important reasons, please cancel your participation early via the app („My Events"). Places are allocated automatically from the waiting list – in order – and you will receive an automatic email.</li>
  <li>Please only register if you can really take part in the B2RUN. <strong>We pay a participation fee for each registration that is not refunded in the event of a no-show.</strong></li>
</ul>
<p><strong>Event details:</strong></p>
<ul>
  <li>${dateLineEn} at <strong>${venue}</strong></li>
  <li>Start of the event: 3:00 p.m.</li>
  <li>Deloitte team meeting, race-number handout and warm-up: 4:00 p.m.</li>
  <li>Deloitte start time: 5:00 p.m.</li>
  <li>Distance: 5.3 km</li>
  <li>Afterwards: get-together, team photos and catering plus after-party inside the stadium (from 8:00 p.m.)</li>
</ul>
<p><em>We will share more detailed information about times and meeting points in mid/late August.</em></p>
<p><strong>Starting fields</strong></p>
<p>B2Run has two starting fields: „Funstarter" and „Durchstarter".</p>
<ul>
  <li>The <strong>Durchstarter</strong> field is intended for fast and ambitious runners who want a „clear track" (guideline times: men &lt;4 min/km, women &lt;5 min/km).</li>
  <li>The <strong>Funstarter</strong> field has no guideline times – the focus here is on having fun.</li>
  <li>If you would like to take part in the Durchstarter run, please select this when registering. The distance for both fields is 5.3 kilometres.</li>
</ul>
<p><strong>Running shirts</strong></p>
<p>Every runner receives a Deloitte running shirt. Please select your size when registering.</p>
<p>Further information can be found on the ${b2runSite}.</p>
<p>If you have any questions, please contact our group mailbox directly.</p>
<p>On your marks, get set, go.</p>
<p>With sporting regards<br />${signatureNames}</p>`.trim();
    // v26.90: Zweisprachig — Hauptsprache oben, die jeweils andere Version unten
    // (per Trennlinie abgesetzt), wie es die B2Run-Kommunikation üblicherweise macht.
    const divider = '<p style="margin:28px 0 20px;border-top:1px solid #d0d0ce;"></p>';
    return {
      subject: isDe ? `Einladung zu ${ev.title}` : `Invitation to ${ev.title}`,
      heading: isDe ? `Einladung zu ${ev.title}` : `Invitation to ${ev.title}`,
      subheading: isDe ? `${dateLineDe} · ${plaetze} Startplätze` : `${dateLineEn} · ${plaetze} starting places`,
      body: isDe ? `${de}\n${divider}\n${en}` : `${en}\n${divider}\n${de}`,
    };
  };
  const buildInviteDefaults = (ev: DeloitteEvent): { subject: string; heading: string; subheading: string; body: string } => {
    const appUrl = `${siteUrl}/SitePages/DEX.aspx?env=WebView`;
    const linkHtml = `<a href="${appUrl}" style="color:#86bc25;font-weight:600;">${appUrl}</a>`;
    const orgList = (ev.organizers || []).map(s => (s || '').trim()).filter(Boolean);
    const teamLine = isDe ? `Das ${ev.title} Orga Team` : `The ${ev.title} Organizer Team`;
    const signatureNames = orgList.length > 0 ? `${teamLine}<br />${orgList.join('<br />')}` : teamLine;
    // v26.89: B2Run-Köln-Events erhalten den spezialisierten Vorschlag.
    if (isB2RunKoelnTitle(ev.title)) {
      return buildB2RunKoelnInviteDefaults(ev, appUrl, signatureNames);
    }
    const body = isDe
      ? `<p>Hallo,</p>\n<p>wir laden dich herzlich zum Event <strong>${ev.title}</strong> ein.</p>\n<p>Du kannst dich ab sofort über unsere Event-Plattform anmelden:</p>\n<p>${linkHtml}</p>\n<p>Falls du dich im Nachgang doch nicht beteiligen kannst, ist eine <strong>Abmeldung jederzeit über dieselbe Plattform</strong> möglich — bitte gib uns rechtzeitig Bescheid, damit Wartelisten-Plätze nachrücken können.</p>\n<p>Bei Rückfragen meld dich gern bei uns.</p>\n<p>Viele Grüße<br />${signatureNames}</p>`
      : `<p>Hello,</p>\n<p>we would like to invite you to the event <strong>${ev.title}</strong>.</p>\n<p>You can register via our event platform:</p>\n<p>${linkHtml}</p>\n<p>If you change your mind, you can <strong>cancel anytime via the same platform</strong> — please let us know early so people on the waitlist can move up.</p>\n<p>Feel free to reach out if you have any questions.</p>\n<p>Best regards<br />${signatureNames}</p>`;
    return {
      subject: isDe ? `Einladung: ${ev.title}` : `Invitation: ${ev.title}`,
      heading: isDe ? `Einladung zu ${ev.title}` : `Invitation to ${ev.title}`,
      subheading: '',
      body,
    };
  };
  const applyInviteDraftOrDefaults = (ev: DeloitteEvent): void => {
    inviteHydratingRef.current = true;
    let loaded: { subject?: string; heading?: string; subheading?: string; body?: string; target?: string } | null = null;
    try {
      const raw = window.localStorage.getItem(inviteDraftKey(ev.id));
      if (raw) loaded = JSON.parse(raw);
    } catch { /* localStorage evtl. blockiert */ }
    const def = buildInviteDefaults(ev);
    setInviteSubject(loaded && typeof loaded.subject === 'string' ? loaded.subject : def.subject);
    setInviteHeading(loaded && typeof loaded.heading === 'string' ? loaded.heading : def.heading);
    setInviteSubheading(loaded && typeof loaded.subheading === 'string' ? loaded.subheading : def.subheading);
    setInviteBody(loaded && typeof loaded.body === 'string' ? loaded.body : def.body);
    setInviteTarget(loaded && loaded.target === 'audience' ? 'audience' : 'organizer');
    // v29.37: dito für Einladung und Erinnerung.
    setInviteHeaderImage(p => ({ ...p, ...eventHeaderImageLayout(ev.emailTemplateOverrides) }));
    // Hydration-Flag im nächsten Tick freigeben, damit das Auto-Speichern erst
    // auf echte Nutzer-Edits reagiert (nicht auf das initiale Laden).
    window.setTimeout(() => { inviteHydratingRef.current = false; }, 0);
  };
  /**
   * v29.32: Verteiler des Events in einzelne Personen auflösen — derselbe Weg
   * wie beim Einladungsversand (Graph-Gruppenmitglieder, Fallback auf die beim
   * Event-Speichern eingefrorene Liste). Ausgeschlossene Adressen fliegen raus,
   * denn wer auf der Ausschluss-Liste steht, sieht das Event nie.
   *
   * Wichtig: Ein reiner STANDORT-Filter lässt sich hier nicht abzählen — dafür
   * müsste man das ganze Verzeichnis lesen. Die Zeile sagt das dann auch, statt
   * eine Zahl zu erfinden.
   */
  const resolveAudienceEmails = async (ev: DeloitteEvent): Promise<AudiencePerson[]> => {
    const entries = (ev.audienceFilter || []).map(s => (s || '').trim()).filter(Boolean);
    const excluded = new Set((ev.excludedUsers || []).map(e => (e || '').toLowerCase().trim()).filter(Boolean));
    const out: AudiencePerson[] = [];
    const seen = new Set<string>();
    // v29.36: Name/Position/Standort mitnehmen, wenn die Verteiler-Abfrage sie
    // liefert — der Nachfass-Schritt zeigt Personen, keine Adressliste.
    const push = (e: string, extra?: Partial<AudiencePerson>): void => {
      const lc = (e || '').trim().toLowerCase();
      if (lc && lc.indexOf('@') > 0 && !seen.has(lc) && !excluded.has(lc)) {
        seen.add(lc);
        out.push({ email: lc, displayName: extra?.displayName || '', jobTitle: extra?.jobTitle || '', location: extra?.location || '' });
      }
    };
    for (const entry of entries) {
      if (entry.indexOf('@') < 0) continue; // Standort-Pattern — nicht auflösbar
      try {
        const grp = await getGroupMembers(entry);
        if (grp && grp.members && grp.members.length > 0) {
          grp.members.forEach(m => push(m.email, { displayName: m.displayName, jobTitle: m.jobTitle, location: m.location }));
        } else push(entry);
      } catch { push(entry); }
    }
    if (out.length === 0) (ev.audienceResolvedEmails || []).forEach(e => push(e));
    return out;
  };

  /**
   * v29.32: „Wer hat noch nicht geantwortet?" — die aufgelöste Sichtbarkeits-
   * Liste minus aller Personen, die sich bereits geäußert haben. Als Antwort
   * zählt JEDE Zeile im Event: Anmeldung, Warteliste, Check-in, Abmeldung und
   * die proaktive Absage („Ich nehme nicht teil"). Bei einer Klammer zählen
   * auch die Zeilen der Sub-Events — wer dort gebucht hat, hat geantwortet.
   *
   * Das Ergebnis geht in den bestehenden Einladungs-Dialog (Modus „Nachfassen",
   * Empfängerliste editierbar, Mailtext editierbar). Bewusst KEIN zweiter
   * Versand-Dialog daneben — der bestehende kann das alles bereits.
   */
  const openPendingReminder = async (): Promise<void> => {
    if (!selectedEvent || pendingCheckBusy) return;
    setPendingCheckBusy(true);
    try {
      const audience = visibilityResolved || await resolveAudienceEmails(selectedEvent);
      if (!visibilityResolved) setVisibilityResolved(audience);
      if (audience.length === 0) {
        showAlert(isDe
          ? 'Für dieses Event kennt DEX keine Namen. Das ist so, wenn die Sichtbarkeit nur über den Standort läuft — dahinter steht keine Liste einzelner Personen. Trage in Schritt 3 des Event-Edits zusätzlich einen Mailverteiler oder einzelne Personen ein, dann kann DEX nachfassen.'
          : 'DEX does not know any names for this event. That is the case when visibility runs via location only — there is no list of individual people behind it. Add a distribution list or individual people in step 3 of the event edit, then DEX can follow up.',
          { variant: 'info' });
        return;
      }
      // Wer hat schon geantwortet? Alle Zeilen des Events + (bei einer Klammer)
      // der Sub-Events, unabhängig vom Status.
      const decided = new Set<string>();
      registrations.forEach(r => { const e = (r.ParticipantEmail || '').toLowerCase().trim(); if (e) decided.add(e); });
      Object.keys(subEventRegsByEventId || {}).forEach(k => {
        (subEventRegsByEventId[k] || []).forEach(r => {
          const e = (r.ParticipantEmail || '').toLowerCase().trim(); if (e) decided.add(e);
        });
      });
      // Organizer-Team zählt nicht als offener Fall — es organisiert das Event.
      const team = new Set<string>([
        ...(selectedEvent.organizerEmails || []),
        ...(selectedEvent.coOrganizerEmails || []),
      ].map(e => (e || '').toLowerCase().trim()).filter(Boolean));
      const pending = audience.filter(p => !decided.has(p.email) && !team.has(p.email));
      if (pending.length === 0) {
        showAlert(isDe
          ? `Alle ${audience.length} Personen, die das Event sehen können, haben bereits geantwortet — angemeldet, abgemeldet oder abgesagt. Es gibt niemanden zum Erinnern.`
          : `All ${audience.length} people who can see this event have already responded — registered, cancelled or declined. There is nobody to remind.`,
          { variant: 'success' });
        return;
      }
      // v29.36: ERST die Übersicht, wer fehlt (Foto, Name, Position) — die Mail
      // kommt im zweiten Schritt. Vorher landete man direkt im Mail-Dialog und
      // musste einer Adressliste glauben, ohne zu sehen, wen man da anschreibt.
      setPendingPeople({ people: pending, reachable: audience.length });
    } catch (err) {
      showAlert((isDe ? 'Prüfung fehlgeschlagen: ' : 'Check failed: ') + String((err as Error)?.message || err), { variant: 'error' });
    } finally {
      setPendingCheckBusy(false);
    }
  };

  const openInviteModal = (): void => {
    if (!selectedEvent) return;
    applyInviteDraftOrDefaults(selectedEvent);
    // v28.37: Anpassungen der letzten Runde nicht mitschleppen und die
    // bereits verschickten Einladungen im Hintergrund nachladen (für den
    // Modus „Nur an noch nicht Eingeladene").
    setInviteCustomEmails(null);
    setInviteAddInput('');
    setInviteAudienceOpen(false);
    setInvitedLc(null);
    if (eventServiceRef) {
      eventServiceRef.getInvitedRecipients(selectedEvent.id)
        .then(list => setInvitedLc(new Set(list)))
        .catch(() => setInvitedLc(new Set<string>()));
    } else {
      setInvitedLc(new Set<string>());
    }
    setShowInviteModal(true);
  };
  const resetInviteDraft = (): void => {
    if (!selectedEvent) return;
    try { window.localStorage.removeItem(inviteDraftKey(selectedEvent.id)); } catch { /* */ }
    inviteHydratingRef.current = true;
    const def = buildInviteDefaults(selectedEvent);
    setInviteSubject(def.subject);
    setInviteHeading(def.heading);
    setInviteSubheading(def.subheading);
    setInviteBody(def.body);
    setInviteDraftSaved(false);
    window.setTimeout(() => { inviteHydratingRef.current = false; }, 0);
  };
  // v22.5/v22.6: expliziter „Entwurf speichern"-Klick — schreibt den aktuellen
  // Stand sofort in localStorage und zeigt kurz „Gespeichert".
  const saveInviteDraft = (): void => {
    if (!selectedEvent) return;
    try {
      window.localStorage.setItem(inviteDraftKey(selectedEvent.id), JSON.stringify({
        subject: inviteSubject, heading: inviteHeading, subheading: inviteSubheading,
        body: inviteBody, target: inviteTarget,
      }));
      setInviteDraftSaved(true);
      window.setTimeout(() => setInviteDraftSaved(false), 2500);
    } catch { /* localStorage evtl. blockiert */ }
  };
  // Auto-Speichern, solange das Modal offen ist — beim nächsten Öffnen wird der
  // Entwurf wiederhergestellt.
  React.useEffect(() => {
    if (!showInviteModal || !selectedEvent || inviteHydratingRef.current) return;
    try {
      window.localStorage.setItem(inviteDraftKey(selectedEvent.id), JSON.stringify({
        subject: inviteSubject, heading: inviteHeading, subheading: inviteSubheading,
        body: inviteBody, target: inviteTarget,
      }));
    } catch { /* */ }
  }, [showInviteModal, selectedEvent, inviteSubject, inviteHeading, inviteSubheading, inviteBody, inviteTarget]);

  // v22.9: Massenmail-Entwurf — Default-Texte, laden/speichern (localStorage pro
  // Event), Picker öffnen, zurücksetzen, Testmail an die Organizer.
  const massmailDraftKey = (id: string): string => `dex_massmail_draft_${id}`;
  const buildMassmailDefaults = (ev: DeloitteEvent): { subject: string; heading: string; body: string } => ({
    subject: `${ev.title} - Info`,
    heading: ev.title,
    body: '',
  });
  const applyMassmailDraftOrDefaults = (ev: DeloitteEvent): void => {
    massmailHydratingRef.current = true;
    let loaded: { subject?: string; heading?: string; subheading?: string; body?: string } | null = null;
    try {
      const raw = window.localStorage.getItem(massmailDraftKey(ev.id));
      if (raw) loaded = JSON.parse(raw);
    } catch { /* localStorage evtl. blockiert */ }
    const def = buildMassmailDefaults(ev);
    setEmailSubject(loaded && typeof loaded.subject === 'string' ? loaded.subject : def.subject);
    setEmailHeading(loaded && typeof loaded.heading === 'string' ? loaded.heading : def.heading);
    setMassmailSubheading(loaded && typeof loaded.subheading === 'string' ? loaded.subheading : '');
    setEmailBody(loaded && typeof loaded.body === 'string' ? loaded.body : def.body);
    // v29.37: Kopfbild-Größe aus dem Event übernehmen statt fest 180/30/30.
    setMassmailHeaderImage(p => ({ ...p, ...eventHeaderImageLayout(ev.emailTemplateOverrides) }));
    window.setTimeout(() => { massmailHydratingRef.current = false; }, 0);
  };
  const openMassmailPicker = (): void => {
    if (selectedEvent) applyMassmailDraftOrDefaults(selectedEvent);
    setMassmailAudience('active');
    setMassmailPasteRaw('');
    setMassmailTestMsg(null);
    setMassmailMode('pick');
  };
  const resetMassmailDraft = (): void => {
    if (!selectedEvent) return;
    try { window.localStorage.removeItem(massmailDraftKey(selectedEvent.id)); } catch { /* */ }
    massmailHydratingRef.current = true;
    const def = buildMassmailDefaults(selectedEvent);
    setEmailSubject(def.subject);
    setEmailHeading(def.heading);
    setMassmailSubheading('');
    setEmailBody(def.body);
    // v30.51: „Zurücksetzen" setzt auch das zusätzliche CC zurück — sonst
    // bliebe ein Verteiler stehen, den niemand mehr im Text sieht.
    setMassmailCc([]);
    setMassmailDraftSaved(false);
    window.setTimeout(() => { massmailHydratingRef.current = false; }, 0);
  };
  const saveMassmailDraft = (): void => {
    if (!selectedEvent) return;
    try {
      window.localStorage.setItem(massmailDraftKey(selectedEvent.id), JSON.stringify({
        subject: emailSubject, heading: emailHeading, subheading: massmailSubheading, body: emailBody,
      }));
      setMassmailDraftSaved(true);
      window.setTimeout(() => setMassmailDraftSaved(false), 2500);
    } catch { /* */ }
  };
  // Auto-Speichern, solange der Massenmail-Editor offen ist.
  React.useEffect(() => {
    if (massmailMode !== 'editor' || !showEmailModal || !selectedEvent || massmailHydratingRef.current) return;
    try {
      window.localStorage.setItem(massmailDraftKey(selectedEvent.id), JSON.stringify({
        subject: emailSubject, heading: emailHeading, subheading: massmailSubheading, body: emailBody,
      }));
    } catch { /* */ }
  }, [massmailMode, showEmailModal, selectedEvent, emailSubject, emailHeading, massmailSubheading, emailBody]);
  // v26.78: Event-Foto als Base64 vorladen, sobald der Massenmail-Editor
  // geöffnet wird — für die Live-Vorschau und den Versand (Bild-im-Kopf-Wahl).
  // Wechselt der Nutzer das Event, wird die Wahl auf „Standard" zurückgesetzt.
  React.useEffect(() => {
    if (!showEmailModal || !selectedEvent) return;
    setMassmailHeaderImage(p => ({ ...p, hero: 'logo' }));
    setMassmailEventPhotoB64('');
    const url = selectedEvent.imageUrl;
    if (!url) return;
    let cancelled = false;
    getCachedImage(url)
      .then(b64 => { if (!cancelled && b64 && b64.indexOf('data:') === 0) setMassmailEventPhotoB64(b64); })
      .catch(() => { /* Event-Foto nicht ladbar → Option bleibt ohne Vorschau/deaktiviert */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [showEmailModal, selectedEvent && selectedEvent.id]);
  // v26.88: Event-Foto AUCH für die Einladungsmail vorladen (Bild-im-Kopf-Wahl),
  // sobald das Einladungs-Fenster geöffnet wird. Beim Event-Wechsel zurück auf
  // „DEX-Logo".
  React.useEffect(() => {
    if (!showInviteModal || !selectedEvent) return;
    setInviteHeaderImage(p => ({ ...p, hero: 'logo' }));
    setInviteEventPhotoB64('');
    const url = selectedEvent.imageUrl;
    if (!url) return;
    let cancelled = false;
    getCachedImage(url)
      .then(b64 => { if (!cancelled && b64 && b64.indexOf('data:') === 0) setInviteEventPhotoB64(b64); })
      .catch(() => { /* Event-Foto nicht ladbar → Option bleibt deaktiviert */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [showInviteModal, selectedEvent && selectedEvent.id]);
  // v26.78: Ersetzt den {{ORB_URL}}-Platzhalter im gewickelten Mail-HTML durch
  // das gewählte Kopf-Bild. Bei „Event-Foto" (+ geladenem Foto) wird das
  // Event-Bild als Base64 fest eingebacken; sonst bleibt {{ORB_URL}} erhalten,
  // damit der Flow wie gehabt das Standard-Bild (DEX-Logo/Orb bzw. das
  // konfigurierte Mail-Logo des Events) einsetzt.
  // v30.52: EINE Umsetzung für beide (und die QR-Mail) — s. utils/mailHeaderImage.
  const applyMassmailHero = (wrappedHtml: string): string =>
    applyHeroImage(wrappedHtml, massmailHeaderImage, massmailEventPhotoB64);
  const applyInviteHero = (wrappedHtml: string): string =>
    applyHeroImage(wrappedHtml, inviteHeaderImage, inviteEventPhotoB64);
  // v29.37: Steht im Kopf ein eigenes Bild? Entweder das eingebackene Event-Foto
  // oder — wenn {{ORB_URL}} stehen bleibt — das Mail-Logo des Events, das der
  // Flow einsetzt. Nur dann darf die volle Breite gelten (sonst Orb-Deckel).
  const massmailHasOwnImage = hasOwnHeaderImage(massmailHeaderImage, massmailEventPhotoB64, selectedEvent && selectedEvent.mailImageBase64);
  const inviteHasOwnImage = hasOwnHeaderImage(inviteHeaderImage, inviteEventPhotoB64, selectedEvent && selectedEvent.mailImageBase64);
  const massmailHeaderOpts = mailHeaderOpts(massmailHeaderImage, massmailHasOwnImage);
  const inviteHeaderOpts = mailHeaderOpts(inviteHeaderImage, inviteHasOwnImage);
  // Testmail mit dem aktuellen Stand an die Organizer (zur Kontrolle vor dem
  // echten Massenversand). Geht NICHT an die Teilnehmer.
  const sendMassmailTestToOrganizers = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent) return;
    const orgEmails = (selectedEvent.organizerEmails || []).filter(Boolean);
    const to = orgEmails.length > 0 ? orgEmails.join(';') : (currentUser.email || '');
    if (!to) {
      setMassmailTestMsg(isDe ? 'Keine Organizer-E-Mail hinterlegt — Test nicht möglich.' : 'No organizer email available — test not possible.');
      return;
    }
    if (!emailSubject.trim() || !emailBody.trim()) {
      setMassmailTestMsg(isDe ? 'Bitte Betreff und Text ausfüllen.' : 'Please fill in subject and body.');
      return;
    }
    setMassmailTesting(true);
    setMassmailTestMsg(null);
    try {
      const previewVars: Record<string, string> = { EventTitle: selectedEvent.title, Organizer: (selectedEvent.organizers || []).join(', ') };
      const resolvedSubject = `[TEST] ${replacePlaceholders(emailSubject, previewVars)}`;
      const resolvedHeading = replacePlaceholders(emailHeading, previewVars);
      const resolvedBody = replacePlaceholders(emailBody, previewVars);
      const resolvedSub = massmailSubheading.trim() ? replacePlaceholders(massmailSubheading, previewVars) : `Event ${selectedEvent.title}`;
      const fullBody = applyMassmailHero(wrapTemplate('#86bc25', resolvedHeading, resolvedSub, resolvedBody, undefined, massmailHeaderOpts));
      await eventServiceRef.queueEmail(resolvedSubject, to, 'Organizer (Test)', fullBody, 'Massenmail', selectedEvent.title, selectedEvent.id);
      setMassmailTestMsg(isDe ? `Testmail an die Organizer (${to.split(';').length}) verschickt — bitte Postfach prüfen.` : `Test email sent to the organizers (${to.split(';').length}) — please check the mailbox.`);
    } catch (err) {
      setMassmailTestMsg((isDe ? 'Fehler beim Test-Versand: ' : 'Error during test send: ') + (err instanceof Error ? err.message : String(err)));
    }
    setMassmailTesting(false);
  };


  // v22.6: QR-Versand-Aktionen als benannte Funktionen (vorher inline im Modal) —
  // macht das neue kompakte Querformat-Layout lesbar. Verhalten unverändert.
  // v22.18: pro-Event angepasster QR-Mail-Text — Override-Key 'QRCode' im
  // EmailTemplateOverrides-JSON des Events (übersteht den Wizard-Roundtrip,
  // weil Nicht-Unterstrich-Keys dort erhalten bleiben). QR-Block bleibt fix.
  const getQrMailOverride = (ev: DeloitteEvent | null): QrEmailOverride | undefined => {
    if (!ev) return undefined;
    try {
      const all = JSON.parse(ev.emailTemplateOverrides || '{}');
      const ov = all && all['QRCode'];
      if (ov && (ov.subject || ov.heading || ov.subheading || ov.bodyHtml)) return ov as QrEmailOverride;
    } catch { /* kein Override */ }
    return undefined;
  };
  // Editor öffnen: Felder aus Override (falls vorhanden) oder den Standard-
  // Texten vorbelegen + Beispiel-QR-Block für die Live-Vorschau erzeugen.
  // v29.26: optionales target — vom Hauptevent aus lassen sich die QR-Mails
  // der Sub-Events einzeln gestalten (je Sub-Event ein eigener Override auf
  // dessen Zeile; Versand und Auto-Versand des Sub-Events lesen genau den).
  const openQrMailEditor = async (target?: DeloitteEvent): Promise<void> => {
    const tgt = target || selectedEvent;
    if (!tgt) return;
    setQrEditTarget(target || null);
    const ov = getQrMailOverride(tgt);
    const def = qrEmailDefaults(tgt.emailLanguage || 'EN');
    setQrEditSubject((ov && ov.subject) || def.subject);
    setQrEditHeading((ov && ov.heading) || def.heading);
    setQrEditSubheading((ov && ov.subheading) || def.subheading);
    setQrEditBody((ov && ov.bodyHtml) || def.body);
    // v30.52: gespeichertes Kopf-Bild laden; Event-Foto für die Auswahl
    // nachziehen (leer = „Event-Foto" bleibt deaktiviert).
    setQrHeaderImage(normalizeMailHeaderImage(ov && ov.headerImage));
    setQrBlockLang((ov && ov.blockLang) || '');
    setQrBlockNote((ov && ov.blockNote) || '');
    setQrEventPhotoB64('');
    if (tgt.imageUrl) {
      getCachedImage(tgt.imageUrl)
        .then(b64 => { if (b64 && b64.indexOf('data:') === 0) setQrEventPhotoB64(b64); })
        .catch(() => { /* Foto nicht ladbar → Option bleibt deaktiviert */ });
    }
    // Beispiel-QR (eigene Daten) für die Vorschau — gleicher Aufbau wie im Versand.
    const myName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
    const qrData = `DEX|${tgt.eventNumber}|${currentUser.email}`;
    // v30.33: Beispiel-ID, damit im Editor sichtbar ist, dass die
    // Teilnehmer-ID mitgeschickt wird — beim echten Versand steht dort die
    // tatsaechliche Nummer der Person.
    const qrImageHtml = await buildQrImageHtml(qrData);
    setQrEditSampleImg(qrImageHtml);
    setQrEditSampleBlock(buildQrBlockHtml(qrImageHtml, myName, SAMPLE_QR_ID, ((ov && ov.blockLang) || tgt.emailLanguage || 'EN'), (ov && ov.blockNote) || ''));
    // v22.19: Versand-Modal schließen — der Editor zeigt die Versand-Aktionen
    // in einer eigenen linken Spalte (nebeneinander statt übereinander).
    // Beim Schließen des Editors öffnet das Versand-Modal wieder.
    setQrSendModalOpen(false);
    setQrEditOpen(true);
  };
  const closeQrMailEditor = (): void => {
    setQrEditOpen(false);
    setQrEditTarget(null);
    setQrSendModalOpen(true);
  };
  // Speichern: Override in das EmailTemplateOverrides-JSON des Events mergen
  // (andere Keys + Piggybacks bleiben erhalten). Entspricht alles den
  // Standard-Texten, wird der Key entfernt (= zurück auf Standard).
  const saveQrMailOverride = async (): Promise<void> => {
    // v29.26: schreibt auf das EDITOR-ZIEL — das kann ein vom Hauptevent aus
    // geöffnetes Sub-Event sein (qrEditTarget), sonst das Event selbst.
    const tgt = qrEditTarget || selectedEvent;
    if (!tgt || qrEditSaving) return;
    setQrEditSaving(true);
    try {
      const def = qrEmailDefaults(tgt.emailLanguage || 'EN');
      // v30.52: Das Kopf-Bild zählt mit. Ohne diese Bedingung würde eine
      // geänderte Bildbreite bei sonst unveränderten Texten den Override
      // LÖSCHEN — die Einstellung wäre nach dem Speichern weg.
      const isDefault = qrEditSubject.trim() === def.subject.trim()
        && qrEditHeading.trim() === def.heading.trim()
        && qrEditSubheading.trim() === def.subheading.trim()
        && qrEditBody.trim() === def.body.trim()
        // v30.60: Ohne diese Bedingung würde eine allein geänderte
        // Block-Sprache den Override löschen — die Auswahl wäre nach dem
        // Speichern wieder weg (dieselbe Falle wie beim Kopf-Bild, v30.52).
        && !qrBlockLang
        && !qrBlockNote.trim()
        && isDefaultMailHeaderImage(qrHeaderImage);
      let all: Record<string, unknown> = {};
      try { all = JSON.parse(tgt.emailTemplateOverrides || '{}') || {}; } catch { all = {}; }
      if (isDefault) {
        delete all['QRCode'];
      } else {
        all['QRCode'] = {
          subject: qrEditSubject, heading: qrEditHeading, subheading: qrEditSubheading, bodyHtml: qrEditBody,
          // Nur Auswahl + Zahlen — NIE das Foto selbst (s. QrEmailOverride).
          headerImage: { ...qrHeaderImage },
          ...(qrBlockLang ? { blockLang: qrBlockLang } : {}),
          ...(qrBlockNote.trim() ? { blockNote: qrBlockNote.trim() } : {}),
        };
      }
      const json = JSON.stringify(all);
      const ok = await updateEvent(tgt.id, { 'EmailTemplateOverrides': json });
      if (ok) {
        if (!qrEditTarget || (selectedEvent && qrEditTarget.id === selectedEvent.id)) {
          setSelectedEvent(prev => prev ? { ...prev, emailTemplateOverrides: json } : prev);
        }
        // Lokale Kopie des Ziels nachziehen — sonst vergleicht der Editor
        // gegen den alten Stand und meldet weiter „ungespeichert".
        setQrEditTarget(prev => (prev && prev.id === tgt.id) ? { ...prev, emailTemplateOverrides: json } : prev);
        await refreshEvents();
        showAlert(isDe
          ? (isDefault
            ? 'QR-Mail auf den Standardtext zurückgesetzt.'
            : 'QR-Mail-Text gespeichert — gilt ab jetzt für Vorschau, Versand UND den automatischen QR-Versand bei neuen Anmeldungen.')
          : (isDefault ? 'QR email reset to the default text.' : 'QR email text saved — now used for preview, sending AND the automatic QR send for new registrations.'), { variant: 'success' });
      } else {
        showAlert(isDe ? 'Speichern fehlgeschlagen — vermutlich fehlen Schreibrechte auf der Event-Liste.' : 'Saving failed — you probably lack write permission on the event list.', { variant: 'error' });
      }
    } finally { setQrEditSaving(false); }
  };
  /**
   * v30.33: Unter den QR-Code kommt die Teilnehmer-ID — groß und vorlesbar.
   *
   * Sie ist nicht nur Deko: Seit v30.33 kann das Check-in-Team die ID ins
   * Suchfeld tippen und die Person damit einchecken. Das ist der Weg, der ohne
   * Kamera auskommt — und auf verwalteten Geräten ist die Kamera nicht überall
   * erreichbar (SharePoint-App-WebView, Android-Foto-Picker). Steht die ID
   * nicht in der Mail, kann niemand sie nennen, und der ganze Weg läuft leer.
   *
   * Bewusst monospace und groß: Die Zahl wird am Einlass vorgelesen und
   * abgetippt, nicht gelesen.
   */
  /**
   * v30.36: Erzeugung liegt jetzt in `utils/qrWithMark` — gemeinsam mit dem
   * Auto-Versand im EventContext. Vorher gab es zwei Erzeuger, die
   * auseinanderliefen (unterschiedliche Fehlerkorrektur), und das faellt
   * niemandem auf: Der Code scannt einfach schlechter.
   */
  const buildQrImageHtml = async (qrData: string): Promise<string> => {
    const qrDataUrl = await buildParticipantQrDataUrl(qrData, 300);
    if (!qrDataUrl) {
      return `<p style="font-family:monospace;font-size:1.2rem;background:#f5f5f5;padding:12px;border-radius:8px;text-align:center;">${qrData}</p>`;
    }
    return `<img src="${qrDataUrl}" alt="QR-Code" style="width:300px;max-width:100%;height:auto;" />`;
  };

  /**
   * v30.52: Löst das Kopf-Bild der QR-Mail auf.
   *
   * Nur wenn im Override „Event-Foto" gewählt ist, wird das Bild als Base64
   * geholt und fest eingebacken; sonst bleibt `{{ORB_URL}}` stehen und der
   * Flow setzt wie bisher das Standard-Bild. `getCachedImage` cacht — pro
   * Sitzung fällt der Abruf also einmal an, nicht je Teilnehmer.
   */
  const qrHeroPhotoFor = async (ev: DeloitteEvent, override?: QrEmailOverride): Promise<string> => {
    const hdr = normalizeMailHeaderImage(override && override.headerImage);
    if (hdr.hero !== 'event' || !ev.imageUrl) return '';
    try {
      const b64 = await getCachedImage(ev.imageUrl);
      return (b64 && b64.indexOf('data:') === 0) ? b64 : '';
    } catch { return ''; }
  };

  const qrPreviewAction = async (): Promise<void> => {
    if (!selectedEvent) return;
    setQrPreviewLoading(true);
    try {
      const orgEmail = currentUser.email;
      const orgFullName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || orgEmail;
      const orgFirstName = currentUser.firstName || orgFullName.split(/\s+/)[0] || orgFullName;
      const qrData = `DEX|${selectedEvent.eventNumber}|${orgEmail}`;
      const qrImageHtml = await buildQrImageHtml(qrData);
      const qrOv = getQrMailOverride(selectedEvent);
      const emailData = qrCodeEmail(orgFirstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN', orgFullName, qrOv, SAMPLE_QR_ID, await qrHeroPhotoFor(selectedEvent, qrOv));
      let eventOrb = '';
      try {
        const ov = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
        if (ov && typeof ov._eventLogo === 'string') eventOrb = ov._eventLogo;
      } catch { /* */ }
      const previewBody = emailData.body.replace(/\{\{ORB_URL\}\}/g, eventOrb || getCachedOrbBase64() || '');
      setQrPreviewSubject(emailData.subject);
      setQrPreviewHtml(previewBody);
      setQrPreviewOpen(true);
    } finally { setQrPreviewLoading(false); }
  };
  // v22.19: optionaler liveOverride — der Test-Versand aus dem Mail-Editor
  // nutzt den AKTUELLEN (ggf. ungespeicherten) Editor-Text, damit Test = Vorschau.
  // v29.26: optionales target — Test aus dem Editor eines Sub-Events nutzt
  // dessen Titel/Sprache/Event-Nummer statt der des geöffneten Events.
  const qrTestSendAction = async (liveOverride?: QrEmailOverride, target?: DeloitteEvent): Promise<void> => {
    const ev = target || selectedEvent;
    if (!eventServiceRef || !ev) return;
    setIsSendingQR(true); setQrSendResult(null); setQrSentCount(0);
    try {
      // v24.99: Test-Mail geht an ALLE Organisatoren des Events (vorher nur an
      // den eingeloggten User). Fallback: wenn keine Organisatoren hinterlegt
      // sind, an mich selbst. Jeder bekommt einen QR mit der EIGENEN Adresse.
      const orgEmails = (ev.organizerEmails || []).map(e => (e || '').trim()).filter(Boolean);
      const orgNames = ev.organizers || [];
      const recipients = orgEmails.length > 0
        ? orgEmails.map((em, i) => ({ email: em, rawName: orgNames[i] || em }))
        : [{ email: currentUser.email, rawName: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email }];
      // v30.52: Test = Vorschau — also auch beim Kopf-Bild den AKTUELLEN
      // Editor-Stand nehmen, wenn der Test aus dem Editor kommt.
      const testOverride = liveOverride || getQrMailOverride(ev);
      const testHeroPhoto = await qrHeroPhotoFor(ev, testOverride);
      let sent = 0;
      for (const r of recipients) {
        const raw = (r.rawName || '').trim();
        // Deloitte-Displayname „Nachname, Vorname" → „Vorname Nachname" + Vorname.
        const fullName = raw.indexOf(',') >= 0 ? raw.split(',').reverse().map(s => s.trim()).join(' ') : (raw || r.email);
        const firstName = raw.indexOf(',') >= 0 ? (raw.substring(raw.indexOf(',') + 1).trim().split(/\s+/)[0] || fullName) : (fullName.split(/\s+/)[0] || fullName);
        const qrData = `DEX|${ev.eventNumber}|${r.email}`;
        const qrImageHtml = await buildQrImageHtml(qrData);
        const emailData = qrCodeEmail(firstName, ev.title, qrImageHtml, ev.emailLanguage || 'EN', fullName, testOverride, SAMPLE_QR_ID, testHeroPhoto);
        await eventServiceRef.queueEmail(emailData.subject, r.email, fullName, emailData.body, 'QRCode', ev.title, ev.id);
        sent++; setQrSentCount(sent);
      }
      setQrSendResult(isDe
        ? `Test-Mail an ${sent} Organisator${sent === 1 ? '' : 'en'} verschickt — bitte im Postfach prüfen.`
        : `Test email sent to ${sent} organizer${sent === 1 ? '' : 's'} — please check the mailbox.`);
    } catch (err) {
      setQrSendResult((isDe ? 'Fehler beim Test-Versand: ' : 'Error during test send: ') + (err instanceof Error ? err.message : String(err)));
    }
    setIsSendingQR(false);
  };
  const qrFullSendAction = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent) return;
    const eligible = registrations.filter(r => r.Status === 'Angemeldet');
    if (eligible.length === 0) {
      setQrSendResult(isDe ? 'Alle Teilnehmer haben bereits einen QR-Code — nichts zu senden.' : 'All participants already have a QR code — nothing to send.');
      return;
    }
    if (!(await confirmDialog(isDe ? `QR-Code an ${eligible.length} Teilnehmer ohne Code senden?` : `Send the QR code to ${eligible.length} participants without a code?`, { confirmLabel: isDe ? 'Senden' : 'Send' }))) return;
    setIsSendingQR(true); setQrSendResult(null); setQrSentCount(0);
    let sent = 0; let extCount = 0;
    for (const reg of eligible) {
      const qrData = `DEX|${selectedEvent.eventNumber}|${reg.ParticipantEmail}`;
      const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
      const firstName = reg.Vorname || (reg.ParticipantName || '').trim().split(/\s+/)[0] || name;
      const qrImageHtml = await buildQrImageHtml(qrData);
      const sendOv = getQrMailOverride(selectedEvent);
      const emailData = qrCodeEmail(firstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN', name, sendOv, reg.TeilnehmerID, await qrHeroPhotoFor(selectedEvent, sendOv));
      // v27.11: Member-Firm-Adressen zählen als intern → QR-Mail direkt.
      const isExternal = isExternalEmail(reg.ParticipantEmail);
      if (isExternal) {
        const orgEmails = (selectedEvent.organizerEmails || []).filter(Boolean);
        const orgRecipient = orgEmails.length > 0 ? orgEmails.join(';') : currentUser.email;
        const orgSubject = `[Externer Teilnehmer] QR-Code für ${name} — ${selectedEvent.title}`;
        const qrExternalHint = `<div style="margin:0 0 16px;padding:12px 16px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;line-height:1.55;color:#7a4a00;">`
          + `<strong>QR-Code für externen Teilnehmer.</strong><br>`
          + `Eigentlich für <strong>${reg.ParticipantEmail}</strong> (${name}). Da externe Adressen keinen Mail-Versand bekommen, landet der QR-Code bei dir — drucke ihn aus oder leite die Mail intern an den Empfänger weiter (Datenschutzrichtlinien Deloitte Deutschland beachten).`
          + `</div>`;
        const qrBody = injectIntoEmailContent(emailData.body, qrExternalHint);
        await eventServiceRef.queueEmail(orgSubject, orgRecipient, 'Organizer', qrBody, 'QRCode', selectedEvent.title, selectedEvent.id);
        extCount++;
      } else {
        await eventServiceRef.queueEmail(emailData.subject, reg.ParticipantEmail, name, emailData.body, 'QRCode', selectedEvent.title, selectedEvent.id);
      }
      if (selectedEvent.subsiteUrl) {
        await eventServiceRef.setQRSentStatus(selectedEvent.subsiteUrl, reg.Id);
      }
      sent++; setQrSentCount(sent);
    }
    // v21: Erster Massen-Versand startet die QR-Phase (AutoSendQRCode=true).
    try { await eventServiceRef.updateEvent(parseInt(selectedEvent.id, 10), { AutoSendQRCode: true }); } catch { /* */ }
    const regs = await getAllRegistrations(selectedEvent.id);
    setRegistrations(regs);
    setIsSendingQR(false);
    setQrSendResult(extCount > 0
      ? (isDe
        ? `${sent} QR-Codes verschickt (davon ${extCount} an dich/Organizer umgeleitet — externe Adressen).`
        : `${sent} QR codes sent (${extCount} of them redirected to you/the organizer — external addresses).`)
      : (isDe ? `${sent} QR-Codes verschickt.` : `${sent} QR codes sent.`));
  };

  const saveSelfCheckInWindow = async (): Promise<void> => {
    if (!selectedEvent || sciBusy) return;
    if (sciFrom && sciTo && new Date(sciFrom).getTime() >= new Date(sciTo).getTime()) {
      showAlert(isDe ? '„Bis" muss zeitlich nach „Von" liegen.' : '"Until" must be after "From".');
      return;
    }
    setSciBusy(true);
    try {
      const ok = await updateEvent(selectedEvent.id, {
        'SelfCheckInFrom': sciFrom ? new Date(sciFrom).toISOString() : null,
        'SelfCheckInTo': sciTo ? new Date(sciTo).toISOString() : null,
      });
      setSciSaveMsg(ok
        ? (isDe ? 'Zeitfenster gespeichert.' : 'Time window saved.')
        : (isDe ? 'Speichern fehlgeschlagen — bitte erneut versuchen.' : 'Saving failed — please try again.'));
    } finally { setSciBusy(false); }
  };

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
  // Seit v6.5: getrennte Wartelisten bei B2Run-Split-Kapazitäten (Durchstarter/Funstarter).
  // Die Split-Aktivierung erkennen wir daran, dass beide Kapazitäts-Felder gesetzt und > 0 sind.
  const isSplitCapacity = !!selectedEvent
    && typeof selectedEvent.durchstarterCapacity === 'number'
    && typeof selectedEvent.funstarterCapacity === 'number'
    && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
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
  const waitlistTruePos: Record<number, number> = (() => {
    const map: Record<number, number> = {};
    const rank = (arr: SPRegistration[]): void => {
      arr.slice()
        .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0))
        .forEach((r, idx) => { if (typeof r.Id === 'number') map[r.Id] = idx + 1; });
    };
    const all = registrations.filter(r => r.Status === 'Warteliste');
    if (isSplitCapacity) {
      rank(all.filter(r => r.PreferredStarterType === 'Durchstarter'));
      rank(all.filter(r => r.PreferredStarterType === 'Funstarter'));
      rank(all.filter(r => !r.PreferredStarterType || (r.PreferredStarterType !== 'Durchstarter' && r.PreferredStarterType !== 'Funstarter')));
    } else {
      rank(all);
    }
    return map;
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
    closeEditModal, editError, editForm, editingReg, isDe, isSavingEdit,
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
    massmailAudience, massmailMode, massmailStatuses, registrations, selectedEvent, setMassmailAudience,
    setMassmailMode, setMassmailPasteRaw, setMassmailStatuses, setShowEmailModal,
  };
  const massmailPasteModalProps = {
    massmailMode, massmailPasteRaw, registrations, selectedEvent, setMassmailMode, setMassmailPasteRaw,
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
    dupCancelBusy, dupCancelReg, isDe, performSilentDuplicateDelete, performStandardCancel, selectedEvent,
    setDupCancelBusy, setDupCancelReg, showAlert,
  };
  const overbookDecisionModalProps = {
    obBusy, obKeepVariant, obMailBody, obMailLang, obMailSubject, obRemoveCalendar,
    obWithMail, overbookModal, runOverbookResolution, selectedEvent, setObKeepVariant, setObMailBody,
    setObMailLang, setObMailSubject, setObRemoveCalendar, setObWithMail, setOverbookModal,
  };
  const waitlistPositionModalProps = {
    eventServiceRef, getAllRegistrations, isDe, selectedEvent, setRegistrations, setWlPosBusy,
    setWlPosModal, setWlPosValue, showAlert, wlPosBusy, wlPosModal, wlPosValue,
  };
  const adminAddMemberModalProps = {
    addTeamMember, adminAddCcOrganizer, adminAddLeadRegId, adminAddMemberBusy, adminAddMemberConsent, adminAddMemberDialog,
    adminAddMemberError, adminAddMemberIncludeIntl, adminAddMemberPick, adminAddMemberQuery, adminAddMemberQueryTimer, adminAddMemberResults,
    adminAddMemberSearching, adminAddNewPersonMail, adminAddNotifyOthers, adminAddNotifyScope, adminAddSendMail, adminAddTeamlessPicks,
    assignTeamlessToTeam, currentUser, getAllRegistrations, isDe, notifyExistingTeamMembers, registrations,
    searchUsers, selectedEvent, setAdminAddCcOrganizer, setAdminAddLeadRegId, setAdminAddMemberBusy, setAdminAddMemberConsent,
    setAdminAddMemberDialog, setAdminAddMemberError, setAdminAddMemberIncludeIntl, setAdminAddMemberPick, setAdminAddMemberQuery, setAdminAddMemberResults,
    setAdminAddMemberSearching, setAdminAddNewPersonMail, setAdminAddNotifyOthers, setAdminAddNotifyScope, setAdminAddSendMail, setAdminAddTeamlessPicks,
    setRegistrations, setTeamsToast,
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
    getAllRegistrations, isAdmin, isDe, isLoadingRegs, isMobile, isOrganizerFor,
    leadTransferBusy, leadTransferOpenFor, moveRegToTeam, onTeamDrop, openTeamMailDialog, registrations,
    selectedEvent, setAdminAddCcOrganizer, setAdminAddLeadRegId, setAdminAddMemberConsent, setAdminAddMemberDialog, setAdminAddMemberError,
    setAdminAddMemberPick, setAdminAddMemberQuery, setAdminAddMemberResults, setAdminAddNewPersonMail, setAdminAddNotifyOthers, setAdminAddNotifyScope,
    setAdminAddSendMail, setAdminAddTeamlessPicks, setDragOverTid, setDragRegId, setLeadTransferBusy, setLeadTransferOpenFor,
    setRegistrations, setTeamEditOpenFor, setTeamsCollapsed, setTeamsToast, showAlert, teamEditOpenFor,
    teamsCollapsed, transferTeamLead,
  };
  const teamMailModalProps = {
    getActiveTeams, isDe, selectedEvent, sendTeamMails, setTeamMailBody, setTeamMailInfoByTid,
    setTeamMailOpen, setTeamMailSubject, teamMailBody, teamMailInfoByTid, teamMailOpen, teamMailSending,
    teamMailSubject,
  };
  const participantTableProps = {
    activeRegs, allEvents, attachmentsByReg, availableColumns, colToggleHover, columnOrder,
    computeRoommatePairs, confirmDialog, duplicateEmails, eventServiceRef, getAllRegistrations, getRoommateInfo,
    handleSort, hasRoommateColumn, hiddenColumns, hideColumn, highlightMatch, inactiveAccounts,
    isDe, isSplitCapacity, moveColumn, openEditModal, orgPastLock, parentEventForSelected,
    parentRegsByEmail, performStandardCancel, personalColsCollapsed, query, registrations, selectedEvent,
    setAttachmentsModalReg, setColToggleHover, setDupCancelReg, setParticipantDetail, setPersonalColsCollapsed, setRegistrations,
    setShowColumnPicker, setSplitParticipantsView, showAlert, showColumn, showColumnPicker, showMatches,
    sortIcon, splitParticipantsView, stripLocPrefix,
  };
  const waitlistTablesProps = {
    buildCancellationMail, confirmDialog, currentUser, eventServiceRef, getAllRegistrations, isDe,
    isSplitCapacity, query, selectedEvent, setRegistrations, setWaitlistSortAsc, setWaitlistSortColumn,
    setWlPosModal, setWlPosValue, showAlert, waitlistDurch, waitlistFun, waitlistRegs,
    waitlistSortAsc, waitlistSortColumn, waitlistTruePos, waitlistUnassigned, wlPosBusy,
  };
  const cancelledListProps = {
    cancelledRegs, cancelledSortAsc, cancelledSortColumn, confirmDialog, consolidatedChildren, eventServiceRef,
    getAllRegistrations, hasWaitlistActivity, isAdmin, isConsolidatedMode, isDe, isOrganizerFor,
    registrations, selectedEvent, setCancelledSortAsc, setCancelledSortColumn, setRegistrations, setSubRegReloadTick,
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
    activeRegs, childEventsOf, confirmDialog, consolidatedFiltered, detailCardRef, events,
    evTabHover, handleSelectEvent, isAdmin, isConsolidatedMode, isDe, isImpersonating,
    isLoadingRegs, isMobile, isOrganizerFor, navigate, openTabGroup, registrations,
    reservedDetailHeight, reservedDetailWidth, selectedEvent, setCheckInHubOpen, setCheckInHubStep, setEvTabHover,
    setOpenTabGroup, subEventRegsByEventId, t, toggleDraftStatus, waitlistRegs,
  };
  const nextStepsBoxProps = {
    childEventsOf, isAdmin, isDe, isOrganizerFor, openInviteModal, selectedEvent,
    setVisListOpen, visListOpen,
  };
  const billingStatusStripProps = {
    isAdmin, isDe, isFA, isOrganizerFor, isQRScannerOnlyForSelected, navigate,
    selectedEvent, setBillingPanelOpen,
  };
  const adminActionsCardProps = {
    adminEvents, allEvents, childEventsOf, confirmDialog, copiedDeepLink, copiedEmails,
    detectOverbookResult, eventServiceRef, fixColumnsResult, fixFieldsResult, getAllRegistrations, isAdmin,
    isCheckingDeclines, isDe, isDetectingOverbook, isFixingColumns, isFixingFields, isOrganizerFor,
    isPromoting, isQRScannerOnlyForSelected, isRefreshingProfiles, isReorderingIDs, isRepairingAccess, isRepairingNames,
    isRepairingOrganizers, isRepairingPerms, isResettingCounter, isSendingQR, isSplitCapacity, isSyncingRegistry,
    navigate, openChangeLogForEvent, openCommsModal, openInviteModal, openMassmailPicker, promoteResult,
    qrSentCount, refreshEvents, refreshProfilesResult, registrations, reorderResult, repairAccessResult,
    repairNamesResult, repairOrganizersResult, repairPermsResult, resetCounterResult, runIdReorder, runManualPromote,
    searchUsers, selectedEvent, setAccessFixModal, setB2runTodoOpen, setBibImportOpen, setBillingPanelOpen,
    setCheckInHubOpen, setCheckInHubStep, setCopiedDeepLink, setCopiedEmails, setDeclineCopied, setDeclineResult,
    setDetectOverbookResult, setExcelAudience, setExcelTargetModal, setFixColumnsResult, setFixFieldsResult, setIsCheckingDeclines,
    setIsDetectingOverbook, setIsFixingColumns, setIsFixingFields, setIsRefreshingProfiles, setIsRepairingAccess, setIsRepairingNames,
    setIsRepairingOrganizers, setIsRepairingPerms, setIsResettingCounter, setIsSyncingRegistry, setNameFixModal, setRefreshProfilesResult,
    setRegistrations, setRepairAccessResult, setRepairNamesResult, setRepairOrganizersResult, setRepairPermsResult, setResetCounterResult,
    setShirtSizeOpen, setShowDeclineModal, setShowExportMenu, setSubRegReloadTick, setSyncRegistryResult, shirtFieldExists,
    showAlert, showExportMenu, siteUrl, spServiceRef, syncRegistryResult, t,
    updateEvent,
  };
  const kpiTilesProps = {
    isConsolidatedMode, isSplitCapacity, registrations, selectedEvent, subEventRegsByEventId, t,
  };
  const hotelPlanningSectionProps = {
    childEventsOf, confirmDialog, getAllRegistrations, hotelPanelOpen, isAdmin, isDe,
    isOrganizerFor, refreshEvents, registrations, selectedEvent, setHotelPanelOpen, setRegistrations,
    showAlert, subEventRegsByEventId,
  };
  const quizStatsSectionProps = {
    registrations, selectedEvent,
  };
  const activeEventHintsBoxProps = {
    childEventsOf, expandedHintIds, hintLangBusy, hintsDismissTick, isAdmin, isDe,
    isOrganizerFor, parentEventForSelected, refreshEvents, selectedEvent, setExpandedHintIds, setHintLangBusy,
    setHintsDismissTick, setQrSendModalOpen, setSelectedEvent, showAlert, updateEvent,
  };
  const audienceVisibilityRowProps = {
    isAdmin, isDe, isOrganizerFor, openPendingReminder, orgPastLock, pendingCheckBusy,
    resolveAudienceEmails, selectedEvent, setVisibilityAllAddresses, setVisibilityBusy, setVisibilityOpen, setVisibilityResolved,
    visibilityAllAddresses, visibilityBusy, visibilityOpen, visibilityResolved,
  };
  const consolidatedViewProps = {
    addAllToKlammer, addingToKlammer, addToKlammer, bulkKlammerProgress, colToggleHover, confirmDialog,
    consolidatedChildren, consolidatedFiltered, consolidatedRows, consolidatedSort, consolidatedSortAsc, expandedConsolidatedEmail,
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
              void (async () => {
                try {
                  const regs = await getAllRegistrations(selectedEvent.id);
                  setRegistrations(regs);
                } catch { /* Anzeige aktualisiert sich beim nächsten Öffnen */ }
              })();
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
        {deniedSubEventLists.length > 0 && (
          <div style={{
            border: '1px solid var(--dex-red)', background: '#fff5f5', borderRadius: 8,
            padding: '12px 14px', marginBottom: 12, fontSize: 13, lineHeight: 1.5,
          }}>
            <strong style={{ color: 'var(--dex-red)' }}>
              {isDe
                ? `Kein Zugriff auf ${deniedSubEventLists.length} Teilnehmerliste(n)`
                : `No access to ${deniedSubEventLists.length} participant list(s)`}
            </strong>
            <div style={{ marginTop: 6 }}>
              {isDe
                ? 'Die Zahlen unten sind deshalb unvollständig — betroffene Termine erscheinen mit 0 Teilnehmern, obwohl dort Anmeldungen liegen können. Grund ist fast immer, dass du erst nachträglich als Organizer benannt wurdest: Die Berechtigung wurde dann nur auf dem Haupt-Event gesetzt, nicht auf den einzelnen Terminen. Ein Admin oder der Haupt-Organizer behebt das über die Aktion „Organizer-Berechtigungen reparieren“.'
                : 'The numbers below are therefore incomplete — affected dates show 0 participants even though registrations may exist. This almost always happens when you were named organizer after the event was created: permissions were then set on the main event only, not on the individual dates. An admin or the main organizer can fix this via the action „Repair organizer permissions“.'}
            </div>
            <div style={{ marginTop: 6, color: 'var(--dex-gray-500)' }}>
              {deniedSubEventLists.slice(0, 8).join(' · ')}
              {deniedSubEventLists.length > 8 ? ` … (+${deniedSubEventLists.length - 8})` : ''}
            </div>
          </div>
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
