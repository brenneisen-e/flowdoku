/**
 * Event-Erstellung (nur für Organizer / SuperAdmin)
 *
 * Erstellt ein Event in der DEX_Events-Liste und eine
 * separate Teilnehmerliste mit Item-Level Permissions.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents, formatOrganizerList } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
// v20.4: moderne Confirm-/Alert-Modals statt window.confirm/alert.
import { useDialog } from '../context/DialogContext';
import { EventService, CustomField } from '../services/EventService';
import { isThrottled } from '../utils/spThrottle';
// v26.48: zentrale B2Run-Köln-Vorlage (Titel-Erkennung + 7 Meldefelder mit
// deterministischen IDs für den offiziellen Excel-Export).
import { isB2RunKoelnTitle, b2runKoelnTemplateFields } from '../data/b2runKoeln';
import { eventCreatedEmail, buildOutlookBody, stripOutlookWrapper, parseOutlookHeadings, replacePlaceholders, getCachedOrbBase64, normalizeMadeWithLink } from '../services/EmailTemplates';
import { exportSummaryAsPdf, exportSummaryAsDoc, SummaryData } from '../services/EventSummaryExport';
import { EventType, AgendaItem } from '../types';
import { Trash2, Send, Plus, X, Users, Check } from './Icons';
import { RichText } from '@pnp/spfx-controls-react/lib/controls/richText';
import { HtmlEditorModal } from './HtmlEditorModal';
import { RegisterPreviewModal } from './RegisterPreviewModal';
import { InfoTooltip } from './InfoTooltip';
import WizardHint from './WizardHint';
import BulkUserImportModal from './BulkUserImportModal';
import AudiencePicker from './AudiencePicker';
import ImageCropModal from './ImageCropModal';
import Modal from './Modal';
import InternationalSearchToggle from './InternationalSearchToggle';
import OrganizerList from './OrganizerList';
// v20.2: Self-Check-in ist aus dem Wizard ausgezogen — Aktivierung läuft
// automatisch beim ersten Klick auf die Aktionen (Check-in-Seite, Admin
// Center, QR-Kachel im Event-Detail); Zeitfenster + Deaktivieren im
// Kachel-Modal des Admin Centers.
import { buildOutlookLocation } from '../utils/eventFormat';
import { setActiveWizardStep } from '../utils/wizardStepContext';
import { canEditBilling } from '../data/billingFields';
import { BundledComm, bundledCommOf, bundledCommConfig } from '../utils/bundledComm';
// v28.98: Sperrt „Zurück" im Header, solange gespeichert wird.
import { setSaveInProgress } from '../utils/saveGuard';
import { shortSubEventTitle } from '../utils/subEventTitle';
import { DESCRIPTION_TEMPLATES } from '../data/descriptionTemplates';
import { CustomFieldInput } from './wizard/customFieldInput';
import { BillingStep } from './wizard/steps/BillingStep';
import { DocumentsStep } from './wizard/steps/DocumentsStep';
import { FunZoneStep } from './wizard/steps/FunZoneStep';
import { TeamStep } from './wizard/steps/TeamStep';
// v28.94: Unterkomponenten des Assistenten liegen jetzt in ./wizard —
// sie kennen den Wizard-State nicht und liessen sich deshalb ohne
// Verhaltensaenderung herausloesen.
import { StickyTabStrip } from './wizard/StickyTabStrip';
import { StepBadge } from './wizard/StepBadge';
import { LocationMultiSelect } from './wizard/LocationMultiSelect';
import { FieldDescEditor } from './wizard/FieldDescEditor';
import { FieldTypeSuggestion } from './wizard/FieldTypeSuggestion';
import { useIsMobile } from '../utils/useIsMobile';
import { Icon } from '@fluentui/react/lib/Icon';
import DatePicker, { registerLocale } from 'react-datepicker';
import { de } from 'date-fns/locale';
import 'react-datepicker/dist/react-datepicker.css';
import { dlog } from '../utils/debugLog';
import { ImgView, SubEventDraft, OutlookConfirmItem, SuggestedCategory, SuggestedEntry } from './wizard/wizardTypes';
import { EmailOverrideEntry } from './wizard/emailOverrideEntry';
import { compressImage } from '../utils/imageCompress';
import { BasicsStep } from './wizard/steps/BasicsStep';
import { DetailsStep } from './wizard/steps/DetailsStep';
import { LocationProgramStep } from './wizard/steps/LocationProgramStep';
import { SubEventsSection } from './wizard/steps/SubEventsSection';
import { CapacityStep } from './wizard/steps/CapacityStep';
import { FieldsStep } from './wizard/steps/FieldsStep';
import { CommunicationStep } from './wizard/steps/CommunicationStep';
import { outlookLogoPiggyback, readOutlookLogo, serializeCustomFields, reinsertOrganizerPlaceholder, resolveAudienceMembersToCsv } from './wizard/wizardHelpers';
import { detectOutlookRelevantChangesImpl } from './wizard/logic/outlookChanges';
import { runWizardSubmit } from './wizard/logic/wizardSubmit';
import { persistSubEventsForParentImpl } from './wizard/logic/persistSubEvents';
import { WizardTermsModal } from './wizard/WizardTermsModal';
import { WizardModals } from './wizard/WizardModals';

// Deutsche Locale registrieren
registerLocale('de', de);

// v22.36: Die kuratierten Agenda-Icon-Listen (AGENDA_ICONS/EXTENDED_ICONS)
// sind entfallen — Agenda-Schritte sind durchnummeriert statt bebildert.

/**
 * Komprimiert ein Bild clientseitig via Canvas.
 * Max 1200px Breite, JPEG 80% Qualität.
 */














export default function EventCreationPage(): React.ReactElement {
  const { goBack, selectedEventId, currentPage, setNavigationGuard, navigate } = useNavigation();
  const { events, childEventsOf, ensureEventDocuments, createEvent, updateEvent, getLastEventUpdateError, deleteEvent, deleteEventItemOnly, refreshEvents, requestCoOrganizerApprovals, notifyNewCoOrganizers, notifyAdminsExternalAudienceAccess } = useEvents();
  const { currentUser } = useCurrentUser();
  // searchGroups + searchUsersByLocation werden seit v19.x ausschließlich im
  // ausgelagerten <AudiencePicker> verwendet (eigener useRoles-Hook dort).
  const { searchUsers, getGroupMembers, canCreateEvents, isAdmin, originalIsAdmin, isFA } = useRoles();
  // v29.66: F&A-Pilot — der komplette Abrechnungs-Teil ist bewusst NUR fuer
  // Admins sichtbar (Testphase laut Fachkonzept; vor dem Rollout sind noch
  // Abstimmungsschleifen geplant). originalIsAdmin deckt den Demo-Modus ab.
  const adminLike = isAdmin || originalIsAdmin;
  // v26.34: Der „Benötigst du Hilfe?"-Ball (Power-User-Hilfe) unten rechts auf
  // Wizard-Seite 1 wurde auf Wunsch entfernt (inkl. powerUsers-Memo + State).
  // v13.0: Frühe Permission-Prüfung — vorher konnte ein Demo-User die
  // Seite öffnen und das Save würde erst beim SP-Write scheitern. Mit
  // Guard zurück zur Start-Seite, falls keine Organizer-Rechte.
  React.useEffect(() => {
    if (!canCreateEvents) goBack();
  }, [canCreateEvents, goBack]);
  // v19.x: Audience-Such-/Chip-/Member-Modal-State ist nach <AudiencePicker>
  // gewandert (Hauptevent + jedes Sub-Event halten dort ihren eigenen State).
  // Hier bleibt nur die persistierte Audience selbst (siehe `audience` weiter
  // unten) und die Ausschluss-Liste (`excludedUsers`).

  // Nutzungsbedingungen: Beim Erstellen eines neuen Events muss der Organizer
  // zuerst eine Bestätigungs-Maske mit den Nutzungs- und Datenschutz-
  // bedingungen akzeptieren. Nicht relevant beim Bearbeiten bestehender Events.
  const [tcAccepted, setTcAccepted] = React.useState(false);
  const [tcCheckbox, setTcCheckbox] = React.useState(false);
  // v28.41: Zweite, bewusst getrennte Bestätigung — der Organizer muss aktiv
  // erklären, dass es ein internes Event ist bzw. die Deloitte-Teilnahme an
  // einer externen Veranstaltung koordiniert. Absichtlich NICHT mit der
  // Nutzungsbedingungs-Zustimmung zusammengelegt: Wer beides in einem Haken
  // abnickt, hat den Einsatzbereich nicht gelesen.
  const [internalCheckbox, setInternalCheckbox] = React.useState(false);
  const [tcExpanded, setTcExpanded] = React.useState(false);
  const { t, locale } = useLanguage();
  // v20.4: App-Modals statt nativer Browser-Dialoge.
  const { confirmDialog, showAlert } = useDialog();
  // Mobile-Breakpoint (≤768px) für responsive Grid-/Flex-Anpassungen im Wizard.
  const isMobile = useIsMobile();
  const isDe = locale === 'de';

  // Edit-Modus: wenn wir auf 'edit-event' sind und eine selectedEventId haben
  const isEditMode = currentPage === 'edit-event' && !!selectedEventId;
  const editEvent = isEditMode ? events.find(e => e.id === selectedEventId) : null;
  // v30.46: Wer darf den Abrechnungs-Schritt sehen? EINE Ableitung für alle
  // Stellen — Schritt-Array, Rendering, Speichern-Dialog und die Obergrenze des
  // Sprung-Index hängen ab jetzt hier dran, nicht mehr je einzeln an
  // `adminLike`. Der Schalter dahinter (`FA_BILLING_STEP_FOR_ORGANIZERS` in
  // `data/billingFields.ts`) steht im Pilot auf `false`; ihn auf `true` zu
  // setzen öffnet den Schritt für Organizer des eigenen Events, ohne dass hier
  // etwas angefasst werden muss.
  //
  // Steht bewusst HINTER `editEvent`: Weiter oben (neben `adminLike`) wäre es
  // ein TDZ-Fehler — dieselbe Falle, die v29.71 an genau dieser Stelle schon
  // einmal erwischt hat.
  const canBilling = canEditBilling(
    adminLike,
    (editEvent?.organizerEmails || []).some(e => (e || '').toLowerCase() === (currentUser.email || '').toLowerCase()),
    isFA
  );

  // v29.71 BUG-FIX: Dieser Block stand VOR der editEvent-Deklaration (Zeile
  // oben). Die useState-Initialisierer laufen synchron beim ersten Render —
  // der Zugriff auf editEvent warf dort einen ReferenceError (TDZ), den das
  // try/catch STILL schluckte: alle Abrechnungs- und Freischalt-States
  // starteten immer leer. Beim Wieder-Oeffnen war der Haken weg, und der
  // naechste Save schrieb den leeren Zustand zurueck und LOESCHTE damit die
  // zuvor gespeicherte Einstellung. Ein catch um einen Programmierfehler ist
  // kein Netz, sondern ein Versteck — deshalb steht der Block jetzt hier.
  // v29.66: Abrechnungsrelevanz (F&A-Pilot, nur Admins). Persistiert als
  // Piggyback `_billing` in EmailTemplateOverrides — beim Laden gestrippt
  // (s. Strip-Block), beim Speichern an BEIDEN Pfaden gemergt.
  // `relevant` kennt drei Zustaende: null = nie beantwortet (keine
  // Vorauswahl, das verlangt das Konzept ausdruecklich), true/false =
  // aktive Entscheidung des Organizers.
  const [billingRelevant, setBillingRelevant] = React.useState<boolean | null>(() => {
    try {
      const b = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._billing;
      if (b && b.relevant === true) return true;
      if (b && b.relevant === false) return false;
    } catch { /* kein Blob, kein Zustand */ }
    return null;
  });
  // Versandart: Konzept-Default ist Option 2 (manuell).
  const [billingSendMode, setBillingSendMode] = React.useState<'auto' | 'manual'>(() => {
    try {
      const b = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._billing;
      if (b && b.sendMode === 'auto') return 'auto';
    } catch { /* */ }
    return 'manual';
  });
  const [billingFields, setBillingFields] = React.useState<Record<string, string>>(() => {
    try {
      const b = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._billing;
      if (b && b.fields && typeof b.fields === 'object') return b.fields as Record<string, string>;
    } catch { /* */ }
    return {};
  });
  // Der Frage-Dialog nach den Nutzungsbedingungen (nur neue Events, nur Admins).
  const [billingPromptOpen, setBillingPromptOpen] = React.useState(false);
  // v29.67: Freischalt-Regel fuer Kalender-Termine. Bei sehr vielen Tagen
  // sollen Teilnehmer nicht Monate im Voraus buchen — die Tage oeffnen erst
  // X Tage vorher (je Termin) oder X Tage vor dem Montag der jeweiligen
  // Woche (dann oeffnet die ganze KW gemeinsam).
  const [openRuleEnabled, setOpenRuleEnabled] = React.useState<boolean>(() => {
    try { const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subEventOpenRule; return !!(r && r.days > 0); } catch { return false; }
  });
  const [openRuleMode, setOpenRuleMode] = React.useState<'day' | 'week'>(() => {
    try { const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subEventOpenRule; return r && r.mode === 'week' ? 'week' : 'day'; } catch { return 'day'; }
  });
  const [openRuleDays, setOpenRuleDays] = React.useState<number>(() => {
    try { const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subEventOpenRule; return (r && typeof r.days === 'number' && r.days > 0) ? r.days : 7; } catch { return 7; }
  });
  // v29.76: „Anmeldung ab" kann statt rollierend auch ein FESTES Datum sein
  // (alle Termine oeffnen gemeinsam an diesem Tag). Gleiches Piggyback,
  // eigener mode 'fixed' — die Anmeldeseite wertet beide Formen aus.
  const [openRuleFixedDate, setOpenRuleFixedDate] = React.useState<string>(() => {
    try {
      const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subEventOpenRule;
      return (r && r.mode === 'fixed' && r.date) ? isoToLocal(r.date) : '';
    } catch { return ''; }
  });
  const subEventOpenRulePiggyback = (): Record<string, unknown> => {
    if (!subEventsOptIn) return {};
    if (openRuleEnabled && openRuleDays > 0) return { _subEventOpenRule: { mode: openRuleMode, days: openRuleDays } };
    if (!openRuleEnabled && openRuleFixedDate) return { _subEventOpenRule: { mode: 'fixed', date: berlinLocalToUtcIso(openRuleFixedDate) || '' } };
    return {};
  };
  // v29.76: Rollierende Fristen fuer Kalender-Termine — „Anmeldung bis" und
  // „Abmeldung bis" als Abstand zum jeweiligen Termin statt festem Datum
  // (Gegenstueck zu „Anmeldung ab"). Die Regel wird MATERIALISIERT: ein
  // Effect rechnet je Sub-Event startDate minus Abstand und schreibt das
  // Ergebnis in registrationDeadline/lastDeregisterDate des Drafts —
  // Anmeldeseite, Organizer Center und Flows lesen weiter die echten
  // Spalten, es gibt keinen zweiten Auswertungsort. Das Piggyback traegt
  // nur die Regel selbst: fuer die UI beim Wieder-Oeffnen, und damit neu
  // angeklickte Kalender-Tage sie automatisch bekommen.
  const [regRuleEnabled, setRegRuleEnabled] = React.useState<boolean>(() => {
    try { const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subDeadlineRule; return !!(r && r.reg && r.reg.amount > 0); } catch { return false; }
  });
  const [regRuleAmount, setRegRuleAmount] = React.useState<number>(() => {
    try { const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subDeadlineRule; return (r && r.reg && typeof r.reg.amount === 'number' && r.reg.amount > 0) ? r.reg.amount : 1; } catch { return 1; }
  });
  const [regRuleUnit, setRegRuleUnit] = React.useState<'days' | 'hours'>(() => {
    try { const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subDeadlineRule; return r && r.reg && r.reg.unit === 'hours' ? 'hours' : 'days'; } catch { return 'days'; }
  });
  const [cancelRuleEnabled, setCancelRuleEnabled] = React.useState<boolean>(() => {
    try { const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subDeadlineRule; return !!(r && r.cancel && r.cancel.amount > 0); } catch { return false; }
  });
  const [cancelRuleAmount, setCancelRuleAmount] = React.useState<number>(() => {
    try { const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subDeadlineRule; return (r && r.cancel && typeof r.cancel.amount === 'number' && r.cancel.amount > 0) ? r.cancel.amount : 1; } catch { return 1; }
  });
  const [cancelRuleUnit, setCancelRuleUnit] = React.useState<'days' | 'hours'>(() => {
    try { const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subDeadlineRule; return r && r.cancel && r.cancel.unit === 'hours' ? 'hours' : 'days'; } catch { return 'days'; }
  });
  // v29.77: Abmelden darf auch NACH dem Termin-Beginn noch erlaubt sein
  // („bis 1 Stunde nach dem Termin") — Richtung als eigener Schalter.
  const [cancelRuleAfter, setCancelRuleAfter] = React.useState<boolean>(() => {
    try { const r = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._subDeadlineRule; return !!(r && r.cancel && r.cancel.after === true); } catch { return false; }
  });
  const subDeadlineRulePiggyback = (): Record<string, unknown> => {
    if (!subEventsOptIn) return {};
    const rule: Record<string, unknown> = {};
    if (regRuleEnabled && regRuleAmount > 0) rule.reg = { amount: regRuleAmount, unit: regRuleUnit };
    if (cancelRuleEnabled && cancelRuleAmount > 0 && userCancelAllowed) rule.cancel = { amount: cancelRuleAmount, unit: cancelRuleUnit, ...(cancelRuleAfter ? { after: true } : {}) };
    return Object.keys(rule).length ? { _subDeadlineRule: rule } : {};
  };
  // Abstand relativ zum Termin-Start — bewusst exakt (X Tage = X * 24 h),
  // damit die Regel bei Ganztags-Terminen (Start 00:00) glatte Tagesgrenzen
  // liefert. after=true rechnet VORWAERTS (Abmelden nach Beginn).
  const rollingDeadlineIso = (startIso: string, amount: number, unit: 'days' | 'hours', after?: boolean): string => {
    const t2 = new Date(startIso || '').getTime();
    if (!isFinite(t2) || !(amount > 0)) return '';
    return new Date(t2 + (after ? 1 : -1) * amount * (unit === 'hours' ? 3600000 : 86400000)).toISOString();
  };
  // v29.75: „Sichtbarkeit gilt für alle Sub-Events" — der Haken hält
  // Standortfilter, Verteiler und Verknüpfung der Sub-Events mit der
  // Klammer synchron (Spiegel-Effect weiter unten, nach den States).
  // Persistiert als Piggyback, damit der Haken beim Wieder-Öffnen
  // gesetzt bleibt und künftige Klammer-Änderungen weiter durchreichen.
  const [visAllSubs, setVisAllSubs] = React.useState<boolean>(() => {
    try { return JSON.parse(editEvent?.emailTemplateOverrides || '{}')._visAllSubs === true; } catch { return false; }
  });
  // v30.7: Der Haken ging beim Speichern verloren (User-Meldung: gesetzt,
  // gespeichert, beim Wieder-Öffnen weg). Zwei Absicherungen: (1) Das
  // fruehere `&& subEventsOptIn`-Gate entfaellt — ein gesetzter Haken wird
  // IMMER geschrieben (ohne Sub-Events wirkt er schlicht nicht). (2) Solange
  // der Organizer den Haken in DIESER Sitzung nicht angefasst hat, wird der
  // Server-Stand weitergetragen — selbst wenn der State-Init ihn (z.B. durch
  // einen unter Drosselung veralteten editEvent) nicht mitbekommen hat,
  // loescht ein argloser Save ihn dann nicht mehr.
  const visAllSubsTouchedRef = React.useRef(false);
  const visAllSubsPiggyback = (): Record<string, unknown> => {
    if (visAllSubs) return { _visAllSubs: true };
    if (!visAllSubsTouchedRef.current) {
      try {
        if (JSON.parse(editEvent?.emailTemplateOverrides || '{}')._visAllSubs === true) return { _visAllSubs: true };
      } catch { /* */ }
    }
    return {};
  };
  // v30.5: Alles AUSSER relevant/sendMode/fields (Versand-Historie, Stempel,
  // Snapshots, Abschluss durch F&A) pflegen die F&A-Flows über
  // patchEventOverridesValue — der Wizard darf diese Schlüssel beim
  // Speichern nicht verlieren. Beim Öffnen einfrieren, beim Bauen des
  // Piggybacks wieder unterlegen.
  const billingExtraRef = React.useRef<Record<string, unknown>>((() => {
    try {
      const b = JSON.parse(editEvent?.emailTemplateOverrides || '{}')._billing;
      if (b && typeof b === 'object') {
        const { relevant, sendMode, fields, ...extra } = b as Record<string, unknown>;
        void relevant; void sendMode; void fields;
        return extra;
      }
    } catch { /* */ }
    return {};
  })());
  // v30.5: Protokollierung (Fachkonzept Abschnitt 13) — geloggt wird beim
  // Speichern, wenn sich Kennzeichnung oder Angaben geändert haben.
  const billingInitialRef = React.useRef<string>(JSON.stringify({ r: billingRelevant, m: billingSendMode, f: billingFields }));
  const billingPiggyback = (): Record<string, unknown> => {
    if (billingRelevant === null) return {};
    const snap = JSON.stringify({ r: billingRelevant, m: billingSendMode, f: billingFields });
    if (snap !== billingInitialRef.current) {
      try {
        const prev = JSON.parse(billingInitialRef.current) as { r: boolean | null };
        const by = `${currentUser?.firstName || ''} ${currentUser?.surname || ''}`.trim() || currentUser?.email || '';
        const log = Array.isArray(billingExtraRef.current.log) ? (billingExtraRef.current.log as unknown[]) : [];
        const action = prev.r !== billingRelevant
          ? (billingRelevant ? 'Event als abrechnungsrelevant markiert' : 'Event nicht mehr als abrechnungsrelevant markiert')
          : 'Abrechnungsinformationen geändert';
        billingExtraRef.current = { ...billingExtraRef.current, log: [...log, { ts: new Date().toISOString(), by, action }].slice(-60) };
      } catch { /* Log ist best-effort */ }
      billingInitialRef.current = snap;
    }
    return { _billing: { relevant: billingRelevant, sendMode: billingSendMode, fields: billingFields, ...billingExtraRef.current } };
  };
  // Alle elf Felder sind laut Fachkonzept Pflicht. Unvollstaendig ist ein
  // STATUS („Abrechnungsrelevante Informationen unvollständig"), kein
  // Speicher-Blocker — deshalb bewusst KEIN getStepErrors-Fall fuer Schritt 10.
  // v30.5: Definition nach data/billingFields.ts gezogen (Wizard, F&A Center
  // und F&A-Mails brauchen dieselbe Liste).

  // ========== Zeitzonen-Handling (Europe/Berlin, browser-TZ-unabhängig) ==========
  //
  // Hintergrund: Der datetime-local-Input liefert einen naiven String ohne TZ-Suffix
  // (z.B. "2026-04-23T19:00"). Wenn wir diesen mit `new Date(str).toISOString()`
  // konvertieren, interpretiert JavaScript den String in der BROWSER-Zeitzone. Bei
  // einem Browser auf UTC oder in einer VM/Citrix mit falscher TZ führt das zu einem
  // 2h-Shift: 19:00 wird als UTC interpretiert statt als MESZ, SP speichert 19:00Z
  // statt 17:00Z, und Outlook zeigt dann 21:00 MESZ.
  //
  // Fix: Die App interpretiert ALLE Event-Zeiten explizit als Europe/Berlin, egal
  // welche Zeitzone der Browser hat. Wir nutzen Intl.DateTimeFormat um den Offset
  // für einen konkreten Zeitpunkt zu bestimmen (DST-aware).

  /** Gibt den Offset von Europe/Berlin zu UTC an dem gegebenen Zeitpunkt in ms zurück.
   *  Im Winter: +3600000 (+1h). Im Sommer: +7200000 (+2h). */
  const berlinOffsetMs = (dateUtc: Date): number => {
    const dtf = new Intl.DateTimeFormat('en-US', {
      timeZone: 'Europe/Berlin',
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
    });
    const parts = dtf.formatToParts(dateUtc);
    const get = (type: string): number => parseInt(parts.find(p => p.type === type)?.value || '0', 10);
    let h = get('hour');
    if (h === 24) h = 0; // en-US hour12:false liefert manchmal 24 statt 0
    const asIfUtc = Date.UTC(get('year'), get('month') - 1, get('day'), h, get('minute'), get('second'));
    return asIfUtc - dateUtc.getTime();
  };

  /** datetime-local-String ("2026-04-23T19:00") als Europe/Berlin interpretieren
   *  und nach UTC-ISO konvertieren ("2026-04-23T17:00:00.000Z"). */
  const berlinLocalToUtcIso = (localStr: string): string => {
    if (!localStr) return '';
    // Parse den String erstmal als ob er UTC wäre -> das sind UTC-Zahlen die den Berlin-Werten entsprechen
    const asUtc = new Date(localStr.length === 16 ? localStr + ':00Z' : localStr + 'Z');
    if (isNaN(asUtc.getTime())) return '';
    // Der echte UTC-Zeitpunkt ist asUtc minus Berlin-Offset an diesem Zeitpunkt
    const offset = berlinOffsetMs(asUtc);
    return new Date(asUtc.getTime() - offset).toISOString();
  };

  /** UTC-ISO ("2026-04-23T17:00:00.000Z") nach datetime-local in Europe/Berlin
   *  ("2026-04-23T19:00") konvertieren — für das Input-Feld. */
  const isoToLocal = (iso: string): string => {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    const dtf = new Intl.DateTimeFormat('en-US', {
      timeZone: 'Europe/Berlin',
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hour12: false,
    });
    const parts = dtf.formatToParts(d);
    const get = (type: string): string => parts.find(p => p.type === type)?.value || '00';
    let hour = get('hour');
    if (hour === '24') hour = '00';
    return `${get('year')}-${get('month')}-${get('day')}T${hour}:${get('minute')}`;
  };

  // Deadline-Datum als Ende-des-Tages (23:59 Europe/Berlin) speichern, damit:
  //  a) Die Uhrzeit-Anzeige in der EventCard nicht mehr "02:00" zeigt
  //  b) Die Deadline-Prüfung "new Date(deadline) < new Date()" wirklich den
  //     gesamten ausgewählten Tag als gültig behandelt (statt nur bis UTC-Mitternacht).
  const deadlineToEndOfDayIso = (dateStr: string): string | null => {
    if (!dateStr) return null;
    // dateStr im Format "YYYY-MM-DD" (date-Input) - wir behandeln als 23:59 Europe/Berlin
    const localStr = dateStr.length === 10 ? `${dateStr}T23:59` : dateStr;
    const utcIso = berlinLocalToUtcIso(localStr);
    return utcIso || null;
  };

  const [title, setTitle] = React.useState(editEvent ? editEvent.title : '');
  // Mehrere Organizer werden mit '; ' getrennt gespeichert (innerhalb eines Namens
  // kann ',' vorkommen, z.B. 'Maerzluft, Petra').
  //
  // Auto-Heal bei Längen-Mismatch (Legacy-Korruption aus v10.0–v10.2-Closure-Bug):
  // pad auf max(names.length, emails.length) statt truncate auf min. Dadurch
  // verliert der User keine Organizer-Einträge beim Edit-Load — fehlende Emails
  // bleiben als leere Strings erhalten und können vom User einzeln per Picker
  // nachgepflegt werden. Die Warning-Box (siehe weiter unten in der UI) macht das
  // Mismatch sichtbar.
  const [organizer, setOrganizer] = React.useState(() => {
    if (!editEvent) return `${currentUser.firstName} ${currentUser.surname}`;
    const names = editEvent.organizers || [];
    const emails = editEvent.organizerEmails || [];
    const max = Math.max(names.length, emails.length);
    const padded: string[] = [];
    for (let i = 0; i < max; i++) padded.push(names[i] || '');
    return padded.join('; ');
  });
  const [organizerResults, setOrganizerResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const [organizerSearch, setOrganizerSearch] = React.useState('');
  // Beim Edit: organizerEmails aus dem gespeicherten Event übernehmen, nicht auf currentUser
  // zurücksetzen. Sonst überschreibt ein Edit+Save die gesamte Organizer-Email-Liste mit
  // nur der Mail des aktuellen Editors — alle anderen Organizer würden stumm aus der
  // Late-Cancel- / Organizer-Mail-Verteilung rausfallen.
  //
  // Auto-Heal: wenn organizers (Names) und organizerEmails unterschiedliche Längen haben
  // (Symptom des Closure-Bugs aus v10.0–v10.2: in dem Fenster wurden Emails per prev =>
  // korrekt akkumuliert, Names aber nur einmal geschrieben → mehr Emails als Namen
  // gespeichert), wird das längere Array auf die Länge des kürzeren abgeschnitten. Sonst
  // produziert die Index-basierte Render-Logik Phantom-Chips ohne Namen oder zeigt Fotos
  // zur falschen Email — und der Picker grayt Personen aus, die sichtbar gar nicht in
  // der Liste sind, weil ihre Email noch im Array liegt.
  const [organizerEmails, setOrganizerEmails] = React.useState<string[]>(() => {
    if (!editEvent || !editEvent.organizerEmails || editEvent.organizerEmails.length === 0) {
      return editEvent && editEvent.organizers && editEvent.organizers.length > 0
        ? editEvent.organizers.map(() => '')
        : [currentUser.email];
    }
    const names = editEvent.organizers || [];
    const emails = editEvent.organizerEmails;
    if (names.length === emails.length) return emails.slice();
    if (names.length !== emails.length) {
      console.warn(
        `[DEX] EventCreationPage: organizers/organizerEmails Längen-Mismatch (${names.length} vs ${emails.length}) — `
        + `padding auf max=${Math.max(names.length, emails.length)} mit leeren Slots. `
        + `Ursache: Legacy-Daten aus v10.0–v10.2-Closure-Bug. User muss fehlende Emails per Picker nachfüllen, dann Save heilt.`
      );
    }
    const max = Math.max(names.length, emails.length);
    const padded: string[] = [];
    for (let i = 0; i < max; i++) padded.push(emails[i] || '');
    return padded;
  });
  // isSearchingOrganizer entfällt seit v4.8.0 — Filter läuft sync gegen den
  // bereits geladenen DEX_Roles-State, kein Async-Spinner mehr nötig.
  const isSearchingOrganizer = false;
  const organizerTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  // v10.16: Optionaler Ansprechpartner pro Event. Reines Anzeige-Feld
  // (kein App-Login, keine Permissions) — z.B. die Person vor Ort die
  // Teilnehmer bei Fragen anrufen sollen. Erscheint auf Registration- +
  // MyEvents-Seite zusätzlich zu den Organizern.
  const [contactName, setContactName] = React.useState<string>(editEvent ? (editEvent.contactName || '') : '');
  const [contactEmail, setContactEmail] = React.useState<string>(editEvent ? (editEvent.contactEmail || '') : '');
  const [contactInfo, setContactInfo] = React.useState<string>(editEvent ? (editEvent.contactInfo || '') : '');
  // v28.5: EIN Organizer kann als Rückfragen-Kontakt markiert werden —
  // bekommt auf der Anmeldeseite einen orangen Badge + Legende (SP-Spalte
  // ContactOrganizerEmail, v26.18; UI erst jetzt verdrahtet).
  const [contactOrganizerEmail, setContactOrganizerEmail] = React.useState<string>(editEvent ? (editEvent.contactOrganizerEmail || '') : '');
  // v24.10 (Q2): Ansprechpartner standardmäßig eingeklappt — nur aufklappen,
  // wenn beim Bearbeiten bereits Daten hinterlegt sind.
  const [contactExpanded, setContactExpanded] = React.useState<boolean>(
    !!(editEvent && ((editEvent.contactName || '') || (editEvent.contactEmail || '') || (editEvent.contactInfo || '')))
  );

  // v6.19: QR-Code-Scanner pro Event (beliebiger Deloitte-User, kein Admin/Organizer nötig).
  // Getrennte State-Arrays für Namen + Emails (index-synchron). Sucht via Graph-API.
  const [qrScannerNames, setQrScannerNames] = React.useState<string[]>(
    editEvent && editEvent.qrScannerNames ? editEvent.qrScannerNames.slice() : []
  );
  const [qrScannerEmails, setQrScannerEmails] = React.useState<string[]>(
    editEvent && editEvent.qrScannerEmails ? editEvent.qrScannerEmails.slice() : []
  );
  const [qrScannerSearch, setQrScannerSearch] = React.useState('');
  const [qrScannerResults, setQrScannerResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  // v9.18: Debounce-Timer für Graph-Search (statt nur Role-Filter)
  const qrScannerTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  // v9.18/v9.20: Co-Organizer-State obsolet — Organizer-Picker selbst nimmt
  // jetzt alle Deloitte-User per Graph-Search. Felder bleiben für
  // Backward-Compat: Events vor v9.20 können noch _coOrganizers haben,
  // Access-Checks lesen sie weiterhin.
  const coOrganizerNames: string[] = (editEvent && editEvent.coOrganizerNames) ? editEvent.coOrganizerNames.slice() : [];
  const coOrganizerEmails: string[] = (editEvent && editEvent.coOrganizerEmails) ? editEvent.coOrganizerEmails.slice() : [];

  // v9.21: Test-Team pro Event — Personen die das Event im Entwurfsmodus
  // sehen + sich anmelden dürfen.
  const [testTeamNames, setTestTeamNames] = React.useState<string[]>(
    editEvent && editEvent.testTeamNames ? editEvent.testTeamNames.slice() : []
  );
  const [testTeamEmails, setTestTeamEmails] = React.useState<string[]>(
    editEvent && editEvent.testTeamEmails ? editEvent.testTeamEmails.slice() : []
  );
  const [testTeamSearch, setTestTeamSearch] = React.useState('');
  const [testTeamResults, setTestTeamResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const testTeamTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  const [organizerIncludeIntl, setOrganizerIncludeIntl] = React.useState(false);
  const [testTeamIncludeIntl, setTestTeamIncludeIntl] = React.useState(false);
  const [qrScannerIncludeIntl, setQrScannerIncludeIntl] = React.useState(false);

  // Massenimport-Modale für die drei Team-Felder (Co-Organizer, Test-Team,
  // Check-In Team). Pattern analog zum Audience-Massenimport (Sichtbarkeits-
  // Reiter), aber pro Team-Liste eigenes Modal mit eigenem Import-Target.
  const [bulkOrganizerOpen, setBulkOrganizerOpen] = React.useState(false);
  const [bulkTestTeamOpen, setBulkTestTeamOpen] = React.useState(false);
  const [bulkQrScannerOpen, setBulkQrScannerOpen] = React.useState(false);

  // v9.21: Active-From-Datum (optional) — Event auto-aktiv ab diesem Zeitpunkt.
  const [activeFrom, setActiveFrom] = React.useState(editEvent ? (editEvent.activeFrom || '') : '');
  // v23.14: Vorschau vor Aktivierung — nur relevant bei gesetztem „Aktiv ab".
  // true = Teilnehmer sehen das Event schon vorher als Vorschau (mit Hinweis
  // „Anmeldung ab …", noch nicht buchbar); false = vorher komplett unsichtbar.
  const [previewBeforeActive, setPreviewBeforeActive] = React.useState(editEvent ? !!editEvent.previewBeforeActive : false);
  const [location, setLocation] = React.useState(editEvent ? editEvent.location : '');
  // Strukturierte Adresse (Straße, Hausnummer, PLZ, Ort) - separat zum freien Location-Feld
  const [addrStreet, setAddrStreet] = React.useState(editEvent?.locationAddress?.street || '');
  const [addrHouseNo, setAddrHouseNo] = React.useState(editEvent?.locationAddress?.houseNo || '');
  const [addrZip, setAddrZip] = React.useState(editEvent?.locationAddress?.zip || '');
  const [addrCity, setAddrCity] = React.useState(editEvent?.locationAddress?.city || '');
  // v18.40: Überschreibbarer Outlook-Termin-Ort. Leer = automatischer Standard
  // (Veranstaltungsort + Adresse). Gefüllt = manueller Wert. Beim Edit nur als
  // Override vorbelegen, wenn der gespeicherte Wert vom Auto-Standard abweicht —
  // sonst bleibt das Feld leer und der Ort zieht weiter automatisch nach.
  const [outlookLocationOverride, setOutlookLocationOverride] = React.useState<string>(() => {
    if (!editEvent) return '';
    const auto = buildOutlookLocation(editEvent.location, editEvent.locationAddress);
    const stored = editEvent.outlookLocation || '';
    return (stored && stored !== auto) ? stored : '';
  });
  const [locationFilter, setLocationFilter] = React.useState(
    editEvent ? editEvent.locationAudience.join(', ') : ''
  );
  const [audience, setAudience] = React.useState(
    editEvent && editEvent.audienceFilter ? editEvent.audienceFilter.join(', ') : ''
  );
  // Default für neue Events: 'OR' — konsistent mit EventContext-Read-Fallback
  // und konservativer (UND-Verknüpfung kann Mitarbeiter unbeabsichtigt
  // ausschliessen). Bestehende Events behalten ihren gespeicherten Wert.
  const [filterMode, setFilterMode] = React.useState<'AND' | 'OR'>(
    editEvent ? editEvent.filterMode : 'OR'
  );
  const [description, setDescription] = React.useState(editEvent ? editEvent.description : '');
  // v28.7: „Keine Beschreibung nutzen" — reiner UI-Schalter im Wizard
  // (Default: Beschreibung nutzen). Anhaken leert die Beschreibung und
  // blendet den Editor-Zugang aus; gespeichert wird schlicht ''.
  const [noDescription, setNoDescription] = React.useState<boolean>(() => {
    // v28.79: beim Bearbeiten aus dem gespeicherten Flag vorbelegen —
    // sonst stand der Schalter nach dem Neuladen wieder auf „Beschreibung
    // nutzen", obwohl der Organizer sie bewusst weggelassen hatte.
    if (!editEvent) return false;
    if ((editEvent.description || '').trim()) return false;
    try {
      const ov = JSON.parse(editEvent.emailTemplateOverrides || '{}');
      return !!(ov && ov._noDescription);
    } catch { return false; }
  });
  // EventType wird nicht mehr als UI-Feld abgefragt (v5.2) — neue Events:
  // aus Template abgeleitet (b2run → 'B2Run', sonst → 'Other'). Bei Edit:
  // den gespeicherten Wert beibehalten. Die Variable wird weiterhin für
  // Card-Gradient + B2Run-spezifische Admin-Funktionen gebraucht.
  const [storedEventType] = React.useState<EventType>(editEvent ? editEvent.type : 'Other');
  const [startDate, setStartDate] = React.useState(editEvent ? isoToLocal(editEvent.startDate) : '');
  const [endDate, setEndDate] = React.useState(editEvent ? isoToLocal(editEvent.endDate) : '');
  const [registrationDeadline, setRegistrationDeadline] = React.useState(
    editEvent ? isoToLocal(editEvent.registrationDeadline) : ''
  );
  // v28.20: EXPLIZITE Anmeldefrist der Klammer (Piggyback _klammerDeadline).
  // Optional — leer heißt wie bisher: offen, solange ein Sub-Event offen ist.
  // Gesetzt + abgelaufen = Anmeldung fürs GESAMTE Event geschlossen.
  const [klammerDeadline, setKlammerDeadline] = React.useState(
    editEvent && editEvent.klammerDeadline ? isoToLocal(editEvent.klammerDeadline) : ''
  );
  const [lastDeregisterDate, setLastDeregisterDate] = React.useState(editEvent ? isoToLocal(editEvent.lastDeregisterDate) : '');
  // v29.25: Selbst-Abmeldung, zweistufig. Stufe 1: „Abmeldung durch User
  // ermöglichen" (Default ja; bei Nein gibt es keine Abmeldefrist und nur
  // Organizer/Admins melden ab — Piggyback _noSelfCancel). Stufe 2 (nur bei
  // Ja mit gesetzter Frist): „auch nach der Abmeldefrist erlauben" (Default
  // ja = Late-Cancel mit Organizer-Mail; bei Nein Piggyback
  // _noCancelAfterDeadline).
  // v29.38: Optionaler Teams-Besprechungslink. DEX legt KEIN Teams-Meeting an —
  // der Organizer fuegt den Link seiner eigenen Besprechung ein. Er landet als
  // Teilnahme-Block im Outlook-Termin.
  const [teamsLink, setTeamsLink] = React.useState<string>(editEvent?.teamsLink || '');
  /**
   * v30.26: Online-Meeting-Modus des Events — die Entscheidung trifft der
   * Organizer pro Event, NICHT der Flow global.
   *
   *  'none' — kein Online-Meeting (Präsenz).
   *  'own'  — der Organizer legt die Besprechung selbst in Outlook/Teams an
   *           und trägt den Link ein (v29.38-Feld). Er behält damit alle
   *           Besprechungsoptionen (Lobby, Aufzeichnung, Referenten).
   *  'auto' — DEX lässt den Flow eine echte Teams-Besprechung erzeugen
   *           (Spalte OutlookIsOnlineMeeting). Bequem und mit „Teilnehmen"-
   *           Knopf im Kalender, ABER der Termin gehört dem Gruppen-/
   *           No-Reply-Postfach: An den Besprechungsoptionen kann danach
   *           niemand mehr etwas ändern.
   *
   * Abgeleitet aus dem gespeicherten Stand, damit Bestandsevents (nur
   * teamsLink gepflegt) unverändert als 'own' erscheinen.
   */
  const [onlineMeetingMode, setOnlineMeetingMode] = React.useState<'none' | 'own' | 'auto'>(() => {
    if (editEvent?.outlookIsOnlineMeeting) return 'auto';
    if ((editEvent?.teamsLink || '').trim()) return 'own';
    return 'none';
  });
  const [userCancelAllowed, setUserCancelAllowed] = React.useState<boolean>(!(editEvent && editEvent.noSelfCancel));
  const [noCancelAfterDeadline, setNoCancelAfterDeadline] = React.useState<boolean>(!!(editEvent && editEvent.noCancelAfterDeadline));
  // v9.22: Auto-Fill der Deadlines wenn Start-Datum gesetzt wird und die
  // Deadlines noch leer sind. Default-Logik:
  //   - RegistrationDeadline: 7 Tage vor Event-Start
  //   - LastDeregisterDate: 3 Tage vor Event-Start
  // Der Organizer kann beides überschreiben — wir aktualisieren NICHT,
  // wenn der User schon einen Wert gesetzt hat.
  const autoFillRanRef = React.useRef(false);
  React.useEffect(() => {
    if (autoFillRanRef.current) return;
    if (!startDate) return;
    if (registrationDeadline || lastDeregisterDate) return;
    try {
      const start = new Date(startDate);
      if (isNaN(start.getTime())) return;
      const fmt = (d: Date): string => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}T${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
      const reg = new Date(start.getTime() - 7 * 24 * 60 * 60 * 1000);
      const lastCancel = new Date(start.getTime() - 3 * 24 * 60 * 60 * 1000);
      setRegistrationDeadline(fmt(reg));
      setLastDeregisterDate(fmt(lastCancel));
      autoFillRanRef.current = true;
    } catch { /* */ }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [startDate]);
  // v26.51: Wird das START-Datum GEÄNDERT (v.a. beim Bearbeiten: Event wird
  // verschoben), wandern gesetzte An-/Abmeldefristen relativ mit — gleiche
  // Logik wie vorher: War die Anmeldefrist 1 Woche vor dem Event, liegt sie
  // nach der Verschiebung wieder 1 Woche vor dem (neuen) Event-Beginn.
  const prevStartForShiftRef = React.useRef(startDate);
  React.useEffect(() => {
    const prev = prevStartForShiftRef.current;
    prevStartForShiftRef.current = startDate;
    if (!prev || !startDate || prev === startDate) return;
    const oldTs = new Date(prev).getTime();
    const newTs = new Date(startDate).getTime();
    if (!isFinite(oldTs) || !isFinite(newTs)) return;
    const delta = newTs - oldTs;
    if (!delta) return;
    const fmt = (d: Date): string => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}T${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
    const shift = (val: string): string => {
      if (!val) return val;
      const t = new Date(val).getTime();
      if (!isFinite(t)) return val;
      return fmt(new Date(t + delta));
    };
    setRegistrationDeadline(v => shift(v));
    setLastDeregisterDate(v => shift(v));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [startDate]);
  const [maxParticipants, setMaxParticipants] = React.useState(
    editEvent && editEvent.maxParticipants ? editEvent.maxParticipants.toString() : ''
  );
  const [unlimitedParticipants, setUnlimitedParticipants] = React.useState(
    !editEvent || !editEvent.maxParticipants || editEvent.maxParticipants === 0
  );
  // v22.37: Neues Event startet UNBEGRENZT → standardmäßig KEINE Warteliste.
  // Erst wenn der Organizer die Teilnehmerzahl begrenzt, wird die Warteliste
  // automatisch aktiviert (Default ja bei begrenzter Kapazität — siehe
  // Unbegrenzt-Toggle-onChange).
  const [waitlistEnabled, setWaitlistEnabled] = React.useState(
    editEvent && typeof editEvent.waitlistEnabled !== 'undefined' ? editEvent.waitlistEnabled : false
  );
  const [eventImageUrl, setEventImageUrl] = React.useState(editEvent ? (editEvent.imageUrl || '') : '');
  const [imageFile, setImageFile] = React.useState<File | null>(null);
  const [imagePreview, setImagePreview] = React.useState(editEvent ? (editEvent.imageUrl || '') : '');
  // v28.11: Frisch hochgeladenes ORIGINAL (vor dem Zuschnitt) + dessen
  // Seitenverhältnis. Wird nur persistiert (zweites Attachment
  // __eventimgorig__ + Piggyback _imageOrigUrl), wenn ein Querformat-
  // Original per Zuschnitt rund/quadratisch wurde — die Event-Liste zeigt
  // dann das Original als Kachel-Hintergrund.
  const [imageOrigFile, setImageOrigFile] = React.useState<File | null>(null);
  const [imageOrigAspect, setImageOrigAspect] = React.useState<number | null>(null);
  // v23.15: Bild-Editor (Zuschneiden / Kreis) offen?
  const [imageEditOpen, setImageEditOpen] = React.useState(false);
  // v26.97: Zuschneiden des Mail-/Outlook-Kopfbildes (nutzt dasselbe
  // ImageCropModal wie das Event-Bild). Ziel = welches Logo gerade zugeschnitten
  // wird ('email' oder 'outlook').
  const [logoCropTarget, setLogoCropTarget] = React.useState<'email' | 'outlook' | null>(null);
  // v28.30: Merker, ob das aktuell gesetzte Kopfbild per „Event-Foto
  // übernehmen" entstanden ist. Nur für die Optik des Knopfs (gruen + Haken =
  // „ist übernommen"). Bewusst NICHT persistiert: nach dem Neuladen zeigt die
  // Vorschau darunter ohnehin das echte Bild, und ein zweiter Klick auf den
  // Knopf ist folgenlos (er setzt dasselbe Bild noch einmal).
  const [emailLogoFromPhoto, setEmailLogoFromPhoto] = React.useState(false);
  const [outlookLogoFromPhoto, setOutlookLogoFromPhoto] = React.useState(false);
  // v27.11: Zuschneiden eines Sub-Event-Bildes — Index des Sub-Events in
  // `subEvents`, dessen Bild gerade im ImageCropModal offen ist (null = zu).
  const [subImageCropIdx, setSubImageCropIdx] = React.useState<number | null>(null);
  // v23.19: Optionale Pro-Ansicht-Darstellung (Zoom + vertikale Position).
  // Default leer = Standard (cover/zentriert) — nur auf Wunsch eingestellt.
  const [imageDisplay, setImageDisplay] = React.useState<{ card?: ImgView; hero?: ImgView }>(editEvent && editEvent.imageDisplay ? editEvent.imageDisplay : {});
  const [imageDisplayOpen, setImageDisplayOpen] = React.useState(false);

  // v29.47: Der Boot lädt die Anhänge nicht mehr mit. Beim Bearbeiten braucht
  // der Wizard sie (Schritt „Dokumente") — hier für das bearbeitete Event
  // nachholen, damit die Liste nicht fälschlich leer aussieht und ein Save die
  // bestehenden Dateien nicht als „entfernt" behandelt.
  React.useEffect(() => {
    if (editEvent?.id) void ensureEventDocuments([editEvent.id]);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editEvent?.id]);

  // v28.5: Bild als Banner über den Event-Infos (statt kompakt links) —
  // Organizer-Wahl, sinnvoll für breite Querformat-Fotos. Piggyback _imageBanner.
  const [imageBanner, setImageBanner] = React.useState<boolean>(!!(editEvent && editEvent.imageBanner));
  // v28.91: Sub-Events sind Termine (ein Tag je Sub-Event). Der Organizer legt
  // sie über einen Kalender an, die Anmeldeseite zeigt sie als Kalender.
  // Piggyback _subEventCalendar.
  const [subEventCalendar, setSubEventCalendar] = React.useState<boolean>(!!(editEvent && editEvent.subEventCalendar));
  // v29.22: Zum Löschen vorgemerkte, GESPEICHERTE Termine (Drafts mit dbId,
  // per Kalender-Klick oder X abgewählt). Sie bleiben hier geparkt statt
  // einfach zu verschwinden: Der Kalender zeigt sie ORANGE („wird beim
  // Speichern gelöscht"), ein erneuter Klick holt den Draft mitsamt allen
  // Einstellungen zurück. Vorher war ein abgewählter gespeicherter Tag
  // optisch nicht von „nie dagewesen" zu unterscheiden — und durch die
  // keyboard-selected-Färbung des DatePickers sah er sogar weiter grün aus,
  // als hätte das Abwählen nicht funktioniert. Die Speicher-Mechanik ändert
  // sich nicht: Nicht in subEvents = nicht in keptDbIds = wird beim Save
  // (nach der bestehenden Rückfrage) gelöscht.
  const [removedSavedSubs, setRemovedSavedSubs] = React.useState<SubEventDraft[]>([]);
  // v29.22: Die Terminliste unter dem Kalender ist standardmäßig EINGEKLAPPT
  // — bei 20 Terminen war sie eine Bildschirmseite Wiederholung dessen, was
  // der Kalender schon zeigt. Aufklappen nur bei Bedarf (Bearbeiten/Details).
  const [terminListOpen, setTerminListOpen] = React.useState(false);
  // v28.97: Genau EIN Sub-Event waehlbar statt beliebig vieler.
  const [subEventSingleChoice, setSubEventSingleChoice] = React.useState<boolean>(!!(editEvent && editEvent.subEventSingleChoice));
  // v30.61: Gebündelte Kommunikation (Mail / Kalender / QR getrennt schaltbar).
  // Gelesen aus den Overrides der Klammer — siehe utils/bundledComm.
  const [bundledComm, setBundledComm] = React.useState<BundledComm>(() => bundledCommOf(editEvent));
  // v28.10: Seitenverhältnis des Wizard-Bilds — die Banner-Option ist nur für
  // Querformat-Fotos sinnvoll und wird nur dann angeboten (Ratio >= 1.2).
  const [wizardImgAspect, setWizardImgAspect] = React.useState<number | null>(null);
  React.useEffect(() => {
    if (!imagePreview) { setWizardImgAspect(null); return; }
    let cancelled = false;
    const img = new Image();
    img.onload = () => { if (!cancelled && img.naturalHeight > 0) setWizardImgAspect(img.naturalWidth / img.naturalHeight); };
    img.src = imagePreview;
    return () => { cancelled = true; };
  }, [imagePreview]);
  React.useEffect(() => {
    // Nicht-Querformat (z.B. nach Kreis-Zuschnitt) → Banner-Flag zurücknehmen,
    // sonst bliebe ein unsichtbar gesetztes _imageBanner am Event hängen.
    if (wizardImgAspect != null && wizardImgAspect < 1.2 && imageBanner) setImageBanner(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [wizardImgAspect]);
  // v11.20: Re-sync useEffect aus v11.19 wieder rausgenommen — der hat
  // den Wizard-State mit stale-editEvent-Daten überschrieben (re-sync 2
  // mit helpText="" wurde im Maintainer-DevTools beobachtet, obwohl SP
  // nachweislich helpText="Test123" hatte). Das Aufrufen von
  // setCustomFields aus dem Effect heraus war zu fragil. Stattdessen
  // verlassen wir uns wieder auf den useState-Initializer + zusätzlich
  // ein detaillierteres Save-Log um zu sehen was *wirklich* an SP geht.
  const [customFields, setCustomFields] = React.useState<CustomFieldInput[]>(
    editEvent ? editEvent.eventSpecificFields.map(f => ({
      id: f.id, label: f.label, type: f.type, required: f.required,
      options: f.options ? [...f.options] : [], visible: true,
      ...(f.multi ? { multi: true } : {}),
      ...(f.defaultValue ? { defaultValue: f.defaultValue } : {}),
      ...(f.optionCategories && f.optionCategories.length > 0 ? { optionCategories: [...f.optionCategories] } : {}),
      ...(f.prefilterLabel ? { prefilterLabel: f.prefilterLabel } : {}),
      ...(f.helpText ? { helpText: f.helpText } : {}),
      ...(f.helpTextStyle === 'inline' ? { helpTextStyle: 'inline' as const } : {}),
      ...(f.showIf ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } } : {}),
      ...(f.onlyForGroup ? { onlyForGroup: f.onlyForGroup } : {}),
      // v11.94: confirmLabel beim Edit-Mount mit-übernehmen.
      ...(f.confirmLabel ? { confirmLabel: f.confirmLabel } : {}),
      // v17.20: EN-Varianten beim Edit-Mount mit-übernehmen.
      ...(f.labelEn ? { labelEn: f.labelEn } : {}),
      ...(f.helpTextEn ? { helpTextEn: f.helpTextEn } : {}),
      ...(f.confirmLabelEn ? { confirmLabelEn: f.confirmLabelEn } : {}),
      ...(f.optionsEn && f.optionsEn.length > 0 ? { optionsEn: [...f.optionsEn] } : {}),
      ...(f.externalLinks && f.externalLinks.length > 0 ? { externalLinks: f.externalLinks.map(x => ({ ...x })) } : {}),
      // v18.41: CC-bei-Mail-Flag beim Edit-Mount mit-übernehmen.
      ...(f.ccOnEmails ? { ccOnEmails: true } : {}),
      // v26.60: abgeschaltete Roommate-Benachrichtigung mit-übernehmen.
      ...(f.notifyRoommate === false ? { notifyRoommate: false } : {}),
      // v29.40: Verteiler-Begrenzung des Personen-Feldes mit-übernehmen.
      ...(f.audienceOnly ? { audienceOnly: true } : {}),
      // v29.20 (Audit A3): withTime (v24.25) und die daterange-Grenzen
      // (v28.63) fehlten in DIESEM Mapper — serializeCustomFields schreibt
      // nur, was im Draft steht, also entfernte jeder Edit-Save die
      // Einstellungen still: „Datum + Uhrzeit" wurde zum reinen Datum, das
      // buchbare Übernachtungs-Fenster verschwand. Die historische
      // Drop-Klasse (v11.21/v18.20/v19.20) saß diesmal im Lade-Pfad.
      ...(f.withTime ? { withTime: true } : {}),
      ...(f.rangeStart ? { rangeStart: f.rangeStart } : {}),
      ...(f.rangeEnd ? { rangeEnd: f.rangeEnd } : {}),
      ...(typeof f.maxNights === 'number' && f.maxNights > 0 ? { maxNights: f.maxNights } : {}),
    })) : []
  );
  const [outlookBody, setOutlookBody] = React.useState(editEvent ? reinsertOrganizerPlaceholder(stripOutlookWrapper(editEvent.outlookBody || ''), editEvent.organizers || []) : '');
  // Outlook-Termin-Header: beide Überschriften sind pro Event editierbar.
  // Default: eventTitle + formatiertes Startdatum. Parsed aus bestehendem
  // OutlookBody, falls der User sie schon angepasst hat.
  const [outlookHeading, setOutlookHeading] = React.useState(() => {
    if (editEvent) {
      const p = parseOutlookHeadings(editEvent.outlookBody || '');
      if (p.heading) return p.heading;
    }
    return editEvent ? (editEvent.title || '') : '';
  });
  const [outlookSubheading, setOutlookSubheading] = React.useState(() => {
    if (editEvent) {
      const p = parseOutlookHeadings(editEvent.outlookBody || '');
      if (p.subheading && p.subheading !== 'Event Details') return p.subheading;
    }
    return '';
  });
  // v18.42: Betreff des Outlook-Termins (leer = Event-Titel). Per-Tab gespiegelt
  // wie outlookHeading; persistiert in der DEX_Events-Spalte OutlookSubject.
  const [outlookSubject, setOutlookSubject] = React.useState<string>(editEvent?.outlookSubject || '');
  // v18.44: abweichendes Outlook-Datum (Top-Level). Leer = Event-Start/-Ende.
  // Als ISO gespeichert (wie Sub-Event-Datum); DatePicker konvertiert via isoToLocal.
  // v29.52: Ganztägiger Termin (Hauptevent). Die Krücke 00:00–23:59 blockiert in
  // Outlook den Tag als normalen Termin statt als Ganztags-Eintrag im Kopf —
  // der Unterschied fällt erst im Kalender des Teilnehmers auf.
  const [allDay, setAllDay] = React.useState<boolean>(!!(editEvent && editEvent.allDay));
  // v29.54: Der Termin blockiert den Kalender (Outlook `showAs`). Default ist
  // beschäftigt; bei Ganztags-Terminen ist das oft nicht gewollt, weil dann
  // ein ganzer Arbeitstag als belegt gilt.
  const [showAsFree, setShowAsFree] = React.useState<boolean>(!!(editEvent && editEvent.showAsFree));
  // v29.55: Bekommen die Organizer den Outlook-Termin JEDES Sub-Events? Der
  // Flow setzt requiredAttendees aus OrganizerEmail, und die steht auf jeder
  // Sub-Event-Zeile — bei 21 Tagen sind das 21 Blocker im Kalender, für Tage
  // ohne eigene Buchung. Positiv im UI, negativ gespeichert (skipOrganizerInvite).
  // Die Einstellung gilt event-weit: Klammer und alle Sub-Events bekommen
  // denselben Wert. Bestandsevents kommen mit false an und bleiben unverändert.
  const [orgGetsSubInvites, setOrgGetsSubInvites] = React.useState<boolean>(
    editEvent ? !editEvent.skipOrganizerInvite : true,
  );
  // Hat der Organizer die Entscheidung selbst getroffen? Dann nie überschreiben.
  const orgInvitesTouchedRef = React.useRef<boolean>(!!editEvent);
  // v29.56: Stand beim Oeffnen — daraus leitet sich beim Speichern ab, ob die
  // Organizer an BESTEHENDEN Outlook-Terminen nachträglich an- oder
  // abgemeldet werden müssen. SkipOrganizerInvite wirkt nämlich nur beim
  // ANLEGEN (requiredAttendees); ein bestehender Termin behält seine
  // Teilnehmerliste, bis jemand Einladen/Ausladen queued.
  const initialOrgGetsSubInvitesRef = React.useRef<boolean>(editEvent ? !editEvent.skipOrganizerInvite : true);
  const [outlookStartOverride, setOutlookStartOverride] = React.useState<string>(editEvent?.outlookStart || '');
  const [outlookEndOverride, setOutlookEndOverride] = React.useState<string>(editEvent?.outlookEnd || '');
  // Modal-State für den HTML-Editor (Outlook-Body + E-Mail-Templates)
  const [htmlEditorOpen, setHtmlEditorOpen] = React.useState(false);
  const [htmlEditorMode, setHtmlEditorMode] = React.useState<'outlook' | 'email' | 'description'>('outlook');
  const [htmlEditorTemplateType, setHtmlEditorTemplateType] = React.useState<string>('');
  const [emailLanguage, setEmailLanguage] = React.useState(
    editEvent
      ? (editEvent.emailLanguage || (locale === 'de' ? 'DE' : 'EN'))
      : (locale === 'de' ? 'DE' : 'EN')
  );
  const [disableEmails, setDisableEmails] = React.useState(editEvent ? !!editEvent.disableEmails : false);
  // v19.21: granulare Sub-Schalter unter dem Master „Bestätigungs-E-Mails":
  // einzeln die Anmelde- bzw. Abmelde-Bestätigung abschaltbar (Top-Level-Event).
  const [disableRegistrationEmail, setDisableRegistrationEmail] = React.useState(editEvent ? !!editEvent.disableRegistrationEmail : false);
  const [disableCancellationEmail, setDisableCancellationEmail] = React.useState(editEvent ? !!editEvent.disableCancellationEmail : false);
  // v19.23: Outlook-Absage = automatische Abmeldung vom Event (Flow-getrieben,
  // Top-Level-Event). Persistiert als DEX_Events-Spalte; der
  // DEX_OutlookDeclineHandler-Flow liest die Spalte und meldet die Person ab.
  const [autoDeregisterOnDecline, setAutoDeregisterOnDecline] = React.useState(editEvent ? !!editEvent.autoDeregisterOnDecline : false);
  const [inactiveHandling, setInactiveHandling] = React.useState<'notify' | 'autoderegister'>(editEvent && editEvent.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify');
  const [disableOutlook, setDisableOutlook] = React.useState(editEvent ? !!editEvent.disableOutlook : false);
  // v14.4: Acknowledgement, dass bei Top-Level-Kommunikation = AUS die
  // Teilnehmer sich für mindestens ein Sub-Event anmelden müssen. Vorausgewählt
  // für Events, die schon mit deaktivierter Kommunikation gespeichert sind
  // (alter Lauf ist bereits durch die Gate-Logik durchgekommen). Bei neuen
  // Events / frisch umgeschaltetem Toggle bleibt der Haken aus, der Save
  // ist dann blockiert bis bestätigt.
  const [mainCommDisabledAck, setMainCommDisabledAck] = React.useState<boolean>(
    !!editEvent && (!!editEvent.disableEmails || !!editEvent.disableOutlook),
  );
  // v14.5: Toggle „Anmeldung für mindestens ein Sub-Event verpflichtend".
  // Wird im RegistrationForm erzwungen — der Submit-Button blockiert, bis
  // der Teilnehmer ein Sub-Event angehakt hat. Sinnvoll wenn die Haupt-
  // Event-Kommunikation aus ist und alle Mails/Outlook-Termine nur über
  // die Sub-Events laufen.
  // v15.3: setRequireSubEventSelection wird nicht mehr direkt von der UI
  // aufgerufen — der Flag wird beim Save aus dem subEventsOnlyMode-Toggle
  // in Schritt 2 abgeleitet. State bleibt als Read-only für die Save-
  // Logik (siehe handleSubmit).
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [requireSubEventSelection, setRequireSubEventSelection] = React.useState<boolean>(
    !!editEvent && !!editEvent.requireSubEventSelection,
  );
  // v14.8: „Nur Sub-Events"-Modus. Wenn true, ist die Hauptevent-Anmeldung im
  // Teilnehmerformular ausgeblendet. Impliziert requireSubEventSelection=true.
  const [subEventsOnlyMode, setSubEventsOnlyMode] = React.useState<boolean>(
    !!editEvent && !!editEvent.subEventsOnlyMode,
  );
  // v14.8: Organizer-konfigurierbarer Begriff für die untergeordneten Events
  // (Standard „Sub-Event" / „Sub-Events", alternativ Workshop / Session etc.).
  // v15.9: separater `customTermMode`-Flag, damit „Eigene Bezeichnung…"
  // im Dropdown auch dann angeklebt bleibt, wenn beide Inputs noch leer
  // sind (sonst kippt die Heuristik unten auf 'subevent' zurück und die
  // Custom-Inputs verschwinden bevor der User tippen kann).
  const [customTermMode, setCustomTermMode] = React.useState<boolean>(false);
  // v29.60: Grammatisches Geschlecht der Bezeichnung. '' = wie bisher raten.
  // Gebraucht wird der AKKUSATIV („wähle mindestens einen Office-Tag aus"),
  // und den kann man aus dem Wort nicht ableiten — deshalb gefragt statt geraten.
  const [childGender, setChildGender] = React.useState<'' | 'm' | 'f' | 'n'>(
    (editEvent && editEvent.childEventTermGender) || '',
  );
  const [childTermSingular, setChildTermSingular] = React.useState<string>(
    (editEvent && editEvent.childEventTermSingular) || '',
  );
  const [childTermPlural, setChildTermPlural] = React.useState<string>(
    (editEvent && editEvent.childEventTermPlural) || '',
  );
  // v8.5: Organizer-BCC-Modi (Anmeldung + Abmeldung).
  const [notifyOrgRegisterMode, setNotifyOrgRegisterMode] = React.useState<'never' | 'always' | 'fromDate'>(
    editEvent ? (editEvent.notifyOrgRegisterMode || 'never') : 'never'
  );
  // v29.19: Wie alle anderen Datumsfelder über isoToLocal laden — der rohe
  // SP-Wert ist UTC-ISO mit „Z", und berlinLocalToUtcIso beim Save hängt ein
  // weiteres „Z" an → ungültig → es wurde '' in die Spalte geschrieben. Wer
  // das Feld beim Edit nicht neu anfasste, verlor sein „BCC ab"-Datum still;
  // der Modus blieb auf FromDate stehen und die Organizer-Kopie feuerte nie.
  const [notifyOrgRegisterFromDate, setNotifyOrgRegisterFromDate] = React.useState<string>(
    editEvent && editEvent.notifyOrgRegisterFromDate ? isoToLocal(editEvent.notifyOrgRegisterFromDate) : ''
  );
  const [notifyOrgCancelMode, setNotifyOrgCancelMode] = React.useState<'never' | 'always' | 'afterDeadline'>(
    // v10.17+: Default für neue Events ist 'afterDeadline' (Erst nach der
    // letzten Abmeldemöglichkeit) — sonst flutet jede Stornierung den
    // Organizer-Posteingang. User-Wunsch.
    editEvent ? (editEvent.notifyOrgCancelMode || 'never') : 'afterDeadline'
  );
  // v8.6: Exclude-Liste — explizit ausgeschlossene User (überschreiben den
  // Sichtbarkeits-Filter). UI: Modal "Sichtbare Personen anzeigen".
  const [excludedUsers, setExcludedUsers] = React.useState<string[]>(
    editEvent ? (editEvent.excludedUsers || []) : []
  );
  // v11.88: Demo-Auswahl-Modal — der „Demo"-Button öffnet einen Dialog
  // mit vier Vorlagen-Karten (Standard, Mit Gruppen, Mit Sub-Event,
  // Mit Sub-Event + Team). Klick auf eine Karte füllt das Formular
  // mit der jeweiligen Variante und schliesst das Modal.
  const [showDemoVariantModal, setShowDemoVariantModal] = React.useState<boolean>(false);
  // v24.9 (E): „Eigenes Event als Vorlage" — aufklappbare Kachelgalerie der
  // bisherigen Events des Organizers; Auswahl lädt Einstellungen + Bild.
  const [showTemplatePicker, setShowTemplatePicker] = React.useState<boolean>(false);
  const [templateLoadingId, setTemplateLoadingId] = React.useState<string | null>(null);
  // v17.21: Modal nach erfolgreichem Speichern — fragt den Organizer, ob er
  // eine A4-Zusammenfassung des Events herunterladen möchte. Pending-Payload
  // hält die Info für den `dex-event-submit-success`-Dispatch, der erst
  // gefeuert wird, wenn der User im Modal eine Auswahl getroffen hat.
  const [showSummaryModal, setShowSummaryModal] = React.useState<boolean>(false);
  const [pendingSuccessDispatch, setPendingSuccessDispatch] = React.useState<{
    title: string; eventId: string; type: 'create' | 'update';
  } | null>(null);
  // v17.22: Unmount-Safety. Der Success-Dispatch (dex-event-submit-success,
  // treibt Erfolgs-Banner + Auto-Navigation in DexEventPlatform) läuft erst,
  // wenn der User im Summary-Modal eine Auswahl trifft. Verlässt er den
  // Wizard vorher (Header-Navigation, Browser-Back, Tab-Eviction), würde der
  // Dispatch sonst verloren gehen — Folge: kein Banner, kein Redirect, User
  // denkt der Save sei fehlgeschlagen. Dieser Ref + Cleanup-Effect feuert den
  // Dispatch beim Unmount nach, falls er noch aussteht.
  const pendingSuccessDispatchRef = React.useRef<{ title: string; eventId: string; type: 'create' | 'update' } | null>(null);
  React.useEffect(() => {
    return () => {
      const pending = pendingSuccessDispatchRef.current;
      if (pending) {
        pendingSuccessDispatchRef.current = null;
        try {
          window.dispatchEvent(new CustomEvent('dex-event-submit-success', { detail: pending }));
        } catch { /* */ }
      }
    };
  }, []);
  // v19.x: Der gesamte Ausschluss-Modal-State (resolved Members, Suche,
  // Tabellen-Filter, Sortierung, Pagination) ist nach <AudiencePicker>
  // gewandert. Hier bleibt nur die persistierte `excludedUsers`-Liste (oben),
  // die als Prop in den Picker durchgereicht wird.
  // v9.16: neue Events starten standardmäßig als Test-Event — der Organizer
  // kann sich erst alles in Ruhe anschauen, das Test-Team probiert die
  // Anmeldung durch, und erst wenn alles passt wird der Schalter rausgenommen.
  const [isFictive, setIsFictive] = React.useState(editEvent ? !!editEvent.isFictive : true);
  // v18.9: Organizer-Anzeige (Chips mit Name + Foto) auf Anmelde-Seite +
  // „Meine Events" ausblenden. Rein visuell — Rechte/Mails unberührt.
  const [hideOrganizer, setHideOrganizer] = React.useState(editEvent ? !!editEvent.hideOrganizer : false);
  // v24.15: Wenn „Organizer ausblenden" an ist UND es mehrere Organizer gibt,
  // kann der Organizer stattdessen NUR EINZELNE ausblenden (Piggyback
  // `_hideOrgIndividual`). false = alle ausblenden; true = nur die angeklickten.
  const [hideOrganizerIndividualOnly, setHideOrganizerIndividualOnly] = React.useState(editEvent ? !!editEvent.hideOrganizerIndividualOnly : false);
  // v23.6: „Assistenzen sehen das Event generell" (Piggyback _assistantsCanSee).
  // Wenn aktiv, sehen Personen mit dem Job-Title „Assistenz" das Event auch
  // dann, wenn Standort-/Verteiler-Filter sie sonst ausschließen würden —
  // damit sie stellvertretend (z.B. für einen Partner) anmelden können.
  const [assistantsCanSee, setAssistantsCanSee] = React.useState(editEvent ? !!editEvent.assistantsCanSee : false);
  // v23.25: Organizer auf der Anmeldeseite groß (Foto + Mail + Rolle direkt
  // sichtbar) statt klein als Chip mit Hover (Piggyback _organizerDisplayLarge).
  const [organizerDisplayLarge, setOrganizerDisplayLarge] = React.useState(editEvent ? !!editEvent.organizerDisplayLarge : false);
  // v24.8 (J): EINZELNE Organizer von der Anzeige ausnehmen (Klick auf den Chip).
  // Sie behalten alle Rechte/Mails — sie werden nur nicht als Ansprechpartner
  // auf der Anmelde-Seite gezeigt. Piggyback `_hiddenOrganizers` (E-Mails, lc).
  const [hiddenOrganizerEmails, setHiddenOrganizerEmails] = React.useState<string[]>(
    editEvent && editEvent.hiddenOrganizerEmails ? editEvent.hiddenOrganizerEmails.map(e => (e || '').toLowerCase()).filter(Boolean) : []
  );
  const toggleOrganizerHidden = (email: string): void => {
    const lc = (email || '').toLowerCase();
    if (!lc) return;
    setHiddenOrganizerEmails(prev => prev.indexOf(lc) >= 0 ? prev.filter(x => x !== lc) : [...prev, lc]);
  };
  // Nur im Edit-Modus: standardmäßig wird der Outlook-Termin NICHT angefasst,
  // damit bei kleinen Änderungen (z.B. Description) nicht unnötig eine
  // "Updated meeting"-Benachrichtigung an alle Teilnehmer geht. Der Organizer
  // muss die Checkbox aktiv setzen wenn er möchte dass Titel/Start/Ende im
  // Outlook-Termin aktualisiert werden.
  const [triggerOutlookUpdate, setTriggerOutlookUpdate] = React.useState(false);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [emailTemplates, setEmailTemplates] = React.useState<Array<{ id: number; templateType: string; language: string; subject: string; heading: string; headingColor: string; bodyHtml: string }>>([]);
  const [emailTemplateOverrides, setEmailTemplateOverrides] = React.useState<Record<string, EmailOverrideEntry>>(
    editEvent?.emailTemplateOverrides ? (() => {
      try {
        const parsed = JSON.parse(editEvent.emailTemplateOverrides);
        // v11.39: Alle Piggyback-Keys rausstrippen — sie werden in separaten
        // States gehalten (emailLogoPreview, outlookLogoPreview, testTeamEmails
        // etc.) und beim Speichern frisch dazugemerged. Wenn sie hier
        // mitgeschleppt werden, überschreibt der spread `...emailTemplateOverrides`
        // am Ende von handleSubmit die frisch berechneten Werte und das
        // Entfernen z.B. eines Test-Team-Mitglieds bleibt ohne Wirkung.
        const {
          _eventLogo, _outlookLogo, _outlookLogoSameAsMail, _b2run,
          _qrScanners, _coOrganizers, _testTeam,
          _splitDisplayOrderReversed,
          _requireSubEventSelection,
          _subEventsOnlyMode, _subEventsDisabled, _imageBanner, _childEventTerm,
          _inheritFlags, _hideOrganizer, _headerImageLayout,
          // v22.78/v23.6: diese Piggyback-Keys MÜSSEN ebenfalls gestrippt
          // werden — sonst überschreibt der stale Wert aus dem geladenen Blob
          // beim Edit-Save (letzter Spread `...topOverrides`) das frisch
          // berechnete Flag, d.h. Abwählen bliebe ohne Wirkung.
          _teamTerm, _teamMembersCannotCreate, _assistantsCanSee, _previewBeforeActive, _imageDisplay,
          _organizerDisplayLarge, _hiddenOrganizers, _hideOrgIndividual, _mainEventLabel,
          _imageOrigUrl, _klammerDeadline,
          // v28.39: Hotel-Planung wird ausschliesslich im Organizer Center gepflegt.
          // Stripping verhindert, dass ein parallel offener Wizard beim Speichern
          // einen veralteten Stand zurückschreibt.
          _hotels, _hotelStays, _hotelVisible, _hotelRules,
          // v28.79: „Keine Beschreibung nutzen"-Flag (s. noDescriptionConfig).
          _noDescription,
          // v28.91: Kalender-Modus der Sub-Events (s. subEventCalendarConfig).
          _subEventCalendar, _subEventSingleChoice,
          // v29.25: Abmelde-Sperren (s. userCancelAllowed / noCancelAfterDeadline).
          _noSelfCancel, _noCancelAfterDeadline,
          // v29.38: Teams-Link (s. teamsLinkConfig).
          _teamsLink,
          // v29.66: Abrechnungs-Piggyback (F&A-Pilot) — lebt in eigenen States.
          _billing,
          // v29.67: Freischalt-Regel der Kalender-Termine — eigene States.
          _subEventOpenRule,
          // v29.75: Sichtbarkeit-für-alle-Sub-Events-Haken — eigener State.
          _visAllSubs,
          // v29.76: Rollierende Fristen der Kalender-Termine — eigene States.
          _subDeadlineRule,
          // v30.61: Gebündelte Kommunikation — eigene States (s. bundledComm).
          _commBundledMail, _commBundledOutlook, _commBundledQr,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          ...rest
        } = parsed as Record<string, unknown>;
        // Variablen nur destrukturiert, um sie aus `rest` zu entfernen.
        void _eventLogo; void _outlookLogo; void _outlookLogoSameAsMail; void _b2run;
        void _billing; void _subEventOpenRule; void _visAllSubs; void _subDeadlineRule;
        void _qrScanners; void _coOrganizers; void _testTeam;
        void _splitDisplayOrderReversed; void _requireSubEventSelection;
        void _subEventsOnlyMode; void _subEventsDisabled; void _imageBanner; void _childEventTerm;
        void _inheritFlags; void _hideOrganizer; void _headerImageLayout;
        void _teamTerm; void _teamMembersCannotCreate; void _assistantsCanSee; void _previewBeforeActive; void _imageDisplay;
        void _organizerDisplayLarge; void _hiddenOrganizers; void _hideOrgIndividual; void _mainEventLabel;
        void _imageOrigUrl; void _klammerDeadline; void _noDescription;
        void _subEventCalendar; void _subEventSingleChoice;
        void _noSelfCancel; void _noCancelAfterDeadline; void _teamsLink;
        void _hotels; void _hotelStays; void _hotelVisible; void _hotelRules;
        void _commBundledMail; void _commBundledOutlook; void _commBundledQr;
        return rest as Record<string, EmailOverrideEntry>;
      } catch { return {}; }
    })() : {}
  );
  // editingTemplate state entfällt seit Modal-Migration v4.7.0
  // Custom Event-Logo für E-Mails (ersetzt das DEX-Orb in E-Mails).
  const [emailLogoPreview, setEmailLogoPreview] = React.useState(() => {
    if (!editEvent?.emailTemplateOverrides) return '';
    try { const o = JSON.parse(editEvent.emailTemplateOverrides); return o._eventLogo || ''; } catch { return ''; }
  });
  // Custom Event-Logo für Outlook-Termin (ersetzt das DEX-Orb im Termin-Body).
  // Separat vom Mail-Logo, damit man z.B. in Mails das neutrale DEX-Logo lassen
  // und im Outlook-Termin ein event-spezifisches Bild anzeigen kann.
  const [outlookLogoPreview, setOutlookLogoPreview] = React.useState(() => {
    if (!editEvent?.emailTemplateOverrides) return '';
    try { return readOutlookLogo(JSON.parse(editEvent.emailTemplateOverrides)); } catch { return ''; }
  });
  // v18.73: Header-Bild (Event-Bild = {{ORB_URL}}) Größe + Innenabstand pro
  // Event. Gilt für Mail- UND Outlook-Termin-Kopf. Persistiert als Piggyback
  // `_headerImageLayout` in EmailTemplateOverrides. Default = bisheriges
  // Layout (Breite 180, Innenabstand 30/30).
  // v29.29: NEUE Events starten mit dem Vollbild-Kopf (Bild über die ganze
  // Mailbreite) — das ist der Regelfall, das kleine zentrierte Bild war eher
  // die Ausnahme. Bestehende Events ohne gespeichertes Layout behalten
  // bewusst 180/30/30: Ihre Mails sähen sonst nach dem nächsten Speichern
  // ungefragt anders aus.
  const [headerImageLayout, setHeaderImageLayout] = React.useState<{ width: number; paddingV: number; paddingH: number }>(() => {
    const legacyDef = { width: 180, paddingV: 30, paddingH: 30 };
    const fullWidth = { width: 600, paddingV: 0, paddingH: 0 };
    if (!editEvent) return fullWidth;
    if (!editEvent.emailTemplateOverrides) return legacyDef;
    try {
      const o = JSON.parse(editEvent.emailTemplateOverrides);
      const il = o._headerImageLayout || {};
      return {
        width: typeof il.width === 'number' && il.width > 0 ? il.width : 180,
        paddingV: typeof il.paddingV === 'number' && il.paddingV >= 0 ? il.paddingV : 30,
        paddingH: typeof il.paddingH === 'number' && il.paddingH >= 0 ? il.paddingH : 30,
      };
    } catch { return legacyDef; }
  });
  // v19.20: Snapshot des initialen Header-Bild-Layouts (Breite/Innenabstand)
  // beim Edit-Mount. Eine reine Layout-Änderung verändert NICHT den rohen
  // Outlook-Body-Text (das Layout wird erst beim Wrappen via buildOutlookBody
  // angewendet) — der Outlook-Änderungs-Detektor hätte sie deshalb übersehen.
  // Wir vergleichen das aktuelle Layout gegen diesen Snapshot, damit eine
  // Größen-/Abstands-Änderung das „Outlook-Termin aktualisieren?"-Modal genauso
  // öffnet wie eine Textänderung. useRef fixiert den Wert beim ersten Render.
  const initialHeaderImageLayoutRef = React.useRef<{ width: number; paddingV: number; paddingH: number }>(headerImageLayout);
  // v29.38: Gleiche Mechanik für den Teams-Link — er steht nicht im rohen
  // Termin-Text (er wird erst beim Wrappen angehängt), ändert den Termin aber
  // sichtbar. Ohne Snapshot bliebe eine reine Link-Änderung für den
  // Update-Detektor unsichtbar und der Termin behielte den alten Stand.
  const effTeamsLink = (): string => (onlineMeetingMode === 'own' ? teamsLink.trim() : '');
  /**
   * Link, der in den TERMIN-BODY wandert — nicht derselbe wie `effTeamsLink()`,
   * der nur den gespeicherten `_teamsLink` steuert.
   *
   * **v30.40: Im Modus „DEX erzeugt den Link" steht hier wieder nichts.** Der
   * Weg über die Marke `{{TEAMS_URL}}` (v30.27–v30.39) ist gescheitert, und
   * zwar erst im letzten möglichen Moment — im fertigen Termin beim Teilnehmer:
   *
   * 1. Die App schrieb `<a href="{{TEAMS_URL}}">…</a>` in den Body.
   * 2. Der Flow holte den Body per Graph und ersetzte die Marke durch die echte
   *    `joinUrl`. Der PATCH lief mit 200 durch.
   * 3. Im Termin stand danach `[https://teams.microsoft.com/l/meetup-join/…]An
   *    Microsoft-Teams-Besprechung teilnehmen` — der Anker war zu Text
   *    zerfallen. Vorher, mit der unersetzten Marke im href, war es noch ein
   *    Knopf.
   *
   * Beobachtet, nicht bewiesen: Exchange normalisiert den Body eines
   * Online-Meetings und lässt einen Anker auf die eigene joinUrl nicht stehen.
   * Was sich prüfen ließ, spricht dafür — die Degradierung trat exakt mit
   * unserem PATCH ein, an keiner früheren Stelle.
   *
   * Entscheidend ist aber nicht die Ursache, sondern dass der Kasten, den
   * Exchange unter die Karte hängt, ohnehin bleiben MUSS: Ihn zu entfernen
   * hieße, den Meeting-Blob aus dem Body zu werfen, und das deaktiviert die
   * Besprechung (Graph-Referenz `event: update`). Er trägt Join-Link,
   * Meeting-ID und Passcode — mehr, als der DEX-Block je hatte. Ein zweiter,
   * kaputter Link darüber macht den Termin nur schlechter.
   *
   * Für einen SELBST eingetragenen Link bleibt alles wie bisher: Da gibt es
   * keine Exchange-Normalisierung, der Block rendert sauber, und ohne ihn stünde
   * der Link nirgends.
   */
  const outlookTeamsLink = (): string => (
    onlineMeetingMode === 'own' ? teamsLink.trim() : ''
  );
  const initialTeamsLinkRef = React.useRef<string>(teamsLink);
  // v30.26: Der Modus zählt für den Outlook-Update-Detektor genauso wie der
  // Link selbst — ein Wechsel von „eigener Link" auf „DEX erzeugt" ändert
  // sowohl den Termin-Text (Link fällt weg) als auch den Termin-Typ
  // (isOnlineMeeting). Ohne diesen Vergleich bliebe der bestehende Termin
  // stehen, weil sich der reine teamsLink-String nicht bewegt hat.
  const initialOnlineMeetingModeRef = React.useRef<'none' | 'own' | 'auto'>(onlineMeetingMode);
  const onlineMeetingChanged = (): boolean =>
    onlineMeetingMode !== initialOnlineMeetingModeRef.current
    || effTeamsLink() !== (initialOnlineMeetingModeRef.current === 'own' ? initialTeamsLinkRef.current.trim() : '');
  // v18.73: Piggyback-Konfig für den Save (leer wenn alles auf Default steht —
  // dann wird der Key gar nicht geschrieben). Wird in Create- UND Edit-Pfad
  // sowie in die Sub-Event-Overrides gemerged.
  const headerImageLayoutConfig = (headerImageLayout.width !== 180 || headerImageLayout.paddingV !== 30 || headerImageLayout.paddingH !== 30)
    ? { _headerImageLayout: { width: headerImageLayout.width, paddingV: headerImageLayout.paddingV, paddingH: headerImageLayout.paddingH } }
    : {};
  /**
   * v28.29: Kopfbild-Layout für EINEN konkreten Outlook-/Mail-Body. Breite und
   * Innenabstand stellt der Organizer für sein FOTO ein („Volle Breite" = 600px).
   * Fällt ein Termin mangels eigenem Bild auf das Standard-DEX-Logo (Orb) zurück,
   * wurde dieses bisher ebenfalls 600px breit gerendert — in Outlook ein
   * bildschirmfüllender, unten abgeschnittener Orb. Ohne eigenes Bild deshalb
   * max. 180px mit Mindestabstand.
   */
  const headerLayoutFor = (logoB64: string): { imageWidth: number; imagePaddingV: number; imagePaddingH: number } => {
    const hasOwn = !!(logoB64 && logoB64.trim());
    return {
      imageWidth: hasOwn ? headerImageLayout.width : Math.min(headerImageLayout.width, 180),
      imagePaddingV: hasOwn ? headerImageLayout.paddingV : Math.max(headerImageLayout.paddingV, 20),
      imagePaddingH: hasOwn ? headerImageLayout.paddingH : Math.max(headerImageLayout.paddingH, 20),
    };
  };
  // v26.95: Das Event-Foto als Mail-/Outlook-Kopfbild übernehmen. Quelle ist der
  // frisch gewählte File (imageFile), sonst die Vorschau (Data-URL direkt, http-
  // URL bestehender Events wird geladen). In JEDEM Fall auf 600px komprimiert,
  // damit die Base64-Größe für die Mail-Pipeline handhabbar bleibt.
  const fileToBase64 = (file: File): Promise<string> =>
    new Promise<string>(resolve => { const r = new FileReader(); r.onload = e => resolve((e.target?.result as string) || ''); r.onerror = () => resolve(''); r.readAsDataURL(file); });
  // v28.10: Base64-Logos hart auf Mail-taugliche Größe bringen. Ungebremste
  // Logos landeten bis zu DREIMAL im selben Save-Payload (OutlookBody via
  // {{ORB_URL}}, EmailTemplateOverrides._eventLogo/_outlookLogo und
  // EmailImageBase64) und rissen das SharePoint-REST-Limit von 2 MB
  // („The request message is too big"). Ab ~400 KB wird auf 600px
  // runterskaliert; schlägt das fehl, bleibt der Originalwert.
  const shrinkLogoB64 = async (b64: string): Promise<string> => {
    // v28.31: Schwelle von 400 KB auf 200 KB und in Stufen verkleinern, bis das
    // Bild unter ~250 KB liegt. Dasselbe Base64 steckt DREIMAL im Save-Payload
    // (OutlookBody, _eventLogo, _outlookLogo) — bei 800 KB je Vorkommen reisst
    // ein einziges Event das SharePoint-Limit von 2 MB. Ausgabe immer als JPEG
    // auf weissem Grund (siehe compressImage): PNG-Fotos sind der Hauptgrund
    // für die Ausreisser.
    const TARGET = 250_000;
    if (!b64 || b64.indexOf('data:') !== 0 || b64.length <= 200_000) return b64;
    let best = b64;
    for (const w of [600, 480, 360]) {
      try {
        const resp = await fetch(best);
        const blob = await resp.blob();
        const f = new File([blob], 'logo.jpg', { type: blob.type || 'image/jpeg' });
        const out = await fileToBase64(await compressImage(f, w, 0.82, true));
        if (out && out.length < best.length) best = out;
        if (best.length <= TARGET) break;
      } catch { break; }
    }
    return best;
  };
  const applyEventPhotoToLogo = async (setter: (b64: string) => void): Promise<string> => {
    try {
      let b64 = '';
      // v28.29 BUG-FIX: „Event-Foto verwenden" nahm bisher IMMER den Zuschnitt
      // (imageFile/imagePreview = das runde bzw. quadratisch beschnittene
      // Event-Bild). Im Mail-/Outlook-Kopf steht aber ein RECHTECK — das Foto
      // kam dort sichtbar abgeschnitten an, ohne dass der Organizer das
      // gewollt hätte. Wenn ein unbeschnittenes Original existiert (frischer
      // Upload: imageOrigFile; gespeichertes Event: editEvent.imageOrigUrl),
      // wird jetzt DIESES übernommen.
      if (imageOrigFile) {
        b64 = await fileToBase64(await compressImage(imageOrigFile, 600, 0.85, true));
      } else if (editEvent && editEvent.imageOrigUrl) {
        try {
          const resp = await fetch(editEvent.imageOrigUrl, { credentials: 'include' });
          const blob = await resp.blob();
          const f = new File([blob], 'event-photo.jpg', { type: blob.type || 'image/jpeg' });
          b64 = await fileToBase64(await compressImage(f, 600, 0.85, true));
        } catch { /* Original nicht ladbar → unten auf den Zuschnitt zurückfallen */ }
      }
      if (b64) {
        setter(b64);
        return b64;
      }
      if (imageFile) {
        b64 = await fileToBase64(await compressImage(imageFile, 600, 0.85, true));
      } else if (imagePreview && imagePreview.indexOf('data:') === 0) {
        // v28.10: Frischer Zuschnitt (Data-URL) ebenfalls komprimieren —
        // vorher ging das volle Bild unkomprimiert ins Logo (2-MB-Falle).
        b64 = await shrinkLogoB64(imagePreview);
      } else if (imagePreview) {
        const resp = await fetch(imagePreview, { credentials: 'include' });
        const blob = await resp.blob();
        const f = new File([blob], 'event-photo.jpg', { type: blob.type || 'image/jpeg' });
        b64 = await fileToBase64(await compressImage(f, 600, 0.85, true));
      }
      if (b64) setter(b64);
      else showAlert(isDe ? 'Kein Event-Foto vorhanden — bitte zuerst oben ein Bild hochladen.' : 'No event photo yet — please upload an image above first.', { variant: 'error' });
      return b64;
    } catch {
      showAlert(isDe ? 'Das Event-Foto konnte nicht übernommen werden.' : 'Could not use the event photo.', { variant: 'error' });
      return '';
    }
  };
  // v27.2: Größensteuerung fürs Kopfbild als wiederverwendbarer Block (Schritt 23
  // UND 24) — inkl. verkleinerter Live-Vorschau, die zeigt, wie groß das Bild im
  // Mail-/Outlook-Kopf steht. `headerImageLayout` gilt event-weit (Mail + Outlook).
  const renderHeaderSizeControl = (previewSrc: string, note?: string): React.ReactElement => {
    const PREV_W = 260; const sc = PREV_W / 600;
    const isFullWidthPreset = headerImageLayout.width === 600 && headerImageLayout.paddingV === 0 && headerImageLayout.paddingH === 0;
    const isDefaultPreset = headerImageLayout.width === 180 && headerImageLayout.paddingV === 30 && headerImageLayout.paddingH === 30;
    const numInput = (val: number, min: number, max: number, def: number, set: (n: number) => void): React.ReactElement => (
      <input type="number" min={min} max={max} step={min === 80 ? 10 : 2} value={val}
        onChange={e => set(Math.max(min, Math.min(max, parseInt(e.target.value, 10) || def)))}
        style={{ width: 78, height: 28, fontSize: '0.82rem', borderRadius: 4, border: '1px solid var(--dex-gray-300)', padding: '0 8px' }} />
    );
    const lbl: React.CSSProperties = { display: 'flex', flexDirection: 'column', gap: 3, fontSize: '0.7rem', color: 'var(--dex-gray-600)' };
    return (
      <div style={{ marginTop: 12, padding: '10px 12px', background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)', borderRadius: 8 }}>
        <div style={{ fontSize: '0.7rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', letterSpacing: 0.3, marginBottom: 8 }}>
          {isDe ? 'BILDGRÖSSE IM KOPF — gilt für Mail & Outlook-Termin' : 'HEADER IMAGE SIZE — applies to mail & Outlook invite'}
        </div>
        <div style={{ display: 'flex', gap: 18, flexWrap: 'wrap', alignItems: 'flex-start' }}>
          <div style={{ display: 'flex', alignItems: 'flex-end', gap: 12, flexWrap: 'wrap' }}>
            <label style={lbl}>{isDe ? 'Breite (px)' : 'Width (px)'}{numInput(headerImageLayout.width, 80, 600, 180, n => setHeaderImageLayout(p => ({ ...p, width: n })))}</label>
            <label style={lbl}>{isDe ? 'Abstand seitl.' : 'Padding sides'}{numInput(headerImageLayout.paddingH, 0, 80, 0, n => setHeaderImageLayout(p => ({ ...p, paddingH: n })))}</label>
            <label style={lbl}>{isDe ? 'Abstand ob./unt.' : 'Padding top/bot.'}{numInput(headerImageLayout.paddingV, 0, 80, 0, n => setHeaderImageLayout(p => ({ ...p, paddingV: n })))}</label>
            {/* v28.31: Beide Voreinstellungen zeigen jetzt an, WELCHE gerade
                aktiv ist. Vorher war „Volle Breite" immer gruen und „Standard"
                immer grau — auch wenn tatsaechlich 180/30/30 (= Standard) stand. */}
            <button type="button" onClick={() => setHeaderImageLayout({ width: 600, paddingV: 0, paddingH: 0 })}
              title={isDe ? 'Bild füllt den Kopf über die volle Breite' : 'Image fills the header full width'}
              style={{ height: 28, padding: '0 12px', fontSize: '0.72rem', fontWeight: isFullWidthPreset ? 700 : 600, cursor: 'pointer', background: isFullWidthPreset ? 'var(--dex-green, #86bc25)' : 'transparent', color: isFullWidthPreset ? '#fff' : 'var(--dex-gray-600)', border: isFullWidthPreset ? 'none' : '1px solid var(--dex-gray-300)', borderRadius: 6 }}>
              {isFullWidthPreset ? '✓ ' : ''}{isDe ? 'Volle Breite' : 'Full width'}
            </button>
            <button type="button" onClick={() => setHeaderImageLayout({ width: 180, paddingV: 30, paddingH: 30 })}
              style={{ height: 28, padding: '0 12px', fontSize: '0.72rem', fontWeight: isDefaultPreset ? 700 : 600, cursor: 'pointer', background: isDefaultPreset ? 'var(--dex-green, #86bc25)' : 'transparent', color: isDefaultPreset ? '#fff' : 'var(--dex-gray-600)', border: isDefaultPreset ? 'none' : '1px solid var(--dex-gray-300)', borderRadius: 6 }}>
              {isDefaultPreset ? '✓ ' : ''}{isDe ? 'Standard' : 'Default'}
            </button>
          </div>
          {previewSrc && (
            <div style={{ width: PREV_W, flexShrink: 0, border: '1px solid var(--dex-gray-200)', borderRadius: 4, overflow: 'hidden', background: '#fff' }}>
              <div style={{ textAlign: 'center', padding: `${Math.round(headerImageLayout.paddingV * sc)}px ${Math.round(headerImageLayout.paddingH * sc)}px` }}>
                <img src={previewSrc} alt="" style={{ display: 'inline-block', width: '100%', maxWidth: Math.max(20, Math.round(headerImageLayout.width * sc)), height: 'auto' }} />
              </div>
              <div style={{ borderTop: '2px solid var(--dex-green, #86bc25)' }} />
              <div style={{ fontSize: '0.6rem', color: 'var(--dex-gray-400)', textAlign: 'center', padding: '2px 0' }}>{isDe ? 'So groß im Mail-Kopf (verkleinert)' : 'Size in the mail header (scaled)'}</div>
            </div>
          )}
        </div>
        {/* v28.29: sagt, WOHER das gezeigte Bild kommt (eigenes / vom Hauptevent
            geerbt / Standardlogo). Vorher zeigte die Vorschau kommentarlos das
            Event-Foto, obwohl gespeichert etwas anderes wurde. */}
        {note && (
          <div style={{ marginTop: 8, fontSize: '0.72rem', color: 'var(--dex-gray-600)', lineHeight: 1.45 }}>
            {note}
          </div>
        )}
      </div>
    );
  };
  const [dragFieldId, setDragFieldId] = React.useState<string | null>(null);
  // v18.55: Pro-Feld Ein-/Ausklapp-Status für Schritt 5 (Felder). Default =
  // eingeklappt (kompakte Karte: nur Nummer + Label + Typ + Pflicht + Aktionen);
  // Detail-Einstellungen (Hilfetext, Optionen, Bedingung, CC, EN-Variante …)
  // erst beim Aufklappen. Neu hinzugefügte Felder starten aufgeklappt.
  const [fieldExpandOverride, setFieldExpandOverride] = React.useState<Record<string, boolean>>({});
  const toggleFieldExpand = (id: string, current: boolean): void =>
    setFieldExpandOverride(prev => ({ ...prev, [id]: !current }));
  const [dragOverFieldId, setDragOverFieldId] = React.useState<string | null>(null);
  // v9.28: Reorder-Mode toggelt die Hoch/Runter-Pfeile pro Custom-Field.
  // Standardmäßig aus — sonst sieht das Feld-Listing zu unruhig aus.
  const [reorderMode, setReorderMode] = React.useState(false);
  // v9.28: Modal für neuen Quiz-Bereich (statt window.prompt)
  const [newSectionModalOpen, setNewSectionModalOpen] = React.useState(false);
  const [newSectionName, setNewSectionName] = React.useState('');
  const [newSectionError, setNewSectionError] = React.useState('');
  const [agenda, setAgenda] = React.useState<AgendaItem[]>(
    editEvent && editEvent.agenda ? [...editEvent.agenda] : []
  );
  const [transferTimes, setTransferTimes] = React.useState<Array<{id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string}>>(
    editEvent?.transferTimes?.map(t => ({...t, meetingPoint: t.meetingPoint || '', address: t.address || '', arrivalTime: t.arrivalTime || '', description: t.description || ''})) || []
  );
  const [documents, setDocuments] = React.useState<Array<{name: string; file?: File; url: string; size: number}>>(
    editEvent?.documents?.map(d => ({...d, size: d.size || 0})) || []
  );
  // Snapshot der beim Edit-Start vorhandenen Dokument-Namen, um beim Speichern
  // entfernte Attachments aus SharePoint löschen zu können.
  const [initialDocumentNames] = React.useState<string[]>(
    editEvent?.documents?.map(d => d.name) || []
  );
  const [quiz, setQuiz] = React.useState<Array<{id: string; question: string; options: string[]; correctIndices: number[]; imageBase64?: string; section?: string}>>(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    editEvent?.quiz?.map(q => ({...q, correctIndices: q.correctIndices || [(q as any).correctIndex || 0], imageBase64: (q as any).imageBase64, section: (q as any).section})) || []
  );
  const quizClusterSize = editEvent?.quizClusterSize || 1;
  // Sub-Events-Drafts im UI. Seit v6.4: Sub-Events sind eigene DEX_Events-Items.
  // Beim Edit laden wir die bestehenden Child-Events und mappen sie auf Drafts.
  // Beim Save werden Drafts mit `dbId` als updateEvent, ohne als createEvent geschrieben;
  // in der DB verbliebene Child-Events, die nicht mehr im Draft sind, werden gelöscht.
  const [subEvents, setSubEvents] = React.useState<SubEventDraft[]>(() => {
    if (!editEvent) return [];
    const kids = childEventsOf(editEvent.id);
    return kids.map(k => {
      // v11.57: Pro-Sub-Event Logo-Bilder aus EmailTemplateOverrides
      // (Piggyback-Pattern, gleich wie Top-Level-Event).
      // v14.4: zusätzlich die Mail-Text-Overrides (Anmeldung/Warteliste/
      // Abmeldung/Nachrücken) — vorher landeten Edits auf Sub-Event-Tabs
      // versehentlich beim Haupt-Event.
      let emailLogo = '';
      let outlookLogo = '';
      let subOverrides: Record<string, EmailOverrideEntry> = {};
      // v15.0: Inheritance-Flags aus dem Piggyback-JSON lesen. Wenn der
      // Flag nicht persistiert wurde (alte Events) fällt die App auf
      // datenbasierte Heuristik zurück (siehe weiter unten).
      let inheritFlagsRaw: { capacity?: boolean; fields?: boolean; location?: boolean } | undefined;
      try {
        const ov = JSON.parse(k.emailTemplateOverrides || '{}') as Record<string, unknown>;
        emailLogo = (ov?._eventLogo as string) || '';
        outlookLogo = readOutlookLogo(ov);
        inheritFlagsRaw = (ov?._inheritFlags as { capacity?: boolean; fields?: boolean; location?: boolean } | undefined);
        // Piggyback-Keys (mit Unterstrich-Prefix) rausstrippen, der Rest sind
        // die echten Mail-Template-Overrides pro TemplateType.
        const filtered: Record<string, EmailOverrideEntry> = {};
        for (const key of Object.keys(ov)) {
          if (key.startsWith('_')) continue;
          const val = ov[key] as Partial<EmailOverrideEntry> | undefined;
          if (val && (val.subject || val.heading || val.bodyHtml || val.headingColor || val.headingFontSize || val.subheading !== undefined || val.headingBold !== undefined || val.headingItalic !== undefined || val.subheadingColor || val.subheadingFontSize || val.subheadingBold !== undefined || val.subheadingItalic !== undefined)) {
            filtered[key] = {
              subject: val.subject || '',
              heading: val.heading || '',
              bodyHtml: val.bodyHtml || '',
              // v18.19/v18.22: Überschrift-Farbe/-Größe/-Stil + Subheading-
              // Formatierung mit-übernehmen.
              ...(val.subheading !== undefined ? { subheading: val.subheading } : {}),
              ...(val.headingColor ? { headingColor: val.headingColor } : {}),
              ...(val.headingFontSize ? { headingFontSize: val.headingFontSize } : {}),
              ...(val.headingBold !== undefined ? { headingBold: val.headingBold } : {}),
              ...(val.headingItalic !== undefined ? { headingItalic: val.headingItalic } : {}),
              ...(val.subheadingColor ? { subheadingColor: val.subheadingColor } : {}),
              ...(val.subheadingFontSize ? { subheadingFontSize: val.subheadingFontSize } : {}),
              ...(val.subheadingBold !== undefined ? { subheadingBold: val.subheadingBold } : {}),
              ...(val.subheadingItalic !== undefined ? { subheadingItalic: val.subheadingItalic } : {}),
            };
          }
        }
        subOverrides = filtered;
      } catch { /* */ }
      const parsedHeads = parseOutlookHeadings(k.outlookBody || '');
      // v15.0: Inheritance-Heuristik für Bestands-Events: wenn das
      // Piggyback-Flag fehlt UND das jeweilige Datenfeld nicht-leer ist,
      // gilt es als „eigener Wert" (nicht vom Hauptevent geerbt). Wenn
      // das Feld leer ist, default = übernehmen.
      const inheritCap = inheritFlagsRaw && typeof inheritFlagsRaw.capacity === 'boolean'
        ? inheritFlagsRaw.capacity
        : !(k.maxParticipants && k.maxParticipants > 0);
      const inheritFields = inheritFlagsRaw && typeof inheritFlagsRaw.fields === 'boolean'
        ? inheritFlagsRaw.fields
        : !((k.eventSpecificFields || []).length > 0);
      const inheritLoc = inheritFlagsRaw && typeof inheritFlagsRaw.location === 'boolean'
        ? inheritFlagsRaw.location
        : !(k.location && k.location.trim().length > 0);
      return {
      id: k.id,
      dbId: k.id,
      title: k.title,
      description: k.description,
      location: k.location,
      startDate: k.startDate,
      endDate: k.endDate,
      maxParticipants: k.maxParticipants || 0,
      registrationDeadline: k.registrationDeadline,
      mandatory: !!k.mandatoryRegistration, // v24.64: Pflicht-Sub-Event
      disableEmails: k.disableEmails,
      disableRegistrationEmail: k.disableRegistrationEmail,
      disableCancellationEmail: k.disableCancellationEmail,
      autoDeregisterOnDecline: k.autoDeregisterOnDecline,
      inactiveHandling: k.inactiveHandling,
      disableOutlook: k.disableOutlook,
      // v11.57: pro-Sub-Event Kommunikations-Felder laden
      emailLanguage: k.emailLanguage || (locale === 'de' ? 'DE' : 'EN'),
      emailLogoBase64: emailLogo,
      outlookLogoBase64: outlookLogo,
      emailTemplateOverrides: subOverrides,
      outlookBody: stripOutlookWrapper(k.outlookBody || ''),
      outlookHeading: parsedHeads.heading || k.title || '',
      outlookSubheading: parsedHeads.subheading && parsedHeads.subheading !== 'Event Details' ? parsedHeads.subheading : '',
      outlookSubject: k.outlookSubject || '',
      outlookStart: k.outlookStart || '',
      outlookEnd: k.outlookEnd || '',
      outlookLocation: k.outlookLocation || '',
      allDay: !!k.allDay,
      showAsFree: !!k.showAsFree,
      // v11.57: Snapshot der initialen Outlook-relevanten Felder
      initialOutlookEventId: k.outlookEventId || '',
      // v11.61: CalendarLink (iCalUId) als Outlook-Existenz-Indikator. Der
      // Flow schreibt OutlookEventId nicht — auf erfolgreichen Sub-Events
      // ist nur CalendarLink gefüllt.
      initialCalendarLink: k.calendarLink || '',
      initialTitle: k.title || '',
      initialAllDay: !!k.allDay,
      initialShowAsFree: !!k.showAsFree,
      initialStartDate: k.startDate || '',
      initialEndDate: k.endDate || '',
      initialOutlookBody: k.outlookBody || '',
      // v28.30: Kopfbild-Snapshot (gleicher Piggyback-Key wie beim Hauptevent).
      initialOutlookLogoBase64: ((): string => {
        if (!k.emailTemplateOverrides) return '';
        try { return readOutlookLogo(JSON.parse(k.emailTemplateOverrides)); } catch { return ''; }
      })(),
      customFields: (k.eventSpecificFields || []).map(f => ({
        id: f.id,
        label: f.label,
        type: f.type as CustomFieldInput['type'],
        required: !!f.required,
        options: f.options || [],
        // EventSpecificField hat kein 'visible'-Feld — default auf true.
        // Sichtbarkeit ist im Storage immer „shown" (default), nur im Wizard
        // kann der User Felder ausblenden.
        visible: true,
        externalLinks: f.externalLinks,
        multi: f.multi,
        helpText: f.helpText,
        helpTextStyle: f.helpTextStyle,
        showIf: f.showIf,
        // v29.20 (Audit A3): Dieser Mapper übernahm nur eine Teilmenge der
        // Feld-Eigenschaften — der nächste Save eines Klammer-Events schrieb
        // die Sub-Event-CustomFields dann OHNE den Rest zurück (die
        // v11.21-Drop-Klasse, hier im Lade-Pfad). „Felder vom Hauptevent
        // kopieren" und der Sub-Feld-Editor setzen all diese Properties.
        ...(f.onlyForGroup ? { onlyForGroup: f.onlyForGroup } : {}),
        ...(f.confirmLabel ? { confirmLabel: f.confirmLabel } : {}),
        ...(f.defaultValue ? { defaultValue: f.defaultValue } : {}),
        ...(f.optionCategories && f.optionCategories.length > 0 ? { optionCategories: [...f.optionCategories] } : {}),
        ...(f.prefilterLabel ? { prefilterLabel: f.prefilterLabel } : {}),
        ...(f.labelEn ? { labelEn: f.labelEn } : {}),
        ...(f.helpTextEn ? { helpTextEn: f.helpTextEn } : {}),
        ...(f.confirmLabelEn ? { confirmLabelEn: f.confirmLabelEn } : {}),
        ...(f.optionsEn && f.optionsEn.length > 0 ? { optionsEn: [...f.optionsEn] } : {}),
        ...(f.ccOnEmails ? { ccOnEmails: true } : {}),
        ...(f.notifyRoommate === false ? { notifyRoommate: false } : {}),
        ...(f.audienceOnly ? { audienceOnly: true } : {}),
        ...(f.withTime ? { withTime: true } : {}),
        ...(f.rangeStart ? { rangeStart: f.rangeStart } : {}),
        ...(f.rangeEnd ? { rangeEnd: f.rangeEnd } : {}),
        ...(typeof f.maxNights === 'number' && f.maxNights > 0 ? { maxNights: f.maxNights } : {}),
      })),
      // v15.3: pro-Sub-Event Felder aus dem Event-Datenmodell laden. Alle
      // Sub-Events haben jetzt eigene Adresse, Agenda, Transferzeiten,
      // Deadline, Standortfilter, Audience, Filter-Modus, Warteliste und
      // Anrede-Toggle — wie der Hauptevent.
      locationAddress: k.locationAddress ? {
        street: k.locationAddress.street || '',
        houseNo: k.locationAddress.houseNo || '',
        zip: k.locationAddress.zip || '',
        city: k.locationAddress.city || '',
      } : { street: '', houseNo: '', zip: '', city: '' },
      agenda: (k.agenda || []) as AgendaItem[],
      transferTimes: (k.transferTimes || []).map(tt => ({
        id: tt.id,
        location: tt.location || '',
        meetingPoint: tt.meetingPoint || '',
        address: tt.address || '',
        date: tt.date || '',
        departureTime: tt.departureTime || '',
        arrivalTime: tt.arrivalTime || '',
        description: tt.description || '',
      })),
      lastDeregisterDate: k.lastDeregisterDate || '',
      // Form-Felder für Standortfilter / Mailverteiler sind comma-separated
      // Strings, persistiert im Event aber als Arrays — siehe Top-Level-Mapping.
      locationFilter: (k.locationAudience || []).join(', '),
      audience: (k.audienceFilter || []).join(', '),
      filterMode: (k.filterMode === 'AND' ? 'AND' : 'OR') as 'AND' | 'OR',
      // v22.10: Ausschluss-Liste des Sub-Events laden (vorher nicht persistiert).
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      excludedUsers: ((k as any).excludedUsers || []) as string[],
      waitlistEnabled: typeof k.waitlistEnabled === 'boolean' ? k.waitlistEnabled : true,
      askSalutation: !!k.askSalutation,
      // v27.11: bestehendes Sub-Event-Bild (SP-URL) als Vorschau laden.
      imagePreview: k.imageUrl || '',
      // v15.0 (legacy): Inheritance-Flags werden seit v15.3 nicht mehr
      // ausgewertet. Bleiben in den geparsten Drafts, weil das Schema
      // sie noch erlaubt — Wirkung gleich Null.
      inheritLocationFromParent: inheritLoc,
      inheritCapacityFromParent: inheritCap,
      inheritCustomFieldsFromParent: inheritFields,
    };
    });
  });
  // v29.75: Solange der „Sichtbarkeit gilt für alle Sub-Events"-Haken gesetzt
  // ist, spiegelt jede Änderung an Standortfilter/Verteiler/Verknüpfung der
  // Klammer sofort in alle Sub-Event-Drafts — auch das Setzen des Hakens
  // selbst. Die Sub-Event-Sichtbarkeits-UI ist währenddessen gesperrt (sonst
  // würden dortige Eingaben beim nächsten Klammer-Edit stillschweigend
  // überschrieben). Unveränderte Drafts behalten ihre Referenz, damit der
  // v29.57-Skip („nichts geändert → nicht schreiben") nicht anschlägt.
  React.useEffect(() => {
    if (!visAllSubs) return;
    setSubEvents(prev => {
      let changed = false;
      const next = prev.map(s => {
        if ((s.locationFilter || '') === locationFilter && (s.audience || '') === audience && (s.filterMode || 'OR') === filterMode) return s;
        changed = true;
        return { ...s, locationFilter, audience, filterMode };
      });
      return changed ? next : prev;
    });
  }, [visAllSubs, locationFilter, audience, filterMode]);
  // v29.76: Rollierende Fristen materialisieren — je Sub-Event wird
  // startDate minus Abstand in registrationDeadline/lastDeregisterDate
  // geschrieben. subEvents steht bewusst MIT in den Deps: neu angeklickte
  // Kalender-Tage und geaenderte Termine bekommen die Frist sofort; die
  // Referenz-Stabilitaet (return prev bei 0 Aenderungen) beendet die
  // Kette nach einem Durchlauf und haelt den v29.57-Skip sauber.
  // ACHTUNG: subEventsOptIn ist hier NICHT referenzierbar (Deklaration erst
  // weiter unten — TDZ, die v29.71-Falle). Ein eigener Guard ist auch nicht
  // noetig: ohne Sub-Event-Drafts ist die Schleife leer.
  // v29.77: gilt fuer ALLE Sub-Event-Events, nicht mehr nur den Kalender.
  // v30.6: Die Regel ist nur noch eine VORBELEGUNG. Bei unveraenderter Regel
  // werden ausschliesslich LEERE Fristen gefuellt (neu angeklickte Tage) —
  // ein manuell ueberschriebener Wert im Sub-Reiter bleibt stehen, auch
  // ueber Speichern/Wieder-Oeffnen hinweg (die Abweichung lebt in den
  // materialisierten Spalten selbst, es braucht kein neues Flag). Aendert
  // der Organizer die REGEL, rechnet sie bewusst wieder ALLE Termine neu.
  // deadlineRuleKeyRef startet mit dem GELADENEN Regelstand, damit das
  // Wieder-Oeffnen des Wizards nicht als Regel-Aenderung zaehlt.
  const deadlineRuleKey = JSON.stringify({
    r: (regRuleEnabled && regRuleAmount > 0) ? [regRuleAmount, regRuleUnit] : null,
    c: (cancelRuleEnabled && cancelRuleAmount > 0 && userCancelAllowed) ? [cancelRuleAmount, cancelRuleUnit, cancelRuleAfter] : null,
  });
  const deadlineRuleKeyRef = React.useRef(deadlineRuleKey);
  React.useEffect(() => {
    const force = deadlineRuleKeyRef.current !== deadlineRuleKey;
    deadlineRuleKeyRef.current = deadlineRuleKey;
    if (!regRuleEnabled && !cancelRuleEnabled) return;
    const near = (a: string, b: string): boolean => {
      const ta = new Date(a || '').getTime(); const tb = new Date(b || '').getTime();
      return isFinite(ta) && isFinite(tb) && Math.abs(ta - tb) < 60000;
    };
    setSubEvents(prev => {
      let changed = false;
      const next = prev.map(s => {
        if (!s.startDate) return s;
        const patch: Partial<SubEventDraft> = {};
        if (regRuleEnabled && regRuleAmount > 0) {
          const iso = rollingDeadlineIso(s.startDate, regRuleAmount, regRuleUnit);
          const cur = s.registrationDeadline || '';
          if (iso && (force ? !near(cur, iso) : !cur)) patch.registrationDeadline = iso;
        }
        if (cancelRuleEnabled && cancelRuleAmount > 0 && userCancelAllowed) {
          const iso = rollingDeadlineIso(s.startDate, cancelRuleAmount, cancelRuleUnit, cancelRuleAfter);
          const cur = s.lastDeregisterDate || '';
          if (iso && (force ? !near(cur, iso) : !cur)) patch.lastDeregisterDate = iso;
        }
        if (Object.keys(patch).length === 0) return s;
        changed = true;
        return { ...s, ...patch };
      });
      return changed ? next : prev;
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [deadlineRuleKey, regRuleEnabled, regRuleAmount, regRuleUnit, cancelRuleEnabled, cancelRuleAmount, cancelRuleUnit, cancelRuleAfter, userCancelAllowed, subEvents]);
  // v11.57: aktiv ausgewählter Tab in Step 6 (Kommunikation, v11.80 Renumbering). 0 = Haupt-Event,
  // N>0 = subEvents[N-1]. Beim Tab-Wechsel werden die Step-5-Felder zwischen
  // dem Top-Level-State und der jeweiligen Sub-Event-Slice gespiegelt — siehe
  // switchCommTab-Helper weiter unten.
  const [activeCommTabIdx, setActiveCommTabIdx] = React.useState<number>(0);
  // v15.0: pro-Event-Tabs in Schritt 3 (Ort), Schritt 4 (Kapazität) und
  // Schritt 6 (Felder). 0 = Haupt-Event, N>0 = subEvents[N-1]. Im
  // Gegensatz zu Step 6 (Kommunikation) gibt es hier KEIN Mirror-Pattern,
  // weil die per-Tab-Werte direkt im subEvents[]-State (location,
  // maxParticipants, customFields) bzw. in den Top-Level-States stehen —
  // jedes Eingabefeld liest/schreibt direkt aus seinem Zielort.
  const [activeLocationTabIdx, setActiveLocationTabIdx] = React.useState<number>(0);
  const [activeCapacityTabIdx, setActiveCapacityTabIdx] = React.useState<number>(0);
  const [activeFieldsTabIdx, setActiveFieldsTabIdx] = React.useState<number>(0);
  // v24.62: Manueller „Outlook-Termin jetzt aktualisieren"-Trigger (Schritt 6,
  // Sektionen Outlook-Bild + Outlook-Text). Schickt für das Event/Sub-Event des
  // aktiven Kommunikations-Tabs ein Update an die Kalender der Teilnehmer — für
  // den Fall, dass der Outlook-Termin noch den alten Stand zeigt. Nutzt den
  // zuletzt GESPEICHERTEN Stand (der Flow liest das Event aus der Verwaltung).
  const [outlookUpdateBusy, setOutlookUpdateBusy] = React.useState(false);
  // v28.28: Was wurde zuletzt angestoßen? Ohne diese Rückmeldung blieb im
  // Wizard nur der unveränderte Knopf stehen — Organizer lasen das als
  // „hat nicht funktioniert" und klickten wieder und wieder.
  const [outlookUpdateDone, setOutlookUpdateDone] = React.useState<string>('');
  /**
   * v28.28: Sub-Events haben EIGENE Outlook-Termine. Der Knopf aktualisiert
   * bisher ausschließlich den Termin des GERADE GEÖFFNETEN Tabs — steht man auf
   * dem Hauptevent, bleiben die Sub-Event-Termine unangetastet. Das war die
   * Ursache für „Update bei Sub-Events funktioniert nicht sauber": Es wurde
   * schlicht ein anderer Termin aktualisiert als erwartet. Deshalb sagt der
   * Knopf jetzt, für WELCHEN Termin er gilt, und es gibt einen zweiten für
   * „alle Termine dieses Events".
   */
  const outlookUpdateTargets = (): Array<{ id: string; title: string }> => {
    const out: Array<{ id: string; title: string }> = [];
    if (editEvent && editEvent.disableOutlook !== true && (editEvent.outlookEventId || editEvent.calendarLink)) {
      out.push({ id: editEvent.id, title: title || editEvent.title || '' });
    }
    for (const s of subEventsRef.current) {
      if (!s.dbId || s.disableOutlook) continue;
      if (!s.initialOutlookEventId && !s.initialCalendarLink) continue;
      out.push({ id: s.dbId, title: s.title || '' });
    }
    return out;
  };
  /**
   * v28.67: Gegenstück zu outlookUpdateTargets() — Events, für die ein
   * Outlook-Termin vorgesehen ist, aber KEINER existiert (weder OutlookEventId
   * noch CalendarLink). Bisher fielen die einfach aus der Zaehlung: Bei einem
   * Event mit Hauptevent + vier Sub-Events stand dann „Alle 2 Termine
   * aktualisieren", ohne zu sagen, dass drei Kalendereinträge schlicht
   * fehlen — das las sich wie ein Zaehlfehler der App statt wie der Befund,
   * der es ist. Häufigste Ursache war der v28.66-Bug (Sub-Event ohne Zeiten
   * -> „Create event (V4)" bricht ab). Hier wird der Befund benannt.
   */
  /**
   * v28.66/v28.69: Start-/Endzeit des Hauptevents als UTC-ISO. Fallback-Quelle
   * für Sub-Events ohne eigene Zeiten — sowohl beim Speichern als auch beim
   * nachträglichen Anlegen fehlender Termine.
   */
  const parentTimesIso = (): { start: string; end: string } => ({
    start: startDate ? berlinLocalToUtcIso(startDate) : '',
    end: endDate ? berlinLocalToUtcIso(endDate) : '',
  });
  /**
   * v28.69: dbIds von Sub-Events, deren Outlook-Termin beim nächsten Speichern
   * ERZWUNGEN neu angelegt werden soll. Der Flow DEX_CreateOutlookEvent
   * triggert nur auf NEUE DEX_Events-Zeilen — ein reines Update stoesst ihn
   * nie an. persistSubEventsForParent nimmt die IDs auf und läuft dafür in
   * den nicht-destruktiven Recreate-Pfad aus v11.69 (Zeile löschen, mit
   * existingSubsiteUrl neu anlegen; Anmeldungen, TeilnehmerIDs und Subsite
   * bleiben unberuehrt). Wird nach dem Durchlauf geleert.
   */
  const forceOutlookRecreateRef = React.useRef<Set<string>>(new Set<string>());
  const outlookMissingTargets = (): Array<{ id: string; title: string }> => {
    const out: Array<{ id: string; title: string }> = [];
    if (editEvent && editEvent.disableOutlook !== true && !editEvent.outlookEventId && !editEvent.calendarLink) {
      out.push({ id: editEvent.id, title: title || editEvent.title || '' });
    }
    for (const s of subEventsRef.current) {
      if (!s.dbId || s.disableOutlook) continue;
      if (s.initialOutlookEventId || s.initialCalendarLink) continue;
      out.push({ id: s.dbId, title: s.title || '' });
    }
    return out;
  };
  /**
   * v28.69: Fehlende Sub-Event-Termine direkt anlegen. Bisher stand im Kasten
   * nur die Anleitung „Haken aus- und wieder einschalten und erneut speichern"
   * — das ist der v11.69-Recreate-Pfad, den der Organizer von Hand ausloesen
   * musste. Jetzt macht der Knopf genau das: fehlende Zeiten aus dem
   * Hauptevent übernehmen (v28.66), Recreate erzwingen, speichern.
   * Das Hauptevent selbst wird bewusst NICHT neu angelegt — seine Item-Id
   * steht in ParentEventId aller Sub-Events, ein Recreate würde die Kinder
   * zu Waisen machen.
   */
  const createMissingOutlookAppointments = async (): Promise<void> => {
    const mainId = editEvent?.id || '';
    const missing = outlookMissingTargets().filter(m => m.id && m.id !== mainId);
    if (missing.length === 0) return;
    const pt = parentTimesIso();
    const fmt = (iso: string): string => {
      if (!iso) return '—';
      const d = new Date(iso);
      return isNaN(d.getTime()) ? '—' : d.toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    };
    const plan = missing.map(m => {
      const d = subEventsRef.current.find(s => s.dbId === m.id);
      const start = (d && d.startDate) || pt.start || '';
      const end = (d && d.endDate) || (d && d.startDate) || pt.end || pt.start || '';
      const inherited = !(d && d.startDate);
      return { ...m, start, end, inherited };
    });
    if (plan.some(p => !p.start)) {
      showAlert(isDe
        ? 'Das Hauptevent hat keine Startzeit — ohne die lässt sich kein Termin anlegen. Bitte zuerst in Schritt 1 die Zeiten setzen und speichern.'
        : 'The main event has no start time — no appointment can be created without it. Please set the times in step 1 first and save.', { variant: 'error' });
      return;
    }
    const list = plan.map(p => `• ${p.title || '?'}: ${fmt(p.start)} – ${fmt(p.end)}${p.inherited ? (isDe ? '  (Zeiten des Hauptevents)' : '  (main event times)') : ''}`).join('\n');
    const ok = await confirmDialog(
      isDe
        ? `Für diese ${plan.length} ${plan.length === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} wird jetzt ein Outlook-Termin angelegt:\n\n${list}\n\nSub-Events ohne eigene Zeiten übernehmen die Zeiten des Hauptevents. Wenn du genauere Zeiten willst, brich hier ab, trage sie oben ein und klicke danach erneut.\n\nAnmeldungen, Teilnehmer-IDs und Teilnehmerlisten bleiben dabei unverändert.`
        : `An Outlook appointment will now be created for these ${plan.length} sub-event(s):\n\n${list}\n\nSub-events without their own times inherit the main event's times. If you want more precise times, cancel here, enter them above and click again.\n\nRegistrations, attendee IDs and attendee lists stay untouched.`,
      { confirmLabel: isDe ? 'Termine anlegen' : 'Create appointments' },
    );
    if (!ok) return;
    plan.forEach(p => forceOutlookRecreateRef.current.add(p.id));
    // v29.20 (Audit): wie attemptSubmit VOR dem Save flushen — dieser Pfad
    // rief handleSubmit direkt, und die Kommunikations-Eingaben des gerade
    // offenen Reiters lagen dann noch nicht im Draft (CLAUDE.md-Falle):
    // Der frisch angelegte Outlook-Termin trug den Text von VOR der letzten
    // Bearbeitung. Ebenso die pendingOutlook*-Reste eines früheren
    // Modal-Durchlaufs zurücksetzen, die hier sonst ungefragt nachwirkten.
    flushActiveCommTabToState();
    pendingOutlookDirtyWriteRef.current = null;
    pendingOutlookDirtyWriteRefs.current = {};
    pendingOutlookUpdateForTopRef.current = false;
    pendingOutlookUpdateForSubEventsRef.current = [];
    pendingOutlookRecreateForSubEventsRef.current = [];
    setOutlookUpdateBusy(true);
    try {
      await handleSubmit();
    } finally {
      setOutlookUpdateBusy(false);
      forceOutlookRecreateRef.current.clear();
    }
  };
  const triggerOutlookUpdateNow = async (): Promise<void> => {
    let targetDbId = '';
    let targetTitle = title;
    let hasAppointment = false;
    if (activeCommTabIdx > 0) {
      const sub = subEvents[activeCommTabIdx - 1];
      if (sub) {
        targetDbId = sub.dbId || '';
        targetTitle = sub.title || title;
        hasAppointment = !!(sub.initialOutlookEventId || sub.initialCalendarLink);
      }
    } else {
      targetDbId = editEvent?.id || '';
      hasAppointment = !!(editEvent?.outlookEventId || editEvent?.calendarLink);
    }
    if (!editEvent || !targetDbId) {
      showAlert(isDe ? 'Den Outlook-Termin gibt es erst, nachdem das Event gespeichert wurde.' : 'The Outlook appointment only exists after the event has been saved.', { variant: 'info' });
      return;
    }
    if (disableOutlook) {
      showAlert(isDe ? 'Für diesen Tab ist der Outlook-Termin deaktiviert (Schalter weiter oben in Schritt 6).' : 'The Outlook appointment is disabled for this tab (toggle further up in step 6).', { variant: 'info' });
      return;
    }
    if (!hasAppointment) {
      showAlert(isDe ? 'Für dieses Event wurde noch kein Outlook-Termin angelegt — er entsteht beim Speichern.' : 'No Outlook appointment has been created for this event yet — it is created on save.', { variant: 'info' });
      return;
    }
    const ok = await confirmDialog(
      isDe
        ? `Der Outlook-Termin von „${targetTitle}" wird bei allen Teilnehmern mit dem zuletzt GESPEICHERTEN Stand aktualisiert. Falls du gerade etwas geändert hast, speichere bitte zuerst und klicke dann erneut hier.\n\nHinweis: Sub-Events haben eigene Termine — die aktualisierst du im jeweiligen Tab oder über „Alle Termine aktualisieren".`
        : `The Outlook appointment of „${targetTitle}" will be updated for all attendees with the last SAVED state. If you just changed something, please save first and then click here again.\n\nNote: sub-events have their own appointments — update them in their tab or via „Update all appointments".`,
      { confirmLabel: isDe ? 'Jetzt aktualisieren' : 'Update now' },
    );
    if (!ok) return;
    setOutlookUpdateBusy(true);
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      const svc = new EventService(ctx);
      await svc.queueOutlookEvent('', targetDbId, targetTitle, 'UpdateEvent');
      setOutlookUpdateDone(isDe
        ? `Angestoßen für „${targetTitle}" (${new Date().toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })} Uhr). Die Kalender der Teilnehmer aktualisieren sich in Kürze.`
        : `Triggered for „${targetTitle}" (${new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })}). Attendees' calendars will refresh shortly.`);
      showAlert(isDe ? 'Outlook-Aktualisierung wurde angestoßen — die Kalender der Teilnehmer aktualisieren sich in Kürze.' : 'Outlook update triggered — attendees will see the refreshed appointment shortly.', { variant: 'success' });
    } catch {
      showAlert(isDe ? 'Aktualisierung fehlgeschlagen — bitte erneut versuchen.' : 'Update failed — please try again.', { variant: 'error' });
    } finally {
      setOutlookUpdateBusy(false);
    }
  };
  /** v28.28: Haupt-Termin UND alle Sub-Event-Termine in einem Rutsch. */
  const triggerOutlookUpdateAll = async (): Promise<void> => {
    const targets = outlookUpdateTargets();
    if (targets.length === 0) return;
    const list = targets.map(t => `• ${t.title || '?'}`).join('\n');
    const ok = await confirmDialog(
      isDe
        ? `Alle ${targets.length} Outlook-Termine dieses Events mit dem zuletzt GESPEICHERTEN Stand aktualisieren?\n\n${list}\n\nJede/r Teilnehmer/in bekommt pro Termin, für den sie/er angemeldet ist, eine „Aktualisierter Termin"-Benachrichtigung.`
        : `Update all ${targets.length} Outlook appointments of this event with the last SAVED state?\n\n${list}\n\nEach attendee receives an „updated meeting" notification per appointment they are registered for.`,
      { confirmLabel: isDe ? 'Alle aktualisieren' : 'Update all' },
    );
    if (!ok) return;
    setOutlookUpdateBusy(true);
    let done = 0;
    let failed = 0;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      const svc = new EventService(ctx);
      for (const t of targets) {
        try { await svc.queueOutlookEvent('', t.id, t.title, 'UpdateEvent'); done += 1; }
        catch { failed += 1; }
      }
    } finally {
      setOutlookUpdateBusy(false);
    }
    const stamp = new Date().toLocaleTimeString(isDe ? 'de-DE' : 'en-GB', { hour: '2-digit', minute: '2-digit' });
    setOutlookUpdateDone(isDe
      ? `${done} Termin(e) angestoßen${failed > 0 ? `, ${failed} fehlgeschlagen` : ''} (${stamp} Uhr). Die Kalender der Teilnehmer aktualisieren sich in Kürze.`
      : `${done} appointment(s) triggered${failed > 0 ? `, ${failed} failed` : ''} (${stamp}). Attendees' calendars will refresh shortly.`);
    showAlert(
      isDe ? `${done} Outlook-Termin(e) angestoßen${failed > 0 ? `, ${failed} fehlgeschlagen` : ''}.` : `${done} Outlook appointment(s) triggered${failed > 0 ? `, ${failed} failed` : ''}.`,
      { variant: failed > 0 ? 'error' : 'success' },
    );
  };
  /**
   * v28.28: Kompakte Schalter-Zeile für „Was soll automatisch verschickt
   * werden?". Vorher stand unter jedem Haken ein drei- bis vierzeiliger
   * Fließtext — auf dem Schirm eine Textwüste, in der man die eigentlichen
   * Schalter kaum noch fand. Jetzt: fette Bezeichnung, EIN kurzer Satz, alle
   * Details im Info-Tooltip daneben.
   */
  const commToggleRow = (opts: {
    checked: boolean;
    onChange: (v: boolean) => void;
    label: string;
    short: string;
    info: React.ReactNode;
    accent?: string;
  }): React.ReactElement => (
    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', padding: '6px 0' }}>
      <input
        type="checkbox"
        checked={opts.checked}
        onChange={e => opts.onChange(e.target.checked)}
        style={{ width: 18, height: 18, cursor: 'pointer', marginTop: 2, flexShrink: 0, accentColor: opts.accent || 'var(--dex-green, #86bc25)' }}
      />
      <span style={{ fontSize: '0.88rem', minWidth: 0 }}>
        <strong>{opts.label}</strong>
        <InfoTooltip text={opts.info} />
        <span style={{ display: 'block', fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.45, marginTop: 1 }}>
          {opts.short}
        </span>
      </span>
    </label>
  );

  // Wiederverwendbarer Button-Block für die Outlook-Sektionen (Bild + Text).
  const renderOutlookUpdateButton = (): React.ReactNode => {
    if (!editEvent) return null; // nur beim Bearbeiten sinnvoll (Neu-Event hat noch keinen Termin)
    // v28.28: Der Kasten war orange umrandet und wurde dadurch als Warnung
    // („da steht noch was aus") gelesen — obwohl er nur ein dauerhaft
    // verfügbares Werkzeug ist und nach dem Klick natürlich stehen bleibt.
    // Jetzt neutral, mit Ziel-Angabe und sichtbarer Erfolgsmeldung.
    const tabTitle = activeCommTabIdx > 0
      ? (subEvents[activeCommTabIdx - 1]?.title || (childTermSingular || 'Sub-Event'))
      : (title || editEvent.title || (isDe ? 'Hauptevent' : 'main event'));
    const allTargets = outlookUpdateTargets();
    const showAll = allTargets.length > 1;
    // v28.67: fehlende Termine benennen (s. outlookMissingTargets).
    const missingTargets = outlookMissingTargets();
    const totalTargets = allTargets.length + missingTargets.length;
    // v28.69: nachanlegbar sind nur Sub-Events — das Hauptevent nicht, seine
    // Item-Id steht in ParentEventId aller Kinder (s. createMissingOutlookAppointments).
    const missingSubIds = missingTargets.filter(m => m.id && m.id !== (editEvent?.id || ''));
    return (
      <div style={{ marginTop: 14, padding: 12, borderRadius: 8, background: 'var(--dex-gray-50, #f8f9fa)', border: '1px solid var(--dex-gray-200)' }}>
        <div style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-800)', marginBottom: 6 }}>
          {isDe ? 'Outlook-Termin manuell nachschicken (optional)' : 'Re-send the Outlook appointment manually (optional)'}
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button
            type="button"
            className="btn btn-secondary"
            disabled={outlookUpdateBusy}
            onClick={() => { void triggerOutlookUpdateNow(); }}
            style={{ fontSize: '0.82rem', padding: '7px 14px' }}
          >
            {outlookUpdateBusy
              ? (isDe ? 'Wird aktualisiert…' : 'Updating…')
              : (isDe ? `Termin von „${tabTitle}" aktualisieren` : `Update appointment of „${tabTitle}"`)}
          </button>
          {showAll && (
            <button
              type="button"
              className="btn btn-secondary"
              disabled={outlookUpdateBusy}
              onClick={() => { void triggerOutlookUpdateAll(); }}
              style={{ fontSize: '0.82rem', padding: '7px 14px' }}
            >
              {isDe
                ? `Alle ${allTargets.length} Termine aktualisieren`
                : `Update all ${allTargets.length} appointments`}
            </button>
          )}
        </div>
        {missingTargets.length > 0 && (
          <div style={{
            marginTop: 8, padding: '8px 10px', borderRadius: 6, fontSize: '0.76rem', lineHeight: 1.5,
            background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
          }}>
            {isDe
              ? <>Für <strong>{missingTargets.length} von {totalTargets}</strong> Terminen dieses Events gibt es noch <strong>keinen</strong> Kalendereintrag — deshalb steht oben nur „{allTargets.length}“: {missingTargets.map(m => m.title || '?').join(', ')}. Häufigste Ursache: das {childTermSingular || 'Sub-Event'} wurde ohne Start-/Endzeit gespeichert, dann kann kein Termin erzeugt werden. {missingSubIds.length > 0 ? <>Der Knopf unten legt die fehlenden Termine jetzt an — Sub-Events ohne eigene Zeiten übernehmen dabei die Zeiten des Hauptevents. Anmeldungen und Teilnehmerlisten bleiben unverändert.</> : <>Für das Hauptevent selbst lässt sich das hier nicht nachholen — bitte beim Support melden.</>}</>
              : <>There is <strong>no</strong> calendar entry yet for <strong>{missingTargets.length} of {totalTargets}</strong> appointments of this event — that is why it says „{allTargets.length}“ above: {missingTargets.map(m => m.title || '?').join(', ')}. Most common cause: the sub-event was saved without a start/end time, so no appointment can be created. {missingSubIds.length > 0 ? <>The button below creates the missing appointments now — sub-events without their own times inherit the main event&apos;s times. Registrations and attendee lists stay untouched.</> : <>This cannot be repaired here for the main event itself — please contact support.</>}</>}
            {missingSubIds.length > 0 && (
              <div style={{ marginTop: 8 }}>
                <button
                  type="button"
                  className="btn btn-primary"
                  disabled={outlookUpdateBusy}
                  onClick={() => { void createMissingOutlookAppointments(); }}
                  style={{ fontSize: '0.82rem', padding: '7px 14px' }}
                >
                  {outlookUpdateBusy
                    ? (isDe ? 'Wird angelegt…' : 'Creating…')
                    : (isDe ? `${missingSubIds.length} fehlende Termine jetzt anlegen` : `Create ${missingSubIds.length} missing appointments now`)}
                </button>
              </div>
            )}
          </div>
        )}
        {outlookUpdateDone && (
          <div style={{
            marginTop: 8, padding: '6px 10px', borderRadius: 6, fontSize: '0.76rem', fontWeight: 600,
            background: '#f1f7e8', border: '1px solid var(--dex-green, #86bc25)', color: 'var(--dex-green-dark, #4a7c1f)',
          }}>
            ✓ {outlookUpdateDone}
          </div>
        )}
        <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)', marginTop: 8, lineHeight: 1.5 }}>
          {isDe
            ? <>Nur nötig, wenn der Kalendereintrag der Teilnehmer noch veraltet ist — der Termin wird dann mit dem zuletzt <strong>gespeicherten</strong> Stand neu verschickt. Erst speichern, dann klicken. <strong>Wichtig:</strong> {childTermPlural || 'Sub-Events'} haben eigene Termine; dieser Knopf betrifft nur „{tabTitle}“{showAll ? ' — für alle auf einmal den zweiten Knopf nutzen' : ''}. Der Kasten bleibt dauerhaft stehen, er ist keine Fehlermeldung.</>
            : <>Only needed if the attendees’ calendar entry is still outdated — the appointment is re-sent with the last <strong>saved</strong> state. Save first, then click. <strong>Important:</strong> sub-events have their own appointments; this button only affects „{tabTitle}“{showAll ? ' — use the second button for all at once' : ''}. This box is always here, it is not an error message.</>}
        </div>
      </div>
    );
  };
  // v11.60: synchroner Spiegel von subEvents für Save/Detect-Pfade. React-
  // State-Updates sind async — wenn flushActiveCommTabToState() per
  // setSubEvents(prev=>...) den aktiven Tab in die jeweilige Slice
  // zurückschreibt, sieht das direkt danach laufende
  // detectOutlookRelevantChanges() (und auch persistSubEventsForParent
  // unter handleSubmit) noch die alte Array aus dem Closure. Ergebnis:
  // Modal kommt nicht, und die Sub-Event-Änderung wird beim Schreiben
  // wieder mit dem Original überschrieben. Der Ref hält die jeweils
  // aktuellste Array synchron — alle Code-Pfade, die nach einem Flush
  // lesen, gehen über `subEventsRef.current`.
  const subEventsRef = React.useRef<typeof subEvents>(subEvents);
  React.useEffect(() => { subEventsRef.current = subEvents; }, [subEvents]);
  // v22.54: Im Klammer-Modus ist die Anmeldefrist der Klammer nicht buchbar/
  // wirkungslos — die effektive Frist ergibt sich aus dem am längsten laufenden
  // Sub-Event (späteste Sub-Event-Anmeldefrist). Wird im ausgegrauten
  // Anmeldefrist-Feld der Klammer rein anzeigend gebunden.
  const effectiveKlammerDeadline = React.useMemo<string>(() => {
    const times = subEvents
      .map(s => s.registrationDeadline)
      .filter((d): d is string => !!d)
      .map(d => new Date(d).getTime())
      .filter(n => Number.isFinite(n));
    if (!times.length) return '';
    const max = new Date(Math.max(...times));
    return `${max.getFullYear()}-${String(max.getMonth() + 1).padStart(2, '0')}-${String(max.getDate()).padStart(2, '0')}T${String(max.getHours()).padStart(2, '0')}:${String(max.getMinutes()).padStart(2, '0')}`;
  }, [subEvents]);
  // Snapshot der beim Edit-Start vorhandenen Sub-Event-DB-IDs, um beim Save
  // entfernte Sub-Events zu löschen.
  const [initialSubEventDbIds] = React.useState<string[]>(() => {
    if (!editEvent) return [];
    return childEventsOf(editEvent.id).map(k => k.id);
  });
  // Snapshot der initialen Outlook-Metadaten pro Sub-Event (DisableOutlook +
  // OutlookEventId + SubsiteUrl). Wird beim Save gebraucht, um zu erkennen, ob
  // ein Sub-Event nachträglich von „Outlook deaktiviert" auf „Outlook
  // aktiviert" gedreht wurde — der DEX_CreateOutlookEvent-Flow lauscht nur
  // auf NEUE DEX_Events-Items (GetOnNewItems-Trigger), deshalb muss das
  // betroffene Sub-Event in diesem Fall gelöscht und neu angelegt werden,
  // damit überhaupt ein Outlook-Termin entsteht.
  const [initialSubEventOutlookMeta] = React.useState<Record<string, { disableOutlook: boolean; outlookEventId: string; subsiteUrl: string; registrationListName: string }>>(() => {
    if (!editEvent) return {};
    const acc: Record<string, { disableOutlook: boolean; outlookEventId: string; subsiteUrl: string; registrationListName: string }> = {};
    for (const k of childEventsOf(editEvent.id)) {
      acc[k.id] = {
        disableOutlook: !!k.disableOutlook,
        outlookEventId: k.outlookEventId || '',
        subsiteUrl: k.subsiteUrl || '',
        // v11.69: Subsite-Events nutzen immer die Standard-Teilnehmerliste
        // "Teilnehmer" (siehe REG_LIST_NAME in EventService). Wird beim
        // Recreate-Pfad an `createEvent({ existingRegistrationListName })`
        // mitgegeben, damit der Reuse-Branch in createEvent greift.
        registrationListName: 'Teilnehmer',
      };
    }
    return acc;
  });
  // v11.57: Snapshot der initialen Outlook-relevanten Felder des Top-Level-
  // Events (Title, Start, End, OutlookBody). Wird beim Save mit den aktuellen
  // Werten verglichen — Änderung löst das Update-Confirm-Modal aus.
  // Im Ref, weil wir das einmal beim Mount fixieren und nicht bei Re-Renders
  // neu setzen wollen.
  const initialOutlookSnapshot = React.useRef<{ title: string; startDate: string; endDate: string; outlookBody: string; outlookLocation: string; outlookSubject: string; outlookStart: string; outlookEnd: string; organizers: string; outlookLogo: string; allDay: boolean; showAsFree: boolean }>({
    showAsFree: !!(editEvent && editEvent.showAsFree), // v29.54
    // v29.52: „Ganztägig" gehört in den Snapshot. Der Haken ändert weder Titel
    // noch Start/Ende (die stehen ohnehin auf 00:00/23:59) — ohne diesen
    // Vergleich bliebe er für den Detektor unsichtbar, es würde KEIN
    // UpdateEvent gequeued, und der bereits verschickte Outlook-Termin bliebe
    // für immer ein Zeitblock. Der Flow käme nie zum Zug.
    allDay: !!(editEvent && editEvent.allDay),
    // v28.30: Kopfbild des Outlook-Termins in den Snapshot. Ein Bildwechsel
    // ändert weder den rohen Termin-Text noch das Layout — ohne diesen
    // Vergleich blieb er für den Save-Detektor unsichtbar.
    outlookLogo: ((): string => {
      if (!editEvent?.emailTemplateOverrides) return '';
      try { return readOutlookLogo(JSON.parse(editEvent.emailTemplateOverrides)); } catch { return ''; }
    })(),
    // v22.48: Organizer-Namen in den Snapshot — eine Organizer-Änderung ändert
    // den Outlook-Standardtext („wendet euch bitte an …") und soll daher das
    // Outlook-Update-Modal öffnen.
    organizers: (editEvent?.organizers || []).join(';'),
    title: editEvent?.title || '',
    startDate: editEvent?.startDate || '',
    endDate: editEvent?.endDate || '',
    outlookBody: editEvent?.outlookBody || '',
    // v18.44: abweichendes Outlook-Datum in den Snapshot — eine Override-Änderung
    // soll das Update-Modal öffnen.
    outlookStart: editEvent?.outlookStart || '',
    outlookEnd: editEvent?.outlookEnd || '',
    // v18.34/v18.40: effektiver Ort in den Snapshot (gespeicherte Override ODER
    // Auto). Eine reine Ort-Änderung soll das Outlook-Update-Modal öffnen.
    outlookLocation: editEvent?.outlookLocation || buildOutlookLocation(editEvent?.location, editEvent?.locationAddress),
    // v18.42: Betreff in den Snapshot — eine reine Betreff-Änderung soll das
    // Outlook-Update-Modal ebenfalls öffnen.
    outlookSubject: editEvent?.outlookSubject || '',
  });
  // v11.57: Update-Confirm-Modal-State. Beim Save mit Outlook-relevanten
  // Änderungen öffnen wir das Modal und warten auf die Entscheidung des
  // Organizers. v11.63: Statt einem globalen "Outlook-Update senden ja/nein"
  // listet das Modal jetzt jedes geänderte Event einzeln (Hauptevent +
  // betroffene Sub-Events) und der Organizer setzt pro Event einen Haken.
  const [outlookConfirmOpen, setOutlookConfirmOpen] = React.useState(false);
  // v11.63: Snapshot der Detect-Items zum Modal-Open-Zeitpunkt. Jeder Eintrag
  // beschreibt ein Event (Hauptevent oder Sub-Event) mit Outlook-relevanten
  // Änderungen — Title, Start, End oder OutlookBody — und listet, welche
  // Felder sich geändert haben (für die Anzeige als Sub-Text pro Item).
  const [outlookConfirmItems, setOutlookConfirmItems] = React.useState<OutlookConfirmItem[]>([]);
  // v11.63: Pro Event-ID, ob die Checkbox im Modal angehakt ist.
  // true = UpdateEvent in Queue + OutlookDirty=false setzen.
  // false (oder nicht im Map) = kein UpdateEvent, OutlookDirty=true setzen.
  const [outlookConfirmChecks, setOutlookConfirmChecks] = React.useState<Record<string, boolean>>({});
  // v11.63: Top-Level-Outlook-Update-Entscheidung. true = nach erfolgreichem
  // updateEvent ein DEX_Outlook 'UpdateEvent' in die Queue schreiben.
  const pendingOutlookUpdateForTopRef = React.useRef<boolean>(false);
  // Sub-Event-IDs, für die ein DEX_Outlook 'UpdateEvent' angefordert wurde.
  const pendingOutlookUpdateForSubEventsRef = React.useRef<string[]>([]);
  // v11.69: Sub-Event-IDs, für die ein *Recreate* des DEX_Events-Items
  // angefordert wurde (Outlook-Termin nachträglich anlegen ohne Teilnehmer-
  // Verlust). Werden in `persistSubEventsForParent` aufgegriffen: das alte
  // DEX_Events-Item wird per `deleteEventItemOnly` entfernt (Subsite +
  // Teilnehmerliste bleiben unangetastet), dann wird per
  // `createEvent({ existingSubsiteUrl, existingRegistrationListName })` ein
  // neues DEX_Events-Item angelegt, das die alte Subsite wiederverwendet.
  // Der `DEX_CreateOutlookEvent`-Flow triggert auf das neue Item und legt den
  // Outlook-Termin an.
  const pendingOutlookRecreateForSubEventsRef = React.useRef<string[]>([]);
  // v11.63: Pro Event-ID der gewünschte OutlookDirty-Wert.
  // Nur Eventd-IDs, die im Modal waren, werden hier gesetzt — alle anderen
  // bleiben unberührt (kein OutlookDirty-Patch).
  const pendingOutlookDirtyWriteRefs = React.useRef<Record<string, boolean>>({});
  // v11.57 (kompatibel): Schreibwert für OutlookDirty für das Top-Level-
  // Event im nächsten updateEvent-Call. null = nicht setzen, false/true =
  // setzen. Wird aus pendingOutlookDirtyWriteRefs[topId] abgeleitet.
  const pendingOutlookDirtyWriteRef = React.useRef<boolean | null>(null);
  // Bereiche, die per "+ Bereich"-Button angelegt aber noch nicht mit einer
  // Frage belegt wurden. Sobald eine Frage per Drag&Drop reinkommt, ergibt sich
  // der Section-Name aus dem question.section-Feld selbst — pendingSections
  // hält nur die noch leeren Zwischen-Buckets.
  const [pendingSections, setPendingSections] = React.useState<string[]>([]);
  const [draggedQuestionId, setDraggedQuestionId] = React.useState<string | null>(null);
  // v6.35: Handbuch-Previews können einen bestimmten Wizard-Schritt gezielt
  // zeigen, indem sie vor dem Mount `window.__dexPreviewInitialStep = <n>`
  // setzen (0..6). Nur für Read-only-Previews; in der echten App ist das
  // Flag nie gesetzt, dann bleibt der Default 0 (Step 1 "Grundlagen").
  const [currentStep, setCurrentStep] = React.useState<number>(() => {
    if (typeof window !== 'undefined') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const init = (window as any).__dexPreviewInitialStep;
      // v29.21 (Audit B5): 9 Schritte, Indizes 0..8 — die alte 7er-Grenze
      // verwarf Schritt 9 (Fun-Zone); Ticket-/Handbuch-Previews darauf
      // zeigten stattdessen Grundlagen.
      // v30.44: …und jetzt 0..9, weil Schritt 10 (Abrechnung) dazugekommen ist.
      // Ohne diese Grenze landete „Angaben ergänzen" aus dem Organizer Center
      // stumm auf Schritt 1 statt bei der Abrechnung — der Wert wurde
      // verworfen, nicht gemeldet. Wer einen Schritt ergänzt, muss diese Zahl
      // mitziehen (dieselbe Falle wie `SCOPE_AWARE_STEPS`, `getStepErrors` …).
      // Auf den letzten Schritt begrenzen, den DIESE Person sieht: Schritt 10
      // (Index 9) haengt an `adminLike`. Ein zu grosser Wert wuerde sonst
      // jede Anzeige-Bedingung `currentStep === N` verfehlen — die Seite
      // bliebe leer, ohne dass irgendwo etwas meldet, warum.
      const maxStep = canBilling ? 9 : 8;
      if (typeof init === 'number' && init >= 0) return Math.min(init, maxStep);
    }
    return 0;
  });
  // v30.44: Die Marke ist ein EINMAL-Signal. Sie wird bewusst hier abgeräumt
  // und nicht beim Setzen: Bliebe sie stehen, öffnete auch das nächste „Event
  // anlegen" auf Schritt 10. Die Vorschau-Modals löschen sie zusätzlich beim
  // Schließen — doppeltes Löschen ist folgenlos.
  React.useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    try { delete (window as any).__dexPreviewInitialStep; } catch { /* */ }
  }, []);
  // v26.30: aktuellen Wizard-Schritt (1-basiert, passend zur DexTicket-
  // Konvention) für den Header-Frage-Button melden — so wird die Frage eines
  // Organizers direkt dem Wizard-Schritt zugeordnet. Beim Verlassen zurücksetzen.
  React.useEffect(() => {
    setActiveWizardStep(currentStep + 1);
    return () => setActiveWizardStep(null);
  }, [currentStep]);
  const [selectedTemplate, setSelectedTemplate] = React.useState<'blank' | 'b2run'>('blank');
  // EventType wird bei neuen Events aus dem Template abgeleitet; bei Edit
  // bleibt der gespeicherte Wert erhalten.
  const eventType: EventType = editEvent ? storedEventType : (selectedTemplate === 'b2run' ? 'B2Run' : 'Other');
  const [b2runStartblocks, setB2runStartblocks] = React.useState<string[]>([]);
  const [newStartblock, setNewStartblock] = React.useState<string>('');
  const [durchstarterCapacity, setDurchstarterCapacity] = React.useState<string>(
    editEvent && typeof editEvent.durchstarterCapacity === 'number' ? String(editEvent.durchstarterCapacity) : ''
  );
  const [funstarterCapacity, setFunstarterCapacity] = React.useState<string>(
    editEvent && typeof editEvent.funstarterCapacity === 'number' ? String(editEvent.funstarterCapacity) : ''
  );
  // v10.20: frei wählbare Bezeichnungen für die zwei Kapazitäts-Gruppen.
  // Default leer; wenn der User die Split-Capacity einschaltet ohne Label
  // zu setzen, fallen Wizard und RegistrationPage auf 'Durchstarter' /
  // 'Funstarter' zurück (Backward-Compat für B2Run-Events vor v10.20).
  const [splitLabelA, setSplitLabelA] = React.useState<string>(
    (editEvent && editEvent.splitLabelA) || ''
  );
  const [splitLabelB, setSplitLabelB] = React.useState<string>(
    (editEvent && editEvent.splitLabelB) || ''
  );
  // v26.72: frei konfigurierbare Beschreibung pro Gruppe (mehrzeilig) —
  // erscheint unter dem Gruppen-Namen in der Auswahl-Karte auf der Anmeldeseite.
  const [splitDescA, setSplitDescA] = React.useState<string>(
    (editEvent && editEvent.splitDescA) || ''
  );
  const [splitDescB, setSplitDescB] = React.useState<string>(
    (editEvent && editEvent.splitDescB) || ''
  );
  // v26.83: frei wählbarer Hinweistext über der Gruppen-Auswahl (ersetzt den
  // Standardsatz „Wähle eine der zwei Gruppen aus…" auf der Anmeldeseite).
  const [splitHelpText, setSplitHelpText] = React.useState<string>(
    (editEvent && editEvent.splitHelpText) || ''
  );
  // v26.83: frei wählbare Überschrift der Gruppen-Auswahl (statt „Gruppen-Auswahl").
  const [splitSectionTitle, setSplitSectionTitle] = React.useState<string>(
    (editEvent && editEvent.splitSectionTitle) || ''
  );
  // v10.20: Warteliste-Modus bei aktiver Split-Capacity. Default false =
  // getrennte Wartelisten pro Gruppe (alter B2Run-Stil). true = eine
  // gemeinsame Warteliste, FIFO über beide Gruppen hinweg.
  const [splitSharedWaitlist, setSplitSharedWaitlist] = React.useState<boolean>(
    !!editEvent?.splitSharedWaitlist
  );
  // v11.25: Anzeige-Reihenfolge der zwei Gruppen-Karten in der Registrierung
  // umkehren. Pure UI-Toggle — interne Daten (splitLabelA/B, Kapazitäten,
  // StarterType auf Anmeldungen) bleiben unangetastet.
  const [splitDisplayOrderReversed, setSplitDisplayOrderReversed] = React.useState<boolean>(
    !!editEvent?.splitDisplayOrderReversed
  );
  // v11.0: Teilnehmer-Upload aktivieren + optionaler Hinweistext
  const [allowAttendeeUpload, setAllowAttendeeUpload] = React.useState<boolean>(
    !!editEvent?.allowAttendeeUpload
  );
  const [attendeeUploadHint, setAttendeeUploadHint] = React.useState<string>(
    editEvent?.attendeeUploadHint || ''
  );
  const [attendeeUploadLabel, setAttendeeUploadLabel] = React.useState<string>(
    editEvent?.attendeeUploadLabel || ''
  );
  // v11.80: Anrede im Registrierungsformular abfragen (Default false). Wenn
  // false, wird das Anrede-Dropdown ausgeblendet und ein leerer String als
  // Anrede gespeichert. Wird im neuen Schritt 5 (Felder) konfiguriert.
  const [askSalutation, setAskSalutation] = React.useState<boolean>(
    !!editEvent?.askSalutation
  );
  // v18.75: Sicherheitshinweis vor dem Absenden der Anmeldung (Schritt 5, ganz
  // unten). Default aus. Modus 'summary' = Auswahl-Übersicht (Haupt-/Sub-Events
  // mit De-/Selektieren), 'freetext' = eigener Hinweis-Text.
  const [confirmDialogEnabled, setConfirmDialogEnabled] = React.useState<boolean>(!!editEvent?.confirmDialogEnabled);
  const [confirmDialogMode, setConfirmDialogMode] = React.useState<string>(editEvent?.confirmDialogMode || 'summary');
  const [confirmDialogText, setConfirmDialogText] = React.useState<string>(editEvent?.confirmDialogText || '');
  // v18.35: Anmeldesprache vorgeben. '' = App-Sprache (Default), 'de' / 'en' =
  // Anmeldeseite (inkl. Disclaimer) immer in dieser Sprache anzeigen.
  const [registrationLanguage, setRegistrationLanguage] = React.useState<'' | 'de' | 'en'>(
    editEvent?.registrationLanguage === 'de' || editEvent?.registrationLanguage === 'en' ? editEvent.registrationLanguage : ''
  );
  // v20.2: Self-Check-in-States aus dem Wizard entfernt — Aktivierung läuft
  // automatisch beim ersten Klick auf die Aktionen (Check-in-Seite, Admin
  // Center, QR-Kachel im Event-Detail); Zeitfenster + Deaktivieren im
  // Kachel-Modal des Admin Centers. Der Wizard fasst die SelfCheckIn*-Spalten
  // weder beim Create noch beim Edit an, damit die dort gesetzten Werte einen
  // Wizard-Save überleben.
  // v11.80: Team-Anmeldung — eine Person meldet ein ganzes Team an.
  // Konfiguration im neuen Schritt 4 (Team-Anmeldung). Die tatsächliche
  // Multi-Person-Anmelde-Logik folgt mit v11.81+; aktuell wird nur die
  // Konfiguration persistiert.
  const [teamRegistrationEnabled, setTeamRegistrationEnabled] = React.useState<boolean>(
    !!editEvent?.teamRegistrationEnabled
  );
  const [teamSize, setTeamSize] = React.useState<number>(
    typeof editEvent?.teamSize === 'number' && editEvent.teamSize > 0 ? editEvent.teamSize : 4
  );
  const [askTeamName, setAskTeamName] = React.useState<boolean>(
    !!editEvent?.askTeamName
  );
  // v11.81: Erweiterte Team-Konfiguration — Beitritts-Modus, Sichtbarkeit
  // offener Slots, Lead-Approval. Die tatsächliche Team-Anmelde-Logik
  // (Multi-Person-Form, Mails, Outlook) folgt mit v11.82+.
  const [teamPartialAllowed, setTeamPartialAllowed] = React.useState<boolean>(
    !!editEvent?.teamPartialAllowed
  );
  const [teamOpenSlotsVisible, setTeamOpenSlotsVisible] = React.useState<boolean>(
    !!editEvent?.teamOpenSlotsVisible
  );
  const [teamJoinRequiresApproval, setTeamJoinRequiresApproval] = React.useState<boolean>(
    !!editEvent?.teamJoinRequiresApproval
  );
  // v22.78: frei benennbarer Team-Begriff (z.B. „Break-Out Session") +
  // „Teilnehmer dürfen keine neuen Teams erstellen".
  const [teamTermSingular, setTeamTermSingular] = React.useState<string>(editEvent?.teamTermSingular || '');
  const [teamTermPlural, setTeamTermPlural] = React.useState<string>(editEvent?.teamTermPlural || '');
  const [teamMembersCannotCreate, setTeamMembersCannotCreate] = React.useState<boolean>(!!editEvent?.teamMembersCannotCreate);
  // v24.58: Anzeige-Bezeichnung des Haupt-Events in der Sub-Event-Auswahl.
  const [mainEventLabelMode, setMainEventLabelMode] = React.useState<'default' | 'custom' | 'none'>(editEvent?.mainEventLabelMode || 'default');
  const [mainEventLabel, setMainEventLabel] = React.useState<string>(editEvent?.mainEventLabel || '');
  // v17.20: Bilingual-Toggle — wenn an, kann der Organizer pro Custom-Field
  // (Label, Help-Text, Checkbox-Confirm-Text, Dropdown-Optionen) eine
  // englische Variante hinterlegen. Wird im Wizard-Schritt 5 ganz oben als
  // separater Toggle eingestellt; die EN-Inputs blenden pro Card auf, wenn
  // der Toggle aktiv ist.
  const [bilingualFields, setBilingualFields] = React.useState<boolean>(
    !!editEvent?.bilingualFields
  );
  // v6.15: Starter-Typ → Startblock-Zuordnung + Leistungsnachweis-Pflicht
  const [durchstarterStartblock, setDurchstarterStartblock] = React.useState<string>(
    editEvent?.durchstarterStartblock || ''
  );
  const [funstarterStartblock, setFunstarterStartblock] = React.useState<string>(
    editEvent?.funstarterStartblock || ''
  );
  // v10.24: setDurchstarterRequiresProof wird nicht mehr aufgerufen — der UI-
  // Toggle in Schritt 3 ist entfallen, das Feature wird durch Pro-Gruppe-
  // Custom-Fields in Schritt 4 ersetzt. State bleibt erhalten, damit
  // bestehende Events mit gesetztem Wert nicht beim Save den Wert verlieren
  // (durchstarterRequiresProof wird beim Persist mitgeschrieben falls
  // editEvent das Flag schon hatte).
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [durchstarterRequiresProof, setDurchstarterRequiresProof] = React.useState<boolean>(
    !!editEvent?.durchstarterRequiresProof
  );
  const [showPreview, setShowPreview] = React.useState(false);
  const [showRegisterPreview, setShowRegisterPreview] = React.useState(false);
  const [triedNext, setTriedNext] = React.useState(false);
  // v26.86: Die Blöcke im Schritt „Kapazität & Sichtbarkeit" sind einzeln
  // einklappbar und beim ersten Aufruf EINGEKLAPPT (Set enthält die offenen
  // Keys → leer = alles zu), damit der Schritt nicht überfordert. Der äußere
  // Block-<div> (inkl. zebraS3Bg-Alternation) bleibt erhalten; nur der Body
  // wird ein-/ausgeblendet, die Überschrift ist der Klappschalter.
  // Bei ausgelöster Validierung (triedNext) klappt automatisch ALLES auf, damit
  // keine Fehlermeldung in einem eingeklappten Block versteckt bleibt.
  const [expandedVisBlocks, setExpandedVisBlocks] = React.useState<Set<string>>(() => new Set());
  const isVisOpen = (k: string): boolean => triedNext || expandedVisBlocks.has(k);
  const toggleVis = (k: string): void => setExpandedVisBlocks(prev => { const n = new Set(prev); if (n.has(k)) n.delete(k); else n.add(k); return n; });
  const visHeader = (key: string, badge: React.ReactNode, title: React.ReactNode): React.ReactElement => (
    <button type="button" onClick={() => toggleVis(key)} aria-expanded={isVisOpen(key)}
      style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 8, background: 'none', border: 'none', padding: 0, marginBottom: isVisOpen(key) ? 8 : 0, cursor: 'pointer', textAlign: 'left' }}>
      {badge}
      <span className="form-label" style={{ margin: 0, fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: 8 }}>{title}</span>
      <span style={{ marginLeft: 'auto', color: 'var(--dex-gray-400)', fontSize: '0.85rem', transform: isVisOpen(key) ? 'rotate(90deg)' : 'none', transition: 'transform .15s' }}>▶</span>
    </button>
  );
  // v22.62/v22.63: Beim „Weiter"/Speichern fragt ein Modal, ob die geänderte
  // Klammer-/Hauptevent-Sichtbarkeit auf alle Sub-Events übernommen werden soll
  // — IMMER wenn die Klammer-Sichtbarkeit geändert/neu gesetzt wurde UND von
  // den Sub-Events abweicht. `visSnapshotRef` hält den zuletzt „abgehandelten"
  // Sichtbarkeits-Stand (Baseline beim Mount via Effekt unten).
  const [visCopyModalOpen, setVisCopyModalOpen] = React.useState(false);
  const visCopyPendingRef = React.useRef<(() => void) | null>(null);
  const visSnapshotRef = React.useRef<string | null>(null);
  const [previewSections, setPreviewSections] = React.useState<Array<{ id: string; label: string }>>([
    { id: 'event', label: 'Event-Karte' },
    { id: 'personal', label: 'Personal Information' },
    { id: 'specific', label: 'Event specific Information' },
    { id: 'actions', label: 'Buttons' },
  ]);
  const [dragSectionId, setDragSectionId] = React.useState<string | null>(null);
  const [dragOverSectionId, setDragOverSectionId] = React.useState<string | null>(null);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [progress, setProgress] = React.useState(0);
  const [progressLabel, setProgressLabel] = React.useState('');
  // v19.x: „Sichtbarkeit prüfen"-Modal-State (Verteiler-Cache, Testpersonen-
  // Suche) ist nach <AudiencePicker> gewandert.
  const [submitted, setSubmitted] = React.useState(false);
  const [error, setError] = React.useState('');
  const [imageUploadError, setImageUploadError] = React.useState('');

  const locationOptions = ['Berlin', 'Dresden', 'Düsseldorf', 'Frankfurt', 'Görlitz', 'Halle', 'Hamburg', 'Hannover', 'Köln', 'Leipzig', 'Magdeburg', 'Mannheim', 'München', 'Nürnberg', 'Stuttgart', 'Walldorf', 'All'];

  const addCustomField = (): void => {
    const newId = `cf-${Date.now()}`;
    setCustomFields([...customFields, {
      id: newId, label: '', type: 'text',
      required: false, options: [], visible: true,
    }]);
    // v18.55: neues Feld direkt aufgeklappt, damit man es sofort ausfüllen kann.
    setFieldExpandOverride(prev => ({ ...prev, [newId]: true }));
  };

  /**
   * Deloitte-Standard-Vorschläge als Katalog. Der Organizer wählt über ein
   * Modal mit Checkboxen aus, welche dieser Felder hinzugefügt werden sollen.
   * Ausgewählte Felder werden ans Ende der aktuellen customFields angehängt.
   */
  // Bilingual: Labels + Optionen der Felder werden in der Event-Sprache (DE/EN)
  // angelegt, passend zum Locale beim Klick auf 'Vorgeschlagene Felder'.
  // v10.21: Catalog mit Kategorien — 'general' (default ausgeklappt) und
  // 'b2run' (default eingeklappt). Damit ersetzt das Suggested-Modal den
  // alten Template-Dropdown: User wählt fokussiert die Felder, die er
  // wirklich braucht, statt einen B2Run-Block auf einmal aufzuziehen.
  // v10.23: jeder Suggested-Field-Eintrag hat ein Fluent-UI-Icon (visuelles
  // Erkennungsmerkmal in der Auswahl-Liste) und einen ausführlicheren
  // Tooltip-Text — der erklärt dem Organizer, was das Feld in der App
  // bewirkt, ohne dass er es erst hinzufügen muss.
  const SUGGESTED_FIELDS_CATALOG: SuggestedEntry[] = isDe ? [
    {
      // v22.38: Sonder-Eintrag — schaltet das Standard-Anrede-Feld an
      // (askSalutation-Flag) statt ein Custom-Field anzulegen. Wird in
      // addSelectedSuggestedFields gesondert behandelt.
      key: 'salutation', category: 'general', icon: 'Contact',
      label: 'Anrede',
      description: 'Pflicht-Dropdown Frau / Herr / Divers / Keine Angabe — erscheint über dem Vornamen',
      build: (n) => ({ id: `cf-${n}`, label: 'Anrede', type: 'select', required: true, options: [], visible: true }),
    },
    {
      key: 'tshirt', category: 'general', icon: 'Tag',
      label: 'T-Shirt Größe',
      description: 'Dropdown mit Kein T-Shirt / XS–XXL',
      build: (n) => ({ id: `cf-${n}`, label: 'T-Shirt Größe', type: 'select', required: false, options: ['Habe bereits ein T-Shirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'allergies', category: 'general', icon: 'Warning',
      label: 'Allergien',
      description: 'Freitextfeld für Allergien/Unverträglichkeiten',
      build: (n) => ({ id: `cf-${n}`, label: 'Allergien', type: 'text', required: false, options: [], visible: true }),
    },
    {
      key: 'diet', category: 'general', icon: 'EatDrink',
      label: 'Essenspräferenzen',
      description: 'Dropdown: Keine Präferenzen / Vegetarisch / Vegan / Pescetarisch',
      build: (n) => ({ id: `cf-${n}`, label: 'Essenspräferenzen', type: 'select', required: false, options: ['Keine Präferenzen', 'Vegetarisch', 'Vegan', 'Pescetarisch'], visible: true }),
    },
    {
      key: 'hotel', category: 'general', icon: 'Hotel',
      label: 'Hotel benötigt',
      description: 'Checkbox: Teilnehmer benötigt ein Hotel',
      build: (n) => ({ id: `cf-${n}`, label: 'Hotel benötigt', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'roomtype', category: 'general', icon: 'Room',
      label: 'Zimmerart',
      description: 'Dropdown: Keine Präferenz / Einzelzimmer / Doppelzimmer',
      build: (n) => ({ id: `cf-${n}`, label: 'Zimmerart (falls Hotel benötigt)', type: 'select', required: false, options: ['Keine Präferenz', 'Einzelzimmer', 'Doppelzimmer'], visible: true }),
    },
    {
      key: 'roommate', category: 'general', icon: 'People',
      label: 'Bevorzugter Zimmerpartner',
      description: 'Personen-Suche; Match-Erkennung im Admin Center',
      build: (n) => ({ id: `cf-${n}`, label: 'Bevorzugter Zimmerpartner (bei Doppelzimmer)', type: 'roommate', required: false, options: [], visible: true }),
    },
    // B2Run-Pakete — nur für Lauf-Events relevant. Sektion ist im Modal
    // standardmäßig eingeklappt, damit der Standard-Organizer sie nicht
    // versehentlich aktiviert.
    {
      key: 'b2run_startblock', category: 'b2run', icon: 'Running',
      label: 'Startblock',
      description: 'Dropdown der Startblöcke. Optionen werden nachträglich im Wizard gepflegt.',
      build: (_n) => ({ id: `b2run_startblock`, label: 'Startblock', type: 'select', required: true, options: [], visible: true }),
    },
    {
      key: 'b2run_gruppe', category: 'b2run', icon: 'BulletedList',
      label: 'Gruppe',
      description: 'Dropdown: offene Klasse / Nordic Walker / Damen / Herren',
      build: (_n) => ({ id: `b2run_gruppe`, label: 'Gruppe', type: 'select', required: true, options: ['offene Klasse', 'Nordic Walker', 'Damen', 'Herren'], visible: true }),
    },
    {
      key: 'b2run_altersklasse', category: 'b2run', icon: 'Calendar',
      label: 'Altersklasse',
      description: 'Dropdown: unter 18 / 18-29 / 30-39 / 40-49 / 50-59 / 60+',
      build: (_n) => ({ id: `b2run_altersklasse`, label: 'Altersklasse', type: 'select', required: true, options: ['unter 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true }),
    },
    {
      key: 'b2run_infoservice', category: 'b2run', icon: 'CellPhone',
      label: 'Infoservice (SMS)',
      description: 'Checkbox: aktiviert die Mobilnummer-Pflicht für den B2Run-SMS-Service',
      build: (_n) => ({ id: `b2run_infoservice`, label: 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_mobilnummer', category: 'b2run', icon: 'Phone',
      label: 'Mobilnummer',
      description: 'Freitext, dynamisch Pflicht wenn Infoservice aktiv',
      build: (_n) => ({ id: `b2run_mobilnummer`, label: 'Mobilnummer (nur bei aktiviertem Infoservice)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } }),
    },
    {
      key: 'b2run_anonym', category: 'b2run', icon: 'Hide3',
      label: 'Anonym teilnehmen',
      description: 'Checkbox: Teilnehmer in Ergebnislisten anonymisieren',
      build: (_n) => ({ id: `b2run_anonym`, label: 'Anonym teilnehmen', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_laufshirt', category: 'b2run', icon: 'Sport',
      label: 'Deloitte-Laufshirt',
      description: 'Dropdown: vorhandenes Shirt / XS-XXL',
      build: (_n) => ({ id: `b2run_laufshirt`, label: 'Deloitte-Laufshirt', type: 'select', required: true, options: ['Habe bereits ein Laufshirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'b2run_datenschutz', category: 'b2run', icon: 'LockShield',
      label: 'AGB / Datenschutz',
      description: 'Pflicht-Checkbox mit Links zu B2Run-AGB und Datenschutzerklärung',
      build: (_n) => ({
        id: `b2run_datenschutz`,
        label: 'Zustimmung AGB, Datenschutz & Bildaufnahmen',
        type: 'checkbox', required: true, options: [], visible: true,
        externalLinks: [
          { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
          { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
        ],
      }),
    },
  ] : [
    {
      key: 'salutation', category: 'general', icon: 'Contact',
      label: 'Salutation',
      description: 'Required dropdown Mrs / Mr / Diverse / Prefer not to say — shown above the first name',
      build: (n) => ({ id: `cf-${n}`, label: 'Salutation', type: 'select', required: true, options: [], visible: true }),
    },
    {
      key: 'tshirt', category: 'general', icon: 'Tag',
      label: 'T-Shirt size',
      description: 'Dropdown: No t-shirt needed / XS–XXL',
      build: (n) => ({ id: `cf-${n}`, label: 'T-Shirt size', type: 'select', required: false, options: ['I already have one', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'allergies', category: 'general', icon: 'Warning',
      label: 'Allergies',
      description: 'Free-text field for allergies / intolerances',
      build: (n) => ({ id: `cf-${n}`, label: 'Allergies', type: 'text', required: false, options: [], visible: true }),
    },
    {
      key: 'diet', category: 'general', icon: 'EatDrink',
      label: 'Dietary preferences',
      description: 'Dropdown: No preference / Vegetarian / Vegan / Pescetarian',
      build: (n) => ({ id: `cf-${n}`, label: 'Dietary preferences', type: 'select', required: false, options: ['No preference', 'Vegetarian', 'Vegan', 'Pescetarian'], visible: true }),
    },
    {
      key: 'hotel', category: 'general', icon: 'Hotel',
      label: 'Hotel required',
      description: 'Checkbox: participant needs a hotel room',
      build: (n) => ({ id: `cf-${n}`, label: 'Hotel required', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'roomtype', category: 'general', icon: 'Room',
      label: 'Room type',
      description: 'Dropdown: No preference / Single room / Double room',
      build: (n) => ({ id: `cf-${n}`, label: 'Room type (if hotel needed)', type: 'select', required: false, options: ['No preference', 'Single room', 'Double room'], visible: true }),
    },
    {
      key: 'roommate', category: 'general', icon: 'People',
      label: 'Preferred roommate',
      description: 'People search; match detection in the admin center',
      build: (n) => ({ id: `cf-${n}`, label: 'Preferred roommate (for double room)', type: 'roommate', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_startblock', category: 'b2run', icon: 'Running',
      label: 'Start block',
      description: 'Dropdown of start blocks. Options are added later in the wizard.',
      build: (_n) => ({ id: `b2run_startblock`, label: 'Start block', type: 'select', required: true, options: [], visible: true }),
    },
    {
      key: 'b2run_gruppe', category: 'b2run', icon: 'BulletedList',
      label: 'Category',
      description: 'Dropdown: Open class / Nordic Walker / Women / Men',
      build: (_n) => ({ id: `b2run_gruppe`, label: 'Category', type: 'select', required: true, options: ['Open class', 'Nordic Walker', 'Women', 'Men'], visible: true }),
    },
    {
      key: 'b2run_altersklasse', category: 'b2run', icon: 'Calendar',
      label: 'Age group',
      description: 'Dropdown: under 18 / 18-29 / 30-39 / 40-49 / 50-59 / 60+',
      build: (_n) => ({ id: `b2run_altersklasse`, label: 'Age group', type: 'select', required: true, options: ['under 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true }),
    },
    {
      key: 'b2run_infoservice', category: 'b2run', icon: 'CellPhone',
      label: 'Info service (SMS)',
      description: 'Checkbox: enables the mandatory mobile-number for the B2Run SMS service',
      build: (_n) => ({ id: `b2run_infoservice`, label: 'Use B2Run info service (SMS — mobile number required)', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_mobilnummer', category: 'b2run', icon: 'Phone',
      label: 'Mobile number',
      description: 'Free text, dynamically required when info service is active',
      build: (_n) => ({ id: `b2run_mobilnummer`, label: 'Mobile number (only if info service is enabled)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } }),
    },
    {
      key: 'b2run_anonym', category: 'b2run', icon: 'Hide3',
      label: 'Anonymous participation',
      description: 'Checkbox: anonymise attendee in result lists',
      build: (_n) => ({ id: `b2run_anonym`, label: 'Participate anonymously', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_laufshirt', category: 'b2run', icon: 'Sport',
      label: 'Deloitte running shirt',
      description: 'Dropdown: existing shirt / XS-XXL',
      build: (_n) => ({ id: `b2run_laufshirt`, label: 'Deloitte running shirt', type: 'select', required: true, options: ['I already have one', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'b2run_datenschutz', category: 'b2run', icon: 'LockShield',
      label: 'Terms / privacy',
      description: 'Required checkbox with links to B2Run terms and privacy policy',
      build: (_n) => ({
        id: `b2run_datenschutz`,
        label: 'I agree to the terms, privacy policy and photo/video recordings',
        type: 'checkbox', required: true, options: [], visible: true,
        externalLinks: [
          { label: 'Terms (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
          { label: 'Privacy (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
        ],
      }),
    },
  ];

  const [showSuggestedModal, setShowSuggestedModal] = React.useState(false);
  const [suggestedSelection, setSuggestedSelection] = React.useState<Record<string, boolean>>({});
  // v10.21: B2Run-Sektion im Suggested-Modal default eingeklappt — die meisten
  // Organizer brauchen sie nicht; soll nicht visuell übernehmen.
  const [showB2runSuggested, setShowB2runSuggested] = React.useState(false);

  const openSuggestedModal = (): void => {
    // v9.17: Standard ist KEINS ausgewählt — User wählt aktiv aus, was er
    // wirklich braucht. Vorher waren alle vorgewählt, was zu unbeabsichtigt
    // viele übernommenen Feldern führte.
    setSuggestedSelection({});
    setShowSuggestedModal(true);
  };

  const addSelectedSuggestedFields = (): void => {
    const selected = SUGGESTED_FIELDS_CATALOG.filter(s => suggestedSelection[s.key]);
    if (selected.length === 0) { setShowSuggestedModal(false); return; }
    // v22.38: Sonder-Eintrag „Anrede" schaltet das Standard-Anrede-Feld an
    // (askSalutation-Flag, Pseudo-Zeile in der Feld-Liste) statt ein
    // Custom-Field anzulegen.
    if (selected.some(s => s.key === 'salutation')) setAskSalutation(true);
    const buildable = selected.filter(s => s.key !== 'salutation');
    if (buildable.length === 0) { setShowSuggestedModal(false); return; }
    const now = Date.now();
    const newFields: CustomFieldInput[] = buildable.map((s, i) => s.build(now + i));
    // v10.21: B2Run-Felder haben deterministische IDs (b2run_startblock etc.).
    // Wenn ein Feld mit gleicher ID schon im customFields-Array steht, skippen
    // wir es — sonst entstehen Duplikate, wenn der User das Modal mehrfach
    // öffnet. Allgemeine Felder (cf-<timestamp>) bekommen eindeutige IDs und
    // werden immer angehängt.
    const existingIds = new Set(customFields.map(f => f.id));
    const dedupedNewFields = newFields.filter(f => !existingIds.has(f.id));
    setCustomFields([...customFields, ...dedupedNewFields]);
    setShowSuggestedModal(false);
  };

  const removeCustomField = (id: string): void => {
    setCustomFields(customFields.filter(f => f.id !== id));
  };

  const updateCustomField = (id: string, updates: Partial<CustomFieldInput>): void => {
    setCustomFields(customFields.map(f => f.id === id ? { ...f, ...updates } : f));
  };
  // v24.16: Sichtbarkeitsbedingung (showIf) als wiederverwendbarer Helfer —
  // genutzt vom Hauptevent UND von Sub-Event-Feldern (vorher fehlte die UI bei
  // Sub-Events komplett, daher liessen sich Bedingungen dort nie setzen).
  const renderShowIfConfig = (field: CustomFieldInput, idx: number, allFields: CustomFieldInput[], onUpdate: (u: Partial<CustomFieldInput>) => void): React.ReactElement => {
                      const candidateSources = allFields.slice(0, idx).filter(other =>
                        (other.type === 'select' || other.type === 'checkbox') && (other.label || '').trim().length > 0
                      );
                      const sourceField = field.showIf
                        ? allFields.find(o => o.id === field.showIf!.fieldId)
                        : null;
                      const removeShowIf = (): void => {
                        // showIf gezielt löschen: updateCustomField macht ein
                        // shallow-merge, also setzen wir undefined und filtern
                        // beim Save raus.
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        onUpdate({ showIf: undefined as any });
                      };
                      return (
                        <div style={{ marginLeft: 32, marginTop: 10, padding: '10px 12px', background: 'rgba(21,101,192,0.04)', border: '1px dashed var(--dex-gray-300)', borderRadius: 8 }}>
                          {!field.showIf ? (
                            <span style={{ display: 'inline-flex', alignItems: 'center' }}>
                              <button
                                type="button"
                                onClick={() => {
                                  if (candidateSources.length === 0) {
                                    showAlert(isDe
                                      ? 'Es gibt noch kein Dropdown- oder Checkbox-Feld VOR diesem hier, an das die Sichtbarkeit geknüpft werden könnte. Lege zuerst ein passendes Feld weiter oben an.'
                                      : 'There is no dropdown or checkbox field BEFORE this one yet that visibility could depend on. Please add a suitable field above first.');
                                    return;
                                  }
                                  const first = candidateSources[0];
                                  onUpdate({
                                    showIf: {
                                      fieldId: first.id,
                                      values: first.type === 'checkbox' ? ['true'] : (first.options[0] ? [first.options[0]] : []),
                                    },
                                  });
                                }}
                                style={{
                                  background: 'none', border: 'none', padding: 0,
                                  color: 'var(--dex-green-dark, #4a7c1f)',
                                  fontSize: '0.8rem', fontWeight: 600, cursor: 'pointer',
                                }}
                              >
                                + {isDe ? 'Sichtbarkeitsbedingung hinzufügen' : 'Add visibility condition'}
                              </button>
                              <InfoTooltip
                                text={isDe
                                  ? 'Dieses Feld wird nur angezeigt, wenn die Antwort auf eine andere (zuvor angelegte) Frage einem von dir festgelegten Wert entspricht. Beispiel: „Roommate" wird nur gefragt, wenn die Frage „Zimmerart" mit „Doppelzimmer" beantwortet wurde. Andernfalls bleibt das Feld komplett verborgen — und blockiert auch nicht die Pflichtfeld-Validierung.'
                                  : 'This field is shown only when the answer to another (previously added) question matches a value you specify. Example: "Roommate" is only asked when the question "Room type" is answered with "Double room". Otherwise the field stays fully hidden — and does not block the required-field validation either.'}
                              />
                            </span>
                          ) : (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', fontWeight: 600, display: 'flex', alignItems: 'center' }}>
                                {isDe ? 'Diese Frage nur anzeigen wenn:' : 'Only show this question when:'}
                                <InfoTooltip
                                  text={isDe
                                    ? 'Dieses Feld wird nur angezeigt, wenn die Antwort auf die Quell-Frage einem der gewählten Werte entspricht. Bei Mehrfachauswahl-Quellen reicht ein Treffer. Pflichtfeld-Validierung wird übersprungen, solange das Feld verborgen ist.'
                                    : 'This field is shown only when the answer to the source question matches one of the chosen values. With multi-select sources a single match is enough. Required-field validation is skipped as long as the field stays hidden.'}
                                />
                              </div>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                                <select
                                  className="form-select"
                                  value={field.showIf.fieldId}
                                  onChange={e => {
                                    const newSrc = allFields.find(o => o.id === e.target.value);
                                    if (!newSrc) return;
                                    onUpdate({
                                      showIf: {
                                        fieldId: newSrc.id,
                                        values: newSrc.type === 'checkbox' ? ['true'] : (newSrc.options[0] ? [newSrc.options[0]] : []),
                                      },
                                    });
                                  }}
                                  style={{ fontSize: '0.82rem', padding: '4px 8px', minWidth: 180, maxWidth: 320 }}
                                >
                                  {candidateSources.map(o => (
                                    <option key={o.id} value={o.id}>
                                      {allFields.findIndex(c => c.id === o.id) + 1}. {o.label}
                                    </option>
                                  ))}
                                  {/* fallback wenn die ausgewählte Quelle hinter dem Feld gelandet
                                      ist (z.B. nach einem Move) — option in der Liste anzeigen,
                                      aber als ungültig markiert lassen. */}
                                  {sourceField && !candidateSources.find(c => c.id === sourceField.id) && (
                                    <option value={sourceField.id} disabled>
                                      ⚠ {sourceField.label} ({isDe ? 'liegt hinter diesem Feld' : 'is positioned after this field'})
                                    </option>
                                  )}
                                </select>
                                <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                                  {isDe ? '=' : '='}
                                </span>
                                {sourceField && sourceField.type === 'checkbox' ? (
                                  <select
                                    className="form-select"
                                    value={field.showIf.values[0] || 'true'}
                                    onChange={e => onUpdate({
                                      showIf: { fieldId: field.showIf!.fieldId, values: [e.target.value] },
                                    })}
                                    style={{ fontSize: '0.82rem', padding: '4px 8px', minWidth: 130 }}
                                  >
                                    <option value="true">{isDe ? 'angehakt' : 'checked'}</option>
                                    <option value="false">{isDe ? 'nicht angehakt' : 'unchecked'}</option>
                                  </select>
                                ) : sourceField ? (
                                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                                    {(sourceField.options || []).filter(Boolean).map(opt => {
                                      const checked = field.showIf!.values.indexOf(opt) >= 0;
                                      return (
                                        <label
                                          key={opt}
                                          style={{
                                            display: 'inline-flex', alignItems: 'center', gap: 4,
                                            padding: '4px 10px', borderRadius: 999,
                                            fontSize: '0.78rem', cursor: 'pointer',
                                            border: `1px solid ${checked ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                            background: checked ? 'rgba(134,188,37,0.10)' : '#fff',
                                            color: checked ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                                            fontWeight: 600,
                                          }}
                                        >
                                          <input
                                            type="checkbox"
                                            checked={checked}
                                            onChange={() => {
                                              const next = checked
                                                ? field.showIf!.values.filter(v => v !== opt)
                                                : [...field.showIf!.values, opt];
                                              onUpdate({
                                                showIf: { fieldId: field.showIf!.fieldId, values: next },
                                              });
                                            }}
                                            style={{ display: 'none' }}
                                          />
                                          {checked ? '✓' : '○'} {opt}
                                        </label>
                                      );
                                    })}
                                  </div>
                                ) : null}
                                <button
                                  type="button"
                                  onClick={removeShowIf}
                                  title={isDe ? 'Bedingung entfernen' : 'Remove condition'}
                                  style={{
                                    background: 'none', border: 'none', cursor: 'pointer',
                                    color: 'var(--dex-red, #c00)', fontSize: '0.8rem',
                                    padding: '4px 6px', marginLeft: 'auto',
                                  }}
                                >
                                  ✕ {isDe ? 'entfernen' : 'remove'}
                                </button>
                              </div>
                            </div>
                          )}
                        </div>
                      );
  };


  // === Sub-Event Custom-Field Helpers (v10.11+) =============================
  // Per-Sub-Event Custom-Fields ersetzen die hardcoded Funstarter/Durchstarter-
  // Frage. Pattern parallel zu den Hauptevent-Helpers — operieren aber auf dem
  // `customFields[]` eines spezifischen SubEventDraft (nach Client-`id`
  // identifiziert). Funktional minimaler als die Hauptevent-Variante (kein
  // Suggested-Modal, kein showIf für v1), reicht aber für „Auswahlfrage pro
  // Sub-Event mit individuellem Label + Optionen".
  const addSubEventCustomField = (subEventId: string): void => {
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: [...(se.customFields || []), {
        id: `cf_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
        label: '',
        type: 'select',
        required: false,
        options: ['', ''],
        visible: true,
      }],
    })));
  };
  const removeSubEventCustomField = (subEventId: string, fieldId: string): void => {
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: (se.customFields || []).filter(f => f.id !== fieldId),
    })));
  };
  const updateSubEventCustomField = (subEventId: string, fieldId: string, updates: Partial<CustomFieldInput>): void => {
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: (se.customFields || []).map(f => f.id === fieldId ? { ...f, ...updates } : f),
    })));
  };
  const copyParentFieldsToSubEvent = (subEventId: string): void => {
    // Dupliziert die Hauptevent-Felder ins Sub-Event mit frischen IDs (sonst
    // kollidieren Field-IDs zwischen Parent und Children, was bei Validierungs-
    // Logik und showIf-Refs zu Konflikten führen würde).
    const cloned: CustomFieldInput[] = customFields.map(f => ({
      ...f,
      id: `cf_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      options: f.options.slice(),
      externalLinks: f.externalLinks ? f.externalLinks.map(x => ({ ...x })) : undefined,
      showIf: undefined,  // showIf-Refs würden auf Parent-Field-IDs zeigen, droppen
    }));
    setSubEvents(prev => prev.map(se => se.id !== subEventId ? se : ({
      ...se,
      customFields: cloned,
    })));
  };

  /**
   * Template-Auswahl: setzt EventType und Custom Fields automatisch.
   * B2Run: legt alle Pflichtfelder für die Anmeldung bei b2run.com an
   * (laut Excel "Deloitte_Teilnehmer_innen_B2Run_Koeln_2025_v4.xlsx").
   *
   * v10.21: Template-Dropdown im Wizard entfällt; B2Run-Felder werden über
   * das Suggested-Felder-Modal einzeln gewählt. Diese Funktion bleibt für
   * eventuelle programmatische Aufrufer (Edit-Modus, Migrations-Skripte)
   * erhalten — sie wird im aktuellen UI nicht mehr aufgerufen.
   */
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const applyTemplate = (template: 'blank' | 'b2run'): void => {
    setSelectedTemplate(template);
    if (template === 'blank') {
      // v7.20-Fix: NICHT alle Fields löschen — nur die B2Run-spezifischen
      // (Präfix "b2run_"). So gehen Custom-Felder, die der Organizer manuell
      // angelegt hat, beim Deselect des B2Run-Templates nicht verloren.
      setCustomFields(prev => prev.filter(f => !f.id.startsWith('b2run_')));
      setB2runStartblocks([]);
      return;
    }
    if (template === 'b2run') {
      // Custom Fields in der Reihenfolge der B2Run-Excel-Spalten
      // Hinweis: Strasse/PLZ/Stadt werden NICHT abgefragt (werden leer in der Excel stehen)
      // Locale-abhängige Labels/Optionen. IDs bleiben konstant, damit die
      // B2Run-Logik (Infoservice -> Mobilnummer, CSV-Export etc.) unabhängig
      // von der Sprache funktioniert.
      const fields: CustomFieldInput[] = isDe ? [
        { id: 'b2run_startblock', label: 'Startblock', type: 'select', required: true, options: [...b2runStartblocks], visible: true },
        { id: 'b2run_gruppe', label: 'Gruppe', type: 'select', required: true, options: ['offene Klasse', 'Nordic Walker', 'Damen', 'Herren'], visible: true },
        { id: 'b2run_altersklasse', label: 'Altersklasse', type: 'select', required: true, options: ['unter 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true },
        { id: 'b2run_infoservice', label: 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_mobilnummer', label: 'Mobilnummer (nur bei aktiviertem Infoservice)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } },
        { id: 'b2run_anonym', label: 'Anonym teilnehmen', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_laufshirt', label: 'Deloitte-Laufshirt', type: 'select', required: true, options: ['Habe bereits ein Laufshirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true },
        {
          id: 'b2run_datenschutz',
          label: 'Zustimmung AGB, Datenschutz & Bildaufnahmen',
          type: 'checkbox',
          required: true,
          options: [],
          visible: true,
          externalLinks: [
            { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
            { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
          ],
        },
      ] : [
        { id: 'b2run_startblock', label: 'Start block', type: 'select', required: true, options: [...b2runStartblocks], visible: true },
        { id: 'b2run_gruppe', label: 'Category', type: 'select', required: true, options: ['Open class', 'Nordic Walker', 'Women', 'Men'], visible: true },
        { id: 'b2run_altersklasse', label: 'Age group', type: 'select', required: true, options: ['under 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true },
        { id: 'b2run_infoservice', label: 'Use B2Run info service (SMS — mobile number required)', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_mobilnummer', label: 'Mobile number (only if info service is enabled)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } },
        { id: 'b2run_anonym', label: 'Participate anonymously', type: 'checkbox', required: false, options: [], visible: true },
        { id: 'b2run_laufshirt', label: 'Deloitte running shirt', type: 'select', required: true, options: ['I already have one', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true },
        {
          id: 'b2run_datenschutz',
          label: 'I agree to the terms, privacy policy and photo/video recordings',
          type: 'checkbox',
          required: true,
          options: [],
          visible: true,
          externalLinks: [
            { label: 'Terms (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
            { label: 'Privacy (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
          ],
        },
      ];
      // v7.20-Fix: bestehende NON-b2run-Felder erhalten und die B2Run-Felder
      // anhängen (vorher: setCustomFields(fields) hat alles überschrieben).
      setCustomFields(prev => {
        const nonB2run = prev.filter(f => !f.id.startsWith('b2run_'));
        return [...nonB2run, ...fields];
      });
    }
  };

  // Startblöcke-Änderung direkt in das Custom Field übernehmen
  React.useEffect(() => {
    if (selectedTemplate !== 'b2run' && !(isEditMode && customFields.some(f => f.id === 'b2run_startblock'))) return;
    setCustomFields(prev => prev.map(f =>
      f.id === 'b2run_startblock' ? { ...f, options: [...b2runStartblocks] } : f
    ));
  }, [b2runStartblocks]);

  // Edit-Mode: Wenn das Event B2Run-Custom-Fields hat, Startblöcke aus dem Field laden
  React.useEffect(() => {
    if (!isEditMode) return;
    const sb = customFields.find(f => f.id === 'b2run_startblock');
    if (sb && b2runStartblocks.length === 0 && sb.options && sb.options.length > 0) {
      const parts = sb.options.map(s => s.trim()).filter(Boolean);
      setB2runStartblocks(parts);
      setSelectedTemplate('b2run');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEditMode]);

  const addStartblock = (): void => {
    const trimmed = newStartblock.trim();
    if (!trimmed) return;
    if (b2runStartblocks.indexOf(trimmed) >= 0) { setNewStartblock(''); return; }
    setB2runStartblocks([...b2runStartblocks, trimmed]);
    setNewStartblock('');
  };

  const removeStartblock = (block: string): void => {
    setB2runStartblocks(b2runStartblocks.filter(b => b !== block));
  };

  // Bei B2Run: maxParticipants automatisch aus Summe von Durchstarter + Funstarter berechnen
  const isB2runTemplate = selectedTemplate === 'b2run' || (isEditMode && customFields.some(f => f.id === 'b2run_startblock'));
  // Seit v6.5: explizite Checkbox in Schritt 3 ("Lauf-Event mit getrennten
  // Starter-Kapazitäten") statt versteckt über das Template gesteuert.
  // Initial-Wert: beim Edit aus vorhandenen Kapazitäten abgeleitet, bei neuem
  // Event true wenn B2Run-Template gewählt wurde.
  const [useSplitCapacities, setUseSplitCapacities] = React.useState<boolean>(() => {
    if (editEvent) {
      return typeof editEvent.durchstarterCapacity === 'number'
        && typeof editEvent.funstarterCapacity === 'number'
        && (editEvent.durchstarterCapacity > 0 || editEvent.funstarterCapacity > 0);
    }
    return selectedTemplate === 'b2run';
  });
  // Automatisch aktivieren wenn B2Run-Template nachträglich gewählt wird.
  React.useEffect(() => {
    if (!editEvent && selectedTemplate === 'b2run') setUseSplitCapacities(true);
  }, [selectedTemplate, editEvent]);

  React.useEffect(() => {
    if (!useSplitCapacities) return;
    const d = parseInt(durchstarterCapacity, 10) || 0;
    const f = parseInt(funstarterCapacity, 10) || 0;
    const sum = d + f;
    if (sum > 0) setMaxParticipants(String(sum));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [durchstarterCapacity, funstarterCapacity, useSplitCapacities]);

  // v11.88: Helpers für Datums-Formatierung — werden von allen Demo-
  // Varianten + dem alten fillDemo geteilt.
  const fmtDatetime = (d: Date): string => {
    const pad = (n: number): string => (n < 10 ? '0' : '') + n;
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  };
  const fmtDate = (d: Date): string => {
    const pad = (n: number): string => (n < 10 ? '0' : '') + n;
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  };

  // v11.88: Reset-Helfer — setzt alle Team-, Split- und sonstigen
  // Variant-spezifischen Felder auf neutralen Default zurück, damit die
  // Demo-Varianten nicht versehentlich Zustand der vorigen Variante erben.
  const resetDemoVariantBaseState = (): void => {
    // v28.7: Demo-Vorlagen setzen immer eine Beschreibung — der
    // „Keine Beschreibung"-Schalter darf dann nicht angehakt bleiben.
    setNoDescription(false);
    // v28.20: keine Klammer-Frist aus einer vorherigen Vorlage mitschleppen.
    setKlammerDeadline('');
    setUseSplitCapacities(false);
    setSplitLabelA('Teilnehmergruppe 1');
    setSplitLabelB('Teilnehmergruppe 2');
    setDurchstarterCapacity('0');
    setFunstarterCapacity('0');
    setSplitSharedWaitlist(false);
    setTeamRegistrationEnabled(false);
    setTeamSize(4);
    setAskTeamName(false);
    setTeamPartialAllowed(false);
    setTeamOpenSlotsVisible(false);
    setTeamJoinRequiresApproval(false);
    setAskSalutation(false);
    setConfirmDialogEnabled(false); // v18.75: Sicherheitshinweis-Default
    setConfirmDialogMode('summary');
    setConfirmDialogText('');
    setSubEvents([]);
    setRemovedSavedSubs([]);
    setCustomFields([]);
    setAgenda([]);
    setTransferTimes([]);
    setLocationFilter('');
    setAudience('');
    setEventImageUrl('');
    setContactName('');
    setContactEmail('');
    setContactInfo('');
    setWaitlistEnabled(true);
    setEmailLanguage('DE');
  };

  // v24.9 (E): bestehendes Event als Vorlage übernehmen — Einstellungen + Bild
  // ins neue Formular laden (KEINE Datumswerte, KEINE Sub-Events — die legt der
  // Organizer fürs neue Event frisch fest). Das Bild wird vom (gleichen
  // SharePoint-)Anhang gefetcht und als Datei für den Re-Upload übernommen.
  const applyEventTemplate = async (ev: (typeof events)[number]): Promise<void> => {
    setTemplateLoadingId(ev.id);
    try {
      resetDemoVariantBaseState();
      setTitle(`${ev.title || ''} (Kopie)`.trim());
      setDescription(ev.description || '');
      setLocation(ev.location || '');
      setAddrStreet(ev.locationAddress?.street || '');
      setAddrHouseNo(ev.locationAddress?.houseNo || '');
      setAddrZip(ev.locationAddress?.zip || '');
      setAddrCity(ev.locationAddress?.city || '');
      setLocationFilter((ev.locationAudience || []).join(', '));
      setAudience((ev.audienceFilter || []).join(', '));
      setFilterMode(ev.filterMode || 'OR');
      if (ev.maxParticipants && ev.maxParticipants > 0) { setUnlimitedParticipants(false); setMaxParticipants(String(ev.maxParticipants)); }
      else { setUnlimitedParticipants(true); setMaxParticipants(''); }
      setWaitlistEnabled(!!ev.waitlistEnabled);
      setAskSalutation(!!ev.askSalutation);
      if (ev.agenda && ev.agenda.length > 0) setAgenda([...ev.agenda]);
      // Geteilte Kapazität übernehmen, falls vorhanden.
      if ((ev.splitLabelA || '').trim() || (ev.splitLabelB || '').trim() || (ev.durchstarterCapacity || 0) > 0 || (ev.funstarterCapacity || 0) > 0) {
        setUseSplitCapacities(true);
        setSplitLabelA(ev.splitLabelA || '');
        setSplitLabelB(ev.splitLabelB || '');
        setSplitDescA(ev.splitDescA || '');
        setSplitDescB(ev.splitDescB || '');
        setSplitHelpText(ev.splitHelpText || '');
        setSplitSectionTitle(ev.splitSectionTitle || '');
        setDurchstarterCapacity(String(ev.durchstarterCapacity || 0));
        setFunstarterCapacity(String(ev.funstarterCapacity || 0));
        setSplitSharedWaitlist(!!ev.splitSharedWaitlist);
      }
      setCustomFields((ev.eventSpecificFields || []).map(f => ({
        id: f.id, label: f.label, type: f.type, required: f.required,
        options: f.options ? [...f.options] : [], visible: true,
        ...(f.multi ? { multi: true } : {}),
        ...(f.helpText ? { helpText: f.helpText } : {}),
        ...(f.helpTextStyle === 'inline' ? { helpTextStyle: 'inline' as const } : {}),
        ...(f.showIf ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } } : {}),
        ...(f.onlyForGroup ? { onlyForGroup: f.onlyForGroup } : {}),
        ...(f.confirmLabel ? { confirmLabel: f.confirmLabel } : {}),
        ...(f.labelEn ? { labelEn: f.labelEn } : {}),
        ...(f.helpTextEn ? { helpTextEn: f.helpTextEn } : {}),
        ...(f.confirmLabelEn ? { confirmLabelEn: f.confirmLabelEn } : {}),
        ...(f.optionsEn && f.optionsEn.length > 0 ? { optionsEn: [...f.optionsEn] } : {}),
        ...(f.externalLinks && f.externalLinks.length > 0 ? { externalLinks: f.externalLinks.map(x => ({ ...x })) } : {}),
        ...(f.ccOnEmails ? { ccOnEmails: true } : {}),
        ...(f.notifyRoommate === false ? { notifyRoommate: false } : {}),
        ...(f.audienceOnly ? { audienceOnly: true } : {}),
        // v29.20 (Audit A3): auch hier fehlten Vorauswahl, Vorfilter,
        // Uhrzeit-Option und die daterange-Grenzen — die Kopie verlor sie.
        ...(f.defaultValue ? { defaultValue: f.defaultValue } : {}),
        ...(f.optionCategories && f.optionCategories.length > 0 ? { optionCategories: [...f.optionCategories] } : {}),
        ...(f.prefilterLabel ? { prefilterLabel: f.prefilterLabel } : {}),
        ...(f.withTime ? { withTime: true } : {}),
        ...(f.rangeStart ? { rangeStart: f.rangeStart } : {}),
        ...(f.rangeEnd ? { rangeEnd: f.rangeEnd } : {}),
        ...(typeof f.maxNights === 'number' && f.maxNights > 0 ? { maxNights: f.maxNights } : {}),
      })));
      // v29.20 (Audit): Die EN-Varianten oben nützen nur mit dem Schalter —
      // serializeCustomFields schreibt sie ausschließlich bei aktivem
      // bilingualFields. Ohne die Übernahme zeigte der Wizard die EN-Texte
      // der Vorlage an und der Save warf sie still weg. Ebenso übernimmt die
      // Kopie jetzt Mail-Sprache und Ausschluss-Liste — beides gehört zur
      // Konfiguration, die man mit einer Vorlage erwartet.
      setBilingualFields(!!ev.bilingualFields);
      setEmailLanguage((ev.emailLanguage || 'DE').toUpperCase() === 'EN' ? 'EN' : 'DE');
      setExcludedUsers([...(ev.excludedUsers || [])]);
      // Bild: Vorschau sofort, Datei best-effort vom SP-Anhang nachladen.
      // v29.20 (Audit): Vorher NUR gesetzt, wenn die Vorlage ein Bild hat —
      // beim Wechsel von Vorlage A (mit Bild) zu Vorlage B (ohne) blieben
      // imagePreview/imageFile von A stehen, und das neue Event bekam still
      // das Foto des falschen Events. resetDemoVariantBaseState leerte nur
      // eventImageUrl. Jetzt wird immer erst geleert.
      setImagePreview('');
      setImageFile(null);
      setImageOrigFile(null);
      setImageOrigAspect(null);
      if (ev.imageUrl) {
        setImagePreview(ev.imageUrl);
        try {
          const resp = await fetch(ev.imageUrl);
          if (resp.ok) {
            const blob = await resp.blob();
            const ext = (blob.type && blob.type.indexOf('png') >= 0) ? 'png' : 'jpg';
            setImageFile(new File([blob], `vorlage-bild.${ext}`, { type: blob.type || 'image/jpeg' }));
          }
        } catch { /* nur Vorschau, kein Re-Upload */ }
      }
      setShowTemplatePicker(false);
      setCurrentStep(0);
    } finally {
      setTemplateLoadingId(null);
    }
  };

  // v24.5: Demo-Events finden immer am NÄCHSTEN Samstag statt.
  const nextSaturdayAt = (hour: number, minute: number): Date => {
    const d = new Date();
    const day = d.getDay(); // 0=So … 6=Sa
    let add = (6 - day + 7) % 7; // Tage bis zum kommenden Samstag
    if (add === 0) add = 7;      // heute Samstag → nächster Samstag
    d.setDate(d.getDate() + add);
    d.setHours(hour, minute, 0, 0);
    return d;
  };
  // Tag(e) vor dem nächsten Samstag — für Anmelde-/Abmeldefristen.
  const beforeNextSaturday = (daysBefore: number, hour: number, minute: number): Date => {
    const d = nextSaturdayAt(hour, minute);
    d.setDate(d.getDate() - daysBefore);
    return d;
  };

  // v11.88: Vier Demo-Vorlagen — vom „Demo"-Button-Modal aufgerufen.
  // Jede Variante füllt das Formular vollständig (inkl. Reset der
  // Felder, die diese Variante NICHT setzt).
  const loadDemoStandard = (): void => {
    resetDemoVariantBaseState();
    const start = nextSaturdayAt(10, 0);
    const end = nextSaturdayAt(12, 0);
    const deadline = beforeNextSaturday(1, 23, 59);
    setTitle('Demo-Meeting Standard');
    setDescription('Beispielhaftes einfaches Meeting ohne Gruppen und ohne Sub-Events.');
    setLocation('Heinrich Campus Düsseldorf, 6. Etage');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDate(deadline));
    setLastDeregisterDate(fmtDate(deadline));
    setMaxParticipants('50');
    setUseSplitCapacities(false);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    const tDemo = Date.now();
    setCustomFields([
      { id: `cf-${tDemo}`, label: 'Essenspräferenz', type: 'select', required: true,
        options: ['Vegetarisch', 'Vegan', 'Keine Einschränkungen'], visible: true },
    ]);
    setCurrentStep(0);
  };

  const loadDemoGroups = (): void => {
    resetDemoVariantBaseState();
    const start = nextSaturdayAt(9, 0);
    const end = nextSaturdayAt(17, 0);
    const deadline = beforeNextSaturday(2, 23, 59);
    setTitle('Demo-Workshop mit Gruppen');
    setDescription('Workshop mit zwei Teilnehmer-Gruppen (Vormittag/Nachmittag) und gemeinsamer Warteliste.');
    setLocation('Deloitte Office Köln');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDate(deadline));
    setLastDeregisterDate(fmtDate(deadline));
    setMaxParticipants('50');
    setUseSplitCapacities(true);
    setSplitLabelA('Vormittag');
    setSplitLabelB('Nachmittag');
    setDurchstarterCapacity('25');
    setFunstarterCapacity('25');
    setSplitSharedWaitlist(true);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    setCurrentStep(0);
  };

  const loadDemoSubEvent = (): void => {
    resetDemoVariantBaseState();
    const start = nextSaturdayAt(9, 0);
    const end = nextSaturdayAt(17, 0);
    const deadline = beforeNextSaturday(3, 23, 59);
    setTitle('Demo-Conference mit Dinner');
    setDescription('Hauptkonferenz + abendliches Dinner als getrenntes Sub-Event mit eigener Anmeldung.');
    setLocation('Deloitte Office Hamburg');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDate(deadline));
    setLastDeregisterDate(fmtDate(deadline));
    setMaxParticipants('100');
    setUseSplitCapacities(false);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    const tDemo = Date.now();
    setCustomFields([
      { id: `cf-${tDemo}`, label: 'Hotel-Buchung', type: 'select', required: false,
        options: ['Ja, ich brauche ein Hotel', 'Nein, ich reise abends ab'], visible: true },
    ]);
    const dinnerStart = nextSaturdayAt(18, 0);
    const dinnerEnd = nextSaturdayAt(22, 0);
    setSubEvents([
      {
        id: `se-${tDemo}`,
        title: 'Networking-Dinner',
        startDate: berlinLocalToUtcIso(fmtDatetime(dinnerStart)),
        endDate: berlinLocalToUtcIso(fmtDatetime(dinnerEnd)),
        registrationDeadline: '',
        location: 'Restaurant Fischmarkt',
        description: 'Optionales Networking-Dinner im Anschluss an die Konferenz.',
        maxParticipants: 60,
        disableEmails: false,
        disableOutlook: false,
        customFields: [],
      },
    ]);
    setCurrentStep(0);
  };

  const loadDemoSubEventTeam = (): void => {
    resetDemoVariantBaseState();
    const start = nextSaturdayAt(18, 0);
    const end = nextSaturdayAt(22, 0);
    const deadline = beforeNextSaturday(5, 23, 59);
    setTitle('Demo-Kneipenquiz mit Team-Anmeldung');
    setDescription('Quizabend, bei dem ganze Teams über das Anmeldeformular angemeldet werden.');
    setLocation('Heinrich Campus Düsseldorf, 6. Etage, Dachterrasse');
    setStartDate(fmtDatetime(start));
    setEndDate(fmtDatetime(end));
    setRegistrationDeadline(fmtDate(deadline));
    setLastDeregisterDate(fmtDate(deadline));
    setMaxParticipants('80');
    setUseSplitCapacities(false);
    setWaitlistEnabled(true);
    setAskSalutation(false);
    setTeamRegistrationEnabled(true);
    setTeamSize(4);
    setAskTeamName(true);
    setTeamPartialAllowed(true);
    setTeamOpenSlotsVisible(true);
    setTeamJoinRequiresApproval(false);
    const tDemo = Date.now();
    setCustomFields([
      { id: `cf-${tDemo}`, label: 'Essenspräferenz', type: 'select', required: true,
        options: ['Vegetarisch', 'Vegan', 'Keine Einschränkungen'], visible: true },
    ]);
    const briefStart = nextSaturdayAt(17, 0);
    const briefEnd = nextSaturdayAt(17, 30);
    setSubEvents([
      {
        id: `se-${tDemo}`,
        title: 'Vorbereitungs-Briefing (Quizmaster)',
        startDate: berlinLocalToUtcIso(fmtDatetime(briefStart)),
        endDate: berlinLocalToUtcIso(fmtDatetime(briefEnd)),
        registrationDeadline: '',
        location: 'Heinrich Campus Düsseldorf, 6. Etage, Dachterrasse',
        description: 'Kurzes Briefing für die Quizmaster-Helfer vor dem Event.',
        maxParticipants: 10,
        disableEmails: false,
        disableOutlook: false,
        customFields: [],
      },
    ]);
    setCurrentStep(0);
  };

  // v11.88: Variant-Map für den Demo-Button. Key entspricht der Karten-
  // Auswahl im Modal, Value ist die Loader-Funktion oben.
  const DEMO_VARIANTS: Record<'standard' | 'groups' | 'subevent' | 'subeventTeam', () => void> = {
    standard: loadDemoStandard,
    groups: loadDemoGroups,
    subevent: loadDemoSubEvent,
    subeventTeam: loadDemoSubEventTeam,
  };


  const moveCustomField = (id: string, direction: 'up' | 'down'): void => {
    const idx = customFields.findIndex(f => f.id === id);
    if (idx < 0) return;
    const newIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (newIdx < 0 || newIdx >= customFields.length) return;
    const updated = [...customFields];
    [updated[idx], updated[newIdx]] = [updated[newIdx], updated[idx]];
    setCustomFields(updated);
  };

  // ===== Quiz helpers =====
  const addQuizQuestion = (): void => {
    setQuiz([...quiz, { id: `q-${Date.now()}`, question: '', options: ['', ''], correctIndices: [0] }]);
  };
  const removeQuizQuestion = (id: string): void => {
    setQuiz(quiz.filter(q => q.id !== id));
  };
  const updateQuizQuestion = (id: string, updates: Partial<{question: string; options: string[]; correctIndices: number[]; imageBase64: string | undefined; section: string | undefined}>): void => {
    setQuiz(quiz.map(q => q.id === id ? { ...q, ...updates } : q));
  };

  // ===== Agenda helpers =====
  const addAgendaItem = (): void => {
    setAgenda([...agenda, {
      id: `ag-${Date.now()}`,
      date: startDate ? startDate.slice(0, 10) : '',
      time: '',
      endTime: '',
      icon: 'Calendar',
      title: '',
      description: '',
    }]);
  };

  const removeAgendaItem = (id: string): void => {
    setAgenda(agenda.filter(a => a.id !== id));
  };

  const updateAgendaItem = (id: string, updates: Partial<AgendaItem>): void => {
    setAgenda(agenda.map(a => a.id === id ? { ...a, ...updates } : a));
  };

  /**
   * Persistiert die Sub-Event-Drafts nach dem Parent-Save. Seit v6.4 sind Sub-Events
   * eigene DEX_Events-Items mit gesetztem parentEventId.
   *
   * - Drafts **ohne dbId** → `createEvent({ ..., parentEventId })`
   * - Drafts **mit dbId** und Werte-Diff → `updateEvent(dbId, patch)`
   * - Initial vorhandene Sub-Event-DB-IDs, die **nicht** mehr als Draft existieren → `deleteEvent(id)` (kaskadierend inkl. Subsite + Kalendertermin)
   *
   * Alle Sub-Events erben Metadaten (Organizer, Audience, Email-Language,
   * Templates, Logos) vom Parent — eigenständig sind nur Titel, Daten, Ort und Kapazität.
   */
  // v27.11: Pro-Sub-Event-Bild persistieren (gleicher Mechanismus wie das
  // Haupt-Event-Bild: Item-Attachment hochladen + EventImageUrl-MERGE).
  // Läuft NACH createEvent/updateEvent, weil der Attachment-Upload die
  // DEX_Events-Item-Id braucht. Best-effort — ein Bild-Fehler darf den
  // Sub-Event-Save nicht blockieren.
  /**
   * v29.32: Das EIGENE Bild eines Sub-Events als Kopfbild für dessen Mails und
   * Outlook-Termin. Bisher erbte ein Sub-Event ohne eigenes Mail-Logo das Logo
   * des Hauptevents — bei einer Reihe mit unterschiedlichen Terminen (eigenes
   * Foto je Termin) kam in der Bestätigung also das falsche Bild an. Reihenfolge
   * bleibt: eigenes Mail-Logo → eigenes Event-Bild → geerbtes Logo des
   * Hauptevents. Gleiche Aufbereitung wie „Event-Foto übernehmen" beim
   * Hauptevent (auf 600 px komprimiert, damit die Zeile nicht ins
   * SharePoint-2-MB-Limit läuft).
   *
   * Der Cache verhindert, dass dasselbe gespeicherte Bild bei einem Save mit
   * vielen Sub-Events mehrfach geladen und komprimiert wird.
   */
  const subPhotoLogoCache = React.useRef<Record<string, string>>({});
  const subPhotoAsLogo = async (draft: { imageFile?: File | null; imagePreview?: string; imageRemoved?: boolean }): Promise<string> => {
    try {
      if (draft.imageRemoved) return '';
      if (draft.imageFile) return await fileToBase64(await compressImage(draft.imageFile, 600, 0.85, true));
      const prev = (draft.imagePreview || '').trim();
      if (!prev) return '';
      if (prev.indexOf('data:') === 0) return await shrinkLogoB64(prev);
      const cached = subPhotoLogoCache.current[prev];
      if (typeof cached === 'string') return cached;
      const resp = await fetch(prev, { credentials: 'include' });
      const blob = await resp.blob();
      const f = new File([blob], 'sub-event-photo.jpg', { type: blob.type || 'image/jpeg' });
      const b64 = await fileToBase64(await compressImage(f, 600, 0.85, true));
      subPhotoLogoCache.current[prev] = b64;
      return b64;
    } catch { return ''; }
  };
  const persistSubEventImage = async (subDbId: string | number | null | undefined, draft: { imageFile?: File | null; imageRemoved?: boolean }): Promise<void> => {
    const idNum = Number(subDbId);
    if (!subDbId || !isFinite(idNum) || idNum <= 0) return;
    if (!draft.imageFile && !draft.imageRemoved) return;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (!ctx) return;
      const svc = new EventService(ctx);
      if (draft.imageFile) {
        const compressed = await compressImage(draft.imageFile);
        const uploadedUrl = await svc.uploadEventImageAsAttachment(idNum, compressed);
        if (uploadedUrl) await svc.updateEventImageUrl(idNum, uploadedUrl);
      } else if (draft.imageRemoved) {
        await svc.updateEventImageUrl(idNum, '');
      }
    } catch (err) {
      console.warn('[DEX][v27.11] Sub-Event-Bild speichern fehlgeschlagen:', subDbId, err);
    }
  };

  /**
   * v28.98: `onStep` meldet nach JEDEM Sub-Event, wie viele von wie vielen
   * fertig sind. Vorher setzte der Aufrufer einmal 75 % und wartete dann auf
   * den ganzen Durchlauf — bei neun Terminen steht der Balken minutenlang
   * still und sieht aus, als haenge das Speichern.
   */
  const persistSubEventsForParent = async (
    parentEventId: string,
    onStep?: (done: number, total: number, title: string) => void,
  ): Promise<void> => {
    return await persistSubEventsForParentImpl({
      bilingualFields, childEventsOf, confirmDialog, contactEmail, createEvent, deleteEvent,
      deleteEventItemOnly, editEvent, emailLanguage, forceOutlookRecreateRef, headerImageLayoutConfig,
      headerLayoutFor, initialSubEventDbIds, initialSubEventOutlookMeta, initialSubPersistRef, isDe, isFictive,
      onlineMeetingMode, organizer, orgGetsSubInvites, outlookTeamsLink, parentTimesIso, pendingOutlookRecreateForSubEventsRef,
      persistSubEventImage, resolveTopLevelCommState, sanitizeOrganizerPairs, showAlert, shrinkLogoB64,
      subEventCalendar, subEventsRef, subPersistKey, subPhotoAsLogo, subTopGateInitialRef, subTopGateKey,
      title, updateEvent,
    }, parentEventId, onStep);
  };

  /**
   * v11.57: Tab-Wechsel im Schritt 7 (Kommunikation). Der aktuelle Step-6-
   * UI-State (emailLanguage, outlookBody, disableEmails, disableOutlook,
   * Logo-Previews, Outlook-Heading) wird in das ausgehende Tab-Slot
   * geschrieben, danach werden die Felder aus dem neuen Tab-Slot geladen.
   *  - Slot 0 = Top-Level-Event-State (die normalen `emailLanguage`,
   *    `outlookBody` etc. — also der gleiche Speicherort wie heute).
   *  - Slot N>0 = subEvents[N-1] (die Felder aus SubEventDraft).
   * Die Step-5-UI bleibt unverändert an die Top-Level-States gebunden — wir
   * spiegeln nur beim Tab-Wechsel hin und zurück.
   */
  const switchCommTab = (nextIdx: number): void => {
    if (nextIdx === activeCommTabIdx) return;
    // 1) Aktuellen UI-State in das ausgehende Slot schreiben.
    if (activeCommTabIdx > 0) {
      const fromIdx = activeCommTabIdx - 1;
      // v11.60: synchron in den Ref schreiben (siehe flushActiveCommTabToState).
      const flushed = subEventsRef.current.map((s, i) => i === fromIdx ? {
        ...s,
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        autoDeregisterOnDecline,
        inactiveHandling,
        disableOutlook,
        // v14.4: Mail-Text-Overrides pro Sub-Event mitspiegeln.
        emailTemplateOverrides: { ...emailTemplateOverrides },
      } : s);
      subEventsRef.current = flushed;
      setSubEvents(flushed);
    } else {
      // Slot 0 = Top-Level. Der UI-State wird hier direkt vom Top-Level-State
      // gehalten — kein Snapshot nötig, weil setEmailLanguage etc. den Wert
      // schon dort hält. Beim Zurück-Wechsel auf Tab 0 setzen wir die
      // Top-Level-States aus dem `topLevelCommSnapshot`-Ref (siehe unten).
      topLevelCommSnapshot.current = {
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        autoDeregisterOnDecline,
        inactiveHandling,
        disableOutlook,
        emailTemplateOverrides: { ...emailTemplateOverrides },
      };
    }
    // 2) Werte aus dem Ziel-Slot in die Step-5-UI laden.
    if (nextIdx === 0) {
      const snap = topLevelCommSnapshot.current;
      if (snap) {
        setEmailLanguage(snap.emailLanguage);
        setEmailLogoPreview(snap.emailLogoBase64 || '');
        setOutlookLogoPreview(snap.outlookLogoBase64 || '');
        setOutlookBody(snap.outlookBody || '');
        setOutlookHeading(snap.outlookHeading || '');
        setOutlookSubheading(snap.outlookSubheading || '');
        setOutlookSubject(snap.outlookSubject || '');
        setDisableEmails(!!snap.disableEmails);
        setDisableRegistrationEmail(!!snap.disableRegistrationEmail);
        setDisableCancellationEmail(!!snap.disableCancellationEmail);
        setAutoDeregisterOnDecline(!!snap.autoDeregisterOnDecline);
        setInactiveHandling(snap.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify');
        setDisableOutlook(!!snap.disableOutlook);
        setEmailTemplateOverrides(snap.emailTemplateOverrides || {});
      }
    } else {
      const sub = subEvents[nextIdx - 1];
      if (sub) {
        setEmailLanguage(sub.emailLanguage || (locale === 'de' ? 'DE' : 'EN'));
        setEmailLogoPreview(sub.emailLogoBase64 || '');
        setOutlookLogoPreview(sub.outlookLogoBase64 || '');
        setOutlookBody(sub.outlookBody || '');
        setOutlookHeading(sub.outlookHeading || sub.title || '');
        setOutlookSubheading(sub.outlookSubheading || '');
        setOutlookSubject(sub.outlookSubject || '');
        setDisableEmails(!!sub.disableEmails);
        setDisableRegistrationEmail(!!sub.disableRegistrationEmail);
        setDisableCancellationEmail(!!sub.disableCancellationEmail);
        setAutoDeregisterOnDecline(!!sub.autoDeregisterOnDecline);
        setInactiveHandling(sub.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify');
        setDisableOutlook(!!sub.disableOutlook);
        setEmailTemplateOverrides(sub.emailTemplateOverrides || {});
      }
    }
    setActiveCommTabIdx(nextIdx);
  };
  // v11.57: Snapshot des Top-Level-Step-5-States. Wird beim Wechsel auf einen
  // Sub-Event-Tab gesetzt und beim Zurückspringen wieder eingespielt.
  const topLevelCommSnapshot = React.useRef<{
    emailLanguage: string;
    emailLogoBase64: string;
    outlookLogoBase64: string;
    outlookBody: string;
    outlookHeading: string;
    outlookSubheading: string;
    outlookSubject: string;
    disableEmails: boolean;
    disableRegistrationEmail: boolean;
    disableCancellationEmail: boolean;
    autoDeregisterOnDecline: boolean;
    inactiveHandling?: 'notify' | 'autoderegister';
    disableOutlook: boolean;
    emailTemplateOverrides: Record<string, EmailOverrideEntry>;
  } | null>(null);
  // v11.57: Bevor wir submitten, müssen die Werte des aktuell sichtbaren
  // Tabs ins zugehörige Slot zurückgeschrieben werden — sonst gehen die
  // letzten Änderungen verloren.
  const flushActiveCommTabToState = (): void => {
    if (activeCommTabIdx > 0) {
      const fromIdx = activeCommTabIdx - 1;
      // v11.60: synchron in den Ref schreiben — sonst sieht die direkt
      // anschliessende Detect-/Persist-Logik noch die alte Array.
      const flushed = subEventsRef.current.map((s, i) => i === fromIdx ? {
        ...s,
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        autoDeregisterOnDecline,
        inactiveHandling,
        disableOutlook,
        emailTemplateOverrides: { ...emailTemplateOverrides },
      } : s);
      subEventsRef.current = flushed;
      setSubEvents(flushed);
    }
    // Slot 0 (Top-Level) wird ohnehin direkt von den State-Variablen gespeist
    // — kein Snapshot-Flush nötig (resolveTopLevelCommState liest auf Tab 0
    // direkt aus dem State, der Snapshot wird nur für Sub-Tab-Pfade benutzt).
  };

  /**
   * v11.93: Liefert die Top-Level-Kommunikations-Werte zuverlässig — egal auf
   * welchem Tab der User gerade in Schritt 6 steht. Bug-Hintergrund: Die
   * Logo-/Body-/Heading-States werden zwischen Top-Level und Sub-Event-Tabs
   * hin- und hergespiegelt; wenn der User auf einem Sub-Tab Speichern klickt,
   * stehen in den State-Variablen die Sub-Event-Werte — der Top-Level-Save
   * würde diese fälschlich aufs Haupt-Event schreiben (Logo, Outlook-Body
   * etc.). Diese Helper-Funktion entscheidet auf Basis von activeCommTabIdx,
   * ob die aktuellen State-Variablen schon Top-Level sind (Tab 0) oder ob aus
   * dem topLevelCommSnapshot resolved werden muss.
   */
  const resolveTopLevelCommState = (): {
    emailLanguage: string;
    emailLogoBase64: string;
    outlookLogoBase64: string;
    outlookBody: string;
    outlookHeading: string;
    outlookSubheading: string;
    outlookSubject: string;
    disableEmails: boolean;
    disableRegistrationEmail: boolean;
    disableCancellationEmail: boolean;
    autoDeregisterOnDecline: boolean;
    inactiveHandling?: 'notify' | 'autoderegister';
    disableOutlook: boolean;
    emailTemplateOverrides: Record<string, EmailOverrideEntry>;
  } => {
    if (activeCommTabIdx === 0) {
      return {
        emailLanguage,
        emailLogoBase64: emailLogoPreview,
        outlookLogoBase64: outlookLogoPreview,
        outlookBody,
        outlookHeading,
        outlookSubheading,
        outlookSubject,
        disableEmails,
        disableRegistrationEmail,
        disableCancellationEmail,
        autoDeregisterOnDecline,
        inactiveHandling,
        disableOutlook,
        emailTemplateOverrides,
      };
    }
    const snap = topLevelCommSnapshot.current;
    if (snap) return snap;
    // Fallback (sollte praktisch nicht eintreten): wir sind auf einem Sub-Tab,
    // hatten aber noch keinen Snapshot — verwenden die aktuellen State-Werte,
    // damit zumindest kein Crash entsteht.
    return {
      emailLanguage,
      emailLogoBase64: emailLogoPreview,
      outlookLogoBase64: outlookLogoPreview,
      outlookBody,
      outlookHeading,
      outlookSubheading,
      outlookSubject,
      disableEmails,
      disableRegistrationEmail,
      disableCancellationEmail,
      autoDeregisterOnDecline,
      inactiveHandling,
      disableOutlook,
      emailTemplateOverrides,
    };
  };

  /**
   * v28.29: Nach „Event-Foto übernehmen" auf dem HAUPTEVENT-Tab fragen, ob das
   * Bild auch für alle {childTermPlural} gelten soll. Sub-Events haben eigene
   * Mails und eigene Outlook-Termine; ohne diese Abfrage musste der Organizer
   * das Foto in jedem Tab einzeln übernehmen — und merkte den Unterschied erst,
   * wenn die Sub-Event-Termine mit dem Standardlogo bei den Teilnehmern landeten.
   * Sagt er Ja, wird das Bild in JEDEN Sub-Event-Draft geschrieben (überschreibt
   * auch bereits gesetzte eigene Bilder — das ist der Sinn der Frage); sagt er
   * Nein, greift weiterhin die stille Vererbung für Sub-Events OHNE eigenes Bild.
   */
  const offerLogoToSubEvents = async (kind: 'email' | 'outlook', b64: string): Promise<void> => {
    if (!b64 || activeCommTabIdx !== 0) return;
    const named = subEventsRef.current.filter(x => x.title && x.title.trim());
    if (named.length === 0) return;
    const what = kind === 'email'
      ? (isDe ? 'in den E-Mails' : 'in the emails')
      : (isDe ? 'im Outlook-Termin' : 'in the Outlook appointment');
    const term = childTermPlural || (isDe ? 'Sub-Events' : 'sub-events');
    const ok = await confirmDialog(
      isDe
        ? `Das Event-Foto steht jetzt ${what} des Hauptevents.\n\nSoll es auch für die ${named.length} ${term} gelten?\n\n${named.map(x => `• ${x.title.trim()}`).join('\n')}\n\nJa = alle bekommen dieses Bild (ein dort bereits hinterlegtes eigenes Bild wird ersetzt). Nein = die ${term} behalten ihre Einstellung; wer kein eigenes Bild hat, erbt ohnehin das des Hauptevents.`
        : `The event photo is now used ${what} of the main event.\n\nApply it to the ${named.length} ${term} as well?\n\n${named.map(x => `• ${x.title.trim()}`).join('\n')}\n\nYes = all of them get this image (any own image set there is replaced). No = they keep their setting; those without an own image inherit the main event's anyway.`,
      { confirmLabel: isDe ? `Ja, für alle ${term}` : `Yes, for all ${term}` },
    );
    if (!ok) return;
    const next = subEventsRef.current.map(x => {
      if (!x.title || !x.title.trim()) return x;
      return kind === 'email' ? { ...x, emailLogoBase64: b64 } : { ...x, outlookLogoBase64: b64 };
    });
    subEventsRef.current = next;
    setSubEvents(next);
    showAlert(
      isDe ? `Bild für ${named.length} ${term} übernommen — beim Speichern werden deren Termine/Mails mit aktualisiert.`
        : `Image applied to ${named.length} ${term} — their appointments/emails are updated on save.`,
      { variant: 'success' },
    );
  };

  /**
   * v28.29: Welches Bild steht am Ende WIRKLICH im Kopf der Mail bzw. des
   * Outlook-Termins? Die Vorschau in Schritt 23/24 hat bisher stumpf
   * `logo || imagePreview` gezeigt — also das Event-Foto, auch wenn als Logo
   * gar nichts hinterlegt war. Auf Sub-Event-Tabs war das der eigentliche
   * Stolperstein: Die Vorschau zeigte das Foto, gespeichert wurde aber der
   * Standard-Orb. Jetzt liefert dieser Helper genau das Bild, das der Save
   * schreibt, plus einen Hinweis, woher es stammt.
   */
  const effectiveHeaderImage = (kind: 'email' | 'outlook', own: string): { src: string; note: string } => {
    if (own) return { src: own, note: '' };
    // Sub-Event-Tab: erbt seit v28.29 das Kopfbild des Hauptevents.
    if (activeCommTabIdx > 0) {
      const snap = topLevelCommSnapshot.current;
      const parentLogo = snap ? (kind === 'email' ? snap.emailLogoBase64 : snap.outlookLogoBase64) : '';
      if (parentLogo) {
        return {
          src: parentLogo,
          note: isDe
            ? 'Wird vom Hauptevent übernommen — lade oben ein eigenes Bild hoch, wenn dieser Termin ein anderes zeigen soll.'
            : 'Inherited from the main event — upload your own image above if this appointment should show a different one.',
        };
      }
    }
    const orb = getCachedOrbBase64() || '';
    return {
      src: orb,
      note: isDe
        ? 'Kein eigenes Bild hinterlegt — es wird das Deloitte-Standardlogo verwendet. Das Event-Foto wandert NICHT automatisch hierher; dafür auf „Event-Foto übernehmen" klicken.'
        : 'No custom image set — the Deloitte default logo is used. The event photo does not move here automatically; click „Use event photo" for that.',
    };
  };

  /**
   * Save-Side-Sanity: Organizer-Names und -Emails 1:1 paaren bevor sie nach SP
   * geschrieben werden. Pairs ohne BEIDE (Name + Email) fallen raus — verhindert
   * dass eine Mismatch-State (z.B. „Spiegel, Mirjam" gepaart mit
   * „egenctuerk@deloitte.de") in DEX_Events landet. Bisher wurden organizer und
   * organizerEmails unabhängig serialisiert, dadurch konnten Drift-States aus
   * Closure-Bugs / Edit-Pfaden / Move-Bugs in die Persistenz durchschlagen.
   *
   * Returnt sauber serialisierte Strings (`Organizer` mit '; '-Trenner,
   * `OrganizerEmail` mit ';'-Trenner) — exakt das Format das DEX_Events erwartet
   * und der OutlookEventCreate-Flow + DEX_SEND_MAIL als Recipient-Liste lesen.
   */
  const sanitizeOrganizerPairs = React.useCallback((): { orgString: string; orgEmailString: string; droppedCount: number } => {
    const names = (organizer || '').split(';').map(s => s.trim());
    const emails = (organizerEmails || []).map(e => (e || '').trim());
    const max = Math.max(names.length, emails.length);
    const pairs: Array<{ n: string; e: string }> = [];
    let dropped = 0;
    for (let i = 0; i < max; i++) {
      const n = (names[i] || '').trim();
      const e = (emails[i] || '').trim();
      if (n && e) pairs.push({ n, e });
      else if (n || e) dropped++;
    }
    return {
      orgString: pairs.map(p => p.n).join('; '),
      orgEmailString: pairs.map(p => p.e).join(';'),
      droppedCount: dropped,
    };
  }, [organizer, organizerEmails]);

  const handleSubmitInner = async (): Promise<void> => {
    return await runWizardSubmit({
      activeFrom, addrCity, addrHouseNo, addrStreet, addrZip, agenda,
      allDay, allowAttendeeUpload, askSalutation, askTeamName, assistantsCanSee, attendeeUploadHint,
      attendeeUploadLabel, audience, berlinLocalToUtcIso, bilingualFields, billingPiggyback, bundledComm,
      childEventsOf, childGender, childTermPlural, childTermSingular, computeFormSnapshot, confirmDialog,
      confirmDialogEnabled, confirmDialogMode, confirmDialogText, contactEmail, contactInfo, contactName,
      contactOrganizerEmail, coOrganizerEmails, coOrganizerNames, createdEventIdRef, createEvent, currentUser,
      customFields, deadlineToEndOfDayIso, description, disableOutlook, documents, DRAFT_KEY,
      durchstarterCapacity, durchstarterRequiresProof, durchstarterStartblock, editEvent, effTeamsLink, emailLanguage,
      endDate, eventImageUrl, eventType, excludedUsers, filterMode, funstarterCapacity,
      funstarterStartblock, getGroupMembers, getLastEventUpdateError, headerImageLayoutConfig, headerLayoutFor, hiddenOrganizerEmails,
      hideOrganizer, hideOrganizerIndividualOnly, imageBanner, imageDisplay, imageFile, imageOrigAspect,
      imageOrigFile, initialDocumentNames, initialFormSnapshotRef, initialOrgGetsSubInvitesRef, initialSubEventDbIds, isB2runTemplate,
      isDe, isEditMode, isFictive, klammerDeadline, lastDeregisterDate, lastDraftJsonRef,
      location, locationFilter, mainCommDisabledAck, mainEventLabel, mainEventLabelMode, maxParticipants,
      noCancelAfterDeadline, noDescription, notifyAdminsExternalAudienceAccess, notifyNewCoOrganizers, notifyOrgCancelMode, notifyOrgRegisterFromDate,
      notifyOrgRegisterMode, onlineMeetingMode, organizer, organizerDisplayLarge, organizerEmails, orgGetsSubInvites,
      outlookEndOverride, outlookLocationOverride, outlookStartOverride, outlookTeamsLink, pendingOutlookDirtyWriteRef, pendingOutlookDirtyWriteRefs,
      pendingOutlookUpdateForSubEventsRef, pendingOutlookUpdateForTopRef, pendingSuccessDispatchRef, persistSubEventsForParent, previewBeforeActive, qrScannerEmails,
      qrScannerNames, quiz, quizClusterSize, refreshEvents, registrationDeadline, registrationLanguage,
      regRuleEnabled, requestCoOrganizerApprovals, requireSubEventSelection, resolveTopLevelCommState, sanitizeOrganizerPairs, selectedEventId,
      setDraftSavedAt, setError, setImageUploadError, setIsSubmitting, setNavigationGuard, setPendingDraft,
      setPendingSuccessDispatch, setProgress, setProgressLabel, setRemovedSavedSubs, setShowSummaryModal, showAlert,
      showAsFree, shrinkLogoB64, splitDescA, splitDescB, splitDisplayOrderReversed, splitHelpText,
      splitLabelA, splitLabelB, splitSectionTitle, splitSharedWaitlist, startDate, subDeadlineRulePiggyback,
      subEventCalendar, subEventOpenRulePiggyback, subEventSingleChoice, subEventsOnlyMode, subEventsOptIn, subEventsRef,
      teamJoinRequiresApproval, teamMembersCannotCreate, teamOpenSlotsVisible, teamPartialAllowed, teamRegistrationEnabled, teamSize,
      teamTermPlural, teamTermSingular, testTeamEmails, testTeamNames, title, transferTimes,
      unlimitedParticipants, updateEvent, userCancelAllowed, useSplitCapacities, visAllSubsPiggyback, waitlistEnabled,
      wizardImgAspect,
    });
  };

  /**
   * v28.71 BUG-FIX: Zwei Riegel gegen versehentlich doppelt angelegte Events.
   *
   * Gemeldet wurde: „bei jedem Speichern wird das Event neu gespeichert" — im
   * Postfach landeten mehrere „Event angelegt"-Mails zum selben Titel. Ursache
   * sind zwei Luecken im Anlege-Pfad:
   *
   * 1. KEIN Re-Entry-Schutz. `handleSubmit` hatte weder eine
   *    `isSubmitting`-Abfrage noch einen Latch, und der Speichern-Knopf wird
   *    nicht deaktiviert. `setIsSubmitting(true)` wirkt erst mit dem nächsten
   *    Render — zwei schnelle Klicks liefen also beide komplett durch und
   *    legten zwei Events an. Ein Ref greift dagegen synchron.
   * 2. NACH dem Anlegen blieb der Wizard im Anlege-Modus. `isEditMode` hängt
   *    an `currentPage === 'edit-event'`; der Wizard verlaesst die Seite aber
   *    erst, wenn das Zusammenfassungs-Fenster geschlossen wird. In diesem
   *    Fenster war `setIsSubmitting(false)` bereits gesetzt — ein weiterer
   *    Klick auf „Speichern" lief erneut in den CREATE-Zweig und legte ein
   *    zweites Event an, statt das eben erstellte zu aktualisieren.
   *
   * Der zweite Klick ist jetzt kein stiller Doppel-Anlegevorgang mehr, sondern
   * bietet an, das bereits angelegte Event zum Bearbeiten zu öffnen.
   */
  const submitInFlightRef = React.useRef<boolean>(false);
  const createdEventIdRef = React.useRef<string>('');
  const handleSubmit = async (): Promise<void> => {
    if (submitInFlightRef.current) return;
    if (!isEditMode && createdEventIdRef.current) {
      const openIt = await confirmDialog(
        isDe
          ? `„${title}" wurde in diesem Durchlauf bereits angelegt. Erneutes Speichern würde ein ZWEITES Event mit denselben Daten erzeugen.\n\nMöchtest du stattdessen das bereits angelegte Event zum Bearbeiten öffnen?`
          : `„${title}" has already been created in this session. Saving again would create a SECOND event with the same data.\n\nDo you want to open the existing event for editing instead?`,
        { confirmLabel: isDe ? 'Event öffnen' : 'Open event' },
      );
      if (openIt) {
        setNavigationGuard(null);
        navigate('edit-event', createdEventIdRef.current);
      }
      return;
    }
    submitInFlightRef.current = true;
    try {
      await handleSubmitInner();
    } finally {
      submitInFlightRef.current = false;
      // v28.98: In JEDEM Fall zuruecknehmen — auch wenn handleSubmitInner
      // früh aussteigt (Validierung, Fehler) oder wirft. Sonst bliebe
      // „Zurück" dauerhaft gesperrt.
      setSaveInProgress(false);
    }
  };

  // v11.57: detect ob beim aktuellen Edit-State Outlook-relevante Änderungen
  // anstehen (Title, Start, End oder OutlookBody). Vergleich gegen den
  // Mount-Snapshot.
  // v11.63: liefert pro betroffenem Event (Hauptevent + Sub-Events) eine
  // Liste der konkret geänderten Felder, damit der User im Modal pro
  // Eintrag sehen kann, was sich wirklich verändert hat.
  const detectOutlookRelevantChanges = (): { items: OutlookConfirmItem[] } => {
    return detectOutlookRelevantChangesImpl({
      activeCommTabIdx, addrCity, addrHouseNo, addrStreet, addrZip, allDay,
      berlinLocalToUtcIso, childEventsOf, editEvent, endDate, headerImageLayout, initialHeaderImageLayoutRef,
      initialOutlookSnapshot, location, onlineMeetingChanged, organizer, outlookBody, outlookEndOverride,
      outlookLocationOverride, outlookStartOverride, resolveTopLevelCommState, showAsFree, startDate, subEventCalendar,
      subEventsOnlyMode, subEventsRef, title,
    });
  };

  // v11.57: Wrapper-Funktion für den Save-Button. Im Edit-Modus mit
  // Outlook-relevanter Änderung wird das Confirm-Modal gezeigt. Sonst
  // direkt handleSubmit.
  const attemptSubmit = (): void => {
    // Aktuelle Tab-Werte zurück ins jeweilige Slot schreiben, damit
    // beim handleSubmit nichts verloren geht.
    flushActiveCommTabToState();
    if (!isEditMode || !editEvent) {
      pendingOutlookDirtyWriteRef.current = null;
      pendingOutlookDirtyWriteRefs.current = {};
      pendingOutlookUpdateForTopRef.current = false;
      pendingOutlookUpdateForSubEventsRef.current = [];
      pendingOutlookRecreateForSubEventsRef.current = [];
      handleSubmit().catch(() => { /* Errors werden in handleSubmit gesetzt */ });
      return;
    }
    const det = detectOutlookRelevantChanges();
    if (det.items.length > 0) {
      // v11.63: Snapshot der Items + leerer Check-Map (alle false). Pro
      // Event entscheidet der Organizer einzeln im Modal.
      setOutlookConfirmItems(det.items);
      setOutlookConfirmChecks({});
      setOutlookConfirmOpen(true);
      return;
    }
    // Keine Outlook-relevante Änderung — dirty-Flag nicht anfassen.
    pendingOutlookDirtyWriteRef.current = null;
    pendingOutlookDirtyWriteRefs.current = {};
    // Wenn der User den expliziten Step-5-Schalter „Outlook-Termin
    // aktualisieren" angehakt hat, überschreibt das die Modal-Logik
    // und triggert ein manuelles UpdateEvent für das Top-Level — auch
    // wenn die Detect-Heuristik nichts Outlook-relevantes gefunden hat.
    pendingOutlookUpdateForTopRef.current = !!triggerOutlookUpdate;
    pendingOutlookUpdateForSubEventsRef.current = [];
    pendingOutlookRecreateForSubEventsRef.current = [];
    handleSubmit().catch(() => { /* */ });
  };

  // v11.57: Confirm-Modal-Handler.
  // v11.63: Liest aus outlookConfirmChecks ab, welche Events der Organizer
  // angehakt hat. Angehakte Events bekommen UpdateEvent + OutlookDirty=false,
  // nicht angehakte (aber im Detect-Items gelistete) bekommen
  // OutlookDirty=true. Events ausserhalb des Detect-Items bleiben unberührt.
  const confirmOutlookSave = (): void => {
    setOutlookConfirmOpen(false);
    const topId = editEvent ? editEvent.id : '';
    const topItem = outlookConfirmItems.find(it => it.kind === 'top');
    const subItems = outlookConfirmItems.filter(it => it.kind === 'sub');
    const topChecked = !!topItem && !!outlookConfirmChecks[topItem.eventId];
    // v11.69: Angehakte Sub-Events trennen in:
    //  - `normalUpdateSubIds`: Sub-Event hat bereits einen Outlook-Termin →
    //    DEX_Outlook 'UpdateEvent' in die Queue schreiben (bestehender Pfad).
    //  - `recreateSubIds`: Sub-Event hat noch keinen Outlook-Termin
    //    (`noOutlookYet`) → DEX_Events-Item per `deleteEventItemOnly` löschen
    //    und mit `existingSubsiteUrl` neu anlegen, damit der
    //    DEX_CreateOutlookEvent-Flow triggert. Teilnehmer-Subsite + Liste
    //    bleiben unangetastet erhalten.
    const checkedSubItems = subItems.filter(it => !!outlookConfirmChecks[it.eventId]);
    const normalUpdateSubIds = checkedSubItems.filter(it => !it.noOutlookYet).map(it => it.eventId);
    const recreateSubIds = checkedSubItems.filter(it => !!it.noOutlookYet).map(it => it.eventId);
    pendingOutlookUpdateForTopRef.current = topChecked;
    pendingOutlookUpdateForSubEventsRef.current = normalUpdateSubIds;
    pendingOutlookRecreateForSubEventsRef.current = recreateSubIds;
    // Pro Event-ID den OutlookDirty-Schreibwert vormerken.
    // v11.69: noOutlookYet-Items werden — egal ob angehakt oder nicht — NICHT
    // dirty markiert. Bei angehakt erfolgt ein Recreate (neues Item hat von
    // Haus aus OutlookDirty=false), bei nicht angehakt existiert immer noch
    // kein Outlook-Termin der "aus-Sync" sein könnte → Marker wäre falsch.
    const dirtyMap: Record<string, boolean> = {};
    for (const it of outlookConfirmItems) {
      if (it.noOutlookYet) continue;
      dirtyMap[it.eventId] = !outlookConfirmChecks[it.eventId];
    }
    pendingOutlookDirtyWriteRefs.current = dirtyMap;
    // Top-Level kompatibel halten: wenn das Top-Event im Modal war, wird
    // OutlookDirty entsprechend gesetzt; sonst null = nicht anfassen.
    if (topItem) {
      pendingOutlookDirtyWriteRef.current = !topChecked;
    } else {
      pendingOutlookDirtyWriteRef.current = null;
    }
    // setTriggerOutlookUpdate steuert in handleSubmit, ob der Top-Level-
    // Outlook-Branch überhaupt betreten wird. v11.63: nur true wenn das
    // Top-Event angehakt wurde ODER mindestens ein Sub-Event angehakt
    // wurde (damit der Sub-Event-Branch im handleSubmit getroffen wird).
    setTriggerOutlookUpdate(topChecked || normalUpdateSubIds.length > 0 || recreateSubIds.length > 0);
    // Verhindern dass topId als „angehakt" interpretiert wird ohne Modal.
    void topId;
    handleSubmit().catch(() => { /* */ });
  };
  const cancelOutlookSave = (): void => {
    setOutlookConfirmOpen(false);
    // Nichts speichern — User bleibt im Wizard.
  };

  // v11.57: bei Sub-Event-Anzahl-Änderung Tab sicher in Range halten.
  // v27.11 BUG-FIX: NICHT direkt setActiveCommTabIdx(0) — das ließ die
  // Step-6-State-Variablen (Mail-Sprache, Outlook-Body, Logos, disableEmails
  // etc.) auf den SUB-EVENT-Werten stehen, die dann als Tab 0 = Top-Level
  // galten. Beim nächsten Speichern wurden die Sub-Event-Kommunikations-
  // Einstellungen still aufs Haupt-Event geschrieben (Datenverlust beim
  // Abschalten des Sub-Event-Toggles). switchCommTab(0) stellt den
  // Top-Level-Snapshot korrekt wieder her.
  React.useEffect(() => {
    if (activeCommTabIdx > subEvents.length) {
      switchCommTab(0);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [subEvents.length, activeCommTabIdx]);
  // v15.0: gleiche Range-Garantie für die neuen Tab-Sets in den
  // Steps 3 (Ort), 4 (Kapazität) und 6 (Felder).
  React.useEffect(() => {
    if (activeLocationTabIdx > subEvents.length) setActiveLocationTabIdx(0);
    if (activeCapacityTabIdx > subEvents.length) setActiveCapacityTabIdx(0);
    if (activeFieldsTabIdx > subEvents.length) setActiveFieldsTabIdx(0);
  }, [subEvents.length, activeLocationTabIdx, activeCapacityTabIdx, activeFieldsTabIdx]);

  // v15: Templates laden wenn Step 6 (Kommunikation, currentStep === 6) erreicht
  // wird. Index hat sich verschoben, weil Team-Anmeldung jetzt NACH Kommunikation
  // kommt (siehe steps-Array).
  // WICHTIG: Dieser useEffect MUSS vor dem early return (if submitted) stehen,
  // da React die gleiche Anzahl Hooks bei jedem Render erwartet (Rules of Hooks).
  React.useEffect(() => {
    // v29.21 (Audit B1): Kommunikation ist Index 5 — der alte Vergleich mit 6
    // (heute Team-Anmeldung) lud die Vorlagen erst EINEN Schritt zu spät;
    // wer sich neu durchklickte, sah im Kommunikations-Schritt leere
    // Betreff-/Text-Defaults. >= statt ===, damit auch ein Direktsprung auf
    // einen späteren Schritt (Kreis-Klick) die Vorlagen nachlädt.
    if (currentStep >= 5 && emailTemplates.length === 0) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (ctx) {
        const svc = new EventService(ctx);
        svc.getAllEmailTemplates().then(setEmailTemplates).catch(() => { /* Templates nicht verfügbar */ });
      }
    }
  }, [currentStep]);

  // v22.6: Sub-Events/Sub-Sections erben standardmäßig die Sichtbarkeit der
  // Klammer / des Hauptevents. Beim Öffnen eines bestehenden Events werden
  // Sub-Events OHNE eigene Sichtbarkeit einmalig mit Standortfilter +
  // Mailverteiler + Verknüpfung des Hauptevents vorbelegt. Sub-Events, die
  // bereits eine eigene Sichtbarkeit haben, bleiben unangetastet (werden NICHT
  // automatisch überschrieben). Danach pro Sub-Event frei änderbar.
  const subVisibilityInheritedRef = React.useRef(false);
  React.useEffect(() => {
    if (subVisibilityInheritedRef.current) return;
    if (!editEvent || subEvents.length === 0) return;
    subVisibilityInheritedRef.current = true;
    const parentLoc = locationFilter.trim();
    const parentAud = audience.trim();
    if (!parentLoc && !parentAud) return; // Hauptevent/Klammer ohne Filter → nichts zu erben
    setSubEvents(prev => prev.map(se => {
      const seHasOwn = (se.locationFilter || '').trim() || (se.audience || '').trim();
      if (seHasOwn) return se; // eigene Sichtbarkeit → nicht überschreiben
      return { ...se, locationFilter, audience, filterMode };
    }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editEvent, subEvents.length, locationFilter, audience, filterMode]);

  /**
   * v28.78: Die globale Scope-Karte unter der Schritt-Leiste.
   *
   * Zwei Zustände, damit die Leiste nicht bei jedem Schritt verschwindet und
   * wieder auftaucht (das Springen war schlimmer als der Nutzen):
   *  - AUSWAHL: in den Schritten, die pro Sub-Event konfiguriert werden.
   *  - GILT-FÜR-ALLE: in allen anderen. Gleiche Position, gleicher Rahmen,
   *    aber ohne Auswahl und mit dem Satz, dass dieser Schritt für das
   *    gesamte Event gilt. Das erklärt das Modell nebenbei mit.
   * Ohne Sub-Events wird gar nichts gezeigt — dann gibt es nichts zu wählen.
   */
  const renderGlobalScopeBar = (): React.ReactElement | null => {
    if (subEvents.length === 0) return null;
    const named = subEvents.filter(s => (s.title || '').trim());
    // v29.21 (Audit): Nicht mehr verstecken, wenn ein Sub-Reiter aktiv ist.
    // Sequenz vorher: „Hinzufügen" (Draft ohne Titel) → „Bearbeiten"
    // (setScope(1)) → die Leiste war null, die Sub-Event-Liste hängt an
    // activeScopeIdx === 0 — keine Bedienung mehr, um zurück auf die Klammer
    // zu kommen. Die Reiter tragen für unbenannte Drafts den Fallback
    // „Sub-Event ohne Titel".
    if (named.length === 0 && activeScopeIdx === 0) return null;
    const applies = SCOPE_AWARE_STEPS.indexOf(currentStep) >= 0;
    const scopeIdx = Math.min(activeScopeIdx, subEvents.length);
    const mainLabel = `${subEventsOnlyMode ? (isDe ? 'Klammer' : 'Bracket') : (isDe ? 'Haupt-Event' : 'Main event')}: ${title || (isDe ? 'Ohne Titel' : 'Untitled')}`;
    // v28.90: Die Karte war grün getönt und mit grünem Rand abgesetzt — sie las
    // sich dadurch wie ein Status („hier stimmt etwas") statt wie das, was sie
    // ist: eine Navigation. Grün bleibt der aktiven Auswahl vorbehalten.
    // Ausserdem `overflow: hidden` (plus `minWidth: 0` weiter innen): Die
    // Reiter-Reihe schob sich bei vielen Sub-Events über den rechten Kartenrand
    // hinaus — Flex-Kinder haben `min-width: auto`, die Scroll-Fläche konnte
    // ihren Container also aufblähen.
    // v28.91: …und ganz ohne eigene Fläche. Die weiße Karte auf grauem Grund
    // war immer noch ein Kasten, der um Aufmerksamkeit konkurriert; die Reiter
    // selbst tragen ihre Form bereits. Transparent, nur Abstand.
    return (
      <div id="dex-scope-bar" style={{
        margin: '18px 0 0', padding: '12px 0 14px', borderRadius: 0,
        background: 'transparent',
        border: 'none',
        overflow: 'hidden',
      }}>
        {applies ? (
          renderPerEventTabStrip(
            scopeIdx,
            setScope,
            mainLabel,
            isDe ? 'Event-Ebene wechseln' : 'Switch event level',
          )
        ) : (
          <>
            <div style={{
              fontSize: '0.78rem', fontWeight: 700, letterSpacing: '0.03em',
              textTransform: 'uppercase', color: 'var(--dex-gray-500)', marginBottom: 6,
            }}>
              {isDe ? 'Welches (Sub-)Event bearbeitest du gerade?' : 'Which (sub-)event are you editing?'}
            </div>
            <div style={{
              display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap',
              padding: '9px 14px', borderRadius: 10,
              background: '#fff', border: '1px dashed var(--dex-gray-300)',
              fontSize: '0.84rem', color: 'var(--dex-gray-700)',
            }}>
              <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>
                {isDe ? 'Dieser Schritt gilt für das gesamte Event' : 'This step applies to the entire event'}
              </strong>
              <span style={{ color: 'var(--dex-gray-600)' }}>
                {isDe
                  ? `— ${subEventsOnlyMode ? 'Klammer' : 'Haupt-Event'} und alle ${named.length} ${named.length === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} gemeinsam. Eine Auswahl gibt es hier nicht.`
                  : `— ${subEventsOnlyMode ? 'bracket' : 'main event'} and all ${named.length} sub-events together. There is nothing to pick here.`}
              </span>
            </div>
          </>
        )}
      </div>
    );
  };

  // v15.6: Hinweis-Banner für den Hauptevent-Tab in den Steps 3/4/5, wenn
  // subEventsOnlyMode aktiv ist. Der Hauptevent ist dann nicht buchbar — die
  // Einstellungen aus diesem Tab werden zur Laufzeit ignoriert. Der Tab bleibt
  // sichtbar (Konsistenz mit den restlichen Steps), wird aber ausgegraut, und
  // ein gelber Hinweis-Banner sagt explizit, dass der Organizer auf einen
  // [childTermPlural]-Tab wechseln soll, um die Konfiguration zu pflegen.
  //
  // WICHTIG: kein Hook — reines, render-loses Helper. Damit fügt diese
  // Funktion keinen zusätzlichen useState/useEffect/useMemo-Call hinzu und
  // verletzt die Rules of Hooks nicht (Rules of Hooks: alle Hooks vor dem
  // early return weiter unten).
  const renderHauptGreyoutBanner = (): React.ReactElement | null => {
    if (!subEventsOnlyMode) return null;
    const termPlural = (childTermPlural || (isDe ? 'Sub-Events' : 'sub-events')).trim() || (isDe ? 'Sub-Events' : 'sub-events');
    return (
      <WizardHint
        isDe={isDe}
        title={isDe ? 'Anmeldung läuft über die Event-Sections' : 'Registration runs via the event sections'}
        style={{ marginBottom: 16 }}
      >
        <div>
          {isDe ? (
            <>
              Du hast für dieses Event eingestellt, dass sich Teilnehmer <strong>nur für die {termPlural} (Sub-Sections) anmelden</strong> — die Klammer selbst ist <strong>nicht buchbar</strong>, sie fasst die {termPlural} nur zusammen. (Eingestellt in Schritt 1 „Grundlagen“ unter „Wie sollen sich Teilnehmer anmelden?“ → „Nur für {termPlural}“.) <strong>Entscheidend ist die Sichtbarkeit direkt hier unten</strong> (Standortfilter + Mailverteiler): Sie legt fest, <strong>wer das Event überhaupt sieht und sich für die {termPlural} anmelden kann</strong> — die {termPlural} übernehmen diese Sichtbarkeit standardmäßig. <strong>Plätze und Fristen</strong> stellst du nicht hier auf der Klammer ein, sondern <strong>je {childTermSingular || 'Sub-Event'}-Tab</strong> (wie viele Plätze, bis wann an-/abmelden) — die Kapazitäts- und Fristen-Felder weiter unten sind darum ausgegraut.
            </>
          ) : (
            <>
              You configured this event so that participants <strong>register only for the {termPlural} (sub-sections)</strong> — the bracket itself is <strong>not bookable</strong>, it only groups the {termPlural}. (Set in step 1 “Basics” under “How should participants register?” → “Sub-{termPlural} only”.) <strong>The decisive setting is the visibility right below this note</strong> (location filter + mailing lists): it determines <strong>who can see the event at all and register for the {termPlural}</strong> — the {termPlural} inherit this visibility by default. You set <strong>seats and deadlines</strong> not here on the bracket but <strong>per {childTermSingular || 'sub-event'} tab</strong> (how many seats, until when to register/cancel) — which is why the capacity and deadline fields further down are greyed out.
            </>
          )}
        </div>
      </WizardHint>
    );
  };

  // v22.22: Live-Zusammenfassung der aktuell eingestellten Sichtbarkeit —
  // gleiche Klartext-Logik wie die „Nächste Schritte"-Box im Organizer
  // Center. Wird unter der Sichtbarkeits-Überschrift gerendert (Hauptevent/
  // Klammer UND je Sub-Event-Tab) und aktualisiert sich live mit der Auswahl.
  // Render-loses Helper, kein Hook (Rules of Hooks).
  const renderVisibilitySummaryBox = (
    locList: string[],
    audienceStr: string,
    mode: 'AND' | 'OR',
    excludedCount: number
  ): React.ReactElement => {
    const locs = (locList || []).filter(Boolean);
    const auds = (audienceStr || '').split(',').map(s => s.trim()).filter(Boolean);
    // v28.76: Klartext statt Stichworten. Vorher stand hier „Sichtbar für
    // 1 Verteiler/Personen." — grammatisch schief und inhaltlich unklar
    // (1 Verteiler? 1 Person? beides?). Jetzt ein ganzer Satz, der sagt, WER
    // das Event sieht.
    let text: string;
    if (locs.length === 0 && auds.length === 0) {
      text = isDe
        ? 'Das Event sehen alle Mitarbeiter von Deloitte Deutschland.'
        : 'Everyone at Deloitte Germany can see this event.';
    } else {
      const parts: string[] = [];
      if (locs.length) {
        parts.push(isDe
          ? (locs.length === 1 ? `Mitarbeiter am Standort ${locs[0]}` : `Mitarbeiter an den Standorten ${locs.join(', ')}`)
          : (locs.length === 1 ? `employees at location ${locs[0]}` : `employees at the locations ${locs.join(', ')}`));
      }
      if (auds.length) {
        parts.push(isDe
          ? (auds.length === 1 ? 'die Mitglieder des hinterlegten Verteilers bzw. die hinterlegte Person' : `die Mitglieder der ${auds.length} hinterlegten Verteiler bzw. Personen`)
          : (auds.length === 1 ? 'the members of the selected distribution list or the selected person' : `the members of the ${auds.length} selected distribution lists / people`));
      }
      const joiner = parts.length > 1
        ? (mode === 'AND' ? (isDe ? ' und gleichzeitig ' : ' and at the same time ') : (isDe ? ' oder ' : ' or '))
        : '';
      text = (isDe ? 'Das Event sehen nur ' : 'Only ') + parts.join(joiner) + (isDe ? '.' : ' can see this event.');
    }
    return (
      <div style={{
        marginTop: 10, padding: '8px 12px', borderRadius: 8,
        background: 'rgba(134,188,37,0.07)', border: '1px solid var(--dex-green, #86bc25)',
        fontSize: '0.78rem', color: 'var(--dex-gray-700)', lineHeight: 1.5,
      }}>
        <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>
          {isDe ? 'Aktuell eingestellt: ' : 'Currently configured: '}
        </strong>
        {text}
        {excludedCount > 0 && (
          <> {isDe
            ? `${excludedCount} Person${excludedCount === 1 ? ' ist' : 'en sind'} ausgeschlossen.`
            : `${excludedCount} ${excludedCount === 1 ? 'person is' : 'people are'} excluded.`}</>
        )}
      </div>
    );
  };

  /**
   * v28.76: Widerspruch zwischen Klammer und Sub-Events benennen.
   *
   * Der Zugang läuft IMMER über die Klammer: Wer sie nicht sieht, kommt an
   * kein Sub-Event — und die Sub-Events schränken danach weiter ein. Daraus
   * folgen zwei Zustände, die in der Maske bisher unsichtbar waren:
   *
   *  a) Klammer offen, alle Sub-Events auf denselben Standort → jeder sieht
   *     das Event, findet darin aber nichts zum Anmelden. Sieht nach einem
   *     Fehler aus, ist aber die logische Folge der Einstellungen.
   *  b) Ein Sub-Event lässt mehr zu als die Klammer → der Überschuss kommt
   *     nie an, weil die Klammer vorher filtert. Die Einstellung im
   *     Sub-Event wirkt dann schlicht nicht.
   *
   * Beides wird hier gemeldet, mit dem passenden Ein-Klick-Ausweg.
   */
  const renderKlammerVisibilityMismatch = (): React.ReactElement | null => {
    if (!subEventsOnlyMode || subEvents.length === 0) return null;
    const split = (s: string): string[] => (s || '').split(',').map(x => x.trim()).filter(Boolean);
    const parentLocs = split(locationFilter);
    const parentAuds = split(audience);
    const parentOpen = parentLocs.length === 0 && parentAuds.length === 0;
    const childLocSets = subEvents.map(s => split(s.locationFilter || ''));

    // (a) Klammer offen, aber JEDES Sub-Event schränkt ein.
    if (parentOpen && childLocSets.every(l => l.length > 0)) {
      const union = Array.from(new Set(childLocSets.reduce((a, b) => a.concat(b), [])));
      return (
        <div style={{
          marginTop: 10, padding: '10px 12px', borderRadius: 8,
          background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
          fontSize: '0.78rem', lineHeight: 1.55,
        }}>
          <strong>{isDe ? 'Die Klammer lässt mehr zu als ihre Sub-Events' : 'The bracket is broader than its sub-events'}</strong>
          <div style={{ marginTop: 4 }}>
            {isDe
              ? <>Hier ist <strong>kein Standort</strong> gesetzt, das Event ist also für alle sichtbar — aber <strong>alle {subEvents.length} Sub-Events</strong> sind auf {union.length === 1 ? <>den Standort <strong>{union[0]}</strong></> : <>die Standorte <strong>{union.join(', ')}</strong></>} beschränkt. Wer nicht dazugehört, sieht das Event in der Übersicht, findet darin aber <strong>nichts, wofür er sich anmelden kann</strong>.</>
              : <>No location is set here, so the event is visible to everyone — but <strong>all {subEvents.length} sub-events</strong> are restricted to {union.join(', ')}. People outside see the event but find nothing they can register for.</>}
          </div>
          <button
            type="button"
            className="btn btn-primary"
            style={{ fontSize: '0.78rem', padding: '5px 12px', marginTop: 8 }}
            onClick={() => setLocationFilter(union.join(', '))}
          >
            {isDe
              ? `Klammer ebenfalls auf ${union.join(', ')} setzen`
              : `Restrict the bracket to ${union.join(', ')} as well`}
          </button>
        </div>
      );
    }

    // (b) Ein Sub-Event lässt mehr zu, als die Klammer durchlässt.
    if (!parentOpen && parentLocs.length > 0) {
      const lc = (s: string): string => s.toLowerCase();
      const parentLc = parentLocs.map(lc);
      const offenders = subEvents
        .map((s, i) => ({ s, extra: childLocSets[i].filter(l => parentLc.indexOf(lc(l)) < 0) }))
        .filter(x => x.extra.length > 0);
      if (offenders.length > 0) {
        const extras = Array.from(new Set(offenders.reduce<string[]>((a, b) => a.concat(b.extra), [])));
        return (
          <div style={{
            marginTop: 10, padding: '10px 12px', borderRadius: 8,
            background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
            fontSize: '0.78rem', lineHeight: 1.55,
          }}>
            <strong>{isDe ? 'Einstellungen in Sub-Events, die nicht greifen können' : 'Sub-event settings that cannot take effect'}</strong>
            <div style={{ marginTop: 4 }}>
              {isDe
                ? <>{offenders.length === 1 ? 'Ein Sub-Event lässt' : `${offenders.length} Sub-Events lassen`} {extras.length === 1 ? <>den Standort <strong>{extras[0]}</strong></> : <>die Standorte <strong>{extras.join(', ')}</strong></>} zu — die Klammer aber nicht. Der Zugang läuft immer über die Klammer, deshalb bleiben diese Personen <strong>trotzdem draußen</strong>. Entweder hier ergänzen oder im Sub-Event entfernen.</>
                : <>{offenders.length} sub-event(s) allow {extras.join(', ')}, but the bracket does not. Access always goes through the bracket, so those people stay out anyway.</>}
            </div>
            <button
              type="button"
              className="btn btn-primary"
              style={{ fontSize: '0.78rem', padding: '5px 12px', marginTop: 8 }}
              onClick={() => setLocationFilter(Array.from(new Set(parentLocs.concat(extras))).join(', '))}
            >
              {isDe
                ? `${extras.join(', ')} hier ergänzen`
                : `Add ${extras.join(', ')} here`}
            </button>
          </div>
        );
      }
    }
    return null;
  };

  // v15.6: Style-Helfer für den ausgegrauten Hauptevent-Tab-Inhalt. Bei
  // subEventsOnlyMode wird der gesamte Hauptevent-Tab-Inhalt mit Opacity 0.55
  // und pointer-events:none umhüllt, damit der Organizer optisch sofort sieht,
  // dass dieser Bereich aktuell wirkungslos ist. Der Hinweis-Banner steht
  // außerhalb der Hülle, bleibt also klar lesbar.
  const hauptGreyoutWrapperStyle = (): React.CSSProperties => (
    subEventsOnlyMode
      ? { opacity: 0.55, pointerEvents: 'none', userSelect: 'none' }
      : {}
  );

  // v17.3: Unsaved-Changes-Tracking + Navigation-Guard. Wir snapshotten
  // beim Mount die wichtigsten Form-Felder und prüfen bei Bedarf, ob sich
  // etwas geändert hat. Beim Zurück-Klick fragt ein Modal nach.
  const [unsavedConfirmOpen, setUnsavedConfirmOpen] = React.useState<null | { resolve: (_ok: boolean) => void }>(null);
  const initialFormSnapshotRef = React.useRef<string>('');
  const computeFormSnapshot = React.useCallback((): string => {
    return JSON.stringify({
      title, description, location,
      addrStreet, addrHouseNo, addrZip, addrCity,
      organizer, organizerEmails: organizerEmails.join(';'),
      startDate, endDate, registrationDeadline, lastDeregisterDate,
      maxParticipants, waitlistEnabled,
      audience, locationFilter, filterMode,
      contactName, contactEmail, contactInfo,
      eventImageUrl,
      teamRegistrationEnabled, teamSize, askTeamName, teamPartialAllowed,
      teamOpenSlotsVisible, teamJoinRequiresApproval,
      askSalutation, requireSubEventSelection,
      // Custom-Fields nur via Anzahl + Labels — JSON.stringify auf das
      // gesamte Array wäre instabil bei id-Änderungen.
      customFieldsHash: (customFields || []).map(f => `${f.id}:${f.label}:${f.type}:${f.required}`).join('|'),
      agendaLen: (agenda || []).length,
      docsLen: (documents || []).length,
      subEventsLen: (subEvents || []).length,
      outlookBody, outlookHeading, outlookSubheading, outlookSubject,
      disableEmails, disableRegistrationEmail, disableCancellationEmail, autoDeregisterOnDecline, inactiveHandling, disableOutlook,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      emailTemplateOverridesHash: JSON.stringify(emailTemplateOverrides || {}),
      // v29.66: F&A-Pilot — ohne diesen Eintrag warnt der Ungespeichert-
      // Waechter nicht, wenn nur Abrechnungsfelder geaendert wurden.
      billingHash: JSON.stringify({ billingRelevant, billingSendMode, billingFields }),
      subOpenRuleHash: JSON.stringify({ openRuleEnabled, openRuleMode, openRuleDays, openRuleFixedDate }), // v29.67/v29.76
      visAllSubs, // v29.75
      subDeadlineRuleHash: JSON.stringify({ regRuleEnabled, regRuleAmount, regRuleUnit, cancelRuleEnabled, cancelRuleAmount, cancelRuleUnit, cancelRuleAfter }), // v29.76/77
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    title, description, location, addrStreet, addrHouseNo, addrZip, addrCity,
    organizer, organizerEmails, startDate, endDate, registrationDeadline, lastDeregisterDate,
    maxParticipants, waitlistEnabled, audience, locationFilter, filterMode,
    contactName, contactEmail, contactInfo, eventImageUrl,
    teamRegistrationEnabled, teamSize, askTeamName, teamPartialAllowed,
    teamOpenSlotsVisible, teamJoinRequiresApproval, askSalutation, requireSubEventSelection,
    customFields, agenda, documents, subEvents,
    outlookBody, outlookHeading, outlookSubheading, outlookSubject, disableEmails, disableRegistrationEmail, disableCancellationEmail, autoDeregisterOnDecline, inactiveHandling, disableOutlook,
    emailTemplateOverrides,
    billingRelevant, billingSendMode, billingFields, // v29.66
    openRuleEnabled, openRuleMode, openRuleDays, openRuleFixedDate, // v29.67/v29.76
    visAllSubs, // v29.75
    regRuleEnabled, regRuleAmount, regRuleUnit, cancelRuleEnabled, cancelRuleAmount, cancelRuleUnit, cancelRuleAfter, // v29.76/77
  ]);
  React.useEffect(() => {
    // Initial-Snapshot ein paar Ticks nach dem ersten Render setzen, damit
    // alle initialen useEffect-Loads (z.B. editEvent → State-Hydration) durch sind.
    const t = setTimeout(() => { initialFormSnapshotRef.current = computeFormSnapshot(); }, 100);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editEvent?.id, isEditMode]);
  React.useEffect(() => {
    if (submitted) {
      setNavigationGuard(null);
      return;
    }
    const guard = async (): Promise<boolean> => {
      const isDirty = initialFormSnapshotRef.current !== '' && computeFormSnapshot() !== initialFormSnapshotRef.current;
      if (!isDirty) return true;
      return new Promise<boolean>(resolve => { setUnsavedConfirmOpen({ resolve }); });
    };
    setNavigationGuard(guard);
    return () => setNavigationGuard(null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [submitted, computeFormSnapshot]);

  // v22.22: Schwebender Weiter-Button — solange die Aktions-Zeile (Zurück /
  // Vorschau / Weiter) unten noch NICHT im Viewport ist, schwebt unten rechts
  // ein fixierter Weiter-Button. Sobald die echte Zeile sichtbar wird, blendet
  // er weich aus (Opacity + Transform) — wirkt wie ein sanftes „Andocken" an
  // seine eigentliche Stelle.
  const actionRowRef = React.useRef<HTMLDivElement | null>(null);
  const [actionRowVisible, setActionRowVisible] = React.useState(true);
  React.useEffect(() => {
    const el = actionRowRef.current;
    if (!el || typeof IntersectionObserver === 'undefined') {
      setActionRowVisible(true);
      return undefined;
    }
    const obs = new IntersectionObserver(entries => {
      setActionRowVisible(entries[0] ? entries[0].isIntersecting : true);
    }, { threshold: 0.3 });
    obs.observe(el);
    return () => obs.disconnect();
  }, [isSubmitting, submitted]);

  // v22.23: Das Organizer-Tutorial steuert den aktiven Wizard-Schritt von
  // außen (TutorialGuide dispatcht ein CustomEvent pro Tour-Schritt), damit
  // die Tour alle 9 Schritte nacheinander zeigen kann. Bewusst entkoppelt —
  // ohne laufende Tour feuert das Event nie.
  React.useEffect(() => {
    const onTourStep = (e: Event): void => {
      const detail = (e as CustomEvent).detail;
      // v29.21 (Audit B5): 0..8 statt 0..7 — die Tour-Karten der letzten
      // beiden Schritte wurden sonst verworfen und die Tour blieb auf
      // „Dokumente" stehen.
      // v30.59: Gegen die Schritte klemmen, die DIESE Person wirklich hat.
      // Seit v30.44 gibt es Schritt 10 (Abrechnung) nur für Admins — für alle
      // anderen endet der Assistent bei Index 8. Eine feste 9 hier setzte den
      // Schritt auf einen, den es für sie nicht gibt: Jede Anzeige-Bedingung
      // `currentStep === N` verfehlt dann, und die Seite bleibt LEER, ohne dass
      // irgendwo etwas meldet. Genau davor warnt der Kommentar an `maxStep` —
      // nur war diese Stelle nicht mitgezogen.
      if (typeof detail === 'number' && detail >= 0) {
        setCurrentStep(Math.min(detail, canBilling ? 9 : 8));
        setTriedNext(false);
      }
    };
    window.addEventListener('dex-tutorial-wizard-step', onTourStep);
    return () => window.removeEventListener('dex-tutorial-wizard-step', onTourStep);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // v22.36: Ausgefüllte Eingaben pastellgrün markieren — gleiche Optik wie
  // auf der Anmeldeseite (inputStyleGreen, v19.0). Statt ~100 Inputs einzeln
  // zu verdrahten: delegierte input/change-Listener + periodischer Sweep
  // (fängt auch programmatische Befüllung wie Demo-Daten/Edit-Hydration ab).
  // Die Klasse .dex-filled färbt Rand + Hintergrund (CSS im Step-Bar-Style).
  // v22.36: „Prüfen"-Modal — Übersicht aller Einstellungen inkl. Standards
  // und noch leerer optionaler Punkte.
  const [showConfigCheck, setShowConfigCheck] = React.useState(false);

  // v22.36: Sub-Events als Opt-in — Schritt 2 fragt zuerst, OB Sub-Events
  // genutzt werden sollen (Default: nein); erst dann erscheint die ganze
  // Konfiguration. Beim Bearbeiten/Demo-Befüllen mit vorhandenen Sub-Events
  // schaltet der Effekt einmalig automatisch auf „ja" (Ref verhindert, dass
  // er ein bewusstes Abschalten sofort wieder überschreibt).
  const [subEventsOptIn, setSubEventsOptIn] = React.useState<boolean>(false);
  // v29.55: Ein einzelner Termin gehört in den Kalender des Organizers — das
  // bleibt der Default. Sobald es eine Reihe wird, kippt er: Bei zwanzig Tagen
  // sind zwanzig Blocker für Tage ohne eigene Buchung genau das, worüber sich
  // Organizer beschwert haben. Hat der Organizer selbst entschieden
  // (orgInvitesTouchedRef) oder wird ein bestehendes Event bearbeitet, bleibt
  // sein Wert stehen.
  React.useEffect(() => {
    if (orgInvitesTouchedRef.current) return;
    setOrgGetsSubInvites(!(subEventsOptIn && subEvents.length > 0));
  }, [subEventsOptIn, subEvents.length]);

  // ============================================================
  // v30.0: Entwurfs-Zwischenspeicher fuer die Neu-Anlage.
  //
  // Wer DEX mitten in der Event-Erstellung schliesst, verlor bisher alles.
  // Jetzt sichert ein Debounce den Formularstand laufend nach localStorage;
  // beim naechsten Oeffnen der Event-Erstellung fragt ein Dialog, ob der
  // Entwurf fortgesetzt werden soll. Bewusste Grenzen:
  //  - NUR Neu-Anlage. Ein Edit-Entwurf koennte beim Wiederherstellen einen
  //    inzwischen live geaenderten Stand ueberschreiben.
  //  - Bilder (File-Objekte) lassen sich nicht serialisieren und fehlen im
  //    Entwurf — der Dialog sagt das dazu.
  //  - localStorage ist ein Komfort, kein Vertrag: Quota/Privacy-Fehler
  //    werden geschluckt, der Entwurf verfaellt nach 14 Tagen.
  // Der Block steht bewusst NACH subEventsOptIn (Zeile oben) — frueher
  // platzierte States waeren die v29.71-TDZ-Falle.
  // ============================================================
  const DRAFT_KEY = 'dex_event_creation_draft_v1';
  const draftPromptShownRef = React.useRef(false);
  // v30.1: Zeitpunkt der letzten Zwischenspeicherung — fuer die Anzeige
  // „Zwischengespeichert am …" ueber dem Formular. lastDraftJsonRef
  // verhindert die Endlosschleife Autosave → setState → Re-Render →
  // Autosave: unveraenderter Stand wird weder geschrieben noch gemeldet.
  const [draftSavedAt, setDraftSavedAt] = React.useState<number | null>(null);
  const lastDraftJsonRef = React.useRef('');
  // v30.4: Statt eines Modal-Dialogs beim Öffnen zeigt eine KACHEL unter
  // „Eigenes Event als Vorlage nutzen?" den gefundenen Entwurf — mit
  // Fortsetzen- und Löschen-Button. pendingDraft hält ihn, bis der User
  // entscheidet oder das eigene Tippen ihn überschreibt (draftSavedAt).
  const [pendingDraft, setPendingDraft] = React.useState<{ savedAt: number; data: Record<string, unknown> } | null>(null);
  const buildDraftPayload = (): Record<string, unknown> => ({
    title, description, location, addrStreet, addrHouseNo, addrZip, addrCity,
    organizer, organizerEmails, contactName, contactEmail, contactInfo,
    startDate, endDate, registrationDeadline, lastDeregisterDate, klammerDeadline, activeFrom,
    maxParticipants, waitlistEnabled, audience, locationFilter, filterMode, excludedUsers,
    customFields, agenda,
    // File-Objekte serialisieren zu {} — bewusst strippen statt Muell speichern.
    subEvents: subEvents.map(sd => ({ ...sd, imageFile: null })),
    subEventsOptIn, subEventsOnlyMode, subEventCalendar, subEventSingleChoice,
    requireSubEventSelection, askSalutation,
    teamRegistrationEnabled, teamSize, askTeamName,
    userCancelAllowed, noCancelAfterDeadline, teamsLink, onlineMeetingMode, // v30.26
    disableEmails, disableOutlook,
    emailTemplateOverrides,
    openRuleEnabled, openRuleMode, openRuleDays, openRuleFixedDate,
    regRuleEnabled, regRuleAmount, regRuleUnit,
    cancelRuleEnabled, cancelRuleAmount, cancelRuleUnit, cancelRuleAfter,
    visAllSubs,
    billingRelevant, billingSendMode, billingFields,
    currentStep,
  });
  const applyDraftPayload = (d: Record<string, unknown>): void => {
    const str = (v: unknown): string => (typeof v === 'string' ? v : '');
    const bool = (v: unknown, dflt: boolean): boolean => (typeof v === 'boolean' ? v : dflt);
    const num = (v: unknown, dflt: number): number => (typeof v === 'number' && isFinite(v) ? v : dflt);
    setTitle(str(d.title)); setDescription(str(d.description)); setLocation(str(d.location));
    setAddrStreet(str(d.addrStreet)); setAddrHouseNo(str(d.addrHouseNo)); setAddrZip(str(d.addrZip)); setAddrCity(str(d.addrCity));
    setOrganizer(str(d.organizer));
    if (Array.isArray(d.organizerEmails)) setOrganizerEmails(d.organizerEmails as string[]);
    setContactName(str(d.contactName)); setContactEmail(str(d.contactEmail)); setContactInfo(str(d.contactInfo));
    setStartDate(str(d.startDate)); setEndDate(str(d.endDate));
    setRegistrationDeadline(str(d.registrationDeadline)); setLastDeregisterDate(str(d.lastDeregisterDate));
    setKlammerDeadline(str(d.klammerDeadline)); setActiveFrom(str(d.activeFrom));
    setMaxParticipants(str(d.maxParticipants)); setWaitlistEnabled(bool(d.waitlistEnabled, false));
    setAudience(str(d.audience)); setLocationFilter(str(d.locationFilter));
    setFilterMode(d.filterMode === 'AND' ? 'AND' : 'OR');
    if (Array.isArray(d.excludedUsers)) setExcludedUsers(d.excludedUsers as string[]);
    if (Array.isArray(d.customFields)) setCustomFields(d.customFields as CustomFieldInput[]);
    if (Array.isArray(d.agenda)) setAgenda(d.agenda as AgendaItem[]);
    if (Array.isArray(d.subEvents)) setSubEvents((d.subEvents as SubEventDraft[]).map(x => ({ ...x, imageFile: null })));
    setSubEventsOptIn(bool(d.subEventsOptIn, false));
    setSubEventsOnlyMode(bool(d.subEventsOnlyMode, false));
    setSubEventCalendar(bool(d.subEventCalendar, false));
    setSubEventSingleChoice(bool(d.subEventSingleChoice, false));
    setRequireSubEventSelection(bool(d.requireSubEventSelection, false));
    setAskSalutation(bool(d.askSalutation, false));
    setTeamRegistrationEnabled(bool(d.teamRegistrationEnabled, false));
    setTeamSize(num(d.teamSize, 2)); setAskTeamName(bool(d.askTeamName, false));
    setUserCancelAllowed(bool(d.userCancelAllowed, true));
    setNoCancelAfterDeadline(bool(d.noCancelAfterDeadline, false));
    setTeamsLink(str(d.teamsLink));
    // v30.26: Modus aus dem Entwurf; Alt-Entwürfe kennen ihn nicht — dort
    // ergibt sich 'own' aus einem vorhandenen Link, sonst 'none'.
    const draftOmMode = str(d.onlineMeetingMode);
    setOnlineMeetingMode(
      draftOmMode === 'auto' || draftOmMode === 'own' || draftOmMode === 'none'
        ? draftOmMode
        : (str(d.teamsLink).trim() ? 'own' : 'none'),
    );
    setDisableEmails(bool(d.disableEmails, false)); setDisableOutlook(bool(d.disableOutlook, false));
    if (d.emailTemplateOverrides && typeof d.emailTemplateOverrides === 'object') {
      setEmailTemplateOverrides(d.emailTemplateOverrides as Record<string, EmailOverrideEntry>);
    }
    setOpenRuleEnabled(bool(d.openRuleEnabled, false));
    setOpenRuleMode(d.openRuleMode === 'week' ? 'week' : 'day');
    setOpenRuleDays(num(d.openRuleDays, 7)); setOpenRuleFixedDate(str(d.openRuleFixedDate));
    setRegRuleEnabled(bool(d.regRuleEnabled, false)); setRegRuleAmount(num(d.regRuleAmount, 1));
    setRegRuleUnit(d.regRuleUnit === 'hours' ? 'hours' : 'days');
    setCancelRuleEnabled(bool(d.cancelRuleEnabled, false)); setCancelRuleAmount(num(d.cancelRuleAmount, 1));
    setCancelRuleUnit(d.cancelRuleUnit === 'hours' ? 'hours' : 'days');
    setCancelRuleAfter(bool(d.cancelRuleAfter, false));
    setVisAllSubs(bool(d.visAllSubs, false));
    if (typeof d.billingRelevant === 'boolean') setBillingRelevant(d.billingRelevant);
    setBillingSendMode(d.billingSendMode === 'auto' ? 'auto' : 'manual');
    if (d.billingFields && typeof d.billingFields === 'object') setBillingFields(d.billingFields as Record<string, string>);
    // v30.59: Auch beim Entwurf nach OBEN klemmen. Ein Entwurf, der auf
    // Schritt 10 gespeichert wurde (als Admin, oder mit einem Build, in dem es
    // den Schritt noch für alle gab), führte beim Wieder-Öffnen als Organizer
    // auf einen Schritt, den es dort nicht gibt — und der Assistent zeigte
    // gar nichts mehr an. Dasselbe Muster wie beim Tour-Schritt oben.
    setCurrentStep(Math.min(Math.max(0, num(d.currentStep, 0)), canBilling ? 9 : 8));
  };
  // Beim Betreten der Neu-Anlage EINMAL den letzten Entwurf laden. v30.4:
  // kein Modal mehr — der Entwurf erscheint als Kachel in Schritt 1
  // (unter der Vorlagen-Kachel), dort entscheidet der User per Button.
  React.useEffect(() => {
    if (isEditMode || draftPromptShownRef.current) return;
    draftPromptShownRef.current = true;
    try {
      const raw = localStorage.getItem(DRAFT_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw) as { savedAt?: number; data?: Record<string, unknown> };
      const data = parsed?.data;
      const age = Date.now() - (parsed?.savedAt || 0);
      const hasSubstance = !!data && ((typeof data.title === 'string' && data.title.trim().length > 0) || (Array.isArray(data.subEvents) && data.subEvents.length > 0));
      if (!data || !hasSubstance || !(age >= 0) || age > 14 * 86400000) {
        localStorage.removeItem(DRAFT_KEY);
        return;
      }
      setPendingDraft({ savedAt: parsed.savedAt || 0, data });
    } catch { /* localStorage gesperrt o.ä. — kein Entwurf, kein Drama */ }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEditMode]);
  // Autosave: laeuft nach JEDEM Render, der Timer wird dabei neu aufgezogen —
  // gespeichert wird also ~1,5 s nach der letzten Aenderung (Tipp-Pause).
  // Kein Deps-Array: bei ~200 State-Variablen waere jede Liste sofort veraltet.
  React.useEffect(() => {
    // v30.4: createdEventIdRef — nach erfolgreichem Anlegen darf der Autosave
    // den soeben gelöschten Entwurf nicht aus dem noch gefüllten Formular
    // neu erzeugen (der Wizard bleibt für das Summary-Modal gemountet).
    if (isEditMode || submitted || createdEventIdRef.current) return undefined;
    const t = setTimeout(() => {
      try {
        const hasSubstance = title.trim().length > 0 || description.trim().length > 0 || subEvents.length > 0;
        if (!hasSubstance) return;
        const json = JSON.stringify(buildDraftPayload());
        if (json === lastDraftJsonRef.current) return; // nichts Neues — nicht schreiben
        lastDraftJsonRef.current = json;
        const now = Date.now();
        localStorage.setItem(DRAFT_KEY, JSON.stringify({ savedAt: now, data: JSON.parse(json) }));
        setDraftSavedAt(now);
      } catch { /* Quota voll / Privacy-Modus — Autosave ist best-effort */ }
    }, 1500);
    return () => clearTimeout(t);
  });
  // Nach erfolgreichem Anlegen ist der Entwurf erledigt.
  React.useEffect(() => {
    if (submitted && !isEditMode) {
      try { localStorage.removeItem(DRAFT_KEY); } catch { /* */ }
    }
  }, [submitted, isEditMode]);

  /**
   * v29.57 — Unveränderte Sub-Events beim Speichern überspringen.
   *
   * Ein Save schrieb bisher IMMER alle Sub-Events. Bei einer Terminreihe mit 21
   * Tagen sind das 21 MERGE-Requests plus Bild-Uploads, auch wenn der Organizer
   * nur an einem Tag die Frist geändert hat — und genau diese Masse hat die
   * SharePoint-Drosselung (HTTP 429) ausgelöst, die in v29.48 abgefangen
   * werden musste. Weniger schreiben ist die bessere Lösung als schneller
   * wiederholen.
   *
   * Uebersprungen wird nur, wenn ZWEI Bedingungen zusammenkommen:
   *
   *  1. Der Entwurf ist Zeichen für Zeichen der, der beim Oeffnen geladen
   *     wurde (`subPersistKey`, ohne die reinen Snapshot-/UI-Felder).
   *  2. Am Hauptevent hat sich NICHTS geändert (`subTopGateKey`).
   *
   * Punkt 2 ist der wichtige: Ein Sub-Event erbt eine Menge vom Hauptevent —
   * Organizer, Zeiten (bei leeren Sub-Zeiten), Mail-Logo, Kopfbild-Layout,
   * Teams-Link, Sprache, Zweisprachigkeit der Felder. Aendert sich davon etwas,
   * ändert sich der Schreibvorgang JEDES Sub-Events, obwohl kein einziger
   * Entwurf angefasst wurde. Deshalb ist die Schranke bewusst UEBERINKLUSIV:
   * Sie nimmt den kompletten Formular-Snapshot plus alles, was der nicht
   * abdeckt. Zu oft schreiben kostet Zeit, zu selten schreiben verliert Daten.
   *
   * WER HIER ETWAS ERGAENZT, das vom Hauptevent in ein Sub-Event fließt, muss
   * es in `subTopGateKey` aufnehmen — sonst wird die Aenderung still verworfen.
   */
  const subPersistKey = (d: SubEventDraft): string => {
    const out: Record<string, unknown> = {};
    for (const k of Object.keys(d).sort()) {
      // `id` ist eine Client-ID, `initial*` sind Vergleichs-Snapshots — beide
      // sagen nichts darüber aus, was in SharePoint landet.
      if (k === 'id' || k.indexOf('initial') === 0) continue;
      const v = (d as unknown as Record<string, unknown>)[k];
      // WICHTIG: `JSON.stringify(File)` liefert `{}`. Ein frisch ausgewähltes
      // Bild wäre in diesem Schlüssel also unsichtbar, das Sub-Event würde
      // als unverändert gelten und der Upload fiele still aus. Deshalb Datei-
      // Werte über ihre Kennzeichen vergleichen.
      out[k] = (typeof File !== 'undefined' && v instanceof File)
        ? `file:${v.name}:${v.size}:${v.lastModified}`
        : v;
    }
    return JSON.stringify(out);
  };
  const subTopGateKey = (): string => {
    const pair = sanitizeOrganizerPairs();
    return JSON.stringify({
      form: computeFormSnapshot(),
      org: `${pair.orgString}|${pair.orgEmailString}`,
      orgInvites: orgGetsSubInvites,
      start: startDate, end: endDate,
      bilingual: bilingualFields,
      emailLang: emailLanguage,
      header: headerImageLayout,
      teams: teamsLink,
      allDay, showAsFree,
      logos: `${(emailLogoPreview || '').length}:${(outlookLogoPreview || '').length}`,
    });
  };
  // v29.57: Die Schimmer-Keyframes werden sonst nur im Boot-Bildschirm
  // eingehängt (DexEventPlatform). Wer den Wizard nach einem gecachten Start
  // öffnet, hätte den Balken ohne Animation. Idempotent über die id.
  React.useEffect(() => {
    if (typeof document === 'undefined') return;
    if (document.getElementById('dex-progress-pulse-keyframes')) return;
    const style = document.createElement('style');
    style.id = 'dex-progress-pulse-keyframes';
    style.textContent = '@keyframes dexProgressShimmer { 0% { transform: translateX(-120%); } 100% { transform: translateX(320%); } }';
    document.head.appendChild(style);
  }, []);

  /** Entwürfe, wie sie beim Oeffnen geladen wurden (Schlüssel = dbId). */
  const initialSubPersistRef = React.useRef<Record<string, string>>({});
  const subTopGateInitialRef = React.useRef<string>('');
  React.useEffect(() => {
    if (!editEvent) return;
    const map: Record<string, string> = {};
    for (const se of subEventsRef.current) if (se.dbId) map[se.dbId] = subPersistKey(se);
    initialSubPersistRef.current = map;
    subTopGateInitialRef.current = subTopGateKey();
    // Absichtlich nur beim Mount — spätere Aenderungen sind genau das, was
    // erkannt werden soll.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  // v28.2 SOFT-DISABLE: Der Toggle verwirft keine Drafts mehr (v27.11-Stash
  // entfällt) — `subEventsOptIn === false` bei vorhandenen Drafts heißt nur
  // noch „deaktiviert": beim Speichern wird das Piggyback-Flag
  // _subEventsDisabled gesetzt, die Kinder bleiben vollständig erhalten.
  const subOptInHydratedRef = React.useRef(false);
  React.useEffect(() => {
    if (!subOptInHydratedRef.current && subEvents.length > 0) {
      subOptInHydratedRef.current = true;
      // v28.2: Soft-deaktivierte Events starten mit Toggle AUS — die Drafts
      // sind trotzdem geladen und kommen beim Einschalten sofort wieder.
      setSubEventsOptIn(!(editEvent && editEvent.subEventsDisabled));
    }
  }, [subEvents.length]);

  const wizardRootRef = React.useRef<HTMLDivElement | null>(null);
  React.useEffect(() => {
    const root = wizardRootRef.current;
    if (!root) return undefined;
    const SELECTOR = 'input.form-input, textarea.form-input, select.form-select, select.form-input';
    const apply = (el: Element): void => {
      const input = el as HTMLInputElement;
      const type = (input.type || '').toLowerCase();
      if (type === 'checkbox' || type === 'radio' || type === 'file' || type === 'button') return;
      input.classList.toggle('dex-filled', !!(input.value || '').trim());
    };
    const sweep = (): void => { root.querySelectorAll(SELECTOR).forEach(apply); };
    const onInput = (e: Event): void => {
      const t = e.target as HTMLElement;
      if (t && typeof t.matches === 'function' && t.matches(SELECTOR)) apply(t);
    };
    sweep();
    const iv = window.setInterval(sweep, 1200);
    root.addEventListener('input', onInput, true);
    root.addEventListener('change', onInput, true);
    return () => {
      window.clearInterval(iv);
      root.removeEventListener('input', onInput, true);
      root.removeEventListener('change', onInput, true);
    };
  }, []);

  // v30.40: Diese fünf Hooks standen bis hierher UNTER `if (submitted)`.
  // `submitted` kippt beim Speichern — damit sank die Hook-Anzahl mitten im
  // Leben der Komponente, und das ist derselbe Fehler, der in v30.3 nach jeder
  // Anmeldung den Bildschirm geleert hat (React #300). Gefunden mit
  // `react-hooks/rules-of-hooks`, die in diesem Projekt bis v30.40 nicht lief.
  //
  // Verschoben wurde nur die POSITION, kein Zeichen am Inhalt. Möglich war das,
  // weil keiner der fünf beim Anlegen etwas liest, das erst weiter unten
  // entsteht: vier `useState` mit Literalen, ein `useMemo` über reine
  // Literale, und der `useEffect` hat `[]` und läuft ohnehin erst nach dem
  // Render. NEUE Hooks gehören ebenfalls hierher — nie unter den Return.
  // Tooltip-State: welcher Step zeigt gerade seinen Hint-Tooltip an?
  const [hintStepIdx, setHintStepIdx] = React.useState<number | null>(null);
  // Baseline der Klammer-Sichtbarkeit beim Mount festhalten (Original-Stand des
  // Events bzw. leer bei Neuanlage) — damit später erkannt wird, ob der
  // Organizer die Sichtbarkeit wirklich geändert/neu gesetzt hat.
  React.useEffect(() => {
    if (visSnapshotRef.current === null) visSnapshotRef.current = visKey(locationFilter, audience, filterMode);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const SUB_TRANSFER_GROUPS: Array<{ key: string; de: string; en: string; fields: string[] }> = React.useMemo(() => ([
    { key: 'visibility', de: 'Sichtbarkeit (Standortfilter, Mailverteiler, Verknüpfung, Ausschlüsse)', en: 'Visibility (location filter, mailing lists, link mode, exclusions)', fields: ['locationFilter', 'audience', 'filterMode', 'excludedUsers'] },
    { key: 'capacity', de: 'Teilnehmerzahl & Warteliste', en: 'Capacity & waitlist', fields: ['maxParticipants', 'waitlistEnabled'] },
    { key: 'regDeadline', de: 'Anmeldefrist', en: 'Registration deadline', fields: ['registrationDeadline'] },
    { key: 'deregDeadline', de: 'Abmeldefrist', en: 'Cancellation deadline', fields: ['lastDeregisterDate'] },
    { key: 'place', de: 'Ort & Adresse', en: 'Location & address', fields: ['location', 'locationAddress'] },
    { key: 'mandatory', de: 'Pflichtanmeldung', en: 'Mandatory registration', fields: ['mandatory'] },
    { key: 'communication', de: 'Kommunikation (Logo, Outlook-Text, Überschriften, Betreff, Mail-Sprache, Mail-Schalter)', en: 'Communication (logo, Outlook text, headings, subject, mail language, mail toggles)', fields: ['emailLanguage', 'emailLogoBase64', 'outlookLogoBase64', 'outlookBody', 'outlookHeading', 'outlookSubheading', 'outlookSubject', 'disableEmails', 'disableRegistrationEmail', 'disableCancellationEmail', 'autoDeregisterOnDecline', 'inactiveHandling', 'disableOutlook', 'emailTemplateOverrides'] },
    { key: 'times', de: 'Zeiten (Start & Ende) — überschreibt die Termine!', en: 'Times (start & end) — overwrites the dates!', fields: ['startDate', 'endDate'] },
  ]), []);
  const [subTransfer, setSubTransfer] = React.useState<null | { fromIdx: number; groups: string[]; targets: number[] }>(null);
  const [activeScopeIdx, setActiveScopeIdx] = React.useState<number>(0);

  if (submitted) {
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '64px 32px' }}>
          <h2>{isEditMode ? 'Event erfolgreich aktualisiert!' : 'Event erfolgreich erstellt!'}</h2>
          <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>
            &bdquo;{title}&ldquo; wurde {isEditMode ? 'aktualisiert' : 'angelegt'}.
          </p>
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button
              className="btn btn-primary"
              onClick={() => {
                // v9.43: Hard-Reload mit Deep-Link statt Soft-Navigate. Grund:
                // nach Event-Erstellung braucht SharePoint ein paar Sekunden, bis
                // die neue Subsite + Listen API-konsistent sind, und der
                // Permission-Cache des Browsers hat für die frische Subsite noch
                // keine gültigen Tokens. Ein Soft-Navigate führt deshalb in einen
                // Render-Crash mit weißem Screen (React #300).
                //
                // Der Hard-Reload räumt den Permission-Cache auf. Damit der User
                // nicht auf der Landing-Seite landet und manuell zur Eventliste
                // klicken muss, hängen wir einen Deep-Link ?action=event-created
                // dran. Der Bootstrap der App liest diesen Parameter, navigiert
                // automatisch zur Eventliste und zeigt eine grüne Erfolgs-Banner-
                // Meldung mit dem Event-Titel.
                const action = isEditMode ? 'event-updated' : 'event-created';
                const targetEventId = selectedEventId || '';
                const url = window.location.pathname + '?action=' + action + (targetEventId ? '&event=' + encodeURIComponent(targetEventId) : '');
                window.location.href = url;
              }}
            >
              Zur Übersicht
            </button>
            <button className="btn btn-secondary" onClick={() => { setSubmitted(false); setTitle(''); }}>Weiteres Event erstellen</button>
          </div>
          <p style={{ marginTop: 20, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
            <em>Hinweis: Beim Klick auf {'„Zur Übersicht“'} wird die Seite einmal neu geladen, damit SharePoint die frisch erstellte Subsite überall sauber einbindet — du landest direkt in der Eventliste mit einer Erfolgs-Meldung.</em>
          </p>
        </div>
      </div>
    );
  }

  // Hilfsfunktion für die Vorschau
  const formatPreviewDate = (val: string): string => {
    if (!val) return '--';
    const d = new Date(val);
    return d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
      ' ' + d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  };

  // Vorschau-Sektion rendern
  const renderPreviewSection = (sectionId: string): React.ReactElement | null => {
    switch (sectionId) {
      case 'event':
        return (
          <div className="registration-event" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
            <div className="section-header section-header--red">Selected Event</div>
            <div className="registration-event__card">
              <div className="registration-event__image" style={{
                background: eventImageUrl
                  ? `url(${eventImageUrl}) center/cover`
                  : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
              }}>
                <div className="registration-event__overlay">
                  <h4>{title || 'Event Titel'}</h4>
                  <p>{formatPreviewDate(startDate)} until<br />{formatPreviewDate(endDate)}</p>
                </div>
              </div>
            </div>
          </div>
        );
      case 'personal':
        return (
          <div className="registration-form" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
            <div className="section-header">Personal Information</div>
            <div style={{ padding: '16px 20px' }}>
              <div className="form-group"><label className="form-label"><span className="required">*</span> Salutation</label><select className="form-select" disabled><option>Please select</option></select></div>
              <div className="form-group"><label className="form-label"><span className="required">*</span> First Name</label><input className="form-input" disabled placeholder="First Name" /></div>
              <div className="form-group"><label className="form-label"><span className="required">*</span> Surname</label><input className="form-input" disabled placeholder="Surname" /></div>
              <div className="form-group"><label className="form-label"><span className="required">*</span> E-Mail</label><input className="form-input" disabled placeholder="email@deloitte.de" /></div>
            </div>
          </div>
        );
      case 'specific':
        return (
          <div className="registration-specific" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
            <div className="section-header">Event specific Information</div>
            <div style={{ padding: '16px 20px' }}>
              {customFields.filter(f => f.label).length === 0 ? (
                <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', fontSize: '0.9rem' }}>No additional information required.</p>
              ) : (
                customFields.filter(f => f.label).map(field => (
                  <div className="form-group" key={field.id}>
                    <label className="form-label">{field.required && <span className="required">*</span>}{field.label}</label>
                    {field.type === 'select' && field.multi ? (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 4, padding: 8, border: '1px solid var(--dex-gray-200)', borderRadius: 6, background: '#fff' }}>
                        {field.options.map(o => o.trim()).filter(Boolean).map(opt => (
                          <label key={opt} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                            <input type="checkbox" disabled />
                            <span>{opt}</span>
                          </label>
                        ))}
                        <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-400)', marginTop: 2 }}>Mehrere Auswahl möglich</span>
                      </div>
                    ) : field.type === 'select' ? (
                      <select className="form-select" disabled><option>Please select</option>{field.options.map(o => o.trim()).filter(Boolean).map(opt => <option key={opt}>{opt}</option>)}</select>
                    ) : field.type === 'checkbox' ? (
                      <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.9rem' }}><input type="checkbox" disabled /> {field.label}</label>
                    ) : (
                      <input className="form-input" disabled placeholder={field.label} type={field.type === 'number' ? 'number' : 'text'} />
                    )}
                  </div>
                ))
              )}
            </div>
          </div>
        );
      case 'actions':
        return (
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 16 }}>
            <button className="btn btn-danger" disabled style={{ opacity: 0.5 }}><Trash2 size={16} /> Delete</button>
            <button className="btn btn-primary" disabled style={{ opacity: 0.5 }}><Send size={16} /> Register</button>
          </div>
        );
      default:
        return null;
    }
  };


  // Hint-Bullets pro Step. Werden über das i-Icon in der Progress-Bar
  // (Mouseover) als Tooltip eingeblendet — vorher wurden sie als
  // dauerhafte grüne Hinweis-Box am Anfang jedes Steps angezeigt
  // (renderStepIntro), das war für geübte Organizer zu viel Rauschen.
  const STEP_HINTS_DE: string[][] = [
    [
      'Event-Titel und Beschreibung — werden auf der Eventliste und der Registrierungsseite angezeigt',
      'Event-Bild hochladen (wird oben auf der Detailseite und in den Mails verwendet)',
      'Als Entwurf speichern — taucht dann nur für Admins, Organizer und Test-Team auf',
    ],
    [
      // v24.12 Schritt 2: Organizer & Team
      'Organizer auswählen — bekommen alle Organizer-Mails und sehen das Event im Admin Center; einzelne lassen sich von der Anmeldeseite ausblenden',
      'Anzeige der Organizer auf dem Anmeldeformular wählen (klein/groß) — mit Live-Vorschau',
      'Optional: externen Ansprechpartner (z.B. Service-Mail) angeben',
      'Test-Team: sieht das Event schon im Entwurf',
      'Check-in-Team: bedient am Event-Tag nur das Check-in-Tool',
    ],
    [
      // v15 Step 3: Ort & Programm (mit Tabs pro Sub-Event)
      'Veranstaltungsort und Adresse erfassen — pro Sub-Event optional eigener Ort (per Tab)',
      'Start- und End-Datum (mit Uhrzeit) festlegen',
      'Agenda mehrtägig pflegen (Drag-Reihenfolge pro Tag)',
      'Transferzeiten — Bus / Shuttle / Bahn von/zum Veranstaltungsort',
    ],
    [
      // v15 Step 4: Kapazität & Sichtbarkeit (mit Tabs pro Sub-Event)
      'Maximale Teilnehmerzahl festlegen (oder Unbegrenzt) — pro Sub-Event eigene Kapazität per Tab (Default: vom Hauptevent übernehmen)',
      'Anmeldefrist setzen — pro Sub-Event eigene Deadline möglich (leer = Hauptevent-Deadline gilt)',
      'Optional: Letzte Abmeldemöglichkeit — die kommunizierte Abmeldefrist; danach bleibt die Abmeldung bis zum Event-Ende möglich, die Organizer werden aber automatisch informiert',
      'Warteliste aktivieren — voll besetzte Events nehmen weitere Anmeldungen auf, bis ein Platz frei wird',
      'Optional: Geteilte Kapazität — zwei frei benannte Gruppen mit eigener Platzzahl + eigener oder gemeinsamer Warteliste',
    ],
    [
      // v15 Step 5: Felder (mit Tabs pro Sub-Event)
      'Feldtyp wählen: Text, Zahl, Dropdown, Checkbox, Personen-Suche oder Roommate (Doppelzimmer)',
      'Mehrfachauswahl bei Dropdowns (z.B. mehrere Allergien anhaken)',
      'Pflichtfeld setzen (rotes Sternchen, Anmeldung blockiert wenn leer)',
      'Beschreibung pro Feld — landet als „i"-Tooltip neben dem Feld-Label',
      'Sichtbarkeitsbedingung: Feld nur dann anzeigen wenn eine andere Frage einen bestimmten Wert hat (z.B. „Zimmerart nur fragen wenn Hotel = ja")',
      'Pro Sub-Event eigene Felder per Tab (Default: vom Hauptevent übernehmen)',
    ],
    [
      // v15 Step 6: Kommunikation
      'E-Mail-Sprache (DE/EN) für die automatischen Mails an die Teilnehmer wählen',
      'Pro Mail-Template (Anmeldung, Storno, Warteliste, Erinnerung, QR-Code…) den Subject/Heading/Body anpassen — mit Live-Vorschau',
      'Eigenes Logo / Header-Bild für Mail und Outlook-Termin hochladen',
      'Outlook-Termin-Body individuell gestalten (Live-Vorschau zeigt wie das Outlook-Element später aussieht)',
      'Benachrichtigungen optional komplett deaktivieren — z.B. für interne Entwurfs-Events',
      'Pro Sub-Event eigene Mail-Texte + Outlook-Body + Disable-Toggles per Tab',
    ],
    [
      // v15 Step 7: Team-Anmeldung
      'Team-Anmeldung erlauben — ein Teilnehmer kann sich für sich + sein Team gleichzeitig anmelden',
      'Team-Größe festlegen (2-N Personen)',
      'Optional Team-Namen abfragen — z.B. für Quiz- oder Lauf-Teams',
      'Beitritts-Modus: nur komplette Teams ODER auch Teil-Teams erlaubt',
      'Optional: offene Slots öffentlich sichtbar — andere Teilnehmer können beitreten (ggf. mit Lead-Approval)',
    ],
    [
      // v15 Step 8: Dokumente
      'Dokumente hochladen (PDF) — Teilnehmer sehen sie auf MyEvents als Inline-Vorschau oder Download',
    ],
    [
      // v15 Step 9: Fun-Zone
      'Quiz-Fragen für das Event anlegen — Multiple-Choice mit beliebig vielen Antwortoptionen',
      'Pro Frage optional ein Bild hochladen (Logo, Foto-Quiz, etc.)',
      'Mehrere richtige Antworten möglich (Mehrfachauswahl) — werden alle für volle Punktzahl gebraucht',
      'Cluster-Größe steuern: wie viele Fragen pro „Spielblock" angezeigt werden — Teilnehmer kann zwischenspeichern und später weitermachen',
      'Live-Highscore + Statistik im Admin Center sehen (welche Fragen am häufigsten falsch beantwortet werden)',
    ],
    // v29.66: Schritt 10 „Abrechnung" (F&A-Pilot, nur Admins sehen den Schritt).
    [
      'Event als abrechnungsrelevant kennzeichnen — die Entscheidung ist jederzeit änderbar',
      'Versandart wählen: automatisch (7 Tage vor/nach dem Event) oder manuell über das Organizer Center',
      'Alle elf Pflichtangaben für Finance & Accounting pflegen — unvollständig blockiert das Speichern nicht',
    ],
  ];
  // v29.21 (Audit B5): Die EN-Liste stand noch auf der 10-Schritt-Zählung von
  // vor v28.87 (alter Grundlagen-Mix, eigener Sub-Events-Schritt) — ab dem
  // dritten Eintrag zeigte jedes i-Icon die Hints des FALSCHEN Schritts, und
  // für Schritt 9 gab es gar keinen Eintrag. Jetzt 1:1 parallel zu
  // STEP_HINTS_DE (9 Einträge).
  const STEP_HINTS_EN: string[][] = [
    [
      'Event title and description — shown on the event list and registration page',
      'Upload an event image (used at the top of the detail page and in emails)',
      'Save as draft — only visible to admins, organizers and the test team',
    ],
    [
      // Step 2: Organizers & Team
      'Pick the organizers — they receive all organizer emails and see the event in the admin center; individual ones can be hidden from the registration page',
      'Choose how organizers appear on the registration form (small/large) — with live preview',
      'Optional: add an external contact (e.g. a service mailbox)',
      'Test team: sees the event while it is still a draft',
      'Check-in team: only operates the check-in tool on event day',
    ],
    [
      // Step 3: Location & Programme (with tabs per sub-event)
      'Enter venue and address — per sub-event an own location is possible (via tab)',
      'Set start and end date (with time)',
      'Maintain a multi-day agenda (drag ordering per day)',
      'Transfer times — bus / shuttle / train to and from the venue',
    ],
    [
      // v15 Step 4: Capacity & Visibility (with tabs per sub-event)
      'Set the maximum number of attendees (or Unlimited) — per sub-event own capacity via tab (default: inherit from main event)',
      'Set the registration deadline — per sub-event own deadline possible (empty = main-event deadline applies)',
      'Optional: last cancellation date — the communicated deadline; cancelling stays possible until the event ends, but organizers are notified automatically',
      'Enable waitlist — full events accept new registrations and promote them once a spot frees up',
      'Optional: split capacity — two freely-named groups with own seat count + own or shared waitlist',
    ],
    [
      // v15 Step 5: Fields (with tabs per sub-event)
      'Pick a field type: text, number, dropdown, checkbox, people search or roommate (double room)',
      'Multi-select for dropdowns (e.g. tick multiple allergies)',
      'Mark required (red asterisk, blocks submit when empty)',
      'Description per field — appears as „i" tooltip next to the field label',
      'Visibility condition: only show this field when another question has a specific value (e.g. „Only ask room type if Hotel = yes")',
      'Per sub-event own fields via tab (default: inherit from main event)',
    ],
    [
      // v15 Step 6: Communication
      'Pick the email language (DE/EN) for automatic emails to attendees',
      'Edit subject / heading / body per email template (registration, cancellation, waitlist, reminder, QR code…) — with live preview',
      'Upload a custom logo / header image for the email and Outlook event',
      'Customise the Outlook event body (live preview shows how the Outlook item will appear)',
      'Optionally disable notifications entirely — e.g. for internal draft events',
      'Per sub-event own mail texts + Outlook body + disable toggles via tab',
    ],
    [
      // v15 Step 7: Team Registration
      'Allow team registration — an attendee can register themselves + their team at once',
      'Set team size (2-N people)',
      'Optionally ask for a team name — e.g. quiz or running teams',
      'Join mode: complete teams only OR partial teams allowed',
      'Optional: open slots publicly visible — other attendees can join (with optional lead approval)',
    ],
    [
      // v15 Step 8: Documents
      'Upload documents (PDF) — attendees see them on MyEvents as inline preview or download',
    ],
    [
      // v15 Step 9: Fun-Zone
      'Create quiz questions for the event — multiple choice with any number of answer options',
      'Optionally upload an image per question (logo, photo quiz, etc.)',
      'Multiple correct answers are supported — all of them must be picked for full points',
      'Control cluster size: how many questions per „play block" — attendees can save progress and continue later',
      'See live highscore + statistics in the admin center (which questions are most often answered incorrectly)',
    ],
    // v29.66: step 10 "Billing" (F&A pilot, admins only).
    [
      'Mark the event as billing-relevant — the decision can be changed at any time',
      'Pick the delivery mode: automatic (7 days before/after the event) or manual via the organizer center',
      'Maintain all eleven mandatory Finance & Accounting details — incomplete data never blocks saving',
    ],
  ];

  const steps = [
    { label: t('create.step.basics'), icon: '1' },
    // v24.12: Organizer-Einstellungen als eigener Schritt direkt nach Grundlagen.
    { label: isDe ? 'Organizer & Team' : 'Organizers & Team', icon: '2' },
    // v15.0: Sub-Events kommen vor „Ort & Programm". Hintergrund:
    // Steps 4-6 (Ort, Kapazität, Felder) zeigen pro-Sub-Event-Tabs,
    // damit der Organizer pro Sub-Event eigenes Ort / eigene Kapazität /
    // eigene Felder pflegen kann — dafür müssen die Sub-Events schon
    // angelegt sein, deshalb ist Schritt 3 der Sub-Events-Step.
    { label: t('create.step.datetime'), icon: '3' },
    { label: t('create.step.capacity'), icon: '4' },
    { label: t('create.step.fields'), icon: '5' },
    { label: t('create.step.communication'), icon: '6' },
    // v15.0: Team-Anmeldung kommt jetzt nach Kommunikation (vorher nach
    // Kapazität). Reihenfolge spiegelt den realen Setup-Workflow besser
    // wider: erst die Komm-Texte stehen, dann entscheidet der Organizer
    // ob Team-Anmeldung relevant ist.
    { label: t('create.step.team'), icon: '7' },
    { label: t('create.step.documents'), icon: '8' },
    { label: t('create.step.funzone'), icon: '9' },
    // v29.66: F&A-Pilot — Schritt 10 haengt am ENDE, damit kein bestehender
    // Index wandert (die Falle aus CLAUDE.md: currentStep === N,
    // STEP_HINTS, SCOPE_AWARE_STEPS, getStepErrors haengen alle an festen
    // Indizes). Navigation und Speichern-Knopf laufen ueber steps.length
    // und ziehen automatisch mit. `dim` graut den Reiter aus, solange das
    // Event nicht abrechnungsrelevant ist — oeffnen bleibt erlaubt.
    ...(canBilling ? [{
      label: isDe ? 'Abrechnung' : 'Billing',
      icon: '10',
      dim: billingRelevant !== true,
    }] : []),
  ];


  // v29.21 (Audit B3): parametrisiert — der Kreis-Klick muss auch die
  // ÜBERSPRUNGENEN Schritte prüfen können, nicht nur den aktuellen.
  const getStepErrorsFor = (step: number): string[] => {
    const errors: string[] = [];
    switch (step) {
      case 0:
        // Schritt 1 (Grundlagen): Titel + Datum sind Pflicht. Die Datum-Checks
        // laufen hier, weil die DatePicker schon in Grundlagen stehen.
        if (!title) errors.push('title');
        if (!startDate) errors.push('startDate');
        if (!endDate) errors.push('endDate');
        // v29.55 BUG-FIX: Bei einem ganztägigen Termin ist die Uhrzeit
        // bedeutungslos — die DatePicker liefern ohne Zeitauswahl beide Male
        // 00:00, und `<=` meldete dann bei Start = Ende denselben Tag als
        // Fehler. Ganztägig wird deshalb tagesgenau verglichen: Fehler nur,
        // wenn der End-TAG vor dem Start-TAG liegt.
        if (startDate && endDate) {
          const bad = allDay
            ? endDate.slice(0, 10) < startDate.slice(0, 10)
            : new Date(endDate) <= new Date(startDate);
          if (bad) errors.push('endBeforeStart');
        }
        // v9.14: description ist optional — kein Pflichtfeld mehr
        // v28.87: Die Sub-Events stehen seit dem Wegfall von Schritt 3 in
        // Grundlagen — also wird ihre Datumsprüfung hier mitgeführt (v18.36:
        // Ende vor Start laesst den Outlook-Create-Flow mit HTTP 400 scheitern).
        if (subEvents.some(se => se.title && se.title.trim() && se.startDate && se.endDate && new Date(se.endDate) <= new Date(se.startDate))) {
          errors.push('subEventEndBeforeStart');
        }
        break;
      case 1:
        // v24.12: Schritt 2 (Organizer & Team) — mindestens ein Organizer ist Pflicht.
        if (!organizer) errors.push('organizer');
        break;
      case 2:
        // Schritt 3 (Ort & Programm) ist ohne Pflicht-Validierung —
        // Adresse / Agenda / Transferzeiten sind alle optional.
        break;
      case 3:
        // Schritt 4 (Kapazität & Sichtbarkeit).
        // v29.21 (Audit B4): Im „Nur Sub-Events"-Modus sind die geprüften
        // Felder gar nicht bedienbar — der sichtbare DatePicker editiert die
        // Klammer-Frist, die Abmeldefrist ist ausgegraut, der Kapazitätsblock
        // durch die Erklär-Box ersetzt. Die Prüfungen sperrten „Weiter" dann
        // ohne sichtbaren Grund und ohne Ausweg. Gleiches bei geteilter
        // Kapazität: maxParticipants ist dort per Konvention 0/leer, die
        // Fehlermeldung rendert nur im Nicht-Split-Zweig.
        if (!subEventsOnlyMode) {
          if (registrationDeadline && startDate && new Date(registrationDeadline) > new Date(startDate)) errors.push('deadlineAfterStart');
          if (userCancelAllowed && lastDeregisterDate && startDate && new Date(lastDeregisterDate) > new Date(startDate)) errors.push('deregAfterStart');
          if (!useSplitCapacities && !unlimitedParticipants && (maxParticipants === '' || isNaN(Number(maxParticipants)) || Number(maxParticipants) < 0)) errors.push('maxParticipants');
        }
        break;
    }
    return errors;
  };
  const getStepErrors = (): string[] => getStepErrorsFor(currentStep);


  // v22.62: Klammer-Sichtbarkeit auf alle Sub-Events kopieren.
  const applyParentVisibilityToSubs = (): void => {
    setSubEvents(prev => prev.map(se => ({ ...se, locationFilter, audience, filterMode })));
  };
  const visKey = (loc: string, aud: string, mode: string): string => `${(loc || '').trim()}${(aud || '').trim()}${mode || 'AND'}`;
  // v22.63: Fragt, ob die geänderte Klammer-Sichtbarkeit auf die Sub-Events
  // übernommen werden soll — immer wenn (a) die Klammer eine Sichtbarkeit hat,
  // (b) sie sich seit der Baseline geändert hat UND (c) sie von mindestens
  // einem Sub-Event abweicht. Sonst läuft `proceed` direkt.
  const interceptVisibilityCopy = (proceed: () => void): void => {
    const curKey = visKey(locationFilter, audience, filterMode);
    const hasVisibility = locationFilter.trim() !== '' || audience.trim() !== '';
    const changed = visSnapshotRef.current !== null && curKey !== visSnapshotRef.current;
    const differsFromSubs = subEvents.some(se => visKey(se.locationFilter || '', se.audience || '', se.filterMode || 'AND') !== curKey);
    const shouldAsk = subEvents.length > 0 && hasVisibility && changed && differsFromSubs;
    if (!shouldAsk) {
      visSnapshotRef.current = curKey; // Stand als „abgehandelt" merken
      proceed();
      return;
    }
    visCopyPendingRef.current = proceed;
    setVisCopyModalOpen(true);
  };
  // „Weiter" mit Sichtbarkeits-Abfrage.
  const proceedNext = (): void => {
    setTriedNext(true);
    const errs = getStepErrors();
    if (errs.length > 0) {
      // v28.89: Schritt 1 ist scope-fähig — die Pflichtfelder (Titel, Start,
      // Ende) gehören aber zum Hauptevent bzw. der Klammer. Steht der Reiter
      // auf einem Sub-Event, sind sie nicht einmal sichtbar: „Weiter" täte
      // scheinbar nichts. Deshalb auf die Ebene wechseln, auf der der Fehler
      // steht — beim Datums-Dreher andersherum auf das betroffene Sub-Event.
      if (currentStep === 0) {
        const mainFields = ['title', 'startDate', 'endDate', 'endBeforeStart'];
        if (activeScopeIdx > 0 && errs.some(e => mainFields.indexOf(e) >= 0)) {
          setScope(0);
        } else if (errs.indexOf('subEventEndBeforeStart') >= 0 && errs.every(e => mainFields.indexOf(e) < 0)) {
          const bad = subEvents.findIndex(se => se.title && se.title.trim() && se.startDate && se.endDate
            && new Date(se.endDate) <= new Date(se.startDate));
          if (bad >= 0) setScope(bad + 1);
        }
      }
      // v29.21 (Audit B4): dieselbe Ebenen-Logik für Schritt 4 — alle dort
      // geprüften Felder liegen auf der Hauptevent-Ebene (Reiter „Klammer/
      // Haupt-Event"). Stand der Reiter auf einem Sub-Event, saßen die roten
      // Markierungen in einem display:none-Block und „Weiter" wirkte tot.
      if (currentStep === 3 && activeCapacityTabIdx !== 0) {
        setScope(0);
      }
      return;
    }
    setTriedNext(false);
    interceptVisibilityCopy(() => setCurrentStep(s => s + 1));
  };
  const attemptSubmitGuarded = (): void => { interceptVisibilityCopy(attemptSubmit); };
  const closeVisCopy = (apply: boolean): void => {
    // Stand als abgehandelt merken — erst eine erneute Klammer-Änderung fragt
    // wieder. Bei „Übernehmen" matchen die Sub-Events danach ohnehin.
    visSnapshotRef.current = visKey(locationFilter, audience, filterMode);
    setVisCopyModalOpen(false);
    if (apply) applyParentVisibilityToSubs();
    const cont = visCopyPendingRef.current;
    visCopyPendingRef.current = null;
    if (cont) cont();
  };

  const fieldHasError = (fieldName: string): boolean => triedNext && getStepErrors().indexOf(fieldName) >= 0;

  const errorBorderStyle = (fieldName: string): React.CSSProperties =>
    fieldHasError(fieldName) ? { borderColor: 'var(--dex-red)', boxShadow: '0 0 0 2px rgba(218,41,28,0.15)' } : {};

  // v7.23: Intro-Hilfsbox pro Wizard-Step. Zeigt eine Liste was der User in
  // diesem Schritt einstellen kann + Verweis aufs Handbuch. DE/EN bilingual.
  // v7.25: pastell-grüner Hintergrund (statt grau), Feature-Items als kompakte
  // Zeilen mit grünem Check-Icon (statt klassischer Disc-Bullets).
  // v7.26: Items in einem auto-fit-Grid (bis zu 3 Spalten ab Wide-Screen),
  // damit die Box nicht extrem lang wird wenn viele Items drin sind.
  // No-Op seit v7.36: die Hint-Box wird nicht mehr inline am Step-Anfang
  // gerendert. Stattdessen liegen die Hints in STEP_HINTS_DE/EN und werden
  // über das i-Icon in der Progress-Bar (Mouseover) angezeigt. Funktion
  // bleibt aus Kompatibilitätsgründen mit den 7 bestehenden Call-Sites.
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const renderStepIntro = (_bulletsDe: string[], _bulletsEn: string[]): React.ReactElement | null => null;

  // v15.0: kleine TabStrip-Komponente für pro-Sub-Event-Tabs in den
  // Schritten 3 (Ort), 4 (Kapazität) und 6 (Felder). Visuell konsistent
  // mit dem Komm-Tab-Pattern in Schritt 7 (grüne Unterstreichung des
  // aktiven Tabs, leichte Hover/Active-Styles). Tab 0 ist immer das
  // Haupt-Event, Tabs 1..N entsprechen `subEvents[0..N-1]`.
  /**
   * v28.73: Quicklink aus dem Klammer-Info-Tooltip. Springt in Schritt 3
   * (Sub-Events) zur Auswahl „Anmeldung zum Hauptevent oder nur zu den
   * Sub-Events" und hebt sie kurz hervor, damit der Organizer die Stelle
   * findet statt sie zu suchen.
   */
  /**
   * v28.96: „Sub-Events bearbeiten" neben dem Kalender. Der Kalender legt die
   * Termine an — bearbeitet werden sie über die Reiter ganz oben, und die
   * stehen nach neun angelegten Tagen weit außerhalb des Sichtfelds. Der Knopf
   * scrollt dorthin und hebt die Leiste kurz hervor, damit klar ist, WO die
   * Bearbeitung stattfindet. Gleiches Muster wie goToSubEventsMode (v28.73).
   */
  const goToScopeBar = (): void => {
    window.setTimeout(() => {
      const el = document.getElementById('dex-scope-bar');
      if (!el) return;
      el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      el.style.transition = 'box-shadow 0.3s';
      el.style.boxShadow = '0 0 0 3px rgba(134,188,37,0.55)';
      el.style.borderRadius = '12px';
      window.setTimeout(() => { el.style.boxShadow = 'none'; }, 2200);
    }, 60);
  };

  const goToSubEventsMode = (): void => {
    // v29.21 (Audit B2): Der Anmelde-Modus-Umschalter (#dex-subevents-mode)
    // liegt seit v28.87 in SCHRITT 1 auf der Klammer-Ebene — der Sprung nach
    // Schritt 3 landete auf „Ort & Programm" ohne den beworbenen Umschalter
    // (Element dort display:none, scrollIntoView lief ins Leere).
    setCurrentStep(0);
    setScope(0);
    setTriedNext(false);
    window.setTimeout(() => {
      const el = document.getElementById('dex-subevents-mode');
      if (!el) return;
      el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      el.style.transition = 'box-shadow 0.3s';
      el.style.boxShadow = '0 0 0 3px rgba(134,188,37,0.55)';
      el.style.borderRadius = '10px';
      window.setTimeout(() => { el.style.boxShadow = 'none'; }, 2200);
    }, 120);
  };

  /**
   * v28.74: Einstellungen eines Sub-Events auf andere übertragen.
   *
   * Hintergrund: Bei Events mit vielen Sub-Events (z.B. neun Tagen) pflegt der
   * Organizer Sichtbarkeit, Kapazität und Fristen bisher NEUN MAL von Hand.
   * Dabei entstehen unbemerkt Abweichungen — im gemeldeten Fall stand der
   * Standortfilter „Berlin" auf einigen Tagen und auf anderen nicht, was erst
   * in der Zusammenfassung auffiel (und dort auch nur, wenn man genau liest).
   * „Vom Hauptevent kopieren" half nicht, weil die Werte ja gerade NICHT vom
   * Hauptevent kommen sollen, sondern vom Geschwister-Sub-Event.
   *
   * Deshalb zwei Dinge: ein Hinweis, der Abweichungen von sich aus meldet,
   * und ein Dialog, in dem der Organizer auswählt, WAS er auf WELCHE
   * Sub-Events überträgt.
   */
  const asRec = (d: SubEventDraft | undefined): Record<string, unknown> =>
    (d || {}) as unknown as Record<string, unknown>;
  /** Anzahl der ANDEREN Sub-Events, bei denen diese Gruppe abweicht. */
  const subGroupDiffCount = (srcIdx: number, fields: string[]): number => {
    const src = asRec(subEvents[srcIdx]);
    let n = 0;
    for (let i = 0; i < subEvents.length; i++) {
      if (i === srcIdx) continue;
      const other = asRec(subEvents[i]);
      const differs = fields.some(f => JSON.stringify(src[f] ?? '') !== JSON.stringify(other[f] ?? ''));
      if (differs) n++;
    }
    return n;
  };
  const applySubTransfer = (): void => {
    if (!subTransfer) return;
    // v28.80: Die Kommunikationsfelder (Logo, Outlook-Text, Betreff …) stehen
    // NICHT laufend im Draft — sie leben im UI-State und werden erst beim
    // Reiterwechsel in den Slot geschrieben. Ohne diesen Flush würde man den
    // Stand VOR der letzten Bearbeitung kopieren. Der Flush schreibt synchron
    // in subEventsRef, deshalb wird von dort gelesen.
    flushActiveCommTabToState();
    const src = asRec(subEventsRef.current[subTransfer.fromIdx] || subEvents[subTransfer.fromIdx]);
    const fields: string[] = [];
    for (const g of SUB_TRANSFER_GROUPS) {
      if (subTransfer.groups.indexOf(g.key) >= 0) fields.push(...g.fields);
    }
    if (fields.length === 0 || subTransfer.targets.length === 0) { setSubTransfer(null); return; }
    setSubEvents(prev => prev.map((s, i) => {
      if (subTransfer.targets.indexOf(i) < 0) return s;
      const patch: Record<string, unknown> = {};
      for (const f of fields) {
        const v = src[f];
        // v28.80: Objekte (z.B. emailTemplateOverrides) klonen — sonst teilen
        // sich alle Ziel-Sub-Events dieselbe Referenz und eine spätere
        // Aenderung an einem würde die anderen mitziehen.
        patch[f] = (v && typeof v === 'object') ? JSON.parse(JSON.stringify(v)) : v;
      }
      return { ...s, ...(patch as unknown as Partial<SubEventDraft>) };
    }));
    const n = subTransfer.targets.length;
    setSubTransfer(null);
    showAlert(isDe
      ? `Einstellungen auf ${n} ${n === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} übertragen. Nicht vergessen zu speichern.`
      : `Settings transferred to ${n} sub-event(s). Don't forget to save.`,
      { variant: 'success' });
  };

  /**
   * v30.60: Kommunikation des Haupt-Events auf ALLE Termine übernehmen.
   *
   * Nutzer-Frage 01.09.2026: „Kann man die Kommunikation für alle Sub-Events
   * auf einmal einstellen?" Bisher gab es nur den Weg Sub → andere Subs
   * (v28.80). Wer die Mails zentral auf der Klammer gestaltet hatte, musste
   * sie erst auf EINEN Termin kopieren, um sie von dort verteilen zu können.
   *
   * Gelesen wird über `resolveTopLevelCommState()` und NICHT direkt aus den
   * State-Variablen: Steht der Organizer gerade auf einem Sub-Reiter, tragen
   * `outlookBody` & Co. die Werte dieses Sub-Events — man würde einen Termin
   * auf alle anderen kopieren und es „Haupt-Event" nennen.
   *
   * Überschrieben wird bewusst ALLES, auch bereits gepflegte Einzelwerte:
   * Genau das ist die Frage, die der Dialog stellt. Objektwerte werden
   * geklont, sonst teilen sich alle Termine dieselbe Referenz (v28.80).
   */
  const applyCommToAllSubEvents = async (): Promise<void> => {
    const named = subEventsRef.current.filter(x => x.title && x.title.trim());
    if (named.length === 0) return;
    const term = childTermPlural || (isDe ? 'Sub-Events' : 'sub-events');
    const ok = await confirmDialog(isDe
      ? `Die Kommunikations-Einstellungen des Haupt-Events (Mail-Sprache, Logo, Outlook-Text, Überschriften, Betreff und alle Mail-Schalter) werden auf ALLE ${named.length} ${term} übertragen.\n\nBereits einzeln gepflegte Werte werden dabei überschrieben. Fortfahren?`
      : `The main event's communication settings will be applied to ALL ${named.length} sub-events. Individually maintained values will be overwritten. Continue?`,
      { title: isDe ? 'Für alle Termine gleich einstellen' : 'Apply to all dates' });
    if (!ok) return;
    // Erst den aktiven Reiter sichern — sonst kopiert man den Stand vor der
    // letzten Bearbeitung (CLAUDE.md: „Kommunikationsfelder der Sub-Events
    // liegen nicht laufend im Draft").
    flushActiveCommTabToState();
    const src = resolveTopLevelCommState() as unknown as Record<string, unknown>;
    const commGroup = SUB_TRANSFER_GROUPS.filter(g => g.key === 'communication')[0];
    const fields = commGroup ? commGroup.fields : [];
    setSubEvents(prev => prev.map(s => {
      if (!s.title || !s.title.trim()) return s;
      const patch: Record<string, unknown> = {};
      for (const f of fields) {
        const v = src[f];
        patch[f] = (v && typeof v === 'object') ? JSON.parse(JSON.stringify(v)) : v;
      }
      return { ...s, ...(patch as unknown as Partial<SubEventDraft>) };
    }));
    showAlert(isDe
      ? `Die Kommunikation des Haupt-Events gilt jetzt für alle ${named.length} ${term}. Nicht vergessen zu speichern.`
      : `The main event's communication now applies to all ${named.length} sub-events. Don't forget to save.`,
      { variant: 'success' });
  };

  /**
   * v28.78: Der Scope-Umschalter (Klammer / Sub-Events) lebt nicht mehr in
   * jedem Schritt, sondern EINMAL global unter der Schritt-Leiste.
   *
   * Vorher hatte jeder Schritt seinen eigenen Reiter-Index — man landete beim
   * Schrittwechsel wieder auf der Klammer und musste sein Sub-Event neu
   * suchen. Und weil der Umschalter im weissen Inhaltsbereich stand, las er
   * sich als Teil des Schritts statt als das, was er ist: die Ebene, auf der
   * gerade gearbeitet wird — sie trägt durch den ganzen Assistenten.
   *
   * Ein gemeinsamer Index bedeutet: Wer auf „Di. 08.09." steht, bleibt auf
   * „Di. 08.09.", auch wenn er von Kapazität zu Feldern wechselt.
   *
   * Gestalterisch bewusst ANDERS als die Schritt-Leiste: Die Schritte sind
   * eine Fortschritts-Spur (Kreise + Linie, ohne Rahmen), der Scope ist eine
   * abgesetzte Kontext-Karte mit eigener Tönung. Zwei Achsen, zwei
   * Formsprachen — sonst liest man sie als zwei Navigationen.
   */
  // v28.89: Schritt 1 ist jetzt ebenfalls scope-fähig. Titel, Start, Ende,
  // Beschreibung und Bild sind DIESELBEN Eingaben — sie zeigen je nach
  // gewählter Ebene auf den Top-Level-State oder auf den gewählten Sub-Event
  // (siehe scopeSub/patchScopeSub unten). Der in v28.81 versuchte Weg (eine
  // zweite Box neben den vorhandenen Feldern) wurde in v28.82 bewusst
  // zurückgenommen: Gleichartige Angaben sollen gleich aussehen und an
  // derselben Stelle stehen.
  const SCOPE_AWARE_STEPS = [0, 2, 3, 4, 5]; // Grundlagen, Ort & Programm, Kapazität, Felder, Kommunikation
  const setScope = (idx: number): void => {
    setActiveScopeIdx(idx);
    setActiveLocationTabIdx(idx);
    setActiveCapacityTabIdx(idx);
    setActiveFieldsTabIdx(idx);
    // Schritt 7 lagert die Kommunikationsfelder pro Reiter ein und aus —
    // deshalb NICHT den State direkt setzen, sondern den Umschalter rufen.
    switchCommTab(idx);
  };

  /**
   * v28.89: Der gerade gewählte Sub-Event-Draft — oder `undefined`, wenn die
   * Klammer/das Hauptevent bearbeitet wird. `scopeSub` ist die einzige Stelle,
   * an der Schritt 1 entscheidet, wohin eine Eingabe geht.
   *
   * Absichern gegen einen Index, der auf einen gelöschten Sub-Event zeigt:
   * `subEvents` kann sich ändern, während der Scope steht (Karte entfernt).
   */
  const scopeSub = activeScopeIdx > 0 ? subEvents[activeScopeIdx - 1] : undefined;
  const patchScopeSub = (patch: Partial<SubEventDraft>): void => {
    const i = activeScopeIdx - 1;
    if (i < 0) return;
    setSubEvents(prev => prev.map((x, k) => (k === i ? { ...x, ...patch } : x)));
  };
  /** Sub-Event-Zeiten liegen als UTC-ISO vor, der Top-Level-State als
   *  Berliner Lokalzeit „YYYY-MM-DDTHH:MM" — beide Richtungen einmal zentral. */
  const localStrToDate = (s: string): Date | null => {
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  };
  const dateToLocalStr = (d: Date | null): string => (d
    ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}T${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`
    : '');
  const subIsoToDate = (iso: string): Date | null => localStrToDate(iso ? isoToLocal(iso) : '');
  const subDateToIso = (d: Date | null): string => (d ? berlinLocalToUtcIso(dateToLocalStr(d)) : '');

  // Die fünf Grundlagen-Felder, an den Scope gebunden.
  const scTitle = scopeSub ? (scopeSub.title || '') : title;
  const setScTitle = (v: string): void => { if (scopeSub) patchScopeSub({ title: v }); else setTitle(v); };
  const scStart = scopeSub ? subIsoToDate(scopeSub.startDate) : localStrToDate(startDate);
  // v29.55: Ganztägig legt die Zeiten auf die Tagesgrenzen — sonst steht in
  // DEX_Events 00:00/00:00 und jede spätere Auswertung hält den Termin für
  // null Minuten lang.
  const clampAllDay = (d: Date | null, end: boolean): Date | null => {
    if (!d || !scAllDay) return d;
    const c = new Date(d);
    if (end) c.setHours(23, 59, 0, 0); else c.setHours(0, 0, 0, 0);
    return c;
  };
  const setScStart = (d: Date | null): void => { const v = clampAllDay(d, false); if (scopeSub) patchScopeSub({ startDate: subDateToIso(v) }); else setStartDate(dateToLocalStr(v)); };
  const scEnd = scopeSub ? subIsoToDate(scopeSub.endDate) : localStrToDate(endDate);
  const setScEnd = (d: Date | null): void => { const v = clampAllDay(d, true); if (scopeSub) patchScopeSub({ endDate: subDateToIso(v) }); else setEndDate(dateToLocalStr(v)); };
  // v29.52: „Ganztägig" hängt am selben Scope wie Start/Ende — der Haken gilt
  // also für den oben gewählten Reiter, nicht global.
  // v29.56: Sammel-Schalter für alle Termine. Bei einer Reihe über zwanzig
  // Tage ist die Einzel-Einstellung zwar richtig, aber nicht zumutbar — man
  // müsste jeden Reiter anfassen. Die beiden Schalter setzen den Wert auf
  // ALLE Sub-Events; die Einzel-Haken bleiben und können danach abweichen.
  const setAllSubsAllDay = (v: boolean): void => {
    setSubEvents(prev => prev.map(se => {
      if (!!se.allDay === v) return se;
      const next: SubEventDraft = { ...se, allDay: v };
      if (v) {
        // Gleiche Klemmung wie beim Einzel-Haken: ohne Tagesgrenzen stünde in
        // DEX_Events später eine Spanne, die keinen ganzen Tag abdeckt.
        const st = subIsoToDate(se.startDate);
        const en = subIsoToDate(se.endDate) || st;
        if (st) {
          const s0 = new Date(st); s0.setHours(0, 0, 0, 0);
          const e0 = new Date(en || st); e0.setHours(23, 59, 0, 0);
          next.startDate = subDateToIso(s0);
          next.endDate = subDateToIso(e0);
        }
      }
      return next;
    }));
  };
  const setAllSubsShowAsFree = (v: boolean): void => {
    setSubEvents(prev => prev.map(se => (!!se.showAsFree === v ? se : { ...se, showAsFree: v })));
  };
  const scShowAsFree = scopeSub ? !!scopeSub.showAsFree : showAsFree;
  const setScShowAsFree = (v: boolean): void => {
    if (scopeSub) patchScopeSub({ showAsFree: v }); else setShowAsFree(v);
  };
  const scAllDay = scopeSub ? !!scopeSub.allDay : allDay;
  const setScAllDay = (v: boolean): void => {
    // Beim Einschalten die Zeiten auf die Tagesgrenzen legen. Der Flow rechnet
    // daraus die Ganztags-Grenzen; leer lassen wäre falsch, weil ein Sub-Event
    // ohne Zeiten seit v28.66 die Zeiten des Hauptevents erbt.
    const dayOf = (d: Date | null): Date | null => d;
    if (scopeSub) {
      const st = subIsoToDate(scopeSub.startDate);
      const en = subIsoToDate(scopeSub.endDate) || st;
      const patch: Partial<SubEventDraft> = { allDay: v };
      if (v && st) {
        const s0 = new Date(st); s0.setHours(0, 0, 0, 0);
        const e0 = new Date((dayOf(en) || st)); e0.setHours(23, 59, 0, 0);
        patch.startDate = subDateToIso(s0);
        patch.endDate = subDateToIso(e0);
      }
      patchScopeSub(patch);
      return;
    }
    setAllDay(v);
    if (v) {
      const st = localStrToDate(startDate);
      const en = localStrToDate(endDate) || st;
      if (st) {
        const s0 = new Date(st); s0.setHours(0, 0, 0, 0);
        const e0 = new Date((en || st)); e0.setHours(23, 59, 0, 0);
        setStartDate(dateToLocalStr(s0));
        setEndDate(dateToLocalStr(e0));
      }
    }
  };
  const scDescription = scopeSub ? (scopeSub.description || '') : description;
  const setScDescription = (v: string): void => { if (scopeSub) patchScopeSub({ description: v }); else setDescription(v); };
  const scImagePreview = scopeSub ? (scopeSub.imagePreview || '') : imagePreview;

  /**
   * v28.91: Termine als Sub-Events.
   *
   * Ein Termin ist KEIN neuer Datentyp — er ist ein ganz normales Sub-Event
   * mit eigener Teilnehmerliste, Kapazität, Frist und Outlook-Termin. Der
   * Kalender ist nur eine schnellere Art, sie anzulegen (neun Tage anklicken
   * statt neun Karten ausfüllen) und auf der Anmeldeseite anzuzeigen. Damit
   * bleiben Flows, `ParentEventId` und die Warteliste unangetastet.
   */
  const dayKeyOfDate = (d: Date): string =>
    `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const dayKeyOfSub = (se: SubEventDraft): string => {
    if (!se.startDate) return '';
    const local = isoToLocal(se.startDate);
    return local ? local.slice(0, 10) : '';
  };
  const dayLabel = (d: Date): string =>
    d.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' });
  /** Leerer Draft mit den Vorgaben, die auch der „Hinzufügen"-Knopf setzt. */
  const makeSubEventDraft = (patch: Partial<SubEventDraft>): SubEventDraft => ({
    id: (typeof crypto !== 'undefined' && crypto.randomUUID) ? crypto.randomUUID() : `se_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    title: '',
    description: '',
    location: '',
    startDate: '',
    endDate: '',
    maxParticipants: 0,
    disableEmails: false,
    disableOutlook: false,
    locationAddress: { street: '', houseNo: '', zip: '', city: '' },
    agenda: [],
    transferTimes: [],
    // v28.20: Hat die Klammer eine explizite Frist, starten neue Sub-Events
    // mit demselben Anmeldeschluss.
    ...(subEventsOnlyMode && klammerDeadline
      ? { registrationDeadline: berlinLocalToUtcIso(klammerDeadline) || '' }
      : {}),
    lastDeregisterDate: '',
    locationFilter: locationFilter,
    audience: audience,
    filterMode: filterMode,
    excludedUsers: [],
    waitlistEnabled: true,
    askSalutation: false,
    ...patch,
  });
  /**
   * Tag im Kalender umschalten. Anlegen ist harmlos; Entfernen kann eine
   * gespeicherte Teilnehmerliste betreffen (`dbId`) und fragt deshalb nach —
   * dieselbe Rückfrage wie das X an der Karte.
   */
  const toggleDaySubEvent = (d: Date | null): void => {
    if (!d) return;
    const key = dayKeyOfDate(d);
    // v29.17: Die An-/Abwahl-Entscheidung fällt IM Funktions-Updater, auf dem
    // tatsächlichen State — nicht auf dem `subEvents` aus der Render-Closure.
    // Bei 20+ Terminen dauert ein Re-Render des Wizards spürbar; wer in der
    // Zeit erneut klickt, dessen zweiter Klick sah vorher noch den ALTEN
    // Stand: Ein eben abgewählter Tag galt als „nicht vorhanden" und wurde
    // wieder angelegt — das Abwählen wirkte „kaputt". Mit dem Updater ist
    // jeder Klick ein echter Toggle auf dem aktuellen Stand.
    //
    // Scope-Korrektur vorab aus der Closure: Steht der Reiter gerade auf dem
    // Tag, der entfernt wird, zurück auf die Klammer. Im (seltenen) Stale-Fall
    // unterbleibt nur die Korrektur — scopeSub sichert den Index ohnehin ab.
    const closureIdx = subEvents.findIndex(se => dayKeyOfSub(se) === key);
    if (closureIdx >= 0 && activeScopeIdx === closureIdx + 1) setScope(0);
    // v29.22: Abwahl eines GESPEICHERTEN Termins → in removedSavedSubs parken
    // (Orange im Kalender, per Klick rückholbar). Ein ORANGE-Tag → Draft aus
    // dem Park zurückholen statt einen neuen anzulegen. Beide Übergänge sind
    // über die Guards in den Updatern idempotent — schnelle Doppelklicks auf
    // veraltetem Render-Stand (v29.17-Falle) können weder doppelt parken noch
    // doppelt anlegen.
    const closureExisting = closureIdx >= 0 ? subEvents[closureIdx] : undefined;
    if (closureExisting) {
      if (closureExisting.dbId) {
        setRemovedSavedSubs(prev => prev.some(x => x.id === closureExisting.id) ? prev : [...prev, closureExisting]);
      }
      setSubEvents(prev => prev.filter(x => x.id !== closureExisting.id));
      return;
    }
    const stashed = removedSavedSubs.find(x => dayKeyOfSub(x) === key);
    if (stashed) {
      setRemovedSavedSubs(prev => prev.filter(x => x.id !== stashed.id));
      setSubEvents(prev => prev.some(x => x.id === stashed.id) ? prev : [...prev, stashed]);
      return;
    }
    setSubEvents(prev => {
      const existingIdx = prev.findIndex(se => dayKeyOfSub(se) === key);
      if (existingIdx >= 0) {
        // v28.96: KEINE Rückfrage je Klick. Im Kalender wird aus- und
        // abgewählt, oft mehrfach hintereinander — ein Modal bei jedem Klick
        // macht genau das unbenutzbar. Der Datenverlust-Hinweis steht ohnehin
        // schon an der richtigen Stelle: beim SPEICHERN listet
        // handleSubmitInner alle Sub-Events auf, die dabei endgültig gelöscht
        // würden (toDelete), und fragt einmal nach. Bis dahin ist nichts
        // passiert. (Stale-Doppelklick: der Tag wurde eben schon behandelt.)
        return prev;
      }
      // v28.92: Der Termin bekommt die UHRZEIT des Hauptevents, gelegt auf
      // diesen Tag — läuft das Event von 9 bis 17 Uhr, gilt das auch für den
      // einzelnen Tag. Ein Ganztags-Block (00:00–23:59) würde den Teilnehmern
      // den kompletten Kalendertag zustellen.
      //
      // Die Zeitfelder bleiben bewusst NICHT leer: Ein Sub-Event ohne Zeiten
      // erbt seit v28.66 die TERMINE des Hauptevents — bei einer Reihe also
      // 01.09.–01.10. für jeden einzelnen Tag statt des Tages selbst.
      //
      // Zwei Fälle, in denen die Uhrzeit nichts hergibt, fallen auf den ganzen
      // Tag zurück: gar keine Zeiten am Hauptevent, und ein mehrtägiges Event,
      // dessen Endzeit nicht nach der Startzeit liegt (z.B. 01.09. 00:00 bis
      // 01.10. 00:00 — daraus liesse sich für einen Tag keine gültige Spanne
      // bauen).
      const timeOf = (v: string): string => {
        const t = (v || '').slice(11, 16);
        return /^\d{2}:\d{2}$/.test(t) ? t : '';
      };
      const startTime = timeOf(startDate);
      const endTime = timeOf(endDate);
      const usable = !!startTime && !!endTime && endTime > startTime;
      const start = berlinLocalToUtcIso(`${key}T${usable ? startTime : '00:00'}`);
      const end = berlinLocalToUtcIso(`${key}T${usable ? endTime : '23:59'}`);
      // v29.52: Der erzeugte Tag erbt „ganztägig" vom Hauptevent — und ist es
      // auch dann, wenn sich aus dem Zeitraum keine Uhrzeit ableiten ließ
      // (`!usable`, z.B. Klammer 01.09. 00:00 – 25.09. 17:00). Genau dieser
      // Fall hat die ganztägigen Blocker erzeugt, über die sich Organizer
      // beschwert haben: 00:00–23:59 sieht in Outlook aus wie ein Tag
      // Vollsperrung, ist aber technisch ein normaler Termin.
      return prev.concat([makeSubEventDraft({
        title: dayLabel(d), startDate: start, endDate: end,
        allDay: allDay || !usable,
      })]);
    });
  };

  /**
   * v29.13: Ein Sub-Event aus der Liste nehmen — mit derselben Rückfrage,
   * egal von wo. Das hing bisher als anonyme Funktion in der Sub-Event-Karte;
   * die Kalender-Liste braucht genau dieselbe Rückfrage, und zwei Kopien
   * derselben Warnung laufen erfahrungsgemäß auseinander.
   */
  const removeSubEventDraft = (se: SubEventDraft): void => {
    (async () => {
      const seTitle = se.title || (isDe ? 'Ohne Titel' : 'Untitled');
      const msg = se.dbId
        ? (isDe
          ? `Sub-Event „${seTitle}" wirklich entfernen? Beim nächsten SPEICHERN wird es endgültig gelöscht — inklusive Teilnehmerliste und aller Anmeldungen (93 Tage im Papierkorb).`
          : `Really remove sub-event "${seTitle}"? On the next SAVE it will be permanently deleted — including its attendee list and all registrations (recycled for 93 days).`)
        : (isDe
          ? `Sub-Event „${seTitle}" entfernen? Die eingetragenen Angaben gehen verloren.`
          : `Remove sub-event "${seTitle}"? The entered details will be lost.`);
      const ok = await confirmDialog(msg, { danger: true, confirmLabel: isDe ? 'Entfernen' : 'Remove' });
      if (ok) {
        // v28.89: Steht der Scope auf einem Reiter hinter dem gelöschten,
        // zeigt er danach auf ein anderes Sub-Event — zurück auf die Klammer,
        // das ist die einzige Ebene, die es sicher noch gibt.
        setScope(0);
        // v29.22: Gespeicherte Termine parken — der Kalender zeigt sie orange
        // als „wird beim Speichern gelöscht" und macht sie rückholbar.
        if (se.dbId) {
          setRemovedSavedSubs(prev => prev.some(x => x.id === se.id) ? prev : [...prev, se]);
        }
        setSubEvents(prev => prev.filter(x => x.id !== se.id));
      }
    })().catch(() => { /* */ });
  };

  const renderPerEventTabStrip = (
    activeIdx: number,
    onChange: (idx: number) => void,
    mainLabel: string,
    ariaLabel: string,
  ): React.ReactElement | null => {
    if (subEvents.length === 0) return null;
    // v22.5: Der „Haupt"/„Klammer"-Badge links im Tab trägt die Rolle bereits —
    // deshalb das doppelte „Klammer: …"/„Haupt-Event: …"-Präfix aus dem Label
    // strippen (sonst stand „KLAMMER  Klammer: …" doppelt da). Sub-Event-Tabs
    // zeigen nur den reinen Sub-Namen (ohne „<Hauptevent> | "-Präfix).
    const strippedMain = mainLabel.replace(/^(Klammer|Bracket|Haupt-Event|Main event):\s*/i, '').trim();
    const tabs: Array<{ label: string; isMain: boolean }> = [
      { label: strippedMain || mainLabel, isMain: true },
      ...subEvents.map(s => ({
        label: (shortSubEventTitle(s.title, title) || (isDe ? 'Sub-Event ohne Titel' : 'Untitled sub-event')).trim(),
        isMain: false,
      })),
    ];
    // v28.72: Geltungsbereich benennen. Die Reiter standen bisher ohne
    // Erklärung da — sie sehen aus wie eine Beschriftung („dieses Event hat
    // 5 Teile"), nicht wie eine Umschaltung. Organizer stellten deshalb alles
    // am ersten Reiter ein und wunderten sich, dass es für die anderen nicht
    // galt; manche merkten gar nicht, dass ihr Event eine Klammer ist. Zwei
    // Ergänzungen, beide am Blick des Nutzers ausgerichtet:
    //  - eine Frage ÜBER den Reitern, die die Bedienung benennt,
    //  - ein Hinweis UNTER den Reitern, direkt über den Feldern: für wen die
    //    Einstellungen gerade gelten und wo die anderen zu finden sind. Der
    //    steht bewusst bei den Feldern, weil dort hingeschaut wird — nicht
    //    oben in der Leiste.
    const subCount = subEvents.length;
    const activeIsMain = activeIdx === 0;
    const activeLabel = (tabs[activeIdx] || tabs[0]).label;
    const mainWord = subEventsOnlyMode ? (isDe ? 'Klammer' : 'bracket') : (isDe ? 'Haupt-Event' : 'main event');
    const otherSubs = activeIsMain ? subCount : subCount - 1;
    const scopeText = ((): React.ReactNode => {
      if (activeIsMain) {
        if (isDe) {
          return (
            <>Du bearbeitest gerade {subEventsOnlyMode ? <>die <strong>Klammer</strong></> : <>das <strong>Haupt-Event</strong></>} „{activeLabel}“.{' '}
              {subEventsOnlyMode
                ? <>Zur Klammer meldet sich niemand direkt an — Teilnehmer wählen eines der Sub-Events. </>
                : null}
              Die Einstellungen auf dieser Seite gelten <strong>ausschließlich für {subEventsOnlyMode ? 'die Klammer' : 'das Haupt-Event'}</strong>. {otherSubs === 1 ? 'Das andere Sub-Event stellst du' : `Die ${otherSubs} Sub-Events stellst du`} oben über {otherSubs === 1 ? 'seinen Reiter' : 'ihre Reiter'} <strong>separat</strong> ein.</>
          );
        }
        return (
          <>You are editing the <strong>{subEventsOnlyMode ? 'bracket' : 'main event'}</strong> „{activeLabel}“.{' '}
            {subEventsOnlyMode ? <>Nobody registers for the bracket itself — attendees pick one of the sub-events. </> : null}
            These settings apply <strong>only to it</strong>. The {otherSubs === 1 ? 'other sub-event' : `${otherSubs} sub-events`} are configured <strong>separately</strong> via {otherSubs === 1 ? 'its tab' : 'their tabs'} above.</>
        );
      }
      if (isDe) {
        return (
          <>Du bearbeitest gerade das <strong>Sub-Event</strong> „{activeLabel}“. Die Einstellungen auf dieser Seite gelten <strong>ausschließlich für dieses Sub-Event</strong>
            {otherSubs > 0
              ? <> — {mainWord === 'Klammer' ? 'die Klammer' : 'das Haupt-Event'} und {otherSubs === 1 ? 'das weitere Sub-Event' : `die ${otherSubs} weiteren Sub-Events`} stellst du oben über die Reiter separat ein.</>
              : <> — {mainWord === 'Klammer' ? 'die Klammer' : 'das Haupt-Event'} stellst du oben über den Reiter separat ein.</>}</>
        );
      }
      return (
        <>You are editing the <strong>sub-event</strong> „{activeLabel}“. These settings apply <strong>only to it</strong>
          {otherSubs > 0
            ? <> — the {subEventsOnlyMode ? 'bracket' : 'main event'} and the {otherSubs === 1 ? 'other sub-event' : `${otherSubs} other sub-events`} are configured separately via the tabs above.</>
            : <> — the {subEventsOnlyMode ? 'bracket' : 'main event'} is configured separately via its tab above.</>}</>
      );
    })();
    // v22.30: Rendering + Sticky-Pin + gefüllter Aktiv-Tab leben in der
    // Modul-Komponente StickyTabStrip (Hooks pro Instanz).
    return (
      <>
        <div style={{
          fontSize: '0.78rem', fontWeight: 700, letterSpacing: '0.03em',
          textTransform: 'uppercase', color: 'var(--dex-gray-500)', marginBottom: 6,
        }}>
          {isDe ? 'Welches (Sub-)Event bearbeitest du gerade?' : 'Which (sub-)event are you editing?'}
        </div>
        <StickyTabStrip
          tabs={tabs}
          activeIdx={activeIdx}
          onChange={onChange}
          ariaLabel={ariaLabel}
          mainBadge={subEventsOnlyMode ? (isDe ? 'Klammer' : 'Bracket') : (isDe ? 'Haupt' : 'Main')}
          klammer={subEventsOnlyMode}
          klammerWord={isDe ? 'Klammerevent' : 'bracket event'}
          // v29.23: Zähl-Badge rechts in der Klammer-Zeile — ersetzt die frei
          // schwebende Zahl neben den umbrechenden Reitern. Im Kalender-Modus
          // sind die Kinder „Termine", sonst gilt die Event-Bezeichnung.
          countBadge={subCount >= 2 ? `${subCount} ${subEventCalendar
            ? (isDe ? 'Termine' : 'dates')
            : (childTermPlural || (isDe ? 'Sub-Events' : 'sub-events'))}` : undefined}
          klammerInfo={
            <InfoTooltip
              placement="bottom"
              interactive
              text={isDe ? (
                <>
                  <strong>Klammerevent</strong> — zu diesem Event selbst meldet sich <strong>niemand</strong> an. Teilnehmer sehen nur die {childTermPlural || 'Sub-Events'} darunter und melden sich <strong>dort</strong> an. Der Eventname ist die Klammer darüber: Er erscheint in der Übersicht und fasst die {childTermPlural || 'Sub-Events'} zusammen.<br /><br />
                  Deshalb gibt es hier keine eigene Teilnehmerzahl und keine eigene Warteliste — beides pflegst du je {childTermSingular || 'Sub-Event'}.<br /><br />
                  <strong>Du willst, dass man sich auch zum Hauptevent anmelden kann?</strong> Dann stell die Anmeldung in Schritt 1 (&bdquo;Grundlagen&ldquo;) um —{' '}
                  <button
                    type="button"
                    onClick={() => goToSubEventsMode()}
                    style={{
                      background: 'none', border: 'none', padding: 0, font: 'inherit',
                      color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline',
                      cursor: 'pointer', fontWeight: 700,
                    }}
                  >
                    hier direkt hinspringen
                  </button>.
                </>
              ) : (
                <>
                  <strong>Bracket event</strong> — <strong>nobody</strong> registers for this event itself. Attendees only see the sub-events below and register <strong>there</strong>. The event name is the bracket around them: it appears in the overview and groups the sub-events.<br /><br />
                  That is why there is no capacity and no waitlist at this level — you set both per sub-event.<br /><br />
                  <strong>Want people to be able to register for the main event too?</strong> Then change the registration mode in step 1 —{' '}
                  <button
                    type="button"
                    onClick={() => goToSubEventsMode()}
                    style={{
                      background: 'none', border: 'none', padding: 0, font: 'inherit',
                      color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline',
                      cursor: 'pointer', fontWeight: 700,
                    }}
                  >
                    jump there directly
                  </button>.
                </>
              )}
            />
          }
        />
        <div style={{
          margin: '-6px 0 16px', padding: '10px 12px', borderRadius: 8,
          background: 'var(--dex-gray-50, #f8f9fa)',
          border: '1px solid var(--dex-gray-200)',
          borderLeft: '4px solid var(--dex-green, #86bc25)',
          fontSize: '0.82rem', lineHeight: 1.55, color: 'var(--dex-gray-700)',
        }}>
          {scopeText}
        </div>
      </>
    );
  };

  // Nutzungsbedingungen-Modal: zeigt sich beim ersten Aufruf der
  // Create-Event-Seite. Nach Akzeptieren wird die Maske weggeklappt; bei
  // Bearbeiten bestehender Events (isEditMode) wird sie nicht angezeigt.
  const showTermsModal = !isEditMode && !tcAccepted;

  // v10.23: Zebra-Hintergrund für Schritt-3-Blöcke (Kapazität & Sichtbarkeit).
  // Counter wird pro Render zurückgesetzt; conditional Blöcke verschieben den
  // Index nur wenn sie tatsächlich rendern, sodass die Alternation auch dann
  // sauber bleibt, wenn Filterverknüpfung oder Sichtbarkeit-prüfen-Buttons
  // ausgeblendet sind. Wird bewusst NUR in Schritt 3 verwendet — andere Steps
  // bekommen das in einer späteren Iteration nachgezogen.
  const zebraS3Bg = (): string => {
    // v22.29: Aufräum-Pass — kein Zebra-Wechsel mehr (wirkte unruhig).
    // v22.38: Sektions-Flächen wieder NEUTRAL grau (konsistent zu allen
    // anderen Wizard-Schritten) — pastellgrün ist seit v22.36 den
    // AUSGEFÜLLTEN Eingaben vorbehalten (.dex-filled), sonst konkurrieren
    // die Farben. Funktion bleibt die zentrale Farb-Stelle.
    return 'var(--dex-gray-50, #fafafa)';
  };

  // v30.66: Props-Buendel fuer die ausgelagerten Wizard-Schritte. Bewusst hier,
  // direkt vor dem return — davor stehen alle Deklarationen, ein Buendel weiter
  // oben waere ein TDZ-Fehler auf die spaeter deklarierten Handler.
  const basicsStepProps = {
    activeFrom, activeScopeIdx, applyDraftPayload, applyEventTemplate, childEventsOf, childTermSingular,
    currentStep, currentUser, dayKeyOfDate, description, DRAFT_KEY, draftSavedAt,
    editEvent, emailLogoPreview, errorBorderStyle, events, fieldHasError, fileToBase64,
    imageBanner, imageDisplay, imageDisplayOpen, imageEditOpen, imageFile, imageOrigFile,
    imagePreview, imageUploadError, isDe, isEditMode, isFictive, location,
    logoCropTarget, noDescription, outlookLogoPreview, patchScopeSub, pendingDraft, previewBeforeActive,
    renderStepIntro, scAllDay, scDescription, scEnd, scImagePreview, scopeSub,
    scShowAsFree, scStart, scTitle, setActiveFrom, setDescription, setEmailLogoFromPhoto,
    setEmailLogoPreview, setEventImageUrl, setHtmlEditorMode, setHtmlEditorOpen, setImageBanner, setImageDisplay,
    setImageDisplayOpen, setImageEditOpen, setImageFile, setImageOrigAspect, setImageOrigFile, setImagePreview,
    setImageUploadError, setIsFictive, setLogoCropTarget, setNoDescription, setOutlookLogoFromPhoto, setOutlookLogoPreview,
    setPendingDraft, setPreviewBeforeActive, setScAllDay, setScEnd, setScShowAsFree, setScStart,
    setScTitle, setShowDemoVariantModal, setShowTemplatePicker, setSubEvents, setSubImageCropIdx, showTemplatePicker,
    shrinkLogoB64, startDate, subEvents, subEventsOnlyMode, t, templateLoadingId,
    title, wizardImgAspect, zebraS3Bg,
  };
  const detailsStepProps = {
    contactEmail, contactExpanded, contactInfo, contactName, contactOrganizerEmail, currentStep,
    errorBorderStyle, hiddenOrganizerEmails, hideOrganizer, hideOrganizerIndividualOnly, isDe, isSearchingOrganizer,
    location, organizer, organizerDisplayLarge, organizerEmails, organizerIncludeIntl, organizerResults,
    organizerSearch, organizerTimerRef, qrScannerEmails, qrScannerIncludeIntl, qrScannerNames, qrScannerResults,
    qrScannerSearch, qrScannerTimerRef, searchUsers, setBulkOrganizerOpen, setBulkQrScannerOpen, setBulkTestTeamOpen,
    setContactEmail, setContactExpanded, setContactInfo, setContactName, setContactOrganizerEmail, setHideOrganizer,
    setHideOrganizerIndividualOnly, setOrganizer, setOrganizerDisplayLarge, setOrganizerEmails, setOrganizerIncludeIntl, setOrganizerResults,
    setOrganizerSearch, setQrScannerEmails, setQrScannerIncludeIntl, setQrScannerNames, setQrScannerResults, setQrScannerSearch,
    setTestTeamEmails, setTestTeamIncludeIntl, setTestTeamNames, setTestTeamResults, setTestTeamSearch, startDate,
    t, testTeamEmails, testTeamIncludeIntl, testTeamNames, testTeamResults, testTeamSearch,
    testTeamTimerRef, title, toggleOrganizerHidden,
  };
  const locationProgramStepProps = {
    activeLocationTabIdx, addAgendaItem, addrCity, addrHouseNo, addrStreet, addrZip,
    agenda, currentStep, isDe, isMobile, isoToLocal, location,
    locationOptions, onlineMeetingMode, outlookLocationOverride, removeAgendaItem, renderStepIntro, setAddrCity,
    setAddrHouseNo, setAddrStreet, setAddrZip, setLocation, setOnlineMeetingMode, setOutlookLocationOverride,
    setSubEvents, setTeamsLink, setTransferTimes, startDate, subEvents, t,
    teamsLink, transferTimes, updateAgendaItem,
  };
  const subEventsSectionProps = {
    activeScopeIdx, audience, berlinLocalToUtcIso, childGender, childTermPlural, childTermSingular,
    confirmDialog, currentStep, customTermMode, dayKeyOfSub, endDate, filterMode,
    goToScopeBar, isDe, isoToLocal, klammerDeadline, locationFilter, mainEventLabel,
    mainEventLabelMode, openRuleDays, openRuleEnabled, openRuleMode, orgGetsSubInvites, orgInvitesTouchedRef,
    removedSavedSubs, removeSubEventDraft, requireSubEventSelection, setAllSubsAllDay, setAllSubsShowAsFree, setChildGender,
    setChildTermPlural, setChildTermSingular, setCustomTermMode, setEndDate, setMainEventLabel, setMainEventLabelMode,
    setOrgGetsSubInvites, setRemovedSavedSubs, setRequireSubEventSelection, setScope, setStartDate, setSubEventCalendar,
    setSubEvents, setSubEventSingleChoice, setSubEventsOnlyMode, setSubEventsOptIn, setSubImageCropIdx, setTerminListOpen,
    startDate, subEventCalendar, subEvents, subEventSingleChoice, subEventsOnlyMode, subEventsOptIn,
    subImageCropIdx, t, terminListOpen, title, toggleDaySubEvent,
  };
  const capacityStepProps = {
    activeCapacityTabIdx, activeFrom, assistantsCanSee, audience, b2runStartblocks, berlinLocalToUtcIso,
    cancelRuleAfter, cancelRuleAmount, cancelRuleEnabled, cancelRuleUnit, childTermPlural, childTermSingular,
    currentStep, durchstarterCapacity, durchstarterStartblock, effectiveKlammerDeadline, errorBorderStyle, excludedUsers,
    fieldHasError, filterMode, funstarterCapacity, funstarterStartblock, hauptGreyoutWrapperStyle, isDe,
    isVisOpen, klammerDeadline, lastDeregisterDate, locationFilter, locationOptions, maxParticipants,
    noCancelAfterDeadline, openRuleDays, openRuleEnabled, openRuleFixedDate, openRuleMode, registrationDeadline,
    regRuleAmount, regRuleEnabled, regRuleUnit, renderHauptGreyoutBanner, renderKlammerVisibilityMismatch, renderStepIntro,
    renderVisibilitySummaryBox, rollingDeadlineIso, setActiveCapacityTabIdx, setActiveFrom, setAssistantsCanSee, setAudience,
    setCancelRuleAfter, setCancelRuleAmount, setCancelRuleEnabled, setCancelRuleUnit, setDurchstarterCapacity, setDurchstarterStartblock,
    setExcludedUsers, setFilterMode, setFunstarterCapacity, setFunstarterStartblock, setKlammerDeadline, setLastDeregisterDate,
    setLocationFilter, setMaxParticipants, setNoCancelAfterDeadline, setOpenRuleDays, setOpenRuleEnabled, setOpenRuleFixedDate,
    setOpenRuleMode, setRegistrationDeadline, setRegRuleAmount, setRegRuleEnabled, setRegRuleUnit, setSplitDescA,
    setSplitDescB, setSplitDisplayOrderReversed, setSplitHelpText, setSplitLabelA, setSplitLabelB, setSplitSectionTitle,
    setSplitSharedWaitlist, setSubEvents, setSubTransfer, setUnlimitedParticipants, setUserCancelAllowed, setUseSplitCapacities,
    setVisAllSubs, setWaitlistEnabled, splitDescA, splitDescB, splitDisplayOrderReversed, splitHelpText,
    splitLabelA, splitLabelB, splitSectionTitle, splitSharedWaitlist, SUB_TRANSFER_GROUPS, subEvents,
    subEventsOnlyMode, subEventsOptIn, subGroupDiffCount, t, title, unlimitedParticipants,
    userCancelAllowed, useSplitCapacities, visAllSubs, visAllSubsTouchedRef, visHeader, waitlistEnabled,
    zebraS3Bg,
  };
  const fieldsStepProps = {
    activeFieldsTabIdx, addCustomField, addStartblock, addSubEventCustomField, askSalutation, b2runStartblocks,
    bilingualFields, childTermPlural, confirmDialogEnabled, confirmDialogMode, confirmDialogText, copyParentFieldsToSubEvent,
    currentStep, customFields, dragFieldId, dragOverFieldId, fieldExpandOverride, isDe,
    moveCustomField, newStartblock, openSuggestedModal, registrationLanguage, removeCustomField, removeStartblock,
    removeSubEventCustomField, renderShowIfConfig, renderStepIntro, reorderMode, setAskSalutation, setBilingualFields,
    setConfirmDialogEnabled, setConfirmDialogMode, setConfirmDialogText, setCustomFields, setDragFieldId, setDragOverFieldId,
    setNewStartblock, setRegistrationLanguage, setReorderMode, setSubEvents, splitLabelA, splitLabelB,
    subEvents, subEventsOnlyMode, t, title, toggleFieldExpand, updateCustomField,
    updateSubEventCustomField, useSplitCapacities,
  };
  const communicationStepProps = {
    activeCommTabIdx, applyCommToAllSubEvents, applyEventPhotoToLogo, autoDeregisterOnDecline, bundledComm, childTermPlural,
    commToggleRow, confirmDialog, currentStep, disableCancellationEmail, disableEmails, disableOutlook,
    disableRegistrationEmail, durchstarterCapacity, effectiveHeaderImage, emailLanguage, emailLogoFromPhoto, emailLogoPreview,
    emailTemplateOverrides, emailTemplates, funstarterCapacity, imageFile, imagePreview, inactiveHandling,
    isDe, mainCommDisabledAck, maxParticipants, notifyOrgCancelMode, notifyOrgRegisterFromDate, notifyOrgRegisterMode,
    offerLogoToSubEvents, organizer, outlookBody, outlookLogoFromPhoto, outlookLogoPreview, renderHeaderSizeControl,
    renderOutlookUpdateButton, renderStepIntro, setAutoDeregisterOnDecline, setBundledComm, setDisableCancellationEmail, setDisableEmails,
    setDisableOutlook, setDisableRegistrationEmail, setEmailLanguage, setEmailLogoFromPhoto, setEmailLogoPreview, setEmailTemplateOverrides,
    setHtmlEditorMode, setHtmlEditorOpen, setHtmlEditorTemplateType, setInactiveHandling, setLogoCropTarget, setMainCommDisabledAck,
    setNotifyOrgCancelMode, setNotifyOrgRegisterFromDate, setNotifyOrgRegisterMode, setOutlookLogoFromPhoto, setOutlookLogoPreview, setSubTransfer,
    subEvents, subEventsOnlyMode, t, title, unlimitedParticipants, useSplitCapacities,
    waitlistEnabled,
  };
  const wizardTermsModalProps = {
    canBilling, goBack, internalCheckbox, isDe, setBillingPromptOpen, setInternalCheckbox,
    setTcAccepted, setTcCheckbox, setTcExpanded, showTermsModal, tcCheckbox, tcExpanded,
  };
  const wizardModalsProps = {
    allowAttendeeUpload, askTeamName, attendeeUploadHint, attendeeUploadLabel, contactInfo, contactName,
    notifyOrgCancelMode, notifyOrgRegisterFromDate, notifyOrgRegisterMode, quizClusterSize, splitDescA, splitDescB,
    splitDisplayOrderReversed, splitHelpText, splitSectionTitle, teamJoinRequiresApproval, teamOpenSlotsVisible, teamPartialAllowed,
    activeCommTabIdx, activeFrom, addrCity, addrHouseNo, addrStreet, addrZip,
    addSelectedSuggestedFields, agenda, applySubTransfer, askSalutation, attemptSubmit, audience,
    autoDeregisterOnDecline, berlinLocalToUtcIso, bilingualFields, buildDraftPayload, bulkOrganizerOpen, bulkQrScannerOpen,
    bulkTestTeamOpen, cancelOutlookSave, childTermPlural, childTermSingular, closeVisCopy, confirmOutlookSave,
    contactEmail, customFields, DEMO_VARIANTS, description, disableCancellationEmail, disableEmails,
    disableOutlook, disableRegistrationEmail, documents, DRAFT_KEY, dragOverSectionId, dragSectionId,
    durchstarterCapacity, emailLanguage, emailLogoPreview, emailTemplateOverrides, emailTemplates, endDate,
    eventImageUrl, excludedUsers, filterMode, funstarterCapacity, headerImageLayout, htmlEditorMode,
    htmlEditorOpen, htmlEditorTemplateType, imagePreview, inactiveHandling, isDe, isEditMode,
    isFictive, isMobile, isoToLocal, lastDeregisterDate, location, locationFilter,
    maxParticipants, newSectionError, newSectionModalOpen, newSectionName, organizer, organizerEmails,
    outlookBody, outlookConfirmChecks, outlookConfirmItems, outlookConfirmOpen, outlookEndOverride, outlookHeading,
    outlookLocationOverride, outlookLogoPreview, outlookStartOverride, outlookSubheading, outlookSubject, pendingSections,
    pendingSuccessDispatch, pendingSuccessDispatchRef, previewSections, qrScannerEmails, qrScannerNames, quiz,
    registrationDeadline, registrationLanguage, renderPreviewSection, requireSubEventSelection, scDescription, scopeSub,
    searchUsers, setBulkOrganizerOpen, setBulkQrScannerOpen, setBulkTestTeamOpen, setDragOverSectionId, setDragSectionId,
    setEmailTemplateOverrides, setHeaderImageLayout, setHtmlEditorOpen, setNewSectionError, setNewSectionModalOpen, setNewSectionName,
    setOrganizer, setOrganizerEmails, setOutlookBody, setOutlookConfirmChecks, setOutlookEndOverride, setOutlookHeading,
    setOutlookLocationOverride, setOutlookStartOverride, setOutlookSubheading, setOutlookSubject, setPendingSections, setPendingSuccessDispatch,
    setPreviewSections, setQrScannerEmails, setQrScannerNames, setScDescription, setShowB2runSuggested, setShowConfigCheck,
    setShowDemoVariantModal, setShowPreview, setShowRegisterPreview, setShowSuggestedModal, setShowSummaryModal, setSubEvents,
    setSubTransfer, setSuggestedSelection, setTestTeamEmails, setTestTeamNames, setUnsavedConfirmOpen, showB2runSuggested,
    showConfigCheck, showDemoVariantModal, showPreview, showRegisterPreview, showSuggestedModal, showSummaryModal,
    splitLabelA, splitLabelB, splitSharedWaitlist, startDate, SUB_TRANSFER_GROUPS, subEvents,
    subEventsOnlyMode, subGroupDiffCount, subTransfer, SUGGESTED_FIELDS_CATALOG, suggestedSelection, t,
    teamRegistrationEnabled, teamSize, testTeamEmails, testTeamNames, title, transferTimes,
    unlimitedParticipants, unsavedConfirmOpen, useSplitCapacities, visCopyModalOpen, waitlistEnabled,
  };
  return (
    <div ref={wizardRootRef} className="page-container" style={{ maxWidth: 1100, marginLeft: 'auto', marginRight: 'auto' }}>
      <WizardTermsModal {...wizardTermsModalProps} />
      <div>
        {/* ===== Step Progress Bar =====
            v14.8: drei Layout-Fixes für das 9-Schritt-Layout:
            (1) Linie endet exakt auf der Mittelachse des ersten/letzten
                Kreises — vorher fix `left/right: 10%`, was zufällig nur
                für n=8 stimmte; jetzt dynamisch über `100 / (steps.length * 2)`.
            (2) Linie etwas dicker (5 statt 3 px) + abgerundet — sonst
                verschwindet sie bei 9 Schritten optisch.
            (3) `alignItems: flex-start` statt `center` — sonst rutschen
                Kreise nach unten, wenn ein Label (z.B. „Kapazität &
                Sichtbarkeit") umbricht.
            Die Linie sitzt bei top=17, height=5 (Mitte bei 19.5 px) —
            das deckt sich exakt mit der Mitte der 40-px-Kreise. */}
        <div style={{ marginBottom: 32 }}>
          {/* v22.22: Hover-Effekt auf den Schritt-Punkten — hebt den Schritt
              leicht an und färbt Kreis-Rand + Label grün, damit die
              Klickbarkeit sofort erkennbar ist. */}
          <style>{`
            .dex-wizard-step { transition: transform 0.15s ease; }
            .dex-wizard-step:hover { transform: translateY(-2px); }
            .dex-wizard-step:hover .dex-step-circle { border-color: var(--dex-green, #86bc25) !important; box-shadow: 0 4px 12px rgba(134,188,37,0.35) !important; }
            .dex-wizard-step:hover .dex-step-label { color: var(--dex-green-dark, #4a7c1f) !important; }
            /* v22.30: Gefüllter grüner Schritt-Header — sitzt bündig als
               Kopf der weißen Karte (negative Margins überbrücken das
               Karten-Padding): oben rund wie die Karte, unten gerade Kante,
               darunter beginnt der Schritt-Inhalt. */
            .dex-step-head-title {
              margin: -32px -32px 0; padding: 16px 24px 4px;
              background: var(--dex-green, #86bc25); color: #fff;
              font-size: 1.3rem; font-weight: 700;
              border-radius: 15px 15px 0 0;
            }
            .dex-step-head-lead {
              margin: 0 -32px 20px; padding: 0 24px 14px;
              background: var(--dex-green, #86bc25); color: rgba(255,255,255,0.95);
              font-size: 0.85rem; line-height: 1.55;
              border-radius: 0;
            }
            @media (max-width: 768px) {
              .dex-step-head-title { margin: -20px -16px 0; padding: 14px 16px 4px; }
              .dex-step-head-lead { margin: 0 -16px 16px; padding: 0 16px 12px; }
            }
            /* v29.7: Zwischen-Trenner INNERHALB eines Schritts — gleicher
               grüner Balken wie der Schritt-Kopf, nur ohne die runden Ecken
               oben (die gehören dem Kartenanfang) und eine Spur kleiner, damit
               der Schritt-Kopf die Überschrift bleibt. Der frühere Trenner war
               eine graue Haarlinie mit Kleinschrift; die trennt zu leise für
               den Themenwechsel „was ist das Event" → „woraus besteht es". */
            .dex-step-sub-head {
              margin: 32px -32px 0; padding: 13px 24px 3px;
              background: var(--dex-green, #86bc25); color: #fff;
              font-size: 1.1rem; font-weight: 700;
            }
            .dex-step-sub-lead {
              margin: 0 -32px 20px; padding: 0 24px 13px;
              background: var(--dex-green, #86bc25); color: rgba(255,255,255,0.95);
              font-size: 0.85rem; line-height: 1.55;
            }
            @media (max-width: 768px) {
              .dex-step-sub-head { margin: 24px -16px 0; padding: 12px 16px 3px; }
              .dex-step-sub-lead { margin: 0 -16px 16px; padding: 0 16px 11px; }
            }
            /* v22.36: Ausgefüllte Eingaben — pastellgrün wie auf der
               Anmeldeseite (Klasse wird per Sweep/Listener getoggelt). */
            .dex-filled:not(:focus) {
              border-color: var(--dex-green, #86bc25) !important;
              background: rgba(134,188,37,0.07) !important;
            }
          `}</style>
          {(() => {
            const sidePct = 100 / (steps.length * 2);
            const spanPct = 100 - 2 * sidePct;
            return (
          <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', position: 'relative' }}>
            {/* Verbindungslinie */}
            <div style={{ position: 'absolute', top: 17, left: `${sidePct}%`, right: `${sidePct}%`, height: 5, background: 'var(--dex-gray-200)', borderRadius: 3, zIndex: 0 }} />
            <div style={{ position: 'absolute', top: 17, left: `${sidePct}%`, height: 5, background: 'var(--dex-green)', borderRadius: 3, zIndex: 1, width: `${(currentStep / Math.max(1, steps.length - 1)) * spanPct}%`, transition: 'width 0.4s ease' }} />
            {steps.map((step, idx) => (
              <div
                key={idx}
                className="dex-wizard-step"
                data-tour={`wizard-step-${idx}`}
                onClick={() => {
                  // v29.21 (Audit B3): Zurück ist immer frei; nach vorn nur,
                  // wenn ALLE übersprungenen Schritte fehlerfrei sind. Vorher
                  // prüfte der Klick nur den aktuellen Schritt — ein Sprung
                  // von Schritt 1 direkt auf 9 umging z.B. die Organizer-
                  // Pflicht aus Schritt 2, und handleSubmitInner prüft nur den
                  // Titel: Ein Event ohne Organizer war anlegbar. Bei einem
                  // Fehler springt der Wizard auf den ersten fehlerhaften
                  // Schritt, damit die Markierungen sichtbar sind.
                  if (idx <= currentStep) { setCurrentStep(idx); return; }
                  setTriedNext(true);
                  for (let st = currentStep; st < idx; st++) {
                    if (getStepErrorsFor(st).length > 0) { setCurrentStep(st); return; }
                  }
                  setTriedNext(false);
                  setCurrentStep(idx);
                }}
                style={{
                  display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8,
                  zIndex: 2, cursor: 'pointer',
                  flex: 1,
                  // v29.66: ausgegrauter Abrechnungs-Schritt (s. steps-Array).
                  // v30.28: NIE den Schritt ausgrauen, auf dem man gerade steht.
                  // Der aktive Kreis ist grün mit weißer Ziffer — bei 50 %
                  // Deckkraft wird daraus Weiß auf Blassgrün, also unlesbar.
                  // Das Ausgrauen soll „optional" sagen, nicht „unlesbar".
                  opacity: ((step as { dim?: boolean }).dim && idx !== currentStep) ? 0.5 : 1,
                }}
              >
                <div className="dex-step-circle" style={{
                  width: 40, height: 40, borderRadius: '50%',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  fontWeight: 700, fontSize: '1rem',
                  background: idx <= currentStep ? 'var(--dex-green)' : '#fff',
                  color: idx <= currentStep ? '#fff' : 'var(--dex-gray-400)',
                  border: idx <= currentStep ? '3px solid var(--dex-green)' : '3px solid var(--dex-gray-200)',
                  transition: 'all 0.3s ease',
                  boxShadow: idx === currentStep ? '0 0 0 4px rgba(134,188,37,0.2)' : 'none',
                }}>
                  {idx < currentStep ? '✓' : step.icon}
                </div>
                <span className="dex-step-label" style={{
                  fontSize: '0.75rem', fontWeight: idx === currentStep ? 700 : 500,
                  color: idx <= currentStep ? 'var(--dex-green)' : 'var(--dex-gray-400)',
                  transition: 'color 0.3s ease',
                  textAlign: 'center',
                }}>
                  {step.label}
                </span>
                {/* v9.27/v9.37: i-Icon UNTER dem Step-Label (vorher inline rechts daneben).
                    Hover zeigt die Hints für diesen Step.
                    v9.37: Styling identisch zur InfoTooltip-Komponente (serif, 20x20,
                    1.5px-Border) — sonst wirkt das wizard-i im Vergleich klobig. */}
                <span
                  role="button"
                  tabIndex={0}
                  aria-label={isDe ? 'Hinweise zu diesem Schritt' : 'Hints for this step'}
                  onMouseEnter={() => setHintStepIdx(idx)}
                  onMouseLeave={() => setHintStepIdx(null)}
                  onFocus={() => setHintStepIdx(idx)}
                  onBlur={() => setHintStepIdx(null)}
                  onClick={e => { e.stopPropagation(); setHintStepIdx(prev => prev === idx ? null : idx); }}
                  style={{
                    position: 'relative',
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    width: 20, height: 20, borderRadius: '50%',
                    background: hintStepIdx === idx ? 'var(--dex-gray-100, #f0f0f0)' : 'transparent',
                    color: 'var(--dex-gray-700, #555)',
                    border: `1.5px solid ${hintStepIdx === idx ? 'var(--dex-gray-700, #555)' : 'var(--dex-gray-500, #888)'}`,
                    fontSize: '0.7rem', fontWeight: 700, fontFamily: 'serif',
                    cursor: 'help',
                    marginTop: 4,
                    userSelect: 'none',
                    transition: 'background 0.15s, border-color 0.15s',
                  }}
                >
                  i
                  {hintStepIdx === idx && (
                    <div
                      role="tooltip"
                      style={{
                        // v9.40: Styling 1:1 wie InfoTooltip (siehe InfoTooltip.tsx),
                        // damit die zwei Tooltip-Varianten optisch konsistent wirken.
                        position: 'absolute',
                        top: 'calc(100% + 8px)',
                        left: '50%',
                        transform: 'translateX(-50%)',
                        width: 'max-content',
                        maxWidth: 480,
                        minWidth: 280,
                        background: 'rgba(40,40,40,0.96)',
                        color: '#fff',
                        padding: '12px 16px',
                        borderRadius: 8,
                        boxShadow: '0 6px 18px rgba(0,0,0,0.28)',
                        // v15: explizit Sans-Serif — vorher 'inherit', was
                        // den serif-Font des parent „i"-Icons (s. unten
                        // fontFamily:'serif' für das i-Glyph) übernommen
                        // hat und den ganzen Tooltip Times-artig erscheinen
                        // ließ. Jetzt 1:1 wie InfoTooltip.
                        fontFamily: 'Aptos, "Open Sans", "Segoe UI", Arial, Helvetica, sans-serif',
                        fontSize: '0.82rem',
                        lineHeight: 1.55,
                        fontWeight: 400,
                        fontStyle: 'normal',
                        textAlign: 'left',
                        whiteSpace: 'normal',
                        zIndex: 1500,
                        pointerEvents: 'none',
                      }}
                    >
                      <div style={{ fontWeight: 700, marginBottom: 8, color: 'rgba(255,255,255,0.92)' }}>
                        {isDe ? 'Was ich hier einstellen kann' : 'What I can configure here'}
                      </div>
                      <ul style={{ margin: 0, paddingLeft: 18 }}>
                        {(isDe ? STEP_HINTS_DE : STEP_HINTS_EN)[idx]?.map((b, bi) => (
                          <li key={bi} style={{ marginBottom: 4 }}>{b}</li>
                        ))}
                      </ul>
                    </div>
                  )}
                </span>
              </div>
            ))}
          </div>
            );
          })()}
        </div>

        {/* v28.78: Scope-Karte zwischen Schritt-Leiste und Formular — eine
            Ebene für „für wen gilt das hier?", die durch alle Schritte trägt. */}
        {renderGlobalScopeBar()}

        {/* v30.1: Autosave-Anzeige der Neu-Anlage — Speicher-Symbol plus
            Zeitstempel der letzten Zwischenspeicherung, auf jedem Schritt
            sichtbar. Nur Neu-Anlage: im Edit-Modus gibt es keinen
            Entwurfs-Zwischenspeicher (s. v30.0-Block). */}
        {!isEditMode && draftSavedAt !== null && (
          <div style={{ display: 'flex', justifyContent: 'flex-end', margin: '6px 2px 2px' }}>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" style={{ flexShrink: 0 }}>
                <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
                <polyline points="17 21 17 13 7 13 7 21" />
                <polyline points="7 3 7 8 15 8" />
              </svg>
              {isDe
                ? `Zwischengespeichert am ${new Date(draftSavedAt).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })} um ${new Date(draftSavedAt).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })} Uhr`
                : `Auto-saved on ${new Date(draftSavedAt).toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })} at ${new Date(draftSavedAt).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })}`}
            </span>
          </div>
        )}

        {/* ===== Formular ===== */}
        <div>
          <div className="card" style={{ borderRadius: 16, boxShadow: '0 2px 12px rgba(0,0,0,0.06)' }}>
            <div className="creation-form">
              {/* v22.30: marginBottom 48 kompensiert die -32px-Top-Margin
                  des grünen Schritt-Headers darunter (netto 16px Abstand). */}
              {error && (
                <div style={{ padding: '10px 16px', background: '#fce4ec', color: '#c62828', borderRadius: 8, marginBottom: 48, fontSize: '0.85rem' }}>
                  {error}
                </div>
              )}

              {/* ===== Schritt 1: Grundlagen =====
                  v9.32: 1-basierte UI-Nummerierung (in der Logik bleibt
                  currentStep 0-basiert) — siehe ENTWICKLUNG.md. */}
              <BasicsStep visible={currentStep === 0} {...basicsStepProps} />

              {/* v24.12: Organizer-Einstellungen als eigener Wizard-Schritt (Schritt 2). */}
              <DetailsStep visible={currentStep === 1} {...detailsStepProps} />

              {/* ===== Step 3 (v15.0: vormals Step 2): Ort & Programm ===== */}
              <LocationProgramStep visible={currentStep === 2} {...locationProgramStepProps} />

              {/* ===== Step 2 (v15.0, vormals Step 3): Sub-Events =====
                  Sub-Events (Workshops / Sessions / Programmpunkte), plus
                  Bezeichnungs-Dropdown und Anmelde-Modus (Hauptevent +
                  Sub-Events vs. nur Sub-Events).
                  v15.0: vorgezogen vor „Ort & Programm", damit die folgenden
                  Steps pro-Sub-Event-Tabs anbieten können. */}
              {/* v28.87: Frueher Schritt 3 („Sub-Events"). Der Schritt ist
                  entfallen; sein Inhalt hängt jetzt unten an Schritt 1
                  (Grundlagen). Der Block bleibt als Ganzes bestehen — nur
                  seine Anzeige-Bedingung zeigt auf Schritt 1.
                  v28.89: …und nur auf der Klammer-/Hauptevent-Ebene. Die Liste
                  ist die Übersicht ÜBER die Sub-Events; auf dem Reiter eines
                  einzelnen Sub-Events stünde sie unter dessen eigenen
                  Grundlagen und läse sich wie eine Verschachtelung. */}
              <SubEventsSection visible={currentStep === 0 && activeScopeIdx === 0} {...subEventsSectionProps} />

              {/* ===== Step 4 (v14.8: vormals Step 3): Kapazität, Fristen & Sichtbarkeit ===== */}
              <CapacityStep visible={currentStep === 3} {...capacityStepProps} />

              {/* ===== Step 5 (v14.8: vormals Step 4): Team-Anmeldung =====
                  Renderblock für den Wizard-Schritt Team-Anmeldung.
                  Konfiguriert Team-Anmeldung-Toggle + Teamgröße +
                  Team-Name-Frage. v15: Index 4 → 6 (Team kommt jetzt nach
                  Kommunikation). */}
              <TeamStep
                visible={currentStep === 6}
                teamRegistrationEnabled={teamRegistrationEnabled}
                setTeamRegistrationEnabled={setTeamRegistrationEnabled}
                teamSize={teamSize}
                setTeamSize={setTeamSize}
                askTeamName={askTeamName}
                setAskTeamName={setAskTeamName}
                teamTermSingular={teamTermSingular}
                setTeamTermSingular={setTeamTermSingular}
                teamTermPlural={teamTermPlural}
                setTeamTermPlural={setTeamTermPlural}
                teamMembersCannotCreate={teamMembersCannotCreate}
                setTeamMembersCannotCreate={setTeamMembersCannotCreate}
                teamPartialAllowed={teamPartialAllowed}
                setTeamPartialAllowed={setTeamPartialAllowed}
                teamOpenSlotsVisible={teamOpenSlotsVisible}
                setTeamOpenSlotsVisible={setTeamOpenSlotsVisible}
                teamJoinRequiresApproval={teamJoinRequiresApproval}
                setTeamJoinRequiresApproval={setTeamJoinRequiresApproval}
              />

              {/* ===== Step 5 (v15: vormals Step 6): Registrierungsfelder ===== */}
              <FieldsStep visible={currentStep === 4} {...fieldsStepProps} />

              {/* ===== Step 6 (v15: vormals Step 7): Kommunikation ===== */}
              <CommunicationStep visible={currentStep === 5} {...communicationStepProps} />

              {/* ===== Step 8 (v14.8: vormals Step 7): Dokumente ===== */}
              <DocumentsStep
                visible={currentStep === 7}
                documents={documents}
                setDocuments={setDocuments}
                allowAttendeeUpload={allowAttendeeUpload}
                setAllowAttendeeUpload={setAllowAttendeeUpload}
                attendeeUploadLabel={attendeeUploadLabel}
                setAttendeeUploadLabel={setAttendeeUploadLabel}
                attendeeUploadHint={attendeeUploadHint}
                setAttendeeUploadHint={setAttendeeUploadHint}
                renderStepIntro={renderStepIntro}
              />

              {/* ===== Step 9 (v14.8: vormals Step 8): Fun-Zone ===== */}
              <FunZoneStep
                visible={currentStep === 8}
                quiz={quiz}
                addQuizQuestion={addQuizQuestion}
                removeQuizQuestion={removeQuizQuestion}
                updateQuizQuestion={updateQuizQuestion}
                pendingSections={pendingSections}
                setPendingSections={setPendingSections}
                draggedQuestionId={draggedQuestionId}
                setDraggedQuestionId={setDraggedQuestionId}
                setNewSectionName={setNewSectionName}
                setNewSectionError={setNewSectionError}
                setNewSectionModalOpen={setNewSectionModalOpen}
                renderStepIntro={renderStepIntro}
              />

              {/* v29.66: F&A-Pilot — Schritt 10 „Abrechnung" (nur Admins). Haengt
                  als LETZTER Schritt an, damit kein bestehender Index wandert.
                  Nicht abrechnungsrelevant: nur die Frage; „Ja" blendet die
                  Abschnitte sofort ein (reiner State, kein Neuladen).
                  v30.38: Stand bis hierher AUSSERHALB von `.card`/`.creation-form`
                  — der grüne Schritt-Kopf lebt von `margin: -32px` gegen die
                  Karten-Polsterung, und ohne Karte bleute er gegen die Seite.
                  Sichtbar war das als leerer weißer Streifen über dem Kopf: die
                  Karte mit den Schritten 1–9 (alle `display:none`) rendert dann
                  nur noch ihre eigene Polsterung. */}
              {canBilling && (
                <BillingStep
                  visible={currentStep === 9}
                  billingRelevant={billingRelevant}
                  setBillingRelevant={setBillingRelevant}
                  billingSendMode={billingSendMode}
                  setBillingSendMode={setBillingSendMode}
                  billingFields={billingFields}
                  setBillingFields={setBillingFields}
                />
              )}

            </div>{/* close creation-form */}
          </div>{/* close card */}

          {/* v29.66: F&A-Pilot — Frage-Dialog nach den Nutzungsbedingungen.
              Keine Vorauswahl: Das Konzept verlangt eine AKTIVE Entscheidung,
              deshalb zwei gleichrangige Knoepfe statt Radio mit Default. */}
          {canBilling && billingPromptOpen && (
            <div style={{
              position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1300,
              display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
            }}>
              <div className="card" style={{
                width: '100%', maxWidth: 560, padding: 28, borderRadius: 16,
                background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
              }}>
                {/* v29.69: Pilot-Badge — der Dialog erscheint nur Admins;
                    das soll man ihm ansehen, damit im Test niemand glaubt,
                    Organizer bekaemen diese Frage bereits. */}
                <span style={{
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                  background: 'rgba(237,139,0,0.14)', color: 'var(--dex-orange-dark, #b96a00)',
                  border: '1px solid var(--dex-orange, #ed8b00)',
                  borderRadius: 999, padding: '3px 12px', marginBottom: 10,
                  fontSize: '0.72rem', fontWeight: 700, letterSpacing: 0.4, textTransform: 'uppercase',
                }}>
                  Pilot — aktuell nur für Admins sichtbar
                </span>
                <h2 style={{ margin: '0 0 10px', fontSize: '1.15rem' }}>
                  Handelt es sich um ein abrechnungsrelevantes Event?
                </h2>
                <p style={{ fontSize: '0.9rem', color: 'var(--dex-gray-600)', margin: '0 0 18px' }}>
                  Abrechnungsrelevante Events sind Veranstaltungen, deren Kosten oder
                  Bewirtungsaufwendungen gegenüber Finance &amp; Accounting dokumentiert
                  oder abgerechnet werden müssen. Das ist der Fall, wenn im Nachgang
                  <strong> Rechnungen über die Kreditorenbuchhaltung eingereicht
                  werden</strong> — etwa für Catering, eine externe Raumbuchung oder
                  Anmeldegebühren (z.B. Startgelder für Läufer) — oder wenn für das
                  Event <strong>Ariba-Bestellungen</strong> ausgelöst werden.
                </p>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', margin: '0 0 18px' }}>
                  Die Entscheidung lässt sich später jederzeit im Schritt „Abrechnung&ldquo; ändern.
                </p>
                <div style={{ display: 'flex', gap: 12, justifyContent: 'flex-end' }}>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    onClick={() => { setBillingRelevant(false); setBillingPromptOpen(false); }}
                  >
                    Nein
                  </button>
                  <button
                    type="button"
                    className="btn btn-primary"
                    onClick={() => { setBillingRelevant(true); setBillingPromptOpen(false); }}
                  >
                    Ja, abrechnungsrelevant
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Fortschrittsanzeige */}
          {isSubmitting && (
            <div className="mt-24" style={{ padding: '20px 0' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
                  {progressLabel}
                </span>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--dex-green)' }}>
                  {progress}%
                </span>
              </div>
              <div style={{
                width: '100%', height: 8, background: 'var(--dex-gray-200)',
                borderRadius: 4, overflow: 'hidden',
              }}>
                {/* v29.57: Derselbe Schimmer wie im Boot-Balken (v29.41). Beim
                    Speichern eines Events mit vielen Terminen steht der Balken
                    zwischen zwei Abschnitten sekundenlang fast still — ohne
                    Lebenszeichen liest sich das als Hänger, und Organizer
                    klicken dann ein zweites Mal auf Speichern. Bewusst schwach
                    (weiß auf Grün, 45 %) und langsam, kein Blinken. */}
                <div style={{
                  width: `${progress}%`, height: '100%',
                  background: progress === 100
                    ? 'var(--dex-green)'
                    : 'linear-gradient(90deg, var(--dex-green), #0076a8)',
                  borderRadius: 4,
                  transition: 'width 0.5s ease',
                  position: 'relative',
                  overflow: 'hidden',
                }}>
                  {progress < 100 && (
                    <div style={{
                      position: 'absolute', top: 0, bottom: 0, left: 0, width: '35%',
                      background: 'linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.45) 50%, rgba(255,255,255,0) 100%)',
                      animation: 'dexProgressShimmer 2.1s ease-in-out infinite',
                    }} />
                  )}
                </div>
              </div>
            </div>
          )}

          {!isSubmitting && (
            <>
            {/* v22.36: Aufgeräumte Aktions-Leiste — klare Hierarchie statt
                drei loser Gruppen: links die Rück-Navigation, rechts erst die
                ruhigen Werkzeuge (Vorschau / Prüfen), dann durch einen
                Trenner abgesetzt die Haupt-Aktionen (Speichern als Outline,
                Weiter/Anlegen als EINZIGER gefüllter Primär-Button).
                v22.22: ref für den schwebenden Weiter-Button. */}
            <div ref={actionRowRef} style={{
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              gap: 12, marginTop: 28, paddingTop: 16,
              borderTop: '1px solid var(--dex-gray-200)', flexWrap: 'wrap',
            }}>
              <div style={{ display: 'flex', gap: 8 }}>
                {currentStep === 0 ? (
                  <button className="btn btn-danger" onClick={() => goBack()}><Trash2 size={16} /> {t('create.cancel')}</button>
                ) : (
                  <button className="btn btn-secondary" onClick={() => setCurrentStep(currentStep - 1)}>
                    {t('general.back')}
                  </button>
                )}
              </div>

              <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setShowRegisterPreview(true)}
                  disabled={!title}
                  style={{ opacity: title ? 1 : 0.5 }}
                  title={title ? 'So sehen Teilnehmer die Registrierungsseite' : 'Event-Titel eingeben, um die Vorschau zu öffnen'}
                >
                  {t('create.registerpreview')}
                </button>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setShowConfigCheck(true)}
                  disabled={!title}
                  style={{ opacity: title ? 1 : 0.5 }}
                  title={isDe ? 'Alle Einstellungen des Events im Überblick prüfen' : 'Review all event settings at a glance'}
                >
                  {isDe ? 'Prüfen' : 'Review'}
                </button>

                {/* Trenner zwischen Werkzeugen und Haupt-Aktionen */}
                <span aria-hidden="true" style={{ width: 1, height: 26, background: 'var(--dex-gray-200)', margin: '0 4px' }} />

                {/* v17.5: Im Edit-Modus immer speichern können, ohne durch
                    alle Steps zu klicken. v22.36: als Outline-Variante —
                    der gefüllte Primär-Button bleibt Weiter/Anlegen. */}
                {isEditMode && currentStep < steps.length - 1 && (
                  <button
                    className="btn btn-outline"
                    disabled={!title}
                    onClick={attemptSubmitGuarded}
                    style={{ opacity: !title ? 0.5 : 1 }}
                  >
                    <Send size={16} /> {isDe ? 'Speichern & zurück zum Event' : 'Save & return to event'}
                  </button>
                )}

                {currentStep < steps.length - 1 ? (
                  <button
                    className="btn btn-primary"
                    onClick={proceedNext}
                  >
                    {t('create.next')}
                  </button>
                ) : (
                  <button
                    className="btn btn-primary"
                    data-tour="wizard-submit"
                    disabled={!title}
                    onClick={attemptSubmitGuarded}
                    style={{ opacity: !title ? 0.5 : 1 }}
                  >
                    <Send size={16} /> {isEditMode ? t('create.save') : t('create.submit')}
                  </button>
                )}
              </div>
            </div>

            {/* v22.22: Schwebender Weiter-Button — sichtbar nur, solange die
                Aktions-Zeile unten noch nicht im Viewport ist. Blendet beim
                Erreichen des Seitenendes weich aus (der echte Button übernimmt).
                v22.26: Position + Look exakt wie die Jump-Buttons im Organizer
                Center (unten MITTIG, grüne Pille) — bewährtes fixed-Muster,
                kollidiert nicht mit dem Chat-Icon unten rechts. */}
            {currentStep < steps.length - 1 && (
              <button
                type="button"
                aria-hidden={actionRowVisible}
                tabIndex={actionRowVisible ? -1 : 0}
                onClick={proceedNext}
                style={{
                  position: 'fixed', left: '50%', bottom: 20, zIndex: 900,
                  background: 'var(--dex-green, #86bc25)', color: '#fff',
                  border: 'none', padding: '10px 16px', borderRadius: 999,
                  cursor: 'pointer', fontSize: '0.85rem', fontWeight: 600,
                  boxShadow: '0 4px 12px rgba(0,0,0,0.18)',
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                  fontFamily: 'inherit',
                  transition: 'opacity 0.3s ease, transform 0.3s ease',
                  opacity: actionRowVisible ? 0 : 1,
                  transform: actionRowVisible ? 'translate(-50%, 14px)' : 'translate(-50%, 0)',
                  pointerEvents: actionRowVisible ? 'none' : 'auto',
                }}
              >
                {t('create.next')}
              </button>
            )}
            </>
          )}
        </div>
      </div>

      <WizardModals {...wizardModalsProps} />
    </div>
  );
}
