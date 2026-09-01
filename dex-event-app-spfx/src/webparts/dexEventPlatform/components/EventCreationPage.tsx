/**
 * Event-Erstellung (nur für Organizer / SuperAdmin)
 *
 * Erstellt ein Event in der DEX_Events-Liste und eine
 * separate Teilnehmerliste mit Item-Level Permissions.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
// v20.4: moderne Confirm-/Alert-Modals statt window.confirm/alert.
import { useDialog } from '../context/DialogContext';
import { EventService } from '../services/EventService';
// v26.48: zentrale B2Run-Köln-Vorlage (Titel-Erkennung + 7 Meldefelder mit
// deterministischen IDs für den offiziellen Excel-Export).
import { getCachedOrbBase64 } from '../services/EmailTemplates';
import { EventType, AgendaItem } from '../types';
import { InfoTooltip } from './InfoTooltip';
import WizardHint from './WizardHint';
// v20.2: Self-Check-in ist aus dem Wizard ausgezogen — Aktivierung läuft
// automatisch beim ersten Klick auf die Aktionen (Check-in-Seite, Admin
// Center, QR-Kachel im Event-Detail); Zeitfenster + Deaktivieren im
// Kachel-Modal des Admin Centers.
import { buildOutlookLocation } from '../utils/eventFormat';
import { setActiveWizardStep } from '../utils/wizardStepContext';
import { canEditBilling } from '../data/billingFields';
// v28.98: Sperrt „Zurück" im Header, solange gespeichert wird.
import { setSaveInProgress } from '../utils/saveGuard';
// v28.94: Unterkomponenten des Assistenten liegen jetzt in ./wizard —
// sie kennen den Wizard-State nicht und liessen sich deshalb ohne
// Verhaltensaenderung herausloesen.
import { useIsMobile } from '../utils/useIsMobile';
import { registerLocale } from 'react-datepicker';
import { de } from 'date-fns/locale';
import 'react-datepicker/dist/react-datepicker.css';
import { SubEventDraft, OutlookConfirmItem } from './wizard/wizardTypes';
import { EmailOverrideEntry } from './wizard/emailOverrideEntry';
import { compressImage } from '../utils/imageCompress';
import { readOutlookLogo } from './wizard/wizardHelpers';
import { detectOutlookRelevantChangesImpl } from './wizard/logic/outlookChanges';
import { runWizardSubmit } from './wizard/logic/wizardSubmit';
import { persistSubEventsForParentImpl } from './wizard/logic/persistSubEvents';
import { WizardTermsModal } from './wizard/WizardTermsModal';
import { WizardModals } from './wizard/WizardModals';
import { SUB_TRANSFER_GROUPS } from '../data/wizardHints';
import { renderGlobalScopeBarImpl, renderKlammerVisibilityMismatchImpl, renderOutlookUpdateButtonImpl, renderPerEventTabStripImpl, renderPreviewSectionImpl, renderVisibilitySummaryBoxImpl } from './wizard/logic/wizardRenderHelpers';
import { applySubTransferImpl, getStepErrorsForImpl, toggleDaySubEventImpl } from './wizard/logic/wizardMisc';
import { applyCommToAllSubEventsImpl, flushActiveCommTabToStateImpl, resolveTopLevelCommStateImpl, switchCommTabImpl } from './wizard/logic/commTabs';
import { applyDraftPayloadImpl } from './wizard/logic/wizardDraft';
import { confirmOutlookSaveImpl, createMissingOutlookAppointmentsImpl, triggerOutlookUpdateAllImpl, triggerOutlookUpdateNowImpl } from './wizard/logic/outlookActions';
import { loadDemoSubEventImpl, loadDemoSubEventTeamImpl } from './wizard/logic/wizardTemplates';
import { WizardFormShell } from './wizard/WizardFormShell';
import { useWizardEventFieldState } from './wizard/hooks/useWizardEventFieldState';
import { useWizardVisibilityState } from './wizard/hooks/useWizardVisibilityState';
import { useWizardOptionState } from './wizard/hooks/useWizardOptionState';

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
  // v30.66: `activeFrom` war das einzige Datumsfeld, das die Berlin-Konvertierung
  // umging — roh aus SharePoint geladen (UTC-ISO mit Z), aber vom DatePicker als
  // naive Lokalzeit ueberschrieben und mit `new Date(...)` gespeichert. Damit hing
  // der Freischalt-Zeitpunkt an der Zeitzone des Browsers: In einer VDI-Sitzung auf
  // UTC ging ein auf 00:00 gesetztes Event erst um 02:00 Berliner Zeit auf. Jetzt
  // derselbe Pfad wie bei allen anderen Fristen (isoToLocal beim Laden,
  // berlinLocalToUtcIso beim Speichern) — so wie es v29.19 fuer
  // NotifyOrgRegisterFromDate schon gemacht hat.
  const [activeFrom, setActiveFrom] = React.useState(editEvent && editEvent.activeFrom ? isoToLocal(editEvent.activeFrom) : '');
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
  const {
    allDay, audience, autoDeregisterOnDecline, bundledComm, childGender, childTermPlural,
    childTermSingular, customFields, customTermMode, description, disableCancellationEmail, disableEmails,
    disableOutlook, disableRegistrationEmail, emailLanguage, emailLogoFromPhoto, endDate, eventImageUrl,
    excludedUsers, filterMode, htmlEditorMode, htmlEditorOpen, htmlEditorTemplateType, imageBanner,
    imageDisplay, imageDisplayOpen, imageEditOpen, imageFile, imageOrigAspect, imageOrigFile,
    imagePreview, inactiveHandling, initialOrgGetsSubInvitesRef, klammerDeadline, lastDeregisterDate, locationFilter,
    logoCropTarget, mainCommDisabledAck, maxParticipants, noCancelAfterDeadline, noDescription, notifyOrgCancelMode,
    notifyOrgRegisterFromDate, notifyOrgRegisterMode, onlineMeetingMode, orgGetsSubInvites, orgInvitesTouchedRef, outlookBody,
    outlookEndOverride, outlookHeading, outlookLocationOverride, outlookLogoFromPhoto, outlookStartOverride, outlookSubheading,
    outlookSubject, pendingSuccessDispatch, pendingSuccessDispatchRef, registrationDeadline, removedSavedSubs, requireSubEventSelection,
    setAllDay, setAudience, setAutoDeregisterOnDecline, setBundledComm, setChildGender, setChildTermPlural,
    setChildTermSingular, setCustomFields, setCustomTermMode, setDescription, setDisableCancellationEmail, setDisableEmails,
    setDisableOutlook, setDisableRegistrationEmail, setEmailLanguage, setEmailLogoFromPhoto, setEndDate, setEventImageUrl,
    setExcludedUsers, setFilterMode, setHtmlEditorMode, setHtmlEditorOpen, setHtmlEditorTemplateType, setImageBanner,
    setImageDisplay, setImageDisplayOpen, setImageEditOpen, setImageFile, setImageOrigAspect, setImageOrigFile,
    setImagePreview, setInactiveHandling, setKlammerDeadline, setLastDeregisterDate, setLocationFilter, setLogoCropTarget,
    setMainCommDisabledAck, setMaxParticipants, setNoCancelAfterDeadline, setNoDescription, setNotifyOrgCancelMode, setNotifyOrgRegisterFromDate,
    setNotifyOrgRegisterMode, setOnlineMeetingMode, setOrgGetsSubInvites, setOutlookBody, setOutlookEndOverride, setOutlookHeading,
    setOutlookLocationOverride, setOutlookLogoFromPhoto, setOutlookStartOverride, setOutlookSubheading, setOutlookSubject, setPendingSuccessDispatch,
    setRegistrationDeadline, setRemovedSavedSubs, setRequireSubEventSelection, setShowAsFree, setShowDemoVariantModal, setShowSummaryModal,
    setShowTemplatePicker, setStartDate, setSubEventCalendar, setSubEventSingleChoice, setSubEventsOnlyMode, setSubImageCropIdx,
    setTeamsLink, setTemplateLoadingId, setTerminListOpen, setUnlimitedParticipants, setUserCancelAllowed, setWaitlistEnabled,
    showAsFree, showDemoVariantModal, showSummaryModal, showTemplatePicker, startDate, storedEventType,
    subEventCalendar, subEventSingleChoice, subEventsOnlyMode, subImageCropIdx, teamsLink, templateLoadingId,
    terminListOpen, unlimitedParticipants, userCancelAllowed, waitlistEnabled, wizardImgAspect,
  } = useWizardEventFieldState({ editEvent, ensureEventDocuments, isoToLocal, locale });
  const {
    agenda, applyEventPhotoToLogo, assistantsCanSee, documents, dragFieldId, dragOverFieldId,
    effTeamsLink, emailLogoPreview, emailTemplateOverrides, emailTemplates, fieldExpandOverride, fileToBase64,
    headerImageLayout, headerImageLayoutConfig, headerLayoutFor, hiddenOrganizerEmails, hideOrganizer, hideOrganizerIndividualOnly,
    initialDocumentNames, initialHeaderImageLayoutRef, isFictive, newSectionError, newSectionModalOpen, newSectionName,
    onlineMeetingChanged, organizerDisplayLarge, outlookLogoPreview, outlookTeamsLink, quiz, quizClusterSize,
    renderHeaderSizeControl, reorderMode, setAgenda, setAssistantsCanSee, setDocuments, setDragFieldId,
    setDragOverFieldId, setEmailLogoPreview, setEmailTemplateOverrides, setEmailTemplates, setFieldExpandOverride, setHeaderImageLayout,
    setHideOrganizer, setHideOrganizerIndividualOnly, setIsFictive, setNewSectionError, setNewSectionModalOpen, setNewSectionName,
    setOrganizerDisplayLarge, setOutlookLogoPreview, setQuiz, setReorderMode, setSubEvents, setTransferTimes,
    setTriggerOutlookUpdate, shrinkLogoB64, subEvents, toggleFieldExpand, toggleOrganizerHidden, transferTimes,
    triggerOutlookUpdate,
  } = useWizardVisibilityState({ childEventsOf, editEvent, imageFile, imageOrigFile, imagePreview, isDe, locale, onlineMeetingMode, showAlert, teamsLink });
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
    return await createMissingOutlookAppointmentsImpl({
      childTermPlural, childTermSingular, confirmDialog, editEvent, flushActiveCommTabToState, forceOutlookRecreateRef,
      handleSubmit, isDe, outlookMissingTargets, parentTimesIso, pendingOutlookDirtyWriteRef, pendingOutlookDirtyWriteRefs,
      pendingOutlookRecreateForSubEventsRef, pendingOutlookUpdateForSubEventsRef, pendingOutlookUpdateForTopRef, setOutlookUpdateBusy, showAlert, subEventsRef,
    });
  };
  const triggerOutlookUpdateNow = async (): Promise<void> => {
    return await triggerOutlookUpdateNowImpl({
      activeCommTabIdx, confirmDialog, disableOutlook, editEvent, isDe, setOutlookUpdateBusy,
      setOutlookUpdateDone, showAlert, subEvents, title,
    });
  };
  /** v28.28: Haupt-Termin UND alle Sub-Event-Termine in einem Rutsch. */
  const triggerOutlookUpdateAll = async (): Promise<void> => {
    return await triggerOutlookUpdateAllImpl({
      confirmDialog, isDe, outlookUpdateTargets, setOutlookUpdateBusy, setOutlookUpdateDone, showAlert,
    });
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
    return renderOutlookUpdateButtonImpl({
      activeCommTabIdx, childTermPlural, childTermSingular, createMissingOutlookAppointments, editEvent, isDe,
      outlookMissingTargets, outlookUpdateBusy, outlookUpdateDone, outlookUpdateTargets, subEvents, title,
      triggerOutlookUpdateAll, triggerOutlookUpdateNow,
    });
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
  const {
    SUGGESTED_FIELDS_CATALOG, addCustomField, addSelectedSuggestedFields, addStartblock, addSubEventCustomField, allowAttendeeUpload,
    applyEventTemplate, askSalutation, askTeamName, attendeeUploadHint, attendeeUploadLabel, beforeNextSaturday,
    bilingualFields, confirmDialogEnabled, confirmDialogMode, confirmDialogText, copyParentFieldsToSubEvent, dragOverSectionId,
    dragSectionId, durchstarterRequiresProof, durchstarterStartblock, error, fmtDate, fmtDatetime,
    funstarterStartblock, imageUploadError, isB2runTemplate, isSubmitting, isVisOpen, loadDemoGroups,
    loadDemoStandard, locationOptions, mainEventLabel, mainEventLabelMode, nextSaturdayAt, openSuggestedModal,
    previewSections, progress, progressLabel, registrationLanguage, removeCustomField, removeStartblock,
    removeSubEventCustomField, renderShowIfConfig, resetDemoVariantBaseState, setAllowAttendeeUpload, setAskSalutation, setAskTeamName,
    setAttendeeUploadHint, setAttendeeUploadLabel, setBilingualFields, setConfirmDialogEnabled, setConfirmDialogMode, setConfirmDialogText,
    setDragOverSectionId, setDragSectionId, setDurchstarterStartblock, setError, setFunstarterStartblock, setImageUploadError,
    setIsSubmitting, setMainEventLabel, setMainEventLabelMode, setPreviewSections, setProgress, setProgressLabel,
    setRegistrationLanguage, setShowB2runSuggested, setShowPreview, setShowRegisterPreview, setShowSuggestedModal, setSubmitted,
    setSuggestedSelection, setTeamJoinRequiresApproval, setTeamMembersCannotCreate, setTeamOpenSlotsVisible, setTeamPartialAllowed, setTeamRegistrationEnabled,
    setTeamSize, setTeamTermPlural, setTeamTermSingular, setTriedNext, setUseSplitCapacities, setVisCopyModalOpen,
    showB2runSuggested, showPreview, showRegisterPreview, showSuggestedModal, submitted, suggestedSelection,
    teamJoinRequiresApproval, teamMembersCannotCreate, teamOpenSlotsVisible, teamPartialAllowed, teamRegistrationEnabled, teamSize,
    teamTermPlural, teamTermSingular, triedNext, updateCustomField, updateSubEventCustomField, useSplitCapacities,
    visCopyModalOpen, visCopyPendingRef, visHeader, visSnapshotRef,
  } = useWizardOptionState({ b2runStartblocks, customFields, durchstarterCapacity, editEvent, funstarterCapacity, isDe, isEditMode, newStartblock, selectedTemplate, setAddrCity, setAddrHouseNo, setAddrStreet, setAddrZip, setAgenda, setAudience, setB2runStartblocks, setContactEmail, setContactInfo, setContactName, setCurrentStep, setCustomFields, setDescription, setDurchstarterCapacity, setEmailLanguage, setEndDate, setEventImageUrl, setExcludedUsers, setFieldExpandOverride, setFilterMode, setFunstarterCapacity, setImageFile, setImageOrigAspect, setImageOrigFile, setImagePreview, setKlammerDeadline, setLastDeregisterDate, setLocation, setLocationFilter, setMaxParticipants, setNewStartblock, setNoDescription, setRegistrationDeadline, setRemovedSavedSubs, setSelectedTemplate, setShowTemplatePicker, setSplitDescA, setSplitDescB, setSplitHelpText, setSplitLabelA, setSplitLabelB, setSplitSectionTitle, setSplitSharedWaitlist, setStartDate, setSubEvents, setTemplateLoadingId, setTitle, setTransferTimes, setUnlimitedParticipants, setWaitlistEnabled, showAlert });
  const loadDemoSubEvent = (): void => {
    return loadDemoSubEventImpl({
      beforeNextSaturday, berlinLocalToUtcIso, fmtDate, fmtDatetime, nextSaturdayAt, resetDemoVariantBaseState,
      setAskSalutation, setCurrentStep, setCustomFields, setDescription, setEndDate, setLastDeregisterDate,
      setLocation, setMaxParticipants, setRegistrationDeadline, setStartDate, setSubEvents, setTitle,
      setUseSplitCapacities, setWaitlistEnabled,
    });
  };

  const loadDemoSubEventTeam = (): void => {
    return loadDemoSubEventTeamImpl({
      beforeNextSaturday, berlinLocalToUtcIso, fmtDate, fmtDatetime, nextSaturdayAt, resetDemoVariantBaseState,
      setAskSalutation, setAskTeamName, setCurrentStep, setCustomFields, setDescription, setEndDate,
      setLastDeregisterDate, setLocation, setMaxParticipants, setRegistrationDeadline, setStartDate, setSubEvents,
      setTeamJoinRequiresApproval, setTeamOpenSlotsVisible, setTeamPartialAllowed, setTeamRegistrationEnabled, setTeamSize, setTitle,
      setUseSplitCapacities, setWaitlistEnabled,
    });
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
    return switchCommTabImpl({
      activeCommTabIdx, autoDeregisterOnDecline, disableCancellationEmail, disableEmails, disableOutlook, disableRegistrationEmail,
      emailLanguage, emailLogoPreview, emailTemplateOverrides, inactiveHandling, locale, outlookBody,
      outlookHeading, outlookLogoPreview, outlookSubheading, outlookSubject, setActiveCommTabIdx, setAutoDeregisterOnDecline,
      setDisableCancellationEmail, setDisableEmails, setDisableOutlook, setDisableRegistrationEmail, setEmailLanguage, setEmailLogoPreview,
      setEmailTemplateOverrides, setInactiveHandling, setOutlookBody, setOutlookHeading, setOutlookLogoPreview, setOutlookSubheading,
      setOutlookSubject, setSubEvents, subEvents, subEventsRef, topLevelCommSnapshot,
    }, nextIdx);
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
    return flushActiveCommTabToStateImpl({
      activeCommTabIdx, autoDeregisterOnDecline, disableCancellationEmail, disableEmails, disableOutlook, disableRegistrationEmail,
      emailLanguage, emailLogoPreview, emailTemplateOverrides, inactiveHandling, outlookBody, outlookHeading,
      outlookLogoPreview, outlookSubheading, outlookSubject, setSubEvents, subEventsRef,
    });
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
  const resolveTopLevelCommState = (): { emailLanguage: string; emailLogoBase64: string; outlookLogoBase64: string; outlookBody: string; outlookHeading: string; outlookSubheading: string; outlookSubject: string; disableEmails: boolean; disableRegistrationEmail: boolean; disableCancellationEmail: boolean; autoDeregisterOnDecline: boolean; inactiveHandling?: 'notify' | 'autoderegister'; disableOutlook: boolean; emailTemplateOverrides: Record<string, EmailOverrideEntry>; } => {
    return resolveTopLevelCommStateImpl({
      activeCommTabIdx, autoDeregisterOnDecline, disableCancellationEmail, disableEmails, disableOutlook, disableRegistrationEmail,
      emailLanguage, emailLogoPreview, emailTemplateOverrides, inactiveHandling, outlookBody, outlookHeading,
      outlookLogoPreview, outlookSubheading, outlookSubject, topLevelCommSnapshot,
    });
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
    return confirmOutlookSaveImpl({
      editEvent, handleSubmit, outlookConfirmChecks, outlookConfirmItems, pendingOutlookDirtyWriteRef, pendingOutlookDirtyWriteRefs,
      pendingOutlookRecreateForSubEventsRef, pendingOutlookUpdateForSubEventsRef, pendingOutlookUpdateForTopRef, setOutlookConfirmOpen, setTriggerOutlookUpdate,
    });
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
    return renderGlobalScopeBarImpl({
      activeScopeIdx, childTermPlural, childTermSingular, currentStep, isDe, renderPerEventTabStrip,
      SCOPE_AWARE_STEPS, setScope, subEvents, subEventsOnlyMode, title,
    });
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
    return renderVisibilitySummaryBoxImpl({
      isDe,
    }, locList, audienceStr, mode, excludedCount);
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
    return renderKlammerVisibilityMismatchImpl({
      audience, isDe, locationFilter, setLocationFilter, subEvents, subEventsOnlyMode,
    });
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
    return applyDraftPayloadImpl({
      canBilling, setActiveFrom, setAddrCity, setAddrHouseNo, setAddrStreet, setAddrZip,
      setAgenda, setAskSalutation, setAskTeamName, setAudience, setBillingFields, setBillingRelevant,
      setBillingSendMode, setCancelRuleAfter, setCancelRuleAmount, setCancelRuleEnabled, setCancelRuleUnit, setContactEmail,
      setContactInfo, setContactName, setCurrentStep, setCustomFields, setDescription, setDisableEmails,
      setDisableOutlook, setEmailTemplateOverrides, setEndDate, setExcludedUsers, setFilterMode, setKlammerDeadline,
      setLastDeregisterDate, setLocation, setLocationFilter, setMaxParticipants, setNoCancelAfterDeadline, setOnlineMeetingMode,
      setOpenRuleDays, setOpenRuleEnabled, setOpenRuleFixedDate, setOpenRuleMode, setOrganizer, setOrganizerEmails,
      setRegistrationDeadline, setRegRuleAmount, setRegRuleEnabled, setRegRuleUnit, setRequireSubEventSelection, setStartDate,
      setSubEventCalendar, setSubEvents, setSubEventSingleChoice, setSubEventsOnlyMode, setSubEventsOptIn, setTeamRegistrationEnabled,
      setTeamSize, setTeamsLink, setTitle, setUserCancelAllowed, setVisAllSubs, setWaitlistEnabled,
    }, d);
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
    return renderPreviewSectionImpl({
      customFields, endDate, eventImageUrl, formatPreviewDate, startDate, title,
    }, sectionId);
  };


  // Hint-Bullets pro Step. Werden über das i-Icon in der Progress-Bar
  // (Mouseover) als Tooltip eingeblendet — vorher wurden sie als
  // dauerhafte grüne Hinweis-Box am Anfang jedes Steps angezeigt
  // (renderStepIntro), das war für geübte Organizer zu viel Rauschen.
  // v29.21 (Audit B5): Die EN-Liste stand noch auf der 10-Schritt-Zählung von
  // vor v28.87 (alter Grundlagen-Mix, eigener Sub-Events-Schritt) — ab dem
  // dritten Eintrag zeigte jedes i-Icon die Hints des FALSCHEN Schritts, und
  // für Schritt 9 gab es gar keinen Eintrag. Jetzt 1:1 parallel zu
  // STEP_HINTS_DE (9 Einträge).

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
    return getStepErrorsForImpl({
      allDay, endDate, lastDeregisterDate, maxParticipants, organizer, registrationDeadline,
      startDate, subEvents, subEventsOnlyMode, title, unlimitedParticipants, userCancelAllowed,
      useSplitCapacities,
    }, step);
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
    return applySubTransferImpl({
      asRec, childTermPlural, childTermSingular, flushActiveCommTabToState, isDe, setSubEvents,
      setSubTransfer, showAlert, subEvents, subEventsRef, subTransfer,
    });
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
    return await applyCommToAllSubEventsImpl({
      childTermPlural, confirmDialog, flushActiveCommTabToState, isDe, resolveTopLevelCommState, setSubEvents,
      showAlert, subEventsRef,
    });
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
    return toggleDaySubEventImpl({
      activeScopeIdx, allDay, berlinLocalToUtcIso, dayKeyOfDate, dayKeyOfSub, dayLabel,
      endDate, makeSubEventDraft, removedSavedSubs, setRemovedSavedSubs, setScope, setSubEvents,
      startDate, subEvents,
    }, d);
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
    return renderPerEventTabStripImpl({
      ariaLabel, childTermPlural, childTermSingular, goToSubEventsMode, isDe, mainLabel,
      subEventCalendar, subEvents, subEventsOnlyMode, title,
    }, activeIdx, onChange);
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
    currentUser, dayKeyOfDate, description, DRAFT_KEY, draftSavedAt,
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
    contactEmail, contactExpanded, contactInfo, contactName, contactOrganizerEmail,
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
    agenda, isDe, isMobile, isoToLocal, location,
    locationOptions, onlineMeetingMode, outlookLocationOverride, removeAgendaItem, renderStepIntro, setAddrCity,
    setAddrHouseNo, setAddrStreet, setAddrZip, setLocation, setOnlineMeetingMode, setOutlookLocationOverride,
    setSubEvents, setTeamsLink, setTransferTimes, startDate, subEvents, t,
    teamsLink, transferTimes, updateAgendaItem,
  };
  const subEventsSectionProps = {
    activeScopeIdx, audience, berlinLocalToUtcIso, childGender, childTermPlural, childTermSingular,
    confirmDialog, customTermMode, dayKeyOfSub, endDate, filterMode,
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
    durchstarterCapacity, durchstarterStartblock, effectiveKlammerDeadline, errorBorderStyle, excludedUsers,
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
    customFields, dragFieldId, dragOverFieldId, fieldExpandOverride, isDe,
    moveCustomField, newStartblock, openSuggestedModal, registrationLanguage, removeCustomField, removeStartblock,
    removeSubEventCustomField, renderShowIfConfig, renderStepIntro, reorderMode, setAskSalutation, setBilingualFields,
    setConfirmDialogEnabled, setConfirmDialogMode, setConfirmDialogText, setCustomFields, setDragFieldId, setDragOverFieldId,
    setNewStartblock, setRegistrationLanguage, setReorderMode, setSubEvents, splitLabelA, splitLabelB,
    subEvents, subEventsOnlyMode, t, title, toggleFieldExpand, updateCustomField,
    updateSubEventCustomField, useSplitCapacities,
  };
  const communicationStepProps = {
    activeCommTabIdx, applyCommToAllSubEvents, applyEventPhotoToLogo, autoDeregisterOnDecline, bundledComm, childTermPlural,
    commToggleRow, confirmDialog, disableCancellationEmail, disableEmails, disableOutlook,
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
  const wizardFormShellProps = {
    currentStep,
    actionRowRef, actionRowVisible, activeScopeIdx, addQuizQuestion, allowAttendeeUpload, askTeamName,
    attemptSubmitGuarded, attendeeUploadHint, attendeeUploadLabel, basicsStepProps, billingFields, billingPromptOpen,
    billingRelevant, billingSendMode, canBilling, capacityStepProps, communicationStepProps,
    detailsStepProps, documents, draftSavedAt, draggedQuestionId, error, fieldsStepProps,
    getStepErrorsFor, goBack, hintStepIdx, isDe, isEditMode, isSubmitting,
    locationProgramStepProps, pendingSections, proceedNext, progress, progressLabel, quiz,
    removeQuizQuestion, renderGlobalScopeBar, renderStepIntro, setAllowAttendeeUpload, setAskTeamName, setAttendeeUploadHint,
    setAttendeeUploadLabel, setBillingFields, setBillingPromptOpen, setBillingRelevant, setBillingSendMode, setCurrentStep,
    setDocuments, setDraggedQuestionId, setHintStepIdx, setNewSectionError, setNewSectionModalOpen, setNewSectionName,
    setPendingSections, setShowConfigCheck, setShowRegisterPreview, setTeamJoinRequiresApproval, setTeamMembersCannotCreate, setTeamOpenSlotsVisible,
    setTeamPartialAllowed, setTeamRegistrationEnabled, setTeamSize, setTeamTermPlural, setTeamTermSingular, setTriedNext,
    steps, subEventsSectionProps, t, teamJoinRequiresApproval, teamMembersCannotCreate, teamOpenSlotsVisible,
    teamPartialAllowed, teamRegistrationEnabled, teamSize, teamTermPlural, teamTermSingular, title,
    updateQuizQuestion,
  };
  return (
    <div ref={wizardRootRef} className="page-container" style={{ maxWidth: 1100, marginLeft: 'auto', marginRight: 'auto' }}>
      <WizardTermsModal {...wizardTermsModalProps} />
      <WizardFormShell {...wizardFormShellProps} />

      <WizardModals {...wizardModalsProps} />
    </div>
  );
}
