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

// Deutsche Locale registrieren
registerLocale('de', de);

// v22.36: Die kuratierten Agenda-Icon-Listen (AGENDA_ICONS/EXTENDED_ICONS)
// sind entfallen — Agenda-Schritte sind durchnummeriert statt bebildert.

/**
 * Komprimiert ein Bild clientseitig via Canvas.
 * Max 1200px Breite, JPEG 80% Qualität.
 */

/**
 * v28.32: Wenn Mail- und Outlook-Kopfbild identisch sind (der Normalfall, seit
 * v28.29 erst recht: Sub-Events erben das Bild des Hauptevents), stand dasselbe
 * Base64 ZWEIMAL im selben Datensatz — einmal als `_eventLogo`, einmal als
 * `_outlookLogo`. Statt der zweiten Kopie wird jetzt nur noch ein Marker
 * gespeichert; beim Laden wird daraus wieder das Mail-Bild. Spart bei einem
 * Foto-Kopfbild rund ein Drittel des Speicher-Payloads.
 */
const outlookLogoPiggyback = (emailLogo: string, outlookLogo: string): Record<string, unknown> => {
  if (!outlookLogo) return {};
  if (emailLogo && outlookLogo === emailLogo) return { _outlookLogoSameAsMail: true };
  return { _outlookLogo: outlookLogo };
};
/** Gegenstück zu `outlookLogoPiggyback` beim Laden. */
const readOutlookLogo = (ov: Record<string, unknown> | null | undefined): string => {
  if (!ov) return '';
  if (ov._outlookLogoSameAsMail) return (ov._eventLogo as string) || '';
  return (ov._outlookLogo as string) || '';
};



// v17.22: Einziger Serializer für Custom-Fields → CustomFields-JSON.
// Vorher dreimal copy-paste (Create-Save, Edit-Save, Sub-Event-Save), was
// dazu führte, dass der Sub-Event-Pfad die v17.20-EN-Varianten nicht
// mitnahm. Zentral hier, damit alle drei Pfade identisch persistieren.
//
// Wichtig (v17.22-Fix): DE-Optionen UND EN-Optionen werden POSITIONAL
// gepaart gefiltert — vorher wurde `options` per `.filter(Boolean)` von
// Leereinträgen befreit, `optionsEn` aber nicht, wodurch das Index-Mapping
// zwischen DE und EN bei leeren Slots verrutschte (leere/falsche EN-Labels
// auf der Anmeldeseite).
function serializeCustomFields(
  fields: CustomFieldInput[],
  bilingual: boolean
): CustomField[] {
  return fields
    .filter(f => f.label && f.label.trim().length > 0)
    .map(f => {
      let optionsOut: string[] | undefined;
      let optionsEnOut: string[] | undefined;
      // v26.75: Vorfilter-Kategorien POSITIONAL zu den (bereinigten) Optionen.
      let categoriesOut: string[] | undefined;
      if (f.type === 'select') {
        const pairs = (f.options || [])
          .map((o, i) => ({ de: (o || '').trim(), en: ((f.optionsEn || [])[i] || '').trim(), cat: (((f.optionCategories || [])[i]) || '').trim() }))
          .filter(p => p.de.length > 0);
        optionsOut = pairs.map(p => p.de);
        if (bilingual && pairs.some(p => p.en.length > 0)) {
          optionsEnOut = pairs.map(p => p.en);
        }
        if (!f.multi && pairs.some(p => p.cat.length > 0)) {
          categoriesOut = pairs.map(p => p.cat);
        }
      }
      return {
        id: f.id,
        label: f.label.trim(),
        type: f.type,
        required: !!f.required,
        visible: f.visible !== false,
        ...(f.helpText && f.helpText.trim() ? { helpText: f.helpText.trim() } : {}),
        // v18.18: nur persistieren wenn 'inline' (Default 'tooltip' = weglassen).
        ...(f.helpTextStyle === 'inline' ? { helpTextStyle: 'inline' as const } : {}),
        ...(f.showIf && f.showIf.fieldId && f.showIf.values && f.showIf.values.length > 0
          ? { showIf: { fieldId: f.showIf.fieldId, values: [...f.showIf.values] } }
          : {}),
        ...(optionsOut ? { options: optionsOut, ...(f.multi ? { multi: true } : {}) } : {}),
        // v26.74: Vorauswahl nur bei Single-Select persistieren, und nur wenn der
        // Wert eine der (bereinigten) Optionen ist.
        ...(f.type === 'select' && !f.multi && f.defaultValue && (optionsOut || []).indexOf(f.defaultValue) >= 0
          ? { defaultValue: f.defaultValue }
          : {}),
        // v26.75: Vorfilter-Kategorien + optionale Beschriftung (nur Single-Select).
        ...(categoriesOut
          ? { optionCategories: categoriesOut, ...(f.prefilterLabel && f.prefilterLabel.trim() ? { prefilterLabel: f.prefilterLabel.trim() } : {}) }
          : {}),
        // v24.25: Uhrzeit-Flag nur bei Datums-Feldern persistieren.
        ...(f.type === 'date' && f.withTime ? { withTime: true } : {}),
        // v28.63: Buchbares Fenster + Nächte-Limit nur beim Zeitraum-Feld.
        ...(f.type === 'daterange' ? {
          ...(f.rangeStart ? { rangeStart: f.rangeStart } : {}),
          ...(f.rangeEnd ? { rangeEnd: f.rangeEnd } : {}),
          ...(f.maxNights && f.maxNights > 0 ? { maxNights: f.maxNights } : {}),
        } : {}),
        ...(f.onlyForGroup && f.onlyForGroup !== 'all' ? { onlyForGroup: f.onlyForGroup } : {}),
        ...(f.type === 'checkbox' && f.confirmLabel && f.confirmLabel.trim()
          ? { confirmLabel: f.confirmLabel.trim() }
          : {}),
        // v17.20: Englische Varianten — nur wenn der Bilingual-Toggle an ist
        // UND der Organizer Text eingegeben hat.
        ...(bilingual && f.labelEn && f.labelEn.trim() ? { labelEn: f.labelEn.trim() } : {}),
        ...(bilingual && f.helpTextEn && f.helpTextEn.trim() ? { helpTextEn: f.helpTextEn.trim() } : {}),
        ...(bilingual && f.type === 'checkbox' && f.confirmLabelEn && f.confirmLabelEn.trim()
          ? { confirmLabelEn: f.confirmLabelEn.trim() }
          : {}),
        ...(optionsEnOut ? { optionsEn: optionsEnOut } : {}),
        ...(f.externalLinks && f.externalLinks.length > 0
          ? { externalLinks: f.externalLinks.map(x => ({ label: x.label, url: x.url })) }
          : {}),
        // v18.41: CC-bei-Mail nur für People-Picker-Felder persistieren.
        ...((f.type === 'user' || f.type === 'roommate') && f.ccOnEmails ? { ccOnEmails: true } : {}),
        // v26.60: Roommate-Benachrichtigung — nur das explizite ABSCHALTEN
        // persistieren (undefined = an, Bestandsverhalten bleibt unverändert).
        ...(f.type === 'roommate' && f.notifyRoommate === false ? { notifyRoommate: false } : {}),
        // v29.40: Verteiler-Begrenzung des Personen-Feldes mitschreiben.
        ...((f.type === 'user' || f.type === 'roommate') && f.audienceOnly ? { audienceOnly: true } : {}),
      } as CustomField;
    });
}

// v27.3: Der Outlook-Body wird beim Speichern mit fest aufgelöstem {{Organizer}}
// gespeichert. Kommen später Organizer dazu, blieb der alte Name eingebacken.
// Beim Edit-Laden mappen wir den eingebackenen Organizer-Namen wieder auf
// {{Organizer}} zurück — dann löst der nächste Save mit ALLEN aktuellen
// Organizern neu auf. Sicher: findet sich nichts, bleibt der Body unverändert.
function reinsertOrganizerPlaceholder(body: string, organizers: string[]): string {
  if (!body || !organizers || organizers.length === 0) return body;
  if (body.indexOf('{{Organizer}}') >= 0) return body; // schon Platzhalter
  const names = organizers.map(n => (n || '').trim()).filter(Boolean);
  // v29.21 (Audit): Der Save backt den Platzhalter über formatOrganizerList
  // — aus „Nachname, Vorname" (People-Picker-Format der Organizer-Spalte)
  // wird dort „Vorname Nachname". Genau diese Form steht also im
  // gespeicherten Body; die rohen Spalten-Namen matchen beim Regelfall mit
  // Komma NIE, und der v27.3-Mechanismus („Organizer-Wechsel löst beim
  // nächsten Save neu auf") war wirkungslos. Beide Formen probieren.
  const flipped = names.map(n => {
    const c = n.indexOf(',');
    return c > 0 ? `${n.slice(c + 1).trim()} ${n.slice(0, c).trim()}` : n;
  });
  const seenCand: Record<string, boolean> = {};
  const candidates = [...flipped, ...names].filter(n => (seenCand[n] ? false : (seenCand[n] = true)));
  for (const joiner of ['; ', ', ']) {
    const full = candidates.length > 1 ? flipped.join(joiner) : '';
    if (full && body.indexOf(full) >= 0) return body.split(full).join('{{Organizer}}');
  }
  const fullRaw = names.join('; ');
  if (fullRaw && body.indexOf(fullRaw) >= 0) return body.split(fullRaw).join('{{Organizer}}');
  // Bereits „kaputte"/veraltete Bodies: den ersten enthaltenen Organizer-Namen
  // (längster zuerst, um Teil-Treffer zu vermeiden) auf den Platzhalter mappen.
  for (const n of [...candidates].sort((a, b) => b.length - a.length)) {
    if (n.length >= 3 && body.indexOf(n) >= 0) return body.split(n).join('{{Organizer}}');
  }
  return body;
}







/**
 * v16.4: Audience-Liste (kommasepariert) in eine flache, ';'-separierte
 * Liste von Member-E-Mails auflösen. Jede '@'-Eintrag wird via
 * getGroupMembers (Graph) probiert — wenn die Auflösung eine
 * Mitglieder-Liste liefert, werden alle deren E-Mails übernommen, sonst
 * wird der Eintrag als direkte User-E-Mail behandelt. Lowercase + dedupliziert.
 *
 * Wird beim Event-Save aufgerufen und in die SP-Spalte
 * `AudienceResolvedEmails` geschrieben. matchesAudience im
 * EventListPage prüft Sichtbarkeit zur Laufzeit gegen diese Liste.
 */
async function resolveAudienceMembersToCsv(
  audienceCsv: string,
  getGroupMembers: (groupEmail: string) => Promise<{ groupName: string; members: Array<{ email: string }> } | null>,
): Promise<string> {
  const items = (audienceCsv || '').split(',').map(s => s.trim()).filter(Boolean);
  if (items.length === 0) return '';
  const out = new Set<string>();
  for (const item of items) {
    if (item.indexOf('@') < 0) continue; // Gruppen-Patterns (DEKOELN etc.) bleiben Runtime-Match
    try {
      const grp = await getGroupMembers(item);
      if (grp && grp.members && grp.members.length > 0) {
        for (const m of grp.members) {
          const e = (m.email || '').toLowerCase().trim();
          if (e) out.add(e);
        }
      } else {
        // Keine Member zurückgeliefert → behandle als direkte User-Adresse.
        out.add(item.toLowerCase());
      }
    } catch {
      out.add(item.toLowerCase());
    }
  }
  return Array.from(out).join(';');
}


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
    const keptDbIds = new Set<string>();
    const failedSubTitles: string[] = [];
    // v29.57: Einmal je Save auswerten, nicht je Sub-Event — die Schranke
    // hängt nur am Hauptevent. Bei einem NEUEN Event (kein Snapshot) wird nie
    // übersprungen; dort gibt es ohnehin nichts zu vergleichen.
    const subGateUnchanged = !!editEvent
      && subTopGateInitialRef.current !== ''
      && subTopGateInitialRef.current === subTopGateKey();
    let skippedSubCount = 0;
    // v29.74: Drossel-Abbruch (s. Schleifenrumpf).
    let consecutiveSubFailures = 0;
    let abortedForThrottle = false;
    const stepTotal = subEventsRef.current.filter(d => !!(d.title || '').trim()).length;
    let stepDone = 0;
    // v11.87: Sub-Event-Progress-Callback aus dem aufrufenden handleSubmit
    // einspeisen. Der Caller setzt window.__dexSubEventProgress vor dem
    // Aufruf und entfernt es danach. Wenn nicht gesetzt: no-op.
    const subOnProgress = (stage: string): void => {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const cb = (window as any).__dexSubEventProgress;
        if (typeof cb === 'function') cb(stage);
      } catch { /* */ }
    };
    // Sub-Events erben Organizer + OrganizerEmail vom Parent. Einmal sanitisieren
    // statt pro Iteration, identisch für alle Children.
    const sanitizedOrgPair = sanitizeOrganizerPairs();
    // v28.66 BUG-FIX: Zeiten des Hauptevents als Fallback für Sub-Events ohne
    // eigene Zeiten. Die Start-/End-Felder eines Sub-Events sind optional (neue
    // Drafts starten auf '' und der DatePicker ist löschbar) und werden — anders
    // als beim Hauptevent — von der Wizard-Prüfung nicht eingefordert. Bisher
    // landeten StartDate UND EndDate dann leer in DEX_Events; der
    // DEX_CreateOutlookEvent-Flow rechnet convertFromUtc(coalesce(OutlookEnd,
    // EndDate)) = convertFromUtc(null) und „Create event (V4)" bricht ab — für
    // dieses Sub-Event entsteht gar kein Outlook-Termin (und der Flow-Lauf
    // scheitert komplett). Ein Sub-Event findet immer innerhalb seines
    // Hauptevents statt, also sind dessen Zeiten der richtige Default. Nach dem
    // Speichern stehen sie sichtbar in den Sub-Event-Feldern und können dort
    // präzisiert werden.
    const { start: parentStartIso, end: parentEndIso } = parentTimesIso();
    // v28.29 BUG-FIX: Kopfbild-Vererbung vom Hauptevent auf die Sub-Events.
    // Sub-Events haben EIGENE Outlook-Termine und eigene Mails, aber praktisch
    // nie ein eigenes Kopfbild — Schritt 23/24 wird pro Tab gepflegt, und die
    // Vorschau dort zeigte fälschlich das Event-Foto, obwohl im Sub-Tab gar
    // nichts hinterlegt war. Beim Speichern fiel der Sub-Event-Body deshalb
    // auf das Standard-DEX-Logo zurück: Hauptevent-Termin mit Foto,
    // Sub-Event-Termine mit Orb. Jetzt erbt jedes Sub-Event ohne eigenes Bild
    // das Bild des Hauptevents; ein eigenes Bild im Sub-Tab gewinnt weiterhin.
    const parentComm = resolveTopLevelCommState();
    const inheritedEmailLogo = await shrinkLogoB64(parentComm.emailLogoBase64 || '');
    const inheritedOutlookLogo = await shrinkLogoB64(parentComm.outlookLogoBase64 || '');
    // v11.60: aus dem Ref iterieren — der React-State ist beim Save evtl.
    // noch nicht propagiert, weil flushActiveCommTabToState() per setState
    // erst async wirkt. Der Ref hält synchron die letzten Tab-Werte.
    for (const draft of subEventsRef.current) {
      if (!draft.title || !draft.title.trim()) {
        // v29.19: Leere Drafts werden nicht gespeichert — aber ein BESTEHENDES
        // Sub-Event (dbId), dessen Titel gerade nur geleert ist, gilt trotzdem
        // als „behalten". Vorher fiel es in die Lücke zwischen zwei
        // „behalten"-Definitionen: Der Warn-Dialog in handleSubmitInner zählt
        // über dbIds (Draft existiert → kein Dialog), die Aufräum-Schleife
        // unten über keptDbIds (nicht drin → deleteEvent) — das Sub-Event
        // wurde samt Subsite und Anmeldungen OHNE Rückfrage gelöscht, nur
        // weil jemand den Titel zum Neutippen geleert und dann gespeichert
        // hat. Löschen geht weiterhin — über das X an der Karte, mit Dialog.
        if (draft.dbId) keptDbIds.add(draft.dbId);
        continue;
      }
      // v11.57: Pro-Sub-Event Kommunikations-Felder. Wenn der Organizer für
      // den Sub-Event eigene Werte in Step 5 gesetzt hat, verwenden wir die;
      // sonst fallback auf die Top-Level-Werte (Backward-Compat für
      // Sub-Events ohne eigene Communication-Einstellungen).
      const subEmailLang = draft.emailLanguage || emailLanguage;
      const subOutlookBodyRaw = (typeof draft.outlookBody === 'string' && draft.outlookBody !== '') ? draft.outlookBody : '';
      const subOutlookHeading = draft.outlookHeading || draft.title || '';
      const subOutlookSub = draft.outlookSubheading || '';
      // v30.7: Kalender-Tage heissen „Di. 01.09.2026" — ohne eigenen Betreff
      // fiel der Outlook-Flow auf diesen Tages-Titel zurueck, und im Kalender
      // der Teilnehmer stand nur das Datum (das der Termin ohnehin traegt).
      // Default ist deshalb der NAME DES HAUPTEVENTS; ein im Kommunikations-
      // Reiter gesetzter eigener Betreff gewinnt weiterhin.
      const subOutlookSubject = (draft.outlookSubject || '').trim() || (subEventCalendar ? title.trim() : '');
      // v28.29: eigenes Bild des Sub-Events gewinnt, sonst erbt es das
      // Kopfbild des Hauptevents (statt still auf den Orb zu fallen).
      // v29.20 (Audit): auch das EIGENE Sub-Logo verkleinern — der
      // v28.10-Schutz lief nur über die geerbten Parent-Logos. Ein auf dem
      // Sub-Reiter hochgeladenes unkomprimiertes Foto steckte bis zu dreimal
      // im Payload und riss das Sub-Event ins SharePoint-2-MB-Limit.
      // v29.32: Zwischen eigenem Logo und geerbtem Parent-Logo steht jetzt das
      // EIGENE Bild des Sub-Events (s. subPhotoAsLogo) — bei Terminreihen mit
      // Foto je Termin kam sonst überall das Bild des Hauptevents an. Nur
      // aufgerufen, wenn kein eigenes Logo gesetzt ist (spart Laden/Komprimieren).
      const subOwnPhotoLogo = draft.emailLogoBase64 ? '' : await subPhotoAsLogo(draft);
      const subEmailLogo = (draft.emailLogoBase64 ? await shrinkLogoB64(draft.emailLogoBase64) : '') || subOwnPhotoLogo || inheritedEmailLogo;
      const subOutlookLogo = (draft.outlookLogoBase64 ? await shrinkLogoB64(draft.outlookLogoBase64) : '') || subOwnPhotoLogo || inheritedOutlookLogo;
      // Outlook-Body wrappen. v26.59 BUG-FIX: Ohne eigenen Text wurde der Body
      // bisher LEER gespeichert („der Flow setzt einen Default" — stimmte
      // nicht, der Flow mappt 1:1) → die Outlook-Einladung der Sub-Events kam
      // komplett ohne Text an. Jetzt bekommen Sub-Events denselben
      // Default-Body wie das Haupt-Event (v7.4-Pattern: Anmeldebestätigung +
      // Abmelde-Hinweis über die App + Organizer-Kontakt).
      let wrappedSubOutlookBody = '';
      {
        const orgNamesSub = formatOrganizerList([organizer], subEmailLang);
        const vars = {
          EventTitle: draft.title.trim(),
          // v27.5: {{Organizer}} auf normalisierte Namen ("Vorname Nachname",
          // mit „und" verbunden) statt roher „Nachname, Vorname"-Join.
          Organizer: orgNamesSub || organizer,
          ContactEmail: contactEmail.trim(),
          Location: draft.location || '',
          Address: '',
          StartDate: draft.startDate ? new Date(draft.startDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
          EndDate: draft.endDate ? new Date(draft.endDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
        };
        const escHtmlSub = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        const APP_URL_SUB = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';
        const defaultSubBody = subEmailLang === 'EN'
          ? `<p>You are registered for the event <strong>${escHtmlSub(draft.title.trim())}</strong>.</p>`
            + `<p>If you are unable to attend, please cancel your registration in time via the <a href="${APP_URL_SUB}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;My Events&ldquo;).</p>`
            + `<p>For organizational questions please contact <strong>${escHtmlSub(orgNamesSub || 'the organizer')}</strong>.</p>`
          : `<p>Ihr seid für das Event <strong>${escHtmlSub(draft.title.trim())}</strong> angemeldet.</p>`
            + `<p>Falls ihr nicht teilnehmen könnt, meldet euch bitte rechtzeitig über die <a href="${APP_URL_SUB}" style="color:#86bc25;font-weight:600;">DEX App</a> (&bdquo;Meine Events&ldquo;) ab.</p>`
            + `<p>Bei organisatorischen Fragen wendet euch bitte an <strong>${escHtmlSub(orgNamesSub || 'den Organizer')}</strong>.</p>`;
        const resolvedBody = subOutlookBodyRaw ? replacePlaceholders(subOutlookBodyRaw, vars) : defaultSubBody;
        const resolvedHead = subOutlookHeading ? replacePlaceholders(subOutlookHeading, vars) : draft.title.trim();
        // v27.5: Default-Unter-Überschrift = Ort (nicht Datum).
        const resolvedSub2 = subOutlookSub ? replacePlaceholders(subOutlookSub, vars) : (draft.location || undefined);
        // v18.73: Sub-Events erben das Header-Bild-Layout des Hauptevents.
        // v28.29: ohne eigenes/geerbtes Bild wird die Breite gekappt (Orb).
        const wrapped = buildOutlookBody(resolvedHead, resolvedBody, resolvedSub2, headerLayoutFor(subOutlookLogo), outlookTeamsLink(), (subEmailLang || '').toUpperCase() !== 'EN');
        wrappedSubOutlookBody = wrapped.replace(/\{\{ORB_URL\}\}/g, subOutlookLogo || getCachedOrbBase64() || '');
      }
      // Sub-Event-EmailTemplateOverrides: Logo-Piggybacks (Top-Level-Pattern)
      // + ab v14.4 die echten Mail-Text-Overrides pro Sub-Event
      // (Anmeldung/Warteliste/Abmeldung/Nachrücken).
      const subDraftOverrides = draft.emailTemplateOverrides || {};
      // v15.3: Inheritance-Flags entfallen — Sub-Events sind seit v15.3
      // vollwertige Events mit eigener Konfiguration. Der Piggyback-Key
      // `_inheritFlags` wird nicht mehr geschrieben.
      const subOverridesMerged: Record<string, unknown> = {
        ...subDraftOverrides,
        ...(subEmailLogo ? { _eventLogo: subEmailLogo } : {}),
        ...outlookLogoPiggyback(subEmailLogo, subOutlookLogo),
        // v18.73: Sub-Event erbt das Header-Bild-Layout des Hauptevents, damit
        // auch die Sub-Event-Mails den gleichen Bild-Kopf nutzen.
        ...headerImageLayoutConfig,
      };
      const subEmailOverrides = Object.keys(subOverridesMerged).length > 0
        ? JSON.stringify(subOverridesMerged)
        : '';
      // v15.3: Sub-Event-eigene strukturierte Adresse serialisieren (analog
      // zum Hauptevent-Top-Level-Pattern). Wenn alle vier Komponenten leer
      // sind, wird ein leerer String gespeichert.
      const draftAddr = draft.locationAddress || { street: '', houseNo: '', zip: '', city: '' };
      const draftHasAddress = !!(draftAddr.street || draftAddr.houseNo || draftAddr.zip || draftAddr.city);
      const draftLocationAddress = draftHasAddress ? JSON.stringify(draftAddr) : '';
      const draftAgendaJson = JSON.stringify(draft.agenda || []);
      const draftTransfersJson = JSON.stringify(draft.transferTimes || []);
      const childPayload = {
        title: draft.title.trim(),
        type: 'Other',
        status: 'Active',
        description: draft.description || '',
        location: draft.location || '',
        locationAddress: draftLocationAddress,
        locationFilter: draft.locationFilter || '',
        audience: draft.audience || '',
        filterMode: (draft.filterMode === 'AND' ? 'AND' : 'OR'),
        // v22.10: Ausschluss-Liste pro Sub-Event mitpersistieren (createEvent
        // schreibt sie in die Spalte ExcludedUsers).
        excludedUsers: draft.excludedUsers || [],
        // v28.66: ohne eigene Zeit die Zeit des Hauptevents erben (s.o.).
        startDate: draft.startDate || parentStartIso || '',
        // v22.17: NIE ein leeres EndDate persistieren — sonst rechnet der
        // DEX_CreateOutlookEvent-Flow convertFromUtc(coalesce(OutlookEnd,
        // EndDate)) = convertFromUtc(null) und „Create event (V4)" stürzt ab
        // (kein Outlook-Termin). Fallback auf das Start-Datum.
        // v28.66: greift auch, wenn das Sub-Event gar keine eigene Zeit hat —
        // dann kommt das Ende (ersatzweise der Start) des Hauptevents.
        endDate: draft.endDate || draft.startDate || parentEndIso || parentStartIso || '',
        registrationDeadline: draft.registrationDeadline || '',
        lastDeregisterDate: draft.lastDeregisterDate || '',
        maxParticipants: draft.maxParticipants || 0,
        waitlistEnabled: typeof draft.waitlistEnabled === 'boolean' ? draft.waitlistEnabled : true,
        mandatoryRegistration: !!draft.mandatory, // v24.64: Pflicht-Sub-Event
        eventImageUrl: '',
        organizer: sanitizedOrgPair.orgString,
        organizerEmail: sanitizedOrgPair.orgEmailString,
        outlookEventId: '',
        outlookBody: wrappedSubOutlookBody,
        outlookSubject: subOutlookSubject || undefined,
        outlookStart: (draft.outlookStart || '') || undefined,
        outlookEnd: (draft.outlookEnd || '') || undefined,
        outlookLocation: (draft.outlookLocation || '') || undefined,
        // v29.52: ganztägig mitschreiben — sonst kippt der Haken beim Speichern zurück.
        allDay: !!draft.allDay,
        showAsFree: !!draft.showAsFree,
        // v30.26: Die Online-Meeting-Entscheidung gilt event-weit — jeder
        // Termin einer Reihe bekommt dann seine eigene Teams-Besprechung.
        outlookIsOnlineMeeting: onlineMeetingMode === 'auto',
        skipOrganizerInvite: !orgGetsSubInvites, // v29.55
        agenda: draftAgendaJson,
        transfers: draftTransfersJson,
        documents: '[]',
        funZone: '[]',
        quizClusterSize: 1,
        emailLanguage: subEmailLang,
        emailTemplateOverrides: subEmailOverrides,
        disableEmails: !!draft.disableEmails,
        // v19.22: granulare An-/Abmelde-Mail-Schalter pro Sub-Event persistieren.
        disableRegistrationEmail: !!draft.disableRegistrationEmail,
        disableCancellationEmail: !!draft.disableCancellationEmail,
        autoDeregisterOnDecline: !!draft.autoDeregisterOnDecline,
        inactiveHandling: (draft.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify') as 'notify' | 'autoderegister',
        disableOutlook: !!draft.disableOutlook,
        isFictive: isFictive,
        askSalutation: !!draft.askSalutation,
        // v29.20 (Audit): kanonisch serialisieren wie der Update-Zweig —
        // roh gingen leere Felder und leere Options-Slots (neue Drafts
        // starten mit ['','']) direkt auf die Anmeldeseite, EN-Varianten
        // wurden am bilingual-Schalter vorbei geschrieben und Wizard-interne
        // Properties (visible) landeten im Storage-JSON.
        customFields: serializeCustomFields(draft.customFields || [], bilingualFields),
        parentEventId: parentEventId,
      };
      if (draft.dbId) {
        // v11.69: Recreate-Pfad via Modal-Auswahl. Wenn der Organizer im
        // Outlook-Confirm-Modal ein Sub-Event mit `noOutlookYet=true`
        // angehakt hat (es existiert noch kein Outlook-Termin), wird das
        // bestehende DEX_Events-Item per `deleteEventItemOnly` entfernt und
        // eine NEUE DEX_Events-Zeile angelegt — wobei die bestehende
        // Subsite + Teilnehmerliste an die neue Zeile gekoppelt werden.
        // Damit triggert der `DEX_CreateOutlookEvent`-Flow (GetOnNewItems)
        // auf dem neuen Item und legt den Outlook-Termin an. Die alte
        // Subsite mit ALLEN Anmeldungen bleibt unangetastet.
        const initialMeta = initialSubEventOutlookMeta[draft.dbId];
        if (pendingOutlookRecreateForSubEventsRef.current.includes(draft.dbId)) {
          const subsiteUrlForReuse = initialMeta?.subsiteUrl || '';
          const regListNameForReuse = initialMeta?.registrationListName || 'Teilnehmer';
          if (subsiteUrlForReuse && regListNameForReuse) {
            // v29.21 (Audit): Den Rückgabewert PRÜFEN — deleteEventItemOnly
            // wirft nie, es meldet Fehlschlag als false (das catch hier war
            // toter Code). Schlug der Delete fehl (429/403), existierte die
            // alte Zeile weiter, createEvent legte eine ZWEITE Zeile auf
            // dieselbe Subsite, und die Aufräum-Schleife am Ende rief
            // deleteEvent auf die alte Id — KASKADIEREND: Sie recycelte die
            // geteilte Subsite mit allen Anmeldungen, auf die die frische
            // Zeile zeigt. Deshalb: bei Fehlschlag den Recreate für dieses
            // Sub-Event abbrechen, die dbId als behalten registrieren und in
            // den normalen Update-Pfad fallen — der Termin fehlt dann
            // weiterhin, aber es geht nichts verloren.
            const itemDeleted = await deleteEventItemOnly(draft.dbId).catch(() => false);
            if (!itemDeleted) {
              console.warn('[DEX][v29.21] Recreate abgebrochen — deleteEventItemOnly fehlgeschlagen, Sub-Event bleibt unangetastet:', draft.dbId);
              showAlert(isDe
                ? `Der Outlook-Termin für „${draft.title || 'Sub-Event'}" konnte nicht neu angelegt werden (SharePoint hat den Austausch des Eintrags abgelehnt). Das Sub-Event und seine Anmeldungen sind unverändert — bitte versuche es später erneut.`
                : `The Outlook appointment for "${draft.title || 'sub-event'}" could not be recreated (SharePoint rejected replacing the entry). The sub-event and its registrations are untouched — please try again later.`, { variant: 'error' });
            } else {
              // Reuse-Payload: bestehende Subsite + Teilnehmerliste an die
              // neue DEX_Events-Zeile koppeln. disableOutlook explizit false,
              // outlookEventId leer, damit der Flow sauber neu schreibt.
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const reusePayload: any = {
                ...childPayload,
                disableOutlook: false,
                outlookEventId: '',
                existingSubsiteUrl: subsiteUrlForReuse,
                existingRegistrationListName: regListNameForReuse,
                onProgress: subOnProgress,
              };
              try {
                const recreatedId = await createEvent(reusePayload);
                // v27.11: Sub-Event-Bild auch auf dem Recreate-Pfad persistieren.
                await persistSubEventImage(recreatedId, draft);
              } catch (err) {
                console.warn('[DEX][v11.69] Sub-Event-Recreate mit Subsite-Reuse fehlgeschlagen:', draft.dbId, err);
              }
              // NICHT zu keptDbIds hinzufügen — das alte Item wurde gelöscht
              // und die neue Zeile hat eine andere ID.
              continue;
            }
          } else {
            console.warn('[DEX][v11.69] Recreate angefordert aber keine subsiteUrl/registrationListName vorhanden — Sub-Event:', draft.dbId, 'meta:', initialMeta);
            // Fall durch zum normalen Update-Pfad — wenigstens die Felder
            // werden persistiert; Outlook-Termin entsteht aber nicht.
          }
        }
        // Spezialfall: Outlook nachträglich aktivieren via DisableOutlook-
        // Toggle (alter Pfad vor v11.69). Wenn der User auf einem
        // **bestehenden** Sub-Event die "Outlook erstellen"-Checkbox
        // einschaltet (DisableOutlook: true → false) und bisher kein
        // Outlook-Termin angelegt wurde (OutlookEventId leer), muss das
        // Sub-Event neu angelegt werden — der Power-Automate-Flow
        // `DEX_CreateOutlookEvent` triggert ausschließlich auf NEUE
        // DEX_Events-Items (GetOnNewItems). Ein reines MERGE-Update würde
        // den Flow nie anstoßen → kein Outlook-Termin.
        const wasOutlookDisabled = !!initialMeta?.disableOutlook;
        const nowOutlookEnabled = !draft.disableOutlook;
        const hadOutlookEventId = !!(initialMeta?.outlookEventId);
        // v28.69: „Fehlende Termine jetzt anlegen" erzwingt denselben Pfad —
        // ein reines Update triggert den GetOnNewItems-Flow nie.
        const forcedRecreate = forceOutlookRecreateRef.current.has(draft.dbId) && nowOutlookEnabled;
        const needsOutlookRecreate = forcedRecreate || (wasOutlookDisabled && nowOutlookEnabled && !hadOutlookEventId);
        if (needsOutlookRecreate) {
          // v11.69: Seit dem Subsite-Reuse-Pfad muss hier KEINE destruktive
          // Lösch-Aktion mehr passieren. Wir entfernen nur die DEX_Events-
          // Zeile und legen sie mit `existingSubsiteUrl` neu an — alle
          // Anmeldungen, TeilnehmerIDs und die Subsite bleiben unangetastet.
          // Daher auch kein window.confirm mehr nötig.
          const subsiteUrlForReuse = initialMeta?.subsiteUrl || '';
          const regListNameForReuse = initialMeta?.registrationListName || 'Teilnehmer';
          if (subsiteUrlForReuse && regListNameForReuse) {
            try {
              await deleteEventItemOnly(draft.dbId);
            } catch { /* Delete-Fehler: trotzdem versuchen, neu anzulegen */ }
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const reusePayload: any = {
              ...childPayload,
              disableOutlook: false,
              outlookEventId: '',
              existingSubsiteUrl: subsiteUrlForReuse,
              existingRegistrationListName: regListNameForReuse,
              onProgress: subOnProgress,
            };
            try {
              const recreatedId = await createEvent(reusePayload);
              // v27.11: Sub-Event-Bild auch auf dem Legacy-Recreate-Pfad persistieren.
              await persistSubEventImage(recreatedId, draft);
            } catch (err) {
              console.warn('[DEX][v11.69] Legacy-Toggle-Recreate fehlgeschlagen:', draft.dbId, err);
            }
            continue;
          } else {
            // Edge-Case: kein subsiteUrl auf dem alten Item bekannt (sehr
            // alte Events). In dem Fall fallen wir auf den destruktiven
            // Legacy-Pfad zurück — mit Confirm.
            const msg = isDe
              ? `Beim Sub-Event „${draft.title}" wurde „Outlook-Termin erstellen" nachträglich aktiviert, aber es konnte keine bestehende Teilnehmer-Subsite ermittelt werden.\n\n`
                + `Wenn du jetzt fortfährst, wird das Sub-Event komplett neu aufgesetzt — vorhandene Anmeldungen, Teilnehmer-IDs und die Teilnehmer-Subsite gehen verloren (landen 93 Tage im Papierkorb).\n\n`
                + `Trotzdem fortfahren?`
              : `On sub-event "${draft.title}" you turned on "Create Outlook event" after the fact, but no existing participant subsite could be determined.\n\n`
                + `If you continue, the sub-event will be re-created from scratch — existing registrations, participant IDs and the participant subsite will be lost (recycled for 93 days).\n\n`
                + `Continue anyway?`;
            const confirmed = await confirmDialog(msg, { danger: true, title: isDe ? 'Sub-Event neu aufsetzen' : 'Recreate sub-event', confirmLabel: isDe ? 'Trotzdem fortfahren' : 'Continue anyway' });
            if (confirmed) {
              try {
                await deleteEvent(draft.dbId);
              } catch { /* delete-Fehler darf Re-Create nicht blockieren */ }
              const recreatedLegacyId = await createEvent({ ...childPayload, onProgress: subOnProgress });
              // v27.11: Sub-Event-Bild auch hier persistieren.
              await persistSubEventImage(recreatedLegacyId, draft);
              continue;
            } else {
              draft.disableOutlook = true;
              childPayload.disableOutlook = true;
            }
          }
        }
        keptDbIds.add(draft.dbId);
        // v29.57: Unverändertes Sub-Event überspringen — spart bei einer
        // Terminreihe die große Mehrheit der Schreibvorgänge (s. Kommentar
        // an subPersistKey). Nur wenn BEIDES unverändert ist: der Entwurf
        // selbst UND alles am Hauptevent, was ein Sub-Event erbt.
        // v30.10: … UND der wirksame Outlook-Betreff schon in der Zeile
        // steht. Der v30.7-Fallback (Kalender-Tag → Hauptevent-Titel) wird
        // erst BEIM SCHREIBEN berechnet — im Entwurf ändert sich dadurch
        // nichts, der Skip hielt Bestands-Terminreihen deshalb für
        // unverändert, OutlookSubject blieb leer und der UpdateEvent-Flow
        // fiel per coalesce weiter auf den Tages-Titel zurück, egal wie oft
        // aktualisiert wurde. Genau die Falle, vor der der subPersistKey-
        // Kommentar warnt: ein abgeleiteter Wert, den keiner der beiden
        // Schlüssel sieht.
        const storedSubRow = childEventsOf(parentEventId).find(k => k.id === draft.dbId);
        const subjectInSync = subOutlookSubject === ((storedSubRow && storedSubRow.outlookSubject) || '').trim();
        if (subGateUnchanged && subjectInSync && !draft.imageFile && !draft.imageRemoved) {
          const before = initialSubPersistRef.current[draft.dbId];
          if (before !== undefined && before === subPersistKey(draft)) {
            skippedSubCount++;
            stepDone++;
            if (onStep) onStep(stepDone, stepTotal, shortSubEventTitle(draft.title, title) || draft.title);
            continue;
          }
        }
        // Update bestehender Sub-Event: nur geänderte Felder patchen. CustomFields
        // werden als JSON-String serialisiert — v17.22: zentraler
        // serializeCustomFields-Helper, damit der Sub-Event-Pfad dieselben
        // EN-Varianten + Options-Pairing erhält wie Top-Level (vorher droppte
        // dieser Pfad labelEn/helpTextEn/confirmLabelEn/optionsEn still).
        const cfJson = JSON.stringify(serializeCustomFields(draft.customFields || [], bilingualFields));
        // v11.57: Sub-Event-Kommunikations-Felder mit-persistieren — bisher
        // wurden Mail-Sprache, Outlook-Body, Logos pro Sub-Event nur am
        // Top-Level gespeichert. Mit den Tabs in Step 5 kann jeder Sub-Event
        // jetzt seine eigene Konfiguration haben.
        // v11.57: Outlook-Update-Flag pro Sub-Event setzen, wenn dieses
        // Sub-Event als „outlookDirty" markiert wurde (vom Confirm-Modal
        // entschieden).
        const subUpdates: Record<string, unknown> = {
          'Title': childPayload.title,
          'Description': childPayload.description,
          'Location': childPayload.location,
          'StartDate': childPayload.startDate || null,
          // v28.66 BUG-FIX: hier fehlte der v22.17-Schutz — beim Speichern
          // eines BESTEHENDEN Sub-Events ohne End-Zeit wurde EndDate mit null
          // überschrieben (Outlook-Flow-Crash, s.o.). childPayload trägt den
          // Fallback bereits; die zweite Stufe bleibt als Absicherung stehen.
          'EndDate': childPayload.endDate || childPayload.startDate || null,
          'RegistrationDeadline': childPayload.registrationDeadline || null,
          'MaxParticipants': childPayload.maxParticipants,
          // v20.0 BUG-FIX (Audit): diese Felder wurden beim UPDATE bestehender
          // Sub-Events NIE mitgeschrieben (nur beim Create via childPayload) —
          // Änderungen an Sichtbarkeit (v19.27 AudiencePicker), Adresse, Agenda,
          // Transferzeiten, Abmeldefrist, Warteliste und Anrede-Abfrage gingen
          // beim Speichern eines bestehenden Sub-Events still verloren
          // (gleiche Bug-Klasse wie v19.32 bei den Mail-Flags).
          'WaitlistEnabled': childPayload.waitlistEnabled,
          // v29.20 (Audit): Organizer-Aenderungen erreichten bestehende
          // Sub-Events nie — nur der Create-Pfad schrieb sie (childPayload).
          // Sub-Event-Mails und Flows lesen die Zeile des Sub-Events und
          // adressierten nach einem Organizer-Wechsel weiter die alten
          // Personen. Der Doku-Kommentar oben verspricht die Vererbung
          // ausdruecklich („Alle Sub-Events erben Metadaten (Organizer, …)").
          'Organizer': childPayload.organizer,
          'OrganizerEmail': childPayload.organizerEmail,
          'MandatoryRegistration': childPayload.mandatoryRegistration, // v24.64: Pflicht-Sub-Event
          'LastDeregisterDate': childPayload.lastDeregisterDate || null,
          'LocationAddress': childPayload.locationAddress,
          'LocationFilter': childPayload.locationFilter,
          'Audience': childPayload.audience,
          'FilterMode': childPayload.filterMode,
          // v22.10: Ausschluss-Liste pro Sub-Event aktualisieren (semikolon-sep.).
          'ExcludedUsers': (draft.excludedUsers || []).filter(Boolean).join(';'),
          'Agenda': childPayload.agenda,
          'Transfers': childPayload.transfers,
          'AskSalutation': childPayload.askSalutation,
          'DisableEmails': childPayload.disableEmails,
          // v19.32 BUG-FIX: die granularen An-/Abmelde-Mail-Flags + Auto-Abmeldung
          // wurden beim UPDATE bestehender Sub-Events NICHT mitgeschrieben (nur
          // beim Create), daher gingen sie nach dem Speichern verloren.
          'DisableRegistrationEmail': childPayload.disableRegistrationEmail,
          'DisableCancellationEmail': childPayload.disableCancellationEmail,
          'AutoDeregisterOnDecline': childPayload.autoDeregisterOnDecline,
          'InactiveHandling': childPayload.inactiveHandling || 'notify',
          'DisableOutlook': childPayload.disableOutlook,
          'OutlookSubject': subOutlookSubject,
          'OutlookStart': (draft.outlookStart || '') || null,
          'OutlookEnd': (draft.outlookEnd || '') || null,
          'OutlookLocation': (draft.outlookLocation || '') || '',
          // v29.52: ganztägig auch beim UPDATE bestehender Sub-Events — genau
          // die Klasse Fehler, die v19.32/v20.0/v29.20 schon dreimal hatten.
          'AllDay': !!draft.allDay,
          'ShowAsFree': !!draft.showAsFree, // v29.54
          'OutlookIsOnlineMeeting': onlineMeetingMode === 'auto', // v30.26
          'SkipOrganizerInvite': !orgGetsSubInvites, // v29.55
          'EmailLanguage': childPayload.emailLanguage,
          // v29.42: Fußzeilen-Link auch auf dem direkten Sub-Event-Schreibweg
          // normalisieren (der läuft nicht über EventService.updateEvent).
          'OutlookBody': normalizeMadeWithLink(childPayload.outlookBody || ''),
          'EmailTemplateOverrides': childPayload.emailTemplateOverrides,
          'EmailImageBase64': subEmailLogo || '',
          'CustomFields': cfJson,
        };
        // v22.15: Auto-Heilung — steht das Sub-Event auf „Abgeschlossen"
        // (z.B. vom Auto-Cleanup wegen eines alten Testdatums), das End-Datum
        // liegt nach dieser Bearbeitung aber in der Zukunft, zurück auf Aktiv.
        {
          const storedChild = childEventsOf(parentEventId).find(k => k.id === draft.dbId);
          const subEndIso = childPayload.endDate || childPayload.startDate || '';
          if (storedChild && storedChild.status === 'Completed' && subEndIso && new Date(subEndIso).getTime() > Date.now()) {
            subUpdates['EventStatus'] = 'Active';
          }
        }
        // OutlookDirty + Update wird vom Aufrufer (handleSubmit) anhand des
        // jeweiligen Sub-Event-Snapshots gesteuert — siehe pendingSubUpdates.
        // v29.21 (Audit): Ergebnis PRÜFEN — updateEvent wirft nie, es meldet
        // Fehlschlag als false. Vorher lief der Balken weiter und der Save
        // endete mit „Änderungen gespeichert!", auch wenn einzelne Sub-Events
        // (429, zu großer Payload) nie ankamen.
        // v29.77: skipReload — sonst laedt JEDES Sub-Event-Update die komplette
        // Event-Liste (28 MB) neu. Der eine Reload kommt vom Hauptevent-Update.
        const subOk = await updateEvent(draft.dbId, subUpdates, { skipReload: true });
        if (!subOk) failedSubTitles.push(shortSubEventTitle(draft.title, title) || draft.title);
        // v29.74: Bei Drosselung ANHALTEN statt weiterhaemmern. Zwei
        // Fehlschlaege in Folge waehrend aktiver Drossel-Schranke heisst:
        // SharePoint will gerade keine weiteren Schreibzugriffe von diesem
        // Nutzer. Die restlichen Termine trotzdem zu versuchen (jeder mit
        // eigenen Retries) hat einen Organizer bis zur NUTZER-SPERRE
        // (Throttle.htm) eskaliert. Lieber ehrlich abbrechen — gespeichert
        // ist gespeichert, der Rest kommt beim naechsten Save.
        if (!subOk) {
          consecutiveSubFailures++;
          if (consecutiveSubFailures >= 2 && isThrottled()) {
            abortedForThrottle = true;
            break;
          }
        } else {
          consecutiveSubFailures = 0;
        }
        // v29.74: Atempause zwischen den Schreibzugriffen — 19 MERGEs im
        // Renn-Tempo sind genau das Muster, das die Drosselung ausloest.
        await new Promise<void>(r => setTimeout(r, 250));
        // v27.11: eigenes Sub-Event-Bild persistieren (Upload/Entfernen).
        await persistSubEventImage(draft.dbId, draft);
      } else {
        const newSubId = await createEvent({ ...childPayload, onProgress: subOnProgress });
        if (!newSubId) failedSubTitles.push(shortSubEventTitle(draft.title, title) || draft.title);
        // v27.11: Bild fürs frisch angelegte Sub-Event hochladen (braucht die
        // neue DEX_Events-Item-Id aus createEvent).
        await persistSubEventImage(newSubId, draft);
      }
      stepDone++;
      if (onStep) onStep(stepDone, stepTotal, shortSubEventTitle(draft.title, title) || draft.title);
    }
    // Entfernte Sub-Events aufräumen: deleteEvent löscht kaskadierend auch
    // die Subsite (Teilnehmerliste) und queued einen Outlook-DeleteEvent.
    //
    // v29.48 BUG-FIX: Das Ergebnis wurde hier verworfen („Delete-Fehler darf
    // Save nicht blockieren") — deleteEvent WIRFT aber gar nicht, es liefert
    // false. Ein abgelehntes Löschen (typisch: HTTP 429, SharePoint drosselt
    // nach 20 Sub-Event-Schreibvorgängen) war damit unsichtbar: der Wizard
    // meldete „Änderungen gespeichert!", und der abgewählte Tag stand nach dem
    // Neuladen wieder in der Liste. Genau das war der Kalender-Fall aus der
    // Rückmeldung — 28.09. und 30.09. blieben stehen, obwohl sie abgewählt
    // waren, während der dazwischenliegende 29.09. verschwand.
    const failedDeleteTitles: string[] = [];
    for (const oldId of initialSubEventDbIds) {
      if (!keptDbIds.has(oldId)) {
        const gone = await deleteEvent(oldId).catch(() => false);
        if (!gone) {
          const stored = childEventsOf(parentEventId).find(k => k.id === oldId);
          failedDeleteTitles.push(stored ? (shortSubEventTitle(stored.title, title) || stored.title) : `#${oldId}`);
        }
      }
    }
    // v29.21 (Audit): gescheiterte Sub-Events benennen statt still erfolgreich
    // zu wirken — der Organizer entscheidet dann selbst, ob er erneut speichert.
    if (abortedForThrottle) {
      showAlert(isDe
        ? `SharePoint bremst gerade alle Schreibzugriffe (Drosselung). Der Speichervorgang wurde nach ${stepDone} von ${stepTotal} Terminen angehalten, damit dein Konto nicht gesperrt wird. Was gespeichert ist, bleibt gespeichert — bitte warte ein paar Minuten und speichere dann erneut; bereits gesicherte Termine werden dabei übersprungen.`
        : `SharePoint is currently throttling all writes. Saving stopped after ${stepDone} of ${stepTotal} dates to protect your account from being blocked. Everything saved so far is kept — please wait a few minutes and save again; dates already saved will be skipped.`, { variant: 'error' });
    } else if (failedSubTitles.length > 0) {
      showAlert(isDe
        ? `${failedSubTitles.length} Sub-Event${failedSubTitles.length === 1 ? '' : 's'} konnte${failedSubTitles.length === 1 ? '' : 'n'} nicht gespeichert werden: ${failedSubTitles.join(', ')}. Die übrigen Änderungen sind gespeichert — bitte speichere erneut, um es nochmal zu versuchen.`
        : `${failedSubTitles.length} sub-event${failedSubTitles.length === 1 ? '' : 's'} could not be saved: ${failedSubTitles.join(', ')}. All other changes are saved — please save again to retry.`, { variant: 'error' });
    }
    // v29.57: Nach dem Save ist der aktuelle Stand der neue Vergleichspunkt —
    // sonst würde ein zweiter Save in derselben Sitzung alles erneut schreiben.
    {
      const map: Record<string, string> = {};
      for (const se of subEventsRef.current) if (se.dbId) map[se.dbId] = subPersistKey(se);
      initialSubPersistRef.current = map;
      subTopGateInitialRef.current = subTopGateKey();
      if (skippedSubCount > 0) {
        // eslint-disable-next-line no-console
        dlog('perf', `[DEX] ${skippedSubCount} unveränderte Sub-Events übersprungen (kein Schreibvorgang).`);
      }
    }
    if (failedDeleteTitles.length > 0) {
      showAlert(isDe
        ? `${failedDeleteTitles.length} abgewählte${failedDeleteTitles.length === 1 ? 'r Termin konnte' : ' Termine konnten'} nicht gelöscht werden: ${failedDeleteTitles.join(', ')}. ${failedDeleteTitles.length === 1 ? 'Er steht' : 'Sie stehen'} deshalb weiterhin in der Liste — bitte speichere erneut.`
        : `${failedDeleteTitles.length} deselected date${failedDeleteTitles.length === 1 ? '' : 's'} could not be deleted: ${failedDeleteTitles.join(', ')}. ${failedDeleteTitles.length === 1 ? 'It is' : 'They are'} therefore still in the list — please save again.`, { variant: 'error' });
    }
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
      const wrappedOutlook = buildOutlookBody(resolvedOlHeading, resolvedBody, resolvedOlSub, headerLayoutFor(effOutlookLogo), outlookTeamsLink(), (emailLanguage || '').toUpperCase() !== 'EN');
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
      updates['ActiveFrom'] = activeFrom ? new Date(activeFrom).toISOString() : null;
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
        const outlookTotal =
          ((!disableOutlook && pendingOutlookUpdateForTopRef.current) ? 1 : 0)
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
        if (!disableOutlook && pendingOutlookUpdateForTopRef.current) {
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const ctx = (window as any).__dexSpfxContext;
            if (ctx && editEvent?.subsiteUrl) {
              const svc = new EventService(ctx);
              tickOutlook(isDe ? `Outlook wird aktualisiert… (${title || 'Hauptevent'})` : `Updating Outlook… (${title || 'main event'})`);
              await svc.queueOutlookEvent('', selectedEventId, title, 'UpdateEvent');
            }
          } catch { /* Outlook-Update optional */ }
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
                try {
                  await svc.queueOutlookEvent('', subId, subTitle, 'UpdateEvent');
                  await updateEvent(subId, { 'OutlookDirty': false }, { skipReload: true });
                } catch { /* einzelne Sub-Update-Fehler nicht eskalieren */ }
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
        activeFrom: activeFrom ? new Date(activeFrom).toISOString() : undefined,
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
          const wrapped = buildOutlookBody(resolvedHeading, resolvedBody, resolvedSub, headerLayoutFor(effOutlookLogo), outlookTeamsLink(), (emailLanguage || '').toUpperCase() !== 'EN');
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
    const items: OutlookConfirmItem[] = [];
    if (!editEvent) return { items };
    const snap = initialOutlookSnapshot.current;
    // v11.64: Datetime-Vergleich über Date.getTime(), nicht String. Sonst
    // kippt der Vergleich an Format-Unterschieden (snap kommt roh aus SP
    // mit „2026-09-24T16:00:00Z", currentStart geht durch
    // berlinLocalToUtcIso() und wird „2026-09-24T16:00:00.000Z" — gleicher
    // Zeitpunkt, anderer String). Das hat den Hauptevent in v11.63
    // fälschlich als „Startzeit, Endzeit geändert" gemeldet.
    const sameInstant = (a: string, b: string): boolean => {
      if (a === b) return true;
      if (!a && !b) return true;
      if (!a || !b) return false;
      const da = new Date(a).getTime();
      const db = new Date(b).getTime();
      if (isNaN(da) || isNaN(db)) return a === b;
      return da === db;
    };
    const currentTitle = title || '';
    const currentStart = startDate ? berlinLocalToUtcIso(startDate) : '';
    const currentEnd = endDate ? berlinLocalToUtcIso(endDate) : '';
    // Outlook-Body vergleich anhand des „rohen" Body (ohne Wrapper), da der
    // Wrapper bei jedem Save neu gebaut wird und dadurch immer „änderbar"
    // aussehen würde. Vergleich gegen den initial gestrippten Wert.
    // v29.21 (Audit): denselben Reinsert wie die outlookBody-Hydration
    // anwenden — sonst enthält der State den Platzhalter, der Snapshot den
    // gebackenen Namen, und der Vergleich meldete bei JEDEM Save „Termin-Text
    // geändert" (Dauer-False-Positive des Update-Modals).
    const initialStripped = reinsertOrganizerPlaceholder(stripOutlookWrapper(snap.outlookBody || ''), editEvent?.organizers || []);
    // v29.21 (Audit): Auf einem Sub-Tab liegt der AKTUELLE Top-Level-Body im
    // Top-Level-Slot (resolveTopLevelCommState) — der alte Fallback verglich
    // den Mount-Snapshot mit sich selbst, eine Hauptevent-Body-Änderung vor
    // dem Reiterwechsel wurde also nie erkannt: kein Update-Angebot, kein
    // OutlookDirty, Teilnehmer-Kalender blieben still veraltet. Betreff, Logo
    // und DisableOutlook daneben machten es längst so.
    const currentStripped = activeCommTabIdx === 0 ? (outlookBody || '') : (resolveTopLevelCommState().outlookBody || '');
    const currentTopLocation = outlookLocationOverride.trim() || buildOutlookLocation(location, { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity });
    const currentTopSubject = (resolveTopLevelCommState().outlookSubject || '').trim();
    // v19.20: globale Header-Bild-Layout-Änderung (Breite/Innenabstand) erkennen.
    // Das Layout steht NICHT im rohen Body (wird erst beim Wrappen angewendet),
    // betrifft aber den Hero-Bild-Kopf des Outlook-Termins — daher als eigenes
    // Änderungs-Feld „layout" werten, damit das Update-Modal aufgeht (und der
    // Grund klar als „Kopfbild" benannt wird, nicht irreführend als „Termin-Text").
    const initLayout = initialHeaderImageLayoutRef.current;
    const layoutChanged = headerImageLayout.width !== initLayout.width
      || headerImageLayout.paddingV !== initLayout.paddingV
      || headerImageLayout.paddingH !== initLayout.paddingH;
    // v28.30: Kopfbild-Wechsel erkennen. Das Bild steckt weder im rohen
    // Termin-Text (es wird erst beim Wrappen als {{ORB_URL}} eingesetzt) noch
    // im Layout — eine reine Bild-Änderung war für den Detektor deshalb
    // unsichtbar, und das Update-Modal bot nur Events an, bei denen zufällig
    // noch etwas anderes anders war.
    const curTopLogo = resolveTopLevelCommState().outlookLogoBase64 || '';
    const topLogoChanged = curTopLogo !== (snap.outlookLogo || '');
    const topChangedFields: Array<'title' | 'startDate' | 'endDate' | 'outlookBody' | 'location' | 'subject' | 'layout' | 'organizer' | 'logo'> = [];
    if (currentTitle !== (snap.title || '')) topChangedFields.push('title');
    if (!sameInstant(currentStart, snap.startDate || '')) topChangedFields.push('startDate');
    if (!sameInstant(currentEnd, snap.endDate || '')) topChangedFields.push('endDate');
    if (currentStripped !== initialStripped) topChangedFields.push('outlookBody');
    if (layoutChanged) topChangedFields.push('layout');
    // v29.38: reine Teams-Link-Änderung ebenfalls als Kopf-/Layout-Änderung
    // melden (eigener Grund wäre eine weitere Feld-Variante — der Termin-Text
    // ändert sich hier tatsächlich, deshalb 'outlookBody').
    // v30.26: auch der Online-Meeting-Modus (s. onlineMeetingChanged).
    if (onlineMeetingChanged() && topChangedFields.indexOf('outlookBody') < 0) topChangedFields.push('outlookBody');
    if (topLogoChanged) topChangedFields.push('logo');
    // v22.48: Organizer-Änderung. Der Outlook-Standardtext enthält die
    // Organizer-Namen („wendet euch bitte an …"). Solange der Text NICHT
    // individuell überschrieben ist (Body leer = Default), ändert eine
    // Organizer-Änderung den Termin-Text → als Outlook-relevant werten. Bei
    // custom Body bleibt der Text bewusst stehen → nicht melden.
    const currentOrganizers = organizer.split(';').map(s => s.trim()).filter(Boolean).join(';');
    const bodyIsDefault = !currentStripped.trim() && !initialStripped.trim();
    if (bodyIsDefault && currentOrganizers !== (snap.organizers || '')) topChangedFields.push('organizer');
    // v18.34: reine Ort-Änderung gilt ebenfalls als Outlook-relevant.
    if (currentTopLocation !== (snap.outlookLocation || '')) topChangedFields.push('location');
    // v18.42: reine Betreff-Änderung gilt ebenfalls als Outlook-relevant.
    if (currentTopSubject !== (snap.outlookSubject || '').trim()) topChangedFields.push('subject');
    // v18.44: abweichendes Outlook-Datum (Override) gilt als Termin-Änderung.
    if ((outlookStartOverride || '') !== (snap.outlookStart || '') && topChangedFields.indexOf('startDate') < 0) topChangedFields.push('startDate');
    if ((outlookEndOverride || '') !== (snap.outlookEnd || '') && topChangedFields.indexOf('endDate') < 0) topChangedFields.push('endDate');
    // v29.52: Umschalten auf/von „ganztägig" ist eine Termin-Änderung.
    if (!!allDay !== !!snap.allDay && topChangedFields.indexOf('startDate') < 0) topChangedFields.push('startDate');
    // v29.54: Wechsel zwischen Beschäftigt und Frei ändert den bestehenden
    // Termin ebenfalls — ohne diesen Vergleich bliebe er im Kalender stehen.
    if (!!showAsFree !== !!snap.showAsFree && topChangedFields.indexOf('startDate') < 0) topChangedFields.push('startDate');
    // v11.61: Beide Pointer prüfen — DEX_CreateOutlookEvent setzt nur
    // CalendarLink auf Erfolg, OutlookEventId bleibt leer. Wer beides
    // leer hat, hatte nie einen Outlook-Termin.
    const topHasOutlook = !!editEvent.outlookEventId || !!editEvent.calendarLink;
    // v18.45 BUG-FIX: für das Hauptevent IMMER dessen Top-Level-DisableOutlook
    // prüfen — nicht das rohe `disableOutlook` (das hält beim Speichern auf einem
    // Sub-Event-Tab den Sub-Wert). Sonst wurde das Hauptevent fälschlich im
    // Update-Modal gelistet, obwohl dort Outlook deaktiviert ist (z.B. Event mit
    // Outlook nur auf Sub-Event-Ebene).
    const topDisableOutlook = resolveTopLevelCommState().disableOutlook;
    // v18.51: Im „Nur für Sub-Events"-Modus (subEventsOnlyMode) ist das
    // Hauptevent von der Teilnehmer-Anmeldung ausgenommen — niemand meldet sich
    // direkt fürs Hauptevent an. Ein Outlook-Update-Hinweis fürs Hauptevent ist
    // dann sinnlos und wird unterdrückt (Sub-Events bekommen weiter ihre Hinweise).
    if (topChangedFields.length > 0 && !topDisableOutlook && topHasOutlook && !subEventsOnlyMode) {
      items.push({
        kind: 'top',
        eventId: editEvent.id,
        title: currentTitle || editEvent.title || '',
        changedFields: topChangedFields,
      });
    }
    // Sub-Events: pro Sub-Event vergleichen.
    // v11.60: subEventsRef statt subEvents — der Flush hat die aktuellen
    // UI-Werte gerade synchron in den Ref geschrieben, der React-State
    // ist noch nicht propagiert.
    for (const s of subEventsRef.current) {
      if (!s.dbId) continue;
      const initTitle = s.initialTitle || '';
      const initStart = s.initialStartDate || '';
      const initEnd = s.initialEndDate || '';
      const initBodyStripped = stripOutlookWrapper(s.initialOutlookBody || '');
      const curBodyStripped = (s.outlookBody || '');
      const hasOutlookEvId = !!s.initialOutlookEventId || !!s.initialCalendarLink;
      // v28.30: Kopfbild pro Sub-Event vergleichen — und zwar das WIRKSAME.
      // Seit v28.29 erbt ein Sub-Event ohne eigenes Bild das des Hauptevents;
      // wechselt dort das Bild, ändert sich also auch der Sub-Event-Termin,
      // obwohl im Sub-Draft selbst nichts steht.
      const initSubLogo = s.initialOutlookLogoBase64 || '';
      const curSubLogo = s.outlookLogoBase64 || curTopLogo;
      const subChangedFields: Array<'title' | 'startDate' | 'endDate' | 'outlookBody' | 'layout' | 'logo'> = [];
      if ((s.title || '') !== initTitle) subChangedFields.push('title');
      // v11.64: auch hier semantischer Vergleich — gleiche Falle wie oben.
      if (!sameInstant(s.startDate || '', initStart)) subChangedFields.push('startDate');
      if (!sameInstant(s.endDate || '', initEnd)) subChangedFields.push('endDate');
      // v29.52: dasselbe für den Ganztags-Haken je Sub-Event.
      if (!!s.allDay !== !!s.initialAllDay && subChangedFields.indexOf('startDate') < 0) subChangedFields.push('startDate');
      if (!!s.showAsFree !== !!s.initialShowAsFree && subChangedFields.indexOf('startDate') < 0) subChangedFields.push('startDate'); // v29.54
      if (curBodyStripped !== initBodyStripped) subChangedFields.push('outlookBody');
      // v19.20: globale Header-Bild-Layout-Änderung betrifft auch die
      // Sub-Event-Outlook-Termine (gleicher Hero-Bild-Kopf) — als eigenes
      // „layout"-Feld werten, damit das Update-Modal sie mit auflistet.
      if (layoutChanged) subChangedFields.push('layout');
      // v29.38: Der Teams-Link gilt event-weit — er steckt auch in den
      // Sub-Event-Terminen. Ändert er sich, müssen die genauso aktualisiert
      // werden, sonst zeigen sie weiter den alten (oder gar keinen) Link.
      if (onlineMeetingChanged() && subChangedFields.indexOf('outlookBody') < 0) subChangedFields.push('outlookBody');
      if (curSubLogo !== initSubLogo) subChangedFields.push('logo');
      // v30.9: Auch der WIRKSAME Outlook-Betreff zählt als Änderung. Seit
      // v30.7 erben Kalender-Tage den Hauptevent-Titel als Betreff — der
      // Save schreibt das zwar in die OutlookSubject-Spalte, aber ohne
      // UpdateEvent in der Queue bleibt der BESTEHENDE Termin beim alten
      // Namen (Tages-Datum). Verglichen wird gegen den gespeicherten Stand
      // der Sub-Event-Zeile, nicht gegen einen Draft-Schnappschuss — so
      // greift es auch für Events, die vor v30.7 angelegt wurden.
      {
        const storedRow = childEventsOf(editEvent.id).find(c => c.id === s.dbId);
        const effSubSubject = (s.outlookSubject || '').trim() || (subEventCalendar ? (title || editEvent.title || '').trim() : '');
        const storedSubSubject = ((storedRow && storedRow.outlookSubject) || '').trim();
        if (effSubSubject !== storedSubSubject && subChangedFields.indexOf('title') < 0) subChangedFields.push('title');
      }
      // v11.66: Debug-Log für jeden Sub-Event, damit wir in der Browser-
      // Konsole nachvollziehen können, warum das Modal manchmal nicht
      // erscheint. v11.67: JSON.stringify damit der Browser die Werte
      // direkt anzeigt (statt nur „Object" mit Klick zum Aufklappen).
      // v11.79: nur noch sichtbar, wenn der Maintainer in der Console
      // `window.__dexDebug = true` setzt — sonst spammt das Log im
      // Normalbetrieb die DevTools voll.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      if (typeof window !== 'undefined' && (window as any).__dexDebug) {
        // eslint-disable-next-line no-console
        dlog('perf', '[DEX][outlook-detect][sub] ' + JSON.stringify({
          dbId: s.dbId,
          title: s.title,
          subChangedFields,
          disableOutlook: s.disableOutlook,
          hasOutlookEvId,
          initialOutlookEventId: s.initialOutlookEventId,
          initialCalendarLink: s.initialCalendarLink,
          bodyLenInitial: (s.initialOutlookBody || '').length,
          bodyLenCurrent: (s.outlookBody || '').length,
          bodyLenInitStripped: initBodyStripped.length,
          bodyMatch: curBodyStripped === initBodyStripped,
          titleMatch: (s.title || '') === initTitle,
        }));
      }
      if (subChangedFields.length > 0 && !s.disableOutlook) {
        items.push({
          kind: 'sub',
          eventId: s.dbId,
          title: s.title || '',
          changedFields: subChangedFields,
          // v11.68: ohne CalendarLink/OutlookEventId existiert kein Outlook-Termin
          // — Save persistiert den neuen Body in DEX_Events, aber wir können
          // kein UpdateEvent queuen. Modal rendert Info-Eintrag.
          noOutlookYet: !hasOutlookEvId,
        });
      }
    }
    // v14.8: Items aus persistiertem OutlookDirty-Flag nachziehen. Wenn ein
    // Sub-Event oder das Hauptevent in einer früheren Session als „Outlook-
    // Update ausstehend" markiert wurde (User hat den Haken damals nicht
    // gesetzt → OutlookDirty=true wurde in SP geschrieben), soll der nächste
    // Save trotzdem das Modal anbieten — auch ohne neue inhaltliche Änderung
    // in dieser Session. Sonst bleibt der Dirty-Flag ewig hängen und der
    // Yellow-Hint in Schritt 1 wird nie aufgelöst.
    const hasItemForEvent = (id: string): boolean => items.some(it => it.eventId === id);
    // Hauptevent
    // v18.50 BUG-FIX: auch im Dirty-Marker-Pfad das Top-Level-DisableOutlook
    // prüfen (nicht das rohe `disableOutlook`, das beim Speichern auf einem
    // Sub-Event-Tab den Sub-Wert hält) — sonst taucht das Hauptevent im
    // Update-Modal als „Frühere Änderung nicht synchronisiert" auf, obwohl
    // dort Outlook deaktiviert ist (Event mit Outlook nur auf Sub-Event-Ebene).
    // Gleiche Falle wie v18.45 im Changed-Fields-Pfad oben.
    if (editEvent.outlookDirty && !topDisableOutlook && !subEventsOnlyMode
        && (editEvent.outlookEventId || editEvent.calendarLink)
        && !hasItemForEvent(editEvent.id)) {
      items.push({
        kind: 'top',
        eventId: editEvent.id,
        title: title || editEvent.title || '',
        changedFields: [],
      });
    }
    // Sub-Events
    for (const s of subEventsRef.current) {
      if (!s.dbId) continue;
      if (s.disableOutlook) continue;
      const hasOutlookEvId = !!s.initialOutlookEventId || !!s.initialCalendarLink;
      if (!hasOutlookEvId) continue;
      const childEvt = childEventsOf(editEvent.id).find(c => c.id === s.dbId);
      if (childEvt && childEvt.outlookDirty && !hasItemForEvent(s.dbId)) {
        items.push({
          kind: 'sub',
          eventId: s.dbId,
          title: s.title || '',
          changedFields: [],
          noOutlookYet: false,
        });
      }
    }
    // v11.79: gated debug log — siehe oben.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (typeof window !== 'undefined' && (window as any).__dexDebug) {
      // eslint-disable-next-line no-console
      dlog('perf', '[DEX][outlook-detect][result] ' + JSON.stringify({
        itemsCount: items.length,
        items,
        activeCommTabIdx,
        topOutlookBodyLen: (outlookBody || '').length,
        topInitialOutlookBodyLen: (snap.outlookBody || '').length,
        topBodyMatch: currentStripped === initialStripped,
      }));
    }
    return { items };
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
  return (
    <div ref={wizardRootRef} className="page-container" style={{ maxWidth: 1100, marginLeft: 'auto', marginRight: 'auto' }}>
      {showTermsModal && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 720, maxHeight: '90vh', overflow: 'auto',
              padding: 28, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
          >
            <h2 style={{ margin: '0 0 4px', fontSize: '1.3rem' }}>
              {isDe
                ? 'Deloitte Event Experience Platform — Nutzungsbedingungen (Deutschland)'
                : 'Deloitte Event Experience Platform — Terms of Use (Germany)'}
            </h2>
            <p style={{ margin: '0 0 16px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
              {isDe ? 'Letzte Überarbeitung: 05.08.2026' : 'Last revised: 5 August 2026'}
            </p>

            {/* Eingeklappte Kurzfassung — die volle Fassung kann der Nutzer
                über den Toggle ausklappen. Die Checkbox-Bestätigung ist
                trotzdem Pflicht (siehe weiter unten). */}
            <div
              style={{
                background: 'var(--dex-gray-50, #f8f9fa)',
                border: '1px solid var(--dex-gray-200, #e5e7eb)',
                borderRadius: 10,
                padding: '12px 14px',
                fontSize: '0.88rem',
                lineHeight: 1.5,
                color: 'var(--dex-gray-700)',
              }}
            >
              {isDe
                ? <>Bitte gehe sorgfältig mit personenbezogenen Daten der Teilnehmer um, sammle nur das absolut Nötige, nutze die Daten ausschließlich für den vereinbarten Event-Zweck und beachte die Datenschutzregeln von Deloitte Deutschland. Volltext über den Button unten einsehen.</>
                : <>Please handle attendees&apos; personal data with care, collect only what is absolutely necessary, use the data exclusively for the agreed event purpose, and follow Deloitte Germany&apos;s data-protection rules. Use the button below to read the full text.</>}
            </div>

            <button
              type="button"
              onClick={() => setTcExpanded(v => !v)}
              style={{
                marginTop: 10,
                background: 'none',
                border: 'none',
                color: 'var(--dex-green-dark, #4a7c1f)',
                fontWeight: 600,
                fontSize: '0.85rem',
                cursor: 'pointer',
                padding: '4px 0',
                display: 'inline-flex',
                alignItems: 'center',
                gap: 6,
              }}
            >
              {tcExpanded
                ? (isDe ? '▲ Vollständige Bedingungen einklappen' : '▲ Hide full terms')
                : (isDe ? '▼ Vollständige Bedingungen anzeigen' : '▼ Show full terms')}
            </button>

            {tcExpanded && (
              <div style={{ fontSize: '0.88rem', lineHeight: 1.55, color: 'var(--dex-gray-800)', marginTop: 12 }}>
                {isDe ? (
                  <>
                    <p>
                      Der Zugang zur Event Experience Platform wird dir als Mitarbeiter von Deloitte Deutschland gewährt,
                      damit du das Teilnehmermanagement für Veranstaltungen, Events, Workshops oder andere Termine
                      organisieren kannst.
                    </p>

                    <p style={{ marginBottom: 6 }}>Die Plattform dient zur Koordination von:</p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Internen Deloitte Veranstaltungen</li>
                      <li>Externen Veranstaltungen, bei denen das Teilnehmermanagement für Deloitte-Mitarbeiter organisiert wird (bspw. Laufveranstaltungen wie B2Run, oder JPMorgan)</li>
                    </ul>

                    <p>
                      <strong>Wichtiger Hinweis:</strong> Für externe Events mit externen Teilnehmern ist die Plattform
                      nicht vorgesehen. Externe Nicht-Deloitte-Mitarbeiter werden über dieses Tool
                      nicht koordiniert und erhalten keinen Zugang zur Plattform. Alles zu solchen Veranstaltungen findest du im{' '}
                      <a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer">Event Management im DeloitteNet</a>.
                    </p>

                    <p>Jedes Event, das du erstellst, muss den nachfolgenden Richtlinien folgen.</p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Wichtige Datenschutzhinweise</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Die Teilnahme an Events ist immer freiwillig und darf nicht erzwungen werden.</li>
                      <li>Vermeide die Sammlung personenbezogener Daten so weit wie möglich.</li>
                      <li>Sammle nur die Daten, die du unbedingt benötigst, um den Zweck des Events zu erreichen.</li>
                      <li>Reduziere Freitextfelder auf das absolute Minimum, um individuelle Informationen zur Identifizierung von Personen zu vermeiden.</li>
                      <li>Verwende gesammelte Daten ausschließlich für den definierten und genehmigten Zweck. Falls Abweichungen notwendig sind, wende dich im Voraus an das Datenschutz-Team.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Berechtigungen und Datenzugriff</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>Als Event-Ersteller / Administrator:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Du erhältst Admin-Funktionalitäten für dein spezifisches Event.</li>
                      <li>Du kannst auf die gesamte Teilnehmerliste deines Events zugreifen.</li>
                      <li>Diese Berechtigung gilt ausschließlich für das von dir erstellte Event.</li>
                      <li>Du darfst Teilnehmerinformationen nicht mit anderen teilen oder für andere Zwecke verwenden.</li>
                    </ul>

                    <p style={{ marginBottom: 6 }}><strong>Als Event-Teilnehmer:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Du kannst dich für Events an- oder abmelden.</li>
                      <li>Deine Anmeldung ist freiwillig.</li>
                      <li>Du erhältst Informationen zum jeweiligen Event.</li>
                      <li>Du hast keinen Zugriff auf die Teilnehmerliste oder Informationen über andere Teilnehmer.</li>
                      <li>Du siehst nur deine eigenen Event-Anmeldungen und -Daten.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Datenschutzbestimmungen im Detail</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>Beschränkung der Sammlung personenbezogener und vertraulicher Daten:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Nur was unbedingt erforderlich ist, um den beabsichtigten Zweck zu erreichen.</li>
                      <li>Offene Fragen auf das Minimum reduzieren (um die Sammlung unnötiger oder nicht autorisierter Daten zu vermeiden).</li>
                    </ul>

                    <p>
                      <strong>Sammle keine sensiblen personenbezogenen Daten</strong> — das heißt: keine Daten bezüglich
                      Rasse oder ethnischer Herkunft, religiöser oder philosophischer Überzeugungen,
                      Gewerkschaftsmitgliedschaft, politischer Meinungen, medizinischer oder gesundheitlicher Zustände
                      oder Informationen über das Sexualleben oder die sexuelle Orientierung einer Person. Falls sensible
                      personenbezogene Daten gesammelt werden müssen, kontaktiere zuerst das Team unter
                      {' '}<a href="mailto:privacy@deloitte.de">privacy@deloitte.de</a>.
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Besondere Bestimmungen für das Teilnehmermanagement</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Teilnehmerdaten dürfen nur für das spezifische Event verwendet werden, für das sie gesammelt wurden.</li>
                      <li>Die Weitergabe von Teilnehmerlisten an Dritte ist untersagt.</li>
                      <li>Teilnehmerdaten anderer Events sind nicht einsehbar.</li>
                      <li>Nach Abschluss des Events sind Teilnehmerdaten gemäß den Deloitte-Richtlinien zu behandeln.</li>
                    </ul>

                    <p>
                      Ermögliche anonyme Antworten, wann immer möglich. Verwende personenbezogene und vertrauliche Daten,
                      die in einem Event gesammelt wurden, nicht für andere Zwecke als den ursprünglich angegebenen.
                      Sprich dich mit dem Datenschutz-Team ab, falls eine andere Nutzung der Daten beabsichtigt ist
                      (du benötigst die vorherige schriftliche Einwilligung der betroffenen Personen / Teilnehmer
                      unter Verwendung einer entsprechenden Vorlage).
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Kontaktinformationen</h3>
                    <ul style={{ marginTop: 0 }}>
                      {/* v29.43: Funktionspostfach statt persönlichem Konto. */}
                      <li>Kontakt: DEX-Team (<a href="mailto:dex.event@deloitte.de">dex.event@deloitte.de</a>)</li>
                    </ul>

                    <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                      Diese Richtlinien gelten für alle Arten von Events, einschließlich Workshops, Seminare,
                      Webinare, Konferenzen und andere Veranstaltungen, deren Teilnehmermanagement für
                      Deloitte-Mitarbeiter über die Event Experience Platform organisiert wird.
                    </p>
                  </>
                ) : (
                  <>
                    <p>
                      Access to the Event Experience Platform is granted to you as an employee of Deloitte Germany so
                      that you can organise attendee management for events, workshops or other appointments.
                    </p>

                    <p style={{ marginBottom: 6 }}>The platform is used to coordinate:</p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Internal Deloitte events</li>
                      <li>External events for which attendee management is organised on behalf of Deloitte employees (e.g. running events such as B2Run or JPMorgan)</li>
                    </ul>

                    <p>
                      <strong>Important note:</strong> The platform is not intended for external events with external
                      attendees. External non-Deloitte employees are not coordinated through this
                      tool and will not be granted access to the platform. Everything about such events is on{' '}
                      <a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer">Event Management on DeloitteNet</a>.
                    </p>

                    <p>Every event you create must follow the guidelines below.</p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Key data-protection guidance</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Attending events is always voluntary and must never be enforced.</li>
                      <li>Avoid collecting personal data wherever possible.</li>
                      <li>Only collect data that is strictly necessary to achieve the event&apos;s purpose.</li>
                      <li>Keep free-text fields to an absolute minimum to avoid collecting individual information that could identify people.</li>
                      <li>Use collected data exclusively for the defined and approved purpose. If you need to deviate, contact the data-protection team in advance.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Permissions and data access</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>As event creator / administrator:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>You receive admin functionality for your specific event.</li>
                      <li>You can access the entire attendee list of your event.</li>
                      <li>This permission is limited to the event you created.</li>
                      <li>You may not share attendee information with others or use it for other purposes.</li>
                    </ul>

                    <p style={{ marginBottom: 6 }}><strong>As event attendee:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>You can register for or unregister from events.</li>
                      <li>Your registration is voluntary.</li>
                      <li>You receive information about the relevant event.</li>
                      <li>You have no access to the attendee list or information about other attendees.</li>
                      <li>You only see your own event registrations and data.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Data-protection rules in detail</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>Restricting the collection of personal and confidential data:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Only what is strictly necessary to achieve the intended purpose.</li>
                      <li>Reduce open-ended questions to a minimum (to avoid collecting unnecessary or unauthorised data).</li>
                    </ul>

                    <p>
                      <strong>Do not collect sensitive personal data</strong> — that is, no data on race or ethnic origin,
                      religious or philosophical beliefs, trade-union membership, political opinions, medical or health
                      conditions, or information about a person&apos;s sex life or sexual orientation. If sensitive personal
                      data must be collected, contact the team first at
                      {' '}<a href="mailto:privacy@deloitte.de">privacy@deloitte.de</a>.
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Specific rules for attendee management</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Attendee data may only be used for the specific event for which it was collected.</li>
                      <li>Sharing attendee lists with third parties is prohibited.</li>
                      <li>Attendee data of other events is not accessible.</li>
                      <li>After the event, attendee data must be handled in line with Deloitte policy.</li>
                    </ul>

                    <p>
                      Allow anonymous responses wherever possible. Do not use personal or confidential data collected for
                      one event for purposes other than the originally stated one. Coordinate with the data-protection
                      team if you intend to use the data differently (you will need prior written consent from the
                      affected individuals / attendees, using an appropriate template).
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Contact</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Contact: DEX team (<a href="mailto:dex.event@deloitte.de">dex.event@deloitte.de</a>)</li>
                    </ul>

                    <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                      These guidelines apply to all types of events including workshops, seminars, webinars, conferences
                      and any other events whose attendee management for Deloitte employees is organised through the
                      Event Experience Platform.
                    </p>
                  </>
                )}
              </div>
            )}

            <label
              style={{
                display: 'flex', alignItems: 'flex-start', gap: 10,
                marginTop: 20, padding: 14,
                background: tcCheckbox ? 'rgba(134,188,37,0.08)' : 'var(--dex-gray-50, #f8f9fa)',
                border: `1px solid ${tcCheckbox ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200, #e5e7eb)'}`,
                borderRadius: 10, cursor: 'pointer',
                transition: 'all 0.15s ease',
              }}
            >
              <input
                type="checkbox"
                checked={tcCheckbox}
                onChange={e => setTcCheckbox(e.target.checked)}
                style={{ marginTop: 2, width: 18, height: 18, accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
              />
              <span style={{ fontSize: '0.9rem', lineHeight: 1.4 }}>
                {isDe
                  ? 'Ich habe die Nutzungs- und Datenschutzbedingungen gelesen und akzeptiere sie. Ich bestätige, dass ich mich beim Anlegen und Verwalten dieses Events an die Datenschutzbestimmungen halten werde.'
                  : 'I have read and accept the terms of use and data-protection rules. I confirm that I will follow the data-protection rules when creating and managing this event.'}
              </span>
            </label>

            <label
              style={{
                display: 'flex', alignItems: 'flex-start', gap: 10,
                marginTop: 12, padding: 14,
                background: internalCheckbox ? 'rgba(134,188,37,0.08)' : 'var(--dex-gray-50, #f8f9fa)',
                border: `1px solid ${internalCheckbox ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200, #e5e7eb)'}`,
                borderRadius: 10, cursor: 'pointer',
                transition: 'all 0.15s ease',
              }}
            >
              <input
                type="checkbox"
                checked={internalCheckbox}
                onChange={e => setInternalCheckbox(e.target.checked)}
                style={{ marginTop: 2, width: 18, height: 18, accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
              />
              <span style={{ fontSize: '0.9rem', lineHeight: 1.4 }}>
                {isDe
                  ? <>Ich bestätige, dass dies ein <strong>internes Deloitte Event</strong> ist oder die <strong>Deloitte-Teilnahme an einer externen Veranstaltung</strong> koordiniert.
                      <span style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                        Für externe Events mit externen Teilnehmern ist DEX nicht vorgesehen — alles dazu findest du im <a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>Event Management im DeloitteNet</a>.
                      </span></>
                  : <>I confirm that this is a <strong>Deloitte-internal event</strong> or coordinates <strong>Deloitte participation in an external event</strong>.
                      <span style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                        DEX is not intended for external events with external attendees — everything about those is on <a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>Event Management on DeloitteNet</a>.
                      </span></>}
              </span>
            </label>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 12, marginTop: 20 }}>
              <button
                type="button"
                className="btn btn-outline"
                onClick={() => goBack()}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                disabled={!tcCheckbox || !internalCheckbox}
                onClick={() => {
                  setTcAccepted(true);
                  // v29.66: F&A-Pilot — direkt nach dem Akzeptieren fragt der
                  // Dialog nach der Abrechnungsrelevanz (nur Admins, nur beim
                  // Anlegen; im Edit-Modus erscheinen die Bedingungen nicht).
                  if (canBilling) setBillingPromptOpen(true);
                }}
                style={{ opacity: (tcCheckbox && internalCheckbox) ? 1 : 0.5, cursor: (tcCheckbox && internalCheckbox) ? 'pointer' : 'not-allowed' }}
              >
                <Check size={16} /> {isDe ? 'Akzeptieren & weiter' : 'Accept & continue'}
              </button>
            </div>
          </div>
        </div>
      )}
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

      {/* ===== Vollbild-Vorschau Modal ===== */}
      {showPreview && (
        <div className="preview-modal" style={{
          position: 'fixed', inset: 0, background: '#fff', zIndex: 1000,
          display: 'flex', flexDirection: 'column',
        }}>
          <div className="preview-modal-inner" style={{
            background: '#fff', borderRadius: 0, width: '100%', maxWidth: '100%',
            height: '100%', overflow: 'auto', padding: 0,
          }}>
            <div style={{
              padding: '16px 24px', borderBottom: '1px solid var(--dex-gray-200)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              position: 'sticky', top: 0, background: '#fff', zIndex: 1, borderRadius: '16px 16px 0 0',
            }}>
              <div>
                <h3 style={{ margin: 0 }}>Vorschau: Registrierungsseite</h3>
                <p style={{ margin: '4px 0 0', fontSize: '0.8rem', color: 'var(--dex-gray-400)' }}>
                  Sektionen per Drag &amp; Drop verschieben
                </p>
              </div>
              <button onClick={() => setShowPreview(false)} style={{ background: 'none', border: 'none', fontSize: '1.5rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}>
                <X size={24} />
              </button>
            </div>

            <div style={{ padding: 24, display: 'flex', flexDirection: 'column', gap: 16 }}>
              {previewSections.map(section => (
                <div
                  key={section.id}
                  draggable
                  onDragStart={() => setDragSectionId(section.id)}
                  onDragOver={(e) => { e.preventDefault(); setDragOverSectionId(section.id); }}
                  onDragLeave={() => { if (dragOverSectionId === section.id) setDragOverSectionId(null); }}
                  onDrop={() => {
                    if (dragSectionId && dragSectionId !== section.id) {
                      const fromIdx = previewSections.findIndex(s => s.id === dragSectionId);
                      const toIdx = previewSections.findIndex(s => s.id === section.id);
                      if (fromIdx >= 0 && toIdx >= 0) {
                        const updated = [...previewSections];
                        const [moved] = updated.splice(fromIdx, 1);
                        updated.splice(toIdx, 0, moved);
                        setPreviewSections(updated);
                      }
                    }
                    setDragSectionId(null);
                    setDragOverSectionId(null);
                  }}
                  onDragEnd={() => { setDragSectionId(null); setDragOverSectionId(null); }}
                  style={{
                    opacity: dragSectionId === section.id ? 0.4 : 1,
                    borderTop: dragOverSectionId === section.id ? '3px solid var(--dex-green)' : undefined,
                    cursor: 'grab',
                    position: 'relative',
                  }}
                >
                  <div style={{
                    position: 'absolute', top: 4, right: 8, fontSize: '0.65rem',
                    color: 'var(--dex-gray-300)', fontWeight: 600, userSelect: 'none',
                  }}>
                    ⠿ verschieben
                  </div>
                  {renderPreviewSection(section.id)}
                </div>
              ))}
            </div>

            <div style={{
              padding: '16px 24px', borderTop: '1px solid var(--dex-gray-200)',
              display: 'flex', gap: 12, justifyContent: 'flex-end',
              position: 'sticky', bottom: 0, background: '#fff', borderRadius: '0 0 16px 16px',
            }}>
              <button className="btn btn-secondary" onClick={() => setShowPreview(false)}>
                Zurück zum Formular
              </button>
              <button
                className="btn btn-primary"
                disabled={!title}
                onClick={() => { setShowPreview(false); attemptSubmit(); }}
              >
                <Send size={16} /> {isEditMode ? t('create.save') : t('create.submit')}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* HTML-Editor-Modal mit Live-Preview (Outlook-Termin, E-Mail-Template oder Beschreibung).
          v9.39: Mode 'description' für die Event-Beschreibung — wird auf der Anmelde-Seite
          1:1 als HTML gerendert, deshalb hier auch ein Bearbeiten/Vorschau-Modal wie bei den
          Mail-Templates. */}
      {(() => {
        if (!htmlEditorOpen) return null;
        const isOutlook = htmlEditorMode === 'outlook';
        const isDescription = htmlEditorMode === 'description';
        const tType = htmlEditorTemplateType;
        const defaultTpl = (!isOutlook && !isDescription) ? emailTemplates.find(tp => tp.templateType === tType && tp.language === emailLanguage) : undefined;
        const override = (!isOutlook && !isDescription) ? emailTemplateOverrides[tType] : undefined;
        const currentSubject = override?.subject || defaultTpl?.subject || '';
        const currentHeading = override?.heading || defaultTpl?.heading || '';
        // v15.19: Subheading-Override pro Event. Falls override.subheading
        // explizit gesetzt ist (auch leerer String), nutze diesen Wert.
        const currentSubheading = override?.subheading !== undefined ? override.subheading : '';
        // v28.89: Im Beschreibungs-Modus folgt der Editor der gewählten Ebene
        // (scDescription/setScDescription) — er wird ausschließlich aus
        // Schritt 1 geöffnet, wo der Scope-Reiter darüber steht.
        const currentBody = isOutlook
          ? outlookBody
          : isDescription
            ? scDescription
            : (override?.bodyHtml || defaultTpl?.bodyHtml || '');
        // v18.19: Überschrift-Farbe + -Größe (Override > Template-Default).
        const currentHeadingColor = (override?.headingColor) || (defaultTpl?.headingColor) || '#86bc25';
        const currentHeadingFontSize = override?.headingFontSize || '26px';
        // v18.22: Überschrift fett/kursiv + Unter-Überschrift-Formatierung.
        const currentHeadingBold = override?.headingBold;
        const currentHeadingItalic = override?.headingItalic;
        const currentSubheadingColor = override?.subheadingColor || '#000000';
        const currentSubheadingFontSize = override?.subheadingFontSize || '20px';
        const currentSubheadingBold = override?.subheadingBold;
        const currentSubheadingItalic = override?.subheadingItalic;
        // v18.22: zentraler Patch-Helper — merged ein Teil-Update in den
        // Override des aktuellen TemplateTypes und BEWAHRT alle übrigen Felder
        // (vorher droppte z.B. ein Heading-Text-Edit die zuvor gesetzte Farbe).
        const patchOverride = (patch: Partial<EmailOverrideEntry>): void => {
          setEmailTemplateOverrides(prev => {
            const cur = prev[tType];
            return {
              ...prev,
              [tType]: {
                subject: cur?.subject ?? currentSubject,
                heading: cur?.heading ?? currentHeading,
                subheading: cur?.subheading !== undefined ? cur.subheading : currentSubheading,
                bodyHtml: cur?.bodyHtml ?? currentBody,
                ...(cur ? {
                  headingColor: cur.headingColor,
                  headingFontSize: cur.headingFontSize,
                  headingBold: cur.headingBold,
                  headingItalic: cur.headingItalic,
                  subheadingColor: cur.subheadingColor,
                  subheadingFontSize: cur.subheadingFontSize,
                  subheadingBold: cur.subheadingBold,
                  subheadingItalic: cur.subheadingItalic,
                } : {}),
                ...patch,
              },
            };
          });
        };
        // v18.42: read-only Termin/Ort-Labels für den Outlook-Editor — je nach
        // aktivem Tab (Hauptevent oder Sub-Event).
        const olActiveSub = activeCommTabIdx > 0 ? subEvents[activeCommTabIdx - 1] : undefined;
        const olStart = olActiveSub ? olActiveSub.startDate : startDate;
        const olEnd = olActiveSub ? olActiveSub.endDate : endDate;
        const olFmt = (d?: string): string => {
          if (!d) return '';
          try { return new Date(d).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); } catch { return ''; }
        };
        // v18.44: Auto-Ort (= „würde übernommen") als Platzhalter.
        const outlookLocationAuto = olActiveSub
          ? (buildOutlookLocation(olActiveSub.location, olActiveSub.locationAddress) || olActiveSub.location || '')
          : (buildOutlookLocation(location, { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity }));
        // v18.44: aktuelle Override-Werte des aktiven Tabs (leer = übernommen).
        const olLocationOverrideVal = olActiveSub ? (olActiveSub.outlookLocation || '') : outlookLocationOverride;
        const olStartOverrideVal = olActiveSub ? (olActiveSub.outlookStart || '') : outlookStartOverride;
        const olEndOverrideVal = olActiveSub ? (olActiveSub.outlookEnd || '') : outlookEndOverride;
        const setOlLocation = (v: string): void => {
          if (olActiveSub) { const fi = activeCommTabIdx - 1; setSubEvents(prev => prev.map((s, i) => i === fi ? { ...s, outlookLocation: v } : s)); }
          else setOutlookLocationOverride(v);
        };
        const setOlStart = (iso: string): void => {
          if (olActiveSub) { const fi = activeCommTabIdx - 1; setSubEvents(prev => prev.map((s, i) => i === fi ? { ...s, outlookStart: iso } : s)); }
          else setOutlookStartOverride(iso);
        };
        const setOlEnd = (iso: string): void => {
          if (olActiveSub) { const fi = activeCommTabIdx - 1; setSubEvents(prev => prev.map((s, i) => i === fi ? { ...s, outlookEnd: iso } : s)); }
          else setOutlookEndOverride(iso);
        };
        const pad2 = (n: number): string => String(n).padStart(2, '0');
        const olIsoToDate = (iso?: string): Date | null => {
          if (!iso) return null;
          const loc = isoToLocal(iso); if (!loc) return null;
          const [dp, tp] = loc.split('T'); const [y, mo, da] = dp.split('-').map(n => parseInt(n, 10)); const [h, mi] = (tp || '00:00').split(':').map(n => parseInt(n, 10));
          return new Date(y, mo - 1, da, h, mi, 0, 0);
        };
        const olDateToIso = (d: Date | null): string => {
          if (!d) return '';
          return berlinLocalToUtcIso(`${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}`);
        };
        const dpCommon = {
          showTimeSelect: true, timeFormat: 'HH:mm', timeIntervals: 15, timeCaption: 'Uhrzeit',
          dateFormat: 'dd.MM.yyyy, HH:mm', locale: 'de', className: 'form-input',
          wrapperClassName: 'dex-datepicker-wrapper', calendarClassName: 'dex-datepicker-calendar',
          popperPlacement: 'bottom-start' as const, isClearable: true, autoComplete: 'off',
        };
        const outlookDateEditor = (
          <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
              <label style={{ fontSize: '0.68rem', color: 'var(--dex-gray-400)' }}>Start</label>
              <DatePicker {...dpCommon} selected={olIsoToDate(olStartOverrideVal)} onChange={(d: Date | null) => setOlStart(olDateToIso(d))} placeholderText={olStart ? olFmt(olStart) + ' (übernommen)' : 'Start'} />
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
              <label style={{ fontSize: '0.68rem', color: 'var(--dex-gray-400)' }}>Ende</label>
              <DatePicker {...dpCommon} selected={olIsoToDate(olEndOverrideVal)} onChange={(d: Date | null) => setOlEnd(olDateToIso(d))} placeholderText={olEnd ? olFmt(olEnd) + ' (übernommen)' : 'Ende'} />
            </div>
          </div>
        );
        // v18.46: Standard-Body-Vorlage (mit Platzhaltern) für „Standardtext laden"
        // im Outlook-Editor — Sprache folgt der aktiven Mail-Sprache.
        const outlookDefaultBody = (emailLanguage === 'EN')
          ? '<p>You are registered for the event <strong>{{EventTitle}}</strong>.</p>'
            + '<p>If you are unable to attend, please cancel your registration in time via the <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView" style="color:#86bc25;font-weight:600;">DEX App</a> („My Events").</p>'
            + '<p>For organizational questions please contact <strong>{{Organizer}}</strong>.</p>'
          : '<p>Ihr seid für das Event <strong>{{EventTitle}}</strong> angemeldet.</p>'
            + '<p>Falls ihr nicht teilnehmen könnt, meldet euch bitte rechtzeitig über die <a href="https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView" style="color:#86bc25;font-weight:600;">DEX App</a> („Meine Events") ab.</p>'
            + '<p>Bei organisatorischen Fragen wendet euch bitte an <strong>{{Organizer}}</strong>.</p>';
        // v19.2: Einladender Beispieltext für die Beschreibung — über den
        // „Standardtext laden"-Button im Beschreibungs-Editor übernehmbar (statt
        // wie früher als Inline-Box im Wizard).
        const descriptionExampleHtml = isDe
          ? 'Liebe Kolleginnen und Kollegen,<br><br>wir freuen uns sehr, euch herzlich einzuladen! Es erwartet euch ein abwechslungsreiches Programm mit viel Raum für Austausch und Begegnung.<br><br>Wir freuen uns auf einen schönen gemeinsamen Tag mit euch!'
          : 'Dear colleagues,<br><br>we are delighted to invite you! Look forward to a varied programme with plenty of room for exchange and networking.<br><br>We look forward to seeing you there!';
        return (
          <HtmlEditorModal
            open={htmlEditorOpen}
            onClose={() => setHtmlEditorOpen(false)}
            defaultBodyHtml={isOutlook ? outlookDefaultBody : (isDescription ? descriptionExampleHtml : undefined)}
            title={isOutlook
              ? 'Outlook-Termin: Body bearbeiten'
              : isDescription
                ? (scopeSub
                  ? (isDe
                    ? `Beschreibung: ${shortSubEventTitle(scopeSub.title, title) || (childTermSingular || 'Sub-Event')}`
                    : `Description: ${shortSubEventTitle(scopeSub.title, title) || (childTermSingular || 'sub-event')}`)
                  : (isDe ? 'Event-Beschreibung bearbeiten' : 'Edit event description'))
                : `E-Mail-Template: ${tType}`}
            // v28.7: Die Starthilfe (Tipp-Text + Vorschlags-Chips) lebt jetzt
            // HIER im Editor statt als Dauer-Box im Wizard-Schritt.
            headerExtra={isDescription ? (
              <div style={{ padding: '10px 12px', borderRadius: 'var(--dex-radius)', background: 'var(--dex-gray-50, #f7f7f7)', border: '1px solid var(--dex-gray-200)', fontSize: '0.78rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {isDe
                  ? <>Die Beschreibung ist der <strong>einladende Einleitungstext ganz oben auf der Anmeldemaske</strong> — das Erste, was deine Teilnehmenden lesen. Erzähl hier gern, <strong>worum es geht, für wen das Event ist und was man wissen sollte</strong>.<br />Ein kleiner Tipp: <strong>Zeitpunkt, Ort, Organizer und Kontaktperson musst du hier nicht angeben</strong> — die zeigt die App bereits als eigene Felder darüber an. So bleibt dein Text schön schlank und einladend.</>
                  : <>The description is the <strong>inviting intro text right at the top of the registration form</strong> — the first thing your attendees read. Feel free to tell them <strong>what the event is about, who it&rsquo;s for and what to know</strong>.<br />A little tip: <strong>you don&rsquo;t need to add the date, location, organizer or contact person here</strong> — the app already shows those as their own fields above. That keeps your text nice and inviting.</>}
              </div>
            ) : undefined}
            bodyTemplates={isDescription ? DESCRIPTION_TEMPLATES.map(tpl => ({
              key: tpl.key,
              label: isDe ? tpl.labelDe : tpl.labelEn,
              html: isDe ? tpl.de : tpl.en,
              title: (isDe ? tpl.de : tpl.en).replace(/<[^>]+>/g, '').replace(/&rsquo;/g, '’'),
            })) : undefined}
            bodyTemplatesLabel={isDescription ? (isDe ? 'Vorschläge zum Übernehmen (danach frei anpassbar):' : 'Suggestions to use (fully editable afterwards):') : undefined}
            value={currentBody}
            onChange={(html) => {
              if (isOutlook) {
                setOutlookBody(html);
              } else if (isDescription) {
                setScDescription(html);
              } else {
                // v18.22: patchOverride bewahrt alle übrigen Override-Felder
                // (Farbe/Größe/fett/kursiv von Über-/Unter-Überschrift).
                patchOverride({ bodyHtml: html });
              }
            }}
            previewMode={isDescription ? 'plain' : (isOutlook ? 'outlook' : 'email')}
            emailSubject={(!isOutlook && !isDescription) ? currentSubject : undefined}
            onEmailSubjectChange={(!isOutlook && !isDescription) ? (s) => patchOverride({ subject: s }) : undefined}
            emailHeading={(!isOutlook && !isDescription) ? currentHeading : undefined}
            onEmailHeadingChange={(!isOutlook && !isDescription) ? (h) => patchOverride({ heading: h }) : undefined}
            emailSubheading={(!isOutlook && !isDescription) ? currentSubheading : undefined}
            onEmailSubheadingChange={(!isOutlook && !isDescription) ? (s) => patchOverride({ subheading: s }) : undefined}
            emailHeadingColor={(!isOutlook && !isDescription) ? currentHeadingColor : undefined}
            emailHeadingFontSize={(!isOutlook && !isDescription) ? currentHeadingFontSize : undefined}
            onEmailHeadingColorChange={(!isOutlook && !isDescription) ? (hex) => patchOverride({ headingColor: hex }) : undefined}
            onEmailHeadingFontSizeChange={(!isOutlook && !isDescription) ? (px) => patchOverride({ headingFontSize: px }) : undefined}
            emailHeadingBold={(!isOutlook && !isDescription) ? currentHeadingBold : undefined}
            emailHeadingItalic={(!isOutlook && !isDescription) ? currentHeadingItalic : undefined}
            onEmailHeadingBoldChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ headingBold: b }) : undefined}
            onEmailHeadingItalicChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ headingItalic: b }) : undefined}
            emailSubheadingColor={(!isOutlook && !isDescription) ? currentSubheadingColor : undefined}
            emailSubheadingFontSize={(!isOutlook && !isDescription) ? currentSubheadingFontSize : undefined}
            emailSubheadingBold={(!isOutlook && !isDescription) ? currentSubheadingBold : undefined}
            emailSubheadingItalic={(!isOutlook && !isDescription) ? currentSubheadingItalic : undefined}
            onEmailSubheadingColorChange={(!isOutlook && !isDescription) ? (hex) => patchOverride({ subheadingColor: hex }) : undefined}
            onEmailSubheadingFontSizeChange={(!isOutlook && !isDescription) ? (px) => patchOverride({ subheadingFontSize: px }) : undefined}
            onEmailSubheadingBoldChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ subheadingBold: b }) : undefined}
            onEmailSubheadingItalicChange={(!isOutlook && !isDescription) ? (b) => patchOverride({ subheadingItalic: b }) : undefined}
            imageWidth={!isDescription ? headerImageLayout.width : undefined}
            imagePaddingV={!isDescription ? headerImageLayout.paddingV : undefined}
            imagePaddingH={!isDescription ? headerImageLayout.paddingH : undefined}
            onImageWidthChange={!isDescription ? (w) => setHeaderImageLayout(p => ({ ...p, width: w })) : undefined}
            onImagePaddingVChange={!isDescription ? (v) => setHeaderImageLayout(p => ({ ...p, paddingV: v })) : undefined}
            onImagePaddingHChange={!isDescription ? (h) => setHeaderImageLayout(p => ({ ...p, paddingH: h })) : undefined}
            outlookHeading={isOutlook ? outlookHeading : undefined}
            onOutlookHeadingChange={isOutlook ? setOutlookHeading : undefined}
            outlookSubheading={isOutlook ? outlookSubheading : undefined}
            onOutlookSubheadingChange={isOutlook ? setOutlookSubheading : undefined}
            outlookSubject={isOutlook ? outlookSubject : undefined}
            onOutlookSubjectChange={isOutlook ? setOutlookSubject : undefined}
            outlookDateEditor={isOutlook ? outlookDateEditor : undefined}
            outlookLocationValue={isOutlook ? olLocationOverrideVal : undefined}
            onOutlookLocationChange={isOutlook ? setOlLocation : undefined}
            outlookLocationAuto={isOutlook ? outlookLocationAuto : undefined}
            previewVars={{
              // v17.5: Im Sub-Event-Kommunikations-Tab den Titel des
              // aktiven Sub-Events einsetzen, sonst den Hauptevent-Titel.
              EventTitle: (() => {
                if (activeCommTabIdx > 0) {
                  const sub = subEvents[activeCommTabIdx - 1];
                  return (sub && sub.title && sub.title.trim()) || title || 'Event Title';
                }
                return title || 'Event Title';
              })(),
              Name: 'Max Mustermann',
              // v27.5: normalisierte Organizer-Namen ("Vorname Nachname" + „und").
              Organizer: formatOrganizerList([organizer], emailLanguage) || organizer || 'Organisator',
              ContactEmail: contactEmail.trim() || 'kontakt@deloitte.de',
              AppUrl: 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView',
              WaitlistPosition: '1',
              Address: [addrStreet, addrHouseNo].filter(Boolean).join(' ') + ((addrZip || addrCity) ? ', ' + [addrZip, addrCity].filter(Boolean).join(' ') : ''),
              Location: location || 'Veranstaltungsort',
              StartDate: startDate ? new Date(startDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
              EndDate: endDate ? new Date(endDate).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
              EventDate: startDate ? new Date(startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
            }}
            insertableVars={isOutlook ? [
              // v17.16: {{Name}} hier ENTFERNT — der Outlook-Termin geht
              // an alle Teilnehmer gleichzeitig, eine pro-Person-Anrede
              // ist nicht möglich. Vorher konnte der Organizer {{Name}}
              // einfügen, was bei allen Empfängern als unaufgelöster
              // Platzhalter „{{Name}}" stehen blieb.
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
              // v27.5: {{ContactEmail}} nur anbieten, wenn eine Ansprechpartner-
              // Mail hinterlegt ist (Schritt „Ansprechpartner").
              ...(contactEmail.trim() ? [{ key: '{{ContactEmail}}', label: 'Kontakt-Mail' }] : []),
              { key: '{{Location}}', label: 'Ort' },
              { key: '{{Address}}', label: 'Adresse' },
              { key: '{{StartDate}}', label: 'Start' },
              { key: '{{EndDate}}', label: 'Ende' },
              { key: '{{AppUrl}}', label: 'App Link' },
            ] : [
              { key: '{{Name}}', label: 'Name' },
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
              ...(contactEmail.trim() ? [{ key: '{{ContactEmail}}', label: 'Kontakt-Mail' }] : []),
              { key: '{{AppUrl}}', label: 'App Link' },
              { key: '{{WaitlistPosition}}', label: 'Waitlist #' },
            ]}
            imageBase64={(isOutlook ? outlookLogoPreview : emailLogoPreview) || ''}
          />
        );
      })()}

      {/* Register-Page-Preview-Modal (zeigt, was Teilnehmer sehen würden) */}
      <RegisterPreviewModal
        open={showRegisterPreview}
        onClose={() => setShowRegisterPreview(false)}
        data={{
          title,
          description,
          location,
          locationAddress: { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity },
          startDate,
          endDate,
          imagePreview,
          organizers: organizer.split(';').map(s => s.trim()).filter(Boolean),
          organizerEmails,
          maxParticipants: Number(maxParticipants) || 0,
          unlimitedParticipants,
          customFields: customFields.map(f => ({
            id: f.id,
            label: f.label,
            type: f.type,
            required: f.required,
            visible: f.visible !== false,
            options: f.type === 'select' ? f.options : undefined,
            // v26.74: Vorauswahl an die Live-Preview weiterreichen.
            defaultValue: f.type === 'select' && !f.multi ? f.defaultValue : undefined,
            // v26.75: Vorfilter-Kategorien + Beschriftung an die Preview.
            optionCategories: f.type === 'select' && !f.multi ? f.optionCategories : undefined,
            prefilterLabel: f.type === 'select' && !f.multi ? f.prefilterLabel : undefined,
            // v7.24: helpText, multi und showIf an die Live-Preview weiterreichen,
            // damit die echte RegistrationPage genau das rendert was der
            // Teilnehmer später sieht (i-Tooltip, Multi-Select-Liste,
            // Sichtbarkeitsbedingung).
            helpText: f.helpText,
            helpTextStyle: f.helpTextStyle,
            multi: f.multi,
            showIf: f.showIf,
            // v17.20: EN-Varianten an die Preview weiterreichen — sonst sieht
            // der Organizer in der Vorschau nicht, was englische Teilnehmer
            // bekommen würden.
            confirmLabel: f.confirmLabel,
            labelEn: f.labelEn,
            helpTextEn: f.helpTextEn,
            confirmLabelEn: f.confirmLabelEn,
            optionsEn: f.optionsEn,
            // v29.21 (Audit): Uhrzeit-Option + Übernachtungs-Fenster an die
            // Vorschau — sonst zeigte sie einen reinen Datums-Picker bzw.
            // einen Zeitraum ohne Grenzen, anders als die echte Anmeldeseite.
            withTime: f.withTime,
            rangeStart: f.rangeStart,
            rangeEnd: f.rangeEnd,
            maxNights: f.maxNights,
          })),
          isFictive,
          // v14.10: Sub-Events + Sub-Only-Mode + Bezeichnungs-Term an die
          // Vorschau weiterreichen, damit der Organizer auch die Sub-Event-
          // Auswahl im Anmeldeformular sieht (vorher fehlte sie komplett).
          subEvents: subEvents.map(s => ({
            id: s.id,
            title: s.title,
            location: s.location,
            startDate: s.startDate,
            endDate: s.endDate,
            maxParticipants: s.maxParticipants,
            description: s.description,
            customFields: (s.customFields || []).map(f => ({
              id: f.id,
              label: f.label,
              type: f.type,
              required: f.required,
              visible: f.visible !== false,
              options: f.type === 'select' ? f.options : undefined,
              helpText: f.helpText,
              helpTextStyle: f.helpTextStyle,
              multi: f.multi,
              showIf: f.showIf,
              // v29.21 (Audit): wie beim Hauptevent — Vorauswahl, EN-Varianten
              // und Datums-Optionen fehlten in der Sub-Event-Vorschau.
              defaultValue: f.type === 'select' && !f.multi ? f.defaultValue : undefined,
              confirmLabel: f.confirmLabel,
              labelEn: f.labelEn,
              helpTextEn: f.helpTextEn,
              confirmLabelEn: f.confirmLabelEn,
              optionsEn: f.optionsEn,
              withTime: f.withTime,
              rangeStart: f.rangeStart,
              rangeEnd: f.rangeEnd,
              maxNights: f.maxNights,
            })),
          })),
          subEventsOnlyMode,
          requireSubEventSelection: requireSubEventSelection || subEventsOnlyMode,
          childEventTermSingular: childTermSingular,
          childEventTermPlural: childTermPlural,
          // v17.22: Bilingual-Flag an die Vorschau — sonst rendert die
          // Preview die EN-Varianten nie (useEnVariants prüft event.bilingualFields).
          bilingualFields,
          // v22.36: Geteilte Kapazität an die Vorschau — sonst fehlt die
          // Gruppenauswahl im Vorschau-Formular.
          ...(useSplitCapacities ? {
            durchstarterCapacity: Number(durchstarterCapacity) || 0,
            funstarterCapacity: Number(funstarterCapacity) || 0,
            splitLabelA,
            splitLabelB,
            splitDescA,
            splitDescB,
            splitHelpText,
            splitSectionTitle,
            splitDisplayOrderReversed,
            splitSharedWaitlist,
          } : {}),
          // v29.21 (Audit): Die Vorschau verspricht „1:1 das, was der
          // Teilnehmer bekommt" — ohne diese Props fehlten Anrede-Dropdown,
          // Gruppen-Beschreibungen und die gedrehte Gruppen-Reihenfolge.
          askSalutation,
        }}
      />

      {/* Massenimport-Modale — eine generische Komponente, mehrere Aufruf-Stellen.
          Teams speichern parallele Names[] + Emails[]-Arrays. Die onAdd-Callbacks
          übersetzen jeweils zwischen Modal-Output (Email + DisplayName) und der
          jeweiligen State-Form. Der Audience-/Sichtbarkeits-Massenimport ist nach
          <AudiencePicker> gewandert (self-contained pro Instanz). */}
      <BulkUserImportModal
        open={bulkOrganizerOpen}
        onClose={() => setBulkOrganizerOpen(false)}
        title="Massenimport — Co-Organizer"
        description={(
          <p style={{ marginTop: 0 }}>
            Mehrere <strong>Co-Organizer</strong> auf einmal hinzufügen. Reihenfolge
            spielt eine Rolle — der erste Eintrag in der Liste bleibt der Haupt-Organizer.
            Massenimport hängt neue Personen <strong>hinten</strong> an.
          </p>
        )}
        existingEmails={organizerEmails}
        searchUsers={searchUsers}
        onAdd={({ email, displayName }) => {
          // WICHTIG: functional setState für `organizer`-String, sonst sehen
          // schnelle Sequenz-Calls (Massenimport mit 10+ Namen) alle dieselbe
          // closure-stale Version und nur der letzte Name landet, während
          // organizerEmails über `prev => ...` korrekt akkumuliert. Das führte
          // zu out-of-sync orgNames/orgEmails-Arrays mit falscher Namen-Email-
          // Zuordnung im Duplikat-Hinweis.
          setOrganizer(prev => {
            const existingNames = (prev || '').split(';').map(s => s.trim()).filter(Boolean);
            return [...existingNames, displayName].join('; ');
          });
          setOrganizerEmails(prev => [...prev, email]);
        }}
      />
      <BulkUserImportModal
        open={bulkTestTeamOpen}
        onClose={() => setBulkTestTeamOpen(false)}
        title="Massenimport — Test-Team"
        description={(
          <p style={{ marginTop: 0 }}>
            Mehrere <strong>Test-Team-Mitglieder</strong> auf einmal hinzufügen. Test-Team
            sieht das Event schon im Entwurfsmodus und kann sich testweise anmelden.
          </p>
        )}
        existingEmails={testTeamEmails}
        searchUsers={searchUsers}
        onAdd={({ email, displayName }) => {
          setTestTeamNames(prev => [...prev, displayName]);
          setTestTeamEmails(prev => [...prev, email]);
        }}
      />
      <BulkUserImportModal
        open={bulkQrScannerOpen}
        onClose={() => setBulkQrScannerOpen(false)}
        title="Massenimport — Check-In Team"
        description={(
          <p style={{ marginTop: 0 }}>
            Mehrere <strong>Check-In-Team-Mitglieder</strong> auf einmal hinzufügen. Diese
            Personen dürfen am Eventtag den QR-Scanner / Check-In-Tool benutzen, haben aber
            keine weiteren Admin-Rechte.
          </p>
        )}
        existingEmails={qrScannerEmails}
        searchUsers={searchUsers}
        onAdd={({ email, displayName }) => {
          setQrScannerNames(prev => [...prev, displayName]);
          setQrScannerEmails(prev => [...prev, email]);
        }}
      />

      {/* Das Gruppen-Mitglieder-Modal ist nach <AudiencePicker> gewandert
          (die Audience-Chips, die es öffnen, leben jetzt dort). */}

      {/* v11.88: Demo-Auswahl-Modal — Ersatz für den früheren
          direkten „Demo"-Klick. Vier Karten-Optionen: Standard,
          Mit Gruppen, Mit Sub-Event, Mit Sub-Event + Team. Klick auf
          eine Karte schliesst das Modal und füllt das Formular mit
          der jeweiligen Variante. */}

      {/* v22.36: „Prüfen"-Modal — Übersicht aller Einstellungen: was ist
          gesetzt, wo greifen Standards, welche optionalen Punkte sind leer,
          welche Pflichtangaben fehlen. Rein lesend aus dem Wizard-State. */}
      {showConfigCheck && (() => {
        type CheckStatus = 'ok' | 'default' | 'empty' | 'missing';
        interface CheckRow { label: string; value: React.ReactNode; status: CheckStatus }
        const fmtDt = (v: string): string => {
          if (!v) return '';
          const d = new Date(v);
          return Number.isFinite(d.getTime()) ? d.toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : v;
        };
        const plainDesc = (description || '').replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').trim();
        const orgList = organizer.split(';').map(s => s.trim()).filter(Boolean);
        const locList = locationFilter.split(',').map(s => s.trim()).filter(Boolean);
        const audList = (audience || '').split(',').map(s => s.trim()).filter(Boolean);
        const sections: Array<{ title: string; rows: CheckRow[] }> = [];
        sections.push({
          title: isDe ? 'Schritt 1 — Grundlagen' : 'Step 1 — Basics',
          rows: [
            { label: isDe ? 'Event-Titel' : 'Event title', value: title || '—', status: title ? 'ok' : 'missing' },
            { label: isDe ? 'Zeitraum' : 'Dates', value: startDate ? `${fmtDt(startDate)}${endDate ? ` – ${fmtDt(endDate)}` : (isDe ? ' (kein Ende — Outlook-Termin nicht möglich)' : ' (no end — Outlook invite not possible)')}` : '—', status: !startDate ? 'missing' : (endDate ? 'ok' : 'empty') },
            { label: isDe ? 'Beschreibung' : 'Description', value: plainDesc ? `${plainDesc.slice(0, 80)}${plainDesc.length > 80 ? '…' : ''}` : '—', status: plainDesc ? 'ok' : 'empty' },
            { label: isDe ? 'Event-Bild' : 'Event image', value: imagePreview ? (isDe ? 'hochgeladen' : 'uploaded') : '—', status: imagePreview ? 'ok' : 'empty' },
            { label: 'Status', value: isFictive ? (activeFrom ? (isDe ? `Entwurf — geht automatisch live am ${fmtDt(activeFrom)}` : `Draft — goes live automatically on ${fmtDt(activeFrom)}`) : (isDe ? 'Entwurf (nur Admins, Organizer, Test-Team)' : 'Draft (admins, organizers, test team only)')) : (isDe ? 'Aktiv — für berechtigte Teilnehmer sichtbar' : 'Active — visible to eligible attendees'), status: isFictive ? 'default' : 'ok' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 2 — Organizer & Team' : 'Step 2 — Organizers & Team',
          rows: [
            { label: 'Organizer', value: orgList.length ? `${orgList.length} ${isDe ? 'Person(en)' : 'person(s)'}` : '—', status: orgList.length ? 'ok' : 'missing' },
            { label: isDe ? 'Test-Team' : 'Test team', value: testTeamEmails.length ? `${testTeamEmails.length} ${isDe ? 'Person(en)' : 'person(s)'}` : '—', status: testTeamEmails.length ? 'ok' : 'empty' },
            { label: isDe ? 'Check-In-Team' : 'Check-in team', value: qrScannerEmails.length ? `${qrScannerEmails.length} ${isDe ? 'Person(en)' : 'person(s)'}` : '—', status: qrScannerEmails.length ? 'ok' : 'empty' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 1 — Sub-Events' : 'Step 1 — Sub-events',
          rows: subEvents.length === 0
            ? [{ label: 'Sub-Events', value: isDe ? 'keine' : 'none', status: 'default' }]
            : [
                { label: 'Sub-Events', value: `${subEvents.length} (${subEvents.map(s => s.title || '?').join(', ').slice(0, 90)})`, status: 'ok' },
                { label: isDe ? 'Anmelde-Modus' : 'Registration mode', value: subEventsOnlyMode ? (isDe ? 'Nur für Sub-Events (Klammer nicht buchbar)' : 'Sub-events only (bracket not bookable)') : (isDe ? 'Hauptevent + Sub-Events' : 'Main event + sub-events'), status: 'ok' },
              ],
        });
        sections.push({
          title: isDe ? 'Schritt 3 — Ort & Programm' : 'Step 3 — Location & programme',
          rows: [
            { label: isDe ? 'Veranstaltungsort' : 'Venue', value: location || '—', status: location ? 'ok' : 'empty' },
            { label: isDe ? 'Adresse' : 'Address', value: (addrStreet || addrCity) ? [addrStreet, addrHouseNo, addrZip, addrCity].filter(Boolean).join(' ') : '—', status: (addrStreet || addrCity) ? 'ok' : 'empty' },
            { label: 'Agenda', value: agenda.length ? `${agenda.length} ${isDe ? 'Programmpunkte' : 'items'}` : '—', status: agenda.length ? 'ok' : 'empty' },
            { label: isDe ? 'Transferzeiten' : 'Transfers', value: transferTimes.length ? `${transferTimes.length}` : '—', status: transferTimes.length ? 'ok' : 'empty' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 4 — Kapazität & Sichtbarkeit' : 'Step 4 — Capacity & visibility',
          rows: [
            useSplitCapacities
              ? { label: isDe ? 'Plätze (geteilte Kapazität)' : 'Seats (split capacity)', value: `${splitLabelA || 'Gruppe A'}: ${durchstarterCapacity || 0} · ${splitLabelB || 'Gruppe B'}: ${funstarterCapacity || 0}${splitSharedWaitlist ? (isDe ? ' · gemeinsame Warteliste' : ' · shared waitlist') : ''}`, status: 'ok' }
              : { label: isDe ? 'Plätze' : 'Seats', value: unlimitedParticipants ? (isDe ? 'Unbegrenzt' : 'Unlimited') : String(maxParticipants || 0), status: unlimitedParticipants ? 'default' : 'ok' },
            { label: isDe ? 'Warteliste' : 'Waitlist', value: waitlistEnabled ? (isDe ? 'aktiv' : 'on') : (isDe ? 'aus' : 'off'), status: waitlistEnabled ? 'default' : 'ok' },
            { label: isDe ? 'Anmeldefrist' : 'Registration deadline', value: registrationDeadline ? fmtDt(registrationDeadline) : '—', status: registrationDeadline ? 'ok' : 'empty' },
            { label: isDe ? 'Abmeldefrist (kommuniziert)' : 'Cancellation deadline (communicated)', value: lastDeregisterDate ? fmtDt(lastDeregisterDate) : '—', status: lastDeregisterDate ? 'ok' : 'empty' },
            { label: isDe ? 'Sichtbarkeit' : 'Visibility', value: (locList.length === 0 && audList.length === 0) ? (isDe ? 'alle Mitarbeiter von Deloitte Deutschland' : 'all Deloitte Germany employees') : [locList.length ? `${isDe ? 'Standorte' : 'Locations'}: ${locList.join(', ')}` : '', audList.length ? `${audList.length} ${isDe ? 'Verteiler/Personen' : 'lists/people'}` : ''].filter(Boolean).join(filterMode === 'AND' ? ' UND ' : ' ODER '), status: (locList.length === 0 && audList.length === 0) ? 'default' : 'ok' },
            { label: isDe ? 'Ausgeschlossene Personen' : 'Excluded people', value: excludedUsers.length ? `${excludedUsers.length}` : '—', status: excludedUsers.length ? 'ok' : 'default' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 5 — Felder' : 'Step 5 — Fields',
          rows: [
            { label: isDe ? 'Eigene Abfrage-Felder' : 'Custom fields', value: customFields.length ? `${customFields.length}` : (isDe ? 'keine' : 'none'), status: customFields.length ? 'ok' : 'default' },
            { label: isDe ? 'Anrede abfragen' : 'Ask salutation', value: askSalutation ? (isDe ? 'an' : 'on') : (isDe ? 'aus' : 'off'), status: askSalutation ? 'ok' : 'default' },
            { label: isDe ? 'Zweisprachige Felder (DE+EN)' : 'Bilingual fields (DE+EN)', value: bilingualFields ? (isDe ? 'an' : 'on') : (isDe ? 'aus' : 'off'), status: bilingualFields ? 'ok' : 'default' },
            { label: isDe ? 'Formular-Sprache' : 'Form language', value: registrationLanguage === 'de' ? (isDe ? 'Immer Deutsch' : 'Always German') : registrationLanguage === 'en' ? (isDe ? 'Immer Englisch' : 'Always English') : (isDe ? 'Automatisch (App-Sprache)' : 'Automatic (app language)'), status: registrationLanguage ? 'ok' : 'default' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 6 — Kommunikation' : 'Step 6 — Communication',
          rows: [
            { label: isDe ? 'Mail-Sprache' : 'Email language', value: (emailLanguage || 'EN').toUpperCase() === 'DE' ? 'Deutsch' : 'English', status: 'ok' },
            { label: isDe ? 'Bestätigungs-Mails' : 'Confirmation emails', value: disableEmails ? (isDe ? 'deaktiviert' : 'disabled') : (isDe ? 'aktiv' : 'on'), status: disableEmails ? 'ok' : 'default' },
            ...(!disableEmails && (disableRegistrationEmail || disableCancellationEmail) ? [{ label: isDe ? 'Einzeln deaktiviert' : 'Individually disabled', value: [disableRegistrationEmail ? (isDe ? 'Anmelde-Bestätigung' : 'registration confirmation') : '', disableCancellationEmail ? (isDe ? 'Abmelde-Bestätigung' : 'cancellation confirmation') : ''].filter(Boolean).join(', '), status: 'ok' as CheckStatus }] : []),
            { label: isDe ? 'Outlook-Termin' : 'Outlook invite', value: disableOutlook ? (isDe ? 'deaktiviert' : 'disabled') : (isDe ? 'aktiv' : 'on'), status: disableOutlook ? 'ok' : 'default' },
            { label: isDe ? 'Auto-Abmeldung bei Outlook-Absage' : 'Auto-cancel on Outlook decline', value: autoDeregisterOnDecline ? (isDe ? 'an' : 'on') : (isDe ? 'aus' : 'off'), status: autoDeregisterOnDecline ? 'ok' : 'default' },
            { label: isDe ? 'Person nicht mehr bei Deloitte' : 'Person no longer at Deloitte', value: inactiveHandling === 'autoderegister' ? (isDe ? 'automatisch abmelden' : 'auto-deregister') : (isDe ? 'Organizer informieren' : 'notify organizer'), status: inactiveHandling === 'autoderegister' ? 'ok' : 'default' },
          ],
        });
        sections.push({
          title: isDe ? 'Schritt 7 — Team-Anmeldung' : 'Step 7 — Team registration',
          rows: [{ label: isDe ? 'Team-Anmeldung' : 'Team registration', value: teamRegistrationEnabled ? (isDe ? `aktiv — Teams à ${teamSize}` : `on — teams of ${teamSize}`) : (isDe ? 'aus' : 'off'), status: teamRegistrationEnabled ? 'ok' : 'default' }],
        });
        sections.push({
          title: isDe ? 'Schritt 8 — Dokumente' : 'Step 8 — Documents',
          rows: [{ label: isDe ? 'Dokumente' : 'Documents', value: documents.length ? `${documents.length}` : '—', status: documents.length ? 'ok' : 'empty' }],
        });
        sections.push({
          title: isDe ? 'Schritt 9 — Fun-Zone' : 'Step 9 — Fun zone',
          rows: [{ label: 'Quiz', value: quiz.length ? `${quiz.length} ${isDe ? 'Fragen' : 'questions'}` : '—', status: quiz.length ? 'ok' : 'empty' }],
        });
        const allRows = sections.reduce((acc, s) => acc + s.rows.length, 0);
        void allRows;
        const missingCount = sections.reduce((acc, s) => acc + s.rows.filter(r => r.status === 'missing').length, 0);
        const emptyCount = sections.reduce((acc, s) => acc + s.rows.filter(r => r.status === 'empty').length, 0);
        const chip = (st: CheckStatus): React.ReactElement | null => {
          if (st === 'default') return <span style={{ fontSize: '0.66rem', fontWeight: 700, padding: '1px 8px', borderRadius: 999, background: 'var(--dex-gray-100)', color: 'var(--dex-gray-500)', flexShrink: 0 }}>{isDe ? 'Standard' : 'Default'}</span>;
          if (st === 'empty') return <span style={{ fontSize: '0.66rem', fontWeight: 700, padding: '1px 8px', borderRadius: 999, background: 'rgba(237,139,0,0.12)', color: 'var(--dex-orange-dark, #b35a00)', flexShrink: 0 }}>{isDe ? 'leer (optional)' : 'empty (optional)'}</span>;
          if (st === 'missing') return <span style={{ fontSize: '0.66rem', fontWeight: 700, padding: '1px 8px', borderRadius: 999, background: 'rgba(218,41,28,0.12)', color: 'var(--dex-red, #c00)', flexShrink: 0 }}>{isDe ? 'fehlt' : 'missing'}</span>;
          return null;
        };
        return (
          <div
            role="dialog"
            aria-modal="true"
            onClick={() => setShowConfigCheck(false)}
            style={{
              position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1250,
              display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
            }}
          >
            <div
              onClick={e => e.stopPropagation()}
              className="card"
              style={{ width: '100%', maxWidth: 760, maxHeight: '88vh', overflow: 'auto', padding: 24, borderRadius: 16, background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.25)' }}
            >
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 4 }}>
                <h3 style={{ margin: 0, fontSize: '1.1rem' }}>{isDe ? 'Event prüfen — alle Einstellungen im Überblick' : 'Review event — all settings at a glance'}</h3>
                <button type="button" onClick={() => setShowConfigCheck(false)} style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4 }} aria-label={isDe ? 'Schließen' : 'Close'}><X size={20} /></button>
              </div>
              <p style={{ margin: '0 0 14px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {missingCount > 0
                  ? (isDe ? <><strong style={{ color: 'var(--dex-red, #c00)' }}>{missingCount} Pflichtangabe(n) fehlen</strong>{emptyCount > 0 ? <> · {emptyCount} optionale Punkte sind noch leer</> : null}.</> : <><strong style={{ color: 'var(--dex-red, #c00)' }}>{missingCount} required item(s) missing</strong>{emptyCount > 0 ? <> · {emptyCount} optional items still empty</> : null}.</>)
                  : emptyCount > 0
                    ? (isDe ? <>Alle Pflichtangaben sind gesetzt — <strong>{emptyCount} optionale Punkte</strong> sind noch leer.</> : <>All required items are set — <strong>{emptyCount} optional items</strong> are still empty.</>)
                    : (isDe ? 'Alles gesetzt — keine offenen Punkte.' : 'Everything set — nothing open.')}
              </p>
              {sections.map((sec, si) => (
                <div key={si} style={{ marginBottom: 14 }}>
                  <div style={{ fontWeight: 700, fontSize: '0.85rem', color: 'var(--dex-green-dark, #4a7c1f)', borderBottom: '1px solid var(--dex-gray-100)', paddingBottom: 4, marginBottom: 6 }}>{sec.title}</div>
                  {sec.rows.map((r, ri) => (
                    <div key={ri} style={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', alignItems: 'flex-start', gap: isMobile ? 2 : 10, padding: '4px 0', fontSize: '0.8rem' }}>
                      <span style={{ flex: isMobile ? '0 0 auto' : '0 0 230px', width: isMobile ? '100%' : undefined, color: 'var(--dex-gray-500)' }}>{r.label}</span>
                      <span style={{ flex: 1, color: 'var(--dex-gray-800)', minWidth: 0, overflowWrap: 'anywhere' }}>{r.value}</span>
                      {chip(r.status)}
                    </div>
                  ))}
                </div>
              ))}
              <div style={{ textAlign: 'right', marginTop: 6 }}>
                <button className="btn btn-primary" onClick={() => setShowConfigCheck(false)}>{isDe ? 'Schließen' : 'Close'}</button>
              </div>
            </div>
          </div>
        );
      })()}

      {showDemoVariantModal && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
          onClick={() => setShowDemoVariantModal(false)}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 760, maxHeight: '90vh', overflow: 'auto',
              padding: 24, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                {isDe ? 'Demo-Daten laden' : 'Load demo data'}
              </h3>
              <button
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-600)' }}
                onClick={() => setShowDemoVariantModal(false)}
                aria-label={isDe ? 'Schließen' : 'Close'}
              >
                <X size={20} />
              </button>
            </div>
            <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
              {isDe
                ? 'Wähle eine Vorlage. Die ausgewählte Variante überschreibt deine aktuellen Eingaben — verworfen wird nichts, falls du noch nichts gespeichert hast.'
                : 'Choose a template. The selected variant overrides your current input — nothing is lost if you haven\'t saved yet.'}
            </p>
            {(() => {
              const cards: Array<{ key: keyof typeof DEMO_VARIANTS; titleDe: string; titleEn: string; descDe: string; descEn: string }> = [
                {
                  key: 'standard',
                  titleDe: 'Standard',
                  titleEn: 'Standard',
                  descDe: 'Ein Event, eine Gruppe. Typisches Meeting / Lunch.',
                  descEn: 'One event, one group. Typical meeting or lunch.',
                },
                {
                  key: 'groups',
                  titleDe: 'Mit Gruppen',
                  titleEn: 'With groups',
                  descDe: 'Event mit zwei Teilnehmer-Gruppen (Split Capacity), z.B. Vormittag / Nachmittag.',
                  descEn: 'Event with two participant groups (split capacity), e.g. morning / afternoon.',
                },
                {
                  key: 'subevent',
                  titleDe: 'Mit Sub-Event',
                  titleEn: 'With sub-event',
                  descDe: 'Haupt-Event + 1 Sub-Event, z.B. Conference + Dinner.',
                  descEn: 'Main event + 1 sub-event, e.g. conference + dinner.',
                },
                {
                  key: 'subeventTeam',
                  titleDe: 'Mit Sub-Event + Team',
                  titleEn: 'With sub-event + team',
                  descDe: 'Wie links, aber mit Team-Anmeldung (Teams à 4 Personen).',
                  descEn: 'Same as on the left, but with team registration (teams of 4 people).',
                },
              ];
              return (
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: 12 }}>
                  {cards.map(card => (
                    <button
                      key={card.key}
                      type="button"
                      onClick={() => {
                        DEMO_VARIANTS[card.key]();
                        setShowDemoVariantModal(false);
                      }}
                      style={{
                        textAlign: 'left',
                        padding: 16,
                        borderRadius: 12,
                        border: '1px solid var(--dex-gray-200)',
                        background: '#fff',
                        cursor: 'pointer',
                        transition: 'all 0.15s ease',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: 8,
                        minHeight: 120,
                      }}
                      onMouseEnter={e => {
                        e.currentTarget.style.borderColor = 'var(--dex-green-dark, #4a7c1f)';
                        e.currentTarget.style.boxShadow = '0 6px 16px rgba(74,124,31,0.12)';
                        e.currentTarget.style.transform = 'translateY(-2px)';
                      }}
                      onMouseLeave={e => {
                        e.currentTarget.style.borderColor = 'var(--dex-gray-200)';
                        e.currentTarget.style.boxShadow = 'none';
                        e.currentTarget.style.transform = 'translateY(0)';
                      }}
                    >
                      <div style={{ fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
                        {isDe ? card.titleDe : card.titleEn}
                      </div>
                      <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                        {isDe ? card.descDe : card.descEn}
                      </div>
                    </button>
                  ))}
                </div>
              );
            })()}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 20, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
              <button
                type="button"
                className="btn btn-outline"
                onClick={() => setShowDemoVariantModal(false)}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* v17.21: A4-Zusammenfassungs-Modal nach erfolgreichem Save — fragt
          den Organizer, ob er das gesamte Event als PDF oder Word herunter-
          laden möchte (z.B. zur Durchsicht durch einen Partner). Beim
          Klick auf eine Option läuft der Export sofort, danach feuert der
          eigentliche „Wizard verlassen"-Dispatch (`dex-event-submit-success`).
          „Nein, danke" springt direkt zum Dispatch. */}
      {/* v28.74: „Einstellungen auf andere übertragen" — Auswahl WAS und WOHIN. */}
      {subTransfer && (() => {
        const srcName = shortSubEventTitle(subEvents[subTransfer.fromIdx]?.title || '', title) || (childTermSingular || 'Sub-Event');
        const toggleGroup = (key: string): void => setSubTransfer(prev => prev && ({
          ...prev,
          groups: prev.groups.indexOf(key) >= 0 ? prev.groups.filter(k => k !== key) : [...prev.groups, key],
        }));
        const toggleTarget = (i: number): void => setSubTransfer(prev => prev && ({
          ...prev,
          targets: prev.targets.indexOf(i) >= 0 ? prev.targets.filter(x => x !== i) : [...prev.targets, i],
        }));
        const canApply = subTransfer.groups.length > 0 && subTransfer.targets.length > 0;
        return (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 1300, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
            <div className="card" style={{ width: '100%', maxWidth: 680, maxHeight: '88vh', overflow: 'auto', padding: 24, borderRadius: 14, background: '#fff' }}>
              <h3 style={{ margin: '0 0 4px', fontSize: '1.1rem' }}>
                {isDe ? 'Einstellungen übertragen' : 'Transfer settings'}
              </h3>
              <p style={{ margin: '0 0 16px', fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {isDe
                  ? <>Die ausgewählten Einstellungen von <strong>„{srcName}“</strong> werden auf die ausgewählten {childTermPlural || 'Sub-Events'} übertragen und <strong>überschreiben</strong> die dortigen Werte. Gespeichert wird erst, wenn du den Assistenten speicherst.</>
                  : <>The selected settings of <strong>„{srcName}“</strong> are applied to the selected sub-events and <strong>overwrite</strong> their values. Nothing is stored until you save the wizard.</>}
              </p>

              <div style={{ fontWeight: 700, fontSize: '0.85rem', marginBottom: 8 }}>
                {isDe ? '1. Was soll übertragen werden?' : '1. What should be transferred?'}
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 18 }}>
                {SUB_TRANSFER_GROUPS.map(g => {
                  const on = subTransfer.groups.indexOf(g.key) >= 0;
                  const n = subGroupDiffCount(subTransfer.fromIdx, g.fields);
                  return (
                    <label key={g.key} style={{
                      display: 'flex', alignItems: 'flex-start', gap: 10, padding: '8px 12px', borderRadius: 8,
                      border: `1px solid ${on ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                      background: on ? 'rgba(134,188,37,0.06)' : '#fff', cursor: 'pointer', fontSize: '0.84rem',
                    }}>
                      <input type="checkbox" checked={on} onChange={() => toggleGroup(g.key)} style={{ marginTop: 3 }} />
                      <span>
                        {isDe ? g.de : g.en}
                        {n > 0 && (
                          <span style={{ marginLeft: 8, fontSize: '0.72rem', fontWeight: 700, padding: '1px 7px', borderRadius: 10, background: '#fff3d6', color: '#7a5a12', border: '1px solid #e0b34d' }}>
                            {isDe ? `${n}× abweichend` : `${n}× differing`}
                          </span>
                        )}
                      </span>
                    </label>
                  );
                })}
              </div>

              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                <div style={{ fontWeight: 700, fontSize: '0.85rem' }}>
                  {isDe ? '2. Auf welche übertragen?' : '2. Transfer to which ones?'}
                </div>
                <button
                  type="button"
                  className="btn btn-secondary"
                  style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                  onClick={() => setSubTransfer(prev => prev && ({
                    ...prev,
                    targets: prev.targets.length === subEvents.length - 1
                      ? []
                      : subEvents.map((_, i) => i).filter(i => i !== prev.fromIdx),
                  }))}
                >
                  {subTransfer.targets.length === subEvents.length - 1
                    ? (isDe ? 'Keine' : 'None')
                    : (isDe ? 'Alle' : 'All')}
                </button>
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 20 }}>
                {subEvents.map((s, i) => {
                  if (i === subTransfer.fromIdx) return null;
                  const on = subTransfer.targets.indexOf(i) >= 0;
                  const nm = shortSubEventTitle(s.title, title) || (isDe ? 'Ohne Titel' : 'Untitled');
                  return (
                    <label key={s.id || i} style={{
                      display: 'inline-flex', alignItems: 'center', gap: 6, padding: '6px 10px', borderRadius: 8,
                      border: `1px solid ${on ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                      background: on ? 'rgba(134,188,37,0.10)' : '#fff', cursor: 'pointer', fontSize: '0.8rem',
                    }}>
                      <input type="checkbox" checked={on} onChange={() => toggleTarget(i)} />
                      {nm}
                    </label>
                  );
                })}
              </div>

              <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button type="button" className="btn btn-secondary" onClick={() => setSubTransfer(null)}>
                  {isDe ? 'Abbrechen' : 'Cancel'}
                </button>
                <button type="button" className="btn btn-primary" disabled={!canApply} onClick={applySubTransfer}>
                  {isDe
                    ? `Auf ${subTransfer.targets.length} übertragen`
                    : `Transfer to ${subTransfer.targets.length}`}
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {showSummaryModal && pendingSuccessDispatch && (() => {
        const closeAndDispatch = (): void => {
          // v17.22: Ref VOR dem Dispatch leeren, damit der Unmount-Cleanup
          // nicht ein zweites Mal feuert (Doppel-Navigation/-Banner).
          const payload = pendingSuccessDispatchRef.current || pendingSuccessDispatch;
          pendingSuccessDispatchRef.current = null;
          setShowSummaryModal(false);
          setPendingSuccessDispatch(null);
          try {
            window.dispatchEvent(new CustomEvent('dex-event-submit-success', {
              detail: payload,
            }));
          } catch { /* */ }
        };
        const buildData = (): SummaryData => {
          // Bild als DataURL (falls noch nicht Base64): unten reicht der
          // bestehende imagePreview, der bei neu hochgeladenen Bildern
          // bereits eine Data-URL ist und bei bestehenden Events die
          // SharePoint-URL. Letztere wird im PDF/Doc-Export im Print-View
          // i.d.R. nicht geladen (CORS) — wir bauen einen Fallback-Text.
          const subEventsForSummary = subEvents.map(se => ({
            title: se.title || '',
            startDate: se.startDate,
            endDate: se.endDate,
            location: se.location,
            description: se.description,
            maxParticipants: typeof se.maxParticipants === 'number' ? se.maxParticipants : undefined,
            waitlistEnabled: !!se.waitlistEnabled,
          }));
          const customFieldsForSummary = customFields
            .filter(f => f.label && f.label.trim().length > 0)
            .map(f => ({
              id: f.id,
              label: f.label,
              type: f.type,
              required: !!f.required,
              helpText: f.helpText,
              helpTextStyle: f.helpTextStyle,
              confirmLabel: f.confirmLabel,
              options: f.type === 'select' ? f.options : undefined,
              multi: !!f.multi,
              onlyForGroup: f.onlyForGroup,
              labelEn: f.labelEn,
              helpTextEn: f.helpTextEn,
              confirmLabelEn: f.confirmLabelEn,
              optionsEn: f.optionsEn,
              showIf: f.showIf,
            }));
          // Transferzeiten + Agenda werden in den Summary-Helper als
          // vereinfachte Spalten gemappt — das Detail-Schema bleibt im
          // Wizard, der Export nimmt die für Reviewer relevanten Spalten.
          const transfersForSummary = transferTimes.map(t => ({
            time: [t.date, t.departureTime].filter(Boolean).join(' '),
            description: [t.location, t.description, t.meetingPoint].filter(Boolean).join(' — '),
          }));
          const agendaForSummary = agenda.map(a => ({
            time: [a.date, a.time, a.endTime ? ` – ${a.endTime}` : ''].filter(Boolean).join(' '),
            topic: a.title,
            speaker: a.description,
          }));
          const quizForSummary = quiz.map(q => ({
            question: q.question,
            options: q.options,
            correctIndex: (q.correctIndices && q.correctIndices.length > 0) ? q.correctIndices[0] : undefined,
          }));
          const documentsForSummary = documents.map(doc => ({
            name: doc.name,
            size: doc.size,
          }));
          return {
            title,
            description,
            imageDataUrl: imagePreview || eventImageUrl || undefined,
            startDate,
            endDate,
            organizers: organizer.split(';').map(s => s.trim()).filter(Boolean),
            organizerEmails,
            contactName,
            contactEmail,
            contactInfo,
            testTeam: testTeamNames,
            qrScanners: qrScannerNames,
            isFictive,
            activeFrom,
            location,
            address: { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity },
            agenda: agendaForSummary,
            transfers: transfersForSummary,
            // v29.21 (Audit): beide Strings sind im Wizard KOMMA-separiert —
            // split(';') lieferte immer ein 1-elementiges Array, der Export
            // meldete bei fuenf Verteilern „1 Eintrag".
            locationFilter: locationFilter ? locationFilter.split(',').map(s => s.trim()).filter(Boolean) : [],
            audience: audience ? audience.split(',').map(s => s.trim()).filter(Boolean) : [],
            filterMode,
            excludedUsers,
            registrationDeadline,
            lastDeregisterDate,
            maxParticipants: Number(maxParticipants) || 0,
            unlimitedParticipants,
            waitlistEnabled,
            durchstarterCapacity: Number(durchstarterCapacity) || 0,
            funstarterCapacity: Number(funstarterCapacity) || 0,
            splitLabelA, splitLabelB,
            splitSharedWaitlist,
            teamRegistrationEnabled,
            teamSize,
            askTeamName,
            teamPartialAllowed,
            teamOpenSlotsVisible,
            teamJoinRequiresApproval,
            askSalutation,
            bilingualFields,
            customFields: customFieldsForSummary,
            allowAttendeeUpload,
            attendeeUploadHint,
            attendeeUploadLabel,
            emailLanguage,
            disableEmails,
            disableOutlook,
            outlookHeading,
            outlookSubheading,
            outlookBody,
            notifyOrgRegisterMode,
            notifyOrgRegisterFromDate,
            notifyOrgCancelMode,
            documents: documentsForSummary,
            funZone: quizForSummary,
            quizClusterSize,
            subEvents: subEventsForSummary,
            childTermSingular,
            childTermPlural,
            subEventsOnlyMode,
            requireSubEventSelection,
            generatedAt: new Date().toISOString(),
            locale: isDe ? 'de' : 'en',
          };
        };
        const onPdf = (): void => {
          try { exportSummaryAsPdf(buildData()); } catch (err) {
            console.warn('[DEX] exportSummaryAsPdf failed:', err);
          }
          closeAndDispatch();
        };
        const onDoc = (): void => {
          try { exportSummaryAsDoc(buildData()); } catch (err) {
            console.warn('[DEX] exportSummaryAsDoc failed:', err);
          }
          closeAndDispatch();
        };
        return (
          <div
            style={{
              position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1300,
              display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
            }}
            onClick={closeAndDispatch}
          >
            <div
              className="card"
              style={{
                width: '100%', maxWidth: 560, padding: 24, borderRadius: 16,
                background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
              }}
              onClick={e => e.stopPropagation()}
            >
              <h3 style={{ margin: 0, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                {isDe ? 'Event-Zusammenfassung herunterladen?' : 'Download event summary?'}
              </h3>
              <p style={{ marginTop: 12, color: 'var(--dex-gray-700)', lineHeight: 1.55, fontSize: '0.95rem' }}>
                {isDe
                  ? <>Das Event wurde gespeichert. Möchtest du jetzt eine <strong>A4-Zusammenfassung</strong> mit allen Sektionen (Foto, Beschreibung, Sichtbarkeit, Felder, Kommunikation, Dokumente, Sub-Events…) herunterladen? Du kannst sie z.B. einem Partner zur Durchsicht weiterleiten.</>
                  : <>The event has been saved. Would you like to download a <strong>one-page A4 summary</strong> with every section (photo, description, visibility, fields, communication, documents, sub-events…)? You can forward it to a partner for review.</>}
              </p>
              <div style={{
                marginTop: 18, padding: '10px 14px', background: 'rgba(0,90,156,0.06)',
                border: '1px solid rgba(0,90,156,0.25)', borderRadius: 8,
                fontSize: '0.82rem', color: 'var(--dex-gray-700)',
              }}>
                {isDe
                  ? <><strong>Hinweis:</strong> Beim PDF-Export öffnet sich der Browser-Druckdialog. Wähle dort <strong>&bdquo;Als PDF speichern&ldquo;</strong> als Ziel. Word-Export lädt direkt eine .doc-Datei herunter.</>
                  : <><strong>Note:</strong> The PDF export opens the browser print dialog — pick <strong>&ldquo;Save as PDF&rdquo;</strong> as the destination. Word export downloads a .doc file directly.</>}
              </div>
              <div style={{ marginTop: 22, display: 'flex', flexWrap: 'wrap', gap: 10, justifyContent: 'flex-end' }}>
                <button className="btn btn-outline" onClick={closeAndDispatch}>
                  {isDe ? 'Nein, danke' : 'No, thanks'}
                </button>
                <button className="btn btn-secondary" onClick={onDoc}>
                  {isDe ? 'Als Word (.doc)' : 'As Word (.doc)'}
                </button>
                <button className="btn btn-primary" onClick={onPdf}>
                  {isDe ? 'Als PDF' : 'As PDF'}
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* v19.x: Das „Personen ausschließen"-Modal UND das „Sichtbarkeit prüfen"-Modal
          sind nach <AudiencePicker> gewandert (self-contained pro Instanz für
          Hauptevent + jedes Sub-Event). */}

      {/* v20.2: Self-Check-in-Erklär-Modal (v18.33) entfernt — die Erklärung
          lebt jetzt im Kachel-Modal des Admin Centers und im Handbuch. */}

      {/* v9.28/v13.4: Modal — neuer Quiz-Bereich anlegen, jetzt über <Modal>-Wrapper. */}
      <Modal
        open={newSectionModalOpen}
        onClose={() => setNewSectionModalOpen(false)}
        maxWidth={460}
        ariaLabel="Neuen Quiz-Bereich anlegen"
      >
        {newSectionModalOpen && (
          <>
            <h3 style={{ marginTop: 0, marginBottom: 8, fontSize: '1.15rem' }}>
              Neuen Bereich anlegen
            </h3>
            <p style={{ margin: '0 0 14px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              Bereiche bündeln Quiz-Fragen auf einer gemeinsamen Seite. Vergib einen
              kurzen, sprechenden Namen — z.B. <em>Orte</em>, <em>Geschichte</em> oder <em>Foto-Quiz</em>.
            </p>
            <input
              type="text"
              className="form-input"
              autoFocus
              value={newSectionName}
              placeholder='z.B. "Orte"'
              onChange={e => { setNewSectionName(e.target.value); setNewSectionError(''); }}
              onKeyDown={e => {
                if (e.key === 'Enter') {
                  const name = newSectionName.trim();
                  if (!name) { setNewSectionError('Bitte einen Namen eingeben.'); return; }
                  const existing = new Set<string>();
                  for (const q of quiz) if (q.section) existing.add(q.section);
                  for (const p of pendingSections) existing.add(p);
                  if (existing.has(name)) { setNewSectionError('Ein Bereich mit diesem Namen existiert bereits.'); return; }
                  setPendingSections([...pendingSections, name]);
                  setNewSectionModalOpen(false);
                } else if (e.key === 'Escape') {
                  setNewSectionModalOpen(false);
                }
              }}
              style={{ fontSize: '0.95rem', marginBottom: 8 }}
            />
            {newSectionError && (
              <div style={{ color: 'var(--dex-red, #c00)', fontSize: '0.78rem', marginBottom: 10 }}>{newSectionError}</div>
            )}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 12 }}>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => setNewSectionModalOpen(false)}
              >
                Abbrechen
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => {
                  const name = newSectionName.trim();
                  if (!name) { setNewSectionError('Bitte einen Namen eingeben.'); return; }
                  const existing = new Set<string>();
                  for (const q of quiz) if (q.section) existing.add(q.section);
                  for (const p of pendingSections) existing.add(p);
                  if (existing.has(name)) { setNewSectionError('Ein Bereich mit diesem Namen existiert bereits.'); return; }
                  setPendingSections([...pendingSections, name]);
                  setNewSectionModalOpen(false);
                }}
              >
                <Plus size={14} /> Bereich anlegen
              </button>
            </div>
          </>
        )}
      </Modal>

      {/* Modal: Vorgeschlagene Felder auswählen (Multi-Select) — v13.4 über <Modal>. */}
      <Modal
        open={showSuggestedModal}
        onClose={() => setShowSuggestedModal(false)}
        maxWidth={540}
        ariaLabel="Vorgeschlagene Felder auswählen"
      >
        {showSuggestedModal && (
          <div
            style={{
              display: 'flex', flexDirection: 'column', gap: 14,
            }}
          >
            <h2 style={{ margin: 0, fontSize: '1.25rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Vorgeschlagene Felder' : 'Suggested fields'}
            </h2>
            <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
              {isDe
                ? 'Wähle aus dem Katalog, welche Felder dem Event hinzugefügt werden sollen. Du kannst die Felder danach weiter anpassen.'
                : 'Pick the fields you want to add to the event. You can still tweak them afterwards.'}
            </p>
            {/* v10.21: Catalog gruppiert nach Kategorie. Allgemeine Felder
                immer ausgeklappt, B2Run-Felder default eingeklappt mit
                Toggle. Jeder Eintrag bekommt ein Badge mit der Kategorie. */}
            {(() => {
              const generalEntries = SUGGESTED_FIELDS_CATALOG.filter(s => s.category === 'general');
              const b2runEntries = SUGGESTED_FIELDS_CATALOG.filter(s => s.category === 'b2run');
              const renderEntry = (s: SuggestedEntry): React.ReactElement => (
                <label
                  key={s.key}
                  style={{
                    display: 'flex', gap: 10, alignItems: 'flex-start',
                    padding: '10px 12px', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 8, cursor: 'pointer',
                    background: suggestedSelection[s.key] ? 'var(--dex-gray-50, #fafafa)' : '#fff',
                  }}
                >
                  <input
                    type="checkbox"
                    checked={!!suggestedSelection[s.key]}
                    onChange={e => setSuggestedSelection({ ...suggestedSelection, [s.key]: e.target.checked })}
                    style={{ marginTop: 3, flexShrink: 0 }}
                  />
                  {/* v10.23: passendes Fluent-UI-Icon links neben dem Label,
                      damit die Auswahl auf einen Blick visuell wiedererkennbar
                      ist. Farbe analog zur Kategorie (grün=Allgemein,
                      orange=B2Run). */}
                  <Icon
                    iconName={s.icon}
                    style={{
                      fontSize: 20, flexShrink: 0, marginTop: 2,
                      color: s.category === 'b2run' ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)',
                    }}
                  />
                  <span style={{ flex: 1 }}>
                    <span style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                      <strong style={{ fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>{s.label}</strong>
                      <span style={{
                        fontSize: '0.65rem', fontWeight: 600,
                        padding: '2px 8px', borderRadius: 999,
                        textTransform: 'uppercase', letterSpacing: 0.5,
                        background: s.category === 'b2run' ? 'rgba(237,139,0,0.12)' : 'rgba(134,188,37,0.12)',
                        color: s.category === 'b2run' ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)',
                      }}>
                        {s.category === 'b2run' ? 'B2Run' : (isDe ? 'Allgemein' : 'General')}
                      </span>
                      {/* v10.23: i-Tooltip mit ausführlichem Hinweis was das
                          Feld in der App tut — verhindert Klick-und-Probier-
                          Modus, weil der Organizer schon vor Auswahl sieht
                          welche Frage-Form (Dropdown / Freitext / Pflicht-
                          Checkbox) und welcher Effekt (Anzeige im Admin-Center,
                          Excel-Export, etc.) entsteht. Klick auf das Label
                          (das `<label>`-Wrapping) würde die Checkbox togglen
                          — das `onClick`-stopPropagation des InfoTooltip
                          verhindert das. */}
                      <span onClick={e => e.preventDefault()} style={{ display: 'inline-flex' }}>
                        <InfoTooltip text={s.tooltip || s.description} />
                      </span>
                    </span>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>{s.description}</div>
                  </span>
                </label>
              );
              return (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginTop: 4 }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                    {generalEntries.map(renderEntry)}
                  </div>
                  <div style={{ borderTop: '1px solid var(--dex-gray-200)', paddingTop: 14 }}>
                    <button
                      type="button"
                      onClick={() => setShowB2runSuggested(v => !v)}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        background: 'none', border: 'none', padding: 0,
                        fontSize: '0.85rem', fontWeight: 600, cursor: 'pointer',
                        color: 'var(--dex-gray-700)',
                      }}
                    >
                      <span style={{ display: 'inline-flex', transform: showB2runSuggested ? 'rotate(90deg)' : 'rotate(0deg)', transition: 'transform 0.15s' }}>▶</span>
                      {isDe ? 'B2Run-spezifische Felder' : 'B2Run-specific fields'}
                      <span style={{
                        fontSize: '0.65rem', fontWeight: 600,
                        padding: '2px 8px', borderRadius: 999,
                        background: 'rgba(237,139,0,0.12)',
                        color: 'var(--dex-orange-dark, #b35a00)',
                      }}>
                        B2Run · {b2runEntries.length}
                      </span>
                    </button>
                    {showB2runSuggested && (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 10 }}>
                        <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)', lineHeight: 1.45 }}>
                          {isDe
                            ? 'Diese Felder sind speziell für B2Run-Lauf-Events vorgesehen (Startblock, Altersklasse, Datenschutz-Checkbox mit b2run.de-Links etc.). Bei normalen Events brauchst du sie nicht.'
                            : 'These fields are intended for B2Run running events (start block, age group, B2Run-specific privacy checkbox etc.). Skip them for standard events.'}
                        </p>
                        {b2runEntries.map(renderEntry)}
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'space-between', marginTop: 6 }}>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  type="button"
                  className="btn btn-outline"
                  style={{ fontSize: '0.8rem', padding: '4px 12px' }}
                  onClick={() => {
                    const all: Record<string, boolean> = {};
                    for (const s of SUGGESTED_FIELDS_CATALOG) all[s.key] = true;
                    setSuggestedSelection(all);
                  }}
                >
                  {isDe ? 'Alle' : 'All'}
                </button>
                <button
                  type="button"
                  className="btn btn-outline"
                  style={{ fontSize: '0.8rem', padding: '4px 12px' }}
                  onClick={() => setSuggestedSelection({})}
                >
                  {isDe ? 'Keine' : 'None'}
                </button>
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setShowSuggestedModal(false)}
                >
                  {isDe ? 'Abbrechen' : 'Cancel'}
                </button>
                <button
                  type="button"
                  className="btn btn-primary"
                  onClick={addSelectedSuggestedFields}
                  disabled={!Object.values(suggestedSelection).some(Boolean)}
                >
                  {isDe ? 'Hinzufügen' : 'Add'}
                </button>
              </div>
            </div>
          </div>
        )}
      </Modal>

      {/* v22.62: „Sichtbarkeit auf Sub-Events übernehmen?" — erscheint beim
          ersten „Weiter"/Speichern, sobald die Klammer eine Sichtbarkeit hat
          und Sub-Events existieren. */}
      <Modal
        open={visCopyModalOpen}
        onClose={() => closeVisCopy(false)}
        maxWidth={560}
        dismissable={false}
        ariaLabel={isDe ? 'Sichtbarkeit übernehmen' : 'Apply visibility'}
      >
        {visCopyModalOpen && (
          <div>
            <h2 style={{ margin: '0 0 10px', fontSize: '1.15rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {isDe ? 'Sichtbarkeit auf alle Sub-Events übernehmen?' : 'Apply visibility to all sub-events?'}
            </h2>
            <p style={{ margin: '0 0 16px', fontSize: '0.9rem', color: 'var(--dex-gray-700)', lineHeight: 1.55 }}>
              {isDe
                ? <>Du hast für {subEventsOnlyMode ? 'die Klammer' : 'das Hauptevent'} eine Sichtbarkeit gesetzt. Sollen <strong>alle {subEvents.length} Sub-Events</strong> dieselbe Sichtbarkeit (Standortfilter + Mailverteiler + Verknüpfung) übernehmen?<br /><br />Das ist meist sinnvoll, damit jeder, der das Event sehen soll, auch die Sub-Events erreicht — der Zugang läuft ohnehin über die Sichtbarkeit des Gesamt-Events. Bereits gesetzte, abweichende Sub-Event-Sichtbarkeiten werden dabei <strong>überschrieben</strong>.</>
                : <>You set a visibility for {subEventsOnlyMode ? 'the bracket' : 'the main event'}. Should <strong>all {subEvents.length} sub-events</strong> adopt the same visibility (location filter + mailing lists + combination)?<br /><br />This usually makes sense so that everyone who should see the event can also reach the sub-events — access runs through the overall event’s visibility anyway. Any existing, differing sub-event visibilities will be <strong>overwritten</strong>.</>}
            </p>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, flexWrap: 'wrap' }}>
              <button className="btn btn-secondary" onClick={() => closeVisCopy(false)}>
                {isDe ? 'Nein, eigene behalten' : 'No, keep their own'}
              </button>
              <button className="btn btn-primary" onClick={() => closeVisCopy(true)}>
                {isDe ? `Ja, auf alle ${subEvents.length} übernehmen` : `Yes, apply to all ${subEvents.length}`}
              </button>
            </div>
          </div>
        )}
      </Modal>

      {/* v11.57 / v11.63 / v13.4: Outlook-Update-Confirm-Modal über <Modal>-Wrapper.
          dismissable=false, da Schließen nur über Cancel-Button erlaubt. */}
      <Modal
        open={outlookConfirmOpen}
        onClose={cancelOutlookSave}
        maxWidth={620}
        dismissable={false}
        ariaLabel="Outlook-Update bestätigen"
      >
        {outlookConfirmOpen && (
          <div>
            <h2 id="outlook-confirm-title" style={{
              margin: '0 0 10px', fontSize: '1.15rem', fontWeight: 700,
              color: 'var(--dex-green-dark, #4a7c1f)',
            }}>
              {isDe ? 'Outlook-Termin der Teilnehmer aktualisieren?' : 'Update Outlook invite for attendees?'}
            </h2>
            <p style={{ margin: '0 0 14px', fontSize: '0.9rem', color: 'var(--dex-gray-700)', lineHeight: 1.55 }}>
              {isDe
                ? 'Du hast Felder geändert, die für die Teilnehmer-Outlook-Termine relevant sind. Wähle aus, welche Termine du jetzt neu rausschicken willst — der Rest wird gespeichert, aber Outlook bleibt unangetastet (du kannst das später jederzeit nachholen).'
                : 'You changed fields that are relevant to the attendees’ Outlook invites. Pick which invites you want to resend now — everything else is saved, but Outlook is left alone (you can resend later at any time).'}
            </p>
            <div style={{
              border: '1px solid var(--dex-gray-200)',
              borderRadius: 8,
              marginBottom: 14,
              background: 'var(--dex-gray-50, #f8f9fa)',
            }}>
              {outlookConfirmItems.map((it, idx) => {
                const isLast = idx === outlookConfirmItems.length - 1;
                const fieldLabelMap: Record<'title'|'startDate'|'endDate'|'outlookBody'|'location'|'subject'|'layout'|'organizer'|'logo', { de: string; en: string }> = {
                  title: { de: 'Titel', en: 'Title' },
                  startDate: { de: 'Startzeit', en: 'Start time' },
                  endDate: { de: 'Endzeit', en: 'End time' },
                  outlookBody: { de: 'Termin-Text', en: 'Calendar body' },
                  location: { de: 'Ort', en: 'Location' },
                  subject: { de: 'Betreff', en: 'Subject' },
                  layout: { de: 'Kopfbild (Größe/Abstand)', en: 'Header image (size/spacing)' },
                  organizer: { de: 'Organizer (im Termin-Text)', en: 'Organizer (in calendar body)' },
                  logo: { de: 'Kopfbild', en: 'Header image' },
                };
                const changedLabels = it.changedFields.map(f => isDe ? fieldLabelMap[f].de : fieldLabelMap[f].en).join(', ');
                const checked = !!outlookConfirmChecks[it.eventId];
                // v11.69: noOutlookYet-Items bekommen wieder eine Checkbox.
                // Default UNCHECKED. Beim Anhaken wird das Sub-Event in der
                // Eventverwaltung komplett neu angelegt (DEX_Events-Item
                // delete + create mit `existingSubsiteUrl`), damit der
                // Outlook-Termin entsteht. Die bestehende Teilnehmerliste
                // mit allen Anmeldungen bleibt unangetastet.
                if (it.noOutlookYet) {
                  return (
                    <label
                      key={it.eventId}
                      style={{
                        display: 'flex', alignItems: 'flex-start', gap: 12,
                        padding: '12px 14px',
                        borderBottom: isLast ? 'none' : '1px solid var(--dex-gray-200)',
                        cursor: 'pointer',
                        background: '#fffaf0',
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={e => {
                          const next = e.target.checked;
                          setOutlookConfirmChecks(prev => ({ ...prev, [it.eventId]: next }));
                        }}
                        style={{ width: 18, height: 18, marginTop: 2, cursor: 'pointer', flexShrink: 0 }}
                      />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: '0.92rem', fontWeight: 600, color: 'var(--dex-gray-800)', wordBreak: 'break-word' }}>
                          {isDe ? `Sub-Event: ${it.title}` : `Sub-event: ${it.title}`}
                        </div>
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 3 }}>
                          {isDe ? 'Geändert: ' : 'Changed: '}{changedLabels}
                        </div>
                        <div style={{ fontSize: '0.76rem', color: '#8a6d3b', marginTop: 6, lineHeight: 1.45, background: '#fcf8e3', border: '1px solid #faebcc', borderRadius: 4, padding: '6px 8px' }}>
                          {isDe
                            ? <>Für dieses Sub-Event gibt es noch keinen Outlook-Termin. Wenn du den Haken setzt, wird das Sub-Event in der Eventverwaltung neu angelegt, damit der Outlook-Termin entsteht. <strong>Die bestehende Teilnehmerliste mit allen Anmeldungen bleibt erhalten</strong> — nur die DEX_Events-Zeile bekommt eine neue ID.</>
                            : <>This sub-event does not have an Outlook event yet. If you tick the box, the sub-event is re-created in the event admin so the Outlook event can be generated. <strong>The existing participant list with all registrations stays intact</strong> — only the DEX_Events row gets a new ID.</>}
                        </div>
                      </div>
                    </label>
                  );
                }
                // v15.3: leere changedFields-Liste = Item kommt aus dem
                // persistierten OutlookDirty-Flag (frühere Session,
                // wurde damals nicht synchronisiert). Klartext-Hinweis
                // statt leerer „Geändert:"-Zeile.
                const isFromPersistedDirty = it.changedFields.length === 0;
                return (
                  <label
                    key={it.eventId}
                    style={{
                      display: 'flex', alignItems: 'flex-start', gap: 12,
                      padding: '12px 14px',
                      borderBottom: isLast ? 'none' : '1px solid var(--dex-gray-200)',
                      cursor: 'pointer',
                      background: isFromPersistedDirty ? '#fff8e8' : '#fff',
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={e => {
                        const next = e.target.checked;
                        setOutlookConfirmChecks(prev => ({ ...prev, [it.eventId]: next }));
                      }}
                      style={{ width: 18, height: 18, marginTop: 2, cursor: 'pointer', flexShrink: 0 }}
                    />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '0.92rem', fontWeight: 600, color: 'var(--dex-gray-800)', wordBreak: 'break-word' }}>
                        {it.kind === 'top'
                          ? (isDe ? `Hauptevent: ${it.title}` : `Main event: ${it.title}`)
                          : (isDe ? `Sub-Event: ${it.title}` : `Sub-event: ${it.title}`)}
                      </div>
                      {isFromPersistedDirty ? (
                        <div style={{ fontSize: '0.78rem', color: '#8a6d3b', marginTop: 4, lineHeight: 1.45 }}>
                          {isDe
                            ? <>⏳ <strong>Frühere Änderung nicht synchronisiert</strong> — beim letzten Speichern dieses Events wurden Outlook-relevante Felder geändert, der Outlook-Sync wurde aber damals übersprungen. Haken setzen, um die Teilnehmer jetzt nachträglich per Outlook-Update zu informieren.</>
                            : <>⏳ <strong>Earlier change not yet synced</strong> — Outlook-relevant fields were changed in a previous save of this event, but the Outlook sync was skipped at the time. Tick the box to send the catch-up Outlook update to attendees now.</>}
                        </div>
                      ) : (
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 3 }}>
                          {isDe ? 'Geändert: ' : 'Changed: '}{changedLabels}
                        </div>
                      )}
                    </div>
                  </label>
                );
              })}
            </div>
            <p style={{
              margin: '0 0 12px', fontSize: '0.8rem', color: 'var(--dex-gray-500)',
              lineHeight: 1.5,
            }}>
              {isDe
                ? 'Bei angehakten Events bekommen die Teilnehmer eine „Aktualisierter Termin"-Benachrichtigung von Outlook. Nicht angehakte Termine werden für später als „ausstehender Outlook-Sync" markiert.'
                : 'Ticked events trigger an “updated meeting” notification from Outlook for attendees. Unticked invites are flagged as “pending Outlook sync” for later.'}
            </p>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 8 }}>
              <button
                className="btn btn-secondary"
                onClick={cancelOutlookSave}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                className="btn btn-primary"
                style={{ background: 'var(--dex-green, #86bc25)', borderColor: 'var(--dex-green, #86bc25)' }}
                onClick={() => confirmOutlookSave()}
              >
                {isDe ? 'Speichern' : 'Save'}
              </button>
            </div>
          </div>
        )}
      </Modal>

      {/* v17.3: Unsaved-Changes-Confirm-Modal. Erscheint, wenn der User
          auf „Zurück" klickt und das Formular gegenüber dem Initial-
          Snapshot Änderungen hat.
          v30.1: Drei modusabhängige Wege, gestapelt statt nebeneinander:
           - Neu-Anlage: „Entwurf speichern" legt den Stand in den
             Entwurfs-Zwischenspeicher (v30.0) und verlässt den Wizard —
             beim nächsten Öffnen der Event-Erstellung wird er angeboten.
             „Event verwerfen" löscht den Entwurf endgültig.
           - Edit: „Änderungen speichern" = attemptSubmit wie bisher
             (blockt die Back-Nav, nach erfolgreichem Save navigiert der
             submit-success-Dispatch); „Änderungen verwerfen" verlässt
             ohne Speichern. */}
      {unsavedConfirmOpen && (
        <Modal
          open={true}
          onClose={() => { unsavedConfirmOpen.resolve(false); setUnsavedConfirmOpen(null); }}
          maxWidth={480}
          padding={24}
          ariaLabel={isDe ? 'Ungespeicherte Änderungen' : 'Unsaved changes'}
        >
          <h3 style={{ margin: '0 0 12px', fontSize: '1.1rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
            {isDe
              ? (isEditMode ? 'Ungespeicherte Änderungen' : 'Entwurf noch nicht gespeichert')
              : (isEditMode ? 'Unsaved changes' : 'Draft not saved yet')}
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: '0.9rem', lineHeight: 1.5, color: 'var(--dex-gray-700)' }}>
            {isDe
              ? (isEditMode
                ? <>Du hast Änderungen am Event vorgenommen, die noch <strong>nicht gespeichert</strong> sind. Was möchtest du tun?</>
                : <>Dein Event ist noch <strong>nicht angelegt</strong>. Du kannst den Stand als Entwurf behalten — beim nächsten Öffnen der Event-Erstellung machst du genau hier weiter. Hochgeladene Bilder sind im Entwurf nicht enthalten.</>)
              : (isEditMode
                ? <>You have made changes to this event that are <strong>not saved yet</strong>. What do you want to do?</>
                : <>Your event is <strong>not created yet</strong>. You can keep this state as a draft — next time you open event creation you continue right here. Uploaded images are not part of the draft.</>)}
          </p>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            {isEditMode ? (
              /* v17.7: blockt die laufende Back-Nav (resolve(false)) und
                 triggert attemptSubmit; nach erfolgreichem Save dispatched
                 EventCreationPage „dex-event-submit-success" und
                 DexEventPlatform navigiert zum Organizer-Menü. */
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => {
                  unsavedConfirmOpen.resolve(false);
                  setUnsavedConfirmOpen(null);
                  window.setTimeout(() => { attemptSubmit(); }, 0);
                }}
                style={{ fontSize: '0.9rem', width: '100%', justifyContent: 'center' }}
              >
                <Send size={14} /> {isDe ? 'Änderungen speichern' : 'Save changes'}
              </button>
            ) : (
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => {
                  // Sofort schreiben — der 1,5-s-Debounce des Autosaves hat
                  // die letzten Eingaben sonst evtl. noch nicht gesichert.
                  try { localStorage.setItem(DRAFT_KEY, JSON.stringify({ savedAt: Date.now(), data: buildDraftPayload() })); } catch { /* best-effort */ }
                  unsavedConfirmOpen.resolve(true);
                  setUnsavedConfirmOpen(null);
                }}
                style={{ fontSize: '0.9rem', width: '100%', justifyContent: 'center' }}
              >
                <Send size={14} /> {isDe ? 'Entwurf speichern' : 'Save draft'}
              </button>
            )}
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => { unsavedConfirmOpen.resolve(false); setUnsavedConfirmOpen(null); }}
              style={{ fontSize: '0.9rem', width: '100%', justifyContent: 'center' }}
            >
              {isDe
                ? (isEditMode ? 'Bearbeitung fortsetzen' : 'Eventerstellung fortsetzen')
                : (isEditMode ? 'Continue editing' : 'Continue creating')}
            </button>
            <button
              type="button"
              className="btn btn-danger"
              onClick={() => {
                if (!isEditMode) {
                  // Verwerfen heisst verwerfen — auch den Entwurfs-
                  // Zwischenspeicher, sonst bietet ihn der naechste
                  // Besuch wieder an.
                  try { localStorage.removeItem(DRAFT_KEY); } catch { /* */ }
                }
                unsavedConfirmOpen.resolve(true);
                setUnsavedConfirmOpen(null);
              }}
              style={{ fontSize: '0.9rem', width: '100%', justifyContent: 'center' }}
            >
              {isDe
                ? (isEditMode ? 'Änderungen verwerfen' : 'Event verwerfen')
                : (isEditMode ? 'Discard changes' : 'Discard event')}
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}
