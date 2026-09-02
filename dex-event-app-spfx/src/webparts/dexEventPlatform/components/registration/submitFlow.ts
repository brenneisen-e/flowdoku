/**
 * Submit-Pfad der Anmeldeseite — aus RegistrationPage.tsx ausgelagert (v30.66).
 *
 * `handleSubmit` (Validierung, Rueckfragen, Modals) und die drei Wege dahinter:
 * Team-Anmeldung, Beitritt zu einem offenen Team und die eigentliche
 * Registrierung. Die vier Funktionskoerper sind ZEICHENGLEICH uebernommen.
 *
 * Warum eine Fabrik und keine freien Funktionen mit Parametern: die Koerper
 * lesen 91 Bezeichner aus dem Komponenten-Scope. Als `c.xyz` haette jede
 * einzelne Zeile angefasst werden muessen — in einer Datei ohne Tests genau der
 * Umbau, der stillschweigend etwas verdreht. Stattdessen destrukturiert die
 * Fabrik das Kontext-Objekt einmal oben; darunter steht der Code wie vorher.
 * Der Aufruf steht an genau der Stelle, an der `handleSubmit` frueher
 * deklariert war, und laeuft wie vorher bei jedem Render — die Funktionen
 * sehen also denselben Render-Stand wie zuvor.
 */
import * as React from 'react';
import { isExternalEmailAddr, isPlausibleEmail, stripPrefilterKeys } from './regHelpers';
import { addPendingShadowParent, removePendingShadowParent } from '../../utils/shadowHeal';
import { BUNDLED_COMM_DEFAULT, BundledItem, bundledCommOf } from '../../utils/bundledComm';
import { collectCcEmailsFromFields } from '../../context/EventContext';
import { selfCancelLocked } from '../../utils/cancelPolicy';
import { Locale } from '../../context/LanguageContext';
import { DeloitteEvent, Salutation } from '../../types';

export interface SubmitFlowCtx {
  assistantModalDecidedRef: React.MutableRefObject<boolean>;
  cancelRegistration: (eventId: string, opts?: { suppressNotifications?: boolean; skipReload?: boolean; }) => Promise<boolean>;
  canCreateEvents: boolean;
  canDelegateAssistant: boolean;
  ccSelfDecidedRef: React.MutableRefObject<boolean>;
  ccSelfRef: React.MutableRefObject<boolean>;
  childEvents: DeloitteEvent[];
  childOneDe: string;
  childTermPlural: string;
  childTermSingular: string;
  confirmDialog: (message: React.ReactNode, opts?: import("../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  confirmDialogConfirmedRef: React.MutableRefObject<boolean>;
  createTeamJoinRequest: (eventId: string, teamId: string, customData?: Record<string, string>) => Promise<{ ok: boolean; itemId?: number; reason?: string; }>;
  currentUser: import("../../types/index").User;
  delegateAssistEnabled: boolean;
  delegateAssistValue: string;
  delegateChoiceRef: React.MutableRefObject<{ enabled: boolean; value: string; }>;
  delegateRegistrationToAssistant: (eventId: string, assistant: { email: string; name: string; }) => Promise<void>;
  durchCap: number;
  email: string;
  event: DeloitteEvent;
  eventSpecific: Record<string, string>;
  externalEmailConfirmedRef: React.MutableRefObject<boolean>;
  externalPerson: boolean;
  firstName: string;
  funCap: number;
  getEventNumbersForEmail: (email: string) => Promise<{ registered: number[]; waitlisted: number[]; }>;
  hiddenChildCount: number;
  isAllowedTargetForAssistant: (jt: string) => boolean;
  isAssistant: boolean;
  isDeadlinePassed: boolean;
  isSplitGroup: boolean;
  isTeamMode: boolean;
  joinTeam: (eventId: string, teamId: string, teamName: string, customData?: Record<string, string>) => Promise<{ ok: boolean; status?: "Angemeldet" | "Warteliste"; reason?: string; }>;
  locale: Locale;
  nothingToSubmit: boolean;
  otherConsentConfirmed: boolean;
  parentAlreadyRegistered: boolean;
  parentFullNoWaitlist: boolean;
  parentRegBlocked: boolean;
  pendingDocFiles: Record<string, File>;
  pendingJoinTeam: { teamId: string; teamName: string; };
  preferredStarterType: string;
  previewAsUser: boolean;
  recordProxyDelegation: (eventId: string, participant: { email: string; name: string; }) => Promise<void>;
  refreshEvents: () => Promise<void>;
  registerForEvent: (eventId: string, customData: Record<string, string>, participantFirstName?: string, participantLastName?: string, participantEmail?: string, preferredStarterType?: string, opts?: { skipShadowParent?: boolean; suppressMail?: boolean; suppressOutlook?: boolean; extraCc?: string; proxyConsentConfirmed?: boolean; actorAllowedAsAssistant?: boolean; skipReload?: boolean; bundledItems?: BundledItem[]; }) => Promise<{ ok: boolean; status: "Angemeldet" | "Warteliste"; reason?: string; }>;
  registerForOther: boolean;
  registerForParent: boolean;
  registerTeam: (eventId: string, leadData: { firstName: string; lastName: string; email: string; salutation?: string; customData: Record<string, string>; preferredStarterType?: string; }, members: { email: string; displayName: string; customData?: Record<string, string>; }[], teamName: string) => Promise<{ ok: boolean; teamId?: string; status?: "Angemeldet" | "Warteliste"; reason?: string; }>;
  salutation: "" | Salutation;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  selectedEventId: string;
  selectedSessions: Set<string>;
  sendBundledUpdateMail: (parentEvent: DeloitteEvent, recipientEmail: string, recipientName: string, items: BundledItem[]) => Promise<boolean>;
  sessionFieldValues: Record<string, Record<string, string>>;
  sessionMeta: Record<string, { count: number; wasRegistered: boolean; }>;
  /** v30.67: Abweichung zwischen Vorbelegung und Auswahl — eine leere Auswahl kann eine Abmeldung sein. */
  sessionsChanged: boolean;
  setAssistantModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setCcSelfModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setConfirmDialogAck: React.Dispatch<React.SetStateAction<boolean>>;
  setConfirmDialogOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setConfirmDraftParent: React.Dispatch<React.SetStateAction<boolean>>;
  setConfirmDraftSessions: React.Dispatch<React.SetStateAction<Set<string>>>;
  setError: React.Dispatch<React.SetStateAction<string>>;
  setExternalEmailWarning: React.Dispatch<React.SetStateAction<boolean>>;
  setFallbackDialog: React.Dispatch<React.SetStateAction<{ wunsch: string; alt: string; altFree: number; }>>;
  setIsSubmitting: React.Dispatch<React.SetStateAction<boolean>>;
  setSessionsOnlySubmitted: React.Dispatch<React.SetStateAction<boolean>>;
  setShowErrors: React.Dispatch<React.SetStateAction<boolean>>;
  setSubmitProgress: React.Dispatch<React.SetStateAction<number>>;
  setSubmitProgressLabel: React.Dispatch<React.SetStateAction<string>>;
  setSubmitted: React.Dispatch<React.SetStateAction<boolean>>;
  setSubmittedAsWaitlist: React.Dispatch<React.SetStateAction<boolean>>;
  setSubmittedJoinKind: React.Dispatch<React.SetStateAction<"joined" | "requested">>;
  showAlert: (message: React.ReactNode, opts?: import("../../context/DialogContext").AlertOptions) => void;
  /** v30.67: null, solange die Belegung nicht ermittelt ist — damit darf nicht gerechnet werden. */
  starterCounts: { durch: number; fun: number; durchWait: number; funWait: number; } | null;
  submittedSessionsRef: React.MutableRefObject<Set<string>>;
  subOnlyTerms: boolean;
  surname: string;
  t: (key: string) => string;
  teamMemberFields: Record<number, Record<string, string>>;
  teamMembersParsed: { displayName: string; email: string; }[];
  teamName: string;
  teamValidation: { ok: boolean; reason?: string; };
  thirdPartyCheck: { alreadyRegistered: boolean; notInAudience: boolean; registeredName?: string; registeredDate?: string; };
  updateMyRegistration: (eventId: string, customData: Record<string, string>) => Promise<boolean>;
  uploadFieldDocument: (eventId: string, fieldId: string, file: File, participantEmail?: string) => Promise<boolean>;
  userResults: { email: string; displayName: string; location: string; jobTitle: string; }[];
  userSearchIncludeIntl: boolean;
  willRegisterParent: boolean;
}

export interface SubmitFlow {
  /** Klick auf „Anmelden": validiert, stellt die Rueckfragen und waehlt den Weg. */
  handleSubmit: () => Promise<void>;
  /** Eigentliche Registrierung — auch aus dem Starter-Typ-Fallback-Dialog aufgerufen. */
  performRegistration: (starterTypeToUse: string) => Promise<void>;
}

export function createSubmitFlow(c: SubmitFlowCtx): SubmitFlow {
  const {
    assistantModalDecidedRef, cancelRegistration, canCreateEvents, canDelegateAssistant, ccSelfDecidedRef, ccSelfRef,
    childEvents, childOneDe, childTermPlural, childTermSingular, confirmDialog, confirmDialogConfirmedRef,
    createTeamJoinRequest, currentUser, delegateAssistEnabled, delegateAssistValue, delegateChoiceRef, delegateRegistrationToAssistant,
    durchCap, email, event, eventSpecific, externalEmailConfirmedRef, externalPerson,
    firstName, funCap, getEventNumbersForEmail, hiddenChildCount, isAllowedTargetForAssistant, isAssistant,
    isDeadlinePassed, isSplitGroup, isTeamMode, joinTeam, locale,
    nothingToSubmit, otherConsentConfirmed, parentAlreadyRegistered, parentFullNoWaitlist, parentRegBlocked, pendingDocFiles,
    pendingJoinTeam, preferredStarterType, previewAsUser, recordProxyDelegation, refreshEvents, registerForEvent,
    registerForOther, registerForParent, registerTeam, salutation, searchUsers, selectedEventId,
    selectedSessions, sendBundledUpdateMail, sessionFieldValues, sessionMeta, sessionsChanged, setAssistantModalOpen, setCcSelfModalOpen,
    setConfirmDialogAck, setConfirmDialogOpen, setConfirmDraftParent, setConfirmDraftSessions, setError, setExternalEmailWarning,
    setFallbackDialog, setIsSubmitting, setSessionsOnlySubmitted, setShowErrors, setSubmitProgress, setSubmitProgressLabel,
    setSubmitted, setSubmittedAsWaitlist, setSubmittedJoinKind, showAlert, starterCounts, submittedSessionsRef,
    subOnlyTerms, surname, t, teamMemberFields, teamMembersParsed, teamName,
    teamValidation, thirdPartyCheck, updateMyRegistration, uploadFieldDocument, userResults, userSearchIncludeIntl,
    willRegisterParent,
  } = c;

  const handleSubmit = async (): Promise<void> => {
    // v30.4: User-Vorschau ist NUR zum Ansehen — hier ist die letzte
    // gemeinsame Stelle vor jeder echten Anmeldung (selbst, für Dritte,
    // Team), deshalb sitzt die Sperre hier statt an jedem Button.
    if (previewAsUser) {
      setError(locale === 'de'
        ? 'Vorschau-Modus: Du siehst die Anmeldemaske gerade so, wie reguläre User sie sehen — Anmelden ist hier deaktiviert. Beende die Vorschau über den blauen Balken oben, um dich anzumelden.'
        : 'Preview mode: you are viewing the form as regular users see it — registering is disabled here. End the preview via the blue bar at the top to register.');
      return;
    }
    // v17.25: Demo-Showcase-Event — keine echte Anmeldung. Freundlicher
    // Hinweis statt SP-Roundtrip; der Context-Guard würde ohnehin no-oppen.
    if (event?.isDemoShowcase) {
      setError(locale === 'de'
        ? 'Dies ist ein Demo-Event — es wird keine echte Anmeldung gespeichert. Du kannst die Bereiche oben frei ausprobieren.'
        : 'This is a demo event — no real registration is stored. Feel free to explore the sections above.');
      return;
    }
    // v30.9: Auswahl fuer die Erfolgsseite einfrieren (s. submittedSessionsRef).
    submittedSessionsRef.current = new Set(selectedSessions);
    // Validierung Pflichtfelder
    setShowErrors(true);

    // Wenn der Haupt-Event-Checkbox aus ist und keine Session ausgewählt ist,
    // gibt es nichts zu tun.
    //
    // v28.88: …aber der GRUND muss zum Fall passen. „Bitte wähle mindestens das
    // Haupt-Event oder ein Sub-Event aus" setzt voraus, dass es etwas zu wählen
    // gibt. Bei einer bestehenden Anmeldung ist `willRegisterParent` immer
    // false (parentAlreadyRegistered, s.o.), und hat das Event keine
    // Sub-Events, steht auf der Seite überhaupt keine Auswahl. Wer bereits
    // angemeldet war und auf „Registrieren" klickte, bekam deshalb eine
    // Aufforderung, die ins Leere zeigt. Dasselbe im „Nur Sub-Events"-Modus und
    // bei gesperrtem Hauptevent (Frist abgelaufen / voll ohne Warteliste): Das
    // Haupt-Event ist dort gar nicht wählbar, die Meldung nannte es trotzdem.
    if (nothingToSubmit) {
      const oneSub = childTermSingular || (locale === 'de' ? 'Sub-Event' : 'sub-event');
      const hasSubs = childEvents.length > 0;
      if (parentAlreadyRegistered) {
        setError(locale === 'de'
          ? (hasSubs
            ? `Du bist für dieses Event bereits angemeldet. Möchtest du zusätzlich ${childOneDe} buchen, wähle es oben aus — abmelden kannst du dich über „Meine Events".`
            : 'Du bist für dieses Event bereits angemeldet — es gibt nichts weiter abzuschicken. Abmelden kannst du dich über „Meine Events".')
          : (hasSubs
            ? `You are already registered for this event. To additionally book a ${oneSub}, pick it above — you can cancel via „My events".`
            : 'You are already registered for this event — there is nothing further to submit. You can cancel via „My events".'));
        return;
      }
      if ((event && event.subEventsOnlyMode) || parentRegBlocked) {
        if (hasSubs) {
          // v29.13: Der Zusatz „das Haupt-Event ist hier nicht buchbar" hilft
          // nur, wenn es für den Teilnehmer sichtbar EIN Haupt-Event gibt. Im
          // reinen Sub-Event-Modus gibt es das nicht — dort ist die Auswahl
          // oben schlicht die Liste der Events, und der Nachsatz erfände eine
          // zweite Ebene, die nirgends auftaucht.
          setError(locale === 'de'
            ? (subOnlyTerms
              ? `Bitte wähle mindestens ${childOneDe} aus, um dich anzumelden.`
              : `Bitte wähle mindestens ${childOneDe} aus, um dich anzumelden — das Haupt-Event ist hier nicht buchbar.`)
            : (subOnlyTerms
              ? `Please select at least one ${oneSub} to register.`
              : `Please select at least one ${oneSub} to register — the main event cannot be booked here.`));
          return;
        }
        // v29.9: „keines angelegt" nur sagen, wenn wirklich keines existiert.
        // Sind welche da, aber keines für diese Person freigegeben, ist das der
        // Grund — und die Person soll wissen, dass sie nach einer Freigabe
        // fragen muss und nicht nach einem fehlenden Termin.
        setError(locale === 'de'
          ? (parentFullNoWaitlist
            ? 'Alle Plätze sind belegt und die Warteliste ist für dieses Event deaktiviert — eine Anmeldung ist nicht mehr möglich.'
            : (isDeadlinePassed
              ? 'Die Anmeldefrist dieses Events ist abgelaufen — eine Anmeldung ist nicht mehr möglich.'
              : (hiddenChildCount > 0
                ? 'Die Anmeldung läuft hier ausschließlich über die einzelnen Programmpunkte — für dich ist aktuell keiner davon freigegeben. Wenn du teilnehmen möchtest, wende dich bitte an die Organizer.'
                : 'Für dieses Event läuft die Anmeldung ausschließlich über Sub-Events — aktuell ist keines angelegt. Bitte wende dich an die Organizer.')))
          : (parentFullNoWaitlist
            ? 'All seats are taken and the waitlist is disabled for this event — registration is no longer possible.'
            : (isDeadlinePassed
              ? 'The registration deadline for this event has passed — registration is no longer possible.'
              : (hiddenChildCount > 0
                ? 'Registration here runs exclusively via the individual programme items — none of them is currently released for you. If you would like to attend, please contact the organizers.'
                : 'Registration for this event runs exclusively via sub-events — none exists yet. Please contact the organizers.'))));
        return;
      }
      setError(t('reg.nothing.selected') || 'Bitte wähle mindestens Haupt-Event oder eine Session aus.');
      return;
    }

    // v24.64: Pflicht-Sub-Events. Pro Sub-Event kann der Organizer im Wizard
    // („Sub-Events"-Schritt) „Pflichtanmeldung" setzen — ein so markiertes
    // Sub-Event MUSS ausgewählt sein, sonst ist die Anmeldung nicht möglich.
    // (Löst das alte, in der UI nicht mehr einstellbare
    // requireSubEventSelection ab — das wird hier bewusst NICHT mehr geprüft.)
    const subShortName = (c: { title?: string }): string => {
      const tt = (c.title || '').trim();
      const parts = tt.split('|');
      return (parts.length > 1 ? parts[parts.length - 1] : tt).trim();
    };
    const mandatoryMissing = childEvents.filter(c => c.mandatoryRegistration && !selectedSessions.has(c.id));
    if (mandatoryMissing.length > 0) {
      const names = mandatoryMissing.map(subShortName).filter(Boolean).join(', ');
      setError(locale === 'de'
        ? `Bitte wähle die Pflicht-${mandatoryMissing.length === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} aus, um dich anzumelden: ${names}.`
        : `Please select the mandatory ${mandatoryMissing.length === 1 ? (childTermSingular || 'sub-event') : (childTermPlural || 'sub-events')} to register: ${names}.`);
      return;
    }
    // v24.64: Im „Nur Sub-Events"-Modus ist das Haupt-Event nicht buchbar —
    // dann muss mindestens ein Sub-Event gewählt sein.
    // v30.67: … es sei denn, die leere Auswahl IST die Änderung: Wer alle
    // gebuchten Termine abwählt, will sich abmelden (`sessionsChanged`, seit
    // v28.88 bekannt). Der Guard nannte ihm bisher einen Grund („mindestens
    // ein Event auswählen"), der auf seine Absicht nicht zutraf — die letzte
    // Anmeldung war über die Anmeldeseite nicht kündbar.
    if (event && event.subEventsOnlyMode && childEvents.length > 0 && selectedSessions.size === 0 && !sessionsChanged) {
      setError(locale === 'de'
        ? `Bitte wähle mindestens ${childOneDe} aus, um dich anzumelden.`
        : `Please select at least one ${childTermSingular || 'sub-event'} to register.`);
      return;
    }

    // v15.16: Pflicht-Bestätigung bei „Für andere registrieren" —
    // analog zur Team-Anmeldung muss die Zustimmung der Person
    // explizit bestätigt werden.
    if (registerForOther && !otherConsentConfirmed) {
      setError(locale === 'de'
        ? 'Bitte bestätige, dass die Person ihrer stellvertretenden Anmeldung zugestimmt hat.'
        : 'Please confirm that the person has consented to this registration.');
      return;
    }

    // Basis-Felder sind immer Pflicht (Name + Email), auch im Sessions-Only-Modus.
    if (!firstName.trim() || !surname.trim() || !email.trim()) {
      setError(t('reg.requiredfields'));
      return;
    }
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) {
      setError(t('reg.invalidemail') || 'Ungültige E-Mail-Adresse');
      return;
    }

    // v19.6: Stellvertretende Anmeldung — verhindern, dass der Organizer
    // versehentlich SICH SELBST als „andere Person" einträgt. Im
    // „Für andere registrieren"-Modus MUSS eine andere Person ausgewählt sein.
    // Ist die Teilnehmer-E-Mail leer ODER identisch zur eigenen, läuft sonst die
    // Doppel-Anmelde-Prüfung gegen den eingeloggten User und meldet
    // irreführend „bereits angemeldet" (Beobachtung des Users: „bei einer
    // Anmeldung einer dritten Person wird weiterhin geprüft, ob ich selber
    // angemeldet werde"). Hier klar abbrechen statt still die Selbst-Prüfung
    // auszulösen.
    if (registerForOther && email.trim().toLowerCase() === (currentUser.email || '').toLowerCase()) {
      setError(t('reg.error.selfasother'));
      return;
    }

    // v19.8: Bereits angemeldete Zielperson hart blocken — kein CC-Modal, keine
    // Anmeldung. Greift auch, falls der Button-Disable wegen einer noch
    // laufenden Vorab-Prüfung kurz nicht aktiv war.
    if (registerForOther && thirdPartyCheck && thirdPartyCheck.alreadyRegistered) {
      setError(t('reg.error.other'));
      return;
    }

    // v18.74: Bei stellvertretender Anmeldung einer EXTERNEN Adresse zwei
    // Stufen: (1) strengere Plausibilitätsprüfung gegen Tippfehler (fehlende
    // TLD, doppelte Punkte, mehrere @, …) → harter Fehler; (2) ein
    // Bestätigungs-Dialog, der die Adresse groß anzeigt und zum Gegenlesen
    // auffordert (man kann an externe Adressen keinen Tippfehler korrigieren).
    if (registerForOther && email.trim() && isExternalEmailAddr(email)) {
      if (!isPlausibleEmail(email)) {
        setError(locale === 'de'
          ? 'Die externe E-Mail-Adresse sieht ungültig aus — bitte auf Tippfehler prüfen (z.B. fehlende Domain-Endung wie „.de" oder „.com").'
          : 'The external email address looks invalid — please check for typos (e.g. a missing domain ending like „.de" or „.com").');
        return;
      }
      if (!externalEmailConfirmedRef.current) {
        setExternalEmailWarning(true);
        return; // Bestätigungs-Dialog (Tippfehler-Gegenlesen) zeigen
      }
    }

    // Nur wenn der User das Haupt-Event (neu) anmelden möchte, gelten
    // Anrede + Custom-Fields + B2Run-Starter-Typ als Pflicht.
    // v18.53 BUG-FIX: Im subEventsOnlyMode wird das Hauptevent zwar nicht
    // direkt angemeldet, aber als „Schatten-Registrierung" mitgeschrieben
    // (s.u., shouldShadowRegisterParent) — inkl. der Hauptevent-Custom-Fields
    // (z.B. „Selection as above confirmed", Food Preferences, Hotel). Diese
    // werden in dem Modus trotzdem angezeigt und persistiert, also MÜSSEN sie
    // auch validiert werden. Vorher konnte man im Nur-Sub-Events-Modus mit
    // leerem Pflichtfeld absenden, weil `willRegisterParent` hier immer false
    // ist. Dieselbe Bedingung wie shouldShadowRegisterParent unten verwenden.
    const isSubOnlyModeValidate = !!(event && event.subEventsOnlyMode);
    // v18.57 BUG-FIX: Bedingung robust gemacht. Vorher hing sie an
    // `sessionsBeingAddedValidate` (= sessionMeta[...]?.wasRegistered), das in
    // manchen Re-Submit-/Reload-Fällen unzuverlässig war → Validierungsblock
    // wurde übersprungen → Pflichtfeld-Bypass. Jetzt reicht: subEventsOnlyMode
    // aktiv UND mindestens eine Section ausgewählt. Die übergreifenden
    // Hauptevent-Pflichtfelder (insb. die pro Absenden neu leere „Bestätigung"-
    // Checkbox) werden in dem Modus IMMER angezeigt und müssen IMMER validiert
    // werden. KEIN `!myParentReg`, KEIN sessionMeta-Abhängigkeit mehr.
    // v29.27: Pflichtfelder der ausgewählten Sub-Events prüfen — die Fragen
    // stehen jetzt inline in der Karte, also muss der Submit erzwingen, was
    // vorher das Bestätigen-Modal erzwungen hat. Bereits bestehende
    // Anmeldungen (wasRegistered) sind ausgenommen: ihre Antworten liegen in
    // der Teilnehmer-Zeile und werden hier nicht neu erfasst.
    {
      const subMissing: string[] = [];
      childEvents.forEach(ce => {
        if (!selectedSessions.has(ce.id)) return;
        if (sessionMeta[ce.id]?.wasRegistered) return;
        const values = sessionFieldValues[ce.id] || {};
        (ce.eventSpecificFields || []).filter(f => f && f.label && f.required && f.type !== 'document').forEach(f => {
          if (f.showIf && f.showIf.fieldId) {
            const raw = (values[f.showIf.fieldId] || '').trim();
            const answers = !raw ? [] : (raw.indexOf(' | ') >= 0 ? raw.split(' | ').map(s => s.trim()).filter(Boolean) : [raw]);
            if (!answers.some(a => f.showIf!.values.indexOf(a) >= 0)) return;
          }
          const filled = f.type === 'checkbox' ? values[f.id] === 'true' : !!(values[f.id] || '').trim();
          if (!filled) subMissing.push(`${ce.title || (locale === 'de' ? 'Sub-Event' : 'sub-event')}: ${f.label}`);
        });
      });
      if (subMissing.length > 0) {
        setError(`${t('reg.requiredcustom')}: ${subMissing.join(', ')}`);
        return;
      }
    }
    const willCollectMainFields = willRegisterParent || registerForOther
      || (isSubOnlyModeValidate && selectedSessions.size > 0 && !registerForOther)
      // v18.73: Beim vorgemerkten Team-Beitritt gelten dieselben Pflichtfelder
      // (Anrede + event-spezifische Felder) wie bei einer normalen Anmeldung —
      // der ganze Sinn der Änderung ist, dass diese Infos nicht übersprungen
      // werden können.
      || !!pendingJoinTeam;
    if (willCollectMainFields) {
      // v11.80: Anrede ist nur dann Pflichtfeld, wenn das Event das
      // Anrede-Dropdown auch tatsächlich abfragt (event.askSalutation === true).
      // Sonst wird die Anrede gar nicht gerendert und bleibt leer.
      if (event.askSalutation && !salutation) {
        setError(t('reg.requiredfields'));
        return;
      }

      // Pflicht-Custom-Fields validieren. Checkbox-Pflichtfelder müssen 'true' sein,
      // alle anderen dürfen nach trim nicht leer sein.
      // B2Run: Mobilnummer ist nur Pflicht wenn Infoservice aktiviert; ansonsten
      // gilt das Feld als versteckt und wird übersprungen.
      const missingRequired = event.eventSpecificFields
        .filter(f => {
          if (f.id === 'b2run_mobilnummer') {
            if (eventSpecific['b2run_infoservice'] !== 'true') return false;
            return !eventSpecific[f.id]?.trim();
          }
          // v7.21: Felder mit nicht erfüllter Sichtbarkeitsbedingung sind
          // ausgeblendet und dürfen die Validation nicht blockieren.
          if (f.showIf && f.showIf.fieldId) {
            const raw = (eventSpecific[f.showIf.fieldId] || '').trim();
            const answers = !raw
              ? []
              : raw.indexOf(' | ') >= 0
                ? raw.split(' | ').map(s => s.trim()).filter(Boolean)
                : [raw];
            const conditionMet = answers.some(a => f.showIf!.values.indexOf(a) >= 0);
            if (!conditionMet) return false;
          }
          // v10.24: Pro-Gruppe-Constraint — wenn das Feld auf eine andere
          // Gruppe als die aktuell vom User gewählte beschränkt ist, ist
          // es ausgeblendet und blockt die Validation nicht.
          if (f.onlyForGroup && f.onlyForGroup !== 'all' && isSplitGroup) {
            const want = f.onlyForGroup === 'A' ? 'Durchstarter' : 'Funstarter';
            if (preferredStarterType !== want) return false;
          }
          if (!f.required) return false;
          // v19.0: Pflicht-Dokument — bei NEUER Anmeldung muss eine Datei gewählt
          // sein. Bei bereits angemeldeter Person läuft die Verwaltung über
          // „Meine Events", daher dort nicht blockieren.
          if (f.type === 'document') return !parentAlreadyRegistered && !pendingDocFiles[f.id];
          return f.type === 'checkbox'
            ? eventSpecific[f.id] !== 'true'
            : !eventSpecific[f.id]?.trim();
        });
      if (missingRequired.length > 0) {
        setError(`${t('reg.requiredcustom')}: ${missingRequired.map(f => f.label).join(', ')}`);
        return;
      }

      // B2Run: Starter-Typ Pflichtfeld. v18.73: Beim Team-Beitritt NICHT
      // erzwingen — der Beitretende erbt die Gruppe des Teams (siehe
      // addTeamMember), er wählt sie nicht selbst.
      if (isSplitGroup && !preferredStarterType && !pendingJoinTeam) {
        // v11.7: generische Fehlermeldung — vorher hatte der Translation-Key
        // 'B2Run Starter-Typ' als Fallback. Bei generischen Split-Capacity-
        // Events mit eigenen Labels (z.B. 'Vormittag' / 'Nachmittag') passt
        // das nicht. Inline-Text mit den frei wählbaren Labels.
        const lblA = (event?.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
        const lblB = (event?.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
        setError(locale === 'de'
          ? `Bitte wähle eine der zwei Gruppen aus (${lblA} oder ${lblB}).`
          : `Please pick one of the two groups (${lblA} or ${lblB}).`);
        return;
      }

      // v6.15: Leistungsnachweis-Pflicht bei Durchstarter (Admin-Option)
      if (event.durchstarterRequiresProof && preferredStarterType === 'Durchstarter' && eventSpecific['b2run_leistungsnachweis'] !== 'true') {
        setError(t('reg.starter.proof.required') || 'Bitte Leistungsnachweis bestätigen.');
        return;
      }
    }

    // v11.10: Sub-Events erben preferredStarterType vom oberen
    // Group-Selection-Block. Wenn das Event Split-Capacity hat und Sessions
    // ausgewählt sind, muss eine Gruppe gewählt sein — egal ob Parent dabei
    // ist oder nur Sessions registriert werden.
    if (isSplitGroup && selectedSessions.size > 0 && !preferredStarterType && !pendingJoinTeam) {
      setError(locale === 'de'
        ? 'Bitte wähle eine Gruppe.'
        : 'Please pick a group.');
      return;
    }

    // Assistant-Ausnahme: defense-in-depth check beim Submit — Target muss
    // Partner oder Director sein. Der Fall tritt nur ein wenn der User weder
    // Organizer des Events noch Admin ist, aber via Assistant-Ausnahme für
    // eine andere Person registrieren will. JobTitle entweder aus dem zuletzt
    // geladenen Search-Result oder per Live-Lookup.
    if (registerForOther && isAssistant && !canCreateEvents) {
      try {
        const emailLc = email.trim().toLowerCase();
        let targetJobTitle = userResults.find(u => u.email.toLowerCase() === emailLc)?.jobTitle || '';
        if (!targetJobTitle) {
          const fresh = await searchUsers(email.trim(), userSearchIncludeIntl);
          targetJobTitle = fresh.find(u => u.email.toLowerCase() === emailLc)?.jobTitle || '';
        }
        if (!isAllowedTargetForAssistant(targetJobTitle)) {
          setError('As an Assistant you can only register Partners or Directors for events.');
          return;
        }
      } catch {
        setError('Unable to verify job title of selected person. Please select again from the dropdown.');
        return;
      }
    }

    // Seit v6.5: B2Run-Split-Kapazitäten. Wenn der gewählte Starter-Typ voll ist,
    // aber der andere Typ noch freie Plätze hat, zeigen wir einen Dialog —
    // der User entscheidet explizit:
    //   (a) auf den anderen Typ umsteigen, oder
    //   (b) auf die Warteliste für den gewünschten Typ.
    // Kein stiller Auto-Fallback mehr. Beide Typen voll → direkt auf Warteliste
    // (kein Dialog, Logik in EventContext setzt Status=Warteliste).
    // v30.67: `starterCounts` ist null, solange der Belegungs-Effect nicht
    // durch ist (oder keine Zahl liefern darf). Dann wird der Vorab-Dialog
    // übersprungen statt mit erfundenen Zahlen zu rechnen — vorher warf die
    // Zeile darunter einen TypeError in eine unbehandelte Promise-Rejection,
    // noch vor `setIsSubmitting(true)`: kein Overlay, keine Meldung, der
    // „Anmelden"-Knopf wirkte tot. Die harte Grenze zieht ohnehin
    // `reserveSeat` serverseitig.
    if ((willRegisterParent || registerForOther) && isSplitGroup && preferredStarterType && !pendingJoinTeam && starterCounts) {
      const durchFree = Math.max(0, durchCap - starterCounts.durch);
      const funFree = Math.max(0, funCap - starterCounts.fun);
      const wunschFree = preferredStarterType === 'Durchstarter' ? durchFree : funFree;
      const altType = preferredStarterType === 'Durchstarter' ? 'Funstarter' : 'Durchstarter';
      const altFree = altType === 'Durchstarter' ? durchFree : funFree;
      if (wunschFree === 0 && altFree > 0) {
        // Dialog zeigen, Submit warten.
        setFallbackDialog({ wunsch: preferredStarterType, alt: altType, altFree });
        return;
      }
    }

    // v18.73: Vorgemerkter Team-Beitritt — committet hier (mit den oben
    // ausgefüllten persönlichen + event-spezifischen Feldern), statt einer
    // normalen Einzel-Anmeldung. Steht vor dem Team-Anmelde-Pfad, weil beide
    // sich gegenseitig ausschließen.
    if (pendingJoinTeam) {
      await performJoinSelectedTeam();
      return;
    }

    // v11.82: Team-Anmeldung — separater Submit-Pfad.
    if (isTeamMode) {
      if (!teamValidation.ok) {
        setError(teamValidation.reason || (locale === 'de' ? 'Team-Anmeldung unvollständig.' : 'Team registration incomplete.'));
        return;
      }
      await performTeamRegistration(preferredStarterType);
      return;
    }

    // v18.75: Sicherheitshinweis vor dem Absenden (nur normale Anmeldung —
    // Team-/Beitritts-Pfade sind oben bereits abgehandelt). Beim ersten
    // Submit öffnet sich der Dialog; nach Bestätigung läuft handleSubmit erneut
    // (Ref gesetzt) und überspringt den Dialog.
    if (event && event.confirmDialogEnabled && !confirmDialogConfirmedRef.current) {
      setConfirmDraftParent(willRegisterParent || registerForOther);
      setConfirmDraftSessions(new Set(selectedSessions));
      setConfirmDialogAck(false);
      setConfirmDialogOpen(true);
      return;
    }

    // v19.6: Stellvertretende Anmeldung einer INTERNEN Person — fragen, ob
    // der/die Anmeldende selbst auf CC der Bestätigungs-Mail soll. Nur einmal
    // pro Submit-Durchlauf (ccSelfDecidedRef); externe Personen sind ausgenommen
    // (dort ist der Organizer-Kreis ohnehin schon auf CC). Nach der Entscheidung
    // läuft handleSubmit erneut und überspringt das Modal.
    if (registerForOther && !externalPerson && !ccSelfDecidedRef.current) {
      setCcSelfModalOpen(true);
      return;
    }

    // v24.48: Assistenz-Abfrage als Modal beim Register-Klick — nur für
    // Partner/Director (canDelegateAssistant), einmal pro Submit-Durchlauf.
    // Nach der Entscheidung läuft handleSubmit erneut und überspringt das Modal.
    if (canDelegateAssistant && !assistantModalDecidedRef.current) {
      setAssistantModalOpen(true);
      return;
    }

    await performRegistration(preferredStarterType);
  };

  // v11.82: Team-Anmeldung absenden — Lead + N-1 Mitglieder per registerTeam.
  // Sub-Events werden im Team-Modus NICHT mitangemeldet — das ist Phase 2
  // (siehe Manual). Lead und Member bekommen jeweils nur den Hauptevent.
  const performTeamRegistration = async (starterTypeToUse: string): Promise<void> => {
    setError('');
    // v23.2: Harter Riegel gegen Doppel-Anmeldung über den Team-Pfad. Wer als
    // eingeloggte Person bereits aktiv beim Event angemeldet ist (solo oder in
    // einem anderen Team), darf NICHT erneut ein Team anlegen — der Solo-Pfad
    // ist seit jeher so abgesichert, der Team-Pfad war es nicht (Ursache der
    // Doppel-Anmeldung bei Team-Events). Die Team-Karte ist in dem Fall bereits
    // ausgeblendet; das hier ist das Sicherheitsnetz, falls der Status erst
    // nach dem Aufklappen geladen wurde.
    if (parentAlreadyRegistered) {
      setError(locale === 'de'
        ? 'Du bist bereits für dieses Event angemeldet — eine zusätzliche Team-Anmeldung ist nicht möglich. Bitte zuerst über „Meine Events" abmelden, falls du in ein anderes Team wechseln möchtest.'
        : 'You are already registered for this event — an additional team registration is not possible. Please cancel via „My Events" first if you want to switch to another team.');
      return;
    }
    setIsSubmitting(true);
    setSubmitProgress(5);
    setSubmitProgressLabel(locale === 'de' ? 'Team-Anmeldung wird vorbereitet…' : 'Preparing team registration…');
    try {
      const customData: Record<string, string> = { salutation, ...stripPrefilterKeys(eventSpecific) };
      const leadEmail = email.trim();
      const leadFirstName = firstName.trim();
      const leadLastName = surname.trim();
      // v18.12: Custom-Field-Antworten pro Mitglied (nach Slot-Index) mitgeben.
      const members = teamMembersParsed
        .map((m, idx) => m ? { email: m.email, displayName: m.displayName, customData: stripPrefilterKeys(teamMemberFields[idx] || {}) } : null)
        .filter((m): m is { displayName: string; email: string; customData: Record<string, string> } => !!m);
      setSubmitProgress(30);
      setSubmitProgressLabel(locale === 'de'
        ? `Team wird angemeldet (${1 + members.length} Personen)…`
        : `Registering team (${1 + members.length} people)…`);
      const result = await registerTeam(
        selectedEventId!,
        {
          firstName: leadFirstName,
          lastName: leadLastName,
          email: leadEmail,
          salutation,
          customData,
          preferredStarterType: starterTypeToUse || undefined,
        },
        members,
        event?.askTeamName ? (teamName.trim() || undefined) : undefined
      );
      setSubmitProgress(90);
      if (!result.ok) {
        if (result.reason && result.reason.indexOf('already-registered:') === 0) {
          const dupEmail = result.reason.substring('already-registered:'.length);
          setError(locale === 'de'
            ? `Person bereits angemeldet: ${dupEmail}. Bitte aus dem Team entfernen und erneut versuchen.`
            : `Person already registered: ${dupEmail}. Please remove from the team and try again.`);
        } else if (result.reason && result.reason.indexOf('partial-insert:') === 0) {
          // v30.67: Ein Teilerfolg ist keine "fehlgeschlagene Team-Anmeldung" —
          // die uebrigen Personen SIND angemeldet. Wer hier "Fehler" liest und
          // es noch einmal versucht, erzeugt Doppel-Anmeldungen.
          const failed = result.reason.substring('partial-insert:'.length);
          setError(locale === 'de'
            ? `Für folgende Personen ist die Anmeldung nicht durchgekommen: ${failed} — bitte einzeln nachmelden. Die übrigen Teammitglieder sind angemeldet.`
            : `Registration did not go through for: ${failed} — please register them individually. The other team members are registered.`);
        } else if (result.reason === 'dup-check-failed') {
          setError(locale === 'de'
            ? 'Die Prüfung auf bereits angemeldete Personen konnte nicht durchgeführt werden — es wurde niemand angemeldet. Bitte später erneut versuchen.'
            : 'The check for already registered people could not be run — nobody was registered. Please try again later.');
        } else {
          setError(t('reg.error') || (locale === 'de' ? 'Fehler bei der Team-Anmeldung.' : 'Team registration failed.'));
        }
        return;
      }
      setSubmitted(true);
    } catch {
      setError(t('reg.genericerror') || (locale === 'de' ? 'Unerwarteter Fehler.' : 'Unexpected error.'));
    } finally {
      setSubmitProgress(100);
      setTimeout(() => {
        setIsSubmitting(false);
        setSubmitProgress(0);
        setSubmitProgressLabel('');
      }, 250);
    }
  };

  // v18.73: Vorgemerkten Team-Beitritt absenden — committet erst hier (auf
  // „Anmelden"), inkl. der ausgefüllten event-spezifischen Felder. Direkter
  // Beitritt (joinTeam) bzw. Anfrage an den Lead (createTeamJoinRequest), je
  // nach event.teamJoinRequiresApproval.
  const performJoinSelectedTeam = async (): Promise<void> => {
    if (!pendingJoinTeam || !event) return;
    setError('');
    setIsSubmitting(true);
    setSubmitProgress(10);
    setSubmitProgressLabel(locale === 'de' ? 'Beitritt wird verarbeitet…' : 'Processing your join…');
    try {
      // Event-spezifische Antworten des Beitretenden (wie bei der normalen
      // Anmeldung) — werden an den Team-Beitritt durchgereicht.
      const customData: Record<string, string> = { salutation, ...stripPrefilterKeys(eventSpecific) };
      setSubmitProgress(50);
      if (event.teamJoinRequiresApproval) {
        const r = await createTeamJoinRequest(event.id, pendingJoinTeam.teamId, customData);
        if (!r.ok) {
          setError(r.reason === 'already-registered'
            ? (locale === 'de' ? 'Du bist bereits beim Event angemeldet.' : 'You are already registered for this event.')
            : (locale === 'de' ? 'Beitritts-Anfrage fehlgeschlagen.' : 'Join request failed.'));
          return;
        }
        setSubmittedJoinKind('requested');
        setSubmitted(true);
      } else {
        const r = await joinTeam(event.id, pendingJoinTeam.teamId, pendingJoinTeam.teamName, customData);
        if (!r.ok) {
          setError(r.reason && r.reason.startsWith('already-registered')
            ? (locale === 'de' ? 'Du bist bereits beim Event angemeldet.' : 'You are already registered for this event.')
            : r.reason === 'team-full'
              ? (locale === 'de' ? 'Das Team ist inzwischen voll.' : 'The team has filled up in the meantime.')
              : (locale === 'de' ? 'Beitritt fehlgeschlagen.' : 'Joining failed.'));
          return;
        }
        setSubmittedJoinKind('joined');
        setSubmittedAsWaitlist(r.status === 'Warteliste');
        setSubmitted(true);
      }
    } catch {
      setError(locale === 'de' ? 'Unerwarteter Fehler beim Beitritt.' : 'Unexpected error while joining.');
    } finally {
      setSubmitProgress(100);
      setTimeout(() => { setIsSubmitting(false); setSubmitProgress(0); setSubmitProgressLabel(''); }, 250);
    }
  };

  // Eigentliche Registrierung — entkoppelt vom Validation/Submit-Trigger,
  // damit sie auch vom Fallback-Dialog aufgerufen werden kann (mit ggf. geändertem Starter-Typ).
  const performRegistration = async (starterTypeToUse: string): Promise<void> => {
    setError('');
    setIsSubmitting(true);
    setSubmitProgress(5);
    setSubmitProgressLabel(locale === 'de' ? 'Anmeldung wird vorbereitet…' : 'Preparing registration…');
    try {
      const customData: Record<string, string> = {
        salutation,
        ...stripPrefilterKeys(eventSpecific),
      };
      const participantEmail = email.trim();
      const firstTrim = firstName.trim();
      const surnameTrim = surname.trim();

      // v19.6: CC-Wunsch bei stellvertretender INTERNER Anmeldung. Wenn der/die
      // Anmeldende im Modal „Ja, auf CC" gewählt hat, landet die eigene E-Mail
      // als CC auf der Bestätigungs-Mail (NICHT im Outlook-Termin — wie bei den
      // Feld-CCs). Externe Anmeldungen sind ausgenommen.
      const ccSelfEmail = (registerForOther && !externalPerson && ccSelfRef.current)
        ? (currentUser.email || '').trim()
        : '';

      // v24.41: Delegation an die eigene Assistenz (nur Selbst-Anmeldung von
      // Admin/Director, Assistenz im Picker gewählt). Die Assistenz kommt auf CC
      // der Bestätigung; die eigentliche Zugriffs-Übergabe (Zeilen-Autor) läuft
      // nach erfolgreicher Anmeldung über delegateRegistrationToAssistant.
      // v24.49: Auswahl aus dem Ref (synchron gesetzt) lesen, Fallback auf State.
      const choice = delegateChoiceRef.current;
      const chosenVal = choice ? choice.value : (delegateAssistEnabled ? delegateAssistValue : '');
      const chosenEnabled = choice ? choice.enabled : delegateAssistEnabled;
      const chosenParsed = (() => {
        const m = (chosenVal || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
        return m ? { name: m[1].trim(), email: m[2].trim() } : null;
      })();
      const delegateAssist = (canDelegateAssistant && chosenEnabled && chosenParsed
        && chosenParsed.email.toLowerCase() !== participantEmail.toLowerCase())
        ? chosenParsed : null;
      const delegateCc = delegateAssist ? delegateAssist.email : '';

      let anySuccess = false;
      let parentOk = true;
      let lastSubReason: string | undefined;
      // v29.48: Schatten-/Klammer-Zeile konnte nicht angelegt werden (s.
      // doParentRegistration). Die Sub-Event-Anmeldungen bleiben gültig, aber
      // die übergreifenden Hauptevent-Antworten fehlen — das muss gesagt werden.
      let shadowParentFailed = false;
      // v23.10: Assistenz-Proxy-Anmeldung — der Client hat (Picker-Greyout +
      // Submit-Validierung oben) bereits sichergestellt, dass eine Assistenz nur
      // Partner/Director anmeldet. Dieses Flag wird der Registrierung als
      // vertrauenswürdig durchgereicht, damit die fragile serverseitige
      // Job-Title-Ableitung legitime Assistenzen nicht mehr fälschlich ablehnt.
      const actorAllowedAsAssistant = registerForOther && isAssistant && !canCreateEvents;
      // v23.9: Übersetzt den konkreten Misserfolgs-Grund aus registerForEvent in
      // eine verständliche Meldung — statt pauschal „bereits registriert".
      const regFailMessage = (reason?: string): string => {
        if (reason === 'not-allowed') {
          return locale === 'de'
            ? 'Du bist nicht berechtigt, diese Person für dieses Event anzumelden. Bitte wende dich an die Organizer des Events.'
            : 'You are not allowed to register this person for this event. Please contact the event organizers.';
        }
        if (reason === 'deadline') {
          return locale === 'de'
            ? 'Die Anmeldefrist dieses Events ist abgelaufen — eine Anmeldung ist nicht mehr möglich. Organizer und Admins können nach Ablauf weiterhin anmelden.'
            : 'The registration deadline for this event has passed — registration is no longer possible. Organizers and admins can still register after the deadline.';
        }
        // v30.14: Duplikat-Prüfung selbst gescheitert (Drosselung) — die
        // Anmeldung wird bewusst NICHT durchgeführt, statt eine Doppel-Zeile
        // zu riskieren.
        if (reason === 'dup-check-failed') {
          return locale === 'de'
            ? 'SharePoint ist gerade stark ausgelastet — wir konnten nicht sicher prüfen, ob diese Anmeldung schon existiert, und haben deshalb NICHT angemeldet (keine Doppel-Anmeldung riskieren). Bitte versuche es in ein paar Minuten erneut.'
            : 'SharePoint is currently under heavy load — we could not verify whether this registration already exists, so nothing was registered (to avoid a duplicate). Please try again in a few minutes.';
        }
        if (reason === 'insert-failed') {
          return locale === 'de'
            ? 'Die Anmeldung konnte nicht gespeichert werden (technischer Fehler an der Teilnehmerliste). Bitte erneut versuchen; hält es an, im Organizer Center „Spalten fixen" ausführen.'
            : 'The registration could not be saved (technical error on the participant list). Please try again; if it persists, run „Fix columns" in the organizer center.';
        }
        // v27.11: aktiver Duplikat-Treffer (deckt jetzt auch Externe ab).
        if (reason === 'already-registered') {
          return registerForOther
            ? t('reg.thirdparty.alreadyregistered')
            : (locale === 'de'
              ? 'Du bist für dieses Event bereits angemeldet.'
              : 'You are already registered for this event.');
        }
        // v27.11: Event voll und Warteliste vom Organizer abgeschaltet.
        if (reason === 'full') {
          return locale === 'de'
            ? 'Alle Plätze sind belegt und die Warteliste ist für dieses Event deaktiviert — eine Anmeldung ist nicht mehr möglich.'
            : 'All seats are taken and the waitlist is disabled for this event — registration is no longer possible.';
        }
        // Fallback (unbekannt / kein Grund) — bisherige generische Meldung.
        return t(registerForOther ? 'reg.error.other' : 'reg.error');
      };

      // Verfeinerte Progress-Stufen je nach Anzahl Sub-Events:
      // - parent: 5 → 30 → 50 (wenn Parent-Anmeldung lief)
      // - sub-events: 50 → 90 (gleichmäßig verteilt)
      // - finalize: 90 → 100
      const subOps = childEvents.filter(ce => {
        const wasReg = sessionMeta[ce.id]?.wasRegistered;
        const isSel = selectedSessions.has(ce.id);
        return (isSel && !wasReg) || (!isSel && wasReg && !registerForOther);
      }).length;

      // 1) Haupt-Event anmelden (nur wenn Checkbox an und noch nicht angemeldet).
      // v15.25: Im subEventsOnlyMode wird die Parent-Anmeldung trotzdem
      // durchgeführt — als „Schatten-Registrierung" rein zur Daten-
      // Vollständigkeit. Damit hat jeder Teilnehmer auch im Parent-
      // Teilnehmer-Schema eine Zeile mit den Antworten auf die Hauptevent-
      // Custom-Fields (Food Preferences, Hotel, Travel etc.). Mails +
      // Outlook werden für diese Schatten-Anmeldung in EventContext
      // unterdrückt — der User soll keine Bestätigung fürs Hauptevent
      // bekommen, da er da gar nicht „teilnimmt", sondern nur für Sub-
      // Events.
      const isSubOnlyMode = !!(event && event.subEventsOnlyMode);
      const sessionsBeingAdded = childEvents.some(ce => selectedSessions.has(ce.id) && !sessionMeta[ce.id]?.wasRegistered);
      // v30.67: Nicht „gibt es eine Zeile", sondern „gibt es eine AKTIVE Zeile
      // der Person, die hier angemeldet wird". Zwei Fälle liefen daran vorbei:
      // (1) Eine ABGEMELDETE Klammer-Zeile zählte als vorhanden — die Klammer
      //     wurde nie reaktiviert, `updateMyRegistration` (unten) schrieb die
      //     frischen Antworten in die abgemeldete Zeile, und die Hotelplanung
      //     (filtert auf aktive Stati) kannte die Person nicht mehr, obwohl
      //     sie „Yes, I need accommodation" beantwortet hatte.
      // (2) `myParentReg` ist IMMER die Zeile des eingeloggten Users; im
      //     Stellvertreter-Modus sagt sie nichts über die Zielperson. Hatte
      //     sich die Organizerin selbst angemeldet, bekam die andere Person
      //     Termin-Zeilen, aber keine Klammer — und weil diese Seite
      //     `skipShadowParent` setzt, zog auch das zentrale Netz nicht nach.
      // `registerForEvent` ist für die Klammer idempotent (aktive Zeile →
      // ok, kein zweiter Insert) und reaktiviert eine abgemeldete Zeile samt
      // CustomData — der Schritt-3-Block darf also in beiden Fällen laufen.
      const parentAlreadyHasRow = !registerForOther && parentAlreadyRegistered;
      // v26.67 (B): deckt jetzt Selbst- UND Fremd-Anmeldung ab (das frühere
      // `!registerForOther` entfällt — die Klammer läuft im subEventsOnly-Modus
      // in beiden Fällen ZUM SCHLUSS über den Schritt-3-Block unten).
      const shouldShadowRegisterParent = isSubOnlyMode && sessionsBeingAdded && !parentAlreadyHasRow;

      // v30.57: Den Nachzug-Merker SCHARFSTELLEN, BEVOR der erste Sub-Event
      // geschrieben wird — nicht erst, wenn die Klammer-Zeile scheitert.
      //
      // Bis hier entstand er nur im Fehlerzweig von `doParentRegistration`.
      // Das deckt den Fall ab, dass SharePoint den Schreibvorgang ablehnt —
      // aber nicht den Fall, den die Fehlermeldung im Organizer Center selbst
      // nennt: die ABGEBROCHENE Anmeldung. Wer den Tab schließt, das WLAN
      // verliert oder auf dem Handy die App wechselt, während die 19 Termine
      // geschrieben werden, kommt nie bis zum Klammer-Schritt. Dann gibt es
      // keinen Fehler, keinen Merker und keine Selbstheilung — nur eine Person
      // mit Sub-Event-Zeilen ohne Klammer-Zeile. Genau das Bild aus dem
      // roten Kasten.
      //
      // Umgekehrt scharfstellen kostet nichts: Der Merker wird unten wieder
      // entfernt, sobald die Klammer-Zeile steht, und ein überzähliger Nachzug
      // ist folgenlos — `registerForEvent` ist für die Klammer idempotent.
      if (shouldShadowRegisterParent) {
        try {
          addPendingShadowParent({
            eventId: selectedEventId!,
            customData: { ...customData },
            firstName: firstTrim,
            lastName: surnameTrim,
            email: participantEmail,
            proxy: registerForOther,
            ts: Date.now(),
          });
        } catch { /* localStorage gesperrt → es bleibt beim bisherigen Netz */ }
      }

      // v28.22: UNSICHTBARE Doppel-Anmeldung abfangen.
      //
      // Die Teilnehmerlisten laufen mit Item-Level-Security („nur eigene
      // Elemente", geprüft am Zeilen-AUTOR). Meldet eine Assistenz jemanden an,
      // bleibt sie Autor der Zeile, solange der nachträgliche Autor-Wechsel
      // mangels Rechten scheitert (Contribute reicht dafür nicht; der
      // DEX_AccessFix-Flow bzw. der Admin-Auto-Fix zieht ihn erst später nach).
      // Bis dahin ist die Zeile für die betroffene Person UNSICHTBAR — weder in
      // „Meine Events" noch für den Vorab-Check beim Anmelden. Der Check läuft
      // fail-open (lieber eine Zeile zu viel als eine blockierte Anmeldung) und
      // legte deshalb eine ZWEITE Anmeldung an.
      //
      // Gegenmittel: zusätzlich DEX_Participants fragen. Die Liste liegt auf der
      // Haupt-Site, kennt keine Item-Level-Security und wird bei JEDER An-/
      // Abmeldung mitgeschrieben — sie sieht also auch fremd angelegte Zeilen.
      // Bewusst nur eine RÜCKFRAGE, keine harte Sperre: Sollte der Eintrag mal
      // veraltet sein (Abmeldung ohne erfolgreiches Nachziehen), bleibt eine
      // legitime Anmeldung möglich.
      const hiddenDupTitles: string[] = [];
      try {
        const nums = await getEventNumbersForEmail(participantEmail);
        const knownNumbers = new Set<number>([...nums.registered, ...nums.waitlisted]);
        const isKnown = (n?: number): boolean => typeof n === 'number' && n > 0 && knownNumbers.has(n);
        // Hauptevent/Klammer: nur prüfen, wenn wir jetzt wirklich eine Zeile
        // anlegen würden und uns keine sichtbare bekannt ist.
        const willTouchParent = (registerForParent && !parentAlreadyRegistered) || shouldShadowRegisterParent;
        if (willTouchParent && !parentAlreadyHasRow && isKnown(event.eventNumber)) {
          hiddenDupTitles.push(event.title);
        }
        for (const ce of childEvents) {
          if (!selectedSessions.has(ce.id)) continue;
          if (sessionMeta[ce.id]?.wasRegistered) continue;
          if (isKnown(ce.eventNumber)) hiddenDupTitles.push(ce.title);
        }
      } catch { /* best-effort — im Zweifel wie bisher weiter */ }
      if (hiddenDupTitles.length > 0) {
        const list = hiddenDupTitles.map(x => `• ${x}`).join('\n');
        const who = registerForOther ? (`${firstTrim} ${surnameTrim}`.trim() || participantEmail) : '';
        const proceed = await confirmDialog(
          locale === 'de'
            ? (registerForOther
              ? `${who} ist laut unseren Daten hier bereits angemeldet:\n\n${list}\n\nMöglicherweise hat sich die Person selbst angemeldet oder eine andere Assistenz hat das übernommen — dann siehst du die Zeile wegen der Zugriffsrechte auf der Teilnehmerliste nicht. Eine erneute Anmeldung würde einen ZWEITEN Platz belegen.\n\nTrotzdem anmelden?`
              : `Du bist laut unseren Daten hier bereits angemeldet:\n\n${list}\n\nMöglicherweise hat dich jemand angemeldet (z.B. deine Assistenz) — dann siehst du die Anmeldung wegen der Zugriffsrechte auf der Teilnehmerliste nicht in „Meine Events". Eine erneute Anmeldung würde einen ZWEITEN Platz belegen.\n\nTrotzdem anmelden?`)
            : (registerForOther
              ? `According to our records ${who} is already registered for:\n\n${list}\n\nThe person may have registered themselves, or another assistant did it — in that case the row is hidden from you by the attendee list's permissions. Registering again would take a SECOND seat.\n\nRegister anyway?`
              : `According to our records you are already registered for:\n\n${list}\n\nSomeone may have registered you (e.g. your assistant) — in that case the attendee list's permissions hide it from „My events". Registering again would take a SECOND seat.\n\nRegister anyway?`),
          { danger: true, confirmLabel: locale === 'de' ? 'Trotzdem anmelden' : 'Register anyway' },
        );
        if (!proceed) {
          setIsSubmitting(false);
          setSubmitProgress(0);
          setSubmitProgressLabel('');
          return;
        }
      }
      // v26.67 (B): Gemeinsame Klammer-/Parent-Anmelde-Routine. `bestEffort` =
      // true bei der subEventsOnly-Schatten-Zeile, die JETZT NACH den Sub-Events
      // angelegt wird — ein Fehlschlag darf die (gültigen) Sub-Event-Anmeldungen
      // nicht als Fehler markieren.
      // v30.61: Was in dieser Anmeldung tatsächlich gebucht wurde — Grundlage
      // der Terminliste in der Sammelmail. Gesammelt wird NUR bei Erfolg; eine
      // Zeile für einen Termin, der nicht geschrieben wurde, wäre eine Zusage,
      // die die Daten nicht decken.
      //
      // Steht HIER und nicht bei der Sub-Event-Schleife: `doParentRegistration`
      // schließt darüber, und bei normalen Events (nicht subEventsOnly) läuft
      // die Klammer-Anmeldung VOR der Schleife — eine spätere Deklaration wäre
      // dort ein TDZ-Fehler (dieselbe Falle wie v29.71).
      const bookedItems: BundledItem[] = [];
      // Gebündelt wird nur im Klammer-Modus. Bei einem normalen Event mit
      // Sub-Events ist das Hauptevent selbst buchbar und verschickt ohnehin
      // seine eigene Mail — „statt zehn eine" gibt es dort nicht zu holen.
      const bundledMode = (event && event.subEventsOnlyMode)
        ? bundledCommOf(event)
        : BUNDLED_COMM_DEFAULT;
      const doParentRegistration = async (bestEffort: boolean): Promise<void> => {
        const parentResult = await registerForEvent(
          selectedEventId!,
          customData,
          firstTrim,
          surnameTrim,
          participantEmail,
          starterTypeToUse || undefined,
          // v18.74: Bei stellvertretender Anmeldung den Zustimmungs-Nachweis
          // mitschreiben (Pflicht-Checkbox wurde oben validiert).
          // v19.6: ccSelfEmail (Anmeldende:r auf CC der Bestätigungs-Mail).
          // v30.9: skipReload — der 28-MB-Volllade-Refresh läuft EINMAL am
          // Ende des Absendens (fire-and-forget), nicht nach jedem Schreiben.
          {
            ...(registerForOther
              ? { proxyConsentConfirmed: true, actorAllowedAsAssistant, ...(ccSelfEmail ? { extraCc: ccSelfEmail } : {}) }
              : { ...(delegateCc ? { extraCc: delegateCc } : {}) }),
            skipReload: true,
            // v30.61: Bei gebündelter Kommunikation trägt die Klammer-Mail die
            // Liste der gebuchten Termine. Nur DIESE Stelle kennt sie — der
            // EventContext sieht immer nur ein einzelnes Event.
            ...(bundledMode.mail && bookedItems.length > 0 ? { bundledItems: bookedItems } : {}),
          }
        );
        if (bestEffort) {
          // Schatten-Klammer: Erfolg zählt mit, aber kein Fehler-Durchschlag.
          if (parentResult.ok) { anySuccess = true; return; }
          // v29.48: … der Fehlschlag darf aber auch nicht spurlos verschwinden.
          // Genau hier entsteht „Fehlende Klammer-Anmeldung" im Organizer
          // Center: Die Schattenzeile ist der LETZTE Schreibvorgang einer
          // Anmeldung, also der, bei dem das SharePoint-Kontingent am ehesten
          // erschöpft ist (HTTP 429, s. utils/spThrottle). Bis v29.47 wurde das
          // Ergebnis verworfen, der Teilnehmer sah „Anmeldung erfolgreich",
          // und die übergreifenden Hauptevent-Antworten fehlten still.
          shadowParentFailed = true;
          return;
        }
        parentOk = parentResult.ok;
        if (parentOk) {
          anySuccess = true;
          // v18.67: echten Status fürs Ergebnis-Modal merken (nicht isFull).
          setSubmittedAsWaitlist(parentResult.status === 'Warteliste');
        }
        // v23.9: KONKRETE Fehlermeldung statt pauschal „bereits registriert" —
        // der echte Grund (Berechtigung / Deadline / technischer Fehler) wird
        // jetzt aus registerForEvent durchgereicht.
        else setError(regFailMessage(parentResult.reason));
      };
      // v26.67 (B) BUG-FIX: Im subEventsOnly-Modus ist die „Parent"-Anmeldung nur
      // eine Schatten-/Klammer-Zeile (Daten-Vollständigkeit) — sie wird jetzt ZUM
      // SCHLUSS angelegt (siehe Schritt 3), nachdem mind. ein Sub-Event steht.
      // Bricht der Vorgang vorher ab, ist die Person sichtbar in ihren Sub-Events
      // angemeldet statt als unsichtbarer, blockierender „Geist" zurückzubleiben.
      // Bei NORMALEN Events (nicht subEventsOnly) bleibt der Parent die eigentliche
      // Anmeldung und läuft weiterhin ZUERST. (willRegisterParent ist im
      // subEventsOnly-Modus immer false.)
      if (!isSubOnlyMode && (willRegisterParent || registerForOther)) {
        setSubmitProgress(30);
        setSubmitProgressLabel(locale === 'de' ? 'Haupt-Event wird angemeldet…' : 'Registering for main event…');
        await doParentRegistration(false);
        setSubmitProgress(50);
      } else if (isSubOnlyMode && parentAlreadyHasRow && sessionsBeingAdded && !registerForOther) {
        // v18.59: Die Schatten-Parent-Zeile existiert bereits (frühere
        // Section-Anmeldung) → sie wird NICHT neu registriert. Trotzdem die
        // übergreifenden Hauptevent-Antworten (Food Preferences, Hotel etc.)
        // mit den aktuellen Formular-Werten aktualisieren, damit Änderungen
        // beim Nach-Anmelden einer weiteren Section persistiert werden. Vorher
        // gingen sie verloren (Audit-Befund #2).
        setSubmitProgress(40);
        try { await updateMyRegistration(selectedEventId!, customData); } catch { /* best-effort */ }
        setSubmitProgress(50);
      } else {
        setSubmitProgress(50);
      }

      // 2) Ausgewählte Sessions an-/abmelden (unabhängig vom Parent).
      //    - Session ausgewählt + nicht angemeldet → anmelden
      //    - Session nicht ausgewählt + angemeldet → abmelden
      //    - Starter-Typ: wenn der User sich gleichzeitig fürs Haupt-Event anmeldet,
      //      wird dessen Starter-Typ auch auf die Session-Teilnehmerliste geschrieben
      //      (shared) — sonst die pro-Session-Auswahl. So steht in der TN-Liste jeder
      //      Session korrekt, ob der Teilnehmer Durchstarter oder Funstarter ist.
      // v11.10: Sub-Events erben grundsätzlich preferredStarterType
      // (bzw. starterTypeToUse vom Fallback-Dialog). Vorher hingen sie
      // an sessionStarterType pro-Session, was zu redundanten UI-Radios
      // pro Sub-Event geführt hat.
      const inheritedStarterType = starterTypeToUse || preferredStarterType || '';
      // v18.53: Im subEventsOnlyMode sind die Hauptevent-CC-Felder (z.B.
      // „Your assistant") übergreifend — sie gelten für die Sub-Events. Daher
      // die CC einmal aus dem Hauptformular ziehen und an jede Sub-Event-
      // Anmeldung mitgeben, damit deren Bestätigungsmails ebenfalls an die
      // Assistenz auf CC gehen (das „Hauptevent" ist nicht anmeldbar und seine
      // Schatten-Registrierung verschickt keine Mail).
      const crossCutCc = (event && event.subEventsOnlyMode)
        ? collectCcEmailsFromFields(event.eventSpecificFields, customData, participantEmail)
        : '';
      // v10.15+: Sub-Event-Anmeldungen laufen auch beim Stellvertreter-Modus
      // (registerForOther) durch. registerForEvent akzeptiert ja participantFirstName/
      // -LastName/-Email als Argumente, daher kann der Assistent jede beliebige
      // Person sowohl auf das Hauptevent als auch auf alle gewählten Sub-Events
      // anmelden. Vorher war der Sub-Event-Loop hinter !registerForOther
      // versteckt — Beobachtung des Users: 'beim register for someone else kann
      // man nur fürs Main Event anmelden, nicht für die Sub-Events'. Fix.
      let subOpsDone = 0;
      // v26.67 (B): mind. eine NEUE Sub-Event-Anmeldung erfolgreich? Gate für
      // die nachgelagerte Schatten-Klammer.
      let anySubRegSuccess = false;
      // v29.25: Abwahlen, die wegen der Abmelde-Sperre NICHT abgemeldet wurden.
      const lockedCancelTitles: string[] = [];
      for (const ce of childEvents) {
        const wasReg = sessionMeta[ce.id]?.wasRegistered;
        const isSel = selectedSessions.has(ce.id);
        if (isSel && !wasReg) {
          setSubmitProgressLabel(locale === 'de'
            ? `${childTermSingular || 'Sub-Event'} „${ce.title || '?'}" wird angemeldet…`
            : `Registering for ${childTermSingular || 'sub-event'} „${ce.title || '?'}"…`);
          const sType = isSplitGroup ? (inheritedStarterType || undefined) : undefined;
          // Pro-Sub-Event Custom-Field-Werte aus dem Modal-Flow (sessionFieldValues
          // wird beim Bestätigen des Sub-Event-Modals befüllt). Default: {}.
          // v11.34: Anrede (salutation) zusätzlich mitgeben — vorher fehlte sie
          // bei Sub-Event-Anmeldungen, die Teilnehmerliste hatte dann „-" in
          // der Anrede-Spalte. Salutation kommt aus dem Hauptformular und ist
          // pro User identisch für alle Sub-Event-Anmeldungen.
          // v15.25: Im subEventsOnlyMode landen die Hauptevent-CF-Antworten
          // jetzt in der Schatten-Parent-Registrierung (s.o.) — die Sub-
          // Events bekommen nur ihre eigenen CFs aus dem Modal-Flow.
          const seFieldValues = { salutation, ...(sessionFieldValues[ce.id] || {}) };
          // v18.74: extraCc (übergreifende CC) + proxyConsentConfirmed (Nachweis
          // bei stellvertretender Anmeldung) zusammen in die Opts.
          // v19.6: ccSelfEmail zusätzlich in die CC der Sub-Event-Bestätigung
          // mergen (deduppt serverseitig).
          const seExtraCc = [crossCutCc, ccSelfEmail, delegateCc].filter(Boolean).join(';');
          // v30.9: skipReload — bei Kalender-Events lief nach JEDEM Tag ein
          // kompletter loadEvents (28 MB); jetzt ein Sammel-Refresh am Ende.
          const seOpts = {
            ...(seExtraCc ? { extraCc: seExtraCc } : {}),
            // v30.61: Gebündelte Kommunikation — der Termin wird still
            // angelegt, die Klammer verschickt einmal für alles. Mail und
            // Kalender getrennt, weil beides getrennt schaltbar ist.
            ...(bundledMode.mail ? { suppressMail: true } : {}),
            ...(bundledMode.outlook ? { suppressOutlook: true } : {}),
            ...(registerForOther ? { proxyConsentConfirmed: true, actorAllowedAsAssistant } : {}),
            skipReload: true,
            // v30.42: Diese Seite legt die Klammer-Zeile SELBST an — zuletzt und
            // mit den übergreifenden Antworten (Hotel, Verpflegung, Anreise).
            // Die zentrale Absicherung im EventContext würde hier vorher eine
            // LEERE Schattenzeile schreiben; die Zeile wäre dann belegt und die
            // Antworten würden nie geschrieben. Genau deshalb gibt es das Flag.
            skipShadowParent: true,
          };
          const subRes = await registerForEvent(ce.id, seFieldValues, firstTrim, surnameTrim, participantEmail, sType, seOpts);
          if (subRes.ok) {
            anySuccess = true; anySubRegSuccess = true;
            bookedItems.push({ title: ce.title || '', startDate: ce.startDate, endDate: ce.endDate, location: ce.location });
          }
          else lastSubReason = subRes.reason;
          subOpsDone++;
          setSubmitProgress(50 + Math.floor((subOpsDone / Math.max(subOps, 1)) * 40));
        } else if (!isSel && wasReg && !registerForOther) {
          // v29.25: Selbst-Abmeldung nach der Frist gesperrt (Organizer-
          // Option) — das Abwählen darf hier nicht still abmelden. Der
          // Haken bleibt technisch abgewählt, die Anmeldung besteht weiter;
          // die betroffenen Termine werden nach dem Absenden benannt.
          if (selfCancelLocked(ce, event)) {
            lockedCancelTitles.push(ce.title || (locale === 'de' ? 'Sub-Event' : 'sub-event'));
            subOpsDone++;
            setSubmitProgress(50 + Math.floor((subOpsDone / Math.max(subOps, 1)) * 40));
            continue;
          }
          setSubmitProgressLabel(locale === 'de'
            ? `${childTermSingular || 'Sub-Event'} „${ce.title || '?'}" wird abgemeldet…`
            : `Cancelling ${childTermSingular || 'sub-event'} „${ce.title || '?'}"…`);
          // Cancel-Pfad bleibt aufs Selbst-Anmelden begrenzt: ein Stellvertreter
          // soll nicht aus Versehen einen Sub-Event-Slot des Anderen freigeben
          // weil er den Haken nicht gesetzt hat. Wer einen TN abmelden will,
          // macht das aktiv im Admin Center.
          await cancelRegistration(ce.id, { skipReload: true });
          anySuccess = true;
          subOpsDone++;
          setSubmitProgress(50 + Math.floor((subOpsDone / Math.max(subOps, 1)) * 40));
        }
      }
      // v29.25: Gesperrte Abmeldungen benennen — sonst sähe das Absenden wie
      // eine erfolgreiche Abmeldung aus, obwohl die Anmeldung weiter besteht.
      if (lockedCancelTitles.length > 0) {
        showAlert((event && event.noSelfCancel)
          ? (locale === 'de'
            ? `Nicht abgemeldet: ${lockedCancelTitles.join(', ')}. Die Organizer haben die Selbst-Abmeldung für dieses Event deaktiviert — bitte wende dich zum Abmelden an die Organizer. Deine Anmeldung bleibt bestehen.`
            : `Not cancelled: ${lockedCancelTitles.join(', ')}. The organizers have disabled self-cancellation for this event — please contact the organizers to cancel. Your registration remains in place.`)
          : (locale === 'de'
            ? `Nicht abgemeldet: ${lockedCancelTitles.join(', ')}. Die Abmeldefrist ist abgelaufen und die Organizer haben die Selbst-Abmeldung danach deaktiviert — bitte wende dich zum Abmelden an die Organizer. Deine Anmeldung bleibt bestehen.`
            : `Not cancelled: ${lockedCancelTitles.join(', ')}. The cancellation deadline has passed and the organizers have disabled self-cancellation after it — please contact the organizers to cancel. Your registration remains in place.`),
          { variant: 'error' });
      }
      // v30.67: Wer im „Nur Sub-Events"-Modus den LETZTEN gebuchten Termin
      // abwählt, will sich ganz abmelden — bis hierher hatte der Guard oben
      // das verhindert. Bleibt kein Termin mehr, muss auch die Schatten-
      // Klammer weg, sonst führt „Meine Events" die Person weiter als
      // angemeldet (derselbe Fall wie v15.26 in MyEventSubEvents). Still,
      // weil die Termin-Abmeldungen ihre Mails schon verschickt haben. Bei
      // ausgeblendeten Terminen (hiddenChildCount) ist nicht sicher, ob noch
      // etwas gebucht ist — dann bleibt die Klammer lieber stehen.
      if (isSubOnlyMode && !registerForOther && parentAlreadyHasRow && anySuccess
        && lockedCancelTitles.length === 0 && hiddenChildCount === 0
        && !childEvents.some(ce => selectedSessions.has(ce.id))) {
        setSubmitProgressLabel(locale === 'de' ? 'Hauptevent-Eintrag wird entfernt…' : 'Removing main-event entry…');
        try { await cancelRegistration(selectedEventId!, { suppressNotifications: true, skipReload: true }); }
        catch (err) { console.warn('[DEX] shadow-parent cancel failed:', err); }
      }
      // v26.67 (B) Schritt 3: Schatten-/Klammer-Zeile im subEventsOnly-Modus
      // JETZT anlegen — erst nachdem mind. ein Sub-Event erfolgreich angemeldet
      // wurde. So kann kein unsichtbarer „Geist" entstehen (Klammer ohne
      // Sub-Event). Deckt Selbst- (shouldShadowRegisterParent) UND
      // Fremd-Anmeldung (registerForOther) ab; ohne neue Sub-Event-Anmeldung
      // (z. B. nur Abmeldungen) wird keine leere Klammer erzeugt.
      if (shouldShadowRegisterParent && anySubRegSuccess) {
        setSubmitProgress(92);
        setSubmitProgressLabel(locale === 'de' ? 'Hauptevent-Daten werden gespeichert…' : 'Saving main-event data…');
        await doParentRegistration(true);
        // v29.48: Ein zweiter Versuch, bevor wir aufgeben. Der erste scheitert
        // fast immer an der Drosselung; withThrottleRetry hat dann schon
        // gewartet, das Kontingent ist wieder frei.
        if (shadowParentFailed) {
          shadowParentFailed = false;
          await doParentRegistration(true);
        }
        // v30.57: Steht die Zeile, ist der vorab gesetzte Merker erledigt.
        if (!shadowParentFailed) {
          try { removePendingShadowParent(selectedEventId!, participantEmail); } catch { /* egal */ }
        }
      } else if (shouldShadowRegisterParent) {
        // Kein einziger Sub-Event ging durch — dann gibt es auch nichts
        // nachzuziehen, der Merker muss wieder weg.
        try { removePendingShadowParent(selectedEventId!, participantEmail); } catch { /* egal */ }
      }
      // v30.16: Bleibt die Klammer-Zeile aus, heilt sie sich jetzt SELBST —
      // der v29.48-Fehler-Dialog („SharePoint rejected the last write, bitte
      // erneut speichern") ist weg. Der Fall trat im Soft Opening real auf,
      // obwohl schon zweimal versucht wurde: Die Klammer ist der letzte
      // Schreibvorgang und trifft die Drossel am härtesten. Statt den
      // Teilnehmer zu belasten: (1) Merker mit den Formular-Antworten in
      // localStorage (verlustfrei), (2) stille Wiederholungen im Hintergrund
      // der Erfolgsseite (20 s / 60 s / 180 s), (3) geht der Browser vorher
      // zu, arbeitet der EventContext den Merker beim nächsten App-Start ab.
      // registerForEvent ist für die Klammer idempotent — doppelter Nachzug
      // fügt nichts doppelt ein. Letztes Netz bleibt das Organizer-Panel
      // („Fehlende Klammer-Anmeldung" + Sammel-Fix, v30.14).
      if (shadowParentFailed) {
        const shadowEntry = {
          eventId: selectedEventId!,
          customData: { ...customData },
          firstName: firstTrim,
          lastName: surnameTrim,
          email: participantEmail,
          proxy: registerForOther,
          ts: Date.now(),
        };
        addPendingShadowParent(shadowEntry);
        const retryShadow = (attempt: number): void => {
          window.setTimeout(() => {
            registerForEvent(
              shadowEntry.eventId, shadowEntry.customData, shadowEntry.firstName, shadowEntry.lastName,
              shadowEntry.email, undefined,
              { ...(shadowEntry.proxy ? { proxyConsentConfirmed: true, actorAllowedAsAssistant } : {}), skipReload: true }
            ).then(r => {
              if (r.ok) removePendingShadowParent(shadowEntry.eventId, shadowEntry.email);
              else if (attempt < 3) retryShadow(attempt + 1);
            }).catch(() => { if (attempt < 3) retryShadow(attempt + 1); });
          }, attempt === 1 ? 20000 : attempt === 2 ? 60000 : 180000);
        };
        retryShadow(1);
      }
      // v30.61: Gebündelt UND die Klammer-Zeile stand schon (also lief oben
      // keine Anmeldung, die eine Mail ausgelöst hätte) → eine aktualisierte
      // Sammelmail mit dem VOLLSTÄNDIGEN Stand. `bookedItems` trägt nur die
      // gerade neu gebuchten Termine; für „Deine Anmeldung ist jetzt so" muss
      // alles hinein, was aktuell gebucht ist — sonst läse sich die Mail wie
      // eine Abmeldung von allem anderen.
      if (bundledMode.mail && parentAlreadyHasRow && event) {
        const stillBooked: BundledItem[] = childEvents
          .filter(ce => selectedSessions.has(ce.id))
          .map(ce => ({ title: ce.title || '', startDate: ce.startDate, endDate: ce.endDate, location: ce.location }));
        const changed = bookedItems.length > 0 || lockedCancelTitles.length === 0;
        if (changed && stillBooked.length > 0) {
          const fullName = `${firstTrim} ${surnameTrim}`.trim() || participantEmail;
          void sendBundledUpdateMail(event, participantEmail, fullName, stillBooked);
        }
      }
      setSubmitProgress(95);
      setSubmitProgressLabel(locale === 'de' ? 'Bestätigungen werden versandt…' : 'Confirmations are being queued…');

      if (anySuccess) {
        // v19.0: ausgewählte Dokument-Dateien als Attachment an die Teilnehmer-
        // Zeile hängen — das Item existiert jetzt. Bei stellvertretender
        // Anmeldung an die Teilnehmer-E-Mail, sonst an den eingeloggten User.
        const docFields = (event?.eventSpecificFields || []).filter(f => f.type === 'document');
        const anyDoc = docFields.some(df => !!pendingDocFiles[df.id]);
        if (anyDoc) {
          setSubmitProgressLabel(locale === 'de' ? 'Dokumente werden hochgeladen…' : 'Uploading documents…');
          for (const df of docFields) {
            const file = pendingDocFiles[df.id];
            if (!file) continue;
            try { await uploadFieldDocument(selectedEventId!, df.id, file, registerForOther ? participantEmail : undefined); }
            catch { /* best-effort — Anmeldung bleibt gültig, Upload kann später über „Meine Events" nachgeholt werden */ }
          }
        }
        // Flag: wenn ausschließlich Sessions angemeldet/geändert wurden (kein
        // Parent diesmal oder schon vorher angemeldet), zeigen wir auf der
        // Success-Seite den Sessions-Only-Hinweis.
        setSessionsOnlySubmitted(!willRegisterParent && !registerForOther);
        // v24.41 Szenario A: Assistenz verknüpfen (Info + Anforderung). Der
        // Owner bleibt der/die Anmeldende; die Assistenz sieht es als Info.
        if (delegateAssist) {
          try { await delegateRegistrationToAssistant(selectedEventId!, delegateAssist); }
          catch { /* best-effort — Anmeldung bleibt gültig */ }
        }
        // v24.46: Hat das Event ein Assistenz-CC-Feld (Organizer hat es selbst
        // eingebaut → KEIN Modal) und der User dort eine Person angegeben, läuft
        // dieselbe Info-Freischaltung automatisch über dieses Feld — für JEDEN
        // Anmelder, nicht nur Partner/Director. Greift nur bei Selbst-Anmeldung
        // (für andere: die andere Person ist die angemeldete, nicht der CC).
        if (!registerForOther) {
          const ccFields = (event?.eventSpecificFields || []).filter(f => (f.type === 'user' || f.type === 'roommate') && !!f.ccOnEmails);
          const seenAssist = new Set<string>();
          for (const f of ccFields) {
            const raw = (customData[f.id] || '').trim();
            for (const part of raw.split(';').map(s => s.trim()).filter(Boolean)) {
              const m = part.match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
              const aEmail = m ? m[2].trim() : '';
              const aName = m ? m[1].trim() : '';
              if (!aEmail) continue;
              const key = aEmail.toLowerCase();
              if (key === participantEmail.toLowerCase() || seenAssist.has(key)) continue;
              seenAssist.add(key);
              try { await delegateRegistrationToAssistant(selectedEventId!, { email: aEmail, name: aName }); }
              catch { /* best-effort */ }
            }
          }
        }
        // v24.41 Szenario B: Bei stellvertretender Anmeldung (für andere) einen
        // Info-Link anlegen — der/die Anmeldende ist Owner, die angemeldete
        // Person sieht die Anmeldung als Info unter „Meine Events".
        if (registerForOther && !externalPerson && participantEmail
          && participantEmail.toLowerCase() !== (currentUser.email || '').toLowerCase()) {
          try { await recordProxyDelegation(selectedEventId!, { email: participantEmail, name: `${firstTrim} ${surnameTrim}`.trim() || participantEmail }); }
          catch { /* best-effort */ }
        }
        // v30.9: EIN Refresh für alle Schreibvorgänge dieses Absendens (die
        // Einzel-Aufrufe oben laufen mit skipReload). Bewusst NICHT awaiten —
        // die Erfolgsseite soll sofort erscheinen; Kacheln/Zähler ziehen im
        // Hintergrund nach.
        void refreshEvents().catch(() => { /* best-effort */ });
        setSubmitted(true);
      } else if (!parentOk) {
        // Parent-Fehler wurde schon in setError oben gesetzt.
      } else {
        setError(regFailMessage(lastSubReason));
      }
    } catch {
      setError(t('reg.genericerror'));
    } finally {
      setSubmitProgress(100);
      setSubmitProgressLabel(locale === 'de' ? 'Fertig!' : 'Done!');
      // v19.6: CC-Frage-Entscheidung zurücksetzen, damit der nächste
      // Submit-Durchlauf (z.B. nächste stellvertretende Anmeldung) wieder fragt.
      ccSelfDecidedRef.current = false;
      ccSelfRef.current = false;
      // v24.48: Assistenz-Entscheidung zurücksetzen (nächster Submit fragt neu).
      assistantModalDecidedRef.current = false;
      delegateChoiceRef.current = null;
      // Kleine Verzögerung damit der User die 100%-Anzeige kurz sieht
      // bevor das Overlay wieder verschwindet.
      setTimeout(() => {
        setIsSubmitting(false);
        setSubmitProgress(0);
        setSubmitProgressLabel('');
      }, 250);
    }
  };

  return { handleSubmit, performRegistration };
}
