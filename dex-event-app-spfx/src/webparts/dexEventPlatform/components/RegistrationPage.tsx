/**
 * Registrierungsseite für ein einzelnes Event
 *
 * Drei-Spalten-Layout: Event-Info | persönliche Daten | eventspezifische Felder
 * Speichert die Registrierung in der SharePoint-Teilnehmerliste des Events.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
// v22.10: Sub-Sections nach ihrer EIGENEN Sichtbarkeit filtern (gleiche Logik
// wie die Event-Liste) — sonst sieht jeder Hauptevent-Teilnehmer alle Sub-Events.
import { isEventVisibleForUser } from './EventListPage';
import { useCachedImage, useCachedImageWithFallback } from '../utils/imageCache';
import { useIsMobile } from '../utils/useIsMobile';
import { isRegistrationFullyClosed } from '../utils/eventFormat';
import { useLanguage, translations as appTranslations, Locale } from '../context/LanguageContext';
// v20.4: modernes Alert-Modal statt window.alert.
import { useDialog } from '../context/DialogContext';
import { Salutation, EventSpecificField, DeloitteEvent } from '../types';
import { Icon } from '@fluentui/react/lib/Icon';
import { InfoTooltip } from './InfoTooltip';
import { MultiSelectDropdown } from './MultiSelectDropdown';
import OrganizerList from './OrganizerList';
import { UserFieldPicker } from './UserFieldPicker';
// v29.51: nachgeladen — zieht react-datepicker + date-fns aus dem Boot-Bundle.
import StayRangePicker from './StayRangePickerLazy';
import { dlog } from '../utils/debugLog';

// v30.66: Modul-Ebene (Formatierer, Sanitizer, CollapsibleSection, Bild-Cache)
// liegt in einer eigenen Datei — die ausgelagerten Teilbaeume der Seite brauchen
// dieselben Helfer und duerfen sie nicht aus der Seite zurueckimportieren.
import { formatDate, isExternalEmailAddr, renderFieldDescHtml, IMG_ASPECT_CACHE } from './registration/regHelpers';
import { SubEventFieldsModal } from './registration/SubEventFieldsModal';
import { AssistantModal, CcSelfModal, ExternalEmailWarningModal, FallbackDialogModal } from './registration/SmallModals';
import { SubmitConfirmModal } from './registration/SubmitConfirmModal';
import { MassImportModal } from './registration/MassImportModal';
import { ProxyWizardModal } from './registration/ProxyWizardModal';
import { PrivacyNote, RegistrationActionBar } from './registration/RegistrationActionBar';
import { EventSpecificSection } from './registration/EventSpecificSection';
import { OpenTeamsList, TeamRegistrationCard } from './registration/TeamSection';
import { PersonalDataSection } from './registration/PersonalDataSection';
import { EventCard } from './registration/EventCard';
import { DeadlineBanner, DemoBanner, LocationBanner, SubmitOverlay } from './registration/RegistrationBanners';
import { createSubmitFlow } from './registration/submitFlow';

export default function RegistrationPage(): React.ReactElement {
  // v11.98: Beim Mount nach oben scrollen. Sonst behält der scrollende
  // .main-content-Container die Position aus der vorherigen Seite (z.B.
  // wenn man weit unten in der Events-Kachel war und dann auf Register
  // klickt — die Register-Page erscheint dann mittendrin statt am Anfang).
  React.useEffect(() => {
    const main = document.querySelector('.main-content');
    if (main && typeof (main as HTMLElement).scrollTo === 'function') {
      (main as HTMLElement).scrollTo({ top: 0, behavior: 'auto' });
    } else if (main) {
      (main as HTMLElement).scrollTop = 0;
    }
    if (typeof window !== 'undefined') {
      window.scrollTo({ top: 0, behavior: 'auto' });
    }
  }, []);

  const { selectedEventId, navigate, navIntent, clearIntent } = useNavigation();
  const { events, isEventsLoading, registerForEvent, registerTeam, cancelRegistration, declineEvent, checkRegistrationByEmail, getMyRegistration, getAllRegistrations, childEventsOf, listOpenTeamsForEvent, joinTeam, createTeamJoinRequest, updateMyRegistration, uploadFieldDocument, delegateRegistrationToAssistant, recordProxyDelegation, getLiveCounterStats, subscribeEventRealtime, getEventNumbersForEmail, refreshEvents, sendBundledUpdateMail } = useEvents();
  const { currentUser, groupEmails } = useCurrentUser();
  // v30.3: previewAsUser = „Übersicht als User sehen" (Organizer-/Admin-
  // Vorschau). isAdmin kommt aus dem RoleContext bereits absenkt zurück;
  // nur der per-Event-Organizer-Check unten muss lokal mitziehen.
  const { searchUsers, searchUser, isAdmin, previewAsUser } = useRoles();
  const { locale: appLocale } = useLanguage();
  // v20.4: App-Modal statt nativem Browser-Alert.
  const { showAlert, confirmDialog } = useDialog();
  // Mobile-Breakpoint für kompaktere Handy-Darstellung (einklappbare Sektionen etc.).
  const isMobile = useIsMobile();
  const event = events.find(e => e.id === selectedEventId);

  // v18.35: Erzwungene Anmeldesprache. Hat der Organizer für dieses Event eine
  // feste Anmeldesprache gesetzt ('de'/'en'), wird die GESAMTE Anmeldeseite
  // (App-Chrome, Form-Chrome, Inline-Texte UND Disclaimer) in dieser Sprache
  // angezeigt — unabhängig von der App-Sprache des Teilnehmers. Wir
  // überschreiben dazu `locale` und `t` lokal; alle bestehenden Verwendungen
  // im Rest der Datei greifen damit automatisch auf die erzwungene Sprache zu.
  const forcedRegLang: Locale | undefined =
    (event?.registrationLanguage === 'de' || event?.registrationLanguage === 'en') ? event.registrationLanguage : undefined;
  const locale: Locale = forcedRegLang || appLocale;
  const t = React.useCallback(
    (key: string): string => appTranslations[locale][key] || appTranslations['en'][key] || appTranslations['de'][key] || key,
    [locale]
  );

  // v11.56/v17.20/v19.18: tEvent() liefert die Form-Chrome-Strings (Placeholder,
  // Hints, Sub-Event-/Auswahl-Sektion).
  // v19.18 FIX: Form-Chrome folgt jetzt IMMER der App-/erzwungenen Anmeldesprache
  // (`locale`) — NICHT mehr der Event-MAIL-Sprache (`emailLanguage`). Vorher
  // entstand eine verwirrende Misch-Anzeige: der Großteil der Anmeldeseite in der
  // App-Sprache (z.B. DE), aber die Sub-Event-Auswahl + Platzhalter in der
  // Event-Mail-Sprache (z.B. EN bei einem B2Run-Event mit emailLanguage='EN').
  // Die Anmeldeseite zeigt der Person jetzt durchgängig EINE Sprache: die
  // erzwungene Anmeldesprache (falls per Event gesetzt), sonst die App-Sprache
  // des Teilnehmers. Die Mail-Sprache (`emailLanguage`) steuert weiterhin NUR die
  // tatsächlichen E-Mails — nicht die Formular-Anzeige. Der Bilingual-Toggle
  // steuert davon unberührt weiter die EN-Varianten der Custom-Field-Labels
  // (siehe `useEnVariants` unten).
  const eventLocale: Locale = locale;
  // v29.13: Die Anmeldeseite zeigt das Event-Bild aus Schritt 1. Es ist NICHT
  // dasselbe Bild wie das Mail-Logo aus dem Kommunikations-Schritt — Mails und
  // Outlook-Termin nehmen das, die Seite hier nicht. Wer nur eines von beiden
  // pflegt, pflegt meist das Mail-Logo (man sieht es sofort im Postfach) und
  // wundert sich, warum die Anmeldeseite den generischen DEX-Kreis zeigt.
  // Deshalb: kein Event-Bild → das Mail-Logo des Events übernehmen. Es bleibt
  // ein Rückfall; ist ein Event-Bild da, hat es immer Vorrang.
  const heroImgUrl = (event?.imageUrl || '') || (event?.mailImageBase64 || '');
  const usesMailImage = !event?.imageUrl && !!event?.mailImageBase64;
  // v19.22: Event-Bild über den IndexedDB-Cache (sofort beim zweiten Aufruf).
  // Base64 gehört nicht in den Cache — es liegt bereits vollständig vor.
  const cachedImage0 = useCachedImage(event?.imageUrl);
  const cachedImage = usesMailImage ? heroImgUrl : cachedImage0;
  // v28.11: Vergrößerte Hover-Ansicht des Event-Bilds — zeigt bevorzugt das
  // unbeschnittene Querformat-Original (falls vorhanden), sonst das Event-Bild.
  // v29.34: mit Rückfall auf das Event-Bild — die Original-URL kann ins Leere
  // zeigen (siehe useCachedImageWithFallback), die Lupe blieb dann leer.
  const cachedZoomImage0 = useCachedImageWithFallback(event?.imageOrigUrl, event?.imageUrl);
  const cachedZoomImage = usesMailImage ? heroImgUrl : cachedZoomImage0;
  // v28.12: Kein Auto-Zoom mehr beim Hover — der Hover zeigt nur ein
  // Lupen-Icon, erst der KLICK darauf öffnet die Großansicht (Lightbox).
  const [imgHovered, setImgHovered] = React.useState(false);
  const [imgZoomed, setImgZoomed] = React.useState(false);
  const tEvent = React.useCallback((key: string): string => {
    return appTranslations[eventLocale][key] || appTranslations['en'][key] || appTranslations['de'][key] || t(key) || key;
  }, [eventLocale, t]);
  // v17.20: Lookup-Helfer für die EN-Varianten eines Custom-Fields. Greift
  // nur, wenn der Bilingual-Toggle des Events an ist UND die App-Locale des
  // Teilnehmers `en` ist. Sonst still Fallback auf den DE-Wert. Index-Mapping
  // der Optionen ist positional — DE-Option i ↔ EN-Option i.
  // v17.22: `useEnVariants` steuert NUR die Anzeige-Labels. Die gespeicherten
  // Werte bleiben in JEDEM Fall die kanonischen DE-Originale: Single-Select
  // rendert `<option value={DE-Wert}>{EN-Anzeige}</option>`, Multi-Select gibt
  // `options={field.options}` (DE) als Wert weiter und nutzt `optionLabels`
  // nur für die Darstellung. Deshalb ist auch der „Register-for-Other"-Pfad
  // unkritisch: meldet ein EN-Organizer eine DE-Person an, sieht der Organizer
  // die EN-Labels (er füllt das Formular), gespeichert wird aber der
  // DE-Wert — die Zielperson und die Bestätigungs-Mail (event.emailLanguage)
  // bekommen also keine sprachlich falschen Daten.
  const useEnVariants = !!event?.bilingualFields && locale === 'en';
  const pickFieldLabel = React.useCallback((f: EventSpecificField): string =>
    (useEnVariants && f.labelEn && f.labelEn.trim()) ? f.labelEn : f.label,
  [useEnVariants]);
  const pickFieldHelp = React.useCallback((f: EventSpecificField): string | undefined =>
    (useEnVariants && f.helpTextEn && f.helpTextEn.trim()) ? f.helpTextEn : f.helpText,
  [useEnVariants]);
  const pickFieldConfirmLabel = React.useCallback((f: EventSpecificField): string | undefined =>
    (useEnVariants && f.confirmLabelEn && f.confirmLabelEn.trim()) ? f.confirmLabelEn : f.confirmLabel,
  [useEnVariants]);
  const pickOptionLabel = React.useCallback((f: EventSpecificField, optIdx: number, fallback: string): string => {
    if (useEnVariants && f.optionsEn && f.optionsEn[optIdx] && f.optionsEn[optIdx].trim()) {
      return f.optionsEn[optIdx];
    }
    return fallback;
  }, [useEnVariants]);

  // Per-Event-Organizer-Check: ist der eingeloggte User Haupt- ODER Co-Organizer
  // dieses Events? Nur dann darf er a) nach Deadline registrieren und b)
  // "Register for another person" nutzen. Ein Organizer von EVENT A darf NICHT
  // für EVENT B solche Admin-Aktionen ausführen. Admin darf global alles.
  // v19.6: Co-Organizer (event.coOrganizerEmails) zählen hier ausdrücklich
  // mit — vorher sah ein Co-Organizer den „Für andere registrieren"-Button
  // nicht, obwohl er das Event mitorganisiert. Serverseitig wird derselbe
  // Personenkreis in canRegisterForOthers() akzeptiert.
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  // v30.3: In der User-Vorschau zählt auch die per-Event-Organizer-
  // Eigenschaft nicht — sonst blieben Frist-/Freischalt-Bypässe aktiv
  // und die Seite sähe eben NICHT aus wie für reguläre User.
  const isEventOrganizerReal = !!event && (
    event.organizerEmails.some(e => (e || '').toLowerCase() === currentEmailLc) ||
    (event.coOrganizerEmails || []).some(e => (e || '').toLowerCase() === currentEmailLc)
  );
  const isEventOrganizer = !previewAsUser && isEventOrganizerReal;
  const isOrganizer = isEventOrganizer; // alten Namen behalten für Referenzen unten
  const canCreateEvents = isEventOrganizer || isAdmin; // statt tenant-weitem Organizer

  // Assistant-Ausnahme: User mit JobTitle "Assistant" / "Senior Assistant" dürfen
  // "Register for another person" nutzen, allerdings NUR für Director/Partner und
  // NUR für Events für die sie sich eh selber anmelden könnten (also nicht nach
  // Deadline). Der Deadline-Schutz greift automatisch, weil RegistrationPage für
  // normale User nach Deadline komplett die "closed"-Seite zeigt und gar nicht
  // zum Button-Rendering kommt.
  const currentJobTitleLc = (currentUser.jobTitle || '').toLowerCase();
  const isAssistant = currentJobTitleLc.includes('assistant');
  const ALLOWED_TARGET_TITLES = ['partner', 'director'];
  const isAllowedTargetForAssistant = (jt: string): boolean => {
    const lc = (jt || '').toLowerCase();
    return ALLOWED_TARGET_TITLES.some(t => lc === t || lc.indexOf(t) >= 0);
  };
  const canRegisterForOther = canCreateEvents || isAssistant;

  // Sichtbarkeits-Check: Würde dieses Event dem User als normaler User angezeigt werden?
  const showLocationBanner = canCreateEvents && event && (() => {
    const locFilters = event.locationAudience;
    // Audience-Filter normalisieren: 'All'/'DEALL' = "kein Audience-Filter"
    const audFilters = (event.audienceFilter || [])
      .map(s => s.trim())
      .filter(s => s && s.toLowerCase() !== 'all' && s.toLowerCase() !== 'deall');
    const hasLoc = locFilters.length > 0;
    const hasAud = audFilters.length > 0;
    if (!hasLoc && !hasAud) return false; // kein Filter = alle sehen es

    const loc = (currentUser.location || '').toLowerCase();
    const email = currentUser.email.toLowerCase();

    const locMatch = !hasLoc || locFilters.some(f => {
      const fl = f.trim().toLowerCase();
      if (fl === 'all') return true;
      const norm = fl.replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ä/g, 'ae');
      return loc.indexOf(fl) >= 0 || loc.indexOf(norm) >= 0;
    });

    const audMatch = !hasAud || audFilters.some(f => {
      const fl = f.trim().toLowerCase();
      if (fl.indexOf('@') >= 0) return email === fl;
      if (fl.startsWith('de')) {
        const city = fl.substring(2);
        const norm = city.replace(/oe/g, 'ö').replace(/ue/g, 'ü').replace(/ae/g, 'ä');
        return loc.indexOf(city) >= 0 || loc.indexOf(norm) >= 0;
      }
      return false;
    });

    // Default: AND (Schnittmenge). Nur OR wenn explizit gesetzt.
    // Wichtig: wenn nur EIN Filter gesetzt ist, zählt nur dieser - egal ob AND/OR.
    const mode = event.filterMode || 'AND';
    let visible: boolean;
    if (mode === 'OR') {
      if (hasLoc && hasAud) visible = locMatch || audMatch;
      else if (hasLoc) visible = locMatch;
      else visible = audMatch;
    } else {
      // AND
      if (hasLoc && hasAud) visible = locMatch && audMatch;
      else if (hasLoc) visible = locMatch;
      else visible = audMatch;
    }
    return !visible;
  })();

  const [salutation, setSalutation] = React.useState<Salutation | ''>('');
  const [firstName, setFirstName] = React.useState(currentUser.firstName);
  const [surname, setSurname] = React.useState(currentUser.surname);
  const [email, setEmail] = React.useState(currentUser.email);
  const [registerForOther, setRegisterForOther] = React.useState(false);
  // v24.41: „Meine Assistenz beauftragen" — Admins/Directoren können bei der
  // eigenen Anmeldung eine Assistenz angeben, die danach die Anmeldung in ihrer
  // „Assistenz"-Kachel verwaltet (Delegation) und auf CC der Bestätigung kommt.
  const [delegateAssistEnabled, setDelegateAssistEnabled] = React.useState(false);
  const [delegateAssistValue, setDelegateAssistValue] = React.useState('');
  // v18.74: „Person außerhalb Deloitte" — explizit eine externe Person
  // stellvertretend anmelden. Blendet den Deloitte-People-Picker aus und macht
  // Vorname/Nachname/E-Mail frei eintragbar. Die Zustimmung ist hier SCHRIFTLICH
  // einzuholen; es wird kein Outlook-Termin versendet (Bestätigungs-Mail mit
  // Organizer auf CC).
  const [externalPerson, setExternalPerson] = React.useState(false);

  // Wenn die Seite mit Intent 'register-other' geöffnet wird (z.B. via "Register another person"
  // Button auf einer Karte, für die der Organizer/Admin schon selbst registriert ist),
  // direkt in den "Für andere registrieren"-Modus springen und Felder leeren.
  React.useEffect(() => {
    if (navIntent === 'register-other' && (canCreateEvents)) {
      setRegisterForOther(true);
      setFirstName(''); setSurname(''); setEmail('');
      clearIntent();
    }
  }, [navIntent, canCreateEvents]);
  const [eventSpecific, setEventSpecific] = React.useState<Record<string, string>>({});
  // v19.0: Pro Dokument-Custom-Feld die ausgewählte Datei (vor dem Absenden).
  // Wird NICHT in customData geschrieben — nach erfolgreicher Anmeldung als
  // Attachment an die Teilnehmer-Zeile gehängt.
  const [pendingDocFiles, setPendingDocFiles] = React.useState<Record<string, File | null>>({});
  const [preferredStarterType, setPreferredStarterType] = React.useState<string>('');
  // Seit v6.5: Fallback-Dialog wenn B2Run-Wunschtyp voll, aber Alternative frei.
  const [fallbackDialog, setFallbackDialog] = React.useState<{ wunsch: string; alt: string; altFree: number } | null>(null);
  // v19.19: zusätzlich zu den aktiven Belegungen pro Gruppe (durch/fun) auch
  // die Wartelisten-Zahlen pro Gruppe (durchWait/funWait) — für die
  // Kapazitäts-/Warteliste-Anzeige in der Gruppen-Auswahl.
  const [starterCounts, setStarterCounts] = React.useState<{ durch: number; fun: number; durchWait: number; funWait: number } | null>(null);
  const [submitted, setSubmitted] = React.useState(false);
  // v18.11: „Ich nehme nicht teil"-Absage.
  const [declined, setDeclined] = React.useState(false);
  const [isDeclining, setIsDeclining] = React.useState(false);
  // v18.13/v18.14: Massenimport von Drittpersonen (nur Organizer/Admin im
  // „Für andere registrieren"-Modus). Zwei Schritte: (1) Liste einfügen →
  // gegen das Verzeichnis auflösen, (2) Vorschau-Tabelle prüfen → anmelden.
  const [massImportOpen, setMassImportOpen] = React.useState(false);
  const [massImportText, setMassImportText] = React.useState('');
  const [massImportMode, setMassImportMode] = React.useState<'mail' | 'nomail' | 'silent'>('mail');
  const [massImportStep, setMassImportStep] = React.useState<'input' | 'preview'>('input');
  const [massImportResolving, setMassImportResolving] = React.useState(false);
  const [massImportRows, setMassImportRows] = React.useState<Array<{
    email: string; firstName: string; lastName: string; jobTitle: string; location: string;
    status: 'ok' | 'duplicate' | 'notfound'; raw: string;
  }>>([]);
  const [massImportBusy, setMassImportBusy] = React.useState(false);
  const [massImportProgress, setMassImportProgress] = React.useState('');
  const [massImportResult, setMassImportResult] = React.useState<{ ok: number; failed: string[] } | null>(null);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  // v30.19: Während des Absendens warnt der Browser vor dem Schließen des
  // Fensters/Tabs — ein abgebrochener Submit hinterlässt halbe Anmeldungen
  // (Tage ohne Klammer, fehlende Bestätigungen). Der Text im Dialog kommt
  // vom Browser; entscheidend ist preventDefault + returnValue.
  // WICHTIG: Hook steht VOR den frühen Returns (v30.4-Regel dieser Datei).
  React.useEffect(() => {
    if (!isSubmitting) return undefined;
    const warnBeforeUnload = (e: BeforeUnloadEvent): void => { e.preventDefault(); e.returnValue = ''; };
    window.addEventListener('beforeunload', warnBeforeUnload);
    return () => window.removeEventListener('beforeunload', warnBeforeUnload);
  }, [isSubmitting]);
  // v11.33: Submit-Overlay mit Fortschrittsanzeige (Prozent + Label).
  // Bei vielen Sub-Events / vielen Custom-Fields kann der Submit
  // mehrere Sekunden dauern — vorher hat der User nur einen disabled
  // Button gesehen ohne Feedback was gerade passiert.
  const [submitProgress, setSubmitProgress] = React.useState(0);
  const [submitProgressLabel, setSubmitProgressLabel] = React.useState('');
  // v29.29: Angezeigter Fortschritt. `submitProgress` ist der ZIELWERT, den die
  // Anmelde-Schritte setzen — er springt naturgemäß (5 → 30 → 50 → 100), weil
  // dazwischen einzelne, unterschiedlich lange SharePoint-Aufrufe liegen. Der
  // Balken lief dadurch in drei Sätzen statt gleichmäßig. Der Anzeigewert
  // nähert sich dem Ziel jetzt in kleinen Schritten und kriecht, solange ein
  // Schritt dauert, langsam weiter (gedeckelt vor der nächsten Stufe) — so
  // steht der Balken nie still und überholt den echten Stand auch nicht.
  const [displayProgress, setDisplayProgress] = React.useState(0);
  // Start und Ende eines Laufs setzen die Anzeige zurück. Das gehört in einen
  // EIGENEN Effekt: Der Animations-Effekt unten läuft bei jeder Ziel-Änderung
  // neu an und würde die Anzeige sonst mitten im Lauf auf 0 zurückwerfen.
  React.useEffect(() => { setDisplayProgress(0); }, [isSubmitting]);
  React.useEffect(() => {
    if (!isSubmitting) return undefined;
    let tick = 0;
    const id = window.setInterval(() => {
      tick++;
      setDisplayProgress(prev => {
        const target = Math.min(100, Math.max(0, submitProgress));
        // Abschluss: zügig auf 100 aufziehen.
        if (target >= 100) return Math.min(100, prev + Math.max(1, Math.ceil((100 - prev) / 4)));
        // Annäherung ans Ziel — je größer der Rückstand, desto größer der Schritt.
        if (prev < target) return Math.min(target, prev + Math.max(1, Math.ceil((target - prev) / 6)));
        // Ziel erreicht, der Schritt läuft noch: langsam weiterkriechen.
        //
        // Die Stufen liegen weit auseinander (30 → 50 → 95): Die Anmeldung des
        // Haupt-Events ist EIN langer Aufruf ohne Zwischenstand. Ein Deckel von
        // +8 (erste Fassung) ließ den Balken deshalb bei 38 stehen, bis am Ende
        // alles auf einmal kam — „38 % und dann direkt auf 100 %". Der Kriech-
        // Bereich deckt jetzt bis zu 25 Punkte ab und wird dabei immer
        // langsamer (erst alle ~0,25 s ein Prozent, später über eine Sekunde),
        // damit er die Lücke füllt, ohne dem echten Stand davonzulaufen.
        //
        // WICHTIG: Die Anzeige darf hier NICHT wieder auf den Zielwert
        // zurückgezogen werden. Genau das tat eine frühere Fassung („liegt die
        // Anzeige über dem Ziel, übernimm das Ziel") — der Kriech-Schritt ging
        // auf 31, die nächste Runde sprang zurück auf 30, und der Balken
        // flackerte 30/31/30/31. Ein Lauf startet ohnehin bei 0 (Effekt oben).
        const creepCap = Math.min(95, target + 25);
        if (prev >= creepCap) return prev;
        const every = 4 + (prev - target); // je weiter vorgekrochen, desto träger
        return (tick % Math.max(4, every) === 0) ? prev + 1 : prev;
      });
    }, 60);
    return () => window.clearInterval(id);
  }, [isSubmitting, submitProgress]);
  const [error, setError] = React.useState('');
  const [showErrors, setShowErrors] = React.useState(false);
  // v11.91: showDescription wurde entfernt — Beschreibung ist immer offen.
  const [thirdPartyCheck, setThirdPartyCheck] = React.useState<{ alreadyRegistered: boolean; notInAudience: boolean; registeredName?: string; registeredDate?: string } | null>(null);
  // v27.13: Profil-Karte („Persönliche Informationen") — Toggle für die
  // vollständige Liste der automatisch übernommenen Profildaten.
  // v28.1: standardmäßig AUSGEKLAPPT (Wunsch E.B.) — volle Transparenz ohne
  // Klick; über den Minus-Button weiterhin einklappbar.
  const [profileCardExpanded, setProfileCardExpanded] = React.useState(true);

  // Seit v6.14: integrierte Session-Auswahl direkt auf der Registrierungsseite.
  // Der User kann auf EINER Seite wählen, ob er sich für das Haupt-Event und/oder
  // einzelne Sub-Events anmelden möchte. Bei B2Run-Parents zusätzlich pro Session
  // eine Durchstarter/Funstarter-Auswahl.
  // v22.10 (Bugfix): Sub-Sections werden jetzt nach ihrer EIGENEN Sichtbarkeit
  // gefiltert. Vorher sah jeder, der das Hauptevent sehen konnte, ALLE
  // Sub-Events — auch wenn das Sub-Event einen eigenen Empfängerkreis hatte.
  // Organizer/Admins (und der „Für andere registrieren"-Modus) sehen weiterhin
  // alle Sub-Sections, damit sie stellvertretend buchen können.
  const childEvents = React.useMemo(() => {
    if (!event) return [];
    // v28.2: Sub-Events SOFT-deaktiviert (_subEventsDisabled) — für ALLE
    // ausblenden (auch Organizer/Stellvertreter: es soll niemand mehr auf
    // deaktivierte Sub-Events gebucht werden). Bestehende Anmeldungen
    // bleiben unberührt (MyEvents/Admin lesen die Kinder direkt).
    if (event.subEventsDisabled) return [];
    const all = childEventsOf(event.id);
    if (canCreateEvents || registerForOther) return all;
    // v22.68: Sub-Events im Entwurf (isFictive) sind für reguläre Teilnehmer
    // NICHT buchbar — vorher wurden sie nicht gefiltert und waren trotz
    // „Entwurf" buchbar, solange die Klammer sichtbar war. Organizer/
    // Stellvertreter (oben) sehen Entwürfe weiterhin.
    return all.filter(ce => !ce.isFictive && isEventVisibleForUser(ce, currentUser.email, currentUser.location, groupEmails, currentUser.jobTitle));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [event?.id, canCreateEvents, registerForOther, currentUser.email, currentUser.location, groupEmails]);
  /**
   * v29.77: „Anmeldung ab" gilt jetzt fuer ALLE Sub-Event-Darstellungen,
   * nicht nur die Kalender-Kacheln — deshalb ein gemeinsamer Rechner.
   * mode 'fixed' = ein Zeitpunkt fuer alle; 'day'/'week' = rollierend
   * relativ zum Start des jeweiligen Sub-Events (bei 'week' zum Montag
   * seiner Woche, Mitternacht lokal — identisch zur Kachel-Logik).
   */
  const subOpenFrom = (startIso?: string): Date | null => {
    const rule = event?.subEventOpenRule;
    if (!rule) return null;
    if (rule.mode === 'fixed') {
      const d = new Date(rule.date || '');
      return isFinite(d.getTime()) ? d : null;
    }
    if (!((rule.days || 0) > 0)) return null;
    const base = new Date(startIso || '');
    if (!isFinite(base.getTime())) return null;
    const d = new Date(base.getFullYear(), base.getMonth(), base.getDate());
    if (rule.mode === 'week') d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
    d.setDate(d.getDate() - (rule.days || 0));
    d.setHours(0, 0, 0, 0);
    return d;
  };
  /**
   * v29.9: Wie viele buchbare Sub-Events hat das Event, die dieser Person der
   * Zielgruppen-Filter oben WEGGENOMMEN hat?
   *
   * Das ist der Unterschied zwischen „es gibt keine" und „für dich ist keines
   * freigegeben" — und der ist wichtig: Ist die Klammer weiter gefasst als ihre
   * Sub-Events, kann jemand die Anmeldeseite öffnen und findet nichts zum
   * Anklicken. Die Meldung dazu behauptete bis v29.8, es sei „aktuell keines
   * angelegt". Das ist für die betroffene Person nachweislich falsch und
   * schickt sie mit der falschen Frage zu den Organizern.
   */
  const hiddenChildCount = React.useMemo(() => {
    if (!event || event.subEventsDisabled) return 0;
    if (canCreateEvents || registerForOther) return 0;
    const bookable = childEventsOf(event.id).filter(ce => !ce.isFictive);
    return Math.max(0, bookable.length - childEvents.length);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [event?.id, event?.subEventsDisabled, canCreateEvents, registerForOther, childEvents.length]);
  // v15.10: vom Organizer konfigurierbare Bezeichnung (z.B. „Event-Sections",
  // „Workshops"). Wenn gesetzt überschreibt das die Default-Übersetzung
  // („Sessions" / „Sub-Events") überall im RegistrationPage-UI.
  // v29.13: Besteht das Event ausschließlich aus Sub-Events, ist „Sub-Event"
  // die falsche Bezeichnung — es gibt kein Haupt-Event, unter dem sie hingen.
  // Für den Teilnehmer sind das schlicht DIE Events. Ohne eigenen Begriff des
  // Organizers wird deshalb hier der Default umgestellt; ein gesetzter eigener
  // Begriff hat weiterhin Vorrang. Weil damit alle Texte über diese beiden
  // Konstanten laufen, verschwindet „Sub-Event" in diesem Fall überall auf der
  // Anmeldeseite auf einmal — statt an einem Dutzend Einzelstellen.
  const subOnlyTerms = !!(event && event.subEventsOnlyMode);
  const childTermSingular = (event && event.childEventTermSingular)
    || (subOnlyTerms ? (locale === 'de' ? 'Event' : 'event') : '');
  const childTermPlural = (event && event.childEventTermPlural)
    || (subOnlyTerms ? (locale === 'de' ? 'Events' : 'events') : '');
  /**
   * v29.13: „ein Event" — aber „eine Session". Der unbestimmte Artikel hing
   * bisher fest an „eine …", was schon mit dem Default „Sub-Event" falsch war
   * („eine Sub-Event") und mit dem neuen Default „Event" auffällt. Wir raten
   * das Geschlecht nicht, sondern führen die wenigen femininen Begriffe, die
   * als Bezeichnung realistisch vorkommen; alles andere ist maskulin/neutral.
   */
  const childOneDe = React.useMemo(() => {
    const term = childTermSingular || 'Sub-Event';
    // v29.60: Gepflegtes Geschlecht schlaegt die Heuristik. Und: Alle Saetze,
    // in denen childOneDe vorkommt, brauchen den AKKUSATIV
    // („Bitte wähle mindestens … aus") — maskulin heisst das
    // „einen", nicht „ein". Genau daran ist
    // „mindestens ein Office-Tag" aufgefallen.
    const g = event && event.childEventTermGender;
    if (g === 'm') return `einen ${term}`;
    if (g === 'f') return `eine ${term}`;
    if (g === 'n') return `ein ${term}`;
    // Ohne Angabe wie bisher raten. Die Liste bleibt unveraendert, damit sich
    // an bestehenden Events nichts still verschiebt; maskuline Begriffe
    // liefern hier weiterhin „ein" — dafuer gibt es jetzt die Auswahl
    // im Assistenten.
    return /(session|veranstaltung|einheit|runde|reihe|tour|führung|schicht|woche|gruppe|stunde)$/i.test(term)
      ? `eine ${term}`
      : `ein ${term}`;
  }, [childTermSingular, event]);
  // v24.58: Anzeige-Präfix des Haupt-Events in der Sub-Event-Auswahl.
  // 'none' → kein Präfix (null), 'custom' → freier Text, sonst der mitgegebene
  // Default („Haupt-Event"/„Main event").
  const resolveMainEventLabel = React.useCallback((defaultLabel: string): string | null => {
    const mode = event && event.mainEventLabelMode;
    if (mode === 'none') return null;
    if (mode === 'custom' && event && event.mainEventLabel && event.mainEventLabel.trim()) return event.mainEventLabel.trim();
    return defaultLabel;
  }, [event]);
  // v24.73: Live-Plätze aus dem (für alle lesbaren) Sitzplatz-Counter. Die
  // Teilnehmerliste selbst ist item-level-gesichert — ein normaler Teilnehmer
  // sieht darüber NICHT die echte Gesamtzahl. Der Counter (aktiv = SeatsTaken,
  // Warteliste = WaitlistTaken) ist für alle lesbar und liefert die korrekten
  // Werte. Wird beim Öffnen + bei Fenster-Fokus leise nachgeladen (kein
  // sichtbares Nachladen — nur die Zahl ändert sich). Der Live-Push folgt in v24.74.
  // v29.62: Hover-Zustand der Tages-Kacheln im Anmelde-Kalender. Inline-Styles
  // koennen kein :hover (CLAUDE.md) — ohne diesen State lesen sich die Kacheln
  // wie Beschriftungen, nicht wie etwas Anklickbares. Schluessel ist der
  // Tagesschluessel (YYYY-MM-DD), nicht der Index: Der Index waere ueber
  // mehrere Monatsraster hinweg mehrdeutig.
  const [dayHoverKey, setDayHoverKey] = React.useState<string>('');
  const [liveStats, setLiveStats] = React.useState<{ active: number; waitlist: number } | null>(null);
  React.useEffect(() => {
    if (!event || !event.id || !(event.maxParticipants > 0)) { setLiveStats(null); return undefined; }
    let cancelled = false;
    const load = (): void => {
      getLiveCounterStats(event.id).then(s => { if (!cancelled && s) setLiveStats(s); }).catch(() => { /* best-effort */ });
    };
    load();
    const onFocus = (): void => load();
    window.addEventListener('focus', onFocus);
    // v24.75: Echtzeit-Push auf den (für alle lesbaren) Counter — bei jeder
    // An-/Abmeldung am Event meldet SharePoint die Counter-Änderung → der Wert
    // aktualisiert sich live, ohne Polling. Best-effort: klappt der Socket nicht,
    // bleibt der Lade-/Fokus-Refresh.
    let cleanupSocket: (() => void) | null = null;
    subscribeEventRealtime(event.id, 'counter', load)
      .then(c => { if (cancelled) c(); else cleanupSocket = c; })
      .catch(() => { /* best-effort */ });
    return () => {
      cancelled = true;
      window.removeEventListener('focus', onFocus);
      if (cleanupSocket) cleanupSocket();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [event?.id, event?.maxParticipants]);
  const [registerForParent, setRegisterForParent] = React.useState(true);
  const [selectedSessions, setSelectedSessions] = React.useState<Set<string>>(new Set());
  // v30.9: Schnappschuss der Auswahl zum Submit-Zeitpunkt. Die Erfolgsseite
  // listet „die folgenden Office-Tage" — las dafür aber den LIVE-State, und
  // der kann zwischen Submit und Render geleert/umgebaut werden (Refresh-
  // Effekte nach der Anmeldung). Ergebnis war eine Erfolgsseite OHNE die
  // Tages-Liste. Der Ref friert den Stand beim Klick auf „Anmelden" ein.
  const submittedSessionsRef = React.useRef<Set<string>>(new Set());
  // v30.66: Der Wert wird nur noch geschrieben (Reset beim Event-Wechsel) —
  // gelesen hat ihn allein die mit v30.66 entfernte Alt-Auswahl `{false && …}`.
  const [, setSessionStarterType] = React.useState<Record<string, string>>({});
  // v10.12+: pro Sub-Event eigene Custom-Field-Werte. Wird beim Check eines
  // Sub-Events in dem Pop-Up-Modal abgefragt (siehe pendingSubEventModal weiter
  // unten) und beim Submit pro Sub-Event-Registrierung an registerForEvent
  // weitergereicht.
  const [sessionFieldValues, setSessionFieldValues] = React.useState<Record<string, Record<string, string>>>({});
  // Modal-State: wenn ein Sub-Event angecheckt wird das Custom-Fields hat,
  // wird hier die ID gemerkt + ein Draft der Field-Werte. Beim „Bestätigen"
  // wandern die Werte in sessionFieldValues + die Session in selectedSessions.
  // Beim „Abbrechen" wird der Modal geschlossen und die Session NICHT angecheckt.
  const [pendingSubEventModal, setPendingSubEventModal] = React.useState<{
    subEventId: string;
    draftValues: Record<string, string>;
  } | null>(null);
  /**
   * v30.62: `count` ist bewusst `number | null`.
   *
   * Die Belegung kam bis hierher aus `getAllRegistrations(ce.id)` — also aus der
   * Teilnehmerliste der Subsite. Die hat Zeilen-Sicherheit: Ein normaler
   * Teilnehmer liest dort NUR die eigene Zeile, und `getAllRegistrations` wirft
   * bei 403 nicht, sondern liefert das bis dahin Gelesene (CLAUDE.md). Ergebnis
   * war `count = 0` für JEDEN Tag — also „80 frei" überall, während Organizer,
   * die die ganze Liste lesen dürfen, die echten Zahlen sahen. Genau dieser
   * Widerspruch wurde gemeldet.
   *
   * `null` heißt jetzt „nicht ermittelbar" und wird als Strich angezeigt. Eine
   * Zahl, die aus einem Leseverbot entsteht, ist keine Aussage über die Daten.
   */
  const [sessionMeta, setSessionMeta] = React.useState<Record<string, { count: number | null; wasRegistered: boolean }>>({});
  const [myParentReg, setMyParentReg] = React.useState<{ Status?: string } | null>(null);
  const [sessionsOnlySubmitted, setSessionsOnlySubmitted] = React.useState(false);
  // v18.67: echtes Anmelde-Ergebnis (Angemeldet/Warteliste) aus der
  // Haupt-Registrierung — das Ergebnis-Modal nutzt das statt der gecachten
  // isFull-Schätzung, die nach Cancel/Re-Register veraltet sein konnte und
  // fälschlich „Warteliste" zeigte, obwohl der User angemeldet wurde.
  const [submittedAsWaitlist, setSubmittedAsWaitlist] = React.useState(false);

  // v11.82: Team-Anmeldung — UI-State.
  // - isTeamMode: User hat den Toggle „Ich melde mich + mein Team an" angehakt.
  // - teamName: optionaler Team-Name (nur sichtbar wenn event.askTeamName).
  // - teamMembers: N-1 People-Picker-Slots, jeder „<DisplayName> <email>".
  // - teamConsentConfirmed: Pflicht-Checkbox „alle Mitglieder haben zugestimmt".
  // v14.5: Wenn der Organizer `requireSubEventSelection` aktiviert hat, ist
  // Team-Anmeldung nicht kombinierbar — der Team-Flow registriert nur fürs
  // Hauptevent (keine Sub-Event-Auswahl möglich), würde also entweder am
  // Submit-Gate scheitern (verwirrend) oder das Event-Setup unterlaufen.
  // Deshalb hier den Toggle hart ausblenden, damit die Inkonsistenz gar
  // nicht erst entsteht.
  const isTeamCapable = !!event?.teamRegistrationEnabled
    && (event?.teamSize || 0) >= 2
    // v22.78: Wenn Teilnehmer keine neuen Teams erstellen dürfen, wird der
    // „Ich melde mich + mein Team an"-Toggle ausgeblendet (Organizer ordnet zu).
    && !event?.teamMembersCannotCreate
    && !(event?.requireSubEventSelection && childEvents.length > 0);
  const teamSize = event?.teamSize || 0;
  const teamPartialAllowed = !!event?.teamPartialAllowed;
  const [isTeamMode, setIsTeamMode] = React.useState(false);
  const [teamName, setTeamName] = React.useState('');
  const [teamMembers, setTeamMembers] = React.useState<string[]>([]);
  // v18.12: Custom-Field-Antworten pro Team-Mitglied (Slot-Index → {fieldId: value}).
  // So kann der Lead z.B. die Essenspräferenz auch für jedes Teammitglied angeben.
  const [teamMemberFields, setTeamMemberFields] = React.useState<Record<number, Record<string, string>>>({});
  const [teamConsentConfirmed, setTeamConsentConfirmed] = React.useState(false);
  // v15.16: Bei „Für andere registrieren" (registerForOther) braucht es
  // ebenfalls eine explizite Bestätigung, dass die Person der Anmeldung
  // zugestimmt hat — analog zur Team-Anmelde-Pflicht.
  const [otherConsentConfirmed, setOtherConsentConfirmed] = React.useState(false);
  // v26.76: Geführter Wizard für die stellvertretende Anmeldung (interner Fall):
  // 0 = geschlossen, 1 = Person suchen, 2 = Zustimmung. Nach „OK" ist die Person
  // übernommen und die persönlichen Felder vorbefüllt.
  const [proxyStep, setProxyStep] = React.useState<0 | 1 | 2>(0);
  // v27.6/v30.66: Der frühere INLINE-Ablauf für „Für andere Person anmelden"
  // (Personensuche + Extern-Umschalter + Zustimmungs-Box direkt im Formular) ist
  // vollständig in den geführten Wizard-Modal gewandert; die Alt-UI hing an einem
  // `useState(false)` ohne Setter und ist mit v30.66 entfernt.
  // v11.83: Offene Teams (Slots-frei) + Beitritts-Flow.
  const [openTeams, setOpenTeams] = React.useState<Array<{ teamId: string; teamName: string; activeCount: number; teamSize: number; leadEmail: string; leadDisplayName: string }>>([]);
  const [openTeamsLoaded, setOpenTeamsLoaded] = React.useState(false);
  // v18.73: Beitritt zu einem offenen Team wird nur VORGEMERKT — die eigentliche
  // Anmeldung (inkl. der ausgefüllten persönlichen + event-spezifischen Felder)
  // passiert erst beim Klick auf „Anmelden" unten (performJoinSelectedTeam).
  // Vorher wurde der Beitritt sofort beim Klick committet, ohne dass der User
  // seine event-spezifischen Infos angeben konnte.
  const [pendingJoinTeam, setPendingJoinTeam] = React.useState<{ teamId: string; teamName: string } | null>(null);
  // v18.73: Erfolgsscreen-Variante bei Team-Beitritt ('joined' = direkt
  // angemeldet, 'requested' = Anfrage an den Team-Lead gesendet).
  const [submittedJoinKind, setSubmittedJoinKind] = React.useState<null | 'joined' | 'requested'>(null);
  // Beim Aktivieren des Team-Modus: Member-Slots initialisieren (teamSize-1 Slots).
  React.useEffect(() => {
    if (isTeamMode && teamMembers.length !== Math.max(0, teamSize - 1)) {
      setTeamMembers(Array.from({ length: Math.max(0, teamSize - 1) }, () => ''));
    }
    if (!isTeamMode) {
      setTeamMembers([]);
      setTeamName('');
      setTeamConsentConfirmed(false);
      setTeamMemberFields({});
    }
  }, [isTeamMode, teamSize]);
  // Parser für People-Picker-Values im Format „DisplayName <email>".
  const parseTeamMember = (v: string): { displayName: string; email: string } | null => {
    const m = (v || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
    if (!m) return null;
    return { displayName: m[1].trim(), email: m[2].trim().toLowerCase() };
  };
  const teamMembersParsed = teamMembers.map(parseTeamMember);
  // v18.12: Custom-Fields, die pro Team-Mitglied abgefragt werden — alle
  // event-spezifischen Felder AUSSER Personen-Pickern (user/roommate) und
  // B2Run-Spezialfeldern; gruppen-spezifische Felder nur „für alle".
  const teamMemberApplicableFields = (event?.eventSpecificFields || []).filter(f =>
    f.type !== 'user' && f.type !== 'roommate' &&
    f.id !== 'b2run_startblock' && f.id !== 'b2run_mobilnummer' &&
    (!f.onlyForGroup || f.onlyForGroup === 'all')
  );
  // Validation des Team-Submits — Lead-Email darf nicht in der Member-Liste
  // sein, Member-Emails müssen untereinander disjunkt sein, im Pflicht-Modus
  // müssen alle Slots gefüllt sein.
  const teamValidation = ((): { ok: boolean; reason?: string } => {
    if (!isTeamMode) return { ok: true };
    const leadEmail = (email || '').trim().toLowerCase();
    const filled = teamMembersParsed.filter(m => !!m);
    if (!teamPartialAllowed && filled.length < (teamSize - 1)) {
      return { ok: false, reason: locale === 'de' ? `Bitte alle ${teamSize - 1} Team-Mitglieder auswählen.` : `Please pick all ${teamSize - 1} team members.` };
    }
    const seen = new Set<string>();
    if (leadEmail) seen.add(leadEmail);
    for (const m of filled) {
      if (!m) continue;
      if (seen.has(m.email)) {
        return { ok: false, reason: locale === 'de' ? `„${m.displayName}" ist doppelt im Team.` : `„${m.displayName}" appears twice in the team.` };
      }
      seen.add(m.email);
    }
    if (event?.askTeamName && !teamName.trim()) {
      return { ok: false, reason: locale === 'de' ? 'Bitte Team-Name angeben.' : 'Please enter a team name.' };
    }
    if (teamName.trim().length > 60) {
      return { ok: false, reason: locale === 'de' ? 'Team-Name max. 60 Zeichen.' : 'Team name must be 60 characters or fewer.' };
    }
    if (!teamConsentConfirmed) {
      return { ok: false, reason: locale === 'de' ? 'Bitte bestätige, dass alle Mitglieder zugestimmt haben.' : 'Please confirm that all members have consented.' };
    }
    return { ok: true };
  })();

  // v11.83: Offene Teams nachladen — sobald wir wissen, dass das Event
  // Team-Anmeldung erlaubt UND der Organizer „Offene Slots oeffentlich
  // sichtbar" aktiviert hat UND der User selbst noch nicht angemeldet
  // ist. Lazy: einmal pro Event-Wechsel.
  React.useEffect(() => {
    setOpenTeamsLoaded(false);
    setOpenTeams([]);
    setPendingJoinTeam(null); // v18.73: Vormerkung beim Event-/Modus-Wechsel zurücksetzen
    if (!event) return;
    if (!event.teamRegistrationEnabled || !event.teamOpenSlotsVisible) return;
    if (registerForOther) return; // Stellvertreter-Modus nicht unterstützt für Beitritt
    (async () => {
      try {
        const list = await listOpenTeamsForEvent(event.id);
        setOpenTeams(list);
      } catch {
        setOpenTeams([]);
      } finally {
        setOpenTeamsLoaded(true);
      }
    })().catch(() => setOpenTeamsLoaded(true));
  }, [event?.id, event?.teamRegistrationEnabled, event?.teamOpenSlotsVisible, registerForOther, listOpenTeamsForEvent]);

  // v18.73: Team-Beitritt nur VORMERKEN (Toggle). Erneuter Klick auf dasselbe
  // Team hebt die Vormerkung wieder auf. Gegenseitig exklusiv zum „Ich melde
  // mich + mein Team an"-Modus (man kann nicht gleichzeitig ein neues Team
  // anlegen und einem bestehenden beitreten). Die eigentliche Anmeldung läuft
  // erst über den „Anmelden"-Button (performJoinSelectedTeam).
  const togglePendingJoinTeam = (teamId: string, teamName: string): void => {
    setError('');
    setPendingJoinTeam(prev => (prev && prev.teamId === teamId) ? null : { teamId, teamName });
    setIsTeamMode(false);
  };

  // Vorbelegen: Parent-Reg prüfen + Sessions-Meta laden (bereits-registrierte
  // Sessions werden als angehakt voreingestellt).
  React.useEffect(() => {
    if (!event) return;
    // Parent-Reg-Vorbelegung nur im Selbst-Modus — sie beruht auf
    // getMyRegistration des eingeloggten Users und ist im Stellvertreter-
    // Modus bedeutungslos.
    if (!registerForOther) {
      (async () => {
        try {
          const r = await getMyRegistration(event.id) as { Status?: string; StarterType?: string; PreferredStarterType?: string } | null;
          setMyParentReg(r);
          if (r && r.Status !== 'Abgemeldet') {
            setRegisterForParent(false);
            // v11.10: Bei bereits angemeldetem Parent den existierenden
            // Starter-Typ in die Group-Selection vorladen, damit auch im
            // Sessions-Only-Modus eine Gruppe sichtbar gewählt ist und
            // Sub-Events sauber davon erben.
            const existing = r.StarterType || r.PreferredStarterType;
            if (existing && (existing === 'Durchstarter' || existing === 'Funstarter')) {
              setPreferredStarterType(existing);
            }
          }
        } catch { /* */ }
      })();
    }
    if (childEvents.length > 0) {
      (async () => {
        try {
          const meta: Record<string, { count: number | null; wasRegistered: boolean }> = {};
          const preselect = new Set<string>();
          const starterPre: Record<string, string> = {};
          // v29.65: Diese Schleife lief STRENG NACHEINANDER, und jeder Durchlauf
          // macht zwei Requests (eigene Anmeldung + alle Anmeldungen). Bei einem
          // Kalender-Event mit 21 Tagen sind das 42 Roundtrips, bevor
          // `setSessionMeta` ueberhaupt einmal aufgerufen wird — und bis dahin
          // faellt jede Tages-Kachel auf `{ count: 0, wasRegistered: false }`
          // zurueck: alle zeigen die volle Platzzahl, und die eigenen, schon
          // gebuchten Tage sind NICHT vorausgewaehlt. Genau das ist der Eindruck
          // „die Maske sieht nicht so aus, wie sie soll".
          //
          // Jetzt laufen die Tage zu sechst (dieselbe Grenze wie beim
          // Teilnehmerzahlen-Nachlauf im EventContext) und die Zahlen werden
          // schrittweise veroeffentlicht, sobald ein Tag fertig ist.
          //
          // Die VORAUSWAHL bleibt bewusst ein einziger Aufruf am Ende: Sie
          // ueberschreibt die Auswahl des Nutzers, und die haeppchenweise
          // nachzuziehen wuerde mit dem konkurrieren, was er waehrenddessen
          // anklickt.
          /**
           * v30.62: Die Belegung eines Termins — aus einer Quelle, die JEDER
           * lesen darf.
           *
           * Erste Wahl ist der Platzzähler `DEX_TeilnehmerCounter` der Subsite.
           * Er ist genau dafür da und für alle lesbar; das Haupt-Event benutzt
           * ihn längst (`liveStats`). Die Termin-Kacheln zählten stattdessen
           * über die Teilnehmerliste — und die ist zeilenweise gesichert. Wer
           * nur die eigene Zeile sehen darf, zählt eine oder keine, und die
           * Kachel meldete „80 frei", während die Organizerin daneben die
           * echten Zahlen sah.
           *
           * Die Liste bleibt der Rückfall, aber nur mit geprüftem Status: Ohne
           * `onHttpError` wäre ein Leseverbot von einer leeren Liste nicht zu
           * unterscheiden — genau die Verwechslung, die den Fehler erzeugt hat.
           * Trägt keine der beiden Quellen, kommt `null` zurück und die Kachel
           * zeigt einen Strich statt einer erfundenen Zahl.
           */
          const occupancyOf = async (subId: string): Promise<number | null> => {
            // v30.63: Nachvollziehbar in der Browser-Konsole. Bei einer
            // gemeldeten Abweichung soll man SEHEN, aus welcher Quelle die Zahl
            // kam — sonst rät man wieder zwischen „Daten falsch" und „Anzeige
            // falsch", und genau das hat diesen Fehler zwei Runden gekostet.
            const title = (childEvents.find(c => c.id === subId) || { title: subId }).title;
            const cap = (childEvents.find(c => c.id === subId) || { maxParticipants: 0 }).maxParticipants || 0;
            try {
              const stats = await getLiveCounterStats(subId);
              // `seatsKnown` trennt „null Anmeldungen" von „nie geschrieben".
              // Ohne diese Unterscheidung wäre der ungenutzte Zähler genau die
              // 0, die den ursprünglichen Fehler erzeugt hat.
              if (stats && stats.seatsKnown && typeof stats.active === 'number' && stats.active >= 0) {
                // eslint-disable-next-line no-console
                dlog('seats', `[DEX][seats] "${title}" — Quelle: Platzzähler · belegt ${stats.active}/${cap} · frei ${Math.max(0, cap - stats.active)} · Warteliste ${stats.waitlist < 0 ? 'unbekannt' : stats.waitlist}`);
                return stats.active;
              }
              // eslint-disable-next-line no-console
              dlog('seats', `[DEX][seats] "${title}" — Platzzähler ${stats ? 'nie beschrieben (seatsKnown=false)' : 'nicht lesbar'} → Rückfall auf die Teilnehmerliste`);
            } catch (e) {
              // eslint-disable-next-line no-console
              dlog('seats', `[DEX][seats] "${title}" — Platzzähler-Abruf fehlgeschlagen → Rückfall auf die Teilnehmerliste`, e);
            }
            let readable = true;
            try {
              const regs = await getAllRegistrations(subId, (status) => {
                readable = false;
                // eslint-disable-next-line no-console
                dlog('seats', `[DEX][seats] "${title}" — Teilnehmerliste nicht (vollständig) lesbar, HTTP ${status}. Keine Zahl anzeigen.`);
              });
              if (!readable) return null;
              const n = (regs || []).filter(r => {
                const st = r.Status || '';
                return st === 'Angemeldet' || st === 'QR versendet' || st === 'Eingecheckt';
              }).length;
              // eslint-disable-next-line no-console
              dlog('seats', `[DEX][seats] "${title}" — Quelle: Teilnehmerliste · belegt ${n}/${cap} · frei ${Math.max(0, cap - n)} · ${regs.length} Zeile(n) gelesen`);
              return n;
            } catch (e) {
              // eslint-disable-next-line no-console
              dlog('seats', `[DEX][seats] "${title}" — auch die Teilnehmerliste nicht lesbar. Keine Zahl anzeigen.`, e);
              return null;
            }
          };
          const runLimited = async (items: typeof childEvents, limit: number, fn: (x: typeof childEvents[number]) => Promise<void>): Promise<void> => {
            let next = 0;
            const workers = new Array(Math.min(limit, items.length)).fill(0).map(async () => {
              for (;;) {
                const i = next++;
                if (i >= items.length) return;
                await fn(items[i]);
              }
            });
            await Promise.all(workers);
          };
          await runLimited(childEvents, 6, async (ce) => {
            if (registerForOther) {
              // v18.37: Stellvertreter-Modus — nur die Belegungszahl laden.
              // KEINE Self-Vorbelegung, weil getMyRegistration die Daten des
              // eingeloggten Users liefert, nicht die der angemeldeten Person.
              meta[ce.id] = { count: await occupancyOf(ce.id), wasRegistered: false };
            } else {
              const [myReg, count] = await Promise.all([
                getMyRegistration(ce.id) as Promise<{ Status?: string; StarterType?: string; PreferredStarterType?: string } | null>,
                occupancyOf(ce.id),
              ]);
              const wasRegistered = !!myReg && myReg.Status !== 'Abgemeldet';
              meta[ce.id] = { count, wasRegistered };
              if (wasRegistered) {
                preselect.add(ce.id);
                const existingType = myReg?.StarterType || myReg?.PreferredStarterType;
                if (existingType) starterPre[ce.id] = existingType;
              }
            }
            // Schrittweise veroeffentlichen — die Kachel dieses Tages zeigt ihre
            // echten Zahlen, sobald sie da sind, statt auf alle anderen zu warten.
            const justDone = meta[ce.id];
            if (justDone) setSessionMeta(prev => ({ ...prev, [ce.id]: justDone }));
          });
          setSessionMeta(meta);
          // Im Stellvertreter-Modus startet die Auswahl leer (frische Anmeldung).
          setSelectedSessions(registerForOther ? new Set<string>() : preselect);
          setSessionStarterType(prev => ({ ...starterPre, ...prev }));
        } catch { /* */ }
      })();
    }
  }, [event?.id, registerForOther]);

  // v28.4: Seitenverhältnis des Event-Bildes erkennen — Querformat-Fotos
  // bekommen im „Geführte Schritte"-Layout einen BREITEREN Bild-Slot (420px
  // statt 300px), damit sie nicht winzig in der Ecke hängen; Hochkant/
  // Quadrat bleibt beim kompakten 300er-Slot. Das Bild sitzt vertikal
  // mittig neben den Infos (kein toter Leerraum mehr unter dem Foto).
  // v28.19: Kein Layout-Umspringen mehr — die Analyse läuft asynchron, daher:
  // (a) Ergebnis pro URL im Modul-Cache (IMG_ASPECT_CACHE), damit der zweite
  //     Besuch synchron im ersten Render die richtige Form kennt, und
  // (b) `imgAspectReady` als Gate: Der Bild-Slot wird erst gerendert, wenn
  //     die Form feststeht (oder die Analyse fehlschlug) — das Bild erscheint
  //     dann direkt an der richtigen Stelle statt kurz rechts zu starten.
  const [imgProbe, setImgProbe] = React.useState<{ url: string; ratio: number | null } | null>(null);
  const imgAspectCached = heroImgUrl ? IMG_ASPECT_CACHE[heroImgUrl] : undefined;
  const imgAspect: number | null = imgAspectCached !== undefined
    ? imgAspectCached
    : (imgProbe && imgProbe.url === heroImgUrl ? imgProbe.ratio : null);
  const imgAspectReady = imgAspectCached !== undefined
    || (!!imgProbe && imgProbe.url === heroImgUrl);
  React.useEffect(() => {
    if (!heroImgUrl) return undefined;
    if (IMG_ASPECT_CACHE[heroImgUrl] !== undefined) return undefined;
    let cancelled = false;
    const probeUrl = heroImgUrl;
    const img = new Image();
    img.onload = () => {
      if (cancelled || img.naturalHeight <= 0) return;
      const fileRatio = img.naturalWidth / img.naturalHeight;
      let ratio = fileRatio;
      // v28.9: CONTENT-Ratio statt reiner Datei-Ratio. Logos/Kreis-Grafiken
      // liegen oft mit transparentem oder einfarbigem Rand in einer breiten
      // Datei — die Datei-Ratio sortierte sie als „Querformat" ein und das
      // Kreis-Layout (v28.7) griff nie. Wir rastern das Bild klein, prüfen
      // ob die vier Ecken einen einheitlichen Rand bilden (transparent oder
      // eine Farbe), trimmen diesen Rand und nehmen das Seitenverhältnis
      // des sichtbaren Inhalts. Randlose Fotos (uneinige Ecken) und
      // Canvas-Fehler behalten die Datei-Ratio.
      try {
        const W = 96;
        const H = Math.max(1, Math.round(W / fileRatio));
        const canvas = document.createElement('canvas');
        canvas.width = W; canvas.height = H;
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.drawImage(img, 0, 0, W, H);
          const data = ctx.getImageData(0, 0, W, H).data;
          const px = (x: number, y: number): number[] => {
            const i = (y * W + x) * 4;
            return [data[i], data[i + 1], data[i + 2], data[i + 3]];
          };
          const corners = [px(0, 0), px(W - 1, 0), px(0, H - 1), px(W - 1, H - 1)];
          const dist = (a: number[], b: number[]): number =>
            Math.max(Math.abs(a[0] - b[0]), Math.abs(a[1] - b[1]), Math.abs(a[2] - b[2]));
          const allTransparent = corners.every(c => c[3] <= 24);
          const uniform = allTransparent || corners.every(c => c[3] > 24 && dist(c, corners[0]) <= 20);
          if (uniform) {
            const bg = corners[0];
            let minX = W, minY = H, maxX = -1, maxY = -1;
            for (let y = 0; y < H; y++) {
              for (let x = 0; x < W; x++) {
                const p = px(x, y);
                const isContent = p[3] > 24 && (allTransparent || dist(p, bg) > 28);
                if (isContent) {
                  if (x < minX) minX = x;
                  if (x > maxX) maxX = x;
                  if (y < minY) minY = y;
                  if (y > maxY) maxY = y;
                }
              }
            }
            if (maxX >= minX && maxY >= minY) {
              const bw = maxX - minX + 1;
              const bh = maxY - minY + 1;
              // Nur übernehmen, wenn wirklich Rand weggefallen ist — sonst
              // ist die Datei-Ratio genauer (voller Inhalt bis zur Kante).
              if (bh > 0 && (bw < W || bh < H)) ratio = bw / bh;
            }
          }
        }
      } catch { /* tainted canvas o.ä. → Datei-Ratio behalten */ }
      IMG_ASPECT_CACHE[probeUrl] = ratio;
      setImgProbe({ url: probeUrl, ratio });
    };
    // Ladefehler: Form bleibt unbekannt (ratio null → Standard-Slot), aber
    // das Gate öffnet, damit das Bild/Fallback nicht dauerhaft versteckt ist.
    img.onerror = () => { if (!cancelled) setImgProbe({ url: probeUrl, ratio: null }); };
    img.src = probeUrl;
    return () => { cancelled = true; };
  }, [heroImgUrl]);
  // v28.6: Slot-Größe hängt von der BILDFORM ab — Kreis-/Quadrat-Bilder
  // (Ratio ~1, z.B. aus dem Zuschnitt-Tool) brauchen keinen 300er-Block,
  // Querformat bekommt Breite, Hochkant Höhe.
  const imgSlotW = imgAspect == null ? 280 : (imgAspect >= 1.2 ? 420 : (imgAspect >= 0.8 ? 210 : 240));
  const imgSlotH = imgAspect == null ? 260 : (imgAspect >= 1.2 ? 260 : (imgAspect >= 0.8 ? 210 : 300));
  // v28.7: Kreis-/Quadrat-Bilder (Ratio ~1, typisch der Kreis-Zuschnitt aus
  // dem Wizard mit transparenten Ecken) sitzen NICHT mehr seitlich neben den
  // Infos, sondern als eigener Kreis OBEN MITTIG, der die Oberkante der
  // Event-Karte überlappt („eingebautes" Profilbild-Muster). Banner-Modus
  // hat weiter Vorrang; Quer-/Hochformat behält den Seiten-Slot.
  // v29.70: Der Kreis ist jetzt der DEFAULT — nicht mehr nur fuer ~quadratische
  // Bilder (Ratio 0,8-1,2). Vorher musste der Organizer sein Foto im Wizard
  // extra rund zuschneiden, damit es den Kreis bekam; jedes normale Querformat
  // landete im Seiten-Slot. Jetzt schneidet das CSS (objectFit: cover in der
  // runden Maske) jedes Bild mittig in den Kreis. Zwei bewusste Ausnahmen, weil
  // sie AKTIVE Entscheidungen des Organizers sind: der Banner-Modus und eine
  // gepflegte Hero-Darstellung (Zoom/Hoehe aus v23.19). Wer das ganze Bild
  // sehen will, hat weiterhin die Lupe (Lightbox, v28.12).
  // Nebeneffekt: Die Bedingung haengt nicht mehr an imgAspect — der Kreis
  // steht sofort, statt erst nach der Bildvermessung das Layout zu wechseln.
  const imgCircleNotch = !!heroImgUrl && !event?.imageBanner && !event?.imageDisplay?.hero;
  const circleSize = isMobile ? 140 : 170;
  // v28.91: Kein Event-Foto → das DEX-Bild steht als KREIS oben mittig,
  // genau dort, wo auch ein rundes Event-Logo sitzt (imgCircleNotch). Im
  // Seiten-Slot rechts wirkte es wie ein Foto des Events, das es nicht ist.
  const showOrbPlaceholder = !heroImgUrl;

  // B2Run Split-Capacity: aktuelle Auslastung pro Typ laden
  // Split-UI nur wenn BEIDE Starter-Typen verfügbar sind (>0). Wenn der Admin eine
  // Kapazität auf 0 gesetzt hat, gibt es faktisch nur einen Typ — dann keine Auswahl
  // anzeigen und den einzig verfügbaren Typ automatisch setzen (siehe useEffect unten).
  const durchCap = (event && typeof event.durchstarterCapacity === 'number') ? event.durchstarterCapacity : 0;
  const funCap = (event && typeof event.funstarterCapacity === 'number') ? event.funstarterCapacity : 0;
  const isSplitGroup = !!event && durchCap > 0 && funCap > 0;
  // v10.20: frei wählbare Bezeichnungen aus dem Event laden, mit Fallback auf
  // die historischen B2Run-Defaults 'Durchstarter' / 'Funstarter'. Die internen
  // Werte für SP-Persistenz (StarterType-Spalte) bleiben unverändert — das
  // Label ist reines UI.
  const splitLabelA = (event?.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
  const splitLabelB = (event?.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
  const singleStarterType: string = (!event || (durchCap <= 0 && funCap <= 0))
    ? '' // kein B2Run-Event überhaupt
    : (durchCap > 0 && funCap <= 0) ? 'Durchstarter'
    : (funCap > 0 && durchCap <= 0) ? 'Funstarter'
    : ''; // beide > 0 -> User muss wählen (Split-UI)

  // Auto-Set: wenn nur ein Starter-Typ verfügbar ist, direkt diesen Typ als
  // preferredStarterType speichern — damit registerForEvent ihn trotzdem auf den
  // Teilnehmer-Eintrag schreiben kann, obwohl das Split-UI nicht angezeigt wird.
  React.useEffect(() => {
    if (singleStarterType && preferredStarterType !== singleStarterType) {
      setPreferredStarterType(singleStarterType);
    }
  }, [singleStarterType]);

  // v6.15: Starter-Typ → Startblock-Auto-Mapping. Wenn der Admin für dieses Event
  // einen Block an den Starter-Typ gebunden hat, wird das zugehörige
  // b2run_startblock-Custom-Field automatisch gesetzt — der User muss den Block
  // nicht extra auswählen (das Custom-Field wird dann im UI ausgeblendet).
  const durchstarterBlock = event?.durchstarterStartblock || '';
  const funstarterBlock = event?.funstarterStartblock || '';
  const hasStarterBlockMapping = !!(durchstarterBlock || funstarterBlock);
  React.useEffect(() => {
    if (!hasStarterBlockMapping || !preferredStarterType) return;
    const mappedBlock = preferredStarterType === 'Durchstarter' ? durchstarterBlock : funstarterBlock;
    if (!mappedBlock) return;
    if (eventSpecific.b2run_startblock === mappedBlock) return;
    setEventSpecific(prev => ({ ...prev, b2run_startblock: mappedBlock }));
  }, [preferredStarterType, durchstarterBlock, funstarterBlock, hasStarterBlockMapping]);
  // v26.74: Vorauswahl bei Single-Select-Feldern — den vom Organizer gesetzten
  // Default einmal vorbelegen (nur wenn der Teilnehmer das Feld noch nicht
  // berührt hat; ein bewusstes Leeren bleibt erhalten). Läuft, sobald die
  // Event-Felder verfügbar sind bzw. sich ihre Defaults ändern.
  const selectDefaultsSig = (event?.eventSpecificFields || [])
    .filter(f => f.type === 'select' && !f.multi && f.defaultValue)
    .map(f => `${f.id}=${f.defaultValue}`).join('|');
  React.useEffect(() => {
    const fields = event?.eventSpecificFields || [];
    const defaults: Record<string, string> = {};
    for (const f of fields) {
      if (f.type === 'select' && !f.multi && f.defaultValue && (f.options || []).indexOf(f.defaultValue) >= 0) {
        defaults[f.id] = f.defaultValue;
      }
    }
    if (Object.keys(defaults).length === 0) return;
    setEventSpecific(prev => {
      let changed = false;
      const next = { ...prev };
      for (const k of Object.keys(defaults)) {
        if (next[k] === undefined) { next[k] = defaults[k]; changed = true; }
      }
      return changed ? next : prev;
    });
  }, [selectDefaultsSig]);
  React.useEffect(() => {
    if (!isSplitGroup || !event?.subsiteUrl) return;
    (async () => {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ctx = (window as any).__dexSpfxContext;
        if (!ctx) return;
        const { EventService } = await import('../services/EventService');
        const svc = new EventService(ctx);
        const allRegs = await svc.getAllRegistrations(event.subsiteUrl!);
        const active = allRegs.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        // v19.19: Warteliste pro Gruppe über die effektive Gruppe
        // (StarterType || PreferredStarterType) — Wartelisten-Einträge haben
        // i.d.R. noch keinen StarterType, ihr Gruppen-Wunsch steht in
        // PreferredStarterType.
        const waiting = allRegs.filter(r => r.Status === 'Warteliste');
        const effGroup = (r: { StarterType?: string; PreferredStarterType?: string }): string => r.StarterType || r.PreferredStarterType || '';
        setStarterCounts({
          durch: active.filter(r => r.StarterType === 'Durchstarter').length,
          fun: active.filter(r => r.StarterType === 'Funstarter').length,
          durchWait: waiting.filter(r => effGroup(r) === 'Durchstarter').length,
          funWait: waiting.filter(r => effGroup(r) === 'Funstarter').length,
        });
      } catch { /* ignore */ }
    })();
  }, [isSplitGroup, event?.subsiteUrl]);
  // v19.17: Der frühere 5-Sekunden-Poll auf dem Anmelde-Screen wurde wieder
  // entfernt — er verursachte einen sichtbaren Re-Render. Die Belegungszahl
  // kommt jetzt aus dem Context-Stand beim Öffnen (die Übersicht lädt sie beim
  // Navigieren frisch nach). Kein Polling, kein Refresh, Formular bleibt stabil.
  // Deloitte-Mitarbeitersuche
  const [userSearch, setUserSearch] = React.useState('');
  const [userResults, setUserResults] = React.useState<Array<{ email: string; displayName: string; location: string; jobTitle: string }>>([]);
  // v11.97: nach Picker-Auswahl im "Für andere Person registrieren"-Modus
  // halten wir das volle Profil (Department + Mobile zusätzlich), damit
  // die Personal-Info-Card die gleichen Read-only-Felder zeigt wie beim
  // Self-Register-Modus.
  const [pickedUserProfile, setPickedUserProfile] = React.useState<{
    jobTitle?: string;
    department?: string;
    location?: string;
    mobilePhone?: string;
    // v28.11: Unternehmenszugehörigkeit der ausgewählten Person — vorher
    // fehlte das Feld und die Profil-Karte zeigte „— nicht hinterlegt".
    company?: string;
  } | null>(null);
  const [isSearchingUser, setIsSearchingUser] = React.useState(false);
  const [userSearchIncludeIntl, setUserSearchIncludeIntl] = React.useState(false);
  const searchTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  // v30.4: Die folgenden Hooks standen bis v30.3 weiter unten — HINTER den
  // bedingten Returns (!event, notYetActive, isFullyClosed). Das ging nur gut,
  // solange diese Bedingungen während eines Seitenbesuchs konstant blieben;
  // die neue User-Vorschau kippt isOrganizer/isAdmin aber zur LAUFZEIT und
  // damit die Return-Bedingungen → React #300/#310 (weißer Bildschirm, siehe
  // v30.3). Deshalb: In dieser Komponente stehen ALLE Hooks oberhalb von
  // `if (!event)` — keine Ausnahme.
  // v9.22: Warning-Modal für externe Email-Anmeldung (durch Organizer für
  // Drittpersonen die noch kein Deloitte-Postfach haben). Default: nicht
  // erlaubt; Organizer kann nach Bestätigung trotzdem fortfahren — die
  // Bestätigungsmail geht dann nicht an die externe Adresse, sondern an
  // den Organizer mit Datenschutz-Hinweis-Header.
  const [externalEmailWarning, setExternalEmailWarning] = React.useState(false);
  // v18.75: Sicherheitshinweis vor dem Absenden (pro Event konfiguriert). Der
  // Dialog erscheint nach dem „Anmelden"-Klick und vor der eigentlichen
  // (Normal-)Anmeldung. confirmDraft* halten die — in der Auswahl-Übersicht
  // editierbare — Auswahl, bis der User bestätigt.
  const [confirmDialogOpen, setConfirmDialogOpen] = React.useState(false);
  const [confirmDialogAck, setConfirmDialogAck] = React.useState(false);
  const [confirmDraftParent, setConfirmDraftParent] = React.useState(true);
  const [confirmDraftSessions, setConfirmDraftSessions] = React.useState<Set<string>>(new Set());
  const confirmDialogConfirmedRef = React.useRef(false);
  const externalEmailConfirmedRef = React.useRef(false);
  // v19.6: Bei stellvertretender Anmeldung einer INTERNEN Person (Deloitte)
  // fragt nach dem „Anmelden"-Klick ein Modal, ob der/die Anmeldende (Organizer,
  // Co-Organizer oder Assistenz) selbst auf CC der Bestätigungs-Mail gesetzt
  // werden soll. ccSelfDecidedRef merkt sich, dass die Frage in diesem
  // Submit-Durchlauf bereits beantwortet wurde (analog confirmDialogConfirmedRef);
  // ccSelfRef hält die Entscheidung (true = auf CC).
  const [ccSelfModalOpen, setCcSelfModalOpen] = React.useState(false);
  const ccSelfDecidedRef = React.useRef(false);
  const ccSelfRef = React.useRef(false);
  // v24.48: Assistenz-Abfrage als Modal beim Register-Klick (Partner/Director).
  const [assistantModalOpen, setAssistantModalOpen] = React.useState(false);
  const assistantModalDecidedRef = React.useRef(false);
  // v24.49: Auswahl SYNCHRON im Ref festhalten — der Re-Submit aus dem Modal
  // läuft sonst mit dem alten State-Wert (setState ist async) und die CC würde
  // verloren gehen. { enabled, value } wird beim Klick im Modal gesetzt.
  const delegateChoiceRef = React.useRef<{ enabled: boolean; value: string } | null>(null);
  /**
   * v29.40: Personensuche, die nur Personen aus dem Verteilerkreis des Events
   * liefert — für Felder mit `audienceOnly` (typisch: Zimmerpartner). Es ist
   * dieselbe Prüfung wie beim Anmelden für andere (`isEventVisibleForUser`),
   * damit Feld und Anmeldung nicht unterschiedlich urteilen.
   *
   * Die Treffer liefern Standort und Position mit; ohne diese Angaben kann ein
   * reiner Standortfilter nicht greifen — dann bleibt die Person draußen, was
   * die sichere Richtung ist (lieber jemanden zu wenig anbieten als eine
   * Person, die gar nicht eingeladen ist).
   */
  const searchUsersInAudience = React.useCallback(async (q: string, includeIntl?: boolean) => {
    const res = await searchUsers(q, includeIntl);
    if (!event) return res;
    return res.filter(u => isEventVisibleForUser(event, u.email, u.location || '', [], u.jobTitle || ''));
  }, [searchUsers, event]);

  if (!event) {
    // v28.7: Beim Browser-Refresh restauriert der NavigationContext die
    // Anmeldeseite SOFORT, während die Events noch aus SharePoint laden —
    // vorher stand dann fälschlich „Event nicht gefunden". Solange die
    // Events laden, zeigen wir den Spinner (gleiches Muster wie die
    // Event-Liste); „nicht gefunden" kommt erst, wenn das Event nach dem
    // Laden wirklich fehlt.
    if (isEventsLoading) {
      return (
        <div className="page-container text-center">
          <div style={{ padding: 48 }}>
            <svg width={48} height={48} viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" style={{ display: 'block', margin: '0 auto 16px' }}>
              <circle cx={24} cy={24} r={20} fill="none" stroke="rgba(134,188,37,0.20)" strokeWidth={4} />
              <path d="M 24 4 A 20 20 0 0 1 44 24" fill="none" stroke="#86bc25" strokeWidth={4} strokeLinecap="round">
                <animateTransform attributeName="transform" type="rotate" from="0 24 24" to="360 24 24" dur="1s" repeatCount="indefinite" />
              </path>
            </svg>
            <p style={{ color: 'var(--dex-gray-400)' }}>{locale === 'de' ? 'Event wird geladen …' : 'Loading event …'}</p>
          </div>
        </div>
      );
    }
    return (
      <div className="page-container text-center">
        <h2>{t('reg.eventnotfound')}</h2>
        <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
          {t('reg.backtoevents')}
        </button>
      </div>
    );
  }

  // Registrierungs-Deadline prüfen.
  // v22.54: „Anmeldung geschlossen" greift nur, wenn das Hauptevent UND alle
  // Sub-Events zu sind. Solange mindestens ein Sub-Event noch offen ist, kommt
  // der Teilnehmer rein und kann sich für die offenen Sub-Events anmelden —
  // auch wenn die (Klammer-/Hauptevent-)Frist abgelaufen ist.
  // v28.20: Auch die explizite Klammer-Frist zählt (Organizer/Admin-Banner +
  // Parent-Reg-Block; für reguläre User greift ohnehin die Fully-Closed-Seite).
  const isDeadlinePassed = (!!event.registrationDeadline && new Date(event.registrationDeadline) < new Date())
    || (!!event.klammerDeadline && new Date(event.klammerDeadline) < new Date());
  // v30.20: Über ALLE buchbaren Sub-Events rechnen, NICHT über die
  // sichtbarkeitsgefilterte childEvents-Liste. Sonst kippt die Entscheidung
  // mit der Sichtbarkeit des Betrachters: In der User-Vorschau (und für User
  // mit engem Verteiler) fielen alle Tage aus der Liste, die leere Liste
  // zählte als „kein Sub-Event mehr offen", und die Alt-Frist der
  // Klammer-Spalte (25.08.) zeigte „Anmeldefrist abgelaufen" — obwohl
  // Termine offen waren, die die Person nur nicht sehen darf. Für „für dich
  // ist keines freigegeben" gibt es die v29.9-Meldung, nicht die Frist-Seite.
  const isFullyClosed = isRegistrationFullyClosed(
    event,
    event.subEventsDisabled ? [] : childEventsOf(event.id).filter(ce => !ce.isFictive)
  );

  // v23.14: Vorschau vor Aktivierung — reguläre User dürfen die Anmeldeseite
  // erst ab dem „Aktiv ab"-Zeitpunkt öffnen (Deep-Link-Schutz; die Karte
  // blockiert den Klick ohnehin). Organizer/Admin dürfen vorbereiten.
  const notYetActive = !!event.activeFrom && new Date(event.activeFrom) > new Date();
  if (notYetActive && !isOrganizer && !isAdmin) {
    const activeFromStr = new Date(event.activeFrom as string).toLocaleString(locale === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    return (
      <div className="page-container">
        <div className="card" style={{ position: 'relative', overflow: 'hidden' }}>
          <div style={{
            height: 200,
            background: event.imageUrl ? `url(${cachedImage}) center/cover no-repeat` : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
            borderRadius: '16px 16px 0 0',
          }} />
          <div style={{ padding: 32, textAlign: 'center' }}>
            <Icon iconName="Clock" style={{ fontSize: 48, color: 'var(--dex-orange)', marginBottom: 16 }} />
            <h2 style={{ marginBottom: 8 }}>{event.title}</h2>
            <p style={{ color: 'var(--dex-gray-700)', marginBottom: 8, fontWeight: 600 }}>
              {locale === 'de' ? 'Die Anmeldung ist noch nicht geöffnet.' : 'Registration is not open yet.'}
            </p>
            <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.9rem' }}>
              {locale === 'de' ? 'Anmeldung ab' : 'Registration opens'}: <strong>{activeFromStr}</strong>
            </p>
            <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
              {t('reg.backtoevents')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (isFullyClosed && !isOrganizer && !isAdmin) {
    return (
      <div className="page-container">
        <div className="card" style={{ position: 'relative', overflow: 'hidden' }}>
          <div style={{
            height: 200,
            background: event.imageUrl
              ? `url(${cachedImage}) center/cover no-repeat`
              : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
            borderRadius: '16px 16px 0 0',
          }} />
          <div style={{ padding: 32, textAlign: 'center' }}>
            <Icon iconName="Clock" style={{ fontSize: 48, color: 'var(--dex-orange)', marginBottom: 16 }} />
            <h2 style={{ marginBottom: 8 }}>{t('reg.deadlinepassed.title')}</h2>
            <p style={{ color: 'var(--dex-gray-600)', marginBottom: 8 }}>
              {t('reg.deadlinepassed.text')}
            </p>
            {/* v28.20: Bei Klammern mit expliziter Frist DIE anzeigen — die
                Spalten-Frist ist dort ein wirkungsloser Alt-Wert (und kann
                leer sein → Invalid Date). */}
            {(() => {
              const d = event.klammerDeadline || event.registrationDeadline;
              return d ? (
                <p style={{ color: 'var(--dex-gray-400)', fontSize: '0.85rem' }}>
                  {t('reg.deadlinepassed.date')}: {new Date(d).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}
                </p>
              ) : null;
            })()}
            <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
              {t('reg.backtoevents')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  const errorBorder = { border: '2px solid var(--dex-red)' };

  const parentAlreadyRegistered = !!(myParentReg && myParentReg.Status && myParentReg.Status !== 'Abgemeldet');
  // "Hauptevent wird jetzt angemeldet" gilt nur, wenn der Parent-Checkbox an ist
  // UND der User nicht bereits angemeldet ist. Bei bereits angemeldetem Parent
  // wird die Parent-Registrierung nicht nochmal ausgelöst.
  // v22.54: Ist die Hauptevent-Frist abgelaufen, kann ein normaler Teilnehmer
  // das Hauptevent nicht mehr buchen — die offenen Sub-Events bleiben aber
  // wählbar. Organizer/Admins dürfen weiterhin (manuelle Anmeldung).
  // v24.41: „Meine Assistenz beauftragen"-Prompt — nur bei der EIGENEN
  // Anmeldung (kein Stellvertreter-/Team-Modus), nur für Admins ODER Directoren,
  // und NICHT, wenn das Event bereits ein Assistenz-CC-Feld hat (dann hätte der
  // Organizer den CC schon eingebaut → kein doppeltes Abfragen).
  // v24.45: Partner ODER Director (P/D) — vorher wurde nur „director" geprüft,
  // Partner gingen leer aus. „Senior Director"/„Associate Partner" matchen mit.
  const isPartnerOrDirector = /(partner|director)/i.test(currentUser.jobTitle || '');
  // v24.48: Unterdrücken, wenn das Event bereits ein eigenes Assistenz-Feld hat
  // — entweder ein People-Picker mit „auf CC setzen" ODER ein People-Picker,
  // dessen Bezeichnung auf „Assistenz/Assistant" hindeutet (z.B. „Your assistant").
  const hasAssistantCcField = (event?.eventSpecificFields || []).some(f => {
    if (f.type !== 'user' && f.type !== 'roommate') return false;
    if (f.ccOnEmails) return true;
    return /assist|assistenz|assistenten?/i.test(`${f.label || ''} ${f.labelEn || ''}`);
  });
  // v27.8: NUR noch Partner/Director — die Abfrage „Für Partner & Directoren"
  // erschien vorher auch jedem Admin bei der eigenen Anmeldung, obwohl ein
  // Admin nicht zwangsläufig P/D ist. Admins, die zugleich P/D sind, matchen
  // weiter über isPartnerOrDirector.
  const canDelegateAssistant = isPartnerOrDirector && !registerForOther && !isTeamMode && !pendingJoinTeam && !hasAssistantCcField;
  const parsedDelegateAssist = (() => {
    const m = (delegateAssistValue || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
    return m ? { name: m[1].trim(), email: m[2].trim() } : null;
  })();
  // v27.11: Voll & Warteliste vom Organizer deaktiviert → Hauptevent nicht
  // mehr buchbar. Vorher lief die Anmeldung still auf die (abgeschaltete)
  // Warteliste — der WaitlistEnabled-Toggle war wirkungslos. Sub-Events
  // bleiben über die parentRegBlocked-Mechanik weiterhin einzeln buchbar.
  const pfActive = liveStats ? liveStats.active : (event ? (event.currentParticipants || 0) : 0);
  const pfWaitlist = (liveStats && liveStats.waitlist >= 0) ? liveStats.waitlist : (event ? (event.waitlistCount || 0) : 0);
  const parentFullNoWaitlist = !!event && event.maxParticipants > 0 && event.waitlistEnabled === false
    && Math.max(0, event.maxParticipants - pfActive - pfWaitlist) <= 0;
  const parentRegBlocked = ((isDeadlinePassed && !isOrganizer && !isAdmin) || parentFullNoWaitlist) && !parentAlreadyRegistered;
  const willRegisterParent = registerForParent && !parentAlreadyRegistered && !registerForOther && !(event && event.subEventsOnlyMode) && !parentRegBlocked;
  // Fürs Registrieren für andere bleibt der alte Flow: Parent wird immer registriert,
  // keine Session-Auswahl (siehe Render).
  const isSessionsOnlyMode = !willRegisterParent && !registerForOther && !parentAlreadyRegistered;

  // v28.88: Gibt es überhaupt etwas abzuschicken?
  //
  // Bisher hing die Sperre allein an `selectedSessions.size === 0`. Das ist zu
  // grob: Wer bereits gebuchte Sub-Events ALLE abwählt, will sie abmelden —
  // die Auswahl ist dann leer, es gibt aber sehr wohl etwas zu tun (der
  // Abmelde-Pfad im Sub-Event-Loop weiter unten). Deshalb zählt hier jede
  // Abweichung zwischen Vorbelegung (sessionMeta.wasRegistered) und aktueller
  // Auswahl als Änderung.
  const sessionsChanged = childEvents.some(ce => {
    const wasReg = !!sessionMeta[ce.id]?.wasRegistered;
    const isSel = selectedSessions.has(ce.id);
    return (isSel && !wasReg) || (!isSel && wasReg && !registerForOther);
  });
  // Nichts anzumelden, nichts zu ändern, kein Team-Vorgang → der
  // „Registrieren"-Klick hätte keine Wirkung.
  const nothingToSubmit = !willRegisterParent && !registerForOther && !isTeamMode
    && !pendingJoinTeam && selectedSessions.size === 0 && !sessionsChanged;

  // v29.27: Sub-Event-Fragen INLINE in der Sub-Event-Karte — nicht mehr im
  // Bestätigen-Modal. Der Teilnehmer sieht damit direkt an der Kachel, welche
  // Frage zu welchem Termin gehört (die Hauptevent-Fragen stehen darunter mit
  // eigener Überschrift). Die Werte hängen live an sessionFieldValues[ce.id];
  // die Pflicht-Prüfung, die vorher das Modal erzwang, sitzt jetzt im Submit.
  // Kalender-Modus und der Team-Beitritts-Dialog nutzen weiter das Modal —
  // dort gibt es keine Karte, die die Felder tragen könnte.
  const renderSubEventInlineFields = (ce: DeloitteEvent): React.ReactElement | null => {
    const values = sessionFieldValues[ce.id] || {};
    const useEnHere = locale === 'en' && !!ce.bilingualFields;
    const fLabel = (f: EventSpecificField): string =>
      (useEnHere && f.labelEn && f.labelEn.trim()) ? f.labelEn : f.label;
    const fHelp = (f: EventSpecificField): string | undefined =>
      (useEnHere && f.helpTextEn && f.helpTextEn.trim()) ? f.helpTextEn : f.helpText;
    const fOpt = (f: EventSpecificField, opt: string, idx: number): string =>
      (useEnHere && f.optionsEn && f.optionsEn[idx] && f.optionsEn[idx].trim()) ? f.optionsEn[idx] : opt;
    const fields = (ce.eventSpecificFields || [])
      .filter(f => f && f.label)
      .filter(f => {
        if (!f.showIf || !f.showIf.fieldId) return true;
        const raw = (values[f.showIf.fieldId] || '').trim();
        if (!raw) return false;
        const answers = raw.indexOf(' | ') >= 0
          ? raw.split(' | ').map(s => s.trim()).filter(Boolean)
          : [raw];
        return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
      });
    if (fields.length === 0) return null;
    const setValue = (fieldId: string, value: string): void => {
      setSessionFieldValues(prev => ({ ...prev, [ce.id]: { ...(prev[ce.id] || {}), [fieldId]: value } }));
    };
    return (
      <div style={{ marginTop: 10, padding: '10px 12px', borderRadius: 8, background: '#fff', border: '1px solid var(--dex-gray-200)' }}>
        <div style={{ fontSize: '0.78rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', marginBottom: 8 }}>
          {locale === 'de'
            ? `Fragen zu diesem ${childTermSingular || 'Sub-Event'}`
            : `Questions for this ${childTermSingular || 'sub-event'}`}
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          {fields.map(f => {
            const val = values[f.id] || '';
            const missing = showErrors && f.required && (f.type === 'checkbox' ? val !== 'true' : !val.trim());
            return (
              <div key={f.id}>
                <label className="form-label" style={{ display: 'block', fontSize: '0.82rem', marginBottom: 4, ...(missing ? { color: 'var(--dex-red, #c00)' } : {}) }}>
                  {fLabel(f)}
                  {f.required && <span style={{ color: 'var(--dex-red, #c00)', marginLeft: 4 }}>*</span>}
                  {fHelp(f) && <InfoTooltip text={fHelp(f)} />}
                </label>
                {f.type === 'select' && f.multi ? (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    {(f.options || []).map((opt, optIdx) => {
                      const current = val.split(' | ').map(s => s.trim()).filter(Boolean);
                      const checked = current.indexOf(opt) >= 0;
                      return (
                        <label key={opt} style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.84rem', cursor: 'pointer' }}>
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={e => {
                              const next = e.target.checked ? [...current, opt] : current.filter(x => x !== opt);
                              setValue(f.id, next.join(' | '));
                            }}
                          />
                          {fOpt(f, opt, optIdx)}
                        </label>
                      );
                    })}
                  </div>
                ) : f.type === 'select' ? (
                  <select className="form-input" value={val} onChange={e => setValue(f.id, e.target.value)} style={{ width: '100%', fontSize: '0.88rem' }}>
                    <option value="">{locale === 'de' ? '— bitte wählen —' : '— please select —'}</option>
                    {(f.options || []).map((opt, optIdx) => <option key={opt} value={opt}>{fOpt(f, opt, optIdx)}</option>)}
                  </select>
                ) : f.type === 'checkbox' ? (
                  <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.84rem' }}>
                    <input type="checkbox" checked={val === 'true'} onChange={e => setValue(f.id, e.target.checked ? 'true' : 'false')} />
                    {locale === 'de' ? 'Ja' : 'Yes'}
                  </label>
                ) : f.type === 'number' ? (
                  <input type="number" className="form-input" value={val} onChange={e => setValue(f.id, e.target.value)} style={{ width: '100%', fontSize: '0.88rem' }} />
                ) : (
                  <input type="text" className="form-input" value={val} onChange={e => setValue(f.id, e.target.value)} style={{ width: '100%', fontSize: '0.88rem' }} />
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  // v29.28: Der Hauptevent-Felder-Block als Render-Funktion — er wird an
  // ZWEI möglichen Orten gebraucht: bei Events MIT Sub-Events direkt unter
  // der Haupt-Event-Kachel in der Auswahl-Box (dort, wo die Fragen
  // hingehören — die Sub-Event-Fragen stecken seit v29.27 in deren Karten),
  // sonst an der bisherigen Stelle unter der Auswahl. Inhalt 1:1 der
  // bisherige Block (v11.2/v26.91) — als Ganzes gehoben, nicht geschnitten.
  const renderMainFieldsSection = (): React.ReactElement => (
    event.eventSpecificFields.length === 0 && !isSplitGroup ? (
      <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{t('reg.noadditional')}</p>
    ) : (
      // v11.2 / v11.5: Custom-Fields ohne Pro-Gruppe-Constraint im
      // 2-Spalten-Grid. Group-spezifische Felder werden bereits
      // oben innerhalb der Gruppen-Auswahl-Box gerendert und hier
      // ausgefiltert.
      <div className="dex-reg-fields-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
      {/* v29.27: Zuordnung klarmachen — die Sub-Event-Fragen stehen
          in den Karten, diese hier gehören zum Haupt-Event (bzw. bei
          einer Klammer zur Anmeldung insgesamt). */}
      {childEvents.length > 0 && event.eventSpecificFields.length > 0 && (
        <div style={{ gridColumn: '1 / -1', fontSize: '0.8rem', fontWeight: 700, color: 'var(--dex-gray-600)', marginBottom: -6 }}>
          {event.subEventsOnlyMode
            ? (locale === 'de' ? 'Allgemeine Fragen zur Anmeldung' : 'General questions for your registration')
            : (locale === 'de' ? 'Fragen zum Haupt-Event' : 'Questions for the main event')}
        </div>
      )}
      {(() => {
        // v26.91: Zuerst die WIRKLICH sichtbaren Felder ermitteln, dann mit
        // Index rendern — so kann renderRegField pro 2-Spalten-Zeile
        // entscheiden, ob es leeren Beschreibungs-Platz reservieren muss.
        const visibleSpecificFields = event.eventSpecificFields
          .filter(f => f.id !== 'b2run_mobilnummer' || eventSpecific['b2run_infoservice'] === 'true')
          .filter(f => !(f.id === 'b2run_startblock' && hasStarterBlockMapping))
          .filter(f => {
            if (!f.showIf || !f.showIf.fieldId) return true;
            const raw = (eventSpecific[f.showIf.fieldId] || '').trim();
            if (!raw) return false;
            const answers = raw.indexOf(' | ') >= 0
              ? raw.split(' | ').map(s => s.trim()).filter(Boolean)
              : [raw];
            return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
          })
          // v11.5: Group-Spec NICHT hier rendern — die kommen oben
          // in der Gruppen-Auswahl-Box. Hier nur 'all' / undefined
          // (oder Events ohne Split-Capacity).
          .filter(f => {
            const grp = f.onlyForGroup;
            if (!grp || grp === 'all') return true;
            if (!isSplitGroup) return true;
            return false;
          });
        return visibleSpecificFields.map((f, i) => renderRegField(f, undefined, undefined, i, visibleSpecificFields));
      })()}
      </div>
    )
  );

  // v30.66: Der Submit-Pfad (handleSubmit + Team-/Beitritts-/Registrierungs-Weg,
  // zusammen ueber 1.100 Zeilen) liegt in `./registration/submitFlow`. Die Fabrik
  // bekommt den Render-Stand als Kontext-Objekt und wird — wie die frueheren
  // Deklarationen an dieser Stelle — bei jedem Render neu gebaut. Der Aufruf muss
  // HIER stehen: darueber sind alle uebergebenen Werte deklariert, darunter
  // erwarten Aufrufer bereits `handleSubmit` und `performRegistration`.
  const { handleSubmit, performRegistration } = createSubmitFlow({
    assistantModalDecidedRef, cancelRegistration, canCreateEvents, canDelegateAssistant, ccSelfDecidedRef, ccSelfRef,
    childEvents, childOneDe, childTermPlural, childTermSingular, confirmDialog, confirmDialogConfirmedRef,
    createTeamJoinRequest, currentUser, delegateAssistEnabled, delegateAssistValue, delegateChoiceRef, delegateRegistrationToAssistant,
    durchCap, email, event, eventSpecific, externalEmailConfirmedRef, externalPerson,
    firstName, funCap, getEventNumbersForEmail, hiddenChildCount, isAllowedTargetForAssistant, isAssistant,
    isDeadlinePassed, isSplitGroup, isTeamMode, joinTeam, locale, myParentReg,
    nothingToSubmit, otherConsentConfirmed, parentAlreadyRegistered, parentFullNoWaitlist, parentRegBlocked, pendingDocFiles,
    pendingJoinTeam, preferredStarterType, previewAsUser, recordProxyDelegation, refreshEvents, registerForEvent,
    registerForOther, registerForParent, registerTeam, salutation, searchUsers, selectedEventId,
    selectedSessions, sendBundledUpdateMail, sessionFieldValues, sessionMeta, setAssistantModalOpen, setCcSelfModalOpen,
    setConfirmDialogAck, setConfirmDialogOpen, setConfirmDraftParent, setConfirmDraftSessions, setError, setExternalEmailWarning,
    setFallbackDialog, setIsSubmitting, setSessionsOnlySubmitted, setShowErrors, setSubmitProgress, setSubmitProgressLabel,
    setSubmitted, setSubmittedAsWaitlist, setSubmittedJoinKind, showAlert, starterCounts, submittedSessionsRef,
    subOnlyTerms, surname, t, teamMemberFields, teamMembersParsed, teamName,
    teamValidation, thirdPartyCheck, updateMyRegistration, uploadFieldDocument, userResults, userSearchIncludeIntl,
    willRegisterParent,
  });

  // v26.37: „Zurücksetzen"-Button (und handleClear) entfernt — auf Wunsch.

  // v18.11: Proaktive Absage („Ich nehme nicht teil"). Keine Pflichtfelder
  // nötig — der User signalisiert nur, dass er nicht kommt.
  const handleDecline = async (): Promise<void> => {
    if (!event || isDeclining) return;
    // v30.4: Auch die Absage ist ein Schreibvorgang — in der Vorschau gesperrt.
    if (previewAsUser) {
      setError(locale === 'de'
        ? 'Vorschau-Modus: nur zum Ansehen — auch Absagen ist deaktiviert. Beende die Vorschau über den blauen Balken oben.'
        : 'Preview mode: view only — declining is disabled as well. End the preview via the blue bar at the top.');
      return;
    }
    if (event.isDemoShowcase) {
      setError(locale === 'de'
        ? 'Dies ist ein Demo-Event — es wird nichts gespeichert.'
        : 'This is a demo event — nothing is stored.');
      return;
    }
    setIsDeclining(true);
    setError('');
    try {
      // v29.32: Die Absage gilt für das GANZE Event — Klammer/Haupt-Event UND
      // alle sichtbaren Sub-Events. Vorher landete sie nur in der
      // Hauptevent-Liste: Im „Nur Sub-Events"-Modus ist das eine Schattenzeile,
      // und in den Sub-Event-Listen (die der Organizer tatsächlich auswertet)
      // stand die Person weiter als „hat nicht geantwortet". Eine Auswahl ist
      // dafür bewusst NICHT nötig — wer absagt, sagt für alles ab.
      // declineEvent je Ziel macht das Richtige: bestehende Anmeldung →
      // regulärer Abmelde-Pfad (Platz frei, Mail, Nachrücken), sonst eine
      // Absage-Zeile.
      const ok = await declineEvent(event.id);
      let subFailed = 0;
      for (const ce of childEvents) {
        try { if (!(await declineEvent(ce.id))) subFailed++; }
        catch { subFailed++; }
      }
      if (ok && subFailed === 0) setDeclined(true);
      else if (ok) {
        // Klammer steht, einzelne Sub-Events nicht — den Teilablauf benennen,
        // statt eine vollständige Absage zu behaupten.
        setDeclined(true);
        setError(locale === 'de'
          ? `Deine Absage ist erfasst — bei ${subFailed} ${subFailed === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} hat es nicht geklappt. Bitte melde dich bei den Organizern.`
          : `Your decline was recorded — it failed for ${subFailed} sub-event(s). Please contact the organizers.`);
      }
      else setError(t('reg.genericerror') || 'Ein Fehler ist aufgetreten. Bitte versuche es erneut.');
    } catch {
      setError(t('reg.genericerror') || 'Ein Fehler ist aufgetreten. Bitte versuche es erneut.');
    } finally {
      setIsDeclining(false);
    }
  };

  // v18.14: Schritt 1 — eingefügte Liste gegen das Verzeichnis auflösen und
  // eine Vorschau-Tabelle (Vorname/Nachname/Position/Standort/E-Mail) bauen.
  // Pro Zeile: E-Mail erkennen (dann Profil-Lookup per E-Mail) ODER nur ein
  // Name (dann Personensuche → bester Treffer). Duplikate + nicht auflösbare
  // Zeilen werden markiert.
  const splitName = (raw: string): { firstName: string; lastName: string } => {
    const dn = (raw || '').trim();
    if (!dn) return { firstName: '', lastName: '' };
    if (dn.indexOf(',') >= 0) { const p = dn.split(',').map(s => s.trim()); return { firstName: p[1] || '', lastName: p[0] || '' }; }
    const p = dn.split(/\s+/).filter(Boolean);
    if (p.length <= 1) return { firstName: p[0] || '', lastName: '' };
    return { firstName: p[0], lastName: p.slice(1).join(' ') };
  };
  const resolveMassImport = async (): Promise<void> => {
    if (massImportResolving) return;
    const lines = massImportText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    const EMAIL_RE = /([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/;
    setMassImportResolving(true);
    const rows: typeof massImportRows = [];
    const seen = new Set<string>();
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      setMassImportProgress(`${i + 1} / ${lines.length}`);
      const m = line.match(EMAIL_RE);
      let email = m ? m[1].toLowerCase() : '';
      const nameRaw = line.replace(m ? m[1] : '', '').replace(/[<>;,\t]/g, ' ').trim();
      let jobTitle = ''; let location = ''; let displayName = nameRaw;
      if (email) {
        // Profil per E-Mail nachschlagen (für Position + Standort + Name).
        try { const p = await searchUser(email); if (p) { displayName = p.displayName || nameRaw; jobTitle = p.jobTitle || ''; location = p.location || ''; } } catch { /* */ }
      } else if (nameRaw) {
        // Kein E-Mail in der Zeile → Personensuche, besten Treffer nehmen.
        try {
          const results = await searchUsers(nameRaw, false);
          if (results && results.length > 0) {
            const best = results[0];
            email = (best.email || '').toLowerCase();
            displayName = best.displayName || nameRaw;
            jobTitle = best.jobTitle || '';
            location = best.location || '';
          }
        } catch { /* */ }
      }
      const { firstName, lastName } = splitName(displayName);
      let status: 'ok' | 'duplicate' | 'notfound';
      if (!email) status = 'notfound';
      else if (seen.has(email)) status = 'duplicate';
      else { seen.add(email); status = 'ok'; }
      rows.push({ email, firstName, lastName, jobTitle, location, status, raw: line });
    }
    setMassImportRows(rows);
    setMassImportStep('preview');
    setMassImportResolving(false);
    setMassImportProgress('');
  };

  // v18.14: Schritt 2 — die aufgelösten, gültigen Zeilen anmelden.
  const runMassImport = async (): Promise<void> => {
    if (!event || massImportBusy) return;
    const toRegister = massImportRows.filter(r => r.status === 'ok');
    if (toRegister.length === 0) return;
    const suppressMail = massImportMode === 'nomail' || massImportMode === 'silent';
    const suppressOutlook = massImportMode === 'silent';
    setMassImportBusy(true);
    setMassImportResult(null);
    let ok = 0;
    const failed: string[] = [];
    for (let i = 0; i < toRegister.length; i++) {
      const r = toRegister[i];
      setMassImportProgress(`${i + 1} / ${toRegister.length} — ${r.email}`);
      try {
        // v30.9: skipReload — sonst zieht JEDE Zeile des Massenimports einen
        // kompletten loadEvents nach sich (N × 28 MB → Drosselung).
        const success = (await registerForEvent(event.id, {}, r.firstName, r.lastName, r.email, undefined, { suppressMail, suppressOutlook, skipReload: true })).ok;
        if (success) ok++; else failed.push(r.email);
      } catch { failed.push(r.email); }
    }
    // Ein Sammel-Refresh statt N Einzel-Reloads; nicht awaiten — das Ergebnis-
    // Panel soll sofort erscheinen, die Zähler ziehen im Hintergrund nach.
    if (ok > 0) void refreshEvents().catch(() => { /* best-effort */ });
    setMassImportBusy(false);
    setMassImportProgress('');
    setMassImportResult({ ok, failed });
  };

  if (declined) {
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 640, margin: '0 auto' }}>
          <h2 style={{ marginTop: 0 }}>
            {locale === 'de' ? 'Absage erfasst' : 'Decline recorded'}
          </h2>
          <p className="mt-8" style={{ color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
            {locale === 'de'
              ? <>Danke für die Rückmeldung — wir haben vermerkt, dass du <strong>nicht</strong> an &bdquo;{event?.title}&ldquo; teilnimmst. Falls sich das ändert, kannst du dich jederzeit über diese Seite anmelden.</>
              : <>Thanks for letting us know — we noted that you will <strong>not</strong> attend &bdquo;{event?.title}&ldquo;. If that changes, you can register any time via this page.</>}
          </p>
          <div style={{ marginTop: 28 }}>
            <button className="btn btn-primary" onClick={() => navigate('register')}>
              {t('reg.backtoevents') || (locale === 'de' ? 'Zurück zu Events' : 'Back to events')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  // v18.73: Eigener Erfolgsscreen für den Team-Beitritt (direkt angemeldet
  // bzw. Beitritts-Anfrage gesendet) — vor der generischen Anmelde-Logik.
  if (submitted && submittedJoinKind) {
    const isReq = submittedJoinKind === 'requested';
    const headline = isReq
      ? (locale === 'de' ? 'Beitritts-Anfrage gesendet' : 'Join request sent')
      : (submittedAsWaitlist
          ? (locale === 'de' ? 'Auf der Warteliste' : 'On the waitlist')
          : (locale === 'de' ? 'Team-Beitritt erfolgreich' : 'Joined the team'));
    const body = isReq
      ? (locale === 'de'
          ? `Deine Anfrage zum Beitritt wurde an den Team-Kapitän gesendet. Sobald er entscheidet, bekommst du eine E-Mail mit dem Ergebnis. Deine Angaben werden bei der Bestätigung automatisch übernommen.`
          : `Your join request has been sent to the team lead. Once they decide, you will receive an email with the result. Your details will be applied automatically upon approval.`)
      : (submittedAsWaitlist
          ? (locale === 'de'
              ? `Das Team war voll — du stehst jetzt auf der Warteliste für „${event.title}". Sobald ein Platz frei wird, rückst du automatisch nach und bekommst eine Bestätigung. Details findest du unter „Meine Events".`
              : `The team was full — you are now on the waitlist for „${event.title}". You will be moved up automatically when a spot opens and receive a confirmation. See „My Events" for details.`)
          : (locale === 'de'
              ? `Du bist dem Team beigetreten und für „${event.title}" angemeldet. Du bekommst eine Bestätigungs-E-Mail und einen Outlook-Termin. Details findest du unter „Meine Events".`
              : `You joined the team and are registered for „${event.title}". You will receive a confirmation email and an Outlook invite. See „My Events" for details.`));
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 720, margin: '0 auto' }}>
          {event.imageUrl && (
            <div style={{
              width: '100%', maxWidth: 480, height: 200, margin: '0 auto 24px',
              borderRadius: 'var(--dex-radius-lg)',
              background: `url(${cachedImage}) center/cover no-repeat`,
            }} />
          )}
          <h2 style={{ marginTop: 0 }}>{headline}</h2>
          <p style={{ fontSize: '0.95rem', color: 'var(--dex-gray-700)', lineHeight: 1.6, maxWidth: 560, margin: '0 auto' }}>{body}</p>
        </div>
      </div>
    );
  }

  if (submitted) {
    const sessionsOnlyHint = sessionsOnlySubmitted;
    // v26.71: Externe stellvertretende Anmeldung — die Person ist NICHT final
    // angemeldet, sondern nur in der Teilnehmerliste hinterlegt (Datenschutz-
    // rückmeldung offen). Es geht KEINE Mail an die externe Adresse und KEIN
    // Kalendereintrag; der Anmelder verschickt die Einladung selbst und
    // bestätigt danach. Deshalb hier NICHT „erfolgreich registriert" texten.
    const isExternalProxy = registerForOther && isExternalEmailAddr((email || '').trim());
    const proxyName = `${firstName} ${surname}`.trim() || email;
    const successHeadline = isExternalProxy
      ? (locale === 'de' ? 'In der Teilnehmerliste hinterlegt' : 'Added to the participant list')
      : sessionsOnlyHint
      ? (childTermPlural
          ? (locale === 'de' ? `Für ${childTermPlural} angemeldet` : `Registered for ${childTermPlural}`)
          : (t('reg.success.sessionsonly.title') || 'Für Sessions angemeldet'))
      : (submittedAsWaitlist ? t('reg.waitlisttitle') : t('reg.success'));
    const successBody = isExternalProxy
      ? (locale === 'de'
          ? `${proxyName} wurde in der Teilnehmerliste hinterlegt — mit dem Status „Angemeldet (Datenschutzrückmeldung offen)". Da „${email}" eine externe Adresse ist, versendet die App KEINE Mail dorthin und KEINEN Kalendereintrag. Du bekommst eine E-Mail (Organisator:innen in Kopie) mit einem Button, über den du den fertigen Einladungs-Entwurf direkt herunterlädst — leite ihn aus deinem eigenen Postfach an ${proxyName} weiter und bestätige nach ihrer Rückmeldung die Datenschutz-Zustimmung in der Teilnehmerliste.`
          : `${proxyName} has been added to the participant list — with status „Registered (privacy confirmation pending)". Since „${email}" is an external address, the app sends NO email there and NO calendar entry. You'll receive an email (organizers in copy) with a button to download the ready-made invitation draft directly — forward it from your own mailbox to ${proxyName}, and confirm the privacy consent in the participant list once they reply.`)
      : sessionsOnlyHint
      ? (event.subEventsOnlyMode
          ? (childTermSingular
              ? (locale === 'de'
                  ? `Du hast dich für die ausgewählten ${childTermPlural || `${childTermSingular}s`} im Rahmen von „${event.title}" angemeldet. Du bekommst pro ${childTermSingular} eine separate Bestätigungsmail und einen eigenen Outlook-Kalendereintrag.`
                  : `You registered for the selected items within "${event.title}". You will receive a separate confirmation email and Outlook calendar entry per ${childTermSingular}.`)
              : (locale === 'de'
                  ? `Du hast dich für die ausgewählten Sub-Events im Rahmen von „${event.title}" angemeldet. Du bekommst pro Sub-Event eine separate Bestätigungsmail und einen eigenen Outlook-Kalendereintrag.`
                  : `You registered for the selected sub-events within "${event.title}". You will receive a separate confirmation email and Outlook calendar entry per sub-event.`))
          : (childTermPlural && childTermSingular
              ? (locale === 'de'
                  ? `Du hast dich ausschließlich für die ausgewählten ${childTermPlural} angemeldet — NICHT für das Haupt-Event „${event.title}". Du bekommst pro ${childTermSingular} eine separate Bestätigungsmail und einen eigenen Outlook-Kalendereintrag.`
                  : `You registered exclusively for the selected ${childTermPlural} — NOT for the main event "${event.title}". You will receive a separate confirmation email and Outlook calendar entry per ${childTermSingular}.`)
              : (t('reg.success.sessionsonly.msg') || 'Du hast dich ausschließlich für die ausgewählten Sessions angemeldet — NICHT für das Haupt-Event "{title}". Du bekommst pro Session eine separate Bestätigungsmail und einen eigenen Outlook-Kalendereintrag.').replace('{title}', event.title)))
      : (submittedAsWaitlist
          ? (registerForOther
              ? t('reg.waitlistmsg.other').replace('{name}', `${firstName} ${surname}`.trim()).replace('{title}', event.title).replace('{email}', email)
              : t('reg.waitlistmsg').replace('{title}', event.title))
          : (() => {
              // v28.16: Statt der nackten E-Mail-Adresse konkret sagen, was
              // automatisch verschickt wird — Mail-Bestätigung und Outlook-
              // Termin jeweils nur, wenn sie für das Event aktiv sind.
              const mailActive = !event.disableEmails && !event.disableRegistrationEmail;
              const outlookActive = !event.disableOutlook && !!event.startDate;
              let confirmTail = '';
              if (registerForOther) {
                const who = `${firstName} ${surname}`.trim() || email;
                if (mailActive && outlookActive) confirmTail = locale === 'de'
                  ? ` ${who} erhält automatisch eine Bestätigung per E-Mail (an ${email}) sowie einen Outlook-Kalendereintrag.`
                  : ` ${who} will automatically receive a confirmation email (to ${email}) and an Outlook calendar invitation.`;
                else if (mailActive) confirmTail = locale === 'de'
                  ? ` ${who} erhält automatisch eine Bestätigung per E-Mail an ${email}.`
                  : ` ${who} will automatically receive a confirmation email to ${email}.`;
                else if (outlookActive) confirmTail = locale === 'de'
                  ? ` ${who} erhält automatisch einen Outlook-Kalendereintrag.`
                  : ` ${who} will automatically receive an Outlook calendar invitation.`;
              } else {
                if (mailActive && outlookActive) confirmTail = locale === 'de'
                  ? ' Du erhältst automatisch eine Bestätigung per E-Mail sowie einen Outlook-Kalendereintrag.'
                  : ' You will automatically receive a confirmation email and an Outlook calendar invitation.';
                else if (mailActive) confirmTail = locale === 'de'
                  ? ' Du erhältst automatisch eine Bestätigung per E-Mail.'
                  : ' You will automatically receive a confirmation email.';
                else if (outlookActive) confirmTail = locale === 'de'
                  ? ' Du erhältst automatisch einen Outlook-Kalendereintrag.'
                  : ' You will automatically receive an Outlook calendar invitation.';
              }
              const base = registerForOther
                ? t('reg.successmsg.other').replace('{name}', `${firstName} ${surname}`.trim()).replace('{title}', event.title)
                : t('reg.successmsg').replace('{title}', event.title);
              return base + confirmTail;
            })());
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 720, margin: '0 auto' }}>
          {/* v15.20: Event-Foto + Organizer-Info auch auf der
              Erfolgs-Seite anzeigen — analog zur Event-Karte. */}
          {event.imageUrl && (
            <div style={{
              width: '100%', maxWidth: 480, height: 200, margin: '0 auto 24px',
              borderRadius: 'var(--dex-radius-lg)',
              background: `url(${cachedImage}) center/cover no-repeat`,
            }} />
          )}
          <h2 style={{ marginTop: 0 }}>{successHeadline}</h2>
          {/* v18.54: Im subEventsOnlyMode strukturierter Bestätigungstext —
              Begrüßung, Verweis auf das (nicht anwählbare) Hauptevent, Bullet-
              Liste der gewählten Sections (dynamische Organizer-Bezeichnung) und
              der Mail/Outlook-Satz NUR wenn für die gewählten Sections wirklich
              Mail bzw. Outlook aktiv ist. */}
          {!isExternalProxy && sessionsOnlyHint && event.subEventsOnlyMode ? (() => {
            const selectedChildren = childEvents.filter(ce => selectedSessions.has(ce.id) || submittedSessionsRef.current.has(ce.id));
            const anyEmail = selectedChildren.some(ce => !ce.disableEmails);
            const anyOutlook = selectedChildren.some(ce => !ce.disableOutlook);
            const sectionPlural = childTermPlural || (locale === 'de' ? 'Event-Sections' : 'event-sections');
            const sectionSingular = childTermSingular || (locale === 'de' ? 'Event-Section' : 'event-section');
            const greetingName = (firstName || '').trim();
            let confirmLine = '';
            if (anyEmail && anyOutlook) confirmLine = locale === 'de'
              ? `Du erhältst pro ${sectionSingular} eine E-Mail-Bestätigung und einen Outlook-Termin.`
              : `You will receive a confirmation email and an Outlook invitation per ${sectionSingular}.`;
            else if (anyEmail) confirmLine = locale === 'de'
              ? `Du erhältst pro ${sectionSingular} eine E-Mail-Bestätigung.`
              : `You will receive a confirmation email per ${sectionSingular}.`;
            else if (anyOutlook) confirmLine = locale === 'de'
              ? `Du erhältst pro ${sectionSingular} einen Outlook-Termin.`
              : `You will receive an Outlook invitation per ${sectionSingular}.`;
            return (
              <div className="mt-8" style={{ color: 'var(--dex-gray-700)', textAlign: 'left', maxWidth: 520, margin: '8px auto 0', lineHeight: 1.6 }}>
                <p style={{ margin: '0 0 10px' }}>
                  {locale === 'de'
                    ? <>Hallo{greetingName ? <> <strong>{greetingName}</strong></> : ''},</>
                    : <>Hi{greetingName ? <> <strong>{greetingName}</strong></> : ''},</>}
                </p>
                <p style={{ margin: '0 0 10px' }}>
                  {locale === 'de'
                    ? <>du hast dich erfolgreich für das <strong>{event.title}</strong> angemeldet. Wir haben deine Anmeldung für die folgenden {sectionPlural} erhalten:</>
                    : <>you have successfully registered for <strong>{event.title}</strong>. We received your registration for the following {sectionPlural}:</>}
                </p>
                <ul style={{ margin: '0 0 10px', paddingLeft: 22 }}>
                  {selectedChildren.map(ce => {
                    // v19.33: nur den reinen Section-Namen zeigen (Parent-Präfix
                    // „<Hauptevent> | …" strippen) + dahinter „ | <Datum>".
                    const full = (ce.title || '').trim();
                    const pipe = full.lastIndexOf('|');
                    const name = (pipe >= 0 ? full.substring(pipe + 1).trim() : full) || (locale === 'de' ? 'ohne Titel' : 'untitled');
                    // v30.4: Kalender-Tage tragen das Datum bereits im Titel
                    // („Di. 01.09.2026") — „Titel | Datum" verdoppelte dieselbe
                    // Angabe, und die 00:00-Uhrzeit eines ganztägigen Termins
                    // sagt nichts. Nennt der Titel den Tag schon, hängen wir
                    // höchstens eine echte Uhrzeit an.
                    const d = ce.startDate ? new Date(ce.startDate) : null;
                    const valid = !!d && isFinite(d.getTime());
                    const dayStr = valid ? d!.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '';
                    const timeStr = valid ? d!.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }) : '';
                    const titleHasDate = !!dayStr && name.indexOf(dayStr) >= 0;
                    const label = titleHasDate
                      ? (timeStr && timeStr !== '00:00' ? `${name} | ${timeStr}` : name)
                      : (valid ? `${name} | ${formatDate(ce.startDate)}` : name);
                    return (
                      <li key={ce.id} style={{ marginBottom: 3 }}>{label}</li>
                    );
                  })}
                </ul>
                {confirmLine && <p style={{ margin: 0 }}>{confirmLine}</p>}
              </div>
            );
          })() : (
            <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>{successBody}</p>
          )}
          {(() => {
            // v24.15: „Organizer ausblenden" ohne Einzel-Modus = ALLE aus.
            if (event.hideOrganizer && !event.hideOrganizerIndividualOnly) return null;
            // Organizer als Chips mit Foto (gleicher Stil wie auf der
            // Anmelde-Seite). „Nachname, Vorname" → „Vorname Nachname".
            const orgs = event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => {
              const trimmed = o.trim();
              const parts = trimmed.split(',').map(s => s.trim());
              return parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
            }).filter(Boolean);
            if (orgs.length === 0) return null;
            return (
              <div style={{ marginTop: 24, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8 }}>
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, fontWeight: 600 }}>Organizer</div>
                <OrganizerList names={orgs} emails={event.organizerEmails} hiddenEmails={(event.hideOrganizer && event.hideOrganizerIndividualOnly) ? event.hiddenOrganizerEmails : []} forceIsDe={locale === 'de'} size="md" display={event.organizerDisplayLarge ? 'card' : 'chip'} nameFontSize="1.05rem" hideContactPrompt={!!(event.contactName || event.contactEmail || event.contactInfo)} />
              </div>
            );
          })()}
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button className="btn btn-primary" onClick={() => navigate('my-events')}>
              {t('myevents.title')}
            </button>
            <button className="btn btn-secondary" onClick={() => navigate('register')}>
              {t('reg.registeranother')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  // v26.76: gemeinsame Auswahl-Logik für die stellvertretende Anmeldung —
  // wird sowohl im Wizard als auch (Fallback) inline genutzt. Übernimmt Name/
  // E-Mail aus dem Suchtreffer, lädt das Profil nach und prüft Doppel-Anmeldung
  // + Gästekreis (thirdPartyCheck).
  const pickProxyUser = (u: { email: string; displayName: string; location?: string; jobTitle?: string }): void => {
    let uFirstName = '';
    let uSurname = '';
    if (u.displayName.includes(',')) {
      const parts = u.displayName.split(',').map(s => s.trim());
      uSurname = parts[0] || '';
      uFirstName = parts[1] || '';
    } else {
      const parts = u.displayName.split(' ');
      uFirstName = parts[0] || '';
      uSurname = parts.slice(1).join(' ') || '';
    }
    setFirstName(uFirstName);
    setSurname(uSurname);
    setEmail(u.email);
    setUserSearch(u.displayName);
    setUserResults([]);
    setPickedUserProfile({ jobTitle: u.jobTitle || '', location: u.location || '' });
    searchUser(u.email).then(p => {
      if (p) {
        setPickedUserProfile({
          jobTitle: p.jobTitle || u.jobTitle || '',
          department: p.department || '',
          location: p.location || u.location || '',
          mobilePhone: p.mobilePhone || '',
          company: p.company || '',
        });
      }
    }).catch(() => { /* silent */ });
    setThirdPartyCheck(null);
    if (event) {
      (async () => {
        const existing = await checkRegistrationByEmail(event.id, u.email).catch(() => null);
        const alreadyRegistered = !!existing && existing.Status !== 'Abgemeldet';
        const notInAudience = !isEventVisibleForUser(event, u.email, u.location || '', [], u.jobTitle || '');
        setThirdPartyCheck({
          alreadyRegistered,
          notInAudience,
          registeredName: (existing && (existing.ParticipantName || `${existing.Vorname || ''} ${existing.Nachname || ''}`.trim())) || u.displayName || '',
          registeredDate: (existing && existing.RegistrationDate) || '',
        });
      })();
    }
  };
  // v11.5: Custom-Field-Renderer extrahiert — wird zweimal verwendet:
  // einmal direkt in der Gruppen-Auswahl-Box für Felder mit
  // onlyForGroup-Constraint, einmal im Eventspez-2-Spalten-Grid für
  // alle anderen Felder. Die Filter-Logik dazu unten in den
  // groupSpecificFields- bzw. generalFields-Konstanten.
  // v13.2: fRaw jetzt typsicher als EventSpecificField (vorher any).
  const renderRegField = (fRaw: EventSpecificField, store?: Record<string, string>, setStore?: (next: Record<string, string>) => void, rowIndex?: number, rowList?: EventSpecificField[]): React.ReactElement => {
    // v18.12: optionaler Wert-Store — für die Custom-Fields pro Team-Mitglied.
    // Default = eventSpecific/setEventSpecific (Lead bzw. Solo-Anmeldung).
    const vals = store || eventSpecific;
    const setVals = setStore || setEventSpecific;
    // Dynamisch Required erzwingen: bei aktivem Infoservice ist die
    // Mobilnummer Pflicht.
    let field: EventSpecificField = fRaw;
    // Mobilnummer bei aktiviertem Infoservice dynamisch zur Pflicht
    if (fRaw.id === 'b2run_mobilnummer' && vals['b2run_infoservice'] === 'true') {
      field = { ...field, required: true };
    }
    // v26.96: Die frühere automatische Injektion von AGB-/Datenschutz-Links
    // für b2run_datenschutz wurde entfernt — der Organizer hinterlegt Links
    // jetzt selbst (z.B. direkt in der Feld-Beschreibung mit dem Editor).
    // Laufshirt/T-Shirt-Feld bei B2Run ist Pflicht (falls in alten Events
    // noch nicht so markiert)
    if ((fRaw.id === 'b2run_laufshirt' || /laufshirt/i.test(fRaw.label || '')) && !fRaw.required) {
      field = { ...field, required: true };
    }
    // v17.20: vor jedem Render die EN-Variante einziehen, sofern verfügbar.
    const displayLabel = pickFieldLabel(field);
    const displayHelp = pickFieldHelp(field);
    const displayConfirmLabel = pickFieldConfirmLabel(field);
    // v18.18: 'inline' = Erklär-Text unter dem Label (nicht fett), sonst
    // weiterhin "i"-Hover-Box neben dem Label.
    const isInlineHelp = field.helpTextStyle === 'inline';
    const inlineHelpEl = (displayHelp && isInlineHelp)
      // v18.77: Inline-Hilfe reserviert mind. 2 Zeilen Höhe (minHeight). Dadurch
      // stehen die Eingaben benachbarter Felder auf gleicher Höhe, wenn sich die
      // Beschreibungen um eine Zeile unterscheiden — OHNE die Nachbar-Eingabe
      // (wie zuvor mit flexGrow) bis ganz nach unten zu ziehen, was bei Feldern
      // mit Inhalt UNTER der Eingabe (People-Picker mit „international suchen") zu
      // großen Lücken führte.
      // v26.91: Beschreibung darf **fett** + Links enthalten (renderFieldDescHtml
      // escaped alles andere — der Organizer-Text ist sicherer Origin).
      ? <div style={{ fontSize: '0.78rem', fontWeight: 400, color: 'var(--dex-gray-500)', lineHeight: 1.45, marginTop: 2, marginBottom: 6, minHeight: '2.9em' }} dangerouslySetInnerHTML={{ __html: renderFieldDescHtml(displayHelp) }} />
      : null;
    // v26.16: Felder OHNE Inline-Beschreibung bekommen einen leeren Platzhalter
    // gleicher Höhe, SOBALD irgendein Feld im Formular eine Inline-Beschreibung
    // hat — damit stehen die Eingaben benachbarter Felder im 2-Spalten-Grid auf
    // gleicher Höhe (z.B. „Dressing" auf Höhe von „Gerichtauswahl"). Nur dann,
    // damit ohne Beschreibungen keine unnötigen Lücken entstehen.
    // v26.91: Der leere Platzhalter für eine fehlende Inline-Beschreibung wird
    // jetzt PRO ZEILE entschieden: nur reservieren, wenn der NEBEN diesem Feld
    // stehende Partner in derselben 2-Spalten-Zeile eine Beschreibung hat — sonst
    // (beide Felder ohne Beschreibung) entsteht keine leere Lücke mehr. Kennt der
    // Aufrufer die Zeile nicht (andere Render-Kontexte), gilt das bisherige
    // globale Verhalten (irgendein Feld hat eine Inline-Beschreibung).
    const fieldHasInlineDesc = (ff: EventSpecificField): boolean => ff.helpTextStyle === 'inline' && !!pickFieldHelp(ff);
    let reserveHelpSpace: boolean;
    if (typeof rowIndex === 'number' && rowList) {
      const partnerIdx = rowIndex % 2 === 0 ? rowIndex + 1 : rowIndex - 1;
      const partner = rowList[partnerIdx];
      reserveHelpSpace = !!(partner && fieldHasInlineDesc(partner));
    } else {
      reserveHelpSpace = (event?.eventSpecificFields || []).some(fieldHasInlineDesc);
    }
    const inlineHelpSlot = inlineHelpEl || (reserveHelpSpace
      // v26.17: Der leere Platzhalter muss dieselben Font-Metriken (fontSize/
      // lineHeight) wie der echte Inline-Hilfetext tragen, da sich 'minHeight'
      // in 'em' auf die EIGENE font-size bezieht. Ohne fontSize erbte der
      // Platzhalter die größere .form-group-Schrift und reservierte ~8px mehr
      // Höhe → die Eingabe eines Feldes OHNE Beschreibung stand tiefer als die
      // des Nachbarfeldes MIT Beschreibung (z.B. „Food Allergies" vs. „Hotel").
      ? <div aria-hidden="true" style={{ fontSize: '0.78rem', lineHeight: 1.45, marginTop: 2, marginBottom: 6, minHeight: '2.9em' }} />
      : null);
    // v19.0: Ausgefüllte Felder bekommen die gleiche grüne Hervorhebung wie die
    // ausgewählten Event-Sections (grüner Rand + zarter grüner Hintergrund).
    const fieldVal = vals[field.id];
    const isFieldFilled = field.type === 'checkbox' ? fieldVal === 'true' : !!(fieldVal && fieldVal.trim());
    const greenFilledStyle: React.CSSProperties = { borderColor: 'var(--dex-green, #86bc25)', boxShadow: '0 0 0 1px var(--dex-green, #86bc25) inset', background: 'rgba(134,188,37,0.06)' };
    const isErrEmpty = !!(showErrors && field.required && (field.type === 'checkbox' ? fieldVal !== 'true' : !fieldVal?.trim()));
    const inputStyleGreen: React.CSSProperties = isErrEmpty ? errorBorder : (isFieldFilled ? greenFilledStyle : {});
    return (
  <div className="form-group" key={field.id}>
    {field.type !== 'checkbox' && (
      <>
      <label className="form-label">
        {field.required && <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>}
        {displayLabel}
        {/* v9.17: konsistenter InfoTooltip statt simples i-Icon —
            gibt schönes Hover-Popover mit der vom Organizer
            beim Event-Anlegen hinterlegten Beschreibung.
            v18.18: nur im 'tooltip'-Modus; 'inline' rendert darunter. */}
        {displayHelp && !isInlineHelp && <InfoTooltip text={displayHelp.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim()} />}
      </label>
      {inlineHelpSlot}
      </>
    )}
    {field.type === 'select' && field.multi ? (
      // v11.89: Multi-Select-Dropdown — gleicher Look wie Single-Select,
      // beim Aufklappen Checkboxen pro Option. Werte werden weiterhin
      // " | "-getrennt im selben Feld vals[field.id]
      // gespeichert (kompatibel mit Record<string,string>).
      (() => {
        const sep = ' | ';
        const raw = (vals[field.id] || '').trim();
        const selected = raw ? raw.split(sep).map(s => s.trim()).filter(Boolean) : [];
        const isErr = !!(showErrors && field.required && selected.length === 0);
        // v17.20: Anzeige-Labels für den EN-Modus mappen. Der gespeicherte
        // Wert bleibt weiterhin der DE-Wert (positional gemappt), damit alle
        // anderen Stellen (Mails, Excel-Export, Admin-Center) konsistent
        // bleiben — wir tauschen ausschliesslich das Display-Label.
        const displayOptions = (field.options || []).map((o, i) => ({
          value: o,
          label: pickOptionLabel(field, i, o),
        }));
        return (
          <MultiSelectDropdown
            options={field.options || []}
            optionLabels={useEnVariants ? displayOptions.map(d => d.label) : undefined}
            value={selected}
            onChange={next => setVals({ ...vals, [field.id]: next.join(sep) })}
            placeholder={tEvent('reg.pleaseselect')}
            error={isErr}
          />
        );
      })()
    ) : field.type === 'select' && field.optionCategories && field.optionCategories.some(c => (c || '').trim()) ? (
      // v26.75: Vorfilter — zuerst Kategorie wählen, dann nur die passenden
      // Optionen zeigen (kürzere Liste). Die Kategorie-Auswahl liegt transient
      // in vals['<id>__cat']; gespeichert wird nur der eigentliche Optionswert.
      (() => {
        // v26.96: EINE Kombibox statt zwei Auswahlfeldern — die Kategorien sind
        // <optgroup>-Überschriften, darunter die zugehörigen Optionen. Optionen
        // OHNE Kategorie (leer = immer sichtbar) stehen ungruppiert am Ende.
        const cats = field.optionCategories || [];
        const opts = field.options || [];
        const distinctCats = Array.from(new Set(cats.map(c => (c || '').trim()).filter(Boolean)));
        return (
          <select className="form-select" value={vals[field.id] || ''} onChange={e => setVals({ ...vals, [field.id]: e.target.value })} style={inputStyleGreen}>
            <option value="" disabled hidden>{tEvent('reg.pleaseselect')}</option>
            {distinctCats.map(cat => (
              <optgroup key={cat} label={cat}>
                {opts.map((opt, i) => ((cats[i] || '').trim() === cat && (opt || '').trim())
                  ? <option key={`${cat}-${i}`} value={`${cat} ${opt}`}>{cat} {pickOptionLabel(field, i, opt)}</option>
                  : null)}
              </optgroup>
            ))}
            {opts.map((opt, i) => ((cats[i] || '').trim() === '' && (opt || '').trim())
              ? <option key={`nocat-${i}`} value={opt}>{pickOptionLabel(field, i, opt)}</option>
              : null)}
          </select>
        );
      })()
    ) : field.type === 'select' ? (
      <select className="form-select" value={vals[field.id] || ''} onChange={e => setVals({ ...vals, [field.id]: e.target.value })} style={inputStyleGreen}>
        <option value="" disabled hidden>{tEvent('reg.pleaseselect')}</option>
        {field.options && field.options.map((opt, i) => <option key={opt} value={opt}>{pickOptionLabel(field, i, opt)}</option>)}
      </select>
    ) : field.type === 'user' || field.type === 'roommate' ? (
      // v7.17: 'roommate' nutzt denselben Picker wie 'user' — der
      // einzige Unterschied ist dass 'roommate' beim Anmelden
      // automatisch eine Zimmerpartner-Mail an die ausgewählte
      // Person triggert (siehe EventContext). 'user' ist der
      // generische Personen-Picker ohne Mail-Versand.
      <UserFieldPicker
        value={vals[field.id] || ''}
        onChange={v => setVals({ ...vals, [field.id]: v })}
        // v29.40: Ist das Feld auf den Verteilerkreis begrenzt, findet die
        // Suche nur Personen, die das Event auch sehen. Der Picker übernimmt
        // ausschließlich Treffer aus der Liste (kein freier Text), damit
        // reicht das Filtern der Suche — es gibt keinen Umweg daran vorbei.
        searchUsers={field.audienceOnly ? searchUsersInAudience : searchUsers}
        searchUserByEmail={searchUser}
        placeholder={tEvent('reg.userfield.placeholder')}
        errorStyle={showErrors && field.required && !vals[field.id]?.trim() ? errorBorder : {}}
        // v26.60: „Person wird benachrichtigt"-Hinweis nur, wenn die
        // Zimmerpartner-Anfrage-Mail nicht abgeschaltet wurde.
        hint={[
          field.type === 'roommate' && field.notifyRoommate !== false ? tEvent('reg.userfield.notifyhint') : '',
          field.audienceOnly ? (locale === 'de' ? 'Auswählbar sind nur Personen, die zu diesem Event eingeladen sind.' : 'Only people invited to this event can be selected.') : '',
        ].filter(Boolean).join(' ') || undefined}
        forcedIsDe={locale === 'de'}
      />
    ) : field.type === 'checkbox' ? (
      // v11.91: Checkbox bekommt jetzt eine ordentliche Karten-Box mit
      // gleichem Look wie die Dropdown-Inputs — vorher war die Mini-
      // Checkbox neben den Dropdowns visuell „verloren". Der Label-Text
      // sitzt oben (analog zu den anderen Feldern, damit die Zeilen
      // horizontal aligned sind), drinnen ein deutlich vergrößerter
      // Checkbox + kurzer „Ja, bestätigen"-Hinweis.
      <>
        <label className="form-label">
          {field.required && <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>}
          {displayLabel}
          {displayHelp && !isInlineHelp && <InfoTooltip text={displayHelp.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim()} />}
        </label>
        {inlineHelpSlot}
        <label
          style={{
            display: 'flex', alignItems: 'center', gap: 10,
            cursor: 'pointer',
            // v11.93: exakt gleiche Höhe wie .form-select — 12px 16px Padding,
            // 1.5px Border, 12px Radius. Vorher 44px minHeight = optisch
            // höher als die Dropdowns daneben.
            padding: '12px 16px',
            border: showErrors && field.required && vals[field.id] !== 'true'
              ? '1.5px solid var(--dex-red)'
              : '1.5px solid var(--dex-gray-200)',
            borderRadius: 12,
            background: vals[field.id] === 'true' ? 'rgba(134,188,37,0.10)' : 'var(--dex-white, #fff)',
            transition: 'background 0.12s',
          }}
        >
          <input
            type="checkbox"
            checked={vals[field.id] === 'true'}
            onChange={e => setVals({ ...vals, [field.id]: e.target.checked ? 'true' : 'false' })}
            style={{ width: 16, height: 16, accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer', flexShrink: 0 }}
          />
          <span style={{ fontSize: '0.95rem', color: 'var(--dex-gray-800)' }}>
            {/* v11.94: Organizer kann den Text neben der Checkbox im Wizard
                pro Feld setzen (field.confirmLabel). Default: „Ja, bestätigen".
                v17.20: pickFieldConfirmLabel zieht im EN-Modus den
                confirmLabelEn-Wert; fällt sonst auf den DE-Wert. */}
            {(displayConfirmLabel && displayConfirmLabel.trim())
              || (eventLocale === 'de' ? 'Ja, bestätigen' : 'Yes, confirm')}
          </span>
        </label>
        {field.externalLinks && field.externalLinks.length > 0 && (
          <div style={{ marginTop: 4, fontSize: '0.78rem' }}>
            {field.externalLinks.map((l, i) => (
              <span key={l.url}>
                {i > 0 && <span style={{ color: 'var(--dex-gray-300)', margin: '0 6px' }}>|</span>}
                <a
                  href={l.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  onClick={e => e.stopPropagation()}
                  style={{ color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline' }}
                >
                  {l.label}
                </a>
              </span>
            ))}
          </div>
        )}
      </>
    ) : field.type === 'document' ? (
      // v19.0: Dokument-Upload (PDF/Bild). Datei wird in pendingDocFiles
      // gehalten und nach erfolgreicher Anmeldung als Attachment hochgeladen.
      // Im Team-Mitglied-Kontext (store gesetzt) nicht unterstützt.
      store ? (
        <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
          {locale === 'de' ? 'Dokument-Upload pro Team-Mitglied wird nicht unterstützt.' : 'Per-member document upload is not supported.'}
        </div>
      ) : (() => {
        const picked = pendingDocFiles[field.id] || null;
        const docErr = !!(showErrors && field.required && !picked && !parentAlreadyRegistered);
        return (
          <div>
            {picked ? (
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', border: '1.5px solid var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.06)', borderRadius: 12 }}>
                <Icon iconName="Attach" style={{ fontSize: 16, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                <span style={{ flex: 1, fontSize: '0.88rem', wordBreak: 'break-all' }}>{picked.name}</span>
                <button
                  type="button"
                  onClick={() => setPendingDocFiles(prev => ({ ...prev, [field.id]: null }))}
                  title={locale === 'de' ? 'Datei entfernen' : 'Remove file'}
                  style={{ background: 'var(--dex-gray-200)', border: 'none', color: 'var(--dex-gray-700)', width: 22, height: 22, borderRadius: '50%', cursor: 'pointer', fontSize: '0.9rem', lineHeight: 1, flexShrink: 0 }}
                >×</button>
              </div>
            ) : (
              <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, padding: '10px 16px', border: `1.5px dashed ${docErr ? 'var(--dex-red)' : 'var(--dex-gray-300)'}`, borderRadius: 12, cursor: 'pointer', fontSize: '0.88rem', color: 'var(--dex-gray-600)' }}>
                <Icon iconName="Upload" style={{ fontSize: 16 }} />
                {locale === 'de' ? 'Datei wählen (PDF, JPG, PNG)' : 'Choose file (PDF, JPG, PNG)'}
                <input
                  type="file"
                  accept="application/pdf,image/png,image/jpeg,.pdf,.png,.jpg,.jpeg"
                  style={{ display: 'none' }}
                  onChange={e => {
                    const f = e.target.files?.[0] || null;
                    if (f && f.size > 10 * 1024 * 1024) {
                      // eslint-disable-next-line no-alert
                      showAlert(locale === 'de' ? 'Die Datei ist zu groß (max. 10 MB).' : 'The file is too large (max. 10 MB).', { variant: 'error' });
                      e.target.value = '';
                      return;
                    }
                    setPendingDocFiles(prev => ({ ...prev, [field.id]: f }));
                  }}
                />
              </label>
            )}
            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>
              {locale === 'de'
                ? 'Wird nach dem Absenden hochgeladen. Du kannst es auch später über „Meine Events" ergänzen oder ersetzen.'
                : 'Uploaded after submitting. You can also add or replace it later via „My Events".'}
            </div>
          </div>
        );
      })()
    ) : field.type === 'daterange' ? (
      // v28.63: Übernachtungs-Zeitraum — Anreise + Abreise in einem Feld, die
      // Nächte ergeben sich daraus. Ersetzt die frühere Kombination aus
      // „Hotel: ja/nein" plus „Zusatznächte"-Auswahlliste; die Hotel-Planung
      // im Organizer Center liest den Zeitraum direkt aus.
      <StayRangePicker
        value={vals[field.id] || ''}
        onChange={(next: string) => setVals({ ...vals, [field.id]: next })}
        isDe={locale === 'de'}
        rangeStart={field.rangeStart}
        rangeEnd={field.rangeEnd}
        maxNights={field.maxNights}
        required={field.required}
      />
    ) : field.type === 'date' ? (
      // v24.25: Datums-Feld — Kalender-Auswahl. Mit withTime zusätzlich Uhrzeit
      // (datetime-local). Der Wert wird als String gespeichert (wie alle
      // Custom-Field-Antworten).
      <input
        className="form-input"
        type={field.withTime ? 'datetime-local' : 'date'}
        value={vals[field.id] || ''}
        onChange={e => setVals({ ...vals, [field.id]: e.target.value })}
        style={inputStyleGreen}
      />
    ) : (
      <input className="form-input" value={vals[field.id] || ''} onChange={e => setVals({ ...vals, [field.id]: e.target.value })} placeholder={displayLabel} style={inputStyleGreen} />
    )}
  </div>
    );
  };

  // v30.66: Props-Buendel fuer die ausgelagerten Teilbaeume der Anmeldeseite.
  // Bewusst hier, direkt vor dem return — davor stehen alle Deklarationen; ein
  // Buendel weiter oben waere ein TDZ-Fehler auf die spaeter deklarierten Handler.
  // Die Anzeige-Bedingungen bleiben beim Aufrufer, die Komponenten rendern unbedingt.
  const fallbackDialogModalProps = {
    fallbackDialog, locale, performRegistration, setFallbackDialog, setPreferredStarterType, splitLabelA,
    splitLabelB,
  };
  const proxyWizardModalProps = {
    canCreateEvents, checkRegistrationByEmail, currentUser, email, event, externalEmailConfirmedRef,
    externalPerson, firstName, isAllowedTargetForAssistant, isAssistant, isSearchingUser, locale,
    otherConsentConfirmed, pickedUserProfile, pickProxyUser, proxyStep, searchTimerRef, searchUsers,
    setEmail, setExternalPerson, setFirstName, setIsSearchingUser, setMassImportOpen, setMassImportResult,
    setMassImportRows, setMassImportStep, setOtherConsentConfirmed, setPickedUserProfile, setProxyStep, setRegisterForOther,
    setSurname, setThirdPartyCheck, setUserResults, setUserSearch, setUserSearchIncludeIntl, showAlert,
    surname, t, thirdPartyCheck, userResults, userSearch, userSearchIncludeIntl,
  };
  const massImportModalProps = {
    locale, massImportBusy, massImportMode, massImportOpen, massImportProgress, massImportResolving,
    massImportResult, massImportRows, massImportStep, massImportText, resolveMassImport, runMassImport,
    setMassImportMode, setMassImportOpen, setMassImportResult, setMassImportRows, setMassImportStep, setMassImportText,
  };
  const submitConfirmModalProps = {
    childEvents, childTermPlural, confirmDialogAck, confirmDialogConfirmedRef, confirmDialogOpen, confirmDraftParent,
    confirmDraftSessions, event, handleSubmit, locale, registerForOther, resolveMainEventLabel,
    selectedSessions, sessionFieldValues, setConfirmDialogAck, setConfirmDialogOpen, setConfirmDraftParent, setConfirmDraftSessions,
    setPendingSubEventModal, setRegisterForParent, setSelectedSessions, willRegisterParent,
  };
  const externalEmailWarningModalProps = {
    email, externalEmailConfirmedRef, externalEmailWarning, handleSubmit, locale, setExternalEmailWarning,
  };
  const ccSelfModalProps = {
    ccSelfDecidedRef, ccSelfModalOpen, ccSelfRef, firstName, handleSubmit, locale,
    setCcSelfModalOpen, surname,
  };
  const assistantModalProps = {
    assistantModalDecidedRef, assistantModalOpen, delegateAssistValue, delegateChoiceRef, handleSubmit, locale,
    parsedDelegateAssist, searchUser, searchUsers, setAssistantModalOpen, setDelegateAssistEnabled, setDelegateAssistValue,
  };
  const subEventFieldsModalProps = {
    childEvents, childTermSingular, locale, pendingSubEventModal, setPendingSubEventModal, setSelectedSessions,
    setSessionFieldValues,
  };
  const locationBannerProps = {
    currentUser, event, t,
  };
  const deadlineBannerProps = {
    event, isFullyClosed, locale, t,
  };
  const submitOverlayProps = {
    displayProgress, locale, submitProgressLabel,
  };
  const demoBannerProps = {
    locale,
  };
  const eventCardProps = {
    cachedImage, cachedZoomImage, circleSize, currentUser, event, heroImgUrl,
    imgAspectReady, imgCircleNotch, imgHovered, imgSlotH, imgSlotW, imgZoomed,
    isMobile, locale, setImgHovered, setImgZoomed, showOrbPlaceholder, usesMailImage,
  };
  const personalDataSectionProps = {
    canCreateEvents, canRegisterForOther, ccSelfDecidedRef, ccSelfRef, currentUser, email,
    errorBorder, event, externalEmailConfirmedRef, externalPerson, firstName, isAssistant,
    isMobile, isTeamCapable, isTeamMode, locale, parentAlreadyRegistered, pickedUserProfile,
    profileCardExpanded, proxyStep, registerForOther, salutation, setEmail, setExternalPerson,
    setFirstName, setIsTeamMode, setMassImportOpen, setMassImportResult, setMassImportRows, setMassImportStep,
    setOtherConsentConfirmed, setPendingJoinTeam, setPickedUserProfile, setProfileCardExpanded, setProxyStep, setRegisterForOther,
    setSalutation, setSurname, setThirdPartyCheck, setUserResults, setUserSearch, showErrors,
    surname, t, teamSize, thirdPartyCheck,
  };
  const teamRegistrationCardProps = {
    errorBorder, event, isMobile, isTeamMode, locale, parseTeamMember,
    renderRegField, searchUser, searchUsers, setTeamConsentConfirmed, setTeamMemberFields, setTeamMembers,
    setTeamName, showErrors, teamConsentConfirmed, teamMemberApplicableFields, teamMemberFields, teamMembers,
    teamName, teamPartialAllowed,
  };
  const openTeamsListProps = {
    event, isMobile, locale, openTeams, pendingJoinTeam, togglePendingJoinTeam,
  };
  const eventSpecificSectionProps = {
    childEvents, childOneDe, childTermPlural, childTermSingular, dayHoverKey, durchCap,
    event, eventSpecific, funCap, hasStarterBlockMapping, hiddenChildCount, isAdmin,
    isMobile, isOrganizer, isSessionsOnlyMode, isSplitGroup, locale, parentAlreadyRegistered,
    parentFullNoWaitlist, parentRegBlocked, preferredStarterType, registerForOther, registerForParent, renderMainFieldsSection,
    renderRegField, renderSubEventInlineFields, resolveMainEventLabel, selectedSessions, sessionFieldValues, sessionMeta,
    setDayHoverKey, setEventSpecific, setPendingSubEventModal, setPreferredStarterType, setRegisterForParent, setSelectedSessions,
    setSessionFieldValues, showErrors, splitLabelA, splitLabelB, starterCounts, subOpenFrom,
    t, tEvent,
  };
  const registrationActionBarProps = {
    childEvents, childOneDe, childTermPlural, childTermSingular, email, event,
    handleDecline, handleSubmit, isDeclining, isSubmitting, isTeamMode, liveStats,
    locale, nothingToSubmit, otherConsentConfirmed, parentAlreadyRegistered, pendingJoinTeam, registerForOther,
    resolveMainEventLabel, selectedSessions, t, teamMembersParsed, teamValidation, thirdPartyCheck,
    willRegisterParent,
  };
  const privacyNoteProps = {
    event, t,
  };
  return (
    <div className="page-container">
      {/* v30.43: Der blaue Hinweiskasten „Du bist Organizer bzw. Admin …" ist
          hier entfallen. Er stand vier Zeilen breit über der Seite, nur um EINE
          Einstellung anzubieten, und gab es ausschließlich auf dieser Seite.
          Der Wechselschalter Organizer-/User-Ansicht sitzt jetzt im Header
          (`Header.tsx`, v30.43): Der Zustand ist dort immer sichtbar, gilt
          appweit und ist mit einem Klick umgelegt. `previewAsUser` und die
          Rollen-Absenkung dahinter sind unverändert — nur der Einstieg ist
          umgezogen. */}
      {showLocationBanner && <LocationBanner {...locationBannerProps} />}
      {/* Deadline-Banner für Organizer/Admin: die Registrierungsfrist ist abgelaufen,
          das Formular wird aber trotzdem angezeigt (Admin/Organizer darf weiter registrieren).
          Der Ton entspricht dem Location-Banner: "als normaler User könntest du dich nicht registrieren". */}
      {isDeadlinePassed && (isOrganizer || isAdmin) && <DeadlineBanner {...deadlineBannerProps} />}
      {/* v11.33: Submit-Overlay mit Spinner + Prozent + Live-Label.
          Wird während des gesamten Anmelde-Flows (Parent + alle Sub-Events
          + Bestätigungen) eingeblendet, sodass der User auch bei langen
          Submits klares Feedback bekommt. */}
      {isSubmitting && <SubmitOverlay {...submitOverlayProps} />}
      {/* v18: Demo-Hinweis-Banner. Das Demo-Event wird ansonsten exakt wie ein
          echtes Event gerendert (so wie es in der Realität nutzbar wäre) —
          keine künstlichen Showcase-Elemente. Nur die echte Anmeldung ist
          deaktiviert. */}
      {event && event.isDemoShowcase && <DemoBanner {...demoBannerProps} />}
      <div className="registration-layout">
        {/* v28.2 „Geführte Schritte": Station 1 — Dein Event. Der frühere
            „Ausgewähltes Event"-Pill-Header entfällt (Step-Label ersetzt ihn). */}
        <div className="reg-step-head">
          <span className="reg-step-num">1</span>
          <span className="reg-step-label">{locale === 'de' ? 'Dein Event' : 'Your event'}</span>
        </div>
        <EventCard {...eventCardProps} />

        {/* v18.73: Die „Offene Teams"-Box ist nach UNTEN gewandert (unter die
            „Ich melde mich + mein Team an"-Karte) — siehe weiter unten. Oben
            steht jetzt immer zuerst die persönliche Daten-Karte, dann die
            event-spezifischen Infos. */}

        {/* v28.2: Station 2 — Deine Daten. */}
        <div className="reg-step-head">
          <span className="reg-step-num">2</span>
          <span className="reg-step-label">{locale === 'de' ? 'Deine Daten — automatisch aus M365' : 'Your details — automatically from M365'}</span>
        </div>
        {/* Persönliche Daten */}
        <PersonalDataSection {...personalDataSectionProps} />

        {/* v11.82: Team-Anmeldung-Card — separat unter „Persönliche Daten",
            nur sichtbar wenn der Toggle aktiv ist. */}
        {isTeamCapable && isTeamMode && !registerForOther && !parentAlreadyRegistered && <TeamRegistrationCard {...teamRegistrationCardProps} />}

        {/* v18.73: Offene Teams — sichtbar wenn der Organizer „Offene Slots
            öffentlich sichtbar" aktiviert hat und es Teams gibt, denen Plätze
            fehlen. Steht jetzt UNTER der „Ich melde mich + mein Team an"-Karte.
            Klick auf „Vormerken" wählt ein Team nur vor — die eigentliche
            Anmeldung (inkl. der oben/unten ausgefüllten persönlichen +
            event-spezifischen Felder) passiert erst über den „Anmelden"-Button.
            */}
        {event && event.teamRegistrationEnabled && event.teamOpenSlotsVisible && !registerForOther && openTeamsLoaded && openTeams.length > 0 && !parentAlreadyRegistered && <OpenTeamsList {...openTeamsListProps} />}

        {/* Eventspezifische Felder (inkl. Split-Capacity Starter-Typ-Auswahl wenn
            beide Kapazitäten > 0; bei nur einem verfügbaren Typ wird dieser
            automatisch gesetzt und gar nicht angezeigt). v10.20: Sessions-/
            Hauptevent-Auswahl ist hierher gewandert (vorher links unter der
            Event-Karte). */}
        {/* v18.73: Die „Event-spezifische Informationen"-Karte nur anzeigen,
            wenn es dort tatsächlich etwas auszufüllen/auszuwählen gibt — also
            Custom-Felder, eine Gruppen-Auswahl (Split) ODER eine Sub-Event-
            Auswahl. Sonst (leeres „Keine zusätzlichen Informationen
            erforderlich") wird die Karte komplett ausgeblendet. */}
        {/* v28.2: Station 3 — Anmeldung abschließen (immer sichtbar; die
            Event-Felder-Karte darunter nur, wenn es etwas auszufüllen gibt). */}
        <div className="reg-step-head">
          <span className="reg-step-num">3</span>
          <span className="reg-step-label">{locale === 'de' ? 'Anmeldung abschließen' : 'Complete your registration'}</span>
        </div>
        {/* v29.9: …und die Karte muss auch dann erscheinen, wenn für diese
            Person KEIN Programmpunkt sichtbar ist — sonst fehlt der Hinweis
            darunter genau in dem Fall, für den er gedacht ist. */}
        {(event.eventSpecificFields.length > 0 || isSplitGroup || childEvents.length > 0
          || (hiddenChildCount > 0 && event.subEventsOnlyMode)) && <EventSpecificSection {...eventSpecificSectionProps} />}
      </div>

      {/* v11.4: Fehlermeldung + Action-Buttons stehen jetzt direkt unter
          dem registration-layout (also unter der Eventspez-Karte) und
          NICHT mehr unterhalb des Datenschutz-Hinweises. Der Hinweis ist
          eine Fußnote und gehört ans Seitenende — die Aktions-Buttons
          gehören thematisch zur Anmelde-Maske. */}

      {/* Fehlermeldung */}
      {error && (
        <div className="mt-16" style={{ padding: '12px 16px', background: 'rgba(239,68,68,0.1)', border: '1px solid var(--dex-red)', borderRadius: 'var(--dex-radius-md)', color: 'var(--dex-red)', fontSize: '0.9rem' }}>
          {error}
        </div>
      )}

      {/* v24.48: Die „Meine Assistenz"-Abfrage ist von inline auf ein Modal
          beim Register-Klick umgestellt (siehe assistantModalOpen unten). */}

      {/* Buttons + Platz-Badge (v24.57) — Badge über den zentrierten Buttons. */}
      <RegistrationActionBar {...registrationActionBarProps} />

      {/* Datenschutz-Hinweis als Fußnote ganz unten.
          v11.93: Breite auf 1100px begrenzt + zentriert (analog
          .registration-layout), damit der Text nicht über die ganze
          App-Breite läuft. */}
      <PrivacyNote {...privacyNoteProps} />

      {/* Fallback-Dialog (seit v6.5): Wunsch-Starter-Typ voll, aber Alternative frei.
          User entscheidet explizit zwischen Umsteigen oder Warteliste. */}
      {fallbackDialog && <FallbackDialogModal {...fallbackDialogModalProps} />}

      {/* v9.22: Modal für externe Email-Anmeldung */}
      {/* v18.13: Massenimport-Modal. */}
      {/* v26.76: Geführter Wizard für die stellvertretende Anmeldung (interner
          Fall): Schritt 1 Person suchen (mit Foto), Schritt 2 Zustimmung. Nach
          „OK" ist die Person übernommen und die persönlichen Felder vorbefüllt.
          Externe Person / Massenimport bleiben als eigene Wege erhalten. */}
      {proxyStep > 0 && <ProxyWizardModal {...proxyWizardModalProps} />}
      {massImportOpen && <MassImportModal {...massImportModalProps} />}

      {/* v18.75: Sicherheitshinweis-Dialog vor dem Absenden (pro Event). */}
      {confirmDialogOpen && event && <SubmitConfirmModal {...submitConfirmModalProps} />}

      {externalEmailWarning && <ExternalEmailWarningModal {...externalEmailWarningModalProps} />}

      {/* v19.6: CC-Frage bei stellvertretender INTERNER Anmeldung. Erscheint
          nach dem „Anmelden"-Klick und vor der eigentlichen Anmeldung — der/die
          Anmeldende (Organizer, Co-Organizer oder Assistenz) entscheidet, ob er/
          sie selbst auf CC der Bestätigungs-Mail soll. */}
      {ccSelfModalOpen && <CcSelfModal {...ccSelfModalProps} />}

      {/* v24.48: Assistenz-Abfrage — erscheint nach „Register"-Klick für
          Partner/Director. People-Picker (Suche nach Name/E-Mail). */}
      {assistantModalOpen && <AssistantModal {...assistantModalProps} />}

      {/* v10.12: Sub-Event Custom-Fields Modal — wird geöffnet wenn ein
          Sub-Event mit eigenen Custom-Fields angecheckt wird. Der User muss die
          Antworten ausfüllen + bestätigen, dann wandert die Session in
          selectedSessions und die Werte in sessionFieldValues. Beim „Abbrechen"
          wird die Session NICHT angecheckt. */}
      {pendingSubEventModal && <SubEventFieldsModal {...subEventFieldsModalProps} />}
    </div>
  );
}

