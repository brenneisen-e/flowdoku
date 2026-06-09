/**
 * Registrierungsseite fuer ein einzelnes Event
 *
 * Drei-Spalten-Layout: Event-Info | persoenliche Daten | eventspezifische Felder
 * Speichert die Registrierung in der SharePoint-Teilnehmerliste des Events.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents, collectCcEmailsFromFields } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage, translations as appTranslations, Locale } from '../context/LanguageContext';
import { Salutation, EventSpecificField } from '../types';
import { Icon } from '@fluentui/react/lib/Icon';
import { Trash2, Send, X } from './Icons';
import { InfoTooltip } from './InfoTooltip';
import { MultiSelectDropdown } from './MultiSelectDropdown';
import OrganizerList from './OrganizerList';
import Modal from './Modal';
import InternationalSearchToggle from './InternationalSearchToggle';
import { UserFieldPicker } from './UserFieldPicker';

function formatDate(iso: string): string {
  const d = new Date(iso);
  return (
    d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
    ' ' +
    d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
  );
}

// v11.94: Kompakt-Format für die Datum-Badge in der Event-Card —
// wenn Start- und End-Tag identisch sind, nur einmal das Datum +
// "HH:MM - HH:MM". Sonst beide voll mit "-" dazwischen.
function formatDateRange(startIso: string, endIso: string): string {
  if (!startIso) return '';
  const start = new Date(startIso);
  const end = endIso ? new Date(endIso) : null;
  const dayFmt = { day: '2-digit', month: '2-digit', year: 'numeric' } as const;
  const timeFmt = { hour: '2-digit', minute: '2-digit' } as const;
  if (!end || isNaN(end.getTime())) {
    return `${start.toLocaleDateString('de-DE', dayFmt)} ${start.toLocaleTimeString('de-DE', timeFmt)}`;
  }
  const sameDay = start.toDateString() === end.toDateString();
  if (sameDay) {
    return `${start.toLocaleDateString('de-DE', dayFmt)} ${start.toLocaleTimeString('de-DE', timeFmt)} – ${end.toLocaleTimeString('de-DE', timeFmt)}`;
  }
  return `${start.toLocaleDateString('de-DE', dayFmt)} ${start.toLocaleTimeString('de-DE', timeFmt)} – ${end.toLocaleDateString('de-DE', dayFmt)} ${end.toLocaleTimeString('de-DE', timeFmt)}`;
}

// v18.74: Externe Adresse = kein Deloitte-Deutschland-Postfach (@deloitte.de).
const isExternalEmailAddr = (e: string): boolean => {
  const v = (e || '').trim();
  return !!v && !/@(.*\.)?deloitte\.de$/i.test(v);
};

// v18.74: Strengere Plausibilitätsprüfung gegen Tippfehler bei externen
// Adressen — fängt fehlende/zu kurze TLD, doppelte Punkte, mehrere @, führende/
// abschließende Punkte und Whitespace/Kommas ab. Verifiziert NICHT die Existenz
// des Postfachs (das geht clientseitig nicht), aber blockt offensichtliche
// Vertipper.
const isPlausibleEmail = (e: string): boolean => {
  const v = (e || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[a-zA-Z]{2,}$/.test(v)) return false;
  if ((v.match(/@/g) || []).length !== 1) return false;
  if (v.indexOf('..') >= 0) return false;
  if (/[\s,;]/.test(v)) return false;
  const [local, domain] = v.split('@');
  if (!local || !domain) return false;
  if (local.startsWith('.') || local.endsWith('.')) return false;
  if (domain.startsWith('.') || domain.endsWith('.') || domain.startsWith('-')) return false;
  return true;
};

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
  const { events, registerForEvent, registerTeam, cancelRegistration, declineEvent, checkRegistrationByEmail, getMyRegistration, getAllRegistrations, childEventsOf, listOpenTeamsForEvent, joinTeam, createTeamJoinRequest, updateMyRegistration, uploadFieldDocument } = useEvents();
  const { currentUser } = useCurrentUser();
  const { searchUsers, searchUser, isAdmin } = useRoles();
  const { locale: appLocale } = useLanguage();
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
  const tEvent = React.useCallback((key: string): string => {
    return appTranslations[eventLocale][key] || appTranslations['en'][key] || appTranslations['de'][key] || t(key) || key;
  }, [eventLocale, t]);
  // v17.20: Lookup-Helfer fuer die EN-Varianten eines Custom-Fields. Greift
  // nur, wenn der Bilingual-Toggle des Events an ist UND die App-Locale des
  // Teilnehmers `en` ist. Sonst still Fallback auf den DE-Wert. Index-Mapping
  // der Optionen ist positional — DE-Option i ↔ EN-Option i.
  // v17.22: `useEnVariants` steuert NUR die Anzeige-Labels. Die gespeicherten
  // Werte bleiben in JEDEM Fall die kanonischen DE-Originale: Single-Select
  // rendert `<option value={DE-Wert}>{EN-Anzeige}</option>`, Multi-Select gibt
  // `options={field.options}` (DE) als Wert weiter und nutzt `optionLabels`
  // nur fuer die Darstellung. Deshalb ist auch der „Register-for-Other"-Pfad
  // unkritisch: meldet ein EN-Organizer eine DE-Person an, sieht der Organizer
  // die EN-Labels (er fuellt das Formular), gespeichert wird aber der
  // DE-Wert — die Zielperson und die Bestaetigungs-Mail (event.emailLanguage)
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
  // fuer EVENT B solche Admin-Aktionen ausfuehren. Admin darf global alles.
  // v19.6: Co-Organizer (event.coOrganizerEmails) zaehlen hier ausdruecklich
  // mit — vorher sah ein Co-Organizer den „Für andere registrieren"-Button
  // nicht, obwohl er das Event mitorganisiert. Serverseitig wird derselbe
  // Personenkreis in canRegisterForOthers() akzeptiert.
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isEventOrganizer = !!event && (
    event.organizerEmails.some(e => (e || '').toLowerCase() === currentEmailLc) ||
    (event.coOrganizerEmails || []).some(e => (e || '').toLowerCase() === currentEmailLc)
  );
  const isOrganizer = isEventOrganizer; // alten Namen behalten fuer Referenzen unten
  const canCreateEvents = isEventOrganizer || isAdmin; // statt tenant-weitem Organizer

  // Assistant-Ausnahme: User mit JobTitle "Assistant" / "Senior Assistant" duerfen
  // "Register for another person" nutzen, allerdings NUR fuer Director/Partner und
  // NUR fuer Events fuer die sie sich eh selber anmelden koennten (also nicht nach
  // Deadline). Der Deadline-Schutz greift automatisch, weil RegistrationPage fuer
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
    // Wichtig: wenn nur EIN Filter gesetzt ist, zaehlt nur dieser - egal ob AND/OR.
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
  // v18.74: „Person außerhalb Deloitte" — explizit eine externe Person
  // stellvertretend anmelden. Blendet den Deloitte-People-Picker aus und macht
  // Vorname/Nachname/E-Mail frei eintragbar. Die Zustimmung ist hier SCHRIFTLICH
  // einzuholen; es wird kein Outlook-Termin versendet (Bestätigungs-Mail mit
  // Organizer auf CC).
  const [externalPerson, setExternalPerson] = React.useState(false);

  // Wenn die Seite mit Intent 'register-other' geoeffnet wird (z.B. via "Register another person"
  // Button auf einer Karte, fuer die der Organizer/Admin schon selbst registriert ist),
  // direkt in den "Fuer andere registrieren"-Modus springen und Felder leeren.
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
  // v11.33: Submit-Overlay mit Fortschrittsanzeige (Prozent + Label).
  // Bei vielen Sub-Events / vielen Custom-Fields kann der Submit
  // mehrere Sekunden dauern — vorher hat der User nur einen disabled
  // Button gesehen ohne Feedback was gerade passiert.
  const [submitProgress, setSubmitProgress] = React.useState(0);
  const [submitProgressLabel, setSubmitProgressLabel] = React.useState('');
  const [error, setError] = React.useState('');
  const [showErrors, setShowErrors] = React.useState(false);
  // v11.91: showDescription wurde entfernt — Beschreibung ist immer offen.
  const [thirdPartyCheck, setThirdPartyCheck] = React.useState<{ alreadyRegistered: boolean; notInAudience: boolean; registeredName?: string; registeredDate?: string } | null>(null);

  // Seit v6.14: integrierte Session-Auswahl direkt auf der Registrierungsseite.
  // Der User kann auf EINER Seite wählen, ob er sich für das Haupt-Event und/oder
  // einzelne Sub-Events anmelden möchte. Bei B2Run-Parents zusätzlich pro Session
  // eine Durchstarter/Funstarter-Auswahl.
  const childEvents = React.useMemo(() => event ? childEventsOf(event.id) : [], [event?.id]);
  // v15.10: vom Organizer konfigurierbare Bezeichnung (z.B. „Event-Sections",
  // „Workshops"). Wenn gesetzt überschreibt das die Default-Übersetzung
  // („Sessions" / „Sub-Events") überall im RegistrationPage-UI.
  const childTermSingular = (event && event.childEventTermSingular) || '';
  const childTermPlural = (event && event.childEventTermPlural) || '';
  const [registerForParent, setRegisterForParent] = React.useState(true);
  const [selectedSessions, setSelectedSessions] = React.useState<Set<string>>(new Set());
  const [sessionStarterType, setSessionStarterType] = React.useState<Record<string, string>>({});
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
  const [sessionMeta, setSessionMeta] = React.useState<Record<string, { count: number; wasRegistered: boolean }>>({});
  const [myParentReg, setMyParentReg] = React.useState<{ Status?: string } | null>(null);
  const [sessionsOnlySubmitted, setSessionsOnlySubmitted] = React.useState(false);
  // v18.67: echtes Anmelde-Ergebnis (Angemeldet/Warteliste) aus der
  // Haupt-Registrierung — das Ergebnis-Modal nutzt das statt der gecachten
  // isFull-Schaetzung, die nach Cancel/Re-Register veraltet sein konnte und
  // faelschlich „Warteliste" zeigte, obwohl der User angemeldet wurde.
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
    && !(event?.requireSubEventSelection && childEvents.length > 0);
  const teamSize = event?.teamSize || 0;
  const teamPartialAllowed = !!event?.teamPartialAllowed;
  const [isTeamMode, setIsTeamMode] = React.useState(false);
  const [teamName, setTeamName] = React.useState('');
  const [teamMembers, setTeamMembers] = React.useState<string[]>([]);
  // v18.12: Custom-Field-Antworten pro Team-Mitglied (Slot-Index → {fieldId: value}).
  // So kann der Lead z.B. die Essenspraeferenz auch fuer jedes Teammitglied angeben.
  const [teamMemberFields, setTeamMemberFields] = React.useState<Record<number, Record<string, string>>>({});
  const [teamConsentConfirmed, setTeamConsentConfirmed] = React.useState(false);
  // v15.16: Bei „Für andere registrieren" (registerForOther) braucht es
  // ebenfalls eine explizite Bestätigung, dass die Person der Anmeldung
  // zugestimmt hat — analog zur Team-Anmelde-Pflicht.
  const [otherConsentConfirmed, setOtherConsentConfirmed] = React.useState(false);
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
  // Parser fuer People-Picker-Values im Format „DisplayName <email>".
  const parseTeamMember = (v: string): { displayName: string; email: string } | null => {
    const m = (v || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
    if (!m) return null;
    return { displayName: m[1].trim(), email: m[2].trim().toLowerCase() };
  };
  const teamMembersParsed = teamMembers.map(parseTeamMember);
  // v18.12: Custom-Fields, die pro Team-Mitglied abgefragt werden — alle
  // event-spezifischen Felder AUSSER Personen-Pickern (user/roommate) und
  // B2Run-Spezialfeldern; gruppen-spezifische Felder nur „fuer alle".
  const teamMemberApplicableFields = (event?.eventSpecificFields || []).filter(f =>
    f.type !== 'user' && f.type !== 'roommate' &&
    f.id !== 'b2run_startblock' && f.id !== 'b2run_mobilnummer' &&
    (!f.onlyForGroup || f.onlyForGroup === 'all')
  );
  // Validation des Team-Submits — Lead-Email darf nicht in der Member-Liste
  // sein, Member-Emails muessen untereinander disjunkt sein, im Pflicht-Modus
  // muessen alle Slots gefuellt sein.
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
    if (registerForOther) return; // Stellvertreter-Modus nicht unterstuetzt fuer Beitritt
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
          const meta: Record<string, { count: number; wasRegistered: boolean }> = {};
          const preselect = new Set<string>();
          const starterPre: Record<string, string> = {};
          for (const ce of childEvents) {
            if (registerForOther) {
              // v18.37: Stellvertreter-Modus — nur die Belegungszahl laden
              // (für die „X/Y belegt"-/Voll-Anzeige). KEINE Self-Vorbelegung,
              // weil getMyRegistration die Daten des eingeloggten Users liefert,
              // nicht die der angemeldeten Person. Der Assistent wählt die
              // Sub-Events frisch aus.
              const allRegs = await getAllRegistrations(ce.id);
              const count = (allRegs || []).filter(r => {
                const s = r.Status || '';
                return s === 'Angemeldet' || s === 'QR versendet' || s === 'Eingecheckt';
              }).length;
              meta[ce.id] = { count, wasRegistered: false };
            } else {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const [myReg, allRegs] = await Promise.all([
                getMyRegistration(ce.id) as Promise<{ Status?: string; StarterType?: string; PreferredStarterType?: string } | null>,
                getAllRegistrations(ce.id),
              ]);
              const wasRegistered = !!myReg && myReg.Status !== 'Abgemeldet';
              const count = (allRegs || []).filter(r => {
                const s = r.Status || '';
                return s === 'Angemeldet' || s === 'QR versendet' || s === 'Eingecheckt';
              }).length;
              meta[ce.id] = { count, wasRegistered };
              if (wasRegistered) {
                preselect.add(ce.id);
                const existingType = myReg?.StarterType || myReg?.PreferredStarterType;
                if (existingType) starterPre[ce.id] = existingType;
              }
            }
          }
          setSessionMeta(meta);
          // Im Stellvertreter-Modus startet die Auswahl leer (frische Anmeldung).
          setSelectedSessions(registerForOther ? new Set<string>() : preselect);
          setSessionStarterType(prev => ({ ...starterPre, ...prev }));
        } catch { /* */ }
      })();
    }
  }, [event?.id, registerForOther]);

  // Bild-Orientierung erkennen (Hochkant -> links | Querformat -> oben)
  const [imgOrientation, setImgOrientation] = React.useState<'portrait' | 'landscape'>('landscape');
  React.useEffect(() => {
    if (!event?.imageUrl) { setImgOrientation('landscape'); return; }
    const img = new Image();
    img.onload = () => {
      setImgOrientation(img.naturalHeight > img.naturalWidth ? 'portrait' : 'landscape');
    };
    img.src = event.imageUrl;
  }, [event?.imageUrl]);

  // B2Run Split-Capacity: aktuelle Auslastung pro Typ laden
  // Split-UI nur wenn BEIDE Starter-Typen verfuegbar sind (>0). Wenn der Admin eine
  // Kapazitaet auf 0 gesetzt hat, gibt es faktisch nur einen Typ — dann keine Auswahl
  // anzeigen und den einzig verfuegbaren Typ automatisch setzen (siehe useEffect unten).
  const durchCap = (event && typeof event.durchstarterCapacity === 'number') ? event.durchstarterCapacity : 0;
  const funCap = (event && typeof event.funstarterCapacity === 'number') ? event.funstarterCapacity : 0;
  const isSplitGroup = !!event && durchCap > 0 && funCap > 0;
  // v10.20: frei waehlbare Bezeichnungen aus dem Event laden, mit Fallback auf
  // die historischen B2Run-Defaults 'Durchstarter' / 'Funstarter'. Die internen
  // Werte fuer SP-Persistenz (StarterType-Spalte) bleiben unveraendert — das
  // Label ist reines UI.
  const splitLabelA = (event?.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
  const splitLabelB = (event?.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
  const singleStarterType: string = (!event || (durchCap <= 0 && funCap <= 0))
    ? '' // kein B2Run-Event ueberhaupt
    : (durchCap > 0 && funCap <= 0) ? 'Durchstarter'
    : (funCap > 0 && durchCap <= 0) ? 'Funstarter'
    : ''; // beide > 0 -> User muss waehlen (Split-UI)

  // Auto-Set: wenn nur ein Starter-Typ verfuegbar ist, direkt diesen Typ als
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
  } | null>(null);
  const [isSearchingUser, setIsSearchingUser] = React.useState(false);
  const [userSearchIncludeIntl, setUserSearchIncludeIntl] = React.useState(false);
  const searchTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  if (!event) {
    return (
      <div className="page-container text-center">
        <h2>{t('reg.eventnotfound')}</h2>
        <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
          {t('reg.backtoevents')}
        </button>
      </div>
    );
  }

  // Registrierungs-Deadline pruefen
  const isDeadlinePassed = event.registrationDeadline && new Date(event.registrationDeadline) < new Date();

  if (isDeadlinePassed && !isOrganizer && !isAdmin) {
    return (
      <div className="page-container">
        <div className="card" style={{ position: 'relative', overflow: 'hidden' }}>
          <div style={{
            height: 200,
            background: event.imageUrl
              ? `url(${event.imageUrl}) center/cover no-repeat`
              : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
            borderRadius: '16px 16px 0 0',
          }} />
          <div style={{ padding: 32, textAlign: 'center' }}>
            <Icon iconName="Clock" style={{ fontSize: 48, color: 'var(--dex-orange)', marginBottom: 16 }} />
            <h2 style={{ marginBottom: 8 }}>{t('reg.deadlinepassed.title')}</h2>
            <p style={{ color: 'var(--dex-gray-600)', marginBottom: 8 }}>
              {t('reg.deadlinepassed.text')}
            </p>
            <p style={{ color: 'var(--dex-gray-400)', fontSize: '0.85rem' }}>
              {t('reg.deadlinepassed.date')}: {new Date(event.registrationDeadline).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}
            </p>
            <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>
              {t('reg.backtoevents')}
            </button>
          </div>
        </div>
      </div>
    );
  }

  const isFull = event.maxParticipants > 0 && event.currentParticipants >= event.maxParticipants;
  const errorBorder = { border: '2px solid var(--dex-red)' };

  const parentAlreadyRegistered = !!(myParentReg && myParentReg.Status && myParentReg.Status !== 'Abgemeldet');
  // "Hauptevent wird jetzt angemeldet" gilt nur, wenn der Parent-Checkbox an ist
  // UND der User nicht bereits angemeldet ist. Bei bereits angemeldetem Parent
  // wird die Parent-Registrierung nicht nochmal ausgelöst.
  const willRegisterParent = registerForParent && !parentAlreadyRegistered && !registerForOther && !(event && event.subEventsOnlyMode);
  // Fürs Registrieren für andere bleibt der alte Flow: Parent wird immer registriert,
  // keine Session-Auswahl (siehe Render).
  const isSessionsOnlyMode = !willRegisterParent && !registerForOther && !parentAlreadyRegistered;

  // v9.22: Warning-Modal fuer externe Email-Anmeldung (durch Organizer fuer
  // Drittpersonen die noch kein Deloitte-Postfach haben). Default: nicht
  // erlaubt; Organizer kann nach Bestaetigung trotzdem fortfahren — die
  // Bestaetigungsmail geht dann nicht an die externe Adresse, sondern an
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
  // ccSelfRef haelt die Entscheidung (true = auf CC).
  const [ccSelfModalOpen, setCcSelfModalOpen] = React.useState(false);
  const ccSelfDecidedRef = React.useRef(false);
  const ccSelfRef = React.useRef(false);

  const handleSubmit = async (): Promise<void> => {
    // v17.25: Demo-Showcase-Event — keine echte Anmeldung. Freundlicher
    // Hinweis statt SP-Roundtrip; der Context-Guard wuerde ohnehin no-oppen.
    if (event?.isDemoShowcase) {
      setError(locale === 'de'
        ? 'Dies ist ein Demo-Event — es wird keine echte Anmeldung gespeichert. Du kannst die Bereiche oben frei ausprobieren.'
        : 'This is a demo event — no real registration is stored. Feel free to explore the sections above.');
      return;
    }
    // Validierung Pflichtfelder
    setShowErrors(true);

    // Wenn der Haupt-Event-Checkbox aus ist und keine Session ausgewählt ist,
    // gibt es nichts zu tun.
    if (!willRegisterParent && !registerForOther && selectedSessions.size === 0 && !pendingJoinTeam) {
      setError(t('reg.nothing.selected') || 'Bitte wähle mindestens Haupt-Event oder eine Session aus.');
      return;
    }

    // v14.5: Wenn der Organizer `requireSubEventSelection` aktiviert hat
    // (Toggle in Schritt 6), MUSS der Teilnehmer mindestens ein Sub-Event
    // angehakt haben — sonst landet er ohne Kommunikation in der Liste,
    // weil die Hauptevent-Mails/Outlook deaktiviert sind.
    if (event && event.requireSubEventSelection && childEvents.length > 0 && selectedSessions.size === 0) {
      setError(t('reg.require.subevent') || (locale === 'de'
        ? `Für dieses Event musst du mindestens ein ${childTermSingular || 'Sub-Event'} auswählen — sonst kannst du dich nicht anmelden.`
        : `For this event you must pick at least one ${childTermSingular || 'sub-event'} — otherwise you cannot register.`));
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
    const willCollectMainFields = willRegisterParent || registerForOther
      || (isSubOnlyModeValidate && selectedSessions.size > 0 && !registerForOther)
      // v18.73: Beim vorgemerkten Team-Beitritt gelten dieselben Pflichtfelder
      // (Anrede + event-spezifische Felder) wie bei einer normalen Anmeldung —
      // der ganze Sinn der Änderung ist, dass diese Infos nicht übersprungen
      // werden können.
      || !!pendingJoinTeam;
    if (willCollectMainFields) {
      // v11.80: Anrede ist nur dann Pflichtfeld, wenn das Event das
      // Anrede-Dropdown auch tatsaechlich abfragt (event.askSalutation === true).
      // Sonst wird die Anrede gar nicht gerendert und bleibt leer.
      if (event.askSalutation && !salutation) {
        setError(t('reg.requiredfields'));
        return;
      }

      // Pflicht-Custom-Fields validieren. Checkbox-Pflichtfelder muessen 'true' sein,
      // alle anderen duerfen nach trim nicht leer sein.
      // B2Run: Mobilnummer ist nur Pflicht wenn Infoservice aktiviert; ansonsten
      // gilt das Feld als versteckt und wird uebersprungen.
      const missingRequired = event.eventSpecificFields
        .filter(f => {
          if (f.id === 'b2run_mobilnummer') {
            if (eventSpecific['b2run_infoservice'] !== 'true') return false;
            return !eventSpecific[f.id]?.trim();
          }
          // v7.21: Felder mit nicht erfuellter Sichtbarkeitsbedingung sind
          // ausgeblendet und duerfen die Validation nicht blockieren.
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
    // Organizer des Events noch Admin ist, aber via Assistant-Ausnahme fuer
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
    if ((willRegisterParent || registerForOther) && isSplitGroup && preferredStarterType && !pendingJoinTeam) {
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

    await performRegistration(preferredStarterType);
  };

  // v11.82: Team-Anmeldung absenden — Lead + N-1 Mitglieder per registerTeam.
  // Sub-Events werden im Team-Modus NICHT mitangemeldet — das ist Phase 2
  // (siehe Manual). Lead und Member bekommen jeweils nur den Hauptevent.
  const performTeamRegistration = async (starterTypeToUse: string): Promise<void> => {
    setError('');
    setIsSubmitting(true);
    setSubmitProgress(5);
    setSubmitProgressLabel(locale === 'de' ? 'Team-Anmeldung wird vorbereitet…' : 'Preparing team registration…');
    try {
      const customData: Record<string, string> = { salutation, ...eventSpecific };
      const leadEmail = email.trim();
      const leadFirstName = firstName.trim();
      const leadLastName = surname.trim();
      // v18.12: Custom-Field-Antworten pro Mitglied (nach Slot-Index) mitgeben.
      const members = teamMembersParsed
        .map((m, idx) => m ? { email: m.email, displayName: m.displayName, customData: { ...(teamMemberFields[idx] || {}) } } : null)
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
      const customData: Record<string, string> = { salutation, ...eventSpecific };
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
        ...eventSpecific,
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

      let anySuccess = false;
      let parentOk = true;

      // Verfeinerte Progress-Stufen je nach Anzahl Sub-Events:
      // - parent: 5 → 30 → 50 (wenn Parent-Anmeldung lief)
      // - sub-events: 50 → 90 (gleichmaessig verteilt)
      // - finalize: 90 → 100
      const subOps = childEvents.filter(ce => {
        const wasReg = sessionMeta[ce.id]?.wasRegistered;
        const isSel = selectedSessions.has(ce.id);
        return (isSel && !wasReg) || (!isSel && wasReg && !registerForOther);
      }).length;

      // 1) Haupt-Event anmelden (nur wenn Checkbox an und noch nicht angemeldet).
      // v15.25: Im subEventsOnlyMode wird die Parent-Anmeldung trotzdem
      // durchgefuehrt — als „Schatten-Registrierung" rein zur Daten-
      // Vollstaendigkeit. Damit hat jeder Teilnehmer auch im Parent-
      // Teilnehmer-Schema eine Zeile mit den Antworten auf die Hauptevent-
      // Custom-Fields (Food Preferences, Hotel, Travel etc.). Mails +
      // Outlook werden fuer diese Schatten-Anmeldung in EventContext
      // unterdrueckt — der User soll keine Bestaetigung fuers Hauptevent
      // bekommen, da er da gar nicht „teilnimmt", sondern nur fuer Sub-
      // Events.
      const isSubOnlyMode = !!(event && event.subEventsOnlyMode);
      const sessionsBeingAdded = childEvents.some(ce => selectedSessions.has(ce.id) && !sessionMeta[ce.id]?.wasRegistered);
      const parentAlreadyHasRow = !!myParentReg;
      const shouldShadowRegisterParent = isSubOnlyMode && sessionsBeingAdded && !parentAlreadyHasRow && !registerForOther;
      if (willRegisterParent || registerForOther || shouldShadowRegisterParent) {
        setSubmitProgress(30);
        setSubmitProgressLabel(locale === 'de' ? 'Haupt-Event wird angemeldet…' : 'Registering for main event…');
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
          registerForOther
            ? { proxyConsentConfirmed: true, ...(ccSelfEmail ? { extraCc: ccSelfEmail } : {}) }
            : undefined
        );
        parentOk = parentResult.ok;
        if (parentOk) {
          anySuccess = true;
          // v18.67: echten Status fuers Ergebnis-Modal merken (nicht isFull).
          setSubmittedAsWaitlist(parentResult.status === 'Warteliste');
        }
        // v19.6: Bei stellvertretender Anmeldung die personenbezogene Meldung
        // („Diese Person ist möglicherweise bereits angemeldet") statt der
        // Selbst-Meldung („du bist bereits angemeldet") zeigen.
        else setError(t(registerForOther ? 'reg.error.other' : 'reg.error'));
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
          // pro User identisch fuer alle Sub-Event-Anmeldungen.
          // v15.25: Im subEventsOnlyMode landen die Hauptevent-CF-Antworten
          // jetzt in der Schatten-Parent-Registrierung (s.o.) — die Sub-
          // Events bekommen nur ihre eigenen CFs aus dem Modal-Flow.
          const seFieldValues = { salutation, ...(sessionFieldValues[ce.id] || {}) };
          // v18.74: extraCc (übergreifende CC) + proxyConsentConfirmed (Nachweis
          // bei stellvertretender Anmeldung) zusammen in die Opts.
          // v19.6: ccSelfEmail zusätzlich in die CC der Sub-Event-Bestätigung
          // mergen (deduppt serverseitig).
          const seExtraCc = [crossCutCc, ccSelfEmail].filter(Boolean).join(';');
          const seOpts = (seExtraCc || registerForOther)
            ? { ...(seExtraCc ? { extraCc: seExtraCc } : {}), ...(registerForOther ? { proxyConsentConfirmed: true } : {}) }
            : undefined;
          const ok = (await registerForEvent(ce.id, seFieldValues, firstTrim, surnameTrim, participantEmail, sType, seOpts)).ok;
          if (ok) anySuccess = true;
          subOpsDone++;
          setSubmitProgress(50 + Math.floor((subOpsDone / Math.max(subOps, 1)) * 40));
        } else if (!isSel && wasReg && !registerForOther) {
          setSubmitProgressLabel(locale === 'de'
            ? `${childTermSingular || 'Sub-Event'} „${ce.title || '?'}" wird abgemeldet…`
            : `Cancelling ${childTermSingular || 'sub-event'} „${ce.title || '?'}"…`);
          // Cancel-Pfad bleibt aufs Selbst-Anmelden begrenzt: ein Stellvertreter
          // soll nicht aus Versehen einen Sub-Event-Slot des Anderen freigeben
          // weil er den Haken nicht gesetzt hat. Wer einen TN abmelden will,
          // macht das aktiv im Admin Center.
          await cancelRegistration(ce.id);
          anySuccess = true;
          subOpsDone++;
          setSubmitProgress(50 + Math.floor((subOpsDone / Math.max(subOps, 1)) * 40));
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
        setSubmitted(true);
      } else if (!parentOk) {
        // Parent-Fehler wurde schon in setError oben gesetzt.
      } else {
        setError(t(registerForOther ? 'reg.error.other' : 'reg.error'));
      }
    } catch {
      setError(t('reg.genericerror'));
    } finally {
      setSubmitProgress(100);
      setSubmitProgressLabel(locale === 'de' ? 'Fertig!' : 'Done!');
      // v19.6: CC-Frage-Entscheidung zuruecksetzen, damit der naechste
      // Submit-Durchlauf (z.B. naechste stellvertretende Anmeldung) wieder fragt.
      ccSelfDecidedRef.current = false;
      ccSelfRef.current = false;
      // Kleine Verzoegerung damit der User die 100%-Anzeige kurz sieht
      // bevor das Overlay wieder verschwindet.
      setTimeout(() => {
        setIsSubmitting(false);
        setSubmitProgress(0);
        setSubmitProgressLabel('');
      }, 250);
    }
  };

  const handleClear = (): void => {
    setSalutation('');
    setFirstName('');
    setSurname('');
    setEmail('');
    setEventSpecific({});
    setRegisterForOther(false);
    // v18.15: „Zurücksetzen" leert jetzt auch die Event-Section-/Sub-Event-
    // Auswahl, die Gruppen-Wahl und den Team-Modus — vorher blieb die Auswahl
    // bestehen (der Button schien „kaputt").
    setSelectedSessions(new Set());
    setSessionFieldValues({});
    setRegisterForParent(true);
    setPreferredStarterType('');
    setIsTeamMode(false);
    setTeamName('');
    setTeamMembers([]);
    setTeamMemberFields({});
    setTeamConsentConfirmed(false);
    setShowErrors(false);
    setError('');
  };

  // v18.11: Proaktive Absage („Ich nehme nicht teil"). Keine Pflichtfelder
  // nötig — der User signalisiert nur, dass er nicht kommt.
  const handleDecline = async (): Promise<void> => {
    if (!event || isDeclining) return;
    if (event.isDemoShowcase) {
      setError(locale === 'de'
        ? 'Dies ist ein Demo-Event — es wird nichts gespeichert.'
        : 'This is a demo event — nothing is stored.');
      return;
    }
    setIsDeclining(true);
    setError('');
    try {
      const ok = await declineEvent(event.id);
      if (ok) setDeclined(true);
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
        const success = (await registerForEvent(event.id, {}, r.firstName, r.lastName, r.email, undefined, { suppressMail, suppressOutlook })).ok;
        if (success) ok++; else failed.push(r.email);
      } catch { failed.push(r.email); }
    }
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
              background: `url(${event.imageUrl}) center/cover no-repeat`,
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
    const successHeadline = sessionsOnlyHint
      ? (childTermPlural
          ? (locale === 'de' ? `Für ${childTermPlural} angemeldet` : `Registered for ${childTermPlural}`)
          : (t('reg.success.sessionsonly.title') || 'Für Sessions angemeldet'))
      : (submittedAsWaitlist ? t('reg.waitlisttitle') : t('reg.success'));
    const successBody = sessionsOnlyHint
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
          : (registerForOther
              ? t('reg.successmsg.other').replace('{name}', `${firstName} ${surname}`.trim()).replace('{title}', event.title).replace('{email}', email)
              : t('reg.successmsg').replace('{title}', event.title).replace('{email}', email)));
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 720, margin: '0 auto' }}>
          {/* v15.20: Event-Foto + Organizer-Info auch auf der
              Erfolgs-Seite anzeigen — analog zur Event-Karte. */}
          {event.imageUrl && (
            <div style={{
              width: '100%', maxWidth: 480, height: 200, margin: '0 auto 24px',
              borderRadius: 'var(--dex-radius-lg)',
              background: `url(${event.imageUrl}) center/cover no-repeat`,
            }} />
          )}
          <h2 style={{ marginTop: 0 }}>{successHeadline}</h2>
          {/* v18.54: Im subEventsOnlyMode strukturierter Bestätigungstext —
              Begrüßung, Verweis auf das (nicht anwählbare) Hauptevent, Bullet-
              Liste der gewählten Sections (dynamische Organizer-Bezeichnung) und
              der Mail/Outlook-Satz NUR wenn für die gewählten Sections wirklich
              Mail bzw. Outlook aktiv ist. */}
          {sessionsOnlyHint && event.subEventsOnlyMode ? (() => {
            const selectedChildren = childEvents.filter(ce => selectedSessions.has(ce.id));
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
                  {selectedChildren.map(ce => (
                    <li key={ce.id} style={{ marginBottom: 3 }}>{ce.title || (locale === 'de' ? 'ohne Titel' : 'untitled')}</li>
                  ))}
                </ul>
                {confirmLine && <p style={{ margin: 0 }}>{confirmLine}</p>}
              </div>
            );
          })() : (
            <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>{successBody}</p>
          )}
          {(() => {
            // v18.9: Organizer-Anzeige optional ausgeblendet.
            if (event.hideOrganizer) return null;
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
                <OrganizerList names={orgs} emails={event.organizerEmails} size="md" />
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

  // v11.5: Custom-Field-Renderer extrahiert — wird zweimal verwendet:
  // einmal direkt in der Gruppen-Auswahl-Box für Felder mit
  // onlyForGroup-Constraint, einmal im Eventspez-2-Spalten-Grid für
  // alle anderen Felder. Die Filter-Logik dazu unten in den
  // groupSpecificFields- bzw. generalFields-Konstanten.
  // v13.2: fRaw jetzt typsicher als EventSpecificField (vorher any).
  const renderRegField = (fRaw: EventSpecificField, store?: Record<string, string>, setStore?: (next: Record<string, string>) => void): React.ReactElement => {
    // v18.12: optionaler Wert-Store — fuer die Custom-Fields pro Team-Mitglied.
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
    // Fallback: b2run_datenschutz ohne gespeicherte externalLinks ->
    // AGB + Datenschutz-Links zur Laufzeit injizieren, damit aeltere
    // Events ohne 'Felder reparieren' trotzdem die Links zeigen.
    if (fRaw.id === 'b2run_datenschutz' && (!fRaw.externalLinks || fRaw.externalLinks.length === 0)) {
      field = {
        ...field,
        externalLinks: [
          { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
          { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
        ],
      };
    }
    // Laufshirt/T-Shirt-Feld bei B2Run ist Pflicht (falls in alten Events
    // noch nicht so markiert)
    if ((fRaw.id === 'b2run_laufshirt' || /laufshirt/i.test(fRaw.label || '')) && !fRaw.required) {
      field = { ...field, required: true };
    }
    // v17.20: vor jedem Render die EN-Variante einziehen, sofern verfuegbar.
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
      ? <div style={{ fontSize: '0.78rem', fontWeight: 400, color: 'var(--dex-gray-500)', lineHeight: 1.45, marginTop: 2, marginBottom: 6, minHeight: '2.9em' }}>{displayHelp}</div>
      : null;
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
            gibt schoenes Hover-Popover mit der vom Organizer
            beim Event-Anlegen hinterlegten Beschreibung.
            v18.18: nur im 'tooltip'-Modus; 'inline' rendert darunter. */}
        {displayHelp && !isInlineHelp && <InfoTooltip text={displayHelp} />}
      </label>
      {inlineHelpEl}
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
        // v17.20: Anzeige-Labels fuer den EN-Modus mappen. Der gespeicherte
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
    ) : field.type === 'select' ? (
      <select className="form-select" value={vals[field.id] || ''} onChange={e => setVals({ ...vals, [field.id]: e.target.value })} style={inputStyleGreen}>
        <option value="">{tEvent('reg.pleaseselect')}</option>
        {field.options && field.options.map((opt, i) => <option key={opt} value={opt}>{pickOptionLabel(field, i, opt)}</option>)}
      </select>
    ) : field.type === 'user' || field.type === 'roommate' ? (
      // v7.17: 'roommate' nutzt denselben Picker wie 'user' — der
      // einzige Unterschied ist dass 'roommate' beim Anmelden
      // automatisch eine Zimmerpartner-Mail an die ausgewaehlte
      // Person triggert (siehe EventContext). 'user' ist der
      // generische Personen-Picker ohne Mail-Versand.
      <UserFieldPicker
        value={vals[field.id] || ''}
        onChange={v => setVals({ ...vals, [field.id]: v })}
        searchUsers={searchUsers}
        searchUserByEmail={searchUser}
        placeholder={tEvent('reg.userfield.placeholder')}
        errorStyle={showErrors && field.required && !vals[field.id]?.trim() ? errorBorder : {}}
        hint={field.type === 'roommate' ? tEvent('reg.userfield.notifyhint') : undefined}
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
          {displayHelp && !isInlineHelp && <InfoTooltip text={displayHelp} />}
        </label>
        {inlineHelpEl}
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
                confirmLabelEn-Wert; faellt sonst auf den DE-Wert. */}
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
                      alert(locale === 'de' ? 'Die Datei ist zu groß (max. 10 MB).' : 'The file is too large (max. 10 MB).');
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
    ) : (
      <input className="form-input" value={vals[field.id] || ''} onChange={e => setVals({ ...vals, [field.id]: e.target.value })} placeholder={displayLabel} style={inputStyleGreen} />
    )}
  </div>
    );
  };

  return (
    <div className="page-container">
      {showLocationBanner && (
        <div style={{
          padding: '10px 16px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
          background: 'rgba(237,139,0,0.1)', border: '1px solid var(--dex-orange)',
          color: 'var(--dex-orange)', fontSize: '0.85rem',
        }}>
          {t('reg.locationnotice')}
          {event && event.locationAudience.length > 0 && <> {t('reg.locationfilter')}: <strong>{event.locationAudience.join(', ')}</strong>.</>}
          {/* v9.17: bei Einzel-E-Mail-Whitelists in audienceFilter wuerden bei
              groesseren Verteilern (50+ Adressen) der Banner zugekleistert.
              Statt alle Mails auflisten: nur die Anzahl + die ersten 3
              Adressen zeigen, der Rest als "+N weitere". Gruppen-/Group-Namen
              (ohne "@") werden weiterhin alle aufgefuehrt — die sind kurz. */}
          {event && event.audienceFilter && event.audienceFilter.length > 0 && (() => {
            const items = event.audienceFilter;
            const emails = items.filter(s => s.includes('@'));
            const groups = items.filter(s => !s.includes('@'));
            const showLabel = (() => {
              if (emails.length === 0) return groups.join(', ');
              if (emails.length <= 3) return [...groups, ...emails].join(', ');
              const head = emails.slice(0, 3).join(', ');
              const more = emails.length - 3;
              const tail = `${head} (+${more} ${t('reg.audience.more') || 'weitere E-Mail-Adressen'})`;
              return groups.length > 0 ? `${groups.join(', ')}, ${tail}` : tail;
            })();
            return <> {t('reg.audience')}: <strong>{showLabel}</strong>.</>;
          })()}
          {event && event.filterMode === 'AND' && <> ({t('reg.andmode')})</>}
          {' '}{t('reg.yourlocation')}: {currentUser.location || t('reg.unknown')}.
        </div>
      )}
      {/* Deadline-Banner fuer Organizer/Admin: die Registrierungsfrist ist abgelaufen,
          das Formular wird aber trotzdem angezeigt (Admin/Organizer darf weiter registrieren).
          Der Ton entspricht dem Location-Banner: "als normaler User koenntest du dich nicht registrieren". */}
      {isDeadlinePassed && (isOrganizer || isAdmin) && (
        <div style={{
          padding: '10px 16px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
          background: 'rgba(237,139,0,0.1)', border: '1px solid var(--dex-orange)',
          color: 'var(--dex-orange)', fontSize: '0.85rem',
        }}>
          {t('reg.deadlinepassed.adminnotice')}
          {event && event.registrationDeadline && (
            <> {t('reg.deadlinepassed.date')}: <strong>{new Date(event.registrationDeadline).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })}</strong>.</>
          )}
        </div>
      )}
      {/* v11.33: Submit-Overlay mit Spinner + Prozent + Live-Label.
          Wird waehrend des gesamten Anmelde-Flows (Parent + alle Sub-Events
          + Bestaetigungen) eingeblendet, sodass der User auch bei langen
          Submits klares Feedback bekommt. */}
      {isSubmitting && (
        <div
          role="dialog"
          aria-modal="true"
          aria-label={locale === 'de' ? 'Anmeldung läuft' : 'Submitting registration'}
          style={{
            position: 'fixed', inset: 0, zIndex: 2000,
            background: 'rgba(0,0,0,0.55)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            padding: 16,
          }}
        >
          <div style={{
            background: '#fff', borderRadius: 12, padding: '28px 32px',
            maxWidth: 460, width: '100%',
            boxShadow: '0 16px 48px rgba(0,0,0,0.35)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16,
          }}>
            {/* v15.13: Doppel-Ladebalken entfernt — die deterministische
                Progress-Bar weiter unten (0-100% mit Label) reicht aus.
                Die zusätzliche indeterministische „Pulse"-Bar war für den
                User verwirrend. */}
            <div style={{ fontWeight: 700, fontSize: '1rem', color: 'var(--dex-gray-800)' }}>
              {locale === 'de' ? 'Anmeldung läuft …' : 'Submitting registration …'}
            </div>
            <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', textAlign: 'center', minHeight: 18 }}>
              {submitProgressLabel}
            </div>
            <div style={{ width: '100%', height: 8, borderRadius: 4, background: 'var(--dex-gray-200)', overflow: 'hidden' }}>
              <div style={{
                width: `${Math.min(100, Math.max(0, submitProgress))}%`,
                height: '100%',
                background: 'var(--dex-green, #86bc25)',
                transition: 'width 0.25s ease',
              }} />
            </div>
            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontVariantNumeric: 'tabular-nums' }}>
              {submitProgress}%
            </div>
          </div>
          <style>{`@keyframes dexProgressSlide { 0% { left: -40%; } 100% { left: 100%; } }`}</style>
        </div>
      )}
      {/* v18: Demo-Hinweis-Banner. Das Demo-Event wird ansonsten exakt wie ein
          echtes Event gerendert (so wie es in der Realität nutzbar wäre) —
          keine künstlichen Showcase-Elemente. Nur die echte Anmeldung ist
          deaktiviert. */}
      {event && event.isDemoShowcase && (
        <div style={{
          padding: '10px 16px', marginBottom: 16, borderRadius: 'var(--dex-radius, 12px)',
          background: 'rgba(0,118,168,0.08)', border: '1px solid var(--dex-blue, #0076a8)',
          color: 'var(--dex-gray-800)', fontSize: '0.85rem', lineHeight: 1.55,
          display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap',
        }}>
          <span style={{
            padding: '2px 8px', borderRadius: 999, background: 'var(--dex-blue, #0076a8)',
            color: '#fff', fontSize: '0.66rem', fontWeight: 700, letterSpacing: 1,
          }}>DEMO</span>
          {locale === 'de'
            ? <span>Dies ist ein <strong>Demo-Event</strong> — es wird genau so angezeigt wie ein echtes Event. Du kannst die Anmeldemaske ansehen, aber <strong>keine echte Anmeldung</strong> absenden.</span>
            : <span>This is a <strong>demo event</strong> — shown exactly like a real one. You can explore the registration form, but <strong>cannot submit a real registration</strong>.</span>}
        </div>
      )}
      <div className="registration-layout">
        {/* Event-Info links */}
        <div className="registration-event">
          <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
            <Icon iconName="Calendar" style={{ fontSize: 16 }} />
            {t('reg.selectedevent')}
          </div>
          <div
            className="registration-event__card"
            style={{
              display: 'flex',
              // Hochkant -> Bild links + Inhalt rechts | Querformat -> Bild oben + Inhalt drunter
              flexDirection: imgOrientation === 'portrait' ? 'row' : 'column',
              gap: 12,
              alignItems: 'stretch',
            }}
          >
            <div
              className="registration-event__image"
              style={{
                position: 'relative',
                // v11.91: Hintergrund auf Weiß gesetzt — PNGs mit Transparenz
                // zeigten vorher den hellgrauen Hintergrund durch, was wie ein
                // unsauberer „grauer Rand" um Logos aussah.
                background: event.imageUrl
                  ? '#fff'
                  : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
                borderRadius: 'var(--dex-radius)',
                overflow: 'hidden',
                // Hochkant: schmal links + volle Karten-Hoehe
                // Querformat: volle Breite oben, Hoehe richtet sich nach Bild-Aspect (kein Crop)
                ...(imgOrientation === 'portrait'
                  ? { flex: '0 0 220px', alignSelf: 'stretch', minHeight: 360, display: 'flex' }
                  : { width: '100%', display: 'flex', justifyContent: 'center' }),
              }}
            >
              {event.imageUrl && (
                <img
                  src={event.imageUrl}
                  alt={event.title}
                  // v11.56: Auch im Portrait-Modus 'contain', damit das Bild
                  // vollstaendig sichtbar bleibt (vorher wurde der obere
                  // Bildteil weggecroppt). Hintergrundfarbe der Hülle ist
                  // bereits gray-100 — das ergibt einen sauberen, neutralen
                  // Letterbox-Rahmen statt eines Bildschnitts.
                  style={imgOrientation === 'portrait'
                    ? { width: '100%', height: '100%', objectFit: 'contain', display: 'block' }
                    : { width: '100%', height: 'auto', maxHeight: 480, objectFit: 'contain', display: 'block' }
                  }
                />
              )}
              {/* v11.91: Info-Button entfernt — die Beschreibung ist jetzt
                  immer ausgeklappt, kein Toggle mehr nötig. */}
            </div>
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 10, padding: '4px 4px 4px 0' }}>
              <h4 style={{ fontSize: '1rem', margin: 0 }}>{event.title}</h4>
              {/* v11.91: Datum + Ort als prominente Badges mit Icon.
                  v11.93: Datum einzeilig (nowrap) — Box wächst auf
                  natürliche Breite. Der Ort-Kasten streckt sich auf
                  dieselbe Breite, damit beide Boxen visuell aligniert
                  sind. inline-flex + alignItems:stretch sorgt für gleiche
                  Breite ohne festen Wert. */}
              {/* v11.94: alignSelf:stretch + maxWidth:100% damit die Box
                  nicht über den Card-Rand rausragt; gleichzeitig wächst
                  sie auf die natürliche Breite des längeren Inhalts und
                  beide Boxen sind gleich breit. */}
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'stretch', gap: 6, maxWidth: '100%' }}>
                <div style={{
                  display: 'flex', alignItems: 'center', gap: 8,
                  padding: '8px 12px', borderRadius: 8,
                  background: 'rgba(134,188,37,0.10)', color: 'var(--dex-green-dark, #4a7c1f)',
                  fontSize: '0.88rem', fontWeight: 600,
                }}>
                  <Icon iconName="Calendar" style={{ fontSize: 16, flexShrink: 0 }} />
                  <span>
                    {/* v11.94: kompaktes Range-Format („–" statt „until"),
                        bei gleichem Tag nur einmal Datum + „HH:MM - HH:MM". */}
                    {formatDateRange(event.startDate, event.endDate)}
                  </span>
                </div>
                {(event.location || (event.locationAddress && (event.locationAddress.street || event.locationAddress.city))) && (
                  <div style={{
                    display: 'flex', alignItems: 'flex-start', gap: 8,
                    padding: '8px 12px', borderRadius: 8,
                    background: 'rgba(0,86,166,0.08)', color: '#0a3766',
                    fontSize: '0.88rem',
                  }}>
                    <Icon iconName="POI" style={{ fontSize: 16, marginTop: 2, flexShrink: 0 }} />
                    <span>
                      {event.location && (
                        <span style={{ fontWeight: 700 }}>{event.location}</span>
                      )}
                      {event.locationAddress && (event.locationAddress.street || event.locationAddress.city) && (
                        <>
                          <br />
                          <span style={{ fontWeight: 400 }}>
                            {[event.locationAddress.street, event.locationAddress.houseNo].filter(Boolean).join(' ')}
                            {(event.locationAddress.zip || event.locationAddress.city) && <br />}
                            {[event.locationAddress.zip, event.locationAddress.city].filter(Boolean).join(' ')}
                          </span>
                        </>
                      )}
                    </span>
                  </div>
                )}
              </div>
              {(() => {
                // v18.9: Organizer-Anzeige optional ausgeblendet.
                if (event.hideOrganizer) return null;
                // Organizer als Chips mit Foto (Hover-Enlarge). Namen werden von "Nachname, Vorname"
                // in "Vorname Nachname" normalisiert. v11.91: Label + Chip größer für bessere Lesbarkeit.
                const orgs = event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => {
                  const trimmed = o.trim();
                  const parts = trimmed.split(',').map(s => s.trim());
                  return parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
                }).filter(Boolean);
                if (orgs.length === 0) return null;
                return (
                  <div style={{ marginTop: 6 }}>
                    <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 6, fontWeight: 600 }}>Organizer</div>
                    <OrganizerList names={orgs} emails={event.organizerEmails} size="md" />
                  </div>
                );
              })()}
              {/* v10.16: Optionaler Ansprechpartner — frei eingegebene Person
                  außerhalb des App-User-Pools. Reines Anzeige-Feld; Mailto-Link
                  wenn Email gesetzt. Wird nur gerendert wenn mindestens Name
                  oder Email gepflegt sind. */}
              {(event.contactName || event.contactEmail || event.contactInfo) && (
                <div style={{ marginTop: 12, padding: '10px 12px', background: 'var(--dex-gray-50, #f7f7f7)', borderRadius: 8, border: '1px solid var(--dex-gray-200)' }}>
                  {/* v11.97: Gleiche Schriftgröße + Gewicht wie das
                      ORGANIZER-Label darüber. */}
                  <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 6, fontWeight: 600 }}>
                    {locale === 'de' ? 'Ansprechpartner' : 'Contact'}
                  </div>
                  {event.contactName && (
                    <div style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>{event.contactName}</div>
                  )}
                  {event.contactEmail && (
                    <div style={{ fontSize: '0.78rem', marginTop: 2 }}>
                      <a href={`mailto:${event.contactEmail}`} style={{ color: 'var(--dex-green, #86bc25)', textDecoration: 'none' }}>{event.contactEmail}</a>
                    </div>
                  )}
                  {event.contactInfo && (
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-700)', marginTop: 4, whiteSpace: 'pre-wrap', lineHeight: 1.45 }}>{event.contactInfo}</div>
                  )}
                </div>
              )}
              {/* Verfuegbare Plaetze anzeigen — nur wenn es eine Obergrenze gibt.
                  Gilt sowohl fuer normale Deloitte-Events als auch fuer B2Run
                  (dort ist maxParticipants = durchCap + funCap). */}
              {event.maxParticipants > 0 && (() => {
                const free = Math.max(0, event.maxParticipants - (event.currentParticipants || 0));
                const isFullAll = free <= 0;
                const nearlyFull = !isFullAll && free <= Math.max(1, Math.round(event.maxParticipants * 0.1));
                const color = isFullAll
                  ? 'var(--dex-red, #c00)'
                  : nearlyFull
                    ? 'var(--dex-orange, #ff8c00)'
                    : 'var(--dex-green-dark, #6b9a1e)';
                // v11.93: "Event voll"-Hinweis nicht mehr hier rendern —
                // der ausführliche "All places are taken. Currently X on
                // waiting list"-Block unter der Beschreibung übernimmt das
                // bereits, sonst doppelte Meldung.
                if (isFullAll) return null;
                // v16.2: Bei Team-Event Anzahl freier Teams zusaetzlich
                // ausgeben (Plaetze / TeamSize), damit die Quoten-Anzeige
                // bei z.B. „80 / 80" auch klar macht „= 20 Teams frei".
                const isTeamEvent = !!(event.teamRegistrationEnabled && event.teamSize && event.teamSize > 1);
                const teamsFree = isTeamEvent ? Math.floor(free / (event.teamSize || 1)) : 0;
                return (
                  <div style={{
                    fontSize: '0.78rem',
                    color,
                    marginTop: 6,
                    fontWeight: 600,
                  }}>
                    {`${free} / ${event.maxParticipants} ${t('reg.seats.available') || 'Plätze frei'}`}
                    {isTeamEvent && (
                      <span style={{ marginLeft: 6, color: 'var(--dex-gray-600)', fontWeight: 500 }}>
                        ({teamsFree} {teamsFree === 1 ? (locale === 'de' ? 'Team' : 'team') : (locale === 'de' ? 'Teams' : 'teams')} {locale === 'de' ? 'frei' : 'free'})
                      </span>
                    )}
                  </div>
                );
              })()}
            </div>
          </div>
          {/* v11.91: Beschreibung immer ausgeklappt — kein Toggle mehr. */}
          {event.description && (
            // v9.25: Beschreibung darf HTML enthalten (RichText-Editor im
            // EventCreation/Edit). Wir rendern als HTML statt Plain-Text,
            // damit Formatierung wie Listen, Links, Fett etc. funktioniert.
            // Die Description kommt aus dem eigenen Tenant — sicherer Origin.
            <div
              className="dex-event-desc"
              style={{
                padding: '12px 16px', color: 'var(--dex-gray-700)',
                background: 'var(--dex-gray-50)', borderRadius: '0 0 var(--dex-radius) var(--dex-radius)',
                borderTop: '1px solid var(--dex-gray-200)',
                wordBreak: 'break-word',
              }}
              dangerouslySetInnerHTML={{
                __html: (() => {
                  // v11.91: Email-Adressen in der Beschreibung automatisch
                  // in mailto-Links umwandeln. Funktioniert sowohl für
                  // Plain-Text als auch für HTML — Emails in bereits
                  // verlinktem Text (innerhalb von href="...") werden
                  // übersprungen.
                  const raw = event.description || '';
                  const isHtml = /<[a-z][\s\S]*>/i.test(raw);
                  const base = isHtml
                    ? raw
                    : raw.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
                  const EMAIL_RE = /\b([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})\b/g;
                  // Skip emails that already sit inside href="..." (already linked)
                  return base.replace(EMAIL_RE, (match, _email, offset, full) => {
                    const beforeWindow = full.slice(Math.max(0, offset - 80), offset);
                    if (/href\s*=\s*["'][^"']*$/.test(beforeWindow)) return match;
                    if (/>$/.test(beforeWindow) && /<a [^>]*$/i.test(beforeWindow)) return match;
                    // v11.95: Deloitte-Grün (--dex-green #86bc25) statt dem
                    // dunkleren Olive — gleiche Farbe wie die Card-Header.
                    return `<a href="mailto:${match}" style="color:#86bc25;text-decoration:underline">${match}</a>`;
                  });
                })(),
              }}
            />
          )}
          {isFull && (
            <p className="text-red text-center mt-8" style={{ padding: '0 12px 12px', fontWeight: 600, fontSize: '0.85rem' }}>
              {t('reg.allplaces')
                .replace('{count}', String(event.waitlistCount))
                .replace('{personLabel}', event.waitlistCount === 1
                  ? (locale === 'en' ? 'person' : 'Person')
                  : (locale === 'en' ? 'people' : 'Personen'))}
            </p>
          )}
          {/* v10.20: Sessions-/Hauptevent-Auswahl ist nun in die rechte Spalte
              ('registration-specific') eingebettet — siehe weiter unten unter
              dem section-header "reg.eventinfo". Vorher stand der Block hier
              in der linken Event-Karten-Spalte; das Layout war optisch
              ungleich (links wuchs unbegrenzt, rechts gar nichts) und der
              User musste zwischen Spalten hin- und her-springen. Jetzt sind
              alle Anmelde-Inputs (Sessions, Starter-Typ, Custom-Fields) in
              einer Spalte gebuendelt. */}
          {false && (
            <div style={{ marginTop: 16, border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 16 }}>
              <h4 style={{ marginTop: 0, marginBottom: 4, fontSize: '0.95rem' }}>{t('reg.selection.title') || 'Wofür möchtest du dich anmelden?'}</h4>
              <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                {t('reg.selection.hint') || 'Haupt-Event und Sessions können unabhängig voneinander an- oder abgewählt werden.'}
              </p>

              {/* v15.3.1: Haupt-Event-Checkbox nur zeigen wenn Anmeldung
                  fürs Hauptevent überhaupt möglich ist. Im subEventsOnlyMode
                  läuft die Anmeldung exklusiv über Sub-Events. */}
              {!event.subEventsOnlyMode && (
              <label style={{
                display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
                borderRadius: 8,
                border: `1px solid ${registerForParent && !parentAlreadyRegistered ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                background: registerForParent && !parentAlreadyRegistered ? 'rgba(134,188,37,0.06)' : '#fff',
                cursor: parentAlreadyRegistered ? 'default' : 'pointer',
              }}>
                <input
                  type="checkbox"
                  checked={parentAlreadyRegistered ? true : registerForParent}
                  disabled={parentAlreadyRegistered}
                  onChange={e => setRegisterForParent(e.target.checked)}
                  style={{ marginTop: 2 }}
                />
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700 }}>{t('reg.selection.mainevent') || 'Haupt-Event'}: {event.title}</div>
                  {parentAlreadyRegistered && (
                    <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                      {t('reg.selection.alreadyregistered') || 'Du bist bereits für das Haupt-Event angemeldet.'}
                    </div>
                  )}
                </div>
              </label>
              )}

              {/* Sessions */}
              {childEvents.length > 0 && (
                <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', fontWeight: 600 }}>{childTermPlural || t('reg.selection.sessions') || 'Sessions'}</div>
                  {/* v14.5: Pflicht-Hinweis wenn Organizer requireSubEventSelection aktiv hat. */}
                  {event && event.requireSubEventSelection && (
                    <div style={{
                      padding: '8px 12px', borderRadius: 8,
                      background: 'rgba(237,139,0,0.10)',
                      border: '1px solid var(--dex-orange, #ed8b00)',
                      fontSize: '0.82rem', color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600,
                    }}>
                      {locale === 'de'
                        ? `Pflicht: bitte mindestens ein ${childTermSingular || 'Sub-Event'} auswählen.`
                        : `Required: please pick at least one ${childTermSingular || 'sub-event'}.`}
                    </div>
                  )}
                  {childEvents.map(ce => {
                    const meta = sessionMeta[ce.id] || { count: 0, wasRegistered: false };
                    const isSel = selectedSessions.has(ce.id);
                    const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
                    const isSessionFull = hasCap && meta.count >= (ce.maxParticipants || 0);
                    const deadlinePassed = !!(ce.registrationDeadline && new Date(ce.registrationDeadline) < new Date());
                    const disabled = (isSessionFull && !isSel) || (deadlinePassed && !isSel);
                    // Erbt vom Haupt-Event wenn gleichzeitig angemeldet wird.
                    const inheritsStarter = isSplitGroup && (willRegisterParent || registerForOther);
                    const sType = sessionStarterType[ce.id] || '';

                    return (
                      <div key={ce.id} style={{
                        padding: 10, borderRadius: 8,
                        border: `1px solid ${isSel ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                        background: isSel ? 'rgba(134,188,37,0.06)' : '#fff',
                      }}>
                        <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: disabled ? 'not-allowed' : 'pointer' }}>
                          <input
                            type="checkbox"
                            checked={isSel}
                            disabled={disabled}
                            onChange={e => {
                              if (e.target.checked) {
                                // Wenn das Sub-Event eigene Custom-Fields hat:
                                // erst das Modal öffnen, der User muss die Antworten
                                // bestätigen. Bei „Bestätigen" landet die Session in
                                // selectedSessions + die Werte in sessionFieldValues.
                                // Sub-Events ohne Custom-Fields: direkt selektieren.
                                const hasCustomFields = (ce.eventSpecificFields || []).length > 0;
                                if (hasCustomFields) {
                                  setPendingSubEventModal({
                                    subEventId: ce.id,
                                    draftValues: { ...(sessionFieldValues[ce.id] || {}) },
                                  });
                                } else {
                                  const next = new Set(selectedSessions);
                                  next.add(ce.id);
                                  setSelectedSessions(next);
                                }
                              } else {
                                // Uncheck: Session entfernen + gespeicherte Field-Werte
                                // wegräumen damit beim erneuten Checken ein frisches
                                // Modal kommt (kein Stale-State).
                                const next = new Set(selectedSessions);
                                next.delete(ce.id);
                                setSelectedSessions(next);
                                setSessionFieldValues(prev => {
                                  const copy = { ...prev };
                                  delete copy[ce.id];
                                  return copy;
                                });
                              }
                            }}
                            style={{ marginTop: 2 }}
                          />
                          <div style={{ flex: 1 }}>
                            <div style={{ fontWeight: 700 }}>{ce.title || tEvent('reg.subevents.untitled')}</div>
                            {ce.description && (
                              // v11.97: gleiche Schriftgröße wie der Titel
                              // (Standard-Body). Vorher 0.78rem klein.
                              <div style={{ color: 'var(--dex-gray-600)', marginTop: 2 }}>{ce.description}</div>
                            )}
                            {/* v11.94: Datum + Ort mit Icons (analog zum
                                Haupt-Event-Header), damit Sub-Events visuell
                                konsistent sind und auf einen Blick lesbar. */}
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, marginTop: 4, color: 'var(--dex-gray-600)' }}>
                              {ce.startDate && (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                  {/* v11.97: Icon-Größe an Standard-Body angepasst. */}
                                  <Icon iconName="Calendar" style={{ fontSize: 15, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                                  {formatDate(ce.startDate)}
                                </span>
                              )}
                              {ce.location && (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                  <Icon iconName="POI" style={{ fontSize: 15, color: '#0a3766' }} />
                                  {ce.location}
                                </span>
                              )}
                            </div>
                            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {hasCap && (() => {
                                // v9.8: Klartext-Anzeige damit der User auf einen Blick
                                // sieht, wie viele Plaetze noch frei sind. Vorher stand
                                // dort nur "0/25" ohne Label, was die User-Frage
                                // "warum steht da 0/25?" ausgeloest hat.
                                const sessionFree = Math.max(0, (ce.maxParticipants || 0) - (meta.count || 0));
                                return (
                                  <> · <span style={{ color: isSessionFull ? 'var(--dex-red)' : 'inherit', fontWeight: 600 }}>
                                    {/* v19.19: belegt-Zahl bei der Kapazität deckeln,
                                        damit eine Überbuchung (count > cap) auf der
                                        Anmeldeseite NICHT sichtbar wird — die echte
                                        Überbuchungszahl sieht nur der Organizer/Admin. */}
                                    {Math.min(meta.count, ce.maxParticipants || 0)}/{ce.maxParticipants} {t('reg.subevents.taken')}
                                  </span>
                                  {!isSessionFull && (
                                    <span style={{ color: 'var(--dex-green-dark)' }}> — {sessionFree} {t('reg.free')}</span>
                                  )}
                                  </>
                                );
                              })()}
                            </div>
                            {deadlinePassed && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-orange)', marginTop: 2 }}>
                                {t('reg.subevents.deadlinepassed')}
                              </div>
                            )}
                            {isSessionFull && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-red)', marginTop: 2 }}>
                                {t('reg.subevents.sessionfull')}
                              </div>
                            )}
                            {/* v10.20: Gruppen-Auswahl pro Session — nur wenn NICHT vom Parent geerbt.
                                Dynamische Labels splitLabelA / splitLabelB. */}
                            {isSel && isSplitGroup && !inheritsStarter && (
                              <div style={{ marginTop: 8, display: 'flex', gap: 10, fontSize: '0.8rem' }}>
                                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
                                  <input
                                    type="radio"
                                    name={`starter-${ce.id}`}
                                    checked={sType === 'Durchstarter'}
                                    onChange={() => setSessionStarterType({ ...sessionStarterType, [ce.id]: 'Durchstarter' })}
                                  />
                                  {splitLabelA}
                                </label>
                                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
                                  <input
                                    type="radio"
                                    name={`starter-${ce.id}`}
                                    checked={sType === 'Funstarter'}
                                    onChange={() => setSessionStarterType({ ...sessionStarterType, [ce.id]: 'Funstarter' })}
                                  />
                                  {splitLabelB}
                                </label>
                              </div>
                            )}
                            {/* v10.25: B2Run-spezifischer Vererbungs-Hinweis
                                entfernt — bei generischer Split-Capacity
                                (z.B. "Vormittag/Nachmittag") wirkt der
                                Hinweis "Starter-Typ wird vom Haupt-Event
                                übernommen" verwirrend. Die Logik bleibt
                                (Sub-Event übernimmt die Gruppen-Wahl des
                                Parents), nur die explizite UI-Zeile ist
                                weg. */}
                          </div>
                        </label>
                      </div>
                    );
                  })}
                </div>
              )}

              {isSessionsOnlyMode && selectedSessions.size > 0 && !event.subEventsOnlyMode && (
                <div style={{
                  marginTop: 12, padding: '8px 10px', borderRadius: 6,
                  background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                  color: 'var(--dex-orange)', fontSize: '0.78rem',
                }}>
                  {childTermPlural
                    ? (locale === 'de'
                        ? `Du meldest dich ausschließlich für ${childTermPlural} an — NICHT für das Haupt-Event.`
                        : `You are registering exclusively for ${childTermPlural} — NOT for the main event.`)
                    : (t('reg.selection.sessionsonlyhint') || 'Du meldest dich ausschließlich für Sessions an — NICHT für das Haupt-Event.')}
                </div>
              )}
            </div>
          )}
        </div>

        {/* v18.73: Die „Offene Teams"-Box ist nach UNTEN gewandert (unter die
            „Ich melde mich + mein Team an"-Karte) — siehe weiter unten. Oben
            steht jetzt immer zuerst die persönliche Daten-Karte, dann die
            event-spezifischen Infos. */}

        {/* Persoenliche Daten */}
        <div className="registration-form">
          {/* v11.97: Section-Header + Register-for-other-Toggle in einer
              Zeile (grünes Section-Header-Pill links, Toggle als Link
              rechts daneben). Vorher saß der Toggle unter dem Header
              im Body — wenig auffällig. „* = Required field"-Legende
              ist hier weg und sitzt jetzt am Event-Specific-Header. */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
            <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Icon iconName="ContactInfo" style={{ fontSize: 16 }} />
              {t('reg.personalinfo')}
            </div>
            {canRegisterForOther && (
              <button
                type="button"
                onClick={() => {
                  setRegisterForOther(!registerForOther);
                  setThirdPartyCheck(null);
                  setPickedUserProfile(null);
                  setOtherConsentConfirmed(false);
                  setExternalPerson(false); // v18.74: Extern-Modus beim Wechsel zurücksetzen
                  // v19.6: CC-Frage-Entscheidung beim Moduswechsel zuruecksetzen.
                  ccSelfDecidedRef.current = false;
                  ccSelfRef.current = false;
                  if (!registerForOther) { setFirstName(''); setSurname(''); setEmail(''); setUserSearch(''); setUserResults([]); }
                  else { setFirstName(currentUser.firstName); setSurname(currentUser.surname); setEmail(currentUser.email); setUserSearch(''); setUserResults([]); }
                }}
                style={{
                  background: 'none', border: 'none', padding: '4px 12px',
                  color: 'var(--dex-green-dark)', fontSize: '0.85rem',
                  textDecoration: 'underline', cursor: 'pointer',
                  fontWeight: 600,
                }}
              >
                {registerForOther ? t('reg.registerself') : t('reg.registerother')}
              </button>
            )}
            {/* v18.13: Massenimport — nur Organizer/Admin im „Für andere"-Modus. */}
            {registerForOther && canCreateEvents && (
              <button
                type="button"
                onClick={() => { setMassImportResult(null); setMassImportRows([]); setMassImportStep('input'); setMassImportOpen(true); }}
                style={{
                  background: 'none', border: 'none', padding: '4px 12px',
                  color: 'var(--dex-blue, #0076a8)', fontSize: '0.85rem',
                  textDecoration: 'underline', cursor: 'pointer', fontWeight: 600,
                }}
              >
                {locale === 'de' ? 'Massenimport' : 'Bulk import'}
              </button>
            )}
          </div>
          <div style={{ padding: '24px 20px' }}>
            {canRegisterForOther && (
              <>
                {registerForOther && isAssistant && !canCreateEvents && (
                  <div style={{
                    padding: '8px 12px', marginBottom: 12, borderRadius: 'var(--dex-radius-md)',
                    background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                    color: 'var(--dex-orange)', fontSize: '0.8rem',
                  }}>
                    As an Assistant you can only register <strong>Partners</strong> or <strong>Directors</strong> for this event.
                  </div>
                )}
                {/* v18.74: Umschalter „Person außerhalb Deloitte" — nur für
                    Organizer/Admins (nicht für Assistenten, die auf interne
                    Partner/Directors beschränkt sind). Blendet den People-Picker
                    aus und macht Vorname/Nachname/E-Mail frei eintragbar. */}
                {registerForOther && canCreateEvents && (
                  <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 16, cursor: 'pointer', fontSize: '0.85rem', color: 'var(--dex-gray-800)' }}>
                    <input
                      type="checkbox"
                      checked={externalPerson}
                      onChange={e => {
                        const on = e.target.checked;
                        setExternalPerson(on);
                        // Picker- und Feld-State frisch zurücksetzen (manuelle Eingabe).
                        setUserSearch(''); setUserResults([]); setPickedUserProfile(null);
                        setThirdPartyCheck(null); setOtherConsentConfirmed(false);
                        setFirstName(''); setSurname(''); setEmail('');
                      }}
                    />
                    <span>
                      {locale === 'de'
                        ? <>Person <strong>außerhalb Deloitte</strong> anmelden (externe E-Mail-Adresse)</>
                        : <>Register a person <strong>outside Deloitte</strong> (external email address)</>}
                    </span>
                  </label>
                )}
                {registerForOther && !externalPerson && (
                  <div className="form-group" style={{ position: 'relative', marginBottom: 20 }}>
                    {/* v11.97: Label entfernt — Suche ist selbsterklärend (Placeholder). */}
                    <input
                      className="form-input"
                      value={userSearch}
                      onChange={e => {
                        const val = e.target.value;
                        setUserSearch(val);
                        if (searchTimerRef.current) clearTimeout(searchTimerRef.current);
                        if (val.length >= 2) {
                          searchTimerRef.current = setTimeout(async () => {
                            setIsSearchingUser(true);
                            const results = await searchUsers(val, userSearchIncludeIntl);
                            setUserResults(results);
                            setIsSearchingUser(false);
                          }, 300);
                        } else {
                          setUserResults([]);
                        }
                      }}
                      placeholder={t('reg.searchplaceholder') || 'Name oder E-Mail eingeben...'}
                    />
                    <InternationalSearchToggle
                      checked={userSearchIncludeIntl}
                      onChange={async next => {
                        setUserSearchIncludeIntl(next);
                        const val = userSearch.trim();
                        if (val.length >= 2) {
                          setIsSearchingUser(true);
                          try {
                            const results = await searchUsers(val, next);
                            setUserResults(results);
                          } catch { /* */ }
                          setIsSearchingUser(false);
                        }
                      }}
                      isDe={locale === 'de'}
                    />
                    {isSearchingUser && (
                      <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>Suche...</div>
                    )}
                    {userResults.length > 0 && (
                      <div style={{
                        position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
                        background: '#fff', border: '1px solid var(--dex-gray-200)',
                        borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                        maxHeight: 200, overflowY: 'auto',
                      }}>
                        {userResults.map(u => {
                          // "Nachname, Vorname" oder "Vorname Nachname" Format
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
                          // Assistant-Einschraenkung: User darf nur Partner/Director auswaehlen,
                          // andere Treffer werden grau + nicht-klickbar angezeigt.
                          const assistantOnly = isAssistant && !canCreateEvents;
                          const targetAllowed = !assistantOnly || isAllowedTargetForAssistant(u.jobTitle);
                          return (
                            <div
                              key={u.email}
                              style={{
                                padding: '8px 12px', cursor: targetAllowed ? 'pointer' : 'not-allowed', fontSize: '0.85rem',
                                borderBottom: '1px solid var(--dex-gray-100)',
                                opacity: targetAllowed ? 1 : 0.45,
                              }}
                              onMouseEnter={e => { if (targetAllowed) (e.currentTarget as HTMLElement).style.background = 'var(--dex-gray-50)'; }}
                              onMouseLeave={e => { (e.currentTarget as HTMLElement).style.background = '#fff'; }}
                              onMouseDown={() => {
                                if (!targetAllowed) return;
                                setFirstName(uFirstName);
                                setSurname(uSurname);
                                setEmail(u.email);
                                setUserSearch(u.displayName);
                                setUserResults([]);
                                // v11.97: Profil mit Department + Mobile nachladen
                                // — searchUsers liefert nur jobTitle+location.
                                setPickedUserProfile({
                                  jobTitle: u.jobTitle || '',
                                  location: u.location || '',
                                });
                                searchUser(u.email).then(p => {
                                  if (p) {
                                    setPickedUserProfile({
                                      jobTitle: p.jobTitle || u.jobTitle || '',
                                      department: p.department || '',
                                      location: p.location || u.location || '',
                                      mobilePhone: p.mobilePhone || '',
                                    });
                                  }
                                }).catch(() => { /* silent */ });
                                // Frueh-Check: bereits angemeldet? Im Verteiler?
                                setThirdPartyCheck(null);
                                if (event) {
                                  (async () => {
                                    const existing = await checkRegistrationByEmail(event.id, u.email).catch(() => null);
                                    const alreadyRegistered = !!existing && existing.Status !== 'Abgemeldet';
                                    const locFilters = event.locationAudience || [];
                                    const audFilters = (event.audienceFilter || [])
                                      .map(s => s.trim())
                                      .filter(s => s && s.toLowerCase() !== 'all' && s.toLowerCase() !== 'deall');
                                    const hasAnyFilter = locFilters.length > 0 || audFilters.length > 0;
                                    let notInAudience = false;
                                    if (hasAnyFilter) {
                                      const loc = (u.location || '').toLowerCase();
                                      const uEmail = u.email.toLowerCase();
                                      const locMatch = locFilters.length === 0 || locFilters.some(f => {
                                        const fl = f.trim().toLowerCase();
                                        if (fl === 'all') return true;
                                        const norm = fl.replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ä/g, 'ae');
                                        return loc.indexOf(fl) >= 0 || loc.indexOf(norm) >= 0;
                                      });
                                      const audMatch = audFilters.length === 0 || audFilters.some(f => {
                                        const fl = f.trim().toLowerCase();
                                        if (fl.indexOf('@') >= 0) return uEmail === fl;
                                        if (fl.startsWith('de')) {
                                          const city = fl.substring(2);
                                          const norm = city.replace(/oe/g, 'ö').replace(/ue/g, 'ü').replace(/ae/g, 'ä');
                                          return loc.indexOf(city) >= 0 || loc.indexOf(norm) >= 0;
                                        }
                                        return false;
                                      });
                                      const mode = event.filterMode || 'AND';
                                      const hasLoc = locFilters.length > 0;
                                      const hasAud = audFilters.length > 0;
                                      let visible: boolean;
                                      if (mode === 'OR') {
                                        if (hasLoc && hasAud) visible = locMatch || audMatch;
                                        else if (hasLoc) visible = locMatch;
                                        else visible = audMatch;
                                      } else {
                                        if (hasLoc && hasAud) visible = locMatch && audMatch;
                                        else if (hasLoc) visible = locMatch;
                                        else visible = audMatch;
                                      }
                                      notInAudience = !visible;
                                    }
                                    setThirdPartyCheck({
                                      alreadyRegistered,
                                      notInAudience,
                                      // v19.8: Name + Anmeldedatum der bestehenden
                                      // Registrierung fuer eine konkrete Hinweis-Box.
                                      registeredName: (existing && (existing.ParticipantName || `${existing.Vorname || ''} ${existing.Nachname || ''}`.trim())) || u.displayName || '',
                                      registeredDate: (existing && existing.RegistrationDate) || '',
                                    });
                                  })();
                                }
                              }}
                              title={targetAllowed ? '' : 'Assistants can only register Partners or Directors for events.'}
                            >
                              {/* v11.3: People-Picker-Reihe mit Foto — analog
                                  zum Wizard-Organizer-Picker. SP-userphoto.aspx
                                  liefert das Profilbild zum E-Mail-Account. */}
                              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                                <img
                                  src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=S`}
                                  alt={u.displayName}
                                  onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                  style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                                />
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                                  <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                    {u.email}
                                    {u.jobTitle ? ` · ${u.jobTitle}` : ''}
                                    {u.location ? ` · ${u.location}` : ''}
                                  </div>
                                </div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                )}
                {/* v15.16: Pflicht-Bestätigungs-Box bei „Für andere
                    registrieren" — analog zur Team-Anmeldung. Erscheint
                    sobald eine Person ausgewählt ist (E-Mail gefüllt) und
                    enthält zusätzlich einen Spezial-Hinweis für externe
                    Adressen (kein Outlook-Termin, Mail zur Weiterleitung
                    an die Organizer). */}
                {registerForOther && (() => {
                  // v15.22: Hinweis-Box bereits anzeigen, sobald „Für andere
                  // registrieren" aktiv ist — nicht erst wenn die E-Mail
                  // gefüllt ist. Der User soll sofort sehen, dass eine
                  // Zustimmung nötig ist.
                  // v18.74: Extern, sobald der Extern-Modus aktiv ist ODER die
                  // eingegebene E-Mail keine Deloitte-DE-Adresse ist.
                  const isExternal = externalPerson || (!!email.trim() && !/@(.*\.)?deloitte\.de$/i.test(email.trim()));
                  const pickedName = `${firstName} ${surname}`.trim();
                  return (
                    <div style={{
                      marginBottom: 16,
                      padding: '14px 16px',
                      background: 'rgba(237,139,0,0.10)',
                      border: '2px solid var(--dex-orange, #ed8b00)',
                      borderRadius: 'var(--dex-radius-md)',
                      color: '#7a4a00',
                      fontSize: '0.88rem',
                      lineHeight: 1.55,
                    }}>
                      <div style={{ fontWeight: 700, marginBottom: 8, fontSize: '0.95rem' }}>
                        {locale === 'de'
                          ? (isExternal ? 'Schriftliche Zustimmung der Person einholen' : 'Vorab die Zustimmung der Person einholen')
                          : (isExternal ? 'Get the person\'s written consent' : 'Get the person\'s consent up front')}
                      </div>
                      {externalPerson && (
                        <div style={{ marginBottom: 8, fontWeight: 600 }}>
                          {locale === 'de'
                            ? <>Trage <strong>Vorname, Nachname und E-Mail-Adresse</strong> der externen Person unten selbst ein.</>
                            : <>Enter the external person&rsquo;s <strong>first name, last name and email address</strong> in the fields below.</>}
                        </div>
                      )}
                      <div style={{ marginBottom: 8 }}>
                        {locale === 'de'
                          ? <>Mit dem Absenden meldest du {pickedName ? <><strong>{pickedName}</strong></> : <>die {externalPerson ? 'externe ' : 'ausgewählte '}Person</>} stellvertretend an. Bitte stelle sicher, dass die Person ihrer Anmeldung <strong>VORHER {isExternal ? 'schriftlich ' : ''}zugestimmt</strong> hat — eine Anmeldung ohne Einverständnis ist nicht erlaubt.</>
                          : <>By submitting you register {pickedName ? <><strong>{pickedName}</strong></> : <>the {externalPerson ? 'external ' : 'selected '}person</>} on their behalf. Please make sure the person has <strong>{isExternal ? 'consented in writing' : 'consented up front'}</strong> — registering people without their consent is not allowed.</>}
                      </div>
                      {isExternal && (
                        <div style={{
                          marginTop: 10, padding: '10px 12px',
                          background: '#fff', border: '1px dashed var(--dex-orange, #ed8b00)',
                          borderRadius: 6,
                        }}>
                          <div style={{ fontWeight: 700, marginBottom: 4 }}>
                            {locale === 'de'
                              ? 'Externe Person — kein Kalendereintrag'
                              : 'External person — no calendar invite'}
                          </div>
                          {locale === 'de'
                            ? <>{email.trim() ? <>Die Adresse <strong>{email}</strong> gehört nicht zum Deloitte-Deutschland-Tenant. </> : <>Eine externe Adresse gehört nicht zum Deloitte-Deutschland-Tenant. </>}Die Person bekommt deshalb <strong>keinen Outlook-Termin</strong> (an externe Adressen können wir keine Termine versenden). Die <strong>Anmeldebestätigung per E-Mail</strong> geht direkt an die Person, mit den <strong>Organizern auf CC</strong> als Nachweis. Die Zustimmung ist <strong>schriftlich</strong> einzuholen.</>
                            : <>{email.trim() ? <>The address <strong>{email}</strong> is not part of the Deloitte Germany tenant. </> : <>An external address is not part of the Deloitte Germany tenant. </>}The person will therefore <strong>not receive an Outlook calendar invite</strong> (we cannot send invites to external addresses). The <strong>confirmation email</strong> is sent directly to the person, with the <strong>organizers on CC</strong> as a record. Consent must be obtained <strong>in writing</strong>.</>}
                        </div>
                      )}
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginTop: 12, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={otherConsentConfirmed}
                          onChange={e => setOtherConsentConfirmed(e.target.checked)}
                          style={{ marginTop: 3 }}
                        />
                        <span style={{ flex: 1, color: 'var(--dex-gray-800)' }}>
                          <span style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                          {locale === 'de'
                            ? (isExternal
                                ? 'Ich bestätige, dass die schriftliche Zustimmung der Person zur stellvertretenden Anmeldung vorliegt.'
                                : 'Ich bestätige, dass die Person ihrer stellvertretenden Anmeldung zugestimmt hat.')
                            : (isExternal
                                ? 'I confirm that the person\'s written consent to this registration on their behalf is on file.'
                                : 'I confirm that the person has consented to this registration on their behalf.')}
                        </span>
                      </label>
                    </div>
                  );
                })()}
                {registerForOther && thirdPartyCheck && (thirdPartyCheck.alreadyRegistered || thirdPartyCheck.notInAudience) && (
                  <div style={{
                    padding: '10px 14px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
                    background: thirdPartyCheck.alreadyRegistered ? 'rgba(200,30,30,0.07)' : 'rgba(237,139,0,0.08)',
                    border: `1px solid ${thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)'}`,
                    color: thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)',
                    fontSize: '0.85rem',
                  }}>
                    {thirdPartyCheck.alreadyRegistered && (() => {
                      // v19.8: Konkrete Meldung mit Name + Anmeldedatum statt
                      // generischem „Diese Person ist bereits angemeldet".
                      const nm = (`${firstName} ${surname}`.trim()) || thirdPartyCheck.registeredName || (locale === 'de' ? 'Diese Person' : 'This person');
                      const d = thirdPartyCheck.registeredDate ? new Date(thirdPartyCheck.registeredDate) : null;
                      const dateStr = d && !isNaN(d.getTime())
                        ? d.toLocaleDateString(locale === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })
                        : '';
                      return (
                        <div>
                          <strong>
                            {nm}{locale === 'de' ? ' ist bereits für das Event angemeldet' : ' is already registered for this event'}
                            {dateStr ? ` (${locale === 'de' ? 'Anmeldedatum' : 'Registered'}: ${dateStr})` : ''}.
                          </strong>
                          <div style={{ marginTop: 4, fontWeight: 400 }}>
                            {locale === 'de'
                              ? 'Eine erneute Anmeldung ist nicht möglich. Bitte wähle eine andere Person.'
                              : 'Registering this person again is not possible. Please pick a different person.'}
                          </div>
                        </div>
                      );
                    })()}
                    {thirdPartyCheck.notInAudience && (
                      <div style={{ marginTop: thirdPartyCheck.alreadyRegistered ? 6 : 0 }}>
                        <strong>{t('reg.thirdparty.notinaudience')}</strong>
                      </div>
                    )}
                  </div>
                )}
              </>
            )}

            {/* v11.80: Anrede-Dropdown nur rendern, wenn das Event das Feld
                explizit abfragt (event.askSalutation === true). Default: aus —
                viele Events brauchen die Anrede nicht. Wenn nicht gerendert,
                bleibt salutation '' und wird so in die Teilnehmer-Zeile
                geschrieben. */}
            {event.askSalutation && (
              <div className="form-group">
                <label className="form-label"><span className="required">*</span> {t('reg.salutation')}</label>
                <select className="form-select" value={salutation} onChange={e => setSalutation(e.target.value as Salutation)} style={showErrors && !salutation ? errorBorder : {}}>
                  <option value="">{t('reg.pleaseselect')}</option>
                  <option value="Herr">{locale === 'de' ? 'Herr' : 'Mr'}</option>
                  <option value="Frau">{locale === 'de' ? 'Frau' : 'Mrs'}</option>
                  <option value="Divers">{locale === 'de' ? 'Divers' : 'Diverse'}</option>
                  <option value="Keine Angabe">{locale === 'de' ? 'Keine Angabe' : 'Prefer not to say'}</option>
                </select>
              </div>
            )}

            {/* v11.97: Sternchen entfernt — die Felder Vorname, Nachname,
                E-Mail werden read-only aus dem SP-Profil befüllt. Der User
                kann sie ohnehin nicht ändern (außer im „Für andere Person
                registrieren"-Modus, dort kommen die echten Required-Marker
                über die Validation). */}
            <div className="form-group">
              <label className="form-label">{t('reg.firstname')}</label>
              <input className="form-input" value={firstName} onChange={e => { if (externalPerson) setFirstName(e.target.value); }} placeholder={t('reg.firstname')} disabled={!externalPerson} style={{ background: externalPerson ? '#fff' : 'var(--dex-gray-100)', ...(showErrors && !firstName.trim() ? errorBorder : {}) }} />
            </div>

            <div className="form-group">
              <label className="form-label">{t('reg.surname')}</label>
              <input className="form-input" value={surname} onChange={e => { if (externalPerson) setSurname(e.target.value); }} placeholder={t('reg.surname')} disabled={!externalPerson} style={{ background: externalPerson ? '#fff' : 'var(--dex-gray-100)', ...(showErrors && !surname.trim() ? errorBorder : {}) }} />
            </div>

            <div className="form-group">
              <label className="form-label">{t('reg.email')}</label>
              <input className="form-input" type="email" value={email} onChange={e => { if (externalPerson) { setEmail(e.target.value); externalEmailConfirmedRef.current = false; /* v18.74: Tippfehler-Check bei Änderung erneut erzwingen */ } }} placeholder={externalPerson ? 'name@firma.de' : 'email@deloitte.de'} disabled={!externalPerson} style={{ background: externalPerson ? '#fff' : 'var(--dex-gray-100)', ...(showErrors && !email.trim() ? errorBorder : {}) }} />
            </div>

            {/* v11.94/v11.97/v12.0: Zusätzliche read-only-Profildaten aus dem
                SP-Profil — Job Title, Geschäftsbereich, Büro.
                Self-Register: aus useCurrentUser(), nur Felder mit Wert.
                For-other-Register: aus pickedUserProfile, alle Felder
                rendern sobald jemand ausgewählt ist (auch leere — sonst
                rätselt der Stellvertreter ob die App das Profil überhaupt
                geladen hat). v18.50: Mobil wird nicht mehr angezeigt — wird
                bei der eigenen Anmeldung auch nicht abgefragt. */}
            {(() => {
              const profile = registerForOther ? pickedUserProfile : currentUser;
              // v15.22: Bei „Für andere registrieren" die Felder auch
              // anzeigen, wenn noch kein Profil gewählt wurde — sonst
              // wundert sich der Stellvertreter, warum nur Name + Mail
              // sichtbar sind. Die Felder bleiben dann leer mit
              // Placeholder „aus SP-Profil — nicht hinterlegt".
              if (!profile && !registerForOther) return null;
              // v18.74: Bei externen Personen gibt es kein Deloitte-Profil —
              // Position/Geschäftsbereich/Büro gar nicht erst anzeigen.
              if (externalPerson) return null;
              const jt = profile ? ((profile as { jobTitle?: string }).jobTitle || '') : '';
              const dept = profile ? ((profile as { department?: string }).department || '') : '';
              const loc = profile ? ((profile as { location?: string }).location || '') : '';
              if (!registerForOther && !jt && !dept && !loc) return null;
              const placeholder = locale === 'de' ? 'aus SP-Profil — nicht hinterlegt' : 'from SP profile — not set';
              const renderField = (label: string, value: string): React.ReactElement => (
                <div className="form-group">
                  <label className="form-label">{label}</label>
                  <input
                    className="form-input"
                    value={value}
                    placeholder={registerForOther && !value ? placeholder : undefined}
                    disabled
                    style={{ background: 'var(--dex-gray-100)' }}
                  />
                </div>
              );
              return (
                <>
                  {(jt || registerForOther) && renderField(locale === 'de' ? 'Position' : 'Job Title', jt)}
                  {(dept || registerForOther) && renderField(locale === 'de' ? 'Geschäftsbereich' : 'Business Area', dept)}
                  {(loc || registerForOther) && renderField(locale === 'de' ? 'Büro' : 'Office', loc)}
                </>
              );
            })()}

            {/* v11.82: Team-Anmeldung-Toggle. Nur sichtbar wenn der Organizer
                in Schritt 4 die Team-Anmeldung aktiviert hat UND der User sich
                NICHT für eine andere Person registriert (Team-fuer-Andere wird
                nicht unterstuetzt — der Stellvertreter-Pfad ist auf eine
                Einzel-Person ausgelegt). */}
            {isTeamCapable && !registerForOther && (
              <div className="form-group" style={{ marginTop: 16, marginBottom: 0 }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={isTeamMode}
                    onChange={e => { setIsTeamMode(e.target.checked); if (e.target.checked) setPendingJoinTeam(null); }}
                    style={{ marginTop: 3 }}
                  />
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)' }}>
                      {locale === 'de'
                        ? `Ich melde mich + mein Team an (Team-Anmeldung)`
                        : 'Register me + my team (team registration)'}
                    </div>
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4, lineHeight: 1.4 }}>
                      {locale === 'de'
                        ? `Belegt bis zu ${teamSize} Plätze auf einmal. Jedes Mitglied bekommt automatisch Bestätigungsmail, Outlook-Termin und sieht das Event in „Meine Events".`
                        : `Books up to ${teamSize} seats at once. Each member automatically receives a confirmation email, an Outlook invite, and sees the event in „My Events".`}
                    </div>
                  </div>
                </label>
              </div>
            )}
          </div>
        </div>

        {/* v11.82: Team-Anmeldung-Card — separat unter „Persönliche Daten",
            nur sichtbar wenn der Toggle aktiv ist. */}
        {isTeamCapable && isTeamMode && !registerForOther && (
          <div className="registration-form" style={{ marginTop: 24 }}>
            <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Icon iconName="People" style={{ fontSize: 16 }} />
              {locale === 'de' ? 'Team-Anmeldung' : 'Team registration'}
            </div>
            <div style={{ padding: '24px 20px' }}>
              {/* Pflicht-Hinweis-Box ganz oben — auffällig orange. */}
              <div style={{
                marginBottom: 20,
                padding: '14px 16px',
                background: 'rgba(237,139,0,0.10)',
                border: '2px solid var(--dex-orange, #ed8b00)',
                borderRadius: 'var(--dex-radius-md)',
                color: '#7a4a00',
                fontSize: '0.88rem',
                lineHeight: 1.55,
              }}>
                <div style={{ fontWeight: 700, marginBottom: 8, fontSize: '0.95rem' }}>
                  {locale === 'de'
                    ? 'Vorab die Zustimmung jedes Teammitglieds einholen'
                    : 'Get every team member\'s consent up front'}
                </div>
                <div style={{ marginBottom: 8 }}>
                  {locale === 'de'
                    ? 'Mit dem Absenden meldest du nicht nur dich selbst an, sondern auch alle weiter unten eingetragenen Personen. Jedes Teammitglied erhält automatisch:'
                    : 'By submitting you register yourself AND every person you add below. Each team member automatically receives:'}
                </div>
                <ul style={{ margin: '0 0 8px 18px', padding: 0 }}>
                  <li>{locale === 'de' ? 'eine Anmeldebestätigung per Mail' : 'a confirmation email'}</li>
                  <li>{locale === 'de' ? 'einen Outlook-Termin im Kalender' : 'an Outlook calendar invite'}</li>
                  <li>{locale === 'de' ? 'den Event in „Meine Events"' : 'the event in „My Events"'}</li>
                </ul>
                <div style={{ marginTop: 4 }}>
                  {locale === 'de'
                    ? <>Bitte stelle sicher, dass alle Teilnehmer ihrer Anmeldung <strong>VORHER zugestimmt</strong> haben — eine Anmeldung ohne Einverständnis ist nicht erlaubt.</>
                    : <>Please make sure every participant has <strong>consented up front</strong> — registering people without their consent is not allowed.</>}
                </div>
              </div>

              {event?.askTeamName && (
                <div className="form-group">
                  <label className="form-label">
                    <span className="required">*</span> {locale === 'de' ? 'Team-Name' : 'Team name'}
                  </label>
                  <input
                    className="form-input"
                    value={teamName}
                    onChange={e => setTeamName(e.target.value.slice(0, 60))}
                    placeholder={locale === 'de' ? 'z.B. „Die Schnellen"' : 'e.g. „The Quick Ones"'}
                    style={showErrors && isTeamMode && event?.askTeamName && !teamName.trim() ? errorBorder : {}}
                  />
                  <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>
                    {locale === 'de' ? 'Max. 60 Zeichen — wird in der Teilnehmerliste mitgespeichert.' : 'Max 60 characters — stored on the participant list.'}
                  </div>
                </div>
              )}

              {/* Member-Slots */}
              <div style={{ marginTop: 8 }}>
                {teamMembers.map((mv, idx) => {
                  const slotRequired = !teamPartialAllowed;
                  const parsed = parseTeamMember(mv);
                  const isErr = showErrors && isTeamMode && slotRequired && !parsed;
                  return (
                    <div className="form-group" key={`team-slot-${idx}`}>
                      <label className="form-label">
                        {slotRequired && <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>}
                        {locale === 'de'
                          ? `Mitglied ${idx + 2}${slotRequired ? '' : ' (optional)'}`
                          : `Member ${idx + 2}${slotRequired ? '' : ' (optional)'}`}
                      </label>
                      <UserFieldPicker
                        value={mv}
                        onChange={v => {
                          const next = [...teamMembers];
                          next[idx] = v;
                          setTeamMembers(next);
                        }}
                        searchUsers={async (q, includeIntl) => {
                          const results = await searchUsers(q, includeIntl);
                          return results.map(r => ({ email: r.email, displayName: r.displayName, location: r.location, jobTitle: r.jobTitle }));
                        }}
                        searchUserByEmail={searchUser}
                        placeholder={locale === 'de' ? 'Name oder E-Mail eingeben...' : 'Type a name or email...'}
                        errorStyle={isErr ? errorBorder : {}}
                        forcedIsDe={locale === 'de'}
                      />
                      {/* v18.12: Custom-Fields pro Team-Mitglied — erscheinen,
                          sobald die Person ausgewählt ist (z.B. Essenspräferenz). */}
                      {parsed && teamMemberApplicableFields.length > 0 && (
                        <div style={{ marginTop: 8, marginLeft: 8, paddingLeft: 12, borderLeft: '2px solid var(--dex-gray-200)' }}>
                          {teamMemberApplicableFields
                            .filter(f => {
                              if (!f.showIf) return true;
                              const mv = teamMemberFields[idx] || {};
                              const v = mv[f.showIf.fieldId] || '';
                              const parts = v.split(' | ').map(s => s.trim());
                              return f.showIf.values.some(x => v === x || parts.indexOf(x) >= 0);
                            })
                            .map(f => renderRegField(
                              f,
                              teamMemberFields[idx] || {},
                              next => setTeamMemberFields(prev => ({ ...prev, [idx]: next }))
                            ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* Pflicht-Bestätigungs-Checkbox */}
              <div className="form-group" style={{ marginTop: 18, marginBottom: 0 }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={teamConsentConfirmed}
                    onChange={e => setTeamConsentConfirmed(e.target.checked)}
                    style={{ marginTop: 3 }}
                  />
                  <div style={{ flex: 1, fontSize: '0.88rem', color: 'var(--dex-gray-800)', lineHeight: 1.5 }}>
                    <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                    {locale === 'de'
                      ? 'Ich bestätige, dass alle eingetragenen Teammitglieder ihrer Anmeldung zugestimmt haben.'
                      : 'I confirm that every listed team member has consented to this registration.'}
                  </div>
                </label>
              </div>
            </div>
          </div>
        )}

        {/* v18.73: Offene Teams — sichtbar wenn der Organizer „Offene Slots
            öffentlich sichtbar" aktiviert hat und es Teams gibt, denen Plätze
            fehlen. Steht jetzt UNTER der „Ich melde mich + mein Team an"-Karte.
            Klick auf „Vormerken" wählt ein Team nur vor — die eigentliche
            Anmeldung (inkl. der oben/unten ausgefüllten persönlichen +
            event-spezifischen Felder) passiert erst über den „Anmelden"-Button.
            */}
        {event && event.teamRegistrationEnabled && event.teamOpenSlotsVisible && !registerForOther && openTeamsLoaded && openTeams.length > 0 && !parentAlreadyRegistered && (
          <div className="registration-form" style={{ marginBottom: 16 }}>
            <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Icon iconName="People" style={{ fontSize: 16 }} />
              {locale === 'de' ? 'Offene Teams — einem unvollständigen Team beitreten' : 'Open teams — join an incomplete team'}
            </div>
            <div style={{ padding: '20px' }}>
              <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)', marginTop: 0, marginBottom: 12, lineHeight: 1.5 }}>
                {locale === 'de'
                  ? 'Andere Personen haben Teams angemeldet, denen noch Plätze fehlen. Du kannst eines vormerken — fülle dann oben deine persönlichen Daten und unten die event-spezifischen Angaben aus und klicke auf „Anmelden", um beizutreten.'
                  : 'Other people have registered teams with open slots. Pre-select one — then fill in your personal details above and the event-specific information below, and click „Register" to join.'}
                {event.teamJoinRequiresApproval && (
                  <> {locale === 'de'
                    ? <><br /><strong>Hinweis:</strong> der Team-Kapitän muss deinen Beitritt erst bestätigen.</>
                    : <><br /><strong>Note:</strong> the team lead has to approve your join.</>}
                  </>
                )}
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {openTeams.map(t => {
                  const free = t.teamSize - t.activeCount;
                  const isPicked = !!pendingJoinTeam && pendingJoinTeam.teamId === t.teamId;
                  return (
                    <div key={t.teamId} style={{
                      display: 'flex', alignItems: 'center', gap: 12,
                      padding: '10px 12px',
                      background: isPicked ? 'rgba(134,188,37,0.10)' : 'var(--dex-gray-50, #f7f7f7)',
                      borderRadius: 6,
                      border: isPicked ? '2px solid var(--dex-green, #86bc25)' : '1px solid var(--dex-gray-200)',
                    }}>
                      <Icon iconName="Group" style={{ fontSize: 16, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                      <div style={{ flex: 1, fontSize: '0.88rem' }}>
                        <div style={{ fontWeight: 600 }}>
                          {locale === 'de'
                            ? `Team „${t.teamName || 'ohne Namen'}"`
                            : `Team „${t.teamName || 'unnamed'}"`}
                        </div>
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                          {locale === 'de'
                            ? `${t.activeCount}/${t.teamSize} belegt — ${free} Slot${free === 1 ? '' : 's'} frei`
                            : `${t.activeCount}/${t.teamSize} taken — ${free} slot${free === 1 ? '' : 's'} free`}
                        </div>
                      </div>
                      <button
                        type="button"
                        className={isPicked ? 'btn btn-secondary' : 'btn btn-primary'}
                        onClick={() => togglePendingJoinTeam(t.teamId, t.teamName)}
                        style={{ fontSize: '0.82rem', padding: '6px 12px' }}
                      >
                        {isPicked
                          ? (locale === 'de' ? 'Vorgemerkt ✓ — entfernen' : 'Pre-selected ✓ — remove')
                          : (locale === 'de' ? 'Vormerken' : 'Pre-select')}
                      </button>
                    </div>
                  );
                })}
              </div>
              <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 10, marginBottom: 0, lineHeight: 1.4 }}>
                {locale === 'de'
                  ? 'Mitgliedernamen werden aus Privatsphäre-Gründen nicht angezeigt.'
                  : 'Member names are hidden for privacy reasons.'}
              </p>
              {pendingJoinTeam && (
                <div style={{
                  marginTop: 12,
                  padding: '10px 12px',
                  borderRadius: 6,
                  background: 'rgba(134,188,37,0.10)',
                  border: '1px solid var(--dex-green, #86bc25)',
                  color: 'var(--dex-green-dark, #3f5f10)',
                  fontSize: '0.85rem',
                  lineHeight: 1.5,
                }}>
                  {locale === 'de'
                    ? <>Team <strong>„{pendingJoinTeam.teamName || 'ohne Namen'}“</strong> ist vorgemerkt. Fülle deine Angaben aus und klicke unten auf <strong>„Anmelden“</strong>, um den Beitritt abzuschließen.</>
                    : <>Team <strong>“{pendingJoinTeam.teamName || 'unnamed'}”</strong> is pre-selected. Fill in your details and click <strong>“Register”</strong> below to complete your join.</>}
                </div>
              )}
            </div>
          </div>
        )}

        {/* Eventspezifische Felder (inkl. Split-Capacity Starter-Typ-Auswahl wenn
            beide Kapazitaeten > 0; bei nur einem verfuegbaren Typ wird dieser
            automatisch gesetzt und gar nicht angezeigt). v10.20: Sessions-/
            Hauptevent-Auswahl ist hierher gewandert (vorher links unter der
            Event-Karte). */}
        {/* v18.73: Die „Event-spezifische Informationen"-Karte nur anzeigen,
            wenn es dort tatsächlich etwas auszufüllen/auszuwählen gibt — also
            Custom-Felder, eine Gruppen-Auswahl (Split) ODER eine Sub-Event-
            Auswahl. Sonst (leeres „Keine zusätzlichen Informationen
            erforderlich") wird die Karte komplett ausgeblendet. */}
        {(event.eventSpecificFields.length > 0 || isSplitGroup || childEvents.length > 0) && (
        <div className="registration-specific">
          {/* v11.97: Section-Header + „* = Required field"-Legende in
              einer Zeile. Legende mit ROTEM Stern (vorher war der Stern
              in der Erklärung grau, jetzt im Deloitte-Rot wie alle echten
              Required-Marker). */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
            <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Icon iconName="EditNote" style={{ fontSize: 16 }} />
              {t('reg.eventinfo')}
            </div>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 12, padding: '0 12px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                <span style={{ color: 'var(--dex-red, #da291c)', fontWeight: 700, marginRight: 2 }}>*</span> = {t('reg.requiredfield')}
              </span>
              {/* v18.12: „Zurücksetzen" ersetzt den frueheren „Löschen"-Button
                  aus der Aktions-Zeile — leert das Formular (Eingaben). */}
              <button
                type="button"
                className="btn btn-secondary"
                onClick={handleClear}
                disabled={isSubmitting}
                style={{ fontSize: '0.78rem', padding: '4px 12px' }}
                title={locale === 'de' ? 'Eingaben zurücksetzen' : 'Reset inputs'}
              >
                <Trash2 size={14} /> {locale === 'de' ? 'Zurücksetzen' : 'Reset'}
              </button>
            </span>
          </div>
          <div style={{ padding: '24px 20px' }}>
            {/* v11.10: Group-Selection ist ein eigener, IMMER sichtbarer
                Block (sofern das Event Split-Capacity hat). Vorher war er
                inkorrekt INNERHALB der Sub-Events-Auswahl genistet, sodass
                Events ohne Sub-Events keine Gruppen-Buttons hatten und
                Drittpersonen-Registrierungen die Gruppen-Wahl gar nicht
                anzeigten. Sub-Events erben jetzt einfach
                preferredStarterType — keine Pro-Sub-Event-Radios mehr. */}
            {isSplitGroup && (
              <div style={{ marginBottom: 20, border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 16 }}>
                <label className="form-label" style={{ fontWeight: 700, marginBottom: 6 }}>
                  <span className="required">*</span> {locale === 'de' ? 'Gruppen-Auswahl' : 'Group selection'}
                </label>
                <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 10 }}>
                  {locale === 'de'
                    ? `Wähle eine der zwei Gruppen aus. Ist die Wunsch-Gruppe voll, kannst du automatisch in die andere wechseln oder auf der Warteliste warten.`
                    : 'Pick one of the two groups. If your preferred group is full, you can either switch to the other or join the waitlist.'}
                </p>
                {/* v19.19: Gesamt-Kapazitäts-Zusammenfassung — Gesamtzahl der
                    Plätze, aktuell freie Plätze (geklammert ≥ 0) und die Zahl
                    der Personen auf der Warteliste. WICHTIG: hier wird NIE eine
                    Überbuchung sichtbar — freie Plätze sind bei 0 gedeckelt,
                    bei Vollbelegung steht „ausgebucht". Die echte
                    Überbuchungszahl ist ausschließlich dem Organizer/Admin
                    im Admin Center vorbehalten. */}
                {(() => {
                  const totalCap = durchCap + funCap;
                  const dActive = starterCounts?.durch ?? 0;
                  const fActive = starterCounts?.fun ?? 0;
                  const totalFree = Math.max(0, durchCap - dActive) + Math.max(0, funCap - fActive);
                  const totalWait = (starterCounts?.durchWait ?? 0) + (starterCounts?.funWait ?? 0);
                  return (
                    <div style={{
                      display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap',
                      marginBottom: 12, padding: '8px 12px', borderRadius: 8,
                      background: 'var(--dex-gray-50, #f7f7f7)', border: '1px solid var(--dex-gray-200)',
                      fontSize: '0.8rem',
                    }}>
                      <span style={{ color: 'var(--dex-gray-700)', fontWeight: 700 }}>
                        {locale === 'de' ? 'Gesamtkapazität:' : 'Total capacity:'} {totalCap} {locale === 'de' ? 'Plätze' : 'seats'}
                      </span>
                      <span style={{ color: 'var(--dex-gray-300)' }}>·</span>
                      {totalFree > 0 ? (
                        <span style={{ color: 'var(--dex-green-dark, #6b9a1e)', fontWeight: 700 }}>
                          {totalFree} {locale === 'de' ? 'frei' : 'free'}
                        </span>
                      ) : (
                        <span style={{ color: 'var(--dex-red, #c00)', fontWeight: 700 }}>
                          {locale === 'de' ? 'ausgebucht' : 'fully booked'}
                        </span>
                      )}
                      {totalWait > 0 && (
                        <>
                          <span style={{ color: 'var(--dex-gray-300)' }}>·</span>
                          <span style={{ color: 'var(--dex-gray-600)' }}>
                            {totalWait} {locale === 'de'
                              ? (totalWait === 1 ? 'Person auf der Warteliste' : 'Personen auf der Warteliste')
                              : (totalWait === 1 ? 'person on the waitlist' : 'people on the waitlist')}
                            {event.splitSharedWaitlist ? (locale === 'de' ? ' (gemeinsam)' : ' (shared)') : ''}
                          </span>
                        </>
                      )}
                    </div>
                  );
                })()}
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                  {(() => {
                    const optA = { id: 'Durchstarter', label: splitLabelA, desc: splitLabelA === 'Durchstarter' ? t('reg.starter.durch.desc') : '', cap: durchCap, count: starterCounts?.durch ?? 0, wait: starterCounts?.durchWait ?? 0, color: 'var(--dex-green-dark, #6b9a1e)' };
                    const optB = { id: 'Funstarter', label: splitLabelB, desc: splitLabelB === 'Funstarter' ? t('reg.starter.fun.desc') : '', cap: funCap, count: starterCounts?.fun ?? 0, wait: starterCounts?.funWait ?? 0, color: 'var(--dex-orange, #ff8c00)' };
                    // v11.25: pure UI-Reihenfolge — bei reversed wird Karte B
                    // zuerst gerendert. Interne IDs/Capacities/StarterType der
                    // Anmeldungen bleiben unangetastet.
                    return event.splitDisplayOrderReversed ? [optB, optA] : [optA, optB];
                  })().map(opt => {
                    const free = opt.cap - opt.count;
                    const isFull = free <= 0;
                    const isActive = preferredStarterType === opt.id;
                    return (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => setPreferredStarterType(opt.id)}
                        style={{
                          padding: 14, textAlign: 'left',
                          borderRadius: 'var(--dex-radius, 12px)',
                          border: isActive ? `2px solid ${opt.color}` : '2px solid var(--dex-gray-200)',
                          background: isActive ? 'var(--dex-green-light, #f0fdf4)' : '#fff',
                          cursor: 'pointer', transition: 'all 0.15s',
                          position: 'relative',
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                          <strong style={{ color: opt.color, fontSize: '0.95rem' }}>{opt.label}</strong>
                          {isActive && <span style={{ color: opt.color, fontSize: '0.8rem' }}>✓</span>}
                        </div>
                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginBottom: 6 }}>{opt.desc}</div>
                        <div style={{ fontSize: '0.78rem' }}>
                          {isFull ? (
                            <span style={{ color: 'var(--dex-red, #c00)', fontWeight: 600 }}>{t('reg.starter.full')}</span>
                          ) : (
                            // v19.19: nie negativ — bei Überbuchung greift der
                            // isFull-Zweig oben (zeigt „Voll"), die echte
                            // Überbuchungszahl bleibt dem Organizer/Admin vorbehalten.
                            <span style={{ color: opt.color }}>{`${Math.max(0, free)} / ${opt.cap} ${t('reg.starter.free')}`}</span>
                          )}
                          {/* v19.19: Warteliste pro Gruppe — nur bei GETRENNTEN
                              Wartelisten. Bei gemeinsamer Warteliste steht die
                              Zahl gesammelt in der Kapazitäts-Zusammenfassung. */}
                          {!event.splitSharedWaitlist && opt.wait > 0 && (
                            <span style={{ display: 'block', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {opt.wait} {locale === 'de'
                                ? (opt.wait === 1 ? 'Person auf der Warteliste' : 'Personen auf der Warteliste')
                                : (opt.wait === 1 ? 'person on the waitlist' : 'people on the waitlist')}
                            </span>
                          )}
                        </div>
                      </button>
                    );
                  })}
                </div>

                {event.durchstarterRequiresProof && preferredStarterType === 'Durchstarter' && (
                  <div style={{ marginTop: 12, padding: '10px 12px', background: 'rgba(237,139,0,0.06)', border: '1px solid var(--dex-orange)', borderRadius: 8 }}>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
                      <input
                        type="checkbox"
                        checked={eventSpecific['b2run_leistungsnachweis'] === 'true'}
                        onChange={e => setEventSpecific({ ...eventSpecific, b2run_leistungsnachweis: e.target.checked ? 'true' : 'false' })}
                        style={{ marginTop: 3 }}
                      />
                      <span>
                        <strong>{t('reg.starter.proof') || 'Leistungsnachweis vorhanden'} <span className="required">*</span></strong>
                        <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                          {t('reg.starter.proof.hint') || 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.'}
                        </span>
                      </span>
                    </label>
                    {showErrors && eventSpecific['b2run_leistungsnachweis'] !== 'true' && (
                      <div style={{ marginTop: 6, fontSize: '0.75rem', color: 'var(--dex-red)' }}>
                        {t('reg.starter.proof.required') || 'Bitte Leistungsnachweis bestätigen.'}
                      </div>
                    )}
                  </div>
                )}
                {/* v11.12: Custom-Fields mit onlyForGroup-Constraint
                    direkt INNERHALB der Gruppen-Auswahl-Box rendern.
                    Klappt erst auf, wenn der User eine Gruppe gewählt
                    hat und das Feld der gewählten Gruppe entspricht.
                    Gleicher orange-getönter Style wie der Legacy-
                    Leistungsnachweis-Block — sodass jede Gruppen-
                    spezifische Abfrage optisch klar als „Folge der
                    Gruppen-Wahl" erkennbar ist. */}
                {preferredStarterType && (() => {
                  const groupSpec = event.eventSpecificFields
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
                    .filter(f => {
                      const grp = f.onlyForGroup;
                      if (!grp || grp === 'all') return false;
                      if (grp === 'A') return preferredStarterType === 'Durchstarter';
                      if (grp === 'B') return preferredStarterType === 'Funstarter';
                      return false;
                    });
                  if (groupSpec.length === 0) return null;
                  const labelA = (event.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
                  const labelB = (event.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
                  const grpLabel = preferredStarterType === 'Durchstarter' ? labelA : labelB;
                  return (
                    <div style={{
                      marginTop: 12, padding: '12px 14px',
                      background: 'rgba(237,139,0,0.06)',
                      border: '1px solid var(--dex-orange)',
                      borderRadius: 8,
                      display: 'flex', flexDirection: 'column', gap: 12,
                    }}>
                      <div style={{ fontSize: '0.78rem', fontWeight: 700, color: 'var(--dex-orange, #ed8b00)' }}>
                        {locale === 'de'
                          ? `Zusätzliche Angaben für „${grpLabel}"`
                          : `Additional details for „${grpLabel}"`}
                      </div>
                      {groupSpec.map(f => renderRegField(f))}
                    </div>
                  );
                })()}
              </div>
            )}

            {/* v11.10: Sub-Events-Auswahl als eigener Block.
                v18.37: jetzt AUCH im Stellvertreter-Modus („Für andere
                registrieren") sichtbar — der Submit-Pfad meldet ausgewählte
                Sub-Events längst für die Zielperson an (registerForEvent mit
                participantEmail), nur die Auswahl-UI war versehentlich hinter
                !registerForOther versteckt. Die Belegungszahlen werden im
                Stellvertreter-Modus über getAllRegistrations geladen, ohne
                Self-Vorbelegung. Der Haupt-Event-Checkbox wird im
                Stellvertreter-Modus ausgeblendet (das Haupt-Event wird dort
                ohnehin immer mit angemeldet). Sub-Events erben
                preferredStarterType vom Group-Selection-Block oben. */}
            {childEvents.length > 0 && (
              <div style={{ marginBottom: 20, border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 16 }}>
                {/* v15.11: im subEventsOnlyMode ist die Hauptevent-Anmeldung
                    deaktiviert — Überschrift + Hinweis entsprechend
                    anpassen, sonst lesen sich „Haupt-Event und … können
                    unabhängig" widersprüchlich. */}
                <h4 style={{ marginTop: 0, marginBottom: 4, fontSize: '0.95rem' }}>
                  {event.subEventsOnlyMode
                    ? (childTermPlural
                        ? (locale === 'de' ? `${childTermPlural} auswählen` : `Select ${childTermPlural}`)
                        : (locale === 'de' ? 'Sub-Events auswählen' : 'Select sub-events'))
                    : (tEvent('reg.selection.title') || 'Wofür möchtest du dich anmelden?')}
                </h4>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                  {event.subEventsOnlyMode
                    ? (childTermPlural
                        ? (locale === 'de'
                            ? `Bitte wähle mindestens eine ${childTermSingular || 'Sub-Event'}, für die du dich anmelden möchtest.`
                            : `Please pick at least one ${childTermSingular || 'sub-event'} you want to register for.`)
                        : (locale === 'de'
                            ? 'Bitte wähle mindestens ein Sub-Event, für das du dich anmelden möchtest.'
                            : 'Please pick at least one sub-event you want to register for.'))
                    : registerForOther
                      ? (locale === 'de'
                          ? `Die Person wird für das Haupt-Event angemeldet. Wähle zusätzlich die gewünschten ${childTermPlural || 'Sub-Events'} aus.`
                          : `The person will be registered for the main event. Additionally pick the desired ${childTermPlural || 'sub-events'}.`)
                    : (childTermPlural
                        ? (locale === 'de'
                            ? `Haupt-Event und ${childTermPlural} können unabhängig voneinander an- oder abgewählt werden.`
                            : `Main event and ${childTermPlural} can be selected or deselected independently.`)
                        : (tEvent('reg.selection.hint') || 'Haupt-Event und Sessions können unabhängig voneinander an- oder abgewählt werden.'))}
                </p>

                {/* v15.7: Hauptevent-Card auch hier ausblenden bei
                    subEventsOnlyMode — gleicher Fix wie der primäre Pfad
                    weiter oben. Vorher wurde dieser Render-Pfad (Register
                    for someone else) übersehen.
                    v18.37: im Stellvertreter-Modus ebenfalls ausblenden — die
                    Person wird dort immer für das Haupt-Event angemeldet, ein
                    steuerbarer Haken wäre irreführend. */}
                {!event.subEventsOnlyMode && !registerForOther && (
                <label style={{
                  display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
                  borderRadius: 8,
                  border: `1px solid ${registerForParent && !parentAlreadyRegistered ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                  background: registerForParent && !parentAlreadyRegistered ? 'rgba(134,188,37,0.06)' : '#fff',
                  cursor: parentAlreadyRegistered ? 'default' : 'pointer',
                }}>
                  <input
                    type="checkbox"
                    checked={parentAlreadyRegistered ? true : registerForParent}
                    disabled={parentAlreadyRegistered}
                    onChange={e => setRegisterForParent(e.target.checked)}
                    style={{ marginTop: 2 }}
                  />
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 700 }}>{tEvent('reg.selection.mainevent') || 'Haupt-Event'}: {event.title}</div>
                    {parentAlreadyRegistered && (
                      <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                        {tEvent('reg.selection.alreadyregistered') || 'Du bist bereits für das Haupt-Event angemeldet.'}
                      </div>
                    )}
                  </div>
                </label>
                )}

                {/* Sessions */}
                <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', fontWeight: 600 }}>{childTermPlural || tEvent('reg.selection.sessions') || 'Sessions'}</div>
                  {childEvents.map(ce => {
                    const meta = sessionMeta[ce.id] || { count: 0, wasRegistered: false };
                    const isSel = selectedSessions.has(ce.id);
                    const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
                    const isSessionFull = hasCap && meta.count >= (ce.maxParticipants || 0);
                    const deadlinePassed = !!(ce.registrationDeadline && new Date(ce.registrationDeadline) < new Date());
                    const disabled = (isSessionFull && !isSel) || (deadlinePassed && !isSel);

                    return (
                      <div key={ce.id} style={{
                        padding: 10, borderRadius: 8,
                        border: `1px solid ${isSel ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                        background: isSel ? 'rgba(134,188,37,0.06)' : '#fff',
                      }}>
                        <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: disabled ? 'not-allowed' : 'pointer' }}>
                          <input
                            type="checkbox"
                            checked={isSel}
                            disabled={disabled}
                            onChange={e => {
                              if (e.target.checked) {
                                const hasCustomFields = (ce.eventSpecificFields || []).length > 0;
                                if (hasCustomFields) {
                                  setPendingSubEventModal({
                                    subEventId: ce.id,
                                    draftValues: { ...(sessionFieldValues[ce.id] || {}) },
                                  });
                                } else {
                                  const next = new Set(selectedSessions);
                                  next.add(ce.id);
                                  setSelectedSessions(next);
                                }
                              } else {
                                const next = new Set(selectedSessions);
                                next.delete(ce.id);
                                setSelectedSessions(next);
                                setSessionFieldValues(prev => {
                                  const copy = { ...prev };
                                  delete copy[ce.id];
                                  return copy;
                                });
                              }
                            }}
                            style={{ marginTop: 2 }}
                          />
                          <div style={{ flex: 1 }}>
                            <div style={{ fontWeight: 700 }}>{ce.title || tEvent('reg.subevents.untitled')}</div>
                            {ce.description && (
                              // v11.97: gleiche Schriftgröße wie der Titel
                              // (Standard-Body). Vorher 0.78rem klein.
                              <div style={{ color: 'var(--dex-gray-600)', marginTop: 2 }}>{ce.description}</div>
                            )}
                            {/* v11.94: gleiches Icon-Layout wie oben (anderer
                                Render-Pfad für Team-Modus). */}
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, marginTop: 4, color: 'var(--dex-gray-600)' }}>
                              {ce.startDate && (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                  {/* v11.97: Icon-Größe an Standard-Body angepasst. */}
                                  <Icon iconName="Calendar" style={{ fontSize: 15, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                                  {formatDate(ce.startDate)}
                                </span>
                              )}
                              {ce.location && (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                  <Icon iconName="POI" style={{ fontSize: 15, color: '#0a3766' }} />
                                  {ce.location}
                                </span>
                              )}
                            </div>
                            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {hasCap && (() => {
                                const sessionFree = Math.max(0, (ce.maxParticipants || 0) - (meta.count || 0));
                                return (
                                  <> · <span style={{ color: isSessionFull ? 'var(--dex-red)' : 'inherit', fontWeight: 600 }}>
                                    {/* v19.19: belegt-Zahl bei der Kapazität deckeln (s.o.) —
                                        Überbuchung nie auf der Anmeldeseite anzeigen. */}
                                    {Math.min(meta.count, ce.maxParticipants || 0)}/{ce.maxParticipants} {tEvent('reg.subevents.taken')}
                                  </span>
                                  {!isSessionFull && (
                                    <span style={{ color: 'var(--dex-green-dark)' }}> — {sessionFree} {tEvent('reg.free')}</span>
                                  )}
                                  </>
                                );
                              })()}
                            </div>
                            {deadlinePassed && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-orange)', marginTop: 2 }}>
                                {tEvent('reg.subevents.deadlinepassed')}
                              </div>
                            )}
                            {isSessionFull && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-red)', marginTop: 2 }}>
                                {tEvent('reg.subevents.sessionfull')}
                              </div>
                            )}
                            {/* v11.10: Hardcoded Sub-Event-Gruppen-Radios entfernt.
                                Sub-Events erben jetzt grundsätzlich
                                preferredStarterType vom Group-Selection-Block
                                oben. Pro-Sub-Event-Gruppe ist konzeptionell
                                Quatsch — die Gruppe gehört zum Teilnehmer
                                (z.B. „Vormittag/Nachmittag"), nicht zur
                                Session. */}
                          </div>
                        </label>
                      </div>
                    );
                  })}
                </div>

                {isSessionsOnlyMode && selectedSessions.size > 0 && !event.subEventsOnlyMode && (
                  <div style={{
                    marginTop: 12, padding: '8px 10px', borderRadius: 6,
                    background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                    color: 'var(--dex-orange)', fontSize: '0.78rem',
                  }}>
                    {childTermPlural
                      ? (locale === 'de'
                          ? `Du meldest dich ausschließlich für ${childTermPlural} an — NICHT für das Haupt-Event.`
                          : `You are registering exclusively for ${childTermPlural} — NOT for the main event.`)
                      : (tEvent('reg.selection.sessionsonlyhint') || 'Du meldest dich ausschließlich für Sessions an — NICHT für das Haupt-Event.')}
                  </div>
                )}
              </div>
            )}
            {event.eventSpecificFields.length === 0 && !isSplitGroup ? (
              <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{t('reg.noadditional')}</p>
            ) : (
              // v11.2 / v11.5: Custom-Fields ohne Pro-Gruppe-Constraint im
              // 2-Spalten-Grid. Group-spezifische Felder werden bereits
              // oben innerhalb der Gruppen-Auswahl-Box gerendert und hier
              // ausgefiltert.
              <div className="dex-reg-fields-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
              {event.eventSpecificFields
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
                })
                .map(f => renderRegField(f))
              }
              </div>
            )}
          </div>
        </div>
        )}
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

      {/* Buttons */}
      <div className="registration-actions mt-24" style={{ maxWidth: 1100, margin: '24px auto 0' }}>
        {(() => {
          // v15.11: im subEventsOnlyMode (Hauptevent nicht anmeldbar) muss
          // mindestens ein Sub-Event ausgewählt sein, sonst Button ausgrauen
          // + Hinweis statt „Registrieren (Haupt-Event)" zeigen.
          const isSubOnly = !!(event && event.subEventsOnlyMode) && !registerForOther;
          const nothingPicked = isSubOnly && selectedSessions.size === 0;
          // v15.16: Consent-Pflicht bei „Für andere registrieren".
          const needsOtherConsent = registerForOther && !!email.trim() && !otherConsentConfirmed;
          // v19.8: Bei stellvertretender Anmeldung den Button sperren, wenn die
          // ausgewählte Person bereits angemeldet ist — vorher konnte man
          // trotz Hinweis auf „Registrieren" klicken (und es kam danach noch
          // die CC-Frage). Jetzt klare Blockade direkt am Button.
          const targetAlreadyRegistered = registerForOther && !!(thirdPartyCheck && thirdPartyCheck.alreadyRegistered);
          // v18: Demo-Event — Register-Button ist bewusst NICHT auswaehlbar
          // (keine echte Anmeldung; reine Showcase-Ansicht).
          const isDemo = !!(event && event.isDemoShowcase);
          const isDisabled = isDemo || isSubmitting || (isTeamMode && !teamValidation.ok) || nothingPicked || needsOtherConsent || targetAlreadyRegistered;
          const titleAttr = isDemo
            ? (locale === 'de' ? 'Demo-Event — eine echte Anmeldung ist nicht möglich.' : 'Demo event — real registration is not possible.')
            : (targetAlreadyRegistered
            ? (locale === 'de' ? 'Diese Person ist bereits für das Event angemeldet.' : 'This person is already registered for this event.')
            : (isTeamMode && !teamValidation.ok
            ? (teamValidation.reason || '')
            : (nothingPicked
                ? (locale === 'de'
                    ? `Bitte mindestens ${childTermSingular ? `eine ${childTermSingular}` : 'ein Sub-Event'} auswählen.`
                    : `Please pick at least one ${childTermSingular || 'sub-event'}.`)
                : (needsOtherConsent
                    ? (locale === 'de'
                        ? 'Bitte bestätige die Zustimmung der Person.'
                        : 'Please confirm the person\'s consent.')
                    : ''))));
          return (
            <button
              className="btn btn-primary"
              onClick={handleSubmit}
              disabled={isDisabled}
              title={titleAttr}
            >
              <Send size={16} /> {(() => {
                if (isSubmitting) return t('reg.submitting');
                // v18.73: Vorgemerkter Team-Beitritt — eigener Button-Text.
                if (pendingJoinTeam) {
                  return event?.teamJoinRequiresApproval
                    ? (locale === 'de' ? 'Beitritt anfragen' : 'Request to join')
                    : (locale === 'de' ? 'Team beitreten & anmelden' : 'Join team & register');
                }
                // v11.82: Team-Modus — eigener Button-Text mit Personen-Zahl.
                if (isTeamMode) {
                  const n = 1 + teamMembersParsed.filter(Boolean).length;
                  return locale === 'de'
                    ? `Team anmelden (${n} ${n === 1 ? 'Person' : 'Personen'})`
                    : `Register team (${n} ${n === 1 ? 'person' : 'people'})`;
                }
                if (nothingPicked) {
                  return locale === 'de'
                    ? `Bitte mindestens ${childTermSingular ? `eine ${childTermSingular}` : 'ein Sub-Event'} auswählen`
                    : `Please pick at least one ${childTermSingular || 'sub-event'}`;
                }
                if (registerForOther) return t('reg.register');
                // v7.3: Kein Selection-Block → einfacher "Registrieren"-Text ohne
                // Parantheses-Info. Erst wenn Sub-Events existieren, zeigen wir
                // detailliert an, was gerade submittet wird.
                if (childEvents.length === 0) return t('reg.register');
                const parts: string[] = [];
                if (willRegisterParent) parts.push(t('reg.selection.mainevent') || 'Haupt-Event');
                if (selectedSessions.size > 0) {
                  parts.push(`${selectedSessions.size} ${selectedSessions.size === 1 ? (childTermSingular || t('reg.selection.sessioncount.one') || 'Session') : (childTermPlural || t('reg.selection.sessioncount.many') || 'Sessions')}`);
                }
                if (parts.length === 0) return t('reg.register');
                return `${t('reg.register')} (${parts.join(' + ')})`;
              })()}
            </button>
          );
        })()}
        {/* v18.11: „Ich nehme nicht teil" — proaktive Absage. Nur bei
            Selbst-Anmeldung (nicht „für andere", nicht Team-Modus, kein
            Demo-Event). Braucht keine Pflichtfelder. */}
        {!registerForOther && !isTeamMode && !pendingJoinTeam && !(event && event.isDemoShowcase) && (
          <button
            type="button"
            className="btn btn-secondary"
            onClick={handleDecline}
            disabled={isDeclining || isSubmitting}
            title={locale === 'de'
              ? 'Melde zurück, dass du nicht teilnehmen wirst (keine Anmeldung).'
              : 'Let us know you will not attend (no registration).'}
            style={{ color: 'var(--dex-red, #c00)' }}
          >
            {isDeclining
              ? (locale === 'de' ? 'Wird gesendet…' : 'Submitting…')
              : (locale === 'de' ? 'Ich nehme nicht teil' : 'I will not attend')}
          </button>
        )}
      </div>

      {/* Datenschutz-Hinweis als Fußnote ganz unten.
          v11.93: Breite auf 1100px begrenzt + zentriert (analog
          .registration-layout), damit der Text nicht über die ganze
          App-Breite läuft. */}
      <div
        className="footer-disclaimer mt-24"
        style={{ borderRadius: 'var(--dex-radius-lg)', maxWidth: 1100, margin: '24px auto 0' }}
      >
        <p>
          {t('reg.privacy').replace('{title}', event.title)}
        </p>
      </div>

      {/* Fallback-Dialog (seit v6.5): Wunsch-Starter-Typ voll, aber Alternative frei.
          User entscheidet explizit zwischen Umsteigen oder Warteliste. */}
      {fallbackDialog && (
        <Modal
          open={true}
          onClose={() => setFallbackDialog(null)}
          maxWidth={480}
          padding={24}
          ariaLabel="Plätze voll"
        >
            {(() => {
              // v10.20: Label-Mapping fuer die freie Bezeichnung — wunsch/alt
              // sind interne IDs ('Durchstarter' / 'Funstarter'); die Anzeige
              // nimmt splitLabelA / splitLabelB.
              const wunschLabel = fallbackDialog.wunsch === 'Durchstarter' ? splitLabelA : splitLabelB;
              const altLabel = fallbackDialog.alt === 'Durchstarter' ? splitLabelA : splitLabelB;
              return (
                <>
                  {/* v17.22: Attendee-facing → bilingual. Vorher war dieser
                      Fallback-Dialog (Wunsch-Gruppe voll) rein deutsch. */}
                  <h3 style={{ margin: 0, marginBottom: 10 }}>
                    {locale === 'de' ? `${wunschLabel}-Plätze sind voll` : `${wunschLabel} is full`}
                  </h3>
                  <p style={{ color: 'var(--dex-gray-700)', lineHeight: 1.5, marginBottom: 8 }}>
                    {locale === 'de'
                      ? <>Für <strong>{wunschLabel}</strong> gibt es aktuell keine freien Plätze mehr.</>
                      : <>There are currently no free spots left for <strong>{wunschLabel}</strong>.</>}
                  </p>
                  <p style={{ color: 'var(--dex-gray-700)', lineHeight: 1.5, marginBottom: 20 }}>
                    {locale === 'de'
                      ? <>Es sind allerdings noch <strong>{fallbackDialog.altFree}</strong> Plätze als <strong>{altLabel}</strong> frei. Möchtest du stattdessen als <strong>{altLabel}</strong> starten?</>
                      : <>However, there are still <strong>{fallbackDialog.altFree}</strong> spots available as <strong>{altLabel}</strong>. Would you like to join as <strong>{altLabel}</strong> instead?</>}
                  </p>
                  <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                    <button
                      className="btn btn-secondary"
                      style={{ fontSize: '0.9rem' }}
                      onClick={async () => {
                        const wunsch = fallbackDialog.wunsch;
                        setFallbackDialog(null);
                        // Wunsch beibehalten → landet auf Warteliste für den Wunsch-Typ.
                        await performRegistration(wunsch);
                      }}
                    >
                      {locale === 'de' ? `Auf ${wunschLabel}-Warteliste` : `Join ${wunschLabel} waitlist`}
                    </button>
                    <button
                      className="btn btn-primary"
                      style={{ fontSize: '0.9rem' }}
                      onClick={async () => {
                        const alt = fallbackDialog.alt;
                        setFallbackDialog(null);
                        // Preferred auf den Alt-Typ setzen, damit sowohl Anzeige
                        // als auch das Register-Payload den neuen Wunsch nutzen.
                        setPreferredStarterType(alt);
                        await performRegistration(alt);
                      }}
                    >
                      {locale === 'de' ? `Als ${altLabel} starten` : `Join as ${altLabel}`}
                    </button>
                  </div>
                </>
              );
            })()}
        </Modal>
      )}

      {/* v9.22: Modal fuer externe Email-Anmeldung */}
      {/* v18.13: Massenimport-Modal. */}
      {massImportOpen && (
        <Modal
          open={massImportOpen}
          onClose={() => { if (!massImportBusy) setMassImportOpen(false); }}
          maxWidth={600}
          padding={24}
          dismissable={!massImportBusy}
          ariaLabel={locale === 'de' ? 'Teilnehmer-Massenimport' : 'Bulk participant import'}
        >
          <h3 style={{ margin: '0 0 8px', fontSize: '1.1rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
            {locale === 'de' ? 'Teilnehmer-Massenimport' : 'Bulk participant import'}
          </h3>

          {massImportStep === 'input' && (
            <>
              <p style={{ margin: '0 0 12px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {locale === 'de'
                  ? <>Namen und/oder E-Mail-Adressen einfügen — <strong>eine Person pro Zeile</strong>. Das Tool gleicht jede Zeile mit dem Deloitte-Verzeichnis ab und zeigt dir danach eine <strong>Vorschau-Tabelle</strong> (Vorname, Nachname, Position, Standort, E-Mail) zum Prüfen, bevor angemeldet wird.</>
                  : <>Paste names and/or email addresses — <strong>one person per line</strong>. The tool matches each line against the Deloitte directory and then shows a <strong>preview table</strong> (first name, last name, position, location, email) to review before registering.</>}
              </p>
              <textarea
                className="form-input"
                value={massImportText}
                onChange={e => setMassImportText(e.target.value)}
                disabled={massImportResolving}
                rows={8}
                placeholder={'Mustermann, Max\nerika.musterfrau@deloitte.de\nMax Mustermann; max.mustermann@deloitte.de'}
                style={{ width: '100%', fontFamily: 'monospace', fontSize: '0.82rem', resize: 'vertical' }}
              />
              {massImportResolving && (
                <div style={{ marginTop: 10, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                  {locale === 'de' ? 'Verzeichnis-Abgleich läuft…' : 'Matching against the directory…'} {massImportProgress}
                </div>
              )}
              <div style={{ marginTop: 18, display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
                <button className="btn btn-secondary" onClick={() => setMassImportOpen(false)} disabled={massImportResolving}>
                  {locale === 'de' ? 'Abbrechen' : 'Cancel'}
                </button>
                <button className="btn btn-primary" onClick={resolveMassImport} disabled={massImportResolving || !massImportText.trim()}>
                  {massImportResolving ? (locale === 'de' ? 'Wird abgeglichen…' : 'Matching…') : (locale === 'de' ? 'Abgleichen & Vorschau' : 'Match & preview')}
                </button>
              </div>
            </>
          )}

          {massImportStep === 'preview' && (() => {
            const okCount = massImportRows.filter(r => r.status === 'ok').length;
            const dupCount = massImportRows.filter(r => r.status === 'duplicate').length;
            const nfCount = massImportRows.filter(r => r.status === 'notfound').length;
            const removeRow = (idx: number): void => setMassImportRows(prev => prev.filter((_, i) => i !== idx));
            const tdStyle: React.CSSProperties = { padding: '6px 8px', fontSize: '0.8rem', borderBottom: '1px solid var(--dex-gray-100)' };
            const thStyle: React.CSSProperties = { padding: '6px 8px', fontSize: '0.75rem', textAlign: 'left', color: 'var(--dex-gray-500)', borderBottom: '2px solid var(--dex-gray-200)', textTransform: 'uppercase', letterSpacing: 0.4 };
            return (
              <>
                <p style={{ margin: '0 0 10px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {locale === 'de'
                    ? <><strong>{okCount}</strong> bereit zum Anmelden{dupCount > 0 ? `, ${dupCount} Duplikat(e)` : ''}{nfCount > 0 ? `, ${nfCount} nicht gefunden` : ''}. Prüfe die Tabelle — nicht passende Zeilen kannst du entfernen.</>
                    : <><strong>{okCount}</strong> ready to register{dupCount > 0 ? `, ${dupCount} duplicate(s)` : ''}{nfCount > 0 ? `, ${nfCount} not found` : ''}. Review the table — remove rows that don&apos;t fit.</>}
                </p>
                <div style={{ maxHeight: 320, overflowY: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 8 }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr>
                        <th style={thStyle}>{locale === 'de' ? 'Vorname' : 'First name'}</th>
                        <th style={thStyle}>{locale === 'de' ? 'Nachname' : 'Last name'}</th>
                        <th style={thStyle}>{locale === 'de' ? 'Position' : 'Position'}</th>
                        <th style={thStyle}>{locale === 'de' ? 'Standort' : 'Location'}</th>
                        <th style={thStyle}>E-Mail</th>
                        <th style={thStyle}>{locale === 'de' ? 'Status' : 'Status'}</th>
                        <th style={thStyle} />
                      </tr>
                    </thead>
                    <tbody>
                      {massImportRows.map((r, idx) => (
                        <tr key={`${r.email || r.raw}-${idx}`} style={{ opacity: r.status === 'ok' ? 1 : 0.6 }}>
                          <td style={tdStyle}>{r.firstName || '–'}</td>
                          <td style={tdStyle}>{r.lastName || '–'}</td>
                          <td style={{ ...tdStyle, color: 'var(--dex-gray-600)' }}>{r.jobTitle || '–'}</td>
                          <td style={{ ...tdStyle, color: 'var(--dex-gray-600)' }}>{r.location || '–'}</td>
                          <td style={{ ...tdStyle, color: 'var(--dex-gray-600)' }}>{r.email || <span style={{ color: 'var(--dex-red, #c00)' }}>{r.raw}</span>}</td>
                          <td style={tdStyle}>
                            {r.status === 'ok' && <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-green-dark, #4a7c1f)' }}>{locale === 'de' ? 'OK' : 'OK'}</span>}
                            {r.status === 'duplicate' && <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-orange-dark, #b35a00)' }}>{locale === 'de' ? 'Duplikat' : 'Duplicate'}</span>}
                            {r.status === 'notfound' && <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-red, #c00)' }}>{locale === 'de' ? 'Nicht gefunden' : 'Not found'}</span>}
                          </td>
                          <td style={tdStyle}>
                            <button type="button" onClick={() => removeRow(idx)} disabled={massImportBusy} title={locale === 'de' ? 'Zeile entfernen' : 'Remove row'} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400)', padding: 2 }}>
                              <X size={14} />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div style={{ marginTop: 14, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
                    {locale === 'de' ? 'Benachrichtigung' : 'Notification'}
                  </div>
                  {([
                    { v: 'mail', de: 'Bestätigungsmail senden (+ Outlook-Termin)', en: 'Send confirmation email (+ Outlook invite)' },
                    { v: 'nomail', de: 'Ohne Bestätigungsmail (aber Outlook-Termin)', en: 'No confirmation email (but Outlook invite)' },
                    { v: 'silent', de: 'Stille Anmeldung (keine Mail, kein Kalendereintrag)', en: 'Silent registration (no email, no calendar invite)' },
                  ] as const).map(opt => (
                    <label key={opt.v} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                      <input type="radio" name="massImportMode" checked={massImportMode === opt.v} onChange={() => setMassImportMode(opt.v)} disabled={massImportBusy} />
                      {locale === 'de' ? opt.de : opt.en}
                    </label>
                  ))}
                </div>

                {massImportBusy && (
                  <div style={{ marginTop: 12, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                    {locale === 'de' ? 'Anmeldung läuft…' : 'Registering…'} {massImportProgress}
                  </div>
                )}
                {massImportResult && (
                  <div style={{
                    marginTop: 12, padding: '10px 12px', borderRadius: 8, fontSize: '0.82rem',
                    background: massImportResult.failed.length > 0 ? 'rgba(237,139,0,0.08)' : 'rgba(134,188,37,0.10)',
                    border: `1px solid ${massImportResult.failed.length > 0 ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)'}`,
                    color: 'var(--dex-gray-700)', lineHeight: 1.5,
                  }}>
                    {locale === 'de' ? <><strong>{massImportResult.ok}</strong> Person(en) angemeldet.</> : <><strong>{massImportResult.ok}</strong> person(s) registered.</>}
                    {massImportResult.failed.length > 0 && (
                      <div style={{ marginTop: 4 }}>
                        {locale === 'de' ? 'Nicht angemeldet (bereits angemeldet / Fehler): ' : 'Not registered (already registered / error): '}
                        {massImportResult.failed.join(', ')}
                      </div>
                    )}
                  </div>
                )}

                <div style={{ marginTop: 18, display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
                  {massImportResult ? (
                    <button className="btn btn-primary" onClick={() => setMassImportOpen(false)}>
                      {locale === 'de' ? 'Schließen' : 'Close'}
                    </button>
                  ) : (
                    <>
                      <button className="btn btn-secondary" onClick={() => { setMassImportStep('input'); setMassImportResult(null); }} disabled={massImportBusy}>
                        {locale === 'de' ? 'Zurück' : 'Back'}
                      </button>
                      <button className="btn btn-primary" onClick={runMassImport} disabled={massImportBusy || okCount === 0}>
                        {massImportBusy ? (locale === 'de' ? 'Läuft…' : 'Running…') : (locale === 'de' ? `${okCount} anmelden` : `Register ${okCount}`)}
                      </button>
                    </>
                  )}
                </div>
              </>
            );
          })()}
        </Modal>
      )}

      {/* v18.75: Sicherheitshinweis-Dialog vor dem Absenden (pro Event). */}
      {confirmDialogOpen && event && (() => {
        const isFree = event.confirmDialogMode === 'freetext';
        // v18.76: ALLE Sub-Events zeigen (auch nicht ausgewählte), damit der
        // Teilnehmer im Dialog ab- UND zuwählen kann.
        const allChildren = childEvents;
        const showParent = willRegisterParent || registerForOther;
        const parentEditable = willRegisterParent && !registerForOther; // proxy: Parent fix
        const canConfirm = isFree
          ? confirmDialogAck
          : (confirmDraftParent || confirmDraftSessions.size > 0 || (showParent && !parentEditable));
        // v18.76: Datum + Uhrzeit pro Eintrag anzeigen.
        const fmtDT = (iso?: string): string => {
          if (!iso) return '';
          try { return new Date(iso).toLocaleString(locale === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); } catch { return ''; }
        };
        const dtRange = (s?: string, e?: string): string => {
          const a = fmtDT(s); const b = fmtDT(e);
          return a && b ? `${a} – ${b}` : (a || b);
        };
        return (
          <Modal
            open={confirmDialogOpen}
            onClose={() => setConfirmDialogOpen(false)}
            maxWidth={560}
            padding={24}
            ariaLabel={locale === 'de' ? 'Anmeldung bestätigen' : 'Confirm registration'}
          >
            <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
              {locale === 'de' ? 'Bitte bestätigen' : 'Please confirm'}
            </h3>
            {isFree ? (
              <>
                <div style={{
                  margin: '0 0 14px', padding: '12px 14px', whiteSpace: 'pre-wrap',
                  background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)',
                  borderRadius: 8, fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-800)',
                }}>
                  {(event.confirmDialogText || '').trim() || (locale === 'de' ? 'Bitte bestätige deine Anmeldung.' : 'Please confirm your registration.')}
                </div>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginBottom: 4 }}>
                  <input type="checkbox" checked={confirmDialogAck} onChange={e => setConfirmDialogAck(e.target.checked)} style={{ marginTop: 3 }} />
                  <span style={{ flex: 1, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
                    {locale === 'de' ? 'Ich habe den Hinweis gelesen und bestätige.' : 'I have read and acknowledge the note.'}
                  </span>
                </label>
              </>
            ) : (() => {
              // v19.0: statt generischem „Punkte/items" den konfigurierten
              // Section-Begriff verwenden (Default „Event-Sections").
              const sectionTerm = childTermPlural || (locale === 'de' ? 'Event-Sections' : 'event sections');
              return (
              <>
                <p style={{ margin: '0 0 12px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
                  {locale === 'de'
                    ? `Du meldest dich für die angehakten ${sectionTerm} an. Du kannst vor dem Absenden einzelne ${sectionTerm} ab- oder zuwählen:`
                    : `You are registering for the checked ${sectionTerm}. You can de-/select ${sectionTerm} before submitting:`}
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 14 }}>
                  {showParent && (
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: parentEditable ? 'pointer' : 'default', padding: '8px 10px', background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)', borderRadius: 6 }}>
                      <input
                        type="checkbox"
                        checked={parentEditable ? confirmDraftParent : true}
                        disabled={!parentEditable}
                        onChange={e => setConfirmDraftParent(e.target.checked)}
                        style={{ marginTop: 2 }}
                      />
                      <span style={{ flex: 1 }}>
                        <span style={{ fontSize: '0.88rem', fontWeight: 600, display: 'block' }}>{event.title} <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>{locale === 'de' ? '(Haupt-Event)' : '(main event)'}</span></span>
                        {dtRange(event.startDate, event.endDate) && (
                          <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', display: 'block', marginTop: 1 }}>{dtRange(event.startDate, event.endDate)}</span>
                        )}
                      </span>
                    </label>
                  )}
                  {allChildren.map(ce => (
                    <label key={ce.id} style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', padding: '8px 10px', background: confirmDraftSessions.has(ce.id) ? 'rgba(134,188,37,0.06)' : 'var(--dex-gray-50, #f7f7f5)', border: `1px solid ${confirmDraftSessions.has(ce.id) ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`, borderRadius: 6 }}>
                      <input
                        type="checkbox"
                        checked={confirmDraftSessions.has(ce.id)}
                        style={{ marginTop: 2 }}
                        onChange={e => {
                          if (e.target.checked) {
                            // v18.76: Sub-Event mit eigenen Pflichtfeldern erst über
                            // das Sub-Event-Modal erfassen, damit keine leeren
                            // Pflicht-Antworten entstehen. Dialog schließen, Modal
                            // öffnen; nach dem Ausfüllen erscheint der Dialog erneut.
                            const hasCF = (ce.eventSpecificFields || []).length > 0;
                            if (hasCF && !sessionFieldValues[ce.id] && !selectedSessions.has(ce.id)) {
                              confirmDialogConfirmedRef.current = false;
                              setConfirmDialogOpen(false);
                              setPendingSubEventModal({ subEventId: ce.id, draftValues: { ...(sessionFieldValues[ce.id] || {}) } });
                            } else {
                              setConfirmDraftSessions(prev => { const n = new Set(prev); n.add(ce.id); return n; });
                            }
                          } else {
                            setConfirmDraftSessions(prev => { const n = new Set(prev); n.delete(ce.id); return n; });
                          }
                        }}
                      />
                      <span style={{ flex: 1 }}>
                        <span style={{ fontSize: '0.88rem', display: 'block' }}>{ce.title}</span>
                        {dtRange(ce.startDate, ce.endDate) && (
                          <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', display: 'block', marginTop: 1 }}>{dtRange(ce.startDate, ce.endDate)}</span>
                        )}
                      </span>
                    </label>
                  ))}
                </div>
                {!canConfirm && (
                  <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-red, #c00)' }}>
                    {locale === 'de' ? `Bitte mindestens eine ${sectionTerm} auswählen.` : `Please select at least one of the ${sectionTerm}.`}
                  </p>
                )}
              </>
            );
            })()}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
              <button className="btn btn-secondary" onClick={() => setConfirmDialogOpen(false)} style={{ fontSize: '0.85rem' }}>
                {locale === 'de' ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                className="btn btn-primary"
                disabled={!canConfirm}
                onClick={() => {
                  // v18.75: In der Auswahl-Übersicht die (ggf. angepasste)
                  // Auswahl in den echten State übernehmen, dann Submit erneut
                  // anstoßen (Ref überspringt den Dialog).
                  if (!isFree) {
                    if (parentEditable) setRegisterForParent(confirmDraftParent);
                    setSelectedSessions(new Set(confirmDraftSessions));
                  }
                  confirmDialogConfirmedRef.current = true;
                  setConfirmDialogOpen(false);
                  setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 60);
                }}
                style={{ fontSize: '0.85rem' }}
              >
                {locale === 'de' ? 'Anmeldung bestätigen' : 'Confirm registration'}
              </button>
            </div>
          </Modal>
        );
      })()}

      {externalEmailWarning && (
        <Modal
          open={externalEmailWarning}
          onClose={() => setExternalEmailWarning(false)}
          maxWidth={540}
          padding={24}
          ariaLabel={locale === 'de' ? 'E-Mail-Adresse prüfen' : 'Check the email address'}
        >
            {/* v18.74: Tippfehler-Gegenlesen — die externe Adresse groß
                anzeigen und zur Bestätigung auffordern. */}
            <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
              {locale === 'de' ? 'E-Mail-Adresse prüfen' : 'Check the email address'}
            </h3>
            <p style={{ margin: '0 0 10px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
              {locale === 'de'
                ? <>Du meldest eine <strong>externe Person</strong> an. Bitte lies die Adresse genau gegen — an externe Adressen lässt sich ein <strong>Tippfehler nachträglich nicht korrigieren</strong>:</>
                : <>You are registering an <strong>external person</strong>. Please read the address carefully — a <strong>typo cannot be corrected afterwards</strong> for external addresses:</>}
            </p>
            <div style={{
              margin: '0 0 12px', padding: '12px 14px', textAlign: 'center',
              background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)',
              borderRadius: 8, fontSize: '1.05rem', fontWeight: 700, wordBreak: 'break-all',
              color: 'var(--dex-gray-900, #222)',
            }}>
              {email}
            </div>
            <p style={{ margin: '0 0 12px', fontSize: '0.82rem', lineHeight: 1.55, color: 'var(--dex-gray-600)' }}>
              {locale === 'de'
                ? <>Die <strong>Anmeldebestätigung</strong> geht direkt an diese Adresse, mit den <strong>Organizern auf CC</strong>. Ein <strong>Outlook-Termin</strong> wird an externe Adressen nicht versendet.</>
                : <>The <strong>confirmation email</strong> is sent directly to this address, with the <strong>organizers on CC</strong>. An <strong>Outlook invite</strong> is not sent to external addresses.</>}
            </p>
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setExternalEmailWarning(false)}
                style={{ fontSize: '0.85rem' }}
              >
                {locale === 'de' ? 'Zurück, korrigieren' : 'Back, edit'}
              </button>
              <button
                className="btn btn-primary"
                onClick={() => {
                  externalEmailConfirmedRef.current = true;
                  setExternalEmailWarning(false);
                  // Re-trigger submit via short timeout
                  setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
                }}
                style={{ fontSize: '0.85rem' }}
              >
                {locale === 'de' ? 'Adresse ist korrekt' : 'Address is correct'}
              </button>
            </div>
        </Modal>
      )}

      {/* v19.6: CC-Frage bei stellvertretender INTERNER Anmeldung. Erscheint
          nach dem „Anmelden"-Klick und vor der eigentlichen Anmeldung — der/die
          Anmeldende (Organizer, Co-Organizer oder Assistenz) entscheidet, ob er/
          sie selbst auf CC der Bestätigungs-Mail soll. */}
      {ccSelfModalOpen && (
        <Modal
          open={ccSelfModalOpen}
          onClose={() => setCcSelfModalOpen(false)}
          maxWidth={520}
          padding={24}
          ariaLabel={locale === 'de' ? 'Auf Kopie der Bestätigung?' : 'Copy on the confirmation?'}
        >
          <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
            {locale === 'de' ? 'Möchtest du eine Kopie der Bestätigung?' : 'Do you want a copy of the confirmation?'}
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
            {locale === 'de'
              ? <>Du meldest {`${firstName} ${surname}`.trim() ? <strong>{`${firstName} ${surname}`.trim()}</strong> : <>die ausgewählte Person</>} stellvertretend an. Möchtest du selbst auf <strong>CC der Bestätigungs-Mail</strong> gesetzt werden? Du bekommst dann eine Kopie der Anmeldebestätigung.<br /><br />Der <strong>Outlook-Termin</strong> wird davon nicht berührt — die CC gilt nur für die Bestätigungs-Mail.</>
              : <>You are registering {`${firstName} ${surname}`.trim() ? <strong>{`${firstName} ${surname}`.trim()}</strong> : <>the selected person</>} on their behalf. Would you like to be added to the <strong>CC of the confirmation email</strong>? You will then receive a copy of the confirmation.<br /><br />The <strong>Outlook invite</strong> is not affected — the CC only applies to the confirmation email.</>}
          </p>
          <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
            <button
              className="btn btn-secondary"
              onClick={() => {
                ccSelfRef.current = false;
                ccSelfDecidedRef.current = true;
                setCcSelfModalOpen(false);
                setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              {locale === 'de' ? 'Nein, ohne CC' : 'No, without CC'}
            </button>
            <button
              className="btn btn-primary"
              onClick={() => {
                ccSelfRef.current = true;
                ccSelfDecidedRef.current = true;
                setCcSelfModalOpen(false);
                setTimeout(() => { handleSubmit().catch(() => { /* */ }); }, 50);
              }}
              style={{ fontSize: '0.85rem' }}
            >
              {locale === 'de' ? 'Ja, mich auf CC setzen' : 'Yes, add me to CC'}
            </button>
          </div>
        </Modal>
      )}

      {/* v10.12: Sub-Event Custom-Fields Modal — wird geöffnet wenn ein
          Sub-Event mit eigenen Custom-Fields angecheckt wird. Der User muss die
          Antworten ausfüllen + bestätigen, dann wandert die Session in
          selectedSessions und die Werte in sessionFieldValues. Beim „Abbrechen"
          wird die Session NICHT angecheckt. */}
      {pendingSubEventModal && (() => {
        const ce = childEvents.find(c => c.id === pendingSubEventModal.subEventId);
        if (!ce) return null;
        const fields = (ce.eventSpecificFields || []).filter(f => f && f.label);
        const draft = pendingSubEventModal.draftValues;
        const setDraft = (next: Record<string, string>): void => {
          setPendingSubEventModal(prev => prev ? { ...prev, draftValues: next } : prev);
        };
        const updateFieldValue = (fieldId: string, value: string): void => {
          setDraft({ ...draft, [fieldId]: value });
        };
        // v17.22: EN-Varianten auch im Sub-Event-Modal respektieren — geknüpft
        // an die Bilingual-Einstellung DES Sub-Events (ce), nicht des Parents.
        const useEnHere = locale === 'en' && !!ce.bilingualFields;
        const fLabel = (f: EventSpecificField): string =>
          (useEnHere && f.labelEn && f.labelEn.trim()) ? f.labelEn : f.label;
        const fHelp = (f: EventSpecificField): string | undefined =>
          (useEnHere && f.helpTextEn && f.helpTextEn.trim()) ? f.helpTextEn : f.helpText;
        const fOpt = (f: EventSpecificField, opt: string, idx: number): string =>
          (useEnHere && f.optionsEn && f.optionsEn[idx] && f.optionsEn[idx].trim()) ? f.optionsEn[idx] : opt;
        const requiredMissing = fields.filter(f => f.required && !((draft[f.id] || '').trim())).map(f => fLabel(f));
        const canSubmit = requiredMissing.length === 0;
        const onConfirm = (): void => {
          if (!canSubmit) return;
          setSessionFieldValues(prev => ({ ...prev, [ce.id]: { ...draft } }));
          setSelectedSessions(prev => {
            const next = new Set(prev);
            next.add(ce.id);
            return next;
          });
          setPendingSubEventModal(null);
        };
        const onCancel = (): void => setPendingSubEventModal(null);

        return (
          <Modal
            open={true}
            onClose={onCancel}
            maxWidth={520}
            padding={24}
            ariaLabel={ce.title || childTermSingular || (locale === 'de' ? 'Sub-Event' : 'Sub-event')}
          >
              <h3 style={{ margin: '0 0 6px', fontSize: '1.1rem' }}>
                {ce.title || childTermSingular || (locale === 'de' ? 'Sub-Event' : 'Sub-event')}
              </h3>
              <p style={{ margin: '0 0 18px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {locale === 'de'
                  ? `Bitte beantworte die Fragen für dieses ${childTermSingular || 'Sub-Event'}:`
                  : `Please answer the questions for this ${childTermSingular || 'sub-event'}:`}
              </p>

              <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginBottom: 20 }}>
                {fields.map(f => {
                  const val = draft[f.id] || '';
                  return (
                    <div key={f.id}>
                      <label className="form-label" style={{ display: 'block', fontSize: '0.85rem', marginBottom: 4 }}>
                        {fLabel(f)}
                        {f.required && <span style={{ color: 'var(--dex-red, #c00)', marginLeft: 4 }}>*</span>}
                        {/* v11.16: konsistenter InfoTooltip statt grauer
                            Inline-Beschreibung — gleicher Look wie auf
                            der Haupt-Register-Page. */}
                        {fHelp(f) && <InfoTooltip text={fHelp(f)} />}
                      </label>
                      {f.type === 'select' && f.multi ? (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          {(f.options || []).map((opt, optIdx) => {
                            const current = val.split(' | ').map(s => s.trim()).filter(Boolean);
                            const checked = current.indexOf(opt) >= 0;
                            return (
                              <label key={opt} style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={e => {
                                    const next = e.target.checked
                                      ? [...current, opt]
                                      : current.filter(x => x !== opt);
                                    updateFieldValue(f.id, next.join(' | '));
                                  }}
                                />
                                {fOpt(f, opt, optIdx)}
                              </label>
                            );
                          })}
                        </div>
                      ) : f.type === 'select' ? (
                        <select
                          className="form-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                          style={{ width: '100%', fontSize: '0.9rem' }}
                        >
                          <option value="">{locale === 'de' ? '— bitte wählen —' : '— please select —'}</option>
                          {(f.options || []).map((opt, optIdx) => <option key={opt} value={opt}>{fOpt(f, opt, optIdx)}</option>)}
                        </select>
                      ) : f.type === 'checkbox' ? (
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
                          <input
                            type="checkbox"
                            checked={val === 'true'}
                            onChange={e => updateFieldValue(f.id, e.target.checked ? 'true' : 'false')}
                          />
                          {locale === 'de' ? 'Ja' : 'Yes'}
                        </label>
                      ) : f.type === 'number' ? (
                        <input
                          type="number"
                          className="form-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                          style={{ width: '100%', fontSize: '0.9rem' }}
                        />
                      ) : (
                        <input
                          type="text"
                          className="form-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                          style={{ width: '100%', fontSize: '0.9rem' }}
                        />
                      )}
                    </div>
                  );
                })}
              </div>

              {!canSubmit && requiredMissing.length > 0 && (
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-red, #c00)', marginBottom: 12 }}>
                  {locale === 'de' ? 'Pflichtfelder fehlen: ' : 'Required fields missing: '}{requiredMissing.join(', ')}
                </div>
              )}

              <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button type="button" className="btn btn-secondary" onClick={onCancel}>
                  {locale === 'de' ? 'Abbrechen' : 'Cancel'}
                </button>
                <button type="button" className="btn btn-primary" onClick={onConfirm} disabled={!canSubmit}>
                  {locale === 'de' ? 'Bestätigen' : 'Confirm'}
                </button>
              </div>
          </Modal>
        );
      })()}
    </div>
  );
}

