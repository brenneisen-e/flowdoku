/**
 * Registrierungsseite fuer ein einzelnes Event
 *
 * Drei-Spalten-Layout: Event-Info | persoenliche Daten | eventspezifische Felder
 * Speichert die Registrierung in der SharePoint-Teilnehmerliste des Events.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { Salutation } from '../types';
import { Icon } from '@fluentui/react/lib/Icon';
import { Info, Trash2, Send } from './Icons';
import { InfoTooltip } from './InfoTooltip';
import OrganizerList from './OrganizerList';

function formatDate(iso: string): string {
  const d = new Date(iso);
  return (
    d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
    ' ' +
    d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
  );
}

export default function RegistrationPage(): React.ReactElement {
  const { selectedEventId, navigate, navIntent, clearIntent } = useNavigation();
  const { events, registerForEvent, cancelRegistration, checkRegistrationByEmail, getMyRegistration, getAllRegistrations, childEventsOf } = useEvents();
  const { currentUser } = useCurrentUser();
  const { searchUsers, isAdmin } = useRoles();
  const { t, locale } = useLanguage();
  const event = events.find(e => e.id === selectedEventId);

  // Per-Event-Organizer-Check: ist der eingeloggte User in event.organizerEmails?
  // Nur dann darf er a) nach Deadline registrieren und b) "Register for another
  // person" nutzen. Ein Organizer von EVENT A darf NICHT fuer EVENT B solche
  // Admin-Aktionen ausfuehren. Admin darf global alles.
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isEventOrganizer = !!event && event.organizerEmails.some(e => e.toLowerCase() === currentEmailLc);
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
  const [preferredStarterType, setPreferredStarterType] = React.useState<string>('');
  // Seit v6.5: Fallback-Dialog wenn B2Run-Wunschtyp voll, aber Alternative frei.
  const [fallbackDialog, setFallbackDialog] = React.useState<{ wunsch: string; alt: string; altFree: number } | null>(null);
  const [starterCounts, setStarterCounts] = React.useState<{ durch: number; fun: number } | null>(null);
  const [submitted, setSubmitted] = React.useState(false);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [error, setError] = React.useState('');
  const [showErrors, setShowErrors] = React.useState(false);
  const [showDescription, setShowDescription] = React.useState(true);
  const [thirdPartyCheck, setThirdPartyCheck] = React.useState<{ alreadyRegistered: boolean; notInAudience: boolean } | null>(null);

  // Seit v6.14: integrierte Session-Auswahl direkt auf der Registrierungsseite.
  // Der User kann auf EINER Seite wählen, ob er sich für das Haupt-Event und/oder
  // einzelne Sub-Events anmelden möchte. Bei B2Run-Parents zusätzlich pro Session
  // eine Durchstarter/Funstarter-Auswahl.
  const childEvents = React.useMemo(() => event ? childEventsOf(event.id) : [], [event?.id]);
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

  // Vorbelegen: Parent-Reg prüfen + Sessions-Meta laden (bereits-registrierte
  // Sessions werden als angehakt voreingestellt).
  React.useEffect(() => {
    if (!event || registerForOther) return;
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
    if (childEvents.length > 0) {
      (async () => {
        try {
          const meta: Record<string, { count: number; wasRegistered: boolean }> = {};
          const preselect = new Set<string>();
          const starterPre: Record<string, string> = {};
          for (const ce of childEvents) {
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
          setSessionMeta(meta);
          setSelectedSessions(preselect);
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
        setStarterCounts({
          durch: active.filter(r => r.StarterType === 'Durchstarter').length,
          fun: active.filter(r => r.StarterType === 'Funstarter').length,
        });
      } catch { /* ignore */ }
    })();
  }, [isSplitGroup, event?.subsiteUrl]);
  // Deloitte-Mitarbeitersuche
  const [userSearch, setUserSearch] = React.useState('');
  const [userResults, setUserResults] = React.useState<Array<{ email: string; displayName: string; location: string; jobTitle: string }>>([]);
  const [isSearchingUser, setIsSearchingUser] = React.useState(false);
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
  const willRegisterParent = registerForParent && !parentAlreadyRegistered && !registerForOther;
  // Fürs Registrieren für andere bleibt der alte Flow: Parent wird immer registriert,
  // keine Session-Auswahl (siehe Render).
  const isSessionsOnlyMode = !willRegisterParent && !registerForOther && !parentAlreadyRegistered;

  // v9.22: Warning-Modal fuer externe Email-Anmeldung (durch Organizer fuer
  // Drittpersonen die noch kein Deloitte-Postfach haben). Default: nicht
  // erlaubt; Organizer kann nach Bestaetigung trotzdem fortfahren — die
  // Bestaetigungsmail geht dann nicht an die externe Adresse, sondern an
  // den Organizer mit Datenschutz-Hinweis-Header.
  const [externalEmailWarning, setExternalEmailWarning] = React.useState(false);
  const externalEmailConfirmedRef = React.useRef(false);

  const handleSubmit = async (): Promise<void> => {
    // Validierung Pflichtfelder
    setShowErrors(true);

    // Wenn der Haupt-Event-Checkbox aus ist und keine Session ausgewählt ist,
    // gibt es nichts zu tun.
    if (!willRegisterParent && !registerForOther && selectedSessions.size === 0) {
      setError(t('reg.nothing.selected') || 'Bitte wähle mindestens Haupt-Event oder eine Session aus.');
      return;
    }

    // v9.22: Externe Email-Adresse bei "Für andere Person registrieren" —
    // Warnung anzeigen bevor der Anmelde-Flow startet.
    if (registerForOther && email && !externalEmailConfirmedRef.current) {
      const emLow = email.trim().toLowerCase();
      const isDel = /@(.*\.)?deloitte\.de$/.test(emLow);
      if (!isDel) {
        setExternalEmailWarning(true);
        return; // Modal zeigen, User muss bestaetigen
      }
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

    // Nur wenn der User das Haupt-Event (neu) anmelden möchte, gelten
    // Anrede + Custom-Fields + B2Run-Starter-Typ als Pflicht.
    if (willRegisterParent || registerForOther) {
      if (!salutation) {
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
          return f.type === 'checkbox'
            ? eventSpecific[f.id] !== 'true'
            : !eventSpecific[f.id]?.trim();
        });
      if (missingRequired.length > 0) {
        setError(`${t('reg.requiredcustom')}: ${missingRequired.map(f => f.label).join(', ')}`);
        return;
      }

      // B2Run: Starter-Typ Pflichtfeld
      if (isSplitGroup && !preferredStarterType) {
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
    if (isSplitGroup && selectedSessions.size > 0 && !preferredStarterType) {
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
          const fresh = await searchUsers(email.trim());
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
    if ((willRegisterParent || registerForOther) && isSplitGroup && preferredStarterType) {
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

    await performRegistration(preferredStarterType);
  };

  // Eigentliche Registrierung — entkoppelt vom Validation/Submit-Trigger,
  // damit sie auch vom Fallback-Dialog aufgerufen werden kann (mit ggf. geändertem Starter-Typ).
  const performRegistration = async (starterTypeToUse: string): Promise<void> => {
    setError('');
    setIsSubmitting(true);
    try {
      const customData: Record<string, string> = {
        salutation,
        ...eventSpecific,
      };
      const participantEmail = email.trim();
      const firstTrim = firstName.trim();
      const surnameTrim = surname.trim();

      let anySuccess = false;
      let parentOk = true;

      // 1) Haupt-Event anmelden (nur wenn Checkbox an und noch nicht angemeldet).
      if (willRegisterParent || registerForOther) {
        parentOk = await registerForEvent(
          selectedEventId!,
          customData,
          firstTrim,
          surnameTrim,
          participantEmail,
          starterTypeToUse || undefined
        );
        if (parentOk) anySuccess = true;
        else setError(t('reg.error'));
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
      // v10.15+: Sub-Event-Anmeldungen laufen auch beim Stellvertreter-Modus
      // (registerForOther) durch. registerForEvent akzeptiert ja participantFirstName/
      // -LastName/-Email als Argumente, daher kann der Assistent jede beliebige
      // Person sowohl auf das Hauptevent als auch auf alle gewählten Sub-Events
      // anmelden. Vorher war der Sub-Event-Loop hinter !registerForOther
      // versteckt — Beobachtung des Users: 'beim register for someone else kann
      // man nur fürs Main Event anmelden, nicht für die Sub-Events'. Fix.
      for (const ce of childEvents) {
        const wasReg = sessionMeta[ce.id]?.wasRegistered;
        const isSel = selectedSessions.has(ce.id);
        if (isSel && !wasReg) {
          const sType = isSplitGroup ? (inheritedStarterType || undefined) : undefined;
          // Pro-Sub-Event Custom-Field-Werte aus dem Modal-Flow (sessionFieldValues
          // wird beim Bestätigen des Sub-Event-Modals befüllt). Default: {}.
          const seFieldValues = sessionFieldValues[ce.id] || {};
          const ok = await registerForEvent(ce.id, seFieldValues, firstTrim, surnameTrim, participantEmail, sType);
          if (ok) anySuccess = true;
        } else if (!isSel && wasReg && !registerForOther) {
          // Cancel-Pfad bleibt aufs Selbst-Anmelden begrenzt: ein Stellvertreter
          // soll nicht aus Versehen einen Sub-Event-Slot des Anderen freigeben
          // weil er den Haken nicht gesetzt hat. Wer einen TN abmelden will,
          // macht das aktiv im Admin Center.
          await cancelRegistration(ce.id);
          anySuccess = true;
        }
      }

      if (anySuccess) {
        // Flag: wenn ausschließlich Sessions angemeldet/geändert wurden (kein
        // Parent diesmal oder schon vorher angemeldet), zeigen wir auf der
        // Success-Seite den Sessions-Only-Hinweis.
        setSessionsOnlySubmitted(!willRegisterParent && !registerForOther);
        setSubmitted(true);
      } else if (!parentOk) {
        // Parent-Fehler wurde schon in setError oben gesetzt.
      } else {
        setError(t('reg.error'));
      }
    } catch {
      setError(t('reg.genericerror'));
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleClear = (): void => {
    setSalutation('');
    setFirstName('');
    setSurname('');
    setEmail('');
    setEventSpecific({});
    setRegisterForOther(false);
  };

  if (submitted) {
    const sessionsOnlyHint = sessionsOnlySubmitted;
    const successHeadline = sessionsOnlyHint
      ? (t('reg.success.sessionsonly.title') || 'Für Sessions angemeldet')
      : (isFull ? t('reg.waitlisttitle') : t('reg.success'));
    const successBody = sessionsOnlyHint
      ? (t('reg.success.sessionsonly.msg') || 'Du hast dich ausschließlich für die ausgewählten Sessions angemeldet — NICHT für das Haupt-Event "{title}". Du bekommst pro Session eine separate Bestätigungsmail und einen eigenen Outlook-Kalendereintrag.').replace('{title}', event.title)
      : (isFull
          ? (registerForOther
              ? t('reg.waitlistmsg.other').replace('{name}', `${firstName} ${surname}`.trim()).replace('{title}', event.title).replace('{email}', email)
              : t('reg.waitlistmsg').replace('{title}', event.title))
          : (registerForOther
              ? t('reg.successmsg.other').replace('{name}', `${firstName} ${surname}`.trim()).replace('{title}', event.title).replace('{email}', email)
              : t('reg.successmsg').replace('{title}', event.title).replace('{email}', email)));
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '64px 32px' }}>
          <h2>{successHeadline}</h2>
          <p className="mt-8" style={{ color: 'var(--dex-gray-600)' }}>{successBody}</p>
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center' }}>
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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const renderRegField = (fRaw: any): React.ReactElement => {
    // Dynamisch Required erzwingen: bei aktivem Infoservice ist die
    // Mobilnummer Pflicht.
    let field = fRaw;
    // Mobilnummer bei aktiviertem Infoservice dynamisch zur Pflicht
    if (fRaw.id === 'b2run_mobilnummer' && eventSpecific['b2run_infoservice'] === 'true') {
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
    return (
  <div className="form-group" key={field.id}>
    {field.type !== 'checkbox' && (
      <label className="form-label">
        {field.required && <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>}
        {field.label}
        {/* v9.17: konsistenter InfoTooltip statt simples i-Icon —
            gibt schoenes Hover-Popover mit der vom Organizer
            beim Event-Anlegen hinterlegten Beschreibung. */}
        {field.helpText && <InfoTooltip text={field.helpText} />}
      </label>
    )}
    {field.type === 'select' && field.multi ? (
      // v7.11: Mehrfachauswahl als Checkbox-Liste. Werte werden
      // " | "-getrennt im selben Feld eventSpecific[field.id]
      // gespeichert (kompatibel mit Record<string,string>).
      (() => {
        const sep = ' | ';
        const raw = (eventSpecific[field.id] || '').trim();
        const selected = raw ? raw.split(sep).map(s => s.trim()).filter(Boolean) : [];
        const toggle = (opt: string): void => {
          const next = selected.indexOf(opt) >= 0
            ? selected.filter(s => s !== opt)
            : [...selected, opt];
          setEventSpecific({ ...eventSpecific, [field.id]: next.join(sep) });
        };
        const isErr = showErrors && field.required && selected.length === 0;
        return (
          <div
            style={{
              display: 'flex', flexDirection: 'column', gap: 6,
              padding: 10, borderRadius: 8,
              border: isErr ? '1px solid var(--dex-red)' : '1px solid var(--dex-gray-200)',
              background: '#fff',
            }}
          >
            {(field.options || []).map(opt => (
              <label key={opt} style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.9rem', color: 'var(--dex-gray-700)' }}>
                <input
                  type="checkbox"
                  checked={selected.indexOf(opt) >= 0}
                  onChange={() => toggle(opt)}
                />
                <span>{opt}</span>
              </label>
            ))}
            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', marginTop: 2 }}>
              {t('reg.multiselect.hint') || 'Mehrere Auswahl möglich'}
            </div>
          </div>
        );
      })()
    ) : field.type === 'select' ? (
      <select className="form-select" value={eventSpecific[field.id] || ''} onChange={e => setEventSpecific({ ...eventSpecific, [field.id]: e.target.value })} style={showErrors && field.required && !eventSpecific[field.id]?.trim() ? errorBorder : {}}>
        <option value="">{t('reg.pleaseselect')}</option>
        {field.options && field.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
      </select>
    ) : field.type === 'user' || field.type === 'roommate' ? (
      // v7.17: 'roommate' nutzt denselben Picker wie 'user' — der
      // einzige Unterschied ist dass 'roommate' beim Anmelden
      // automatisch eine Zimmerpartner-Mail an die ausgewaehlte
      // Person triggert (siehe EventContext). 'user' ist der
      // generische Personen-Picker ohne Mail-Versand.
      <UserFieldPicker
        value={eventSpecific[field.id] || ''}
        onChange={v => setEventSpecific({ ...eventSpecific, [field.id]: v })}
        searchUsers={searchUsers}
        placeholder={t('reg.userfield.placeholder')}
        errorStyle={showErrors && field.required && !eventSpecific[field.id]?.trim() ? errorBorder : {}}
        hint={field.type === 'roommate' ? t('reg.userfield.notifyhint') : undefined}
      />
    ) : field.type === 'checkbox' ? (
      <label
        style={{
          display: 'flex', alignItems: 'flex-start', gap: 8,
          cursor: 'pointer', fontSize: '0.9rem', color: 'var(--dex-gray-700)',
          padding: showErrors && field.required && eventSpecific[field.id] !== 'true' ? '6px 8px' : 0,
          border: showErrors && field.required && eventSpecific[field.id] !== 'true' ? '1px solid var(--dex-red)' : 'none',
          borderRadius: showErrors && field.required && eventSpecific[field.id] !== 'true' ? 6 : 0,
        }}
      >
        <input
          type="checkbox"
          checked={eventSpecific[field.id] === 'true'}
          onChange={e => setEventSpecific({ ...eventSpecific, [field.id]: e.target.checked ? 'true' : 'false' })}
          style={{ marginTop: 3, flexShrink: 0 }}
        />
        <span>
          {field.required && <span className="required" style={{ color: 'var(--dex-red)', marginRight: 4, fontWeight: 700 }}>*</span>}
          {field.label}
          {field.helpText && <span className="info-icon" title={field.helpText} style={{ marginLeft: 8 }}>i</span>}
          {field.externalLinks && field.externalLinks.length > 0 && (
            <span style={{ display: 'block', marginTop: 4, fontSize: '0.78rem' }}>
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
            </span>
          )}
        </span>
      </label>
    ) : (
      <input className="form-input" value={eventSpecific[field.id] || ''} onChange={e => setEventSpecific({ ...eventSpecific, [field.id]: e.target.value })} placeholder={field.label} style={showErrors && field.required && !eventSpecific[field.id]?.trim() ? errorBorder : {}} />
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
                background: event.imageUrl
                  ? 'var(--dex-gray-100)'
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
                  style={imgOrientation === 'portrait'
                    ? { width: '100%', height: '100%', objectFit: 'cover', display: 'block' }
                    : { width: '100%', height: 'auto', maxHeight: 480, objectFit: 'contain', display: 'block' }
                  }
                />
              )}
              <button className="event-card__info-btn" aria-label="Event info" onClick={() => setShowDescription(!showDescription)}>
                <Info size={16} />
              </button>
            </div>
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 6, padding: '4px 4px 4px 0' }}>
              <h4 style={{ fontSize: '1rem', margin: 0 }}>{event.title}</h4>
              <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: 0 }}>
                {formatDate(event.startDate)} {t('reg.until')}<br />
                {formatDate(event.endDate)}
              </p>
              {/* Veranstaltungsort: Name fett + strukturierte Adresse darunter */}
              {(event.location || (event.locationAddress && (event.locationAddress.street || event.locationAddress.city))) && (
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', margin: '2px 0 0' }}>
                  {event.location && (
                    <div style={{ fontWeight: 700, color: 'var(--dex-gray-700)' }}>{event.location}</div>
                  )}
                  {event.locationAddress && (event.locationAddress.street || event.locationAddress.city) && (
                    <div>
                      {[event.locationAddress.street, event.locationAddress.houseNo].filter(Boolean).join(' ')}
                      {(event.locationAddress.zip || event.locationAddress.city) && <br />}
                      {[event.locationAddress.zip, event.locationAddress.city].filter(Boolean).join(' ')}
                    </div>
                  )}
                </div>
              )}
              {(() => {
                // Organizer als Chips mit Foto (Hover-Enlarge). Namen werden von "Nachname, Vorname"
                // in "Vorname Nachname" normalisiert.
                const orgs = event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => {
                  const trimmed = o.trim();
                  const parts = trimmed.split(',').map(s => s.trim());
                  return parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
                }).filter(Boolean);
                if (orgs.length === 0) return null;
                return (
                  <div style={{ marginTop: 6 }}>
                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4 }}>Organizer</div>
                    <OrganizerList names={orgs} emails={event.organizerEmails} size="sm" />
                  </div>
                );
              })()}
              {/* v10.16: Optionaler Ansprechpartner — frei eingegebene Person
                  außerhalb des App-User-Pools. Reines Anzeige-Feld; Mailto-Link
                  wenn Email gesetzt. Wird nur gerendert wenn mindestens Name
                  oder Email gepflegt sind. */}
              {(event.contactName || event.contactEmail || event.contactInfo) && (
                <div style={{ marginTop: 12, padding: '10px 12px', background: 'var(--dex-gray-50, #f7f7f7)', borderRadius: 8, border: '1px solid var(--dex-gray-200)' }}>
                  <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4 }}>
                    {locale === 'de' ? 'Ansprechpartner' : 'Contact'}
                  </div>
                  {event.contactName && (
                    <div style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>{event.contactName}</div>
                  )}
                  {event.contactEmail && (
                    <div style={{ fontSize: '0.78rem', marginTop: 2 }}>
                      <a href={`mailto:${event.contactEmail}`} style={{ color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'none' }}>{event.contactEmail}</a>
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
                return (
                  <div style={{
                    fontSize: '0.78rem',
                    color,
                    marginTop: 6,
                    fontWeight: 600,
                  }}>
                    {isFullAll
                      ? t('reg.seats.full') || 'Event voll — Anmeldung geht auf die Warteliste'
                      : `${free} / ${event.maxParticipants} ${t('reg.seats.available') || 'Plätze frei'}`}
                  </div>
                );
              })()}
            </div>
          </div>
          {showDescription && event.description && (
            // v9.25: Beschreibung darf HTML enthalten (RichText-Editor im
            // EventCreation/Edit). Wir rendern als HTML statt Plain-Text,
            // damit Formatierung wie Listen, Links, Fett etc. funktioniert.
            // Die Description kommt aus dem eigenen Tenant — sicherer Origin.
            <div
              style={{
                padding: '12px 16px', fontSize: '0.85rem', color: 'var(--dex-gray-700)',
                background: 'var(--dex-gray-50)', borderRadius: '0 0 var(--dex-radius) var(--dex-radius)',
                borderTop: '1px solid var(--dex-gray-200)',
                wordBreak: 'break-word',
                lineHeight: 1.55,
              }}
              dangerouslySetInnerHTML={{
                __html: /<[a-z][\s\S]*>/i.test(event.description || '')
                  ? event.description
                  : (event.description || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>'),
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

              {/* Haupt-Event-Checkbox */}
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

              {/* Sessions */}
              {childEvents.length > 0 && (
                <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', fontWeight: 600 }}>{t('reg.selection.sessions') || 'Sessions'}</div>
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
                            <div style={{ fontWeight: 700 }}>{ce.title || t('reg.subevents.untitled')}</div>
                            {ce.description && (
                              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>{ce.description}</div>
                            )}
                            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {ce.startDate && <>{formatDate(ce.startDate)}</>}
                              {ce.location && <> · {ce.location}</>}
                              {hasCap && (() => {
                                // v9.8: Klartext-Anzeige damit der User auf einen Blick
                                // sieht, wie viele Plaetze noch frei sind. Vorher stand
                                // dort nur "0/25" ohne Label, was die User-Frage
                                // "warum steht da 0/25?" ausgeloest hat.
                                const sessionFree = Math.max(0, (ce.maxParticipants || 0) - (meta.count || 0));
                                return (
                                  <> · <span style={{ color: isSessionFull ? 'var(--dex-red)' : 'inherit', fontWeight: 600 }}>
                                    {meta.count}/{ce.maxParticipants} {t('reg.subevents.taken')}
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

              {isSessionsOnlyMode && selectedSessions.size > 0 && (
                <div style={{
                  marginTop: 12, padding: '8px 10px', borderRadius: 6,
                  background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                  color: 'var(--dex-orange)', fontSize: '0.78rem',
                }}>
                  {t('reg.selection.sessionsonlyhint') || 'Du meldest dich ausschließlich für Sessions an — NICHT für das Haupt-Event.'}
                </div>
              )}
            </div>
          )}
        </div>

        {/* Persoenliche Daten */}
        <div className="registration-form">
          <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
            <Icon iconName="ContactInfo" style={{ fontSize: 16 }} />
            {t('reg.personalinfo')}
          </div>
          <div style={{ padding: '24px 20px' }}>
            <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 12 }}>
              <span className="required">*</span> = {t('reg.requiredfield')}
            </p>

            {canRegisterForOther && (
              <>
                {/* v9.17: registerForOther-Toggle als unscheinbarer Text-Link
                    statt prominentem Button — die Mehrheit registriert sich
                    selbst, der Link ist nur fuer den Sonderfall gedacht. */}
                <button
                  type="button"
                  onClick={() => {
                    setRegisterForOther(!registerForOther);
                    setThirdPartyCheck(null);
                    if (!registerForOther) { setFirstName(''); setSurname(''); setEmail(''); setUserSearch(''); setUserResults([]); }
                    else { setFirstName(currentUser.firstName); setSurname(currentUser.surname); setEmail(currentUser.email); setUserSearch(''); setUserResults([]); }
                  }}
                  style={{
                    background: 'none', border: 'none', padding: 0,
                    color: 'var(--dex-green-dark)', fontSize: '0.78rem',
                    textDecoration: 'underline', cursor: 'pointer',
                    marginBottom: 16, display: 'inline-block',
                  }}
                >
                  {registerForOther ? t('reg.registerself') : t('reg.registerother')}
                </button>
                {registerForOther && isAssistant && !canCreateEvents && (
                  <div style={{
                    padding: '8px 12px', marginBottom: 12, borderRadius: 'var(--dex-radius-md)',
                    background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                    color: 'var(--dex-orange)', fontSize: '0.8rem',
                  }}>
                    As an Assistant you can only register <strong>Partners</strong> or <strong>Directors</strong> for this event.
                  </div>
                )}
                {registerForOther && (
                  <div className="form-group" style={{ position: 'relative', marginBottom: 20 }}>
                    <label className="form-label">{t('reg.searchemployee') || 'Deloitte Mitarbeiter suchen'}</label>
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
                            const results = await searchUsers(val);
                            setUserResults(results);
                            setIsSearchingUser(false);
                          }, 300);
                        } else {
                          setUserResults([]);
                        }
                      }}
                      placeholder={t('reg.searchplaceholder') || 'Name oder E-Mail eingeben...'}
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
                                    setThirdPartyCheck({ alreadyRegistered, notInAudience });
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
                {registerForOther && thirdPartyCheck && (thirdPartyCheck.alreadyRegistered || thirdPartyCheck.notInAudience) && (
                  <div style={{
                    padding: '10px 14px', marginBottom: 16, borderRadius: 'var(--dex-radius-md)',
                    background: thirdPartyCheck.alreadyRegistered ? 'rgba(200,30,30,0.07)' : 'rgba(237,139,0,0.08)',
                    border: `1px solid ${thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)'}`,
                    color: thirdPartyCheck.alreadyRegistered ? 'var(--dex-red)' : 'var(--dex-orange)',
                    fontSize: '0.85rem',
                  }}>
                    {thirdPartyCheck.alreadyRegistered && (
                      <div><strong>{t('reg.thirdparty.alreadyregistered')}</strong></div>
                    )}
                    {thirdPartyCheck.notInAudience && (
                      <div style={{ marginTop: thirdPartyCheck.alreadyRegistered ? 6 : 0 }}>
                        <strong>{t('reg.thirdparty.notinaudience')}</strong>
                      </div>
                    )}
                  </div>
                )}
              </>
            )}

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

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> {t('reg.firstname')}</label>
              <input className="form-input" value={firstName} onChange={e => { if (registerForOther) setFirstName(e.target.value); }} placeholder={t('reg.firstname')} disabled={!registerForOther} style={{ background: 'var(--dex-gray-100)', ...(showErrors && !firstName.trim() ? errorBorder : {}) }} />
            </div>

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> {t('reg.surname')}</label>
              <input className="form-input" value={surname} onChange={e => { if (registerForOther) setSurname(e.target.value); }} placeholder={t('reg.surname')} disabled={!registerForOther} style={{ background: 'var(--dex-gray-100)', ...(showErrors && !surname.trim() ? errorBorder : {}) }} />
            </div>

            <div className="form-group">
              <label className="form-label"><span className="required">*</span> {t('reg.email')}</label>
              <input className="form-input" type="email" value={email} onChange={e => { if (registerForOther) setEmail(e.target.value); }} placeholder="email@deloitte.de" disabled={!registerForOther} style={{ background: 'var(--dex-gray-100)', ...(showErrors && !email.trim() ? errorBorder : {}) }} />
            </div>
          </div>
        </div>

        {/* Eventspezifische Felder (inkl. Split-Capacity Starter-Typ-Auswahl wenn
            beide Kapazitaeten > 0; bei nur einem verfuegbaren Typ wird dieser
            automatisch gesetzt und gar nicht angezeigt). v10.20: Sessions-/
            Hauptevent-Auswahl ist hierher gewandert (vorher links unter der
            Event-Karte). */}
        <div className="registration-specific">
          <div className="section-header" style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
            <Icon iconName="EditNote" style={{ fontSize: 16 }} />
            {t('reg.eventinfo')}
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
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                  {([
                    { id: 'Durchstarter', label: splitLabelA, desc: splitLabelA === 'Durchstarter' ? t('reg.starter.durch.desc') : '', cap: durchCap, count: starterCounts?.durch ?? 0, color: 'var(--dex-green-dark, #6b9a1e)' },
                    { id: 'Funstarter', label: splitLabelB, desc: splitLabelB === 'Funstarter' ? t('reg.starter.fun.desc') : '', cap: funCap, count: starterCounts?.fun ?? 0, color: 'var(--dex-orange, #ff8c00)' },
                  ]).map(opt => {
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
                            <span style={{ color: opt.color }}>{`${free} / ${opt.cap} ${t('reg.starter.free')}`}</span>
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
                      {groupSpec.map(renderRegField)}
                    </div>
                  );
                })()}
              </div>
            )}

            {/* v11.10: Sub-Events-Auswahl als eigener Block — nur bei
                Self-Registration mit Sub-Events. Bei „Für andere
                registrieren" bleibt der alte Flow (nur Parent), weil die
                Session-Zuordnung über getMyRegistration nur für den
                eingeloggten User funktioniert. Der Parent-Event-Checkbox
                ist hier weiterhin enthalten — er steuert, ob das Haupt-
                Event registriert werden soll. Hardcoded Pro-Sub-Event-
                Gruppen-Radios sind weg, Sub-Events erben grundsätzlich
                preferredStarterType vom Group-Selection-Block oben. */}
            {childEvents.length > 0 && !registerForOther && (
              <div style={{ marginBottom: 20, border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 16 }}>
                <h4 style={{ marginTop: 0, marginBottom: 4, fontSize: '0.95rem' }}>{t('reg.selection.title') || 'Wofür möchtest du dich anmelden?'}</h4>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                  {t('reg.selection.hint') || 'Haupt-Event und Sessions können unabhängig voneinander an- oder abgewählt werden.'}
                </p>

                {/* Haupt-Event-Checkbox */}
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

                {/* Sessions */}
                <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', fontWeight: 600 }}>{t('reg.selection.sessions') || 'Sessions'}</div>
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
                            <div style={{ fontWeight: 700 }}>{ce.title || t('reg.subevents.untitled')}</div>
                            {ce.description && (
                              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>{ce.description}</div>
                            )}
                            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {ce.startDate && <>{formatDate(ce.startDate)}</>}
                              {ce.location && <> · {ce.location}</>}
                              {hasCap && (() => {
                                const sessionFree = Math.max(0, (ce.maxParticipants || 0) - (meta.count || 0));
                                return (
                                  <> · <span style={{ color: isSessionFull ? 'var(--dex-red)' : 'inherit', fontWeight: 600 }}>
                                    {meta.count}/{ce.maxParticipants} {t('reg.subevents.taken')}
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

                {isSessionsOnlyMode && selectedSessions.size > 0 && (
                  <div style={{
                    marginTop: 12, padding: '8px 10px', borderRadius: 6,
                    background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                    color: 'var(--dex-orange)', fontSize: '0.78rem',
                  }}>
                    {t('reg.selection.sessionsonlyhint') || 'Du meldest dich ausschließlich für Sessions an — NICHT für das Haupt-Event.'}
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
                .map(renderRegField)
              }
              </div>
            )}
          </div>
        </div>
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
        <button className="btn btn-danger" onClick={handleClear} disabled={isSubmitting}><Trash2 size={16} /> {t('reg.delete')}</button>
        <button className="btn btn-primary" onClick={handleSubmit} disabled={isSubmitting}>
          <Send size={16} /> {(() => {
            if (isSubmitting) return t('reg.submitting');
            if (registerForOther) return t('reg.register');
            // v7.3: Kein Selection-Block → einfacher "Registrieren"-Text ohne
            // Parantheses-Info. Erst wenn Sub-Events existieren, zeigen wir
            // detailliert an, was gerade submittet wird.
            if (childEvents.length === 0) return t('reg.register');
            const parts: string[] = [];
            if (willRegisterParent) parts.push(t('reg.selection.mainevent') || 'Haupt-Event');
            if (selectedSessions.size > 0) {
              parts.push(`${selectedSessions.size} ${selectedSessions.size === 1 ? (t('reg.selection.sessioncount.one') || 'Session') : (t('reg.selection.sessioncount.many') || 'Sessions')}`);
            }
            if (parts.length === 0) return t('reg.register');
            return `${t('reg.register')} (${parts.join(' + ')})`;
          })()}
        </button>
      </div>

      {/* Datenschutz-Hinweis als Fußnote ganz unten */}
      <div className="footer-disclaimer mt-24" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
        <p>
          {t('reg.privacy').replace('{title}', event.title)}
        </p>
      </div>

      {/* Fallback-Dialog (seit v6.5): Wunsch-Starter-Typ voll, aber Alternative frei.
          User entscheidet explizit zwischen Umsteigen oder Warteliste. */}
      {fallbackDialog && (
        <div
          role="dialog"
          aria-modal="true"
          style={{
            position: 'fixed', inset: 0, zIndex: 1000,
            background: 'rgba(0,0,0,0.5)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            padding: 16,
          }}
          onClick={() => setFallbackDialog(null)}
        >
          <div
            onClick={e => e.stopPropagation()}
            style={{
              background: '#fff', borderRadius: 'var(--dex-radius, 12px)',
              padding: 24, maxWidth: 480, width: '100%',
              boxShadow: '0 10px 40px rgba(0,0,0,0.25)',
            }}
          >
            {(() => {
              // v10.20: Label-Mapping fuer die freie Bezeichnung — wunsch/alt
              // sind interne IDs ('Durchstarter' / 'Funstarter'); die Anzeige
              // nimmt splitLabelA / splitLabelB.
              const wunschLabel = fallbackDialog.wunsch === 'Durchstarter' ? splitLabelA : splitLabelB;
              const altLabel = fallbackDialog.alt === 'Durchstarter' ? splitLabelA : splitLabelB;
              return (
                <>
                  <h3 style={{ margin: 0, marginBottom: 10 }}>
                    {wunschLabel}-Plätze sind voll
                  </h3>
                  <p style={{ color: 'var(--dex-gray-700)', lineHeight: 1.5, marginBottom: 8 }}>
                    Für <strong>{wunschLabel}</strong> gibt es aktuell keine freien Plätze mehr.
                  </p>
                  <p style={{ color: 'var(--dex-gray-700)', lineHeight: 1.5, marginBottom: 20 }}>
                    Es sind allerdings noch <strong>{fallbackDialog.altFree}</strong> Plätze als <strong>{altLabel}</strong> frei.
                    Möchtest du stattdessen als <strong>{altLabel}</strong> starten?
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
                      Auf {wunschLabel}-Warteliste
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
                      Als {altLabel} starten
                    </button>
                  </div>
                </>
              );
            })()}
          </div>
        </div>
      )}

      {/* v9.22: Modal fuer externe Email-Anmeldung */}
      {externalEmailWarning && (
        <div
          role="dialog"
          aria-modal="true"
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 2000,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
          }}
          onClick={() => setExternalEmailWarning(false)}
        >
          <div
            onClick={e => e.stopPropagation()}
            style={{
              background: '#fff', borderRadius: 'var(--dex-radius, 8px)',
              maxWidth: 540, width: '100%', padding: 24,
              boxShadow: 'var(--dex-shadow-hover)',
            }}
          >
            <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
              Externe E-Mail-Adresse
            </h3>
            <p style={{ margin: '0 0 12px', fontSize: '0.9rem', lineHeight: 1.55, color: 'var(--dex-gray-700)' }}>
              Die Adresse <strong>{email}</strong> gehört nicht zum Deloitte-Deutschland-Tenant (@deloitte.de).
              Standardmäßig sind Anmeldungen für externe Personen nicht vorgesehen — die Plattform
              ist nur für DEALL-Mitarbeiter freigeschaltet.
            </p>
            <p style={{ margin: '0 0 12px', fontSize: '0.85rem', lineHeight: 1.55, color: 'var(--dex-gray-600)' }}>
              Wenn du diese Person trotzdem als Teilnehmer erfassen möchtest (z.B. neue Mitarbeiter
              die noch nicht angestellt sind, externe Berater die am Event teilnehmen), kannst du fortfahren.
              Die Bestätigungsmail wird dann <strong>nicht an die externe Adresse</strong> versendet,
              sondern landet bei dir als Organizer in der Inbox — du kannst sie unter Beachtung
              der <strong>Deloitte-Datenschutzrichtlinien</strong> ggf. weiterleiten.
            </p>
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setExternalEmailWarning(false)}
                style={{ fontSize: '0.85rem' }}
              >
                Abbrechen
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
                Trotzdem anmelden
              </button>
            </div>
          </div>
        </div>
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
        const requiredMissing = fields.filter(f => f.required && !((draft[f.id] || '').trim())).map(f => f.label);
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
          <div
            role="dialog"
            aria-modal="true"
            style={{
              position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 2000,
              display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
            }}
            onClick={onCancel}
          >
            <div
              onClick={e => e.stopPropagation()}
              className="card"
              style={{
                width: '100%', maxWidth: 520, maxHeight: '85vh', overflow: 'auto',
                padding: 24, borderRadius: 'var(--dex-radius, 8px)', background: '#fff',
              }}
            >
              <h3 style={{ margin: '0 0 6px', fontSize: '1.1rem' }}>
                {ce.title || (locale === 'de' ? 'Sub-Event' : 'Sub-event')}
              </h3>
              <p style={{ margin: '0 0 18px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {locale === 'de'
                  ? 'Bitte beantworte die Fragen für dieses Sub-Event:'
                  : 'Please answer the questions for this sub-event:'}
              </p>

              <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginBottom: 20 }}>
                {fields.map(f => {
                  const val = draft[f.id] || '';
                  return (
                    <div key={f.id}>
                      <label className="form-label" style={{ display: 'block', fontSize: '0.85rem', marginBottom: 4 }}>
                        {f.label}
                        {f.required && <span style={{ color: 'var(--dex-red, #c00)', marginLeft: 4 }}>*</span>}
                      </label>
                      {f.helpText && (
                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>{f.helpText}</div>
                      )}
                      {f.type === 'select' && f.multi ? (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          {(f.options || []).map(opt => {
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
                                {opt}
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
                          {(f.options || []).map(opt => <option key={opt} value={opt}>{opt}</option>)}
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
            </div>
          </div>
        );
      })()}
    </div>
  );
}

function UserFieldPicker(props: {
  value: string;
  onChange: (v: string) => void;
  searchUsers: (q: string) => Promise<Array<{ email: string; displayName: string; location?: string }>>;
  placeholder: string;
  errorStyle: React.CSSProperties;
  hint?: string;
}): React.ReactElement {
  const [query, setQuery] = React.useState(props.value);
  const [results, setResults] = React.useState<Array<{ email: string; displayName: string; location?: string }>>([]);
  const [isSearching, setIsSearching] = React.useState(false);
  const [hasSelection, setHasSelection] = React.useState(/.+ <.+@.+>/.test(props.value));
  const timerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  React.useEffect(() => {
    setQuery(props.value);
    setHasSelection(/.+ <.+@.+>/.test(props.value));
  }, [props.value]);
  return (
    <div style={{ position: 'relative' }}>
      <input
        className="form-input"
        value={query}
        onChange={e => {
          const val = e.target.value;
          setQuery(val);
          // Freitext ist kein gueltiger Wert; erst nach Dropdown-Auswahl.
          setHasSelection(false);
          props.onChange('');
          if (timerRef.current) clearTimeout(timerRef.current);
          if (val.length >= 2) {
            timerRef.current = setTimeout(async () => {
              setIsSearching(true);
              try { setResults(await props.searchUsers(val)); }
              catch { setResults([]); }
              setIsSearching(false);
            }, 300);
          } else {
            setResults([]);
          }
        }}
        onBlur={() => {
          // Wenn keine gueltige Person ausgewaehlt wurde, Feld leeren.
          setTimeout(() => {
            if (!hasSelection) {
              setQuery('');
              props.onChange('');
              setResults([]);
            }
          }, 150);
        }}
        placeholder={props.placeholder}
        style={props.errorStyle}
      />
      {isSearching && (
        <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>Suche...</div>
      )}
      {props.hint && (
        <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 4, marginBottom: 0, fontStyle: 'italic' }}>
          {props.hint}
        </p>
      )}
      {results.length > 0 && (
        <div style={{
          position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
          background: '#fff', border: '1px solid var(--dex-gray-200)',
          borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
          maxHeight: 200, overflowY: 'auto',
        }}>
          {results.map(u => (
            <div
              key={u.email}
              style={{
                padding: '8px 12px', cursor: 'pointer', fontSize: '0.85rem',
                borderBottom: '1px solid var(--dex-gray-100)',
              }}
              onMouseDown={() => {
                const formatted = `${u.displayName} <${u.email}>`;
                setQuery(formatted);
                props.onChange(formatted);
                setHasSelection(true);
                setResults([]);
              }}
            >
              <div style={{ fontWeight: 600 }}>{u.displayName}</div>
              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>{u.email}{u.location ? ` · ${u.location}` : ''}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
