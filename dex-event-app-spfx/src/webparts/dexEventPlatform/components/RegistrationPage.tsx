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
  const { t } = useLanguage();
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
  const [sessionMeta, setSessionMeta] = React.useState<Record<string, { count: number; wasRegistered: boolean }>>({});
  const [myParentReg, setMyParentReg] = React.useState<{ Status?: string } | null>(null);
  const [sessionsOnlySubmitted, setSessionsOnlySubmitted] = React.useState(false);

  // Vorbelegen: Parent-Reg prüfen + Sessions-Meta laden (bereits-registrierte
  // Sessions werden als angehakt voreingestellt).
  React.useEffect(() => {
    if (!event || registerForOther) return;
    (async () => {
      try {
        const r = await getMyRegistration(event.id);
        setMyParentReg(r);
        if (r && r.Status !== 'Abgemeldet') setRegisterForParent(false);
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
  const isB2runSplit = !!event && durchCap > 0 && funCap > 0;
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
    if (!isB2runSplit || !event?.subsiteUrl) return;
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
  }, [isB2runSplit, event?.subsiteUrl]);
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

  const handleSubmit = async (): Promise<void> => {
    // Validierung Pflichtfelder
    setShowErrors(true);

    // Wenn der Haupt-Event-Checkbox aus ist und keine Session ausgewählt ist,
    // gibt es nichts zu tun.
    if (!willRegisterParent && !registerForOther && selectedSessions.size === 0) {
      setError(t('reg.nothing.selected') || 'Bitte wähle mindestens Haupt-Event oder eine Session aus.');
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
      if (isB2runSplit && !preferredStarterType) {
        setError(t('reg.starter.required'));
        return;
      }

      // v6.15: Leistungsnachweis-Pflicht bei Durchstarter (Admin-Option)
      if (event.durchstarterRequiresProof && preferredStarterType === 'Durchstarter' && eventSpecific['b2run_leistungsnachweis'] !== 'true') {
        setError(t('reg.starter.proof.required') || 'Bitte Leistungsnachweis bestätigen.');
        return;
      }
    }

    // B2Run-Parent: jede AUSGEWÄHLTE Session muss auch einen Starter-Typ gewählt haben.
    // Ausnahme: wenn der User sich gleichzeitig fürs Haupt-Event anmeldet, erben
    // die Sessions automatisch den Haupt-Event-Starter-Typ — dann keine Extra-Abfrage.
    const sharedStarterTypeFromParent = (willRegisterParent || registerForOther) ? preferredStarterType : '';
    if (isB2runSplit && !sharedStarterTypeFromParent && selectedSessions.size > 0) {
      const missingStarter = Array.from(selectedSessions).some(sid => !sessionStarterType[sid]);
      if (missingStarter) {
        setError(t('reg.sessions.starter.required') || 'Bitte für jede ausgewählte Session den Starter-Typ wählen.');
        return;
      }
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
    if ((willRegisterParent || registerForOther) && isB2runSplit && preferredStarterType) {
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
      const inheritedStarterType = (willRegisterParent || registerForOther) ? (starterTypeToUse || preferredStarterType) : '';
      if (!registerForOther) {
        for (const ce of childEvents) {
          const wasReg = sessionMeta[ce.id]?.wasRegistered;
          const isSel = selectedSessions.has(ce.id);
          if (isSel && !wasReg) {
            const sType = (isB2runSplit && inheritedStarterType) ? inheritedStarterType : (sessionStarterType[ce.id] || undefined);
            const ok = await registerForEvent(ce.id, {}, firstTrim, surnameTrim, participantEmail, sType);
            if (ok) anySuccess = true;
          } else if (!isSel && wasReg) {
            await cancelRegistration(ce.id);
            anySuccess = true;
          }
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
          {event && event.audienceFilter && event.audienceFilter.length > 0 && <> {t('reg.audience')}: <strong>{event.audienceFilter.join(', ')}</strong>.</>}
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
          <div className="section-header section-header--red">{t('reg.selectedevent')}</div>
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
            <div style={{
              padding: '12px 16px', fontSize: '0.85rem', color: 'var(--dex-gray-700)',
              background: 'var(--dex-gray-50)', borderRadius: '0 0 var(--dex-radius) var(--dex-radius)',
              borderTop: '1px solid var(--dex-gray-200)',
              // pre-wrap: erhaelt Zeilenumbrueche aus dem Description-Textarea
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-word',
            }}>
              {event.description}
            </div>
          )}
          {isFull && (
            <p className="text-red text-center mt-8" style={{ padding: '0 12px 12px', fontWeight: 600, fontSize: '0.85rem' }}>
              {t('reg.allplaces').replace('{count}', String(event.waitlistCount))}
            </p>
          )}
          {/* v6.14: Integrierte Event-Auswahl — Parent + Sessions auf einer Seite.
              Bei "für andere Person anmelden" bleibt der alte Flow (nur Parent, keine
              Session-Auswahl), weil die Session-Zuordnung über getMyRegistration
              nur für den eingeloggten User funktioniert.
              Checkbox "Hauptevent anmelden" erscheint immer, ist bei bereits
              Angemeldeten automatisch aus + disabled. Sessions erscheinen mit
              eigener Checkbox; bei B2Run-Parents nur dann mit Durchstarter-/
              Funstarter-Auswahl, wenn der User sich NICHT gleichzeitig fürs
              Haupt-Event anmeldet (sonst wird der Haupt-Event-Starter-Typ
              automatisch auf die Session-Anmeldung übernommen). */}
          {/* v7.3: Selection-Block nur rendern, wenn es tatsächlich Sub-Events
              gibt. Bei einem Event ohne Sessions ist die Checkbox "Haupt-Event"
              redundant (es gibt nichts alternatives zum Abwählen), also
              komplett weglassen. */}
          {!registerForOther && childEvents.length > 0 && (
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
                    const inheritsStarter = isB2runSplit && (willRegisterParent || registerForOther);
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
                              const next = new Set(selectedSessions);
                              if (e.target.checked) next.add(ce.id); else next.delete(ce.id);
                              setSelectedSessions(next);
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
                              {hasCap && (
                                <> · <span style={{ color: isSessionFull ? 'var(--dex-red)' : 'inherit' }}>{meta.count}/{ce.maxParticipants}</span></>
                              )}
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
                            {/* B2Run Starter-Typ pro Session — nur wenn NICHT vom Parent geerbt */}
                            {isSel && isB2runSplit && !inheritsStarter && (
                              <div style={{ marginTop: 8, display: 'flex', gap: 10, fontSize: '0.8rem' }}>
                                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
                                  <input
                                    type="radio"
                                    name={`starter-${ce.id}`}
                                    checked={sType === 'Durchstarter'}
                                    onChange={() => setSessionStarterType({ ...sessionStarterType, [ce.id]: 'Durchstarter' })}
                                  />
                                  Durchstarter
                                </label>
                                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
                                  <input
                                    type="radio"
                                    name={`starter-${ce.id}`}
                                    checked={sType === 'Funstarter'}
                                    onChange={() => setSessionStarterType({ ...sessionStarterType, [ce.id]: 'Funstarter' })}
                                  />
                                  Funstarter
                                </label>
                              </div>
                            )}
                            {isSel && isB2runSplit && inheritsStarter && preferredStarterType && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 4, fontStyle: 'italic' }}>
                                {(t('reg.selection.starterinherited') || 'Starter-Typ wird vom Haupt-Event übernommen').replace('{type}', preferredStarterType)}
                              </div>
                            )}
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
          <div className="section-header">{t('reg.personalinfo')}</div>
          <div style={{ padding: '24px 20px' }}>
            <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', marginBottom: 12 }}>
              <span className="required">*</span> = {t('reg.requiredfield')}
            </p>

            {canRegisterForOther && (
              <>
                <button
                  className="btn btn-outline"
                  style={{ marginBottom: 20, fontSize: '0.85rem' }}
                  onClick={() => {
                    setRegisterForOther(!registerForOther);
                    setThirdPartyCheck(null);
                    if (!registerForOther) { setFirstName(''); setSurname(''); setEmail(''); setUserSearch(''); setUserResults([]); }
                    else { setFirstName(currentUser.firstName); setSurname(currentUser.surname); setEmail(currentUser.email); setUserSearch(''); setUserResults([]); }
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
                              <strong>{u.displayName}</strong>
                              <span style={{ color: 'var(--dex-gray-400)', marginLeft: 8 }}>{u.email}</span>
                              {u.jobTitle && <span style={{ color: 'var(--dex-gray-500)', marginLeft: 8, fontSize: '0.78rem', fontStyle: 'italic' }}>{u.jobTitle}</span>}
                              {u.location && <span style={{ color: 'var(--dex-gray-400)', marginLeft: 8, fontSize: '0.8rem' }}>({u.location})</span>}
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
                <option value="Herr">Herr</option>
                <option value="Frau">Frau</option>
                <option value="Divers">Divers</option>
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

        {/* Eventspezifische Felder (inkl. B2Run Starter-Typ-Auswahl wenn beide
            Kapazitaeten > 0; bei nur einem verfuegbaren Typ wird dieser automatisch
            gesetzt und gar nicht angezeigt). */}
        <div className="registration-specific">
          <div className="section-header">{t('reg.eventinfo')}</div>
          <div style={{ padding: '24px 20px' }}>
            {isB2runSplit && (
              <div className="form-group" style={{ marginBottom: 20 }}>
                <label className="form-label" style={{ fontWeight: 700, marginBottom: 6 }}>
                  <span className="required">*</span> {t('reg.starter.title')}
                </label>
                <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 10 }}>
                  {t('reg.starter.hint')}
                </p>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                  {([
                    { id: 'Durchstarter', label: t('reg.starter.durch'), desc: t('reg.starter.durch.desc'), cap: durchCap, count: starterCounts?.durch ?? 0, color: 'var(--dex-green-dark, #6b9a1e)' },
                    { id: 'Funstarter', label: t('reg.starter.fun'), desc: t('reg.starter.fun.desc'), cap: funCap, count: starterCounts?.fun ?? 0, color: 'var(--dex-orange, #ff8c00)' },
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

                {/* v6.15: Leistungsnachweis-Pflicht bei Durchstarter. Admin hat
                    die Option pro Event aktiviert — User muss dann beim Wählen
                    von Durchstarter bestätigen, dass ein Leistungsnachweis
                    vorliegt. Ohne Bestätigung wird die Anmeldung blockiert. */}
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
              </div>
            )}
            {event.eventSpecificFields.length === 0 && !isB2runSplit ? (
              <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{t('reg.noadditional')}</p>
            ) : (
              event.eventSpecificFields
                // B2Run-Sonderregel: Mobilnummer nur zeigen wenn Infoservice aktiviert
                .filter(f => f.id !== 'b2run_mobilnummer' || eventSpecific['b2run_infoservice'] === 'true')
                // v6.15: Startblock ausblenden, wenn er automatisch aus dem Starter-Typ abgeleitet wird
                .filter(f => !(f.id === 'b2run_startblock' && hasStarterBlockMapping))
                .map(fRaw => {
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
                      {field.helpText && <span className="info-icon" title={field.helpText} style={{ marginLeft: 8 }}>i</span>}
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
                  ) : field.type === 'user' ? (
                    <UserFieldPicker
                      value={eventSpecific[field.id] || ''}
                      onChange={v => setEventSpecific({ ...eventSpecific, [field.id]: v })}
                      searchUsers={searchUsers}
                      placeholder={t('reg.userfield.placeholder')}
                      errorStyle={showErrors && field.required && !eventSpecific[field.id]?.trim() ? errorBorder : {}}
                      hint={t('reg.userfield.notifyhint')}
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
                })
            )}
          </div>
        </div>
      </div>

      {/* Datenschutz-Hinweis */}
      <div className="footer-disclaimer mt-24" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
        <p>
          {t('reg.privacy').replace('{title}', event.title)}
        </p>
      </div>

      {/* Fehlermeldung */}
      {error && (
        <div className="mt-16" style={{ padding: '12px 16px', background: 'rgba(239,68,68,0.1)', border: '1px solid var(--dex-red)', borderRadius: 'var(--dex-radius-md)', color: 'var(--dex-red)', fontSize: '0.9rem' }}>
          {error}
        </div>
      )}

      {/* Buttons */}
      <div className="registration-actions mt-24">
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
            <h3 style={{ margin: 0, marginBottom: 10 }}>
              {fallbackDialog.wunsch}-Plätze sind voll
            </h3>
            <p style={{ color: 'var(--dex-gray-700)', lineHeight: 1.5, marginBottom: 8 }}>
              Für <strong>{fallbackDialog.wunsch}</strong> gibt es aktuell keine freien Plätze mehr.
            </p>
            <p style={{ color: 'var(--dex-gray-700)', lineHeight: 1.5, marginBottom: 20 }}>
              Es sind allerdings noch <strong>{fallbackDialog.altFree}</strong> Plätze als <strong>{fallbackDialog.alt}</strong> frei.
              Möchtest du stattdessen als <strong>{fallbackDialog.alt}</strong> starten?
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
                Auf {fallbackDialog.wunsch}-Warteliste
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
                Als {fallbackDialog.alt} starten
              </button>
            </div>
          </div>
        </div>
      )}
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
