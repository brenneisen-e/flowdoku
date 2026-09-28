/**
 * v22.21: Geführtes Tutorial (Onboarding-Tour).
 *
 * - `TutorialProvider` stellt `openTutorial()` bereit (Landing-Page-Button).
 *   Gibt es nur die User-Tour, startet sie direkt; Organizer/Admins bekommen
 *   zuerst eine Tour-Auswahl (User- ODER Organizer-Tutorial).
 * - Das Overlay navigiert pro Schritt auf die richtige Seite, sucht das
 *   Ziel-Element (CSS-Selektor, mit Polling weil Seiten lazy laden) und legt
 *   einen Spotlight darüber: abgedunkelter Backdrop mit „Loch" über dem
 *   Element (Box-Shadow-Trick) plus Schritt-Karte daneben. Ohne Ziel-Element
 *   erscheint die Karte zentriert.
 * - z-index 10800 — bewusst ÜBER dem shared Modal (9999), damit die Tour
 *   nie hinter App-Modals verschwindet.
 */

import * as React from 'react';
import { useNavigation } from '../../context/NavigationContext';
import { useLanguage } from '../../context/LanguageContext';
import { useRoles } from '../../context/RoleContext';
import { useEvents } from '../../context/EventContext';
import { useCurrentUser } from '../../context/UserContext';
import Modal from '../Modal';
import { TUTORIAL_TOURS, TutorialTour, TutorialTourId, TutorialStep } from './tutorialTours';
import { WizardCoach, CoachSnapshot, COACH_STATIONS, StationBase } from './WizardCoach';

/** v32.1.0: Eigener Entwurfsschlüssel des Mitmach-Tutorials — sonst
 *  überschriebe der Autosave einen echten, unfertigen Entwurf. */
export const COACH_DRAFT_KEY = 'dex_event_creation_draft_tutorial_v1';

const TOUR_Z_INDEX = 10800;
const CARD_WIDTH = 380;

interface TutorialContextType {
  /** Öffnet das Tutorial — bei mehreren verfügbaren Touren zuerst die Auswahl. */
  openTutorial: () => void;
  startTour: (id: TutorialTourId) => void;
  availableTours: TutorialTourId[];
  /** v32.1.0: Mitmach-Tutorial (WizardCoach) — darf die Person es starten? */
  canCoach: boolean;
  /** Läuft das Mitmach-Tutorial gerade? Der Assistent schaltet dann in den
   *  Tutorial-Modus (eigener Entwurf, Sicherheitsnetz beim Anlegen). */
  coachActive: boolean;
  startCoach: () => void;
  stopCoach: () => void;
  /** Der Assistent meldet seinen Zustand; gesetzt wird nur bei Änderung. */
  reportWizard: (snap: CoachSnapshot | null) => void;
}

const TutorialContext = React.createContext<TutorialContextType | undefined>(undefined);

export function useTutorial(): TutorialContextType {
  const ctx = React.useContext(TutorialContext);
  if (!ctx) {
    // Fallback (z.B. Handbuch-Previews ohne Provider): No-op statt Crash.
    return {
      openTutorial: () => { /* kein Provider */ }, startTour: () => { /* kein Provider */ }, availableTours: ['user'],
      canCoach: false, coachActive: false, startCoach: () => { /* */ }, stopCoach: () => { /* */ }, reportWizard: () => { /* */ },
    };
  }
  return ctx;
}

export function TutorialProvider(props: { children: React.ReactNode }): React.ReactElement {
  const { canCreateEvents, isAdmin, isImpersonating } = useRoles();
  const { events, setTutorialDemoActive, deleteEvent, countExternalRegistrations, refreshEvents, getLastEventDeleteError } = useEvents();
  const { currentUser } = useCurrentUser();
  const { currentPage, selectedEventId, navigate } = useNavigation();
  const { locale } = useLanguage();
  const [chooserOpen, setChooserOpen] = React.useState(false);
  const [activeTour, setActiveTour] = React.useState<TutorialTour | null>(null);

  // ---------- v32.1.0: Mitmach-Tutorial (WizardCoach) ----------
  const canCoach = canCreateEvents || isAdmin;
  const [coachActive, setCoachActive] = React.useState(false);
  const [stationIdx, setStationIdxState] = React.useState(0);
  const [snapshot, setSnapshot] = React.useState<CoachSnapshot | null>(null);
  const [testEventId, setTestEventId] = React.useState('');
  const [base, setBase] = React.useState<StationBase>({ customFieldCount: 0 });
  const [deleteBusy, setDeleteBusy] = React.useState(false);
  const [deleteError, setDeleteError] = React.useState<string | null>(null);
  const [deleteDone, setDeleteDone] = React.useState(false);
  const lastSnapJsonRef = React.useRef('');
  const snapshotRef = React.useRef<CoachSnapshot | null>(null);
  snapshotRef.current = snapshot;

  const reportWizard = React.useCallback((snap: CoachSnapshot | null): void => {
    const j = JSON.stringify(snap);
    if (j === lastSnapJsonRef.current) return;
    lastSnapJsonRef.current = j;
    setSnapshot(snap);
    if (snap && snap.createdEventId) setTestEventId(prev => prev || snap.createdEventId);
  }, []);

  // Beim Betreten einer Station den Ausgangsstand merken (z.B. wie viele
  // Fragen es schon gab — „eine Frage hinzugefügt" heißt: eine MEHR).
  const setStationIdx = React.useCallback((i: number): void => {
    const snap = snapshotRef.current;
    setBase({ customFieldCount: snap ? snap.customFieldCount : 0 });
    setStationIdxState(Math.max(0, Math.min(i, COACH_STATIONS.length - 1)));
  }, []);

  const clearCoachDraft = (): void => { try { window.localStorage.removeItem(COACH_DRAFT_KEY); } catch { /* */ } };

  const startCoach = React.useCallback((): void => {
    setChooserOpen(false);
    setActiveTour(null);
    setTutorialDemoActive(false);
    clearCoachDraft();
    lastSnapJsonRef.current = '';
    setSnapshot(null);
    setTestEventId('');
    setDeleteBusy(false); setDeleteError(null); setDeleteDone(false);
    setStationIdxState(0);
    setBase({ customFieldCount: 0 });
    setCoachActive(true);
    // Der Assistent startet im Tutorial-Modus neu (key-Wechsel in
    // DexEventPlatform) — auch wenn er gerade schon offen ist.
    navigate('create-event');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [navigate]);

  const stopCoach = React.useCallback((): void => {
    setCoachActive(false);
    clearCoachDraft();
    lastSnapJsonRef.current = '';
    setSnapshot(null);
  }, []);

  // Die Id des angelegten Test-Events kommt sicher aus dem Erfolgs-Ereignis.
  React.useEffect(() => {
    if (!coachActive) return undefined;
    const onOk = (e: Event): void => {
      const d = (e as CustomEvent).detail as { eventId?: string; type?: string } | undefined;
      if (d && d.type === 'create' && d.eventId) setTestEventId(String(d.eventId));
    };
    window.addEventListener('dex-event-submit-success', onOk);
    return () => window.removeEventListener('dex-event-submit-success', onOk);
  }, [coachActive]);

  // Löschen erst, wenn der Outlook-Termin angelegt ist — sonst fehlt der
  // `CalendarLink`, niemand sagt den Termin ab, und er bleibt im Kalender
  // stehen. Solange er fehlt, alle 15 s nachladen (höchstens zwölfmal).
  const testEvent = testEventId ? (events || []).find(e => e.id === testEventId) : undefined;
  const onFinish = coachActive && COACH_STATIONS[stationIdx] && COACH_STATIONS[stationIdx].id === 'finish';
  const outlookAusstehend = !!testEvent && !testEvent.disableOutlook && !(testEvent.calendarLink || '').trim();
  const isDeCoach = locale === 'de';
  const blockedReason = !testEventId ? null
    : !testEvent ? (isDeCoach ? 'Dein Test-Event wird noch geladen …' : 'Your test event is still loading …')
      : outlookAusstehend ? (isDeCoach ? 'Der Outlook-Termin wird gerade noch angelegt — Löschen geht in einem Moment, damit er mit abgesagt wird.' : 'The Outlook invite is still being created — deleting works in a moment so it gets cancelled too.')
        : null;
  React.useEffect(() => {
    if (!onFinish || deleteDone || !blockedReason) return undefined;
    let n = 0;
    const iv = window.setInterval(() => {
      n += 1;
      if (n > 12) { window.clearInterval(iv); return; }
      refreshEvents().catch(() => { /* nächster Versuch */ });
    }, 15000);
    return () => window.clearInterval(iv);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [onFinish, deleteDone, !!blockedReason]);

  const deleteTest = async (): Promise<void> => {
    if (!testEvent || deleteBusy) return;
    setDeleteBusy(true); setDeleteError(null);
    try {
      // Dieselbe Schranke wie beim Entwurf-Löschen auf der Startseite: Hat
      // sich jemand außer dem Organizer-Team angemeldet, wird hier nicht gelöscht.
      const ext = await countExternalRegistrations(testEvent);
      if (ext > 0) {
        setDeleteError(isDeCoach ? 'Am Test-Event sind Anmeldungen von anderen Personen — es wird hier nicht gelöscht.' : 'Other people registered for the test event — it is not deleted here.');
        return;
      }
      const gone = await deleteEvent(testEvent.id);
      if (!gone) {
        setDeleteError(getLastEventDeleteError(isDeCoach ? 'de' : 'en') || (isDeCoach ? 'Löschen fehlgeschlagen — bitte erneut versuchen.' : 'Deletion failed — please try again.'));
        return;
      }
      setDeleteDone(true);
      try { await refreshEvents(); } catch { /* */ }
      navigate('start');
    } catch {
      setDeleteError(isDeCoach ? 'Löschen fehlgeschlagen — bitte erneut versuchen.' : 'Deletion failed — please try again.');
    } finally { setDeleteBusy(false); }
  };

  // Organizer-Tour sichtbar für alle mit Organizer-Einstieg — gleiche Logik
  // wie die Organizer-Kachel der Startseite (inkl. Co-Organizer + Demo-Modus).
  const currentEmailLc = (currentUser?.email || '').toLowerCase();
  const isOrganizerOfAnyEvent = !!currentEmailLc && (events || []).some(e => {
    if ((e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc)) return true;
    return (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
  });
  const hasOrganizerTour = canCreateEvents || isAdmin || isOrganizerOfAnyEvent || isImpersonating;

  const value = React.useMemo<TutorialContextType>(() => {
    const availableTours: TutorialTourId[] = hasOrganizerTour ? ['user', 'organizer'] : ['user'];
    const startTour = (id: TutorialTourId): void => {
      setChooserOpen(false);
      // v22.23: Organizer-Tour blendet ein Demo-Event in die Organizer-Liste
      // ein (nur client-seitig, verschwindet beim Beenden der Tour wieder).
      setTutorialDemoActive(id === 'organizer');
      setActiveTour(TUTORIAL_TOURS[id]);
    };
    const openTutorial = (): void => {
      if (availableTours.length === 1 && !canCoach) {
        startTour('user');
      } else {
        setChooserOpen(true);
      }
    };
    return { openTutorial, startTour, availableTours, canCoach, coachActive, startCoach, stopCoach, reportWizard };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hasOrganizerTour, canCoach, coachActive, startCoach, stopCoach, reportWizard]);

  // v32.0.14: Einstieg von außen (Deep-Link ?action=tutorial aus der
  // Onboarding-Mail). Ein Fenster-Ereignis statt eines Props, damit
  // DexEventPlatform die Tour starten kann, ohne im Provider zu stecken.
  // detail 'organizer' startet die Organizer-Tour direkt, wenn verfügbar.
  React.useEffect(() => {
    const onOpen = (e: Event): void => {
      const want = (e as CustomEvent).detail;
      // v32.1.0: Wer Events anlegen darf, landet direkt im Mitmach-Tutorial.
      if (want === 'organizer' && value.canCoach) value.startCoach();
      else if (want === 'organizer' && value.availableTours.indexOf('organizer') >= 0) value.startTour('organizer');
      else value.openTutorial();
    };
    window.addEventListener('dex-open-tutorial', onOpen);
    return () => window.removeEventListener('dex-open-tutorial', onOpen);
  }, [value]);

  const closeTour = React.useCallback((): void => {
    setActiveTour(null);
    setTutorialDemoActive(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <TutorialContext.Provider value={value}>
      {props.children}
      {chooserOpen && (
        <TourChooser
          onPick={value.startTour}
          onPickCoach={canCoach ? startCoach : undefined}
          tours={availableToursList(hasOrganizerTour)}
          onClose={() => setChooserOpen(false)}
        />
      )}
      {activeTour && (
        <TutorialOverlay tour={activeTour} onClose={closeTour} />
      )}
      {coachActive && (
        <WizardCoach
          snapshot={snapshot}
          env={{ page: currentPage, selectedEventId: selectedEventId || null, testEventId, myEmail: currentUser?.email || '', isDe: isDeCoach }}
          stationIdx={stationIdx}
          setStationIdx={setStationIdx}
          base={base}
          onClose={stopCoach}
          onGoWizardStep={(n: number) => { try { window.dispatchEvent(new CustomEvent('dex-tutorial-wizard-step', { detail: n })); } catch { /* */ } }}
          onGoToWizard={() => navigate('create-event')}
          onDeleteTest={() => { void deleteTest(); }}
          onKeepTest={stopCoach}
          onRealEvent={() => { stopCoach(); window.setTimeout(() => navigate('create-event'), 0); }}
          deleteState={{ busy: deleteBusy, blockedReason, error: deleteError, done: deleteDone }}
        />
      )}
    </TutorialContext.Provider>
  );
}

function availableToursList(hasOrganizerTour: boolean): TutorialTourId[] {
  return hasOrganizerTour ? ['user', 'organizer'] : ['user'];
}

/** Auswahl-Dialog, wenn mehrere Touren verfügbar sind (Organizer/Admin).
 *  v32.1.0: Wer Events anlegen darf, sieht zuerst das Mitmach-Tutorial
 *  (empfohlen) — die beiden Rundgänge darunter. */
function TourChooser(props: { onPick: (id: TutorialTourId) => void; onPickCoach?: () => void; tours: TutorialTourId[]; onClose: () => void }): React.ReactElement {
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const tours = props.tours;
  const kachel = (key: string, titel: string, text: string, onClick: () => void, empfohlen?: boolean): React.ReactElement => (
    <button key={key} type="button" onClick={onClick} className={`dex-ui-choice${empfohlen ? ' is-active' : ''}`}
      style={{ textAlign: 'left', cursor: 'pointer', fontFamily: 'inherit', width: '100%', display: 'block' }}>
      <span style={{ display: 'flex', alignItems: 'center', gap: 8, fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-gray-800)', marginBottom: 4 }}>
        {titel}
        {empfohlen && <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'Empfohlen' : 'Recommended'}</span>}
      </span>
      <span style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', lineHeight: 1.5 }}>{text}</span>
    </button>
  );
  return (
    <Modal open={true} onClose={props.onClose} maxWidth={560} ariaLabel={isDe ? 'Tutorial auswählen' : 'Choose tutorial'}
      title={isDe ? 'Welches Tutorial möchtest du starten?' : 'Which tutorial would you like to start?'}
      subtitle={isDe ? 'Du kannst jedes jederzeit neu starten.' : 'You can restart each of them anytime.'}>
      <div style={{ display: 'grid', gap: 10 }}>
        {props.onPickCoach && kachel('coach',
          isDe ? 'Test-Event gemeinsam anlegen' : 'Create a test event together',
          isDe ? 'Mitmachen statt zuschauen: Du tippst und klickst selbst, ich zeige dir wo. Am Ende steht ein echtes Test-Event, das nur du siehst — und das du mit einem Klick wieder löschst. Etwa fünf Minuten.'
            : 'Hands-on instead of watching: you type and click, I show you where. At the end there is a real test event only you can see — deleted again with one click. About five minutes.',
          props.onPickCoach, true)}
        {tours.map(id => {
          const tour = TUTORIAL_TOURS[id];
          return kachel(id, isDe ? tour.labelDe : tour.labelEn, isDe ? tour.descDe : tour.descEn, () => props.onPick(id));
        })}
      </div>
    </Modal>
  );
}

/** Das eigentliche Tour-Overlay: Spotlight + Schritt-Karte + Navigation. */
function TutorialOverlay(props: { tour: TutorialTour; onClose: () => void }): React.ReactElement {
  const { tour, onClose } = props;
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const { currentPage, navigate } = useNavigation();
  const [stepIdx, setStepIdx] = React.useState(0);
  const step: TutorialStep = tour.steps[stepIdx];
  // rect = gemessene Position des Ziel-Elements; null + settled = zentriert.
  const [rect, setRect] = React.useState<{ top: number; left: number; width: number; height: number } | null>(null);
  const [settled, setSettled] = React.useState(false);
  const targetRef = React.useRef<HTMLElement | null>(null);

  const measure = React.useCallback((): void => {
    const el = targetRef.current;
    if (!el || !el.isConnected) return;
    const r = el.getBoundingClientRect();
    setRect({ top: r.top, left: r.left, width: r.width, height: r.height });
  }, []);

  // Pro Schritt: Seite ansteuern, dann Ziel-Element pollen (Seiten laden lazy).
  React.useEffect(() => {
    setSettled(false);
    setRect(null);
    targetRef.current = null;
    if (currentPage !== step.page) {
      // v31.59: ersetzend — die Tour soll keine Kette von Rücksprungzielen
      // hinterlassen, sonst führt „Zurück" nach der Tour durch alle Tour-Seiten.
      navigate(step.page, undefined, undefined, { replace: true });
      // Fallback: Wenn die Navigation blockiert wird (z.B. Unsaved-Changes-
      // Guard einer Seite), nach kurzer Zeit trotzdem die Karte zentriert
      // zeigen — sonst bliebe nur der dunkle Backdrop ohne Ausweg.
      const t = window.setTimeout(() => setSettled(true), 1500);
      return () => window.clearTimeout(t); // Effekt läuft nach dem Seitenwechsel erneut (currentPage-Dep).
    }
    // v22.23: Wizard-Klick-Through — auf der create-event-Seite stellt die
    // Tour den Wizard-Schritt per CustomEvent um, bevor der Spotlight sucht.
    if (typeof step.wizardStep === 'number') {
      try {
        window.dispatchEvent(new CustomEvent('dex-tutorial-wizard-step', { detail: step.wizardStep }));
      } catch { /* ältere Browser ohne CustomEvent-Konstruktor */ }
    }
    if (!step.selector) {
      setSettled(true);
      return;
    }
    let cancelled = false;
    let tries = 0;
    // v22.23: Karte nicht erst nach dem Polling-Ende zeigen — nach 0,7 s
    // erscheint sie (zunächst zentriert), der Spotlight rückt nach, sobald
    // das Ziel-Element gefunden ist. Vorher wartete der User bei fehlendem
    // Element bis zu 4 s auf einen dunklen Bildschirm.
    const earlyTimer = window.setTimeout(() => { if (!cancelled) setSettled(true); }, 700);
    const poll = (): void => {
      if (cancelled) return;
      const el = document.querySelector(step.selector as string) as HTMLElement | null;
      if (el) {
        targetRef.current = el;
        try { el.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch { /* ältere Browser */ }
        // Nach dem Scroll messen (smooth braucht einen Moment).
        window.setTimeout(() => { if (!cancelled) { measure(); setSettled(true); } }, 350);
        return;
      }
      tries += 1;
      if (tries >= 25) { setSettled(true); return; } // ~4s — dann zentriert.
      window.setTimeout(poll, 160);
    };
    poll();
    return () => { cancelled = true; window.clearTimeout(earlyTimer); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [stepIdx, currentPage, tour.id]);

  // Bei Resize/Scroll nachmessen, damit der Spotlight am Element klebt.
  React.useEffect(() => {
    const onMove = (): void => measure();
    window.addEventListener('resize', onMove);
    window.addEventListener('scroll', onMove, true);
    return () => {
      window.removeEventListener('resize', onMove);
      window.removeEventListener('scroll', onMove, true);
    };
  }, [measure]);

  // ESC beendet die Tour.
  React.useEffect(() => {
    const onKey = (e: KeyboardEvent): void => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);

  const isFirst = stepIdx === 0;
  const isLast = stepIdx === tour.steps.length - 1;
  const goNext = (): void => { if (isLast) { onClose(); } else { setStepIdx(i => i + 1); } };
  const goBackStep = (): void => { if (!isFirst) setStepIdx(i => i - 1); };

  // Karten-Position: unter dem Spotlight wenn Platz, sonst darüber; seitlich
  // an den Viewport geklemmt. Ohne Spotlight: mittig.
  // v22.25: Ziel-Rechteck auf den Viewport clampen — bei Zielen GRÖSSER als
  // der Viewport (z.B. das ganze Event-Raster) zeigten „unter dem Ziel" UND
  // „über dem Ziel" beide aus dem Bildschirm heraus → Karte unsichtbar, Tour
  // wirkte eingefroren. Reicht der Platz weder unten noch oben, wird die
  // Karte unten mittig ÜBER dem Spotlight fixiert.
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  const pad = 8; // Luft um das Ziel-Element im Spotlight
  let cardStyle: React.CSSProperties;
  if (rect) {
    const cardW = Math.min(CARD_WIDTH, vw - 32);
    const rectTopClamped = Math.max(rect.top, 0);
    const rectBottomClamped = Math.min(rect.top + rect.height, vh);
    const spaceBelow = vh - rectBottomClamped;
    const spaceAbove = rectTopClamped;
    const left = Math.min(Math.max(16, rect.left + rect.width / 2 - cardW / 2), Math.max(16, vw - cardW - 16));
    // Großzügig geschätzter Platzbedarf der Karte inkl. Abstand.
    const CARD_SPACE = 280;
    if (spaceBelow >= CARD_SPACE) {
      cardStyle = { position: 'fixed', top: rectBottomClamped + pad + 16, left, width: cardW };
    } else if (spaceAbove >= CARD_SPACE) {
      cardStyle = { position: 'fixed', bottom: vh - rectTopClamped + pad + 16, left, width: cardW };
    } else {
      cardStyle = { position: 'fixed', bottom: 24, left: '50%', transform: 'translateX(-50%)', width: cardW };
    }
  } else {
    cardStyle = {
      position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)',
      width: Math.min(440, vw - 32),
    };
  }

  return (
    <div aria-live="polite" style={{ position: 'fixed', inset: 0, zIndex: TOUR_Z_INDEX }}>
      {/* Klick-Fänger: blockiert die App-Interaktion während der Tour. */}
      <div style={{ position: 'absolute', inset: 0 }} onClick={() => { /* Klicks schlucken */ }} />
      {rect && settled ? (
        // Spotlight: das „Loch" entsteht durch den riesigen Box-Shadow.
        <div
          style={{
            position: 'fixed',
            top: rect.top - pad, left: rect.left - pad,
            width: rect.width + pad * 2, height: rect.height + pad * 2,
            borderRadius: 12,
            boxShadow: '0 0 0 100000px rgba(15,23,42,0.62)',
            border: '2px solid var(--dex-green, #86bc25)',
            pointerEvents: 'none',
            transition: 'top 0.25s ease, left 0.25s ease, width 0.25s ease, height 0.25s ease',
          }}
        />
      ) : (
        <div style={{ position: 'absolute', inset: 0, background: 'rgba(15,23,42,0.62)', pointerEvents: 'none' }} />
      )}

      {/* Schritt-Karte — erscheint erst, wenn der Schritt „angekommen" ist. */}
      {settled && (
        <div
          role="dialog"
          aria-label={isDe ? step.titleDe : step.titleEn}
          style={{
            ...cardStyle,
            background: '#fff', borderRadius: 16, padding: '18px 20px 16px',
            boxShadow: '0 16px 48px rgba(0,0,0,0.35)',
            fontFamily: 'inherit',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
            <span style={{
              fontSize: '0.7rem', fontWeight: 700, letterSpacing: '0.06em', textTransform: 'uppercase',
              color: 'var(--dex-green-dark, #4a7c1f)',
            }}>
              {isDe ? tour.labelDe : tour.labelEn} · {isDe ? 'Schritt' : 'Step'} {stepIdx + 1}/{tour.steps.length}
            </span>
            <button
              type="button"
              onClick={onClose}
              aria-label={isDe ? 'Tour beenden' : 'End tour'}
              style={{
                background: 'none', border: 'none', cursor: 'pointer',
                color: 'var(--dex-gray-400)', fontSize: '1.1rem', lineHeight: 1, padding: 4,
              }}
            >×</button>
          </div>
          {/* Fortschrittsbalken */}
          <div style={{ background: 'var(--dex-gray-100, #f0f0f0)', borderRadius: 999, height: 5, overflow: 'hidden', marginBottom: 12 }}>
            <div style={{
              width: `${Math.round(((stepIdx + 1) / tour.steps.length) * 100)}%`,
              height: '100%', background: 'var(--dex-green, #86bc25)', borderRadius: 999,
              transition: 'width 0.25s ease',
            }} />
          </div>
          <h3 style={{ margin: '0 0 8px', fontSize: '1.05rem', color: 'var(--dex-gray-800)' }}>
            {isDe ? step.titleDe : step.titleEn}
          </h3>
          <p style={{ margin: '0 0 16px', fontSize: '0.86rem', color: 'var(--dex-gray-600)', lineHeight: 1.6 }}>
            {isDe ? step.bodyDe : step.bodyEn}
          </p>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
            <button
              type="button"
              onClick={onClose}
              style={{
                background: 'none', border: 'none', cursor: 'pointer', padding: '6px 0',
                color: 'var(--dex-gray-400)', fontSize: '0.78rem', fontFamily: 'inherit',
              }}
            >
              {isDe ? 'Tour beenden' : 'End tour'}
            </button>
            <div style={{ display: 'flex', gap: 8 }}>
              {!isFirst && (
                <button className="btn btn-secondary" type="button" onClick={goBackStep} style={{ fontSize: '0.84rem', padding: '8px 16px' }}>
                  {isDe ? 'Zurück' : 'Back'}
                </button>
              )}
              <button className="btn btn-primary" type="button" onClick={goNext} style={{ fontSize: '0.84rem', padding: '8px 18px' }}>
                {isLast
                  ? (isDe ? 'Fertig' : 'Done')
                  : isFirst
                    ? (isDe ? 'Los geht’s' : 'Let’s go')
                    : (isDe ? 'Weiter' : 'Next')}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
