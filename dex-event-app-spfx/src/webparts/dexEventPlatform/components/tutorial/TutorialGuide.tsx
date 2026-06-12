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

const TOUR_Z_INDEX = 10800;
const CARD_WIDTH = 380;

interface TutorialContextType {
  /** Öffnet das Tutorial — bei mehreren verfügbaren Touren zuerst die Auswahl. */
  openTutorial: () => void;
  startTour: (id: TutorialTourId) => void;
  availableTours: TutorialTourId[];
}

const TutorialContext = React.createContext<TutorialContextType | undefined>(undefined);

export function useTutorial(): TutorialContextType {
  const ctx = React.useContext(TutorialContext);
  if (!ctx) {
    // Fallback (z.B. Handbuch-Previews ohne Provider): No-op statt Crash.
    return { openTutorial: () => { /* kein Provider */ }, startTour: () => { /* kein Provider */ }, availableTours: ['user'] };
  }
  return ctx;
}

export function TutorialProvider(props: { children: React.ReactNode }): React.ReactElement {
  const { canCreateEvents, isAdmin, isImpersonating } = useRoles();
  const { events } = useEvents();
  const { currentUser } = useCurrentUser();
  const [chooserOpen, setChooserOpen] = React.useState(false);
  const [activeTour, setActiveTour] = React.useState<TutorialTour | null>(null);

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
      setActiveTour(TUTORIAL_TOURS[id]);
    };
    const openTutorial = (): void => {
      if (availableTours.length === 1) {
        startTour('user');
      } else {
        setChooserOpen(true);
      }
    };
    return { openTutorial, startTour, availableTours };
  }, [hasOrganizerTour]);

  return (
    <TutorialContext.Provider value={value}>
      {props.children}
      {chooserOpen && (
        <TourChooser
          onPick={value.startTour}
          onClose={() => setChooserOpen(false)}
        />
      )}
      {activeTour && (
        <TutorialOverlay tour={activeTour} onClose={() => setActiveTour(null)} />
      )}
    </TutorialContext.Provider>
  );
}

/** Auswahl-Dialog, wenn mehrere Touren verfügbar sind (Organizer/Admin). */
function TourChooser(props: { onPick: (id: TutorialTourId) => void; onClose: () => void }): React.ReactElement {
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const tours: TutorialTourId[] = ['user', 'organizer'];
  return (
    <Modal open={true} onClose={props.onClose} maxWidth={560} ariaLabel={isDe ? 'Tutorial auswählen' : 'Choose tutorial'}>
      <h3 style={{ margin: '0 0 4px', fontSize: '1.1rem' }}>
        {isDe ? 'Welches Tutorial möchtest du starten?' : 'Which tutorial would you like to start?'}
      </h3>
      <p style={{ margin: '0 0 14px', fontSize: '0.84rem', color: 'var(--dex-gray-500)' }}>
        {isDe
          ? 'Beide dauern nur rund zwei Minuten — und du kannst sie jederzeit neu starten.'
          : 'Both take only about two minutes — and you can restart them anytime.'}
      </p>
      <div style={{ display: 'grid', gap: 10 }}>
        {tours.map(id => {
          const tour = TUTORIAL_TOURS[id];
          return (
            <button
              key={id}
              type="button"
              onClick={() => props.onPick(id)}
              style={{
                textAlign: 'left', cursor: 'pointer', fontFamily: 'inherit',
                background: '#fff', border: '1.5px solid var(--dex-gray-200)',
                borderRadius: 12, padding: '14px 16px',
                transition: 'border-color 0.15s, box-shadow 0.15s',
              }}
              onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--dex-green, #86bc25)'; }}
              onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--dex-gray-200)'; }}
            >
              <span style={{ display: 'block', fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-gray-800)', marginBottom: 4 }}>
                {isDe ? tour.labelDe : tour.labelEn}
              </span>
              <span style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', lineHeight: 1.5 }}>
                {isDe ? tour.descDe : tour.descEn}
              </span>
            </button>
          );
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
      navigate(step.page);
      // Fallback: Wenn die Navigation blockiert wird (z.B. Unsaved-Changes-
      // Guard einer Seite), nach kurzer Zeit trotzdem die Karte zentriert
      // zeigen — sonst bliebe nur der dunkle Backdrop ohne Ausweg.
      const t = window.setTimeout(() => setSettled(true), 1500);
      return () => window.clearTimeout(t); // Effekt läuft nach dem Seitenwechsel erneut (currentPage-Dep).
    }
    if (!step.selector) {
      setSettled(true);
      return;
    }
    let cancelled = false;
    let tries = 0;
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
    return () => { cancelled = true; };
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
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  const pad = 8; // Luft um das Ziel-Element im Spotlight
  let cardStyle: React.CSSProperties;
  if (rect) {
    const below = rect.top + rect.height + pad + 16;
    const placeBelow = below < vh * 0.62;
    const left = Math.min(Math.max(16, rect.left + rect.width / 2 - CARD_WIDTH / 2), Math.max(16, vw - CARD_WIDTH - 16));
    cardStyle = placeBelow
      ? { position: 'fixed', top: below, left, width: Math.min(CARD_WIDTH, vw - 32) }
      : { position: 'fixed', bottom: vh - rect.top + pad + 16, left, width: Math.min(CARD_WIDTH, vw - 32) };
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
