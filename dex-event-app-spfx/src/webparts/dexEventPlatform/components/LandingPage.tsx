// Landing Page - Startbildschirm mit animiertem Orb und Willkommensnachricht

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
// v18.25: personalisierte Begrüßung (Vorname) im Landing-Hero — useCurrentUser
// wird dafür wieder hier benötigt. (v18.26: KPI-Zeile wieder entfernt — die
// Einsatz-Zahlen stehen nur auf dem Boot-Loader davor.)
import { useCurrentUser } from '../context/UserContext';
import { APP_VERSION } from '../version';
import { Info, Mail } from './Icons';
import LandingInfoModal from './LandingInfoModal';
import InquiryModal from './InquiryModal';

export default function LandingPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { locale, setLocale, t } = useLanguage();
  const isDe = locale === 'de';
  // v18.25: Vorname für die persönliche Begrüßung.
  const { currentUser } = useCurrentUser();
  const firstName = (currentUser?.firstName || '').trim();
  const [showInfo, setShowInfo] = React.useState(false);
  // v13.3: Inquiry-Modal lebt jetzt komplett in der wiederverwendbaren
  // InquiryModal-Komponente — eigene States hier entfallen.
  const [showInquiry, setShowInquiry] = React.useState(false);

  // Keyframes als inline style-Tag injizieren, da SPFx SCSS-Module
  // @keyframes innerhalb von :global manchmal nicht korrekt emittieren
  React.useEffect(() => {
    const id = 'dex-orb-keyframes';
    if (!document.getElementById(id)) {
      const style = document.createElement('style');
      style.id = id;
      style.textContent = `
        @keyframes dexOrbSpin { to { transform: rotate(360deg); } }
      `;
      document.head.appendChild(style);
    }
  }, []);

  // v6.26: Die Mobile-Check-In-Bubble lebt jetzt im Header neben dem vorhandenen
  // QR-Icon — nicht mehr hier auf der LandingPage. Damit kommt sie direkt mit
  // dem App-Start und bleibt auf jeder Seite konsistent an einem festen Ort.

  return (
    <div className="landing" style={{ position: 'relative' }}>
      <span style={{
        position: 'absolute', top: 12, right: 16,
        fontSize: '0.7rem', color: 'var(--dex-gray-300)',
      }}>
        v{APP_VERSION}
      </span>

      <div className="landing__hero">
        <div className="landing__card" style={{ position: 'relative' }}>
          {/* Sprachauswahl - oben links in der weissen Card */}
          <div style={{
            position: 'absolute', top: 16, left: 16,
            display: 'flex', gap: 4,
          }}>
            <button
              onClick={() => setLocale('de')}
              style={{
                background: locale === 'de' ? 'var(--dex-green)' : 'var(--dex-gray-100)',
                border: 'none',
                borderRadius: 6, padding: '6px 12px', cursor: 'pointer', fontSize: '0.8rem',
                color: locale === 'de' ? '#fff' : 'var(--dex-gray-500)',
                fontWeight: locale === 'de' ? 700 : 500,
                transition: 'all 0.2s',
              }}
              title="Deutsch"
            >
              DE
            </button>
            <button
              onClick={() => setLocale('en')}
              style={{
                background: locale === 'en' ? 'var(--dex-green)' : 'var(--dex-gray-100)',
                border: 'none',
                borderRadius: 6, padding: '6px 12px', cursor: 'pointer', fontSize: '0.8rem',
                color: locale === 'en' ? '#fff' : 'var(--dex-gray-500)',
                fontWeight: locale === 'en' ? 700 : 500,
                transition: 'all 0.2s',
              }}
              title="English"
            >
              EN
            </button>
          </div>
          <div className="landing__orb">
            <div className="landing__orb-inner" />
          </div>
          <div className="landing__text">
            <h1>
              {isDe ? 'Hallo' : 'Hi'}{firstName ? <> <strong>{firstName}</strong></> : ''}{', '}
              <span style={{ whiteSpace: 'nowrap' }}>
                {isDe ? 'willkommen bei ' : 'welcome to '}<strong>DEX</strong>.
              </span>
            </h1>
            <p>
              {isDe
                ? 'Deine zentrale Plattform für Deloitte Events – von der Einladung über die Anmeldung bis zum Check-in.'
                : 'Your central platform for Deloitte events – from invitation through registration to check-in.'}
            </p>
          </div>
          <button className="btn btn-lg btn-block btn-outline" onClick={() => navigate('start')}>
            {t('landing.start')}
          </button>
          <div
            className="landing__actions"
            style={{
              display: 'flex', gap: 12, justifyContent: 'center', alignItems: 'center',
              flexWrap: 'wrap',
            }}
          >
            <button
              onClick={() => setShowInfo(!showInfo)}
              style={{
                background: 'none', border: '2px solid var(--dex-gray-300)', borderRadius: '50%',
                width: 40, height: 40, display: 'flex', alignItems: 'center', justifyContent: 'center',
                cursor: 'pointer', color: 'var(--dex-gray-400)', transition: 'all 0.2s',
                flexShrink: 0,
              }}
              title={locale === 'de' ? 'Über die App' : 'About the app'}
            >
              <Info size={18} />
            </button>
            <button
              type="button"
              onClick={() => setShowInquiry(true)}
              style={{
                background: 'none', border: '2px solid var(--dex-gray-300)', borderRadius: '50%',
                width: 40, height: 40, display: 'flex', alignItems: 'center', justifyContent: 'center',
                cursor: 'pointer', color: 'var(--dex-gray-400)', transition: 'all 0.2s',
                flexShrink: 0, padding: 0,
              }}
              title={locale === 'de' ? 'DEX App für dein Event anfragen' : 'Request the DEX App for your event'}
            >
              <Mail size={18} />
            </button>
            <button
              type="button"
              onClick={() => setShowInquiry(true)}
              className="landing__bubble"
              style={{
                position: 'relative',
                background: 'var(--dex-green)',
                color: '#fff',
                padding: '10px 14px',
                borderRadius: 12,
                fontSize: '0.82rem',
                lineHeight: 1.35,
                maxWidth: 280,
                boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
                border: 'none',
                cursor: 'pointer',
                textAlign: 'left',
                fontFamily: 'inherit',
              }}
              title={locale === 'de' ? 'Anfrage senden' : 'Send inquiry'}
            >
              {locale === 'de'
                ? 'Möchtest du die DEX App auch für dein Event nutzen? Melde dich gerne bei uns!'
                : 'Want to use the DEX App for your event too? Just reach out to us!'}
            </button>
          </div>

          <div
            style={{
              fontSize: '0.95rem',
              color: 'var(--dex-gray-400)',
              textAlign: 'center',
              marginTop: 6,
              lineHeight: 1.3,
            }}
          >
            {locale === 'de' ? 'Entwickelt von ' : 'Built by '}
            <span style={{ fontWeight: 600, color: 'var(--dex-gray-500)' }}>
              <DevName name="Eike Brenneisen" email="ebrenneisen@deloitte.de" />
              {', '}
              <DevName name="Andreas Enk" email="aenk@deloitte.de" />
              {' '}{locale === 'de' ? 'und' : 'and'}{' '}
              <DevName name="Nils Felten" email="nifelten@deloitte.de" />
            </span>
          </div>
        </div>
      </div>
      <LandingInfoModal open={showInfo} locale={locale === 'de' ? 'de' : 'en'} onClose={() => setShowInfo(false)} />
      {/* v13.3: Inquiry-Modal aus der wiederverwendbaren Komponente. */}
      <InquiryModal open={showInquiry} onClose={() => setShowInquiry(false)} />
    </div>
  );
}

// v11.47: KPI-Box-Reihe ueber dem "Entwickelt von ..."-Block. Drei Boxen
// nebeneinander: gehostete Events, Teilnehmer, App-Aufrufe. Jede Box mit
// einer AnimatedCounter-Komponente, die beim ersten Verfuegbarwerden des
// Werts von 0 zum Zielwert ease-out hochzaehlt (~1.6s). Solange Daten noch
// laden, steht ein dezenter Skeleton-Punkt im Wert-Feld.
export function KpiRow(props: {
  locale: string;
  eventsLoading: boolean;
  participantsLoading: boolean;
  events: number;
  participants: number;
}): React.ReactElement {
  const isDe = props.locale === 'de';
  const labels = isDe
    ? { ev: 'Events', pa: 'Teilnehmer' }
    : { ev: 'Events', pa: 'Attendees' };
  return (
    <div style={{
      display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 10,
    }}>
      <KpiBox label={labels.ev} value={props.events} loading={props.eventsLoading} />
      <KpiBox label={labels.pa} value={props.participants} loading={props.participantsLoading} />
    </div>
  );
}

function KpiBox(props: { label: string; value: number; loading: boolean }): React.ReactElement {
  return (
    <div style={{
      display: 'flex', flexDirection: 'column', alignItems: 'center',
      padding: '12px 8px',
      background: 'linear-gradient(135deg, rgba(134,188,37,0.08), rgba(0,118,168,0.04))',
      border: '1px solid var(--dex-gray-200)',
      borderRadius: 12,
      minWidth: 0,
    }}>
      <div style={{
        fontSize: 'clamp(1.4rem, 4vw, 1.9rem)', fontWeight: 800,
        color: 'var(--dex-green-dark, #4a7c1f)',
        lineHeight: 1.1, letterSpacing: '-0.02em',
        fontVariantNumeric: 'tabular-nums',
      }}>
        {props.loading
          ? <SkeletonDots />
          : <AnimatedCounter value={props.value} />}
      </div>
      <div style={{
        marginTop: 6,
        fontSize: '0.72rem', color: 'var(--dex-gray-500)',
        fontWeight: 500, textAlign: 'center', lineHeight: 1.2,
      }}>
        {props.label}
      </div>
    </div>
  );
}

function SkeletonDots(): React.ReactElement {
  return (
    <span aria-hidden="true" style={{ display: 'inline-flex', gap: 4, opacity: 0.55 }}>
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', animation: 'dexKpiPulse 1.2s ease-in-out infinite' }} />
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', animation: 'dexKpiPulse 1.2s ease-in-out 0.2s infinite' }} />
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', animation: 'dexKpiPulse 1.2s ease-in-out 0.4s infinite' }} />
      <style>{`
        @keyframes dexKpiPulse {
          0%, 100% { transform: scale(0.6); opacity: 0.3; }
          50%      { transform: scale(1);   opacity: 1; }
        }
      `}</style>
    </span>
  );
}

/** Ease-out-Cubic-Animation von 0 (bzw. dem letzten geanimierten Wert) auf
 *  `value`. Wenn `value` sich aendert, startet eine neue Animation; bei
 *  schnellen Aenderungen wird die laufende Animation sauber durch eine neue
 *  ersetzt (requestAnimationFrame + AbortFlag). */
function AnimatedCounter(props: { value: number; durationMs?: number }): React.ReactElement {
  const target = Math.max(0, Math.floor(props.value || 0));
  // v11.79: Default-Dauer von 1600 ms → 600 ms reduziert. Seit der App-Boot
  // unter ~1.6 s liegt, wirkte das laengere Count-Up-Tempo traege; die Zahl
  // tickt jetzt knackiger hoch ohne hektisch zu wirken.
  const duration = props.durationMs ?? 600;
  const [shown, setShown] = React.useState<number>(0);
  const startRef = React.useRef<number>(0);
  const fromRef = React.useRef<number>(0);
  const rafRef = React.useRef<number | null>(null);
  React.useEffect(() => {
    if (rafRef.current !== null) cancelAnimationFrame(rafRef.current);
    fromRef.current = shown;
    startRef.current = 0;
    const step = (ts: number): void => {
      if (startRef.current === 0) startRef.current = ts;
      const elapsed = ts - startRef.current;
      const t = Math.min(1, elapsed / duration);
      // ease-out-cubic
      const eased = 1 - Math.pow(1 - t, 3);
      const value = fromRef.current + (target - fromRef.current) * eased;
      setShown(Math.round(value));
      if (t < 1) {
        rafRef.current = requestAnimationFrame(step);
      } else {
        rafRef.current = null;
      }
    };
    rafRef.current = requestAnimationFrame(step);
    return () => {
      if (rafRef.current !== null) cancelAnimationFrame(rafRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [target, duration]);
  return <span>{shown.toLocaleString('de-DE')}</span>;
}

// v7.1: Entwickler-Name mit Hover-Popover auf der LandingPage.
// Analog zu OrganizerList: bei Hover erscheint ein kleines Popover mit
// größerem Foto, Name und Rolle/Standort. Die Daten (jobTitle + location)
// werden per SharePoint-User-Profil-Lookup live nachgeladen (einmalig
// beim ersten Hover) — damit zeigen wir immer den aktuellen Rollen-Stand
// aus dem AD, keine hardcoded Strings.
function DevName(props: { name: string; email: string }): React.ReactElement {
  const [hovered, setHovered] = React.useState(false);
  const [failed, setFailed] = React.useState(false);
  const [profile, setProfile] = React.useState<{ jobTitle: string; location: string } | null>(null);
  const loadedRef = React.useRef(false);
  React.useEffect(() => {
    if (!hovered || loadedRef.current) return;
    loadedRef.current = true;
    (async () => {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ctx = (window as any).__dexSpfxContext;
        if (!ctx) return;
        const { SharePointService } = await import('../services/SharePointService');
        const svc = new SharePointService(ctx);
        const result = await svc.searchUserByEmail(props.email);
        if (result) {
          setProfile({
            jobTitle: result.jobTitle || '',
            location: result.location || '',
          });
        }
      } catch { /* Profil-Lookup fehlgeschlagen — Fallback bleibt leer */ }
    })().catch(() => { /* ignore */ });
  }, [hovered, props.email]);

  const roleLine = profile
    ? [profile.jobTitle, profile.location].filter(Boolean).join(' · ')
    : '';

  return (
    <span
      style={{ position: 'relative', cursor: 'default', display: 'inline-block' }}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
    >
      <span style={{ textDecoration: hovered ? 'underline' : 'none', textDecorationColor: 'var(--dex-green)' }}>
        {props.name}
      </span>
      {hovered && (
        <span style={{
          position: 'absolute', bottom: 'calc(100% + 10px)', left: '50%',
          transform: 'translateX(-50%)',
          background: '#fff', borderRadius: 10, padding: '10px 12px',
          boxShadow: '0 8px 24px rgba(0,0,0,0.15)',
          display: 'flex', alignItems: 'center', gap: 10,
          whiteSpace: 'nowrap', zIndex: 20, pointerEvents: 'none',
          border: '1px solid var(--dex-gray-200)',
        }}>
          {failed ? (
            <span style={{
              width: 48, height: 48, borderRadius: '50%',
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              background: 'linear-gradient(135deg, #86bc25, #0076a8)',
              color: '#fff', fontWeight: 700, fontSize: '0.9rem', flexShrink: 0,
            }}>
              {props.name.split(' ').map(p => p[0]).slice(0, 2).join('')}
            </span>
          ) : (
            <img
              src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(props.email)}&size=L`}
              alt={props.name}
              onError={() => setFailed(true)}
              style={{ width: 48, height: 48, borderRadius: '50%', objectFit: 'cover', flexShrink: 0, background: 'var(--dex-gray-200)' }}
            />
          )}
          <span style={{ textAlign: 'left', fontSize: '0.82rem', lineHeight: 1.35 }}>
            <span style={{ display: 'block', fontWeight: 700, color: 'var(--dex-gray-800)' }}>{props.name}</span>
            {roleLine && (
              <span style={{ display: 'block', color: 'var(--dex-gray-500)', fontSize: '0.75rem' }}>{roleLine}</span>
            )}
            <span style={{ display: 'block', color: 'var(--dex-gray-400)', fontSize: '0.72rem' }}>{props.email}</span>
          </span>
        </span>
      )}
    </span>
  );
}
