// Landing Page - Startbildschirm mit animiertem Orb und Willkommensnachricht

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
import { useCurrentUser } from '../context/UserContext';
import { useEvents } from '../context/EventContext';
import { APP_VERSION } from '../version';
import { Info, Mail } from './Icons';
import LandingInfoModal from './LandingInfoModal';

export default function LandingPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { locale, setLocale, t } = useLanguage();
  const { currentUser } = useCurrentUser();
  const { sendAdminInquiry } = useEvents();
  const [showInfo, setShowInfo] = React.useState(false);
  const [showInquiry, setShowInquiry] = React.useState(false);
  const userFullName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim();
  const [inquiryName, setInquiryName] = React.useState(userFullName);
  const [inquiryEvent, setInquiryEvent] = React.useState('');
  const [inquiryMessage, setInquiryMessage] = React.useState('');
  const [inquirySending, setInquirySending] = React.useState(false);
  const [inquiryStatus, setInquiryStatus] = React.useState<'' | 'success' | 'error'>('');
  React.useEffect(() => {
    if (showInquiry && !inquiryName && userFullName) setInquiryName(userFullName);
  }, [showInquiry, userFullName]);

  async function handleInquirySubmit(): Promise<void> {
    if (!inquiryEvent.trim() || !inquiryMessage.trim() || inquirySending) return;
    setInquirySending(true);
    setInquiryStatus('');
    const ok = await sendAdminInquiry(
      inquiryName.trim() || userFullName,
      currentUser.email || '',
      inquiryEvent.trim(),
      inquiryMessage.trim(),
    );
    setInquirySending(false);
    if (ok) {
      setInquiryStatus('success');
      setInquiryEvent('');
      setInquiryMessage('');
      setTimeout(() => { setShowInquiry(false); setInquiryStatus(''); }, 1800);
    } else {
      setInquiryStatus('error');
    }
  }

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
              {t('landing.welcome')}{' '}
              <strong style={{ whiteSpace: 'nowrap' }}>{t('landing.platform')}.</strong>
            </h1>
            <p>{t('landing.subtitle')}</p>
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
      {showInquiry && (
        <div
          role="dialog"
          aria-modal="true"
          onClick={() => { if (!inquirySending) setShowInquiry(false); }}
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            zIndex: 9999, padding: 16,
          }}
        >
          <div
            onClick={e => e.stopPropagation()}
            style={{
              background: '#fff', borderRadius: 16, padding: '24px 28px',
              maxWidth: 480, width: '100%', boxShadow: '0 12px 48px rgba(0,0,0,0.18)',
              display: 'flex', flexDirection: 'column', gap: 14,
            }}
          >
            <h2 style={{ margin: 0, fontSize: '1.3rem', color: 'var(--dex-gray-800)' }}>
              {locale === 'de' ? 'DEX App für dein Event anfragen' : 'Request the DEX App for your event'}
            </h2>
            <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
              {locale === 'de'
                ? 'Wir melden uns kurz bei dir und besprechen, wie wir dein Event unterstützen können.'
                : 'We will get back to you and discuss how we can support your event.'}
            </p>
            <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
              {locale === 'de' ? 'Dein Name' : 'Your name'}
              <input
                type="text"
                value={inquiryName}
                onChange={e => setInquiryName(e.target.value)}
                disabled={inquirySending}
                style={{
                  padding: '8px 10px', border: '1px solid var(--dex-gray-300)',
                  borderRadius: 8, fontSize: '0.9rem',
                }}
              />
            </label>
            <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
              {locale === 'de' ? 'Event-Name' : 'Event name'}
              <input
                type="text"
                value={inquiryEvent}
                onChange={e => setInquiryEvent(e.target.value)}
                disabled={inquirySending}
                placeholder={locale === 'de' ? 'z.B. Summer Party 2026' : 'e.g. Summer Party 2026'}
                style={{
                  padding: '8px 10px', border: '1px solid var(--dex-gray-300)',
                  borderRadius: 8, fontSize: '0.9rem',
                }}
              />
            </label>
            <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
              {locale === 'de' ? 'Was brauchst du?' : 'What do you need?'}
              <textarea
                value={inquiryMessage}
                onChange={e => setInquiryMessage(e.target.value)}
                disabled={inquirySending}
                rows={5}
                placeholder={locale === 'de'
                  ? 'Kurz beschreiben: Anzahl Teilnehmer, Termin, gewünschte Funktionen...'
                  : 'Briefly describe: number of participants, date, features needed...'}
                style={{
                  padding: '8px 10px', border: '1px solid var(--dex-gray-300)',
                  borderRadius: 8, fontSize: '0.9rem', fontFamily: 'inherit', resize: 'vertical',
                }}
              />
            </label>
            {inquiryStatus === 'success' && (
              <div style={{ color: 'var(--dex-green)', fontSize: '0.85rem' }}>
                {locale === 'de' ? 'Anfrage gesendet — wir melden uns!' : 'Request sent — we will get back to you!'}
              </div>
            )}
            {inquiryStatus === 'error' && (
              <div style={{ color: 'var(--dex-red)', fontSize: '0.85rem' }}>
                {locale === 'de' ? 'Senden fehlgeschlagen. Bitte später erneut versuchen.' : 'Sending failed. Please try again later.'}
              </div>
            )}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 4 }}>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => setShowInquiry(false)}
                disabled={inquirySending}
              >
                {locale === 'de' ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={handleInquirySubmit}
                disabled={inquirySending || !inquiryEvent.trim() || !inquiryMessage.trim()}
              >
                {inquirySending
                  ? (locale === 'de' ? 'Wird gesendet...' : 'Sending...')
                  : (locale === 'de' ? 'Anfrage senden' : 'Send inquiry')}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
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
