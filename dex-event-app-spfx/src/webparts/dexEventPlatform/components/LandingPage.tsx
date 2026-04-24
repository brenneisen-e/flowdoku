// Landing Page - Startbildschirm mit animiertem Orb und Willkommensnachricht

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
import { useCurrentUser } from '../context/UserContext';
import { useEvents } from '../context/EventContext';
import { useRoles } from '../context/RoleContext';
import { APP_VERSION } from '../version';
import { Info, Mail } from './Icons';
import LandingInfoModal from './LandingInfoModal';

export default function LandingPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { locale, setLocale, t } = useLanguage();
  const { currentUser } = useCurrentUser();
  const { sendAdminInquiry, events } = useEvents();
  const { isAdmin } = useRoles();
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

  // v6.22: Mobile-QR-Shortcut. Organizer/Admin/QR-Scanner sehen auf Mobile-
  // Geräten oben rechts ein QR-Icon + Sprechblase "Geht's zum Check-in".
  // Navigation:
  //  - genau 1 Event zugänglich → direkt auf die Scanner-Seite
  //  - >1 Event oder Admin → zur allgemeinen Check-In-Seite (Event-Picker)
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const checkInEvents = React.useMemo(() => {
    if (!currentEmailLc && !isAdmin) return [];
    return (events || []).filter(e => {
      if (isAdmin) return true;
      const orgMatch = (e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
      const qrMatch = (e.qrScannerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
      return orgMatch || qrMatch;
    });
  }, [events, currentEmailLc, isAdmin]);
  const [isMobile, setIsMobile] = React.useState<boolean>(() => {
    if (typeof window === 'undefined') return false;
    return window.matchMedia && window.matchMedia('(max-width: 768px)').matches;
  });
  React.useEffect(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return;
    const mq = window.matchMedia('(max-width: 768px)');
    const handler = (e: MediaQueryListEvent): void => setIsMobile(e.matches);
    // addEventListener bevorzugt; addListener als Fallback für ältere Browser
    if (mq.addEventListener) mq.addEventListener('change', handler);
    // eslint-disable-next-line @typescript-eslint/no-deprecated
    else mq.addListener(handler);
    return () => {
      if (mq.removeEventListener) mq.removeEventListener('change', handler);
      // eslint-disable-next-line @typescript-eslint/no-deprecated
      else mq.removeListener(handler);
    };
  }, []);
  const showCheckInShortcut = isMobile && checkInEvents.length > 0;
  const goToCheckIn = (): void => {
    if (checkInEvents.length === 1 && !isAdmin) {
      navigate('check-in', checkInEvents[0].id);
    } else {
      navigate('check-in');
    }
  };

  return (
    <div className="landing" style={{ position: 'relative' }}>
      <span style={{
        position: 'absolute', top: 12, right: 16,
        fontSize: '0.7rem', color: 'var(--dex-gray-300)',
      }}>
        v{APP_VERSION}
      </span>

      {/* v6.22: Mobile-QR-Shortcut für Organizer/Admin/QR-Scanner.
          Sprechblase + QR-Icon oben rechts — ein Tap führt direkt zum Check-In
          (oder zum Event-Picker, wenn mehrere Events eingecheckt werden können). */}
      {showCheckInShortcut && (
        <div style={{
          position: 'absolute', top: 30, right: 12, zIndex: 10,
          display: 'flex', alignItems: 'center', gap: 8,
        }}>
          <button
            type="button"
            onClick={goToCheckIn}
            style={{
              background: 'var(--dex-green)', color: '#fff',
              padding: '8px 12px', borderRadius: 12,
              fontSize: '0.78rem', fontWeight: 600, lineHeight: 1.2,
              border: 'none', cursor: 'pointer',
              boxShadow: '0 4px 12px rgba(134,188,37,0.28)',
              position: 'relative',
              fontFamily: 'inherit',
              whiteSpace: 'nowrap',
            }}
            aria-label={locale === 'de' ? 'Zum Check-in' : 'Go to check-in'}
          >
            {locale === 'de' ? "Geht's zum Check-in" : 'Go to check-in'}
            {/* Kleines Dreieck als Sprechblasen-Spitze */}
            <span style={{
              position: 'absolute', top: '50%', right: -6, transform: 'translateY(-50%)',
              width: 0, height: 0,
              borderTop: '6px solid transparent',
              borderBottom: '6px solid transparent',
              borderLeft: '6px solid var(--dex-green)',
            }} />
          </button>
          <button
            type="button"
            onClick={goToCheckIn}
            style={{
              background: 'var(--dex-green)', color: '#fff',
              border: 'none', borderRadius: '50%',
              width: 44, height: 44,
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              cursor: 'pointer',
              boxShadow: '0 4px 12px rgba(134,188,37,0.35)',
              padding: 0,
              flexShrink: 0,
            }}
            aria-label={locale === 'de' ? 'Zum Check-in' : 'Go to check-in'}
          >
            {/* Inline QR-Icon (quadratisches QR-Muster) */}
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
              <rect x="3" y="3" width="7" height="7" stroke="currentColor" strokeWidth="2" />
              <rect x="14" y="3" width="7" height="7" stroke="currentColor" strokeWidth="2" />
              <rect x="3" y="14" width="7" height="7" stroke="currentColor" strokeWidth="2" />
              <rect x="6" y="6" width="1.5" height="1.5" fill="currentColor" />
              <rect x="17" y="6" width="1.5" height="1.5" fill="currentColor" />
              <rect x="6" y="17" width="1.5" height="1.5" fill="currentColor" />
              <rect x="14" y="14" width="2" height="2" fill="currentColor" />
              <rect x="18" y="14" width="2" height="2" fill="currentColor" />
              <rect x="14" y="18" width="2" height="2" fill="currentColor" />
              <rect x="18" y="18" width="2" height="2" fill="currentColor" />
            </svg>
          </button>
        </div>
      )}
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
            {locale === 'de' ? 'Built by ' : 'Built by '}
            <strong style={{ fontWeight: 600, color: 'var(--dex-gray-500)' }}>
              Eike Brenneisen, Andreas Enk {locale === 'de' ? 'und' : 'and'} Nils Felten
            </strong>
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
