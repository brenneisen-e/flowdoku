// Landing Page - Startbildschirm mit animiertem Orb und Willkommensnachricht

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
import { APP_VERSION } from '../version';
import { Info, Mail } from './Icons';

export default function LandingPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { locale, setLocale, t } = useLanguage();
  const [showInfo, setShowInfo] = React.useState(true);

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

  return (
    <div className="landing" style={{ position: 'relative' }}>
      <span style={{
        position: 'absolute', top: 12, right: 16,
        fontSize: '0.7rem', color: 'var(--dex-gray-300)',
      }}>
        v{APP_VERSION}
      </span>
      {/* Sprachauswahl */}
      <div style={{
        position: 'absolute', top: 12, left: 16,
        display: 'flex', gap: 4,
      }}>
        <button
          onClick={() => setLocale('de')}
          style={{
            background: locale === 'de' ? 'var(--dex-green)' : 'rgba(255,255,255,0.1)',
            border: 'none',
            borderRadius: 6, padding: '6px 12px', cursor: 'pointer', fontSize: '0.8rem',
            color: locale === 'de' ? '#fff' : 'var(--dex-gray-400)',
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
            background: locale === 'en' ? 'var(--dex-green)' : 'rgba(255,255,255,0.1)',
            border: 'none',
            borderRadius: 6, padding: '6px 12px', cursor: 'pointer', fontSize: '0.8rem',
            color: locale === 'en' ? '#fff' : 'var(--dex-gray-400)',
            fontWeight: locale === 'en' ? 700 : 500,
            transition: 'all 0.2s',
          }}
          title="English"
        >
          EN
        </button>
      </div>
      <div className="landing__hero">
        <div className="landing__orb">
          <div className="landing__orb-inner" />
        </div>
        <div className="landing__text">
          <h1>
            {t('landing.welcome')} <strong>{t('landing.platform')}</strong>.
          </h1>
          <p>{t('landing.subtitle')}</p>
        </div>
        <button className="btn btn-lg btn-block btn-outline" onClick={() => navigate('start')}>
          {t('landing.start')}
        </button>
        <div style={{ display: 'flex', gap: 16, marginTop: 20, justifyContent: 'center', alignItems: 'center' }}>
          <button
            onClick={() => setShowInfo(!showInfo)}
            style={{
              background: 'none', border: '2px solid var(--dex-gray-300)', borderRadius: '50%',
              width: 40, height: 40, display: 'flex', alignItems: 'center', justifyContent: 'center',
              cursor: 'pointer', color: 'var(--dex-gray-400)', transition: 'all 0.2s',
            }}
            title="Über die App"
          >
            <Info size={18} />
          </button>
          <a
            href="mailto:ebrenneisen@deloitte.de;nifelten@deloitte.de;aenk@deloitte.de?subject=DEX Event Experience Platform – Feedback"
            style={{
              background: 'none', border: '2px solid var(--dex-gray-300)', borderRadius: '50%',
              width: 40, height: 40, display: 'flex', alignItems: 'center', justifyContent: 'center',
              cursor: 'pointer', color: 'var(--dex-gray-400)', textDecoration: 'none', transition: 'all 0.2s',
            }}
            title="Kontakt aufnehmen"
          >
            <Mail size={18} />
          </a>
        </div>
        {showInfo && (
          <p style={{
            color: 'var(--dex-gray-500)', fontSize: '0.78rem', marginTop: 12,
            textAlign: 'center', lineHeight: 1.5, maxWidth: 600,
            background: 'var(--dex-white)', padding: '12px 20px',
            borderRadius: 'var(--dex-radius-lg)', border: '1px solid var(--dex-gray-200)',
          }}>
            The Event Experience Platform is a solution for managing participants at Deloitte events.
            Developed by Eike Brenneisen, Andreas Enk and Nils Felten. Currently in pilot phase.
          </p>
        )}
      </div>
    </div>
  );
}
