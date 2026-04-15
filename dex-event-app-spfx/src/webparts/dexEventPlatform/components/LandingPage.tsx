// Landing Page - Startbildschirm mit animiertem Orb und Willkommensnachricht

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
import { APP_VERSION } from '../version';
import { Info, Mail } from './Icons';
import LandingInfoModal from './LandingInfoModal';

export default function LandingPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { locale, setLocale, t } = useLanguage();
  const [showInfo, setShowInfo] = React.useState(false);

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
        <div className="landing__card">
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
          <div style={{ display: 'flex', gap: 16, justifyContent: 'center', alignItems: 'center' }}>
            <button
              onClick={() => setShowInfo(!showInfo)}
              style={{
                background: 'none', border: '2px solid var(--dex-gray-300)', borderRadius: '50%',
                width: 40, height: 40, display: 'flex', alignItems: 'center', justifyContent: 'center',
                cursor: 'pointer', color: 'var(--dex-gray-400)', transition: 'all 0.2s',
              }}
              title={locale === 'de' ? 'Über die App' : 'About the app'}
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
              title={locale === 'de' ? 'Kontakt aufnehmen' : 'Get in touch'}
            >
              <Mail size={18} />
            </a>
          </div>
        </div>
      </div>
      <LandingInfoModal open={showInfo} locale={locale === 'de' ? 'de' : 'en'} onClose={() => setShowInfo(false)} />
    </div>
  );
}
