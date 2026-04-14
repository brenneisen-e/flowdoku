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
      {/* Hinweis-Box rechts: neue Version + Link zur alten App */}
      <div className="landing__notice" style={{
        position: 'absolute', top: 48, right: 16,
        maxWidth: 560, padding: '32px 36px',
        background: 'rgba(255,255,255,0.95)', borderRadius: 'var(--dex-radius-lg)',
        border: '1px solid var(--dex-gray-200)',
        boxShadow: '0 4px 16px rgba(0,0,0,0.10)',
        fontSize: '1rem', lineHeight: 1.55, color: 'var(--dex-gray-700)',
        zIndex: 5,
      }}>
        <div style={{ fontWeight: 700, color: 'var(--dex-green)', marginBottom: 14, fontSize: '1.35rem' }}>
          {locale === 'de' ? 'Neue Version der DEX App' : 'New version of the DEX App'}
        </div>
        <div style={{ marginBottom: 20 }}>
          {locale === 'de'
            ? 'Dies ist die neue Version der DEX App. Falls du Zugriff auf ein altes Event benötigst, das noch nicht über die neue Version verwaltet wird, nutze bitte die alte App.'
            : 'This is the new version of the DEX App. If you need access to an older event that is not yet managed via the new version, please use the legacy app.'}
        </div>
        <a
          href="https://apps.powerapps.com/play/e/5bbcd5e1-8573-e5f8-b7b1-e4866693b255/a/796780aa-feb8-4579-b1aa-1978a4faa85b?tenantId=36da45f1-dd2c-4d1f-af13-5abe46b99921&hint=0dfc47c9-7457-463f-851a-1002de314739&sourcetime=1766053193836"
          target="_blank"
          rel="noopener noreferrer"
          style={{
            display: 'inline-block', padding: '12px 24px', borderRadius: 8,
            background: 'var(--dex-green)', color: '#fff', textDecoration: 'none',
            fontWeight: 600, fontSize: '0.95rem',
          }}
        >
          {locale === 'de' ? 'Alte App öffnen' : 'Open legacy app'}
        </a>
      </div>
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
      <LandingInfoModal open={showInfo} locale={locale === 'de' ? 'de' : 'en'} onClose={() => setShowInfo(false)} />
    </div>
  );
}
