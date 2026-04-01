// Landing Page - Startbildschirm mit animiertem Orb und Willkommensnachricht

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { APP_VERSION } from '../version';
import { Info, Mail } from './Icons';

export default function LandingPage(): React.ReactElement {
  const { navigate } = useNavigation();
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
        @keyframes dexOrbPulse {
          0%, 100% { transform: rotate(0deg) scale(1); }
          50% { transform: rotate(180deg) scale(1.06); }
        }
        @keyframes dexOrbGlow {
          0%, 100% { filter: drop-shadow(0 0 8px rgba(134, 188, 37, 0.3)); }
          50% { filter: drop-shadow(0 0 24px rgba(0, 118, 168, 0.5)); }
        }
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
      <div className="landing__hero">
        <div className="landing__orb">
          <div className="landing__orb-inner" />
        </div>
        <div className="landing__text">
          <h1>
            Welcome to the new <strong>Event Experience Platform</strong>.
          </h1>
          <p>Enjoy the new app to handle your registration for your Deloitte Event.</p>
        </div>
        <button className="btn btn-lg btn-block btn-outline" onClick={() => navigate('start')}>
          Start
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
