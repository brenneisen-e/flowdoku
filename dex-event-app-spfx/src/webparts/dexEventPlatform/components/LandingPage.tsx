// Landing Page - Startbildschirm mit animiertem Orb und Willkommensnachricht

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { APP_VERSION } from '../version';

export default function LandingPage(): React.ReactElement {
  const { navigate } = useNavigation();

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
        <p style={{ color: 'var(--dex-gray-400)', fontSize: '0.75rem', marginTop: 16, textAlign: 'center', lineHeight: 1.4 }}>
          The Event Experience Platform is a solution for managing participants at Deloitte events.
          Developed by Eike Brenneisen, Andreas Enk and Nils Felten. Currently in pilot phase &ndash;
          questions or feedback? Feel free to get in touch!
        </p>
      </div>
    </div>
  );
}
