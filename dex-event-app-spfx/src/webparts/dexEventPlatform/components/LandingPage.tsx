// Landing Page - Startbildschirm mit animiertem Orb und Willkommensnachricht

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { APP_VERSION } from '../version';

export default function LandingPage(): React.ReactElement {
  const { navigate } = useNavigation();

  return (
    <div className="landing">
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
        <button className="btn btn-lg btn-block btn-secondary" onClick={() => navigate('start')}>
          Start
        </button>
      </div>
      <footer style={{
        display: 'flex', justifyContent: 'space-between', alignItems: 'center',
        padding: '8px 24px', fontSize: '0.7rem', color: 'var(--dex-gray-400, #a0a0a0)',
        borderTop: '1px solid var(--dex-gray-200, #eee)', flexShrink: 0,
      }}>
        <span>Developed by Eike Brenneisen, Andreas Enk &amp; Nils Felten</span>
        <span>v{APP_VERSION}</span>
      </footer>
    </div>
  );
}
