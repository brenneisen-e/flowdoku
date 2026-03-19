/**
 * Header-Komponente
 *
 * Sticky Header mit dynamischem Titel je nach Route.
 * Auf der Landing Page wird das Deloitte-Logo angezeigt,
 * auf allen anderen Seiten ein Zurueck-Button.
 */

import { useNavigate, useLocation } from 'react-router-dom';
import { ChevronLeft, Settings, Mail } from 'lucide-react';
import { currentUser } from '../data/mockData';

export default function Header() {
  const navigate = useNavigate();
  const location = useLocation();
  const isLanding = location.pathname === '/';
  const isStart = location.pathname === '/start';

  // Titel-Mapping je nach aktuellem Pfad
  const getTitle = () => {
    const path = location.pathname;
    if (path === '/start') return 'Event Experience Platform';
    if (path === '/register') return 'Registration | Available events at your location';
    if (path.startsWith('/register/')) return 'Registration Deloitte Events';
    if (path === '/my-events') return 'My Events';
    if (path === '/create-event') return 'Deloitte Event Creation';
    if (path === '/settings') return 'Settings';
    return '';
  };

  return (
    <header className="header">
      <div className="header-left">
        {isLanding ? (
          <div className="header-logo">
            Deloitte<span>.</span>
          </div>
        ) : (
          <>
            <button className="back-btn" onClick={() => navigate(-1)} aria-label="Back">
              <ChevronLeft size={20} />
            </button>
            <span className="header-title" style={{ border: 'none', paddingLeft: 0, fontWeight: 500 }}>
              {getTitle()}
            </span>
          </>
        )}
      </div>
      <div className="header-right">
        {isLanding && (
          <button className="header-icon-btn" onClick={() => navigate('/start')}>
            <Mail size={20} />
          </button>
        )}
        {!isLanding && !isStart && (
          <button className="header-icon-btn" onClick={() => navigate('/settings')}>
            <Settings size={20} />
          </button>
        )}
        {/* User-Avatar mit Initialen */}
        <div className="header-avatar" title={`${currentUser.firstName} ${currentUser.surname}`}>
          {currentUser.firstName[0]}{currentUser.surname[0]}
        </div>
      </div>
    </header>
  );
}
