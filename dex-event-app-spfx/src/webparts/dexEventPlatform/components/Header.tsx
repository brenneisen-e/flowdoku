/**
 * Header-Komponente
 *
 * Sticky Header mit dynamischem Titel je nach aktueller Seite.
 * Auf der Landing Page wird das Deloitte-Logo angezeigt,
 * auf allen anderen Seiten ein Zurueck-Button.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { ChevronLeft, Settings, Mail } from './Icons';
import { currentUser } from '../data/mockData';

export default function Header(): React.ReactElement {
  const { currentPage, navigate, goBack } = useNavigation();
  const isLanding = currentPage === 'landing';
  const isStart = currentPage === 'start';

  // Titel-Mapping je nach aktuellem Seitenstatus
  const getTitle = (): string => {
    switch (currentPage) {
      case 'start': return 'Event Experience Platform';
      case 'register': return 'Registration | Available events at your location';
      case 'registration': return 'Registration Deloitte Events';
      case 'my-events': return 'My Events';
      case 'create-event': return 'Deloitte Event Creation';
      case 'settings': return 'Settings';
      default: return '';
    }
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
            <button className="back-btn" onClick={() => goBack()} aria-label="Back">
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
          <button className="header-icon-btn" onClick={() => navigate('start')}>
            <Mail size={20} />
          </button>
        )}
        {!isLanding && !isStart && (
          <button className="header-icon-btn" onClick={() => navigate('settings')}>
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
