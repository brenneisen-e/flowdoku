/**
 * Header-Komponente
 *
 * Sticky Header mit dynamischem Titel je nach aktueller Seite.
 * Auf der Landing Page wird das Deloitte-Logo angezeigt,
 * auf allen anderen Seiten ein Zurueck-Button.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { ChevronLeft, Settings, Mail } from './Icons';

export default function Header(): React.ReactElement {
  const { currentPage, navigate, goBack } = useNavigation();
  const { currentUser } = useCurrentUser();
  const { currentUserRole, isAdmin } = useRoles();
  const [showPopup, setShowPopup] = React.useState(false);
  const isLanding = currentPage === 'landing';

  // Titel-Mapping je nach aktuellem Seitenstatus
  const getTitle = (): string => {
    switch (currentPage) {
      case 'start': return 'Event Experience Platform';
      case 'register': return 'Registration | Available events at your location';
      case 'registration': return 'Registration Deloitte Events';
      case 'my-events': return 'My Events';
      case 'create-event': return 'Deloitte Event Creation';
      case 'edit-event': return 'Event bearbeiten';
      case 'settings': return 'Settings';
      case 'profile': return 'My Profile';
      case 'admin': return 'Admin';
      case 'role-matrix': return 'Rollen-Matrix';
      default: return '';
    }
  };

  // Rollen-Farbe
  const roleColors: Record<string, { bg: string; color: string }> = {
    'Admin': { bg: '#e8f5e9', color: '#2e7d32' },
    'Organizer': { bg: '#e3f2fd', color: '#1565c0' },
    'User': { bg: '#f5f5f5', color: '#666' },
  };
  const rc = roleColors[currentUserRole] || roleColors['User'];

  return (
    <header className="header">
      <div className="header-left">
        {isLanding ? (
          <div className="header-logo">
            Deloitte<span>.</span>
          </div>
        ) : (
          <>
            <button className="back-btn" onClick={() => navigate('start')} aria-label="Back">
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
        {!isLanding && (
          <button
            className="header-icon-btn"
            onClick={() => navigate('settings')}
            title="Settings"
            style={currentPage === 'settings' ? { background: 'var(--dex-gray-200)' } : {}}
          >
            <Settings size={20} />
          </button>
        )}
        {/* User-Avatar mit Initialen + Popup */}
        <div style={{ position: 'relative' }}>
          <div
            className="header-avatar"
            title={`${currentUser.firstName} ${currentUser.surname}`}
            onClick={() => setShowPopup(!showPopup)}
            style={{ cursor: 'pointer' }}
          >
            {currentUser.firstName ? currentUser.firstName[0] : ''}{currentUser.surname ? currentUser.surname[0] : ''}
          </div>
          {showPopup && (
            <div style={{
              position: 'absolute', right: 0, top: '100%', marginTop: 8,
              background: '#fff', borderRadius: 'var(--dex-radius-lg, 12px)',
              boxShadow: '0 8px 32px rgba(0,0,0,0.18)', padding: '20px 24px',
              minWidth: 260, zIndex: 1000,
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16 }}>
                <div style={{
                  width: 48, height: 48, borderRadius: '50%',
                  background: 'linear-gradient(135deg, #86bc25, #0076a8)',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  color: '#fff', fontWeight: 700, fontSize: '1.1rem',
                }}>
                  {currentUser.firstName ? currentUser.firstName[0] : ''}{currentUser.surname ? currentUser.surname[0] : ''}
                </div>
                <div>
                  <div style={{ fontWeight: 600, fontSize: '1rem' }}>{currentUser.firstName} {currentUser.surname}</div>
                  <div style={{ color: '#666', fontSize: '0.85rem' }}>{currentUser.email}</div>
                </div>
              </div>
              {currentUser.location && (
                <div style={{ fontSize: '0.85rem', color: '#555', marginBottom: 8 }}>
                  Location: {currentUser.location}
                </div>
              )}
              <div style={{ fontSize: '0.85rem', marginBottom: 16 }}>
                <span style={{
                  display: 'inline-block', padding: '2px 10px', borderRadius: 12,
                  background: rc.bg, color: rc.color,
                  fontSize: '0.8rem', fontWeight: 500,
                }}>
                  {currentUserRole}
                </span>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                <button
                  className="btn btn-secondary btn-block"
                  style={{ fontSize: '0.85rem' }}
                  onClick={() => { setShowPopup(false); navigate('profile'); }}
                >
                  View full profile
                </button>
                {isAdmin && (
                  <button
                    className="btn btn-primary btn-block"
                    style={{ fontSize: '0.85rem' }}
                    onClick={() => { setShowPopup(false); navigate('settings'); }}
                  >
                    <Settings size={14} /> Role Management
                  </button>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    </header>
  );
}
