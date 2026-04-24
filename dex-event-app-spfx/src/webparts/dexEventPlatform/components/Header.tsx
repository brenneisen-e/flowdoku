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
import { useEvents } from '../context/EventContext';
import { useLanguage } from '../context/LanguageContext';
import { ChevronLeft, Settings, Book, QrCode } from './Icons';

export default function Header(): React.ReactElement {
  const { currentPage, navigate } = useNavigation();
  const { currentUser, photoUrl } = useCurrentUser();
  const { currentUserRole, isAdmin, isOrganizer } = useRoles();
  const { events } = useEvents();
  // Check-In-Button (Admin / Organizer / QR-Scanner): schneller Einstieg in den
  // QR-Scanner ohne vorher ein konkretes Event auszuwaehlen. CheckInPage liest
  // das Event aus dem gescannten QR-Code selbst (`DEX|<eventNumber>|<email>`).
  // v6.26: Zugriff auch fuer User, die per E-Mail in event.qrScannerEmails
  // mindestens eines Events eingetragen sind — ohne globale Organizer-Rolle.
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isQRScannerOfAny = !!currentEmailLc && (events || []).some(
    e => (e.qrScannerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc),
  );
  const canCheckIn = isAdmin || isOrganizer || isQRScannerOfAny;
  const { t } = useLanguage();
  const [showPopup, setShowPopup] = React.useState(false);
  const isLanding = currentPage === 'landing';
  const isStart = currentPage === 'start';

  // v6.26: Mobile-Detection fuer die "Jetzt einchecken"-Sprechblase neben dem
  // QR-Icon. Wird nur auf Mobilgeraeten angezeigt (Viewport <= 768px), auf
  // Desktop bleibt der Header schlank.
  const [isMobile, setIsMobile] = React.useState<boolean>(() => {
    if (typeof window === 'undefined') return false;
    return !!(window.matchMedia && window.matchMedia('(max-width: 768px)').matches);
  });
  React.useEffect(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return;
    const mq = window.matchMedia('(max-width: 768px)');
    const handler = (e: MediaQueryListEvent): void => setIsMobile(e.matches);
    if (mq.addEventListener) mq.addEventListener('change', handler);
    // eslint-disable-next-line @typescript-eslint/no-deprecated
    else mq.addListener(handler);
    return () => {
      if (mq.removeEventListener) mq.removeEventListener('change', handler);
      // eslint-disable-next-line @typescript-eslint/no-deprecated
      else mq.removeListener(handler);
    };
  }, []);

  // Titel-Mapping je nach aktuellem Seitenstatus
  const getTitle = (): string => {
    switch (currentPage) {
      case 'start': return t('start.title');
      case 'register': return t('header.registration');
      case 'registration': return 'Registration Deloitte Events';
      case 'my-events': return t('myevents.title');
      case 'create-event': return t('header.createevent');
      case 'edit-event': return t('header.editevent');
      case 'settings': return t('header.settings');
      case 'profile': return t('header.profile');
      case 'admin': return t('header.admin');
      case 'role-matrix': return t('header.rolematrix');
      case 'participants': return t('header.participants');
      case 'flowcharts': return t('header.flowcharts');
      case 'check-in': return t('header.checkin');
      case 'manual': return t('header.manual');
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
            <button className="back-btn" onClick={() => navigate(isStart ? 'landing' : 'start')} aria-label="Back">
              <ChevronLeft size={20} />
            </button>
            <span className="header-title" style={{ border: 'none', paddingLeft: 0, fontWeight: 500 }}>
              {getTitle()}
            </span>
          </>
        )}
      </div>
      <div className="header-right">
        {canCheckIn && isLanding && isMobile && (
          <button
            type="button"
            onClick={() => navigate('check-in')}
            aria-label={t('header.checkin')}
            style={{
              background: 'var(--dex-green)', color: '#fff',
              border: 'none', borderRadius: 12,
              padding: '6px 10px',
              fontSize: '0.78rem', fontWeight: 600, lineHeight: 1.2,
              cursor: 'pointer', marginRight: 6,
              boxShadow: '0 2px 8px rgba(134,188,37,0.28)',
              fontFamily: 'inherit', whiteSpace: 'nowrap',
              position: 'relative',
            }}
          >
            {t('header.checkin.bubble') || 'Jetzt einchecken'}
            <span style={{
              position: 'absolute', top: '50%', right: -6, transform: 'translateY(-50%)',
              width: 0, height: 0,
              borderTop: '6px solid transparent',
              borderBottom: '6px solid transparent',
              borderLeft: '6px solid var(--dex-green)',
            }} />
          </button>
        )}
        {canCheckIn && (
          <button
            className="header-icon-btn"
            onClick={() => navigate('check-in')}
            title={t('header.checkin')}
            style={currentPage === 'check-in' ? { background: 'var(--dex-gray-200)' } : {}}
            aria-label={t('header.checkin')}
          >
            <QrCode size={20} />
          </button>
        )}
        <button
          className="header-icon-btn"
          onClick={() => navigate('manual')}
          title={t('header.manual')}
          style={currentPage === 'manual' ? { background: 'var(--dex-gray-200)' } : {}}
        >
          <Book size={20} />
        </button>
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
          {photoUrl ? (
            <img
              src={photoUrl}
              alt={`${currentUser.firstName} ${currentUser.surname}`}
              onClick={() => setShowPopup(!showPopup)}
              style={{ width: 36, height: 36, borderRadius: '50%', objectFit: 'cover', cursor: 'pointer' }}
              onError={(e) => { (e.target as HTMLImageElement).style.display = 'none'; }}
            />
          ) : (
            <div
              className="header-avatar"
              title={`${currentUser.firstName} ${currentUser.surname}`}
              onClick={() => setShowPopup(!showPopup)}
              style={{ cursor: 'pointer' }}
            >
              {currentUser.firstName ? currentUser.firstName[0] : ''}{currentUser.surname ? currentUser.surname[0] : ''}
            </div>
          )}
          {showPopup && (
            <div style={{
              position: 'absolute', right: 0, top: '100%', marginTop: 8,
              background: '#fff', borderRadius: 'var(--dex-radius-lg, 12px)',
              boxShadow: '0 8px 32px rgba(0,0,0,0.18)', padding: '20px 24px',
              minWidth: 260, zIndex: 1000,
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16 }}>
                {photoUrl ? (
                  <img src={photoUrl} alt="Profile" style={{ width: 48, height: 48, borderRadius: '50%', objectFit: 'cover' }} />
                ) : (
                  <div style={{
                    width: 48, height: 48, borderRadius: '50%',
                    background: 'linear-gradient(135deg, #86bc25, #0076a8)',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    color: '#fff', fontWeight: 700, fontSize: '1.1rem',
                  }}>
                    {currentUser.firstName ? currentUser.firstName[0] : ''}{currentUser.surname ? currentUser.surname[0] : ''}
                  </div>
                )}
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
                  {t('profile.viewfull')}
                </button>
                {isAdmin && (
                  <button
                    className="btn btn-primary btn-block"
                    style={{ fontSize: '0.85rem' }}
                    onClick={() => { setShowPopup(false); navigate('settings'); }}
                  >
                    <Settings size={14} /> {t('settings.rolemanagement')}
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
