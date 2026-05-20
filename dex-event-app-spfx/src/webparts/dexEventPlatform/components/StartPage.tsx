// Start-Seite - Navigation zu Registration, My Events und Admin

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useLanguage } from '../context/LanguageContext';
import { Calendar, Pin, Settings } from './Icons';

export default function StartPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { canCreateEvents } = useRoles();
  const { events } = useEvents();
  const { currentUser } = useCurrentUser();
  const { t } = useLanguage();

  // v11.38: Co-Organizer pro Event (per-Event-Rolle, ohne globale Organizer-
  // Rolle in DEX_Roles) sehen die Organizer-Kachel ebenfalls — AdminPage
  // gewaehrt ihnen ohnehin Zugriff auf "ihre" Events (siehe isOrganizerFor
  // dort), aber ohne Kachel im Startmenue gab es bisher keinen Einstieg.
  const currentEmailLc = (currentUser.email || '').toLowerCase();
  const isCoOrganizerOfAny = !!currentEmailLc && (events || []).some(
    e => (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc),
  );
  const showAdminTile = canCreateEvents || isCoOrganizerOfAny;

  return (
    <div className="page-container">
      {/* v9.36: Keyframes inline injizieren — SPFx hasht sonst die Names im
          .module.scss und die Animation findet sie nicht. */}
      <style>{`
        @keyframes dexStartIconBounce {
          0%   { transform: translateY(0) scale(1); }
          30%  { transform: translateY(-8px) scale(1.1); }
          60%  { transform: translateY(0) scale(1); }
          80%  { transform: translateY(-3px) scale(1.04); }
          100% { transform: translateY(0) scale(1); }
        }
        @keyframes dexStartIconWiggle {
          0%   { transform: rotate(0deg) scale(1); }
          20%  { transform: rotate(-10deg) scale(1.08); }
          40%  { transform: rotate(8deg) scale(1.08); }
          60%  { transform: rotate(-6deg) scale(1.05); }
          80%  { transform: rotate(4deg) scale(1.03); }
          100% { transform: rotate(0deg) scale(1); }
        }
        @keyframes dexStartIconSpin {
          0%   { transform: rotate(0deg) scale(1); }
          100% { transform: rotate(360deg) scale(1); }
        }
      `}</style>
      <div className={`start-grid${showAdminTile ? ' start-grid--with-admin' : ''}`}>
        <div className="card card-clickable start-card" onClick={() => navigate('register')}>
          <div className="start-card__icon">
            <Calendar size={64} strokeWidth={1} />
          </div>
          <h2>{t('start.register')}</h2>
          <p style={{ whiteSpace: 'nowrap' }}>{t('start.register.desc')}</p>
        </div>
        <div className="card card-clickable start-card" onClick={() => navigate('my-events')}>
          <div className="start-card__icon">
            <Pin size={64} strokeWidth={1} />
          </div>
          <h2>{t('start.myevents')}</h2>
          <p style={{ whiteSpace: 'nowrap' }}>{t('start.myevents.desc')}</p>
        </div>
        {showAdminTile && (
          <div className="card card-clickable start-card start-card--admin" onClick={() => navigate('admin')}>
            <div className="start-card__icon">
              <Settings size={64} strokeWidth={1} />
            </div>
            <h2>{t('start.admin')}</h2>
            <p style={{ whiteSpace: 'nowrap' }}>{t('start.admin.desc')}</p>
          </div>
        )}
      </div>
    </div>
  );
}
