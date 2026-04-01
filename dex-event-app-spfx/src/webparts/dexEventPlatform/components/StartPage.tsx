// Start-Seite - Navigation zu Registration, My Events und Admin

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { Calendar, Pin, Settings } from './Icons';

export default function StartPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { canCreateEvents } = useRoles();
  const { t } = useLanguage();

  return (
    <div className="page-container">
      <div className="start-grid" style={canCreateEvents ? { gridTemplateColumns: '1fr 1fr 1fr' } : undefined}>
        <div className="card card-clickable start-card" onClick={() => navigate('register')}>
          <div className="start-card__icon">
            <Calendar size={64} strokeWidth={1} />
          </div>
          <h2>{t('start.register')}</h2>
          <p>{t('start.register.desc')}</p>
        </div>
        <div className="card card-clickable start-card" onClick={() => navigate('my-events')}>
          <div className="start-card__icon">
            <Pin size={64} strokeWidth={1} />
          </div>
          <h2>{t('start.myevents')}</h2>
          <p>{t('start.myevents.desc')}</p>
        </div>
        {canCreateEvents && (
          <div className="card card-clickable start-card" onClick={() => navigate('admin')}>
            <div className="start-card__icon">
              <Settings size={64} strokeWidth={1} />
            </div>
            <h2>{t('start.admin')}</h2>
            <p>{t('start.admin.desc')}</p>
          </div>
        )}
      </div>
    </div>
  );
}
