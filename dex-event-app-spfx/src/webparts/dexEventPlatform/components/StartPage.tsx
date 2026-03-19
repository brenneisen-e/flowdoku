// Start-Seite - Navigation zu Registration, My Events und Admin

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { Calendar, Pin, Settings } from './Icons';

export default function StartPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { canCreateEvents } = useRoles();

  return (
    <div className="page-container">
      <div className="start-grid" style={canCreateEvents ? { gridTemplateColumns: '1fr 1fr 1fr' } : undefined}>
        <div className="card card-clickable start-card" onClick={() => navigate('register')}>
          <div className="start-card__icon">
            <Calendar size={64} strokeWidth={1} />
          </div>
          <h2>Registration</h2>
          <p>Register for a Deloitte Event</p>
        </div>
        <div className="card card-clickable start-card" onClick={() => navigate('my-events')}>
          <div className="start-card__icon">
            <Pin size={64} strokeWidth={1} />
          </div>
          <h2>My Events</h2>
          <p>Check Registration Status / Cancel</p>
        </div>
        {canCreateEvents && (
          <div className="card card-clickable start-card" onClick={() => navigate('admin')}>
            <div className="start-card__icon">
              <Settings size={64} strokeWidth={1} />
            </div>
            <h2>Admin / Organizer</h2>
            <p>Manage Events &amp; Participants</p>
          </div>
        )}
      </div>
    </div>
  );
}
