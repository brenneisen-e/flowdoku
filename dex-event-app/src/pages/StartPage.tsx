import { useNavigate } from 'react-router-dom';
import { Calendar, Pin } from 'lucide-react';

export default function StartPage() {
  const navigate = useNavigate();

  return (
    <div className="page-container">
      <div className="start-grid">
        <div className="card card-clickable start-card" onClick={() => navigate('/register')}>
          <div className="start-card__icon">
            <Calendar size={64} strokeWidth={1} />
          </div>
          <h2>Registration</h2>
          <p>Register for a Deloitte Event</p>
        </div>
        <div className="card card-clickable start-card" onClick={() => navigate('/my-events')}>
          <div className="start-card__icon">
            <Pin size={64} strokeWidth={1} />
          </div>
          <h2>My Events</h2>
          <p>Check Registration Status / Cancel</p>
        </div>
      </div>
    </div>
  );
}
