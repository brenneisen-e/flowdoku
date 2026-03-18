import { useNavigate } from 'react-router-dom';
import { currentUser } from '../data/mockData';
import { Plus, FileText, Users } from 'lucide-react';

export default function SettingsPage() {
  const navigate = useNavigate();

  return (
    <div className="page-container">
      <div className="settings-grid">
        {/* User Info */}
        <div className="card">
          <h3 className="mb-16">User Information</h3>
          <div className="settings-info">
            <div className="settings-info__row">
              <span className="settings-info__label">Name</span>
              <span>{currentUser.firstName} {currentUser.surname}</span>
            </div>
            <div className="settings-info__row">
              <span className="settings-info__label">Email</span>
              <span>{currentUser.email}</span>
            </div>
            <div className="settings-info__row">
              <span className="settings-info__label">Location</span>
              <span>{currentUser.location}</span>
            </div>
            <div className="settings-info__row">
              <span className="settings-info__label">Role</span>
              <span className={`badge ${currentUser.isAdmin ? 'badge-green' : 'badge-gray'}`}>
                {currentUser.isAdmin ? 'Admin' : 'User'}
              </span>
            </div>
          </div>
        </div>

        {/* Admin Actions */}
        {currentUser.isAdmin && (
          <div className="card">
            <h3 className="mb-16">Admin Actions</h3>
            <div className="settings-actions">
              <button className="btn btn-primary btn-block" onClick={() => navigate('/create-event')}>
                <Plus size={18} /> Create New Event
              </button>
              <button className="btn btn-secondary btn-block mt-8">
                <FileText size={18} /> View All Events (Admin)
              </button>
              <button className="btn btn-secondary btn-block mt-8">
                <Users size={18} /> Extract Mail Addresses
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
