import { useNavigate } from 'react-router-dom';
import { currentUser } from '../data/mockData';
import { Plus, FileText, Users, Download } from 'lucide-react';

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
        {/* Download App */}
        <div className="card">
          <h3 className="mb-16">Offline Version</h3>
          <p style={{ color: 'var(--dex-gray-600)', fontSize: '0.9rem', marginBottom: 16 }}>
            Download the complete application as a ZIP file to run it locally on your machine.
          </p>
          <button
            className="btn btn-primary btn-block"
            onClick={() => window.open('https://github.com/brenneisen-e/flowdoku/archive/refs/heads/main.zip', '_blank')}
          >
            <Download size={18} /> Download App as ZIP
          </button>
          <p style={{ color: 'var(--dex-gray-400)', fontSize: '0.8rem', marginTop: 12 }}>
            After downloading, extract the ZIP, open a terminal in the <code>dex-event-app</code> folder,
            and run <code>npm install && npm run dev</code>.
          </p>
        </div>
      </div>
    </div>
  );
}
