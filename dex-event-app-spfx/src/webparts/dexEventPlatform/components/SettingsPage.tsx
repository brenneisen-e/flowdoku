/**
 * Settings-Seite
 * Zeigt User-Infos, Admin-Aktionen und Download-Link.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { currentUser } from '../data/mockData';
import { Plus, FileText, Users, Download, Copy, Check } from './Icons';

const DOWNLOAD_URL = 'https://github.com/brenneisen-e/flowdoku/archive/refs/heads/main.zip';

export default function SettingsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const [copied, setCopied] = React.useState(false);

  // Link in die Zwischenablage kopieren (mit Fallback fuer aeltere Browser)
  const handleCopyLink = async (): Promise<void> => {
    try {
      await navigator.clipboard.writeText(DOWNLOAD_URL);
    } catch {
      const input = document.createElement('input');
      input.value = DOWNLOAD_URL;
      document.body.appendChild(input);
      input.select();
      document.execCommand('copy');
      document.body.removeChild(input);
    }
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="page-container">
      <div className="settings-grid">

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

        {currentUser.isAdmin && (
          <div className="card">
            <h3 className="mb-16">Admin Actions</h3>
            <div className="settings-actions">
              <button className="btn btn-primary btn-block" onClick={() => navigate('create-event')}>
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

        <div className="card">
          <h3 className="mb-16">Offline Version</h3>
          <p style={{ color: 'var(--dex-gray-600)', fontSize: '0.9rem', marginBottom: 16 }}>
            Copy the download link below and open it in a new browser tab to download the app as ZIP.
          </p>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, background: 'var(--dex-gray-100)', borderRadius: 'var(--dex-radius)', padding: '10px 12px', fontSize: '0.85rem', wordBreak: 'break-all' as any }}>
            <Download size={16} />
            <span style={{ flex: 1, color: 'var(--dex-gray-800)', userSelect: 'all' }}>{DOWNLOAD_URL}</span>
            <button className="btn btn-primary" onClick={handleCopyLink} style={{ flexShrink: 0, padding: '6px 12px', fontSize: '0.8rem' }}>
              {copied ? <span><Check size={14} /> Copied!</span> : <span><Copy size={14} /> Copy Link</span>}
            </button>
          </div>
          <p style={{ color: 'var(--dex-gray-400)', fontSize: '0.8rem', marginTop: 12 }}>
            After downloading, extract the ZIP, open a terminal in the <code>dex-event-app</code> folder,
            and run <code>npm install &amp;&amp; npm run dev</code>.
          </p>
        </div>
      </div>
    </div>
  );
}
