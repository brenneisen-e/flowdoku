/**
 * Settings-Seite
 * Zeigt User-Infos, Rollenmanagement (SuperAdmin), Admin-Aktionen und Download-Link.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { UserRole } from '../types';
import { Plus, FileText, Users, Download, Copy, Check, Trash2 } from './Icons';

const DOWNLOAD_URL = 'https://github.com/brenneisen-e/flowdoku/archive/refs/heads/main.zip';

export default function SettingsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { currentUser } = useCurrentUser();
  const {
    roles, currentUserRole, isSuperAdmin, canCreateEvents,
    addRole, updateRole, removeRole, isRolesLoading, siteUrl,
  } = useRoles();
  const [copied, setCopied] = React.useState(false);

  // Formular-State fuer neue Rolle
  const [newEmail, setNewEmail] = React.useState('');
  const [newName, setNewName] = React.useState('');
  const [newRole, setNewRole] = React.useState<UserRole>('EventAdmin');
  const [newLocation, setNewLocation] = React.useState('');
  const [isAdding, setIsAdding] = React.useState(false);
  const [showAddForm, setShowAddForm] = React.useState(false);
  const [statusMsg, setStatusMsg] = React.useState('');

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

  const handleAddRole = async (): Promise<void> => {
    if (!newEmail || !newName) return;
    setIsAdding(true);
    setStatusMsg('');
    const success = await addRole(newEmail, newName, newRole, newLocation);
    if (success) {
      setStatusMsg('Role assigned successfully.');
      setNewEmail('');
      setNewName('');
      setNewLocation('');
      setShowAddForm(false);
    } else {
      setStatusMsg('Error: Could not assign role. Please try again.');
    }
    setIsAdding(false);
    setTimeout(() => setStatusMsg(''), 4000);
  };

  const handleRemoveRole = async (itemId: number, userName: string): Promise<void> => {
    const confirmed = confirm(`Remove role for "${userName}"?`);
    if (!confirmed) return;
    await removeRole(itemId);
  };

  const handleChangeRole = async (itemId: number, role: UserRole): Promise<void> => {
    await updateRole(itemId, role);
  };

  // Rollen-Badge Farbe
  const roleBadge = (role: string): React.ReactElement => {
    const colors: Record<string, { bg: string; color: string }> = {
      'SuperAdmin': { bg: '#e8f5e9', color: '#2e7d32' },
      'EventAdmin': { bg: '#e3f2fd', color: '#1565c0' },
      'User': { bg: '#f5f5f5', color: '#666' },
    };
    const c = colors[role] || colors['User'];
    return (
      <span style={{
        display: 'inline-block', padding: '2px 10px', borderRadius: 12,
        background: c.bg, color: c.color, fontSize: '0.8rem', fontWeight: 500,
      }}>
        {role}
      </span>
    );
  };

  return (
    <div className="page-container">
      <div className="settings-grid">

        {/* User Information */}
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
              <span>{currentUser.location || '-'}</span>
            </div>
            <div className="settings-info__row">
              <span className="settings-info__label">Role</span>
              {roleBadge(currentUserRole)}
            </div>
          </div>
        </div>

        {/* Admin Actions - sichtbar fuer EventAdmin und SuperAdmin */}
        {canCreateEvents && (
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

        {/* Rollenmanagement - nur fuer SuperAdmin */}
        {isSuperAdmin && (
          <div className="card">
            <h3 className="mb-16">Role Management</h3>
            <p style={{ color: 'var(--dex-gray-500, #888)', fontSize: '0.85rem', marginBottom: 16 }}>
              Manage who can create events. Roles are stored in the SharePoint list "DEX_Roles".
              {' '}
              <a
                href={`${siteUrl}/Lists/DEX_Roles/AllItems.aspx`}
                target="_blank"
                rel="noopener noreferrer"
                style={{ color: 'var(--dex-green-dark, #6b9a1e)', fontWeight: 600, textDecoration: 'underline' }}
              >
                Open SharePoint List
              </a>
            </p>

            {statusMsg && (
              <div style={{
                padding: '8px 12px', borderRadius: 8, marginBottom: 16, fontSize: '0.85rem',
                background: statusMsg.startsWith('Error') ? '#fce4ec' : '#e8f5e9',
                color: statusMsg.startsWith('Error') ? '#c62828' : '#2e7d32',
              }}>
                {statusMsg}
              </div>
            )}

            {/* Rollen-Tabelle */}
            {isRolesLoading ? (
              <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>Loading roles...</p>
            ) : (
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--dex-gray-200, #eee)' }}>
                      <th style={{ textAlign: 'left', padding: '8px 8px 8px 0', color: 'var(--dex-gray-500)' }}>Name</th>
                      <th style={{ textAlign: 'left', padding: 8, color: 'var(--dex-gray-500)' }}>Email</th>
                      <th style={{ textAlign: 'left', padding: 8, color: 'var(--dex-gray-500)' }}>Role</th>
                      <th style={{ textAlign: 'left', padding: 8, color: 'var(--dex-gray-500)' }}>Location</th>
                      <th style={{ textAlign: 'right', padding: '8px 0 8px 8px', color: 'var(--dex-gray-500)' }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {roles.map(r => (
                      <tr key={r.id} style={{ borderBottom: '1px solid var(--dex-gray-100, #f0f0f0)' }}>
                        <td style={{ padding: '10px 8px 10px 0', fontWeight: 500 }}>{r.userName}</td>
                        <td style={{ padding: 10, color: 'var(--dex-gray-600)' }}>{r.userEmail}</td>
                        <td style={{ padding: 10 }}>
                          <select
                            value={r.role}
                            onChange={e => handleChangeRole(r.id, e.target.value as UserRole)}
                            style={{
                              padding: '4px 8px', borderRadius: 6, border: '1px solid var(--dex-gray-200, #ddd)',
                              fontSize: '0.8rem', background: '#fff',
                            }}
                            disabled={r.userEmail.toLowerCase() === currentUser.email.toLowerCase()}
                          >
                            <option value="SuperAdmin">SuperAdmin</option>
                            <option value="EventAdmin">EventAdmin</option>
                            <option value="User">User</option>
                          </select>
                        </td>
                        <td style={{ padding: 10, color: 'var(--dex-gray-500)' }}>{r.location || '-'}</td>
                        <td style={{ padding: '10px 0 10px 8px', textAlign: 'right' }}>
                          {r.userEmail.toLowerCase() !== currentUser.email.toLowerCase() && (
                            <button
                              onClick={() => handleRemoveRole(r.id, r.userName)}
                              style={{
                                border: 'none', background: 'none', cursor: 'pointer',
                                color: 'var(--dex-danger, #e53935)', padding: 4,
                              }}
                              title="Remove role"
                            >
                              <Trash2 size={16} />
                            </button>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {/* Neue Rolle hinzufuegen */}
            {!showAddForm ? (
              <button
                className="btn btn-primary mt-16"
                onClick={() => setShowAddForm(true)}
                style={{ fontSize: '0.85rem' }}
              >
                <Plus size={16} /> Add User Role
              </button>
            ) : (
              <div style={{
                marginTop: 16, padding: 16, background: 'var(--dex-gray-50, #fafafa)',
                borderRadius: 'var(--dex-radius, 8px)', border: '1px solid var(--dex-gray-200, #eee)',
              }}>
                <h4 style={{ margin: '0 0 12px', fontSize: '0.9rem' }}>Assign New Role</h4>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>
                      Name <span style={{ color: 'var(--dex-danger, red)' }}>*</span>
                    </label>
                    <input
                      className="form-input"
                      value={newName}
                      onChange={e => setNewName(e.target.value)}
                      placeholder="Max Mustermann"
                      style={{ fontSize: '0.85rem' }}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>
                      Email <span style={{ color: 'var(--dex-danger, red)' }}>*</span>
                    </label>
                    <input
                      className="form-input"
                      type="email"
                      value={newEmail}
                      onChange={e => setNewEmail(e.target.value)}
                      placeholder="mmustermann@deloitte.de"
                      style={{ fontSize: '0.85rem' }}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>Role</label>
                    <select
                      className="form-select"
                      value={newRole}
                      onChange={e => setNewRole(e.target.value as UserRole)}
                      style={{ fontSize: '0.85rem' }}
                    >
                      <option value="EventAdmin">EventAdmin</option>
                      <option value="SuperAdmin">SuperAdmin</option>
                    </select>
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>Location</label>
                    <input
                      className="form-input"
                      value={newLocation}
                      onChange={e => setNewLocation(e.target.value)}
                      placeholder="Duesseldorf"
                      style={{ fontSize: '0.85rem' }}
                    />
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 8, marginTop: 16, justifyContent: 'flex-end' }}>
                  <button
                    className="btn btn-secondary"
                    onClick={() => { setShowAddForm(false); setNewEmail(''); setNewName(''); setNewLocation(''); }}
                    style={{ fontSize: '0.85rem' }}
                  >
                    Cancel
                  </button>
                  <button
                    className="btn btn-primary"
                    onClick={handleAddRole}
                    disabled={isAdding || !newEmail || !newName}
                    style={{ fontSize: '0.85rem' }}
                  >
                    {isAdding ? 'Saving...' : 'Assign Role'}
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Offline Version */}
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
