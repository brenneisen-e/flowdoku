/**
 * Settings-Seite
 * Zeigt User-Infos, Rollenmanagement (SuperAdmin), Admin-Aktionen und Download-Link.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { UserRole } from '../types';
import { Plus, FileText, Users, Trash2 } from './Icons';

export default function SettingsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { currentUser } = useCurrentUser();
  const {
    roles, currentUserRole, isSuperAdmin, canCreateEvents,
    addRole, updateRole, removeRole, isRolesLoading, siteUrl, searchUsers,
  } = useRoles();
  // Formular-State für neue Rolle
  const [newEmail, setNewEmail] = React.useState('');
  const [newName, setNewName] = React.useState('');
  const [newRole, setNewRole] = React.useState<UserRole>('EventAdmin');
  const [newLocation, setNewLocation] = React.useState('');
  const [isAdding, setIsAdding] = React.useState(false);
  const [showAddForm, setShowAddForm] = React.useState(false);
  const [statusMsg, setStatusMsg] = React.useState('');
  const [isSearching, setIsSearching] = React.useState(false);
  const [userFound, setUserFound] = React.useState<boolean | null>(null);
  const [suggestions, setSuggestions] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const [showSuggestions, setShowSuggestions] = React.useState(false);
  const searchTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  // Email-Eingabe: User-Autocomplete mit Debounce
  const handleEmailChange = (query: string): void => {
    setNewEmail(query);
    setUserFound(null);

    if (searchTimer.current) clearTimeout(searchTimer.current);

    if (query.length >= 2) {
      searchTimer.current = setTimeout(async () => {
        setIsSearching(true);
        const results = await searchUsers(query);
        setSuggestions(results);
        setShowSuggestions(results.length > 0);
        setIsSearching(false);
      }, 400);
    } else {
      setSuggestions([]);
      setShowSuggestions(false);
    }
  };

  const selectSuggestion = (suggestion: { email: string; displayName: string; location: string }): void => {
    setNewEmail(suggestion.email);
    setNewName(suggestion.displayName);
    setNewLocation(suggestion.location);
    setUserFound(true);
    setShowSuggestions(false);
    setSuggestions([]);
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

  const [isRemoving, setIsRemoving] = React.useState<number | null>(null);

  const handleRemoveRole = async (itemId: number, userName: string): Promise<void> => {
    const confirmed = confirm(`Rolle für "${userName}" entfernen?`);
    if (!confirmed) return;
    setIsRemoving(itemId);
    const success = await removeRole(itemId);
    setIsRemoving(null);
    if (success) {
      setStatusMsg(`Rolle für ${userName} entfernt.`);
    } else {
      setStatusMsg('Error: Rolle konnte nicht entfernt werden.');
    }
    setTimeout(() => setStatusMsg(''), 4000);
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
                              disabled={isRemoving === r.id}
                              style={{
                                border: 'none', background: 'none', cursor: isRemoving === r.id ? 'wait' : 'pointer',
                                color: 'var(--dex-danger, #e53935)', padding: 4,
                                opacity: isRemoving === r.id ? 0.4 : 1,
                              }}
                              title="Rolle entfernen"
                            >
                              {isRemoving === r.id ? '...' : <Trash2 size={16} />}
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
                {/* Email mit Autocomplete */}
                <div style={{ marginBottom: 12 }}>
                  <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>
                    Email <span style={{ color: 'var(--dex-danger, red)' }}>*</span>
                  </label>
                  <div style={{ position: 'relative' }}>
                    <input
                      className="form-input"
                      value={newEmail}
                      onChange={e => handleEmailChange(e.target.value)}
                      onFocus={() => { if (suggestions.length > 0) setShowSuggestions(true); }}
                      onBlur={() => setTimeout(() => setShowSuggestions(false), 200)}
                      placeholder="Name oder Email eingeben..."
                      style={{
                        fontSize: '0.85rem',
                        borderColor: userFound === true ? 'var(--dex-green)' : undefined,
                      }}
                      autoFocus
                      autoComplete="off"
                    />
                    {isSearching && (
                      <span style={{ position: 'absolute', right: 12, top: '50%', transform: 'translateY(-50%)', fontSize: '0.8rem', color: 'var(--dex-gray-400)' }}>
                        Suche...
                      </span>
                    )}
                    {showSuggestions && suggestions.length > 0 && (
                      <div style={{
                        position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 200,
                        background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 'var(--dex-radius)',
                        boxShadow: 'var(--dex-shadow-hover)', maxHeight: 240, overflowY: 'auto',
                      }}>
                        {suggestions.map((s, i) => (
                          <div
                            key={i}
                            onMouseDown={() => selectSuggestion(s)}
                            style={{
                              padding: '10px 12px', cursor: 'pointer', borderBottom: '1px solid var(--dex-gray-100)',
                              fontSize: '0.85rem',
                            }}
                            onMouseEnter={e => { (e.target as HTMLDivElement).style.background = 'var(--dex-gray-100)'; }}
                            onMouseLeave={e => { (e.target as HTMLDivElement).style.background = '#fff'; }}
                          >
                            <div style={{ fontWeight: 600 }}>{s.displayName}</div>
                            <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>
                              {s.email}{s.location ? ` · ${s.location}` : ''}
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                  {userFound === true && (
                    <span style={{ fontSize: '0.75rem', color: 'var(--dex-green-dark)', marginTop: 4, display: 'block' }}>
                      Gefunden: {newName}{newLocation ? ` (${newLocation})` : ''}
                    </span>
                  )}
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>
                      Name {userFound !== true && <span style={{ color: 'var(--dex-danger, red)' }}>*</span>}
                    </label>
                    <input
                      className="form-input"
                      value={newName}
                      onChange={e => setNewName(e.target.value)}
                      placeholder={userFound === true ? '' : 'Max Mustermann'}
                      style={{
                        fontSize: '0.85rem',
                        background: userFound === true ? 'var(--dex-gray-100)' : undefined,
                      }}
                      readOnly={userFound === true}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>Location</label>
                    <input
                      className="form-input"
                      value={newLocation}
                      onChange={e => setNewLocation(e.target.value)}
                      placeholder={userFound === true ? '' : 'Düsseldorf'}
                      style={{
                        fontSize: '0.85rem',
                        background: userFound === true ? 'var(--dex-gray-100)' : undefined,
                      }}
                      readOnly={userFound === true}
                    />
                  </div>
                </div>

                <div style={{ marginTop: 12 }}>
                  <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>Role</label>
                  <select
                    className="form-select"
                    value={newRole}
                    onChange={e => setNewRole(e.target.value as UserRole)}
                    style={{ fontSize: '0.85rem', maxWidth: 200 }}
                  >
                    <option value="EventAdmin">EventAdmin</option>
                    <option value="SuperAdmin">SuperAdmin</option>
                  </select>
                </div>

                <div style={{ display: 'flex', gap: 8, marginTop: 16, justifyContent: 'flex-end' }}>
                  <button
                    className="btn btn-secondary"
                    onClick={() => { setShowAddForm(false); setNewEmail(''); setNewName(''); setNewLocation(''); setUserFound(null); }}
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

      </div>
    </div>
  );
}
