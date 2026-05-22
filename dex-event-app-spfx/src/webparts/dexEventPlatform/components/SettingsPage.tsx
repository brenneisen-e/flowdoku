/**
 * Settings-Seite
 * Zeigt User-Infos, Rollenmanagement (Admin), Admin-Aktionen und Download-Link.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { UserRole } from '../types';
import { Plus, FileText, Trash2, X } from './Icons';

export default function SettingsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { currentUser } = useCurrentUser();
  const {
    roles, currentUserRole, isAdmin, canCreateEvents,
    addRole, updateRole, updateRoleLocation, removeRole, isRolesLoading, siteUrl, searchUsers,
  } = useRoles();
  const { events, sendOrganizerOnboarding } = useEvents();

  /**
   * Map: organizer-email-lowercase -> Liste von Event-Titeln, die diese Person koordiniert.
   * Nutzt den gleichen Substring-Match wie die AdminPage Filter-Logik:
   * Organizer-Eintrag enthaelt entweder den Vor+Nachnamen oder nur den Nachnamen.
   */
  const organizerEventMap = React.useMemo<Record<string, string[]>>(() => {
    const map: Record<string, string[]> = {};
    for (const role of roles) {
      if (role.role !== 'Organizer' && role.role !== 'Admin') continue;
      const emailLc = (role.userEmail || '').toLowerCase();
      if (!emailLc) continue;
      const matched: string[] = [];
      for (const evt of events) {
        // v6.20: strikt per E-Mail — kein Namens-Substring-Match mehr.
        // Der alte Heuristik-Match per Nachname hat bei häufigen Nachnamen
        // False-Positives produziert (z.B. Assistentin mit gleichem Nachnamen
        // wie ein Organizer bekam Zugriff zugeordnet).
        if (evt.organizerEmails && evt.organizerEmails.some(e => e.toLowerCase() === emailLc)) {
          matched.push(evt.title);
        }
      }
      if (matched.length > 0) map[emailLc] = matched;
    }
    return map;
  }, [roles, events]);

  /**
   * v10.16: Per-Event-Co-Organizers — Personen die im Wizard-Picker eines Events
   * als Organizer eingetragen sind, aber KEINEN globalen Eintrag in DEX_Roles
   * haben. Sie haben für DAS jeweilige Event vollen Zugriff (über die Event-
   * Subsite-Permissions), tauchen aber nicht in der regulären roles-Tabelle auf.
   * Damit der Admin sieht „wer hat per Event Zugriff ohne globalen Status",
   * werden sie hier eingesammelt: pro Email die Liste der Events + ein
   * Display-Name (aus dem ersten Event-Match abgeleitet).
   */
  const coOrganizersList = React.useMemo<Array<{ email: string; name: string; events: string[] }>>(() => {
    const knownEmails = new Set(roles.map(r => (r.userEmail || '').toLowerCase()));
    const accumulator: Record<string, { name: string; events: string[] }> = {};
    for (const evt of events) {
      const orgEmails = evt.organizerEmails || [];
      const orgNames = evt.organizers || [];
      for (let i = 0; i < orgEmails.length; i++) {
        const emailLc = (orgEmails[i] || '').toLowerCase();
        if (!emailLc) continue;
        if (knownEmails.has(emailLc)) continue;  // schon in der Roles-Tabelle
        const name = orgNames[i] || emailLc;
        if (!accumulator[emailLc]) accumulator[emailLc] = { name, events: [] };
        accumulator[emailLc].events.push(evt.title);
      }
    }
    return Object.keys(accumulator).sort().map(emailLc => ({
      email: emailLc,
      name: accumulator[emailLc].name,
      events: accumulator[emailLc].events,
    }));
  }, [roles, events]);
  // Formular-State für neue Rolle
  const [newEmail, setNewEmail] = React.useState('');
  const [newName, setNewName] = React.useState('');
  const [newRole, setNewRole] = React.useState<UserRole>('Organizer');
  const [newLocation, setNewLocation] = React.useState('');
  const [isAdding, setIsAdding] = React.useState(false);
  const [showAddForm, setShowAddForm] = React.useState(false);
  const [statusMsg, setStatusMsg] = React.useState('');
  const [isSearching, setIsSearching] = React.useState(false);
  const [userFound, setUserFound] = React.useState<boolean | null>(null);
  const [suggestions, setSuggestions] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const [showSuggestions, setShowSuggestions] = React.useState(false);
  const searchTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  // Timer cleanup bei Unmount
  React.useEffect(() => {
    return () => { if (searchTimer.current) clearTimeout(searchTimer.current); };
  }, []);

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

  // Onboarding-Mail-Prompt: erscheint nach erfolgreicher Zuweisung einer
  // Organizer- oder Admin-Rolle. Ein Klick auf "Mail senden" verschickt eine
  // Begruessungsmail mit Links zur App und zum Handbuch (siehe
  // organizerOnboardingEmail in EmailTemplates.ts) — Cc geht automatisch an
  // die DEX-Verantwortlichen.
  const [onboardingPrompt, setOnboardingPrompt] = React.useState<
    { email: string; name: string; role: 'Organizer' | 'Admin' } | null
  >(null);
  const [isSendingOnboarding, setIsSendingOnboarding] = React.useState(false);

  const handleAddRole = async (): Promise<void> => {
    if (!newEmail || !newName) return;
    setIsAdding(true);
    setStatusMsg('');
    const assignedEmail = newEmail;
    const assignedName = newName;
    const assignedRole = newRole;
    const success = await addRole(assignedEmail, assignedName, assignedRole, newLocation);
    if (success) {
      setStatusMsg('Role assigned successfully.');
      setNewEmail('');
      setNewName('');
      setNewLocation('');
      setShowAddForm(false);
      // User/-Rolle bekommt keine Onboarding-Mail — die Mail erklaert
      // Organizer-/Admin-Funktionen, die Standard-User gar nicht haben.
      if (assignedRole === 'Organizer' || assignedRole === 'Admin') {
        setOnboardingPrompt({ email: assignedEmail, name: assignedName, role: assignedRole });
      }
    } else {
      setStatusMsg('Error: Could not assign role. Please try again.');
    }
    setIsAdding(false);
    setTimeout(() => setStatusMsg(''), 4000);
  };

  const handleSendOnboarding = async (): Promise<void> => {
    if (!onboardingPrompt) return;
    setIsSendingOnboarding(true);
    const ok = await sendOrganizerOnboarding(onboardingPrompt.email, onboardingPrompt.name, onboardingPrompt.role);
    setIsSendingOnboarding(false);
    setOnboardingPrompt(null);
    setStatusMsg(ok ? 'Onboarding-Mail wurde verschickt.' : 'Onboarding-Mail konnte nicht versendet werden.');
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
      'Admin': { bg: '#e8f5e9', color: '#2e7d32' },
      'Organizer': { bg: '#e3f2fd', color: '#1565c0' },
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

        {/* Admin Actions - sichtbar fuer Organizer und Admin */}
        {canCreateEvents && (
          <div className="card">
            <h3 className="mb-16">Admin Actions</h3>
            <div className="settings-actions">
              <button className="btn btn-primary btn-block" onClick={() => navigate('create-event')}>
                <Plus size={18} /> Create New Event
              </button>
              <button className="btn btn-secondary btn-block mt-8" onClick={() => navigate('admin')}>
                <FileText size={18} /> View All Events (Admin)
              </button>
              {isAdmin && (
                <button className="btn btn-secondary btn-block mt-8" onClick={() => navigate('role-matrix')}>
                  <FileText size={18} /> Rollen-Matrix anzeigen
                </button>
              )}
            </div>
          </div>
        )}

        {/* v9.21: Rollenmanagement collapsible — Admin kann die ganze Liste
            zusammenklappen wenn er sie gerade nicht braucht. Default: zu. */}
        {isAdmin && (
          <details className="card" style={{ cursor: 'default' }} open>
            <summary style={{
              cursor: 'pointer', fontWeight: 600, fontSize: '1rem',
              padding: '4px 0', listStyle: 'revert',
              color: 'var(--dex-gray-800)',
            }}>
              Role Management
            </summary>
            <div style={{ marginTop: 16 }}>
            <p style={{ color: 'var(--dex-gray-500, #888)', fontSize: '0.85rem', marginBottom: 16 }}>
              Manage who can create events. Roles are stored in the SharePoint list &ldquo;DEX_Roles&rdquo;.
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

            {/* Neue Rolle hinzufuegen — v11.72: nach OBEN verschoben, damit der
                Admin nicht erst durch die Tabelle scrollen muss. Form ist auf
                EINEN People-Picker reduziert (analog zum Organizer-Picker im
                EventCreation-Wizard). */}
            {!showAddForm ? (
              <button
                className="btn btn-primary"
                onClick={() => setShowAddForm(true)}
                style={{ fontSize: '0.85rem', marginBottom: 16 }}
              >
                <Plus size={16} /> Add User Role
              </button>
            ) : (
              <div style={{
                marginBottom: 16, padding: 16, background: 'var(--dex-gray-50, #fafafa)',
                borderRadius: 'var(--dex-radius, 8px)', border: '1px solid var(--dex-gray-200, #eee)',
              }}>
                <h4 style={{ margin: '0 0 12px', fontSize: '0.9rem' }}>Assign New Role</h4>

                {/* People-Picker — Such-Input ODER ausgewählter Chip */}
                <div style={{ marginBottom: 12 }}>
                  <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>
                    Person <span style={{ color: 'var(--dex-danger, red)' }}>*</span>
                  </label>
                  {userFound === true && newEmail && newName ? (
                    <div style={{
                      display: 'inline-flex', alignItems: 'center', gap: 10,
                      padding: '6px 10px 6px 6px', background: '#fff',
                      border: '1px solid var(--dex-green)', borderRadius: 999,
                      maxWidth: '100%',
                    }}>
                      <img
                        src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(newEmail)}&size=S`}
                        alt={newName}
                        onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                        style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                      />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 600, fontSize: '0.85rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {newName}
                        </div>
                        <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {newEmail}{newLocation ? ` · ${newLocation}` : ''}
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => { setNewEmail(''); setNewName(''); setNewLocation(''); setUserFound(null); }}
                        title="Auswahl entfernen"
                        style={{
                          border: 'none', background: 'var(--dex-gray-100)', cursor: 'pointer',
                          width: 24, height: 24, borderRadius: '50%',
                          display: 'flex', alignItems: 'center', justifyContent: 'center',
                          color: 'var(--dex-gray-600)', flexShrink: 0,
                        }}
                      >
                        <X size={16} />
                      </button>
                    </div>
                  ) : (
                    <div style={{ position: 'relative' }}>
                      <input
                        className="form-input"
                        value={newEmail}
                        onChange={e => handleEmailChange(e.target.value)}
                        onFocus={() => { if (suggestions.length > 0) setShowSuggestions(true); }}
                        onBlur={() => setTimeout(() => setShowSuggestions(false), 200)}
                        placeholder="Name oder Email eingeben..."
                        style={{ fontSize: '0.85rem' }}
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
                          boxShadow: 'var(--dex-shadow-hover)', maxHeight: 320, overflowY: 'auto', marginTop: 2,
                        }}>
                          {suggestions.map((s, i) => (
                            <div
                              key={i}
                              onMouseDown={() => selectSuggestion(s)}
                              style={{
                                display: 'flex', alignItems: 'center', gap: 10,
                                padding: '8px 12px', cursor: 'pointer',
                                borderBottom: i < suggestions.length - 1 ? '1px solid var(--dex-gray-100)' : 'none',
                              }}
                              onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = 'var(--dex-gray-100)'; }}
                              onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = '#fff'; }}
                            >
                              <img
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(s.email)}&size=S`}
                                alt={s.displayName}
                                onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                              />
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontWeight: 600, fontSize: '0.85rem' }}>{s.displayName}</div>
                                <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                  {s.email}{s.location ? ` · ${s.location}` : ''}
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>

                <div style={{ marginTop: 12 }}>
                  <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>Role</label>
                  <select
                    className="form-select"
                    value={newRole}
                    onChange={e => setNewRole(e.target.value as UserRole)}
                    style={{ fontSize: '0.85rem', maxWidth: 200 }}
                  >
                    <option value="Organizer">Organizer</option>
                    <option value="Admin">Admin</option>
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
                      <th style={{ textAlign: 'left', padding: 8, color: 'var(--dex-gray-500)' }}>Coordinated Events</th>
                      <th style={{ textAlign: 'right', padding: '8px 0 8px 8px', color: 'var(--dex-gray-500)' }} />
                    </tr>
                  </thead>
                  <tbody>
                    {[...roles]
                      .sort((a, b) => {
                        // Admins zuerst, dann Organizer, dann User - jeweils alphabetisch
                        const order: Record<string, number> = { Admin: 0, Organizer: 1, User: 2 };
                        const diff = (order[a.role] ?? 3) - (order[b.role] ?? 3);
                        return diff !== 0 ? diff : a.userName.localeCompare(b.userName);
                      })
                      .map((r, i, arr) => {
                        // Trennlinie zwischen Rollen-Gruppen
                        const prevRole = i > 0 ? arr[i - 1].role : null;
                        const showSeparator = prevRole && prevRole !== r.role;
                        return (
                          <React.Fragment key={r.id}>
                            {showSeparator && (
                              <tr><td colSpan={6} style={{ padding: 0 }}><hr style={{ border: 'none', borderTop: '2px solid var(--dex-gray-300)', margin: '4px 0' }} /></td></tr>
                            )}
                            <tr style={{ borderBottom: '1px solid var(--dex-gray-100, #f0f0f0)' }}>
                        <td style={{ padding: '10px 8px 10px 0', fontWeight: 500 }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                            <img
                              src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(r.userEmail)}&size=S`}
                              alt={r.userName}
                              onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                              style={{
                                width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', flexShrink: 0,
                                background: 'var(--dex-gray-100)',
                                transition: 'transform 0.15s ease, box-shadow 0.15s ease',
                                transformOrigin: 'left center',
                                cursor: 'zoom-in',
                              }}
                              onMouseEnter={e => {
                                (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.4)';
                                (e.currentTarget as HTMLImageElement).style.boxShadow = '0 4px 14px rgba(0,0,0,0.18)';
                                (e.currentTarget as HTMLImageElement).style.zIndex = '50';
                                (e.currentTarget as HTMLImageElement).style.position = 'relative';
                              }}
                              onMouseLeave={e => {
                                (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)';
                                (e.currentTarget as HTMLImageElement).style.boxShadow = 'none';
                                (e.currentTarget as HTMLImageElement).style.zIndex = '';
                                (e.currentTarget as HTMLImageElement).style.position = '';
                              }}
                            />
                            <span>{r.userName}</span>
                          </div>
                        </td>
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
                            <option value="Admin">Admin</option>
                            <option value="Organizer">Organizer</option>
                            <option value="User">User</option>
                          </select>
                        </td>
                        <td style={{ padding: 10 }}>
                          <input
                            className="form-input"
                            key={`loc-${r.id}-${r.location}`}
                            defaultValue={r.location || ''}
                            style={{ fontSize: '0.85rem', padding: '4px 8px', width: '100%', minWidth: 120 }}
                            placeholder="Standort"
                            onBlur={async (e) => {
                              const newLoc = e.target.value.trim();
                              if (newLoc !== (r.location || '')) {
                                await updateRoleLocation(r.id, newLoc);
                              }
                            }}
                          />
                        </td>
                        <td style={{ padding: 10, fontSize: '0.78rem', color: 'var(--dex-gray-600)', maxWidth: 280 }}>
                          {(() => {
                            const evts = organizerEventMap[r.userEmail.toLowerCase()] || [];
                            if (r.role === 'User') return <span style={{ color: 'var(--dex-gray-300)' }}>—</span>;
                            if (evts.length === 0) return <span style={{ color: 'var(--dex-gray-300)' }}>keine</span>;
                            return (
                              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                                {evts.map((title, idx) => (
                                  <span key={idx} style={{
                                    display: 'inline-block', padding: '2px 8px', borderRadius: 10,
                                    background: 'rgba(134,188,37,0.14)', color: 'var(--dex-green-dark)',
                                    fontSize: '0.74rem', fontWeight: 600,
                                  }}>{title}</span>
                                ))}
                              </div>
                            );
                          })()}
                        </td>
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
                          </React.Fragment>
                        );
                      })}
                  </tbody>
                </table>

                {/* v10.16: Per-Event-Co-Organizer — Personen mit Event-Zugriff
                    aber ohne globalen Eintrag in DEX_Roles. Sie wurden vom
                    Hauptorganizer per Wizard-Picker zu einem (oder mehreren)
                    Events hinzugefügt und haben für DIESE Events Vollzugriff,
                    ohne dafür „Organizer"-Rolle global haben zu müssen. */}
                {coOrganizersList.length > 0 && (
                  <div style={{ marginTop: 24 }}>
                    <h4 style={{ margin: '0 0 4px', fontSize: '0.95rem', color: 'var(--dex-gray-800)' }}>
                      Per-Event-Co-Organizer ({coOrganizersList.length})
                    </h4>
                    <p style={{ margin: '0 0 12px', fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                      Diese Personen sind im Wizard-Picker eines oder mehrerer Events als Organizer eingetragen, haben aber KEINEN globalen Eintrag in DEX_Roles. Sie können das jeweilige Event verwalten (Teilnehmer, Bearbeiten, Mails) — aber keine NEUEN Events anlegen. Wenn du jemandem permanent &bdquo;Organizer&ldquo;-Status geben willst, fügst du sie über das Formular oben mit Role &bdquo;Organizer&ldquo; hinzu.
                    </p>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--dex-gray-200, #eee)' }}>
                          <th style={{ textAlign: 'left', padding: '8px 8px 8px 0', color: 'var(--dex-gray-500)' }}>Name</th>
                          <th style={{ textAlign: 'left', padding: 8, color: 'var(--dex-gray-500)' }}>Email</th>
                          <th style={{ textAlign: 'left', padding: 8, color: 'var(--dex-gray-500)' }}>Status</th>
                          <th style={{ textAlign: 'left', padding: 8, color: 'var(--dex-gray-500)' }}>Coordinated Events</th>
                        </tr>
                      </thead>
                      <tbody>
                        {coOrganizersList.map(co => (
                          <tr key={co.email} style={{ borderBottom: '1px solid var(--dex-gray-100, #f0f0f0)' }}>
                            <td style={{ padding: '10px 8px 10px 0', fontWeight: 500 }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                                <img
                                  src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(co.email)}&size=S`}
                                  alt={co.name}
                                  onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                  style={{
                                    width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', flexShrink: 0,
                                    background: 'var(--dex-gray-100)',
                                    transition: 'transform 0.15s ease, box-shadow 0.15s ease',
                                    transformOrigin: 'left center',
                                    cursor: 'zoom-in',
                                  }}
                                  onMouseEnter={e => {
                                    (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.4)';
                                    (e.currentTarget as HTMLImageElement).style.boxShadow = '0 4px 14px rgba(0,0,0,0.18)';
                                    (e.currentTarget as HTMLImageElement).style.zIndex = '50';
                                    (e.currentTarget as HTMLImageElement).style.position = 'relative';
                                  }}
                                  onMouseLeave={e => {
                                    (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)';
                                    (e.currentTarget as HTMLImageElement).style.boxShadow = 'none';
                                    (e.currentTarget as HTMLImageElement).style.zIndex = '';
                                    (e.currentTarget as HTMLImageElement).style.position = '';
                                  }}
                                />
                                <span>{co.name}</span>
                              </div>
                            </td>
                            <td style={{ padding: 10, color: 'var(--dex-gray-600)' }}>{co.email}</td>
                            <td style={{ padding: 10 }}>
                              <span style={{
                                display: 'inline-block', padding: '3px 10px', borderRadius: 999,
                                background: 'rgba(237,139,0,0.15)', color: 'var(--dex-orange-dark, #b35a00)',
                                fontSize: '0.74rem', fontWeight: 600,
                              }}>Per-Event</span>
                            </td>
                            <td style={{ padding: 10, fontSize: '0.78rem', color: 'var(--dex-gray-600)', maxWidth: 280 }}>
                              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                                {co.events.map((title, idx) => (
                                  <span key={idx} style={{
                                    display: 'inline-block', padding: '2px 8px', borderRadius: 10,
                                    background: 'rgba(134,188,37,0.14)', color: 'var(--dex-green-dark)',
                                    fontSize: '0.74rem', fontWeight: 600,
                                  }}>{title}</span>
                                ))}
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            </div>
          </details>
        )}

        {/* v9.16/v9.21: Test-Team war hier global, ist jetzt per-Event
            (im EventCreation-Wizard). TestTeamManager entfernt. */}

        {/* Berechtigungs-Übersicht - nur fuer Admin */}
        {isAdmin && <PermissionsViewer siteUrl={siteUrl} />}

      </div>

      {/* Onboarding-Mail-Prompt nach erfolgreicher Rollen-Zuweisung */}
      {onboardingPrompt && (
        <div
          role="dialog"
          aria-modal="true"
          style={{
            position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
            background: 'rgba(0,0,0,0.45)', zIndex: 2000,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
          }}
          onClick={() => { if (!isSendingOnboarding) setOnboardingPrompt(null); }}
        >
          <div
            onClick={e => e.stopPropagation()}
            style={{
              background: '#fff', borderRadius: 'var(--dex-radius, 8px)',
              maxWidth: 460, width: '100%', padding: 24,
              boxShadow: 'var(--dex-shadow-hover)',
            }}
          >
            <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem' }}>
              Onboarding-Mail an {onboardingPrompt.name}?
            </h3>
            <p style={{ margin: '0 0 12px', fontSize: '0.9rem', lineHeight: 1.5, color: 'var(--dex-gray-700)' }}>
              <strong>{onboardingPrompt.name}</strong> wurde als <strong>{onboardingPrompt.role}</strong> hinzugefügt.
              Möchtest du eine Begrüßungsmail mit Link zur App, zum Handbuch und einer kurzen
              Anleitung zum ersten Test-Event verschicken?
            </p>
            <p style={{ margin: '0 0 20px', fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
              Empfänger: {onboardingPrompt.email}<br />
              Cc: ebrenneisen@deloitte.de, nifelten@deloitte.de
            </p>
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setOnboardingPrompt(null)}
                disabled={isSendingOnboarding}
                style={{ fontSize: '0.85rem' }}
              >
                Nicht senden
              </button>
              <button
                className="btn btn-primary"
                onClick={handleSendOnboarding}
                disabled={isSendingOnboarding}
                style={{ fontSize: '0.85rem' }}
              >
                {isSendingOnboarding ? 'Sende...' : 'Mail senden'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/**
 * Zeigt die Berechtigungen aller DEX-Listen an.
 */
function PermissionsViewer(props: { siteUrl: string }): React.ReactElement {
  const { siteUrl } = props;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ctx = (window as any).__dexSpfxContext;
  const SPHttpClient = require('@microsoft/sp-http').SPHttpClient;

  type ListPerms = { listName: string; perms: Array<{ name: string; type: string; level: string }>; loading: boolean; error?: string };
  const [lists, setLists] = React.useState<ListPerms[]>([]);
  const [isOpen, setIsOpen] = React.useState(false);

  const listNames = ['DEX_Events', 'DEX_Roles', 'DEX_Emails', 'DEX_Outlook', 'DEX_IDReorder', 'DEX_Participants'];

  async function loadPermissions(): Promise<void> {
    if (!ctx) return;
    const results: ListPerms[] = [];

    for (const listName of listNames) {
      try {
        const resp = await ctx.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments?$expand=Member,RoleDefinitionBindings&$select=Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name`,
          SPHttpClient.configurations.v1
        );
        if (!resp.ok) {
          results.push({ listName, perms: [], loading: false, error: `${resp.status}` });
          continue;
        }
        const data = await resp.json();
        const items = data.value || data.d?.results || [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const perms = items.map((item: any) => {
          const member = item.Member || {};
          const bindings = item.RoleDefinitionBindings || item.RoleDefinitionBindings?.results || [];
          const roleNames = (Array.isArray(bindings) ? bindings : bindings.results || [])
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            .map((b: any) => b.Name).filter((n: string) => n !== 'Limited Access');
          const pType = member.PrincipalType === 8 ? 'Gruppe' : 'User';
          return { name: member.Title || '?', type: pType, level: roleNames.join(', ') || '-' };
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        }).filter((p: any) => p.level !== '-');
        results.push({ listName, perms, loading: false });
      } catch {
        results.push({ listName, perms: [], loading: false, error: 'Fehler' });
      }
    }
    setLists(results);
  }

  return (
    <div className="card" style={{ marginTop: 16 }}>
      <h3
        className="mb-16"
        style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 8 }}
        onClick={() => { setIsOpen(!isOpen); if (!isOpen && lists.length === 0) loadPermissions(); }}
      >
        🔒 Listen-Berechtigungen {isOpen ? '▾' : '▸'}
      </h3>
      {isOpen && (
        <div>
          {lists.length === 0 && <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>Lade Berechtigungen...</p>}
          {lists.map(list => (
            <div key={list.listName} style={{ marginBottom: 16 }}>
              <h4 style={{ fontSize: '0.85rem', marginBottom: 4 }}>{list.listName}</h4>
              {list.error ? (
                <p style={{ color: 'var(--dex-danger, red)', fontSize: '0.8rem' }}>Fehler: {list.error}</p>
              ) : (
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.8rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '1px solid var(--dex-gray-200)' }}>
                      <th style={{ textAlign: 'left', padding: 4, color: 'var(--dex-gray-500)' }}>Name</th>
                      <th style={{ textAlign: 'left', padding: 4, color: 'var(--dex-gray-500)' }}>Typ</th>
                      <th style={{ textAlign: 'left', padding: 4, color: 'var(--dex-gray-500)' }}>Berechtigung</th>
                    </tr>
                  </thead>
                  <tbody>
                    {list.perms.map((p, i) => (
                      <tr key={i} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                        <td style={{ padding: 4 }}>{p.name}</td>
                        <td style={{ padding: 4, color: 'var(--dex-gray-500)' }}>{p.type}</td>
                        <td style={{ padding: 4 }}>
                          <span style={{
                            padding: '2px 6px', borderRadius: 4, fontSize: '0.75rem',
                            background: p.level.includes('Full Control') ? '#e8f5e9' : p.level.includes('Contribute') ? '#fff3e0' : '#e3f2fd',
                            color: p.level.includes('Full Control') ? '#2e7d32' : p.level.includes('Contribute') ? '#e65100' : '#1565c0',
                          }}>
                            {p.level}
                          </span>
                        </td>
                      </tr>
                    ))}
                    {list.perms.length === 0 && !list.error && (
                      <tr><td colSpan={3} style={{ padding: 4, color: 'var(--dex-gray-400)' }}>Erbt Parent-Berechtigungen</td></tr>
                    )}
                  </tbody>
                </table>
              )}
            </div>
          ))}
          <button
            className="btn btn-secondary mt-8"
            style={{ fontSize: '0.8rem' }}
            onClick={loadPermissions}
          >
            Aktualisieren
          </button>
        </div>
      )}
    </div>
  );
}


