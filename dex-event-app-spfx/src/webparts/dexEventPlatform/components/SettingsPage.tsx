/**
 * Settings-Seite
 * Zeigt User-Infos, Rollenmanagement (Admin), Admin-Aktionen und Download-Link.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useLanguage } from '../context/LanguageContext';
// v20.4: moderne Confirm-Modals statt window.confirm.
import { useDialog } from '../context/DialogContext';
import { UserRole } from '../types';
import { Plus, Trash2, X } from './Icons';
import Modal from './Modal';
import InternationalSearchToggle from './InternationalSearchToggle';
import { PersonContactHover } from './PersonContactHover';

export default function SettingsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { currentUser } = useCurrentUser();
  const {
    roles, isAdmin, originalIsAdmin,
    addRole, updateRole, setPowerUser, removeRole, isRolesLoading, siteUrl, searchUsers, searchUser,
  } = useRoles();
  const { events, sendOrganizerOnboarding } = useEvents();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  // v20.4: App-Modal statt window.confirm.
  const { confirmDialog } = useDialog();
  // v13.0: Settings/Rollenverwaltung ist Admin-only. Vorher konnte ein
  // Demo-User die Seite öffnen — Admin-Controls waren zwar versteckt,
  // aber der Seitenzugriff selbst war frei. Wir nutzen originalIsAdmin
  // damit der Admin-im-Demo-Modus seine eigene Einstellungen weiterhin
  // testen kann.
  React.useEffect(() => {
    if (!originalIsAdmin) navigate('start');
  }, [originalIsAdmin, navigate]);

  /**
   * Map: organizer-email-lowercase -> Liste von Event-Titeln, die diese Person koordiniert.
   * Nutzt den gleichen Substring-Match wie die AdminPage Filter-Logik:
   * Organizer-Eintrag enthält entweder den Vor+Nachnamen oder nur den Nachnamen.
   */
  const organizerEventMap = React.useMemo<Record<string, string[]>>(() => {
    const map: Record<string, string[]> = {};
    for (const role of roles) {
      if (role.role !== 'Organizer' && role.role !== 'Admin' && role.role !== 'IT-Admin') continue;
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
    // v24.84: alphabetisch nach Name sortieren (vorher nach E-Mail-Schlüssel).
    return Object.keys(accumulator).map(emailLc => ({
      email: emailLc,
      name: accumulator[emailLc].name,
      events: accumulator[emailLc].events,
    })).sort((a, b) => (a.name || '').localeCompare(b.name || '', 'de'));
  }, [roles, events]);
  // v19.26: Die Per-Event-Organizer-Namen stammen aus dem index-basierten
  // Pairing von Organizer-Namen ↔ -E-Mails im Event. Driftet dieses Pairing
  // (z.B. weil ein Feld einen leeren Slot hatte und unabhängig gefiltert wurde),
  // steht der Name neben der falschen E-Mail — und das Foto (lädt per E-Mail)
  // zeigt die falsche Person. Wir lösen den Anzeige-Namen daher zuverlässig über
  // die E-Mail auf (gleicher Schlüssel wie das Foto); Fallback bleibt der
  // Index-Name. So passen Name + Foto immer zusammen.
  const [resolvedOrgNames, setResolvedOrgNames] = React.useState<Record<string, string>>({});
  const orgNameAttemptedRef = React.useRef<Set<string>>(new Set());
  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      for (const co of coOrganizersList) {
        if (orgNameAttemptedRef.current.has(co.email)) continue;
        orgNameAttemptedRef.current.add(co.email);
        try {
          // includeInternational=true → deckt @deloitte.de UND @deloitte.com ab.
          const results = await searchUsers(co.email, true);
          const match = results.find(r => (r.email || '').toLowerCase() === co.email);
          if (match && match.displayName && !cancelled) {
            setResolvedOrgNames(prev => ({ ...prev, [co.email]: match.displayName }));
          }
        } catch { /* Fallback: Index-Name aus co.name */ }
      }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [coOrganizersList]);
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
  const [includeIntl, setIncludeIntl] = React.useState(false);
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
        const results = await searchUsers(query, includeIntl);
        setSuggestions(results);
        setShowSuggestions(results.length > 0);
        setIsSearching(false);
      }, 400);
    } else {
      setSuggestions([]);
      setShowSuggestions(false);
    }
  };

  React.useEffect(() => {
    if (newEmail.length >= 2 && userFound !== true) {
      (async () => {
        setIsSearching(true);
        try {
          const results = await searchUsers(newEmail, includeIntl);
          setSuggestions(results);
          setShowSuggestions(results.length > 0);
        } catch {
          setSuggestions([]);
        }
        setIsSearching(false);
      })().catch(() => {});
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [includeIntl]);

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
  // Begrüßungsmail mit Links zur App und zum Handbuch (siehe
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
    // v24.84: Duplikaterkennung — hat die Person schon eine Rolle, nicht erneut
    // hinzufügen (sonst doppelte Zeile in DEX_Roles). Stattdessen Hinweis, die
    // bestehende Rolle in der Liste unten zu ändern.
    const existing = roles.find(r => (r.userEmail || '').toLowerCase() === assignedEmail.toLowerCase());
    if (existing) {
      setStatusMsg(`Error: ${assignedName} hat bereits eine Rolle (${existing.role}). Bitte ändere die bestehende Rolle in der Liste unten, statt die Person erneut hinzuzufügen.`);
      setIsAdding(false);
      setTimeout(() => setStatusMsg(''), 6000);
      return;
    }
    const success = await addRole(assignedEmail, assignedName, assignedRole, newLocation);
    if (success) {
      setStatusMsg('Role assigned successfully.');
      setNewEmail('');
      setNewName('');
      setNewLocation('');
      setShowAddForm(false);
      // User/-Rolle bekommt keine Onboarding-Mail — die Mail erklärt
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
    const confirmed = await confirmDialog(`Rolle für "${userName}" entfernen?`, { danger: true, confirmLabel: isDe ? 'Entfernen' : 'Remove' });
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

  // v24.87: roleBadge entfernt (war nur für die gelöschte „User Information"-Karte).

  // ===================== v24.85: Role-Management nach Kategorien =====================
  // Test-Team + Check-in-Team über ALLE Events aggregieren (read-only, analog
  // zu den Per-Event-Co-Organizern). Datenwerte: event.testTeamEmails/-Names
  // bzw. event.qrScannerEmails/-Names.
  const aggregateTeam = React.useCallback((pick: (e: (typeof events)[number]) => { emails?: string[]; names?: string[] }) => {
    const acc: Record<string, { name: string; events: string[] }> = {};
    for (const evt of events) {
      const sel = pick(evt); const emails = sel.emails || []; const names = sel.names || [];
      for (let i = 0; i < emails.length; i++) {
        const lc = (emails[i] || '').toLowerCase(); if (!lc) continue;
        if (!acc[lc]) acc[lc] = { name: names[i] || lc, events: [] };
        if (acc[lc].events.indexOf(evt.title) < 0) acc[lc].events.push(evt.title);
      }
    }
    return Object.keys(acc).map(lc => ({ email: lc, name: acc[lc].name, events: acc[lc].events }))
      .sort((a, b) => (a.name || '').localeCompare(b.name || '', 'de'));
  }, [events]);
  const testersList = React.useMemo(() => aggregateTeam(e => ({ emails: e.testTeamEmails, names: e.testTeamNames })), [aggregateTeam]);
  const checkinList = React.useMemo(() => aggregateTeam(e => ({ emails: e.qrScannerEmails, names: e.qrScannerNames })), [aggregateTeam]);

  // Position (Job Title) + Standort pro Person live nachladen — DEX_Roles
  // speichert die Position nicht. Best-effort, 1× pro E-Mail gecacht.
  const [profiles, setProfiles] = React.useState<Record<string, { jobTitle?: string; location?: string }>>({});
  const profileAttemptedRef = React.useRef<Set<string>>(new Set());
  React.useEffect(() => {
    const emails = new Set<string>();
    roles.forEach(r => emails.add((r.userEmail || '').toLowerCase()));
    coOrganizersList.forEach(c => emails.add(c.email));
    testersList.forEach(c => emails.add(c.email));
    checkinList.forEach(c => emails.add(c.email));
    let cancelled = false;
    (async () => {
      for (const em of Array.from(emails)) {
        if (!em || profileAttemptedRef.current.has(em)) continue;
        profileAttemptedRef.current.add(em);
        try {
          const u = await searchUser(em);
          if (u && !cancelled && (u.jobTitle || u.location)) {
            setProfiles(prev => ({ ...prev, [em]: { jobTitle: u.jobTitle, location: u.location } }));
          }
        } catch { /* best-effort */ }
      }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [roles, coOrganizersList, testersList, checkinList]);

  // Klapp-Status pro Kategorie (Default: alle offen).
  const [openSections, setOpenSections] = React.useState<Set<string>>(() => new Set(['admins', 'organizer', 'coorg', 'tester', 'checkin', 'user']));
  const toggleSection = (k: string): void => setOpenSections(prev => { const n = new Set(prev); if (n.has(k)) n.delete(k); else n.add(k); return n; });
  // v24.87: Freitext-Suche über alle Kategorien (Name/E-Mail/Position/Standort).
  const [roleSearch, setRoleSearch] = React.useState('');
  // Deloitte-Displayname „Nachname, Vorname" → { first, last }.
  const splitName = (full: string): { first: string; last: string } => {
    const n = (full || '').trim();
    if (!n) return { first: '', last: '' };
    const c = n.indexOf(',');
    if (c >= 0) return { last: n.substring(0, c).trim(), first: n.substring(c + 1).trim() };
    const parts = n.split(/\s+/);
    return parts.length > 1 ? { first: parts[0], last: parts.slice(1).join(' ') } : { first: n, last: '' };
  };

  const renderRoleSections = (): React.ReactElement => {
    const thS: React.CSSProperties = { textAlign: 'left', padding: 8, color: 'var(--dex-gray-500)', fontSize: '0.76rem', fontWeight: 600, whiteSpace: 'nowrap' };
    const tdS: React.CSSProperties = { padding: 8, verticalAlign: 'middle' };
    const byName = (a: { userName: string }, b: { userName: string }): number => (a.userName || '').localeCompare(b.userName || '', 'de');
    // v24.87: Freitext-Filter (Name/E-Mail/Position/Standort) über alle Kategorien.
    const q = roleSearch.trim().toLowerCase();
    const hit = (parts: Array<string | undefined>): boolean => !q || parts.some(s => (s || '').toLowerCase().indexOf(q) >= 0);
    const matchRole = (r: typeof roles[number]): boolean => { const p = profiles[(r.userEmail || '').toLowerCase()] || {}; return hit([r.userName, r.userEmail, p.jobTitle, p.location, r.location]); };
    const matchAgg = (x: { email: string; name: string }): boolean => { const p = profiles[x.email] || {}; return hit([x.name, x.email, p.jobTitle, p.location]); };
    // v26.33: IT-Admins zählen zur Admin-Gruppe (gleiche Rechte).
    const admins = [...roles].filter(r => r.role === 'Admin' || r.role === 'IT-Admin').filter(matchRole).sort(byName);
    const organizers = [...roles].filter(r => r.role === 'Organizer').filter(matchRole).sort(byName);
    const usersLeft = [...roles].filter(r => r.role !== 'Admin' && r.role !== 'IT-Admin' && r.role !== 'Organizer').filter(matchRole).sort(byName);
    const eventBadges = (titles: string[]): React.ReactNode => titles.length === 0
      ? <span style={{ color: 'var(--dex-gray-300)' }}>—</span>
      : <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>{titles.map((t2, idx) => (
          <span key={idx} style={{ display: 'inline-block', padding: '2px 8px', borderRadius: 10, background: 'rgba(134,188,37,0.14)', color: 'var(--dex-green-dark)', fontSize: '0.72rem', fontWeight: 600 }}>{t2}</span>
        ))}</div>;
    const catPill = (label: string, bg: string, fg: string): React.ReactNode => (
      <span style={{ display: 'inline-block', padding: '2px 10px', borderRadius: 999, background: bg, color: fg, fontSize: '0.74rem', fontWeight: 600, whiteSpace: 'nowrap' }}>{label}</span>
    );
    const headRow = (
      <tr style={{ borderBottom: '2px solid var(--dex-gray-200, #eee)' }}>
        <th style={thS}>{isDe ? 'Vorname' : 'First name'}</th>
        <th style={thS}>{isDe ? 'Nachname' : 'Last name'}</th>
        <th style={thS}>Email</th>
        <th style={thS}>{isDe ? 'Position' : 'Position'}</th>
        <th style={thS}>{isDe ? 'Standort' : 'Location'}</th>
        <th style={thS}>Role</th>
        <th style={thS}>Status</th>
        <th style={thS}>Power User</th>
        <th style={thS}>Coordinated Events</th>
        <th style={{ ...thS, textAlign: 'right' }} />
      </tr>
    );
    // v24.87: Status-Badges (Admin/Organizer/User + Co-Organizer/Tester/Check-in).
    const adminPill = catPill('Admin', '#e8f5e9', '#2e7d32');
    // v26.33: IT-Admin — volle Admin-Rechte, aber kein Empfänger der Mails.
    const itAdminPill = catPill('IT-Admin', '#ede7f6', '#5e35b1');
    const organizerPill = catPill('Organizer', 'rgba(0,118,168,0.10)', 'var(--dex-blue, #0076a8)');
    const userPill = catPill('User', '#f5f5f5', '#666');
    const coOrgPill = catPill('Co-Organizer', 'rgba(237,139,0,0.15)', 'var(--dex-orange-dark, #b35a00)');
    // Editierbare Zeile (DEX_Roles: Admins / Organizer / User)
    const editableRow = (r: typeof roles[number]): React.ReactElement => {
      const { first, last } = splitName(r.userName);
      const emailLc = (r.userEmail || '').toLowerCase();
      const prof = profiles[emailLc] || {};
      const pos = prof.jobTitle || '';
      const isSelf = emailLc === currentUser.email.toLowerCase();
      const evts = organizerEventMap[emailLc] || [];
      return (
        <tr key={r.id} style={{ borderBottom: '1px solid var(--dex-gray-100, #f0f0f0)' }}>
          <td style={tdS}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
              <PersonContactHover email={r.userEmail} name={r.userName} size={30} subline={pos} isDe={isDe} />
              <span style={{ fontWeight: 500 }}>{first || '-'}</span>
            </div>
          </td>
          <td style={{ ...tdS, fontWeight: 500 }}>{last || '-'}</td>
          <td style={{ ...tdS, color: 'var(--dex-gray-600)' }}>{r.userEmail}</td>
          <td style={{ ...tdS, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{pos || '—'}</td>
          {/* v24.87: Standort ist NICHT editierbar — er ergibt sich aus dem
              Profil zur E-Mail-Adresse (live nachgeladen, Fallback DEX_Roles). */}
          <td style={{ ...tdS, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{prof.location || r.location || '—'}</td>
          <td style={tdS}>
            <select
              value={r.role}
              onChange={e => handleChangeRole(r.id, e.target.value as UserRole)}
              style={{ padding: '4px 8px', borderRadius: 6, border: '1px solid var(--dex-gray-200, #ddd)', fontSize: '0.8rem', background: '#fff' }}
              disabled={isSelf}
            >
              <option value="Admin">Admin</option>
              <option value="IT-Admin">IT-Admin</option>
              <option value="Organizer">Organizer</option>
              <option value="User">User</option>
            </select>
          </td>
          <td style={tdS}>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
              {r.role === 'Admin' && adminPill}
              {r.role === 'IT-Admin' && itAdminPill}
              {r.role === 'Organizer' && organizerPill}
              {r.role === 'User' && userPill}
              {/* v24.87: globaler Organizer, der zusätzlich Events betreut → auch Co-Organizer. */}
              {r.role === 'Organizer' && evts.length > 0 && coOrgPill}
            </div>
          </td>
          <td style={tdS}>
            {(r.role === 'Organizer' || r.role === 'Admin') ? (
              <button
                type="button"
                disabled={isSelf}
                onClick={() => setPowerUser(r.id, !r.isPowerUser)}
                title={r.isPowerUser ? (isDe ? 'Power-User-Status entfernen' : 'Remove power-user status') : (isDe ? 'Als Power User markieren' : 'Mark as power user')}
                style={{ display: 'inline-flex', alignItems: 'center', gap: 4, padding: '4px 10px', borderRadius: 999, cursor: isSelf ? 'default' : 'pointer', fontSize: '0.74rem', fontWeight: 600, whiteSpace: 'nowrap', border: `1px solid ${r.isPowerUser ? '#b35a00' : 'var(--dex-gray-300)'}`, background: r.isPowerUser ? '#fff4e5' : '#fff', color: r.isPowerUser ? '#b35a00' : 'var(--dex-gray-600)' }}
              >
                <span aria-hidden="true">{r.isPowerUser ? '★' : '☆'}</span>
                {r.isPowerUser ? 'Power User' : (isDe ? 'Power User?' : 'Power user?')}
              </button>
            ) : <span style={{ color: 'var(--dex-gray-300)' }}>—</span>}
          </td>
          <td style={{ ...tdS, fontSize: '0.78rem', color: 'var(--dex-gray-600)', maxWidth: 280 }}>
            {r.role === 'User' ? <span style={{ color: 'var(--dex-gray-300)' }}>—</span> : eventBadges(evts)}
          </td>
          <td style={{ ...tdS, textAlign: 'right' }}>
            {!isSelf && (
              <button onClick={() => handleRemoveRole(r.id, r.userName)} disabled={isRemoving === r.id} style={{ border: 'none', background: 'none', cursor: isRemoving === r.id ? 'wait' : 'pointer', color: 'var(--dex-danger, #e53935)', padding: 4, opacity: isRemoving === r.id ? 0.4 : 1 }} title="Rolle entfernen">
                {isRemoving === r.id ? '...' : <Trash2 size={16} />}
              </button>
            )}
          </td>
        </tr>
      );
    };
    // Read-only Zeile (aggregiert: Co-Organizer / Tester / Check-in)
    const aggRow = (p: { email: string; name: string; events: string[] }, rolePill: React.ReactNode): React.ReactElement => {
      const { first, last } = splitName(p.name);
      const prof = profiles[p.email] || {};
      const displayName = resolvedOrgNames[p.email] || p.name;
      const dn = splitName(displayName);
      return (
        <tr key={p.email} style={{ borderBottom: '1px solid var(--dex-gray-100, #f0f0f0)' }}>
          <td style={tdS}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
              <PersonContactHover email={p.email} name={displayName} size={30} subline={prof.jobTitle || ''} isDe={isDe} />
              <span style={{ fontWeight: 500 }}>{dn.first || first || '-'}</span>
            </div>
          </td>
          <td style={{ ...tdS, fontWeight: 500 }}>{dn.last || last || '-'}</td>
          <td style={{ ...tdS, color: 'var(--dex-gray-600)' }}>{p.email}</td>
          <td style={{ ...tdS, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{prof.jobTitle || '—'}</td>
          <td style={{ ...tdS, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{prof.location || '—'}</td>
          <td style={tdS}><span style={{ color: 'var(--dex-gray-300)' }}>—</span></td>
          <td style={tdS}>{rolePill}</td>
          <td style={tdS}><span style={{ color: 'var(--dex-gray-300)' }}>—</span></td>
          <td style={{ ...tdS, fontSize: '0.78rem', color: 'var(--dex-gray-600)', maxWidth: 280 }}>{eventBadges(p.events)}</td>
          <td style={{ ...tdS, textAlign: 'right' }} />
        </tr>
      );
    };
    const coorgF = coOrganizersList.filter(matchAgg);
    const testersF = testersList.filter(matchAgg);
    const checkinF = checkinList.filter(matchAgg);
    // v24.87: Organizer + Co-Organizer in EINEM Abschnitt — globale Organizer
    // (editierbar) zuerst, danach die reinen Per-Event-Co-Organizer (read-only).
    // Die Status-Spalte zeigt „Organizer" und/oder „Co-Organizer".
    const organizerSectionBody = [
      ...organizers.map(editableRow),
      ...coorgF.map(p => aggRow(p, coOrgPill)),
    ];
    const sections: Array<{ key: string; title: string; body: React.ReactNode; count: number }> = [
      { key: 'admins', title: 'Admins', count: admins.length, body: admins.map(editableRow) },
      { key: 'organizer', title: isDe ? 'Organizer & Co-Organizer' : 'Organizers & co-organizers', count: organizers.length + coorgF.length, body: organizerSectionBody },
      { key: 'tester', title: isDe ? 'Tester' : 'Testers', count: testersF.length, body: testersF.map(p => aggRow(p, catPill(isDe ? 'Tester' : 'Tester', 'rgba(0,118,168,0.10)', 'var(--dex-blue, #0076a8)'))) },
      { key: 'checkin', title: 'Check-in', count: checkinF.length, body: checkinF.map(p => aggRow(p, catPill('Check-in', 'rgba(134,188,37,0.16)', 'var(--dex-green-dark, #4a7c1f)'))) },
    ];
    if (usersLeft.length > 0) sections.push({ key: 'user', title: isDe ? 'Weitere (User)' : 'Other (users)', count: usersLeft.length, body: usersLeft.map(editableRow) });
    return (
      <div>
        {sections.map(sec => {
          const open = q ? true : openSections.has(sec.key);
          return (
            <div key={sec.key} style={{ marginBottom: 12, border: '1px solid var(--dex-gray-200, #eee)', borderRadius: 8, overflow: 'hidden' }}>
              <button
                type="button"
                onClick={() => toggleSection(sec.key)}
                style={{ width: '100%', textAlign: 'left', display: 'flex', alignItems: 'center', gap: 8, padding: '10px 14px', background: 'var(--dex-gray-50, #fafafa)', border: 'none', cursor: 'pointer', fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-gray-800)' }}
              >
                <span style={{ color: 'var(--dex-gray-500)' }}>{open ? '▾' : '▸'}</span>
                {sec.title} <span style={{ color: 'var(--dex-gray-400)', fontWeight: 500 }}>({sec.count})</span>
              </button>
              {open && (
                <div style={{ overflowX: 'auto', padding: '0 6px 6px' }}>
                  {sec.count === 0 ? (
                    <p style={{ padding: '8px 10px', margin: 0, color: 'var(--dex-gray-400)', fontStyle: 'italic', fontSize: '0.82rem' }}>{isDe ? 'keine' : 'none'}</p>
                  ) : (
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                      <thead>{headRow}</thead>
                      <tbody>{sec.body}</tbody>
                    </table>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <div className="page-container">
      <div>

        {/* v24.87: „User Information"-Karte entfernt — diese Seite ist reine
            Rollenverwaltung (die eigenen User-Infos stehen im Header-Avatar-Menü). */}

        {/* v24.84: „Admin Actions"-Karte (Create New Event / View All Events /
            Rollen-Matrix) entfernt — diese Wege gibt es bereits über die
            Start-/Admin-Kacheln. Diese Seite ist jetzt reine Rollenverwaltung. */}

        {/* v9.21: Rollenmanagement collapsible — Admin kann die ganze Liste
            zusammenklappen wenn er sie gerade nicht braucht. Default: zu. */}
        {isAdmin && (
          <div className="card">
            {/* v24.85: nicht mehr die ganze Karte einklappbar — stattdessen je
                Kategorie ein eigener aufklappbarer Abschnitt (renderRoleSections). */}
            <h2 style={{ margin: '0 0 12px', fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Rollenverwaltung' : 'Role management'}
            </h2>
            <div style={{ marginTop: 4 }}>
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

            {/* Neue Rolle hinzufügen — v11.72: nach OBEN verschoben, damit der
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
                        src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(newEmail)}&size=L`}
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
                      <div style={{ marginTop: 2 }}>
                        <InternationalSearchToggle query={newEmail} checked={includeIntl} onChange={setIncludeIntl} isDe={isDe} />
                      </div>
                      {/* v11.75: explizite „Keine Treffer"-Box wenn die Suche
                          fertig ist und 0 Treffer hat — sonst wirkt der Picker
                          stumm und der Admin weiss nicht, ob die Suche lief. */}
                      {!isSearching && newEmail && newEmail.length >= 2 && suggestions.length === 0 && (
                        <div style={{
                          marginTop: 6, padding: '8px 12px', borderRadius: 'var(--dex-radius)',
                          border: '1px dashed var(--dex-gray-300)', background: 'var(--dex-gray-50)',
                          color: 'var(--dex-gray-600)', fontSize: '0.78rem',
                        }}>
                          Keine Treffer für &bdquo;{newEmail}&ldquo;. Versuche es mit dem vollen Namen oder der E-Mail-Adresse.
                        </div>
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
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(s.email)}&size=L`}
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
                    <option value="IT-Admin">IT-Admin (volle Rechte, keine Mails)</option>
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

            {/* v24.87: Freitext-Suche über alle Kategorien. */}
            {!isRolesLoading && (
              <div style={{ position: 'relative', marginBottom: 14, maxWidth: 360 }}>
                <input
                  className="form-input"
                  value={roleSearch}
                  onChange={e => setRoleSearch(e.target.value)}
                  placeholder={isDe ? 'Suchen (Name, E-Mail, Position, Standort)…' : 'Search (name, email, position, location)…'}
                  style={{ fontSize: '0.85rem', paddingRight: roleSearch ? 30 : undefined }}
                />
                {roleSearch && (
                  <button
                    type="button"
                    onClick={() => setRoleSearch('')}
                    title={isDe ? 'Suche zurücksetzen' : 'Clear search'}
                    style={{ position: 'absolute', right: 6, top: '50%', transform: 'translateY(-50%)', border: 'none', background: 'var(--dex-gray-100)', cursor: 'pointer', width: 22, height: 22, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--dex-gray-600)' }}
                  >
                    <X size={14} />
                  </button>
                )}
              </div>
            )}

            {/* v24.85: Rollen je Kategorie als eigene aufklappbare Abschnitte. */}
            {isRolesLoading ? (
              <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{isDe ? 'Rollen werden geladen…' : 'Loading roles...'}</p>
            ) : renderRoleSections()}

            </div>
          </div>
        )}

        {/* v9.16/v9.21: Test-Team war hier global, ist jetzt per-Event
            (im EventCreation-Wizard). TestTeamManager entfernt. */}

        {/* v24.86: „Default-Mail-Templates re-seed" und „Wochenbericht — Test-
            Versand" sind ins Admin Center (Aktionen-Dropdown) gewandert.
            v24.84: „Listen-Berechtigungen"-Übersicht entfernt. */}

      </div>

      {/* Onboarding-Mail-Prompt nach erfolgreicher Rollen-Zuweisung */}
      {/* v13.4: Onboarding-Prompt jetzt über das <Modal>-Wrapper-Component. */}
      <Modal
        open={!!onboardingPrompt}
        onClose={() => setOnboardingPrompt(null)}
        dismissable={!isSendingOnboarding}
        maxWidth={460}
        ariaLabel="Onboarding-Mail senden"
      >
        {onboardingPrompt && (
          <>
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
          </>
        )}
      </Modal>
    </div>
  );
}

