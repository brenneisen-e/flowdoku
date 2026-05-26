/**
 * v12.7: Demo-Impersonation-Modal
 *
 * Admin-only Tool: erlaubt es, die App temporär als ein anderer User
 * (mit gewähltem Standort) anzuzeigen. Wird über das Header-User-Menü
 * unter „Role Management" geöffnet. Beim Speichern landet die Auswahl
 * in localStorage und die Seite wird neu geladen — danach gilt für
 * RoleContext: role='User', isAdmin=false; UserContext: Name + Email
 * + Location aus den gewählten Werten.
 *
 * Beendet wird die Impersonation über den oben angedockten Banner
 * (siehe DexEventPlatform.tsx → ImpersonationBanner).
 */
import * as React from 'react';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import Modal from './Modal';

interface ImpersonateModalProps {
  open: boolean;
  onClose: () => void;
}

const STORAGE_KEY = 'dex_demo_impersonation';

export default function ImpersonateModal({ open, onClose }: ImpersonateModalProps): React.ReactElement | null {
  const { searchUsers } = useRoles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const [query, setQuery] = React.useState('');
  const [results, setResults] = React.useState<Array<{ email: string; displayName: string; location: string; jobTitle: string }>>([]);
  const [picked, setPicked] = React.useState<{ email: string; displayName: string; location: string } | null>(null);
  const [locationOverride, setLocationOverride] = React.useState('');
  const [isSearching, setIsSearching] = React.useState(false);
  const timer = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  React.useEffect(() => {
    if (!open) {
      setQuery('');
      setResults([]);
      setPicked(null);
      setLocationOverride('');
    }
  }, [open]);

  const handleQueryChange = (val: string): void => {
    setQuery(val);
    if (timer.current) clearTimeout(timer.current);
    if (val.length >= 2) {
      timer.current = setTimeout(async () => {
        setIsSearching(true);
        try { setResults(await searchUsers(val)); }
        catch { setResults([]); }
        setIsSearching(false);
      }, 300);
    } else {
      setResults([]);
    }
  };

  const handleActivate = (): void => {
    if (!picked) return;
    const parts = (picked.displayName || '').split(',');
    let firstName = '';
    let surname = '';
    if (parts.length === 2) {
      surname = parts[0].trim();
      firstName = parts[1].trim();
    } else {
      const words = picked.displayName.split(/\s+/);
      firstName = words[0] || '';
      surname = words.slice(1).join(' ') || '';
    }
    const finalLocation = (locationOverride || picked.location || '').trim();
    if (!finalLocation) {
      // eslint-disable-next-line no-alert
      alert(isDe ? 'Bitte einen Standort eingeben — das Standortfilter steuert was der User in der App sieht.' : 'Please enter a location — the location filter controls what the user sees.');
      return;
    }
    const payload = JSON.stringify({
      email: picked.email,
      firstName,
      surname,
      location: finalLocation,
    });
    try { window.localStorage.setItem(STORAGE_KEY, payload); }
    catch { /* */ }
    // Reload, damit RoleContext + UserContext die Impersonation lesen.
    window.location.reload();
  };

  // v13.1: Modal-Wrapper-Komponente — kapselt Backdrop/Escape/Padding.
  return (
    <Modal
      open={open}
      onClose={onClose}
      maxWidth={520}
      ariaLabel={isDe ? 'Demo-Modus' : 'Demo mode'}
    >
        <h2 style={{ margin: 0, fontSize: '1.2rem', color: 'var(--dex-gray-800)' }}>
          {isDe ? 'Demo-Modus: als User testen' : 'Demo mode: act as a user'}
        </h2>
        <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
          {isDe
            ? 'Wähle einen Deloitte-Kollegen und einen Standort — die App tut danach so, als wärst du dieser User (Rolle „User", Standortfilter aktiv). Beenden über den orangen Banner oben.'
            : 'Pick a Deloitte colleague and a location — the app will then act as if you were that user (role "User", location filter active). End via the orange banner at the top.'}
        </p>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {isDe ? 'Person suchen' : 'Find person'}
          <input
            type="text"
            value={query}
            onChange={e => handleQueryChange(e.target.value)}
            placeholder={isDe ? 'Name oder E-Mail eingeben…' : 'Type a name or email…'}
            style={{ padding: '8px 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.9rem' }}
          />
          {isSearching && (
            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)', marginTop: 2 }}>
              {isDe ? 'Suche…' : 'Searching…'}
            </div>
          )}
          {results.length > 0 && !picked && (
            <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, marginTop: 4, maxHeight: 220, overflowY: 'auto' }}>
              {results.map(u => (
                <div
                  key={u.email}
                  onClick={() => {
                    setPicked({ email: u.email, displayName: u.displayName, location: u.location });
                    setLocationOverride(u.location || '');
                    setResults([]);
                    setQuery(u.displayName);
                  }}
                  style={{ padding: '8px 12px', cursor: 'pointer', borderBottom: '1px solid var(--dex-gray-100)', fontSize: '0.85rem' }}
                >
                  <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                  <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem' }}>
                    {u.email}{u.location ? ` · ${u.location}` : ''}
                  </div>
                </div>
              ))}
            </div>
          )}
        </label>
        {picked && (
          <div style={{ padding: '10px 12px', background: 'var(--dex-gray-50, #f7f7f7)', border: '1px solid var(--dex-gray-200)', borderRadius: 8, fontSize: '0.85rem' }}>
            <strong>{picked.displayName}</strong>
            <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem' }}>{picked.email}</div>
          </div>
        )}
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {isDe ? 'Standort (steuert den Standortfilter — Pflicht)' : 'Location (drives the location filter — required)'}
          <input
            type="text"
            value={locationOverride}
            onChange={e => setLocationOverride(e.target.value)}
            placeholder={isDe ? 'z.B. „DE - München"' : 'e.g. „DE - Munich"'}
            style={{ padding: '8px 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.9rem' }}
          />
        </label>
        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 4 }}>
          <button type="button" className="btn btn-secondary" onClick={onClose}>
            {isDe ? 'Abbrechen' : 'Cancel'}
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={handleActivate}
            disabled={!picked || !locationOverride.trim()}
          >
            {isDe ? 'Demo-Modus starten' : 'Start demo mode'}
          </button>
        </div>
    </Modal>
  );
}
