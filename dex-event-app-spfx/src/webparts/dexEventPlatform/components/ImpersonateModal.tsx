/**
 * v12.7: Demo-Impersonation-Modal
 *
 * Admin-only Tool: erlaubt es, die App temporär als ein synthetischer
 * User mit gewähltem Standort anzuzeigen. Wird über das Header-User-
 * Menü unter „Rollenverwaltung" geöffnet. Beim Speichern landet die
 * Auswahl in localStorage und die Seite wird neu geladen — danach gilt
 * für RoleContext: role='User', isAdmin=false; UserContext: Demo-
 * Name + Demo-Email + Location aus den gewählten Werten.
 *
 * v13.7: kein People-Picker mehr. Nur noch Standort-Dropdown.
 *
 * v13.11: optional zusätzlich „Check-In-Team"-Modus — der Demo-User
 * wird per qrScannerEmails synthetisch einem konkreten Event
 * zugeordnet. Damit kann der Admin testen, wie die Start-Seite und
 * die Check-In-Seite für reine Check-In-Helfer aussieht.
 *
 * Beendet wird die Impersonation über den oben angedockten Banner
 * (siehe DexEventPlatform.tsx → ImpersonationBanner).
 */
import * as React from 'react';
import { useLanguage } from '../context/LanguageContext';
import { useEvents } from '../context/EventContext';
import Modal from './Modal';

interface ImpersonateModalProps {
  open: boolean;
  onClose: () => void;
}

const STORAGE_KEY = 'dex_demo_impersonation';

const LOCATION_OPTIONS = [
  'Berlin', 'Dresden', 'Düsseldorf', 'Frankfurt', 'Görlitz', 'Halle',
  'Hamburg', 'Hannover', 'Köln', 'Leipzig', 'Magdeburg', 'Mannheim',
  'München', 'Nürnberg', 'Stuttgart', 'Walldorf',
];

export default function ImpersonateModal({ open, onClose }: ImpersonateModalProps): React.ReactElement | null {
  const { locale } = useLanguage();
  const { events } = useEvents();
  const isDe = locale === 'de';
  const [location, setLocation] = React.useState('');
  const [isCheckInTeam, setIsCheckInTeam] = React.useState(false);
  const [checkInEventId, setCheckInEventId] = React.useState('');

  React.useEffect(() => {
    if (!open) {
      setLocation('');
      setIsCheckInTeam(false);
      setCheckInEventId('');
    }
  }, [open]);

  const activeEvents = React.useMemo(() => {
    return (events || []).filter(e => e.status === 'Active');
  }, [events]);

  const handleActivate = (): void => {
    if (!location.trim()) return;
    if (isCheckInTeam && !checkInEventId) return;
    const payload = JSON.stringify({
      email: 'demo.user@deloitte.de',
      firstName: 'Demo',
      surname: 'User',
      location: location.trim(),
      checkInEventId: isCheckInTeam ? checkInEventId : '',
    });
    try { window.localStorage.setItem(STORAGE_KEY, payload); }
    catch { /* */ }
    window.location.reload();
  };

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
            ? 'Wähle einen Standort — die App tut danach so, als wärst du ein Mitarbeiter dieses Standorts (Rolle „User", Standortfilter aktiv). Beenden über den orangen Banner oben.'
            : 'Pick a location — the app will then act as if you were an employee from that office (role „User", location filter active). End via the orange banner at the top.'}
        </p>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {isDe ? 'Standort (steuert den Standortfilter — Pflicht)' : 'Location (drives the location filter — required)'}
          <select
            value={location}
            onChange={e => setLocation(e.target.value)}
            style={{ padding: '8px 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.9rem', background: '#fff' }}
          >
            <option value="">{isDe ? '— Standort wählen —' : '— Pick a location —'}</option>
            {LOCATION_OPTIONS.map(loc => (
              <option key={loc} value={loc}>{loc}</option>
            ))}
          </select>
        </label>
        <div style={{
          padding: '10px 12px',
          background: 'var(--dex-gray-50, #f7f7f7)',
          border: '1px solid var(--dex-gray-200)',
          borderRadius: 10,
          display: 'flex', flexDirection: 'column', gap: 8,
        }}>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', color: 'var(--dex-gray-700)', cursor: 'pointer' }}>
            <input
              type="checkbox"
              checked={isCheckInTeam}
              onChange={e => {
                setIsCheckInTeam(e.target.checked);
                if (!e.target.checked) setCheckInEventId('');
              }}
              style={{ cursor: 'pointer' }}
            />
            <span>
              {isDe
                ? 'Ich gehöre zum Check-In-Team eines Events'
                : 'I belong to the check-in team of an event'}
            </span>
          </label>
          {isCheckInTeam && (
            <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
              {isDe ? 'Event auswählen (Pflicht)' : 'Pick event (required)'}
              <select
                value={checkInEventId}
                onChange={e => setCheckInEventId(e.target.value)}
                style={{ padding: '8px 10px', border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.9rem', background: '#fff' }}
              >
                <option value="">{isDe ? '— Event wählen —' : '— Pick event —'}</option>
                {activeEvents.map(ev => (
                  <option key={ev.id} value={ev.id}>
                    {ev.title}{ev.eventNumber ? ` (#${ev.eventNumber})` : ''}
                  </option>
                ))}
              </select>
              {activeEvents.length === 0 && (
                <span style={{ fontSize: '0.75rem', color: 'var(--dex-orange, #ed8b00)' }}>
                  {isDe ? 'Keine aktiven Events verfügbar.' : 'No active events available.'}
                </span>
              )}
            </label>
          )}
        </div>
        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 4 }}>
          <button type="button" className="btn btn-secondary" onClick={onClose}>
            {isDe ? 'Abbrechen' : 'Cancel'}
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={handleActivate}
            disabled={!location.trim() || (isCheckInTeam && !checkInEventId)}
          >
            {isDe ? 'Demo-Modus starten' : 'Start demo mode'}
          </button>
        </div>
    </Modal>
  );
}
