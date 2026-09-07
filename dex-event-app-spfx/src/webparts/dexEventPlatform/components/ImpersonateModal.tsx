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
 *
 * v31.2: Optik auf die dex-ui-Klassen umgestellt (Modal-Kopf/-Fuß, Standort
 * als Chip-Reihe, Rollen als Schalter-Zeilen mit Folge-Satz). Verhalten,
 * State und Payload sind unverändert.
 */
import * as React from 'react';
import { useLanguage } from '../context/LanguageContext';
import { useEvents } from '../context/EventContext';
import Modal from './Modal';
// v31.2: Gemeinsame UI-Klassen (Chips, Schalter-Zeilen, Callout) — Hover und
// Fokus kommen aus dem Stylesheet, nicht aus Inline-Styles (die können kein
// :hover). Das Symbol für den Dialog-Kopf kommt aus Icons.tsx.
import { cx } from './dexUi';
import { Users } from './Icons';

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
  // v29.73: Demo-User gehoert zum Verteilerkreis. Events mit Verteiler-
  // Sichtbarkeit waren im Demo-Modus unsichtbar bzw. gesperrt — der
  // synthetische demo.user steht in keinem Verteiler. Damit liess sich die
  // Anmeldemaske solcher Events nie aus Teilnehmersicht ansehen.
  const [isAudienceMember, setIsAudienceMember] = React.useState(false);
  const [isCheckInTeam, setIsCheckInTeam] = React.useState(false);
  const [checkInEventId, setCheckInEventId] = React.useState('');

  React.useEffect(() => {
    if (!open) {
      setLocation('');
      setIsAudienceMember(false);
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
      audienceMember: isAudienceMember,
      checkInEventId: isCheckInTeam ? checkInEventId : '',
    });
    try { window.localStorage.setItem(STORAGE_KEY, payload); }
    catch { /* */ }
    window.location.reload();
  };

  // v31.2: Der Startknopf war gesperrt, ohne zu sagen warum. Der Fuß nennt
  // jetzt die fehlende Angabe — dieselbe Bedingung wie `disabled` unten.
  const missingLocation = !location.trim();
  const missingCheckInEvent = isCheckInTeam && !checkInEventId;
  const blockerHint = missingLocation
    ? (isDe ? 'Wähle zuerst einen Standort.' : 'Pick a location first.')
    : missingCheckInEvent
      ? (isDe ? 'Wähle noch das Event für das Check-In-Team.' : 'Pick the event for the check-in team.')
      : '';

  return (
    <Modal
      open={open}
      onClose={onClose}
      maxWidth={520}
      ariaLabel={isDe ? 'Demo-Modus' : 'Demo mode'}
      icon={<Users size={20} />}
      title={isDe ? 'Als Teilnehmer testen' : 'Test as an attendee'}
      subtitle={isDe
        ? 'Die App zeigt dir danach alles so, wie es ein Mitarbeiter des gewählten Standorts sieht — Rolle „User“, Standortfilter aktiv.'
        : 'The app then shows you everything the way an employee of the chosen office sees it — role “User”, location filter active.'}
      footer={<>
        {blockerHint && <span className="dex-ui-modal-foot-left dex-ui-muted">{blockerHint}</span>}
        <button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" onClick={handleActivate} disabled={!location.trim() || (isCheckInTeam && !checkInEventId)}>
          {isDe ? 'Demo-Modus starten' : 'Start demo mode'}
        </button>
      </>}
    >
      {/* v31.2: Pflichtfrage zuerst. 16 kurze Werte sind eine Chip-Reihe
          (Leitfaden 2b), kein Dropdown — man sieht alle Standorte auf einen
          Blick und jeder Chip hat einen Hover. Bindung unverändert:
          `location` ↔ `setLocation`. */}
      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{isDe ? 'Standort' : 'Location'}</div>
        <div className="dex-ui-field">
          <div className="dex-ui-label">
            {isDe ? 'An welchem Standort arbeitet die Testperson?' : 'Which office does the test person work at?'}
          </div>
          <div className="dex-ui-inline" role="group" aria-label={isDe ? 'Standort wählen' : 'Pick a location'}>
            {LOCATION_OPTIONS.map(loc => (
              <button key={loc} type="button" className={cx('dex-ui-chip', location === loc && 'is-active')} aria-pressed={location === loc} onClick={() => setLocation(loc)}>
                {loc}
              </button>
            ))}
          </div>
          <div className="dex-ui-help">
            {isDe ? 'Pflicht — der Standort steuert den Standortfilter der Event-Übersicht.' : 'Required — the office drives the location filter of the event overview.'}
          </div>
        </div>
      </div>

      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{isDe ? 'Zusätzliche Rollen (optional)' : 'Additional roles (optional)'}</div>
        <div className="dex-ui-stack">
          {/* v29.73: Verteilerkreis-Simulation. Ohne sie sind Events mit
              Verteiler-Sichtbarkeit im Demo-Modus unsichtbar — demo.user steht
              in keinem Verteiler, die Anmeldemaske solcher Events liess sich
              also nie aus Teilnehmersicht pruefen. Der STANDORT-Filter bleibt
              bewusst aktiv, sonst waere die Auswahl oben wirkungslos. */}
          <label className={cx('dex-ui-toggle-row', isAudienceMember && 'is-active')}>
            <input type="checkbox" checked={isAudienceMember} onChange={e => setIsAudienceMember(e.target.checked)} />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">
                {isDe ? 'Ich gehöre zum Verteilerkreis der Events' : 'I belong to the audience of the events'}
              </span>
              <span className="dex-ui-toggle-row-desc">
                {isDe
                  ? 'Dann gilt die Verteiler-Sichtbarkeit für dich als erfüllt: Du siehst die Anmeldemaske wie ein Mitglied des Verteilers. Der Standortfilter bleibt aktiv.'
                  : 'Then audience visibility counts as satisfied for you: you see the registration page as an audience member would. The location filter stays active.'}
              </span>
            </span>
          </label>

          {/* v31.2: Die Event-Auswahl folgt direkt ihrem Schalter (Leitfaden 2a). */}
          <div className={cx(!isCheckInTeam && 'dex-ui-card--muted')}>
            <label className={cx('dex-ui-toggle-row', isCheckInTeam && 'is-active')}>
              <input
                type="checkbox"
                checked={isCheckInTeam}
                onChange={e => { setIsCheckInTeam(e.target.checked); if (!e.target.checked) setCheckInEventId(''); }}
              />
              <span className="dex-ui-toggle-row-body">
                <span className="dex-ui-toggle-row-title">
                  {isDe ? 'Ich gehöre zum Check-In-Team eines Events' : 'I belong to the check-in team of an event'}
                </span>
                <span className="dex-ui-toggle-row-desc">
                  {isDe
                    ? 'Dann siehst du Startseite und Check-In-Seite so, wie sie ein reiner Check-In-Helfer sieht.'
                    : 'Then you see the start page and the check-in page the way a pure check-in helper does.'}
                </span>
              </span>
            </label>
            {isCheckInTeam && (
              <div className="dex-ui-field dex-ui-fade-in" style={{ marginTop: 10, paddingLeft: 14 }}>
                <label className="dex-ui-label" htmlFor="dex-impersonate-checkin-event">
                  {isDe ? 'Für welches Event?' : 'For which event?'}
                </label>
                <select id="dex-impersonate-checkin-event" className="dex-ui-select" value={checkInEventId} onChange={e => setCheckInEventId(e.target.value)}>
                  <option value="">{isDe ? '— Event wählen —' : '— Pick event —'}</option>
                  {activeEvents.map(ev => (
                    <option key={ev.id} value={ev.id}>{ev.title}{ev.eventNumber ? ` (#${ev.eventNumber})` : ''}</option>
                  ))}
                </select>
                <div className="dex-ui-help">
                  {activeEvents.length === 0
                    ? <span style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>{isDe ? 'Keine aktiven Events verfügbar.' : 'No active events available.'}</span>
                    : (isDe ? 'Pflicht — nur aktive Events stehen zur Auswahl.' : 'Required — only active events are listed.')}
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* v31.2: Was nach dem Klick passiert, stand bisher im Einleitungstext —
          hier steht es dort, wo der Knopf ist. */}
      <div className="dex-ui-callout dex-ui-callout--neutral">
        <span>{isDe
          ? 'Beim Start lädt die Seite neu. Beenden kannst du den Demo-Modus jederzeit über den orangen Banner oben.'
          : 'Starting reloads the page. You can end demo mode any time via the orange banner at the top.'}</span>
      </div>
    </Modal>
  );
}
