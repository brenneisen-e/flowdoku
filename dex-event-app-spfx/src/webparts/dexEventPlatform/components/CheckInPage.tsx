/**
 * Check-In Seite - Oeffnet den QR-Code Scanner in einem neuen Fenster
 *
 * Der Scanner laeuft als eigenstaendige HTML-Seite (checkin.html in SiteAssets),
 * damit die Kamera ohne iFrame-Einschraenkungen funktioniert.
 * Fallback: Manuelle Code-Eingabe direkt in der App.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { EventService, SPRegistration } from '../services/EventService';
import { useLanguage } from '../context/LanguageContext';

export default function CheckInPage(): React.ReactElement {
  const { events } = useEvents();
  const { selectedEventId } = useNavigation();
  const { isAdmin, isOrganizer, siteUrl } = useRoles();
  const { t } = useLanguage();
  const [manualCode, setManualCode] = React.useState('');
  const [resultMessage, setResultMessage] = React.useState('');
  const [resultType, setResultType] = React.useState<'success' | 'error' | 'info' | ''>('');
  const [isProcessing, setIsProcessing] = React.useState(false);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const context = (window as any).__dexSpfxContext;
  const eventService = React.useMemo(() => context ? new EventService(context) : null, []);

  const selectedEvent = selectedEventId ? events.find(e => e.id === selectedEventId) : null;

  // Scanner in neuem Fenster oeffnen
  const openScanner = (): void => {
    const checkinUrl = `${siteUrl}/SiteAssets/checkin.html`;
    const params = new URLSearchParams();
    params.set('siteUrl', siteUrl);
    if (selectedEvent) {
      params.set('subsiteUrl', selectedEvent.subsiteUrl || '');
      params.set('eventTitle', selectedEvent.title);
      params.set('eventNumber', (selectedEvent.eventNumber || 0).toString());
    }
    window.open(`${checkinUrl}?${params.toString()}`, '_blank', 'width=500,height=800,menubar=no,toolbar=no');
  };

  // Manuelle Code-Verarbeitung
  const handleManualSubmit = async (): Promise<void> => {
    if (!manualCode.trim() || !eventService) return;
    setIsProcessing(true);
    setResultMessage('');
    setResultType('');

    const parts = manualCode.trim().split('|');
    if (parts.length !== 3 || parts[0] !== 'DEX') {
      setResultMessage('Ungültiger Code. Format: DEX|EventNumber|Email');
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const eventNumber = parseInt(parts[1], 10);
    const email = parts[2];
    const event = events.find(e => e.eventNumber === eventNumber);

    if (!event || !event.subsiteUrl) {
      setResultMessage(`Event #${eventNumber} nicht gefunden.`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const reg = await eventService.getRegistrationByEmail(event.subsiteUrl, email);
    if (!reg) {
      setResultMessage(`${email} ist nicht für dieses Event registriert.`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    if (reg.Status === 'Eingecheckt') {
      const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
      setResultMessage(`${name} — bereits eingecheckt.`);
      setResultType('info');
      setIsProcessing(false);
      return;
    }

    if (reg.Status === 'Abgemeldet') {
      setResultMessage(`Teilnehmer ist abgemeldet — Check-in nicht möglich.`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const success = await eventService.checkInParticipant(event.subsiteUrl, reg.Id);
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
    if (success) {
      setResultMessage(`${name} — erfolgreich eingecheckt!`);
      setResultType('success');
    } else {
      setResultMessage(`Check-in fehlgeschlagen für ${name}.`);
      setResultType('error');
    }

    setManualCode('');
    setIsProcessing(false);
  };

  if (!isAdmin && !isOrganizer) {
    return (
      <div className="page-container text-center">
        <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>Nur für Organizer und Admins verfügbar.</p>
      </div>
    );
  }

  return (
    <div className="page-container">
      <h2 style={{ marginBottom: 8 }}>
        Check-in {selectedEvent ? `— ${selectedEvent.title}` : ''}
      </h2>
      <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.85rem', marginBottom: 24 }}>
        Öffne den QR-Code Scanner in einem neuen Fenster für vollen Kamera-Zugriff.
      </p>

      {/* Scanner-Button */}
      <div className="card" style={{ padding: 32, textAlign: 'center', marginBottom: 24 }}>
        <button
          className="btn btn-primary"
          onClick={openScanner}
          style={{ fontSize: '1.1rem', padding: '16px 40px' }}
        >
          QR-Code Scanner öffnen
        </button>
        <p style={{ color: 'var(--dex-gray-400)', fontSize: '0.8rem', marginTop: 12 }}>
          Öffnet ein neues Fenster mit Kamera-Zugriff für kontinuierliches Scannen.
          <br />Die Datei <code>checkin.html</code> muss in SiteAssets hochgeladen sein.
        </p>
      </div>

      {/* Manuelle Eingabe */}
      <div className="card" style={{ padding: 24 }}>
        <h3 style={{ marginBottom: 12 }}>Manuelle Eingabe</h3>
        <div style={{ display: 'flex', gap: 8 }}>
          <input
            className="form-input"
            value={manualCode}
            onChange={e => setManualCode(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') handleManualSubmit(); }}
            placeholder="DEX|1|email@deloitte.de"
            style={{ flex: 1 }}
            disabled={isProcessing}
          />
          <button
            className="btn btn-secondary"
            onClick={handleManualSubmit}
            disabled={isProcessing || !manualCode.trim()}
          >
            {isProcessing ? '...' : 'Einchecken'}
          </button>
        </div>

        {resultMessage && (
          <div style={{
            marginTop: 12, padding: '12px 16px', borderRadius: 8,
            background: resultType === 'success' ? '#e8f5e9' : resultType === 'error' ? '#ffebee' : '#e3f2fd',
            color: resultType === 'success' ? '#2e7d32' : resultType === 'error' ? '#c62828' : '#1565c0',
            fontWeight: 600,
          }}>
            {resultMessage}
          </div>
        )}
      </div>
    </div>
  );
}
