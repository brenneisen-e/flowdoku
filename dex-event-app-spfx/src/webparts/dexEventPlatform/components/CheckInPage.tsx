/**
 * Check-In Seite - QR-Code Scanner fuer Event-Check-in
 *
 * Nutzt die BarcodeDetector API (Chrome/Edge) zum Scannen.
 * QR-Code Format: DEX|{EventNumber}|{Email}
 * Nach dem Scan wird der Teilnehmer angezeigt und kann eingecheckt werden.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { EventService, SPRegistration } from '../services/EventService';
import { DeloitteEvent } from '../types';

interface ScanResult {
  eventNumber: number;
  email: string;
  event: DeloitteEvent | null;
  registration: SPRegistration | null;
  error?: string;
}

export default function CheckInPage(): React.ReactElement {
  const { events } = useEvents();
  const { navigate, selectedEventId } = useNavigation();
  const { isAdmin, isOrganizer } = useRoles();
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const [isScanning, setIsScanning] = React.useState(false);
  const [scanResult, setScanResult] = React.useState<ScanResult | null>(null);
  const [isCheckingIn, setIsCheckingIn] = React.useState(false);
  const [checkInSuccess, setCheckInSuccess] = React.useState(false);
  const [manualCode, setManualCode] = React.useState('');
  const [statusMessage, setStatusMessage] = React.useState('');
  const [checkedInCount, setCheckedInCount] = React.useState(0);
  const scanIntervalRef = React.useRef<number | null>(null);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const context = (window as any).__dexSpfxContext;
  const eventService = React.useMemo(() => context ? new EventService(context) : null, []);

  // Event-Filter: Wenn von der AdminPage mit eventId navigiert, nur dieses Event
  const selectedEvent = selectedEventId ? events.find(e => e.id === selectedEventId) : null;

  // EventNumber -> Event Map
  const eventByNumber = React.useMemo(() => {
    const map: Record<number, DeloitteEvent> = {};
    for (const e of events) {
      if (e.eventNumber) map[e.eventNumber] = e;
    }
    return map;
  }, [events]);

  // Kamera starten
  const startScanning = async (): Promise<void> => {
    // Pruefen ob getUserMedia verfuegbar ist
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      setStatusMessage('Kamera-API nicht verfügbar. Die Seite muss über HTTPS geladen werden.');
      return;
    }

    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: 'environment' }, width: { ideal: 640 }, height: { ideal: 480 } },
      });
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        await videoRef.current.play();
        setIsScanning(true);
        startDetection();
      }
    } catch (err) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const error = err as any;
      if (error?.name === 'NotAllowedError') {
        setStatusMessage('Kamera-Berechtigung wurde verweigert. Bitte in den Browser-Einstellungen erlauben und Seite neu laden.');
      } else if (error?.name === 'NotFoundError') {
        setStatusMessage('Keine Kamera gefunden. Bitte ein Gerät mit Kamera verwenden.');
      } else if (error?.name === 'NotReadableError') {
        setStatusMessage('Kamera wird bereits von einer anderen App verwendet. Bitte andere Apps schließen.');
      } else {
        setStatusMessage(`Kamera konnte nicht gestartet werden: ${error?.message || 'Unbekannter Fehler'}. Bitte Code manuell eingeben.`);
      }
    }
  };

  // Kamera stoppen
  const stopScanning = (): void => {
    if (scanIntervalRef.current) {
      clearInterval(scanIntervalRef.current);
      scanIntervalRef.current = null;
    }
    if (videoRef.current && videoRef.current.srcObject) {
      const tracks = (videoRef.current.srcObject as MediaStream).getTracks();
      tracks.forEach(t => t.stop());
      videoRef.current.srcObject = null;
    }
    setIsScanning(false);
  };

  // BarcodeDetector Loop
  const startDetection = (): void => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const BarcodeDetectorClass = (window as any).BarcodeDetector;
    if (!BarcodeDetectorClass) {
      setStatusMessage('BarcodeDetector wird von diesem Browser nicht unterstützt. Bitte Chrome oder Edge verwenden, oder Code manuell eingeben.');
      return;
    }

    const detector = new BarcodeDetectorClass({ formats: ['qr_code'] });

    scanIntervalRef.current = window.setInterval(async () => {
      if (!videoRef.current || videoRef.current.readyState !== 4) return;
      try {
        const barcodes = await detector.detect(videoRef.current);
        if (barcodes.length > 0) {
          const code = barcodes[0].rawValue;
          if (code) {
            stopScanning();
            await processCode(code);
          }
        }
      } catch { /* ignore detection errors */ }
    }, 500);
  };

  // QR-Code verarbeiten
  const processCode = async (code: string): Promise<void> => {
    setScanResult(null);
    setCheckInSuccess(false);
    setStatusMessage('');

    // Format: DEX|{EventNumber}|{Email}
    const parts = code.split('|');
    if (parts.length !== 3 || parts[0] !== 'DEX') {
      setScanResult({ eventNumber: 0, email: '', event: null, registration: null, error: 'Ungültiger QR-Code. Format: DEX|EventNumber|Email' });
      return;
    }

    const eventNumber = parseInt(parts[1], 10);
    const email = parts[2];

    if (isNaN(eventNumber) || !email) {
      setScanResult({ eventNumber: 0, email: '', event: null, registration: null, error: 'Ungültige Daten im QR-Code.' });
      return;
    }

    const event = eventByNumber[eventNumber];
    if (!event) {
      setScanResult({ eventNumber, email, event: null, registration: null, error: `Event #${eventNumber} nicht gefunden.` });
      return;
    }

    // Registrierung laden
    if (!eventService || !event.subsiteUrl) {
      setScanResult({ eventNumber, email, event, registration: null, error: 'Event-Subsite nicht verfügbar.' });
      return;
    }

    const reg = await eventService.getRegistrationByEmail(event.subsiteUrl, email);
    if (!reg) {
      setScanResult({ eventNumber, email, event, registration: null, error: `${email} ist nicht für dieses Event registriert.` });
      return;
    }

    setScanResult({ eventNumber, email, event, registration: reg });
  };

  // Check-in durchfuehren
  const handleCheckIn = async (): Promise<void> => {
    if (!scanResult?.event || !scanResult?.registration || !eventService) return;
    setIsCheckingIn(true);

    const success = await eventService.checkInParticipant(
      scanResult.event.subsiteUrl!,
      scanResult.registration.Id
    );

    if (success) {
      setCheckInSuccess(true);
      setCheckedInCount(prev => prev + 1);
      // Update result to show new status
      setScanResult({
        ...scanResult,
        registration: { ...scanResult.registration, Status: 'Eingecheckt' },
      });
    } else {
      setStatusMessage('Check-in fehlgeschlagen. Bitte erneut versuchen.');
    }
    setIsCheckingIn(false);
  };

  // Manueller Code eingeben
  const handleManualSubmit = async (): Promise<void> => {
    if (!manualCode.trim()) return;
    await processCode(manualCode.trim());
    setManualCode('');
  };

  // Neuen Scan starten
  const resetScan = (): void => {
    setScanResult(null);
    setCheckInSuccess(false);
    setStatusMessage('');
  };

  // Cleanup
  React.useEffect(() => {
    return () => {
      if (scanIntervalRef.current) clearInterval(scanIntervalRef.current);
      if (videoRef.current && videoRef.current.srcObject) {
        const tracks = (videoRef.current.srcObject as MediaStream).getTracks();
        tracks.forEach(t => t.stop());
      }
    };
  }, []);

  if (!isAdmin && !isOrganizer) {
    return (
      <div className="page-container text-center">
        <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>Nur für Organizer und Admins verfügbar.</p>
      </div>
    );
  }

  const getStatusColor = (status: string): string => {
    switch (status) {
      case 'Eingecheckt': return 'var(--dex-green)';
      case 'Angemeldet': case 'QR versendet': return '#1565c0';
      case 'Warteliste': return 'var(--dex-orange)';
      case 'Abgemeldet': return 'var(--dex-red)';
      default: return 'var(--dex-gray-400)';
    }
  };

  const displayName = (reg: SPRegistration): string =>
    (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;

  return (
    <div className="page-container">
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
        <h2 style={{ margin: 0 }}>
          Check-in {selectedEvent ? `— ${selectedEvent.title}` : ''}
        </h2>
        {checkedInCount > 0 && (
          <span style={{ fontSize: '0.9rem', fontWeight: 600, color: 'var(--dex-green)' }}>
            {checkedInCount} eingecheckt in dieser Session
          </span>
        )}
      </div>

      {/* Scanner-Bereich */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        {!isScanning && !scanResult && (
          <div style={{ textAlign: 'center' }}>
            <button className="btn btn-primary" onClick={startScanning} style={{ fontSize: '1rem', padding: '12px 32px' }}>
              📷 QR-Code scannen
            </button>
            <p style={{ color: 'var(--dex-gray-400)', fontSize: '0.85rem', marginTop: 12 }}>
              Oder Code manuell eingeben:
            </p>
            <div style={{ display: 'flex', gap: 8, justifyContent: 'center', marginTop: 8 }}>
              <input
                className="form-input"
                value={manualCode}
                onChange={e => setManualCode(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter') handleManualSubmit(); }}
                placeholder="DEX|1|email@deloitte.de"
                style={{ maxWidth: 300 }}
              />
              <button className="btn btn-secondary" onClick={handleManualSubmit}>Prüfen</button>
            </div>
          </div>
        )}

        {isScanning && (
          <div style={{ textAlign: 'center' }}>
            <video
              ref={videoRef}
              style={{ width: '100%', maxWidth: 400, borderRadius: 'var(--dex-radius)', border: '3px solid var(--dex-green)' }}
              playsInline
              muted
            />
            <p style={{ color: 'var(--dex-green)', fontWeight: 600, marginTop: 8 }}>Scanne QR-Code...</p>
            <button className="btn btn-secondary" onClick={stopScanning} style={{ marginTop: 8 }}>Abbrechen</button>
          </div>
        )}

        {statusMessage && (
          <p style={{ color: 'var(--dex-orange)', textAlign: 'center', marginTop: 12 }}>{statusMessage}</p>
        )}
      </div>

      {/* Scan-Ergebnis */}
      {scanResult && (
        <div className="card" style={{ padding: 24 }}>
          {scanResult.error ? (
            <div style={{ textAlign: 'center' }}>
              <p style={{ color: 'var(--dex-red)', fontWeight: 600, fontSize: '1.1rem' }}>
                {scanResult.error}
              </p>
              <button className="btn btn-primary" onClick={resetScan} style={{ marginTop: 16 }}>
                Neuer Scan
              </button>
            </div>
          ) : (
            <div>
              {/* Teilnehmer-Info */}
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 }}>
                <div>
                  <h3 style={{ margin: '0 0 4px 0', fontSize: '1.3rem' }}>
                    {scanResult.registration ? displayName(scanResult.registration) : scanResult.email}
                  </h3>
                  <p style={{ color: 'var(--dex-gray-500)', margin: '0 0 8px 0' }}>{scanResult.email}</p>
                  <p style={{ margin: 0 }}>
                    <strong>Event:</strong> {scanResult.event?.title || `#${scanResult.eventNumber}`}
                  </p>
                </div>
                <span style={{
                  padding: '6px 16px',
                  borderRadius: 20,
                  fontSize: '0.9rem',
                  fontWeight: 700,
                  color: '#fff',
                  background: scanResult.registration ? getStatusColor(scanResult.registration.Status) : 'var(--dex-gray-400)',
                }}>
                  {scanResult.registration?.Status || 'Unbekannt'}
                </span>
              </div>

              {/* Check-in Aktion */}
              {scanResult.registration && scanResult.registration.Status !== 'Eingecheckt' && scanResult.registration.Status !== 'Abgemeldet' && (
                <button
                  className="btn btn-primary"
                  onClick={handleCheckIn}
                  disabled={isCheckingIn}
                  style={{ width: '100%', fontSize: '1.1rem', padding: '14px 0', background: 'var(--dex-green)' }}
                >
                  {isCheckingIn ? 'Wird eingecheckt...' : 'Einchecken'}
                </button>
              )}

              {checkInSuccess && (
                <div style={{ textAlign: 'center', marginTop: 16, padding: 16, background: '#e8f5e9', borderRadius: 'var(--dex-radius)' }}>
                  <p style={{ color: '#2e7d32', fontWeight: 700, fontSize: '1.1rem', margin: 0 }}>
                    Erfolgreich eingecheckt!
                  </p>
                </div>
              )}

              {scanResult.registration?.Status === 'Eingecheckt' && !checkInSuccess && (
                <div style={{ textAlign: 'center', padding: 16, background: '#fff3e0', borderRadius: 'var(--dex-radius)' }}>
                  <p style={{ color: '#e65100', fontWeight: 600, margin: 0 }}>
                    Bereits eingecheckt
                  </p>
                </div>
              )}

              {scanResult.registration?.Status === 'Abgemeldet' && (
                <div style={{ textAlign: 'center', padding: 16, background: '#ffebee', borderRadius: 'var(--dex-radius)' }}>
                  <p style={{ color: 'var(--dex-red)', fontWeight: 600, margin: 0 }}>
                    Teilnehmer ist abgemeldet — Check-in nicht möglich
                  </p>
                </div>
              )}

              <button className="btn btn-secondary" onClick={resetScan} style={{ width: '100%', marginTop: 12 }}>
                Nächsten Teilnehmer scannen
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
