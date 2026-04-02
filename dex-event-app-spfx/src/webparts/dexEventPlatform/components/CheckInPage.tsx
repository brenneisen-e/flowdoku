/**
 * Check-In Seite - QR-Code Scanner + Foto-Upload + Manuelle Eingabe
 *
 * Versucht zuerst die Kamera direkt zu nutzen (funktioniert in vielen SPFx-Umgebungen).
 * Fallback 1: Foto-Upload (User macht Foto vom QR-Code, App liest es aus).
 * Fallback 2: Manuelle Code-Eingabe.
 */

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { EventService, SPRegistration } from '../services/EventService';
import { useLanguage } from '../context/LanguageContext';
import { Html5Qrcode } from 'html5-qrcode';

export default function CheckInPage(): React.ReactElement {
  const { events } = useEvents();
  const { selectedEventId } = useNavigation();
  const { isAdmin, isOrganizer, siteUrl } = useRoles();
  const { t } = useLanguage();
  const scannerRef = React.useRef<Html5Qrcode | null>(null);
  const [manualCode, setManualCode] = React.useState('');
  const [resultMessage, setResultMessage] = React.useState('');
  const [resultType, setResultType] = React.useState<'success' | 'error' | 'info' | ''>('');
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [isScanning, setIsScanning] = React.useState(false);
  const [cameraError, setCameraError] = React.useState('');
  const [checkedInCount, setCheckedInCount] = React.useState(0);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const context = (window as any).__dexSpfxContext;
  const eventService = React.useMemo(() => context ? new EventService(context) : null, []);

  const selectedEvent = selectedEventId ? events.find(e => e.id === selectedEventId) : null;

  // Erkennen ob die App in der SharePoint Mobile App laeuft
  const isSharePointMobileApp = React.useMemo(() => {
    const ua = navigator.userAgent;
    return /SharePoint/i.test(ua) && /Mobile|Android|iPhone|iPad/i.test(ua);
  }, []);

  // URL fuer den Browser-Link generieren
  const getBrowserUrl = (): string => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    if (ctx) {
      return ctx.pageContext.site.absoluteUrl + ctx.pageContext.site.serverRelativeUrl.replace(ctx.pageContext.site.serverRelativeUrl, '') + '/SitePages/' +
        (window.location.pathname.split('/').pop() || 'Test_App.aspx');
    }
    return window.location.href;
  };

  // Scanner in neuem Fenster oeffnen (aspx-Seite in SiteAssets)
  const openExternalScanner = (): void => {
    const checkinUrl = `${siteUrl}/SiteAssets/checkin.aspx`;
    const params = new URLSearchParams();
    params.set('siteUrl', siteUrl);
    if (selectedEvent) {
      params.set('subsiteUrl', selectedEvent.subsiteUrl || '');
      params.set('eventTitle', selectedEvent.title);
      params.set('eventNumber', (selectedEvent.eventNumber || 0).toString());
    }
    window.open(`${checkinUrl}?${params.toString()}`, '_blank');
  };

  const eventByNumber = React.useMemo(() => {
    const map: Record<number, { id: string; title: string; subsiteUrl: string; eventNumber: number }> = {};
    for (const e of events) {
      if (e.eventNumber) map[e.eventNumber] = { id: e.id, title: e.title, subsiteUrl: e.subsiteUrl || '', eventNumber: e.eventNumber };
    }
    return map;
  }, [events]);

  // html5-qrcode Scanner starten
  const startCamera = async (): Promise<void> => {
    setCameraError('');
    try {
      const scanner = new Html5Qrcode('qr-reader');
      scannerRef.current = scanner;
      await scanner.start(
        { facingMode: 'environment' },
        { fps: 10, qrbox: { width: 250, height: 250 } },
        async (decodedText) => {
          // Erfolgreich gescannt — pausieren, verarbeiten, dann weiterscannen
          try {
            await scanner.pause();
          } catch { /* */ }
          await processCode(decodedText);
          // Nach 2 Sekunden weiterscannen
          setTimeout(() => {
            try { scanner.resume(); } catch { /* */ }
          }, 2000);
        },
        () => { /* ignore scan failures */ }
      );
      setIsScanning(true);
    } catch (err) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const error = err as any;
      const msg = typeof error === 'string' ? error : error?.message || 'Unbekannter Fehler';
      setCameraError(`Kamera konnte nicht gestartet werden: ${msg}`);
    }
  };

  const stopCamera = async (): Promise<void> => {
    try {
      if (scannerRef.current) {
        await scannerRef.current.stop();
        scannerRef.current = null;
      }
    } catch { /* */ }
    setIsScanning(false);
  };

  // Foto-Upload: QR-Code aus Bild lesen
  const handlePhotoUpload = async (file: File): Promise<void> => {
    try {
      const tempScanner = new Html5Qrcode('qr-reader-temp');
      const result = await tempScanner.scanFile(file, true);
      await tempScanner.clear();
      if (result) {
        await processCode(result);
      } else {
        setResultMessage('Kein QR-Code im Bild erkannt.');
        setResultType('error');
      }
    } catch {
      setResultMessage('Kein QR-Code im Bild erkannt. Bitte erneut versuchen.');
      setResultType('error');
    }
  };

  // Code verarbeiten und einchecken
  const processCode = async (code: string): Promise<void> => {
    if (isProcessing) return;
    setIsProcessing(true);
    setResultMessage('');
    setResultType('');

    const parts = code.split('|');
    if (parts.length !== 3 || parts[0] !== 'DEX') {
      setResultMessage('Ungültiger QR-Code.');
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const eventNumber = parseInt(parts[1], 10);
    const email = parts[2];
    const event = eventByNumber[eventNumber];

    if (!event || !event.subsiteUrl || !eventService) {
      setResultMessage(`Event #${eventNumber} nicht gefunden.`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const reg = await eventService.getRegistrationByEmail(event.subsiteUrl, email);
    if (!reg) {
      setResultMessage(`${email} — nicht registriert.`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;

    if (reg.Status === 'Eingecheckt') {
      setResultMessage(`${name} — ${t('checkin.alreadycheckedin')}`);
      setResultType('info');
      setIsProcessing(false);
      return;
    }
    if (reg.Status === 'Abgemeldet') {
      setResultMessage(`${name} — ${t('checkin.cancelled')}`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const success = await eventService.checkInParticipant(event.subsiteUrl, reg.Id);
    if (success) {
      setCheckedInCount(prev => prev + 1);
      setResultMessage(`${name} — ${t('checkin.success')}`);
      setResultType('success');
    } else {
      setResultMessage(`${name} — Check-in fehlgeschlagen.`);
      setResultType('error');
    }
    setIsProcessing(false);
  };

  const handleManualSubmit = async (): Promise<void> => {
    if (!manualCode.trim()) return;
    await processCode(manualCode.trim());
    setManualCode('');
  };

  React.useEffect(() => {
    return () => {
      if (scannerRef.current) {
        scannerRef.current.stop().catch(() => {});
      }
    };
  }, []);

  if (!isAdmin && !isOrganizer) {
    return (
      <div className="page-container text-center">
        <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>Admin / Organizer only.</p>
      </div>
    );
  }

  return (
    <div className="page-container">
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
        <h2 style={{ margin: 0 }}>
          {t('checkin.title')} {selectedEvent ? `— ${selectedEvent.title}` : ''}
        </h2>
        {checkedInCount > 0 && (
          <span style={{ fontSize: '0.9rem', fontWeight: 600, color: 'var(--dex-green)' }}>
            {checkedInCount} {t('checkin.sessioncount')}
          </span>
        )}
      </div>

      {/* Ergebnis-Anzeige */}
      {resultMessage && (
        <div style={{
          padding: '16px 20px', borderRadius: 12, marginBottom: 16, fontWeight: 600, fontSize: '1rem',
          background: resultType === 'success' ? '#e8f5e9' : resultType === 'error' ? '#ffebee' : '#e3f2fd',
          color: resultType === 'success' ? '#2e7d32' : resultType === 'error' ? '#c62828' : '#1565c0',
          border: resultType === 'success' ? '2px solid #86bc25' : resultType === 'error' ? '2px solid #ef5350' : '2px solid #42a5f5',
        }}>
          {resultMessage}
        </div>
      )}

      {/* SharePoint Mobile App Warnung */}
      {isSharePointMobileApp && (
        <div className="card" style={{
          padding: 20, marginBottom: 16,
          background: '#fff3e0', border: '2px solid #ff9800', borderRadius: 12,
        }}>
          <h3 style={{ margin: '0 0 8px', color: '#e65100', fontSize: '1rem' }}>
            Kamera nicht verfügbar in der SharePoint App
          </h3>
          <p style={{ color: '#bf360c', fontSize: '0.85rem', lineHeight: 1.6, margin: '0 0 12px' }}>
            Die SharePoint Mobile App unterstützt keinen Kamera-Zugriff für Webparts.
            Bitte öffne diese Seite in <strong>Edge</strong> oder <strong>Safari</strong> auf deinem Handy — dort funktioniert der QR-Scanner.
          </p>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <button
              className="btn btn-primary"
              style={{ fontSize: '0.85rem' }}
              onClick={() => {
                const url = window.location.href;
                navigator.clipboard.writeText(url).then(() => {
                  setResultMessage('Link kopiert! Öffne ihn in Edge oder Safari.');
                  setResultType('info');
                }).catch(() => {
                  window.prompt('Link kopieren und in Edge/Safari öffnen:', url);
                });
              }}
            >
              Link kopieren
            </button>
          </div>
          <p style={{ color: '#bf360c', fontSize: '0.75rem', marginTop: 8, marginBottom: 0 }}>
            Tipp: Lege dir die Seite als Lesezeichen in Edge an für schnellen Zugriff beim Event.
          </p>
        </div>
      )}

      {/* QR-Code scannen — zwei Optionen */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        <h3 style={{ marginBottom: 8 }}>QR-Code scannen</h3>
        <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.85rem', marginBottom: 16 }}>
          QR-Code fotografieren oder aus der Galerie wählen — Check-in wird sofort ausgeführt.
        </p>
        <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
          <label style={{
            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10,
            padding: '16px 24px', borderRadius: 12, border: 'none',
            background: 'var(--dex-green)', color: '#fff', cursor: 'pointer',
            fontSize: '1rem', fontWeight: 700, flex: '1 1 200px',
          }}>
            Kamera öffnen
            <input
              type="file"
              accept="image/*"
              capture="environment"
              style={{ display: 'none' }}
              onChange={e => {
                const file = e.target.files && e.target.files[0];
                if (file) handlePhotoUpload(file);
                e.target.value = '';
              }}
            />
          </label>
          <label style={{
            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10,
            padding: '16px 24px', borderRadius: 12, border: '2px solid var(--dex-gray-300)',
            background: '#fff', color: 'var(--dex-gray-700)', cursor: 'pointer',
            fontSize: '1rem', fontWeight: 600, flex: '1 1 200px',
          }}>
            Aus Galerie wählen
            <input
              type="file"
              accept="image/*"
              style={{ display: 'none' }}
              onChange={e => {
                const file = e.target.files && e.target.files[0];
                if (file) handlePhotoUpload(file);
                e.target.value = '';
              }}
            />
          </label>
        </div>
      </div>

      {/* Live-Kamera Scanner (html5-qrcode) */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        <h3 style={{ marginBottom: 12 }}>Live-Scanner</h3>
        {!isScanning ? (
          <div style={{ textAlign: 'center' }}>
            <button className="btn btn-primary" onClick={startCamera} style={{ fontSize: '1rem', padding: '12px 32px' }}>
              {t('checkin.scan')}
            </button>
            {cameraError && (
              <p style={{ color: 'var(--dex-orange)', fontSize: '0.85rem', marginTop: 12 }}>{cameraError}</p>
            )}
          </div>
        ) : (
          <div style={{ textAlign: 'center' }}>
            <p style={{ color: 'var(--dex-green)', fontWeight: 600, marginBottom: 8 }}>{t('checkin.scanning')}</p>
            <button className="btn btn-secondary" onClick={stopCamera} style={{ marginBottom: 12 }}>
              {t('general.cancel')}
            </button>
          </div>
        )}
        <div id="qr-reader" style={{ width: '100%', maxWidth: 400, margin: '0 auto' }} />
      </div>

      <div id="qr-reader-temp" style={{ display: 'none' }} />

      {/* Manuelle Eingabe */}
      <div className="card" style={{ padding: 24 }}>
        <h3 style={{ marginBottom: 12 }}>{t('checkin.manual')}</h3>
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
            {t('checkin.checkinbtn')}
          </button>
        </div>
      </div>
    </div>
  );
}
