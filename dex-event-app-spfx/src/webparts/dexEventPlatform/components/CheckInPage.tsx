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

export default function CheckInPage(): React.ReactElement {
  const { events } = useEvents();
  const { selectedEventId } = useNavigation();
  const { isAdmin, isOrganizer, siteUrl } = useRoles();
  const { t } = useLanguage();
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const canvasRef = React.useRef<HTMLCanvasElement>(null);
  const [manualCode, setManualCode] = React.useState('');
  const [resultMessage, setResultMessage] = React.useState('');
  const [resultType, setResultType] = React.useState<'success' | 'error' | 'info' | ''>('');
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [isScanning, setIsScanning] = React.useState(false);
  const [cameraError, setCameraError] = React.useState('');
  const [checkedInCount, setCheckedInCount] = React.useState(0);
  const scanIntervalRef = React.useRef<number | null>(null);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const context = (window as any).__dexSpfxContext;
  const eventService = React.useMemo(() => context ? new EventService(context) : null, []);

  const selectedEvent = selectedEventId ? events.find(e => e.id === selectedEventId) : null;

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

  // Kamera starten
  const startCamera = async (): Promise<void> => {
    setCameraError('');
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      setCameraError('Kamera-API nicht verfügbar.');
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
      setCameraError(error?.name === 'NotAllowedError'
        ? 'Kamera-Berechtigung verweigert. Bitte Foto-Upload oder manuelle Eingabe nutzen.'
        : `Kamera nicht verfügbar: ${error?.message || error?.name}. Bitte Foto-Upload nutzen.`);
    }
  };

  const stopCamera = (): void => {
    if (scanIntervalRef.current) { clearInterval(scanIntervalRef.current); scanIntervalRef.current = null; }
    if (videoRef.current && videoRef.current.srcObject) {
      (videoRef.current.srcObject as MediaStream).getTracks().forEach(t => t.stop());
      videoRef.current.srcObject = null;
    }
    setIsScanning(false);
  };

  const startDetection = (): void => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const BarcodeDetectorClass = (window as any).BarcodeDetector;
    if (!BarcodeDetectorClass) {
      setCameraError('BarcodeDetector nicht unterstützt. Bitte Chrome/Edge verwenden oder Foto-Upload nutzen.');
      stopCamera();
      return;
    }
    const detector = new BarcodeDetectorClass({ formats: ['qr_code'] });
    scanIntervalRef.current = window.setInterval(async () => {
      if (isProcessing || !videoRef.current || videoRef.current.readyState !== 4) return;
      try {
        const barcodes = await detector.detect(videoRef.current);
        if (barcodes.length > 0 && barcodes[0].rawValue) {
          await processCode(barcodes[0].rawValue);
        }
      } catch { /* */ }
    }, 500);
  };

  // Foto-Upload: QR-Code aus Bild lesen
  const handlePhotoUpload = async (file: File): Promise<void> => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const BarcodeDetectorClass = (window as any).BarcodeDetector;
    if (!BarcodeDetectorClass) {
      setResultMessage('BarcodeDetector nicht unterstützt. Bitte Code manuell eingeben.');
      setResultType('error');
      return;
    }
    const detector = new BarcodeDetectorClass({ formats: ['qr_code'] });
    const img = new Image();
    img.onload = async () => {
      try {
        const barcodes = await detector.detect(img);
        if (barcodes.length > 0 && barcodes[0].rawValue) {
          await processCode(barcodes[0].rawValue);
        } else {
          setResultMessage('Kein QR-Code im Bild erkannt. Bitte erneut versuchen.');
          setResultType('error');
        }
      } catch {
        setResultMessage('Fehler beim Lesen des QR-Codes.');
        setResultType('error');
      }
      URL.revokeObjectURL(img.src);
    };
    img.src = URL.createObjectURL(file);
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
      if (scanIntervalRef.current) clearInterval(scanIntervalRef.current);
      if (videoRef.current && videoRef.current.srcObject) {
        (videoRef.current.srcObject as MediaStream).getTracks().forEach(t => t.stop());
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

      {/* Foto-Upload (funktioniert ueberall — primaere Methode auf Mobile) */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        <h3 style={{ marginBottom: 8 }}>QR-Code scannen</h3>
        <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.85rem', marginBottom: 16 }}>
          Kamera öffnet sich automatisch — QR-Code fotografieren und Check-in wird sofort ausgeführt.
        </p>
        <label style={{
          display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10,
          padding: '16px 24px', borderRadius: 12, border: 'none',
          background: 'var(--dex-green)', color: '#fff', cursor: 'pointer',
          fontSize: '1.1rem', fontWeight: 700, width: '100%',
        }}>
          QR-Code fotografieren
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
      </div>

      {/* Kamera-Scanner (optional — funktioniert nicht in allen SPFx-Umgebungen) */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        <h3 style={{ marginBottom: 12 }}>Live-Kamera (optional)</h3>
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
            <video ref={videoRef} playsInline muted style={{
              width: '100%', maxWidth: 400, borderRadius: 12, border: '3px solid var(--dex-green)',
            }} />
            <canvas ref={canvasRef} style={{ display: 'none' }} />
            <p style={{ color: 'var(--dex-green)', fontWeight: 600, marginTop: 8 }}>{t('checkin.scanning')}</p>
            <button className="btn btn-secondary" onClick={stopCamera} style={{ marginTop: 8 }}>
              {t('general.cancel')}
            </button>
          </div>
        )}
      </div>

      {/* Foto-Upload */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        <h3 style={{ marginBottom: 12 }}>Foto-Upload</h3>
        <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.85rem', marginBottom: 12 }}>
          {t('checkin.photodesc') || 'QR-Code fotografieren oder aus Galerie wählen.'}
        </p>
        <label style={{
          display: 'inline-flex', alignItems: 'center', gap: 8,
          padding: '10px 20px', borderRadius: 8, border: '2px dashed var(--dex-gray-300)',
          cursor: 'pointer', fontSize: '0.9rem', color: 'var(--dex-gray-600)',
        }}>
          Foto auswählen
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
      </div>

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
