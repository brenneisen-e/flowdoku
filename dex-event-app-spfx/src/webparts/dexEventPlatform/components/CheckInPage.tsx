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
import { EventService } from '../services/EventService';
import { useLanguage } from '../context/LanguageContext';
import QrScanner from 'qr-scanner';

export default function CheckInPage(): React.ReactElement {
  const { events } = useEvents();
  const { selectedEventId } = useNavigation();
  const { isAdmin, isOrganizer, siteUrl } = useRoles();
  const { t } = useLanguage();
  const scannerRef = React.useRef<QrScanner | null>(null);
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const [manualCode, setManualCode] = React.useState('');
  const [resultMessage, setResultMessage] = React.useState('');
  const [resultType, setResultType] = React.useState<'success' | 'error' | 'info' | ''>('');
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [isScanning, setIsScanning] = React.useState(false);
  const [cameraError, setCameraError] = React.useState('');
  const [checkedInCount, setCheckedInCount] = React.useState(0);
  const confirmCardRef = React.useRef<HTMLDivElement>(null);
  const [pendingCheckIn, setPendingCheckIn] = React.useState<{
    name: string; email: string; event: { subsiteUrl: string; title: string };
    regId: number; status: string; department?: string; jobTitle?: string; location?: string; photoUrl?: string;
  } | null>(null);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const context = (window as any).__dexSpfxContext;
  const eventService = React.useMemo(() => context ? new EventService(context) : null, []);

  const selectedEvent = selectedEventId ? events.find(e => e.id === selectedEventId) : null;

  // Erkennen ob die App in der SharePoint Mobile App laeuft
  const isSharePointMobileApp = React.useMemo(() => {
    const ua = navigator.userAgent;
    return /SharePoint/i.test(ua) && /Mobile|Android|iPhone|iPad/i.test(ua);
  }, []);

  // URL fuer den Browser-Link generieren (fuer zukuenftige Nutzung)
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const getBrowserUrl = (): string => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    if (ctx) {
      return ctx.pageContext.site.absoluteUrl + ctx.pageContext.site.serverRelativeUrl.replace(ctx.pageContext.site.serverRelativeUrl, '') + '/SitePages/' +
        (window.location.pathname.split('/').pop() || 'DEX.aspx');
    }
    return window.location.href;
  };

  // Scanner in neuem Fenster oeffnen (aspx-Seite in SiteAssets, fuer zukuenftige Nutzung)
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
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

  const lastScannedRef = React.useRef<string>('');
  const processingRef = React.useRef<boolean>(false);

  // qr-scanner starten
  const startCamera = async (): Promise<void> => {
    setCameraError('');
    if (!videoRef.current) return;

    try {
      const scanner = new QrScanner(
        videoRef.current,
        async (result) => {
          const code = result.data;
          if (!code || code === lastScannedRef.current || processingRef.current) return;
          lastScannedRef.current = code;
          processingRef.current = true;
          // Vibration bei Erkennung
          try { navigator.vibrate(200); } catch { /* */ }
          await processCode(code);
          processingRef.current = false;
          setTimeout(() => { lastScannedRef.current = ''; }, 3000);
        },
        {
          preferredCamera: 'environment',
          highlightScanRegion: true,
          highlightCodeOutline: true,
          maxScansPerSecond: 5,
        }
      );
      scannerRef.current = scanner;
      await scanner.start();
      setIsScanning(true);
    } catch (err) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const error = err as any;
      const msg = typeof error === 'string' ? error : error?.message || 'Unbekannter Fehler';
      setCameraError(`Kamera konnte nicht gestartet werden: ${msg}`);
    }
  };

  const stopCamera = (): void => {
    if (scannerRef.current) {
      scannerRef.current.stop();
      scannerRef.current.destroy();
      scannerRef.current = null;
    }
    setIsScanning(false);
  };

  // Foto-Upload: QR-Code aus Bild lesen
  const handlePhotoUpload = async (file: File): Promise<void> => {
    try {
      const result = await QrScanner.scanImage(file, { returnDetailedScanResult: true });
      if (result && result.data) {
        await processCode(result.data);
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
      setResultMessage(`Ungültiger QR-Code: "${code}"`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }

    const eventNumber = parseInt(parts[1], 10);
    const email = parts[2];

    // EventNumber oder SP-ID Lookup
    let event = eventByNumber[eventNumber];

    // Fallback: Wenn EventNumber 0 oder nicht gefunden, versuche SP-ID
    if (!event) {
      const byId = events.find(e => e.id === parts[1]);
      if (byId) {
        event = { id: byId.id, title: byId.title, subsiteUrl: byId.subsiteUrl || '', eventNumber: byId.eventNumber };
      }
    }

    if (!event || !event.subsiteUrl || !eventService) {
      setResultMessage(`Event #${eventNumber} nicht gefunden. (Code: ${code})`);
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

    // Statt sofort einzuchecken, Info-Karte anzeigen und auf Bestätigung warten
    // Profilbild laden
    let photoUrl = '';
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (ctx) {
        const siteBase = ctx.pageContext.web.absoluteUrl;
        photoUrl = `${siteBase}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;
      }
    } catch { /* */ }

    setPendingCheckIn({
      name,
      email,
      event: { subsiteUrl: event.subsiteUrl, title: event.title },
      regId: reg.Id,
      status: reg.Status,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      department: (reg as any).Department || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      jobTitle: (reg as any).JobTitle || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      location: (reg as any).Location || '',
      photoUrl,
    });
    setResultMessage('');
    setResultType('');
    setIsProcessing(false);
    // Zum Bestätigungs-Dialog scrollen
    setTimeout(() => {
      confirmCardRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 100);
  };

  // Check-in bestätigen
  const confirmCheckIn = async (): Promise<void> => {
    if (!pendingCheckIn || !eventService) return;
    try {
      await eventService.checkInParticipant(pendingCheckIn.event.subsiteUrl, pendingCheckIn.regId);
      setCheckedInCount(prev => prev + 1);
      setResultMessage(`${pendingCheckIn.name} — ${t('checkin.success')}`);
      setResultType('success');
    } catch {
      setResultMessage(`${pendingCheckIn.name} — Check-in fehlgeschlagen.`);
      setResultType('error');
    }
    setPendingCheckIn(null);
    processingRef.current = false;
  };

  const cancelCheckIn = (): void => {
    setPendingCheckIn(null);
    lastScannedRef.current = '';
    processingRef.current = false;
  };

  const handleManualSubmit = async (): Promise<void> => {
    if (!manualCode.trim()) return;
    await processCode(manualCode.trim());
    setManualCode('');
  };

  React.useEffect(() => {
    return () => {
      if (scannerRef.current) {
        scannerRef.current.stop();
        scannerRef.current.destroy();
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
    <div className="page-container" role="main">
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

      {/* Bestätigungs-Dialog nach Scan */}
      {pendingCheckIn && (
        <div ref={confirmCardRef} className="card" role="dialog" aria-modal="true" aria-label="Check-in bestätigen" style={{
          padding: 24, marginBottom: 16, border: '2px solid var(--dex-green)',
          borderRadius: 16, background: '#fff',
        }}>
          <div style={{ display: 'flex', gap: 16, alignItems: 'center', marginBottom: 16 }}>
            {pendingCheckIn.photoUrl ? (
              <img
                src={pendingCheckIn.photoUrl}
                alt={pendingCheckIn.name}
                style={{ width: 72, height: 72, borderRadius: '50%', objectFit: 'cover', flexShrink: 0 }}
                onError={(e) => { (e.target as HTMLImageElement).style.display = 'none'; }}
              />
            ) : (
              <div style={{
                width: 72, height: 72, borderRadius: '50%', flexShrink: 0,
                background: 'linear-gradient(135deg, #86bc25, #0076a8)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                color: '#fff', fontWeight: 700, fontSize: '1.4rem',
              }}>
                {pendingCheckIn.name.split(' ').map(n => n[0]).join('').substring(0, 2)}
              </div>
            )}
            <div style={{ flex: 1 }}>
              <h3 style={{ margin: '0 0 4px', fontSize: '1.2rem' }}>{pendingCheckIn.name}</h3>
              <p style={{ margin: '0 0 2px', color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>{pendingCheckIn.email}</p>
              {pendingCheckIn.jobTitle && (
                <p style={{ margin: '0 0 2px', fontSize: '0.85rem' }}>{pendingCheckIn.jobTitle}</p>
              )}
              {pendingCheckIn.department && (
                <p style={{ margin: '0 0 2px', fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>{pendingCheckIn.department}</p>
              )}
              {pendingCheckIn.location && (
                <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>{pendingCheckIn.location}</p>
              )}
            </div>
            <span style={{
              padding: '6px 14px', borderRadius: 20, fontSize: '0.8rem', fontWeight: 700,
              background: pendingCheckIn.status === 'Eingecheckt' ? '#e8f5e9' : '#e3f2fd',
              color: pendingCheckIn.status === 'Eingecheckt' ? '#2e7d32' : '#1565c0',
            }}>
              {pendingCheckIn.status}
            </span>
          </div>
          <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', margin: '0 0 16px' }}>
            Event: <strong>{pendingCheckIn.event.title}</strong>
          </p>
          <div style={{ display: 'flex', gap: 12 }}>
            <button
              className="btn btn-primary"
              onClick={confirmCheckIn}
              style={{ flex: 1, fontSize: '1rem', padding: '12px 0', background: 'var(--dex-green)' }}
            >
              Einchecken bestätigen
            </button>
            <button
              className="btn btn-secondary"
              onClick={cancelCheckIn}
              style={{ padding: '12px 20px' }}
            >
              Abbrechen
            </button>
          </div>
        </div>
      )}

      {/* Live-Scanner — Kamerabild + Steuerung */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        {!isScanning ? (
          <>
            <h3 style={{ marginBottom: 12 }}>Live-Scanner</h3>
            <div style={{ textAlign: 'center' }}>
              <button className="btn btn-primary" onClick={startCamera} style={{ fontSize: '1.1rem', padding: '14px 36px' }}>
                {t('checkin.scan')}
              </button>
              {cameraError && (
                <p style={{ color: 'var(--dex-orange)', fontSize: '0.85rem', marginTop: 12 }}>{cameraError}</p>
              )}
            </div>
          </>
        ) : (
          <>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
              <h3 style={{ margin: 0, color: 'var(--dex-green)' }}>{t('checkin.scanning')}</h3>
              <button className="btn btn-secondary" onClick={stopCamera} style={{ fontSize: '0.85rem' }}>
                Scanner stoppen
              </button>
            </div>
          </>
        )}
        <div style={{ position: 'relative', width: '100%', maxWidth: 500, margin: '0 auto', overflow: 'hidden', borderRadius: 12 }}>
          <video
            ref={videoRef}
            style={{
              width: '100%', height: isScanning ? 400 : 0,
              display: 'block',
              border: isScanning ? '3px solid var(--dex-green)' : 'none',
              objectFit: 'cover', background: '#000',
              transition: 'height 0.3s ease',
              borderRadius: 12,
            }}
            playsInline
            muted
          />
        </div>
      </div>

      {/* Foto-Upload — nur sichtbar wenn Live-Scanner NICHT aktiv */}
      {!isScanning && (
        <div className="card" style={{ padding: 24, marginBottom: 16 }}>
          <h3 style={{ marginBottom: 8 }}>QR-Code scannen</h3>
          <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.85rem', marginBottom: 16 }}>
            Alternativ: QR-Code fotografieren oder aus der Galerie wählen.
          </p>
          <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
            <label style={{
              display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10,
              padding: '14px 24px', borderRadius: 12, border: 'none',
              background: 'var(--dex-green)', color: '#fff', cursor: 'pointer',
              fontSize: '0.95rem', fontWeight: 700, flex: '1 1 180px',
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
              padding: '14px 24px', borderRadius: 12, border: '2px solid var(--dex-gray-300)',
              background: '#fff', color: 'var(--dex-gray-700)', cursor: 'pointer',
              fontSize: '0.95rem', fontWeight: 600, flex: '1 1 180px',
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
      )}

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
