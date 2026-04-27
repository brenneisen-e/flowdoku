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
  const { selectedEventId, navigate } = useNavigation();
  const { isAdmin, isOrganizer, siteUrl } = useRoles();
  const { t } = useLanguage();
  // v6.22: aktueller User-E-Mail aus SPFx-Kontext, brauchen wir für Organizer-/
  // QR-Scanner-Zuordnung (statt nur tenant-weite isOrganizer/isAdmin-Flags).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const spCtx = (window as any).__dexSpfxContext;
  const currentEmailLc: string = (spCtx?.pageContext?.user?.email || '').toLowerCase();
  // Events, die der aktuelle User einchecken darf: Admin = alle, sonst nur die
  // Events in denen er Organizer oder QR-Code-Scanner ist (per E-Mail-Match).
  const accessibleEvents = React.useMemo(() => {
    return (events || []).filter(e => {
      if (isAdmin) return true;
      const orgMatch = (e.organizerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
      const qrMatch = (e.qrScannerEmails || []).some(x => (x || '').toLowerCase() === currentEmailLc);
      return orgMatch || qrMatch;
    });
  }, [events, isAdmin, currentEmailLc]);
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

  // v7.12: Name-Suche fuer manuelles Einchecken — wenn der QR-Scanner in der
  // SP-App nicht funktioniert (Camera-API gesperrt) oder der Teilnehmer den
  // QR-Code nicht zur Hand hat, kann der Helfer nach Namen / E-Mail suchen
  // und per Tap "Einchecken" ausloesen. Die Registrierungen werden pro Event
  // lazy nachgeladen und in `searchRegsCache` zwischengespeichert.
  const [nameSearchQuery, setNameSearchQuery] = React.useState('');
  const [nameSearchEventId, setNameSearchEventId] = React.useState<string>(selectedEventId || '');
  const [searchRegsCache, setSearchRegsCache] = React.useState<Record<string, import('../services/EventService').SPRegistration[]>>({});
  const [isLoadingSearchRegs, setIsLoadingSearchRegs] = React.useState(false);
  const [searchLoadError, setSearchLoadError] = React.useState('');
  React.useEffect(() => {
    if (selectedEventId && !nameSearchEventId) setNameSearchEventId(selectedEventId);
  }, [selectedEventId]); // eslint-disable-line react-hooks/exhaustive-deps

  const loadRegsForSearch = React.useCallback(async (eventId: string): Promise<void> => {
    if (!eventId || !eventService) return;
    if (searchRegsCache[eventId]) return; // bereits geladen
    const ev = events.find(e => e.id === eventId);
    if (!ev || !ev.subsiteUrl) {
      setSearchLoadError('Event nicht gefunden oder ohne Subsite.');
      return;
    }
    setIsLoadingSearchRegs(true);
    setSearchLoadError('');
    try {
      const regs = await eventService.getAllRegistrations(ev.subsiteUrl);
      setSearchRegsCache(prev => ({ ...prev, [eventId]: regs }));
    } catch {
      setSearchLoadError('Teilnehmerliste konnte nicht geladen werden.');
    }
    setIsLoadingSearchRegs(false);
  }, [eventService, events, searchRegsCache]);

  const onSearchFocus = (): void => {
    if (nameSearchEventId) loadRegsForSearch(nameSearchEventId);
  };
  React.useEffect(() => {
    if (nameSearchQuery && nameSearchEventId && !searchRegsCache[nameSearchEventId]) {
      loadRegsForSearch(nameSearchEventId);
    }
  }, [nameSearchQuery, nameSearchEventId, searchRegsCache, loadRegsForSearch]);

  const searchHits = React.useMemo(() => {
    const q = nameSearchQuery.trim().toLowerCase();
    if (q.length < 2 || !nameSearchEventId) return [];
    const regs = searchRegsCache[nameSearchEventId] || [];
    const matches = regs.filter(r => {
      const full = `${r.Vorname || ''} ${r.Nachname || ''} ${r.ParticipantName || ''} ${r.ParticipantEmail || ''}`.toLowerCase();
      return full.indexOf(q) >= 0;
    });
    return matches.slice(0, 8);
  }, [nameSearchQuery, nameSearchEventId, searchRegsCache]);

  const startManualCheckInFromSearch = (reg: import('../services/EventService').SPRegistration): void => {
    const ev = events.find(e => e.id === nameSearchEventId);
    if (!ev || !ev.subsiteUrl) return;
    if (reg.Status === 'Abgemeldet') {
      setResultMessage(`${reg.ParticipantName || reg.ParticipantEmail} — ${t('checkin.cancelled')}`);
      setResultType('error');
      return;
    }
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || reg.ParticipantEmail);
    let photoUrl = '';
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (ctx) {
        const siteBase = ctx.pageContext.web.absoluteUrl;
        photoUrl = `${siteBase}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(reg.ParticipantEmail || '')}`;
      }
    } catch { /* */ }
    setPendingCheckIn({
      name,
      email: reg.ParticipantEmail || '',
      event: { subsiteUrl: ev.subsiteUrl, title: ev.title },
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
    setNameSearchQuery('');
    setTimeout(() => {
      confirmCardRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 100);
  };

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

    // 1. Browser-API-Check: manche Embedded-WebViews (SharePoint-Mobile-App)
    //    stellen mediaDevices gar nicht bereit. Dort gleich mit klarer Meldung
    //    abbrechen statt auf qr-scanner-Fehler zu warten.
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      setCameraError(
        'Dein Browser stellt keinen Kamera-Zugriff bereit. '
        + 'Bitte öffne diese Seite direkt in Edge oder Safari (nicht in der SharePoint-App / Teams).'
      );
      return;
    }

    // 2. Explizit Berechtigung anfragen (triggert Permission-Prompt). Damit
    //    bekommen wir sauber unterscheidbare Fehler statt einer generischen
    //    qr-scanner-Exception. Das Test-Stream wird danach sofort geschlossen
    //    und qr-scanner startet seinen eigenen Stream.
    try {
      const testStream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: 'environment' } },
        audio: false,
      });
      // Test-Stream sofort stoppen, damit qr-scanner seinen eigenen aufbauen kann
      testStream.getTracks().forEach(track => track.stop());
    } catch (err) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const e = err as any;
      const name = e?.name || '';
      let msg: string;
      if (name === 'NotAllowedError' || name === 'PermissionDeniedError') {
        msg = 'Kamera-Berechtigung wurde abgelehnt. Bitte in den Browser-Einstellungen '
          + 'für diese Seite die Kamera erlauben und dann erneut versuchen. '
          + '(iOS Safari: aA-Icon links in der Adresszeile → Website-Einstellungen → Kamera: Erlauben. '
          + 'Android Chrome: Schloss-Icon in der Adresszeile → Berechtigungen → Kamera: Zulassen.)';
      } else if (name === 'NotFoundError' || name === 'DevicesNotFoundError') {
        msg = 'Keine Kamera gefunden. Stelle sicher, dass dein Gerät eine Kamera hat und kein anderes Programm sie blockiert.';
      } else if (name === 'NotReadableError' || name === 'TrackStartError') {
        msg = 'Die Kamera ist bereits in Benutzung (z.B. Teams-Anruf oder eine andere App). Bitte schließe andere Apps und versuche es erneut.';
      } else if (name === 'OverconstrainedError' || name === 'ConstraintNotSatisfiedError') {
        // Kein Environment-Facing-Camera verfuegbar -> Fallback auf beliebige Kamera
        try {
          const fallback = await navigator.mediaDevices.getUserMedia({ video: true, audio: false });
          fallback.getTracks().forEach(track => track.stop());
        } catch {
          setCameraError('Keine passende Kamera gefunden.');
          return;
        }
        // Fallback OK -> weiter mit qr-scanner-Start (ohne early return)
        msg = '';
      } else if (name === 'SecurityError') {
        msg = 'Kamera-Zugriff vom Browser blockiert (vermutlich unsichere Verbindung oder eingebetteter iframe). Öffne die Seite direkt in Edge/Safari.';
      } else {
        msg = `Kamera konnte nicht gestartet werden: ${e?.message || String(err) || 'Unbekannter Fehler'}`;
      }
      if (msg) {
        setCameraError(msg);
        return;
      }
    }

    // 3. qr-scanner starten (nutzt jetzt die bereits erteilte Permission)
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
      setCameraError(`Scanner konnte nicht gestartet werden: ${msg}`);
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

  // v6.22: Zugriffskontrolle — wer kein zugängliches Event hat UND nicht Admin ist,
  // kommt gar nicht erst in die Scanner-Maske. Deckt User ohne Rolle ab UND auch
  // Organizer/Admin ohne eigene Events.
  if (!isAdmin && !isOrganizer && accessibleEvents.length === 0) {
    return (
      <div className="page-container text-center">
        <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>
          {t('checkin.noaccess') || 'Du hast keinen Zugriff auf den Check-In. Wende dich an einen Admin, wenn du als Organizer oder QR-Scanner eingetragen werden solltest.'}
        </p>
        <button className="btn btn-secondary" onClick={() => navigate('landing')}>{t('reg.backtoevents') || 'Zurück'}</button>
      </div>
    );
  }

  // Kein Event ausgewählt → Event-Picker (nur relevante Events: nur die, die
  // der User einchecken darf; bei exakt einem Event wird automatisch weiter
  // navigiert, weil die LandingPage das schon macht. Wir fangen hier aber
  // trotzdem den Fall ab, wenn jemand direkt via Header-Button ohne Eventauswahl
  // herkommt.)
  if (!selectedEvent) {
    if (accessibleEvents.length === 1) {
      // Auto-select: direkt weiterleiten statt Liste mit einem Eintrag zeigen.
      navigate('check-in', accessibleEvents[0].id);
      return (
        <div className="page-container text-center">
          <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>…</p>
        </div>
      );
    }
    return (
      <div className="page-container" role="main">
        <h2 className="mb-16">{t('checkin.title') || 'Check-In'}</h2>
        <p style={{ color: 'var(--dex-gray-600)', marginBottom: 16, fontSize: '0.9rem' }}>
          {t('checkin.pickevent') || 'Wähle das Event, für das du eincheckst:'}
        </p>
        {accessibleEvents.length === 0 ? (
          <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
            {t('checkin.noevents') || 'Keine Events verfügbar.'}
          </p>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            {accessibleEvents.map(ev => (
              <button
                key={ev.id}
                className="card"
                onClick={() => navigate('check-in', ev.id)}
                style={{
                  textAlign: 'left', padding: 16,
                  background: '#fff', border: '1px solid var(--dex-gray-200)',
                  borderRadius: 12, cursor: 'pointer',
                  display: 'flex', flexDirection: 'column', gap: 4,
                }}
              >
                <div style={{ fontWeight: 600, fontSize: '0.95rem' }}>{ev.title}</div>
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                  {ev.startDate ? new Date(ev.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : ''}
                  {ev.location ? ` · ${ev.location}` : ''}
                </div>
              </button>
            ))}
          </div>
        )}
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
              Einchecken
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
        {/* Wichtig: Video nicht mit fixer Hoehe + objectFit:cover croppen, sonst landet
            die vom qr-scanner eingeblendete Scan-Region-Box verschoben, weil die
            Library Overlay-Koordinaten aus dem nativen Video-Aspect berechnet.
            Video soll seine natuerliche Aspect Ratio behalten (width:100%, height:auto).
            Vor dem Scan-Start verstecken wir den Container via max-height:0 - so
            bleibt der Video-Ref stabil und die Library kann nach getUserMedia die
            Dimensionen korrekt bestimmen. */}
        <div style={{
          position: 'relative', width: '100%', maxWidth: 500, margin: '0 auto',
          overflow: 'hidden', borderRadius: 12,
          maxHeight: isScanning ? '80vh' : 0,
          border: isScanning ? '3px solid var(--dex-green)' : 'none',
          background: '#000',
          transition: 'max-height 0.3s ease',
        }}>
          <video
            ref={videoRef}
            style={{
              width: '100%', height: 'auto', display: 'block',
              background: '#000',
              borderRadius: 9,
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

      {/* v7.12: Name-Suche + Manuell einchecken — als Fallback wenn kein QR
          zur Hand ist oder die Camera-API in der SP-App blockiert ist. */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        <h3 style={{ marginBottom: 12 }}>Nach Name suchen</h3>
        {accessibleEvents.length > 1 && (
          <select
            className="form-input"
            value={nameSearchEventId}
            onChange={e => { setNameSearchEventId(e.target.value); setNameSearchQuery(''); }}
            style={{ marginBottom: 10, padding: '8px 12px', fontSize: '0.9rem', width: '100%' }}
          >
            <option value="">— Event auswählen —</option>
            {accessibleEvents.map(ev => (
              <option key={ev.id} value={ev.id}>{ev.title}</option>
            ))}
          </select>
        )}
        <input
          className="form-input"
          value={nameSearchQuery}
          onFocus={onSearchFocus}
          onChange={e => setNameSearchQuery(e.target.value)}
          placeholder="Vorname, Nachname oder E-Mail eingeben…"
          disabled={!nameSearchEventId}
          style={{ width: '100%', padding: '10px 14px', fontSize: '0.95rem' }}
        />
        {!nameSearchEventId && accessibleEvents.length > 1 && (
          <p style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
            Bitte zuerst ein Event wählen.
          </p>
        )}
        {isLoadingSearchRegs && (
          <p style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
            Teilnehmerliste wird geladen…
          </p>
        )}
        {searchLoadError && (
          <p style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-red)' }}>
            {searchLoadError}
          </p>
        )}
        {nameSearchQuery.trim().length >= 2 && (
          <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
            {searchHits.length === 0 && !isLoadingSearchRegs && (
              <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)', fontStyle: 'italic', margin: 0 }}>
                Kein Treffer — bitte anders schreiben oder per QR-Code einchecken.
              </p>
            )}
            {searchHits.map(reg => {
              const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || reg.ParticipantEmail || '-');
              const status = reg.Status;
              const alreadyIn = status === 'Eingecheckt';
              const cancelled = status === 'Abgemeldet';
              const waitlist = status === 'Warteliste';
              const statusBg = alreadyIn ? 'rgba(134,188,37,0.15)'
                : cancelled ? 'rgba(204,0,0,0.10)'
                : waitlist ? 'rgba(237,139,0,0.12)'
                : 'rgba(21,101,192,0.10)';
              const statusFg = alreadyIn ? 'var(--dex-green-dark, #4a7c1f)'
                : cancelled ? 'var(--dex-red, #c00)'
                : waitlist ? 'var(--dex-orange, #ed8b00)'
                : '#1565c0';
              return (
                <div
                  key={reg.Id}
                  style={{
                    display: 'flex', alignItems: 'center', gap: 10,
                    padding: '10px 12px', border: '1px solid var(--dex-gray-200)',
                    borderRadius: 10, background: '#fff',
                  }}
                >
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: '0.9rem', color: 'var(--dex-gray-800)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                      {name}
                    </div>
                    <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                      {reg.ParticipantEmail}
                    </div>
                  </div>
                  <span style={{
                    fontSize: '0.7rem', padding: '3px 8px', borderRadius: 999,
                    background: statusBg, color: statusFg, fontWeight: 600, whiteSpace: 'nowrap',
                  }}>{status}</span>
                  <button
                    type="button"
                    className="btn btn-primary"
                    style={{ fontSize: '0.78rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
                    disabled={alreadyIn || cancelled || isProcessing}
                    onClick={() => startManualCheckInFromSearch(reg)}
                  >
                    {alreadyIn ? '✓ Eingecheckt' : cancelled ? 'Abgemeldet' : 'Einchecken'}
                  </button>
                </div>
              );
            })}
          </div>
        )}
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
