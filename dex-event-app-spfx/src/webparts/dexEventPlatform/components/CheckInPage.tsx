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
// v20.1: Self-Check-in direkt aus der Check-in-Seite (Live-QR + PDF-Druck).
import { generateSelfCheckInToken } from '../utils/selfCheckIn';
import { downloadSelfCheckInPdf } from '../utils/selfCheckInPdf';
// v20.4: modernes Alert-Modal statt window.alert.
import { useDialog } from '../context/DialogContext';
import { useRoles } from '../context/RoleContext';
import { useCurrentUser } from '../context/UserContext';
import { EventService } from '../services/EventService';
import { useLanguage } from '../context/LanguageContext';
import OrganizerList from './OrganizerList';
// v20.0 (Audit): qr-scanner nur noch als Typ statisch importieren — die
// eigentliche Bibliothek wird erst beim Kamera-Start dynamisch nachgeladen.
import type QrScanner from 'qr-scanner';

export default function CheckInPage(): React.ReactElement {
  const { events, getAllRegistrations, updateEvent } = useEvents();
  // v20.4: App-Modal statt nativem Browser-Alert.
  const { showAlert, confirmDialog } = useDialog();
  const { selectedEventId, navigate } = useNavigation();
  const { isAdmin, isOrganizer, siteUrl } = useRoles();
  const { currentUser } = useCurrentUser();
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  // v6.22 / v13.11: aktueller User-E-Mail über UserContext — der respektiert
  // die Demo-Impersonation (sonst greift hier immer die echte SPFx-Identität
  // des Admins, und Demo-Modus „Check-In-Team" käme nie an die Check-In-
  // Seite ran, weil die qrScannerEmails-Injektion auf die Demo-Mail zeigt).
  const currentEmailLc: string = (currentUser.email || '').toLowerCase();
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
  // v15.3: Picker zeigt per Default nur aktive Events (gleicher Filter wie
  // EventListPage). Toggle „Nur aktive" oben rechts, der das aufweicht.
  const [onlyActiveCheckIn, setOnlyActiveCheckIn] = React.useState<boolean>(true);
  const visibleCheckInEvents = React.useMemo(() => {
    if (!onlyActiveCheckIn) return accessibleEvents;
    const now = Date.now();
    return accessibleEvents.filter(e => {
      if (e.status !== 'Active') return false;
      if (e.isFictive) return false;
      const activeFromTs = e.activeFrom ? new Date(e.activeFrom).getTime() : 0;
      if (activeFromTs > 0 && activeFromTs > now) return false;
      // Vergangene Events ebenfalls ausblenden (EndDate < heute - 1 Tag)
      const endTs = e.endDate ? new Date(e.endDate).getTime() : 0;
      if (endTs > 0 && endTs < now - 24 * 60 * 60 * 1000) return false;
      return true;
    });
  }, [accessibleEvents, onlyActiveCheckIn]);
  const scannerRef = React.useRef<QrScanner | null>(null);
  const videoRef = React.useRef<HTMLVideoElement>(null);
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

  // v7.12: Name-Suche für manuelles Einchecken — wenn der QR-Scanner in der
  // SP-App nicht funktioniert (Camera-API gesperrt) oder der Teilnehmer den
  // QR-Code nicht zur Hand hat, kann der Helfer nach Namen / E-Mail suchen
  // und per Tap "Einchecken" auslösen. Die Registrierungen werden pro Event
  // lazy nachgeladen und in `searchRegsCache` zwischengespeichert.
  const [nameSearchQuery, setNameSearchQuery] = React.useState('');
  const [nameSearchEventId, setNameSearchEventId] = React.useState<string>(selectedEventId || '');
  const [searchRegsCache, setSearchRegsCache] = React.useState<Record<string, import('../services/EventService').SPRegistration[]>>({});
  const [isLoadingSearchRegs, setIsLoadingSearchRegs] = React.useState(false);
  const [searchLoadError, setSearchLoadError] = React.useState('');
  // v20.1: Busy-Flag für die Self-Check-in-Aktionen (Live-QR / PDF).
  const [selfCheckInBusy, setSelfCheckInBusy] = React.useState(false);
  React.useEffect(() => {
    if (selectedEventId && !nameSearchEventId) setNameSearchEventId(selectedEventId);
  }, [selectedEventId]); // eslint-disable-line react-hooks/exhaustive-deps

  const loadRegsForSearch = React.useCallback(async (eventId: string): Promise<void> => {
    if (!eventId) return;
    if (searchRegsCache[eventId]) return; // bereits geladen
    setIsLoadingSearchRegs(true);
    setSearchLoadError('');
    try {
      const regs = await getAllRegistrations(eventId);
      setSearchRegsCache(prev => ({ ...prev, [eventId]: regs }));
    } catch {
      setSearchLoadError('Teilnehmerliste konnte nicht geladen werden.');
    }
    setIsLoadingSearchRegs(false);
  }, [getAllRegistrations, searchRegsCache]);

  // v7.14: Sobald nameSearchEventId gesetzt ist, Teilnehmerliste vorab laden,
  // damit die Live-Liste ohne Vorab-Tippen sichtbar ist.
  React.useEffect(() => {
    if (nameSearchEventId && !searchRegsCache[nameSearchEventId]) {
      loadRegsForSearch(nameSearchEventId);
    }
  }, [nameSearchEventId, searchRegsCache, loadRegsForSearch]);

  // v7.16: Quick-Filter "Nur offene" — wenn aktiv, werden Eingecheckte,
  // Wartelistler und Abgemeldete aus der Liste ausgeblendet, sodass der
  // Helfer am Eingang nur die noch offenen Anmeldungen sieht.
  const [onlyOpen, setOnlyOpen] = React.useState(false);

  // KPI-Zähler: angemeldet (= Status Angemeldet/QR versendet/Eingecheckt)
  // und eingecheckt. Wird im KPI-Bereich der Check-In-Page angezeigt.
  const checkInKpis = React.useMemo(() => {
    if (!nameSearchEventId) return { registered: 0, checkedIn: 0, noShow: 0 };
    const regs = searchRegsCache[nameSearchEventId] || [];
    let registered = 0;
    let checkedIn = 0;
    let noShow = 0;
    for (const r of regs) {
      if (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt') registered++;
      if (r.Status === 'Eingecheckt') checkedIn++;
      if (r.Status === 'No-Show') noShow++;
    }
    return { registered, checkedIn, noShow };
  }, [nameSearchEventId, searchRegsCache]);

  // v7.14: Live-Filter ueber die ganze Liste — leerer Query zeigt alle
  // Teilnehmer. Sortiert nach Status (Aktive zuerst), dann Nachname.
  const searchHits = React.useMemo(() => {
    if (!nameSearchEventId) return [];
    const regs = searchRegsCache[nameSearchEventId] || [];
    const q = nameSearchQuery.trim().toLowerCase();
    const matchesQuery = q.length === 0
      ? regs
      : regs.filter(r => {
          const full = `${r.Vorname || ''} ${r.Nachname || ''} ${r.ParticipantName || ''} ${r.ParticipantEmail || ''}`.toLowerCase();
          return full.indexOf(q) >= 0;
        });
    // v7.16: optionaler Quick-Filter "nur offene Anmeldungen"
    const filtered = onlyOpen
      ? matchesQuery.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet')
      : matchesQuery;
    // Sortierung: Angemeldet/QR versendet zuerst, dann Eingecheckt, dann
    // Warteliste, dann Abgemeldet. Innerhalb der Gruppe alphabetisch nach
    // Nachname.
    const statusRank = (s: string): number => {
      if (s === 'Angemeldet' || s === 'QR versendet') return 0;
      if (s === 'Eingecheckt') return 1;
      if (s === 'Warteliste') return 2;
      return 3; // Abgemeldet & Sonstige
    };
    return filtered.slice().sort((a, b) => {
      const sa = statusRank(a.Status);
      const sb = statusRank(b.Status);
      if (sa !== sb) return sa - sb;
      const na = (a.Nachname || a.ParticipantName || '').toLowerCase();
      const nb = (b.Nachname || b.ParticipantName || '').toLowerCase();
      return na.localeCompare(nb);
    });
  }, [nameSearchQuery, nameSearchEventId, searchRegsCache, onlyOpen]);

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

  // v23.28: Teilnehmer als „No-Show" markieren (nicht erschienen). Direkt aus
  // der Suchliste; nach Bestätigung wird der lokale Cache aktualisiert.
  const markNoShowFromSearch = async (reg: import('../services/EventService').SPRegistration): Promise<void> => {
    const ev = events.find(e => e.id === nameSearchEventId);
    if (!ev || !ev.subsiteUrl || !eventService) return;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || reg.ParticipantEmail || '-');
    const ok = await confirmDialog(
      isDe
        ? `„${name}" als nicht erschienen (No-Show) markieren?`
        : `Mark „${name}" as a no-show?`,
      { confirmLabel: isDe ? 'Als No-Show markieren' : 'Mark as no-show' }
    );
    if (!ok) return;
    const success = await eventService.markNoShowParticipant(ev.subsiteUrl, reg.Id);
    if (success) {
      setSearchRegsCache(prev => {
        const list = prev[nameSearchEventId] || [];
        return { ...prev, [nameSearchEventId]: list.map(r => r.Id === reg.Id ? { ...r, Status: 'No-Show' } : r) };
      });
      setResultMessage(isDe ? `${name} — als No-Show markiert.` : `${name} — marked as no-show.`);
      setResultType('info');
    } else {
      // v23.29: No-Show gibt es nur für Events, die ab v23.28 NEU angelegt
      // wurden — Bestands-Events kennen die Choice nicht (HTTP 400).
      setResultMessage(isDe
        ? `${name} — No-Show ist für dieses Event nicht verfügbar (nur für neu angelegte Events).`
        : `${name} — no-show is not available for this event (only for newly created events).`);
      setResultType('error');
    }
  };

  // Erkennen ob die App in der SharePoint Mobile App läuft
  const isSharePointMobileApp = React.useMemo(() => {
    const ua = navigator.userAgent;
    return /SharePoint/i.test(ua) && /Mobile|Android|iPhone|iPad/i.test(ua);
  }, []);

  // URL für den Browser-Link generieren (für zukünftige Nutzung)
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

  // Scanner in neuem Fenster oeffnen (aspx-Seite in SiteAssets, für zukünftige Nutzung)
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
        // Kein Environment-Facing-Camera verfügbar -> Fallback auf beliebige Kamera
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
      // v20.0: Bibliothek lazy laden (eigener Chunk, nur bei Kamera-Start).
      const QrScannerCls = (await import('qr-scanner')).default;
      const scanner = new QrScannerCls(
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
      <div className="page-container" role="main" style={{ maxWidth: 1100, marginLeft: 'auto', marginRight: 'auto' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 12, marginBottom: 8 }}>
          <h2 style={{ margin: 0 }}>{t('checkin.title') || 'Check-In'}</h2>
          {/* v15.3: Toggle „Nur aktive" — identisches Pattern wie EventListPage. */}
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
            <input
              type="checkbox"
              checked={onlyActiveCheckIn}
              onChange={e => setOnlyActiveCheckIn(e.target.checked)}
              style={{ width: 16, height: 16, accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
            />
            <span>{isDe ? 'Nur aktive Events' : 'Active events only'}</span>
          </label>
        </div>
        <p style={{ color: 'var(--dex-gray-600)', marginBottom: 16, fontSize: '0.9rem' }}>
          {t('checkin.pickevent') || 'Wähle das Event, für das du eincheckst:'}
        </p>
        {visibleCheckInEvents.length === 0 ? (
          <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>
            {onlyActiveCheckIn
              ? (isDe ? 'Keine aktiven Events. Toggle oben deaktivieren, um auch ältere zu sehen.' : 'No active events. Untoggle „Active events only" to see older ones.')
              : (t('checkin.noevents') || 'Keine Events verfügbar.')}
          </p>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            {visibleCheckInEvents.map(ev => (
              <button
                key={ev.id}
                className="card"
                onClick={() => navigate('check-in', ev.id)}
                style={{
                  textAlign: 'left', padding: 14,
                  background: '#fff', border: '1px solid var(--dex-gray-200)',
                  borderRadius: 12, cursor: 'pointer',
                  display: 'flex', alignItems: 'center', gap: 14,
                  transition: 'border-color 0.15s, box-shadow 0.15s, transform 0.1s',
                }}
                onMouseEnter={e => { e.currentTarget.style.borderColor = 'var(--dex-green, #86bc25)'; e.currentTarget.style.boxShadow = '0 4px 12px rgba(134,188,37,0.12)'; }}
                onMouseLeave={e => { e.currentTarget.style.borderColor = 'var(--dex-gray-200)'; e.currentTarget.style.boxShadow = 'none'; }}
              >
                {/* v15.3: Event-Bild als Thumbnail links — Fallback grünes Gradient mit Initialen. */}
                <div style={{
                  flex: '0 0 auto', width: 84, height: 84, borderRadius: 10,
                  background: ev.imageUrl
                    ? `url(${ev.imageUrl}) center/cover no-repeat`
                    : 'linear-gradient(135deg, var(--dex-green, #86bc25), var(--dex-blue, #0076a8))',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  color: '#fff', fontWeight: 700, fontSize: '1.4rem',
                  overflow: 'hidden', flexShrink: 0,
                }}>
                  {!ev.imageUrl && ev.title ? ev.title.charAt(0).toUpperCase() : ''}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontSize: '1rem', color: 'var(--dex-gray-800)', marginBottom: 2 }}>{ev.title}</div>
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
                    {ev.startDate ? new Date(ev.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : ''}
                    {ev.location ? ` · ${ev.location}` : ''}
                  </div>
                  {/* v15.3: Organizer-Chips mit Hover-Profilfoto (gleiches OrganizerList wie auf der Registrierungsseite). */}
                  {(ev.organizers && ev.organizers.length > 0) && (
                    <div style={{ marginTop: 6 }}>
                      <OrganizerList
                        names={ev.organizers}
                        emails={ev.organizerEmails || []}
                        size="sm"
                        compact
                      />
                    </div>
                  )}
                </div>
              </button>
            ))}
          </div>
        )}
      </div>
    );
  }

  // v20.1: Self-Check-in — grundsätzlich immer verfügbar. Hat das Event noch
  // keinen aktiven Token (Wizard-Toggle nie gesetzt), wird Self-Check-in beim
  // ersten Klick automatisch aktiviert: Token erzeugen + am Event speichern.
  // Das Persistieren braucht Schreibrechte auf DEX_Events (Organizer/Admin) —
  // reine Check-in-Helfer bekommen in dem Fall einen klaren Hinweis. Existiert
  // der Token bereits, funktionieren beide Aktionen ohne Schreibzugriff.
  const ensureSelfCheckInReady = async (): Promise<string | null> => {
    if (selectedEvent.selfCheckInEnabled && selectedEvent.selfCheckInToken) {
      return selectedEvent.selfCheckInToken;
    }
    const token = selectedEvent.selfCheckInToken || generateSelfCheckInToken();
    let ok = false;
    try {
      ok = await updateEvent(selectedEvent.id, { 'SelfCheckInEnabled': true, 'SelfCheckInToken': token });
    } catch { ok = false; }
    if (!ok) {
      showAlert(isDe
        ? 'Der Self-Check-in-QR für dieses Event konnte noch nicht eingerichtet werden (fehlende Berechtigung zum Speichern am Event). Bitte einmalig von einem Organizer oder Admin öffnen lassen — danach kann auch das Check-in-Team QR-Anzeige und PDF nutzen.'
        : 'The self check-in QR for this event could not be set up yet (missing permission to save on the event). Please have an organizer or admin open it once — afterwards the check-in team can use the QR display and PDF as well.', { variant: 'error' });
      return null;
    }
    return token;
  };
  const openSelfCheckInDisplay = async (): Promise<void> => {
    if (selfCheckInBusy) return;
    setSelfCheckInBusy(true);
    try {
      const token = await ensureSelfCheckInReady();
      if (token) navigate('self-checkin-display', selectedEvent.id);
    } finally { setSelfCheckInBusy(false); }
  };
  const downloadSelfCheckInQrPdf = async (): Promise<void> => {
    if (selfCheckInBusy) return;
    setSelfCheckInBusy(true);
    try {
      const token = await ensureSelfCheckInReady();
      if (!token) return;
      const dateLabel = selectedEvent.startDate
        ? new Date(selectedEvent.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })
        : '';
      await downloadSelfCheckInPdf({
        eventTitle: selectedEvent.title || 'Event',
        eventDateLabel: dateLabel,
        locationLabel: selectedEvent.location || '',
        token,
      });
    } finally { setSelfCheckInBusy(false); }
  };

  // v20.3: Breite begrenzen — vorher lief die Check-in-Ansicht auf großen
  // Monitoren über die volle Viewport-Breite (Event-Picker-Branch hatte
  // bereits maxWidth 1100, die Haupt-Ansicht nicht). Schmaler = lesbarer.
  return (
    <div className="page-container" role="main" style={{ maxWidth: 900, marginLeft: 'auto', marginRight: 'auto', width: '100%' }}>
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
            {isDe ? 'Kamera nicht verfügbar in der SharePoint App' : 'Camera not available in the SharePoint app'}
          </h3>
          <p style={{ color: '#bf360c', fontSize: '0.85rem', lineHeight: 1.6, margin: '0 0 12px' }}>
            {isDe
              ? <>Die SharePoint Mobile App unterstützt keinen Kamera-Zugriff für Webparts. Bitte öffne diese Seite in <strong>Edge</strong> oder <strong>Safari</strong> auf deinem Handy — dort funktioniert der QR-Scanner.</>
              : <>The SharePoint mobile app does not support camera access for web parts. Please open this page in <strong>Edge</strong> or <strong>Safari</strong> on your phone — the QR scanner works there.</>}
          </p>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <button
              className="btn btn-primary"
              style={{ fontSize: '0.85rem' }}
              onClick={() => {
                const url = window.location.href;
                navigator.clipboard.writeText(url).then(() => {
                  setResultMessage(isDe ? 'Link kopiert! Öffne ihn in Edge oder Safari.' : 'Link copied! Open it in Edge or Safari.');
                  setResultType('info');
                }).catch(() => {
                  showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{url}</span>, { title: isDe ? 'Link kopieren und in Edge/Safari öffnen' : 'Copy link and open in Edge/Safari' });
                });
              }}
            >
              {isDe ? 'Link kopieren' : 'Copy link'}
            </button>
          </div>
          <p style={{ color: '#bf360c', fontSize: '0.75rem', marginTop: 8, marginBottom: 0 }}>
            {isDe
              ? 'Tipp: Lege dir die Seite als Lesezeichen in Edge an für schnellen Zugriff beim Event.'
              : 'Tip: Bookmark this page in Edge for quick access during the event.'}
          </p>
        </div>
      )}

      {/* Bestätigungs-Dialog nach Scan */}
      {pendingCheckIn && (
        <div ref={confirmCardRef} className="card" role="dialog" aria-modal="true" aria-label={isDe ? 'Check-in bestätigen' : 'Confirm check-in'} style={{
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
            {isDe ? 'Event: ' : 'Event: '}<strong>{pendingCheckIn.event.title}</strong>
          </p>
          <div style={{ display: 'flex', gap: 12 }}>
            <button
              className="btn btn-primary"
              onClick={confirmCheckIn}
              style={{ flex: 1, fontSize: '1rem', padding: '12px 0', background: 'var(--dex-green)' }}
            >
              {isDe ? 'Einchecken' : 'Check in'}
            </button>
            <button
              className="btn btn-secondary"
              onClick={cancelCheckIn}
              style={{ padding: '12px 20px' }}
            >
              {isDe ? 'Abbrechen' : 'Cancel'}
            </button>
          </div>
        </div>
      )}

      {/* Live-Scanner — Kamerabild + Steuerung */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        {!isScanning ? (
          <>
            <h3 style={{ marginBottom: 12 }}>{isDe ? 'Live-Scanner' : 'Live scanner'}</h3>
            <div style={{ textAlign: 'center' }}>
              <button className="btn btn-primary" onClick={startCamera} style={{ fontSize: '1.1rem', padding: '14px 36px' }}>
                {t('checkin.scan')}
              </button>
              {cameraError && (
                <div style={{ marginTop: 12 }}>
                  <p style={{ color: 'var(--dex-orange)', fontSize: '0.85rem', margin: 0 }}>{cameraError}</p>
                  <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', marginTop: 8, fontWeight: 600 }}>
                    Alternativ: Du kannst Teilnehmer auch in der Liste unten suchen
                    und per Klick auf &quot;Einchecken&quot; manuell einchecken.
                  </p>
                </div>
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
        {/* Wichtig: Video nicht mit fixer Höhe + objectFit:cover croppen, sonst landet
            die vom qr-scanner eingeblendete Scan-Region-Box verschoben, weil die
            Library Overlay-Koordinaten aus dem nativen Video-Aspect berechnet.
            Video soll seine natürliche Aspect Ratio behalten (width:100%, height:auto).
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

      {/* v20.1: Self-Check-in — prominent direkt unter dem Live-Scanner.
          Teilnehmer scannen den Event-QR mit der NATIVEN Handy-Kamera (kein
          Kamera-Zugriff in der App nötig) und checken sich selbst ein. */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        <h3 style={{ marginBottom: 6 }}>Self-Check-in</h3>
        <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', margin: '0 0 14px' }}>
          {isDe
            ? 'Teilnehmer scannen den Event-QR mit der normalen Handy-Kamera und checken sich selbst ein — ohne Scanner-Team und ohne Kamera-Freigabe in der App. Live-Anzeige für einen Bildschirm am Eingang (Code rotiert, foto-sicher) oder PDF zum Ausdrucken und Aushängen.'
            : 'Attendees scan the event QR with their regular phone camera and check themselves in — no scanner team and no in-app camera access needed. Live display for a screen at the entrance (rotating code, photo-safe) or a printable PDF to post.'}
        </p>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'center' }}>
          <button
            className="btn btn-primary"
            disabled={selfCheckInBusy}
            onClick={openSelfCheckInDisplay}
            style={{ fontSize: '0.95rem', padding: '12px 24px' }}
          >
            {isDe ? 'Live-QR anzeigen' : 'Show live QR'}
          </button>
          <button
            className="btn btn-secondary"
            disabled={selfCheckInBusy}
            onClick={downloadSelfCheckInQrPdf}
            style={{ fontSize: '0.95rem', padding: '12px 24px' }}
          >
            {isDe ? 'QR-PDF herunterladen (drucken)' : 'Download QR PDF (print)'}
          </button>
        </div>
        {selfCheckInBusy && (
          <p style={{ textAlign: 'center', fontSize: '0.8rem', color: 'var(--dex-gray-400)', margin: '10px 0 0' }}>
            {isDe ? 'Wird vorbereitet…' : 'Preparing…'}
          </p>
        )}
      </div>

      {/* v7.16: Foto-Upload-Fallback (Galerie / Kamera-Capture per File-Input)
          ist raus. Wenn der Live-Scanner nicht startet, weisen wir in der
          Camera-Error-Message direkt auf die Live-Teilnehmerliste hin —
          dort kann man per Klick manuell einchecken. */}

      {/* v7.14: Live-Teilnehmerliste mit Foto / Position / Standort + Filter.
          Liste wird sofort beim Auswählen des Events geladen, kein "ab 2
          Zeichen tippen" mehr — der Helfer sieht direkt alle Leute, kann den
          gesuchten Eintrag scrollen oder das Suchfeld zum Filtern nutzen. */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        <h3 style={{ marginBottom: 12 }}>Teilnehmer einchecken</h3>
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
        {/* v7.16: KPI-Bereich — angemeldet vs. eingecheckt, plus Quick-Filter
            "Nur offene anzeigen" der die Liste auf Angemeldet/QR versendet
            reduziert. Sichtbar nur wenn Event gewählt + Liste geladen. */}
        {nameSearchEventId && (searchRegsCache[nameSearchEventId] || []).length > 0 && (
          <div style={{
            display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 8,
            marginBottom: 10,
          }}>
            <div style={{
              padding: '10px 12px', borderRadius: 10,
              background: 'rgba(21,101,192,0.08)', textAlign: 'center',
            }}>
              <div style={{ fontSize: '1.4rem', fontWeight: 700, color: '#1565c0', lineHeight: 1 }}>
                {checkInKpis.registered}
              </div>
              <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                Angemeldet
              </div>
            </div>
            <div style={{
              padding: '10px 12px', borderRadius: 10,
              background: 'rgba(134,188,37,0.12)', textAlign: 'center',
            }}>
              <div style={{ fontSize: '1.4rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', lineHeight: 1 }}>
                {checkInKpis.checkedIn}
                <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-400)', fontWeight: 500 }}>
                  /{checkInKpis.registered}
                </span>
              </div>
              <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                Eingecheckt
              </div>
            </div>
            {/* v23.28: No-Show-Zähler. */}
            <div style={{
              padding: '10px 12px', borderRadius: 10,
              background: 'rgba(96,96,96,0.10)', textAlign: 'center',
            }}>
              <div style={{ fontSize: '1.4rem', fontWeight: 700, color: 'var(--dex-gray-700, #444)', lineHeight: 1 }}>
                {checkInKpis.noShow}
              </div>
              <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                No-Show
              </div>
            </div>
          </div>
        )}
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          <input
            className="form-input"
            value={nameSearchQuery}
            onChange={e => setNameSearchQuery(e.target.value)}
            placeholder="Filter: Vorname, Nachname oder E-Mail…"
            disabled={!nameSearchEventId}
            style={{ flex: '1 1 200px', padding: '10px 14px', fontSize: '0.95rem' }}
          />
          <button
            type="button"
            onClick={() => setOnlyOpen(v => !v)}
            disabled={!nameSearchEventId}
            title={onlyOpen ? 'Alle Teilnehmer zeigen' : 'Nur offene Anmeldungen zeigen (Eingecheckte ausblenden)'}
            style={{
              padding: '8px 14px', borderRadius: 10,
              border: `1px solid ${onlyOpen ? 'var(--dex-green)' : 'var(--dex-gray-300)'}`,
              background: onlyOpen ? 'rgba(134,188,37,0.10)' : '#fff',
              color: onlyOpen ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-700)',
              cursor: nameSearchEventId ? 'pointer' : 'not-allowed',
              fontSize: '0.82rem', fontWeight: 600, fontFamily: 'inherit',
              whiteSpace: 'nowrap', flexShrink: 0,
              opacity: nameSearchEventId ? 1 : 0.5,
            }}
          >
            {onlyOpen ? '✓ Nur offene' : 'Nur offene'}
          </button>
        </div>
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
        {nameSearchEventId && !isLoadingSearchRegs && (
          <div style={{ marginTop: 12 }}>
            {searchHits.length === 0 ? (
              <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)', fontStyle: 'italic', margin: 0 }}>
                {nameSearchQuery.trim().length > 0
                  ? 'Kein Treffer — bitte anders schreiben oder per QR-Code einchecken.'
                  : 'Keine Teilnehmer für dieses Event.'}
              </p>
            ) : (
              <>
                <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-400)', marginBottom: 6 }}>
                  {searchHits.length} {searchHits.length === 1 ? 'Teilnehmer' : 'Teilnehmer'}
                  {nameSearchQuery.trim().length > 0 ? ' (gefiltert)' : ''}
                </div>
                <div style={{
                  display: 'flex', flexDirection: 'column', gap: 8,
                  // Maximalhöhe + Scroll, damit lange Events nicht die ganze Seite sprengen.
                  maxHeight: 480, overflowY: 'auto',
                  paddingRight: 4,
                }}>
                  {searchHits.map(reg => {
                    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || reg.ParticipantEmail || '-');
                    const status = reg.Status;
                    const alreadyIn = status === 'Eingecheckt';
                    const cancelled = status === 'Abgemeldet';
                    const waitlist = status === 'Warteliste';
                    const noShow = status === 'No-Show';
                    const statusBg = alreadyIn ? 'rgba(134,188,37,0.15)'
                      : cancelled ? 'rgba(204,0,0,0.10)'
                      : noShow ? 'rgba(96,96,96,0.14)'
                      : waitlist ? 'rgba(237,139,0,0.12)'
                      : 'rgba(21,101,192,0.10)';
                    const statusFg = alreadyIn ? 'var(--dex-green-dark, #4a7c1f)'
                      : cancelled ? 'var(--dex-red, #c00)'
                      : noShow ? 'var(--dex-gray-700, #444)'
                      : waitlist ? 'var(--dex-orange, #ed8b00)'
                      : '#1565c0';
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const jobTitle = (reg as any).JobTitle || '';
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const location = (reg as any).Location || '';
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const ctx = (window as any).__dexSpfxContext;
                    const photoSrc = ctx && reg.ParticipantEmail
                      ? `${ctx.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(reg.ParticipantEmail)}`
                      : '';
                    return (
                      <div
                        key={reg.Id}
                        style={{
                          display: 'flex', alignItems: 'center', gap: 12,
                          padding: '10px 12px', border: '1px solid var(--dex-gray-200)',
                          borderRadius: 10, background: '#fff',
                        }}
                      >
                        {/* Foto-Avatar mit Initials-Fallback. Initialen liegen
                            im Hintergrund, das <img> deckt sie ab — schlägt
                            der Image-Load fehl, kommen sie zum Vorschein. */}
                        <div style={{
                          position: 'relative',
                          width: 44, height: 44, borderRadius: '50%',
                          flexShrink: 0, overflow: 'hidden',
                          background: 'var(--dex-gray-100)',
                          display: 'flex', alignItems: 'center', justifyContent: 'center',
                          fontWeight: 600, fontSize: '0.85rem', color: 'var(--dex-gray-500)',
                        }}>
                          <span style={{ pointerEvents: 'none' }}>
                            {(`${(reg.Vorname || reg.ParticipantName || '').charAt(0)}${(reg.Nachname || '').charAt(0)}`.toUpperCase()) || '?'}
                          </span>
                          {photoSrc && (
                            <img
                              src={photoSrc}
                              alt={name}
                              style={{
                                position: 'absolute', inset: 0,
                                width: '100%', height: '100%', objectFit: 'cover',
                              }}
                              onError={e => {
                                (e.currentTarget as HTMLImageElement).style.display = 'none';
                              }}
                            />
                          )}
                        </div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 600, fontSize: '0.9rem', color: 'var(--dex-gray-800)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                            {name}
                          </div>
                          {(jobTitle || location) && (
                            <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                              {[jobTitle, location].filter(Boolean).join(' · ')}
                            </div>
                          )}
                          <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-400)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                            {reg.ParticipantEmail}
                          </div>
                        </div>
                        <span style={{
                          fontSize: '0.7rem', padding: '3px 8px', borderRadius: 999,
                          background: statusBg, color: statusFg, fontWeight: 600, whiteSpace: 'nowrap',
                        }}>{status}</span>
                        {/* v23.28: Check-in UND No-Show nebeneinander. */}
                        <div style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
                          <button
                            type="button"
                            className="btn btn-primary"
                            style={{ fontSize: '0.78rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
                            disabled={alreadyIn || cancelled || isProcessing}
                            onClick={() => startManualCheckInFromSearch(reg)}
                          >
                            {alreadyIn ? '✓ Eingecheckt' : cancelled ? 'Abgemeldet' : 'Einchecken'}
                          </button>
                          <button
                            type="button"
                            className="btn btn-secondary"
                            style={{ fontSize: '0.78rem', padding: '6px 12px', whiteSpace: 'nowrap', color: 'var(--dex-gray-700, #444)' }}
                            disabled={cancelled || noShow || isProcessing}
                            onClick={() => { void markNoShowFromSearch(reg); }}
                            title={isDe ? 'Teilnehmer als nicht erschienen markieren' : 'Mark attendee as a no-show'}
                          >
                            {noShow ? (isDe ? 'No-Show' : 'No-show') : 'No-Show'}
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </div>
        )}
      </div>

      {/* v7.14: Die alte "Manuell QR-Code als String tippen"-Card ist raus.
          Die Live-Teilnehmerliste oben deckt das Manuelle Einchecken
          benutzerfreundlicher ab. */}
    </div>
  );
}
