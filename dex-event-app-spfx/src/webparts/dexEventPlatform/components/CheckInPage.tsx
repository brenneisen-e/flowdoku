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
import { checkInExtras, parseCustomData, CheckInExtra, shirtAllocate, parseShirtStock, ShirtAllocationResult } from '../utils/checkInExtras';
import { parseAgendaCheckIns, parseAgendaMarks, parseAgendaNoShows, formatMarkTime, suggestCurrentAgendaItem } from '../utils/agendaCheckIns';
import Modal from './Modal';
import { agendaGroups, groupLabel, groupDateLabel } from '../utils/agendaGroups';
import { useLanguage } from '../context/LanguageContext';
import { useIsMobile } from '../utils/useIsMobile';
import OrganizerList from './OrganizerList';
import { ChevronDown, ChevronUp } from './Icons';
// v20.0 (Audit): qr-scanner nur noch als Typ statisch importieren — die
// eigentliche Bibliothek wird erst beim Kamera-Start dynamisch nachgeladen.
import type QrScanner from 'qr-scanner';
import { shortSubEventTitle } from '../utils/subEventTitle';


export default function CheckInPage(): React.ReactElement {
  const { events, getAllRegistrations, updateEvent } = useEvents();
  // v20.4: App-Modal statt nativem Browser-Alert.
  const { showAlert, confirmDialog } = useDialog();
  const { selectedEventId, navigate } = useNavigation();
  const { isAdmin, isOrganizer, siteUrl } = useRoles();
  const { currentUser } = useCurrentUser();
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  const isMobile = useIsMobile();
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
  // v23.45: Sub-Events werden im Picker unter ihrem Hauptevent gruppiert und
  // standardmäßig eingeklappt — statt flacher Liste. Ein Hauptevent mit
  // Sub-Events bekommt einen Aufklapp-Pfeil; Klick auf die Karte selbst checkt
  // weiterhin direkt in das (Haupt-)Event ein.
  const groupedCheckInEvents = React.useMemo(() => {
    const visibleIds = new Set(visibleCheckInEvents.map(e => e.id));
    const childrenByParent: Record<string, typeof visibleCheckInEvents> = {};
    visibleCheckInEvents.forEach(e => {
      if (e.parentEventId && visibleIds.has(e.parentEventId)) {
        (childrenByParent[e.parentEventId] = childrenByParent[e.parentEventId] || []).push(e);
      }
    });
    // Top-Level = Events ohne sichtbaren Parent (echte Hauptevents ODER
    // verwaiste Sub-Events, deren Parent ausgeblendet/gefiltert ist).
    const topLevel = visibleCheckInEvents.filter(e => !(e.parentEventId && visibleIds.has(e.parentEventId)));
    return topLevel.map(parent => ({ parent, children: childrenByParent[parent.id] || [] }));
  }, [visibleCheckInEvents]);
  const [expandedCheckInParents, setExpandedCheckInParents] = React.useState<Record<string, boolean>>({});
  const scannerRef = React.useRef<QrScanner | null>(null);
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const [resultMessage, setResultMessage] = React.useState('');
  const [resultType, setResultType] = React.useState<'success' | 'error' | 'info' | ''>('');
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [isScanning, setIsScanning] = React.useState(false);
  const [cameraError, setCameraError] = React.useState('');
  // v30.29: Merkt sich, ob der Kamera-Fehler aus der iframe-Einbettung kam —
  // nur dann hilft der „Seite im Browser öffnen"-Knopf, sonst wäre er ein
  // Angebot, das nichts löst.
  const [cameraErrorInIframe, setCameraErrorInIframe] = React.useState(false);
  // v30.30: Foto-Weg — native Kamera-App per File-Input statt getUserMedia.
  const photoInputRef = React.useRef<HTMLInputElement>(null);
  // v30.31: Nur für die `accept`-Erweiterung unten — bewusst eng gefasste
  // UA-Prüfung, weil der Workaround eine Android-Eigenheit adressiert.
  const isAndroid = typeof navigator !== 'undefined' && /Android/i.test(navigator.userAgent || '');
  const [photoBusy, setPhotoBusy] = React.useState(false);
  const [photoHint, setPhotoHint] = React.useState('');
  // v30.35: Direkteingabe der Teilnehmer-ID (s. Block unter dem Scan-Knopf).
  const [idInput, setIdInput] = React.useState('');
  const [idError, setIdError] = React.useState('');
  const [checkedInCount, setCheckedInCount] = React.useState(0);
  const confirmCardRef = React.useRef<HTMLDivElement>(null);
  type PendingCheckInInfo = {
    name: string; email: string; event: { id?: string; subsiteUrl: string; title: string };
    regId: number; status: string; department?: string; jobTitle?: string; location?: string; photoUrl?: string;
    /** v30.53: Startnummer, Trikotgröße, Gruppe — was am Tisch gebraucht wird. */
    extras?: CheckInExtra[];
    /** v30.91: Programmpunkt, an dem eingecheckt wird (statt Event-Status). */
    agendaItemId?: string;
    agendaLabel?: string;
  };
  const [pendingCheckIn, setPendingCheckIn] = React.useState<PendingCheckInInfo | null>(null);
  // v31.1: „Letzte Check-ins" dieser Sitzung — mit Rückgängig. Nur was HIER
  // eingecheckt wurde (Scan, ID, Liste); der vorherige Status wird gemerkt,
  // damit der Revert nichts erfindet.
  // v31.2: `kind` — auch ein No-Show landet hier und ist rücknehmbar (Nutzer
  // 07.09.2026: „No-Show soll auch rückgängig machbar sein").
  type RecentCheckIn = {
    key: string; at: string; name: string; regId: number; eventId: string; subsiteUrl: string;
    prevStatus: string; agendaItemId?: string; agendaLabel?: string; kind?: 'checkin' | 'noshow';
  };
  const [recentCheckIns, setRecentCheckIns] = React.useState<RecentCheckIn[]>([]);
  const [undoBusyKey, setUndoBusyKey] = React.useState<string>('');
  // v31.2: Rückfrage im Programmpunkt-Modus — No-Show nur am gewählten Punkt
  // oder für das ganze Event? Beides sind andere Daten (Marke vs. Status).
  const [noShowAsk, setNoShowAsk] = React.useState<{ reg: import('../services/EventService').SPRegistration; name: string } | null>(null);
  // v31.1: Live-Scanner und Self-Check-in sind Kacheln zum Aufklappen —
  // Standard zu (Nutzer 07.09.2026). Läuft der Scanner, ist die Kachel offen.
  const [scannerOpen, setScannerOpen] = React.useState(false);
  const [selfOpen, setSelfOpen] = React.useState(false);


  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const context = (window as any).__dexSpfxContext;
  const eventService = React.useMemo(() => context ? new EventService(context) : null, []);

  const selectedEvent = selectedEventId ? events.find(e => e.id === selectedEventId) : null;

  /**
   * v30.53: Startnummer + Trikotgröße + Gruppe zu einer Zeile.
   *
   * Am Check-in-Tisch eines Laufs wird nicht nur abgehakt, sondern auch das
   * Trikot ausgegeben und die Startnummer genannt. Ohne diese Angaben muss
   * daneben eine Excel offen sein — genau dort entsteht die Schlange. Die
   * Regel, WELCHE Felder das sind, steht in utils/checkInExtras.
   */
  // v30.88: Trikot-Verteilung je Event — dieselbe Rechnung wie in der Aktion
  // „Benötigte T-Shirts" (utils/checkInExtras.shirtAllocate), gecacht je
  // Registrierungs-Array. Der Cache der Teilnehmerlisten wird weiter unten
  // deklariert; extrasFor liest ihn über eine Ref, damit die Closure nie einen
  // veralteten Stand sieht.
  const searchRegsCacheRef = React.useRef<Record<string, import('../services/EventService').SPRegistration[]>>({});
  const shirtAllocCacheRef = React.useRef<Map<object, ShirtAllocationResult>>(new Map());
  const shirtAllocFor = React.useCallback((eventId: string): ShirtAllocationResult | null => {
    const ev = events.find(e => e.id === eventId);
    const rs = searchRegsCacheRef.current[eventId];
    if (!ev || !rs) return null;
    const stock = parseShirtStock(ev.emailTemplateOverrides);
    if (Object.keys(stock).length === 0) return null;
    const hit = shirtAllocCacheRef.current.get(rs);
    if (hit) return hit;
    const res = shirtAllocate(ev.eventSpecificFields, rs, stock);
    shirtAllocCacheRef.current.set(rs, res);
    return res;
  }, [events]);
  const extrasFor = React.useCallback((
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    reg: any,
    eventId: string,
  ): CheckInExtra[] => {
    const ev = events.find(e => e.id === eventId);
    const out = checkInExtras(
      ev?.eventSpecificFields,
      parseCustomData(reg?.CustomData),
      reg,
      { bib: isDe ? 'Startnummer' : 'Bib number', group: isDe ? 'Gruppe' : 'Group' },
    );
    // v30.88: Gegenvorschlag, wenn die Wunschgröße laut Bestand nicht reicht —
    // die Antwort auf „passt es überhaupt?" gehört an den Tisch, nicht in eine Excel.
    const alloc = shirtAllocFor(eventId);
    const em = String((reg && reg.ParticipantEmail) || '').toLowerCase().trim();
    const a = alloc && em ? alloc.byEmail[em] : undefined;
    if (a && a.short) {
      out.push({
        label: isDe ? 'Trikot-Vorschlag' : 'Shirt proposal',
        value: a.proposal
          ? (isDe ? `${a.proposal} statt ${a.wish} (nicht mehr vorrätig)` : `${a.proposal} instead of ${a.wish} (out of stock)`)
          : (isDe ? `${a.wish} nicht mehr vorrätig — keine Ausweichgröße` : `${a.wish} out of stock — no alternative left`),
        strong: true,
      });
    }
    return out;
  }, [events, isDe, shirtAllocFor]);

  // v7.12: Name-Suche für manuelles Einchecken — wenn der QR-Scanner in der
  // SP-App nicht funktioniert (Camera-API gesperrt) oder der Teilnehmer den
  // QR-Code nicht zur Hand hat, kann der Helfer nach Namen / E-Mail suchen
  // und per Tap "Einchecken" auslösen. Die Registrierungen werden pro Event
  // lazy nachgeladen und in `searchRegsCache` zwischengespeichert.
  const [nameSearchQuery, setNameSearchQuery] = React.useState('');
  const [nameSearchEventId, setNameSearchEventId] = React.useState<string>(selectedEventId || '');
  const [searchRegsCache, setSearchRegsCache] = React.useState<Record<string, import('../services/EventService').SPRegistration[]>>({});
  searchRegsCacheRef.current = searchRegsCache; // v30.88 (s. shirtAllocFor)

  // v30.91: Programmpunkte (Konzept docs/konzept-programmpunkte.md, Stufe 2).
  // Bei einem Event mit `agendaCheckIn` wird nicht das Event, sondern EIN
  // Programmpunkt eingecheckt: Das Team wählt oben den Punkt, jeder Scan,
  // jede ID und jeder Klick setzt nur diesen Punkt in `AgendaCheckIns`. Der
  // Event-Status bleibt unberührt (Nutzer-Entscheidung 07.09.2026).
  const agendaEv = React.useMemo(() => events.find(e => e.id === nameSearchEventId) || null, [events, nameSearchEventId]);
  const agendaItems = React.useMemo(() => (agendaEv && agendaEv.agendaCheckIn
    ? (agendaEv.agenda || []).slice().sort((a, b) => ((a.date || '') + (a.time || '')).localeCompare((b.date || '') + (b.time || '')))
    : []), [agendaEv]);
  const agendaMode = agendaItems.length > 0;
  const agendaTermSingular = (agendaEv && agendaEv.agendaTermSingular) || (isDe ? 'Programmpunkt' : 'Agenda item');
  const [agendaPointId, setAgendaPointId] = React.useState<string>('');
  React.useEffect(() => {
    if (!agendaMode) { setAgendaPointId(''); return; }
    let stored = '';
    try { stored = window.localStorage.getItem(`dex_checkin_point_${nameSearchEventId}`) || ''; } catch { /* */ }
    if (agendaItems.some(a => a.id === stored)) { setAgendaPointId(stored); return; }
    const s = suggestCurrentAgendaItem(agendaItems);
    setAgendaPointId(s ? s.id : '');
  }, [agendaMode, nameSearchEventId, agendaItems]);
  const choosePoint = (id: string): void => {
    setAgendaPointId(id);
    try { window.localStorage.setItem(`dex_checkin_point_${nameSearchEventId}`, id); } catch { /* */ }
  };
  const agendaPoint = agendaItems.find(a => a.id === agendaPointId) || null;
  const presentAt = (reg: { AgendaCheckIns?: string } | null | undefined, pointId: string): boolean =>
    !!(reg && pointId && parseAgendaCheckIns(reg.AgendaCheckIns)[pointId]);
  // Zähler je Punkt (aktive Anmeldungen) — für die Chips der Punkt-Wahl.
  const agendaCounts = React.useMemo((): Record<string, number> => {
    const out: Record<string, number> = {};
    if (!agendaMode) return out;
    const regs = searchRegsCache[nameSearchEventId] || [];
    for (const r of regs) {
      if (r.Status === 'Abgemeldet') continue;
      const marks = parseAgendaCheckIns(r.AgendaCheckIns);
      Object.keys(marks).forEach(k => { out[k] = (out[k] || 0) + 1; });
    }
    return out;
  }, [agendaMode, searchRegsCache, nameSearchEventId]);
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
      // v30.67: `getAllRegistrations` wirft bei HTTP-Fehlern NICHT, sondern
      // liefert die bis dahin gelesenen Zeilen — bei 403/429 also `[]`. Das
      // landete hier als gültige, leere Liste im Cache: „Keine Teilnehmer",
      // „Keine Anmeldung mit der Teilnehmer-ID …", und weil der Cache-
      // Schlüssel gesetzt war, kein zweiter Versuch mehr bis zum Seiten-
      // Reload. Nur der Rückruf unterscheidet „leer" von „verboten"; ohne
      // geprüften Status kommt nichts in den Cache (CLAUDE.md, dritter Pfad).
      let readable = true;
      let httpStatus = 0;
      const regs = await getAllRegistrations(eventId, (status) => { readable = false; httpStatus = status; });
      if (!readable) {
        // v30.67 (Review): zweisprachig wie der Nachbarpfad checkInByParticipantId.
        setSearchLoadError(httpStatus === 403
          ? (isDe
            ? 'Keine Leseberechtigung auf der Teilnehmerliste dieses Termins — bitte Organizer/Admin um Freigabe bitten.'
            : 'No read permission on this date\'s attendee list — please ask an organizer/admin for access.')
          : (isDe
            ? `Teilnehmerliste konnte nicht gelesen werden (${httpStatus ? 'HTTP ' + httpStatus : 'keine Teilnehmerliste gefunden'}) — bitte erneut versuchen.`
            : `The attendee list could not be read (${httpStatus ? 'HTTP ' + httpStatus : 'no attendee list found'}) — please try again.`));
      } else {
        setSearchRegsCache(prev => ({ ...prev, [eventId]: regs }));
      }
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
    } else if (searchLoadError) {
      // v30.67: Der Ladefehler gehört zu dem Event, bei dem er entstand —
      // beim Wechsel auf ein bereits geladenes Event darf er nicht stehen
      // bleiben (loadRegsForSearch kehrt bei Cache-Treffer vorher zurück).
      setSearchLoadError('');
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
      // v30.91: Im Programmpunkt-Modus zählt die Kachel die Anwesenden am
      // gewählten Punkt — der Event-Status sagt dort nichts.
      if (agendaMode ? presentAt(r, agendaPointId) : r.Status === 'Eingecheckt') checkedIn++;
      // v31.2: Im Programmpunkt-Modus zählt die Kachel die No-Shows AM PUNKT
      // (Marke), sonst den Event-Status.
      if (agendaMode && agendaPointId ? !!parseAgendaNoShows(r.AgendaCheckIns)[agendaPointId] : r.Status === 'No-Show') noShow++;
    }
    return { registered, checkedIn, noShow };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nameSearchEventId, searchRegsCache, agendaMode, agendaPointId]);

  // v7.14: Live-Filter über die ganze Liste — leerer Query zeigt alle
  // Teilnehmer. Sortiert nach Status (Aktive zuerst), dann Nachname.
  const searchHits = React.useMemo(() => {
    if (!nameSearchEventId) return [];
    const regs = searchRegsCache[nameSearchEventId] || [];
    const q = nameSearchQuery.trim().toLowerCase();
    // v30.33: Die Teilnehmer-ID ist jetzt suchbar — und zwar EXAKT, nicht als
    // Teiltreffer. Sie steht schon heute unter jedem QR-Code in der Mail; sie
    // eintippen zu können macht den Check-in unabhängig von der Kamera, die auf
    // verwalteten Geräten nicht überall erreichbar ist.
    //
    // Exakt deshalb, weil ein Teiltreffer bei „17" auch 117 und 170 liefern
    // würde — am Einlass die falsche Person einzuchecken ist schlimmer als
    // einmal mehr zu tippen. Die Namens-/Mail-Suche bleibt zusätzlich als
    // Teiltreffer bestehen, damit „17" auch eine Mail mit 17 darin findet.
    //
    // Eindeutig ist die Zahl, weil der Check-in immer AUF EIN EVENT bezogen
    // ist (Event-Picker davor) und TeilnehmerID je Event fortlaufend vergeben
    // wird. Ein Event-Präfix braucht es hier deshalb nicht.
    // v30.83: numerisch vergleichen — „005" (so steht die Nummer mit
    // führenden Nullen in der QR-Mail) ist dieselbe ID wie 5. Der String-
    // Vergleich lieferte „Kein Treffer" (Befund 07.09.2026).
    const numericQ = /^\d+$/.test(q) ? parseInt(q, 10) : NaN;
    const matchesQuery = q.length === 0
      ? regs
      : regs.filter(r => {
          if (isFinite(numericQ) && r.TeilnehmerID !== undefined && r.TeilnehmerID !== null && Number(r.TeilnehmerID) === numericQ) return true;
          const full = `${r.Vorname || ''} ${r.Nachname || ''} ${r.ParticipantName || ''} ${r.ParticipantEmail || ''}`.toLowerCase();
          return full.indexOf(q) >= 0;
        });
    // v7.16: optionaler Quick-Filter "nur offene Anmeldungen"
    const filtered = onlyOpen
      ? matchesQuery.filter(r => agendaMode
        ? (r.Status !== 'Abgemeldet' && r.Status !== 'Warteliste' && !presentAt(r, agendaPointId))
        : (r.Status === 'Angemeldet' || r.Status === 'QR versendet'))
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
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nameSearchQuery, nameSearchEventId, searchRegsCache, onlyOpen, agendaMode, agendaPointId]);

  /**
   * v30.35: Check-in über die eingetippte Teilnehmer-ID.
   *
   * Nutzt bewusst denselben Weg wie ein Klick in der Trefferliste
   * (`startManualCheckInFromSearch`) — also dieselbe Bestätigungskarte mit
   * Foto und Status. Am Einlass will man vor dem Einchecken sehen, WEN man
   * eincheckt; eine ID ohne Gesicht wäre schneller und riskanter.
   *
   * Die Liste ist zu diesem Zeitpunkt geladen (`loadRegsForSearch` läuft beim
   * Auswählen des Events). Ist sie es nicht, sagt die Meldung genau das,
   * statt „ID nicht gefunden" zu behaupten.
   */
  const checkInByParticipantId = async (): Promise<void> => {
    const raw = idInput.trim();
    if (!raw) return;
    setIdError('');
    if (!nameSearchEventId) {
      setIdError(isDe ? 'Bitte zuerst oben das Event auswählen.' : 'Please pick the event above first.');
      return;
    }
    const regs = searchRegsCache[nameSearchEventId];
    if (!regs) {
      // v30.67: Ohne Cache-Eintrag gibt es zwei Zustände — „lädt noch" und
      // „konnte nicht gelesen werden" (403/429, s. loadRegsForSearch). Beide
      // erlauben KEINE Aussage über die ID; der zweite darf aber nicht als
      // Warten verkauft werden, das nie endet.
      setIdError(searchLoadError
        ? (isDe
          ? `Die Teilnehmerliste konnte nicht gelesen werden — ob die ID ${raw} angemeldet ist, lässt sich so nicht sagen. ${searchLoadError}`
          : `The attendee list could not be read — whether ID ${raw} is registered cannot be told. Please retry or ask an organizer/admin for access.`)
        : (isDe ? 'Teilnehmerliste wird noch geladen — bitte kurz warten und erneut versuchen.' : 'Attendee list is still loading — please wait a moment and try again.'));
      void loadRegsForSearch(nameSearchEventId);
      return;
    }
    // v30.83: numerisch — „005" aus der QR-Mail ist ID 5.
    const rawNum = parseInt(raw, 10);
    const hit = regs.filter(r => r.TeilnehmerID !== undefined && r.TeilnehmerID !== null && Number(r.TeilnehmerID) === rawNum);
    if (hit.length === 0) {
      setIdError(isDe
        ? `Keine Anmeldung mit der Teilnehmer-ID ${raw} bei diesem Event. Bitte die Nummer aus der QR-Mail prüfen — oder unten nach dem Namen suchen.`
        : `No registration with attendee ID ${raw} for this event. Please check the number in the QR email — or search by name below.`);
      return;
    }
    if (hit.length > 1) {
      // Sollte nicht vorkommen (ID ist je Event fortlaufend), wäre aber ein
      // Datenfehler, den man am Einlass nicht stillschweigend raten darf.
      setIdError(isDe
        ? `Mehrere Anmeldungen mit der ID ${raw} gefunden — bitte unten über den Namen einchecken und das den DEX-Admins melden.`
        : `Several registrations share ID ${raw} — please check in by name below and report this to the DEX admins.`);
      return;
    }
    setIdInput('');
    setNameSearchQuery(''); // v30.87: Live-Filter der Liste zurücksetzen
    startManualCheckInFromSearch(hit[0]);
  };

  const startManualCheckInFromSearch = (reg: import('../services/EventService').SPRegistration): void => {
    const ev = events.find(e => e.id === nameSearchEventId);
    if (!ev || !ev.subsiteUrl) return;
    if (reg.Status === 'Abgemeldet') {
      setResultMessage(`${reg.ParticipantName || reg.ParticipantEmail} — ${t('checkin.cancelled')}`);
      setResultType('error');
      return;
    }
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || reg.ParticipantEmail);
    // v30.91: Programmpunkt-Modus — ohne gewählten Punkt kein Check-in;
    // schon erfasst → Hinweis mit Uhrzeit, kein zweiter Schreibvorgang.
    if (agendaMode) {
      if (!agendaPoint) {
        setResultMessage(isDe ? `Bitte oben den ${agendaTermSingular} wählen, an dem eingecheckt wird.` : `Please pick the ${agendaTermSingular.toLowerCase()} above first.`);
        setResultType('error');
        return;
      }
      const marks = parseAgendaCheckIns(reg.AgendaCheckIns);
      if (marks[agendaPoint.id]) {
        setResultMessage(`${name} — ${isDe ? 'bereits erfasst' : 'already recorded'} (${agendaPoint.title}, ${formatMarkTime(marks[agendaPoint.id].at)})`);
        setResultType('info');
        return;
      }
    }
    let photoUrl = '';
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (ctx) {
        const siteBase = ctx.pageContext.web.absoluteUrl;
        photoUrl = `${siteBase}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(reg.ParticipantEmail || '')}`;
      }
    } catch { /* */ }
    // v31.1: Liste und Teilnehmer-ID checken DIREKT ein — die Person ist hier
    // schon eindeutig gewählt, eine zweite Bestätigung war nur ein Klick mehr.
    const info: PendingCheckInInfo = {
      name,
      email: reg.ParticipantEmail || '',
      event: { id: ev.id, subsiteUrl: ev.subsiteUrl, title: ev.title },
      regId: reg.Id,
      status: reg.Status,
      agendaItemId: agendaMode && agendaPoint ? agendaPoint.id : undefined,
      agendaLabel: agendaMode && agendaPoint ? agendaPoint.title : undefined,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      department: (reg as any).Department || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      jobTitle: (reg as any).JobTitle || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      location: (reg as any).Location || '',
      photoUrl,
      extras: extrasFor(reg, ev.id),
    };
    setResultMessage('');
    setResultType('');
    setNameSearchQuery('');
    setIsProcessing(true);
    void performCheckIn(info).finally(() => setIsProcessing(false));
  };

  // v23.28: Teilnehmer als „No-Show" markieren (nicht erschienen). Direkt aus
  // der Suchliste; nach Bestätigung wird der lokale Cache aktualisiert.
  // v31.2: Der No-Show landet in „Letzte Check-ins" und ist dort rücknehmbar.
  const rememberNoShow = (reg: import('../services/EventService').SPRegistration, name: string, ev: { id: string; subsiteUrl?: string }, agendaItemId?: string, agendaLabel?: string): void => {
    const entry: RecentCheckIn = {
      key: `${reg.Id}:${agendaItemId || 'status'}:noshow:${Date.now()}`,
      at: new Date().toISOString(), name, regId: reg.Id, eventId: ev.id || '', subsiteUrl: ev.subsiteUrl || '',
      prevStatus: reg.Status, agendaItemId, agendaLabel, kind: 'noshow',
    };
    setRecentCheckIns(prev => [entry, ...prev].slice(0, 30));
  };
  // v31.2: No-Show NUR am gewählten Programmpunkt — Marke mit noShow, der
  // Event-Status bleibt (die Person kann beim nächsten Punkt wieder da sein).
  const applyPointNoShow = async (reg: import('../services/EventService').SPRegistration, name: string): Promise<void> => {
    const ev = events.find(e => e.id === nameSearchEventId);
    if (!ev || !ev.subsiteUrl || !eventService || !agendaPoint) return;
    const pointId = agendaPoint.id;
    const r = await eventService.markAgendaNoShow(ev.subsiteUrl, reg.Id, pointId);
    if (!r.ok) {
      setResultMessage(isDe
        ? `${name} — No-Show konnte nicht gespeichert werden${r.status ? ` (HTTP ${r.status})` : ''}. ${r.status === 400 ? 'Fehlt die Spalte AgendaCheckIns? Organizer: „Spalten fixen" ausführen.' : 'Bitte erneut versuchen.'}`
        : `${name} — no-show could not be saved${r.status ? ` (HTTP ${r.status})` : ''}.`);
      setResultType('error');
      return;
    }
    const at = r.at || new Date().toISOString();
    setSearchRegsCache(prev => {
      const list = prev[nameSearchEventId] || [];
      return { ...prev, [nameSearchEventId]: list.map(x => x.Id === reg.Id
        ? { ...x, AgendaCheckIns: JSON.stringify({ ...parseAgendaMarks(x.AgendaCheckIns), [pointId]: { at, by: '', noShow: true } }) }
        : x) };
    });
    rememberNoShow(reg, name, ev, pointId, agendaPoint.title);
    setResultMessage(isDe ? `${name} — No-Show bei ${agendaPoint.title}.` : `${name} — no-show at ${agendaPoint.title}.`);
    setResultType('info');
  };
  const applyEventNoShow = async (reg: import('../services/EventService').SPRegistration, name: string): Promise<void> => {
    const ev = events.find(e => e.id === nameSearchEventId);
    if (!ev || !ev.subsiteUrl || !eventService) return;
    const success = await eventService.markNoShowParticipant(ev.subsiteUrl, reg.Id);
    if (success) {
      setSearchRegsCache(prev => {
        const list = prev[nameSearchEventId] || [];
        return { ...prev, [nameSearchEventId]: list.map(r => r.Id === reg.Id ? { ...r, Status: 'No-Show' } : r) };
      });
      rememberNoShow(reg, name, ev);
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
  const markNoShowFromSearch = async (reg: import('../services/EventService').SPRegistration): Promise<void> => {
    const ev = events.find(e => e.id === nameSearchEventId);
    if (!ev || !ev.subsiteUrl || !eventService) return;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || reg.ParticipantEmail || '-');
    // v31.2: Mit gewähltem Programmpunkt erst fragen — Punkt oder Event?
    if (agendaMode && agendaPoint) { setNoShowAsk({ reg, name }); return; }
    const ok = await confirmDialog(
      isDe
        ? `„${name}" als nicht erschienen (No-Show) markieren?`
        : `Mark „${name}" as a no-show?`,
      { confirmLabel: isDe ? 'Als No-Show markieren' : 'Mark as no-show' }
    );
    if (!ok) return;
    await applyEventNoShow(reg, name);
  };


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

  // Scanner in neuem Fenster öffnen (aspx-Seite in SiteAssets, für zukünftige Nutzung)
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
    setCameraErrorInIframe(false); // v30.29
    if (!videoRef.current) return;

    // 1. Browser-API-Check: manche Embedded-WebViews (SharePoint-Mobile-App)
    //    stellen mediaDevices gar nicht bereit. Dort gleich mit klarer Meldung
    //    abbrechen statt auf qr-scanner-Fehler zu warten.
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      setCameraError(
        'Dein Browser stellt keinen Kamera-Zugriff bereit. '
        + 'Bitte öffne diese Seite direkt in Edge, angemeldet mit deinem Arbeitskonto (nicht in der SharePoint-App / Teams) — oder nimm den Foto-Weg unten.'
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
        // v30.29: `NotAllowedError` hat ZWEI sehr verschiedene Ursachen, und
        // die alte Meldung nannte immer nur die eine.
        //
        //  (a) Die Person hat den Dialog wirklich abgelehnt → Browser-
        //      Einstellung zurücksetzen hilft.
        //  (b) Die Seite steckt in einem iframe OHNE `allow="camera"`
        //      (Teams-Registerkarte, eingebettete Web-Part-Ansicht, manche
        //      SharePoint-Rahmen). Dann wird gar nicht erst GEFRAGT — Chrome
        //      wirft denselben Fehlernamen. Wer daraufhin das Schloss-Icon
        //      sucht, findet dort keinen Kamera-Eintrag und hält das Gerät
        //      für kaputt. Der Rahmen gibt die Kamera nicht frei; helfen
        //      kann nur, die Seite direkt im Browser zu öffnen.
        //
        // Unterschieden wird über die Permissions-API: `prompt` heißt „nie
        // gefragt" und damit (b). Fehlt die API (Safari < 16), entscheidet
        // die iframe-Erkennung allein.
        let inIframe = false;
        try { inIframe = window.self !== window.top; } catch { inIframe = true; /* Cross-Origin-Zugriff wirft → sicher im iframe */ }
        let neverAsked = false;
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const perms = (navigator as any).permissions;
          if (perms?.query) {
            const st = await perms.query({ name: 'camera' });
            neverAsked = st?.state === 'prompt';
          }
        } catch { /* Permissions-API kennt 'camera' nicht — dann bleibt es bei der iframe-Erkennung */ }

        setCameraErrorInIframe(inIframe);
        if (inIframe) {
          msg = 'Diese Seite läuft in einem eingebetteten Rahmen (Teams-Registerkarte oder SharePoint-App). '
            + 'Der Rahmen gibt die Kamera nicht frei — du wurdest deshalb gar nicht erst gefragt. '
            + 'Öffne die Seite direkt in Edge (Arbeitskonto) und starte den Scan dort erneut — oder nimm den Foto-Weg, der ohne Kamera-Freigabe auskommt. '
            + 'In Teams: die drei Punkte oben rechts an der Registerkarte → „Im Browser öffnen".';
        } else if (neverAsked) {
          msg = 'Der Browser hat die Kamera blockiert, ohne zu fragen. '
            + 'Das passiert bei unsicheren Verbindungen oder wenn die Seite eingebettet ist — '
            + 'öffne sie direkt in Edge (Arbeitskonto) — oder nimm den Foto-Weg, der ohne Kamera-Freigabe auskommt.';
        } else {
          msg = 'Kamera-Berechtigung wurde abgelehnt. Bitte in den Browser-Einstellungen '
            + 'für diese Seite die Kamera erlauben und dann erneut versuchen. '
            + '(iOS Safari: aA-Icon links in der Adresszeile → Website-Einstellungen → Kamera: Erlauben. '
            + 'Android Chrome: Schloss-Icon in der Adresszeile → Berechtigungen → Kamera: Zulassen.)';
        }
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
        msg = 'Kamera-Zugriff vom Browser blockiert (vermutlich unsichere Verbindung oder eingebetteter iframe). Öffne die Seite direkt in Edge (Arbeitskonto) — oder nimm den Foto-Weg.';
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
  /**
   * v30.30: Foto-Weg. Der Live-Scanner braucht `getUserMedia` und hängt damit
   * an einer Kette, die wir nicht kontrollieren (Browserwahl → CA-Policy →
   * WebView → Kamera-Permission). Ein File-Input mit `capture="environment"`
   * ruft stattdessen die native Kamera-App auf: Die App fragt DEX nie nach
   * einer Kamera-Berechtigung, sie bekommt ein fertiges Bild.
   *
   * Ab hier ist alles identisch zum Live-Scan — dieselbe `processCode`, dieselbe
   * Ergebniskarte, dieselbe Doppel-Scan-Erkennung. Der Foto-Weg ist bewusst nur
   * eine zweite EINGABE, kein zweiter Ablauf.
   *
   * Die Bibliothek kommt aus demselben Lazy-Chunk wie beim Live-Scanner (v20.0),
   * das Bundle wächst dadurch nicht.
   */
  const handlePhotoPicked = async (file: File | null): Promise<void> => {
    setPhotoHint('');
    if (!file) {
      // Kein File trotz Klick: Auf verwalteten Geräten kann die Richtlinie den
      // Rückweg aus der Kamera-App unterbinden („Receive data from other apps").
      // Das sieht aus wie ein Abbruch durch den Nutzer und ist keiner — deshalb
      // beide Möglichkeiten nennen statt zu raten.
      setPhotoHint(isDe
        ? 'Es kam kein Foto zurück. Entweder hast du abgebrochen — oder eine Geräte-Richtlinie erlaubt die Übergabe aus der Kamera-App nicht. Dann bleibt die Suche in der Teilnehmerliste unten.'
        : 'No photo was returned. Either you cancelled — or a device policy blocks handing files over from the camera app. In that case use the participant search below.');
      return;
    }
    setPhotoBusy(true);
    try {
      const QrScannerCls = (await import('qr-scanner')).default;
      const res = await QrScannerCls.scanImage(file, { returnDetailedScanResult: true });
      await processCode(res.data);
    } catch {
      // scanImage wirft, wenn im Bild kein Code steckt — das ist der Normalfall
      // eines unscharfen oder zu weit entfernten Fotos, kein technischer Fehler.
      setPhotoHint(isDe
        ? 'Kein QR-Code im Foto erkannt. Geh näher ran, halte das Handy ruhig und achte darauf, dass der Code nicht gespiegelt oder überstrahlt ist.'
        : 'No QR code found in the photo. Move closer, hold steady, and make sure the code is not mirrored or washed out by glare.');
    } finally {
      setPhotoBusy(false);
    }
  };

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

    if (reg.Status === 'Abgemeldet') {
      setResultMessage(`${name} — ${t('checkin.cancelled')}`);
      setResultType('error');
      setIsProcessing(false);
      return;
    }
    // v30.91: Programmpunkt-Modus des GESCANNTEN Events. Der Punkt wird
    // unten beim Event gewählt; ein Scan zu einem anderen Event oder ohne
    // Punkt wird nicht geraten, sondern abgelehnt.
    const scanEv = events.find(e => e.id === event.id);
    const scanAgenda = !!(scanEv && scanEv.agendaCheckIn && (scanEv.agenda || []).length > 0);
    if (scanAgenda) {
      if (event.id !== nameSearchEventId || !agendaPoint) {
        setResultMessage(isDe
          ? `Dieses Event hat Programmpunkte — bitte unten das Event und den ${agendaTermSingular} wählen, an dem eingecheckt wird.`
          : `This event has agenda items — please pick the event and the ${agendaTermSingular.toLowerCase()} below first.`);
        setResultType('error');
        setIsProcessing(false);
        return;
      }
      const marks = parseAgendaCheckIns(reg.AgendaCheckIns);
      if (marks[agendaPoint.id]) {
        setResultMessage(`${name} — ${isDe ? 'bereits erfasst' : 'already recorded'} (${agendaPoint.title}, ${formatMarkTime(marks[agendaPoint.id].at)})`);
        setResultType('info');
        setIsProcessing(false);
        return;
      }
    } else if (reg.Status === 'Eingecheckt') {
      setResultMessage(`${name} — ${t('checkin.alreadycheckedin')}`);
      setResultType('info');
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
      event: { id: event.id, subsiteUrl: event.subsiteUrl, title: event.title },
      regId: reg.Id,
      status: reg.Status,
      agendaItemId: scanAgenda && agendaPoint ? agendaPoint.id : undefined,
      agendaLabel: scanAgenda && agendaPoint ? agendaPoint.title : undefined,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      department: (reg as any).Department || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      jobTitle: (reg as any).JobTitle || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      location: (reg as any).Location || '',
      photoUrl,
      // v30.53: auch auf dem QR-Weg — der Tisch braucht dieselben Angaben,
      // egal ob gescannt oder gesucht wurde.
      extras: extrasFor(reg, event.id),
    });
    setResultMessage('');
    setResultType('');
    setIsProcessing(false);
    // Zum Bestätigungs-Dialog scrollen
    setTimeout(() => {
      confirmCardRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 100);
  };

  // Check-in ausführen. v31.1: aus `confirmCheckIn` herausgelöst — Liste und
  // Teilnehmer-ID checken direkt ein (Nutzer 07.09.2026: „wenn ich auf
  // Einchecken klicke, soll er direkt eingecheckt sein"); nur der Scan zeigt
  // vorher die Bestätigungskarte, weil dort die Person erst identifiziert wird.
  const performCheckIn = async (pendingCheckIn: PendingCheckInInfo): Promise<void> => {
    if (!eventService) return;
    const remember = (): void => {
      setRecentCheckIns(prev => [{
        key: `${pendingCheckIn.regId}:${pendingCheckIn.agendaItemId || 'status'}:${Date.now()}`,
        at: new Date().toISOString(), name: pendingCheckIn.name, regId: pendingCheckIn.regId,
        eventId: pendingCheckIn.event.id || '', subsiteUrl: pendingCheckIn.event.subsiteUrl,
        prevStatus: pendingCheckIn.status, agendaItemId: pendingCheckIn.agendaItemId, agendaLabel: pendingCheckIn.agendaLabel,
      }, ...prev].slice(0, 30));
    };
    try {
      // v30.67: `checkInParticipant` wirft nie — es liefert `response.ok` bzw.
      // false. Der catch unten fängt also nur Unerwartetes; ob der MERGE
      // (403 ohne Schreibrecht, 429 in der Einlasswelle, 400 ohne Spalte)
      // durchging, steht NUR im Rückgabewert. Vorher stieg der Zähler und die
      // grüne Meldung kam auch dann, wenn in der Liste weiter „Angemeldet"
      // stand — und die Person tauchte später in der No-Show-Auswertung auf.
      // v30.91: Programmpunkt — nur der Punkt wird gesetzt, der Status bleibt.
      if (pendingCheckIn.agendaItemId) {
        const pointId = pendingCheckIn.agendaItemId;
        const label = pendingCheckIn.agendaLabel || '';
        const r = await eventService.checkInAgendaItem(pendingCheckIn.event.subsiteUrl, pendingCheckIn.regId, pointId);
        if (!r.ok) {
          setResultMessage(isDe
            ? `${pendingCheckIn.name} — Anwesenheit konnte nicht gespeichert werden${r.status ? ` (HTTP ${r.status})` : ''}. ${r.status === 400 ? 'Fehlt die Spalte AgendaCheckIns? Organizer: „Spalten fixen" ausführen.' : 'Bitte erneut versuchen.'}`
            : `${pendingCheckIn.name} — attendance could not be saved${r.status ? ` (HTTP ${r.status})` : ''}.`);
          setResultType('error');
          setPendingCheckIn(null);
          processingRef.current = false;
          return;
        }
        const evId = pendingCheckIn.event.id;
        const regId = pendingCheckIn.regId;
        const at = r.already || new Date().toISOString();
        if (evId) {
          setSearchRegsCache(prev => {
            const list = prev[evId];
            if (!list) return prev;
            return { ...prev, [evId]: list.map(x => x.Id === regId
              ? { ...x, AgendaCheckIns: JSON.stringify({ ...parseAgendaMarks(x.AgendaCheckIns), [pointId]: { at, by: '' } }) }
              : x) };
          });
        }
        if (!r.already) { setCheckedInCount(prev => prev + 1); remember(); }
        setResultMessage(r.already
          ? `${pendingCheckIn.name} — ${isDe ? 'bereits erfasst' : 'already recorded'} (${label}, ${formatMarkTime(r.already)})`
          : `${pendingCheckIn.name} — ${isDe ? 'anwesend bei' : 'present at'} ${label}`);
        setResultType(r.already ? 'info' : 'success');
        setPendingCheckIn(null);
        processingRef.current = false;
        return;
      }
      const ok = await eventService.checkInParticipant(pendingCheckIn.event.subsiteUrl, pendingCheckIn.regId);
      if (!ok) {
        setResultMessage(isDe
          ? `${pendingCheckIn.name} — Check-in fehlgeschlagen (bitte erneut versuchen oder Organizer informieren).`
          : `${pendingCheckIn.name} — check-in failed (please retry or inform an organizer).`);
        setResultType('error');
        setPendingCheckIn(null);
        processingRef.current = false;
        return;
      }
      setCheckedInCount(prev => prev + 1);
      remember();
      // v30.67: Den neuen Status auch in der Trefferliste nachführen — sie
      // liest nur aus dem Cache, und der wurde bisher nur beim No-Show
      // gepatcht. Ohne den Patch blieb die Person „Angemeldet", die KPI-
      // Kachel stand, und unter „Nur offene" stand sie weiter bei den Offenen;
      // ein zweiter Helfer checkte sie erneut ein.
      const evId = pendingCheckIn.event.id;
      const regId = pendingCheckIn.regId;
      if (evId) {
        setSearchRegsCache(prev => {
          const list = prev[evId];
          if (!list) return prev;
          return { ...prev, [evId]: list.map(r => r.Id === regId ? { ...r, Status: 'Eingecheckt' } : r) };
        });
      }
      setResultMessage(`${pendingCheckIn.name} — ${t('checkin.success')}`);
      setResultType('success');
    } catch {
      setResultMessage(`${pendingCheckIn.name} — Check-in fehlgeschlagen.`);
      setResultType('error');
    }
    setPendingCheckIn(null);
    processingRef.current = false;
  };
  const confirmCheckIn = (): Promise<void> => pendingCheckIn ? performCheckIn(pendingCheckIn) : Promise.resolve();

  /** v31.1: Check-in aus „Letzte Check-ins" zurücknehmen. Programmpunkt →
   *  Anwesenheit entfernen; Event-Status → zurück auf den gemerkten Stand.
   *  Das Ergebnis wird geprüft, der Cache nachgeführt, der Zähler korrigiert. */
  const undoCheckIn = async (e: RecentCheckIn): Promise<void> => {
    if (!eventService || undoBusyKey) return;
    setUndoBusyKey(e.key);
    try {
      let ok = false;
      const isNoShow = e.kind === 'noshow';
      if (e.agendaItemId) {
        // v31.2: entfernt die Marke am Punkt — Anwesenheit ODER No-Show.
        ok = (await eventService.removeAgendaCheckIn(e.subsiteUrl, e.regId, e.agendaItemId)).ok;
        if (ok && e.eventId) {
          setSearchRegsCache(prev => {
            const list = prev[e.eventId];
            if (!list) return prev;
            return { ...prev, [e.eventId]: list.map(x => {
              if (x.Id !== e.regId) return x;
              const m = parseAgendaMarks(x.AgendaCheckIns);
              delete m[e.agendaItemId as string];
              return { ...x, AgendaCheckIns: Object.keys(m).length ? JSON.stringify(m) : '' };
            }) };
          });
        }
      } else {
        ok = await eventService.revertCheckIn(e.subsiteUrl, e.regId, e.prevStatus);
        if (ok && e.eventId) {
          const back = (e.prevStatus === 'QR versendet' || (isNoShow && e.prevStatus === 'Eingecheckt')) ? e.prevStatus : 'Angemeldet';
          setSearchRegsCache(prev => {
            const list = prev[e.eventId];
            if (!list) return prev;
            return { ...prev, [e.eventId]: list.map(x => x.Id === e.regId ? { ...x, Status: back } : x) };
          });
        }
      }
      if (!ok) {
        setResultMessage(isDe ? `${e.name} — Rückgängig fehlgeschlagen, bitte erneut versuchen.` : `${e.name} — undo failed, please retry.`);
        setResultType('error');
        return;
      }
      setRecentCheckIns(prev => prev.filter(x => x.key !== e.key));
      if (!isNoShow) setCheckedInCount(prev => Math.max(0, prev - 1));
      setResultMessage(isDe
        ? `${e.name} — ${isNoShow ? 'No-Show' : 'Check-in'} zurückgenommen${e.agendaLabel ? ` (${e.agendaLabel})` : ''}.`
        : `${e.name} — ${isNoShow ? 'no-show' : 'check-in'} reverted${e.agendaLabel ? ` (${e.agendaLabel})` : ''}.`);
      setResultType('info');
    } finally { setUndoBusyKey(''); }
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

  // v23.45: gemeinsamer Karten-Renderer für Haupt- und Sub-Events im Picker.
  // `isChild` rendert eine kompaktere, eingerückte Variante.
  const renderCheckInEventCard = (ev: typeof events[number], isChild: boolean): React.ReactElement => {
    const thumb = isChild ? 60 : 84;
    return (
      <button
        key={ev.id}
        className="card"
        onClick={() => navigate('check-in', ev.id)}
        style={{
          textAlign: 'left', padding: isChild ? 12 : 14, width: '100%',
          background: isChild ? 'var(--dex-gray-50, #f7f8f9)' : '#fff',
          border: '1px solid var(--dex-gray-200)',
          borderRadius: 12, cursor: 'pointer',
          display: 'flex', alignItems: 'center', gap: 14,
          transition: 'border-color 0.15s, box-shadow 0.15s, transform 0.1s',
        }}
        onMouseEnter={e => { e.currentTarget.style.borderColor = 'var(--dex-green, #86bc25)'; e.currentTarget.style.boxShadow = '0 4px 12px rgba(134,188,37,0.12)'; }}
        onMouseLeave={e => { e.currentTarget.style.borderColor = 'var(--dex-gray-200)'; e.currentTarget.style.boxShadow = 'none'; }}
      >
        {/* v15.3: Event-Bild als Thumbnail links — Fallback grünes Gradient mit Initialen. */}
        <div style={{
          flex: '0 0 auto', width: thumb, height: thumb, borderRadius: 10,
          background: ev.imageUrl
            ? `url(${ev.imageUrl}) center/cover no-repeat`
            : 'linear-gradient(135deg, var(--dex-green, #86bc25), var(--dex-blue, #0076a8))',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          color: '#fff', fontWeight: 700, fontSize: isChild ? '1.1rem' : '1.4rem',
          overflow: 'hidden', flexShrink: 0,
        }}>
          {!ev.imageUrl && ev.title ? ev.title.charAt(0).toUpperCase() : ''}
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          {isChild && (
            <span style={{
              display: 'inline-block', fontSize: '0.66rem', fontWeight: 700, letterSpacing: 0.3,
              color: 'var(--dex-green-dark, #4a7c1f)', background: 'rgba(134,188,37,0.14)',
              borderRadius: 6, padding: '1px 7px', marginBottom: 4, textTransform: 'uppercase',
            }}>{isDe ? 'Sub-Event' : 'Sub-event'}</span>
          )}
          <div style={{ fontWeight: 700, fontSize: isChild ? '0.92rem' : '1rem', color: 'var(--dex-gray-800)', marginBottom: 2 }}>{shortSubEventTitle(ev.title)}</div>
          <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
            {ev.startDate ? new Date(ev.startDate).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : ''}
            {ev.location ? ` · ${ev.location}` : ''}
          </div>
          {/* v15.3: Organizer-Chips mit Hover-Profilfoto (gleiches OrganizerList wie auf der Registrierungsseite). */}
          {/* v29.48: „Organizer ausblenden" wurde hier nicht ausgewertet — die
              Selbst-Check-in-Seite ist teilnehmersichtbar, also gilt die
              Einstellung auch hier (gleiche Prüfung wie Anmeldeseite/Kachel). */}
          {!isChild && (ev.organizers && ev.organizers.length > 0) && !(ev.hideOrganizer && !ev.hideOrganizerIndividualOnly) && (
            <div style={{ marginTop: 6 }}>
              <OrganizerList
                names={ev.organizers}
                emails={ev.organizerEmails || []}
                hiddenEmails={(ev.hideOrganizer && ev.hideOrganizerIndividualOnly) ? (ev.hiddenOrganizerEmails || []) : []}
                size="sm"
                compact
              />
            </div>
          )}
        </div>
      </button>
    );
  };

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
            {groupedCheckInEvents.map(({ parent, children }) => {
              const expanded = !!expandedCheckInParents[parent.id];
              const hasChildren = children.length > 0;
              return (
                <div key={parent.id} style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ position: 'relative' }}>
                    {renderCheckInEventCard(parent, false)}
                    {hasChildren && (
                      <button
                        type="button"
                        onClick={e => { e.stopPropagation(); setExpandedCheckInParents(p => ({ ...p, [parent.id]: !p[parent.id] })); }}
                        title={isDe ? (expanded ? 'Sub-Events einklappen' : 'Sub-Events anzeigen') : (expanded ? 'Collapse sub-events' : 'Show sub-events')}
                        style={{
                          ...(isMobile
                            ? { marginTop: 8, alignSelf: 'flex-start' }
                            : { position: 'absolute', top: '50%', right: 14, transform: 'translateY(-50%)' }),
                          display: 'inline-flex', alignItems: 'center', gap: 6,
                          background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 999,
                          padding: '6px 12px', cursor: 'pointer', color: 'var(--dex-gray-700)',
                          fontSize: '0.78rem', fontWeight: 600,
                        }}
                      >
                        <span>{children.length} {isDe ? 'Sub-Events' : 'sub-events'}</span>
                        {expanded ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                      </button>
                    )}
                  </div>
                  {expanded && hasChildren && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginLeft: 18, paddingLeft: 18, borderLeft: '2px solid var(--dex-gray-200)' }}>
                      {children.map(ch => renderCheckInEventCard(ch, true))}
                    </div>
                  )}
                </div>
              );
            })}
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

  // v31.1: Aufbau wie das Anmeldeformular (Nutzer 07.09.2026: „oben den Kreis
  // des Event-Logos und dann kommen die sinnvollen Check-in-Sektionen") —
  // nummerierte Abschnitte, das Event mit rundem Bild zuerst. Die Reihenfolge
  // kommt über CSS `order`, damit die bestehenden Karten (Scanner, Self-
  // Check-in, Liste) nicht im JSX verschoben werden müssen.
  const heroEv = events.find(e => e.id === nameSearchEventId) || selectedEvent || null;
  const heroImg = ((): string => {
    if (!heroEv) return '';
    if (heroEv.imageUrl) return heroEv.imageUrl;
    try {
      const o = JSON.parse(heroEv.emailTemplateOverrides || '{}');
      if (o && typeof o._eventLogo === 'string' && o._eventLogo) return o._eventLogo;
    } catch { /* */ }
    return heroEv.mailImageBase64 || '';
  })();
  const heroDate = ((): string => {
    if (!heroEv || !heroEv.startDate) return '';
    const f = (iso: string): string => new Date(iso).toLocaleDateString('de-DE', { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' });
    const a = f(heroEv.startDate);
    const b = heroEv.endDate ? f(heroEv.endDate) : '';
    return b && b !== a ? `${a} – ${b}` : a;
  })();
  const sectionLabel = (n: number, label: string): React.ReactElement => (
    <div style={{ display: 'flex', alignItems: 'center', gap: 12, margin: '6px 0 10px' }}>
      <span style={{ width: 34, height: 34, borderRadius: '50%', background: 'var(--dex-green, #86bc25)', color: '#fff', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: '0.95rem', boxShadow: '0 0 0 4px rgba(134,188,37,0.18)', flexShrink: 0 }}>{n}</span>
      <span style={{ fontSize: '0.78rem', fontWeight: 700, letterSpacing: '0.08em', textTransform: 'uppercase', color: 'var(--dex-green-dark, #4a7c1f)' }}>{label}</span>
    </div>
  );
  // v31.2: Trenner OHNE Nummer für die Werkzeuge unter dem Ablauf. Nutzer
  // 07.09.2026: „das sollte nicht Schritt 4 und 5 sein, sondern einfach
  // optisch getrennt sein — sonst denkt man, das wäre chronologisch."
  const sideLabel = (label: string): React.ReactElement => (
    <div style={{ display: 'flex', alignItems: 'center', gap: 12, margin: '22px 0 10px' }}>
      <span style={{ fontSize: '0.74rem', fontWeight: 700, letterSpacing: '0.08em', textTransform: 'uppercase', color: 'var(--dex-gray-500)' }}>{label}</span>
      <span style={{ flex: 1, height: 1, background: 'var(--dex-gray-200)' }} />
    </div>
  );
  const cardToggle = (title: string, open: boolean, onToggle: () => void, hint: string): React.ReactElement => (
    <button type="button" onClick={onToggle} aria-expanded={open} style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 12, background: 'none', border: 'none', padding: 0, cursor: 'pointer', textAlign: 'left', color: 'inherit', font: 'inherit' }}>
      <h3 style={{ margin: 0 }}>{title}</h3>
      {!open && <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{hint}</span>}
      {open && <span style={{ flex: 1 }} />}
      <span style={{ color: 'var(--dex-gray-400)', display: 'inline-flex' }}>{open ? <ChevronUp size={18} /> : <ChevronDown size={18} />}</span>
    </button>
  );

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

      {/* v30.35: Der orange Warnkasten „Kamera nicht verfuegbar in der
          SharePoint App" ist raus. Er war ein Alarm fuer einen Zustand,
          der in der SharePoint-App der NORMALFALL ist — und seit v30.33
          gibt es dort einen Weg, der einfach funktioniert (Teilnehmer-ID).
          Eine Warnung, die bei jedem Start erscheint und auf etwas
          hinweist, das man nicht aendern kann, wird ueberlesen und macht
          die Seite nur enger. Die Geraete-Erwartung steht jetzt ruhig im
          Hinweistext unter dem Scan-Knopf. */}

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
              {/* v30.53: Startnummer + Trikotgröße gut sichtbar — sie werden
                  hier vorgelesen bzw. ausgegeben, nicht nur nachgeschlagen. */}
              {(pendingCheckIn.extras || []).length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginTop: 8 }}>
                  {(pendingCheckIn.extras || []).map((x, i) => (
                    <span key={i} style={{
                      display: 'inline-flex', alignItems: 'baseline', gap: 6,
                      padding: x.strong ? '5px 12px' : '4px 10px',
                      borderRadius: 8,
                      background: x.strong ? 'rgba(134,188,37,0.14)' : 'var(--dex-gray-100, #f5f5f5)',
                      border: x.strong ? '1px solid var(--dex-green, #86bc25)' : '1px solid var(--dex-gray-200)',
                    }}>
                      <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)' }}>{x.label}</span>
                      <strong style={{
                        fontSize: x.strong ? '1.15rem' : '0.9rem',
                        fontFamily: x.strong ? "'Courier New',Courier,monospace" : 'inherit',
                        letterSpacing: x.strong ? '0.04em' : undefined,
                        color: x.strong ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-800)',
                      }}>{x.value}</strong>
                    </span>
                  ))}
                </div>
              )}
            </div>
          </div>
          <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', margin: '0 0 16px' }}>
            {isDe ? 'Event: ' : 'Event: '}<strong>{pendingCheckIn.event.title}</strong>
            {pendingCheckIn.agendaLabel && (
              <><br />{agendaTermSingular}: <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>{pendingCheckIn.agendaLabel}</strong></>
            )}
          </p>
          <div style={{ display: 'flex', gap: 12 }}>
            <button
              className="btn btn-primary"
              onClick={confirmCheckIn}
              style={{ flex: 1, fontSize: '1rem', padding: '12px 0', background: 'var(--dex-green)' }}
            >
              {pendingCheckIn.agendaItemId ? (isDe ? 'Anwesenheit erfassen' : 'Record attendance') : (isDe ? 'Einchecken' : 'Check in')}
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

      <div style={{ display: 'flex', flexDirection: 'column' }}>
      {/* v31.1 — Abschnitt 1: das Event, wie auf der Anmeldeseite mit rundem Bild. */}
      <div style={{ order: 1 }}>
        {sectionLabel(1, isDe ? 'Event' : 'Event')}
        <div className="card" style={{ padding: heroImg ? '86px 24px 20px' : 24, marginBottom: 16, marginTop: heroImg ? 72 : 0, position: 'relative', overflow: 'visible' }}>
          {heroImg && (
            <div style={{ position: 'absolute', top: -72, left: '50%', transform: 'translateX(-50%)', width: 144, height: 144, borderRadius: '50%', background: '#fff', boxShadow: '0 8px 26px rgba(0,0,0,0.14)', padding: 6 }}>
              <img src={heroImg} alt={heroEv ? heroEv.title : ''} style={{ width: '100%', height: '100%', borderRadius: '50%', objectFit: 'cover', display: 'block' }} onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }} />
            </div>
          )}
          {heroEv ? (
            <div style={{ textAlign: 'center', marginBottom: (accessibleEvents.length > 1 || agendaMode) ? 14 : 0 }}>
              <div style={{ fontSize: '1.15rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>{heroEv.title}</div>
              <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                {heroDate}{heroEv.location ? ` · ${heroEv.location}` : ''}
              </div>
            </div>
          ) : (
            <p style={{ margin: 0, textAlign: 'center', color: 'var(--dex-gray-500)', fontSize: '0.9rem' }}>{isDe ? 'Bitte ein Event wählen.' : 'Please pick an event.'}</p>
          )}
          {accessibleEvents.length > 1 && (
            <select
              className="form-input"
              value={nameSearchEventId}
              onChange={e => { setNameSearchEventId(e.target.value); setNameSearchQuery(''); }}
              style={{ marginBottom: agendaMode ? 12 : 0, padding: '8px 12px', fontSize: '0.9rem', width: '100%' }}
            >
              <option value="">— Event auswählen —</option>
              {accessibleEvents.map(ev => (
                <option key={ev.id} value={ev.id}>{ev.title}</option>
              ))}
            </select>
          )}
          {/* v30.91: Programmpunkt wählen — hier wird eingecheckt. Vorschlag ist
              der Punkt, der gerade läuft (suggestCurrentAgendaItem); die Wahl
              bleibt je Event im Gerät gespeichert. */}
          {agendaMode && (
            <div style={{ padding: '10px 12px', borderRadius: 10, background: 'rgba(134,188,37,0.08)', border: '1px solid rgba(134,188,37,0.45)' }}>
              <div style={{ fontSize: '0.8rem', fontWeight: 700, marginBottom: 6 }}>
                {isDe ? `${agendaTermSingular} wählen — hier wird eingecheckt` : `Pick the ${agendaTermSingular.toLowerCase()} — attendance is recorded there`}
              </div>
              {/* v30.94: Chips je Cluster (utils/agendaGroups). */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {agendaGroups(agendaItems).map((grp, gi) => (
                  <div key={grp.key}>
                    {(agendaGroups(agendaItems).length > 1 || grp.cluster) && (
                      <div style={{ fontSize: '0.74rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', marginBottom: 4 }}>
                        {groupLabel(grp, gi, isDe)} <span style={{ fontWeight: 500, color: 'var(--dex-gray-500)' }}>· {groupDateLabel(grp, isDe, false)}</span>
                      </div>
                    )}
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                      {grp.items.map(it => {
                        const active = it.id === agendaPointId;
                        const cnt = agendaCounts[it.id] || 0;
                        return (
                          <button key={it.id} type="button" onClick={() => choosePoint(it.id)} style={{
                            display: 'inline-flex', alignItems: 'center', gap: 6, padding: '6px 10px', borderRadius: 999, cursor: 'pointer',
                            border: `1px solid ${active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                            background: active ? 'var(--dex-green, #86bc25)' : '#fff', color: active ? '#fff' : 'var(--dex-gray-800)',
                            fontSize: '0.8rem', fontWeight: active ? 700 : 500, textAlign: 'left',
                          }}>
                            <span style={{ fontVariantNumeric: 'tabular-nums', opacity: 0.85 }}>{it.time}</span>
                            <span>{it.title || (isDe ? '(ohne Titel)' : '(untitled)')}</span>
                            <span style={{ fontSize: '0.72rem', padding: '1px 7px', borderRadius: 999, background: active ? 'rgba(255,255,255,0.25)' : 'var(--dex-gray-100)', color: active ? '#fff' : 'var(--dex-gray-600)' }}>{cnt}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>
              <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)', marginTop: 6 }}>
                {isDe
                  ? 'Jeder Scan, jede ID und jeder Klick setzt nur diesen Punkt. Der Event-Status der Person bleibt unverändert.'
                  : 'Every scan, ID and click records only this item. The person\'s event status stays unchanged.'}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* v31.1 — Live-Scanner, eingeklappt bis gebraucht. v31.2: ohne Nummer —
          Scanner und Self-Check-in sind Werkzeuge neben dem Ablauf, keine
          Schritte danach. */}
      <div style={{ order: 4 }}>
      {sideLabel(isDe ? 'Weitere Wege zum Einchecken' : 'Other ways to check in')}
      {/* Live-Scanner — Kamerabild + Steuerung */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        {!isScanning ? (
          <>
            {cardToggle(isDe ? 'Live-Scanner' : 'Live scanner', scannerOpen, () => setScannerOpen(o => !o), isDe ? 'QR-Code mit der Kamera scannen oder Foto vom Code machen' : 'Scan the QR code with the camera or take a photo of it')}
            {scannerOpen && (
            <div style={{ textAlign: 'center', marginTop: 14 }}>
              {/* v30.35: EIN Knopf statt zwei. Zwei gleich große Scan-Knöpfe
                  nebeneinander sind zwei Bedienwege für dieselbe Absicht — man
                  muss am Einlass erst entscheiden, statt zu scannen. Der
                  Live-Scanner ist der Standard; der Foto-Weg steht als schmaler
                  Text-Link darunter, für den Fall, dass der Scanner nicht
                  aufgeht. Wenn beides scheitert, trägt die Teilnehmer-ID
                  (v30.33) — die steht im Hinweis darunter. */}
              <button className="btn btn-primary" onClick={startCamera} style={{ fontSize: '1.1rem', padding: '14px 36px' }}>
                {t('checkin.scan')}
              </button>
              <div style={{ marginTop: 10 }}>
                <button
                  type="button"
                  onClick={() => photoInputRef.current?.click()}
                  disabled={photoBusy}
                  style={{
                    background: 'none', border: 'none', padding: 4, cursor: photoBusy ? 'default' : 'pointer',
                    color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.88rem', textDecoration: 'underline',
                    font: 'inherit', fontFamily: 'inherit',
                  }}
                >
                  {photoBusy
                    ? (isDe ? 'Foto wird gelesen…' : 'Reading photo…')
                    : (isDe ? 'Stattdessen Foto vom QR-Code machen' : 'Take a photo of the QR code instead')}
                </button>
              </div>

              {/* v31.1: Das Teilnehmer-ID-Feld steht jetzt in Abschnitt 2
                  „Einchecken" — dort, wo auch die Liste ist. */}
              {/* v30.30: `capture="environment"` öffnet die NATIVE Kamera-App des
                  Geräts und liefert eine Datei zurück — ohne getUserMedia, ohne
                  Kamera-Berechtigung für die Seite. Damit ist es der einzige
                  Scan-Weg, der weder an der WebView-Sperre der SharePoint-App
                  noch an der iframe-Freigabe einer Teams-Registerkarte hängt.
                  Bewusst KEIN `multiple` — eine Person pro Foto.

                  v30.31: `accept` wird auf Android erweitert. Seit Android 14
                  leiten Chrome und Edge ein reines `accept="image/*"` in den
                  System-Foto-Picker um — und der hat gar keinen Kamera-Eintrag,
                  auch mit `capture` nicht. Ein zusätzlicher, nicht-standardisierter
                  Wert bricht diese Umleitung und bringt die Kamera zurück (der
                  dokumentierte Workaround, s. Release Notes). Auf iOS bleibt es
                  beim sauberen `image/*` — dort greift `capture` normal, und ein
                  unbekannter MIME-Wert wäre nur ein Risiko ohne Nutzen. */}
              <input
                ref={photoInputRef}
                type="file"
                accept={isAndroid ? 'image/*,android/allowCamera' : 'image/*'}
                capture="environment"
                style={{ display: 'none' }}
                onChange={e => { void handlePhotoPicked(e.target.files?.[0] || null); e.target.value = ''; }}
              />
              {photoHint && (
                <p style={{ color: 'var(--dex-orange)', fontSize: '0.85rem', margin: '10px 0 0' }}>{photoHint}</p>
              )}
              <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', margin: '10px 0 0' }}>
                {/* v30.33: Erwartung ehrlich setzen statt jeden erst scheitern
                    lassen. Auf iPhones läuft der Scanner in aller Regel; auf
                    Android hängt er an Dingen, die wir nicht kontrollieren
                    (WebView der SharePoint-App, Foto-Picker ab Android 14).
                    Deshalb dort gleich den Weg nennen, der immer funktioniert. */}
                {isDe
                  ? <>Auf <strong>iPhones</strong> funktioniert der Live-Scanner in der Regel zuverlässig.{isAndroid ? <> Auf <strong>Android</strong> ist er oft blockiert — dann tippe die <strong>Teilnehmer-ID</strong> unten ins Suchfeld: Sie steht in jeder QR-Mail unter dem Code und funktioniert immer.</> : <> Klappt er nicht, tippe die <strong>Teilnehmer-ID</strong> unten ins Suchfeld — sie steht in jeder QR-Mail unter dem Code.</>} Der Foto-Weg benutzt die normale Kamera-App und kommt ohne Kamera-Freigabe für die Seite aus.{isAndroid ? <> Zeigt Android eine Auswahl statt der Kamera, tippe dort auf <strong>Kamera</strong>.</> : null}</>
                  : <>On <strong>iPhones</strong> the live scanner usually works reliably.{isAndroid ? <> On <strong>Android</strong> it is often blocked — then type the <strong>attendee ID</strong> into the search field below: it is printed under the code in every QR email and always works.</> : <> If it does not start, type the <strong>attendee ID</strong> into the search field below — it is printed under the code in every QR email.</>} The photo route uses the regular camera app and needs no camera permission for the page.{isAndroid ? <> If Android shows a chooser instead of the camera, pick <strong>Camera</strong> there.</> : null}</>}
              </p>
              {cameraError && (
                <div style={{ marginTop: 12 }}>
                  <p style={{ color: 'var(--dex-orange)', fontSize: '0.85rem', margin: 0 }}>{cameraError}</p>
                  {/* v30.29: Steckt die Seite im iframe, ist „direkt im Browser
                      öffnen" die einzige Abhilfe — dann auch einen Knopf dafür
                      anbieten statt nur den Weg zu beschreiben. window.open aus
                      dem iframe erzeugt einen Top-Level-Tab, dort greift die
                      Kamera-Freigabe der Seite selbst. */}
                  {cameraErrorInIframe && (
                    <button
                      className="btn btn-secondary"
                      style={{ marginTop: 10, fontSize: '0.85rem' }}
                      onClick={() => window.open(window.location.href, '_blank', 'noopener')}
                    >
                      Seite im Browser öffnen
                    </button>
                  )}
                  <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)', marginTop: 8, fontWeight: 600 }}>
                    Alternativ: Du kannst Teilnehmer auch in der Liste oben suchen
                    und per Klick auf &quot;Einchecken&quot; manuell einchecken.
                  </p>
                </div>
              )}
            </div>
            )}
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

      </div>

      {/* v31.1 — Self-Check-in, eingeklappt bis gebraucht (v31.2: ohne
          Nummer, direkt unter dem Live-Scanner). */}
      <div style={{ order: 5 }}>
      {/* v20.1: Self-Check-in — prominent direkt unter dem Live-Scanner.
          Teilnehmer scannen den Event-QR mit der NATIVEN Handy-Kamera (kein
          Kamera-Zugriff in der App nötig) und checken sich selbst ein. */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        {cardToggle('Self-Check-in', selfOpen, () => setSelfOpen(o => !o), isDe ? 'Live-QR für einen Bildschirm am Eingang oder PDF zum Aushängen' : 'Live QR for a screen at the entrance or a printable PDF')}
        {selfOpen && (<>
        <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', margin: '12px 0 14px' }}>
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
        </>)}
      </div>
      </div>

      {/* v31.1 — Abschnitt 2: Einchecken (ID, Suche, Liste). Event-Auswahl und
          Programmpunkt stehen jetzt in Abschnitt 1. */}
      <div style={{ order: 2 }}>
      {sectionLabel(2, isDe ? 'Einchecken' : 'Check in')}
      {/* v7.14: Live-Teilnehmerliste mit Foto / Position / Standort + Filter.
          Liste wird sofort beim Auswählen des Events geladen, kein "ab 2
          Zeichen tippen" mehr — der Helfer sieht direkt alle Leute, kann den
          gesuchten Eintrag scrollen oder das Suchfeld zum Filtern nutzen. */}
      <div className="card" style={{ padding: 24, marginBottom: 16 }}>
        {/* v30.35/v31.1: Die Teilnehmer-ID steht in der Mail groß unter dem
            QR-Code — also gehört sie hier genauso groß hin. „Einchecken" checkt
            direkt ein (kein zweiter Dialog). Die Nummer filtert die Liste
            darunter live mit (v30.87). */}
        <form
          onSubmit={e => { e.preventDefault(); void checkInByParticipantId(); }}
          style={{ display: 'flex', gap: 8, justifyContent: 'center', alignItems: 'center', flexWrap: 'wrap', marginBottom: 6 }}
        >
          <label htmlFor="dex-checkin-id" style={{ fontSize: '0.78rem', letterSpacing: '0.06em', textTransform: 'uppercase', color: 'var(--dex-gray-500)', flex: '1 1 100%', textAlign: 'center' }}>
            {isDe ? 'Teilnehmer-ID aus der QR-Mail' : 'Attendee ID from the QR email'}
          </label>
          <input
            id="dex-checkin-id"
            value={idInput}
            onChange={e => { const v = e.target.value.replace(/\D/g, ''); setIdInput(v); setIdError(''); setNameSearchQuery(v); }}
            inputMode="numeric"
            pattern="[0-9]*"
            placeholder="17"
            disabled={!nameSearchEventId}
            aria-label={isDe ? 'Teilnehmer-ID' : 'Attendee ID'}
            style={{
              width: 130, padding: '12px 14px', textAlign: 'center',
              fontFamily: "'Courier New', Courier, monospace", fontSize: '1.5rem', fontWeight: 700,
              border: '2px solid var(--dex-gray-300)', borderRadius: 10,
            }}
          />
          <button type="submit" className="btn btn-primary" disabled={!idInput.trim() || isProcessing} style={{ fontSize: '1rem', padding: '12px 24px' }}>
            {agendaMode ? (isDe ? 'Anwesend erfassen' : 'Record') : (isDe ? 'Einchecken' : 'Check in')}
          </button>
        </form>
        {idError && (
          <p style={{ color: 'var(--dex-orange)', fontSize: '0.85rem', margin: '4px 0 10px', textAlign: 'center' }}>{idError}</p>
        )}
        <div style={{ borderTop: '1px solid var(--dex-gray-200)', margin: '12px 0 14px' }} />
        {/* v7.16: KPI-Bereich — angemeldet vs. eingecheckt, plus Quick-Filter
            "Nur offene anzeigen" der die Liste auf Angemeldet/QR versendet
            reduziert. Sichtbar nur wenn Event gewählt + Liste geladen. */}
        {nameSearchEventId && (searchRegsCache[nameSearchEventId] || []).length > 0 && (
          <div style={{
            display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(96px, 1fr))', gap: 8,
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
                {agendaMode ? (agendaPoint ? `${isDe ? 'Anwesend' : 'Present'} · ${agendaPoint.title}` : (isDe ? 'Anwesend' : 'Present')) : 'Eingecheckt'}
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
            placeholder="Teilnehmer-ID, Vorname, Nachname oder E-Mail…"
            inputMode="text"
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
            {/* v30.67: Ohne Cache-Eintrag ist ein zweiter Versuch möglich — der
                Knopf erspart den Seiten-Reload, der vorher der einzige Weg war. */}
            {nameSearchEventId && !isLoadingSearchRegs && (
              <button
                type="button"
                className="btn btn-secondary"
                style={{ marginLeft: 8, fontSize: '0.74rem', padding: '2px 8px' }}
                onClick={() => { void loadRegsForSearch(nameSearchEventId); }}
              >
                {isDe ? 'Erneut laden' : 'Reload'}
              </button>
            )}
          </p>
        )}
        {/* v30.67: Bei einem Ladefehler gibt es keine Liste — also auch kein
            „Keine Teilnehmer für dieses Event", das wäre eine Aussage über
            Daten, die nie gelesen wurden. */}
        {nameSearchEventId && !isLoadingSearchRegs && !searchLoadError && (
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
                    // v30.91: Im Programmpunkt-Modus heißt „drin": am gewählten Punkt erfasst.
                    const alreadyIn = agendaMode ? presentAt(reg, agendaPointId) : status === 'Eingecheckt';
                    const cancelled = status === 'Abgemeldet';
                    const waitlist = status === 'Warteliste';
                    // v31.2: Im Programmpunkt-Modus zählt der No-Show AM PUNKT
                    // (Marke) — der Event-Status bleibt sichtbar, wenn er es ist.
                    const noShowAtPoint = agendaMode && !!agendaPointId && !!parseAgendaNoShows(reg.AgendaCheckIns)[agendaPointId];
                    const noShow = status === 'No-Show' || noShowAtPoint;
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
                          display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap',
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
                          {/* v30.53: Startnummer + Trikotgröße schon in der
                              Trefferliste — beim B2Run wird beides am selben
                              Tisch gebraucht wie der Check-in selbst. */}
                          {(() => {
                            const ex = extrasFor(reg, nameSearchEventId);
                            if (ex.length === 0) return null;
                            return (
                              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 4 }}>
                                {ex.map((x, i) => (
                                  <span key={i} style={{
                                    display: 'inline-flex', alignItems: 'baseline', gap: 5,
                                    fontSize: '0.7rem', padding: '2px 8px', borderRadius: 6,
                                    background: x.strong ? 'rgba(134,188,37,0.14)' : 'var(--dex-gray-100, #f5f5f5)',
                                    color: 'var(--dex-gray-700)',
                                  }}>
                                    <span style={{ color: 'var(--dex-gray-500)' }}>{x.label}</span>
                                    <strong style={{
                                      fontFamily: x.strong ? "'Courier New',Courier,monospace" : 'inherit',
                                      color: x.strong ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-800)',
                                    }}>{x.value}</strong>
                                  </span>
                                ))}
                              </div>
                            );
                          })()}
                        </div>
                        <span style={{
                          fontSize: '0.7rem', padding: '3px 8px', borderRadius: 999,
                          background: statusBg, color: statusFg, fontWeight: 600, whiteSpace: 'nowrap',
                        }}>{noShowAtPoint && status !== 'No-Show' ? (isDe ? 'No-Show hier' : 'No-show here') : status}</span>
                        {/* v23.28: Check-in UND No-Show nebeneinander. */}
                        <div style={{ display: 'flex', gap: 6, flexShrink: 0, flex: isMobile ? '1 1 100%' : undefined }}>
                          <button
                            type="button"
                            className="btn btn-primary"
                            style={{ fontSize: '0.78rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
                            disabled={alreadyIn || cancelled || isProcessing}
                            onClick={() => startManualCheckInFromSearch(reg)}
                          >
                            {alreadyIn ? (agendaMode ? (isDe ? '✓ Anwesend' : '✓ Present') : '✓ Eingecheckt') : cancelled ? 'Abgemeldet' : (agendaMode ? (isDe ? 'Anwesend erfassen' : 'Record') : 'Einchecken')}
                          </button>
                          <button
                            type="button"
                            className="btn btn-secondary"
                            style={{ fontSize: '0.78rem', padding: '6px 12px', whiteSpace: 'nowrap', color: 'var(--dex-gray-700, #444)' }}
                            disabled={cancelled || noShow || alreadyIn || isProcessing}
                            onClick={() => { void markNoShowFromSearch(reg); }}
                            title={isDe
                              ? (agendaMode ? 'Nicht erschienen — für diesen Punkt oder das ganze Event' : 'Teilnehmer als nicht erschienen markieren')
                              : (agendaMode ? 'No-show — for this point or the whole event' : 'Mark attendee as a no-show')}
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
      </div>

      {/* v31.1 — Abschnitt 3: Letzte Check-ins dieser Sitzung, mit Rückgängig
          (Nutzer 07.09.2026). Nur was hier eingecheckt wurde — der Revert setzt
          den gemerkten vorherigen Status bzw. entfernt die Punkt-Anwesenheit. */}
      <div style={{ order: 3 }}>
        {sectionLabel(3, isDe ? 'Letzte Check-ins' : 'Recent check-ins')}
        <div className="card" style={{ padding: recentCheckIns.length ? '14px 20px' : '14px 20px', marginBottom: 16 }}>
          {recentCheckIns.length === 0 ? (
            <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
              {isDe ? 'Noch kein Check-in in dieser Sitzung. Jeder Check-in erscheint hier und lässt sich zurücknehmen.' : 'No check-in in this session yet. Every check-in appears here and can be reverted.'}
            </p>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4, maxHeight: 260, overflowY: 'auto' }}>
              {recentCheckIns.map(e => (
                <div key={e.key} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '6px 8px', borderRadius: 8, background: 'var(--dex-gray-50, #fafafa)' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontVariantNumeric: 'tabular-nums', width: 44, flexShrink: 0 }}>{formatMarkTime(e.at)}</span>
                  <span style={{ fontWeight: 600, fontSize: '0.88rem', flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {e.name}
                    {e.agendaLabel && <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}> · {e.agendaLabel}</span>}
                  </span>
                  {/* v31.2: No-Shows stehen mit in der Liste — als graue Pille erkennbar. */}
                  {e.kind === 'noshow' && (
                    <span className="dex-ui-pill dex-ui-pill--gray">No-Show</span>
                  )}
                  <button type="button" className="btn btn-secondary" disabled={!!undoBusyKey} onClick={() => { void undoCheckIn(e); }} style={{ fontSize: '0.76rem', padding: '4px 10px', whiteSpace: 'nowrap' }}
                    title={isDe
                      ? (e.agendaItemId
                        ? (e.kind === 'noshow' ? 'No-Show an diesem Punkt entfernen' : 'Anwesenheit an diesem Punkt entfernen')
                        : `Status zurück auf „${(e.prevStatus === 'QR versendet' || (e.kind === 'noshow' && e.prevStatus === 'Eingecheckt')) ? e.prevStatus : 'Angemeldet'}“`)
                      : (e.kind === 'noshow' ? 'Revert this no-show' : 'Revert this check-in')}>
                    {undoBusyKey === e.key ? '…' : (isDe ? 'Rückgängig' : 'Undo')}
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
      </div>

      {/* v7.14: Die alte "Manuell QR-Code als String tippen"-Card ist raus.
          Die Live-Teilnehmerliste oben deckt das Manuelle Einchecken
          benutzerfreundlicher ab. */}

      {/* v31.2: Rückfrage im Programmpunkt-Modus — Nutzer 07.09.2026: „man soll
          gefragt werden, ob das für das ganze Event No-Show ist oder nur für
          den Programmpunkt". Zwei Kacheln, jede sagt, was sie tut. */}
      <Modal
        open={!!noShowAsk}
        onClose={() => setNoShowAsk(null)}
        maxWidth={520}
        title={isDe ? `${noShowAsk ? noShowAsk.name : ''} — nicht erschienen` : `${noShowAsk ? noShowAsk.name : ''} — no-show`}
        subtitle={isDe ? 'Wofür gilt der No-Show?' : 'What does the no-show apply to?'}
        footer={<button type="button" className="btn btn-secondary" onClick={() => setNoShowAsk(null)}>{isDe ? 'Abbrechen' : 'Cancel'}</button>}
      >
        <div className="dex-ui-stack">
          <button
            type="button"
            className="dex-ui-choice"
            onClick={() => { const a = noShowAsk; setNoShowAsk(null); if (a) void applyPointNoShow(a.reg, a.name); }}
          >
            <span className="dex-ui-choice-body">
              <span className="dex-ui-choice-title">{isDe ? `Nur für „${agendaPoint ? agendaPoint.title : ''}“` : `Only for “${agendaPoint ? agendaPoint.title : ''}”`}</span>
              <span className="dex-ui-choice-desc">
                {isDe
                  ? `Die Person fehlt bei diesem ${agendaTermSingular}. Ihr Event-Status bleibt, beim nächsten ${agendaTermSingular} kann sie wieder erfasst werden.`
                  : `The person is missing at this ${agendaTermSingular.toLowerCase()}. The event status stays; they can be recorded at the next one.`}
              </span>
            </span>
          </button>
          <button
            type="button"
            className="dex-ui-choice"
            onClick={() => { const a = noShowAsk; setNoShowAsk(null); if (a) void applyEventNoShow(a.reg, a.name); }}
          >
            <span className="dex-ui-choice-body">
              <span className="dex-ui-choice-title">{isDe ? 'Für das ganze Event' : 'For the whole event'}</span>
              <span className="dex-ui-choice-desc">
                {isDe
                  ? 'Der Event-Status wird auf „No-Show“ gesetzt — die Person gilt für das gesamte Event als nicht erschienen.'
                  : 'The event status is set to “No-Show” — the person counts as absent for the entire event.'}
              </span>
            </span>
          </button>
          <p className="dex-ui-muted" style={{ margin: 0 }}>
            {isDe ? 'Beides lässt sich unter „Letzte Check-ins“ zurücknehmen.' : 'Both can be reverted under “Recent check-ins”.'}
          </p>
        </div>
      </Modal>
    </div>
  );
}
