/**
 * Meine Events - zeigt alle Events für die der User registriert ist.
 * Lädt Registrierungen aus den jeweiligen Teilnehmerlisten.
 * Ermöglicht Abmeldung mit Zwei-Schritt-Bestätigung.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import OrganizerList from './OrganizerList';
import { CachedImg } from './CachedImage';

import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useRoles } from '../context/RoleContext';
import { UserFieldPicker } from './UserFieldPicker';
import { useCurrentUser } from '../context/UserContext';
// v22.13: Sub-Events in „Meine Events" nach ihrer EIGENEN Sichtbarkeit filtern
// (gleiche Logik wie Anmeldeseite/Event-Liste).
import { isEventVisibleForUser } from './EventListPage';
import { DeloitteEvent, EventSpecificField, AgendaItem, TransferTime } from '../types';
import { SPRegistration, EventCommRow } from '../services/EventService';
import { wrapTemplate } from '../services/EmailTemplates';
import { isEventOver } from '../utils/eventFormat';
import { selfCancelLocked, selfCancelLockReason } from '../utils/cancelPolicy';
import { useLanguage } from '../context/LanguageContext';
// v20.4: moderne Confirm-/Alert-Modals statt window.confirm/alert.
import { useDialog } from '../context/DialogContext';
// v11.99: RefreshCw nicht mehr benötigt (Page-Level-Refresh-Button entfernt).
import { X, Pencil, QrCode, Mail } from './Icons';
import Modal from './Modal';
import InternationalSearchToggle from './InternationalSearchToggle';
import { buildDemoShowcaseEvents, buildDemoMyRegistration } from '../services/demoShowcaseEvent';
import { TeamsJoinButton } from './TeamsJoinButton';
import { eventTeamsLink, locationWithoutTeamsUrl } from '../utils/teamsLink';
import { formatAllDayPeriod } from '../utils/eventFormat';
import { dlog } from '../utils/debugLog';
import { FieldAnswerTag, MyEventEntry, formatDate, formatDateRange, getStatusBadgeClass, getStatusLabel } from './myEvents/myEventsHelpers';
import DocumentsViewer from './myEvents/DocumentsViewer';
import QuizPlayer from './myEvents/QuizPlayer';
import MyEventSubEvents from './myEvents/MyEventSubEvents';
import MyEventDocField from './myEvents/MyEventDocField';
import MyEventUpload from './myEvents/MyEventUpload';
import CancelledEventsCollapsible from './myEvents/CancelledEventsCollapsible';
import MyEventCard from './myEvents/MyEventCard';
import { AddMemberModal, ManageTeamModal, CascadeCancelModal, MyQrModal, EventCommsModal } from './myEvents/MyEventsModals';

export default function MyEventsPage(): React.ReactElement {
  const { navigate, selectedEventId, navIntent, clearIntent } = useNavigation();
  const { topLevelEvents, childEventsOf, isEventsLoading, ensureEventDocuments, getMyRegistration, getMyEventNumbers, cancelRegistration, cancelTeamMember, updateMyRegistration, switchSplitGroup, listMyEventAttachments, uploadMyEventAttachment, deleteMyEventAttachment, uploadFieldDocument, listFieldDocuments, deleteFieldDocument, registerForEvent, getAllRegistrations, getTeamMembers, addTeamMember, listTeamJoinRequestsForEvent, decideTeamJoinRequest, getMyAssistantLinks, requestAssistantChange, resolveAssistantRequest, getEventComms } = useEvents();
  const { currentUser } = useCurrentUser();
  const currentUserEmail = (currentUser?.email || '').toLowerCase();
  // v24.41: Assistenz-Verknüpfungen — INFO-Ansicht für Anmeldungen, die jemand
  // ANDERES verwaltet (proxy: jemand hat MICH angemeldet → ich sehe nur Info).
  const [assistantLinks, setAssistantLinks] = React.useState<import('../services/EventService').AssistantLink[]>([]);
  React.useEffect(() => {
    let cancelled = false;
    getMyAssistantLinks().then(list => { if (!cancelled) setAssistantLinks(list || []); }).catch(() => { /* */ });
    return () => { cancelled = true; };
  }, [getMyAssistantLinks]);
  // INFO-Anmeldungen (ich = angemeldete Person, aber jemand anderes ist Owner).
  const infoAsParticipant = (assistantLinks || []).filter(l =>
    (l.participantEmail || '').toLowerCase() === currentUserEmail
    && (l.ownerEmail || '').toLowerCase() !== currentUserEmail
    && l.status === 'Active');
  // Offene Anforderungen AN MICH (ich verwalte die Anmeldung, jemand bittet um
  // Änderung/Abmeldung).
  const openRequestsToMe = (assistantLinks || []).filter(l =>
    (l.ownerEmail || '').toLowerCase() === currentUserEmail
    && l.status === 'Active' && l.requestStatus === 'Open' && !!l.requestType);
  const reloadAssistantLinks = React.useCallback(() => {
    getMyAssistantLinks().then(list => setAssistantLinks(list || [])).catch(() => { /* */ });
  }, [getMyAssistantLinks]);
  // Anforderung stellen (als Info-Empfänger) bzw. erledigen (als Owner).
  const submitAssistantRequest = async (link: import('../services/EventService').AssistantLink, type: 'change' | 'cancel'): Promise<void> => {
    const note = await promptDialog(
      isDe ? (type === 'cancel' ? 'Abmeldung anfordern — kurze Anmerkung (optional):' : 'Änderung anfordern — was soll geändert werden?')
           : (type === 'cancel' ? 'Request cancellation — short note (optional):' : 'Request change — what should change?'),
      { defaultValue: '' });
    if (note === null) return;
    const ok = await requestAssistantChange(link, type, note || '');
    if (ok) { showAlert(isDe ? 'Anforderung gesendet — die verwaltende Person bekommt eine Mail mit Direktlink.' : 'Request sent — the managing person receives a mail with a direct link.', { variant: 'success' }); reloadAssistantLinks(); }
    else showAlert(isDe ? 'Anforderung konnte nicht gesendet werden.' : 'Request could not be sent.', { variant: 'error' });
  };
  const resolveReq = async (link: import('../services/EventService').AssistantLink, decision: 'Done' | 'Rejected'): Promise<void> => {
    const ok = await resolveAssistantRequest(link.id, decision);
    if (ok) { showAlert(isDe ? (decision === 'Done' ? 'Als erledigt markiert.' : 'Anforderung abgelehnt.') : (decision === 'Done' ? 'Marked as done.' : 'Request rejected.'), { variant: 'success' }); reloadAssistantLinks(); }
  };
  // v11.82: Team-Mitglieder pro Event-Karte cachen — Lazy-Load via getTeamMembers.
  // Key = `${eventId}|${teamId}`. Belastet das initiale loadMyRegistrations
  // nicht — nur für Events mit gesetzter TeamId im eigenen Eintrag.
  const [teamMembersCache, setTeamMembersCache] = React.useState<Record<string, SPRegistration[]>>({});
  const teamCacheKeyRef = React.useRef<Set<string>>(new Set());
  const enqueueTeamFetch = React.useCallback((eventId: string, teamId: string): void => {
    if (!eventId || !teamId) return;
    const key = `${eventId}|${teamId}`;
    if (teamCacheKeyRef.current.has(key)) return;
    teamCacheKeyRef.current.add(key);
    getTeamMembers(eventId, teamId).then(list => {
      setTeamMembersCache(prev => ({ ...prev, [key]: list }));
    }).catch(() => {
      setTeamMembersCache(prev => ({ ...prev, [key]: [] }));
    });
  }, [getTeamMembers]);
  // v11.99: Page-Level-Refresh-State entfernt — Header-Button übernimmt.

  // v11.83: Add-Member-Modal + Join-Requests-Cache + Helpers.
  // searchUsers wird für den Add-Member-Picker gebraucht (gleiche API wie
  // im Registrierungs-Formular).
  const { searchUsers, searchUser, isImpersonating, isAdmin } = useRoles();
  const [addMemberDialog, setAddMemberDialog] = React.useState<{
    eventId: string;
    teamId: string;
    teamName: string;
    freeSlots: number;
  } | null>(null);
  const [addMemberPick, setAddMemberPick] = React.useState<{ email: string; displayName: string } | null>(null);
  const [addMemberQuery, setAddMemberQuery] = React.useState('');
  const [addMemberResults, setAddMemberResults] = React.useState<Array<{ email: string; displayName: string }>>([]);
  const [addMemberSearching, setAddMemberSearching] = React.useState(false);
  const [addMemberConsent, setAddMemberConsent] = React.useState(false);
  const [addMemberBusy, setAddMemberBusy] = React.useState(false);
  const [addMemberError, setAddMemberError] = React.useState('');
  const [addMemberIncludeIntl, setAddMemberIncludeIntl] = React.useState(false);
  const addMemberQueryTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  React.useEffect(() => {
    const q = addMemberQuery.trim();
    if (q.length >= 2 && !addMemberPick) {
      (async () => {
        setAddMemberSearching(true);
        try {
          const res = await searchUsers(q, addMemberIncludeIntl);
          setAddMemberResults(res.map(r => ({ email: r.email, displayName: r.displayName })));
        } catch { setAddMemberResults([]); }
        setAddMemberSearching(false);
      })().catch(() => {});
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [addMemberIncludeIntl]);
  const closeAddMemberDialog = (): void => {
    setAddMemberDialog(null);
    setAddMemberPick(null);
    setAddMemberQuery('');
    setAddMemberResults([]);
    setAddMemberConsent(false);
    setAddMemberError('');
    setAddMemberBusy(false);
  };
  const submitAddMember = async (): Promise<void> => {
    if (!addMemberDialog || !addMemberPick || !addMemberConsent || addMemberBusy) return;
    setAddMemberBusy(true);
    setAddMemberError('');
    try {
      const res = await addTeamMember(
        addMemberDialog.eventId,
        addMemberDialog.teamId,
        addMemberDialog.teamName || undefined,
        addMemberPick
      );
      if (!res.ok) {
        if (res.reason && res.reason.startsWith('already-registered')) {
          setAddMemberError(isDe
            ? 'Diese Person ist bereits beim Event angemeldet — bitte abmelden lassen, bevor du sie zum Team hinzufügst.'
            : 'This person is already registered for the event — please have them cancel first before adding to the team.');
        } else if (res.reason === 'team-full') {
          setAddMemberError(isDe ? 'Das Team ist bereits voll.' : 'The team is already full.');
        } else {
          setAddMemberError(isDe ? 'Hinzufügen fehlgeschlagen.' : 'Adding failed.');
        }
        setAddMemberBusy(false);
        return;
      }
      // Team-Cache invalidieren, damit das Badge neu lädt.
      const key = `${addMemberDialog.eventId}|${addMemberDialog.teamId}`;
      teamCacheKeyRef.current.delete(key);
      setTeamMembersCache(prev => {
        const next = { ...prev };
        delete next[key];
        return next;
      });
      closeAddMemberDialog();
    } catch {
      setAddMemberError(isDe ? 'Hinzufügen fehlgeschlagen.' : 'Adding failed.');
      setAddMemberBusy(false);
    }
  };

  // Join-Request-Cache pro (eventId|teamId). Nur für Leads relevant.
  const [joinRequestsCache, setJoinRequestsCache] = React.useState<Record<string, Array<{ Id: number; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string }>>>({});
  const joinReqKeyRef = React.useRef<Set<string>>(new Set());
  const enqueueJoinReqFetch = React.useCallback((eventId: string, teamId: string): void => {
    if (!eventId || !teamId) return;
    const key = `${eventId}|${teamId}`;
    if (joinReqKeyRef.current.has(key)) return;
    joinReqKeyRef.current.add(key);
    listTeamJoinRequestsForEvent(eventId, teamId).then(list => {
      setJoinRequestsCache(prev => ({ ...prev, [key]: list }));
    }).catch(() => {
      setJoinRequestsCache(prev => ({ ...prev, [key]: [] }));
    });
  }, [listTeamJoinRequestsForEvent]);
  const [joinReqBusyId, setJoinReqBusyId] = React.useState<number | null>(null);
  const handleDecideJoinRequest = async (eventId: string, teamId: string, requestId: number, decision: 'Approved' | 'Rejected'): Promise<void> => {
    if (joinReqBusyId !== null) return;
    setJoinReqBusyId(requestId);
    try {
      await decideTeamJoinRequest(requestId, decision);
      const key = `${eventId}|${teamId}`;
      joinReqKeyRef.current.delete(key);
      teamCacheKeyRef.current.delete(key);
      setJoinRequestsCache(prev => {
        const next = { ...prev };
        delete next[key];
        return next;
      });
      setTeamMembersCache(prev => {
        const next = { ...prev };
        delete next[key];
        return next;
      });
    } finally {
      setJoinReqBusyId(null);
    }
  };
  // v11.86: „Team verwalten"-Modal. Lead kann pro Mitglied einen
  // Trash-Button anklicken — ein zweites Confirm-Modal fordert die
  // Bestätigung an und ruft anschliessend cancelTeamMember auf.
  const [manageTeamDialog, setManageTeamDialog] = React.useState<{
    eventId: string;
    teamId: string;
    teamName: string;
    teamSize: number;
  } | null>(null);
  const [manageTeamMembers, setManageTeamMembers] = React.useState<SPRegistration[]>([]);
  const [manageTeamConfirm, setManageTeamConfirm] = React.useState<SPRegistration | null>(null);
  const [manageTeamBusyId, setManageTeamBusyId] = React.useState<number | null>(null);
  const closeManageTeamDialog = (): void => {
    setManageTeamDialog(null);
    setManageTeamMembers([]);
    setManageTeamConfirm(null);
    setManageTeamBusyId(null);
  };
  const openManageTeamDialog = (eventId: string, teamId: string, teamName: string, teamSize: number, members: SPRegistration[]): void => {
    setManageTeamDialog({ eventId, teamId, teamName, teamSize });
    setManageTeamMembers(members);
  };
  const performManageTeamCancel = async (member: SPRegistration): Promise<void> => {
    if (!manageTeamDialog || manageTeamBusyId !== null) return;
    setManageTeamBusyId(member.Id);
    try {
      const ok = await cancelTeamMember(manageTeamDialog.eventId, member);
      if (ok) {
        // Cache invalidieren, damit das Team-Badge neu lädt.
        const key = `${manageTeamDialog.eventId}|${manageTeamDialog.teamId}`;
        teamCacheKeyRef.current.delete(key);
        setTeamMembersCache(prev => {
          const next = { ...prev };
          delete next[key];
          return next;
        });
        // Aktualisierte Mitgliederliste neu laden.
        try {
          const fresh = await getTeamMembers(manageTeamDialog.eventId, manageTeamDialog.teamId);
          setTeamMembersCache(prev => ({ ...prev, [key]: fresh }));
          setManageTeamMembers(fresh);
        } catch { /* */ }
        setManageTeamConfirm(null);
      }
    } finally {
      setManageTeamBusyId(null);
    }
  };

  const { t, locale } = useLanguage();

  // v29.47: Dokumente werden beim Start nicht mehr für ALLE Events geladen
  // (das war ein Request je Event und einer der Gründe für den langen Boot).
  // Hier braucht die Seite sie wirklich — also für die sichtbaren Events
  // nachholen, sobald sie feststehen. `ensureEventDocuments` merkt sich, was
  // schon geladen wurde, und bündelt parallele Aufrufe.
  React.useEffect(() => {
    const ids = (topLevelEvents || []).map(e => e.id).filter(Boolean);
    if (ids.length === 0) return;
    void ensureEventDocuments(ids);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [topLevelEvents.length]);

  // v20.4: App-Modals statt nativer Browser-Dialoge.
  const { confirmDialog, showAlert, promptDialog } = useDialog();
  // v20.7: Persönlicher Check-in-QR-Code unter „Meine Events" — gleicher
  // Payload wie die QR-Mail (DEX|<EventNr>|<E-Mail>), client-seitig erzeugt.
  // Damit hat jeder aktive Teilnehmer seinen QR jederzeit zur Hand, auch
  // ohne die Mail zu suchen.
  const [myQrModal, setMyQrModal] = React.useState<{ dataUrl: string; name: string; tid?: number; eventTitle: string } | null>(null);
  const openMyQr = async (ev: DeloitteEvent, reg: SPRegistration): Promise<void> => {
    try {
      const QRCode = await import('qrcode');
      const email = reg.ParticipantEmail || currentUser.email || '';
      const dataUrl = await QRCode.toDataURL(`DEX|${ev.eventNumber}|${email}`, { width: 320, margin: 2 });
      const name = `${reg.Vorname || currentUser.firstName || ''} ${reg.Nachname || currentUser.surname || ''}`.trim() || email;
      setMyQrModal({ dataUrl, name, tid: reg.TeilnehmerID || undefined, eventTitle: ev.title || '' });
    } catch {
      showAlert(isDe ? 'QR-Code konnte nicht erzeugt werden.' : 'QR code could not be generated.', { variant: 'error' });
    }
  };
  // Teilnehmer-Nachrichten-Ansicht: pro Event die Broadcast-Mails (Einladung,
  // Ankündigungen) aus dem dauerhaften Kommunikations-Log lesen.
  const [commsModal, setCommsModal] = React.useState<{ eventId: string; eventTitle: string } | null>(null);
  const [commsRows, setCommsRows] = React.useState<EventCommRow[]>([]);
  const [commsLoading, setCommsLoading] = React.useState(false);
  const [commsOpenId, setCommsOpenId] = React.useState<number | null>(null);
  const openComms = (ev: DeloitteEvent): void => {
    setCommsModal({ eventId: ev.id, eventTitle: ev.title || '' });
    setCommsRows([]);
    setCommsOpenId(null);
    setCommsLoading(true);
    getEventComms(ev.id)
      .then(rows => {
        const sorted = (rows || []).slice().sort((a, b) => new Date(b.created).getTime() - new Date(a.created).getTime());
        setCommsRows(sorted);
      })
      .catch(() => { setCommsRows([]); })
      .then(() => { setCommsLoading(false); });
  };
  const isDe = locale === 'de';
  const [myEvents, setMyEvents] = React.useState<MyEventEntry[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [cancellingId, setCancellingId] = React.useState<string | null>(null);
  const [isCancelling, setIsCancelling] = React.useState(false);
  // v22.46: Erfolgs-Screen nach erfolgreicher Selbst-Abmeldung — analog zum
  // Anmelde-Success-Screen (persönliche Ansprache, Event-Bild, Organizer).
  const [cancelSuccess, setCancelSuccess] = React.useState<null | {
    title: string; imageUrl?: string; organizers: string[]; organizerEmails: string[];
    // v24.94: War die abgemeldete Anmeldung auf der Warteliste? Dann wurde KEIN
    // Platz frei und es rückt niemand nach — die Erfolgsmeldung muss das sagen.
    wasWaitlisted?: boolean;
    // v28.16: Wartelisten-Zustand des Events zum Abmelde-Zeitpunkt — die
    // Erfolgsmeldung sagt konkret, ob jemand nachrückt, statt „falls eine
    // Warteliste besteht …" zu raten.
    waitlistEnabled?: boolean;
    waitlistCount?: number;
    unlimited?: boolean;
    // v29.48: „Organizer ausblenden" galt auf diesem Erfolgs-Screen nicht —
    // hier standen die Namen weiter, obwohl das Event sie überall sonst
    // versteckt. Die drei Flags kommen deshalb mit in den Screen-State.
    hideOrganizer?: boolean;
    hideOrganizerIndividualOnly?: boolean;
    hiddenOrganizerEmails?: string[];
  }>(null);
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [editData, setEditData] = React.useState<Record<string, string>>({});
  const [isSaving, setIsSaving] = React.useState(false);
  const [loadError, setLoadError] = React.useState('');
  // v17.23: Event-Beschreibung auf der MyEvents-Karte ist standardmäßig
  // eingeklappt — pro Event-Id gemerkt, ob der User sie aufgeklappt hat.
  const [descExpanded, setDescExpanded] = React.useState<Record<string, boolean>>({});
  // v11.34: Cascade-Cancel-Dialog (Parent → Sub-Events). State hält
  // den Dialog-Inhalt + die Resolver-Funktion, damit performCancel auf
  // die User-Wahl awaiten kann statt window.confirm.
  const [cascadeDialog, setCascadeDialog] = React.useState<{
    parentTitle: string;
    // v15.8: zusätzlich Start-Datum und Ort pro Sub-Event mitliefern,
    // damit das Modal eine vollständige Info-Liste zeigen kann (vorher
    // nur Titel). Felder optional, fallbacks im Render.
    subEvents: { id: string; title: string; startDate?: string; location?: string }[];
    resolve: (_choice: 'cascade' | 'parent-only' | 'abort') => void;
    /** v14.7: Wenn das Event `requireSubEventSelection` hat, ist „nur
     *  Hauptevent abmelden, Sub-Events behalten" sinnlos (Teilnehmer
     *  konnte sich gar nicht „nur Hauptevent" anmelden). Stattdessen
     *  zeigen wir nur die zwei sinnvollen Optionen — alles abmelden ODER
     *  abbrechen und stattdessen einzelne Sections über „Anmeldung
     *  ändern" abwählen. Die Begriffe „Haupt-/Sub-Event" werden in
     *  diesem Modus durch „Event" + „Event-Sections" ersetzt. */
    isSectionedEvent?: boolean;
  } | null>(null);

  React.useEffect(() => {
    // Warten bis Events fertig geladen sind, sonst zeigen wir Fehler obwohl nur noch geladen wird
    if (isEventsLoading) {
      setIsLoading(true);
      setLoadError('');
      return;
    }
    // v11.79: Stale-while-revalidate. Wenn vor < 60 s schon mal MyEvents
    // geladen wurden, sofort den letzten Stand aus dem sessionStorage
    // rendern (Skeleton-Spinner übersprungen) und im Hintergrund frisch
    // nachladen. Beim erneuten Klick auf "Meine Events" fühlt sich die
    // Seite damit instantan an, ohne Stale-Daten zu riskieren.
    try {
      const raw = window.sessionStorage.getItem('dex:myevents:cache');
      if (raw) {
        const cache = JSON.parse(raw) as { ts: number; entries: MyEventEntry[] };
        if (cache && Array.isArray(cache.entries) && (Date.now() - cache.ts) < 60_000) {
          setMyEvents(cache.entries);
          setIsLoading(false);
          // im Hintergrund refreshen, kein "loading"-Flag
          loadMyRegistrations(true).catch(() => { /* ignore */ });
          return;
        }
      }
    } catch { /* sessionStorage kann disabled sein — dann normaler Pfad */ }
    loadMyRegistrations();
  }, [topLevelEvents, isEventsLoading]);

  async function loadMyRegistrations(silent: boolean = false): Promise<void> {
    // v11.79: Performance-Logs + Promise.all-Parallelisierung.
    // Vorher: pro angemeldetem Event eine sequentielle getMyRegistration —
    // bei N Anmeldungen N Roundtrips in Serie. Jetzt: alle parallel via
    // Promise.all, dazu Phase-Timer pro Block für ein Vorher/Nachher-
    // Profil in der Browser-Console.
    const tStart = performance.now();
    if (!silent) {
      setIsLoading(true);
      setLoadError('');
    }
    const entries: MyEventEntry[] = [];

    // v18: Im Demo-Modus immer einen Demo-Eintrag in „Meine Events" zeigen
    // (zweite der drei Demo-Sichten). Nicht an eine echte Anmeldung gekoppelt
    // — die Register-Anmeldung ist im Demo deaktiviert; dieser Eintrag zeigt
    // nur, wie eine Anmeldung in „Meine Events" aussieht.
    if (isImpersonating) {
      try {
        const demoEvent = buildDemoShowcaseEvents(isDe ? 'de' : 'en')[0];
        entries.push({
          event: demoEvent,
          registration: buildDemoMyRegistration(currentUser?.email || 'demo.user@deloitte.de', `${currentUser?.firstName || 'Demo'} ${currentUser?.surname || 'User'}`.trim()),
        });
      } catch { /* */ }
    }

    // Schneller Pfad: DEX_Participants abfragen
    const tNums = performance.now();
    const myNumbers = await getMyEventNumbers();
    const allMyNumbers = [...myNumbers.registered, ...myNumbers.waitlisted];
    // v20.0 (Audit): Set statt wiederholtem Array.indexOf in den Filter-Schleifen
    // (vorher O(Events × Anmeldungen) pro Ladevorgang).
    const myNumSet = new Set(allMyNumbers);
    // eslint-disable-next-line no-console
    dlog('perf', `[DEX][perf][myevents] getMyEventNumbers = ${Math.round(performance.now() - tNums)} ms (n=${allMyNumbers.length})`);

    if (allMyNumbers.length > 0) {
      // Nur Events laden die in DEX_Participants stehen
      // Seit v6.4: Sub-Events sind eigene DEX_Events-Items. In "My Events" zeigen
      // wir nur Top-Level-Events; Sub-Event-Anmeldungen erscheinen verschachtelt
      // unter ihrem Parent über childEventsOf().
      const relevantEvents = topLevelEvents.filter(e => e.eventNumber && myNumSet.has(e.eventNumber));
      const tRel = performance.now();
      // v11.79: parallele getMyRegistration-Calls statt sequentielle Schleife.
      const relevantRegs = await Promise.all(relevantEvents.map(async (event) => {
        try { return { event, reg: await getMyRegistration(event.id) }; }
        catch { return { event, reg: null as SPRegistration | null }; }
      }));
      // eslint-disable-next-line no-console
      dlog('perf', `[DEX][perf][myevents] getMyRegistration relevant n=${relevantEvents.length} = ${Math.round(performance.now() - tRel)} ms (parallel)`);
      for (const { event, reg } of relevantRegs) {
        if (!reg) {
          // v28.23: Das zentrale Teilnehmer-Register kennt die Anmeldung, die
          // Zeile selbst ist aber nicht lesbar (stellvertretend angelegt, Autor
          // noch die Assistenz → Item-Level-Security). Früher: `continue` — der
          // Eintrag verschwand komplett, die Person hielt sich für nicht
          // angemeldet und meldete sich ein zweites Mal an. Jetzt: Karte mit
          // Platzhalter-Registrierung anzeigen, ohne Detaildaten/Selbst-Aktionen.
          //
          // v28.99: …aber nur, wenn „nicht lesbar" überhaupt die Erklärung sein
          // KANN. Wer Admin oder Organizer dieses Events ist, sieht die
          // Teilnehmerliste vollständig — findet er dort keine Zeile, gibt es
          // keine. Dann ist der Eintrag im Register ein Überbleibsel (z.B. eine
          // Abmeldung, bei der das Nachziehen scheiterte, oder eine von Hand in
          // SharePoint gelöschte Zeile), und die Karte behauptet eine Anmeldung,
          // die es nicht gibt. In dem Fall lieber nichts anzeigen — das Register
          // räumt die Admin-Aktion „Teilnehmer-Register bereinigen" auf.
          const myMailLc = (currentUser?.email || '').trim().toLowerCase();
          const isOrganizerOfEvent = (event.organizerEmails || [])
            .some(e => (e || '').trim().toLowerCase() === myMailLc);
          if (isAdmin || isOrganizerOfEvent) continue;
          const onWaitlist = !!(event.eventNumber && myNumbers.waitlisted.indexOf(event.eventNumber) >= 0);
          entries.push({
            event,
            registration: {
              Id: 0,
              Title: '',
              ParticipantName: `${currentUser?.firstName || ''} ${currentUser?.surname || ''}`.trim(),
              ParticipantEmail: currentUser?.email || '',
              Status: onWaitlist ? 'Warteliste' : 'Angemeldet',
              RegistrationDate: '',
              CancellationDate: '',
              CustomData: '',
            },
            hiddenRow: true,
          });
          continue;
        }
        // v10.22: Sonderfall — Parent-Reg ist 'Abgemeldet', aber der User
        // hat noch aktive Sub-Event-Registrierungen (Hauptevent abgemeldet,
        // Sub-Events behalten). Dann zeigen wir das Parent als
        // sessionsOnly-Container damit die Sub-Event-Liste sichtbar
        // bleibt und der User weitere Sub-Events nachbuchen oder einzeln
        // abmelden kann. Wenn alles abgemeldet ist, normales Cancelled-
        // Verhalten.
        if (reg.Status === 'Abgemeldet') {
          const kids = childEventsOf(event.id);
          const activeKids = kids.filter(k => k.eventNumber && myNumSet.has(k.eventNumber));
          if (activeKids.length > 0) {
            entries.push({
              event,
              registration: { ...reg, Status: 'Angemeldet' },
              sessionsOnly: true,
              subEventTitles: activeKids.map(k => k.title || (isDe ? 'ohne Titel' : 'untitled')),
            });
          } else {
            entries.push({ event, registration: reg });
          }
        } else {
          entries.push({ event, registration: reg });
        }
      }
      // v6.14: Sessions-Only-Entries. Wenn der User nur Sub-Events eines Parent-Events
      // angemeldet hat (nicht das Parent selbst), zeigen wir den Parent trotzdem als
      // Container mit Badge "Nur Sessions", damit die Session-Anmeldungen sichtbar und
      // managebar bleiben. Wir iterieren über alle Top-Level-Events, für die der Parent
      // KEINE eigene aktive Registrierung hat, und prüfen, ob mindestens ein Child-Event
      // des Parents in allMyNumbers enthalten ist.
      const parentsAlreadyAdded = new Set<string>(entries.map(e => e.event.id));
      for (const parent of topLevelEvents) {
        if (parentsAlreadyAdded.has(parent.id)) continue;
        const kids = childEventsOf(parent.id);
        if (kids.length === 0) continue;
        const activeKids = kids.filter(k => k.eventNumber && myNumSet.has(k.eventNumber));
        if (activeKids.length === 0) continue;
        // Virtuelle Registration — reicht als Platzhalter. Status 'Abgemeldet' würde den
        // Eintrag in den "past"-Bucket werfen; wir nutzen 'Angemeldet' mit sessionsOnly=true
        // als Marker, damit er in "activeEntries" erscheint aber die Parent-Aktionen ausgeblendet werden.
        const virtualReg: SPRegistration = {
          Id: 0,
          Title: '',
          ParticipantName: '',
          ParticipantEmail: '',
          Status: 'Angemeldet',
          RegistrationDate: '',
          CancellationDate: '',
          CustomData: '',
        };
        // v18.53/v18.56: Im subEventsOnlyMode hält die Schatten-Parent-
        // Registrierung die übergreifenden Hauptevent-Antworten (Food
        // Preferences, Hotel, Assistenz etc.) — die brauchen wir für die
        // „Angaben ergänzen"-Vorbefüllung. ABER: der Aktiv/Abgemeldet-Bucket
        // dieser Karte richtet sich nach den AKTIVEN Sub-Event-Anmeldungen
        // (sessionsOnly), NICHT nach dem Status der Schatten-Parent-Zeile (die
        // kann z.B. aus einer alten Abmeldung 'Abgemeldet' sein). Daher NUR die
        // CustomData (+ Id für den Edit-Pfad) übernehmen und den aktiven
        // virtualReg-Status behalten. v18.56-Fix: vorher wurde die komplette
        // realParentReg übernommen → bei Status='Abgemeldet' verschwand das
        // Event fälschlich aus „Meine Events".
        let parentRegForEntry: SPRegistration = virtualReg;
        if (parent.subEventsOnlyMode) {
          try {
            const realParentReg = await getMyRegistration(parent.id);
            if (realParentReg && realParentReg.CustomData) {
              parentRegForEntry = { ...virtualReg, CustomData: realParentReg.CustomData, Id: realParentReg.Id };
            }
          } catch { /* Fallback virtualReg */ }
        }
        entries.push({
          event: parent,
          registration: parentRegForEntry,
          sessionsOnly: true,
          subEventTitles: activeKids.map(k => k.title || (isDe ? 'ohne Titel' : 'untitled')),
        });
      }
      // Zusatzschleife: abgemeldete Events finden. DEX_Participants hält nur
      // EventRegistered/EventOnWaitlist - bei Abmeldung wird die EventNumber dort
      // entfernt. Ohne diese Schleife wären alte Abmeldungen im "My Events"-Tab
      // unsichtbar, sobald der User noch für mind. ein Event angemeldet ist.
      // v10.22: Skip-Check für Events, die wir oben bereits als sessionsOnly-
      // Container eingetragen haben (sonst Doppel-Eintrag mit Status 'Angemeldet'
      // UND 'Abgemeldet' für dasselbe Parent-Event).
      const handledParentIds = new Set(entries.map(e => e.event.id));
      const remainingEvents = topLevelEvents.filter(e =>
        (!e.eventNumber || !myNumSet.has(e.eventNumber)) && !handledParentIds.has(e.id)
      );
      const tRem = performance.now();
      // v11.79: auch die Abgemeldet-Suche parallelisieren. Vorher: pro nicht-
      // angemeldetem Event ein Roundtrip — bei vielen Events der teuerste
      // Block, weil oft 80%+ "kein Eintrag" zurückkommen. Parallel schneidet
      // die Gesamtdauer auf max(slowest), nicht sum().
      const remainingRegs = await Promise.all(remainingEvents.map(async (event) => {
        try { return { event, reg: await getMyRegistration(event.id) }; }
        catch { return { event, reg: null as SPRegistration | null }; }
      }));
      // eslint-disable-next-line no-console
      dlog('perf', `[DEX][perf][myevents] getMyRegistration remaining n=${remainingEvents.length} = ${Math.round(performance.now() - tRem)} ms (parallel, Abgemeldet-Suche)`);
      for (const { event, reg } of remainingRegs) {
        if (reg && reg.Status === 'Abgemeldet') {
          // v10.22: Auch hier den hasActiveChild-Sonderfall prüfen — falls
          // DEX_Participants die Parent-EventNumber noch nicht entfernt hat
          // bzw. der User nur Sub-Event-Anmeldungen hat und kein DEX_Participants-
          // Eintrag für das Parent existiert.
          const kids = childEventsOf(event.id);
          const activeKids = kids.filter(k => k.eventNumber && myNumSet.has(k.eventNumber));
          if (activeKids.length > 0) {
            entries.push({
              event,
              registration: { ...reg, Status: 'Angemeldet' },
              sessionsOnly: true,
              subEventTitles: activeKids.map(k => k.title || (isDe ? 'ohne Titel' : 'untitled')),
            });
          } else {
            entries.push({ event, registration: reg });
          }
        }
      }
    } else {
      // Fallback: Alter Weg für Altdaten ohne DEX_Participants-Eintrag.
      // v11.79: ebenfalls parallelisiert.
      const tFb = performance.now();
      const fbRegs = await Promise.all(topLevelEvents.map(async (event) => {
        try { return { event, reg: await getMyRegistration(event.id) }; }
        catch { return { event, reg: null as SPRegistration | null }; }
      }));
      // eslint-disable-next-line no-console
      dlog('perf', `[DEX][perf][myevents] getMyRegistration fallback n=${topLevelEvents.length} = ${Math.round(performance.now() - tFb)} ms (parallel, Altdaten-Pfad)`);
      for (const { event, reg } of fbRegs) {
        if (reg) entries.push({ event, registration: reg });
      }
    }

    if (entries.length === 0 && allMyNumbers.length > 0) {
      setLoadError('Registrierungen konnten nicht geladen werden.');
    }
    setMyEvents(entries);
    if (!silent) setIsLoading(false);
    // v11.79: Cache für Stale-while-revalidate. Beim nächsten Mount
    // innerhalb 60 s wird sofort der letzte Stand gerendert.
    try {
      window.sessionStorage.setItem('dex:myevents:cache', JSON.stringify({ ts: Date.now(), entries }));
    } catch { /* ignore */ }
    const tTotal = Math.round(performance.now() - tStart);
    // eslint-disable-next-line no-console
    dlog('perf', `[DEX][perf][myevents] total = ${tTotal} ms (entries=${entries.length}, silent=${silent})`);
  }

  // Eigentliche Cancel-Logik (direkt ausführen, ohne 2-Klick-Bestätigung).
  // Wird sowohl von handleCancel (beim 2. Klick) als auch vom Auto-Cancel-
  // Deep-Link (direkt nach Navigation) genutzt.
  const performCancel = async (eventId: string): Promise<void> => {
    // v22.22: Abmeldung von bereits vergangenen Events ist nicht mehr
    // möglich — auch nicht über den Auto-Cancel-Deep-Link aus der Mail
    // (der läuft ebenfalls über performCancel).
    const guardEntry = myEvents.find(e => e.event.id === eventId);
    if (guardEntry && isEventOver(guardEntry.event)) {
      showAlert(isDe
        ? 'Dieses Event liegt bereits in der Vergangenheit — eine Abmeldung ist nicht mehr möglich.'
        : 'This event is already in the past — cancelling is no longer possible.',
        { variant: 'error' });
      setCancellingId(null);
      setIsCancelling(false);
      return;
    }
    // v29.25: Selbst-Abmeldung gesperrt (Organizer-Option) — komplett oder
    // nach der Frist. Der Guard sitzt in performCancel, damit auch der
    // Auto-Cancel-Deep-Link aus der Mail dieselbe Antwort bekommt wie der
    // Knopf.
    const lockReason = guardEntry ? selfCancelLockReason(guardEntry.event) : null;
    if (lockReason) {
      showAlert(lockReason === 'always'
        ? (isDe
          ? 'Bei diesem Event ist die Selbst-Abmeldung deaktiviert. Bitte wende dich zum Abmelden direkt an die Organizer.'
          : 'Self-cancellation is disabled for this event. Please contact the organizers directly to cancel.')
        : (isDe
          ? 'Die Abmeldefrist ist abgelaufen und die Organizer haben die Selbst-Abmeldung danach deaktiviert. Bitte wende dich zum Abmelden direkt an die Organizer.'
          : 'The cancellation deadline has passed and the organizers have disabled self-cancellation after it. Please contact the organizers directly to cancel.'),
        { variant: 'error' });
      setCancellingId(null);
      setIsCancelling(false);
      return;
    }
    setCancellingId(eventId);
    setIsCancelling(true);

    // Check if this is a late cancellation
    const entry = myEvents.find(e => e.event.id === eventId);
    const isLateCancellation = entry?.event.lastDeregisterDate && new Date(entry.event.lastDeregisterDate) < new Date();

    // v11.33/v11.34: Cascade-Prompt via gestyltem Modal — wenn das
    // Hauptevent Sub-Events hat für die der User aktiv angemeldet ist,
    // fragen ob diese auch abgemeldet werden sollen. Drei Auswahlen:
    // - 'cascade'      → Parent + alle Sub-Events abmelden
    // - 'parent-only'  → nur Parent abmelden, Sub-Events behalten
    // - 'abort'        → den ganzen Cancel-Vorgang abbrechen
    const childIdsToCancel: string[] = [];
    if (entry) {
      const kids = childEventsOf(eventId);
      // v15.8: zusätzlich startDate und location pro Sub-Event mitschicken
      // damit das Cancel-Modal Ort + Zeit anzeigen kann.
      const activeKids: { id: string; title: string; startDate?: string; location?: string }[] = [];
      for (const ce of kids) {
        try {
          // v29.25: Sub-Events mit gesperrter Selbst-Abmeldung gar nicht erst
          // in den Kaskaden-Dialog aufnehmen — der Cancel würde sonst dort
          // scheitern, wo der User ihn nicht mehr sieht.
          if (selfCancelLocked(ce, entry?.event)) continue;
          const reg = await getMyRegistration(ce.id);
          if (reg && reg.Status !== 'Abgemeldet') {
            activeKids.push({
              id: ce.id,
              title: ce.title || (isDe ? 'Sub-Event' : 'Sub-event'),
              startDate: ce.startDate || '',
              location: ce.location || '',
            });
          }
        } catch { /* ignore */ }
      }
      if (activeKids.length > 0) {
        const choice = await new Promise<'cascade' | 'parent-only' | 'abort'>(resolve => {
          setCascadeDialog({
            parentTitle: entry.event.title,
            subEvents: activeKids,
            resolve,
            isSectionedEvent: !!entry.event.requireSubEventSelection,
          });
        });
        setCascadeDialog(null);
        if (choice === 'abort') {
          setCancellingId(null);
          setIsCancelling(false);
          return;
        }
        if (choice === 'cascade') {
          for (const k of activeKids) childIdsToCancel.push(k.id);
        }
      }
    }

    const success = await cancelRegistration(eventId);
    if (success) {
      // Late cancellation: alle Organizer zusammen benachrichtigen (EINE Mail an
      // die semikolon-separierte Liste), im Deloitte-Layout via wrapTemplate,
      // in der konfigurierten Event-Sprache.
      // entry.event.organizers = NAMEN, entry.event.organizerEmails = E-Mails.
      if (isLateCancellation && entry && entry.event.organizerEmails.length > 0) {
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ctx = (window as any).__dexSpfxContext;
          if (ctx) {
            const { EventService } = await import('../services/EventService');
            const svc = new EventService(ctx);
            const userName = `${entry.registration.Vorname || ''} ${entry.registration.Nachname || ''}`.trim() || entry.registration.ParticipantEmail;
            const userEmail = entry.registration.ParticipantEmail || entry.registration.Title;
            const isDe = (entry.event.emailLanguage || 'EN').toUpperCase() === 'DE';
            const deadlineStr = new Date(entry.event.lastDeregisterDate).toLocaleDateString(isDe ? 'de-DE' : 'en-GB');
            const subject = isDe
              ? `Verspätete Abmeldung: ${entry.event.title}`
              : `Late cancellation: ${entry.event.title}`;
            const innerBody = isDe
              ? `<p><strong>${userName}</strong> hat die Anmeldung für <strong>${entry.event.title}</strong> nach Ablauf der Abmeldefrist (${deadlineStr}) storniert.</p><p><strong>E-Mail:</strong> <a href="mailto:${userEmail}">${userEmail}</a></p>`
              : `<p><strong>${userName}</strong> has cancelled their registration for <strong>${entry.event.title}</strong> after the cancellation deadline (${deadlineStr}).</p><p><strong>E-Mail:</strong> <a href="mailto:${userEmail}">${userEmail}</a></p>`;
            const heading = isDe ? 'Verspätete Abmeldung' : 'Late cancellation';
            const subheading = entry.event.title;
            const body = wrapTemplate('#ed8b00', heading, subheading, innerBody);
            // EINE Mail mit ';'-separierter Recipient-Liste - die Recipient-Spalte
            // ist Multi-Line (Note), kann also mehrere E-Mails enthalten. So sehen
            // alle Organizer die Mail gemeinsam (statt N separate Einzel-Mails).
            const toList = entry.event.organizerEmails.join(';');
            const toNames = entry.event.organizers.join(', ') || toList;
            // EmailType 'Info' (Choice-Feld lässt nur Anmeldung/Abmeldung/
            // Warteliste/Nachrücken/Info zu).
            await svc.queueEmail(subject, toList, toNames, body, 'Info', entry.event.title, eventId)
              .catch(err => console.warn('[DEX] LateCancel queueEmail failed:', err));
          }
        } catch { /* Email-Fehler ignorieren */ }
      }
      // v11.33: nach erfolgreicher Parent-Abmeldung optional die ausgewählten
      // Sub-Events kaskadieren. Best-effort — Fehler einzelner Sub-Events
      // brechen den Reload nicht ab.
      for (const childId of childIdsToCancel) {
        try { await cancelRegistration(childId); }
        catch (err) { console.warn('[DEX] cascade-cancel sub-event failed:', childId, err); }
      }
      await loadMyRegistrations();
      // v22.46: Erfolgs-Screen mit persönlicher Ansprache anzeigen.
      if (entry) {
        setCancelSuccess({
          title: entry.event.title,
          imageUrl: entry.event.imageUrl,
          organizers: entry.event.organizers || [],
          organizerEmails: entry.event.organizerEmails || [],
          wasWaitlisted: entry.registration.Status === 'Warteliste',
          // v28.16: konkreter Wartelisten-Zustand für die Erfolgsmeldung.
          waitlistEnabled: entry.event.waitlistEnabled !== false,
          waitlistCount: entry.event.waitlistCount || 0,
          unlimited: !(entry.event.maxParticipants > 0)
            && !(((entry.event.durchstarterCapacity || 0) + (entry.event.funstarterCapacity || 0)) > 0),
          hideOrganizer: !!entry.event.hideOrganizer,
          hideOrganizerIndividualOnly: !!entry.event.hideOrganizerIndividualOnly,
          hiddenOrganizerEmails: entry.event.hiddenOrganizerEmails || [],
        });
      }
    }
    setCancellingId(null);
    setIsCancelling(false);
  };

  const handleCancel = async (eventId: string): Promise<void> => {
    if (cancellingId === eventId) {
      await performCancel(eventId);
    } else {
      setCancellingId(eventId);
    }
  };

  // Auto-Cancel: wenn die Seite via Deep-Link mit Intent 'auto-cancel' geöffnet
  // wurde (z.B. aus einer Outlook-Decline-Reminder-Mail), die Registrierung
  // direkt stornieren - OHNE dass der User zusätzlich auf "Abmeldung
  // bestätigen" klicken muss. Der Klick auf den Link in der Mail gilt als
  // Bestätigung. Da der User eingeloggt sein muss und nur seine eigene
  // Registrierung cancelt, ist das sicher.
  const didAutoOpen = React.useRef(false);
  React.useEffect(() => {
    if (didAutoOpen.current) return;
    if (navIntent !== 'auto-cancel' || !selectedEventId) return;
    // Warten bis die Registrierungen geladen sind, sonst findet performCancel
    // den entry nicht (late-cancel-check schlägt fehl).
    if (isLoading) return;
    didAutoOpen.current = true;
    clearIntent();
    // Event-Karte einscrollen damit der User den aktualisierten Status sieht
    const el = document.getElementById(`dex-myevent-${selectedEventId}`);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    // Direkt cancellen
    performCancel(selectedEventId).catch(err => console.warn('[DEX] auto-cancel failed:', err));
  }, [navIntent, selectedEventId, isLoading]);

  const activeEntries = myEvents.filter(e => e.registration.Status !== 'Abgemeldet');
  const cancelledEntries = myEvents.filter(e => e.registration.Status === 'Abgemeldet');
  // v22.22: Cluster „Aktive Events" / „Vergangene Events" — gleiche Karte,
  // getrennte Sektionen (plus die bestehende „Abgemeldete Events"-Sektion).
  const upcomingEntries = activeEntries.filter(e => !isEventOver(e.event));
  const pastEntries = activeEntries.filter(e => isEventOver(e.event));

  if (isLoading) {
    // v11.79: Border-Ring-Spinner (v11.33/v11.70) durch eine saubere
    // indeterminierte Progress-Bar ersetzt — der border-radius/border-top-
    // Trick erzeugte in der SP-Hostpage trotz inline-block-Fix einen
    // mitrotierenden vertikalen Strich im Zentrum. Stattdessen jetzt
    // dasselbe Markup wie im App-Boot-Loader (siehe DexEventPlatform.tsx,
    // dexProgressSlide-Keyframe): eine endlos durchlaufende grüne Zone
    // über einer grauen Spur. Kein Rotations-Element, kein Strich-Quirk.
    return (
      <div className="page-container text-center" style={{ padding: '64px 16px' }}>
        <div style={{
          display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 16,
          width: 'min(320px, 80%)',
        }}>
          <div
            aria-hidden="true"
            style={{
              width: '100%', height: 6, borderRadius: 3,
              background: '#e5e5e5',
              overflow: 'hidden', position: 'relative',
            }}
          >
            <div style={{
              position: 'absolute', top: 0, bottom: 0,
              width: '40%',
              background: '#86bc25',
              borderRadius: 3,
              animation: 'dexProgressSlide 1.2s ease-in-out infinite',
            }} />
          </div>
          <p style={{ color: 'var(--dex-gray-600)', margin: 0, fontSize: '0.95rem' }}>
            {t('myevents.loading')}
          </p>
        </div>
        <style>{`@keyframes dexProgressSlide { 0% { left: -40%; } 100% { left: 100%; } }`}</style>
      </div>
    );
  }

  // v22.46: Abmelde-Erfolgs-Screen — gleiche Optik wie der Anmelde-Success
  // (Event-Bild, persönliche Ansprache, Organizer, Buttons).
  if (cancelSuccess) {
    const greet = (currentUser?.firstName || '').trim();
    const orgs = (cancelSuccess.organizers || [])
      .reduce<string[]>((acc, o) => [...acc, ...o.split(';')], [])
      .map(o => {
        const trimmed = o.trim();
        const parts = trimmed.split(',').map(s => s.trim());
        return parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
      })
      .filter(Boolean);
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 720, margin: '0 auto' }}>
          {cancelSuccess.imageUrl && (
            <div style={{
              width: '100%', maxWidth: 480, height: 200, margin: '0 auto 24px',
              borderRadius: 'var(--dex-radius-lg)',
              background: `url(${cancelSuccess.imageUrl}) center/cover no-repeat`,
              filter: 'grayscale(0.4)', opacity: 0.85,
            }} />
          )}
          <h2 style={{ marginTop: 0, color: 'var(--dex-red, #c00)' }}>
            {isDe ? 'Abmeldung bestätigt' : 'Cancellation confirmed'}
          </h2>
          <div className="mt-8" style={{ color: 'var(--dex-gray-700)', textAlign: 'left', maxWidth: 520, margin: '8px auto 0', lineHeight: 1.6 }}>
            <p style={{ margin: '0 0 10px' }}>
              {isDe
                ? <>Hallo{greet ? <> <strong>{greet}</strong></> : ''},</>
                : <>Hi{greet ? <> <strong>{greet}</strong></> : ''},</>}
            </p>
            <p style={{ margin: '0 0 10px' }}>
              {cancelSuccess.wasWaitlisted
                ? (isDe
                  ? <>du hast dich erfolgreich von der <strong>Warteliste</strong> des Events <strong>„{cancelSuccess.title}“</strong> abgemeldet.</>
                  : <>you have successfully removed yourself from the <strong>waitlist</strong> for <strong>“{cancelSuccess.title}”</strong>.</>)
                : (() => {
                  // v28.16: konkrete Aussage statt „falls eine Warteliste
                  // besteht …" — die App KENNT den Zustand: unbegrenzte
                  // Events haben keine Plätze, bei besetzter Warteliste
                  // rückt die nächste Person nach, sonst ist der Platz
                  // einfach wieder frei.
                  const seatTail = cancelSuccess.unlimited
                    ? null
                    : (cancelSuccess.waitlistEnabled && (cancelSuccess.waitlistCount || 0) > 0)
                      ? (isDe
                        ? <> Dein Platz ist wieder frei — die nächste Person auf der Warteliste rückt automatisch nach.</>
                        : <> Your spot is free again — the next person on the waitlist is promoted automatically.</>)
                      : (isDe
                        ? <> Dein Platz ist wieder frei.</>
                        : <> Your spot is free again.</>);
                  return isDe
                    ? <>du hast dich erfolgreich vom Event <strong>„{cancelSuccess.title}“</strong> abgemeldet.{seatTail}</>
                    : <>you have successfully cancelled your registration for <strong>“{cancelSuccess.title}”</strong>.{seatTail}</>;
                })()}
            </p>
            <p style={{ margin: 0 }}>
              {isDe
                ? <>Falls du es dir anders überlegst, kannst du dich jederzeit über den Anmelde-Bereich erneut anmelden.</>
                : <>If you change your mind, you can register again anytime via the registration area.</>}
            </p>
          </div>
          {orgs.length > 0 && !(cancelSuccess.hideOrganizer && !cancelSuccess.hideOrganizerIndividualOnly) && (
            <div style={{ marginTop: 24, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8 }}>
              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, fontWeight: 600 }}>Organizer</div>
              <OrganizerList
                names={orgs}
                emails={cancelSuccess.organizerEmails}
                hiddenEmails={(cancelSuccess.hideOrganizer && cancelSuccess.hideOrganizerIndividualOnly) ? (cancelSuccess.hiddenOrganizerEmails || []) : []}
                size="md"
              />
            </div>
          )}
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button className="btn btn-primary" onClick={() => setCancelSuccess(null)}>
              {t('myevents.title')}
            </button>
            <button className="btn btn-secondary" onClick={() => { setCancelSuccess(null); navigate('register'); }}>
              {isDe ? 'Zur Anmeldung' : 'Go to registration'}
            </button>
          </div>
        </div>
      </div>
    );
  }

  // v30.66: Props-Buendel fuer die ausgelagerte Karte (MyEventCard) — die
  // Liste unten reicht genau das durch, was der Karten-Rumpf frueher aus
  // diesem Scope gelesen hat.
  const myEventCardProps = {
    cancellingId,
    cancelRegistration,
    childEventsOf,
    confirmDialog,
    deleteFieldDocument,
    deleteMyEventAttachment,
    descExpanded,
    editData,
    editingId,
    enqueueJoinReqFetch,
    enqueueTeamFetch,
    getAllRegistrations,
    getMyRegistration,
    handleCancel,
    handleDecideJoinRequest,
    isCancelling,
    isDe,
    isSaving,
    joinReqBusyId,
    joinRequestsCache,
    listFieldDocuments,
    listMyEventAttachments,
    loadMyRegistrations,
    locale,
    openComms,
    openManageTeamDialog,
    openMyQr,
    registerForEvent,
    searchUser,
    searchUsers,
    setAddMemberConsent,
    setAddMemberDialog,
    setAddMemberError,
    setAddMemberPick,
    setAddMemberQuery,
    setAddMemberResults,
    setCancellingId,
    setDescExpanded,
    setEditData,
    setEditingId,
    setIsSaving,
    setMyEvents,
    showAlert,
    switchSplitGroup,
    t,
    teamMembersCache,
    updateMyRegistration,
    uploadFieldDocument,
    uploadMyEventAttachment,
  };

  return (
    // v9.9: max-width damit "Meine Events" auf Desktop nicht die volle Breite
    // einnimmt — die einspaltige Karten-Liste sieht sonst auf >1400px-Screens
    // unangenehm gestreckt aus. Inline-style damit das globale .page-container
    // (max-width:100%) andere Seiten nicht beeinflusst.
    <div className="page-container" style={{ maxWidth: 1100, marginLeft: 'auto', marginRight: 'auto' }}>
      <style>{`
        @keyframes dex-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        /* v9.9: Hover-Effekt für den Abmelden-Button — füllt rot, hebt leicht
           an und wirft einen weicheren Schatten, damit die Affordance klar ist.
           Im 2. Klick-Zustand (cancellingId === event.id) bleibt er ohnehin
           rot — nur der idle-State braucht Feedback. */
        .dex-cancel-btn { transition: background 120ms ease, color 120ms ease, transform 120ms ease, box-shadow 120ms ease; }
        .dex-cancel-btn:not(.dex-cancel-btn--armed):hover { background: var(--dex-red) !important; color: #fff !important; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(218,41,28,0.25); }
        .dex-cancel-btn--armed:hover { transform: translateY(-1px); box-shadow: 0 4px 14px rgba(218,41,28,0.4) !important; }
      `}</style>
      {/* v11.99: Page-Level-Refresh-Button entfernt — der Header oben
          rechts hat bereits einen Aktualisieren-Button, doppelt verwirrt. */}
      <div style={{ marginBottom: 16 }}>
        <h2 style={{ margin: 0 }}>{t('myevents.title')}</h2>
      </div>

      {loadError && (
        <div className="card" style={{ padding: 16, marginBottom: 16, color: 'var(--dex-red)' }}>
          {loadError}
        </div>
      )}

      {activeEntries.length === 0 && cancelledEntries.length === 0 && !loadError && (
        <div className="card text-center" style={{ padding: 48 }}>
          <p style={{ color: 'var(--dex-gray-400)' }}>{t('myevents.empty')}</p>
          <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>{t('myevents.browse')}</button>
        </div>
      )}

      {activeEntries.length > 0 && (() => {
        // v22.22: Der Karten-Renderer ist die frühere map-Callback-Funktion —
        // unverändert, nur extrahiert, damit „Aktive Events" und „Vergangene
        // Events" dieselbe Karte in getrennten Clustern rendern können.
        const renderMyEventCard = (entry: MyEventEntry): React.ReactElement | null => (
          <MyEventCard key={entry.event.id} entry={entry} {...myEventCardProps} />
        );
        const clusterHeadingStyle: React.CSSProperties = {
          margin: '0 0 12px', fontSize: '1.05rem',
          color: 'var(--dex-gray-700)', display: 'flex', alignItems: 'center', gap: 8,
        };
        // v24.41: INFO-Anmeldungen — jemand anderes (Assistenz) hat MICH
        // angemeldet und verwaltet die Anmeldung; ich sehe nur Info. Nur die
        // zeigen, die nicht ohnehin schon als eigene Karte erscheinen.
        const shownEventIds = new Set<string>([...upcomingEntries, ...pastEntries].map(e => e.event.id));
        const visibleInfo = (infoAsParticipant || []).filter(d => d.eventId && !shownEventIds.has(d.eventId));
        // Pro Event nur EINEN Info-Eintrag (Klammer + Sub-Events teilen sich
        // dieselbe verwaltende Person).
        const infoByEvent = new Map<string, typeof visibleInfo[number]>();
        for (const d of visibleInfo) { if (!infoByEvent.has(d.eventId)) infoByEvent.set(d.eventId, d); }
        const infoList = Array.from(infoByEvent.values());
        return (
          <>
            {/* v24.42: Offene Anforderungen AN MICH (ich verwalte die Anmeldung). */}
            {openRequestsToMe.length > 0 && (
              <div style={{ marginBottom: 24, background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 12, padding: '14px 16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                  <Icon iconName="Mail" style={{ fontSize: 18, color: 'var(--dex-orange, #ed8b00)' }} />
                  <strong style={{ fontSize: '0.95rem', color: '#b35a00' }}>
                    {isDe ? `Offene Anforderungen (${openRequestsToMe.length})` : `Open requests (${openRequestsToMe.length})`}
                  </strong>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                  {openRequestsToMe.map(l => (
                    <div key={l.id} style={{ fontSize: '0.85rem', borderTop: '1px solid rgba(237,139,0,0.25)', paddingTop: 8 }}>
                      <div>
                        <strong>{l.requestType === 'cancel' ? (isDe ? 'Abmeldung' : 'Cancellation') : (isDe ? 'Änderung' : 'Change')}</strong>
                        {' · '}{l.eventTitle || l.eventId} · {l.participantName || l.participantEmail}
                      </div>
                      <div style={{ color: 'var(--dex-gray-600)', fontSize: '0.8rem', marginTop: 2 }}>
                        {isDe ? 'von' : 'from'} {l.requestedByName || l.requestedByEmail}{l.requestNote ? ` — „${l.requestNote}"` : ''}
                      </div>
                      <div style={{ display: 'flex', gap: 8, marginTop: 6 }}>
                        <button type="button" className="btn btn-primary" style={{ fontSize: '0.78rem', padding: '4px 12px' }} onClick={() => { void resolveReq(l, 'Done'); }}>
                          {isDe ? 'Als erledigt markieren' : 'Mark as done'}
                        </button>
                        <button type="button" className="btn btn-secondary" style={{ fontSize: '0.78rem', padding: '4px 12px' }} onClick={() => { void resolveReq(l, 'Rejected'); }}>
                          {isDe ? 'Ablehnen' : 'Reject'}
                        </button>
                      </div>
                      <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.76rem', marginTop: 4 }}>
                        {isDe ? 'Bitte führe die Änderung/Abmeldung wie gewohnt aus (oben in dieser Liste bzw. in der „Assistenz"-Kachel), dann hier als erledigt markieren.' : 'Please perform the change/cancellation as usual, then mark it done here.'}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {infoList.length > 0 && (
              <div style={{ marginBottom: 24, background: 'rgba(134,188,37,0.06)', border: '1px solid var(--dex-green, #86bc25)', borderRadius: 12, padding: '14px 16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                  <Icon iconName="ContactCard" style={{ fontSize: 18, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                  <strong style={{ fontSize: '0.95rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
                    {isDe ? 'Von deiner Assistenz verwaltet' : 'Managed by your assistant'}
                  </strong>
                </div>
                <p style={{ margin: '0 0 10px', fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Diese Anmeldungen wurden für dich vorgenommen und werden von der angegebenen Person verwaltet. Du siehst sie hier zur Info und kannst eine Änderung oder Abmeldung anfordern (die verwaltende Person bekommt eine Mail mit Direktlink).'
                    : 'These registrations were made for you and are managed by the person below. Shown here for your information; you can request a change or cancellation (the managing person gets a mail with a direct link).'}
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                  {infoList.map((d, i) => (
                    <div key={`${d.eventId}-${i}`} style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', fontSize: '0.86rem' }}>
                      <strong>{d.eventTitle || d.eventId}</strong>
                      <span style={{ color: 'var(--dex-gray-500)' }}>
                        · {isDe ? 'verwaltet von' : 'managed by'} {d.assistantName || d.assistantEmail}
                      </span>
                      {d.requestStatus === 'Open' ? (
                        <span style={{ color: 'var(--dex-orange-dark, #b35a00)', fontSize: '0.78rem' }}>
                          · {isDe ? 'Anforderung gesendet' : 'Request sent'}
                        </span>
                      ) : (
                        <span style={{ display: 'inline-flex', gap: 6 }}>
                          <button type="button" className="btn btn-secondary" style={{ fontSize: '0.76rem', padding: '3px 10px' }} onClick={() => { void submitAssistantRequest(d, 'change'); }}>
                            {isDe ? 'Änderung anfordern' : 'Request change'}
                          </button>
                          <button type="button" className="btn btn-secondary" style={{ fontSize: '0.76rem', padding: '3px 10px', color: 'var(--dex-red, #c00)' }} onClick={() => { void submitAssistantRequest(d, 'cancel'); }}>
                            {isDe ? 'Abmeldung anfordern' : 'Request cancellation'}
                          </button>
                        </span>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}
            {upcomingEntries.length > 0 && (
              <>
                <h3 style={clusterHeadingStyle}>
                  {isDe ? 'Aktive Events' : 'Active events'}
                  <span style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-400)' }}>({upcomingEntries.length})</span>
                </h3>
                <div className="my-events-list">{upcomingEntries.map(renderMyEventCard)}</div>
              </>
            )}
            {pastEntries.length > 0 && (
              <>
                <h3 style={{ ...clusterHeadingStyle, marginTop: upcomingEntries.length > 0 ? 28 : 0 }}>
                  {isDe ? 'Vergangene Events' : 'Past events'}
                  <span style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-400)' }}>({pastEntries.length})</span>
                </h3>
                <p style={{ margin: '-6px 0 12px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                  {isDe
                    ? 'Diese Events liegen in der Vergangenheit — eine Abmeldung ist hier nicht mehr möglich.'
                    : 'These events are in the past — cancelling is no longer possible here.'}
                </p>
                <div className="my-events-list">{pastEntries.map(renderMyEventCard)}</div>
              </>
            )}
          </>
        );
      })()}

      {cancelledEntries.length > 0 && (
        // v11.97: Cancelled-Liste einklappbar — Default eingeklappt, damit
        // die Liste nicht den ganzen Screen mit abgemeldeten Events flutet.
        <CancelledEventsCollapsible
          count={cancelledEntries.length}
          title={t('myevents.cancelledevents')}
          locale={t('myevents.cancelledon')}
          entries={cancelledEntries}
          formatDate={formatDate}
          statusLabel={t('status.cancelled')}
          hintText={locale === 'de'
            ? 'Diese Events hast du abgemeldet. Möchtest du wieder teilnehmen, melde dich im Anmelde-Bereich einfach neu an.'
            : 'You cancelled these events. If you want to attend again, simply register anew in the registration area.'}
          reRegisterLabel={locale === 'de' ? 'Zur Anmeldung' : 'Go to registration'}
          onReRegister={() => navigate('register')}
        />
      )}

      {/* v11.34: Cascade-Cancel-Modal — ersetzt das früher genutzte
          window.confirm. Drei klare Aktionen: alles abmelden, nur
          Hauptevent, Abbrechen. */}
      {/* v11.83: Add-Member-Modal (Team-Lead fügt nachträglich ein
          Mitglied hinzu). Layout-Logik analog zum cascadeDialog. */}
      {addMemberDialog && (
        <AddMemberModal
          addMemberBusy={addMemberBusy}
          addMemberConsent={addMemberConsent}
          addMemberError={addMemberError}
          addMemberIncludeIntl={addMemberIncludeIntl}
          addMemberPick={addMemberPick}
          addMemberQuery={addMemberQuery}
          addMemberQueryTimer={addMemberQueryTimer}
          addMemberResults={addMemberResults}
          addMemberSearching={addMemberSearching}
          closeAddMemberDialog={closeAddMemberDialog}
          isDe={isDe}
          searchUsers={searchUsers}
          setAddMemberConsent={setAddMemberConsent}
          setAddMemberIncludeIntl={setAddMemberIncludeIntl}
          setAddMemberPick={setAddMemberPick}
          setAddMemberQuery={setAddMemberQuery}
          setAddMemberResults={setAddMemberResults}
          setAddMemberSearching={setAddMemberSearching}
          submitAddMember={submitAddMember}
        />
      )}
      {/* v11.86: Manage-Team-Modal — Lead bearbeitet sein Team
          (Mitglieder abmelden). Pflicht-Confirm vor dem eigentlichen
          Cancel; der Cancel selbst läuft über cancelTeamMember im
          EventContext. */}
      {manageTeamDialog && (
        <ManageTeamModal
          closeManageTeamDialog={closeManageTeamDialog}
          currentUserEmail={currentUserEmail}
          isDe={isDe}
          manageTeamBusyId={manageTeamBusyId}
          manageTeamConfirm={manageTeamConfirm}
          manageTeamDialog={manageTeamDialog}
          manageTeamMembers={manageTeamMembers}
          performManageTeamCancel={performManageTeamCancel}
          setManageTeamConfirm={setManageTeamConfirm}
        />
      )}
      {cascadeDialog && (
        <CascadeCancelModal
          cascadeDialog={cascadeDialog}
          isDe={isDe}
        />
      )}

      {/* v20.7: Modal mit dem persönlichen Check-in-QR-Code (wie in der
          QR-Mail: großer Code + Name + Teilnehmer-Nr. als Fallback für den
          manuellen Check-in). */}
      {myQrModal && (
        <MyQrModal
          isDe={isDe}
          myQrModal={myQrModal}
          setMyQrModal={setMyQrModal}
        />
      )}

      {/* Nachrichten-zum-Event-Modal: listet die Broadcast-Mails (Einladung,
          Ankündigungen) aus dem Kommunikations-Log; ein Klick öffnet den vollen
          HTML-Body (im isolierten iframe gerendert). */}
      {commsModal && (
        <EventCommsModal
          commsLoading={commsLoading}
          commsModal={commsModal}
          commsOpenId={commsOpenId}
          commsRows={commsRows}
          isDe={isDe}
          setCommsModal={setCommsModal}
          setCommsOpenId={setCommsOpenId}
        />
      )}
    </div>
  );
}
