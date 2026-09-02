/* useEventSelection — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 3417-3665 des
 * Stands vor dem Schnitt). Der Rumpf ist zeichengleich uebernommen; was die
 * Gruppe aus dem Komponenten-Scope liest, kommt als `ctx` herein, was sie
 * nach aussen liefert, geht als Objekt zurueck.
 */
import * as React from 'react';
import { ACCESS_DENIED_MSG } from '../../admin/adminConstants';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';
import { buildStaticCheckInUrl, defaultCheckInWindow, generateSelfCheckInToken } from '../../../utils/selfCheckIn';
import { localizeStatus } from '../../../utils/eventStatus';

export interface UseEventSelectionCtx {
  adminEvents: DeloitteEvent[];
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  detailCardRef: React.MutableRefObject<HTMLDivElement>;
  eventServiceRef: EventService;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  isDe: boolean;
  navigate: (page: import("../../../context/NavigationContext").Page, eventId?: string, intent?: import("../../../context/NavigationContext").NavIntent) => void;
  refreshEvents: () => Promise<void>;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  selectedEventId: string;
  setDeniedSubEventLists: React.Dispatch<React.SetStateAction<string[]>>;
  setIsLoadingRegs: React.Dispatch<React.SetStateAction<boolean>>;
  setRegLoadError: React.Dispatch<React.SetStateAction<string>>;
  setRegistrations: React.Dispatch<React.SetStateAction<SPRegistration[]>>;
  setReservedDetailHeight: React.Dispatch<React.SetStateAction<number>>;
  setSelectedEvent: React.Dispatch<React.SetStateAction<DeloitteEvent>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  updateEvent: (eventId: string, updates: Record<string, unknown>, opts?: { skipReload?: boolean; }) => Promise<boolean>;
}

export interface UseEventSelectionResult {
  handleSelectEvent: (event: DeloitteEvent) => Promise<void>;
  openSelfCheckInModal: () => Promise<void>;
  sciBusy: boolean;
  sciFrom: string;
  sciModalOpen: boolean;
  sciModalQr: string;
  sciSaveMsg: string;
  sciTo: string;
  sciToken: string;
  setSciBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setSciFrom: React.Dispatch<React.SetStateAction<string>>;
  setSciModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setSciSaveMsg: React.Dispatch<React.SetStateAction<string>>;
  setSciTo: React.Dispatch<React.SetStateAction<string>>;
  toggleDraftStatus: () => Promise<void>;
}

export function useEventSelection(ctx: UseEventSelectionCtx): UseEventSelectionResult {
  const {
    adminEvents, childEventsOf, confirmDialog, detailCardRef, eventServiceRef, getAllRegistrations,
    isDe, navigate, refreshEvents, registrations, selectedEvent, selectedEventId,
    setDeniedSubEventLists, setIsLoadingRegs, setRegLoadError, setRegistrations,
    setReservedDetailHeight, setSelectedEvent, showAlert, updateEvent,
  } = ctx;
  const handleSelectEvent = async (event: DeloitteEvent): Promise<void> => {
    // v18.24: aktuelle Card-Höhe einfrieren, BEVOR der State wechselt (DOM
    // zeigt noch den alten Stand) — verhindert das Zusammenklappen während
    // die Teilnehmer des neuen Events geladen werden.
    setReservedDetailHeight(detailCardRef.current?.offsetHeight);
    setSelectedEvent(event);
    // v10.19: NavigationContext.selectedEventId mitziehen, damit Header die
    // Page-ID granular ableiten kann (admin-center vs. admin-event) und der
    // Deep-Link-Kopier-Button immer die echte Item-ID des aktuell offenen
    // Events kennt. Skip falls bereits synchron — sonst doppelter History-
    // Eintrag beim Auto-Select via Deep-Link.
    if (selectedEventId !== event.id) {
      navigate('admin', event.id);
    }
    setIsLoadingRegs(true);
    setRegLoadError('');
    setDeniedSubEventLists([]);
    try {
      // v30.37: Auch hier zählt der HTTP-Status. Ein 403 kam bisher als leere
      // Liste an und wurde als „Noch keine Teilnehmer registriert." gerendert —
      // die freundlichste denkbare Lüge.
      // v30.67: JEDER Rückruf zählt — 429 und 5xx fehlten in der Liste und
      // kamen als „Noch keine Teilnehmer registriert." an. Und `st === 0`
      // (Netzfehler) war zwar gelistet, machte `ownDenied` aber zu 0 = falsy,
      // der Hinweis blieb also gerade im Netzfehler-Fall aus.
      let ownDenied = false;
      const regs = await getAllRegistrations(event.id, () => { ownDenied = true; });
      setRegistrations(regs);
      if (ownDenied) setRegLoadError(ACCESS_DENIED_MSG);
    } catch {
      setRegistrations([]);
      setRegLoadError('Teilnehmerliste konnte nicht geladen werden.');
    }
    setIsLoadingRegs(false);
    // Reservierung freigeben — der neue Inhalt steht jetzt, die Card nimmt
    // im selben Render die echte neue Höhe an (kein Zwischen-Kollaps).
    setReservedDetailHeight(undefined);
  };

  // v6.31: wenn navigation.selectedEventId gesetzt ist beim Mount (z.B. vom
  // Handbuch-Preview oder einem Deep-Link), direkt in die Detail-Ansicht
  // springen statt auf die Event-Auswahl-Liste.
  const didAutoSelectRef = React.useRef(false);
  React.useEffect(() => {
    if (didAutoSelectRef.current) return;
    if (!selectedEventId || selectedEvent) return;
    const match = adminEvents.find(e => e.id === selectedEventId);
    if (match) {
      didAutoSelectRef.current = true;
      handleSelectEvent(match).catch(() => { /* Fehler wird intern gesetzt */ });
    }
  }, [selectedEventId, adminEvents, selectedEvent]);

  // v22.40: Auto-Heilung stale Überbuchungs-Marker. Hat sich seit dem
  // „Überbuchung prüfen"-Lauf jemand abgemeldet, passt eine vorher als
  // überbucht markierte Person womöglich wieder in die Kapazität (oder ist
  // selbst nicht mehr aktiv). Solche `OverbookReview='Pending'`-Marker werden
  // hier still entfernt — sonst zeigt die Review-Box (und die orange Tabellen-
  // Markierung) jemanden als „über Kapazität", der längst regulär drinsteht.
  const overbookHealRef = React.useRef(false);
  React.useEffect(() => {
    if (overbookHealRef.current) { overbookHealRef.current = false; return; }
    if (!selectedEvent || !selectedEvent.subsiteUrl || !eventServiceRef) return;
    const flagged = registrations.filter(r => r.OverbookReview === 'Pending');
    if (flagged.length === 0) return;
    const isSplit = typeof selectedEvent.durchstarterCapacity === 'number'
      && typeof selectedEvent.funstarterCapacity === 'number'
      && ((selectedEvent.durchstarterCapacity || 0) > 0 || (selectedEvent.funstarterCapacity || 0) > 0);
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const groupOf = (r: SPRegistration): string => r.StarterType || r.PreferredStarterType || '';
    const keyOf = (r: SPRegistration): string => isSplit ? (groupOf(r) || '?') : 'all';
    const capOf = (key: string): number => {
      if (!isSplit) return selectedEvent.maxParticipants || 0;
      if (key === 'Durchstarter') return selectedEvent.durchstarterCapacity || 0;
      if (key === 'Funstarter') return selectedEvent.funstarterCapacity || 0;
      return 0;
    };
    const activeByGroup: Record<string, SPRegistration[]> = {};
    registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0).slice().sort((a, b) => a.Id - b.Id)
      .forEach(r => { const k = keyOf(r); (activeByGroup[k] = activeByGroup[k] || []).push(r); });
    // Stale = nicht (mehr) aktiv ODER Position passt wieder in die Kapazität.
    const stale = flagged.filter(r => {
      const k = keyOf(r); const cap = capOf(k); const bucket = activeByGroup[k] || [];
      const idx = bucket.findIndex(x => x.Id === r.Id);
      if (idx < 0) return true;
      return !(cap > 0 && (idx + 1) > cap);
    });
    if (stale.length === 0) return;
    let cancelled = false;
    (async () => {
      for (const r of stale) {
        try { await eventServiceRef.clearOverbookMark(selectedEvent.subsiteUrl, r.Id); }
        catch (err) { console.warn('[DEX] clearOverbookMark (auto-heal) failed:', err); }
      }
      if (cancelled) return;
      overbookHealRef.current = true; // nächsten Effekt-Lauf nach Reload überspringen
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    })().catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [registrations, selectedEvent?.id]);

  // Soft-Refresh-Sync: wenn `events` durch refreshEvents() aktualisiert wurde
  // (z.B. nach Event-deaktivieren, Event-aktivieren, Edit-save), den lokalen
  // `selectedEvent`-State aus der frischen Liste neu derivieren. Sonst bleibt
  // der Status-Badge (z.B. „Aktiv" vs „Entwurf") nach Toggle stale, weil
  // `selectedEvent` ein eigener useState ist und nicht aus `events` derived.
  React.useEffect(() => {
    if (!selectedEvent) return;
    const fresh = adminEvents.find(e => e.id === selectedEvent.id);
    if (fresh && fresh !== selectedEvent) {
      setSelectedEvent(fresh);
    }
  }, [adminEvents]);

  // v17.9: Map regId → Beitritts-Position (1-basiert, sortiert nach
  // v17.15: joinOrderById useMemo entfernt — wurde mit der Beitritts-#-
  // Spalte (v17.9) eingeführt, die der User in v17.10 wieder rausgeworfen
  // hat. Damit kein Hook mehr, der bei /joinOrder/ stale referenziert war.

  // v20.0 (Audit): ungenutzte Helper-Funktion getRegListUrl entfernt
  // (war seit Jahren nie aufgerufen, lieferte ohnehin nur `<base>/Lists`).

  // v20.1: Self-Check-in auto-aktivieren, falls das Event noch keinen aktiven
  // Token hat (Wizard-Toggle nie gesetzt): Token erzeugen + am Event
  // persistieren. Damit sind QR-PDF + Live-Anzeige grundsätzlich immer
  // verfügbar — der Klick auf die Aktion IST die Aktivierung.
  const ensureSelfCheckInReady = async (ev: DeloitteEvent): Promise<string | null> => {
    if (ev.selfCheckInEnabled && ev.selfCheckInToken) return ev.selfCheckInToken;
    const token = ev.selfCheckInToken || generateSelfCheckInToken();
    let ok = false;
    try {
      ok = await updateEvent(ev.id, { 'SelfCheckInEnabled': true, 'SelfCheckInToken': token });
    } catch { ok = false; }
    if (!ok) {
      showAlert(isDe
        ? 'Self-Check-in konnte nicht aktiviert werden (Speichern am Event fehlgeschlagen). Bitte erneut versuchen.'
        : 'Self check-in could not be activated (saving to the event failed). Please try again.');
      return null;
    }
    return token;
  };

  // v20.2: Self-Check-in-QR-Kachel unter dem Event-Logo + Erklär-/Einstell-Modal.
  // Die Kachel erscheint ab 5 Tagen vor Event-Start ODER sobald QR-Codes
  // versendet wurden; Klick öffnet das Modal mit großem QR, PDF-/Live-Aktionen
  // und dem editierbaren Check-in-Zeitfenster (Von/Bis).
  const [sciModalOpen, setSciModalOpen] = React.useState(false);
  const [sciModalQr, setSciModalQr] = React.useState('');
  const [sciToken, setSciToken] = React.useState('');
  const [sciFrom, setSciFrom] = React.useState('');
  const [sciTo, setSciTo] = React.useState('');
  const [sciBusy, setSciBusy] = React.useState(false);
  const [sciSaveMsg, setSciSaveMsg] = React.useState('');
  // v30.38: Der Mini-QR ist entfallen. Er wurde beim Öffnen JEDES Events
  // erzeugt (qrcode-Chunk + Canvas), nur um in der Kachel als 64-px-Vorschau zu
  // stehen — die Kachel führt jetzt in den Einstieg „QR-Codes und Check-In" und
  // zeigt ein Icon. Eine Vorschau des Self-Check-in-Codes an einer Stelle, an der
  // man sich noch gar nicht für Self-Check-in entschieden hat, nahm die Auswahl
  // vorweg; den großen QR gibt es im Modal weiterhin (`sciModalQr`).
  const isoToLocalInput = (iso?: string): string => {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    const p = (n: number): string => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
  };
  const openSelfCheckInModal = async (): Promise<void> => {
    if (!selectedEvent || sciBusy) return;
    setSciBusy(true);
    try {
      const token = await ensureSelfCheckInReady(selectedEvent);
      if (!token) return;
      setSciToken(token);
      // v20.3: Von/Bis immer vorbelegen — gespeicherte Werte ODER der
      // Standard (2 Stunden vor Event-Start bis Event-Ende). Der Standard
      // gilt auch zur Laufzeit, solange nichts anderes gespeichert ist
      // (isWithinCheckInWindow) — Anzeige und Verhalten sind damit deckungsgleich.
      const def = defaultCheckInWindow(selectedEvent.startDate, selectedEvent.endDate);
      setSciFrom(isoToLocalInput(selectedEvent.selfCheckInFrom) || (def.opensAt ? isoToLocalInput(def.opensAt.toISOString()) : ''));
      setSciTo(isoToLocalInput(selectedEvent.selfCheckInTo) || (def.closesAt ? isoToLocalInput(def.closesAt.toISOString()) : ''));
      setSciSaveMsg('');
      try {
        const QRCode = await import('qrcode');
        setSciModalQr(await QRCode.toDataURL(buildStaticCheckInUrl(token), { width: 560, margin: 1 }));
      } catch { setSciModalQr(''); }
      setSciModalOpen(true);
    } finally { setSciBusy(false); }
  };
  // v20.3: Der Status-Badge neben dem Event-Titel ist klickbar — Aktiv ⇄
  // Entwurf (ersetzt den früheren Eintrag im Aktionen-Menü). Gleiche Logik
  // wie der alte v11.89-Toggle: IsFictive flippen, beim Live-Schalten
  // Legacy-EventStatus auf 'Active' setzen.
  const toggleDraftStatus = async (): Promise<void> => {
    if (!selectedEvent) return;
    const isDraft = !!selectedEvent.isFictive;
    // v22.15: Abgeschlossen/Abgesagt → zurück auf Aktiv (Reaktivierung).
    // Vorher waren diese Zustände eine Sackgasse — auch für Admins.
    const isFinalState = !isDraft && (selectedEvent.status === 'Completed' || selectedEvent.status === 'Cancelled');
    if (isFinalState) {
      const fromLabel = isDe ? localizeStatus(selectedEvent.status) : selectedEvent.status;
      if (!(await confirmDialog(
        isDe
          ? `Event von „${fromLabel}" wieder auf Aktiv setzen? Danach ist es für die Berechtigten wieder sichtbar und buchbar. Hinweis: Liegt das End-Datum in der Vergangenheit, setzt der automatische Aufräum-Lauf das Event beim nächsten App-Start erneut auf „Abgeschlossen" — dann zuerst das Datum korrigieren.`
          : `Set event from "${fromLabel}" back to Active? It will be visible and bookable for eligible users again. Note: if the end date is in the past, the automatic cleanup will set it back to "Completed" on the next app start — fix the date first in that case.`,
        { title: isDe ? 'Event reaktivieren' : 'Reactivate event', confirmLabel: isDe ? 'Auf Aktiv setzen' : 'Set to Active' },
      ))) return;
      const ok = await updateEvent(selectedEvent.id, { 'EventStatus': 'Active' });
      if (ok) {
        setSelectedEvent(prev => prev ? { ...prev, status: 'Active' } : prev);
        await refreshEvents();
      } else {
        showAlert(isDe
          ? 'Der Status konnte nicht geändert werden. Vermutlich fehlen dir Schreibrechte auf der Event-Liste — bitte einen Haupt-Organizer oder Admin den Status umschalten lassen.'
          : 'The status could not be changed. You probably lack write permission on the event list — please ask a main organizer or admin to switch the status.', { variant: 'error' });
      }
      return;
    }
    const nextIsFictive = !isDraft;
    const confirmMsg = nextIsFictive
      ? (isDe ? 'Event auf "Entwurf" zurücksetzen? Reguläre User sehen das Event danach nicht mehr.' : 'Reset event to "draft"? Regular users will no longer see the event afterwards.')
      : (isDe ? 'Event live schalten? Alle Berechtigten können sich danach anmelden.' : 'Publish event? All eligible users can register afterwards.');
    if (!(await confirmDialog(confirmMsg, { title: isDe ? 'Event-Status ändern' : 'Change event status', confirmLabel: nextIsFictive ? (isDe ? 'Auf Entwurf setzen' : 'Set to draft') : (isDe ? 'Live schalten' : 'Publish') }))) return;
    const patch: Record<string, unknown> = { 'IsFictive': nextIsFictive };
    if (!nextIsFictive) patch['EventStatus'] = 'Active';
    const ok = await updateEvent(selectedEvent.id, patch);
    if (ok) {
      // Badge sofort umschalten — selectedEvent ist lokaler State und wird
      // durch refreshEvents nicht automatisch ersetzt.
      setSelectedEvent(prev => prev ? { ...prev, isFictive: nextIsFictive, ...(nextIsFictive ? {} : { status: 'Active' }) } : prev);
      // v22.67: Beim Live-Schalten eines Events mit Sub-Events werden die
      // Sub-Events automatisch mit live geschaltet (Entwurf → Aktiv) — sonst
      // bliebe das Event sichtbar, aber die Sub-Events wären für Teilnehmer
      // nicht buchbar.
      if (!nextIsFictive) {
        for (const c of childEventsOf(selectedEvent.id)) {
          if (c.isFictive) {
            try { await updateEvent(c.id, { 'IsFictive': false, 'EventStatus': 'Active' }); } catch { /* best-effort */ }
          }
        }
      }
      await refreshEvents();
    } else {
      // v22.14: vorher scheiterte der Klick STUMM — der Organizer dachte,
      // der Status lasse sich nicht ändern, ohne zu erfahren warum.
      showAlert(isDe
        ? 'Der Status konnte nicht geändert werden. Vermutlich fehlen dir Schreibrechte auf der Event-Liste (z.B. als Co-Organizer ohne Organizer-Rolle) — bitte einen Haupt-Organizer oder Admin den Status umschalten lassen.'
        : 'The status could not be changed. You probably lack write permission on the event list (e.g. co-organizer without the organizer role) — please ask a main organizer or admin to switch the status.', { variant: 'error' });
    }
  };
  return {
    handleSelectEvent, openSelfCheckInModal, sciBusy, sciFrom, sciModalOpen, sciModalQr,
    sciSaveMsg, sciTo, sciToken, setSciBusy, setSciFrom, setSciModalOpen, setSciSaveMsg, setSciTo,
    toggleDraftStatus,
  };
}

