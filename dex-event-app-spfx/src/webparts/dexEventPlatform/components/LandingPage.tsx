// Landing Page - Startbildschirm mit animiertem Orb und Willkommensnachricht

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
// v18.25: personalisierte Begrüßung (Vorname) im Landing-Hero — useCurrentUser
// wird dafür wieder hier benötigt. (v18.26: KPI-Zeile wieder entfernt — die
// Einsatz-Zahlen stehen nur auf dem Boot-Loader davor.)
import { useCurrentUser } from '../context/UserContext';
import { APP_VERSION } from '../version';
import InquiryModal from './InquiryModal';
// v22: Archivierungs-Info für Admins (rechts auf der Landing Page) —
// zählt beim App-Start die archivreifen Zeilen abgelaufener Events und
// bietet das Verschieben ins admin-only DEX_Archive mit Fortschrittsmodal.
import { useEvents } from '../context/EventContext';
import { useRoles } from '../context/RoleContext';
import { useDialog } from '../context/DialogContext';
import Modal from './Modal';
import { useIsMobile } from '../utils/useIsMobile';

export default function LandingPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { locale, setLocale, t } = useLanguage();
  const isDe = locale === 'de';
  const isMobile = useIsMobile();
  // v18.25: Vorname für die persönliche Begrüßung.
  const { currentUser } = useCurrentUser();
  const firstName = (currentUser?.firstName || '').trim();
  // v24.20: Tageszeitabhängige Begrüßung (Morgen/Tag/Abend) statt fixem „Hallo".
  const greetHour = new Date().getHours();
  const greeting = isDe
    ? (greetHour < 5 ? 'Hallo' : greetHour < 11 ? 'Guten Morgen' : greetHour < 18 ? 'Guten Tag' : 'Guten Abend')
    : (greetHour < 5 ? 'Hi' : greetHour < 11 ? 'Good morning' : greetHour < 18 ? 'Good afternoon' : 'Good evening');
  // v13.3: Inquiry-Modal lebt jetzt komplett in der wiederverwendbaren
  // InquiryModal-Komponente — eigene States hier entfallen.
  const [showInquiry, setShowInquiry] = React.useState(false);
  // v26: Hover-Zustand für die „DEX für dein Event nutzen"-CTA (dunkelgrüne Kontur).
  const [ctaHover, setCtaHover] = React.useState(false);

  // ==================== v22: Archivierung (Admin) ====================
  const { isAdmin, canCreateEvents } = useRoles();
  // v24.23: Die „DEX für dein Event nutzen"-Box (Werde Organizer) ist für
  // Organizer überflüssig — sie haben die Funktionen schon. Admins sehen sie
  // bewusst weiter (um die normale User-Ansicht der Landing Page zu prüfen).
  const showOrganizerCta = !canCreateEvents || isAdmin;
  const { isEventsLoading, getArchivableCount, runArchiveExpired, scanInactiveAccounts, notifyOrganizerOfInactive, autoDeregisterInactive, getSentInactiveNotices, getDeletableArchiveCount, runDeleteOldArchive, getParticipantDeletionWarnings, getParticipantDeletionDue, runParticipantDeletion, maybeSendParticipantDeletionWarnings, deleteEvent, countExternalRegistrations, refreshEvents } = useEvents();
  // v26.40: Modal-Hinweis nach automatischer Abmeldung von Ex-Deloitte-Personen.
  const [autoDeregModal, setAutoDeregModal] = React.useState<Array<{ title: string; people: Array<{ email: string; name: string }> }> | null>(null);
  // v24.51: „Organizer benachrichtigen" pro Event (inaktive Konten).
  const [notifyBusyId, setNotifyBusyId] = React.useState<string | null>(null);
  const [notifyResult, setNotifyResult] = React.useState<Record<string, string>>({});
  const { confirmDialog, showAlert } = useDialog();
  const [archInfo, setArchInfo] = React.useState<{ total: number; perList: Record<string, number> } | null>(null);
  // v23.40: Löschkonzept — Anzahl DEX_Archive-Einträge älter als 1 Monat (v23.48).
  const [delArchCount, setDelArchCount] = React.useState(0);
  const [delArchBusy, setDelArchBusy] = React.useState(false);
  // v26.32: Löschkonzept — Teilnehmerlisten (3 Monate): Vorwarn- + Fällig-Zähler.
  const [pdWarn, setPdWarn] = React.useState(0);
  const [pdDue, setPdDue] = React.useState(0);
  const [pdBusy, setPdBusy] = React.useState(false);
  // v22.45: Warnung über Teilnehmer ohne aktives Deloitte-Konto — pro Event,
  // für Organizer (eigene Events) und Admins (alle aktiven Events).
  const [inactiveSummary, setInactiveSummary] = React.useState<Array<{ eventId: string; title: string; people: Array<{ email: string; name: string }> }>>([]);
  const [archModal, setArchModal] = React.useState<null | {
    running: boolean;
    listIdx: number; listTotal: number; listName: string;
    done: number; total: number;
    summary: string[] | null;
  }>(null);
  // v22.2: Abbruch-Mechanik — Ref wird vom Abbrechen-Button gesetzt und vom
  // Lauf pro Zeile geprüft (bereits verschobene Zeilen bleiben archiviert).
  const archCancelRef = React.useRef(false);
  const [archCancelRequested, setArchCancelRequested] = React.useState(false);
  // Beim App-Start (sobald Events geladen) als Admin die archivreifen
  // Zeilen zählen — nur dann erscheint die Box mit dem Button.
  React.useEffect(() => {
    if (!isAdmin || isEventsLoading) return undefined;
    let cancelled = false;
    getArchivableCount()
      .then(r => { if (!cancelled) setArchInfo(r); })
      .catch(() => { /* Zählung best-effort — Box bleibt dann aus */ });
    // v23.40: parallel die löschbaren (alten) Archiv-Einträge zählen.
    getDeletableArchiveCount()
      .then(n => { if (!cancelled) setDelArchCount(n); })
      .catch(() => { /* best-effort */ });
    // v26.32: Teilnehmerlisten-Löschkonzept — Vorwarn-/Fällig-Zähler laden und die
    // Vorwarn-Mails an die Organizer automatisch (Queue-entdoppelt) auslösen.
    getParticipantDeletionWarnings()
      .then(list => { if (!cancelled) setPdWarn(list.length); })
      .catch(() => { /* best-effort */ });
    getParticipantDeletionDue()
      .then(list => { if (!cancelled) setPdDue(list.length); })
      .catch(() => { /* best-effort */ });
    maybeSendParticipantDeletionWarnings().catch(() => { /* best-effort */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isAdmin, isEventsLoading]);

  // v23.40: Alte Archiv-Einträge (älter als 1 Monat, v23.48) löschen.
  const startDeleteOldArchive = async (): Promise<void> => {
    if (delArchBusy || delArchCount === 0) return;
    const ok = await confirmDialog(
      isDe
        ? `${delArchCount} Archiv-Einträge, die älter als 1 Monat sind, endgültig löschen?\n\nDas Archiv (DEX_Archive) ist die letzte Ablage — diese alten Einträge werden unwiderruflich entfernt.`
        : `Permanently delete ${delArchCount} archive entries older than 1 month?\n\nThe archive (DEX_Archive) is the final storage — these old entries are removed irreversibly.`,
      { danger: true, confirmLabel: isDe ? 'Endgültig löschen' : 'Delete permanently' },
    );
    if (!ok) return;
    setDelArchBusy(true);
    try {
      const r = await runDeleteOldArchive();
      setDelArchCount(0);
      showAlert(
        isDe
          ? `${r.deleted} alte Archiv-Einträge gelöscht${r.failed ? `, ${r.failed} fehlgeschlagen` : ''}.`
          : `${r.deleted} old archive entries deleted${r.failed ? `, ${r.failed} failed` : ''}.`,
        { variant: r.failed ? 'error' : 'success' },
      );
      try { const n = await getDeletableArchiveCount(); setDelArchCount(n); } catch { /* */ }
    } catch {
      showAlert(isDe ? 'Löschen fehlgeschlagen — bitte erneut versuchen.' : 'Deletion failed — please try again.', { variant: 'error' });
    } finally { setDelArchBusy(false); }
  };
  // v26.32: Fällige Teilnehmerlisten löschen — KPIs ins DEX_EventStats archivieren,
  // dann die Teilnehmer-Subsite recyceln. Das Event bleibt in DEX_Events erhalten.
  const startParticipantDeletion = async (): Promise<void> => {
    if (pdBusy || pdDue === 0) return;
    const ok = await confirmDialog(
      isDe
        ? `Bei ${pdDue} ${pdDue === 1 ? 'Event' : 'Events'} die Teilnehmerliste endgültig löschen?\n\nDie wichtigsten Kennzahlen werden zuvor ins Statistik-Archiv übernommen. Das Event bleibt bestehen, die Teilnehmerliste wird in den SharePoint-Papierkorb verschoben.`
        : `Delete the attendee list for ${pdDue} event(s)?\n\nThe key KPIs are archived to the statistics archive first. The event is kept; the attendee list is moved to the SharePoint recycle bin.`,
      { danger: true, confirmLabel: isDe ? 'Endgültig löschen' : 'Delete permanently' },
    );
    if (!ok) return;
    setPdBusy(true);
    try {
      const r = await runParticipantDeletion();
      showAlert(
        isDe
          ? `${r.deleted} Teilnehmerliste(n) archiviert & gelöscht${r.failed ? `, ${r.failed} fehlgeschlagen` : ''}.`
          : `${r.deleted} attendee list(s) archived & deleted${r.failed ? `, ${r.failed} failed` : ''}.`,
        { variant: r.failed ? 'error' : 'success' },
      );
      try { const list = await getParticipantDeletionDue(); setPdDue(list.length); } catch { /* */ }
      try { const w = await getParticipantDeletionWarnings(); setPdWarn(w.length); } catch { /* */ }
      try { await refreshEvents(); } catch { /* */ }
    } catch {
      showAlert(isDe ? 'Löschen fehlgeschlagen — bitte erneut versuchen.' : 'Deletion failed — please try again.', { variant: 'error' });
    } finally { setPdBusy(false); }
  };
  // ==================== v24.1: Entwurf-Aufräumen (Organizer) ====================
  // Entwurf-Events (nie aktiv), deren Datum > 1 Tag her ist, kann der Organizer
  // hier direkt löschen — mit einfacher Ja-Bestätigung (kein Titel-Eintippen).
  // Sicherheits-Check beim Löschen: nie ein Event mit Anmeldungen über das
  // Organizer-Team hinaus (das wäre ein „ehemals aktives" Event → geschützt).
  const { events: _eventsForDrafts } = useEvents();
  const [draftDeleteBusyId, setDraftDeleteBusyId] = React.useState<string | null>(null);
  const staleDrafts = React.useMemo(() => {
    if (isEventsLoading) return [];
    const emailLc = (currentUser?.email || '').toLowerCase();
    const now = Date.now();
    const dayMs = 24 * 60 * 60 * 1000;
    return (_eventsForDrafts || []).filter(e => {
      if (!e.isFictive) return false;       // nur Entwürfe
      if (e.parentEventId) return false;     // nur Top-Level (Sub-Events kaskadieren mit)
      const isOrg = isAdmin
        || (e.organizerEmails || []).some(x => (x || '').toLowerCase() === emailLc)
        || (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === emailLc);
      if (!isOrg) return false;
      const endRaw = e.endDate || e.startDate;
      const endTs = endRaw ? new Date(endRaw).getTime() : 0;
      return endTs > 0 && endTs < now - dayMs; // ab 1 Tag nach Ablauf
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [_eventsForDrafts, isEventsLoading, isAdmin, currentUser?.email]);
  // v26.32: Die frühere „Events älter als 1 Jahr löschen"-Regel entfällt — das
  // Löschkonzept behält das Event und löscht nur die Teilnehmerliste nach 3
  // Monaten (Zähler pdWarn/pdDue via getParticipantDeletionWarnings/Due).
  const deleteStaleDraft = async (ev: typeof staleDrafts[number]): Promise<void> => {
    if (draftDeleteBusyId) return;
    const ok = await confirmDialog(
      isDe
        ? `Entwurf „${ev.title}" wirklich löschen? Das kann nicht rückgängig gemacht werden.`
        : `Delete draft „${ev.title}"? This cannot be undone.`,
      { danger: true, confirmLabel: isDe ? 'Ja, löschen' : 'Yes, delete' },
    );
    if (!ok) return;
    setDraftDeleteBusyId(ev.id);
    try {
      // Schutz: Events mit echten Anmeldungen dürfen hier NICHT gelöscht werden.
      const ext = await countExternalRegistrations(ev);
      if (ext > 0) {
        showAlert(isDe
          ? 'Dieses Event hat Anmeldungen über das Organizer-Team hinaus und kann hier nicht gelöscht werden (nur durch einen Admin, frühestens 1 Jahr nach dem Event).'
          : 'This event has registrations beyond the organizer team and cannot be deleted here.', { variant: 'error' });
        return;
      }
      await deleteEvent(ev.id);
      showAlert(isDe ? 'Entwurf gelöscht.' : 'Draft deleted.', { variant: 'success' });
      try { await refreshEvents(); } catch { /* */ }
    } catch {
      showAlert(isDe ? 'Löschen fehlgeschlagen — bitte erneut versuchen.' : 'Deletion failed — please try again.', { variant: 'error' });
    } finally { setDraftDeleteBusyId(null); }
  };

  // ==================== v22.1: Check-in-Hinweisbox (Landing) ====================
  // Ab 2 Tagen vor Event-Start UND sobald der QR-Massen-Versand lief (eigener
  // Status 'QR versendet'), zeigt die Landing Page über dem Start-Button eine
  // Hinweisbox „Check-in für <Event>" mit kleinem QR — Klick öffnet ihn groß.
  const { getMyRegistration, events } = useEvents();

  // v24.59: Events/Personen ausblenden, deren Organizer bereits über das
  // inaktive Konto benachrichtigt wurde (Dedup-Marker DEX_InactiveNotices).
  // Pro Event werden nur noch-nicht-benachrichtigte Personen behalten; bleibt
  // keine übrig, fällt das ganze Event aus der Box.
  type InactiveItem = { eventId: string; title: string; people: Array<{ email: string; name: string }> };
  const filterNotified = React.useCallback(async (items: InactiveItem[]): Promise<InactiveItem[]> => {
    const out: InactiveItem[] = [];
    for (const it of items) {
      let sent = new Set<string>();
      try { sent = await getSentInactiveNotices(it.eventId); } catch { sent = new Set<string>(); }
      const remaining = it.people.filter(p => !sent.has((p.email || '').toLowerCase().trim()));
      if (remaining.length > 0) out.push({ ...it, people: remaining });
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // v22.45: Inaktive-Konten-Scan für Organizer/Admins. Gedrosselt 1×/24h via
  // localStorage; Ergebnis sofort aus dem Cache angezeigt, im Hintergrund
  // aktualisiert. Organizer scannen ihre eigenen aktiven Events, Admins alle.
  React.useEffect(() => {
    if (isEventsLoading) return undefined;
    const emailLc = (currentUser?.email || '').toLowerCase();
    if (!emailLc) return undefined;
    const isOrgOf = (e: typeof events[number]): boolean =>
      (e.organizerEmails || []).some(x => (x || '').toLowerCase() === emailLc)
      || (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === emailLc);
    const relevant = (events || []).filter(e =>
      e.status === 'Active' && (e.subsiteUrl || '').trim() && (isAdmin || isOrgOf(e)));
    if (relevant.length === 0) { setInactiveSummary([]); return undefined; }
    // v26.40: Event-Setting — 'autoderegister' = automatisch abmelden (+ Modal),
    // 'notify' (Default) = im Organizer-Hinweis anzeigen.
    const isAutoDereg = (eventId: string): boolean =>
      (events || []).find(e => e.id === eventId)?.inactiveHandling === 'autoderegister';
    // v26.42: Cache-Key auf _v2 — alte Scans enthielten Fehlalarme für UMBENANNTE
    // Konten (z.B. Heirat); die neue Alias-Prüfung braucht einen frischen Scan.
    const CACHE = 'dex_inactivesummary_v2';
    // v22.46: Signatur der relevanten Events — ändert sich die Event-Liste
    // (neues Event, Status-Wechsel), wird neu gescannt statt 24h zu warten.
    const sig = relevant.map(e => e.id).sort().join('|');
    let cancelled = false;
    // Sofort aus dem Cache zeigen (falls vorhanden), auf aktuelle Events gefiltert.
    let stale = true;
    try {
      const raw = window.localStorage.getItem(CACHE);
      if (raw) {
        const parsed = JSON.parse(raw) as { ts?: number; sig?: string; items?: InactiveItem[] };
        if (parsed && typeof parsed.ts === 'number' && Array.isArray(parsed.items)) {
          const liveIds = new Set(relevant.map(e => e.id));
          // v26.40: Nur 'notify'-Events in den Organizer-Hinweis; 'autoderegister'
          // wird beim frischen Scan automatisch abgemeldet (nicht hier anzeigen).
          const cached = parsed.items.filter(it => liveIds.has(it.eventId) && !isAutoDereg(it.eventId));
          // v24.59: bereits benachrichtigte Konten gleich ausblenden.
          filterNotified(cached).then(f => { if (!cancelled) setInactiveSummary(f); }).catch(() => { if (!cancelled) setInactiveSummary(cached); });
          if (Date.now() - parsed.ts < 24 * 60 * 60 * 1000 && parsed.sig === sig) stale = false;
        }
      }
    } catch { /* */ }
    if (!stale) return () => { cancelled = true; };
    const t = window.setTimeout(() => {
      scanInactiveAccounts(relevant.map(e => ({ id: e.id, title: e.title, subsiteUrl: e.subsiteUrl })))
        .then(async items => {
          if (cancelled) return;
          // Rohscan cachen (für die nächste Anzeige), aber benachrichtigte ausblenden.
          try { window.localStorage.setItem(CACHE, JSON.stringify({ ts: Date.now(), sig, items })); } catch { /* */ }
          // v26.40: Events mit 'autoderegister' → erkannte Ex-Deloitte-Personen
          // automatisch abmelden und dem Organizer als Modal melden.
          const autoItems = items.filter(it => isAutoDereg(it.eventId));
          if (autoItems.length > 0) {
            const report: Array<{ title: string; people: Array<{ email: string; name: string }> }> = [];
            for (const it of autoItems) {
              try { const removed = await autoDeregisterInactive(it.eventId, it.people); if (removed.length > 0) report.push({ title: it.title, people: removed }); } catch { /* */ }
            }
            if (!cancelled && report.length > 0) setAutoDeregModal(report);
          }
          // Nur 'notify'-Events landen im Organizer-Hinweis.
          const notifyItems = items.filter(it => !isAutoDereg(it.eventId));
          const filtered = await filterNotified(notifyItems);
          if (!cancelled) setInactiveSummary(filtered);
        })
        .catch(() => { /* best-effort */ });
    }, 3500); // dem Boot Vorrang geben
    return () => { cancelled = true; window.clearTimeout(t); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEventsLoading, isAdmin, currentUser?.email]);

  const [checkInBoxes, setCheckInBoxes] = React.useState<Array<{
    eventId: string; title: string; qrSmall: string; qrData: string;
    name: string; tid?: number;
  }>>([]);
  const [qrBigModal, setQrBigModal] = React.useState<{ dataUrl: string; title: string; name: string; tid?: number } | null>(null);
  React.useEffect(() => {
    if (isEventsLoading) return undefined;
    const myEmail = (currentUser?.email || '').trim();
    if (!myEmail) return undefined;
    let cancelled = false;
    (async () => {
      const now = Date.now();
      // Kandidaten: Events, deren Start höchstens 2 Tage entfernt ist und die
      // noch nicht vorbei sind (Ende, Fallback: Ende des Start-Tages).
      // v22.3: kalendertägig statt exakt 48h — die Box erscheint ab 00:00 des
      // Tages, der zwei Tage vor dem Start-Tag liegt (Beispiel: Event am
      // 12.06. 18:00 → Box ab 10.06. 00:00 sichtbar, nicht erst ab 18:00).
      const candidates = events.filter(e => {
        if (!e.eventNumber || !e.startDate) return false;
        const startDt = new Date(e.startDate);
        const start = startDt.getTime();
        if (!Number.isFinite(start)) return false;
        let end = e.endDate ? new Date(e.endDate).getTime() : NaN;
        if (!Number.isFinite(end)) {
          end = new Date(startDt.getFullYear(), startDt.getMonth(), startDt.getDate(), 23, 59, 59).getTime();
        }
        const windowOpens = new Date(startDt.getFullYear(), startDt.getMonth(), startDt.getDate() - 2, 0, 0, 0).getTime();
        return now >= windowOpens && now <= end;
      }).slice(0, 4);
      if (candidates.length === 0) { if (!cancelled) setCheckInBoxes([]); return; }
      const boxes: Array<{ eventId: string; title: string; qrSmall: string; qrData: string; name: string; tid?: number }> = [];
      const QRCode = await import('qrcode');
      for (const ev of candidates) {
        try {
          const reg = await getMyRegistration(ev.id);
          // Nur wenn der QR-Versand für diese Person bereits lief — Status
          // 'QR versendet' (Massen-Versand ODER Auto-Versand nach Phase-Start).
          if (!reg || reg.Status !== 'QR versendet') continue;
          const qrData = `DEX|${ev.eventNumber}|${reg.ParticipantEmail || myEmail}`;
          const qrSmall = await QRCode.toDataURL(qrData, { width: 132, margin: 1 });
          const name = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || (reg.ParticipantEmail || myEmail);
          boxes.push({ eventId: ev.id, title: ev.title || '', qrSmall, qrData, name, tid: reg.TeilnehmerID || undefined });
        } catch { /* einzelner Lookup-Fehler — Box entfällt */ }
        if (cancelled) return;
      }
      if (!cancelled) setCheckInBoxes(boxes);
    })().catch(() => { /* best-effort */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEventsLoading, currentUser?.email]);

  // v24.13/v24.18: „Du bist angemeldet"-Box für aktive Events, für die man
  // angemeldet ist — Klick springt zu „Meine Events". Quelle ist die echte
  // Anmelde-Zeile pro Event (getMyRegistration), inkl. Sub-Event-/Klammer-
  // Anmeldungen (auf das Klammer-Event abgebildet). Events mit bereits
  // sichtbarer Check-in-/QR-Box werden beim Rendern ausgeblendet (sonst doppelt).
  const [myRegBoxes, setMyRegBoxes] = React.useState<Array<{ eventId: string; title: string; imageUrl?: string; startDate: string; location?: string }>>([]);
  const [nowTick, setNowTick] = React.useState(Date.now());
  React.useEffect(() => {
    const id = window.setInterval(() => setNowTick(Date.now()), 60000);
    return () => window.clearInterval(id);
  }, []);
  React.useEffect(() => {
    if (isEventsLoading) return undefined;
    const myEmail = (currentUser?.email || '').trim();
    if (!myEmail) return undefined;
    let cancelled = false;
    (async () => {
      const now = Date.now();
      // v24.18 BUG-FIX: Statt aus dem (best-effort, oft unvollständigen)
      // zentralen Teilnehmer-Register (getMyEventNumbers) zu lesen — was bei
      // Sub-Event-/Klammer-Anmeldungen leer blieb und die Box gar nicht
      // erscheinen ließ — prüfen wir die Anmeldung direkt pro Event über
      // getMyRegistration (derselbe verlässliche Pfad wie die Check-in-Box).
      // Sub-Event-Anmeldungen werden auf das KLAMMER-/Hauptevent abgebildet
      // (dessen Name + Bild wird angezeigt).
      const topCandidates = (events || [])
        .filter(e => !e.parentEventId && e.status === 'Active' && !e.isFictive && !!e.startDate)
        .filter(e => {
          const startTs = new Date(e.startDate).getTime();
          if (!Number.isFinite(startTs)) return false;
          let endTs = e.endDate ? new Date(e.endDate).getTime() : NaN;
          if (!Number.isFinite(endTs)) { const d = new Date(e.startDate); endTs = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59).getTime(); }
          return endTs >= now; // noch nicht vorbei
        })
        .sort((a, b) => new Date(a.startDate).getTime() - new Date(b.startDate).getTime())
        .slice(0, 20);
      const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
      const isActiveReg = (r: { Status?: string } | null | undefined): boolean =>
        !!r && ACTIVE.indexOf(r.Status || '') >= 0;
      const boxes: Array<{ eventId: string; title: string; imageUrl?: string; startDate: string; location?: string }> = [];
      for (const top of topCandidates) {
        if (cancelled) return;
        let registered = false;
        try {
          // Eigene Anmeldung am Haupt-/Klammer-Event (bei „Nur Sub-Events"
          // existiert hier eine Schatten-Zeile, sobald ein Sub-Event gebucht
          // wurde) — fängt den gemeldeten Fall direkt ab.
          registered = isActiveReg(await getMyRegistration(top.id));
        } catch { /* Lookup-Fehler — weiter mit Sub-Events */ }
        if (!registered) {
          // Fallback: irgendein Sub-Event aktiv angemeldet?
          const children = (events || []).filter(c => c.parentEventId === top.id && !c.isFictive);
          for (const child of children) {
            if (cancelled) return;
            try {
              if (isActiveReg(await getMyRegistration(child.id))) { registered = true; break; }
            } catch { /* einzelner Sub-Event-Lookup-Fehler */ }
          }
        }
        if (registered) {
          boxes.push({ eventId: top.id, title: top.title || '', imageUrl: top.imageUrl, startDate: top.startDate, location: top.location });
        }
        if (boxes.length >= 6) break;
      }
      if (!cancelled) setMyRegBoxes(boxes);
    })().catch(() => { /* best-effort */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEventsLoading, currentUser?.email, events]);

  const openBigQr = async (box: { qrData: string; title: string; name: string; tid?: number }): Promise<void> => {
    try {
      const QRCode = await import('qrcode');
      const dataUrl = await QRCode.toDataURL(box.qrData, { width: 320, margin: 2 });
      setQrBigModal({ dataUrl, title: box.title, name: box.name, tid: box.tid });
    } catch { /* */ }
  };

  const startArchive = async (): Promise<void> => {
    if (!archInfo || archInfo.total === 0) return;
    const ok = await confirmDialog(
      isDe
        ? `${archInfo.total} Zeilen abgelaufener Events jetzt ins Archiv verschieben?\n\nDie Zeilen werden aus den Arbeitslisten (E-Mails, Outlook, ID-Vergabe, Änderungsprotokoll, Zugriffs-Queue) entfernt und sind danach nur noch für Admins in der Liste DEX_Archive einsehbar.`
        : `Move ${archInfo.total} rows of expired events to the archive now?\n\nThe rows are removed from the working lists (emails, Outlook, ID assignment, change log, access queue) and are afterwards only visible to admins in the DEX_Archive list.`,
      { title: isDe ? 'Archivierung' : 'Archiving', confirmLabel: isDe ? 'Jetzt archivieren' : 'Archive now' }
    );
    if (!ok) return;
    archCancelRef.current = false;
    setArchCancelRequested(false);
    setArchModal({ running: true, listIdx: 0, listTotal: 5, listName: '', done: 0, total: 0, summary: null });
    try {
      const res = await runArchiveExpired((listIdx, listTotal, listName, done, total) => {
        setArchModal(prev => prev ? { ...prev, listIdx: listIdx + 1, listTotal, listName, done, total } : prev);
      }, () => archCancelRef.current);
      const parts: string[] = [];
      if (res.cancelled) {
        parts.push(isDe
          ? `Abgebrochen — ${res.archived} Zeilen wurden bereits verschoben, der Rest bleibt in den Arbeitslisten und steht beim nächsten Lauf wieder an.`
          : `Cancelled — ${res.archived} rows were already moved, the rest stays in the working lists and will be offered again on the next run.`);
      } else {
      parts.push(isDe ? `${res.archived} Zeilen ins Archiv verschoben.` : `${res.archived} rows moved to the archive.`);
      }
      const perListLine = Object.keys(res.perList).filter(k => res.perList[k] > 0).map(k => `${k}: ${res.perList[k]}`).join(' · ');
      if (perListLine) parts.push(perListLine);
      if (res.failed > 0) {
        parts.push(isDe
          ? `${res.failed} Zeile(n) konnten nicht verschoben werden — sie bleiben in der Quell-Liste und werden beim nächsten Lauf erneut versucht.`
          : `${res.failed} row(s) could not be moved — they stay in the source list and will be retried on the next run.`);
      }
      setArchModal(prev => prev ? { ...prev, running: false, summary: parts } : prev);
      getArchivableCount().then(r => setArchInfo(r)).catch(() => { /* */ });
    } catch {
      setArchModal(prev => prev ? { ...prev, running: false, summary: [isDe ? 'Fehler bei der Archivierung — bitte erneut versuchen.' : 'Error during archiving — please try again.'] } : prev);
    }
  };

  // Keyframes als inline style-Tag injizieren, da SPFx SCSS-Module
  // @keyframes innerhalb von :global manchmal nicht korrekt emittieren
  React.useEffect(() => {
    const id = 'dex-orb-keyframes';
    if (!document.getElementById(id)) {
      const style = document.createElement('style');
      style.id = id;
      style.textContent = `
        @keyframes dexOrbSpin { to { transform: rotate(360deg); } }
      `;
      document.head.appendChild(style);
    }
  }, []);

  // v6.26: Die Mobile-Check-In-Bubble lebt jetzt im Header neben dem vorhandenen
  // QR-Icon — nicht mehr hier auf der LandingPage. Damit kommt sie direkt mit
  // dem App-Start und bleibt auf jeder Seite konsistent an einem festen Ort.

  return (
    <div className="landing" style={{ position: 'relative' }}>
      <span style={{
        position: 'absolute', top: 12, right: 16,
        fontSize: '0.7rem', color: 'var(--dex-gray-300)',
      }}>
        v{APP_VERSION}
      </span>

      {/* v22 / v22.45: Hinweis-Boxen oben rechts auf der Landing Page —
          gestapelt in einem gemeinsamen Container (Archivierung für Admin,
          Inaktive-Konten-Warnung für Organizer/Admin). */}
      {((isAdmin && archInfo && archInfo.total > 0) || (isAdmin && delArchCount > 0) || inactiveSummary.length > 0 || staleDrafts.length > 0 || (isAdmin && (pdWarn > 0 || pdDue > 0))) && (
      <div style={{
        ...(isMobile
          ? { position: 'static', width: '100%', margin: '0 auto 16px' }
          : { position: 'absolute', top: 34, right: 16, width: 300 }),
        maxWidth: 'calc(100vw - 32px)', zIndex: 6,
        display: 'flex', flexDirection: 'column', gap: 12,
      }}>
      {/* v22.45: Teilnehmer ohne aktives Deloitte-Konto. */}
      {inactiveSummary.length > 0 && (() => {
        const totalPeople = inactiveSummary.reduce((acc, it) => acc + it.people.length, 0);
        return (
          <div style={{
            width: '100%', boxSizing: 'border-box',
            background: '#fff3e0', border: '1px solid var(--dex-orange, #ed8b00)',
            borderRadius: 12, boxShadow: '0 8px 24px rgba(0,0,0,0.10)',
            padding: '14px 16px', textAlign: 'left',
          }}>
            <div style={{ fontWeight: 800, fontSize: '0.9rem', color: 'var(--dex-orange-dark, #b35a00)', marginBottom: 6 }}>
              {isDe ? 'Inaktive Deloitte-Konten' : 'Inactive Deloitte accounts'}
            </div>
            <p style={{ margin: '0 0 8px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
              {isDe
                ? <><strong>{totalPeople}</strong> {totalPeople === 1 ? 'Person hat' : 'Personen haben'} womöglich Deloitte verlassen — Mails/Outlook kommen ggf. nicht an. Bitte im Event prüfen und ggf. abmelden.</>
                : <><strong>{totalPeople}</strong> {totalPeople === 1 ? 'person has' : 'people have'} possibly left Deloitte — emails/Outlook may not arrive. Please review and deregister in the event.</>}
            </p>
            <ul style={{ margin: 0, padding: 0, listStyle: 'none', display: 'flex', flexDirection: 'column', gap: 6 }}>
              {inactiveSummary.map(it => (
                <li key={it.eventId}>
                  <button
                    type="button"
                    onClick={() => navigate('admin', it.eventId)}
                    title={isDe ? 'Event im Organizer Center öffnen' : 'Open event in the Organizer Center'}
                    style={{
                      display: 'block', width: '100%', textAlign: 'left', cursor: 'pointer',
                      background: 'rgba(237,139,0,0.08)', border: '1px solid rgba(237,139,0,0.4)',
                      borderRadius: 8, padding: '6px 10px', fontFamily: 'inherit',
                    }}
                  >
                    <span style={{ display: 'block', fontWeight: 700, fontSize: '0.78rem', color: 'var(--dex-gray-800)' }}>
                      {it.title} <span style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>({it.people.length})</span>
                    </span>
                    <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                      {it.people.map(p => p.name).join(', ')}
                    </span>
                  </button>
                  {/* v24.51: Organizer per Mail benachrichtigen (Dedup: nur 1x je
                      Event+Person, egal welcher Admin klickt). */}
                  {notifyResult[it.eventId] ? (
                    <div style={{ fontSize: '0.7rem', color: 'var(--dex-green-dark, #4a7c1f)', marginTop: 4, paddingLeft: 2 }}>
                      {notifyResult[it.eventId]}
                    </div>
                  ) : (
                    <button
                      type="button"
                      disabled={notifyBusyId === it.eventId}
                      onClick={async () => {
                        setNotifyBusyId(it.eventId);
                        try {
                          const res = await notifyOrganizerOfInactive(it.eventId, it.people);
                          if (res.noOrganizer) {
                            // Kein Organizer → Mail nicht möglich, Box-Eintrag bleibt (mit Hinweis).
                            setNotifyResult(prev => ({ ...prev, [it.eventId]: isDe ? 'Kein Organizer hinterlegt — keine Mail möglich.' : 'No organizer on file — no mail possible.' }));
                          } else {
                            // v24.59: Erfolgreich versendet ODER bereits benachrichtigt →
                            // Event sofort aus der Landing-Box entfernen (und den Marker
                            // greift beim nächsten Aufruf dauerhaft über den Filter).
                            setInactiveSummary(prev => prev.filter(x => x.eventId !== it.eventId));
                          }
                        } catch {
                          setNotifyResult(prev => ({ ...prev, [it.eventId]: isDe ? 'Fehler beim Versenden.' : 'Sending failed.' }));
                        } finally { setNotifyBusyId(null); }
                      }}
                      style={{
                        marginTop: 5, fontSize: '0.7rem', fontWeight: 600, cursor: 'pointer',
                        background: 'var(--dex-orange, #ed8b00)', color: '#fff', border: 'none',
                        borderRadius: 6, padding: '4px 9px', fontFamily: 'inherit', width: '100%',
                      }}
                    >
                      {notifyBusyId === it.eventId ? '…' : (isDe ? 'Organizer benachrichtigen' : 'Notify organizer')}
                    </button>
                  )}
                </li>
              ))}
            </ul>
          </div>
        );
      })()}
      {/* v24.1: Entwurf-Aufräumen — abgelaufene Entwürfe des Organizers. */}
      {staleDrafts.length > 0 && (
        <div style={{
          width: '100%', boxSizing: 'border-box',
          background: '#fff', border: '1px solid rgba(237,139,0,0.5)',
          borderRadius: 12, boxShadow: '0 8px 24px rgba(0,0,0,0.10)',
          padding: '14px 16px', textAlign: 'left',
        }}>
          <div style={{ fontWeight: 800, fontSize: '0.9rem', color: 'var(--dex-gray-800)', marginBottom: 6 }}>
            {isDe ? 'Entwürfe aufräumen' : 'Clean up drafts'}
          </div>
          <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <><strong>{staleDrafts.length}</strong> {staleDrafts.length === 1 ? 'Entwurf ist' : 'Entwürfe sind'} abgelaufen (Datum vorbei) und {staleDrafts.length === 1 ? 'wurde' : 'wurden'} nie aktiviert. Du kannst {staleDrafts.length === 1 ? 'ihn' : 'sie'} hier löschen.</>
              : <><strong>{staleDrafts.length}</strong> draft(s) expired and were never activated. You can delete them here.</>}
          </p>
          <ul style={{ margin: 0, padding: 0, listStyle: 'none', display: 'flex', flexDirection: 'column', gap: 6 }}>
            {staleDrafts.map(ev => (
              <li key={ev.id} style={{ display: 'flex', alignItems: 'center', gap: 8, justifyContent: 'space-between' }}>
                <span style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-800)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {ev.title || (isDe ? 'Ohne Titel' : 'Untitled')}
                </span>
                <button
                  type="button"
                  className="btn btn-secondary"
                  style={{ fontSize: '0.74rem', padding: '4px 10px', color: 'var(--dex-red, #c00)', flexShrink: 0 }}
                  disabled={draftDeleteBusyId === ev.id}
                  onClick={() => { deleteStaleDraft(ev).catch(() => { /* */ }); }}
                >
                  {draftDeleteBusyId === ev.id ? (isDe ? 'Löscht…' : 'Deleting…') : (isDe ? 'Löschen' : 'Delete')}
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}
      {/* v22: Archivierungs-Info (nur Admin, nur wenn Zeilen anstehen). */}
      {isAdmin && archInfo && archInfo.total > 0 && (
        <div style={{
          width: '100%', boxSizing: 'border-box',
          background: '#fff', border: '1px solid rgba(134,188,37,0.5)',
          borderRadius: 12, boxShadow: '0 8px 24px rgba(0,0,0,0.10)',
          padding: '14px 16px', textAlign: 'left',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ fontWeight: 800, fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Archivierung' : 'Archiving'}
            </span>
            <span style={{
              fontSize: '0.66rem', padding: '1px 7px', borderRadius: 999, fontWeight: 700,
              background: 'rgba(237,139,0,0.12)', color: 'var(--dex-orange, #ed8b00)',
            }}>{isDe ? 'Nur Admin' : 'Admin only'}</span>
          </div>
          <p style={{ margin: '0 0 8px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <><strong>{archInfo.total}</strong> Zeilen aus abgelaufenen Events stehen zur Archivierung an.</>
              : <><strong>{archInfo.total}</strong> rows from expired events are ready for archiving.</>}
          </p>
          <ul style={{ margin: '0 0 10px', paddingLeft: 16, fontSize: '0.72rem', color: 'var(--dex-gray-500)', lineHeight: 1.5 }}>
            {Object.keys(archInfo.perList).filter(k => archInfo.perList[k] > 0).map(k => (
              <li key={k}>{k}: {archInfo.perList[k]}</li>
            ))}
          </ul>
          <button
            className="btn btn-primary"
            style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%' }}
            onClick={() => { startArchive().catch(() => { /* */ }); }}
          >
            {isDe ? 'Jetzt archivieren' : 'Archive now'}
          </button>
        </div>
      )}
      {/* v23.40: Löschkonzept — alte Archiv-Einträge (älter als 1 Monat). */}
      {isAdmin && delArchCount > 0 && (
        <div style={{
          width: '100%', boxSizing: 'border-box',
          background: '#fff', border: '1px solid rgba(218,41,28,0.4)',
          borderRadius: 12, boxShadow: '0 8px 24px rgba(0,0,0,0.10)',
          padding: '14px 16px', textAlign: 'left',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ fontWeight: 800, fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Archiv aufräumen' : 'Clean up archive'}
            </span>
            <span style={{
              fontSize: '0.66rem', padding: '1px 7px', borderRadius: 999, fontWeight: 700,
              background: 'rgba(237,139,0,0.12)', color: 'var(--dex-orange, #ed8b00)',
            }}>{isDe ? 'Nur Admin' : 'Admin only'}</span>
          </div>
          <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <><strong>{delArchCount}</strong> Archiv-Einträge sind älter als 1 Monat und können endgültig gelöscht werden.</>
              : <><strong>{delArchCount}</strong> archive entries are older than 1 month and can be permanently deleted.</>}
          </p>
          <button
            className="btn btn-secondary"
            style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%', color: 'var(--dex-red, #c00)' }}
            disabled={delArchBusy}
            onClick={() => { startDeleteOldArchive().catch(() => { /* */ }); }}
          >
            {delArchBusy ? (isDe ? 'Wird gelöscht…' : 'Deleting…') : (isDe ? 'Alte Einträge löschen' : 'Delete old entries')}
          </button>
        </div>
      )}
      {/* v26.32: Löschkonzept — Vorwarnung (Teilnehmerliste wird in ~1 Woche gelöscht). */}
      {isAdmin && pdWarn > 0 && (
        <div style={{
          width: '100%', boxSizing: 'border-box',
          background: '#fff', border: '1px solid rgba(237,139,0,0.5)',
          borderRadius: 12, boxShadow: '0 8px 24px rgba(0,0,0,0.10)',
          padding: '14px 16px', textAlign: 'left',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ fontWeight: 800, fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Teilnehmerlisten laufen ab' : 'Attendee lists expiring'}
            </span>
            <span style={{
              fontSize: '0.66rem', padding: '1px 7px', borderRadius: 999, fontWeight: 700,
              background: 'rgba(237,139,0,0.12)', color: 'var(--dex-orange, #ed8b00)',
            }}>{isDe ? 'Nur Admin' : 'Admin only'}</span>
          </div>
          <p style={{ margin: '0 0 4px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <>Bei <strong>{pdWarn}</strong> {pdWarn === 1 ? 'Event' : 'Events'} wird die Teilnehmerliste in etwa einer Woche gelöscht (3 Monate nach dem Event). Die Organizer wurden automatisch informiert, die Liste noch herunterzuladen. Event &amp; Kennzahlen bleiben im Statistik-Archiv erhalten.</>
              : <>For <strong>{pdWarn}</strong> event(s) the attendee list will be deleted in about a week (3 months after the event). The organizers were notified automatically to download it. Event &amp; KPIs are kept in the statistics archive.</>}
          </p>
        </div>
      )}
      {/* v26.32: Löschkonzept — fällige Teilnehmerlisten löschen (Event bleibt, KPIs ins Archiv). */}
      {isAdmin && pdDue > 0 && (
        <div style={{
          width: '100%', boxSizing: 'border-box',
          background: '#fff', border: '1px solid rgba(218,41,28,0.4)',
          borderRadius: 12, boxShadow: '0 8px 24px rgba(0,0,0,0.10)',
          padding: '14px 16px', textAlign: 'left',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ fontWeight: 800, fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Teilnehmerlisten löschen' : 'Delete attendee lists'}
            </span>
            <span style={{
              fontSize: '0.66rem', padding: '1px 7px', borderRadius: 999, fontWeight: 700,
              background: 'rgba(237,139,0,0.12)', color: 'var(--dex-orange, #ed8b00)',
            }}>{isDe ? 'Nur Admin' : 'Admin only'}</span>
          </div>
          <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <>Bei <strong>{pdDue}</strong> {pdDue === 1 ? 'Event ist' : 'Events sind'} die Aufbewahrungsfrist (3 Monate) abgelaufen. Beim Löschen werden die wichtigsten Kennzahlen ins Statistik-Archiv übernommen und dann die Teilnehmerliste entfernt — das Event bleibt erhalten.</>
              : <><strong>{pdDue}</strong> event(s) have passed the 3-month retention. Deleting archives the key KPIs to the statistics archive and then removes the attendee list — the event itself is kept.</>}
          </p>
          <button
            className="btn btn-secondary"
            style={{ fontSize: '0.82rem', padding: '8px 16px', width: '100%', color: 'var(--dex-red, #c00)' }}
            disabled={pdBusy}
            onClick={() => { startParticipantDeletion().catch(() => { /* */ }); }}
          >
            {pdBusy ? (isDe ? 'Wird gelöscht…' : 'Deleting…') : (isDe ? 'Teilnehmerlisten löschen' : 'Delete attendee lists')}
          </button>
        </div>
      )}
      </div>
      )}

      {/* v22: Fortschritts-/Ergebnis-Modal der Archivierung. */}
      {/* v26.40: Modal-Hinweis nach automatischer Abmeldung von Ex-Deloitte-Personen. */}
      {autoDeregModal && (
        <Modal
          open={true}
          onClose={() => setAutoDeregModal(null)}
          dismissable={true}
          maxWidth={560}
          ariaLabel={isDe ? 'Automatisch abgemeldet' : 'Auto-deregistered'}
        >
          <h3 style={{ margin: '0 0 10px', fontSize: '1.05rem' }}>
            {isDe ? 'Automatisch abgemeldet: Konten ohne Deloitte-Zugang' : 'Auto-deregistered: accounts without Deloitte access'}
          </h3>
          <p style={{ margin: '0 0 12px', fontSize: '0.9rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
            {isDe
              ? 'Für die folgenden Events ist eingestellt, dass Personen ohne aktives Deloitte-Konto automatisch abgemeldet werden. Diese Personen wurden soeben abgemeldet (die Warteliste rückt ggf. nach):'
              : 'These events are set to auto-deregister people without an active Deloitte account. The following people were just deregistered (the waitlist may move up):'}
          </p>
          {autoDeregModal.map((ev, i) => (
            <div key={i} style={{ marginBottom: 12 }}>
              <div style={{ fontWeight: 700, fontSize: '0.92rem', marginBottom: 4 }}>{ev.title}</div>
              <ul style={{ margin: 0, paddingLeft: 18, fontSize: '0.88rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {ev.people.map((p, j) => <li key={j}>{p.name}{p.name && p.name !== p.email ? ` (${p.email})` : ''}</li>)}
              </ul>
            </div>
          ))}
          <div style={{ marginTop: 14, textAlign: 'right' }}>
            <button className="btn btn-primary" style={{ fontSize: '0.88rem', padding: '9px 18px' }} onClick={() => setAutoDeregModal(null)}>
              {isDe ? 'Verstanden' : 'Got it'}
            </button>
          </div>
        </Modal>
      )}
      {archModal && (
        <Modal
          open={true}
          onClose={() => { if (!archModal.running) setArchModal(null); }}
          dismissable={!archModal.running}
          maxWidth={520}
          ariaLabel={isDe ? 'Archivierung' : 'Archiving'}
        >
          <h3 style={{ margin: '0 0 10px', fontSize: '1.05rem' }}>
            {isDe ? 'Archivierung abgelaufener Event-Zeilen' : 'Archiving expired event rows'}
          </h3>
          {archModal.running ? (
            <>
              <p style={{ margin: '0 0 6px', fontSize: '0.9rem', color: 'var(--dex-gray-700)' }}>
                {isDe ? 'Liste' : 'List'} {archModal.listIdx}/{archModal.listTotal}: <strong>{archModal.listName || '…'}</strong>
              </p>
              <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>
                {archModal.total > 0
                  ? (isDe ? `Zeile ${archModal.done}/${archModal.total} wird verschoben…` : `Moving row ${archModal.done}/${archModal.total}…`)
                  : (isDe ? 'Liste wird geladen…' : 'Loading list…')}
              </p>
              {(() => {
                const base = Math.max(0, archModal.listIdx - 1);
                const inner = archModal.total > 0 ? archModal.done / archModal.total : 0;
                const pct = Math.min(100, Math.round(((base + inner) / Math.max(1, archModal.listTotal)) * 100));
                return (
                  <div style={{ background: 'var(--dex-gray-100, #f0f0f0)', borderRadius: 999, height: 10, overflow: 'hidden' }}>
                    <div style={{ width: `${pct}%`, height: '100%', background: 'var(--dex-green, #86bc25)', borderRadius: 999, transition: 'width 0.2s ease' }} />
                  </div>
                );
              })()}
              <p style={{ margin: '10px 0 0', fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                {isDe ? 'Bitte das Fenster geöffnet lassen, bis die Archivierung abgeschlossen ist.' : 'Please keep this window open until archiving completes.'}
              </p>
              {/* v22.2: Abbrechen — stoppt nach der aktuellen Zeile. */}
              <div style={{ textAlign: 'right', marginTop: 12 }}>
                <button
                  className="btn btn-secondary"
                  disabled={archCancelRequested}
                  onClick={() => { archCancelRef.current = true; setArchCancelRequested(true); }}
                  style={{ fontSize: '0.85rem' }}
                >
                  {archCancelRequested
                    ? (isDe ? 'Wird abgebrochen…' : 'Cancelling…')
                    : (isDe ? 'Abbrechen' : 'Cancel')}
                </button>
              </div>
            </>
          ) : (
            <>
              <ul style={{ margin: '0 0 14px', paddingLeft: 18, fontSize: '0.88rem', color: 'var(--dex-gray-700)', lineHeight: 1.6 }}>
                {(archModal.summary || []).map((s, i) => <li key={i}>{s}</li>)}
              </ul>
              <div style={{ textAlign: 'right' }}>
                <button className="btn btn-primary" onClick={() => setArchModal(null)} style={{ fontSize: '0.88rem', padding: '9px 18px' }}>
                  {isDe ? 'Schließen' : 'Close'}
                </button>
              </div>
            </>
          )}
        </Modal>
      )}

      <div className="landing__hero">
        <div className="landing__card" style={{ position: 'relative' }}>
          {/* Sprachauswahl - oben links in der weissen Card */}
          <div style={{
            position: 'absolute', top: 16, left: 16,
            display: 'flex', gap: 4,
          }}>
            <button
              onClick={() => setLocale('de')}
              style={{
                background: locale === 'de' ? 'var(--dex-green)' : 'var(--dex-gray-100)',
                border: 'none',
                borderRadius: 6, padding: '6px 12px', cursor: 'pointer', fontSize: '0.8rem',
                color: locale === 'de' ? '#fff' : 'var(--dex-gray-500)',
                fontWeight: locale === 'de' ? 700 : 500,
                transition: 'all 0.2s',
              }}
              title="Deutsch"
            >
              DE
            </button>
            <button
              onClick={() => setLocale('en')}
              style={{
                background: locale === 'en' ? 'var(--dex-green)' : 'var(--dex-gray-100)',
                border: 'none',
                borderRadius: 6, padding: '6px 12px', cursor: 'pointer', fontSize: '0.8rem',
                color: locale === 'en' ? '#fff' : 'var(--dex-gray-500)',
                fontWeight: locale === 'en' ? 700 : 500,
                transition: 'all 0.2s',
              }}
              title="English"
            >
              EN
            </button>
          </div>
          <div className="landing__orb">
            <div className="landing__orb-inner" />
          </div>
          <div className="landing__text">
            <h1>
              {greeting}{firstName ? <>, <strong>{firstName}</strong></> : ''}.
            </h1>
            <p>
              {isDe
                ? <>Willkommen bei <strong>DEX</strong>. Unsere neue App für die Organisation von <span style={{ whiteSpace: 'nowrap' }}>Deloitte Events</span>. Von der Anmeldung, bis zum Check-in. Alles an einer Stelle.</>
                : <>Welcome to <strong>DEX</strong>. Our new app for organising <span style={{ whiteSpace: 'nowrap' }}>Deloitte events</span>. From registration to check-in. Everything in one place.</>}
            </p>
          </div>
          {/* v22.1: Check-in-Hinweisbox(en) — ab 2 Tage vor dem Event, sobald
              der eigene QR-Code versendet wurde. Klick auf den kleinen QR
              öffnet ihn groß im Modal (zum Vorzeigen am Eingang). */}
          {checkInBoxes.map(box => (
            <button
              key={box.eventId}
              type="button"
              onClick={() => { openBigQr(box).catch(() => { /* */ }); }}
              style={{
                display: 'flex', alignItems: 'center', gap: 14, width: '100%',
                textAlign: 'left', marginBottom: 12, padding: '12px 14px',
                background: 'rgba(134,188,37,0.08)',
                border: '1.5px solid var(--dex-green, #86bc25)', borderRadius: 12,
                cursor: 'pointer',
              }}
              title={isDe ? 'QR-Code groß anzeigen' : 'Show QR code enlarged'}
            >
              <img src={box.qrSmall} alt="QR" style={{ width: 66, height: 66, flexShrink: 0, borderRadius: 6, background: '#fff', border: '1px solid var(--dex-gray-200)' }} />
              <span style={{ minWidth: 0 }}>
                <span style={{ display: 'block', fontWeight: 700, fontSize: '0.92rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
                  {isDe ? 'Check-in für' : 'Check-in for'} {box.title}
                </span>
                <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                  {isDe
                    ? 'Dein persönlicher QR-Code — antippen zum Vergrößern und am Eingang vorzeigen.'
                    : 'Your personal QR code — tap to enlarge and show at the entrance.'}
                </span>
              </span>
            </button>
          ))}
          {/* v24.13: „Du bist angemeldet"-Boxen — aktive Events, für die man
              angemeldet ist; ausgeblendet, wenn schon eine Check-in/QR-Box läuft. */}
          {(() => {
            const checkInIds = new Set(checkInBoxes.map(b => b.eventId));
            const list = myRegBoxes.filter(b => !checkInIds.has(b.eventId));
            return list.map(box => {
              const startTs = new Date(box.startDate).getTime();
              const diff = startTs - nowTick;
              let countdown: string; let urgent = false;
              if (diff <= 0) { countdown = isDe ? 'läuft gerade' : 'happening now'; urgent = true; }
              else if (diff < 24 * 60 * 60 * 1000) { const h = Math.max(1, Math.ceil(diff / (60 * 60 * 1000))); countdown = isDe ? `noch ${h} ${h === 1 ? 'Stunde' : 'Stunden'}` : `in ${h} ${h === 1 ? 'hour' : 'hours'}`; urgent = true; }
              else { const d = Math.floor(diff / (24 * 60 * 60 * 1000)); countdown = isDe ? `noch ${d} ${d === 1 ? 'Tag' : 'Tage'}` : `in ${d} ${d === 1 ? 'day' : 'days'}`; }
              const dateLabel = new Date(box.startDate).toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
              return (
                <button
                  key={box.eventId}
                  type="button"
                  onClick={() => navigate('my-events')}
                  style={{ display: 'flex', alignItems: 'center', gap: 14, width: '100%', textAlign: 'left', marginBottom: 12, padding: '12px 14px', background: '#fff', border: '1.5px solid var(--dex-gray-200)', borderRadius: 12, cursor: 'pointer' }}
                  title={isDe ? 'Zu „Meine Events"' : 'Go to My Events'}
                >
                  <div style={{ width: 66, height: 66, flexShrink: 0, borderRadius: 8, background: box.imageUrl ? `url(${box.imageUrl}) center/cover no-repeat` : 'linear-gradient(135deg, var(--dex-green, #86bc25), var(--dex-blue, #0076a8))', border: '1px solid var(--dex-gray-200)' }} />
                  <span style={{ minWidth: 0, flex: 1 }}>
                    <span style={{ display: 'block', fontWeight: 700, fontSize: '0.92rem', color: 'var(--dex-gray-800)' }}>
                      {isDe
                        ? <>Du bist für <span style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>&bdquo;{box.title}&ldquo;</span> angemeldet</>
                        : <>You are registered for <span style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>&bdquo;{box.title}&ldquo;</span></>}
                    </span>
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                      {dateLabel}{box.location ? ` · ${box.location}` : ''}
                    </span>
                  </span>
                  <span style={{ flexShrink: 0, fontSize: '0.78rem', fontWeight: 800, padding: '4px 10px', borderRadius: 999, background: urgent ? 'rgba(218,41,28,0.10)' : 'rgba(134,188,37,0.12)', color: urgent ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)' }}>
                    {countdown}
                  </span>
                </button>
              );
            });
          })()}
          <button className="btn btn-lg btn-block btn-outline" data-tour="landing-start" onClick={() => navigate('start')} style={{ maxWidth: 360 }}>
            {t('landing.start')}
          </button>
          {/* v26.37: „DEX für dein Event nutzen"-Box auf dem Handy ausblenden —
              spart Platz; die Anfrage-Funktion bleibt am Desktop erhalten. */}
          {showOrganizerCta && !isMobile && (
          <div
            className="landing__actions"
            style={{
              display: 'flex', gap: 12, justifyContent: 'center', alignItems: 'center',
              flexWrap: 'wrap',
            }}
          >
            {/* v24.19: „DEX für dein Event nutzen"-Box (öffnet die Anfrage),
                mit integriertem DEX-Orb-Logo. Ersetzt die früheren Info-/Mail-
                Icons + die Tutorial-Bubble (das Tutorial liegt jetzt mittig im
                Header). „Über die App" steht jetzt als Textlink unter den
                Entwickler-Namen. */}
            <button
              type="button"
              onClick={() => setShowInquiry(true)}
              onMouseEnter={() => setCtaHover(true)}
              onMouseLeave={() => setCtaHover(false)}
              className="landing__bubble"
              style={{
                display: 'flex', alignItems: 'center', gap: 12,
                background: 'var(--dex-green)', color: '#fff',
                padding: '12px 18px', borderRadius: 14,
                boxShadow: ctaHover ? '0 4px 14px rgba(0,0,0,0.16)' : '0 2px 8px rgba(0,0,0,0.10)',
                border: ctaHover ? '2px solid var(--dex-green-dark, #4a7c1f)' : '2px solid transparent',
                cursor: 'pointer', textAlign: 'left', transition: 'all 0.15s ease',
                transform: ctaHover ? 'translateY(-1px)' : 'none',
                fontFamily: 'inherit', width: '100%', maxWidth: 360,
              }}
              title={locale === 'de' ? 'DEX App für dein Event anfragen' : 'Request the DEX App for your event'}
            >
              <span style={{
                position: 'relative', flexShrink: 0, width: 38, height: 38, borderRadius: '50%',
                background: 'conic-gradient(from 0deg, #86bc25, #0076a8 90deg, #00bcd4 150deg, #4caf50 210deg, #8bc34a 270deg, #ffeb3b 320deg, #86bc25 360deg)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                boxShadow: '0 0 0 2px rgba(255,255,255,0.45)',
              }}>
                <span style={{ width: 16, height: 16, borderRadius: '50%', background: '#fff' }} />
              </span>
              <span style={{ minWidth: 0 }}>
                <span style={{ display: 'block', fontWeight: 800, fontSize: '0.98rem', lineHeight: 1.25 }}>
                  {locale === 'de' ? 'DEX für dein Event nutzen' : 'Use DEX for your event'}
                </span>
                <span style={{ display: 'block', fontSize: '0.82rem', opacity: 0.95, lineHeight: 1.3, marginTop: 2 }}>
                  {locale === 'de' ? 'Werde Organizer und nutze alle Funktionen.' : 'Become an organizer and use all features.'}
                </span>
              </span>
            </button>
          </div>
          )}

          <div
            style={{
              fontSize: '0.95rem',
              color: 'var(--dex-gray-400)',
              textAlign: 'center',
              marginTop: 6,
              lineHeight: 1.3,
            }}
          >
            {locale === 'de' ? 'Entwickelt von ' : 'Built by '}
            <span style={{ fontWeight: 600, color: 'var(--dex-gray-500)' }}>
              <DevName name="Eike Brenneisen" email="ebrenneisen@deloitte.de" isMobile={isMobile} />
              {' '}{locale === 'de' ? 'und' : 'and'}{' '}
              <DevName name="Nils Felten" email="nifelten@deloitte.de" isMobile={isMobile} />
            </span>
          </div>
        </div>
      </div>
      {/* v22.1: Groß-Ansicht des persönlichen Check-in-QR (Klick auf die
          Hinweisbox) — gleicher Aufbau wie „Mein QR-Code" in Meine Events. */}
      {qrBigModal && (
        <Modal
          open={true}
          onClose={() => setQrBigModal(null)}
          maxWidth={420}
          ariaLabel={isDe ? 'Mein QR-Code' : 'My QR code'}
        >
          <h3 style={{ margin: 0, fontSize: '1.05rem', textAlign: 'center' }}>
            {isDe ? 'Mein Check-in-QR-Code' : 'My check-in QR code'}
          </h3>
          <p style={{ margin: 0, fontSize: '0.82rem', color: 'var(--dex-gray-500)', textAlign: 'center' }}>
            {qrBigModal.title}
          </p>
          <div style={{ textAlign: 'center' }}>
            <img
              src={qrBigModal.dataUrl}
              alt="QR-Code"
              style={{ width: 280, maxWidth: '90%', height: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 10, background: '#fff' }}
            />
          </div>
          <p style={{ margin: 0, textAlign: 'center', fontWeight: 700, fontSize: '0.95rem' }}>
            {qrBigModal.name}
            {qrBigModal.tid ? <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}> · Nr. {qrBigModal.tid}</span> : null}
          </p>
          <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)', textAlign: 'center', lineHeight: 1.5 }}>
            {isDe
              ? 'Zeig diesen Code am Eingang dem Check-in-Team — er ist derselbe wie in deiner QR-Mail. Falls der Scan nicht klappt, reicht dein Name.'
              : 'Show this code to the check-in team at the entrance — it is the same as in your QR email. If the scan fails, your name is enough.'}
          </p>
          <div style={{ textAlign: 'right' }}>
            <button className="btn btn-secondary" onClick={() => setQrBigModal(null)} style={{ fontSize: '0.85rem' }}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          </div>
        </Modal>
      )}

      {/* v13.3: Inquiry-Modal aus der wiederverwendbaren Komponente. */}
      <InquiryModal open={showInquiry} onClose={() => setShowInquiry(false)} />
    </div>
  );
}

// v11.47: KPI-Box-Reihe ueber dem "Entwickelt von ..."-Block. Drei Boxen
// nebeneinander: gehostete Events, Teilnehmer, App-Aufrufe. Jede Box mit
// einer AnimatedCounter-Komponente, die beim ersten Verfügbarwerden des
// Werts von 0 zum Zielwert ease-out hochzählt (~1.6s). Solange Daten noch
// laden, steht ein dezenter Skeleton-Punkt im Wert-Feld.
export function KpiRow(props: {
  locale: string;
  eventsLoading: boolean;
  participantsLoading: boolean;
  events: number;
  participants: number;
}): React.ReactElement {
  const isDe = props.locale === 'de';
  const labels = isDe
    ? { ev: 'Events', pa: 'Teilnehmer' }
    : { ev: 'Events', pa: 'Attendees' };
  return (
    <div style={{
      display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 10,
    }}>
      <KpiBox label={labels.ev} value={props.events} loading={props.eventsLoading} />
      <KpiBox label={labels.pa} value={props.participants} loading={props.participantsLoading} />
    </div>
  );
}

function KpiBox(props: { label: string; value: number; loading: boolean }): React.ReactElement {
  return (
    <div style={{
      display: 'flex', flexDirection: 'column', alignItems: 'center',
      padding: '12px 8px',
      background: 'linear-gradient(135deg, rgba(134,188,37,0.08), rgba(0,118,168,0.04))',
      border: '1px solid var(--dex-gray-200)',
      borderRadius: 12,
      minWidth: 0,
    }}>
      <div style={{
        fontSize: 'clamp(1.4rem, 4vw, 1.9rem)', fontWeight: 800,
        color: 'var(--dex-green-dark, #4a7c1f)',
        lineHeight: 1.1, letterSpacing: '-0.02em',
        fontVariantNumeric: 'tabular-nums',
      }}>
        {props.loading
          ? <SkeletonDots />
          : <AnimatedCounter value={props.value} />}
      </div>
      <div style={{
        marginTop: 6,
        fontSize: '0.72rem', color: 'var(--dex-gray-500)',
        fontWeight: 500, textAlign: 'center', lineHeight: 1.2,
      }}>
        {props.label}
      </div>
    </div>
  );
}

function SkeletonDots(): React.ReactElement {
  return (
    <span aria-hidden="true" style={{ display: 'inline-flex', gap: 4, opacity: 0.55 }}>
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', animation: 'dexKpiPulse 1.2s ease-in-out infinite' }} />
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', animation: 'dexKpiPulse 1.2s ease-in-out 0.2s infinite' }} />
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'currentColor', animation: 'dexKpiPulse 1.2s ease-in-out 0.4s infinite' }} />
      <style>{`
        @keyframes dexKpiPulse {
          0%, 100% { transform: scale(0.6); opacity: 0.3; }
          50%      { transform: scale(1);   opacity: 1; }
        }
      `}</style>
    </span>
  );
}

/** Ease-out-Cubic-Animation von 0 (bzw. dem letzten geanimierten Wert) auf
 *  `value`. Wenn `value` sich aendert, startet eine neue Animation; bei
 *  schnellen Aenderungen wird die laufende Animation sauber durch eine neue
 *  ersetzt (requestAnimationFrame + AbortFlag). */
function AnimatedCounter(props: { value: number; durationMs?: number }): React.ReactElement {
  const target = Math.max(0, Math.floor(props.value || 0));
  // v11.79: Default-Dauer von 1600 ms → 600 ms reduziert. Seit der App-Boot
  // unter ~1.6 s liegt, wirkte das längere Count-Up-Tempo träge; die Zahl
  // tickt jetzt knackiger hoch ohne hektisch zu wirken.
  const duration = props.durationMs ?? 600;
  const [shown, setShown] = React.useState<number>(0);
  const startRef = React.useRef<number>(0);
  const fromRef = React.useRef<number>(0);
  const rafRef = React.useRef<number | null>(null);
  React.useEffect(() => {
    if (rafRef.current !== null) cancelAnimationFrame(rafRef.current);
    fromRef.current = shown;
    startRef.current = 0;
    const step = (ts: number): void => {
      if (startRef.current === 0) startRef.current = ts;
      const elapsed = ts - startRef.current;
      const t = Math.min(1, elapsed / duration);
      // ease-out-cubic
      const eased = 1 - Math.pow(1 - t, 3);
      const value = fromRef.current + (target - fromRef.current) * eased;
      setShown(Math.round(value));
      if (t < 1) {
        rafRef.current = requestAnimationFrame(step);
      } else {
        rafRef.current = null;
      }
    };
    rafRef.current = requestAnimationFrame(step);
    return () => {
      if (rafRef.current !== null) cancelAnimationFrame(rafRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [target, duration]);
  return <span>{shown.toLocaleString('de-DE')}</span>;
}

// v7.1: Entwickler-Name mit Hover-Popover auf der LandingPage.
// Analog zu OrganizerList: bei Hover erscheint ein kleines Popover mit
// größerem Foto, Name und Rolle/Standort. Die Daten (jobTitle + location)
// werden per SharePoint-User-Profil-Lookup live nachgeladen (einmalig
// beim ersten Hover) — damit zeigen wir immer den aktuellen Rollen-Stand
// aus dem AD, keine hardcoded Strings.
function DevName(props: { name: string; email: string; isMobile?: boolean }): React.ReactElement {
  const [hovered, setHovered] = React.useState(false);
  const [failed, setFailed] = React.useState(false);
  const [profile, setProfile] = React.useState<{ jobTitle: string; location: string } | null>(null);
  const loadedRef = React.useRef(false);
  const wrapRef = React.useRef<HTMLSpanElement>(null);
  const isMobile = !!props.isMobile;
  // v26.37: Auf dem Handy gibt es kein Hover — die Karte wird per Tap geöffnet
  // und schließt beim Tap außerhalb.
  React.useEffect(() => {
    if (!isMobile || !hovered) return;
    const onDocDown = (e: Event): void => {
      if (wrapRef.current && e.target instanceof Node && !wrapRef.current.contains(e.target)) {
        setHovered(false);
      }
    };
    document.addEventListener('pointerdown', onDocDown);
    return () => document.removeEventListener('pointerdown', onDocDown);
  }, [isMobile, hovered]);
  React.useEffect(() => {
    if (!hovered || loadedRef.current) return;
    loadedRef.current = true;
    (async () => {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ctx = (window as any).__dexSpfxContext;
        if (!ctx) return;
        const { SharePointService } = await import('../services/SharePointService');
        const svc = new SharePointService(ctx);
        const result = await svc.searchUserByEmail(props.email);
        if (result) {
          setProfile({
            jobTitle: result.jobTitle || '',
            location: result.location || '',
          });
        }
      } catch { /* Profil-Lookup fehlgeschlagen — Fallback bleibt leer */ }
    })().catch(() => { /* ignore */ });
  }, [hovered, props.email]);

  const roleLine = profile
    ? [profile.jobTitle, profile.location].filter(Boolean).join(' · ')
    : '';

  return (
    <span
      ref={wrapRef}
      style={{ position: 'relative', cursor: isMobile ? 'pointer' : 'default', display: 'inline-block' }}
      onMouseEnter={isMobile ? undefined : () => setHovered(true)}
      onMouseLeave={isMobile ? undefined : () => setHovered(false)}
      onClick={isMobile ? (e) => { e.stopPropagation(); setHovered(h => !h); } : undefined}
    >
      <span style={{ textDecoration: hovered ? 'underline' : 'none', textDecorationColor: 'var(--dex-green)' }}>
        {props.name}
      </span>
      {hovered && (
        <span style={{
          position: 'absolute', bottom: 'calc(100% + 10px)', left: '50%',
          transform: 'translateX(-50%)',
          background: '#fff', borderRadius: 10, padding: '10px 12px',
          boxShadow: '0 8px 24px rgba(0,0,0,0.15)',
          display: 'flex', alignItems: 'center', gap: 10,
          whiteSpace: isMobile ? 'normal' : 'nowrap',
          maxWidth: isMobile ? 'calc(100vw - 24px)' : undefined,
          width: isMobile ? 'max-content' : undefined,
          zIndex: 20, pointerEvents: 'none',
          border: '1px solid var(--dex-gray-200)',
        }}>
          {failed ? (
            <span style={{
              width: 48, height: 48, borderRadius: '50%',
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              background: 'linear-gradient(135deg, #86bc25, #0076a8)',
              color: '#fff', fontWeight: 700, fontSize: '0.9rem', flexShrink: 0,
            }}>
              {props.name.split(' ').map(p => p[0]).slice(0, 2).join('')}
            </span>
          ) : (
            <img
              src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(props.email)}&size=L`}
              alt={props.name}
              onError={() => setFailed(true)}
              style={{ width: 48, height: 48, borderRadius: '50%', objectFit: 'cover', flexShrink: 0, background: 'var(--dex-gray-200)' }}
            />
          )}
          <span style={{ textAlign: 'left', fontSize: '0.82rem', lineHeight: 1.35 }}>
            <span style={{ display: 'block', fontWeight: 700, color: 'var(--dex-gray-800)' }}>{props.name}</span>
            {roleLine && (
              <span style={{ display: 'block', color: 'var(--dex-gray-500)', fontSize: '0.75rem' }}>{roleLine}</span>
            )}
            <span style={{ display: 'block', color: 'var(--dex-gray-400)', fontSize: '0.72rem' }}>{props.email}</span>
          </span>
        </span>
      )}
    </span>
  );
}
