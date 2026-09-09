// Landing Page - Startbildschirm mit animiertem Orb und Willkommensnachricht

import * as React from 'react';
import DexLogo from './DexLogo';
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
import { AlertCircle, ChevronDown, GraduationCap } from './Icons';
import { INACTIVE_SUMMARY_CACHE_KEY } from '../utils/accountCheckCache';
import { cx, ensureDexUiStyles } from './dexUi';

export default function LandingPage(): React.ReactElement {
  // v31.4 (Review): Der Hinweiskasten „Code nicht ladbar" unten nutzt
  // `dex-ui-callout` — das Stylesheet hängt sonst an einem geöffneten Modal.
  ensureDexUiStyles();
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
  // v31.9: Der Hover-State der „DEX für dein Event nutzen"-Box (v26) ist weg —
  // den Hover macht jetzt `dex-ui-card--hover`. Ein State, der nur die Optik
  // umschaltet, rendert die ganze Seite bei jeder Mausbewegung neu.
  // v31.9: Aufklapper der Verwaltungs-Hinweise. Sie sind Organizer-Arbeit und
  // standen als bis zu sechs gestapelte Kästen mit vier destruktiven Knöpfen
  // über der Seite; zugeklappt nennt der Kopf die Zahl, offen steht alles da.
  const [adminHintsOpen, setAdminHintsOpen] = React.useState(false);

  // ==================== v22: Archivierung (Admin) ====================
  const { isAdmin, canCreateEvents } = useRoles();
  // v24.23: Die „DEX für dein Event nutzen"-Box (Werde Organizer) ist für
  // Organizer überflüssig — sie haben die Funktionen schon.
  // v30.68: Auch für Admins aus (Nutzer-Ansage 02.09.2026) — sie sind
  // Organizer mit mehr Rechten; die User-Ansicht prüft man über die
  // Rollen-Vorschau, nicht über einen Kasten, der einen selbst wirbt.
  const showOrganizerCta = !canCreateEvents;
  const { isEventsLoading, getArchivableCount, runArchiveExpired, scanInactiveAccounts, notifyOrganizerOfInactive, autoDeregisterInactive, getSentInactiveNotices, getDeletableArchiveCount, runDeleteOldArchive, getParticipantDeletionWarnings, getParticipantDeletionDue, runParticipantDeletion, maybeSendParticipantDeletionWarnings, deleteEvent, getLastEventDeleteError, countExternalRegistrations, refreshEvents, getAllRegistrations } = useEvents();
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
  // v26.32: Löschkonzept — Teilnehmerlisten (3 Monate): Vorwarn- + Fällig-Liste.
  // v28.17: Es werden die betroffenen Events selbst aufgehoben (nicht nur die
  // Anzahl), damit die Admin-Boxen die Event-Titel nennen können.
  const [pdWarnEvents, setPdWarnEvents] = React.useState<Array<{ id: string; title: string; startDate?: string; endDate?: string }>>([]);
  const [pdDueEvents, setPdDueEvents] = React.useState<Array<{ id: string; title: string; startDate?: string; endDate?: string }>>([]);
  const pdWarn = pdWarnEvents.length;
  const pdDue = pdDueEvents.length;
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
      .then(list => { if (!cancelled) setPdWarnEvents(list); })
      .catch(() => { /* best-effort */ });
    getParticipantDeletionDue()
      .then(list => { if (!cancelled) setPdDueEvents(list); })
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
    const pdDueTitles = pdDueEvents.map(e => `• ${e.title}`).join('\n');
    const ok = await confirmDialog(
      isDe
        ? `Bei ${pdDue} ${pdDue === 1 ? 'Event' : 'Events'} die Teilnehmerliste endgültig löschen?\n\n${pdDueTitles}\n\nDie wichtigsten Kennzahlen werden zuvor ins Statistik-Archiv übernommen. Das Event bleibt bestehen, die Teilnehmerliste wird in den SharePoint-Papierkorb verschoben.`
        : `Delete the attendee list for ${pdDue} event(s)?\n\n${pdDueTitles}\n\nThe key KPIs are archived to the statistics archive first. The event is kept; the attendee list is moved to the SharePoint recycle bin.`,
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
      try { const list = await getParticipantDeletionDue(); setPdDueEvents(list); } catch { /* */ }
      try { const w = await getParticipantDeletionWarnings(); setPdWarnEvents(w); } catch { /* */ }
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
      // v30.67 (Review): deleteEvent liefert false, wenn bewusst nichts (oder
      // nur ein Teil) gelöscht wurde — Termine nicht lesbar, ein Termin nicht
      // löschbar, Klammer-Recycle abgelehnt. Das stand vorher grün als
      // „Entwurf gelöscht" da, und der Entwurf blieb in der Liste.
      const gone = await deleteEvent(ev.id);
      if (!gone) {
        showAlert(getLastEventDeleteError(isDe ? 'de' : 'en')
          || (isDe ? 'Löschen fehlgeschlagen — bitte erneut versuchen.' : 'Deletion failed — please try again.'), { variant: 'error' });
        return;
      }
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
  /**
   * v30.42: Sub-Events auf ihre Klammer zusammenziehen.
   *
   * Der Scan läuft je Event, und bei einem Klammer-Event ist jeder Kalender-Tag
   * ein eigenes Event. Dieselbe Person erschien deshalb einmal PRO Tag — bei
   * 19 Office-Tagen also bis zu 19 Kacheln für einen einzigen Befund (Nutzer-
   * Screenshot: vier Kacheln, alle „Jiva Dimitrova-Micha"). Für den Organizer
   * ist das EINE Information: Diese Person hat womöglich Deloitte verlassen.
   *
   * Zusammengezogen wird ABSICHTLICH erst nach dem Aufteilen in
   * `autoItems`/`notifyItems`: Das automatische Abmelden muss weiter am
   * einzelnen Termin ansetzen (dort liegt die Anmeldung), nur die ANZEIGE und
   * die Organizer-Mail gehören auf die Klammer. Damit landet auch der
   * Dedup-Merker auf der Klammer und die Kacheln kommen nach dem Benachrichtigen
   * nicht einzeln zurück.
   */
  const collapseToUmbrella = React.useCallback((items: InactiveItem[]): InactiveItem[] => {
    const byId: Record<string, typeof events[number]> = {};
    (events || []).forEach(e => { byId[e.id] = e; });
    const out: InactiveItem[] = [];
    const idx: Record<string, number> = {};
    for (const it of items) {
      const own = byId[it.eventId];
      const parent = own?.parentEventId ? byId[own.parentEventId] : undefined;
      // Nur echte Klammer-Events einsammeln. Ein Sub-Event unter einem normalen
      // Hauptevent bleibt eigenständig — dort ist der Termin die Aussage.
      const root = (parent && parent.subEventsOnlyMode) ? parent : own;
      const rootId = root?.id || it.eventId;
      const rootTitle = root?.title || it.title;
      if (idx[rootId] === undefined) {
        idx[rootId] = out.length;
        out.push({ eventId: rootId, title: rootTitle, people: [] });
      }
      const bucket = out[idx[rootId]];
      for (const p of it.people) {
        const key = (p.email || '').toLowerCase().trim();
        if (!key) continue;
        if (bucket.people.some(x => (x.email || '').toLowerCase().trim() === key)) continue;
        bucket.people.push(p);
      }
    }
    return out.filter(it => it.people.length > 0);
  }, [events]);

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

  // v29.31: Gegenprobe gegen den AKTUELLEN Anmeldestand. Der Scan ist 24 h
  // gecacht — wer inzwischen abgemeldet wurde, stand hier trotzdem weiter als
  // offener Fall („steht immer noch, obwohl ich die Person gerade abgemeldet
  // habe"). Das Verwerfen des Caches nach einer Abmeldung greift nur im
  // eigenen Browser; deshalb zusätzlich lesen: eine Abfrage je Event, das
  // überhaupt einen Fund hat — kein erneuter Konten-Check über Graph.
  const ACTIVE_REG = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
  const filterStillRegistered = React.useCallback(async (items: InactiveItem[]): Promise<InactiveItem[]> => {
    const out: InactiveItem[] = [];
    for (const it of items) {
      try {
        const regs = await getAllRegistrations(it.eventId);
        const activeEmails = new Set(regs
          .filter(r => ACTIVE_REG.indexOf(r.Status) >= 0)
          .map(r => (r.ParticipantEmail || '').trim().toLowerCase()));
        const remaining = it.people.filter(p => activeEmails.has((p.email || '').toLowerCase().trim()));
        if (remaining.length > 0) out.push({ ...it, people: remaining });
      } catch {
        // Lesefehler: den gecachten Stand lieber zeigen als still schlucken.
        out.push(it);
      }
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
    // v29.31: Schlüssel zentral (utils/accountCheckCache) — nach einer
    // Abmeldung im Organizer Center wird genau dieser Eintrag verworfen.
    const CACHE = INACTIVE_SUMMARY_CACHE_KEY;
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
          const cached = collapseToUmbrella(parsed.items.filter(it => liveIds.has(it.eventId) && !isAutoDereg(it.eventId)));
          // v24.59: bereits benachrichtigte Konten gleich ausblenden.
          // v29.31: …und inzwischen abgemeldete Personen (Gegenprobe live).
          filterNotified(cached)
            .then(f => filterStillRegistered(f))
            .then(f => { if (!cancelled) setInactiveSummary(f); })
            .catch(() => { if (!cancelled) setInactiveSummary(cached); });
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
          const notifyItems = collapseToUmbrella(items.filter(it => !isAutoDereg(it.eventId)));
          const filtered = await filterNotified(notifyItems);
          if (!cancelled) setInactiveSummary(filtered);
          // Der frische Scan liest die Anmeldungen ohnehin selbst — hier ist
          // keine zweite Gegenprobe nötig (die greift nur beim Cache-Pfad).
        })
        .catch(() => { /* best-effort */ });
    }, 3500); // dem Boot Vorrang geben
    return () => { cancelled = true; window.clearTimeout(t); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEventsLoading, isAdmin, currentUser?.email]);

  const [checkInBoxes, setCheckInBoxes] = React.useState<Array<{
    eventId: string; title: string; qrSmall: string; qrData: string;
    name: string; tid?: number;
    /** v31.4: Person ist schon eingecheckt — die Box bleibt trotzdem stehen. */
    checkedIn?: boolean;
  }>>([]);
  const [qrBigModal, setQrBigModal] = React.useState<{ dataUrl: string; title: string; name: string; tid?: number } | null>(null);
  /**
   * v31.4 (Review): Ein Lesefehler ist keine Null. `getMyRegistration` liefert
   * bei 403/429/500 dasselbe `null` wie bei „keine Zeile" — die Box wäre dann
   * wortlos weg, und der neue Kommentar unten erhebt sie ausdrücklich zur
   * einzigen Fundstelle für den eigenen Code. Wer nichts sieht, schließt
   * daraus „ich bin nicht angemeldet" und stellt sich am falschen Tisch an.
   * Deshalb: unbekannt wird BENANNT, statt als „nichts" gerendert.
   */
  const [checkInBoxError, setCheckInBoxError] = React.useState(false);
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
        // v31.4 (Review): Die Klammer eines `subEventsOnlyMode`-Events ist
        // keine Anmeldeeinheit — ihre Zeile ist die Schattenzeile und
        // konstruktionsbedingt IMMER 'Angemeldet' (kein `reserveSeat`, der
        // Wartelisten-Riegel nimmt den Modus aus, und seit v30.68 wird sie
        // VOR dem Termin geschrieben). Ein QR-Kasten auf der Klammer hieße
        // deshalb „Einlass zugesagt" auch für Warteliste und für „Klammer
        // ohne Termin". Der Termin selbst hat eine eigene EventNumber und
        // damit einen eigenen Kasten — dieselbe Regel, nach der „Meine
        // Events" den QR-Knopf auf der Klammer seit v28.7 auslässt.
        if (e.subEventsOnlyMode) return false;
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
      if (candidates.length === 0) { if (!cancelled) { setCheckInBoxes([]); setCheckInBoxError(false); } return; }
      const boxes: Array<{ eventId: string; title: string; qrSmall: string; qrData: string; name: string; tid?: number; checkedIn?: boolean }> = [];
      let readFailed = false;
      const QRCode = await import('qrcode');
      for (const ev of candidates) {
        try {
          // Nur die Status, die „versuch es gleich nochmal" heißen, zählen als
          // Lesefehler: Drosselung (429), Server-Fehler (5xx) und 0
          // (Netz/Abbruch). 401/403/404 sind hier KEINE Störung — die
          // Kandidatenliste enthält auch Events, für die die Person keine
          // Anmeldung und keinen Zugriff hat; daraus einen Warnkasten zu
          // bauen hieße, ihn allen zu zeigen.
          const reg = await getMyRegistration(ev.id, (status) => {
            if (status === 0 || status === 429 || status >= 500) readFailed = true;
          });
          // v31.4: Die Box hing an EINEM Status ('QR versendet') — und
          // verschwand damit in drei Lagen, in denen der Code weiter gilt:
          // nach dem eigenen Check-in ('Eingecheckt'), nach einem Auschecken
          // im Organizer Center (setzt hart 'Angemeldet' zurück) und bei
          // jeder Anmeldung vor dem QR-Massenversand. Der Code selbst hängt
          // an gar keinem Status: Er ist `DEX|<EventNr>|<E-Mail>` und wird
          // hier im Browser erzeugt; der Check-in-Tisch nimmt ihn immer an.
          // Das Fenster ist ohnehin schon eng (zwei Tage vor dem Event), und
          // ein aktiv Angemeldeter, der seinen Code sucht, findet ihn sonst
          // nirgends. Warteliste, Abmeldung und No-Show bleiben draußen —
          // dort wäre der Code eine falsche Zusage.
          const st = reg ? reg.Status : '';
          if (!reg || (st !== 'QR versendet' && st !== 'Angemeldet' && st !== 'Eingecheckt')) continue;
          const qrData = `DEX|${ev.eventNumber}|${reg.ParticipantEmail || myEmail}`;
          const qrSmall = await QRCode.toDataURL(qrData, { width: 132, margin: 1 });
          const name = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || (reg.ParticipantEmail || myEmail);
          // v31.4 (Review): Wenn eine QR-Mail verschickt wurde, gilt DEREN
          // Nummer — nicht die laufende von heute. Der Check-in löst die
          // vorgelesene Zahl seit v31.4 zuerst über `QrSentId` auf; stünde
          // hier die laufende Nummer, hätte dieselbe Person zwei
          // verschiedene Zahlen (App und Mail), und eine davon zeigt am
          // Tisch auf jemand anderen. Wortgleich zu MyEventsPage.
          boxes.push({ eventId: ev.id, title: ev.title || '', qrSmall, qrData, name, tid: reg.QrSentId || reg.TeilnehmerID || undefined, checkedIn: st === 'Eingecheckt' });
        } catch { readFailed = true; /* einzelner Lookup-Fehler — Box entfällt, aber nicht wortlos */ }
        if (cancelled) return;
      }
      if (!cancelled) { setCheckInBoxes(boxes); setCheckInBoxError(readFailed && boxes.length === 0); }
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
  // v28.14: Der Einführungs-Hinweis darf erst rendern, wenn die eigenen
  // Anmeldungen geladen sind — sonst blitzt er kurz auf und verschwindet
  // wieder, sobald die „Du bist angemeldet"-Box das Event übernimmt.
  const [myRegBoxesLoaded, setMyRegBoxesLoaded] = React.useState(false);
  const [nowTick, setNowTick] = React.useState(Date.now());
  React.useEffect(() => {
    const id = window.setInterval(() => setNowTick(Date.now()), 60000);
    return () => window.clearInterval(id);
  }, []);
  // v30.68: SOFORT zeigen, was beim letzten Mal galt — Merker je Person im
  // Browser —, dann im Hintergrund frisch prüfen. Bis hierher erschienen die
  // Kacheln erst, wenn alle Events geladen UND bis zu zwanzig Teilnehmer-
  // listen nacheinander gelesen waren: mehrere Sekunden, in denen die
  // meisten längst auf „Start" geklickt hatten (Nutzer-Befund 02.09.2026).
  // Der Merker ist Komfort, keine Wahrheit: vergangene Events fallen beim
  // Lesen raus, die frische Prüfung ersetzt ihn, sobald sie da ist.
  const regBoxCacheKey = (email: string): string => `dex_landing_regboxes_v1:${email.toLowerCase()}`;
  React.useEffect(() => {
    const myEmail = (currentUser?.email || '').trim();
    if (!myEmail) return;
    try {
      const raw = window.localStorage.getItem(regBoxCacheKey(myEmail));
      if (!raw) return;
      const parsed = JSON.parse(raw) as { ts?: number; items?: Array<{ eventId: string; title: string; imageUrl?: string; startDate: string; location?: string }> };
      if (!parsed || !Array.isArray(parsed.items)) return;
      if (Date.now() - (parsed.ts || 0) > 14 * 24 * 60 * 60 * 1000) return;
      const now = Date.now();
      const stillAhead = parsed.items.filter(b => {
        const d = new Date(b.startDate);
        if (!Number.isFinite(d.getTime())) return false;
        return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59).getTime() >= now;
      });
      if (stillAhead.length > 0) setMyRegBoxes(prev => (prev.length > 0 ? prev : stillAhead));
    } catch { /* Merker nicht lesbar — dann eben erst die frische Prüfung */ }
  }, [currentUser?.email]);
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
      // v30.68: Parallel mit kleiner Obergrenze statt streng nacheinander —
      // zwanzig Kandidaten hintereinander waren der zweite Grund für die
      // Wartezeit. Nicht Promise.all über alles: Das wäre die Drosselung
      // (s. mapLimited im EventContext, v29.47).
      const limited = async <T,>(items: T[], n: number, fn: (x: T) => Promise<void>): Promise<void> => {
        let next = 0;
        const worker = async (): Promise<void> => {
          for (;;) {
            const k = next++;
            if (k >= items.length || cancelled) return;
            await fn(items[k]);
          }
        };
        await Promise.all(Array.from({ length: Math.max(1, Math.min(n, items.length)) }, () => worker()));
      };
      const registeredIds = new Set<string>();
      await limited(topCandidates, 4, async (top) => {
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
          await limited(children, 2, async (child) => {
            if (registered) return;
            try {
              if (isActiveReg(await getMyRegistration(child.id))) registered = true;
            } catch { /* einzelner Sub-Event-Lookup-Fehler */ }
          });
        }
        if (registered) registeredIds.add(top.id);
      });
      if (cancelled) return;
      const boxes = topCandidates
        .filter(top => registeredIds.has(top.id))
        .slice(0, 6)
        .map(top => ({ eventId: top.id, title: top.title || '', imageUrl: top.imageUrl, startDate: top.startDate, location: top.location }));
      if (!cancelled) {
        setMyRegBoxes(boxes);
        setMyRegBoxesLoaded(true);
        try { window.localStorage.setItem(regBoxCacheKey(myEmail), JSON.stringify({ ts: Date.now(), items: boxes })); } catch { /* Merker ist Komfort */ }
      }
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
      {/* v31.9: Die Sprachwahl sitzt jetzt am Seitenrand statt absolut in der
          Karte. Sie ist — wie die Versionsmarke gegenüber — ein Randelement
          ohne Platz in der Lesereihenfolge; in der Karte lag sie über der
          ersten Zeile und hätte den nach oben gezogenen Check-in-Kasten
          verdeckt. */}
      <div style={{ position: 'absolute', top: 12, left: 16, display: 'flex', gap: 4, zIndex: 7 }}>
        <button
          type="button"
          className={cx('dex-ui-chip', locale === 'de' && 'is-active')}
          onClick={() => setLocale('de')}
          title="Deutsch"
        >
          DE
        </button>
        <button
          type="button"
          className={cx('dex-ui-chip', locale === 'en' && 'is-active')}
          onClick={() => setLocale('en')}
          title="English"
        >
          EN
        </button>
      </div>

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
          {/* v31.9.1: Orb und Begrüßung stehen wieder ganz oben.
              v31.9 hatte den Check-in-Kasten und &bdquo;Du bist angemeldet&ldquo;
              davor gezogen, damit der QR-Code am Eventmorgen ohne Scrollen da
              ist. Der Preis war das Gesicht der App — Nutzer-Entscheidung
              09.09.2026: &bdquo;zurück wie vorher&ldquo;. Die Kästen folgen
              direkt unter der Begrüßung und stehen damit weiterhin VOR dem
              Start-Knopf; das war der eigentliche Gewinn und bleibt. */}
          {/* v28.33: animiertes DEX-Logo (Canvas) statt des rotierenden
              Farbring-Platzhalters. Zeichnet die Höhenlinien-Kugel des
              DEX-Logos live, pausiert automatisch bei „Bewegung reduzieren",
              ausserhalb des Viewports und im inaktiven Tab. */}
          <div className="landing__orb">
            <DexLogo title="DEX" motion="oscillate" style={{ width: '100%' }} />
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
          {/* v31.9: Was heute zu tun ist, steht vor dem Start-Knopf. Bis v31.8
              lagen der Check-in-Kasten mit QR und Einlassnummer und
              &bdquo;Du bist angemeldet&ldquo; ganz unten, hinter dem
              Werbekasten — am Eventmorgen musste man daran vorbeiscrollen.
              (v31.9 hatte sie zusätzlich vor Orb und Begrüßung gezogen; das ist
              seit v31.9.1 wieder zurückgenommen, siehe Kommentar oben.) */}
          {/* v22.1: Check-in-Hinweisbox(en) — ab 2 Tage vor dem Event, sobald
              der eigene QR-Code versendet wurde. Klick auf den kleinen QR
              öffnet ihn groß im Modal (zum Vorzeigen am Eingang).
              v28.13: ALLE Hinweis-Boxen (Check-in + „Du bist angemeldet") in
              EINEM Wrapper — vorher schob das Karten-gap (bis 40px) + 12px
              marginBottom jede Box weit auseinander. Jetzt sitzt die Gruppe
              kompakt (10px innen), der große Karten-Abstand gilt nur noch
              einmal um die ganze Gruppe. */}
          {(() => {
            const checkInIds = new Set(checkInBoxes.map(b => b.eventId));
            const regList = myRegBoxes.filter(b => !checkInIds.has(b.eventId));
            if (checkInBoxes.length === 0 && regList.length === 0 && !checkInBoxError) return null;
            return (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 10, width: '100%' }}>
                {/* v31.4 (Review): Der Kasten fehlte bei einem Lesefehler
                    wortlos — und wortlos liest sich wie „du bist nicht
                    angemeldet". Am Eventmorgen ist eine Drosselung (429) auf
                    der Teilnehmerliste der Normalfall, und dann stellt sich
                    die Person am Organizer-Tisch an statt am Check-in. */}
                {checkInBoxError && (
                  <div className="dex-ui-callout dex-ui-callout--warn">
                    {isDe
                      ? 'Dein Check-in-Code konnte gerade nicht geladen werden — bitte lade die Seite neu. Deinen Code findest du auch in „Meine Events" und in deiner QR-Mail.'
                      : 'Your check-in code could not be loaded right now — please reload the page. You will also find your code in “My events” and in your QR email.'}
                  </div>
                )}
                {checkInBoxes.map(box => (
                  <button
                    key={box.eventId}
                    type="button"
                    onClick={() => { openBigQr(box).catch(() => { /* */ }); }}
                    // v31.9: Weiße Karte mit grüner Kante statt grüner Fläche
                    // (Grundsatz 1.1); der Hover steckt jetzt in der Klasse und
                    // nicht mehr in einem Inline-Style, der keinen kennt.
                    className="dex-ui-card dex-ui-card--accent dex-ui-card--hover"
                    style={{
                      display: 'flex', alignItems: 'center', gap: 14, width: '100%',
                      textAlign: 'left', cursor: 'pointer', fontFamily: 'inherit',
                    }}
                    title={isDe ? 'QR-Code groß anzeigen' : 'Show QR code enlarged'}
                  >
                    <img src={box.qrSmall} alt="QR" style={{ width: 66, height: 66, flexShrink: 0, borderRadius: 6, background: '#fff', border: '1px solid var(--dex-gray-200)' }} />
                    <span style={{ minWidth: 0 }}>
                      <span style={{ display: 'block', fontWeight: 700, fontSize: '0.92rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>
                        {box.checkedIn
                          ? (isDe ? 'Eingecheckt für' : 'Checked in for')
                          : (isDe ? 'Check-in für' : 'Check-in for')} {box.title}
                      </span>
                      <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                        {box.checkedIn
                          ? (isDe
                            ? 'Du bist schon eingecheckt. Der Code bleibt hier, falls du ihn noch einmal brauchst.'
                            : 'You are already checked in. The code stays here in case you need it again.')
                          : (isDe
                            ? 'Dein persönlicher QR-Code — antippen zum Vergrößern und am Eingang vorzeigen.'
                            : 'Your personal QR code — tap to enlarge and show at the entrance.')}
                      </span>
                      {/* v31.4: Die Nummer gehört sichtbar in die Box, nicht erst
                          ins große Modal. Auf Android scheitert der Kamera-Scan
                          in der SharePoint-App regelmäßig (s. CLAUDE.md
                          „Kamera-Scan"); dann ist genau diese Zahl der Weg, und
                          man will sie vorlesen können, ohne erst zu tippen. */}
                      {box.tid !== undefined && (
                        <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                          {isDe ? 'Deine Nummer am Einlass: ' : 'Your number at the entrance: '}
                          <strong style={{ fontFamily: "'Courier New', Courier, monospace", fontSize: '0.9rem', letterSpacing: '0.04em', color: 'var(--dex-gray-800)' }}>
                            {`00${box.tid}`.slice(-Math.max(3, String(box.tid).length))}
                          </strong>
                        </span>
                      )}
                    </span>
                  </button>
                ))}
                {/* v24.13: „Du bist angemeldet"-Boxen — aktive Events, für die man
                    angemeldet ist; ausgeblendet, wenn schon eine Check-in/QR-Box läuft. */}
                {regList.map(box => {
                  const startTs = new Date(box.startDate).getTime();
                  const diff = startTs - nowTick;
                  // v29.18: Kein Rot mehr, wenn das Event kurz bevorsteht. Rot
                  // heißt in der App „Problem/Fehler" — ein Event, auf das man
                  // sich angemeldet hat und das morgen stattfindet, ist aber
                  // eine Vorfreude-Information, keine Warnung. Der Countdown
                  // bleibt grün, egal wie nah der Termin ist.
                  let countdown: string;
                  if (diff <= 0) { countdown = isDe ? 'läuft gerade' : 'happening now'; }
                  else if (diff < 24 * 60 * 60 * 1000) { const h = Math.max(1, Math.ceil(diff / (60 * 60 * 1000))); countdown = isDe ? `noch ${h} ${h === 1 ? 'Stunde' : 'Stunden'}` : `in ${h} ${h === 1 ? 'hour' : 'hours'}`; }
                  else { const d = Math.floor(diff / (24 * 60 * 60 * 1000)); countdown = isDe ? `noch ${d} ${d === 1 ? 'Tag' : 'Tage'}` : `in ${d} ${d === 1 ? 'day' : 'days'}`; }
                  const dateLabel = new Date(box.startDate).toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
                  return (
                    <button
                      key={box.eventId}
                      type="button"
                      onClick={() => navigate('my-events')}
                      // v31.9: dieselbe Kartenklasse wie der Check-in-Kasten —
                      // der Hover sagt „hier kannst du klicken", der Inline-
                      // Style konnte das nicht.
                      className="dex-ui-card dex-ui-card--hover"
                      style={{ display: 'flex', alignItems: 'center', gap: 14, width: '100%', textAlign: 'left', cursor: 'pointer', fontFamily: 'inherit' }}
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
                      {/* v31.9: Anzeige, kein Schalter — also dex-ui-pill
                          (Grundsatz 3) statt handgebauter Pille. */}
                      <span className="dex-ui-pill dex-ui-pill--green" style={{ flexShrink: 0 }}>
                        {countdown}
                      </span>
                    </button>
                  );
                })}
              </div>
            );
          })()}
          {/* v31.9: „Start" ist die Handlung dieser Seite und damit der einzige
              Primär-Knopf (Grundsatz 1.5). Bis v31.8 war er `btn-outline`,
              während der Werbekasten darunter vollflächig grün war — die
              Rangfolge stand auf dem Kopf. */}
          <button className="btn btn-lg btn-block btn-primary" data-tour="landing-start" onClick={() => navigate('start')} style={{ maxWidth: 360 }}>
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
            {/* v31.9: Aus der vollflächig grünen Kachel wird eine weiße Karte
                mit Hover. Grün ist in dieser App die Farbe der Handlung — hier
                lag sie auf der Werbung und nicht auf „Start" darüber
                (Grundsatz 1.1 und 1.5). Der Hover steckt jetzt in
                `dex-ui-card--hover`; der `ctaHover`-State war reine Optik und
                entfällt (Grundsatz 3). */}
            <button
              type="button"
              onClick={() => setShowInquiry(true)}
              className="landing__bubble dex-ui-card dex-ui-card--hover"
              style={{
                display: 'flex', alignItems: 'center', gap: 12,
                cursor: 'pointer', textAlign: 'left',
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
                <span style={{ display: 'block', fontWeight: 800, fontSize: '0.98rem', lineHeight: 1.25, color: 'var(--dex-gray-800)' }}>
                  {locale === 'de' ? 'DEX für dein Event nutzen' : 'Use DEX for your event'}
                </span>
                <span style={{ display: 'block', fontSize: '0.82rem', lineHeight: 1.3, marginTop: 2, color: 'var(--dex-gray-600)' }}>
                  {locale === 'de' ? 'Werde Organizer und nutze alle Funktionen.' : 'Become an organizer and use all features.'}
                </span>
                {/* v28.15: Einführungs-Hinweis als weiße Pille IN der Box
                    (statt separater Kachel, v28.13/14) — Klick führt direkt
                    zur Anmeldeseite des Einführungs-Events. Kein <button>
                    (verschachtelte Buttons sind invalide) — span[role=button]
                    mit stopPropagation, damit nicht die Anfrage aufgeht. */}
                {(() => {
                  if (!myRegBoxesLoaded) return null;
                  const isIntroTitle = (t: string): boolean => /dex/i.test(t) && /einf(ü|u)hrung|introduction|onboarding/i.test(t);
                  const intro = (events || [])
                    .filter(e => e.status === 'Active' && !e.parentEventId && !e.isFictive && !!e.startDate
                      && new Date(e.startDate).getTime() > nowTick && isIntroTitle(e.title || ''))
                    .sort((a, b) => new Date(a.startDate).getTime() - new Date(b.startDate).getTime())[0];
                  if (!intro) return null;
                  if (myRegBoxes.some(b => b.eventId === intro.id) || checkInBoxes.some(b => b.eventId === intro.id)) return null;
                  const dateLabel = new Date(intro.startDate).toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
                  const goRegister = (): void => navigate('registration', intro.id);
                  return (
                    <span
                      role="button"
                      tabIndex={0}
                      onClick={e => { e.stopPropagation(); goRegister(); }}
                      onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.stopPropagation(); e.preventDefault(); goRegister(); } }}
                      title={intro.title}
                      // v31.9: Die Pille war weiß auf grüner Fläche — auf der
                      // weißen Karte wäre sie unsichtbar. Sie ist schaltbar
                      // (führt zur Anmeldung), also ein Chip, keine Pille.
                      className="dex-ui-chip"
                      style={{ marginTop: 8 }}
                    >
                      <GraduationCap size={13} strokeWidth={2.5} />
                      {isDe ? `Für virtuelles Training anmelden · ${dateLabel}` : `Register for the virtual training · ${dateLabel}`}
                    </span>
                  );
                })()}
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

      {/* v22 / v22.45: Hinweis-Boxen oben rechts auf der Landing Page —
          gestapelt in einem gemeinsamen Container (Archivierung für Admin,
          Inaktive-Konten-Warnung für Organizer/Admin).
          v31.9: Zwei Änderungen. (1) Der Block steht jetzt NACH der Karte im
          DOM. Am Desktop ändert das nichts (er hängt absolut oben rechts an
          `.landing`), auf dem Handy ist er statisch und schob bis zu sechs
          Kästen vor die eigene Anmeldung — Organizer-Arbeit vor dem QR-Code
          am Eventmorgen. (2) Aus sechs gestapelten Karten mit vier
          destruktiven Knöpfen wird EIN Warnkasten mit Aufklapper; zugeklappt
          nennt er, was ansteht, offen steht alles Bisherige da. */}
      {((isAdmin && archInfo && archInfo.total > 0) || (isAdmin && delArchCount > 0) || inactiveSummary.length > 0 || staleDrafts.length > 0 || (isAdmin && (pdWarn > 0 || pdDue > 0))) && (() => {
        // Kurzfassung für den zugeklappten Zustand: Was steckt drin? Ohne sie
        // wäre der Aufklapper eine Tür ohne Schild.
        const hints: string[] = [];
        const inactivePeople = inactiveSummary.reduce((acc, it) => acc + it.people.length, 0);
        if (inactivePeople > 0) hints.push(isDe
          ? `${inactivePeople} ${inactivePeople === 1 ? 'inaktives Konto' : 'inaktive Konten'}`
          : `${inactivePeople} inactive account${inactivePeople === 1 ? '' : 's'}`);
        if (staleDrafts.length > 0) hints.push(isDe
          ? `${staleDrafts.length} ${staleDrafts.length === 1 ? 'abgelaufener Entwurf' : 'abgelaufene Entwürfe'}`
          : `${staleDrafts.length} expired draft${staleDrafts.length === 1 ? '' : 's'}`);
        if (isAdmin && archInfo && archInfo.total > 0) hints.push(isDe
          ? `${archInfo.total} Zeilen zum Archivieren`
          : `${archInfo.total} rows to archive`);
        if (isAdmin && delArchCount > 0) hints.push(isDe
          ? `${delArchCount} alte Archiv-Einträge`
          : `${delArchCount} old archive entries`);
        if (isAdmin && pdWarn > 0) hints.push(isDe
          ? `${pdWarn} Teilnehmerlisten laufen ab`
          : `${pdWarn} attendee lists expiring`);
        if (isAdmin && pdDue > 0) hints.push(isDe
          ? `${pdDue} Teilnehmerlisten zum Löschen`
          : `${pdDue} attendee lists to delete`);
        return (
      <div style={{
        ...(isMobile
          ? { position: 'static', width: '100%', margin: '16px auto 0' }
          : { position: 'absolute', top: 34, right: 16, width: 300 }),
        maxWidth: 'calc(100vw - 32px)', zIndex: 6,
      }}>
      <div className="dex-ui-callout dex-ui-callout--warn" style={{ boxShadow: '0 8px 24px rgba(0,0,0,0.10)' }}>
        <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
        <div className="dex-ui-callout-body">
          <button
            type="button"
            className={cx('dex-ui-disclosure', adminHintsOpen && 'is-open')}
            onClick={() => setAdminHintsOpen(o => !o)}
            style={{ color: 'inherit' }}
          >
            <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
            {isDe ? 'Aufgaben aus der Verwaltung' : 'Administration tasks'}
            <span className="dex-ui-disclosure-count">{hints.length}</span>
          </button>
          {!adminHintsOpen && (
            <div style={{ marginTop: 2 }}>{hints.join(' · ')}</div>
          )}
          {adminHintsOpen && (
          <div className="dex-ui-disclosure-body" style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
      {/* v22.45: Teilnehmer ohne aktives Deloitte-Konto. */}
      {inactiveSummary.length > 0 && (() => {
        const totalPeople = inactiveSummary.reduce((acc, it) => acc + it.people.length, 0);
        // v31.9: Karte im Warnkasten (dex-ui-card--sm) statt eigener
        // Kasten-Optik — Rahmen und Schatten trägt der Kasten drumherum.
        return (
          <div className="dex-ui-card dex-ui-card--sm" style={{ width: '100%', textAlign: 'left' }}>
            <div style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-gray-800)', marginBottom: 6 }}>
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
                  {/* v31.9: klickbare Zeile mit Hover statt Inline-Kasten —
                      der Klick öffnet das Event, das sagt jetzt auch die
                      Optik (Grundsatz 3). */}
                  <button
                    type="button"
                    onClick={() => navigate('admin', it.eventId)}
                    title={isDe ? 'Event im Organizer Center öffnen' : 'Open event in the Organizer Center'}
                    className="dex-ui-rowbtn dex-ui-row dex-ui-row--framed"
                    style={{ padding: '6px 10px' }}
                  >
                    <span className="dex-ui-row-main">
                      <span className="dex-ui-row-title dex-ui-row-title--wrap" style={{ fontSize: '0.78rem' }}>
                        {it.title} <span style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>({it.people.length})</span>
                      </span>
                      <span className="dex-ui-row-sub" style={{ fontSize: '0.72rem' }}>
                        {it.people.map(p => p.name).join(', ')}
                      </span>
                    </span>
                  </button>
                  {/* v24.51: Organizer per Mail benachrichtigen (Dedup: nur 1x je
                      Event+Person, egal welcher Admin klickt). */}
                  {notifyResult[it.eventId] ? (
                    <div className="dex-ui-callout dex-ui-callout--sm dex-ui-callout--neutral" style={{ marginTop: 4 }}>
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
                      // v31.9: normaler Nebenknopf statt oranger Fläche — der
                      // Primär-Knopf dieser Seite ist „Start" (Grundsatz 1.5).
                      className="btn btn-secondary dex-ui-btn-sm"
                      style={{ marginTop: 5, width: '100%' }}
                    >
                      {/* v31.9: „…" sagt nicht, was gerade passiert (6d). */}
                      {notifyBusyId === it.eventId
                        ? (isDe ? 'Wird gesendet…' : 'Sending…')
                        : (isDe ? 'Organizer benachrichtigen' : 'Notify organizer')}
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
        <div className="dex-ui-card dex-ui-card--sm" style={{ width: '100%', textAlign: 'left' }}>
          <div style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-gray-800)', marginBottom: 6 }}>
            {isDe ? 'Entwürfe aufräumen' : 'Clean up drafts'}
          </div>
          <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <><strong>{staleDrafts.length}</strong> {staleDrafts.length === 1 ? 'Entwurf ist' : 'Entwürfe sind'} abgelaufen (Datum vorbei) und {staleDrafts.length === 1 ? 'wurde' : 'wurden'} nie aktiviert. Du kannst {staleDrafts.length === 1 ? 'ihn' : 'sie'} hier löschen.</>
              : <><strong>{staleDrafts.length}</strong> draft(s) expired and were never activated. You can delete them here.</>}
          </p>
          {/* v31.9: Zeile ohne Hover — sie ist Anzeige, nur der Knopf rechts
              handelt (Löschen ist die Ausnahme, die rechts außen stehen darf,
              2a′). */}
          <ul style={{ margin: 0, padding: 0, listStyle: 'none', display: 'flex', flexDirection: 'column', gap: 6 }}>
            {staleDrafts.map(ev => (
              <li key={ev.id} className="dex-ui-row dex-ui-row--static dex-ui-row--framed" style={{ padding: '6px 8px 6px 10px', gap: 8 }}>
                <span className="dex-ui-row-main dex-ui-row-title" style={{ fontSize: '0.78rem' }}>
                  {ev.title || (isDe ? 'Ohne Titel' : 'Untitled')}
                </span>
                {/* v31.9: `btn-danger` statt roter Schrift auf `btn-secondary`
                    — in dieser App ist der Gefahren-Knopf bewusst grau, die
                    Warnung trägt die Rückfrage (confirmDialog, danger). */}
                <button
                  type="button"
                  className="btn btn-danger dex-ui-btn-sm"
                  style={{ flexShrink: 0 }}
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
        <div className="dex-ui-card dex-ui-card--sm" style={{ width: '100%', textAlign: 'left' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Archivierung' : 'Archiving'}
            </span>
            <span className="dex-ui-pill dex-ui-pill--orange dex-ui-pill--sm">{isDe ? 'Nur Admin' : 'Admin only'}</span>
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
          {/* v31.9: kein zweiter Primär-Knopf auf der Seite — „Start" ist der
              eine (Grundsatz 1.5). */}
          <button
            type="button"
            className="btn btn-secondary dex-ui-btn-sm"
            style={{ width: '100%' }}
            onClick={() => { startArchive().catch(() => { /* */ }); }}
          >
            {isDe ? 'Jetzt archivieren' : 'Archive now'}
          </button>
        </div>
      )}
      {/* v23.40: Löschkonzept — alte Archiv-Einträge (älter als 1 Monat). */}
      {isAdmin && delArchCount > 0 && (
        <div className="dex-ui-card dex-ui-card--sm" style={{ width: '100%', textAlign: 'left' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Archiv aufräumen' : 'Clean up archive'}
            </span>
            <span className="dex-ui-pill dex-ui-pill--orange dex-ui-pill--sm">{isDe ? 'Nur Admin' : 'Admin only'}</span>
          </div>
          <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <><strong>{delArchCount}</strong> Archiv-Einträge sind älter als 1 Monat und können endgültig gelöscht werden.</>
              : <><strong>{delArchCount}</strong> archive entries are older than 1 month and can be permanently deleted.</>}
          </p>
          {/* v31.9: `btn-danger` statt roter Schrift auf `btn-secondary` — der
              Gefahren-Knopf ist in dieser App bewusst grau, gewarnt wird im
              Rückfrage-Dialog. */}
          <button
            type="button"
            className="btn btn-danger dex-ui-btn-sm"
            style={{ width: '100%' }}
            disabled={delArchBusy}
            onClick={() => { startDeleteOldArchive().catch(() => { /* */ }); }}
          >
            {delArchBusy ? (isDe ? 'Wird gelöscht…' : 'Deleting…') : (isDe ? 'Alte Einträge löschen' : 'Delete old entries')}
          </button>
        </div>
      )}
      {/* v26.32: Löschkonzept — Vorwarnung (Teilnehmerliste wird in ~1 Woche gelöscht). */}
      {isAdmin && pdWarn > 0 && (
        <div className="dex-ui-card dex-ui-card--sm" style={{ width: '100%', textAlign: 'left' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Teilnehmerlisten laufen ab' : 'Attendee lists expiring'}
            </span>
            <span className="dex-ui-pill dex-ui-pill--orange dex-ui-pill--sm">{isDe ? 'Nur Admin' : 'Admin only'}</span>
          </div>
          <p style={{ margin: '0 0 6px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <>Bei {pdWarn === 1 ? 'diesem Event' : <>diesen <strong>{pdWarn}</strong> Events</>} wird die Teilnehmerliste in etwa einer Woche gelöscht (3 Monate nach dem Event). Die Organizer wurden automatisch informiert, die Liste noch herunterzuladen. Event &amp; Kennzahlen bleiben im Statistik-Archiv erhalten:</>
              : <>For {pdWarn === 1 ? 'this event' : <>these <strong>{pdWarn}</strong> events</>} the attendee list will be deleted in about a week (3 months after the event). The organizers were notified automatically to download it. Event &amp; KPIs are kept in the statistics archive:</>}
          </p>
          <ul style={{ margin: 0, paddingLeft: 16, fontSize: '0.76rem', color: 'var(--dex-gray-600)', lineHeight: 1.6 }}>
            {pdWarnEvents.map(ev => (
              <li key={ev.id}>
                <strong>{ev.title}</strong>
                {(ev.endDate || ev.startDate)
                  ? ` — ${isDe ? 'Event vom' : 'event on'} ${new Date(ev.endDate || ev.startDate || '').toLocaleDateString(isDe ? 'de-DE' : 'en-GB')}`
                  : ''}
              </li>
            ))}
          </ul>
        </div>
      )}
      {/* v26.32: Löschkonzept — fällige Teilnehmerlisten löschen (Event bleibt, KPIs ins Archiv). */}
      {isAdmin && pdDue > 0 && (
        <div className="dex-ui-card dex-ui-card--sm" style={{ width: '100%', textAlign: 'left' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
            <span style={{ fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Teilnehmerlisten löschen' : 'Delete attendee lists'}
            </span>
            <span className="dex-ui-pill dex-ui-pill--orange dex-ui-pill--sm">{isDe ? 'Nur Admin' : 'Admin only'}</span>
          </div>
          <p style={{ margin: '0 0 6px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <>Bei {pdDue === 1 ? 'diesem Event ist' : <>diesen <strong>{pdDue}</strong> Events ist</>} die Aufbewahrungsfrist (3 Monate) abgelaufen. Beim Löschen werden die wichtigsten Kennzahlen ins Statistik-Archiv übernommen und dann die Teilnehmerliste entfernt — das Event bleibt erhalten:</>
              : <>{pdDue === 1 ? 'This event has' : <>These <strong>{pdDue}</strong> events have</>} passed the 3-month retention. Deleting archives the key KPIs to the statistics archive and then removes the attendee list — the event itself is kept:</>}
          </p>
          <ul style={{ margin: '0 0 10px', paddingLeft: 16, fontSize: '0.76rem', color: 'var(--dex-gray-600)', lineHeight: 1.6 }}>
            {pdDueEvents.map(ev => (
              <li key={ev.id}>
                <strong>{ev.title}</strong>
                {(ev.endDate || ev.startDate)
                  ? ` — ${isDe ? 'Event vom' : 'event on'} ${new Date(ev.endDate || ev.startDate || '').toLocaleDateString(isDe ? 'de-DE' : 'en-GB')}`
                  : ''}
              </li>
            ))}
          </ul>
          {/* v31.9: siehe oben — `btn-danger`, kein roter Text auf grauem Knopf. */}
          <button
            type="button"
            className="btn btn-danger dex-ui-btn-sm"
            style={{ width: '100%' }}
            disabled={pdBusy}
            onClick={() => { startParticipantDeletion().catch(() => { /* */ }); }}
          >
            {pdBusy ? (isDe ? 'Wird gelöscht…' : 'Deleting…') : (isDe ? 'Teilnehmerlisten löschen' : 'Delete attendee lists')}
          </button>
        </div>
      )}
          </div>
          )}
        </div>
      </div>
      </div>
        );
      })()}
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

// v11.47: KPI-Box-Reihe über dem "Entwickelt von ..."-Block. Drei Boxen
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
 *  `value`. Wenn `value` sich ändert, startet eine neue Animation; bei
 *  schnellen Änderungen wird die laufende Animation sauber durch eine neue
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
