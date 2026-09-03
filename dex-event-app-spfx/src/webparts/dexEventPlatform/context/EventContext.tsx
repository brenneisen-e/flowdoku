/**
 * Event Context - zentraler State für alle Events
 *
 * Lädt Events aus der SharePoint-Liste DEX_Events.
 * Erstellt die Liste automatisch beim ersten Start.
 * Verwaltet Registrierungen über Event-Subsites mit Teilnehmerlisten.
 *
 * - Eike, Maerz 2026
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DeloitteEvent } from '../types';
import { EventService, SPEvent, SPRegistration, ReseedSummary } from '../services/EventService';
import { verifyRotatingCode, isWithinCheckInWindow } from '../utils/selfCheckIn';
import { buildHashDeepLink } from '../utils/deepLink';
import { isEventOver } from '../utils/eventFormat';
import { isExternalEmail } from '../utils/deloitteDomain';
import { registrationEmail, externalInviteInstructionEmail, externalInvitationEmail, waitlistEmail, buildEmailFromTemplate, loadLogosAsBase64, wrapTemplate, qrCodeEmail, teamInfoBlockHtml, injectIntoEmailContent } from '../services/EmailTemplates';
import { buildUnsentEmlDraft } from '../utils/emlDraft';
import { readPendingShadowParents, removePendingShadowParent, addPendingShadowParent } from '../utils/shadowHeal';
import { withParentTitleSubject } from '../utils/mailSubject';
import { APP_VERSION } from '../version';
import { BundledItem, bundledCommOf, bundledItemsTableHtml, bundledItemsHeading } from '../utils/bundledComm';
import { buildDemoShowcaseEvents, isDemoShowcaseId, buildDemoRegistrations } from '../services/demoShowcaseEvent';
import { looksLikeClaimName, resolveMyDisplayName, safeDisplayName } from '../utils/displayName';
import { emitBootStage } from '../utils/bootProgress';
import { getCachedImage } from '../utils/imageCache';
import { dlog, isDebug } from '../utils/debugLog';

// v30.66: Reine Modul-Helfer (kein State-Zugriff) liegen jetzt in
// `eventTextHelpers.ts`. Hier re-exportiert, damit bestehende Importe aus
// `EventContext` unveraendert tragen.
import { applyEventTemplateOverride, stripSpNoteWrapper, formatOrganizerList, collectCcEmailsFromFields, mergeCcLists, summarizeCustomFields, buildEventUpdateDiff } from './eventTextHelpers';
export { applyEventTemplateOverride, stripSpNoteWrapper, formatOrganizerList, collectCcEmailsFromFields, mergeCcLists, summarizeCustomFields, buildEventUpdateDiff };
// v30.66: Das SP->App-Mapping liegt in `eventMapping.ts`; es bekommt die
// Subsite-Map als Ref herein, weil es sich die Subsite-URL je Event merkt.
import { mapSPEventToDeloitteEvent } from './eventMapping';
import { makeBillingActions } from './actions/billing';
import { makeOrganizerRoleActions } from './actions/organizerRoles';
import { makeInactiveAccountActions } from './actions/inactiveAccounts';
import { makeMaintenanceActions } from './actions/maintenance';
import { makeAutoMailActions } from './actions/autoMails';
import { makeCancellationActions } from './actions/cancellation';
import { makeArchiveActions } from './actions/archiveAndPurge';
import { makeAssistantActions } from './actions/assistantAndProxy';
import { makeParticipantFileActions } from './actions/participantFiles';
import { makeMailActions } from './actions/mails';

import { EventContextType, CreateEventInput, SelfCheckInParams, SelfCheckInStatus, SelfCheckInResult, EventStatsRow, FixColumnsDetail } from './eventContextTypes';
import { CounterStats } from '../services/events/seats';
// v30.66: Die Typen liegen jetzt in `eventContextTypes.ts`; hier nur noch
// re-exportiert, damit bestehende Importe aus `EventContext` weiter tragen.
export { EventContextType, CreateEventInput, SelfCheckInParams, SelfCheckInStatus, SelfCheckInResult, EventStatsRow, FixColumnsDetail };

export const EventContext = React.createContext<EventContextType | undefined>(undefined);

export function EventProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const [events, setEvents] = React.useState<DeloitteEvent[]>([]);
  const [isEventsLoading, setIsEventsLoading] = React.useState(true);
  // Map von EventId -> SubsiteUrl für schnellen Zugriff
  const subsiteMap = React.useRef<Record<string, string>>({});

  const eventService = React.useMemo(() => new EventService(props.context), []);

  // v9.16/v9.21: Test-Team war kurz global (TestTeamEmails in _Config),
  // ist jetzt per-Event (event.testTeamEmails). Globaler State raus.
  const currentUserEmail = props.context.pageContext.user.email;
  /**
   * v28.64: `pageContext.user.displayName` liefert bei einzelnen Personen das
   * Claims-Login-Token statt des Namens („0#.f|membership|user@deloitte.de").
   * Der Wert kommt aus der versteckten „User Information List" der Site, in
   * die SharePoint den Namen beim ersten Kontakt stempelt — stand dort damals
   * kein Anzeigename, bleibt das Token dauerhaft stehen, auch wenn das Profil
   * längst stimmt. Deshalb: erkennen, aus dem Benutzerprofil nachladen und bis
   * dahin die E-Mail zeigen statt des Tokens. Details in utils/displayName.ts.
   */
  const rawUserName = props.context.pageContext.user.displayName;
  const [profileUserName, setProfileUserName] = React.useState('');
  React.useEffect(() => {
    if (!looksLikeClaimName(rawUserName)) return;
    resolveMyDisplayName(props.context)
      .then(n => { if (n) setProfileUserName(n); })
      .catch(() => { /* Fallback bleibt die E-Mail */ });
  }, [rawUserName]);
  const currentUserName = profileUserName || safeDisplayName(rawUserName, currentUserEmail);
  // Vorname für E-Mail-Anreden ({{Name}} im Template).
  // Deloitte-displayName ist "Nachname, Vorname" (mit Komma) -> Teil nach Komma.
  // Fallback: displayName ohne Komma -> erstes Wort (vereinzelt "Vorname Nachname").
  const getFirstName = (displayName: string): string => {
    if (!displayName) return '';
    const commaIdx = displayName.indexOf(',');
    if (commaIdx >= 0) return displayName.substring(commaIdx + 1).trim().split(/\s+/)[0];
    return displayName.trim().split(/\s+/)[0];
  };
  const currentUserFirstName = getFirstName(currentUserName);

  // v30.18: Kalender-Eltern eines Tages auflösen — für den Mail-Betreff
  // (withParentTitleSubject, s. utils/mailSubject).
  const calDayParentOf = (ev: { parentEventId?: string } | undefined): DeloitteEvent | undefined => {
    if (!ev || !ev.parentEventId) return undefined;
    const p = events.find(x => x.id === ev.parentEventId);
    return p && p.subEventCalendar ? p : undefined;
  };

  // ==================== Sub-Event-Helper (v6.4+) ====================
  // Seit v6.4 sind Sub-Events keine separaten JSON-Arrays mehr, sondern
  // eigene DEX_Events-Items mit gesetztem parentEventId. Damit funktionieren
  // alle bestehenden Flows (DEX_CreateOutlookEvent, DEX_Outlook_Einladungen,
  // Teilnehmerliste, Organizer-Kalendereinladungen, Declines, QR-Codes,
  // Warteliste, ...) unverändert — ein Sub-Event ist einfach ein Event.
  // v13.11: topLevelEvents wird unten direkt aus `eventsForConsumer`
  // gefiltert, weil die Demo-Impersonation pro Event qrScannerEmails
  // shadowen muss.
  const childEventsOf = React.useCallback(
    (parentEventId: string): DeloitteEvent[] => {
      if (!parentEventId) return [];
      return events
        .filter(e => e.parentEventId === parentEventId)
        .slice()
        .sort((a, b) => (a.startDate || '').localeCompare(b.startDate || ''));
    },
    [events]
  );

  React.useEffect(() => {
    initEvents().catch(() => setIsEventsLoading(false));
  }, []);

  async function initEvents(): Promise<void> {
    // v11.74: Performance-Profiling — misst jede Boot-Phase und gibt am
    // Ende eine sortierte Tabelle aus. Nur in der Console sichtbar
    // (DevTools → Console), kein UI-Impact. Hilft beim Identifizieren
    // der echten Bottlenecks (ensure-Listen vs. getEvents vs. counts
    // vs. attachments).
    //
    // v11.76: Schema-Ensure-Gate. Die idempotenten `ensure*`/`upgrade*`-
    // Wartungs-Calls müssen NICHT bei jedem Page-Load laufen. Sie sind nur
    // beim ersten Boot nach einer App-Version mit Schema-Änderungen nötig.
    // Wir verwenden einen versions-gebundenen localStorage-Key — sobald
    // wir die Version 11.76 erfolgreich durchgeschmurgelt haben, sparen
    // wir uns alle ensure-Calls beim nächsten Boot.
    //
    // Außerdem: wenn die ensure-Calls DOCH laufen, parallelisieren wir sie
    // (Stage 1 = ensureEventsList alleine; Stage 2 = alles andere parallel
    // via Promise.allSettled), statt sie sequentiell hintereinander zu
    // ketten. Spart bei 11 Calls und ca. 6.7 s seriell ca. 4-5 s.
    const ENSURE_FLAG_KEY = 'dex.schema.ensured.v' + APP_VERSION;
    // v29.47: Zweiter Schlüssel OHNE Version — er sagt nur „diese App lief hier
    // schon einmal". Das entscheidet, ob die Schema-Pflege den Start blockieren
    // darf: Beim allerersten Start muss sie es (ohne Listen kein Lesen), nach
    // einem Update dagegen sind die Listen längst da, und die 22 Prüf-Anfragen
    // haben den Boot nur ausgebremst. Sie laufen dann NACH dem ersten Bild.
    const ENSURE_EVER_KEY = 'dex.schema.everEnsured';
    let skipEnsure = false;
    let everEnsured = false;
    try {
      if (typeof window !== 'undefined') {
        skipEnsure = window.localStorage.getItem(ENSURE_FLAG_KEY) === '1';
        everEnsured = window.localStorage.getItem(ENSURE_EVER_KEY) === '1';
      }
    } catch { /* localStorage disabled */ }

    const perfMarks: Array<{ name: string; ms: number }> = [];
    const tBoot = performance.now();
    const stage = async (name: string, fn: () => Promise<unknown>): Promise<void> => {
      const t0 = performance.now();
      try { await fn(); } catch { /* swallow — matches existing try/catch pattern */ }
      perfMarks.push({ name, ms: Math.round(performance.now() - t0) });
    };
    const safeRun = async (name: string, fn: () => Promise<unknown>, acc: Array<{ name: string; ms: number }>): Promise<void> => {
      const t0 = performance.now();
      try { await fn(); } catch { /* swallow */ }
      acc.push({ name, ms: Math.round(performance.now() - t0) });
    };

    const runEnsureStage = async (): Promise<void> => {
      // v29.41: Der Start-Balken bekommt echte Abschnitte statt einer reinen
      // Zeitschätzung — diese Stage ist die teuerste und läuft nur beim ersten
      // Boot je Version.
      emitBootStage('schema');
      // Stage 1: DEX_Events anlegen/sichern (Listen-Erstellung muss als erstes;
      // die upgrade*-Calls operieren auf DEX_Events).
      await stage('ensureEventsList', () => eventService.ensureEventsList());
      // Stage 2: alles andere parallel — keine inter-Abhängigkeiten.
      const parallelMarks: Array<{ name: string; ms: number }> = [];
      const tPar = performance.now();
      // Hinweis: safeRun() swallowt Exceptions intern und resolved IMMER. Daher
      // ist Promise.all hier sicher (kein early-reject) und auch in ES2018-
      // Targets verfügbar — Promise.allSettled wäre erst ab ES2020.
      await Promise.all([
        safeRun('upgradeAudienceFieldToNote', () => eventService.upgradeAudienceFieldToNote(), parallelMarks),
        safeRun('upgradeOrganizerFieldsToNote', () => eventService.upgradeOrganizerFieldsToNote(), parallelMarks),
        safeRun('upgradeOverflowTextFieldsToNote', () => eventService.upgradeOverflowTextFieldsToNote(), parallelMarks),
        safeRun('ensureEmailsList', () => eventService.ensureEmailsList(), parallelMarks),
        safeRun('ensureOutlookList', () => eventService.ensureOutlookList(), parallelMarks),
        safeRun('ensureParticipantsList', () => eventService.ensureParticipantsList(), parallelMarks),
        safeRun('ensureEmailTemplatesList', () => eventService.ensureEmailTemplatesList(), parallelMarks),
        safeRun('ensureIDReorderList', () => eventService.ensureIDReorderList(), parallelMarks),
        safeRun('ensureChangeLogList', () => eventService.ensureChangeLogList(), parallelMarks),
        safeRun('ensureEventCommsList', () => eventService.ensureEventCommsList(), parallelMarks),
        safeRun('ensureTeamJoinRequestsList', () => eventService.ensureTeamJoinRequestsList(), parallelMarks),
        safeRun('ensureOutlookLocksList', () => eventService.ensureOutlookLocksList(), parallelMarks),
        safeRun('ensureAccessFixList', () => eventService.ensureAccessFixList(), parallelMarks),
        safeRun('ensureAssistantAccessList', () => eventService.ensureAssistantAccessList(), parallelMarks),
        safeRun('ensureInactiveNoticesList', () => eventService.ensureInactiveNoticesList(), parallelMarks),
        safeRun('ensureArchiveList', () => eventService.ensureArchiveList(), parallelMarks),
        safeRun('ensureWeeklyReportsList', () => eventService.ensureWeeklyReportsList(), parallelMarks),
        safeRun('ensureTicketsList', () => eventService.ensureTicketsList(), parallelMarks),
        safeRun('ensureOrganizerRequestsList', () => eventService.ensureOrganizerRequestsList(), parallelMarks),
        safeRun('ensureOrganizerArchivedList', () => eventService.ensureOrganizerArchivedList(), parallelMarks),
        safeRun('ensureAssetsFolders', () => eventService.ensureAssetsFolders(), parallelMarks),
        safeRun('ensureLogosInConfig', () => eventService.ensureLogosInConfig(), parallelMarks),
      ]);
      const dPar = Math.round(performance.now() - tPar);
      // eslint-disable-next-line no-console
      dlog('perf', `[DEX][perf][boot] ensure-parallel-stage = ${dPar} ms`);
      // Einzelne Sub-Zeiten in die Gesamt-Tabelle übernehmen.
      for (const m of parallelMarks) perfMarks.push(m);
      // Erfolg markieren — nächster Boot überspringt die ensure-Calls.
      try {
        if (typeof window !== 'undefined') {
          window.localStorage.setItem(ENSURE_FLAG_KEY, '1');
          window.localStorage.setItem(ENSURE_EVER_KEY, '1');
        }
      } catch { /* localStorage disabled */ }
    };

    // v29.47: Erstinstallation → blockierend (ohne Listen gibt es nichts zu
    // lesen). Alle späteren Fälle → nach dem ersten Bild, im Hintergrund.
    const ensureAfterBoot = !skipEnsure && everEnsured;
    if (!skipEnsure && !everEnsured) {
      await runEnsureStage();
    }

    // loadLogosAsBase64 ist KEIN ensure-Call — es füllt den In-Memory-Cache
    // mit den Logo-Daten, die für Mail-/Outlook-Templates gebraucht werden.
    // Muss bei jedem Boot laufen.
    // v29.47: Die Logos stecken NUR in Mail- und Outlook-Vorlagen — für die
    // Anzeige braucht sie niemand. Sie blockierten den Start trotzdem, also
    // laufen sie jetzt nebenher; wer eine Mail baut, liest sie ohnehin aus dem
    // Cache (und lädt sie im Zweifel selbst nach).
    void stage('loadLogosAsBase64', () => loadLogosAsBase64(props.context.spHttpClient, eventService.siteUrl));
    await stage('loadEvents (Anzeige-Kette)', () => loadEvents());
    setIsEventsLoading(false);
    // Schema-Pflege nach dem ersten Bild — sie sichert nur Spalten und Listen,
    // die zum Lesen längst existieren.
    if (ensureAfterBoot) {
      void runEnsureStage().catch(err => console.warn('[DEX] Schema-Pflege im Hintergrund fehlgeschlagen:', err));
    }
    const tTotal = Math.round(performance.now() - tBoot);
    const sorted = [...perfMarks].sort((a, b) => b.ms - a.ms);
    if (skipEnsure) {
      // eslint-disable-next-line no-console
      dlog('perf', `[DEX][perf][boot] total = ${tTotal} ms (schema-ensure SKIPPED, version=v${APP_VERSION})`);
    } else {
      // eslint-disable-next-line no-console
      dlog('perf', `[DEX][perf][boot] total = ${tTotal} ms (schema-ensure RAN, version=v${APP_VERSION})`);
    }
    // v30.65: Auch die Tabelle nur auf Anforderung — sie war die „Array(2)"-Zeile
    // im gemeldeten Log.
    if (isDebug('perf')) {
      // eslint-disable-next-line no-console
      console.table(sorted.map(m => ({ stage: m.name, ms: m.ms })));
    }
  }

  async function loadEvents(): Promise<void> {
    // v29.47 (Performance-Audit): Der Boot hing an den PRO-EVENT-Abfragen.
    // Bei ~40 Events waren das über 80 Requests — je Event einmal die
    // Teilnehmerzahl aus der Subsite und einmal die Anhänge —, und ALLE davon
    // wurden abgewartet, bevor überhaupt etwas auf dem Bildschirm stand.
    // SharePoint drosselt bei so vielen gleichzeitigen Anfragen zusätzlich, das
    // war der Unterschied zwischen „zwei Sekunden" und „zehn Sekunden".
    //
    // Jetzt gilt: Was zum ANZEIGEN nötig ist, blockiert. Alles andere kommt
    // danach und aktualisiert die Ansicht nach.
    emitBootStage('events');
    const tGet = performance.now();
    const spEvents = await eventService.getEvents();
    const dGet = Math.round(performance.now() - tGet);
    // eslint-disable-next-line no-console
    dlog('perf', `[DEX][perf][loadEvents] getEvents = ${dGet} ms (n=${spEvents.length})`);
    emitBootStage('mapping');
    const tMap = performance.now();
    // v9.41: jedes Event-Mapping einzeln in try/catch wrappen — wenn EIN
    // Event-Mapping fehlschlägt (z.B. weil eine frisch erstellte Subsite noch
    // nicht API-konsistent ist), kippt nicht die ganze Eventliste in einen
    // Fehlerzustand. Stattdessen wird der einzelne kaputte Event ausgelassen
    // und beim nächsten Refresh erneut versucht. (Kein Promise.allSettled
    // benutzt, weil die SPFx-tsconfig auf ES2018 steht.)
    const safeMapped = await Promise.all(spEvents.map(async (e) => {
      try {
        return await mapSPEventToDeloitteEvent(e, subsiteMap);
      } catch (err) {
        console.warn('[DEX] mapSPEventToDeloitteEvent fehlgeschlagen für Event', e?.Id, err);
        return null;
      }
    }));
    const mapped = safeMapped.filter((x): x is DeloitteEvent => x !== null);
    const dMap = Math.round(performance.now() - tMap);
    // eslint-disable-next-line no-console
    dlog('perf', `[DEX][perf][loadEvents] mapSPEventToDeloitteEvent x ${spEvents.length} = ${dMap} ms`);

    // AB HIER ist die App bedienbar: Titel, Zeiten, Sichtbarkeit, Bilder und
    // die zuletzt gespeicherte Teilnehmerzahl (Spalte CurrentParticipants)
    // stehen im Mapping. Die Liste geht deshalb sofort raus.
    // v30.67: Bereits nachgeladene Anhänge behalten — s. keepLoadedDocuments.
    setEvents(prev => keepLoadedDocuments(prev, inheritParentImages(mapped)));

    // Nachlauf, sichtbar nur als Zahlen, die sich still aktualisieren.
    void (async () => {
      try {
        const tCnt = performance.now();
        const withCounts = await loadParticipantCountsForEvents(mapped);
        // eslint-disable-next-line no-console
        dlog('perf', `[DEX][perf][loadEvents] participantCounts (nachgelagert) = ${Math.round(performance.now() - tCnt)} ms`);
        // v30.67: Nicht den Snapshot von VOR einem parallel eingetroffenen
        // ensureEventDocuments zurückschreiben — s. keepLoadedDocuments.
        setEvents(prev => keepLoadedDocuments(prev, inheritParentImages(withCounts)));
      } catch (err) { console.warn('[DEX] Teilnehmerzahlen-Nachlauf fehlgeschlagen:', err); }
    })();
  }

  /**
   * v29.47: Anhänge (Dokumente) eines Events bei Bedarf nachladen.
   *
   * Vorher lud der Boot die Anhänge ALLER Events — ein Request pro Event, nur
   * damit sie im Objekt stehen. Gelesen werden sie an genau zwei Stellen:
   * „Meine Events" (Download-Liste) und der Wizard beim Bearbeiten. Beides
   * betrifft eine Handvoll Events, nicht alle.
   *
   * Mehrfachaufrufe für dasselbe Event sind billig: Was einmal geladen wurde,
   * wird gemerkt (`documentsLoadedRef`), und parallele Aufrufe teilen sich
   * dieselbe laufende Abfrage.
   */
  const documentsLoadedRef = React.useRef<Set<string>>(new Set());
  const documentsInflightRef = React.useRef<Map<string, Promise<void>>>(new Map());
  async function ensureEventDocuments(eventIds: string[]): Promise<void> {
    const todo = (eventIds || [])
      .filter(id => !!id && /^\d+$/.test(id))
      .filter(id => !documentsLoadedRef.current.has(id));
    if (todo.length === 0) return;
    await Promise.all(todo.map(async (id) => {
      const running = documentsInflightRef.current.get(id);
      if (running) return running;
      const p = (async () => {
        try {
          const attachments = await eventService.getEventAttachments(Number(id));
          documentsLoadedRef.current.add(id);
          setEvents(prev => prev.map(e => (e.id === id ? { ...e, documents: attachments } : e)));
        } catch { /* ohne Dokumente weiterarbeiten — sie sind Beiwerk */ }
        finally { documentsInflightRef.current.delete(id); }
      })();
      documentsInflightRef.current.set(id, p);
      return p;
    }));
  }

  /**
   * v30.67 (Review): Anhänge eines Events neu laden, nachdem der Wizard sie
   * am Context vorbei geändert hat (eigene EventService-Instanz). Ohne das
   * trug keepLoadedDocuments die Liste von VOR dem Speichern in den frischen
   * State: gelöschte Datei mit totem Link, neue Datei fehlt — bis zum F5, und
   * ein erneut geöffneter Wizard zeigte die Änderung als nicht geschehen.
   */
  async function refreshEventDocuments(eventId: string): Promise<void> {
    const id = String(eventId || '');
    if (!id) return;
    const running = documentsInflightRef.current.get(id);
    if (running) { try { await running; } catch { /* */ } }
    documentsLoadedRef.current.delete(id);
    await ensureEventDocuments([id]);
  }

  /**
   * v30.68: Sub-Events ohne eigenes Bild tragen das Bild des Hauptevents
   * (Nutzer-Ansage 02.09.2026: „im Default auch das Hauptevent-Foto"). Nur
   * Anzeige — nichts wird kopiert oder hochgeladen; ein eigenes Bild am
   * Sub-Event gewinnt weiterhin. `imageInherited` sagt dem Wizard, dass das
   * Feld dort leer bleibt. Zoom/Ausschnitt (`imageDisplay`) wandern mit,
   * damit die Kachel gleich beschnitten ist.
   */
  function inheritParentImages(list: DeloitteEvent[]): DeloitteEvent[] {
    const byId: Record<string, DeloitteEvent> = {};
    for (const e of list) byId[e.id] = e;
    return list.map(e => {
      if (!e.parentEventId || (e.imageUrl || '').trim()) return e;
      const parent = byId[e.parentEventId];
      if (!parent || !(parent.imageUrl || '').trim()) return e;
      return { ...e, imageUrl: parent.imageUrl, imageDisplay: e.imageDisplay || parent.imageDisplay, imageInherited: true };
    });
  }

  /**
   * v30.67: Nachgeladene Anhänge über einen loadEvents hinweg behalten.
   *
   * Das Mapping setzt immer `documents: []`, der Merker `documentsLoadedRef`
   * sagt aber weiter „geladen" — zwei Zustände über dieselben Daten, die
   * nicht synchron waren. Nach jedem loadEvents (Antworten ändern, Sub-Event
   * abmelden) verschwand deshalb die Download-Box in „Meine Events" bis zum
   * F5, weil niemand mehr nachlud (der Effect dort hängt an der Anzahl der
   * Events, und die ändert sich nicht). Derselbe Verlust ohne Nutzeraktion,
   * wenn der Zähl-Nachlauf einen Snapshot von VOR einem parallel
   * eingetroffenen ensureEventDocuments zurückschrieb. Was einmal geladen
   * ist, wandert hier in die frisch gemappten Objekte.
   */
  function keepLoadedDocuments(prev: DeloitteEvent[], next: DeloitteEvent[]): DeloitteEvent[] {
    if (!prev || prev.length === 0) return next;
    const docsById: Record<string, DeloitteEvent['documents']> = {};
    for (const p of prev) {
      if (p.documents && p.documents.length > 0) docsById[p.id] = p.documents;
    }
    return next.map(e => (docsById[e.id] && (!e.documents || e.documents.length === 0)) ? { ...e, documents: docsById[e.id] } : e);
  }


/**
 * v29.47: `Promise.all` über eine ganze Eventliste stellt vierzig Anfragen
 * gleichzeitig — SharePoint drosselt das, und gedrosselte Anfragen sind
 * langsamer als der Reihe nach gestellte. Diese Variante hält höchstens
 * `limit` Aufrufe in der Luft und behält die Reihenfolge der Ergebnisse bei.
 */
async function mapLimited<T, R>(items: T[], limit: number, fn: (item: T, index: number) => Promise<R>): Promise<R[]> {
  const out: R[] = new Array(items.length);
  let next = 0;
  const worker = async (): Promise<void> => {
    for (;;) {
      const i = next++;
      if (i >= items.length) return;
      out[i] = await fn(items[i], i);
    }
  };
  const n = Math.max(1, Math.min(limit, items.length));
  await Promise.all(Array.from({ length: n }, () => worker()));
  return out;
}

  async function loadParticipantCountsForEvents(evts: DeloitteEvent[]): Promise<DeloitteEvent[]> {
    // v29.47: Höchstens sechs Abfragen gleichzeitig. Vorher gingen alle auf
    // einmal raus — bei vierzig Events drosselt SharePoint, und gedrosselte
    // Anfragen dauern länger als der Reihe nach gestellte.
    const results = await mapLimited(evts, 6, async (evt) => {
        if (!evt.subsiteUrl) return evt;
        // v29.47: Vergangene Events brauchen keine frische Zahl — sie ändert
        // sich nicht mehr, und die gespeicherte steht bereits im Objekt.
        if (isEventOver(evt)) return evt;
        try {
          const counts = await eventService.getRegistrationCount(evt.subsiteUrl);
          // v26.63: Frische Zahl best-effort nach DEX_Events.CurrentParticipants
          // zurückschreiben, wenn sie vom gespeicherten Wert abweicht. Klappt nur
          // für Organizer/Admins (Schreibrecht) — bei normalen Usern schlägt der
          // MERGE still fehl. Nur Haupt-/Sub-Events mit numerischer Item-Id.
          if (counts.registered !== evt.currentParticipants && /^\d+$/.test(evt.id)) {
            eventService.persistCurrentParticipants(Number(evt.id), counts.registered).catch(() => { /* best-effort */ });
          }
          return { ...evt, currentParticipants: counts.registered, waitlistCount: counts.waitlist };
        } catch {
          return evt;
        }
      });
    // v22.74: Klammer-Events („Nur Sub-Events") zeigen in der Listen-Karte die
    // EINDEUTIGE Personenzahl über alle Sub-Events — eine Person, die sich für
    // mehrere Sub-Events anmeldet, zählt EINMAL (nicht die Summe). Der eigene
    // Subsite-Counter der Klammer (Schatten-Zeilen) ist falsch; deshalb werden
    // die Sub-Event-Teilnehmer-Mails geladen und entdoppelt. Nur für Klammern
    // (Minderheit) → vertretbarer Extra-Aufwand.
    const klammerParents = results.filter(e => e.subEventsOnlyMode);
    if (klammerParents.length === 0) return results;
    const childrenByParent = new Map<string, DeloitteEvent[]>();
    for (const e of results) {
      if (e.parentEventId) {
        const arr = childrenByParent.get(e.parentEventId) || [];
        arr.push(e);
        childrenByParent.set(e.parentEventId, arr);
      }
    }
    const uniqueByParent = new Map<string, { active: number; waitlist: number }>();
    await Promise.all(klammerParents.map(async (p) => {
      const kids = childrenByParent.get(p.id) || [];
      const activeSet = new Set<string>();
      const waitSet = new Set<string>();
      await Promise.all(kids.map(async (c) => {
        if (!c.subsiteUrl) return;
        try {
          const r = await eventService.getParticipantEmailsByStatus(c.subsiteUrl);
          for (const em of r.active) activeSet.add(em);
          for (const em of r.waitlist) waitSet.add(em);
        } catch { /* best-effort */ }
      }));
      // Warteliste nur Personen, die NIRGENDS aktiv sind.
      for (const a of Array.from(activeSet)) waitSet.delete(a);
      uniqueByParent.set(p.id, { active: activeSet.size, waitlist: waitSet.size });
    }));
    return results.map(e => uniqueByParent.has(e.id)
      ? { ...e, currentParticipants: uniqueByParent.get(e.id)!.active, waitlistCount: uniqueByParent.get(e.id)!.waitlist }
      : e);
  }

  // v24.73: Live-Plätze aus dem Counter (für alle lesbar). Aktiv = SeatsTaken
  // (wird von reserveSeat für JEDE Anmeldung gepflegt → korrekt für alle),
  // Warteliste = WaitlistTaken. null → Counter (noch) nicht da → Aufrufer nutzt
  // den bisherigen Wert.
  async function getLiveCounterStats(eventId: string): Promise<CounterStats | null> {
    const event = events.find(e => e.id === eventId);
    const subsiteUrl = subsiteMap.current[eventId] || event?.subsiteUrl;
    if (!subsiteUrl) return null;
    const isSplit = !!event && typeof event.durchstarterCapacity === 'number'
      && typeof event.funstarterCapacity === 'number'
      && ((event.durchstarterCapacity || 0) > 0 || (event.funstarterCapacity || 0) > 0);
    try { return await eventService.getCounterStats(subsiteUrl, isSplit); } catch { return null; }
  }

  // v24.75: Echtzeit-Push auf eine Liste der Event-Subsite abonnieren.
  async function subscribeEventRealtime(eventId: string, kind: 'counter' | 'participants', onChange: () => void): Promise<() => void> {
    const event = events.find(e => e.id === eventId);
    const subsiteUrl = subsiteMap.current[eventId] || event?.subsiteUrl;
    if (!subsiteUrl) return () => { /* */ };
    return eventService.subscribeListRealtime(subsiteUrl, kind, onChange);
  }

  // v24.74: Counter aller aktiven Kapazitäts-Events frischziehen (privilegiert).
  async function reconcileCounters(opts?: { onlyMine?: boolean }): Promise<void> {
    const seen = new Set<string>();
    // v30.63: Organizer heilen ihre EIGENEN Events mit.
    //
    // Bisher lief der Abgleich nur beim Admin-Start, höchstens alle sechs
    // Stunden. Zwischen zwei Admin-Anmeldungen konnte der für alle lesbare
    // Zähler driften — nach oben, wenn eine Selbst-Abmeldung ihn bei
    // unbekanntem Wartelisten-Stand nicht herunterzählen durfte, nach unten,
    // weil der Nachrück-Flow Warteliste→Angemeldet promotet, ohne SeatsTaken
    // zu erhöhen. Seit die Anmeldeseite diesen Zähler ANZEIGT, ist das nicht
    // mehr nur eine Frage der Überbuchungs-Sperre.
    //
    // Gefährlich wäre das nur, wenn jemand ohne Vollzugriff den Zähler mit
    // einer beschnittenen Sicht überschreibt — genau der Fehler vom 15.07.
    // Das kann hier nicht passieren: `getActiveCounts` vergleicht die
    // gelesene Zeilenzahl mit dem echten ItemCount der Liste und wirft bei
    // unvollständiger Sicht, `syncSeatsToActiveCount` schreibt dann nichts.
    // Ein Organizer ohne Rechte auf eine Subsite heilt sie also schlicht
    // nicht — er kann sie nicht kaputtmachen.
    const mine = (e: DeloitteEvent): boolean => {
      const me = (currentUserEmail || '').toLowerCase();
      if (!me) return false;
      return [...(e.organizerEmails || []), ...(e.coOrganizerEmails || [])]
        .some(x => (x || '').toLowerCase() === me);
    };
    // v30.67: Effektive Kapazität statt maxParticipants. Bei geteilten
    // Kapazitäten ist maxParticipants per Konvention 0 (CLAUDE.md) — der
    // Filter schloss damit genau die Events aus, die als einzige ZWEI Zähler
    // synchron halten müssen; das `isSplit` unten war nie erreichbar. Folge:
    // Eine Selbst-Abmeldung bei unbekanntem Wartelisten-Stand ließ den
    // Gruppenzähler dauerhaft zu hoch, reserveSeat meldete 'full' bei freien
    // Plätzen, und die App heilte das von sich aus nie.
    const effectiveCap = (e: DeloitteEvent): number => (e.maxParticipants || 0) > 0
      ? (e.maxParticipants || 0)
      : ((e.durchstarterCapacity || 0) + (e.funstarterCapacity || 0));
    const targets = (events || []).filter(e =>
      e.status === 'Active' && effectiveCap(e) > 0 && (e.subsiteUrl || '').trim()
      && (!opts?.onlyMine || mine(e)));
    for (const e of targets) {
      const sub = e.subsiteUrl as string;
      if (seen.has(sub)) continue;
      seen.add(sub);
      const isSplit = typeof e.durchstarterCapacity === 'number'
        && typeof e.funstarterCapacity === 'number'
        && ((e.durchstarterCapacity || 0) > 0 || (e.funstarterCapacity || 0) > 0);
      // Sequentiell wegen SP-Throttling; best-effort, blockiert nie.
      try { await eventService.syncSeatsToActiveCount(sub, { isSplit }); } catch { /* */ }
    }
  }

  async function refreshParticipantCounts(eventId?: string): Promise<void> {
    if (eventId) {
      const subsiteUrl = subsiteMap.current[eventId];
      if (!subsiteUrl) return;
      try {
        const counts = await eventService.getRegistrationCount(subsiteUrl);
        // v19.17: referenzschonend — nur das betroffene Event ersetzen, und nur
        // wenn sich die Zahl tatsächlich geändert hat. Bleibt alles gleich, wird
        // dieselbe Array-Referenz zurückgegeben → React rendert NICHT neu (kein
        // sichtbares Flackern).
        setEvents(current => {
          let changed = false;
          const next = current.map(e => {
            if (e.id === eventId && (e.currentParticipants !== counts.registered || e.waitlistCount !== counts.waitlist)) {
              changed = true;
              return { ...e, currentParticipants: counts.registered, waitlistCount: counts.waitlist };
            }
            return e;
          });
          return changed ? next : current;
        });
      } catch { /* default bleibt */ }
    } else {
      // Alle Events: frische Zahlen laden, dann NUR die geänderten Events durch
      // neue Objekte ersetzen (unveränderte behalten ihre Referenz). Ändert sich
      // gar nichts, bleibt die Array-Referenz gleich → kein Re-Render.
      setEvents(current => {
        loadParticipantCountsForEvents(current).then(updated => {
          setEvents(prev => {
            let changed = false;
            const next = prev.map(e => {
              const u = updated.find(x => x.id === e.id);
              if (u && (u.currentParticipants !== e.currentParticipants || u.waitlistCount !== e.waitlistCount)) {
                changed = true;
                return { ...e, currentParticipants: u.currentParticipants, waitlistCount: u.waitlistCount };
              }
              return e;
            });
            return changed ? next : prev;
          });
        }).catch(() => { /* ignore */ });
        return current;
      });
    }
  }


  async function createEvent(input: CreateEventInput): Promise<number | null> {
    const eventId = await eventService.createEvent(input);
    if (eventId) {
      // v9.0: Audit-Log (fire-and-forget — Save-Flow nicht blocken)
      eventService.writeChangeLog({
        action: 'EventCreated',
        targetType: 'Event',
        targetId: String(eventId),
        targetName: input.title,
        eventId: String(eventId),
        eventTitle: input.title,
        // v26.19: Beschreibung mit ins Audit-Log aufnehmen (falls vorhanden) —
        // als reiner Text gekürzt, damit der Eintrag lesbar bleibt. Bei
        // EventUpdated wird sie ohnehin über den Feld-Diff erfasst.
        details: (() => {
          const descPlain = (input.description || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const fieldsSummary = summarizeCustomFields((input as any).customFields);
          return {
            eventType: input.type, location: input.location, startDate: input.startDate, maxParticipants: input.maxParticipants,
            ...(descPlain ? { description: descPlain.length > 300 ? `${descPlain.slice(0, 300)}…` : descPlain } : {}),
            ...(fieldsSummary && fieldsSummary !== '(keine Felder)' ? { fields: fieldsSummary } : {}),
          };
        })(),
      }).catch(() => { /* */ });
      // v11.53: KPI-Counter sofort hochzählen — nur für nicht-fictive Events
      // (Test-Events zählen nicht in der LandingPage-KPI).
      // v23.38: Sub-Events (parentEventId) NICHT mitzählen — die Events-KPI
      // zählt nur eigenständige Haupt-/Klammer-Events (sonst bläht jedes
      // Sub-Event den Zähler wieder auf, gegen den die Boot-Neuberechnung läuft).
      if (!input.isFictive && !input.parentEventId) {
        eventService.bumpKpiEvents(1).catch(() => { /* best-effort */ });
      }
      // v9.41: KEIN Auto-Refresh mehr direkt nach Create. Grund: SharePoint braucht
      // einige Sekunden, bis die frisch angelegte Subsite + Teilnehmerliste +
      // DEX_TeilnehmerCounter-Liste API-seitig konsistent abrufbar sind. Wenn wir
      // hier sofort getEvents() + mapSPEventToDeloitteEvent() laufen lassen, fallen
      // die Subsite-Reads mit 400/404 ins Leere und das nachfolgende Event-List-
      // Rendering kann in eine Render-Loop laufen (React #300 → weißer Screen).
      //
      // Stattdessen wird der Refresh erst getriggert, wenn der User auf der Success-
      // Seite auf 'Events anzeigen' klickt — bis dahin hatte SP genug Zeit zum
      // Propagieren. Falls jemand auf der Erfolgs-Seite stehen bleibt und nichts
      // klickt, wird beim nächsten Page-Mount ohnehin gerefreshed.
    }
    return eventId;
  }

  /**
   * v30.67: Aktiv-Prüfung mit drei Antworten — true = angemeldet, false = frei,
   * null = nicht prüfbar (429/403) → der Aufrufer lehnt ab (fail-closed).
   *
   * Der Service liefert heute bei JEDEM Fehler `false` (`if (!response.ok)
   * return false; catch { return false; }`). Damit war der v30.14-Zweig
   * `alreadyActive === null` in registerForEvent toter Code, und in der
   * Anmeldewelle kam die zweite Zeile durch — genau der Befund, den v30.14
   * beheben sollte. Die Aufrufstellen hier sind dreiwertig; sobald
   * `isUserAlreadyOnEvent` bei Fehler `null` liefert (Service-Änderung, im
   * v30.67-Bericht beschrieben), greift der Zweig ohne weiteres Zutun.
   */
  const probeAlreadyActive = async (subsiteUrl: string, email: string): Promise<boolean | null> =>
    eventService.isUserAlreadyOnEvent(subsiteUrl, email);

  // v30.42: Je Sitzung EINMAL je (Klammer, Person) die Schattenzeile prüfen.
  // Ohne diesen Merker liefe ein Massen-Lauf über 19 Termine 19-mal in
  // dieselbe Prüfung — auf einem Pfad, der ohnehin drosselungsempfindlich ist.
  // v30.68: Sitzungs-Merker MIT Verfall. Ein Lauf über 19 Termine prüft die
  // Klammer-Zeile einmal, nicht 19-mal — aber nicht für immer: Wird die
  // Klammer in derselben Sitzung abgemeldet (Organizer Center), sagte ein
  // ewiger Merker weiter „steht", und der nächste Termin hätte keine Klammer.
  const shadowEnsuredRef = React.useRef<Map<string, number>>(new Map<string, number>());
  const SHADOW_ENSURED_TTL_MS = 5 * 60 * 1000;
  async function registerForEvent(
    eventId: string,
    customData: Record<string, string>,
    participantFirstName?: string,
    participantLastName?: string,
    participantEmail?: string,
    preferredStarterType?: string, // B2Run: 'Durchstarter' | 'Funstarter'
    // v18.13: Massenimport — pro Anmeldung Bestätigungsmail bzw. Outlook-Termin
    // unterdrücken („stille Anmeldung").
    // v18.53: extraCc — zusätzliche CC-Adressen, die NICHT aus den
    // event-eigenen Feldern stammen. Genutzt im subEventsOnlyMode: das
    // „Hauptevent" ist dort nicht anmeldbar, die CC-Felder (z.B. Assistenz)
    // sind übergreifend und gelten für die Sub-Events — also müssen die
    // Sub-Event-Bestätigungsmails ebenfalls an die Assistenz auf CC gehen.
    // v18.74: proxyConsentConfirmed — bei stellvertretender Anmeldung wurde die
    // Zustimmung der Person bestätigt (Pflicht-Checkbox auf der Anmeldeseite).
    // Wird als Nachweis in die SP-Spalte ProxyConsent geschrieben.
    // v30.9: skipReload — jeder Aufruf zog am Ende ein volles loadEvents
    // (28 MB + participantCounts) nach sich. Die Anmeldeseite meldet bei
    // Kalender-Events MEHRFACH an (Tage + Schatten-Klammer) und lud damit
    // pro Klick mehrere Male den kompletten Bestand; genau dieselbe Bremse
    // wie v29.77 im Wizard. Schleifen skippen und refreshen EINMAL am Ende.
    opts?: { skipShadowParent?: boolean; suppressMail?: boolean; suppressOutlook?: boolean; extraCc?: string; proxyConsentConfirmed?: boolean; actorAllowedAsAssistant?: boolean; skipReload?: boolean; bundledItems?: BundledItem[] }
  ): Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste'; reason?: string }> {
    // v17.25: Demo-Showcase-Event → No-Op, kein SP-Roundtrip. Die Register-
    // Seite blockt den Submit ohnehin mit einem Demo-Hinweis; dieser Guard
    // ist die zweite Verteidigungslinie.
    if (isDemoShowcaseId(eventId)) return { ok: true, status: 'Angemeldet' };
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, status: 'Warteliste' };

    // Vorname/Nachname aus displayName extrahieren falls nicht übergeben.
    // Deloitte-Profile liefern den Namen typischerweise als "Nachname, Vorname"
    // (Komma-Format aus dem Active Directory). Früher haben wir mit Space
    // gesplittet — das tauschte Vor- und Nachname und führte u.a. dazu, dass
    // bei Sub-Event-Anmeldungen (die ohne explizite Vor-/Nachname-Args laufen)
    // die "Anrede" mit dem Nachnamen geschrieben wurde.
    const parseDisplayName = (raw: string): { firstName: string; lastName: string } => {
      const dn = (raw || '').trim();
      if (!dn) return { firstName: '', lastName: '' };
      if (dn.indexOf(',') >= 0) {
        const parts = dn.split(',').map(s => s.trim());
        return { firstName: parts[1] || '', lastName: parts[0] || '' };
      }
      const parts = dn.split(/\s+/).filter(Boolean);
      if (parts.length === 0) return { firstName: '', lastName: '' };
      if (parts.length === 1) return { firstName: parts[0], lastName: '' };
      return { firstName: parts[0], lastName: parts.slice(1).join(' ') };
    };
    const parsed = parseDisplayName(currentUserName);
    const firstNameToUse = participantFirstName || parsed.firstName;
    const lastNameToUse = participantLastName || parsed.lastName;
    const emailToUse = participantEmail || currentUserEmail;
    const nameToUse = `${firstNameToUse} ${lastNameToUse}`.trim();

    // Prüfen ob schon registriert
    const event = events.find(e => e.id === eventId);
    // v27.11 (Bug „Externe können mehrfach angemeldet werden"): status-
    // gefilterter Aktiv-Check über isUserAlreadyOnEvent — greift für ALLE
    // Anmeldepfade (auch externe Personen, Massenimport, Admin/Assistenz),
    // ist unabhängig von der Zeilen-Reihenfolge (das $top=1-Read unten konnte
    // eine alte Abgemeldet-Zeile erwischen und aktive Duplikate übersehen)
    // und blockt aktive Doppel-Anmeldungen VOR jeder Sitzplatz-Reservierung.
    // Fail-open wie bisher (Item-Level-Security kann fremde Zeilen verbergen),
    // aber strikt weniger Duplikate als vorher.
    // v30.14: Fail-CLOSED statt fail-open, wenn die Prüfung selbst scheitert
    // (typisch: 429 während einer Anmeldewelle). Genau in dem Moment, in dem
    // die Prüfung wegen Drosselung ausfiel, hat der User es „noch einmal
    // versucht" — und die zweite Zeile kam durch (belegte Doppel-Warteliste
    // im Soft Opening, 3 Minuten Abstand). Lieber die Anmeldung mit klarer
    // Meldung ablehnen als eine Doppel-Zeile riskieren. Ausnahme bleibt die
    // Klammer-Schatten-Zeile (subEventsOnlyMode): dort ist eine doppelte
    // Zeile harmlos (belegt keinen Platz, Bereinigen-Knopf existiert), ein
    // Abbruch würde dagegen die „Fehlende Klammer"-Fälle vermehren.
    // v30.17: Für die Klammer-Schatten-Zeile (subEventsOnlyMode) läuft die
    // isUserAlreadyOnEvent-Prüfung gar nicht mehr — jedes ihrer Ergebnisse
    // führte ohnehin zu „weiter" (aktiv → ok:true, Fehler → exempt, nicht
    // aktiv → getMyRegistration entscheidet). Das spart einen SharePoint-
    // Request pro Anmeldung; die Idempotenz sichert weiterhin der
    // getMyRegistration-Check unten (v23.9-Guard). Für ECHTE Anmeldungen
    // (Sub-Events, normale Events) bleibt die Prüfung inkl. fail-closed.
    const isShadowParent = !!(event && event.subEventsOnlyMode);
    const alreadyActive = isShadowParent
      ? false
      : await eventService.isUserAlreadyOnEvent(subsiteUrl, (emailToUse || '').trim()).catch(() => null);
    if (alreadyActive === null) {
      return { ok: false, status: 'Warteliste', reason: 'dup-check-failed' };
    }
    if (alreadyActive) {
      return { ok: false, status: 'Warteliste', reason: 'already-registered' };
    }

    // v30.67 (Review): Auch dieser Lesevorgang ist keine Aussage, wenn er
    // scheitert — ein 429 hier hieß „keine Zeile" und führte in den Insert
    // statt in die Reaktivierung: zweite Zeile zur selben Adresse. Die
    // Klammer-Schattenzeile bleibt ausgenommen (dort ist die Dublette
    // harmlos; ein Abbruch vermehrte nur die „Fehlende Klammer"-Fälle).
    let regReadFailed = false;
    const existing = await eventService.getMyRegistration(subsiteUrl, emailToUse, () => { regReadFailed = true; });
    if (regReadFailed && !isShadowParent) {
      return { ok: false, status: 'Warteliste', reason: 'dup-check-failed' };
    }
    if (existing && existing.Status !== 'Abgemeldet') {
      // v23.9: Im Klammer-Modus (subEventsOnlyMode) ist die Parent-Zeile NUR
      // ein Schatten zur Datenvollständigkeit — die echte Anmeldung sind die
      // Sub-Events. Ein bereits vorhandener (nicht abgemeldeter) Schatten darf
      // deshalb die (Sub-Event-)Anmeldung NICHT blockieren. Das war die Ursache
      // des „bereits registriert"-Falls: eine Person mit aktiver Schatten-Zeile
      // (z.B. aus einer abgebrochenen Anmeldung), die in den Sub-Events nur eine
      // ABGEMELDETE Zeile hat, tauchte weder in der Liste noch im Geister-Kasten
      // auf, blockierte aber jede Neu-Anmeldung. Jetzt: Schatten als „schon da"
      // behandeln (ok:true, kein zweiter Insert), Sub-Events laufen weiter.
      if (event && event.subEventsOnlyMode) {
        return { ok: true, status: existing.Status === 'Warteliste' ? 'Warteliste' : 'Angemeldet' };
      }
      return { ok: false, status: 'Warteliste' };
    }

    // v30.73: Zweite Quelle für die Klammer-Schattenzeile — das Register
    // DEX_Participants auf der Haupt-Site. Die Teilnehmerlisten haben
    // Zeilen-Sicherheit („nur eigene Elemente", geprüft am Zeilen-AUTOR):
    // Eine Assistenz sieht die Klammer-Zeile nicht, die eine andere Assistenz
    // oder die Person selbst angelegt hat — `getMyRegistration` oben meldet
    // dann „keine Zeile", und der Insert unten legte eine zweite an. Dazu kam
    // seit v30.68, dass ein GESCHEITERTER Lesevorgang für die Schattenzeile
    // toleriert wurde („Dublette ist harmlos") und ebenfalls in den Insert
    // lief — bei 19 Office-Tagen mit Drosselung oft genug für drei Zeilen je
    // Person (Befund 03.09.2026, „Doppelte Klammer-Zeilen (3)").
    //
    // Das Register kennt keine Zeilen-Sicherheit und wird bei JEDER An- und
    // Abmeldung mitgeschrieben — auch für die Schattenzeile (upsertParticipant
    // unten). Steht die Klammer-Nummer dort, existiert die Zeile, nur nicht
    // sichtbar: Zielzustand, kein Insert. Ist EINE der beiden Quellen nicht
    // lesbar und die andere sagt nicht „vorhanden", wird NICHT eingefügt —
    // der Aufrufer (Klammer-zuerst in registerForEvent / submitFlow) zieht
    // über den zweiten Versuch und den Nachzug-Merker nach. Ein leeres
    // Ergebnis ohne geprüften Status ist keine Aussage (CLAUDE.md).
    if (isShadowParent && !existing) {
      let registryReadFailed = false;
      const rec = await eventService.getParticipantByEmail((emailToUse || '').trim(), () => { registryReadFailed = true; });
      const nums = (v?: string): number[] => v ? v.split(',').map(x => parseInt(x.trim(), 10)).filter(n => !isNaN(n)) : [];
      const parentNo = event ? (event.eventNumber || 0) : 0;
      const knownInRegistry = parentNo > 0 && !!rec
        && (nums(rec.EventRegistered).indexOf(parentNo) >= 0 || nums(rec.EventOnWaitlist).indexOf(parentNo) >= 0);
      if (knownInRegistry) {
        dlog('seats', '[DEX] Klammer-Zeile laut Register vorhanden (für diesen Nutzer unsichtbar) — kein zweiter Insert:', emailToUse, parentNo);
        return { ok: true, status: 'Angemeldet' };
      }
      if (regReadFailed || registryReadFailed) {
        console.warn('[DEX] Klammer-Zeile: Bestand nicht prüfbar (Liste/Register nicht lesbar) — kein Insert, Nachzug folgt:', emailToUse, parentNo);
        return { ok: false, status: 'Warteliste', reason: 'dup-check-failed' };
      }
    }

    // v30.68: Die Klammer-Zeile wird VOR dem Termin geschrieben — als
    // erster Schreibvorgang, nicht als letzter. Bis v30.67 kam sie zuletzt,
    // also genau dort, wo die Drosselung nach 19 Terminen greift, und ein
    // zugeklappter Tab kam nie bis dahin. Die eigentliche Ursache des roten
    // Kastens war aber die Fristprüfung im Service (registration.ts, Check
    // B): Sie wies die Schattenzeile mit der Frist des HAUPTEVENTS ab — seit
    // v30.68 gilt die Frist dort nicht für die Klammer eines „Nur Sub-
    // Events"-Events.
    //
    // Nutzer-Entscheidung 02.09.2026: Die Termine werden IMMER geschrieben,
    // auch wenn die Klammer scheitert — „wenn mit der Klammer was nicht
    // stimmt, werden die Leute nicht angemeldet, das will ich nicht". Dann
    // folgt nach dem Termin ein zweiter Versuch, danach der Nachzug-Merker
    // (utils/shadowHeal, beim nächsten App-Start). Der Sitzungs-Merker
    // verhindert, dass ein Lauf über 19 Termine 19-mal dieselbe Zeile prüft.
    //
    // `skipShadowParent` setzt nur die Anmeldeseite — sie schreibt die Klammer
    // selbst zuerst, MIT den übergreifenden Antworten (Hotel, Verpflegung,
    // Anreise).
    let umbrellaPending = false;
    if (event && event.parentEventId && !opts?.skipShadowParent) {
      const parentEv = events.find(e => e.id === event.parentEventId);
      if (parentEv && parentEv.subEventsOnlyMode) {
        const guardKey = `${parentEv.id}|${(emailToUse || '').toLowerCase().trim()}`;
        const ensuredAt = shadowEnsuredRef.current.get(guardKey) || 0;
        if (Date.now() - ensuredAt > SHADOW_ENSURED_TTL_MS) {
          let umbrellaOk = false;
          try {
            const shadow = await registerForEvent(
              parentEv.id, {}, firstNameToUse, lastNameToUse, emailToUse, undefined,
              {
                suppressMail: true, suppressOutlook: true, skipReload: true, skipShadowParent: true,
                ...(opts?.proxyConsentConfirmed
                  ? { proxyConsentConfirmed: true, actorAllowedAsAssistant: !!opts?.actorAllowedAsAssistant }
                  : {}),
              }
            );
            // `already-registered` heißt: Die Zeile steht — Zielzustand.
            umbrellaOk = shadow.ok || shadow.reason === 'already-registered';
          } catch (err) {
            console.warn('[DEX] Klammer-Zeile konnte nicht sichergestellt werden:', err);
          }
          if (umbrellaOk) shadowEnsuredRef.current.set(guardKey, Date.now());
          else {
            umbrellaPending = true;
            console.warn('[DEX] registerForEvent: Klammer-Zeile steht (noch) nicht — Termin wird trotzdem angelegt, Klammer wird nachgezogen:', parentEv.id, emailToUse);
          }
        }
      }
    }

    // Prüfen ob Platz frei oder Waitlist
    let status: 'Angemeldet' | 'Warteliste' = 'Angemeldet';
    let effectiveStarterType: string | undefined = preferredStarterType;

    // B2Run Split-Capacity Logik (seit v6.5): getrennte Wartelisten pro StarterType.
    // Wenn der Wunsch-Typ noch freie Plätze hat → direkt angemeldet mit diesem Typ.
    // Wenn der Wunsch-Typ voll ist → landet auf der Warteliste MIT gesetztem
    // PreferredStarterType (kein stiller Fallback auf den anderen Typ mehr).
    // Die Entscheidung "möchte ich auf den anderen Typ umsteigen" trifft der User
    // explizit im UI (RegistrationPage Pre-Check-Dialog), bevor er hier reinkommt —
    // dann hat preferredStarterType bereits den neuen Wunsch.
    const isSplitGroup = !!event && typeof event.durchstarterCapacity === 'number' && typeof event.funstarterCapacity === 'number'
      && (event.durchstarterCapacity > 0 || event.funstarterCapacity > 0);
    // v11.36: Atomare Sitzplatz-Reservierung statt des alten read-then-write
    // (TOCTOU-Race + fail-open-catch → Massen-Überbuchung bei Anmeldewellen).
    // reserveSeat inkrementiert den Gruppen-Counter per ETag-CAS:
    //   'reserved' → Platz sicher belegt → Angemeldet
    //   'full'     → Gruppe voll → Warteliste
    //   'error'    → Counter nicht nutzbar → FAIL-CLOSED → Warteliste
    //                (NIE optimistisch Angemeldet — genau das war der Bug).
    // v30.67: Merken, WELCHER Zähler wirklich erhöht wurde — für die Rückgabe,
    // falls der Insert danach scheitert (s.u.). cap <= 0 heißt „unbegrenzt":
    // reserveSeat liefert dann 'reserved', ohne den Zähler anzufassen.
    let reservedSeatField: 'SeatsTaken' | 'SeatsTakenDurch' | 'SeatsTakenFun' | null = null;
    if (event && isSplitGroup && preferredStarterType) {
      const cap = preferredStarterType === 'Durchstarter'
        ? (event.durchstarterCapacity || 0)
        : (event.funstarterCapacity || 0);
      const seat = await eventService.reserveSeat(subsiteUrl, preferredStarterType as 'Durchstarter' | 'Funstarter', cap);
      if (seat === 'reserved') {
        effectiveStarterType = preferredStarterType;
        if (cap > 0) reservedSeatField = eventService.seatFieldFor(preferredStarterType);
      } else {
        // 'full' ODER 'error' → fail-closed auf Warteliste für genau diesen Typ.
        status = 'Warteliste';
        effectiveStarterType = undefined; // wird erst beim Nachrücken gesetzt
      }
    } else if (event && isSplitGroup && !preferredStarterType) {
      // Split-Event ohne Gruppenwahl: sicherste Variante ist Warteliste
      // (die UI erzwingt normalerweise eine Gruppenwahl; das ist der Schutz
      // gegen den Pfad, der früher ungebremst auf Angemeldet fiel).
      status = 'Warteliste';
      effectiveStarterType = undefined;
    } else if (event && event.maxParticipants > 0) {
      const seat = await eventService.reserveSeat(subsiteUrl, '', event.maxParticipants);
      if (seat !== 'reserved') {
        status = 'Warteliste';
      } else {
        reservedSeatField = 'SeatsTaken';
      }
    }

    // v27.11: Warteliste abgeschaltet → volle Events nehmen KEINE neuen
    // Anmeldungen mehr an. Vorher fiel die Anmeldung trotz deaktivierter
    // Warteliste still auf Status 'Warteliste' — der Toggle war wirkungslos.
    // (Kein Counter-Rollback nötig: status 'Warteliste' heißt, reserveSeat
    // hat NICHT reserviert. Schatten-Zeilen im subEventsOnlyMode bleiben
    // ausgenommen — sie sind keine echte Parent-Anmeldung.)
    if (status === 'Warteliste' && event && event.waitlistEnabled === false && !event.subEventsOnlyMode) {
      return { ok: false, status: 'Warteliste', reason: 'full' };
    }

    // FieldMap aus Custom Fields extrahieren (cf.id -> spInternalName)
    const fieldMap: Record<string, string> = {};
    if (event) {
      for (const f of event.eventSpecificFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const spName = (f as any).spInternalName;
        if (spName) fieldMap[f.id] = spName;
      }
    }

    // Audit-Trail: wer klickt gerade "Register"? = der eingeloggte User.
    // Bei Self-Registration ist das = der Teilnehmer selbst, bei "Für andere
    // Person registrieren" ist das der Organizer/Admin (Teilnehmer-Daten
    // wurden über participantFirstName/participantEmail übergeben).
    const actorName = currentUserName;
    const actorEmail = currentUserEmail;

    // v18.74: Nachweis der Zustimmung bei stellvertretender Anmeldung. Eine
    // Anmeldung gilt als „stellvertretend", wenn die Teilnehmer-E-Mail von der
    // des eingeloggten Users abweicht. Bei externen Adressen (kein Deloitte-
    // Postfach) ist die Zustimmung schriftlich einzuholen — das wird im
    // Nachweis vermerkt. v27.11: Member-Firm-Adressen (@deloitte.at, .com, …)
    // zählen als intern — nur echte Fremd-Domains laufen durch den Extern-Flow.
    const isProxyRegistration = (emailToUse || '').toLowerCase() !== (currentUserEmail || '').toLowerCase();
    const isExternalParticipant = isExternalEmail(emailToUse);
    const proxyConsentStr = (isProxyRegistration && opts?.proxyConsentConfirmed)
      ? `${isExternalParticipant ? 'Schriftliche ' : ''}Zustimmung der Person zur stellvertretenden Anmeldung bestätigt durch ${actorName} (${actorEmail}) am ${new Date().toLocaleString('de-DE')}`
      : '';

    // v19.9: Der Client weiß zuverlässig (aus dem geladenen Event-Objekt), ob
    // der/die Anmeldende Haupt- oder Co-Organizer dieses Events ist. Diese
    // Info an EventService durchreichen — sie hat dort Vorrang vor der fragilen
    // serverseitigen Organizer-Ableitung (SubsiteUrl/Note-Feld/pageContext),
    // die im Tenant gelegentlich legitime Organizer abgelehnt hat.
    const actorEmailLc = (currentUserEmail || '').toLowerCase();
    // v30.67 (Review): Parent-Fallback wie AdminPage.isOrganizerFor.
    // `_coOrganizers` steht nur auf der Klammer-Zeile; ein Sub-Event erbt
    // OrganizerEmail, aber keine Co-Organizer. Ohne den Fallback war die
    // Frist-Ausnahme für Co-Organizer (v30.67) auf Terminen wirkungslos —
    // „Teilnehmer hinzufügen" nach Fristablauf meldete je Termin „deadline".
    const isOrganizerOf = (ev: DeloitteEvent | undefined): boolean => !!ev && (
      (ev.organizerEmails || []).some(e => (e || '').toLowerCase() === actorEmailLc) ||
      (ev.coOrganizerEmails || []).some(e => (e || '').toLowerCase() === actorEmailLc)
    );
    const parentOfEvent = (event && event.parentEventId) ? events.find(p => p.id === event.parentEventId) : undefined;
    const actorIsEventOrganizer = isOrganizerOf(event) || isOrganizerOf(parentOfEvent);
    let success: boolean;
    let failReason: string | undefined;
    if (existing && existing.Status === 'Abgemeldet') {
      // v30.67: Gruppe mitgeben — bei geteilten Kapazitaeten landete eine
      // Re-Anmeldung sonst ohne StarterType und damit in keiner Gruppe.
      success = await eventService.reactivateRegistration(subsiteUrl, existing.Id, firstNameToUse, lastNameToUse, customData, status, fieldMap, actorName, actorEmail, proxyConsentStr, effectiveStarterType, preferredStarterType);
      if (!success) failReason = 'error';
    } else {
      const r = await eventService.registerForEvent(
        subsiteUrl, firstNameToUse, lastNameToUse, emailToUse, customData, status, fieldMap,
        effectiveStarterType, preferredStarterType, actorName, actorEmail, proxyConsentStr,
        actorIsEventOrganizer, !!opts?.actorAllowedAsAssistant
      );
      success = r.ok;
      failReason = r.reason;
    }
    // v30.67: Reservierten Platz zurückgeben, wenn der Insert scheitert.
    // reserveSeat schreibt bewusst VOR dem Insert (atomar per ETag-CAS); der
    // Erfolgspfad war gepflegt, der Fehlerpfad nicht. Bei 429 in der
    // Anmeldewelle, 403 auf einer frischen Subsite oder 400 wegen einer
    // fehlenden Custom-Field-Spalte blieb SeatsTaken erhöht — drei Versuche,
    // drei Phantom-Plätze, und ab der Kapazität ging jede echte Anmeldung auf
    // die Warteliste. Geheilt hat das nur der privilegierte Reconcile
    // (Admin-Boot, 1×/6 h) — ein reiner Teilnehmer-Pfad nie. Dasselbe Muster
    // wie der Rollback in switchSplitGroup.
    if (!success && reservedSeatField) {
      try { await eventService.adjustSeatCounterField(subsiteUrl, reservedSeatField, -1); }
      catch (err) { console.warn('[DEX] Sitzplatz-Rückgabe nach fehlgeschlagenem Insert fehlgeschlagen:', err); }
    }

    if (success && event) {
      // v22.20: Eine Reaktivierung vergibt bewusst eine FRISCHE Counter-ID —
      // die Person landet damit nummerisch hinter der Warteliste, obwohl sie
      // als Angemeldet zurückkommt. Direkt einen ID-Reorder anstoßen, damit
      // der Flow sie sofort korrekt einsortiert (statt erst bei der nächsten
      // Abmeldung). Best-effort, blockiert die Anmeldung nicht.
      if (existing && existing.Status === 'Abgemeldet' && status === 'Angemeldet') {
        eventService.queueIDReorder(eventId, event.eventNumber, subsiteUrl, event.title)
          .catch(err => console.warn('[DEX] queueIDReorder (reactivate) failed:', err));
      }
      // v9.0: Audit-Log (fire-and-forget)
      eventService.writeChangeLog({
        action: existing ? 'ParticipantReactivated' : 'ParticipantRegistered',
        targetType: 'Participant',
        targetId: emailToUse,
        targetName: nameToUse,
        eventId: eventId,
        eventTitle: event.title,
        details: { status, asActor: emailToUse !== currentUserEmail ? 'on-behalf-of' : 'self' },
      }).catch(() => { /* */ });
      // Warteliste-Position ermitteln
      let waitlistPosition = 0;
      if (status === 'Warteliste') {
        // v24.73: WaitlistTaken im (für alle lesbaren) Counter hochzählen — damit
        // die Live-Warteliste-Zahl auch für normale Teilnehmer stimmt (informativ,
        // nicht überbuchungs-relevant; privilegierter Reconcile heilt Drift).
        // v27.10: Bump ABWARTEN und die Position aus dem Counter lesen — die
        // frühere Listen-Zählung (getRegistrationCount) sieht für normale User
        // seit v26.87 nur die eigenen Zeilen und lieferte immer „1".
        try { await eventService.adjustWaitlistCounter(subsiteUrl, +1); } catch { /* best-effort */ }
        try {
          const stats = await eventService.getCounterStats(subsiteUrl, isSplitGroup);
          if (stats && stats.waitlist > 0) {
            waitlistPosition = stats.waitlist;
          } else {
            const counts = await eventService.getRegistrationCount(subsiteUrl);
            waitlistPosition = counts.waitlist;
          }
        } catch { /* Position nicht ermittelbar */ }
      }

      // Dual-Write: DEX_Participants aktualisieren
      if (event.eventNumber) {
        eventService.upsertParticipant(
          firstNameToUse, lastNameToUse, emailToUse, event.eventNumber, status
        ).catch(err => console.warn('[DEX] upsertParticipant failed:', err));
      }
      // E-Mail in Queue eintragen (SharePoint-Template, Fallback auf Code-Template)
      const templateType = status === 'Warteliste' ? 'Warteliste' : 'Anmeldung';
      const lang = event.emailLanguage || 'EN';
      const posText = waitlistPosition > 0 ? String(waitlistPosition) : '';
      // {{Name}} in E-Mail-Anreden: nur Vorname (firstNameToUse ist bei Self-Reg
      // aus dem displayName gesplittet, bei "Für andere registrieren" explizit gesetzt).
      const vars = { Name: firstNameToUse, EventTitle: event.title, Organizer: formatOrganizerList(event.organizers, lang), AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`, WaitlistPosition: posText };
      // v26.47: Externe Dritte (kein Deloitte-Postfach), die stellvertretend
      // angemeldet wurden — der Mail-Flow kann externe Adressen NICHT erreichen.
      // Deshalb: (1) Registrierung als „Datenschutzrückmeldung offen" markieren
      // (ConsentReview='Pending'), (2) INSTRUKTIONS-Mail an die ANMELDENDE
      // Person (intern, zustellbar) mit den 3 Schritten — den fertigen
      // Einladungs-Entwurf (.eml) lädt sie in der Teilnehmerliste herunter und
      // verschickt ihn aus dem eigenen Postfach.
      const isExternalInvite = status === 'Angemeldet' && isProxyRegistration && isExternalParticipant;
      // v26.73: Download-Deeplink für den .eml-Entwurf. Der Anhang darf per
      // Deloitte-Mail-Regel NICHT direkt mitgeschickt werden (NDR: „…cannot send
      // email attachments"). Statt dessen: den fertigen Entwurf als Attachment an
      // der Teilnehmer-Zeile ablegen und in der Instruktions-Mail einen Deeplink
      // setzen, der die Datei beim Klick in der App herunterlädt.
      let inviteDownloadUrl = '';
      if (isExternalInvite) {
        try { await eventService.markConsentPendingByEmail(subsiteUrl, emailToUse); }
        catch (err) { console.warn('[DEX] markConsentPendingByEmail failed:', err); }
        try {
          const inv = externalInvitationEmail(
            nameToUse, event.title, currentUserName || '',
            lang.toUpperCase() === 'DE',
            { startDate: event.startDate, endDate: event.endDate, location: event.location }
          );
          const eml = buildUnsentEmlDraft({
            to: [emailToUse],
            cc: ['no_reply.events@deloitte.de', ...Array.from(new Set([...(event.organizerEmails || []), ...(event.coOrganizerEmails || [])].filter(Boolean)))],
            subject: inv.subject,
            html: inv.body,
          });
          const invItemId = await eventService.storeInviteEmlByEmail(subsiteUrl, emailToUse, eml).catch(() => 0);
          if (invItemId > 0) {
            inviteDownloadUrl = buildHashDeepLink(`${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`, { action: 'downloadinvite', event: eventId, item: invItemId, email: emailToUse, name: nameToUse });
          }
        } catch (emlErr) { console.warn('[DEX] invite .eml store failed:', emlErr); }
      }
      let emailData: { subject: string; body: string };
      // v26.73: Die eigentliche Einladung (.eml) wird NICHT an die Mail gehängt
      // (Mail-Regel blockt Anhänge). Sie liegt als Attachment an der Teilnehmer-
      // Zeile; der Anmelder holt sie per Deeplink-Button in der App und leitet sie
      // aus dem eigenen Postfach weiter (die App/der Flow sendet NIE an Externe).
      const spTemplateRaw = isExternalInvite ? null : await eventService.getEmailTemplate(templateType, lang).catch(() => null);
      const spTemplate = applyEventTemplateOverride(spTemplateRaw, event.emailTemplateOverrides, templateType);
      if (isExternalInvite) {
        const orgCenterUrl = buildHashDeepLink(`${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`, { action: 'admin', event: eventId });
        const registrantFirst = (currentUserName || '').split(' ')[0] || currentUserName;
        emailData = externalInviteInstructionEmail(
          registrantFirst, nameToUse, emailToUse, event.title,
          lang.toUpperCase() === 'DE', orgCenterUrl, inviteDownloadUrl
        );
      } else if (spTemplate) {
        emailData = buildEmailFromTemplate(spTemplate, vars);
      } else {
        emailData = status === 'Warteliste'
          ? waitlistEmail(firstNameToUse, event.title, waitlistPosition)
          : registrationEmail(firstNameToUse, event.title);
      }
      // v15.25: Im subEventsOnlyMode wird das Hauptevent nur als
      // „Schatten-Registrierung" angelegt (Daten-Zeile für Parent-CFs).
      // Der User nimmt nicht am Hauptevent teil und soll dafür KEINE
      // Bestätigungs-Mail und KEINEN Outlook-Termin bekommen — die
      // tatsächlichen Teilnahme-Mails kommen pro Sub-Event.
      // v30.61: Gebündelte Kommunikation kehrt genau diese Unterdrückung um.
      //
      // Bisher galt: Die Klammer ist im subEventsOnlyMode nur eine Datenzeile,
      // also schweigt sie — jeder Termin verschickt selbst. Steht der Modus auf
      // gebündelt, ist die Klammer der EINZIGE Absender, und die Termine werden
      // still angelegt (die Anmeldeseite setzt dafür suppressMail/-Outlook).
      // Mail, Kalender und QR sind getrennt schaltbar (Nutzer-Entscheidung
      // 01.09.2026), deshalb drei Ableitungen statt einer.
      const bundledMode = bundledCommOf(event);
      const suppressParentNotifications = !!event.subEventsOnlyMode && !bundledMode.mail;
      const suppressParentOutlook = !!event.subEventsOnlyMode && !bundledMode.outlook;
      const suppressParentQr = !!event.subEventsOnlyMode && !bundledMode.qr;
      // …und die Gegenrichtung: Steht die Klammer auf „ein QR fürs Gesamt-Event",
      // darf der Termin KEINEN eigenen Code schicken — sonst bekäme die Person
      // beides, und am Einlass gäbe es zwei gültige Codes für eine Person.
      // Dasselbe für Mail und Kalender: Die Anmeldeseite unterdrückt sie zwar
      // schon beim Aufruf, aber sie ist nicht der einzige Weg, auf dem eine
      // Termin-Anmeldung entsteht (Organizer-Modal, Assistenz, Massenimport).
      // Diese Prüfung hier greift für alle.
      const parentOfThis = event.parentEventId ? events.find(e => e.id === event.parentEventId) : undefined;
      const parentBundled = bundledCommOf(parentOfThis);
      const childSuppressedByParent = !!event.parentEventId && !!parentOfThis && !!parentOfThis.subEventsOnlyMode;
      // v19.21: disableRegistrationEmail = nur die Anmelde-Bestätigung
      // unterdrücken (granulares Sub-Häkchen unter dem Master „E-Mails").
      if (!event.disableEmails && !event.disableRegistrationEmail && !suppressParentNotifications && !opts?.suppressMail
        && !(childSuppressedByParent && parentBundled.mail)) {
        // v8.5: Organizer-Mitlese-Modus auswerten. Bei 'always' immer,
        // bei 'fromDate' nur wenn das konfigurierte Datum bereits erreicht
        // ist, bei 'never'/undefined gar nicht.
        // v28.28: Die Kopie geht jetzt auf CC statt BCC — der Organizer soll
        // für die Teilnehmer:innen SICHTBAR mit im Verteiler stehen (bewusste
        // Produktentscheidung: Transparenz, wer die Anmeldung betreut, und
        // „Allen antworten" landet direkt beim richtigen Ansprechpartner).
        let orgCopyCc = '';
        const mode = event.notifyOrgRegisterMode || 'never';
        if (mode === 'always' || (mode === 'fromDate' && event.notifyOrgRegisterFromDate && new Date() >= new Date(event.notifyOrgRegisterFromDate))) {
          const orgEmails = (event.organizerEmails || []).filter(Boolean);
          if (orgEmails.length > 0) orgCopyCc = orgEmails.join(';');
        }
        // v9.22: Externe Mail-Adresse erkennen — kein Deloitte-Postfach.
        // v18.74: Bei externen Empfängern wird die Bestätigungsmail jetzt
        // DIREKT an die externe Person versendet (vorher an den Organizer
        // umgeleitet) — mit dem Organizer auf CC (Nachweis/Kopie). Ein
        // Outlook-Kalendereintrag wird für externe Adressen weiterhin NICHT
        // versendet (Microsoft blockt das ohne Federation, s.u.
        // skipOutlookForExternal). v27.11: Member-Firm-Adressen (@deloitte.at,
        // @deloitte.com, …) zählen jetzt als intern.
        const isExternalRecipient = isExternalEmail(emailToUse);
        // v26.47: Bei externer Einladung geht die INSTRUKTIONS-Mail an die
        // ANMELDENDE Person (intern, zustellbar) — an externe Adressen kann der
        // Mail-Flow nicht zustellen. Die eigentliche Einladung verschickt die
        // anmeldende Person als .eml-Entwurf aus dem eigenen Postfach
        // (Download in der Teilnehmerliste).
        const finalRecipient = isExternalInvite ? currentUserEmail : emailToUse;
        const finalSubject = emailData.subject;
        let finalBody = emailData.body;
        // v30.61: Bei gebündelter Kommunikation trägt DIESE eine Mail alles,
        // wofür die Person angemeldet ist. Ohne die Liste wäre die Bündelung
        // ein Rückschritt: Zehn Einzelmails sagen wenigstens, worum es geht.
        // Die Liste kommt vom Aufrufer (nur er weiß, was gerade gebucht wurde)
        // und wird direkt in den Fließtext gesetzt.
        if (bundledMode.mail && opts?.bundledItems && opts.bundledItems.length > 0) {
          const isDeMail = (lang || 'EN').toUpperCase() === 'DE';
          const block = `<p style="margin:18px 0 0;font-weight:700;">`
            + bundledItemsHeading(opts.bundledItems.length, isDeMail, event.childEventTermPlural)
            + `</p>`
            + bundledItemsTableHtml(opts.bundledItems, isDeMail);
          finalBody = injectIntoEmailContent(finalBody, block);
        }
        const finalRecipientName = isExternalInvite ? (currentUserName || currentUserEmail) : nameToUse;
        // CC-Adressen, die zusätzlich zu den Feld-CCs gelten (Organizer bei
        // externer Anmeldung).
        let externalCcExtra = '';
        if (isExternalRecipient) {
          // v26.33: ALLE Organizer (inkl. Co-Organizer) auf CC.
          const allOrganizers = [...(event.organizerEmails || []), ...(event.coOrganizerEmails || [])].filter(Boolean);
          if (isExternalInvite) {
            // v26.71: Instruktions-Mail geht an den Anmelder (To); die
            // Organisator:innen kommen zusätzlich auf CC (Kopie/Nachweis). Es
            // steht KEINE externe Adresse in To/CC — die App sendet nie an
            // Externe (dedupe/To-Ausschluss passiert weiter unten in ccMerged).
            externalCcExtra = allOrganizers.join(';');
          } else {
            externalCcExtra = allOrganizers.join(';');
            // Hinweis-Box VOR dem Original-Body — adressiert an die externe Person.
            const externalHint = `<div style="margin:0 0 16px;padding:12px 16px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;line-height:1.55;color:#7a4a00;">`
              + `<strong>Externe Anmeldung — kein automatischer Kalendereintrag.</strong><br>`
              + `Diese Anmeldebestätigung gehört zu <strong>${nameToUse}</strong> (<strong>${emailToUse}</strong>). Da es sich um eine externe Adresse (kein Deloitte-Postfach) handelt, wird <strong>kein Outlook-Kalendereintrag</strong> versendet — bitte trage dir den Termin manuell in deinen Kalender ein. `
              + `Der Organizer ist zur Bestätigung auf Kopie (CC).`
              + `</div>`;
            // Body kommt schon als komplett-gewickeltes HTML (Deloitte-Template).
            // Wir injecten den Hinweis direkt nach dem opening-<body>-Tag.
            finalBody = injectIntoEmailContent(finalBody, externalHint);
          }
        }
        // v26.37: Standard-Tipp in der Anmeldebestätigung interner User —
        // DEX auf dem Handy am besten über die offizielle Microsoft-SharePoint-
        // App öffnen (der mobile Browser lädt die Seite teils nicht zuverlässig,
        // v.a. Android). Nicht für externe Empfänger/Einladungen (nutzen DEX nicht)
        // und nur bei der eigentlichen Anmeldebestätigung ('Angemeldet').
        if (status === 'Angemeldet' && !isExternalInvite && !isExternalRecipient) {
          const isDeMail = (lang || 'EN').toUpperCase() === 'DE';
          const mobileAppTip = `<div style="margin:0 0 16px;padding:12px 16px;background:#f1f7e8;border:1px solid #86bc25;border-radius:8px;font-size:13px;line-height:1.55;color:#3d5a1a;">`
            + (isDeMail
              ? `<strong>💡 Tipp: DEX auf dem Handy nutzen.</strong><br>Wenn du DEX auf deinem Smartphone öffnen möchtest, installiere vorher die offizielle <strong>Microsoft-SharePoint-App</strong> aus dem App Store (iPhone/iPad) bzw. Google Play Store (Android) und öffne DEX darüber. Im normalen Handy-Browser lädt die Seite teilweise nicht zuverlässig.`
              : `<strong>💡 Tip: Using DEX on your phone.</strong><br>If you'd like to open DEX on your smartphone, first install the official <strong>Microsoft SharePoint app</strong> from the App Store (iPhone/iPad) or Google Play Store (Android) and open DEX from there. In the regular mobile browser the page sometimes does not load reliably.`)
            + `</div>`;
          // v26.39: Tipp GANZ UNTEN einfügen — direkt unter der „Made with DEX
          // App"-Zeile (als eigene Tabellen-Row), statt oben über der Anrede.
          const tipRow = `<tr><td style="padding:4px 30px 20px 30px;">${mobileAppTip}</td></tr>`;
          const madeWithRe = /(Made with DEX App<\/a>\s*<\/td>\s*<\/tr>)/i;
          if (madeWithRe.test(finalBody)) {
            finalBody = finalBody.replace(madeWithRe, `$1\n${tipRow}`);
          } else if (/<\/body>/i.test(finalBody)) {
            // Kein „Made with DEX App" in der Vorlage → als Fallback vor </body>.
            finalBody = finalBody.replace(/<\/body>/i, `${tipRow}</body>`);
          } else {
            finalBody = injectIntoEmailContent(finalBody, mobileAppTip);
          }
        }
        // v26.41: Wurde zu diesem Event bereits eine Rundmail versendet (Einladung/
        // Massenmail/Ankündigung), bekommen Spät-Anmelder:innen (und via denselben
        // Pfad Nachrücker) in ihrer Anmeldebestätigung den Hinweis, dass sie die
        // bisherige Kommunikation in der App unter „Meine Events" nachlesen können.
        if (status === 'Angemeldet' && !isExternalInvite && !isExternalRecipient) {
          let priorComms = false;
          // v29.11: Die Einladung zählt nicht mit. Sie ist der Weg, über den die
          // meisten überhaupt hier gelandet sind — auf sie zu verweisen sagt
          // nichts. Der Hinweis erscheint erst, wenn es darüber hinaus eine
          // Rundmail gab (Ankündigung, Update, Massenmail).
          try { priorComms = await eventService.hasEventComms(eventId, ['Einladung']); } catch { priorComms = false; }
          if (priorComms) {
            const isDeComm = (lang || 'EN').toUpperCase() === 'DE';
            const commsBox = `<div style="margin:0 0 16px;padding:12px 16px;background:#eef4fb;border:1px solid #0076a8;border-radius:8px;font-size:13px;line-height:1.55;color:#0b4a6f;">`
              + (isDeComm
                ? `<strong>Bereits versendete Infos zu diesem Event.</strong><br>Zu diesem Event wurde vorab schon per Mail kommuniziert (z.&nbsp;B. eine Einladung oder Ankündigung). Du findest diese bisherige Kommunikation jederzeit in der DEX App unter <strong>&bdquo;Meine Events&ldquo;</strong> beim Event — so bist du auf dem gleichen Stand.`
                : `<strong>Earlier updates for this event.</strong><br>Some information about this event was already sent out by email (e.g. an invitation or announcement). You can read this previous communication any time in the DEX App under <strong>&bdquo;My Events&ldquo;</strong> on the event — so you're fully up to date.`)
              + `</div>`;
            // v26.69: Hinweis ans Ende verschoben (vorher oben über der Anrede) —
            // als eigene Tabellen-Row direkt UNTER „Made with DEX App" und damit
            // ÜBER der grünen Mobil-Tipp-Box. Die grüne Box wurde oben bereits nach
            // „Made with DEX App" eingefügt; ein zweiter Insert an derselben Stelle
            // landet DAVOR (blau über grün).
            const commsRow = `<tr><td style="padding:4px 30px 4px 30px;">${commsBox}</td></tr>`;
            const madeWithReComms = /(Made with DEX App<\/a>\s*<\/td>\s*<\/tr>)/i;
            if (madeWithReComms.test(finalBody)) {
              finalBody = finalBody.replace(madeWithReComms, `$1\n${commsRow}`);
            } else if (/<\/body>/i.test(finalBody)) {
              finalBody = finalBody.replace(/<\/body>/i, `${commsRow}</body>`);
            } else {
              finalBody = injectIntoEmailContent(finalBody, commsBox);
            }
          }
        }
        // v18.41: People-Picker-Felder mit „CC bei Mail" → ausgewählte
        // Person(en) auf CC der An-/Warteliste-Mail (NICHT im Outlook-Termin).
        // v18.53: opts.extraCc (übergreifende CC aus dem Hauptformular im
        // subEventsOnlyMode) mit den event-eigenen CC-Feldern mergen + deduppen.
        const ccOwn = collectCcEmailsFromFields(event.eventSpecificFields, customData, emailToUse);
        const ccMerged = (() => {
          const seen = new Set<string>();
          const out: string[] = [];
          // v26.45: ALLE An-Empfänger ausschließen (bei der externen Einladung
          // steht auch der/die Anmelder:in im An-Feld — nicht zusätzlich auf CC).
          const toSet = new Set(finalRecipient.split(';').map(s => s.trim().toLowerCase()).filter(Boolean));
          // v18.74: externalCcExtra (Organizer bei externer Anmeldung) mitmergen.
          // v28.28: orgCopyCc = Organizer-Mitlese-Kopie (früher BCC).
          for (const part of [ccOwn, opts?.extraCc || '', externalCcExtra, orgCopyCc]) {
            for (const e of part.split(';').map(s => s.trim()).filter(Boolean)) {
              const lc = e.toLowerCase();
              if (!toSet.has(lc) && !seen.has(lc)) { seen.add(lc); out.push(e); }
            }
          }
          return out.join(';');
        })();
        const ccFromFields = ccMerged || undefined;
        eventService.queueEmail(
          withParentTitleSubject(finalSubject, calDayParentOf(event)), finalRecipient, finalRecipientName, finalBody,
          templateType, event.title, eventId, ccFromFields, undefined,
          undefined,
          // v26.71: KEIN Anhang mehr — die Deloitte-Mail-Flow-Regel blockt
          // Power-Automate-Mails mit Anhang (NDR). Der .eml-Entwurf wird im
          // Organizer Center heruntergeladen und aus dem eigenen Postfach
          // weitergeleitet.
          undefined
        ).catch(err => console.warn('[DEX] queueEmail failed:', err));
      }

      // v9.15/v20.7/v20.10/v21: Auto-Send QR-Code — unabhängig vom
      // Bestätigungs-Mail-Block (greift auch bei stiller Massen-Anmeldung),
      // aber NUR in der „QR-Phase" des Events. v21 BUG-FIX: vorher feuerte
      // der QR bei JEDER Anmeldung an JEDEM Event — auch Monate im Voraus.
      // QR-Phase = der Organizer hat den QR-Massen-Versand bereits gestartet
      // (setzt AutoSendQRCode=true am Event, siehe AdminPage) ODER die
      // Anmeldefrist ist vorbei (Nachzügler). Master „DisableEmails" und die
      // Schatten-Registrierung (subEventsOnly) heben ihn weiterhin auf; nur
      // für Status='Angemeldet' (Wartelistler sind noch nicht confirmed).
      // v21 HOTFIX (User): QR-Phase startet AUSSCHLIESSLICH mit dem ersten
      // manuellen QR-Massen-Versand des Organizers (der setzt
      // AutoSendQRCode=true am Event). Kein Deadline-Trigger — vor dem ersten
      // Versand bekommt KEINE Anmeldung automatisch einen QR-Code.
      const qrPhaseActive = event.autoSendQRCode === true;
      if (qrPhaseActive && status === 'Angemeldet' && !event.disableEmails && !suppressParentQr
        && !(childSuppressedByParent && parentBundled.qr)) {
        // v27.11: Member-Firm-Adressen zählen als intern — QR-Mail direkt.
        const isExternalRecipientQr = isExternalEmail(emailToUse);
        (async (): Promise<void> => {
            try {
              const qrData = `DEX|${event.eventNumber}|${emailToUse}`;
              let qrImageHtml = `<p style="font-family:monospace;font-size:1.2rem;background:#f5f5f5;padding:12px;border-radius:8px;text-align:center;">${qrData}</p>`;
              try {
                // v30.36: gemeinsamer Erzeuger mit dem Massen-Versand
                // (utils/qrWithMark) — hohe Fehlerkorrektur UND Deloitte-D. in
                // der Mitte. Vorher lief dieser Pfad eigenstaendig und hatte
                // beides nicht; der Unterschied faellt nur als schlechteres
                // Scan-Ergebnis auf und wird nie zurueckverfolgt. Die
                // Bibliothek bleibt dort lazy geladen (Boot-Pfad).
                const { buildParticipantQrDataUrl } = await import('../utils/qrWithMark');
                const qrDataUrl = await buildParticipantQrDataUrl(qrData, 300);
                if (qrDataUrl) {
                  qrImageHtml = `<img src="${qrDataUrl}" alt="QR-Code" style="width:300px;max-width:100%;height:auto;" />`;
                }
              } catch (qrErr) { console.warn('[DEX] QR-Erzeugung fehlgeschlagen — Text-Fallback:', qrErr); }
              // v22.18: pro-Event angepasster QR-Mail-Text (Override-Key
              // 'QRCode' im EmailTemplateOverrides-JSON) — gilt damit auch
              // für den Auto-Versand, nicht nur den manuellen Versand.
              let qrOverride;
              try {
                const ovAll = JSON.parse(event.emailTemplateOverrides || '{}');
                const ov = ovAll && ovAll['QRCode'];
                // v30.52: `headerImage` zählt jetzt AUCH als Override — sonst
                // ginge die gespeicherte Kopf-Bild-Einstellung beim
                // automatischen Versand verloren, wenn die Texte auf dem
                // Standard stehen.
                if (ov && (ov.subject || ov.heading || ov.subheading || ov.bodyHtml || ov.headerImage)) qrOverride = ov;
              } catch { /* kein Override */ }
              // v30.52: Hat der Organizer das Event-Foto als Kopf-Bild
              // gewählt, wird es hier aufgelöst und fest eingebacken — sonst
              // bleibt {{ORB_URL}} stehen und der Flow setzt das Standardbild.
              // Der Abruf ist gecacht, kostet pro Sitzung also einmal.
              let qrHeroPhoto = '';
              if (qrOverride && qrOverride.headerImage && qrOverride.headerImage.hero === 'event' && event.imageUrl) {
                try {
                  const b64 = await getCachedImage(event.imageUrl);
                  if (b64 && b64.indexOf('data:') === 0) qrHeroPhoto = b64;
                } catch { /* Foto nicht ladbar → Standardbild */ }
              }
              // v30.67: Die Registrierung VOR dem Mailaufbau laden und die
              // Teilnehmer-ID durchreichen. Der 7. Parameter stand hart auf
              // undefined — die Auto-QR-Mail der Nachzügler kam ohne die
              // ID-Zeile unter dem Code und ohne den Hinweis „Nummer am Einlass
              // nennen". Genau die Personen, bei denen der Kamera-Scan
              // (Android/SharePoint-App) scheitert, hatten keine Ersatznummer;
              // Massenversand und Test hatten sie. Die Zeile wurde ohnehin
              // gleich darauf für setQRSentStatus geladen — jetzt nur einmal.
              const myReg = event.subsiteUrl ? await eventService.getMyRegistration(event.subsiteUrl, emailToUse) : null;
              const qrMail = qrCodeEmail(firstNameToUse, event.title, qrImageHtml, lang, nameToUse, qrOverride, myReg?.TeilnehmerID, qrHeroPhoto);
              // v9.22: Auto-Send-QR für externe Empfänger ebenfalls an den
              // Organizer umleiten (mit klarem Subject-Präfix), nicht an den
              // externen Mail-Empfänger.
              if (isExternalRecipientQr) {
                const orgEmails = (event.organizerEmails || []).filter(Boolean);
                const orgRecipient = orgEmails.length > 0 ? orgEmails.join(';') : currentUserEmail;
                const orgSubject = `[Externer Teilnehmer] QR-Code für ${nameToUse} — ${event.title}`;
                // Hinweis-Box vor dem QR-Code-Body — analog zur Bestätigungsmail.
                const qrExternalHint = `<div style="margin:0 0 16px;padding:12px 16px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;line-height:1.55;color:#7a4a00;">`
                  + `<strong>QR-Code für externen Teilnehmer.</strong><br>`
                  + `Eigentlich für <strong>${emailToUse}</strong> (${nameToUse}). Da externe Adressen keinen Auto-Versand bekommen, landet der QR-Code bei dir — drucke ihn aus oder leite die Mail intern an den Empfänger weiter (Datenschutzrichtlinien Deloitte Deutschland beachten).`
                  + `</div>`;
                const qrBody = injectIntoEmailContent(qrMail.body, qrExternalHint);
                await eventService.queueEmail(
                  orgSubject, orgRecipient, 'Organizer', qrBody,
                  'QRCode', event.title, eventId
                );
              } else {
                await eventService.queueEmail(
                  qrMail.subject, emailToUse, nameToUse, qrMail.body,
                  'QRCode', event.title, eventId
                );
              }
              // Status auf 'QR versendet' setzen, damit der Admin-Center-View
              // sofort zeigt, dass der QR-Code raus ist (analog zum
              // manuellen Massen-QR-Versand).
              if (event.subsiteUrl && myReg && myReg.Id) {
                await eventService.setQRSentStatus(event.subsiteUrl, myReg.Id);
              }
            } catch (err) { console.warn('[DEX] auto-send QR failed:', err); }
          })().catch(err => console.warn('[DEX] auto-send QR outer failed:', err));
        }
      // Roommate-Benachrichtigung: nur Custom-Fields vom Typ 'roommate'
      // durchsuchen (seit v7.17 eigener Feldtyp; vorher waren es alle 'user'-
      // Felder, was bei "Assistent"-, "Mentor"- etc. Pickern zu ungewollten
      // Roommate-Mails führte). Für jede ausgewählte E-Mail eine Roommate-
      // Anfrage-Mail im Deloitte-Template queuen.
      if (!event.disableEmails) {
        for (const f of event.eventSpecificFields) {
          if (f.type !== 'roommate') continue;
          // v26.60 BUG-FIX (Organizer-Feedback): Diese Anfrage-Mail lief bisher
          // IMMER — der CC-Schalter des Felds betraf nur die Kopie der
          // An-/Abmelde-Mail. Jetzt schaltet der neue Wizard-Schalter
          // „Ausgewählte Person automatisch benachrichtigen" sie ab.
          if (f.notifyRoommate === false) continue;
          const v = customData[f.id];
          if (!v) continue;
          const m = v.match(/<([^>]+@[^>]+)>/);
          const partnerEmail = m ? m[1].trim() : '';
          if (!partnerEmail || partnerEmail.toLowerCase() === emailToUse.toLowerCase()) continue;
          const partnerName = v.replace(/<[^>]*>/, '').trim() || partnerEmail;
          const partnerFirstName = partnerName.includes(',')
            ? (partnerName.split(',')[1] || '').trim()
            : (partnerName.split(/\s+/)[0] || '');
          const registrantFullName = `${firstNameToUse} ${lastNameToUse}`.trim() || nameToUse;
          const isDe = (lang || 'EN').toUpperCase() === 'DE';
          // v13.0: Roommate-Mail aus TemplateType=RoommateRequest.
          const tpl = await eventService.getEmailTemplate('RoommateRequest', lang).catch(() => null);
          const vars: Record<string, string> = {
            Name: partnerFirstName || partnerName,
            RegistrantName: registrantFullName,
            EventTitle: event.title,
            AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
          };
          let mail: { subject: string; body: string };
          if (tpl) {
            mail = buildEmailFromTemplate(tpl, vars);
          } else {
            const inner = isDe
              ? `<p>Hallo ${partnerFirstName || partnerName},</p><p><strong>${registrantFullName}</strong> hat dich als <strong>Zimmerpartner</strong> für das Event <strong>${event.title}</strong> angegeben.</p>`
              : `<p>Hello ${partnerFirstName || partnerName},</p><p><strong>${registrantFullName}</strong> has selected you as their <strong>roommate</strong> for the event <strong>${event.title}</strong>.</p>`;
            mail = {
              subject: isDe ? `Zimmerpartner-Anfrage: ${event.title}` : `Roommate request: ${event.title}`,
              body: wrapTemplate('#86bc25', isDe ? 'Zimmerpartner-Anfrage' : 'Roommate request', event.title, inner),
            };
          }
          eventService.queueEmail(
            mail.subject, partnerEmail, partnerName, mail.body, 'RoommateRequest', event.title, eventId
          ).catch(err => console.warn('[DEX] roommate queueEmail failed:', err));
        }
      }
      // Outlook-Termin-Einladung in Queue eintragen
      // v15.16: Für externe Empfänger (kein Deloitte-Postfach) keinen
      // Outlook-Termin queuen — Microsoft blockt das Versenden an externe
      // Adressen ohne Federation, deshalb ist der Eintrag immer ein
      // Bounce. Der Organizer bekommt stattdessen die Bestätigungsmail
      // mit Betreff „Weiterleitung notwendig" (s.o.) und kann darüber
      // den externen Teilnehmer informieren. v27.11: Member-Firm-Postfächer
      // zählen als intern und bekommen den Termin.
      const skipOutlookForExternal = isExternalEmail(emailToUse);
      // v15.25: Schatten-Parent-Registrierung im subEventsOnlyMode bekommt
      // keinen Outlook-Termin (s.o. — der User „nimmt teil" an Sub-Events,
      // nicht am Parent).
      if (status !== 'Warteliste' && !event.disableOutlook && !skipOutlookForExternal && !suppressParentOutlook && !opts?.suppressOutlook
        && !(childSuppressedByParent && parentBundled.outlook)) {
        eventService.queueOutlookEvent(
          emailToUse, eventId, event.title, 'Einladen'
        ).catch(err => console.warn('[DEX] queueOutlookEvent failed:', err));
      }
      // v11.53: KPI-Counter sofort hochzählen, damit der nächste Boot-
      // Loader die neue Zahl ohne Verzögerung zeigt. Nur für 'Angemeldet'-
      // Status (Warteliste zählt nicht in 'Teilnehmer').
      if (status === 'Angemeldet') {
        eventService.bumpKpiParticipants(1).catch(() => { /* best-effort */ });
      }
      // v30.9: s. opts.skipReload — Schleifen laden EINMAL am Ende neu.
      if (!opts?.skipReload) await loadEvents();
    }
    // v30.42 → v30.68: Die Klammer-Zeile wird zentral sichergestellt — seit
    // v30.68 VOR dem Termin (s. oben). Scheiterte sie dort, hier der zweite
    // Versuch nach dem Termin (der Termin ist geschrieben, die Person ist
    // angemeldet); scheitert auch der, holt der Nachzug-Merker die Zeile beim
    // nächsten App-Start nach (utils/shadowHeal).
    if (success && umbrellaPending && event && event.parentEventId) {
      const parentEv2 = events.find(e => e.id === event.parentEventId);
      if (parentEv2) {
        const guardKey2 = `${parentEv2.id}|${(emailToUse || '').toLowerCase().trim()}`;
        let healed = false;
        try {
          const again = await registerForEvent(
            parentEv2.id, {}, firstNameToUse, lastNameToUse, emailToUse, undefined,
            {
              suppressMail: true, suppressOutlook: true, skipReload: true, skipShadowParent: true,
              ...(opts?.proxyConsentConfirmed
                ? { proxyConsentConfirmed: true, actorAllowedAsAssistant: !!opts?.actorAllowedAsAssistant }
                : {}),
            }
          );
          healed = again.ok || again.reason === 'already-registered';
        } catch (err) { console.warn('[DEX] Klammer-Zeile auch im zweiten Versuch nicht:', err); }
        if (healed) shadowEnsuredRef.current.set(guardKey2, Date.now());
        else {
          try {
            addPendingShadowParent({
              eventId: parentEv2.id, customData: {},
              firstName: firstNameToUse || '', lastName: lastNameToUse || '',
              email: emailToUse || '', proxy: !!opts?.proxyConsentConfirmed, ts: Date.now(),
            });
          } catch { /* localStorage gesperrt → bleibt für das Organizer-Panel */ }
        }
      }
    }
    // v18.67: echten Status zurückgeben (Angemeldet/Warteliste), damit die
    // RegistrationPage das Ergebnis-Modal nicht mehr aus der gecachten
    // currentParticipants-Schätzung (isFull) ableiten muss — die war nach
    // Cancel/Re-Register veraltet und zeigte fälschlich "Warteliste".
    return { ok: success, status, reason: success ? undefined : failReason };
  }

  /**
   * v11.82: Team-Anmeldung — eine Person meldet sich + N-1 Mitglieder
   * gleichzeitig an. N Plätze werden atomar reserviert (per `reserveSeat`
   * mit count=N). Sind nicht genug Plätze frei, geht das gesamte Team
   * auf die Warteliste — kein Teil-Anmelden eines vollen Events.
   *
   * Jedes Mitglied bekommt einen eigenen Eintrag in der Subsite-Teilnehmer-
   * liste mit identischer `TeamId`. Genau ein Eintrag (der Lead, also der
   * Submitter) ist `TeamLead=true`. Jeder Mitglied bekommt eine eigene
   * Bestätigungs-Mail (mit Hinweis dass er als Teil eines Teams angemeldet
   * wurde) und einen eigenen Outlook-Termin (sofern aktiviert).
   *
   * Die Member-Einträge bekommen leere Custom-Field-Antworten — nur der
   * Lead beantwortet event-spezifische Fragen. Pflicht-Custom-Fields sollten
   * organisatorisch nicht mit Team-Anmeldung kombiniert werden; die App
   * setzt das nicht hart durch, der Wizard sollte den Organizer im Manual
   * darauf hinweisen.
   */
  async function registerTeam(
    eventId: string,
    leadData: {
      firstName: string;
      lastName: string;
      email: string;
      salutation?: string;
      customData: Record<string, string>;
      preferredStarterType?: string;
    },
    members: Array<{ email: string; displayName: string; customData?: Record<string, string> }>,
    teamName: string | undefined
  ): Promise<{ ok: boolean; teamId?: string; status?: 'Angemeldet' | 'Warteliste'; reason?: string }> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, reason: 'event-not-found' };
    const event = events.find(e => e.id === eventId);
    if (!event) return { ok: false, reason: 'event-not-found' };

    // Doppel-Anmelde-Prüfung: weder der Lead noch ein Member darf bereits
    // (aktiv) angemeldet sein. v11.83: konsolidiert auf den zentralen Helper
    // `isUserAlreadyOnEvent`, der genau die blockierenden Status-Werte
    // berücksichtigt (Angemeldet/QR versendet/Eingecheckt/Warteliste). Pfad
    // ist nicht performance-kritisch — sequentiell ist OK bei N ≤ 20.
    // v23.2: Der Lead ist (fast immer) der eingeloggte User selbst — und für
    // die EIGENE Anmeldung haben wir einen verlässlichen Self-Read
    // (`getMyRegistration` liest die eigene Zeile auch unter Item-Level-
    // Security). Das ist der harte Riegel gegen die Team-Doppel-Anmeldung:
    // `isUserAlreadyOnEvent` fällt bei transienten SP-Fehlern (Throttling
    // während einer Anmelde-Welle) auf „nicht blockieren" zurück — der
    // Self-Read tut das nicht. Wenn die eigene Anmeldung aktiv ist, hart
    // abweisen, BEVOR ein Sitzplatz reserviert wird.
    if (leadData.email.trim().toLowerCase() === currentUserEmail.trim().toLowerCase()) {
      try {
        const myReg = await getMyRegistration(eventId);
        if (myReg && myReg.Status !== 'Abgemeldet') {
          return { ok: false, reason: `already-registered:${leadData.email}` };
        }
      } catch { /* Self-Read fehlgeschlagen → fällt unten auf isUserAlreadyOnEvent zurück */ }
    }

    const allEmails = [leadData.email, ...members.map(m => m.email)];
    for (const em of allEmails) {
      // v30.67: null = nicht prüfbar → ablehnen statt durchwinken (s. probeAlreadyActive).
      const blocked = await probeAlreadyActive(subsiteUrl, em);
      if (blocked === null) return { ok: false, reason: 'dup-check-failed' };
      if (blocked) {
        return { ok: false, reason: `already-registered:${em}` };
      }
    }

    // TeamId generieren — bevorzugt crypto.randomUUID, sonst Fallback.
    let teamId: string;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const c: any = (typeof crypto !== 'undefined') ? crypto : null;
    if (c && typeof c.randomUUID === 'function') {
      teamId = c.randomUUID();
    } else {
      teamId = `team-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
    }

    const teamCount = 1 + members.length;
    let status: 'Angemeldet' | 'Warteliste' = 'Angemeldet';
    let effectiveStarterType: string | undefined = leadData.preferredStarterType;

    // Atomar N Plätze reservieren — Split-Group oder klassisch.
    const isSplitGroup = typeof event.durchstarterCapacity === 'number' && typeof event.funstarterCapacity === 'number'
      && (event.durchstarterCapacity > 0 || event.funstarterCapacity > 0);
    // v30.67: Welcher Zähler um teamCount erhöht wurde — für die Rückgabe
    // nicht belegter Plätze, falls Inserts scheitern (s.u.).
    let reservedSeatField: 'SeatsTaken' | 'SeatsTakenDurch' | 'SeatsTakenFun' | null = null;
    if (isSplitGroup && leadData.preferredStarterType) {
      const cap = leadData.preferredStarterType === 'Durchstarter'
        ? (event.durchstarterCapacity || 0)
        : (event.funstarterCapacity || 0);
      const seat = await eventService.reserveSeat(subsiteUrl, leadData.preferredStarterType as 'Durchstarter' | 'Funstarter', cap, teamCount);
      if (seat !== 'reserved') {
        status = 'Warteliste';
        effectiveStarterType = undefined;
      } else if (cap > 0) {
        reservedSeatField = eventService.seatFieldFor(leadData.preferredStarterType);
      }
    } else if (event.maxParticipants > 0) {
      const seat = await eventService.reserveSeat(subsiteUrl, '', event.maxParticipants, teamCount);
      if (seat !== 'reserved') status = 'Warteliste';
      else reservedSeatField = 'SeatsTaken';
    }
    // v30.67: Reservierte, aber nicht belegte Plätze zurückgeben — sonst
    // blockieren sie als Phantom-Plätze andere Anmeldungen, bis jemand
    // „Zähler abgleichen" klickt (adjustSeatCounterField ist additiv + ETag-CAS).
    const releaseSeats = async (n: number): Promise<void> => {
      if (!reservedSeatField || n <= 0) return;
      try { await eventService.adjustSeatCounterField(subsiteUrl, reservedSeatField, -n); }
      catch (err) { console.warn('[DEX] registerTeam: Sitzplatz-Rückgabe fehlgeschlagen:', err); }
    };

    // FieldMap aus Custom Fields extrahieren (cf.id -> spInternalName)
    const fieldMap: Record<string, string> = {};
    for (const f of event.eventSpecificFields) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const spName = (f as any).spInternalName;
      if (spName) fieldMap[f.id] = spName;
    }

    const actorName = currentUserName;
    const actorEmail = currentUserEmail;

    // Lead-Profil + Member-Profile parallel laden.
    const leadProfilePromise = leadData.email.toLowerCase() === currentUserEmail.toLowerCase()
      ? eventService.getCurrentUserProfile()
      : eventService.getUserProfileByEmail(leadData.email);
    const memberProfilePromises = members.map(m => eventService.getUserProfileByEmail(m.email));
    const [leadProfile, ...memberProfiles] = await Promise.all([leadProfilePromise, ...memberProfilePromises]);

    // Parse "Lastname, Firstname" → { firstName, lastName }. Deloitte-AD-Format.
    const parseDisplayName = (raw: string): { firstName: string; lastName: string } => {
      const dn = (raw || '').trim();
      if (!dn) return { firstName: '', lastName: '' };
      if (dn.indexOf(',') >= 0) {
        const parts = dn.split(',').map(s => s.trim());
        return { firstName: parts[1] || '', lastName: parts[0] || '' };
      }
      const parts = dn.split(/\s+/).filter(Boolean);
      if (parts.length === 0) return { firstName: '', lastName: '' };
      if (parts.length === 1) return { firstName: parts[0], lastName: '' };
      return { firstName: parts[0], lastName: parts.slice(1).join(' ') };
    };

    // v30.67: Inserts NACHEINANDER statt als Promise.all, mit Fehlerzähler.
    // Fünf gleichzeitige POSTs in der Anmeldewelle sind genau die Konstellation,
    // die SharePoint mit 429 drosselt; registerTeamMember fängt das ab und
    // liefert {ok:false}. Vorher galt `some(r => r.ok)` als Erfolg: Zwei von
    // fünf Zeilen fehlten, der Lead sah die Erfolgsseite, die zwei Personen
    // erfuhren nichts und standen in keiner Liste — und SeatsTaken stand auf
    // fünf. Dasselbe Muster, das CLAUDE.md für deleteParticipantData als
    // Ursache des Register-Rückstands beschreibt. Jetzt: Lead zuerst — ohne
    // Lead kein Team, dann wird gar nichts angelegt; scheitert ein Mitglied,
    // laufen die übrigen weiter, der Teilerfolg wird als solcher gemeldet und
    // die nicht belegten Plätze gehen zurück. (Die Counter-CAS in
    // getNextTeilnehmerId garantiert weiterhin eindeutige TeilnehmerIDs.)
    type TeamInsertResult = { ok: boolean; email: string; firstName: string; lastName: string };
    const results: TeamInsertResult[] = [];
    // Lead
    {
      const r = await eventService.registerTeamMember(subsiteUrl, {
        firstName: leadData.firstName,
        lastName: leadData.lastName,
        email: leadData.email,
        profile: leadProfile,
        status,
        teamId,
        teamLead: true,
        teamName,
        customData: leadData.customData,
        customFieldMap: fieldMap,
        starterType: effectiveStarterType,
        preferredStarterType: leadData.preferredStarterType,
        registeredByName: actorName,
        registeredByEmail: actorEmail,
        salutation: leadData.salutation,
      });
      results.push({ ok: r.ok, email: leadData.email, firstName: leadData.firstName, lastName: leadData.lastName });
      if (!r.ok) {
        await releaseSeats(teamCount);
        return { ok: false, reason: 'insert-failed' };
      }
    }
    // Members
    for (let idx = 0; idx < members.length; idx++) {
      const m = members[idx];
      const profile = memberProfiles[idx] || { department: '', location: '', jobTitle: '', phone: '' };
      const parsed = parseDisplayName(m.displayName);
      const r = await eventService.registerTeamMember(subsiteUrl, {
        firstName: parsed.firstName,
        lastName: parsed.lastName,
        email: m.email,
        profile,
        status,
        teamId,
        teamLead: false,
        teamName,
        // v18.12: Custom-Field-Antworten des Mitglieds (z.B. Essenspräferenz).
        customData: m.customData || {},
        customFieldMap: fieldMap,
        starterType: effectiveStarterType,
        preferredStarterType: leadData.preferredStarterType,
        registeredByName: actorName,
        registeredByEmail: actorEmail,
        // Anrede der Mitglieder bleibt leer — kein Picker für Member-Anreden.
        salutation: '',
      });
      results.push({ ok: r.ok, email: m.email, firstName: parsed.firstName, lastName: parsed.lastName });
    }
    const failedInserts = results.filter(r => !r.ok);
    if (failedInserts.length > 0) {
      console.warn('[DEX] registerTeam: Insert fehlgeschlagen für', failedInserts.map(f => f.email).join(', '));
      await releaseSeats(failedInserts.length);
    }

    // Pro erfolgreiche Anmeldung: Bestätigungs-Mail + Outlook-Termin queuen.
    const lang = event.emailLanguage || 'EN';
    const isDe = lang.toUpperCase() === 'DE';
    // v11.87: Team-Info-Block — Mitglieder-Liste, Belegung, Cancel-Hinweis.
    // Baue die Mitglieder-Liste aus den erfolgreichen Inserts auf — Reihenfolge
    // entspricht dem Insert-Pfad (Lead zuerst, dann Members in der Eingabe-
    // Reihenfolge). TeamSize aus dem Event-Config, Fallback auf die Anzahl
    // der tatsächlich angemeldeten Personen.
    const successResults = results.filter(r => r.ok);
    const teamMembersForBlock = successResults.map((r, i) => ({
      firstName: r.firstName,
      lastName: r.lastName,
      isLead: i === 0,
    }));
    const teamSizeForBlock = (typeof event.teamSize === 'number' && event.teamSize > 0)
      ? event.teamSize
      : successResults.length;

    for (const r of results) {
      if (!r.ok) continue;
      const templateType: string = status === 'Warteliste' ? 'Warteliste' : 'Anmeldung';
      const vars = {
        Name: r.firstName,
        EventTitle: event.title,
        Organizer: formatOrganizerList(event.organizers, lang),
        AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
        WaitlistPosition: '',
      };
      let emailData: { subject: string; body: string };
      const spTemplateRaw = await eventService.getEmailTemplate(templateType, lang).catch(() => null);
      const spTemplate = applyEventTemplateOverride(spTemplateRaw, event.emailTemplateOverrides, templateType);
      if (spTemplate) {
        emailData = buildEmailFromTemplate(spTemplate, vars);
      } else {
        emailData = status === 'Warteliste'
          ? waitlistEmail(r.firstName, event.title, 0)
          : registrationEmail(r.firstName, event.title);
      }
      // v11.87: Team-Info-Block + Consent-Hinweis injecten.
      // v16.3: Vorher wurde der Block direkt nach <body> eingefügt — damit
      // landete er VOR dem Deloitte-Template-Header (Logo, grüner Balken,
      // Headline). Stattdessen jetzt INNERHALB des Content-<td> einsetzen,
      // also direkt vor dem eigentlichen Mail-Body. Wir matchen das Content-
      // Padding (`padding:0 30px 30px 30px`) als Marker — dieser Style ist
      // im wrapTemplate eindeutig für das Body-Td.
      const teamInfoHtml = teamInfoBlockHtml({
        teamName,
        members: teamMembersForBlock,
        teamSize: teamSizeForBlock,
        isDe,
        registeredByName: currentUserName,
        consentRequired: true,
      });
      const bodyWithHint = injectIntoEmailContent(emailData.body, teamInfoHtml);
      if (!event.disableEmails) {
        const fullName = `${r.firstName} ${r.lastName}`.trim();
        eventService.queueEmail(
          emailData.subject, r.email, fullName, bodyWithHint,
          templateType, event.title, eventId
        ).catch(err => console.warn('[DEX] team queueEmail failed:', err));
      }
      if (status !== 'Warteliste' && !event.disableOutlook) {
        eventService.queueOutlookEvent(
          r.email, eventId, event.title, 'Einladen'
        ).catch(err => console.warn('[DEX] team queueOutlookEvent failed:', err));
      }
      if (event.eventNumber) {
        eventService.upsertParticipant(
          r.firstName, r.lastName, r.email, event.eventNumber, status
        ).catch(err => console.warn('[DEX] team upsertParticipant failed:', err));
      }
    }

    // Audit-Log (fire-and-forget).
    eventService.writeChangeLog({
      action: 'TeamRegistered',
      targetType: 'Participant',
      targetId: leadData.email,
      targetName: `${leadData.firstName} ${leadData.lastName}`.trim(),
      eventId,
      eventTitle: event.title,
      details: { teamId, teamSize: teamCount, status, members: members.map(m => m.email), ...(failedInserts.length > 0 ? { failed: failedInserts.map(f => f.email) } : {}) },
    }).catch(() => { /* */ });

    if (status === 'Angemeldet') {
      // v30.67: Nur die Zeilen zählen, die es wirklich gibt.
      eventService.bumpKpiParticipants(successResults.length).catch(() => { /* best-effort */ });
    }
    await loadEvents();
    // v30.67: Teilerfolg ist kein Erfolg — die Anmeldeseite darf nicht die
    // Erfolgsseite zeigen, während zwei Personen fehlen. Die Adressen stehen
    // im Grund, damit der Lead sie einzeln nachmelden kann.
    if (failedInserts.length > 0) {
      return { ok: false, teamId, status, reason: `partial-insert:${failedInserts.map(f => f.email).join(',')}` };
    }
    return { ok: true, teamId, status };
  }

  /**
   * v11.83: Einzelnes Mitglied zu einem bestehenden Team hinzufügen.
   * Wird vom „+ Mitglied"-Button im MyEvents-Team-Badge benutzt — nur für
   * Leads sichtbar, daher hier keine separate Lead-Authorisierung; die
   * UI versteckt den Button.
   *
   * Schritte:
   *   1) Doppel-Anmelde-Check via `isUserAlreadyOnEvent`. Wenn die Person
   *      schon angemeldet ist, brechen wir mit klarem Reason ab.
   *   2) Atomar 1 Sitzplatz reservieren — split-aware. Bei Vollbelegung
   *      landet das neue Mitglied auf der Warteliste (kein Hard-Fail).
   *   3) `registerTeamMember` mit identischer TeamId, `teamLead=false`.
   *   4) Bestätigungs-Mail + Outlook-Termin queuen.
   *   5) Optional: Info-Mail an die anderen Mitglieder „X ist eurem Team
   *      beigetreten" (best-effort).
   */
  async function assignTeamlessToTeam(
    eventId: string,
    teamId: string,
    teamName: string | undefined,
    existingRegId: number,
    isLead: boolean = false,
    // v22.40: optionaler Mail-Versand an die zugeordnete (bereits angemeldete)
    // Person. Die Empfänger-Daten kommen vom Aufrufer (AdminPage hat die
    // Registrierungs-Zeile bereits geladen) — so muss niemand die E-Mail
    // erneut eingeben. Best-effort, blockiert die Zuordnung nicht.
    opts?: { sendMail?: boolean; recipientEmail?: string; recipientFirstName?: string; recipientLastName?: string; ccEmail?: string },
  ): Promise<boolean> {
    const event = events.find(e => e.id === eventId);
    if (!event || !event.subsiteUrl) return false;
    const ok = await eventService.assignRegistrationToTeam(event.subsiteUrl, existingRegId, teamId, teamName, isLead);
    if (ok && opts?.sendMail && opts.recipientEmail && !event.disableEmails) {
      try {
        const lang = event.emailLanguage || 'EN';
        const isDe = (lang || 'EN').toUpperCase() === 'DE';
        const members = await eventService.getTeamMembers(event.subsiteUrl, teamId).catch(() => []);
        const active = members.filter(m => m.Status !== 'Abgemeldet');
        const teamSizeForBlock = (typeof event.teamSize === 'number' && event.teamSize > 0) ? event.teamSize : active.length;
        const teamInfoHtml = teamInfoBlockHtml({
          teamName,
          members: active.map(m => ({ firstName: m.Vorname || '', lastName: m.Nachname || '', isLead: !!m.TeamLead })),
          teamSize: teamSizeForBlock,
          isDe,
          registeredByName: currentUserName,
          consentRequired: false,
        });
        const first = (opts.recipientFirstName || '').trim();
        const fullName = `${first} ${opts.recipientLastName || ''}`.trim() || opts.recipientEmail;
        const teamNameStr = teamName ? `„${teamName}"` : (isDe ? 'einem Team' : 'a team');
        const inner = isDe
          ? `<p>Hallo ${first || fullName},</p><p>du wurdest für das Event <strong>${event.title}</strong> dem Team ${teamNameStr} zugeordnet. Deine bestehende Anmeldung bleibt unverändert — du musst nichts weiter tun.</p>`
          : `<p>Hello ${first || fullName},</p><p>you have been assigned to team ${teamNameStr} for the event <strong>${event.title}</strong>. Your existing registration stays unchanged — nothing else to do.</p>`;
        const subject = isDe ? `Team-Zuordnung: ${event.title}` : `Team assignment: ${event.title}`;
        const body = wrapTemplate('#86bc25', isDe ? 'Team-Zuordnung' : 'Team assignment', `Event ${event.title}`, inner + teamInfoHtml);
        await eventService.queueEmail(subject, opts.recipientEmail, fullName, body, 'TeamMemberJoined', event.title, eventId, opts.ccEmail || undefined);
      } catch (err) { console.warn('[DEX] assignTeamlessToTeam mail failed:', err); }
    }
    return ok;
  }

  // v22.49: „Neues Mitglied"-Info an die BESTEHENDEN Team-Mitglieder, wenn der
  // Organizer im Zuordnungs-Modal „auch die übrigen informieren" wählt. scope:
  // 'all' = alle bisherigen aktiven Mitglieder, 'lead' = nur der Team-Lead.
  // excludeEmails = die gerade neu zugeordneten Personen (nicht sich selbst
  // benachrichtigen). Best-effort, gated über event.disableEmails.
  async function notifyExistingTeamMembers(
    eventId: string,
    teamId: string,
    teamName: string | undefined,
    newMemberNames: string[],
    excludeEmails: string[],
    scope: 'all' | 'lead' = 'all',
  ): Promise<void> {
    const event = events.find(e => e.id === eventId);
    if (!event || !event.subsiteUrl || event.disableEmails) return;
    try {
      const lang = event.emailLanguage || 'EN';
      const isDe = (lang || 'EN').toUpperCase() === 'DE';
      const members = await eventService.getTeamMembers(event.subsiteUrl, teamId).catch(() => []);
      const exclude = new Set(excludeEmails.map(e => (e || '').toLowerCase()));
      let others = members.filter(m => m.Status !== 'Abgemeldet' && !exclude.has((m.ParticipantEmail || '').toLowerCase()));
      if (scope === 'lead') others = others.filter(m => !!m.TeamLead);
      if (others.length === 0) return;
      const tpl = await eventService.getEmailTemplate('TeamMemberJoined', lang).catch(() => null);
      const newStr = newMemberNames.filter(Boolean).join(isDe ? ' und ' : ' and ');
      const teamNameStr = teamName ? `„${teamName}"` : '';
      for (const other of others) {
        const otherFirst = other.Vorname || '';
        const otherFull = `${other.Vorname || ''} ${other.Nachname || ''}`.trim() || other.ParticipantEmail;
        let mail: { subject: string; body: string };
        if (tpl) {
          mail = buildEmailFromTemplate(tpl, {
            Name: otherFirst || otherFull,
            NewMemberName: newStr,
            TeamName: teamNameStr,
            EventTitle: event.title,
            AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
          });
        } else {
          const inner = isDe
            ? `<p>Hallo ${otherFirst},</p><p><strong>${newStr}</strong> ${newMemberNames.length > 1 ? 'sind' : 'ist'} eurem Team ${teamNameStr} beigetreten.</p>`
            : `<p>Hello ${otherFirst},</p><p><strong>${newStr}</strong> joined your team ${teamNameStr}.</p>`;
          mail = {
            subject: isDe ? `Neues Team-Mitglied — ${event.title}` : `New team member — ${event.title}`,
            body: wrapTemplate('#86bc25', isDe ? 'Team-Update' : 'Team update', `Event ${event.title}`, inner),
          };
        }
        await eventService.queueEmail(mail.subject, other.ParticipantEmail, otherFull, mail.body, 'TeamMemberJoined', event.title, eventId).catch(() => { /* */ });
      }
    } catch (err) { console.warn('[DEX] notifyExistingTeamMembers failed:', err); }
  }

  async function addTeamMember(
    eventId: string,
    teamId: string,
    teamName: string | undefined,
    member: { email: string; displayName: string },
    // v18.73: optionale event-spezifische Antworten des Beitretenden — werden
    // beim Insert mitgeschrieben (vorher immer leer). Genutzt vom Self-Join
    // über die Anmeldeseite, damit der Beitritt die Pflicht-Custom-Felder nicht
    // mehr überspringt.
    customData?: Record<string, string>,
    // v22.49: Kommunikation optional steuerbar (vom Admin-Zuordnungs-Modal):
    // suppressMemberMail = keine Anmeldebestätigung + kein Outlook an die neue
    // Person; suppressOthersMail = die eingebaute „neues Mitglied"-Info an die
    // übrigen Mitglieder NICHT senden (das übernimmt dann notifyExistingTeam-
    // Members mit der gewählten Reichweite); ccEmail = CC der Bestätigung.
    opts?: { suppressMemberMail?: boolean; suppressOthersMail?: boolean; ccEmail?: string }
  ): Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, reason: 'event-not-found' };
    const event = events.find(e => e.id === eventId);
    if (!event) return { ok: false, reason: 'event-not-found' };
    if (!teamId) return { ok: false, reason: 'invalid-team-id' };

    // v30.67: null = nicht prüfbar → ablehnen statt durchwinken (s. probeAlreadyActive).
    const blocked = await probeAlreadyActive(subsiteUrl, member.email);
    if (blocked === null) return { ok: false, reason: 'dup-check-failed' };
    if (blocked) return { ok: false, reason: `already-registered:${member.email}` };

    // Vorhandene Mitglieder laden — um die richtige Gruppe (Split) und
    // die existierenden Custom-Field-Antworten als Vorlage zu erben.
    const existingMembers = await eventService.getTeamMembers(subsiteUrl, teamId);
    const activeMembers = existingMembers.filter(m => m.Status !== 'Abgemeldet');
    const teamSizeCfg = event.teamSize || (activeMembers.length + 1);
    if (activeMembers.length >= teamSizeCfg) {
      return { ok: false, reason: 'team-full' };
    }
    const inheritedStarterType = activeMembers.find(m => !!m.PreferredStarterType)?.PreferredStarterType || '';

    let status: 'Angemeldet' | 'Warteliste' = 'Angemeldet';
    let effectiveStarterType: string | undefined = inheritedStarterType || undefined;
    const isSplitGroup = typeof event.durchstarterCapacity === 'number' && typeof event.funstarterCapacity === 'number'
      && (event.durchstarterCapacity > 0 || event.funstarterCapacity > 0);
    // v30.67: Welcher Zähler erhöht wurde — Rückgabe bei gescheitertem Insert (s.u.).
    let reservedSeatField: 'SeatsTaken' | 'SeatsTakenDurch' | 'SeatsTakenFun' | null = null;
    if (isSplitGroup && inheritedStarterType) {
      const cap = inheritedStarterType === 'Durchstarter'
        ? (event.durchstarterCapacity || 0)
        : (event.funstarterCapacity || 0);
      const seat = await eventService.reserveSeat(subsiteUrl, inheritedStarterType as 'Durchstarter' | 'Funstarter', cap, 1);
      if (seat !== 'reserved') {
        status = 'Warteliste';
        effectiveStarterType = undefined;
      } else if (cap > 0) {
        reservedSeatField = eventService.seatFieldFor(inheritedStarterType);
      }
    } else if (event.maxParticipants > 0) {
      const seat = await eventService.reserveSeat(subsiteUrl, '', event.maxParticipants, 1);
      if (seat !== 'reserved') status = 'Warteliste';
      else reservedSeatField = 'SeatsTaken';
    }

    // Profil laden + DisplayName parsen.
    const profile = await eventService.getUserProfileByEmail(member.email);
    const parseDisplayName = (raw: string): { firstName: string; lastName: string } => {
      const dn = (raw || '').trim();
      if (!dn) return { firstName: '', lastName: '' };
      if (dn.indexOf(',') >= 0) {
        const parts = dn.split(',').map(s => s.trim());
        return { firstName: parts[1] || '', lastName: parts[0] || '' };
      }
      const parts = dn.split(/\s+/).filter(Boolean);
      if (parts.length === 0) return { firstName: '', lastName: '' };
      if (parts.length === 1) return { firstName: parts[0], lastName: '' };
      return { firstName: parts[0], lastName: parts.slice(1).join(' ') };
    };
    const parsed = parseDisplayName(member.displayName);

    const r = await eventService.registerTeamMember(subsiteUrl, {
      firstName: parsed.firstName,
      lastName: parsed.lastName,
      email: member.email,
      profile,
      status,
      teamId,
      teamLead: false,
      teamName,
      customData: customData || {},
      starterType: effectiveStarterType,
      preferredStarterType: inheritedStarterType || undefined,
      registeredByName: currentUserName,
      registeredByEmail: currentUserEmail,
      salutation: '',
    });
    if (!r.ok) {
      // v30.67: Reservierten Platz zurückgeben — s. registerForEvent.
      if (reservedSeatField) {
        try { await eventService.adjustSeatCounterField(subsiteUrl, reservedSeatField, -1); }
        catch (err) { console.warn('[DEX] addTeamMember: Sitzplatz-Rückgabe fehlgeschlagen:', err); }
      }
      return { ok: false, reason: 'insert-failed' };
    }

    // Bestätigungs-Mail + Outlook + DEX_Participants — same pattern as
    // registerTeam aber für EINEN Member.
    const lang = event.emailLanguage || 'EN';
    const isDe = (lang || 'EN').toUpperCase() === 'DE';
    const templateType = status === 'Warteliste' ? 'Warteliste' : 'Anmeldung';
    const vars = {
      Name: parsed.firstName,
      EventTitle: event.title,
      Organizer: formatOrganizerList(event.organizers, lang),
      AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      WaitlistPosition: '',
    };
    let emailData: { subject: string; body: string };
    const spTplRaw = await eventService.getEmailTemplate(templateType, lang).catch(() => null);
    const spTpl = applyEventTemplateOverride(spTplRaw, event.emailTemplateOverrides, templateType);
    if (spTpl) {
      emailData = buildEmailFromTemplate(spTpl, vars);
    } else {
      emailData = status === 'Warteliste'
        ? waitlistEmail(parsed.firstName, event.title, 0)
        : registrationEmail(parsed.firstName, event.title);
    }
    // v11.87: Team-Info-Block mit allen aktiven Mitgliedern inkl. dem neuen.
    // activeMembers wurde vor dem Insert geladen — wir hängen den frisch
    // angemeldeten User vorne dran als Nicht-Lead an. Den Lead identifizieren
    // wir per TeamLead-Flag.
    const allActiveForBlock: Array<{ firstName: string; lastName: string; isLead: boolean }> = [
      ...activeMembers.map(m => ({
        firstName: m.Vorname || '',
        lastName: m.Nachname || '',
        isLead: !!m.TeamLead,
      })),
      { firstName: parsed.firstName, lastName: parsed.lastName, isLead: false },
    ];
    // Lead zuerst sortieren, danach in Insert-Reihenfolge belassen.
    allActiveForBlock.sort((a, b) => (a.isLead === b.isLead) ? 0 : (a.isLead ? -1 : 1));
    const teamSizeForBlock = (typeof event.teamSize === 'number' && event.teamSize > 0)
      ? event.teamSize
      : allActiveForBlock.length;
    const teamInfoHtml = teamInfoBlockHtml({
      teamName,
      members: allActiveForBlock,
      teamSize: teamSizeForBlock,
      isDe,
      registeredByName: currentUserName,
      consentRequired: true,
    });
    const bodyWithHint = injectIntoEmailContent(emailData.body, teamInfoHtml);
    if (!event.disableEmails && !opts?.suppressMemberMail) {
      const fullName = `${parsed.firstName} ${parsed.lastName}`.trim() || member.email;
      eventService.queueEmail(
        emailData.subject, member.email, fullName, bodyWithHint,
        templateType, event.title, eventId, opts?.ccEmail || undefined
      ).catch(err => console.warn('[DEX] addTeamMember queueEmail failed:', err));
    }
    if (status !== 'Warteliste' && !event.disableOutlook && !opts?.suppressMemberMail) {
      eventService.queueOutlookEvent(
        member.email, eventId, event.title, 'Einladen'
      ).catch(err => console.warn('[DEX] addTeamMember queueOutlookEvent failed:', err));
    }
    if (event.eventNumber) {
      eventService.upsertParticipant(
        parsed.firstName, parsed.lastName, member.email, event.eventNumber, status
      ).catch(err => console.warn('[DEX] addTeamMember upsertParticipant failed:', err));
    }

    // v12.14: Info-Mail an verbleibende Mitglieder kommt jetzt aus
    // DEX_EmailTemplates (TemplateType=TeamMemberJoined). Pre-Wrap +
    // Variable-Substitution durch buildEmailFromTemplate.
    if (!event.disableEmails && !opts?.suppressOthersMail) {
      const tpl = await eventService.getEmailTemplate('TeamMemberJoined', lang).catch(() => null);
      const newMemberFullName = `${parsed.firstName} ${parsed.lastName}`.trim();
      const teamNameStr = teamName ? `„${teamName}"` : '';
      for (const other of activeMembers) {
        const otherFirst = other.Vorname || '';
        const otherFull = `${other.Vorname || ''} ${other.Nachname || ''}`.trim() || other.ParticipantEmail;
        const vars: Record<string, string> = {
          Name: otherFirst || otherFull,
          NewMemberName: newMemberFullName,
          TeamName: teamNameStr,
          EventTitle: event.title,
          AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
        };
        let mail: { subject: string; body: string };
        if (tpl) {
          mail = buildEmailFromTemplate(tpl, vars);
        } else {
          // Fallback: alter Inline-Text falls Template nicht geseeded ist.
          const inner = isDe
            ? `<p>Hallo ${otherFirst},</p><p><strong>${newMemberFullName}</strong> ist eurem Team ${teamNameStr} beigetreten.</p>`
            : `<p>Hello ${otherFirst},</p><p><strong>${newMemberFullName}</strong> joined your team ${teamNameStr}.</p>`;
          mail = {
            subject: isDe ? `Neues Team-Mitglied — ${event.title}` : `New team member — ${event.title}`,
            body: wrapTemplate('#86bc25', isDe ? 'Team-Update' : 'Team update', `Event ${event.title}`, inner),
          };
        }
        eventService.queueEmail(
          mail.subject, other.ParticipantEmail, otherFull,
          mail.body, 'TeamMemberJoined', event.title, eventId
        ).catch(() => { /* best-effort */ });
      }
    }

    if (status === 'Angemeldet') {
      eventService.bumpKpiParticipants(1).catch(() => { /* best-effort */ });
    }
    eventService.writeChangeLog({
      action: 'TeamMemberAdded',
      targetType: 'Participant',
      targetId: member.email,
      targetName: `${parsed.firstName} ${parsed.lastName}`.trim(),
      eventId,
      eventTitle: event.title,
      details: { teamId, addedBy: currentUserEmail, status },
    }).catch(() => { /* */ });

    await loadEvents();
    return { ok: true, status };
  }

  /**
   * v11.83: Direkter Team-Beitritt aus der Anmeldeseite (ohne Approval).
   * Funktional identisch zu `addTeamMember`, aber läuft mit dem
   * eingeloggten User als Member. Der Submit-Pfad in RegistrationPage
   * unterscheidet zwischen `joinTeam` (Approval OFF) und
   * `createTeamJoinRequest` (Approval ON).
   */
  async function joinTeam(
    eventId: string,
    teamId: string,
    teamName: string | undefined,
    // v18.73: event-spezifische Antworten des Beitretenden durchreichen.
    customData?: Record<string, string>
  ): Promise<{ ok: boolean; status?: 'Angemeldet' | 'Warteliste'; reason?: string }> {
    return addTeamMember(eventId, teamId, teamName, {
      email: currentUserEmail,
      displayName: currentUserName,
    }, customData);
  }

  /**
   * v11.84: Lead-Rolle innerhalb eines Teams übergeben. Nur für Admin
   * Center gedacht — die UI versteckt den Button für alle anderen Rollen.
   * Wirft kein Mail zur "alten" Person, sondern eine Info-Mail an alle
   * aktiven Team-Mitglieder mit dem Hinweis "Die Team-Lead-Rolle wurde
   * an <Name> übergeben". Audit-Eintrag im ChangeLog.
   */
  async function transferTeamLead(
    eventId: string,
    teamId: string,
    newLeadEmail: string
  ): Promise<{ ok: boolean; reason?: string }> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, reason: 'event-not-found' };
    const event = events.find(e => e.id === eventId);
    if (!event) return { ok: false, reason: 'event-not-found' };
    if (!teamId || !newLeadEmail) return { ok: false, reason: 'invalid-input' };

    const members = await eventService.getTeamMembers(subsiteUrl, teamId);
    const active = members.filter(m => m.Status !== 'Abgemeldet');
    const oldLead = active.find(m => !!m.TeamLead);
    const newLead = active.find(m => (m.ParticipantEmail || '').toLowerCase() === newLeadEmail.toLowerCase());
    if (!newLead) return { ok: false, reason: 'new-lead-not-found' };
    if (!oldLead) {
      // Kein aktiver Lead — einfach den neuen promoten, kein Demote nötig.
      const okPromote = await eventService.promoteToTeamLead(subsiteUrl, newLead.Id);
      if (!okPromote) return { ok: false, reason: 'promote-failed' };
    } else {
      if (oldLead.Id === newLead.Id) return { ok: false, reason: 'already-lead' };
      const ok = await eventService.transferTeamLead(subsiteUrl, oldLead.Id, newLead.Id);
      if (!ok) return { ok: false, reason: 'transfer-failed' };
    }

    // v12.14: Info-Mails an alle aktiven Mitglieder kommen jetzt aus
    // TemplateType=TeamLeadTransferred. {{NewLeadBlock}}-Platzhalter wird
    // pro Empfänger gefüllt — für den neuen Lead ein zusätzlicher Hinweis,
    // sonst leer.
    if (!event.disableEmails) {
      const lang = event.emailLanguage || 'EN';
      const isDe = lang.toUpperCase() === 'DE';
      const tpl = await eventService.getEmailTemplate('TeamLeadTransferred', lang).catch(() => null);
      const newLeadName = `${newLead.Vorname || ''} ${newLead.Nachname || ''}`.trim() || newLead.ParticipantEmail;
      const teamName = members.find(m => !!m.TeamName)?.TeamName || '';
      const teamNameStr = teamName ? `„${teamName}"` : '';
      const newLeadBlockHtml = isDe
        ? `<p style="padding:10px 14px;background:rgba(134,188,37,0.10);border:1px solid #86bc25;border-radius:6px;"><strong>Du bist ab jetzt Team-Lead.</strong> Du kannst neue Mitglieder über „Meine Events" hinzufügen und ggf. Beitritts-Anfragen entscheiden.</p>`
        : `<p style="padding:10px 14px;background:rgba(134,188,37,0.10);border:1px solid #86bc25;border-radius:6px;"><strong>You are now the team lead.</strong> You can add new members via „My Events" and decide on join requests if any.</p>`;
      for (const other of active) {
        const otherFirst = other.Vorname || '';
        const otherFull = `${other.Vorname || ''} ${other.Nachname || ''}`.trim() || other.ParticipantEmail;
        const isNewLeadMember = other.Id === newLead.Id;
        const vars: Record<string, string> = {
          Name: otherFirst || otherFull,
          NewLeadName: newLeadName,
          TeamName: teamNameStr,
          EventTitle: event.title,
          NewLeadBlock: isNewLeadMember ? newLeadBlockHtml : '',
          AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
        };
        let mail: { subject: string; body: string };
        if (tpl) {
          mail = buildEmailFromTemplate(tpl, vars);
        } else {
          const inner = isDe
            ? `<p>Hallo ${otherFirst},</p><p>Die Team-Lead-Rolle in deinem Team ${teamNameStr} wurde an <strong>${newLeadName}</strong> übergeben.</p>${isNewLeadMember ? newLeadBlockHtml : ''}`
            : `<p>Hello ${otherFirst},</p><p>The team lead role in your team ${teamNameStr} has been transferred to <strong>${newLeadName}</strong>.</p>${isNewLeadMember ? newLeadBlockHtml : ''}`;
          mail = {
            subject: isDe ? `Team-Lead-Wechsel — ${event.title}` : `Team lead change — ${event.title}`,
            body: wrapTemplate('#86bc25', isDe ? 'Team-Lead-Wechsel' : 'Team lead change', `Event ${event.title}`, inner),
          };
        }
        eventService.queueEmail(
          mail.subject, other.ParticipantEmail, otherFull,
          mail.body, 'TeamLeadTransferred', event.title, eventId
        ).catch(() => { /* best-effort */ });
      }
    }

    eventService.writeChangeLog({
      action: 'TeamLeadTransferred',
      targetType: 'Participant',
      targetId: newLeadEmail,
      targetName: `${newLead.Vorname || ''} ${newLead.Nachname || ''}`.trim(),
      eventId,
      eventTitle: event.title,
      details: { teamId, fromLeadEmail: oldLead?.ParticipantEmail || '', toLeadEmail: newLeadEmail, actor: currentUserEmail },
    }).catch(() => { /* */ });

    return { ok: true };
  }

  /**
   * v11.83: Eine Beitritts-Anfrage in DEX_TeamJoinRequests anlegen +
   * Lead-Notification queuen. Die App liest die TeamId vom UI, holt sich
   * den Lead aus der Subsite-Teilnehmerliste (TeamId-Match,
   * TeamLead=true) und schreibt dann eine Mail an die Lead-Email.
   */
  async function createTeamJoinRequest(
    eventId: string,
    teamId: string,
    // v18.73: event-spezifische Antworten des Anfragenden — werden in der
    // Request-Zeile gespeichert und beim Approve auf den neuen Member angewandt.
    customData?: Record<string, string>
  ): Promise<{ ok: boolean; itemId?: number; reason?: string }> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, reason: 'event-not-found' };
    const event = events.find(e => e.id === eventId);
    if (!event) return { ok: false, reason: 'event-not-found' };

    // v30.67: null = nicht prüfbar → ablehnen statt durchwinken (s. probeAlreadyActive).
    const blocked = await probeAlreadyActive(subsiteUrl, currentUserEmail);
    if (blocked === null) return { ok: false, reason: 'dup-check-failed' };
    if (blocked) return { ok: false, reason: 'already-registered' };

    // Lead finden — die Mail-Notification soll an ihn gehen.
    const members = await eventService.getTeamMembers(subsiteUrl, teamId);
    const lead = members.find(m => !!m.TeamLead && m.Status !== 'Abgemeldet');
    if (!lead) return { ok: false, reason: 'team-has-no-lead' };

    const teamNameFromMembers = members.find(m => !!m.TeamName)?.TeamName || '';
    const result = await eventService.createTeamJoinRequest({
      eventId,
      eventTitle: event.title,
      teamId,
      requesterEmail: currentUserEmail,
      requesterDisplayName: currentUserName,
      // v18.73: Antworten als JSON mitschreiben (leer = '{}').
      customData: JSON.stringify(customData || {}),
    });
    if (!result.ok) return { ok: false, reason: 'queue-failed' };

    // v12.14: Lead-Notification kommt jetzt aus TemplateType=TeamJoinRequest.
    if (!event.disableEmails) {
      const lang = event.emailLanguage || 'EN';
      const isDe = lang.toUpperCase() === 'DE';
      const tpl = await eventService.getEmailTemplate('TeamJoinRequest', lang).catch(() => null);
      const leadFirst = lead.Vorname || '';
      const leadFull = `${lead.Vorname || ''} ${lead.Nachname || ''}`.trim() || lead.ParticipantEmail;
      const appUrl = buildHashDeepLink(`${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`, { action: 'teamjoin', request: result.itemId || 0 });
      const teamNameStr = teamNameFromMembers ? `„${teamNameFromMembers}"` : '';
      const vars: Record<string, string> = {
        Name: leadFirst || leadFull,
        RequesterName: currentUserName,
        TeamName: teamNameStr,
        EventTitle: event.title,
        ApproveUrl: `${appUrl}&decision=approve`,
        RejectUrl: `${appUrl}&decision=reject`,
        AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      };
      let mail: { subject: string; body: string };
      if (tpl) {
        mail = buildEmailFromTemplate(tpl, vars);
      } else {
        const inner = isDe
          ? `<p>Hallo ${leadFirst},</p><p><strong>${currentUserName}</strong> möchte deinem Team ${teamNameStr} beitreten.</p><p style="text-align:center;margin:18px 0;"><a href="${appUrl}&decision=approve" style="display:inline-block;padding:10px 18px;background:#86bc25;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;margin-right:8px;">Bestätigen</a> <a href="${appUrl}&decision=reject" style="display:inline-block;padding:10px 18px;background:#999;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;">Ablehnen</a></p>`
          : `<p>Hello ${leadFirst},</p><p><strong>${currentUserName}</strong> would like to join your team ${teamNameStr}.</p><p style="text-align:center;margin:18px 0;"><a href="${appUrl}&decision=approve" style="display:inline-block;padding:10px 18px;background:#86bc25;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;margin-right:8px;">Approve</a> <a href="${appUrl}&decision=reject" style="display:inline-block;padding:10px 18px;background:#999;color:#fff;font-weight:600;text-decoration:none;border-radius:6px;">Reject</a></p>`;
        mail = {
          subject: isDe ? `Team-Beitritts-Anfrage — ${event.title}` : `Team join request — ${event.title}`,
          body: wrapTemplate('#86bc25', isDe ? 'Team-Beitritts-Anfrage' : 'Team join request', `Event ${event.title}`, inner),
        };
      }
      eventService.queueEmail(
        mail.subject, lead.ParticipantEmail, leadFull,
        mail.body, 'TeamJoinRequest', event.title, eventId
      ).catch(() => { /* best-effort */ });
    }

    return { ok: true, itemId: result.itemId };
  }

  /**
   * v11.83: Pending-Anfragen für ein bestimmtes Team eines Events
   * abrufen — wird in der „Beitritts-Anfragen"-Box im MyEvents-Team-
   * Badge angezeigt (nur Leads sehen sie).
   */
  async function listTeamJoinRequestsForEvent(
    eventId: string,
    teamId: string
  ): Promise<Array<{ Id: number; EventId: string; TeamId: string; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string }>> {
    if (!eventId || !teamId) return [];
    const items = await eventService.listTeamJoinRequests({ eventId, teamId, status: 'Pending' });
    return items;
  }

  /**
   * v11.83: Approve/Reject einer Beitritts-Anfrage durch den Team-Lead.
   * Bei Approve: Member-Anmeldung via `addTeamMember`-Logik + Mail an
   * Anfragenden „du wurdest aufgenommen". Bei Reject: kurze Absage-Mail.
   * Beide Pfade setzen anschliessend den Status der DEX_TeamJoinRequests-
   * Zeile auf Approved/Rejected.
   */
  async function decideTeamJoinRequest(
    requestId: number,
    decision: 'Approved' | 'Rejected'
  ): Promise<boolean> {
    // Erst die Request-Zeile holen, damit wir Event-/Team-Kontext kennen.
    const all = await eventService.listTeamJoinRequests({ status: 'Pending' });
    const req = all.find(r => r.Id === requestId);
    if (!req) return false;
    const event = events.find(e => e.id === req.EventId);
    if (!event) return false;
    const subsiteUrl = subsiteMap.current[req.EventId];
    if (!subsiteUrl) return false;

    if (decision === 'Approved') {
      // Bestehenden Team-Namen ableiten.
      const members = await eventService.getTeamMembers(subsiteUrl, req.TeamId);
      const teamName = members.find(m => !!m.TeamName)?.TeamName || '';
      // v18.73: bei der Anfrage gespeicherte event-spezifische Antworten
      // wiederherstellen und auf den neuen Member anwenden.
      let reqCustomData: Record<string, string> | undefined;
      try { reqCustomData = req.CustomData ? JSON.parse(req.CustomData) : undefined; } catch { reqCustomData = undefined; }
      const addRes = await addTeamMember(req.EventId, req.TeamId, teamName || undefined, {
        email: req.RequesterEmail,
        displayName: req.RequesterDisplayName,
      }, reqCustomData);
      if (!addRes.ok) {
        // Wir markieren die Anfrage trotzdem als Approved, wenn der Add
        // fehlschlug — der Lead bekommt ein UI-Feedback und kann manuell
        // nachsetzen. Status bleibt Pending nur bei System-Fehlern auf der
        // List-Selbst.
        return false;
      }
      await eventService.decideTeamJoinRequest(requestId, 'Approved', currentUserEmail);
      // „Du wurdest aufgenommen"-Mail wurde bereits durch addTeamMember
      // gequeued (Bestätigungs-Mail), daher hier keine doppelte Mail.
      return true;
    }

    // Reject
    const ok = await eventService.decideTeamJoinRequest(requestId, 'Rejected', currentUserEmail);
    // v12.14: Absage-Mail aus TemplateType=TeamJoinRejected.
    if (!event.disableEmails) {
      const lang = event.emailLanguage || 'EN';
      const isDe = lang.toUpperCase() === 'DE';
      const tpl = await eventService.getEmailTemplate('TeamJoinRejected', lang).catch(() => null);
      const requesterFirst = req.RequesterDisplayName.split(',').pop()?.trim() || req.RequesterDisplayName;
      const vars: Record<string, string> = {
        Name: requesterFirst,
        EventTitle: event.title,
        AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      };
      let mail: { subject: string; body: string };
      if (tpl) {
        mail = buildEmailFromTemplate(tpl, vars);
      } else {
        const inner = isDe
          ? `<p>Hallo ${requesterFirst},</p><p>deine Beitritts-Anfrage zum Team beim Event „${event.title}" wurde vom Team-Lead abgelehnt.</p>`
          : `<p>Hello ${requesterFirst},</p><p>your join request for the team at event „${event.title}" was declined by the team lead.</p>`;
        mail = {
          subject: isDe ? `Team-Beitritts-Anfrage abgelehnt — ${event.title}` : `Team join request declined — ${event.title}`,
          body: wrapTemplate('#ed8b00', isDe ? 'Team-Beitritts-Anfrage abgelehnt' : 'Team join request declined', `Event ${event.title}`, inner),
        };
      }
      eventService.queueEmail(
        mail.subject, req.RequesterEmail, req.RequesterDisplayName,
        mail.body, 'TeamJoinRejected', event.title, req.EventId
      ).catch(() => { /* best-effort */ });
    }
    return ok;
  }

  /**
   * v11.83: Aktive Teams eines Events für die „Offene Teams"-Box.
   * Filter: nur Teams mit aktivem Mitglied-Count < event.teamSize.
   * Mitgliedernamen werden bewusst NICHT zurückgegeben (Privatsphäre) —
   * nur Belegungs-Anzahl, TeamName und LeadEmail (LeadEmail wird ohnehin
   * gebraucht, weil der Beitritts-Pfad eine Lead-Notification queued).
   */
  async function listOpenTeamsForEvent(eventId: string): Promise<Array<{ teamId: string; teamName: string; activeCount: number; teamSize: number; leadEmail: string; leadDisplayName: string }>> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return [];
    const event = events.find(e => e.id === eventId);
    if (!event || !event.teamRegistrationEnabled) return [];
    const teamSizeCfg = event.teamSize || 0;
    if (teamSizeCfg < 2) return [];

    const all = await eventService.getAllRegistrations(subsiteUrl);
    // Gruppieren nach TeamId.
    const byTeam: Record<string, SPRegistration[]> = {};
    for (const r of all) {
      if (!r.TeamId) continue;
      if (r.Status === 'Abgemeldet') continue;
      if (!byTeam[r.TeamId]) byTeam[r.TeamId] = [];
      byTeam[r.TeamId].push(r);
    }
    const open: Array<{ teamId: string; teamName: string; activeCount: number; teamSize: number; leadEmail: string; leadDisplayName: string }> = [];
    for (const tid of Object.keys(byTeam)) {
      const list = byTeam[tid];
      if (list.length >= teamSizeCfg) continue;
      const lead = list.find(m => !!m.TeamLead) || list[0];
      open.push({
        teamId: tid,
        teamName: list.find(m => !!m.TeamName)?.TeamName || '',
        activeCount: list.length,
        teamSize: teamSizeCfg,
        leadEmail: lead?.ParticipantEmail || '',
        leadDisplayName: `${lead?.Vorname || ''} ${lead?.Nachname || ''}`.trim() || lead?.ParticipantEmail || '',
      });
    }
    return open;
  }

  // v30.66: Ausgelagert nach `cancellation.ts` (Modularisierung Stufe 3).
  // Die Funktionskoerper sind unveraendert; sie bekommen ihre Umgebung
  // ueber ein explizites deps-Objekt statt ueber die Closure.
  const {
    cancelRegistration, cancelTeamMember, getMyProxyRegistrations, cancelProxyRegistration, updateProxyRegistration, handBackToParticipant, declineEvent,
  } = makeCancellationActions({ eventService, events, subsiteMap, currentUserEmail, currentUserName, currentUserFirstName, calDayParentOf, getMyRegistration, loadEvents });

  async function getMyRegistration(eventId: string): Promise<SPRegistration | null> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return null;
    return eventService.getMyRegistration(subsiteUrl, currentUserEmail);
  }

  async function checkRegistrationByEmail(eventId: string, email: string): Promise<SPRegistration | null> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl || !email) return null;
    return eventService.getMyRegistration(subsiteUrl, email);
  }

  // v18.33: Self-Check-in über gescannten QR-Deep-Link.
  async function selfCheckIn(params: SelfCheckInParams): Promise<SelfCheckInResult> {
    try {
      // 1) Event auflösen — per Token (statischer QR) ODER Event-Nummer (rotierend).
      let sp: SPEvent | null = null;
      if (params.token) {
        sp = await eventService.getEventBySelfCheckInToken(params.token);
      } else if (typeof params.eventNumber === 'number' && !isNaN(params.eventNumber)) {
        sp = await eventService.getEventByEventNumber(params.eventNumber);
      }
      if (!sp) return { status: 'not-found' };

      const eventTitle = sp.Title;
      const eventStart = sp.StartDate;

      // 2) Self-Check-in muss aktiviert sein.
      if (!sp.SelfCheckInEnabled) {
        return { status: 'disabled', eventTitle, eventStart };
      }

      // 3) Rotierender Code: Frische + HMAC prüfen.
      if (!params.token) {
        const secret = sp.SelfCheckInToken || '';
        const ok = await verifyRotatingCode(
          secret,
          params.code || '',
          typeof params.windowIndex === 'number' ? params.windowIndex : NaN
        );
        if (!ok) return { status: 'expired', eventTitle, eventStart };
      }

      // 4) Zeitfenster prüfen (from/to ODER Default „Event-Tag").
      const win = isWithinCheckInWindow(
        sp.StartDate,
        sp.EndDate,
        sp.SelfCheckInFrom,
        sp.SelfCheckInTo
      );
      if (!win.ok) {
        return {
          status: 'closed',
          eventTitle,
          eventStart,
          opensAt: win.opensAt ? win.opensAt.toISOString() : undefined,
          closesAt: win.closesAt ? win.closesAt.toISOString() : undefined,
        };
      }

      // 5) Eigene Registrierung auf der Subsite finden.
      const subsiteUrl = sp.SubsiteUrl;
      if (!subsiteUrl) return { status: 'error', eventTitle, eventStart };
      // v30.67 (Review): Am Einlass scannen hunderte Personen in Minuten —
      // ein 429 lieferte null, und die Seite sagte „Keine Anmeldung gefunden"
      // zu jemandem, der angemeldet ist. Nicht lesbar = 'error' (mit „später
      // erneut versuchen"), nicht 'not-registered'.
      let regReadFailed = false;
      const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail, () => { regReadFailed = true; });
      if (regReadFailed) return { status: 'error', eventTitle, eventStart };
      if (!myReg) return { status: 'not-registered', eventTitle, eventStart };
      if (myReg.Status === 'Eingecheckt') return { status: 'already', eventTitle, eventStart };
      if (myReg.Status === 'Warteliste') return { status: 'on-waitlist', eventTitle, eventStart };
      if (myReg.Status === 'Abgemeldet') return { status: 'not-registered', eventTitle, eventStart };

      // 6) Einchecken (eigener Eintrag — Item-Level-Security erlaubt das Schreiben).
      const ok = await eventService.checkInParticipant(subsiteUrl, myReg.Id);
      return { status: ok ? 'success' : 'error', eventTitle, eventStart };
    } catch (e) {
      console.warn('[DEX] selfCheckIn error:', e);
      return { status: 'error' };
    }
  }

  async function getAllRegistrations(eventId: string, onHttpError?: (_status: number) => void): Promise<SPRegistration[]> {
    // v18: Demo-Event → synthetische Teilnehmerliste (~25 Demo-User inkl.
    // Team, Warteliste, Abmeldungen), damit der Admin die Teilnehmer-
    // Verwaltung im Demo-Modus durchspielen kann.
    if (isDemoShowcaseId(eventId)) return buildDemoRegistrations();
    const subsiteUrl = subsiteMap.current[eventId];
    // v30.37: Auch die fehlende Subsite ist ein Fehler, kein leeres Event —
    // sonst ist „Event hat keine Teilnehmerliste" von „niemand angemeldet"
    // wieder nicht zu unterscheiden. Status 0 = kein HTTP-Versuch möglich.
    if (!subsiteUrl) { if (onHttpError) onHttpError(0); return []; }
    return eventService.getAllRegistrations(subsiteUrl, onHttpError);
  }

  // v24.0: Team-E-Mails eines Events (Organizer + Co-Organizer + Test-Team +
  // Check-in-/QR-Team) — alles, was NICHT als „echter Teilnehmer" zählt.
  function teamEmailSetFor(ev: DeloitteEvent): Set<string> {
    const s = new Set<string>();
    const add = (arr?: string[]): void => { (arr || []).forEach(e => { const x = (e || '').toLowerCase().trim(); if (x) s.add(x); }); };
    add(ev.organizerEmails); add(ev.coOrganizerEmails); add(ev.testTeamEmails); add(ev.qrScannerEmails);
    return s;
  }

  // v24.0: Zählt Anmeldungen, die ÜBER das Organizer-Team hinausgehen — also
  // „echte" Teilnehmer. Status-unabhängig (auch Abmeldungen zählen: eine fremde
  // Person, die sich je angemeldet hatte, belegt, dass das Event aktiv/öffentlich
  // war). Hauptevent UND alle Sub-Events werden geprüft. Grundlage für die
  // Lösch-Sperre: solche Events darf nur ein Admin und frühestens 1 Jahr nach
  // dem Event-Ende löschen.
  async function countExternalRegistrations(event: DeloitteEvent): Promise<number> {
    if (isDemoShowcaseId(event.id)) return 0;
    // v24.47: Entwürfe (isFictive = nie live geschaltet) waren nie öffentlich
    // sichtbar — etwaige Anmeldungen sind reine Test-Daten des Organizer-Teams
    // (z.B. der Organizer meldet sich selbst probehalber an). Für solche Events
    // greift die Aufbewahrungs-Sperre nicht → immer löschbar.
    if (event.isFictive) return 0;
    const children = events.filter(e => e.parentEventId === event.id);
    const baseTeam = teamEmailSetFor(event);
    // v24.47: NUR aktive Anmeldungen zählen. Abgemeldete Zeilen (z.B. ein
    // Test-Anmelden + wieder Abmelden auf einem Entwurf) blockierten sonst
    // fälschlich die Löschung („Anmeldungen über das Organizer-Team hinaus",
    // obwohl 0 aktive Teilnehmer).
    const ACTIVE = new Set(['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste']);
    let count = 0;
    for (const ev of [event, ...children]) {
      const team = new Set<string>(baseTeam);
      teamEmailSetFor(ev).forEach(e => team.add(e));
      // v30.66: Ein Lesefehler darf hier NICHT zu einer leeren Liste werden.
      // Der Rueckgabewert ist die Loesch-Sperre: 0 heisst "keine fremden
      // Anmeldungen, darf geloescht werden". Wer die Teilnehmerliste nicht lesen
      // kann (403 auf einer Sub-Event-Subsite), bekam bisher genau diese 0 — und
      // damit die Freigabe zum Loeschen eines Events, das sehr wohl Anmeldungen
      // hat. Jetzt schlaegt der Fehler bis zum Aufrufer durch, der ihn als
      // "unbekannt" behandeln MUSS (siehe AdminPage/LandingPage).
      let regs: SPRegistration[] = [];
      let readError = 0;
      regs = await getAllRegistrations(ev.id, (status: number) => { readError = status; });
      if (readError) throw new Error('countExternalRegistrations: Teilnehmerliste nicht lesbar (HTTP ' + readError + ')');
      for (const r of regs) {
        if (!ACTIVE.has(r.Status || '')) continue;
        const email = (r.ParticipantEmail || '').toLowerCase().trim();
        if (email && !team.has(email)) count++;
      }
    }
    return count;
  }

  // v24.6: Organizer-Archiv (pro Person ausblenden) — reiner Anzeige-Filter.
  async function getOrganizerArchivedEventIds(): Promise<Set<string>> {
    if (!currentUserEmail) return new Set<string>();
    return eventService.getOrganizerArchivedEventIds(currentUserEmail);
  }
  async function archiveEventForOrganizer(eventId: string): Promise<boolean> {
    if (!currentUserEmail) return false;
    return eventService.archiveEventForOrganizer(eventId, currentUserEmail);
  }
  async function unarchiveEventForOrganizer(eventId: string): Promise<boolean> {
    if (!currentUserEmail) return false;
    return eventService.unarchiveEventForOrganizer(eventId, currentUserEmail);
  }

  // v26.51: Klartext-Grund des letzten fehlgeschlagenen updateEvent (Service).
  function getLastEventUpdateError(): string {
    try { return eventService.lastUpdateEventError || ''; } catch { return ''; }
  }

  // v30.67: Klartext-Grund des letzten fehlgeschlagenen deleteEvent — damit
  // die Oberfläche die nicht gelöschten Termine namentlich nennen kann statt
  // nur „fehlgeschlagen" (Gegenstück zu getLastEventUpdateError).
  // v30.67 (Review): beide Sprachen halten — der Provider kennt die Locale
  // nicht, die Ansichten (DangerZoneModal, LandingPage) schon.
  const lastDeleteEventErrorRef = React.useRef<{ de: string; en: string }>({ de: '', en: '' });
  function getLastEventDeleteError(lang?: 'de' | 'en'): string {
    return lang === 'en' ? lastDeleteEventErrorRef.current.en : lastDeleteEventErrorRef.current.de;
  }

  // v29.77: opts.skipReload — der Wizard schreibt beim Speichern eines
  // Kalender-Events VIELE Items nacheinander (je Sub-Event Update + Outlook-
  // Dirty-Flags). Bis jetzt zog JEDER dieser Schreibvorgaenge ein volles
  // loadEvents nach sich: alle ~94 Events (28 MB JSON) + Teilnehmerzaehler
  // ueber alle Subsites. Bei 19 Terminen waren das ~20 Komplett-Reloads —
  // DAS hat das Request-Kontingent verbrannt und die 429-Drossel (bis hin
  // zur Nutzer-Sperre) ausgeloest, nicht die 19 kleinen POSTs. Schleifen
  // setzen skipReload und laden am Ende EINMAL.
  async function updateEvent(eventId: string, updates: Record<string, unknown>, opts?: { skipReload?: boolean }): Promise<boolean> {
    // v19.33: Roh-Stand VOR dem Update holen, damit das Audit-Log nur die
    // WIRKLICH geänderten Felder protokolliert (Vorher → Nachher). Vorher loggte
    // es alle Payload-Keys — der Wizard schreibt aber immer den kompletten Payload.
    let oldItem: Record<string, unknown> = {};
    try {
      const raw = await eventService.getEvent(Number(eventId));
      if (raw) oldItem = raw as unknown as Record<string, unknown>;
    } catch { /* Diff bleibt leer, Update läuft trotzdem */ }
    const success = await eventService.updateEvent(Number(eventId), updates);
    if (success) {
      // v9.0/v19.33: Audit-Log (fire-and-forget — UI-Save soll nicht hängen
      // falls SP-ChangeLog-Liste fehlt oder Permissions fehlen). Nur die echten
      // Änderungen werden geloggt; gab es keine, entfällt der Eintrag.
      const ev = events.find(e => e.id === eventId);
      const changes = buildEventUpdateDiff(oldItem, updates);
      if (Object.keys(changes).length > 0) {
        eventService.writeChangeLog({
          action: 'EventUpdated',
          targetType: 'Event',
          targetId: eventId,
          targetName: ev?.title || '',
          eventId: eventId,
          eventTitle: ev?.title || '',
          details: { changes },
        }).catch(() => { /* */ });
      }
      // v9.41: loadEvents im try/catch — wenn ein einzelner Event-Mapping (z.B.
      // ein frisch erstellter Sibling) fehlschlägt, soll das den updateEvent-
      // Erfolg nicht zu einem white-screen-blow-up führen. allSettled in loadEvents
      // selbst sollte das auch schon abfangen, hier nur belt-and-suspenders.
      if (!opts?.skipReload) {
        try { await loadEvents(); } catch (err) { console.warn('[DEX] post-update loadEvents fehlgeschlagen:', err); }
      }
    }
    return success;
  }

  // v30.66: Ausgelagert nach `archiveAndPurge.ts` (Modularisierung Stufe 3).
  // Die Funktionskoerper sind unveraendert; sie bekommen ihre Umgebung
  // ueber ein explizites deps-Objekt statt ueber die Closure.
  const {
    getArchivableCount, runArchiveExpired, getDeletableArchiveCount, runDeleteOldArchive, getParticipantDeletionDue, getParticipantDeletionWarnings, runParticipantDeletion, maybeSendParticipantDeletionWarnings,
  } = makeArchiveActions({ eventService, events, loadEvents, props });

  /** v26.33: Liest das Statistik-Archiv (DEX_EventStats) für die Admin-Kachel. */
  async function getEventStats(): Promise<EventStatsRow[]> {
    if (!eventService) return [];
    try { return await eventService.getEventStats(); }
    catch (e) { console.warn('[DEX] getEventStats failed:', e); return []; }
  }

  async function deleteEvent(eventId: string): Promise<boolean> {
    // v18.3: Demo-Showcase-Event → No-Op (kein SP-Backend). Defense in depth;
    // die UI blendet den Löschen-Button für das Demo-Event ohnehin aus.
    if (isDemoShowcaseId(eventId)) return false;
    // Seit v6.4: Sub-Events sind eigene DEX_Events-Items. Vor dem Löschen des
    // Parent-Events müssen alle Child-Events gelöscht werden, damit auch deren
    // Outlook-Kalendertermine, Subsites und Teilnehmerlisten aufgeräumt werden.
    const ev = events.find(e => e.id === eventId);
    // v30.67: Kinder vom SERVER, nicht aus dem Client-State — der ist auf 100
    // Events gekappt und laesst Zeilen mit Mapping-Fehler aus. Ist die Abfrage
    // nicht lesbar, wird gar nichts geloescht: Ein uebersehenes Kind bliebe
    // sonst mit ParentEventId auf ein geloeschtes Item verwaist zurueck.
    const serverChildIds = ev && !ev.isFictive ? await eventService.getChildEventIds(Number(eventId)) : [];
    if (serverChildIds === null) {
      lastDeleteEventErrorRef.current = {
        de: 'Nicht gelöscht: Die Termine dieses Events konnten nicht gelesen werden — bitte später erneut versuchen.',
        en: 'Not deleted: the dates of this event could not be read — please try again later.',
      };
      console.warn('[DEX] deleteEvent abgebrochen — Kind-Events nicht lesbar', eventId);
      return false;
    }
    const clientChildren = events.filter(e => e.parentEventId === eventId);
    const knownIds = new Set(clientChildren.map(c => c.id));
    const children: DeloitteEvent[] = clientChildren.concat(
      serverChildIds.filter(id => !knownIds.has(String(id))).map(id => ({ id: String(id), title: `#${id}`, parentEventId: eventId } as unknown as DeloitteEvent))
    );
    // v30.67: Das Ergebnis je Kind AUSWERTEN. eventService.deleteEvent wirft
    // nicht, es liefert false (429 auf dem Item-Recycle nach einem Dutzend
    // Terminen) — das catch war toter Code, die Schleife lief durch, und die
    // Klammer wurde trotzdem gelöscht. Die übrigen Kinder blieben mit einem
    // ParentEventId auf ein nicht mehr existierendes Item zurück: Alle
    // Übersichten filtern !parentEventId, die Termine waren in der App
    // unerreichbar, ihre Subsites, Anmeldungen und Outlook-Termine liefen
    // weiter. Scheitert ein Kind, bleibt die Klammer deshalb stehen — dann ist
    // alles noch bedienbar, und der Admin kann es später erneut versuchen.
    const failedChildren: string[] = [];
    const deletedChildren: DeloitteEvent[] = [];
    for (const child of children) {
      let okChild = false;
      try {
        okChild = await eventService.deleteEvent(Number(child.id));
      } catch (err) {
        console.warn('[DEX] Child-Event-Delete fehlgeschlagen:', child.id, err);
      }
      if (okChild) {
        delete subsiteMap.current[child.id];
        deletedChildren.push(child);
      } else {
        failedChildren.push(child.title || child.id);
      }
    }
    // v11.53: vor dem Löschen merken, wie viele aktive Anmeldungen wir
    // vom KPI-Counter abziehen müssen — Parent + alle Children, nur
    // nicht-fictive Events. Wird im Hintergrund einmalig abgezogen.
    const childActive = deletedChildren
      .filter(c => !c.isFictive)
      .reduce((s, c) => s + (c.currentParticipants || 0), 0);
    if (failedChildren.length > 0) {
      lastDeleteEventErrorRef.current = failedChildren.length === 1
        ? {
          de: `Nicht gelöscht: Der Termin „${failedChildren[0]}" konnte nicht gelöscht werden — bitte später erneut löschen.`,
          en: `Not deleted: the date "${failedChildren[0]}" could not be deleted — please try deleting again later.`,
        }
        : {
          de: `Nicht gelöscht: ${failedChildren.length} Termine konnten nicht gelöscht werden (${failedChildren.join(', ')}) — bitte später erneut löschen.`,
          en: `Not deleted: ${failedChildren.length} dates could not be deleted (${failedChildren.join(', ')}) — please try deleting again later.`,
        };
      console.warn('[DEX] deleteEvent abgebrochen —', lastDeleteEventErrorRef.current.de);
      // Die Teilnehmer der bereits gelöschten Termine sind weg — das gehört
      // aus dem KPI heraus, auch wenn die Klammer stehen bleibt.
      if (childActive > 0) eventService.bumpKpiParticipants(-childActive).catch(() => { /* */ });
      await loadEvents();
      return false;
    }
    lastDeleteEventErrorRef.current = { de: '', en: '' };
    const parentActive = (ev && !ev.isFictive) ? (ev.currentParticipants || 0) : 0;
    // v30.67: Nur das Top-Level-Event zählen. createEvent erhöht den Events-
    // KPI seit v23.38 nur für Haupt-/Klammer-Events (`!input.parentEventId`),
    // die Gegenbuchung hier zog aber je Kind eins ab: Nach dem Löschen einer
    // Klammer mit 8 Terminen stand „Events durchgeführt" um 8 zu niedrig (bei
    // mehreren Kalender-Events negativ), bis zufällig ein Admin die App
    // öffnete. recomputeEventKpiOnly zählt ebenfalls nur !parentEventId.
    const childEventsToDecrement = (ev && !ev.isFictive) ? 1 : 0;
    const success = await eventService.deleteEvent(Number(eventId));
    delete subsiteMap.current[eventId];
    if (success) {
      if (childEventsToDecrement > 0) {
        eventService.bumpKpiEvents(-childEventsToDecrement).catch(() => { /* */ });
      }
      const totalActive = childActive + parentActive;
      if (totalActive > 0) {
        eventService.bumpKpiParticipants(-totalActive).catch(() => { /* */ });
      }
    } else {
      lastDeleteEventErrorRef.current = {
        de: 'Das Event konnte nicht gelöscht werden — bitte später erneut versuchen.',
        en: 'The event could not be deleted — please try again later.',
      };
    }
    // Events immer neu laden, auch wenn Subsite-Löschung fehlschlug
    await loadEvents();
    return success;
  }

  /**
   * v11.69: Löscht ausschliesslich das DEX_Events-Listenitem — KEIN Cascade
   * auf Subsite, Teilnehmerliste oder Outlook-DeleteEvent-Queue.
   * Wird genutzt, um ein Sub-Event mit `existingSubsiteUrl` an einer neuen
   * DEX_Events-Zeile wieder anzulegen, damit der `DEX_CreateOutlookEvent`-
   * Flow (GetOnNewItems-Trigger) triggert — die alte Subsite mit allen
   * Teilnehmer-Anmeldungen bleibt unangetastet erhalten.
   */
  async function deleteEventItemOnly(eventId: string): Promise<boolean> {
    const success = await eventService.deleteEventItemOnly(eventId);
    if (success) {
      delete subsiteMap.current[eventId];
      eventService.writeChangeLog({
        action: 'EventItemOnlyDeleted',
        targetType: 'Event',
        targetId: eventId,
        targetName: (events.find(e => e.id === eventId)?.title) || '',
        eventId: eventId,
        eventTitle: (events.find(e => e.id === eventId)?.title) || '',
        details: { reason: 'Outlook-Recreate ohne Subsite-Verlust (v11.69)' },
      }).catch(() => { /* */ });
    }
    return success;
  }

  async function updateMyRegistration(eventId: string, customData: Record<string, string>): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;

    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return false;

    // FieldMap aus Event extrahieren
    const event = events.find(e => e.id === eventId);
    const fieldMap: Record<string, string> = {};
    if (event) {
      for (const f of event.eventSpecificFields) {
        if (f.spInternalName) fieldMap[f.id] = f.spInternalName;
      }
    }

    // Alte Daten und Labels für ChangeLog
    let oldCustomData: Record<string, string> = {};
    try {
      if (myReg.CustomData) oldCustomData = JSON.parse(myReg.CustomData);
    } catch { /* */ }

    const fieldLabelMap: Record<string, string> = {};
    if (event) {
      for (const f of event.eventSpecificFields) {
        fieldLabelMap[f.id] = f.label;
      }
    }

    const success = await eventService.updateRegistrationData(subsiteUrl, myReg.Id, customData, fieldMap, oldCustomData, fieldLabelMap);
    if (success) await loadEvents();
    return success;
  }

  // v30.66: Ausgelagert nach `assistantAndProxy.ts` (Modularisierung Stufe 3).
  // Die Funktionskoerper sind unveraendert; sie bekommen ihre Umgebung
  // ueber ein explizites deps-Objekt statt ueber die Closure.
  const {
    delegateRegistrationToAssistant, recordProxyDelegation, getMyAssistantLinks, requestAssistantChange, resolveAssistantRequest,
  } = makeAssistantActions({ eventService, events, subsiteMap, currentUserEmail, currentUserName, childEventsOf });

  // v30.66: Ausgelagert nach `participantFiles.ts` (Modularisierung Stufe 3).
  // Die Funktionskoerper sind unveraendert; sie bekommen ihre Umgebung
  // ueber ein explizites deps-Objekt statt ueber die Closure.
  const {
    getAllParticipants, getEventNumbersForEmail, getMyEventNumbers, refreshEvents, listMyEventAttachments, uploadMyEventAttachment, deleteMyEventAttachment, uploadFieldDocument, listFieldDocuments, deleteFieldDocument,
  } = makeParticipantFileActions({ eventService, subsiteMap, currentUserEmail, loadEvents });

  // v10.27: Split-Capacity-Gruppen-Wechsel — wrappt EventService.switchSplitGroup,
  // ergänzt um Mail/Outlook-Sideeffects und Reload.
  async function switchSplitGroup(eventId: string, newType: 'Durchstarter' | 'Funstarter'): Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' | 'Failed'; full: boolean }> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return { ok: false, status: 'Failed', full: false };
    const event = events.find(e => e.id === eventId);
    if (!event) return { ok: false, status: 'Failed', full: false };
    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return { ok: false, status: 'Failed', full: false };
    const result = await eventService.switchSplitGroup(
      subsiteUrl,
      myReg.Id,
      newType,
      event.durchstarterCapacity || 0,
      event.funstarterCapacity || 0,
    );
    if (result.ok) {
      // Audit-Log + Mail/Outlook anstoßen — analog zu cancelRegistration.
      eventService.writeChangeLog({
        action: 'ParticipantSwitchedGroup',
        targetType: 'Participant',
        targetId: currentUserEmail,
        targetName: currentUserName,
        eventId, eventTitle: event.title,
        details: { from: myReg.StarterType || myReg.PreferredStarterType || '', to: newType, finalStatus: result.status },
      }).catch(() => { /* */ });
      if (!event.disableEmails) {
        try {
          const lang = event.emailLanguage || 'EN';
          const isDeMail = lang.toUpperCase() === 'DE';
          const labelA = (event.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
          const labelB = (event.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
          const newLabel = newType === 'Durchstarter' ? labelA : labelB;
          // v13.0: Mail aus DB-Template — Confirmed bzw. Waitlist-Variante.
          const isWaitlist = result.status === 'Warteliste';
          const templateType = isWaitlist ? 'GroupSwitchWaitlist' : 'GroupSwitchConfirmed';
          const tpl = await eventService.getEmailTemplate(templateType, lang).catch(() => null);
          const firstName = currentUserName.split(',').pop()?.trim().split(' ')[0] || currentUserName.split(' ')[0];
          const vars: Record<string, string> = {
            Name: firstName,
            GroupLabel: newLabel,
            EventTitle: event.title,
            AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
          };
          let mail: { subject: string; body: string };
          if (tpl) {
            mail = buildEmailFromTemplate(tpl, vars);
          } else {
            const innerBody = isDeMail
              ? (isWaitlist
                ? `<p>Du hast den Wechsel in die Gruppe <strong>${newLabel}</strong> für <strong>${event.title}</strong> angefragt. Diese Gruppe ist aktuell voll, daher steht deine Anmeldung auf der <strong>Warteliste der Gruppe ${newLabel}</strong>.</p>`
                : `<p>Dein Gruppen-Wechsel zu <strong>${newLabel}</strong> für <strong>${event.title}</strong> ist bestätigt.</p>`)
              : (isWaitlist
                ? `<p>You requested to switch to the <strong>${newLabel}</strong> group for <strong>${event.title}</strong>. The group is currently full, so your registration is on the <strong>${newLabel} waitlist</strong>.</p>`
                : `<p>Your group switch to <strong>${newLabel}</strong> for <strong>${event.title}</strong> is confirmed.</p>`);
            mail = {
              subject: isDeMail
                ? (isWaitlist ? `Gruppen-Wechsel — auf Warteliste: ${event.title}` : `Gruppen-Wechsel bestätigt: ${event.title}`)
                : (isWaitlist ? `Group switch — added to waitlist: ${event.title}` : `Group switch confirmed: ${event.title}`),
              body: wrapTemplate(isWaitlist ? '#ed8b00' : '#86bc25', isDeMail ? 'Gruppen-Wechsel' : 'Group switch', event.title, innerBody),
            };
          }
          await eventService.queueEmail(mail.subject, currentUserEmail, currentUserName, mail.body, templateType, event.title, eventId)
            .catch(err => console.warn('[DEX] switchSplitGroup mail failed:', err));
        } catch { /* */ }
      }
      // v22.20: Nach jedem Gruppenwechsel einen ID-Reorder anstoßen — der Flow
      // zieht die Nummerierung sofort glatt (Wechsler korrekt einsortiert) und
      // besetzt einen in der ALTEN Gruppe frei gewordenen Platz direkt nach,
      // statt erst bei der nächsten Abmeldung. Name/E-Mail des Wechslers gehen
      // mit, damit Nachrück-Mail + Audit den Platz-Vorgänger korrekt benennen.
      try {
        await eventService.queueIDReorder(eventId, event.eventNumber, subsiteUrl, event.title, currentUserName, currentUserEmail);
      } catch (err) { console.warn('[DEX] queueIDReorder (group switch) failed:', err); }
      await loadEvents();
    }
    return result;
  }

  /**
   * Admin-Cleanup: Events mit Status='Active' + EndDate < jetzt auf 'Completed' setzen.
   * Anschliessend wird die Event-Liste neu geladen, damit die UI die neuen Status sieht.
   */
  async function markExpiredEventsAsCompleted(): Promise<number> {
    const n = await eventService.markExpiredEventsAsCompleted();
    if (n > 0) await loadEvents();
    return n;
  }

  /**
   * Anfrage von der Landing Page an die DEX-Admins. Verwendet DEX_SEND_MAIL via
   * der DEX_Emails-Queue, mit dem Anfrager im Cc-Feld. Body wird ins Deloitte-
   * Template (grüner Header, Footer) gewrappt.
   */
  // v12.12: Re-Seed-Aktion durchreichen.
  async function reseedDefaultEmailTemplates(): Promise<ReseedSummary> {
    return eventService.reseedDefaultEmailTemplates();
  }

  async function getAllEmailTemplates(): Promise<Array<{ id: number; templateType: string; language: string; subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string }>> {
    return eventService.getAllEmailTemplates();
  }
  async function updateEmailTemplate(id: number, fields: { subject?: string; heading?: string; subheading?: string; headingColor?: string; bodyHtml?: string }): Promise<boolean> {
    return eventService.updateEmailTemplate(id, fields);
  }

  // v30.66: Ausgelagert nach `mails.ts` (Modularisierung Stufe 3).
  // Die Funktionskoerper sind unveraendert; sie bekommen ihre Umgebung
  // ueber ein explizites deps-Objekt statt ueber die Closure.
  const {
    sendCompleteRegistrationReminder, sendAdminInquiry, notifyAdminsExternalAudienceAccess, sendOrganizerOnboarding,
  } = makeMailActions({ eventService });


  // v13.11: Demo-Impersonation kann den synthetischen Demo-User per
  // qrScannerEmails einem konkreten Event als Check-In-Helfer
  // zuordnen. Wir injizieren die Demo-Mail in das gewählte Event,
  // damit Header/StartPage/AdminPage die übliche Permission-Logik
  // unverändert nutzen können.
  // v22.23: Organizer-Tutorial — Demo-Showcase-Event auch OHNE Impersonation
  // injizieren, solange die Tour läuft (TutorialGuide schaltet das Flag).
  const [tutorialDemoActive, setTutorialDemoActive] = React.useState(false);
  const eventsForConsumer = React.useMemo(() => {
    try {
      if (typeof window === 'undefined') return events;
      const raw = window.localStorage?.getItem('dex_demo_impersonation');
      if (!raw) {
        if (tutorialDemoActive && !events.some(e => e.isDemoShowcase)) {
          const storedLocale = window.localStorage?.getItem('dex-locale');
          const demoLocale: 'de' | 'en' = storedLocale === 'en' ? 'en' : 'de';
          // Eingeloggten User als Organizer eintragen, damit das Demo-Event
          // auch für Nicht-Admins in der Organizer-Center-Liste auftaucht.
          const demo = buildDemoShowcaseEvents(demoLocale).map(e => ({
            ...e,
            organizerEmails: Array.from(new Set([...(e.organizerEmails || []), currentUserEmail].filter(Boolean))),
          }));
          return [...demo, ...events];
        }
        return events;
      }
      const payload = JSON.parse(raw);
      // v17.25: Im Demo-Impersonation-Modus das synthetische Showcase-Event
      // (+ Sub-Event) vorne in die Liste hängen, damit der Admin auf der
      // Register-Seite alle Event-Fähigkeiten durchspielen kann. Existiert
      // nur client-seitig — kein SP-Roundtrip. Sprache aus dem persistierten
      // Locale-Key, Fallback DE.
      let withDemo = events;
      try {
        const storedLocale = window.localStorage?.getItem('dex-locale');
        const demoLocale: 'de' | 'en' = storedLocale === 'en' ? 'en' : 'de';
        if (!events.some(e => e.isDemoShowcase)) {
          withDemo = [...buildDemoShowcaseEvents(demoLocale), ...events];
        }
      } catch { /* */ }
      const targetId: string = payload?.checkInEventId || '';
      const demoEmail: string = (payload?.email || '').toLowerCase();
      if (!targetId || !demoEmail) return withDemo;
      return withDemo.map(e => {
        if (e.id !== targetId) return e;
        const current = (e.qrScannerEmails || []).map(x => x.toLowerCase());
        if (current.indexOf(demoEmail) >= 0) return e;
        return { ...e, qrScannerEmails: [...(e.qrScannerEmails || []), demoEmail] };
      });
    } catch { return events; }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [events, tutorialDemoActive]);

  // v22.42: Automatischer Hintergrund-Fix der Zeilen-Sichtbarkeit. Beim
  // Admin-Start werden Fremd-Anmeldungen (von Assistenz/Organizer für andere
  // angelegt) so repariert, dass der Teilnehmer Autor seiner Zeile wird und
  // sie damit in „Meine Events" sieht + sich selbst abmelden kann. Läuft
  // gedrosselt (1×/24h via localStorage), sequentiell über alle aktiven
  // Subsites, komplett best-effort — der Admin merkt nichts davon.
  // v30.16: Offene Klammer-Nachzüge aus localStorage abarbeiten (s.
  // utils/shadowHeal). Entsteht, wenn der Klammer-Schreibvorgang einer
  // Anmeldung trotz Wiederholungen an der Drosselung scheiterte und der
  // Browser zuging, bevor die Hintergrund-Kette der Erfolgsseite fertig war.
  // Läuft einmal pro Session, 20 s nach dem Boot (nicht in die Boot-Lastspitze),
  // sequentiell; registerForEvent ist für die Klammer idempotent.
  const shadowHealRanRef = React.useRef(false);
  // v30.67: Timer-Id im Ref, Abräumen NUR beim Unmount. Die alte Cleanup
  // `return () => clearTimeout(t)` gehörte zum Effect-Lauf und lief bei JEDER
  // neuen `events`-Referenz — und loadEvents liefert beim Boot garantiert zwei
  // davon (setEvents(mapped), wenige Sekunden später setEvents(withCounts)),
  // lange vor Ablauf der 20 s. Der Guard `shadowHealRanRef` verhinderte
  // danach das Neu-Armen: Der Merker sagte „gelaufen", die Cleanup hatte das
  // Feuern verhindert — zusammen „nie". Die Klammer-Zeilen wurden in keiner
  // Sitzung nachgeholt, der Merker verfiel still nach 14 Tagen. Dieselbe
  // Lehre steht seit v23.12 in DexEventPlatform („KEIN clearTimeout-Cleanup
  // zurückgeben").
  const shadowHealTimerRef = React.useRef<number | null>(null);
  React.useEffect(() => {
    if (shadowHealRanRef.current) return;
    if (!events || events.length === 0) return;
    if (readPendingShadowParents().length === 0) { shadowHealRanRef.current = true; return; }
    shadowHealRanRef.current = true;
    shadowHealTimerRef.current = window.setTimeout(() => {
      void (async () => {
        for (const p of readPendingShadowParents()) {
          try {
            const r = await registerForEvent(
              p.eventId, p.customData || {}, p.firstName || '', p.lastName || '', p.email, undefined,
              { ...(p.proxy ? { proxyConsentConfirmed: true, actorAllowedAsAssistant: true } : {}), suppressMail: true, suppressOutlook: true, skipReload: true }
            );
            if (r.ok) removePendingShadowParent(p.eventId, p.email);
          } catch { /* bleibt im Merker — nächster App-Start versucht es wieder (14-Tage-Verfall) */ }
        }
      })();
    }, 20000);
    // Bewusst KEINE Cleanup hier — s. Kommentar am shadowHealTimerRef.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [events]);
  React.useEffect(() => () => {
    if (shadowHealTimerRef.current !== null) window.clearTimeout(shadowHealTimerRef.current);
  }, []);


  const autoFixStartedRef = React.useRef(false);
  // v30.66: Ausgelagert nach `maintenance.ts` (Modularisierung Stufe 3).
  // Die Funktionskoerper sind unveraendert; sie bekommen ihre Umgebung
  // ueber ein explizites deps-Objekt statt ueber die Closure.
  const {
    autoRepairProxyAccess, fixAllEventColumns, repairAllOrganizerPermissions, restoreCustomFieldDescriptions,
  } = makeMaintenanceActions({ eventService, events, autoFixStartedRef, updateEvent, loadEvents });

  // v30.66: Ausgelagert nach `autoMails.ts` (Modularisierung Stufe 3).
  // Die Funktionskoerper sind unveraendert; sie bekommen ihre Umgebung
  // ueber ein explizites deps-Objekt statt ueber die Closure.
  const {
    maybeSendPostEventOrganizerMails, maybeSendWeeklyReport,
  } = makeAutoMailActions({ eventService, events, currentUserEmail, props });

  // v30.66: Ausgelagert nach `billing.ts` (Modularisierung Stufe 3).
  // Die Funktionskoerper sind unveraendert; sie bekommen ihre Umgebung
  // ueber ein explizites deps-Objekt statt ueber die Closure.
  const {
    getFAConfig, saveFAConfig, sendFAMail, logFAContact, sendBundledUpdateMail, markEventSettled, saveFAPersonalNumbers, maybeSendBillingAutoMails,
  } = makeBillingActions({ eventService, events, setEvents, currentUserEmail, currentUserName, getAllRegistrations });

  // v30.66: Ausgelagert nach `organizerRoles.ts` (Modularisierung Stufe 3).
  // Die Funktionskoerper sind unveraendert; sie bekommen ihre Umgebung
  // ueber ein explizites deps-Objekt statt ueber die Closure.
  const {
    requestOrganizerRole, requestCoOrganizerApprovals, notifyNewCoOrganizers, getOpenOrganizerRequests, markOrganizerRequestDecided,
  } = makeOrganizerRoleActions({ eventService, currentUserEmail, currentUserName, props });

  // v30.66: Ausgelagert nach `inactiveAccounts.ts` (Modularisierung Stufe 3).
  // Die Funktionskoerper sind unveraendert; sie bekommen ihre Umgebung
  // ueber ein explizites deps-Objekt statt ueber die Closure.
  const {
    scanInactiveAccounts, getSentInactiveNotices, notifyOrganizerOfInactive, autoDeregisterInactive, getEventComms,
  } = makeInactiveAccountActions({ eventService, events, currentUserEmail, currentUserName, loadEvents });

  return React.createElement(
    EventContext.Provider,
    {
      value: {
        events: eventsForConsumer,
        topLevelEvents: eventsForConsumer.filter(e => !e.parentEventId),
        childEventsOf, isEventsLoading, ensureEventDocuments, refreshEventDocuments,
        createEvent, registerForEvent, registerTeam,
        getTeamMembers: async (eventId: string, teamId: string): Promise<SPRegistration[]> => {
          const subsiteUrl = subsiteMap.current[eventId];
          if (!subsiteUrl) return [];
          return eventService.getTeamMembers(subsiteUrl, teamId);
        },
        addTeamMember,
        assignTeamlessToTeam, notifyExistingTeamMembers,
        joinTeam,
        transferTeamLead,
        createTeamJoinRequest,
        listTeamJoinRequestsForEvent,
        decideTeamJoinRequest,
        listOpenTeamsForEvent,
        cancelRegistration,
        declineEvent,
        cancelTeamMember,
        getMyRegistration, getMyProxyRegistrations, cancelProxyRegistration, updateProxyRegistration, handBackToParticipant, delegateRegistrationToAssistant, recordProxyDelegation, getMyAssistantLinks, requestAssistantChange, resolveAssistantRequest, selfCheckIn, setTutorialDemoActive, checkRegistrationByEmail, getAllRegistrations, deleteEvent, countExternalRegistrations, getOrganizerArchivedEventIds, archiveEventForOrganizer, unarchiveEventForOrganizer, deleteEventItemOnly, updateEvent, getLastEventUpdateError, getLastEventDeleteError, updateMyRegistration, switchSplitGroup, listMyEventAttachments, uploadMyEventAttachment, deleteMyEventAttachment, uploadFieldDocument, listFieldDocuments, deleteFieldDocument, getMyEventNumbers, getEventNumbersForEmail, getAllParticipants, refreshEvents, refreshParticipantCounts, getLiveCounterStats, reconcileCounters, subscribeEventRealtime, markExpiredEventsAsCompleted, autoRepairProxyAccess, maybeSendWeeklyReport, getFAConfig, saveFAConfig, sendFAMail, markEventSettled, saveFAPersonalNumbers, sendBundledUpdateMail, logFAContact, maybeSendBillingAutoMails, maybeSendPostEventOrganizerMails, scanInactiveAccounts, notifyOrganizerOfInactive, autoDeregisterInactive, getEventComms, getSentInactiveNotices, getArchivableCount, runArchiveExpired, getDeletableArchiveCount, runDeleteOldArchive, getParticipantDeletionWarnings, getParticipantDeletionDue, runParticipantDeletion, maybeSendParticipantDeletionWarnings, getEventStats, fixAllEventColumns, repairAllOrganizerPermissions, restoreCustomFieldDescriptions,
        sendAdminInquiry,
        sendCompleteRegistrationReminder,
        notifyAdminsExternalAudienceAccess,
        requestOrganizerRole, requestCoOrganizerApprovals, notifyNewCoOrganizers, getOpenOrganizerRequests, markOrganizerRequestDecided,
        getOrganizerRequestDetails: (id: number) => eventService.getOrganizerRequestDetails(id),
        reseedDefaultEmailTemplates,
        getAllEmailTemplates,
        updateEmailTemplate,
        sendOrganizerOnboarding,
        getKpiCache: () => eventService.getKpiCache(),
        updateKpiCache: (v) => eventService.updateKpiCache(v),
        recomputeEventKpiOnly: () => eventService.recomputeEventKpiOnly(),
        getKpiTotals: () => eventService.getKpiTotals(),
      },
    },
    props.children
  );
}

export function useEvents(): EventContextType {
  const ctx = React.useContext(EventContext);
  if (!ctx) throw new Error('useEvents must be used within EventProvider');
  return ctx;
}
