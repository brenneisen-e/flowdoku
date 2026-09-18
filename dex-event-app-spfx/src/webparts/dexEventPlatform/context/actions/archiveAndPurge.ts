/**
 * Archivierung abgelaufener Events (DEX_Archive) und das 3-Monats-
 * Loeschkonzept fuer Teilnehmerdaten samt Vorwarn-Mails.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktionskoerper sind unveraendert - statt aus der Provider-Closure
 * beziehen sie ihre Umgebung aus dem `deps`-Objekt.
 */

import { eventHeaderImageOpts } from '../../utils/mailHeaderImage';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DeloitteEvent } from '../../types';
import { EventService } from '../../services/EventService';
import { buildMailButton, wrapTemplate } from '../../services/EmailTemplates';
import { feedbackSchonDa } from '../../utils/dexFeedback';

export interface ArchiveDeps {
  eventService: EventService;
  events: DeloitteEvent[];
  loadEvents: () => Promise<void>;
  props: { context: WebPartContext };
}

export function makeArchiveActions(deps: ArchiveDeps) {
  const { eventService, events, loadEvents, props } = deps;

  // ==================== v21: Archivierung ====================
  // Abgelaufen = End-Datum (Fallback Start-Datum) liegt in der Vergangenheit.
  // Liefert Event-Ids + Subsite-URLs + Titel-Map für die Archiv-Funktionen.
  function getExpiredEventSets(): { ids: Set<string>; subs: Set<string>; titles: Record<string, string>; allIds: Set<string>; allSubs: Set<string> } {
    const now = Date.now();
    const ids = new Set<string>();
    const subs = new Set<string>();
    const titles: Record<string, string> = {};
    // v23.39: zusätzlich ALLE aktuell existierenden Event-IDs/Subsites — damit
    // die Archivierung Zeilen GELÖSCHTER Events (verwaist) ebenfalls erfasst
    // (deren Bezug taucht in allIds/allSubs nicht mehr auf).
    const allIds = new Set<string>();
    const allSubs = new Set<string>();
    for (const e of events) {
      allIds.add(String(e.id));
      if (e.subsiteUrl) allSubs.add(e.subsiteUrl.toLowerCase());
      const endRef = e.endDate || e.startDate;
      const t = endRef ? new Date(endRef).getTime() : 0;
      if (t > 0 && t < now) {
        ids.add(String(e.id));
        titles[String(e.id)] = e.title || '';
        if (e.subsiteUrl) subs.add(e.subsiteUrl.toLowerCase());
      }
    }
    return { ids, subs, titles, allIds, allSubs };
  }

  /** v21: Zählt archivreife Zeilen (Queue-/Log-Listen abgelaufener Events). */
  async function getArchivableCount(): Promise<{ total: number; perList: Record<string, number> }> {
    if (!eventService) return { total: 0, perList: {} };
    const { ids, subs, allIds, allSubs } = getExpiredEventSets();
    // v23.39: Events müssen geladen sein (allIds>0), sonst würde die
    // Verwaist-Erkennung ALLES als gelöscht ansehen. Archivreif sind jetzt
    // Zeilen abgelaufener UND gelöschter (verwaister) Events.
    if (allIds.size === 0) return { total: 0, perList: {} };
    return eventService.countArchivableRows(ids, subs, allIds, allSubs);
  }

  /** v21: Verschiebt alle archivreifen Zeilen ins DEX_Archive (Admin).
   *  v22.2: shouldCancel = Abbruch-Check aus dem Fortschrittsmodal. */
  async function runArchiveExpired(
    onProgress?: (listIdx: number, listTotal: number, listName: string, done: number, total: number) => void,
    shouldCancel?: () => boolean
  ): Promise<{ archived: number; failed: number; cancelled: boolean; perList: Record<string, number>; errors: string[] }> {
    if (!eventService) return { archived: 0, failed: 0, cancelled: false, perList: {}, errors: [] };
    const { ids, subs, titles, allIds, allSubs } = getExpiredEventSets();
    return eventService.archiveExpiredRows(ids, subs, titles, onProgress, shouldCancel, allIds, allSubs);
  }

  // v23.40: Löschkonzept — Stichdatum.
  // v23.48: Frist auf 1 Monat (Wunsch Maintainer) — die archivierten Daten
  // (DEX_Emails/DEX_Outlook/… im DEX_Archive) sollen rund einen Monat nach
  // Ablauf des Events weg. Da sofort archiviert wird, fällt ArchivedAt mit dem
  // Event-Ablauf zusammen.
  function archiveDeleteCutoffIso(): string {
    const d = new Date();
    d.setMonth(d.getMonth() - 1);
    return d.toISOString();
  }
  async function getDeletableArchiveCount(): Promise<number> {
    if (!eventService) return 0;
    return eventService.countDeletableArchiveRows(archiveDeleteCutoffIso());
  }
  async function runDeleteOldArchive(
    onProgress?: (done: number, total: number) => void,
    shouldCancel?: () => boolean
  ): Promise<{ deleted: number; failed: number; cancelled: boolean }> {
    if (!eventService) return { deleted: 0, failed: 0, cancelled: false };
    return eventService.deleteOldArchiveRows(archiveDeleteCutoffIso(), onProgress, shouldCancel);
  }

  // ====================================================================
  // v26.32: Löschkonzept — Teilnehmerliste 3 Monate nach Event-Ende löschen.
  // 3 Mon. − 1 Woche: Vorwarn-Mail an alle Organizer (Download-Hinweis) +
  // Landing-Hinweis. Ab 3 Mon.: Admin löscht per Button die Teilnehmerliste;
  // Event bleibt in DEX_Events, KPIs wandern ins DEX_EventStats-Archiv.
  // ====================================================================
  const PARTICIPANT_RETENTION_MONTHS = 3;
  const PARTICIPANT_WARN_LEAD_MS = 7 * 24 * 60 * 60 * 1000;
  function participantDeleteDueTs(e: DeloitteEvent): number {
    const endRef = e.endDate || e.startDate;
    if (!endRef) return 0;
    const d = new Date(endRef);
    if (isNaN(d.getTime())) return 0;
    d.setMonth(d.getMonth() + PARTICIPANT_RETENTION_MONTHS);
    return d.getTime();
  }
  // Nur echte Haupt-Events mit eigener Teilnehmer-Subsite, keine Entwürfe.
  function isParticipantDeletionCandidate(e: DeloitteEvent): boolean {
    return !e.parentEventId && !e.isFictive && !!e.subsiteUrl;
  }

  /** Events, deren Teilnehmerliste fällig ist (≥3 Mon.) und noch NICHT archiviert. */
  // v26.35: Kandidaten = ab 1 Woche VOR der 3-Monats-Frist (dann läuft die
  // Vorwarnung), noch mit Subsite, kein Entwurf, noch nicht archiviert.
  function participantDeletionCandidates(now: number): DeloitteEvent[] {
    return (events || []).filter(e => {
      const due = participantDeleteDueTs(e);
      return isParticipantDeletionCandidate(e) && due > 0 && (due - PARTICIPANT_WARN_LEAD_MS) <= now;
    });
  }

  /** Löschbar ERST, wenn die Vorwarn-Mail an die Organizer raus ist UND danach
   *  mindestens 1 Woche vergangen ist — sonst könnte gelöscht werden, ohne dass
   *  je gewarnt wurde. */
  async function getParticipantDeletionDue(): Promise<DeloitteEvent[]> {
    if (!eventService) return [];
    const now = Date.now();
    const cands = participantDeletionCandidates(now);
    if (cands.length === 0) return [];
    let archived = new Set<number>();
    try { archived = await eventService.getArchivedStatsEventNumbers(); } catch { archived = new Set(); }
    let warnDates: Record<string, string> = {};
    try { warnDates = await eventService.getParticipantDeletionWarningDates(); } catch { warnDates = {}; }
    return cands.filter(e => {
      if (archived.has(e.eventNumber)) return false;
      const sent = warnDates[String(e.id)];
      const sentTs = sent ? new Date(sent).getTime() : 0;
      return sentTs > 0 && (sentTs + PARTICIPANT_WARN_LEAD_MS) <= now;
    });
  }

  /** Events, deren Löschung ansteht, aber die 1-Wochen-Frist seit der Vorwarnung
   *  noch NICHT abgelaufen ist (bzw. die Warnung noch aussteht). */
  async function getParticipantDeletionWarnings(): Promise<DeloitteEvent[]> {
    if (!eventService) return [];
    const now = Date.now();
    const cands = participantDeletionCandidates(now);
    if (cands.length === 0) return [];
    let archived = new Set<number>();
    try { archived = await eventService.getArchivedStatsEventNumbers(); } catch { archived = new Set(); }
    let warnDates: Record<string, string> = {};
    try { warnDates = await eventService.getParticipantDeletionWarningDates(); } catch { warnDates = {}; }
    return cands.filter(e => {
      if (archived.has(e.eventNumber)) return false;
      const sent = warnDates[String(e.id)];
      const sentTs = sent ? new Date(sent).getTime() : 0;
      // Noch NICHT löschbar → im Warn-/Wartezustand.
      return !(sentTs > 0 && (sentTs + PARTICIPANT_WARN_LEAD_MS) <= now);
    });
  }

  /** Admin-Aktion: KPIs archivieren, dann Teilnehmerliste löschen (Event bleibt). */
  async function runParticipantDeletion(
    onProgress?: (done: number, total: number, label: string) => void
  ): Promise<{ deleted: number; failed: number }> {
    if (!eventService) return { deleted: 0, failed: 0 };
    const due = await getParticipantDeletionDue();
    let deleted = 0, failed = 0;
    for (let i = 0; i < due.length; i++) {
      const e = due[i];
      if (onProgress) onProgress(i, due.length, e.title || '');
      try {
        // Erst KPIs sichern — nur bei Erfolg die (unwiderrufliche) Löschung starten.
        const orgNames = Array.from(new Set(
          [...(e.organizers || []), ...(e.coOrganizerNames || [])].map(x => (x || '').trim()).filter(Boolean)
        )).join(', ');
        const archivedOk = await eventService.archiveEventStats({
          eventNumber: e.eventNumber, eventTitle: e.title, eventType: e.type,
          location: e.location, startDate: e.startDate, endDate: e.endDate,
          maxParticipants: e.maxParticipants, subsiteUrl: e.subsiteUrl,
          organizer: orgNames,
        });
        if (!archivedOk) { failed++; continue; }
        const delOk = await eventService.deleteParticipantData(Number(e.id));
        if (delOk) { deleted++; }
        else {
          // Löschung fehlgeschlagen → Statistik-Zeile zurückrollen, damit das
          // Event beim nächsten Lauf erneut verarbeitet wird. Die Subsite
          // existiert noch (Löschung schlug fehl), daher sind die dann neu
          // berechneten KPIs korrekt — sonst bliebe die Teilnehmerliste (PII)
          // verwaist zurück (Archiv-Eintrag würde sie fälschlich als „erledigt"
          // markieren).
          try { await eventService.deleteEventStatsRow(e.eventNumber); } catch { /* */ }
          failed++;
        }
      } catch { failed++; }
    }
    if (onProgress) onProgress(due.length, due.length, '');
    await loadEvents();
    return { deleted, failed };
  }

  /** Auto-Vorwarnung an alle Organizer (einmalig, Queue-entdoppelt). Wird vom
   *  Admin-App-Open auf der Landing Page ausgelöst. */
  async function maybeSendParticipantDeletionWarnings(): Promise<void> {
    try {
      if (!eventService) return;
      const warns = await getParticipantDeletionWarnings();
      if (warns.length === 0) return;
      let appUrl = '';
      try { appUrl = `${props.context.pageContext.web.absoluteUrl}/SitePages/DEX.aspx`; } catch { appUrl = ''; }
      for (const ev of warns) {
        const key = `dex_participantdelwarn_${ev.id}`;
        let already = false;
        try { already = !!window.localStorage.getItem(key); } catch { already = false; }
        if (already) continue;
        const orgEmails = Array.from(new Set(
          [...(ev.organizerEmails || []), ...(ev.coOrganizerEmails || [])]
            .map(x => (x || '').trim()).filter(x => x.indexOf('@') > 0)
        ));
        if (orgEmails.length === 0) continue;
        const seen = new Set<string>();
        const recipients = orgEmails.filter(e => { const lc = e.toLowerCase(); if (seen.has(lc)) return false; seen.add(lc); return true; });
        let alreadyQueued = false;
        try { alreadyQueued = await eventService.hasQueuedEmail('ParticipantDeletionWarning', ev.id); } catch { alreadyQueued = false; }
        if (alreadyQueued) { try { window.localStorage.setItem(key, String(Date.now())); } catch { /* */ } continue; }
        const linkLine = appUrl
          ? `<p style="margin:0 0 12px;">Ihr könnt die Teilnehmerübersicht jetzt noch im <a href="${appUrl}" style="color:#86bc25;font-weight:600;">Organizer Center der DEX App</a> ansehen und als Excel exportieren.</p>`
          : `<p style="margin:0 0 12px;">Ihr könnt die Teilnehmerübersicht jetzt noch im Organizer Center der DEX App ansehen und als Excel exportieren.</p>`;
        /*
         * v31.31: Auch diese Mail fragt nach Feedback (Nutzer-Ansage
         * 14.09.2026). Sie ist der zweite — und letzte — Anlass, an dem das
         * Event noch einmal auf den Tisch kommt; wer die Mail direkt nach dem
         * Event weggeklickt hat, ist hier drei Monate klüger.
         *
         * ABER nur, wenn für dieses Event noch niemand geantwortet hat. Zum
         * zweiten Mal nach etwas zu fragen, das schon beantwortet ist, liest
         * sich wie „ist nicht angekommen" — und genau das war es nicht.
         * Antworten liegen als Piggyback `_feedback` am Event (utils/dexFeedback).
         */
        const feedbackUrl = appUrl && !feedbackSchonDa(ev.emailTemplateOverrides)
          ? `${appUrl}?env=WebView#action=feedback&e=${encodeURIComponent(String(ev.id))}`
          : '';
        const feedbackBlock = feedbackUrl ? `
          <p style="margin:0 0 10px;">Und wenn ihr zwei Minuten habt: <strong>Wie lief es für euch mit DEX?</strong> Überwiegend zum Anklicken — das hilft uns bei dem, was wir als Nächstes bauen.</p>
          ${buildMailButton(feedbackUrl, 'Feedback geben &rarr;')}` : '';
        const inner = `
          <p style="margin:0 0 12px;">Hallo zusammen,</p>
          <p style="margin:0 0 12px;">für euer Event <strong>&bdquo;${ev.title}&ldquo;</strong> läuft die Aufbewahrungsfrist der Teilnehmerliste ab: <strong>in etwa einer Woche wird die Teilnehmerliste gelöscht</strong> (3 Monate nach dem Event, Datenschutz-/Aufbewahrungsvorgabe).</p>
          <p style="margin:0 0 12px;">Bitte <strong>ladet euch die Liste jetzt herunter</strong>, falls ihr sie noch braucht. Das Event und die wichtigsten Kennzahlen bleiben danach im Statistik-Archiv erhalten.</p>
          ${linkLine}
          ${feedbackBlock}
          <p style="margin:0 0 12px;">Vielen Dank!</p>`;
        const body = wrapTemplate('#86bc25', 'Teilnehmerliste wird bald gelöscht', ev.title, inner, undefined, eventHeaderImageOpts(ev.emailTemplateOverrides, ev.mailImageBase64));
        try {
          await eventService.queueEmail(
            `Teilnehmerliste zu „${ev.title}" wird in ~1 Woche gelöscht`,
            recipients.join('; '), recipients.join('; '), body, 'ParticipantDeletionWarning', ev.title, ev.id,
          );
          try { window.localStorage.setItem(key, String(Date.now())); } catch { /* */ }
        } catch { /* einzelne Mail-Fehler ignorieren */ }
      }
    } catch (e) { console.warn('[DEX] participant deletion warning mail failed:', e); }
  }

  // ====================================================================
  // v31.63: Automatischer Lauf — Archivieren, Archiv aufräumen, fällige
  // Teilnehmerlisten löschen — OHNE Rückfrage, im Hintergrund.
  //
  // Nutzer-Ansage 15.09.2026: „Automatismus mit Archivierung und Löschen
  // beim Aufruf durch Admin (aktuell muss ich das manuell bestätigen) …
  // ich will, dass es einfach im Hintergrund läuft und gemacht wird."
  //
  // Warum das jetzt vertretbar ist und vorher nicht: Bis v31.61 löschten
  // die drei Schritte hart (REST-DELETE). Seit v31.62 gehen Archivzeilen und
  // Teilnehmer-Subsites in den Papierkorb, und die Teilnehmerlöschung ist
  // ohnehin doppelt abgesichert (Vorwarn-Mail an die Organizer, danach eine
  // Woche Frist — `getParticipantDeletionDue`). Die drei Rückfragen im
  // Landing-Kasten und im Admin-Hub bleiben als Handweg bestehen; der
  // Automat ruft dieselben Funktionen, nur ohne Dialog.
  //
  // Reihenfolge = die der Handbedienung: erst Arbeitslisten ins Archiv,
  // dann alte Archivzeilen weg, dann fällige Teilnehmerlisten. Jeder Schritt
  // ist für sich best-effort; ein Fehler im ersten sperrt den zweiten nicht.
  // Der Fortschritt zählt in „Einheiten" (eine Zeile = eine Einheit, eine
  // Teilnehmerliste = eine Einheit), damit oben rechts EINE Prozentzahl
  // steht und nicht drei.
  // ====================================================================
  async function runAutoMaintenance(
    onProgress?: (p: AutoMaintenanceProgress) => void,
    opts?: { nurNeuesSeit?: AutoMaintenanceResult }
  ): Promise<AutoMaintenanceResult> {
    const out: AutoMaintenanceResult = {
      archived: 0, archiveFailed: 0, deleted: 0, deleteFailed: 0,
      participantsDeleted: 0, participantsFailed: 0, nothingToDo: false, error: '',
      archiveErrors: [], finishedAt: 0,
    };
    if (!eventService) { out.nothingToDo = true; return out; }
    const report = (phase: AutoMaintenanceProgress['phase'], doneUnits: number, totalUnits: number, label: string): void => {
      if (!onProgress) return;
      const pct = totalUnits > 0 ? Math.max(0, Math.min(100, Math.round((doneUnits / totalUnits) * 100))) : 0;
      onProgress({ phase, done: doneUnits, total: totalUnits, pct, label });
    };
    try {
      report('zaehlen', 0, 0, '');
      const [arch, delCount, due] = await Promise.all([
        getArchivableCount().catch(() => ({ total: 0, perList: {} as Record<string, number> })),
        getDeletableArchiveCount().catch(() => 0),
        getParticipantDeletionDue().catch(() => [] as DeloitteEvent[]),
      ]);
      const totalUnits = arch.total + delCount + due.length;
      if (totalUnits === 0) { out.nothingToDo = true; report('fertig', 0, 0, ''); return out; }
      // v31.73: Innerhalb der 6-Stunden-Sperre läuft der Automat nur, wenn
      // seit dem letzten Lauf etwas NEU fällig geworden ist. Nutzer-Befund
      // 17.09.2026 („auch TN-Liste löschen Automatismus"): Der Lauf um 9 Uhr
      // hatte 36 archiviert und 310 Archivzeilen entfernt; die Teilnehmerliste
      // des Frühlingsfests wurde erst danach fällig (Vorwarnung + 1 Woche) und
      // blieb bis zum nächsten Lauf mit Knopf im Kasten stehen. Maßstab je
      // Kategorie ist die Fehlzahl des letzten Laufs: Liegen nicht mehr Zeilen
      // an, als damals liegen blieben, sind es dieselben — und die wieder und
      // wieder anzufassen bringt nichts. Liegt in EINER Kategorie mehr an,
      // läuft der ganze Lauf (die Hängenbleiber kosten dann je einen Versuch).
      const alt = opts && opts.nurNeuesSeit;
      if (alt && arch.total <= alt.archiveFailed && delCount <= alt.deleteFailed && due.length <= alt.participantsFailed) {
        out.nothingToDo = true; report('fertig', 0, 0, ''); return out;
      }

      // Schritt 1: Arbeitslisten → DEX_Archive. `runArchiveExpired` meldet
      // je Liste (done/total) — die Summe über die Listen ist der Fortschritt.
      let base = 0;
      if (arch.total > 0) {
        const doneByList: number[] = [];
        const r1 = await runArchiveExpired((listIdx, _listTotal, listName, done) => {
          doneByList[listIdx] = done;
          const sum = doneByList.reduce((a, b) => a + (b || 0), 0);
          report('archiv', Math.min(sum, arch.total), totalUnits, listName);
        });
        out.archived = r1.archived; out.archiveFailed = r1.failed; out.archiveErrors = r1.errors;
        if (r1.failed > 0) console.warn(`[DEX] Automatische Archivierung: ${r1.failed} Zeile(n) nicht verschoben —`, r1.errors);
      }
      base = arch.total;
      report('archiv', base, totalUnits, '');

      // Schritt 2: Archivzeilen älter als 1 Monat → Papierkorb (v31.62).
      if (delCount > 0) {
        const r2 = await runDeleteOldArchive((done) => report('archivloeschen', base + Math.min(done, delCount), totalUnits, ''));
        out.deleted = r2.deleted; out.deleteFailed = r2.failed;
      }
      base += delCount;
      report('archivloeschen', base, totalUnits, '');

      // Schritt 3: fällige Teilnehmerlisten (KPIs ins Statistik-Archiv, dann
      // Subsite in den Papierkorb). `runParticipantDeletion` liest die
      // Fälligen selbst noch einmal — die Vorwarn-Frist wird also nicht aus
      // unserer Zählung von eben übernommen, sondern frisch geprüft.
      if (due.length > 0) {
        const r3 = await runParticipantDeletion((done, _total, label) => report('teilnehmer', base + Math.min(done, due.length), totalUnits, label));
        out.participantsDeleted = r3.deleted; out.participantsFailed = r3.failed;
      }
      report('fertig', totalUnits, totalUnits, '');
    } catch (e) {
      out.error = e instanceof Error ? e.message : String(e);
      console.warn('[DEX] Automatische Archivierung/Löschung fehlgeschlagen:', e);
    }
    out.finishedAt = Date.now();
    // Protokoll: Was der Automat getan hat, steht im Änderungsprotokoll —
    // ohne Dialog gibt es sonst keinen Beleg, dass und wann gelaufen ist.
    // v31.68: samt der ersten Fehlgründe je Zeile (`archivFehlerDetails`).
    try {
      await eventService.writeChangeLog({
        action: 'AutoMaintenanceRun',
        targetType: 'Other',
        targetName: 'Archivierung & Löschen (automatisch)',
        details: {
          archiviert: out.archived, archivFehler: out.archiveFailed,
          archivFehlerDetails: out.archiveErrors.length ? out.archiveErrors : undefined,
          archivGeloescht: out.deleted, archivLoeschFehler: out.deleteFailed,
          teilnehmerlistenGeloescht: out.participantsDeleted, teilnehmerlistenFehler: out.participantsFailed,
          fehler: out.error || undefined,
        },
      });
    } catch { /* Protokoll ist best-effort */ }
    return out;
  }

  return { getArchivableCount, runArchiveExpired, getDeletableArchiveCount, runDeleteOldArchive, getParticipantDeletionDue, getParticipantDeletionWarnings, runParticipantDeletion, maybeSendParticipantDeletionWarnings, runAutoMaintenance };
}

/** v31.63: Fortschritt des automatischen Laufs (Abzeichen oben rechts). */
export interface AutoMaintenanceProgress {
  phase: 'zaehlen' | 'archiv' | 'archivloeschen' | 'teilnehmer' | 'fertig';
  /** Erledigte bzw. gesamte Einheiten (Zeilen + Teilnehmerlisten). */
  done: number;
  total: number;
  pct: number;
  /** Aktuelle Liste bzw. aktueller Event-Titel — nur zur Anzeige. */
  label: string;
}

/** v31.63: Ergebnis des automatischen Laufs. */
export interface AutoMaintenanceResult {
  archived: number; archiveFailed: number;
  deleted: number; deleteFailed: number;
  participantsDeleted: number; participantsFailed: number;
  /** true = nichts stand an; es wurde nichts verändert und nichts protokolliert. */
  nothingToDo: boolean;
  error: string;
  /** v31.68: die ersten Fehlgründe der Archivierung je Zeile (max. 12). */
  archiveErrors: string[];
  /** v31.68: Zeitstempel des Abschlusses (0 = nicht gelaufen). */
  finishedAt: number;
}
