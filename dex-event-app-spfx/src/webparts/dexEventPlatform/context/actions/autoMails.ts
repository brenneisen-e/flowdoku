/**
 * Automatik-Mails: „Danke, wir hoffen es lief gut" nach dem Event und der
 * woechentliche Admin-Bericht.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktionskoerper sind unveraendert - statt aus der Provider-Closure
 * beziehen sie ihre Umgebung aus dem `deps`-Objekt.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DeloitteEvent } from '../../types';
import { EventService } from '../../services/EventService';
import { wrapTemplate } from '../../services/EmailTemplates';
import { isEventOver } from '../../utils/eventFormat';
import { RELEASE_NOTES, splitReleaseNote } from '../../data/releaseNotes';

export interface AutoMailDeps {
  eventService: EventService;
  events: DeloitteEvent[];
  currentUserEmail: string;
  props: { context: WebPartContext };
}

export function makeAutoMailActions(deps: AutoMailDeps) {
  const { eventService, events, currentUserEmail, props } = deps;

  // v23.8: Wöchentlicher Admin-Bericht. Wird beim App-Start (nur Admins,
  // gedrosselt) angestoßen. Quelle der Wahrheit für „wann lief der letzte
  // Bericht" ist die SP-Liste DEX_WeeklyReports — erst wenn der letzte Bericht
  // ≥ 7 Tage her ist, wird ein neuer geschrieben (Vergleichszeitraum = seit dem
  // letzten Bericht; allererster Bericht = letzte 7 Tage). Inhalt: neue Events
  // (+ von wem), Anmeldungen im Zeitraum, neu ernannte Organizer — plus
  // Gesamt-KPIs (Events insgesamt, aktive Anmeldungen insgesamt, hinterlegte
  // Organizer). Best-effort; Fehler blocken nie den Boot.
  // v24.2: „Danke, wir hoffen es lief gut"-Mail an den Organizer, wenn er die
  // App nach dem Event-Tag öffnet. Einmal pro Event und Organizer (localStorage-
  // Drosselung pro Browser). Enthält den Hinweis auf die 1-jährige Aufbewahrung
  // der Teilnehmerübersicht (Datenschutz) + Verweis auf Excel-Export / App.
  async function maybeSendPostEventOrganizerMails(): Promise<void> {
    try {
      const meLc = (currentUserEmail || '').toLowerCase();
      if (!meLc) return;
      const overEvents = (events || []).filter(e => {
        if (e.parentEventId) return false; // nur Hauptevents
        if (e.isFictive) return false;     // keine Entwürfe
        const isOrg = (e.organizerEmails || []).some(x => (x || '').toLowerCase() === meLc)
          || (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === meLc);
        if (!isOrg) return false;
        return isEventOver(e);
      });
      if (overEvents.length === 0) return;
      // v26.39: Persistenter Dedup-Marker (DEX_PostEventMails) statt der
      // transienten DEX_Emails-Queue — die wurde nach Versand + Archivierung
      // wieder leer, wodurch die Mail beim nächsten App-Öffnen ERNEUT rausging.
      const newlyCreated = await eventService.ensurePostEventMailsList().catch(() => false);
      if (newlyCreated) {
        // Erststart nach dem Update: ALLE bereits abgelaufenen (sichtbaren)
        // Events als „erledigt" markieren, damit sie KEINE (weitere) Mail
        // bekommen. Nur ab jetzt neu ablaufende Events erhalten die Mail — dann
        // genau einmal. So endet die Wiederhol-Schleife ohne Extra-Versand.
        const overAll = (events || []).filter(e => !e.parentEventId && !e.isFictive && isEventOver(e));
        for (const e of overAll) { try { await eventService.recordPostEventMail(e.id); } catch { /* */ } }
        return;
      }
      const sentIds = await eventService.getPostEventMailSentEventIds().catch(() => new Set<string>());
      let appUrl = '';
      try { appUrl = `${props.context.pageContext.web.absoluteUrl}/SitePages/DEX.aspx`; } catch { appUrl = ''; }
      for (const ev of overEvents) {
        if (sentIds.has(String(ev.id))) continue; // schon verschickt (persistenter Marker)
        const key = `dex_posteventmail_${ev.id}`;
        let already = false;
        try { already = !!window.localStorage.getItem(key); } catch { already = false; }
        if (already) continue;
        // v26.12: an ALLE Organizer (Haupt- + Co-Organizer), nicht nur an den,
        // der die App öffnet. Doppelversand-Schutz serverseitig über die Queue
        // (DEX_Emails) — so feuert auch bei mehreren Organizern nur EINE Mail.
        const orgEmails = Array.from(new Set(
          [...(ev.organizerEmails || []), ...(ev.coOrganizerEmails || [])]
            .map(x => (x || '').trim())
            .filter(x => x.indexOf('@') > 0)
            .map(x => x)
        ));
        if (orgEmails.length === 0) continue;
        // Dedupe case-insensitiv (eine Adresse nicht doppelt im To).
        const seen = new Set<string>();
        const recipients = orgEmails.filter(e => { const lc = e.toLowerCase(); if (seen.has(lc)) return false; seen.add(lc); return true; });
        const linkLine = appUrl
          ? `<p style="margin:0 0 12px;">Ihr findet die Teilnehmerübersicht jederzeit im <a href="${appUrl}" style="color:#86bc25;font-weight:600;">Organizer Center der DEX App</a> — dort könnt ihr sie auch als Excel exportieren.</p>`
          : `<p style="margin:0 0 12px;">Ihr findet die Teilnehmerübersicht jederzeit im Organizer Center der DEX App — dort könnt ihr sie auch als Excel exportieren.</p>`;
        const inner = `
          <p style="margin:0 0 12px;">Hallo zusammen,</p>
          <p style="margin:0 0 12px;">wir hoffen, euer Event <strong>&bdquo;${ev.title}&ldquo;</strong> ist gut verlaufen und alle hatten eine schöne Zeit!</p>
          <p style="margin:0 0 12px;">Ein kurzer Hinweis zur Aufbewahrung: Die <strong>Teilnehmerübersicht bleibt noch 3 Monate gespeichert</strong> (Datenschutz-/Aufbewahrungsvorgabe). Danach wird sie gelöscht — das Event und die wichtigsten Kennzahlen bleiben im Statistik-Archiv erhalten. Ihr werdet rund eine Woche vorher noch einmal erinnert.</p>
          ${linkLine}
          <p style="margin:0 0 12px;">Vielen Dank, dass ihr das Event organisiert habt!</p>`;
        const body = wrapTemplate('#86bc25', 'Danke für euer Event!', ev.title, inner);
        try {
          await eventService.queueEmail(
            `Dein Event „${ev.title}" — danke & Hinweis zur Aufbewahrung`,
            recipients.join('; '), recipients.join('; '), body, 'PostEventOrganizer', ev.title, ev.id,
          );
          // v26.39: persistenten Marker setzen — verhindert erneuten Versand,
          // auch nachdem die DEX_Emails-Zeile längst archiviert/gelöscht wurde.
          await eventService.recordPostEventMail(ev.id);
          try { window.localStorage.setItem(key, String(Date.now())); } catch { /* */ }
        } catch { /* einzelne Mail-Fehler ignorieren */ }
      }
    } catch (e) { console.warn('[DEX] post-event organizer mail failed:', e); }
  }

  async function maybeSendWeeklyReport(opts?: { force?: boolean }): Promise<{ sent: boolean; admins: number; reason?: string }> {
    const force = !!opts?.force;
    try {
      const SEVEN = 7 * 24 * 60 * 60 * 1000;
      const now = new Date();
      const last = await eventService.getLastWeeklyReport();
      if (!force) {
        // v23.38: Fälligkeit AUSSCHLIESSLICH am Server-Eintrag (DEX_WeeklyReports)
        // festmachen. Der frühere per-Browser-localStorage-Backstop konnte die
        // automatische Auslösung dauerhaft blockieren (ein alter „lastsent"-
        // Merker aus einem früheren Versuch, bei dem nie eine Mail rausging) —
        // das war der Grund, warum nie ein Bericht kam. Ohne Server-Eintrag (oder
        // wenn der letzte ≥ 7 Tage her ist) wird jetzt zuverlässig gesendet; der
        // Doppelversand-Schutz pro App-Session liegt im Boot-Effekt (didWeeklyReport).
        const serverLastTs = (last && last.created) ? new Date(last.created).getTime() : 0;
        if (serverLastTs && (now.getTime() - serverLastTs) < SEVEN) return { sent: false, admins: 0, reason: 'not-due' }; // noch keine 7 Tage
      }
      // v23.38: pro Admin eine eigene Mail mit ECHTEM Namen + persönlicher
      // Anrede (vorher eine Sammelmail mit generischem „Admin").
      const adminRecipients = await eventService.getRoleRecipients('Admin');
      if (adminRecipients.length === 0) return { sent: false, admins: 0, reason: 'no-admins' }; // niemand zum Versenden
      const periodFrom = (last && last.periodTo) ? new Date(last.periodTo) : new Date(now.getTime() - SEVEN);
      const fromIso = periodFrom.toISOString();
      const toIso = now.toISOString();

      // Daten sammeln.
      const [newEvents, newOrganizers, organizerEmails] = await Promise.all([
        eventService.getEventsCreatedSince(fromIso),
        eventService.getRoleItemsCreatedSince('Organizer', fromIso),
        eventService.getRoleEmails('Organizer'),
      ]);
      // Anmeldungen über alle (deduplizierten) Subsites zählen.
      const subs = new Set<string>();
      for (const e of events) { if (e.subsiteUrl) subs.add(e.subsiteUrl); }
      let regSince = 0; let regTotal = 0;
      for (const sub of Array.from(subs)) {
        try { const c = await eventService.countRegistrations(sub, fromIso); regSince += c.since; regTotal += c.total; }
        catch { /* einzelne Subsite-Fehler ignorieren */ }
      }
      const totalEvents = events.length;

      // E-Mail bauen — bewusst anwenderfreundlich formuliert (ganze Sätze,
      // Singular/Plural, „(noch im Entwurf)" hinter Entwurfs-Events).
      const fmtD = (iso: string): string => { try { return new Date(iso).toLocaleDateString('de-DE'); } catch { return iso.slice(0, 10); } };
      const esc = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const draftBadge = '<span style="margin-left:6px;padding:1px 7px;border-radius:8px;background:#fff3e0;color:#b35a00;font-size:11px;font-weight:600;">noch im Entwurf</span>';
      const eventsHtml = newEvents.length > 0
        ? `<p style="margin:6px 0 0;">${newEvents.length === 1 ? 'In dieser Woche wurde <strong>1 neues Event</strong> angelegt:' : `In dieser Woche wurden <strong>${newEvents.length} neue Events</strong> angelegt:`}</p>
           <ul style="margin:8px 0 0 18px;padding:0;">${newEvents.map(e => `<li style="margin-bottom:6px;"><strong>${esc(e.title)}</strong>${e.isDraft ? '&nbsp;' + draftBadge : ''}<br><span style="color:#777;font-size:13px;">angelegt von ${esc(e.author)} am ${fmtD(e.created)}</span></li>`).join('')}</ul>`
        : '<p style="margin:6px 0 0;color:#666;">In dieser Woche wurde kein neues Event angelegt.</p>';
      const orgHtml = newOrganizers.length > 0
        ? `<p style="margin:6px 0 0;">${newOrganizers.length === 1 ? 'Eine Person wurde neu als Organizer berechtigt:' : `${newOrganizers.length} Personen wurden neu als Organizer berechtigt:`}</p>
           <ul style="margin:8px 0 0 18px;padding:0;">${newOrganizers.map(o => `<li style="margin-bottom:4px;">${esc(o.email)} <span style="color:#777;font-size:13px;">(seit ${fmtD(o.created)})</span></li>`).join('')}</ul>`
        : '<p style="margin:6px 0 0;color:#666;">In dieser Woche wurde niemand neu als Organizer berechtigt.</p>';
      // v23.44: Release Notes der Berichtswoche (aus dem TS-Modul, datums-gefiltert).
      const relFromTs = new Date(fromIso).getTime();
      const relToTs = new Date(toIso).getTime();
      const relNotes = RELEASE_NOTES.filter(n => { const t = new Date(n.date).getTime(); return isFinite(t) && t >= relFromTs && t <= relToTs; });
      const relNotesHtml = relNotes.length > 0
        // v30.60: Auch im Wochenbericht gegliedert statt als Textblock —
        // dieselbe Zerlegung wie in der Release-Notes-Tabelle. Ein Bericht,
        // der aus fünf zehnzeiligen Absätzen besteht, wird nicht gelesen.
        ? `<ul style="margin:8px 0 0 18px;padding:0;">${relNotes.map(n => {
          const parts = splitReleaseNote(n.text);
          const inner = parts.points.length > 0
            ? `${parts.lead ? esc(parts.lead) : ''}<ul style="margin:4px 0 0 16px;padding:0;">${parts.points.map(pt => `<li style="margin-bottom:3px;">${esc(pt)}</li>`).join('')}</ul>`
            : `${esc(parts.lead)}${parts.rest ? ` ${esc(parts.rest)}` : ''}`;
          return `<li style="margin-bottom:8px;"><strong>${n.type === 'Bugfix' ? 'Behoben' : 'Neu'}:</strong> ${inner}</li>`;
        }).join('')}</ul>`
        : '';
      const draftCount = newEvents.filter(e => e.isDraft).length;
      const draftHint = draftCount > 0
        ? `<p style="margin:6px 0 0;color:#777;font-size:13px;">Davon ${draftCount === 1 ? 'ist 1 Event noch ein Entwurf' : `sind ${draftCount} Events noch Entwürfe`} und damit für Teilnehmer noch nicht sichtbar.</p>`
        : '';

      // v23.36: Events, die seit dem letzten Bericht von Entwurf auf LIVE
      // gewechselt sind (Snapshot der Entwurfs-IDs aus dem letzten Bericht).
      // Nur Top-Level (Klammern/Hauptevents).
      const lastDraftIds = (last && last.draftEventIds) || [];
      const wentLive = events.filter(e => !e.isFictive && !e.parentEventId && lastDraftIds.indexOf(String(e.id)) >= 0);
      const wentLiveHtml = wentLive.length > 0
        ? `<h3 style="margin:20px 0 0;font-size:15px;color:#2d4a06;">Live gegangene Events</h3>
           <p style="margin:6px 0 0;">${wentLive.length === 1 ? 'Ein zuvor als Entwurf gespeichertes Event ist jetzt <strong>live</strong> und für Teilnehmer sichtbar:' : `${wentLive.length} zuvor als Entwurf gespeicherte Events sind jetzt <strong>live</strong> und für Teilnehmer sichtbar:`}</p>
           <ul style="margin:8px 0 0 18px;padding:0;">${wentLive.map(e => `<li style="margin-bottom:4px;"><strong>${esc(e.title)}</strong></li>`).join('')}</ul>`
        : '';

      // v23.36: Events, die IM ZEITRAUM stattgefunden haben — Klammern/
      // Hauptevents als EINE Zeile mit Auflistung der Sub-Events + Teilnehmerzahl.
      const fromTs = new Date(fromIso).getTime();
      const toTs = new Date(toIso).getTime();
      const shortSub = (t: string): string => { const p = (t || '').split('|'); return (p.length > 1 ? p[p.length - 1] : t).trim(); };
      const evDateTs = (e: DeloitteEvent): number => {
        const end = e.endDate ? new Date(e.endDate).getTime() : 0;
        const start = e.startDate ? new Date(e.startDate).getTime() : 0;
        return (isFinite(end) && end) ? end : ((isFinite(start) && start) ? start : 0);
      };
      const cnt = (e: DeloitteEvent): number => Math.max(0, e.currentParticipants || 0);
      const heldRows = events.filter(e => !e.parentEventId && !e.isFictive).map(e => {
        const kids = events.filter(c => c.parentEventId === e.id);
        let ts = evDateTs(e);
        if (!ts && kids.length) ts = Math.max(0, ...kids.map(evDateTs));
        return { e, kids, ts };
      }).filter(r => r.ts >= fromTs && r.ts <= toTs).sort((a, b) => a.ts - b.ts);
      const heldHtml = heldRows.length > 0
        ? `<p style="margin:6px 0 0;">${heldRows.length === 1 ? 'In diesem Zeitraum hat <strong>1 Event</strong> stattgefunden:' : `In diesem Zeitraum haben <strong>${heldRows.length} Events</strong> stattgefunden:`}</p>
           <ul style="margin:8px 0 0 18px;padding:0;">${heldRows.map(r => `<li style="margin-bottom:6px;"><strong>${esc(r.e.title)}</strong> — ${cnt(r.e)} Teilnehmer <span style="color:#777;font-size:13px;">(${fmtD(new Date(r.ts).toISOString())})</span>${r.kids.length ? `<ul style="margin:4px 0 0 18px;padding:0;color:#555;font-size:13px;">${r.kids.map(k => `<li style="margin-bottom:2px;">${esc(shortSub(k.title))} — ${cnt(k)} Teilnehmer</li>`).join('')}</ul>` : ''}</li>`).join('')}</ul>`
        : '<p style="margin:6px 0 0;color:#666;">In diesem Zeitraum hat kein Event stattgefunden.</p>';

      // v26: Tickets im Berichtszeitraum (Fragen & Antworten der Nutzer).
      const tInPeriod = (iso: string): boolean => { const x = new Date(iso).getTime(); return isFinite(x) && x >= fromTs && x <= toTs; };
      let ticketsInPeriod: Array<{ askerName: string; questions: string[]; status: string; answerText: string; answeredByName: string }> = [];
      try {
        const allTickets = await eventService.getTickets();
        ticketsInPeriod = allTickets.filter(tk => tInPeriod(tk.created) || (!!tk.answeredAt && tInPeriod(tk.answeredAt)));
      } catch { /* best-effort */ }
      const tStatusLabel = (s: string): string => (s === 'Closed' ? 'beantwortet' : (s === 'InProgress' ? 'in Bearbeitung' : 'offen'));
      const answeredInPeriod = ticketsInPeriod.filter(tk => tk.status === 'Closed').length;
      const ticketsHtml = ticketsInPeriod.length > 0
        ? `<p style="margin:6px 0 0;">In diesem Zeitraum ${ticketsInPeriod.length === 1 ? 'wurde <strong>1 Frage</strong> gestellt' : `wurden <strong>${ticketsInPeriod.length} Fragen</strong> gestellt`}${answeredInPeriod > 0 ? `, davon ${answeredInPeriod} beantwortet` : ''}:</p>
           <ul style="margin:8px 0 0 18px;padding:0;">${ticketsInPeriod.map(tk => `<li style="margin-bottom:8px;"><strong>${esc(tk.askerName)}</strong> <span style="color:#777;font-size:13px;">(${tStatusLabel(tk.status)})</span><br><span style="color:#333;">${esc((tk.questions || []).join(' / '))}</span>${tk.status === 'Closed' && tk.answerText ? `<br><span style="color:#2d4a06;font-size:13px;">Antwort${tk.answeredByName ? ` (${esc(tk.answeredByName)})` : ''}: ${esc(tk.answerText)}</span>` : ''}</li>`).join('')}</ul>`
        : '<p style="margin:6px 0 0;color:#666;">In diesem Zeitraum wurden keine Fragen über das Ticketsystem gestellt.</p>';
      const relNotesBottomHtml = relNotes.length > 0
        ? `<p style="margin:6px 0 0;color:#555;">Diese Verbesserungen sind in diesem Zeitraum live gegangen:</p>${relNotesHtml}`
        : '<p style="margin:6px 0 0;color:#666;">In diesem Zeitraum gab es keine neuen Änderungen an der App.</p>';

      const inner = `
        <p style="margin:0 0 6px;">Hallo __GREETING_NAME__,</p>
        <p style="margin:0 0 16px;">hier ist euer wöchentlicher Überblick über die DEX Event Experience Platform — was in den letzten Tagen passiert ist, vom <strong>${fmtD(fromIso)}</strong> bis <strong>${fmtD(toIso)}</strong>.</p>
        <h3 style="margin:18px 0 0;font-size:15px;color:#2d4a06;">Neue Events</h3>
        ${eventsHtml}
        ${draftHint}
        ${wentLiveHtml}
        <h3 style="margin:20px 0 0;font-size:15px;color:#2d4a06;">Stattgefundene Events</h3>
        ${heldHtml}
        <h3 style="margin:20px 0 0;font-size:15px;color:#2d4a06;">Neue Anmeldungen</h3>
        <p style="margin:6px 0 0;">In diesem Zeitraum ${regSince === 1 ? 'ist <strong>1 Anmeldung</strong> eingegangen.' : `sind <strong>${regSince} Anmeldungen</strong> eingegangen.`}</p>
        <h3 style="margin:20px 0 0;font-size:15px;color:#2d4a06;">Neue Organizer</h3>
        ${orgHtml}
        <hr style="border:none;border-top:1px solid #e0e0e0;margin:24px 0;">
        <h3 style="margin:0 0 6px;font-size:15px;color:#2d4a06;">Die Plattform auf einen Blick</h3>
        <p style="margin:0 0 8px;color:#777;font-size:13px;">Gesamtzahlen über alle Events hinweg:</p>
        <table style="border-collapse:collapse;font-size:14px;">
          <tr><td style="padding:3px 16px 3px 0;color:#555;">Events insgesamt</td><td style="padding:3px 0;font-weight:700;">${totalEvents}</td></tr>
          <tr><td style="padding:3px 16px 3px 0;color:#555;">Aktive Anmeldungen insgesamt</td><td style="padding:3px 0;font-weight:700;">${regTotal}</td></tr>
          <tr><td style="padding:3px 16px 3px 0;color:#555;">Berechtigte Organizer</td><td style="padding:3px 0;font-weight:700;">${organizerEmails.length}</td></tr>
        </table>
        <hr style="border:none;border-top:1px solid #e0e0e0;margin:24px 0;">
        <h3 style="margin:0 0 6px;font-size:15px;color:#2d4a06;">Fragen &amp; Antworten</h3>
        ${ticketsHtml}
        <hr style="border:none;border-top:1px solid #e0e0e0;margin:24px 0;">
        <h3 style="margin:0 0 6px;font-size:15px;color:#2d4a06;">Release Notes</h3>
        ${relNotesBottomHtml}
        <p style="margin:24px 0 0;font-size:12px;color:#999;">Diesen Bericht bekommt ihr automatisch einmal pro Woche, weil ihr Admin der DEX Event Experience Platform seid.</p>
      `;
      const subject = `Automatischer Wochenbericht — ${fmtD(toIso)}`;
      // v23.38: EINE Sammelmail an ALLE Admins (alle Adressen im To,
      // Semikolon-getrennt — der DEX_SEND_MAIL-Flow mappt Recipient direkt aufs
      // To-Feld). RecipientName = echte Namen der Admins (statt generisch
      // „Admin"); Anrede generisch „Hallo zusammen". EventId='0' (gültiger
      // OData-Filter im Flow; '' würde den ganzen Lauf abbrechen → keine Mail).
      const toAll = adminRecipients.map(a => a.email).join('; ');
      const namesAll = adminRecipients.map(a => a.name || a.email).join('; ');
      const groupInner = inner.replace('__GREETING_NAME__', 'zusammen');
      const body = wrapTemplate('#86bc25', 'Automatischer Wochenbericht', `${fmtD(fromIso)} – ${fmtD(toIso)}`, groupInner);
      const queued = await eventService.queueEmail(subject, toAll, namesAll, body, 'WeeklyReport', '', '0');
      if (queued) {
        // v23.36: aktuellen Entwurfs-Snapshot mitspeichern, damit der NÄCHSTE
        // Bericht „live gegangen" erkennt (Entwurf jetzt → live beim nächsten Mal).
        const currentDraftIds = events.filter(e => e.isFictive && !e.parentEventId).map(e => String(e.id));
        await eventService.recordWeeklyReport(fromIso, toIso, currentDraftIds);
        return { sent: true, admins: adminRecipients.length };
      }
      console.warn('[DEX] Wochenbericht: queueEmail meldete Misserfolg — kein Protokolleintrag, Retry beim nächsten Boot.');
      return { sent: false, admins: adminRecipients.length, reason: 'queue-failed' };
    } catch (err) {
      console.warn('[DEX] maybeSendWeeklyReport error:', err);
      return { sent: false, admins: 0, reason: 'error' };
    }
  }

  return { maybeSendPostEventOrganizerMails, maybeSendWeeklyReport };
}
