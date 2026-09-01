/**
 * Inaktive Konten (Scan, Organizer-Hinweis, Auto-Abmeldung) und das
 * Kommunikations-Log eines Events.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktionskoerper sind unveraendert - statt aus der Provider-Closure
 * beziehen sie ihre Umgebung aus dem `deps`-Objekt.
 */

import { DeloitteEvent } from '../../types';
import { EventService, SPRegistration, EventCommRow } from '../../services/EventService';
import { buildHashDeepLink } from '../../utils/deepLink';
import { wrapTemplate } from '../../services/EmailTemplates';

export interface InactiveAccountDeps {
  eventService: EventService;
  events: DeloitteEvent[];
  currentUserEmail: string;
  currentUserName: string;
  loadEvents: () => Promise<void>;
}

export function makeInactiveAccountActions(deps: InactiveAccountDeps) {
  const { eventService, events, currentUserEmail, currentUserName, loadEvents } = deps;

  // v22.45: Scan über mehrere Events — pro Event Teilnehmer laden und auf ein
  // aktives Deloitte-Konto prüfen. Für die Landing-Page-Warnung (Organizer/
  // Admin). Sequentiell + best-effort; der Aufrufer (LandingPage) drosselt
  // den Aufruf via localStorage (1×/24h) und übergibt nur Events, die der
  // User auch lesen darf (eigene bzw. — als Admin — alle).
  async function scanInactiveAccounts(
    evs: Array<{ id: string; title: string; subsiteUrl?: string }>
  ): Promise<Array<{ eventId: string; title: string; people: Array<{ email: string; name: string }> }>> {
    const out: Array<{ eventId: string; title: string; people: Array<{ email: string; name: string }> }> = [];
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    for (const ev of evs) {
      const sub = (ev.subsiteUrl || '').trim();
      if (!sub) continue;
      try {
        const regs = await eventService.getAllRegistrations(sub);
        const active = regs.filter(r => ACTIVE.indexOf(r.Status) >= 0);
        const emails = Array.from(new Set(active.map(r => (r.ParticipantEmail || '').trim().toLowerCase()).filter(Boolean)));
        if (emails.length === 0) continue;
        const res = await eventService.checkAccountsActive(emails);
        if (!res.ok || res.inactive.length === 0) continue;
        const people = res.inactive.map(em => {
          const reg = active.find(r => (r.ParticipantEmail || '').trim().toLowerCase() === em);
          const name = reg
            ? ((reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : (reg.ParticipantName || em))
            : em;
          return { email: em, name };
        });
        out.push({ eventId: ev.id, title: ev.title, people });
      } catch (err) { console.warn('[DEX] scanInactiveAccounts failed for', ev.id, err); }
    }
    return out;
  }

  // v24.59: Bereits benachrichtigte Konten eines Events (lowercase E-Mails) —
  // best-effort-Durchgriff auf den Service.
  async function getSentInactiveNotices(eventId: string): Promise<Set<string>> {
    try { return await eventService.getSentInactiveNotices(eventId); }
    catch { return new Set<string>(); }
  }

  // v24.51: Den/die Organizer eines Events per Mail darauf hinweisen, dass eine
  // angemeldete Person womöglich kein aktives Deloitte-Konto mehr hat. Dedup
  // über DEX_InactiveNotices (ReadSecurity=1) — pro Event+Person nur EINE Mail,
  // egal welcher Admin den Button klickt.
  async function notifyOrganizerOfInactive(
    eventId: string,
    people: Array<{ email: string; name: string }>
  ): Promise<{ sent: number; skipped: number; noOrganizer?: boolean }> {
    const event = events.find(e => e.id === eventId);
    if (!event) return { sent: 0, skipped: 0 };
    const orgEmails = Array.from(new Set((event.organizerEmails || []).concat(event.coOrganizerEmails || [])
      .map(e => (e || '').trim()).filter(Boolean)));
    if (orgEmails.length === 0) return { sent: 0, skipped: people.length, noOrganizer: true };
    let sentSet = new Set<string>();
    try { sentSet = await eventService.getSentInactiveNotices(eventId); } catch { sentSet = new Set(); }
    const newPeople = people.filter(p => !sentSet.has((p.email || '').toLowerCase().trim()));
    if (newPeople.length === 0) return { sent: 0, skipped: people.length };
    // Marker setzen (claim) — danach versenden.
    for (const p of newPeople) {
      try { await eventService.recordInactiveNotice(eventId, (p.email || '').toLowerCase().trim()); } catch { /* */ }
    }
    const isDe = !(event.emailLanguage === 'EN');
    const esc = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const list = newPeople.map(p => `<li><strong>${esc(p.name || p.email)}</strong>${p.name ? ` (${esc(p.email)})` : ''}</li>`).join('');
    const heading = isDe ? 'Möglicherweise inaktives Konto' : 'Possibly inactive account';
    const sub = event.title;
    const appUrl = buildHashDeepLink(`${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`, { action: 'admin', event: eventId });
    const body = isDe
      ? `<p>Hallo,</p>
         <p>bei deinem Event <strong>${esc(event.title)}</strong> ${newPeople.length === 1 ? 'gibt es eine angemeldete Person, die' : 'gibt es angemeldete Personen, die'} <strong>womöglich Deloitte verlassen ${newPeople.length === 1 ? 'hat' : 'haben'}</strong> (kein aktives Konto mehr). An ${newPeople.length === 1 ? 'diese Person' : 'diese Personen'} kommen Bestätigungs-Mails und Outlook-Termine <strong>möglicherweise nicht an</strong>:</p>
         <ul>${list}</ul>
         <p>Bitte prüfe das im Organizer Center und melde die ${newPeople.length === 1 ? 'Person' : 'Personen'} bei Bedarf ab.</p>
         <p><a href="${appUrl}" style="display:inline-block;background:#86bc25;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:600;">Event im Organizer Center öffnen</a></p>
         <p>Danke!</p>`
      : `<p>Hello,</p>
         <p>at your event <strong>${esc(event.title)}</strong> there ${newPeople.length === 1 ? 'is a registered person who' : 'are registered people who'} may have <strong>left Deloitte</strong> (no active account). Confirmation emails and Outlook invites <strong>may not arrive</strong> for ${newPeople.length === 1 ? 'them' : 'them'}:</p>
         <ul>${list}</ul>
         <p>Please review this in the Organizer Center and deregister ${newPeople.length === 1 ? 'the person' : 'the people'} if needed.</p>
         <p><a href="${appUrl}" style="display:inline-block;background:#86bc25;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:600;">Open event in the Organizer Center</a></p>
         <p>Thanks!</p>`;
    const subject = isDe ? `Hinweis: möglicherweise inaktives Konto — ${event.title}` : `Heads-up: possibly inactive account — ${event.title}`;
    const wrapped = wrapTemplate('#86bc25', heading, sub, body);
    try {
      // Eine Mail an alle Organizer (Semikolon-getrennt, der Flow mappt To direkt).
      await eventService.queueEmail(subject, orgEmails.join(';'), orgEmails.join(';'), wrapped, 'Info', event.title, eventId);
    } catch (err) { console.warn('[DEX] notifyOrganizerOfInactive mail failed:', err); }
    return { sent: newPeople.length, skipped: people.length - newPeople.length };
  }

  /**
   * v26.40: Erkannte Ex-Deloitte-Personen automatisch abmelden (nur wenn das
   * Event `inactiveHandling === 'autoderegister'` hat). Nutzt denselben Pfad wie
   * das manuelle Abmelden im Organizer Center (Abmelde-Mail/Outlook-Ausladung/
   * Nachrücken via Flow). Gibt die tatsächlich abgemeldeten Personen zurück —
   * die Landing Page zeigt daraus einen Modal-Hinweis für den Organizer.
   */
  async function autoDeregisterInactive(
    eventId: string,
    people: Array<{ email: string; name: string }>
  ): Promise<Array<{ email: string; name: string }>> {
    const event = events.find(e => e.id === eventId);
    if (!event || !event.subsiteUrl) return [];
    const removed: Array<{ email: string; name: string }> = [];
    let regs: SPRegistration[] = [];
    try { regs = await eventService.getAllRegistrations(event.subsiteUrl); } catch { return []; }
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const actorName = (currentUserName || '').trim() || currentUserEmail;
    for (const p of people) {
      const em = (p.email || '').toLowerCase().trim();
      if (!em) continue;
      const reg = regs.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === em && ACTIVE.indexOf(r.Status) >= 0);
      if (!reg) continue;
      try {
        const ok = await eventService.cancelRegistration(event.subsiteUrl, reg.Id, actorName, currentUserEmail);
        if (ok) removed.push({ email: em, name: p.name || em });
      } catch (err) { console.warn('[DEX] autoDeregisterInactive failed for', em, err); }
    }
    if (removed.length > 0) { try { await loadEvents(); } catch { /* */ } }
    return removed;
  }

  // v26.41: Kommunikations-Log (Rundmails) eines Events lesen.
  async function getEventComms(eventId: string): Promise<EventCommRow[]> {
    try { return await eventService.getEventComms(eventId); }
    catch { return []; }
  }

  return { scanInactiveAccounts, getSentInactiveNotices, notifyOrganizerOfInactive, autoDeregisterInactive, getEventComms };
}
