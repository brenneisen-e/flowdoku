/**
 * Organizer-Antraege und Co-Organizer-Benachrichtigungen.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktionskoerper sind unveraendert - statt aus der Provider-Closure
 * beziehen sie ihre Umgebung aus dem `deps`-Objekt.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { EventService } from '../../services/EventService';
import { buildHashDeepLink } from '../../utils/deepLink';
import { wrapTemplate, coOrganizerAddedEmail, organizerOnboardingEmail } from '../../services/EmailTemplates';
import { isDeloitteInternalEmail } from '../../utils/deloitteDomain';
import { DEX_TEAM_RECIPIENTS } from '../../utils/supportContact';

export interface OrganizerRoleDeps {
  eventService: EventService;
  currentUserEmail: string;
  currentUserName: string;
  props: { context: WebPartContext };
}

export function makeOrganizerRoleActions(deps: OrganizerRoleDeps) {
  const { eventService, currentUserEmail, currentUserName, props } = deps;

  // ==================== Organizer-Antrag (v23.37) ====================

  async function requestOrganizerRole(email: string, name: string, location: string, message?: string): Promise<{ ok: boolean; reason?: string }> {
    const mail = (email || '').trim();
    if (!mail) return { ok: false, reason: 'no-email' };
    try {
      // Doppel-Antrag vermeiden: existiert schon ein offener Antrag dieser Person?
      const open = await eventService.getOrganizerRequests(true);
      if (open.some(r => (r.email || '').toLowerCase() === mail.toLowerCase())) {
        return { ok: true, reason: 'already-pending' };
      }
      const created = await eventService.createOrganizerRequest(mail, name || mail, location || '', message || '');
      if (!created.ok) return { ok: false, reason: 'create-failed' };
      // Admins per Mail informieren — mit Deep-Link zum Bestätigen (greift nur als Admin).
      try {
        {
          // v29.64: an das Funktionspostfach, nicht an die Admin-Konten (s.o.).
          const esc = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
          const appBase = `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`;
          const approveUrl = created.itemId ? buildHashDeepLink(appBase, { action: 'approveorg', request: created.itemId }) : appBase;
          const inner = `
            <p style="margin:0 0 12px;">Hallo zusammen,</p>
            <p style="margin:0 0 12px;"><strong>${esc(name || mail)}</strong> möchte <strong>Organizer</strong> werden und kann dann eigene Events anlegen und verwalten.</p>
            <table style="border-collapse:collapse;font-size:14px;margin:8px 0;">
              <tr><td style="padding:3px 16px 3px 0;color:#555;">Name:</td><td>${esc(name || '—')}</td></tr>
              <tr><td style="padding:3px 16px 3px 0;color:#555;">E-Mail:</td><td><a href="mailto:${esc(mail)}">${esc(mail)}</a></td></tr>
              <tr><td style="padding:3px 16px 3px 0;color:#555;">Standort:</td><td>${esc(location || '—')}</td></tr>
            </table>
            ${message ? `<p style="margin:8px 0 0;color:#555;">Nachricht:</p><p style="margin:4px 0 0;">${esc(message).replace(/\r?\n/g, '<br>')}</p>` : ''}
            <p style="margin:20px 0;text-align:center;"><a href="${approveUrl}" style="display:inline-block;padding:12px 26px;background:#86bc25;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Antrag in der App öffnen &amp; bestätigen</a></p>
            <p style="margin:0;color:#777;font-size:13px;">Bestätigen geht nur als Admin. Du findest offene Anträge auch direkt nach dem Öffnen der App oben als Hinweis.</p>
          `;
          const body = wrapTemplate('#86bc25', 'Neuer Organizer-Antrag', esc(name || mail), inner);
          await eventService.queueEmail(`Organizer-Antrag: ${name || mail}`, DEX_TEAM_RECIPIENTS, 'DEX-Team', body, 'OrganizerRequest', '', '0');
        }
      } catch (e) { console.warn('[DEX] requestOrganizerRole admin mail failed:', e); }
      return { ok: true, reason: created.itemId ? undefined : 'no-id' };
    } catch (err) {
      console.warn('[DEX] requestOrganizerRole error:', err);
      return { ok: false, reason: 'error' };
    }
  }

  // v26.24: Co-Organizer-Freigabe. Beim Speichern eines Events ruft der Wizard
  // dies mit der (1:1-gepaarten) Organizer-Namens-/-Mail-Liste auf. Für jede
  // benannte Person, die NOCH KEIN Organizer/Admin ist (und nicht der speichernde
  // User selbst), wird über requestOrganizerRole ein „Organizer werden"-Antrag
  // angelegt — die Admins bekommen die bestehende Antrags-Mail mit Deep-Link und
  // geben die Person frei (dann wird sie Organizer und kann das Event bearbeiten
  // und speichern). requestOrganizerRole entdoppelt offene Anträge selbst, daher
  // ist ein erneuter Save unkritisch. Best-effort — blockt den Save nie.
  async function requestCoOrganizerApprovals(orgNames: string, orgEmails: string, eventTitle: string): Promise<void> {
    try {
      const mails = (orgEmails || '').split(';').map(s => s.trim());
      const names = (orgNames || '').split(';').map(s => s.trim());
      if (mails.filter(Boolean).length === 0) return;
      // v30.67: F&A und IT-Admin haben Organizer-Rechte (utils/roleRank) — sie
      // bekamen trotzdem bei jedem Save einen "Organizer werden"-Antrag samt
      // Admin-Mail, weil hier nur zwei der vier Rollen gezaehlt wurden.
      const [orgs, admins, fa, itAdmins] = await Promise.all([
        eventService.getRoleEmails('Organizer'), eventService.getRoleEmails('Admin'),
        eventService.getRoleEmails('F&A'), eventService.getRoleEmails('IT-Admin'),
      ]);
      const elevated = new Set([...orgs, ...admins, ...fa, ...itAdmins].map(e => e.toLowerCase()));
      const me = (currentUserEmail || '').toLowerCase();
      for (let i = 0; i < mails.length; i++) {
        const mail = mails[i];
        if (!mail) continue;
        const lc = mail.toLowerCase();
        if (lc === me) continue;            // sich selbst nie anfragen
        if (elevated.has(lc)) continue;     // ist schon Organizer/Admin → kann eh speichern
        const nm = names[i] || mail;
        const requester = currentUserName || currentUserEmail || '';
        const msg = `„${requester}" hat diese Person als Co-Organizer für das Event „${eventTitle}" benannt. `
          + `Mit der Freigabe wird sie Organizer und kann das Event in DEX bearbeiten und speichern.`;
        try { await requestOrganizerRole(mail, nm, '', msg); }
        catch (e) { console.warn('[DEX] requestCoOrganizerApprovals: Einzel-Antrag fehlgeschlagen:', mail, e); }
      }
    } catch (e) {
      console.warn('[DEX] requestCoOrganizerApprovals fehlgeschlagen (best-effort):', e);
    }
  }

  /**
   * v26.34: Neu hinzugefügte Co-Organizer benachrichtigen — „Du wurdest von X zum
   * Co-Organizer gemacht und kannst nun auf die Teilnehmerliste zugreifen" — plus
   * eine Outlook-Kalendereinladung zum Event. Wird beim Event-Speichern für die
   * DIFFERENZ (neue Organizer, die vorher nicht dabei waren) aufgerufen. Best-effort.
   */
  async function notifyNewCoOrganizers(
    eventId: string,
    eventTitle: string,
    added: Array<{ name: string; email: string }>,
    isDe: boolean,
    disableOutlook?: boolean
  ): Promise<void> {
    if (!eventService || !added || added.length === 0) return;
    let appUrl = '';
    try { appUrl = `${props.context.pageContext.web.absoluteUrl}/SitePages/DEX.aspx`; } catch { appUrl = ''; }
    // Anzeigename des Anmeldenden „Vorname Nachname" (displayName ist oft „Nachname, Vorname").
    const actorDisplay = (() => {
      const p = (currentUserName || '').split(',').map(s => s.trim());
      return p.length === 2 ? `${p[1]} ${p[0]}` : (currentUserName || '');
    })();
    const seen = new Set<string>();
    for (const person of added) {
      const email = (person.email || '').trim();
      const lc = email.toLowerCase();
      if (!email || email.indexOf('@') < 0 || seen.has(lc)) continue;
      seen.add(lc);
      const name = (person.name || '').trim() || email;
      try {
        const { subject, body } = coOrganizerAddedEmail(name, eventTitle, actorDisplay, isDe, appUrl);
        await eventService.queueEmail(subject, email, name, body, 'CoOrganizerAdded', eventTitle, eventId || '0');
      } catch { /* Mail best-effort */ }
      // Outlook-Kalendereinladung — nur Deloitte-Adressen (v27.11: beliebige
      // Member Firm), nur wenn Outlook aktiv.
      if (!disableOutlook && isDeloitteInternalEmail(email)) {
        try { await eventService.queueOutlookEvent(email, eventId, eventTitle, 'Einladen'); } catch { /* */ }
      }
    }
  }

  async function getOpenOrganizerRequests(): Promise<Array<{ id: number; email: string; name: string; location: string; message: string; created: string }>> {
    try {
      const list = await eventService.getOrganizerRequests(true);
      return list.map(r => ({ id: r.id, email: r.email, name: r.name, location: r.location, message: r.message, created: r.created }));
    } catch { return []; }
  }

  async function markOrganizerRequestDecided(
    id: number, status: 'Approved' | 'Rejected', email: string, name: string,
    opts?: { suppressMail?: boolean },
  ): Promise<boolean> {
    try {
      const ok = await eventService.updateOrganizerRequestStatus(id, status, currentUserEmail || '');
      // v29.63: Hatte die Person die Rechte schon, wird der Antrag nur
      // geschlossen — ohne Mail. Eine Onboarding-Mail fuer etwas, das sich
      // nicht geaendert hat, verwirrt mehr als sie hilft.
      if (opts && opts.suppressMail) return ok;
      // Antragsteller informieren (best-effort, EventId='0').
      try {
        const esc = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        const appBase = `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`;
        if (status === 'Approved') {
          // v28.44 BUG-FIX: Hier ging bisher nur eine kurze Bestätigung raus —
          // die richtige Onboarding-Mail (Links, erstes Test-Event, Handbuch,
          // Ticketsystem, Einsatzbereich) gab es ausschliesslich beim MANUELLEN
          // Zuweisen einer Rolle. Wer über den Antragsweg Organizer wurde, hat
          // sie also nie bekommen. Jetzt bekommen beide Wege dieselbe Mail.
          const onboarding = organizerOnboardingEmail(name || email, 'Organizer');
          await eventService.queueEmail(onboarding.subject, email, name || email, onboarding.body, 'OrganizerApproved', '', '0');
          void appBase;
        } else {
          const inner = `
            <p style="margin:0 0 12px;">Hallo ${esc((name || '').split(' ')[0] || '')},</p>
            <p style="margin:0 0 12px;">vielen Dank für dein Interesse. Dein Antrag, Organizer zu werden, wurde aktuell <strong>nicht freigegeben</strong>. Bei Fragen wende dich gerne an das DEX-Team.</p>
          `;
          const body = wrapTemplate('#86bc25', 'Zu deinem Organizer-Antrag', '', inner);
          await eventService.queueEmail('Zu deinem Organizer-Antrag', email, name || email, body, 'OrganizerRejected', '', '0');
        }
      } catch (e) { console.warn('[DEX] markOrganizerRequestDecided mail failed:', e); }
      return ok;
    } catch (err) {
      console.warn('[DEX] markOrganizerRequestDecided error:', err);
      return false;
    }
  }

  return { requestOrganizerRole, requestCoOrganizerApprovals, notifyNewCoOrganizers, getOpenOrganizerRequests, markOrganizerRequestDecided };
}
