/**
 * Mails ausserhalb des Anmelde-Pfads: Erinnerung an eine unvollstaendige
 * Anmeldung, Anfrage an das DEX-Team, Hinweis auf externe Zielgruppen und
 * die Onboarding-Mail an frisch ernannte Organizer.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktionskoerper sind unveraendert - statt aus der Provider-Closure
 * beziehen sie ihre Umgebung aus dem `deps`-Objekt.
 */

import { EventService } from '../../services/EventService';
import { buildHashDeepLink } from '../../utils/deepLink';
import { wrapTemplate, organizerOnboardingEmail } from '../../services/EmailTemplates';
import { DEX_TEAM_RECIPIENTS } from '../../utils/supportContact';

export interface MailDeps {
  eventService: EventService;
}

export function makeMailActions(deps: MailDeps) {
  const { eventService } = deps;

  // v26.67 (A): Erinnerung an eine verwaiste Klammer-Anmeldung — die Person
  // hat die Anmeldung nicht abgeschlossen (kein Sub-Event gewählt). Bei
  // Fremd-Anmeldung geht die Mail an die ANMELDENDE Person (die kann es in der
  // App abschließen), die Person selbst kommt auf CC; bei Selbst-Anmeldung an
  // die Person. Deep-Link öffnet direkt die Anmeldeseite des Events.
  async function sendCompleteRegistrationReminder(args: { eventId: string; eventTitle: string; participantEmail: string; participantName: string; registeredByEmail?: string; registeredByName?: string }): Promise<boolean> {
    const esc = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const participant = (args.participantEmail || '').trim();
    if (!participant) return false;
    const actor = (args.registeredByEmail || '').trim();
    const isProxy = !!actor && actor.toLowerCase() !== participant.toLowerCase();
    const appBase = `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`;
    const link = buildHashDeepLink(appBase, { action: 'register', event: args.eventId });
    const btn = `<p style="margin:20px 0;text-align:center;"><a href="${link}" style="display:inline-block;padding:12px 26px;background:#86bc25;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Anmeldung jetzt abschließen</a></p>`;
    const firstName = (nm: string): string => (nm || '').includes(',') ? (nm.split(',')[1] || '').trim() : (nm.split(/\s+/)[0] || '');
    let to: string; let toName: string; let cc: string | undefined; let inner: string;
    if (isProxy) {
      to = actor; toName = args.registeredByName || actor;
      cc = participant;
      inner = `
        <p style="margin:0 0 12px;">Hallo ${esc(firstName(args.registeredByName || ''))},</p>
        <p style="margin:0 0 12px;">du hast <strong>${esc(args.participantName || participant)}</strong> für das Event <strong>${esc(args.eventTitle)}</strong> angemeldet — die Anmeldung ist aber <strong>noch nicht abgeschlossen</strong>: Es wurde noch kein Programmpunkt (Sub-Event) ausgewählt. Solange das offen ist, gilt die Person als <strong>nicht angemeldet</strong>.</p>
        <p style="margin:0 0 12px;">Bitte öffne die Anmeldung und wähle die gewünschten Programmpunkte aus, damit die Anmeldung gültig wird.</p>
        ${btn}
        <p style="margin:0;color:#777;font-size:13px;">Die betroffene Person ist auf Kopie (CC).</p>`;
    } else {
      to = participant; toName = args.participantName || participant;
      inner = `
        <p style="margin:0 0 12px;">Hallo ${esc(firstName(args.participantName || ''))},</p>
        <p style="margin:0 0 12px;">deine Anmeldung für das Event <strong>${esc(args.eventTitle)}</strong> ist <strong>noch nicht abgeschlossen</strong>: Es wurde noch kein Programmpunkt (Sub-Event) ausgewählt. Solange das offen ist, bist du <strong>nicht angemeldet</strong>.</p>
        <p style="margin:0 0 12px;">Bitte öffne die Anmeldung und wähle die gewünschten Programmpunkte aus, damit deine Anmeldung gültig wird.</p>
        ${btn}`;
    }
    const body = wrapTemplate('#86bc25', 'Anmeldung noch abschließen', esc(args.eventTitle), inner);
    try {
      return await eventService.queueEmail(
        `Bitte Anmeldung abschließen: ${args.eventTitle}`,
        to, toName, body, 'RegistrationReminder', args.eventTitle, args.eventId, cc || undefined, undefined, 'High'
      );
    } catch (e) { console.warn('[DEX] sendCompleteRegistrationReminder failed:', e); return false; }
  }

  async function sendAdminInquiry(
    requesterName: string,
    requesterEmail: string,
    eventName: string,
    message: string,
    requesterLocation?: string,
    requesterJobTitle?: string
  ): Promise<boolean> {
    // v29.43: Funktionspostfach statt der persönlichen Konten — eine Anfrage
    // darf nicht an einem Urlaub oder Wechsel hängen bleiben.
    const adminTo = DEX_TEAM_RECIPIENTS;
    const subject = `DEX-Anfrage: ${eventName || 'Event ohne Titel'} (von ${requesterName || 'unbekannt'})`;
    const escape = (s: string): string => s
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    const messageHtml = escape(message).replace(/\r?\n/g, '<br>');
    // v26.56: Fachliche Anforderung — die Anfrage-Mail enthält einen Deep-Link,
    // über den Admins die anfragende Person direkt als Organizer freigeben
    // können. Dazu wird (wie beim „Organizer werden"-Antrag) ein nachverfolg-
    // barer Antrag in DEX_OrganizerRequests angelegt und der bestehende
    // approveorg-Deep-Link genutzt. Nur wenn die Person nicht ohnehin schon
    // Organizer/Admin ist; ein bereits offener Antrag wird wiederverwendet
    // statt dupliziert. Fehler hier dürfen die Anfrage-Mail nie blockieren.
    let approveBlock = '';
    try {
      const mailLc = (requesterEmail || '').trim().toLowerCase();
      if (mailLc) {
        const [orgs, admins, itAdmins] = await Promise.all([
          eventService.getRoleEmails('Organizer'),
          eventService.getRoleEmails('Admin'),
          eventService.getRoleEmails('IT-Admin'),
        ]);
        const hasRole = orgs.concat(admins, itAdmins).some(e => (e || '').toLowerCase() === mailLc);
        if (!hasRole) {
          const open = await eventService.getOrganizerRequests(true);
          const existing = open.filter(r => (r.email || '').toLowerCase() === mailLc)[0];
          let requestId = existing ? existing.id : 0;
          if (!requestId) {
            const created = await eventService.createOrganizerRequest(
              requesterEmail.trim(),
              requesterName || requesterEmail,
              requesterLocation || '',
              `DEX-Anfrage zum Event „${eventName || '—'}": ${message}`.slice(0, 500)
            );
            if (created.ok && created.itemId) requestId = created.itemId;
          }
          if (requestId) {
            const appBase = `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`;
            const approveUrl = buildHashDeepLink(appBase, { action: 'approveorg', request: requestId });
            approveBlock = `
      <p style="margin:20px 0 6px;text-align:center;"><a href="${approveUrl}" style="display:inline-block;padding:12px 26px;background:#86bc25;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">${escape(requesterName || requesterEmail)} als Organizer freigeben</a></p>
      <p style="margin:0;color:#777;font-size:13px;text-align:center;">Öffnet die App und bestätigt den automatisch angelegten Organizer-Antrag — Freigeben geht nur als Admin.</p>
    `;
          }
        }
      }
    } catch (e) { console.warn('[DEX] sendAdminInquiry: Freigabe-Deep-Link übersprungen:', e); }
    const bodyInner = `
      <p>Hallo DEX-Team,</p>
      <p>es gibt eine neue Anfrage zur DEX Event Experience Platform:</p>
      <table cellpadding="6" cellspacing="0" style="border-collapse:collapse;margin:8px 0;">
        <tr><td style="color:#555;font-weight:600;vertical-align:top;">Name:</td><td>${escape(requesterName)}</td></tr>
        <tr><td style="color:#555;font-weight:600;vertical-align:top;">E-Mail:</td><td><a href="mailto:${escape(requesterEmail)}">${escape(requesterEmail)}</a></td></tr>
        ${requesterJobTitle && requesterJobTitle.trim() ? `<tr><td style="color:#555;font-weight:600;vertical-align:top;">Position:</td><td>${escape(requesterJobTitle)}</td></tr>` : ''}
        ${requesterLocation && requesterLocation.trim() ? `<tr><td style="color:#555;font-weight:600;vertical-align:top;">Standort:</td><td>${escape(requesterLocation)}</td></tr>` : ''}
        <tr><td style="color:#555;font-weight:600;vertical-align:top;">Event:</td><td>${escape(eventName)}</td></tr>
      </table>
      <p style="color:#555;font-weight:600;margin-bottom:4px;">Worum geht es:</p>
      <p>${messageHtml}</p>
      ${approveBlock}
      <p style="margin-top:24px;color:#888;font-size:0.85rem;">${escape(requesterName)} ist im Cc und kann direkt geantwortet werden.</p>
    `;
    const body = wrapTemplate('#86bc25', 'Neue DEX-Anfrage', `Event: ${eventName || '-'}`, bodyInner);
    // EventId muss '0' sein (nicht ''), damit der DEX_SEND_MAIL Flow Get_Event
    // mit "ID eq 0" als gültigem OData-Filter aufrufen kann. Bei leerem
    // EventId baut der Flow "ID eq " was kein gültiger OData-Ausdruck ist
    // und der Flow direkt in Get_Event failed (clientRequestId-Fehler).
    // Get_Event liefert dann 0 Items, Compose_Image fällt automatisch auf
    // das Default-Bild aus _Config zurück - die Mail geht trotzdem raus.
    // v18.30: Anfrage-Mail mit hoher Wichtigkeit (rotes „!" in Outlook) —
    // der DEX_SEND_MAIL-Flow liest das Importance-Feld aus der Queue aus.
    return eventService.queueEmail(
      subject, adminTo, 'DEX Admin Team', body, 'Info', eventName || 'DEX-Anfrage', '0', requesterEmail, undefined, 'High'
    );
  }

  /**
   * v26.57: Approve-Mail an die Admins, wenn Personen AUSSERHALB von
   * @deloitte.de zur Zielgruppe eines Events hinzugefügt wurden. Hintergrund:
   * Der SharePoint ist im Default nur für Deloitte DE ALL freigeschaltet —
   * internationale Member-Firm-Kolleg:innen (z. B. @deloitte.at) sehen die
   * App sonst gar nicht, selbst wenn sie in der Zielgruppe stehen. Die Mail
   * listet die Personen auf und verlinkt direkt die Site-Berechtigungsseite.
   * Fehler hier dürfen den Event-Save nie blockieren (Aufrufer feuern
   * fire-and-forget mit .catch).
   */
  async function notifyAdminsExternalAudienceAccess(eventTitle: string, persons: string[], requesterName: string): Promise<void> {
    const unique = Array.from(new Set(persons.map(p => (p || '').trim().toLowerCase()).filter(Boolean)));
    if (unique.length === 0) return;
    try {
      // v26.57: Wer schon Site-Zugriff hat (direkt oder über eine Gruppe),
      // braucht keine Freigabe mehr — aus der Mail rausfiltern. Nicht
      // prüfbare Fälle (null) bleiben drin: lieber einmal zu viel
      // benachrichtigen als eine nötige Freigabe verpassen.
      const accessChecks = await Promise.all(unique.map(async p => ({
        email: p,
        hasAccess: await eventService.userHasSiteAccess(p),
      })));
      const needsAccess = accessChecks.filter(c => c.hasAccess !== true).map(c => c.email);
      if (needsAccess.length === 0) return;
      // v29.64: Empfaenger ist das Funktionspostfach, nicht die Liste der
      // Admin-Konten. getRoleEmails('Admin') liefert die PERSOENLICHEN
      // Adressen — genau das, was seit v29.43 nicht mehr sein soll. In v29.43
      // wurden nur zwei von vier Stellen umgestellt; diese hier und die
      // Organizer-Antrags-Mail blieben stehen.
      const esc = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const permissionsUrl = `${eventService.siteUrl}/_layouts/15/user.aspx`;
      // v26.59: Der Button vergibt die Leserechte DIREKT (grantaccess-Deep-Link
      // → GrantAccessHandler in der App) statt auf die rohe SharePoint-
      // Berechtigungsseite zu verlinken. Die manuelle Seite bleibt als
      // Fallback-Link unten drin.
      const appBase = `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`;
      const grantUrl = buildHashDeepLink(appBase, { action: 'grantaccess', emails: needsAccess.join(';') });
      const alreadyOk = unique.length - needsAccess.length;
      const inner = `
        <p style="margin:0 0 12px;">Hallo zusammen,</p>
        <p style="margin:0 0 12px;">beim Event <strong>${esc(eventTitle || '—')}</strong> wurden Personen <strong>außerhalb von @deloitte.de</strong> zur Zielgruppe hinzugefügt${requesterName ? ` (durch ${esc(requesterName)})` : ''}:</p>
        <ul style="margin:0 0 12px;padding-left:20px;">
          ${needsAccess.map(p => `<li style="margin:2px 0;"><a href="mailto:${esc(p)}">${esc(p)}</a></li>`).join('')}
        </ul>
        ${alreadyOk > 0 ? `<p style="margin:0 0 12px;color:#777;font-size:13px;">${alreadyOk} weitere hinzugefügte ${alreadyOk === 1 ? 'Person hat' : 'Personen haben'} bereits Zugriff auf die Site und ${alreadyOk === 1 ? 'ist' : 'sind'} hier nicht aufgeführt.</p>` : ''}
        <p style="margin:0 0 12px;">Der SharePoint ist im Default nur für <strong>Deloitte DE ALL</strong> freigeschaltet — damit diese Personen die DEX App (und damit das Event) überhaupt öffnen können, müssen sie zusätzlich auf der Site berechtigt werden.</p>
        <p style="margin:20px 0;text-align:center;"><a href="${grantUrl}" style="display:inline-block;padding:12px 26px;background:#86bc25;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Zugriff jetzt freigeben (ein Klick)</a></p>
        <p style="margin:0;color:#777;font-size:13px;">Der Button öffnet die DEX App und vergibt automatisch Leserechte über die Besucher-Gruppe der Site — funktioniert nur als Admin. Alternativ manuell über die <a href="${permissionsUrl}" style="color:#4a7c1f;">Site-Berechtigungen</a>.</p>
      `;
      const body = wrapTemplate('#ed8b00', 'SharePoint-Zugriff benötigt', esc(eventTitle || '—'), inner);
      await eventService.queueEmail(
        `SharePoint-Zugriff benötigt: ${unique.length} ${unique.length === 1 ? 'Person' : 'Personen'} außerhalb @deloitte.de (${eventTitle || 'Event'})`,
        DEX_TEAM_RECIPIENTS, 'DEX-Team', body, 'Info', eventTitle || '', '0', undefined, undefined, 'High'
      );
    } catch (e) {
      console.warn('[DEX] notifyAdminsExternalAudienceAccess failed:', e);
    }
  }

  /**
   * Onboarding-Mail an einen neu ernannten Organizer (oder Admin) verschicken.
   * Subject + Body kommen aus EmailTemplates.organizerOnboardingEmail (Deloitte-
   * Layout inkl. Header/Footer). Die DEX-Verantwortlichen werden im Cc
   * informiert. EventId='0' damit der DEX_SEND_MAIL Flow den Get_Event-Step
   * mit gültigem OData-Filter ausführen kann (analog sendAdminInquiry).
   */
  async function sendOrganizerOnboarding(
    recipientEmail: string,
    recipientName: string,
    role: 'Organizer' | 'Admin'
  ): Promise<boolean> {
    if (!recipientEmail || !recipientName) return false;
    const cc = DEX_TEAM_RECIPIENTS;  // v29.43: siehe utils/supportContact
    const { subject, body } = organizerOnboardingEmail(recipientName, role);
    return eventService.queueEmail(
      subject, recipientEmail, recipientName, body, 'Info', 'DEX-Onboarding', '0', cc
    );
  }

  return { sendCompleteRegistrationReminder, sendAdminInquiry, notifyAdminsExternalAudienceAccess, sendOrganizerOnboarding };
}
