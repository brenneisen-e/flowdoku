/**
 * Ticket-Context (v26.0.0)
 *
 * Orchestriert das Ticketsystem: Laden der Tickets (mit Echtzeit-Push über den
 * SharePoint-Socket, Fallback-Poll), Anlegen einer Frage (grüner
 * „Hast du Fragen?"-Header-Button), Beantworten/Übernehmen/Freigeben durch
 * Power-User bzw. Organizer, und der zugehörige Mail-Versand über die
 * bestehende DEX_Emails-Queue.
 *
 * Routing (siehe TICKET_SPEC):
 *  - Frage eines Organizers/Admins  → Audience „PowerUser": Mail an alle
 *    Power-User (To), Admins in CC.
 *  - Frage eines normalen Users zu einem Event → Audience „Organizer": Mail an
 *    die Organizer DIESES Events (To), KEIN CC. Ohne Event-Kontext fällt die
 *    Frage auf die Power-User zurück.
 *
 * Jede Mail enthält einen Deep-Link, der das Ticket direkt zum Beantworten
 * öffnet, und den Hinweis, NICHT per Mail zu antworten, sondern in der App.
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { EventService } from '../services/EventService';
import { DexTicket, TicketAudience } from '../types';
import { wrapTemplate } from '../services/EmailTemplates';
import { subscribeListChanges } from '../utils/spListRealtime';
import { buildHashDeepLink, deepLinkParams } from '../utils/deepLink';
import { useRoles } from './RoleContext';
import { useCurrentUser } from './UserContext';
import { useEvents } from './EventContext';
import { useNavigation } from './NavigationContext';
import { DEX_TEAM_EMAIL } from '../utils/supportContact';

const GREEN = '#86bc25';

export interface AskInput {
  questions: string[];
  /** Optionale Screenshots des aktuellen Bildschirms (ask_-Anhänge). */
  screenshots: File[];
  /** v26.30: Explizit im Frage-Modal gewähltes Event ('' = bewusst „kein Event"
   *  → Power-User). undefined = kein Signal → Seiten-Kontext (selectedEventId). */
  eventId?: string;
  /** v26.30: 1-basierter Wizard-Schritt, in dem die Frage gestellt wurde
   *  (Organizer im Event-Wizard) — null/undefined = nicht im Wizard. */
  askWizardStep?: number | null;
  /** v26.60: 'bug' = Bug-Report — landet im Ticketsystem, die Benachrichtigung
   *  geht aber NUR an die DEX-Maintainer (statt an alle Power-User).
   *  undefined/'question' = inhaltliche Frage (bisheriges Verhalten). */
  category?: 'question' | 'bug';
}

export interface AnswerInput {
  ticket: DexTicket;
  answerText: string;
  /** In der Antwort verlinkte Handbuch-Artikel (id + aufgelöster Titel). */
  articles: Array<{ id: string; title: string }>;
  /** 1-basierter Event-Wizard-Schritt (null = keiner). */
  wizardStep: number | null;
  wizardStepLabel?: string;
  /** v26.52: Markierungsbox auf der Live-Wizard-Vorschau ({x,y,w,h} in %). */
  wizardMarker?: { x: number; y: number; w: number; h: number } | null;
  /** Optionale Screenshots (ans_-Anhänge), z.B. der Wizard-Schritt. */
  screenshots: File[];
}

/** v26.32: Ergebnis einer Ticket-Übernahme. `conflict` = ein anderer Power-User
 *  war schneller (dann steht in `claimedByName`, wer es hat). */
export interface ClaimResult { ok: boolean; conflict?: boolean; claimedByName?: string }

/** v26.33: Anzahl VOLLER Werktage (Mo–Fr, ohne Feiertage) zwischen zwei Daten.
 *  Grundlage für die „≥ 2 Werktage unbeantwortet"-Ticket-Erinnerung. */
function businessDaysBetween(from: Date, to: Date): number {
  if (!(to > from)) return 0;
  const cur = new Date(from.getTime()); cur.setHours(0, 0, 0, 0);
  const end = new Date(to.getTime()); end.setHours(0, 0, 0, 0);
  let count = 0;
  let guard = 0;
  while (cur < end && guard < 3660) {
    cur.setDate(cur.getDate() + 1);
    const dow = cur.getDay(); // 0 = So, 6 = Sa
    if (dow !== 0 && dow !== 6) count++;
    guard++;
  }
  return count;
}

interface TicketContextType {
  tickets: DexTicket[];
  myTickets: DexTicket[];
  isTicketsLoading: boolean;
  /** Aktueller User darf Tickets beantworten (Power-User oder Admin). */
  canAnswerTickets: boolean;
  reloadTickets: () => Promise<void>;
  reloadMyTickets: () => Promise<void>;
  createTicket: (input: AskInput) => Promise<boolean>;
  /** v26.32: übernimmt ein Ticket; `onlyIfOpen` (Klick auf ein offenes Ticket)
   *  verhindert das stille Überschreiben, wenn jemand schneller war. */
  claimTicket: (ticketId: number, opts?: { onlyIfOpen?: boolean }) => Promise<ClaimResult>;
  releaseTicket: (ticketId: number) => Promise<void>;
  answerTicket: (input: AnswerInput) => Promise<boolean>;
  /** v26.8: Auf eine beantwortete Frage erneut antworten (Fragesteller → nur an
   *  die/den Beantwortenden; Beantwortende → an den Fragesteller). */
  replyToTicket: (ticket: DexTicket, text: string) => Promise<boolean>;
  /** v26.8: Ticket schließen, ohne Antwort-Mail (z.B. Rückfrage = nur Danke). */
  closeTicketNoAnswer: (ticketId: number) => Promise<void>;
  /** Tickets eines normalen Users zu einem konkreten Event (Audience=Organizer). */
  ticketsForEvent: (eventId: string) => DexTicket[];
  /** Anzahl noch offener (nicht geschlossener) User-Fragen zu einem Event. */
  openCountForEvent: (eventId: string) => number;
  /** Power-User-Warteschlange (Fragen von Organizern/Admins). */
  powerUserQueue: DexTicket[];
  /** v29.14: Erinnerung an die Power-User schicken (alle unbeantworteten
   *  Tickets in EINER Mail). Ohne Tages-Sperre — hier entscheidet ein Mensch. */
  remindPowerUsers: () => Promise<{ ok: boolean; count: number; recipients: string[]; reason?: 'no-tickets' | 'no-recipients' | 'error' }>;
  /** v29.14: Namen der Personen, die die Erinnerung erreichen würde (ohne
   *  den aktuellen User) — für die Rückfrage vor dem Senden. */
  reminderTargets: () => string[];
}

export const TicketContext = React.createContext<TicketContextType | undefined>(undefined);

function esc(s: string): string {
  return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function firstNameOf(name: string): string {
  const n = (name || '').trim();
  if (!n) return '';
  const c = n.indexOf(',');
  if (c >= 0) return n.substring(c + 1).trim().split(/\s+/)[0];
  return n.split(/\s+/)[0];
}

export function TicketProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const eventService = React.useMemo(() => new EventService(props.context), []);
  const { isAdmin, isPowerUser, roles } = useRoles();
  const { currentUser } = useCurrentUser();
  const { events } = useEvents();
  const { selectedEventId } = useNavigation();

  const [tickets, setTickets] = React.useState<DexTicket[]>([]);
  const [myTickets, setMyTickets] = React.useState<DexTicket[]>([]);
  const [isTicketsLoading, setIsTicketsLoading] = React.useState<boolean>(false);
  // v26.33: true, sobald die Ticket-Liste MINDESTENS EINMAL geladen wurde —
  // Trigger für die tägliche Ticket-Erinnerung (erst nach echtem Load prüfen).
  const [ticketsLoadedOnce, setTicketsLoadedOnce] = React.useState<boolean>(false);

  const myEmailLc = (currentUser.email || '').toLowerCase();

  // Ist der User Organizer/Co-Organizer mindestens eines Events? (sieht dann die
  // an seine Events gerichteten User-Fragen in der Event-Übersicht).
  const isOrganizerOfAnyEvent = React.useMemo(() => {
    if (!myEmailLc) return false;
    return (events || []).some(e =>
      (e.organizerEmails || []).some(x => (x || '').toLowerCase() === myEmailLc) ||
      (e.coOrganizerEmails || []).some(x => (x || '').toLowerCase() === myEmailLc)
    );
  }, [events, myEmailLc]);

  const canAnswerTickets = isAdmin || isPowerUser;
  const shouldLoadAll = canAnswerTickets || isOrganizerOfAnyEvent;

  const reloadTickets = React.useCallback(async (): Promise<void> => {
    setIsTicketsLoading(true);
    try {
      const list = await eventService.getTickets();
      setTickets(list);
      setTicketsLoadedOnce(true);
    } catch { /* best-effort */ }
    setIsTicketsLoading(false);
  }, [eventService]);

  const reloadMyTickets = React.useCallback(async (): Promise<void> => {
    if (!currentUser.email) return;
    try {
      const list = await eventService.getMyTickets(currentUser.email);
      setMyTickets(list);
    } catch { /* best-effort */ }
  }, [eventService, currentUser.email]);

  // Initial-Load + Echtzeit-Push für die Beantwortenden. Pure Fragesteller
  // laden nur ihre eigenen Tickets on-demand (Datenschutz + Performance).
  React.useEffect(() => {
    if (!shouldLoadAll) return;
    let cleanup: (() => void) | null = null;
    let cancelled = false;
    reloadTickets().catch(() => { /* */ });
    subscribeListChanges(props.context.spHttpClient, eventService.siteUrl, 'DEX_Tickets', () => {
      reloadTickets().catch(() => { /* */ });
    }).then(fn => { if (cancelled) { try { fn(); } catch { /* */ } } else { cleanup = fn; } })
      .catch(() => { /* Echtzeit best-effort */ });
    // Fallback-Poll alle 30 s, falls der Socket-Push nicht greift (inoffizieller
    // Endpunkt) — hält „in Bearbeitung" trotzdem nahezu aktuell.
    const poll = window.setInterval(() => { reloadTickets().catch(() => { /* */ }); }, 30000);
    return () => {
      cancelled = true;
      if (cleanup) { try { cleanup(); } catch { /* */ } }
      window.clearInterval(poll);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shouldLoadAll, eventService]);

  // v26.7: Deep-Link-Basis = die SEITE, auf der die App tatsächlich läuft
  // (window.location), statt eines fest verdrahteten „/SitePages/DEX.aspx".
  // Sonst landete der Link auf der Team-Site-Startseite statt in der App.
  const appBase = (() => {
    try {
      if (typeof window !== 'undefined' && window.location && window.location.pathname) {
        return `${window.location.origin}${window.location.pathname}?env=WebView`;
      }
    } catch { /* */ }
    return `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`;
  })();

  // ---- Mail-Bausteine -----------------------------------------------------
  const noReplyHintHtml = `<div style="margin-top:18px;padding:12px 14px;background:#fff3e0;border:1px solid #ed8b00;border-radius:6px;color:#8a4b00;font-size:13px;">
    <strong>Bitte nicht auf diese E-Mail antworten.</strong> Diese Adresse wird nicht gelesen — bearbeite das Ticket direkt in der App über den Button oben.</div>`;

  function ctaButton(href: string, label: string): string {
    return `<p style="margin:22px 0 0;"><a href="${href}" style="display:inline-block;background:${GREEN};color:#ffffff;text-decoration:none;padding:11px 24px;border-radius:6px;font-weight:600;font-size:14px;">${esc(label)}</a></p>`;
  }

  // ---- Ticket-Erinnerung (Automatik v26.33 + Knopf v29.14) ----------------
  /**
   * v29.14: Empfänger einer Ticket-Erinnerung — die Power-User, ersatzweise die
   * Admins. IT-Admins sind bewusst NICHT dabei. `exclude` nimmt eine Adresse
   * heraus (beim Knopf die Person, die ihn drückt: Wer selbst erinnert, braucht
   * die Erinnerung nicht).
   */
  const reminderRecipients = React.useCallback((exclude?: string): { emails: string[]; names: string[] } => {
    const ex = (exclude || '').toLowerCase().trim();
    const pu = roles.filter(r => r.isPowerUser && r.userEmail);
    let list = pu.length > 0 ? pu : roles.filter(r => r.role === 'Admin' && r.userEmail);
    if (ex) list = list.filter(r => (r.userEmail || '').toLowerCase() !== ex);
    return { emails: list.map(r => r.userEmail), names: list.map(r => r.userName || r.userEmail) };
  }, [roles]);

  /**
   * v29.14: Baut und stellt die Sammel-Erinnerung ein. Automatik und Knopf
   * teilen sich diese eine Stelle — der Text der Erinnerung stand sonst
   * zweimal im Code und wäre beim nächsten Anfassen auseinandergelaufen.
   *
   * v29.15: OHNE Absender-Namen. Ich hatte den Namen dessen, der den Knopf
   * drückt, in Betreff und Anrede gesetzt — das macht aus einem sachlichen
   * Systemhinweis eine persönliche Mahnung an Kolleginnen und Kollegen und
   * war so nicht gewünscht. Die Erinnerung sagt jetzt nur, dass ein Ticket
   * offen ist. `overdueOnly` steuert allein den Zusatz „seit mindestens zwei
   * Werktagen", der nur für die Automatik zutrifft.
   */
  const queueTicketReminder = React.useCallback(async (
    list: DexTicket[],
    to: { emails: string[]; names: string[] },
    overdueOnly: boolean,
  ): Promise<boolean> => {
    if (list.length === 0 || to.emails.length === 0) return false;
    const now = new Date();
    const link = buildHashDeepLink(appBase, { action: 'tickets' });
    const rows = list.slice(0, 30).map(t => {
      const q = (t.questions && t.questions[0]) ? t.questions[0] : '(ohne Text)';
      const created = t.created ? new Date(t.created) : now;
      const days = businessDaysBetween(created, now);
      const age = days > 0
        ? `seit ${days} Werktag${days > 1 ? 'en' : ''} offen`
        : 'heute eingegangen';
      return `<li style="margin-bottom:8px;"><strong>${esc(q.slice(0, 140))}</strong><br><span style="color:#777;font-size:12px;">von ${esc(t.askerName || t.askerEmail || 'unbekannt')} · ${age}${t.eventTitle ? ` · Event: ${esc(t.eventTitle)}` : ''}</span></li>`;
    }).join('');
    const n = list.length;
    // v29.15: Ein/Mehrzahl sauber — vorher stand da „1 offene Ticket" im
    // Betreff und „warten 1 Frage" im Text.
    const one = n === 1;
    const intro = one
      ? `<p style="margin:0 0 14px;">im DEX-Ticketsystem wartet <strong>eine Frage</strong>${overdueOnly ? ' seit mindestens zwei Werktagen' : ''} auf eine Antwort:</p>`
      : `<p style="margin:0 0 14px;">im DEX-Ticketsystem warten <strong>${n} Fragen</strong>${overdueOnly ? ' seit mindestens zwei Werktagen' : ''} auf eine Antwort:</p>`;
    const inner = `
      <p style="margin:0 0 6px;">Hallo,</p>
      ${intro}
      <ul style="margin:0 0 4px 18px;padding:0;">${rows}</ul>
      ${n > 30 ? `<p style="margin:6px 0 0;color:#777;font-size:12px;">… und weitere.</p>` : ''}
      ${ctaButton(link, 'Offene Tickets ansehen')}
      ${noReplyHintHtml}
    `;
    const subject = one
      ? 'Erinnerung: ein offenes Ticket im DEX-Ticketsystem'
      : `Erinnerung: ${n} offene Tickets im DEX-Ticketsystem`;
    const body = wrapTemplate(GREEN, 'Offene Tickets warten auf Antwort', 'DEX-Support', inner);
    await eventService.queueEmail(subject, to.emails.join('; '), to.names.join('; '), body, 'TicketReminder', 'DEX-Ticket', '0');
    return true;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eventService, appBase]);

  // ---- Aktionen -----------------------------------------------------------
  const createTicket = React.useCallback(async (input: AskInput): Promise<boolean> => {
    const questions = (input.questions || []).map(q => (q || '').trim()).filter(Boolean);
    if (questions.length === 0) return false;
    const askerRole = isAdmin ? 'Admin' : (roles.find(r => (r.userEmail || '').toLowerCase() === myEmailLc)?.role || 'User');
    const askerIsOrganizerLike = askerRole === 'Organizer' || askerRole === 'Admin';

    // v26.30: Explizit im Frage-Modal gewähltes Event hat Vorrang. '' = bewusst
    // „kein Event" → Power-User. undefined (kein Signal) → Seiten-Kontext
    // (selectedEventId), wie bisher.
    const resolvedEventId = input.eventId !== undefined ? (input.eventId || '') : (selectedEventId || '');
    const ctxEvent = resolvedEventId ? (events || []).find(e => e.id === resolvedEventId) : undefined;

    let audience: TicketAudience = 'PowerUser';
    let eventId = '';
    let eventTitle = '';
    let assignedOrganizers: string[] = [];
    if (askerIsOrganizerLike) {
      audience = 'PowerUser';
      if (ctxEvent) { eventId = ctxEvent.id; eventTitle = ctxEvent.title; }
    } else {
      const orgs = ctxEvent
        ? Array.from(new Set([...(ctxEvent.organizerEmails || []), ...(ctxEvent.coOrganizerEmails || [])]
            .map(x => (x || '').trim()).filter(Boolean)))
        : [];
      if (ctxEvent && orgs.length > 0) {
        audience = 'Organizer';
        eventId = ctxEvent.id; eventTitle = ctxEvent.title; assignedOrganizers = orgs;
      } else {
        audience = 'PowerUser'; // kein Event-Kontext → an die Power-User
      }
    }

    // v26.60: Bug-Reports landen IMMER in der Power-User-Queue (Ticketsystem),
    // die Benachrichtigung geht aber unten nur an die DEX-Maintainer. Ein
    // Event-Bezug bleibt als Kontext erhalten, Organizer-Routing entfällt.
    const category: 'question' | 'bug' = input.category === 'bug' ? 'bug' : 'question';
    if (category === 'bug') {
      audience = 'PowerUser';
      assignedOrganizers = [];
      if (ctxEvent) { eventId = ctxEvent.id; eventTitle = ctxEvent.title; }
    }

    const pageContext = (() => { try { return deepLinkParams().get('action') || window.location.pathname; } catch { return ''; } })();

    const id = await eventService.createTicket({
      questions,
      askerEmail: currentUser.email || '',
      askerName: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || (currentUser.email || ''),
      askerRole,
      askerLocation: currentUser.location || '',
      askerJobTitle: currentUser.jobTitle || '',
      audience, eventId, eventTitle, assignedOrganizers, pageContext,
      askWizardStep: input.askWizardStep ?? null,
      category,
    });
    if (id == null) return false;

    // Screenshots anhängen (best-effort).
    for (const f of (input.screenshots || [])) {
      try { await eventService.addTicketAttachment(id, f, 'ask'); } catch { /* */ }
    }

    // Empfänger + CC bestimmen.
    let toEmails: string[] = [];
    let toNames: string[] = [];
    let cc = '';
    if (category === 'bug') {
      // v26.60: Bug-Reports gehen NUR an das DEX-Team (Wunsch: „Bug ist für
      // das Team, inhaltliche Fragen für die Power-User") — dieselbe Adresse
      // wie die DEX-Anfrage-Mail (sendAdminInquiry).
      // v29.43: Funktionspostfach statt der beiden persönlichen Konten.
      toEmails = [DEX_TEAM_EMAIL];
      toNames = ['DEX-Team'];
    } else if (audience === 'PowerUser') {
      const pu = roles.filter(r => r.isPowerUser && r.userEmail);
      if (pu.length > 0) { toEmails = pu.map(r => r.userEmail); toNames = pu.map(r => r.userName || r.userEmail); }
      const admins = roles.filter(r => r.role === 'Admin' && r.userEmail).map(r => r.userEmail);
      if (toEmails.length === 0) {
        // Keine Power-User hinterlegt → Fallback an die Admins (dann ohne CC-Dopplung).
        toEmails = admins; toNames = admins;
      } else {
        cc = admins.filter(a => toEmails.indexOf(a) < 0).join('; ');
      }
    } else {
      toEmails = assignedOrganizers;
      toNames = assignedOrganizers.map(a => {
        const r = roles.find(x => (x.userEmail || '').toLowerCase() === a.toLowerCase());
        return r?.userName || a;
      });
    }

    if (toEmails.length > 0) {
      const link = audience === 'PowerUser'
        ? buildHashDeepLink(appBase, { action: 'tickets', id })
        : buildHashDeepLink(appBase, { action: 'admin', event: eventId, ticket: id });
      const qHtml = questions.length === 1
        ? `<p style="margin:6px 0 0;">${esc(questions[0])}</p>`
        : `<ul style="margin:8px 0 0 18px;padding:0;">${questions.map(q => `<li style="margin-bottom:6px;">${esc(q)}</li>`).join('')}</ul>`;
      const isBug = category === 'bug';
      const inner = `
        <p style="margin:0 0 6px;">Hallo,</p>
        <p style="margin:0 0 16px;">im DEX-Ticketsystem ist ${isBug ? 'ein neuer <strong>Bug-Report</strong>' : 'eine neue Frage'} eingegangen${eventTitle ? ` (Event: <strong>${esc(eventTitle)}</strong>)` : ''}.</p>
        <p style="margin:0;"><strong>Von:</strong> ${esc(toNamesAskerLabel(currentUser))}</p>
        <p style="margin:14px 0 0;"><strong>${isBug ? 'Beschreibung' : `Frage${questions.length > 1 ? 'n' : ''}`}:</strong></p>
        ${qHtml}
        ${ctaButton(link, 'Ticket öffnen & beantworten')}
        ${noReplyHintHtml}
      `;
      const subject = isBug
        ? `Neuer Bug-Report im DEX-Ticketsystem`
        : questions.length === 1
          ? `Neue Frage im DEX-Ticketsystem`
          : `Neue Fragen (${questions.length}) im DEX-Ticketsystem`;
      const body = wrapTemplate(isBug ? '#ed8b00' : GREEN, isBug ? 'Neuer Bug-Report' : 'Neue Frage im Ticketsystem', eventTitle || 'DEX-Support', inner);
      try {
        await eventService.queueEmail(
          subject, toEmails.join('; '), toNames.join('; '), body,
          'TicketNew', eventTitle || 'DEX-Ticket', eventId || '0', cc || undefined
        );
      } catch { /* Mail best-effort */ }
    }

    // v26.25: Eingangs-Bestätigung an den Fragesteller, wenn die Frage zu einem
    // Event an die Organizer geroutet wurde — mit namentlicher Nennung der
    // Organizer + Hinweis, dass sie über die App antworten und es je nach
    // Verfügbarkeit etwas dauern kann (freundlich formuliert).
    if (audience === 'Organizer' && currentUser.email && toNames.length > 0) {
      const many = toNames.length > 1;
      const orgList = many
        ? `${toNames.slice(0, -1).map(esc).join(', ')} und ${esc(toNames[toNames.length - 1])}`
        : esc(toNames[0]);
      const kuemmern = many
        ? 'kümmern sich darum und antworten dir direkt über die DEX-App'
        : 'kümmert sich darum und antwortet dir direkt über die DEX-App';
      const zeit = many
        ? 'wie schnell die Organizer gerade Zeit finden'
        : 'wie schnell der/die Organizer gerade Zeit findet';
      const qListAsk = questions.length === 1
        ? `<p style="margin:6px 0 0;color:#555;">${esc(questions[0])}</p>`
        : `<ul style="margin:8px 0 0 18px;padding:0;color:#555;">${questions.map(q => `<li style="margin-bottom:4px;">${esc(q)}</li>`).join('')}</ul>`;
      const innerAsk = `
        <p style="margin:0 0 6px;">Hallo ${esc(currentUser.firstName || '')},</p>
        <p style="margin:0 0 16px;">vielen Dank für deine Frage zum Event <strong>${esc(eventTitle)}</strong> — sie ist erfolgreich eingegangen.</p>
        <p style="margin:0;"><strong>Deine Frage${questions.length > 1 ? 'n' : ''}:</strong></p>
        ${qListAsk}
        <p style="margin:16px 0 0;padding:10px 12px;background:#f1f7e8;border:1px solid ${GREEN};border-radius:6px;">Deine Frage liegt jetzt bei <strong>${orgList}</strong> (Organisation dieses Events). ${many ? 'Sie' : 'Die Person'} ${kuemmern}.</p>
        <p style="margin:14px 0 0;color:#555;">Je nachdem, ${zeit}, kann die Antwort etwas dauern — bitte hab ein wenig Geduld. Du bekommst die Antwort automatisch per Mail und findest sie jederzeit in der App über den grünen Button <strong>&bdquo;Hast du Fragen?&ldquo;</strong> (oben rechts) unter <strong>&bdquo;Deine Fragen&ldquo;</strong>.</p>
        ${noReplyHintHtml}
      `;
      const bodyAsk = wrapTemplate(GREEN, 'Deine Frage ist eingegangen', eventTitle || 'DEX-Support', innerAsk);
      try {
        await eventService.queueEmail(
          `Deine Frage zum Event „${eventTitle}" ist eingegangen`,
          currentUser.email,
          `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email,
          bodyAsk,
          'TicketReceived', eventTitle || 'DEX-Ticket', eventId || '0'
        );
      } catch { /* Mail best-effort */ }
    }

    // States aktualisieren.
    reloadMyTickets().catch(() => { /* */ });
    if (shouldLoadAll) reloadTickets().catch(() => { /* */ });
    return true;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eventService, roles, isAdmin, myEmailLc, currentUser, events, selectedEventId, shouldLoadAll]);

  const claimTicket = React.useCallback(async (ticketId: number, opts?: { onlyIfOpen?: boolean }): Promise<ClaimResult> => {
    const name = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || (currentUser.email || '');
    const res = await eventService.claimTicket(ticketId, currentUser.email || '', name, opts);
    // Immer neu laden: bei Erfolg zeigt die Liste „In Bearbeitung" mit meinem
    // Namen, bei Konflikt den tatsächlichen (fremden) Übernehmer.
    await reloadTickets();
    return { ok: res.ok, conflict: res.conflict, claimedByName: res.claimedByName };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eventService, currentUser]);

  const releaseTicket = React.useCallback(async (ticketId: number): Promise<void> => {
    await eventService.releaseTicket(ticketId);
    await reloadTickets();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eventService]);

  const answerTicket = React.useCallback(async (input: AnswerInput): Promise<boolean> => {
    const t = input.ticket;
    const answeredByName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || (currentUser.email || '');
    const ok = await eventService.answerTicket(t.id, {
      answerText: input.answerText || '',
      articleIds: (input.articles || []).map(a => a.id),
      wizardStep: input.wizardStep,
      wizardMarker: input.wizardMarker || null,
      answeredByEmail: currentUser.email || '',
      answeredByName,
      answeredByLocation: currentUser.location || '',
      answeredByJobTitle: currentUser.jobTitle || '',
    });
    if (!ok) return false;
    for (const f of (input.screenshots || [])) {
      try { await eventService.addTicketAttachment(t.id, f, 'ans'); } catch { /* */ }
    }

    // Antwort-Mail an den Fragesteller (mit dem vollen Antworttext, Artikel-
    // Links und Wizard-Hinweis — er muss NICHT in die App, hat die Antwort
    // aber auch dort unter „Hast du Fragen?" → „Deine Fragen").
    if (t.askerEmail) {
      const articleHtml = (input.articles || []).length > 0
        ? `<p style="margin:16px 0 0;"><strong>Passende Handbuch-Artikel:</strong></p>
           <ul style="margin:8px 0 0 18px;padding:0;">${(input.articles || []).map(a => `<li style="margin-bottom:4px;"><a href="${buildHashDeepLink(appBase, { action: 'manual', section: a.id })}" style="color:${GREEN};font-weight:600;">${esc(a.title)}</a></li>`).join('')}</ul>`
        : '';
      const wizardHtml = (input.wizardStep != null)
        ? `<p style="margin:16px 0 0;padding:10px 12px;background:#f1f7e8;border:1px solid ${GREEN};border-radius:6px;"><strong>Im Event-Wizard:</strong> Schritt ${input.wizardStep}${input.wizardStepLabel ? ` — ${esc(input.wizardStepLabel)}` : ''}.</p>`
        : '';
      const qList = (t.questions || []).length === 1
        ? `<p style="margin:6px 0 0;color:#555;">${esc(t.questions[0])}</p>`
        : `<ul style="margin:8px 0 0 18px;padding:0;color:#555;">${(t.questions || []).map(q => `<li>${esc(q)}</li>`).join('')}</ul>`;
      const inner = `
        <p style="margin:0 0 6px;">Hallo ${esc(firstNameOf(t.askerName))},</p>
        <p style="margin:0 0 16px;">deine Frage im DEX-Ticketsystem wurde von <strong>${esc(answeredByName)}</strong> beantwortet.</p>
        <p style="margin:0;"><strong>Deine Frage${(t.questions || []).length > 1 ? 'n' : ''}:</strong></p>
        ${qList}
        <p style="margin:16px 0 0;"><strong>Antwort:</strong></p>
        <div style="margin:6px 0 0;white-space:pre-wrap;">${esc(input.answerText || '').replace(/\n/g, '<br>')}</div>
        ${articleHtml}
        ${wizardHtml}
        ${ctaButton(buildHashDeepLink(appBase, { action: 'ask' }), 'Antwort in der App ansehen')}
        <p style="margin:18px 0 0;color:#666;font-size:13px;">Du findest diese Antwort jederzeit in der App über den grünen Button <strong>&bdquo;Hast du Fragen?&ldquo;</strong> (oben rechts) unter <strong>&bdquo;Deine Fragen&ldquo;</strong>. Falls noch etwas offen ist, stelle dort einfach eine neue Frage.</p>
        ${noReplyHintHtml}
      `;
      const body = wrapTemplate(GREEN, 'Deine Frage wurde beantwortet', t.eventTitle || 'DEX-Support', inner);
      try {
        await eventService.queueEmail(
          'Deine Frage im DEX-Ticketsystem wurde beantwortet',
          t.askerEmail, t.askerName || t.askerEmail, body,
          'TicketAnswered', t.eventTitle || 'DEX-Ticket', t.eventId || '0'
        );
      } catch { /* Mail best-effort */ }
    }

    await reloadTickets();
    return true;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eventService, currentUser, appBase]);

  // v26.8: Auf eine bereits beantwortete Frage erneut antworten (Rückfrage).
  //  - Fragesteller → die Frage wird wieder geöffnet UND geht NUR an die Person,
  //    die geantwortet hat (nicht an alle Power-User/Organizer).
  //  - Beantwortende → Folge-Antwort, Ticket wird wieder geschlossen, Mail an
  //    den Fragesteller.
  const replyToTicket = React.useCallback(async (ticket: DexTicket, text: string): Promise<boolean> => {
    const msg = (text || '').trim();
    if (!msg) return false;
    const myName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || (currentUser.email || '');
    const iAmAsker = (currentUser.email || '').toLowerCase() === (ticket.askerEmail || '').toLowerCase();
    const role: 'asker' | 'answerer' = iAmAsker ? 'asker' : 'answerer';
    const now = new Date().toISOString();
    const newFollowUps = [...(ticket.followUps || []), { byEmail: currentUser.email || '', byName: myName, byRole: role, text: msg, at: now }];

    let ok = false;
    if (iAmAsker) {
      // Wieder öffnen und der/dem ursprünglich Beantwortenden zuweisen.
      ok = await eventService.setTicketFollowUps(ticket.id, newFollowUps, {
        status: 'InProgress',
        claimedByEmail: ticket.answeredByEmail || '',
        claimedByName: ticket.answeredByName || '',
        claimedAt: now,
      });
    } else {
      ok = await eventService.setTicketFollowUps(ticket.id, newFollowUps, { status: 'Closed' });
    }
    if (!ok) return false;

    // Ziel-Empfänger: Rückfrage des Fragestellers → NUR die/der Beantwortende;
    // Folge-Antwort der/des Beantwortenden → der Fragesteller.
    const toEmail = iAmAsker ? (ticket.answeredByEmail || '') : (ticket.askerEmail || '');
    const toName = iAmAsker ? (ticket.answeredByName || ticket.answeredByEmail || '') : (ticket.askerName || ticket.askerEmail || '');
    if (toEmail) {
      const link = iAmAsker
        ? (ticket.audience === 'PowerUser'
            ? buildHashDeepLink(appBase, { action: 'tickets', id: ticket.id })
            : buildHashDeepLink(appBase, { action: 'admin', event: ticket.eventId, ticket: ticket.id }))
        : buildHashDeepLink(appBase, { action: 'ask' });
      const heading = iAmAsker ? 'Rückfrage zu einer beantworteten Frage' : 'Neue Antwort auf deine Frage';
      const introLine = iAmAsker
        ? `<strong>${esc(myName)}</strong> hat auf deine Antwort im DEX-Ticketsystem reagiert${ticket.eventTitle ? ` (Event: <strong>${esc(ticket.eventTitle)}</strong>)` : ''}.`
        : `Deine Frage im DEX-Ticketsystem wurde von <strong>${esc(myName)}</strong> ergänzt/beantwortet.`;
      const ctaLabel = iAmAsker ? 'Rückfrage öffnen & beantworten' : 'Antwort in der App ansehen';
      const inner = `
        <p style="margin:0 0 6px;">Hallo ${esc(firstNameOf(toName))},</p>
        <p style="margin:0 0 16px;">${introLine}</p>
        <p style="margin:0;"><strong>Nachricht:</strong></p>
        <div style="margin:6px 0 0;white-space:pre-wrap;">${esc(msg).replace(/\n/g, '<br>')}</div>
        ${ctaButton(link, ctaLabel)}
        ${noReplyHintHtml}
      `;
      const body = wrapTemplate(GREEN, heading, ticket.eventTitle || 'DEX-Support', inner);
      try {
        await eventService.queueEmail(
          iAmAsker ? 'Rückfrage im DEX-Ticketsystem' : 'Neue Antwort im DEX-Ticketsystem',
          toEmail, toName, body,
          iAmAsker ? 'TicketFollowUp' : 'TicketAnswered',
          ticket.eventTitle || 'DEX-Ticket', ticket.eventId || '0'
        );
      } catch { /* Mail best-effort */ }
    }

    reloadMyTickets().catch(() => { /* */ });
    if (shouldLoadAll) reloadTickets().catch(() => { /* */ });
    return true;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eventService, currentUser, appBase, shouldLoadAll]);

  // v26.8: „Keine Antwort nötig" — Ticket schließen, ohne eine Antwort-Mail zu
  // senden (z.B. wenn die Rückfrage nur ein Dankeschön war).
  const closeTicketNoAnswer = React.useCallback(async (ticketId: number): Promise<void> => {
    await eventService.closeTicketNoAnswer(ticketId);
    await reloadTickets();
    reloadMyTickets().catch(() => { /* */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eventService]);

  const ticketsForEvent = React.useCallback((eventId: string): DexTicket[] => {
    if (!eventId) return [];
    return tickets.filter(t => t.eventId === eventId && t.audience === 'Organizer');
  }, [tickets]);

  const openCountForEvent = React.useCallback((eventId: string): number => {
    return ticketsForEvent(eventId).filter(t => t.status !== 'Closed').length;
  }, [ticketsForEvent]);

  const powerUserQueue = React.useMemo(() => tickets.filter(t => t.audience === 'PowerUser'), [tickets]);

  // v26.33: Ticket-Erinnerung. Öffnet ein Beantworter (Power-User/Admin) die App
  // und liegen Power-User-Fragen ≥ 2 Werktage unbeantwortet, geht EINMAL PRO TAG
  // (queue- + browser-entdoppelt) eine Sammel-Erinnerung an die Power-User.
  const ticketReminderRan = React.useRef(false);
  React.useEffect(() => {
    if (ticketReminderRan.current) return;
    if (!canAnswerTickets || !ticketsLoadedOnce) return;
    ticketReminderRan.current = true;
    void (async (): Promise<void> => {
      try {
        const now = new Date();
        const overdue = powerUserQueue.filter(t => {
          if (t.status === 'Closed') return false;
          if (t.answeredAt) return false;
          const created = t.created ? new Date(t.created) : null;
          if (!created || isNaN(created.getTime())) return false;
          return businessDaysBetween(created, now) >= 2;
        });
        if (overdue.length === 0) return;
        // Browser-Tages-Dedup.
        const dayKey = `dex_ticketreminder_${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()}`;
        try { if (window.localStorage.getItem(dayKey)) return; } catch { /* */ }
        // Queue-Tages-Dedup (falls mehrere Power-User heute schon geöffnet haben).
        const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0).toISOString();
        try { if (await eventService.hasQueuedEmailSince('TicketReminder', startOfToday)) { try { window.localStorage.setItem(dayKey, '1'); } catch { /* */ } return; } } catch { /* */ }
        // v29.14: Empfänger und Mailtext kommen aus derselben Stelle wie beim
        // Erinnerungs-Knopf auf der Tickets-Seite. Die Automatik erinnert das
        // ganze Team, also OHNE Ausschluss.
        const to = reminderRecipients();
        if (to.emails.length === 0) return;
        await queueTicketReminder(overdue, to, true);
        try { window.localStorage.setItem(dayKey, '1'); } catch { /* */ }
      } catch (e) { console.warn('[DEX] ticket reminder failed:', e); }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [canAnswerTickets, ticketsLoadedOnce]);

  /**
   * v29.14: Erinnerung auf Knopfdruck. Die Automatik greift erst nach zwei
   * Werktagen und nur, wenn ein Beantworter die App öffnet — wer als Admin
   * sieht, dass ein Ticket liegen bleibt, konnte bisher nur privat nachfassen.
   *
   * Bewusst ohne Tages-Sperre: Ein Mensch entscheidet hier, nicht ein Timer.
   * Der Tagesschlüssel wird aber GESETZT, damit die Automatik nicht noch eine
   * zweite Erinnerung hinterherschickt.
   */
  const remindPowerUsers = React.useCallback(async (): Promise<{ ok: boolean; count: number; recipients: string[]; reason?: 'no-tickets' | 'no-recipients' | 'error' }> => {
    const openList = powerUserQueue.filter(t => t.status !== 'Closed' && !t.answeredAt);
    if (openList.length === 0) return { ok: false, count: 0, recipients: [], reason: 'no-tickets' };
    const to = reminderRecipients(currentUser.email);
    if (to.emails.length === 0) return { ok: false, count: openList.length, recipients: [], reason: 'no-recipients' };
    try {
      await queueTicketReminder(openList, to, false);
      const now = new Date();
      const dayKey = `dex_ticketreminder_${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()}`;
      try { window.localStorage.setItem(dayKey, '1'); } catch { /* */ }
      return { ok: true, count: openList.length, recipients: to.names };
    } catch (e) {
      console.warn('[DEX] manuelle Ticket-Erinnerung fehlgeschlagen:', e);
      return { ok: false, count: openList.length, recipients: to.names, reason: 'error' };
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [powerUserQueue, reminderRecipients, queueTicketReminder, currentUser]);

  /** v29.14: Wen würde der Erinnerungs-Knopf erreichen? Für die Rückfrage. */
  const reminderTargets = React.useCallback((): string[] => reminderRecipients(currentUser.email).names,
    [reminderRecipients, currentUser.email]);

  const value = React.useMemo<TicketContextType>(() => ({
    tickets, myTickets, isTicketsLoading, canAnswerTickets,
    reloadTickets, reloadMyTickets, createTicket, claimTicket, releaseTicket, answerTicket,
    replyToTicket, closeTicketNoAnswer,
    ticketsForEvent, openCountForEvent, powerUserQueue,
    remindPowerUsers, reminderTargets,
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [tickets, myTickets, isTicketsLoading, canAnswerTickets, powerUserQueue, ticketsForEvent, openCountForEvent, createTicket, answerTicket, replyToTicket, closeTicketNoAnswer, reloadTickets, reloadMyTickets, claimTicket, releaseTicket, remindPowerUsers, reminderTargets]);

  return React.createElement(TicketContext.Provider, { value }, props.children);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function toNamesAskerLabel(u: any): string {
  const name = `${u.firstName || ''} ${u.surname || ''}`.trim();
  const email = u.email || '';
  const loc = u.location ? ` · ${u.location}` : '';
  return `${name || email}${email && name ? ` (${email})` : ''}${loc}`;
}

export function useTickets(): TicketContextType {
  const ctx = React.useContext(TicketContext);
  if (!ctx) throw new Error('useTickets must be used within TicketProvider');
  return ctx;
}
