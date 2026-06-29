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
import { useRoles } from './RoleContext';
import { useCurrentUser } from './UserContext';
import { useEvents } from './EventContext';
import { useNavigation } from './NavigationContext';

const GREEN = '#86bc25';

export interface AskInput {
  questions: string[];
  /** Optionale Screenshots des aktuellen Bildschirms (ask_-Anhänge). */
  screenshots: File[];
}

export interface AnswerInput {
  ticket: DexTicket;
  answerText: string;
  /** In der Antwort verlinkte Handbuch-Artikel (id + aufgelöster Titel). */
  articles: Array<{ id: string; title: string }>;
  /** 1-basierter Event-Wizard-Schritt (null = keiner). */
  wizardStep: number | null;
  wizardStepLabel?: string;
  /** Optionale Screenshots (ans_-Anhänge), z.B. der Wizard-Schritt. */
  screenshots: File[];
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
  claimTicket: (ticketId: number) => Promise<void>;
  releaseTicket: (ticketId: number) => Promise<void>;
  answerTicket: (input: AnswerInput) => Promise<boolean>;
  /** Tickets eines normalen Users zu einem konkreten Event (Audience=Organizer). */
  ticketsForEvent: (eventId: string) => DexTicket[];
  /** Anzahl noch offener (nicht geschlossener) User-Fragen zu einem Event. */
  openCountForEvent: (eventId: string) => number;
  /** Power-User-Warteschlange (Fragen von Organizern/Admins). */
  powerUserQueue: DexTicket[];
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

  const appBase = `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`;

  // ---- Mail-Bausteine -----------------------------------------------------
  const noReplyHintHtml = `<div style="margin-top:18px;padding:12px 14px;background:#fff3e0;border:1px solid #ed8b00;border-radius:6px;color:#8a4b00;font-size:13px;">
    <strong>Bitte nicht auf diese E-Mail antworten.</strong> Diese Adresse wird nicht gelesen — bearbeite das Ticket direkt in der App über den Button oben.</div>`;

  function ctaButton(href: string, label: string): string {
    return `<p style="margin:22px 0 0;"><a href="${href}" style="display:inline-block;background:${GREEN};color:#ffffff;text-decoration:none;padding:11px 24px;border-radius:6px;font-weight:600;font-size:14px;">${esc(label)}</a></p>`;
  }

  // ---- Aktionen -----------------------------------------------------------
  const createTicket = React.useCallback(async (input: AskInput): Promise<boolean> => {
    const questions = (input.questions || []).map(q => (q || '').trim()).filter(Boolean);
    if (questions.length === 0) return false;
    const askerRole = isAdmin ? 'Admin' : (roles.find(r => (r.userEmail || '').toLowerCase() === myEmailLc)?.role || 'User');
    const askerIsOrganizerLike = askerRole === 'Organizer' || askerRole === 'Admin';

    // Event-Kontext (falls auf einer Event-Seite gefragt wurde).
    const ctxEvent = selectedEventId ? (events || []).find(e => e.id === selectedEventId) : undefined;

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

    const pageContext = (() => { try { return new URLSearchParams(window.location.search).get('action') || window.location.pathname; } catch { return ''; } })();

    const id = await eventService.createTicket({
      questions,
      askerEmail: currentUser.email || '',
      askerName: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || (currentUser.email || ''),
      askerRole,
      audience, eventId, eventTitle, assignedOrganizers, pageContext,
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
    if (audience === 'PowerUser') {
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
        ? `${appBase}&action=tickets&id=${id}`
        : `${appBase}&action=admin&event=${encodeURIComponent(eventId)}&ticket=${id}`;
      const qHtml = questions.length === 1
        ? `<p style="margin:6px 0 0;">${esc(questions[0])}</p>`
        : `<ul style="margin:8px 0 0 18px;padding:0;">${questions.map(q => `<li style="margin-bottom:6px;">${esc(q)}</li>`).join('')}</ul>`;
      const inner = `
        <p style="margin:0 0 6px;">Hallo,</p>
        <p style="margin:0 0 16px;">im DEX-Ticketsystem ist eine neue Frage eingegangen${eventTitle ? ` (Event: <strong>${esc(eventTitle)}</strong>)` : ''}.</p>
        <p style="margin:0;"><strong>Von:</strong> ${esc(toNamesAskerLabel(currentUser))}</p>
        <p style="margin:14px 0 0;"><strong>Frage${questions.length > 1 ? 'n' : ''}:</strong></p>
        ${qHtml}
        ${ctaButton(link, 'Ticket öffnen & beantworten')}
        ${noReplyHintHtml}
      `;
      const subject = questions.length === 1
        ? `Neue Frage im DEX-Ticketsystem`
        : `Neue Fragen (${questions.length}) im DEX-Ticketsystem`;
      const body = wrapTemplate(GREEN, 'Neue Frage im Ticketsystem', eventTitle || 'DEX-Support', inner);
      try {
        await eventService.queueEmail(
          subject, toEmails.join('; '), toNames.join('; '), body,
          'TicketNew', eventTitle || 'DEX-Ticket', eventId || '0', cc || undefined
        );
      } catch { /* Mail best-effort */ }
    }

    // States aktualisieren.
    reloadMyTickets().catch(() => { /* */ });
    if (shouldLoadAll) reloadTickets().catch(() => { /* */ });
    return true;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eventService, roles, isAdmin, myEmailLc, currentUser, events, selectedEventId, shouldLoadAll]);

  const claimTicket = React.useCallback(async (ticketId: number): Promise<void> => {
    const name = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || (currentUser.email || '');
    await eventService.claimTicket(ticketId, currentUser.email || '', name);
    await reloadTickets();
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
      answeredByEmail: currentUser.email || '',
      answeredByName,
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
           <ul style="margin:8px 0 0 18px;padding:0;">${(input.articles || []).map(a => `<li style="margin-bottom:4px;"><a href="${appBase}&action=manual&section=${encodeURIComponent(a.id)}" style="color:${GREEN};font-weight:600;">${esc(a.title)}</a></li>`).join('')}</ul>`
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
        <p style="margin:20px 0 0;color:#666;font-size:13px;">Du findest diese Antwort jederzeit in der App über den grünen Button <strong>&bdquo;Hast du Fragen?&ldquo;</strong> (oben rechts) unter <strong>&bdquo;Deine Fragen&ldquo;</strong>. Falls noch etwas offen ist, stelle dort einfach eine neue Frage.</p>
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

  const ticketsForEvent = React.useCallback((eventId: string): DexTicket[] => {
    if (!eventId) return [];
    return tickets.filter(t => t.eventId === eventId && t.audience === 'Organizer');
  }, [tickets]);

  const openCountForEvent = React.useCallback((eventId: string): number => {
    return ticketsForEvent(eventId).filter(t => t.status !== 'Closed').length;
  }, [ticketsForEvent]);

  const powerUserQueue = React.useMemo(() => tickets.filter(t => t.audience === 'PowerUser'), [tickets]);

  const value = React.useMemo<TicketContextType>(() => ({
    tickets, myTickets, isTicketsLoading, canAnswerTickets,
    reloadTickets, reloadMyTickets, createTicket, claimTicket, releaseTicket, answerTicket,
    ticketsForEvent, openCountForEvent, powerUserQueue,
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [tickets, myTickets, isTicketsLoading, canAnswerTickets, powerUserQueue, ticketsForEvent, openCountForEvent, createTicket, answerTicket, reloadTickets, reloadMyTickets, claimTicket, releaseTicket]);

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
