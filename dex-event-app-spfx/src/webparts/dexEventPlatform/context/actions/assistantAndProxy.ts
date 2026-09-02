/**
 * Assistenz und Stellvertretung: Delegation einer Anmeldung, Merker fuer
 * Fremd-Anmeldungen und die Aenderungs-/Kuendigungs-Antraege der Assistenz.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktionskoerper sind unveraendert - statt aus der Provider-Closure
 * beziehen sie ihre Umgebung aus dem `deps`-Objekt.
 */

import { DeloitteEvent } from '../../types';
import { EventService, SPRegistration, AssistantLink } from '../../services/EventService';
import { buildHashDeepLink } from '../../utils/deepLink';
import { wrapTemplate } from '../../services/EmailTemplates';

export interface AssistantDeps {
  eventService: EventService;
  events: DeloitteEvent[];
  subsiteMap: { current: Record<string, string> };
  currentUserEmail: string;
  currentUserName: string;
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
}

export function makeAssistantActions(deps: AssistantDeps) {
  const { eventService, events, subsiteMap, currentUserEmail, currentUserName, childEventsOf } = deps;

  // v24.41: Delegation an eine Assistenz — pro Anmelde-Zeile (Haupt-/Klammer-
  // Event + alle Sub-Events, für die der User angemeldet ist) einen Auftrag in
  // DEX_AssistantAccess anlegen. Der Flow setzt darauf Zeilen-Autor +
  // RegisteredBy auf die Assistenz. Best-effort: blockt die Anmeldung nie.
  async function delegateRegistrationToAssistant(eventId: string, assistant: { email: string; name: string }): Promise<void> {
    const assistEmail = (assistant?.email || '').trim();
    if (!assistEmail) return;
    if (assistEmail.toLowerCase() === (currentUserEmail || '').toLowerCase()) return; // nicht sich selbst
    const assistName = (assistant?.name || '').trim() || assistEmail;
    const participantName = (currentUserName || '').trim() || currentUserEmail;
    // Ziel-Events: Hauptevent + (falls Klammer) alle Sub-Events.
    const targetEventIds = [eventId, ...childEventsOf(eventId).map(c => c.id)];
    for (const evId of targetEventIds) {
      const subsiteUrl = subsiteMap.current[evId];
      if (!subsiteUrl) continue;
      let myReg: SPRegistration | null = null;
      try { myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail); } catch { myReg = null; }
      if (!myReg || !myReg.Id) continue;
      // Nur aktive Zeilen delegieren (keine abgemeldeten).
      if ((myReg.Status || '') === 'Abgemeldet') continue;
      const ev = events.find(e => e.id === evId);
      // Szenario A (ohne Flow): KEIN Owner-Wechsel — der Anmelder bleibt
      // Eigentümer (Owner) und verwaltet selbst. Es wird nur ein Info-
      // Verknüpfungs-Eintrag angelegt, über den die Assistenz die Anmeldung als
      // INFO sieht und eine Änderungs-/Abmelde-Anforderung stellen kann.
      try {
        await eventService.queueAssistantAccess({
          subsiteUrl,
          itemId: myReg.Id,
          eventId: evId,
          eventTitle: ev?.title || '',
          participantEmail: currentUserEmail,
          participantName,
          assistantEmail: assistEmail,
          assistantName: assistName,
          ownerEmail: currentUserEmail,
          linkType: 'delegation',
        });
      } catch (err) { console.warn('[DEX] queueAssistantAccess failed:', err); }
    }
  }

  // v24.41 Szenario B: Nach einer stellvertretenden Anmeldung (Assistenz meldet
  // eine andere Person an) einen Info-Link anlegen — der ANMELDER ist Owner und
  // verwaltet, die angemeldete Person sieht die Anmeldung als INFO unter „Meine
  // Events" und kann eine Anforderung stellen.
  async function recordProxyDelegation(eventId: string, participant: { email: string; name: string }): Promise<void> {
    const partEmail = (participant?.email || '').trim();
    if (!partEmail) return;
    if (partEmail.toLowerCase() === (currentUserEmail || '').toLowerCase()) return; // keine Selbst-Anmeldung
    const partName = (participant?.name || '').trim() || partEmail;
    const ownerName = (currentUserName || '').trim() || currentUserEmail;
    const targetEventIds = [eventId, ...childEventsOf(eventId).map(c => c.id)];
    for (const evId of targetEventIds) {
      const subsiteUrl = subsiteMap.current[evId];
      if (!subsiteUrl) continue;
      let reg: SPRegistration | null = null;
      try { reg = await eventService.getMyRegistration(subsiteUrl, partEmail); } catch { reg = null; }
      if (!reg || !reg.Id || (reg.Status || '') === 'Abgemeldet') continue;
      const ev = events.find(e => e.id === evId);
      try {
        await eventService.queueAssistantAccess({
          subsiteUrl,
          itemId: reg.Id,
          eventId: evId,
          eventTitle: ev?.title || '',
          participantEmail: partEmail,
          participantName: partName,
          assistantEmail: currentUserEmail,   // der Anmelder = verknüpfte „Assistenz"/Owner
          assistantName: ownerName,
          ownerEmail: currentUserEmail,
          linkType: 'proxy',
        });
      } catch (err) { console.warn('[DEX] recordProxyDelegation queue failed:', err); }
    }
  }

  async function getMyAssistantLinks(): Promise<AssistantLink[]> {
    try { return await eventService.getAssistantLinksForUser(currentUserEmail); }
    catch { return []; }
  }

  async function requestAssistantChange(link: AssistantLink, requestType: 'change' | 'cancel', note: string): Promise<boolean> {
    if (!eventService || !link?.id) return false;
    const requesterName = (currentUserName || '').trim() || currentUserEmail;
    const ok = await eventService.setAssistantLinkRequest(link.id, {
      requestType, note,
      requestedByEmail: currentUserEmail,
      requestedByName: requesterName,
    });
    if (!ok) return false;
    // Deeplink-Mail an den OWNER (wer die Anmeldung verwaltet), damit er die
    // Änderung/Abmeldung ausführt. Best-effort — gated über DisableEmails des
    // Events (falls auffindbar).
    try {
      const ev = events.find(e => e.id === link.eventId);
      if (ev && ev.disableEmails) return true; // Mails aus → nur die Liste, keine Mail
      const isDe = !(ev && ev.emailLanguage === 'EN');
      const actionLabel = requestType === 'cancel'
        ? (isDe ? 'Abmeldung' : 'Cancellation')
        : (isDe ? 'Änderung der Angaben' : 'Change of details');
      const deepLink = buildHashDeepLink(`${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`, { action: 'assistreq', id: link.id });
      const esc = (s: string): string => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const heading = isDe ? 'Anforderung an dich' : 'Request for you';
      const sub = isDe ? `${actionLabel} — ${link.eventTitle || ''}` : `${actionLabel} — ${link.eventTitle || ''}`;
      const body = isDe
        ? `<p>Hallo,</p>
           <p><strong>${esc(requesterName)}</strong> bittet dich um eine <strong>${esc(actionLabel)}</strong> für die folgende Anmeldung, die DU verwaltest:</p>
           <p style="margin:12px 0;padding:10px 14px;background:#f3f7ec;border-radius:8px;">
             <strong>Event:</strong> ${esc(link.eventTitle || link.eventId)}<br/>
             <strong>Angemeldete Person:</strong> ${esc(link.participantName || link.participantEmail)}<br/>
             ${note ? `<strong>Anmerkung:</strong> ${esc(note)}` : ''}
           </p>
           <p>Bitte führe die ${esc(actionLabel)} in der DEX-App aus (deine „Assistenz"-Kachel bzw. „Meine Events"):</p>
           <p><a href="${deepLink}" style="display:inline-block;background:#86bc25;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:600;">In DEX öffnen</a></p>
           <p>Vielen Dank!</p>`
        : `<p>Hello,</p>
           <p><strong>${esc(requesterName)}</strong> asks you for a <strong>${esc(actionLabel)}</strong> for the following registration that YOU manage:</p>
           <p style="margin:12px 0;padding:10px 14px;background:#f3f7ec;border-radius:8px;">
             <strong>Event:</strong> ${esc(link.eventTitle || link.eventId)}<br/>
             <strong>Registered person:</strong> ${esc(link.participantName || link.participantEmail)}<br/>
             ${note ? `<strong>Note:</strong> ${esc(note)}` : ''}
           </p>
           <p>Please perform the ${esc(actionLabel)} in the DEX app (your „Assistant" tile or „My Events"):</p>
           <p><a href="${deepLink}" style="display:inline-block;background:#86bc25;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:600;">Open in DEX</a></p>
           <p>Many thanks!</p>`;
      const subject = isDe
        ? `Anforderung: ${actionLabel} — ${link.eventTitle || ''}`
        : `Request: ${actionLabel} — ${link.eventTitle || ''}`;
      const wrapped = wrapTemplate('#86bc25', heading, sub, body);
      await eventService.queueEmail(subject, link.ownerEmail, link.ownerEmail, wrapped, 'Info', link.eventTitle || '', link.eventId);
    } catch (err) { console.warn('[DEX] requestAssistantChange mail failed:', err); }
    return true;
  }

  async function resolveAssistantRequest(linkId: number, decision: 'Done' | 'Rejected'): Promise<boolean> {
    if (!eventService) return false;
    return eventService.resolveAssistantLinkRequest(linkId, decision);
  }

  return { delegateRegistrationToAssistant, recordProxyDelegation, getMyAssistantLinks, requestAssistantChange, resolveAssistantRequest };
}
