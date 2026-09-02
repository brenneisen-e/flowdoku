/**
 * Abmelden und Stellvertretung: eigene Abmeldung, Team-Mitglied abmelden,
 * Proxy-Anmeldungen verwalten und die proaktive Absage.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktionskoerper sind unveraendert - statt aus der Provider-Closure
 * beziehen sie ihre Umgebung aus dem `deps`-Objekt.
 */

import { DeloitteEvent } from '../../types';
import { EventService, SPRegistration } from '../../services/EventService';
import { wrapTemplate, buildEmailFromTemplate, cancellationEmail } from '../../services/EmailTemplates';
import { applyEventTemplateOverride, collectCcEmailsFromFields, mergeCcLists } from '../eventTextHelpers';
import { isEventOver } from '../../utils/eventFormat';
import { isDemoShowcaseId } from '../../services/demoShowcaseEvent';
import { withParentTitleSubject } from '../../utils/mailSubject';

export interface CancellationDeps {
  eventService: EventService;
  events: DeloitteEvent[];
  subsiteMap: { current: Record<string, string> };
  currentUserEmail: string;
  currentUserName: string;
  currentUserFirstName: string;
  calDayParentOf: (ev: { parentEventId?: string } | undefined) => DeloitteEvent | undefined;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  loadEvents: () => Promise<void>;
}

export function makeCancellationActions(deps: CancellationDeps) {
  const { eventService, events, subsiteMap, currentUserEmail, currentUserName, currentUserFirstName, calDayParentOf, getMyRegistration, loadEvents } = deps;

  async function cancelRegistration(eventId: string, opts?: { suppressNotifications?: boolean; skipReload?: boolean }): Promise<boolean> {
    // v17.25: Demo-Showcase-Event → No-Op.
    if (isDemoShowcaseId(eventId)) return true;
    // v22.22: Selbst-Abmeldung von bereits vergangenen Events ist gesperrt —
    // gilt damit auch für den Auto-Cancel-Deep-Link aus Mails und die
    // Sub-Event-Session-Toggles in Meine Events. Organizer/Admins melden
    // über das Admin Center ab (eventService.cancelRegistration direkt) —
    // dieser Pfad ist bewusst NICHT betroffen, läuft bei vergangenen Events
    // aber still (ohne Mail/Outlook/Nachrücken, siehe AdminPage).
    const evForGuard = events.find(e => e.id === eventId);
    if (evForGuard && isEventOver(evForGuard)) {
      console.warn('[DEX] cancelRegistration blocked: event already in the past', eventId);
      return false;
    }
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;

    const myReg = await eventService.getMyRegistration(subsiteUrl, currentUserEmail);
    if (!myReg) return false;

    // Audit: wer klickt gerade 'Abmelden'? = der eingeloggte User. Bei Self-Cancel
    // ist das = der Teilnehmer selbst. Aus der App heraus gibt's aktuell keinen
    // "Abmeldung für andere"-Pfad (das macht der Organizer/Admin über AdminPage,
    // dort wird eventService.cancelRegistration direkt aufgerufen).
    // v11.53: vorherigen Status merken, damit wir den KPI-Counter nur dann
    // dekrementieren, wenn der User tatsächlich 'Angemeldet' war (Wartelist-
    // Cancel berührt den Teilnehmer-KPI nicht).
    const wasActive = myReg.Status === 'Angemeldet';
    // v11.83: Team-Anmeldungs-Kontext snapshotten, BEVOR der eigene Status
    // auf 'Abgemeldet' kippt — danach liefert getTeamMembers den eigenen
    // Eintrag schon mit dem alten Lead-Flag aus und der Promote-Pfad
    // verlässt sich nicht mehr darauf. Wir speichern hier den eigenen
    // TeamId/TeamLead/TeamName-Stand und filtern nach dem Cancel die
    // verbleibenden Mitglieder.
    const wasTeamCancel = !!myReg.TeamId;
    const wasTeamLead = wasTeamCancel && !!myReg.TeamLead;
    const teamId = myReg.TeamId || '';
    const teamName = myReg.TeamName || '';
    const success = await eventService.cancelRegistration(subsiteUrl, myReg.Id, currentUserName, currentUserEmail);
    if (success) {
      const event = events.find(e => e.id === eventId);
      if (wasActive) {
        eventService.bumpKpiParticipants(-1).catch(() => { /* best-effort */ });
      }
      // v9.0: Audit-Log (fire-and-forget)
      eventService.writeChangeLog({
        action: 'ParticipantCancelled',
        targetType: 'Participant',
        targetId: currentUserEmail,
        targetName: currentUserName,
        eventId: eventId,
        eventTitle: event?.title || '',
        details: { participantId: myReg.Id, asActor: 'self' },
      }).catch(() => { /* */ });
      if (event) {
        // Dual-Write: DEX_Participants aktualisieren
        if (event.eventNumber) {
          try {
            await eventService.removeParticipantEvent(currentUserEmail, event.eventNumber);
          } catch (err) { console.warn('[DEX] removeParticipantEvent failed:', err); }
        }
        // Abmelde-E-Mail in Queue eintragen (SharePoint-Template, Fallback auf Code-Template)
        // v17.19/v17.22: Notifications werden NUR unterdrückt, wenn der Aufrufer
        // das explizit anfordert (`opts.suppressNotifications`). Das passiert
        // ausschliesslich beim automatischen Schatten-Parent-Cancel im
        // subEventsOnlyMode (MyEventsPage: letzte Sub-Event-Abmeldung räumt
        // den Schatten-Parent ab — die Sub-Event-Abmeldung hat ihre eigene
        // Mail schon verschickt). v17.22-Fix: vorher wurde pauschal auf
        // `event.subEventsOnlyMode` geprüft, wodurch Alt-Anmeldungen (User
        // hat sich noch im Normal-Modus direkt beim Parent angemeldet, bevor
        // der Organizer das Event auf subEventsOnlyMode umstellte) beim
        // direkten Abmelden weder Bestätigungs-Mail noch Outlook-Ausladen
        // bekamen.
        const suppressParentNotificationsCancel = !!opts?.suppressNotifications;
        // v19.21: disableCancellationEmail = nur die Abmelde-Bestätigung
        // unterdrücken (granulares Sub-Häkchen unter dem Master „E-Mails").
        if (!event.disableEmails && !event.disableCancellationEmail && !suppressParentNotificationsCancel) {
          try {
            const lang = event.emailLanguage || 'EN';
            // {{Name}} in Anreden: nur Vorname (displayName ist im Deloitte-Tenant
            // "Nachname, Vorname" -> getFirstName extrahiert den Vornamen).
            const cancelVars = { Name: currentUserFirstName, EventTitle: event.title, AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView` };
            let emailData: { subject: string; body: string };
            const spTplRaw = await eventService.getEmailTemplate('Abmeldung', lang).catch(() => null);
            const spTpl = applyEventTemplateOverride(spTplRaw, event.emailTemplateOverrides, 'Abmeldung');
            if (spTpl) {
              emailData = buildEmailFromTemplate(spTpl, cancelVars);
            } else {
              emailData = cancellationEmail(currentUserFirstName, event.title);
            }
            // v8.5: Organizer-Mitlesen bei Abmeldung auswerten. 'always' = immer,
            // 'afterDeadline' = nur wenn lastDeregisterDate überschritten ist.
            // v28.28: als CC statt BCC (siehe Anmelde-Pfad).
            let orgCopyCc = '';
            const mode = event.notifyOrgCancelMode || 'never';
            if (mode === 'always' || (mode === 'afterDeadline' && event.lastDeregisterDate && new Date() > new Date(event.lastDeregisterDate))) {
              const orgEmails = (event.organizerEmails || []).filter(Boolean);
              if (orgEmails.length > 0) orgCopyCc = orgEmails.join(';');
            }
            // v18.41: People-Picker-Felder mit „CC bei Mail" → ausgewählte
            // Person(en) auch bei der Abmelde-Mail auf CC (nicht im Outlook-Termin).
            // v18.58: Im subEventsOnlyMode liegen die übergreifenden CC-Felder
            // (z.B. Assistenz) in der Schatten-Parent-Registrierung, NICHT in der
            // Sub-Event-Registrierung. Beim Abmelden einer Section daher
            // zusätzlich die CC aus dem Parent-Event nachladen und mergen, damit
            // die Assistenz auch die Abmelde-Mail in Kopie bekommt (Anmeldung
            // war bereits via extraCc abgedeckt).
            let cancelCc: string | undefined;
            try {
              const cd = myReg.CustomData ? JSON.parse(myReg.CustomData) as Record<string, string> : {};
              const ownCc = collectCcEmailsFromFields(event.eventSpecificFields, cd, currentUserEmail);
              let parentCc = '';
              if (event.parentEventId) {
                const parentEvent = events.find(ev => ev.id === event.parentEventId);
                if (parentEvent && parentEvent.subEventsOnlyMode) {
                  try {
                    const parentReg = await getMyRegistration(event.parentEventId);
                    const pcd = parentReg?.CustomData ? JSON.parse(parentReg.CustomData) as Record<string, string> : {};
                    parentCc = collectCcEmailsFromFields(parentEvent.eventSpecificFields, pcd, currentUserEmail);
                  } catch { /* parent-CC best-effort */ }
                }
              }
              const seen = new Set<string>();
              const merged: string[] = [];
              for (const part of [ownCc, parentCc, orgCopyCc]) {
                for (const em of part.split(';').map(s => s.trim()).filter(Boolean)) {
                  const lc = em.toLowerCase();
                  if (lc !== (currentUserEmail || '').toLowerCase() && !seen.has(lc)) { seen.add(lc); merged.push(em); }
                }
              }
              cancelCc = merged.length ? merged.join(';') : undefined;
            } catch { cancelCc = orgCopyCc || undefined; }
            const emailOk = await eventService.queueEmail(
              withParentTitleSubject(emailData.subject, calDayParentOf(event)), currentUserEmail, currentUserName, emailData.body,
              'Abmeldung', event.title, eventId, cancelCc, undefined
            );
            if (!emailOk) console.warn('[DEX] queueEmail for cancellation returned false');
          } catch (err) { console.warn('[DEX] queueEmail for cancellation failed:', err); }
        }
        // Outlook-Termin-Einladung zurückziehen
        if (!event.disableOutlook && !suppressParentNotificationsCancel) {
          try {
            await eventService.queueOutlookEvent(
              currentUserEmail, eventId, event.title, 'Ausladen'
            );
          } catch (err) { console.warn('[DEX] queueOutlookEvent failed:', err); }
        }
        // Nachrücken wird komplett vom Power-Automate-Flow DEX_IDReorder_TeilnehmerIDs
        // übernommen (seit v6.7). Der Flow ist typen-bewusst für B2Run-Split-
        // Wartelisten: er promotet den ersten Warteliste-Teilnehmer mit passendem
        // PreferredStarterType und verschickt Nachrück-Mail + Outlook-Einladung.
        // Die App macht nur noch Abmeldung + IDReorder-Queue-Trigger — keine
        // parallele Client-Promote-Logik mehr (die vorher zu Race-Conditions mit
        // dem Flow geführt hat).
        // ID-Reorder in Queue eintragen (triggert den DEX_IDReorder-Flow, der
        // danach ID-Neuvergabe + Nachrücken abwickelt).
        if (subsiteUrl) {
          try {
            // v24.71: CancelledName als „Vorname Nachname" schreiben (aus der
            // Registrierung), nicht den SharePoint-Anzeigenamen currentUserName
            // („Nachname, Vorname") — sonst stehen in der Organizer-Nachrück-Mail
            // die Namen uneinheitlich (z.B. „Obermeier, Katrin" abgemeldet vs.
            // „Alexa Sophie Gedaschko" nachgerückt). Der Flow baut PromotedName
            // ebenfalls als „Vorname Nachname".
            const cancelledDisplayName = (myReg.Vorname && myReg.Nachname)
              ? `${myReg.Vorname} ${myReg.Nachname}`
              : (myReg.ParticipantName || currentUserName);
            const reorderOk = await eventService.queueIDReorder(
              eventId, event.eventNumber || 0, subsiteUrl, event.title, cancelledDisplayName, currentUserEmail
            );
            if (!reorderOk) console.warn('[DEX] queueIDReorder returned false');
          } catch (err) { console.warn('[DEX] queueIDReorder failed:', err); }
          // v11.36 → v27.10: Sitzplatz-Counter nach der Abmeldung pflegen.
          // NICHT mehr blind syncSeatsToActiveCount — ein normaler User sieht
          // die Teilnehmerliste seit v26.87 nur beschnitten („nur eigene
          // Elemente") und hätte den Counter mit 0 überschrieben (Ursache der
          // Warteliste-Überholung vom 15.07.). releaseSeatAfterCancel macht
          // den Voll-Reconcile nur bei Vollzugriff und arbeitet sonst additiv.
          try {
            const isSplit = typeof event.durchstarterCapacity === 'number'
              && typeof event.funstarterCapacity === 'number'
              && ((event.durchstarterCapacity || 0) > 0 || (event.funstarterCapacity || 0) > 0);
            await eventService.releaseSeatAfterCancel(subsiteUrl, {
              isSplit,
              previousStatus: myReg.Status || '',
              starterType: myReg.StarterType || undefined,
              // v27.11: Warteliste abgeschaltet → niemand rückt nach, der
              // Platz muss direkt freigegeben werden (sonst Counter-Deadlock).
              waitlistDisabled: event.waitlistEnabled === false,
            });
          } catch { /* best-effort */ }
        }
      } else {
        console.warn('[DEX] cancelRegistration: event not found in state for id', eventId);
      }
      // v11.83: Team-Cancel-Nachlauf — Auto-Promote des früheren Members
      // zum neuen Lead (falls Self-Cancel der Lead war), Info-Mails an die
      // verbleibenden Mitglieder, Hinweis welche Optionen ihnen offenstehen.
      // Der Sitzplatz-Counter wird im normalen Reconcile oben schon
      // dekrementiert — der frei werdende Platz darf von anderen Teilnehmern
      // belegt werden (oder vom Team-Lead nachbesetzt werden, siehe
      // addTeamMember). Die App entscheidet hier bewusst NICHT, ob der
      // Slot für das Team reserviert bleibt — das passt zur Beschreibung
      // im Spec, weil der frei werdende Sitz neutral ist: der Team-Lead
      // kann ihn über "Mitglied hinzufügen" wieder füllen, ansonsten
      // landet er in der normalen Sitzplatz-Verwaltung.
      if (wasTeamCancel && event) {
        await handleTeamCancelPostStep(event, eventId, subsiteUrl, teamId, teamName, wasTeamLead, myReg).catch(err => {
          console.warn('[DEX] team-cancel post-step failed:', err);
        });
      }
      // v30.9: s. registerForEvent — Schleifen skippen, EIN Refresh am Ende.
      if (!opts?.skipReload) await loadEvents();
    }
    return success;
  }

  /**
   * v11.86: Team-Lead meldet stellvertretend ein Team-Mitglied vom Event
   * ab — ausgelöst aus dem „Team verwalten"-Modal in MyEvents. Audit
   * wird auf den eingeloggten Lead geschrieben (CancelledByName/Email),
   * danach läuft derselbe Team-Post-Step wie beim Self-Cancel:
   * Sitzplatz-Reconcile, IDReorder-Queue, Outlook-Ausladung,
   * Abmelde-Bestätigung an die abgemeldete Person und Info-Mails an die
   * uebrigen Team-Mitglieder. Der Lead darf sich über diesen Pfad
   * NICHT selbst löschen — das übernimmt der normale Self-Cancel über
   * `cancelRegistration` (inkl. Auto-Promote des frühesten Members).
   */
  async function cancelTeamMember(
    eventId: string,
    memberRegistration: SPRegistration
  ): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl || !memberRegistration?.Id) return false;
    if (!memberRegistration.TeamId) return false;
    // Self-Schutz: der Lead löscht sich nicht über diesen Pfad — sein
    // eigener Cancel läuft via cancelRegistration mit Auto-Promote.
    if ((memberRegistration.ParticipantEmail || '').toLowerCase() === (currentUserEmail || '').toLowerCase()) {
      console.warn('[DEX] cancelTeamMember: Lead cannot cancel itself via this path');
      return false;
    }
    const wasActive = memberRegistration.Status === 'Angemeldet';
    const teamId = memberRegistration.TeamId;
    const teamName = memberRegistration.TeamName || '';
    // Audit = der eingeloggte Lead (stellvertretender Cancel).
    const ok = await eventService.cancelRegistration(
      subsiteUrl, memberRegistration.Id, currentUserName, currentUserEmail
    );
    if (!ok) return false;
    const event = events.find(e => e.id === eventId);
    if (wasActive) {
      eventService.bumpKpiParticipants(-1).catch(() => { /* */ });
    }
    eventService.writeChangeLog({
      action: 'ParticipantCancelled',
      targetType: 'Participant',
      targetId: memberRegistration.ParticipantEmail,
      targetName: `${memberRegistration.Vorname || ''} ${memberRegistration.Nachname || ''}`.trim() || memberRegistration.ParticipantEmail,
      eventId: eventId,
      eventTitle: event?.title || '',
      details: { participantId: memberRegistration.Id, asActor: 'teamLead', actorEmail: currentUserEmail },
    }).catch(() => { /* */ });
    if (event) {
      // Dual-Write: DEX_Participants aktualisieren
      if (event.eventNumber) {
        try {
          await eventService.removeParticipantEvent(memberRegistration.ParticipantEmail, event.eventNumber);
        } catch (err) { console.warn('[DEX] removeParticipantEvent failed:', err); }
      }
      // Abmelde-Mail an die abgemeldete Person.
      // v19.21: disableCancellationEmail unterdrückt auch hier die Abmelde-Mail.
      if (!event.disableEmails && !event.disableCancellationEmail) {
        try {
          const lang = event.emailLanguage || 'EN';
          const cancelledFirst = memberRegistration.Vorname
            || (memberRegistration.ParticipantName || '').split(/[ ,]+/)[0]
            || '';
          const cancelVars = {
            Name: cancelledFirst,
            EventTitle: event.title,
            AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
          };
          let emailData: { subject: string; body: string };
          const spTplRaw = await eventService.getEmailTemplate('Abmeldung', lang).catch(() => null);
          const spTpl = applyEventTemplateOverride(spTplRaw, event.emailTemplateOverrides, 'Abmeldung');
          if (spTpl) {
            emailData = buildEmailFromTemplate(spTpl, cancelVars);
          } else {
            emailData = cancellationEmail(cancelledFirst, event.title);
          }
          // v28.28: Organizer-Mitlese-Kopie als CC statt BCC (s.o.).
          let orgCopyCc = '';
          const mode = event.notifyOrgCancelMode || 'never';
          if (mode === 'always' || (mode === 'afterDeadline' && event.lastDeregisterDate && new Date() > new Date(event.lastDeregisterDate))) {
            const orgEmails = (event.organizerEmails || []).filter(Boolean);
            if (orgEmails.length > 0) orgCopyCc = orgEmails.join(';');
          }
          // v18.41: CC-Felder der abgemeldeten Person berücksichtigen.
          let memberCc: string | undefined;
          try {
            const cd = memberRegistration.CustomData ? JSON.parse(memberRegistration.CustomData) as Record<string, string> : {};
            memberCc = collectCcEmailsFromFields(event.eventSpecificFields, cd, memberRegistration.ParticipantEmail) || undefined;
          } catch { memberCc = undefined; }
          memberCc = mergeCcLists(memberCc, orgCopyCc, memberRegistration.ParticipantEmail);
          await eventService.queueEmail(
            withParentTitleSubject(emailData.subject, calDayParentOf(event)),
            memberRegistration.ParticipantEmail,
            `${memberRegistration.Vorname || ''} ${memberRegistration.Nachname || ''}`.trim() || memberRegistration.ParticipantEmail,
            emailData.body,
            'Abmeldung', event.title, eventId, memberCc, undefined
          );
        } catch (err) { console.warn('[DEX] queueEmail for team-lead cancel failed:', err); }
      }
      // Outlook-Ausladung.
      if (!event.disableOutlook) {
        try {
          await eventService.queueOutlookEvent(
            memberRegistration.ParticipantEmail, eventId, event.title, 'Ausladen'
          );
        } catch (err) { console.warn('[DEX] queueOutlookEvent (team-lead cancel) failed:', err); }
      }
      // ID-Reorder + Sitzplatz-Sync.
      if (subsiteUrl) {
        try {
          await eventService.queueIDReorder(
            eventId, event.eventNumber || 0, subsiteUrl, event.title,
            `${memberRegistration.Vorname || ''} ${memberRegistration.Nachname || ''}`.trim() || memberRegistration.ParticipantName || undefined,
            memberRegistration.ParticipantEmail || undefined
          );
        } catch (err) { console.warn('[DEX] queueIDReorder (team-lead cancel) failed:', err); }
        // v27.10: ILS-sichere Counter-Pflege statt blindem Voll-Sync (der
        // Team-Lead ist ein normaler User — siehe Kommentar im Self-Cancel).
        try {
          const isSplit = typeof event.durchstarterCapacity === 'number'
            && typeof event.funstarterCapacity === 'number'
            && ((event.durchstarterCapacity || 0) > 0 || (event.funstarterCapacity || 0) > 0);
          await eventService.releaseSeatAfterCancel(subsiteUrl, {
            isSplit,
            previousStatus: memberRegistration.Status || '',
            starterType: memberRegistration.StarterType || undefined,
            waitlistDisabled: event.waitlistEnabled === false, // v27.11
          });
        } catch { /* best-effort */ }
      }
      // Info-Mails an die uebrigen Team-Mitglieder. Wir löschen NICHT
      // den Lead, daher `wasTeamLead = false` → kein Auto-Promote.
      await handleTeamCancelPostStep(event, eventId, subsiteUrl, teamId, teamName, false, memberRegistration)
        .catch(err => { console.warn('[DEX] team-cancel post-step (lead-initiated) failed:', err); });
    }
    await loadEvents();
    return true;
  }

  /**
   * v24.36: Assistenz — alle Fremd-Anmeldungen des eingeloggten Users über
   * alle Subsites (Haupt- + Sub-Events) einsammeln. Dedupliziert pro Subsite
   * (mehrere Events teilen sich keine Subsite, aber Sub-Events haben eigene),
   * sodass jede Teilnehmerliste nur einmal abgefragt wird.
   */
  async function getMyProxyRegistrations(): Promise<Array<{ event: DeloitteEvent; registration: SPRegistration }>> {
    const me = (currentUserEmail || '').toLowerCase().trim();
    if (!me) return [];
    // Subsite → Event-Liste (mehrere Events können theoretisch dieselbe Subsite
    // teilen — z.B. recreate-Pfad; dann ordnen wir die Treffer dem passenden
    // Event über die Event-Nummer in der Zeile nicht zu, sondern nehmen das
    // erste Event der Subsite. In der Praxis = 1:1 Event↔Subsite.)
    const subsiteToEvents = new Map<string, DeloitteEvent[]>();
    for (const ev of events) {
      const sub = subsiteMap.current[ev.id];
      if (!sub) continue;
      const arr = subsiteToEvents.get(sub) || [];
      arr.push(ev);
      subsiteToEvents.set(sub, arr);
    }
    const results: Array<{ event: DeloitteEvent; registration: SPRegistration }> = [];
    const subs = Array.from(subsiteToEvents.keys());
    // Parallel je Subsite abfragen (best-effort, Fehler je Subsite ignorieren).
    await Promise.all(subs.map(async (sub) => {
      let regs: SPRegistration[] = [];
      try { regs = await eventService.getProxyRegistrationsByActor(sub, me); }
      catch { regs = []; }
      if (regs.length === 0) return;
      const evs = subsiteToEvents.get(sub) || [];
      const primary = evs[0];
      if (!primary) return;
      for (const r of regs) {
        results.push({ event: primary, registration: r });
      }
    }));
    return results;
  }

  /**
   * v24.36: Stellvertretende Abmeldung einer Fremd-Anmeldung durch die
   * Assistenz. Spiegelt den Side-Effect-Pfad von `cancelTeamMember`
   * (Abmelde-Mail, Outlook-Ausladen, ID-Reorder, Sitzplatz-Sync), aber ohne
   * Team-Logik. Audit als Akteur „assistant".
   */
  async function cancelProxyRegistration(eventId: string, registration: SPRegistration): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl || !registration?.Id) return false;
    // Self-Schutz: über diesen Pfad meldet sich niemand selbst ab.
    if ((registration.ParticipantEmail || '').toLowerCase() === (currentUserEmail || '').toLowerCase()) {
      console.warn('[DEX] cancelProxyRegistration: cannot cancel own registration via this path');
      return false;
    }
    const event = events.find(e => e.id === eventId);
    // v22.22: Nach Event-Ende keine Abmeldung mehr.
    if (event && isEventOver(event)) {
      console.warn('[DEX] cancelProxyRegistration: event is over, blocked');
      return false;
    }
    const wasActive = registration.Status === 'Angemeldet';
    const ok = await eventService.cancelRegistration(
      subsiteUrl, registration.Id, currentUserName, currentUserEmail
    );
    if (!ok) return false;
    if (wasActive) {
      eventService.bumpKpiParticipants(-1).catch(() => { /* */ });
    }
    eventService.writeChangeLog({
      action: 'RegistrationCancelled',
      targetType: 'Participant',
      targetId: registration.ParticipantEmail,
      targetName: `${registration.Vorname || ''} ${registration.Nachname || ''}`.trim() || registration.ParticipantEmail,
      eventId: eventId,
      eventTitle: event?.title || '',
      details: { participantId: registration.Id, asActor: 'assistant', actorEmail: currentUserEmail },
    }).catch(() => { /* */ });
    if (event) {
      if (event.eventNumber) {
        try { await eventService.removeParticipantEvent(registration.ParticipantEmail, event.eventNumber); }
        catch (err) { console.warn('[DEX] removeParticipantEvent failed:', err); }
      }
      // Abmelde-Mail an die abgemeldete Person.
      if (!event.disableEmails && !event.disableCancellationEmail) {
        try {
          const lang = event.emailLanguage || 'EN';
          const cancelledFirst = registration.Vorname
            || (registration.ParticipantName || '').split(/[ ,]+/)[0] || '';
          const cancelVars = {
            Name: cancelledFirst,
            EventTitle: event.title,
            AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
          };
          let emailData: { subject: string; body: string };
          const spTplRaw = await eventService.getEmailTemplate('Abmeldung', lang).catch(() => null);
          const spTpl = applyEventTemplateOverride(spTplRaw, event.emailTemplateOverrides, 'Abmeldung');
          if (spTpl) emailData = buildEmailFromTemplate(spTpl, cancelVars);
          else emailData = cancellationEmail(cancelledFirst, event.title);
          // v28.28: Organizer-Mitlese-Kopie als CC statt BCC (s.o.).
          let orgCopyCc = '';
          const mode = event.notifyOrgCancelMode || 'never';
          if (mode === 'always' || (mode === 'afterDeadline' && event.lastDeregisterDate && new Date() > new Date(event.lastDeregisterDate))) {
            const orgEmails = (event.organizerEmails || []).filter(Boolean);
            if (orgEmails.length > 0) orgCopyCc = orgEmails.join(';');
          }
          let memberCc: string | undefined;
          try {
            const cd = registration.CustomData ? JSON.parse(registration.CustomData) as Record<string, string> : {};
            memberCc = collectCcEmailsFromFields(event.eventSpecificFields, cd, registration.ParticipantEmail) || undefined;
          } catch { memberCc = undefined; }
          memberCc = mergeCcLists(memberCc, orgCopyCc, registration.ParticipantEmail);
          await eventService.queueEmail(
            withParentTitleSubject(emailData.subject, calDayParentOf(event)),
            registration.ParticipantEmail,
            `${registration.Vorname || ''} ${registration.Nachname || ''}`.trim() || registration.ParticipantEmail,
            emailData.body,
            'Abmeldung', event.title, eventId, memberCc, undefined
          );
        } catch (err) { console.warn('[DEX] queueEmail for proxy cancel failed:', err); }
      }
      // Outlook-Ausladung.
      if (!event.disableOutlook) {
        try {
          await eventService.queueOutlookEvent(registration.ParticipantEmail, eventId, event.title, 'Ausladen');
        } catch (err) { console.warn('[DEX] queueOutlookEvent (proxy cancel) failed:', err); }
      }
      // ID-Reorder + Sitzplatz-Sync (treibt das Nachrücken der Warteliste).
      try {
        await eventService.queueIDReorder(
          eventId, event.eventNumber || 0, subsiteUrl, event.title,
          `${registration.Vorname || ''} ${registration.Nachname || ''}`.trim() || registration.ParticipantName || undefined,
          registration.ParticipantEmail || undefined
        );
      } catch (err) { console.warn('[DEX] queueIDReorder (proxy cancel) failed:', err); }
      // v27.10: ILS-sichere Counter-Pflege statt blindem Voll-Sync (die
      // Assistenz ist ein normaler User — siehe Kommentar im Self-Cancel).
      try {
        const isSplit = typeof event.durchstarterCapacity === 'number'
          && typeof event.funstarterCapacity === 'number'
          && ((event.durchstarterCapacity || 0) > 0 || (event.funstarterCapacity || 0) > 0);
        await eventService.releaseSeatAfterCancel(subsiteUrl, {
          isSplit,
          previousStatus: registration.Status || '',
          starterType: registration.StarterType || undefined,
          waitlistDisabled: event.waitlistEnabled === false, // v27.11
        });
      } catch { /* best-effort */ }
    }
    await loadEvents();
    return true;
  }

  /**
   * v24.36: Custom-Field-Antworten einer Fremd-Anmeldung aktualisieren
   * (Assistenz). Operiert direkt auf der übergebenen Registrierungs-Zeile.
   */
  async function updateProxyRegistration(eventId: string, registration: SPRegistration, customData: Record<string, string>): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl || !registration?.Id) return false;
    const event = events.find(e => e.id === eventId);
    const fieldMap: Record<string, string> = {};
    const fieldLabelMap: Record<string, string> = {};
    if (event) {
      for (const f of event.eventSpecificFields) {
        if (f.spInternalName) fieldMap[f.id] = f.spInternalName;
        fieldLabelMap[f.id] = f.label;
      }
    }
    let oldCustomData: Record<string, string> = {};
    try { if (registration.CustomData) oldCustomData = JSON.parse(registration.CustomData); }
    catch { /* */ }
    const success = await eventService.updateRegistrationData(
      subsiteUrl, registration.Id, customData, fieldMap, oldCustomData, fieldLabelMap
    );
    if (success) await loadEvents();
    return success;
  }

  // v24.43: Eine (stellvertretend angelegte) Anmeldung KOMPLETT an die
  // angemeldete Person selbst übergeben — Owner + Zeilen-Autor + RegisteredBy
  // werden auf die Person gesetzt. Danach sieht/verwaltet NUR noch sie die
  // Anmeldung (über „Meine Events"); sie verschwindet aus der „Assistenz"-
  // Kachel des bisherigen Akteurs. Etwaiger Assistenz-Verknüpfungs-Eintrag
  // wird auf 'Cancelled' gesetzt (kein Assistenz-Bezug mehr).
  async function handBackToParticipant(eventId: string, registration: SPRegistration): Promise<boolean> {
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl || !registration?.Id) return false;
    const pEmail = (registration.ParticipantEmail || '').trim();
    if (!pEmail) return false;
    const pName = `${registration.Vorname || ''} ${registration.Nachname || ''}`.trim() || registration.ParticipantName || pEmail;
    const ok = await eventService.assignRegistrationToAssistant(subsiteUrl, registration.Id, pEmail, pName);
    try { await eventService.setAssistantLinkStatusForRegistration(registration.Id, subsiteUrl, 'Cancelled'); } catch { /* */ }
    if (ok) {
      eventService.writeChangeLog({
        action: 'ParticipantUpdated',
        targetType: 'Participant',
        targetId: pEmail,
        targetName: pName,
        eventId,
        eventTitle: events.find(e => e.id === eventId)?.title || '',
        details: { scope: 'handedBackToParticipant', actorEmail: currentUserEmail },
      }).catch(() => { /* */ });
    }
    await loadEvents();
    return ok;
  }

  /**
   * v11.83: Nach einem Team-Mitglied-Cancel (Self-Cancel) erledigt diese
   * Routine:
   *   1) Verbleibende aktive Team-Mitglieder laden (ohne den gerade
   *      Abgemeldeten, der jetzt 'Abgemeldet' ist).
   *   2) Falls die abgemeldete Person Lead war UND mindestens ein Member
   *      uebrig ist, das früheste aktive Mitglied per MERGE-Patch zum
   *      neuen Lead promoten.
   *   3) Pro verbleibendem Mitglied eine Info-Mail in DEX_Emails queuen,
   *      die den Cancel ankündigt und die nächsten Schritte erklärt.
   *
   * Fail-safe: alle Sub-Operationen sind best-effort und schlucken Fehler
   * still — das Cancel selbst hat oben schon erfolgreich auf dem Item
   * geschrieben, ein Mail-/Promote-Fehler darf den User-Flow nicht
   * blockieren.
   */
  async function handleTeamCancelPostStep(
    event: DeloitteEvent,
    eventId: string,
    subsiteUrl: string,
    teamId: string,
    teamName: string,
    wasTeamLead: boolean,
    cancelledReg: SPRegistration
  ): Promise<void> {
    const members = await eventService.getTeamMembers(subsiteUrl, teamId);
    // Verbleibende = aktive (NICHT 'Abgemeldet') und NICHT der gerade
    // abgemeldete Eintrag (Id-Vergleich, weil ein parallel-Member denselben
    // Vor-/Nachnamen haben könnte).
    const remaining = members.filter(m => m.Status !== 'Abgemeldet' && m.Id !== cancelledReg.Id);
    if (remaining.length === 0) {
      // Team aufgelöst — kein Promote, keine Info-Mails nötig.
      return;
    }

    // Auto-Promote: wenn der Cancel ein Lead war, das früheste aktive
    // Member zum neuen Lead machen. Sortier-Kriterium: kleinste
    // TeilnehmerID, sonst früheste RegistrationDate, sonst kleinste Id.
    let newLeadId: number | null = null;
    if (wasTeamLead) {
      const sorted = [...remaining].sort((a, b) => {
        const aTid = typeof a.TeilnehmerID === 'number' ? a.TeilnehmerID : Number.MAX_SAFE_INTEGER;
        const bTid = typeof b.TeilnehmerID === 'number' ? b.TeilnehmerID : Number.MAX_SAFE_INTEGER;
        if (aTid !== bTid) return aTid - bTid;
        const aRd = new Date(a.RegistrationDate || 0).getTime();
        const bRd = new Date(b.RegistrationDate || 0).getTime();
        if (aRd !== bRd) return aRd - bRd;
        return a.Id - b.Id;
      });
      const promoteTarget = sorted[0];
      if (promoteTarget) {
        try {
          await eventService.promoteToTeamLead(subsiteUrl, promoteTarget.Id);
          newLeadId = promoteTarget.Id;
        } catch (err) {
          console.warn('[DEX] promoteToTeamLead failed:', err);
        }
      }
    }

    // v12.14: Info-Mails kommen aus TemplateType=TeamMemberCancelled.
    // {{NewLeadBlock}}-Platzhalter wird für den Auto-Promote-Empfänger
    // gefüllt, für alle anderen leer.
    if (event.disableEmails) return;
    const lang = event.emailLanguage || 'EN';
    const isDe = lang.toUpperCase() === 'DE';
    const tpl = await eventService.getEmailTemplate('TeamMemberCancelled', lang).catch(() => null);
    const teamSizeCfg = event.teamSize || (remaining.length + 1);
    const cancelledFullName = `${cancelledReg.Vorname || ''} ${cancelledReg.Nachname || ''}`.trim() || cancelledReg.ParticipantEmail;
    const teamNameStr = teamName ? `„${teamName}"` : '';
    const newLeadBlockHtml = isDe
      ? `<p style="padding:10px 14px;background:rgba(134,188,37,0.10);border:1px solid #86bc25;border-radius:6px;"><strong>Du bist jetzt der neue Team-Lead.</strong> Du kannst über „Meine Events" eine neue Person hinzufügen, falls der frei gewordene Platz wieder gefüllt werden soll.</p>`
      : `<p style="padding:10px 14px;background:rgba(134,188,37,0.10);border:1px solid #86bc25;border-radius:6px;"><strong>You are the new team lead now.</strong> You can add a replacement member via „My Events" if you want to fill the freed slot.</p>`;

    for (const m of remaining) {
      const mFirst = m.Vorname || (m.ParticipantName || '').split(/[ ,]+/)[0] || '';
      const mFull = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantEmail;
      const isNewLead = newLeadId !== null && m.Id === newLeadId;
      const vars: Record<string, string> = {
        Name: mFirst || mFull,
        CancelledName: cancelledFullName,
        TeamName: teamNameStr,
        EventTitle: event.title,
        ActiveCount: String(remaining.length),
        TeamSize: String(teamSizeCfg),
        NewLeadBlock: isNewLead ? newLeadBlockHtml : '',
        AppUrl: `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      };
      let mail: { subject: string; body: string };
      if (tpl) {
        mail = buildEmailFromTemplate(tpl, vars);
      } else {
        const inner = isDe
          ? `<p>Hallo ${mFirst},</p><p>${cancelledFullName} hat sich vom Team ${teamNameStr} abgemeldet (${remaining.length}/${teamSizeCfg}).</p>${isNewLead ? newLeadBlockHtml : ''}`
          : `<p>Hello ${mFirst},</p><p>${cancelledFullName} cancelled their registration from team ${teamNameStr} (${remaining.length}/${teamSizeCfg}).</p>${isNewLead ? newLeadBlockHtml : ''}`;
        mail = {
          subject: isDe ? `Team-Update — ${event.title}` : `Team update — ${event.title}`,
          body: wrapTemplate('#ed8b00', isDe ? 'Team-Update' : 'Team update', `Event ${event.title}`, inner),
        };
      }
      try {
        await eventService.queueEmail(
          mail.subject,
          m.ParticipantEmail,
          mFull,
          mail.body,
          'TeamMemberCancelled',
          event.title,
          eventId
        );
      } catch (err) {
        console.warn('[DEX] team-cancel info mail failed:', err);
      }
    }
  }

  // v18.11: Proaktive Absage durch den eingeloggten User („Ich nehme nicht
  // teil"). Wenn schon eine aktive/Warteliste-Anmeldung existiert, wird sie
  // regulär abgemeldet (Seat-Sync, Mail, IDReorder laufen mit). Sonst wird
  // eine reine Absage-Zeile (Status=Abgemeldet, Marker _declined) angelegt.
  async function declineEvent(eventId: string): Promise<boolean> {
    if (isDemoShowcaseId(eventId)) return true;
    const subsiteUrl = subsiteMap.current[eventId];
    if (!subsiteUrl) return false;
    // Name aus displayName ableiten („Nachname, Vorname" oder „Vorname Nachname").
    // v22.57: Claims-Login-Token (z.B. „0#.f|membership|user@…") niemals als
    // Namen verwenden — declineRegistration zieht in dem Fall den sauberen
    // Namen aus dem Benutzerprofil.
    const looksLikeClaim = (s: string): boolean => /\|membership\||0#\.f\||^i:0#/i.test((s || '').trim());
    const dn = looksLikeClaim(currentUserName) ? '' : (currentUserName || '').trim();
    let firstName = ''; let lastName = '';
    if (dn.indexOf(',') >= 0) {
      const p = dn.split(',').map(s => s.trim());
      lastName = p[0] || ''; firstName = p[1] || '';
    } else {
      const p = dn.split(/\s+/).filter(Boolean);
      firstName = p[0] || ''; lastName = p.slice(1).join(' ');
    }
    // v30.67 (Review): Ein 429 hier hieß „keine Zeile" — und darunter wurde
    // eine NEUE Absage-Zeile geschrieben, während die aktive stehen blieb:
    // Platz belegt, QR-Mail und Termin laufen weiter, Person doppelt im
    // Organizer Center, und sie selbst las „Absage erfasst". Nicht lesbar =
    // nichts schreiben; handleDecline zeigt bei false die Fehlermeldung.
    let readFailed = false;
    const existing = await eventService.getMyRegistration(subsiteUrl, currentUserEmail, () => { readFailed = true; });
    if (readFailed) {
      console.warn('[DEX] declineEvent: eigene Zeile nicht lesbar — keine Absage-Zeile geschrieben', eventId);
      return false;
    }
    if (existing) {
      // Bereits abgemeldet/abgesagt → nichts zu tun. Aktiv/Warteliste →
      // regulärer Cancel-Pfad (gibt Sitzplatz frei, Mail, IDReorder).
      if (existing.Status === 'Abgemeldet') return true;
      return await cancelRegistration(eventId);
    }
    const ok = await eventService.declineRegistration(
      subsiteUrl, firstName, lastName, currentUserEmail, currentUserName, currentUserEmail
    );
    if (ok) {
      const event = events.find(e => e.id === eventId);
      eventService.writeChangeLog({
        action: 'ParticipantDeclined',
        targetType: 'Participant',
        targetId: currentUserEmail,
        targetName: `${firstName} ${lastName}`.trim() || currentUserEmail,
        eventId,
        eventTitle: event?.title || '',
        details: { asActor: 'self', proactiveDecline: true },
      }).catch(() => { /* */ });
    }
    return ok;
  }

  return { cancelRegistration, cancelTeamMember, getMyProxyRegistrations, cancelProxyRegistration, updateProxyRegistration, handBackToParticipant, declineEvent };
}
