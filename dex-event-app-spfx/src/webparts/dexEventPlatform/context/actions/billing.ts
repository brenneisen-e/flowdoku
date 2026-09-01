/**
 * F&A-Abrechnung: Versand, Protokoll, Snapshots, Abschluss, Automatik-Mails.
 *
 * v30.66: Aus `EventContext.tsx` herausgezogen (Modularisierung Stufe 3).
 * Die Funktionskoerper sind unveraendert - statt aus der Provider-Closure
 * beziehen sie ihre Umgebung aus dem `deps`-Objekt.
 */

import * as React from 'react';
import { DeloitteEvent } from '../../types';
import { EventService, SPRegistration } from '../../services/EventService';
import { buildHashDeepLink } from '../../utils/deepLink';
import { wrapTemplate } from '../../services/EmailTemplates';
import { BundledItem, bundledCommOf, bundledItemsTableHtml, bundledItemsHeading } from '../../utils/bundledComm';
import { parseBillingOf, missingBillingFields, renderBillingInfoMailBody, renderBillingListMailBody, trimBillingLog, faRowsFromRegistrations, BillingData, BillingLogEntry, FAConfig } from '../../utils/faBilling';

export interface BillingDeps {
  eventService: EventService;
  events: DeloitteEvent[];
  setEvents: React.Dispatch<React.SetStateAction<DeloitteEvent[]>>;
  currentUserEmail: string;
  currentUserName: string;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
}

export function makeBillingActions(deps: BillingDeps) {
  const { eventService, events, setEvents, currentUserEmail, currentUserName, getAllRegistrations } = deps;

  // ==================== F&A-Abrechnung (v30.5) ====================
  // Fachkonzept „Abrechnungsrelevante Events und F&A Integration": Versand
  // und Abschluss laufen über das Piggyback `_billing` des Hauptevents —
  // geschrieben per patchEventOverridesValue (read-modify-write NUR dieses
  // Felds), damit kein voller updateEvent (und kein 28-MB-Reload) nötig ist.

  /** `_billing` nach einem Patch auch im lokalen State nachziehen — bewusst
   *  KEIN loadEvents (v29.77-Lehre: der Voll-Reload war die 429-Ursache). */
  function applyBillingLocally(eventId: string, newB: BillingData): void {
    setEvents(prev => prev.map(e => {
      if (e.id !== eventId) return e;
      let o: Record<string, unknown> = {};
      try { o = JSON.parse(e.emailTemplateOverrides || '{}'); } catch { o = {}; }
      o._billing = newB;
      return { ...e, emailTemplateOverrides: JSON.stringify(o) };
    }));
  }

  async function getFAConfig(): Promise<FAConfig> {
    return eventService.getFAConfig();
  }

  async function saveFAConfig(cfg: FAConfig): Promise<boolean> {
    return eventService.saveFAConfig(cfg);
  }

  /**
   * Abrechnungsinfos ('info') bzw. Teilnehmerliste ('list') an die im F&A
   * Center gepflegten Verteiler senden. Organizer stehen immer in CC.
   * Protokolliert in `_billing.log` inkl. vollständigem Mail-Inhalt und
   * legt den übermittelten Stand als Snapshot ab (die F&A-Detailansicht
   * zeigt laut Konzept NUR Übermitteltes, nie den Live-Stand).
   */
  async function sendFAMail(ev: DeloitteEvent, kind: 'info' | 'list', opts?: { auto?: boolean }): Promise<{ ok: boolean; reason?: string }> {
    const b = parseBillingOf(ev);
    if (!b || !b.relevant) return { ok: false, reason: 'not-relevant' };
    const idNum = parseInt(ev.id, 10);
    if (!isFinite(idNum)) return { ok: false, reason: 'bad-id' };
    const cfg = await eventService.getFAConfig();
    const recipients = (kind === 'info' ? cfg.infoRecipients : cfg.listRecipients).filter(Boolean);
    if (recipients.length === 0) return { ok: false, reason: 'no-recipients' };
    const by = opts?.auto ? 'System (Auto-Versand)' : (currentUserName || currentUserEmail);
    const cc = (ev.organizerEmails || []).filter(Boolean).join('; ');
    const nowIso = new Date().toISOString();
    let body = '';
    let subject = '';
    let stampPatch: Partial<BillingData> = {};
    // v30.24: Deep-Link ins F&A Center statt Datei-Anhang — der Tenant
    // blockt Anhänge aus Power Automate komplett (NDR, s. v26.71). Der Link
    // öffnet das Event im F&A Center; dort liegt der versendete Stand und
    // lässt sich als Excel ziehen. Die personenbezogene Liste bleibt damit
    // in DEX hinter der F&A-Rolle statt in Postfächern zu kursieren.
    const faCenterUrl = buildHashDeepLink(
      `${eventService.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      { action: 'fa', event: ev.id }
    );
    if (kind === 'info') {
      if (missingBillingFields(b).length > 0) return { ok: false, reason: 'incomplete' };
      body = renderBillingInfoMailBody(ev, b, by, faCenterUrl);
      subject = `[DEX] Abrechnungsinformationen: ${ev.title}`;
      stampPatch = { infoSentAt: nowIso, infoSnapshot: { ...(b.fields || {}) }, ...(opts?.auto ? { autoInfoSentAt: nowIso } : {}) };
    } else {
      const regs = await getAllRegistrations(ev.id);
      // v30.50: EINE Abbildung für Versand, Snapshot und Download — vorher
      // hatte jeder Weg seine eigene `.map()`.
      const rows = faRowsFromRegistrations(regs);
      body = renderBillingListMailBody(ev, rows, by, faCenterUrl, b);
      subject = `[DEX] Teilnehmerliste: ${ev.title}`;
      stampPatch = { listSentAt: nowIso, listSnapshot: rows.slice(0, 500), ...(opts?.auto ? { autoListSentAt: nowIso } : {}) };
    }
    const emailType = kind === 'info' ? (opts?.auto ? 'FA_INFO_AUTO' : 'FA_INFO') : (opts?.auto ? 'FA_LIST_AUTO' : 'FA_LIST');
    // Auto-Pfad: serverseitiger Doppelversand-Schutz — mehrere Admins können
    // die App gleichzeitig booten, der localStorage-Stempel reicht dann nicht.
    if (opts?.auto && await eventService.hasQueuedEmail(emailType, ev.id)) {
      return { ok: false, reason: 'already-queued' };
    }
    const ok = await eventService.queueEmail(subject, recipients.join('; '), 'Finance & Accounting', body, emailType, ev.title, ev.id, cc || undefined);
    if (!ok) return { ok: false, reason: 'queue-failed' };
    const entry: BillingLogEntry = {
      ts: nowIso, by,
      action: kind === 'info' ? 'Abrechnungsinformationen an F&A versendet' : 'Teilnehmerliste an F&A versendet',
      mailType: kind, to: recipients.join('; '), cc, subject, body,
    };
    const newB: BillingData = { ...b, ...stampPatch, log: trimBillingLog([...(b.log || []), entry]) };
    const patched = await eventService.patchEventOverridesValue(idNum, '_billing', newB);
    if (patched) applyBillingLocally(ev.id, newB);
    return { ok: true };
  }

  /**
   * v30.53: Rückfrage an F&A in der Kommunikationshistorie festhalten.
   *
   * Das Fachkonzept verlangt für Bereich 3 „Speicherung der Kommunikation in
   * der Kommunikationshistorie". Die Rückfrage entsteht aber im Outlook des
   * Organizers (mailto:) — DEX sieht den Text NIE und bekommt auch keine
   * Bestätigung, dass die Mail rausging. Protokolliert wird deshalb genau
   * das, was belegbar ist: dass und wann eine Rückfrage an wen begonnen
   * wurde. Ein Eintrag „Mail gesendet" wäre eine Behauptung über etwas, das
   * die App weder auslöst noch prüfen kann — und in einer revisionssicheren
   * Historie ist eine solche Behauptung schlimmer als eine Lücke.
   */
  async function logFAContact(ev: DeloitteEvent, to: string, subject: string): Promise<boolean> {
    const b = parseBillingOf(ev);
    if (!b) return false;
    const idNum = parseInt(ev.id, 10);
    if (!isFinite(idNum)) return false;
    const nowIso = new Date().toISOString();
    const entry: BillingLogEntry = {
      ts: nowIso,
      by: currentUserName || currentUserEmail,
      action: 'Rückfrage an F&A im eigenen Postfach geöffnet',
      to, subject,
    };
    const newB: BillingData = { ...b, log: trimBillingLog([...(b.log || []), entry]) };
    const ok = await eventService.patchEventOverridesValue(idNum, '_billing', newB);
    if (ok) applyBillingLocally(ev.id, newB);
    return ok;
  }

  /** „Als abgerechnet markieren" — nur F&A/Admin (UI-seitig gegated). Der
   *  Status bleibt laut Konzept dauerhaft bestehen; Zeitpunkt + Person
   *  werden protokolliert. */
  /**
   * v30.61: Aktualisierte Sammelmail nach einer späteren Änderung.
   *
   * Nutzer-Entscheidung 01.09.2026: Wer später einen Termin dazubucht oder
   * abwählt, soll EINE Mail mit der neuen vollständigen Liste bekommen — nicht
   * gar nichts (dann stimmt die alte Mail nicht mehr) und nicht eine Mail je
   * Änderung (dann ist die Bündelung umsonst).
   *
   * Bewusst eine eigene, schlanke Funktion und kein zweiter Durchlauf durch
   * `registerForEvent`: Dort hängen Kapazitätsprüfung, Wartelisten-Logik,
   * TeilnehmerID-Vergabe und Outlook mit dran. Hier ist nichts zu registrieren
   * — es hat sich nur die Liste geändert, die in der Mail steht.
   *
   * Die Mail geht NUR raus, wenn wirklich noch etwas gebucht ist. Wer alles
   * abgewählt hat, ist abgemeldet; dafür gibt es die Abmelde-Bestätigung, und
   * eine „Deine Termine"-Mail mit leerer Liste wäre eine Verhöhnung.
   */
  async function sendBundledUpdateMail(
    parentEvent: DeloitteEvent,
    recipientEmail: string,
    recipientName: string,
    items: BundledItem[]
  ): Promise<boolean> {
    if (!parentEvent || !recipientEmail) return false;
    if (!bundledCommOf(parentEvent).mail) return false;
    if (!items || items.length === 0) return false;
    if (parentEvent.disableEmails || parentEvent.disableRegistrationEmail) return false;
    const isDeMail = (parentEvent.emailLanguage || 'EN').toUpperCase() === 'DE';
    const first = (recipientName || '').trim().split(/\s+/)[0] || recipientName || '';
    const subject = isDeMail
      ? `Deine Anmeldung wurde aktualisiert — ${parentEvent.title}`
      : `Your registration was updated — ${parentEvent.title}`;
    const intro = isDeMail
      ? `<p>Hallo ${first},</p><p>deine Anmeldung für <strong>${parentEvent.title}</strong> hat sich geändert. Hier ist der aktuelle Stand — diese Übersicht ersetzt die vorherige Bestätigung.</p>`
      : `<p>Hi ${first},</p><p>your registration for <strong>${parentEvent.title}</strong> has changed. Here is the current state — this overview replaces the previous confirmation.</p>`;
    const body = intro
      + `<p style="margin:18px 0 0;font-weight:700;">${bundledItemsHeading(items.length, isDeMail, parentEvent.childEventTermPlural)}</p>`
      + bundledItemsTableHtml(items, isDeMail)
      + (isDeMail
        ? `<p style="margin-top:24px;"><strong>Viele Gr&uuml;&szlig;e</strong><br><br><strong>Dein Event-Team</strong></p>`
        : `<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`);
    // v30.66: Die vier Argumente standen in der falschen Reihenfolge. Die
    // Signatur ist wrapTemplate(headingColor, heading, subheading, bodyHtml) —
    // uebergeben wurde (Ueberschrift, Titel, body, '#0076a8'). Damit landete der
    // gesamte Mailtext im Subheading und als Body stand woertlich '#0076a8' in
    // der Mail. Alle anderen 7 Aufrufstellen im Projekt uebergeben die Farbe
    // zuerst; der Fehler kam mit der gebuendelten Update-Mail in v30.61 herein.
    const wrapped = wrapTemplate(
      '#0076a8',
      isDeMail ? 'Deine Anmeldung wurde aktualisiert' : 'Your registration was updated',
      parentEvent.title, body
    );
    try {
      return await eventService.queueEmail(
        subject, recipientEmail, recipientName, wrapped,
        'RegistrationUpdate', parentEvent.title, parentEvent.id
      );
    } catch (err) {
      console.warn('[DEX] sendBundledUpdateMail failed:', err);
      return false;
    }
  }

  async function markEventSettled(ev: DeloitteEvent): Promise<boolean> {
    const b = parseBillingOf(ev);
    if (!b || b.settled) return false;
    const idNum = parseInt(ev.id, 10);
    if (!isFinite(idNum)) return false;
    const by = currentUserName || currentUserEmail;
    const nowIso = new Date().toISOString();
    const newB: BillingData = {
      ...b,
      settled: { ts: nowIso, by },
      log: trimBillingLog([...(b.log || []), { ts: nowIso, by, action: 'Event als abgerechnet markiert' }]),
    };
    const ok = await eventService.patchEventOverridesValue(idNum, '_billing', newB);
    if (ok) applyBillingLocally(ev.id, newB);
    return ok;
  }

  /**
   * v30.60: Von F&A nachgetragene Personalnummern/Kostenstellen speichern.
   *
   * Nutzer-Ansage 01.09.2026: F&A hat Zugriff auf die Backoffice-Liste
   * „Active Employees" und mappt von dort die Personalnummer je Teilnehmer.
   * Geschrieben wird in den SNAPSHOT (`_billing.listSnapshot`) und nicht in
   * die Teilnehmerliste auf der Subsite — aus zwei Gründen: (1) F&A hat auf
   * die Subsites keinen Zugriff, (2) der Snapshot ist der Stand, der
   * gemeldet wurde, und genau der geht als Excel wieder hinaus.
   *
   * Zusammengeführt wird über die E-Mail-Adresse, nicht über den Index: Der
   * Snapshot kann zwischen Anzeige und Speichern neu versendet worden sein,
   * und ein Index-Treffer würde die Nummer dann der falschen Person
   * zuordnen. Personen, die im aktuellen Snapshot nicht mehr stehen,
   * verlieren ihren Eintrag — das ist gewollt, sie sind nicht mehr gemeldet.
   */
  async function saveFAPersonalNumbers(
    ev: DeloitteEvent,
    values: Record<string, { personalNr?: string; costCenter?: string }>
  ): Promise<boolean> {
    const b = parseBillingOf(ev);
    if (!b || !b.listSnapshot) return false;
    const idNum = parseInt(ev.id, 10);
    if (!isFinite(idNum)) return false;
    const key = (s: string): string => (s || '').toLowerCase().trim();
    let changed = 0;
    const snap = b.listSnapshot.map(r => {
      const v = values[key(r.email)];
      if (!v) return r;
      const pn = (v.personalNr || '').trim();
      const cc = (v.costCenter || '').trim();
      if (pn === (r.personalNr || '') && cc === (r.costCenter || '')) return r;
      changed++;
      return { ...r, personalNr: pn, costCenter: cc };
    });
    if (changed === 0) return true;
    const by = currentUserName || currentUserEmail;
    const nowIso = new Date().toISOString();
    const newB: BillingData = {
      ...b,
      listSnapshot: snap,
      log: trimBillingLog([...(b.log || []), {
        ts: nowIso, by,
        action: `Personalnummern ergänzt (${changed} ${changed === 1 ? 'Person' : 'Personen'})`,
      }]),
    };
    const ok = await eventService.patchEventOverridesValue(idNum, '_billing', newB);
    if (ok) applyBillingLocally(ev.id, newB);
    return ok;
  }

  /**
   * Auto-Versand (Fachkonzept Abschnitt 12): Abrechnungsinfos 7 Kalendertage
   * VOR dem Event (bzw. nach Aktivierung, wenn später erstellt — Entwürfe
   * werden übersprungen), Teilnehmerliste 7 Kalendertage NACH dem Event.
   * Bei unvollständigen Infos geht statt der F&A-Mail EINMAL eine
   * Erinnerung an die Organizer. Läuft beim Admin-Boot (gleicher Ort wie
   * der Wochenbericht); Doppelversand-Schutz über die _billing-Stempel plus
   * die DEX_Emails-Queue. Der 30-Tage-Rückschau-Deckel verhindert, dass
   * Altbestände beim Roll-out plötzlich Mails auslösen.
   */
  async function maybeSendBillingAutoMails(): Promise<{ infoSent: number; listSent: number; reminders: number }> {
    const out = { infoSent: 0, listSent: 0, reminders: 0 };
    const day = 86400000;
    const now = Date.now();
    try {
      for (const ev of events) {
        if (ev.parentEventId || ev.isDemoShowcase || ev.isFictive) continue;
        const b = parseBillingOf(ev);
        if (!b || !b.relevant || b.sendMode !== 'auto' || b.settled) continue;
        const start = new Date(ev.startDate || '').getTime();
        const end = new Date(ev.endDate || ev.startDate || '').getTime();
        if (!isFinite(start) || !isFinite(end)) continue;
        if (end < now - 30 * day) continue;
        // Abrechnungsinfos — ab 7 Tage vor Start, solange noch nichts ging.
        if (now >= start - 7 * day && !b.autoInfoSentAt && !b.infoSentAt) {
          if (missingBillingFields(b).length > 0) {
            if (!b.autoInfoReminderAt) {
              const orgTo = (ev.organizerEmails || []).filter(Boolean).join('; ');
              if (orgTo && !(await eventService.hasQueuedEmail('FA_INFO_REMINDER', ev.id))) {
                const remBody = `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333;line-height:1.5;"><p>Hallo,</p><p>für das abrechnungsrelevante Event <strong>${ev.title}</strong> steht der automatische Versand der Abrechnungsinformationen an Finance &amp; Accounting an — die Pflichtangaben sind aber noch <strong>unvollständig</strong>. Bitte ergänze sie im Event-Wizard (Schritt „Abrechnung"); der Versand wird dann beim nächsten App-Start nachgeholt.</p></div>`;
                const qok = await eventService.queueEmail(`[DEX] Abrechnungsinfos unvollständig: ${ev.title}`, orgTo, '', remBody, 'FA_INFO_REMINDER', ev.title, ev.id);
                if (qok) {
                  const nowIso = new Date().toISOString();
                  const newB: BillingData = { ...b, autoInfoReminderAt: nowIso, log: trimBillingLog([...(b.log || []), { ts: nowIso, by: 'System (Auto-Versand)', action: 'Erinnerung an Organizer: Abrechnungsinfos unvollständig' }]) };
                  const idNum = parseInt(ev.id, 10);
                  if (isFinite(idNum) && await eventService.patchEventOverridesValue(idNum, '_billing', newB)) applyBillingLocally(ev.id, newB);
                  out.reminders++;
                }
              }
            }
          } else {
            const r = await sendFAMail(ev, 'info', { auto: true });
            if (r.ok) out.infoSent++;
          }
          // Bewusst NICHT im selben Lauf auch noch die Liste senden: beide
          // Zweige patchen `_billing`, und der zweite würde auf dem hier
          // veralteten Objekt arbeiten und den Info-Eintrag überschreiben.
          // Die Liste geht dann beim nächsten App-Start raus.
          continue;
        }
        // Teilnehmerliste — ab 7 Tage nach Ende.
        if (now >= end + 7 * day && !b.autoListSentAt && !b.listSentAt) {
          const r = await sendFAMail(ev, 'list', { auto: true });
          if (r.ok) out.listSent++;
        }
      }
    } catch (err) {
      console.warn('[DEX] maybeSendBillingAutoMails error:', err);
    }
    return out;
  }

  return { getFAConfig, saveFAConfig, sendFAMail, logFAContact, sendBundledUpdateMail, markEventSettled, saveFAPersonalNumbers, maybeSendBillingAutoMails };
}
