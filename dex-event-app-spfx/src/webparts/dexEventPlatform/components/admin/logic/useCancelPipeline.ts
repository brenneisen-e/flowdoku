/* useCancelPipeline — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 303-543 des
 * Stands vor dem Schnitt). Der Rumpf ist zeichengleich uebernommen; was die
 * Gruppe aus dem Komponenten-Scope liest, kommt als `ctx` herein, was sie
 * nach aussen liefert, geht als Objekt zurueck.
 */
import * as React from 'react';
import { DUP_ACTIVE_STATI } from '../../admin/adminConstants';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';
import { applyEventTemplateOverride, formatOrganizerList } from '../../../context/EventContext';
import { buildEmailFromTemplate, cancellationEmail, promotionEmail } from '../../../services/EmailTemplates';
import { invalidateInactiveAccountCache } from '../../../utils/accountCheckCache';
import { isEventOver } from '../../../utils/eventFormat';
import { withParentTitleSubject } from '../../../utils/mailSubject';
import { AdminToastState } from '../../admin/adminTypes';

export interface UseCancelPipelineCtx {
  allEvents: DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  currentUser: import("../../../types/index").User;
  duplicateEmails: Set<string>;
  eventServiceRef: EventService;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  isDe: boolean;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setAdminToast: React.Dispatch<React.SetStateAction<AdminToastState>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export interface UseCancelPipelineResult {
  buildCancellationMail: (ev: DeloitteEvent, reg: SPRegistration, fullName: string) => Promise<{    subject: string;    body: string;}>;
  cleanupShadowDuplicates: () => Promise<void>;
  isSyncingRegistry: boolean;
  performSilentDuplicateDelete: (reg: SPRegistration) => Promise<boolean>;
  performStandardCancel: (reg: SPRegistration) => Promise<void>;
  setIsSyncingRegistry: React.Dispatch<React.SetStateAction<boolean>>;
  setSyncRegistryResult: React.Dispatch<React.SetStateAction<string>>;
  shadowDupBusy: boolean;
  syncRegistryResult: string;
}

export function useCancelPipeline(ctx: UseCancelPipelineCtx): UseCancelPipelineResult {
  const {
    allEvents, confirmDialog, currentUser, duplicateEmails, eventServiceRef,
    isDe, registrations, reloadRegistrations, selectedEvent, setAdminToast, showAlert,
  } = ctx;
  /**
   * v29.44: Abmelde-Mail bauen — für ALLE Organizer-Wege gleich.
   *
   * Vorher nahm das Organizer Center überall `cancellationEmail(...)`, also den
   * fest eingebauten Standardtext. Der Selbst-Abmelde-Weg des Teilnehmers löst
   * dagegen seit jeher die gepflegte Vorlage auf (SharePoint-Template +
   * Event-Override). Dieselbe Abmeldung sah damit unterschiedlich aus, je
   * nachdem, WER sie ausgelöst hat — und ein Organizer, der den Text seines
   * Sub-Events angepasst hatte, bekam ihn nie zu sehen.
   */
  const buildCancellationMail = async (
    ev: DeloitteEvent,
    reg: SPRegistration,
    fullName: string,
  ): Promise<{ subject: string; body: string }> => {
    const fallback = cancellationEmail(fullName, ev.title);
    if (!eventServiceRef) return fallback;
    try {
      const lang = ev.emailLanguage || 'EN';
      const spTplRaw = await eventServiceRef.getEmailTemplate('Abmeldung', lang).catch(() => null);
      const spTpl = applyEventTemplateOverride(spTplRaw, ev.emailTemplateOverrides, 'Abmeldung');
      if (!spTpl) return fallback;
      return buildEmailFromTemplate(spTpl, {
        Name: (reg.Vorname || '').trim() || fullName,
        EventTitle: ev.title,
        AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
      });
    } catch { return fallback; }
  };
  const performStandardCancel = async (reg: SPRegistration): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
    const eventWasOver = isEventOver(selectedEvent);
    setAdminToast({ kind: 'cancelling', name });
    const cancelledStarterType = reg.StarterType || '';
    await eventServiceRef.cancelRegistration(selectedEvent.subsiteUrl, reg.Id, `${currentUser.firstName} ${currentUser.surname}`.trim(), currentUser.email);
    // v29.31: Die „Konto inaktiv"-Prüfung ist 24 h gecacht. Ohne Verwerfen
    // meldete die Sammel-Box auf der Startseite die eben abgemeldete Person
    // bis zum nächsten Tag weiter als offenen Fall.
    invalidateInactiveAccountCache([selectedEvent.id, selectedEvent.parentEventId || '']);
    if (reg.ParticipantEmail && !eventWasOver) {
      if (!selectedEvent.disableEmails && !selectedEvent.disableCancellationEmail) {
        const emailData = await buildCancellationMail(selectedEvent, reg, name);
        eventServiceRef.queueEmail(
          emailData.subject, reg.ParticipantEmail, name, emailData.body,
          'Abmeldung', selectedEvent.title, selectedEvent.id
        ).catch(err => console.warn('[DEX]', err));
      }
      if (!selectedEvent.disableOutlook) {
        eventServiceRef.queueOutlookEvent(
          reg.ParticipantEmail, selectedEvent.id, selectedEvent.title, 'Ausladen'
        ).catch(err => console.warn('[DEX]', err));
      }
    }
    if (reg.ParticipantEmail && selectedEvent.eventNumber) {
      eventServiceRef.removeParticipantEvent(
        reg.ParticipantEmail, selectedEvent.eventNumber
      ).catch(err => console.warn('[DEX]', err));
    }
    const isSplitEvent = typeof selectedEvent.durchstarterCapacity === 'number'
      && typeof selectedEvent.funstarterCapacity === 'number'
      && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
    const useTypeFilter = isSplitEvent && !selectedEvent.splitSharedWaitlist;
    // v27.11: WaitlistEnabled=false ist jetzt ein echter Kill-Switch — kein
    // automatisches Nachrücken mehr, wenn der Organizer die Warteliste
    // abgeschaltet hat (manuelles Nachrücken über den Admin-Button bleibt
    // als bewusster Override möglich).
    if (!eventWasOver && selectedEvent.waitlistEnabled !== false) {
      try {
        const promoted = await eventServiceRef.promoteFirstWaitlistItem(
          selectedEvent.subsiteUrl,
          cancelledStarterType || undefined,
          selectedEvent.maxParticipants,
          (useTypeFilter && cancelledStarterType) ? cancelledStarterType : undefined,
          { itemId: reg.Id, participantEmail: reg.ParticipantEmail || '' },
        );
        if (promoted && promoted.success && promoted.email) {
          setAdminToast({ kind: 'promoted', name: promoted.name || promoted.email, email: promoted.email, type: cancelledStarterType || undefined });
          if (!selectedEvent.disableEmails) {
            try {
              const lang = selectedEvent.emailLanguage || 'EN';
              const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
              const promoteVars = {
                Name: promotedFirstName,
                EventTitle: selectedEvent.title,
                Organizer: formatOrganizerList(selectedEvent.organizers, lang),
                AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
                WaitlistPosition: '',
              };
              let emailData: { subject: string; body: string };
              const spTplRaw = await eventServiceRef.getEmailTemplate('Nachruecken', lang).catch(() => null);
              const spTpl = applyEventTemplateOverride(spTplRaw, selectedEvent.emailTemplateOverrides, 'Nachruecken');
              if (spTpl) { emailData = buildEmailFromTemplate(spTpl, promoteVars); }
              else { emailData = promotionEmail(promotedFirstName, selectedEvent.title); }
              await eventServiceRef.queueEmail(
                withParentTitleSubject(emailData.subject, selectedEvent.parentEventId ? allEvents.find(e => e.id === selectedEvent.parentEventId) : undefined),
                promoted.email, promoted.name || '', emailData.body,
                'Nachruecken', selectedEvent.title, selectedEvent.id
              );
            } catch (err) { console.warn('[DEX] promote-email failed:', err); }
          }
          if (!selectedEvent.disableOutlook) {
            try { await eventServiceRef.queueOutlookEvent(promoted.email, selectedEvent.id, selectedEvent.title, 'Einladen'); }
            catch (err) { console.warn('[DEX] promote-outlook failed:', err); }
          }
        } else {
          setAdminToast({ kind: 'no-promote', name });
        }
      } catch (err) {
        console.warn('[DEX] promoteFirstWaitlistItem failed:', err);
        setAdminToast({ kind: 'no-promote', name });
      }
    }
    if (selectedEvent.subsiteUrl && !eventWasOver) {
      try {
        const ok = await eventServiceRef.queueIDReorder(
          selectedEvent.id, selectedEvent.eventNumber || 0,
          selectedEvent.subsiteUrl, selectedEvent.title, name, reg.ParticipantEmail || undefined
        );
        if (!ok) {
          console.warn('[DEX] queueIDReorder returned false');
          showAlert(isDe ? 'Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.' : 'Cancellation successful, but the ID reorder entry could not be written to the queue. Please click "Reassign IDs" once.');
        }
      } catch (err) {
        console.warn('[DEX] queueIDReorder threw:', err);
        showAlert('Abmeldung erfolgreich, aber der ID-Reorder-Eintrag konnte nicht in die Queue geschrieben werden. Bitte einmal "IDs neu vergeben" klicken.');
      }
    }
    // v30.67 (Review): gemeinsamer Nachlade-Pfad — bei 429 nach dem Abmelden
    // ersetzte der Reload die Liste still durch `[]`.
    await reloadRegistrations();
  };

  // v23.2: Stilles Löschen einer doppelten Anmelde-Zeile. Anders als die
  // normale Abmeldung wird die Zeile HART gelöscht (kein „Abgemeldet"-Status,
  // der die Abmeldungs-Liste aufblähen würde) und es laufen KEINE Seiteneffekte
  // (keine Abmelde-Mail, kein Outlook-Ausladen, kein Nachrücken, kein
  // ID-Reorder, kein DEX_Participants-Cleanup) — die Person bleibt über ihre
  // andere Zeile regulär angemeldet. Sitzplatz-Counter wird nachgezogen.
  // v30.67: Liefert, ob die Zeile WIRKLICH weg ist. `deleteRegistration` ist
  // bewusst nicht-werfend (`return resp.ok`, `catch → false`) — das try/catch
  // hier fing also nie etwas, und jeder Aufrufer meldete „entfernt", auch
  // wenn die Zeile noch da war.
  const performSilentDuplicateDelete = async (reg: SPRegistration): Promise<boolean> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return false;
    const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
    setAdminToast({ kind: 'cancelling', name });
    let deleted = false;
    try {
      deleted = await eventServiceRef.deleteRegistration(selectedEvent.subsiteUrl, reg.Id);
      if (!deleted) console.warn('[DEX] performSilentDuplicateDelete: DELETE nicht ok für Item', reg.Id);
      // Audit-Eintrag und Sitzplatz-Sync nur, wenn die Zeile wirklich weg ist —
      // ein „RegistrationDeleted" für eine noch vorhandene Zeile wäre eine
      // falsche Historie.
      if (deleted) {
        try {
          await eventServiceRef.writeChangeLog({
            action: 'RegistrationDeleted',
            targetType: 'Participant',
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            targetId: ((reg as any).ParticipantEmail || '') + '#' + reg.Id,
            targetName: name,
            eventId: selectedEvent.id,
            eventTitle: selectedEvent.title,
            details: { note: 'Doppel-Anmeldung still entfernt (Duplikat). Person bleibt über die zweite Zeile angemeldet.' },
          });
        } catch (err) { console.warn('[DEX] writeChangeLog (dup delete) failed:', err); }
        try {
          const isSplit = typeof selectedEvent.durchstarterCapacity === 'number'
            && typeof selectedEvent.funstarterCapacity === 'number'
            && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
          await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit });
        } catch { /* best-effort */ }
      }
    } catch (err) {
      console.warn('[DEX] performSilentDuplicateDelete failed:', err);
    }
    await reloadRegistrations();
    setAdminToast(null);
    return deleted;
  };
  // v28.23: Doppelte Klammer-Schatten-Zeilen in einem Rutsch bereinigen.
  // Diese Zeilen tauchen in der konsolidierten Klammer-Tabelle NICHT auf (dort
  // steht pro Person genau eine Zeile, aggregiert über die Sub-Events) — ohne
  // diesen Knopf käme der Organizer also gar nicht an sie heran. Behalten wird
  // je Person die Zeile mit den meisten ausgefüllten Hauptevent-Antworten
  // (Tie-Break: die älteste), alle weiteren werden still gelöscht: keine Mail,
  // kein Outlook, kein Nachrücken.
  const [shadowDupBusy, setShadowDupBusy] = React.useState(false);
  // v28.23: Abgleich des zentralen Teilnehmer-Registers (DEX_Participants).
  const [isSyncingRegistry, setIsSyncingRegistry] = React.useState(false);
  const [syncRegistryResult, setSyncRegistryResult] = React.useState<string | null>(null);
  const cleanupShadowDuplicates = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl || shadowDupBusy) return;
    const groups: SPRegistration[][] = [];
    duplicateEmails.forEach(em => {
      const rows = registrations.filter(r => DUP_ACTIVE_STATI.indexOf(r.Status || '') >= 0
        && (r.ParticipantEmail || '').trim().toLowerCase() === em);
      if (rows.length > 1) groups.push(rows);
    });
    if (groups.length === 0) return;
    const extra = groups.reduce((n, rows) => n + rows.length - 1, 0);
    const ok = await confirmDialog(
      isDe
        ? `${extra} doppelte Klammer-Zeile(n) bei ${groups.length} Person(en) entfernen?\n\nJe Person bleibt die Zeile mit den meisten ausgefüllten Hauptevent-Antworten erhalten. Die Entfernung läuft still — ohne Abmelde-Mail, ohne Outlook-Absage, ohne Nachrücken. Die Anmeldungen in den Sub-Events bleiben unberührt.`
        : `Remove ${extra} duplicate overall-event row(s) for ${groups.length} person(s)?\n\nFor each person the row with the most main-event answers is kept. Removal is silent — no cancellation email, no Outlook removal, no waitlist promotion. The sub-event registrations are untouched.`,
      { danger: true, confirmLabel: isDe ? 'Zeilen entfernen' : 'Remove rows' },
    );
    if (!ok) return;
    setShadowDupBusy(true);
    const answerScore = (r: SPRegistration): number => {
      try {
        const o = JSON.parse(r.CustomData || '{}') as Record<string, unknown>;
        return Object.keys(o).filter(k => String(o[k] === null || o[k] === undefined ? '' : o[k]).trim()).length;
      } catch { return 0; }
    };
    let removed = 0;
    // v30.67: `removed` zählte VERSUCHE — `deleteRegistration` wirft nicht,
    // sondern liefert false. Die grüne Meldung „N entfernt" nannte also die
    // Zahl der Aufrufe, nicht der Löschungen.
    let failedDel = 0;
    for (const rows of groups) {
      const sorted = rows.slice().sort((a, b) => (answerScore(b) - answerScore(a)) || (a.Id - b.Id));
      for (const r of sorted.slice(1)) {
        try {
          const ok = await eventServiceRef.deleteRegistration(selectedEvent.subsiteUrl, r.Id);
          if (!ok) { failedDel += 1; console.warn('[DEX] cleanupShadowDuplicates: DELETE nicht ok für Item', r.Id); continue; }
          removed += 1;
          try {
            await eventServiceRef.writeChangeLog({
              action: 'RegistrationDeleted',
              targetType: 'Participant',
              targetId: `${r.ParticipantEmail || ''}#${r.Id}`,
              targetName: (r.Vorname && r.Nachname) ? `${r.Vorname} ${r.Nachname}` : (r.ParticipantName || ''),
              eventId: selectedEvent.id,
              eventTitle: selectedEvent.title,
              details: { note: 'Doppelte Klammer-Schatten-Zeile still entfernt (v28.23). Sub-Event-Anmeldungen unberührt.' },
            });
          } catch { /* best-effort */ }
        } catch (err) { failedDel += 1; console.warn('[DEX] cleanupShadowDuplicates failed for item', r.Id, err); }
      }
    }
    try {
      const isSplit = typeof selectedEvent.durchstarterCapacity === 'number'
        && typeof selectedEvent.funstarterCapacity === 'number'
        && ((selectedEvent.durchstarterCapacity || 0) > 0 || (selectedEvent.funstarterCapacity || 0) > 0);
      await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit });
    } catch { /* best-effort */ }
    try { await reloadRegistrations(); } catch { /* */ }
    setShadowDupBusy(false);
    showAlert(
      failedDel > 0
        ? (isDe
          ? `${removed} doppelte Klammer-Zeile(n) entfernt — ${failedDel} konnten NICHT entfernt werden (fehlende Rechte oder Drosselung) und sind noch da.`
          : `${removed} duplicate overall-event row(s) removed — ${failedDel} could NOT be removed (missing permissions or throttling) and are still there.`)
        : (isDe ? `${removed} doppelte Klammer-Zeile(n) entfernt.` : `${removed} duplicate overall-event row(s) removed.`),
      { variant: failedDel > 0 ? 'error' : 'success' },
    );
  };
  return {
    buildCancellationMail, cleanupShadowDuplicates, isSyncingRegistry,
    performSilentDuplicateDelete, performStandardCancel, setIsSyncingRegistry,
    setSyncRegistryResult, shadowDupBusy, syncRegistryResult,
  };
}

