/* createKlammerActions — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 1484-2059 des
 * Stands vor dem Schnitt). Der Rumpf ist zeichengleich uebernommen; was die
 * Gruppe aus dem Komponenten-Scope liest, kommt als `ctx` herein, was sie
 * nach aussen liefert, geht als Objekt zurueck.
 */
import * as React from 'react';
import { ConsolidatedRow } from '../../admin/adminTypes';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';
import { applyEventTemplateOverride, formatOrganizerList } from '../../../context/EventContext';
import { buildEmailFromTemplate, promotionEmail } from '../../../services/EmailTemplates';
import { invalidateInactiveAccountCache } from '../../../utils/accountCheckCache';
import { isEventOver } from '../../../utils/eventFormat';
import { withParentTitleSubject } from '../../../utils/mailSubject';

export interface CreateKlammerActionsCtx {
  assignAssistRow: ConsolidatedRow;
  assignAssistValue: string;
  buildCancellationMail: (ev: DeloitteEvent, reg: SPRegistration, fullName: string) => Promise<{    subject: string;    body: string;}>;
  bulkKlammerProgress: string;
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  consolidatedChildren: DeloitteEvent[];
  consolidatedRows: ConsolidatedRow[];
  currentUser: import("../../../types/index").User;
  deregModal: { emailKey: string; name: string; email: string; items: { child: DeloitteEvent; reg: SPRegistration; isParent?: boolean; }[]; };
  deregSelected: Set<string>;
  deregSilent: boolean;
  eventServiceRef: EventService;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  inactiveAccounts: string[];
  isDe: boolean;
  mainFieldsEditForm: Record<string, string>;
  mainFieldsEditName: string;
  mainFieldsEditReg: SPRegistration;
  mainFieldsEditSubsite: string;
  mainFieldsEditTargetIsParent: boolean;
  registerForEvent: (eventId: string, customData: Record<string, string>, participantFirstName?: string, participantLastName?: string, participantEmail?: string, preferredStarterType?: string, opts?: { skipShadowParent?: boolean; suppressMail?: boolean; suppressOutlook?: boolean; extraCc?: string; proxyConsentConfirmed?: boolean; actorAllowedAsAssistant?: boolean; skipReload?: boolean; bundledItems?: import("../../../utils/bundledComm").BundledItem[]; }) => Promise<{ ok: boolean; status: "Angemeldet" | "Warteliste"; reason?: string; }>;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setAddingToKlammer: React.Dispatch<React.SetStateAction<string>>;
  setAssignAssistBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setAssignAssistRow: React.Dispatch<React.SetStateAction<ConsolidatedRow>>;
  setAssignAssistValue: React.Dispatch<React.SetStateAction<string>>;
  setBulkKlammerProgress: React.Dispatch<React.SetStateAction<string>>;
  setDeregBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setDeregModal: React.Dispatch<React.SetStateAction<{ emailKey: string; name: string; email: string; items: { child: DeloitteEvent; reg: SPRegistration; isParent?: boolean; }[]; }>>;
  setDeregSelected: React.Dispatch<React.SetStateAction<Set<string>>>;
  setDeregSilent: React.Dispatch<React.SetStateAction<boolean>>;
  setMainFieldsEditError: React.Dispatch<React.SetStateAction<string>>;
  setMainFieldsEditForm: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  setMainFieldsEditName: React.Dispatch<React.SetStateAction<string>>;
  setMainFieldsEditReg: React.Dispatch<React.SetStateAction<SPRegistration>>;
  setMainFieldsEditSaving: React.Dispatch<React.SetStateAction<boolean>>;
  setMainFieldsEditSubsite: React.Dispatch<React.SetStateAction<string>>;
  setMainFieldsEditTargetIsParent: React.Dispatch<React.SetStateAction<boolean>>;
  setRegistrations: React.Dispatch<React.SetStateAction<SPRegistration[]>>;
  setSubEventRegsByEventId: React.Dispatch<React.SetStateAction<Record<string, SPRegistration[]>>>;
  setSubRegReloadTick: React.Dispatch<React.SetStateAction<number>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export interface CreateKlammerActionsResult {
  addAllToKlammer: (rows: ConsolidatedRow[]) => Promise<void>;
  addToKlammer: (row: ConsolidatedRow) => Promise<void>;
  closeDeregModal: () => void;
  closeMainFieldsEdit: () => void;
  openDeregModal: (row: ConsolidatedRow) => void;
  openMainFieldsEdit: (emailKey: string, displayName: string) => void;
  runDeregModal: () => Promise<void>;
  saveMainFieldsEdit: () => Promise<void>;
  submitAssignAssistant: () => Promise<void>;
}

export function createKlammerActions(ctx: CreateKlammerActionsCtx): CreateKlammerActionsResult {
  const {
    assignAssistRow, assignAssistValue, buildCancellationMail, bulkKlammerProgress, childEventsOf,
    confirmDialog, consolidatedChildren, consolidatedRows, currentUser, deregModal, deregSelected,
    deregSilent, eventServiceRef, getAllRegistrations, inactiveAccounts, isDe, mainFieldsEditForm,
    mainFieldsEditName, mainFieldsEditReg, mainFieldsEditSubsite, mainFieldsEditTargetIsParent,
    registerForEvent, registrations, selectedEvent, setAddingToKlammer, setAssignAssistBusy,
    setAssignAssistRow, setAssignAssistValue, setBulkKlammerProgress, setDeregBusy, setDeregModal,
    setDeregSelected, setDeregSilent, setMainFieldsEditError, setMainFieldsEditForm,
    setMainFieldsEditName, setMainFieldsEditReg, setMainFieldsEditSaving, setMainFieldsEditSubsite,
    setMainFieldsEditTargetIsParent, setRegistrations, setSubEventRegsByEventId,
    setSubRegReloadTick, showAlert,
  } = ctx;
  // v19.30 — Feature A: Edit-Modal für die Hauptevent-Custom-Felder einer
  // konsolidierten Zeile öffnen. Die Antworten stehen in der Registrierung der
  // Person auf der Hauptevent-Subsite. Wir suchen sie per E-Mail in
  // `registrations` (das ist die Teilnehmerliste des selektierten Hauptevents).
  // v24.38: Fehlende Klammer-/Hauptevent-Anmeldung nachtragen. Eine Person, die
  // nur in Sub-Events angemeldet ist, aber keine Schatten-Zeile auf der
  // Klammer-Teilnehmerliste hat (Daten-Anomalie), wird hier vom Organizer
  // händisch ergänzt. Das ist genau die „Schatten-Registrierung" des
  // subEventsOnlyMode — `registerForEvent` auf das Klammer-Event unterdrückt
  // bei subEventsOnlyMode automatisch Mail + Outlook (suppressParentNotifications).
  // v30.14: Kern ohne Dialog/Alerts/Reload — wird vom Einzel-Knopf UND vom
  // Sammel-Fix („Alle still nachtragen") benutzt. Liefert true bei Erfolg.
  const addToKlammerCore = async (row: ConsolidatedRow): Promise<boolean> => {
    if (!selectedEvent) return false;
    const name = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email;
    try {
      // v24.39: Realen Registranten aus den Sub-Event-Zeilen ableiten (wer die
      // Sub-Events angemeldet hat). Die Klammer-Schatten-Zeile wird DEMSELBEN
      // zugeschrieben — NICHT dem Admin, der nur die Datenkorrektur macht.
      // Dadurch (1) taucht die Zeile nicht fälschlich im „Assistenz" des Admins
      // auf und (2) sieht die echte Assistenz die Klammer-Anmeldung in IHRER
      // „Assistenz"-Kachel und kann die Klammer-Felder dort pflegen.
      let realByEmail = '';
      let realByName = '';
      const subRegs = (Object.values(row.perChild).filter(Boolean) as SPRegistration[])
        .slice()
        .sort((a, b) => {
          const ta = a.RegistrationDate ? new Date(a.RegistrationDate).getTime() : Number.POSITIVE_INFINITY;
          const tb = b.RegistrationDate ? new Date(b.RegistrationDate).getTime() : Number.POSITIVE_INFINITY;
          return ta - tb;
        });
      for (const r of subRegs) {
        if ((r.RegisteredByEmail || '').trim()) { realByEmail = (r.RegisteredByEmail || '').trim(); realByName = (r.RegisteredByName || '').trim(); break; }
      }
      // Fallback: keine Stellvertreter-Info auf den Subs → als Selbst-Anmeldung
      // behandeln (Schatten gehört dann der Person selbst, nicht dem Admin).
      if (!realByEmail) { realByEmail = row.email; realByName = name; }

      // v30.14: skipReload — der Kern läuft im Sammel-Fix in Serie; der eine
      // Refresh kommt vom Aufrufer, sonst zöge jede Person einen loadEvents.
      // v30.57: KEIN `proxyConsentConfirmed` mehr.
      //
      // Das Nachtragen einer fehlenden Klammer-Zeile ist eine Datenkorrektur —
      // niemand hat dabei jemanden um Zustimmung gefragt. Das Flag schrieb
      // aber genau das in die Spalte `ProxyConsent`: „Zustimmung der Person
      // zur stellvertretenden Anmeldung bestätigt durch <Admin> am <Datum>".
      // Zusammen mit dem `RegisteredBy`-Rückschreiben zwei Zeilen weiter unten
      // entstand ein Datensatz, der sich selbst widerspricht: angemeldet von
      // der Person selbst, Zustimmung bestätigt durch jemand anderen.
      //
      // Ein erfundener Zustimmungsnachweis ist die unangenehmste Sorte
      // falscher Daten — er sieht aus wie ein Beleg. Das Flag steuert
      // ausschließlich diesen Text (s. EventContext, `proxyConsentStr`) und
      // ist KEIN Rechte-Schalter; ohne es bleibt die Spalte leer, und wer die
      // Korrektur ausgelöst hat, steht ohnehin im ChangeLog-Eintrag unten.
      // `actorAllowedAsAssistant` bleibt — das ist die Rechte-Seite.
      const res = await registerForEvent(
        selectedEvent.id, {}, row.vorname || '', row.nachname || '', row.email, undefined,
        { suppressMail: true, suppressOutlook: true, actorAllowedAsAssistant: true, skipReload: true }
      );
      if (res && res.ok) {
        const regs = await getAllRegistrations(selectedEvent.id);
        // Schatten-Zeile dem realen Registranten zuschreiben (registerForEvent
        // hat den eingeloggten Admin als RegisteredBy gesetzt).
        const newParent = regs.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey);
        if (newParent && eventServiceRef && selectedEvent.subsiteUrl
          && realByEmail.toLowerCase() !== (currentUser.email || '').toLowerCase()) {
          try {
            await eventServiceRef.adminUpdateRegistration(
              selectedEvent.subsiteUrl, newParent.Id,
              { RegisteredByEmail: realByEmail, RegisteredByName: realByName },
              { name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email, email: currentUser.email }
            );
          } catch { /* best-effort */ }
        }
        if (eventServiceRef) {
          try {
            await eventServiceRef.writeChangeLog({
              action: 'ParticipantUpdated',
              targetType: 'Participant',
              targetId: (row.email || '') + '#klammer',
              targetName: name,
              eventId: selectedEvent.id,
              eventTitle: selectedEvent.title,
              details: { scope: 'addedMissingKlammerRegistration', actorEmail: currentUser.email, attributedTo: realByEmail },
            });
          } catch { /* */ }
        }
        return true;
      }
      return false;
    } catch (err) {
      console.warn('[DEX] addToKlammer error:', err);
      return false;
    }
  };

  const addToKlammer = async (row: ConsolidatedRow): Promise<void> => {
    if (!selectedEvent) return;
    const name = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email;
    if (!(await confirmDialog(
      isDe
        ? `Fehlende Hauptanmeldung: „${name}" ist nur in Sub-Events angemeldet, fehlt aber am Klammer-Event „${selectedEvent.title}".\n\nDie fehlende Klammer-Anmeldung jetzt ergänzen? (Es wird KEINE Mail und KEIN Outlook-Termin versendet — reine Datenkorrektur.)`
        : `Missing main registration: „${name}" is only in sub-events but missing on the umbrella event „${selectedEvent.title}".\n\nAdd the missing umbrella registration now? (No email and no Outlook invite are sent — data correction only.)`,
      { confirmLabel: isDe ? 'Hinzufügen' : 'Add' }
    ))) return;
    setAddingToKlammer(row.emailKey);
    try {
      const ok = await addToKlammerCore(row);
      if (ok) {
        const regs2 = await getAllRegistrations(selectedEvent.id);
        setRegistrations(regs2);
        showAlert(isDe ? `„${name}" wurde zum Klammer-Event hinzugefügt.` : `„${name}" was added to the umbrella event.`, { variant: 'success' });
      } else {
        showAlert(isDe ? 'Hinzufügen fehlgeschlagen — bitte erneut versuchen.' : 'Adding failed — please try again.', { variant: 'error' });
      }
    } finally {
      setAddingToKlammer(null);
    }
  };

  // v30.14: Sammel-Fix — ALLE fehlenden Klammer-Anmeldungen still nachtragen.
  // Sequentiell mit Fehlerzähler (CLAUDE.md-Regel: prüfbar, kein Promise.all-
  // Feuerwerk unter Drosselung); jede Zeile läuft über denselben Kern wie der
  // Einzel-Knopf (inkl. Zuschreibung an den realen Registranten + ChangeLog).
  const addAllToKlammer = async (rows: ConsolidatedRow[]): Promise<void> => {
    if (!selectedEvent || rows.length === 0 || bulkKlammerProgress) return;
    if (!(await confirmDialog(
      isDe
        ? `Alle ${rows.length} fehlenden Klammer-Anmeldungen jetzt still nachtragen? (Es wird KEINE Mail und KEIN Outlook-Termin versendet — reine Datenkorrektur. Die Zeilen werden der jeweils anmeldenden Person zugeschrieben.)`
        : `Add all ${rows.length} missing umbrella registrations silently now? (No email and no Outlook invite are sent — data correction only. Rows are attributed to whoever registered the sub-events.)`,
      { confirmLabel: isDe ? `Alle ${rows.length} nachtragen` : `Add all ${rows.length}` }
    ))) return;
    let okCount = 0;
    let failCount = 0;
    try {
      for (let i = 0; i < rows.length; i++) {
        setBulkKlammerProgress(`${i + 1}/${rows.length}`);
        if (await addToKlammerCore(rows[i])) okCount++; else failCount++;
      }
      const regs2 = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs2);
      showAlert(
        failCount === 0
          ? (isDe ? `${okCount} Klammer-Anmeldungen nachgetragen.` : `${okCount} umbrella registrations added.`)
          : (isDe
            ? `${okCount} Klammer-Anmeldungen nachgetragen, ${failCount} fehlgeschlagen (typisch: SharePoint-Drosselung) — bitte in ein paar Minuten erneut ausführen, bereits nachgetragene werden übersprungen.`
            : `${okCount} umbrella registrations added, ${failCount} failed (typically SharePoint throttling) — please run again in a few minutes; already-added ones are skipped.`),
        { variant: failCount === 0 ? 'success' : 'error' });
    } finally {
      setBulkKlammerProgress('');
    }
  };

  // v24.40: Eine Person (Klammer + alle aktiven Sub-Event-Anmeldungen) einer
  // gewählten Assistenz zuordnen, damit diese die Anmeldung in ihrer
  // „Assistenz"-Kachel vollständig verwalten kann. Setzt pro betroffener Zeile
  // RegisteredBy + Zeilen-Autor auf die Assistenz (eventService-Helfer).
  const submitAssignAssistant = async (): Promise<void> => {
    if (!assignAssistRow || !selectedEvent || !eventServiceRef) return;
    const m = (assignAssistValue || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
    if (!m) {
      showAlert(isDe ? 'Bitte eine Assistenz aus der Suche auswählen.' : 'Please select an assistant from the search.', { variant: 'error' });
      return;
    }
    const assistName = m[1].trim();
    const assistEmail = m[2].trim();
    const row = assignAssistRow;
    const personName = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email;
    if ((assistEmail || '').toLowerCase() === (row.emailKey || '')) {
      showAlert(isDe ? 'Die Assistenz darf nicht dieselbe Person wie die angemeldete Person sein.' : 'The assistant must not be the same person as the registered person.', { variant: 'error' });
      return;
    }
    setAssignAssistBusy(true);
    try {
      let done = 0;
      let failed = 0;
      // 1) Klammer-/Hauptevent-Zeile.
      const parentReg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey);
      if (parentReg && selectedEvent.subsiteUrl) {
        const ok = await eventServiceRef.assignRegistrationToAssistant(selectedEvent.subsiteUrl, parentReg.Id, assistEmail, assistName);
        if (ok) done += 1; else failed += 1;
      }
      // 2) Alle aktiven Sub-Event-Zeilen der Person.
      const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
      for (const ch of consolidatedChildren) {
        const r = row.perChild[ch.id];
        if (!r || ACTIVE.indexOf(r.Status) < 0 || !ch.subsiteUrl) continue;
        const ok = await eventServiceRef.assignRegistrationToAssistant(ch.subsiteUrl, r.Id, assistEmail, assistName);
        if (ok) done += 1; else failed += 1;
      }
      try {
        await eventServiceRef.writeChangeLog({
          action: 'ParticipantUpdated',
          targetType: 'Participant',
          targetId: (row.email || '') + '#assistant',
          targetName: personName,
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { scope: 'assignedAssistant', assistantEmail: assistEmail, actorEmail: currentUser.email, rowsUpdated: done },
        });
      } catch { /* */ }
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
      setSubRegReloadTick(t => t + 1);
      setAssignAssistRow(null);
      setAssignAssistValue('');
      if (failed === 0) {
        showAlert(isDe ? `„${personName}" wurde der Assistenz „${assistName}" zugeordnet (${done} Anmeldung(en)). Sie kann die Anmeldung jetzt in ihrer „Assistenz"-Kachel verwalten.` : `„${personName}" was assigned to assistant „${assistName}" (${done} registration(s)). They can now manage it in their „Assistant" tile.`, { variant: 'success' });
      } else {
        showAlert(isDe ? `Teilweise zugeordnet: ${done} erfolgreich, ${failed} fehlgeschlagen (evtl. fehlende Rechte). Bei externen/nicht auffindbaren Konten ist die Zuordnung nicht möglich.` : `Partially assigned: ${done} ok, ${failed} failed (possibly missing permissions).`, { variant: 'error' });
      }
    } catch (err) {
      console.warn('[DEX] submitAssignAssistant error:', err);
      showAlert(isDe ? 'Unerwarteter Fehler bei der Zuordnung.' : 'Unexpected error during assignment.', { variant: 'error' });
    } finally {
      setAssignAssistBusy(false);
    }
  };

  const openMainFieldsEdit = (emailKey: string, displayName: string): void => {
    if (!selectedEvent) return;
    setMainFieldsEditError('');
    // 1) Bevorzugt die Hauptevent-Anmeldung (Klammer-Subsite).
    const parentReg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === emailKey) || null;
    let targetReg: SPRegistration | null = parentReg;
    let targetSubsite = selectedEvent.subsiteUrl || '';
    let isParent = true;
    // 2) Fallback (Klammer-Modus ohne Hauptevent-Anmeldung): die früheste
    //    aktive Sub-Event-Zeile der Person als Speicherort nehmen.
    if (!parentReg) {
      const row = consolidatedRows.find(r => r.emailKey === emailKey);
      if (row) {
        const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
        let best: { reg: SPRegistration; sub: string; ts: number } | null = null;
        for (const ch of consolidatedChildren) {
          const r = row.perChild[ch.id];
          if (!r || ACTIVE.indexOf(r.Status) < 0) continue;
          const ts = r.RegistrationDate ? new Date(r.RegistrationDate).getTime() : Number.POSITIVE_INFINITY;
          if (!best || ts < best.ts) best = { reg: r, sub: ch.subsiteUrl || '', ts };
        }
        if (best && best.sub) { targetReg = best.reg; targetSubsite = best.sub; isParent = false; }
      }
    }
    setMainFieldsEditReg(targetReg);
    setMainFieldsEditSubsite(targetSubsite);
    setMainFieldsEditTargetIsParent(isParent);
    setMainFieldsEditName(displayName);
    const initial: Record<string, string> = {};
    if (targetReg) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyReg = targetReg as any;
      let cd: Record<string, unknown> = {};
      if (anyReg.CustomData) { try { cd = JSON.parse(anyReg.CustomData); } catch { /* */ } }
      const parentFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'document' && f.label && f.label.trim());
      for (const f of parentFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = (f as any).spInternalName || '';
        // Bei Sub-Event-Ziel NUR aus CustomData lesen (die Klammer-SP-Spalten
        // existieren auf der Sub-Event-Liste i.d.R. nicht).
        let v: unknown = (isParent && sp) ? anyReg[sp] : undefined;
        if (v === undefined || v === null || v === '') v = cd[f.id];
        initial[f.id] = (v === undefined || v === null) ? '' : String(v);
      }
    }
    setMainFieldsEditForm(initial);
  };
  const closeMainFieldsEdit = (): void => {
    setMainFieldsEditReg(null);
    setMainFieldsEditName('');
    setMainFieldsEditForm({});
    setMainFieldsEditError('');
    setMainFieldsEditSubsite('');
    setMainFieldsEditTargetIsParent(true);
  };
  // v19.30 — Feature A: Speichern der Hauptevent-Custom-Felder. Persistiert
  // über dasselbe `adminUpdateRegistration` wie das reguläre Teilnehmer-Edit
  // (schreibt die SP-Spalten der Hauptevent-Teilnehmerliste) und legt eine
  // Audit-Zeile 'ParticipantUpdated' mit dem Vorher/Nachher-Diff an. Es werden
  // nur geänderte Felder ins Patch aufgenommen — sonst kippt ein unverändertes
  // Choice-Feld den ganzen Save (HTTP 400 'Invalid choice').
  const saveMainFieldsEdit = async (): Promise<void> => {
    const targetSubsite = mainFieldsEditSubsite || selectedEvent?.subsiteUrl || '';
    if (!mainFieldsEditReg || !eventServiceRef || !targetSubsite || !selectedEvent) return;
    const isParentTarget = mainFieldsEditTargetIsParent;
    setMainFieldsEditSaving(true);
    setMainFieldsEditError('');
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyReg = mainFieldsEditReg as any;
      let cd: Record<string, unknown> = {};
      if (anyReg.CustomData) { try { cd = JSON.parse(anyReg.CustomData); } catch { /* */ } }
      const parentFields = (selectedEvent?.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'document' && f.label && f.label.trim());
      const patch: Record<string, unknown> = {};
      const oldValues: Record<string, unknown> = {};
      const fieldLabelMap: Record<string, string> = {};
      const nextCd: Record<string, unknown> = { ...cd };
      let cdChanged = false;
      for (const f of parentFields) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = (f as any).spInternalName || '';
        const newVal = mainFieldsEditForm[f.id] || '';
        let oldVal = '';
        // Bei Sub-Event-Ziel den alten Wert NUR aus CustomData lesen (keine
        // Klammer-SP-Spalten auf der Sub-Event-Liste).
        let oldFromSp: unknown = (isParentTarget && sp) ? anyReg[sp] : undefined;
        if (oldFromSp === undefined || oldFromSp === null || oldFromSp === '') oldFromSp = cd[f.id];
        if (oldFromSp !== undefined && oldFromSp !== null) oldVal = String(oldFromSp);
        if (newVal === oldVal) continue; // unverändert → überspringen
        const keyForAudit = (isParentTarget && sp) ? sp : f.id;
        fieldLabelMap[keyForAudit] = f.label;
        oldValues[keyForAudit] = oldVal;
        // SP-Spalte NUR beim Hauptevent-Ziel patchen (Sub-Event-Liste hat die
        // Klammer-Spalten nicht → würde HTTP 400 werfen). Sonst CustomData-only.
        if (isParentTarget && sp) patch[sp] = newVal;
        nextCd[f.id] = newVal;
        cdChanged = true;
      }
      if (!cdChanged && Object.keys(patch).length === 0) {
        closeMainFieldsEdit();
        return;
      }
      // CustomData immer mitschreiben, damit der konsolidierte View (der bei
      // fehlender SP-Spalte auf CustomData zurückfällt) konsistent bleibt.
      if (cdChanged) patch['CustomData'] = JSON.stringify(nextCd);
      const actor = {
        name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email,
        email: currentUser.email,
      };
      const ok = await eventServiceRef.adminUpdateRegistration(
        targetSubsite, mainFieldsEditReg.Id, patch, actor, oldValues, fieldLabelMap
      );
      if (!ok) {
        setMainFieldsEditError(isDe
          ? 'Speichern fehlgeschlagen — vermutlich fehlt eine SP-Spalte in der Hauptevent-Teilnehmerliste. Klicke einmal „Spalten fixen" für das Hauptevent, dann erneut versuchen.'
          : 'Save failed — likely a missing SP column on the main-event participant list. Click „Fix columns" for the main event once, then retry.');
        return;
      }
      // Audit-Log mit Diff der geänderten Felder (analog saveEdit).
      try {
        const changes: Record<string, { old: unknown; new: unknown }> = {};
        for (const k of Object.keys(oldValues)) {
          changes[k] = { old: oldValues[k], new: (k in patch ? patch[k] : mainFieldsEditForm[k]) };
        }
        await eventServiceRef.writeChangeLog({
          action: 'ParticipantUpdated',
          targetType: 'Participant',
          targetId: (mainFieldsEditReg.ParticipantEmail || '') + '#' + mainFieldsEditReg.Id,
          targetName: `${mainFieldsEditReg.Vorname || ''} ${mainFieldsEditReg.Nachname || ''}`.trim() || mainFieldsEditName,
          eventId: selectedEvent.id,
          eventTitle: selectedEvent.title,
          details: { changes, scope: 'mainEventFields' },
        });
      } catch { /* */ }
      // Neu laden, damit die Klammer-Feld-Spalten die neuen Werte zeigen.
      // Hauptevent-Ziel → Hauptevent-Teilnehmerliste; Sub-Event-Ziel → die
      // konsolidierten Sub-Event-Registrierungen neu ziehen.
      if (isParentTarget) {
        const regs = await getAllRegistrations(selectedEvent.id);
        setRegistrations(regs);
      } else {
        setSubRegReloadTick(t => t + 1);
      }
      closeMainFieldsEdit();
    } catch (err) {
      console.warn('[DEX] saveMainFieldsEdit error:', err);
      setMainFieldsEditError(isDe
        ? 'Unerwarteter Fehler beim Speichern.'
        : 'Unexpected error while saving.');
    } finally {
      setMainFieldsEditSaving(false);
    }
  };

  // v19.30 — Feature B: Sub-Event-Registrierungen neu laden (nach einer
  // Abmeldung im konsolidierten View). Spiegelt den Lade-Effekt von oben.
  const reloadSubEventRegs = async (): Promise<void> => {
    if (!selectedEvent || !selectedEvent.subEventsOnlyMode) return;
    const children = childEventsOf(selectedEvent.id);
    const map: Record<string, SPRegistration[]> = {};
    for (const ch of children) {
      try { map[ch.id] = await getAllRegistrations(ch.id); }
      catch { map[ch.id] = []; }
    }
    setSubEventRegsByEventId(map);
  };
  // v19.30 — Feature B: Abmelde-Modal für eine konsolidierte Zeile öffnen.
  // Sammelt alle Sub-Events, in denen die Person aktiv angemeldet ist.
  const openDeregModal = (row: ConsolidatedRow): void => {
    if (!selectedEvent) return;
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const items: Array<{ child: DeloitteEvent; reg: SPRegistration; isParent?: boolean }> = [];
    // v29.29: Die Klammer-Zeile ZUERST — sie gehört zur Person genauso wie die
    // Sub-Events. Ohne sie blieb nach dem Abmelden aller Sub-Events eine
    // Schatten-Anmeldung auf dem Hauptevent stehen, die die Teilnehmerzahl
    // weiter mitzählte und in „Meine Events" der Person erschien.
    const parentReg = registrations.find(r =>
      (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey && ACTIVE.indexOf(r.Status) >= 0);
    if (parentReg) items.push({ child: selectedEvent, reg: parentReg, isParent: true });
    for (const ch of childEventsOf(selectedEvent.id)) {
      const r = row.perChild[ch.id];
      if (r && ACTIVE.indexOf(r.Status) >= 0) items.push({ child: ch, reg: r });
    }
    setDeregModal({
      emailKey: row.emailKey,
      name: `${row.vorname} ${row.nachname}`.trim() || row.email,
      email: row.email,
      items,
    });
    // Default: alles vorausgewählt — der häufigste Fall ist „ganz abmelden".
    // Der Organizer kann einzelne wieder abwählen.
    setDeregSelected(new Set(items.map(i => i.child.id)));
    // v29.29: Bei einem als inaktiv gemeldeten Konto (Person hat das
    // Unternehmen verlassen) ist die stille Abmeldung der Normalfall.
    setDeregSilent(inactiveAccounts.indexOf(row.emailKey) >= 0);
  };
  const closeDeregModal = (): void => {
    setDeregModal(null);
    setDeregSelected(new Set());
    setDeregSilent(false);
  };
  // v19.30 — Feature B: Abmeldung pro gewähltem Sub-Event durchführen. Spiegelt
  // exakt die Nebenwirkungen des Einzel-Event-Abmeldens (Abmelde-Mail +
  // Outlook 'Ausladen' + DEX_Participants-Cleanup + Nachrücken + ID-Reorder)
  // und schreibt zusätzlich pro Abmeldung eine 'RegistrationCancelled'-
  // Audit-Zeile (die der Einzel-Pfad nicht setzt).
  const runDeregModal = async (): Promise<void> => {
    if (!deregModal || !eventServiceRef) return;
    setDeregBusy(true);
    const actorName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
    const actorEmail = currentUser.email;
    const chosen = deregModal.items.filter(i => deregSelected.has(i.child.id));
    for (const { child, reg, isParent } of chosen) {
      const sub = child.subsiteUrl;
      if (!sub) continue;
      const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
      const cancelledStarterType = reg.StarterType || '';
      try {
        await eventServiceRef.cancelRegistration(sub, reg.Id, actorName, actorEmail);
        // Audit-Zeile (Feature D: Abmeldungen sollen im Event-Log auftauchen).
        try {
          await eventServiceRef.writeChangeLog({
            action: 'RegistrationCancelled',
            targetType: 'Participant',
            targetId: (reg.ParticipantEmail || '') + '#' + reg.Id,
            targetName: name,
            eventId: child.id,
            eventTitle: child.title,
            details: { asActor: 'organizer', via: 'consolidatedDeregister', ...(deregSilent ? { silent: true } : {}), ...(isParent ? { level: 'parent' } : {}) },
          });
        } catch { /* */ }
        // Abmelde-Mail + Outlook 'Ausladen' (event-weite Schalter respektieren).
        // v22.22: Vergangenes Sub-Event → stille Abmeldung (keine Mail, kein
        // Outlook, kein Nachrücken, kein ID-Reorder).
        const childWasOver = isEventOver(child);
        // v29.29: `deregSilent` unterdrückt NUR die Benachrichtigung der
        // ausscheidenden Person (erloschenes Postfach) — Nachrücken und
        // ID-Reorder laufen weiter, der frei gewordene Platz soll ja an die
        // Warteliste gehen und die nachrückende Person ihre Mail bekommen.
        const notifyLeaver = !childWasOver && !deregSilent;
        // v29.44: Auf der KLAMMER keine zweite Abmelde-Bestätigung, wenn im
        // selben Lauf auch Sub-Events abgemeldet werden — dafür ging deren
        // eigene Mail schon raus. Die Klammer-Zeile ist seit v29.29 Teil des
        // Dialogs; seither bekam der Teilnehmer zusätzlich eine Mail mit dem
        // Klammer-Titel, obwohl er sich von einem Termin abgemeldet hat. Im
        // Modus „nur Sub-Events" ist die Klammer ohnehin nur eine
        // Schattenzeile — dort nie eine eigene Mail.
        const skipParentMail = !!isParent && (!!child.subEventsOnlyMode || chosen.some(i => !i.isParent));
        if (reg.ParticipantEmail && notifyLeaver) {
          if (!child.disableEmails && !child.disableCancellationEmail && !skipParentMail) {
            try {
              // v29.44: die für das Event gepflegte Abmelde-Vorlage nehmen —
              // vorher IMMER der Code-Standardtext. Deshalb sah die vom
              // Organizer ausgelöste Abmeldung anders aus als die, die der
              // Teilnehmer beim Selbst-Abmelden bekommt (dort läuft es seit
              // jeher über Vorlage + Event-Override). Gleicher Weg wie beim
              // Nachrücken ein paar Zeilen weiter unten.
              const emailData = await buildCancellationMail(child, reg, name);
              await eventServiceRef.queueEmail(
                emailData.subject, reg.ParticipantEmail, name, emailData.body,
                'Abmeldung', child.title, child.id
              );
            } catch (err) { console.warn('[DEX]', err); }
          }
          if (!child.disableOutlook) {
            try {
              await eventServiceRef.queueOutlookEvent(
                reg.ParticipantEmail, child.id, child.title, 'Ausladen'
              );
            } catch (err) { console.warn('[DEX]', err); }
          }
        }
        // DEX_Participants aufräumen.
        if (reg.ParticipantEmail && child.eventNumber) {
          eventServiceRef.removeParticipantEvent(reg.ParticipantEmail, child.eventNumber)
            .catch(err => console.warn('[DEX]', err));
        }
        // Client-seitiges Nachrücken (typ-bewusst bei Split-Capacity, außer
        // splitSharedWaitlist) — identisch zum Einzel-Event-Abmelden.
        const isSplitEvent = typeof child.durchstarterCapacity === 'number'
          && typeof child.funstarterCapacity === 'number'
          && ((child.durchstarterCapacity || 0) > 0 || (child.funstarterCapacity || 0) > 0);
        const useTypeFilter = isSplitEvent && !child.splitSharedWaitlist;
        if (!childWasOver) {
        // v27.11: Kein automatisches Nachrücken, wenn die Warteliste des
        // Sub-Events abgeschaltet ist (Kill-Switch, s. Einzel-Event-Abmelden).
        // Der ID-Reorder unten läuft weiterhin.
        // v29.29: Auf der KLAMMER nie nachrücken — sie ist keine
        // Anmeldeeinheit (maxParticipants ist dort 0, die Zeile ist eine
        // Schattenzeile). Ein Nachrücken dort würde eine fremde Person auf
        // eine Ebene heben, die niemand bucht.
        if (!isParent && child.waitlistEnabled !== false) {
        try {
          const promoted = await eventServiceRef.promoteFirstWaitlistItem(
            sub,
            cancelledStarterType || undefined,
            child.maxParticipants,
            (useTypeFilter && cancelledStarterType) ? cancelledStarterType : undefined,
            { itemId: reg.Id, participantEmail: reg.ParticipantEmail || '' },
          );
          if (promoted && promoted.success && promoted.email) {
            if (!child.disableEmails) {
              try {
                const lang = child.emailLanguage || 'EN';
                const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
                const promoteVars = {
                  Name: promotedFirstName,
                  EventTitle: child.title,
                  Organizer: formatOrganizerList(child.organizers, lang),
                  AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
                  WaitlistPosition: '',
                };
                let emailData: { subject: string; body: string };
                const spTplRaw = await eventServiceRef.getEmailTemplate('Nachruecken', lang).catch(() => null);
                const spTpl = applyEventTemplateOverride(spTplRaw, child.emailTemplateOverrides, 'Nachruecken');
                if (spTpl) emailData = buildEmailFromTemplate(spTpl, promoteVars);
                else emailData = promotionEmail(promotedFirstName, child.title);
                await eventServiceRef.queueEmail(
                  withParentTitleSubject(emailData.subject, selectedEvent && selectedEvent.subEventCalendar ? selectedEvent : undefined),
                  promoted.email, promoted.name || '', emailData.body,
                  'Nachruecken', child.title, child.id
                );
              } catch (err) { console.warn('[DEX] promote-email failed:', err); }
            }
            if (!child.disableOutlook) {
              try {
                await eventServiceRef.queueOutlookEvent(promoted.email, child.id, child.title, 'Einladen');
              } catch (err) { console.warn('[DEX] promote-outlook failed:', err); }
            }
          }
        } catch (err) { console.warn('[DEX] promoteFirstWaitlistItem failed:', err); }
        }
        // ID-Reorder in die Queue (Flow macht nur noch Reorder).
        try {
          await eventServiceRef.queueIDReorder(
            child.id, child.eventNumber || 0, sub, child.title, name, reg.ParticipantEmail || undefined
          );
        } catch (err) { console.warn('[DEX] queueIDReorder threw:', err); }
        }
      } catch (err) {
        console.warn('[DEX] consolidated deregister failed for child', child.id, err);
      }
    }
    try { await reloadSubEventRegs(); } catch { /* */ }
    // v29.31: Gecachte „Konto inaktiv"-Ergebnisse dieser Event-Familie
    // verwerfen (s. performStandardCancel).
    invalidateInactiveAccountCache(chosen.map(c => c.child.id).concat(selectedEvent ? [selectedEvent.id] : []));
    // v29.29: Auch die Klammer-Teilnehmerliste neu laden — seit die
    // Hauptevent-Zeile mit abgemeldet werden kann, wäre die Kopfzeile
    // („Teilnehmer (N)") sonst bis zum nächsten Öffnen veraltet.
    if (selectedEvent) {
      try { setRegistrations(await getAllRegistrations(selectedEvent.id)); } catch { /* */ }
    }
    setDeregBusy(false);
    closeDeregModal();
  };
  return {
    addAllToKlammer, addToKlammer, closeDeregModal, closeMainFieldsEdit, openDeregModal,
    openMainFieldsEdit, runDeregModal, saveMainFieldsEdit, submitAssignAssistant,
  };
}

