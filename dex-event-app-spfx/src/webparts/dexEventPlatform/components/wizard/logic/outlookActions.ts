import * as React from 'react';
import { OutlookConfirmItem, SubEventDraft } from '../../wizard/wizardTypes';
import { EventService } from '../../../services/EventService';

/* createMissingOutlookAppointments — aus EventCreationPage.tsx ausgelagert (Zeilen 1770-1819 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface CreateMissingOutlookAppointmentsCtx {
  childTermPlural: string;
  childTermSingular: string;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  editEvent: import("../../../types/index").DeloitteEvent;
  flushActiveCommTabToState: () => void;
  forceOutlookRecreateRef: React.MutableRefObject<Set<string>>;
  handleSubmit: () => Promise<void>;
  isDe: boolean;
  outlookMissingTargets: () => { id: string; title: string; }[];
  parentTimesIso: () => {    start: string;    end: string;};
  pendingOutlookDirtyWriteRef: React.MutableRefObject<boolean>;
  pendingOutlookDirtyWriteRefs: React.MutableRefObject<Record<string, boolean>>;
  pendingOutlookRecreateForSubEventsRef: React.MutableRefObject<string[]>;
  pendingOutlookUpdateForSubEventsRef: React.MutableRefObject<string[]>;
  pendingOutlookUpdateForTopRef: React.MutableRefObject<boolean>;
  setOutlookUpdateBusy: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  subEventsRef: React.MutableRefObject<SubEventDraft[]>;
}

export async function createMissingOutlookAppointmentsImpl(ctx: CreateMissingOutlookAppointmentsCtx): Promise<void> {
  const { childTermPlural, childTermSingular, confirmDialog, editEvent, flushActiveCommTabToState, forceOutlookRecreateRef, handleSubmit, isDe, outlookMissingTargets, parentTimesIso, pendingOutlookDirtyWriteRef, pendingOutlookDirtyWriteRefs, pendingOutlookRecreateForSubEventsRef, pendingOutlookUpdateForSubEventsRef, pendingOutlookUpdateForTopRef, setOutlookUpdateBusy, showAlert, subEventsRef } = ctx;
    const mainId = editEvent?.id || '';
    const missing = outlookMissingTargets().filter(m => m.id && m.id !== mainId);
    if (missing.length === 0) return;
    const pt = parentTimesIso();
    const fmt = (iso: string): string => {
      if (!iso) return '—';
      const d = new Date(iso);
      return isNaN(d.getTime()) ? '—' : d.toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    };
    const plan = missing.map(m => {
      const d = subEventsRef.current.find(s => s.dbId === m.id);
      const start = (d && d.startDate) || pt.start || '';
      const end = (d && d.endDate) || (d && d.startDate) || pt.end || pt.start || '';
      const inherited = !(d && d.startDate);
      return { ...m, start, end, inherited };
    });
    if (plan.some(p => !p.start)) {
      showAlert(isDe
        ? 'Das Hauptevent hat keine Startzeit — ohne die lässt sich kein Termin anlegen. Bitte zuerst in Schritt 1 die Zeiten setzen und speichern.'
        : 'The main event has no start time — no appointment can be created without it. Please set the times in step 1 first and save.', { variant: 'error' });
      return;
    }
    const list = plan.map(p => `• ${p.title || '?'}: ${fmt(p.start)} – ${fmt(p.end)}${p.inherited ? (isDe ? '  (Zeiten des Hauptevents)' : '  (main event times)') : ''}`).join('\n');
    const ok = await confirmDialog(
      isDe
        ? `Für diese ${plan.length} ${plan.length === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} wird jetzt ein Outlook-Termin angelegt:\n\n${list}\n\nSub-Events ohne eigene Zeiten übernehmen die Zeiten des Hauptevents. Wenn du genauere Zeiten willst, brich hier ab, trage sie oben ein und klicke danach erneut.\n\nAnmeldungen, Teilnehmer-IDs und Teilnehmerlisten bleiben dabei unverändert.`
        : `An Outlook appointment will now be created for these ${plan.length} sub-event(s):\n\n${list}\n\nSub-events without their own times inherit the main event's times. If you want more precise times, cancel here, enter them above and click again.\n\nRegistrations, attendee IDs and attendee lists stay untouched.`,
      { confirmLabel: isDe ? 'Termine anlegen' : 'Create appointments' },
    );
    if (!ok) return;
    plan.forEach(p => forceOutlookRecreateRef.current.add(p.id));
    // v29.20 (Audit): wie attemptSubmit VOR dem Save flushen — dieser Pfad
    // rief handleSubmit direkt, und die Kommunikations-Eingaben des gerade
    // offenen Reiters lagen dann noch nicht im Draft (CLAUDE.md-Falle):
    // Der frisch angelegte Outlook-Termin trug den Text von VOR der letzten
    // Bearbeitung. Ebenso die pendingOutlook*-Reste eines früheren
    // Modal-Durchlaufs zurücksetzen, die hier sonst ungefragt nachwirkten.
    flushActiveCommTabToState();
    pendingOutlookDirtyWriteRef.current = null;
    pendingOutlookDirtyWriteRefs.current = {};
    pendingOutlookUpdateForTopRef.current = false;
    pendingOutlookUpdateForSubEventsRef.current = [];
    pendingOutlookRecreateForSubEventsRef.current = [];
    setOutlookUpdateBusy(true);
    try {
      await handleSubmit();
    } finally {
      setOutlookUpdateBusy(false);
      forceOutlookRecreateRef.current.clear();
    }
}

/* triggerOutlookUpdateNow — aus EventCreationPage.tsx ausgelagert (Zeilen 1822-1869 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface TriggerOutlookUpdateNowCtx {
  activeCommTabIdx: number;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  disableOutlook: boolean;
  editEvent: import("../../../types/index").DeloitteEvent;
  isDe: boolean;
  setOutlookUpdateBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setOutlookUpdateDone: React.Dispatch<React.SetStateAction<string>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  subEvents: SubEventDraft[];
  title: string;
}

export async function triggerOutlookUpdateNowImpl(ctx: TriggerOutlookUpdateNowCtx): Promise<void> {
  const { activeCommTabIdx, confirmDialog, disableOutlook, editEvent, isDe, setOutlookUpdateBusy, setOutlookUpdateDone, showAlert, subEvents, title } = ctx;
    let targetDbId = '';
    let targetTitle = title;
    let hasAppointment = false;
    if (activeCommTabIdx > 0) {
      const sub = subEvents[activeCommTabIdx - 1];
      if (sub) {
        targetDbId = sub.dbId || '';
        targetTitle = sub.title || title;
        hasAppointment = !!(sub.initialOutlookEventId || sub.initialCalendarLink);
      }
    } else {
      targetDbId = editEvent?.id || '';
      hasAppointment = !!(editEvent?.outlookEventId || editEvent?.calendarLink);
    }
    if (!editEvent || !targetDbId) {
      showAlert(isDe ? 'Den Outlook-Termin gibt es erst, nachdem das Event gespeichert wurde.' : 'The Outlook appointment only exists after the event has been saved.', { variant: 'info' });
      return;
    }
    if (disableOutlook) {
      showAlert(isDe ? 'Für diesen Tab ist der Outlook-Termin deaktiviert (Schalter weiter oben in Schritt 6).' : 'The Outlook appointment is disabled for this tab (toggle further up in step 6).', { variant: 'info' });
      return;
    }
    if (!hasAppointment) {
      showAlert(isDe ? 'Für dieses Event wurde noch kein Outlook-Termin angelegt — er entsteht beim Speichern.' : 'No Outlook appointment has been created for this event yet — it is created on save.', { variant: 'info' });
      return;
    }
    const ok = await confirmDialog(
      isDe
        ? `Der Outlook-Termin von „${targetTitle}" wird bei allen Teilnehmern mit dem zuletzt GESPEICHERTEN Stand aktualisiert. Falls du gerade etwas geändert hast, speichere bitte zuerst und klicke dann erneut hier.\n\nHinweis: Sub-Events haben eigene Termine — die aktualisierst du im jeweiligen Tab oder über „Alle Termine aktualisieren".`
        : `The Outlook appointment of „${targetTitle}" will be updated for all attendees with the last SAVED state. If you just changed something, please save first and then click here again.\n\nNote: sub-events have their own appointments — update them in their tab or via „Update all appointments".`,
      { confirmLabel: isDe ? 'Jetzt aktualisieren' : 'Update now' },
    );
    if (!ok) return;
    setOutlookUpdateBusy(true);
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      const svc = new EventService(ctx);
      await svc.queueOutlookEvent('', targetDbId, targetTitle, 'UpdateEvent');
      setOutlookUpdateDone(isDe
        ? `Angestoßen für „${targetTitle}" (${new Date().toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })} Uhr). Die Kalender der Teilnehmer aktualisieren sich in Kürze.`
        : `Triggered for „${targetTitle}" (${new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })}). Attendees' calendars will refresh shortly.`);
      showAlert(isDe ? 'Outlook-Aktualisierung wurde angestoßen — die Kalender der Teilnehmer aktualisieren sich in Kürze.' : 'Outlook update triggered — attendees will see the refreshed appointment shortly.', { variant: 'success' });
    } catch {
      showAlert(isDe ? 'Aktualisierung fehlgeschlagen — bitte erneut versuchen.' : 'Update failed — please try again.', { variant: 'error' });
    } finally {
      setOutlookUpdateBusy(false);
    }
}

/* triggerOutlookUpdateAll — aus EventCreationPage.tsx ausgelagert (Zeilen 1873-1904 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface TriggerOutlookUpdateAllCtx {
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  isDe: boolean;
  outlookUpdateTargets: () => { id: string; title: string; }[];
  setOutlookUpdateBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setOutlookUpdateDone: React.Dispatch<React.SetStateAction<string>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export async function triggerOutlookUpdateAllImpl(ctx: TriggerOutlookUpdateAllCtx): Promise<void> {
  const { confirmDialog, isDe, outlookUpdateTargets, setOutlookUpdateBusy, setOutlookUpdateDone, showAlert } = ctx;
    const targets = outlookUpdateTargets();
    if (targets.length === 0) return;
    const list = targets.map(t => `• ${t.title || '?'}`).join('\n');
    const ok = await confirmDialog(
      isDe
        ? `Alle ${targets.length} Outlook-Termine dieses Events mit dem zuletzt GESPEICHERTEN Stand aktualisieren?\n\n${list}\n\nJede/r Teilnehmer/in bekommt pro Termin, für den sie/er angemeldet ist, eine „Aktualisierter Termin"-Benachrichtigung.`
        : `Update all ${targets.length} Outlook appointments of this event with the last SAVED state?\n\n${list}\n\nEach attendee receives an „updated meeting" notification per appointment they are registered for.`,
      { confirmLabel: isDe ? 'Alle aktualisieren' : 'Update all' },
    );
    if (!ok) return;
    setOutlookUpdateBusy(true);
    let done = 0;
    let failed = 0;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      const svc = new EventService(ctx);
      for (const t of targets) {
        try { await svc.queueOutlookEvent('', t.id, t.title, 'UpdateEvent'); done += 1; }
        catch { failed += 1; }
      }
    } finally {
      setOutlookUpdateBusy(false);
    }
    const stamp = new Date().toLocaleTimeString(isDe ? 'de-DE' : 'en-GB', { hour: '2-digit', minute: '2-digit' });
    setOutlookUpdateDone(isDe
      ? `${done} Termin(e) angestoßen${failed > 0 ? `, ${failed} fehlgeschlagen` : ''} (${stamp} Uhr). Die Kalender der Teilnehmer aktualisieren sich in Kürze.`
      : `${done} appointment(s) triggered${failed > 0 ? `, ${failed} failed` : ''} (${stamp}). Attendees' calendars will refresh shortly.`);
    showAlert(
      isDe ? `${done} Outlook-Termin(e) angestoßen${failed > 0 ? `, ${failed} fehlgeschlagen` : ''}.` : `${done} Outlook appointment(s) triggered${failed > 0 ? `, ${failed} failed` : ''}.`,
      { variant: failed > 0 ? 'error' : 'success' },
    );
}

/* confirmOutlookSave — aus EventCreationPage.tsx ausgelagert (Zeilen 3729-3773 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface ConfirmOutlookSaveCtx {
  editEvent: import("../../../types/index").DeloitteEvent;
  handleSubmit: () => Promise<void>;
  outlookConfirmChecks: Record<string, boolean>;
  outlookConfirmItems: OutlookConfirmItem[];
  pendingOutlookDirtyWriteRef: React.MutableRefObject<boolean>;
  pendingOutlookDirtyWriteRefs: React.MutableRefObject<Record<string, boolean>>;
  pendingOutlookRecreateForSubEventsRef: React.MutableRefObject<string[]>;
  pendingOutlookUpdateForSubEventsRef: React.MutableRefObject<string[]>;
  pendingOutlookUpdateForTopRef: React.MutableRefObject<boolean>;
  setOutlookConfirmOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setTriggerOutlookUpdate: React.Dispatch<React.SetStateAction<boolean>>;
}

export function confirmOutlookSaveImpl(ctx: ConfirmOutlookSaveCtx): void {
  const { editEvent, handleSubmit, outlookConfirmChecks, outlookConfirmItems, pendingOutlookDirtyWriteRef, pendingOutlookDirtyWriteRefs, pendingOutlookRecreateForSubEventsRef, pendingOutlookUpdateForSubEventsRef, pendingOutlookUpdateForTopRef, setOutlookConfirmOpen, setTriggerOutlookUpdate } = ctx;
    setOutlookConfirmOpen(false);
    const topId = editEvent ? editEvent.id : '';
    const topItem = outlookConfirmItems.find(it => it.kind === 'top');
    const subItems = outlookConfirmItems.filter(it => it.kind === 'sub');
    const topChecked = !!topItem && !!outlookConfirmChecks[topItem.eventId];
    // v11.69: Angehakte Sub-Events trennen in:
    //  - `normalUpdateSubIds`: Sub-Event hat bereits einen Outlook-Termin →
    //    DEX_Outlook 'UpdateEvent' in die Queue schreiben (bestehender Pfad).
    //  - `recreateSubIds`: Sub-Event hat noch keinen Outlook-Termin
    //    (`noOutlookYet`) → DEX_Events-Item per `deleteEventItemOnly` löschen
    //    und mit `existingSubsiteUrl` neu anlegen, damit der
    //    DEX_CreateOutlookEvent-Flow triggert. Teilnehmer-Subsite + Liste
    //    bleiben unangetastet erhalten.
    const checkedSubItems = subItems.filter(it => !!outlookConfirmChecks[it.eventId]);
    const normalUpdateSubIds = checkedSubItems.filter(it => !it.noOutlookYet).map(it => it.eventId);
    const recreateSubIds = checkedSubItems.filter(it => !!it.noOutlookYet).map(it => it.eventId);
    pendingOutlookUpdateForTopRef.current = topChecked;
    pendingOutlookUpdateForSubEventsRef.current = normalUpdateSubIds;
    pendingOutlookRecreateForSubEventsRef.current = recreateSubIds;
    // Pro Event-ID den OutlookDirty-Schreibwert vormerken.
    // v11.69: noOutlookYet-Items werden — egal ob angehakt oder nicht — NICHT
    // dirty markiert. Bei angehakt erfolgt ein Recreate (neues Item hat von
    // Haus aus OutlookDirty=false), bei nicht angehakt existiert immer noch
    // kein Outlook-Termin der "aus-Sync" sein könnte → Marker wäre falsch.
    const dirtyMap: Record<string, boolean> = {};
    for (const it of outlookConfirmItems) {
      if (it.noOutlookYet) continue;
      dirtyMap[it.eventId] = !outlookConfirmChecks[it.eventId];
    }
    pendingOutlookDirtyWriteRefs.current = dirtyMap;
    // Top-Level kompatibel halten: wenn das Top-Event im Modal war, wird
    // OutlookDirty entsprechend gesetzt; sonst null = nicht anfassen.
    if (topItem) {
      pendingOutlookDirtyWriteRef.current = !topChecked;
    } else {
      pendingOutlookDirtyWriteRef.current = null;
    }
    // setTriggerOutlookUpdate steuert in handleSubmit, ob der Top-Level-
    // Outlook-Branch überhaupt betreten wird. v11.63: nur true wenn das
    // Top-Event angehakt wurde ODER mindestens ein Sub-Event angehakt
    // wurde (damit der Sub-Event-Branch im handleSubmit getroffen wird).
    setTriggerOutlookUpdate(topChecked || normalUpdateSubIds.length > 0 || recreateSubIds.length > 0);
    // Verhindern dass topId als „angehakt" interpretiert wird ohne Modal.
    void topId;
    handleSubmit().catch(() => { /* */ });
}

