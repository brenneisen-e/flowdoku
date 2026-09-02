/* useWaitlistActions — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 2070-2467 des
 * Stands vor dem Schnitt). Der Rumpf ist zeichengleich uebernommen; was die
 * Gruppe aus dem Komponenten-Scope liest, kommt als `ctx` herein, was sie
 * nach aussen liefert, geht als Objekt zurueck.
 */
import * as React from 'react';
import { EventService, SPRegistration } from '../../../services/EventService';
import { applyEventTemplateOverride, formatOrganizerList } from '../../../context/EventContext';
import { buildEmailFromTemplate, promotionEmail } from '../../../services/EmailTemplates';
import { withParentTitleSubject } from '../../../utils/mailSubject';
import { DeloitteEvent } from '../../../types';
import { AdminToastState } from '../../admin/adminTypes';
import { buildPromotionPlan, promotionPlanLines, isSplitCapacityOf, PromotionPlan } from '../../../utils/promotionPlan';
import { shortSubEventTitle } from '../../../utils/subEventTitle';

export interface UseWaitlistActionsCtx {
  allEvents: DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  eventServiceRef: EventService;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  isDe: boolean;
  isSplitCapacity: boolean;
  obKeepVariant: "active" | "firstWaitlist";
  obMailBody: string;
  obMailLang: "DE" | "EN";
  obMailSubject: string;
  obRemoveCalendar: boolean;
  obWithMail: boolean;
  overbookModal: { mode: "confirm" | "keep"; targets: SPRegistration[]; };
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  setAdminToast: React.Dispatch<React.SetStateAction<AdminToastState>>;
  setIsPromoting: React.Dispatch<React.SetStateAction<boolean>>;
  setIsReorderingIDs: React.Dispatch<React.SetStateAction<boolean>>;
  setObBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setObMailBody: React.Dispatch<React.SetStateAction<string>>;
  setObMailLang: React.Dispatch<React.SetStateAction<"DE" | "EN">>;
  setObMailSubject: React.Dispatch<React.SetStateAction<string>>;
  setOverbookModal: React.Dispatch<React.SetStateAction<{ mode: "confirm" | "keep"; targets: SPRegistration[]; }>>;
  setPromoteResult: React.Dispatch<React.SetStateAction<string>>;
  setReorderProgress: React.Dispatch<React.SetStateAction<number>>;
  setReorderProgressLabel: React.Dispatch<React.SetStateAction<string>>;
  setReorderResult: React.Dispatch<React.SetStateAction<string>>;
  /** v30.69: Sammel-Heilung „Nachrücken & IDs für ALLE Events nachholen". */
  setIsHealingAll: React.Dispatch<React.SetStateAction<boolean>>;
  setHealAllProgress: React.Dispatch<React.SetStateAction<string | null>>;
  setHealAllResult: React.Dispatch<React.SetStateAction<string | null>>;
}

export interface UseWaitlistActionsResult {
  idFixCheckedForRef: React.MutableRefObject<string>;
  idRecheckBusy: boolean;
  recentCancellation: (regs: SPRegistration[]) => {    recent: boolean;    whenIso: string;    detail: string;};
  reloadRegistrationsForIdCheck: () => Promise<void>;
  runIdReorder: () => Promise<void>;
  runManualPromote: () => Promise<void>;
  runOverbookResolution: () => Promise<void>;
  runHealAllEvents: () => Promise<void>;
}

export function useWaitlistActions(ctx: UseWaitlistActionsCtx): UseWaitlistActionsResult {
  const {
    allEvents, confirmDialog, eventServiceRef, getAllRegistrations, isDe, isSplitCapacity,
    obKeepVariant, obMailBody, obMailLang, obMailSubject, obRemoveCalendar, obWithMail,
    overbookModal, registrations, selectedEvent, setAdminToast, setIsPromoting, setIsReorderingIDs,
    setObBusy, setObMailBody, setObMailLang, setObMailSubject, setOverbookModal, setPromoteResult,
    reloadRegistrations, setReorderProgress, setReorderProgressLabel, setReorderResult,
    setIsHealingAll, setHealAllProgress, setHealAllResult,
  } = ctx;
  // v11.36: Fairer Wartelisten-Rang einer Person in ihrer Gruppe — gleiche
  // Logik wie die Review-Box. Genutzt für die "neue Warteliste-Position" im
  // Mailtext (Vorschlag + Sammel-Versand).
  const getFairWaitlistRank = (reg: SPRegistration): number => {
    if (!selectedEvent) return 0;
    const ACT = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
    const isSplit = isSplitCapacity;
    const keyOf = (r: SPRegistration): string => isSplit ? (r.StarterType || r.PreferredStarterType || '?') : 'all';
    const capOf = (k: string): number => !isSplit
      ? (selectedEvent.maxParticipants || 0)
      : (k === 'Durchstarter' ? (selectedEvent.durchstarterCapacity || 0) : k === 'Funstarter' ? (selectedEvent.funstarterCapacity || 0) : 0);
    const k = keyOf(reg);
    const activeSorted = registrations
      .filter(r => ACT.indexOf(r.Status) >= 0 && keyOf(r) === k)
      .slice().sort((a, b) => a.Id - b.Id);
    const cap = capOf(k);
    const overCap = cap > 0 ? activeSorted.slice(cap) : [];
    const existingWl = registrations.filter(r => r.Status === 'Warteliste' && keyOf(r) === k);
    const fairWl = [...overCap, ...existingWl].sort((a, b) =>
      new Date(a.RegistrationDate).getTime() - new Date(b.RegistrationDate).getTime());
    const idx = fairWl.findIndex(x => x.Id === reg.Id);
    return idx >= 0 ? idx + 1 : 0;
  };

  // v11.36: Beim Öffnen des „Bestätigen"-Dialogs die Mail-Sprache aus dem
  // Event vorbelegen (Default DE wenn nicht explizit EN) — umschaltbar.
  React.useEffect(() => {
    if (overbookModal?.mode === 'confirm') {
      setObMailLang((selectedEvent?.emailLanguage || '').toUpperCase() === 'EN' ? 'EN' : 'DE');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [overbookModal]);

  // v11.36: Mailtext vorbefüllen — reagiert auf Dialog-Öffnen UND Sprachwahl.
  // Enthält die neue Wartelisten-Position der ersten Zielperson.
  // v13.0: buildOverbookApologyEmail ist jetzt async (lädt Template aus
  // DEX_EmailTemplates). Effect wartet auf das Promise und setzt State
  // wenn der Modal noch offen ist.
  React.useEffect(() => {
    if (overbookModal?.mode === 'confirm' && eventServiceRef && selectedEvent) {
      const t = overbookModal.targets[0];
      const nm = t ? ((t.Vorname && t.Nachname) ? `${t.Vorname} ${t.Nachname}` : t.ParticipantName) : '';
      const pos = t ? getFairWaitlistRank(t) : 0;
      let cancelled = false;
      eventServiceRef.buildOverbookApologyEmail(nm, selectedEvent.title, obMailLang, pos).then(m => {
        if (cancelled) return;
        setObMailSubject(m.subject);
        setObMailBody(m.body);
      }).catch(() => { /* */ });
      return () => { cancelled = true; };
    }
    return undefined;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [overbookModal, obMailLang]);

  // v11.36: Überbuchungs-Entscheidung ausführen (einzeln oder Sammel) und
  // danach IDs neu vergeben + Counter/Seat-Sync + Liste neu laden.
  const runOverbookResolution = async (): Promise<void> => {
    if (!overbookModal || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setObBusy(true);
    const sub = selectedEvent.subsiteUrl;
    const isBulk = overbookModal.targets.length > 1;
    try {
      for (const reg of overbookModal.targets) {
        const grp = reg.StarterType || reg.PreferredStarterType || '';
        const nm = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
        if (overbookModal.mode === 'confirm') {
          await eventServiceRef.resolveOverbookToWaitlist(sub, reg.Id, grp);
          if (obWithMail && reg.ParticipantEmail && !selectedEvent.disableEmails) {
            // Einzeln: ggf. vom Admin editierter Text. Sammel: pro Person
            // frisch personalisiert aus dem Standard-Template.
            const mail = isBulk
              ? await eventServiceRef.buildOverbookApologyEmail(nm, selectedEvent.title, obMailLang, getFairWaitlistRank(reg))
              : { subject: obMailSubject, body: obMailBody };
            try {
              await eventServiceRef.queueEmail(
                mail.subject, reg.ParticipantEmail, nm, mail.body,
                'Info', selectedEvent.title, selectedEvent.id
              );
            } catch { /* Mail-Fehler darf Korrektur nicht blockieren */ }
          }
          if (obRemoveCalendar && reg.ParticipantEmail && !selectedEvent.disableOutlook) {
            try {
              await eventServiceRef.queueOutlookEvent(
                reg.ParticipantEmail, selectedEvent.id, selectedEvent.title, 'Ausladen'
              );
            } catch { /* Kalender-Abmeldung best-effort */ }
          }
        } else {
          // Platz behalten
          if (obKeepVariant === 'active') {
            await eventServiceRef.resolveOverbookKeepActive(sub, reg.Id);
          } else {
            await eventServiceRef.resolveOverbookKeepAsFirstWaitlist(sub, reg.Id, grp);
          }
        }
      }
      // IDs neu vergeben (Aktive 1..N, Warteliste N+1..) + Counter + Seat-Sync.
      // Mit Fortschritts-Overlay, damit man bei großen Listen sieht wie weit.
      setReorderProgressLabel('IDs werden neu vergeben…');
      setReorderProgress(0);
      try { await eventServiceRef.reorderParticipantIDs(sub, pct => setReorderProgress(pct)); } catch { /* */ }
      try { await eventServiceRef.syncSeatsToActiveCount(sub, { isSplit: isSplitCapacity }); } catch { /* */ }
      setReorderProgress(null);
      // v30.67 (Review): gemeinsamer Nachlade-Pfad statt `setRegistrations([])` bei 429.
      await reloadRegistrations();
    } catch { /* einzelne Fehler werden geschluckt; Liste wird trotzdem neu geladen */ }
    setObBusy(false);
    setOverbookModal(null);
  };

  // v11.36: TeilnehmerIDs neu vergeben — gemeinsam von der Toolbox-Kachel
  // UND dem Hinweis-Modal genutzt (mit %-Fortschritts-Overlay).
  const runIdReorder = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsReorderingIDs(true);
    setReorderResult(null);
    setReorderProgressLabel(isDe ? 'IDs werden neu vergeben…' : 'Reassigning IDs…');
    setReorderProgress(0);
    try {
      const result = await eventServiceRef.reorderParticipantIDs(
        selectedEvent.subsiteUrl,
        pct => setReorderProgress(pct)
      );
      setReorderResult(isDe
        ? `${result.success} aktualisiert, ${result.errors} Fehler`
        : `${result.success} updated, ${result.errors} errors`);
      await reloadRegistrations();
    } catch {
      setReorderResult(isDe ? 'Fehler beim Neuvergeben der IDs' : 'Error reassigning IDs');
    }
    setReorderProgress(null);
    setIsReorderingIDs(false);
  };

  /**
   * v29.16: Nachrück-Mail + Outlook-Einladung für EINE nachgerückte Person.
   * Aus `runManualPromote` herausgezogen, weil das Füllen mehrerer freier
   * Plätze dieselbe Benachrichtigung je Person braucht — zwei Kopien dieses
   * Mail-Aufbaus wären beim nächsten Template-Wechsel auseinandergelaufen.
   */
  // v30.69: `ev` ist optional und fällt auf `selectedEvent` zurück. Die
  // Sammel-Aktion arbeitet Event für Event ab, ohne die Auswahl umzustellen —
  // ein `setSelectedEvent` je Event würde die ganze Seite neu rendern und die
  // Reihenfolge der Zustandsänderungen unübersichtlich machen.
  const notifyPromoted = async (promoted: { email: string; name?: string }, ev?: DeloitteEvent): Promise<void> => {
    const target = ev || selectedEvent;
    if (!eventServiceRef || !target) return;
    if (!target.disableEmails) {
      try {
        const lang = target.emailLanguage || 'EN';
        const promotedFirstName = (promoted.name || '').trim().split(/\s+/)[0] || '';
        const promoteVars = {
          Name: promotedFirstName,
          EventTitle: target.title,
          Organizer: formatOrganizerList(target.organizers, lang),
          AppUrl: `${eventServiceRef.siteUrl}/SitePages/DEX.aspx?env=WebView`,
          WaitlistPosition: '',
        };
        let emailData: { subject: string; body: string };
        const spTplRaw = await eventServiceRef.getEmailTemplate('Nachruecken', lang).catch(() => null);
        const spTpl = applyEventTemplateOverride(spTplRaw, target.emailTemplateOverrides, 'Nachruecken');
        if (spTpl) {
          emailData = buildEmailFromTemplate(spTpl, promoteVars);
        } else {
          emailData = promotionEmail(promotedFirstName, target.title);
        }
        await eventServiceRef.queueEmail(
          withParentTitleSubject(emailData.subject, target.parentEventId ? allEvents.find(e => e.id === target.parentEventId) : undefined),
          promoted.email, promoted.name || '', emailData.body,
          'Nachruecken', target.title, target.id
        );
      } catch (err) { console.warn('[DEX] promote-email failed:', err); }
    }
    if (!target.disableOutlook) {
      try {
        await eventServiceRef.queueOutlookEvent(
          promoted.email, target.id, target.title, 'Einladen'
        );
      } catch (err) { console.warn('[DEX] promote-outlook failed:', err); }
    }
  };

  /**
   * v18.70 / v29.16: Freie Plätze mit der Warteliste füllen.
   *
   * v18.70 rückte GENAU EINE Person nach und übergab weder Gruppe noch die
   * Gruppen-Kapazität. Auf einem Event mit zwei Gruppen ging das doppelt
   * daneben: Als Obergrenze stand `maxParticipants` drin — das ist bei
   * geteilten Kapazitäten 0, also fand gar keine Prüfung statt — und ohne
   * Typfilter nahm die Abfrage den ersten Wartelistler nach TeilnehmerID,
   * egal aus welcher Gruppe. Eine noch volle Gruppe konnte damit überbucht
   * werden, während die Gruppe mit freien Plätzen leer ausging.
   *
   * Zweiter Punkt, der den Fall überhaupt erst sichtbar macht: Nachgerückt
   * wird sonst NUR beim Abmelden. Erhöht der Organizer eine Gruppengröße,
   * passiert von allein nichts — es gibt kein Ereignis, an dem etwas hinge.
   * Genau dafür ist diese Aktion da, und sie füllt deshalb ALLE freien
   * Plätze auf einmal statt einen pro Klick.
   *
   * Gezählt wird je Gruppe (bzw. einmal gesamt ohne geteilte Kapazität) mit
   * derselben Zuordnung wie überall sonst: StarterType, ersatzweise
   * PreferredStarterType. Bei `splitSharedWaitlist` gibt es nur eine
   * Warteliste — dann entscheidet die Reihenfolge, nicht die Gruppe.
   */
  const runManualPromote = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
    setIsPromoting(true);
    setPromoteResult(null);
    try {
      // Frisch lesen: `registrations` kann Minuten alt sein, und wir leiten
      // daraus ab, wie oft nachgerückt wird. Auf einem veralteten Stand
      // würde die Schleife über die Kapazität hinauslaufen.
      // v30.67 (Review): Nicht lesbar heißt hier nicht „0 Aktive" — genau
      // daraus leitet die Schleife ab, wie viele Plätze frei sind; ein 429
      // hätte die Warteliste in ein volles Event nachrücken lassen.
      let freshFailed = false;
      const fresh = await getAllRegistrations(selectedEvent.id, () => { freshFailed = true; });
      if (freshFailed) {
        setPromoteResult(isDe
          ? 'Nachrücken abgebrochen: Die Teilnehmerliste konnte gerade nicht gelesen werden — bitte später erneut versuchen.'
          : 'Promotion aborted: the attendee list could not be read right now — please try again later.');
        setIsPromoting(false);
        return;
      }
      const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
      const lblA = (selectedEvent.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
      const lblB = (selectedEvent.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
      // Geteilte Kapazität MIT getrennten Wartelisten → je Gruppe rechnen.
      // Gemeinsame Warteliste (splitSharedWaitlist) verhält sich wie ein
      // einzelner Topf: Wer zuerst wartet, rückt nach.
      const perGroup = isSplitCapacity && !selectedEvent.splitSharedWaitlist;
      const groups: Array<{ key?: string; label: string; cap: number }> = perGroup
        ? [
          { key: 'Durchstarter', label: lblA, cap: selectedEvent.durchstarterCapacity || 0 },
          { key: 'Funstarter', label: lblB, cap: selectedEvent.funstarterCapacity || 0 },
        ]
        : [{
          key: undefined,
          label: isDe ? 'Plätze' : 'Seats',
          cap: isSplitCapacity
            // Gemeinsame Warteliste: Die Obergrenze ist die Summe beider
            // Gruppen — `maxParticipants` ist bei geteilten Kapazitäten 0.
            ? (selectedEvent.durchstarterCapacity || 0) + (selectedEvent.funstarterCapacity || 0)
            : (selectedEvent.maxParticipants || 0),
        }];
      const groupOf = (r: SPRegistration): string => r.StarterType || r.PreferredStarterType || '';
      // v30.67: Gemeinsame Warteliste heißt „gemeinsame REIHENFOLGE", nicht
      // „gemeinsamer Kapazitätstopf". Die nachrückende Person bekommt ihren
      // PreferredStarterType als StarterType, und die Gruppengrenzen gelten
      // weiter — `reserveSeat`, `detectOverbooking` und die Gruppen-Karten
      // der Anmeldeseite prüfen alle gegen die EINZELNE Gruppe. Bisher wurde
      // hier gegen die Summe gerechnet und ohne Typfilter nachgerückt: Bei A
      // voll (30/30) und B 25/30 rückten fünf Durchstarter nach, A stand auf
      // 35/30, die Überbuchungs-Box flaggte sie, B blieb leer.
      //
      // Deshalb die Warteliste in ihrer Reihenfolge simulieren: Solange BEIDE
      // Gruppen Platz haben, rückt die nächste Person nach (ohne Filter — das
      // ist genau sie, weil bis dahin niemand übersprungen wurde). Ist eine
      // Gruppe voll, nur noch mit Typfilter auf die offene Gruppe; wer dort
      // nicht hineingehört, bleibt stehen. Ohne Wunsch-Typ zählt die Person
      // auf den Gesamt-Topf und kann nur in Phase 1 nachrücken.
      const sharedInfo = ((): { steps: Array<string | undefined>; activeA: number; activeB: number; capA: number; capB: number } | null => {
        if (!isSplitCapacity || !selectedEvent.splitSharedWaitlist) return null;
        const capA = selectedEvent.durchstarterCapacity || 0;
        const capB = selectedEvent.funstarterCapacity || 0;
        const activeAll = fresh.filter(r => ACTIVE.indexOf(r.Status) >= 0);
        const activeA = activeAll.filter(r => groupOf(r) === 'Durchstarter').length;
        const activeB = activeAll.filter(r => groupOf(r) === 'Funstarter').length;
        let freeA = capA > 0 ? Math.max(0, capA - activeA) : Number.POSITIVE_INFINITY;
        let freeB = capB > 0 ? Math.max(0, capB - activeB) : Number.POSITIVE_INFINITY;
        // Gesamt-Topf als zweite Schranke: Aktive ohne Gruppe belegen einen
        // Platz, ohne in einer Gruppe zu zählen.
        let budget = (capA + capB) > 0 ? Math.max(0, capA + capB - activeAll.length) : Number.POSITIVE_INFINITY;
        const steps: Array<string | undefined> = [];
        const waitingSorted = fresh.filter(r => r.Status === 'Warteliste')
          .sort((a, b) => (a.TeilnehmerID || 0) - (b.TeilnehmerID || 0));
        for (const r of waitingSorted) {
          if (budget <= 0 || (freeA <= 0 && freeB <= 0)) break;
          const g = groupOf(r);
          if (freeA > 0 && freeB > 0) {
            steps.push(undefined);
            if (g === 'Durchstarter') freeA -= 1;
            else if (g === 'Funstarter') freeB -= 1;
            else if (freeA >= freeB) freeA -= 1;
            else freeB -= 1;
            budget -= 1;
          } else if (freeA > 0 && g === 'Durchstarter') {
            steps.push('Durchstarter'); freeA -= 1; budget -= 1;
          } else if (freeB > 0 && g === 'Funstarter') {
            steps.push('Funstarter'); freeB -= 1; budget -= 1;
          }
          // sonst: Gruppe voll → bleibt stehen
        }
        return { steps, activeA, activeB, capA, capB };
      })();
      const plan = groups.map(g => {
        const inGroup = (r: SPRegistration): boolean => !g.key || groupOf(r) === g.key;
        const active = fresh.filter(r => ACTIVE.indexOf(r.Status) >= 0 && inGroup(r)).length;
        const waiting = fresh.filter(r => r.Status === 'Warteliste' && inGroup(r)).length;
        // Kapazität 0 heißt „unbegrenzt" — dann rückt die ganze Warteliste nach.
        const free = g.cap > 0 ? Math.max(0, g.cap - active) : waiting;
        return { ...g, active, waiting, count: sharedInfo ? sharedInfo.steps.length : Math.min(free, waiting), free };
      });
      const total = plan.reduce((n, g) => n + g.count, 0);
      const sharedGroupsText = sharedInfo
        ? ` (${lblA} ${sharedInfo.activeA}/${sharedInfo.capA || '∞'} · ${lblB} ${sharedInfo.activeB}/${sharedInfo.capB || '∞'})`
        : '';

      if (total === 0) {
        // Den tatsächlichen Grund nennen, nicht den erstbesten: „voll" und
        // „niemand wartet" führen zu ganz verschiedenen nächsten Schritten.
        const anyWaiting = plan.some(g => g.waiting > 0);
        setPromoteResult(!anyWaiting
          ? (isDe ? 'Niemand auf der Warteliste.' : 'Nobody on the waitlist.')
          : perGroup
            ? (isDe
              ? `Kein freier Platz in den Gruppen, in denen jemand wartet (${plan.filter(g => g.waiting > 0).map(g => `${g.label}: ${g.active}/${g.cap}`).join(', ')}).`
              : `No free seat in the groups where people are waiting (${plan.filter(g => g.waiting > 0).map(g => `${g.label}: ${g.active}/${g.cap}`).join(', ')}).`)
            : sharedInfo
              // v30.67: gemeinsame Warteliste — die wartenden Personen gehören
              // in eine volle Gruppe, die andere hat Platz, aber niemand will hin.
              ? (isDe
                ? `Kein freier Platz in der Gruppe, in der jemand wartet${sharedGroupsText}.`
                : `No free seat in the group where people are waiting${sharedGroupsText}.`)
              : (isDe ? 'Kein freier Platz — Event ist voll.' : 'No free seat — event is full.'));
        setIsPromoting(false);
        return;
      }

      const lines = plan
        .filter(g => g.waiting > 0 || g.free > 0)
        .map(g => perGroup
          ? `• ${g.label}: ${g.active}/${g.cap || '∞'} belegt · ${g.waiting} auf der Warteliste → ${g.count} rücken nach`
          : `• ${g.active}/${g.cap || '∞'} belegt${sharedGroupsText} · ${g.waiting} auf der Warteliste → ${g.count} rücken nach`)
        .join('\n');
      const ok = await confirmDialog(
        isDe
          ? `${total} ${total === 1 ? 'Person' : 'Personen'} von der Warteliste nachrücken lassen?\n\n${lines}\n\nJede nachgerückte Person bekommt den Status „Angemeldet", eine Nachrück-Mail und eine Outlook-Einladung. Danach werden die TeilnehmerIDs neu vergeben.`
          : `Move ${total} ${total === 1 ? 'person' : 'people'} up from the waitlist?\n\n${lines}\n\nEach of them gets status “Registered”, a promotion email and an Outlook invite. Participant IDs are reassigned afterwards.`,
        { confirmLabel: isDe ? 'Nachrücken' : 'Promote' },
      );
      if (!ok) { setIsPromoting(false); return; }

      const promotedNames: string[] = [];
      let failed = 0;
      // v30.67: Bei gemeinsamer Warteliste die Schrittfolge aus der
      // Simulation oben (je Schritt ein Typfilter oder keiner), sonst wie
      // bisher je Gruppe `count`-mal.
      const runs: Array<{ key?: string; count: number }> = sharedInfo
        ? sharedInfo.steps.map(k => ({ key: k, count: 1 }))
        : plan.map(g => ({ key: g.key, count: g.count }));
      let stop = false;
      for (const g of runs) {
        if (stop) break;
        for (let i = 0; i < g.count; i++) {
          // maxParticipants bewusst NICHT mitgeben: Die Überbuchungs-Sperre
          // im Service zählt über die GANZE Liste und kann eine Gruppe nicht
          // getrennt prüfen. Die Anzahl steht oben schon fest, frisch
          // gerechnet — hier wird nur genau so oft nachgerückt.
          const promoted = await eventServiceRef.promoteFirstWaitlistItem(
            selectedEvent.subsiteUrl,
            undefined,
            undefined,
            g.key,
          );
          if (promoted && promoted.success && promoted.email) {
            promotedNames.push(promoted.name || promoted.email);
            setAdminToast({ kind: 'promoted', name: promoted.name || promoted.email, email: promoted.email });
            await notifyPromoted({ email: promoted.email, name: promoted.name });
          } else {
            // Warteliste unerwartet leer (jemand hat sich zwischendurch
            // abgemeldet) — kein Fehler, nur nichts mehr zu tun. Bei der
            // simulierten Schrittfolge stimmt ab hier die Reihenfolge nicht
            // mehr — dann ganz aufhören statt weiterzuraten.
            failed++;
            if (sharedInfo) stop = true;
            break;
          }
        }
      }

      // IDs neu vergeben + Counter/Seat-Sync + Liste neu laden — einmal am
      // Ende, nicht je Person: Der Reorder schreibt jede Zeile an.
      await runIdReorder();
      try {
        await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit: isSplitCapacity });
      } catch { /* */ }

      const n = promotedNames.length;
      setPromoteResult(n === 0
        ? (isDe ? 'Es konnte niemand nachrücken.' : 'Nobody could be promoted.')
        : n === 1
          ? (isDe ? `${promotedNames[0]} ist nachgerückt.` : `${promotedNames[0]} moved up.`)
          : (isDe
            ? `${n} Personen sind nachgerückt: ${promotedNames.join(', ')}.`
            : `${n} people moved up: ${promotedNames.join(', ')}.`));
      if (failed > 0 && n > 0) {
        console.warn('[DEX] runManualPromote: Warteliste war früher leer als erwartet.');
      }
    } catch (err) {
      console.warn('[DEX] runManualPromote failed:', err);
      setPromoteResult(isDe ? 'Fehler beim Nachrücken.' : 'Error promoting.');
    }
    setIsPromoting(false);
  };

  // v11.70 / v11.71: Hinweis-Box „IDs sind ggf. nicht korrekt" wird jetzt
  // an die tatsächliche TeilnehmerID-Sequenz gekoppelt — nicht mehr an
  // eine 10-Minuten-Zeit-Heuristik nach der letzten Abmeldung.
  //
  // Erwartet: alle nicht-abgemeldeten Einträge (Status in
  // Angemeldet/QR versendet/Eingecheckt/Warteliste) haben TeilnehmerIDs,
  // die nach Sortierung lückenlos 1..N durchlaufen. Sobald
  //   - eine ID fehlt (Lücke),
  //   - eine ID doppelt vorkommt,
  //   - ein nicht-abgemeldeter Eintrag keine (oder ≤0) ID hat,
  // ist der Zustand „IDs evtl. nicht korrekt". Typischer Trigger: gerade
  // erfolgte Abmeldung, der DEX_IDReorder-Flow ist noch nicht fertig.
  // Das gibt einen ehrlichen Status — die Box verschwindet automatisch,
  // sobald der Flow durch ist (statt nach willkürlichen 10 Minuten).
  const recentCancellation = (regs: SPRegistration[]): { recent: boolean; whenIso: string; detail: string } => {
    const active = regs.filter(r => r.Status !== 'Abgemeldet');
    if (active.length === 0) return { recent: false, whenIso: '', detail: '' };
    const ids: number[] = [];
    let noId = 0;
    for (const r of active) {
      const id = Number(r.TeilnehmerID);
      if (!isFinite(id) || id <= 0) { noId++; continue; }
      ids.push(id);
    }
    ids.sort((a, b) => a - b);
    // v22.12: konkrete Diagnose statt nur ja/nein — erste Lücke + Duplikate
    // zählen, damit die Box belegt, WAS in den geladenen Daten falsch ist.
    let dups = 0;
    let firstGapAt = 0;
    for (let i = 0; i < ids.length; i++) {
      if (i > 0 && ids[i] === ids[i - 1]) dups++;
      if (firstGapAt === 0 && ids[i] !== i + 1) firstGapAt = i + 1;
    }
    if (noId === 0 && dups === 0 && firstGapAt === 0) return { recent: false, whenIso: '', detail: '' };
    const parts: string[] = [];
    if (firstGapAt > 0) parts.push(`Nummern nicht durchgängig (erwartet Nr. ${firstGapAt})`);
    if (dups > 0) parts.push(`${dups} doppelte Nummer${dups === 1 ? '' : 'n'}`);
    if (noId > 0) parts.push(`${noId} Eintr${noId === 1 ? 'ag' : 'äge'} ohne Nummer`);
    return {
      recent: true,
      whenIso: latestCancelIso(regs),
      detail: `${active.length} aktive Einträge — ${parts.join(', ')}`,
    };
  };
  // Hilfsfunktion: jüngste CancellationDate aus der Liste (für den
  // optionalen Zeit-Hinweis in der Box).
  const latestCancelIso = (regs: SPRegistration[]): string => {
    let latest = 0;
    for (const r of regs) {
      if (r.Status !== 'Abgemeldet') continue;
      const t = new Date(r.CancellationDate || '').getTime();
      if (!isNaN(t) && t > latest) latest = t;
    }
    return latest > 0 ? new Date(latest).toISOString() : '';
  };
  /**
   * v30.69: Nachrücken & IDs für ALLE Events nachholen (Admin Center).
   *
   * Entstanden am 02.09.2026: Der Flow `DEX_IDReorder_TeilnehmerIDs` brach
   * einen Tag lang jeden Lauf mit 502 ab (Gruppen-Zählung auf Events ohne
   * `StarterType`-Spalte). Jede Abmeldung in dieser Zeit hat einen Platz frei
   * gemacht, den niemand bekommen hat — bei Leuten daneben auf der Warteliste.
   * Pro Event nachklicken hieße 30 Events × „Freie Plätze füllen"; das hier
   * macht es in einem Durchgang und berichtet am Ende.
   *
   * Zwei Phasen, bewusst getrennt:
   *  1. PLANEN — alle Listen lesen, je Event über `buildPromotionPlan`
   *     ausrechnen, wer nachrücken darf, und prüfen, ob die Nummern Lücken
   *     haben. Noch wird nichts geschrieben. Das Ergebnis steht im
   *     Bestätigungs-Dialog: Die Aktion verschickt Nachrück-Mails und
   *     Outlook-Einladungen, und ein Knopf, der ungefragt an 30 Leute mailt,
   *     ist der falsche Knopf.
   *  2. AUSFÜHREN — promoten + benachrichtigen, IDs neu vergeben, Zähler
   *     syncen. Sequentiell (SharePoint-Throttling), mit Fehlerzähler statt
   *     stillem catch — ein 429 hinterließe sonst genau die halben Zustände,
   *     die man später nicht mehr nachrechnen kann (v29.2).
   *
   * Beschnittene Listen werden ÜBERSPRUNGEN, nicht als leer gerechnet: Eine
   * Subsite ohne Vollzugriff meldet sich über `onHttpError`, und „0 Aktive"
   * aus einer 403-Sicht hieße sonst „alle Plätze frei" (v30.62). Ebenso
   * übersprungen: geteilte Kapazität mit GEMEINSAMER Warteliste — dort
   * entscheidet die Reihenfolge, nicht die Gruppe (v30.67), und das rechnet
   * nur `runManualPromote` richtig. Beide Gruppen stehen namentlich im Bericht.
   */
  const runHealAllEvents = async (): Promise<void> => {
    if (!eventServiceRef) return;
    setIsHealingAll(true);
    setHealAllResult(null);
    setHealAllProgress(null);
    try {
      // Alle aktiven Events mit Subsite — auch Sub-Events, jeder Termin hat
      // eigene Liste, eigenen Zähler, eigene Warteliste. Klammern im
      // subEventsOnlyMode raus: dort ist niemand buchbar, die Zeilen sind
      // Schatten (v15.25), und `cap 0` würde als „unbegrenzt" gelesen.
      const seen = new Set<string>();
      const candidates = allEvents.filter(e => {
        const sub = (e.subsiteUrl || '').trim();
        if (!sub || e.status !== 'Active' || e.subEventsOnlyMode) return false;
        if (seen.has(sub)) return false;
        seen.add(sub);
        return true;
      });
      if (candidates.length === 0) {
        setHealAllResult(isDe ? 'Keine aktiven Events mit Teilnehmerliste.' : 'No active events with a participant list.');
        setIsHealingAll(false);
        return;
      }

      // ---- Phase 1: planen -------------------------------------------------
      type Planned = { ev: DeloitteEvent; plan: PromotionPlan; idGap: boolean; blocked: number | null };
      const planned: Planned[] = [];
      for (let i = 0; i < candidates.length; i++) {
        const ev = candidates[i];
        setHealAllProgress(isDe ? `Prüfe ${i + 1}/${candidates.length}: ${ev.title}` : `Checking ${i + 1}/${candidates.length}: ${ev.title}`);
        // Objekt statt `let`, weil TS eine Zuweisung im Rückruf nicht sieht
        // und die Variable sonst als `null` festschreibt.
        const err: { status: number | null } = { status: null };
        const regs = await getAllRegistrations(ev.id, s => { err.status = s; });
        planned.push({
          ev,
          plan: buildPromotionPlan(ev, regs, isDe),
          idGap: recentCancellation(regs).recent,
          blocked: err.status,
        });
      }
      const blocked = planned.filter(p => p.blocked !== null);
      const shared = planned.filter(p => p.blocked === null && p.plan.sharedWaitlist && p.plan.anyWaiting);
      const usable = planned.filter(p => p.blocked === null && !p.plan.sharedWaitlist);
      const withPromote = usable.filter(p => p.plan.total > 0);
      const withGap = usable.filter(p => p.idGap);
      const totalPromote = withPromote.reduce((n, p) => n + p.plan.total, 0);

      const evLabel = (ev: DeloitteEvent): string => {
        const parent = ev.parentEventId ? allEvents.find(e => e.id === ev.parentEventId) : undefined;
        return parent ? `${parent.title} › ${shortSubEventTitle(ev.title, parent.title)}` : ev.title;
      };
      const summaryLines: string[] = [];
      summaryLines.push(isDe ? `${usable.length} ${usable.length === 1 ? 'Event' : 'Events'} geprüft.` : `${usable.length} ${usable.length === 1 ? 'event' : 'events'} checked.`);
      if (totalPromote > 0) {
        summaryLines.push('');
        summaryLines.push(isDe
          ? `${totalPromote} ${totalPromote === 1 ? 'Person rückt' : 'Personen rücken'} nach:`
          : `${totalPromote} ${totalPromote === 1 ? 'person moves' : 'people move'} up:`);
        for (const p of withPromote) {
          summaryLines.push(`• ${evLabel(p.ev)} — ${p.plan.total}`);
          for (const l of promotionPlanLines(p.plan)) summaryLines.push(`    ${l}`);
        }
      } else {
        summaryLines.push(isDe ? 'Niemand muss nachrücken.' : 'Nobody needs to move up.');
      }
      summaryLines.push('');
      summaryLines.push(isDe
        ? `${withGap.length} ${withGap.length === 1 ? 'Event' : 'Events'} mit Nummern-Lücken ${withGap.length === 1 ? 'wird' : 'werden'} neu nummeriert; die Platzzähler aller ${usable.length} Events werden abgeglichen.`
        : `${withGap.length} ${withGap.length === 1 ? 'event' : 'events'} with ID gaps will be renumbered; the seat counters of all ${usable.length} events will be reconciled.`);
      if (shared.length > 0) {
        summaryLines.push('');
        summaryLines.push(isDe
          ? `Nicht automatisch — gemeinsame Warteliste bei geteilten Gruppen, bitte je Event über „Freie Plätze mit Warteliste füllen" (${shared.length}):`
          : `Not automatic — shared waitlist with split groups, please use “Fill free seats from waitlist” per event (${shared.length}):`);
        for (const p of shared) summaryLines.push(`• ${evLabel(p.ev)}`);
      }
      if (blocked.length > 0) {
        summaryLines.push('');
        summaryLines.push(isDe
          ? `Übersprungen — kein Vollzugriff auf die Teilnehmerliste (${blocked.length}):`
          : `Skipped — no full access to the participant list (${blocked.length}):`);
        for (const p of blocked) summaryLines.push(`• ${evLabel(p.ev)} (HTTP ${p.blocked})`);
      }
      if (totalPromote > 0) {
        summaryLines.push('');
        summaryLines.push(isDe
          ? 'Jede nachgerückte Person bekommt den Status „Angemeldet", eine Nachrück-Mail und eine Outlook-Einladung.'
          : 'Each promoted person gets status “Registered”, a promotion email and an Outlook invite.');
      }
      setHealAllProgress(null);
      const ok = await confirmDialog(summaryLines.join('\n'), {
        confirmLabel: totalPromote > 0 ? (isDe ? 'Nachrücken & heilen' : 'Promote & heal') : (isDe ? 'Heilen' : 'Heal'),
      });
      if (!ok) { setIsHealingAll(false); return; }

      // ---- Phase 2: ausführen ---------------------------------------------
      let promotedTotal = 0;
      let reorderedEvents = 0;
      let syncedEvents = 0;
      let errors = 0;
      const promotedNames: string[] = [];
      for (let i = 0; i < usable.length; i++) {
        const { ev, plan, idGap } = usable[i];
        const sub = (ev.subsiteUrl || '').trim();
        setHealAllProgress(isDe ? `Heile ${i + 1}/${usable.length}: ${ev.title}` : `Healing ${i + 1}/${usable.length}: ${ev.title}`);
        let promotedHere = 0;
        for (const g of plan.groups) {
          for (let k = 0; k < g.count; k++) {
            // Obergrenze bewusst NICHT mitgeben — dieselbe Begründung wie in
            // runManualPromote: Der Service zählt über die ganze Liste und
            // kann eine Gruppe nicht trennen; die Anzahl steht oben fest.
            const promoted = await eventServiceRef.promoteFirstWaitlistItem(sub, undefined, undefined, g.key);
            if (promoted && promoted.success && promoted.email) {
              promotedHere++;
              promotedNames.push(promoted.name || promoted.email);
              await notifyPromoted({ email: promoted.email, name: promoted.name }, ev);
            } else {
              // Warteliste unerwartet leer (jemand hat sich zwischendurch
              // abgemeldet) — kein Fehler, nur nichts mehr zu tun.
              break;
            }
          }
        }
        promotedTotal += promotedHere;
        // IDs nur anfassen, wo es nötig ist: Der Reorder schreibt jede
        // geänderte Zeile, und auf 30 sauberen Events wäre das reines Rauschen
        // gegen das SharePoint-Throttling.
        if (idGap || promotedHere > 0) {
          try {
            const r = await eventServiceRef.reorderParticipantIDs(sub, () => { /* Fortschritt je Event nicht nötig */ });
            reorderedEvents++;
            errors += r.errors;
          } catch (e) {
            errors++;
            console.warn('[DEX] healAll: reorder failed', ev.title, e);
          }
        }
        try {
          await eventServiceRef.syncSeatsToActiveCount(sub, { isSplit: isSplitCapacityOf(ev) });
          syncedEvents++;
        } catch (e) {
          errors++;
          console.warn('[DEX] healAll: seat sync failed', ev.title, e);
        }
      }

      const parts: string[] = [];
      parts.push(isDe
        ? `${promotedTotal} ${promotedTotal === 1 ? 'Person' : 'Personen'} nachgerückt`
        : `${promotedTotal} ${promotedTotal === 1 ? 'person' : 'people'} promoted`);
      if (promotedNames.length > 0) parts[0] += ` (${promotedNames.slice(0, 6).join(', ')}${promotedNames.length > 6 ? ', …' : ''})`;
      parts.push(isDe ? `${reorderedEvents} Events neu nummeriert` : `${reorderedEvents} events renumbered`);
      parts.push(isDe ? `${syncedEvents} Zähler abgeglichen` : `${syncedEvents} counters reconciled`);
      if (shared.length > 0) parts.push(isDe ? `${shared.length} mit gemeinsamer Warteliste — bitte einzeln` : `${shared.length} with shared waitlist — please handle individually`);
      if (blocked.length > 0) parts.push(isDe ? `${blocked.length} übersprungen (kein Zugriff)` : `${blocked.length} skipped (no access)`);
      if (errors > 0) parts.push(isDe ? `${errors} Fehler — Konsole prüfen` : `${errors} errors — check console`);
      setHealAllResult(parts.join(' · '));

      // Das gerade geöffnete Event zeigt sonst bis zum nächsten Klick den
      // alten Stand („Teilnehmer (N)", Warn-Box).
      if (selectedEvent && seen.has((selectedEvent.subsiteUrl || '').trim())) {
        try { await reloadRegistrations(); } catch { /* Anzeige bleibt alt, Daten sind geheilt */ }
      }
    } catch (err) {
      console.warn('[DEX] healAll failed:', err);
      setHealAllResult(isDe ? `Fehler: ${err instanceof Error ? err.message : String(err)}` : `Error: ${err instanceof Error ? err.message : String(err)}`);
    }
    setHealAllProgress(null);
    setIsHealingAll(false);
  };

  const idFixCheckedForRef = React.useRef<string | null>(null);
  // v22.12: solange die geladenen Daten kaputte IDs zeigen, lädt die App die
  // Teilnehmerliste automatisch alle 30 Sekunden neu — sobald der
  // DEX_IDReorder-Flow durch ist, verschwindet die Warn-Box von selbst
  // (vorher musste man manuell „Aktualisieren" klicken und hielt den
  // durchgelaufenen Flow fälschlich für kaputt).
  const idRecheckBusyRef = React.useRef(false);
  const [idRecheckBusy, setIdRecheckBusy] = React.useState(false);
  const reloadRegistrationsForIdCheck = React.useCallback(async (): Promise<void> => {
    if (!selectedEvent || idRecheckBusyRef.current) return;
    idRecheckBusyRef.current = true;
    setIdRecheckBusy(true);
    try {
      await reloadRegistrations();
    } catch { /* best-effort */ }
    idRecheckBusyRef.current = false;
    setIdRecheckBusy(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id]);
  return {
    idFixCheckedForRef, idRecheckBusy, recentCancellation, reloadRegistrationsForIdCheck,
    runIdReorder, runManualPromote, runOverbookResolution, runHealAllEvents,
  };
}

