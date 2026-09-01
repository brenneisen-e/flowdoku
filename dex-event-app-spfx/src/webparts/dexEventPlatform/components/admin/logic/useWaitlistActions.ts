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

export interface UseWaitlistActionsCtx {
  allEvents: DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  eventServiceRef: EventService;
  getAllRegistrations: (eventId: string, onHttpError?: (_status: number) => void) => Promise<SPRegistration[]>;
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
  setRegistrations: React.Dispatch<React.SetStateAction<SPRegistration[]>>;
  setReorderProgress: React.Dispatch<React.SetStateAction<number>>;
  setReorderProgressLabel: React.Dispatch<React.SetStateAction<string>>;
  setReorderResult: React.Dispatch<React.SetStateAction<string>>;
}

export interface UseWaitlistActionsResult {
  idFixCheckedForRef: React.MutableRefObject<string>;
  idRecheckBusy: boolean;
  recentCancellation: (regs: SPRegistration[]) => {    recent: boolean;    whenIso: string;    detail: string;};
  reloadRegistrationsForIdCheck: () => Promise<void>;
  runIdReorder: () => Promise<void>;
  runManualPromote: () => Promise<void>;
  runOverbookResolution: () => Promise<void>;
}

export function useWaitlistActions(ctx: UseWaitlistActionsCtx): UseWaitlistActionsResult {
  const {
    allEvents, confirmDialog, eventServiceRef, getAllRegistrations, isDe, isSplitCapacity,
    obKeepVariant, obMailBody, obMailLang, obMailSubject, obRemoveCalendar, obWithMail,
    overbookModal, registrations, selectedEvent, setAdminToast, setIsPromoting, setIsReorderingIDs,
    setObBusy, setObMailBody, setObMailLang, setObMailSubject, setOverbookModal, setPromoteResult,
    setRegistrations, setReorderProgress, setReorderProgressLabel, setReorderResult,
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
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
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
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
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
  const notifyPromoted = async (promoted: { email: string; name?: string }): Promise<void> => {
    if (!eventServiceRef || !selectedEvent) return;
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
        if (spTpl) {
          emailData = buildEmailFromTemplate(spTpl, promoteVars);
        } else {
          emailData = promotionEmail(promotedFirstName, selectedEvent.title);
        }
        await eventServiceRef.queueEmail(
          withParentTitleSubject(emailData.subject, selectedEvent.parentEventId ? allEvents.find(e => e.id === selectedEvent.parentEventId) : undefined),
          promoted.email, promoted.name || '', emailData.body,
          'Nachruecken', selectedEvent.title, selectedEvent.id
        );
      } catch (err) { console.warn('[DEX] promote-email failed:', err); }
    }
    if (!selectedEvent.disableOutlook) {
      try {
        await eventServiceRef.queueOutlookEvent(
          promoted.email, selectedEvent.id, selectedEvent.title, 'Einladen'
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
      const fresh = await getAllRegistrations(selectedEvent.id);
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
      const plan = groups.map(g => {
        const inGroup = (r: SPRegistration): boolean => !g.key || groupOf(r) === g.key;
        const active = fresh.filter(r => ACTIVE.indexOf(r.Status) >= 0 && inGroup(r)).length;
        const waiting = fresh.filter(r => r.Status === 'Warteliste' && inGroup(r)).length;
        // Kapazität 0 heißt „unbegrenzt" — dann rückt die ganze Warteliste nach.
        const free = g.cap > 0 ? Math.max(0, g.cap - active) : waiting;
        return { ...g, active, waiting, count: Math.min(free, waiting), free };
      });
      const total = plan.reduce((n, g) => n + g.count, 0);

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
            : (isDe ? 'Kein freier Platz — Event ist voll.' : 'No free seat — event is full.'));
        setIsPromoting(false);
        return;
      }

      const lines = plan
        .filter(g => g.waiting > 0 || g.free > 0)
        .map(g => perGroup
          ? `• ${g.label}: ${g.active}/${g.cap || '∞'} belegt · ${g.waiting} auf der Warteliste → ${g.count} rücken nach`
          : `• ${g.active}/${g.cap || '∞'} belegt · ${g.waiting} auf der Warteliste → ${g.count} rücken nach`)
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
      for (const g of plan) {
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
            // abgemeldet) — kein Fehler, nur nichts mehr zu tun.
            failed++;
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
      const regs = await getAllRegistrations(selectedEvent.id);
      setRegistrations(regs);
    } catch { /* best-effort */ }
    idRecheckBusyRef.current = false;
    setIdRecheckBusy(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedEvent?.id]);
  return {
    idFixCheckedForRef, idRecheckBusy, recentCancellation, reloadRegistrationsForIdCheck,
    runIdReorder, runManualPromote, runOverbookResolution,
  };
}

