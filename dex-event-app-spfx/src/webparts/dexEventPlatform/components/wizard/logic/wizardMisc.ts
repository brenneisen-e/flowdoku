import * as React from 'react';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { SUB_TRANSFER_GROUPS } from '../../../data/wizardHints';
import { compressImage } from '../../../utils/imageCompress';

/* applyEventPhotoToLogo — aus EventCreationPage.tsx ausgelagert (Zeilen 1271-1312 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface ApplyEventPhotoToLogoCtx {
  editEvent: import("../../../types/index").DeloitteEvent;
  fileToBase64: (file: File) => Promise<string>;
  imageFile: File;
  imageOrigFile: File;
  imagePreview: string;
  isDe: boolean;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  shrinkLogoB64: (b64: string) => Promise<string>;
}

export async function applyEventPhotoToLogoImpl(ctx: ApplyEventPhotoToLogoCtx, setter: (b64: string) => void): Promise<string> {
  const { editEvent, fileToBase64, imageFile, imageOrigFile, imagePreview, isDe, showAlert, shrinkLogoB64 } = ctx;
    try {
      let b64 = '';
      // v28.29 BUG-FIX: „Event-Foto verwenden" nahm bisher IMMER den Zuschnitt
      // (imageFile/imagePreview = das runde bzw. quadratisch beschnittene
      // Event-Bild). Im Mail-/Outlook-Kopf steht aber ein RECHTECK — das Foto
      // kam dort sichtbar abgeschnitten an, ohne dass der Organizer das
      // gewollt hätte. Wenn ein unbeschnittenes Original existiert (frischer
      // Upload: imageOrigFile; gespeichertes Event: editEvent.imageOrigUrl),
      // wird jetzt DIESES übernommen.
      if (imageOrigFile) {
        b64 = await fileToBase64(await compressImage(imageOrigFile, 600, 0.85, true));
      } else if (editEvent && editEvent.imageOrigUrl) {
        try {
          const resp = await fetch(editEvent.imageOrigUrl, { credentials: 'include' });
          const blob = await resp.blob();
          const f = new File([blob], 'event-photo.jpg', { type: blob.type || 'image/jpeg' });
          b64 = await fileToBase64(await compressImage(f, 600, 0.85, true));
        } catch { /* Original nicht ladbar → unten auf den Zuschnitt zurückfallen */ }
      }
      if (b64) {
        setter(b64);
        return b64;
      }
      if (imageFile) {
        b64 = await fileToBase64(await compressImage(imageFile, 600, 0.85, true));
      } else if (imagePreview && imagePreview.indexOf('data:') === 0) {
        // v28.10: Frischer Zuschnitt (Data-URL) ebenfalls komprimieren —
        // vorher ging das volle Bild unkomprimiert ins Logo (2-MB-Falle).
        b64 = await shrinkLogoB64(imagePreview);
      } else if (imagePreview) {
        const resp = await fetch(imagePreview, { credentials: 'include' });
        const blob = await resp.blob();
        const f = new File([blob], 'event-photo.jpg', { type: blob.type || 'image/jpeg' });
        b64 = await fileToBase64(await compressImage(f, 600, 0.85, true));
      }
      if (b64) setter(b64);
      else showAlert(isDe ? 'Kein Event-Foto vorhanden — bitte zuerst oben ein Bild hochladen.' : 'No event photo yet — please upload an image above first.', { variant: 'error' });
      return b64;
    } catch {
      showAlert(isDe ? 'Das Event-Foto konnte nicht übernommen werden.' : 'Could not use the event photo.', { variant: 'error' });
      return '';
    }
}

/* getStepErrorsFor — aus EventCreationPage.tsx ausgelagert (Zeilen 4750-4801 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface GetStepErrorsForCtx {
  allDay: boolean;
  endDate: string;
  lastDeregisterDate: string;
  maxParticipants: string;
  organizer: string;
  registrationDeadline: string;
  startDate: string;
  subEvents: SubEventDraft[];
  subEventsOnlyMode: boolean;
  title: string;
  unlimitedParticipants: boolean;
  userCancelAllowed: boolean;
  useSplitCapacities: boolean;
}

export function getStepErrorsForImpl(ctx: GetStepErrorsForCtx, step: number): string[] {
  const { allDay, endDate, lastDeregisterDate, maxParticipants, organizer, registrationDeadline, startDate, subEvents, subEventsOnlyMode, title, unlimitedParticipants, userCancelAllowed, useSplitCapacities } = ctx;
    const errors: string[] = [];
    switch (step) {
      case 0:
        // Schritt 1 (Grundlagen): Titel + Datum sind Pflicht. Die Datum-Checks
        // laufen hier, weil die DatePicker schon in Grundlagen stehen.
        if (!title) errors.push('title');
        if (!startDate) errors.push('startDate');
        if (!endDate) errors.push('endDate');
        // v29.55 BUG-FIX: Bei einem ganztägigen Termin ist die Uhrzeit
        // bedeutungslos — die DatePicker liefern ohne Zeitauswahl beide Male
        // 00:00, und `<=` meldete dann bei Start = Ende denselben Tag als
        // Fehler. Ganztägig wird deshalb tagesgenau verglichen: Fehler nur,
        // wenn der End-TAG vor dem Start-TAG liegt.
        if (startDate && endDate) {
          const bad = allDay
            ? endDate.slice(0, 10) < startDate.slice(0, 10)
            : new Date(endDate) <= new Date(startDate);
          if (bad) errors.push('endBeforeStart');
        }
        // v9.14: description ist optional — kein Pflichtfeld mehr
        // v28.87: Die Sub-Events stehen seit dem Wegfall von Schritt 3 in
        // Grundlagen — also wird ihre Datumsprüfung hier mitgeführt (v18.36:
        // Ende vor Start laesst den Outlook-Create-Flow mit HTTP 400 scheitern).
        if (subEvents.some(se => se.title && se.title.trim() && se.startDate && se.endDate && new Date(se.endDate) <= new Date(se.startDate))) {
          errors.push('subEventEndBeforeStart');
        }
        break;
      case 1:
        // v24.12: Schritt 2 (Organizer & Team) — mindestens ein Organizer ist Pflicht.
        if (!organizer) errors.push('organizer');
        break;
      case 2:
        // Schritt 3 (Ort & Programm) ist ohne Pflicht-Validierung —
        // Adresse / Agenda / Transferzeiten sind alle optional.
        break;
      case 3:
        // Schritt 4 (Kapazität & Sichtbarkeit).
        // v29.21 (Audit B4): Im „Nur Sub-Events"-Modus sind die geprüften
        // Felder gar nicht bedienbar — der sichtbare DatePicker editiert die
        // Klammer-Frist, die Abmeldefrist ist ausgegraut, der Kapazitätsblock
        // durch die Erklär-Box ersetzt. Die Prüfungen sperrten „Weiter" dann
        // ohne sichtbaren Grund und ohne Ausweg. Gleiches bei geteilter
        // Kapazität: maxParticipants ist dort per Konvention 0/leer, die
        // Fehlermeldung rendert nur im Nicht-Split-Zweig.
        if (!subEventsOnlyMode) {
          if (registrationDeadline && startDate && new Date(registrationDeadline) > new Date(startDate)) errors.push('deadlineAfterStart');
          if (userCancelAllowed && lastDeregisterDate && startDate && new Date(lastDeregisterDate) > new Date(startDate)) errors.push('deregAfterStart');
          if (!useSplitCapacities && !unlimitedParticipants && (maxParticipants === '' || isNaN(Number(maxParticipants)) || Number(maxParticipants) < 0)) errors.push('maxParticipants');
        }
        break;
    }
    return errors;
}

/* applySubTransfer — aus EventCreationPage.tsx ausgelagert (Zeilen 4970-5000 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface ApplySubTransferCtx {
  asRec: (d: SubEventDraft | undefined) => Record<string, unknown>;
  childTermPlural: string;
  childTermSingular: string;
  flushActiveCommTabToState: () => void;
  isDe: boolean;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setSubTransfer: React.Dispatch<React.SetStateAction<{ fromIdx: number; groups: string[]; targets: number[]; }>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  subEvents: SubEventDraft[];
  subEventsRef: React.MutableRefObject<SubEventDraft[]>;
  subTransfer: { fromIdx: number; groups: string[]; targets: number[]; };
}

export function applySubTransferImpl(ctx: ApplySubTransferCtx): void {
  const { asRec, childTermPlural, childTermSingular, flushActiveCommTabToState, isDe, setSubEvents, setSubTransfer, showAlert, subEvents, subEventsRef, subTransfer } = ctx;
    if (!subTransfer) return;
    // v28.80: Die Kommunikationsfelder (Logo, Outlook-Text, Betreff …) stehen
    // NICHT laufend im Draft — sie leben im UI-State und werden erst beim
    // Reiterwechsel in den Slot geschrieben. Ohne diesen Flush würde man den
    // Stand VOR der letzten Bearbeitung kopieren. Der Flush schreibt synchron
    // in subEventsRef, deshalb wird von dort gelesen.
    flushActiveCommTabToState();
    const src = asRec(subEventsRef.current[subTransfer.fromIdx] || subEvents[subTransfer.fromIdx]);
    const fields: string[] = [];
    for (const g of SUB_TRANSFER_GROUPS) {
      if (subTransfer.groups.indexOf(g.key) >= 0) fields.push(...g.fields);
    }
    if (fields.length === 0 || subTransfer.targets.length === 0) { setSubTransfer(null); return; }
    setSubEvents(prev => prev.map((s, i) => {
      if (subTransfer.targets.indexOf(i) < 0) return s;
      const patch: Record<string, unknown> = {};
      for (const f of fields) {
        const v = src[f];
        // v28.80: Objekte (z.B. emailTemplateOverrides) klonen — sonst teilen
        // sich alle Ziel-Sub-Events dieselbe Referenz und eine spätere
        // Aenderung an einem würde die anderen mitziehen.
        patch[f] = (v && typeof v === 'object') ? JSON.parse(JSON.stringify(v)) : v;
      }
      return { ...s, ...(patch as unknown as Partial<SubEventDraft>) };
    }));
    const n = subTransfer.targets.length;
    setSubTransfer(null);
    showAlert(isDe
      ? `Einstellungen auf ${n} ${n === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} übertragen. Nicht vergessen zu speichern.`
      : `Settings transferred to ${n} sub-event(s). Don't forget to save.`,
      { variant: 'success' });
}

/* toggleDaySubEvent — aus EventCreationPage.tsx ausgelagert (Zeilen 5249-5329 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface ToggleDaySubEventCtx {
  activeScopeIdx: number;
  allDay: boolean;
  berlinLocalToUtcIso: (localStr: string) => string;
  dayKeyOfDate: (d: Date) => string;
  dayKeyOfSub: (se: SubEventDraft) => string;
  dayLabel: (d: Date) => string;
  endDate: string;
  makeSubEventDraft: (patch: Partial<SubEventDraft>) => SubEventDraft;
  removedSavedSubs: SubEventDraft[];
  setRemovedSavedSubs: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setScope: (idx: number) => void;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  startDate: string;
  subEvents: SubEventDraft[];
}

export function toggleDaySubEventImpl(ctx: ToggleDaySubEventCtx, d: Date): void {
  const { activeScopeIdx, allDay, berlinLocalToUtcIso, dayKeyOfDate, dayKeyOfSub, dayLabel, endDate, makeSubEventDraft, removedSavedSubs, setRemovedSavedSubs, setScope, setSubEvents, startDate, subEvents } = ctx;
    if (!d) return;
    const key = dayKeyOfDate(d);
    // v29.17: Die An-/Abwahl-Entscheidung fällt IM Funktions-Updater, auf dem
    // tatsächlichen State — nicht auf dem `subEvents` aus der Render-Closure.
    // Bei 20+ Terminen dauert ein Re-Render des Wizards spürbar; wer in der
    // Zeit erneut klickt, dessen zweiter Klick sah vorher noch den ALTEN
    // Stand: Ein eben abgewählter Tag galt als „nicht vorhanden" und wurde
    // wieder angelegt — das Abwählen wirkte „kaputt". Mit dem Updater ist
    // jeder Klick ein echter Toggle auf dem aktuellen Stand.
    //
    // Scope-Korrektur vorab aus der Closure: Steht der Reiter gerade auf dem
    // Tag, der entfernt wird, zurück auf die Klammer. Im (seltenen) Stale-Fall
    // unterbleibt nur die Korrektur — scopeSub sichert den Index ohnehin ab.
    const closureIdx = subEvents.findIndex(se => dayKeyOfSub(se) === key);
    if (closureIdx >= 0 && activeScopeIdx === closureIdx + 1) setScope(0);
    // v29.22: Abwahl eines GESPEICHERTEN Termins → in removedSavedSubs parken
    // (Orange im Kalender, per Klick rückholbar). Ein ORANGE-Tag → Draft aus
    // dem Park zurückholen statt einen neuen anzulegen. Beide Übergänge sind
    // über die Guards in den Updatern idempotent — schnelle Doppelklicks auf
    // veraltetem Render-Stand (v29.17-Falle) können weder doppelt parken noch
    // doppelt anlegen.
    const closureExisting = closureIdx >= 0 ? subEvents[closureIdx] : undefined;
    if (closureExisting) {
      if (closureExisting.dbId) {
        setRemovedSavedSubs(prev => prev.some(x => x.id === closureExisting.id) ? prev : [...prev, closureExisting]);
      }
      setSubEvents(prev => prev.filter(x => x.id !== closureExisting.id));
      return;
    }
    const stashed = removedSavedSubs.find(x => dayKeyOfSub(x) === key);
    if (stashed) {
      setRemovedSavedSubs(prev => prev.filter(x => x.id !== stashed.id));
      setSubEvents(prev => prev.some(x => x.id === stashed.id) ? prev : [...prev, stashed]);
      return;
    }
    setSubEvents(prev => {
      const existingIdx = prev.findIndex(se => dayKeyOfSub(se) === key);
      if (existingIdx >= 0) {
        // v28.96: KEINE Rückfrage je Klick. Im Kalender wird aus- und
        // abgewählt, oft mehrfach hintereinander — ein Modal bei jedem Klick
        // macht genau das unbenutzbar. Der Datenverlust-Hinweis steht ohnehin
        // schon an der richtigen Stelle: beim SPEICHERN listet
        // handleSubmitInner alle Sub-Events auf, die dabei endgültig gelöscht
        // würden (toDelete), und fragt einmal nach. Bis dahin ist nichts
        // passiert. (Stale-Doppelklick: der Tag wurde eben schon behandelt.)
        return prev;
      }
      // v28.92: Der Termin bekommt die UHRZEIT des Hauptevents, gelegt auf
      // diesen Tag — läuft das Event von 9 bis 17 Uhr, gilt das auch für den
      // einzelnen Tag. Ein Ganztags-Block (00:00–23:59) würde den Teilnehmern
      // den kompletten Kalendertag zustellen.
      //
      // Die Zeitfelder bleiben bewusst NICHT leer: Ein Sub-Event ohne Zeiten
      // erbt seit v28.66 die TERMINE des Hauptevents — bei einer Reihe also
      // 01.09.–01.10. für jeden einzelnen Tag statt des Tages selbst.
      //
      // Zwei Fälle, in denen die Uhrzeit nichts hergibt, fallen auf den ganzen
      // Tag zurück: gar keine Zeiten am Hauptevent, und ein mehrtägiges Event,
      // dessen Endzeit nicht nach der Startzeit liegt (z.B. 01.09. 00:00 bis
      // 01.10. 00:00 — daraus liesse sich für einen Tag keine gültige Spanne
      // bauen).
      const timeOf = (v: string): string => {
        const t = (v || '').slice(11, 16);
        return /^\d{2}:\d{2}$/.test(t) ? t : '';
      };
      const startTime = timeOf(startDate);
      const endTime = timeOf(endDate);
      const usable = !!startTime && !!endTime && endTime > startTime;
      const start = berlinLocalToUtcIso(`${key}T${usable ? startTime : '00:00'}`);
      const end = berlinLocalToUtcIso(`${key}T${usable ? endTime : '23:59'}`);
      // v29.52: Der erzeugte Tag erbt „ganztägig" vom Hauptevent — und ist es
      // auch dann, wenn sich aus dem Zeitraum keine Uhrzeit ableiten ließ
      // (`!usable`, z.B. Klammer 01.09. 00:00 – 25.09. 17:00). Genau dieser
      // Fall hat die ganztägigen Blocker erzeugt, über die sich Organizer
      // beschwert haben: 00:00–23:59 sieht in Outlook aus wie ein Tag
      // Vollsperrung, ist aber technisch ein normaler Termin.
      return prev.concat([makeSubEventDraft({
        title: dayLabel(d), startDate: start, endDate: end,
        allDay: allDay || !usable,
      })]);
    });
}

