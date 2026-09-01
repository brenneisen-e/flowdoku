/* outlookChanges — aus EventCreationPage.tsx ausgelagert (Zeilen 6034-6284 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
import * as React from 'react';
import { OutlookConfirmItem, SubEventDraft } from '../../wizard/wizardTypes';
import { reinsertOrganizerPlaceholder } from '../../wizard/wizardHelpers';
import { stripOutlookWrapper } from '../../../services/EmailTemplates';
import { buildOutlookLocation } from '../../../utils/eventFormat';
import { dlog } from '../../../utils/debugLog';
import { EmailOverrideEntry } from '../../wizard/emailOverrideEntry';

export interface OutlookChangesCtx {
  activeCommTabIdx: number;
  addrCity: string;
  addrHouseNo: string;
  addrStreet: string;
  addrZip: string;
  allDay: boolean;
  berlinLocalToUtcIso: (localStr: string) => string;
  childEventsOf: (parentEventId: string) => import("../../../types/index").DeloitteEvent[];
  editEvent: import("../../../types/index").DeloitteEvent;
  endDate: string;
  headerImageLayout: { width: number; paddingV: number; paddingH: number; };
  initialHeaderImageLayoutRef: React.MutableRefObject<{ width: number; paddingV: number; paddingH: number; }>;
  initialOutlookSnapshot: React.MutableRefObject<{ title: string; startDate: string; endDate: string; outlookBody: string; outlookLocation: string; outlookSubject: string; outlookStart: string; outlookEnd: string; organizers: string; outlookLogo: string; allDay: boolean; showAsFree: boolean; }>;
  location: string;
  onlineMeetingChanged: () => boolean;
  organizer: string;
  outlookBody: string;
  outlookEndOverride: string;
  outlookLocationOverride: string;
  outlookStartOverride: string;
  resolveTopLevelCommState: () => { emailLanguage: string; emailLogoBase64: string; outlookLogoBase64: string; outlookBody: string; outlookHeading: string; outlookSubheading: string; outlookSubject: string; disableEmails: boolean; disableRegistrationEmail: boolean; disableCancellationEmail: boolean; autoDeregisterOnDecline: boolean; inactiveHandling?: 'notify' | 'autoderegister'; disableOutlook: boolean; emailTemplateOverrides: Record<string, EmailOverrideEntry>; };
  showAsFree: boolean;
  startDate: string;
  subEventCalendar: boolean;
  subEventsOnlyMode: boolean;
  subEventsRef: React.MutableRefObject<SubEventDraft[]>;
  title: string;
}

export function detectOutlookRelevantChangesImpl(ctx: OutlookChangesCtx): { items: OutlookConfirmItem[] } {
  const { activeCommTabIdx, addrCity, addrHouseNo, addrStreet, addrZip, allDay, berlinLocalToUtcIso, childEventsOf, editEvent, endDate, headerImageLayout, initialHeaderImageLayoutRef, initialOutlookSnapshot, location, onlineMeetingChanged, organizer, outlookBody, outlookEndOverride, outlookLocationOverride, outlookStartOverride, resolveTopLevelCommState, showAsFree, startDate, subEventCalendar, subEventsOnlyMode, subEventsRef, title } = ctx;
    const items: OutlookConfirmItem[] = [];
    if (!editEvent) return { items };
    const snap = initialOutlookSnapshot.current;
    // v11.64: Datetime-Vergleich über Date.getTime(), nicht String. Sonst
    // kippt der Vergleich an Format-Unterschieden (snap kommt roh aus SP
    // mit „2026-09-24T16:00:00Z", currentStart geht durch
    // berlinLocalToUtcIso() und wird „2026-09-24T16:00:00.000Z" — gleicher
    // Zeitpunkt, anderer String). Das hat den Hauptevent in v11.63
    // fälschlich als „Startzeit, Endzeit geändert" gemeldet.
    const sameInstant = (a: string, b: string): boolean => {
      if (a === b) return true;
      if (!a && !b) return true;
      if (!a || !b) return false;
      const da = new Date(a).getTime();
      const db = new Date(b).getTime();
      if (isNaN(da) || isNaN(db)) return a === b;
      return da === db;
    };
    const currentTitle = title || '';
    const currentStart = startDate ? berlinLocalToUtcIso(startDate) : '';
    const currentEnd = endDate ? berlinLocalToUtcIso(endDate) : '';
    // Outlook-Body vergleich anhand des „rohen" Body (ohne Wrapper), da der
    // Wrapper bei jedem Save neu gebaut wird und dadurch immer „änderbar"
    // aussehen würde. Vergleich gegen den initial gestrippten Wert.
    // v29.21 (Audit): denselben Reinsert wie die outlookBody-Hydration
    // anwenden — sonst enthält der State den Platzhalter, der Snapshot den
    // gebackenen Namen, und der Vergleich meldete bei JEDEM Save „Termin-Text
    // geändert" (Dauer-False-Positive des Update-Modals).
    const initialStripped = reinsertOrganizerPlaceholder(stripOutlookWrapper(snap.outlookBody || ''), editEvent?.organizers || []);
    // v29.21 (Audit): Auf einem Sub-Tab liegt der AKTUELLE Top-Level-Body im
    // Top-Level-Slot (resolveTopLevelCommState) — der alte Fallback verglich
    // den Mount-Snapshot mit sich selbst, eine Hauptevent-Body-Änderung vor
    // dem Reiterwechsel wurde also nie erkannt: kein Update-Angebot, kein
    // OutlookDirty, Teilnehmer-Kalender blieben still veraltet. Betreff, Logo
    // und DisableOutlook daneben machten es längst so.
    const currentStripped = activeCommTabIdx === 0 ? (outlookBody || '') : (resolveTopLevelCommState().outlookBody || '');
    const currentTopLocation = outlookLocationOverride.trim() || buildOutlookLocation(location, { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity });
    const currentTopSubject = (resolveTopLevelCommState().outlookSubject || '').trim();
    // v19.20: globale Header-Bild-Layout-Änderung (Breite/Innenabstand) erkennen.
    // Das Layout steht NICHT im rohen Body (wird erst beim Wrappen angewendet),
    // betrifft aber den Hero-Bild-Kopf des Outlook-Termins — daher als eigenes
    // Änderungs-Feld „layout" werten, damit das Update-Modal aufgeht (und der
    // Grund klar als „Kopfbild" benannt wird, nicht irreführend als „Termin-Text").
    const initLayout = initialHeaderImageLayoutRef.current;
    const layoutChanged = headerImageLayout.width !== initLayout.width
      || headerImageLayout.paddingV !== initLayout.paddingV
      || headerImageLayout.paddingH !== initLayout.paddingH;
    // v28.30: Kopfbild-Wechsel erkennen. Das Bild steckt weder im rohen
    // Termin-Text (es wird erst beim Wrappen als {{ORB_URL}} eingesetzt) noch
    // im Layout — eine reine Bild-Änderung war für den Detektor deshalb
    // unsichtbar, und das Update-Modal bot nur Events an, bei denen zufällig
    // noch etwas anderes anders war.
    const curTopLogo = resolveTopLevelCommState().outlookLogoBase64 || '';
    const topLogoChanged = curTopLogo !== (snap.outlookLogo || '');
    const topChangedFields: Array<'title' | 'startDate' | 'endDate' | 'outlookBody' | 'location' | 'subject' | 'layout' | 'organizer' | 'logo'> = [];
    if (currentTitle !== (snap.title || '')) topChangedFields.push('title');
    if (!sameInstant(currentStart, snap.startDate || '')) topChangedFields.push('startDate');
    if (!sameInstant(currentEnd, snap.endDate || '')) topChangedFields.push('endDate');
    if (currentStripped !== initialStripped) topChangedFields.push('outlookBody');
    if (layoutChanged) topChangedFields.push('layout');
    // v29.38: reine Teams-Link-Änderung ebenfalls als Kopf-/Layout-Änderung
    // melden (eigener Grund wäre eine weitere Feld-Variante — der Termin-Text
    // ändert sich hier tatsächlich, deshalb 'outlookBody').
    // v30.26: auch der Online-Meeting-Modus (s. onlineMeetingChanged).
    if (onlineMeetingChanged() && topChangedFields.indexOf('outlookBody') < 0) topChangedFields.push('outlookBody');
    if (topLogoChanged) topChangedFields.push('logo');
    // v22.48: Organizer-Änderung. Der Outlook-Standardtext enthält die
    // Organizer-Namen („wendet euch bitte an …"). Solange der Text NICHT
    // individuell überschrieben ist (Body leer = Default), ändert eine
    // Organizer-Änderung den Termin-Text → als Outlook-relevant werten. Bei
    // custom Body bleibt der Text bewusst stehen → nicht melden.
    const currentOrganizers = organizer.split(';').map(s => s.trim()).filter(Boolean).join(';');
    const bodyIsDefault = !currentStripped.trim() && !initialStripped.trim();
    if (bodyIsDefault && currentOrganizers !== (snap.organizers || '')) topChangedFields.push('organizer');
    // v18.34: reine Ort-Änderung gilt ebenfalls als Outlook-relevant.
    if (currentTopLocation !== (snap.outlookLocation || '')) topChangedFields.push('location');
    // v18.42: reine Betreff-Änderung gilt ebenfalls als Outlook-relevant.
    if (currentTopSubject !== (snap.outlookSubject || '').trim()) topChangedFields.push('subject');
    // v18.44: abweichendes Outlook-Datum (Override) gilt als Termin-Änderung.
    if ((outlookStartOverride || '') !== (snap.outlookStart || '') && topChangedFields.indexOf('startDate') < 0) topChangedFields.push('startDate');
    if ((outlookEndOverride || '') !== (snap.outlookEnd || '') && topChangedFields.indexOf('endDate') < 0) topChangedFields.push('endDate');
    // v29.52: Umschalten auf/von „ganztägig" ist eine Termin-Änderung.
    if (!!allDay !== !!snap.allDay && topChangedFields.indexOf('startDate') < 0) topChangedFields.push('startDate');
    // v29.54: Wechsel zwischen Beschäftigt und Frei ändert den bestehenden
    // Termin ebenfalls — ohne diesen Vergleich bliebe er im Kalender stehen.
    if (!!showAsFree !== !!snap.showAsFree && topChangedFields.indexOf('startDate') < 0) topChangedFields.push('startDate');
    // v11.61: Beide Pointer prüfen — DEX_CreateOutlookEvent setzt nur
    // CalendarLink auf Erfolg, OutlookEventId bleibt leer. Wer beides
    // leer hat, hatte nie einen Outlook-Termin.
    const topHasOutlook = !!editEvent.outlookEventId || !!editEvent.calendarLink;
    // v18.45 BUG-FIX: für das Hauptevent IMMER dessen Top-Level-DisableOutlook
    // prüfen — nicht das rohe `disableOutlook` (das hält beim Speichern auf einem
    // Sub-Event-Tab den Sub-Wert). Sonst wurde das Hauptevent fälschlich im
    // Update-Modal gelistet, obwohl dort Outlook deaktiviert ist (z.B. Event mit
    // Outlook nur auf Sub-Event-Ebene).
    const topDisableOutlook = resolveTopLevelCommState().disableOutlook;
    // v18.51: Im „Nur für Sub-Events"-Modus (subEventsOnlyMode) ist das
    // Hauptevent von der Teilnehmer-Anmeldung ausgenommen — niemand meldet sich
    // direkt fürs Hauptevent an. Ein Outlook-Update-Hinweis fürs Hauptevent ist
    // dann sinnlos und wird unterdrückt (Sub-Events bekommen weiter ihre Hinweise).
    if (topChangedFields.length > 0 && !topDisableOutlook && topHasOutlook && !subEventsOnlyMode) {
      items.push({
        kind: 'top',
        eventId: editEvent.id,
        title: currentTitle || editEvent.title || '',
        changedFields: topChangedFields,
      });
    }
    // Sub-Events: pro Sub-Event vergleichen.
    // v11.60: subEventsRef statt subEvents — der Flush hat die aktuellen
    // UI-Werte gerade synchron in den Ref geschrieben, der React-State
    // ist noch nicht propagiert.
    for (const s of subEventsRef.current) {
      if (!s.dbId) continue;
      const initTitle = s.initialTitle || '';
      const initStart = s.initialStartDate || '';
      const initEnd = s.initialEndDate || '';
      const initBodyStripped = stripOutlookWrapper(s.initialOutlookBody || '');
      const curBodyStripped = (s.outlookBody || '');
      const hasOutlookEvId = !!s.initialOutlookEventId || !!s.initialCalendarLink;
      // v28.30: Kopfbild pro Sub-Event vergleichen — und zwar das WIRKSAME.
      // Seit v28.29 erbt ein Sub-Event ohne eigenes Bild das des Hauptevents;
      // wechselt dort das Bild, ändert sich also auch der Sub-Event-Termin,
      // obwohl im Sub-Draft selbst nichts steht.
      const initSubLogo = s.initialOutlookLogoBase64 || '';
      const curSubLogo = s.outlookLogoBase64 || curTopLogo;
      const subChangedFields: Array<'title' | 'startDate' | 'endDate' | 'outlookBody' | 'layout' | 'logo'> = [];
      if ((s.title || '') !== initTitle) subChangedFields.push('title');
      // v11.64: auch hier semantischer Vergleich — gleiche Falle wie oben.
      if (!sameInstant(s.startDate || '', initStart)) subChangedFields.push('startDate');
      if (!sameInstant(s.endDate || '', initEnd)) subChangedFields.push('endDate');
      // v29.52: dasselbe für den Ganztags-Haken je Sub-Event.
      if (!!s.allDay !== !!s.initialAllDay && subChangedFields.indexOf('startDate') < 0) subChangedFields.push('startDate');
      if (!!s.showAsFree !== !!s.initialShowAsFree && subChangedFields.indexOf('startDate') < 0) subChangedFields.push('startDate'); // v29.54
      if (curBodyStripped !== initBodyStripped) subChangedFields.push('outlookBody');
      // v19.20: globale Header-Bild-Layout-Änderung betrifft auch die
      // Sub-Event-Outlook-Termine (gleicher Hero-Bild-Kopf) — als eigenes
      // „layout"-Feld werten, damit das Update-Modal sie mit auflistet.
      if (layoutChanged) subChangedFields.push('layout');
      // v29.38: Der Teams-Link gilt event-weit — er steckt auch in den
      // Sub-Event-Terminen. Ändert er sich, müssen die genauso aktualisiert
      // werden, sonst zeigen sie weiter den alten (oder gar keinen) Link.
      if (onlineMeetingChanged() && subChangedFields.indexOf('outlookBody') < 0) subChangedFields.push('outlookBody');
      if (curSubLogo !== initSubLogo) subChangedFields.push('logo');
      // v30.9: Auch der WIRKSAME Outlook-Betreff zählt als Änderung. Seit
      // v30.7 erben Kalender-Tage den Hauptevent-Titel als Betreff — der
      // Save schreibt das zwar in die OutlookSubject-Spalte, aber ohne
      // UpdateEvent in der Queue bleibt der BESTEHENDE Termin beim alten
      // Namen (Tages-Datum). Verglichen wird gegen den gespeicherten Stand
      // der Sub-Event-Zeile, nicht gegen einen Draft-Schnappschuss — so
      // greift es auch für Events, die vor v30.7 angelegt wurden.
      {
        const storedRow = childEventsOf(editEvent.id).find(c => c.id === s.dbId);
        const effSubSubject = (s.outlookSubject || '').trim() || (subEventCalendar ? (title || editEvent.title || '').trim() : '');
        const storedSubSubject = ((storedRow && storedRow.outlookSubject) || '').trim();
        if (effSubSubject !== storedSubSubject && subChangedFields.indexOf('title') < 0) subChangedFields.push('title');
      }
      // v11.66: Debug-Log für jeden Sub-Event, damit wir in der Browser-
      // Konsole nachvollziehen können, warum das Modal manchmal nicht
      // erscheint. v11.67: JSON.stringify damit der Browser die Werte
      // direkt anzeigt (statt nur „Object" mit Klick zum Aufklappen).
      // v11.79: nur noch sichtbar, wenn der Maintainer in der Console
      // `window.__dexDebug = true` setzt — sonst spammt das Log im
      // Normalbetrieb die DevTools voll.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      if (typeof window !== 'undefined' && (window as any).__dexDebug) {
        // eslint-disable-next-line no-console
        dlog('perf', '[DEX][outlook-detect][sub] ' + JSON.stringify({
          dbId: s.dbId,
          title: s.title,
          subChangedFields,
          disableOutlook: s.disableOutlook,
          hasOutlookEvId,
          initialOutlookEventId: s.initialOutlookEventId,
          initialCalendarLink: s.initialCalendarLink,
          bodyLenInitial: (s.initialOutlookBody || '').length,
          bodyLenCurrent: (s.outlookBody || '').length,
          bodyLenInitStripped: initBodyStripped.length,
          bodyMatch: curBodyStripped === initBodyStripped,
          titleMatch: (s.title || '') === initTitle,
        }));
      }
      if (subChangedFields.length > 0 && !s.disableOutlook) {
        items.push({
          kind: 'sub',
          eventId: s.dbId,
          title: s.title || '',
          changedFields: subChangedFields,
          // v11.68: ohne CalendarLink/OutlookEventId existiert kein Outlook-Termin
          // — Save persistiert den neuen Body in DEX_Events, aber wir können
          // kein UpdateEvent queuen. Modal rendert Info-Eintrag.
          noOutlookYet: !hasOutlookEvId,
        });
      }
    }
    // v14.8: Items aus persistiertem OutlookDirty-Flag nachziehen. Wenn ein
    // Sub-Event oder das Hauptevent in einer früheren Session als „Outlook-
    // Update ausstehend" markiert wurde (User hat den Haken damals nicht
    // gesetzt → OutlookDirty=true wurde in SP geschrieben), soll der nächste
    // Save trotzdem das Modal anbieten — auch ohne neue inhaltliche Änderung
    // in dieser Session. Sonst bleibt der Dirty-Flag ewig hängen und der
    // Yellow-Hint in Schritt 1 wird nie aufgelöst.
    const hasItemForEvent = (id: string): boolean => items.some(it => it.eventId === id);
    // Hauptevent
    // v18.50 BUG-FIX: auch im Dirty-Marker-Pfad das Top-Level-DisableOutlook
    // prüfen (nicht das rohe `disableOutlook`, das beim Speichern auf einem
    // Sub-Event-Tab den Sub-Wert hält) — sonst taucht das Hauptevent im
    // Update-Modal als „Frühere Änderung nicht synchronisiert" auf, obwohl
    // dort Outlook deaktiviert ist (Event mit Outlook nur auf Sub-Event-Ebene).
    // Gleiche Falle wie v18.45 im Changed-Fields-Pfad oben.
    if (editEvent.outlookDirty && !topDisableOutlook && !subEventsOnlyMode
        && (editEvent.outlookEventId || editEvent.calendarLink)
        && !hasItemForEvent(editEvent.id)) {
      items.push({
        kind: 'top',
        eventId: editEvent.id,
        title: title || editEvent.title || '',
        changedFields: [],
      });
    }
    // Sub-Events
    for (const s of subEventsRef.current) {
      if (!s.dbId) continue;
      if (s.disableOutlook) continue;
      const hasOutlookEvId = !!s.initialOutlookEventId || !!s.initialCalendarLink;
      if (!hasOutlookEvId) continue;
      const childEvt = childEventsOf(editEvent.id).find(c => c.id === s.dbId);
      if (childEvt && childEvt.outlookDirty && !hasItemForEvent(s.dbId)) {
        items.push({
          kind: 'sub',
          eventId: s.dbId,
          title: s.title || '',
          changedFields: [],
          noOutlookYet: false,
        });
      }
    }
    // v11.79: gated debug log — siehe oben.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (typeof window !== 'undefined' && (window as any).__dexDebug) {
      // eslint-disable-next-line no-console
      dlog('perf', '[DEX][outlook-detect][result] ' + JSON.stringify({
        itemsCount: items.length,
        items,
        activeCommTabIdx,
        topOutlookBodyLen: (outlookBody || '').length,
        topInitialOutlookBodyLen: (snap.outlookBody || '').length,
        topBodyMatch: currentStripped === initialStripped,
      }));
    }
    return { items };
}
