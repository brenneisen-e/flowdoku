/**
 * v31.21 — Die Reiter über Haupt-Event und Termine („Bookmark-Tabs", v11.28)
 * als EINE Berechnung.
 *
 * Der Anlass ist ein Wunsch, kein Aufräumen: Die Reiter sollen zusätzlich über
 * der Teilnehmerliste stehen und dort klickbar sein (Nutzer-Ansage 11.09.2026:
 * „zudem möchte ich auf der Teilnehmerliste neben Teilnehmer hinzufügen auch
 * nochmal die Bookmarks angezeigt und klickbar bekommen"). Bei 435 Personen ist
 * der Weg zurück nach oben eine halbe Seite Scrollen, nur um den Termin zu
 * wechseln.
 *
 * Zwei Stellen, die dieselben Zahlen zeigen, sind aber genau die Konstruktion,
 * die auseinanderläuft — in DEX schon zweimal passiert (Handbuch-Artikel neben
 * `ArchitecturePage`, v29.10; zwei Reiter-Reihen im Wizard, v28.88). Die
 * Badge-Zahl ist hier besonders heikel: Sie hat drei Quellen, die sich je nach
 * Zustand abwechseln (Live-Zeilen der geladenen Tabelle, die Zeilen aus
 * `subEventRegsByEventId`, der gespeicherte Zähler des Events), und an jeder
 * einzelnen hängt eine eigene, teuer gelernte Bedingung:
 *
 *  - v23.2: Nicht gewählte Termine zeigen die LIVE-Zeilenzahl, nicht den
 *    veralteten `currentParticipants` — sonst springt der Badge je nachdem,
 *    welcher Reiter gerade offen ist (der 188-vs-190-Effekt).
 *  - v30.42: …aber nur, solange nicht gerade geladen wird. Beim Reiterwechsel
 *    gehört `registrations` noch dem ALTEN Termin; der neue Reiter zeigte
 *    dessen Zahl (Nutzer-Befund: erst 55, dann 51).
 *  - v30.67: dasselbe, wenn die Liste nicht lesbar war. `registrations` ist
 *    dann `[]`, die Live-Zahl wäre eine „0", die keine ist.
 *  - v22.64/v23.3: Im Klammer-Modus zählt der Haupt-Badge eindeutige Personen
 *    über alle Termine — emaillose Zeilen über ihre Zeilen-Id, sonst zeigt die
 *    Klammer weniger als die Tabelle darunter.
 *
 * Diese vier Regeln ein zweites Mal hinzuschreiben hieße, drei davon beim
 * nächsten Mal zu vergessen. Deshalb rechnet das hier, und beide Ansichten
 * rendern nur noch.
 */

import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import { shortSubEventTitle } from '../../../utils/subEventTitle';

/** Die Status, die als „belegt einen Platz" zählen. */
const ACTIVE_TAB_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

export interface EventTab {
  id: string;
  label: string;
  count: number;
  isParent: boolean;
  ev: DeloitteEvent;
}

export interface BuildEventTabsCtx {
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  events: DeloitteEvent[];
  isDe: boolean;
  /** Während des Ladens gehört `registrations` noch dem vorher gewählten Termin. */
  isLoadingRegs: boolean;
  registrations: SPRegistration[];
  /** `true` = die Liste des gewählten Termins war nicht lesbar (dann ist `[]` keine 0). */
  regsUnknown: boolean;
  selectedEvent: DeloitteEvent | null;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
}

/**
 * Liefert die Reiter — Klammer/Hauptevent zuerst, dann die Termine.
 *
 * Leeres Array heißt „hier gibt es nichts umzuschalten": kein Event gewählt,
 * oder ein normales Event ohne Termine.
 */
export function buildEventTabs(ctx: BuildEventTabsCtx): EventTab[] {
  const { childEventsOf, events, isDe, isLoadingRegs, registrations, regsUnknown, selectedEvent, subEventRegsByEventId } = ctx;
  if (!selectedEvent) return [];
  const isChild = !!selectedEvent.parentEventId;
  const siblings = isChild
    ? childEventsOf(selectedEvent.parentEventId || '')
    : childEventsOf(selectedEvent.id);
  if (!isChild && siblings.length === 0) return [];
  const parent = isChild ? events.filter(e => e.id === selectedEvent.parentEventId)[0] : selectedEvent;

  // Die Zahl der GERADE geladenen Tabelle. Gilt nur für den gewählten Reiter
  // und nur, wenn sie überhaupt etwas aussagt (nicht am Laden, lesbar).
  const liveSelectedActive = registrations.filter(r => ACTIVE_TAB_STATI.indexOf(r.Status) >= 0).length;
  const liveGilt = !isLoadingRegs && !regsUnknown;

  const tabs: EventTab[] = [];
  if (parent) {
    let parentCount = parent.currentParticipants || 0;
    const pKids = childEventsOf(parent.id);
    const haveSubData = parent.subEventsOnlyMode && pKids.length > 0 && pKids.every(c => subEventRegsByEventId[c.id] !== undefined);
    if (haveSubData) {
      const activeSet: Record<string, true> = {};
      for (const c of pKids) {
        for (const r of (subEventRegsByEventId[c.id] || [])) {
          if (ACTIVE_TAB_STATI.indexOf(r.Status) < 0) continue;
          const k = (r.ParticipantEmail || '').toLowerCase().trim() || `__noemail#${c.id}#${r.Id}`;
          activeSet[k] = true;
        }
      }
      parentCount = Object.keys(activeSet).length;
    } else if (parent.id === selectedEvent.id && liveGilt) {
      parentCount = liveSelectedActive;
    }
    tabs.push({ id: parent.id, label: parent.title || (isDe ? 'Hauptevent' : 'Main event'), count: parentCount, isParent: true, ev: parent });
  }
  for (const c of siblings) {
    const subRegs = subEventRegsByEventId[c.id];
    const subLiveCount = subRegs
      ? subRegs.filter(r => ACTIVE_TAB_STATI.indexOf(r.Status) >= 0).length
      : (c.currentParticipants || 0);
    tabs.push({
      id: c.id,
      label: shortSubEventTitle(c.title, parent ? parent.title : '') || (isDe ? 'ohne Titel' : 'untitled'),
      count: (c.id === selectedEvent.id && liveGilt) ? liveSelectedActive : subLiveCount,
      isParent: false,
      ev: c,
    });
  }
  return tabs;
}
