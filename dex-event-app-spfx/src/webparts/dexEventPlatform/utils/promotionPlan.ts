/**
 * v30.66 — Wie viele Personen dürfen von der Warteliste nachrücken?
 *
 * Gruppen-Rechnung für die Sammel-Aktion „Nachrücken & IDs für ALLE Events
 * nachholen" (`useWaitlistActions.runHealAllEvents`), nach dem Vorbild von
 * `runManualPromote` (v29.16): Bei geteilten Kapazitäten ist `maxParticipants`
 * 0 (CLAUDE.md), die Obergrenze steht in `durchstarterCapacity`/
 * `funstarterCapacity` — wer `maxParticipants` als Grenze prüft, prüft dort
 * gar nichts und überbucht eine volle Gruppe, während die Gruppe mit freien
 * Plätzen leer ausgeht. Genau dieser Fehler steckte bis v29.16 in der
 * Einzel-Aktion.
 *
 * Bewusst NICHT abgedeckt: geteilte Kapazität mit gemeinsamer Warteliste
 * (`splitSharedWaitlist`) — siehe `PromotionPlan.sharedWaitlist`.
 *
 * Das Modul rechnet nur — es schreibt nichts. Aufrufer entscheiden, ob sie
 * fragen, protokollieren oder ausführen.
 */
import { DeloitteEvent } from '../types';
import { SPRegistration } from '../services/EventService';

/** Belegt einen Platz. Deckungsgleich mit `EventService.ACTIVE_STATI`. */
export const PROMOTE_ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

export interface PromotionGroup {
  /** `undefined` = ein einziger Topf (keine geteilten Gruppen). */
  key?: string;
  /** Anzeigename der Gruppe (`splitLabelA`/`splitLabelB`) bzw. „Plätze". */
  label: string;
  /** Obergrenze. 0 heißt ausdrücklich „unbegrenzt", nicht „voll". */
  cap: number;
  active: number;
  waiting: number;
  free: number;
  /** So viele rücken tatsächlich nach: `min(free, waiting)`. */
  count: number;
}

export interface PromotionPlan {
  /** true = je Gruppe getrennt rechnen (geteilte Kapazität OHNE gemeinsame Warteliste). */
  perGroup: boolean;
  groups: PromotionGroup[];
  /** Summe über alle Gruppen. */
  total: number;
  /** Wartet überhaupt jemand? Trennt „voll" von „niemand da" in der Meldung. */
  anyWaiting: boolean;
  /**
   * true = geteilte Kapazität MIT gemeinsamer Warteliste. Dieses Modul rechnet
   * dafür bewusst NICHT (`total` ist dann 0): Gemeinsame Warteliste heißt
   * „gemeinsame Reihenfolge", nicht „gemeinsamer Kapazitätstopf" (v30.67) —
   * wer nachrückt, muss in SEINE Gruppe passen, und das lässt sich nur durch
   * Simulation der Warteliste in ihrer Reihenfolge entscheiden. Die steht in
   * `useWaitlistActions.runManualPromote`; Aufrufer verweisen dorthin.
   */
  sharedWaitlist: boolean;
}

/**
 * Geteilte Kapazität? Dieselbe Prüfung wie in `AdminPage` und `EventContext` —
 * bewusst hier, damit sie nicht ein viertes Mal abgeschrieben wird.
 */
export function isSplitCapacityOf(ev?: DeloitteEvent | null): boolean {
  return !!ev
    && typeof ev.durchstarterCapacity === 'number'
    && typeof ev.funstarterCapacity === 'number'
    && ((ev.durchstarterCapacity || 0) > 0 || (ev.funstarterCapacity || 0) > 0);
}

/** Gruppenzuordnung einer Zeile — gewählter Typ, ersatzweise der Wunsch-Typ. */
function groupOf(r: SPRegistration): string {
  return r.StarterType || r.PreferredStarterType || '';
}

/**
 * Rechnet aus einer FRISCH gelesenen Teilnehmerliste, wie viele Personen je
 * Gruppe nachrücken dürfen. `regs` muss vollständig sein — eine durch
 * Element-Sicherheit beschnittene Liste ergibt hier zu wenige Aktive und damit
 * zu viele freie Plätze (die Falle aus v30.62). Aufrufer prüfen das über
 * `getAllRegistrations(..., onHttpError)`.
 */
export function buildPromotionPlan(
  ev: DeloitteEvent,
  regs: SPRegistration[],
  isDe: boolean,
): PromotionPlan {
  const isSplit = isSplitCapacityOf(ev);
  const lblA = (ev.splitLabelA && ev.splitLabelA.trim()) || 'Durchstarter';
  const lblB = (ev.splitLabelB && ev.splitLabelB.trim()) || 'Funstarter';
  const sharedWaitlist = isSplit && !!ev.splitSharedWaitlist;
  if (sharedWaitlist) {
    // Siehe `PromotionPlan.sharedWaitlist`: hier gibt es keine gültige Zahl,
    // die ohne Simulation der Reihenfolge herauskäme. Lieber ehrlich 0 als
    // eine Summe, die die Gruppengrenzen ignoriert.
    const waiting = regs.filter(r => r.Status === 'Warteliste').length;
    return { perGroup: false, groups: [], total: 0, anyWaiting: waiting > 0, sharedWaitlist: true };
  }
  const perGroup = isSplit;
  const base: Array<{ key?: string; label: string; cap: number }> = perGroup
    ? [
      { key: 'Durchstarter', label: lblA, cap: ev.durchstarterCapacity || 0 },
      { key: 'Funstarter', label: lblB, cap: ev.funstarterCapacity || 0 },
    ]
    : [{ key: undefined, label: isDe ? 'Plätze' : 'Seats', cap: ev.maxParticipants || 0 }];

  const groups: PromotionGroup[] = base.map(g => {
    const inGroup = (r: SPRegistration): boolean => !g.key || groupOf(r) === g.key;
    const active = regs.filter(r => PROMOTE_ACTIVE_STATI.indexOf(r.Status) >= 0 && inGroup(r)).length;
    const waiting = regs.filter(r => r.Status === 'Warteliste' && inGroup(r)).length;
    // Kapazität 0 heißt „unbegrenzt" — dann rückt die ganze Warteliste nach.
    const free = g.cap > 0 ? Math.max(0, g.cap - active) : waiting;
    return { ...g, active, waiting, free, count: Math.min(free, waiting) };
  });

  return {
    perGroup,
    groups,
    total: groups.reduce((n, g) => n + g.count, 0),
    anyWaiting: groups.some(g => g.waiting > 0),
    sharedWaitlist: false,
  };
}

/** Zeilen für den Bestätigungs-Dialog („• Gruppe: 12/20 belegt · 3 warten → 3 rücken nach"). */
export function promotionPlanLines(plan: PromotionPlan): string[] {
  return plan.groups
    .filter(g => g.waiting > 0 || g.free > 0)
    .map(g => plan.perGroup
      ? `• ${g.label}: ${g.active}/${g.cap || '∞'} belegt · ${g.waiting} auf der Warteliste → ${g.count} rücken nach`
      : `• ${g.active}/${g.cap || '∞'} belegt · ${g.waiting} auf der Warteliste → ${g.count} rücken nach`);
}
