/**
 * v30.54: Offene Aufgaben beim Veranstalter — LIVE aus den Daten abgeleitet.
 *
 * Der Import (`b2runBibImport`) beantwortet die Frage einmal, im Moment des
 * Rücklaufs. Danach geht es weiter: Zwischen Import und Lauftag melden sich
 * Leute ab, andere rücken nach, jemand kommt neu dazu. Jede dieser Bewegungen
 * erzeugt eine Aufgabe beim Veranstalter — und die entsteht, während niemand
 * die Import-Maske offen hat.
 *
 * **Deshalb wird hier nichts gespeichert, sondern gerechnet.** Grundlage sind
 * allein die Teilnehmerzeilen: Wer hat eine Startnummer, wer ist abgemeldet,
 * wer hat wen ersetzt. Eine gespeicherte Aufgabenliste wäre eine zweite
 * Wahrheit neben den Anmeldungen und liefe auseinander, sobald jemand eine
 * Abmeldung rückgängig macht oder eine Nummer von Hand korrigiert. Gespeichert
 * wird nur das Gegenteil: welche Aufgabe der Organizer **abgehakt** hat.
 *
 * Die vier Fälle:
 *
 *  1. `transfer` — Abgemeldeter MIT Nummer, Nachrücker steht fest und hat noch
 *     keine: Die Nummer wandert. In DEX per Knopf, beim Veranstalter von Hand.
 *  2. `assign`   — Abgemeldeter MIT Nummer, kein aufgezeichneter Nachrücker,
 *     aber jemand ist angemeldet OHNE Nummer: Der Platz ist besetzt, die
 *     Zuordnung steht nur nicht in den Daten (Direktanmeldung statt
 *     Warteliste). Vorschlag, keine Feststellung.
 *  3. `unregister` — Abgemeldeter MIT Nummer, und es gibt niemanden, der sie
 *     übernehmen könnte: erst dann beim Veranstalter abmelden.
 *  4. `register` — Angemeldet OHNE Nummer, und es ist keine freie Nummer übrig:
 *     beim Veranstalter nachmelden.
 *
 * Die Reihenfolge ist bewusst so: Fall 3 und 4 dürfen erst greifen, wenn 1
 * und 2 verbraucht sind. Sonst stünde gleichzeitig „Nummer 2912 abmelden" und
 * „Laura nachmelden" da — zwei Aufgaben für einen Vorgang, der eine einzige
 * Ummeldung ist.
 */

import { SPRegistration } from '../services/EventService';

const ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

export type B2RunTodoKind = 'transfer' | 'assign' | 'unregister' | 'register';

export interface B2RunTodo {
  /** Stabiler Schlüssel zum Abhaken — überlebt einen Reload. */
  key: string;
  kind: B2RunTodoKind;
  bib: string;
  /** Person, die die Nummer bisher hat/hatte. */
  from?: SPRegistration;
  /** Person, die sie bekommen soll. */
  to?: SPRegistration;
  /** Steht die Paarung in den Daten (Nachrück-Kette) oder ist sie erschlossen? */
  certain: boolean;
  /** Was der Organizer beim Veranstalter tun muss — in einem Satz. */
  action: string;
}

const lc = (s: string | undefined | null): string => (s || '').toLowerCase().trim();
const bibOf = (r: SPRegistration): string =>
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  String((r as any).Startnummer || '').trim();
const nameOf = (r: SPRegistration): string =>
  `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || r.ParticipantEmail || '—';
const ts = (v: string | undefined): number => {
  const t = new Date(v || '').getTime();
  return isFinite(t) ? t : Number.MAX_SAFE_INTEGER;
};

export function buildB2RunTodos(regs: SPRegistration[] | null | undefined): B2RunTodo[] {
  const rows = (regs || []).filter(r => !!r.ParticipantEmail);
  const byEmail = new Map<string, SPRegistration>();
  for (const r of rows) {
    const k = lc(r.ParticipantEmail);
    const prev = byEmail.get(k);
    // Bei zwei Zeilen zur selben Adresse gewinnt die aktive (Doppel-Adressen-
    // Falle aus CLAUDE.md).
    if (!prev || (ACTIVE_STATI.indexOf(r.Status) >= 0 && ACTIVE_STATI.indexOf(prev.Status) < 0)) {
      byEmail.set(k, r);
    }
  }
  const active = rows.filter(r => ACTIVE_STATI.indexOf(r.Status) >= 0);
  const activeWithoutBib = active
    .filter(r => !bibOf(r))
    .sort((a, b) => ts(a.RegistrationDate) - ts(b.RegistrationDate));
  // Abgemeldete, die eine Nummer tragen — nur die erzeugen überhaupt Arbeit.
  const cancelledWithBib = rows
    .filter(r => r.Status === 'Abgemeldet' && bibOf(r))
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    .sort((a, b) => ts((a as any).CancellationDate) - ts((b as any).CancellationDate));

  const todos: B2RunTodo[] = [];
  const claimed = new Set<string>();

  // 1) Aufgezeichnete Nachrück-Kette zuerst — das ist die einzige Paarung,
  //    die in den Daten steht.
  for (const c of cancelledWithBib) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const succEmail = lc((c as any).ReplacedByParticipantEmail);
    if (!succEmail) continue;
    const succ = byEmail.get(succEmail);
    if (!succ || ACTIVE_STATI.indexOf(succ.Status) < 0) continue;
    if (bibOf(succ)) continue;            // hat die Nummer schon → nichts zu tun
    if (claimed.has(lc(succ.ParticipantEmail))) continue;
    claimed.add(lc(succ.ParticipantEmail));
    todos.push({
      key: `transfer|${bibOf(c)}|${lc(succ.ParticipantEmail)}`,
      kind: 'transfer', bib: bibOf(c), from: c, to: succ, certain: true,
      action: `Startnummer ${bibOf(c)} beim Veranstalter von ${nameOf(c)} auf ${nameOf(succ)} ummelden.`,
    });
  }

  // 2) Freie Nummern den übrigen Personen ohne Nummer zuordnen — zeitliche
  //    Reihenfolge, ausdrücklich als Vorschlag.
  const openBibs = cancelledWithBib.filter(c => !todos.some(t => t.bib === bibOf(c)));
  const openPeople = activeWithoutBib.filter(r => !claimed.has(lc(r.ParticipantEmail)));
  const pairs = Math.min(openBibs.length, openPeople.length);
  for (let i = 0; i < pairs; i++) {
    const c = openBibs[i];
    const p = openPeople[i];
    claimed.add(lc(p.ParticipantEmail));
    todos.push({
      key: `assign|${bibOf(c)}|${lc(p.ParticipantEmail)}`,
      kind: 'assign', bib: bibOf(c), from: c, to: p, certain: false,
      action: `Startnummer ${bibOf(c)} beim Veranstalter von ${nameOf(c)} auf ${nameOf(p)} ummelden — DEX hat hier keinen Nachrücker aufgezeichnet, bitte prüfen.`,
    });
  }

  // 3) Nummern, für die niemand mehr übrig ist.
  for (let i = pairs; i < openBibs.length; i++) {
    const c = openBibs[i];
    todos.push({
      key: `unregister|${bibOf(c)}`,
      kind: 'unregister', bib: bibOf(c), from: c, certain: true,
      action: `Startnummer ${bibOf(c)} (${nameOf(c)}) beim Veranstalter abmelden — es ist niemand da, der sie übernimmt.`,
    });
  }

  // 4) Personen, für die keine Nummer übrig ist.
  for (let i = pairs; i < openPeople.length; i++) {
    const p = openPeople[i];
    todos.push({
      key: `register|${lc(p.ParticipantEmail)}`,
      kind: 'register', bib: '', to: p, certain: true,
      action: `${nameOf(p)} beim Veranstalter nachmelden — angemeldet, aber ohne Startnummer.`,
    });
  }

  return todos;
}

export const B2RUN_TODO_LABELS: Record<B2RunTodoKind, string> = {
  transfer: 'Ummelden',
  assign: 'Ummelden (bitte prüfen)',
  unregister: 'Abmelden',
  register: 'Nachmelden',
};
