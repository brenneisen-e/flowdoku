/**
 * v30.54 / **korrigiert in v30.55**: Offene Aufgaben beim Veranstalter.
 *
 * ## Was in v30.54 falsch war
 *
 * Die Liste wurde AUSSCHLIESSLICH aus den Teilnehmerzeilen abgeleitet, mit der
 * Annahme „ein Abgemeldeter mit Startnummer erzeugt eine Ummeldung". Genau das
 * kann es nach einem Import aber gar nicht geben: Die Startnummern der
 * abgemeldeten Personen stehen **nur in der Datei des Veranstalters**, nie in
 * DEX — der Import schreibt sie ausschließlich der Person, die tatsächlich
 * läuft. Also war `cancelledWithBib` direkt nach dem Import leer und die Liste
 * meldete „Nichts offen", während beim Veranstalter neun Ummeldungen warteten.
 *
 * Der Denkfehler dahinter ist derselbe wie beim „verfallen" in v30.53: Ich habe
 * eine **Verpflichtung** (ich muss bei B2Run etwas ummelden) mit einem
 * **Zustand** (wer hält welche Nummer) verwechselt. Der Zustand lässt sich
 * ableiten, die Verpflichtung nicht — sie ist ein Ereignis der Vergangenheit
 * und muss festgehalten werden. Dass DEX die Nummer geschrieben hat, heißt
 * nicht, dass der Veranstalter davon weiß.
 *
 * ## Wie es jetzt läuft
 *
 * Zwei Quellen, eine Liste:
 *
 *  1. **Festgehalten** (`_b2runTodo`): Was der Import erzeugt hat. Beim
 *     Schreiben der Startnummern legt er für jede Übertragung, jede
 *     Zuordnung, jede verfallene Nummer und jede fehlende Meldung einen
 *     Eintrag an. Das ist die Verpflichtung.
 *  2. **Abgeleitet** (diese Datei): Was SEITHER passiert ist. Meldet sich
 *     jemand ab, der eine Startnummer trägt, entsteht die nächste Ummeldung —
 *     und die soll auftauchen, ohne dass jemand den Import wiederholt. Das
 *     funktioniert, weil die Nummer nach dem Import auf der Zeile der
 *     laufenden Person steht: Meldet die sich ab, ist ihre Zeile eine
 *     „abgemeldet mit Nummer"-Zeile.
 *
 * Beide werden über denselben Schlüssel zusammengeführt; erledigt ist, was
 * abgehakt wurde.
 */

import { SPRegistration } from '../services/EventService';

const ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

export type B2RunTodoKind = 'transfer' | 'assign' | 'unregister' | 'register';

/** Festgehaltene Aufgabe (Piggyback `_b2runTodo`) — bewusst flach und ohne
 *  Objekt-Referenzen, damit sie einen Reload und einen Datenstand-Wechsel
 *  übersteht. */
export interface StoredB2RunTodo {
  key: string;
  kind: B2RunTodoKind;
  bib: string;
  fromName?: string;
  fromEmail?: string;
  toName?: string;
  toEmail?: string;
  certain: boolean;
  action: string;
  /** Wann die Aufgabe entstanden ist. */
  ts: string;
}

export interface B2RunTodo extends StoredB2RunTodo {
  /** Aktuelle Zeile der Person, die die Nummer bekommen soll — für den
   *  Knopf „Nummer in DEX übertragen". Fehlt bei festgehaltenen Aufgaben,
   *  deren Person inzwischen nicht mehr in der Liste steht. */
  toReg?: SPRegistration;
  /** Trägt die Zielperson die Nummer in DEX bereits? */
  bibInDex?: boolean;
}

const lc = (s: string | undefined | null): string => (s || '').toLowerCase().trim();
const bibOf = (r: SPRegistration): string =>
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  String((r as any).Startnummer || '').trim();
export const b2runNameOf = (r: SPRegistration): string =>
  `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || r.ParticipantEmail || '—';
const ts = (v: string | undefined): number => {
  const t = new Date(v || '').getTime();
  return isFinite(t) ? t : Number.MAX_SAFE_INTEGER;
};

/**
 * Aufgaben, die SEIT dem Import entstanden sind — aus den Teilnehmerzeilen.
 *
 * Ausgangspunkt ist die Nummer, nicht die Person: Für jede Startnummer wird
 * gefragt, wer sie hält. Hält sie eine aktive Person, ist nichts zu tun. Hält
 * sie nur noch jemand Abgemeldetes, ist sie frei und sucht einen Nachfolger.
 * Über die Nummer zu gehen ist wichtig, weil dieselbe Nummer im Lauf der Zeit
 * auf mehreren Zeilen auftauchen kann (A meldet sich ab, B bekommt sie, B
 * meldet sich ab) — pro Nummer darf trotzdem nur EINE Aufgabe entstehen.
 */
export function deriveB2RunTodos(regs: SPRegistration[] | null | undefined): StoredB2RunTodo[] {
  const rows = (regs || []).filter(r => !!r.ParticipantEmail);
  const active = rows.filter(r => ACTIVE_STATI.indexOf(r.Status) >= 0);
  // v31.4: Warteliste/No-Show sind rücknehmbar — eine Nummer darauf ist belegt,
  // nicht frei. Vorher stand hier `activeBibs` (nur ACTIVE_STATI); damit galt
  // die Nummer einer Person auf der Warteliste als herrenlos, und die
  // Aufgabenliste hätte sie an jemand anderen weitergereicht, während die
  // Aktion „Startnummern zuteilen" (utils/b2runBibPool) sie sperrt. Beide
  // müssen über dieselbe Nummer dasselbe sagen — die Ableitung wird dadurch
  // ausschließlich STRENGER, nie großzügiger.
  const occupiedBibs = new Set(rows.filter(r => r.Status !== 'Abgemeldet').map(bibOf).filter(Boolean));

  // Nummern, die nur noch auf abgemeldeten Zeilen stehen. Bei mehreren
  // Kandidaten zählt die ZULETZT abgemeldete — sie war die letzte Halterin.
  const freeBib = new Map<string, SPRegistration>();
  for (const r of rows) {
    const b = bibOf(r);
    if (!b || r.Status !== 'Abgemeldet' || occupiedBibs.has(b)) continue;
    const prev = freeBib.get(b);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (!prev || ts((r as any).CancellationDate) > ts((prev as any).CancellationDate)) freeBib.set(b, r);
  }

  const byEmail = new Map<string, SPRegistration>();
  for (const r of rows) {
    const k = lc(r.ParticipantEmail);
    const prev = byEmail.get(k);
    if (!prev || (ACTIVE_STATI.indexOf(r.Status) >= 0 && ACTIVE_STATI.indexOf(prev.Status) < 0)) byEmail.set(k, r);
  }

  const withoutBib = active
    .filter(r => !bibOf(r))
    .sort((a, b) => ts(a.RegistrationDate) - ts(b.RegistrationDate));

  const out: StoredB2RunTodo[] = [];
  const claimed = new Set<string>();
  const nowIso = new Date().toISOString();
  const openBibs = Array.from(freeBib.entries())
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    .sort((a, b) => ts((a[1] as any).CancellationDate) - ts((b[1] as any).CancellationDate));

  // 1) Aufgezeichnete Nachrück-Kette zuerst — die einzige Paarung, die in den
  //    Daten steht.
  const stillOpen: Array<[string, SPRegistration]> = [];
  for (const [bib, from] of openBibs) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const succ = byEmail.get(lc((from as any).ReplacedByParticipantEmail));
    if (succ && ACTIVE_STATI.indexOf(succ.Status) >= 0 && !bibOf(succ) && !claimed.has(lc(succ.ParticipantEmail))) {
      claimed.add(lc(succ.ParticipantEmail));
      out.push({
        key: `transfer|${bib}|${lc(succ.ParticipantEmail)}`,
        kind: 'transfer', bib, certain: true, ts: nowIso,
        fromName: b2runNameOf(from), fromEmail: from.ParticipantEmail,
        toName: b2runNameOf(succ), toEmail: succ.ParticipantEmail,
        action: `Startnummer ${bib} beim Veranstalter von ${b2runNameOf(from)} auf ${b2runNameOf(succ)} ummelden.`,
      });
    } else {
      stillOpen.push([bib, from]);
    }
  }

  // 2) Übrige freie Nummern zeitlich zuordnen — Vorschlag, keine Feststellung.
  const pool = withoutBib.filter(r => !claimed.has(lc(r.ParticipantEmail)));
  const pairs = Math.min(stillOpen.length, pool.length);
  for (let i = 0; i < pairs; i++) {
    const [bib, from] = stillOpen[i];
    const p = pool[i];
    claimed.add(lc(p.ParticipantEmail));
    out.push({
      key: `assign|${bib}|${lc(p.ParticipantEmail)}`,
      kind: 'assign', bib, certain: false, ts: nowIso,
      fromName: b2runNameOf(from), fromEmail: from.ParticipantEmail,
      toName: b2runNameOf(p), toEmail: p.ParticipantEmail,
      action: `Startnummer ${bib} beim Veranstalter von ${b2runNameOf(from)} auf ${b2runNameOf(p)} ummelden — DEX hat hier keinen Nachrücker aufgezeichnet, bitte prüfen.`,
    });
  }

  // 3) Nummern ohne Abnehmer, 4) Personen ohne Nummer. Erst NACH 1 und 2:
  //    sonst stünde „abmelden" neben „nachmelden" für denselben Vorgang.
  for (let i = pairs; i < stillOpen.length; i++) {
    const [bib, from] = stillOpen[i];
    out.push({
      key: `unregister|${bib}`,
      kind: 'unregister', bib, certain: true, ts: nowIso,
      fromName: b2runNameOf(from), fromEmail: from.ParticipantEmail,
      action: `Startnummer ${bib} (${b2runNameOf(from)}) beim Veranstalter abmelden — es ist niemand da, der sie übernimmt.`,
    });
  }
  for (let i = pairs; i < pool.length; i++) {
    const p = pool[i];
    out.push({
      key: `register|${lc(p.ParticipantEmail)}`,
      kind: 'register', bib: '', certain: true, ts: nowIso,
      toName: b2runNameOf(p), toEmail: p.ParticipantEmail,
      action: `${b2runNameOf(p)} beim Veranstalter nachmelden — angemeldet, aber ohne Startnummer.`,
    });
  }
  return out;
}

/**
 * Festgehaltene und abgeleitete Aufgaben zu EINER Liste.
 *
 * Bei gleichem Schlüssel gewinnt die festgehaltene (sie trägt den
 * ursprünglichen Zeitpunkt). Zusätzlich wird für jede Aufgabe nachgesehen, ob
 * die Zielperson die Nummer in DEX inzwischen trägt — das entscheidet, ob der
 * Knopf „Nummer in DEX übertragen" noch etwas zu tun hat.
 */
export function mergeB2RunTodos(
  stored: StoredB2RunTodo[] | null | undefined,
  regs: SPRegistration[] | null | undefined
): B2RunTodo[] {
  const derived = deriveB2RunTodos(regs);
  const map = new Map<string, StoredB2RunTodo>();
  for (const d of derived) map.set(d.key, d);
  for (const s of (stored || [])) map.set(s.key, s);

  const byEmail = new Map<string, SPRegistration>();
  for (const r of (regs || [])) {
    const k = lc(r.ParticipantEmail);
    const prev = byEmail.get(k);
    if (!prev || (ACTIVE_STATI.indexOf(r.Status) >= 0 && ACTIVE_STATI.indexOf(prev.Status) < 0)) byEmail.set(k, r);
  }

  return Array.from(map.values())
    .map(t => {
      const toReg = t.toEmail ? byEmail.get(lc(t.toEmail)) : undefined;
      return { ...t, toReg, bibInDex: !!(toReg && t.bib && bibOf(toReg) === t.bib) };
    })
    .sort((a, b) => {
      const order: B2RunTodoKind[] = ['transfer', 'assign', 'unregister', 'register'];
      const d = order.indexOf(a.kind) - order.indexOf(b.kind);
      return d !== 0 ? d : (a.bib || '').localeCompare(b.bib || '');
    });
}

/**
 * v31.4: Festgehaltene Aufgaben ZUSAMMENFÜHREN statt ersetzen.
 *
 * `_b2runTodo` wurde bis v31.3 an jeder Schreibstelle komplett überschrieben.
 * Solange nur der Import schrieb, fiel das nicht auf; sobald eine zweite Stelle
 * Verpflichtungen anlegt (die Zuteilung), löscht der nächste Import genau die
 * Aufgaben, die noch niemand erledigt hat — der v30.54-Fehler ein zweites Mal,
 * diesmal selbst verursacht. Eine Verpflichtung ist ein Ereignis der
 * Vergangenheit: Sie darf nur verschwinden, wenn jemand sie ausdrücklich
 * entfernt (`removeKeys`), nie als Nebenwirkung eines Schreibvorgangs.
 *
 * Bei gleichem Schlüssel gewinnt `added` — es ist der frischere Stand.
 * Zusätzlich bleibt pro Nummer höchstens EINE transfer/assign-Aufgabe stehen
 * (die jüngere): Zwei davon wären im Aufgaben-Dialog zwei „Nummer in DEX
 * übertragen"-Knöpfe für dieselbe Nummer, und zwei Klicks sind zwei aktive
 * Vergaben derselben Startnummer.
 */
export function mergeStoredTodos(
  existing: StoredB2RunTodo[] | null | undefined,
  added: StoredB2RunTodo[] | null | undefined,
  removeKeys: string[] | null | undefined
): StoredB2RunTodo[] {
  const map = new Map<string, StoredB2RunTodo>();
  for (const t of (existing || [])) if (t && t.key) map.set(t.key, t);
  for (const t of (added || [])) if (t && t.key) map.set(t.key, t);
  for (const k of (removeKeys || [])) map.delete(k);

  // Ein unlesbarer Zeitstempel zählt hier als ALT (0) — anders als bei `ts()`
  // oben, wo „unbekannt" ans Ende sortiert. Eine kaputte Aufgabe darf keine
  // gültige verdrängen.
  const tsOf = (v: string | undefined): number => {
    const t = new Date(v || '').getTime();
    return isFinite(t) ? t : 0;
  };
  // Verglichen wird der ROHWERT der Nummer, nicht der Schlüssel aus
  // `b2runBibPool.bibKey`: Ein Import von dort würde einen Modul-Zyklus
  // aufmachen (b2runBibPool liest StoredB2RunTodo aus dieser Datei), und die
  // Regel zweimal hinzuschreiben wären zwei Wahrheiten. „0900" und „900" in
  // zwei Aufgaben bleiben deshalb hier zwei Einträge — die Zuteilung sperrt
  // diesen Fall ohnehin als mehrdeutig, bevor jemand etwas vergibt.
  const all = Array.from(map.values());
  const newestPerBib = new Map<string, StoredB2RunTodo>();
  for (const t of all) {
    if (t.kind !== 'transfer' && t.kind !== 'assign') continue;
    const b = (t.bib || '').trim();
    if (!b) continue;
    const prev = newestPerBib.get(b);
    if (!prev || tsOf(t.ts) >= tsOf(prev.ts)) newestPerBib.set(b, t);
  }
  return all.filter(t => {
    if (t.kind !== 'transfer' && t.kind !== 'assign') return true;
    const b = (t.bib || '').trim();
    if (!b) return true;
    const keep = newestPerBib.get(b);
    return !keep || keep.key === t.key;
  });
}

/**
 * v31.4: Haken aufräumen, zu denen es keine Aufgabe mehr gibt.
 *
 * `_b2runTodoDone` ist eine reine Schlüsselliste. Verschwindet die zugehörige
 * Aufgabe (die Nummer ist übertragen, die Person wieder angemeldet), bleibt der
 * Haken als Karteileiche liegen — und markiert eine später neu entstehende
 * Aufgabe zur selben Nummer sofort als erledigt. Genau das ist die Art Fehler,
 * die niemandem auffällt: Die Zeile steht dann unter „Erledigt", obwohl beim
 * Veranstalter nichts passiert ist.
 *
 * **`todos` muss die VOLLSTÄNDIGE Liste sein** — festgehalten UND abgeleitet
 * (also `mergeB2RunTodos(...)` bzw. `mergeStoredTodos(...)` plus
 * `deriveB2RunTodos(regs)`). Wer nur die festgehaltenen übergibt, löscht die
 * Haken zu allen abgeleiteten Aufgaben, und der Organizer hakt sie morgen
 * erneut ab.
 */
export function cleanDoneKeys(
  done: string[] | null | undefined,
  todos: StoredB2RunTodo[] | null | undefined
): string[] {
  const known: string[] = [];
  for (const t of (todos || [])) if (t && t.key) known.push(t.key);
  const out: string[] = [];
  for (const k of (done || [])) {
    if (typeof k !== 'string' || !k) continue;
    if (known.indexOf(k) < 0) continue;
    if (out.indexOf(k) < 0) out.push(k);
  }
  return out;
}

export const B2RUN_TODO_LABELS: Record<B2RunTodoKind, string> = {
  transfer: 'Ummelden',
  assign: 'Ummelden (bitte prüfen)',
  unregister: 'Abmelden',
  register: 'Nachmelden',
};
