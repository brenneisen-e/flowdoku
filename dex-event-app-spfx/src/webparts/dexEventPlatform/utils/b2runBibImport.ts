/**
 * v30.48: Import der offiziellen B2Run-Startnummern-Liste.
 *
 * Der Veranstalter schickt die gemeldete Teilnehmerliste zurück — dieselbe
 * Datei, die DEX exportiert hat (`data/b2runKoeln.ts`), erweitert um die Spalte
 * **Startnummer**. Diese Nummer muss zurück in DEX, damit die Teilnehmer ihre
 * echte Nummer sehen.
 *
 * Der interessante Teil ist nicht das Zuordnen, sondern der **Nachrücker-Fall**:
 * Zwischen Meldung und Rücklauf melden sich Leute ab, andere rücken nach. Die
 * Startnummer der abgemeldeten Person gehört dann der nachgerückten — und der
 * Organizer muss beim Veranstalter ummelden.
 *
 * **Diese Zuordnung ist keine Schätzung.** Seit v17.15 schreibt
 * `promoteFirstWaitlistItem` sie in beide Richtungen mit:
 *   - auf der abgemeldeten Person: `ReplacedByParticipantEmail`
 *   - auf der nachgerückten Person: `ReplacedParticipantEmail` + `PromotedDate`
 * Wir folgen also einer festgehaltenen Kette, nicht einer Vermutung. Und weil
 * ein Nachrücker selbst wieder absagen kann, ist es wirklich eine KETTE:
 * A sagt ab → B rückt nach → B sagt ab → C rückt nach. Die Nummer von A gehört
 * dann C. Genau das leistet `resolveSuccessor` unten, mit Zyklus-Schutz —
 * A→B→A wäre sonst eine Endlosschleife, und Datenfehler dieser Art sind bei
 * wiederholtem Ab- und Anmelden nicht ausgeschlossen.
 */

import { SPRegistration } from '../services/EventService';

/** Status, die als „läuft mit" zählen. Abgemeldet ist bewusst nicht dabei. */
const ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

export interface BibRow {
  /** Startnummer aus der Veranstalter-Datei. */
  bib: string;
  email: string;
  firstName: string;
  lastName: string;
  /** Startblock laut Veranstalter (z.B. „17:00 Uhr Funstarter (Grün)"). */
  block: string;
  /** Zeilennummer in der Datei — für Fehlermeldungen, die man wiederfindet. */
  rowNo: number;
}

export type BibMatchKind =
  /** Person läuft, Nummer wird direkt zugeordnet. */
  | 'direct'
  /** Person ist abgemeldet, ein Nachrücker steht fest — Nummer wandert, UMMELDEN. */
  | 'transfer'
  /** Person ist abgemeldet, niemand ist nachgerückt — Nummer verfällt. */
  | 'orphan'
  /** E-Mail aus der Datei ist in DEX unbekannt. */
  | 'unknown';

export interface BibMatch {
  kind: BibMatchKind;
  row: BibRow;
  /** Die Person aus der Datei, sofern in DEX gefunden. */
  listed?: SPRegistration;
  /** Wer die Nummer bekommt (bei 'direct' = listed, bei 'transfer' = Nachrücker). */
  target?: SPRegistration;
  /** Kette der Abmeldungen bis zum Nachrücker — für die Anzeige „A → B → C". */
  chain?: SPRegistration[];
  /** Startblock in DEX weicht vom Veranstalter ab (informativ, kein Fehler). */
  blockMismatch?: { dex: string; file: string };
}

export interface BibImportReport {
  matches: BibMatch[];
  /** In DEX aktiv, steht aber NICHT in der Datei — hat also keine Nummer. */
  missingFromFile: SPRegistration[];
  /** Startnummern, die in der Datei doppelt vorkommen. */
  duplicateBibs: string[];
}

const lc = (s: string | undefined | null): string => (s || '').toLowerCase().trim();

/**
 * Folgt der Nachrück-Kette von einer abgemeldeten Person bis zu der Person,
 * die heute tatsächlich läuft.
 *
 * Gibt `null` zurück, wenn niemand nachgerückt ist oder die Kette in einer
 * ebenfalls abgemeldeten Person ohne weiteren Nachrücker endet — dann ist die
 * Nummer frei und niemand kann sie übernehmen.
 */
export function resolveSuccessor(
  start: SPRegistration,
  byEmail: Map<string, SPRegistration>
): { target: SPRegistration | null; chain: SPRegistration[] } {
  const chain: SPRegistration[] = [];
  const seen = new Set<string>([lc(start.ParticipantEmail)]);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let cur: any = start;
  // Harte Obergrenze zusätzlich zum `seen`-Set: Selbst bei kaputten Daten
  // endet die Schleife garantiert.
  for (let hop = 0; hop < 25; hop++) {
    const nextEmail = lc(cur.ReplacedByParticipantEmail);
    if (!nextEmail) return { target: null, chain };
    if (seen.has(nextEmail)) return { target: null, chain };
    seen.add(nextEmail);
    const next = byEmail.get(nextEmail);
    if (!next) return { target: null, chain };
    chain.push(next);
    if (ACTIVE_STATI.indexOf(next.Status) >= 0) return { target: next, chain };
    cur = next; // auch der Nachrücker hat abgesagt — weitersuchen
  }
  return { target: null, chain };
}

/**
 * Baut den Abgleich zwischen Veranstalter-Datei und DEX-Teilnehmerliste.
 * Schreibt nichts — das Ergebnis wird erst angezeigt und dann bestätigt.
 */
export function buildBibReport(rows: BibRow[], regs: SPRegistration[]): BibImportReport {
  const byEmail = new Map<string, SPRegistration>();
  for (const r of regs) {
    const key = lc(r.ParticipantEmail);
    if (!key) continue;
    // Bei zwei Zeilen zur selben Adresse gewinnt die AKTIVE — sonst gilt eine
    // alte Abmeldung als aktueller Stand (die Doppel-Adressen-Falle aus
    // CLAUDE.md: dieselbe Person unter SMTP und UPN).
    const prev = byEmail.get(key);
    if (!prev || (ACTIVE_STATI.indexOf(r.Status) >= 0 && ACTIVE_STATI.indexOf(prev.Status) < 0)) {
      byEmail.set(key, r);
    }
  }

  const seenBibs = new Set<string>();
  const duplicateBibs: string[] = [];
  const matches: BibMatch[] = [];
  const consumed = new Set<string>();

  for (const row of rows) {
    if (row.bib) {
      if (seenBibs.has(row.bib)) duplicateBibs.push(row.bib);
      seenBibs.add(row.bib);
    }
    const listed = byEmail.get(lc(row.email));
    if (!listed) { matches.push({ kind: 'unknown', row }); continue; }
    consumed.add(lc(row.email));

    const blockOf = (r: SPRegistration): string =>
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      String((r as any).StarterType || (r as any).PreferredStarterType || '');

    if (ACTIVE_STATI.indexOf(listed.Status) >= 0) {
      const dexBlock = blockOf(listed);
      const mismatch = dexBlock && row.block && row.block.toLowerCase().indexOf(dexBlock.toLowerCase()) < 0
        ? { dex: dexBlock, file: row.block }
        : undefined;
      matches.push({ kind: 'direct', row, listed, target: listed, blockMismatch: mismatch });
      continue;
    }

    // Abgemeldet → Nachrücker suchen.
    const { target, chain } = resolveSuccessor(listed, byEmail);
    if (target) {
      consumed.add(lc(target.ParticipantEmail));
      matches.push({ kind: 'transfer', row, listed, target, chain });
    } else {
      matches.push({ kind: 'orphan', row, listed, chain });
    }
  }

  const missingFromFile = regs.filter(r =>
    ACTIVE_STATI.indexOf(r.Status) >= 0 && !consumed.has(lc(r.ParticipantEmail)));

  return { matches, missingFromFile, duplicateBibs };
}

/**
 * Liest die Zeilen aus einer bereits geparsten Tabelle (erste Zeile = Header).
 *
 * Die Spalten werden über ihre ÜBERSCHRIFT gesucht, nicht über ihre Position:
 * Die Rücklauf-Datei des Veranstalters hat mehr Spalten als der Export und in
 * anderer Reihenfolge (im Rücklauf 2026 steht die Startnummer ganz vorn, die
 * E-Mail heißt dort „E-Mail for event registration"). Wer nach Position liest,
 * importiert beim nächsten Jahrgang stillschweigend die falsche Spalte.
 */
export function parseBibSheet(table: unknown[][]): { rows: BibRow[]; error?: string } {
  if (!table || table.length < 2) return { rows: [], error: 'Die Datei enthält keine Datenzeilen.' };
  // Kopfzeile suchen: die erste Zeile, in der sowohl eine Startnummer- als auch
  // eine E-Mail-Spalte steht. Die Veranstalter-Datei hat davor gelegentlich
  // eine Titelzeile.
  const norm = (v: unknown): string => String(v ?? '').replace(/­/g, '').trim().toLowerCase();
  let headerIdx = -1;
  let cBib = -1, cMail = -1, cFirst = -1, cLast = -1, cBlock = -1;
  for (let i = 0; i < Math.min(table.length, 10); i++) {
    const r = (table[i] || []).map(norm);
    const find = (pred: (s: string) => boolean): number => r.findIndex(pred);
    const bib = find(s => s.indexOf('startnummer') >= 0);
    const mail = find(s => s.indexOf('mail') >= 0);
    if (bib >= 0 && mail >= 0) {
      headerIdx = i; cBib = bib; cMail = mail;
      cFirst = find(s => s.indexOf('vorname') >= 0);
      cLast = find(s => s.indexOf('nachname') >= 0);
      cBlock = find(s => s.indexOf('startblock') >= 0);
      break;
    }
  }
  if (headerIdx < 0) {
    return { rows: [], error: 'Keine Kopfzeile mit „Startnummer" und einer E-Mail-Spalte gefunden. Bitte die Original-Datei des Veranstalters verwenden.' };
  }

  const rows: BibRow[] = [];
  for (let i = headerIdx + 1; i < table.length; i++) {
    const r = table[i] || [];
    const cell = (idx: number): string => (idx >= 0 ? String(r[idx] ?? '').trim() : '');
    const email = cell(cMail);
    const bib = cell(cBib);
    if (!email && !bib) continue; // Leerzeile
    rows.push({
      bib, email,
      firstName: cell(cFirst),
      lastName: cell(cLast),
      block: cell(cBlock),
      rowNo: i + 1,
    });
  }
  return { rows };
}
