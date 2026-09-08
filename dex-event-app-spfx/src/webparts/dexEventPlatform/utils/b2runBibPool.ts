/**
 * v31.4: Die EINE Rechnung hinter „Startnummern zuteilen" — Vorrat, Kandidaten,
 * Sperren. Reines Rechenmodul: kein React, kein Service, kein Schreibvorgang.
 *
 * ## Warum eine eigene Datei und nicht ein Feld mehr in `b2runTodos`
 *
 * `deriveB2RunTodos` beantwortet die Frage „was muss ich beim Veranstalter noch
 * tun". Hier geht es um die andere Frage: „welche Nummer darf ich in DEX
 * jemandem geben". Beide lesen dieselben Zeilen, aber die zweite ist ein
 * Eingriff und braucht deshalb strengere Regeln als eine Anzeige. Die Aufgaben
 * gehen trotzdem mit ein — eine Nummer, die beim Veranstalter schon abgemeldet
 * ist, gehört dem Team nicht mehr.
 *
 * ## Was „frei" hier heißt — und was nicht
 *
 * Nutzer-Ansage vom 08.09.2026: „Ich habe ja heute importiert und nun bereits
 * 20 Nummern vergeben bzw. eingecheckt. Die dürfen nun nicht geändert werden."
 * Daraus folgen zwei Regeln, die den ganzen Rest tragen:
 *
 *  1. **Kandidat ist ausschließlich eine aktive Zeile mit LEERER Startnummer.**
 *     Eine Zeile, die eine Nummer trägt, wird nicht angefasst — nicht zum
 *     Korrigieren, nicht zum Lückenschließen.
 *  2. **Belegt ist jede Nummer, die eine nicht-abgemeldete Zeile trägt.** Nicht
 *     `ACTIVE_STATI`: Warteliste und No-Show sind rücknehmbar, die Person kann
 *     morgen wieder mitlaufen. Für den Vorrat gilt deshalb die WEITE Definition
 *     (alles außer 'Abgemeldet' ist belegt), für die Kandidaten die ENGE
 *     (`ACTIVE_STATI`). Die Asymmetrie ist Absicht: Beide Seiten irren
 *     zugunsten von „nichts vergeben".
 *
 * ## Warum kein gespeichertes Nummernregister
 *
 * Der Nummernkreis des Veranstalters ist fix, DEX kennt aber nur die Nummern,
 * die auf einer Zeile stehen. Ihn als Piggyback am Event zu speichern wäre ein
 * Read-Modify-Write ohne ETag (`patchEventOverridesValueEx`) und läge über
 * `EVENT_SELECT` im Browser jedes Teilnehmers — deshalb rechnet dieses Modul
 * ausschließlich auf dem, was in den Zeilen und in den festgehaltenen Aufgaben
 * steht. Es leitet ausdrücklich KEINEN Bereich aus Lücken ab: Nummern, die der
 * Import als 'unknown' gesehen hat, laufen beim Veranstalter auf reale
 * Personen und sind in DEX unsichtbar.
 *
 * ## Was gesperrt ist, ist gesperrt
 *
 * Ein Lesefehler ist keine Null (CLAUDE.md). Hier kommt derselbe Fall in einer
 * zweiten Gestalt vor, die `onHttpError` NICHT sieht: Die Teilnehmerliste ist
 * zeilenweise gesichert (ReadSecurity=2). Wer ohne „Manage Lists" liest,
 * bekommt ein fehlerfreies HTTP 200 mit der eigenen Zeile — und dann sieht
 * JEDE Nummer frei aus. Deshalb der Abgleich gegen die ungefilterte
 * `ItemCount` der Liste (Muster `getActiveCounts`): weniger gelesene Zeilen als
 * die Liste hat = Sicht unvollständig = gar keine Aussage.
 */

import { SPRegistration } from '../services/EventService';
import { StoredB2RunTodo, b2runNameOf } from './b2runTodos';
import { resolveSuccessor } from './b2runBibImport';

/** Status, die als „läuft mit" zählen — die ENGE Definition, nur für Kandidaten. */
const ACTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

const lc = (s: string | undefined | null): string => (s || '').toLowerCase().trim();

/** Rohwert der Startnummer. `Startnummer` steht (noch) nicht im Interface
 *  SPRegistration, kommt aber über `$select=*` mit — wie in `b2runTodos`. */
const rawBib = (r: SPRegistration): string =>
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  String((r as any).Startnummer || '').trim();

/** Gruppe/Startblock einer Zeile. Bei geteilten Kapazitäten steht sie in
 *  `StarterType`, bei Warteliste/Wunsch in `PreferredStarterType`. */
const groupOf = (r: SPRegistration | undefined): string =>
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  r ? String((r as any).StarterType || (r as any).PreferredStarterType || '').trim() : '';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const cancelDate = (r: SPRegistration): string => String((r as any).CancellationDate || '').trim();

/** Zeitpunkt einer Aufgabe. Anders als in `b2runTodos` zählt ein unlesbarer
 *  Zeitstempel hier als ALT (0) und nicht als jüngster — sonst gewinnt eine
 *  kaputte Zeile jeden Vergleich. */
const tsOf = (v: string | undefined): number => {
  const t = new Date(v || '').getTime();
  return isFinite(t) ? t : 0;
};

/** Startnummern sortieren sich nach Zahl, nicht nach Zeichen: sonst steht 100
 *  vor 99. */
const cmpBib = (a: string, b: string): number => a.localeCompare(b, 'de', { numeric: true });

/**
 * Vergleichsschlüssel einer Startnummer: trimmen, führende Nullen weg.
 *
 * v31.4: „0123" und „123" sind auf dem Zettel dieselbe Nummer — würde nur der
 * Rohwert verglichen, gälte eine belegte Nummer in der anderen Schreibweise als
 * frei. Der Rohwert bleibt für die Anzeige erhalten (er steht so auf dem
 * Startnummern-Zettel); zusammengefasst werden zwei Rohwerte NIE, sie sind
 * mehrdeutig und werden gesperrt (s. `BlockedBib.reason === 'ambiguous'`).
 */
export function bibKey(raw: unknown): string {
  const s = (raw === null || raw === undefined) ? '' : String(raw).trim();
  if (!s) return '';
  return s.replace(/^0+(?=.)/, '');
}

/** Eine Nummer, die niemand mehr trägt — mit dem Namen, der noch darauf steht. */
export interface FreeBib {
  /** Rohwert, so wie er in DEX bzw. in der Aufgabe steht. */
  bib: string;
  key: string;
  /** 'row' = eine abgemeldete Zeile in DEX trägt sie noch (dort ist sie zu
   *  leeren), 'todo' = nur aus einer festgehaltenen Aufgabe erinnert. */
  source: 'row' | 'todo';
  fromRow?: SPRegistration;
  fromName: string;
  fromEmail: string;
  /** Startblock der bisherigen Halterin. Bei Quelle 'todo' unbekannt (''). */
  block: string;
  /** v31.4 (Ergänzung zum Bauplan): die aufgezeichnete Nachrück-Kette wird
   *  schon hier aufgelöst. `proposeBibAssignments` bekommt laut Bauplan nur
   *  den Pool — ohne alle Zeilen kann es `resolveSuccessor` nicht selbst
   *  rufen, und ein zweiter Zeilen-Parameter wäre eine zweite Wahrheit. */
  chainTarget?: SPRegistration;
}

export interface BlockedBib {
  bib: string;
  reason: 'done' | 'occupied' | 'ambiguous' | 'multiHolder' | 'noCancelDate';
  detail: string;
}

export interface BibPool {
  /** `false`, sobald `hardBlock` gesetzt ist — dann ist alles andere leer. */
  ok: boolean;
  hardBlock: string;
  free: FreeBib[];
  blocked: BlockedBib[];
  candidates: SPRegistration[];
  skippedPeople: Array<{ reg: SPRegistration; reason: string }>;
  duplicatesActive: Array<{ bib: string; regs: SPRegistration[] }>;
  /** v31.4 (Ergänzung zum Bauplan): der Eingabewert wird mitgeführt, damit
   *  `proposeBibAssignments(pool)` mit EINEM Argument auskommt und nicht mit
   *  einer anderen Annahme rechnet als der Pool. */
  splitGroups: boolean;
}

export interface BibPoolInput {
  regs: SPRegistration[];
  /** Ungefilterte Anzahl der Listeneinträge (`ItemCount`). Negativ oder NaN =
   *  nicht ermittelbar = Sperre. */
  itemCount: number;
  /** Hat `getAllRegistrations` einen HTTP-Fehler gemeldet? */
  readFailed: boolean;
  /** `null` heißt „nicht lesbar/kein Array" — unbekannt, nicht leer. */
  storedTodos: StoredB2RunTodo[] | null;
  doneKeys: string[] | null;
  hasChildren: boolean;
  isChild: boolean;
  splitGroups: boolean;
}

/**
 * Zwei oder mehr nicht-abgemeldete Zeilen mit derselben Nummer.
 *
 * v31.4: Es gibt bis heute keine einzige Stelle im Projekt, die das bemerkt —
 * weder der Import noch der Übertragen-Knopf prüfen vor dem Schreiben, ob die
 * Nummer schon jemand trägt. Das ist die Sicherung gegen die Wege, die niemand
 * vorhergesehen hat: Fällt sie an, teilt die Aktion gar nichts mehr zu, statt
 * auf einer kaputten Grundlage weiterzurechnen.
 */
export function findDuplicateActiveBibs(
  regs: SPRegistration[] | null | undefined
): Array<{ bib: string; regs: SPRegistration[] }> {
  const byKey = new Map<string, SPRegistration[]>();
  for (const r of (regs || [])) {
    if (!r || r.Status === 'Abgemeldet') continue;
    const k = bibKey(rawBib(r));
    if (!k) continue;
    const arr = byKey.get(k) || [];
    arr.push(r);
    byKey.set(k, arr);
  }
  return Array.from(byKey.entries())
    .filter(e => e[1].length > 1)
    // Angezeigt wird der Rohwert der ersten Zeile; gruppiert wurde über den
    // Schlüssel, damit „0123" und „123" als das gemeldet werden, was sie sind:
    // dieselbe Nummer auf zwei Zeilen.
    .map(e => ({ bib: rawBib(e[1][0]) || e[0], regs: e[1] }))
    .sort((a, b) => cmpBib(a.bib, b.bib));
}

/**
 * Baut Vorrat, Kandidaten und Sperren in EINEM Durchgang.
 *
 * Reihenfolge ist Absicht: Erst die Fälle, in denen die Datenlage gar keine
 * Aussage erlaubt (harte Sperre, leerer Pool), dann die Belegung, dann der
 * Vorrat, zuletzt die Personen. Wer die Reihenfolge dreht, rechnet auf einer
 * Sicht, die er noch nicht geprüft hat.
 */
export function buildBibPool(a: BibPoolInput): BibPool {
  const rows = (a.regs || []).filter(r => !!r);
  const duplicatesActive = findDuplicateActiveBibs(rows);
  const empty = (hardBlock: string): BibPool => ({
    ok: false, hardBlock, free: [], blocked: [], candidates: [], skippedPeople: [],
    // Die Dublette wird bewusst MITGEGEBEN, auch wenn sonst alles leer ist:
    // Der Dialog muss beide Namen nennen können, sonst weiß niemand, was zu
    // klären ist.
    duplicatesActive, splitGroups: !!a.splitGroups,
  });

  if (a.readFailed) {
    return empty('Die Teilnehmerliste konnte gerade nicht vollständig gelesen werden — ohne sie weiß DEX nicht, welche Nummer schon jemand trägt. Es wird nichts zugeteilt.');
  }
  // Kein `itemCount < 0`, sondern die Umkehrung: so fällt auch NaN in die
  // Sperre statt still durchzurutschen.
  if (!(a.itemCount >= 0)) {
    return empty('DEX konnte nicht feststellen, wie viele Einträge diese Teilnehmerliste hat. Solange das unbekannt ist, lässt sich eine unvollständige Sicht nicht ausschließen — es wird nichts zugeteilt.');
  }
  if (rows.length < a.itemCount) {
    return empty(`DEX darf nur ${rows.length} von ${a.itemCount} Einträgen dieser Liste lesen. Solange die Sicht unvollständig ist, ist „frei" keine Aussage. Bitte mit Organizer-Rechten öffnen.`);
  }
  if (a.storedTodos === null || a.doneKeys === null) {
    return empty('Die Aufgaben beim Veranstalter oder die Haken dazu konnten nicht gelesen werden. Unbekannt heißt gesperrt, nicht „leer" — es wird nichts zugeteilt.');
  }
  if (a.hasChildren) {
    return empty('Dieses Event hat Termine. Startnummern können auf jedem Termin stehen, diese Ansicht sieht nur eine Liste — deshalb wird hier nichts zugeteilt.');
  }
  if (a.isChild) {
    return empty('Dieser Termin gehört zu einem Klammer-Event. Startnummern können auch auf den anderen Terminen stehen, diese Ansicht sieht nur eine Liste — deshalb wird hier nichts zugeteilt.');
  }
  if (duplicatesActive.length > 0) {
    const d = duplicatesActive[0];
    const names = d.regs.map(b2runNameOf).join(' UND bei ');
    const more = duplicatesActive.length > 1 ? ` (und ${duplicatesActive.length - 1} weitere)` : '';
    return empty(`Achtung: Startnummer ${d.bib} steht bei ${names}${more}. Bitte erst klären — solange das so ist, teilt DEX keine Nummer zu.`);
  }

  // --- Belegung -----------------------------------------------------------
  // Die WEITE Definition: alles außer 'Abgemeldet' hält die Nummer fest.
  // Warteliste, No-Show, leerer oder unbekannter Status sind rücknehmbar.
  const occupiedBy = new Map<string, SPRegistration>();
  for (const r of rows) {
    if (r.Status === 'Abgemeldet') continue;
    const k = bibKey(rawBib(r));
    if (k && !occupiedBy.get(k)) occupiedBy.set(k, r);
  }

  // Mehrdeutigkeit: verschiedene Rohwerte auf denselben Schlüssel. Erfasst wird
  // über ALLE Quellen (Zeilen und Aufgaben) — sonst fällt der Fall „auf der
  // Zeile 0123, in der Aufgabe 123" durch.
  const rawsByKey = new Map<string, string[]>();
  const noteRaw = (raw: string): void => {
    const k = bibKey(raw);
    if (!k) return;
    const arr = rawsByKey.get(k) || [];
    if (arr.indexOf(raw) < 0) arr.push(raw);
    rawsByKey.set(k, arr);
  };
  for (const r of rows) noteRaw(rawBib(r));
  for (const t of a.storedTodos) noteRaw(String(t && t.bib || '').trim());

  // Abgehakte „abmelden"-Aufgaben: die Nummer wurde beim Veranstalter
  // abgemeldet und gehört dem Team nicht mehr. Verglichen wird über den
  // Schlüssel, nicht über den Rohwert — ein Haken auf „0123" muss „123" sperren.
  const doneUnregisterKeys = a.doneKeys
    .filter(k => typeof k === 'string' && k.indexOf('unregister|') === 0)
    .map(k => bibKey(k.substring('unregister|'.length)))
    .filter(Boolean);

  // --- Vorrat -------------------------------------------------------------
  const free: FreeBib[] = [];
  const blocked: BlockedBib[] = [];
  const handled: string[] = []; // Schlüssel, über die schon entschieden wurde

  // Der erste zutreffende Grund gewinnt; die Reihenfolge geht von der härtesten
  // Tatsache (jemand trägt sie) zur Unbestimmbarkeit (wessen Name steht drauf).
  const decide = (key: string, bib: string, fromLabel: string): BlockedBib | null => {
    const holder = occupiedBy.get(key);
    if (holder) {
      return { bib, reason: 'occupied', detail: `Die Nummer steht noch auf der Zeile von ${b2runNameOf(holder)} mit Status ${holder.Status || 'unbekannt'} — die kann jederzeit zurückkommen.` };
    }
    const raws = rawsByKey.get(key) || [];
    if (raws.length > 1) {
      return { bib, reason: 'ambiguous', detail: `In DEX stehen ${raws.length} Schreibweisen derselben Nummer (${raws.join(' und ')}). Welche gemeint ist, lässt sich nicht entscheiden — sie bleiben alle gesperrt.` };
    }
    if (doneUnregisterKeys.indexOf(key) >= 0) {
      return { bib, reason: 'done', detail: `Beim Veranstalter bereits abgemeldet (${fromLabel}) — die Nummer gehört dem Team nicht mehr.` };
    }
    return null;
  };

  // Quelle A: abgemeldete Zeilen mit Nummer. Nur hier gibt es eine Zeile, die
  // beim Vergeben zu LEEREN ist — deshalb gewinnt A gegen B.
  const cancelledByKey = new Map<string, SPRegistration[]>();
  for (const r of rows) {
    if (r.Status !== 'Abgemeldet') continue;
    const raw = rawBib(r);
    const k = bibKey(raw);
    if (!k) continue;
    const arr = cancelledByKey.get(k) || [];
    arr.push(r);
    cancelledByKey.set(k, arr);
  }

  const byEmailActiveFirst = new Map<string, SPRegistration>();
  for (const r of rows) {
    const k = lc(r.ParticipantEmail);
    if (!k) continue;
    const prev = byEmailActiveFirst.get(k);
    // Bei zwei Zeilen zur selben Adresse gewinnt die AKTIVE — sonst gilt eine
    // alte Abmeldung als aktueller Stand (Doppel-Adressen-Falle, CLAUDE.md).
    if (!prev || (ACTIVE_STATI.indexOf(r.Status) >= 0 && ACTIVE_STATI.indexOf(prev.Status) < 0)) {
      byEmailActiveFirst.set(k, r);
    }
  }

  const cancelledKeys = Array.from(cancelledByKey.keys()).sort(cmpBib);
  for (const key of cancelledKeys) {
    const holders = (cancelledByKey.get(key) || []).slice();
    const bib = rawBib(holders[0]);
    handled.push(key);
    const stop = decide(key, bib, b2runNameOf(holders[0]));
    if (stop) { blocked.push(stop); continue; }
    if (holders.length > 1) {
      blocked.push({
        bib, reason: 'multiHolder',
        detail: `Mehrere abgemeldete Zeilen tragen diese Nummer (${holders.map(b2runNameOf).join(', ')}). Welcher Name auf dem Zettel steht, ist nicht bestimmbar — bitte von Hand vergeben.`,
      });
      continue;
    }
    const from = holders[0];
    if (!cancelDate(from)) {
      // v31.4: Ohne Abmeldedatum ist die Zeile nicht belastbar — dieselbe
      // Zeile würde in `deriveB2RunTodos` über die MAX_SAFE_INTEGER-Regel zur
      // vermeintlich jüngsten Halterin. Eine Zuteilung darf sich darauf nicht
      // stützen.
      blocked.push({
        bib, reason: 'noCancelDate',
        detail: `Auf der abgemeldeten Zeile von ${b2runNameOf(from)} fehlt das Abmeldedatum. Solange unklar ist, ob die Abmeldung wirklich stattgefunden hat, wird die Nummer nicht angeboten.`,
      });
      continue;
    }
    const succ = resolveSuccessor(from, byEmailActiveFirst);
    free.push({
      bib, key, source: 'row', fromRow: from,
      fromName: b2runNameOf(from), fromEmail: from.ParticipantEmail || '',
      block: groupOf(from),
      chainTarget: succ.target || undefined,
    });
  }

  // Quelle B: festgehaltene „abmelden"-Aufgaben. Nach einem Import steht die
  // Nummer einer abgemeldeten Person NUR dort — in DEX hat sie nie eine Zeile
  // getragen (s. Dateikopf von `b2runTodos`).
  const todoSeen: string[] = [];
  for (const t of a.storedTodos) {
    if (!t || t.kind !== 'unregister') continue;
    const bib = String(t.bib || '').trim();
    const key = bibKey(bib);
    if (!key || handled.indexOf(key) >= 0 || todoSeen.indexOf(key) >= 0) continue;
    todoSeen.push(key);
    const stop = decide(key, bib, t.fromName || 'unbekannt');
    if (stop) { blocked.push(stop); continue; }
    free.push({
      bib, key, source: 'todo',
      fromName: t.fromName || 'unbekannt', fromEmail: t.fromEmail || '',
      // Der Startblock der Vorgängerin steht in der Aufgabe nicht — unbekannt
      // heißt hier NICHT „passt", sondern wird beim Vorschlag berücksichtigt.
      block: '',
    });
  }

  free.sort((x, y) => cmpBib(x.bib, y.bib));
  blocked.sort((x, y) => cmpBib(x.bib, y.bib));

  // --- Personen -----------------------------------------------------------
  // Aufgaben, die eine Nummer für eine Person reservieren. Abgehakt zählt
  // ausdrücklich MIT: abgehakt heißt „beim Veranstalter umgemeldet", die Person
  // ist dort also bereits mit dieser Nummer geführt und bekommt erst recht
  // keine zweite.
  const reservedFor = new Map<string, StoredB2RunTodo>();
  for (const t of a.storedTodos) {
    if (!t || (t.kind !== 'transfer' && t.kind !== 'assign')) continue;
    const bib = String(t.bib || '').trim();
    const to = lc(t.toEmail);
    if (!bib || !to) continue;
    // Bei zwei Reservierungen für dieselbe Person gilt die jüngere — sie ist
    // der Stand, den der Veranstalter zuletzt gemeldet bekommen hat.
    const prev = reservedFor.get(to);
    if (!prev || tsOf(t.ts) >= tsOf(prev.ts)) reservedFor.set(to, t);
  }

  const rowsByEmail = new Map<string, SPRegistration[]>();
  for (const r of rows) {
    const k = lc(r.ParticipantEmail);
    if (!k) continue;
    const arr = rowsByEmail.get(k) || [];
    arr.push(r);
    rowsByEmail.set(k, arr);
  }

  const candidates: SPRegistration[] = [];
  const skippedPeople: Array<{ reg: SPRegistration; reason: string }> = [];
  const emails = Array.from(rowsByEmail.keys()).sort();
  for (const email of emails) {
    const mine = rowsByEmail.get(email) || [];
    const active = mine.filter(r => ACTIVE_STATI.indexOf(r.Status) >= 0);
    if (active.length === 0) continue; // Abgemeldet/Warteliste/No-Show: kein Kandidat, kein Befund
    const without = active.filter(r => !rawBib(r));
    if (without.length === 0) continue; // trägt bereits eine Nummer — bleibt unverändert
    const target = without[0];

    // v31.4: Dieselbe Person kann unter zwei Schreibweisen in der Liste stehen
    // (SMTP vs. UPN). Trägt EINE ihrer nicht-abgemeldeten Zeilen schon eine
    // Nummer, ist sie versorgt — eine zweite wäre eine doppelte Meldung.
    // Erkennbar ist das nur bei gleicher Adresse; zwei wirklich verschiedene
    // Schreibweisen kann diese Rechnung nicht verbinden.
    const other = mine.filter(r => r.Status !== 'Abgemeldet' && !!rawBib(r))[0];
    if (other) {
      skippedPeople.push({ reg: target, reason: `hat bereits Nummer ${rawBib(other)} auf einer zweiten Zeile` });
      continue;
    }
    const reserved = reservedFor.get(email);
    if (reserved) {
      skippedPeople.push({ reg: target, reason: `hat schon Nummer ${String(reserved.bib || '').trim()} offen — bitte erst übertragen` });
      continue;
    }
    candidates.push(target);
  }

  const tid = (r: SPRegistration): number =>
    (typeof r.TeilnehmerID === 'number' && isFinite(r.TeilnehmerID)) ? r.TeilnehmerID : Number.MAX_SAFE_INTEGER;
  candidates.sort((x, y) => (tid(x) - tid(y)) || (x.Id - y.Id));

  return { ok: true, hardBlock: '', free, blocked, candidates, skippedPeople, duplicatesActive, splitGroups: !!a.splitGroups };
}

export interface BibProposal {
  bib: string;
  key: string;
  free: FreeBib;
  target?: SPRegistration;
  /** 'chain' = in den Daten aufgezeichnet (und nur diese wird vorgehakt),
   *  'suggestion' = erschlossen, 'none' = kein Vorschlag. */
  stufe: 'chain' | 'suggestion' | 'none';
  reason: string;
  /** Vorschlag/Kette überquert den Startblock — am Lauftag ein Startproblem,
   *  kein Datenproblem. */
  blockMismatch: boolean;
}

/** Zwei Startblöcke gelten als gleich, wenn einer den anderen enthält: In DEX
 *  steht „Funstarter", beim Veranstalter „17:00 Uhr Funstarter (Grün)". */
const sameBlock = (a: string, b: string): boolean => {
  const x = lc(a), y = lc(b);
  if (!x || !y) return false;
  return x === y || x.indexOf(y) >= 0 || y.indexOf(x) >= 0;
};

/**
 * Ordnet jeder freien Nummer höchstens eine Person zu — als VORSCHLAG.
 *
 * Drei Vertrauensstufen, weil die Datenlage drei verschiedene Qualitäten hat:
 *
 *  - **'chain'**: `promoteFirstWaitlistItem` schreibt seit v17.15 mit, wer für
 *    wen nachgerückt ist. Das steht in den Daten, es ist nicht geraten — nur
 *    diese Stufe darf im Dialog vorgehakt sein.
 *  - **'suggestion'**: gleiche Gruppe, nächste Nummer, Reihenfolge nach
 *    TeilnehmerID. Eine plausible Annahme und mehr nicht.
 *  - **'none'**: es bleibt nichts übrig, oder es bliebe nur jemand aus einer
 *    anderen Gruppe bei geteilten Kapazitäten.
 *
 * Die Ketten laufen in einem ERSTEN Durchgang, damit eine Feststellung nie
 * einer Vermutung weicht: Sonst könnte ein früher einsortierter Vorschlag die
 * Person wegnehmen, für die eine Nummer nachweislich gedacht ist.
 */
export function proposeBibAssignments(pool: BibPool): BibProposal[] {
  if (!pool || !pool.ok) return [];
  const free = pool.free.slice().sort((a, b) => cmpBib(a.bib, b.bib));
  const cands = pool.candidates.slice();
  const usedIds: number[] = [];
  const byKey = new Map<string, BibProposal>();

  const mismatchWith = (f: FreeBib, r: SPRegistration): boolean =>
    !!f.block && !!groupOf(r) && !sameBlock(f.block, groupOf(r));

  // 1. Durchgang: aufgezeichnete Nachrück-Ketten.
  for (const f of free) {
    const t = f.chainTarget;
    if (!t) continue;
    if (usedIds.indexOf(t.Id) >= 0) continue;
    // Der Nachrücker muss auch heute noch Kandidat sein — hat er inzwischen
    // eine Nummer oder eine offene Reservierung, gilt die Kette nicht mehr.
    if (!cands.some(c => c.Id === t.Id)) continue;
    usedIds.push(t.Id);
    byKey.set(f.key, {
      bib: f.bib, key: f.key, free: f, target: t, stufe: 'chain',
      reason: `Aufgezeichnete Nachrück-Kette: ${f.fromName} hat sich abgemeldet, ${b2runNameOf(t)} ist nachgerückt. Das steht in den Daten, nicht geraten.`,
      blockMismatch: mismatchWith(f, t),
    });
  }

  // 2. Durchgang: Vorschlag je Gruppe, in der Reihenfolge der TeilnehmerID.
  for (const f of free) {
    if (byKey.get(f.key)) continue;
    const open = cands.filter(c => usedIds.indexOf(c.Id) < 0);
    const push = (p: BibProposal): void => { byKey.set(f.key, p); };
    if (open.length === 0) {
      push({ bib: f.bib, key: f.key, free: f, stufe: 'none', reason: 'Es ist niemand ohne Startnummer übrig — diese Nummer bleibt frei.', blockMismatch: false });
      continue;
    }
    const fit = f.block ? open.filter(c => sameBlock(f.block, groupOf(c))) : [];
    if (fit.length > 0) {
      const t = fit[0];
      usedIds.push(t.Id);
      push({
        bib: f.bib, key: f.key, free: f, target: t, stufe: 'suggestion',
        reason: 'Vorschlag, nicht belegt: gleiche Gruppe, nächste freie Nummer. Bitte prüfen.',
        blockMismatch: false,
      });
      continue;
    }
    if (pool.splitGroups) {
      // v31.4: Bei geteilten Kapazitäten wird gruppenfremd NIE vorgeschlagen —
      // ein falscher Startblock fällt erst am Lauftag auf und ist dann kein
      // Datenproblem mehr, sondern ein Startproblem. Wählbar bleibt die Nummer
      // im Dialog trotzdem, aber nur ausdrücklich.
      push({
        bib: f.bib, key: f.key, free: f, stufe: 'none',
        reason: f.block
          ? `Übrig sind nur Personen aus einer anderen Gruppe als „${f.block}" — anderer Startblock, bitte von Hand wählen.`
          : 'Zu dieser Nummer ist kein Startblock bekannt. Bei geteilten Gruppen wird sie deshalb nicht automatisch vorgeschlagen.',
        blockMismatch: true,
      });
      continue;
    }
    const t = open[0];
    usedIds.push(t.Id);
    push({
      bib: f.bib, key: f.key, free: f, target: t, stufe: 'suggestion',
      reason: f.block
        ? 'Anderer Startblock — bitte nur wählen, wenn du weißt, dass es passt.'
        : 'Vorschlag, nicht belegt: nächste freie Nummer in der Reihenfolge der Anmeldung. Bitte prüfen.',
      blockMismatch: mismatchWith(f, t),
    });
  }

  return Array.from(byKey.values()).sort((a, b) => cmpBib(a.bib, b.bib));
}
