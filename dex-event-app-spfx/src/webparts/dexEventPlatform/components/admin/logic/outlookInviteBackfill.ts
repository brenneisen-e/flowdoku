/**
 * v31.36 — „Outlook-Einladungen nachziehen".
 *
 * ## Die Lücke, die das schließt
 *
 * Nutzer-Frage 14.09.2026: Ein Event steht auf „Keine Kommunikation", wird auf
 * „Nur Outlook-Termin" umgestellt, der Termin wird angelegt — bekommen die
 * bereits Angemeldeten ihn dann?
 *
 * Nein. Die Einladung entsteht bei der ANMELDUNG (`queueOutlookEvent(…,
 * 'Einladen')` in `registerForEvent`), nicht beim Anlegen des Termins.
 * `DEX_CreateOutlookEvent` erzeugt den Kalendereintrag; wer darauf steht,
 * entscheiden ausschließlich die `Einladen`-Zeilen der Queue. Wer sich
 * angemeldet hat, während Outlook aus war, hat nie eine bekommen — und es gab
 * keinen einzigen Pfad im Code, der das nachholt (geprüft: alle
 * `'Einladen'`-Aufrufe hängen an einem konkreten Ereignis — Anmeldung,
 * Team-Mitglied, Nachrücken, Organizer-Rolle; `createEvent` schreibt gar
 * keine). Auf dem frischen Termin standen damit nur die Organizer.
 *
 * ## Warum das eine eigene Aktion ist und nicht automatisch läuft
 *
 * Der Vorgang verschickt echte Kalendereinladungen an echte Teilnehmer. Ihn an
 * das Umlegen eines Schalters zu hängen hieße, dass ein Klick im Assistenten
 * unangekündigt hunderte Outlook-Mails auslöst. Deshalb: eigene Aktion, mit
 * Zahlen VOR der Rückfrage.
 *
 * ## Die Regeln, die hier drinstecken
 *
 *  - **Ein Lesefehler ist keine Null** (CLAUDE.md): Gelesen wird über
 *    `reloadRegistrations()`, den EINEN geprüften Nachlade-Pfad des Organizer
 *    Centers. Liefert er `null`, ist die Liste nicht lesbar — dann wird NICHTS
 *    verschickt. Aus einer leeren Liste auf „niemand angemeldet" zu schließen
 *    wäre hier besonders teuer: Die Aktion sähe erfolgreich aus und hätte
 *    niemanden eingeladen.
 *  - **Nur aktive Anmeldungen.** Warteliste bekommt keinen Termin (sie hat
 *    keinen Platz), Abgemeldete und No-Shows erst recht nicht.
 *  - **Externe Adressen werden übersprungen und GEZÄHLT.** Der Outlook-Flow
 *    lädt nur Deloitte-Adressen ein (v27.11: Member Firms zählen als intern);
 *    eine stille Auslassung wäre die Sorte Zahl, die später niemand erklären
 *    kann.
 *  - **Doppelte Adressen einmal.** Dieselbe Person kann unter zwei
 *    Schreibweisen in der Liste stehen (CLAUDE.md: die E-Mail ist der einzige
 *    Schlüssel und nicht eindeutig) — geschlüsselt wird kleingeschrieben.
 *  - **Fehlerzähler statt Schweigen.** `queueOutlookEvent` liefert `false`
 *    statt zu werfen; wer nur den `catch` prüft, meldet Erfolg über eine Zeile,
 *    die nie geschrieben wurde.
 */

import { EventService, SPRegistration } from '../../../services/EventService';
import { DeloitteEvent } from '../../../types';
import { isDeloitteInternalEmail } from '../../../utils/deloitteDomain';
import { parallelLimit } from '../../../utils/parallelLimit';

/** Stati, die einen Platz haben — nur die gehören in einen Kalender. */
const AKTIVE_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

export type BackfillStatus =
  /** Für dieses Event ist der Outlook-Termin abgeschaltet. */
  | 'outlook-aus'
  /** Es gibt noch gar keinen Termin, den man jemandem schicken könnte. */
  | 'kein-termin'
  /** Teilnehmerliste nicht lesbar — es wurde nichts verschickt. */
  | 'nicht-lesbar'
  /** Keine aktive Anmeldung mit interner Adresse. */
  | 'niemand'
  /** Rückfrage abgelehnt. */
  | 'abgebrochen'
  | 'fertig';

export interface BackfillErgebnis {
  status: BackfillStatus;
  /** Wie viele Einladungen in der Queue stehen. */
  eingeladen: number;
  /** Externe Adressen, die Outlook nicht einladen kann. */
  extern: number;
  /** Queue-Zeilen, die nicht geschrieben werden konnten. */
  fehler: number;
}

export interface BackfillCtx {
  svc: EventService;
  event: DeloitteEvent;
  /** Der geprüfte Nachlade-Pfad der Seite — `null` heisst „nicht lesbar". */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  /** Rückfrage mit den echten Zahlen; erst danach wird geschrieben. */
  frage: (_anzahl: number, _extern: number) => Promise<boolean>;
  onProgress?: (_done: number, _total: number) => void;
}

/**
 * Schreibt für jede aktive, interne Anmeldung dieses Termins eine
 * `Einladen`-Zeile in `DEX_Outlook`. Der Flow lädt die Person danach zum
 * bestehenden Kalendereintrag ein.
 *
 * Mehrfach aufzurufen ist ungefährlich: Wer schon auf dem Termin steht, wird
 * von Outlook erneut eingeladen und sieht den Termin unverändert — es entsteht
 * kein zweiter Eintrag.
 */
export async function runOutlookInviteBackfill(ctx: BackfillCtx): Promise<BackfillErgebnis> {
  const { svc, event, reloadRegistrations, frage, onProgress } = ctx;
  const leer: BackfillErgebnis = { status: 'fertig', eingeladen: 0, extern: 0, fehler: 0 };

  if (event.disableOutlook) return { ...leer, status: 'outlook-aus' };
  if (!event.outlookEventId && !event.calendarLink) return { ...leer, status: 'kein-termin' };

  const regs = await reloadRegistrations();
  if (regs === null) return { ...leer, status: 'nicht-lesbar' };

  const gesehen = new Set<string>();
  const intern: string[] = [];
  let extern = 0;
  for (const r of regs) {
    if (AKTIVE_STATI.indexOf(r.Status || '') < 0) continue;
    const mail = (r.ParticipantEmail || r.Title || '').trim();
    const lc = mail.toLowerCase();
    if (!lc || lc.indexOf('@') < 0 || gesehen.has(lc)) continue;
    gesehen.add(lc);
    if (isDeloitteInternalEmail(mail)) intern.push(mail);
    else extern++;
  }

  if (intern.length === 0) return { ...leer, status: 'niemand', extern };
  if (!(await frage(intern.length, extern))) return { ...leer, status: 'abgebrochen', extern };

  let fehler = 0;
  let fertig = 0;
  // Drei gleichzeitig: genug, um bei 400 Personen nicht minutenlang zu laufen,
  // wenig genug, um die Drosselung nicht selbst auszulösen (dieselbe Grenze
  // wie in `outlookQueue`).
  await parallelLimit(intern.map(mail => async () => {
    try {
      const ok = await svc.queueOutlookEvent(mail, event.id, event.title, 'Einladen');
      if (!ok) fehler++;
    } catch { fehler++; }
    fertig++;
    if (onProgress) onProgress(fertig, intern.length);
  }), 3);

  return { status: 'fertig', eingeladen: intern.length - fehler, extern, fehler };
}
