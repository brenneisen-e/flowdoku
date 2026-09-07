/**
 * v30.91: Anwesenheit je Programmpunkt — Spalte `AgendaCheckIns` (Note, JSON)
 * auf der Teilnehmerzeile. Konzept: docs/konzept-programmpunkte.md, 2.2.
 *
 *   { "<agendaItemId>": { "at": "2026-10-12T08:03:12Z", "by": "scanner@deloitte.de" } }
 *
 * Warum in der Teilnehmerzeile und nicht in einer eigenen Liste: Alles, was
 * der Check-in braucht (Name, Foto, Teilnehmer-ID, QR-Code), steht schon auf
 * dieser Zeile, die Rechte stimmen schon (wer heute einchecken darf, schreibt
 * heute schon `Status`), und die Auswertung liest die Liste, die das Organizer
 * Center ohnehin lädt. Ein Check-in ist ein MERGE auf die Zeile — derselbe
 * Mensch zweimal ist idempotent: der erste Zeitstempel bleibt.
 *
 * Nutzer-Entscheidung 07.09.2026: Ein Check-in an einem Punkt setzt NUR
 * diesen Punkt — kein automatischer Event-Check-in („ich möchte wissen, ob die
 * Person bei 1, 3, 5 da war").
 */
export interface AgendaCheckInMark {
  at: string;
  by: string;
  /**
   * v31.2: No-Show an GENAU diesem Punkt (Nutzer 07.09.2026: „man soll
   * gefragt werden, ob das für das ganze Event No-Show ist oder nur für den
   * Programmpunkt"). Eine Marke mit `noShow: true` ist KEINE Anwesenheit —
   * `parseAgendaCheckIns` blendet sie aus, damit Auswertung, Bescheinigung
   * und Zähler sie nie mitzählen. Wer die Zeile zurückschreibt, muss ALLE
   * Marken lesen (`parseAgendaMarks`), sonst löscht er die No-Shows.
   */
  noShow?: boolean;
}

export type AgendaCheckIns = Record<string, AgendaCheckInMark>;

/** v31.2: ALLE Marken (Anwesenheit UND No-Show) — für Lesen-Ändern-Schreiben. */
export function parseAgendaMarks(raw: string | undefined | null): AgendaCheckIns {
  try {
    const o = JSON.parse(raw || '{}');
    if (!o || typeof o !== 'object' || Array.isArray(o)) return {};
    const out: AgendaCheckIns = {};
    Object.keys(o).forEach(k => {
      const v = o[k];
      if (!k || !v || typeof v !== 'object') return;
      const at = typeof v.at === 'string' ? v.at : '';
      if (!at) return;
      const mark: AgendaCheckInMark = { at, by: typeof v.by === 'string' ? v.by : '' };
      if (v.noShow === true) mark.noShow = true;
      out[k] = mark;
    });
    return out;
  } catch { return {}; }
}

/** Nur Anwesenheiten. Defensiv — das Feld ist Freitext, auch Flows oder Hand-Edits könnten es berühren. */
export function parseAgendaCheckIns(raw: string | undefined | null): AgendaCheckIns {
  const all = parseAgendaMarks(raw);
  const out: AgendaCheckIns = {};
  Object.keys(all).forEach(k => { if (!all[k].noShow) out[k] = all[k]; });
  return out;
}

/** v31.2: Nur die No-Show-Marken je Punkt. */
export function parseAgendaNoShows(raw: string | undefined | null): AgendaCheckIns {
  const all = parseAgendaMarks(raw);
  const out: AgendaCheckIns = {};
  Object.keys(all).forEach(k => { if (all[k].noShow) out[k] = all[k]; });
  return out;
}

/** Zeitstempel als „HH:MM" in Berliner Lokalzeit — für Kacheln und Chips. */
export function formatMarkTime(iso: string): string {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
}

/**
 * Welcher Programmpunkt „läuft" gerade? Der erste, dessen Zeitfenster jetzt
 * enthält; sonst der nächste, der beginnt; sonst der letzte. Vorschlag für
 * das Check-in-Team, damit niemand aus zwanzig Punkten den richtigen suchen
 * muss — eine feste Wahl bleibt möglich.
 */
export function suggestCurrentAgendaItem<T extends { id: string; date: string; time: string; endTime?: string }>(items: T[], now: Date = new Date()): T | null {
  if (!items.length) return null;
  const toMs = (date: string, time: string): number => {
    if (!date) return NaN;
    const d = new Date(`${date}T${(time || '00:00')}:00`);
    return d.getTime();
  };
  const sorted = items.slice().sort((a, b) => (toMs(a.date, a.time) || 0) - (toMs(b.date, b.time) || 0));
  const n = now.getTime();
  for (const it of sorted) {
    const s = toMs(it.date, it.time);
    const e = it.endTime ? toMs(it.date, it.endTime) : s + 60 * 60 * 1000;
    if (isFinite(s) && n >= s && n <= e) return it;
  }
  const next = sorted.find(it => { const s = toMs(it.date, it.time); return isFinite(s) && s > n; });
  return next || sorted[sorted.length - 1];
}
