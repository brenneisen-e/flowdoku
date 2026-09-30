/**
 * v32.45 — Wie viele Personen haben eine Frage schon beantwortet?
 *
 * Zwei Nutzer-Ansagen vom 30.09.2026 hängen daran:
 *  1. Wer bei einem bestehenden Event eine Frage entfernt, die schon
 *     beantwortet wurde, bekommt einen Danger-Zone-Dialog. Die Antworten
 *     bleiben zwar im CustomData der Zeilen stehen, aber ohne Feld zeigt sie
 *     keine Ansicht, kein Export und keine Mail mehr — für den Organizer sind
 *     sie weg.
 *  2. Wer eine Frage ERGÄNZT, während es schon Anmeldungen gibt, bekommt den
 *     Hinweis, die Bisherigen per Mail um das Nachtragen unter „Meine Events"
 *     zu bitten (dort steht die Frage seit v32.44 als „Noch offen").
 *
 * Gelesen wird über `getAllRegistrations` MIT `onHttpError` — ein Lesefehler
 * ist keine Null (CLAUDE.md, v30.67). Ist eine Liste nicht lesbar, liefert
 * das Ergebnis `gelesen: false`; der Aufrufer behandelt das als „unbekannt"
 * und warnt trotzdem, statt still zu löschen.
 *
 * Antworten stehen dort, wo angemeldet wurde — bei einem Klammer-Event oft auf
 * den Termin-Zeilen. Für eine Frage des Hauptevents werden deshalb Klammer-
 * UND Termin-Listen gelesen.
 */
import { EventService } from '../../../services/EventService';

/** Status, bei denen die Zeile eine gültige Anmeldung ist (inkl. Warteliste —
 *  auch dort wurde das Formular ausgefüllt). */
const AKTIV = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];

export interface AntwortStand {
  /** false = mindestens eine Liste war nicht lesbar → Zahlen sind Untergrenzen. */
  gelesen: boolean;
  /** Aktive Personen (nach E-Mail, über alle gelesenen Listen). */
  personen: number;
  /** Feld-Id → Anzahl Personen mit einer nicht-leeren Antwort. */
  antworten: Record<string, number>;
}

const cache = new Map<string, { at: number; p: Promise<AntwortStand> }>();
const FRISCH_MS = 60 * 1000;

/** Liest die Listen (Reihenfolge egal) und zählt je Feld die Personen mit Antwort. */
export function leseAntwortStand(subsiteUrls: string[]): Promise<AntwortStand> {
  const urls = subsiteUrls.filter(Boolean).map(u => u.replace(/\/+$/, '')).filter((u, i, a) => a.indexOf(u) === i).sort();
  const key = urls.join('|');
  const hit = cache.get(key);
  if (hit && Date.now() - hit.at < FRISCH_MS) return hit.p;
  const p = (async (): Promise<AntwortStand> => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    if (!ctx || urls.length === 0) return { gelesen: false, personen: 0, antworten: {} };
    const svc = new EventService(ctx);
    let gelesen = true;
    const personen = new Set<string>();
    const jeFeld: Record<string, Set<string>> = {};
    // Sequentiell — bei zwanzig Terminen wären parallele Aufrufe genau die
    // Last, die die Drosselung trifft.
    for (const url of urls) {
      let fehler = 0;
      const rows = await svc.getAllRegistrations(url, (status: number) => { fehler = status; });
      if (fehler) { gelesen = false; continue; }
      for (const r of rows) {
        if (AKTIV.indexOf(r.Status) < 0) continue;
        const mail = String(r.ParticipantEmail || '').toLowerCase().trim();
        if (!mail) continue;
        personen.add(mail);
        let data: Record<string, unknown> = {};
        try { data = r.CustomData ? JSON.parse(r.CustomData) : {}; } catch { data = {}; }
        Object.keys(data).forEach(k => {
          const v = data[k];
          if (v === undefined || v === null || String(v).trim() === '') return;
          (jeFeld[k] = jeFeld[k] || new Set<string>()).add(mail);
        });
      }
    }
    const antworten: Record<string, number> = {};
    Object.keys(jeFeld).forEach(k => { antworten[k] = jeFeld[k].size; });
    return { gelesen, personen: personen.size, antworten };
  })();
  cache.set(key, { at: Date.now(), p });
  // Fehlgeschlagene Läufe nicht eine Minute lang festhalten.
  p.then(s => { if (!s.gelesen) cache.delete(key); }).catch(() => cache.delete(key));
  return p;
}

/** Nach dem Speichern veraltet — beim nächsten Öffnen frisch lesen. */
export function vergissAntwortStand(): void { cache.clear(); }
