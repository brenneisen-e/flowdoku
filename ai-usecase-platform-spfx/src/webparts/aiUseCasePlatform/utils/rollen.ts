/**
 * Rollen: Anzeigename und Speicherwert an EINER Stelle.
 *
 * Nutzer-Entscheidung 29.09.2026: Es gibt Admins, **Use Case Organizer** und
 * User; Organizer legen Use Cases an und pflegen sie (Aufbau wie in DEX, dort
 * heißt die mittlere Rolle „Organizer").
 *
 * Bis v1.2 hieß die mittlere Rolle „Kurator", und `AIUC_Roles.Role` ist eine
 * Auswahlspalte mit genau den Werten `Admin`, `Kurator`, `User`. In der App
 * heißt sie jetzt `Organizer`, **gespeichert bleibt `Kurator`**. Der Grund ist
 * kein Bequemlichkeitsrest:
 *
 *  - Bestehende Zeilen tragen `Kurator`. Würde die App nur noch `Organizer`
 *    kennen, wäre jeder vorhandene Kurator beim Lesen ein `User` — still,
 *    ohne Fehlermeldung, und genau das ist der Fall, den diese Plattform von
 *    Anfang an vermeidet („ein Lesefehler befördert niemanden — und stuft
 *    auch keinen ab").
 *  - Ob SharePoint einen Wert schreibt, der nicht in den Auswahlwerten der
 *    Spalte steht, hängt vom Feld ab; auf einer bestehenden Liste ließe sich
 *    das nur durch Ändern der Spaltendefinition absichern — ein Eingriff in
 *    fremde Daten, für einen Namen.
 *
 * Deshalb kennt NUR diese Datei beide Schreibweisen. Der Rest der App sieht
 * `Organizer`.
 */

import { UserRole } from '../types';

/** Die Rollen in der Reihenfolge ihrer Rechte, höchste zuerst. */
export const ROLLEN_ALLE: UserRole[] = ['Admin', 'Organizer', 'User'];

/** Der Wert, der für `Organizer` in der Spalte `Role` steht (Altbestand). */
const SPEICHERWERT_ORGANIZER = 'Kurator';

/**
 * Rolle aus dem gespeicherten Wert. Unbekanntes wird zu `User` — die
 * vorsichtige Richtung: Wer nicht erkannt wird, bekommt weniger, nie mehr.
 */
export function rolleAusSpeicher(roh: string | undefined | null): UserRole {
  const v = (roh || '').trim();
  if (v === 'Admin') return 'Admin';
  if (v === 'Organizer' || v === SPEICHERWERT_ORGANIZER) return 'Organizer';
  return 'User';
}

/** Der Wert, der für diese Rolle in die Spalte `Role` geschrieben wird. */
export function rolleFuerSpeicher(rolle: UserRole): string {
  return rolle === 'Organizer' ? SPEICHERWERT_ORGANIZER : rolle;
}

/** Anzeigename. „Use Case Organizer" ist ausgeschrieben, wo Platz ist. */
export function rolleLabel(rolle: UserRole, kurz = false): string {
  if (rolle === 'Organizer') return kurz ? 'Organizer' : 'Use Case Organizer';
  return rolle;
}

/**
 * Die Rolle für die Anzeige (Fußzeile, Profil) — „unbekannt", wenn das Lesen der
 * Rollenliste fehlgeschlagen ist (429, 500, Netz). Ein Lesefehler wäre sonst als
 * „User" zu sehen, und ein Admin oder Organizer läse sich selbst als herabgestuft.
 * `forbidden` ist dagegen KEIN Fehler: Wer nicht in der Liste steht, darf sie nicht
 * lesen — und ist zu Recht User.
 */
export function rolleAnzeige(rolle: UserRole, lesestatus: 'loading' | 'ok' | 'forbidden' | 'error', t: (de: string, en: string) => string): string {
  return lesestatus === 'error' ? t('unbekannt', 'unknown') : rolleLabel(rolle);
}

/** Höhere Zahl = mehr Rechte. Für „nie implizit herabstufen". */
export function rolleRang(rolle: UserRole): number {
  return rolle === 'Admin' ? 3 : rolle === 'Organizer' ? 2 : 1;
}
