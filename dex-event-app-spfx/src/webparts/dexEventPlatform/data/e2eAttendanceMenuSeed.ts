/**
 * Seed-Daten fuer die E2E M&A Activation Session Munich:
 * Attendance (Dropdownfeld1) + Menu Preference (Dropdownfeld2) pro Teilnehmer,
 * gekeyed per E-Mail-Adresse.
 *
 * Quelle: `Teilnehmerliste (9).csv` (Export aus alter App, Stand 2026-04-15).
 *
 * Wird von `seedE2EAttendanceMenuByEmail()` in EventService genutzt:
 * pro Teilnehmer in der Subsite-Teilnehmerliste wird per Email gematcht
 * und die beiden Spalten nachgetragen, falls leer (idempotent).
 */
export interface E2EAttendanceMenuEntry {
  email: string;
  attendance: string;
  menuPreference: string;
}

export const E2E_ATTENDANCE_MENU_BY_EMAIL: E2EAttendanceMenuEntry[] = [
  { email: 'ankraus@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'shoerr@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Vegetarian' },
  { email: 'nifelten@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'mrauner@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'kganzen@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'jpreuss@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'tmenzler@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'benmartins@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'efietz@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'sdrees@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'pspyra@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'irnovikova@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'mpankow@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'mvonrueden@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'daschmitt@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Vegetarian' },
  { email: 'ajentgens@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'msailer@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'gogrundmann@deloitte.de', attendance: 'I will only attend the conference.', menuPreference: 'Standard Menu (meat)' },
  { email: 'bpeisl@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'janbakker@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'mahayn@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'micsimon@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'pschuettler@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'mjuellich@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'tbrunn@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'sgeberth@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Vegetarian' },
  { email: 'mstoever@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'stefannguyen@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Vegetarian' },
  { email: 'nelsmann@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'xschiessl@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'jschulzevellinghause@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'mnibler@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'gtillmann@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'slocker@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'adasilvafreitas@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'crhode@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'iugurlu@deloitte.de', attendance: 'I will attend at the conference and the dinner.', menuPreference: 'Standard Menu (meat)' },
  { email: 'hwilken@deloitte.de', attendance: 'I will only attend the conference.', menuPreference: 'Vegetarian' },
];
