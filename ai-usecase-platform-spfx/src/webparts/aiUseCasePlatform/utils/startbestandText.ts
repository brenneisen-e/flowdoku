/**
 * Die Meldung zu einem Lauf von `seedStartUseCases` — EIN Text für Kachelwand und Studio.
 *
 * Die zwei Stellen hatten zwei Meldungen mit zwei Lücken: Die eine sagte bei Teilerfolg
 * nichts, die andere las den Fehlergrund aus einem State, der beim Lesen noch der von
 * VOR dem Lauf war („SharePoint sagt: kein Grund übermittelt", obwohl SharePoint einen
 * genannt hatte). Der Grund kommt jetzt aus dem Ergebnis.
 */

import { StartbestandErgebnis } from '../types';

type Uebersetze = (de: string, en: string) => string;

export interface StartbestandMeldung {
  /** `ok` = alles da, `warn` = Teilerfolg oder nichts angelegt, `info` = nichts zu tun. */
  art: 'ok' | 'warn' | 'info';
  text: string;
}

export function startbestandMeldung(e: StartbestandErgebnis, t: Uebersetze): StartbestandMeldung {
  if (e.fehler === 'laeuft') {
    return { art: 'info', text: t('Das Anlegen läuft schon — einen Moment.', 'Creating is already running — one moment.') };
  }
  if (e.fehler === 'rechte') {
    return { art: 'warn', text: t('Dafür brauchst du die Rolle Use Case Organizer.', 'You need the Use Case Organizer role for this.') };
  }
  if (e.fehlend === 0 && e.angelegt === 0) {
    if (e.fehler) {
      return { art: 'warn', text: t(`Die Liste ließ sich nicht lesen — es wurde nichts angelegt. SharePoint sagt: ${e.fehler}`,
        `The list could not be read — nothing was created. SharePoint says: ${e.fehler}`) };
    }
    return { art: 'info', text: t('Der Startbestand ist vollständig — es fehlt nichts.', 'The starter set is complete — nothing is missing.') };
  }
  if (e.angelegt === e.fehlend) {
    const basis = t(`${e.angelegt} Use Cases angelegt.`, `${e.angelegt} use cases created.`);
    return e.merker
      ? { art: 'ok', text: basis }
      : { art: 'warn', text: `${basis} ${t('Der Merker „schon befüllt" ließ sich nicht im Protokoll speichern — löschst du später alle Use Cases, legt die App den Startbestand beim nächsten Start wieder an.',
        'The "already filled" marker could not be saved to the log — if you later delete all use cases, the app creates the starter set again on the next start.')}` };
  }
  const grund = e.fehler || t('kein Grund übermittelt', 'no reason given');
  return {
    art: 'warn',
    text: t(`Nur ${e.angelegt} von ${e.fehlend} Use Cases angelegt. SharePoint sagt: ${grund}. „Fehlende anlegen" holt den Rest nach — schon vorhandene bleiben unberührt.`,
      `Only ${e.angelegt} of ${e.fehlend} use cases created. SharePoint says: ${grund}. "Add missing ones" creates the rest — existing ones stay untouched.`),
  };
}
