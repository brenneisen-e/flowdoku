/**
 * Speicherwerte in Wörter der Oberfläche — zweisprachig, an EINER Stelle.
 *
 * Status und Bewertung stehen in SharePoint als deutsche Auswahlwerte (`InArbeit`,
 * `Hoch`). Die Detailseite und die Pflegeliste zeigten sie bis v1.3 roh, auch in der
 * englischen Oberfläche („Geplant", „Hoch"); nur die Kachel übersetzte den Status
 * (Review 29.09.2026). Wer einen Wert anzeigt, geht hierdurch.
 */

import { Bewertung, UseCaseStatus } from '../types';

type Uebersetze = (de: string, en: string) => string;

export function statusText(status: UseCaseStatus, t: Uebersetze): string {
  if (status === 'Live') return t('Live', 'Live');
  if (status === 'InArbeit') return t('In Arbeit', 'In progress');
  if (status === 'Geplant') return t('Geplant', 'Planned');
  return t('Archiviert', 'Archived');
}

export function bewertungText(b: Bewertung, t: Uebersetze): string {
  if (b === 'Hoch') return t('Hoch', 'High');
  if (b === 'Mittel') return t('Mittel', 'Medium');
  if (b === 'Niedrig') return t('Niedrig', 'Low');
  return t('unbewertet', 'not assessed');
}
