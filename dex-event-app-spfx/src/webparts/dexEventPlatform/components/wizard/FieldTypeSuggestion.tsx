/**
 * v28.94: Aus `EventCreationPage` herausgeloest. Schlaegt anhand der
 * Beschriftung einen passenden Feldtyp vor.
 */
import * as React from 'react';
import { CustomFieldInput } from './customFieldInput';
import { labelLooksLikeDate, labelLooksLikeName, labelLooksLikeProfile } from '../../utils/fieldHeuristics';

/** v24.25/v24.28: Kleiner Hinweis im Feld-Editor. Drei Fälle, in dieser
 *  Priorität: (1) `profile` — das Feld wird ohnehin schon automatisch erfasst
 *  (kein Umstell-Button, nur Hinweis); (2) `date` — besser Kalender-Feld;
 *  (3) `person` — besser People-Picker (nur Haupt-Felder, `allowPerson`). */
export function FieldTypeSuggestion(props: {
  field: CustomFieldInput;
  isDe: boolean;
  allowPerson: boolean;
  disabled?: boolean;
  onApply: (type: CustomFieldInput['type']) => void;
}): React.ReactElement | null {
  const { field, isDe, allowPerson, disabled, onApply } = props;
  const label = (field.label || '').trim();
  if (!label) return null;
  let kind: 'profile' | 'date' | 'person' | null = null;
  if (labelLooksLikeProfile(label)) kind = 'profile';
  else if (labelLooksLikeDate(label) && field.type !== 'date' && field.type !== 'daterange') kind = 'date';
  // v29.21 (Audit): 'daterange' matcht die Heuristik (Anreise/Abreise/Check-in)
  // naturgemaess — der Tipp hätte das Feld auf 'date' zurueckgestuft und
  // damit rangeStart/rangeEnd/maxNights beim Save verworfen.
  else if (allowPerson && labelLooksLikeName(label) && field.type !== 'user' && field.type !== 'roommate') kind = 'person';
  if (!kind) return null;
  const body = kind === 'profile'
    ? (isDe
        ? <>Das wird vermutlich <strong>schon automatisch erfasst</strong> — Angaben wie Standort, Abteilung oder Unternehmenszugehörigkeit (z.B. Deloitte GmbH / Consulting GmbH) kommen aus dem Profil. Du musst sie meist nicht extra abfragen.</>
        : <>This is probably <strong>already collected automatically</strong> — details like location, department or company affiliation come from the profile. You usually don’t need to ask for them.</>)
    : kind === 'date'
    ? (isDe
        ? <>Das klingt nach einem <strong>Datum</strong>. Mit der Feldart <strong>„Datum“</strong> bekommen Teilnehmer einen Kalender-Auswähler (optional mit Uhrzeit) statt eines Freitextfelds.</>
        : <>This looks like a <strong>date</strong>. With the <strong>„Date“</strong> field type attendees get a calendar picker (optionally with time) instead of a free-text field.</>)
    : (isDe
        ? <>Das klingt nach einer <strong>Person</strong>. Mit der Feldart <strong>„Person“</strong> suchen Teilnehmer die Person direkt (mit Foto &amp; Standort), statt den Namen abzutippen.</>
        : <>This looks like a <strong>person</strong>. With the <strong>„Person“</strong> field type attendees search the person directly (with photo &amp; location) instead of typing the name.</>);
  const applyType: CustomFieldInput['type'] | null = kind === 'date' ? 'date' : kind === 'person' ? 'user' : null;
  const applyLabel = kind === 'date'
    ? (isDe ? 'Auf „Datum" umstellen' : 'Switch to „Date"')
    : (isDe ? 'Auf „Person" umstellen' : 'Switch to „Person"');
  return (
    <div style={{ marginTop: 8, marginLeft: 32, padding: '8px 12px', background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 8, display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
      <span style={{ flex: 1, minWidth: 200, fontSize: '0.78rem', color: 'var(--dex-gray-700)', lineHeight: 1.4 }}>
        <strong>{isDe ? 'Tipp' : 'Tip'}</strong> — „{label}“: {body}
      </span>
      {applyType && (
        <button
          type="button"
          disabled={disabled}
          onClick={() => onApply(applyType)}
          style={{ flexShrink: 0, background: 'var(--dex-green, #86bc25)', color: '#fff', border: 'none', borderRadius: 999, padding: '5px 12px', fontSize: '0.75rem', fontWeight: 700, cursor: disabled ? 'not-allowed' : 'pointer', opacity: disabled ? 0.6 : 1 }}
        >
          {applyLabel}
        </button>
      )}
    </div>
  );
}
