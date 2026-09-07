/* FieldSelectInput — v30.95. Einfach-Auswahl eines Event-Felds im Organizer
 * Center (Teilnehmer bearbeiten, Hauptevent-Felder bearbeiten).
 *
 * Nutzer-Befund 07.09.2026: „warum kann ich den nicht als Herrengröße L
 * speichern?" Die Anmeldeseite speichert bei Auswahlfeldern mit Kategorien
 * (v26.75/v26.96, `optionCategories`) den Wert als „<Kategorie> <Option>",
 * also „Herrengröße L". Die Bearbeiten-Dialoge boten aber die rohen Optionen
 * an („L") — der gespeicherte Wert stand in keiner Option, und „Herrengröße L"
 * war nicht wählbar. Hier deshalb DIESELBE Kombibox wie auf der Anmeldeseite:
 * Kategorien als <optgroup>, Werte `${cat} ${opt}`, Optionen ohne Kategorie
 * ungruppiert dahinter. Ein gespeicherter Wert, der in keiner Option mehr
 * vorkommt (Altbestand „L", umbenannte Option), bleibt als eigener Eintrag
 * sichtbar — sonst zeigt das Feld leer, obwohl etwas gespeichert ist. */
import * as React from 'react';
import { EventSpecificField } from '../../../types';

export interface FieldSelectInputProps {
  field: EventSpecificField;
  value: string;
  onChange: (_v: string) => void;
  isDe: boolean;
}

export function categorizedSelectEntries(field: EventSpecificField): { groups: Array<{ cat: string; entries: Array<{ value: string; label: string }> }>; loose: Array<{ value: string; label: string }> } {
  const cats = field.optionCategories || [];
  const opts = field.options || [];
  const hasCats = cats.some(c => (c || '').trim());
  if (!hasCats) return { groups: [], loose: opts.filter(o => (o || '').trim()).map(o => ({ value: o, label: o })) };
  const distinct = Array.from(new Set(cats.map(c => (c || '').trim()).filter(Boolean)));
  const groups = distinct.map(cat => ({
    cat,
    entries: opts.map((opt, i) => ((cats[i] || '').trim() === cat && (opt || '').trim()) ? { value: `${cat} ${opt}`, label: `${cat} ${opt}` } : null).filter((x): x is { value: string; label: string } => !!x),
  }));
  const loose = opts.map((opt, i) => ((cats[i] || '').trim() === '' && (opt || '').trim()) ? { value: opt, label: opt } : null).filter((x): x is { value: string; label: string } => !!x);
  return { groups, loose };
}

export const FieldSelectInput: React.FC<FieldSelectInputProps> = ({ field, value, onChange, isDe }) => {
  const { groups, loose } = categorizedSelectEntries(field);
  const known = groups.some(g => g.entries.some(e => e.value === value)) || loose.some(e => e.value === value);
  return (
    <select className="form-select" value={value} onChange={e => onChange(e.target.value)} style={{ width: '100%' }}>
      <option value="">{isDe ? '— bitte wählen —' : '— please choose —'}</option>
      {value && !known && (
        <option value={value}>{value} {isDe ? '(gespeicherter Wert — nicht mehr in der Liste)' : '(stored value — no longer in the list)'}</option>
      )}
      {groups.map(g => (
        <optgroup key={g.cat} label={g.cat}>
          {g.entries.map(e => <option key={e.value} value={e.value}>{e.label}</option>)}
        </optgroup>
      ))}
      {loose.map(e => <option key={e.value} value={e.value}>{e.label}</option>)}
    </select>
  );
};

/** Sichtbarkeitsregel `showIf` eines Felds gegen die aktuellen Antworten
 *  auswerten — dieselbe Logik wie auf der Anmeldeseite (Multi-Select mit
 *  „ | ", Checkbox 'true'/'false'). `valueOf` liefert die Antwort zum
 *  Quell-Feld (per Feld-Id). */
export function fieldVisibleByShowIf(field: EventSpecificField, valueOf: (_fieldId: string) => string): boolean {
  if (!field.showIf || !field.showIf.fieldId) return true;
  const raw = (valueOf(field.showIf.fieldId) || '').trim();
  if (!raw) return false;
  const answers = raw.split(' | ').map(s => s.trim()).filter(Boolean);
  return answers.some(a => field.showIf!.values.indexOf(a) >= 0);
}

export default FieldSelectInput;
