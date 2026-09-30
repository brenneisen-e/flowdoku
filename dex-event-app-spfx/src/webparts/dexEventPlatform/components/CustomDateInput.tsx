/**
 * v32.53: Datums-/Zeitfeld für Formular-Antworten — dieselbe Darstellung wie
 * auf der Anmeldeseite (react-datepicker, `dd.MM.yyyy`, nie nativ; s. CLAUDE.md
 * „Wer irgendwo ein neues Datums-/Zeitfeld baut: nie nativ"), dasselbe
 * Speicherformat (utils/customDateValue). Genutzt in „Meine Events → Angaben
 * bearbeiten".
 *
 * Lazy wie auf der Anmeldeseite: react-datepicker und StayRangePicker (bringt
 * `registerLocale('de')` und das Kalender-CSS mit) liegen im selben Chunk.
 */
import * as React from 'react';
import type { ReactDatePickerProps } from 'react-datepicker';
import { parseCustomDateValue, formatCustomDateValue } from '../utils/customDateValue';

const LazyDatePicker = React.lazy(async (): Promise<{ default: React.ComponentType<ReactDatePickerProps> }> => {
  const [dp] = await Promise.all([
    import('react-datepicker'),
    import('./StayRangePicker'),
  ]);
  return { default: dp.default as React.ComponentType<ReactDatePickerProps> };
});

export interface CustomDateInputProps {
  value: string;
  onChange: (next: string) => void;
  withTime?: boolean;
  isDe: boolean;
  error?: boolean;
  required?: boolean;
}

export function CustomDateInput({ value, onChange, withTime, isDe, error, required }: CustomDateInputProps): React.ReactElement {
  return (
    <React.Suspense fallback={<div className="form-input" aria-hidden="true" style={{ minHeight: 48 }} />}>
      <LazyDatePicker
        selected={parseCustomDateValue(value || '')}
        onChange={(d: Date | null) => onChange(formatCustomDateValue(d, !!withTime))}
        showTimeSelect={!!withTime}
        timeFormat="HH:mm"
        timeIntervals={15}
        timeCaption={isDe ? 'Uhrzeit' : 'Time'}
        dateFormat={withTime ? 'dd.MM.yyyy HH:mm' : 'dd.MM.yyyy'}
        locale={isDe ? 'de' : undefined}
        placeholderText={isDe ? (withTime ? 'TT.MM.JJJJ HH:MM' : 'TT.MM.JJJJ') : (withTime ? 'dd/mm/yyyy hh:mm' : 'dd/mm/yyyy')}
        className={error ? 'form-input dex-ui-input--error' : 'form-input'}
        wrapperClassName="dex-datepicker-wrapper"
        calendarClassName="dex-datepicker-calendar"
        popperPlacement="bottom-start"
        autoComplete="off"
        isClearable
        ariaRequired={required ? 'true' : undefined}
      />
    </React.Suspense>
  );
}

export default CustomDateInput;
