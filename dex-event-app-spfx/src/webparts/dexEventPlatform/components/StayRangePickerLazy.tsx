import * as React from 'react';
import type { IStayRangePickerProps } from './StayRangePicker';

/**
 * v29.51 — `StayRangePicker`, nachgeladen statt mitgeliefert.
 *
 * Der Picker zieht `react-datepicker`, `date-fns/locale` und ein CSS mit.
 * Importiert wurde er bisher direkt von `RegistrationPage`, `MyEventsPage` und
 * `AssistantPage` — die ersten beiden hängen NICHT an `React.lazy`, also lag
 * das ganze Paket im Boot-Bundle. Gebraucht wird er aber nur an einer einzigen
 * Stelle: bei einem Formularfeld vom Typ `daterange` (Hotel-Zeitraum), das die
 * wenigsten Events überhaupt haben.
 *
 * Die eigene `Suspense`-Hülle ist Absicht: Die drei Aufrufer stehen teilweise
 * außerhalb der Route-Suspense in `DexEventPlatform`, ohne sie würde der erste
 * Render werfen.
 *
 * `parseStayValue` bleibt bewusst im Original-Modul — die Aufrufer davon
 * (`HotelPlanningPanel`, `HotelSetupWizard`) werden ohnehin nur nachgeladen.
 */
const StayRangePickerInner = React.lazy(() => import('./StayRangePicker'));

/** Platzhalter in Feldhöhe — sonst springt das Formular beim Nachladen. */
const Fallback: React.FC = () => (
  <div style={{
    height: 38, borderRadius: 8,
    background: 'var(--dex-gray-100, #f3f3f3)',
    border: '1px solid var(--dex-gray-200, #e1e1e1)',
  }} />
);

const StayRangePickerLazy: React.FC<IStayRangePickerProps> = (props) => (
  <React.Suspense fallback={<Fallback />}>
    <StayRangePickerInner {...props} />
  </React.Suspense>
);

export default StayRangePickerLazy;
