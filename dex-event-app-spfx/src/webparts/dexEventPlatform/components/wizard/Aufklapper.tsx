/**
 * v32.38: Aufklapper für optionale Wizard-Abschnitte (Nutzer-Ansage
 * 29.09.2026: Transfers und Programm in Schritt 3 einklappen). Hält seinen
 * Zustand selbst, damit die Schritte keine neuen Hooks brauchen. Zu, solange
 * nichts eingetragen ist; offen, sobald es Inhalt gibt — eine getroffene Wahl
 * darf nicht versteckt beginnen (dieselbe Regel wie beim CC der Einladung).
 */
import * as React from 'react';
import { ChevronDown } from '../Icons';
import { cx } from '../dexUi';

export interface AufklapperProps {
  label: React.ReactNode;
  /** Anzahl Einträge — steht zugeklappt rechts im Knopf. */
  anzahl?: number;
  startOffen?: boolean;
  children: React.ReactNode;
}

export const Aufklapper: React.FC<AufklapperProps> = ({ label, anzahl, startOffen, children }) => {
  const [offen, setOffen] = React.useState(!!startOffen);
  return (
    <div>
      <button type="button" className={cx('dex-ui-disclosure', offen && 'is-open')} aria-expanded={offen} onClick={() => setOffen(o => !o)}>
        <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
        {label}
        {typeof anzahl === 'number' && anzahl > 0 && <span className="dex-ui-disclosure-count">{anzahl}</span>}
      </button>
      {offen && <div className="dex-ui-disclosure-body">{children}</div>}
    </div>
  );
};

export default Aufklapper;
