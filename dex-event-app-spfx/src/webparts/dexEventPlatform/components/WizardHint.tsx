import * as React from 'react';

/**
 * v22.22: Einheitliche Hinweis-Box für den Event-Wizard.
 *
 * Aufbau immer gleich: fette Kopfzeile „Hinweis — <Überschrift>" (EN:
 * „Note — <title>"), darunter der Text. Eine Schriftgröße (0.8rem), eine
 * Farbwelt (Orange-Palette), volle Breite — egal an welcher Stelle im
 * Wizard. Interaktive Inhalte (Checkboxen, Buttons, Listen) gehören mit in
 * `children`. Kein Emoji, kein Icon (Klartext-Regel).
 */
export default function WizardHint(props: {
  isDe: boolean;
  title: string;
  children: React.ReactNode;
  /** Abstands-/Sonder-Styles der jeweiligen Einbaustelle (margin etc.). */
  style?: React.CSSProperties;
}): React.ReactElement {
  return (
    <div style={{
      width: '100%',
      boxSizing: 'border-box',
      padding: '12px 14px',
      borderRadius: 8,
      background: 'rgba(237,139,0,0.08)',
      border: '1px solid var(--dex-orange, #ed8b00)',
      fontSize: '0.8rem',
      color: 'var(--dex-gray-700)',
      lineHeight: 1.5,
      ...props.style,
    }}>
      <div style={{ fontWeight: 700, color: 'var(--dex-orange-dark, #b35a00)', marginBottom: 4 }}>
        {props.isDe ? 'Hinweis' : 'Note'} — {props.title}
      </div>
      <div>{props.children}</div>
    </div>
  );
}
