import * as React from 'react';

/**
 * v22.22: Einheitliche Hinweis-Box für den Event-Wizard.
 *
 * Aufbau immer gleich: fette Kopfzeile „Hinweis — <Überschrift>" (EN:
 * „Note — <title>"), darunter der Text. Eine Schriftgröße (0.8rem), eine
 * Farbwelt (Orange-Palette), volle Breite — egal an welcher Stelle im
 * Wizard. Interaktive Inhalte (Checkboxen, Buttons, Listen) gehören mit in
 * `children`. Kein Emoji, kein Icon (Klartext-Regel).
 *
 * v22.27: Boxen sind standardmäßig EINGEKLAPPT — sichtbar ist nur die
 * Kopfzeile (deshalb müssen die Überschriften für sich allein verständlich
 * sein); ein Klick darauf klappt den Text auf/zu. Ausnahme: Boxen mit
 * Pflicht-Interaktion (z.B. Bestätigungs-Checkbox) bekommen
 * `defaultOpen={true}`, damit der Anwender die Pflicht-Aktion nicht
 * übersieht.
 */
export default function WizardHint(props: {
  isDe: boolean;
  title: string;
  children: React.ReactNode;
  /** Abstands-/Sonder-Styles der jeweiligen Einbaustelle (margin etc.). */
  style?: React.CSSProperties;
  /** Aufgeklappt starten — nur für Boxen mit Pflicht-Interaktion nutzen. */
  defaultOpen?: boolean;
}): React.ReactElement {
  const [open, setOpen] = React.useState<boolean>(!!props.defaultOpen);
  return (
    <div style={{
      width: '100%',
      boxSizing: 'border-box',
      borderRadius: 8,
      background: 'rgba(237,139,0,0.08)',
      border: '1px solid var(--dex-orange, #ed8b00)',
      fontSize: '0.8rem',
      color: 'var(--dex-gray-700)',
      lineHeight: 1.5,
      ...props.style,
    }}>
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        aria-expanded={open}
        title={props.isDe
          ? (open ? 'Hinweis einklappen' : 'Hinweis aufklappen')
          : (open ? 'Collapse note' : 'Expand note')}
        style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10,
          width: '100%', boxSizing: 'border-box',
          background: 'none', border: 'none', cursor: 'pointer',
          padding: '10px 14px', margin: 0,
          font: 'inherit', textAlign: 'left',
          fontWeight: 700, color: 'var(--dex-orange-dark, #b35a00)',
        }}
      >
        <span>{props.isDe ? 'Hinweis' : 'Note'} — {props.title}</span>
        <span aria-hidden="true" style={{ flexShrink: 0, fontSize: '0.7rem', fontWeight: 400 }}>
          {open ? '▲' : '▼'}
        </span>
      </button>
      {open && (
        <div style={{ padding: '0 14px 12px' }}>{props.children}</div>
      )}
    </div>
  );
}
