import * as React from 'react';

/**
 * v22.22: Einheitliche Hinweis-Box für den Event-Wizard.
 *
 * v22.27: standardmäßig EINGEKLAPPT — sichtbar ist nur die Kopfzeile, ein
 * Klick klappt auf/zu. Boxen mit Pflicht-Interaktion bekommen
 * `defaultOpen={true}`.
 *
 * v22.30: Einheitliche Farb-Logik im Wizard (verbindlich):
 *   - `variant="hint"` (Default): ORANGE — Hinweise/Warnungen,
 *     Kopfzeile „Hinweis — <Überschrift>".
 *   - `variant="description"`: GRAU — erklärende Beschreibungen („was
 *     passiert hier, wie funktioniert das"), Kopfzeile
 *     „Beschreibung — <Überschrift>".
 *   - Einstellungs-Sektionen (das, was der Organizer aktiv konfiguriert)
 *     liegen auf PASTELLGRÜNEN Karten (siehe `zebraS3Bg()` in
 *     EventCreationPage) — nicht über diese Komponente.
 * Kein Emoji, kein Icon (Klartext-Regel).
 */
export default function WizardHint(props: {
  isDe: boolean;
  title: string;
  children: React.ReactNode;
  /** Abstands-/Sonder-Styles der jeweiligen Einbaustelle (margin etc.). */
  style?: React.CSSProperties;
  /** Aufgeklappt starten — nur für Boxen mit Pflicht-Interaktion nutzen. */
  defaultOpen?: boolean;
  /** Farb-/Bedeutungs-Variante: 'hint' (orange) | 'description' (grau). */
  variant?: 'hint' | 'description';
}): React.ReactElement {
  const [open, setOpen] = React.useState<boolean>(!!props.defaultOpen);
  const isDesc = props.variant === 'description';
  const headColor = isDesc ? 'var(--dex-gray-700, #555)' : 'var(--dex-orange-dark, #b35a00)';
  const prefix = isDesc
    ? (props.isDe ? 'Beschreibung' : 'Description')
    : (props.isDe ? 'Hinweis' : 'Note');
  return (
    <div style={{
      width: '100%',
      boxSizing: 'border-box',
      borderRadius: 8,
      background: isDesc ? 'var(--dex-gray-50, #fafafa)' : 'rgba(237,139,0,0.08)',
      border: isDesc ? '1px solid var(--dex-gray-300, #d2d2d2)' : '1px solid var(--dex-orange, #ed8b00)',
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
          ? (open ? 'Einklappen' : 'Aufklappen')
          : (open ? 'Collapse' : 'Expand')}
        style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10,
          width: '100%', boxSizing: 'border-box',
          background: 'none', border: 'none', cursor: 'pointer',
          padding: '10px 14px', margin: 0,
          font: 'inherit', textAlign: 'left',
          fontWeight: 700, color: headColor,
        }}
      >
        <span>{prefix} — {props.title}</span>
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
