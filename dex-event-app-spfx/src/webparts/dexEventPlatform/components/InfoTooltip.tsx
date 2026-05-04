/**
 * Info-Tooltip Komponente.
 *
 * Ersetzt das bisherige `<span className="info-icon" title="...">i</span>` —
 * das Browser-Default-`title` zeigt nicht in allen Browsern zuverlaessig
 * Tooltips an (vor allem nicht auf Touch und in Edge bei kleinen Hover-Targets).
 *
 * Diese Komponente rendert das gleiche `i`-Icon, blendet beim Hover/Focus aber
 * sofort ein gestyltes Tooltip-Pop unter/ueber dem Icon ein. Funktioniert auch
 * bei Tastatur-Navigation (focus) und Touch (click).
 */
import * as React from 'react';

export interface InfoTooltipProps {
  text: string;
  /** Optional: rechts/links/oben statt default unten */
  placement?: 'top' | 'bottom' | 'right' | 'left';
}

export const InfoTooltip: React.FC<InfoTooltipProps> = ({ text, placement = 'top' }) => {
  const [open, setOpen] = React.useState(false);
  // Touch-Devices: ein Tap toggelt den Tooltip
  const handleClick = (e: React.MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    setOpen(o => !o);
  };

  const tooltipPos: React.CSSProperties = (() => {
    switch (placement) {
      case 'bottom':
        return { top: 'calc(100% + 8px)', left: '50%', transform: 'translateX(-50%)' };
      case 'right':
        return { left: 'calc(100% + 8px)', top: '50%', transform: 'translateY(-50%)' };
      case 'left':
        return { right: 'calc(100% + 8px)', top: '50%', transform: 'translateY(-50%)' };
      case 'top':
      default:
        return { bottom: 'calc(100% + 8px)', left: '50%', transform: 'translateX(-50%)' };
    }
  })();

  return (
    <span
      style={{ position: 'relative', display: 'inline-flex', marginLeft: 8 }}
      onMouseEnter={() => setOpen(true)}
      onMouseLeave={() => setOpen(false)}
      onFocus={() => setOpen(true)}
      onBlur={() => setOpen(false)}
    >
      <span
        className="info-icon"
        tabIndex={0}
        role="button"
        aria-label="Info"
        onClick={handleClick}
        style={{
          display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
          width: 20, height: 20, borderRadius: '50%',
          border: '1.5px solid var(--dex-gray-500, #888)',
          fontSize: '0.7rem', fontWeight: 700, fontFamily: 'serif',
          cursor: 'help', color: 'var(--dex-gray-700, #555)',
          background: 'transparent', userSelect: 'none', flexShrink: 0,
          transition: 'background 0.15s, border-color 0.15s',
          ...(open ? { background: 'var(--dex-gray-100, #f0f0f0)', borderColor: 'var(--dex-gray-700, #555)' } : {}),
        }}
      >
        i
      </span>
      {open && (
        <span
          role="tooltip"
          style={{
            position: 'absolute',
            zIndex: 1500,
            padding: '12px 16px',
            background: 'rgba(40,40,40,0.96)',
            color: '#fff',
            fontSize: '0.82rem',
            fontWeight: 400,
            lineHeight: 1.55,
            borderRadius: 8,
            boxShadow: '0 6px 18px rgba(0,0,0,0.28)',
            // v9.24: breiter als vorher (320 -> 480) damit lange Hilfetexte
            // — vor allem die ausfuehrlicheren Hints aus v9.17/v9.21 — nicht
            // mehr in einer schmalen Spalte hochkant umgebrochen werden.
            width: 'max-content',
            maxWidth: 480,
            minWidth: 280,
            whiteSpace: 'normal',
            textAlign: 'left',
            pointerEvents: 'none',
            ...tooltipPos,
          }}
        >
          {text}
        </span>
      )}
    </span>
  );
};
