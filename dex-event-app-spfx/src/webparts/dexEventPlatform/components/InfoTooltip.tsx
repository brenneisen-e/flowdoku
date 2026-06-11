/**
 * Info-Tooltip Komponente.
 *
 * Ersetzt das bisherige `<span className="info-icon" title="...">i</span>` —
 * das Browser-Default-`title` zeigt nicht in allen Browsern zuverlässig
 * Tooltips an (vor allem nicht auf Touch und in Edge bei kleinen Hover-Targets).
 *
 * Diese Komponente rendert das gleiche `i`-Icon, blendet beim Hover/Focus aber
 * sofort ein gestyltes Tooltip-Pop unter/ueber dem Icon ein. Funktioniert auch
 * bei Tastatur-Navigation (focus) und Touch (click).
 *
 * v11.56: Das Tooltip-Popup wird via React.createPortal direkt an document.body
 * gehängt, damit es nicht von uebergeordneten `overflow:hidden`-Containern
 * (z.B. Karten auf der Registrierungs-Seite) abgeschnitten wird. Die Position
 * wird aus dem getBoundingClientRect() des Trigger-Icons berechnet (fixed).
 */
import * as React from 'react';
import { createPortal } from 'react-dom';

export interface InfoTooltipProps {
  /** Text oder JSX — bei JSX (z.B. <>...<strong>x</strong>...</>) werden
   *  fette Schlagwörter, Listen etc. korrekt gerendert. */
  text: React.ReactNode;
  /** Optional: rechts/links/oben statt default unten */
  placement?: 'top' | 'bottom' | 'right' | 'left';
}

export const InfoTooltip: React.FC<InfoTooltipProps> = ({ text, placement = 'top' }) => {
  const [open, setOpen] = React.useState(false);
  const triggerRef = React.useRef<HTMLSpanElement | null>(null);
  const [coords, setCoords] = React.useState<{ top: number; left: number; transform: string } | null>(null);

  const computePosition = React.useCallback((): void => {
    const el = triggerRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const cx = rect.left + rect.width / 2;
    const cy = rect.top + rect.height / 2;
    const gap = 8;
    // v12.4: Auf schmalen Screens (Mobile) wurde das Tooltip mit
    // translate(-50%) links abgeschnitten. Tooltip-Breite (max 480 /
    // viewport-Width-bezogen) berücksichtigen und so positionieren, dass
    // es nicht über den linken/rechten Rand hinausragt.
    const viewportW = typeof window !== 'undefined' ? window.innerWidth : 1024;
    const sideMargin = 12;
    const tooltipMaxW = Math.min(480, viewportW - sideMargin * 2);
    let top = 0; let left = 0; let transform = '';
    switch (placement) {
      case 'bottom':
        top = rect.bottom + gap;
        left = cx;
        transform = 'translateX(-50%)';
        break;
      case 'right':
        top = cy;
        left = rect.right + gap;
        transform = 'translateY(-50%)';
        break;
      case 'left':
        top = cy;
        left = rect.left - gap;
        transform = 'translate(-100%, -50%)';
        break;
      case 'top':
      default:
        top = rect.top - gap;
        left = cx;
        transform = 'translate(-50%, -100%)';
        break;
    }
    // Clamp links/rechts: wenn die effektive Tooltip-Box (zentriert auf
    // left, breite tooltipMaxW) den Viewport überschreitet, left auf die
    // Bandgrenze ziehen und die Transform-X-Komponente entsprechend
    // korrigieren, damit das Tooltip im sichtbaren Bereich bleibt.
    if (placement === 'top' || placement === 'bottom') {
      const halfW = tooltipMaxW / 2;
      if (left - halfW < sideMargin) {
        left = sideMargin + halfW;
      } else if (left + halfW > viewportW - sideMargin) {
        left = viewportW - sideMargin - halfW;
      }
    }
    setCoords({ top, left, transform });
  }, [placement]);

  React.useEffect(() => {
    if (!open) return;
    computePosition();
    const handler = (): void => computePosition();
    window.addEventListener('scroll', handler, true);
    window.addEventListener('resize', handler);
    return () => {
      window.removeEventListener('scroll', handler, true);
      window.removeEventListener('resize', handler);
    };
  }, [open, computePosition]);

  // Touch-Devices: ein Tap toggelt den Tooltip
  const handleClick = (e: React.MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    setOpen(o => !o);
  };

  const tooltipNode = open && coords && typeof document !== 'undefined' ? createPortal(
    <span
      role="tooltip"
      style={{
        position: 'fixed',
        top: coords.top,
        left: coords.left,
        transform: coords.transform,
        zIndex: 10000,
        padding: '12px 16px',
        background: 'rgba(40,40,40,0.96)',
        color: '#fff',
        // v11.59: explizite Font-Family — durch createPortal landet das Tooltip
        // an document.body und erbt NICHT mehr die Aptos/Open-Sans-Schrift des
        // App-Containers. Ohne diesen Override hat der Browser-Default (Times/
        // System-Sans) gegriffen und das Tooltip sah optisch fremd aus.
        fontFamily: 'Aptos, "Open Sans", "Segoe UI", Arial, Helvetica, sans-serif',
        fontSize: '0.82rem',
        fontWeight: 400,
        lineHeight: 1.55,
        borderRadius: 8,
        boxShadow: '0 6px 18px rgba(0,0,0,0.28)',
        // v9.24: breiter als vorher (320 -> 480) damit lange Hilfetexte
        // — vor allem die ausführlicheren Hints aus v9.17/v9.21 — nicht
        // mehr in einer schmalen Spalte hochkant umgebrochen werden.
        width: 'max-content',
        maxWidth: 'min(480px, calc(100vw - 24px))',
        minWidth: 'min(280px, calc(100vw - 24px))',
        whiteSpace: 'normal',
        textAlign: 'left',
        pointerEvents: 'none',
      }}
    >
      {text}
    </span>,
    document.body
  ) : null;

  return (
    <span
      ref={triggerRef}
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
      {tooltipNode}
    </span>
  );
};
