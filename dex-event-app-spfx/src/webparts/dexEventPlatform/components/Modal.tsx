/**
 * Wiederverwendbares Modal-Wrapper-Komponente (v13.1).
 *
 * Vorher hatte jede Modal-Komponente (~17 Stück in der App) das gleiche
 * Backdrop + Wrapper-Layout selbst implementiert — mit leicht
 * abweichendem z-index, Padding, Border-Radius, Backdrop-Opacity.
 *
 * Diese Komponente kapselt das Standard-Verhalten:
 * - Fixed-Overlay mit halbtransparentem schwarzem Hintergrund.
 * - Inneres Card-Layout (weiß, abgerundet, mit Schatten) auf maximal
 *   `maxWidth` (default 480px) und vollem Width auf Mobile.
 * - Klick auf Backdrop schließt das Modal (es sei denn `dismissable`
 *   ist `false` — z.B. während eines laufenden Submit-Calls).
 * - Klick auf den Card-Body wird gestoppt, damit das Modal beim
 *   internen Klick nicht zugeht.
 * - Escape-Key schließt das Modal (außer dismissable=false).
 *
 * API:
 *   <Modal open onClose={...} maxWidth={520} dismissable={!busy}>
 *     ... eigentlicher Modal-Inhalt ...
 *   </Modal>
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';

interface ModalProps {
  open: boolean;
  onClose: () => void;
  /** Maximalbreite der Card in px. Default 480. */
  maxWidth?: number;
  /** Wenn false, lassen sich Backdrop-Click und Escape ignorieren —
   *  nützlich während async-Submit-Operationen. Default true. */
  dismissable?: boolean;
  /** Optional zusätzliches Card-Padding. Default '24px 28px'. */
  padding?: string | number;
  /** Aria-Label für Screen-Reader; Pflicht für barrierefreie Modals. */
  ariaLabel?: string;
  children: React.ReactNode;
}

export default function Modal({
  open,
  onClose,
  maxWidth = 480,
  dismissable = true,
  padding,
  ariaLabel,
  children,
}: ModalProps): React.ReactElement | null {
  React.useEffect(() => {
    if (!open) return undefined;
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape' && dismissable) onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open, dismissable, onClose]);

  if (!open) return null;

  const overlay = (
    <div
      role="dialog"
      aria-modal="true"
      aria-label={ariaLabel}
      onClick={() => { if (dismissable) onClose(); }}
      style={{
        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        zIndex: 9999, padding: 16,
      }}
    >
      <div
        onClick={e => e.stopPropagation()}
        style={{
          background: '#fff', borderRadius: 16,
          padding: padding ?? '24px 28px',
          maxWidth, width: '100%',
          boxShadow: '0 12px 48px rgba(0,0,0,0.18)',
          maxHeight: 'calc(100vh - 32px)',
          overflowY: 'auto',
          display: 'flex', flexDirection: 'column', gap: 14,
        }}
      >
        {children}
      </div>
    </div>
  );

  // v24.52: Per Portal an document.body rendern — so liegt das Modal AUSSERHALB
  // des Auto-Fit-Zoom-Containers und der position:fixed-Backdrop deckt immer den
  // ganzen Viewport ab (sonst blieben auf skalierten Screens Ränder frei).
  try {
    if (typeof document !== 'undefined' && document.body) {
      return ReactDOM.createPortal(overlay, document.body);
    }
  } catch { /* Fallback: inline rendern */ }
  return overlay;
}
