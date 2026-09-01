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
// v24.65: Die gescopte Root-Klasse des Web Parts. Alle Basis-Styles (Schrift,
// Radio-/Checkbox-/Input-Styling, CSS-Variablen) liegen unter `.dexApp`. Ein per
// Portal an document.body gerendertes Modal liegt AUSSERHALB dieses Scopes —
// deshalb fehlte dem Modal-Inhalt das komplette App-Styling. Indem wir das
// Overlay mit `styles.dexApp` versehen, greift der gesamte gescopte Stylesheet
// wieder (Schrift, Inputs, Variablen, …) — der moderne Look ist zurück, und das
// Portal/Zoom bleibt erhalten.
import styles from './DexEventPlatform.module.scss';

// v24.64: Deterministisches Modal-Button-Styling.
// Hintergrund (recherchiert): SPFx-CSS-Module sind auf den Web-Part-Container
// gescopt; ihre Stylesheets werden über den SPFx-Style-Loader geladen. Ein per
// ReactDOM.createPortal an document.body gerendertes Modal liegt AUSSERHALB
// dieses Containers — die `.btn`-Styles greifen dort nicht zuverlässig, sodass
// die Buttons als nackte Browser-Buttons rendern. Lösung: ein eigenes <style>
// direkt in document.head injizieren (reines DOM, KEIN SPFx-Loader). Mit
// `!important` + festen Werten unter `.dex-modal-overlay` greifen die
// Deloitte-Button-Styles garantiert dokumentweit — auch im Portal.
const MODAL_STYLE_ID = 'dex-modal-global-styles';
function ensureModalStyles(): void {
  if (typeof document === 'undefined') return;
  if (document.getElementById(MODAL_STYLE_ID)) return;
  const el = document.createElement('style');
  el.id = MODAL_STYLE_ID;
  el.textContent = `
.dex-modal-overlay button.btn,
.dex-modal-overlay .btn {
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: 8px !important;
  padding: 10px 24px !important;
  border: none !important;
  border-radius: 12px !important;
  font-size: 0.95rem !important;
  font-weight: 600 !important;
  font-family: inherit !important;
  line-height: 1.2 !important;
  cursor: pointer !important;
  text-decoration: none !important;
  transition: all 0.2s ease !important;
}
.dex-modal-overlay .btn-primary { background: #86bc25 !important; color: #ffffff !important; }
.dex-modal-overlay .btn-primary:hover { background: #6b9a1e !important; color: #ffffff !important; }
.dex-modal-overlay .btn-secondary { background: #e8e8e8 !important; color: #333333 !important; }
.dex-modal-overlay .btn-secondary:hover { background: #d1d1d1 !important; color: #333333 !important; }
.dex-modal-overlay .btn-danger { background: #666666 !important; color: #ffffff !important; }
.dex-modal-overlay .btn-danger:hover { background: #333333 !important; color: #ffffff !important; }
.dex-modal-overlay .btn-outline { background: transparent !important; border: 2px solid #86bc25 !important; color: #6b9a1e !important; }
.dex-modal-overlay .btn-outline:hover { background: #86bc25 !important; color: #ffffff !important; }
.dex-modal-overlay .btn:disabled { opacity: 0.55 !important; cursor: not-allowed !important; }
`;
  document.head.appendChild(el);
}

interface ModalProps {
  open: boolean;
  onClose: () => void;
  /** Maximalbreite der Card in px. Default 480. */
  maxWidth?: number;
  /** Wenn false, lassen sich Backdrop-Click und Escape ignorieren —
   *  nützlich während async-Submit-Operationen. Default true. */
  dismissable?: boolean;
  /**
   * v30.51: Nur den Backdrop-Klick abschalten, Escape aber behalten.
   *
   * Für Dialoge mit Eingabefeldern: Ein versehentlicher Klick daneben wirft
   * dort Getipptes weg, und genau das passiert leicht — wer den vorbelegten
   * Text markiert und die Maustaste einen Millimeter neben der Karte
   * loslässt, hat den Dialog geschlossen. Default true (unverändert).
   */
  backdropClose?: boolean;
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
  backdropClose = true,
  padding,
  ariaLabel,
  children,
}: ModalProps): React.ReactElement | null {
  // v24.64: Globale Modal-Button-Styles einmalig in document.head sicherstellen.
  React.useEffect(() => { ensureModalStyles(); }, []);

  /**
   * v30.51: Der Backdrop schließt nur, wenn die Maus AUF dem Backdrop
   * gedrückt UND losgelassen wurde.
   *
   * Ein DOM-`click` feuert auf dem gemeinsamen Vorfahren von mousedown- und
   * mouseup-Ziel. Wer im Dialog Text markiert und dabei über den Rand der
   * Karte hinauszieht — beim vorbelegten `https://` der Normalfall —, drückt
   * innen und lässt außen los: Der Klick landet dann auf dem Backdrop, und
   * der Dialog ging zu, obwohl niemand danebengeklickt hat. Das
   * `stopPropagation` auf der Karte hilft dagegen nicht, weil das Event die
   * Karte gar nicht erst berührt.
   */
  const downOnBackdropRef = React.useRef(false);

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
      className={`dex-modal-overlay ${styles.dexApp}`}
      onMouseDown={e => { downOnBackdropRef.current = e.target === e.currentTarget; }}
      onClick={e => {
        const wasDownOnBackdrop = downOnBackdropRef.current;
        downOnBackdropRef.current = false;
        if (!dismissable || !backdropClose) return;
        if (e.target !== e.currentTarget || !wasDownOnBackdrop) return;
        onClose();
      }}
      style={{
        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        zIndex: 9999, padding: 16,
        // v24.63: Theme-Variablen direkt am Overlay setzen — so erben alle
        // Modal-Inhalte (Buttons, Eingaben, Texte) die --dex-*-Farben auch dann,
        // wenn das Modal per Portal außerhalb des Web-Part-Containers liegt.
        ['--dex-black' as string]: '#000000',
        ['--dex-white' as string]: '#ffffff',
        ['--dex-green' as string]: '#86bc25',
        ['--dex-green-dark' as string]: '#6b9a1e',
        ['--dex-gray-100' as string]: '#f5f5f5',
        ['--dex-gray-200' as string]: '#e8e8e8',
        ['--dex-gray-300' as string]: '#d1d1d1',
        ['--dex-gray-400' as string]: '#a0a0a0',
        ['--dex-gray-500' as string]: '#808080',
        ['--dex-gray-600' as string]: '#666666',
        ['--dex-gray-700' as string]: '#444444',
        ['--dex-gray-800' as string]: '#333333',
        ['--dex-red' as string]: '#da291c',
        ['--dex-orange' as string]: '#ed8b00',
        ['--dex-orange-dark' as string]: '#b35a00',
        ['--dex-radius' as string]: '12px',
      } as React.CSSProperties}
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
  // WICHTIG (v24.63): Damit am body-Level die Modal-Styles greifen, MÜSSEN die
  // `.btn`-Klassen + die `--dex-*`-CSS-Variablen global (auf :root bzw. ohne
  // Container-Scope) definiert sein — siehe DexEventPlatform.module.scss
  // (:global-Block). Sonst rendern die Buttons als nackte Browser-Buttons.
  try {
    if (typeof document !== 'undefined' && document.body) {
      return ReactDOM.createPortal(overlay, document.body);
    }
  } catch { /* Fallback: inline rendern */ }
  return overlay;
}
