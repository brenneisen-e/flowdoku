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
import { useLocaleSafe } from '../context/LanguageContext';
import { inputLocaleTag } from '../utils/inputLocale';
// v24.65: Die gescopte Root-Klasse des Web Parts. Alle Basis-Styles (Schrift,
// Radio-/Checkbox-/Input-Styling, CSS-Variablen) liegen unter `.dexApp`. Ein per
// Portal an document.body gerendertes Modal liegt AUSSERHALB dieses Scopes —
// deshalb fehlte dem Modal-Inhalt das komplette App-Styling. Indem wir das
// Overlay mit `styles.dexApp` versehen, greift der gesamte gescopte Stylesheet
// wieder (Schrift, Inputs, Variablen, …) — der moderne Look ist zurück, und das
// Portal/Zoom bleibt erhalten.
import styles from './DexEventPlatform.module.scss';
// v31.2: Gemeinsame UI-Klassen (Karten, Chips, Kopf/Fuß …) — dieselbe Quelle
// wie im Wizard, siehe dexUi.ts.
import { ensureDexUiStyles } from './dexUi';
import { X } from './Icons';

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
  ensureDexUiStyles();
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

/* v31.10: Handy (bis 520 px). Auf einem 390-px-Schirm kostet der Rahmen mehr
   als der Inhalt: 16 px Außenabstand plus 26 px Karteninnenabstand lassen von
   390 nur 306 px für Text — jede zweite Zeile bricht zusätzlich um. Weniger
   Rand heißt hier also weniger Gedränge, nicht weniger Ruhe.
   Die Fußzeile verteilt ihre Knöpfe über die Breite, statt sie rechts
   untereinander auszufransen: bei drei Knöpfen lag der Primärknopf sonst
   unterhalb des Bildschirmrands und war nur über einen Rollvorgang
   erreichbar. Ein Knopf allein in seiner Zeile wächst auf die volle Breite.
   Die Regeln stehen hier statt in dexUi.ts, weil sie nur den Modal-Rahmen
   betreffen und das Overlay seine Maße per Inline-Style setzt — dagegen
   kommt eine Klasse nur mit !important an. */
@media (max-width: 520px) {
  .dex-modal-overlay { padding: 10px 8px !important; }
  .dex-modal-overlay .dex-modal-card-pad { padding: 16px 14px !important; }
  .dex-modal-overlay .dex-ui-modal-head { gap: 10px; }
  .dex-modal-overlay .dex-ui-modal-head-icon { width: 32px; height: 32px; border-radius: 10px; }
  /* Tippziel des Schließen-Knopfs: 32 px sind mit dem Finger zu wenig. */
  .dex-modal-overlay .dex-ui-modal-head > .dex-ui-iconbtn { width: 40px; height: 40px; }
  .dex-modal-overlay .dex-ui-modal-foot > .btn { flex: 1 1 45%; padding: 10px 12px !important; }
}
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
  /**
   * v31.2: Einheitlicher Dialog-Kopf — Titel, optional Untertitel und Symbol,
   * rechts der Schließen-Knopf. Vorher baute jeder Dialog seinen eigenen
   * `<h3>` mit eigenen Abständen; die Köpfe sahen in 40 Dialogen 40-mal
   * anders aus. Ohne `title` rendert der Kopf nicht (alte Aufrufer laufen
   * unverändert).
   */
  title?: React.ReactNode;
  subtitle?: React.ReactNode;
  icon?: React.ReactNode;
  /** v31.2: Fußzeile mit den Aktionen (rechtsbündig, Trennlinie oben). */
  footer?: React.ReactNode;
  /** v31.2: Schließen-Knopf im Kopf unterdrücken (z.B. bei Pflicht-Entscheidung). */
  hideClose?: boolean;
  /**
   * v31.3: Stapel-Ebene, Default 9999. Nur setzen, wenn ein Dialog AUS einem
   * anderen heraus geöffnet wird (Wizard-Vorschau aus dem Ticket-Dialog):
   * Bei gleicher Ebene entscheidet die Einhäng-Reihenfolge der Portale — das
   * geht heute gut, ist aber Zufall und nicht Absicht.
   */
  zIndex?: number;
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
  title,
  subtitle,
  icon,
  footer,
  hideClose,
  zIndex = 9999,
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

  // v30.60: Die Sprache für die nativen Datums-/Zeitfelder im Modal. Der
  // Context trägt auch durch ein Portal (React-Kontext folgt dem
  // Komponenten-Baum, nicht dem DOM); das Attribut muss aber trotzdem hier
  // gesetzt werden, weil das Overlay im DOM an document.body hängt und das
  // `lang` der App-Wurzel deshalb nicht erbt.
  const localeSafe = useLocaleSafe();
  const modalLang = inputLocaleTag(localeSafe === 'de');

  React.useEffect(() => {
    if (!open) return undefined;
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape' && dismissable) onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open, dismissable, onClose]);

  if (!open) return null;

  // v31.10: Die Styles auch beim Rendern sicherstellen, nicht nur im Effect.
  // Der Effect läuft NACH dem ersten Anstrich — bis dahin fehlten dem ersten
  // Dialog einer Sitzung Kopf-, Fuß- und (seit dieser Version) die
  // Handy-Regeln, der Dialog sprang also einmal sichtbar um. Der Aufruf ist
  // idempotent (ein getElementById) und kein Hook — dasselbe Muster wie in
  // ParticipantTable und EventListPage.
  ensureModalStyles();

  // v31.10: Ohne eigenen `padding`-Wert trägt die Karte ihre Abstände über
  // diese Klasse — nur so kann die Handy-Regel oben sie verkleinern, ohne
  // Dialogen dazwischenzufunken, die ihren Innenabstand bewusst selbst setzen.
  const hasOwnPadding = padding !== undefined && padding !== null;

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
      // v30.60: Wie die --dex-*-Variablen eine Zeile tiefer muss auch die
      // Sprache am Overlay stehen: Das Modal hängt per Portal an document.body
      // und liegt damit AUSSERHALB des App-Wurzelelements, das `lang` trägt.
      // Ohne das zeigen Datums- und Zeitfelder in Modals wieder MM/DD/YYYY und
      // die 12-Stunden-Uhr (siehe utils/inputLocale).
      lang={modalLang}
      style={{
        // v31.2: dunkles Blau statt reinem Schwarz und ein Hauch Unschärfe —
        // die Seite dahinter bleibt als Kontext erkennbar, der Dialog hebt
        // sich trotzdem klar ab.
        position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.5)',
        backdropFilter: 'blur(2px)', WebkitBackdropFilter: 'blur(2px)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        zIndex, padding: 16,
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
        className={`dex-ui-modal-card${hasOwnPadding ? '' : ' dex-modal-card-pad'}`}
        onClick={e => e.stopPropagation()}
        style={{
          background: '#fff', borderRadius: 18,
          padding: padding ?? '22px 26px',
          maxWidth, width: '100%',
          boxShadow: '0 24px 64px rgba(0,0,0,0.22), 0 0 0 1px rgba(0,0,0,0.04)',
          // v31.10: 100 % des Overlays statt `calc(100vh - 32px)`. Das Overlay
          // liegt fix auf dem ganzen Fenster; sein Innenraum ist genau die
          // Höhe, die der Karte zusteht. Der feste Abzug von 32 px musste den
          // Außenabstand erraten — auf dem Handy sind es seit dieser Version
          // 10 px, und die Karte hätte sonst Höhe verschenkt.
          maxHeight: '100%',
          overflowY: 'auto',
          // v31.10: Rollt der Dialog innen zu Ende, rollt sonst die Seite
          // dahinter weiter (Scroll-Chaining) — auf dem Handy fühlt sich das
          // an, als sei der Dialog weggerutscht. Es rollt entweder der Dialog
          // oder die Seite, nie beides.
          overscrollBehavior: 'contain',
          display: 'flex', flexDirection: 'column', gap: 14,
          boxSizing: 'border-box',
        }}
      >
        {title !== undefined && title !== null && (
          <div className="dex-ui-modal-head">
            {icon && <span className="dex-ui-modal-head-icon" aria-hidden="true">{icon}</span>}
            <div style={{ flex: 1, minWidth: 0 }}>
              <h3 className="dex-ui-modal-title">{title}</h3>
              {/* div statt p: Aufrufer geben auch Blöcke (JSX mit <p>) herein. */}
              {subtitle && <div className="dex-ui-modal-subtitle">{subtitle}</div>}
            </div>
            {!hideClose && (
              <button
                type="button"
                className="dex-ui-iconbtn"
                aria-label={localeSafe === 'de' ? 'Schließen' : 'Close'}
                title={localeSafe === 'de' ? 'Schließen' : 'Close'}
                disabled={!dismissable}
                onClick={onClose}
                style={{ marginTop: -4, marginRight: -6 }}
              >
                <X size={18} />
              </button>
            )}
          </div>
        )}
        {children}
        {footer && <div className="dex-ui-modal-foot">{footer}</div>}
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
