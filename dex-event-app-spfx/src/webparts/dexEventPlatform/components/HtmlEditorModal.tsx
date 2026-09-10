/**
 * Wiederverwendbares Modal mit HTML-Editor + Live-Vorschau.
 *
 * Wird genutzt für:
 *   - Outlook-Termin-Body (Calendar Description)
 *   - E-Mail-Templates (Anmeldung, Warteliste, Abmeldung, Nachrücken)
 *
 * Layout (v31.2, drei Spalten mit ruhigen Überschriften): optional AKTIONEN
 * (`leftPanel`, z.B. QR-Versand) | BEARBEITEN (Kopf-Felder, Kopfbild-Aufklapper,
 * Text mit Toolbar) | VORSCHAU im echten Wrapper (Outlook-Termin bzw.
 * Deloitte-Mail). Nutzer-Befund 07.09.2026 zum QR-Mail-Dialog: „sieht völlig
 * überfordernd aus" — deshalb Feineinstellungen (Schrift der Überschriften,
 * Bildgröße) in Aufklappern, ein Primärknopf im Fuß, Erklärtext je Block eine Zeile.
 *
 * Die Vorschau wird in einem sandboxed iframe gerendert, damit das Wrapper-CSS
 * nicht mit dem App-CSS kollidiert.
 */
import * as React from 'react';
import { X, Mail, Calendar, FileText, Info, ChevronDown, RefreshCw, Link2, ImageIcon } from './Icons';
import { wrapTemplate, replacePlaceholders, replacePlaceholdersPlain, getCachedOrbBase64, getCachedLogoBase64 } from '../services/EmailTemplates';
// v20.4: modernes Confirm-Modal statt window.confirm.
import { useDialog } from '../context/DialogContext';
import { useIsMobile } from '../utils/useIsMobile';
// v31.2: Der Dialog war bis hier rein deutsch (mit einem englischen „Subject");
// jetzt beide Sprachen über isDe wie überall — ohne Provider-Zwang, weil der
// Editor auch aus Vorschauen heraus geöffnet wird.
import { useLocaleSafe } from '../context/LanguageContext';
// v31.2: Gemeinsame UI-Klassen (Chips, Iconknöpfe, Aufklapper, Kopf/Fuß) —
// dieser Dialog baut sein Overlay selbst und läuft nicht durch `Modal`,
// deshalb stellt er das Stylesheet beim Öffnen selbst sicher.
import { ensureDexUiStyles, cx } from './dexUi';
import LinkDialog from './LinkDialog';
// v31.7: Bild im Fließtext — Kompression, Grenzen und das <img> stehen an EINER
// Stelle (utils/inlineMailImage); dort steht auch, warum die Grenzen so hoch
// und nicht höher sind (DEX_Emails.Body ist eine Note-Spalte).
import {
  buildInlineImage, imageFileFromClipboard, inlineImageHtml, countInlineImages, charsToKb,
  INLINE_IMG_MAX_KB, MAIL_BODY_MAX_KB, MAIL_BODY_WARN_KB, InlineImageOutcome,
} from '../utils/inlineMailImage';
// v31.9.8: Positivliste fürs Einfügen aus Outlook/Word — warum sie so eng ist
// und was bewusst wegfällt, steht in der Datei.
import { sanitizePastedHtml } from '../utils/pasteHtmlSanitize';

// v9.40: 'plain' = nur HTML rendern, kein Mail-/Outlook-Wrapper. Wird für die
// Event-Beschreibung im Wizard genutzt — die landet 1:1 auf der Anmelde-Seite.
type PreviewMode = 'outlook' | 'email' | 'plain';

export interface HtmlEditorModalProps {
  open: boolean;
  onClose: () => void;
  title: string;
  value: string;
  onChange: (newValue: string) => void;
  previewMode: PreviewMode;
  emailSubject?: string;
  onEmailSubjectChange?: (s: string) => void;
  emailHeading?: string;
  onEmailHeadingChange?: (s: string) => void;
  /** v15.19: 2. Headline-Zeile (unter der Heading). Default {{EventTitle}}. */
  emailSubheading?: string;
  onEmailSubheadingChange?: (s: string) => void;
  emailHeadingColor?: string;
  /** v18.19: Mail-Überschrift Größe + Farbe einstellbar (pro Event). */
  emailHeadingFontSize?: string;
  onEmailHeadingColorChange?: (hex: string) => void;
  onEmailHeadingFontSizeChange?: (px: string) => void;
  /** v18.22: Überschrift fett/kursiv. */
  emailHeadingBold?: boolean;
  emailHeadingItalic?: boolean;
  onEmailHeadingBoldChange?: (b: boolean) => void;
  onEmailHeadingItalicChange?: (b: boolean) => void;
  /** v18.22: Unter-Überschrift frei formatierbar (Farbe/Größe/fett/kursiv). */
  emailSubheadingColor?: string;
  emailSubheadingFontSize?: string;
  emailSubheadingBold?: boolean;
  emailSubheadingItalic?: boolean;
  onEmailSubheadingColorChange?: (hex: string) => void;
  onEmailSubheadingFontSizeChange?: (px: string) => void;
  onEmailSubheadingBoldChange?: (b: boolean) => void;
  onEmailSubheadingItalicChange?: (b: boolean) => void;
  /** v18.73: Header-Bild (Event-Bild) Größe + Innenabstand — gilt für Mail-
   *  UND Outlook-Kopf, daher in beiden Vorschau-Modi sichtbar. */
  imageWidth?: number;
  imagePaddingV?: number;
  imagePaddingH?: number;
  onImageWidthChange?: (px: number) => void;
  onImagePaddingVChange?: (px: number) => void;
  onImagePaddingHChange?: (px: number) => void;
  /** Outlook-Termin: editierbare Überschrift (<h1>) */
  outlookHeading?: string;
  onOutlookHeadingChange?: (s: string) => void;
  /** Outlook-Termin: editierbare Unter-Überschrift (<h2>) */
  outlookSubheading?: string;
  onOutlookSubheadingChange?: (s: string) => void;
  /** v18.42: Betreff (Titel des Outlook-Termins) — bearbeitbar. Leer = Event-Titel. */
  outlookSubject?: string;
  onOutlookSubjectChange?: (s: string) => void;
  /** v18.44: Termin-Datum-Editor (zwei DatePicker, vom Parent gerendert mit der
   *  gleichen UI wie der Wizard). Leer-Werte zeigen das übernommene Event-Datum. */
  outlookDateEditor?: React.ReactNode;
  /** v18.44: Outlook-Ort — überschreibbar. Leer = aus „Ort & Programm" übernommen. */
  outlookLocationValue?: string;
  onOutlookLocationChange?: (s: string) => void;
  outlookLocationAuto?: string;
  /** v18.46: Standard-Body-Vorlage (HTML mit Platzhaltern) für „Standardtext
   *  laden". Wenn gesetzt, erscheint ein Reset-Button über dem Body-Editor. */
  defaultBodyHtml?: string;
  previewVars?: Record<string, string>;
  /** v22.18: Platzhalter, deren Wert ROHES HTML ist (z.B. {{QR_BLOCK}} im
   *  QR-Mail-Editor) — wird in der Vorschau OHNE HTML-Escaping ersetzt
   *  (previewVars escaped die Werte). */
  previewHtmlVars?: Record<string, string>;
  /** v22.19: Optionale schmale Spalte LINKS neben dem Editor (z.B. die
   *  QR-Versand-Aktionen) — Layout wird dann: Aktionen | Editor | Vorschau.
   *  So stehen Versand und Mail-Anpassung nebeneinander statt als zwei
   *  übereinander gestapelte Modals. */
  leftPanel?: React.ReactNode;
  insertableVars?: Array<{ key: string; label: string }>;
  logoBase64?: string;
  imageBase64?: string;
  /** v22.5: optionale „Briefumschlag"-Kopfzeile über der Mail-Vorschau —
   *  zeigt Empfänger („An") und Betreff, damit man die komplette Mail wie im
   *  Postfach sieht. Nur im previewMode 'email' relevant. */
  previewToLine?: string;
  previewSubjectLine?: string;
  /** Optional: zusätzlicher primärer Button im Footer (z.B. "Senden") */
  extraAction?: {
    label: string;
    onClick: () => void | Promise<void>;
    disabled?: boolean;
    icon?: React.ReactNode;
  };
  /** v11.40: Optionaler React-Knoten oberhalb von Subject/Überschrift im
   *  Editor — z.B. für eine Ziel-Auswahl im Einladungsmail-Modal. */
  headerExtra?: React.ReactNode;
  /** v28.7: Vorlagen-Chips über dem Body-Editor (z.B. Beschreibungs-
   *  Vorschläge aus dem Wizard). Klick ersetzt den Editor-Inhalt — mit
   *  Rückfrage, wenn schon Text drinsteht. */
  bodyTemplates?: Array<{ key: string; label: string; html: string; title?: string }>;
  /** v28.7: Beschriftung über den Vorlagen-Chips. */
  bodyTemplatesLabel?: string;
}

// v18.20: px-basierte Auswahl (wie in Word). v18.23: nur px, keine
// „Klein/Groß"-Beiwörter — die reine px-Angabe ist eindeutiger.
const FONT_SIZE_OPTIONS: number[] = [10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 40, 48];

const COLORS: string[] = [
  '#000000', '#555555', '#86bc25', '#0076a8', '#ed8b00', '#c9302c', '#6b21a8', '#0d6efd',
];

const isHex6 = (v: string): boolean => /^#[0-9a-fA-F]{6}$/.test(v);

// v18.73: Ganzzahl aus einem Number-Input parsen und auf [0, max] begrenzen
// (Untergrenze 0 — kleinere Werte während des Tippens erlaubt; das Rendering
// fällt bei 0 auf die Default-Werte zurück). Bei leerer/ungültiger Eingabe
// kommt der Fallback.
const clampInt = (v: string, max: number, fallback: number): number => {
  const n = parseInt(v, 10);
  if (isNaN(n)) return fallback;
  return Math.max(0, Math.min(max, n));
};

// v18.22: Farbwahl mit Swatches + nativem Farb-Picker + freiem Hex-Code.
// `value` = aktuell aktive Farbe (für Highlight + Picker-Startwert).
const ColorControl: React.FC<{ value: string; onChange: (_hex: string) => void }> = ({ value, onChange }) => {
  const [hexDraft, setHexDraft] = React.useState(value);
  // v31.10: Der native Farbwähler ist dasselbe Tippziel wie ein Farbpunkt und
  // wächst deshalb mit ihm mit.
  const isMobile = useIsMobile();
  React.useEffect(() => { setHexDraft(value); }, [value]);
  const commitHex = (h: string): void => {
    let v = (h || '').trim();
    if (v && !v.startsWith('#')) v = '#' + v;
    if (/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(v)) onChange(v.toLowerCase());
    else setHexDraft(value);
  };
  // v31.2: Swatches als runde Iconknöpfe mit Hover; die aktive Farbe trägt
  // einen dunklen Ring statt eines dickeren Rands (der verschob die Reihe).
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 2, flexWrap: 'wrap' }}>
      {COLORS.map(c => (
        <ColorDot key={c} color={c} active={(value || '').toLowerCase() === c.toLowerCase()} onPick={onChange} />
      ))}
      <input
        type="color"
        value={isHex6(value) ? value : '#000000'}
        onChange={e => onChange(e.target.value.toLowerCase())}
        title="Freie Farbe wählen / Pick any color"
        style={isMobile ? { ...colorPickerStyle, width: 40, height: 40 } : colorPickerStyle}
      />
      <input
        className="dex-ui-input dex-ui-input--sm"
        value={hexDraft}
        onChange={e => setHexDraft(e.target.value)}
        onBlur={() => commitHex(hexDraft)}
        onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); commitHex(hexDraft); } }}
        placeholder="#RRGGBB"
        maxLength={7}
        style={hexInputStyle}
      />
    </span>
  );
};

// v31.2: Ein Farbpunkt in der Palette — 26 px rund, Hover über `dex-ui-iconbtn`.
// `onMouseDown` verhindert, dass der Klick die Editor-Auswahl kollabiert.
// v31.10: Auf dem Handy 40 px, weil ein 26-px-Ziel mit dem Finger zur Lotterie
// wird — acht Farben nebeneinander heißt sonst: dreimal danebengetippt und
// dreimal die falsche Farbe im Text.
const ColorDot: React.FC<{ color: string; active: boolean; onPick: (_hex: string) => void }> = ({ color, active, onPick }) => {
  const isMobile = useIsMobile();
  const box = isMobile ? 40 : 26;
  return (
    <button
      type="button"
      className="dex-ui-iconbtn"
      title={color}
      onMouseDown={e => e.preventDefault()}
      onClick={() => onPick(color)}
      style={{ width: box, height: box }}
    >
      <span style={{
        display: 'block', width: isMobile ? 20 : 16, height: isMobile ? 20 : 16, borderRadius: '50%', background: color,
        boxShadow: active
          ? '0 0 0 2px #fff, 0 0 0 3.5px var(--dex-gray-700, #444)'
          : 'inset 0 0 0 1px rgba(0,0,0,0.12)',
      }} />
    </button>
  );
};
const colorPickerStyle: React.CSSProperties = { width: 26, height: 26, padding: 0, border: '1px solid var(--dex-gray-200, #e8e8e8)', borderRadius: 8, cursor: 'pointer', background: '#fff', marginLeft: 4 };
const hexInputStyle: React.CSSProperties = { width: 84, fontFamily: 'monospace', marginLeft: 4 };

// v18.22: Fett/Kursiv-Umschalter (für Überschrift + Unter-Überschrift).
// v31.2: als Chip mit Hover und `is-active` statt eigener Inline-Farben.
const FmtToggle: React.FC<{ active: boolean; onClick: () => void; title: string; children: React.ReactNode }> = ({ active, onClick, title, children }) => (
  <button
    type="button"
    className={cx('dex-ui-chip', active && 'is-active')}
    title={title}
    aria-pressed={active}
    onMouseDown={e => e.preventDefault()}
    onClick={onClick}
    style={{ minWidth: 30, justifyContent: 'center', padding: '4px 9px' }}
  >{children}</button>
);

export const HtmlEditorModal: React.FC<HtmlEditorModalProps> = (props) => {
  const {
    open, onClose, title,
    value, onChange,
    previewMode,
    emailSubject, onEmailSubjectChange,
    emailHeading, onEmailHeadingChange, emailHeadingColor = '#86bc25',
    emailHeadingFontSize, onEmailHeadingColorChange, onEmailHeadingFontSizeChange,
    emailHeadingBold, emailHeadingItalic, onEmailHeadingBoldChange, onEmailHeadingItalicChange,
    emailSubheading, onEmailSubheadingChange,
    emailSubheadingColor, emailSubheadingFontSize, emailSubheadingBold, emailSubheadingItalic,
    onEmailSubheadingColorChange, onEmailSubheadingFontSizeChange, onEmailSubheadingBoldChange, onEmailSubheadingItalicChange,
    imageWidth, imagePaddingV, imagePaddingH,
    onImageWidthChange, onImagePaddingVChange, onImagePaddingHChange,
    outlookHeading, onOutlookHeadingChange,
    outlookSubheading, onOutlookSubheadingChange,
    outlookSubject, onOutlookSubjectChange,
    outlookDateEditor, outlookLocationValue, onOutlookLocationChange, outlookLocationAuto,
    defaultBodyHtml,
    previewVars = {}, insertableVars = [],
    previewHtmlVars,
    leftPanel,
    logoBase64 = '', imageBase64 = '',
    extraAction,
    headerExtra,
    bodyTemplates, bodyTemplatesLabel,
    previewToLine, previewSubjectLine,
  } = props;

  const editorRef = React.useRef<HTMLDivElement>(null);
  const isMobile = useIsMobile();
  // v31.2: Sprache (s. Import) — `t` hält die Doppeltexte auf einer Zeile.
  const isDe = useLocaleSafe() === 'de';
  const t = (de: string, en: string): string => (isDe ? de : en);
  // v20.4: App-Modals statt window.confirm/window.prompt.
  // v30.51: `promptDialog` wird hier nicht mehr gebraucht — der Link läuft
  // über den eigenen `LinkDialog` (Ziel + Anzeige-Text in einem Fenster).
  const { confirmDialog } = useDialog();
  const savedSelectionRef = React.useRef<Range | null>(null);
  // v18.20: aktuelle Schriftgröße der Auswahl (px) — treibt die „wie in Word"-
  // Anzeige im Größen-Dropdown. null = keine Auswahl im Editor / unbekannt.
  const [currentFontPx, setCurrentFontPx] = React.useState<number | null>(null);
  // v18.22: freier Hex-Code für die Body-Textfarbe (Picker/Eingabe).
  const [bodyHexDraft, setBodyHexDraft] = React.useState('#000000');
  // v31.2: Auszeichnung der aktuellen Auswahl (fett/kursiv/unterstrichen) —
  // reine Anzeige für `is-active` auf den Toolbar-Knöpfen, wie in Word. Wird
  // bei jeder Auswahl-Änderung im Editor und nach jedem Toolbar-Klick gelesen.
  const [activeFmt, setActiveFmt] = React.useState<{ bold: boolean; italic: boolean; underline: boolean }>({ bold: false, italic: false, underline: false });
  // v31.7: Bild einfügen. Der versteckte Datei-Wähler wird vom Toolbar-Knopf
  // ausgelöst; `imgBusy` sperrt ihn, solange komprimiert wird (eine Sekunde
  // reicht für einen Doppelklick, und zwei Läufe auf dieselbe Auswahl fügen
  // zwei Bilder ein). `imgNote` ist die Rückmeldung UNTER dem Editor — sie
  // nennt die erreichte Größe, weil man einem Bild sonst nicht ansieht, ob es
  // 8 oder 38 KB in den Mailtext geschrieben hat.
  // ACHTUNG: Alle Hooks stehen vor `if (!open) return null` — dahinter reißt
  // ein Hook die Reihenfolge (React #300, s. CLAUDE.md/RegistrationPage v30.3).
  const imgInputRef = React.useRef<HTMLInputElement>(null);
  const [imgBusy, setImgBusy] = React.useState(false);
  const [imgNote, setImgNote] = React.useState<{ tone: 'ok' | 'warn'; text: string } | null>(null);
  // v31.10: Auf dem Handy stehen Bearbeiten und Vorschau NICHT nebeneinander,
  // sondern hintereinander — gestapelt blieben dem Schreibfeld rund 80 px,
  // sobald die Tastatur offen war (gemessen bei 400 × 420). Der Umschalter
  // gibt einer Ansicht die volle Höhe. Die Editor-Spalte bleibt dabei immer im
  // Baum (nur `display: none`): Ihr Inhalt wird ausschließlich beim Öffnen aus
  // `value` geladen, ein Aus- und Wiedereinhängen käme also leer zurück und
  // der nächste `onInput` würde den Text überschreiben.
  const [mobileView, setMobileView] = React.useState<'actions' | 'edit' | 'preview'>(leftPanel ? 'actions' : 'edit');
  // v30.51: Offener Link-Dialog samt eingefrorenem Ausgangszustand.
  const [linkState, setLinkState] = React.useState<null | {
    href: string;
    text: string;
    editing: boolean;
    textLocked: boolean;
    anchor: HTMLAnchorElement | null;
  }>(null);

  // External value beim Öffnen in den Editor laden (Re-Open mit anderem Template)
  React.useEffect(() => {
    // v31.2: idempotent — und nötig, falls weder Wizard noch `Modal` das
    // Stylesheet schon eingehängt haben (Editor aus einer Vorschau geöffnet).
    if (open) ensureDexUiStyles();
    // v31.10: Jedes Öffnen beginnt auf der Ansicht, die auch die gestapelte
    // Reihenfolge zuerst zeigte — sonst landet man nach einem Blick in die
    // Vorschau beim nächsten Mal wieder dort statt im Text.
    if (open) setMobileView(leftPanel ? 'actions' : 'edit');
    if (open && editorRef.current) {
      const cur = editorRef.current.innerHTML;
      if (cur !== (value || '')) {
        editorRef.current.innerHTML = value || '';
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  // v29.42: Auswahl über `selectionchange` am Dokument mitschreiben statt nur
  // über mouseup/keyup IM Editor.
  //
  // Der gemeldete Fall („markieren, Fett klicken — mal geht es, mal nicht"):
  // Wer beim Markieren mit der Maus über den unteren/rechten Rand des Editors
  // hinauszieht — bei einer Markierung bis zum Zeilenende der Normalfall —,
  // lässt die Taste AUSSERHALB los. Das `mouseup` trifft dann document, nicht
  // den Editor, `savedSelectionRef` behält die vorherige (oft leere) Auswahl,
  // und der Toolbar-Klick stellt genau die wieder her: Fett landet ins Leere.
  // `selectionchange` feuert unabhängig davon, wo die Maus losgelassen wird.
  React.useEffect(() => {
    if (!open) return undefined;
    const onSelChange = (): void => {
      const el = editorRef.current;
      const sel = window.getSelection();
      if (!el || !sel || sel.rangeCount === 0) return;
      if (el.contains(sel.anchorNode) && el.contains(sel.focusNode)) {
        savedSelectionRef.current = sel.getRangeAt(0).cloneRange();
      }
    };
    document.addEventListener('selectionchange', onSelChange);
    return () => document.removeEventListener('selectionchange', onSelChange);
  }, [open]);

  if (!open) return null;

  const saveSelection = (): void => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0 && editorRef.current?.contains(sel.anchorNode)) {
      savedSelectionRef.current = sel.getRangeAt(0).cloneRange();
    }
  };

  // v18.20: effektive Schriftgröße der aktuellen Auswahl ermitteln (px), damit
  // das Größen-Dropdown — wie in Word — zeigt, wie groß der markierte Text
  // gerade ist. Wir lesen die berechnete font-size des Elements am Cursor.
  const detectFontSize = (): void => {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || !editorRef.current) { setCurrentFontPx(null); return; }
    const node = sel.anchorNode;
    if (!node || !editorRef.current.contains(node)) { setCurrentFontPx(null); return; }
    const el: HTMLElement | null = node.nodeType === Node.ELEMENT_NODE
      ? (node as HTMLElement)
      : node.parentElement;
    if (!el) { setCurrentFontPx(null); return; }
    const px = Math.round(parseFloat(window.getComputedStyle(el).fontSize));
    setCurrentFontPx(isNaN(px) ? null : px);
  };

  // v31.2: Welche Auszeichnung am Cursor gilt — nur für die Optik der Toolbar.
  // `queryCommandState` kann in alten Engines werfen, deshalb je Befehl gekapselt.
  const detectActiveFmt = (): void => {
    const sel = window.getSelection();
    const el = editorRef.current;
    if (!sel || sel.rangeCount === 0 || !el || !el.contains(sel.anchorNode)) {
      setActiveFmt({ bold: false, italic: false, underline: false });
      return;
    }
    const q = (cmd: string): boolean => { try { return document.queryCommandState(cmd); } catch { return false; } };
    setActiveFmt({ bold: q('bold'), italic: q('italic'), underline: q('underline') });
  };

  // Auswahl sichern UND Größen-Anzeige aktualisieren (mouseup/keyup im Editor).
  const syncSelection = (): void => { saveSelection(); detectFontSize(); detectActiveFmt(); };

  const restoreSelection = (): void => {
    const el = editorRef.current;
    if (!el) return;
    const sel = window.getSelection();
    // v29.42: Eine LEBENDE Auswahl im Editor schlägt die gespeicherte. Vorher
    // wurde immer die gespeicherte gesetzt — war die veraltet (siehe
    // selectionchange oben), überschrieb sie die richtige.
    const live = (sel && sel.rangeCount > 0 && el.contains(sel.anchorNode) && el.contains(sel.focusNode))
      ? sel.getRangeAt(0).cloneRange()
      : null;
    const range = live || savedSelectionRef.current;
    // focus() nur, wenn nötig — auf ein bereits fokussiertes Element ist es ein
    // No-op, auf ein unfokussiertes kann es die Auswahl kollabieren lassen.
    if (document.activeElement !== el) el.focus();
    if (range && sel) {
      // Nach dem Fokuswechsel steht die Auswahl evtl. am Anfang — deshalb in
      // jedem Fall neu setzen.
      sel.removeAllRanges();
      sel.addRange(range);
    }
  };

  const fireChange = (): void => {
    if (editorRef.current) onChange(editorRef.current.innerHTML);
  };

  const exec = (cmd: string, arg?: string): void => {
    restoreSelection();
    try { document.execCommand(cmd, false, arg); } catch { /* ignore */ }
    fireChange();
    // v31.2: Nach „Fett" soll der Knopf sofort gedrückt aussehen.
    detectActiveFmt();
  };

  // v18.39: nächstgelegenes <a> der aktuellen Auswahl finden (für „Link
  // bearbeiten" — URL vorbelegen / bestehenden Link erkennen).
  const getSelectionAnchor = (): HTMLAnchorElement | null => {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    let node: Node | null = sel.getRangeAt(0).startContainer;
    while (node && node !== editorRef.current) {
      if (node.nodeType === 1 && (node as HTMLElement).tagName === 'A') return node as HTMLAnchorElement;
      node = node.parentNode;
    }
    return null;
  };

  const escHtml = (s: string): string =>
    s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  // v18.39: Link einfügen / bearbeiten / entfernen — ohne Copy-Paste, damit
  // beim Setzen eines Links keine fremde Formatierung (Zeilenabstand) mehr
  // mit hineinkommt.
  /**
   * v30.51: EIN Dialog für Ziel UND Anzeige-Text (s. components/LinkDialog).
   *
   * Der Zustand des Dialogs wird beim Öffnen eingefroren: `linkState` hält
   * die bestehende href, den Anzeige-Text und — entscheidend — die Auswahl,
   * die beim Klick auf den Knopf galt. Der Dialog stiehlt den Fokus, danach
   * ist die Editor-Auswahl nicht mehr verlässlich; deshalb wird sie beim
   * Übernehmen aus `savedSelectionRef` wiederhergestellt und nicht neu
   * gelesen (dieselbe Falle wie bei den alten Prompts, v20.4).
   */
  const openLinkDialog = (): void => {
    restoreSelection();
    saveSelection();
    const existing = getSelectionAnchor();
    const sel = window.getSelection();
    const hasRange = !!sel && sel.rangeCount > 0;
    const selectedText = hasRange ? sel!.getRangeAt(0).toString() : '';
    // Über mehrere Absätze hinweg kann der Anzeige-Text nicht ersetzt werden,
    // ohne die Formatierung dazwischen wegzuwerfen — dann wird nur verlinkt.
    const multiBlock = selectedText.indexOf('\n') >= 0;
    setLinkState({
      href: existing ? (existing.getAttribute('href') || '') : '',
      text: existing ? (existing.textContent || '') : selectedText,
      editing: !!existing,
      textLocked: multiBlock,
      anchor: existing,
    });
  };

  const applyLink = (href: string, text: string): void => {
    const st = linkState;
    setLinkState(null);
    if (!st) return;
    restoreSelection();
    const sel = window.getSelection();
    const collapsed = !sel || sel.rangeCount === 0 || sel.getRangeAt(0).collapsed;
    const label = text || href.replace(/^mailto:/i, '');
    if (st.anchor) {
      // Bestehender Link: Ziel immer, Anzeige-Text nur wenn er sich geändert
      // hat — sonst würde eine Auszeichnung IM Link (fett, farbig) durch
      // reinen Text ersetzt.
      st.anchor.setAttribute('href', href);
      if (!st.textLocked && text && text !== (st.anchor.textContent || '')) {
        st.anchor.textContent = text;
      }
    } else if (collapsed) {
      try { document.execCommand('insertHTML', false, `<a href="${escHtml(href)}">${escHtml(label)}</a>`); } catch { /* ignore */ }
    } else if (!st.textLocked && text && text !== sel!.getRangeAt(0).toString()) {
      // Markierung vorhanden UND der Anzeige-Text wurde geändert: die Auswahl
      // komplett durch den neuen Link ersetzen.
      try { document.execCommand('insertHTML', false, `<a href="${escHtml(href)}">${escHtml(text)}</a>`); } catch { /* ignore */ }
    } else {
      // Markierten Text unverändert verlinken — erhält die Formatierung darin.
      try { document.execCommand('createLink', false, href); } catch { /* ignore */ }
    }
    fireChange();
  };

  const removeLink = (): void => {
    const st = linkState;
    setLinkState(null);
    restoreSelection();
    if (st?.anchor) {
      // `unlink` braucht eine Auswahl IM Link; nach dem Dialog ist die nicht
      // mehr sicher — deshalb den Anker direkt durch seinen Inhalt ersetzen.
      const parent = st.anchor.parentNode;
      if (parent) {
        while (st.anchor.firstChild) parent.insertBefore(st.anchor.firstChild, st.anchor);
        parent.removeChild(st.anchor);
      }
    } else {
      try { document.execCommand('unlink'); } catch { /* ignore */ }
    }
    fireChange();
  };

  /**
   * v31.7: Ein Bild an die Cursor-Position — aus dem Knopf ODER aus der
   * Zwischenablage, in beiden Fällen derselbe Weg.
   *
   * Anhänge sind gesperrt (der Deloitte-Mailflow beantwortet Power-Automate-
   * Mails mit Anhang mit einem NDR, v26.71), also ist Base64 im HTML nicht die
   * Notlösung, sondern der Weg, den QR-Code, Deloitte-Logo und Kopfbild schon
   * gehen. Neu ist nur, dass ihn der Organizer selbst nutzen kann.
   *
   * Zwei Grenzen, beide aus `utils/inlineMailImage`: die des einzelnen Bildes
   * (was auch nach der letzten Kompressions-Stufe zu groß ist, wird ABGELEHNT
   * statt stillschweigend verschluckt) und die des ganzen Textes — der landet
   * in einer SharePoint-Note-Spalte, und ist die voll, nimmt die Warteschlange
   * die Mail gar nicht erst an. Deshalb wird vor dem Einfügen gerechnet, nicht
   * danach entschuldigt.
   */
  const insertInlineImage = async (file: File | null): Promise<void> => {
    if (!file || imgBusy) return;
    setImgNote(null);
    setImgBusy(true);
    const res: InlineImageOutcome = await buildInlineImage(file)
      .catch((): InlineImageOutcome => ({ ok: false, reason: 'unreadable', dataUrl: '', width: 0, height: 0, chars: 0 }));
    setImgBusy(false);
    if (!res.ok) {
      setImgNote({
        tone: 'warn',
        text: res.reason === 'too-big'
          ? t(
            `Das Bild ist auch nach dem Verkleinern noch ${charsToKb(res.chars)} KB groß — mehr als ${INLINE_IMG_MAX_KB} KB passen nicht in einen Mailtext. Schneide es zu oder speichere es vorher kleiner ab, dann klappt es.`,
            `Even after shrinking, the image is still ${charsToKb(res.chars)} KB — more than ${INLINE_IMG_MAX_KB} KB does not fit into an email. Crop it or save it smaller first, then it works.`)
          : t(
            'Das Bild konnte nicht gelesen werden. Nimm eine JPG- oder PNG-Datei — SVG und Dateien ohne feste Größe gehen nicht.',
            'The image could not be read. Use a JPG or PNG file — SVG and files without a fixed size do not work.'),
      });
      return;
    }
    const html = inlineImageHtml(res.dataUrl, res.width);
    const before = editorRef.current?.innerHTML.length || 0;
    if (before + html.length > MAIL_BODY_MAX_KB * 1024) {
      setImgNote({
        tone: 'warn',
        text: t(
          `Der Text ist schon ${charsToKb(before)} KB groß, mit diesem Bild wären es ${charsToKb(before + html.length)} KB — höchstens ${MAIL_BODY_MAX_KB} KB nimmt die Mail-Warteschlange an. Lösch ein Bild heraus oder verlink es stattdessen.`,
          `The text is already ${charsToKb(before)} KB, with this image it would be ${charsToKb(before + html.length)} KB — the mail queue accepts at most ${MAIL_BODY_MAX_KB} KB. Remove an image or link to it instead.`),
      });
      return;
    }
    restoreSelection();
    try { document.execCommand('insertHTML', false, html); } catch { /* ignore */ }
    fireChange();
    const after = editorRef.current?.innerHTML.length || (before + html.length);
    setImgNote({
      tone: after > MAIL_BODY_WARN_KB * 1024 ? 'warn' : 'ok',
      text: t(
        `Bild eingefügt: ${res.width} × ${res.height} px, ${charsToKb(res.chars)} KB. Der Text ist jetzt ${charsToKb(after)} von höchstens ${MAIL_BODY_MAX_KB} KB.`,
        `Image inserted: ${res.width} × ${res.height} px, ${charsToKb(res.chars)} KB. The text is now ${charsToKb(after)} of at most ${MAIL_BODY_MAX_KB} KB.`),
    });
  };

  // v18.39: Einfügen IMMER als reiner Text (mit Zeilenumbrüchen als <br>).
  // Verhindert, dass kopierte Inhalte Block-Markup (<div>/<p> mit Außen-
  // abständen) mitbringen, das den Zeilenabstand „plötzlich größer" macht
  // und sich danach nicht mehr korrigieren lässt.
  //
  // v31.9.8: Der Preis dafür war zu hoch. Wer eine fertige Outlook-Mail
  // einfügt, verlor JEDEN Link, alles Fette und alle Aufzählungen; die
  // Listenpunkte kamen als „•"-Zeichen im Fließtext an (gemeldet 10.09.2026
  // an einer Rundmail an 150 Empfänger). Statt der Rücknahme filtert jetzt
  // `sanitizePastedHtml` über eine Positivliste: Link, fett, kursiv,
  // unterstrichen, Aufzählung, Umbruch bleiben — jedes `style`, jede Klasse
  // und jede Breitenangabe fliegt raus. Der Zeilenabstand ist damit weiterhin
  // unantastbar, die Begründung von v18.39 also erfüllt. Bleibt nichts
  // Sichtbares übrig, greift unverändert der Klartext-Weg.
  const handlePaste = (e: React.ClipboardEvent<HTMLDivElement>): void => {
    // v31.7: Enthält die Zwischenablage ein BILD und keinen Text, wird es wie
    // über den Knopf verarbeitet — ein Screenshot mit Strg+V ist die natürliche
    // Geste, und bisher verschwand er stillschweigend. Die Datei muss SYNCHRON
    // geholt werden: nach dem ersten `await` ist `clipboardData` leer. Kommt
    // Text mit (Word legt daneben ein Bild des Bereichs ab), bleibt unten alles
    // beim v18.39-Verhalten — die Begründung dafür gilt unverändert.
    const pastedImage = imageFileFromClipboard(e.clipboardData);
    if (pastedImage) {
      e.preventDefault();
      saveSelection();
      void insertInlineImage(pastedImage);
      return;
    }
    e.preventDefault();
    const text = e.clipboardData?.getData('text/plain') || '';
    const rich = sanitizePastedHtml(e.clipboardData?.getData('text/html') || '');
    const html = rich || (text ? text.split(/\r\n|\r|\n/).map(line => escHtml(line)).join('<br>') : '');
    if (!html) return;
    try { document.execCommand('insertHTML', false, html); } catch {
      try { document.execCommand('insertText', false, text); } catch { /* ignore */ }
    }
    fireChange();
  };

  const setFontSize = (px: number): void => {
    restoreSelection();
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (sel.getRangeAt(0).collapsed) return;
    const root = editorRef.current;
    if (!root) return;
    // v18.21: robuste Größen-Zuweisung über execCommand (wie foreColor) statt
    // manuellem range.extractContents() — letzteres warf bei Auswahlen über
    // Element-/Zeilen-Grenzen (z.B. fett markiertes „Location:") eine
    // Exception, die verschluckt wurde → es entstand KEIN Span, die Vorschau
    // blieb unverändert. Trick: styleWithCSS aus → execCommand('fontSize','7')
    // wrappt die Auswahl robust in <font size="7"> (eindeutiger Marker), die
    // wir danach in <span style="font-size:Npx"> umschreiben.
    try { document.execCommand('styleWithCSS', false, 'false'); } catch { /* ignore */ }
    try { document.execCommand('fontSize', false, '7'); } catch { /* ignore */ }
    const markers = root.querySelectorAll('font[size="7"]');
    markers.forEach(node => {
      const span = document.createElement('span');
      span.style.fontSize = `${px}px`;
      while (node.firstChild) span.appendChild(node.firstChild);
      node.parentNode?.replaceChild(span, node);
    });
    // Fallback: hat ein Browser trotz styleWithCSS=false einen span mit
    // font-size:xxx-large erzeugt (Sentinel von Größe 7), auch den umschreiben.
    root.querySelectorAll('span[style*="xxx-large"]').forEach(node => {
      (node as HTMLElement).style.fontSize = `${px}px`;
    });
    fireChange();
    // v18.20: Anzeige sofort auf die gesetzte Größe stellen — der User sieht
    // im Dropdown direkt die neue Größe, und die Live-Vorschau aktualisiert
    // sich über fireChange().
    setCurrentFontPx(px);
  };

  const setColor = (hex: string): void => exec('foreColor', hex);

  // v18.16: Zeilenabstand für die GESAMTE Beschreibung einstellbar. Wir
  // wrappen den kompletten Editor-Inhalt in ein div[data-lh] mit inline
  // line-height — das wird in der gespeicherten innerHTML mitgesichert und
  // überschreibt auf der Anmelde-Seite die Default-Zeilenhöhe.
  const setLineHeight = (lh: string): void => {
    const root = editorRef.current;
    if (!root) return;
    let wrapper = root.querySelector(':scope > div[data-lh]') as HTMLElement | null;
    if (!wrapper || root.childNodes.length !== 1) {
      wrapper = document.createElement('div');
      wrapper.setAttribute('data-lh', '1');
      while (root.firstChild) wrapper.appendChild(root.firstChild);
      root.appendChild(wrapper);
    }
    wrapper.style.lineHeight = lh;
    fireChange();
  };

  const insertVariable = (key: string): void => {
    restoreSelection();
    // execCommand insertText fügt an Cursor ein (oder ersetzt Selektion).
    // Variable als Text — der Server ersetzt sie später per replacePlaceholders.
    try { document.execCommand('insertText', false, key); } catch { /* ignore */ }
    fireChange();
  };

  // v30.94: Orb-Schutz auch in DIESER Vorschau. Ohne eigenes Mail-Logo zeigt
  // der Kopf den DEX-Orb; der wurde mit dem Vollbild-Layout (600/0/0, seit
  // v30.87 Standard) auf volle Breite gezogen — die Vorschau-Karte im Schritt
  // rechnet über headerLayoutFor schon richtig, das Modal nahm die Rohwerte.
  // Dieselbe Regel wie dort: ohne eigenes Bild höchstens 180 px, mindestens
  // 20 px Abstand. Die Eingabefelder zeigen weiter die gespeicherten Werte —
  // sie gelten, sobald ein Logo hochgeladen ist.
  const ownHeaderImage = !!(imageBase64 && imageBase64.trim());
  const effImageWidth = ownHeaderImage || imageWidth === undefined ? imageWidth : Math.min(imageWidth, 180);
  const effImagePaddingV = ownHeaderImage || imagePaddingV === undefined ? imagePaddingV : Math.max(imagePaddingV, 20);
  const effImagePaddingH = ownHeaderImage || imagePaddingH === undefined ? imagePaddingH : Math.max(imagePaddingH, 20);

  const renderPreviewHtml = (): string => {
    let bodyWithVars = replacePlaceholders(value || '', previewVars);
    // v22.18: HTML-Platzhalter (z.B. {{QR_BLOCK}}) RAW ersetzen — nach den
    // escapten previewVars, damit das eingesetzte HTML nicht escaped wird.
    if (previewHtmlVars) {
      for (const k of Object.keys(previewHtmlVars)) {
        bodyWithVars = bodyWithVars.replace(new RegExp(`\\{\\{${k}\\}\\}`, 'g'), previewHtmlVars[k]);
      }
    }
    const cachedLogo = getCachedLogoBase64();
    const cachedOrb = getCachedOrbBase64();

    if (previewMode === 'plain') {
      // v9.40: kein Wrapper — die Beschreibung wird auf der Anmelde-Seite genauso
      // gerendert, eingebettet in die normale Seitentypografie. Wir simulieren das
      // mit minimalen Default-Styles (System-Font + 14px), damit der Editor-Text
      // nicht in irgendeinem rohen Browser-Default aussieht.
      const empty = !bodyWithVars || bodyWithVars.trim() === '';
      const html = empty
        ? `<p style="color:#999;font-style:italic;">${t('Hier erscheint die Beschreibung, sobald du im Editor links etwas tippst — sie wird 1:1 auf der Anmelde-Seite angezeigt.', 'The description appears here as soon as you type in the editor on the left — it is shown 1:1 on the registration page.')}</p>`
        : bodyWithVars;
      return `<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica,Arial,sans-serif;font-size:14px;line-height:1.55;color:#333;margin:0;padding:24px;background:#fff;}img{max-width:100%;height:auto;}a{color:#0076a8;}h1,h2,h3,h4{color:#222;}</style></head><body>${html}</body></html>`;
    }

    if (previewMode === 'email') {
      const heading = replacePlaceholdersPlain(emailHeading || '', previewVars);
      // v15.19: Subheading aus dem Override; Fallback {{EventTitle}}.
      const rawSub = (emailSubheading && emailSubheading.trim()) || '{{EventTitle}}';
      const subheading = replacePlaceholdersPlain(rawSub, previewVars);
      // Manche System-Templates (z.B. Nachrücken, OutlookDeclineReminder)
      // werden bereits komplett Deloitte-gewrappt gespeichert
      // (wrapTemplateForStorage), weil die Power-Automate-Flows den BodyHtml
      // roh versenden. In dem Fall NICHT noch einmal wrappen — sonst tauchen
      // Logo + Event-Titel doppelt auf. Gleiche Logik wie im Outlook-Zweig
      // und in buildEmailFromTemplate (isPreWrapped).
      const isAlreadyWrapped = /^\s*(<!doctype|<html)/i.test(bodyWithVars);
      const wrapped = isAlreadyWrapped
        ? bodyWithVars
        : wrapTemplate(emailHeadingColor, heading, subheading, bodyWithVars, emailHeadingFontSize, {
          headingBold: emailHeadingBold,
          headingItalic: emailHeadingItalic,
          subheadingColor: emailSubheadingColor,
          subheadingFontSize: emailSubheadingFontSize,
          subheadingBold: emailSubheadingBold,
          subheadingItalic: emailSubheadingItalic,
          // v18.73: Header-Bild Größe + Innenabstand live mitvorschauen.
          imageWidth: effImageWidth, imagePaddingV: effImagePaddingV, imagePaddingH: effImagePaddingH,
        });
      return wrapped
        .replace(/\{\{LOGO_URL\}\}/g, logoBase64 || cachedLogo || '')
        .replace(/\{\{ORB_URL\}\}/g, imageBase64 || cachedOrb || '');
    }

    // Outlook-Termin-Vorschau: gleicher wrapTemplate() wie beim echten Versand,
    // damit der User in der Vorschau genau das sieht, was nachher im Termin steht
    // (inkl. Deloitte-Signatur + Legal-Disclaimer im Footer).
    const olHeading = replacePlaceholdersPlain(outlookHeading || '', previewVars) || previewVars.EventTitle || 'Event Title';
    // v27.5: Default-Unter-Überschrift = Ort (nicht mehr Datum/Uhrzeit).
    const olSub = replacePlaceholdersPlain(outlookSubheading || '', previewVars) || previewVars.Location || previewVars.EventDate || 'Event Details';
    // v30.94: Leerer Body zeigt den Standard-Text, der beim Speichern tatsächlich
    // in den Termin kommt (defaultBodyHtml = utils/outlookDefaultBody) — mit dem
    // Hinweis, dass er nur gilt, solange hier nichts steht. Vorher stand hier
    // ein Platzhalter-Satz, die Vorschau-Karte zeigte etwas anderes.
    const bodyForOutlook = bodyWithVars
      || (defaultBodyHtml
        ? replacePlaceholders(defaultBodyHtml, previewVars) + `<p style="color:#999;font-style:italic;font-size:12px;margin-top:18px;">${t('Standardtext — gilt, solange du links nichts eingibst. Klick „Standardtext laden", um ihn zu übernehmen und anzupassen.', 'Default text — applies as long as you leave the editor empty. Click “Load default text” to take it over and adjust it.')}</p>`
        : `<p style="color:#999;font-style:italic;">${t('Hier erscheint der Text — beginne im Editor links zu tippen.', 'Your text appears here — start typing in the editor on the left.')}</p>`);
    // Wenn der Body bereits ein kompletter wrapTemplate-Output ist (z.B. aus editEvent
    // ohne Strip), 1:1 anzeigen — sonst doppelt wickeln.
    const isAlreadyWrapped = /<!doctype|<html/i.test(bodyForOutlook);
    const wrapped = isAlreadyWrapped
      ? bodyForOutlook
      // v18.73: Header-Bild Größe + Innenabstand live mitvorschauen.
      : wrapTemplate('#86bc25', olHeading, olSub, bodyForOutlook, undefined, { imageWidth: effImageWidth, imagePaddingV: effImagePaddingV, imagePaddingH: effImagePaddingH });
    return wrapped
      .replace(/\{\{LOGO_URL\}\}/g, logoBase64 || cachedLogo || '')
      .replace(/\{\{ORB_URL\}\}/g, imageBase64 || cachedOrb || '');
  };

  // v31.2: Kleine Render-Helfer für die Dreiteilung Aktionen | Bearbeiten |
  // Vorschau. Kein State, nur Optik — die Handler oben bleiben unverändert.
  // v31.10: `flexWrap`, damit die Zeile auf 360 px umbricht statt rechts
  // abgeschnitten zu werden; ein leeres `label` lässt die Überschrift weg —
  // auf dem Handy sagt der Umschalter darüber bereits, welche Ansicht offen ist.
  const colHead = (label: string, right?: React.ReactNode): React.ReactElement => (
    <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', padding: isMobile ? '8px 14px' : '9px 18px', minHeight: 40, boxSizing: 'border-box', borderBottom: '1px solid var(--dex-gray-200, #e8e8e8)', background: '#fff', flexShrink: 0 }}>
      {!!label && <span style={{ fontSize: '0.72rem', fontWeight: 700, letterSpacing: '0.08em', textTransform: 'uppercase', color: 'var(--dex-gray-500, #808080)', whiteSpace: 'nowrap' }}>{label}</span>}
      {right}
    </div>
  );
  const field = (label: string, control: React.ReactNode, help?: React.ReactNode, optional?: boolean): React.ReactElement => (
    <div className="dex-ui-field">
      <div className="dex-ui-label">{label}{optional && <span className="dex-ui-label-optional">{t('(optional)', '(optional)')}</span>}</div>
      {control}
      {help && <div className="dex-ui-help">{help}</div>}
    </div>
  );
  // Aufklapper ohne eigenen State: <details> merkt sich offen/zu selbst, und
  // der Dialog wird beim Schließen ohnehin abgebaut (`if (!open) return null`).
  const disclosure = (label: string, count: React.ReactNode, body: React.ReactNode): React.ReactElement => (
    <details>
      <summary className="dex-ui-disclosure" style={{ listStyle: 'none' }}>
        <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
        {label}
        {count !== undefined && <span className="dex-ui-disclosure-count">{count}</span>}
      </summary>
      <div className="dex-ui-disclosure-body">{body}</div>
    </details>
  );
  const smallLabel: React.CSSProperties = { display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.76rem', color: 'var(--dex-gray-600, #666)', fontWeight: 600 };
  const numLabel: React.CSSProperties = { display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.76rem', color: 'var(--dex-gray-600, #666)', fontWeight: 600 };
  // v31.10: `flexShrink: 0`, damit die Auswahlfelder in der rollenden
  // Handy-Toolbar nicht auf Pfeilbreite zusammenfallen; 40 px Höhe als Tippziel.
  const smallSelect: React.CSSProperties = { width: 'auto', paddingRight: 28, flexShrink: 0, ...(isMobile ? { minHeight: 40 } : {}) };
  const envKey: React.CSSProperties = { color: 'var(--dex-gray-500, #808080)', fontWeight: 600, fontSize: '0.72rem', textTransform: 'uppercase', letterSpacing: '0.05em', alignSelf: 'center' };
  // Größe / fett / kursiv / Farbe einer Überschrift in EINER Zeile — einmal
  // für die Überschrift, einmal für die Unter-Überschrift. Die Bold-Semantik
  // bleibt je Aufrufer (Unter-Überschrift: fett, solange nicht ausdrücklich aus).
  const fmtRow = (o: {
    sizeValue: string; sizeDefault: string; sizeOptions: string[]; onSize?: (px: string) => void;
    boldActive: boolean; onBold?: () => void; italicActive: boolean; onItalic?: () => void;
    color: string; onColor?: (hex: string) => void;
  }): React.ReactElement => (
    <div style={{ display: 'flex', alignItems: 'center', gap: 14, flexWrap: 'wrap' }}>
      {o.onSize && (
        <label style={smallLabel}>
          {t('Größe', 'Size')}
          <select className="dex-ui-select dex-ui-input--sm" value={o.sizeValue} onChange={e => { if (o.onSize) o.onSize(e.target.value); }} style={smallSelect}>
            {o.sizeOptions.map(px => <option key={px} value={px}>{px}{px === o.sizeDefault ? t(' (Standard)', ' (default)') : ''}</option>)}
          </select>
        </label>
      )}
      {(o.onBold || o.onItalic) && (
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
          {o.onBold && <FmtToggle active={o.boldActive} onClick={o.onBold} title={t('Fett', 'Bold')}><strong>F</strong></FmtToggle>}
          {o.onItalic && <FmtToggle active={o.italicActive} onClick={o.onItalic} title={t('Kursiv', 'Italic')}><em>{t('K', 'I')}</em></FmtToggle>}
        </span>
      )}
      {o.onColor && <span style={smallLabel}>{t('Farbe', 'Color')}<ColorControl value={o.color} onChange={o.onColor} /></span>}
    </div>
  );
  // Toolbar-Knopf: runder Iconknopf mit Hover; `onMouseDown` verhindert, dass
  // der Klick die Editor-Auswahl kollabiert (dann liefe „Fett" ins Leere).
  // v31.2: `active` setzt `is-active` (gedrückt), wenn die Auswahl die
  // Auszeichnung schon trägt; nur Schriftmaße bleiben inline, weil die Knöpfe
  // Buchstaben statt Symbole zeigen.
  // v31.10: Auf dem Handy rollt die Leiste waagerecht (s. unten), und dann
  // entscheidet die Reihenfolge darüber, was man ohne Rollen erreicht: erst
  // die täglichen Knöpfe (Fett/Kursiv/Unterstrichen, Listen, Link, Bild,
  // Formatierung entfernen), dann die breiten Auswahlfelder, zuletzt die
  // Farbreihe. Umgestellt wird nur per CSS-`order` — die Knöpfe bleiben im
  // Code an ihrem Platz, und am Desktop ändert sich nichts.
  const tbOrder = (n: number): React.CSSProperties => (isMobile ? { order: n } : {});
  const tb = (label: string, onClick: () => void, children: React.ReactNode, active?: boolean, order?: number): React.ReactElement => (
    <button
      type="button"
      className={cx('dex-ui-iconbtn', active && 'is-active')}
      title={label}
      aria-label={label}
      aria-pressed={active === undefined ? undefined : active}
      onMouseDown={e => e.preventDefault()}
      onClick={onClick}
      // v31.10: 40 px auf dem Handy — `dex-ui-iconbtn` ist 32 px, und mit dem
      // Finger ist das unter jeder Tippziel-Grenze.
      style={{ fontWeight: 700, fontSize: '0.85rem', ...(isMobile ? { width: 40, height: 40 } : {}), ...tbOrder(order || 0) }}
    >{children}</button>
  );
  // v31.2: Trenner zwischen den Toolbar-Gruppen — `dex-ui-divider` hochkant
  // (die Klasse ist für die waagerechte Linie gebaut, daher Maße inline).
  const tbDivider = (order: number): React.ReactElement => (
    <span className="dex-ui-divider" aria-hidden="true" style={{ width: 1, height: 20, margin: '0 4px', flexShrink: 0, ...tbOrder(order) }} />
  );

  const modeIcon = previewMode === 'outlook' ? <Calendar size={20} /> : previewMode === 'plain' ? <FileText size={20} /> : <Mail size={20} />;
  const modeLabel = previewMode === 'outlook' ? t('Outlook-Termin', 'Outlook appointment') : previewMode === 'plain' ? t('Anmeldeseite', 'Registration page') : t('Deloitte-Mail', 'Deloitte email');
  // v31.10: Auf dem Handy liegt nichts „links" und nichts „rechts" — dort
  // schaltet der Reiter um. Ein Satz, der eine Anordnung beschreibt, die es
  // gerade nicht gibt, schickt genau die Leute suchen, die ohnehin wenig Platz
  // haben.
  const subtitle = isMobile
    ? (previewMode === 'outlook'
      ? t('Schreib deinen Text — unter „Vorschau" siehst du, wie der Termin im Kalender aussieht.', 'Write your text — under “Preview” you see how the appointment looks in the calendar.')
      : previewMode === 'plain'
        ? t('Schreib deinen Text — unter „Vorschau" siehst du, wie er auf der Anmeldeseite aussieht.', 'Write your text — under “Preview” you see how it looks on the registration page.')
        : t('Schreib deinen Text — unter „Vorschau" siehst du, wie die Mail ankommt.', 'Write your text — under “Preview” you see how the email lands.'))
    : previewMode === 'outlook'
      ? t('Links bearbeiten — rechts siehst du sofort, wie der Termin im Kalender aussieht.', 'Edit on the left — the right side shows instantly how the appointment looks in the calendar.')
      : previewMode === 'plain'
        ? t('Links bearbeiten — rechts siehst du sofort, wie der Text auf der Anmeldeseite aussieht.', 'Edit on the left — the right side shows instantly how the text looks on the registration page.')
        : t('Links bearbeiten — rechts siehst du sofort, wie die Mail bei den Teilnehmenden ankommt.', 'Edit on the left — the right side shows instantly how the email lands with attendees.');
  // v31.10: Welche der drei Spalten gerade sichtbar ist. Am Desktop alle,
  // auf dem Handy genau eine (siehe `mobileView`).
  const showActions = !isMobile || mobileView === 'actions';
  const showEdit = !isMobile || mobileView === 'edit';
  const showPreview = !isMobile || mobileView === 'preview';
  const textSectionTitle = previewMode === 'outlook' ? t('Text des Termins', 'Appointment text') : previewMode === 'plain' ? t('Beschreibung', 'Description') : t('Text der Mail', 'Email text');
  const headingFmt = !!(onEmailHeadingFontSizeChange || onEmailHeadingColorChange || onEmailHeadingBoldChange);
  const subFmt = !!(onEmailSubheadingFontSizeChange || onEmailSubheadingColorChange || onEmailSubheadingBoldChange);
  // Nur Anzeige: Welche der beiden Ein-Klick-Vorgaben gerade gilt.
  const imgW = imageWidth ?? 180; const imgH = imagePaddingH ?? 30; const imgV = imagePaddingV ?? 30;
  const fullWidthOn = imgW === 600 && imgH === 0 && imgV === 0;
  const standardOn = imgW === 180 && imgH === 30 && imgV === 30;
  // v31.7: Wie voll ist der Text? Gemessen wird `value` — genau die Zeichenkette,
  // die gespeichert und in die Mail-Warteschlange geschrieben wird; die
  // Editor-Anzeige ist dafür kein Maß.
  const bodyChars = (value || '').length;
  const inlineImgCount = countInlineImages(value || '');
  // Ersetzt den Editor-Inhalt direkt (innerHTML), weil der contentEditable nur
  // beim Öffnen aus `value` synct — mit Rückfrage, wenn schon Text drinsteht.
  const replaceBody = (html: string, question: string): void => {
    const cur = editorRef.current?.innerHTML || '';
    const isEmpty = cur.replace(/<[^>]*>/g, '').replace(/\s/g, '').trim() === '';
    const apply = (): void => {
      if (editorRef.current) { editorRef.current.innerHTML = html; fireChange(); }
    };
    if (isEmpty) { apply(); return; }
    confirmDialog(question, { confirmLabel: t('Ersetzen', 'Replace') })
      .then(ok => { if (ok) apply(); })
      .catch(() => { /* */ });
  };

  return (
    <div
      style={{
        position: 'fixed', inset: 0, zIndex: 1200,
        background: 'rgba(15,23,42,0.5)', backdropFilter: 'blur(2px)', WebkitBackdropFilter: 'blur(2px)',
        display: 'flex', alignItems: 'stretch', justifyContent: 'center',
        padding: isMobile ? '12px 8px' : '28px 24px',
      }}
      // v18.43: KEIN Schließen per Backdrop-Klick mehr — der Editor enthält
      // viel Arbeit (Body, Betreff, Logos) und wurde durch einen Fehlklick
      // neben das Modal zu leicht geschlossen. Schließen nur noch aktiv über
      // das X oben rechts (oder „Fertig").
    >
      <div
        className="dex-ui-modal-card"
        style={{
          width: '100%', maxWidth: isMobile ? '100vw' : (leftPanel ? 1600 : 1280), maxHeight: '100%',
          display: 'flex', flexDirection: 'column',
          background: '#fff', borderRadius: 18, overflow: 'hidden',
          boxShadow: '0 24px 64px rgba(0,0,0,0.22), 0 0 0 1px rgba(0,0,0,0.04)',
        }}
      >
        {/* v31.2: Kopf wie in `Modal` — Symbol für die Art des Textes, Titel,
            ein Satz, was links und rechts passiert, Schließen-Knopf. */}
        <div style={{ padding: isMobile ? '12px 14px 0' : '16px 20px 0', flexShrink: 0 }}>
          <div className="dex-ui-modal-head">
            <span className="dex-ui-modal-head-icon" aria-hidden="true">{modeIcon}</span>
            <div style={{ flex: 1, minWidth: 0 }}>
              <h3 className="dex-ui-modal-title">{title}</h3>
              <p className="dex-ui-modal-subtitle">{subtitle}</p>
            </div>
            <button type="button" className="dex-ui-iconbtn" onClick={onClose} aria-label={t('Schließen', 'Close')} title={t('Schließen', 'Close')} style={{ marginTop: -4, marginRight: -6 }}>
              <X size={18} />
            </button>
          </div>
          {/* v31.10: Der Umschalter des Handys. Die Reiter tragen bewusst KEIN
              `preventDefault` auf `mouseDown`: Der Editor soll den Fokus
              abgeben und dabei über sein `onBlur` den Stand sichern, bevor die
              Vorschau ihn rendert. Die Auswahl selbst lebt in
              `savedSelectionRef` weiter (selectionchange am Dokument). */}
          {isMobile && (
            <div className="dex-ui-tabs" role="tablist" aria-label={t('Ansicht', 'View')} style={{ display: 'flex', width: '100%', marginTop: 10 }}>
              {leftPanel && (
                <button
                  type="button" role="tab" aria-selected={mobileView === 'actions'}
                  className={cx('dex-ui-tab', mobileView === 'actions' && 'is-active')}
                  onClick={() => setMobileView('actions')}
                  style={{ flex: 1, minHeight: 38 }}
                >{t('Aktionen', 'Actions')}</button>
              )}
              <button
                type="button" role="tab" aria-selected={mobileView === 'edit'}
                className={cx('dex-ui-tab', mobileView === 'edit' && 'is-active')}
                onClick={() => setMobileView('edit')}
                style={{ flex: 1, minHeight: 38 }}
              >{t('Bearbeiten', 'Edit')}</button>
              <button
                type="button" role="tab" aria-selected={mobileView === 'preview'}
                className={cx('dex-ui-tab', mobileView === 'preview' && 'is-active')}
                onClick={() => setMobileView('preview')}
                style={{ flex: 1, minHeight: 38 }}
              >{t('Vorschau', 'Preview')}</button>
            </div>
          )}
        </div>

        <div style={{ flex: 1, display: 'flex', flexDirection: isMobile ? 'column' : 'row', minHeight: 0, overflow: 'hidden' }}>
          {/* === OPTIONALE AKTIONS-SPALTE LINKS (v22.19) === */}
          {leftPanel && (
            <div style={{
              width: isMobile ? '100%' : 280, minWidth: isMobile ? 0 : 280,
              flex: isMobile ? 1 : undefined,
              display: showActions ? 'flex' : 'none', flexDirection: 'column', minHeight: 0,
              borderRight: isMobile ? 'none' : '1px solid var(--dex-gray-200, #e8e8e8)',
              background: 'var(--dex-gray-50, #fafafa)',
            }}>
              {!isMobile && colHead(t('Aktionen', 'Actions'))}
              <div style={{ flex: 1, overflow: 'auto', padding: isMobile ? '14px 14px 18px' : '14px 16px' }}>{leftPanel}</div>
            </div>
          )}

          {/* === BEARBEITEN === */}
          <div style={{
            flex: 1, display: showEdit ? 'flex' : 'none', flexDirection: 'column', minHeight: 0,
            width: isMobile ? '100%' : undefined,
            borderRight: isMobile ? 'none' : '1px solid var(--dex-gray-200, #e8e8e8)',
          }}>
            {!isMobile && colHead(t('Bearbeiten', 'Edit'))}
            {/* v31.10: Auf dem Handy stehen die Blöcke in einer Flex-Spalte, damit
                der Schreibblock über `order` nach oben rutschen kann. Am Desktop
                sieht man alles auf einmal, dort bleibt die gewachsene Reihenfolge
                (Kopf → Kopfbild → Text); auf dem Handy hieße sie: erst einen
                ganzen Bildschirm scrollen, dann tippen. Ein reiner CSS-Umzug —
                die Blöcke selbst bleiben, wo sie im Code stehen. */}
            <div style={{ flex: 1, overflow: 'auto', padding: isMobile ? '12px 14px 18px' : '16px 20px 20px', ...(isMobile ? { display: 'flex', flexDirection: 'column' } : {}) }}>
              {/* Der Aufrufer bringt seinen Block (Versand, Empfänger, Kopfbild-
                  Wahl) fertig gerahmt mit — deshalb hier keine zweite Karte.
                  Er bleibt auch auf dem Handy zuerst: Wer die Empfänger noch
                  wählt, entscheidet damit, was der Text überhaupt sein muss. */}
              {headerExtra && <div className="dex-ui-section" style={isMobile ? { order: -2, flexShrink: 0 } : undefined}>{headerExtra}</div>}

              {previewMode === 'email' && (
                <div className="dex-ui-section">
                  <div className="dex-ui-section-title">{t('Kopf der Mail', 'Email header')}</div>
                  {field(
                    t('Betreff', 'Subject'),
                    <input className="dex-ui-input" value={emailSubject || ''} onChange={e => onEmailSubjectChange && onEmailSubjectChange(e.target.value)} />,
                    t('Steht in der Betreffzeile im Postfach.', 'Shows in the inbox subject line.'),
                  )}
                  <div className="dex-ui-grid-2">
                    <div>{field(
                      t('Überschrift', 'Heading'),
                      <input className="dex-ui-input" value={emailHeading || ''} onChange={e => onEmailHeadingChange && onEmailHeadingChange(e.target.value)} />,
                      t('Die große Zeile über dem Text.', 'The large line above the text.'),
                    )}</div>
                    <div>{field(
                      t('Unter-Überschrift', 'Subheading'),
                      <input className="dex-ui-input" value={emailSubheading || ''} placeholder={previewVars.EventTitle || '{{EventTitle}}'} onChange={e => onEmailSubheadingChange && onEmailSubheadingChange(e.target.value)} />,
                      t('Leer = nur der Event-Titel.', 'Empty = the event title only.'),
                      true,
                    )}</div>
                  </div>
                  {/* v18.19/v18.22: Größe, fett/kursiv und Farbe beider Zeilen —
                      v31.2 zusammen in einem Aufklapper statt zweimal unter
                      den Eingaben; die meisten lassen die Vorgabe. */}
                  {(headingFmt || subFmt) && (
                    <div style={{ marginTop: 6 }}>
                      {disclosure(t('Schrift und Farbe der Überschriften', 'Font and color of the headings'), undefined, (
                        <div className="dex-ui-stack" style={{ gap: 14 }}>
                          {headingFmt && (
                            <div>
                              <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 6 }}>{t('Überschrift', 'Heading')}</div>
                              {fmtRow({
                                sizeValue: emailHeadingFontSize || '26px', sizeDefault: '26px', sizeOptions: ['18px', '22px', '26px', '32px', '40px', '48px'], onSize: onEmailHeadingFontSizeChange,
                                boldActive: !!emailHeadingBold, onBold: onEmailHeadingBoldChange ? () => onEmailHeadingBoldChange(!emailHeadingBold) : undefined,
                                italicActive: !!emailHeadingItalic, onItalic: onEmailHeadingItalicChange ? () => onEmailHeadingItalicChange(!emailHeadingItalic) : undefined,
                                color: emailHeadingColor || '#86bc25', onColor: onEmailHeadingColorChange,
                              })}
                            </div>
                          )}
                          {subFmt && (
                            <div>
                              <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 6 }}>{t('Unter-Überschrift', 'Subheading')}</div>
                              {fmtRow({
                                sizeValue: emailSubheadingFontSize || '20px', sizeDefault: '20px', sizeOptions: ['14px', '16px', '18px', '20px', '24px', '28px', '32px'], onSize: onEmailSubheadingFontSizeChange,
                                boldActive: emailSubheadingBold !== false, onBold: onEmailSubheadingBoldChange ? () => onEmailSubheadingBoldChange(emailSubheadingBold === false) : undefined,
                                italicActive: !!emailSubheadingItalic, onItalic: onEmailSubheadingItalicChange ? () => onEmailSubheadingItalicChange(!emailSubheadingItalic) : undefined,
                                color: emailSubheadingColor || '#000000', onColor: onEmailSubheadingColorChange,
                              })}
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {previewMode === 'outlook' && (
                <>
                  {/* v18.42: Kalender-Metadaten — Betreff, Termin, Ort. */}
                  <div className="dex-ui-section">
                    <div className="dex-ui-section-title">{t('Termin im Kalender', 'Calendar entry')}</div>
                    <p className="dex-ui-section-desc">{t('So erscheint der Termin im Kalender der Teilnehmenden. Leere Felder übernehmen die Werte aus dem Event.', 'This is how the entry appears in attendees’ calendars. Empty fields take the values from the event.')}</p>
                    {onOutlookSubjectChange && field(
                      t('Betreff', 'Subject'),
                      <input className="dex-ui-input" value={outlookSubject || ''} onChange={e => onOutlookSubjectChange(e.target.value)} placeholder={previewVars.EventTitle || t('Event-Titel', 'Event title')} />,
                      t('Titel des Termins im Kalender — leer = Event-Titel.', 'Title of the entry in the calendar — empty = event title.'),
                      true,
                    )}
                    {/* v18.44: Termin (Datum) überschreibbar — DatePicker vom Parent. */}
                    {outlookDateEditor && field(
                      t('Termin', 'Date'),
                      outlookDateEditor,
                      t('Leer = aus Schritt „Grundlagen" bzw. vom Sub-Event-Datum übernommen.', 'Empty = taken from step “Basics” or the sub-event date.'),
                      true,
                    )}
                    {/* v18.44: Ort überschreibbar. */}
                    {onOutlookLocationChange && field(
                      t('Ort', 'Location'),
                      <input className="dex-ui-input" value={outlookLocationValue || ''} onChange={e => onOutlookLocationChange(e.target.value)} placeholder={outlookLocationAuto || t('Ort', 'Location')} />,
                      t('Leer = aus „Ort & Programm" übernommen.', 'Empty = taken from “Location & program”.'),
                      true,
                    )}
                  </div>
                  {/* v28.89: eigener, benannter Block — diese zwei Zeilen landen IM
                      Termin-Text, nicht in den Kalender-Metadaten. */}
                  <div className="dex-ui-section">
                    <div className="dex-ui-section-title">{t('Überschriften im Termin-Text', 'Headings in the appointment text')}</div>
                    <p className="dex-ui-section-desc">{t('Die zwei Zeilen, die im Termin über dem Text stehen — nicht der Betreff im Kalender (der steht oben).', 'The two lines above the text inside the appointment — not the calendar subject (that one is above).')}</p>
                    <div className="dex-ui-grid-2">
                      <div>{field(
                        t('Überschrift (grün)', 'Heading (green)'),
                        <input className="dex-ui-input" value={outlookHeading || ''} onChange={e => onOutlookHeadingChange && onOutlookHeadingChange(e.target.value)} placeholder={previewVars.EventTitle || t('Event-Titel', 'Event title')} />,
                        t('Die große Zeile im Termin-Text.', 'The large line in the appointment text.'),
                      )}</div>
                      <div>{field(
                        t('Unter-Überschrift (schwarz)', 'Subheading (black)'),
                        <input className="dex-ui-input" value={outlookSubheading || ''} onChange={e => onOutlookSubheadingChange && onOutlookSubheadingChange(e.target.value)} placeholder={previewVars.Location || previewVars.EventDate || 'Event Details'} />,
                      )}</div>
                    </div>
                  </div>
                </>
              )}

              {/* v18.73: Header-Bild (Event-Bild) Größe + Innenabstand — gilt
                  für Mail- UND Outlook-Kopf, daher in beiden Modi sichtbar.
                  v31.2: als Aufklapper — Feineinstellung, die selten jemand
                  anfasst; die Zeile zeigt den geltenden Wert. */}
              {(previewMode === 'email' || previewMode === 'outlook') && onImageWidthChange && (
                <div className="dex-ui-section">
                  <div className="dex-ui-section-title">{t('Kopfbild', 'Header image')}</div>
                  {disclosure(
                    t('Größe und Abstand des Bildes im Kopf', 'Size and spacing of the header image'),
                    fullWidthOn ? t('Volle Breite', 'Full width') : `${imgW} px`,
                    <div className="dex-ui-stack">
                      <p className="dex-ui-muted" style={{ margin: 0, lineHeight: 1.5 }}>
                        {t('Gilt für den Mail- und den Outlook-Termin-Kopf. Breites Foto → „Volle Breite" (füllt den Kopf, Höhe passt sich an). Rundes Logo (z.B. DEX-Orb) → kleinere Breite, dann steht es zentriert. Die Vorschau rechts zeigt es sofort.',
                          'Applies to the email and the Outlook header. Wide photo → “Full width” (fills the header, height adjusts). Round logo (e.g. the DEX orb) → smaller width, then it sits centered. The preview on the right shows it instantly.')}
                      </p>
                      {!ownHeaderImage && (
                        <div className="dex-ui-callout dex-ui-callout--neutral">
                          <span className="dex-ui-callout-icon"><Info size={16} /></span>
                          <span>
                            <strong>{t('Noch kein eigenes Mail-Logo:', 'No custom email logo yet:')}</strong>{' '}
                            {t('Der Kopf zeigt den DEX-Orb in fester Größe (max. 180 px). Die Werte hier greifen, sobald du im Reiter „Mail-Logo" ein Bild hinterlegst.',
                              'The header shows the DEX orb at a fixed size (max. 180 px). These values apply once you add an image in the “Email logo” tab.')}
                          </span>
                        </div>
                      )}
                      {/* v26.93: Ein-Klick „Volle Breite" (600/0/0) und zurück auf
                          Standard (180/30/30) — als Chips, die zeigen, was gilt. */}
                      <div className="dex-ui-inline">
                        <button
                          type="button"
                          className={cx('dex-ui-chip', fullWidthOn && 'is-active')}
                          onMouseDown={e => e.preventDefault()}
                          onClick={() => { onImageWidthChange(600); if (onImagePaddingHChange) onImagePaddingHChange(0); if (onImagePaddingVChange) onImagePaddingVChange(0); }}
                          title={t('Bild füllt den Kopf über die volle Breite (Höhe passt sich an)', 'Image fills the header edge to edge (height adjusts)')}
                        >
                          {t('Volle Breite', 'Full width')}
                        </button>
                        <button
                          type="button"
                          className={cx('dex-ui-chip', standardOn && 'is-active')}
                          onMouseDown={e => e.preventDefault()}
                          onClick={() => { onImageWidthChange(180); if (onImagePaddingHChange) onImagePaddingHChange(30); if (onImagePaddingVChange) onImagePaddingVChange(30); }}
                          title={t('Auf die Standard-Werte zurücksetzen (Breite 180, Abstand 30)', 'Back to the defaults (width 180, spacing 30)')}
                        >
                          {t('Standard (180 px, zentriert)', 'Default (180 px, centered)')}
                        </button>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'flex-end', gap: 14, flexWrap: 'wrap' }}>
                        <label style={numLabel}>
                          {t('Breite (px)', 'Width (px)')}
                          <input className="dex-ui-input dex-ui-input--sm" type="number" min={80} max={600} step={10} value={imageWidth ?? 180} onChange={e => onImageWidthChange(clampInt(e.target.value, 600, 180))} style={{ width: 96 }} />
                        </label>
                        {onImagePaddingHChange && (
                          <label style={numLabel}>
                            {t('Abstand seitlich (px)', 'Side spacing (px)')}
                            <input className="dex-ui-input dex-ui-input--sm" type="number" min={0} max={80} step={2} value={imagePaddingH ?? 30} onChange={e => onImagePaddingHChange(clampInt(e.target.value, 80, 30))} style={{ width: 96 }} />
                          </label>
                        )}
                        {onImagePaddingVChange && (
                          <label style={numLabel}>
                            {t('Abstand oben/unten (px)', 'Top/bottom spacing (px)')}
                            <input className="dex-ui-input dex-ui-input--sm" type="number" min={0} max={80} step={2} value={imagePaddingV ?? 30} onChange={e => onImagePaddingVChange(clampInt(e.target.value, 80, 30))} style={{ width: 96 }} />
                          </label>
                        )}
                        <span className="dex-ui-help" style={{ marginTop: 0, paddingBottom: 8 }}>{t('Breite 80–600 px, Abstände 0–80 px.', 'Width 80–600 px, spacing 0–80 px.')}</span>
                      </div>
                    </div>,
                  )}
                </div>
              )}

              {/* v28.89: Der eigentliche Text ist das, weswegen der Dialog
                  geöffnet wird — v31.2 in allen drei Modi ein eigener
                  Abschnitt, „Standardtext laden" rechts in der Titelzeile. */}
              {/* v31.10: Innerhalb des Schreibblocks rutschen auf dem Handy die
                  Chip-Reihen (Platzhalter, Vorlagen) UNTER den Editor — sie
                  standen sonst mit ihren zwei Zeilen Erklärung zwischen
                  Überschrift und Schreibfeld, und bei offener Tastatur war das
                  Feld damit vom Bildschirm geschoben. Eingefügt wird ohnehin an
                  der Cursor-Position, die Reihenfolge ist also frei. */}
              <div className="dex-ui-section" style={isMobile ? { order: -1, flexShrink: 0, display: 'flex', flexDirection: 'column' } : undefined}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 10 }}>
                  <div className="dex-ui-section-title" style={{ flex: 1, marginBottom: 0 }}>{textSectionTitle}</div>
                  {defaultBodyHtml && (
                    <button
                      type="button"
                      className="dex-ui-textbtn dex-ui-textbtn--muted"
                      onMouseDown={e => e.preventDefault()}
                      onClick={() => replaceBody(defaultBodyHtml, t('Den aktuellen Text durch den Standardtext ersetzen?', 'Replace the current text with the default text?'))}
                      title={t('Setzt den Text auf die Standard-Vorlage zurück (mit Platzhaltern)', 'Resets the text to the default template (with placeholders)')}
                    >
                      <RefreshCw size={14} /> {t('Standardtext laden', 'Load default text')}
                    </button>
                  )}
                </div>
                {insertableVars.length > 0 && (
                  <div className="dex-ui-field" style={isMobile ? { order: 1, marginTop: 14 } : undefined}>
                    <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 6 }}>
                      {t('Platzhalter einfügen — landet an der Cursor-Position und wird beim Versand durch den echten Wert ersetzt:', 'Insert a placeholder — it lands at the cursor and is replaced by the real value when sending:')}
                    </div>
                    <div className="dex-ui-inline" style={{ gap: 6 }}>
                      {insertableVars.map(v => (
                        <button
                          key={v.key}
                          type="button"
                          className="dex-ui-chip"
                          onMouseDown={e => { e.preventDefault(); }}
                          onClick={() => insertVariable(v.key)}
                          title={t(`${v.key} an der Cursor-Position einfügen`, `Insert ${v.key} at the cursor`)}
                        >
                          {v.label}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
                {/* v28.7: Vorlagen-Chips (z.B. Beschreibungs-Vorschläge). */}
                {bodyTemplates && bodyTemplates.length > 0 && (
                  <div className="dex-ui-field" style={isMobile ? { order: 1 } : undefined}>
                    {bodyTemplatesLabel && <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 6 }}>{bodyTemplatesLabel}</div>}
                    <div className="dex-ui-inline" style={{ gap: 6 }}>
                      {bodyTemplates.map(tpl => (
                        <button
                          key={tpl.key}
                          type="button"
                          className="dex-ui-chip"
                          onMouseDown={e => e.preventDefault()}
                          title={tpl.title}
                          onClick={() => replaceBody(tpl.html, t('Den aktuellen Text durch die Vorlage ersetzen?', 'Replace the current text with this template?'))}
                        >
                          {tpl.label}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
                {/* Toolbar — in Gruppen: Schrift | Größe/Abstand | Farbe | Listen | Link/Löschen
                    v31.10: Auf dem Handy EINE waagerecht rollende Zeile statt
                    Umbruch. Umgebrochen belegten dieselben Knöpfe auf 400 px
                    fünf Zeilen — rund 200 px, die dem Schreibfeld fehlten.
                    Gerollt wird mit dem Finger, und `order` (s. `tbOrder`)
                    stellt die Gruppen so um, dass die Knopf-Gruppen ohne Rollen
                    dastehen und die breiten Auswahlfelder hinten liegen. */}
                <div style={{
                  display: 'flex', flexWrap: isMobile ? 'nowrap' : 'wrap', gap: 2, padding: '6px 8px', alignItems: 'center',
                  ...(isMobile ? { overflowX: 'auto' as const, WebkitOverflowScrolling: 'touch' as const } : {}),
                  border: '1px solid var(--dex-gray-200, #e8e8e8)', borderBottom: 'none',
                  borderRadius: '10px 10px 0 0', background: 'var(--dex-gray-50, #fafafa)',
                }}>
                  {tb(t('Fett', 'Bold'), () => exec('bold'), <strong>B</strong>, activeFmt.bold)}
                  {tb(t('Kursiv', 'Italic'), () => exec('italic'), <em>I</em>, activeFmt.italic)}
                  {tb(t('Unterstrichen', 'Underline'), () => exec('underline'), <span style={{ textDecoration: 'underline' }}>U</span>, activeFmt.underline)}
                  {tbDivider(2)}
                  {/* v18.20: kontrolliertes Größen-Dropdown — zeigt (wie in Word)
                      die Größe des aktuell markierten Texts und setzt sie beim
                      Ändern. Live-Vorschau aktualisiert sich über fireChange(). */}
                  <select
                    className="dex-ui-select dex-ui-input--sm"
                    title={t('Schriftgröße der Auswahl', 'Font size of the selection')}
                    value={currentFontPx != null ? String(currentFontPx) : ''}
                    onMouseDown={() => saveSelection()}
                    onChange={e => { if (e.target.value) setFontSize(parseInt(e.target.value, 10)); }}
                    style={{ ...smallSelect, ...tbOrder(2) }}
                  >
                    <option value="" disabled>{t('Größe', 'Size')}</option>
                    {/* Aktuelle Größe sicher im Menü vorhalten, auch wenn sie nicht
                        in der Standardliste steht (z.B. 15px aus Alt-Inhalten). */}
                    {currentFontPx != null && FONT_SIZE_OPTIONS.indexOf(currentFontPx) === -1 && (
                      <option value={currentFontPx}>{currentFontPx} px</option>
                    )}
                    {FONT_SIZE_OPTIONS.map(px => (
                      <option key={px} value={px}>{px} px</option>
                    ))}
                  </select>
                  {/* v18.16: Zeilenabstand für die gesamte Beschreibung. */}
                  <select
                    className="dex-ui-select dex-ui-input--sm"
                    title={t('Zeilenabstand', 'Line spacing')}
                    onChange={e => { if (e.target.value) { setLineHeight(e.target.value); e.target.value = ''; } }}
                    defaultValue=""
                    style={{ ...smallSelect, ...tbOrder(2) }}
                  >
                    <option value="" disabled>{t('Zeilenabstand', 'Line spacing')}</option>
                    <option value="1.0">{t('Eng (1.0)', 'Tight (1.0)')}</option>
                    <option value="1.15">{t('Kompakt (1.15)', 'Compact (1.15)')}</option>
                    <option value="1.3">{t('Normal (1.3)', 'Normal (1.3)')}</option>
                    <option value="1.5">{t('Locker (1.5)', 'Relaxed (1.5)')}</option>
                    <option value="2.0">{t('Weit (2.0)', 'Wide (2.0)')}</option>
                  </select>
                  {tbDivider(3)}
                  {/* v31.10: Farbpunkte, Farbwähler und Hex-Feld in EINER Hülle —
                      damit die Gruppe auf dem Handy als Ganzes ans Ende der
                      rollenden Leiste wandern kann (order). Am Desktop ist die
                      Hülle nur ein zusätzliches inline-flex ohne eigene Optik. */}
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 2, flexShrink: 0, ...tbOrder(3) }}>
                  {COLORS.map(c => <ColorDot key={c} color={c} active={false} onPick={setColor} />)}
                  {/* v18.22: freie Body-Textfarbe — nativer Picker + Hex-Eingabe.
                      Auswahl wird vor dem Fokuswechsel gesichert (onMouseDown). */}
                  <input
                    type="color"
                    value={isHex6(bodyHexDraft) ? bodyHexDraft : '#000000'}
                    title={t('Freie Textfarbe wählen', 'Pick any text color')}
                    onMouseDown={() => saveSelection()}
                    onChange={e => { setBodyHexDraft(e.target.value.toLowerCase()); setColor(e.target.value.toLowerCase()); }}
                    style={isMobile ? { ...colorPickerStyle, width: 40, height: 40, flexShrink: 0 } : colorPickerStyle}
                  />
                  <input
                    className="dex-ui-input dex-ui-input--sm"
                    value={bodyHexDraft}
                    onMouseDown={() => saveSelection()}
                    onChange={e => setBodyHexDraft(e.target.value)}
                    onKeyDown={e => {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        let v = bodyHexDraft.trim();
                        if (v && !v.startsWith('#')) v = '#' + v;
                        if (/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(v)) setColor(v.toLowerCase());
                      }
                    }}
                    placeholder="#RRGGBB"
                    maxLength={7}
                    title={t('Hex-Code eingeben + Enter — färbt die Auswahl', 'Type a hex code + Enter — colors the selection')}
                    style={isMobile ? { ...hexInputStyle, minHeight: 40, flexShrink: 0 } : hexInputStyle}
                  />
                  </span>
                  {tbDivider(1)}
                  {tb(t('Aufzählung', 'Bullet list'), () => exec('insertUnorderedList'), '•', undefined, 1)}
                  {tb(t('Nummerierte Liste', 'Numbered list'), () => exec('insertOrderedList'), '1.', undefined, 1)}
                  {tbDivider(1)}
                  {/* v18.39: Link einfügen/bearbeiten — Text markieren + klicken,
                      oder ohne Auswahl klicken für einen neuen Link. Bestehenden
                      Link: Cursor hineinsetzen → URL ändern. */}
                  {tb(t('Link oder E-Mail-Adresse einfügen / bearbeiten', 'Insert / edit a link or email address'), openLinkDialog, <Link2 size={15} />, undefined, 1)}
                  {/* v31.7: Bild an die Cursor-Position. Anhänge gehen nicht
                      (der Deloitte-Mailflow schickt Mails mit Anhang als NDR
                      zurück) — das Bild reist deshalb als Base64 IM HTML mit,
                      genau wie QR-Code und Kopfbild. */}
                  {tb(
                    imgBusy
                      ? t('Bild wird verkleinert …', 'Shrinking the image …')
                      : t('Bild einfügen — es reist in der Mail mit, ohne Anhang', 'Insert an image — it travels inside the email, no attachment'),
                    () => { if (imgBusy) return; saveSelection(); imgInputRef.current?.click(); },
                    <ImageIcon size={15} />,
                    undefined,
                    1,
                  )}
                  {tb(t('Formatierung entfernen', 'Clear formatting'), () => exec('removeFormat'), '⌫', undefined, 1)}
                </div>
                {/* Der Datei-Wähler hängt am Knopf oben; `value` wird nach jeder
                    Wahl geleert, sonst löst dieselbe Datei beim zweiten Mal kein
                    `change` aus. */}
                <input
                  ref={imgInputRef}
                  type="file"
                  accept="image/*"
                  style={{ display: 'none' }}
                  onChange={e => {
                    const f = (e.target.files && e.target.files[0]) || null;
                    e.target.value = '';
                    void insertInlineImage(f);
                  }}
                />
                <div
                  ref={editorRef}
                  contentEditable
                  suppressContentEditableWarning
                  onInput={fireChange}
                  onPaste={handlePaste}
                  onBlur={() => { saveSelection(); fireChange(); }}
                  onMouseUp={syncSelection}
                  onKeyUp={syncSelection}
                  onKeyDown={e => {
                    // Enter = <br> (E-Mail-konform); Shift+Enter = neuer Absatz
                    if (e.key === 'Enter' && !e.shiftKey) {
                      e.preventDefault();
                      try { document.execCommand('insertLineBreak'); } catch {
                        try { document.execCommand('insertHTML', false, '<br>'); } catch { /* ignore */ }
                      }
                      fireChange();
                    }
                  }}
                  style={{
                    // v31.10: Auf dem Handy 200 px statt 280 — mit offener
                    // Tastatur bleibt vom Dialog wenig übrig, und ein
                    // Schreibfeld, das man erst scrollen muss, ist keines.
                    minHeight: isMobile ? 200 : 280, padding: '12px 14px',
                    border: '1px solid var(--dex-gray-200, #e8e8e8)',
                    borderRadius: '0 0 10px 10px',
                    // v18.19: WYSIWYG — bei Mail/Outlook-Bodies exakt die echte
                    // Mail-Basis (Aptos 14px / line-height 1.6) verwenden, damit
                    // „was du im Editor siehst = was der Empfänger bekommt".
                    // Vorher 0.9rem ≈ 14.4px → Organizer haben Größen
                    // „nachjustiert", was zwischen Sub-Events zu uneinheitlichen
                    // Schriftgrößen führte. Bei der Beschreibung ('plain') bleibt
                    // die etwas größere Editor-Schrift.
                    ...(previewMode === 'email' || previewMode === 'outlook'
                      ? { fontFamily: 'Aptos, Arial, Helvetica, sans-serif', fontSize: '14px', lineHeight: 1.6 }
                      : { fontSize: '0.9rem', lineHeight: 1.5 }),
                    outline: 'none', overflowY: 'auto',
                    background: '#fff',
                  }}
                />
                {/* v31.7: Was gerade passiert ist — und wie voll der Text ist.
                    Ein eingefügtes Bild sieht man, seine Größe nicht; und die
                    Grenze taucht sonst erst beim Senden auf, wenn SharePoint
                    die Zeile ablehnt. Die Dauerzeile erscheint nur, wenn
                    wirklich ein Bild im Text steckt — ein reiner Textbrief hat
                    mit dem Budget nichts zu tun. */}
                {imgBusy && (
                  <div className="dex-ui-help" role="status">{t('Bild wird verkleinert …', 'Shrinking the image …')}</div>
                )}
                {imgNote && (
                  <div
                    className={cx('dex-ui-callout', 'dex-ui-callout--sm', imgNote.tone === 'ok' ? 'dex-ui-callout--success' : 'dex-ui-callout--warn')}
                    style={{ marginTop: 8 }}
                    role="status"
                  >
                    <span>{imgNote.text}</span>
                  </div>
                )}
                {inlineImgCount > 0 && (
                  <div className="dex-ui-help" style={{ marginTop: 6 }}>
                    {isDe
                      ? `${inlineImgCount === 1 ? 'Ein Bild' : `${inlineImgCount} Bilder`} im Text · insgesamt ${charsToKb(bodyChars)} von höchstens ${MAIL_BODY_MAX_KB} KB.${bodyChars > MAIL_BODY_WARN_KB * 1024 ? ' Viel Platz ist nicht mehr — ein weiteres Bild passt vermutlich nicht.' : ''}`
                      : `${inlineImgCount === 1 ? 'One image' : `${inlineImgCount} images`} in the text · ${charsToKb(bodyChars)} of at most ${MAIL_BODY_MAX_KB} KB in total.${bodyChars > MAIL_BODY_WARN_KB * 1024 ? ' Not much room left — another image probably will not fit.' : ''}`}
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* === VORSCHAU ===
              v31.10: Auf dem Handy wird sie nur gebaut, wenn ihr Reiter offen
              ist. Anders als die Editor-Spalte darf sie das: Der iframe hält
              keinen Zustand, sein Inhalt entsteht bei jedem Rendern neu aus
              `value` — und genau deshalb kostet ein verstecktes Neubauen bei
              jedem Tastendruck auf dem Handy unnötig Rechenzeit. */}
          {showPreview && (
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', width: isMobile ? '100%' : undefined, background: 'var(--dex-gray-100, #f5f5f5)', minHeight: 0 }}>
            {colHead(isMobile ? '' : t('Vorschau', 'Preview'), (
              <>
                <span className="dex-ui-pill dex-ui-pill--gray">{modeLabel}</span>
                <span className="dex-ui-muted" style={{ fontSize: '0.74rem' }}>{t('Platzhalter mit Beispielwerten', 'Placeholders filled with sample values')}</span>
              </>
            ))}
            {/* v22.5: „Briefumschlag"-Kopf über der Mail-Vorschau — Empfänger
                („An") + Betreff, damit man die komplette Mail wie im Postfach
                sieht. Nur wenn vom Aufrufer befüllt (Einladungsmail). */}
            {previewMode === 'email' && (previewToLine || previewSubjectLine) && (
              <div style={{ background: '#fff', borderBottom: '1px solid var(--dex-gray-200, #e8e8e8)', padding: '10px 18px', fontSize: '0.82rem', color: 'var(--dex-gray-700, #444)', display: 'grid', gridTemplateColumns: 'auto 1fr', columnGap: 12, rowGap: 4 }}>
                {previewToLine && (
                  <>
                    <span style={envKey}>{t('An', 'To')}</span>
                    <span style={{ wordBreak: 'break-word' }}>{previewToLine}</span>
                  </>
                )}
                {previewSubjectLine && (
                  <>
                    <span style={envKey}>{t('Betreff', 'Subject')}</span>
                    <span style={{ fontWeight: 600, color: 'var(--dex-gray-800, #333)', wordBreak: 'break-word' }}>{previewSubjectLine}</span>
                  </>
                )}
              </div>
            )}
            <iframe
              title={t('Vorschau', 'Preview')}
              srcDoc={renderPreviewHtml()}
              sandbox=""
              // v31.10: Auf dem Handy niedrigere Untergrenze — 360 px sprengen
              // den Dialog, sobald der Bildschirm kurz ist.
              style={{ flex: 1, border: 'none', width: '100%', minHeight: isMobile ? 220 : 360, background: '#f5f5f5' }}
            />
          </div>
          )}
        </div>

        {/* v31.2: Fuß wie in `Modal` — genau ein Primärknopf (die Aktion des
            Aufrufers), links daneben Abbrechen bzw. „Fertig", wenn es nichts
            zu senden gibt. Der Hinweis „Vorschau zeigt jede Änderung sofort"
            stand hier doppelt — der Untertitel im Kopf sagt dasselbe. */}
        <div className="dex-ui-modal-foot" style={{ padding: '12px 20px 14px', marginTop: 0, flexShrink: 0 }}>
          <button type="button" className="btn btn-secondary" onClick={onClose}>{extraAction ? t('Abbrechen', 'Cancel') : t('Fertig', 'Done')}</button>
          {extraAction && (
            <button
              type="button"
              className="btn btn-primary"
              disabled={extraAction.disabled}
              onClick={() => { void extraAction.onClick(); }}
              style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}
            >
              {extraAction.icon}
              {extraAction.label}
            </button>
          )}
        </div>
      </div>

      {/* v30.51: Link-Dialog — Ziel, Anzeige-Text und Web/E-Mail in EINEM
          Fenster. Liegt bewusst innerhalb des Editor-Modals im React-Baum;
          gerendert wird er ohnehin per Portal an document.body. */}
      <LinkDialog
        open={!!linkState}
        initialHref={linkState?.href}
        initialText={linkState?.text}
        editing={!!linkState?.editing}
        textLocked={!!linkState?.textLocked}
        onCancel={() => { setLinkState(null); restoreSelection(); }}
        onApply={r => applyLink(r.href, r.text)}
        onRemove={removeLink}
      />
    </div>
  );
};
