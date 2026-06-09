/**
 * Wiederverwendbares Modal mit HTML-Editor + Live-Vorschau.
 *
 * Wird genutzt fuer:
 *   - Outlook-Termin-Body (Calendar Description)
 *   - E-Mail-Templates (Anmeldung, Warteliste, Abmeldung, Nachruecken)
 *
 * Layout (split): links der Editor (contentEditable + Toolbar + Variable-Buttons),
 * rechts die gerenderte Vorschau im echten Wrapper (Outlook-Calendar bzw. Deloitte-Mail).
 *
 * Die Vorschau wird in einem sandboxed iframe gerendert, damit das Wrapper-CSS
 * nicht mit dem App-CSS kollidiert.
 */
import * as React from 'react';
import { X } from './Icons';
import { Icon } from '@fluentui/react/lib/Icon';
import { wrapTemplate, replacePlaceholders, replacePlaceholdersPlain, getCachedOrbBase64, getCachedLogoBase64 } from '../services/EmailTemplates';
// v20.4: modernes Confirm-Modal statt window.confirm.
import { useDialog } from '../context/DialogContext';

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
  /** Outlook-Termin: editierbare Ueberschrift (<h1>) */
  outlookHeading?: string;
  onOutlookHeadingChange?: (s: string) => void;
  /** Outlook-Termin: editierbare Unter-Ueberschrift (<h2>) */
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
  insertableVars?: Array<{ key: string; label: string }>;
  logoBase64?: string;
  imageBase64?: string;
  /** Optional: zusaetzlicher primaerer Button im Footer (z.B. "Senden") */
  extraAction?: {
    label: string;
    onClick: () => void | Promise<void>;
    disabled?: boolean;
    icon?: React.ReactNode;
  };
  /** v11.40: Optionaler React-Knoten oberhalb von Subject/Ueberschrift im
   *  Editor — z.B. fuer eine Ziel-Auswahl im Einladungsmail-Modal. */
  headerExtra?: React.ReactNode;
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
  React.useEffect(() => { setHexDraft(value); }, [value]);
  const commitHex = (h: string): void => {
    let v = (h || '').trim();
    if (v && !v.startsWith('#')) v = '#' + v;
    if (/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(v)) onChange(v.toLowerCase());
    else setHexDraft(value);
  };
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
      {COLORS.map(c => (
        <button
          key={c}
          type="button"
          title={c}
          onMouseDown={e => e.preventDefault()}
          onClick={() => onChange(c)}
          style={{
            width: 20, height: 20, borderRadius: '50%', background: c, cursor: 'pointer', padding: 0,
            border: (value || '').toLowerCase() === c.toLowerCase() ? '2px solid var(--dex-gray-700)' : '1px solid var(--dex-gray-300)',
            boxShadow: 'inset 0 0 0 1px rgba(0,0,0,0.08)',
          }}
        />
      ))}
      <input
        type="color"
        value={isHex6(value) ? value : '#000000'}
        onChange={e => onChange(e.target.value.toLowerCase())}
        title="Freie Farbe wählen"
        style={{ width: 26, height: 26, padding: 0, border: '1px solid var(--dex-gray-300)', borderRadius: 4, cursor: 'pointer', background: '#fff' }}
      />
      <input
        value={hexDraft}
        onChange={e => setHexDraft(e.target.value)}
        onBlur={() => commitHex(hexDraft)}
        onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); commitHex(hexDraft); } }}
        placeholder="#RRGGBB"
        maxLength={7}
        style={{ width: 76, height: 24, fontSize: '0.72rem', fontFamily: 'monospace', borderRadius: 4, border: '1px solid var(--dex-gray-300)', padding: '0 6px' }}
      />
    </span>
  );
};

// v18.22: Fett/Kursiv-Umschalter (für Überschrift + Unter-Überschrift).
const FmtToggle: React.FC<{ active: boolean; onClick: () => void; title: string; children: React.ReactNode }> = ({ active, onClick, title, children }) => (
  <button
    type="button"
    title={title}
    onMouseDown={e => e.preventDefault()}
    onClick={onClick}
    style={{
      minWidth: 28, height: 26, padding: '0 8px', borderRadius: 4, cursor: 'pointer',
      border: active ? '1px solid var(--dex-green, #86bc25)' : '1px solid var(--dex-gray-300)',
      background: active ? 'var(--dex-green, #86bc25)' : '#fff',
      color: active ? '#fff' : 'var(--dex-gray-800)', fontSize: '0.82rem',
    }}
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
    logoBase64 = '', imageBase64 = '',
    extraAction,
    headerExtra,
  } = props;

  const editorRef = React.useRef<HTMLDivElement>(null);
  // v20.4: App-Modals statt window.confirm/window.prompt.
  const { confirmDialog, promptDialog } = useDialog();
  const savedSelectionRef = React.useRef<Range | null>(null);
  // v18.20: aktuelle Schriftgröße der Auswahl (px) — treibt die „wie in Word"-
  // Anzeige im Größen-Dropdown. null = keine Auswahl im Editor / unbekannt.
  const [currentFontPx, setCurrentFontPx] = React.useState<number | null>(null);
  // v18.22: freier Hex-Code für die Body-Textfarbe (Picker/Eingabe).
  const [bodyHexDraft, setBodyHexDraft] = React.useState('#000000');

  // External value beim Oeffnen in den Editor laden (Re-Open mit anderem Template)
  React.useEffect(() => {
    if (open && editorRef.current) {
      const cur = editorRef.current.innerHTML;
      if (cur !== (value || '')) {
        editorRef.current.innerHTML = value || '';
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
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

  // Auswahl sichern UND Größen-Anzeige aktualisieren (mouseup/keyup im Editor).
  const syncSelection = (): void => { saveSelection(); detectFontSize(); };

  const restoreSelection = (): void => {
    const range = savedSelectionRef.current;
    if (range && editorRef.current) {
      editorRef.current.focus();
      const sel = window.getSelection();
      if (sel) { sel.removeAllRanges(); sel.addRange(range); }
    } else {
      editorRef.current?.focus();
    }
  };

  const fireChange = (): void => {
    if (editorRef.current) onChange(editorRef.current.innerHTML);
  };

  const exec = (cmd: string, arg?: string): void => {
    restoreSelection();
    try { document.execCommand(cmd, false, arg); } catch { /* ignore */ }
    fireChange();
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
  const insertLink = (): void => {
    restoreSelection();
    const existing = getSelectionAnchor();
    // v20.4: App-Prompt statt window.prompt. WICHTIG: nach jedem Dialog die
    // gespeicherte Editor-Selektion wiederherstellen (das Modal stiehlt den
    // Fokus), sonst landet execCommand ins Leere.
    promptDialog('Link-Adresse (URL) — leer lassen, um den Link zu entfernen:', {
      title: 'Link',
      defaultValue: existing ? (existing.getAttribute('href') || '') : 'https://',
      confirmLabel: 'Übernehmen',
    }).then(async url => {
      if (url === null) return; // Abbrechen
      restoreSelection();
      const sel = window.getSelection();
      const collapsed = !sel || sel.rangeCount === 0 || sel.getRangeAt(0).collapsed;
      if (url.trim() === '') {
        try { document.execCommand('unlink'); } catch { /* ignore */ }
        fireChange();
        return;
      }
      if (existing) {
        // Bestehender Link: nur die Ziel-URL aktualisieren, Anzeige-Text bleibt.
        existing.setAttribute('href', url.trim());
      } else if (collapsed) {
        // Keine Auswahl: Anzeige-Text erfragen und sauberen Link einfügen.
        const text = (await promptDialog('Anzeige-Text des Links:', { title: 'Link', defaultValue: url.trim(), confirmLabel: 'Einfügen' })) || url.trim();
        restoreSelection();
        try { document.execCommand('insertHTML', false, `<a href="${url.trim()}">${escHtml(text)}</a>`); } catch { /* ignore */ }
      } else {
        // Markierten Text in einen Link verwandeln.
        try { document.execCommand('createLink', false, url.trim()); } catch { /* ignore */ }
      }
      fireChange();
    }).catch(() => { /* */ });
  };

  // v18.39: Einfügen IMMER als reiner Text (mit Zeilenumbrüchen als <br>).
  // Verhindert, dass kopierte Inhalte Block-Markup (<div>/<p> mit Außen-
  // abständen) mitbringen, das den Zeilenabstand „plötzlich größer" macht
  // und sich danach nicht mehr korrigieren lässt.
  const handlePaste = (e: React.ClipboardEvent<HTMLDivElement>): void => {
    e.preventDefault();
    const text = e.clipboardData?.getData('text/plain') || '';
    if (!text) return;
    const html = text.split(/\r\n|\r|\n/).map(line => escHtml(line)).join('<br>');
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
    // execCommand insertText fuegt an Cursor ein (oder ersetzt Selektion).
    // Variable als Text — der Server ersetzt sie spaeter per replacePlaceholders.
    try { document.execCommand('insertText', false, key); } catch { /* ignore */ }
    fireChange();
  };

  const renderPreviewHtml = (): string => {
    const bodyWithVars = replacePlaceholders(value || '', previewVars);
    const cachedLogo = getCachedLogoBase64();
    const cachedOrb = getCachedOrbBase64();

    if (previewMode === 'plain') {
      // v9.40: kein Wrapper — die Beschreibung wird auf der Anmelde-Seite genauso
      // gerendert, eingebettet in die normale Seitentypografie. Wir simulieren das
      // mit minimalen Default-Styles (System-Font + 14px), damit der Editor-Text
      // nicht in irgendeinem rohen Browser-Default aussieht.
      const empty = !bodyWithVars || bodyWithVars.trim() === '';
      const html = empty
        ? '<p style="color:#999;font-style:italic;">Hier erscheint die Beschreibung, sobald du im Editor links etwas tippst — sie wird 1:1 auf der Anmelde-Seite angezeigt.</p>'
        : bodyWithVars;
      return `<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica,Arial,sans-serif;font-size:14px;line-height:1.55;color:#333;margin:0;padding:24px;background:#fff;}img{max-width:100%;height:auto;}a{color:#0076a8;}h1,h2,h3,h4{color:#222;}</style></head><body>${html}</body></html>`;
    }

    if (previewMode === 'email') {
      const heading = replacePlaceholdersPlain(emailHeading || '', previewVars);
      // v15.19: Subheading aus dem Override; Fallback {{EventTitle}}.
      const rawSub = (emailSubheading && emailSubheading.trim()) || '{{EventTitle}}';
      const subheading = replacePlaceholdersPlain(rawSub, previewVars);
      // Manche System-Templates (z.B. Nachruecken, OutlookDeclineReminder)
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
          imageWidth, imagePaddingV, imagePaddingH,
        });
      return wrapped
        .replace(/\{\{LOGO_URL\}\}/g, logoBase64 || cachedLogo || '')
        .replace(/\{\{ORB_URL\}\}/g, imageBase64 || cachedOrb || '');
    }

    // Outlook-Termin-Vorschau: gleicher wrapTemplate() wie beim echten Versand,
    // damit der User in der Vorschau genau das sieht, was nachher im Termin steht
    // (inkl. Deloitte-Signatur + Legal-Disclaimer im Footer).
    const olHeading = replacePlaceholdersPlain(outlookHeading || '', previewVars) || previewVars.EventTitle || 'Event Title';
    const olSub = replacePlaceholdersPlain(outlookSubheading || '', previewVars) || previewVars.EventDate || 'Event Details';
    const bodyForOutlook = bodyWithVars || '<p style="color:#999;font-style:italic;">Hier erscheint der Body — beginne im Editor links zu tippen.</p>';
    // Wenn der Body bereits ein kompletter wrapTemplate-Output ist (z.B. aus editEvent
    // ohne Strip), 1:1 anzeigen — sonst doppelt wickeln.
    const isAlreadyWrapped = /<!doctype|<html/i.test(bodyForOutlook);
    const wrapped = isAlreadyWrapped
      ? bodyForOutlook
      // v18.73: Header-Bild Größe + Innenabstand live mitvorschauen.
      : wrapTemplate('#86bc25', olHeading, olSub, bodyForOutlook, undefined, { imageWidth, imagePaddingV, imagePaddingH });
    return wrapped
      .replace(/\{\{LOGO_URL\}\}/g, logoBase64 || cachedLogo || '')
      .replace(/\{\{ORB_URL\}\}/g, imageBase64 || cachedOrb || '');
  };

  const tbBtn: React.CSSProperties = {
    minWidth: 30, height: 28, padding: '0 8px',
    border: '1px solid var(--dex-gray-300)', borderRadius: 4,
    background: '#fff', cursor: 'pointer',
    fontSize: '0.85rem', fontWeight: 600,
  };

  return (
    <div
      style={{
        position: 'fixed', inset: 0, zIndex: 1200,
        background: 'rgba(0,0,0,0.55)',
        display: 'flex', alignItems: 'stretch', justifyContent: 'center',
        padding: '32px 24px',
      }}
      // v18.43: KEIN Schließen per Backdrop-Klick mehr — der Editor enthält
      // viel Arbeit (Body, Betreff, Logos) und wurde durch einen Fehlklick
      // neben das Modal zu leicht geschlossen. Schließen nur noch aktiv über
      // das X oben rechts (oder „Fertig").
    >
      <div
        className="card"
        style={{
          width: '100%', maxWidth: 1280, maxHeight: '100%',
          display: 'flex', flexDirection: 'column',
          background: '#fff', borderRadius: 'var(--dex-radius)', overflow: 'hidden',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 20px', borderBottom: '1px solid var(--dex-gray-200)' }}>
          <h3 style={{ margin: 0, fontSize: '1.05rem' }}>{title}</h3>
          <button
            type="button"
            onClick={onClose}
            style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4 }}
            aria-label="Schließen"
          >
            <X size={20} />
          </button>
        </div>

        <div style={{ flex: 1, display: 'flex', minHeight: 0, overflow: 'hidden' }}>
          {/* === EDITOR === */}
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', borderRight: '1px solid var(--dex-gray-200)', overflow: 'auto' }}>
            <div style={{ padding: '14px 20px', display: 'flex', flexDirection: 'column', gap: 10 }}>
              {headerExtra}
              {previewMode === 'email' && (
                <>
                  <div>
                    <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>Subject</label>
                    <input
                      className="form-input"
                      value={emailSubject || ''}
                      onChange={e => onEmailSubjectChange && onEmailSubjectChange(e.target.value)}
                      style={{ fontSize: '0.85rem' }}
                    />
                  </div>
                  <div>
                    <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>Überschrift</label>
                    <input
                      className="form-input"
                      value={emailHeading || ''}
                      onChange={e => onEmailHeadingChange && onEmailHeadingChange(e.target.value)}
                      style={{ fontSize: '0.85rem' }}
                    />
                    {/* v18.19/v18.22: Überschrift Größe + freie Farbe + fett/kursiv. */}
                    {(onEmailHeadingFontSizeChange || onEmailHeadingColorChange || onEmailHeadingBoldChange) && (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 6, flexWrap: 'wrap' }}>
                        {onEmailHeadingFontSizeChange && (
                          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>
                            Größe
                            <select
                              value={emailHeadingFontSize || '26px'}
                              onChange={e => onEmailHeadingFontSizeChange(e.target.value)}
                              style={{ height: 26, fontSize: '0.78rem', borderRadius: 4, border: '1px solid var(--dex-gray-300)' }}
                            >
                              {['18px', '22px', '26px', '32px', '40px', '48px'].map(px => (
                                <option key={px} value={px}>{px}{px === '26px' ? ' (Standard)' : ''}</option>
                              ))}
                            </select>
                          </label>
                        )}
                        {(onEmailHeadingBoldChange || onEmailHeadingItalicChange) && (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                            {onEmailHeadingBoldChange && (
                              <FmtToggle active={!!emailHeadingBold} onClick={() => onEmailHeadingBoldChange(!emailHeadingBold)} title="Fett"><strong>F</strong></FmtToggle>
                            )}
                            {onEmailHeadingItalicChange && (
                              <FmtToggle active={!!emailHeadingItalic} onClick={() => onEmailHeadingItalicChange(!emailHeadingItalic)} title="Kursiv"><em>K</em></FmtToggle>
                            )}
                          </span>
                        )}
                        {onEmailHeadingColorChange && (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>
                            Farbe
                            <ColorControl value={emailHeadingColor || '#86bc25'} onChange={onEmailHeadingColorChange} />
                          </span>
                        )}
                      </div>
                    )}
                  </div>
                  <div>
                    <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>
                      Unter-Überschrift <span style={{ color: 'var(--dex-gray-400)', fontWeight: 400 }}>(leer = nur Event-Titel)</span>
                    </label>
                    <input
                      className="form-input"
                      value={emailSubheading || ''}
                      placeholder={previewVars.EventTitle || '{{EventTitle}}'}
                      onChange={e => onEmailSubheadingChange && onEmailSubheadingChange(e.target.value)}
                      style={{ fontSize: '0.85rem' }}
                    />
                    {/* v18.22: Unter-Überschrift Größe + freie Farbe + fett/kursiv. */}
                    {(onEmailSubheadingFontSizeChange || onEmailSubheadingColorChange || onEmailSubheadingBoldChange) && (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 6, flexWrap: 'wrap' }}>
                        {onEmailSubheadingFontSizeChange && (
                          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>
                            Größe
                            <select
                              value={emailSubheadingFontSize || '20px'}
                              onChange={e => onEmailSubheadingFontSizeChange(e.target.value)}
                              style={{ height: 26, fontSize: '0.78rem', borderRadius: 4, border: '1px solid var(--dex-gray-300)' }}
                            >
                              {['14px', '16px', '18px', '20px', '24px', '28px', '32px'].map(px => (
                                <option key={px} value={px}>{px}{px === '20px' ? ' (Standard)' : ''}</option>
                              ))}
                            </select>
                          </label>
                        )}
                        {(onEmailSubheadingBoldChange || onEmailSubheadingItalicChange) && (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                            {onEmailSubheadingBoldChange && (
                              <FmtToggle active={emailSubheadingBold !== false} onClick={() => onEmailSubheadingBoldChange(emailSubheadingBold === false)} title="Fett"><strong>F</strong></FmtToggle>
                            )}
                            {onEmailSubheadingItalicChange && (
                              <FmtToggle active={!!emailSubheadingItalic} onClick={() => onEmailSubheadingItalicChange(!emailSubheadingItalic)} title="Kursiv"><em>K</em></FmtToggle>
                            )}
                          </span>
                        )}
                        {onEmailSubheadingColorChange && (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>
                            Farbe
                            <ColorControl value={emailSubheadingColor || '#000000'} onChange={onEmailSubheadingColorChange} />
                          </span>
                        )}
                      </div>
                    )}
                  </div>
                </>
              )}

              {previewMode === 'outlook' && (
                <>
                  {/* v18.42: Outlook-Übersicht — Betreff (bearbeitbar) + Termin/Ort (read-only). */}
                  <div style={{ background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200, #eee)', borderRadius: 8, padding: '12px 14px', marginBottom: 6 }}>
                    <div style={{ fontSize: '0.72rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', letterSpacing: 0.4, marginBottom: 8 }}>
                      OUTLOOK-TERMIN
                    </div>
                    {onOutlookSubjectChange && (
                      <div style={{ marginBottom: 10 }}>
                        <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>
                          Betreff <span style={{ color: 'var(--dex-gray-400)', fontWeight: 400 }}>(Titel des Termins im Kalender — leer = Event-Titel)</span>
                        </label>
                        <input
                          className="form-input"
                          value={outlookSubject || ''}
                          onChange={e => onOutlookSubjectChange(e.target.value)}
                          placeholder={previewVars.EventTitle || 'Event-Titel'}
                          style={{ fontSize: '0.85rem' }}
                        />
                      </div>
                    )}
                    {/* v18.44: Termin (Datum) überschreibbar — DatePicker vom Parent. */}
                    {outlookDateEditor && (
                      <div style={{ marginBottom: 10 }}>
                        <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>
                          Termin <span style={{ color: 'var(--dex-gray-400)', fontWeight: 400 }}>(leer = aus Schritt &bdquo;Grundlagen&ldquo; bzw. Sub-Event-Datum übernommen)</span>
                        </label>
                        {outlookDateEditor}
                      </div>
                    )}
                    {/* v18.44: Ort überschreibbar. */}
                    {onOutlookLocationChange && (
                      <div>
                        <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>
                          Ort <span style={{ color: 'var(--dex-gray-400)', fontWeight: 400 }}>(leer = aus Reiter &bdquo;Ort &amp; Programm&ldquo; übernommen)</span>
                        </label>
                        <input
                          className="form-input"
                          value={outlookLocationValue || ''}
                          onChange={e => onOutlookLocationChange(e.target.value)}
                          placeholder={outlookLocationAuto || 'Ort'}
                          style={{ fontSize: '0.85rem' }}
                        />
                      </div>
                    )}
                  </div>
                  <div>
                    <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>Überschrift (grün) <span style={{ color: 'var(--dex-gray-400)', fontWeight: 400 }}>(große Zeile IM Termin-Text)</span></label>
                    <input
                      className="form-input"
                      value={outlookHeading || ''}
                      onChange={e => onOutlookHeadingChange && onOutlookHeadingChange(e.target.value)}
                      placeholder={previewVars.EventTitle || 'Event-Titel'}
                      style={{ fontSize: '0.85rem' }}
                    />
                  </div>
                  <div>
                    <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>Unter-Überschrift (schwarz)</label>
                    <input
                      className="form-input"
                      value={outlookSubheading || ''}
                      onChange={e => onOutlookSubheadingChange && onOutlookSubheadingChange(e.target.value)}
                      placeholder={previewVars.EventDate || 'Event Details'}
                      style={{ fontSize: '0.85rem' }}
                    />
                  </div>
                </>
              )}

              {/* v18.73: Header-Bild (Event-Bild) Größe + Innenabstand — gilt
                  für Mail- UND Outlook-Kopf, daher in beiden Modi sichtbar. */}
              {(previewMode === 'email' || previewMode === 'outlook') && onImageWidthChange && (
                <div style={{ background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200, #eee)', borderRadius: 8, padding: '12px 14px' }}>
                  <div style={{ fontSize: '0.72rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', letterSpacing: 0.4, marginBottom: 6 }}>
                    HEADER-BILD
                  </div>
                  <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', margin: '0 0 10px', lineHeight: 1.45 }}>
                    Größe und Innenabstand des Bildes im Kopf. Gilt für den <strong>Mail-</strong> und den <strong>Outlook-Termin-Kopf</strong>. Tipp: ein kleiner Innenabstand lässt das Bild fast bis an den Rand laufen.
                  </p>
                  <div style={{ display: 'flex', alignItems: 'flex-end', gap: 14, flexWrap: 'wrap' }}>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: 3, fontSize: '0.72rem', color: 'var(--dex-gray-600)' }}>
                      Breite (px)
                      <input
                        type="number" min={80} max={600} step={10}
                        value={imageWidth ?? 180}
                        onChange={e => onImageWidthChange(clampInt(e.target.value, 600, 180))}
                        style={{ width: 86, height: 30, fontSize: '0.82rem', borderRadius: 4, border: '1px solid var(--dex-gray-300)', padding: '0 8px' }}
                      />
                    </label>
                    {onImagePaddingHChange && (
                      <label style={{ display: 'flex', flexDirection: 'column', gap: 3, fontSize: '0.72rem', color: 'var(--dex-gray-600)' }}>
                        Abstand seitlich (px)
                        <input
                          type="number" min={0} max={80} step={2}
                          value={imagePaddingH ?? 30}
                          onChange={e => onImagePaddingHChange(clampInt(e.target.value, 80, 30))}
                          style={{ width: 86, height: 30, fontSize: '0.82rem', borderRadius: 4, border: '1px solid var(--dex-gray-300)', padding: '0 8px' }}
                        />
                      </label>
                    )}
                    {onImagePaddingVChange && (
                      <label style={{ display: 'flex', flexDirection: 'column', gap: 3, fontSize: '0.72rem', color: 'var(--dex-gray-600)' }}>
                        Abstand oben/unten (px)
                        <input
                          type="number" min={0} max={80} step={2}
                          value={imagePaddingV ?? 30}
                          onChange={e => onImagePaddingVChange(clampInt(e.target.value, 80, 30))}
                          style={{ width: 86, height: 30, fontSize: '0.82rem', borderRadius: 4, border: '1px solid var(--dex-gray-300)', padding: '0 8px' }}
                        />
                      </label>
                    )}
                    <button
                      type="button"
                      onMouseDown={e => e.preventDefault()}
                      onClick={() => { onImageWidthChange(180); if (onImagePaddingHChange) onImagePaddingHChange(30); if (onImagePaddingVChange) onImagePaddingVChange(30); }}
                      style={{ height: 30, padding: '0 12px', fontSize: '0.72rem', fontWeight: 600, cursor: 'pointer', background: 'transparent', color: 'var(--dex-gray-600)', border: '1px solid var(--dex-gray-300)', borderRadius: 6 }}
                      title="Auf die Standard-Werte zurücksetzen (Breite 180, Abstand 30)"
                    >
                      Standard
                    </button>
                  </div>
                </div>
              )}

              {insertableVars.length > 0 && (
                <div>
                  <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>
                    Variable einfügen (klicken — wird an aktueller Cursor-Position eingefügt):
                  </label>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                    {insertableVars.map(v => (
                      <button
                        key={v.key}
                        type="button"
                        onMouseDown={e => { e.preventDefault(); }}
                        onClick={() => insertVariable(v.key)}
                        style={{
                          fontSize: '0.7rem', padding: '3px 10px', borderRadius: 12,
                          border: '1px solid var(--dex-green)', background: 'rgba(134,188,37,0.08)',
                          color: 'var(--dex-green-dark)', cursor: 'pointer', fontFamily: 'monospace',
                        }}
                        title={`${v.key} an Cursor einfügen`}
                      >
                        {v.label}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              <div>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 4 }}>
                  <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>Body</label>
                  {defaultBodyHtml && (
                    <button
                      type="button"
                      onMouseDown={e => e.preventDefault()}
                      onClick={() => {
                        const cur = editorRef.current?.innerHTML || '';
                        const isEmpty = cur.replace(/<[^>]*>/g, '').replace(/\s/g, '').trim() === '';
                        const apply = (): void => {
                          if (editorRef.current) { editorRef.current.innerHTML = defaultBodyHtml; fireChange(); }
                        };
                        if (isEmpty) { apply(); return; }
                        confirmDialog('Den aktuellen Body-Text durch den Standardtext ersetzen?', { confirmLabel: 'Ersetzen' })
                          .then(ok => { if (ok) apply(); })
                          .catch(() => { /* */ });
                      }}
                      style={{
                        fontSize: '0.72rem', fontWeight: 600, cursor: 'pointer',
                        background: 'transparent', color: 'var(--dex-gray-600)',
                        border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '3px 9px',
                        display: 'inline-flex', alignItems: 'center', gap: 4,
                      }}
                      title="Setzt den Body auf die Standard-Vorlage zurück (mit Variablen)"
                    >
                      <Icon iconName="Refresh" style={{ fontSize: 11 }} /> Standardtext laden
                    </button>
                  )}
                </div>
                {/* Toolbar */}
                <div style={{
                  display: 'flex', flexWrap: 'wrap', gap: 4, padding: 6,
                  border: '1px solid var(--dex-gray-300)', borderBottom: 'none',
                  borderTopLeftRadius: 6, borderTopRightRadius: 6, background: 'var(--dex-gray-50)',
                  alignItems: 'center',
                }}>
                  <button type="button" style={tbBtn} title="Fett" onMouseDown={e => e.preventDefault()} onClick={() => exec('bold')}><strong>B</strong></button>
                  <button type="button" style={tbBtn} title="Kursiv" onMouseDown={e => e.preventDefault()} onClick={() => exec('italic')}><em>I</em></button>
                  <button type="button" style={tbBtn} title="Unterstrichen" onMouseDown={e => e.preventDefault()} onClick={() => exec('underline')}><span style={{ textDecoration: 'underline' }}>U</span></button>
                  <span style={{ width: 1, height: 22, background: 'var(--dex-gray-300)', margin: '0 4px' }} />
                  {/* v18.20: kontrolliertes Größen-Dropdown — zeigt (wie in Word)
                      die Größe des aktuell markierten Texts und setzt sie beim
                      Ändern. Live-Vorschau aktualisiert sich über fireChange(). */}
                  <select
                    title="Schriftgröße der Auswahl"
                    value={currentFontPx != null ? String(currentFontPx) : ''}
                    onMouseDown={() => saveSelection()}
                    onChange={e => { if (e.target.value) setFontSize(parseInt(e.target.value, 10)); }}
                    style={{ height: 28, fontSize: '0.78rem', borderRadius: 4, border: '1px solid var(--dex-gray-300)' }}
                  >
                    <option value="" disabled>Größe</option>
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
                    title="Zeilenabstand"
                    onChange={e => { if (e.target.value) { setLineHeight(e.target.value); e.target.value = ''; } }}
                    defaultValue=""
                    style={{ height: 28, fontSize: '0.78rem', borderRadius: 4, border: '1px solid var(--dex-gray-300)' }}
                  >
                    <option value="" disabled>Zeilenabstand</option>
                    <option value="1.0">Eng (1.0)</option>
                    <option value="1.15">Kompakt (1.15)</option>
                    <option value="1.3">Normal (1.3)</option>
                    <option value="1.5">Locker (1.5)</option>
                    <option value="2.0">Weit (2.0)</option>
                  </select>
                  <span style={{ width: 1, height: 22, background: 'var(--dex-gray-300)', margin: '0 4px' }} />
                  <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>Farbe:</span>
                  {COLORS.map(c => (
                    <button
                      key={c}
                      type="button"
                      title={c}
                      onMouseDown={e => e.preventDefault()}
                      onClick={() => setColor(c)}
                      style={{
                        width: 22, height: 22, borderRadius: '50%',
                        background: c, border: c === '#000000' ? '1px solid #fff' : '1px solid var(--dex-gray-300)',
                        boxShadow: 'inset 0 0 0 1px rgba(0,0,0,0.1)',
                        cursor: 'pointer', padding: 0,
                      }}
                    />
                  ))}
                  {/* v18.22: freie Body-Textfarbe — nativer Picker + Hex-Eingabe.
                      Auswahl wird vor dem Fokuswechsel gesichert (onMouseDown). */}
                  <input
                    type="color"
                    value={isHex6(bodyHexDraft) ? bodyHexDraft : '#000000'}
                    title="Freie Textfarbe wählen"
                    onMouseDown={() => saveSelection()}
                    onChange={e => { setBodyHexDraft(e.target.value.toLowerCase()); setColor(e.target.value.toLowerCase()); }}
                    style={{ width: 24, height: 24, padding: 0, border: '1px solid var(--dex-gray-300)', borderRadius: 4, cursor: 'pointer', background: '#fff' }}
                  />
                  <input
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
                    title="Hex-Code eingeben + Enter — färbt die Auswahl"
                    style={{ width: 76, height: 24, fontSize: '0.72rem', fontFamily: 'monospace', borderRadius: 4, border: '1px solid var(--dex-gray-300)', padding: '0 6px' }}
                  />
                  <span style={{ width: 1, height: 22, background: 'var(--dex-gray-300)', margin: '0 4px' }} />
                  <button type="button" style={tbBtn} title="Liste" onMouseDown={e => e.preventDefault()} onClick={() => exec('insertUnorderedList')}>•</button>
                  <button type="button" style={tbBtn} title="Nummerierte Liste" onMouseDown={e => e.preventDefault()} onClick={() => exec('insertOrderedList')}>1.</button>
                  <span style={{ width: 1, height: 22, background: 'var(--dex-gray-300)', margin: '0 4px' }} />
                  {/* v18.39: Link einfügen/bearbeiten — Text markieren + klicken,
                      oder ohne Auswahl klicken für einen neuen Link. Bestehenden
                      Link: Cursor hineinsetzen → URL ändern. Leere URL entfernt
                      den Link. So muss kein Link mehr per Copy-Paste eingefügt
                      werden (das brach den Zeilenabstand). */}
                  <button type="button" style={tbBtn} title="Link einfügen / bearbeiten" onMouseDown={e => e.preventDefault()} onClick={insertLink}><Icon iconName="Link" style={{ fontSize: 13 }} /></button>
                  <button type="button" style={tbBtn} title="Formatierung entfernen" onMouseDown={e => e.preventDefault()} onClick={() => exec('removeFormat')}>⌫</button>
                </div>
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
                    minHeight: 280, padding: '10px 12px',
                    border: '1px solid var(--dex-gray-300)',
                    borderBottomLeftRadius: 6, borderBottomRightRadius: 6,
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
              </div>
            </div>
          </div>

          {/* === PREVIEW === */}
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--dex-gray-50)', minHeight: 0 }}>
            <div style={{ padding: '10px 16px', fontSize: '0.75rem', color: 'var(--dex-gray-500)', borderBottom: '1px solid var(--dex-gray-200)', background: '#fff' }}>
              Live-Vorschau {previewMode === 'outlook' ? '(Outlook-Termin)' : previewMode === 'plain' ? '(Anmelde-Seite)' : '(Deloitte-Mail)'} — Variablen werden mit Beispielwerten ersetzt
            </div>
            <iframe
              title="Vorschau"
              srcDoc={renderPreviewHtml()}
              sandbox=""
              style={{ flex: 1, border: 'none', width: '100%', minHeight: 360, background: '#f5f5f5' }}
            />
          </div>
        </div>

        <div style={{ padding: '12px 20px', display: 'flex', justifyContent: 'flex-end', gap: 8, borderTop: '1px solid var(--dex-gray-200)' }}>
          <button type="button" className="btn btn-secondary" onClick={onClose}>{extraAction ? 'Abbrechen' : 'Fertig'}</button>
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
    </div>
  );
};
