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
import { wrapTemplate, replacePlaceholders, replacePlaceholdersPlain, getCachedOrbBase64, getCachedLogoBase64 } from '../services/EmailTemplates';

type PreviewMode = 'outlook' | 'email';

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
  emailHeadingColor?: string;
  previewVars?: Record<string, string>;
  insertableVars?: Array<{ key: string; label: string }>;
  logoBase64?: string;
  imageBase64?: string;
}

const FONT_SIZES: Array<{ label: string; px: number }> = [
  { label: 'Klein', px: 12 },
  { label: 'Normal', px: 14 },
  { label: 'Groß', px: 18 },
  { label: 'Sehr groß', px: 24 },
];

const COLORS: string[] = [
  '#000000', '#555555', '#86bc25', '#0076a8', '#ed8b00', '#c9302c', '#6b21a8', '#0d6efd',
];

export const HtmlEditorModal: React.FC<HtmlEditorModalProps> = (props) => {
  const {
    open, onClose, title,
    value, onChange,
    previewMode,
    emailSubject, onEmailSubjectChange,
    emailHeading, onEmailHeadingChange, emailHeadingColor = '#86bc25',
    previewVars = {}, insertableVars = [],
    logoBase64 = '', imageBase64 = '',
  } = props;

  const editorRef = React.useRef<HTMLDivElement>(null);
  const savedSelectionRef = React.useRef<Range | null>(null);

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

  const setFontSize = (px: number): void => {
    restoreSelection();
    // Wrap selection in <span style="font-size:Xpx">. execCommand 'fontSize' nutzt nur 1-7
    // (HTML-Tags font size=...) und ist nicht praezise — eigener Wrapper ist robuster.
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    const range = sel.getRangeAt(0);
    if (range.collapsed) return;
    const span = document.createElement('span');
    span.style.fontSize = `${px}px`;
    try {
      span.appendChild(range.extractContents());
      range.insertNode(span);
      sel.removeAllRanges();
      const r = document.createRange();
      r.selectNodeContents(span);
      sel.addRange(r);
    } catch { /* ignore */ }
    fireChange();
  };

  const setColor = (hex: string): void => exec('foreColor', hex);

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

    if (previewMode === 'email') {
      const heading = replacePlaceholdersPlain(emailHeading || '', previewVars);
      const subheading = `Event ${previewVars.EventTitle || ''}`;
      const wrapped = wrapTemplate(emailHeadingColor, heading, subheading, bodyWithVars);
      return wrapped
        .replace(/\{\{LOGO_URL\}\}/g, logoBase64 || cachedLogo || '')
        .replace(/\{\{ORB_URL\}\}/g, imageBase64 || cachedOrb || '');
    }

    // Outlook-Termin-Vorschau: gleicher wrapTemplate() wie beim echten Versand,
    // damit der User in der Vorschau genau das sieht, was nachher im Termin steht
    // (inkl. Deloitte-Signatur + Legal-Disclaimer im Footer).
    const outlookHeading = previewVars.EventTitle || 'Event Title';
    const outlookSubheading = previewVars.EventDate || 'Event Details';
    const bodyForOutlook = bodyWithVars || '<p style="color:#999;font-style:italic;">Hier erscheint der Body — beginne im Editor links zu tippen.</p>';
    // Wenn der Body bereits ein kompletter wrapTemplate-Output ist (z.B. aus editEvent
    // ohne Strip), 1:1 anzeigen — sonst doppelt wickeln.
    const isAlreadyWrapped = /<!doctype|<html/i.test(bodyForOutlook);
    const wrapped = isAlreadyWrapped
      ? bodyForOutlook
      : wrapTemplate('#86bc25', outlookHeading, outlookSubheading, bodyForOutlook);
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
      onClick={onClose}
    >
      <div
        className="card"
        style={{
          width: '100%', maxWidth: 1280, maxHeight: '100%',
          display: 'flex', flexDirection: 'column',
          background: '#fff', borderRadius: 'var(--dex-radius)', overflow: 'hidden',
        }}
        onClick={e => e.stopPropagation()}
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
                  </div>
                </>
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
                <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>Body</label>
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
                  <select
                    title="Schriftgröße"
                    onMouseDown={() => saveSelection()}
                    onChange={e => { setFontSize(parseInt(e.target.value, 10)); e.target.value = ''; }}
                    defaultValue=""
                    style={{ height: 28, fontSize: '0.78rem', borderRadius: 4, border: '1px solid var(--dex-gray-300)' }}
                  >
                    <option value="" disabled>Größe</option>
                    {FONT_SIZES.map(f => <option key={f.px} value={f.px}>{f.label} ({f.px}px)</option>)}
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
                  <span style={{ width: 1, height: 22, background: 'var(--dex-gray-300)', margin: '0 4px' }} />
                  <button type="button" style={tbBtn} title="Liste" onMouseDown={e => e.preventDefault()} onClick={() => exec('insertUnorderedList')}>•</button>
                  <button type="button" style={tbBtn} title="Nummerierte Liste" onMouseDown={e => e.preventDefault()} onClick={() => exec('insertOrderedList')}>1.</button>
                  <button type="button" style={tbBtn} title="Formatierung entfernen" onMouseDown={e => e.preventDefault()} onClick={() => exec('removeFormat')}>⌫</button>
                </div>
                <div
                  ref={editorRef}
                  contentEditable
                  suppressContentEditableWarning
                  onInput={fireChange}
                  onBlur={() => { saveSelection(); fireChange(); }}
                  onMouseUp={saveSelection}
                  onKeyUp={saveSelection}
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
                    fontSize: '0.9rem', lineHeight: 1.5,
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
              Live-Vorschau {previewMode === 'outlook' ? '(Outlook-Termin)' : '(Deloitte-Mail)'} — Variablen werden mit Beispielwerten ersetzt
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
          <button type="button" className="btn btn-primary" onClick={onClose}>Fertig</button>
        </div>
      </div>
    </div>
  );
};
