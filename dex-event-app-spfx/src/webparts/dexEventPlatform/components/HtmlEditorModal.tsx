/**
 * Wiederverwendbares Modal mit HTML-Editor + Live-Vorschau.
 *
 * Wird genutzt fuer:
 *   - Outlook-Termin-Body (Calendar Description)
 *   - E-Mail-Templates (Anmeldung, Warteliste, Abmeldung, Nachruecken)
 *
 * Layout (split): links der Editor (pnp RichText + Variable-Buttons), rechts
 * die gerenderte Vorschau im echten Wrapper (Outlook-Calendar bzw. Deloitte-Mail).
 *
 * Die Vorschau wird in einem sandboxed iframe gerendert, damit das Wrapper-CSS
 * nicht mit dem App-CSS kollidiert und der User sieht wie es im echten Mail-
 * client aussieht.
 */
import * as React from 'react';
import { RichText } from '@pnp/spfx-controls-react/lib/controls/richText';
import { X } from './Icons';
import { wrapTemplate, replacePlaceholders, replacePlaceholdersPlain } from '../services/EmailTemplates';

type PreviewMode = 'outlook' | 'email';

export interface HtmlEditorModalProps {
  open: boolean;
  onClose: () => void;
  /** Modal-Titel oben */
  title: string;
  /** Aktueller Body-HTML-Inhalt (oder Plain-Text fuer Outlook) */
  value: string;
  onChange: (newValue: string) => void;
  /** Vorschau-Modus: 'outlook' = Outlook-Termin-Wrapper, 'email' = Deloitte-Mail-Wrapper */
  previewMode: PreviewMode;
  /** Optional: bei 'email'-Mode der Subject + Heading + HeadingColor fuer Preview */
  emailSubject?: string;
  onEmailSubjectChange?: (s: string) => void;
  emailHeading?: string;
  onEmailHeadingChange?: (s: string) => void;
  emailHeadingColor?: string;
  /** Variablen die in der Vorschau ersetzt werden ({{Name}}, {{EventTitle}}, etc.) */
  previewVars?: Record<string, string>;
  /** Optional: Variablen-Schnellauswahl-Buttons (key + label) */
  insertableVars?: Array<{ key: string; label: string }>;
  /** Logo-Base64 fuer Preview-Header (DEX-Logo) — falls leer, zeigt Modal nur den Wrapper ohne Logo */
  logoBase64?: string;
  /** ORB-/Event-Bild fuer Preview (Outlook: Event-Bild, Mail: ORB) */
  imageBase64?: string;
}

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

  if (!open) return null;

  // Wrapper rendern: Outlook-Calendar-Stil oder Deloitte-Mail-Stil
  const renderPreviewHtml = (): string => {
    // Variablen ersetzen — Mail wird HTML-escaped, Outlook-Body bleibt as-is (schon HTML aus Editor)
    const bodyWithVars = replacePlaceholders(value || '', previewVars);
    if (previewMode === 'email') {
      const heading = replacePlaceholdersPlain(emailHeading || '', previewVars);
      const subheading = `Event ${previewVars.EventTitle || ''}`;
      const wrapped = wrapTemplate(emailHeadingColor, heading, subheading, bodyWithVars);
      return wrapped
        .replace(/\{\{LOGO_URL\}\}/g, logoBase64 || '')
        .replace(/\{\{ORB_URL\}\}/g, imageBase64 || '');
    }
    // Outlook-Calendar-Mode: schlankerer Wrapper (kein Footer-Legal-Block, da Outlook
    // den Body in den Termin-Detail-Bereich rendert)
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
      body { margin:0; padding:0; font-family: Aptos, Arial, Helvetica, sans-serif; color:#333; background:#f5f5f5; }
      .container { max-width: 600px; margin: 0 auto; background:#fff; }
      .hdr { background:#000; padding:16px 24px; border-bottom:2px solid #86bc25; }
      .hdr img { max-width: 160px; height: auto; }
      .hero { text-align:center; padding:24px; background:#fff; }
      .hero img { max-width:140px; height:auto; }
      .ttl { padding:24px 24px 0; }
      .ttl h1 { margin:0; font-size:22px; color:#86bc25; font-weight:400; }
      .ttl h2 { margin:0 0 16px; font-size:16px; color:#000; font-weight:700; }
      .body { padding: 0 24px 24px; font-size: 14px; line-height: 1.6; }
      .body p { margin: 0 0 12px; }
      .ftr { padding:12px 24px; font-size:11px; color:#999; border-top:1px solid #eee; }
    </style></head><body>
      <div class="container">
        <div class="hdr">${logoBase64 ? `<img src="${logoBase64}" alt="Deloitte" />` : '<span style="color:#fff;font-weight:700;">Deloitte.</span>'}</div>
        ${imageBase64 ? `<div class="hero"><img src="${imageBase64}" alt="Event" /></div>` : ''}
        <div class="ttl">
          <h1>${escapeHtml(previewVars.EventTitle || 'Event Title')}</h1>
          <h2>${escapeHtml(previewVars.EventDate || '')}</h2>
        </div>
        <div class="body">${bodyWithVars || '<p style="color:#999;font-style:italic;">Hier erscheint der Body — beginne im Editor links zu tippen.</p>'}</div>
        <div class="ftr">Outlook-Termin-Vorschau · DEX Event Experience Platform</div>
      </div>
    </body></html>`;
  };

  const insertVariable = (key: string): void => {
    // Variable an Cursor-Position einfuegen ist mit RichText von pnp aufwendig, daher
    // einfacher: in Zwischenablage kopieren und User pastet selbst (Standard im Projekt)
    try { navigator.clipboard.writeText(key); } catch { /* no-op */ }
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
        {/* Header */}
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

        {/* Body: Editor links, Preview rechts */}
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
                    Variable in Zwischenablage kopieren (dann an gewünschter Stelle einfügen):
                  </label>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                    {insertableVars.map(v => (
                      <button
                        key={v.key}
                        type="button"
                        onClick={() => insertVariable(v.key)}
                        style={{
                          fontSize: '0.7rem', padding: '3px 10px', borderRadius: 12,
                          border: '1px solid var(--dex-green)', background: 'rgba(134,188,37,0.08)',
                          color: 'var(--dex-green-dark)', cursor: 'pointer', fontFamily: 'monospace',
                        }}
                        title={`${v.key} kopieren`}
                      >
                        {v.label}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              <div>
                <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 4 }}>Body (HTML)</label>
                <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, minHeight: 280, padding: '0 4px' }}>
                  <RichText
                    value={value || ''}
                    onChange={(text: string) => { onChange(text); return text; }}
                  />
                </div>
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

        {/* Footer */}
        <div style={{ padding: '12px 20px', display: 'flex', justifyContent: 'flex-end', gap: 8, borderTop: '1px solid var(--dex-gray-200)' }}>
          <button type="button" className="btn btn-primary" onClick={onClose}>Fertig</button>
        </div>
      </div>
    </div>
  );
};

function escapeHtml(s: string): string {
  return (s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
