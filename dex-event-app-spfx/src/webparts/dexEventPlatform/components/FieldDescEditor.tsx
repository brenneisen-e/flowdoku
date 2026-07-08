import * as React from 'react';

/**
 * v27.10 (Refactor): unverändert aus EventCreationPage.tsx extrahiert.
 *
 * v27.4: Kompakter Beschreibungs-Editor mit DAUERHAFT sichtbarer Mini-Leiste
 * (Fett + Link). Die Buttons formatieren die Markierung (kein Markdown-Tippen);
 * gespeichert wird ein kleines Markdown-Subset, das die Anmeldeseite rendert.
 */
export default function FieldDescEditor({ value, onChange, isDe }: { value: string; onChange: (v: string) => void; isDe: boolean }): React.ReactElement {
  const ref = React.useRef<HTMLTextAreaElement>(null);
  const applyWrap = (before: string, after: string, ph: string): void => {
    const ta = ref.current; if (!ta) return;
    const s = ta.selectionStart; const e = ta.selectionEnd;
    const sel = value.substring(s, e) || ph;
    const next = value.substring(0, s) + before + sel + after + value.substring(e);
    onChange(next);
    window.setTimeout(() => { try { ta.focus(); ta.setSelectionRange(s + before.length, s + before.length + sel.length); } catch { /* */ } }, 0);
  };
  const insertLink = (): void => {
    const ta = ref.current; if (!ta) return;
    const s = ta.selectionStart; const e = ta.selectionEnd;
    const url = (window.prompt(isDe ? 'Link-Adresse (https://…):' : 'Link URL (https://…):', 'https://') || '').trim();
    if (!url) return;
    const sel = value.substring(s, e) || (isDe ? 'Link-Text' : 'link text');
    const next = value.substring(0, s) + `[${sel}](${url})` + value.substring(e);
    onChange(next);
    window.setTimeout(() => { try { ta.focus(); } catch { /* */ } }, 0);
  };
  const btn: React.CSSProperties = { height: 26, minWidth: 30, padding: '0 8px', fontSize: '0.8rem', cursor: 'pointer', background: '#fff', border: '1px solid var(--dex-gray-300)', borderRadius: 4, color: 'var(--dex-gray-700)' };
  return (
    <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, background: '#fff' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '4px 6px', borderBottom: '1px solid var(--dex-gray-200)' }}>
        <button type="button" title={isDe ? 'Fett' : 'Bold'} onMouseDown={e => e.preventDefault()} onClick={() => applyWrap('**', '**', isDe ? 'fetter Text' : 'bold text')} style={{ ...btn, fontWeight: 800 }}>F</button>
        <button type="button" title={isDe ? 'Link einfügen' : 'Insert link'} onMouseDown={e => e.preventDefault()} onClick={insertLink} style={{ ...btn, display: 'inline-flex', alignItems: 'center', gap: 5 }}>🔗 {isDe ? 'Link' : 'Link'}</button>
        <span style={{ marginLeft: 'auto', fontSize: '0.66rem', color: 'var(--dex-gray-400)' }}>{isDe ? 'Text markieren, dann Button' : 'Select text, then button'}</span>
      </div>
      <textarea
        ref={ref}
        value={value}
        onChange={e => onChange(e.target.value)}
        placeholder={isDe ? 'Beschreibung (optional)' : 'Description (optional)'}
        rows={2}
        style={{ width: '100%', boxSizing: 'border-box', border: 'none', outline: 'none', resize: 'vertical', fontSize: '0.85rem', lineHeight: 1.45, padding: '8px 10px', fontFamily: 'inherit', background: 'transparent', color: 'var(--dex-gray-800)' }}
      />
    </div>
  );
}
