/**
 * v28.94: Aus `EventCreationPage` herausgeloest. Multi-Select für den
 * Standortfilter (loeste in v8.0 die Pillen-Buttons ab).
 */
import * as React from 'react';

// v8.0: Multi-Select-Dropdown für den Standortfilter (löst die alten
// Pillen-Buttons ab — kompakter und mit Suche bei vielen Optionen).
export function LocationMultiSelect({
  options, selected, onChange, isDe,
}: {
  options: string[];
  selected: string[];
  onChange: (next: string[]) => void;
  isDe: boolean;
}): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const [query, setQuery] = React.useState('');
  const ref = React.useRef<HTMLDivElement>(null);

  // Click-Outside zum Schliessen
  React.useEffect(() => {
    if (!open) return undefined;
    const handler = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false);
        setQuery('');
      }
    };
    window.addEventListener('mousedown', handler);
    return () => window.removeEventListener('mousedown', handler);
  }, [open]);

  const toggle = (loc: string): void => {
    if (selected.indexOf(loc) >= 0) onChange(selected.filter(l => l !== loc));
    else onChange([...selected, loc]);
  };

  const filtered = query.trim()
    ? options.filter(o => o.toLowerCase().indexOf(query.trim().toLowerCase()) >= 0)
    : options;

  return (
    <div ref={ref} style={{ position: 'relative', maxWidth: 520 }}>
      <button
        type="button"
        onClick={() => setOpen(v => !v)}
        style={{
          width: '100%', minHeight: 42, padding: '6px 12px',
          background: '#fff',
          border: `1.5px solid ${open ? 'var(--dex-green)' : 'var(--dex-gray-300)'}`,
          borderRadius: 8, cursor: 'pointer', textAlign: 'left',
          display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 6,
          fontSize: '0.88rem', color: 'var(--dex-gray-800)',
          transition: 'border-color 0.15s ease',
        }}
      >
        {selected.length === 0 ? (
          <span style={{ color: 'var(--dex-gray-400)' }}>
            {isDe ? 'Standorte auswählen…' : 'Select locations…'}
          </span>
        ) : (
          selected.map(loc => (
            <span
              key={loc}
              style={{
                display: 'inline-flex', alignItems: 'center', gap: 4,
                padding: '3px 4px 3px 10px',
                background: 'var(--dex-green)', color: '#fff',
                borderRadius: 999, fontSize: '0.78rem',
              }}
            >
              {loc}
              <span
                role="button"
                aria-label={`${loc} entfernen`}
                onClick={e => { e.stopPropagation(); toggle(loc); }}
                style={{
                  width: 18, height: 18, borderRadius: '50%',
                  background: 'rgba(255,255,255,0.25)',
                  display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                  fontSize: '0.85rem', lineHeight: 1, cursor: 'pointer',
                }}
              >×</span>
            </span>
          ))
        )}
        <span style={{ marginLeft: 'auto', color: 'var(--dex-gray-500)', fontSize: '0.7rem' }}>
          {open ? '▲' : '▼'}
        </span>
      </button>
      {open && (
        <div style={{
          position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0,
          background: '#fff', border: '1px solid var(--dex-gray-200)',
          borderRadius: 8, boxShadow: '0 8px 24px rgba(0,0,0,0.12)',
          maxHeight: 320, overflowY: 'auto', zIndex: 50,
        }}>
          {options.length > 6 && (
            <div style={{ padding: 8, borderBottom: '1px solid var(--dex-gray-100)', position: 'sticky', top: 0, background: '#fff' }}>
              <input
                autoFocus
                value={query}
                onChange={e => setQuery(e.target.value)}
                placeholder={isDe ? 'Suchen…' : 'Search…'}
                style={{
                  width: '100%', padding: '6px 10px',
                  border: '1px solid var(--dex-gray-200)',
                  borderRadius: 6, fontSize: '0.85rem',
                }}
              />
            </div>
          )}
          {filtered.length === 0 ? (
            <div style={{ padding: 12, fontSize: '0.82rem', color: 'var(--dex-gray-400)' }}>
              {isDe ? 'Keine Treffer.' : 'No matches.'}
            </div>
          ) : (
            filtered.map(loc => {
              const isChecked = selected.indexOf(loc) >= 0;
              return (
                <label
                  key={loc}
                  style={{
                    display: 'flex', alignItems: 'center', gap: 10,
                    padding: '8px 12px', cursor: 'pointer', fontSize: '0.88rem',
                    background: isChecked ? 'rgba(134,188,37,0.08)' : 'transparent',
                  }}
                  onMouseEnter={e => { if (!isChecked) (e.currentTarget as HTMLLabelElement).style.background = 'var(--dex-gray-50)'; }}
                  onMouseLeave={e => { if (!isChecked) (e.currentTarget as HTMLLabelElement).style.background = 'transparent'; }}
                >
                  <input
                    type="checkbox"
                    checked={isChecked}
                    onChange={() => toggle(loc)}
                    style={{ width: 16, height: 16, accentColor: 'var(--dex-green)', cursor: 'pointer' }}
                  />
                  <span style={{ color: 'var(--dex-gray-800)' }}>{loc}</span>
                </label>
              );
            })
          )}
          {selected.length > 0 && (
            <div style={{
              padding: 8, borderTop: '1px solid var(--dex-gray-100)',
              position: 'sticky', bottom: 0, background: '#fff',
              display: 'flex', justifyContent: 'flex-end',
            }}>
              <button
                type="button"
                onClick={() => onChange([])}
                style={{
                  background: 'none', border: 'none', cursor: 'pointer',
                  fontSize: '0.78rem', color: 'var(--dex-gray-500)',
                  textDecoration: 'underline', padding: '4px 8px',
                }}
              >
                {isDe ? 'Auswahl leeren' : 'Clear selection'}
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
