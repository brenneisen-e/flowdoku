// MultiSelect-Dropdown im Stil der normalen Single-Select-Dropdowns
// (Anzeige als geschlossener „Please select"-Button, Klick öffnet ein
// Panel mit Checkboxen pro Option, ausgewählte Werte als Komma-Liste
// im Header). Ersetzt seit v11.89 die früheren Checkbox-Listen, damit
// die UI zwischen Single- und Multi-Select konsistent aussieht.
import * as React from 'react';

export interface MultiSelectDropdownProps {
  options: string[];
  value: string[];
  onChange: (next: string[]) => void;
  placeholder?: string;
  error?: boolean;
  disabled?: boolean;
}

export const MultiSelectDropdown: React.FC<MultiSelectDropdownProps> = ({
  options,
  value,
  onChange,
  placeholder = 'Please select',
  error = false,
  disabled = false,
}) => {
  const [open, setOpen] = React.useState(false);
  const rootRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    if (!open) return undefined;
    const onDocClick = (e: MouseEvent): void => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', onDocClick);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDocClick);
      document.removeEventListener('keydown', onKey);
    };
  }, [open]);

  const toggle = (opt: string): void => {
    if (value.indexOf(opt) >= 0) {
      onChange(value.filter(v => v !== opt));
    } else {
      onChange([...value, opt]);
    }
  };

  const label = value.length === 0 ? placeholder : value.join(', ');
  const isEmpty = value.length === 0;

  return (
    <div ref={rootRef} style={{ position: 'relative', width: '100%' }}>
      <button
        type="button"
        disabled={disabled}
        onClick={() => !disabled && setOpen(o => !o)}
        style={{
          // v11.97: exakt gleiche Geometrie wie .form-select (siehe SCSS):
          // padding 12/16, border 1.5px, border-radius 12, font 0.95rem,
          // min-height 48px, line-height 1.45.
          width: '100%',
          textAlign: 'left',
          padding: '12px 36px 12px 16px',
          border: error ? '1.5px solid var(--dex-red)' : '1.5px solid var(--dex-gray-200)',
          borderRadius: 12,
          background: disabled ? 'var(--dex-gray-50)' : 'var(--dex-white, #fff)',
          color: isEmpty ? 'var(--dex-gray-400)' : 'var(--dex-gray-800)',
          fontSize: '0.95rem',
          lineHeight: 1.45,
          minHeight: 48,
          boxSizing: 'border-box',
          cursor: disabled ? 'not-allowed' : 'pointer',
          outline: open ? '2px solid var(--dex-green, #86bc25)' : 'none',
          outlineOffset: -1,
          position: 'relative',
          overflow: 'hidden',
          whiteSpace: 'nowrap',
          textOverflow: 'ellipsis',
          fontFamily: 'inherit',
          // v11.98: gleicher Chevron-Look wie .form-select (inline SVG
          // als background-image rechts mittig). Vorher Unicode-▾ —
          // visuell anders als der Single-Select-Pfeil.
          backgroundImage: "url(\"data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3e%3cpath fill='%23666' d='M6 8L1 3h10z'/%3e%3c/svg%3e\")",
          backgroundRepeat: 'no-repeat',
          backgroundPosition: 'right 12px center',
          backgroundSize: '12px 12px',
        }}
      >
        {label}
      </button>
      {open && (
        <div
          style={{
            position: 'absolute',
            top: 'calc(100% + 4px)',
            left: 0,
            right: 0,
            zIndex: 1000,
            background: '#fff',
            border: '1px solid var(--dex-gray-200)',
            borderRadius: 8,
            boxShadow: '0 4px 12px rgba(0,0,0,0.08)',
            maxHeight: 280,
            overflowY: 'auto',
          }}
        >
          {options.length === 0 ? (
            <div style={{ padding: 12, color: 'var(--dex-gray-400)', fontSize: '0.9rem' }}>
              —
            </div>
          ) : (
            options.map(opt => {
              const selected = value.indexOf(opt) >= 0;
              return (
                <label
                  key={opt}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 10,
                    padding: '8px 12px',
                    cursor: 'pointer',
                    fontSize: '0.9rem',
                    color: 'var(--dex-gray-700)',
                    background: selected ? 'rgba(134,188,37,0.10)' : 'transparent',
                  }}
                  onMouseEnter={e => { if (!selected) (e.currentTarget as HTMLLabelElement).style.background = 'var(--dex-gray-50)'; }}
                  onMouseLeave={e => { if (!selected) (e.currentTarget as HTMLLabelElement).style.background = 'transparent'; }}
                >
                  <input
                    type="checkbox"
                    checked={selected}
                    onChange={() => toggle(opt)}
                    style={{ accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
                  />
                  <span>{opt}</span>
                </label>
              );
            })
          )}
        </div>
      )}
    </div>
  );
};
