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
  /** v17.20: Optionale positional gemappte Anzeige-Labels (z.B. für den
   *  bilingualen Modus, in dem das englische Label angezeigt wird, der
   *  gespeicherte Wert aber weiterhin das deutsche Original ist). Wenn
   *  gesetzt, MUSS `optionLabels.length === options.length`. */
  optionLabels?: string[];
}

export const MultiSelectDropdown: React.FC<MultiSelectDropdownProps> = ({
  options,
  value,
  onChange,
  placeholder = 'Please select',
  error = false,
  disabled = false,
  optionLabels,
}) => {
  const [open, setOpen] = React.useState(false);
  const rootRef = React.useRef<HTMLDivElement | null>(null);
  // v22.4: Das Optionen-Panel als position:fixed an die Button-Bounding-Box
  // ankern, damit es NICHT vom `overflow:hidden` der umgebenden Card
  // abgeschnitten wird (Bug: die letzte(n) Option(en) eines Dropdowns waren
  // unsichtbar). Gleiches Muster wie der People-Picker (UserFieldPicker,
  // v18.5). Bei wenig Platz unten klappt es nach oben auf, sonst nach unten;
  // die Höhe wird auf den verfügbaren Platz begrenzt (intern scrollbar).
  const [menuPos, setMenuPos] = React.useState<{ left: number; width: number; openUp: boolean; anchorTop: number; anchorBottom: number; maxHeight: number } | null>(null);
  const recalcMenu = React.useCallback((): void => {
    const el = rootRef.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const spaceBelow = window.innerHeight - r.bottom - 8;
    const spaceAbove = r.top - 8;
    const openUp = spaceBelow < 200 && spaceAbove > spaceBelow;
    const maxHeight = Math.max(140, Math.min(300, openUp ? spaceAbove : spaceBelow));
    setMenuPos({ left: r.left, width: r.width, openUp, anchorTop: r.bottom + 4, anchorBottom: window.innerHeight - r.top + 4, maxHeight });
  }, []);

  React.useEffect(() => {
    if (!open) return undefined;
    recalcMenu();
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
    // Beim Scrollen/Resizen neu ausrichten (capture=true: auch innere Scroll-
    // Container der Anmeldeseite erfassen).
    window.addEventListener('scroll', recalcMenu, true);
    window.addEventListener('resize', recalcMenu);
    return () => {
      document.removeEventListener('mousedown', onDocClick);
      document.removeEventListener('keydown', onKey);
      window.removeEventListener('scroll', recalcMenu, true);
      window.removeEventListener('resize', recalcMenu);
    };
  }, [open, recalcMenu]);

  const toggle = (opt: string): void => {
    if (value.indexOf(opt) >= 0) {
      onChange(value.filter(v => v !== opt));
    } else {
      onChange([...value, opt]);
    }
  };

  // v17.20: wenn optionLabels gesetzt sind, im Header die Anzeige-Labels
  // statt der Werte zusammensetzen (DE-Wert -> EN-Anzeige).
  const labelFor = (v: string): string => {
    if (!optionLabels) return v;
    const idx = options.indexOf(v);
    if (idx >= 0 && optionLabels[idx] && optionLabels[idx].trim()) return optionLabels[idx];
    return v;
  };
  const label = value.length === 0 ? placeholder : value.map(labelFor).join(', ');
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
          // v19.0: Bei Auswahl grüne Hervorhebung (wie ausgefüllte Felder).
          border: error ? '1.5px solid var(--dex-red)' : (!isEmpty && !disabled ? '1.5px solid var(--dex-green, #86bc25)' : '1.5px solid var(--dex-gray-200)'),
          borderRadius: 12,
          background: disabled ? 'var(--dex-gray-50)' : (!isEmpty ? 'rgba(134,188,37,0.06)' : 'var(--dex-white, #fff)'),
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
      {open && menuPos && (
        <div
          style={{
            // v22.4: position:fixed an der Bounding-Box — escaped das
            // overflow:hidden der Card; flippt bei Platzmangel nach oben.
            position: 'fixed',
            left: menuPos.left,
            width: menuPos.width,
            ...(menuPos.openUp ? { bottom: menuPos.anchorBottom } : { top: menuPos.anchorTop }),
            zIndex: 3000,
            background: '#fff',
            border: '1px solid var(--dex-gray-200)',
            borderRadius: 8,
            boxShadow: '0 8px 24px rgba(0,0,0,0.12)',
            maxHeight: menuPos.maxHeight,
            overflowY: 'auto',
          }}
        >
          {options.length === 0 ? (
            <div style={{ padding: 12, color: 'var(--dex-gray-400)', fontSize: '0.9rem' }}>
              —
            </div>
          ) : (
            options.map((opt, optIdx) => {
              const selected = value.indexOf(opt) >= 0;
              const shown = (optionLabels && optionLabels[optIdx] && optionLabels[optIdx].trim()) ? optionLabels[optIdx] : opt;
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
                  <span>{shown}</span>
                </label>
              );
            })
          )}
        </div>
      )}
    </div>
  );
};
