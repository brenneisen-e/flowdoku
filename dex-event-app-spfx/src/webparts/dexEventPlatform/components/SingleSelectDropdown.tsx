/**
 * v32.34: Einzelauswahl im Stil von MultiSelectDropdown — mit UMBRECHENDER
 * Anzeige.
 *
 * Ein natives `<select>` zeigt die gewählte Antwort immer einzeilig und
 * schneidet sie ab (Nutzer-Befund 29.09.2026: „Dinner only on 19 November
 * AND overnight st…"). Das lässt sich per CSS nicht ändern — der Browser
 * zeichnet das Feld selbst. Deshalb ein Knopf, der die ganze Antwort zeigt,
 * und eine Liste, die als `position: fixed` an ihm hängt (dasselbe Muster wie
 * MultiSelectDropdown v22.4, damit `overflow: hidden` der Karte sie nicht
 * abschneidet). Gruppen (früher `<optgroup>`) werden als Zwischenüberschrift
 * gezeigt.
 */
import * as React from 'react';

export interface SingleSelectOption {
  value: string;
  label: string;
  /** Optionale Gruppe — Optionen mit gleicher Gruppe stehen unter einer Überschrift. */
  group?: string;
}

export interface SingleSelectDropdownProps {
  options: SingleSelectOption[];
  value: string;
  onChange: (next: string) => void;
  placeholder?: string;
  error?: boolean;
  disabled?: boolean;
  /** Zusätzliche Stile für den Knopf (z. B. der Fehlerrahmen des Aufrufers). */
  style?: React.CSSProperties;
  ariaLabel?: string;
}

const CHEVRON = "url(\"data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3e%3cpath fill='%23666' d='M6 8L1 3h10z'/%3e%3c/svg%3e\")";

export const SingleSelectDropdown: React.FC<SingleSelectDropdownProps> = ({
  options, value, onChange, placeholder = 'Please select', error = false, disabled = false, style, ariaLabel,
}) => {
  const [open, setOpen] = React.useState(false);
  const [hover, setHover] = React.useState(-1);
  const rootRef = React.useRef<HTMLDivElement | null>(null);
  const [menuPos, setMenuPos] = React.useState<{ left: number; width: number; openUp: boolean; anchorTop: number; anchorBottom: number; maxHeight: number } | null>(null);
  const recalcMenu = React.useCallback((): void => {
    const el = rootRef.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const spaceBelow = window.innerHeight - r.bottom - 8;
    const spaceAbove = r.top - 8;
    const openUp = spaceBelow < 200 && spaceAbove > spaceBelow;
    const maxHeight = Math.max(140, Math.min(320, openUp ? spaceAbove : spaceBelow));
    setMenuPos({ left: r.left, width: r.width, openUp, anchorTop: r.bottom + 4, anchorBottom: window.innerHeight - r.top + 4, maxHeight });
  }, []);

  React.useEffect(() => {
    if (!open) return undefined;
    recalcMenu();
    const onDocClick = (e: MouseEvent): void => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener('mousedown', onDocClick);
    window.addEventListener('scroll', recalcMenu, true);
    window.addEventListener('resize', recalcMenu);
    return () => {
      document.removeEventListener('mousedown', onDocClick);
      window.removeEventListener('scroll', recalcMenu, true);
      window.removeEventListener('resize', recalcMenu);
    };
  }, [open, recalcMenu]);

  const current = options.find(o => o.value === value);
  const isEmpty = !current;
  const waehle = (v: string): void => { onChange(v); setOpen(false); };

  // Tastatur wie beim nativen Feld: Pfeile wandern, Enter wählt, Escape schließt.
  const onKeyDown = (e: React.KeyboardEvent): void => {
    if (disabled) return;
    if (e.key === 'Escape') { setOpen(false); return; }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      if (!open) { setOpen(true); setHover(Math.max(0, options.findIndex(o => o.value === value))); return; }
      setHover(h => {
        const n = options.length;
        if (n === 0) return -1;
        return e.key === 'ArrowDown' ? Math.min(n - 1, h + 1) : Math.max(0, h - 1);
      });
      return;
    }
    if ((e.key === 'Enter' || e.key === ' ') && open && hover >= 0 && options[hover]) {
      e.preventDefault();
      waehle(options[hover].value);
    }
  };

  let letzteGruppe: string | undefined;
  return (
    <div ref={rootRef} style={{ position: 'relative', width: '100%' }}>
      <button
        type="button"
        disabled={disabled}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-label={ariaLabel}
        onClick={() => !disabled && setOpen(o => !o)}
        onKeyDown={onKeyDown}
        style={{
          width: '100%',
          textAlign: 'left',
          padding: '12px 36px 12px 16px',
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
          // Der Kern der Sache: die Antwort bricht um, statt abgeschnitten zu werden.
          whiteSpace: 'normal',
          overflowWrap: 'anywhere',
          fontFamily: 'inherit',
          backgroundImage: CHEVRON,
          backgroundRepeat: 'no-repeat',
          backgroundPosition: 'right 12px center',
          backgroundSize: '12px 12px',
          ...style,
        }}
      >
        {current ? (current.group ? `${current.group} ${current.label}` : current.label) : placeholder}
      </button>
      {open && menuPos && (
        <div
          role="listbox"
          style={{
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
          {options.length === 0 && <div style={{ padding: 12, color: 'var(--dex-gray-400)', fontSize: '0.9rem' }}>—</div>}
          {options.map((o, i) => {
            const kopf = o.group && o.group !== letzteGruppe
              ? <div key={`g-${o.group}-${i}`} style={{ padding: '8px 12px 4px', fontSize: '0.75rem', fontWeight: 700, color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: '0.04em' }}>{o.group}</div>
              : null;
            letzteGruppe = o.group;
            const selected = o.value === value;
            return (
              <React.Fragment key={`${o.group || ''}-${o.value}-${i}`}>
                {kopf}
                <div
                  role="option"
                  aria-selected={selected}
                  onMouseDown={e => e.preventDefault()}
                  onClick={() => waehle(o.value)}
                  onMouseEnter={() => setHover(i)}
                  style={{
                    padding: o.group ? '8px 12px 8px 22px' : '8px 12px',
                    cursor: 'pointer',
                    fontSize: '0.9rem',
                    lineHeight: 1.4,
                    color: 'var(--dex-gray-800)',
                    background: selected ? 'rgba(134,188,37,0.14)' : (hover === i ? 'var(--dex-gray-50, #fafafa)' : 'transparent'),
                    fontWeight: selected ? 600 : 400,
                  }}
                >
                  {o.label}
                </div>
              </React.Fragment>
            );
          })}
        </div>
      )}
    </div>
  );
};

export default SingleSelectDropdown;
