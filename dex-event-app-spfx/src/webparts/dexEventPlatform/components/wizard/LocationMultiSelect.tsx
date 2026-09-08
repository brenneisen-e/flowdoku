/**
 * v28.94: Aus `EventCreationPage` herausgeloest. Multi-Select für den
 * Standortfilter (loeste in v8.0 die Pillen-Buttons ab).
 *
 * v31.2: Optik auf die `dex-ui-*`-Klassen umgestellt (Chips, Zeilen mit Hover,
 * Zähl-Pill, Textknöpfe) — vorher hatte jede Zeile ihren eigenen onMouseEnter-
 * Hover und keinen Fokus-Ring. Dazu „Alle wählen" (17 Standorte, „alle bis auf
 * zwei" waren 15 Klicks). Suche ab 7 Optionen, Klick daneben, Props: unverändert.
 */
import * as React from 'react';
import { cx } from '../dexUi';
import { ChevronDown, X } from '../Icons';

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

  // v31.2: „Alle wählen" ergänzt die bisherige Auswahl um die sichtbaren
  // Einträge. Bei aktiver Suche sind das nur die Treffer — der Knopf heißt
  // dann auch so, damit niemand denkt, er hätte alle 17 Standorte gesetzt.
  const hasQuery = query.trim().length > 0;
  const allVisibleChosen = filtered.every(o => selected.indexOf(o) >= 0);
  const selectVisible = (): void => {
    const next = selected.slice();
    filtered.forEach(o => { if (next.indexOf(o) < 0) next.push(o); });
    onChange(next);
  };

  return (
    <div ref={ref} style={{ position: 'relative', maxWidth: 520 }}>
      {/* v31.2: Auslöser wie ein Eingabefeld (dex-ui-input: Hover-Rand, Fokus-Ring),
          die gewählten Standorte als aktive Chips darin — Zustand und Bedienung auf einen Blick. */}
      <button
        type="button"
        className="dex-ui-input"
        aria-haspopup="listbox"
        aria-expanded={open}
        onClick={() => setOpen(v => !v)}
        style={{
          minHeight: 44, padding: '6px 12px', cursor: 'pointer', textAlign: 'left',
          display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 6,
          ...(open ? { borderColor: 'var(--dex-green, #86bc25)', boxShadow: '0 0 0 3px rgba(134,188,37,0.15)' } : {}),
        }}
      >
        {selected.length === 0 ? (
          <span style={{ color: 'var(--dex-gray-400, #a0a0a0)' }}>
            {isDe ? 'Standorte auswählen …' : 'Select locations …'}
          </span>
        ) : (
          selected.map(loc => (
            <span key={loc} className="dex-ui-chip is-active" style={{ paddingRight: 4 }}>
              {loc}
              <span
                role="button"
                aria-label={isDe ? `${loc} entfernen` : `Remove ${loc}`}
                title={isDe ? 'Entfernen' : 'Remove'}
                onClick={e => { e.stopPropagation(); toggle(loc); }}
                style={{
                  width: 18, height: 18, borderRadius: '50%', background: 'rgba(255,255,255,0.25)',
                  display: 'inline-flex', alignItems: 'center', justifyContent: 'center', lineHeight: 1, cursor: 'pointer',
                }}
              ><X size={11} /></span>
            </span>
          ))
        )}
        <span style={{ marginLeft: 'auto', display: 'inline-flex', alignItems: 'center', gap: 8 }}>
          {selected.length > 0 && (
            <span className="dex-ui-pill dex-ui-pill--green">
              {isDe ? `${selected.length} von ${options.length}` : `${selected.length} of ${options.length}`}
            </span>
          )}
          <span aria-hidden="true" style={{ display: 'inline-flex', color: 'var(--dex-gray-500, #808080)', transition: 'transform 0.18s ease', transform: open ? 'rotate(180deg)' : 'none' }}>
            <ChevronDown size={16} />
          </span>
        </span>
      </button>
      {open && (
        <div
          className="dex-ui-card dex-ui-fade-in"
          role="group"
          aria-label={isDe ? 'Standorte' : 'Locations'}
          style={{
            position: 'absolute', top: 'calc(100% + 6px)', left: 0, right: 0, padding: 6,
            boxShadow: '0 12px 32px rgba(0,0,0,0.12)', maxHeight: 320, overflowY: 'auto', zIndex: 50,
          }}
        >
          {options.length > 6 && (
            <div style={{ position: 'sticky', top: 0, zIndex: 1, margin: '-6px -6px 4px', padding: 8, background: '#fff', borderBottom: '1px solid var(--dex-gray-100, #f5f5f5)' }}>
              <input
                autoFocus
                className="dex-ui-input dex-ui-input--sm"
                value={query}
                onChange={e => setQuery(e.target.value)}
                placeholder={isDe ? 'Standort suchen …' : 'Search location …'}
                aria-label={isDe ? 'Standort suchen' : 'Search location'}
              />
            </div>
          )}
          {filtered.length === 0 ? (
            <div className="dex-ui-muted" style={{ padding: '12px 10px' }}>
              {isDe
                ? <>Kein Standort passt zu &bdquo;{query.trim()}&ldquo;.</>
                : <>No location matches &bdquo;{query.trim()}&ldquo;.</>}
            </div>
          ) : (
            filtered.map(loc => {
              const isChecked = selected.indexOf(loc) >= 0;
              return (
                // v31.2: dex-ui-row bringt den Hover mit — kein onMouseEnter/Leave-Inline-Style mehr.
                <label key={loc} className="dex-ui-row" style={{ cursor: 'pointer', padding: '8px 10px' }}>
                  <input
                    type="checkbox"
                    checked={isChecked}
                    onChange={() => toggle(loc)}
                    style={{ width: 16, height: 16, margin: 0, flexShrink: 0, accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
                  />
                  <span className="dex-ui-row-main dex-ui-row-title" style={isChecked ? { color: 'var(--dex-green-darker, #4a7c1f)' } : { fontWeight: 500 }}>
                    {loc}
                  </span>
                </label>
              );
            })
          )}
          {/* v31.2: Fußzeile immer da — links die Zählung, rechts „Alle wählen" und
              „Auswahl leeren": der Stand ist sichtbar, ohne die Liste hochzuscrollen. */}
          <div style={{
            position: 'sticky', bottom: 0, margin: '4px -6px -6px', padding: '6px 8px', background: '#fff',
            borderTop: '1px solid var(--dex-gray-100, #f5f5f5)',
            display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap',
          }}>
            <span className={cx('dex-ui-pill', selected.length > 0 ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')}>
              {selected.length === 0
                ? (isDe ? 'Kein Standort gewählt' : 'No location selected')
                : (isDe ? `${selected.length} von ${options.length} gewählt` : `${selected.length} of ${options.length} selected`)}
            </span>
            <span style={{ marginLeft: 'auto', display: 'inline-flex', gap: 2 }}>
              {filtered.length > 0 && !allVisibleChosen && (
                <button type="button" className="dex-ui-textbtn" onClick={selectVisible}>
                  {hasQuery ? (isDe ? 'Treffer wählen' : 'Select matches') : (isDe ? 'Alle wählen' : 'Select all')}
                </button>
              )}
              {selected.length > 0 && (
                <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => onChange([])}>
                  {isDe ? 'Auswahl leeren' : 'Clear selection'}
                </button>
              )}
            </span>
          </div>
        </div>
      )}
    </div>
  );
}
