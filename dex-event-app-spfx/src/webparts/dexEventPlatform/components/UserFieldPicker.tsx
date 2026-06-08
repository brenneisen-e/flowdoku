// People-Picker für Custom-Fields vom Typ `user` / `roommate`.
// v18.61: Aus RegistrationPage.tsx in eine eigene Datei ausgelagert, damit
// auch das „Angaben ergänzen"-Edit-Formular in MyEventsPage denselben Picker
// (mit Profilfoto + Such-Dropdown) nutzen kann — vorher wurde das Feld dort
// als simpler Text-Input gerendert („Name <email>" als Rohtext).
import * as React from 'react';
import { useLanguage } from '../context/LanguageContext';
import InternationalSearchToggle from './InternationalSearchToggle';

export function UserFieldPicker(props: {
  value: string;
  onChange: (v: string) => void;
  searchUsers: (q: string, includeIntl?: boolean) => Promise<Array<{ email: string; displayName: string; location?: string; jobTitle?: string }>>;
  // v11.98: Profil-Lookup für die selektierte Person, damit der Chip
  // auch nach Reload den JobTitle + Standort zeigt (props.value enthält
  // nur "Name <email>" — Standort/Title wären sonst weg).
  searchUserByEmail?: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string } | null>;
  placeholder: string;
  errorStyle: React.CSSProperties;
  hint?: string;
  // v18.56: erzwungene Anmeldesprache des Events durchreichen, damit auch der
  // „Auch international suchen"-Toggle der Event-Sprache folgt (nicht der
  // App-Sprache des Teilnehmers). undefined = App-Sprache.
  forcedIsDe?: boolean;
}): React.ReactElement {
  // Parse "Name <email>" aus dem gespeicherten Wert (z.B. nach Remount/Reload),
  // damit der Chip mit Foto sofort wieder erscheint.
  const parseValue = (v: string): { name: string; email: string } | null => {
    const m = (v || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
    if (!m) return null;
    return { name: m[1].trim(), email: m[2].trim() };
  };
  const { locale } = useLanguage();
  const isDe = props.forcedIsDe !== undefined ? props.forcedIsDe : (locale === 'de');
  const initialParsed = parseValue(props.value);
  const [query, setQuery] = React.useState(initialParsed ? '' : (props.value || ''));
  const [results, setResults] = React.useState<Array<{ email: string; displayName: string; location?: string; jobTitle?: string }>>([]);
  const [isSearching, setIsSearching] = React.useState(false);
  const [includeIntl, setIncludeIntl] = React.useState(false);
  // v11.91: Selected merkt sich zusätzlich Standort und JobTitle, damit
  // der Chip dieselben Infos zeigt wie die Dropdown-Treffer.
  const [selected, setSelected] = React.useState<{ name: string; email: string; location?: string; jobTitle?: string } | null>(initialParsed);
  const timerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  React.useEffect(() => {
    const parsed = parseValue(props.value);
    setSelected(parsed);
    if (parsed) setQuery('');
    else setQuery(props.value || '');
    // v11.98: nach Remount / Reload jobTitle + Standort lazy nachladen,
    // damit der Chip die volle Profil-Info zeigt (props.value hält nur
    // „Name <email>", deshalb fehlen die Profil-Properties initial).
    if (parsed && props.searchUserByEmail) {
      props.searchUserByEmail(parsed.email).then(p => {
        if (p && (p.jobTitle || p.location)) {
          setSelected({ name: parsed.name, email: parsed.email, jobTitle: p.jobTitle || '', location: p.location || '' });
        }
      }).catch(() => { /* silent */ });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.value]);
  const hasSelection = !!selected;
  const clearSelection = (): void => {
    setSelected(null);
    setQuery('');
    props.onChange('');
    setResults([]);
  };
  // v18.5: Ergebnis-Dropdown als position:fixed rendern, anker an die
  // Wrapper-Bounding-Box. Vorher (position:absolute) wurde es vom
  // `overflow:hidden` der Karte abgeschnitten.
  const rootRef = React.useRef<HTMLDivElement | null>(null);
  const [menuPos, setMenuPos] = React.useState<{ left: number; top: number; width: number } | null>(null);
  const recalcMenuPos = React.useCallback(() => {
    const el = rootRef.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    setMenuPos({ left: r.left, top: r.bottom + 2, width: r.width });
  }, []);
  const menuOpen = !hasSelection && results.length > 0;
  React.useEffect(() => {
    if (!menuOpen) return undefined;
    recalcMenuPos();
    window.addEventListener('scroll', recalcMenuPos, true);
    window.addEventListener('resize', recalcMenuPos);
    return () => {
      window.removeEventListener('scroll', recalcMenuPos, true);
      window.removeEventListener('resize', recalcMenuPos);
    };
  }, [menuOpen, recalcMenuPos]);
  return (
    <div ref={rootRef} style={{ position: 'relative' }}>
      {hasSelection && selected ? (
        <div style={{
          display: 'inline-flex', alignItems: 'center', gap: 10,
          padding: '8px 12px 8px 8px',
          // v19.0: grüne Hervorhebung wie ausgefüllte Felder.
          border: '1px solid var(--dex-green, #86bc25)',
          borderRadius: 'var(--dex-radius)',
          background: 'rgba(134,188,37,0.06)',
          maxWidth: '100%',
        }}>
          <img
            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(selected.email)}&size=L`}
            alt={selected.name}
            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
            style={{ width: 40, height: 40, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0, transition: 'transform 0.15s', transformOrigin: 'left center' }}
            onMouseEnter={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.4)'; (e.currentTarget as HTMLImageElement).style.zIndex = '20'; (e.currentTarget as HTMLImageElement).style.position = 'relative'; (e.currentTarget as HTMLImageElement).style.boxShadow = '0 6px 18px rgba(0,0,0,0.18)'; }}
            onMouseLeave={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)'; (e.currentTarget as HTMLImageElement).style.boxShadow = 'none'; }}
          />
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontWeight: 600, fontSize: '0.9rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {selected.name}
            </div>
            <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              <a href={`mailto:${selected.email}`} onClick={e => e.stopPropagation()} style={{ color: 'inherit', textDecoration: 'none' }}>{selected.email}</a>
            </div>
            {(selected.jobTitle || selected.location) && (
              <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.72rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {[selected.jobTitle, selected.location].filter(Boolean).join(' · ')}
              </div>
            )}
          </div>
          <button
            type="button"
            onClick={clearSelection}
            title="Auswahl entfernen"
            style={{
              background: 'var(--dex-gray-200)', border: 'none', color: 'var(--dex-gray-700)',
              width: 22, height: 22, borderRadius: '50%', cursor: 'pointer',
              fontSize: '0.9rem', lineHeight: 1,
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              flexShrink: 0,
            }}
          >×</button>
        </div>
      ) : (
        <input
          className="form-input"
          value={query}
          onChange={e => {
            const val = e.target.value;
            setQuery(val);
            // Freitext ist kein gültiger Wert; erst nach Dropdown-Auswahl.
            props.onChange('');
            if (timerRef.current) clearTimeout(timerRef.current);
            if (val.length >= 2) {
              timerRef.current = setTimeout(async () => {
                setIsSearching(true);
                try { setResults(await props.searchUsers(val, includeIntl)); }
                catch { setResults([]); }
                setIsSearching(false);
              }, 300);
            } else {
              setResults([]);
            }
          }}
          onBlur={() => {
            // Wenn keine gültige Person ausgewählt wurde, Feld leeren.
            setTimeout(() => {
              if (!selected) {
                setQuery('');
                props.onChange('');
                setResults([]);
              }
            }, 150);
          }}
          placeholder={props.placeholder}
          style={props.errorStyle}
        />
      )}
      {!hasSelection && (
        <InternationalSearchToggle
          checked={includeIntl}
          onChange={async next => {
            setIncludeIntl(next);
            const val = query.trim();
            if (val.length >= 2) {
              setIsSearching(true);
              try { setResults(await props.searchUsers(val, next)); }
              catch { setResults([]); }
              setIsSearching(false);
            }
          }}
          isDe={isDe}
          compact
        />
      )}
      {isSearching && !hasSelection && (
        <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>{isDe ? 'Suche…' : 'Searching…'}</div>
      )}
      {props.hint && (
        <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 4, marginBottom: 0, fontStyle: 'italic' }}>
          {props.hint}
        </p>
      )}
      {menuOpen && (
        <div style={{
          position: 'fixed',
          left: menuPos?.left ?? 0, top: menuPos?.top ?? 0, width: menuPos?.width ?? 'auto',
          zIndex: 3000,
          background: '#fff', border: '1px solid var(--dex-gray-200)',
          borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
          maxHeight: 280, overflowY: 'auto', marginTop: 0,
        }}>
          {results.map(u => (
            <div
              key={u.email}
              style={{
                padding: '8px 12px', cursor: 'pointer', fontSize: '0.85rem',
                borderBottom: '1px solid var(--dex-gray-100)',
                display: 'flex', alignItems: 'center', gap: 10,
              }}
              onMouseDown={() => {
                const formatted = `${u.displayName} <${u.email}>`;
                setSelected({ name: u.displayName, email: u.email, location: u.location, jobTitle: u.jobTitle });
                setQuery('');
                props.onChange(formatted);
                setResults([]);
              }}
            >
              <img
                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(u.email)}&size=L`}
                alt={u.displayName}
                onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                style={{ width: 40, height: 40, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0, transition: 'transform 0.15s', transformOrigin: 'left center' }}
                onMouseEnter={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.4)'; (e.currentTarget as HTMLImageElement).style.zIndex = '20'; (e.currentTarget as HTMLImageElement).style.position = 'relative'; (e.currentTarget as HTMLImageElement).style.boxShadow = '0 6px 18px rgba(0,0,0,0.18)'; }}
                onMouseLeave={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)'; (e.currentTarget as HTMLImageElement).style.boxShadow = 'none'; }}
              />
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {u.email}
                </div>
                {(u.jobTitle || u.location) && (
                  <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.72rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {[u.jobTitle, u.location].filter(Boolean).join(' · ')}
                  </div>
                )}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default UserFieldPicker;
