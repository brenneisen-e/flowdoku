/**
 * v28.94: Aus `AdminPage` herausgeloest (547 Zeilen von 16.076).
 *
 * Das Aktionen-Menue des Organizer Centers: Die einzelnen Kacheln
 * (`ActionTile`) melden sich beim Mount über einen Context an, das
 * Dropdown (`ActionsDropdown`) liest die Registrierung und baut daraus die
 * gruppierte Liste. Die Gruppe hängt zusammen und ist deshalb EINE Datei —
 * sie kennt nichts vom Seiten-State ausser dem, was sie als Props bekommt.
 */
import * as React from 'react';
import { Search, X } from '../Icons';
import { ActionCategoryKey, ACTION_CATEGORY_ORDER, ACTION_CATEGORY_LABELS } from '../../data/actionCategories';

export interface ActionTileProps {
  icon: React.ReactNode;
  title: string;
  desc: string;
  badge: 'organizer' | 'admin';
  onClick?: () => void;
  href?: string;
  disabled?: boolean;
  busy?: boolean;
  result?: string | null;
  resultIsError?: boolean;
  // v9.19: filled-Variante für Highlight-Aktionen (z.B. Event aktivieren).
  // accent='green' = grün gefüllt, accent='red' = rot gefüllt.
  accent?: 'green' | 'red';
  // v20.3: Kategorie + optionale Unterkategorie (z.B. „Self-Check-in"
  // innerhalb von Check-in) für das gruppierte Aktionen-Dropdown.
  category?: ActionCategoryKey;
  subCategory?: string;
  // children: zusätzlicher Inhalt, der unterhalb der Standard-Tile-Inhalte
  // gerendert wird (z.B. das Excel-Dropdown-Menü).
  children?: React.ReactNode;
}
export function ActionTile(props: ActionTileProps): React.ReactElement | null {
  // v12.7: Wenn ActionTile innerhalb eines ActionsRegistryProvider gerendert
  // wird, registriert er sich dort (title, desc, onClick) statt eine
  // Kachel zu zeichnen. Children-Mode (Excel-Sub-Dropdown) bleibt
  // sichtbar — sonst gingen Modals/Dropdowns verloren.
  const registry = React.useContext(ActionsRegistryContext);
  const registered = !!registry && !props.children;
  const [hover, setHover] = React.useState(false);
  // v22.6: NUR die (stabilen) register/unregister-Funktionen als Effekt-Deps —
  // nicht das ganze Context-Objekt. Das war vorher bei jedem Provider-Render ein
  // neues Objekt und ließ den Effekt endlos neu feuern (Render-Schleife → das
  // Suchfeld im Aktionen-Dropdown war dadurch unbeschreibbar).
  const registryRegister = registry?.register;
  const registryUnregister = registry?.unregister;
  React.useEffect(() => {
    if (!registryRegister || !registryUnregister || !registered) return undefined;
    const key = props.title;
    registryRegister({
      key,
      title: props.title,
      desc: props.desc,
      badge: props.badge,
      onClick: props.onClick,
      href: props.href,
      disabled: props.disabled || props.busy,
      // v20.3: Kategorie-Zuordnung fürs gruppierte Dropdown (Fallback: Event).
      category: props.category || 'event',
      subCategory: props.subCategory,
    });
    return () => registryUnregister(key);
  }, [registryRegister, registryUnregister, registered, props.title, props.desc, props.badge, props.onClick, props.href, props.disabled, props.busy, props.category, props.subCategory]);
  if (registered) return null;
  const isInteractive = !props.disabled && !props.busy;
  const greenAccent = isInteractive && hover;
  // v9.19/v9.20: filled-Look — Tile dezent eingefärbt für
  // Highlight-Aktionen. Pastell statt voll gesättigt, damit nicht
  // alarmierend wirkt.
  const isFilled = !!props.accent;
  const filledBg = props.accent === 'green' ? '#e3f0c5' : props.accent === 'red' ? '#ffe5e5' : '';
  const filledBorder = props.accent === 'green' ? 'var(--dex-green, #86bc25)' : props.accent === 'red' ? 'var(--dex-red, #da291c)' : '';
  const borderColor = isFilled ? filledBorder : (greenAccent ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200, #e5e7eb)');
  const bg = isFilled ? filledBg : (greenAccent ? 'rgba(134,188,37,0.06)' : '#fff');
  // v9.20: bei pastell-gefüllten Tiles Text/Icon dunkel halten — auf
  // hellem Pastell-Hintergrund gut lesbar (im Gegensatz zum vorherigen
  // weiß auf saturated-Color).
  const filledIconColor = props.accent === 'green' ? 'var(--dex-green-dark, #4a7c1f)' : props.accent === 'red' ? '#a01e15' : 'var(--dex-gray-500, #6b7280)';
  const iconColor = isFilled ? filledIconColor : (greenAccent ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500, #6b7280)');
  const filledTextColor = isFilled
    ? (props.accent === 'green' ? 'var(--dex-green-dark, #3f5f10)' : props.accent === 'red' ? '#a01e15' : 'var(--dex-gray-800, #1f2937)')
    : 'var(--dex-gray-800, #1f2937)';
  const badgeLabel = props.badge === 'admin' ? 'Nur Admin' : 'Organizer';
  const badgeColors = props.badge === 'admin'
    ? { bg: 'rgba(237,139,0,0.12)', fg: 'var(--dex-orange, #ed8b00)' }
    : { bg: 'rgba(134,188,37,0.12)', fg: 'var(--dex-green-dark, #4a7c1f)' };
  const sharedStyle: React.CSSProperties = {
    textAlign: 'left', textDecoration: 'none', color: 'inherit',
    background: bg, border: `1px solid ${borderColor}`,
    borderRadius: 12, padding: 14,
    cursor: isInteractive ? 'pointer' : 'not-allowed',
    opacity: isInteractive ? 1 : 0.55,
    display: 'flex', flexDirection: 'column', gap: 8,
    fontFamily: 'inherit', fontSize: 'inherit',
    transition: 'all 0.15s ease',
    boxShadow: greenAccent ? '0 4px 12px rgba(134,188,37,0.18)' : 'none',
    position: 'relative',
    // width:100% sorgt dafür, dass die Kachel auch in einem flex-Wrapper
    // (z.B. Excel-Export hat einen <div display:flex>-Wrapper für das
    // Dropdown-Positioning) auf die volle Grid-Zellen-Breite gestreckt
    // wird — sonst sieht sie schmaler aus als die direkten Grid-Geschwister.
    width: '100%',
    boxSizing: 'border-box',
  };
  const inner = (
    <>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8, color: iconColor, transition: 'color 0.15s ease' }}>
          {props.icon}
          <span style={{ fontWeight: 600, fontSize: '0.88rem', color: filledTextColor }}>{props.title}</span>
        </span>
        <span style={{
          fontSize: '0.65rem', padding: '2px 8px', borderRadius: 999,
          // v9.20: Badge auf pastell Tiles in normalem badge-Look (auf hellem
          // Hintergrund gut sichtbar, im Gegensatz zur vorherigen
          // semi-transparenten weissen Variante auf saturated bg).
          background: badgeColors.bg,
          color: badgeColors.fg, fontWeight: 600,
          whiteSpace: 'nowrap', flexShrink: 0, letterSpacing: '0.02em',
        }}>{badgeLabel}</span>
      </div>
      {props.result && (
        <p style={{
          margin: 0, fontSize: '0.72rem',
          color: props.resultIsError ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)',
          fontStyle: 'italic',
        }}>{props.result}</p>
      )}
      {props.children}
      {hover && props.desc && (
        <div
          role="tooltip"
          style={{
            position: 'absolute',
            top: 'calc(100% + 8px)',
            left: 0,
            right: 0,
            zIndex: 50,
            background: 'var(--dex-gray-900, #1f2937)',
            color: '#fff',
            padding: '10px 12px',
            borderRadius: 8,
            fontSize: '0.76rem',
            lineHeight: 1.45,
            fontWeight: 400,
            boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
            pointerEvents: 'none',
            whiteSpace: 'normal',
          }}
        >
          {props.desc}
        </div>
      )}
    </>
  );
  if (props.href) {
    return (
      <a
        href={props.href}
        target="_blank"
        rel="noopener noreferrer"
        style={sharedStyle}
        onMouseEnter={() => setHover(true)}
        onMouseLeave={() => setHover(false)}
      >
        {inner}
      </a>
    );
  }
  return (
    <button
      type="button"
      disabled={!isInteractive}
      onClick={props.onClick}
      style={sharedStyle}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
    >
      {inner}
    </button>
  );
}

// v11.98: Pill-Toggle für die Aktiv-Teilnehmer-Tabelle bei Split-Kapazität.
// Default 'split' = getrennte Tabellen pro Gruppe. 'merged' = einzelne
// Tabelle (alter Look).
export function SplitMergeToggle(props: {
  view: 'split' | 'merged';
  setView: (v: 'split' | 'merged') => void;
  isDe: boolean;
}): React.ReactElement {
  const pill = (active: boolean): React.CSSProperties => ({
    padding: '5px 12px',
    borderRadius: 999,
    fontSize: '0.78rem',
    fontWeight: 600,
    cursor: 'pointer',
    border: `1px solid ${active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
    background: active ? 'rgba(134,188,37,0.10)' : '#fff',
    color: active ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
    transition: 'all 0.12s ease',
  });
  return (
    <div style={{ display: 'inline-flex', gap: 4, alignItems: 'center' }}>
      <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginRight: 6 }}>
        {props.isDe ? 'Ansicht:' : 'View:'}
      </span>
      <button type="button" onClick={() => props.setView('split')} style={pill(props.view === 'split')}>
        {props.isDe ? 'Getrennt' : 'Split'}
      </button>
      <button type="button" onClick={() => props.setView('merged')} style={pill(props.view === 'merged')}>
        {props.isDe ? 'Zusammen' : 'Merged'}
      </button>
    </div>
  );
}

// v12.7: Sammel-Card-Wrapper aus v12.6 entfernt — Aktionen leben jetzt
// als alphabetische Dropdown-Liste innerhalb der Event-Detail-Card
// (siehe ActionsDropdown weiter unten). Diese Komponente bleibt im
// Code für Backward-Compat, ihre Children werden display:none gerendert
// damit React-State + onClick-Handler weiterhin funktionieren.
export function ActionsCollapsibleCard(props: {
  isDe: boolean;
  children: React.ReactNode;
}): React.ReactElement {
  // v12.7: nicht mehr in eigener Card — wir verstecken die ganze Box
  // (display:none) und die ActionTiles registrieren sich via Context
  // im ActionsDropdown.
  void props.isDe;
  return (
    <div style={{ display: 'none' }}>
      {props.children}
    </div>
  );
}

// v12.7: Action-Registry — ActionTile-Instanzen melden sich beim Mount
// hier an. Der ActionsDropdown unten in der Event-Detail-Card liest den
// registry-State und rendert alle Einträge als alphabetisch sortierte
// Dropdown-Liste mit Hover-Tooltip (desc).
export interface RegisteredAction {
  key: string;
  title: string;
  desc: string;
  badge: 'organizer' | 'admin';
  onClick?: () => void;
  href?: string;
  disabled?: boolean;
  // v20.3: Kategorie + optionale Unterkategorie fürs gruppierte Dropdown.
  category: ActionCategoryKey;
  subCategory?: string;
}
export const ActionsRegistryContext = React.createContext<{
  register: (_a: RegisteredAction) => void;
  unregister: (_key: string) => void;
  actions: RegisteredAction[];
} | null>(null);

export function ActionsRegistryProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [actions, setActions] = React.useState<RegisteredAction[]>([]);
  const register = React.useCallback((a: RegisteredAction) => {
    setActions(prev => {
      const filtered = prev.filter(x => x.key !== a.key);
      return [...filtered, a];
    });
  }, []);
  const unregister = React.useCallback((key: string) => {
    setActions(prev => prev.filter(x => x.key !== key));
  }, []);
  // v22.6: Context-Value memoisieren — sonst entsteht bei jedem Render ein neues
  // Objekt, das die ActionTile-Register-Effekte erneut feuern lässt → Render-
  // Schleife (machte zuvor das Suchfeld im Aktionen-Dropdown unbeschreibbar).
  const value = React.useMemo(() => ({ register, unregister, actions }), [register, unregister, actions]);
  return React.createElement(ActionsRegistryContext.Provider, { value }, props.children);
}

// v22.50: Sprung aus der globalen Header-Suche in eine konkrete Aktion. Die
// Suche legt den Aktions-Key in localStorage ab; hier öffnen wir das Dropdown
// und filtern es auf den passenden Begriff vor. DE-/EN-Seed muss ein
// Teilstring des registrierten Aktions-Titels sein (Substring-Filter).
export const ACTION_FOCUS_SEED: Record<string, { de: string; en: string }> = {
  export: { de: 'Excel-Export', en: 'Excel export' },
  qr: { de: 'QR-Codes versenden', en: 'Send QR codes' },
  massmail: { de: 'E-Mail versenden', en: 'Send email' },
  invite: { de: 'Einladungsmail', en: 'Invitation email' },
  audit: { de: 'Audit-Log', en: 'Audit log' },
  selfcheckin: { de: 'Self-Check-in', en: 'self check-in' },
  idreorder: { de: 'IDs neu vergeben', en: 'Reassign IDs' },
  overbook: { de: 'Überbuchung', en: 'overbooking' },
  accessfix: { de: 'Zugriff reparieren', en: 'repair access' },
  fixcols: { de: 'Spalten fixen', en: 'Fix columns' },
};

export function ActionsDropdown(props: { isDe: boolean }): React.ReactElement | null {
  const ctx = React.useContext(ActionsRegistryContext);
  const [open, setOpen] = React.useState(false);
  // v20.3: aufklappbare Kategorien + Unterkategorien (z.B. „Self-Check-in"
  // unter Check-in) statt flacher Alphabet-Liste. Einträge sind mehrzeilig:
  // Titel fett, Beschreibung darunter — der frühere Hover-Tooltip entfällt.
  const [expanded, setExpanded] = React.useState<Set<string>>(new Set());
  // v22.5: Freitext-Suche über alle Aktionen (Titel + Beschreibung). Solange
  // etwas eingetippt ist, werden alle Kategorien automatisch aufgeklappt.
  const [query, setQuery] = React.useState('');
  const rootRef = React.useRef<HTMLDivElement | null>(null);
  const focusSeededRef = React.useRef(false);
  React.useEffect(() => {
    if (!open) return undefined;
    const onDocClick = (e: MouseEvent): void => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', onDocClick);
    return () => document.removeEventListener('mousedown', onDocClick);
  }, [open]);
  // v22.5: Suchfeld leeren, sobald das Dropdown geschlossen wird.
  React.useEffect(() => { if (!open) setQuery(''); }, [open]);
  // v22.50: Auto-Open + Vorfilter, wenn die Header-Suche eine Aktion angesteuert
  // hat. Einmalig, sobald Aktionen registriert sind.
  React.useEffect(() => {
    if (focusSeededRef.current) return;
    if (!ctx || ctx.actions.length === 0) return;
    let hint = '';
    try { hint = window.localStorage.getItem('dex_search_focus_action') || ''; } catch { /* */ }
    if (!hint) return;
    try { window.localStorage.removeItem('dex_search_focus_action'); } catch { /* */ }
    const seed = ACTION_FOCUS_SEED[hint];
    focusSeededRef.current = true;
    if (!seed) return;
    // v24.67: Wenn genau EINE registrierte Aktion auf den Seed-Begriff passt,
    // diese direkt auslösen (öffnet z.B. das „E-Mail versenden"-Modal) — statt
    // nur das gefilterte Aktionen-Dropdown zu zeigen. Sonst Fallback: Dropdown
    // vorgefiltert öffnen.
    const seedTerm = (props.isDe ? seed.de : seed.en).toLowerCase().trim();
    const matches = ctx.actions.filter(a => a.title.toLowerCase().indexOf(seedTerm) >= 0);
    if (matches.length === 1 && matches[0].onClick && !matches[0].disabled) {
      // kurz warten, damit der Admin-View vollständig gemountet ist.
      const target = matches[0];
      window.setTimeout(() => { try { target.onClick?.(); } catch { /* */ } }, 60);
      return;
    }
    setQuery(props.isDe ? seed.de : seed.en); setOpen(true);
  }, [ctx, props.isDe]);
  if (!ctx || ctx.actions.length === 0) return null;
  const lang = props.isDe ? 'de' : 'en';
  // v22.5: Kategorien alphabetisch nach lokalisiertem Label sortieren.
  const sortedCats = ACTION_CATEGORY_ORDER.slice().sort((a, b) => {
    const la = props.isDe ? ACTION_CATEGORY_LABELS[a].de : ACTION_CATEGORY_LABELS[a].en;
    const lb = props.isDe ? ACTION_CATEGORY_LABELS[b].de : ACTION_CATEGORY_LABELS[b].en;
    return la.localeCompare(lb, lang);
  });
  // v22.5: Aktiver Suchbegriff (klein geschrieben) + Treffer-Filter.
  const q = query.trim().toLowerCase();
  const matchesQuery = (a: RegisteredAction): boolean =>
    !q || a.title.toLowerCase().indexOf(q) >= 0 || (!!a.desc && a.desc.toLowerCase().indexOf(q) >= 0);
  const visibleActions = ctx.actions.filter(matchesQuery);
  const toggleKey = (k: string): void => {
    setExpanded(prev => {
      const next = new Set(prev);
      if (next.has(k)) next.delete(k); else next.add(k);
      return next;
    });
  };
  const runAction = (a: RegisteredAction): void => {
    if (a.disabled) return;
    setOpen(false);
    if (a.href) {
      window.open(a.href, '_blank', 'noopener,noreferrer');
    } else if (a.onClick) {
      a.onClick();
    }
  };
  const renderActionRow = (a: RegisteredAction, indent: number): React.ReactElement => {
    const adminOnly = a.badge === 'admin';
    return (
      <div
        key={a.key}
        onClick={() => runAction(a)}
        onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = 'rgba(134,188,37,0.07)'; }}
        onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = '#fff'; }}
        style={{
          padding: `9px 12px 9px ${indent}px`,
          cursor: a.disabled ? 'not-allowed' : 'pointer',
          borderBottom: '1px solid var(--dex-gray-100)',
          opacity: a.disabled ? 0.5 : 1,
          background: '#fff',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
          <span style={{ fontWeight: 700, fontSize: '0.87rem', color: 'var(--dex-gray-800)' }}>{a.title}</span>
          <span style={{
            fontSize: '0.68rem', padding: '2px 8px', borderRadius: 999,
            background: adminOnly ? 'rgba(237,139,0,0.12)' : 'rgba(134,188,37,0.12)',
            color: adminOnly ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green-dark, #4a7c1f)',
            fontWeight: 700, whiteSpace: 'nowrap', flexShrink: 0,
          }}>
            {adminOnly ? (props.isDe ? 'Nur Admin' : 'Admin only') : 'Organizer'}
          </span>
        </div>
        {a.desc && (
          <div style={{ marginTop: 3, fontSize: '0.76rem', color: 'var(--dex-gray-500)', lineHeight: 1.45 }}>
            {a.desc}
          </div>
        )}
      </div>
    );
  };
  return (
    <div ref={rootRef} style={{ position: 'relative', marginTop: 12 }}>
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        style={{
          width: '100%', textAlign: 'left', padding: '10px 14px',
          // v19.27: grün hinterlegt, damit die Aktionen-Auswahl deutlich auffällt.
          border: '1.5px solid var(--dex-green, #86bc25)', borderRadius: 10,
          background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-dark, #4a7c1f)',
          fontSize: '0.92rem', fontWeight: 700, cursor: 'pointer',
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12,
        }}
      >
        <span>{props.isDe ? `Aktion auswählen (${ctx.actions.length})` : `Pick an action (${ctx.actions.length})`}</span>
        <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.85rem', transform: open ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.15s ease' }}>▾</span>
      </button>
      {open && (
        <div style={{
          position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0,
          background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 10,
          boxShadow: '0 8px 24px rgba(0,0,0,0.12)', zIndex: 50,
          maxHeight: 480, overflowY: 'auto',
        }}>
          {/* v22.5: Suchfeld — filtert alle Aktionen quer über die Kategorien. */}
          <div style={{
            position: 'sticky', top: 0, zIndex: 2, background: '#fff',
            padding: 10, borderBottom: '1px solid var(--dex-gray-200)',
          }}>
            <div style={{ position: 'relative' }}>
              <span style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--dex-gray-400)', display: 'inline-flex', pointerEvents: 'none' }}>
                <Search size={15} />
              </span>
              <input
                type="text"
                autoFocus
                value={query}
                onChange={e => setQuery(e.target.value)}
                placeholder={props.isDe ? 'Aktion suchen…' : 'Search action…'}
                style={{
                  width: '100%', boxSizing: 'border-box', padding: '8px 30px 8px 32px',
                  border: '1px solid var(--dex-gray-300)', borderRadius: 8, fontSize: '0.85rem',
                  outline: 'none',
                }}
              />
              {query && (
                <button
                  type="button"
                  onClick={() => setQuery('')}
                  aria-label={props.isDe ? 'Suche leeren' : 'Clear search'}
                  style={{ position: 'absolute', right: 6, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400)', display: 'inline-flex', padding: 4 }}
                >
                  <X size={14} />
                </button>
              )}
            </div>
          </div>
          {q && visibleActions.length === 0 && (
            <div style={{ padding: '18px 14px', textAlign: 'center', color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
              {props.isDe ? 'Keine Aktion gefunden.' : 'No action found.'}
            </div>
          )}
          {sortedCats.map(catKey => {
            const inCat = visibleActions.filter(a => a.category === catKey);
            if (inCat.length === 0) return null;
            const catLabel = props.isDe ? ACTION_CATEGORY_LABELS[catKey].de : ACTION_CATEGORY_LABELS[catKey].en;
            const catDesc = props.isDe ? ACTION_CATEGORY_LABELS[catKey].descDe : ACTION_CATEGORY_LABELS[catKey].descEn;
            // v22.5: bei aktiver Suche alle Treffer-Kategorien automatisch öffnen.
            const catOpen = q ? true : expanded.has(catKey);
            const direct = inCat.filter(a => !a.subCategory).slice().sort((a, b) => a.title.localeCompare(b.title, lang));
            const subNames = Array.from(new Set(inCat.filter(a => !!a.subCategory).map(a => a.subCategory as string))).sort((a, b) => a.localeCompare(b, lang));
            return (
              <div key={catKey}>
                <div
                  onClick={() => toggleKey(catKey)}
                  onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = 'rgba(134,188,37,0.10)'; }}
                  onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = 'var(--dex-gray-50, #fafafa)'; }}
                  style={{
                    padding: '10px 12px', cursor: 'pointer', userSelect: 'none',
                    display: 'flex', alignItems: 'center', gap: 8,
                    background: 'var(--dex-gray-50, #fafafa)',
                    borderBottom: '1px solid var(--dex-gray-200)',
                    transition: 'background 0.12s ease',
                  }}
                >
                  <span style={{ width: 14, color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.8rem', flexShrink: 0 }}>{catOpen ? '▾' : '▸'}</span>
                  <span style={{ minWidth: 0, flex: 1 }}>
                    <span style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <span style={{ fontWeight: 800, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>{catLabel}</span>
                      <span style={{
                        fontSize: '0.68rem', padding: '1px 7px', borderRadius: 999,
                        background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700,
                      }}>{inCat.length}</span>
                    </span>
                    {/* v20.4: Kurzbeschreibung, was in der Kategorie steckt. */}
                    <span style={{ display: 'block', fontSize: '0.74rem', color: 'var(--dex-gray-500)', fontWeight: 400, marginTop: 1, lineHeight: 1.4 }}>
                      {catDesc}
                    </span>
                  </span>
                </div>
                {catOpen && direct.map(a => renderActionRow(a, 30))}
                {catOpen && subNames.map(sub => {
                  const subKey = `${catKey}::${sub}`;
                  const subOpen = q ? true : expanded.has(subKey);
                  const subActions = inCat.filter(a => a.subCategory === sub).slice().sort((a, b) => a.title.localeCompare(b.title, lang));
                  return (
                    <div key={subKey}>
                      <div
                        onClick={() => toggleKey(subKey)}
                        onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = 'rgba(134,188,37,0.12)'; }}
                        onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = 'rgba(134,188,37,0.05)'; }}
                        style={{
                          padding: '8px 12px 8px 30px', cursor: 'pointer', userSelect: 'none',
                          display: 'flex', alignItems: 'center', gap: 8,
                          background: 'rgba(134,188,37,0.05)',
                          borderBottom: '1px solid var(--dex-gray-100)',
                          transition: 'background 0.12s ease',
                        }}
                      >
                        <span style={{ width: 14, color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.75rem' }}>{subOpen ? '▾' : '▸'}</span>
                        <span style={{ fontWeight: 700, fontSize: '0.84rem', color: 'var(--dex-gray-700)' }}>{sub}</span>
                        <span style={{
                          fontSize: '0.66rem', padding: '1px 6px', borderRadius: 999,
                          background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700,
                        }}>{subActions.length}</span>
                      </div>
                      {subOpen && subActions.map(a => renderActionRow(a, 46))}
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// v12.6: Sammel-Card für alle Aktionen + Quick-Actions unter „Currently
// registered". v12.7: ersetzt durch ActionsDropdown + ActionsRegistry —
// die folgende Funktion bleibt aus Kompatibilitätsgründen als no-op.

