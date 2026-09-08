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
import { Search, X, ChevronDown, ExternalLink } from '../Icons';
import { cx } from '../dexUi';
import { ActionCategoryKey, ACTION_CATEGORY_ORDER, ACTION_CATEGORY_LABELS } from '../../data/actionCategories';

// v31.3: Beschreibungen sind bis zu fünf Zeilen lang (z.B. „Zugriff
// reparieren"). Sichtbar bleiben zwei — der volle Text steht im
// title-Attribut des Knopfs, und die Suche filtert weiter über den ganzen
// Text; es geht also keine Aussage verloren, die Liste wird nur lesbar.
const DESC_CLAMP: React.CSSProperties = { display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden' };

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
  // v31.3: kein Hover-State mehr — der Hover kommt aus `dex-ui-action`.
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
  // v31.3: Die Kachel ist ein `dex-ui-action` — Symbol, Titel, eine Zeile
  // Folge (Rest im title-Attribut), Rolle als Pill rechts. Hover kommt aus
  // der Klasse; der frühere Hover-Tooltip entfällt, weil die Beschreibung
  // jetzt sichtbar ist. `accent='red'` ist die Gefahren-Variante der Klasse;
  // `accent='green'` (Hervorheben, v9.19) bleibt eine grüne Kante links —
  // nur der linke Rand, damit der Hover-Rahmen der Klasse weiter greift.
  const badgeLabel = props.badge === 'admin' ? 'Nur Admin' : 'Organizer';
  const className = cx('dex-ui-action', props.accent === 'red' && 'dex-ui-action--danger');
  const sharedStyle: React.CSSProperties = {
    // width:100% sorgt dafür, dass die Kachel auch in einem flex-Wrapper
    // (z.B. Excel-Export hat einen <div display:flex>-Wrapper für das
    // Dropdown-Positioning) auf die volle Grid-Zellen-Breite gestreckt
    // wird — sonst sieht sie schmaler aus als die direkten Grid-Geschwister.
    width: '100%', textDecoration: 'none', position: 'relative',
    ...(props.accent === 'green' ? { borderLeft: '4px solid var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.06)' } : {}),
    // Ein Link kennt kein :disabled — gedämpft wie der Knopf, Klick bleibt (wie bisher).
    ...(!isInteractive ? { opacity: 0.5, cursor: 'not-allowed' } : {}),
  };
  const inner = (
    <>
      <span className="dex-ui-action-icon" aria-hidden="true">{props.icon}</span>
      <span className="dex-ui-action-body">
        <span className="dex-ui-action-title">{props.title}</span>
        {props.desc && <span className="dex-ui-action-desc" style={DESC_CLAMP}>{props.desc}</span>}
        {props.result && (
          <span
            className="dex-ui-action-desc"
            role="status"
            style={{ fontStyle: 'italic', color: props.resultIsError ? 'var(--dex-red, #c00)' : 'var(--dex-green-darker, #4a7c1f)' }}
          >{props.result}</span>
        )}
        {props.children}
      </span>
      <span className={cx('dex-ui-pill', 'dex-ui-action-badge', props.badge === 'admin' ? 'dex-ui-pill--orange' : 'dex-ui-pill--gray')}>{badgeLabel}</span>
    </>
  );
  if (props.href) {
    return (
      <a
        href={props.href}
        target="_blank"
        rel="noopener noreferrer"
        className={className}
        title={props.desc || undefined}
        aria-disabled={!isInteractive || undefined}
        style={sharedStyle}
      >
        {inner}
      </a>
    );
  }
  return (
    <button
      type="button"
      className={className}
      disabled={!isInteractive}
      aria-busy={props.busy || undefined}
      onClick={props.onClick}
      title={props.desc || undefined}
      style={sharedStyle}
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
  // v31.3: Segment-Reiter (`dex-ui-tabs`) statt zweier Pillen — Hover aus der
  // Klasse. Die Beschriftung sagt, was der Klick tut („eine Tabelle je
  // Gruppe" bzw. „eine gemeinsame Tabelle"), nicht nur „getrennt/zusammen".
  const isSplit = props.view === 'split';
  return (
    <div style={{ display: 'inline-flex', gap: 8, alignItems: 'center' }}>
      <span className="dex-ui-muted">{props.isDe ? 'Ansicht:' : 'View:'}</span>
      <div className="dex-ui-tabs" role="group" aria-label={props.isDe ? 'Ansicht der Teilnehmerliste' : 'Participant list view'}>
        <button type="button" className={cx('dex-ui-tab', isSplit && 'is-active')} aria-pressed={isSplit} onClick={() => props.setView('split')}>
          {props.isDe ? 'Je Gruppe' : 'Per group'}
        </button>
        <button type="button" className={cx('dex-ui-tab', !isSplit && 'is-active')} aria-pressed={!isSplit} onClick={() => props.setView('merged')}>
          {props.isDe ? 'Eine Tabelle' : 'One table'}
        </button>
      </div>
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
  // v31.3: Jede Aktion ist ein `dex-ui-action`-Knopf (Hover aus der Klasse,
  // per Tastatur erreichbar — die früheren <div onClick> waren es nicht).
  // Ein Symbol gibt es hier nicht: Die Registrierung trägt keines, und
  // `RegisteredAction` bleibt unverändert. Die Rolle steht als Pill rechts —
  // „Nur Admin" orange, weil es eine Einschränkung ist; „Organizer" grau,
  // weil es der Normalfall ist und Grün nur bedeuten soll, was aktiv ist.
  const renderActionRow = (a: RegisteredAction): React.ReactElement => {
    const adminOnly = a.badge === 'admin';
    return (
      <button
        key={a.key}
        type="button"
        className="dex-ui-action"
        disabled={!!a.disabled}
        onClick={() => runAction(a)}
        title={a.desc || undefined}
      >
        <span className="dex-ui-action-body">
          <span className="dex-ui-action-title">
            {a.title}
            {a.href && (
              <span
                style={{ display: 'inline-flex', verticalAlign: 'middle', marginLeft: 5, color: 'var(--dex-gray-400, #a0a0a0)' }}
                aria-label={props.isDe ? 'Öffnet in neuem Tab' : 'Opens in a new tab'}
                title={props.isDe ? 'Öffnet in neuem Tab' : 'Opens in a new tab'}
              >
                <ExternalLink size={12} />
              </span>
            )}
          </span>
          {a.desc && <span className="dex-ui-action-desc" style={DESC_CLAMP}>{a.desc}</span>}
        </span>
        <span className={cx('dex-ui-pill', 'dex-ui-action-badge', adminOnly ? 'dex-ui-pill--orange' : 'dex-ui-pill--gray')}>
          {adminOnly ? (props.isDe ? 'Nur Admin' : 'Admin only') : 'Organizer'}
        </span>
      </button>
    );
  };
  // v31.3: Gruppen- und Untergruppen-Köpfe sind Aufklapper-Knöpfe (Chevron
  // dreht über die Klasse, `aria-expanded` sagt dem Screenreader den Zustand).
  const renderGroupHead = (label: React.ReactNode, desc: string | undefined, count: number, isOpen: boolean, onToggle: () => void, sub: boolean): React.ReactElement => (
    <button
      type="button"
      className={cx('dex-ui-disclosure', isOpen && 'is-open')}
      aria-expanded={isOpen}
      onClick={onToggle}
      style={{ margin: 0, padding: sub ? '6px 6px' : '8px 6px', alignItems: 'flex-start' }}
    >
      <span className="dex-ui-disclosure-chevron" style={{ marginTop: sub ? 1 : 2 }}><ChevronDown size={16} /></span>
      <span style={{ flex: 1, minWidth: 0 }}>
        {sub
          ? <span style={{ display: 'block', fontWeight: 600, fontSize: '0.84rem', color: 'var(--dex-gray-700, #444)' }}>{label}</span>
          : <span className="dex-ui-action-group-title" style={{ margin: 0, display: 'block' }}>{label}</span>}
        {/* v20.4: Kurzbeschreibung, was in der Kategorie steckt. */}
        {desc && <span className="dex-ui-action-desc" style={{ fontWeight: 400 }}>{desc}</span>}
      </span>
      <span className="dex-ui-pill dex-ui-pill--gray" style={{ flexShrink: 0 }}>{count}</span>
    </button>
  );
  return (
    <div ref={rootRef} style={{ position: 'relative', marginTop: 12 }}>
      {/* v19.27: grün umrandet, damit die Aktionen-Auswahl deutlich auffällt.
          v31.3: als `btn-outline` — Hover (grün gefüllt) kommt aus der Klasse. */}
      <button
        type="button"
        className="btn btn-outline btn-block"
        aria-haspopup="true"
        aria-expanded={open}
        onClick={() => setOpen(o => !o)}
        style={{ justifyContent: 'space-between', textAlign: 'left', cursor: 'pointer' }}
      >
        <span>{props.isDe ? `Aktion auswählen (${ctx.actions.length})` : `Pick an action (${ctx.actions.length})`}</span>
        <span className={cx('dex-ui-disclosure-chevron', open && 'is-open')} style={{ color: 'inherit' }} aria-hidden="true"><ChevronDown size={16} /></span>
      </button>
      {open && (
        <div
          className="dex-ui-card dex-ui-fade-in"
          role="region"
          aria-label={props.isDe ? 'Aktionen' : 'Actions'}
          style={{
            position: 'absolute', top: 'calc(100% + 6px)', left: 0, right: 0, zIndex: 50,
            padding: 0, maxHeight: 480, overflowY: 'auto',
            boxShadow: '0 12px 32px rgba(0,0,0,0.12)',
          }}
        >
          {/* v22.5: Suchfeld — filtert alle Aktionen quer über die Kategorien. */}
          <div style={{ position: 'sticky', top: 0, zIndex: 2, background: '#fff', padding: 10, borderBottom: '1px solid var(--dex-gray-100, #f5f5f5)' }}>
            <div className="dex-ui-searchbar" style={{ maxWidth: 'none' }}>
              <span className="dex-ui-searchbar-icon"><Search size={15} /></span>
              <input
                type="text"
                className="dex-ui-input dex-ui-input--sm"
                autoFocus
                value={query}
                onChange={e => setQuery(e.target.value)}
                placeholder={props.isDe ? 'Aktion suchen…' : 'Search action…'}
                aria-label={props.isDe ? 'Aktion suchen' : 'Search action'}
                style={{ paddingRight: 34 }}
              />
              {query && (
                <button
                  type="button"
                  className="dex-ui-iconbtn"
                  onClick={() => setQuery('')}
                  aria-label={props.isDe ? 'Suche leeren' : 'Clear search'}
                  style={{ position: 'absolute', right: 3, top: '50%', transform: 'translateY(-50%)', width: 26, height: 26 }}
                >
                  <X size={14} />
                </button>
              )}
            </div>
          </div>
          <div style={{ padding: '6px 10px 10px' }}>
            {q && visibleActions.length === 0 && (
              <div className="dex-ui-empty" style={{ padding: '18px 14px' }}>
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
                <div key={catKey} className="dex-ui-action-group" style={{ marginTop: 4 }}>
                  {renderGroupHead(catLabel, catDesc, inCat.length, catOpen, () => toggleKey(catKey), false)}
                  {catOpen && direct.length > 0 && (
                    <div className="dex-ui-action-grid" style={{ padding: '4px 0 6px' }}>
                      {direct.map(a => renderActionRow(a))}
                    </div>
                  )}
                  {catOpen && subNames.map(sub => {
                    const subKey = `${catKey}::${sub}`;
                    const subOpen = q ? true : expanded.has(subKey);
                    const subActions = inCat.filter(a => a.subCategory === sub).slice().sort((a, b) => a.title.localeCompare(b.title, lang));
                    return (
                      <div key={subKey} style={{ paddingLeft: 22 }}>
                        {renderGroupHead(sub, undefined, subActions.length, subOpen, () => toggleKey(subKey), true)}
                        {subOpen && (
                          <div className="dex-ui-action-grid" style={{ padding: '4px 0 6px' }}>
                            {subActions.map(a => renderActionRow(a))}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

// v12.6: Sammel-Card für alle Aktionen + Quick-Actions unter „Currently
// registered". v12.7: ersetzt durch ActionsDropdown + ActionsRegistry —
// die folgende Funktion bleibt aus Kompatibilitätsgründen als no-op.

