/* AgendaEditor — v30.94. Der Programm-Editor des Wizards (Schritt 3), für das
 * Hauptevent UND die Sub-Events dieselbe Komponente.
 *
 * Nutzer-Befund 07.09.2026 (27 Punkte über vier Tage): „das ist doch nicht
 * übersichtlich", „warum ist das englische Schreibweise?" und „es sollte
 * sowas geben wie Cluster, als Vorschlag Day 1, und dann kann man die
 * Programmpunkte darunter sinnvoll zusammenführen". Drei Ursachen:
 *
 *  1. Jeder Punkt war eine eigene Karte mit sechs beschrifteten Feldern, das
 *     Datum in jeder Zeile neu. 27 Karten × 6 Labels = eine Wand, in der man
 *     den Tag nicht mehr sieht. Jetzt: CLUSTER (utils/agendaGroups). Ein
 *     Cluster hat einen Namen (Vorschlag „Tag 1", frei änderbar) und ein
 *     Datum, das genau einmal im Kopf steht; die Zeilen darunter tragen nur
 *     Start, Ende, Titel, Raum. Beschreibung ist eine zweite Zeile zum
 *     Aufklappen. Punkte lassen sich per Auswahl in einen anderen Cluster
 *     verschieben; ein Präfix wie „Day 1 - " in den Titeln entfernt ein Knopf.
 *
 *  2. Datum und Zeit waren native `<input type="date">`/`type="time">`. Deren
 *     Anzeige (MM/DD/YYYY, 11:00 AM) bestimmt der Browser nach seiner
 *     Sprache — nicht die App. Das `lang`-Attribut aus v30.60 sollte das
 *     richten; der Screenshot zeigt, dass Chrome/Edge es hier ignorieren.
 *     Deshalb keine nativen Felder mehr: das Datum kommt aus demselben
 *     react-datepicker wie in Schritt 1 (dd.MM.yyyy, locale de), die Zeiten
 *     sind Textfelder im 24-Stunden-Format, die „9", „930", „9.30" und
 *     „09:30" gleichermaßen verstehen und beim Verlassen zu HH:MM machen.
 *
 *  3. Die Titel trugen den Cluster als Behelf („Day 1 - Welcome") — mit dem
 *     Feld `cluster` gehört das dorthin, nicht in jeden Titel.
 *
 * Datenmodell: `AgendaItem[]` mit `date`, `time`/`endTime` und neu `cluster`.
 * Der Cluster ist keine eigene Struktur, sondern die Gruppe aller Punkte mit
 * demselben Namen (ohne Namen: demselben Datum). Cluster-Aktionen (Umbenennen,
 * Datum ändern, kopieren, löschen) wirken deshalb auf alle Punkte der Gruppe.
 *
 * v31.2 (UI-Leitfaden): dex-ui-Klassen statt Inline-Styles — Zeilen und
 * Knöpfe mit Hover, beschrifteter Kopf, leerer Zustand, Rückfrage mit Folge.
 * Verhalten, Props und Datenmodell unverändert. */
import * as React from 'react';
import DatePicker from 'react-datepicker';
import { AgendaItem } from '../../types';
import { Plus, X, Copy, Trash2, FileText, Calendar, Info, AlertCircle } from '../Icons';
import { cx } from '../dexUi';
import { agendaGroups, AgendaGroup, sortAgenda, suggestClusterName, stripClusterPrefix, countClusterPrefixed, nextClusterName } from '../../utils/agendaGroups';

export interface AgendaEditorProps {
  items: AgendaItem[];
  /** Funktionaler Updater — der Aufrufer reicht `setAgenda` bzw. einen
   *  Wrapper für das Sub-Event-Draft durch. */
  onChange: (updater: (prev: AgendaItem[]) => AgendaItem[]) => void;
  isDe: boolean;
  isMobile: boolean;
  termSingular: string;
  termPlural: string;
  /** YYYY-MM-DD — Vorbelegung des ersten Tags (Start des Events/Termins). */
  defaultDate: string;
}

/** „9", „9:5", „930", „09.30", „9h30" → „09:30"; leer → ''; unlesbar → null. */
export function normalizeTimeInput(raw: string): string | null {
  const s = (raw || '').trim().replace(/\s+/g, '').replace(/[.,;hH]/g, ':');
  if (!s) return '';
  let h = NaN; let mi = NaN;
  let m = /^(\d{1,2}):?(\d{2})$/.exec(s);
  if (m) { h = +m[1]; mi = +m[2]; }
  else {
    m = /^(\d{1,2}):?$/.exec(s);
    if (!m) return null;
    h = +m[1]; mi = 0;
  }
  if (h > 23 || mi > 59) return null;
  return `${String(h).padStart(2, '0')}:${String(mi).padStart(2, '0')}`;
}

function toYmd(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function fromYmd(ymd: string): Date | null {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return null;
  const d = new Date(ymd + 'T00:00:00');
  return isNaN(d.getTime()) ? null : d;
}
function addDays(ymd: string, n: number): string {
  const d = fromYmd(ymd) || new Date();
  d.setDate(d.getDate() + n);
  return toYmd(d);
}
function newId(): string {
  return `ag-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
}
function weekdayOf(ymd: string, isDe: boolean): string {
  const d = fromYmd(ymd);
  if (!d) return '';
  return d.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { weekday: 'long' });
}

/** Zeit-Textfeld: lokaler Tipp-Zustand, Normalisierung beim Verlassen. Ungültig
 *  bleibt rot stehen und wird NICHT geschrieben — ein halb getipptes „1" darf
 *  nicht als 01:00 in der Agenda landen. */
const TimeField: React.FC<{ value: string; onCommit: (_v: string) => void; placeholder: string; title: string }> = ({ value, onCommit, placeholder, title }) => {
  const [txt, setTxt] = React.useState(value);
  const [bad, setBad] = React.useState(false);
  React.useEffect(() => { setTxt(value); setBad(false); }, [value]);
  const commit = (): void => {
    const n = normalizeTimeInput(txt);
    if (n === null) { setBad(true); return; }
    setBad(false);
    setTxt(n);
    if (n !== value) onCommit(n);
  };
  return (
    <input
      type="text"
      inputMode="numeric"
      className="dex-ui-input dex-ui-input--sm"
      value={txt}
      placeholder={placeholder}
      title={bad ? `${title} — HH:MM` : title}
      aria-label={title}
      aria-invalid={bad || undefined}
      onChange={e => setTxt(e.target.value)}
      onBlur={commit}
      onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); (e.target as HTMLInputElement).blur(); } }}
      style={{ padding: '6px 4px', textAlign: 'center', fontVariantNumeric: 'tabular-nums', borderColor: bad ? 'var(--dex-red, #da291c)' : undefined, background: bad ? 'rgba(218,41,28,0.06)' : undefined }}
    />
  );
};

export const AgendaEditor: React.FC<AgendaEditorProps> = (p) => {
  const { items, onChange, isDe, isMobile, defaultDate } = p;
  const termS = p.termSingular.trim() || (isDe ? 'Programmpunkt' : 'Agenda item');
  const termP = p.termPlural.trim() || (isDe ? 'Programmpunkte' : 'Agenda items');
  const [descOpen, setDescOpen] = React.useState<Record<string, boolean>>({});
  const [confirmKey, setConfirmKey] = React.useState<string>('');
  // Cluster-Name wird lokal getippt und beim Verlassen geschrieben — sonst
  // wechselt die Gruppe bei jedem Zeichen ihren Schlüssel und das Feld
  // verliert den Fokus.
  const [nameEdit, setNameEdit] = React.useState<{ key: string; value: string } | null>(null);
  const focusRef = React.useRef<string | null>(null);

  const groups = React.useMemo(() => agendaGroups(items), [items]);
  const sortedAll = React.useMemo(() => sortAgenda(items), [items]);
  const globalNo = (id: string): number => sortedAll.findIndex(a => a.id === id) + 1;
  const labelOf = (g: AgendaGroup, gi: number): string => g.cluster || suggestClusterName(gi, isDe);

  const blank = (date: string, time: string, cluster: string): AgendaItem =>
    ({ id: newId(), date, time, endTime: '', icon: 'Calendar', title: '', description: '', location: '', cluster: cluster || undefined });

  const patch = (id: string, upd: Partial<AgendaItem>): void => onChange(prev => prev.map(a => a.id === id ? { ...a, ...upd } : a));
  const remove = (id: string): void => onChange(prev => prev.filter(a => a.id !== id));
  const patchGroup = (g: AgendaGroup, upd: (a: AgendaItem) => AgendaItem): void => {
    const ids = g.items.map(a => a.id);
    onChange(prev => prev.map(a => ids.indexOf(a.id) >= 0 ? upd(a) : a));
  };

  /** Neuer Punkt in einer Gruppe — Startzeit = Ende des Vorgängers, damit ein
   *  Tagesablauf ohne Nachtippen weiterläuft. Ohne gespeicherten Cluster
   *  bekommt der Punkt den vorgeschlagenen Namen NICHT — die Gruppe bleibt
   *  eine Datums-Gruppe, bis der Organizer sie benennt. */
  const addTo = (g: AgendaGroup, afterId?: string): void => {
    const anchor = afterId ? g.items.find(a => a.id === afterId) : g.items[g.items.length - 1];
    const n = blank((anchor && anchor.date) || g.date, (anchor && anchor.endTime) || '', g.cluster);
    focusRef.current = n.id;
    onChange(prev => {
      if (!afterId) return [...prev, n];
      const idx = prev.findIndex(a => a.id === afterId);
      return idx < 0 ? [...prev, n] : [...prev.slice(0, idx + 1), n, ...prev.slice(idx + 1)];
    });
  };
  /** Neuer Cluster: Vorschlagsname „Tag N", Datum = letzter Tag + 1. */
  const addGroup = (): void => {
    const dates = groups.map(g => g.dates[g.dates.length - 1]).filter(Boolean).sort();
    const date = dates.length ? addDays(dates[dates.length - 1], 1) : (defaultDate || toYmd(new Date()));
    let name = suggestClusterName(groups.length, isDe);
    const taken = groups.map(g => labelOf(g, groups.indexOf(g)));
    let k = groups.length;
    while (taken.indexOf(name) >= 0) { k++; name = suggestClusterName(k, isDe); }
    const n = blank(date, '', name);
    focusRef.current = n.id;
    onChange(prev => [...prev, n]);
  };
  const renameGroup = (g: AgendaGroup, name: string): void => {
    const v = name.trim();
    if (v === g.cluster) return;
    patchGroup(g, a => ({ ...a, cluster: v || undefined }));
  };
  const redateGroup = (g: AgendaGroup, to: string): void => {
    if (!to) return;
    // Mehrtägiger Cluster: alle Tage um dieselbe Differenz verschieben.
    const from = g.date || g.dates[0] || to;
    const diff = Math.round((new Date(to + 'T00:00:00').getTime() - new Date(from + 'T00:00:00').getTime()) / 86400000);
    patchGroup(g, a => ({ ...a, date: a.date ? addDays(a.date, diff) : to }));
  };
  /** Kopie der Gruppe auf den nächsten freien Tag; Name zählt hoch („Tag 1"
   *  → „Tag 2"). Neue Ids, damit Check-ins (v30.91) am richtigen Punkt hängen. */
  const duplicateGroup = (g: AgendaGroup, gi: number): void => {
    const used = groups.reduce<string[]>((acc, x) => acc.concat(x.dates), []);
    let next = addDays(g.date || defaultDate || toYmd(new Date()), 1);
    while (used.indexOf(next) >= 0) next = addDays(next, 1);
    const base = labelOf(g, gi);
    let name = nextClusterName(base) || suggestClusterName(groups.length, isDe);
    const taken = groups.map((x, i) => labelOf(x, i));
    while (taken.indexOf(name) >= 0) name = nextClusterName(name);
    const clones = g.items.map(a => ({ ...a, id: newId(), date: next, cluster: name }));
    onChange(prev => [...prev, ...clones]);
  };
  const deleteGroup = (g: AgendaGroup): void => {
    const ids = g.items.map(a => a.id);
    onChange(prev => prev.filter(a => ids.indexOf(a.id) < 0));
    setConfirmKey('');
  };
  const stripPrefixes = (g: AgendaGroup): void => {
    patchGroup(g, a => ({ ...a, title: stripClusterPrefix(a.title || '', g.cluster) }));
  };
  const moveItem = (id: string, targetKey: string): void => {
    const target = groups.find(g => g.key === targetKey);
    if (!target) return;
    patch(id, { cluster: target.cluster || undefined, date: target.date || '' });
  };

  const focusIfPending = (el: HTMLInputElement | null, id: string): void => {
    if (el && focusRef.current === id) { focusRef.current = null; el.focus(); }
  };

  // Spalten: # | Start | Ende | Titel | Raum | Cluster-Wechsel | Beschreibung | Löschen
  // v31.2: Cluster-Wechsel zeigt den Zielnamen statt eines 34-px-Pfeils; Symbol-Knöpfe 32 px.
  const cols = isMobile
    ? '24px 60px 60px 1fr 32px 32px'
    : `28px 66px 66px minmax(160px, 1fr) minmax(110px, 190px) ${groups.length > 1 ? 'minmax(96px, 120px) ' : ''}32px 32px`;
  const head: React.CSSProperties = { fontSize: '0.68rem', fontWeight: 600, color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.3, padding: '0 4px', marginBottom: 2 };

  const renderRow = (item: AgendaItem, g: AgendaGroup): React.ReactElement => {
    const showDesc = descOpen[item.id] || !!(item.description || '').trim();
    const moveTitle = isDe ? 'In einen anderen Cluster verschieben' : 'Move to another cluster';
    return (
      <div key={item.id} style={{ borderTop: '1px solid var(--dex-gray-100)', padding: '2px 0' }}>
        {/* v31.2: dex-ui-row liefert den Zeilen-Hover, das Raster hält die
            Spalten unter dem Kopf — deshalb Klasse UND display:grid. */}
        <div className="dex-ui-row" style={{ display: 'grid', gridTemplateColumns: cols, gap: 6, padding: '3px 4px' }}>
          <span style={{ width: 24, height: 24, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', borderRadius: '50%', background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 700, fontSize: '0.74rem', lineHeight: 1 }}>{globalNo(item.id)}</span>
          <TimeField value={item.time || ''} onCommit={v => patch(item.id, { time: v })} placeholder="09:00" title="Start" />
          <TimeField value={item.endTime || ''} onCommit={v => patch(item.id, { endTime: v })} placeholder="10:30" title={isDe ? 'Ende' : 'End'} />
          <input
            type="text" className="dex-ui-input dex-ui-input--sm" value={item.title}
            ref={el => focusIfPending(el, item.id)}
            onChange={e => patch(item.id, { title: e.target.value })}
            onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); addTo(g, item.id); } }}
            aria-label={isDe ? 'Titel' : 'Title'}
            placeholder={isDe ? `Titel — Enter legt den nächsten ${termS} an` : `Title — Enter adds the next ${termS.toLowerCase()}`}
            style={{ fontSize: '0.88rem' }}
          />
          {!isMobile && (
            <input type="text" className="dex-ui-input dex-ui-input--sm" value={item.location || ''} onChange={e => patch(item.id, { location: e.target.value })} aria-label={isDe ? 'Raum' : 'Room'} placeholder={isDe ? 'Raum (optional)' : 'Room (optional)'} />
          )}
          {!isMobile && groups.length > 1 && (
            <select value={g.key} onChange={e => moveItem(item.id, e.target.value)} title={moveTitle} aria-label={moveTitle}
              className="dex-ui-select dex-ui-select--sm" style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', cursor: 'pointer' }}>
              {groups.map((x, i) => <option key={x.key} value={x.key}>{i + 1}: {labelOf(x, i)}</option>)}
            </select>
          )}
          <button type="button" onClick={() => setDescOpen(o => ({ ...o, [item.id]: !showDesc }))}
            className={cx('dex-ui-iconbtn', showDesc && 'dex-ui-iconbtn--green is-active')}
            title={isDe ? 'Beschreibung ein-/ausblenden' : 'Toggle description'}
            aria-label={isDe ? 'Beschreibung' : 'Description'}
            aria-pressed={showDesc}>
            <FileText size={15} />
          </button>
          <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => remove(item.id)} aria-label={isDe ? 'Punkt entfernen' : 'Remove item'} title={isDe ? 'Punkt entfernen' : 'Remove item'}>
            <X size={14} />
          </button>
        </div>
        {(showDesc || isMobile) && (
          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '24px 1fr' : '28px 1fr 70px', gap: 6, margin: '2px 0 4px', padding: '0 4px' }}>
            <span />
            <div className="dex-ui-inline" style={{ gap: 6 }}>
              {isMobile && <input type="text" className="dex-ui-input dex-ui-input--sm" value={item.location || ''} onChange={e => patch(item.id, { location: e.target.value })} aria-label={isDe ? 'Raum' : 'Room'} placeholder={isDe ? 'Raum (optional)' : 'Room (optional)'} style={{ flex: '0 1 40%' }} />}
              {isMobile && groups.length > 1 && (
                <select value={g.key} onChange={e => moveItem(item.id, e.target.value)} title={moveTitle} aria-label={moveTitle} className="dex-ui-select dex-ui-select--sm" style={{ flex: '0 1 40%' }}>
                  {groups.map((x, i) => <option key={x.key} value={x.key}>{labelOf(x, i)}</option>)}
                </select>
              )}
              {showDesc && <input type="text" className="dex-ui-input dex-ui-input--sm" value={item.description || ''} onChange={e => patch(item.id, { description: e.target.value })} aria-label={isDe ? 'Beschreibung' : 'Description'} placeholder={isDe ? 'Beschreibung (optional) — sehen Teilnehmer unter dem Titel' : 'Description (optional) — shown to attendees under the title'} style={{ flex: '1 1 200px' }} />}
            </div>
          </div>
        )}
      </div>
    );
  };

  const renderHeadRow = (): React.ReactElement | null => isMobile ? null : (
    <div style={{ display: 'grid', gridTemplateColumns: cols, gap: 6, marginBottom: 2, padding: '0 4px' }}>
      <span /><span style={head}>Start</span><span style={head}>{isDe ? 'Ende' : 'End'}</span><span style={head}>{isDe ? 'Titel' : 'Title'}</span><span style={head}>{isDe ? 'Raum' : 'Room'}</span>{groups.length > 1 && <span style={head}>Cluster</span>}<span /><span />
    </div>
  );

  // v31.2: Der Knopf für den ersten Punkt bzw. den nächsten Cluster ist EIN
  // Handler an zwei Stellen: im leeren Zustand mitten im Kasten (dort sucht
  // ihn der Blick), sonst unter der Liste.
  const addGroupButton = (
    <button type="button" className="btn btn-outline dex-ui-btn-sm" onClick={addGroup}>
      <Plus size={14} /> {groups.length === 0 ? (isDe ? `Ersten ${termS} anlegen` : `Add first ${termS.toLowerCase()}`) : (isDe ? 'Weiterer Cluster' : 'Another cluster')}
    </button>
  );

  return (
    <div className="dex-ui-stack" style={{ gap: 12 }}>
      {/* v31.2: Leerer Zustand statt einsamem Knopf — sagt, wo Teilnehmer das Programm sehen und dass es Cluster gibt. */}
      {groups.length === 0 && (
        <div className="dex-ui-empty">
          <span className="dex-ui-empty-icon"><Calendar size={20} /></span>
          <div className="dex-ui-empty-title">{isDe ? `Noch keine ${termP}` : `No ${termP.toLowerCase()} yet`}</div>
          <div style={{ marginBottom: 14, lineHeight: 1.5 }}>
            {isDe
              ? <>Teilnehmer sehen das Programm auf der Anmeldeseite und unter &bdquo;Meine Events&ldquo;. Du ordnest die {termP} in Cluster — z.B. Tag 1, Vormittag oder Track A.</>
              : <>Attendees see the programme on the registration page and under &ldquo;My events&rdquo;. You group the {termP.toLowerCase()} into clusters — e.g. Day 1, Morning or Track A.</>}
          </div>
          {addGroupButton}
        </div>
      )}
      {groups.map((g, gi) => {
        const d = fromYmd(g.date);
        const nameValue = nameEdit && nameEdit.key === g.key ? nameEdit.value : g.cluster;
        const prefixed = countClusterPrefixed(g);
        const undated = g.dates.length === 0;
        const label = labelOf(g, gi);
        return (
          <div key={g.key} className="dex-ui-card" style={{ padding: 0, overflow: 'hidden', borderColor: undated ? 'var(--dex-orange, #ed8b00)' : undefined, borderStyle: undated ? 'dashed' : undefined }}>
            {/* v31.2: Kopf als weiche Karte; Name und Datum beschriftet, damit niemand
                raten muss, was das erste Feld ist. Aktionen rechts als Symbol-Knöpfe. */}
            <div className="dex-ui-card dex-ui-card--soft" style={{ border: 'none', borderRadius: 0, borderBottom: '1px solid var(--dex-gray-200)', padding: '8px 14px 10px', display: 'flex', alignItems: 'flex-end', gap: 10, flexWrap: 'wrap' }}>
              <span style={{ fontSize: '0.7rem', fontWeight: 700, width: 22, height: 22, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', borderRadius: '50%', background: 'var(--dex-green-dark, #4a7c1f)', color: '#fff', marginBottom: 5 }}>{gi + 1}</span>
              <div>
                <div style={head}>Cluster</div>
                <input
                  type="text" className="dex-ui-input dex-ui-input--sm" value={nameValue}
                  placeholder={suggestClusterName(gi, isDe)}
                  aria-label={isDe ? 'Cluster-Name' : 'Cluster name'}
                  title={isDe ? 'Cluster-Name — steht als Überschrift über diesen Punkten (z.B. Tag 1, Vormittag, Track A)' : 'Cluster name — heading above these items (e.g. Day 1, Morning, Track A)'}
                  onChange={e => setNameEdit({ key: g.key, value: e.target.value })}
                  onBlur={() => { if (nameEdit && nameEdit.key === g.key) { renameGroup(g, nameEdit.value); setNameEdit(null); } }}
                  onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); (e.target as HTMLInputElement).blur(); } }}
                  style={{ fontWeight: 700, width: 160 }}
                />
              </div>
              <div>
                <div style={head}>{isDe ? 'Datum' : 'Date'}</div>
                <div style={{ width: 128 }}>
                  <DatePicker
                    selected={d}
                    onChange={(nd: Date | null) => { if (nd) redateGroup(g, toYmd(nd)); }}
                    dateFormat="dd.MM.yyyy"
                    locale="de"
                    className="dex-ui-input dex-ui-input--sm"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    placeholderText="TT.MM.JJJJ"
                    autoComplete="off"
                    title={isDe ? 'Datum — verschiebt alle Punkte dieses Clusters' : 'Date — moves every item of this cluster'}
                  />
                </div>
              </div>
              <div className="dex-ui-inline" style={{ gap: 6, marginBottom: 4 }}>
                {undated
                  ? <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'ohne Datum' : 'no date'}</span>
                  : <span className="dex-ui-muted" style={{ minWidth: 70 }}>{weekdayOf(g.date, isDe)}{g.dates.length > 1 ? ` +${g.dates.length - 1}` : ''}</span>}
                <span className="dex-ui-pill dex-ui-pill--gray">{g.items.length} {g.items.length === 1 ? termS : termP}</span>
              </div>
              <span style={{ flex: 1 }} />
              <div className="dex-ui-inline" style={{ gap: 2, marginBottom: 1 }}>
                <button type="button" className="dex-ui-textbtn" onClick={() => addTo(g)} title={isDe ? `${termS} am Ende dieses Clusters anlegen` : `Add ${termS.toLowerCase()} at the end of this cluster`}><Plus size={14} /> {termS}</button>
                <button type="button" className="dex-ui-iconbtn" onClick={() => duplicateGroup(g, gi)} aria-label={isDe ? 'Cluster kopieren' : 'Copy cluster'} title={isDe ? 'Alle Punkte dieses Clusters auf den nächsten freien Tag kopieren' : 'Copy all items of this cluster to the next free day'}><Copy size={15} /></button>
                <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => setConfirmKey(g.key)} aria-label={isDe ? 'Cluster löschen' : 'Delete cluster'} title={isDe ? 'Diesen Cluster mit allen Punkten entfernen' : 'Remove this cluster with all its items'}><Trash2 size={15} /></button>
              </div>
            </div>
            {/* v31.2: Das Präfix-Angebot ist eine Frage an den Organizer, kein
                Knopf zwischen den Aktionen — deshalb eine eigene Hinweiszeile
                mit dem Grund und dem Knopf daneben. */}
            {prefixed > 0 && (
              <div className="dex-ui-callout dex-ui-callout--info dex-ui-callout--flush" style={{ alignItems: 'center', flexWrap: 'wrap' }}>
                <span className="dex-ui-callout-icon"><Info size={15} /></span>
                <span style={{ flex: 1, minWidth: 200 }}>{isDe ? `„${g.cluster} - “ steht noch in ${prefixed} Titeln — der Cluster-Name übernimmt das jetzt.` : `“${g.cluster} - ” still prefixes ${prefixed} titles — the cluster name now carries that.`}</span>
                <button type="button" className="dex-ui-textbtn" onClick={() => stripPrefixes(g)}>{isDe ? `Präfix aus ${prefixed} Titeln entfernen` : `Strip prefix from ${prefixed} titles`}</button>
              </div>
            )}
            {/* v31.2: Rückfrage nennt Cluster, Anzahl und Folge — statt
                „5 Punkte löschen?" neben dem X. */}
            {confirmKey === g.key && (
              <div className="dex-ui-callout dex-ui-callout--danger dex-ui-callout--flush" style={{ alignItems: 'center', flexWrap: 'wrap' }}>
                <span className="dex-ui-callout-icon"><AlertCircle size={15} /></span>
                <span style={{ flex: 1, minWidth: 200 }}>
                  <strong>{isDe ? `Cluster „${label}“ mit ${g.items.length} ${g.items.length === 1 ? termS : termP} löschen?` : `Delete cluster “${label}” with ${g.items.length} ${(g.items.length === 1 ? termS : termP).toLowerCase()}?`}</strong>{' '}
                  {isDe ? 'Die Punkte verschwinden aus dem Programm.' : 'The items disappear from the programme.'}
                </span>
                <button type="button" className="btn btn-danger dex-ui-btn-sm" onClick={() => deleteGroup(g)}>{isDe ? 'Ja, löschen' : 'Yes, delete'}</button>
                <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => setConfirmKey('')}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
              </div>
            )}
            <div style={{ padding: '6px 10px 8px' }}>
              {renderHeadRow()}
              {g.items.map(item => renderRow(item, g))}
            </div>
          </div>
        );
      })}

      {groups.length > 0 && (
        <div className="dex-ui-inline">
          {addGroupButton}
          <span className="dex-ui-muted" style={{ fontSize: '0.76rem' }}>
            {isDe
              ? `${sortedAll.length} ${termP} in ${groups.length} ${groups.length === 1 ? 'Cluster' : 'Clustern'} · Zeiten im 24-Stunden-Format · Enter im Titel legt den nächsten Punkt an.`
              : `${sortedAll.length} ${termP.toLowerCase()} in ${groups.length} ${groups.length === 1 ? 'cluster' : 'clusters'} · 24-hour times · Enter in the title adds the next item.`}
          </span>
        </div>
      )}
    </div>
  );
};

export default AgendaEditor;
