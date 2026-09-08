/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import Modal from './Modal';
import { SPRegistration } from '../services/EventService';
import { DexHotel } from '../types';
// v31.3: Klassen statt Inline-Styles — nur so gibt es Hover (docs/ui-leitfaden.md).
import { cx } from './dexUi';
import { AlertCircle, FileText } from './Icons';

/**
 * HotelImportModal (v28.49)
 * -------------------------
 * Übernimmt eine bestehende Hotel-Liste (Excel/CSV oder direkt aus Excel
 * kopiert) in die DEX-Hotelplanung.
 *
 * Der Ablauf ist bewusst zweistufig: erst **erkennen und zeigen**, dann auf
 * Knopfdruck schreiben. Eine Import-Funktion, die sofort losschreibt, ist bei
 * fremden Listen gefährlich — Tippfehler in E-Mail-Adressen, vertauschte
 * Datumsformate und unbekannte Hotelnamen fallen sonst erst hinterher auf.
 * Deshalb zeigt die Vorschau je Zeile, was passieren wird, und was nicht
 * zugeordnet werden konnte.
 *
 * v31.3: Der Ablauf ist jetzt auch sichtbar dreistufig — „Liste einlesen",
 * „Spalten prüfen", „Vorschau"; Übernehmen ist der einzige Primär-Knopf und
 * sitzt in der Modal-Fußzeile. Erkennung und Zuordnung sind unverändert.
 *
 * Erkannt wird:
 *  - **Spalten** über die Kopfzeile (deutsche und englische Schreibweisen)
 *    und, falls die Kopfzeile nichts hergibt, über den Inhalt (eine Spalte mit
 *    '@' ist die E-Mail, Datumsspalten über parsebare Werte).
 *  - **Personen** über die E-Mail-Adresse (Groß-/Kleinschreibung egal).
 *  - **Hotels** über den Namen, unscharf verglichen (Groß-/Kleinschreibung,
 *    Leerzeichen, Sonderzeichen) — sonst legt jede Schreibweise ein neues
 *    Hotel an.
 *  - **Daten** in den gängigen Formaten (TT.MM.JJJJ, JJJJ-MM-TT, TT/MM/JJJJ)
 *    sowie als Excel-Seriennummer.
 */

export interface IHotelImportResultRow {
  email: string;
  hotel: string;
  from: string;
  to: string;
  /** Teilnehmerzeile, falls über die E-Mail gefunden. */
  reg: SPRegistration | null;
  /** Grund, warum die Zeile nicht übernommen wird (leer = wird übernommen). */
  problem: string;
}

export interface IHotelImportModalProps {
  open: boolean;
  onClose: () => void;
  isDe: boolean;
  people: SPRegistration[];
  hotels: DexHotel[];
  /** Übernimmt die geprüften Zeilen; legt fehlende Hotels vorher an. */
  onApply: (rows: IHotelImportResultRow[], newHotelNames: string[]) => Promise<void>;
  busy: boolean;
  progress: { done: number; total: number } | null;
}

/** Normalisiert für den unscharfen Vergleich von Hotelnamen/Spaltenköpfen. */
const norm = (v: string): string =>
  (v || '').toLowerCase().replace(/[\s\-_.]+/g, '').replace(/[^a-z0-9äöüß@]/g, '');

/** Erkennt die gängigen Datumsformate + Excel-Seriennummer → 'YYYY-MM-DD'. */
const parseDay = (raw: any): string => {
  if (raw === null || raw === undefined) return '';
  // Excel liefert bei Datums-Zellen je nach Einstellung ein Date-Objekt …
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return new Date(Date.UTC(raw.getFullYear(), raw.getMonth(), raw.getDate()))
      .toISOString().substring(0, 10);
  }
  // … oder die Seriennummer seit 1899-12-30.
  if (typeof raw === 'number' && raw > 20000 && raw < 60000) {
    const ms = Math.round((raw - 25569) * 86400000);
    return new Date(ms).toISOString().substring(0, 10);
  }
  const s = String(raw).trim();
  if (!s) return '';
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = /^(\d{1,2})[./](\d{1,2})[./](\d{2,4})/.exec(s);
  if (m) {
    let y = m[3];
    if (y.length === 2) y = `20${y}`;
    return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  }
  return '';
};

const nights = (from: string, to: string): number => {
  if (!from || !to) return 0;
  const a = Date.parse(`${from}T00:00:00Z`);
  const b = Date.parse(`${to}T00:00:00Z`);
  if (isNaN(a) || isNaN(b) || b <= a) return 0;
  return Math.round((b - a) / 86400000);
};

/** v31.3: Initialen für die Personen-Zelle — Import-Zeilen haben kein Foto. */
const initialsOf = (v: string): string => {
  const parts = (v || '').split(/[\s@._-]+/).filter(Boolean);
  if (parts.length === 0) return '?';
  return (parts[0][0] + (parts.length > 1 ? parts[1][0] : '')).toUpperCase();
};

/** Kopfzeilen-Muster je Zielspalte, deutsch + englisch. */
const HEAD: Record<'email' | 'hotel' | 'from' | 'to', string[]> = {
  email: ['email', 'emailadresse', 'mail', 'mailadresse', 'emailaddress', 'teilnehmeremail', 'participantemail', 'upn'],
  hotel: ['hotel', 'hotelname', 'unterkunft', 'accommodation', 'lodging', 'haus'],
  from: ['anreise', 'von', 'checkin', 'arrival', 'from', 'anreisedatum', 'startdatum', 'start'],
  to: ['abreise', 'bis', 'checkout', 'departure', 'to', 'abreisedatum', 'enddatum', 'ende', 'end'],
};

export const HotelImportModal: React.FC<IHotelImportModalProps> = (props: IHotelImportModalProps) => {
  const { open, onClose, isDe, people, hotels, onApply, busy, progress } = props;

  const [raw, setRaw] = React.useState('');
  const [grid, setGrid] = React.useState<any[][]>([]);
  const [fileName, setFileName] = React.useState('');
  const [parseError, setParseError] = React.useState('');
  /** Spalten-Zuordnung; -1 = nicht vorhanden. Nach der Erkennung änderbar. */
  const [map, setMap] = React.useState<{ email: number; hotel: number; from: number; to: number }>(
    { email: -1, hotel: -1, from: -1, to: -1 },
  );
  const [hasHeader, setHasHeader] = React.useState(true);

  const reset = (): void => {
    setRaw(''); setGrid([]); setFileName(''); setParseError('');
    setMap({ email: -1, hotel: -1, from: -1, to: -1 }); setHasHeader(true);
  };

  React.useEffect(() => { if (open) reset(); }, [open]);

  /** Spalten aus der Kopfzeile erkennen, sonst aus dem Inhalt raten. */
  const autoMap = (rows: any[][]): { email: number; hotel: number; from: number; to: number; header: boolean } => {
    const res = { email: -1, hotel: -1, from: -1, to: -1, header: false };
    if (rows.length === 0) return res;
    const head = rows[0].map(c => norm(String(c ?? '')));
    (Object.keys(HEAD) as Array<keyof typeof HEAD>).forEach(key => {
      const idx = head.findIndex(h => h && HEAD[key].some(p => h === p || h.indexOf(p) >= 0));
      if (idx >= 0) res[key] = idx;
    });
    res.header = res.email >= 0 || res.hotel >= 0 || res.from >= 0 || res.to >= 0;
    // Kein brauchbarer Kopf → über den Inhalt raten (erste Datenzeilen).
    const body = res.header ? rows.slice(1) : rows;
    const colCount = Math.max(...rows.map(r => r.length));
    if (res.email < 0) {
      for (let c = 0; c < colCount; c++) {
        if (body.some(r => String(r[c] ?? '').indexOf('@') > 0)) { res.email = c; break; }
      }
    }
    if (res.from < 0 || res.to < 0) {
      const dateCols: number[] = [];
      for (let c = 0; c < colCount; c++) {
        if (c === res.email) continue;
        const hits = body.filter(r => parseDay(r[c])).length;
        if (hits > 0 && hits >= Math.max(1, Math.floor(body.length * 0.5))) dateCols.push(c);
      }
      if (res.from < 0 && dateCols.length > 0) res.from = dateCols[0];
      if (res.to < 0 && dateCols.length > 1) res.to = dateCols[1];
    }
    if (res.hotel < 0) {
      for (let c = 0; c < colCount; c++) {
        if (c === res.email || c === res.from || c === res.to) continue;
        // Textspalte mit wenigen verschiedenen Werten = sehr wahrscheinlich das Hotel.
        const vals = body.map(r => String(r[c] ?? '').trim()).filter(Boolean);
        if (vals.length === 0) continue;
        const uniq = new Set(vals.map(norm));
        if (uniq.size > 0 && uniq.size <= Math.max(6, vals.length / 3) && !vals.some(v => v.indexOf('@') > 0)) {
          res.hotel = c; break;
        }
      }
    }
    return res;
  };

  const applyGrid = (rows: any[][]): void => {
    const cleaned = rows.filter(r => r.some(c => String(c ?? '').trim() !== ''));
    setGrid(cleaned);
    const m = autoMap(cleaned);
    setMap({ email: m.email, hotel: m.hotel, from: m.from, to: m.to });
    setHasHeader(m.header);
  };

  /** Aus Excel kopierter Text bzw. CSV — Trennzeichen automatisch erkennen. */
  const parseText = (text: string): void => {
    setParseError('');
    const lines = text.replace(/\r\n?/g, '\n').split('\n').filter(l => l.trim() !== '');
    if (lines.length === 0) { setGrid([]); return; }
    const first = lines[0];
    const delim = first.indexOf('\t') >= 0 ? '\t'
      : (first.split(';').length > first.split(',').length ? ';' : ',');
    const rows = lines.map(l => l.split(delim).map(c => c.trim().replace(/^"(.*)"$/, '$1')));
    applyGrid(rows);
  };

  const onFile = async (file: File): Promise<void> => {
    setParseError('');
    setFileName(file.name);
    const lower = file.name.toLowerCase();
    try {
      if (/\.(xlsx|xlsm|xls)$/.test(lower)) {
        // Bundle-Regel: xlsx MUSS dynamisch importiert werden.
        const XLSX = await import('xlsx');
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }) as any[][];
        applyGrid(rows);
      } else {
        const text = await file.text();
        setRaw(text);
        parseText(text);
      }
    } catch (err) {
      setParseError(isDe
        ? `Die Datei konnte nicht gelesen werden (${err instanceof Error ? err.message : 'unbekannter Fehler'}). Alternativ: Inhalt in Excel markieren, kopieren und unten einfügen.`
        : `Could not read the file (${err instanceof Error ? err.message : 'unknown error'}). Alternative: select the content in Excel, copy it and paste below.`);
    }
  };

  /* ---------------- Vorschau berechnen ---------------- */

  const byEmail = React.useMemo(() => {
    const m: Record<string, SPRegistration> = {};
    for (const p of people) {
      const e = (p.ParticipantEmail || '').trim().toLowerCase();
      if (e) m[e] = p;
    }
    return m;
  }, [people]);

  const hotelByNorm = React.useMemo(() => {
    const m: Record<string, DexHotel> = {};
    for (const h of hotels) m[norm(h.name)] = h;
    return m;
  }, [hotels]);

  const rows: IHotelImportResultRow[] = React.useMemo(() => {
    const body = hasHeader ? grid.slice(1) : grid;
    const out: IHotelImportResultRow[] = [];
    const seen = new Set<string>();
    for (const r of body) {
      const email = map.email >= 0 ? String(r[map.email] ?? '').trim().toLowerCase() : '';
      const hotelRaw = map.hotel >= 0 ? String(r[map.hotel] ?? '').trim() : '';
      const from = map.from >= 0 ? parseDay(r[map.from]) : '';
      const to = map.to >= 0 ? parseDay(r[map.to]) : '';
      const reg = email ? (byEmail[email] || null) : null;
      // Bekannte Schreibweise eines vorhandenen Hotels auf dessen Namen ziehen,
      // damit nicht „Marriott" und „marriott " als zwei Hotels enden.
      const known = hotelRaw ? hotelByNorm[norm(hotelRaw)] : undefined;
      const hotel = known ? known.name : hotelRaw;
      let problem = '';
      if (!email) problem = isDe ? 'keine E-Mail in der Zeile' : 'no email in row';
      else if (!reg) problem = isDe ? 'keine aktive Anmeldung zu dieser E-Mail' : 'no active registration for this email';
      else if (seen.has(email)) problem = isDe ? 'E-Mail kommt mehrfach vor' : 'duplicate email';
      else if (!hotel) problem = isDe ? 'kein Hotel in der Zeile' : 'no hotel in row';
      else if (from && to && nights(from, to) <= 0) problem = isDe ? 'Abreise liegt nicht nach der Anreise' : 'departure not after arrival';
      if (!problem && email) seen.add(email);
      out.push({ email, hotel, from, to, reg, problem });
    }
    return out;
  }, [grid, hasHeader, map, byEmail, hotelByNorm, isDe]);

  const okRows = rows.filter(r => !r.problem);
  const badRows = rows.filter(r => r.problem);
  const newHotels = React.useMemo(() => {
    const set: string[] = [];
    for (const r of okRows) {
      if (!r.hotel) continue;
      if (hotelByNorm[norm(r.hotel)]) continue;
      if (!set.some(x => norm(x) === norm(r.hotel))) set.push(r.hotel);
    }
    return set;
  }, [okRows, hotelByNorm]);
  const noDates = okRows.filter(r => !r.from || !r.to).length;

  const colOptions = React.useMemo(() => {
    const n = grid.length > 0 ? Math.max(...grid.map(r => r.length)) : 0;
    return Array.from({ length: n }, (_, i) => i);
  }, [grid]);
  const colLabel = (i: number): string => {
    const head = hasHeader && grid.length > 0 ? String(grid[0][i] ?? '').trim() : '';
    return head ? `${i + 1}: ${head}` : `${isDe ? 'Spalte' : 'Column'} ${i + 1}`;
  };

  // v31.3: Die Frage statt des Feldnamens — der Organizer beantwortet hier,
  // welche Spalte seiner Liste wohin gehört. `req` markiert die beiden Spalten,
  // ohne die keine Zeile übernommen wird (die Prüfung bleibt in `rows`).
  const colFields: Array<{ k: 'email' | 'hotel' | 'from' | 'to'; l: string; req: boolean }> = [
    { k: 'email', l: isDe ? 'Wo steht die E-Mail?' : 'Where is the email?', req: true },
    { k: 'hotel', l: isDe ? 'Wo steht das Hotel?' : 'Where is the hotel?', req: true },
    { k: 'from', l: isDe ? 'Wo steht die Anreise?' : 'Where is the arrival?', req: false },
    { k: 'to', l: isDe ? 'Wo steht die Abreise?' : 'Where is the departure?', req: false },
  ];

  return (
    <Modal
      open={open} onClose={onClose} dismissable={!busy} maxWidth={720}
      ariaLabel={isDe ? 'Hotelliste importieren' : 'Import hotel list'}
      title={isDe ? 'Hotelliste importieren' : 'Import hotel list'}
      subtitle={isDe
        ? 'Wir lesen deine Liste, ordnen die Personen über die E-Mail-Adresse zu und zeigen dir Zeile für Zeile, was passiert. Geschrieben wird nichts, bevor du unten bestätigst.'
        : 'We read your list, match people by email address and show you row by row what will happen. Nothing is written before you confirm below.'}
      icon={<FileText size={20} />}
      // v31.3: Übernehmen ist der einzige Primär-Knopf und steht in der
      // Modal-Fußzeile — vorher stand die Knopfzeile frei im Inhalt und der
      // Grund für einen gesperrten Knopf stand nirgends.
      footer={<>
        {grid.length > 0 && okRows.length === 0 && (
          <span className="dex-ui-modal-foot-left dex-ui-muted">
            {isDe ? 'Noch ist keine Zeile übernehmbar — prüfe die Spalten oben.' : 'No row can be applied yet — check the columns above.'}
          </span>
        )}
        <button type="button" className="btn btn-secondary" onClick={onClose} disabled={busy}>
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        <button
          type="button"
          className="btn btn-primary"
          disabled={busy || okRows.length === 0}
          onClick={() => { void onApply(okRows, newHotels); }}
        >
          {busy
            ? (isDe ? 'Wird übernommen…' : 'Applying…')
            : (isDe
              ? `${okRows.length} ${okRows.length === 1 ? 'Zuordnung' : 'Zuordnungen'} übernehmen`
              : `Apply ${okRows.length} assignment${okRows.length === 1 ? '' : 's'}`)}
        </button>
      </>}
    >
      {/* ---- Schritt 1: Quelle ---- */}
      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{isDe ? 'Schritt 1 — Liste einlesen' : 'Step 1 — Load the list'}</div>
        <label className="dex-ui-dropzone" style={busy ? { cursor: 'wait' } : undefined}>
          <span className="dex-ui-choice-icon" aria-hidden="true"><FileText size={18} /></span>
          <span style={{ fontWeight: 700 }}>{isDe ? 'Datei wählen — Excel oder CSV' : 'Choose a file — Excel or CSV'}</span>
          <span className="dex-ui-muted">
            {fileName || (isDe ? '.xlsx, .xls, .csv oder .txt — die Spalten erkennt die App selbst.' : '.xlsx, .xls, .csv or .txt — the app detects the columns itself.')}
          </span>
          <input type="file" accept=".xlsx,.xlsm,.xls,.csv,.txt" style={{ display: 'none' }} disabled={busy}
            onChange={e => { const f = e.target.files && e.target.files[0]; if (f) void onFile(f); e.target.value = ''; }} />
        </label>

        <div className="dex-ui-field" style={{ marginTop: 14 }}>
          <label className="dex-ui-label" htmlFor="dexHotelImportPaste">
            {isDe ? 'Oder direkt aus Excel einfügen' : 'Or paste straight from Excel'}
          </label>
          <textarea
            id="dexHotelImportPaste"
            className="dex-ui-textarea"
            value={raw}
            disabled={busy}
            onChange={e => { setRaw(e.target.value); setFileName(''); parseText(e.target.value); }}
            placeholder={isDe ? 'E-Mail | Hotel | Anreise | Abreise' : 'Email | Hotel | Arrival | Departure'}
            rows={4}
            style={{ fontFamily: 'monospace', fontSize: '0.8rem' }}
          />
          <div className="dex-ui-help">
            {isDe ? 'Zeilen in Excel markieren, kopieren, hier einfügen — am besten mit Kopfzeile.' : 'Select the rows in Excel, copy them and paste them here — ideally with a header row.'}
          </div>
        </div>
      </div>

      {parseError && (
        <div className="dex-ui-callout dex-ui-callout--danger">
          <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>
          <div>{parseError}</div>
        </div>
      )}

      {/* ---- Schritt 2: Spalten-Zuordnung ---- */}
      {grid.length > 0 && (
        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{isDe ? 'Schritt 2 — Spalten prüfen' : 'Step 2 — Check the columns'}</div>
          <div className="dex-ui-section-desc">
            {isDe ? 'Das haben wir automatisch erkannt. Stimmt eine Spalte nicht, stell sie hier um.' : 'This is what we detected automatically. If a column is wrong, change it here.'}
          </div>
          <label className={cx('dex-ui-toggle-row', hasHeader && 'is-active')}>
            <input type="checkbox" checked={hasHeader} onChange={e => setHasHeader(e.target.checked)} />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">{isDe ? 'Die erste Zeile ist eine Kopfzeile' : 'The first row is a header'}</span>
              <span className="dex-ui-toggle-row-desc">
                {isDe ? 'Dann wird sie nicht übernommen, und ihre Texte stehen unten in der Spaltenauswahl.' : 'Then it is not imported, and its texts appear in the column pickers below.'}
              </span>
            </span>
          </label>
          <div className="dex-ui-grid-2" style={{ marginTop: 12 }}>
            {colFields.map(c => (
              <div className="dex-ui-field" key={c.k}>
                <label className="dex-ui-label" htmlFor={`dexHotelImportCol-${c.k}`}>
                  {c.l}
                  {c.req
                    ? <span className="dex-ui-label-required">*</span>
                    : <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>}
                </label>
                <select id={`dexHotelImportCol-${c.k}`} className="dex-ui-select dex-ui-select--sm"
                  value={map[c.k]} disabled={busy}
                  onChange={e => setMap(m => ({ ...m, [c.k]: parseInt(e.target.value, 10) }))}>
                  <option value={-1}>{isDe ? '— nicht vorhanden —' : '— not present —'}</option>
                  {colOptions.map(i => <option key={i} value={i}>{colLabel(i)}</option>)}
                </select>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ---- Schritt 3: Vorschau ---- */}
      {grid.length > 0 && (
        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{isDe ? 'Schritt 3 — Vorschau' : 'Step 3 — Preview'}</div>
          <div className="dex-ui-inline" style={{ marginBottom: 10 }}>
            <span className="dex-ui-pill dex-ui-pill--green">{okRows.length} {isDe ? 'werden übernommen' : 'will be applied'}</span>
            {badRows.length > 0 && (
              <span className="dex-ui-pill dex-ui-pill--orange">{badRows.length} {isDe ? 'übersprungen' : 'skipped'}</span>
            )}
            {newHotels.length > 0 && (
              <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'Hotels neu angelegt: ' : 'hotels newly created: '}<strong>{newHotels.join(', ')}</strong></span>
            )}
          </div>

          {/* v31.3: Was Aufmerksamkeit braucht, steht als Warnung über der
              Tabelle — als graue Pille neben den Zählern ging es unter. */}
          {(badRows.length > 0 || noDates > 0) && (
            <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 10 }}>
              <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>
              <div>
                {badRows.length > 0 && (
                  <div>{isDe ? `${badRows.length} Zeilen werden übersprungen — den Grund nennt die Spalte Status.` : `${badRows.length} rows are skipped — the status column names the reason.`}</div>
                )}
                {noDates > 0 && (
                  <div>{isDe ? `${noDates} ohne Zeitraum — Hotel wird gesetzt, Datum bleibt leer.` : `${noDates} without dates — hotel set, dates left empty.`}</div>
                )}
              </div>
            </div>
          )}

          <div className="dex-ui-table-wrap dex-ui-table-wrap--sticky" style={{ maxHeight: 300 }}>
            <table className="dex-ui-table dex-ui-table--compact">
              <thead>
                <tr>
                  <th>{isDe ? 'Person' : 'Person'}</th>
                  <th>Hotel</th>
                  <th>{isDe ? 'Zeitraum' : 'Stay'}</th>
                  <th>{isDe ? 'Status' : 'Status'}</th>
                </tr>
              </thead>
              <tbody>
                {rows.slice(0, 300).map((r, i) => {
                  // Ohne Treffer in der Teilnehmerliste bleibt die E-Mail der
                  // Name — die Zeile sagt so trotzdem, um wen es geht.
                  const who = r.reg ? (r.reg.ParticipantName || r.reg.ParticipantEmail) : (r.email || '—');
                  return (
                    <tr key={i} className={cx(r.problem && 'is-muted')}>
                      <td>
                        <span className="dex-ui-person">
                          <span className="dex-ui-avatar" aria-hidden="true">{initialsOf(who)}</span>
                          <span style={{ minWidth: 0 }}>
                            <span className="dex-ui-person-name" style={{ display: 'block' }}>{who}</span>
                            {r.reg && r.reg.ParticipantName && r.email
                              && <span className="dex-ui-person-sub" style={{ display: 'block' }}>{r.email}</span>}
                          </span>
                        </span>
                      </td>
                      <td>{r.hotel || '—'}</td>
                      <td style={{ whiteSpace: 'nowrap' }}>
                        {r.from && r.to ? `${r.from} – ${r.to} · ${nights(r.from, r.to)} ${isDe ? 'N.' : 'n.'}` : '—'}
                      </td>
                      <td>
                        <span className={cx('dex-ui-pill', r.problem ? 'dex-ui-pill--orange' : 'dex-ui-pill--green')}>
                          {r.problem || (isDe ? 'wird übernommen' : 'will be applied')}
                        </span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div className="dex-ui-table-foot">
            <span>{isDe ? `${rows.length} Zeilen erkannt` : `${rows.length} rows detected`}</span>
            {rows.length > 300 && (
              <span>{isDe ? 'Vorschau zeigt die ersten 300 — übernommen werden alle.' : 'Preview shows the first 300 — all of them are applied.'}</span>
            )}
          </div>
        </div>
      )}

      {progress && (
        <div>
          <div className="dex-ui-progress">
            <div className="dex-ui-progress-bar"
              style={{ width: `${progress.total > 0 ? Math.round((progress.done / progress.total) * 100) : 0}%` }} />
          </div>
          <div className="dex-ui-help">
            {progress.done}/{progress.total} {isDe ? 'übernommen …' : 'applied …'}
          </div>
        </div>
      )}
    </Modal>
  );
};

export default HotelImportModal;
