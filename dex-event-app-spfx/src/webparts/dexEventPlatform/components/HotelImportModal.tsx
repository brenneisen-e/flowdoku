/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import Modal from './Modal';
import { SPRegistration } from '../services/EventService';
import { DexHotel } from '../types';

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

  const inp: React.CSSProperties = {
    height: 30, fontSize: '0.8rem', padding: '0 8px',
    border: '1px solid var(--dex-gray-300)', borderRadius: 6,
  };
  const colOptions = React.useMemo(() => {
    const n = grid.length > 0 ? Math.max(...grid.map(r => r.length)) : 0;
    return Array.from({ length: n }, (_, i) => i);
  }, [grid]);
  const colLabel = (i: number): string => {
    const head = hasHeader && grid.length > 0 ? String(grid[0][i] ?? '').trim() : '';
    return head ? `${i + 1}: ${head}` : `${isDe ? 'Spalte' : 'Column'} ${i + 1}`;
  };

  return (
    <Modal open={open} onClose={onClose} dismissable={!busy} ariaLabel={isDe ? 'Hotelliste importieren' : 'Import hotel list'}>
      <h2 style={{ margin: 0, fontSize: '1.25rem', color: 'var(--dex-gray-800)' }}>
        {isDe ? 'Hotelliste importieren' : 'Import hotel list'}
      </h2>
      <p style={{ margin: '4px 0 0', fontSize: '0.84rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
        {isDe
          ? 'Lade deine bestehende Liste hoch (Excel oder CSV) oder kopiere sie direkt aus Excel hier hinein. Die App erkennt die Spalten, ordnet die Personen über die E-Mail-Adresse zu und legt fehlende Hotels an. Nichts wird geschrieben, bevor du die Vorschau bestätigst.'
          : 'Upload your existing list (Excel or CSV) or paste it straight from Excel. The app detects the columns, matches people by email address and creates missing hotels. Nothing is written before you confirm the preview.'}
      </p>

      {/* ---- Quelle ---- */}
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center', marginTop: 14 }}>
        <label style={{
          display: 'inline-flex', alignItems: 'center', gap: 8, padding: '8px 14px',
          border: '2px dashed var(--dex-gray-300)', borderRadius: 8, cursor: busy ? 'wait' : 'pointer',
          fontSize: '0.84rem', color: 'var(--dex-gray-700)',
        }}>
          {isDe ? '+ Datei wählen (.xlsx / .csv)' : '+ Choose file (.xlsx / .csv)'}
          <input type="file" accept=".xlsx,.xlsm,.xls,.csv,.txt" style={{ display: 'none' }} disabled={busy}
            onChange={e => { const f = e.target.files && e.target.files[0]; if (f) void onFile(f); e.target.value = ''; }} />
        </label>
        {fileName && <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>{fileName}</span>}
      </div>

      <textarea
        value={raw}
        disabled={busy}
        onChange={e => { setRaw(e.target.value); setFileName(''); parseText(e.target.value); }}
        placeholder={isDe
          ? '… oder hier aus Excel einfügen (mit Kopfzeile, z.B.: E-Mail | Hotel | Anreise | Abreise)'
          : '… or paste from Excel here (with a header row, e.g.: Email | Hotel | Arrival | Departure)'}
        rows={4}
        style={{ width: '100%', marginTop: 10, fontSize: '0.8rem', padding: 8, border: '1px solid var(--dex-gray-300)', borderRadius: 6, fontFamily: 'monospace' }}
      />

      {parseError && (
        <div style={{ marginTop: 10, padding: '8px 10px', borderRadius: 6, background: '#fef3f2', border: '1px solid var(--dex-red, #c00)', color: '#7a1f1c', fontSize: '0.8rem' }}>
          {parseError}
        </div>
      )}

      {/* ---- Spalten-Zuordnung ---- */}
      {grid.length > 0 && (
        <div style={{ marginTop: 14, padding: 12, borderRadius: 8, background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)' }}>
          <div style={{ fontSize: '0.8rem', fontWeight: 700, marginBottom: 8, color: 'var(--dex-gray-700)' }}>
            {isDe ? 'Erkannte Spalten — bei Bedarf korrigieren' : 'Detected columns — correct if needed'}
          </div>
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.78rem', marginBottom: 8, cursor: 'pointer' }}>
            <input type="checkbox" checked={hasHeader} onChange={e => setHasHeader(e.target.checked)} />
            {isDe ? 'Erste Zeile ist eine Kopfzeile' : 'First row is a header'}
          </label>
          <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
            {([
              { k: 'email' as const, l: isDe ? 'E-Mail *' : 'Email *' },
              { k: 'hotel' as const, l: 'Hotel *' },
              { k: 'from' as const, l: isDe ? 'Anreise' : 'Arrival' },
              { k: 'to' as const, l: isDe ? 'Abreise' : 'Departure' },
            ]).map(c => (
              <label key={c.k} style={{ display: 'flex', flexDirection: 'column', gap: 3, fontSize: '0.74rem', color: 'var(--dex-gray-600)' }}>
                {c.l}
                <select style={inp} value={map[c.k]} disabled={busy}
                  onChange={e => setMap(m => ({ ...m, [c.k]: parseInt(e.target.value, 10) }))}>
                  <option value={-1}>{isDe ? '— nicht vorhanden —' : '— not present —'}</option>
                  {colOptions.map(i => <option key={i} value={i}>{colLabel(i)}</option>)}
                </select>
              </label>
            ))}
          </div>
        </div>
      )}

      {/* ---- Vorschau ---- */}
      {grid.length > 0 && (
        <div style={{ marginTop: 14 }}>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 8 }}>
            <span style={{ padding: '6px 12px', borderRadius: 8, background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.8rem', fontWeight: 700 }}>
              {okRows.length} {isDe ? 'werden übernommen' : 'will be applied'}
            </span>
            {badRows.length > 0 && (
              <span style={{ padding: '6px 12px', borderRadius: 8, background: '#fff6e5', color: '#b35a00', fontSize: '0.8rem', fontWeight: 700 }}>
                {badRows.length} {isDe ? 'übersprungen' : 'skipped'}
              </span>
            )}
            {newHotels.length > 0 && (
              <span style={{ padding: '6px 12px', borderRadius: 8, background: 'var(--dex-gray-100)', fontSize: '0.8rem' }}>
                {isDe ? 'neu angelegt: ' : 'newly created: '}<strong>{newHotels.join(', ')}</strong>
              </span>
            )}
            {noDates > 0 && (
              <span style={{ padding: '6px 12px', borderRadius: 8, background: 'var(--dex-gray-100)', fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
                {isDe ? `${noDates} ohne Zeitraum — Hotel wird gesetzt, Datum bleibt leer` : `${noDates} without dates — hotel set, dates left empty`}
              </span>
            )}
          </div>

          <div style={{ maxHeight: 260, overflowY: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 8 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.76rem' }}>
              <thead>
                <tr style={{ textAlign: 'left', color: 'var(--dex-gray-600)', background: 'var(--dex-gray-50, #f7f7f5)' }}>
                  <th style={{ padding: '5px 8px' }}>{isDe ? 'Person' : 'Person'}</th>
                  <th style={{ padding: '5px 8px' }}>Hotel</th>
                  <th style={{ padding: '5px 8px' }}>{isDe ? 'Zeitraum' : 'Stay'}</th>
                  <th style={{ padding: '5px 8px' }}>{isDe ? 'Status' : 'Status'}</th>
                </tr>
              </thead>
              <tbody>
                {rows.slice(0, 300).map((r, i) => (
                  <tr key={i} style={{ borderTop: '1px solid var(--dex-gray-100)', background: r.problem ? '#fffaf2' : undefined }}>
                    <td style={{ padding: '5px 8px' }}>
                      {r.reg ? (r.reg.ParticipantName || r.reg.ParticipantEmail) : (r.email || '—')}
                    </td>
                    <td style={{ padding: '5px 8px' }}>{r.hotel || '—'}</td>
                    <td style={{ padding: '5px 8px', whiteSpace: 'nowrap' }}>
                      {r.from && r.to ? `${r.from} – ${r.to} · ${nights(r.from, r.to)} ${isDe ? 'N.' : 'n.'}` : '—'}
                    </td>
                    <td style={{ padding: '5px 8px', color: r.problem ? '#b35a00' : 'var(--dex-green-dark, #4a7c1f)' }}>
                      {r.problem || (isDe ? 'wird übernommen' : 'will be applied')}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {rows.length > 300 && (
            <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 6 }}>
              {isDe ? `Vorschau zeigt die ersten 300 von ${rows.length} Zeilen — übernommen werden alle.` : `Preview shows the first 300 of ${rows.length} rows — all of them are applied.`}
            </div>
          )}
        </div>
      )}

      {progress && (
        <div style={{ marginTop: 12 }}>
          <div style={{ height: 8, background: 'var(--dex-gray-100)', borderRadius: 999, overflow: 'hidden' }}>
            <div style={{
              height: '100%', transition: 'width 0.2s', background: 'var(--dex-green, #86bc25)',
              width: `${progress.total > 0 ? Math.round((progress.done / progress.total) * 100) : 0}%`,
            }} />
          </div>
          <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
            {progress.done}/{progress.total} {isDe ? 'übernommen …' : 'applied …'}
          </div>
        </div>
      )}

      <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 16 }}>
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
            : (isDe ? `${okRows.length} Zuordnung(en) übernehmen` : `Apply ${okRows.length} assignment(s)`)}
        </button>
      </div>
    </Modal>
  );
};

export default HotelImportModal;
