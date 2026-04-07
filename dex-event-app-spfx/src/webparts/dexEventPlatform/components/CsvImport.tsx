/**
 * CSV Event Import - importiert Events aus einer CSV-Datei in DEX_Events.
 *
 * Nur fuer Admins sichtbar.
 * Mappt CSV-Spalten auf DEX_Events Felder inkl. CustomFields.
 */

import * as React from 'react';
import { EventService, CustomField } from '../services/EventService';

interface CsvEvent {
  title: string;
  status: string;
  type: string;
  description: string;
  location: string;
  locationFilter: string;
  audience: string;
  startDate: string;
  endDate: string;
  registrationDeadline: string;
  maxParticipants: number;
  organizer: string;
  outlookEventId: string;
  customFields: CustomField[];
  // Rohdaten fuer Anzeige
  raw: Record<string, string>;
}

function parseCsvDate(val: string): string {
  if (!val) return '';
  // US-Format: M/D/YYYY H:MM
  const match = val.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})$/);
  if (match) {
    const [, m, d, y, h, min] = match;
    return new Date(Number(y), Number(m) - 1, Number(d), Number(h), Number(min)).toISOString();
  }
  // Fallback
  const dt = new Date(val);
  return isNaN(dt.getTime()) ? '' : dt.toISOString();
}

function mapEventType(contentType: string): string {
  if (contentType.toLowerCase().includes('b2run')) return 'B2Run';
  if (contentType.toLowerCase().includes('morgan')) return 'JPMorgan';
  return 'Other';
}

function mapStatus(status: string): string {
  if (status === 'Active') return 'Active';
  if (status === 'Under Construction') return 'Under Construction';
  if (status === 'Out dated') return 'Completed';
  if (status === 'Fictive') return 'Cancelled';
  return 'Completed';
}

function parseCsv(text: string): Record<string, string>[] {
  // BOM entfernen
  const clean = text.replace(/^\uFEFF/, '');
  const lines = clean.split('\n');
  if (lines.length < 2) return [];

  // Header parsen
  const headers = parseCsvLine(lines[0]);
  const rows: Record<string, string>[] = [];

  let i = 1;
  while (i < lines.length) {
    let line = lines[i];
    // Multiline-Felder: wenn Anzahl Anführungszeichen ungerade, nächste Zeile anhängen
    while ((line.match(/"/g) || []).length % 2 !== 0 && i + 1 < lines.length) {
      i++;
      line += '\n' + lines[i];
    }
    if (line.trim()) {
      const values = parseCsvLine(line);
      const row: Record<string, string> = {};
      headers.forEach((h, idx) => { row[h] = (values[idx] || '').trim(); });
      rows.push(row);
    }
    i++;
  }
  return rows;
}

function parseCsvLine(line: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && i + 1 < line.length && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  result.push(current);
  return result;
}

function mapCsvToEvent(row: Record<string, string>): CsvEvent {
  const customFields: CustomField[] = [];

  // T-Shirt
  if (row['T Shirt Größe visible'] === 'True' && row['T-Shirtgrößen']) {
    try {
      const options = JSON.parse(row['T-Shirtgrößen']);
      customFields.push({ id: 'tshirt', label: 'T-Shirt Größe', type: 'select', required: false, options, visible: true });
    } catch { /* */ }
  }

  // Startblock
  if (row['Startblock']) {
    try {
      const options = JSON.parse(row['Startblock']);
      if (options.length > 0) {
        customFields.push({ id: 'startblock', label: 'Startblock', type: 'select', required: false, options, visible: true });
      }
    } catch { /* */ }
  }

  // Allergien
  if (row['Allergien visible'] === 'True') {
    customFields.push({ id: 'allergien', label: 'Allergien / Unverträglichkeiten', type: 'text', required: false, visible: true });
  }

  // Freitextfelder
  if (row['Frage Freitextfeld 1']) {
    customFields.push({ id: 'freitext1', label: row['Frage Freitextfeld 1'], type: 'text', required: false, visible: true });
  }
  if (row['Frage Freitextfeld 2']) {
    customFields.push({ id: 'freitext2', label: row['Frage Freitextfeld 2'], type: 'text', required: false, visible: true });
  }

  // Dropdownfelder
  if (row['Frage Dropdownfeld 1'] && row['Dropdownfeld 1']) {
    try {
      const options = JSON.parse(row['Dropdownfeld 1']);
      customFields.push({ id: 'dropdown1', label: row['Frage Dropdownfeld 1'], type: 'select', required: false, options, visible: true });
    } catch { /* */ }
  }
  if (row['Frage Dropdownfeld 2'] && row['Dropdownfeld 2']) {
    try {
      const options = JSON.parse(row['Dropdownfeld 2']);
      customFields.push({ id: 'dropdown2', label: row['Frage Dropdownfeld 2'], type: 'select', required: false, options, visible: true });
    } catch { /* */ }
  }

  return {
    title: row['Title'] || '',
    status: mapStatus(row['Status'] || ''),
    type: mapEventType(row['Content Type'] || ''),
    description: row['Event description'] || '',
    location: row['Standort'] || '',
    locationFilter: row['Standort Filter'] || '',
    audience: row['Audience'] || '',
    startDate: parseCsvDate(row['Event time'] || ''),
    endDate: parseCsvDate(row['Event end time'] || ''),
    registrationDeadline: parseCsvDate(row['Last Register Date Time'] || ''),
    maxParticipants: parseInt((row['Max. parcipients'] || '0').replace(/,/g, ''), 10) || 0,
    organizer: (row['Organzier'] || '').replace(/;/g, ', '),
    outlookEventId: row['Outlook EventID'] || '',
    customFields,
    raw: row,
  };
}

interface Props {
  siteUrl: string;
}

export default function CsvImport({ siteUrl: _siteUrl }: Props): React.ReactElement {
  const [isOpen, setIsOpen] = React.useState(false);
  const [csvEvents, setCsvEvents] = React.useState<CsvEvent[]>([]);
  const [selected, setSelected] = React.useState<Set<number>>(new Set());
  const [isImporting, setIsImporting] = React.useState(false);
  const [importLog, setImportLog] = React.useState<string[]>([]);
  const [fileName, setFileName] = React.useState('');

  function handleFileUpload(e: React.ChangeEvent<HTMLInputElement>): void {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (ev) => {
      const text = ev.target?.result as string;
      const rows = parseCsv(text);
      const mapped = rows.map(mapCsvToEvent).filter(e => e.title);
      setCsvEvents(mapped);
      setSelected(new Set(mapped.map((_, i) => i)));
      setImportLog([]);
    };
    reader.readAsText(file, 'utf-8');
  }

  async function handleImport(): Promise<void> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    if (!ctx) return;
    const svc = new EventService(ctx);
    setIsImporting(true);
    const log: string[] = [];

    const toImport = csvEvents.filter((_, i) => selected.has(i));
    for (let i = 0; i < toImport.length; i++) {
      const ev = toImport[i];
      log.push(`[${i + 1}/${toImport.length}] Importiere: ${ev.title}...`);
      setImportLog([...log]);

      try {
        const eventId = await svc.createEvent({
          title: ev.title,
          type: ev.type,
          status: ev.status,
          description: ev.description,
          location: ev.location,
          locationFilter: ev.locationFilter,
          audience: ev.audience,
          filterMode: 'OR',
          startDate: ev.startDate,
          endDate: ev.endDate,
          registrationDeadline: ev.registrationDeadline,
          lastDeregisterDate: '',
          maxParticipants: ev.maxParticipants,
          waitlistEnabled: true,
          eventImageUrl: '',
          organizer: ev.organizer,
          organizerEmail: '',
          outlookEventId: ev.outlookEventId, // CalendarLink uebernehmen fuer Outlook-Einladungen bei Anmeldung
          outlookBody: '',
          emailLanguage: 'EN',
          customFields: ev.customFields,
        });
        if (eventId) {
          log[log.length - 1] += ` ✓ (ID: ${eventId})`;
        } else {
          log[log.length - 1] += ' ✗ (Fehler beim Erstellen)';
        }
      } catch (err) {
        log[log.length - 1] += ` ✗ (${err})`;
      }
      setImportLog([...log]);
    }

    log.push(`\nFertig! ${toImport.length} Events verarbeitet.`);
    setImportLog([...log]);
    setIsImporting(false);
  }

  return (
    <div className="card" style={{ marginTop: 16 }}>
      <h3
        style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 8, margin: 0 }}
        onClick={() => setIsOpen(!isOpen)}
      >
        📥 CSV Event-Import {isOpen ? '▾' : '▸'}
      </h3>

      {isOpen && (
        <div style={{ marginTop: 16 }}>
          <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 12 }}>
            Importiere Events aus einer CSV-Datei. Die CSV muss die Spalten aus der alten Event-Liste enthalten
            (Title, Status, Content Type, Standort, Event time, Event end time, Max. parcipients, Organzier, Event description).
          </p>

          <input
            type="file"
            accept=".csv"
            onChange={handleFileUpload}
            style={{ marginBottom: 12 }}
          />
          {fileName && <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginLeft: 8 }}>{fileName}</span>}

          {csvEvents.length > 0 && (
            <>
              <div style={{ marginTop: 12, marginBottom: 8, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 600 }}>{csvEvents.length} Events gefunden, {selected.size} ausgewählt</span>
                <div style={{ display: 'flex', gap: 4 }}>
                  <button className="btn btn-secondary" style={{ fontSize: '0.7rem', padding: '2px 8px' }}
                    onClick={() => setSelected(new Set(csvEvents.map((_, i) => i)))}>Alle</button>
                  <button className="btn btn-secondary" style={{ fontSize: '0.7rem', padding: '2px 8px' }}
                    onClick={() => setSelected(new Set())}>Keine</button>
                </div>
              </div>

              <div style={{ maxHeight: 400, overflowY: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 8 }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.75rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--dex-gray-200)', position: 'sticky', top: 0, background: '#fff' }}>
                      <th style={{ padding: 6, width: 30 }} />
                      <th style={{ textAlign: 'left', padding: 6 }}>Titel</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>Typ</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>Status</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>Ort</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>Start</th>
                      <th style={{ textAlign: 'right', padding: 6 }}>Max</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>Organizer</th>
                      <th style={{ textAlign: 'right', padding: 6 }}>Felder</th>
                    </tr>
                  </thead>
                  <tbody>
                    {csvEvents.map((ev, i) => (
                      <tr key={i} style={{ borderBottom: '1px solid var(--dex-gray-100)', opacity: selected.has(i) ? 1 : 0.4 }}>
                        <td style={{ padding: 6, textAlign: 'center' }}>
                          <input type="checkbox" checked={selected.has(i)} onChange={() => {
                            const next = new Set(selected);
                            if (next.has(i)) next.delete(i); else next.add(i);
                            setSelected(next);
                          }} />
                        </td>
                        <td style={{ padding: 6, fontWeight: 500 }}>{ev.title}</td>
                        <td style={{ padding: 6 }}>{ev.type}</td>
                        <td style={{ padding: 6 }}>
                          <span style={{
                            padding: '1px 6px', borderRadius: 4, fontSize: '0.7rem',
                            background: ev.status === 'Active' ? '#e8f5e9' : ev.status === 'Completed' ? '#e3f2fd' : '#fff3e0',
                            color: ev.status === 'Active' ? '#2e7d32' : ev.status === 'Completed' ? '#1565c0' : '#e65100',
                          }}>
                            {ev.status}
                          </span>
                        </td>
                        <td style={{ padding: 6 }}>{ev.location}</td>
                        <td style={{ padding: 6 }}>{ev.startDate ? new Date(ev.startDate).toLocaleDateString('de-DE') : '-'}</td>
                        <td style={{ padding: 6, textAlign: 'right' }}>{ev.maxParticipants}</td>
                        <td style={{ padding: 6, maxWidth: 150, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{ev.organizer}</td>
                        <td style={{ padding: 6, textAlign: 'right' }}>{ev.customFields.length}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div style={{ marginTop: 12, display: 'flex', gap: 8 }}>
                <button
                  className="btn btn-primary"
                  disabled={isImporting || selected.size === 0}
                  onClick={handleImport}
                >
                  {isImporting ? 'Importiere...' : `${selected.size} Events importieren`}
                </button>
              </div>
            </>
          )}

          {importLog.length > 0 && (
            <div style={{
              marginTop: 12, background: '#1e1e1e', color: '#0f0', padding: 12,
              borderRadius: 8, fontSize: '0.7rem', fontFamily: 'monospace',
              maxHeight: 300, overflowY: 'auto', whiteSpace: 'pre-wrap',
            }}>
              {importLog.join('\n')}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
