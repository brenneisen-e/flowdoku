/**
 * v30.48: Import der offiziellen B2Run-Startnummern.
 *
 * Zwei Schritte, bewusst getrennt: erst **prüfen und anzeigen**, dann auf
 * Knopfdruck schreiben. Der Grund ist der Nachrücker-Fall — dort wandert eine
 * Startnummer von einer Person zur anderen, und das muss der Organizer sehen
 * BEVOR es in den Daten steht, weil er dieselbe Ummeldung anschließend beim
 * Veranstalter machen muss. Ein Import, der sofort schreibt, würde ihm genau
 * die Liste vorenthalten, für die er den Import gemacht hat.
 *
 * Die Zuordnung „wer ersetzt wen" wird nicht geraten: Sie steht seit v17.15 in
 * `ReplacedByParticipantEmail` / `ReplacedParticipantEmail` auf den Zeilen.
 * Siehe `utils/b2runBibImport.ts`.
 */
import * as React from 'react';
import Modal from '../Modal';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { useCurrentUser } from '../../context/UserContext';
import { DeloitteEvent } from '../../types';
import { SPRegistration, EventService } from '../../services/EventService';
import { parseBibSheet, buildBibReport, BibImportReport, BibMatch } from '../../utils/b2runBibImport';

const nameOf = (r?: SPRegistration): string =>
  r ? `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || (r.ParticipantEmail || '') : '—';

const box: React.CSSProperties = {
  border: '1px solid var(--dex-gray-200)', borderRadius: 10,
  padding: '12px 14px', marginBottom: 12,
};

export default function B2RunBibImportModal(props: {
  event: DeloitteEvent;
  service: EventService;
  onClose: () => void;
  onDone: () => void;
}): React.ReactElement {
  const { getAllRegistrations } = useEvents();
  const { showAlert } = useDialog();
  const { currentUser } = useCurrentUser();
  const [busy, setBusy] = React.useState(false);
  const [progress, setProgress] = React.useState('');
  const [report, setReport] = React.useState<BibImportReport | null>(null);
  const [fileName, setFileName] = React.useState('');
  const [written, setWritten] = React.useState<{ ok: number; failed: number } | null>(null);

  const readFile = async (file: File): Promise<void> => {
    setBusy(true); setProgress('Datei wird gelesen…'); setReport(null); setWritten(null);
    try {
      const buf = await file.arrayBuffer();
      // Bundle-Regel: xlsx MUSS dynamisch importiert werden (s. HotelImportModal).
      const XLSX = await import('xlsx');
      const wb = XLSX.read(buf, { type: 'array', cellDates: false });
      const ws = wb.Sheets[wb.SheetNames[0]];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const table = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' }) as any[][];
      const parsed = parseBibSheet(table);
      if (parsed.error) { await showAlert(parsed.error, { variant: 'error' }); return; }
      setProgress('Teilnehmer werden geladen…');
      const regs = await getAllRegistrations(props.event.id);
      if (regs.length === 0) {
        await showAlert('Die Teilnehmerliste dieses Events ist leer oder nicht lesbar — ohne sie kann nichts zugeordnet werden.', { variant: 'error' });
        return;
      }
      setFileName(file.name);
      setReport(buildBibReport(parsed.rows, regs));
    } catch (err) {
      console.warn('[DEX] B2Run-Import failed:', err);
      await showAlert('Die Datei konnte nicht gelesen werden. Bitte die unveränderte Excel des Veranstalters verwenden.', { variant: 'error' });
    } finally { setBusy(false); setProgress(''); }
  };

  const write = async (): Promise<void> => {
    if (!report || !props.event.subsiteUrl) return;
    const toWrite = report.matches.filter(m => (m.kind === 'direct' || m.kind === 'transfer') && m.target && m.row.bib);
    setBusy(true);
    // Die Spalte gibt es in Bestands-Listen nicht — erst anlegen, sonst laufen
    // alle Schreibvorgänge auf einen Feldfehler und der Import meldet 0/300.
    setProgress('Spalte „Startnummer" wird geprüft…');
    const colOk = await props.service.ensureStartNumberColumn(props.event.subsiteUrl);
    if (!colOk) {
      setBusy(false); setProgress('');
      await showAlert('Die Spalte „Startnummer" konnte in der Teilnehmerliste nicht angelegt werden. Ohne sie kann nichts geschrieben werden.', { variant: 'error' });
      return;
    }
    let ok = 0; let failed = 0;
    // Sequentiell: Die Teilnehmerliste laeuft mit Item-Level-Security, und ein
    // Promise.all ueber 300 Zeilen ist die 429-Welle, vor der CLAUDE.md warnt.
    for (let i = 0; i < toWrite.length; i++) {
      const m = toWrite[i];
      setProgress(`Startnummern werden geschrieben… ${i + 1} / ${toWrite.length}`);
      try {
        await props.service.adminUpdateRegistration(
          props.event.subsiteUrl, m.target!.Id,
          { Startnummer: m.row.bib },
          { name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email, email: currentUser.email },
        );
        ok++;
      } catch { failed++; }
    }
    setBusy(false); setProgress('');
    setWritten({ ok, failed });
    props.onDone();
  };

  const group = (kind: BibMatch['kind']): BibMatch[] => (report ? report.matches.filter(m => m.kind === kind) : []);
  const transfers = group('transfer');
  const orphans = group('orphan');
  const unknowns = group('unknown');
  const directs = group('direct');
  const mismatches = directs.filter(m => m.blockMismatch);

  return (
    <Modal open onClose={props.onClose} maxWidth={860} ariaLabel="Startnummern importieren (B2Run)">
      <div style={{ fontSize: '0.87rem', lineHeight: 1.55 }}>
        <h3 style={{ margin: '0 0 10px', fontSize: '1.05rem' }}>Startnummern importieren (B2Run)</h3>
        {!report && (
          <>
            <p style={{ marginTop: 0 }}>
              Lade die Rücklauf-Datei des Veranstalters hoch — dieselbe Liste, die du gemeldet hast,
              ergänzt um die Spalte <strong>Startnummer</strong>. DEX ordnet die Nummern über die
              E-Mail-Adresse zu und zeigt dir vor dem Schreiben, was passieren würde.
            </p>
            <p style={{ color: 'var(--dex-gray-600)' }}>
              Besonders wichtig: Wer sich nach der Meldung abgemeldet hat, taucht in der Datei noch auf.
              DEX sagt dir, <strong>wer für diese Person nachgerückt ist</strong> und die Nummer übernimmt —
              genau die Ummeldungen, die du beim Veranstalter noch machen musst.
            </p>
            <label className="btn btn-primary" style={{ cursor: busy ? 'wait' : 'pointer', display: 'inline-block' }}>
              {busy ? (progress || 'Bitte warten…') : '+ Datei wählen (.xlsx / .csv)'}
              <input
                type="file"
                accept=".xlsx,.xlsm,.xls,.csv"
                style={{ display: 'none' }}
                disabled={busy}
                onChange={e => { const f = e.target.files?.[0]; if (f) void readFile(f); e.currentTarget.value = ''; }}
              />
            </label>
          </>
        )}

        {report && (
          <>
            <p style={{ marginTop: 0, color: 'var(--dex-gray-600)' }}>
              Datei: <strong>{fileName}</strong> · {report.matches.length} Zeilen gelesen
            </p>

            {transfers.length > 0 && (
              <div style={{ ...box, borderColor: 'var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.07)' }}>
                <strong style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>
                  Beim Veranstalter ummelden ({transfers.length})
                </strong>
                <p style={{ margin: '4px 0 8px', fontSize: '0.82rem' }}>
                  Diese Personen stehen auf der Liste, haben sich aber abgemeldet. Die Startnummer geht
                  an die Person, die nachgerückt ist — in DEX wird sie beim Schreiben übertragen, beim
                  Veranstalter musst du sie noch ummelden.
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  {transfers.map(m => (
                    <div key={m.row.rowNo} style={{ fontSize: '0.82rem' }}>
                      <span style={{ fontFamily: 'monospace', fontWeight: 700 }}>{m.row.bib}</span>
                      {'  '}
                      <span style={{ textDecoration: 'line-through', color: 'var(--dex-gray-500)' }}>{nameOf(m.listed)}</span>
                      {' → '}
                      <strong>{nameOf(m.target)}</strong>
                      <span style={{ color: 'var(--dex-gray-500)' }}> · {m.target?.ParticipantEmail}</span>
                      {m.chain && m.chain.length > 1 && (
                        <span style={{ color: 'var(--dex-gray-500)' }}> (über {m.chain.length - 1} weitere Abmeldung{m.chain.length - 1 === 1 ? '' : 'en'})</span>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {orphans.length > 0 && (
              <div style={{ ...box, borderColor: 'var(--dex-red, #da291c)', background: 'rgba(218,41,28,0.06)' }}>
                <strong style={{ color: 'var(--dex-red, #da291c)' }}>Nummer ohne Nachrücker ({orphans.length})</strong>
                <p style={{ margin: '4px 0 8px', fontSize: '0.82rem' }}>
                  Abgemeldet, und niemand ist nachgerückt. Diese Startnummern bleiben ungenutzt — beim
                  Veranstalter abmelden oder jemanden nachmelden.
                </p>
                {orphans.map(m => (
                  <div key={m.row.rowNo} style={{ fontSize: '0.82rem' }}>
                    <span style={{ fontFamily: 'monospace', fontWeight: 700 }}>{m.row.bib}</span>{'  '}{nameOf(m.listed)}
                  </div>
                ))}
              </div>
            )}

            {report.missingFromFile.length > 0 && (
              <div style={{ ...box, borderColor: 'var(--dex-orange, #ed8b00)' }}>
                <strong>Angemeldet, aber ohne Startnummer ({report.missingFromFile.length})</strong>
                <p style={{ margin: '4px 0 8px', fontSize: '0.82rem' }}>
                  In DEX angemeldet, steht aber nicht in der Datei — hier fehlt die Meldung beim Veranstalter.
                </p>
                {report.missingFromFile.map(r => (
                  <div key={r.Id} style={{ fontSize: '0.82rem' }}>{nameOf(r)} · {r.ParticipantEmail}</div>
                ))}
              </div>
            )}

            {unknowns.length > 0 && (
              <div style={box}>
                <strong>In DEX unbekannt ({unknowns.length})</strong>
                <p style={{ margin: '4px 0 8px', fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                  Diese Adressen stehen in der Datei, aber in keiner DEX-Zeile. Häufigster Grund: eine
                  zweite Schreibweise derselben Person (SMTP-Adresse gegen Alias).
                </p>
                {unknowns.map(m => (
                  <div key={m.row.rowNo} style={{ fontSize: '0.82rem' }}>
                    <span style={{ fontFamily: 'monospace' }}>{m.row.bib}</span>{'  '}
                    {m.row.firstName} {m.row.lastName} · {m.row.email || '(keine E-Mail)'}
                  </div>
                ))}
              </div>
            )}

            {mismatches.length > 0 && (
              <div style={box}>
                <strong>Abweichender Startblock ({mismatches.length})</strong>
                <p style={{ margin: '4px 0 8px', fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                  Nur ein Hinweis — die Startnummer wird trotzdem zugeordnet.
                </p>
                {mismatches.map(m => (
                  <div key={m.row.rowNo} style={{ fontSize: '0.82rem' }}>
                    {nameOf(m.listed)}: DEX &bdquo;{m.blockMismatch!.dex}&ldquo; · Liste &bdquo;{m.blockMismatch!.file}&ldquo;
                  </div>
                ))}
              </div>
            )}

            {report.duplicateBibs.length > 0 && (
              <div style={{ ...box, borderColor: 'var(--dex-red, #da291c)' }}>
                <strong style={{ color: 'var(--dex-red, #da291c)' }}>Doppelte Startnummern in der Datei</strong>
                <div style={{ fontSize: '0.82rem' }}>{report.duplicateBibs.join(', ')}</div>
              </div>
            )}

            <div style={{ ...box, borderColor: 'var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.07)' }}>
              <strong>Wird geschrieben: {directs.length + transfers.length} Startnummern</strong>
              <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                {directs.length} direkt zugeordnet, {transfers.length} an Nachrücker übertragen.
              </div>
            </div>

            {written && (
              <p style={{ color: written.failed ? 'var(--dex-red)' : 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>
                {written.ok} Startnummer(n) geschrieben{written.failed ? `, ${written.failed} fehlgeschlagen` : ''}.
              </p>
            )}

            <div style={{ display: 'flex', gap: 10, marginTop: 14, flexWrap: 'wrap' }}>
              <button
                className="btn btn-primary"
                disabled={busy || !!written}
                onClick={() => { void write(); }}
              >
                {busy ? (progress || 'Bitte warten…') : 'Startnummern jetzt schreiben'}
              </button>
              <button className="btn btn-secondary" disabled={busy} onClick={props.onClose}>
                {written ? 'Schließen' : 'Abbrechen'}
              </button>
            </div>
          </>
        )}
      </div>
    </Modal>
  );
}
