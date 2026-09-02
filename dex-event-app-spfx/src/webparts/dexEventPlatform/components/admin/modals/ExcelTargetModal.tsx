/* ExcelTargetModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 15129-15241 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

export interface ExcelTargetModalProps {
  consolidatedChildren: DeloitteEvent[];
  excelAudience: "active" | "activePlusWait" | "waitOnly" | "withCancelled";
  excelIncludeMatrix: boolean;
  excelSubIds: Set<string>;
  excelTargetModal: { mode: "deloitte" | "b2run"; chooseMode?: boolean; };
  exportConsolidatedExcel: (audience: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled', includeMatrix: boolean, subIds: string[]) => void;
  exportCsv: (mode: 'deloitte' | 'b2run', audience?: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled') => void;
  isConsolidatedMode: boolean;
  isDe: boolean;
  selectedEvent: DeloitteEvent;
  setExcelAudience: React.Dispatch<React.SetStateAction<"active" | "activePlusWait" | "waitOnly" | "withCancelled">>;
  setExcelIncludeMatrix: React.Dispatch<React.SetStateAction<boolean>>;
  setExcelSubIds: React.Dispatch<React.SetStateAction<Set<string>>>;
  setExcelTargetModal: React.Dispatch<React.SetStateAction<{ mode: "deloitte" | "b2run"; chooseMode?: boolean; }>>;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
}

export const ExcelTargetModal: React.FC<ExcelTargetModalProps> = (p) => {
  const { consolidatedChildren, excelAudience, excelIncludeMatrix, excelSubIds, excelTargetModal, exportConsolidatedExcel, exportCsv, isConsolidatedMode, isDe, selectedEvent, setExcelAudience, setExcelIncludeMatrix, setExcelSubIds, setExcelTargetModal, subEventRegsByEventId } = p;
        const closeAll = (): void => setExcelTargetModal(null);
        // v20.4: Im Klammer-Modus entscheidet das Modal, WAS exportiert wird —
        // konsolidierte Matrix und/oder einzelne Sub-Event-Blätter.
        const consolidatedExportPossible = isConsolidatedMode && excelTargetModal.mode === 'deloitte' && consolidatedChildren.length > 0;
        const proceed = (): void => {
          const mode = excelTargetModal.mode;
          setExcelTargetModal(null);
          if (consolidatedExportPossible && (excelIncludeMatrix || excelSubIds.size > 0)) {
            exportConsolidatedExcel(excelAudience, excelIncludeMatrix, Array.from(excelSubIds));
          } else {
            exportCsv(mode, excelAudience);
          }
        };
        const toggleSubId = (id: string): void => {
          setExcelSubIds(prev => {
            const next = new Set(prev);
            if (next.has(id)) next.delete(id); else next.add(id);
            return next;
          });
        };
        const Row = (props: { value: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled'; label: string; desc: string }): React.ReactElement => (
          <label style={{
            display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
            borderRadius: 8, border: `1px solid ${excelAudience === props.value ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
            background: excelAudience === props.value ? 'rgba(134,188,37,0.08)' : '#fff',
            cursor: 'pointer', marginBottom: 8,
          }}>
            <input type="radio" name="excel-target" checked={excelAudience === props.value} onChange={() => setExcelAudience(props.value)} style={{ marginTop: 3 }} />
            <div>
              <div style={{ fontWeight: 600, fontSize: '0.92rem' }}>{props.label}</div>
              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>{props.desc}</div>
            </div>
          </label>
        );
        return (
          <Modal open={true} onClose={closeAll} maxWidth={520} padding={24} ariaLabel="Excel-Export Zielgruppe">
            <h3 style={{ margin: '0 0 14px', fontSize: '1.1rem' }}>Excel-Export</h3>
            {/* v27.9: Format-Auswahl (Deloitte/B2Run) direkt im Modal — vorher
                im Anker-Dropdown, das vom „Aktion auswählen"-Menü abgeschnitten
                wurde. */}
            {excelTargetModal.chooseMode && (
              <div style={{ marginBottom: 14, paddingBottom: 12, borderBottom: '1px solid var(--dex-gray-200)' }}>
                <div style={{ fontWeight: 700, fontSize: '0.92rem', marginBottom: 8 }}>{isDe ? 'Welches Format?' : 'Which format?'}</div>
                {([
                  { m: 'b2run' as const, label: 'B2Run View', desc: isDe ? 'Genau das Veranstalter-Format: EIN Arbeitsblatt, 16 Spalten (Nr., Anrede, Name, E-Mail, Startblock, …) — direkt bei b2run.com importierbar.' : 'Exact organizer format: ONE worksheet, 16 columns — importable at b2run.com.' },
                  { m: 'deloitte' as const, label: isDe ? 'Deloitte Felder' : 'Deloitte fields', desc: isDe ? 'Alle internen Spalten + alle Custom-Fields des Events (für intern).' : 'All internal columns + all event custom fields (internal use).' },
                ]).map(opt => (
                  <label key={opt.m} style={{ display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10, borderRadius: 8, border: `1px solid ${excelTargetModal.mode === opt.m ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`, background: excelTargetModal.mode === opt.m ? 'rgba(134,188,37,0.08)' : '#fff', cursor: 'pointer', marginBottom: 8 }}>
                    <input type="radio" name="excel-mode" checked={excelTargetModal.mode === opt.m} onChange={() => setExcelTargetModal({ mode: opt.m, chooseMode: true })} style={{ marginTop: 3 }} />
                    <div>
                      <div style={{ fontWeight: 600, fontSize: '0.92rem' }}>{opt.label}</div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>{opt.desc}</div>
                    </div>
                  </label>
                ))}
              </div>
            )}
            <div style={{ fontWeight: 700, fontSize: '0.92rem', marginBottom: 8 }}>{isDe ? 'Wen sollen wir exportieren?' : 'Who should we export?'}</div>
            <Row value="active" label="Teilnehmer (alle aktiven)" desc="Status: Angemeldet, QR versendet, Eingecheckt — Default für den Check-In / die Vor-Ort-Liste." />
            <Row value="activePlusWait" label="Teilnehmer + Warteliste" desc="Alle aktiven + Wartelistler in einem Sheet, sortiert nach TeilnehmerID." />
            <Row value="waitOnly" label="Nur Warteliste" desc="Nur die Wartelistler — z.B. für Briefing." />
            <Row value="withCancelled" label="Alles inkl. Abmeldungen" desc="Alle Einträge inklusive abgemeldeter Personen — der Status steht pro Zeile in der Status-Spalte." />
            {/* v20.4: Klammer-Modus — wählen, was in die Datei kommt. */}
            {consolidatedExportPossible && (
              <div style={{ marginTop: 14, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
                <div style={{ fontWeight: 700, fontSize: '0.92rem', marginBottom: 8 }}>
                  {isDe ? 'Was soll in die Datei?' : 'What goes into the file?'}
                </div>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, padding: '8px 10px', borderRadius: 8, border: `1px solid ${excelIncludeMatrix ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`, background: excelIncludeMatrix ? 'rgba(134,188,37,0.08)' : '#fff', cursor: 'pointer', marginBottom: 8 }}>
                  <input type="checkbox" checked={excelIncludeMatrix} onChange={e => setExcelIncludeMatrix(e.target.checked)} style={{ marginTop: 3 }} />
                  <span>
                    <span style={{ display: 'block', fontWeight: 600, fontSize: '0.9rem' }}>{isDe ? 'Konsolidierte Matrix' : 'Consolidated matrix'}</span>
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                      {isDe ? 'Eine Zeile pro Person — mit den übergreifenden Feldern und pro Sub-Event dem Status + den Sub-Event-Antworten (wie die Tabelle in der Klammer-Ansicht).' : 'One row per person — with the cross-cutting fields and per sub-event the status + the sub-event answers (like the table in the consolidated view).'}
                    </span>
                  </span>
                </label>
                <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)', margin: '6px 0 6px' }}>
                  {isDe ? 'Zusätzlich einzelne Sub-Event-Blätter:' : 'Additionally individual sub-event sheets:'}
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 180, overflowY: 'auto' }}>
                  {consolidatedChildren.map(child => {
                    const checked = excelSubIds.has(child.id);
                    const short = shortSubEventTitle(child.title, selectedEvent.title) || child.title || '?';
                    return (
                      <label key={child.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '6px 10px', borderRadius: 8, border: `1px solid ${checked ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`, background: checked ? 'rgba(134,188,37,0.08)' : '#fff', cursor: 'pointer' }}>
                        <input type="checkbox" checked={checked} onChange={() => toggleSubId(child.id)} />
                        <span style={{ fontSize: '0.88rem' }}>{short}</span>
                        <span style={{ marginLeft: 'auto', fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                          {(subEventRegsByEventId[child.id] || []).length} {isDe ? 'Einträge' : 'entries'}
                        </span>
                      </label>
                    );
                  })}
                </div>
                <div style={{ display: 'flex', gap: 12, marginTop: 6 }}>
                  <button type="button" onClick={() => setExcelSubIds(new Set(consolidatedChildren.map(c => c.id)))} style={{ background: 'none', border: 'none', color: 'var(--dex-green-dark)', fontSize: '0.78rem', fontWeight: 600, cursor: 'pointer', padding: 0 }}>
                    {isDe ? 'Alle auswählen' : 'Select all'}
                  </button>
                  <button type="button" onClick={() => setExcelSubIds(new Set())} style={{ background: 'none', border: 'none', color: 'var(--dex-gray-500)', fontSize: '0.78rem', fontWeight: 600, cursor: 'pointer', padding: 0 }}>
                    {isDe ? 'Keine' : 'None'}
                  </button>
                </div>
              </div>
            )}
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 14 }}>
              <button type="button" className="btn btn-secondary" onClick={closeAll} style={{ fontSize: '0.85rem' }}>Abbrechen</button>
              <button type="button" className="btn btn-primary" onClick={proceed} style={{ fontSize: '0.85rem' }}>Excel herunterladen</button>
            </div>
          </Modal>
        );
};

