/* AttachmentsModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16192-16326 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Download, FileText, Plus, X } from '../../Icons';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';

export interface AttachmentsModalProps {
  attachmentsBusy: boolean;
  attachmentsByReg: Record<number, { fileName: string; serverRelativeUrl: string; }[]>;
  attachmentsModalReg: SPRegistration;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  eventServiceRef: EventService;
  isDe: boolean;
  selectedEvent: DeloitteEvent;
  setAttachmentsBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setAttachmentsByReg: React.Dispatch<React.SetStateAction<Record<number, { fileName: string; serverRelativeUrl: string; }[]>>>;
  setAttachmentsModalReg: React.Dispatch<React.SetStateAction<SPRegistration>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export const AttachmentsModal: React.FC<AttachmentsModalProps> = (p) => {
  const { attachmentsBusy, attachmentsByReg, attachmentsModalReg, confirmDialog, eventServiceRef, isDe, selectedEvent, setAttachmentsBusy, setAttachmentsByReg, setAttachmentsModalReg, showAlert } = p;
        const reg = attachmentsModalReg;
        const list = attachmentsByReg[reg.Id] || [];
        const close = (): void => setAttachmentsModalReg(null);
        // v19.0: Dokument-Feld-Attachments tragen einen `dxf-<fieldId>--`-Präfix.
        // Für die Anzeige den Präfix + Timestamp strippen und das Feld-Label
        // ermitteln, damit der Organizer sieht, zu welchem Dokument-Feld die
        // Datei gehört.
        const docFields = (selectedEvent?.eventSpecificFields || []).filter(f => f.type === 'document');
        const fieldLabelForFile = (fileName: string): string => {
          const m = fileName.match(/^dxf-([a-zA-Z0-9]+)--/);
          if (!m) return '';
          const df = docFields.find(f => (f.id || '').replace(/[^a-zA-Z0-9]/g, '') === m[1]);
          return df ? df.label : '';
        };
        const prettyFileName = (fileName: string): string =>
          fileName
            .replace(/^dxf-[a-zA-Z0-9]+--\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}_/, '')
            .replace(/^dxf-[a-zA-Z0-9]+--/, '')
            .replace(/^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}_/, '');
        const refreshOne = async (regId: number): Promise<void> => {
          if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
          try {
            const fresh = await eventServiceRef.listRegistrationAttachments(selectedEvent.subsiteUrl, regId);
            setAttachmentsByReg(prev => ({ ...prev, [regId]: fresh }));
          } catch { /* */ }
        };
        const onDelete = async (fileName: string): Promise<void> => {
          if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
          if (!(await confirmDialog(isDe ? `Datei „${fileName}" wirklich löschen?` : `Really delete file „${fileName}"?`, { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' }))) return;
          setAttachmentsBusy(true);
          try {
            await eventServiceRef.deleteRegistrationAttachment(selectedEvent.subsiteUrl, reg.Id, fileName);
            await refreshOne(reg.Id);
          } finally { setAttachmentsBusy(false); }
        };
        const onAdd = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
          const f = e.target.files && e.target.files[0];
          e.target.value = '';
          if (!f || !eventServiceRef || !selectedEvent?.subsiteUrl) return;
          if (f.size > 10 * 1024 * 1024) {
            showAlert(isDe ? 'Datei ist größer als 10 MB.' : 'File is larger than 10 MB.');
            return;
          }
          setAttachmentsBusy(true);
          try {
            await eventServiceRef.addRegistrationAttachment(selectedEvent.subsiteUrl, reg.Id, f);
            await refreshOne(reg.Id);
          } finally { setAttachmentsBusy(false); }
        };
        const fullName = `${reg.Vorname || ''} ${reg.Nachname || ''}`.trim() || reg.ParticipantEmail || '–';
        return (
          <Modal
            open={true}
            onClose={close}
            dismissable={!attachmentsBusy}
            maxWidth={560}
            padding={24}
            ariaLabel={isDe ? 'Hochgeladene Dateien' : 'Uploaded files'}
          >
              <h3 style={{ margin: '0 0 4px', fontSize: '1.1rem' }}>
                {isDe ? 'Hochgeladene Dateien' : 'Uploaded files'}
              </h3>
              <p style={{ margin: '0 0 14px', fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                {fullName}{reg.ParticipantEmail ? ` · ${reg.ParticipantEmail}` : ''}
              </p>
              {list.length === 0 ? (
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', fontStyle: 'italic', margin: '12px 0' }}>
                  {isDe ? 'Noch keine Dateien hochgeladen.' : 'No files uploaded yet.'}
                </p>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 16 }}>
                  {list.map(f => (
                    <div key={f.fileName} style={{
                      display: 'flex', alignItems: 'center', gap: 10,
                      padding: '8px 12px', borderRadius: 6,
                      background: 'rgba(134,188,37,0.08)',
                      border: '1px solid rgba(134,188,37,0.30)',
                      fontSize: '0.85rem',
                    }}>
                      <FileText size={16} />
                      <a
                        href={f.serverRelativeUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{ flex: 1, color: 'var(--dex-gray-800)', textDecoration: 'none', wordBreak: 'break-all' }}
                      >
                        {fieldLabelForFile(f.fileName) && (
                          <span style={{ display: 'inline-block', fontSize: '0.68rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', background: 'rgba(134,188,37,0.15)', borderRadius: 4, padding: '1px 6px', marginRight: 6 }}>
                            {fieldLabelForFile(f.fileName)}
                          </span>
                        )}
                        {prettyFileName(f.fileName)}
                      </a>
                      <a
                        href={f.serverRelativeUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.72rem', padding: '2px 10px', textDecoration: 'none' }}
                      >
                        <Download size={12} /> {isDe ? 'Download' : 'Download'}
                      </a>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.72rem', padding: '2px 10px', color: 'var(--dex-red, #c00)' }}
                        disabled={attachmentsBusy}
                        onClick={() => onDelete(f.fileName)}
                        title={isDe ? 'Löschen' : 'Delete'}
                      >
                        <X size={12} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
              <div style={{ display: 'flex', gap: 8, justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap' }}>
                <label className="btn btn-outline" style={{ fontSize: '0.82rem', padding: '6px 14px', cursor: attachmentsBusy ? 'wait' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <Plus size={14} /> {attachmentsBusy ? (isDe ? 'Wird übertragen…' : 'Uploading…') : (isDe ? 'Datei hinzufügen' : 'Add file')}
                  <input
                    type="file"
                    accept="application/pdf,image/*,.doc,.docx"
                    style={{ display: 'none' }}
                    onChange={onAdd}
                    disabled={attachmentsBusy}
                  />
                </label>
                <button className="btn btn-primary" onClick={close} disabled={attachmentsBusy}>
                  {isDe ? 'Schließen' : 'Close'}
                </button>
              </div>
          </Modal>
        );
};

