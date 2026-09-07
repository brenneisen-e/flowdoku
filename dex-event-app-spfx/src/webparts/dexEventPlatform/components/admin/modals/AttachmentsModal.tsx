/* AttachmentsModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 16192-16326 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { Download, FileText, Plus, Trash2 } from '../../Icons';
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
          // v31.2: Die Rückfrage nennt die Folge und den lesbaren Namen — der
          // rohe Name trägt Präfix und Zeitstempel, die niemand wiedererkennt.
          const shown = prettyFileName(fileName);
          if (!(await confirmDialog(
            isDe
              ? `Datei „${shown}“ löschen? Sie wird aus dieser Anmeldung entfernt.`
              : `Delete file “${shown}”? It will be removed from this registration.`,
            { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' },
          ))) return;
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
        // v31.2: Die Liste kennt keine Dateigröße — aber die Endung sagt dem
        // Organizer vor dem Klick, was aufgeht (PDF, Bild, Word).
        const kindOf = (fileName: string): string => {
          const ext = (fileName.split('.').pop() || '').toLowerCase();
          if (ext === 'pdf') return 'PDF';
          if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'heic', 'bmp', 'svg'].indexOf(ext) >= 0) return isDe ? 'Bild' : 'Image';
          if (ext === 'doc' || ext === 'docx') return 'Word';
          return ext ? ext.toUpperCase() : (isDe ? 'Datei' : 'File');
        };
        const dlLabel = isDe ? 'Herunterladen' : 'Download';
        const delLabel = isDe ? 'Löschen' : 'Delete';
        // v31.2: Kopf und Fuß kommen vom Modal; Zeilen mit Hover und Symbol-
        // Knöpfen statt grüner Kästen, Leerzustand statt kursivem Satz, Upload
        // als Kachel mit Formaten und Grenze — vorher stand „10 MB" nur im
        // Fehlerdialog NACH dem gescheiterten Versuch.
        return (
          <Modal open={true} onClose={close} dismissable={!attachmentsBusy} maxWidth={560}
            ariaLabel={isDe ? 'Hochgeladene Dateien' : 'Uploaded files'}
            title={isDe ? 'Dateien zur Anmeldung' : 'Files for this registration'}
            subtitle={<>{fullName}{reg.ParticipantEmail ? ` · ${reg.ParticipantEmail}` : ''}</>}
            icon={<FileText size={20} />}
            footer={
              <button type="button" className="btn btn-primary" onClick={close} disabled={attachmentsBusy}>
                {isDe ? 'Schließen' : 'Close'}
              </button>
            }
          >
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  {isDe ? 'Hochgeladen' : 'Uploaded'}
                  <span className="dex-ui-pill dex-ui-pill--gray">{list.length}</span>
                </div>
                {list.length === 0 ? (
                  <div className="dex-ui-empty">
                    <span className="dex-ui-empty-icon"><FileText size={20} /></span>
                    <div className="dex-ui-empty-title">{isDe ? 'Noch keine Dateien' : 'No files yet'}</div>
                    {isDe
                      ? 'Was die Person beim Anmelden hochlädt oder du hier hinzufügst, erscheint in dieser Liste.'
                      : 'Whatever the person uploads when registering or you add here shows up in this list.'}
                  </div>
                ) : (
                  <div className="dex-ui-card" style={{ padding: '4px 6px' }}>
                    {list.map(f => {
                      const label = fieldLabelForFile(f.fileName);
                      return (
                        <div key={f.fileName} className="dex-ui-row dex-ui-row--bordered">
                          <span style={{ color: 'var(--dex-gray-500)', display: 'inline-flex', flexShrink: 0 }}><FileText size={18} /></span>
                          <div className="dex-ui-row-main">
                            <a href={f.serverRelativeUrl} target="_blank" rel="noopener noreferrer" className="dex-ui-row-title"
                              title={isDe ? 'Datei öffnen' : 'Open file'} style={{ display: 'block', color: 'inherit', textDecoration: 'none' }}>
                              {prettyFileName(f.fileName)}
                            </a>
                            <div className="dex-ui-row-sub dex-ui-inline" style={{ gap: 6 }}>
                              {label && <span className="dex-ui-pill dex-ui-pill--green">{label}</span>}
                              <span>{kindOf(f.fileName)}</span>
                            </div>
                          </div>
                          <div className="dex-ui-row-actions">
                            <a href={f.serverRelativeUrl} target="_blank" rel="noopener noreferrer" className="dex-ui-iconbtn" title={dlLabel} aria-label={dlLabel}>
                              <Download size={16} />
                            </a>
                            <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" disabled={attachmentsBusy}
                              onClick={() => onDelete(f.fileName)} title={delLabel} aria-label={delLabel}>
                              <Trash2 size={16} />
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Datei hinzufügen' : 'Add a file'}</div>
                <label className={cx('dex-ui-choice', attachmentsBusy && 'is-disabled')} style={{ cursor: attachmentsBusy ? 'wait' : 'pointer' }}>
                  <span className="dex-ui-choice-icon"><Plus size={18} /></span>
                  <span className="dex-ui-choice-body">
                    <span className="dex-ui-choice-title" style={{ display: 'block' }}>
                      {attachmentsBusy
                        ? (isDe ? 'Wird übertragen…' : 'Uploading…')
                        : (isDe ? 'Datei auswählen und hochladen' : 'Choose a file and upload it')}
                    </span>
                    <span className="dex-ui-choice-desc" style={{ display: 'block' }}>
                      {isDe
                        ? 'PDF, Bild oder Word, höchstens 10 MB. Die Datei wird dieser Anmeldung angehängt.'
                        : 'PDF, image or Word, up to 10 MB. The file is attached to this registration.'}
                    </span>
                  </span>
                  <input type="file" accept="application/pdf,image/*,.doc,.docx" style={{ display: 'none' }} onChange={onAdd} disabled={attachmentsBusy} />
                </label>
              </div>
          </Modal>
        );
};

