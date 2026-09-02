/* MyEventDocField — aus MyEventsPage.tsx ausgelagert (Zeilen 3800-3866 des
 * urspruenglichen Stands, v30.65). Upload-Block fuer EIN Dokument-Feld.
 * Der Code ist zeichengleich uebernommen.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DeloitteEvent } from '../../types';
import { useDialog } from '../../context/DialogContext';

// v19.0: Upload-Block für EIN Dokument-Custom-Feld. Wird pro Dokument-Feld des
// Events gerendert, damit der Teilnehmer die Datei auch nachträglich
// ergänzen/ersetzen/löschen kann. Datei = SP-Item-Attachment (pro Feld über
// Dateinamen-Präfix zugeordnet) — der Organizer sieht sie im Admin Center.
export default function MyEventDocField(props: {
  event: DeloitteEvent;
  field: { id: string; label: string; required?: boolean };
  list: (eventId: string, fieldId: string, participantEmail?: string) => Promise<Array<{ fileName: string; serverRelativeUrl: string; displayName: string }>>;
  upload: (eventId: string, fieldId: string, file: File, participantEmail?: string) => Promise<boolean>;
  remove: (eventId: string, fileName: string, participantEmail?: string) => Promise<boolean>;
}): React.ReactElement {
  const { event, field } = props;
  const isDe = (event.emailLanguage || 'EN').toUpperCase() === 'DE';
  // v20.4: App-Modal statt window.confirm.
  const { confirmDialog } = useDialog();
  const [files, setFiles] = React.useState<Array<{ fileName: string; serverRelativeUrl: string; displayName: string }>>([]);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState('');
  const refresh = React.useCallback(async () => {
    try { setFiles(await props.list(event.id, field.id)); } catch { /* ignore */ }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [event.id, field.id]);
  React.useEffect(() => { refresh().catch(() => { /* */ }); }, [refresh]);
  const onPick = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!f) return;
    if (f.size > 10 * 1024 * 1024) { setError(isDe ? 'Datei ist größer als 10 MB. Bitte kleinere Variante hochladen.' : 'File is larger than 10 MB. Please upload a smaller version.'); return; }
    setError(''); setBusy(true);
    try {
      const ok = await props.upload(event.id, field.id, f);
      if (!ok) setError(isDe ? 'Upload fehlgeschlagen.' : 'Upload failed.');
      await refresh();
    } finally { setBusy(false); }
  };
  const onDelete = async (fileName: string, displayName: string): Promise<void> => {
    // eslint-disable-next-line no-alert
    if (!(await confirmDialog(isDe ? `Datei „${displayName}" wirklich löschen?` : `Really delete file „${displayName}"?`, { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' }))) return;
    setBusy(true);
    try { await props.remove(event.id, fileName); await refresh(); } finally { setBusy(false); }
  };
  return (
    <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 8, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Icon iconName="Attach" style={{ fontSize: 14 }} />
        {field.label}{field.required ? ' *' : ''}
      </div>
      {files.length > 0 && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 10 }}>
          {files.map(f => (
            <div key={f.fileName} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '6px 10px', borderRadius: 6, background: 'rgba(134,188,37,0.08)', border: '1px solid rgba(134,188,37,0.30)', fontSize: '0.82rem' }}>
              <Icon iconName="Page" style={{ fontSize: 16, color: 'var(--dex-green-dark, #4a7c1f)' }} />
              <a href={f.serverRelativeUrl} target="_blank" rel="noopener noreferrer" style={{ flex: 1, color: 'var(--dex-gray-800)', textDecoration: 'none', wordBreak: 'break-all' }}>{f.displayName}</a>
              <button type="button" className="btn btn-secondary" style={{ fontSize: '0.72rem', padding: '2px 10px' }} disabled={busy} onClick={() => onDelete(f.fileName, f.displayName)} title={isDe ? 'Löschen' : 'Delete'}>✕</button>
            </div>
          ))}
        </div>
      )}
      <label className="btn btn-outline" style={{ fontSize: '0.82rem', padding: '6px 14px', cursor: busy ? 'wait' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Icon iconName="Upload" style={{ fontSize: 14 }} />
        {busy ? (isDe ? 'Wird übertragen…' : 'Uploading…') : files.length > 0 ? (isDe ? 'Weitere Datei hochladen' : 'Upload another file') : (isDe ? 'Datei hochladen (PDF/Bild)' : 'Upload file (PDF/image)')}
        <input type="file" accept="application/pdf,image/png,image/jpeg,.pdf,.png,.jpg,.jpeg" style={{ display: 'none' }} onChange={onPick} disabled={busy} />
      </label>
      {error && <div style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-red, #c00)' }}>{error}</div>}
    </div>
  );
}
