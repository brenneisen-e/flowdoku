/* MyEventUpload — aus MyEventsPage.tsx ausgelagert (Zeilen 3594-3596 und
 * 3676-3798 des urspruenglichen Stands, v30.65). Teilnehmer-Upload an der
 * eigenen Anmeldezeile. Der Code ist zeichengleich uebernommen.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DeloitteEvent } from '../../types';
import { useDialog } from '../../context/DialogContext';

// ==================== Datei-Upload ("Meine Events") ====================
// v11.0: Wenn der Organizer beim Event den Toggle „Teilnehmer-Upload
// erlauben" gesetzt hat, sieht der Teilnehmer hier einen Upload-Block.

// Dateien werden als SP-Item-Attachment direkt an die eigene Teilnehmer-
// zeile gehängt — der Admin sieht sie im Admin Center.
export default function MyEventUpload(props: {
  event: DeloitteEvent;
  list: (eventId: string) => Promise<Array<{ fileName: string; serverRelativeUrl: string }>>;
  upload: (eventId: string, file: File) => Promise<boolean>;
  remove: (eventId: string, fileName: string) => Promise<boolean>;
}): React.ReactElement | null {
  const { event } = props;
  const isDe = (event.emailLanguage || 'EN').toUpperCase() === 'DE';
  // v20.4: App-Modal statt window.confirm.
  const { confirmDialog } = useDialog();
  const [files, setFiles] = React.useState<Array<{ fileName: string; serverRelativeUrl: string }>>([]);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>('');

  const refresh = React.useCallback(async () => {
    try {
      const list = await props.list(event.id);
      setFiles(list);
    } catch { /* ignore */ }
  }, [event.id]);

  React.useEffect(() => { refresh().catch(() => { /* */ }); }, [refresh]);

  if (!event.allowAttendeeUpload) return null;

  const label = (event.attendeeUploadLabel || '').trim() || (isDe ? 'Dokumenten-Upload' : 'Document upload');
  const hint = (event.attendeeUploadHint || '').trim();

  const onPick = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!f) return;
    // Soft-Limit 10 MB — SharePoint Item-Attachments sind technisch bis 250 MB,
    // aber große PDFs sprengen die App-UX und Power-Automate-Workflows.
    if (f.size > 10 * 1024 * 1024) {
      setError(isDe ? 'Datei ist größer als 10 MB. Bitte komprimieren oder kleinere Variante hochladen.' : 'File is larger than 10 MB. Please compress or upload a smaller version.');
      return;
    }
    setError('');
    setBusy(true);
    try {
      const ok = await props.upload(event.id, f);
      if (!ok) setError(isDe ? 'Upload fehlgeschlagen.' : 'Upload failed.');
      await refresh();
    } finally {
      setBusy(false);
    }
  };

  const onDelete = async (fileName: string): Promise<void> => {
    if (!(await confirmDialog(isDe ? `Datei „${fileName}" wirklich löschen?` : `Really delete file „${fileName}"?`, { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' }))) return;
    setBusy(true);
    try {
      await props.remove(event.id, fileName);
      await refresh();
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Icon iconName="Attach" style={{ fontSize: 14 }} />
        {label}
      </div>
      {hint && (
        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginBottom: 10, lineHeight: 1.45 }}>
          {hint}
        </div>
      )}
      {files.length > 0 && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 10 }}>
          {files.map(f => (
            <div key={f.fileName} style={{
              display: 'flex', alignItems: 'center', gap: 10,
              padding: '6px 10px', borderRadius: 6,
              background: 'rgba(134,188,37,0.08)',
              border: '1px solid rgba(134,188,37,0.30)',
              fontSize: '0.82rem',
            }}>
              <Icon iconName="PDF" style={{ fontSize: 16, color: 'var(--dex-red, #c00)' }} />
              <a
                href={f.serverRelativeUrl}
                target="_blank"
                rel="noopener noreferrer"
                style={{ flex: 1, color: 'var(--dex-gray-800)', textDecoration: 'none', wordBreak: 'break-all' }}
              >
                {f.fileName}
              </a>
              <button
                type="button"
                className="btn btn-secondary"
                style={{ fontSize: '0.72rem', padding: '2px 10px' }}
                disabled={busy}
                onClick={() => onDelete(f.fileName)}
                title={isDe ? 'Löschen' : 'Delete'}
              >
                ✕
              </button>
            </div>
          ))}
        </div>
      )}
      <label className="btn btn-outline" style={{ fontSize: '0.82rem', padding: '6px 14px', cursor: busy ? 'wait' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Icon iconName="Upload" style={{ fontSize: 14 }} />
        {busy ? (isDe ? 'Wird übertragen…' : 'Uploading…') : files.length > 0 ? (isDe ? 'Weitere Datei hochladen' : 'Upload another file') : (isDe ? 'Datei hochladen' : 'Upload file')}
        <input
          type="file"
          accept="application/pdf,image/*,.doc,.docx"
          style={{ display: 'none' }}
          onChange={onPick}
          disabled={busy}
        />
      </label>
      {error && (
        <div style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-red, #c00)' }}>{error}</div>
      )}
    </div>
  );
}
