/* MyEventUpload — aus MyEventsPage.tsx ausgelagert (Zeilen 3594-3596 und
 * 3676-3798 des urspruenglichen Stands, v30.65). Teilnehmer-Upload an der
 * eigenen Anmeldezeile.
 *
 * v31.8: Optik auf die dex-ui-Klassen umgestellt (docs/ui-leitfaden.md 6a-6d) —
 * die Handler, ihre Bedingungen und die Props sind unveraendert. Der Kasten war
 * vorher komplett inline gestylt und hatte deshalb nirgends einen Hover; die
 * Dateiliste sah aus wie eine Erfolgsmeldung (gruen hinterlegt), obwohl sie eine
 * Liste mit Aktionen ist. Das Stylesheet injiziert die Seite
 * (`MyEventsPage`) ueber `ensureDexUiStyles()`; Unterkomponenten rufen es nie.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DeloitteEvent } from '../../types';
import { useDialog } from '../../context/DialogContext';
import { useLocaleSafe } from '../../context/LanguageContext';
import { AlertCircle, FileText, Trash2 } from '../Icons';

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
  // v31.8: Die Anzeigesprache kommt aus dem LanguageContext, nicht aus
  // `emailLanguage` — das ist die Sprache der MAILS. Vorher konnte dieser
  // Kasten englisch sein, während die Seite drumherum deutsch war.
  const isDe = useLocaleSafe() === 'de';
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
  // v31.8: Nur eine Ableitung fürs Rendern — die Liste selbst bleibt unberührt.
  const hasFiles = files.length > 0;

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
    // v31.8: Die Rückfrage nennt jetzt die Folge (Leitfaden 2b) — vorher stand
    // dort nur die Frage, ob gelöscht werden soll.
    if (!(await confirmDialog(
      isDe
        ? `Datei „${fileName}“ wirklich löschen? Sie wird von deiner Anmeldung entfernt, der Organizer sieht sie dann nicht mehr.`
        : `Really delete file „${fileName}“? It is removed from your registration and the organizer will no longer see it.`,
      { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' }))) return;
    setBusy(true);
    try {
      await props.remove(event.id, fileName);
      await refresh();
    } finally {
      setBusy(false);
    }
  };

  return (
    // v31.8: `dex-ui-section` statt handgebautem Kasten; der Trennstrich nach
    // oben bleibt inline, weil er den Block von den Geschwisterblöcken der
    // Karte abgrenzt und dort dieselbe Linie sitzt.
    <div className="dex-ui-section" style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }} aria-busy={busy}>
      <div className="dex-ui-section-title">
        <Icon iconName="Attach" style={{ fontSize: 14 }} />
        {label}
      </div>
      {hint && (
        // Organizer-Text — bleibt wortgleich, bekommt nur die Abschnittsoptik.
        <div className="dex-ui-section-desc">{hint}</div>
      )}

      {hasFiles && (
        <div className="dex-ui-card dex-ui-card--list" style={{ marginBottom: 10 }}>
          {files.map(f => (
            // v31.8: Zeile mit Hover statt grün hinterlegtem Kasten — grün
            // heißt in der App „aktiv/Erfolg", eine Dateizeile ist beides nicht.
            <div key={f.fileName} className="dex-ui-row dex-ui-row--bordered">
              <span style={{ display: 'inline-flex', color: 'var(--dex-gray-500)', flexShrink: 0 }}>
                <FileText size={16} />
              </span>
              <a
                className="dex-ui-row-main"
                href={f.serverRelativeUrl}
                target="_blank"
                rel="noopener noreferrer"
                style={{ color: 'var(--dex-gray-800)', fontWeight: 600, fontSize: '0.85rem', textDecoration: 'none', wordBreak: 'break-word' }}
              >
                {f.fileName}
              </a>
              <span className="dex-ui-row-actions">
                <button
                  type="button"
                  className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                  disabled={busy}
                  onClick={() => onDelete(f.fileName)}
                  title={isDe ? 'Datei löschen' : 'Delete file'}
                  aria-label={isDe ? `Datei ${f.fileName} löschen` : `Delete file ${f.fileName}`}
                >
                  <Trash2 size={15} />
                </button>
              </span>
            </div>
          ))}
        </div>
      )}

      {/* v31.8: Ohne Datei ist der Upload die Handlung der Sektion und bekommt
          die Ablagefläche; liegt schon etwas da, reicht ein kompakter Knopf.
          Bewusst KEIN Leer-Hinweis „noch keine Datei" — `refresh` schluckt
          Lesefehler, eine leere Liste wäre also keine Aussage (CLAUDE.md). */}
      <label
        className={hasFiles ? 'btn btn-outline dex-ui-btn-sm' : 'dex-ui-dropzone'}
        style={hasFiles
          ? { cursor: busy ? 'wait' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }
          : { cursor: busy ? 'wait' : 'pointer' }}
      >
        <Icon iconName="Upload" style={{ fontSize: hasFiles ? 14 : 20 }} />
        <span style={{ fontWeight: 600 }}>
          {busy
            ? (isDe ? 'Wird übertragen…' : 'Uploading…')
            : hasFiles
              ? (isDe ? 'Weitere Datei hochladen' : 'Upload another file')
              : (isDe ? 'Datei hochladen' : 'Upload file')}
        </span>
        <input
          type="file"
          accept="application/pdf,image/*,.doc,.docx"
          style={{ display: 'none' }}
          onChange={onPick}
          disabled={busy}
        />
      </label>
      {/* v31.8: Grenze und erlaubte Formate stehen jetzt VOR dem Fehlerfall —
          bisher erfuhr man von den 10 MB erst, wenn der Upload abgelehnt wurde. */}
      <div className="dex-ui-help">
        {isDe ? 'PDF, Bild oder Word-Dokument, höchstens 10 MB.' : 'PDF, image or Word document, 10 MB at most.'}
      </div>

      {error && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert" style={{ marginTop: 10 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>{error}</span>
        </div>
      )}
    </div>
  );
}
