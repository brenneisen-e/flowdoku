/* MyEventDocField — aus MyEventsPage.tsx ausgelagert (Zeilen 3800-3866 des
 * urspruenglichen Stands, v30.65). Upload-Block fuer EIN Dokument-Feld.
 *
 * v31.8: Optik auf die dex-ui-Klassen umgestellt, baugleich zu MyEventUpload
 * (docs/ui-leitfaden.md 6a-6d). Handler, Bedingungen und Props sind
 * unveraendert — die beiden Kaesten standen direkt untereinander in derselben
 * Karte und sahen trotzdem unterschiedlich aus (anderes Symbol, andere Farbe).
 * Das Stylesheet injiziert die Seite (`MyEventsPage`) ueber
 * `ensureDexUiStyles()`; Unterkomponenten rufen es nie.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DeloitteEvent } from '../../types';
import { useDialog } from '../../context/DialogContext';
import { useLocaleSafe } from '../../context/LanguageContext';
import { AlertCircle, FileText, Trash2 } from '../Icons';

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
  // v31.8: Die Anzeigesprache kommt aus dem LanguageContext, nicht aus
  // `emailLanguage` — das ist die Sprache der MAILS. Vorher konnte dieser
  // Kasten englisch sein, während die Seite drumherum deutsch war.
  const isDe = useLocaleSafe() === 'de';
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
  // v31.8: Nur eine Ableitung fürs Rendern — die Liste selbst bleibt unberührt.
  const hasFiles = files.length > 0;

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
    // v31.8: Die Rückfrage nennt jetzt die Folge (Leitfaden 2b) — vorher stand
    // dort nur die Frage, ob gelöscht werden soll.
    if (!(await confirmDialog(
      isDe
        ? `Datei „${displayName}“ wirklich löschen? Sie wird von deiner Anmeldung entfernt, der Organizer sieht sie dann nicht mehr.`
        : `Really delete file „${displayName}“? It is removed from your registration and the organizer will no longer see it.`,
      { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' }))) return;
    setBusy(true);
    try { await props.remove(event.id, fileName); await refresh(); } finally { setBusy(false); }
  };

  return (
    // v31.8: `dex-ui-section` statt handgebautem Kasten; der Trennstrich nach
    // oben bleibt inline, weil er den Block von den Geschwisterblöcken der
    // Karte abgrenzt und dort dieselbe Linie sitzt.
    <div className="dex-ui-section" style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }} aria-busy={busy}>
      <div className="dex-ui-section-title">
        <Icon iconName="Attach" style={{ fontSize: 14 }} />
        {/* v31.8 (Nachzug): Beschriftung und Stern zusammen in EINEM Element.
            `dex-ui-section-title` ist ein Flex-Container mit `gap: 10px` —
            als zwei Kinder stand der Stern zwölf Pixel neben dem Wort und sah
            aus wie ein eigener Hinweis. */}
        <span>{field.label}{field.required ? <span className="dex-ui-label-required">*</span> : ''}</span>
      </div>

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
                {f.displayName}
              </a>
              <span className="dex-ui-row-actions">
                <button
                  type="button"
                  className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                  disabled={busy}
                  onClick={() => onDelete(f.fileName, f.displayName)}
                  title={isDe ? 'Datei löschen' : 'Delete file'}
                  aria-label={isDe ? `Datei ${f.displayName} löschen` : `Delete file ${f.displayName}`}
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
        <input type="file" accept="application/pdf,image/png,image/jpeg,.pdf,.png,.jpg,.jpeg" style={{ display: 'none' }} onChange={onPick} disabled={busy} />
      </label>
      {/* v31.8: Format, Grenze und Pflicht stehen jetzt VOR dem Fehlerfall —
          bisher trug „(PDF/Bild)" der Knopf und von den 10 MB erfuhr man erst
          nach der Ablehnung. */}
      <div className="dex-ui-help">
        {field.required
          ? (isDe ? 'Pflichtfeld — PDF oder Bild, höchstens 10 MB.' : 'Required — PDF or image, 10 MB at most.')
          : (isDe ? 'PDF oder Bild, höchstens 10 MB.' : 'PDF or image, 10 MB at most.')}
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
