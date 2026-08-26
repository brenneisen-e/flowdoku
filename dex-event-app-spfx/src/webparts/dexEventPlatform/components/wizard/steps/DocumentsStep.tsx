/**
 * v30.13 — Modularisierung Stufe 3: Schritt „Dokumente" (Step 8, Index 7)
 * als eigene Komponente. JSX 1:1 aus EventCreationPage; der Schritt hängt
 * an vier State-Paaren (Dokumentliste, Teilnehmer-Upload-Toggle samt
 * Anzeigename/Hinweis). `visible` ersetzt `currentStep === 7` —
 * display:none statt unmount, damit Eingaben beim Schrittwechsel erhalten
 * bleiben. Sprache kommt wie überall aus useLanguage; renderStepIntro
 * bleibt Prop, weil der Stub am Wizard lebt (aktuell bewusst null).
 */
import * as React from 'react';
import { Plus, X } from '../../Icons';
import { InfoTooltip } from '../../InfoTooltip';
import { StepBadge } from '../StepBadge';
import { Icon } from '@fluentui/react/lib/Icon';
import { useLanguage } from '../../../context/LanguageContext';

export interface DocumentsStepProps {
  visible: boolean;
  documents: Array<{ name: string; file?: File; url: string; size: number }>;
  setDocuments: (docs: Array<{ name: string; file?: File; url: string; size: number }>) => void;
  allowAttendeeUpload: boolean;
  setAllowAttendeeUpload: (v: boolean) => void;
  attendeeUploadLabel: string;
  setAttendeeUploadLabel: (v: string) => void;
  attendeeUploadHint: string;
  setAttendeeUploadHint: (v: string) => void;
  renderStepIntro: (bulletsDe: string[], bulletsEn: string[]) => React.ReactElement | null;
}

export const DocumentsStep: React.FC<DocumentsStepProps> = ({
  visible, documents, setDocuments,
  allowAttendeeUpload, setAllowAttendeeUpload,
  attendeeUploadLabel, setAttendeeUploadLabel,
  attendeeUploadHint, setAttendeeUploadHint,
  renderStepIntro,
}) => {
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  return (
    <div style={{ display: visible ? 'block' : 'none' }}>
      <h2 className="dex-step-head-title">
        {isDe ? 'Schritt 8 — Dokumente' : 'Step 8 — Documents'}
      </h2>
      <p className="dex-step-head-lead">
        {isDe
          ? 'Hier lädst du alle Unterlagen hoch, die deine Teilnehmer rund um das Event brauchen — von der Agenda bis zur Anfahrt.'
          : 'Here you upload all documents attendees might need around the event — from the agenda to the travel directions.'}
      </p>
      {renderStepIntro(
        [
          'Programm / Agenda pflegen (mehrtägig möglich, Drag-Reihenfolge pro Tag)',
          'Transferzeiten — Bus / Shuttle / Bahn von/zum Veranstaltungsort',
          'Dokumente hochladen (PDF) — Teilnehmer sehen sie auf MyEvents als Inline-Vorschau oder Download',
        ],
        [
          'Maintain the event programme / agenda (multi-day supported, drag-reorder per day)',
          'Transfer times — bus / shuttle / train to and from the venue',
          'Upload documents (PDF) — attendees see them on MyEvents as inline preview or download',
        ]
      )}
      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
        <StepBadge n={33} />
        {isDe ? 'Dokumente hochladen' : 'Upload documents'}
      </label>
      {/* v9.28: Schlagwörter fett rendern für bessere Lesbarkeit. */}
      <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', marginBottom: 16, lineHeight: 1.6 }}>
        {isDe ? (
          <>
            Lade hier alle Unterlagen hoch, die deine Teilnehmer rund um das Event brauchen — z.B.
            die <strong>Detail-Agenda als PDF</strong>, eine <strong>Anfahrtsbeschreibung mit Karte</strong>,
            die <strong>Hausordnung</strong> des Veranstaltungsorts, eine <strong>Packliste</strong>, das <strong>Teilnehmer-Briefing</strong>
            {' '}oder eine <strong>Datenschutz-/Foto-Einverständniserklärung</strong>. Die Dokumente erscheinen
            automatisch unter <strong>{'„Meine Events“'}</strong> in der Detail-Ansicht des Teilnehmers — dort sehen sie eine
            <strong> Inline-Vorschau</strong> (bei PDFs) und können das Dokument einzeln <strong>herunterladen</strong>.
            Mehrere Dateien gleichzeitig hochladen geht per <strong>Drag &amp; Drop</strong> oder <strong>Mehrfachauswahl</strong>.
            Du kannst Dokumente auch <strong>nach dem Event-Live-Gang</strong> noch hinzufügen oder austauschen — die
            Teilnehmer sehen immer die aktuelle Version. <strong>Tipp:</strong> für Dokumente, die nur intern für die
            Organizer wichtig sind (z.B. Kontaktliste vom Caterer), nutze eine geteilte
            <strong> Teams-/SharePoint-Ablage außerhalb von DEX</strong>, da hier alle Teilnehmer Lese-Zugriff haben.
          </>
        ) : (
          <>
            Upload everything attendees might need around the event — e.g.
            the <strong>detailed agenda as PDF</strong>, <strong>directions with a map</strong>,
            the venue&apos;s <strong>house rules</strong>, a <strong>packing list</strong>, the <strong>attendee briefing</strong>
            {' '}or a <strong>privacy/photo consent form</strong>. Documents show up automatically under
            <strong>{' „My Events“'}</strong> in the attendee detail view — they get an
            <strong> inline preview</strong> (for PDFs) and can <strong>download</strong> each file individually.
            Multiple files can be uploaded at once via <strong>drag &amp; drop</strong> or <strong>multi-select</strong>.
            You can keep adding or replacing documents <strong>after the event has gone live</strong> — attendees always
            see the latest version. <strong>Tip:</strong> for documents only meant for organizers (e.g. caterer
            contact list), use a shared <strong>Teams/SharePoint folder outside DEX</strong>, because every attendee has read access here.
          </>
        )}
      </p>

      {documents.map((doc, idx) => (
        <div key={idx} style={{
          display: 'flex', alignItems: 'center', gap: 8, padding: '8px 12px', marginBottom: 6,
          background: 'var(--dex-gray-50, #fafafa)', borderRadius: 'var(--dex-radius)',
          border: '1px solid var(--dex-gray-200)',
        }}>
          <Icon iconName="Page" style={{ fontSize: 16, color: 'var(--dex-gray-600)' }} />
          <span style={{ flex: 1, fontSize: '0.85rem' }}>{doc.name}</span>
          {doc.size > 0 && <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)' }}>{(doc.size / 1024).toFixed(0)} KB</span>}
          <button type="button" onClick={() => setDocuments(documents.filter((_, i) => i !== idx))} style={{
            background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)',
            fontSize: '1.1rem', padding: '4px', lineHeight: 1,
          }} title={t('general.delete')}>
            <X size={16} />
          </button>
        </div>
      ))}

      <label className="btn btn-outline" style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Plus size={14} /> {t('create.documents.upload')}
        <input
          type="file"
          multiple
          style={{ display: 'none' }}
          onChange={(e) => {
            const files = e.target.files;
            if (!files) return;
            const newDocs = Array.from(files).map(f => ({ name: f.name, file: f, url: '', size: f.size }));
            setDocuments([...documents, ...newDocs]);
            e.target.value = '';
          }}
        />
      </label>

      {/* v11.0: Teilnehmer-Upload-Toggle. Default OFF — wird nur
          bei expliziter Aktivierung in „Meine Events" als Upload-
          Bereich für die Anmeldung sichtbar. Anzeigename und
          Hinweistext sind frei konfigurierbar. */}
      <div style={{ marginTop: 32, paddingTop: 20, borderTop: '2px solid var(--dex-gray-100)' }}>
        <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
          <StepBadge n={34} />
          {isDe ? 'Teilnehmer-Upload erlauben' : 'Allow attendee upload'}
          <InfoTooltip text={isDe ? (
            <>
              <strong>Was du hier einstellst:</strong> ob jeder Teilnehmer in &bdquo;Meine Events&ldquo; eine eigene Datei (z.B. PDF) zu seiner Anmeldung hochladen darf.<br /><br />
              <strong>Beispiele:</strong> Reisekostenbeleg, unterschriebener Datenschutzbogen, Foto-Einverständnis, Zertifikat als Voraussetzung für die Teilnahme.<br /><br />
              <strong>Ablauf:</strong> Teilnehmer sieht nach der Anmeldung in &bdquo;Meine Events&ldquo; einen Upload-Block mit deinem Anzeigenamen + Hinweistext. Hochgeladene Dateien werden direkt als <strong>Item-Attachment</strong> an die Teilnehmer-Zeile in der SharePoint-Subsite gehängt — nicht in einer Sammeldatei. Der Teilnehmer kann seine Datei jederzeit ersetzen oder löschen.<br /><br />
              <strong>Admin-Sicht:</strong> du siehst im Admin-Center pro Teilnehmer-Zeile alle hochgeladenen Dateien als Liste mit Download-Link. Du kannst auch fremde Uploads löschen.<br /><br />
              <strong>Default: aus.</strong> Nur einschalten, wenn du tatsächlich ein Dokument von Teilnehmern brauchst.
            </>
          ) : (
            <>
              <strong>What you set here:</strong> whether every attendee can upload a file (e.g. PDF) to their registration via &ldquo;My Events&rdquo;.<br /><br />
              <strong>Examples:</strong> travel-expense receipt, signed privacy form, photo-consent, certificate as a prerequisite to attend.<br /><br />
              <strong>Flow:</strong> after registering, the attendee sees an upload block in &ldquo;My Events&rdquo; with the display name and hint text you configure. Uploaded files attach directly as <strong>item attachments</strong> on the attendee&apos;s row in the SharePoint subsite — not into a collection file. Attendees can replace or delete their own file any time.<br /><br />
              <strong>Admin view:</strong> you see every uploaded file per attendee in the admin center with a download link. You can also delete attendee uploads.<br /><br />
              <strong>Default: off.</strong> Enable only when you actually need a document from attendees.
            </>
          )} />
        </label>
        <label
          style={{
            display: 'inline-flex', alignItems: 'center', gap: 8,
            padding: '8px 14px', borderRadius: 999,
            fontSize: '0.85rem', fontWeight: 600, whiteSpace: 'nowrap',
            cursor: 'pointer', userSelect: 'none',
            border: `1px solid ${allowAttendeeUpload ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
            background: allowAttendeeUpload ? 'rgba(134,188,37,0.10)' : '#fff',
            color: allowAttendeeUpload ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
            transition: 'all 0.15s ease',
          }}
        >
          <input
            type="checkbox"
            checked={allowAttendeeUpload}
            onChange={e => setAllowAttendeeUpload(e.target.checked)}
            style={{ display: 'none' }}
          />
          <span style={{ fontSize: '0.9rem', lineHeight: 1 }}>{allowAttendeeUpload ? '✓' : '○'}</span>
          {isDe ? 'Teilnehmer dürfen Datei hochladen' : 'Attendees may upload a file'}
        </label>

        {allowAttendeeUpload && (
          <div style={{ marginTop: 14, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
            <div>
              <label className="form-label" style={{ fontSize: '0.82rem', marginBottom: 4 }}>
                {isDe ? 'Anzeige-Name (z.B. „Reisekostenbeleg")' : 'Display name (e.g. „Travel-expense receipt")'}
              </label>
              <input
                className="form-input"
                value={attendeeUploadLabel}
                onChange={e => setAttendeeUploadLabel(e.target.value)}
                placeholder={isDe ? 'z.B. Reisekostenbeleg, Datenschutz-Erklärung' : 'e.g. Travel receipt, privacy form'}
                maxLength={80}
              />
            </div>
            <div>
              <label className="form-label" style={{ fontSize: '0.82rem', marginBottom: 4 }}>
                {isDe ? 'Hinweistext für Teilnehmer (optional)' : 'Hint text for attendees (optional)'}
              </label>
              <input
                className="form-input"
                value={attendeeUploadHint}
                onChange={e => setAttendeeUploadHint(e.target.value)}
                placeholder={isDe ? 'z.B. Bitte unterschrieben hochladen' : 'e.g. Please upload signed'}
                maxLength={240}
              />
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
