/**
 * v30.13 — Modularisierung Stufe 3: Schritt „Dokumente" (Step 8, Index 7)
 * als eigene Komponente. JSX 1:1 aus EventCreationPage; der Schritt hängt
 * an vier State-Paaren (Dokumentliste, Teilnehmer-Upload-Toggle samt
 * Anzeigename/Hinweis). `visible` ersetzt `currentStep === 7` —
 * display:none statt unmount, damit Eingaben beim Schrittwechsel erhalten
 * bleiben. Sprache kommt wie überall aus useLanguage; renderStepIntro
 * bleibt Prop, weil der Stub am Wizard lebt (aktuell bewusst null).
 *
 * v31.2: Zwei Abschnitte mit je einer Frage — Unterlagen (Ablage-Zone mit
 * Drag & Drop, Liste mit Hover) und Teilnehmer-Upload (Schalter-Zeile, abhängige
 * Felder gedämpft statt ausgeblendet). Erklärtext im Tooltip; Props unverändert.
 */
import * as React from 'react';
import { Download, FileText, Plus, X } from '../../Icons';
import { InfoTooltip } from '../../InfoTooltip';
import { StepBadge } from '../StepBadge';
import { useLanguage } from '../../../context/LanguageContext';
// v31.2: Gemeinsame UI-Klassen (Abschnitt, Zeile, Schalter-Zeile, Karte) statt
// Inline-Style-Stapel — nur so bekommt jedes klickbare Element einen Hover.
import { cx } from '../../dexUi';

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
  // v31.2: Drag-over hat keine CSS-Pseudoklasse — deshalb ein State für die Zone.
  const [dragOver, setDragOver] = React.useState(false);
  // v31.2: EIN Pfad für Datei-Dialog und Drop — dieselbe Abbildung wie bisher im
  // onChange, damit der Save-Pfad beide Wege gleich behandelt.
  const addFiles = (files: FileList | null): void => {
    if (!files) return;
    const newDocs = Array.from(files).map(f => ({ name: f.name, file: f, url: '', size: f.size }));
    setDocuments([...documents, ...newDocs]);
  };
  return (
    <div style={{ display: visible ? 'block' : 'none' }}>
      <h2 className="dex-step-head-title">
        <span className="dex-step-eyebrow">{isDe ? 'Schritt 8 von 9' : 'Step 8 of 9'}</span>
        {isDe ? 'Dokumente' : 'Documents'}
      </h2>
      <p className="dex-step-head-lead">
        {isDe
          ? <><strong>Optional.</strong> Unterlagen, die deine Teilnehmer rund um das Event brauchen — von der Agenda bis zur Anfahrt. Und falls du etwas von ihnen brauchst: ein Upload-Feld für sie.</>
          : <><strong>Optional.</strong> Documents your attendees need around the event — from the agenda to directions. And if you need something from them: an upload field for attendees.</>}
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
      {/* v9.28: Schlagwörter fett rendern für bessere Lesbarkeit. */}
      {/* v31.2: Zwei Zeilen sichtbar, der Rest (Live-Gang, Tipp zu internen Dokumenten) im Tooltip. */}
      <div className="dex-ui-section">
        <div className="dex-ui-section-title">
          <StepBadge n={33} />
          {isDe ? 'Unterlagen für Teilnehmer' : 'Documents for attendees'}
          {documents.length > 0 && <span className="dex-ui-pill dex-ui-pill--gray">{documents.length}</span>}
        </div>
        <div className="dex-ui-label" style={{ marginBottom: 4 }}>
          {isDe ? 'Welche Unterlagen sollen deine Teilnehmer bekommen?' : 'Which documents should your attendees get?'}
          <InfoTooltip text={isDe ? (
            <>
              <strong>Was du hier einstellst:</strong> alle Unterlagen rund um das Event — z.B. die <strong>Detail-Agenda als PDF</strong>, eine <strong>Anfahrtsbeschreibung mit Karte</strong>, die <strong>Hausordnung</strong> des Veranstaltungsorts, eine <strong>Packliste</strong>, das <strong>Teilnehmer-Briefing</strong> oder eine <strong>Datenschutz-/Foto-Einverständniserklärung</strong>. Mehrere Dateien gleichzeitig gehen per <strong>Drag &amp; Drop</strong> oder <strong>Mehrfachauswahl</strong>.<br /><br />
              <strong>Anzeige in der App:</strong> die Dokumente erscheinen automatisch unter &bdquo;Meine Events&ldquo; in der Detail-Ansicht des Teilnehmers — mit <strong>Inline-Vorschau</strong> (bei PDFs) und einzelnem <strong>Download</strong>.<br /><br />
              <strong>Auswirkung:</strong> du kannst Dokumente auch <strong>nach dem Live-Gang</strong> hinzufügen oder austauschen — Teilnehmer sehen immer die aktuelle Version.<br /><br />
              <strong>Tipp:</strong> für Dokumente, die nur intern für die Organizer wichtig sind (z.B. Kontaktliste vom Caterer), nutze eine geteilte <strong>Teams-/SharePoint-Ablage außerhalb von DEX</strong>, da hier alle Teilnehmer Lese-Zugriff haben.
            </>
          ) : (
            <>
              <strong>What you set here:</strong> everything attendees need around the event — e.g. the <strong>detailed agenda as PDF</strong>, <strong>directions with a map</strong>, the venue&apos;s <strong>house rules</strong>, a <strong>packing list</strong>, the <strong>attendee briefing</strong> or a <strong>privacy/photo consent form</strong>. Several files at once via <strong>drag &amp; drop</strong> or <strong>multi-select</strong>.<br /><br />
              <strong>Where you see it:</strong> documents show up automatically under &ldquo;My Events&rdquo; in the attendee detail view — with an <strong>inline preview</strong> (for PDFs) and individual <strong>download</strong>.<br /><br />
              <strong>Effect:</strong> you can keep adding or replacing documents <strong>after the event has gone live</strong> — attendees always see the latest version.<br /><br />
              <strong>Tip:</strong> for documents only meant for organizers (e.g. caterer contact list), use a shared <strong>Teams/SharePoint folder outside DEX</strong>, because every attendee has read access here.
            </>
          )} />
        </div>
        <p className="dex-ui-help" style={{ marginTop: 0, marginBottom: 12 }}>
          {isDe
            ? <>Zum Beispiel <strong>Agenda</strong>, <strong>Anfahrt</strong>, <strong>Hausordnung</strong>, <strong>Packliste</strong>, <strong>Briefing</strong> oder <strong>Einverständniserklärung</strong>. Teilnehmer finden alles unter &bdquo;Meine Events&ldquo; — als Vorschau (PDF) und zum Herunterladen. Alle Teilnehmer haben Lesezugriff.</>
            : <>For example the <strong>agenda</strong>, <strong>directions</strong>, <strong>house rules</strong>, a <strong>packing list</strong>, the <strong>briefing</strong> or a <strong>consent form</strong>. Attendees find everything under &bdquo;My Events&ldquo; — as a preview (PDF) and for download. Every attendee has read access.</>}
        </p>

        {documents.length > 0 && (
          <div className="dex-ui-card" style={{ padding: '4px 6px', marginBottom: 12 }}>
            {documents.map((doc, idx) => (
              <div key={idx} className="dex-ui-row dex-ui-row--bordered">
                <span style={{ color: 'var(--dex-gray-500)', display: 'inline-flex', flexShrink: 0 }}><FileText size={18} /></span>
                <div className="dex-ui-row-main">
                  <div className="dex-ui-row-title" title={doc.name}>{doc.name}</div>
                  {doc.size > 0 && <div className="dex-ui-row-sub">{(doc.size / 1024).toFixed(0)} KB</div>}
                </div>
                {/* v31.2: Frisch gewählt (noch ohne URL) heißt: wird erst beim Speichern hochgeladen. */}
                {doc.file && <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'Neu' : 'New'}</span>}
                <div className="dex-ui-row-actions">
                  <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" title={t('general.delete')} aria-label={t('general.delete')}
                    onClick={() => setDocuments(documents.filter((_, i) => i !== idx))}>
                    <X size={16} />
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* v31.2: Der Erklärtext versprach Drag & Drop, es gab aber nur den Datei-Dialog.
            Die Zone nimmt jetzt beides und führt es durch denselben Pfad (addFiles). */}
        <label
          className="dex-ui-empty"
          style={{ display: 'block', cursor: 'pointer', padding: '22px 16px', transition: 'border-color 0.18s ease, background 0.18s ease',
            borderColor: dragOver ? 'var(--dex-green, #86bc25)' : undefined, background: dragOver ? 'rgba(134,188,37,0.08)' : undefined }}
          onDragOver={e => { e.preventDefault(); if (!dragOver) setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={e => { e.preventDefault(); setDragOver(false); addFiles(e.dataTransfer.files); }}
        >
          <span className="dex-ui-empty-icon"><Download size={20} /></span>
          <div className="dex-ui-empty-title">{isDe ? 'Dateien hier ablegen' : 'Drop files here'}</div>
          <div style={{ marginBottom: 12 }}>
            {isDe ? 'Mehrere auf einmal — per Drag & Drop oder Mehrfachauswahl im Datei-Dialog.' : 'Several at once — via drag & drop or multi-select in the file dialog.'}
          </div>
          <span className="btn btn-outline dex-ui-btn-sm" style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
            <Plus size={14} /> {t('create.documents.upload')}
          </span>
          <input type="file" multiple style={{ display: 'none' }}
            onChange={(e) => { addFiles(e.target.files); e.target.value = ''; }} />
        </label>
      </div>

      {/* v11.0: Teilnehmer-Upload-Toggle. Default OFF — wird nur
          bei expliziter Aktivierung in „Meine Events" als Upload-
          Bereich für die Anmeldung sichtbar. Anzeigename und
          Hinweistext sind frei konfigurierbar. */}
      {/* v31.2: Ja/Nein als Schalter-Zeile mit einer Zeile Folge. Die abhängigen Felder
          bleiben AUS-geschaltet sichtbar (gedämpft, gesperrt), damit der Organizer sieht,
          was er beim Einschalten bekommt — vorher tauchten sie erst nach dem Klick auf. */}
      <div className="dex-ui-section">
        <div className="dex-ui-section-title">
          <StepBadge n={34} />
          {isDe ? 'Upload durch Teilnehmer' : 'Upload by attendees'}
        </div>
        <label className={cx('dex-ui-toggle-row', allowAttendeeUpload && 'is-active')}>
          <input type="checkbox" checked={allowAttendeeUpload} onChange={e => setAllowAttendeeUpload(e.target.checked)} />
          <span className="dex-ui-toggle-row-body">
            <span className="dex-ui-toggle-row-title">
              {isDe ? 'Teilnehmer dürfen eine eigene Datei hochladen' : 'Attendees may upload a file of their own'}
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
            </span>
            <span className="dex-ui-toggle-row-desc">
              {isDe
                ? <>Dann sehen sie nach der Anmeldung in &bdquo;Meine Events&ldquo; einen Upload-Block mit deiner Beschriftung — z.B. für einen Reisekostenbeleg oder eine unterschriebene Einverständniserklärung. Die Datei landet an ihrer Teilnehmer-Zeile; du siehst sie im Organizer Center. Standard: aus.</>
                : <>Then, after registering, they see an upload block in &bdquo;My Events&ldquo; with your label — e.g. for a travel-expense receipt or a signed consent form. The file attaches to their attendee row; you see it in the organizer center. Default: off.</>}
            </span>
          </span>
        </label>

        <div className={cx('dex-ui-card', !allowAttendeeUpload && 'dex-ui-card--muted')} style={{ marginTop: 12 }}>
          <div className="dex-ui-grid-2">
            <div className="dex-ui-field">
              <label className="dex-ui-label" htmlFor="dex-doc-attendee-upload-label">
                {isDe ? 'Wie soll der Upload-Block heißen?' : 'What should the upload block be called?'}
              </label>
              <input
                id="dex-doc-attendee-upload-label"
                className="form-input"
                value={attendeeUploadLabel}
                disabled={!allowAttendeeUpload}
                onChange={e => setAttendeeUploadLabel(e.target.value)}
                placeholder={isDe ? 'z.B. Reisekostenbeleg, Datenschutz-Erklärung' : 'e.g. Travel receipt, privacy form'}
                maxLength={80}
              />
              <div className="dex-ui-help">{isDe ? <>Steht als Überschrift über dem Upload-Feld in &bdquo;Meine Events&ldquo;.</> : <>Shown as the heading above the upload field in &bdquo;My Events&ldquo;.</>}</div>
            </div>
            <div className="dex-ui-field">
              <label className="dex-ui-label" htmlFor="dex-doc-attendee-upload-hint">
                {isDe ? 'Hinweis für Teilnehmer' : 'Hint for attendees'}
                <span className="dex-ui-label-optional">(optional)</span>
              </label>
              <input
                id="dex-doc-attendee-upload-hint"
                className="form-input"
                value={attendeeUploadHint}
                disabled={!allowAttendeeUpload}
                onChange={e => setAttendeeUploadHint(e.target.value)}
                placeholder={isDe ? 'z.B. Bitte unterschrieben hochladen' : 'e.g. Please upload signed'}
                maxLength={240}
              />
              <div className="dex-ui-help">{isDe ? 'Steht unter der Überschrift — z.B. was unterschrieben sein muss oder bis wann.' : 'Shown below the heading — e.g. what has to be signed or by when.'}</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
