/* QrEditModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13809-14075 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { QrEmailOverride, buildQrBlockHtml, qrEmailDefaults } from '../../../services/EmailTemplates';
import { MailHeaderImage, normalizeMailHeaderImage } from '../../../utils/mailHeaderImage';
import MailHeaderImageChooser from '../../admin/MailHeaderImageChooser';
import { SAMPLE_QR_ID } from '../../admin/adminConstants';
import { HtmlEditorModal } from '../../HtmlEditorModal';
import { Check } from '../../Icons';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

export interface QrEditModalProps {
  closeQrMailEditor: () => void;
  currentUser: import("../../../types/index").User;
  getQrMailOverride: (ev: DeloitteEvent | null) => QrEmailOverride | undefined;
  isDe: boolean;
  isSendingQR: boolean;
  qrBlockLang: "" | "DE" | "EN";
  qrBlockNote: string;
  qrEditBody: string;
  qrEditHeading: string;
  qrEditOpen: boolean;
  qrEditSampleBlock: string;
  qrEditSampleImg: string;
  qrEditSaving: boolean;
  qrEditSubheading: string;
  qrEditSubject: string;
  qrEditTarget: DeloitteEvent;
  qrEventPhotoB64: string;
  qrFullSendAction: () => Promise<void>;
  qrHeaderImage: MailHeaderImage;
  qrSendResult: string;
  qrSentCount: number;
  qrTestSendAction: (liveOverride?: QrEmailOverride, target?: DeloitteEvent) => Promise<void>;
  registrations: SPRegistration[];
  saveQrMailOverride: () => Promise<void>;
  selectedEvent: DeloitteEvent;
  setComposerCrop: React.Dispatch<React.SetStateAction<"invite" | "massmail" | "qr">>;
  setQrBlockLang: React.Dispatch<React.SetStateAction<"" | "DE" | "EN">>;
  setQrBlockNote: React.Dispatch<React.SetStateAction<string>>;
  setQrEditBody: React.Dispatch<React.SetStateAction<string>>;
  setQrEditHeading: React.Dispatch<React.SetStateAction<string>>;
  setQrEditSampleBlock: React.Dispatch<React.SetStateAction<string>>;
  setQrEditSubheading: React.Dispatch<React.SetStateAction<string>>;
  setQrEditSubject: React.Dispatch<React.SetStateAction<string>>;
  setQrHeaderImage: React.Dispatch<React.SetStateAction<MailHeaderImage>>;
}

export const QrEditModal: React.FC<QrEditModalProps> = (p) => {
  const { closeQrMailEditor, currentUser, getQrMailOverride, isDe, isSendingQR, qrBlockLang, qrBlockNote, qrEditBody, qrEditHeading, qrEditOpen, qrEditSampleBlock, qrEditSampleImg, qrEditSaving, qrEditSubheading, qrEditSubject, qrEditTarget, qrEventPhotoB64, qrFullSendAction, qrHeaderImage, qrSendResult, qrSentCount, qrTestSendAction, registrations, saveQrMailOverride, selectedEvent, setComposerCrop, setQrBlockLang, setQrBlockNote, setQrEditBody, setQrEditHeading, setQrEditSampleBlock, setQrEditSubheading, setQrEditSubject, setQrHeaderImage } = p;
        // v29.26: Editor-Ziel — das Event selbst ODER ein vom Hauptevent aus
        // geöffnetes Sub-Event (qrEditTarget). Alle Texte/Vergleiche laufen
        // gegen das Ziel; die Versand-Spalte links gehört dagegen zum
        // GEÖFFNETEN Event und wird bei einem Sub-Ziel durch einen Hinweis
        // ersetzt (dessen Teilnehmerliste ist hier nicht geladen).
        const qrTgt = qrEditTarget || selectedEvent;
        const isSubTarget = qrTgt.id !== selectedEvent.id;
        const myName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
        const previewVars: Record<string, string> = {
          EventTitle: qrTgt.title,
          Vorname: currentUser.firstName || myName,
          Name: myName,
        };
        const customLogo = (() => {
          try {
            const o = JSON.parse(qrTgt.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })();
        const resolvePlain = (s: string): string => s
          .replace(/\{\{EventTitle\}\}/g, qrTgt.title)
          .replace(/\{\{Vorname\}\}/g, previewVars.Vorname)
          .replace(/\{\{Name\}\}/g, myName);
        const def = qrEmailDefaults(qrTgt.emailLanguage || 'EN');
        // Versand-Spalte links: ungespeicherte Änderungen sperren den
        // Massen-Versand (der nutzt den GESPEICHERTEN Text) — der Test an
        // mich nutzt bewusst den aktuellen Editor-Text (Test = Vorschau).
        const savedOv = getQrMailOverride(qrTgt);
        const savedSubject = (savedOv && savedOv.subject) || def.subject;
        const savedHeading = (savedOv && savedOv.heading) || def.heading;
        const savedSubheading = (savedOv && savedOv.subheading) || def.subheading;
        const savedBody = (savedOv && savedOv.bodyHtml) || def.body;
        const savedHeaderImage = normalizeMailHeaderImage(savedOv && savedOv.headerImage);
        const qrEditDirty = qrEditSubject.trim() !== savedSubject.trim()
          || qrEditHeading.trim() !== savedHeading.trim()
          || qrEditSubheading.trim() !== savedSubheading.trim()
          || qrEditBody.trim() !== savedBody.trim()
          // v30.52: Auch eine geänderte Kopf-Bild-Einstellung ist eine
          // ungespeicherte Änderung — sonst sperrt der Versand nicht, obwohl
          // er den alten Stand verschicken würde.
          || qrHeaderImage.hero !== savedHeaderImage.hero
          || qrHeaderImage.width !== savedHeaderImage.width
          || qrHeaderImage.paddingV !== savedHeaderImage.paddingV
          || qrHeaderImage.paddingH !== savedHeaderImage.paddingH;
        const noCodeCount = registrations.filter(r => r.Status === 'Angemeldet').length;
        const withCodeCount = registrations.filter(r => r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length;
        const leftPanel = (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <div style={{ fontWeight: 700, fontSize: '0.92rem' }}>{isDe ? 'QR-Code-Versand' : 'QR code sending'}</div>
            {/* v29.26: Bei einem Sub-Event-Ziel gehören die Zähler und der
                Massen-Versand zum FALSCHEN Event (hier ist die Liste des
                geöffneten Events geladen) — stattdessen sagt ein Hinweis,
                wo der Versand mit diesem Text stattfindet. */}
            {isSubTarget && (
              <div style={{ background: 'rgba(134,188,37,0.10)', border: '1px solid var(--dex-green, #86bc25)', borderRadius: 'var(--dex-radius)', padding: '8px 10px', fontSize: '0.74rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {isDe
                  ? <>Du gestaltest die QR-Mail des Sub-Events <strong>{qrTgt.title}</strong>. Der gespeicherte Text gilt für dessen manuellen Versand (Sub-Event im Organizer Center öffnen → &bdquo;QR-Codes versenden&ldquo;) und den automatischen Versand bei neuen Anmeldungen.</>
                  : <>You are customizing the QR email of the sub-event <strong>{qrTgt.title}</strong>. The saved text applies to its manual sending (open the sub-event in the Organizer Center → “Send QR codes”) and the automatic send for new registrations.</>}
              </div>
            )}
            {!isSubTarget && (
            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
              <span style={{ background: '#eef6e3', color: 'var(--dex-green-dark, #4a7c1f)', borderRadius: 12, padding: '3px 10px', fontSize: '0.74rem', fontWeight: 600 }}>
                <strong>{noCodeCount}</strong> {isDe ? 'ohne Code' : 'without code'}
              </span>
              <span style={{ background: 'var(--dex-gray-100, #f0f0f0)', color: 'var(--dex-gray-600)', borderRadius: 12, padding: '3px 10px', fontSize: '0.74rem', fontWeight: 600 }}>
                <strong>{withCodeCount}</strong> {isDe ? 'mit Code' : 'with code'}
              </span>
            </div>
            )}
            <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', lineHeight: 1.5 }}>
              {isDe
                ? 'Die Live-Vorschau rechts zeigt deinen aktuellen Text. Der Test an dich nutzt ebenfalls den aktuellen Text — der Versand an die Teilnehmer immer den gespeicherten.'
                : 'The live preview on the right shows your current text. The test to yourself also uses the current text — sending to participants always uses the saved one.'}
            </div>
            <button
              type="button"
              className="btn btn-outline"
              disabled={isSendingQR}
              onClick={() => { qrTestSendAction({ subject: qrEditSubject, heading: qrEditHeading, subheading: qrEditSubheading, bodyHtml: qrEditBody, headerImage: { ...qrHeaderImage } }, isSubTarget ? qrTgt : undefined).catch(() => { /* */ }); }}
              style={{ fontSize: '0.82rem', width: '100%' }}
            >
              {isDe ? 'Test an Organisatoren (aktueller Text)' : 'Test to organizers (current text)'}
            </button>
            {!isSubTarget && (
            <button
              type="button"
              className="btn btn-primary"
              disabled={isSendingQR || qrEditDirty || noCodeCount === 0}
              onClick={() => { qrFullSendAction().catch(() => { /* */ }); }}
              style={{ fontSize: '0.82rem', width: '100%', fontWeight: 700 }}
              title={qrEditDirty ? (isDe ? 'Erst speichern — der Versand nutzt den gespeicherten Text.' : 'Save first — sending uses the saved text.') : undefined}
            >
              {isSendingQR
                ? `${isDe ? 'Versende' : 'Sending'}… (${qrSentCount})`
                : (noCodeCount === 0
                  ? (isDe ? 'Alle haben ihren QR-Code' : 'Everyone has their QR code')
                  : (isDe ? `An ${noCodeCount} Teilnehmer senden` : `Send to ${noCodeCount} participant${noCodeCount === 1 ? '' : 's'}`))}
            </button>
            )}
            {qrEditDirty && (
              <div style={{ background: '#fff3e0', border: '1px solid #ed8b00', borderRadius: 'var(--dex-radius)', padding: '8px 10px', fontSize: '0.72rem', color: '#7a4a00', lineHeight: 1.5 }}>
                {isDe
                  ? 'Ungespeicherte Änderungen — erst „Für dieses Event speichern" klicken, dann an die Teilnehmer senden.'
                  : 'Unsaved changes — click "Save for this event" first, then send to participants.'}
              </div>
            )}
            {qrSendResult && (
              <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, borderTop: '1px solid var(--dex-gray-200)', paddingTop: 8 }}>
                {qrSendResult}
              </div>
            )}
            <button
              type="button"
              className="btn btn-secondary"
              disabled={isSendingQR}
              onClick={closeQrMailEditor}
              style={{ fontSize: '0.78rem', width: '100%', marginTop: 4 }}
            >
              {isDe ? 'Zurück zum Versand-Modal' : 'Back to the send dialog'}
            </button>
          </div>
        );
        const headerExtra = (
          <div style={{ padding: 12, background: 'var(--dex-gray-50, #fafafa)', border: '1px solid var(--dex-gray-200)', borderRadius: 'var(--dex-radius)', marginBottom: 4, fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
            {isDe
              ? <><strong>Fester Bestandteil:</strong> Der Platzhalter <code>{'{{QR_BLOCK}}'}</code> steht für den persönlichen QR-Code mit Name + Event als Klartext — er lässt sich verschieben, aber nicht entfernen (fehlt er im Text, wird der Block beim Versand automatisch ans Ende gesetzt). Verfügbare Platzhalter: <code>{'{{Vorname}}'}</code>, <code>{'{{Name}}'}</code>, <code>{'{{EventTitle}}'}</code>. <strong>Der gespeicherte Text gilt für alle QR-Mails dieses Events</strong> — manueller Versand UND automatischer Versand bei neuen Anmeldungen.</>
              : <><strong>Fixed element:</strong> the placeholder <code>{'{{QR_BLOCK}}'}</code> represents the personal QR code with name + event as plain text — it can be moved but not removed (if missing, the block is appended automatically when sending). Available placeholders: <code>{'{{Vorname}}'}</code>, <code>{'{{Name}}'}</code>, <code>{'{{EventTitle}}'}</code>. <strong>The saved text applies to all QR emails of this event</strong> — manual sending AND the automatic send for new registrations.</>}
            {/* v30.52: Kopf-Bild — dieselbe Auswahl wie in Massen- und
                Einladungsmail. Hier wird sie MITGESPEICHERT, weil die QR-Mail
                auch automatisch rausgeht. */}
            <div style={{ marginTop: 10, paddingTop: 8, borderTop: '1px dashed var(--dex-gray-200)' }}>
              <MailHeaderImageChooser
                value={qrHeaderImage}
                onChange={setQrHeaderImage}
                eventPhotoB64={qrEventPhotoB64}
                disabled={qrEditSaving || isSendingQR}
                onCrop={() => setComposerCrop('qr')}
                isDe={isDe}
              />
            </div>
            {/* v30.60: Sprache des Blocks NEBEN dem Code. Er trägt „Name",
                „ID" und den Hinweis zur Nummer und stand bisher immer auf
                Deutsch — auch unter einer englischen Mail. Voreinstellung
                bleibt „wie die Mail-Sprache des Events", damit hier kein
                zweiter Schalter für dieselbe Frage entsteht. */}
            <div style={{ marginTop: 10, paddingTop: 8, borderTop: '1px dashed var(--dex-gray-200)' }}>
              <div style={{ fontWeight: 600, marginBottom: 6 }}>
                {isDe ? 'Sprache neben dem QR-Code' : 'Language next to the QR code'}
              </div>
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                {([
                  { v: '' as const, de: `Wie die Mail-Sprache (${(qrTgt.emailLanguage || 'EN').toUpperCase()})`, en: `Follow the event language (${(qrTgt.emailLanguage || 'EN').toUpperCase()})` },
                  { v: 'DE' as const, de: 'Immer Deutsch', en: 'Always German' },
                  { v: 'EN' as const, de: 'Immer Englisch', en: 'Always English' },
                ]).map(opt => {
                  const on = qrBlockLang === opt.v;
                  return (
                    <button
                      key={opt.v || 'auto'}
                      type="button"
                      disabled={qrEditSaving || isSendingQR}
                      onClick={() => {
                        setQrBlockLang(opt.v);
                        // Vorschau sofort mitziehen — sonst wählt man eine
                        // Sprache und sieht rechts weiter die alte.
                        if (qrEditSampleImg) {
                          const myNm = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
                          setQrEditSampleBlock(buildQrBlockHtml(qrEditSampleImg, myNm, SAMPLE_QR_ID, opt.v || qrTgt.emailLanguage || 'EN', qrBlockNote));
                        }
                      }}
                      style={{
                        padding: '4px 12px', borderRadius: 999, cursor: 'pointer',
                        border: `1px solid ${on ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                        background: on ? 'var(--dex-green, #86bc25)' : '#fff',
                        color: on ? '#fff' : 'var(--dex-gray-600)',
                        fontSize: '0.74rem', fontWeight: 600,
                      }}
                    >{isDe ? opt.de : opt.en}</button>
                  );
                })}
              </div>
              <div style={{ marginTop: 6, color: 'var(--dex-gray-500)' }}>
                {isDe
                  ? 'Betrifft „Name", „ID" und den Hinweis unter der Nummer — nicht deinen Mailtext.'
                  : 'Affects “Name”, “ID” and the note below the number — not your email copy.'}
              </div>
              {/* v30.61: Der Hinweis unter der ID stand fest im Code, während
                  der Mailtext direkt darüber frei ist. „Am Einlass" heißt beim
                  B2Run „bei der Trikot- und Startnummernübergabe" und bei einer
                  Konferenz „an der Registrierung". */}
              <div style={{ marginTop: 10 }}>
                <label style={{ display: 'block', fontWeight: 600, marginBottom: 4 }}>
                  {isDe ? 'Hinweis unter der Teilnehmer-ID' : 'Note below the participant ID'}
                </label>
                <input
                  type="text"
                  value={qrBlockNote}
                  disabled={qrEditSaving || isSendingQR}
                  onChange={e => {
                    setQrBlockNote(e.target.value);
                    if (qrEditSampleImg) {
                      const myNm = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
                      setQrEditSampleBlock(buildQrBlockHtml(qrEditSampleImg, myNm, SAMPLE_QR_ID, qrBlockLang || qrTgt.emailLanguage || 'EN', e.target.value));
                    }
                  }}
                  placeholder={isDe
                    ? 'Leer = „Falls der Scan nicht klappt: einfach diese Nummer am Einlass nennen."'
                    : 'Empty = the default note'}
                  style={{ width: '100%', padding: '6px 10px', border: '1px solid var(--dex-gray-200)', borderRadius: 6, fontSize: '0.82rem' }}
                />
                <div style={{ marginTop: 4, color: 'var(--dex-gray-500)' }}>
                  {isDe
                    ? 'Leer lassen für den Standardsatz. Ein einzelner Bindestrich (-) blendet den Hinweis ganz aus.'
                    : 'Leave empty for the default. A single hyphen (-) hides the note entirely.'}
                </div>
              </div>
            </div>
          </div>
        );
        return (
          <HtmlEditorModal
            open={qrEditOpen}
            onClose={() => { if (!qrEditSaving && !isSendingQR) closeQrMailEditor(); }}
            title={isDe ? `QR-Mail anpassen: ${qrTgt.title}` : `Customize QR email: ${qrTgt.title}`}
            leftPanel={leftPanel}
            value={qrEditBody}
            onChange={setQrEditBody}
            previewMode="email"
            emailSubject={qrEditSubject}
            onEmailSubjectChange={setQrEditSubject}
            emailHeading={qrEditHeading}
            onEmailHeadingChange={setQrEditHeading}
            emailSubheading={qrEditSubheading}
            onEmailSubheadingChange={setQrEditSubheading}
            emailHeadingColor="#86bc25"
            previewVars={previewVars}
            previewHtmlVars={{ QR_BLOCK: qrEditSampleBlock }}
            previewToLine={`${currentUser.email} ${isDe ? '(Beispiel — du selbst)' : '(example — yourself)'}`}
            previewSubjectLine={resolvePlain(qrEditSubject)}
            defaultBodyHtml={def.body}
            insertableVars={[
              { key: '{{Vorname}}', label: isDe ? 'Vorname' : 'First name' },
              { key: '{{Name}}', label: isDe ? 'Voller Name' : 'Full name' },
              { key: '{{EventTitle}}', label: isDe ? 'Event-Titel' : 'Event title' },
              { key: '{{QR_BLOCK}}', label: isDe ? 'QR-Code-Block (fix)' : 'QR code block (fixed)' },
            ]}
            imageBase64={(qrHeaderImage.hero === 'event' && qrEventPhotoB64) ? qrEventPhotoB64 : customLogo}
            imageWidth={qrHeaderImage.width}
            imagePaddingV={qrHeaderImage.paddingV}
            imagePaddingH={qrHeaderImage.paddingH}
            onImageWidthChange={(w) => setQrHeaderImage(p => ({ ...p, width: w }))}
            onImagePaddingVChange={(v) => setQrHeaderImage(p => ({ ...p, paddingV: v }))}
            onImagePaddingHChange={(h) => setQrHeaderImage(p => ({ ...p, paddingH: h }))}
            headerExtra={headerExtra}
            extraAction={{
              label: qrEditSaving
                ? (isDe ? 'Speichert…' : 'Saving…')
                : (isDe ? 'Für dieses Event speichern' : 'Save for this event'),
              onClick: saveQrMailOverride,
              disabled: qrEditSaving || !qrEditSubject.trim() || !qrEditBody.trim(),
              icon: <Check size={16} />,
            }}
          />
        );
};

