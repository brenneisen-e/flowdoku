/* QrEditModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 13809-14075 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { QrEmailOverride, buildQrBlockHtml, qrEmailDefaults } from '../../../services/EmailTemplates';
import { MailHeaderImage, resolveMailHeaderImage } from '../../../utils/mailHeaderImage';
import MailHeaderImageChooser from '../../admin/MailHeaderImageChooser';
import { SAMPLE_QR_ID } from '../../admin/adminConstants';
import { HtmlEditorModal } from '../../HtmlEditorModal';
import { Check, ChevronDown, ChevronLeft, QrCode, Send } from '../../Icons';
import { InfoTooltip } from '../../InfoTooltip';
import { cx } from '../../dexUi';
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
        // v31.2: Aufklapper „Block neben dem QR-Code anpassen" — standardmäßig
        // zu, weil Sprache und Hinweis selten geändert werden. Einziger Hook der
        // Komponente; sie hat keinen frühen Return.
        const [fineOpen, setFineOpen] = React.useState(false);
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
        // v31.0: derselbe Default wie beim Öffnen (eigenes Mail-Logo → volle
        // Breite), sonst gälte der Editor sofort als „ungespeichert geändert".
        const savedHeaderImage = resolveMailHeaderImage(savedOv && savedOv.headerImage, qrTgt.emailTemplateOverrides, qrTgt.mailImageBase64);
        // v31.0: eigene Teilnehmer-ID in der Vorschau, sonst Beispiel-ID.
        const ownId = ((): number => {
          const me = (currentUser.email || '').toLowerCase();
          const mine = me ? registrations.find(r => (r.ParticipantEmail || '').toLowerCase() === me) : undefined;
          return mine && mine.TeilnehmerID ? mine.TeilnehmerID : SAMPLE_QR_ID;
        })();
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
        // v31.2: Die Spalte links ist auf das Nötige reduziert: Zähler,
        // Testmail (sekundär), Versand (der einzige Primär-Knopf) und Zurück
        // als Textknopf. Der Satz zur Live-Vorschau steht im Tooltip neben der
        // Überschrift — als Fließtext hat er die Knöpfe auseinandergeschoben.
        const leftPanel = (
          <div className="dex-ui-stack">
            <div className="dex-ui-inline" style={{ justifyContent: 'space-between' }}>
              <span style={{ fontWeight: 700, fontSize: '0.92rem' }}>{isDe ? 'QR-Codes versenden' : 'Send QR codes'}</span>
              <InfoTooltip placement="right" text={isDe
                ? 'Die Live-Vorschau rechts zeigt deinen aktuellen Text. Die Testmail nutzt ebenfalls den aktuellen Text — der Versand an die Teilnehmer immer den gespeicherten.'
                : 'The live preview on the right shows your current text. The test email also uses the current text — sending to participants always uses the saved one.'} />
            </div>
            {/* v29.26: Bei einem Sub-Event-Ziel gehören die Zähler und der
                Massen-Versand zum FALSCHEN Event (hier ist die Liste des
                geöffneten Events geladen) — stattdessen sagt ein Hinweis,
                wo der Versand mit diesem Text stattfindet. */}
            {isSubTarget && (
              <div className="dex-ui-callout dex-ui-callout--info" style={{ fontSize: '0.76rem' }}>
                <span>
                  {isDe
                    ? <>Du gestaltest die QR-Mail des Sub-Events <strong>{qrTgt.title}</strong>. Der gespeicherte Text gilt für dessen manuellen Versand (Sub-Event im Organizer Center öffnen → &bdquo;QR-Codes versenden&ldquo;) und den automatischen Versand bei neuen Anmeldungen.</>
                    : <>You are customizing the QR email of the sub-event <strong>{qrTgt.title}</strong>. The saved text applies to its manual sending (open the sub-event in the Organizer Center → “Send QR codes”) and the automatic send for new registrations.</>}
                </span>
              </div>
            )}
            {!isSubTarget && (
              <div className="dex-ui-inline">
                <span className="dex-ui-pill dex-ui-pill--green"><strong>{noCodeCount}</strong> {isDe ? 'ohne Code' : 'without code'}</span>
                <span className="dex-ui-pill dex-ui-pill--gray"><strong>{withCodeCount}</strong> {isDe ? 'mit Code' : 'with code'}</span>
              </div>
            )}
            <div>
              <button
                type="button"
                className="btn btn-outline dex-ui-btn-sm"
                disabled={isSendingQR}
                onClick={() => { qrTestSendAction({ subject: qrEditSubject, heading: qrEditHeading, subheading: qrEditSubheading, bodyHtml: qrEditBody, headerImage: { ...qrHeaderImage } }, isSubTarget ? qrTgt : undefined).catch(() => { /* */ }); }}
                style={{ width: '100%' }}
              >
                {isDe ? 'Testmail an Organisatoren' : 'Test email to organizers'}
              </button>
              <div className="dex-ui-help">{isDe ? 'Nutzt deinen aktuellen Text — auch ungespeichert.' : 'Uses your current text — even if unsaved.'}</div>
            </div>
            {!isSubTarget && (
              <button
                type="button"
                className="btn btn-primary"
                disabled={isSendingQR || qrEditDirty || noCodeCount === 0}
                onClick={() => { qrFullSendAction().catch(() => { /* */ }); }}
                style={{ width: '100%' }}
                title={qrEditDirty ? (isDe ? 'Erst speichern — der Versand nutzt den gespeicherten Text.' : 'Save first — sending uses the saved text.') : undefined}
              >
                <Send size={16} />
                {isSendingQR
                  ? `${isDe ? 'Versende' : 'Sending'}… (${qrSentCount})`
                  : (noCodeCount === 0
                    ? (isDe ? 'Alle haben ihren QR-Code' : 'Everyone has their QR code')
                    : (isDe ? `An ${noCodeCount} Teilnehmer senden` : `Send to ${noCodeCount} participant${noCodeCount === 1 ? '' : 's'}`))}
              </button>
            )}
            {qrEditDirty && (
              <div className="dex-ui-callout dex-ui-callout--warn" style={{ fontSize: '0.76rem' }}>
                <span>
                  {isDe
                    ? <>Noch nicht gespeichert — erst &bdquo;Für dieses Event speichern&ldquo; klicken, dann an die Teilnehmer senden.</>
                    : <>Not saved yet — click &ldquo;Save for this event&rdquo; first, then send to participants.</>}
                </span>
              </div>
            )}
            {qrSendResult && (
              <div className="dex-ui-muted" style={{ lineHeight: 1.5, borderTop: '1px solid var(--dex-gray-200)', paddingTop: 8 }}>
                {qrSendResult}
              </div>
            )}
            <button
              type="button"
              className="dex-ui-textbtn dex-ui-textbtn--muted"
              disabled={isSendingQR}
              onClick={closeQrMailEditor}
              style={{ alignSelf: 'flex-start', marginLeft: -8 }}
            >
              <ChevronLeft size={14} />
              {isDe ? 'Zurück zum Versand-Dialog' : 'Back to the send dialog'}
            </button>
          </div>
        );
        // v31.2: Der Kopf über dem Editor war ein Textblock aus fünf Sätzen
        // („sieht völlig überfordernd aus", Screenshot 07.09.2026). Sichtbar
        // bleibt EIN Satz zum festen Platzhalter; alles Weitere (Platzhalter-
        // Liste, Geltung für manuellen UND automatischen Versand) steht im
        // Tooltip — keine Aussage entfällt, sie steht nur nicht mehr im Weg.
        const headerExtra = (
          <div className="dex-ui-stack">
            <div className="dex-ui-callout dex-ui-callout--neutral">
              <span className="dex-ui-callout-icon"><QrCode size={16} /></span>
              <span style={{ flex: 1, minWidth: 0 }}>
                {isDe
                  ? <>Der Platzhalter <code>{'{{QR_BLOCK}}'}</code> ist der persönliche QR-Code — er bleibt immer in der Mail und lässt sich im Text verschieben.</>
                  : <>The placeholder <code>{'{{QR_BLOCK}}'}</code> is the personal QR code — it always stays in the email and can be moved within the text.</>}
                {' '}
                <InfoTooltip placement="bottom" text={isDe
                  ? <><strong>Was der Block zeigt:</strong> den QR-Code mit Name und Event als Klartext. Fehlt <code>{'{{QR_BLOCK}}'}</code> im Text, wird er beim Versand automatisch ans Ende gesetzt.<br /><strong>Weitere Platzhalter:</strong> <code>{'{{Vorname}}'}</code>, <code>{'{{Name}}'}</code>, <code>{'{{EventTitle}}'}</code>.<br /><strong>Auswirkung:</strong> Der gespeicherte Text gilt für alle QR-Mails dieses Events — manueller Versand UND automatischer Versand bei neuen Anmeldungen.</>
                  : <><strong>What the block shows:</strong> the QR code with name and event as plain text. If <code>{'{{QR_BLOCK}}'}</code> is missing, it is appended automatically when sending.<br /><strong>More placeholders:</strong> <code>{'{{Vorname}}'}</code>, <code>{'{{Name}}'}</code>, <code>{'{{EventTitle}}'}</code>.<br /><strong>Effect:</strong> the saved text applies to all QR emails of this event — manual sending AND the automatic send for new registrations.</>} />
              </span>
            </div>
            {/* v30.52: Kopf-Bild — dieselbe Auswahl wie in Massen- und
                Einladungsmail. Hier wird sie MITGESPEICHERT, weil die QR-Mail
                auch automatisch rausgeht. */}
            <MailHeaderImageChooser
              value={qrHeaderImage}
              onChange={setQrHeaderImage}
              eventPhotoB64={qrEventPhotoB64}
              disabled={qrEditSaving || isSendingQR}
              onCrop={() => setComposerCrop('qr')}
              isDe={isDe}
            />
            {/* v31.2: Sprache und Hinweis betreffen nur den Block NEBEN dem
                Code — selten angefasst, deshalb im Aufklapper (standardmäßig
                zu). Die Vorschau rechts zeigt jede Änderung sofort. */}
            <div>
              <button
                type="button"
                className={cx('dex-ui-disclosure', fineOpen && 'is-open')}
                aria-expanded={fineOpen}
                onClick={() => setFineOpen(o => !o)}
              >
                <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                {isDe ? 'Block neben dem QR-Code anpassen' : 'Customize the block next to the QR code'}
                <span className="dex-ui-disclosure-count">2</span>
              </button>
              {fineOpen && (
                <div className="dex-ui-disclosure-body">
                  {/* v30.60: Sprache des Blocks NEBEN dem Code. Er trägt „Name",
                      „ID" und den Hinweis zur Nummer und stand bisher immer auf
                      Deutsch — auch unter einer englischen Mail. Voreinstellung
                      bleibt „wie die Mail-Sprache des Events", damit hier kein
                      zweiter Schalter für dieselbe Frage entsteht. */}
                  <div className="dex-ui-field">
                    <div className="dex-ui-label">
                      {isDe ? 'In welcher Sprache steht der Block neben dem Code?' : 'Which language should the block next to the code use?'}
                    </div>
                    <div className="dex-ui-inline">
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
                            className={cx('dex-ui-chip', on && 'is-active')}
                            aria-pressed={on}
                            disabled={qrEditSaving || isSendingQR}
                            onClick={() => {
                              setQrBlockLang(opt.v);
                              // Vorschau sofort mitziehen — sonst wählt man eine
                              // Sprache und sieht rechts weiter die alte.
                              if (qrEditSampleImg) {
                                const myNm = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
                                setQrEditSampleBlock(buildQrBlockHtml(qrEditSampleImg, myNm, ownId,opt.v || qrTgt.emailLanguage || 'EN', qrBlockNote));
                              }
                            }}
                          >{isDe ? opt.de : opt.en}</button>
                        );
                      })}
                    </div>
                    <div className="dex-ui-help">
                      {isDe
                        ? 'Betrifft „Name", „ID" und den Hinweis unter der Nummer — nicht deinen Mailtext.'
                        : 'Affects “Name”, “ID” and the note below the number — not your email copy.'}
                    </div>
                  </div>
                  {/* v30.61: Der Hinweis unter der ID stand fest im Code, während
                      der Mailtext direkt darüber frei ist. „Am Einlass" heißt beim
                      B2Run „bei der Trikot- und Startnummernübergabe" und bei einer
                      Konferenz „an der Registrierung". */}
                  <div className="dex-ui-field">
                    <label className="dex-ui-label" htmlFor="dex-qr-block-note">
                      {isDe ? 'Was steht unter der Teilnehmer-ID?' : 'What should appear below the participant ID?'}
                      <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                    </label>
                    <input
                      id="dex-qr-block-note"
                      type="text"
                      className="dex-ui-input dex-ui-input--sm"
                      value={qrBlockNote}
                      disabled={qrEditSaving || isSendingQR}
                      onChange={e => {
                        setQrBlockNote(e.target.value);
                        if (qrEditSampleImg) {
                          const myNm = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
                          setQrEditSampleBlock(buildQrBlockHtml(qrEditSampleImg, myNm, ownId,qrBlockLang || qrTgt.emailLanguage || 'EN', e.target.value));
                        }
                      }}
                      placeholder={isDe
                        ? 'Leer = „Falls der Scan nicht klappt: einfach diese Nummer am Einlass nennen."'
                        : 'Empty = the default note'}
                    />
                    <div className="dex-ui-help">
                      {isDe
                        ? 'Erscheint in der Mail direkt unter der Teilnehmer-ID. Leer lassen für den Standardsatz; ein einzelner Bindestrich (-) blendet den Hinweis ganz aus.'
                        : 'Shown in the email right below the participant ID. Leave empty for the default; a single hyphen (-) hides the note entirely.'}
                    </div>
                  </div>
                </div>
              )}
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

