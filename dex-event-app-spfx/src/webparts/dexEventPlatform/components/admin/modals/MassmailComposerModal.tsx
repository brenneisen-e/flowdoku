/* MassmailComposerModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 15244-15482 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { replacePlaceholders, wrapTemplate } from '../../../services/EmailTemplates';
import { formatOrganizerList } from '../../../context/eventTextHelpers';
import { HtmlEditorModal } from '../../HtmlEditorModal';
import RecipientPicker from '../../admin/RecipientPicker';
import MailHeaderImageChooser from '../../admin/MailHeaderImageChooser';
import { Check, ChevronDown, Send, Users } from '../../Icons';
import { cx } from '../../dexUi';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';
import { MailHeaderImage } from '../../../utils/mailHeaderImage';
import { MassmailAudience } from '../adminTypes';
import { buildInlineImage, charsToKb } from '../../../utils/inlineMailImage';
// v31.10: Dieselbe Rechnung wie die Anmeldeseite — wer dort ausgeblendet ist,
// steht auch nicht im CC. Die Regel liegt in EINER Datei, nicht hier.
import { visibleOrganizerEmails } from '../../../utils/organizerVisibility';
import { PollComposerSection } from '../PollComposerSection';

export interface MassmailComposerModalProps {
  applyMassmailHero: (wrappedHtml: string) => string;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  emailBody: string;
  emailHeading: string;
  emailSending: boolean;
  emailSubject: string;
  eventServiceRef: EventService;
  isDe: boolean;
  // v31.9.6: Der gemeinsame Typ statt einer zweiten Aufzaehlung. Die Kopie
  // hier hat den neuen Wert `everyone` nicht mitbekommen — `tsc` hat es
  // gemeldet, sonst haette der Composer eine Auswahl bekommen, die er
  // nicht kennt.
  massmailAudience: MassmailAudience;
  massmailCc: string[];
  massmailDraftSaved: boolean;
  massmailEventPhotoB64: string;
  massmailHeaderImage: MailHeaderImage;
  massmailHeaderOpts: { imageWidth: number; imagePaddingV: number; imagePaddingH: number; };
  massmailPasteRaw: string;
  massmailStatuses: Set<string>;
  massmailSubheading: string;
  massmailTesting: boolean;
  massmailTestMsg: string;
  registrations: SPRegistration[];
  resetMassmailDraft: () => void;
  saveMassmailDraft: () => void;
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string; }>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  selectedEvent: DeloitteEvent;
  sendMassmailTestToOrganizers: () => Promise<void>;
  setComposerCrop: React.Dispatch<React.SetStateAction<"invite" | "massmail" | "qr">>;
  setEmailBody: React.Dispatch<React.SetStateAction<string>>;
  setEmailHeading: React.Dispatch<React.SetStateAction<string>>;
  setEmailSending: React.Dispatch<React.SetStateAction<boolean>>;
  setEmailSubject: React.Dispatch<React.SetStateAction<string>>;
  setMassmailCc: React.Dispatch<React.SetStateAction<string[]>>;
  setMassmailHeaderImage: React.Dispatch<React.SetStateAction<MailHeaderImage>>;
  /** v31.9.7: Kopfbild nur fuer diese Mail. */
  massmailCustomHeaderB64: string;
  setMassmailCustomHeaderB64: React.Dispatch<React.SetStateAction<string>>;
  setMassmailMode: React.Dispatch<React.SetStateAction<"closed" | "pick" | "paste" | "editor">>;
  setMassmailPasteRaw: React.Dispatch<React.SetStateAction<string>>;
  setMassmailSubheading: React.Dispatch<React.SetStateAction<string>>;
  setShowEmailModal: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  showEmailModal: boolean;
}

export const MassmailComposerModal: React.FC<MassmailComposerModalProps> = (p) => {
  const { applyMassmailHero, confirmDialog, emailBody, emailHeading, emailSending, emailSubject, eventServiceRef, isDe, massmailAudience, massmailCc, massmailCustomHeaderB64, setMassmailCustomHeaderB64, massmailDraftSaved, massmailEventPhotoB64, massmailHeaderImage, massmailHeaderOpts, massmailPasteRaw, massmailStatuses, massmailSubheading, massmailTesting, massmailTestMsg, registrations, resetMassmailDraft, saveMassmailDraft, searchUser, searchUsers, selectedEvent, sendMassmailTestToOrganizers, setComposerCrop, setEmailBody, setEmailHeading, setEmailSending, setEmailSubject, setMassmailCc, setMassmailHeaderImage, setMassmailMode, setMassmailPasteRaw, setMassmailSubheading, setShowEmailModal, showAlert, showEmailModal } = p;
        // v31.2: Das zusätzliche CC ist selten nötig und steht deshalb in
        // einem Aufklapper — offen nur, wenn schon jemand eingetragen ist,
        // damit ein gesetzter Verteiler nie unsichtbar mitfährt.
        const [ccOpen, setCcOpen] = React.useState<boolean>(massmailCc.length > 0);
        // v31.9.7: Eigenes Kopfbild. Dieselbe Kompressions-Leiter wie die
        // Inline-Bilder aus v31.7 — nur mit 600 px, weil der Mail-KOPF die
        // volle Tabellenbreite hat (die Inhaltszelle ist schmaler).
        const [headerBusy, setHeaderBusy] = React.useState(false);
        const [headerNote, setHeaderNote] = React.useState('');
        const pickCustomHeader = async (file: File): Promise<void> => {
          setHeaderBusy(true); setHeaderNote('');
          try {
            const out = await buildInlineImage(file, 600);
            if (!out.ok) {
              // Der Grund wird BENANNT — „hat nicht geklappt" laesst den
              // Organizer raten, ob es am Bild oder an der App lag.
              setHeaderNote(out.reason === 'too-big'
                ? (isDe ? `Auch verkleinert noch ${Math.round(charsToKb(out.chars))} KB — bitte ein einfacheres Bild nehmen (weniger Details, kein Screenshot).` : `Still ${Math.round(charsToKb(out.chars))} KB after compression — please use a simpler image.`)
                : (isDe ? 'Diese Datei konnte nicht gelesen werden.' : 'This file could not be read.'));
              return;
            }
            setMassmailCustomHeaderB64(out.dataUrl);
            // Eigenes Bild heisst volle Breite ohne Rand — der Orb-Deckel
            // von 180 px gilt nur fuer das Standard-Logo.
            setMassmailHeaderImage(prev => ({ ...prev, hero: 'custom', width: 600, paddingV: 0, paddingH: 0 }));
            setHeaderNote(isDe ? `Übernommen — ${out.width}×${out.height} px, ${Math.round(charsToKb(out.chars))} KB.` : `Applied — ${out.width}×${out.height} px, ${Math.round(charsToKb(out.chars))} KB.`);
          } finally {
            setHeaderBusy(false);
          }
        };
        // v17.10: Empfänger-Filter abhängig vom gewählten massmailAudience.
        const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];
        const recipients = (() => {
          if (massmailAudience === 'custom') {
            return registrations.filter(r => massmailStatuses.has(r.Status));
          }
          if (massmailAudience === 'waitOnly') {
            return registrations.filter(r => r.Status === 'Warteliste');
          }
          if (massmailAudience === 'activePlusWait') {
            return registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0 || r.Status === 'Warteliste');
          }
          if (massmailAudience === 'everyone') {
            // v31.9.6: Wirklich JEDE Zeile — keine Status-Aufzählung, sonst
            // fehlt jeder Status, den jemand später in SharePoint ergänzt.
            // Ohne diesen Zweig wäre die Auswahl still auf „nur aktive"
            // zurückgefallen (der Rückfall unten) und hätte damit etwas
            // anderes verschickt, als der Organizer angeklickt hat.
            return registrations;
          }
          if (massmailAudience === 'nachruecker') {
            // Aktive minus die in der eingefügten Liste enthaltenen E-Mails.
            const matches = (massmailPasteRaw || '').match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || [];
            const pastedSet = new Set(matches.map(m => m.toLowerCase()));
            return registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0 && !pastedSet.has((r.ParticipantEmail || '').toLowerCase()));
          }
          return registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0);
        })();
        // v30.67: Über `formatOrganizerList` wie alle anderen Mail-Stellen —
        // `join(', ')` auf den Roh-Werten („Nachname, Vorname") ergab
        // „Sathasivam, Philipp, Oesterle, Ines", Vor- und Nachnamen nicht
        // mehr auseinanderzuhalten.
        const orgNames = formatOrganizerList(selectedEvent.organizers || [], selectedEvent.emailLanguage || 'EN') || (selectedEvent.organizers || []).join(', ');
        const previewVars: Record<string, string> = {
          EventTitle: selectedEvent.title,
          Organizer: orgNames,
        };
        /**
         * v30.51.1: Das tatsächliche CC — EINMAL berechnet, für Anzeige UND
         * Versand.
         *
         * Dass die Liste hier oben steht und nicht erst im Versand, ist der
         * eigentliche Punkt: Vorher stand im Dialog eine ZUSAGE („Organizer
         * kommen automatisch auf CC"), während der Versand etwas anderes tat.
         * Jetzt zeigt der Dialog genau die Liste, die verschickt wird.
         *
         * v31.9.9: Der Organizer steht IMMER im CC — auch dann, wenn er als
         * Teilnehmer schon im An-Feld steht (Nutzer-Entscheidung 10.09.2026).
         * Bis v31.9.8 wurde er in genau diesem Fall herausgefiltert, mit der
         * Begründung „niemanden doppelt eintragen". Der Normalfall ist aber,
         * dass ein Organizer bei seinem eigenen Event angemeldet ist — das CC
         * war deshalb fast immer leer, und im Postfach steht der Organizer
         * anonym zwischen 150 Teilnehmern statt sichtbar als Absender-Seite.
         * Zugestellt wird dadurch nichts doppelt: Exchange liefert eine
         * Adresse einmal aus, auch wenn sie in To und Cc steht — dieselbe
         * Zusicherung, auf der schon das nicht gefilterte manuelle CC beruht.
         *
         * v31.10: Nur die Organizer, die dem Teilnehmer auch ANGEZEIGT werden
         * (Nutzer-Ansage 10.09.2026). Wer im Assistenten ausgeblendet ist,
         * gehört nicht ins CC einer Mail an ebendiese Teilnehmer — sonst
         * verrät der Mailkopf genau den Namen, den die Anmeldeseite
         * absichtlich verschweigt. Sind alle ausgeblendet, bleibt das
         * automatische CC leer; ein von Hand eingetragenes CC ist davon
         * unberührt, das ist eine bewusste Einzelentscheidung.
         *
         * `seen` bleibt: Eine Adresse, die zweimal unter den Organizern steht
         * oder zusätzlich von Hand ins CC getippt wurde, erscheint einmal.
         */
        const massmailCcPreview = ((): string[] => {
          const seen = new Set<string>();
          const out: string[] = [];
          for (const raw of visibleOrganizerEmails(selectedEvent)) {
            const e = (raw || '').trim();
            const lc = e.toLowerCase();
            if (!e || seen.has(lc)) continue;
            seen.add(lc);
            out.push(e);
          }
          for (const raw of massmailCc) {
            const e = (raw || '').trim();
            const lc = e.toLowerCase();
            if (!e || seen.has(lc)) continue;
            seen.add(lc);
            out.push(e);
          }
          return out;
        })();
        const customLogo = (() => {
          try {
            const o = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })();
        const sendAction = async (): Promise<void> => {
          if (!eventServiceRef || !selectedEvent) return;
          // v31.2: Meldungen zweisprachig; die Rückfrage nennt die Folge
          // (Leitfaden 2b), nicht nur die Zahl.
          if (recipients.length === 0) { showAlert(isDe ? 'In der gewählten Gruppe ist niemand — es gibt keine Empfänger.' : 'The selected group is empty — there are no recipients.'); return; }
          if (!(await confirmDialog(
            isDe
              ? `An ${recipients.length} Empfänger senden? Die Mail geht so raus, wie du sie jetzt in der Vorschau siehst.`
              : `Send to ${recipients.length} recipients? The email goes out exactly as you see it in the preview now.`,
            { confirmLabel: isDe ? 'Senden' : 'Send' },
          ))) return;
          setEmailSending(true);
          // Variablen einmalig auflösen (Massenmail geht an alle zusammen)
          const resolvedSubject = replacePlaceholders(emailSubject, previewVars);
          const resolvedHeading = replacePlaceholders(emailHeading, previewVars);
          const resolvedBody = replacePlaceholders(emailBody, previewVars);
          // v22.11: editierbare Unter-Überschrift (leer = "Event <Titel>").
          const resolvedSubheading = massmailSubheading.trim()
            ? replacePlaceholders(massmailSubheading, previewVars)
            : `Event ${selectedEvent.title}`;
          const fullBody = applyMassmailHero(wrapTemplate('#86bc25', resolvedHeading, resolvedSubheading, resolvedBody, undefined, massmailHeaderOpts));
          const allEmails = recipients.map(r => r.ParticipantEmail).join(';');
          const ccList = massmailCcPreview;
          const ccString = ccList.length > 0 ? ccList.join(';') : undefined;
          try {
            await eventServiceRef.queueEmail(
              resolvedSubject, allEmails, 'Alle Teilnehmer', fullBody,
              'Massenmail', selectedEvent.title, selectedEvent.id,
              ccString,
            );
            try { await eventServiceRef.logEventComm({ eventId: selectedEvent.id, eventTitle: selectedEvent.title, subject: resolvedSubject, bodyHtml: fullBody, emailType: 'Massenmail' }); } catch { /* */ }
            setEmailSending(false);
            // v30.51: Die Meldung nennt das ZUSÄTZLICHE CC getrennt — sonst
            // liest sich eine höhere Zahl so, als hätte das Event plötzlich
            // mehr Organizer.
            const ccInfo = ccString
              ? (isDe ? ` Auf CC: ${ccList.join(', ')}.` : ` On CC: ${ccList.join(', ')}.`)
              : (isDe ? ' Niemand auf CC — alle Organizer stehen schon im An-Feld.' : ' Nobody on CC — all organizers are already in the To field.');
            showAlert(isDe
              ? `Die Mail an ${recipients.length} Empfänger steht in der Warteschlange und geht in Kürze raus.${ccInfo}`
              : `The email to ${recipients.length} recipients is queued and goes out shortly.${ccInfo}`);
            setShowEmailModal(false);
            setMassmailMode('closed');
            setMassmailPasteRaw('');
          } catch {
            setEmailSending(false);
            showAlert(isDe ? 'Die Mail konnte nicht in die Warteschlange eingetragen werden. Bitte versuch es noch einmal.' : 'The email could not be queued. Please try again.');
          }
        };
        // v22.11: „Briefumschlag"-Kopf über der Vorschau — wie bei der
        // Einladungsmail (An: Empfängergruppe, Betreff: aufgelöster Subject).
        const audienceLabel = massmailAudience === 'custom'
          ? Array.from(massmailStatuses).join(', ')
          : massmailAudience === 'waitOnly' ? (isDe ? 'Nur Warteliste' : 'Waitlist only')
          : massmailAudience === 'activePlusWait' ? (isDe ? 'Teilnehmer + Warteliste' : 'Attendees + waitlist')
          : massmailAudience === 'everyone' ? (isDe ? 'Alle — auch Abgemeldete' : 'Everyone — incl. cancellations')
          : massmailAudience === 'nachruecker' ? (isDe ? 'Nachrücker (manueller Abgleich)' : 'Replacements (manual match)')
          : (isDe ? 'Alle aktiven Teilnehmer' : 'All active attendees');
        // v30.51.1: Die Vorschau nennt die WIRKLICHE CC-Zahl (s. massmailCcPreview).
        const ccCount = massmailCcPreview.length;
        const ccNobody = isDe ? 'niemand in CC' : 'nobody in CC';
        const previewToLine = `${recipients.length} ${isDe ? 'Empfänger' : 'recipients'} — ${audienceLabel} · ${ccCount > 0 ? `${ccCount} in CC` : ccNobody}`;
        const previewSubjectLine = replacePlaceholders(emailSubject, previewVars);
        // v31.2: Ob die Testmail-Rückmeldung ein Erfolg war — dieselbe Prüfung
        // wie bisher, nur einmal benannt statt zweimal im JSX gerechnet.
        const testOk = !!massmailTestMsg && (massmailTestMsg.indexOf('verschickt') >= 0 || massmailTestMsg.indexOf('sent') >= 0);
        return (
          <HtmlEditorModal
            open={showEmailModal}
            onClose={() => !emailSending && setShowEmailModal(false)}
            title={`${isDe ? `Massenmail an ${recipients.length} Empfänger` : `Mass email to ${recipients.length} recipients`} · ${selectedEvent.title}`}
            value={emailBody}
            onChange={setEmailBody}
            previewMode="email"
            emailSubject={emailSubject}
            onEmailSubjectChange={setEmailSubject}
            emailHeading={emailHeading}
            onEmailHeadingChange={setEmailHeading}
            emailSubheading={massmailSubheading}
            onEmailSubheadingChange={setMassmailSubheading}
            previewToLine={previewToLine}
            previewSubjectLine={previewSubjectLine}
            emailHeadingColor="#86bc25"
            previewVars={previewVars}
            insertableVars={[
              { key: '{{EventTitle}}', label: 'Event' },
              { key: '{{Organizer}}', label: 'Organizer' },
            ]}
            imageBase64={(massmailHeaderImage.hero === 'event' && massmailEventPhotoB64) ? massmailEventPhotoB64 : customLogo}
            imageWidth={massmailHeaderImage.width}
            imagePaddingV={massmailHeaderImage.paddingV}
            imagePaddingH={massmailHeaderImage.paddingH}
            onImageWidthChange={(w) => setMassmailHeaderImage(p => ({ ...p, width: w }))}
            onImagePaddingVChange={(v) => setMassmailHeaderImage(p => ({ ...p, paddingV: v }))}
            onImagePaddingHChange={(h) => setMassmailHeaderImage(p => ({ ...p, paddingH: h }))}
            headerExtra={(
              // v31.2: Drei Fragen in der Reihenfolge, in der ein Organizer sie
              // beantwortet: An wen? → Wie sieht der Kopf aus? → Was prüfe ich vor
              // dem Senden? Betreff und Text folgen darunter im Editor, der
              // Sende-Knopf sitzt im Fuß. Klassen statt Inline-Kästen (Hover kommt mit).
              <div className="dex-ui-stack" style={{ gap: 20 }}>
                <div>
                  <div className="dex-ui-section-title">{isDe ? 'An wen geht die Mail?' : 'Who receives the email?'}</div>
                  <div className="dex-ui-inline">
                    <span className={cx('dex-ui-pill', recipients.length > 0 ? 'dex-ui-pill--green' : 'dex-ui-pill--red')}><Users size={13} /> {recipients.length} {isDe ? 'Empfänger' : 'recipients'}</span>
                    {/* v31.10: `--wrap`, weil `audienceLabel` bei „Eigene Auswahl"
                        die angehakten Status aneinanderreiht („Angemeldet, QR
                        versendet, Eingecheckt, …"). Eine Pille bricht sonst nie um
                        und schob den Dialog auf dem Handy 90 px nach rechts aus
                        dem Bild — im Harness bei 390 px gemessen. */}
                    <span className="dex-ui-pill dex-ui-pill--gray dex-ui-pill--wrap">{audienceLabel}</span>
                    <span className={cx('dex-ui-pill', ccCount > 0 ? 'dex-ui-pill--blue' : 'dex-ui-pill--gray')}>{ccCount > 0 ? `${ccCount} CC` : ccNobody}</span>
                  </div>
                  {/* v30.51.1: Was WIRKLICH ins CC geht, statt einer Zusage.
                      v31.9.9: Seit die Organizer immer auf CC stehen, ist die Liste
                      genau die Zusage — leer ist sie nur noch, wenn am Event keine
                      Organizer-Adresse hinterlegt ist. Genau das sagt der Text dann
                      auch, statt einen Grund zu nennen, der nicht mehr gilt. */}
                  <div className="dex-ui-help" style={{ wordBreak: 'break-word' }}>
                    {recipients.length === 0
                      ? (isDe ? 'In dieser Gruppe ist niemand. Schließe den Dialog und wähle im Schritt davor eine andere Gruppe.' : 'This group is empty. Close the dialog and pick another group in the step before.')
                      : (isDe
                        ? <>Die Gruppe hast du im Schritt davor gewählt. <strong>CC:</strong> {ccCount === 0 ? 'niemand — an diesem Event ist keine Organizer-Adresse hinterlegt.' : massmailCcPreview.join(', ')}</>
                        : <>You picked the group in the step before. <strong>CC:</strong> {ccCount === 0 ? 'nobody — this event has no organizer address on file.' : massmailCcPreview.join(', ')}</>)}
                  </div>
                  {/* v30.51: Zusätzliches CC direkt unter der Empfänger-Zeile — CC ist
                      eine Aussage über den Verteiler, nicht über die Gestaltung. */}
                  <button type="button" className={cx('dex-ui-disclosure', ccOpen && 'is-open')} onClick={() => setCcOpen(o => !o)} aria-expanded={ccOpen} style={{ marginTop: 6 }}>
                    <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                    {isDe ? 'Zusätzlich jemanden auf CC setzen' : 'Put someone else on CC'}
                    {massmailCc.length > 0 && <span className="dex-ui-disclosure-count">{massmailCc.length}</span>}
                  </button>
                  {ccOpen && (
                    <div className="dex-ui-disclosure-body">
                      <RecipientPicker
                        label={isDe ? 'Zusätzlich auf CC' : 'Additional CC'}
                        hint={isDe
                          ? 'Personen über die Suche, Funktionspostfächer im Feld darunter. Die Organizer des Events sind ohnehin auf CC und müssen hier nicht eingetragen werden.'
                          : 'People via search, shared mailboxes in the field below. The event organizers are on CC anyway.'}
                        emptyText={isDe ? 'Kein zusätzliches CC — es gehen nur die Organizer mit.' : 'No additional CC — only the organizers.'}
                        value={massmailCc} onChange={setMassmailCc}
                        searchUsers={searchUsers} searchUserByEmail={searchUser}
                        disabled={emailSending}
                      />
                    </div>
                  )}
                </div>

                <div>
                  <div className="dex-ui-section-title">{isDe ? 'Wie sieht der Kopf der Mail aus?' : 'What does the email header look like?'}</div>
                  {/* v30.52: gemeinsame Auswahl (s. admin/MailHeaderImageChooser) — vorher
                      stand dieselbe Reiter-Reihe hier und in der Einladungsmail zweimal. */}
                  <MailHeaderImageChooser
                    value={massmailHeaderImage} onChange={setMassmailHeaderImage}
                    eventPhotoB64={massmailEventPhotoB64} disabled={emailSending}
                    onCrop={() => setComposerCrop('massmail')} isDe={isDe}
                    customB64={massmailCustomHeaderB64}
                    onPickCustom={(f) => { pickCustomHeader(f).catch(() => setHeaderBusy(false)); }}
                    onRemoveCustom={() => {
                      setMassmailCustomHeaderB64('');
                      setHeaderNote('');
                      // Zurueck auf das Standard-Logo — „custom" ohne Bild
                      // waere eine Auswahl, die still den Platzhalter zeigt.
                      setMassmailHeaderImage(prev => ({ ...prev, hero: 'logo' }));
                    }}
                    customBusy={headerBusy}
                    customNote={headerNote}
                  />
                  <div className="dex-ui-help">
                    {/* v31.10: ohne „neben der Vorschau" — auf dem Handy stehen die
                        Felder weiter unten in derselben Spalte, nicht daneben. */}
                    {isDe ? 'Breite und Abstand des Bildes stellst du weiter unten ein.' : 'Width and spacing of the image are set further down.'}
                  </div>
                </div>

                {/* v31.12: Die Umfrage sitzt hier — nach dem Kopf, vor der
                    Absende-Prüfung. Nutzer-Ansage 11.09.2026: erreichbar über
                    „E-Mail versenden", nicht als eigener Weg daneben. */}
                <PollComposerSection
                  selectedEvent={selectedEvent}
                  eventServiceRef={eventServiceRef}
                  isDe={isDe}
                  disabled={emailSending}
                  emailBody={emailBody}
                  setEmailBody={setEmailBody}
                />

                <div>
                  <div className="dex-ui-section-title">{isDe ? 'Bevor du sendest' : 'Before you send'}</div>
                  {/* v31.10: `flexWrap` plus eine Mindestbreite für den Textblock —
                      `dex-ui-step` ist eine Zeile ohne Umbruch, und `dex-ui-step-body`
                      trägt `min-width: 0`. Auf dem Handy schrumpfte der Text deshalb
                      neben dem Knopf auf ein Wort je Zeile („Dein / Text / wird /
                      ohnehin …"), und die Pille „Gespeichert" lag darüber. Mit der
                      Mindestbreite rutscht der KNOPF in die zweite Zeile, wo er als
                      Ganzes hingehört; auf dem Desktop ändert sich nichts. */}
                  <div className="dex-ui-stack">
                    <div className={cx('dex-ui-step', massmailDraftSaved && 'is-done')} style={{ flexWrap: 'wrap' }}>
                      <span className="dex-ui-step-num">{massmailDraftSaved ? <Check size={14} /> : 1}</span>
                      <div className="dex-ui-step-body" style={{ minWidth: 170 }}>
                        <div className="dex-ui-step-title">{isDe ? 'Zwischenstand sichern' : 'Save your progress'}</div>
                        <div className="dex-ui-step-hint">{isDe ? 'Dein Text wird ohnehin automatisch gespeichert und beim nächsten Öffnen wiederhergestellt.' : 'Your text is saved automatically anyway and restored next time you open it.'}</div>
                      </div>
                      <div className="dex-ui-step-action dex-ui-inline" style={{ justifyContent: 'flex-end' }}>
                        {massmailDraftSaved && <span className="dex-ui-pill dex-ui-pill--green"><Check size={12} /> {isDe ? 'Gespeichert' : 'Saved'}</span>}
                        <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={saveMassmailDraft} disabled={emailSending}>
                          <Check size={14} /> {isDe ? 'Entwurf speichern' : 'Save draft'}
                        </button>
                      </div>
                    </div>
                    <div className={cx('dex-ui-step', massmailTesting && 'is-pending')} style={{ flexWrap: 'wrap' }}>
                      <span className="dex-ui-step-num">2</span>
                      <div className="dex-ui-step-body" style={{ minWidth: 170 }}>
                        <div className="dex-ui-step-title">{isDe ? 'Erst an die Organizer testen' : 'Test with the organizers first'}</div>
                        <div className="dex-ui-step-hint">{isDe ? 'Schickt die Mail so, wie sie jetzt ist, mit [TEST] im Betreff — nur an die Organizer des Events.' : 'Sends the email as it is now, with [TEST] in the subject — to the event organizers only.'}</div>
                        {massmailTestMsg && (
                          <div className={cx('dex-ui-callout', testOk ? 'dex-ui-callout--success' : 'dex-ui-callout--warn')} style={{ marginTop: 8, padding: '6px 10px' }} role="status">{massmailTestMsg}</div>
                        )}
                      </div>
                      <div className="dex-ui-step-action">
                        <button type="button" className="btn btn-outline dex-ui-btn-sm" onClick={() => { sendMassmailTestToOrganizers().catch(() => { /* */ }); }} disabled={emailSending || massmailTesting}>
                          <Send size={14} /> {massmailTesting ? (isDe ? 'Sendet…' : 'Sending…') : (isDe ? 'Testmail senden' : 'Send test email')}
                        </button>
                      </div>
                    </div>
                  </div>
                  {/* v31.10: Was „Auf Vorlage zurücksetzen" wegwirft, stand nur im
                      `title` — auf dem Handy gibt es kein Überfahren, der Hinweis war
                      dort also unerreichbar (Leitfaden 6b). Er steht jetzt als Text
                      unter dem Knopf; der `title` bleibt als Zugabe. Und der Verweis
                      auf den Sende-Knopf nennt die Fußzeile statt „unten rechts" —
                      auf dem Handy sitzt der Knopf unten über die volle Breite. */}
                  <div style={{ marginTop: 8 }}>
                    <div className="dex-ui-muted" style={{ fontSize: '0.76rem' }}>{isDe ? 'Der Versand selbst ist der grüne Knopf in der Fußzeile dieses Dialogs.' : 'Sending itself is the green button in the footer of this dialog.'}</div>
                    <div style={{ marginTop: 4 }}>
                      <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" style={{ marginLeft: -8 }} onClick={resetMassmailDraft} disabled={emailSending}
                        title={isDe ? 'Setzt Betreff, Überschrift, Text und zusätzliches CC auf die Vorlage zurück.' : 'Resets subject, heading, text and additional CC to the template.'}>
                        {isDe ? 'Auf Vorlage zurücksetzen' : 'Reset to template'}
                      </button>
                      {/* Die Folge steht UNTER dem Knopf, nicht daneben: nebeneinander
                          bricht sie auf dem Handy ohnehin um, und dann steht der
                          Knopf als erste Zeile über einem Fließtext und liest sich
                          als Überschrift. */}
                      <div className="dex-ui-help" style={{ marginTop: 2 }}>
                        {isDe ? 'Setzt Betreff, Überschrift, Text und zusätzliches CC auf die Vorlage zurück.' : 'Resets subject, heading, text and additional CC to the template.'}
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}
            extraAction={{
              label: emailSending ? (isDe ? 'Wird eingetragen…' : 'Queuing…') : (isDe ? `An ${recipients.length} Empfänger senden` : `Send to ${recipients.length} recipients`),
              onClick: sendAction,
              disabled: emailSending || !emailSubject.trim() || !emailBody.trim() || recipients.length === 0,
              icon: <Send size={16} />,
            }}
          />
        );
};

