/* MassmailComposerModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 15244-15482 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { replacePlaceholders, wrapTemplate } from '../../../services/EmailTemplates';
import { HtmlEditorModal } from '../../HtmlEditorModal';
import RecipientPicker from '../../admin/RecipientPicker';
import MailHeaderImageChooser from '../../admin/MailHeaderImageChooser';
import { Check, Send } from '../../Icons';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';
import { MailHeaderImage } from '../../../utils/mailHeaderImage';

export interface MassmailComposerModalProps {
  applyMassmailHero: (wrappedHtml: string) => string;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  emailBody: string;
  emailHeading: string;
  emailSending: boolean;
  emailSubject: string;
  eventServiceRef: EventService;
  isDe: boolean;
  massmailAudience: "active" | "activePlusWait" | "waitOnly" | "nachruecker" | "custom";
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
  setMassmailMode: React.Dispatch<React.SetStateAction<"closed" | "pick" | "paste" | "editor">>;
  setMassmailPasteRaw: React.Dispatch<React.SetStateAction<string>>;
  setMassmailSubheading: React.Dispatch<React.SetStateAction<string>>;
  setShowEmailModal: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  showEmailModal: boolean;
}

export const MassmailComposerModal: React.FC<MassmailComposerModalProps> = (p) => {
  const { applyMassmailHero, confirmDialog, emailBody, emailHeading, emailSending, emailSubject, eventServiceRef, isDe, massmailAudience, massmailCc, massmailDraftSaved, massmailEventPhotoB64, massmailHeaderImage, massmailHeaderOpts, massmailPasteRaw, massmailStatuses, massmailSubheading, massmailTesting, massmailTestMsg, registrations, resetMassmailDraft, saveMassmailDraft, searchUser, searchUsers, selectedEvent, sendMassmailTestToOrganizers, setComposerCrop, setEmailBody, setEmailHeading, setEmailSending, setEmailSubject, setMassmailCc, setMassmailHeaderImage, setMassmailMode, setMassmailPasteRaw, setMassmailSubheading, setShowEmailModal, showAlert, showEmailModal } = p;
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
          if (massmailAudience === 'nachruecker') {
            // Aktive minus die in der eingefügten Liste enthaltenen E-Mails.
            const matches = (massmailPasteRaw || '').match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || [];
            const pastedSet = new Set(matches.map(m => m.toLowerCase()));
            return registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0 && !pastedSet.has((r.ParticipantEmail || '').toLowerCase()));
          }
          return registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0);
        })();
        const orgNames = (selectedEvent.organizers || []).join(', ');
        const previewVars: Record<string, string> = {
          EventTitle: selectedEvent.title,
          Organizer: orgNames,
        };
        /**
         * v30.51.1: Das tatsächliche CC — EINMAL berechnet, für Anzeige UND
         * Versand.
         *
         * Zwei Regeln, die nicht dieselbe sind:
         *  - Das AUTOMATISCHE Organizer-CC (v17.10) wird gegen die Empfänger
         *    entdoppelt. Es soll niemanden doppelt eintragen, den ohnehin
         *    jemand anschreibt.
         *  - Ein SELBST eingetragenes CC wird nicht gefiltert. Der gemeldete
         *    Fall: zwei Personen eingetragen, angekommen ist eine — die andere
         *    war selbst Teilnehmer und stand damit schon im An-Feld, also warf
         *    der Filter sie still hinaus. Wer jemanden ausdrücklich auf CC
         *    setzt, hat sich dabei etwas gedacht. Doppelt zugestellt wird
         *    nichts: Exchange liefert eine Adresse einmal aus, auch wenn sie
         *    in To und Cc steht.
         *
         * Dass die Liste hier oben steht und nicht erst im Versand, ist der
         * eigentliche Punkt: Vorher stand im Dialog eine ZUSAGE („Organizer
         * kommen automatisch auf CC"), während der Versand etwas anderes tat.
         * Jetzt zeigt der Dialog genau die Liste, die verschickt wird.
         */
        const massmailCcPreview = ((): string[] => {
          const recipientSet = new Set(recipients.map(r => (r.ParticipantEmail || '').toLowerCase()));
          const seen = new Set<string>();
          const out: string[] = [];
          for (const raw of (selectedEvent.organizerEmails || [])) {
            const e = (raw || '').trim();
            const lc = e.toLowerCase();
            if (!e || recipientSet.has(lc) || seen.has(lc)) continue;
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
          if (recipients.length === 0) { showAlert('Keine Empfänger in der gewählten Auswahl.'); return; }
          if (!(await confirmDialog(`E-Mail an ${recipients.length} Teilnehmer senden?`, { confirmLabel: isDe ? 'Senden' : 'Send' }))) return;
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
              ? ` (auf CC: ${ccList.join(', ')})`
              : ' (niemand auf CC — alle Organizer stehen schon im An-Feld)';
            showAlert(`E-Mail an ${recipients.length} Empfänger in die Warteschlange eingetragen.${ccInfo}`);
            setShowEmailModal(false);
            setMassmailMode('closed');
            setMassmailPasteRaw('');
          } catch {
            setEmailSending(false);
            showAlert('Fehler beim Eintragen der E-Mail.');
          }
        };
        // v22.11: „Briefumschlag"-Kopf über der Vorschau — wie bei der
        // Einladungsmail (An: Empfängergruppe, Betreff: aufgelöster Subject).
        const audienceLabel = massmailAudience === 'custom'
          ? Array.from(massmailStatuses).join(', ')
          : massmailAudience === 'waitOnly' ? 'Nur Warteliste'
          : massmailAudience === 'activePlusWait' ? 'Teilnehmer + Warteliste'
          : massmailAudience === 'nachruecker' ? 'Nachrücker (manueller Abgleich)'
          : 'Alle aktiven Teilnehmer';
        // v30.51.1: Die Vorschau nennt die WIRKLICHE CC-Zahl (s. massmailCcPreview).
        const previewToLine = `${recipients.length} Empfänger — ${audienceLabel}${massmailCcPreview.length > 0 ? ` · ${massmailCcPreview.length} in CC` : ' · niemand in CC'}`;
        const previewSubjectLine = replacePlaceholders(emailSubject, previewVars);
        return (
          <HtmlEditorModal
            open={showEmailModal}
            onClose={() => !emailSending && setShowEmailModal(false)}
            title={`Massenmail an ${recipients.length} Teilnehmer: ${selectedEvent.title}`}
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
              <div style={{ padding: 12, background: 'var(--dex-gray-50, #fafafa)', border: '1px solid var(--dex-gray-200)', borderRadius: 'var(--dex-radius)', marginBottom: 4 }}>
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginBottom: 8 }}>
                  {isDe
                    ? <>Geht an <strong>{recipients.length}</strong> Empfänger (die oben gewählte Gruppe).</>
                    : <>Goes to <strong>{recipients.length}</strong> recipients (the group selected above).</>}
                </div>
                {/* v30.51.1: Was WIRKLICH ins CC geht, statt einer Zusage.
                    Vorher stand hier „Organizer kommen automatisch auf CC" —
                    das stimmt aber genau dann nicht, wenn die Organizer selbst
                    am Event teilnehmen (der Normalfall beim eigenen Event):
                    Dann stehen sie schon im An-Feld und werden nicht noch
                    einmal ins CC gesetzt. Wer das nicht weiß, sucht den Fehler
                    in der App. */}
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginBottom: 8 }}>
                  <strong style={{ color: 'var(--dex-gray-700)' }}>CC: </strong>
                  {massmailCcPreview.length === 0
                    ? <span style={{ color: 'var(--dex-gray-500)' }}>niemand — alle Organizer stehen bereits im An-Feld.</span>
                    : <span style={{ wordBreak: 'break-word' }}>{massmailCcPreview.join(', ')}</span>}
                </div>
                {/* v30.51: Zusätzliches CC. Bewusst hier oben, direkt unter der
                    Empfänger-Zeile — CC ist eine Aussage über den Verteiler,
                    nicht über die Gestaltung. */}
                <div style={{ marginBottom: 10 }}>
                  <RecipientPicker
                    label={isDe ? 'Zusätzlich auf CC' : 'Additional CC'}
                    hint={isDe
                      ? 'Personen über die Suche, Funktionspostfächer im Feld darunter. Die Organizer des Events sind ohnehin auf CC und müssen hier nicht eingetragen werden.'
                      : 'People via search, shared mailboxes in the field below. The event organizers are on CC anyway.'}
                    emptyText={isDe ? 'Kein zusätzliches CC — es gehen nur die Organizer mit.' : 'No additional CC — only the organizers.'}
                    value={massmailCc}
                    onChange={setMassmailCc}
                    searchUsers={searchUsers}
                    searchUserByEmail={searchUser}
                    disabled={emailSending}
                  />
                </div>
                {/* v30.52: gemeinsame Auswahl (s. admin/MailHeaderImageChooser) —
                    vorher stand dieselbe Reiter-Reihe hier und in der
                    Einladungsmail wortgleich ein zweites Mal. */}
                <MailHeaderImageChooser
                  value={massmailHeaderImage}
                  onChange={setMassmailHeaderImage}
                  eventPhotoB64={massmailEventPhotoB64}
                  disabled={emailSending}
                  onCrop={() => setComposerCrop('massmail')}
                  isDe={isDe}
                />
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                  <button type="button" className="btn btn-secondary" onClick={saveMassmailDraft} disabled={emailSending} style={{ fontSize: '0.78rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                    <Check size={14} /> {isDe ? 'Entwurf speichern' : 'Save draft'}
                  </button>
                  {massmailDraftSaved && (
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, fontSize: '0.78rem' }}>
                      <Check size={14} /> {isDe ? 'Gespeichert' : 'Saved'}
                    </span>
                  )}
                  <button type="button" className="btn btn-outline" onClick={() => { sendMassmailTestToOrganizers().catch(() => { /* */ }); }} disabled={emailSending || massmailTesting} style={{ fontSize: '0.78rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                    <Send size={14} /> {massmailTesting ? (isDe ? 'Sendet…' : 'Sending…') : (isDe ? 'Testmail an Organizer' : 'Test email to organizers')}
                  </button>
                  <span style={{ flex: 1 }} />
                  <button type="button" onClick={resetMassmailDraft} disabled={emailSending} style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600, fontSize: '0.74rem', textDecoration: 'underline' }}>
                    {isDe ? 'Zurücksetzen' : 'Reset'}
                  </button>
                </div>
                {massmailTestMsg && (
                  <div style={{ marginTop: 8, fontSize: '0.78rem', color: (massmailTestMsg.indexOf('verschickt') >= 0 || massmailTestMsg.indexOf('sent') >= 0) ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-orange-dark, #b35a00)' }}>
                    {massmailTestMsg}
                  </div>
                )}
                <div style={{ marginTop: 8, fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                  {isDe
                    ? 'Dein Text wird zusätzlich automatisch gespeichert und beim nächsten Öffnen wiederhergestellt.'
                    : 'Your text is also saved automatically and restored next time you open it.'}
                </div>
              </div>
            )}
            extraAction={{
              label: emailSending ? 'Wird eingetragen…' : `An ${recipients.length} Teilnehmer senden`,
              onClick: sendAction,
              disabled: emailSending || !emailSubject.trim() || !emailBody.trim() || recipients.length === 0,
              icon: <Send size={16} />,
            }}
          />
        );
};

