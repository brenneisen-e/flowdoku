/* InviteComposerModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 15485-16084 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { getBlockedInviteRecipients } from '../../../utils/inviteGuards';
import { replacePlaceholders, wrapTemplate } from '../../../services/EmailTemplates';
import { formatOrganizerList } from '../../../context/eventTextHelpers';
import RecipientPicker from '../../admin/RecipientPicker';
import MailHeaderImageChooser from '../../admin/MailHeaderImageChooser';
import { Check, Send } from '../../Icons';
import { HtmlEditorModal } from '../../HtmlEditorModal';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';
import { MailHeaderImage } from '../../../utils/mailHeaderImage';

export interface InviteComposerModalProps {
  applyInviteHero: (wrappedHtml: string) => string;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  currentUser: import("../../../types/index").User;
  eventServiceRef: EventService;
  getGroupMembers: (groupEmail: string) => Promise<{ groupName: string; members: { email: string; displayName: string; firstName?: string; lastName?: string; jobTitle?: string; location?: string; }[]; }>;
  inviteAddInput: string;
  inviteAudienceOpen: boolean;
  inviteBody: string;
  inviteCc: string[];
  inviteCustomEmails: string[];
  invitedLc: Set<string>;
  inviteDraftSaved: boolean;
  inviteEventPhotoB64: string;
  inviteHeaderImage: MailHeaderImage;
  inviteHeaderOpts: { imageWidth: number; imagePaddingV: number; imagePaddingH: number; };
  inviteHeading: string;
  inviteSending: boolean;
  inviteSubheading: string;
  inviteSubject: string;
  inviteTarget: "organizer" | "audience" | "pending" | "uninvited";
  isDe: boolean;
  refreshEvents: () => Promise<void>;
  registrations: SPRegistration[];
  resetInviteDraft: () => void;
  saveInviteDraft: () => void;
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string; }>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  selectedEvent: DeloitteEvent;
  setComposerCrop: React.Dispatch<React.SetStateAction<"invite" | "massmail" | "qr">>;
  setInviteAddInput: React.Dispatch<React.SetStateAction<string>>;
  setInviteAudienceOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setInviteBody: React.Dispatch<React.SetStateAction<string>>;
  setInviteCc: React.Dispatch<React.SetStateAction<string[]>>;
  setInviteCustomEmails: React.Dispatch<React.SetStateAction<string[]>>;
  setInviteHeaderImage: React.Dispatch<React.SetStateAction<MailHeaderImage>>;
  setInviteHeading: React.Dispatch<React.SetStateAction<string>>;
  setInviteSending: React.Dispatch<React.SetStateAction<boolean>>;
  setInviteSubheading: React.Dispatch<React.SetStateAction<string>>;
  setInviteSubject: React.Dispatch<React.SetStateAction<string>>;
  setInviteTarget: React.Dispatch<React.SetStateAction<"organizer" | "audience" | "pending" | "uninvited">>;
  setShowInviteModal: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  showInviteModal: boolean;
  siteUrl: string;
  updateEvent: (eventId: string, updates: Record<string, unknown>, opts?: { skipReload?: boolean; }) => Promise<boolean>;
}

export const InviteComposerModal: React.FC<InviteComposerModalProps> = (p) => {
  const { applyInviteHero, confirmDialog, currentUser, eventServiceRef, getGroupMembers, inviteAddInput, inviteAudienceOpen, inviteBody, inviteCc, inviteCustomEmails, invitedLc, inviteDraftSaved, inviteEventPhotoB64, inviteHeaderImage, inviteHeaderOpts, inviteHeading, inviteSending, inviteSubheading, inviteSubject, inviteTarget, isDe, refreshEvents, registrations, resetInviteDraft, saveInviteDraft, searchUser, searchUsers, selectedEvent, setComposerCrop, setInviteAddInput, setInviteAudienceOpen, setInviteBody, setInviteCc, setInviteCustomEmails, setInviteHeaderImage, setInviteHeading, setInviteSending, setInviteSubheading, setInviteSubject, setInviteTarget, setShowInviteModal, showAlert, showInviteModal, siteUrl, updateEvent } = p;
        const audienceEmails = (selectedEvent.audienceFilter || [])
          .map(s => (s || '').trim())
          .filter(Boolean);
        const myEmail = currentUser.email || '';
        const myDisplayName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || myEmail;
        // v28.37: „Nur an noch nicht Angemeldete" — der Verteiler abzueglich
        // aller, die im Event schon eine Zeile haben (angemeldet, Warteliste,
        // eingecheckt ODER abgemeldet). Genau der Nachfass-Fall: erinnern, ohne
        // die anzuschreiben, die sich längst entschieden haben. Verteiler-
        // Adressen ohne '@' (Gruppen) bleiben drin — sie werden beim Senden
        // ohnehin in Mitglieder aufgeloest, und wer darin schon angemeldet ist,
        // laesst sich vorher nicht herausrechnen.
        const decidedLc = new Set(
          registrations
            .map(r => (r.ParticipantEmail || '').trim().toLowerCase())
            .filter(Boolean)
        );
        const pendingEmails = audienceEmails.filter(e => {
          const lc = e.toLowerCase();
          if (lc.indexOf('@') < 0) return true; // Verteiler/Gruppe → mitnehmen
          return !decidedLc.has(lc);
        });
        const alreadyDecidedCount = audienceEmails.length - pendingEmails.length;
        // v28.37: Zweiter Abgleich — gegen die bereits verschickten
        // Einladungsmails. Trifft den Fall „Verteiler wurde nachträglich
        // erweitert, jetzt nur die Neuen anschreiben".
        const uninvitedEmails = audienceEmails.filter(e => {
          const lc = e.toLowerCase();
          if (lc.indexOf('@') < 0) return true; // Gruppe → nicht aufloesbar, mitnehmen
          return !(invitedLc && invitedLc.has(lc));
        });
        const alreadyInvitedCount = audienceEmails.length - uninvitedEmails.length;
        const modeEmails = inviteTarget === 'pending'
          ? pendingEmails
          : (inviteTarget === 'uninvited' ? uninvitedEmails : audienceEmails);
        // Handanpassung sticht die Radio-Auswahl.
        const effectiveEmails = inviteCustomEmails || modeEmails;
        const targetEmails = inviteTarget === 'organizer'
          ? [myEmail].filter(Boolean)
          : effectiveEmails;
        /** Echter Massenversand (Verteiler ODER Nachfass) — nur „An mich" nicht. */
        const isBroadcast = inviteTarget !== 'organizer';
        // v11.43: Organizer-Mails als CC mitschicken — damit alle Organizer
        // sehen, dass die Einladung raus ist und ggf. auf Rückfragen
        // antworten können. Duplikate gegenüber TO werden rausgefiltert
        // (z.B. wenn der Sender selbst Organizer ist und 'An mich' wählt).
        const toLcSet = new Set(targetEmails.map(e => (e || '').toLowerCase()));
        // v30.51.1: Das AUTOMATISCHE Organizer-CC wird gegen die Empfänger
        // entdoppelt, ein selbst eingetragenes CC NICHT — sonst verschwindet
        // eine ausdrücklich gewählte Person still, nur weil sie ohnehin
        // Empfänger ist (s. Massenmail).
        const ccEmails = ((): string[] => {
          const seen = new Set<string>();
          const out: string[] = [];
          for (const raw of (selectedEvent.organizerEmails || [])) {
            const e = (raw || '').trim();
            const lc = e.toLowerCase();
            if (!e || toLcSet.has(lc) || seen.has(lc)) continue;
            seen.add(lc);
            out.push(e);
          }
          for (const raw of inviteCc) {
            const e = (raw || '').trim();
            const lc = e.toLowerCase();
            if (!e || seen.has(lc)) continue;
            seen.add(lc);
            out.push(e);
          }
          return out;
        })();
        // v11.41: Blocked-Check für den aktuell gewählten Empfänger-Modus.
        // 'organizer'-Modus blockt eigentlich nie — die eigene Mail ist immer
        // eine Person, kein Verteiler — aber wir laufen das defensiv mit.
        const blockedInTargets = getBlockedInviteRecipients(targetEmails);
        const blockedInAudience = getBlockedInviteRecipients(audienceEmails);
        // v28.38 BUG-FIX: Auch die Beschriftung des Senden-Knopfs muss die
        // WIRKLICH adressierte Liste nennen — sie zeigte bisher stur die
        // Verteilergröße, selbst nachdem Adressen entfernt wurden.
        const nRecipients = targetEmails.length;
        const recipientLabel = inviteTarget === 'organizer'
          ? (isDe ? `An mich (${myEmail})` : `To me (${myEmail})`)
          : (isDe
            ? `${inviteCustomEmails
              ? 'An angepasste Auswahl'
              : inviteTarget === 'uninvited' ? 'An noch nicht Eingeladene'
              : inviteTarget === 'pending' ? 'An noch nicht Angemeldete'
              : 'An Mailverteiler'} (${nRecipients === 0 ? 'leer' : nRecipients + ' Empfänger'})`
            : `${inviteCustomEmails
              ? 'To adjusted selection'
              : inviteTarget === 'uninvited' ? 'To not-yet-invited'
              : inviteTarget === 'pending' ? 'To not-yet-registered'
              : 'To mail distribution'} (${nRecipients === 0 ? 'empty' : nRecipients + ' recipients'})`);
        // v30.67: s. MassmailComposerModal — „Nachname, Vorname"-Mus vermeiden.
        const orgNames = formatOrganizerList(selectedEvent.organizers || [], selectedEvent.emailLanguage || 'EN') || (selectedEvent.organizers || []).join(', ');
        const appUrl = `${siteUrl}/SitePages/DEX.aspx?env=WebView`;
        const previewVars: Record<string, string> = {
          EventTitle: selectedEvent.title,
          Organizer: orgNames,
          Link: appUrl,
        };
        const customLogo = (() => {
          try {
            const o = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })();
        const sendAction = async (): Promise<void> => {
          if (!eventServiceRef || !selectedEvent) return;
          if (targetEmails.length === 0) {
            showAlert(isDe
              ? (isBroadcast
                ? 'Es ist kein Mailverteiler auf dem Event hinterlegt. Bitte zuerst in Schritt 3 (Sichtbarkeit) Empfänger ergänzen.'
                : 'Keine eigene E-Mail-Adresse verfügbar.')
              : (isBroadcast
                ? 'No mail distribution list configured on the event. Please add recipients in step 3 (Visibility) first.'
                : 'No own email address available.'));
            return;
          }
          // v11.41: Hart blocken — Einladungsmail darf nie an pauschale
          // Standort-/All-Verteiler ('deall', 'all', 'de.<stadt>') gehen.
          if (blockedInTargets.length > 0) {
            const lines = blockedInTargets.map(b => `• ${b.email}  (${b.reason})`).join('\n');
            showAlert(isDe
              ? `Die Einladungs-Mail darf NICHT an pauschale Standort- oder All-Verteiler verschickt werden.\n\nFolgende Empfänger sind blockiert:\n\n${lines}\n\nBitte entferne diese Adressen aus dem Mailverteiler in Schritt 3 des Event-Edits oder nutze die Option „An mich (zum Weiterleiten)".`
              : `The invitation email must NOT be sent to entire location or all-distribution lists.\n\nThe following recipients are blocked:\n\n${lines}\n\nPlease remove these addresses from the mail distribution in step 3 of event edit, or use the option "To me (for forwarding)".`);
            return;
          }
          // v27.11 (Bug-Report): Verteiler VOR dem Versand in einzelne
          // Mitglieder-Adressen auflösen. Vorher ging die Verteiler-Adresse
          // roh ins To-Feld — Exchange lehnt das ab, sobald der Verteiler nur
          // autorisierte Absender zulässt (die Shared Mailbox
          // no_reply.events@deloitte.de ist das i.d.R. nicht); der NDR landete
          // unsichtbar in der Shared Mailbox und für den Organizer sah alles
          // erfolgreich aus. Auflösung via Graph (transitive Mitglieder,
          // gleicher Resolver wie die Sichtbarkeits-Auflösung v16.4);
          // nicht auflösbare Einträge bleiben als Direktadresse erhalten.
          let resolvedRecipients: string[] = targetEmails;
          if (isBroadcast) {
            setInviteSending(true);
            const out: string[] = [];
            const seen = new Set<string>();
            const push = (e: string): void => {
              const lc = (e || '').trim().toLowerCase();
              if (lc && lc.indexOf('@') > 0 && !seen.has(lc)) { seen.add(lc); out.push(lc); }
            };
            for (const entry of targetEmails) {
              if ((entry || '').indexOf('@') < 0) continue; // Standort-Pattern o.ä. — nicht mailbar
              try {
                const grp = await getGroupMembers(entry);
                if (grp && grp.members && grp.members.length > 0) {
                  for (const m of grp.members) push(m.email);
                } else {
                  push(entry); // Einzelperson ODER nicht auflösbarer Verteiler
                }
              } catch { push(entry); }
            }
            // Fallback: Live-Auflösung ergab nichts (z.B. fehlende
            // Group.Read.All) → beim Event-Save eingefrorene Liste verwenden.
            if (out.length === 0 && (selectedEvent.audienceResolvedEmails || []).length > 0) {
              for (const e of selectedEvent.audienceResolvedEmails || []) push(e);
            }
            if (out.length > 0) resolvedRecipients = out;
            setInviteSending(false);
          }
          const confirmMsg = isDe
            ? (inviteTarget === 'organizer'
              ? `Einladungs-Mail an dich selbst (${myEmail}) senden? Du kannst sie anschließend aus Outlook an deinen Verteiler weiterleiten.`
              : `Einladungs-Mail an ${resolvedRecipients.length} aufgelöste Empfänger des Mailverteilers senden?\n\nDie Verteiler wurden in einzelne Mitglieder-Adressen aufgelöst; die Empfänger stehen im Bcc (sehen einander nicht), du selbst im An-Feld.\n\n${resolvedRecipients.slice(0, 12).join(', ')}${resolvedRecipients.length > 12 ? `, … (+${resolvedRecipients.length - 12})` : ''}`)
            : (inviteTarget === 'organizer'
              ? `Send invitation email to yourself (${myEmail})? You can then forward it from Outlook to your distribution list.`
              : `Send invitation email to ${resolvedRecipients.length} resolved recipients of the mail distribution?\n\nDistribution lists were resolved into individual member addresses; recipients are on Bcc (cannot see each other), you are in the To field.\n\n${resolvedRecipients.slice(0, 12).join(', ')}${resolvedRecipients.length > 12 ? `, … (+${resolvedRecipients.length - 12})` : ''}`);
          if (!(await confirmDialog(confirmMsg, { confirmLabel: isDe ? 'Senden' : 'Send' }))) return;
          setInviteSending(true);
          const resolvedSubject = replacePlaceholders(inviteSubject, previewVars);
          const resolvedHeading = replacePlaceholders(inviteHeading, previewVars);
          const resolvedBody = replacePlaceholders(inviteBody, previewVars);
          // v22.5: editierbare Unter-Überschrift verwenden (leer = „Event <Titel>").
          const resolvedSubheading = inviteSubheading && inviteSubheading.trim()
            ? replacePlaceholders(inviteSubheading, previewVars)
            : `Event ${selectedEvent.title}`;
          const fullBody = applyInviteHero(wrapTemplate('#86bc25', resolvedHeading, resolvedSubheading, resolvedBody, undefined, inviteHeaderOpts));
          const ccString = ccEmails.join(';');
          const recipientName = inviteTarget === 'organizer'
            ? myDisplayName
            : (inviteTarget === 'uninvited'
              ? (isDe ? 'Noch nicht Eingeladene' : 'Not-yet-invited')
              : inviteTarget === 'pending'
              ? (isDe ? 'Noch nicht Angemeldete' : 'Not-yet-registered')
              : (isDe ? 'Mailverteiler' : 'Mail distribution'));
          try {
            if (isBroadcast) {
              // v27.11: Aufgelöste Mitglieder in Chunks (Exchange-Limit ~500
              // Empfänger/Mail) per Bcc verschicken — wie beim Verteiler sehen
              // die Mitglieder einander nicht. To = der auslösende Organizer,
              // CC (übrige Organizer) nur auf dem ersten Chunk.
              const CHUNK = 450;
              for (let i = 0; i < resolvedRecipients.length; i += CHUNK) {
                const chunk = resolvedRecipients.slice(i, i + CHUNK);
                await eventServiceRef.queueEmail(
                  resolvedSubject, myEmail, recipientName, fullBody,
                  'Einladung', selectedEvent.title, selectedEvent.id,
                  (i === 0 && ccString) ? ccString : undefined,
                  chunk.join(';'),
                );
              }
            } else {
              await eventServiceRef.queueEmail(
                resolvedSubject, targetEmails.join(';'), recipientName, fullBody,
                'Einladung', selectedEvent.title, selectedEvent.id,
                ccString || undefined,
              );
            }
            // v26.69: NUR echte Broadcasts an den Mailverteiler ins Kommunikations-
            // Log schreiben. Der „An mich (zum Weiterleiten)"-Selbstversand
            // (inviteTarget === 'organizer') geht nur an die eigene Mailbox — das
            // ist eine Vorbereitung, KEINE Kommunikation an die Teilnehmer. Solche
            // Selbstversände dürfen den „Bereits versendete Infos"-Hinweis in
            // späteren Anmeldebestätigungen nicht auslösen und sollen auch nicht in
            // den event-bezogenen Nachrichten der Teilnehmer auftauchen.
            if (isBroadcast) {
              try { await eventServiceRef.logEventComm({ eventId: selectedEvent.id, eventTitle: selectedEvent.title, subject: resolvedSubject, bodyHtml: fullBody, emailType: 'Einladung' }); } catch { /* */ }
            }
            setInviteSending(false);
            showAlert(isDe
              ? `Einladungs-Mail an ${isBroadcast ? resolvedRecipients.length : targetEmails.length} Empfänger in die Warteschlange eingetragen.`
              : `Invitation email queued for ${isBroadcast ? resolvedRecipients.length : targetEmails.length} recipient(s).`);
            setShowInviteModal(false);
          } catch {
            setInviteSending(false);
            showAlert(isDe ? 'Fehler beim Eintragen der E-Mail.' : 'Error queueing the email.');
          }
        };
        const headerExtra = (
          <div style={{
            padding: 12,
            background: 'var(--dex-gray-50, #fafafa)',
            border: '1px solid var(--dex-gray-200)',
            borderRadius: 'var(--dex-radius)',
            marginBottom: 4,
          }}>
            <div style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 8 }}>
              {isDe ? 'Empfänger' : 'Recipient'}
            </div>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, cursor: 'pointer', fontSize: '0.82rem' }}>
              <input
                type="radio"
                name="inviteTarget"
                checked={inviteTarget === 'organizer'}
                onChange={() => setInviteTarget('organizer')}
                style={{ marginTop: 3 }}
              />
              <span>
                <strong>{isDe ? 'An mich — zum Weiterleiten' : 'To me — for forwarding'}</strong>
                <br />
                <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem' }}>
                  {myEmail}
                </span>
              </span>
            </label>
            {/* v28.37: Zwei Nachfass-Modi. Beide arbeiten auf dem Verteiler und
                ziehen davon ab, wer schon „durch" ist — einmal gemessen an den
                bereits verschickten Einladungen, einmal an der Teilnehmerliste. */}
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, cursor: audienceEmails.length === 0 ? 'not-allowed' : 'pointer', fontSize: '0.82rem', opacity: audienceEmails.length === 0 ? 0.55 : 1 }}>
              <input
                type="radio"
                name="inviteTarget"
                checked={inviteTarget === 'uninvited'}
                onChange={() => { setInviteTarget('uninvited'); setInviteCustomEmails(null); }}
                disabled={audienceEmails.length === 0}
                style={{ marginTop: 3 }}
              />
              <span style={{ flex: 1 }}>
                <strong>{isDe ? `Nur an noch nicht Eingeladene (${uninvitedEmails.length})` : `Only to not-yet-invited (${uninvitedEmails.length})`}</strong>
                <br />
                <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem' }}>
                  {invitedLc === null
                    ? (isDe ? 'Frühere Einladungen werden geladen …' : 'Loading earlier invitations …')
                    : (isDe
                      ? `Abgleich gegen bereits verschickte Einladungsmails — ${alreadyInvitedCount} Adresse(n) fallen raus. Hinweis: Versendete Mails werden nach rund einem Monat archiviert; ältere Einladungsrunden sind darin nicht mehr enthalten.`
                      : `Compared against invitations already sent — ${alreadyInvitedCount} address(es) excluded. Note: sent mails are archived after about a month, so older rounds are no longer included.`)}
                </span>
              </span>
            </label>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, cursor: audienceEmails.length === 0 ? 'not-allowed' : 'pointer', fontSize: '0.82rem', opacity: audienceEmails.length === 0 ? 0.55 : 1 }}>
              <input
                type="radio"
                name="inviteTarget"
                checked={inviteTarget === 'pending'}
                onChange={() => { setInviteTarget('pending'); setInviteCustomEmails(null); }}
                disabled={audienceEmails.length === 0}
                style={{ marginTop: 3 }}
              />
              <span style={{ flex: 1 }}>
                <strong>{isDe ? `Nur an noch nicht Angemeldete (${pendingEmails.length})` : `Only to not-yet-registered (${pendingEmails.length})`}</strong>
                <br />
                <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem' }}>
                  {isDe
                    ? `Abgleich gegen die Teilnehmerliste — ${alreadyDecidedCount} Adresse(n) haben sich bereits an- oder abgemeldet und fallen raus.`
                    : `Compared against the participant list — ${alreadyDecidedCount} address(es) have already registered or cancelled and are excluded.`}
                </span>
              </span>
            </label>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: audienceEmails.length === 0 ? 'not-allowed' : 'pointer', fontSize: '0.82rem', opacity: audienceEmails.length === 0 ? 0.55 : 1 }}>
              <input
                type="radio"
                name="inviteTarget"
                checked={inviteTarget === 'audience'}
                onChange={() => { setInviteTarget('audience'); setInviteCustomEmails(null); }}
                disabled={audienceEmails.length === 0}
                style={{ marginTop: 3 }}
              />
              <span style={{ flex: 1 }}>
                <strong>
                  {isDe
                    ? `An Mailverteiler des Events (${audienceEmails.length})`
                    : `To event mail distribution (${audienceEmails.length})`}
                </strong>
                <br />
                <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem', wordBreak: 'break-word' }}>
                  {audienceEmails.length === 0
                    ? (isDe
                      ? 'Kein Mailverteiler auf dem Event hinterlegt — in Schritt 3 (Sichtbarkeit) im Event-Edit ergänzen.'
                      : 'No mail distribution configured — add recipients in step 3 (Visibility) of event edit.')
                    : (isDe ? 'Die Adressen stehen unten und lassen sich vor dem Senden anpassen.' : 'The addresses are listed below and can be adjusted before sending.')}
                </span>
                {blockedInAudience.length > 0 && (
                  <div style={{
                    marginTop: 6, padding: '6px 8px',
                    background: '#fef3f2', border: '1px solid #c9302c',
                    borderRadius: 6, color: '#7a1f1c',
                    fontSize: '0.75rem', lineHeight: 1.4,
                  }}>
                    <strong>
                      {isDe ? '⚠ Blockierte Empfänger im Mailverteiler:' : '⚠ Blocked recipients in the distribution list:'}
                    </strong>
                    <ul style={{ margin: '4px 0 0 16px', padding: 0 }}>
                      {blockedInAudience.map(b => (
                        <li key={b.email}><code>{b.email}</code> — {b.reason}</li>
                      ))}
                    </ul>
                    <div style={{ marginTop: 4 }}>
                      {isDe
                        ? 'Pauschale Standort- oder All-Verteiler sind für Einladungs-Mails nicht zulässig. Bitte aus dem Mailverteiler entfernen (Event-Edit, Schritt 3) — sonst wird das Senden blockiert.'
                        : 'Entire location or all-distribution lists are not allowed for invitation emails. Please remove from the distribution list (event edit, step 3) — otherwise sending is blocked.'}
                    </div>
                  </div>
                )}
              </span>
            </label>
            {/* v28.37: Empfaengerliste — eingeklappt (sie kann mehrere hundert
                Adressen haben und schob den Dialog vorher auseinander) und vor
                dem Senden anpassbar: einzelne rausnehmen oder ergaenzen. Eine
                ergaenzte Adresse kann auf Wunsch direkt in den Event-Verteiler
                übernommen werden, damit sie beim nächsten Mal automatisch
                dabei ist. */}
            {inviteTarget !== 'organizer' && audienceEmails.length > 0 && (
              <div style={{ marginTop: 10, border: '1px solid var(--dex-gray-200)', borderRadius: 8, overflow: 'hidden' }}>
                <button
                  type="button"
                  onClick={() => setInviteAudienceOpen(o => !o)}
                  style={{ width: '100%', textAlign: 'left', background: 'var(--dex-gray-50, #f7f7f5)', border: 'none', cursor: 'pointer', padding: '8px 10px', fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-700)', display: 'flex', alignItems: 'center', gap: 8 }}
                >
                  <span style={{ transform: inviteAudienceOpen ? 'rotate(90deg)' : 'none', transition: 'transform .15s' }}>▶</span>
                  {isDe ? `Empfänger anzeigen und anpassen (${effectiveEmails.length})` : `Show and adjust recipients (${effectiveEmails.length})`}
                  {inviteCustomEmails && (
                    <span style={{ marginLeft: 'auto', fontSize: '0.7rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                      {isDe ? 'angepasst' : 'adjusted'}
                    </span>
                  )}
                </button>
                {inviteAudienceOpen && (
                  <div style={{ padding: 10 }}>
                    {effectiveEmails.length === 0 && (
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginBottom: 8 }}>
                        {isDe ? 'Keine Empfänger übrig — es würde niemand angeschrieben.' : 'No recipients left — nobody would be contacted.'}
                      </div>
                    )}
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, maxHeight: 190, overflowY: 'auto' }}>
                      {effectiveEmails.map(em => (
                        <span key={em} style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '3px 6px 3px 9px', borderRadius: 999, background: 'var(--dex-gray-100, #f0f0ee)', fontSize: '0.74rem' }}>
                          {em}
                          <button
                            type="button"
                            title={isDe ? 'Aus dieser Mail entfernen' : 'Remove from this mail'}
                            onClick={() => setInviteCustomEmails(effectiveEmails.filter(x => x !== em))}
                            style={{ border: 'none', background: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)', fontSize: '0.85rem', lineHeight: 1, padding: 0 }}
                          >×</button>
                        </span>
                      ))}
                    </div>
                    <div style={{ display: 'flex', gap: 6, marginTop: 10 }}>
                      <input
                        type="text"
                        value={inviteAddInput}
                        onChange={e => setInviteAddInput(e.target.value)}
                        placeholder={isDe ? 'Adresse ergänzen …' : 'Add address …'}
                        style={{ flex: 1, minWidth: 0, height: 30, fontSize: '0.78rem', padding: '0 8px', border: '1px solid var(--dex-gray-300)', borderRadius: 6 }}
                      />
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.76rem', padding: '5px 12px' }}
                        onClick={() => {
                          (async () => {
                            const addr = inviteAddInput.trim();
                            if (!addr) return;
                            if (addr.indexOf('@') <= 0) {
                              showAlert(isDe ? 'Bitte eine gültige E-Mail-Adresse eingeben.' : 'Please enter a valid email address.', { variant: 'error' });
                              return;
                            }
                            const lc = addr.toLowerCase();
                            if (effectiveEmails.some(x => x.toLowerCase() === lc)) {
                              showAlert(isDe ? 'Diese Adresse steht bereits in der Liste.' : 'That address is already in the list.', { variant: 'info' });
                              setInviteAddInput('');
                              return;
                            }
                            setInviteCustomEmails(effectiveEmails.concat([addr]));
                            setInviteAddInput('');
                            // Noch nicht im Event-Verteiler? Dann anbieten, sie
                            // dauerhaft aufzunehmen — sonst fällt sie beim
                            // nächsten Versand wieder raus.
                            if (!audienceEmails.some(x => x.toLowerCase() === lc)) {
                              const ok = await confirmDialog(
                                isDe
                                  ? `„${addr}" ist noch nicht im Mailverteiler des Events.\n\nSoll die Adresse dauerhaft in den Verteiler aufgenommen werden? Dann sieht die Person das Event auch in ihrer Übersicht und ist bei künftigen Mails automatisch dabei.\n\nNein = die Adresse bekommt nur diese eine Mail.`
                                  : `„${addr}" is not in the event mail distribution yet.\n\nAdd it permanently? The person will then also see the event in their overview and be included in future mails.\n\nNo = the address only receives this one mail.`,
                                { confirmLabel: isDe ? 'In den Verteiler aufnehmen' : 'Add to distribution' },
                              );
                              if (ok) {
                                const next = audienceEmails.concat([addr]);
                                const saved = await updateEvent(selectedEvent.id, { 'Audience': next.join(',') });
                                if (saved) {
                                  await refreshEvents();
                                  showAlert(isDe ? 'Adresse in den Mailverteiler des Events aufgenommen.' : 'Address added to the event mail distribution.', { variant: 'success' });
                                } else {
                                  showAlert(isDe ? 'Der Verteiler konnte nicht gespeichert werden — die Adresse bekommt nur diese Mail.' : 'Could not save the distribution list — the address only receives this mail.', { variant: 'error' });
                                }
                              }
                            }
                          })().catch(() => { /* */ });
                        }}
                      >
                        {isDe ? 'Hinzufügen' : 'Add'}
                      </button>
                    </div>
                    {inviteCustomEmails && (
                      <button
                        type="button"
                        onClick={() => setInviteCustomEmails(null)}
                        style={{ marginTop: 8, border: 'none', background: 'none', cursor: 'pointer', padding: 0, fontSize: '0.74rem', color: 'var(--dex-gray-600)', textDecoration: 'underline' }}
                      >
                        {isDe ? 'Anpassungen verwerfen und Auswahl oben verwenden' : 'Discard changes and use the selection above'}
                      </button>
                    )}
                  </div>
                )}
              </div>
            )}
            {ccEmails.length > 0 && (
              <div style={{
                marginTop: 10, paddingTop: 8,
                borderTop: '1px dashed var(--dex-gray-200)',
                fontSize: '0.78rem', color: 'var(--dex-gray-600)',
              }}>
                <strong style={{ color: 'var(--dex-gray-700)' }}>{isDe ? 'CC' : 'CC'}: </strong>
                <span style={{ wordBreak: 'break-word' }}>{ccEmails.join(', ')}</span>
                <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 3 }}>
                  {isDe
                    ? 'Alle Organizer dieses Events werden automatisch in CC gesetzt.'
                    : 'All organizers of this event are automatically added in CC.'}
                </div>
              </div>
            )}
            {/* v30.51: Zusätzliches CC per Personensuche — dieselbe Bedienung
                wie bei der Massenmail und im F&A Center. */}
            <div style={{ marginTop: 10, paddingTop: 8, borderTop: '1px dashed var(--dex-gray-200)' }}>
              <RecipientPicker
                label={isDe ? 'Zusätzlich auf CC' : 'Additional CC'}
                hint={isDe
                  ? 'Personen über die Suche, Funktionspostfächer im Feld darunter. Die Organizer stehen ohnehin auf CC.'
                  : 'People via search, shared mailboxes in the field below. The organizers are on CC anyway.'}
                emptyText={isDe ? 'Kein zusätzliches CC — es gehen nur die Organizer mit.' : 'No additional CC — only the organizers.'}
                value={inviteCc}
                onChange={setInviteCc}
                searchUsers={searchUsers}
                searchUserByEmail={searchUser}
                disabled={inviteSending}
              />
            </div>
            {/* v30.52: dieselbe Auswahl wie in Massen- und QR-Mail. */}
            <div style={{ marginTop: 10, paddingTop: 8, borderTop: '1px dashed var(--dex-gray-200)' }}>
              <MailHeaderImageChooser
                value={inviteHeaderImage}
                onChange={setInviteHeaderImage}
                eventPhotoB64={inviteEventPhotoB64}
                disabled={inviteSending}
                onCrop={() => setComposerCrop('invite')}
                isDe={isDe}
              />
            </div>
            {/* v22.5/v22.6: Entwurf speichern (Button) + Auto-Speichern-Hinweis
                + Zurücksetzen. */}
            <div style={{ marginTop: 10, paddingTop: 8, borderTop: '1px dashed var(--dex-gray-200)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={saveInviteDraft}
                  style={{ fontSize: '0.78rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }}
                >
                  <Check size={14} />
                  {isDe ? 'Entwurf speichern' : 'Save draft'}
                </button>
                {inviteDraftSaved && (
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, fontSize: '0.78rem' }}>
                    <Check size={14} /> {isDe ? 'Gespeichert' : 'Saved'}
                  </span>
                )}
                <span style={{ flex: 1 }} />
                <button
                  type="button"
                  onClick={resetInviteDraft}
                  style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600, fontSize: '0.74rem', textDecoration: 'underline' }}
                >
                  {isDe ? 'Auf Standardtext zurücksetzen' : 'Reset to default text'}
                </button>
              </div>
              <div style={{ marginTop: 6, fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                {isDe
                  ? 'Dein Text wird zusätzlich automatisch gespeichert und beim nächsten Öffnen wiederhergestellt.'
                  : 'Your text is also saved automatically and restored next time you open it.'}
              </div>
            </div>
          </div>
        );
        // v28.38 BUG-FIX: Die „An:"-Zeile der Vorschau zeigte immer die volle
        // Verteilergröße — auch in den Nachfass-Modi und nach dem Entfernen
        // einzelner Adressen. Sie muss die WIRKLICH adressierte Liste nennen,
        // sonst widerspricht sie dem Senden-Knopf direkt daneben.
        const previewToLine = inviteTarget === 'organizer'
          ? myEmail
          : (isDe
            ? `${targetEmails.length} ${targetEmails.length === 1 ? 'Empfänger' : 'Empfänger'}${
              inviteCustomEmails ? ' (angepasste Auswahl)'
                : inviteTarget === 'uninvited' ? ' — noch nicht eingeladen'
                : inviteTarget === 'pending' ? ' — noch nicht angemeldet'
                : ' des Mailverteilers'}`
            : `${targetEmails.length} recipient(s)${
              inviteCustomEmails ? ' (adjusted selection)'
                : inviteTarget === 'uninvited' ? ' — not yet invited'
                : inviteTarget === 'pending' ? ' — not yet registered'
                : ' of the mail distribution'}`);
        const previewSubjectLine = replacePlaceholders(inviteSubject, previewVars);
        return (
          <HtmlEditorModal
            open={showInviteModal}
            onClose={() => !inviteSending && setShowInviteModal(false)}
            title={isDe
              ? `Einladungsmail: ${selectedEvent.title}`
              : `Invitation email: ${selectedEvent.title}`}
            value={inviteBody}
            onChange={setInviteBody}
            previewMode="email"
            emailSubject={inviteSubject}
            onEmailSubjectChange={setInviteSubject}
            emailHeading={inviteHeading}
            onEmailHeadingChange={setInviteHeading}
            emailSubheading={inviteSubheading}
            onEmailSubheadingChange={setInviteSubheading}
            emailHeadingColor="#86bc25"
            previewToLine={previewToLine}
            previewSubjectLine={previewSubjectLine}
            previewVars={previewVars}
            insertableVars={[
              { key: '{{EventTitle}}', label: isDe ? 'Event-Titel' : 'Event title' },
              { key: '{{Link}}', label: isDe ? 'Anmelde-Link' : 'Registration link' },
              { key: '{{Organizer}}', label: 'Organizer' },
            ]}
            imageBase64={(inviteHeaderImage.hero === 'event' && inviteEventPhotoB64) ? inviteEventPhotoB64 : customLogo}
            imageWidth={inviteHeaderImage.width}
            imagePaddingV={inviteHeaderImage.paddingV}
            imagePaddingH={inviteHeaderImage.paddingH}
            onImageWidthChange={(w) => setInviteHeaderImage(p => ({ ...p, width: w }))}
            onImagePaddingVChange={(v) => setInviteHeaderImage(p => ({ ...p, paddingV: v }))}
            onImagePaddingHChange={(h) => setInviteHeaderImage(p => ({ ...p, paddingH: h }))}
            headerExtra={headerExtra}
            extraAction={{
              label: inviteSending
                ? (isDe ? 'Wird eingetragen…' : 'Queueing…')
                : (isDe ? `Senden — ${recipientLabel}` : `Send — ${recipientLabel}`),
              onClick: sendAction,
              disabled: inviteSending
                || !inviteSubject.trim()
                || !inviteBody.trim()
                || targetEmails.length === 0,
              icon: <Send size={16} />,
            }}
          />
        );
};

