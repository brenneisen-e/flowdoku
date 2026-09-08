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
import { AlertCircle, Check, ChevronRight, Plus, Send, X } from '../../Icons';
import { HtmlEditorModal } from '../../HtmlEditorModal';
// v31.2: Gemeinsame UI-Klassen — das Stylesheet hängt HtmlEditorModal beim
// Öffnen selbst ein (`ensureDexUiStyles`), hier braucht es nur `cx`.
import { cx } from '../../dexUi';
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
  /** v30.67 (Review): undefined = lädt, null = nicht lesbar, Set = bekannt. */
  invitedLc: Set<string> | null | undefined;
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
        // v31.2: Der einzige Hook steht vor allem anderen; die Komponente hat
        // keine frühen Returns, die Reihenfolge ist damit fest. Das CC ist
        // Feineinstellung — der Aufklapper startet zu, der Zähler im Knopf
        // zeigt trotzdem, wie viele eine Kopie bekommen. Ausnahme: Steht beim
        // Öffnen schon ein Zusatz-CC, startet er offen — eine getroffene Wahl
        // darf nicht versteckt beginnen.
        const [ccOpen, setCcOpen] = React.useState(inviteCc.length > 0);
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
        // v30.67 (Review): Solange die Einladungen nicht BEKANNT sind (lädt
        // oder nicht lesbar), gibt es keine „noch nicht Eingeladenen" — der
        // Modus ist gesperrt und der Sende-Pfad blockt zusätzlich (s.u.).
        const invitedKnown = !!invitedLc;
        const uninvitedEmails = invitedKnown
          ? audienceEmails.filter(e => {
            const lc = e.toLowerCase();
            if (lc.indexOf('@') < 0) return true; // Gruppe → nicht aufloesbar, mitnehmen
            return !invitedLc.has(lc);
          })
          : [];
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
              : 'An alle im Mailverteiler'} (${nRecipients === 0 ? 'leer' : nRecipients + ' Empfänger'})`
            : `${inviteCustomEmails
              ? 'To adjusted selection'
              : inviteTarget === 'uninvited' ? 'To not-yet-invited'
              : inviteTarget === 'pending' ? 'To not-yet-registered'
              : 'To everyone on the mail distribution'} (${nRecipients === 0 ? 'empty' : nRecipients + ' recipients'})`);
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
          // v30.67 (Review): Der Modus kann noch gewählt sein, wenn das Lesen
          // der Einladungen NACH der Auswahl gescheitert ist — dann nie an
          // „alle" senden, sondern abbrechen.
          if (inviteTarget === 'uninvited' && !invitedKnown) {
            showAlert(isDe
              ? 'Die bereits verschickten Einladungen konnten nicht gelesen werden — ein Abgleich „Nur an noch nicht Eingeladene" ist gerade nicht möglich. Bitte schließe den Dialog und öffne ihn erneut, oder wähle einen anderen Empfängerkreis.'
              : 'The invitations already sent could not be read — "Only to not-yet-invited" cannot be determined right now. Please close and reopen the dialog, or choose a different recipient group.', { variant: 'error' });
            return;
          }
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
        // v31.2: Der „Hinzufügen"-Pfad steht einmal hier statt inline im
        // Knopf — der Block ist mit Rückfrage und Verteiler-Speichern zu lang
        // für ein onClick und wäre beim Umbau sonst zweimal zu prüfen.
        const addRecipient = (): void => {
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
        };
        // v31.2: Die Empfänger-Frage als Auswahl-Kacheln (eine von vier, jede
        // mit einer Zeile Folge) statt vier Radio-Zeilen mit Fließtext. Die
        // Kachel bindet dieselben Setter wie vorher das Radio: nur „An mich"
        // lässt eine Handanpassung der Liste stehen, die drei Verteiler-Modi
        // setzen sie zurück (v28.37). Ein Klick auf die schon aktive Kachel
        // ist ein No-op: Ein gesetztes Radio feuerte kein onChange, ein Button
        // feuert onClick immer — ohne den Guard verwarf ein Doppelklick die
        // Handanpassung der Empfängerliste.
        type InviteTargetKey = InviteComposerModalProps['inviteTarget'];
        const renderTargetTile = (opt: { key: InviteTargetKey; title: string; count?: string | number; desc: string; disabled?: boolean }): React.ReactElement => {
          const active = inviteTarget === opt.key;
          return (
            <button
              key={opt.key}
              type="button"
              role="radio"
              aria-checked={active}
              disabled={opt.disabled}
              className={cx('dex-ui-choice', active && 'is-active', opt.disabled && 'is-disabled')}
              onClick={() => {
                if (active) return;
                setInviteTarget(opt.key);
                if (opt.key !== 'organizer') setInviteCustomEmails(null);
              }}
            >
              <div className="dex-ui-choice-body">
                <div className="dex-ui-choice-title" style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  {opt.title}
                  {opt.count !== undefined && (
                    <span className={cx('dex-ui-pill', active ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')}>{opt.count}</span>
                  )}
                </div>
                <div className="dex-ui-choice-desc" style={{ wordBreak: 'break-word' }}>{opt.desc}</div>
              </div>
              <span className="dex-ui-choice-check" aria-hidden="true">{active && <Check size={12} />}</span>
            </button>
          );
        };
        const noAudience = audienceEmails.length === 0;
        const headerExtra = (
          <div>
            {/* ---- 1. An wen? ------------------------------------------ */}
            <div className="dex-ui-section">
              <div className="dex-ui-section-title">{isDe ? 'An wen geht die Mail?' : 'Who receives the email?'}</div>
              <div className="dex-ui-grid-2" role="radiogroup" aria-label={isDe ? 'Empfänger' : 'Recipients'}>
                {/* v31.2: „Nur an mich" ist die Vorgabe (inviteTarget-Default
                    'organizer') und steht deshalb oben links — die Blickführung
                    beginnt dort, und der Senden-Knopf „An mich (…)" überrascht
                    nicht mehr. */}
                {renderTargetTile({
                  key: 'organizer',
                  title: isDe ? 'Nur an mich' : 'Only to me',
                  desc: isDe
                    ? `${myEmail} — zum Prüfen oder zum Weiterleiten aus Outlook an deinen Verteiler.`
                    : `${myEmail} — to check it, or to forward it from Outlook to your distribution list.`,
                })}
                {renderTargetTile({
                  key: 'audience',
                  title: isDe ? 'An alle im Mailverteiler' : 'Everyone on the mail distribution',
                  count: audienceEmails.length,
                  disabled: noAudience,
                  desc: noAudience
                    ? (isDe
                      ? 'Kein Mailverteiler auf dem Event hinterlegt — ergänze ihn im Event-Edit, Schritt 3 (Sichtbarkeit).'
                      : 'No mail distribution configured — add recipients in event edit, step 3 (Visibility).')
                    : (isDe
                      ? 'Der komplette Verteiler des Events. Die Adressen stehen unten und lassen sich vor dem Senden anpassen.'
                      : 'The full event distribution list. The addresses are listed below and can be adjusted before sending.'),
                })}
                {/* v28.37: Zwei Nachfass-Modi. Beide arbeiten auf dem Verteiler und
                    ziehen davon ab, wer schon „durch" ist — einmal gemessen an den
                    bereits verschickten Einladungen, einmal an der Teilnehmerliste. */}
                {renderTargetTile({
                  key: 'uninvited',
                  title: isDe ? 'Nur an noch nicht Eingeladene' : 'Only those not yet invited',
                  count: invitedKnown ? uninvitedEmails.length : '–',
                  disabled: noAudience || !invitedKnown,
                  desc: invitedLc === undefined
                    ? (isDe ? 'Frühere Einladungen werden geladen …' : 'Loading earlier invitations …')
                    : invitedLc === null
                    ? (isDe
                      ? 'Abgleich nicht möglich — die bereits verschickten Einladungen konnten nicht gelesen werden. Schließe den Dialog und öffne ihn erneut.'
                      : 'Comparison not possible — the invitations already sent could not be read. Close and reopen the dialog.')
                    : (isDe
                      ? `Wer schon eine Einladungsmail bekommen hat, fällt raus — ${alreadyInvitedCount} Adresse(n).`
                      : `Whoever already received an invitation is excluded — ${alreadyInvitedCount} address(es).`),
                })}
                {renderTargetTile({
                  key: 'pending',
                  title: isDe ? 'Nur an noch nicht Angemeldete' : 'Only those not yet registered',
                  count: pendingEmails.length,
                  disabled: noAudience,
                  desc: isDe
                    ? `Wer sich schon an- oder abgemeldet hat, fällt raus — ${alreadyDecidedCount} Adresse(n) laut Teilnehmerliste.`
                    : `Whoever already registered or cancelled is excluded — ${alreadyDecidedCount} address(es) per participant list.`,
                })}
              </div>
              {/* v31.2: Die Grenze des Abgleichs gehört zum Zähler der Kachel,
                  nicht erst zur Wahl — sie steht, sobald die Kachel wählbar ist. */}
              {!noAudience && invitedKnown && (
                <div className="dex-ui-help">
                  {isDe
                    ? '„Nur an noch nicht Eingeladene" zählt nur die letzten rund vier Wochen: Versendete Mails werden nach etwa einem Monat archiviert, ältere Einladungsrunden sind im Abgleich nicht mehr enthalten.'
                    : '"Only those not yet invited" covers only the last four weeks or so: sent mails are archived after about a month, so older invitation rounds are no longer part of the comparison.'}
                </div>
              )}
              {blockedInAudience.length > 0 && (
                <div className="dex-ui-callout dex-ui-callout--danger" role="alert" style={{ marginTop: 10 }}>
                  <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                  <div style={{ minWidth: 0 }}>
                    <strong>{isDe ? 'Blockierte Empfänger im Mailverteiler' : 'Blocked recipients in the distribution list'}</strong>
                    <ul style={{ margin: '4px 0 0 16px', padding: 0 }}>
                      {blockedInAudience.map(b => (
                        <li key={b.email}><code>{b.email}</code> — {b.reason}</li>
                      ))}
                    </ul>
                    <div style={{ marginTop: 4 }}>
                      {isDe
                        ? 'Pauschale Standort- oder All-Verteiler sind für Einladungs-Mails nicht zulässig. Entferne sie aus dem Mailverteiler (Event-Edit, Schritt 3) — sonst bleibt das Senden blockiert.'
                        : 'Entire location or all-distribution lists are not allowed for invitation emails. Remove them from the distribution list (event edit, step 3) — otherwise sending stays blocked.'}
                    </div>
                  </div>
                </div>
              )}
              {/* v28.37: Empfaengerliste — eingeklappt (sie kann mehrere hundert
                  Adressen haben und schob den Dialog vorher auseinander) und vor
                  dem Senden anpassbar: einzelne rausnehmen oder ergaenzen. Eine
                  ergaenzte Adresse kann auf Wunsch direkt in den Event-Verteiler
                  übernommen werden, damit sie beim nächsten Mal automatisch
                  dabei ist. */}
              {inviteTarget !== 'organizer' && audienceEmails.length > 0 && (
                <div style={{ marginTop: 8 }}>
                  <button
                    type="button"
                    className={cx('dex-ui-disclosure', inviteAudienceOpen && 'is-open')}
                    aria-expanded={inviteAudienceOpen}
                    onClick={() => setInviteAudienceOpen(o => !o)}
                  >
                    <span className="dex-ui-disclosure-chevron"><ChevronRight size={16} /></span>
                    {isDe ? 'Empfänger anzeigen und anpassen' : 'Show and adjust recipients'}
                    {inviteCustomEmails && (
                      <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'angepasst' : 'adjusted'}</span>
                    )}
                    <span className="dex-ui-disclosure-count">{effectiveEmails.length}</span>
                  </button>
                  {inviteAudienceOpen && (
                    <div className="dex-ui-disclosure-body">
                      {effectiveEmails.length === 0 && (
                        <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 8 }}>
                          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                          <span>{isDe ? 'Keine Empfänger übrig — es würde niemand angeschrieben.' : 'No recipients left — nobody would be contacted.'}</span>
                        </div>
                      )}
                      <div className="dex-ui-inline" style={{ gap: 6, maxHeight: 190, overflowY: 'auto' }}>
                        {effectiveEmails.map(em => (
                          <span key={em} className="dex-ui-pill dex-ui-pill--gray" style={{ paddingRight: 3 }}>
                            {em}
                            <button
                              type="button"
                              className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                              style={{ width: 20, height: 20 }}
                              title={isDe ? 'Aus dieser Mail entfernen' : 'Remove from this mail'}
                              aria-label={isDe ? `${em} aus dieser Mail entfernen` : `Remove ${em} from this mail`}
                              onClick={() => setInviteCustomEmails(effectiveEmails.filter(x => x !== em))}
                            ><X size={11} /></button>
                          </span>
                        ))}
                      </div>
                      <div style={{ display: 'flex', gap: 6, marginTop: 10 }}>
                        <input
                          type="text"
                          className="dex-ui-input dex-ui-input--sm"
                          value={inviteAddInput}
                          onChange={e => setInviteAddInput(e.target.value)}
                          placeholder={isDe ? 'Adresse ergänzen, z. B. vorname.nachname@deloitte.de' : 'Add an address, e.g. first.last@deloitte.de'}
                          aria-label={isDe ? 'Adresse ergänzen' : 'Add address'}
                          style={{ flex: 1, minWidth: 0 }}
                        />
                        <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={addRecipient}>
                          <Plus size={14} />
                          {isDe ? 'Hinzufügen' : 'Add'}
                        </button>
                      </div>
                      <div className="dex-ui-help">
                        {isDe ? 'Die ergänzte Adresse bekommt diese Mail — auf Wunsch nimmst du sie dauerhaft in den Verteiler auf.' : 'The added address receives this mail — you can also add it to the distribution list permanently.'}
                      </div>
                      {inviteCustomEmails && (
                        <button
                          type="button"
                          className="dex-ui-textbtn dex-ui-textbtn--muted"
                          style={{ marginTop: 6, marginLeft: -8 }}
                          onClick={() => setInviteCustomEmails(null)}
                        >
                          {isDe ? 'Anpassungen verwerfen und Auswahl oben verwenden' : 'Discard changes and use the selection above'}
                        </button>
                      )}
                    </div>
                  )}
                </div>
              )}
              {/* ---- Kopie (CC) — eine Empfänger-Frage, deshalb hier und nicht
                   zwischen Optik und Inhalt; selten geändert, deshalb zu. ------ */}
              {/* v31.2: Automatisches Organizer-CC und das Zusatz-CC (v30.51)
                  standen als zwei Blöcke untereinander — jetzt EIN Aufklapper;
                  der Zähler im Knopf nennt, wie viele eine Kopie bekommen. */}
              <div style={{ marginTop: 4 }}>
                <button
                  type="button"
                  className={cx('dex-ui-disclosure', ccOpen && 'is-open')}
                  aria-expanded={ccOpen}
                  onClick={() => setCcOpen(o => !o)}
                >
                  <span className="dex-ui-disclosure-chevron"><ChevronRight size={16} /></span>
                  {isDe ? 'Wer bekommt eine Kopie (CC)?' : 'Who gets a copy (CC)?'}
                  <span className="dex-ui-disclosure-count">
                    {ccEmails.length > 0
                      ? (isDe ? `${ccEmails.length} Person(en)` : `${ccEmails.length} person(s)`)
                      : (isDe ? 'niemand' : 'nobody')}
                  </span>
                </button>
                {ccOpen && (
                  <div className="dex-ui-disclosure-body dex-ui-stack">
                    {ccEmails.length > 0 && (
                      <div className="dex-ui-callout dex-ui-callout--neutral">
                        <div style={{ minWidth: 0 }}>
                          <strong>CC: </strong>
                          <span style={{ wordBreak: 'break-word' }}>{ccEmails.join(', ')}</span>
                          <div className="dex-ui-help" style={{ marginTop: 3 }}>
                            {isDe
                              ? 'Alle Organizer dieses Events werden automatisch in CC gesetzt.'
                              : 'All organizers of this event are automatically added in CC.'}
                          </div>
                        </div>
                      </div>
                    )}
                    {/* v30.51: Zusätzliches CC per Personensuche — dieselbe Bedienung
                        wie bei der Massenmail und im F&A Center. */}
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
                )}
              </div>
            </div>

            {/* ---- 2. Wie sieht die Mail aus? --------------------------- */}
            {/* v30.52: dieselbe Auswahl wie in Massen- und QR-Mail. */}
            <div className="dex-ui-section">
              <div className="dex-ui-section-title">{isDe ? 'Wie sieht die Mail aus?' : 'What does the email look like?'}</div>
              <MailHeaderImageChooser
                value={inviteHeaderImage}
                onChange={setInviteHeaderImage}
                eventPhotoB64={inviteEventPhotoB64}
                disabled={inviteSending}
                onCrop={() => setComposerCrop('invite')}
                isDe={isDe}
              />
              <div className="dex-ui-help" style={{ marginTop: 0 }}>
                {isDe
                  ? 'Die Vorschau rechts zeigt Kopfbild, Überschrift und Text so, wie die Mail ankommt.'
                  : 'The preview on the right shows header image, heading and text as the email will arrive.'}
              </div>
            </div>

            {/* ---- 3. Was steht drin? — Betreff, Überschrift und Text folgen
                 direkt unter diesem Block (HtmlEditorModal rendert sie). ------ */}
            {/* v22.5/v22.6: Entwurf speichern (Button) + Auto-Speichern-Hinweis
                + Zurücksetzen. */}
            <div className="dex-ui-section" style={{ marginBottom: 4 }}>
              <div className="dex-ui-section-title">{isDe ? 'Was steht in der Mail?' : 'What does the email say?'}</div>
              <div className="dex-ui-inline">
                <span className="dex-ui-help" style={{ marginTop: 0, flex: 1, minWidth: 200 }}>
                  {isDe
                    ? 'Betreff, Überschrift und Text stehen direkt darunter. Dein Text wird automatisch gespeichert und beim nächsten Öffnen wiederhergestellt.'
                    : 'Subject, heading and text follow right below. Your text is saved automatically and restored next time you open it.'}
                </span>
                <button type="button" className="dex-ui-textbtn" onClick={saveInviteDraft}>
                  <Check size={14} />
                  {isDe ? 'Entwurf speichern' : 'Save draft'}
                </button>
                {inviteDraftSaved && (
                  <span className="dex-ui-pill dex-ui-pill--green"><Check size={12} /> {isDe ? 'Gespeichert' : 'Saved'}</span>
                )}
                <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={resetInviteDraft}>
                  {isDe ? 'Auf Standardtext zurücksetzen' : 'Reset to default text'}
                </button>
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

