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
import { AlertCircle, Check, ChevronDown, Plus, Send, Users, X } from '../../Icons';
import Modal from '../../Modal';
import { parsePastedRecipients } from '../../../utils/pastedRecipients';
import { HtmlEditorModal } from '../../HtmlEditorModal';
// v31.2: Gemeinsame UI-Klassen — das Stylesheet hängt HtmlEditorModal beim
// Öffnen selbst ein (`ensureDexUiStyles`), hier braucht es nur `cx`.
import { cx } from '../../dexUi';
import { DeloitteEvent } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';
import { MailHeaderImage, applyHeroImage, hasOwnHeaderImage, kopfBildVorschau, kopfMasseFuerBild, mailHeaderOpts } from '../../../utils/mailHeaderImage';
import { ladeKopfbild } from '../../../utils/inlineMailImage';

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
  /** v32.33: Zurück zur Frage „Was für eine Mail?" (MailTypeModal). */
  onZurueckZurArt?: () => void;
}

export const InviteComposerModal: React.FC<InviteComposerModalProps> = (p) => {
  const { applyInviteHero, confirmDialog, currentUser, eventServiceRef, getGroupMembers, inviteAddInput, inviteAudienceOpen, inviteBody, inviteCc, inviteCustomEmails, invitedLc, inviteDraftSaved, inviteEventPhotoB64, inviteHeaderImage, inviteHeading, inviteSending, inviteSubheading, inviteSubject, inviteTarget, isDe, refreshEvents, registrations, resetInviteDraft, saveInviteDraft, searchUser, searchUsers, selectedEvent, setComposerCrop, setInviteAddInput, setInviteAudienceOpen, setInviteBody, setInviteCc, setInviteCustomEmails, setInviteHeaderImage, setInviteHeading, setInviteSending, setInviteSubheading, setInviteSubject, setInviteTarget, setShowInviteModal, showAlert, showInviteModal, siteUrl, updateEvent, onZurueckZurArt } = p;
        // v31.2: Der einzige Hook steht vor allem anderen; die Komponente hat
        // keine frühen Returns, die Reihenfolge ist damit fest. Das CC ist
        // Feineinstellung — der Aufklapper startet zu, der Zähler im Knopf
        // zeigt trotzdem, wie viele eine Kopie bekommen. Ausnahme: Steht beim
        // Öffnen schon ein Zusatz-CC, startet er offen — eine getroffene Wahl
        // darf nicht versteckt beginnen.
        const [ccOpen, setCcOpen] = React.useState(inviteCc.length > 0);
        // v32.26: Empfänger und Kopfbild zugeklappt — wie im Massenmail-Editor
        // (Nutzer-Ansage 29.09.2026); die Zeile nennt Ziel und Anzahl trotzdem.
        const [aufAn, setAufAn] = React.useState(false);
        const [aufBild, setAufBild] = React.useState(false);
        // v32.29: „Nur intern“ hat drei Kreise — ich, alle Organizer, Organizer
        // plus Test-Team (Nutzer-Ansage 29.09.2026). Bewusst KEINE zwei neuen
        // Kacheln („das werden dann viele Kacheln“), sondern eine Kachel mit
        // Chip-Zeile darunter; `inviteTarget` bleibt 'organizer', der Kreis
        // lebt nur hier im Dialog. Intern heißt weiter: kein Bcc-Versand, kein
        // Eintrag im Kommunikations-Log.
        // v32.30: Mehrfachauswahl (Chips) — ich / Organizer / Test-Team sind
        // additiv und lassen sich mit einer Verteiler-Gruppe kombinieren.
        const [intern, setIntern] = React.useState<Set<'ich' | 'orgs' | 'test'>>(new Set());
        // Vorauswahl einmal, sobald die früheren Einladungen gelesen sind (Nutzer-
        // Ansage 29.09.2026: „sinnvoll vorausgewählt“): mit Verteiler „Noch nicht
        // Eingeladene“ (sonst „Alle im Verteiler“), ohne Verteiler „An mich“. Eine
        // vorgesetzte Liste (Offene Personen) oder eine eigene Wahl gewinnt.
        const vorgewaehltRef = React.useRef(false);
        // v32.30: „Woher kommt der Verteiler?“ — ein Klick zeigt ihn (Nutzer-
        // Ansage 29.09.2026), statt den Organizer in den Assistenten zu schicken.
        const [verteilerZeigen, setVerteilerZeigen] = React.useState(false);
        // v32.33: Eigenes Kopfbild auch in der Einladung (Nutzer-Ansage
        // 29.09.2026) — dieselbe Kachel und Leiter wie in Massen- und QR-Mail.
        const [eigenesKopfB64, setEigenesKopfB64] = React.useState('');
        const [kopfBusy, setKopfBusy] = React.useState(false);
        const [kopfNote, setKopfNote] = React.useState('');
        // v32.31: Eigene Liste (Nutzer-Ansage 29.09.2026: „bei der Einladungsmail
        // eine eigene Liste reinladen … mit der Sichtbarkeit abgleichen und
        // fragen, ob die Delta-Personen auch in die Sichtbarkeit sollen“). Die
        // Liste ist ein Zusatz zu den Gruppen, kein eigener Modus.
        // v32.31: „Sichtbarkeit prüfen“ — die Personen mit Namen, nach Nachname
        // sortiert, und ein Suchfeld „sieht X das Event?“ (Nutzer-Ansage
        // 29.09.2026). null = noch nicht geladen.
        const [sichtPersonen, setSichtPersonen] = React.useState<Array<{ email: string; vorname: string; nachname: string; standort: string; quelle: string }> | null>(null);
        const [sichtSuche, setSichtSuche] = React.useState('');
        const [eigeneOffen, setEigeneOffen] = React.useState(false);
        const [eigeneRoh, setEigeneRoh] = React.useState('');
        const [eigene, setEigene] = React.useState<string[]>([]);
        const [eigeneAn, setEigeneAn] = React.useState(false);
        const [eigeneSpeichern, setEigeneSpeichern] = React.useState(false);
        // v32.29: Erst „An wen?“, dann der Editor — wie bei Info-Mail und
        // Reminder (Nutzer-Ansage 29.09.2026). Kommt der Dialog mit einer schon
        // gesetzten Auswahl (Kasten „Offene Personen“ setzt die Liste vorab),
        // ist die Frage beantwortet und der Editor öffnet direkt.
        const [schritt, setSchritt] = React.useState<'wer' | 'mail'>(inviteCustomEmails ? 'mail' : 'wer');
        // v32.27: Verteiler in ihre Mitglieder auflösen (Nutzer-Befund 29.09.2026:
        // „An alle im Mailverteiler“ zeigte 1 — die Verteiler-Adresse selbst). Ein
        // Verteiler hat ein @ und lief deshalb als einzelne Person durch; damit
        // konnten „noch nicht Eingeladene/Angemeldete“ nichts herausrechnen. Die
        // Mitglieder löst der Assistent beim Speichern auf (AudienceResolvedEmails,
        // Personen bleiben darin sie selbst). Muster ohne @ (DEKOELN) stehen dort
        // nicht und werden wie bisher mitgenommen. Ohne aufgelöste Liste (altes
        // Event, nie gespeichert) gilt die Zielgruppe wie eingetragen.
        const audienceRoh = (selectedEvent.audienceFilter || [])
          .map(s => (s || '').trim())
          .filter(Boolean);
        const aufgeloest = selectedEvent.audienceResolvedEmails || [];
        // Ausgeschlossene sehen das Event nie — sie bekommen auch keine Einladung
        // (dieselbe Regel wie resolveAudienceEmails in useMailComposers).
        const ausgeschlossen = new Set((selectedEvent.excludedUsers || []).map(e => (e || '').toLowerCase().trim()).filter(Boolean));
        const audienceEmails = (aufgeloest.length > 0
          ? Array.from(new Set([...aufgeloest, ...audienceRoh.filter(e => e.indexOf('@') < 0)]))
          : audienceRoh).filter(e => !ausgeschlossen.has(e.toLowerCase()));
        const myEmail = currentUser.email || '';
        const myDisplayName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || myEmail;
        const eindeutig = (list: string[]): string[] => {
          const seen = new Set<string>();
          const out: string[] = [];
          for (const raw of list) {
            const e = (raw || '').trim();
            const lc = e.toLowerCase();
            if (!e || lc.indexOf('@') < 0 || seen.has(lc)) continue;
            seen.add(lc);
            out.push(e);
          }
          return out;
        };
        const internAdr = {
          ich: eindeutig([myEmail]),
          orgs: eindeutig([...(selectedEvent.organizerEmails || []), ...(selectedEvent.coOrganizerEmails || [])]),
          test: eindeutig(selectedEvent.testTeamEmails || []),
        };
        const internEmails = eindeutig([
          ...(intern.has('ich') ? internAdr.ich : []),
          ...(intern.has('orgs') ? internAdr.orgs : []),
          ...(intern.has('test') ? internAdr.test : []),
        ]);
        const internLabel = [
          intern.has('ich') ? (isDe ? 'mich' : 'me') : '',
          intern.has('orgs') ? 'Organizer' : '',
          intern.has('test') ? (isDe ? 'Test-Team' : 'test team') : '',
        ].filter(Boolean).join(' + ');
        const nurIch = intern.size === 1 && intern.has('ich');
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
        const eigeneAktiv = eigeneAn && eigene.length > 0;
        const targetEmails = inviteTarget === 'organizer'
          ? eindeutig([...internEmails, ...(eigeneAktiv ? eigene : [])])
          : eindeutig([...effectiveEmails.filter(e => e.indexOf('@') > 0), ...internEmails, ...(eigeneAktiv ? eigene : [])]).concat(effectiveEmails.filter(e => e.indexOf('@') < 0));
        /** Echter Massenversand (Verteiler ODER Nachfass) — nur „An mich" nicht. */
        // v32.31: Eine eigene Liste ist ein echter Versand an Teilnehmende —
        // Bcc, Auflösung, Kommunikations-Log wie beim Verteiler.
        const isBroadcast = inviteTarget !== 'organizer' || eigeneAktiv;
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
        const recipientLabel = !isBroadcast
          ? (nurIch
            ? (isDe ? `An mich (${myEmail})` : `To me (${myEmail})`)
            : targetEmails.length === 0 ? (isDe ? 'Noch niemand gewählt' : 'Nobody selected yet')
            : `${isDe ? 'An' : 'To'} ${internLabel} (${targetEmails.length} ${isDe ? 'Empfänger' : 'recipients'})`)
          : (isDe
            ? `${inviteCustomEmails
              ? 'An angepasste Auswahl'
              : inviteTarget === 'organizer' ? 'An deine eigene Liste'
              : inviteTarget === 'uninvited' ? 'An noch nicht Eingeladene'
              : inviteTarget === 'pending' ? 'An noch nicht Angemeldete'
              : 'An alle im Mailverteiler'} (${nRecipients === 0 ? 'leer' : nRecipients + ' Empfänger'})`
            : `${inviteCustomEmails
              ? 'To adjusted selection'
              : inviteTarget === 'organizer' ? 'To your own list'
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
        const mailLogo = selectedEvent.mailImageBase64 || customLogo;
        // v32.33: Kopf lokal rechnen — `applyInviteHero`/`inviteHeaderOpts` aus
        // dem Hook kennen das hochgeladene Bild dieser Einladung nicht.
        const kopfEinsetzen = (html: string): string =>
          (inviteHeaderImage.hero === 'custom' || inviteHeaderImage.hero === 'orb')
            ? applyHeroImage(html, inviteHeaderImage, inviteEventPhotoB64, eigenesKopfB64)
            : applyInviteHero(html);
        const kopfOpts = mailHeaderOpts(inviteHeaderImage, hasOwnHeaderImage(inviteHeaderImage, inviteEventPhotoB64, mailLogo, eigenesKopfB64));
        const eigenesKopfLaden = async (file: File): Promise<void> => {
          setKopfBusy(true); setKopfNote('');
          try {
            const out = await ladeKopfbild(file, isDe);
            setKopfNote(out.note);
            if (!out.dataUrl) return;
            setEigenesKopfB64(out.dataUrl);
            setInviteHeaderImage(prev => ({ ...prev, hero: 'custom', ...kopfMasseFuerBild(out.width, out.height) }));
          } finally { setKopfBusy(false); }
        };
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
                : 'Noch niemand gewählt — wähle unter „An wen geht die Mail?“ mindestens eine Gruppe.')
              : (isBroadcast
                ? 'No mail distribution list configured on the event. Please add recipients in step 3 (Visibility) first.'
                : 'Nobody selected yet — pick at least one group under “Who receives the email?”.'));
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
            ? (!isBroadcast
              ? (nurIch
                ? `Einladungs-Mail an dich selbst (${myEmail}) senden? Du kannst sie anschließend aus Outlook an deinen Verteiler weiterleiten.`
                : `Einladungs-Mail zur Probe an ${internLabel} senden (${targetEmails.length} Empfänger)?\n\n${targetEmails.join(', ')}\n\nDas ist ein interner Versand — er zählt nicht als Einladung an die Teilnehmenden.`)
              : `Einladungs-Mail an ${resolvedRecipients.length} aufgelöste Empfänger des Mailverteilers senden?\n\nDie Verteiler wurden in einzelne Mitglieder-Adressen aufgelöst; die Empfänger stehen im Bcc (sehen einander nicht), du selbst im An-Feld.\n\n${resolvedRecipients.slice(0, 12).join(', ')}${resolvedRecipients.length > 12 ? `, … (+${resolvedRecipients.length - 12})` : ''}`)
            : (inviteTarget === 'organizer'
              ? (nurIch
                ? `Send invitation email to yourself (${myEmail})? You can then forward it from Outlook to your distribution list.`
                : `Send a test invitation to ${internLabel} (${targetEmails.length} recipients)?\n\n${targetEmails.join(', ')}\n\nThis is an internal send — it does not count as an invitation to participants.`)
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
          const fullBody = kopfEinsetzen(wrapTemplate('#86bc25', resolvedHeading, resolvedSubheading, resolvedBody, undefined, kopfOpts));
          const ccString = ccEmails.join(';');
          const recipientName = !isBroadcast
            ? (nurIch ? myDisplayName : internLabel)
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
                // v32.31: die EINGETRAGENEN Einträge fortschreiben, nicht die
                // aufgelösten Mitglieder — sonst ersetzt ein Klick den Verteiler
                // durch seine Einzeladressen.
                const next = audienceRoh.concat([addr]);
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
        const noAudience = audienceEmails.length === 0;
        const keinVersandVerteiler = (() => {
          const mitAt = audienceRoh.filter(e => e.indexOf('@') > 0);
          if (mitAt.length === 0) return true; // leer oder nur Standort-Muster
          const gesperrt = new Set(getBlockedInviteRecipients(mitAt).map(b => b.email.toLowerCase()));
          return mitAt.every(e => gesperrt.has(e.toLowerCase()));
        })();
        const namenAus = (dn: string, first?: string, last?: string): { vorname: string; nachname: string } => {
          if (last || first) return { vorname: (first || '').trim(), nachname: (last || '').trim() };
          const d = (dn || '').replace(/\s*\(.*\)\s*$/, '').trim();
          const k = d.indexOf(',');
          if (k > 0) return { nachname: d.slice(0, k).trim(), vorname: d.slice(k + 1).trim() };
          const teile = d.split(/\s+/).filter(Boolean);
          return { vorname: teile.slice(0, -1).join(' '), nachname: teile.slice(-1)[0] || '' };
        };
        const sichtbarkeitLaden = (): void => {
          setVerteilerZeigen(true);
          if (sichtPersonen) return;
          (async () => {
            const map = new Map<string, { email: string; vorname: string; nachname: string; standort: string; quelle: string }>();
            for (const entry of audienceRoh) {
              if (entry.indexOf('@') < 0) continue; // Standort-Muster — keine Personenliste
              try {
                const grp = await getGroupMembers(entry);
                if (grp && grp.members && grp.members.length > 0) {
                  for (const m of grp.members) {
                    const lc = (m.email || '').toLowerCase();
                    if (!lc || ausgeschlossen.has(lc) || map.has(lc)) continue;
                    map.set(lc, { email: m.email, ...namenAus(m.displayName, m.firstName, m.lastName), standort: m.location || '', quelle: grp.groupName || entry });
                  }
                  continue;
                }
              } catch { /* Einzelperson oder nicht auflösbar */ }
              const lc = entry.toLowerCase();
              if (ausgeschlossen.has(lc) || map.has(lc)) continue;
              let dn = ''; let standort = '';
              try { const u = await searchUser(entry); dn = (u && u.displayName) || ''; standort = (u && u.location) || ''; } catch { /* */ }
              map.set(lc, { email: entry, ...namenAus(dn), standort, quelle: isDe ? 'Einzelperson' : 'Individual' });
            }
            // Beim Speichern eingefrorene Mitglieder, die die Live-Auflösung nicht
            // lieferte (fehlende Rechte), trotzdem zeigen — Name aus der Adresse.
            for (const e of aufgeloest) {
              const lc = (e || '').toLowerCase();
              if (!lc || map.has(lc) || ausgeschlossen.has(lc)) continue;
              map.set(lc, { email: e, vorname: '', nachname: e.split('@')[0], standort: '', quelle: isDe ? 'gespeicherte Auflösung' : 'saved resolution' });
            }
            const liste = Array.from(map.values()).sort((a, b) =>
              (a.nachname || a.email).localeCompare(b.nachname || b.email, 'de', { sensitivity: 'base' })
              || a.vorname.localeCompare(b.vorname, 'de', { sensitivity: 'base' }));
            setSichtPersonen(liste);
          })().catch(() => setSichtPersonen([]));
        };
        const bekanntLc = new Set([...audienceEmails, ...audienceRoh].map(e => e.toLowerCase()));
        const eigeneUebernehmen = (): void => {
          (async () => {
            const liste = eindeutig(parsePastedRecipients(eigeneRoh).map(x => x.email));
            if (liste.length === 0) {
              showAlert(isDe ? 'In der eingefügten Liste wurde keine E-Mail-Adresse gefunden.' : 'No email address found in the pasted list.', { variant: 'error' });
              return;
            }
            setEigene(liste);
            setEigeneAn(true);
            vorgewaehltRef.current = true;
            setEigeneOffen(false);
            // Abgleich mit der Sichtbarkeit: Wer nicht darin steht, sieht das
            // Event nicht und kommt über den Link nicht zur Anmeldung.
            const delta = liste.filter(e => !bekanntLc.has(e.toLowerCase()));
            if (delta.length === 0) return;
            const vorschau = delta.slice(0, 12).join(', ') + (delta.length > 12 ? `, … (+${delta.length - 12})` : '');
            const ok = await confirmDialog(
              isDe
                ? `${delta.length} von ${liste.length} Personen stehen noch nicht in der Sichtbarkeit des Events — sie sehen es nicht in ihrer Übersicht und können sich nicht anmelden.\n\n${vorschau}\n\nSollen sie in die Sichtbarkeit (Mailverteiler) aufgenommen werden?`
                : `${delta.length} of ${liste.length} people are not in the event’s visibility yet — they cannot see it or register.\n\n${vorschau}\n\nAdd them to the visibility (mail distribution)?`,
              { confirmLabel: isDe ? 'In die Sichtbarkeit aufnehmen' : 'Add to visibility' },
            );
            if (!ok) return;
            setEigeneSpeichern(true);
            try {
              const next = eindeutig([...audienceRoh, ...delta]).concat(audienceRoh.filter(e => e.indexOf('@') < 0));
              const saved = await updateEvent(selectedEvent.id, { 'Audience': Array.from(new Set(next)).join(',') });
              if (saved) {
                setSichtPersonen(null);
                await refreshEvents();
                showAlert(isDe ? `${delta.length} Personen in die Sichtbarkeit aufgenommen.` : `${delta.length} people added to the visibility.`, { variant: 'success' });
              } else {
                showAlert(isDe ? 'Die Sichtbarkeit konnte nicht gespeichert werden — die Personen bekommen die Einladung trotzdem, sehen das Event aber nicht.' : 'The visibility could not be saved — the people still get the invitation but cannot see the event.', { variant: 'error' });
              }
            } finally { setEigeneSpeichern(false); }
          })().catch(() => { setEigeneSpeichern(false); });
        };
        const zielAuswahl = (
            <>
              {/* v32.30: Chips statt Kacheln (Nutzer-Ansage 29.09.2026). Die drei
                  Verteiler-Gruppen schließen sich aus (jede ist eine Teilmenge der
                  ersten), ein zweiter Klick wählt ab; „Intern“ ist additiv. */}
              {(() => {
                const zeile = (titel: string, kinder: React.ReactNode): React.ReactElement => (
                  <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start', flexWrap: 'wrap', marginBottom: 10 }}>
                    <span className="dex-ui-muted" style={{ fontSize: '0.8rem', fontWeight: 600, minWidth: 110, paddingTop: 6 }}>{titel}</span>
                    <div className="dex-ui-inline" style={{ gap: 6, flexWrap: 'wrap', flex: '1 1 300px' }}>{kinder}</div>
                  </div>
                );
                const chip = (key: string, label: string, n: number | string, on: boolean, click: () => void, off?: boolean, title?: string): React.ReactElement => (
                  <button key={key} type="button" className={cx('dex-ui-chip', on && 'is-active')} aria-pressed={on} disabled={off} title={title} onClick={click}>
                    {on && <Check size={12} />} {label} <span style={{ opacity: 0.75, marginLeft: 2 }}>{n}</span>
                  </button>
                );
                const waehleVerteiler = (k: InviteTargetKey): void => {
                  vorgewaehltRef.current = true;
                  setInviteCustomEmails(null);
                  setInviteTarget(inviteTarget === k ? 'organizer' : k);
                };
                const toggleIntern = (k: 'ich' | 'orgs' | 'test'): void => {
                  vorgewaehltRef.current = true;
                  setIntern(prev => { const n = new Set(prev); if (n.has(k)) n.delete(k); else n.add(k); return n; });
                };
                const verteilerTitel = noAudience ? (isDe ? 'Kein Mailverteiler auf dem Event hinterlegt — ergänze ihn im Event-Edit, Schritt 3 (Sichtbarkeit).' : 'No mail distribution configured — add it in event edit, step 3 (Visibility).') : undefined;
                return (
                  <>
                    {zeile(isDe ? 'Mailverteiler' : 'Distribution', <>
                      {chip('audience', isDe ? 'Alle im Verteiler' : 'Everyone on the list', audienceEmails.length, inviteTarget === 'audience', () => waehleVerteiler('audience'), noAudience || keinVersandVerteiler, verteilerTitel)}
                      {chip('uninvited', isDe ? 'Noch nicht Eingeladene' : 'Not yet invited', invitedKnown ? uninvitedEmails.length : (invitedLc === undefined ? '…' : '–'), inviteTarget === 'uninvited', () => waehleVerteiler('uninvited'), noAudience || keinVersandVerteiler || !invitedKnown,
                        invitedLc === null ? (isDe ? 'Die bereits verschickten Einladungen konnten nicht gelesen werden — schließe den Dialog und öffne ihn erneut.' : 'The invitations already sent could not be read — close and reopen the dialog.') : (isDe ? `Wer schon eine Einladungsmail bekommen hat, fällt raus (${alreadyInvitedCount}).` : `Whoever already got an invitation is excluded (${alreadyInvitedCount}).`))}
                      {chip('pending', isDe ? 'Noch nicht Angemeldete' : 'Not yet registered', pendingEmails.length, inviteTarget === 'pending', () => waehleVerteiler('pending'), noAudience || keinVersandVerteiler,
                        isDe ? `Wer sich schon an- oder abgemeldet hat, fällt raus (${alreadyDecidedCount}).` : `Whoever already registered or cancelled is excluded (${alreadyDecidedCount}).`)}
                    </>)}
                    {/* v32.33: Ohne anschreibbaren Verteiler (nur Standort-Muster oder
                        pauschale Verteiler wie DEALL) kann DEX die Einladung nicht
                        verschicken — das Gruppenpostfach darf an diese großen
                        Verteiler nicht senden (Nutzer-Ansage 29.09.2026). Das muss
                        HIER stehen, nicht erst beim Senden. */}
                    {keinVersandVerteiler && (
                      <div className="dex-ui-callout dex-ui-callout--warn" style={{ margin: '-2px 0 12px' }}>
                        <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                        <div>
                          <strong>{isDe ? 'An diesen Verteiler kann DEX nicht senden.' : 'DEX cannot send to this distribution.'}</strong>{' '}
                          {isDe
                            ? 'Die Sichtbarkeit läuft über Standorte oder große Verteiler (z. B. DEALL). Über das Gruppenpostfach können wir an diese großen Verteiler nicht schicken — die Einladung schreibst du als Organizer deshalb außerhalb von DEX, z. B. direkt aus Outlook. Tipp: „An mich“ schicken und aus Outlook an den Verteiler weiterleiten.'
                            : 'Visibility runs via locations or large distribution lists (e.g. DEALL). The group mailbox cannot send to these — please send the invitation yourself outside DEX, e.g. from Outlook. Tip: send it “to me” and forward it from Outlook.'}
                        </div>
                      </div>
                    )}
                    <div className="dex-ui-help" style={{ margin: '-4px 0 10px', paddingLeft: 120 }}>
                      {isDe ? 'Der Mailverteiler kommt aus der Sichtbarkeit des Events (Assistent, Schritt 3).' : 'The distribution comes from the event’s visibility (wizard, step 3).'}
                      {' '}
                      <button type="button" className="dex-ui-textbtn" onClick={sichtbarkeitLaden}>
                        {isDe ? 'Sichtbarkeit prüfen' : 'Check visibility'}
                      </button>
                    </div>
                    {zeile(isDe ? 'Intern' : 'Internal', <>
                      {chip('ich', isDe ? 'An mich' : 'To me', internAdr.ich.length, intern.has('ich'), () => toggleIntern('ich'), internAdr.ich.length === 0)}
                      {chip('orgs', 'Organizer', internAdr.orgs.length, intern.has('orgs'), () => toggleIntern('orgs'), internAdr.orgs.length === 0)}
                      {chip('test', isDe ? 'Test-Team' : 'Test team', internAdr.test.length, intern.has('test'), () => toggleIntern('test'), internAdr.test.length === 0,
                        internAdr.test.length === 0 ? (isDe ? 'Für dieses Event ist kein Test-Team eingetragen.' : 'No test team on this event.') : undefined)}
                    </>)}
                    {zeile(isDe ? 'Eigene Liste' : 'Own list', <>
                      {eigene.length > 0 && chip('eigene', isDe ? 'Eigene Liste' : 'Own list', eigene.length, eigeneAn, () => { vorgewaehltRef.current = true; setEigeneAn(o => !o); })}
                      <button type="button" className="dex-ui-textbtn" disabled={eigeneSpeichern} onClick={() => setEigeneOffen(true)}>
                        {eigene.length > 0 ? (isDe ? 'Liste ändern …' : 'Change list …') : (isDe ? 'Adressliste einfügen …' : 'Paste address list …')}
                      </button>
                      {eigeneAktiv && (() => {
                        const n = eigene.filter(e => !bekanntLc.has(e.toLowerCase())).length;
                        return n > 0 ? <span className="dex-ui-help" style={{ margin: 0 }}>{isDe ? `${n} davon nicht in der Sichtbarkeit` : `${n} of them not in the visibility`}</span> : null;
                      })()}
                    </>)}
                    {eigeneOffen && (
                      <Modal open={true} onClose={() => setEigeneOffen(false)} maxWidth={720}
                        title={isDe ? 'Eigene Liste einfügen' : 'Paste your own list'}
                        subtitle={isDe ? 'Aus Outlook kopiert oder beliebig formatiert — DEX erkennt die Adressen und gleicht sie mit der Sichtbarkeit ab.' : 'Copied from Outlook or any format — DEX picks out the addresses and checks them against the visibility.'}
                        icon={<Users size={20} />}
                        footer={<>
                          <button type="button" className="btn btn-secondary" onClick={() => setEigeneOffen(false)}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
                          <button type="button" className="btn btn-primary" onClick={eigeneUebernehmen} disabled={!eigeneRoh.trim()}>
                            {isDe ? `Übernehmen (${parsePastedRecipients(eigeneRoh).length})` : `Apply (${parsePastedRecipients(eigeneRoh).length})`}
                          </button>
                        </>}>
                        <textarea className="form-control" rows={10} value={eigeneRoh} onChange={e => setEigeneRoh(e.target.value)}
                          placeholder={isDe ? 'Max Mustermann <mmustermann@deloitte.de>; erika@deloitte.de …' : 'Jane Doe <jdoe@deloitte.com>; john@deloitte.com …'}
                          style={{ width: '100%', fontFamily: 'inherit', fontSize: '0.9rem' }} />
                      </Modal>
                    )}
                    <div className="dex-ui-inline" style={{ gap: 6 }}>
                      <span className={cx('dex-ui-pill', targetEmails.length > 0 ? 'dex-ui-pill--green' : 'dex-ui-pill--red')}>{targetEmails.length} {isDe ? 'Empfänger' : 'recipients'}</span>
                      {!isBroadcast && intern.size > 0 && (
                        <span className="dex-ui-help" style={{ margin: 0 }}>{isDe ? 'Nur intern — zählt nicht als Einladung an die Teilnehmenden.' : 'Internal only — does not count as an invitation to participants.'}</span>
                      )}
                    </div>
                    {!isBroadcast && intern.size > 0 && !nurIch && (
                      <div className="dex-ui-help" style={{ wordBreak: 'break-word' }}>{targetEmails.join(', ')}</div>
                    )}
                    {verteilerZeigen && (() => {
                      const q = sichtSuche.trim().toLowerCase();
                      const treffer = (sichtPersonen || []).filter(x => !q
                        || x.email.toLowerCase().indexOf(q) >= 0
                        || `${x.vorname} ${x.nachname}`.toLowerCase().indexOf(q) >= 0
                        || `${x.nachname}, ${x.vorname}`.toLowerCase().indexOf(q) >= 0);
                      const muster = audienceRoh.filter(e => e.indexOf('@') < 0);
                      return (
                        <Modal open={true} onClose={() => setVerteilerZeigen(false)} maxWidth={860}
                          title={isDe ? 'Sichtbarkeit prüfen' : 'Check visibility'}
                          subtitle={isDe ? 'Wer das Event sieht — aus der Sichtbarkeit (Assistent, Schritt 3). Ändern kannst du sie dort oder über „Eigene Liste“.' : 'Who can see the event — from the visibility settings (wizard, step 3). Change them there or via “Own list”.'}
                          icon={<Users size={20} />}
                          footer={<button type="button" className="btn btn-secondary" onClick={() => setVerteilerZeigen(false)}>{isDe ? 'Schließen' : 'Close'}</button>}>
                          <div className="dex-ui-inline" style={{ gap: 6, flexWrap: 'wrap', marginBottom: 10 }}>
                            <span className="dex-ui-muted" style={{ fontSize: '0.8rem', fontWeight: 600 }}>{isDe ? 'Eingetragen:' : 'Entered:'}</span>
                            {audienceRoh.length === 0
                              ? <span className="dex-ui-help" style={{ margin: 0 }}>{isDe ? 'nichts — die Sichtbarkeit läuft nur über den Standort.' : 'nothing — visibility runs via location only.'}</span>
                              : audienceRoh.map(e => <span key={e} className="dex-ui-pill dex-ui-pill--gray dex-ui-pill--wrap">{e}</span>)}
                          </div>
                          {muster.length > 0 && (
                            <div className="dex-ui-help">{isDe ? `Standort-Muster (${muster.join(', ')}) sehen alle an diesem Standort — sie stehen nicht einzeln in der Liste.` : `Location patterns (${muster.join(', ')}) include everyone at that location — they are not listed individually.`}</div>
                          )}
                          <input className="form-control" value={sichtSuche} onChange={e => setSichtSuche(e.target.value)}
                            placeholder={isDe ? 'Name oder E-Mail — sieht diese Person das Event?' : 'Name or email — can this person see the event?'}
                            style={{ width: '100%', margin: '6px 0 8px' }} />
                          {q && sichtPersonen && (
                            <div className={cx('dex-ui-callout', treffer.length > 0 ? 'dex-ui-callout--success' : 'dex-ui-callout--warn')} style={{ marginBottom: 8 }}>
                              <span>{treffer.length > 0
                                ? (isDe ? `${treffer.length} Treffer — sieht das Event.` : `${treffer.length} match(es) — can see the event.`)
                                : (isDe ? 'Kein Treffer — diese Person sieht das Event nicht (außer über ein Standort-Muster).' : 'No match — this person cannot see the event (unless via a location pattern).')}</span>
                            </div>
                          )}
                          {sichtPersonen === null
                            ? <div className="dex-ui-help">{isDe ? 'Verteiler werden aufgelöst …' : 'Resolving distribution lists …'}</div>
                            : (
                              <div style={{ maxHeight: 360, overflowY: 'auto', border: '1px solid var(--dex-gray-200, #e5e5e5)', borderRadius: 8 }}>
                                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                                  <thead>
                                    <tr style={{ position: 'sticky', top: 0, background: 'var(--dex-gray-50, #fafafa)', textAlign: 'left' }}>
                                      <th style={{ padding: '6px 10px' }}>{isDe ? 'Nachname' : 'Last name'}</th>
                                      <th style={{ padding: '6px 10px' }}>{isDe ? 'Vorname' : 'First name'}</th>
                                      <th style={{ padding: '6px 10px' }}>{isDe ? 'E-Mail' : 'Email'}</th>
                                      <th style={{ padding: '6px 10px' }}>{isDe ? 'Standort' : 'Location'}</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {treffer.map(x => (
                                      <tr key={x.email} style={{ borderTop: '1px solid var(--dex-gray-100, #f0f0f0)' }}>
                                        <td style={{ padding: '5px 10px', fontWeight: 600 }}>{x.nachname}</td>
                                        <td style={{ padding: '5px 10px' }}>{x.vorname}</td>
                                        <td style={{ padding: '5px 10px', wordBreak: 'break-all' }}>{x.email}</td>
                                        <td style={{ padding: '5px 10px' }}>{x.standort}</td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            )}
                          <div className="dex-ui-help">
                            {sichtPersonen ? (isDe ? `${sichtPersonen.length} Personen` : `${sichtPersonen.length} people`) : ''}
                            {ausgeschlossen.size > 0 ? (isDe ? ` · ${ausgeschlossen.size} ausgeschlossene nicht enthalten` : ` · ${ausgeschlossen.size} excluded not included`) : ''}
                          </div>
                        </Modal>
                      );
                    })()}
                  </>
                );
              })()}
              {/* v31.2: Die Grenze des Abgleichs gehört zum Zähler der Kachel,
                  nicht erst zur Wahl — sie steht, sobald die Kachel wählbar ist. */}
              {!noAudience && !keinVersandVerteiler && invitedKnown && (
                <div className="dex-ui-help">
                  {isDe
                    ? '„Noch nicht Eingeladene" zählt nur die letzten rund vier Wochen: Versendete Mails werden nach etwa einem Monat archiviert, ältere Einladungsrunden sind im Abgleich nicht mehr enthalten.'
                    : '"Not yet invited" covers only the last four weeks or so: sent mails are archived after about a month, so older invitation rounds are no longer part of the comparison.'}
                </div>
              )}
              {blockedInAudience.length > 0 && !keinVersandVerteiler && (
                <div className="dex-ui-callout dex-ui-callout--danger" role="alert" style={{ marginTop: 10 }}>
                  <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                  <div style={{ minWidth: 0 }}>
                    <strong>{isDe ? 'Blockierte Empfänger im Mailverteiler' : 'Blocked recipients in the distribution list'}</strong>
                    {/* v31.10: `wordBreak` — eine gesperrte Verteiler-Adresse ist oft
                        länger als eine Handy-Zeile, und `<code>` bricht von sich aus
                        nicht um. */}
                    <ul style={{ margin: '4px 0 0 16px', padding: 0, wordBreak: 'break-word' }}>
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
            </>
        );
        const headerExtra = (
          <div>
            {/* ---- 1. An wen? ------------------------------------------ */}
            <div className="dex-ui-section">
              <button type="button" className={cx('dex-ui-disclosure', aufAn && 'is-open')} onClick={() => setAufAn(o => !o)} aria-expanded={aufAn}>
                <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                {isDe ? 'An wen geht die Mail?' : 'Who receives the email?'}
                <span className="dex-ui-pill dex-ui-pill--green" style={{ marginLeft: 8 }}>{targetEmails.length}</span>
                <span className="dex-ui-muted" style={{ marginLeft: 6, fontWeight: 400, fontSize: '0.8rem' }}>{recipientLabel}</span>
              </button>
              {aufAn && (
              <div className="dex-ui-disclosure-body">
              {zielAuswahl}
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
                    {/* v31.10: ChevronDown wie überall sonst. `.is-open` dreht den
                        Pfeil um 180° — aus dem ChevronRight wurde geöffnet ein Pfeil
                        nach LINKS, also die eine Richtung, die nichts bedeutet. */}
                    <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
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
                      {/* v31.10: Zwei Handy-Korrekturen an derselben Zeile.
                          `--wrap`: eine Pille bricht sonst nie um, und
                          „bernd.aussergewoehnlichlangername@deloitte.de" ragte bei
                          390 px über den Rand hinaus. Und das „×" war mit 20 px ein
                          Tippziel, das man mit dem Finger nicht trifft — 28 px sind
                          das Maximum, das die Zeilenhöhe einer Pille noch trägt. */}
                      <div className="dex-ui-inline" style={{ gap: 6, maxHeight: 190, overflowY: 'auto' }}>
                        {effectiveEmails.map(em => (
                          <span key={em} className="dex-ui-pill dex-ui-pill--gray dex-ui-pill--wrap" style={{ paddingRight: 3 }}>
                            {em}
                            <button
                              type="button"
                              className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                              style={{ width: 28, height: 28 }}
                              title={isDe ? 'Aus dieser Mail entfernen' : 'Remove from this mail'}
                              aria-label={isDe ? `${em} aus dieser Mail entfernen` : `Remove ${em} from this mail`}
                              onClick={() => setInviteCustomEmails(effectiveEmails.filter(x => x !== em))}
                            ><X size={12} /></button>
                          </span>
                        ))}
                      </div>
                      {/* v31.10: Umbruch statt Quetschen — ohne ihn schrumpfte das
                          Eingabefeld neben „Hinzufügen" auf dem Handy so weit, dass
                          vom Platzhalter nichts mehr lesbar war. */}
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 10 }}>
                        <input
                          type="text"
                          className="dex-ui-input dex-ui-input--sm"
                          value={inviteAddInput}
                          onChange={e => setInviteAddInput(e.target.value)}
                          placeholder={isDe ? 'Adresse ergänzen, z. B. vorname.nachname@deloitte.de' : 'Add an address, e.g. first.last@deloitte.de'}
                          aria-label={isDe ? 'Adresse ergänzen' : 'Add address'}
                          style={{ flex: '1 1 200px', minWidth: 160 }}
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
                  <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
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
              )}
            </div>

            {/* ---- 2. Wie sieht die Mail aus? --------------------------- */}
            {/* v30.52: dieselbe Auswahl wie in Massen- und QR-Mail. */}
            <div className="dex-ui-section">
              <button type="button" className={cx('dex-ui-disclosure', aufBild && 'is-open')} onClick={() => setAufBild(o => !o)} aria-expanded={aufBild}>
                <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                {isDe ? 'Wie sieht die Mail aus? (Kopfbild)' : 'What does the email look like? (header image)'}
              </button>
              {aufBild && (
              <div className="dex-ui-disclosure-body">
              <MailHeaderImageChooser
                value={inviteHeaderImage}
                onChange={setInviteHeaderImage}
                eventPhotoB64={inviteEventPhotoB64}
                mailLogoB64={mailLogo}
                disabled={inviteSending}
                onCrop={() => setComposerCrop('invite')}
                isDe={isDe}
                customB64={eigenesKopfB64}
                onPickCustom={(f) => { eigenesKopfLaden(f).catch(() => setKopfBusy(false)); }}
                onRemoveCustom={() => { setEigenesKopfB64(''); setKopfNote(''); setInviteHeaderImage(prev => ({ ...prev, hero: 'logo' })); }}
                customBusy={kopfBusy}
                customNote={kopfNote}
              />
              <div className="dex-ui-help" style={{ marginTop: 0 }}>
                {isDe
                  /* v31.10: ohne „rechts" — auf dem Handy steht die Vorschau nicht
                     daneben, sondern hinter dem Reiter „Vorschau". */
                  ? 'Die Vorschau zeigt Kopfbild, Überschrift und Text so, wie die Mail ankommt.'
                  : 'The preview shows header image, heading and text as the email will arrive.'}
              </div>
              </div>
              )}
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
        const previewToLine = !isBroadcast
          ? (nurIch ? myEmail : targetEmails.join(', '))
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
        React.useEffect(() => {
          if (vorgewaehltRef.current || inviteCustomEmails) return;
          if (noAudience || keinVersandVerteiler) { vorgewaehltRef.current = true; setIntern(new Set<'ich' | 'orgs' | 'test'>(['ich'])); return; }
          if (invitedLc === undefined) return; // Einladungen werden noch gelesen
          vorgewaehltRef.current = true;
          setInviteTarget(invitedKnown && uninvitedEmails.length > 0 ? 'uninvited' : 'audience');
          // eslint-disable-next-line react-hooks/exhaustive-deps
        }, [invitedLc, noAudience, keinVersandVerteiler]);
        if (schritt === 'wer') {
          return (
            <Modal
              open={true}
              onClose={() => setShowInviteModal(false)}
              maxWidth={980}
              title={isDe ? 'An wen soll die Einladung gehen?' : 'Who should receive the invitation?'}
              subtitle={isDe ? 'Wähle die Empfänger — den Text schreibst du danach im Editor.' : 'Choose the recipients — you write the text in the editor next.'}
              icon={<Users size={20} />}
              footer={
                <>
                  {onZurueckZurArt && (
                    <button type="button" className="btn btn-outline" style={{ marginRight: 'auto' }} onClick={() => { setShowInviteModal(false); onZurueckZurArt(); }}>
                      {isDe ? 'Zurück: Art der Mail' : 'Back: type of mail'}
                    </button>
                  )}
                  <button type="button" className="btn btn-secondary" onClick={() => setShowInviteModal(false)}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
                  <button type="button" className="btn btn-primary" disabled={targetEmails.length === 0} onClick={() => setSchritt('mail')}>
                    {isDe ? `Weiter: Einladung schreiben (${targetEmails.length})` : `Next: write invitation (${targetEmails.length})`}
                  </button>
                </>
              }
            >
              {zielAuswahl}
            </Modal>
          );
        }
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
            imageBase64={kopfBildVorschau(inviteHeaderImage, { photo: inviteEventPhotoB64, custom: eigenesKopfB64, mailLogo })}
            imageWidth={inviteHeaderImage.width}
            imagePaddingV={inviteHeaderImage.paddingV}
            imagePaddingH={inviteHeaderImage.paddingH}
            // v31.79: Größe steht in der 4. Kachel der Bildwahl (MailHeaderImageChooser).
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

