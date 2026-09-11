/* AdminActionsCard — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 7972-9375 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3: Die Kacheln sind nach dem Ablauf eines Events gruppiert (Event →
 * Kommunikation → Check-in → Teilnehmer → Daten & Export → Wartung), jede
 * Beschreibung beginnt mit der Folge des Klicks, gesperrte Aktionen nennen
 * den Grund (`whenList`), und DE/EN sagen dasselbe. Handler, Bedingungen,
 * Bestätigungstexte und Titel (Handbuch, Tour und Suche verweisen darauf)
 * sind unverändert.
 */
import * as React from 'react';
import { ActionTile, ActionsCollapsibleCard } from '../../admin/ActionsMenu';
import { AlertCircle, Calendar, Check, Columns, Copy, Download, ExternalLink, FileText, Hash, Link2, Mail, Pencil, QrCode, RefreshCw, Send, Shirt, Users, Wrench } from '../../Icons';
import { parseBillingOf } from '../../../utils/faBilling';
import { WAITLIST_BLOCKER_KEY, waitlistBlockerEnabled } from '../../../services/events/waitlistBlocker';
import { buildHashDeepLink } from '../../../utils/deepLink';
import { isB2RunKoelnTitle } from '../../../data/b2runKoeln';
import { EventService, REG_LIST_NAME, SPRegistration } from '../../../services/EventService';
import { DeloitteEvent } from '../../../types';
import { SharePointService } from '../../../services/SharePointService';

export interface AdminActionsCardProps {
  adminEvents: DeloitteEvent[];
  allEvents: DeloitteEvent[];
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  copiedDeepLink: boolean;
  copiedEmails: boolean;
  detectOverbookResult: string;
  eventServiceRef: EventService;
  fixColumnsResult: string;
  fixFieldsResult: string;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  isAdmin: boolean;
  isCheckingDeclines: boolean;
  isDe: boolean;
  isDetectingOverbook: boolean;
  isFixingColumns: boolean;
  isFixingFields: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  isPromoting: boolean;
  isRefreshingProfiles: boolean;
  isReorderingIDs: boolean;
  isRepairingAccess: boolean;
  isRepairingNames: boolean;
  isRepairingOrganizers: boolean;
  isRepairingPerms: boolean;
  isResettingCounter: boolean;
  isSendingQR: boolean;
  isSplitCapacity: boolean;
  isSyncingRegistry: boolean;
  navigate: (page: import("../../../context/NavigationContext").Page, eventId?: string, intent?: import("../../../context/NavigationContext").NavIntent) => void;
  openChangeLogForEvent: () => void;
  openCommsModal: () => void;
  openInviteModal: () => void;
  openMassmailPicker: () => void;
  promoteResult: string;
  qrSentCount: number;
  refreshEvents: () => Promise<void>;
  refreshProfilesResult: string;
  registrations: SPRegistration[];
  reorderResult: string;
  repairAccessResult: string;
  repairNamesResult: string;
  repairOrganizersResult: string;
  repairPermsResult: string;
  resetCounterResult: string;
  runIdReorder: () => Promise<void>;
  runManualPromote: () => Promise<void>;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<{ email: string; displayName: string; location: string; jobTitle: string; }[]>;
  selectedEvent: DeloitteEvent;
  setAccessFixModal: React.Dispatch<React.SetStateAction<{ running: boolean; evIdx: number; evTotal: number; evTitle: string; itemDone: number; itemTotal: number; summary: string[]; }>>;
  /** v31.4: Dialog „Startnummern zuteilen" (s. B2RunAssignBibsModal). */
  setAssignBibsOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setB2runTodoOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setBibImportOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setBillingPanelOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setCheckInHubOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setCheckInHubStep: React.Dispatch<React.SetStateAction<"choose" | "checkin">>;
  setCopiedDeepLink: React.Dispatch<React.SetStateAction<boolean>>;
  setCopiedEmails: React.Dispatch<React.SetStateAction<boolean>>;
  setDeclineCopied: React.Dispatch<React.SetStateAction<boolean>>;
  setDeclineResult: React.Dispatch<React.SetStateAction<{ declinedAndRegistered: { email: string; name: string; reg: SPRegistration; }[]; declinedTotal: number; error: string; }>>;
  setDetectOverbookResult: React.Dispatch<React.SetStateAction<string>>;
  setExcelAudience: React.Dispatch<React.SetStateAction<"active" | "activePlusWait" | "waitOnly" | "withCancelled">>;
  setExcelTargetModal: React.Dispatch<React.SetStateAction<{ mode: "deloitte" | "b2run"; chooseMode?: boolean; }>>;
  setFixColumnsResult: React.Dispatch<React.SetStateAction<string>>;
  setFixFieldsResult: React.Dispatch<React.SetStateAction<string>>;
  setIsCheckingDeclines: React.Dispatch<React.SetStateAction<boolean>>;
  setIsDetectingOverbook: React.Dispatch<React.SetStateAction<boolean>>;
  setIsFixingColumns: React.Dispatch<React.SetStateAction<boolean>>;
  setIsFixingFields: React.Dispatch<React.SetStateAction<boolean>>;
  setIsRefreshingProfiles: React.Dispatch<React.SetStateAction<boolean>>;
  setIsRepairingAccess: React.Dispatch<React.SetStateAction<boolean>>;
  setIsRepairingNames: React.Dispatch<React.SetStateAction<boolean>>;
  setIsRepairingOrganizers: React.Dispatch<React.SetStateAction<boolean>>;
  setIsRepairingPerms: React.Dispatch<React.SetStateAction<boolean>>;
  setIsResettingCounter: React.Dispatch<React.SetStateAction<boolean>>;
  setIsSyncingRegistry: React.Dispatch<React.SetStateAction<boolean>>;
  setNameFixModal: React.Dispatch<React.SetStateAction<{ running: boolean; step: string; evIdx: number; evTotal: number; summary: string[]; }>>;
  setRefreshProfilesResult: React.Dispatch<React.SetStateAction<string>>;
  setRepairAccessResult: React.Dispatch<React.SetStateAction<string>>;
  setRepairNamesResult: React.Dispatch<React.SetStateAction<string>>;
  setRepairOrganizersResult: React.Dispatch<React.SetStateAction<string>>;
  setRepairPermsResult: React.Dispatch<React.SetStateAction<string>>;
  setResetCounterResult: React.Dispatch<React.SetStateAction<string>>;
  setShirtSizeOpen: React.Dispatch<React.SetStateAction<boolean>>;
  /** v31.4: Dialog „QR-Nummern nachtragen“ (gedruckte Nummern aus DEX_Emails). */
  setQrBackfillOpen: React.Dispatch<React.SetStateAction<boolean>>;
  /** v30.93: Kopie in ein neues Event mit Programmpunkten (Programmpunkte, Stufe 4). */
  setCopyToAgendaOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setShowDeclineModal: React.Dispatch<React.SetStateAction<boolean>>;
  setShowExportMenu: React.Dispatch<React.SetStateAction<boolean>>;
  setSubRegReloadTick: React.Dispatch<React.SetStateAction<number>>;
  setSyncRegistryResult: React.Dispatch<React.SetStateAction<string>>;
  shirtFieldExists: boolean;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  showExportMenu: boolean;
  siteUrl: string;
  spServiceRef: SharePointService;
  syncRegistryResult: string;
  t: (key: string) => string;
  updateEvent: (eventId: string, updates: Record<string, unknown>, opts?: { skipReload?: boolean; }) => Promise<boolean>;
}

export const AdminActionsCard: React.FC<AdminActionsCardProps> = (p) => {
  // v31.15: Wartelisten-Platzhalter ein-/ausschalten. Der Wert ist ein
  // Piggyback in EmailTemplateOverrides — `patchEventOverridesValue` liest den
  // aktuellen Stand und schreibt NUR diesen Schluessel zurueck, damit der
  // Schalter nicht die Mail-Vorlagen des Events ueberbuegelt.
  const [blockerBusy, setBlockerBusy] = React.useState(false);

  const { adminEvents, allEvents, childEventsOf, confirmDialog, copiedDeepLink, copiedEmails, detectOverbookResult, eventServiceRef, fixColumnsResult, fixFieldsResult, isAdmin, isCheckingDeclines, isDe, isDetectingOverbook, isFixingColumns, isFixingFields, isOrganizerFor, isPromoting, isRefreshingProfiles, isReorderingIDs, isRepairingAccess, isRepairingNames, isRepairingOrganizers, isRepairingPerms, isResettingCounter, isSendingQR, isSplitCapacity, isSyncingRegistry, navigate, openChangeLogForEvent, openCommsModal, openInviteModal, openMassmailPicker, promoteResult, qrSentCount, refreshEvents, refreshProfilesResult, registrations, reloadRegistrations, reorderResult, repairAccessResult, repairNamesResult, repairOrganizersResult, repairPermsResult, resetCounterResult, runIdReorder, runManualPromote, searchUsers, selectedEvent, setAccessFixModal, setB2runTodoOpen, setBibImportOpen, setBillingPanelOpen, setCheckInHubOpen, setCheckInHubStep, setCopiedDeepLink, setCopiedEmails, setDeclineCopied, setDeclineResult, setDetectOverbookResult, setExcelAudience, setExcelTargetModal, setFixColumnsResult, setFixFieldsResult, setIsCheckingDeclines, setIsDetectingOverbook, setIsFixingColumns, setIsFixingFields, setIsRefreshingProfiles, setIsRepairingAccess, setIsRepairingNames, setIsRepairingOrganizers, setIsRepairingPerms, setIsResettingCounter, setIsSyncingRegistry, setNameFixModal, setRefreshProfilesResult, setRepairAccessResult, setRepairNamesResult, setRepairOrganizersResult, setRepairPermsResult, setResetCounterResult, setQrBackfillOpen, setShirtSizeOpen, setShowDeclineModal, setShowExportMenu, setSubRegReloadTick, setSyncRegistryResult, shirtFieldExists, showAlert, showExportMenu, siteUrl, spServiceRef, syncRegistryResult, t, updateEvent } = p;

  const waitlistBlockerOn = waitlistBlockerEnabled(selectedEvent.emailTemplateOverrides);
  const toggleWaitlistBlocker = async (): Promise<void> => {
    const an = !waitlistBlockerOn;
    const ok = await confirmDialog(
      an
        ? (isDe
          ? 'Wartelisten-Platzhalter einschalten? Wer ab jetzt auf der Warteliste landet, bekommt einen eigenen Kalendereintrag „mit Vorbehalt". Wichtig: Das wirkt nur, wenn der Flow DEX_Outlook_Einladungen die Auftragsarten BlockerSetzen und BlockerLoeschen kennt — sonst entstehen Aufträge, die niemand abarbeitet. Bereits Wartende bekommen rückwirkend keinen Platzhalter.'
          : 'Turn on the waitlist placeholder? Anyone landing on the waitlist from now on gets their own calendar entry marked “tentative”. Important: this only works once the DEX_Outlook_Einladungen flow knows the BlockerSetzen and BlockerLoeschen action types — otherwise jobs pile up unprocessed. People already waiting get no placeholder retroactively.')
        : (isDe
          ? 'Wartelisten-Platzhalter ausschalten? Neue Wartende bekommen keinen Kalendereintrag mehr. Bereits gesetzte Platzhalter bleiben stehen, bis die Person nachrückt oder sich abmeldet — sie werden NICHT rückwirkend entfernt.'
          : 'Turn off the waitlist placeholder? New waitlisted people no longer get a calendar entry. Placeholders already set remain until the person moves up or cancels — they are NOT removed retroactively.'),
      { confirmLabel: an ? (isDe ? 'Einschalten' : 'Turn on') : (isDe ? 'Ausschalten' : 'Turn off') },
    );
    if (!ok) return;
    setBlockerBusy(true);
    try {
      const gespeichert = await eventServiceRef.patchEventOverridesValue(Number(selectedEvent.id), WAITLIST_BLOCKER_KEY, an);
      if (!gespeichert) {
        showAlert(isDe ? 'Die Einstellung konnte nicht gespeichert werden. Bitte versuch es noch einmal.' : 'The setting could not be saved. Please try again.');
        return;
      }
      await refreshEvents();
    } finally {
      setBlockerBusy(false);
    }
  };
  const { setCopyToAgendaOpen, setAssignBibsOpen } = p;
  // v31.3: Gesperrte Aktionen bleiben sichtbar — mit dem Grund in der
  // Folgezeile (Leitfaden 5a, Punkt 5). Fast alle Sperren haben dieselbe
  // Ursache: Das Event hat (noch) keine Teilnehmerliste. Der Grund steht VOR
  // der Beschreibung, weil das Dropdown die Zeile gedämpft rendert und der
  // Leser zuerst wissen will, warum er nicht klicken kann — nicht, was er
  // bekäme, wenn er könnte.
  const hasList = !!selectedEvent?.subsiteUrl;
  const noListReason = isDe
    ? 'Gerade nicht möglich: Dieses Event hat noch keine Teilnehmerliste. '
    : 'Not available right now: this event has no participant list yet. ';
  const whenList = (desc: string): string => (hasList ? desc : noListReason + desc);
  return (
        <ActionsCollapsibleCard isDe={isDe}>
          {/* v31.3: Die Kacheln registrieren sich im ActionsRegistryProvider und
              werden im ActionsDropdown (EventDetailCard) nach `category`
              gruppiert angezeigt — dieser Container bleibt display:none.
              Die Reihenfolge HIER folgt dem Ablauf eines Events (Event →
              Kommunikation → Check-in → Teilnehmer → Daten & Export →
              Wartung), damit ein Leser die Aktionen so findet, wie ein
              Organizer sie braucht. Sichtbar wirken aus dieser Datei nur
              `category`, `title`, `desc`, `disabled` und `badge`. */}
          <div className="admin-actions-grid dex-ui-action-grid">
            {/* v31.3 · Gruppe EVENT — das Event selbst: bearbeiten, teilen, nachvollziehen. */}
            {/* 1. Event bearbeiten */}
            <ActionTile
              icon={<Pencil size={18} />}
              category="event"
              title={t('admin.editbutton') || 'Event bearbeiten'}
              desc={isDe
                ? 'Öffnet das Event im Schritt-für-Schritt-Wizard: Titel, Datum, Ort, Plätze, Anmeldefragen, Mail-Vorlagen und Quiz nachträglich anpassen.'
                : 'Opens the event in the step-by-step wizard: adjust title, date, location, seats, registration questions, mail templates and quiz afterwards.'}
              badge="organizer"
              onClick={() => navigate('edit-event', selectedEvent.id)}
            />

            {/* v31.15: Wartelisten-Platzhalter im Kalender. Sitzt hier und
                nicht im Assistenten, weil er eine Entscheidung ueber den
                LAUFENDEN Betrieb ist — und weil er eine Voraussetzung hat,
                die der Assistent nicht pruefen kann: Der Flow
                DEX_Outlook_Einladungen muss die beiden neuen Auftragsarten
                kennen. Deshalb nennt die Rueckfrage sie ausdruecklich; ohne
                die Flow-Aenderung entstehen Auftraege, die niemand
                abarbeitet. Vorgabe ist AUS. */}
            {!!selectedEvent.waitlistEnabled && !selectedEvent.disableOutlook && (
              <ActionTile
                icon={<Calendar size={18} />}
                category="event"
                title={waitlistBlockerOn
                  ? (isDe ? 'Wartelisten-Platzhalter: an' : 'Waitlist placeholder: on')
                  : (isDe ? 'Wartelisten-Platzhalter: aus' : 'Waitlist placeholder: off')}
                desc={isDe
                  ? 'Wer auf der Warteliste landet, bekommt einen eigenen Kalendereintrag „mit Vorbehalt", der den Termin freihält — kein Platz im echten Outlook-Termin, keine Zusage. Rückt die Person nach oder meldet sie sich ab, verschwindet der Platzhalter wieder. Setzt voraus, dass der Flow DEX_Outlook_Einladungen die Auftragsarten BlockerSetzen und BlockerLoeschen kennt.'
                  : 'Anyone landing on the waitlist gets their own calendar entry marked “tentative” that holds the slot — no seat in the real Outlook invite, no commitment. It disappears when the person moves up or cancels. Requires the DEX_Outlook_Einladungen flow to know the BlockerSetzen and BlockerLoeschen action types.'}
                badge="organizer"
                busy={blockerBusy}
                onClick={() => { void toggleWaitlistBlocker(); }}
              />
            )}

            {/* v11.89/v20.3: Der Event-Live/Entwurf-Toggle ist aus dem
                Aktionen-Menü ausgezogen — der Status-Badge neben dem
                Event-Titel ist jetzt selbst der klickbare Umschalter. */}

            {/* v10.19: Deep-Link kopieren — Organizer/Admin können den Link
                des aktuell offenen Events in die Zwischenablage legen und z.B.
                an Co-Organizer / Helfer weitergeben. Zielseite ist exakt
                dieses Admin-Center-Detail (?action=admin&event=<SP-ID>). Beim
                Aufruf landet der Empfänger nach Login direkt auf der gleichen
                Detail-Ansicht statt in der Event-Auswahl-Liste. */}
            <ActionTile
              icon={<Link2 size={18} />}
              category="event"
              title={copiedDeepLink ? (t('admin.copied') || 'Kopiert') : (isDe ? 'Deep-Link kopieren' : 'Copy deep link')}
              desc={isDe
                ? 'Legt den direkten Link auf diese Event-Ansicht in die Zwischenablage. Per Mail oder Teams an Co-Organizer schicken — sie landen nach dem Login direkt hier, ohne die Event-Liste durchzuklicken.'
                : 'Copies the direct link to this event view to the clipboard. Send it via email or Teams to co-organizers — after login they land right here without clicking through the event list.'}
              badge="organizer"
              onClick={() => {
                const base = (typeof window !== 'undefined' && window.location)
                  ? `${window.location.origin}${window.location.pathname}`
                  : `${siteUrl}/SitePages/DEX.aspx`;
                // v26.33: Deep-Link-Parameter im Hash (Outlook/Teams-resistent).
                const url = buildHashDeepLink(`${base}?env=WebView`, { action: 'admin', event: selectedEvent.id });
                if (typeof navigator !== 'undefined' && navigator.clipboard) {
                  navigator.clipboard.writeText(url).then(() => {
                    setCopiedDeepLink(true);
                    setTimeout(() => setCopiedDeepLink(false), 2000);
                  }).catch(() => { showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{url}</span>, { title: isDe ? 'Deep-Link manuell kopieren' : 'Copy deep link manually' }); });
                } else {
                  showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{url}</span>, { title: isDe ? 'Deep-Link manuell kopieren' : 'Copy deep link manually' });
                }
              }}
            />

            {/* v19.30 (Feature D): Audit-Log / Änderungsprotokoll dieses
                Events öffnen — vorgefiltert auf den Event-Titel. Sichtbar für
                Admin oder Organizer dieses Events. Zeigt pro Eintrag Zeitpunkt,
                Akteur, Aktion, Ziel-Teilnehmer und bei Daten-Änderungen das
                Vorher → Nachher je Feld. */}
            {(isAdmin || isOrganizerFor(selectedEvent)) && (
              <ActionTile
                icon={<FileText size={18} />}
                category="event"
                title={isDe ? 'Audit-Log / Änderungsprotokoll' : 'Audit log / change history'}
                desc={isDe
                  ? 'Öffnet das Änderungsprotokoll vorgefiltert auf dieses Event. Du siehst pro Eintrag: wann, wer, welche Aktion (z.B. bearbeitet, abgemeldet, gelöscht), welcher Teilnehmer betroffen war und bei Daten-Änderungen den genauen Vorher → Nachher-Vergleich je Feld.'
                  : 'Opens the change history pre-filtered to this event. Each entry shows: when, who, which action (e.g. edited, deregistered, deleted), which participant was affected and — for data changes — the exact before → after comparison per field.'}
                badge="organizer"
                onClick={openChangeLogForEvent}
              />
            )}

            {/* 5c-3. v30.93: Programmpunkte, Stufe 4 — ein Event mit Sub-Events
                in ein NEUES Event mit Programmpunkten überführen. Nur bei
                Klammern mit Terminen; das alte Event bleibt unangetastet. */}
            {selectedEvent && !selectedEvent.parentEventId && childEventsOf(selectedEvent.id).length > 0 && (
              <ActionTile
                icon={<Copy size={18} />}
                category="event"
                title={isDe ? 'Als neues Event mit Programmpunkten kopieren' : 'Copy as new event with agenda items'}
                desc={isDe
                  ? 'Legt ein NEUES Event an, dessen Programmpunkte die bisherigen Sub-Events sind: Anmeldungen werden still kopiert (keine Mail, kein Outlook), Check-ins werden zu Anwesenheiten. Dieses Event bleibt vollständig erhalten.'
                  : 'Creates a NEW event whose agenda items are the current sub-events: registrations are copied silently (no mail, no Outlook), check-ins become attendances. This event stays intact.'}
                badge="organizer"
                onClick={() => setCopyToAgendaOpen(true)}
              />
            )}

            {/* v30.5: Event-Abrechnung (Fachkonzept Abschnitt 6) — erscheint
                AUSSCHLIESSLICH bei abrechnungsrelevanten Events. Versand an
                F&A + Versandhistorie liegen im Modal (BillingActionPanel). */}
            {parseBillingOf(selectedEvent)?.relevant === true && (
              <ActionTile
                icon={<Send size={18} />}
                category="event"
                title="Event-Abrechnung"
                desc={isDe
                  ? 'Sendet Abrechnungsinformationen oder die Teilnehmerliste an Finance & Accounting und zeigt die Versandhistorie. Erscheint nur bei abrechnungsrelevanten Events.'
                  : 'Sends billing information or the participant list to Finance & Accounting and shows the send history. Appears only for billing-relevant events.'}
                badge="organizer"
                onClick={() => setBillingPanelOpen(true)}
              />
            )}

            {/* v31.3 · Gruppe KOMMUNIKATION — Einladen, Rundmails, Adressen. */}
            {/* v11.40: 4b. Einladungsmail — Mail mit Anmelde-Link an dich
                (zum Weiterleiten an Kollegen / Teams / externe Adressen)
                oder direkt an den auf dem Event hinterlegten Mailverteiler.
                Default-Text + Link werden vorbefüllt, sind aber im RichText-
                Editor frei editierbar. */}
            <ActionTile
              icon={<Send size={18} />}
              category="mails"
              title={isDe ? 'Einladungsmail' : 'Invitation email'}
              desc={isDe
                ? 'Versendet eine Einladungs-Mail mit Anmelde-Link — an dich zum Weiterleiten oder direkt an den hinterlegten Mailverteiler des Events.'
                : 'Sends an invitation email with the registration link — to yourself for forwarding or directly to the configured mail distribution list of the event.'}
              badge="organizer"
              onClick={openInviteModal}
            />

            {/* 4. Massenmail an alle aktiven Teilnehmer */}
            <ActionTile
              icon={<Mail size={18} />}
              category="mails"
              title={isDe ? 'E-Mail versenden' : 'Send email'}
              desc={isDe
                ? 'Öffnet den Editor mit der Deloitte-Mail-Vorlage. Die Mail geht an alle aktiven Teilnehmer — nicht an Warteliste oder Abgemeldete. Hier legst du auch eine Umfrage oder Abstimmung an: Antworten zum Anklicken direkt in der Mail.'
                : 'Opens the editor with the Deloitte mail template. The mail goes to all active participants — not to the waitlist or cancelled people. This is also where you add a poll or survey: answers to click right inside the email.'}
              badge="organizer"
              onClick={openMassmailPicker}
            />

            {/* 4c. Gesendete Rundmails — durabler Kommunikations-Log
                (DEX_EventComms). Zeigt alle versendeten Broadcast-Mails
                (Einladung / Massenmail) mit Zeitstempel + Absender; Klick
                auf eine Zeile blendet den kompletten HTML-Body ein. */}
            <ActionTile
              icon={<Mail size={18} />}
              category="mails"
              title={isDe ? 'Gesendete Mails' : 'Sent emails'}
              desc={isDe
                ? 'Zeigt alle Rundmails zu diesem Event (Einladung, Massenmail) mit Zeitpunkt und Absender. Ein Klick auf eine Zeile öffnet den kompletten Mail-Text.'
                : 'Shows all broadcast emails for this event (invitation, mass mail) with time and sender. Clicking a row opens the full mail body.'}
              badge="organizer"
              onClick={openCommsModal}
            />

            {/* 3. E-Mail-Adressen kopieren */}
            <ActionTile
              icon={<Copy size={18} />}
              category="mails"
              title={copiedEmails ? (t('admin.copied') || 'Kopiert') : (t('admin.copyemails') || 'E-Mails kopieren')}
              desc={isDe
                ? 'Legt die Adressen aller aktiven Teilnehmer (mit Semikolon getrennt) in die Zwischenablage — direkt in Outlook-Empfänger oder andere Werkzeuge einfügbar. Warteliste und Abgemeldete sind nicht dabei.'
                : 'Copies the addresses of all active participants (semicolon-separated) to the clipboard — paste them straight into Outlook recipients or other tools. Waitlist and cancelled people are not included.'}
              badge="organizer"
              onClick={() => {
                const emails = registrations
                  .filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt')
                  .map(r => r.ParticipantEmail)
                  .join('; ');
                if (emails) {
                  navigator.clipboard.writeText(emails).then(() => {
                    setCopiedEmails(true);
                    setTimeout(() => setCopiedEmails(false), 2000);
                  }).catch(() => { showAlert(<span style={{ userSelect: 'all', wordBreak: 'break-all', fontFamily: 'monospace', fontSize: '0.8rem' }}>{emails}</span>, { title: isDe ? 'E-Mail-Adressen manuell kopieren' : 'Copy email addresses manually' }); });
                }
              }}
            />

            {/* v31.3 · Gruppe CHECK-IN — der Event-Tag. */}
            {/* v30.36: Ein Einstieg statt fuenf Kacheln. „QR-Codes versenden"
                und die drei Self-Check-in-Kacheln standen gleichrangig
                nebeneinander und haben das Aktionen-Grid dominiert, obwohl sie
                zusammengehoeren und meist nur EINE davon gebraucht wird. Jetzt
                ein Knopf, dahinter eine Entscheidung: Codes verschicken oder
                Check-in am Event-Tag. Die Self-Check-in-Varianten (PDF,
                Live-Anzeige) erscheinen erst, wenn man sich fuer Check-in
                entschieden hat — vorher sind sie nur Rauschen. */}
            <ActionTile
              icon={<QrCode size={18} />}
              category="checkin"
              title={isSendingQR ? (isDe ? `QR-Codes werden versendet... (${qrSentCount})` : `Sending QR codes... (${qrSentCount})`) : (isDe ? 'QR-Codes und Check-In' : 'QR codes and check-in')}
              desc={isDe
                ? 'Alles rund um den Event-Tag an einer Stelle: persönliche QR-Codes an die Teilnehmer verschicken — oder das Check-in vorbereiten und starten (Team scannt, oder Teilnehmer checken sich selbst ein).'
                : 'Everything about event day in one place: send personal QR codes to attendees — or prepare and start check-in (your team scans, or attendees check themselves in).'}
              badge="organizer"
              busy={isSendingQR}
              onClick={() => { setCheckInHubStep('choose'); setCheckInHubOpen(true); }}
            />

            {/* v9.20: Check-In starten — prominent als erster Tile.
                Sowohl Organizer als auch Check-In-Team-Mitglieder dürfen
                diese Aktion auslösen (siehe Header.canCheckIn-Logik). */}
            <ActionTile
              icon={<Hash size={18} />}
              category="checkin"
              title={t('admin.checkin')}
              desc={isDe
                ? 'Öffnet das Check-in-Tool: QR-Codes scannen, manuell ein- und auschecken, Live-Zahlen sehen (angemeldet / eingecheckt / ausstehend). Am Eventtag das wichtigste Werkzeug.'
                : 'Opens the check-in tool: scan QR codes, check in and out manually, see live numbers (registered / checked in / pending). The most important tool on event day.'}
              badge="organizer"
              onClick={() => navigate('check-in', selectedEvent.id)}
            />

            {/* v31.3 · Gruppe TEILNEHMER — Plätze, Nummern, Prüfungen. */}
            {/* 7a2. Nachrücken — Admin ODER Organizer des Events (v18.70).
                v29.16: Füllt ALLE freien Plätze, bei geteilten Kapazitäten je
                Gruppe getrennt gerechnet. Die Rückfrage mit der Aufstellung
                steht in runManualPromote — dort sind die Zahlen bekannt; eine
                zweite Rückfrage davor wäre nur eine Frage ohne Inhalt. */}
            {(isAdmin || (!!selectedEvent && isOrganizerFor(selectedEvent))) && (
              <ActionTile
                icon={<Users size={18} />}
                category="participants"
                title={isPromoting ? (isDe ? 'Rückt nach…' : 'Promoting…') : (isDe ? 'Freie Plätze mit Warteliste füllen' : 'Fill free seats from waitlist')}
                desc={(isReorderingIDs && hasList
                  ? (isDe ? 'Gerade nicht möglich: Die IDs werden gerade neu vergeben — danach wieder verfügbar. ' : 'Not available right now: IDs are being reassigned — available again afterwards. ')
                  : '') + whenList(isDe
                  ? 'Rückt so viele Personen von der Warteliste nach, wie Plätze frei sind — in der Reihenfolge der TeilnehmerIDs. Bei zwei Gruppen wird je Gruppe getrennt gerechnet, es rückt also niemand in eine noch volle Gruppe. Vor dem Ausführen siehst du die Aufstellung. Jede nachgerückte Person bekommt Status „Angemeldet", eine Nachrück-Mail und eine Outlook-Einladung; danach werden die IDs neu vergeben. Nötig vor allem, wenn du eine Kapazität erhöht hast — von allein rückt nur bei einer Abmeldung jemand nach.'
                  : 'Moves as many people up from the waitlist as there are free seats, in participant-ID order. With two groups each group is calculated separately, so nobody moves into a group that is still full. You see the breakdown before it runs. Everyone promoted gets status “Registered”, a promotion email and an Outlook invite; IDs are reassigned afterwards. Mainly needed after you raised a capacity — on its own, promotion only happens on a cancellation.')}
                badge="organizer"
                busy={isPromoting}
                disabled={!selectedEvent?.subsiteUrl || isPromoting || isReorderingIDs}
                result={promoteResult}
                resultIsError={!!promoteResult && (promoteResult.indexOf('Fehler') >= 0 || promoteResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  await runManualPromote();
                }}
              />
            )}

            {/* 7c. Überbuchung prüfen — Admin ODER Organizer des Events (v11.36).
                Markiert pro Gruppe (bzw. gesamt) die zuletzt über Kapazität
                Angemeldeten mit OverbookReview='Pending'. Ändert KEINEN
                Status — Admin/Organizer entscheidet danach pro Person über
                die Buttons in der „Überbuchung – zu prüfen"-Box oben in der
                Teilnehmerliste. Organizer dürfen das für eigene Events, weil
                es Teilnehmerverwaltung ist (analog Abmelden/QR/Massenmail). */}
            {(isAdmin || (!!selectedEvent && isOrganizerFor(selectedEvent))) && (
              <ActionTile
                icon={<Users size={18} />}
                category="participants"
                title={isDetectingOverbook ? (isDe ? 'Wird geprüft…' : 'Checking…') : (isDe ? 'Überbuchung prüfen' : 'Check overbooking')}
                desc={whenList(isDe
                  ? 'Findet pro Gruppe (Durchstarter/Funstarter, bzw. gesamt) die zuletzt angemeldeten Personen ÜBER der Kapazität und markiert sie zur Prüfung. Es wird nichts automatisch geändert — danach entscheidest du pro Person (auf Warteliste / Platz behalten) über die Knöpfe oben in der Teilnehmerliste.'
                  : 'Finds, per group (Durchstarter/Funstarter, or overall), the most recently registered people OVER capacity and marks them for review. Nothing is changed automatically — afterwards you decide per person (move to waitlist / keep seat) via the buttons at the top of the participant list.')}
                badge="organizer"
                busy={isDetectingOverbook}
                disabled={!selectedEvent?.subsiteUrl}
                result={detectOverbookResult}
                resultIsError={!!detectOverbookResult && (detectOverbookResult.indexOf('Fehler') >= 0 || detectOverbookResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!(await confirmDialog(isDe ? 'Überbuchung prüfen und betroffene Personen markieren? (ändert keinen Status)' : 'Check overbooking and mark affected people? (does not change any status)', { confirmLabel: isDe ? 'Prüfen' : 'Check' }))) return;
                  setIsDetectingOverbook(true);
                  setDetectOverbookResult(null);
                  try {
                    const res = await eventServiceRef.detectOverbooking(selectedEvent.subsiteUrl, {
                      isSplit: isSplitCapacity,
                      maxParticipants: selectedEvent.maxParticipants || 0,
                      durchstarterCapacity: selectedEvent.durchstarterCapacity || 0,
                      funstarterCapacity: selectedEvent.funstarterCapacity || 0,
                    });
                    // Counter mit echtem Bestand abgleichen (best-effort).
                    try { await eventServiceRef.syncSeatsToActiveCount(selectedEvent.subsiteUrl, { isSplit: isSplitCapacity }); } catch { /* */ }
                    const parts = res.groups
                      .map(g => `${g.group}: ${g.activeBefore}/${g.cap || '∞'} → ${g.marked} ${isDe ? 'markiert' : 'marked'}`)
                      .join(' · ');
                    setDetectOverbookResult(res.total > 0
                      ? (isDe
                        ? `${res.total} markiert (${parts})${res.errors ? ` — ${res.errors} Fehler` : ''}`
                        : `${res.total} marked (${parts})${res.errors ? ` — ${res.errors} errors` : ''}`)
                      : (isDe ? `Keine Überbuchung gefunden (${parts})` : `No overbooking found (${parts})`));
                    await reloadRegistrations();
                  } catch {
                    setDetectOverbookResult(isDe ? 'Fehler beim Prüfen der Überbuchung' : 'Error checking overbooking');
                  }
                  setIsDetectingOverbook(false);
                }}
              />
            )}

            {/* 6. Outlook-Absagen prüfen — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<AlertCircle size={18} />}
                category="participants"
                title={isCheckingDeclines ? (isDe ? 'Outlook wird geprüft…' : 'Checking Outlook…') : (isDe ? 'Outlook-Absagen prüfen' : 'Check Outlook declines')}
                desc={isDe
                  ? 'Liest die Absagen des Outlook-Termins aus dem Postfach no_reply.events und gleicht sie mit den aktiven Teilnehmern ab: Du siehst, wer abgesagt hat, aber noch angemeldet ist — und kannst diese Personen gezielt ansprechen oder abmelden.'
                  : 'Reads the declines of the Outlook appointment from the no_reply.events mailbox and matches them against active participants: you see who declined but is still registered — and can contact or deregister those people directly.'}
                badge="admin"
                busy={isCheckingDeclines}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent) return;
                  setIsCheckingDeclines(true);
                  setDeclineResult(null);
                  setDeclineCopied(false);
                  try {
                    const result = await eventServiceRef.getDeclinedAttendees(selectedEvent.id);
                    if (result.ok) {
                      const activeByEmail = new Map<string, SPRegistration>();
                      for (const r of registrations) {
                        if (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt') {
                          activeByEmail.set(String(r.ParticipantEmail || '').toLowerCase(), r);
                        }
                      }
                      const hits: Array<{ email: string; name: string; reg: SPRegistration }> = [];
                      for (const d of result.attendees) {
                        const reg = activeByEmail.get(d.email);
                        if (reg) hits.push({ email: d.email, name: d.name, reg });
                      }
                      setDeclineResult({
                        declinedAndRegistered: hits,
                        declinedTotal: result.attendees.length,
                        error: null,
                      });
                      setShowDeclineModal(true);
                    } else {
                      let msg = result.message || 'Unbekannter Fehler beim Lesen des Outlook-Termins.';
                      if (!result.message) {
                        if (result.reason === 'no-pointer') {
                          msg = 'Für dieses Event ist kein Outlook-Termin verknüpft (OutlookEventId / CalendarLink fehlen).';
                        } else if (result.reason === 'not-found') {
                          msg = 'Outlook-Termin wurde im Postfach no_reply.events@deloitte.de nicht gefunden.';
                        } else if (result.reason === 'forbidden') {
                          msg = 'Graph-API-Zugriff abgelehnt (HTTP 403). Tenant-Admin muss "Calendars.Read.Shared" genehmigen, und der User braucht Reviewer-Rechte auf dem Postfach-Kalender.';
                        }
                      }
                      setDeclineResult({ declinedAndRegistered: [], declinedTotal: 0, error: msg });
                      setShowDeclineModal(true);
                    }
                  } catch (err) {
                    setDeclineResult({
                      declinedAndRegistered: [],
                      declinedTotal: 0,
                      error: err instanceof Error ? err.message : String(err),
                    });
                    setShowDeclineModal(true);
                  }
                  setIsCheckingDeclines(false);
                }}
              />
            )}

            {/* 7. TeilnehmerIDs neu vergeben — Admin ODER Organizer des Events (v11.36) */}
            {(isAdmin || (!!selectedEvent && isOrganizerFor(selectedEvent))) && (
              <ActionTile
                icon={<Hash size={18} />}
                category="participants"
                title={isReorderingIDs ? (isDe ? 'IDs werden vergeben…' : 'Assigning IDs…') : (isDe ? 'IDs neu vergeben' : 'Reassign IDs')}
                desc={whenList(isDe
                  ? 'Vergibt die TeilnehmerIDs fortlaufend (1, 2, 3, …) nach Erstellungsreihenfolge: schließt Lücken nach Abmeldungen und sortiert die Liste sauber durch. Nicht ausführen, während gerade viele Anmeldungen laufen — erst wenn die Anmeldewelle vorbei ist.'
                  : 'Assigns the participant IDs consecutively (1, 2, 3, …) by creation order: closes gaps after cancellations and sorts the list cleanly. Do not run while many registrations are coming in — wait until the registration wave is over.')}
                badge="organizer"
                busy={isReorderingIDs}
                disabled={!selectedEvent?.subsiteUrl}
                result={reorderResult}
                resultIsError={!!reorderResult && (reorderResult.indexOf('Fehler') >= 0 || reorderResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!(await confirmDialog(isDe
                    ? 'TeilnehmerIDs neu vergeben (1, 2, 3, …)? Sortierung nach Erstellungsreihenfolge.\n\nNICHT ausführen, während gerade viele Anmeldungen laufen — bitte erst wenn die Anmeldewelle vorbei ist.'
                    : 'Reassign participant IDs (1, 2, 3, …)? Sorted by creation order.\n\nDo NOT run while many registrations are coming in — please wait until the registration wave is over.'))) return;
                  await runIdReorder();
                }}
              />
            )}

            {/* v31.3 · Gruppe DATEN & EXPORT — Listen für andere Werkzeuge und Veranstalter. */}
            {/* 5. Excel-Download (mit Dropdown Deloitte/B2Run-View)
                Wrapper braucht display:flex, damit der innere Button auf die
                volle Grid-Zellen-Höhe gestreckt wird — sonst sieht die Kachel
                niedriger aus als ihre Nachbarn, die zwei Zeilen Titel haben. */}
            <div style={{ position: 'relative', display: 'flex' }}>
              <ActionTile
                icon={<Download size={18} />}
                category="participants"
                title={isDe ? 'Excel-Export' : 'Excel export'}
                desc={selectedEvent && (selectedEvent.type === 'B2Run' || isB2RunKoelnTitle(selectedEvent.title))
                  ? (isDe
                    ? 'Lädt die Teilnehmerliste als Excel. Du wählst vorher die Zielgruppe und das Format: „Deloitte Felder“ (alle internen Spalten + Anmeldefragen) oder „B2Run View“ (direkt in b2run.com importierbar).'
                    : 'Downloads the participant list as Excel. You pick the audience and the format first: “Deloitte fields” (all internal columns + registration questions) or “B2Run view” (importable straight into b2run.com).')
                  : (isDe
                    ? 'Lädt die Teilnehmerliste als Excel mit allen internen Spalten und den Anmeldefragen des Events. Die Zielgruppe (z.B. nur aktive, mit Warteliste, mit Abgemeldeten) wählst du vorher.'
                    : 'Downloads the participant list as Excel with all internal columns and the event’s registration questions. You pick the audience (e.g. active only, with waitlist, with cancelled) first.')}
                badge="organizer"
                onClick={() => {
                  // v17.12: Erst Zielgruppe abfragen, dann erst exportieren.
                  // v27.9: Bei B2Run die Format-Auswahl (Deloitte/B2Run) DIREKT
                  // im Modal treffen — der frühere Anker-Dropdown wurde vom
                  // „Aktion auswählen"-Menü (overflow:auto) abgeschnitten, daher
                  // kam die Auswahl gar nicht erst zum Vorschein.
                  setExcelAudience('active');
                  if (selectedEvent && (selectedEvent.type === 'B2Run' || isB2RunKoelnTitle(selectedEvent.title))) {
                    setExcelTargetModal({ mode: 'b2run', chooseMode: true });
                  } else {
                    setExcelTargetModal({ mode: 'deloitte' });
                  }
                }}
              />
              {/* v31.3: Das Format-Untermenü als Zeilen-Knöpfe mit Klassen-Hover
                  statt onMouseEnter-Inline-Styles (Leitfaden 1.3). Seit v27.9
                  wird die Format-Wahl im Excel-Modal getroffen; `showExportMenu`
                  bleibt als Prop bestehen, damit der Aufrufer unverändert läuft. */}
              {showExportMenu && (
                <div className="dex-ui-card dex-ui-card--list" style={{ position: 'absolute', top: '100%', left: 0, right: 0, marginTop: 4, zIndex: 100, boxShadow: '0 8px 24px rgba(0,0,0,0.15)' }}>
                  <button
                    type="button"
                    className="dex-ui-rowbtn dex-ui-row dex-ui-row--bordered"
                    onClick={() => { setShowExportMenu(false); setExcelAudience('active'); setExcelTargetModal({ mode: 'deloitte' }); }}
                  >
                    <span className="dex-ui-row-main">
                      <span className="dex-ui-row-title">{isDe ? 'Deloitte Felder' : 'Deloitte fields'}</span>
                      <span className="dex-ui-row-sub" style={{ whiteSpace: 'normal' }}>
                        {isDe
                          ? 'Alle internen Felder: Name, E-Mail, Bereich, Standort, Position, Status, Anmeldedatum + alle Anmeldefragen des Events.'
                          : 'All internal fields: name, email, department, location, position, status, registration date + all registration questions of the event.'}
                      </span>
                    </span>
                  </button>
                  {selectedEvent && (selectedEvent.type === 'B2Run' || isB2RunKoelnTitle(selectedEvent.title)) && (
                    <button
                      type="button"
                      className="dex-ui-rowbtn dex-ui-row dex-ui-row--bordered"
                      onClick={() => { setShowExportMenu(false); setExcelAudience('active'); setExcelTargetModal({ mode: 'b2run' }); }}
                    >
                      <span className="dex-ui-row-main">
                        <span className="dex-ui-row-title">B2Run View</span>
                        <span className="dex-ui-row-sub" style={{ whiteSpace: 'normal' }}>
                          {isDe
                            ? 'Spaltenformat exakt wie die B2Run-Excel-Vorlage (Nr, Anrede, Name, E-Mail, Startblock, AGB, Gruppe, Mobilnummer, Altersklasse, …) — direkt in b2run.com importierbar.'
                            : 'Column format exactly like the B2Run Excel template (no., salutation, name, email, start block, terms, group, mobile number, age class, …) — importable straight into b2run.com.'}
                        </span>
                      </span>
                    </button>
                  )}
                </div>
              )}
            </div>

            {/* 5c-2. v30.60: Bestellliste der Trikots. Bewusst NICHT an B2Run
                gehängt: Ein Feld mit Konfektionsgröße gibt es auch bei
                Team-Events mit Hoodie oder Poloshirt. Die Kachel erscheint,
                sobald es ein passendes Abfragefeld gibt — sonst wäre sie eine
                Aktion, die nur eine Fehlermeldung zeigt. */}
            {selectedEvent && shirtFieldExists && (
              <ActionTile
                icon={<Shirt size={18} />}
                category="participants"
                title={isDe ? 'Benötigte T-Shirts' : 'Required T-shirts'}
                desc={isDe
                  ? 'Zählt die angegebenen Trikot-/Konfektionsgrößen aller angemeldeten Personen zusammen — die Bestellliste je Größe, mit Namen und als Excel. Wer keine Größe angegeben hat, wird namentlich ausgewiesen statt weggelassen.'
                  : 'Totals the shirt sizes of all registered people — the order list per size, with names and as Excel. People without a size are listed by name instead of dropped.'}
                badge="organizer"
                onClick={() => setShirtSizeOpen(true)}
              />
            )}

            {/* 5b. v30.48: Startnummern-Rücklauf einlesen — nur B2Run Köln.
                Bewusst NICHT an `type === 'B2Run'` gehängt: Das Spaltenformat
                der Rücklauf-Datei ist das des Köln-Exports. */}
            {selectedEvent && isB2RunKoelnTitle(selectedEvent.title) && (
              <ActionTile
                icon={<Hash size={18} />}
                category="participants"
                title={isDe ? 'Startnummern importieren' : 'Import bib numbers'}
                desc={isDe
                  ? 'Liest die Rücklauf-Datei des Veranstalters ein und schreibt die echten Startnummern zu den Teilnehmern. Zeigt vorher, welche Nummer wegen einer Abmeldung an einen Nachrücker geht — die musst du beim Veranstalter ummelden.'
                  : 'Reads the organiser\'s return file and writes the real bib numbers to the participants. Shows beforehand which number moves to a waitlist promotion after a cancellation.'}
                badge="organizer"
                onClick={() => setBibImportOpen(true)}
              />
            )}

            {/* 5c. v30.54: Was beim Veranstalter noch zu tun ist. Eigene
                Aktion, weil die Liste NACH dem Import weiterlebt: Jede spätere
                Abmeldung erzeugt eine Ummeldung, und die entsteht, wenn kein
                Import-Fenster offen ist. */}
            {selectedEvent && isB2RunKoelnTitle(selectedEvent.title) && (
              <ActionTile
                icon={<Check size={18} />}
                category="participants"
                title={isDe ? 'Offen beim Veranstalter' : 'Open with the organiser'}
                desc={isDe
                  ? 'Zeigt, welche Startnummern du beim Veranstalter noch ummelden, abmelden oder nachmelden musst — jeweils mit Namen und Nummer. Wird bei jedem Öffnen neu berechnet, spätere Abmeldungen tauchen also von selbst auf. Erledigtes lässt sich abhaken.'
                  : 'Shows which bib numbers still need to be transferred, cancelled or newly registered with the organiser. Recalculated each time you open it, so later cancellations show up by themselves.'}
                badge="organizer"
                onClick={() => setB2runTodoOpen(true)}
              />
            )}

            {/* 5d. v31.4: Freie Startnummern an Personen ohne Nummer.
                Frei ist ausschließlich, was DEX gesehen hat — Nummern auf
                abgemeldeten Zeilen und Nummern aus festgehaltenen
                „abmelden"-Aufgaben. Es wird kein Nummernkreis gespeichert und
                keiner aus Lücken abgeleitet; bereits vergebene Nummern
                bleiben unangetastet (s. B2RunAssignBibsModal). Dieselbe
                Gate-Bedingung wie die beiden Nachbarn. */}
            {selectedEvent && isB2RunKoelnTitle(selectedEvent.title) && (
              <ActionTile
                icon={<Hash size={18} />}
                category="participants"
                title={isDe ? 'Startnummern zuteilen' : 'Assign bib numbers'}
                desc={isDe
                  ? 'Zeigt dir, welche freien Startnummern es gibt und wer noch keine hat — vergeben wird erst, wenn du bestätigst. Frei ist nur, was durch eine Abmeldung frei geworden ist; bereits vergebene Nummern werden nicht angefasst.'
                  : 'Shows which bib numbers are free and who still has none — nothing is assigned until you confirm. Free means freed up by a cancellation; numbers already handed out are never touched.'}
                badge="organizer"
                onClick={() => setAssignBibsOpen(true)}
              />
            )}

            {/* v31.3 · Gruppe WARTUNG (dieses Event) — Reparaturen an Rechten, Spalten, Feldern. */}
            {/* v24.97: „Default-Mail-Vorlagen zurücksetzen" + „Wochenbericht
                jetzt senden" sind GLOBALE Admin-Aktionen und liegen jetzt im
                Admin-Hub (AdminHubPage, Sektion „E-Mails & Berichte") statt hier
                im per-Event-Aktionsmenü. */}

            {/* v30.37: Organizer-Berechtigungen über Klammer UND alle Termine
                neu setzen. Bis v30.36 lief der Sync beim Speichern nur über
                die Klammer-Subsite — nachträglich benannte (Co-)Organizer
                hatten auf keinem einzigen Sub-Event Leserecht und sahen das
                Event als leer. Idempotent: wer die Rechte hat, behält sie.
                Ausführen kann das nur, wer selbst Full Control hat (Admin
                oder Haupt-Organizer) — die betroffene Person kann sich die
                Rechte naturgemäß nicht selbst geben. */}
            {(isAdmin || isOrganizerFor(selectedEvent)) && (
              <ActionTile
                icon={<RefreshCw size={18} />}
                category="maintenance"
                title={isDe ? 'Organizer-Berechtigungen reparieren' : 'Repair organizer permissions'}
                desc={whenList(isDe
                  ? 'Setzt für alle Organizer und Co-Organizer dieses Events das Leserecht auf der Teilnehmerliste — auf dem Haupt-Event UND auf jedem einzelnen Termin. Nötig, wenn jemand nachträglich als Organizer dazugekommen ist und überall „0 Teilnehmer" sieht, obwohl Anmeldungen vorliegen.'
                  : 'Grants every organizer and co-organizer of this event read access to the participant list — on the main event AND on every single date. Needed when someone was added as organizer later and sees “0 participants” everywhere although registrations exist.')}
                badge="organizer"
                busy={isRepairingPerms}
                disabled={!selectedEvent?.subsiteUrl}
                result={repairPermsResult}
                resultIsError={!!repairPermsResult && (repairPermsResult.indexOf('Fehler') >= 0 || repairPermsResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  const emails = (selectedEvent.organizerEmails || [])
                    .concat(selectedEvent.coOrganizerEmails || [])
                    .map(e => (e || '').trim()).filter(Boolean);
                  if (emails.length === 0) {
                    setRepairPermsResult(isDe ? 'Keine Organizer-Adressen hinterlegt' : 'No organizer addresses on file');
                    return;
                  }
                  const sites = [selectedEvent.subsiteUrl]
                    .concat(childEventsOf(selectedEvent.id).map(k => k.subsiteUrl || ''))
                    .filter(Boolean);
                  setIsRepairingPerms(true);
                  setRepairPermsResult(null);
                  try {
                    const r = await eventServiceRef.ensureOrganizerPermissionsMulti(sites, emails.join(';'));
                    const unresolved = r.unresolved.length
                      ? (isDe ? ` · ${r.unresolved.length} Adresse(n) nicht gefunden: ${r.unresolved.join(', ')}` : ` · ${r.unresolved.length} address(es) not found: ${r.unresolved.join(', ')}`)
                      : '';
                    // v30.67 (Review): `failed` auswerten. Das Ergebnis trägt seit
                    // v30.67 je fehlgeschlagene Zuweisung Subsite, Scope (web/list)
                    // und HTTP-Status — die Kachel las nur users/sites und meldete
                    // bei 40× HTTP 403 grün „3 Person(en) auf 20 Liste(n)
                    // berechtigt". `grants` ist die Zahl der wirklich gesetzten
                    // Rechte; `users` sagt nur, wie viele Adressen aufgelöst wurden.
                    // `resultIsError` hängt am Wort „Fehler"/„Error" — deshalb steht
                    // es vorn im Text. Vorbild: repairAllOrganizerPermissions
                    // (context/actions/maintenance.ts). Die Termin-Listen werden nur
                    // nachgeladen, wenn nichts fehlgeschlagen ist — sonst zeigte der
                    // Reload dieselben „0", und das Banner verschwände zu Unrecht.
                    const failed = r.failed || [];
                    const siteName = (s: string): string => s.replace(/\/+$/, '').split('/').pop() || s;
                    if (failed.length > 0) {
                      const shown = failed.slice(0, 5).map(f => `${siteName(f.site)} [${f.scope}] HTTP ${f.status}`).join(', ');
                      const more = failed.length > 5 ? ' …' : '';
                      setRepairPermsResult(isDe
                        ? `Fehler: ${failed.length} Zuweisung(en) fehlgeschlagen — ${shown}${more} · ${r.grants} Recht(e) gesetzt${unresolved}`
                        : `Error: ${failed.length} assignment(s) failed — ${shown}${more} · ${r.grants} grant(s) set${unresolved}`);
                    } else {
                      setRepairPermsResult(isDe
                        ? `${r.grants} Recht(e) für ${r.users} Person(en) auf ${r.sites} Liste(n) gesetzt${unresolved}`
                        : `${r.grants} grant(s) for ${r.users} person(s) on ${r.sites} list(s) set${unresolved}`);
                      setSubRegReloadTick(t => t + 1);
                    }
                  } catch {
                    setRepairPermsResult(isDe ? 'Fehler beim Setzen der Berechtigungen' : 'Error setting permissions');
                  }
                  setIsRepairingPerms(false);
                }}
              />
            )}

            {/* v31.4: QR-Nummern nachtragen. Die Spalte `QrSentId` (welche
                Nummer stand in der versendeten QR-Mail?) gibt es erst seit
                v31.4 — für jedes Event, dessen QR-Mails vorher rausgingen,
                ist sie leer. Genau dort tut der Fehler weh: Jede Abmeldung
                nummeriert die Liste neu, die gedruckte Mail nicht, und der
                Tisch checkt beim Abtippen der Mail-Nummer die falsche Person
                ein. Die Nummern stehen noch in der Mail-Warteschlange. */}
            {(isAdmin || isOrganizerFor(selectedEvent)) && (
              <ActionTile
                icon={<QrCode size={18} />}
                category="maintenance"
                title={isDe ? 'QR-Nummern nachtragen' : 'Backfill QR numbers'}
                desc={whenList(isDe
                  ? 'Liest aus den bereits verschickten QR-Mails zurück, welche Nummer bei welcher Person gedruckt stand, und trägt sie in der Teilnehmerliste nach — für dieses Event und alle Termine. Danach greift am Check-in die Nummer aus der Mail, auch wenn Abmeldungen die laufende Nummer längst verschoben haben. Zeigt vorher eine Vorschau; geschrieben wird erst nach deiner Bestätigung.'
                  : 'Recovers from the QR emails already sent which number was printed for which person and writes it back to the attendee list — for this event and every date. After that, check-in resolves the number from the email even when cancellations have shifted the running number. Shows a preview first; nothing is written before you confirm.')}
                badge="organizer"
                disabled={!selectedEvent?.subsiteUrl}
                onClick={() => setQrBackfillOpen(true)}
              />
            )}

            {/* v28.23: Teilnehmer-Register nachziehen. Der Dual-Write nach
                DEX_Participants läuft best-effort — fehlt ein Eintrag, sieht die
                Person das Event NICHT in „Meine Events" und die Doppel-Anmelde-
                Vorwarnung (v28.22) greift für sie nicht. Diese Aktion gleicht
                Klammer + alle Sub-Events in einem Lauf ab. */}
            {(isAdmin || (!!selectedEvent && isOrganizerFor(selectedEvent))) && (
              <ActionTile
                icon={<RefreshCw size={18} />}
                category="maintenance"
                title={isSyncingRegistry
                  ? (isDe ? 'Register wird abgeglichen…' : 'Syncing registry…')
                  : (isDe ? 'Teilnehmer-Register nachziehen' : 'Sync participant registry')}
                desc={whenList(isDe
                  ? 'Gleicht die zentrale Teilnehmer-Übersicht mit den echten Anmeldungen dieses Events ab (inkl. aller Sub-Events). Nötig, wenn jemand angemeldet ist, das Event aber nicht in „Meine Events" sieht — dann fehlt der zentrale Eintrag, und auch die Warnung vor doppelter Anmeldung greift für diese Person nicht. Ergänzt nur fehlende Einträge; es wird nichts abgemeldet und niemand benachrichtigt.'
                  : 'Reconciles the central participant registry with this event\'s actual registrations (including all sub-events). Needed when someone is registered but does not see the event in “My events” — the central entry is missing then, and the duplicate-registration warning does not work for that person. Only adds missing entries; nothing is cancelled and nobody is notified.')}
                badge="organizer"
                busy={isSyncingRegistry}
                disabled={!selectedEvent?.subsiteUrl}
                result={syncRegistryResult}
                resultIsError={!!syncRegistryResult && (syncRegistryResult.indexOf('Fehler') >= 0 || syncRegistryResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  const targets = [selectedEvent, ...childEventsOf(selectedEvent.id)]
                    .filter(e => !!e.subsiteUrl && !!e.eventNumber);
                  if (!(await confirmDialog(isDe
                    ? `Teilnehmer-Register für ${targets.length} Liste(n) abgleichen (Event + Sub-Events)?\n\nFehlende Einträge werden ergänzt, damit die Betroffenen ihre Anmeldung in „Meine Events" sehen. Es werden keine Anmeldungen geändert und keine Mails verschickt. Bei vielen Teilnehmern kann das ein paar Minuten dauern.`
                    : `Reconcile the participant registry for ${targets.length} list(s) (event + sub-events)?\n\nMissing entries are added so people see their registration in „My events". No registrations are changed and no emails are sent. With many participants this can take a few minutes.`,
                    { confirmLabel: isDe ? 'Abgleichen' : 'Reconcile' }))) return;
                  setIsSyncingRegistry(true);
                  setSyncRegistryResult(null);
                  let activeTotal = 0;
                  let fixedTotal = 0;
                  let failedTotal = 0;
                  const errored: string[] = [];
                  // v28.25: Den ECHTEN Fehlertext festhalten (z.B. „DEX_Participants
                  // nicht lesbar (HTTP 500)") — vorher blieb nur ein nacktes
                  // „Fehler bei: <Event>" übrig und die Ursache war unsichtbar.
                  let firstErrText = '';
                  for (const ev of targets) {
                    try {
                      const r = await eventServiceRef.backfillParticipantRegistry(
                        ev.subsiteUrl as string,
                        ev.eventNumber as number,
                      );
                      activeTotal += r.active;
                      fixedTotal += r.fixed;
                      failedTotal += r.failed;
                    } catch (err) {
                      errored.push((ev.title || '?').slice(0, 40));
                      if (!firstErrText) firstErrText = (err instanceof Error ? err.message : String(err || '')).slice(0, 300);
                    }
                  }
                  setSyncRegistryResult(isDe
                    ? `${activeTotal} aktive Anmeldung(en) geprüft, ${fixedTotal} Register-Eintrag/Einträge ergänzt${failedTotal > 0 ? `, ${failedTotal} fehlgeschlagen` : ''}.${errored.length > 0 ? ` Fehler bei: ${errored.join(', ')}. ${firstErrText}` : ''}`
                    : `${activeTotal} active registration(s) checked, ${fixedTotal} registry entr(y/ies) added${failedTotal > 0 ? `, ${failedTotal} failed` : ''}.${errored.length > 0 ? ` Errors on: ${errored.join(', ')}. ${firstErrText}` : ''}`);
                  setIsSyncingRegistry(false);
                }}
              />
            )}

            {/* 8. Spalten fixen — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<Columns size={18} />}
                category="maintenance"
                title={isFixingColumns ? (isDe ? 'Spalten werden gefixt…' : 'Fixing columns…') : (isDe ? 'Spalten fixen' : 'Fix columns')}
                desc={whenList(isDe
                  ? 'Bringt die Teilnehmerliste (und die aller Sub-Events) auf den aktuellen Stand: legt fehlende Spalten an, räumt überflüssige weg (z.B. StarterType ohne Gruppen) und sortiert die Spalten-Reihenfolge. Nutzen, wenn in der Liste Spalten fehlen oder Antworten nicht ankommen. Leere Duplikat-Spalten werden nach Rückfrage gelöscht.'
                  : 'Brings the participant list (and those of all sub-events) up to date: creates missing columns, removes superfluous ones (e.g. StarterType without groups) and fixes the column order. Use it when columns are missing or answers do not arrive. Empty duplicate columns are deleted after confirmation.')}
                badge="admin"
                busy={isFixingColumns}
                disabled={!selectedEvent?.subsiteUrl}
                result={fixColumnsResult}
                resultIsError={!!fixColumnsResult && (fixColumnsResult.indexOf('Fehler') >= 0 || fixColumnsResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  setIsFixingColumns(true);
                  setFixColumnsResult(null);
                  try {
                    const isB2Run = !!(selectedEvent.durchstarterCapacity || selectedEvent.funstarterCapacity);
                    const hasQuiz = !!(selectedEvent.quiz && selectedEvent.quiz.length > 0);
                    const customFields = (selectedEvent.eventSpecificFields || []).map(f => ({
                      id: f.id, label: f.label, type: f.type, required: f.required, options: f.options,
                      visible: true,
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      spInternalName: (f as any).spInternalName || '',
                    }));
                    const result = await eventServiceRef.fixRegistrationListColumns(
                      selectedEvent.subsiteUrl,
                      { isB2Run, hasQuiz, customFields },
                      (count, titles) => {
                        const preview = titles.slice(0, 8).map(t => `„${t}"`).join(', ');
                        const more = titles.length > 8 ? (isDe ? ` …und ${titles.length - 8} weitere` : ` …and ${titles.length - 8} more`) : '';
                        // v20.4: App-Modal statt window.confirm — der Service
                        // akzeptiert boolean | Promise<boolean> und awaitet.
                        return confirmDialog(isDe
                          ? `${count} überflüssige (leere) Duplikat-Spalten in der Teilnehmerliste gefunden ` +
                            `(${titles.length} Titel betroffen: ${preview}${more}).\n\n` +
                            `Diese werden jetzt gelöscht (irreversibel). Spalten mit Daten bleiben erhalten ` +
                            `und werden zur manuellen Prüfung gemeldet.\n\nFortfahren?`
                          : `Found ${count} redundant (empty) duplicate columns in the participant list ` +
                            `(${titles.length} titles affected: ${preview}${more}).\n\n` +
                            `These will now be deleted (irreversible). Columns with data are kept ` +
                            `and reported for manual review.\n\nProceed?`
                        );
                      }
                    );
                    const msgs: string[] = [];
                    if (result.added.length > 0) msgs.push(isDe ? `Spalten hinzugefügt: ${result.added.join(', ')}` : `Columns added: ${result.added.join(', ')}`);
                    if (result.removed.length > 0) msgs.push(isDe ? `Spalten entfernt: ${result.removed.join(', ')}` : `Columns removed: ${result.removed.join(', ')}`);
                    if (result.duplicatesRemoved && result.duplicatesRemoved.length > 0) {
                      msgs.push(isDe ? `${result.duplicatesRemoved.length} leere Duplikate gelöscht` : `${result.duplicatesRemoved.length} empty duplicates deleted`);
                    }
                    if (result.duplicatesWithData && result.duplicatesWithData.length > 0) {
                      const list = result.duplicatesWithData.map(t => `„${t}"`).join(', ');
                      msgs.push(isDe ? `${result.duplicatesWithData.length} Duplikate mit Daten — bitte manuell prüfen: ${list}` : `${result.duplicatesWithData.length} duplicates with data — please review manually: ${list}`);
                    }
                    if (result.viewFixed) msgs.push(isDe ? 'View-Reihenfolge korrigiert' : 'View order fixed');
                    if (result.customFieldMap && Object.keys(result.customFieldMap).length > 0) {
                      // v26.13 DATENVERLUST-FIX: aus den VOLLEN geparsten Feldern
                      // bauen (helpText/showIf/EN-Varianten erhalten), nicht aus
                      // dem gestrippten `customFields`. Sonst löscht „Spalten fixen"
                      // alle Beschreibungen.
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const updatedCf = (selectedEvent.eventSpecificFields || []).map((f: any) => {
                        const sp = result.customFieldMap![f.id];
                        return sp ? { ...f, spInternalName: sp } : { ...f };
                      });
                      try {
                        await updateEvent(selectedEvent.id, { 'CustomFields': JSON.stringify(updatedCf) });
                        msgs.push(isDe ? `Custom-Field-Zuordnung aktualisiert (${Object.keys(result.customFieldMap).length})` : `Custom field mapping updated (${Object.keys(result.customFieldMap).length})`);
                      } catch {
                        msgs.push(isDe ? 'WARN: Custom-Field-Mapping konnte nicht am Event gespeichert werden' : 'WARN: custom field mapping could not be saved on the event');
                      }
                    }
                    // v24.32: „Spalten fixen" deckt jetzt auch ALLE Sub-Events ab
                    // (eigene Teilnehmerlisten) — sonst fehlt z.B. die neue
                    // Company-Spalte dort. Pro Sub-Event mit eigener Subsite die
                    // Spalten + View fixen und das Custom-Field-Mapping nachziehen.
                    const subKids = childEventsOf(selectedEvent.id).filter(c => (c.subsiteUrl || '').trim());
                    let subFixed = 0; let subFailed = 0;
                    for (const child of subKids) {
                      try {
                        const childCf = (child.eventSpecificFields || []).map(f => ({
                          id: f.id, label: f.label, type: f.type, required: f.required, options: f.options,
                          visible: true,
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          spInternalName: (f as any).spInternalName || '',
                        }));
                        const childB2 = !!(child.durchstarterCapacity || child.funstarterCapacity);
                        const childQuiz = !!(child.quiz && child.quiz.length > 0);
                        const cr = await eventServiceRef.fixRegistrationListColumns(child.subsiteUrl, { isB2Run: childB2, hasQuiz: childQuiz, customFields: childCf });
                        if (cr.customFieldMap && Object.keys(cr.customFieldMap).length > 0) {
                          // v26.13 DATENVERLUST-FIX: volle Sub-Event-Felder erhalten
                          // (helpText etc.), nicht aus dem gestrippten childCf bauen.
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const updCf = (child.eventSpecificFields || []).map((f: any) => { const sp = cr.customFieldMap![f.id]; return sp ? { ...f, spInternalName: sp } : { ...f }; });
                          try { await updateEvent(child.id, { 'CustomFields': JSON.stringify(updCf) }); } catch { /* best-effort */ }
                        }
                        subFixed++;
                      } catch (err) { subFailed++; console.warn('[DEX] Spalten fixen (Sub-Event) fehlgeschlagen:', child.id, err); }
                    }
                    if (subKids.length > 0) {
                      msgs.push(isDe
                        ? `Sub-Events geprüft: ${subFixed}/${subKids.length}${subFailed ? ` (${subFailed} mit Fehler)` : ''}`
                        : `Sub-events checked: ${subFixed}/${subKids.length}${subFailed ? ` (${subFailed} with errors)` : ''}`);
                    }
                    const finalMsg = msgs.length > 0 ? msgs.join(' | ') : (isDe ? 'Alles OK, keine Änderungen nötig.' : 'All OK, no changes needed.');
                    setFixColumnsResult(finalMsg);
                    // v19.10: Ergebnis zusätzlich als Dialog zeigen. Im Aktionen-
                    // Dropdown rendert die ActionTile (und damit ihr `result`-Text)
                    // `null` — vorher kam nach „Spalten fixen" daher GAR KEINE
                    // sichtbare Rückmeldung. Ein window.alert ist garantiert sichtbar.
                    showAlert((isDe ? '„Spalten fixen" — Ergebnis:\n\n' : 'Fix columns — result:\n\n') + finalMsg);
                  } catch {
                    const errMsg = isDe ? 'Fehler beim Fixen der Spalten.' : 'Error fixing columns.';
                    setFixColumnsResult(errMsg);
                    showAlert(errMsg);
                  }
                  setIsFixingColumns(false);
                }}
              />
            )}

            {/* 9. Felder reparieren — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<Wrench size={18} />}
                category="maintenance"
                title={isFixingFields ? (isDe ? 'Felder werden repariert…' : 'Repairing fields…') : (isDe ? 'Felder reparieren' : 'Repair fields')}
                desc={isDe
                  ? 'Räumt die Anmeldefelder dieses Events automatisch auf: AGB/Datenschutz wird eine echte Checkbox, die T-Shirt-Auswahl bekommt „Ohne T-Shirt“ und wird optional, bei B2Run-Events werden die Sonderfelder ergänzt, doppelte „(Pflicht)“-Zusätze verschwinden.'
                  : 'Cleans up this event’s registration fields automatically: terms/privacy becomes a real checkbox, the T-shirt choice gets “Ohne T-Shirt” and becomes optional, B2Run events get their special fields added, redundant “(required)” suffixes disappear.'}
                badge="admin"
                busy={isFixingFields}
                disabled={!selectedEvent}
                result={fixFieldsResult}
                resultIsError={!!fixFieldsResult && (fixFieldsResult.startsWith('Fehler') || fixFieldsResult.startsWith('Update fehl') || fixFieldsResult.startsWith('Error') || fixFieldsResult.startsWith('Update failed'))}
                onClick={async () => {
                  if (!selectedEvent) return;
                  setIsFixingFields(true);
                  setFixFieldsResult(null);
                  try {
                    const changes: string[] = [];
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const raw: any[] = (selectedEvent.eventSpecificFields || []).map((f: any) => ({ ...f }));
                    const hasField = (id: string): boolean => raw.some(f => f.id === id);
                    const isB2Run = raw.some(f => String(f.id || '').indexOf('b2run_') === 0);
                    if (isB2Run) {
                      if (!hasField('b2run_infoservice')) {
                        raw.push({ id: 'b2run_infoservice', label: 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)', type: 'checkbox', required: false, options: [], visible: true });
                        changes.push("Feld ergänzt: 'Infoservice'");
                      }
                      if (!hasField('b2run_anonym')) {
                        raw.push({ id: 'b2run_anonym', label: 'Anonym teilnehmen', type: 'checkbox', required: false, options: [], visible: true });
                        changes.push("Feld ergänzt: 'Anonym teilnehmen'");
                      }
                      const hasLaufshirt = raw.some(f => f.id === 'b2run_laufshirt' || /laufshirt/i.test(String(f.label || '')));
                      if (!hasLaufshirt) {
                        raw.push({ id: 'b2run_laufshirt', label: 'Deloitte-Laufshirt', type: 'select', required: true, options: ['Habe bereits ein Laufshirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true });
                        changes.push("Feld ergänzt: 'Deloitte-Laufshirt' (Pflicht)");
                      }
                    }
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const fixed = raw.map((f: any) => {
                      const nf = { ...f };
                      const label = String(nf.label || '');
                      const lowLabel = label.toLowerCase();
                      const isConsent = lowLabel.indexOf('zustimmung') >= 0
                        || lowLabel.indexOf('agb') >= 0
                        || lowLabel.indexOf('datenschutz') >= 0;
                      const isB2RunCheckbox = ['b2run_infoservice', 'b2run_anonym', 'b2run_datenschutz'].indexOf(nf.id) >= 0;
                      if ((isConsent || isB2RunCheckbox) && nf.type !== 'checkbox') {
                        nf.type = 'checkbox';
                        nf.options = [];
                        changes.push(`${label} -> Checkbox`);
                      }
                      // v30.67: Idempotent machen. Der Guard prüfte auf 'kein' — ein
                      // Wort, das in keinem der beiden eingefügten Texte vorkommt
                      // ('Ohne T-Shirt', 'Habe bereits ein Laufshirt'). Jeder Klick
                      // hängte deshalb eine weitere Kopie vorn an. Und 'Laufshirt'
                      // enthält 'shirt': Ein Laufshirt-Feld lief durch BEIDE Zweige,
                      // bekam 'Ohne T-Shirt' UND ein zweites 'Habe bereits …', und
                      // `required` kippte erst auf false, dann zurück auf true.
                      const uniqOpts = (arr: string[]): string[] => arr.filter((o, i) => arr.indexOf(o) === i);
                      const hasOpt = (arr: string[], text: string): boolean =>
                        arr.some((o: string) => o.trim().toLowerCase() === text.toLowerCase() || o.toLowerCase().indexOf('kein') >= 0);
                      const isLaufshirt = nf.id === 'b2run_laufshirt' || /laufshirt/i.test(label);
                      const isShirt = !isLaufshirt && (lowLabel.indexOf('t-shirt') >= 0 || lowLabel.indexOf('tshirt') >= 0 || lowLabel.indexOf('shirt') >= 0);
                      if (isShirt && nf.type === 'select') {
                        const rawOpts: string[] = Array.isArray(nf.options) ? nf.options.slice() : [];
                        const opts = uniqOpts(rawOpts);
                        if (opts.length !== rawOpts.length) {
                          nf.options = opts;
                          changes.push(`${label}: doppelte Optionen entfernt`);
                        }
                        if (!hasOpt(opts, 'Ohne T-Shirt')) {
                          opts.unshift('Ohne T-Shirt');
                          nf.options = opts;
                          changes.push(`${label} -> 'Ohne T-Shirt'-Option`);
                        }
                        if (nf.required) {
                          nf.required = false;
                          changes.push(`${label} -> optional`);
                        }
                      }
                      const stripped = label.replace(/\s*\((?:pflicht|mandatory|required)\)\s*$/i, '').trim();
                      if (stripped && stripped !== label) {
                        nf.label = stripped;
                        changes.push(`Label "${label}" -> "${stripped}"`);
                      }
                      if (nf.id === 'b2run_mobilnummer') {
                        if (nf.required) { nf.required = false; changes.push('Mobilnummer -> optional'); }
                        if (nf.label === 'Mobilnummer') {
                          nf.label = 'Mobilnummer (nur bei aktiviertem Infoservice)';
                          changes.push("Mobilnummer-Label präzisiert");
                        }
                      }
                      if (nf.id === 'b2run_infoservice' && nf.label && nf.label.indexOf('benötigt') >= 0) {
                        nf.label = 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)';
                        changes.push('Infoservice-Label modernisiert');
                      }
                      if (nf.id === 'b2run_datenschutz') {
                        const needLinks = !Array.isArray(nf.externalLinks) || nf.externalLinks.length === 0;
                        if (needLinks) {
                          nf.externalLinks = [
                            { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
                            { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
                          ];
                          changes.push('B2Run-Datenschutz: AGB + Datenschutz Links ergänzt');
                        }
                      }
                      if (isLaufshirt) {
                        if (!nf.required) {
                          nf.required = true;
                          changes.push(`${label || nf.id}: als Pflichtfeld markiert`);
                        }
                        if (nf.type === 'select') {
                          const rawOpts: string[] = Array.isArray(nf.options) ? nf.options.slice() : [];
                          const opts = uniqOpts(rawOpts);
                          if (opts.length !== rawOpts.length) {
                            nf.options = opts;
                            changes.push(`${label || nf.id}: doppelte Optionen entfernt`);
                          }
                          if (!hasOpt(opts, 'Habe bereits ein Laufshirt')) {
                            opts.unshift('Habe bereits ein Laufshirt');
                            nf.options = opts;
                            changes.push(`${label || nf.id}: 'Habe bereits ein Laufshirt'-Option hinzugefügt`);
                          }
                        }
                      }
                      return nf;
                    });
                    const dsIdx = fixed.findIndex((f: { id: string }) => f.id === 'b2run_datenschutz');
                    if (dsIdx >= 0 && dsIdx !== fixed.length - 1) {
                      const [ds] = fixed.splice(dsIdx, 1);
                      fixed.push(ds);
                      changes.push('Zustimmung-Checkbox ans Ende verschoben');
                    }
                    const ok = await updateEvent(selectedEvent.id, { CustomFields: JSON.stringify(fixed) });
                    if (ok) {
                      setFixFieldsResult(changes.length > 0
                        ? (isDe ? `Geändert: ${changes.join(' | ')}` : `Changed: ${changes.join(' | ')}`)
                        : (isDe ? 'Keine Änderungen nötig.' : 'No changes needed.'));
                    } else {
                      setFixFieldsResult(isDe ? 'Update fehlgeschlagen.' : 'Update failed.');
                    }
                  } catch (err) {
                    setFixFieldsResult((isDe ? 'Fehler: ' : 'Error: ') + (err instanceof Error ? err.message : String(err)));
                  }
                  setIsFixingFields(false);
                }}
              />
            )}

            {/* 7b. Counter zurücksetzen — Admin only (v9.13 → v11.27).
                Recovery-Button um den DEX_TeilnehmerCounter EXAKT auf
                max(TeilnehmerID) der Subsite zu setzen. Bidirektional:
                Counter wird hochgezogen wenn er drunter steht (gegen
                Doppel-IDs), oder runtergesetzt wenn er drüber steht
                (z.B. nach vielen Abmeldungen, die TIDs gefressen
                haben). Vorher (vor v11.27) lief es nur monotonic-up,
                weshalb ein zu hoher Counter (Counter=11, Max-TID=4)
                nicht zurückgesetzt wurde — Klick auf den Button
                hatte dann keinen sichtbaren Effekt. */}
            {isAdmin && (
              <ActionTile
                icon={<Hash size={18} />}
                category="maintenance"
                title={isResettingCounter ? (isDe ? 'Counter wird zurückgesetzt…' : 'Resetting counter…') : (isDe ? 'Counter zurücksetzen' : 'Reset counter')}
                desc={whenList(isDe
                  ? 'Setzt den Zähler der Teilnehmer-Nummern exakt auf die höchste vergebene Nummer — nach oben wie nach unten. Neue Anmeldungen bekommen danach wieder die nächste passende Nummer. Nutzen, wenn neue Anmeldungen mit offensichtlich falschen Nummern starten (viel zu hoch nach vielen Abmeldungen, oder wieder bei 1).'
                  : 'Sets the participant-number counter exactly to the highest number in use — up or down. New registrations then get the next fitting number again. Use it when new registrations start with obviously wrong numbers (far too high after many cancellations, or back at 1).')}
                badge="admin"
                busy={isResettingCounter}
                disabled={!selectedEvent?.subsiteUrl}
                result={resetCounterResult}
                resultIsError={!!resetCounterResult && (resetCounterResult.indexOf('Fehler') >= 0 || resetCounterResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  if (!(await confirmDialog(isDe ? 'Counter auf aktuellen Max-Wert zurücksetzen?' : 'Reset counter to the current max value?'))) return;
                  setIsResettingCounter(true);
                  setResetCounterResult(null);
                  try {
                    const result = await eventServiceRef.resetCounterToMax(selectedEvent.subsiteUrl);
                    setResetCounterResult(isDe
                      ? `Counter steht jetzt auf ${result.counter} (Max-TID: ${result.max})`
                      : `Counter is now at ${result.counter} (max ID: ${result.max})`);
                  } catch {
                    setResetCounterResult(isDe ? 'Fehler beim Zurücksetzen des Counters' : 'Error resetting the counter');
                  }
                  setIsResettingCounter(false);
                }}
              />
            )}

            {/* 10. Profile neu laden — Admin only */}
            {isAdmin && (
              <ActionTile
                icon={<RefreshCw size={18} />}
                category="maintenance"
                title={isRefreshingProfiles ? (isDe ? 'Teilnehmer werden nachgeladen…' : 'Reloading attendees…') : (isDe ? 'Teilnehmer nachladen (Daten reparieren)' : 'Reload attendees (repair data)')}
                desc={whenList(isDe
                  ? 'Fragt, wie viele der letzten Anmeldungen geprüft werden sollen, und lädt für sie Name, Position, Standort, Bereich und Telefon frisch aus dem Microsoft-365-Benutzerprofil. Repariert auch kaputte Namen — z.B. wenn statt des Vornamens ein technisches Anmelde-Kürzel in der Liste steht.'
                  : 'Asks how many of the most recent registrations to check and reloads their name, job title, location, department and phone from the Microsoft 365 user profile. Also repairs broken names — e.g. when a technical login token appears instead of the first name.')}
                badge="admin"
                busy={isRefreshingProfiles}
                disabled={!selectedEvent?.subsiteUrl}
                result={refreshProfilesResult}
                resultIsError={!!refreshProfilesResult && (refreshProfilesResult.indexOf('Fehler') >= 0 || refreshProfilesResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent?.subsiteUrl) return;
                  const ans = prompt(isDe
                    ? 'Wie viele der letzten Teilnehmer sollen aus dem Benutzerprofil neu geladen werden? (JobTitle, Standort, Department, Phone)'
                    : 'How many of the most recent participants should be reloaded from the user profile? (job title, location, department, phone)', '20');
                  if (!ans) return;
                  const n = parseInt(ans, 10);
                  if (isNaN(n) || n <= 0) { showAlert(isDe ? 'Bitte eine positive Zahl eingeben.' : 'Please enter a positive number.'); return; }
                  setIsRefreshingProfiles(true);
                  setRefreshProfilesResult(null);
                  try {
                    const result = await eventServiceRef.fixEventParticipantsProfileData(selectedEvent.subsiteUrl, n);
                    setRefreshProfilesResult(isDe
                      ? `${result.scanned} geprüft, ${result.updated} aktualisiert, ${result.failedLookups} Profil-Lookups fehlgeschlagen`
                      : `${result.scanned} checked, ${result.updated} updated, ${result.failedLookups} profile lookups failed`);
                    await reloadRegistrations();
                  } catch {
                    setRefreshProfilesResult(isDe ? 'Fehler beim Auffrischen der Profile' : 'Error refreshing profiles');
                  }
                  setIsRefreshingProfiles(false);
                }}
              />
            )}

            {/* v11.11: Custom-Fields aus Versionsverlauf zurückholen.
                Hilft den Admins, denen die v11.9-Migration die b2run_*-
                Felder (Altersgruppe, T-Shirt-Größe etc.) versehentlich
                aus customFields entfernt hat. Liest die SP-Versionen des
                Event-Items, sucht die jüngste Version mit b2run_*-
                Feldern und mergt diese zurück in das aktuelle
                CustomFields-Array. Bestehende Felder bleiben unverändert
                — es werden NUR fehlende b2run_*-Felder ergänzt. */}

            {isAdmin && selectedEvent && (
              <ActionTile
                icon={<RefreshCw size={18} />}
                category="maintenance"
                title={isDe ? 'Custom-Fields aus Versionsverlauf zurückholen' : 'Restore custom fields from version history'}
                desc={isDe
                  ? 'Liest den Versionsverlauf des Events und holt verloren gegangene b2run_*-Anmeldefelder zurück (Altersgruppe, T-Shirt-Größe, Startblock, Mobilnummer …) — nötig nach der v11.9-Migration, die diese Felder versehentlich entfernt hat. Bestehende Felder bleiben unangetastet, nur Fehlendes wird ergänzt.'
                  : 'Reads the event’s version history and restores lost b2run_* registration fields (age group, T-shirt size, start block, mobile number …) — needed after the v11.9 migration that deleted these fields by accident. Existing fields stay untouched, only missing ones are added.'}
                badge="admin"
                onClick={async () => {
                  if (!eventServiceRef || !selectedEvent) return;
                  try {
                    const history = await eventServiceRef.getEventCustomFieldsHistory(parseInt(selectedEvent.id, 10));
                    if (history.length === 0) {
                      showAlert(isDe ? 'Kein Versionsverlauf gefunden — entweder hat das Event keine Versionen oder der Zugriff wurde verweigert.' : 'No version history found — the event has no versions or access was denied.');
                      return;
                    }
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const currentFields: any[] = (selectedEvent.eventSpecificFields || []).map(f => ({ ...f }));
                    const currentIds = new Set(currentFields.map(f => String(f.id || '').toLowerCase()));
                    // Jüngste Version mit b2run_*-Feldern finden, die noch
                    // NICHT in currentFields stecken.
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    let foundFields: any[] = [];
                    let foundVersion = '';
                    let foundModified = '';
                    for (const v of history) {
                      const missingB2run = v.customFields.filter(f => {
                        const id = String(f.id || '').toLowerCase();
                        return id.indexOf('b2run_') === 0 && !currentIds.has(id);
                      });
                      if (missingB2run.length > 0) {
                        foundFields = missingB2run;
                        foundVersion = v.versionLabel;
                        foundModified = v.modified;
                        break;
                      }
                    }
                    if (foundFields.length === 0) {
                      showAlert(isDe ? 'Keine fehlenden b2run_*-Felder im Versionsverlauf gefunden — entweder sind alle Felder schon vorhanden oder es gab nie welche.' : 'No missing b2run_* fields found in the version history — either all fields already exist or there never were any.');
                      return;
                    }
                    const fieldList = foundFields.map(f => `• ${f.label || f.id}`).join('\n');
                    const modifiedDate = foundModified ? new Date(foundModified).toLocaleString(isDe ? 'de-DE' : 'en-GB') : '?';
                    if (!(await confirmDialog(isDe
                      ? `Folgende ${foundFields.length} Custom-Field(s) aus Version ${foundVersion} (${modifiedDate}) zurückholen?\n\n${fieldList}\n\nDie Felder werden ans Ende deiner aktuellen Felder-Liste angehängt. Du kannst sie danach im Wizard frei umbenennen, neu sortieren oder löschen.`
                      : `Restore the following ${foundFields.length} custom field(s) from version ${foundVersion} (${modifiedDate})?\n\n${fieldList}\n\nThe fields are appended to the end of your current field list. You can rename, reorder or delete them afterwards in the wizard.`, { confirmLabel: isDe ? 'Zurückholen' : 'Restore' }))) {
                      return;
                    }
                    const merged = [...currentFields, ...foundFields];
                    const ok = await updateEvent(selectedEvent.id, { 'CustomFields': JSON.stringify(merged) });
                    if (!ok) {
                      showAlert(isDe ? 'Update fehlgeschlagen — siehe Browser-Console.' : 'Update failed — see browser console.');
                      return;
                    }
                    // Subsite-Spalten gleich mit-syncen, damit die b2run_*-
                    // Spalten in der Teilnehmerliste wieder existieren.
                    if (selectedEvent.subsiteUrl) {
                      try {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const cfForFix: any[] = merged.map((f: any) => ({
                          id: f.id,
                          label: f.label,
                          type: f.type,
                          required: !!f.required,
                          visible: true,
                          options: f.options || [],
                          spInternalName: f.spInternalName || '',
                          ...(f.helpText ? { helpText: f.helpText } : {}),
                          ...(f.multi ? { multi: true } : {}),
                          ...(f.showIf ? { showIf: f.showIf } : {}),
                        }));
                        const splitActive = (selectedEvent.durchstarterCapacity || 0) > 0 && (selectedEvent.funstarterCapacity || 0) > 0;
                        await eventServiceRef.fixRegistrationListColumns(selectedEvent.subsiteUrl, {
                          isB2Run: splitActive,
                          hasQuiz: (selectedEvent.quiz || []).length > 0,
                          customFields: cfForFix,
                        });
                      } catch (err) { console.warn('[DEX] fixRegistrationListColumns nach Restore fehlgeschlagen:', err); }
                    }
                    await refreshEvents();
                    showAlert(isDe
                      ? `${foundFields.length} Custom-Field(s) erfolgreich aus Version ${foundVersion} zurückgeholt.`
                      : `${foundFields.length} custom field(s) successfully restored from version ${foundVersion}.`);
                  } catch (err) {
                    console.warn('[DEX] restore custom fields from history failed:', err);
                    showAlert(isDe ? 'Zurückholen fehlgeschlagen — siehe Browser-Console.' : 'Restore failed — see browser console.');
                  }
                }}
              />
            )}

            {/* v27.13: „In SharePoint öffnen" entfernt — alle Teilnehmer-
                Aktionen (Bearbeiten, Export, Massenimport, Audit) laufen über
                die App. Direktes Editieren in der rohen SP-Liste erzeugte
                Zeilen ohne Audit-Felder und ohne Format-Validierung (siehe
                Feedback Datenschutz-Review 07/2026). */}

            {/* 5d. v30.56: Die SharePoint-Liste direkt öffnen — Admin only.
                Für die Fälle, die keine App-Ansicht abbildet: eine Spalte
                nachsehen, eine Zeile von Hand korrigieren, den echten
                Datenstand gegen eine Anzeige halten. Bewusst nur für Admins:
                Die Liste kennt keine der Schutzregeln der App. */}
            {isAdmin && selectedEvent && selectedEvent.subsiteUrl && (
              <ActionTile
                icon={<ExternalLink size={18} />}
                category="maintenance"
                title={isDe ? 'Teilnehmerliste in SharePoint öffnen' : 'Open participant list in SharePoint'}
                desc={isDe
                  ? 'Öffnet die zugrunde liegende SharePoint-Liste dieses Events in einem neuen Tab — für Nachschauen und Korrekturen, die die App nicht abbildet. Achtung: Dort gelten die Prüfungen der App nicht, Änderungen wirken sofort und werden nicht ins Änderungsprotokoll geschrieben.'
                  : 'Opens the underlying SharePoint list of this event in a new tab — for lookups and corrections the app does not cover. Caution: none of the app’s checks apply there, changes take effect immediately and are not written to the change history.'}
                badge="admin"
                onClick={() => {
                  const url = `${selectedEvent.subsiteUrl}/Lists/${encodeURIComponent(REG_LIST_NAME)}/AllItems.aspx`;
                  window.open(url, '_blank', 'noopener,noreferrer');
                }}
              />
            )}

            {/* v31.3 · Gruppe WARTUNG (alle Events) — Admin-Läufe über den gesamten Bestand. */}
            {/* v20.6: Fremd-Anmeldungen: Zugriff reparieren (alle aktiven
                Events) — Admin only. Geht alle Teilnehmerlisten durch, stellt
                die "nur eigene Elemente"-Sicherheit sicher und setzt bei
                Anmeldungen durch Dritte den Zeilen-Autor auf den Teilnehmer,
                damit die angemeldete Person ihre Anmeldung in "Meine Events"
                sieht und sich selbst abmelden kann (v20.5 rückwirkend). */}
            {isAdmin && (
              <ActionTile
                icon={<Wrench size={18} />}
                category="maintenance"
                title={isRepairingAccess ? (isDe ? 'Prüfung läuft…' : 'Check running…') : (isDe ? 'Fremd-Anmeldungen: Zugriff reparieren (alle aktiven Events)' : 'Proxy registrations: repair access (all active events)')}
                desc={isDe
                  ? 'Geht alle Teilnehmerlisten aller aktiven Events (inkl. Sub-Events) durch und prüft zwei Dinge: (1) dass jede Liste auf „nur eigene Elemente" steht — also niemand fremde Anmeldedaten sehen kann; (2) dass bei Anmeldungen, die jemand FÜR eine andere Person gemacht hat, die angemeldete Person ihre eigene Anmeldung sehen und sich selbst abmelden kann. Gefundene Probleme werden direkt repariert. Externe Personen ohne Deloitte-Login können dabei nicht berücksichtigt werden.'
                  : 'Walks the participant lists of all active events (incl. sub-events) and checks two things: (1) that every list is set to „own items only" — so nobody can see other people’s registration data; (2) that for registrations someone made FOR another person, the registered person can see their own registration and cancel it themselves. Found issues are repaired directly. External people without a Deloitte login cannot be covered.'}
                badge="admin"
                busy={isRepairingAccess}
                result={repairAccessResult}
                resultIsError={!!repairAccessResult && (repairAccessResult.indexOf('Fehler') >= 0 || repairAccessResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef) return;
                  // Alle aktiven Events (inkl. Sub-Events) mit Teilnehmerliste,
                  // dedupliziert nach Subsite (Reuse-Pfade teilen sich eine).
                  const seen = new Set<string>();
                  const targets = allEvents.filter(ev => {
                    if (ev.status !== 'Active' || !ev.subsiteUrl) return false;
                    const key = ev.subsiteUrl.toLowerCase();
                    if (seen.has(key)) return false;
                    seen.add(key);
                    return true;
                  });
                  if (targets.length === 0) {
                    setRepairAccessResult(isDe ? 'Keine aktiven Events mit Teilnehmerliste gefunden.' : 'No active events with a participant list found.');
                    return;
                  }
                  if (!(await confirmDialog(isDe
                    ? `Zugriffs-Prüfung über ${targets.length} aktive Event-Teilnehmerlisten starten?\n\nGeprüft wird pro Liste die „nur eigene Elemente"-Sicherheit, und bei Anmeldungen durch Dritte bekommt die angemeldete Person Zugriff auf ihre eigene Anmeldung. Je nach Teilnehmerzahl kann das einige Minuten dauern.`
                    : `Start the access check across ${targets.length} active event participant lists?\n\nEach list is checked for „own items only" security, and for registrations made by a third party the registered person gets access to their own registration. Depending on participant counts this can take a few minutes.`, { confirmLabel: isDe ? 'Prüfen & reparieren' : 'Check & repair' }))) return;
                  setIsRepairingAccess(true);
                  setRepairAccessResult(null);
                  // v20.7: Fortschritts-Modal öffnen.
                  setAccessFixModal({ running: true, evIdx: 0, evTotal: targets.length, evTitle: '', itemDone: 0, itemTotal: 0, summary: null });
                  // v22: Globale Queue-Listen (DEX_Outlook/DEX_IDReorder) auf
                  // „nur eigene Elemente" härten — einmal pro Lauf, idempotent.
                  let queueIlsLine = '';
                  try {
                    const q = await eventServiceRef.hardenQueueListsIls();
                    queueIlsLine = q.failed.length === 0
                      ? (isDe ? 'Globale Queue-Listen (Outlook/IDReorder) stehen auf „nur eigene Elemente".' : 'Global queue lists (Outlook/IDReorder) set to „own items only".')
                      : (isDe ? `Queue-Listen-Härtung fehlgeschlagen bei: ${q.failed.join(', ')}.` : `Queue list hardening failed for: ${q.failed.join(', ')}.`);
                  } catch { /* best-effort */ }
                  let listsChecked = 0;
                  let ilsWrong = 0;
                  let ilsFixed = 0;
                  let proxyFound = 0;
                  let authorFixed = 0;
                  let authorFailed = 0;
                  const errorEvents: string[] = [];
                  try {
                    for (let i = 0; i < targets.length; i++) {
                      const ev = targets[i];
                      const shortTitle = (ev.title || '?').slice(0, 40);
                      setAccessFixModal(prev => prev ? { ...prev, evIdx: i + 1, evTitle: shortTitle, itemDone: 0, itemTotal: 0 } : prev);
                      try {
                        const r = await eventServiceRef.repairProxyRegistrationAccess(ev.subsiteUrl as string, (done, total) => {
                          if (done % 10 === 0 || done === total) {
                            setAccessFixModal(prev => prev ? { ...prev, itemDone: done, itemTotal: total } : prev);
                          }
                        });
                        listsChecked++;
                        if (r.ilsWasWrong) {
                          ilsWrong++;
                          if (r.ilsFixed) ilsFixed++;
                        }
                        proxyFound += r.proxyFound;
                        authorFixed += r.authorFixed;
                        authorFailed += r.authorFailed;
                      } catch {
                        errorEvents.push(shortTitle);
                      }
                    }
                    const parts: string[] = [];
                    if (isDe) {
                      parts.push(`${listsChecked} Listen geprüft.`);
                      parts.push(ilsWrong === 0
                        ? 'Listen-Sicherheit überall korrekt („nur eigene Elemente").'
                        : `Listen-Sicherheit bei ${ilsWrong} Liste(n) falsch — ${ilsFixed} davon repariert${ilsFixed < ilsWrong ? ', Rest bitte manuell prüfen!' : '.'}`);
                      parts.push(`${proxyFound} Anmeldungen durch Dritte gefunden, ${authorFixed} Zugriffe repariert${authorFailed > 0 ? `, ${authorFailed} nicht möglich (z.B. externe Personen)` : ''}.`);
                      if (errorEvents.length > 0) parts.push(`Fehler bei: ${errorEvents.join(', ')}`);
                    } else {
                      parts.push(`${listsChecked} lists checked.`);
                      parts.push(ilsWrong === 0
                        ? 'List security correct everywhere („own items only").'
                        : `List security wrong on ${ilsWrong} list(s) — ${ilsFixed} repaired${ilsFixed < ilsWrong ? ', please check the rest manually!' : '.'}`);
                      parts.push(`${proxyFound} third-party registrations found, ${authorFixed} access repaired${authorFailed > 0 ? `, ${authorFailed} not possible (e.g. external people)` : ''}.`);
                      if (errorEvents.length > 0) parts.push(`Errors on: ${errorEvents.join(', ')}`);
                    }
                    if (queueIlsLine) parts.push(queueIlsLine);
                    setRepairAccessResult(parts.join(' '));
                    // v20.7: Summary im Fortschritts-Modal anzeigen.
                    setAccessFixModal(prev => prev ? { ...prev, running: false, summary: parts } : prev);
                  } catch {
                    const err = isDe ? 'Fehler bei der Zugriffs-Prüfung.' : 'Error during the access check.';
                    setRepairAccessResult(err);
                    setAccessFixModal(prev => prev ? { ...prev, running: false, summary: [err] } : prev);
                  } finally {
                    setIsRepairingAccess(false);
                  }
                }}
              />
            )}

            {/* 8b. Organizer-Mails reparieren (alle Events) — Admin only.
                Findet Events mit Längen-Mismatch zwischen organizers (Names) und
                organizerEmails — typisch nach Legacy-Korruption aus v10.0–v10.2-
                Closure-Bug. Versucht via Graph-Search die fehlenden Emails per
                Lastname-Match nachzufüllen. Persistiert das gefixte Pair-Mapping
                via updateEvent. Bricht NICHT bei einzelnen Fehlern ab — ein Event
                mit unauflösbarem Namen wird übersprungen, der Rest läuft weiter.
                Operiert über ALLE adminEvents (nicht nur das gerade ausgewählte). */}
            {isAdmin && (
              <ActionTile
                icon={<Wrench size={18} />}
                category="maintenance"
                title={isRepairingOrganizers ? (isDe ? 'Reparatur läuft…' : 'Repair running…') : (isDe ? 'Organizer-Mails reparieren (alle Events)' : 'Repair organizer emails (all events)')}
                desc={isDe
                  ? 'Prüft alle Events auf Organizer-Namen ohne E-Mail (Altlast früherer App-Versionen), sucht die fehlenden Adressen über den Nachnamen im Tenant und speichert die vervollständigten Paare. Personen ohne eindeutig auffindbare Adresse fallen aus dem Organizer-Feld und werden namentlich gemeldet — die trägst du im Wizard neu ein.'
                  : 'Checks all events for organizer names without an email (legacy data from earlier app versions), looks up the missing addresses by last name in the tenant and saves the completed pairs. People without an unambiguous address are dropped from the organizer field and reported by name — you re-add them in the wizard.'}
                badge="admin"
                busy={isRepairingOrganizers}
                result={repairOrganizersResult}
                resultIsError={!!repairOrganizersResult && (repairOrganizersResult.indexOf('Fehler') >= 0 || repairOrganizersResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef) return;
                  if (!(await confirmDialog(isDe
                    ? `Organizer-Mails über ALLE ${adminEvents.length} Events reparieren? Dauert je nach Anzahl ca. 1–2 Minuten.`
                    : `Repair organizer emails across ALL ${adminEvents.length} events? Depending on the count this takes about 1–2 minutes.`, { confirmLabel: isDe ? 'Reparieren' : 'Repair' }))) return;
                  setIsRepairingOrganizers(true);
                  setRepairOrganizersResult(null);
                  let scanned = 0;
                  let mismatched = 0;
                  let eventsUpdated = 0;
                  let orgsRecovered = 0;
                  const unresolvedNames: string[] = [];
                  // v30.67: Namen, die beim Schreiben aus dem Organizer-Feld
                  // fielen (s. finalPairs unten) — getrennt gemeldet, weil der
                  // Admin sie im Wizard neu hinzufügen muss.
                  const droppedNames: string[] = [];
                  let namesFromEmail = 0;
                  try {
                    for (const ev of adminEvents) {
                      scanned++;
                      const names = (ev.organizers || []).slice();
                      const emails = (ev.organizerEmails || []).slice();
                      // Pad das kürzere Array auf max() — nichts geht verloren
                      const max = Math.max(names.length, emails.length);
                      while (names.length < max) names.push('');
                      while (emails.length < max) emails.push('');
                      // Mismatch erkannt? Mindestens ein Name ohne Email oder eine Email ohne Name
                      const hasMismatch = names.some((n, i) => (n || '').trim() && !((emails[i] || '').trim()))
                        || emails.some((e, i) => (e || '').trim() && !((names[i] || '').trim()));
                      if (!hasMismatch) continue;
                      mismatched++;
                      // Für jeden Slot mit Name aber ohne Email: Graph-Search nach Lastname
                      // pro Slot, EINS-zu-EINS-Match wenn Local-Part den Lastname enthält.
                      let recoveredHere = 0;
                      const unresolvedHere: string[] = [];
                      const fixedNames = names.slice();
                      const fixedEmails = emails.slice();
                      for (let i = 0; i < max; i++) {
                        const name = (fixedNames[i] || '').trim();
                        const email = (fixedEmails[i] || '').trim();
                        if (!name || email) continue;
                        // Lastname extrahieren — egal ob "Lastname, Firstname" oder
                        // "Firstname Lastname", als Suchquery für Graph nehmen wir
                        // den ganzen Namen (Graph ist tolerant).
                        try {
                          // Lastname als Suchterm — Graph-Search ist tolerant für
                          // 'Lastname' als Query und liefert eindeutigere Ergebnisse
                          // als die kombinierte Form 'Lastname, Firstname'.
                          const queryRaw = name.indexOf(',') >= 0 ? name.split(',')[0].trim() : name;
                          const hits = await searchUsers(queryRaw);
                          // Lastname-Substring-Match: filtere die Hits auf Personen,
                          // deren Email-Local-Part den Lastname enthält. Damit greifen
                          // wir den richtigen Eintrag auch bei Häufigkeitsnamen.
                          const lastname = queryRaw.toLowerCase().split(/\s+/).filter(t => t.length >= 3).pop() || '';
                          const matched = lastname
                            ? hits.filter(h => ((h.email || '').toLowerCase().split('@')[0]).indexOf(lastname) >= 0)
                            : hits;
                          if (matched.length === 1 && matched[0].email) {
                            fixedEmails[i] = matched[0].email;
                            recoveredHere++;
                            orgsRecovered++;
                          } else if (matched.length === 0 && hits.length === 1 && hits[0].email) {
                            // Kein Lastname-Match aber genau 1 Treffer überhaupt — übernehmen
                            fixedEmails[i] = hits[0].email;
                            recoveredHere++;
                            orgsRecovered++;
                          } else {
                            // Mehrdeutig oder nichts gefunden — leer lassen
                            unresolvedHere.push(`${name} (Event ${ev.eventNumber})`);
                          }
                        } catch {
                          unresolvedHere.push(`${name} (Event ${ev.eventNumber})`);
                        }
                      }
                      // Nichts wiederhergestellt? Skip Update — Storage ist eh schon
                      // im aktuellen Zustand, Pad allein bringt keinen Mehrwert
                      // (bei Save aus Wizard heilt sich das ohnehin).
                      if (recoveredHere === 0) { unresolvedNames.push(...unresolvedHere); continue; }
                      // v30.67: Slots „E-Mail ohne Namen" vor dem Schreiben mit dem
                      // Anzeigenamen füllen (ersatzweise die Adresse selbst) — die
                      // E-Mail ist berechtigungsrelevant und darf nicht wegfallen.
                      for (let i = 0; i < max; i++) {
                        const e = (fixedEmails[i] || '').trim();
                        if (!e || (fixedNames[i] || '').trim()) continue;
                        let resolved = '';
                        try { resolved = eventServiceRef ? await eventServiceRef.displayNameForEmail(e) : ''; } catch { /* Adresse als Name */ }
                        fixedNames[i] = resolved || e;
                        namesFromEmail++;
                      }
                      // v30.67: NUR vollständige Paare schreiben — wie
                      // `sanitizeOrganizerPairs` im Wizard. Organizer und
                      // OrganizerEmail sind zwei POSITIONSGEBUNDENE Strings, aber die
                      // Lesekante (`eventMapping.ts`, `.filter(s => s)`) verwirft leere
                      // Segmente: 'anna;;carla' wird zu [anna, carla], und ab dann trägt
                      // „Bernd" Carlas Adresse; `remove(idx)` auf Bernds Chip löscht
                      // Carlas E-Mail. Die Reparatur erzeugte damit genau die
                      // Verschiebung, die sie beheben soll — und der zweite Lauf
                      // zementierte sie. Unauflösbare Namen fallen deshalb aus dem
                      // Feld und werden unten namentlich gemeldet.
                      const finalPairs = fixedNames.map((n, i) => ({ n: (n || '').trim(), e: (fixedEmails[i] || '').trim() }))
                        .filter(p => p.n && p.e);
                      droppedNames.push(...unresolvedHere);
                      const finalNames = finalPairs.map(p => p.n).join('; ');
                      const finalEmails = finalPairs.map(p => p.e).join(';');
                      try {
                        const ok = await updateEvent(ev.id, { 'Organizer': finalNames, 'OrganizerEmail': finalEmails });
                        if (ok) eventsUpdated++;
                      } catch {
                        // Update fehlgeschlagen — counts trotzdem belassen, einfach
                        // skip dieses Event.
                      }
                    }
                    const lines = isDe
                      ? [`Gescannt: ${scanned}`, `Mit Mismatch: ${mismatched}`, `Aktualisiert: ${eventsUpdated}`, `Emails wiederhergestellt: ${orgsRecovered}`]
                      : [`Scanned: ${scanned}`, `With mismatch: ${mismatched}`, `Updated: ${eventsUpdated}`, `Emails recovered: ${orgsRecovered}`];
                    if (namesFromEmail > 0) {
                      lines.push(isDe ? `Namen zu E-Mails ergänzt: ${namesFromEmail}` : `Names added for emails: ${namesFromEmail}`);
                    }
                    if (unresolvedNames.length > 0) {
                      lines.push(isDe
                        ? `E-Mail manuell nachziehen (${unresolvedNames.length}): ${unresolvedNames.slice(0, 5).join(', ')}${unresolvedNames.length > 5 ? '…' : ''}`
                        : `Add email manually (${unresolvedNames.length}): ${unresolvedNames.slice(0, 5).join(', ')}${unresolvedNames.length > 5 ? '…' : ''}`);
                    }
                    if (droppedNames.length > 0) {
                      lines.push(isDe
                        ? `Ohne E-Mail nicht speicherbar — aus dem Organizer-Feld entfernt, im Wizard neu hinzufügen (${droppedNames.length}): ${droppedNames.slice(0, 5).join(', ')}${droppedNames.length > 5 ? '…' : ''}`
                        : `Not storable without email — removed from the organizer field, re-add in the wizard (${droppedNames.length}): ${droppedNames.slice(0, 5).join(', ')}${droppedNames.length > 5 ? '…' : ''}`);
                    }
                    setRepairOrganizersResult(lines.join(' · '));
                  } catch (err) {
                    setRepairOrganizersResult(isDe ? `Fehler: ${err instanceof Error ? err.message : String(err)}` : `Error: ${err instanceof Error ? err.message : String(err)}`);
                  }
                  setIsRepairingOrganizers(false);
                }}
              />
            )}

            {/* v28.65: Login-Tokens in Namen reparieren.
                SharePoint stempelt in seine versteckte „User Information List"
                bei manchen Personen das Claims-Login („0#.f|membership|
                user@deloitte.de") statt des Anzeigenamens — und korrigiert das
                nie rückwirkend. Bis v28.64 landete dieser Wert in
                Teilnehmerzeilen, Organizer-Feldern und der Rollenliste. Diese
                Aktion sucht die betroffenen Einträge und zieht die Namen aus
                dem Benutzerprofil nach. Neue Einträge sind seit v28.64 sauber;
                das hier räumt den Bestand auf. */}
            {isAdmin && (
              <ActionTile
                icon={<Wrench size={18} />}
                category="maintenance"
                title={isRepairingNames
                  ? (isDe ? 'Namen werden repariert…' : 'Repairing names…')
                  : (isDe ? 'Login-Tokens in Namen reparieren (alle Events)' : 'Repair login tokens in names (all events)')}
                desc={isDe
                  ? 'Sucht Einträge, bei denen statt des Namens ein technisches SharePoint-Login steht („0#.f|membership|user@deloitte.de") — in den Teilnehmerlisten aller Events, in den Organizer-Namen der Events und in der Rollenverwaltung. Die Namen werden aus dem Benutzerprofil nachgezogen; ist eine Person dort nicht auflösbar, steht statt des Tokens wenigstens die E-Mail. Seit v28.64 entstehen keine neuen Fälle mehr — das hier räumt den Bestand auf.'
                  : 'Finds entries where a technical SharePoint login („0#.f|membership|user@deloitte.de") is stored instead of the name — in the participant lists of all events, in the organizer names and in role management. Names are pulled from the user profile; if a person cannot be resolved, the email is used instead of the token. Since v28.64 no new cases occur — this cleans up the existing data.'}
                badge="admin"
                busy={isRepairingNames}
                result={repairNamesResult}
                resultIsError={!!repairNamesResult && (repairNamesResult.indexOf('Fehler') >= 0 || repairNamesResult.indexOf('Error') >= 0)}
                onClick={async () => {
                  if (!eventServiceRef) return;
                  if (!(await confirmDialog(isDe
                    ? `Namen über ALLE ${adminEvents.length} Events, die Organizer-Felder und die Rollenverwaltung prüfen und reparieren?\n\nGeändert werden ausschließlich Einträge, in denen ein Login-Token statt eines Namens steht. Je nach Anzahl der Teilnehmer dauert das einige Minuten.`
                    : `Check and repair names across ALL ${adminEvents.length} events, the organizer fields and role management?\n\nOnly entries containing a login token instead of a name are changed. Depending on the number of attendees this takes a few minutes.`,
                    { confirmLabel: isDe ? 'Reparieren' : 'Repair' }))) return;
                  setIsRepairingNames(true);
                  setRepairNamesResult(null);
                  setNameFixModal({ running: true, step: isDe ? 'Teilnehmerlisten' : 'Participant lists', evIdx: 0, evTotal: adminEvents.length, summary: null });
                  let regHits = 0; let regFixed = 0; let regFailed = 0;
                  let orgFixed = 0; let orgEvents = 0;
                  // v30.67: Slots ohne E-Mail, die beim Schreiben wegfielen (s.u.).
                  let orgDropped = 0;
                  let roleFixed = 0; let roleHits = 0;
                  try {
                    // 1. Teilnehmerlisten aller Events (Sub-Events haben eigene Listen
                    //    und stehen selbst in adminEvents — daher genügt ein Durchlauf).
                    for (let i = 0; i < adminEvents.length; i++) {
                      const ev = adminEvents[i];
                      setNameFixModal({ running: true, step: `${isDe ? 'Teilnehmerliste' : 'Participant list'}: ${ev.title}`, evIdx: i + 1, evTotal: adminEvents.length, summary: null });
                      if (!ev.subsiteUrl) continue;
                      try {
                        const r = await eventServiceRef.repairClaimNamesInRegistrations(ev.subsiteUrl);
                        regHits += r.hits; regFixed += r.fixed; regFailed += r.failed;
                      } catch { /* Liste nicht lesbar — weiter */ }
                    }
                    // 2. Organizer-Namen der Events.
                    setNameFixModal({ running: true, step: isDe ? 'Organizer-Namen' : 'Organizer names', evIdx: adminEvents.length, evTotal: adminEvents.length, summary: null });
                    const looksLikeClaim = (x: string): boolean => /\|membership\b|^i:0[#|]|^c:0|0#\.[a-z]\||^\d+#\./i.test((x || '').trim());
                    const mailFromClaim = (x: string): string => {
                      const m = (x || '').match(/\|([^|]+@[^|\s]+)\s*$/);
                      return m ? m[1].trim().toLowerCase() : '';
                    };
                    for (const ev of adminEvents) {
                      const names = (ev.organizers || []).slice();
                      const emails = (ev.organizerEmails || []).slice();
                      if (!names.some(n => looksLikeClaim(n))) continue;
                      let changed = 0;
                      for (let i = 0; i < names.length; i++) {
                        if (!looksLikeClaim(names[i])) continue;
                        const mail = (emails[i] || '').trim() || mailFromClaim(names[i]);
                        const resolved = mail ? await eventServiceRef.displayNameForEmail(mail) : '';
                        names[i] = resolved || mail || names[i];
                        if (!emails[i] && mail) emails[i] = mail;
                        changed++; orgFixed++;
                      }
                      if (changed === 0) continue;
                      // v30.67: nur vollständige Paare schreiben (s. „Organizer-Mails
                      // reparieren"). `emails` war hier zudem ein SPARSE Array, wenn es
                      // kürzer als `names` war — `join(';')` macht aus den Löchern leere
                      // Segmente, die Lesekante verwirft sie, alle Adressen dahinter
                      // rutschen um einen Slot nach vorn.
                      // v30.67 (Review): über max(names, emails) laufen — wie im
                      // Nachbar-Tile „Organizer-Mails reparieren". `names.map` lief nur
                      // über die Namen: Ein Event mit MEHR Adressen als Namen verlor
                      // die überzähligen Adressen stumm, und die Adresse ist das, was
                      // berechtigt (Rechteverlust ohne Meldung). Ein Slot „E-Mail ohne
                      // Name" bekommt deshalb den Anzeigenamen (ersatzweise die
                      // Adresse); verworfen und gezählt wird nur „Name ohne E-Mail".
                      const slots = Math.max(names.length, emails.length);
                      const pairs: Array<{ n: string; e: string }> = [];
                      for (let i = 0; i < slots; i++) {
                        let n = (names[i] || '').trim();
                        const e = (emails[i] || '').trim();
                        if (e && !n) {
                          let resolved = '';
                          try { resolved = await eventServiceRef.displayNameForEmail(e); } catch { /* Adresse als Name */ }
                          n = resolved || e;
                        }
                        pairs.push({ n, e });
                      }
                      const complete = pairs.filter(p => p.n && p.e);
                      orgDropped += pairs.length - complete.length;
                      try {
                        const ok = await updateEvent(ev.id, { 'Organizer': complete.map(p => p.n).join('; '), 'OrganizerEmail': complete.map(p => p.e).join(';') });
                        if (ok) orgEvents++;
                      } catch { /* Event überspringen */ }
                    }
                    // 3. Rollenverwaltung.
                    setNameFixModal({ running: true, step: isDe ? 'Rollenverwaltung' : 'Role management', evIdx: adminEvents.length, evTotal: adminEvents.length, summary: null });
                    try {
                      const rr = spServiceRef ? await spServiceRef.repairClaimNamesInRoles() : { scanned: 0, hits: 0, fixed: 0, failed: 0 };
                      roleHits = rr.hits; roleFixed = rr.fixed;
                    } catch { /* Rollenliste nicht schreibbar */ }

                    const lines = isDe
                      ? [
                        `Teilnehmerzeilen: ${regHits} betroffen, ${regFixed} repariert${regFailed > 0 ? `, ${regFailed} fehlgeschlagen` : ''}`,
                        `Organizer-Namen: ${orgFixed} in ${orgEvents} Event(s)${orgDropped > 0 ? `, ${orgDropped} Slot(s) ohne E-Mail nicht übernommen (im Wizard nachtragen)` : ''}`,
                        `Rollenverwaltung: ${roleHits} betroffen, ${roleFixed} repariert`,
                      ]
                      : [
                        `Attendee rows: ${regHits} affected, ${regFixed} repaired${regFailed > 0 ? `, ${regFailed} failed` : ''}`,
                        `Organizer names: ${orgFixed} in ${orgEvents} event(s)${orgDropped > 0 ? `, ${orgDropped} slot(s) without email not written (add in the wizard)` : ''}`,
                        `Role management: ${roleHits} affected, ${roleFixed} repaired`,
                      ];
                    setRepairNamesResult(lines.join(' · '));
                    setNameFixModal({ running: false, step: '', evIdx: 0, evTotal: 0, summary: lines });
                  } catch (err) {
                    setRepairNamesResult(isDe ? `Fehler: ${err instanceof Error ? err.message : String(err)}` : `Error: ${err instanceof Error ? err.message : String(err)}`);
                    setNameFixModal(null);
                  }
                  setIsRepairingNames(false);
                }}
              />
            )}
          </div>
        </ActionsCollapsibleCard>
  );
};

