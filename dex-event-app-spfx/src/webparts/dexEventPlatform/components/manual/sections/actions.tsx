import * as React from 'react';
import { ManualSection } from '../types';

/**
 * v22.53: Handbuch-Sektion, die ALLE Aktionen aus dem Aktionen-Dropdown des
 * Organizer Centers erklärt — gruppiert nach denselben Kategorien wie das
 * Dropdown (Event / Teilnehmer / E-Mails / Check-in / Wartung & Reparatur).
 * Klartext, kein Tech-Jargon (Tooltip-Regel).
 */
export function actionsSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'actions',
    title: isDe ? 'Aktionen im Organizer Center' : 'Actions in the Organizer Center',
    category: 'organizer',
    description: isDe
      ? 'Alle Aktionen aus dem Aktionen-Dropdown erklärt — Teilnehmerliste exportieren, QR-Codes, Mails, Check-in und Reparatur-Werkzeuge.'
      : 'Every action from the actions dropdown explained — export the attendee list, QR codes, emails, check-in and repair tools.',
    visibleFor: ['Organizer', 'Admin'],
    keywords: isDe
      ? 'Aktionen Aktion auswählen Dropdown Menü Werkzeuge Organizer Center Admin Center Kategorien Suchfeld Excel-Export Teilnehmerliste xlsx herunterladen Zielgruppe Matrix Sub-Event-Blätter IDs neu vergeben Nummern renummerieren Von Warteliste nachrücken Überbuchung prüfen Outlook-Absagen prüfen Massenmail E-Mail versenden Einladungsmail E-Mails kopieren Check-in starten QR-Codes versenden Self-Check-in Event bearbeiten In SharePoint öffnen Deep-Link kopieren Audit-Log Änderungsprotokoll Spalten fixen Zugriff reparieren Counter zurücksetzen Custom-Fields reparieren Wartung Reparatur'
      : 'actions pick an action dropdown menu tools organizer center admin center categories search box Excel export attendee list xlsx download target group matrix sub-event sheets reassign IDs renumber promote from waitlist check overbooking check Outlook declines mass email send email invitation email copy emails start check-in send QR codes self check-in edit event open in SharePoint copy deep link audit log change history fix columns repair access reset counter migrate B2Run repair custom fields maintenance repair',
    perspectives: [
      {
        perspective: 'organizer',
        intro: isDe
          ? 'Im Organizer Center sitzt unter den Event-Details der grüne Knopf „Aktion auswählen". Er öffnet ein Menü mit allen Werkzeugen für dein Event — nach Kategorien gruppiert und mit einem Suchfeld oben. Tippe z.B. „Excel" oder „SharePoint", um direkt zum Export zu springen. Hier ist jede Aktion erklärt.'
          : 'In the Organizer Center, below the event details, sits the green "Pick an action" button. It opens a menu with every tool for your event — grouped by category with a search box at the top. Type e.g. "Excel" or "SharePoint" to jump straight to the export. Here is what each action does.',
        steps: [
          {
            number: 1,
            title: isDe ? 'Kategorie „Teilnehmer" — inkl. Excel-Export' : 'Category "Participants" — incl. Excel export',
            description: (
              <>
                {isDe ? (
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                    <li><strong>Teilnehmerliste exportieren (Excel)</strong> — lädt die Teilnehmer als Excel-Datei (.xlsx) direkt auf deinen Rechner herunter. Vor dem Export wählst du die <strong>Zielgruppe</strong>: nur Angemeldete, inklusive Warteliste oder <strong>„Alles inkl. Abmeldungen“</strong> (der Status steht dann pro Zeile in einer eigenen Spalte). Bei Events mit Sub-Events (Klammer) kannst du zusätzlich entscheiden, ob du <strong>eine konsolidierte Matrix</strong> (eine Zeile pro Person über alle Sub-Events) und/oder <strong>einzelne Sub-Event-Blätter</strong> bekommst.</li>
                    <li><strong>IDs neu vergeben</strong> — vergibt die TeilnehmerIDs sauber neu (Angemeldete 1…N, Warteliste danach). Normalerweise passiert das automatisch — der Knopf ist für den Fall, dass du sofort eine lückenlose Reihenfolge brauchst.</li>
                    <li><strong>Von Warteliste nachrücken</strong> — holt die erste Person von der Warteliste manuell in den Aktiv-Bereich (z.B. wenn ein Platz frei wurde und du nicht auf die Automatik warten willst).</li>
                    <li><strong>Überbuchung prüfen</strong> — findet Personen, die über der Kapazität angemeldet sind, und markiert sie zur Prüfung. Du entscheidest dann pro Person, ob sie auf die Warteliste wandert oder den Platz behält.</li>
                    <li><strong>Outlook-Absagen prüfen</strong> — gleicht ab, wer den Kalender-Termin abgesagt hat, ohne sich in der App abzumelden.</li>
                  </ul>
                ) : (
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                    <li><strong>Export attendee list (Excel)</strong> — downloads the attendees as an Excel file (.xlsx) straight to your computer. Before exporting you pick the <strong>target group</strong>: registered only, including the waitlist, or <strong>“All incl. cancellations”</strong> (the status then appears per row in its own column). For events with sub-events (bracket) you can additionally choose a <strong>consolidated matrix</strong> (one row per person across all sub-events) and/or <strong>individual sub-event sheets</strong>.</li>
                    <li><strong>Reassign attendee IDs</strong> — renumbers the attendee IDs cleanly (registered 1…N, waitlist after). This usually happens automatically — the button is for when you need a gap-free order right away.</li>
                    <li><strong>Promote from waitlist</strong> — manually pulls the first waitlisted person into the active area (e.g. when a seat opened and you don’t want to wait for the automation).</li>
                    <li><strong>Check overbooking</strong> — finds people registered above capacity and flags them for review. You then decide per person whether they move to the waitlist or keep the seat.</li>
                    <li><strong>Check Outlook declines</strong> — reconciles who declined the calendar invite without cancelling in the app.</li>
                  </ul>
                )}
              </>
            ),
            tip: isDe
              ? 'Der Excel-Export läuft komplett lokal in deinem Browser — die Datei landet in deinem Download-Ordner. Es werden keine Teilnehmerdaten an Dritte verschickt.'
              : 'The Excel export runs entirely locally in your browser — the file lands in your download folder. No attendee data is sent to third parties.',
          },
          {
            number: 2,
            title: isDe ? 'Kategorie „E-Mails"' : 'Category "Emails"',
            description: (
              <>
                {isDe ? (
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                    <li><strong>E-Mail versenden an Teilnehmergruppen</strong> — schreibt eine Massenmail an ausgewählte Gruppen (z.B. alle Angemeldeten oder nur die Warteliste) im Deloitte-Layout, mit Rich-Text-Editor.</li>
                    <li><strong>Einladungsmail</strong> — verschickt eine Mail mit dem Anmelde-Link — an dich selbst zum Testen oder an einen Mailverteiler.</li>
                    <li><strong>E-Mails kopieren</strong> — legt alle aktiven Teilnehmer-Adressen semikolon-getrennt in die Zwischenablage, direkt einfügbar in Outlook.</li>
                  </ul>
                ) : (
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                    <li><strong>Send email to participant groups</strong> — sends a mass email to selected groups (e.g. all registered or just the waitlist) in the Deloitte layout, with a rich-text editor.</li>
                    <li><strong>Invitation email</strong> — sends an email with the registration link — to yourself for testing or to a mailing list.</li>
                    <li><strong>Copy emails</strong> — puts all active attendee addresses semicolon-separated into your clipboard, ready to paste into Outlook.</li>
                  </ul>
                )}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Kategorie „Check-in"' : 'Category "Check-in"',
            description: (
              <>
                {isDe ? (
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                    <li><strong>Check-in starten</strong> — öffnet die Check-in-Seite, um Teilnehmer am Event-Tag einzuchecken (per Scan oder Namenssuche).</li>
                    <li><strong>QR-Codes versenden</strong> — schickt jedem aktiven Teilnehmer seinen persönlichen Check-in-QR per Mail. Ab dem ersten Versand bekommen auch neue Anmeldungen ihren Code automatisch.</li>
                    <li><strong>Self-Check-in einstellen</strong> — öffnet das Fenster für den Self-Check-in: großer QR, druckbares PDF, rotierende Live-Anzeige und das Check-in-Zeitfenster (Von/Bis).</li>
                  </ul>
                ) : (
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                    <li><strong>Start check-in</strong> — opens the check-in page to check attendees in on event day (by scan or name search).</li>
                    <li><strong>Send QR codes</strong> — emails each active attendee their personal check-in QR. From the first send on, new registrations also get their code automatically.</li>
                    <li><strong>Set up self check-in</strong> — opens the self check-in panel: large QR, printable PDF, rotating live display and the check-in time window (from/to).</li>
                  </ul>
                )}
              </>
            ),
          },
          {
            number: 4,
            title: isDe ? 'Kategorie „Event"' : 'Category "Event"',
            description: (
              <>
                {isDe ? (
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                    <li><strong>Event bearbeiten</strong> — öffnet den Wizard, um alle Einstellungen des Events zu ändern.</li>
                    <li><strong>Deep-Link kopieren</strong> — kopiert den direkten Link zur Anmeldeseite des Events, z.B. für eine eigene Einladung.</li>
                    <li><strong>Audit-Log / Änderungsprotokoll</strong> — öffnet das Protokoll aller Änderungen an diesem Event (wer hat wann was geändert), bereits auf das Event vorgefiltert.</li>
                  </ul>
                ) : (
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                    <li><strong>Edit event</strong> — opens the wizard to change all settings of the event.</li>
                    <li><strong>Copy deep link</strong> — copies the direct link to the event’s registration page, e.g. for your own invitation.</li>
                    <li><strong>Audit log / change history</strong> — opens the log of all changes to this event (who changed what and when), pre-filtered to the event.</li>
                  </ul>
                )}
              </>
            ),
          },
          {
            number: 5,
            title: isDe ? 'Kategorie „Wartung & Reparatur" (meist Admin)' : 'Category "Maintenance & repair" (mostly Admin)',
            description: (
              <>
                {isDe ? (
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                    <li><strong>Spalten fixen</strong> — legt fehlende Spalten in der Teilnehmerliste an und bringt sie in die richtige Reihenfolge (nach Schema-Änderungen).</li>
                    <li><strong>Fremd-Anmeldungen: Zugriff reparieren</strong> — sorgt dafür, dass jemand, der von einer anderen Person angemeldet wurde, seine Anmeldung in „Meine Events“ sieht und sich selbst abmelden kann.</li>
                    <li><strong>Counter zurücksetzen</strong> — setzt den internen Platz-Zähler neu auf den echten Bestand (Sonderfall nach Störungen).</li>
                    <li><strong>Custom-Fields / Felder reparieren</strong> — holt verloren geglaubte eigene Abfragen zurück bzw. repariert beschädigte Feld-Definitionen.</li>
                  </ul>
                ) : (
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                    <li><strong>Fix columns</strong> — creates missing columns in the attendee list and puts them in the right order (after schema changes).</li>
                    <li><strong>Repair proxy-registration access</strong> — makes sure someone who was registered by another person sees their registration in “My Events” and can cancel it themselves.</li>
                    <li><strong>Reset counter</strong> — resets the internal seat counter to the real headcount (edge case after disruptions).</li>
                    <li><strong>Restore / repair custom fields</strong> — brings back custom questions believed lost or repairs broken field definitions.</li>
                  </ul>
                )}
              </>
            ),
            warning: isDe
              ? 'Reparatur-Aktionen sind für Sonderfälle gedacht. Im Normalbetrieb brauchst du sie nicht — frag im Zweifel kurz nach, bevor du sie ausführst.'
              : 'Repair actions are for edge cases. In normal operation you won\'t need them — when in doubt, check before running them.',
          },
        ],
      },
    ],
  };
}
