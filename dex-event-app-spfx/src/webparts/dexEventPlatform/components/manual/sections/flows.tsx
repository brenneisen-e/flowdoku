import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';

export function flowsSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'flows',
    title: isDe ? 'Power Automate Flows' : 'Power Automate Flows',
    category: 'architecture',
    description: isDe
      ? 'Die sechs Hintergrund-Flows, die E-Mails und Outlook-Termine steuern.'
      : 'The six background flows that drive emails and Outlook appointments.',
    visibleFor: ['Admin'],
    keywords: isDe
      ? 'Power Automate Flows Hintergrund-Flows Queue DEX_Emails DEX_Outlook DEX_SEND_MAIL DEX_CreateOutlookEvent DEX_Outlook_Einladungen DEX_OutlookDeclineHandler DEX_OutlookForwardHandler DEX_IDReorder_TeilnehmerIDs Outlook-Termin Kalender ORB_URL flow-jsons Monitoring no_reply.events'
      : 'Power Automate flows background flows queue DEX_Emails DEX_Outlook DEX_SEND_MAIL DEX_CreateOutlookEvent DEX_Outlook_Einladungen DEX_OutlookDeclineHandler DEX_OutlookForwardHandler DEX_IDReorder_TeilnehmerIDs outlook appointment calendar ORB_URL flow-jsons monitoring no_reply.events',
    perspectives: [
      {
        perspective: 'admin',
        intro: (
          <>
            {isDe
              ? 'Die App speichert E-Mail- und Outlook-Intents in SharePoint-Queues (DEX_Emails, DEX_Outlook). Sechs Power-Automate-Flows arbeiten diese ab. Volle Flow-JSON-Definitionen findest du in docs/flow-jsons.md (einzige Quelle der Wahrheit).'
              : 'The app writes email and Outlook intents into SharePoint queues (DEX_Emails, DEX_Outlook). Six Power Automate flows drain them. Full flow definitions live in docs/flow-jsons.md (single source of truth).'}
          </>
        ),
        steps: [
          {
            number: 1,
            title: 'DEX_SEND_MAIL',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Trigger: neuer Eintrag in DEX_Emails. Der Body kommt aus der App bereits komplett gewickelt (Deloitte-Template, Header, Footer, Logo inline) mit nur einem offenen Platzhalter:'
                    : 'Trigger: new item in DEX_Emails. The body arrives from the app already fully wrapped (Deloitte template, header, footer, logo inlined) with one remaining placeholder:'}
                </p>
                <ul style={{ paddingLeft: 18, margin: '0 0 6px 0', lineHeight: 1.7 }}>
                  <li><code>{'{{ORB_URL}}'}</code> — {isDe ? 'wird durch Event-spezifisches EmailImageBase64 ersetzt (oder DefaultImageBase64 aus _Config wenn das Event kein Bild hat).' : 'replaced with event-specific EmailImageBase64 (or DefaultImageBase64 from _Config if the event has no image).'}</li>
                </ul>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Versand über die Shared Mailbox no_reply.events@deloitte.de. Status danach auf "Sent" oder "Failed".'
                    : 'Sent via the shared mailbox no_reply.events@deloitte.de. Status afterwards "Sent" or "Failed".'}
                </p>
              </>
            ),
          },
          {
            number: 2,
            title: 'DEX_CreateOutlookEvent',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Trigger: neues DEX_Events-Item. Erzeugt einmalig den Outlook-Kalendereintrag im Postfach des Hauptorganisators und schreibt die iCalUId in DEX_Events.CalendarLink zurück.'
                    : 'Trigger: new DEX_Events item. Creates the initial Outlook calendar entry in the primary organizer\'s mailbox and writes back iCalUId to DEX_Events.CalendarLink.'}
                </p>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Wichtig: das Body-Feld in der "Create event (V4)"-Aktion DARF KEINEN <p class="editor-paragraph">…</p>-Wrapper enthalten — sonst wird das Deloitte-HTML-Layout in einen einzigen <p>-Tag gepresst und Outlook strippt fast alles weg. Erwartet wird der reine OutlookBody mit Platzhalter-Replace für LOGO_URL und ORB_URL.'
                    : 'Important: the body field in "Create event (V4)" must NOT have a <p class="editor-paragraph">…</p> wrapper — otherwise the Deloitte HTML layout collapses into a single <p> tag and Outlook strips most of it. Expected: the raw OutlookBody with placeholder replacement for LOGO_URL and ORB_URL.'}
                </p>
              </>
            ),
            warning: isDe
              ? 'Vor 2026-04-17 enthielt der Flow den <p>-Wrapper-Bug. Wenn du Termine vor diesem Datum siehst die "fast leer" sind, ist das die Ursache. Bei einem Body-Update aus der App reicht das Klicken von "Outlook event aktualisieren" um den korrekten Body nachzuziehen.'
              : 'Before 2026-04-17 the flow had the <p> wrapper bug. Calendar entries from before that date that look "almost empty" are caused by it. Click "Update Outlook event" in the app to backfill the correct body.',
          },
          {
            number: 3,
            title: 'DEX_Outlook_Einladungen',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Trigger: neuer Eintrag in DEX_Outlook (Queue). Vier ActionTypes:'
                    : 'Trigger: new item in DEX_Outlook (queue). Four ActionTypes:'}
                </p>
                <ul style={{ paddingLeft: 18, margin: '0 0 6px 0', lineHeight: 1.7 }}>
                  <li><strong>Einladen</strong> — {isDe ? 'fügt einen einzelnen Teilnehmer zum Outlook-Termin hinzu (PATCH attendees-Array).' : 'adds a single attendee to the Outlook event (PATCH attendees array).'}</li>
                  <li><strong>Ausladen</strong> — {isDe ? 'entfernt einen einzelnen Teilnehmer (PATCH gefiltertes attendees-Array).' : 'removes a single attendee (PATCH filtered attendees array).'}</li>
                  <li><strong>UpdateEvent</strong> — {isDe ? 'patcht Titel, Start, Ende UND Body (seit 2026-04-17). Der Body kommt aus DEX_Events.OutlookBody mit aufgelöstem {{ORB_URL}}. Vorher wurde der Body NICHT mitgeschickt → Aenderungen waren unsichtbar.' : 'patches subject, start, end AND body (since 2026-04-17). Body comes from DEX_Events.OutlookBody with {{ORB_URL}} resolved. Previously body was NOT sent → changes were invisible.'}</li>
                  <li><strong>DeleteEvent</strong> — {isDe ? 'löscht den Outlook-Termin per Graph DELETE. Item-CalendarLink wird direkt im Queue-Item mitgegeben (weil das DEX_Events-Item bei Lauf bereits weg ist).' : 'deletes the Outlook event via Graph DELETE. CalendarLink is included in the queue item (because the DEX_Events item is already gone when the flow runs).'}</li>
                </ul>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Concurrency: 1 (sequentielle Verarbeitung damit Race-Conditions auf demselben Outlook-Termin vermieden werden).'
                    : 'Concurrency: 1 (sequential processing to avoid race conditions on the same Outlook event).'}
                </p>
              </>
            ),
            tip: isDe
              ? 'UpdateEvent-Body-Expression nutzt json(string(...)) um HTML-Quotes sauber zu escapen. Wenn der PATCH 400 zurückgibt, ist meistens ein un-escaped Anführungszeichen im OutlookBody die Ursache.'
              : 'The UpdateEvent body expression uses json(string(...)) to escape HTML quotes. A 400 response on PATCH is usually caused by an unescaped quote in OutlookBody.',
          },
          {
            number: 4,
            title: 'DEX_OutlookDeclineHandler',
            description: (
              <>
                {isDe
                  ? 'Erkennt, wenn ein Teilnehmer die Outlook-Einladung ablehnt (auch ueber weitergeleitete Decline-Mails wie FW:/WG:). Statt automatisch in der App abzumelden, wird eine Reminder-Mail gequeued ("war das Absicht? Wenn ja, klicke hier zum Abmelden") — Nutzer bleibt in der App angemeldet bis er aktiv bestätigt.'
                  : 'Detects when an attendee declines the Outlook invite (including forwarded decline notifications like FW:/Re:). Instead of auto-cancelling in the app, queues a reminder email ("did you mean it? click here to confirm cancellation") — user stays registered until they actively confirm.'}
              </>
            ),
          },
          {
            number: 5,
            title: 'DEX_OutlookForwardHandler',
            description: (
              <>
                {isDe
                  ? 'Reagiert auf Meeting-Forward-Notifications. Wenn der weitergeleitete Empfänger NICHT in der App für das Event angemeldet ist, geht eine FYI-Mail an den Organizer ("X hat den Termin an Y weitergeleitet — Y ist aber nicht angemeldet"). Hilft Organizern bei der Sichtkontrolle.'
                  : 'Reacts to meeting forward notifications. If the forwarded-to person is NOT registered for the event in the app, sends an FYI to the organizer ("X forwarded the invite to Y — but Y is not registered"). Helps organizers spot-check.'}
              </>
            ),
          },
          {
            number: 6,
            title: 'DEX_IDReorder_TeilnehmerIDs',
            description: (
              <>
                {isDe
                  ? 'Wird durch den Admin-Button "IDs neu vergeben" getriggert. Nummeriert alle aktiven Teilnehmer eines Events sequentiell (lücken-frei) und rückt Wartelistler entsprechend nach. Letzter Schritt vor größeren Events um sauberere Teilnehmer-IDs zu haben.'
                  : 'Triggered by the admin "Renumber IDs" button. Renumbers active attendees sequentially (gap-free) and promotes waitlist members accordingly. Last step before bigger events to get clean attendee IDs.'}
              </>
            ),
          },
          {
            number: 7,
            title: isDe ? 'Body-/Logo-Replacement-Pipeline' : 'Body & logo replacement pipeline',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Die SPFx-App + Flows teilen sich die Verantwortung für Bilder im HTML-Body:'
                    : 'The SPFx app and the flows share responsibility for images in HTML body content:'}
                </p>
                <ul style={{ paddingLeft: 18, margin: '0 0 6px 0', lineHeight: 1.7 }}>
                  <li>{isDe ? <><strong>SPFx App</strong> baut den Body mit wrapTemplate() und inlined dabei das gecachte Deloitte-Logo direkt als Base64. <code>{'{{ORB_URL}}'}</code> bleibt als Platzhalter.</> : <><strong>SPFx app</strong> builds the body via wrapTemplate() and inlines the cached Deloitte logo directly as Base64. <code>{'{{ORB_URL}}'}</code> remains a placeholder.</>}</li>
                  <li>{isDe ? <><strong>Flow</strong> (DEX_SEND_MAIL / DEX_CreateOutlookEvent / DEX_Outlook_Einladungen) ersetzt nur noch <code>{'{{ORB_URL}}'}</code> mit dem event-spezifischen Bild (EmailImageBase64) oder dem globalen DefaultImageBase64.</> : <><strong>Flow</strong> (DEX_SEND_MAIL / DEX_CreateOutlookEvent / DEX_Outlook_Einladungen) only replaces <code>{'{{ORB_URL}}'}</code> with the event-specific image (EmailImageBase64) or the global DefaultImageBase64.</>}</li>
                </ul>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Vorteil: kein Doppel-Wrapping, kein Wrapper-Wissen im Flow nötig. Beim Edit eines Events wird der Body in der App via stripOutlookWrapper() unwrapped, der User bearbeitet nur den inneren Content, und beim Speichern wird er erneut sauber gewickelt (idempotent).'
                    : 'Advantage: no double-wrapping, no wrapper knowledge needed in the flow. On event edit, the body is unwrapped via stripOutlookWrapper(), the user edits only the inner content, and on save it gets re-wrapped cleanly (idempotent).'}
                </p>
              </>
            ),
          },
          {
            number: 8,
            title: isDe ? 'Monitoring' : 'Monitoring',
            description: (
              <>
                {isDe
                  ? 'Alle Flows laufen unter no_reply.events@deloitte.de. Fehler siehst du im Power-Automate-Admin-Center. Die Queue-Listen (DEX_Emails, DEX_Outlook) zeigen "Failed"-Einträge in der Status-Spalte — regelmäßig sichten und manuell auf "Pending" zurücksetzen falls ein Retry sinnvoll ist.'
                  : 'All flows run under no_reply.events@deloitte.de. Errors are visible in the Power Automate admin center. The queue lists (DEX_Emails, DEX_Outlook) show "Failed" entries in the Status column — review regularly and manually reset to "Pending" for retries.'}
              </>
            ),
            mockup: (
              <Callout variant="warning">
                {isDe
                  ? 'Bei Aenderungen an einem Flow: IMMER das aktuelle JSON in docs/flow-jsons.md pflegen. Das ist die einzige Quelle der Wahrheit für den aktuellen Stand. Power Automate selbst hat keine Versionierung im Sinne von Git.'
                  : 'When modifying a flow: ALWAYS update the current JSON in docs/flow-jsons.md. That file is the single source of truth — Power Automate has no Git-style versioning.'}
              </Callout>
            ),
          },
        ],
      },
    ],
  };
}
