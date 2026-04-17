import * as React from 'react';
import { ManualSection } from '../types';
import { DemoEmailPreview, Callout } from '../ManualMockups';

export function templatesSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'email-templates',
    title: isDe ? 'Globale E-Mail-Templates' : 'Global email templates',
    category: 'admin',
    description: isDe
      ? 'Die vier System-Mails (Anmeldung, Warteliste, Abmeldung, Nachrücken) zentral anpassen.'
      : 'Centrally customize the four system emails (Registration, Waitlist, Cancellation, Promotion).',
    visibleFor: ['Admin'],
    perspectives: [
      {
        perspective: 'admin',
        steps: [
          {
            number: 1,
            title: isDe ? 'Templates öffnen' : 'Open templates',
            description: (
              <>
                {isDe
                  ? 'In den Einstellungen → "E-Mail-Templates". Du siehst alle vier System-Typen in DE und EN, jeweils mit Betreff, Überschrift und Inhalt.'
                  : 'In Settings → "Email templates". You see all four system types in DE and EN, each with subject, heading and content.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Platzhalter nutzen' : 'Use placeholders',
            description: (
              <>
                {isDe
                  ? 'In Templates funktionieren Platzhalter, die automatisch ersetzt werden: {{Name}} = Teilnehmer-Vorname, {{EventTitle}} = Event-Titel, {{Organizer}} = Organizer-Namen, {{AppUrl}} = Link zur App, {{WaitlistPosition}} = Warteliste-Platz.'
                  : 'Templates support placeholders that are auto-replaced: {{Name}} = attendee first name, {{EventTitle}} = event title, {{Organizer}} = organizer names, {{AppUrl}} = link back to the app, {{WaitlistPosition}} = waitlist position.'}
              </>
            ),
            mockup: <DemoEmailPreview />,
            tip: isDe
              ? 'Jeder Organizer kann pro Event die Templates überschreiben — globale Templates sind der Default, Event-spezifische Overrides sind optional.'
              : 'Each organizer can override templates per event — globals are the default, per-event overrides are optional.',
          },
          {
            number: 3,
            title: isDe ? 'Änderungen testen' : 'Test your changes',
            description: (
              <>
                {isDe
                  ? 'Das Vorschau-Fenster zeigt dir den gerenderten HTML-Output mit Mock-Daten. Dort siehst du sofort, ob die Platzhalter korrekt aufgelöst werden.'
                  : 'The preview pane shows rendered HTML with mock data. You\'ll see immediately whether placeholders resolve correctly.'}
              </>
            ),
            warning: isDe
              ? 'Änderungen an globalen Templates gelten ab sofort für alle zukünftigen Mails — auch laufender Events. Teste mit einem Test-Event, bevor du produktive Änderungen machst.'
              : 'Global template changes take effect immediately for all future mails — including running events. Test with a test event before making production changes.',
          },
          {
            number: 4,
            title: isDe ? 'Logos & Default-Bilder' : 'Logos & default images',
            description: (
              <>
                {isDe
                  ? 'Im gleichen Bereich kannst du das Deloitte-Logo für den Mail-Header, das DEX-Logo und Default-Event-Bilder (falls Organizer keines hochladen) verwalten.'
                  : 'In the same area you manage the Deloitte logo for the mail header, the DEX logo, and default event images (used when organizers don\'t upload one).'}
              </>
            ),
            mockup: <Callout variant="info">{isDe ? 'Alle Bilder werden als Base64 in DEX_EmailTemplates gespeichert — keine externen URL-Abhängigkeiten.' : 'All images are stored as Base64 in DEX_EmailTemplates — no external URL dependencies.'}</Callout>,
          },
        ],
      },
    ],
  };
}
