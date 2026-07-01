import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';

export function faqSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  const faqs = isDe
    ? [
        { q: 'Ich sehe ein Event nicht, das Kolleg:innen sehen — warum?', a: 'Der Organizer hat das Event auf bestimmte Standorte oder eine Audience-Liste beschränkt. Dein Standort oder deine E-Mail muss zum Filter passen.' },
        { q: 'Ich bekomme keine Bestätigungs-E-Mail.', a: 'Der Organizer kann Mails pro Event deaktiviert haben. Wenn nicht: prüfe Spam-Ordner und warte 5 Minuten (Queue läuft alle paar Minuten).' },
        { q: 'Mein Outlook-Termin ist verschwunden.', a: 'Wahrscheinlich hast du dich abgemeldet (Outlook-Termin wird automatisch entfernt). Melde dich erneut an, dann kommt er zurück. Oder jemand hat den Termin im Namen des Organizers gelöscht — in dem Fall Organizer anschreiben.' },
        { q: 'Ich möchte ein Event duplizieren.', a: 'Aktuell kein Ein-Klick-Duplizieren. Erstelle ein neues Event und kopiere Felder manuell — Workflow-Verbesserung ist in Arbeit.' },
        { q: 'Wie blockiere ich System-Mails für ein Event?', a: 'Als Organizer im Wizard-Schritt "Kommunikation" (Schritt 7): Schalter "Benachrichtigungen deaktivieren". Nur sinnvoll für Tests oder sehr kleine Events.' },
        { q: 'Wer kann meine Registrierung sehen?', a: 'Nur du selbst, der Event-Organizer und Admins. Item-Level Security schützt deine Daten vor anderen Teilnehmern.' },
        { q: 'Kann ich einen Outlook-Verteiler als Zielgruppe eintragen?', a: 'Ja. Trage im Wizard-Schritt "Kapazität & Sichtbarkeit" (Schritt 5) im Feld "Audience" statt einzelner Personen den Namen eines Outlook-/Mail-Verteilers ein (mehrere kommasepariert) — das System löst die Mitglieder automatisch auf. Details im Artikel "Verteilergruppen / Mail-Verteiler als Zielgruppe".' },
        { q: 'Kann ich DEX am Handy nutzen?', a: 'Ja. DEX ist eine SharePoint-Seite — am Smartphone öffnest du sie am zuverlässigsten über die kostenlose Microsoft-SharePoint-App (besonders unter Android, wo der mobile Browser die Seite teils nicht korrekt lädt). Mehr im Artikel "DEX am Handy nutzen".' },
      ]
    : [
        { q: 'I don\'t see an event my colleagues see — why?', a: 'The organizer scoped the event to certain locations or an audience list. Your location or email must match the filter.' },
        { q: 'I didn\'t receive a confirmation email.', a: 'The organizer may have disabled emails for this event. Otherwise: check your spam folder and wait 5 minutes (the queue runs every few minutes).' },
        { q: 'My Outlook appointment disappeared.', a: 'Likely you cancelled (Outlook appointment is auto-removed). Re-register to get it back. Or someone deleted the appointment in the organizer\'s mailbox — in that case contact the organizer.' },
        { q: 'I want to duplicate an event.', a: 'Currently no one-click duplication. Create a new event and copy fields manually — workflow improvement planned.' },
        { q: 'How do I suppress system mails for an event?', a: 'As organizer in the wizard "Communication" step (step 7): toggle "Disable notifications". Only useful for testing or very small events.' },
        { q: 'Who can see my registration?', a: 'Only you, the event organizer, and admins. Item-level security protects your data from other attendees.' },
        { q: 'Can I use an Outlook distribution list as the audience?', a: 'Yes. In the wizard "Capacity & visibility" step (step 5), in the "Audience" field, enter the name of an Outlook / mail distribution group instead of individual people (several comma-separated) — the system resolves the members automatically. See the article "Distribution groups as audience" for details.' },
        { q: 'Can I use DEX on my phone?', a: 'Yes. DEX is a SharePoint page — on a smartphone it works most reliably via the free Microsoft SharePoint app (especially on Android, where the mobile browser sometimes fails to load the page). See the article "Using DEX on mobile".' },
      ];

  return {
    id: 'faq',
    title: isDe ? 'FAQ & Troubleshooting' : 'FAQ & Troubleshooting',
    category: 'general',
    description: isDe
      ? 'Häufige Fragen, schnelle Antworten.'
      : 'Frequent questions, quick answers.',
    keywords: 'faq häufige fragen troubleshooting fehlerbehebung problem probleme hilfe help frequently asked questions event nicht sichtbar filter audience zielgruppe verteiler bestätigungsmail confirmation email spam outlook termin verschwunden appointment duplizieren duplicate system-mails benachrichtigungen deaktivieren notifications registrierung sichtbar item-level security handy mobil mobile sharepoint-app android',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'user',
        steps: faqs.map((f, i) => ({
          number: i + 1,
          title: f.q,
          description: f.a,
          mockup: i === 0 ? (
            <Callout variant="info">
              {isDe
                ? 'Tipp: Frag deinen Event-Organizer, ob das Event einen Filter hat. Die Antwort gibt dir schnell Klarheit.'
                : 'Tip: Ask the event organizer whether the event has a filter. Their answer clarifies the situation quickly.'}
            </Callout>
          ) : undefined,
        })),
      },
    ],
  };
}
