import * as React from 'react';
import { ManualSection } from '../types';

/**
 * v22.21: Handbuch-Sektion für das geführte Tutorial (Onboarding-Tour).
 */
export function tutorialSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'tutorial',
    title: isDe ? 'Geführtes Tutorial (Onboarding)' : 'Guided tutorial (onboarding)',
    category: 'general',
    description: isDe
      ? 'Die interaktive Tour durch die App — Schritt für Schritt, direkt auf den echten Seiten.'
      : 'The interactive tour through the app — step by step, directly on the real pages.',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'user',
        steps: [
          {
            number: 1,
            title: isDe ? 'Tutorial starten' : 'Start the tutorial',
            description: (
              <>
                {isDe
                  ? 'Auf der Willkommensseite (ganz am Anfang der App) findest du unter dem Start-Button die grüne Sprechblase „Neu hier? Starte das Tutorial". Ein Klick darauf startet die geführte Tour.'
                  : 'On the welcome page (the very first screen of the app) you find the green bubble "New here? Start the tutorial" below the start button. One click starts the guided tour.'}
              </>
            ),
            tip: isDe
              ? 'Du kannst das Tutorial beliebig oft neu starten — es ändert nichts an deinen Daten oder Anmeldungen.'
              : 'You can restart the tutorial as often as you like — it does not change any of your data or registrations.',
          },
          {
            number: 2,
            title: isDe ? 'Der Tour folgen' : 'Follow the tour',
            description: (
              <>
                {isDe
                  ? 'Die Tour führt dich über die echten Seiten der App: Der Bildschirm wird abgedunkelt, das jeweils wichtige Element (z.B. eine Kachel oder ein Button) bleibt hell hervorgehoben, und eine Karte erklärt in wenigen Sätzen, was du dort tun kannst. Mit „Weiter" und „Zurück" bewegst du dich durch die Schritte, „Tour beenden" oder die Escape-Taste brechen jederzeit ab.'
                  : 'The tour guides you across the real pages of the app: the screen dims, the relevant element (e.g. a tile or a button) stays highlighted, and a card explains in a few sentences what you can do there. "Next" and "Back" move you through the steps; "End tour" or the Escape key exit at any time.'}
              </>
            ),
          },
        ],
      },
      {
        perspective: 'organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Tour auswählen' : 'Choose a tour',
            description: (
              <>
                {isDe
                  ? 'Als Organizer (oder Admin) öffnet der Tutorial-Button zuerst eine Auswahl: das Teilnehmer-Tutorial (Events finden, anmelden, QR-Code) oder das Organizer-Tutorial (Event-Wizard, Organizer Center, Check-in am Eventtag). Reguläre User sehen keine Auswahl — bei ihnen startet direkt das Teilnehmer-Tutorial.'
                  : 'As an organizer (or admin) the tutorial button first opens a choice: the attendee tutorial (find events, register, QR code) or the organizer tutorial (event wizard, Organizer Center, check-in on the event day). Regular users see no choice — the attendee tutorial starts right away.'}
              </>
            ),
            tip: isDe
              ? 'Das Organizer-Tutorial zeigt auch den „Demo"-Button im Event-Wizard — damit kannst du nach der Tour gefahrlos ein Beispiel-Event durchspielen.'
              : 'The organizer tutorial also points out the "Demo" button in the event wizard — after the tour you can safely try out a sample event with it.',
          },
        ],
      },
    ],
  };
}
