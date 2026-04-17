import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';

export function idReorderSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'id-reorder',
    title: isDe ? 'IDs neu vergeben & Spalten fixen' : 'Renumber IDs & fix columns',
    category: 'admin',
    description: isDe
      ? 'Admin-Wartungstools für Teilnehmerlisten: sequentielle IDs und fehlende Spalten.'
      : 'Admin maintenance tools for participant lists: sequential IDs and missing columns.',
    visibleFor: ['Admin'],
    perspectives: [
      {
        perspective: 'admin',
        steps: [
          {
            number: 1,
            title: isDe ? '"IDs neu vergeben"' : '"Renumber IDs"',
            description: (
              <>
                {isDe
                  ? 'Durch An- und Abmeldungen können die TeilnehmerIDs einer Liste Lücken haben (1, 2, 4, 7, …). Mit diesem Button triggerst du den Power-Automate-Flow "DEX_IDReorder_TeilnehmerIDs", der alle aktiven Teilnehmer neu sequentiell nummeriert. Gleichzeitig rücken Wartelistler entsprechend nach, falls freie Plätze verfügbar sind.'
                  : 'Registrations and cancellations can cause gaps in attendee IDs (1, 2, 4, 7, …). This button triggers the Power Automate flow "DEX_IDReorder_TeilnehmerIDs" to re-number all active attendees sequentially. At the same time, waitlist members promote accordingly if seats open up.'}
              </>
            ),
            mockup: (
              <Callout variant="warning" title={isDe ? 'Kann Teilnehmer-Mail-Flut auslösen' : 'May trigger an email flood'}>
                {isDe
                  ? 'Wenn viele Wartelistler nachrücken, gehen entsprechend viele Nachrück-Mails raus. Sinnvoll kurz vor dem Event-Tag.'
                  : 'If many waitlisters promote, just as many promotion emails go out. Best used shortly before event day.'}
              </Callout>
            ),
          },
          {
            number: 2,
            title: isDe ? '"Spalten fixen"' : '"Fix columns"',
            description: (
              <>
                {isDe
                  ? 'Bei Schema-Änderungen (neue Pflichtspalten, geänderte View-Reihenfolge) kannst du bestehende Teilnehmerlisten nachträglich anpassen. Der Button ruft fixRegistrationListColumns() auf — fehlende Spalten werden angelegt, View neu sortiert.'
                  : 'When the schema changes (new columns, view reordering), you can retro-fit existing participant lists. The button calls fixRegistrationListColumns() — missing columns are created, view is reordered.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Wann brauche ich das?' : 'When do I need this?',
            description: (
              <>
                {isDe
                  ? 'Nach einem App-Update, das neue Spalten einführt (QuizScore, StarterType usw.). Bestehende Events haben die Spalten sonst nicht, und die Quiz-/B2Run-Features funktionieren dort nicht.'
                  : 'After an app update that introduces new columns (QuizScore, StarterType, etc.). Existing events otherwise lack those columns, and the corresponding features don\'t work there.'}
              </>
            ),
            tip: isDe
              ? 'Das Tool ist idempotent — du kannst es beliebig oft pro Event ausführen, ohne Daten zu verlieren.'
              : 'The tool is idempotent — you can re-run it any number of times without data loss.',
          },
        ],
      },
    ],
  };
}
