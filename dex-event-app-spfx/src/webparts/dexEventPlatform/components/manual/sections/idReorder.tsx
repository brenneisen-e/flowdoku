import * as React from 'react';
import { ManualSection } from '../types';
import { AppPreview } from '../previews/AppPreview';
import { DEMO_EVENT_ID } from '../previews/PreviewProviders';
import Header from '../../Header';
import AdminPage from '../../AdminPage';

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
              <AppPreview
                label={isDe ? 'Admin Center → IDs neu vergeben / Spalten fixen (echte Ansicht)' : 'Admin center → reorder IDs / fix columns (real view)'}
                role="Admin"
                page="admin"
                selectedEventId={DEMO_EVENT_ID}
                width={1024}
                device="laptop"
              >
                <Header />
                <AdminPage />
              </AppPreview>
            ),
            warning: isDe
              ? 'Wenn viele Wartelistler nachrücken, gehen entsprechend viele Nachrück-Mails raus. Sinnvoll kurz vor dem Event-Tag.'
              : 'If many waitlisters promote, just as many promotion emails go out. Best used shortly before event day.',
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
          {
            number: 4,
            title: isDe ? 'Reaktivierung & TeilnehmerIDs' : 'Reactivation & participant IDs',
            description: (
              <>
                {isDe
                  ? 'Wenn sich ein zuvor abgemeldeter User erneut anmeldet, wird sein bestehender Listen-Eintrag wiederverwendet (statt ein neues Item anzulegen) — ChangeLog und Audit-Historie bleiben so erhalten. Die TeilnehmerID wird dabei direkt aus dem Subsite-Counter neu gezogen: wer mal #12 war, bekommt nach der Reaktivierung die nächst-freie ID am Ende der Liste (z.B. #87). Logisch passt das, weil die Reihenfolge der RegistrationDate folgt — und bei einer Reaktivierung ist das eben das aktuelle Datum.'
                  : 'When a previously cancelled user re-registers, their existing list item is reused (instead of creating a new one) — preserving the change log and audit history. The TeilnehmerID is freshly pulled from the subsite counter: someone who was #12 before will get the next available ID at the end of the list after reactivation (e.g. #87). This matches the registration-date order — and on reactivation, that order entry is "now".'}
              </>
            ),
            tip: isDe
              ? 'Soll die Liste nach mehreren Reaktivierungen wieder lückenlos durchnummeriert sein, einmal "IDs neu vergeben" klicken — der Flow sortiert Aktive (1..N) und Warteliste (N+1..N+M) sauber.'
              : 'After several reactivations, click "Renumber IDs" once to get a gap-free sequence — the flow sorts active (1..N) and waitlist (N+1..N+M) cleanly.',
          },
        ],
      },
    ],
  };
}
