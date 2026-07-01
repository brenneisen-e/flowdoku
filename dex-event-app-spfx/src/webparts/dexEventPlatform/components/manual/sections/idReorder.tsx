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
    title: isDe ? 'IDs neu vergeben & Spalten fixen' : 'Reassign IDs & fix columns',
    category: 'admin',
    description: isDe
      ? 'Admin-Wartungstools für Teilnehmerlisten: sequentielle IDs und fehlende Spalten.'
      : 'Admin maintenance tools for participant lists: sequential IDs and missing columns.',
    visibleFor: ['Admin'],
    keywords: isDe
      ? 'IDs neu vergeben TeilnehmerID TeilnehmerIDs Nummern neu vergeben durchnummerieren Lücken schließen Reorder Neuvergabe Spalten fixen fehlende Spalten fixRegistrationListColumns Schema-Update QuizScore StarterType Reaktivierung erneut anmelden Überbuchung prüfen überbucht Kapazität DEX_IDReorder Power Automate Flow Wartungstool idempotent'
      : 'reassign IDs renumber IDs participant ID participant IDs renumber sequential close gaps reorder fix columns missing columns fixRegistrationListColumns schema update QuizScore StarterType reactivation re-register overbooking check overbooked capacity DEX_IDReorder power automate flow maintenance tool idempotent',
    perspectives: [
      {
        perspective: 'admin',
        steps: [
          {
            number: 1,
            title: isDe ? '"IDs neu vergeben"' : '"Reassign IDs"',
            description: (
              <>
                {isDe
                  ? 'Durch An- und Abmeldungen können die TeilnehmerIDs einer Liste Lücken haben (1, 2, 4, 7, …). Mit diesem Button triggerst du den Power-Automate-Flow "DEX_IDReorder_TeilnehmerIDs", der alle aktiven Teilnehmer neu sequentiell nummeriert. Gleichzeitig rücken Wartelistler entsprechend nach, falls freie Plätze verfügbar sind. Auch wenn unter Massen-Anmeldungen einmal eine TeilnehmerID nicht gesetzt werden konnte (Counter-Liste vorübergehend nicht erreichbar — sehr selten seit v9.10), kannst du mit diesem Button bestehende Lücken oder Null-Werte sauber wieder durchnummerieren. Seit v11.36 dürfen das auch Organizer für ihre eigenen Events (nicht nur Admins), mit Fortschrittsanzeige in Prozent. Die IDs sind durch den Flow ohnehin immer durchlaufend — meist musst du gar nichts tun. Öffnest du ein Event kurz nach einer Abmeldung, erscheint ein Hinweis-Modal: die automatische Korrektur (Nachrücken + ID-Neuvergabe) läuft evtl. noch im Hintergrund — bitte ein paar Minuten warten, statt parallel manuell zu korrigieren (sonst Doppel-Nachrücken möglich).'
                  : 'Registrations and cancellations can cause gaps in attendee IDs (1, 2, 4, 7, …). This button triggers the Power Automate flow "DEX_IDReorder_TeilnehmerIDs" to re-number all active attendees sequentially. At the same time, waitlist members promote accordingly if seats open up. If a participant ID couldn\'t be assigned during a high-load go-live (counter list temporarily unreachable — very rare since v9.10), this button cleans up gaps or null entries with a fresh sequential renumbering. Since v11.36 organizers may do this for their own events too (not just admins), with a percentage progress bar. IDs stay sequential via the flow anyway — usually nothing to do. If you open an event shortly after a cancellation, a hint modal appears: the automatic correction (promotion + ID renumbering) may still be running in the background — please wait a few minutes instead of correcting manually in parallel (otherwise double-promotion possible).'}
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
              : 'After several reactivations, click "Reassign IDs" once to get a gap-free sequence — the flow sorts active (1..N) and waitlist (N+1..N+M) cleanly.',
          },
          {
            number: 5,
            title: isDe ? 'Überbuchung prüfen & bereinigen' : 'Check & fix overbooking',
            description: (
              <>
                {isDe
                  ? 'Bei sehr vielen zeitgleichen Anmeldungen kann es passieren, dass mehr Personen angemeldet werden als Plätze da sind (pro Gruppe bei Zwei-Gruppen-Events). Der Button "Überbuchung prüfen" findet die zuletzt über Kapazität Angemeldeten und markiert sie — es wird zunächst nichts geändert. Oben in der Teilnehmerliste erscheint die Box "Überbuchung – zu prüfen" mit Buttons pro Person und einem Sammel-Button "Alle bestätigen". Pro Person entscheidest du: "Auf Warteliste" (optional mit Entschuldigungs-Mail im Deloitte-Layout und/oder Abmeldung vom Kalendereintrag) oder "Platz behalten" (entweder als Erste(r) auf der Warteliste oder bleibt angemeldet, dann rückt einmal niemand nach bis die Überzahl absorbiert ist). Nach jeder Entscheidung werden die TeilnehmerIDs automatisch neu vergeben — du siehst dabei einen Fortschrittsbalken mit Prozent. Jede Korrektur wird im ChangeLog der Person vermerkt (war fälschlich angemeldet, ursprünglicher Registrierungszeitpunkt), bleibt also nachvollziehbar.'
                  : 'With a large number of simultaneous registrations, more people can end up registered than there are seats (per group for two-group events). The "Check overbooking" button finds the people registered last over capacity and flags them — nothing is changed yet. A box "Overbooking – to review" appears at the top of the participant list with per-person buttons and a bulk "Confirm all". Per person you decide: "To waitlist" (optionally with an apology email in the Deloitte layout and/or removal from the calendar invite) or "Keep seat" (either first on the waitlist, or stays registered — then one promotion is skipped until the surplus is absorbed). After each decision the participant IDs are renumbered automatically, with a percentage progress bar. Every correction is recorded in the person\'s change log (was mistakenly registered, original registration time), so it stays auditable.'}
              </>
            ),
            tip: isDe
              ? 'Das dürfen auch Organizer für ihre eigenen Events (nicht nur Admins) — es zählt als Teilnehmerverwaltung. Vorbeugend gilt: die App reserviert Plätze seit v11.36 atomar pro Gruppe, im Zweifel (Fehler/Überlast) landet eine Anmeldung sicherheitshalber auf der Warteliste statt fälschlich „Angemeldet".'
              : 'Organizers may do this for their own events too (not just admins) — it counts as participant management. Preventively: since v11.36 the app reserves seats atomically per group; in case of error/overload a registration safely goes to the waitlist instead of wrongly "registered".',
          },
        ],
      },
    ],
  };
}
