import * as React from 'react';
import { ManualSection } from '../types';

/**
 * Handbuch-Sektion zur Team-Anmeldung (eingefuehrt v11.82).
 *
 * Erklaert User + Organizer wie der Team-Anmelde-Flow funktioniert:
 * Voraussetzung am Event, Zustimmungs-Pflicht, Wartelisten-Verhalten bei
 * unzureichender Restkapazitaet, Pflicht-Bestaetigungs-Checkbox, automatische
 * Mails + Outlook-Termine pro Mitglied, Sub-Event-Hinweis.
 */
export function teamRegistrationSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  const D = (de: string, en: string): React.ReactNode => <>{isDe ? de : en}</>;
  return {
    id: 'team-registration',
    title: isDe ? 'Team-Anmeldung — sich + mehrere Personen gleichzeitig anmelden' : 'Team registration — register yourself + several people at once',
    category: 'general',
    description: isDe
      ? 'Wie du dich + ein ganzes Team in einem Schritt anmeldest.'
      : 'How to register yourself + an entire team in one go.',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'user',
        intro: D(
          'Bei Events mit aktivierter Team-Anmeldung (Organizer hat das in Schritt 4 erlaubt) kannst du dich und mehrere Personen mit einem einzigen Submit anmelden. Wichtig: jedes Teammitglied muss vorher zugestimmt haben.',
          'For events with team registration enabled (the organizer switches this on in step 4) you can register yourself together with several other people in a single submit. Important: every team member must have consented up front.'
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Toggle „Team-Anmeldung" aktivieren' : 'Enable the „team registration" toggle',
            description: D(
              'Im Anmeldeformular siehst du unter „Persönliche Daten" eine Zeile „Ich melde mich + mein Team an (Team-Anmeldung)". Hake sie an. Eine neue Card „Team-Anmeldung" klappt auf — beginnt mit einer auffälligen orange Pflicht-Hinweis-Box zur Zustimmung jedes Mitglieds.',
              'In the registration form, under „Personal information", you will see a checkbox „Register me + my team (team registration)". Tick it. A new „Team registration" card unfolds — starting with a bold orange consent reminder.'
            ),
          },
          {
            number: 2,
            title: isDe ? 'Team-Name (optional)' : 'Team name (optional)',
            description: D(
              'Wenn der Organizer in Schritt 4 die Option „Team-Namen abfragen" aktiviert hat, erscheint ein Pflichtfeld Team-Name (max. 60 Zeichen). Den Namen siehst du später in „Meine Events" als Badge an der Event-Karte. Wenn der Organizer das deaktiviert hat, taucht die Frage gar nicht auf.',
              'If the organizer enabled „Ask for team name" in step 4, a required input field Team name appears (max 60 characters). The name later shows up as a badge on the event card in „My Events". If the organizer did not enable it, the question is hidden entirely.'
            ),
          },
          {
            number: 3,
            title: isDe ? 'Mitglieder per People-Picker auswählen' : 'Pick members via people-picker',
            description: D(
              'Pro Slot (Mitglied 2, Mitglied 3, …) suchst du den Kollegen über Name/E-Mail und übernimmst ihn aus dem Vorschlag (mit Profilbild). Im Modus „Nur komplette Teams" (Default) sind alle Slots Pflicht — du musst N-1 weitere Personen eintragen. Im Modus „Auch Teil-Teams erlaubt" sind die Slots optional; du kannst dich auch alleine als 1er-Team anmelden oder weniger als die maximale Größe wählen.',
              'For each slot (Member 2, Member 3, …) you search the colleague by name/email and pick them from the suggestion (with profile picture). In „Complete teams only" mode (default) every slot is required — you must add N-1 more people. In „Partial teams allowed" mode the slots are optional; you can also register as a 1-person team or with fewer than the max size.'
            ),
          },
          {
            number: 4,
            title: isDe ? 'Zustimmungs-Checkbox bestätigen' : 'Tick the consent box',
            description: D(
              'Ganz unten in der Team-Card ist eine Pflicht-Checkbox: „Ich bestätige, dass alle eingetragenen Teammitglieder ihrer Anmeldung zugestimmt haben." Ohne diesen Haken ist der Submit-Button deaktiviert. Das ist absichtlich so streng — du meldest gleich mehrere Personen an, ein Versehen wäre unangenehm.',
              'At the bottom of the team card you must tick a required checkbox: „I confirm that every listed team member has consented to this registration." Without this tick the submit button stays disabled. This is on purpose — registering several people at once should never happen by accident.'
            ),
          },
          {
            number: 5,
            title: isDe ? 'Submit — was passiert dann?' : 'Submit — what happens next?',
            description: D(
              'Der Button heißt jetzt „Team anmelden (N Personen)". Beim Klick reserviert die App alle N Plätze auf einmal. Sind genug frei, sind alle Angemeldet. Reicht der Platz nicht, geht das gesamte Team gemeinsam auf die Warteliste — nie wird nur ein Teil aktiviert. Jedes Mitglied bekommt automatisch eine Bestätigungs-Mail (mit Hinweis „Du wurdest als Teil eines Teams angemeldet"), einen Outlook-Termin und sieht das Event sofort in „Meine Events" mit Team-Badge.',
              'The button now reads „Register team (N people)". On click the app reserves all N seats at once. If enough seats are free, everyone is Registered. If capacity is short, the entire team goes to the waitlist together — partial activation never happens. Each member automatically receives a confirmation email (with a „You were registered as part of a team" notice), an Outlook invite, and sees the event immediately in „My Events" with a team badge.'
            ),
          },
          {
            number: 6,
            title: isDe ? 'Mitglied hat nicht zugestimmt?' : 'A member did not consent?',
            description: D(
              'Sollte ein Mitglied versehentlich angemeldet worden sein, kann es sich jederzeit selbst über „Meine Events" wieder abmelden — das funktioniert unabhängig vom Team-Lead. Die Bestätigungs-Mail enthält genau diesen Hinweis. Du als Lead siehst die Liste der aktiven Mitglieder im Badge — wer abgemeldet ist, taucht da nicht mehr auf.',
              'If a member was registered by mistake, they can cancel themselves any time via „My Events" — independent of the team lead. The confirmation email explicitly states this. As lead you see the list of active members in the badge — anyone who cancelled disappears from there.'
            ),
          },
        ],
      },
      {
        perspective: 'organizer',
        intro: D(
          'Du steuerst die Team-Anmeldung über Schritt 4 (Team-Anmeldung) im Event-Wizard. Hinweis: die Pro-Sub-Event-Team-Anmeldung folgt erst in v11.83+ — aktuell ist die Team-Funktion auf das Haupt-Event beschränkt.',
          'You control team registration via step 4 (team registration) in the event wizard. Note: per-sub-event team registration is coming in v11.83+ — for now the team feature is limited to the main event.'
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Settings in Schritt 4' : 'Settings in step 4',
            description: D(
              'Aktiviere „Team-Anmeldung erlauben", setze die Team-Größe (Default 4, Min 2, Max 20) und entscheide ob ein Team-Name abgefragt werden soll. Im Sub-Bereich „Beitritts-Modus" wählst du, ob nur komplette Teams erlaubt sind oder auch Teil-Teams. Die Optionen „Offene Slots öffentlich sichtbar" und „Beitritt erfordert Bestätigung" greifen erst, sobald die Beitritts-Funktion in einer späteren Iteration ausgerollt wird.',
              'Switch on „Allow team registration", set the team size (default 4, min 2, max 20) and decide whether to ask for a team name. In the sub-section „Joining mode" you decide between complete teams only or partial teams allowed. The „Public open slots" and „Joining requires approval" options only take effect once the join feature ships in a later iteration.'
            ),
          },
          {
            number: 2,
            title: isDe ? 'Was sieht der User?' : 'What the user sees',
            description: D(
              'Auf der Anmelde-Seite erscheint unter „Persönliche Daten" der Team-Toggle nur wenn du Team-Anmeldung aktiviert hast und die Team-Größe ≥ 2 ist. Sobald aktiviert: orange Pflicht-Hinweis-Box, optional Team-Name-Feld, N-1 People-Picker, Zustimmungs-Checkbox, Submit-Button mit Personen-Zahl.',
              'On the registration page the team toggle below „Personal information" only shows up if you enabled team registration and team size ≥ 2. Once enabled: orange consent box, optional team name input, N-1 people pickers, consent checkbox, submit button with people count.'
            ),
          },
          {
            number: 3,
            title: isDe ? 'Daten in der Teilnehmerliste' : 'Data in the participant list',
            description: D(
              'Pro Mitglied wird ein eigener Eintrag in der Subsite-Teilnehmerliste angelegt — gemeinsam gruppiert über die Spalte TeamId (UUID, identisch für alle Mitglieder). TeamLead ist genau bei einem Eintrag pro Team true (der anmeldenden Person). TeamName ist optional, enthält den frei gewählten Namen. Du kannst die Spalten in der Default-View der Teilnehmerliste am Ende sehen; bei Bedarf das Admin-Tool „Spalten fixen" laufen lassen, um sie auf bestehenden Events nachzurüsten.',
              'Each member gets a separate entry in the subsite participant list — grouped together via the TeamId column (UUID, identical for all members). TeamLead is true for exactly one entry per team (the registering person). TeamName is optional and holds the freely chosen name. You can see these columns at the end of the participant list default view; run the admin tool „Fix columns" if you need to add them to existing events.'
            ),
          },
          {
            number: 4,
            title: isDe ? 'Limitierungen v11.82' : 'Limitations v11.82',
            description: D(
              'In dieser Iteration: Team-Anmeldung läuft nur fürs Haupt-Event, nicht für Sub-Events (die kommen mit v11.83+ dran). „Für andere Person registrieren" und Team-Anmeldung schließen sich gegenseitig aus — der Toggle ist im Stellvertreter-Modus ausgeblendet. Pflicht-Custom-Felder werden nur für den Lead abgefragt — Mitglieder bekommen leere Custom-Field-Antworten. Wenn du detaillierte Daten pro Person brauchst (z.B. Schuhgröße, Diätwünsche), klär das organisatorisch über deinen Standard-Anmeldepfad statt der Team-Funktion.',
              'In this iteration: team registration only covers the main event, not sub-events (those land in v11.83+). „Register for someone else" and team registration are mutually exclusive — the toggle is hidden in proxy mode. Required custom fields only apply to the lead — member entries get empty custom-field answers. If you need per-person details (e.g. shoe size, dietary preferences), handle that organisationally via your standard registration path instead of the team feature.'
            ),
          },
        ],
      },
    ],
  };
}
