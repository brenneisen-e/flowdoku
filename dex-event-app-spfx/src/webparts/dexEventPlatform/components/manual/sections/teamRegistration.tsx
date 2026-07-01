import * as React from 'react';
import { ManualSection } from '../types';

/**
 * Handbuch-Sektion zur Team-Anmeldung (eingeführt v11.82).
 *
 * Erklärt User + Organizer wie der Team-Anmelde-Flow funktioniert:
 * Voraussetzung am Event, Zustimmungs-Pflicht, Wartelisten-Verhalten bei
 * unzureichender Restkapazität, Pflicht-Bestätigungs-Checkbox, automatische
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
    keywords: isDe
      ? 'Team-Anmeldung Team anmelden Gruppe Gruppen-Anmeldung mehrere Personen anmelden Team-Lead Team-Kapitän Team-Name Mitglied hinzufügen People-Picker Zustimmung Zustimmungs-Checkbox Team beitreten Beitritts-Anfrage offene Slots Team bearbeiten Team verwalten Lead übergeben Break-Out Session Break-Out Sessions Gruppe Tisch Team-Größe komplette Teams Teil-Teams Warteliste Team Schritt 8 Drag and Drop zuordnen Per-Team-Mail Einwahllink'
      : 'team registration register team group registration register multiple people team lead team captain team name add member people picker consent consent checkbox join team join request open slots edit team manage team transfer lead break-out session break-out sessions group table team size complete teams partial teams waitlist team step 8 drag and drop assign per-team mail join link',
    perspectives: [
      {
        perspective: 'user',
        intro: D(
          'Bei Events mit aktivierter Team-Anmeldung (Organizer hat das in Schritt 8 „Team-Anmeldung" erlaubt) kannst du dich und mehrere Personen mit einem einzigen Submit anmelden. Wichtig: jedes Teammitglied muss vorher zugestimmt haben.',
          'For events with team registration enabled (the organizer switches this on in step 8 „Team registration") you can register yourself together with several other people in a single submit. Important: every team member must have consented up front.'
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
              'Wenn der Organizer in Schritt 8 die Option „Team-Namen abfragen" aktiviert hat, erscheint ein Pflichtfeld Team-Name (max. 60 Zeichen). Den Namen siehst du später in „Meine Events" als Badge an der Event-Karte. Wenn der Organizer das deaktiviert hat, taucht die Frage gar nicht auf.',
              'If the organizer enabled „Ask for team name" in step 8, a required input field Team name appears (max 60 characters). The name later shows up as a badge on the event card in „My Events". If the organizer did not enable it, the question is hidden entirely.'
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
          {
            number: 7,
            title: isDe ? 'Mitglied meldet sich ab — was passiert? (v11.83)' : 'A member cancels — what happens? (v11.83)',
            description: D(
              'Sobald sich ein Team-Mitglied über „Meine Events" abmeldet, bekommen alle verbleibenden Mitglieder automatisch eine Info-Mail im Deloitte-Layout. Sie listet den Abgemeldeten, zeigt die aktuelle Belegung (z.B. „2/4") und beschreibt drei Optionen: nichts tun (Platz bleibt erstmal für das Team reserviert), als Team-Lead jemand Neues nachträglich hinzufügen, oder andere Teilnehmer der Event-Anmeldeseite den Slot belegen lassen (sofern der Organizer „Offene Slots öffentlich" aktiviert hat). Wenn der Abgemeldete der Team-Lead war, wird automatisch das früheste verbleibende Mitglied (kleinste TeilnehmerID, sonst früheste Registration-Date) zum neuen Lead promotet — er bekommt in seiner Info-Mail einen extra Hinweis.',
              'When a team member cancels via „My Events", all remaining members automatically receive an info email in the Deloitte layout. It lists who left, shows the current occupancy (e.g. „2/4") and describes three options: do nothing (the seat stays reserved for the team for now), the team lead adds a replacement, or other people fill the slot via the registration page (if the organizer enabled „Public open slots"). If the cancelling person was the team lead, the earliest remaining member (smallest TeilnehmerID, otherwise earliest registration date) is auto-promoted to the new lead — they receive an extra note in their info email.'
            ),
          },
          {
            number: 8,
            title: isDe ? 'Mitglied nachträglich hinzufügen — als Team-Lead (v11.83)' : 'Add a member afterwards — as team lead (v11.83)',
            description: D(
              'Wenn du Team-Lead bist und das Team noch nicht voll ist, siehst du im Team-Badge auf „Meine Events" einen Button „+ Mitglied hinzufügen (N Slot frei)". Klick öffnet ein Modal mit derselben orange Pflicht-Hinweisbox wie die Initial-Team-Anmeldung (Zustimmung des neuen Mitglieds vorher einholen!), einem People-Picker und einer Pflicht-Bestätigungs-Checkbox. Beim Klick auf „Hinzufügen" wird die Person sofort zum Team eingetragen — Anmeldebestätigung, Outlook-Einladung und automatische Info-Mail an die anderen Team-Mitglieder („X ist eurem Team beigetreten") gehen direkt raus. Doppel-Anmelde-Schutz: ist die Person bereits beim Event angemeldet, kommt eine klare Fehlermeldung statt eines Inserts.',
              'If you are team lead and the team is not yet full, the team badge on „My Events" shows a button „+ Add member (N slots free)". Clicking it opens a modal with the same orange consent box as the initial team registration (get the new member\'s consent up front!), a people picker and a required confirmation checkbox. On „Add" the person is immediately joined to the team — confirmation email, Outlook invite and an info mail to all other team members („X joined your team") are queued right away. Duplicate protection: if the person is already registered for the event, you see a clear error message instead of an insert.'
            ),
          },
          {
            number: 9,
            title: isDe ? 'Offenen Team beitreten — als anderer Teilnehmer (v11.83)' : 'Join an open team — as another participant (v11.83)',
            description: D(
              'Wenn der Organizer „Offene Slots öffentlich sichtbar" aktiviert hat, siehst du auf der Event-Anmeldeseite oberhalb des Formulars eine Box „Offene Teams". Sie listet pro Team den Team-Namen (falls vorhanden) und die Belegung (z.B. „2/4 belegt — 2 Slots frei"). Mitgliedernamen werden bewusst NICHT angezeigt (Datenschutz). Klick auf „Beitreten" meldet dich direkt für das Team an — du bekommst Anmeldebestätigung + Outlook-Einladung. Hat der Organizer „Beitritt erfordert Bestätigung durch Team-Kapitän" aktiviert, steht der Button stattdessen auf „Beitritt anfragen": deine Anfrage landet in einer Approve-Queue beim Team-Lead, der entscheidet, und du bekommst eine Mail mit Ergebnis.',
              'When the organizer enabled „Public open slots", the registration page shows a box „Open teams" above the form. Each team is listed with its name (if any) and occupancy (e.g. „2/4 taken — 2 slots free"). Member names are intentionally NOT shown (privacy). Click „Join" to register straight into the team — you receive a confirmation email + Outlook invite. If the organizer enabled „Joining requires approval by team captain", the button reads „Request to join" instead: your request lands in an approve queue with the team lead, they decide, and you receive a result mail.'
            ),
          },
          {
            number: 10,
            title: isDe ? 'Beitritts-Anfragen bearbeiten — als Team-Lead (v11.83)' : 'Handle join requests — as team lead (v11.83)',
            description: D(
              'Wenn dein Event Approval-Pflicht hat, siehst du auf „Meine Events" im Team-Badge einen orange Block „Beitritts-Anfragen (N)" mit einer Liste der offenen Anfragen (Name + E-Mail des Anfragenden). Pro Eintrag stehen zwei Buttons: „Bestätigen" und „Ablehnen". Bestätigen führt sofort die Aufnahme durch (Mail + Outlook für den Beigetretenen, Info-Mail an die anderen Mitglieder). Ablehnen schickt eine kurze Absage-Mail an den Anfragenden mit dem Hinweis, sich einzeln anzumelden oder einem anderen Team beizutreten.',
              'If your event requires approval, „My Events" shows an orange block „Join requests (N)" inside the team badge, listing all open requests (name + email of the requester). Each row has two buttons: „Approve" and „Reject". Approve immediately adds the requester (mail + Outlook for them, info mail to other team members). Reject sends a short decline mail to the requester suggesting they register individually or join another team.'
            ),
          },
          {
            number: 11,
            title: isDe ? 'Team-Lead bearbeitet sein Team von „Meine Events" aus (v11.86)' : 'Team lead manages the team from „My Events" (v11.86)',
            description: D(
              'Ab v11.86 sieht der Team-Lead im Team-Badge der eigenen Event-Karte neben dem „+ Mitglied hinzufügen"-Button einen zweiten Button „Team bearbeiten". Klick öffnet das Modal „Team verwalten" mit allen Mitgliedern als Karten (Profilfoto, Name, E-Mail, Standort, Lead-Badge oder Abgemeldet-Badge). Pro Mitglied außer dem Lead selbst gibt es einen roten Trash-Button. Klick auf den Trash-Button öffnet ein Confirm-Modal mit dem Namen der Person und einem Hinweis, dass die Abmeldung stellvertretend passiert und die Person auch eine eigene Abmelde-Bestätigungs-Mail bekommt. Bestätigen meldet die Person sofort vom Event ab — Abmelde-Mail, Outlook-Termin-Absage, ID-Neuvergabe und Info-Mail an die verbleibenden Mitglieder laufen automatisch. Der Lead selbst sieht keinen Trash-Button — wer sich als Lead abmelden will, nimmt den normalen „Abmelden"-Button auf der Event-Karte; die App promotet dann automatisch das früheste verbleibende Mitglied zum neuen Lead.',
              'From v11.86 onwards the team badge on the lead\'s own event card shows a second button „Edit team" next to „+ Add member". Clicking it opens the „Manage team" modal with all members shown as cards (profile picture, name, email, location, lead badge or cancelled badge). Every member except the lead has a red trash button. Clicking it opens a confirm modal naming the person and pointing out that the cancellation happens on their behalf and that they also receive their own cancellation confirmation email. Confirming cancels the person from the event immediately — cancellation mail, Outlook removal, ID renumbering and info mail to the remaining members all run automatically. The lead has no trash button — to cancel as lead, use the normal „Cancel" button on the event card; the app then auto-promotes the earliest remaining member to the new lead.'
            ),
          },
        ],
      },
      {
        perspective: 'organizer',
        intro: D(
          'Du steuerst die Team-Anmeldung über Schritt 8 (Team-Anmeldung) im Event-Wizard. Hinweis: die Team-Funktion ist auf das Haupt-Event beschränkt — Sub-Events haben keine eigene Team-Anmeldung.',
          'You control team registration via step 8 (team registration) in the event wizard. Note: the team feature is limited to the main event — sub-events do not have their own team registration.'
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Settings in Schritt 8' : 'Settings in step 8',
            description: D(
              'Aktiviere „Team-Anmeldung erlauben", setze die Team-Größe (Default 4, Min 2, Max 20) und entscheide ob ein Team-Name abgefragt werden soll. Im Sub-Bereich „Beitritts-Modus" wählst du, ob nur komplette Teams erlaubt sind oder auch Teil-Teams. Aktivierst du „Auch Teil-Teams erlaubt", kannst du zusätzlich „Offene Slots öffentlich sichtbar" schalten (andere Teilnehmer sehen offene Teams auf der Anmeldeseite und können beitreten) und darunter „Beitritt erfordert Bestätigung durch Team-Kapitän".',
              'Switch on „Allow team registration", set the team size (default 4, min 2, max 20) and decide whether to ask for a team name. In the sub-section „Joining mode" you decide between complete teams only or partial teams allowed. If you enable „Partial teams allowed" you can additionally switch on „Public open slots" (other attendees see open teams on the registration page and can join) and, below it, „Joining requires team captain approval".'
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
          {
            number: 5,
            title: isDe ? 'Admin-Center-Team-Management (v11.84)' : 'Admin Center team management (v11.84)',
            description: D(
              'Im Admin Center erscheint bei Events mit aktivierter Team-Anmeldung über der Teilnehmer-Tabelle eine eigene Sektion „Teams (N)" mit allen aktiven Teams. Pro Team siehst du den Team-Namen (falls vergeben), die Belegung (z.B. „3/4 belegt"), und alle Mitglieder inkl. Profilfoto, Name, E-Mail, Status-Badge und Lead-Badge. Über zwei Buttons kannst du eingreifen: „Person hinzufügen" (nur sichtbar wenn das Team noch freie Slots hat) öffnet das gleiche Modal wie für den Team-Lead in „Meine Events" — orange Pflicht-Hinweisbox zur Zustimmung, People-Picker, Pflicht-Bestätigungs-Checkbox; die neue Person wird sofort angemeldet, bekommt Bestätigungs-Mail + Outlook-Termin, und die anderen Mitglieder bekommen die Info-Mail „X ist eurem Team beigetreten". „Lead-Rolle übergeben" öffnet ein Dropdown mit den anderen aktiven Mitgliedern — Auswahl setzt die alte Lead-Zeile auf TeamLead=false und die neue auf TeamLead=true; alle Mitglieder bekommen eine Info-Mail im Deloitte-Layout, der neue Lead extra mit dem Hinweis, dass er ab jetzt Mitglieder hinzufügen und Beitritts-Anfragen entscheiden darf. Sichtbar nur für Admin oder Organizer des jeweiligen Events.',
              'In the admin center, events with team registration enabled get a dedicated „Teams (N)" section above the participant table, listing all active teams. Per team you see the team name (if any), the occupancy (e.g. „3/4 taken"), and every member including profile picture, name, email, status badge and lead badge. Two buttons let you intervene: „Add person" (only visible if the team has free slots) opens the same modal the team lead uses in „My Events" — orange consent box, people picker, required confirmation checkbox; the new person is immediately registered, receives a confirmation email + Outlook invite, and the other team members get the „X joined your team" info mail. „Transfer lead role" opens a dropdown with the other active members — picking one demotes the old lead and promotes the new one; everyone receives an info email in the Deloitte layout, the new lead with an extra note that they can now add members and decide on join requests. Visible only to admin or the event organizer.'
            ),
          },
          {
            number: 6,
            title: isDe ? 'Teams frei benennen — z.B. „Break-Out Session" (v23.0)' : 'Rename teams freely — e.g. „break-out session" (v23.0)',
            description: D(
              'In Schritt 4 kannst du den Begriff „Team" frei umbenennen — getrennt für Einzahl und Mehrzahl. Trägst du z.B. „Break-Out Session" (Einzahl) und „Break-Out Sessions" (Mehrzahl) ein, heißt die Funktion überall in der App so: im Anmeldeformular, im Organizer Center (Sektion „Break-Out Sessions (N)", „ohne Break-Out Session"-Box), in „Meine Events" (Badge „Break-Out Session „… " — N/M belegt" inkl. „du bist Break-Out Session-Lead"). Lässt du die Felder leer, bleibt es bei „Team"/„Teams".',
              'In step 4 you can freely rename the term „team" — separately for singular and plural. Enter e.g. „break-out session" (singular) and „break-out sessions" (plural), and the feature is labelled that way everywhere: in the registration form, in the organizer center (section „Break-out sessions (N)", „without break-out session" box), and in „My Events" (badge „Break-out session „… " — N/M taken" including „you are break-out session lead"). Leave the fields empty to keep „team"/„teams".'
            ),
          },
          {
            number: 7,
            title: isDe ? 'Teilnehmer dürfen keine eigenen Teams anlegen (v23.0)' : 'Participants may not create their own teams (v23.0)',
            description: D(
              'Direkt darunter in Schritt 4 gibt es den Schalter „Teilnehmer dürfen keine neuen Teams/Gruppen erstellen". Ist er aktiv, verschwindet die „Ich melde mich + mein Team an"-Karte auf der Anmeldeseite komplett — jeder meldet sich einzeln an, und DU teilst die Personen danach selbst in die Gruppen ein (siehe nächster Schritt). Use-Case: Du willst die Break-Out-Aufteilung kontrollieren statt sie den Teilnehmern zu überlassen.',
              'Right below it in step 4 there is a toggle „Participants may not create new teams/groups". When active, the „Register me + my team" card disappears entirely from the registration page — everyone registers individually, and YOU assign people to the groups afterwards (see next step). Use case: you want to control the break-out split instead of leaving it to the participants.'
            ),
          },
          {
            number: 8,
            title: isDe ? 'Personen per Drag & Drop zuordnen (v23.0)' : 'Assign people via drag & drop (v23.0)',
            description: D(
              'In der Teams-Sektion des Organizer Centers kannst du Personen mit der Maus packen und zwischen den Teams/Break-Out-Sessions und der „ohne Team"-Box hin- und herziehen. Das Ziel-Feld färbt sich beim Drüberziehen grün. Loslassen ordnet die Person dem Team zu (oder löst die Zuordnung, wenn du sie in die „ohne Team"-Box ziehst). War der Gezogene ein Team-Lead und bleiben Mitglieder übrig, rückt automatisch das früheste verbleibende Mitglied als neuer Lead nach. Jede Verschiebung wird im Änderungsprotokoll festgehalten.',
              'In the teams section of the organizer center you can grab people with the mouse and drag them between the teams/break-out sessions and the „without team" box. The target area turns green while you hover over it. Dropping assigns the person to the team (or clears the assignment if you drop them into the „without team" box). If the dragged person was a team lead and members remain, the earliest remaining member is auto-promoted to the new lead. Every move is recorded in the change log.'
            ),
          },
          {
            number: 9,
            title: isDe ? 'Per-Team-Mail mit eigenem Einwahllink (v23.0)' : 'Per-team mail with its own join link (v23.0)',
            description: D(
              'Sobald aktive Teams existieren, erscheint in der Teams-Sektion der Button „Mail an <Teams>". Er öffnet ein Modal: oben Betreff + Mail-Text (mit den Platzhaltern {{Vorname}}, {{Name}}, {{TeamName}}, {{EventTitle}} und {{TeamInfo}}), darunter pro Team ein eigenes Info-Feld. In dieses Info-Feld trägst du die team-spezifische Information ein — typischerweise einen eigenen Microsoft-Teams-Einwahllink je Break-Out-Session. Beim Versand bekommt jedes aktive Mitglied eine EIGENE Mail im Deloitte-Layout, in der {{TeamInfo}} durch die Info SEINES Teams ersetzt ist (Links werden automatisch klickbar). Ideal, um nach der Aufteilung allen Gruppen in einem Rutsch ihren jeweiligen Treffpunkt/Link zu schicken. Die Mail respektiert den E-Mail-Schalter aus Schritt 6 — sind E-Mails fürs Event deaktiviert, wird nichts versendet.',
              'As soon as active teams exist, the teams section shows a „Mail to <teams>" button. It opens a modal: subject + mail body at the top (with the placeholders {{Vorname}}, {{Name}}, {{TeamName}}, {{EventTitle}} and {{TeamInfo}}), and below it one info field per team. Into that info field you enter the team-specific information — typically a dedicated Microsoft Teams join link per break-out session. On send, each active member receives their OWN mail in the Deloitte layout, with {{TeamInfo}} replaced by THEIR team\'s info (links are made clickable automatically). Perfect for sending every group its own meeting point/link in one go after the split. The mail respects the email switch from step 6 — if emails are disabled for the event, nothing is sent.'
            ),
          },
        ],
      },
    ],
  };
}
