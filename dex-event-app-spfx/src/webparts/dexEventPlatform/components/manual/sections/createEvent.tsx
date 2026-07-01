import * as React from 'react';
import { ManualSection } from '../types';
import { ClickPath } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import Header from '../../Header';
import EventCreationPage from '../../EventCreationPage';

/**
 * v26.40: Handbuch-Sektion für den Event-Erstellungs-Wizard.
 *
 * Der Wizard hat aktuell 10 Schritte (currentStep 0..9). Diese Sektion
 * dokumentiert jeden Schritt einzeln und in der realen Reihenfolge:
 *   0 Grundlagen · 1 Organizer & Team · 2 Sub-Events · 3 Ort & Programm ·
 *   4 Kapazität & Sichtbarkeit · 5 Felder · 6 Kommunikation ·
 *   7 Team-Anmeldung · 8 Dokumente · 9 Fun-Zone.
 *
 * WICHTIG: id ('create-event'), category und visibleFor bleiben stabil —
 * andere Dateien/Links referenzieren diese id.
 */
export function createEventSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';

  // Kleine Helfer-Factory, damit die 10 Wizard-Previews einheitlich aussehen.
  // Jeder Preview öffnet die echte EventCreationPage im Create-Modus, setzt
  // via initialStep den gewünschten Wizard-Schritt (0..9). Laptop-Frame, weil
  // das Event-Anlegen auf dem Desktop stattfindet.
  const wizardPreview = (stepIndex: number, stepLabel: string): React.ReactElement => (
    <AppPreview
      label={stepLabel}
      role="Organizer"
      page="create-event"
      width={1024}
      device="laptop"
      initialStep={stepIndex}
    >
      <Header />
      <EventCreationPage />
    </AppPreview>
  );

  return {
    id: 'create-event',
    title: isDe ? 'Event erstellen (10-Schritte-Wizard)' : 'Create an event (10-step wizard)',
    category: 'organizer',
    description: isDe
      ? 'Schritt-für-Schritt-Anleitung durch alle zehn Stufen der Event-Erstellung.'
      : 'Step-by-step walkthrough of all ten event-creation stages.',
    visibleFor: ['Admin', 'Organizer'],
    // Such-Stichwörter/Synonyme, damit Suche + Frage-Vorschläge die Sektion finden.
    keywords:
      'Event erstellen anlegen neu Wizard Assistent Schritte Grundlagen Titel Datum Beschreibung Bild ' +
      'Organizer Co-Organizer Ansprechpartner Test-Team Scanner Check-in QR-Code Sub-Event Sub-Events Workshop Session ' +
      'Ort Adresse Veranstaltungsort Agenda Programm Transfer Transferzeiten ' +
      'Kapazität Teilnehmerzahl Warteliste Waitlist geteilte Kapazität Split Sichtbarkeit Audience Zielgruppe Standort-Filter Verteiler Frist Deadline Abmeldefrist ' +
      'Felder Custom Fields Abfragen Dropdown Roommate Anrede Dokument-Upload Sprache ' +
      'Kommunikation Mail E-Mail Outlook Templates Vorlagen Logo Team-Anmeldung Teams Dokumente Anhänge Fun-Zone Quiz ' +
      'create event visibility capacity waitlist split communication documents quiz custom fields team registration',
    perspectives: [
      {
        perspective: 'organizer',
        intro: (
          <>
            <p style={{ margin: '0 0 8px 0' }}>
              {isDe
                ? 'Die Event-Erstellung führt dich durch zehn Schritte. Nur Pflichtfelder (mit *) sind zum Weiterklicken erforderlich — Titel, Start- und Enddatum (Schritt 1) sowie mindestens ein Organizer (Schritt 2). Alles andere kannst du jederzeit später ergänzen; viele Schritte (Sub-Events, Team-Anmeldung, Dokumente, Fun-Zone) sind komplett optional. Jeder Schritt unten hat unter „Vorschau der echten App" einen Button, der dir den jeweiligen Wizard-Step direkt öffnet.'
                : 'Event creation walks you through ten steps. Only required fields (marked *) must be filled before advancing — title, start and end date (step 1) and at least one organizer (step 2). Everything else can be added later; many steps (sub-events, team registration, documents, fun zone) are entirely optional. Each step below has a "Real app preview" button that opens the matching wizard step.'}
            </p>
            <p style={{ margin: '0 0 8px 0' }}>
              {isDe
                ? 'In der Aktionsleiste unten (auf jedem Schritt sichtbar) findest du zwei Werkzeuge: „Vorschau" zeigt live, wie die Registrierungsseite für Teilnehmer aussieht (aktiv, sobald ein Titel eingegeben ist), und „Prüfen" öffnet eine Übersicht aller Event-Einstellungen. Rechts daneben liegen Zurück/Weiter und — auf dem letzten Schritt — „Event erstellen".'
                : 'The action bar at the bottom (visible on every step) offers two tools: "Preview" shows live how the registration page looks to attendees (enabled once a title is set), and "Review" opens an overview of all event settings. To the right sit Back/Next and — on the last step — "Create event".'}
            </p>
            <p style={{ margin: 0 }}>
              {isDe
                ? 'Tipp — Demo-Daten laden: Auf Schritt 1 (Grundlagen) findest du oben rechts im grünen Kopf den Button „Demo". Ein Klick öffnet ein Auswahl-Modal mit fertigen Vorlagen (einfaches Meeting, Event mit Gruppen, Event mit Sub-Event, Sub-Event + Team). Wenn du bereits eigene Events angelegt hast, erscheint darunter außerdem ein Fächer „Eigenes Event als Vorlage nutzen?" — damit übernimmst du Einstellungen und Bild eines vergangenen Events (Datum und Anmeldungen legst du danach neu fest).'
                : 'Tip — load demo data: On step 1 (Basics) the green header holds a "Demo" button in the top-right. Clicking it opens a picker with ready-made templates (simple meeting, event with groups, event with sub-event, sub-event + team). If you already created your own events, a fan "Use one of your events as a template?" appears below — it reuses settings and image from a past event (you set date and registrations fresh afterwards).'}
            </p>
          </>
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Schritt 1: Grundlagen' : 'Step 1: Basics',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Hier legst du das Fundament des Events fest: Titel, Datum, Beschreibung und Bild.'
                    : 'Here you define the foundation of the event: title, date, description and image.'}
                </p>
                <ul style={{ marginTop: 0, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe
                    ? 'Titel * — der Name des Events (Pflichtfeld).'
                    : 'Title * — the name of the event (required).'}</li>
                  <li>{isDe
                    ? 'Start- und Enddatum * — jeweils mit Datum UND Uhrzeit (Berliner Zeit). Diese Werte gehen direkt in den Outlook-Termin. Beim Setzen des Startdatums werden Anmelde-Frist (7 Tage vorher) und letzte Abmeldemöglichkeit (3 Tage vorher) automatisch vorbelegt — anpassbar in Schritt 5.'
                    : 'Start and end date * — each with date AND time (Berlin time). These values flow straight into the Outlook invite. Setting the start date auto-fills the registration deadline (7 days before) and last cancellation (3 days before) — adjustable in step 5.'}</li>
                  <li>{isDe
                    ? 'Beschreibung — optional. Öffne „Bearbeiten & Vorschau" für einen Rich-Text-Editor. Wiederholst du hier nur Titel/Datum/Ort, warnt die App (kein Mehrwert für Teilnehmer).'
                    : 'Description — optional. Open "Edit & Preview" for a rich-text editor. If you only repeat title/date/location, the app warns you (no added value for attendees).'}</li>
                  <li>{isDe
                    ? 'Event-Bild — optional, mit Zuschneiden (Crop) und pro Ansicht (Karte / Hero auf der Anmeldeseite) getrennt einstellbarem Ausschnitt und Zoom.'
                    : 'Event image — optional, with crop and per-view (card / hero on the registration page) separately adjustable framing and zoom.'}</li>
                  <li>{isDe
                    ? 'Entwurf / „Aktiv ab" — standardmäßig ist das Event ein Entwurf und für Teilnehmer unsichtbar. Mit „Aktiv ab" planst du die Veröffentlichung; per Radio wählst du, ob es davor komplett unsichtbar bleibt oder schon als Vorschau in der Event-Liste erscheint.'
                    : 'Draft / "Active from" — by default the event is a draft and hidden from attendees. "Active from" schedules publication; a radio lets you choose whether it stays fully hidden until then or already appears as a preview in the event list.'}</li>
                </ul>
              </>
            ),
            mockup: wizardPreview(0, isDe ? 'Wizard Schritt 1: Grundlagen (echte Ansicht)' : 'Wizard step 1: Basics (real view)'),
            warning: isDe
              ? 'Start- und Enddatum inkl. Uhrzeit gehen direkt in den Outlook-Termin — sorgfältig prüfen, bevor Einladungen rausgehen. Die App verhindert, dass das Ende vor dem Start liegt.'
              : 'Start and end date incl. time flow directly into the Outlook invite — double-check before invitations go out. The app prevents the end being before the start.',
          },
          {
            number: 2,
            title: isDe ? 'Schritt 2: Organizer & Team' : 'Step 2: Organizers & Team',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Wer verantwortet das Event und welches erweiterte Team ist beteiligt. Personen fügst du überall als Chips hinzu: ins Suchfeld tippen (Suche über alle Deloitte-User, per Häkchen auch international), aus der Liste wählen — der Chip erscheint darüber.'
                    : 'Who owns the event and which extended team is involved. You add people everywhere as chips: type into the search field (search across all Deloitte users, optionally international via a checkbox), pick from the list — the chip appears above.'}
                </p>
                <ul style={{ marginTop: 0, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe
                    ? 'Organizer * (grüne Chips) — mindestens einer ist Pflicht; du selbst bist als erster Organizer vorbelegt. Reihenfolge per ◀ ▶, Entfernen per ×. Alle Organizer werden beim Anlegen gemeinsam als Pflicht-Teilnehmer in den Outlook-Termin eingeladen. Optional lassen sich Organizer auf der Anmeldeseite ausblenden (ganz oder einzeln) und die Anzeigegröße (klein/groß) wählen.'
                    : 'Organizers * (green chips) — at least one is required; you are pre-filled as the first organizer. Reorder via ◀ ▶, remove via ×. All organizers are jointly invited as required attendees to the Outlook event on creation. Optionally you can hide organizers on the registration page (all or individually) and pick the display size (small/large).'}</li>
                  <li>{isDe
                    ? 'Ansprechpartner (optional, aufklappbar) — falls jemand Externes für Rückfragen zuständig ist: Name, E-Mail und ein Freitext-Infofeld.'
                    : 'Contact person (optional, collapsible) — if someone external handles queries: name, email and a free-text info field.'}</li>
                  <li>{isDe
                    ? 'Test-Team (blaue Chips) — eine kleine Gruppe, die das Event schon im Entwurfsmodus sieht und sich testweise anmelden darf.'
                    : 'Test team (blue chips) — a small group that sees the event already in draft mode and may register for testing.'}</li>
                  <li>{isDe
                    ? 'QR-Code-Scanner / Check-in-Team (orange Chips) — Personen, die am Event-Tag nur das Check-in-Tool bedienen dürfen, ohne das Event bearbeiten zu können.'
                    : 'QR code scanner / check-in team (orange chips) — people who may only operate the check-in tool on event day, without editing the event.'}</li>
                </ul>
              </>
            ),
            mockup: wizardPreview(1, isDe ? 'Wizard Schritt 2: Organizer & Team (echte Ansicht)' : 'Wizard step 2: Organizers & Team (real view)'),
            tip: isDe
              ? 'Nur Nutzer mit Rolle „Organizer" oder „Admin" erscheinen in der Organizer-Suche. Fehlt jemand → Admin bitten, die Rolle zu setzen. Ist dieselbe Person mehrfach gelistet (z.B. Organizer + Test-Team), warnt die App und bietet Schnell-Entfernen an — Organizer haben Test- und Check-in-Rechte ohnehin.'
              : 'Only users with role "Organizer" or "Admin" appear in the organizer search. If someone is missing → ask an admin to grant the role. If the same person is listed multiple times (e.g. organizer + test team), the app warns and offers quick-remove — organizers already have test and check-in rights.',
          },
          {
            number: 3,
            title: isDe ? 'Schritt 3: Sub-Events' : 'Step 3: Sub-events',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Optional. Gliedere dein Event in einzeln buchbare Programmbausteine (z.B. Workshops, Sessions oder ein Networking-Dinner). Standardmäßig aus — zum Aktivieren den Schalter „Sub-Events nutzen?" umlegen.'
                    : 'Optional. Split your event into individually bookable building blocks (e.g. workshops, sessions or a networking dinner). Off by default — flip the "Use sub-events?" switch to enable.'}
                </p>
                <ul style={{ marginTop: 0, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe
                    ? 'Bezeichnung — wähle, wie die Bausteine heißen (Sub-Event, Workshop, Session, Programmpunkt, Event-Section oder eine eigene Bezeichnung in Ein-/Mehrzahl).'
                    : 'Naming — choose what the blocks are called (sub-event, workshop, session, program item, event section, or a custom singular/plural term).'}</li>
                  <li>{isDe
                    ? 'Anmelde-Modus — „Hauptevent + Bausteine" (Teilnehmer melden sich fürs Hauptevent an und wählen Bausteine dazu) oder „Nur Bausteine" (keine Hauptevent-Anmeldung). Bei „Hauptevent + Bausteine" kannst du zusätzlich das Label des Hauptevents in der Auswahl anpassen.'
                    : 'Registration mode — "Main event + blocks" (attendees register for the main event and add blocks) or "Blocks only" (no main-event registration). With "Main event + blocks" you can also customize the main-event label in the selection.'}</li>
                  <li>{isDe
                    ? 'Pro Sub-Event: Titel, Start-/Enddatum (mit Uhrzeit), Beschreibung sowie ein Häkchen „Pflichtanmeldung für dieses Sub-Event" — dann muss jeder Teilnehmer diesen Baustein mitbuchen.'
                    : 'Per sub-event: title, start/end date (with time), description, plus a checkbox "Mandatory registration for this sub-event" — then every attendee must book that block.'}</li>
                </ul>
              </>
            ),
            mockup: wizardPreview(2, isDe ? 'Wizard Schritt 3: Sub-Events (echte Ansicht)' : 'Wizard step 3: Sub-events (real view)'),
            tip: isDe
              ? 'Sub-Events werden bewusst vor Ort, Kapazität und Feldern konfiguriert: Diese späteren Schritte zeigen pro Sub-Event eigene Tabs, sodass du für jedes Sub-Event eigenen Ort, eigene Kapazität und eigene Felder pflegen kannst — dafür müssen die Sub-Events schon angelegt sein.'
              : 'Sub-events are deliberately configured before location, capacity and fields: those later steps show per-sub-event tabs so you can maintain separate location, capacity and fields for each sub-event — which is why the sub-events must exist first.',
          },
          {
            number: 4,
            title: isDe ? 'Schritt 4: Ort & Programm' : 'Step 4: Location & Programme',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Wo findet das Event statt und wie sieht der Ablauf aus — Ort, Adresse, Agenda und optionale Transferzeiten. Diese Angaben erscheinen auf der Anmeldeseite und in „Meine Events". Alle Felder hier sind optional.'
                    : 'Where the event takes place and how the day is structured — location, address, agenda and optional transfer times. These appear on the registration page and in "My Events". Every field here is optional.'}
                </p>
                <ul style={{ marginTop: 0, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe
                    ? 'Veranstaltungsort — der Name der Location (z.B. „RheinEnergieStadion, Köln").'
                    : 'Location — the venue name (e.g. "RheinEnergieStadion, Cologne").'}</li>
                  <li>{isDe
                    ? 'Adresse — strukturiert in Straße, Hausnummer, PLZ und Ort. Optional lässt sich der „Ort im Outlook-Termin" separat überschreiben (leer = automatisch aus Veranstaltungsort + Adresse).'
                    : 'Address — structured into street, house number, zip and city. Optionally the "Location in the Outlook event" can be overridden separately (empty = automatic from venue + address).'}</li>
                  <li>{isDe
                    ? 'Agenda / Programm — Programmpunkte mit Datum, Start-/Endzeit, Titel und optionaler Beschreibung. Die Punkte werden automatisch chronologisch sortiert und als nummerierte Timeline angezeigt (nicht Teil des Outlook-Texts).'
                    : 'Agenda / programme — items with date, start/end time, title and optional description. Items are auto-sorted chronologically and shown as a numbered timeline (not part of the Outlook body).'}</li>
                  <li>{isDe
                    ? 'Transferzeiten — optionale An-/Abreise-Infos (Stadt, Treffpunkt, Adresse, Datum, Abfahrts-/Ankunftszeit, Beschreibung). Empfohlen für Off-site-Events; bewusst nicht im Outlook-Termin, um Kalenderkonflikte zu vermeiden.'
                    : 'Transfer times — optional arrival/departure info (city, meeting point, address, date, departure/arrival time, description). Recommended for off-site events; deliberately not in the Outlook invite to avoid calendar conflicts.'}</li>
                </ul>
                {isDe
                  ? 'Hat dein Event Sub-Events, kannst du oben über die Tab-Leiste Ort und Programm pro Sub-Event pflegen — inklusive Button „Vom Hauptevent kopieren".'
                  : 'If the event has sub-events, the tab bar on top lets you maintain location and programme per sub-event — including a "Copy from main event" button.'}
              </>
            ),
            mockup: wizardPreview(3, isDe ? 'Wizard Schritt 4: Ort & Programm (echte Ansicht)' : 'Wizard step 4: Location & Programme (real view)'),
          },
          {
            number: 5,
            title: isDe ? 'Schritt 5: Kapazität & Sichtbarkeit' : 'Step 5: Capacity & Visibility',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Wer sieht das Event, wie viele Plätze gibt es und bis wann kann man sich an-/abmelden.'
                    : 'Who sees the event, how many seats are available and by when one can register/cancel.'}
                </p>
                <ul style={{ marginTop: 0, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe
                    ? 'Standort-Filter — wählst du hier einen oder mehrere Standorte, sehen nur Mitarbeiter dieser Standorte das Event. Leer = für alle sichtbar.'
                    : 'Location filter — pick one or more locations and only employees from those locations see the event. Empty = visible to everyone.'}</li>
                  <li>{isDe
                    ? 'Audience (Zielgruppe/Sichtbarkeit) — hier hinterlegst du Mailverteiler und/oder einzelne Personen. Statt jede Person einzeln einzutragen kannst du den Namen eines Outlook-/Mail-Verteilers eingeben — das System löst die Mitglieder automatisch auf. Du kannst außerdem einzelne Personen von der Sichtbarkeit ausschließen.'
                    : 'Audience (target group/visibility) — here you set mailing lists and/or individual people. Instead of adding every person individually you can enter the name of an Outlook / mail distribution group — the system resolves its members automatically. You can also exclude individual people from visibility.'}</li>
                  <li>{isDe
                    ? 'Filterverknüpfung — nur wenn Standort-Filter UND Audience gesetzt sind: „ODER" (einer der Filter reicht, Standard) oder „UND" (beides muss zutreffen).'
                    : 'Filter combination — only when both location filter AND audience are set: "OR" (one filter is enough, default) or "AND" (both must match).'}</li>
                  <li>{isDe
                    ? 'Anmelde-Frist und letzte Abmeldemöglichkeit — jeweils mit Datum + Uhrzeit. Nach der Anmelde-Frist ist der Selbst-Anmelde-Button gesperrt (Organizer können weiterhin manuell hinzufügen); die Abmeldung bleibt bis Event-Ende möglich, Organizer bekommen nach der Frist eine Info-Mail (für Hotel/Catering).'
                    : 'Registration deadline and last cancellation — each with date + time. After the registration deadline the self-register button is locked (organizers can still add people manually); cancellation stays possible until the event ends, and organizers get an info email after the deadline (for hotel/catering).'}</li>
                  <li>{isDe
                    ? 'Max. Teilnehmer / Warteliste — per Schalter „Teilnehmeranzahl begrenzen?" legst du eine Obergrenze fest (Standard: unbegrenzt). Bei begrenzter Kapazität lässt sich die Warteliste aktivieren: Anmeldungen über der Grenze landen in der Warteschlange und rücken automatisch (FIFO) nach, wenn jemand absagt.'
                    : 'Max. participants / waitlist — the "Limit the number of participants?" switch sets a cap (default: unlimited). With limited capacity you can enable the waitlist: registrations over the cap queue up and auto-promote (FIFO) when someone cancels.'}</li>
                  <li>{isDe
                    ? 'Geteilte Kapazität — teile die Plätze optional in zwei frei benannte Gruppen (Gruppe A / Gruppe B, z.B. Vormittag/Nachmittag, VIP/Standard, 5 km/10 km) mit je eigener Platzzahl. Bei aktiver Warteliste wählst du zwischen getrennten Wartelisten pro Gruppe (Standard) und einer gemeinsamen FIFO-Warteliste. Optional lässt sich die Reihenfolge der beiden Boxen im Anmeldeformular umkehren.'
                    : 'Split capacity — optionally split seats into two freely named groups (Group A / Group B, e.g. morning/afternoon, VIP/standard, 5 km/10 km), each with its own seat count. With the waitlist enabled you choose between separate waitlists per group (default) and one shared FIFO waitlist. You can also reverse the order of the two boxes on the registration form.'}</li>
                </ul>
              </>
            ),
            mockup: wizardPreview(4, isDe ? 'Wizard Schritt 5: Kapazität & Sichtbarkeit (echte Ansicht)' : 'Wizard step 5: Capacity & Visibility (real view)'),
            tip: isDe
              ? 'Statt viele Personen einzeln einzutragen, gib im Audience-Feld einfach den Namen eines Verteilers ein — z.B. „DE TT DUESSELDORF". Ändert sich der Verteiler in Outlook, nutzt die App beim nächsten Öffnen den aktuellen Stand. Pflichtfelder pro Gruppe (bei geteilter Kapazität) stellst du in Schritt 6 (Felder) über den Selektor „Sichtbar für Teilnehmergruppe" ein.'
              : 'Instead of adding many people individually, just type a distribution group name in the Audience field — e.g. "DE TT DUESSELDORF". If the group changes in Outlook, the app uses the current membership next time it opens. Per-group required fields (with split capacity) are set in step 6 (Fields) via the "Visible for attendee group" selector.',
          },
          {
            number: 6,
            title: isDe ? 'Schritt 6: Felder' : 'Step 6: Fields',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Hier definierst du eigene Abfragen, die Teilnehmer bei der Anmeldung zusätzlich zu den Standarddaten (Vorname, Nachname, E-Mail sowie Titel/Standort/Abteilung aus dem Profil) sehen. Über „Feld hinzufügen" oder den Katalog „Vorgeschlagene Felder" legst du Felder an; per Drag-Handle oder ▲▼ ordnest du sie um.'
                    : 'Here you define custom questions attendees see at registration in addition to the standard data (first name, last name, email plus job title/location/department from the profile). Add fields via "Add field" or the "Suggested fields" catalogue; reorder them via drag handle or ▲▼.'}
                </p>
                <p style={{ margin: '0 0 4px 0' }}>
                  {isDe ? 'Verfügbare Feldtypen:' : 'Available field types:'}
                </p>
                <ul style={{ marginTop: 0, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe ? 'Text (Freitext), Zahl, Datum (optional mit Uhrzeit), Checkbox.' : 'Text (free text), number, date (optionally with time), checkbox.'}</li>
                  <li>{isDe ? 'Dropdown — mit frei definierbaren Optionen, optional als Mehrfachauswahl.' : 'Dropdown — with freely defined options, optionally multi-select.'}</li>
                  <li>{isDe ? 'Person — Personen-Suche mit Foto und Standort; die gewählte Person kann optional bei An-/Abmelde-Mails auf CC gesetzt werden (nur Mails, nicht der Outlook-Termin).' : 'Person — people search with photo and location; the picked person can optionally be CC\'d on registration/cancellation emails (emails only, not the Outlook invite).'}</li>
                  <li>{isDe ? 'Roommate (Zimmerpartner) — wie Person, löst aber zusätzlich beim Anmelden eine automatische „Zimmerpartner-Anfrage"-Mail an die ausgewählte Person aus; Admin Center zeigt beidseitige Treffer.' : 'Roommate — like Person, but additionally triggers an automatic "roommate request" email to the picked person on registration; the Admin Center shows mutual matches.'}</li>
                  <li>{isDe ? 'Dokument (Upload) — Teilnehmer laden eine Datei (PDF/Bild) hoch, die an ihre Anmeldung angehängt wird.' : 'Document (upload) — attendees upload a file (PDF/image) attached to their registration.'}</li>
                </ul>
                <p style={{ margin: '0 0 4px 0' }}>
                  {isDe ? 'Pro Feld einstellbar:' : 'Configurable per field:'}
                </p>
                <ul style={{ marginTop: 0, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe ? 'Pflichtfeld (blockiert das Absenden, wenn leer).' : 'Required (blocks submit when empty).'}</li>
                  <li>{isDe ? 'Beschreibung — wahlweise als „i"-Info-Tooltip (Hover) oder als Text unter dem Feld-Titel.' : 'Description — either as an "i" info tooltip (hover) or as text below the field title.'}</li>
                  <li>{isDe ? 'Sichtbarkeitsbedingung — Feld nur anzeigen, wenn eine vorherige Frage (Dropdown/Checkbox) einen bestimmten Wert hat. Versteckte Felder blockieren die Pflicht-Validierung nicht.' : 'Visibility condition — show the field only when a previous question (dropdown/checkbox) has a certain value. Hidden fields do not block required-field validation.'}</li>
                  <li>{isDe ? 'Bei geteilter Kapazität (Schritt 5): Selektor „Sichtbar für Teilnehmergruppe" mit „Beide Gruppen", „Nur Gruppe A", „Nur Gruppe B".' : 'With split capacity (step 5): a "Visible for attendee group" selector offering "Both groups", "Group A only", "Group B only".'}</li>
                </ul>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Weitere Einstellungen im Schritt: Toggle „Deutsch und Englisch ermöglichen" (pro Feld einen EN-Text hinterlegen), Dropdown „Sprache des Anmeldeformulars" (Automatisch / Immer Deutsch / Immer Englisch), Toggle „Anrede abfragen?" (Pflicht-Dropdown Frau/Herr/Divers/Keine Angabe) sowie „Sicherheitshinweis vor dem Absenden anzeigen?" — entweder als Auswahl-Übersicht (Haupt-Event + gewählte Sub-Events, ab-/zuwählbar) oder als eigener Hinweistext, den der Teilnehmer per Checkbox bestätigen muss. Hat dein Event Sub-Events, pflegst du die Felder pro Sub-Event über eigene Tabs (mit „Vom Hauptevent kopieren"). Der „Vorgeschlagene Felder"-Katalog enthält eine Allgemein- und eine (eingeklappte) B2Run-Sektion.'
                    : 'Further settings in this step: a toggle "Offer German and English" (add an EN text per field), a "Registration form language" dropdown (Automatic / Always German / Always English), a "Ask for salutation?" toggle (required Mrs/Mr/Diverse/Prefer-not-to-say dropdown) and "Show a confirmation prompt before submitting?" — either as a selection summary (main event + chosen sub-events, de-/selectable) or as a custom hint the attendee must acknowledge via checkbox. If the event has sub-events, you maintain fields per sub-event via their own tabs (with "Copy from main event"). The "Suggested fields" catalogue holds a General and a (collapsed) B2Run section.'}
                </p>
              </>
            ),
            mockup: wizardPreview(5, isDe ? 'Wizard Schritt 6: Felder (echte Ansicht)' : 'Wizard step 6: Fields (real view)'),
            tip: isDe
              ? 'Mit dem „Vorschau"-Button in der Aktionsleiste siehst du jederzeit live, wie das Anmeldeformular für Teilnehmer aussieht — inklusive Sichtbarkeitsbedingungen und i-Tooltips.'
              : 'Use the "Preview" button in the action bar to see live at any time what the attendee\'s registration form looks like — including visibility conditions and i-tooltips.',
            warning: isDe
              ? 'Datenschutz: Erhebe nur Daten, die zwingend für das Event nötig sind, und keine sensiblen personenbezogenen Daten (Gesundheit, Religion, politische Meinung usw.). Bei Unklarheiten immer Rücksprache mit dem Datenschutz halten.'
              : 'Privacy: only collect data strictly necessary for the event, and no sensitive personal data (health, religion, political opinion, etc.). In case of doubt, always check with the data-protection officer.',
          },
          {
            number: 7,
            title: isDe ? 'Schritt 7: Kommunikation' : 'Step 7: Communication',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Steuere, welche E-Mails und Outlook-Einladungen automatisch rausgehen und wie sie aussehen. Ganz oben zeigt eine Übersichts-Box auf einen Blick, was für den gewählten Tab automatisch kommuniziert wird und was nicht.'
                    : 'Control which emails and Outlook invites go out automatically and how they look. An overview box on top shows at a glance what is communicated automatically for the selected tab and what is not.'}
                </p>
                <ul style={{ marginTop: 0, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe
                    ? 'Mail-Sprache — DE oder EN.'
                    : 'Mail language — DE or EN.'}</li>
                  <li>{isDe
                    ? 'Versand-Schalter — der Hauptschalter „Versand von E-Mails und Outlook-Einladungen" (standardmäßig an). Darunter getrennt: Anmelde-Bestätigung und Abmelde-Bestätigung einzeln an-/abschaltbar, sowie „Outlook-Kalendereintrag". Bei aktivem Outlook lässt sich zusätzlich „Outlook-Absage = automatische Abmeldung" aktivieren.'
                    : 'Sending switches — the master switch "Send emails and Outlook invites" (on by default). Below it, separately: registration confirmation and cancellation confirmation each toggle on/off, plus "Outlook calendar entry". With Outlook active you can additionally enable "Outlook decline = automatic deregistration".'}</li>
                  <li>{isDe
                    ? 'Organizer mitlesen (BCC) — pro Anmeldung und pro Abmeldung wählbar: nicht informieren / bei jeder / erst ab einem Datum (bzw. erst nach der letzten Abmeldemöglichkeit).'
                    : 'Loop in organizers (BCC) — choosable per registration and per cancellation: don\'t notify / on every one / only from a date (or only after the last cancellation date).'}</li>
                  <li>{isDe
                    ? 'Logos & Outlook-Text — optionales Event-Logo für den Mail-Header, optionales Logo für den Outlook-Termin und ein anpassbarer Outlook-Termin-Beschreibungstext (Rich-Text-Editor).'
                    : 'Logos & Outlook text — an optional event logo for the mail header, an optional logo for the Outlook invite and an editable Outlook invite description (rich-text editor).'}</li>
                  <li>{isDe
                    ? 'E-Mail-Templates — die System-Vorlagen Anmeldung, Abmeldung und (bei aktiver Warteliste) Warteliste und Nachrücken lassen sich einzeln bearbeiten und wieder zurücksetzen.'
                    : 'Email templates — the system templates Registration, Cancellation and (with the waitlist enabled) Waitlist and Promotion can each be edited and reset.'}</li>
                </ul>
                {isDe
                  ? 'Hat dein Event Sub-Events, gibt es oben eine Tab-Leiste: pro Sub-Event lassen sich eigene Mail-Sprache, eigene Outlook-Texte, eigene Logos und eigene Versand-Einstellungen pflegen — ideal, wenn z.B. ein Sub-Event auf Deutsch und eines auf Englisch läuft.'
                  : 'If the event has sub-events, a tab bar appears on top: each sub-event can have its own mail language, Outlook texts, logos and sending settings — ideal e.g. when one sub-event runs in German and another in English.'}
              </>
            ),
            mockup: wizardPreview(6, isDe ? 'Wizard Schritt 7: Kommunikation (echte Ansicht)' : 'Wizard step 7: Communication (real view)'),
          },
          {
            number: 8,
            title: isDe ? 'Schritt 8: Team-Anmeldung' : 'Step 8: Team Registration',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Optional. Erlaube, dass eine Person ein ganzes Team gleichzeitig anmeldet — z.B. Lauf-Teams, Workshop-Gruppen oder Tische am Abend. Standardmäßig aus; ist der Schalter aus, verhält sich das Event wie gewohnt.'
                    : 'Optional. Let one person register an entire team at once — e.g. running teams, workshop groups or evening tables. Off by default; when the switch is off the event behaves as usual.'}
                </p>
                <ul style={{ marginTop: 0, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe ? 'Team-Anmeldung erlauben — Hauptschalter (Default aus).' : 'Allow team registration — master switch (default off).'}</li>
                  <li>{isDe ? 'Team-Größe — Default 4, Min. 2, Max. 20.' : 'Team size — default 4, min 2, max 20.'}</li>
                  <li>{isDe ? 'Team-Namen abfragen — wenn an, gibt der Team-Lead bei der Anmeldung einen frei wählbaren Team-Namen ein.' : 'Ask for team name — when on, the team lead enters a freely chosen team name at registration.'}</li>
                  <li>{isDe ? 'Bezeichnung statt „Team" — eigener Begriff in Ein- und Mehrzahl (z.B. „Break-Out Session").' : 'Label instead of "Team" — a custom term in singular and plural (e.g. "Break-out session").'}</li>
                  <li>{isDe ? 'Teilnehmer dürfen keine neuen Teams erstellen — empfohlen für Break-Out-Sessions; der Organizer verteilt Personen per Drag & Drop im Admin Center.' : 'Participants cannot create new teams — recommended for break-out sessions; the organizer assigns people via drag & drop in the Admin Center.'}</li>
                  <li>{isDe ? 'Beitritts-Modus — „Nur komplette Teams" (alle Mitglieder gleich eintragen) oder „Auch Teil-Teams erlaubt" (der Lead besetzt z.B. 2 von 4 Slots, der Rest bleibt offen).' : 'Join mode — "Only complete teams" (enter all members at once) or "Partial teams allowed" (the lead fills e.g. 2 of 4 slots, the rest stay open).'}</li>
                  <li>{isDe ? 'Unvollständige Teams öffentlich für Beitritt sichtbar — andere sehen offene Slots als „Team mit X freien Plätzen" (ohne die Namen der bereits Angemeldeten). Zusätzlich aktivierbar: „Beitritt erfordert Bestätigung durch Team-Kapitän" (nur wenn offene Teams sichtbar sind).' : 'Open teams publicly visible for joining — others see open slots as "team with X free seats" (without the names of already-registered members). Additionally: "Joining requires team captain approval" (only when open teams are visible).'}</li>
                </ul>
              </>
            ),
            mockup: wizardPreview(7, isDe ? 'Wizard Schritt 8: Team-Anmeldung (echte Ansicht)' : 'Wizard step 8: Team Registration (real view)'),
          },
          {
            number: 9,
            title: isDe ? 'Schritt 9: Dokumente' : 'Step 9: Documents',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Optional. Lade Dokumente (z.B. PDFs) hoch, die Teilnehmer vor dem Event sehen sollen: Agenda, Anreise-Infos, Sicherheitshinweise. Teilnehmer sehen sie in „Meine Events" als Inline-Vorschau oder Download. Mehrfach-Auswahl und Drag & Drop werden unterstützt.'
                    : 'Optional. Upload documents (e.g. PDFs) attendees should see before the event: agenda, travel info, safety notes. Attendees see them in "My Events" as an inline preview or download. Multi-select and drag & drop are supported.'}
                </p>
                {isDe
                  ? 'Zusätzlich kannst du „Teilnehmer-Upload erlauben" aktivieren: Dann dürfen Teilnehmer eine eigene Datei hochladen (mit anpassbarem Anzeige-Namen wie „Reisekostenbeleg" und einem optionalen Hinweistext).'
                  : 'You can also enable "Allow attendee upload": attendees may upload their own file (with a customizable display name such as "Travel-expense receipt" and an optional hint text).'}
              </>
            ),
            mockup: wizardPreview(8, isDe ? 'Wizard Schritt 9: Dokumente (echte Ansicht)' : 'Wizard step 9: Documents (real view)'),
          },
          {
            number: 10,
            title: isDe ? 'Schritt 10: Fun-Zone (Quiz)' : 'Step 10: Fun Zone (Quiz)',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Optional. Erstelle ein Quiz, das Teilnehmer vor oder während des Events spielen. Fragen lassen sich per Drag & Drop in benannte Bereiche gruppieren — alle Fragen eines Bereichs werden im Quiz zusammen angezeigt.'
                    : 'Optional. Build a quiz for attendees to play before or during the event. Questions can be grouped into named sections via drag & drop — all questions of a section are shown together in the quiz.'}
                </p>
                <ul style={{ marginTop: 0, marginBottom: 6, paddingLeft: 18, lineHeight: 1.55 }}>
                  <li>{isDe ? 'Pro Frage: Fragetext, beliebig viele Antwortoptionen (mind. 2) und ein optionales Bild.' : 'Per question: question text, any number of answer options (at least 2) and an optional image.'}</li>
                  <li>{isDe ? 'Mehrere richtige Antworten sind möglich — für volle Punkte müssen alle markierten gewählt werden.' : 'Multiple correct answers are supported — all marked ones must be picked for full points.'}</li>
                  <li>{isDe ? 'Bereiche fungieren als „Spielblöcke": Teilnehmer können ihren Fortschritt speichern und später weitermachen.' : 'Sections act as "play blocks": attendees can save their progress and continue later.'}</li>
                  <li>{isDe ? 'Live-Highscore und Statistiken (welche Fragen am häufigsten falsch beantwortet werden) siehst du im Admin Center.' : 'Live highscore and statistics (which questions are most often answered incorrectly) are visible in the Admin Center.'}</li>
                </ul>
                {isDe ? 'Details siehe Kapitel „Quiz / Fun-Zone".' : 'See the "Quiz / Fun Zone" chapter for details.'}
              </>
            ),
            mockup: wizardPreview(9, isDe ? 'Wizard Schritt 10: Fun-Zone (echte Ansicht)' : 'Wizard step 10: Fun Zone (real view)'),
          },
          {
            number: 11,
            title: isDe ? 'Abschluss: Vorschau, Prüfen & Event erstellen' : 'Finish: preview, review & create event',
            description: (
              <>
                {isDe
                  ? 'In der Aktionsleiste unten kannst du jederzeit „Vorschau" (so sehen Teilnehmer die Registrierungsseite) und „Prüfen" (alle Einstellungen im Überblick) öffnen. Auf dem letzten Schritt legst du das Event mit „Event erstellen" an: im Hintergrund entsteht automatisch eine SharePoint-Subsite mit der Teilnehmerliste, und dein Event taucht in der Übersicht auf. Im Bearbeiten-Modus kannst du jederzeit über „Speichern & zurück zum Event" sichern, ohne durch alle Schritte zu klicken.'
                  : 'In the action bar at the bottom you can open "Preview" (how attendees see the registration page) and "Review" (all settings at a glance) at any time. On the last step you commit the event with "Create event": a SharePoint subsite with the participant list is auto-provisioned in the background, and your event appears in the overview. In edit mode you can save any time via "Save & return to event" without clicking through every step.'}
              </>
            ),
            mockup: <ClickPath label={isDe ? 'Vorschau' : 'Preview'} label2={isDe ? 'Event erstellen' : 'Create event'} />,
            warning: isDe
              ? 'Die Subsite wird bei erstmaligem Speichern angelegt — das dauert einige Sekunden. Bitte die Seite nicht neu laden, während „Event wird erstellt…" angezeigt wird.'
              : 'The subsite is provisioned on first save — this takes a few seconds. Do not reload while "Creating event..." is shown.',
          },
        ],
      },
    ],
  };
}
