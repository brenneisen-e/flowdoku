// v30.66: Aus EventCreationPage.tsx ausgelagert — reine Textkonstanten ohne
// jeden Bezug auf den Wizard-Zustand. Sie standen im Funktionskoerper und wurden
// damit bei jedem Render neu aufgebaut.

export const STEP_HINTS_DE: string[][] = [
  [
    'Event-Titel und Beschreibung — werden auf der Eventliste und der Registrierungsseite angezeigt',
    'Event-Bild hochladen (wird oben auf der Detailseite und in den Mails verwendet)',
    'Als Entwurf speichern — taucht dann nur für Admins, Organizer und Test-Team auf',
  ],
  [
    // v24.12 Schritt 2: Organizer & Team
    'Organizer auswählen — bekommen alle Organizer-Mails und sehen das Event im Admin Center; einzelne lassen sich von der Anmeldeseite ausblenden',
    'Anzeige der Organizer auf dem Anmeldeformular wählen (klein/groß) — mit Live-Vorschau',
    'Optional: externen Ansprechpartner (z.B. Service-Mail) angeben',
    'Test-Team: sieht das Event schon im Entwurf',
    'Check-in-Team: bedient am Event-Tag nur das Check-in-Tool',
  ],
  [
    // v15 Step 3: Ort & Programm (mit Tabs pro Sub-Event)
    'Veranstaltungsort und Adresse erfassen — pro Sub-Event optional eigener Ort (per Tab)',
    'Start- und End-Datum (mit Uhrzeit) festlegen',
    'Agenda mehrtägig pflegen (Drag-Reihenfolge pro Tag)',
    'Transferzeiten — Bus / Shuttle / Bahn von/zum Veranstaltungsort',
  ],
  [
    // v15 Step 4: Kapazität & Sichtbarkeit (mit Tabs pro Sub-Event)
    'Maximale Teilnehmerzahl festlegen (oder Unbegrenzt) — pro Sub-Event eigene Kapazität per Tab (Default: vom Hauptevent übernehmen)',
    'Anmeldefrist setzen — pro Sub-Event eigene Deadline möglich (leer = Hauptevent-Deadline gilt)',
    'Optional: Letzte Abmeldemöglichkeit — die kommunizierte Abmeldefrist; danach bleibt die Abmeldung bis zum Event-Ende möglich, die Organizer werden aber automatisch informiert',
    'Warteliste aktivieren — voll besetzte Events nehmen weitere Anmeldungen auf, bis ein Platz frei wird',
    'Optional: Geteilte Kapazität — zwei frei benannte Gruppen mit eigener Platzzahl + eigener oder gemeinsamer Warteliste',
  ],
  [
    // v15 Step 5: Felder (mit Tabs pro Sub-Event)
    'Feldtyp wählen: Text, Zahl, Dropdown, Checkbox, Personen-Suche oder Roommate (Doppelzimmer)',
    'Mehrfachauswahl bei Dropdowns (z.B. mehrere Allergien anhaken)',
    'Pflichtfeld setzen (rotes Sternchen, Anmeldung blockiert wenn leer)',
    'Beschreibung pro Feld — landet als „i"-Tooltip neben dem Feld-Label',
    'Sichtbarkeitsbedingung: Feld nur dann anzeigen wenn eine andere Frage einen bestimmten Wert hat (z.B. „Zimmerart nur fragen wenn Hotel = ja")',
    'Pro Sub-Event eigene Felder per Tab (Default: vom Hauptevent übernehmen)',
  ],
  [
    // v15 Step 6: Kommunikation
    'E-Mail-Sprache (DE/EN) für die automatischen Mails an die Teilnehmer wählen',
    'Pro Mail-Template (Anmeldung, Storno, Warteliste, Erinnerung, QR-Code…) den Subject/Heading/Body anpassen — mit Live-Vorschau',
    'Eigenes Logo / Header-Bild für Mail und Outlook-Termin hochladen',
    'Outlook-Termin-Body individuell gestalten (Live-Vorschau zeigt wie das Outlook-Element später aussieht)',
    'Benachrichtigungen optional komplett deaktivieren — z.B. für interne Entwurfs-Events',
    'Pro Sub-Event eigene Mail-Texte + Outlook-Body + Disable-Toggles per Tab',
  ],
  [
    // v15 Step 7: Team-Anmeldung
    'Team-Anmeldung erlauben — ein Teilnehmer kann sich für sich + sein Team gleichzeitig anmelden',
    'Team-Größe festlegen (2-N Personen)',
    'Optional Team-Namen abfragen — z.B. für Quiz- oder Lauf-Teams',
    'Beitritts-Modus: nur komplette Teams ODER auch Teil-Teams erlaubt',
    'Optional: offene Slots öffentlich sichtbar — andere Teilnehmer können beitreten (ggf. mit Lead-Approval)',
  ],
  [
    // v15 Step 8: Dokumente
    'Dokumente hochladen (PDF) — Teilnehmer sehen sie auf MyEvents als Inline-Vorschau oder Download',
  ],
  [
    // v15 Step 9: Fun-Zone
    'Quiz-Fragen für das Event anlegen — Multiple-Choice mit beliebig vielen Antwortoptionen',
    'Pro Frage optional ein Bild hochladen (Logo, Foto-Quiz, etc.)',
    'Mehrere richtige Antworten möglich (Mehrfachauswahl) — werden alle für volle Punktzahl gebraucht',
    'Cluster-Größe steuern: wie viele Fragen pro „Spielblock" angezeigt werden — Teilnehmer kann zwischenspeichern und später weitermachen',
    'Live-Highscore + Statistik im Admin Center sehen (welche Fragen am häufigsten falsch beantwortet werden)',
  ],
  // v29.66: Schritt 10 „Abrechnung" (F&A-Pilot, nur Admins sehen den Schritt).
  [
    'Event als abrechnungsrelevant kennzeichnen — die Entscheidung ist jederzeit änderbar',
    'Versandart wählen: automatisch (7 Tage vor/nach dem Event) oder manuell über das Organizer Center',
    'Alle elf Pflichtangaben für Finance & Accounting pflegen — unvollständig blockiert das Speichern nicht',
  ],
];

export const STEP_HINTS_EN: string[][] = [
  [
    'Event title and description — shown on the event list and registration page',
    'Upload an event image (used at the top of the detail page and in emails)',
    'Save as draft — only visible to admins, organizers and the test team',
  ],
  [
    // Step 2: Organizers & Team
    'Pick the organizers — they receive all organizer emails and see the event in the admin center; individual ones can be hidden from the registration page',
    'Choose how organizers appear on the registration form (small/large) — with live preview',
    'Optional: add an external contact (e.g. a service mailbox)',
    'Test team: sees the event while it is still a draft',
    'Check-in team: only operates the check-in tool on event day',
  ],
  [
    // Step 3: Location & Programme (with tabs per sub-event)
    'Enter venue and address — per sub-event an own location is possible (via tab)',
    'Set start and end date (with time)',
    'Maintain a multi-day agenda (drag ordering per day)',
    'Transfer times — bus / shuttle / train to and from the venue',
  ],
  [
    // v15 Step 4: Capacity & Visibility (with tabs per sub-event)
    'Set the maximum number of attendees (or Unlimited) — per sub-event own capacity via tab (default: inherit from main event)',
    'Set the registration deadline — per sub-event own deadline possible (empty = main-event deadline applies)',
    'Optional: last cancellation date — the communicated deadline; cancelling stays possible until the event ends, but organizers are notified automatically',
    'Enable waitlist — full events accept new registrations and promote them once a spot frees up',
    'Optional: split capacity — two freely-named groups with own seat count + own or shared waitlist',
  ],
  [
    // v15 Step 5: Fields (with tabs per sub-event)
    'Pick a field type: text, number, dropdown, checkbox, people search or roommate (double room)',
    'Multi-select for dropdowns (e.g. tick multiple allergies)',
    'Mark required (red asterisk, blocks submit when empty)',
    'Description per field — appears as „i" tooltip next to the field label',
    'Visibility condition: only show this field when another question has a specific value (e.g. „Only ask room type if Hotel = yes")',
    'Per sub-event own fields via tab (default: inherit from main event)',
  ],
  [
    // v15 Step 6: Communication
    'Pick the email language (DE/EN) for automatic emails to attendees',
    'Edit subject / heading / body per email template (registration, cancellation, waitlist, reminder, QR code…) — with live preview',
    'Upload a custom logo / header image for the email and Outlook event',
    'Customise the Outlook event body (live preview shows how the Outlook item will appear)',
    'Optionally disable notifications entirely — e.g. for internal draft events',
    'Per sub-event own mail texts + Outlook body + disable toggles via tab',
  ],
  [
    // v15 Step 7: Team Registration
    'Allow team registration — an attendee can register themselves + their team at once',
    'Set team size (2-N people)',
    'Optionally ask for a team name — e.g. quiz or running teams',
    'Join mode: complete teams only OR partial teams allowed',
    'Optional: open slots publicly visible — other attendees can join (with optional lead approval)',
  ],
  [
    // v15 Step 8: Documents
    'Upload documents (PDF) — attendees see them on MyEvents as inline preview or download',
  ],
  [
    // v15 Step 9: Fun-Zone
    'Create quiz questions for the event — multiple choice with any number of answer options',
    'Optionally upload an image per question (logo, photo quiz, etc.)',
    'Multiple correct answers are supported — all of them must be picked for full points',
    'Control cluster size: how many questions per „play block" — attendees can save progress and continue later',
    'See live highscore + statistics in the admin center (which questions are most often answered incorrectly)',
  ],
  // v29.66: step 10 "Billing" (F&A pilot, admins only).
  [
    'Mark the event as billing-relevant — the decision can be changed at any time',
    'Pick the delivery mode: automatic (7 days before/after the event) or manual via the organizer center',
    'Maintain all eleven mandatory Finance & Accounting details — incomplete data never blocks saving',
  ],
];

// Die Gruppen des Uebertragen-Dialogs. Vorher ein useMemo mit leerer
// Abhaengigkeitsliste — also schon damals eine Konstante.
export const SUB_TRANSFER_GROUPS: Array<{ key: string; de: string; en: string; fields: string[] }> = [
  { key: 'visibility', de: 'Sichtbarkeit (Standortfilter, Mailverteiler, Verknüpfung, Ausschlüsse)', en: 'Visibility (location filter, mailing lists, link mode, exclusions)', fields: ['locationFilter', 'audience', 'filterMode', 'excludedUsers'] },
  { key: 'capacity', de: 'Teilnehmerzahl & Warteliste', en: 'Capacity & waitlist', fields: ['maxParticipants', 'waitlistEnabled'] },
  { key: 'regDeadline', de: 'Anmeldefrist', en: 'Registration deadline', fields: ['registrationDeadline'] },
  { key: 'deregDeadline', de: 'Abmeldefrist', en: 'Cancellation deadline', fields: ['lastDeregisterDate'] },
  { key: 'place', de: 'Ort & Adresse', en: 'Location & address', fields: ['location', 'locationAddress'] },
  { key: 'mandatory', de: 'Pflichtanmeldung', en: 'Mandatory registration', fields: ['mandatory'] },
  { key: 'communication', de: 'Kommunikation (Logo, Outlook-Text, Überschriften, Betreff, Mail-Sprache, Mail-Schalter)', en: 'Communication (logo, Outlook text, headings, subject, mail language, mail toggles)', fields: ['emailLanguage', 'emailLogoBase64', 'outlookLogoBase64', 'outlookBody', 'outlookHeading', 'outlookSubheading', 'outlookSubject', 'disableEmails', 'disableRegistrationEmail', 'disableCancellationEmail', 'autoDeregisterOnDecline', 'inactiveHandling', 'disableOutlook', 'emailTemplateOverrides'] },
  { key: 'times', de: 'Zeiten (Start & Ende) — überschreibt die Termine!', en: 'Times (start & end) — overwrites the dates!', fields: ['startDate', 'endDate'] },
];
