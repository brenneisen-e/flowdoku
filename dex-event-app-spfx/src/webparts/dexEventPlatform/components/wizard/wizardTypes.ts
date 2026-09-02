// v30.66: Typen, die bis dahin IM Komponentenkoerper von EventCreationPage
// standen. Die ausgelagerten Wizard-Schritte brauchen sie in ihren
// Props-Vertraegen — ein Typ im Funktionskoerper ist von aussen nicht
// referenzierbar, deshalb hier auf Modul-Ebene und exportiert.
import { CustomFieldInput } from './customFieldInput';
import { EmailOverrideEntry } from './emailOverrideEntry';
import { AgendaItem } from '../../types';

export type ImgView = { zoom: number; posY: number; height?: number };

export interface SubEventDraft {
  id: string;                     // Synthetische Client-ID (für React-Keys); nicht = DB-Id
  dbId?: string;                  // DEX_Events-Id wenn das Sub-Event bereits persistiert wurde
  title: string;
  description?: string;
  location?: string;
  startDate: string;
  endDate: string;
  /** v29.52: ganztägiger Termin — der Outlook-Flow macht daraus isAllDay. */
  allDay?: boolean;
  /** v29.54: Termin als „Frei" statt „Beschäftigt" anzeigen. */
  showAsFree?: boolean;
  maxParticipants?: number;
  registrationDeadline?: string;
  /** v24.64: Pflicht-Sub-Event — Teilnehmer MUSS dieses Sub-Event auswählen. */
  mandatory?: boolean;
  disableEmails?: boolean;
  // v19.22: granulare An-/Abmelde-Mail-Schalter jetzt auch pro Sub-Event.
  disableRegistrationEmail?: boolean;
  disableCancellationEmail?: boolean;
  autoDeregisterOnDecline?: boolean;
  inactiveHandling?: 'notify' | 'autoderegister';
  disableOutlook?: boolean;
  /** Per-Sub-Event Custom-Fields (v10.11+). Ersetzt die hardcoded Funstarter/
   *  Durchstarter-Frage bei B2Run — wer eine zusätzliche Auswahl-Frage pro
   *  Sub-Event will, definiert sie hier individuell. Default: leeres Array
   *  (= Sub-Event ohne zusätzliche Frage, Teilnehmer wählen nur den Termin).
   *  Wird unabhängig vom Hauptevent-customFields gespeichert; mit dem Button
   *  „Vom Hauptevent kopieren" kann der Organizer die Felder duplizieren. */
  customFields?: CustomFieldInput[];
  /** v11.57: Pro-Sub-Event Kommunikations-Einstellungen (Step 5 im Wizard).
   *  Jeder Sub-Event kann eigene Mail-Sprache, Versand-Schalter, Logo-Bilder
   *  und Outlook-Termin-Texte haben. Wird beim Speichern in das jeweilige
   *  DEX_Events-Item geschrieben (siehe persistSubEventsForParent). */
  emailLanguage?: string;
  emailLogoBase64?: string;
  outlookLogoBase64?: string;
  outlookBody?: string;
  outlookHeading?: string;
  outlookSubheading?: string;
  /** v18.42: Betreff des Sub-Event-Outlook-Termins (leer = Sub-Event-Titel). */
  outlookSubject?: string;
  /** v18.44: abweichendes Outlook-Datum/-Ort des Sub-Events (ISO/Text). Leer = übernommen. */
  outlookStart?: string;
  outlookEnd?: string;
  outlookLocation?: string;
  /** v14.4: Pro-Sub-Event Mail-Text-Overrides (Anmeldung / Warteliste /
   *  Abmeldung / Nachrücken). Erlaubt es, jedem Sub-Event eigene Subjects,
   *  Headings und Bodies zu geben — Frage von 2026-05 (3 Sub-Events sollen
   *  jeweils eigene An-/Abmelde-Mails versenden können). Vorher landeten
   *  Änderungen auf einem Sub-Tab fälschlicherweise im Top-Level-Override
   *  → die Sub-Events feuerten die Haupt-Event-Texte ab. */
  emailTemplateOverrides?: Record<string, EmailOverrideEntry>;
  /** v11.57: Snapshot der initialen Outlook-relevanten Felder, um beim Save
   *  zu erkennen, ob die Teilnehmer einen Update-Termin bekommen sollen. */
  initialOutlookEventId?: string;
  initialCalendarLink?: string;
  initialTitle?: string;
  /** v29.52: „Ganztägig" beim Laden — Vergleichswert für den Outlook-Detektor. */
  initialAllDay?: boolean;
  /** v29.54: Anzeige-Status beim Laden — ebenfalls Outlook-relevant. */
  initialShowAsFree?: boolean;
  initialStartDate?: string;
  initialEndDate?: string;
  initialOutlookBody?: string;
  /** v28.30: Kopfbild des Outlook-Termins beim Mount. Ohne diesen Snapshot
   *  konnte der Save-Detektor eine reine Bild-Änderung nicht erkennen — das
   *  „Outlook-Termin aktualisieren?"-Modal bot dann nur das Hauptevent an. */
  initialOutlookLogoBase64?: string;
  /** v15.0 (legacy, ungenutzt ab v15.3): Inheritance-Flags für
   *  pro-Sub-Event-Tabs. Mit v15.3 sind Sub-Events vollwertige Events
   *  mit eigener Konfiguration — Inherit-Flags wurden ersatzlos
   *  gestrichen. Felder bleiben optional im Interface, damit
   *  bestehende Piggyback-JSONs ohne Crash gelesen werden können. */
  inheritLocationFromParent?: boolean;
  inheritCapacityFromParent?: boolean;
  inheritCustomFieldsFromParent?: boolean;
  /** v15.3: pro Sub-Event eigene strukturierte Adresse (analog
   *  Hauptevent). Persistiert als JSON in `LocationAddress`-Spalte. */
  locationAddress?: { street: string; houseNo: string; zip: string; city: string };
  /** v15.3: pro Sub-Event eigene Agenda. Persistiert als JSON in
   *  `Agenda`-Spalte. */
  agenda?: AgendaItem[];
  /** v15.3: pro Sub-Event eigene Transferzeiten. Persistiert als JSON
   *  in `Transfers`-Spalte. */
  transferTimes?: Array<{ id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string }>;
  /** v15.3: pro Sub-Event eigene Abmeldefrist. */
  lastDeregisterDate?: string;
  /** v15.3: pro Sub-Event eigener Standortfilter (comma-separated). */
  locationFilter?: string;
  /** v15.3: pro Sub-Event eigene Zielgruppe (audience-String). */
  audience?: string;
  /** v15.3: pro Sub-Event eigene Filterverknüpfung. */
  filterMode?: 'AND' | 'OR';
  /** v22.10: pro Sub-Event ausgeschlossene Personen (E-Mails). Wird im
   *  DEX_Events-Item (Spalte ExcludedUsers) des Sub-Events persistiert —
   *  vorher hielt der Picker den Ausschluss nur intern (ging beim Reload
   *  verloren). */
  excludedUsers?: string[];
  /** v15.3: pro Sub-Event eigene Warteliste an/aus. */
  waitlistEnabled?: boolean;
  /** v15.3: pro Sub-Event eigene Anrede-Abfrage. */
  askSalutation?: boolean;
  /** v27.11: Pro-Sub-Event eigenes Event-Bild (vorher konnte nur das
   *  Haupt-Event ein Bild haben — childPayload schrieb hart '').
   *  imagePreview = Data-URL (frisch gewählt) oder bestehende SP-URL,
   *  imageFile = neues/zugeschnittenes Bild für den Upload beim Speichern,
   *  imageRemoved = bestehendes Bild wurde entfernt (EventImageUrl wird
   *  beim Speichern geleert). */
  imagePreview?: string;
  imageFile?: File | null;
  imageRemoved?: boolean;
}

export type OutlookConfirmItem = {
  kind: 'top' | 'sub';
  eventId: string;
  title: string;
  changedFields: Array<'title' | 'startDate' | 'endDate' | 'outlookBody' | 'location' | 'subject' | 'layout' | 'organizer' | 'logo'>;
  /** v11.68: Sub-Event hat noch keinen Outlook-Termin (kein CalendarLink in
   *  DEX_Events). Body-/Titel-Change wird beim Save in DEX_Events
   *  persistiert, aber es kann KEIN UpdateEvent gequeuet werden — es gibt
   *  keinen Outlook-Termin, an den die Teilnehmer eine Notification kriegen
   *  könnten. Im Modal wird das Item statt mit Checkbox als
   *  Info-Eintrag mit Erklärung gerendert. */
  noOutlookYet?: boolean;
};

export type SuggestedCategory = 'general' | 'b2run';

export type SuggestedEntry = { key: string; label: string; description: string; category: SuggestedCategory; icon: string; tooltip?: string; build: (_now: number) => CustomFieldInput };
