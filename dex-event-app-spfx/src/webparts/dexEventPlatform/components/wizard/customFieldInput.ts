/**
 * v28.94: Aus `EventCreationPage` herausgeloest. Der Entwurfs-Typ eines
 * Abfragefelds im Assistenten (vor dem Serialisieren nach SharePoint).
 */
export interface CustomFieldInput {
  id: string;
  label: string;
  // v19.0: document = Datei-Upload; v24.25: date = Kalender-Auswahl;
  // v28.63: daterange = Übernachtungs-Zeitraum (Anreise + Abreise, Nächte berechnet)
  type: 'text' | 'select' | 'number' | 'checkbox' | 'user' | 'roommate' | 'document' | 'date' | 'daterange';
  required: boolean;
  /** v24.25: Nur für `type === 'date'` — zusätzlich die Uhrzeit abfragen. */
  withTime?: boolean;
  /** v28.63: Nur für `type === 'daterange'` — buchbares Fenster und Nächte-Limit. */
  rangeStart?: string;
  rangeEnd?: string;
  maxNights?: number;
  // Optionen als Array (incl. leerer Slots für "frisch hinzugefügte" Einträge)
  options: string[];
  visible: boolean;
  externalLinks?: Array<{ label: string; url: string }>;
  /** v18.41: Nur People-Picker (user/roommate): ausgewählte Person bei
   *  An-/Abmelde-Mail auf CC setzen (nicht im Outlook-Termin). */
  ccOnEmails?: boolean;
  /** v26.60: Nur roommate — false schaltet die separate
   *  „Zimmerpartner-Anfrage"-Mail ab (undefined = an). */
  notifyRoommate?: boolean;
  /** v29.40: Nur user/roommate — true begrenzt die Personensuche im Feld auf
   *  den Verteilerkreis des Events (siehe EventSpecificField.audienceOnly). */
  audienceOnly?: boolean;
  /** v7.11: Bei type=select erlaubt true Mehrfachauswahl (Checkbox-Liste statt
   *  Single-Dropdown). Wert wird " | "-getrennt gespeichert. */
  multi?: boolean;
  /** v26.74: Vorauswahl bei Single-Select — eine der `options` ist im
   *  Anmeldeformular vorausgewählt (leer = keine Vorauswahl). */
  defaultValue?: string;
  /** v26.75: Vorfilter — Kategorie pro Option (positional zu `options`). */
  optionCategories?: string[];
  /** v26.75: Beschriftung des Vorfilter-Dropdowns. */
  prefilterLabel?: string;
  /** v7.20: Optionale Beschreibung — landet als "i"-Tooltip neben dem
   *  Feld-Label im Registrierungsformular. */
  helpText?: string;
  /** v18.18: 'tooltip' (Default) = "i"-Hover-Box neben dem Label;
   *  'inline' = nicht-fetter Erklär-Text direkt unter dem Label. */
  helpTextStyle?: 'tooltip' | 'inline';
  /** v7.21: Sichtbarkeitsbedingung — Feld nur anzeigen wenn das Quell-Feld
   *  einen der `values` als Antwort hat. */
  showIf?: { fieldId: string; values: string[] };
  /** v10.24: Bei aktiver Split-Capacity Feld nur für eine der zwei
   *  Gruppen sichtbar machen ('A' = Durchstarter / Gruppe A, 'B' =
   *  Funstarter / Gruppe B). 'all' / undefined = beide Gruppen. */
  onlyForGroup?: 'all' | 'A' | 'B';
  /** v11.94: Nur für type='checkbox' — Text neben der Checkbox im
   *  Registrierungsformular (Default „Ja, bestätigen" / „Yes, confirm"). */
  confirmLabel?: string;
  /** v17.20: Englische Varianten — nur relevant wenn der Organizer im
   *  selben Schritt 5 den Toggle „Deutsch und Englisch ermöglichen"
   *  gesetzt hat. */
  labelEn?: string;
  helpTextEn?: string;
  confirmLabelEn?: string;
  optionsEn?: string[];
}
