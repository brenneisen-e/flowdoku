// v30.66: Aus EventCreationPage.tsx ausgelagert. Der Katalog haengt allein an
// der Sprache — als Funktion statt als Konstante im Funktionskoerper, wo er bei
// jedem Render neu gebaut wurde.
import { CustomFieldInput } from '../components/wizard/customFieldInput';
import { SuggestedEntry } from '../components/wizard/wizardTypes';

export function getSuggestedFieldsCatalog(isDe: boolean): SuggestedEntry[] {
  return isDe ? [
    {
      // v22.38: Sonder-Eintrag — schaltet das Standard-Anrede-Feld an
      // (askSalutation-Flag) statt ein Custom-Field anzulegen. Wird in
      // addSelectedSuggestedFields gesondert behandelt.
      key: 'salutation', category: 'general', icon: 'Contact',
      label: 'Anrede',
      description: 'Pflicht-Dropdown Frau / Herr / Divers / Keine Angabe — erscheint über dem Vornamen',
      build: (n) => ({ id: `cf-${n}`, label: 'Anrede', type: 'select', required: true, options: [], visible: true }),
    },
    {
      key: 'tshirt', category: 'general', icon: 'Tag',
      label: 'T-Shirt Größe',
      description: 'Dropdown mit Kein T-Shirt / XS–XXL',
      build: (n) => ({ id: `cf-${n}`, label: 'T-Shirt Größe', type: 'select', required: false, options: ['Habe bereits ein T-Shirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'allergies', category: 'general', icon: 'Warning',
      label: 'Allergien',
      description: 'Freitextfeld für Allergien/Unverträglichkeiten',
      build: (n) => ({ id: `cf-${n}`, label: 'Allergien', type: 'text', required: false, options: [], visible: true }),
    },
    {
      key: 'diet', category: 'general', icon: 'EatDrink',
      label: 'Essenspräferenzen',
      description: 'Dropdown: Keine Präferenzen / Vegetarisch / Vegan / Pescetarisch',
      build: (n) => ({ id: `cf-${n}`, label: 'Essenspräferenzen', type: 'select', required: false, options: ['Keine Präferenzen', 'Vegetarisch', 'Vegan', 'Pescetarisch'], visible: true }),
    },
    {
      key: 'hotel', category: 'general', icon: 'Hotel',
      label: 'Hotel benötigt',
      description: 'Checkbox: Teilnehmer benötigt ein Hotel',
      build: (n) => ({ id: `cf-${n}`, label: 'Hotel benötigt', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'roomtype', category: 'general', icon: 'Room',
      label: 'Zimmerart',
      description: 'Dropdown: Keine Präferenz / Einzelzimmer / Doppelzimmer',
      build: (n) => ({ id: `cf-${n}`, label: 'Zimmerart (falls Hotel benötigt)', type: 'select', required: false, options: ['Keine Präferenz', 'Einzelzimmer', 'Doppelzimmer'], visible: true }),
    },
    {
      key: 'roommate', category: 'general', icon: 'People',
      label: 'Bevorzugter Zimmerpartner',
      description: 'Personen-Suche; Match-Erkennung im Admin Center',
      build: (n) => ({ id: `cf-${n}`, label: 'Bevorzugter Zimmerpartner (bei Doppelzimmer)', type: 'roommate', required: false, options: [], visible: true }),
    },
    // B2Run-Pakete — nur für Lauf-Events relevant. Sektion ist im Modal
    // standardmäßig eingeklappt, damit der Standard-Organizer sie nicht
    // versehentlich aktiviert.
    {
      key: 'b2run_startblock', category: 'b2run', icon: 'Running',
      label: 'Startblock',
      description: 'Dropdown der Startblöcke. Optionen werden nachträglich im Wizard gepflegt.',
      build: (_n) => ({ id: `b2run_startblock`, label: 'Startblock', type: 'select', required: true, options: [], visible: true }),
    },
    {
      key: 'b2run_gruppe', category: 'b2run', icon: 'BulletedList',
      label: 'Gruppe',
      description: 'Dropdown: offene Klasse / Nordic Walker / Damen / Herren',
      build: (_n) => ({ id: `b2run_gruppe`, label: 'Gruppe', type: 'select', required: true, options: ['offene Klasse', 'Nordic Walker', 'Damen', 'Herren'], visible: true }),
    },
    {
      key: 'b2run_altersklasse', category: 'b2run', icon: 'Calendar',
      label: 'Altersklasse',
      description: 'Dropdown: unter 18 / 18-29 / 30-39 / 40-49 / 50-59 / 60+',
      build: (_n) => ({ id: `b2run_altersklasse`, label: 'Altersklasse', type: 'select', required: true, options: ['unter 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true }),
    },
    {
      key: 'b2run_infoservice', category: 'b2run', icon: 'CellPhone',
      label: 'Infoservice (SMS)',
      description: 'Checkbox: aktiviert die Mobilnummer-Pflicht für den B2Run-SMS-Service',
      build: (_n) => ({ id: `b2run_infoservice`, label: 'Infoservice nutzen (SMS von B2Run — Mobilnummer erforderlich)', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_mobilnummer', category: 'b2run', icon: 'Phone',
      label: 'Mobilnummer',
      description: 'Freitext, dynamisch Pflicht wenn Infoservice aktiv',
      build: (_n) => ({ id: `b2run_mobilnummer`, label: 'Mobilnummer (nur bei aktiviertem Infoservice)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } }),
    },
    {
      key: 'b2run_anonym', category: 'b2run', icon: 'Hide3',
      label: 'Anonym teilnehmen',
      description: 'Checkbox: Teilnehmer in Ergebnislisten anonymisieren',
      build: (_n) => ({ id: `b2run_anonym`, label: 'Anonym teilnehmen', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_laufshirt', category: 'b2run', icon: 'Sport',
      label: 'Deloitte-Laufshirt',
      description: 'Dropdown: vorhandenes Shirt / XS-XXL',
      build: (_n) => ({ id: `b2run_laufshirt`, label: 'Deloitte-Laufshirt', type: 'select', required: true, options: ['Habe bereits ein Laufshirt', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'b2run_datenschutz', category: 'b2run', icon: 'LockShield',
      label: 'AGB / Datenschutz',
      description: 'Pflicht-Checkbox mit Links zu B2Run-AGB und Datenschutzerklärung',
      build: (_n) => ({
        id: `b2run_datenschutz`,
        label: 'Zustimmung AGB, Datenschutz & Bildaufnahmen',
        type: 'checkbox', required: true, options: [], visible: true,
        externalLinks: [
          { label: 'AGB (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
          { label: 'Datenschutz (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
        ],
      }),
    },
  ] : [
    {
      key: 'salutation', category: 'general', icon: 'Contact',
      label: 'Salutation',
      description: 'Required dropdown Mrs / Mr / Diverse / Prefer not to say — shown above the first name',
      build: (n) => ({ id: `cf-${n}`, label: 'Salutation', type: 'select', required: true, options: [], visible: true }),
    },
    {
      key: 'tshirt', category: 'general', icon: 'Tag',
      label: 'T-Shirt size',
      description: 'Dropdown: No t-shirt needed / XS–XXL',
      build: (n) => ({ id: `cf-${n}`, label: 'T-Shirt size', type: 'select', required: false, options: ['I already have one', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'allergies', category: 'general', icon: 'Warning',
      label: 'Allergies',
      description: 'Free-text field for allergies / intolerances',
      build: (n) => ({ id: `cf-${n}`, label: 'Allergies', type: 'text', required: false, options: [], visible: true }),
    },
    {
      key: 'diet', category: 'general', icon: 'EatDrink',
      label: 'Dietary preferences',
      description: 'Dropdown: No preference / Vegetarian / Vegan / Pescetarian',
      build: (n) => ({ id: `cf-${n}`, label: 'Dietary preferences', type: 'select', required: false, options: ['No preference', 'Vegetarian', 'Vegan', 'Pescetarian'], visible: true }),
    },
    {
      key: 'hotel', category: 'general', icon: 'Hotel',
      label: 'Hotel required',
      description: 'Checkbox: participant needs a hotel room',
      build: (n) => ({ id: `cf-${n}`, label: 'Hotel required', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'roomtype', category: 'general', icon: 'Room',
      label: 'Room type',
      description: 'Dropdown: No preference / Single room / Double room',
      build: (n) => ({ id: `cf-${n}`, label: 'Room type (if hotel needed)', type: 'select', required: false, options: ['No preference', 'Single room', 'Double room'], visible: true }),
    },
    {
      key: 'roommate', category: 'general', icon: 'People',
      label: 'Preferred roommate',
      description: 'People search; match detection in the admin center',
      build: (n) => ({ id: `cf-${n}`, label: 'Preferred roommate (for double room)', type: 'roommate', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_startblock', category: 'b2run', icon: 'Running',
      label: 'Start block',
      description: 'Dropdown of start blocks. Options are added later in the wizard.',
      build: (_n) => ({ id: `b2run_startblock`, label: 'Start block', type: 'select', required: true, options: [], visible: true }),
    },
    {
      key: 'b2run_gruppe', category: 'b2run', icon: 'BulletedList',
      label: 'Category',
      description: 'Dropdown: Open class / Nordic Walker / Women / Men',
      build: (_n) => ({ id: `b2run_gruppe`, label: 'Category', type: 'select', required: true, options: ['Open class', 'Nordic Walker', 'Women', 'Men'], visible: true }),
    },
    {
      key: 'b2run_altersklasse', category: 'b2run', icon: 'Calendar',
      label: 'Age group',
      description: 'Dropdown: under 18 / 18-29 / 30-39 / 40-49 / 50-59 / 60+',
      build: (_n) => ({ id: `b2run_altersklasse`, label: 'Age group', type: 'select', required: true, options: ['under 18', '18-29', '30-39', '40-49', '50-59', '60+'], visible: true }),
    },
    {
      key: 'b2run_infoservice', category: 'b2run', icon: 'CellPhone',
      label: 'Info service (SMS)',
      description: 'Checkbox: enables the mandatory mobile-number for the B2Run SMS service',
      build: (_n) => ({ id: `b2run_infoservice`, label: 'Use B2Run info service (SMS — mobile number required)', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_mobilnummer', category: 'b2run', icon: 'Phone',
      label: 'Mobile number',
      description: 'Free text, dynamically required when info service is active',
      build: (_n) => ({ id: `b2run_mobilnummer`, label: 'Mobile number (only if info service is enabled)', type: 'text', required: true, options: [], visible: true, showIf: { fieldId: 'b2run_infoservice', values: ['true'] } }),
    },
    {
      key: 'b2run_anonym', category: 'b2run', icon: 'Hide3',
      label: 'Anonymous participation',
      description: 'Checkbox: anonymise attendee in result lists',
      build: (_n) => ({ id: `b2run_anonym`, label: 'Participate anonymously', type: 'checkbox', required: false, options: [], visible: true }),
    },
    {
      key: 'b2run_laufshirt', category: 'b2run', icon: 'Sport',
      label: 'Deloitte running shirt',
      description: 'Dropdown: existing shirt / XS-XXL',
      build: (_n) => ({ id: `b2run_laufshirt`, label: 'Deloitte running shirt', type: 'select', required: true, options: ['I already have one', 'XS', 'S', 'M', 'L', 'XL', 'XXL'], visible: true }),
    },
    {
      key: 'b2run_datenschutz', category: 'b2run', icon: 'LockShield',
      label: 'Terms / privacy',
      description: 'Required checkbox with links to B2Run terms and privacy policy',
      build: (_n) => ({
        id: `b2run_datenschutz`,
        label: 'I agree to the terms, privacy policy and photo/video recordings',
        type: 'checkbox', required: true, options: [], visible: true,
        externalLinks: [
          { label: 'Terms (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/agb/index.html' },
          { label: 'Privacy (b2run.de)', url: 'https://www.b2run.de/run/de/de/organisation/datenschutz/datenschutz-teilnahme-an-veranstaltungen.html' },
        ],
      }),
    },
  ];
}
