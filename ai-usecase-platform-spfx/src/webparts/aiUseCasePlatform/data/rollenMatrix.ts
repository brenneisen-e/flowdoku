/**
 * Die Rechte-Matrix: Was darf welche Rolle?
 *
 * Eine reine ANSICHT. Sie steuert nichts — welche Rolle etwas darf, entscheidet
 * der Code an der Stelle selbst (`isAdmin` / `isOrganizer` aus dem
 * `RoleContext`; `ManagePage`, `LogPage` und `RolePage` prüfen dagegen). Diese
 * Datei beschreibt das nur, damit ein Admin nicht im Quelltext nachlesen muss,
 * was er mit „Use Case Organizer" vergibt.
 *
 * Genau deshalb steht hier nichts, was die App nicht tut. In DEX ist die
 * Matrix eine Sammlung von rund hundert Zeilen geworden, von denen einige
 * Funktionen beschreiben, die es längst nicht mehr gibt — eine Matrix, die
 * mehr verspricht als der Code hält, ist schlimmer als keine. Wer hier eine
 * Zeile ergänzt, prüft sie an der Stelle, die sie behauptet, und wer im Code
 * eine Prüfung ändert, ändert die Zeile im selben Commit.
 *
 * Die Zeilen der Kategorie „SharePoint" müssen zu `RECHTE_SOLL` in
 * `SharePointService.ts` passen — das ist die Tabelle, gegen die „Rechte
 * prüfen" liest. Bewusst zwei Stellen und nicht eine: Der Service kennt
 * Rechte-Stufen als Zahlen, diese Datei kennt Sätze.
 */

/** `true`/`false` = ja/nein; ein Text = ja, aber mit Einschränkung. */
export type MatrixWert = boolean | { de: string; en: string };

export interface MatrixZeile {
  key: string;
  de: string;
  en: string;
  descDe: string;
  descEn: string;
  user: MatrixWert;
  organizer: MatrixWert;
  admin: MatrixWert;
}

export interface MatrixKategorie {
  key: string;
  de: string;
  en: string;
  zeilen: MatrixZeile[];
}

/** Die Spalten der Matrix, in der Reihenfolge ihrer Rechte (niedrigste zuerst). */
export const MATRIX_SPALTEN: Array<'user' | 'organizer' | 'admin'> = ['user', 'organizer', 'admin'];

export const RECHTE_MATRIX: MatrixKategorie[] = [
  {
    key: 'ansehen',
    de: 'Use Cases ansehen und starten',
    en: 'View and launch use cases',
    zeilen: [
      {
        key: 'kachelwand',
        de: 'Kachelwand ansehen, suchen und filtern',
        en: 'Browse, search and filter the tile wall',
        descDe: 'Suche (bindestrich- und umlautunabhängig), Bereichsfilter und „Nur aufrufbare“.',
        descEn: 'Search (hyphen and umlaut insensitive), area filter and “Only callable”.',
        user: true, organizer: true, admin: true,
      },
      {
        key: 'archiviert',
        de: 'Archivierte Use Cases in der Kachelwand sehen',
        en: 'See archived use cases in the tile wall',
        descDe: 'Für User blenden Kachelwand und Suche sie aus. Wer den direkten Link kennt, sieht die Detailseite trotzdem.',
        descEn: 'The tile wall and search hide them from users. Anyone who knows the direct link still sees the detail page.',
        user: false, organizer: true, admin: true,
      },
      {
        key: 'detail',
        de: 'Detailseite öffnen (Beschreibung, Bewertung, Ressourcen)',
        en: 'Open the detail page (description, assessment, resources)',
        descDe: 'Source Code, Deployment-Guide, Wiki und Video sind Links — sie öffnen nur, was der Zielort selbst erlaubt.',
        descEn: 'Source code, deployment guide, wiki and video are links — they only open what the target itself allows.',
        user: true, organizer: true, admin: true,
      },
      {
        key: 'starten',
        de: 'Demo starten (neues Fenster oder eingebettet)',
        en: 'Start a demo (new window or embedded)',
        descDe: 'Nur wenn der Use Case auf „Live“ steht und ein Deployment-Link hinterlegt ist.',
        descEn: 'Only if the use case is “Live” and a deployment link is stored.',
        user: true, organizer: true, admin: true,
      },
      {
        key: 'sprache',
        de: 'Sprache der Oberfläche wechseln (DE / EN)',
        en: 'Switch the interface language (DE / EN)',
        descDe: 'Gilt nur für diese Person und diesen Browser.',
        descEn: 'Applies to this person and this browser only.',
        user: true, organizer: true, admin: true,
      },
    ],
  },
  {
    key: 'pflegen',
    de: 'Use Cases anlegen und pflegen',
    en: 'Create and maintain use cases',
    zeilen: [
      {
        key: 'pflegeseite',
        de: 'Pflegeseite öffnen (Use Cases pflegen)',
        en: 'Open the management page (maintain use cases)',
        descDe: 'Die Pflegeseite ist nur für Use Case Organizer und Admins erreichbar; für User gibt es dort keinen Weg hinein.',
        descEn: 'The management page is only reachable for use case organizers and admins; there is no way in for users.',
        user: false, organizer: true, admin: true,
      },
      {
        key: 'anlegen',
        de: 'Use Case anlegen',
        en: 'Create a use case',
        descDe: 'Titel, Beschreibung, Bereich, Status, Ressourcen-Links, Aufrufart, Bewertung und Schlagworte.',
        descEn: 'Title, description, area, status, resource links, launch mode, assessment and keywords.',
        user: false, organizer: true, admin: true,
      },
      {
        key: 'bearbeiten',
        de: 'Use Case bearbeiten',
        en: 'Edit a use case',
        descDe: 'Auch fremde Use Cases: Betreuer eines Use Cases geben keine eigenen Rechte.',
        descEn: 'Including use cases of others: maintainers of a use case carry no rights of their own.',
        user: false, organizer: true, admin: true,
      },
      {
        key: 'loeschen',
        de: 'Use Case löschen',
        en: 'Delete a use case',
        descDe: 'Landet im SharePoint-Papierkorb. Jeder Use Case Organizer darf jeden Use Case löschen.',
        descEn: 'Goes to the SharePoint recycle bin. Every use case organizer may delete every use case.',
        user: false, organizer: true, admin: true,
      },
      {
        key: 'startbestand',
        de: 'Startbestand aus dem Konzept anlegen',
        en: 'Create the starter set from the concept',
        descDe: 'Der Knopf erscheint, wenn die Kachelwand leer ist.',
        descEn: 'The button appears when the tile wall is empty.',
        user: false, organizer: true, admin: true,
      },
    ],
  },
  {
    key: 'bilder',
    de: 'Kachelbilder',
    en: 'Tile images',
    zeilen: [
      {
        key: 'bild-sehen',
        de: 'Kachelbilder sehen',
        en: 'See tile images',
        descDe: 'Ohne Bild zeigt die Kachel das Kürzel des Titels.',
        descEn: 'Without an image the tile shows the title initials.',
        user: true, organizer: true, admin: true,
      },
      {
        key: 'bild-pflegen',
        de: 'Bild hochladen, zuschneiden, ersetzen oder entfernen',
        en: 'Upload, crop, replace or remove an image',
        descDe: 'Das Bild hängt als Anhang an der Use-Case-Zeile; das alte räumt die App nach dem Speichern weg.',
        descEn: 'The image is an attachment on the use case row; the app removes the old one after saving.',
        user: false, organizer: true, admin: true,
      },
    ],
  },
  {
    key: 'protokoll',
    de: 'Protokoll',
    en: 'Log',
    zeilen: [
      {
        key: 'protokoll-sehen',
        de: 'Protokoll ansehen (wer hat wann was geändert)',
        en: 'View the log (who changed what, and when)',
        descDe: 'Das Protokoll enthält die Adressen der Personen, die Use Cases pflegen — deshalb nicht für User.',
        descEn: 'The log contains the addresses of the people who maintain use cases — hence not for users.',
        user: false, organizer: true, admin: true,
      },
      {
        key: 'protokoll-schreiben',
        de: 'Änderungen werden automatisch protokolliert',
        en: 'Changes are logged automatically',
        descDe: 'Anlegen, Ändern und Löschen schreiben je eine Zeile — mit der Adresse der Person, die es getan hat.',
        descEn: 'Creating, changing and deleting each write a row — with the address of the person who did it.',
        user: false, organizer: true, admin: true,
      },
    ],
  },
  {
    key: 'rollen',
    de: 'Rollenverwaltung',
    en: 'Role management',
    zeilen: [
      {
        key: 'rollen-sehen',
        de: 'Rollenverwaltung öffnen',
        en: 'Open the role management',
        descDe: 'Auch diese Matrix steht nur dort.',
        descEn: 'This matrix also lives only there.',
        user: false, organizer: false, admin: true,
      },
      {
        key: 'rollen-vergeben',
        de: 'Rollen vergeben, ändern und entfernen',
        en: 'Assign, change and remove roles',
        descDe: 'Setzt und entzieht dabei die Rechte auf den Listen und meldet, wenn eines nicht gesetzt werden konnte.',
        descEn: 'Sets and revokes the rights on the lists and reports when one could not be set.',
        user: false, organizer: false, admin: true,
      },
      {
        key: 'rechte-pruefen',
        de: 'Rechte auf den Listen prüfen und nachsetzen',
        en: 'Check and re-grant the rights on the lists',
        descDe: 'Liest je Person nach, ob die Rechte wirklich gesetzt sind — ein Eintrag in der Rollenliste allein reicht nicht.',
        descEn: 'Reads back per person whether the rights are really set — an entry in the roles list alone is not enough.',
        user: false, organizer: false, admin: true,
      },
      {
        key: 'selbst',
        de: 'Die eigene Rolle ändern oder entfernen',
        en: 'Change or remove your own role',
        descDe: 'Gesperrt: Wer sich selbst herabstuft, käme nicht mehr an diese Seite. Das kann nur ein anderer Admin.',
        descEn: 'Locked: whoever downgrades themselves could no longer reach this page. Only another admin can do that.',
        user: false, organizer: false, admin: { de: 'Gesperrt', en: 'Locked' },
      },
    ],
  },
  {
    key: 'vorschau',
    de: 'Vorschau als Nutzer',
    en: 'Preview as user',
    zeilen: [
      {
        key: 'vorschau-an',
        de: 'Zwischen Organizer- und User-Ansicht wechseln',
        en: 'Switch between organizer and user view',
        descDe: 'Die User-Ansicht zeigt die Plattform, wie normale Nutzer sie sehen — ohne Pflege, Protokoll und Rollen. Sie ändert an den Rechten nichts und endet beim Neuladen.',
        descEn: 'The user view shows the platform the way regular users see it — without management, log and roles. It changes no rights and ends on reload.',
        user: false, organizer: true, admin: true,
      },
      {
        key: 'vorschau-schreiben',
        de: 'In der User-Ansicht Use Cases ändern',
        en: 'Change use cases in the user view',
        descDe: 'Gesperrt: In dieser Ansicht zählt die Person als User, das Schreiben ist damit zu.',
        descEn: 'Locked: in this view the person counts as a user, so writing is closed.',
        user: false, organizer: false, admin: false,
      },
    ],
  },
  {
    key: 'sharepoint',
    de: 'SharePoint (Rechte auf den Listen)',
    en: 'SharePoint (rights on the lists)',
    zeilen: [
      {
        key: 'sp-rollen',
        de: 'AIUC_Roles (Rollenliste)',
        en: 'AIUC_Roles (roles list)',
        descDe: 'Die Liste hat eigene Rechte. Eine Rolle wirkt nur mit Leserecht darauf — ohne es ist die Person trotz Eintrag ein normaler User. Vollzugriff hat nur ein Admin, sonst könnte sich jeder selbst hochstufen.',
        descEn: 'The list has rights of its own. A role only works with read access to it — without it the person is a regular user despite the entry. Only an admin has full control, otherwise anyone could promote themselves.',
        user: false, organizer: { de: 'Lesen', en: 'Read' }, admin: { de: 'Vollzugriff', en: 'Full control' },
      },
      {
        key: 'sp-usecases',
        de: 'AIUC_UseCases (Use-Case-Liste)',
        en: 'AIUC_UseCases (use case list)',
        descDe: 'Für User vergibt die App nichts — sie lesen über die normalen Rechte der Site.',
        descEn: 'The app grants nothing to users — they read through the normal rights of the site.',
        user: { de: 'Lesen über die Site', en: 'Read via the site' }, organizer: { de: 'Bearbeiten', en: 'Edit' }, admin: { de: 'Bearbeiten', en: 'Edit' },
      },
      {
        key: 'sp-log',
        de: 'AIUC_Log (Protokoll)',
        en: 'AIUC_Log (log)',
        descDe: 'Zum Schreiben der Protokollzeilen. Für User vergibt die App nichts und zeigt ihnen das Protokoll nicht.',
        descEn: 'For writing the log rows. The app grants users nothing and does not show them the log.',
        user: false, organizer: { de: 'Beitragen', en: 'Contribute' }, admin: { de: 'Beitragen', en: 'Contribute' },
      },
    ],
  },
];
