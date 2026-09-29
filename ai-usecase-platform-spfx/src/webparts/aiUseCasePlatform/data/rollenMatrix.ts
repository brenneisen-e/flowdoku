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
 *
 * **Ansicht und SharePoint sind zwei Dinge (Review-Fund 13).** Ein ✗ in der
 * Spalte „User" kann heißen „SharePoint verweigert es" ODER „die Oberfläche
 * zeigt es nicht". Das zweite schützt nichts: Wer über die Site Rechte auf der
 * Liste hat (Site-Mitglieder haben üblicherweise Bearbeiten), tut es direkt in
 * SharePoint. Zeilen, bei denen NUR die Oberfläche sperrt, tragen
 * `nurOberflaeche: true` und werden in der Rollenverwaltung so gekennzeichnet.
 * Wer eine Zeile ergänzt, entscheidet das ausdrücklich — und beantwortet die
 * Frage „was passiert, wenn jemand die Oberfläche umgeht?".
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
  /**
   * true = NUR die Oberfläche der App steuert das; SharePoint erzwingt es nicht.
   * Die Zeile sagt dann, was die App tut — nicht, was jemand mit Rechten auf der
   * Liste in SharePoint selbst könnte.
   */
  nurOberflaeche?: boolean;
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
        descDe: 'Suche (bindestrich- und umlautunabhängig), Bereichsfilter und „Nur aufrufbare“ (zeigt die Use Cases mit Status Live). Lesen kann jede Person mit Leserecht auf der Liste.',
        descEn: 'Search (hyphen and umlaut insensitive), area filter and “Only callable” (shows use cases with status Live). Anyone with read access to the list can read.',
        user: true, organizer: true, admin: true,
      },
      {
        key: 'archiviert',
        de: 'Archivierte Use Cases sehen',
        en: 'See archived use cases',
        descDe: 'Für User blenden Kachelwand, Suche und Detailseite sie aus; der direkte Link führt zu „Dieser Use Case ist nicht (mehr) da“. Das ist reine Oberfläche: Die Zeile ist in SharePoint für jede Person mit Leserecht lesbar — als Schutz vertraulicher Inhalte taugt das nicht.',
        descEn: 'For users the tile wall, search and detail page hide them; the direct link leads to “This use case does not exist (any more)”. This is interface only: the row is readable in SharePoint for anyone with read access — it is no protection for confidential content.',
        user: false, organizer: true, admin: true, nurOberflaeche: true,
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
        descDe: 'Die Seite prüft die Rolle selbst: Use Case Organizer und Admins arbeiten dort, für User zeigt sie nur den Hinweis „Dafür fehlt dir die Berechtigung“ mit dem Weg, die Rolle anzufragen.',
        descEn: 'The page checks the role itself: use case organizers and admins work there; for users it only shows the note “You are not allowed to do this” with the way to request the role.',
        user: false, organizer: true, admin: true, nurOberflaeche: true,
      },
      {
        key: 'anlegen',
        de: 'Use Case anlegen',
        en: 'Create a use case',
        descDe: 'Titel, Beschreibung, Bereich, Status, Ressourcen-Links, Aufrufart, Bewertung und Schlagworte. Die App sperrt das für User; in SharePoint darf es jede Person mit Bearbeiten-Recht auf der Liste direkt (bei Site-Mitgliedern üblich).',
        descEn: 'Title, description, area, status, resource links, launch mode, assessment and keywords. The app locks this for users; in SharePoint anyone with edit rights on the list can do it directly (usual for site members).',
        user: false, organizer: true, admin: true, nurOberflaeche: true,
      },
      {
        key: 'bearbeiten',
        de: 'Use Case bearbeiten',
        en: 'Edit a use case',
        descDe: 'Auch fremde Use Cases: Betreuer eines Use Cases geben keine eigenen Rechte. Nur die App sperrt das für User; SharePoint erlaubt es jeder Person mit Bearbeiten-Recht auf der Liste.',
        descEn: 'Including use cases of others: maintainers of a use case carry no rights of their own. Only the app locks this for users; SharePoint allows it for anyone with edit rights on the list.',
        user: false, organizer: true, admin: true, nurOberflaeche: true,
      },
      {
        key: 'loeschen',
        de: 'Use Case löschen',
        en: 'Delete a use case',
        descDe: 'Landet im SharePoint-Papierkorb. Jeder Use Case Organizer darf jeden Use Case löschen. Nur die App sperrt das für User; SharePoint erlaubt es jeder Person mit Bearbeiten-Recht auf der Liste.',
        descEn: 'Goes to the SharePoint recycle bin. Every use case organizer may delete every use case. Only the app locks this for users; SharePoint allows it for anyone with edit rights on the list.',
        user: false, organizer: true, admin: true, nurOberflaeche: true,
      },
      {
        key: 'startbestand',
        de: 'Startbestand aus dem Konzept anlegen',
        en: 'Create the starter set from the concept',
        descDe: 'Der Knopf erscheint, wenn die Kachelwand leer ist. Die automatische Erstbefüllung beim ersten Start läuft nur, wenn ein Organizer oder Admin die Plattform öffnet — und erst, nachdem die Rollen geladen sind —, nicht beim Öffnen durch irgendwen.',
        descEn: 'The button appears when the tile wall is empty. The automatic initial fill on first start only runs when an organizer or admin opens the platform — and only after the roles have loaded — not when just anyone opens it.',
        user: false, organizer: true, admin: true, nurOberflaeche: true,
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
        descDe: 'Das Bild hängt als Anhang an der Use-Case-Zeile; das alte räumt die App nach dem Speichern weg. Wer die Zeile bearbeiten darf, darf auch Anhänge ändern — in SharePoint also jede Person mit Bearbeiten-Recht auf der Liste.',
        descEn: 'The image is an attachment on the use case row; the app removes the old one after saving. Whoever may edit the row may also change attachments — in SharePoint that is anyone with edit rights on the list.',
        user: false, organizer: true, admin: true, nurOberflaeche: true,
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
        descDe: 'Das Protokoll enthält die Adressen der Personen, die Use Cases pflegen — deshalb zeigt die App es User nicht. Das ist reine Oberfläche: AIUC_Log erbt (bzw. kopiert) die Rechte der Site, und jede Person mit Leserecht liest die Zeilen samt Adressen über die Listen-Adresse in SharePoint.',
        descEn: 'The log contains the addresses of the people who maintain use cases — hence the app does not show it to users. This is interface only: AIUC_Log inherits (or copies) the site’s rights, and anyone with read access reads the rows including addresses via the list address in SharePoint.',
        user: false, organizer: true, admin: true, nurOberflaeche: true,
      },
      {
        key: 'protokoll-schreiben',
        de: 'Änderungen werden automatisch protokolliert',
        en: 'Changes are logged automatically',
        descDe: 'Anlegen, Ändern (nur wenn sich wirklich etwas geändert hat) und Löschen schreiben je eine Zeile, ebenso jede Rollenänderung (vergeben, geändert, entfernt) — mit der Adresse der Person, die es getan hat. Best-effort: Fehlt das Recht „Beitragen“ auf AIUC_Log, gibt es keine Zeile und keine Meldung; die Änderung selbst gilt trotzdem.',
        descEn: 'Creating, changing (only if something really changed) and deleting each write a row, as does every role change (assigned, changed, removed) — with the address of the person who did it. Best-effort: without “contribute” on AIUC_Log there is no row and no message; the change itself still applies.',
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
        descDe: 'Auch diese Matrix steht nur dort. Die Seite zeigt die App nur Admins; die Rollenliste selbst schützt SharePoint mit eigenen Rechten (Owners und die Vergebenen).',
        descEn: 'This matrix also lives only there. The app shows the page only to admins; the roles list itself is protected by SharePoint with rights of its own (owners and those granted).',
        user: false, organizer: false, admin: true,
      },
      {
        key: 'rollen-vergeben',
        de: 'Rollen vergeben, ändern und entfernen',
        en: 'Assign, change and remove roles',
        descDe: 'Setzt und entzieht dabei die Rechte auf den Listen und meldet, wenn eines nicht gesetzt werden konnte. Beim Herabstufen und Entfernen werden die Rechte ZUERST entzogen (nachgelesen bestätigt), erst dann ändert sich die Zeile; scheitert der Entzug, bleibt die Zeile unverändert (einzelne Listen können schon entzogen sein — „Rechte prüfen“ zeigt den Stand). Dafür hat ein Admin Vollzugriff auf alle drei Listen. Steht dieselbe Adresse mehrfach in der Liste, gilt die erste Zeile.',
        descEn: 'Sets and revokes the rights on the lists and reports when one could not be set. When downgrading or removing, the rights are revoked FIRST (confirmed by reading back), only then does the row change; if the revocation fails, the row stays unchanged (some lists may already be revoked — “Check rights” shows the state). For this an admin has full control on all three lists. If the same address is listed more than once, the first row applies.',
        user: false, organizer: false, admin: true,
      },
      {
        key: 'rechte-pruefen',
        de: 'Rechte auf den Listen prüfen und nachsetzen',
        en: 'Check and re-grant the rights on the lists',
        descDe: 'Liest je Person nach, ob die Rechte wirklich gesetzt sind — ein Eintrag in der Rollenliste allein reicht nicht. Rechte über die Owners-/Members-Gruppe der Liste zählen als gedeckt („über Gruppe gedeckt“). Nachsetzen greift nur bei Zeilen, deren Rolle sich seit der Prüfung nicht geändert hat.',
        descEn: 'Reads back per person whether the rights are really set — an entry in the roles list alone is not enough. Rights through the owners/members group of the list count as covered (“covered by group”). Re-granting only applies to rows whose role has not changed since the check.',
        user: false, organizer: false, admin: true,
      },
      {
        key: 'selbst',
        de: 'Die eigene Rolle ändern oder entfernen',
        en: 'Change or remove your own role',
        descDe: 'Gesperrt: Wer sich selbst herabstuft, käme nicht mehr an diese Seite. Das kann nur ein anderer Admin. Die Sperre gilt für das Konto, nicht für die Schreibweise — auch eine Zeile mit einem Alias derselben Person ist gesperrt, und der Entzug verweigert sich selbst.',
        descEn: 'Locked: whoever downgrades themselves could no longer reach this page. Only another admin can do that. The lock applies to the account, not the spelling — a row with an alias of the same person is locked too, and the revocation refuses itself.',
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
        user: false, organizer: true, admin: true, nurOberflaeche: true,
      },
      {
        key: 'vorschau-schreiben',
        de: 'In der User-Ansicht Use Cases ändern',
        en: 'Change use cases in the user view',
        descDe: 'Gesperrt: In dieser Ansicht zählt die Person als User, das Schreiben ist damit in der App zu. Ihre SharePoint-Rechte bleiben unverändert.',
        descEn: 'Locked: in this view the person counts as a user, so writing is closed in the app. Their SharePoint rights stay unchanged.',
        user: false, organizer: false, admin: false, nurOberflaeche: true,
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
        descDe: 'Die Liste hat eigene Rechte (Owners plus die Vergebenen). Eine Rolle wirkt nur mit Leserecht darauf — ohne es ist die Person trotz Eintrag ein normaler User. Vollzugriff hat nur ein Admin, sonst könnte sich jeder selbst hochstufen. Ein gewöhnlicher User kann die Liste nicht lesen; das ist der Normalfall und kein Fehler.',
        descEn: 'The list has rights of its own (owners plus those granted). A role only works with read access to it — without it the person is a regular user despite the entry. Only an admin has full control, otherwise anyone could promote themselves. A regular user cannot read the list; that is the normal case, not an error.',
        user: false, organizer: { de: 'Lesen', en: 'Read' }, admin: { de: 'Vollzugriff', en: 'Full control' },
      },
      {
        key: 'sp-usecases',
        de: 'AIUC_UseCases (Use-Case-Liste)',
        en: 'AIUC_UseCases (use case list)',
        descDe: 'Für User vergibt die App nichts — sie lesen über die normalen Rechte der Site. Ein Admin braucht Vollzugriff, weil Rechte vergeben, entziehen und lesen auf dieser Liste „Berechtigungen verwalten“ verlangt. ACHTUNG: Die erste Vergabe kappt die Vererbung dieser Liste von der Site (die bisherigen Zuweisungen bleiben als Kopie). Rechte, die später am Web vergeben werden — eine neue Gruppe, „Everyone except external users“ —, wirken hier nicht mehr; neue Leser müssen dann auf der Liste selbst eingetragen werden.',
        descEn: 'The app grants nothing to users — they read through the normal rights of the site. An admin needs full control because granting, revoking and reading rights on this list requires “manage permissions”. NOTE: the first grant breaks this list’s inheritance from the site (the existing assignments stay as a copy). Rights granted on the web later — a new group, “Everyone except external users” — no longer reach it; new readers must then be added on the list itself.',
        user: { de: 'Lesen über die Site', en: 'Read via the site' }, organizer: { de: 'Bearbeiten', en: 'Edit' }, admin: { de: 'Vollzugriff', en: 'Full control' },
      },
      {
        key: 'sp-log',
        de: 'AIUC_Log (Protokoll)',
        en: 'AIUC_Log (log)',
        descDe: 'Zum Schreiben der Protokollzeilen. Für User vergibt die App nichts und zeigt ihnen das Protokoll nicht — SharePoint schützt es aber nicht gegen Lesen (die Liste erbt bzw. kopiert die Rechte der Site). Ein Admin braucht Vollzugriff (Rechte verwalten). ACHTUNG: Auch hier kappt die erste Vergabe die Vererbung von der Site (Kopie der bisherigen Zuweisungen); später am Web vergebene Rechte wirken dort nicht mehr.',
        descEn: 'For writing the log rows. The app grants users nothing and does not show them the log — but SharePoint does not protect it against reading (the list inherits or copies the site’s rights). An admin needs full control (manage rights). NOTE: here too the first grant breaks inheritance from the site (a copy of the existing assignments); rights granted on the web later no longer reach it.',
        user: false, organizer: { de: 'Beitragen', en: 'Contribute' }, admin: { de: 'Vollzugriff', en: 'Full control' },
      },
    ],
  },
];
