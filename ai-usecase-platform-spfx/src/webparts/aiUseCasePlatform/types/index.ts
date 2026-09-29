/**
 * Die Datentypen der AI Use Case Platform.
 *
 * Das Modell kommt aus dem Konzept-Deck „Agentic Banking Demo Hub":
 * je Demo eine Kurzbeschreibung, ein Aufruf und fuenf Ressourcen
 * (Source Code, Deployment-Link, Deployment-Guide, Wiki/Doku, Video),
 * dazu die drei Bewertungsdimensionen der Use-Case-Analyse.
 */

/**
 * Rollen. Aufbau wie DEX: eine Zeile je Person in der Rollenliste.
 *
 * `Organizer` ist der „Use Case Organizer": legt Use Cases an und pflegt sie.
 * In der Liste steht dafür weiterhin `Kurator` (Altbestand) — die Umsetzung in
 * beide Richtungen steht in `utils/rollen.ts` und nirgends sonst.
 */
export type UserRole = 'Admin' | 'Organizer' | 'User';

export interface RoleAssignment {
  id: number;
  userEmail: string;
  userName: string;
  role: UserRole;
  assignedBy: string;
  assignedDate: string;
}

/**
 * Bewertung einer Dimension. Das Deck kennt „Hoch" und „Mittel"; „Niedrig"
 * ist ergaenzt, weil eine Skala ohne unteres Ende beim ersten schwachen
 * Kandidaten zur Luege wird.
 *
 * `unbewertet` ist ein eigener Wert und NICHT dasselbe wie „niedrig" — ein
 * frisch angelegter Use Case ist nicht schlecht bewertet, er ist noch nicht
 * bewertet. (Dieselbe Unterscheidung, an der DEX gelernt hat, dass ein
 * Lesefehler keine Null ist.)
 */
export type Bewertung = 'Hoch' | 'Mittel' | 'Niedrig' | 'unbewertet';

/**
 * Der Lebenszyklus einer Demo. Vor dem Hackathon existiert der Use Case als
 * Idee, danach als lauffaehige Demo — die Plattform muss beides zeigen
 * koennen, sonst steht sie am Tag der Veranstaltung leer da.
 */
export type UseCaseStatus =
  /** Geplant, noch keine Demo. Sichtbar, aber nicht aufrufbar. */
  | 'Geplant'
  /** Wird gerade gebaut (z.B. waehrend des Hackathons). */
  | 'InArbeit'
  /** Fertig und aufrufbar. */
  | 'Live'
  /** Ausgelaufen — bleibt fuer die Historie, aus der Kachelwand verschwunden. */
  | 'Archiviert';

/**
 * Wie eine Demo geoeffnet wird. Nutzer-Entscheidung 10.09.2026:
 * je Use Case waehlbar.
 *
 * Warum das eine echte Entscheidung ist und keine Kosmetik: Permissions
 * Policy reicht Rechte nur ABWAERTS durch. Was das WebView der Host-App nicht
 * hat, kann keine eingebettete Seite gewinnen — in DEX sind daran fuenf
 * Anlaeufe fuer den Kamera-Scan gescheitert. Eine Demo mit Kamera, Mikrofon
 * oder Zwischenablage MUSS `fenster` sein. Dazu kommt, dass viele Ziele das
 * Einbetten selbst per `X-Frame-Options` verbieten — das sieht man erst im
 * Kundentermin, deshalb ist `fenster` die Vorgabe.
 */
export type AufrufArt = 'fenster' | 'eingebettet';

/** Die fuenf Ressourcen je Demo aus dem Konzept-Deck. */
export interface UseCaseRessourcen {
  /** Repository, meist GitHub. */
  sourceCode: string;
  /** Die laufende Demo. Das Ziel des Kachel-Klicks. */
  deployment: string;
  /** Wie man die Demo selbst aufsetzt. */
  deploymentGuide: string;
  /** Fachliche Dokumentation. */
  wiki: string;
  /** Aufzeichnung, wenn die Live-Demo nicht geht. */
  video: string;
}

export interface UseCase {
  /** SharePoint-Item-Id. */
  id: number;
  /** Der Name, der auf der Kachel steht. */
  titel: string;
  /** Ein bis zwei Saetze — was macht die Demo. Steht auf der Kachel. */
  kurzbeschreibung: string;
  /** Die lange Fassung fuer die Detailseite. HTML erlaubt. */
  beschreibung: string;
  /** Fachlicher Bereich, z.B. „Investment / Private Banking". Fuer den Filter. */
  bereich: string;
  status: UseCaseStatus;
  salesRelevanz: Bewertung;
  machbarkeit: Bewertung;
  demoTauglichkeit: Bewertung;
  aufrufArt: AufrufArt;
  ressourcen: UseCaseRessourcen;
  /** Kachelbild als Anhang-URL. Leer = die Kachel zeigt ihr Kuerzel. */
  bildUrl: string;
  /** Reihenfolge in der Kachelwand. Kleiner zuerst. */
  reihenfolge: number;
  /** Wer die Demo betreut — Anzeige, keine Berechtigung. */
  betreuerEmails: string[];
  betreuerNamen: string[];
  /** Freie Schlagworte fuer die Suche. */
  schlagworte: string[];
  geaendertAm: string;
  geaendertVon: string;
}

/** Ein Eintrag im Aenderungsprotokoll. */
export interface LogEintrag {
  id: number;
  /** 0 = gehört zu keinem einzelnen Use Case (z. B. die Erstbefüllung). */
  useCaseId: number;
  aktion: string;
  detail: string;
  wer: string;
  wann: string;
}

/**
 * Was beim Speichern mit dem Kachelbild passieren soll.
 *
 * Ein eigener Typ statt zweier Flags, weil die Kombinationen sich
 * widersprechen können („neues Bild" UND „Bild entfernen"): Der Typ lässt
 * genau eine Absicht zu.
 */
export type BildAenderung =
  | { art: 'unveraendert' }
  | { art: 'neu'; datei: File }
  | { art: 'entfernen' };

/** Ergebnis von `saveUseCase` — der Grund steht dabei, damit die Meldung ihn nennen kann. */
export interface SpeicherErgebnis {
  ok: boolean;
  /** Id des gespeicherten Use Cases (bei „neu" erst nach dem Anlegen bekannt). */
  id: number | null;
  /** Der Use Case ist gespeichert, aber das Bild nicht angekommen. */
  bildFehler: boolean;
  /** Warum `ok` false ist. */
  grund?: 'rechte' | 'speichern' | 'bild';
}
