/**
 * Rollen der AI Use Case Platform.
 *
 * Drei Rollen, Aufbau wie DEX:
 *   Admin   — alles, inklusive Rollenverwaltung
 *   Use Case Organizer — Use Cases anlegen und pflegen (in der Liste: Kurator)
 *   User    — sehen und aufrufen
 *
 * Die zwei Regeln, die DEX teuer gelernt hat, gelten hier von Anfang an:
 *
 * 1. Ein Lesefehler befoerdert niemanden. `getRoles` liefert bei 403 `null`,
 *    nicht `[]` — und nur eine FRISCH angelegte Liste macht die erste Person
 *    zum Admin. In DEX war „getRoles ist leer" einmal gleichbedeutend mit
 *    „Erstinstallation", und ein 403 machte damit jeden Aufrufer zum Admin
 *    (v6.34).
 *
 * 2. Eine Rolle wirkt nur mit Leserecht auf der Rollenliste. Deshalb setzt
 *    jede Zuweisung Read nach — und meldet, wenn das nicht geklappt hat,
 *    statt Erfolg zu behaupten (v30.80/v30.85).
 *
 * Fünf weitere Regeln aus dem Review von v1.3 — jede steht hier, weil man sie
 * beim Umbauen leicht wieder herausnimmt:
 *
 * 3. **Nie Rechte ohne Zeile zurücklassen.** Beim ENTFERNEN und beim
 *    HERABSTUFEN werden die Rechte ZUERST entzogen (nachgelesen bestätigt), erst
 *    danach wird die Zeile gelöscht bzw. geändert. Scheitert der Entzug, bleibt
 *    die Zeile, wie sie war („nichts wurde geändert"). Bis dahin stand es
 *    umgekehrt: Zeile weg, Entzug gescheitert — und „Rechte prüfen" iteriert
 *    nur Zeilen, fand die Person also nie wieder. Bei Erhöhungen ist es
 *    andersherum (erst die Zeile, dann die Rechte): Ein Fehler soll immer
 *    Richtung WENIGER Zugriff scheitern.
 * 4. **Rechte gelten je Adresse, nicht je Zeile.** Steht dieselbe Adresse
 *    zweimal in der Liste, zählt die ERSTE. Vor jedem Rechte-Schritt wird die
 *    Liste frisch gelesen (nicht der State geglaubt); wer eine überzählige
 *    Zeile ändert oder entfernt, ändert nur die Zeile und fasst die Rechte der
 *    Person nicht an.
 * 5. **Selbstschutz über das aufgelöste Konto**, nicht über die Adress-
 *    Zeichenkette: Dieselbe Person kann unter einem Alias in der Liste stehen.
 * 6. **Eine Erstinstallation, die nicht gespeichert wurde, macht niemanden zum
 *    Admin.** Sie wird gemeldet (`erstinstallation`).
 * 7. **Rollenänderungen stehen im Protokoll** (`rolle-vergeben`,
 *    `rolle-geaendert`, `rolle-entfernt`) — wer wann jemanden zum Admin gemacht
 *    hat, muss auffindbar sein.
 */

import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService, PersonenSuche, RechteBericht, RechteLuecke, RechteReparatur, RechteUeberschuss, RechteCode, LeseFehler } from '../services/SharePointService';
import { RoleAssignment, UserRole } from '../types';
import { isCurrentUser } from '../utils/sessionIdentities';
import { rolleAusSpeicher, rolleRang } from '../utils/rollen';

/**
 * Warum ein Schreibvorgang der Rollenverwaltung NICHT (oder nur teilweise) ging —
 * als Code; die Oberfläche formuliert zweisprachig.
 *  - `lesefehler`        — die Rollenliste war nicht frisch lesbar: nichts wurde geändert.
 *  - `selbst`            — das ist das eigene Konto (auch unter einer anderen Schreibweise).
 *  - `entzug`            — der Entzug der Rechte ist gescheitert oder nicht belegt: nichts wurde geändert.
 *  - `zeile`             — die Zeile ließ sich nicht schreiben.
 *  - `zeile-nach-entzug` — die Rechte sind entzogen, die Zeile ließ sich danach nicht ändern.
 *  - `nicht-gefunden`    — die Zeile gibt es nicht (mehr).
 *  - `herabstufung`      — „Person hinzufügen" würde herabstufen und tut es nie.
 *  - `nur-zeile`         — KEIN Fehler: Es war ein überzähliger Eintrag; nur die Zeile wurde geändert.
 */
export type AktionsGrund = '' | 'lesefehler' | 'selbst' | 'entzug' | 'zeile' | 'zeile-nach-entzug' | 'nicht-gefunden' | 'herabstufung' | 'nur-zeile' | 'keine-berechtigung';

/**
 * Wie die Erstinstallation ausging (leere Rollenliste, erste Person wird Admin).
 *  - `keine`             — keine Erstinstallation (Rollenliste hatte Einträge oder war nicht lesbar).
 *  - `gespeichert`       — Zeile und Rechte stehen.
 *  - `rechte-fehlen`     — Zeile steht, aber nicht alle Rechte konnten gesetzt werden (`erstinstallationFehlt`).
 *  - `nicht-gespeichert` — die Zeile ließ sich nicht schreiben: niemand wurde Admin.
 */
export type ErstinstallationsStatus = 'keine' | 'gespeichert' | 'rechte-fehlen' | 'nicht-gespeichert';

/** Ergebnis von „Fehlende Rechte nachsetzen": zusätzlich die Zeilen, deren Rolle sich seit der Prüfung geändert hat. */
export interface RepairErgebnis extends RechteReparatur {
  /** Für diese Zeilen wurde NICHTS gesetzt: Ihre Rolle stimmt nicht mehr mit dem Bericht überein. */
  rolleGeaendert: RechteLuecke[];
  /** Die Rollenliste war nicht lesbar — nichts wurde angefasst. */
  lesefehler: boolean;
}

export type UeberschussGrund = 'rolle-geaendert' | 'nicht-bestaetigt' | 'selbst' | 'lesefehler';

/** Ergebnis von „Überzählige Rechte entziehen" — mit den getrennten Zuständen aus dem Service. */
export interface EntzugsErgebnis {
  erledigt: number;
  /** Nicht entzogen (oder nicht belegt) — mit dem Grund. */
  offen: Array<RechteUeberschuss & { grund: UeberschussGrund }>;
  /** Entzogen, aber das Vorgesehene fehlt jetzt (nachgelesen) — es muss nachgesetzt werden. */
  lesenFehlt: RechteUeberschuss[];
}

interface RoleContextType {
  roles: RoleAssignment[];
  currentUserRole: UserRole;
  isRolesLoading: boolean;
  /** `forbidden` = die Rollenliste ist nicht lesbar. Wer darin steht, ist trotzdem „User". */
  rolesReadStatus: 'loading' | 'ok' | 'forbidden' | 'error';
  isAdmin: boolean;
  isOrganizer: boolean;
  /**
   * Die ECHTE Rolle, unabhängig von der Vorschau „als Nutzer ansehen".
   *
   * `isAdmin`/`isOrganizer` sinken in der Vorschau auf „User" — das ist der
   * Zweck. Das Menü, in dem man die Vorschau wieder verlässt, darf aber nicht
   * mit verschwinden, sonst kommt man aus der Ansicht nicht mehr heraus
   * (dieselbe Regel wie `originalIsAdmin` in DEX).
   */
  originalIsAdmin: boolean;
  originalIsOrganizer: boolean;
  /** Vorschau „als Nutzer sehen" — ohne Identitaetswechsel, endet beim Neuladen. */
  previewAsUser: boolean;
  setPreviewAsUser: (on: boolean) => void;
  addRole: (email: string, name: string, role: UserRole) => Promise<boolean>;
  updateRole: (itemId: number, role: UserRole) => Promise<boolean>;
  removeRole: (itemId: number) => Promise<boolean>;
  refreshRoles: () => Promise<void>;
  /** Welche Rechte beim letzten Schreiben NICHT gesetzt werden konnten (Codes — die Oberfläche formuliert). */
  lastRightsMissing: () => RechteCode[];
  /** Warum das letzte Schreiben nicht (oder nur als Zeile) ging. `''` = ohne Besonderheit. */
  lastAktionGrund: () => AktionsGrund;
  /** Wie die Erstinstallation ausging; bei `rechte-fehlen` nennt `erstinstallationFehlt` die Rechte. */
  erstinstallation: ErstinstallationsStatus;
  erstinstallationFehlt: RechteCode[];
  /** Erstinstallation noch einmal versuchen (nur sinnvoll bei `nicht-gespeichert`). */
  erstinstallationWiederholen: () => Promise<void>;
  /**
   * v1.3: Personensuche über den People-Picker von SharePoint (kein Graph).
   * `null` heißt: Die Suche ging NICHT — das ist etwas anderes als `treffer: []`.
   * Wer daraus „keine Treffer" macht, sagt einem Admin, die Person gebe es
   * nicht, weil eine Anfrage gedrosselt war.
   */
  searchUsers: (query: string, includeInternational?: boolean) => Promise<PersonenSuche | null>;
  /** Laufende Suchen für überholt erklären (Seite verlassen, neuer Suchtext). */
  cancelSearch: () => void;
  /** HTTP-Status der letzten Suche — für die Meldung, WARUM sie nicht ging. */
  lastPersonSearchStatus: () => number;
  /**
   * v1.3: „Rechte prüfen". Liest die Rollenliste frisch und danach je Person
   * nach, ob die Rechte auf den drei Listen wirklich gesetzt sind. Schreibt
   * nichts. `null` = nicht lesbar (`lastRightsReadError` sagt warum) — kein
   * „alles in Ordnung".
   */
  auditRoleRights: (onProgress?: (done: number, total: number) => void) => Promise<RechteBericht | null>;
  /** Woran der letzte Audit beim Lesen scheiterte — als Code (Liste, Art, Status), nicht als Satz. */
  lastRightsReadFehler: () => LeseFehler | null;
  /**
   * v1.3: Setzt fehlende Rechte nach — mit Wiederholung und Nachlesen, nur
   * additiv. Liest die Rollen vorher frisch: Für Zeilen, deren Rolle sich seit
   * der Prüfung geändert hat, wird nichts gesetzt (sonst bekäme eine
   * herabgestufte Person ihre Rechte zurück).
   */
  repairRoleRights: (luecken: RechteLuecke[], onProgress?: (done: number, total: number) => void) => Promise<RepairErgebnis>;
  /**
   * v1.3: Entzieht überzählige Rechte (früherer Admin mit Vollzugriff auf der
   * Rollenliste, User mit Restrechten). Prüft die Rolle vorher frisch: Hat sich
   * die Rolle seit dem Audit geändert, wird nichts angefasst.
   */
  revokeExcessRights: (items: RechteUeberschuss[], onProgress?: (done: number, total: number) => void) => Promise<EntzugsErgebnis>;
  /**
   * Ist diese Adresse die angemeldete Person? Über alle Schreibweisen des
   * Kontos (`user.email`, `loginName`) UND über das aufgelöste Konto: Steht
   * dieselbe Person unter einem Alias in der Liste, erkennt das die Auflösung
   * (im Hintergrund, für Zeilen mit demselben Namen).
   */
  isSelf: (email: string) => boolean;
  siteUrl: string;
  service: SharePointService;
}

const RoleContext = React.createContext<RoleContextType | undefined>(undefined);

// `Kurator` (Altbestand) und `Organizer` sind dieselbe Rolle — siehe utils/rollen.ts.
const normalizeRole = rolleAusSpeicher;
const roleRank = rolleRang;

/** Adresse zum Vergleichen: klein und getrimmt (die Adresse ist der Schlüssel). */
function norm(a: string | undefined | null): string {
  return (a || '').trim().toLowerCase();
}

/** Gleicher Name in beliebiger Wortfolge („Klein, Julia" = „Julia Klein")? Nur ein HINWEIS, welche Zeilen sich lohnen aufzulösen. */
function gleicherNameWie(a: string, b: string): boolean {
  const n = (s: string): string => (s || '').toLowerCase().replace(/,/g, ' ').split(/\s+/).filter(Boolean).sort().join(' ');
  const x = n(a);
  return !!x && x === n(b);
}

export function RoleProvider(props: { context: WebPartContext; children: React.ReactNode }): React.ReactElement {
  const [roles, setRoles] = React.useState<RoleAssignment[]>([]);
  const [currentUserRole, setCurrentUserRole] = React.useState<UserRole>('User');
  const [isRolesLoading, setIsRolesLoading] = React.useState(true);
  const [rolesReadStatus, setRolesReadStatus] = React.useState<'loading' | 'ok' | 'forbidden' | 'error'>('loading');
  const [previewAsUser, setPreviewAsUser] = React.useState(false);
  const [erstinstallation, setErstinstallation] = React.useState<ErstinstallationsStatus>('keine');
  // Die ECHTE Rolle für die Schreibaktionen unten (nicht die der Vorschau „als User ansehen"):
  // Sie stehen in `useCallback`s mit fester Abhängigkeitsliste und lesen den Wert deshalb aus
  // einem Ref. SharePoint hält ohnehin dagegen (Organizer haben nur Lesen auf der Rollenliste),
  // aber ohne diese Sperre stünde die Rollenverwaltung ohne Netz da, sobald die Liste erbt.
  const echterAdminRef = React.useRef(false);
  const [erstinstallationFehlt, setErstinstallationFehlt] = React.useState<RechteCode[]>([]);
  // Adressen (klein), die zum ANGEMELDETEN Konto auflösen, obwohl `isCurrentUser`
  // sie nicht erkennt (Alias in der Rollenliste). Im Ref für die Aktionen (stabile
  // Callbacks), im State für die Oberfläche (Sperre der Zeile).
  const [selbstKonten, setSelbstKonten] = React.useState<string[]>([]);
  const selbstRef = React.useRef<string[]>([]);
  const selbstGeprueft = React.useRef<Record<string, true>>({});
  const rightsMissingRef = React.useRef<RechteCode[]>([]);
  const grundRef = React.useRef<AktionsGrund>('');

  const service = React.useMemo(() => new SharePointService(props.context), []);
  const myEmail = props.context.pageContext.user.email;
  const myName = props.context.pageContext.user.displayName || myEmail;

  const mapRows = (rows: Array<{ Id: number; Title: string; UserName: string; Role: string; AssignedBy: string; AssignedDate: string }>): RoleAssignment[] =>
    rows.map(r => ({
      id: r.Id,
      userEmail: r.Title || '',
      userName: r.UserName || '',
      role: normalizeRole(r.Role),
      assignedBy: r.AssignedBy || '',
      assignedDate: r.AssignedDate || '',
    }));

  React.useEffect(() => {
    init().catch(() => setIsRolesLoading(false));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /** Die Rollenliste FRISCH lesen — nie dem State glauben, er kann veraltet sein (zweiter Admin, zweiter Tab). `null` = nicht lesbar. */
  async function frischeRollen(): Promise<RoleAssignment[] | null> {
    const rows = await service.getRoles();
    return rows === null ? null : mapRows(rows);
  }

  async function init(): Promise<void> {
    // Legt nur bei einem EINDEUTIGEN 404 an; für einen normalen User ist die
    // Rollenliste unlesbar (403) — dann schreibt es nichts (Review-Fund 11).
    await service.ensureRolesList();
    const rows = await service.getRoles();

    if (rows === null) {
      // Nicht lesbar. KEINE Rolle zuordnen — sonst umgeht ein
      // voruebergehender Fehler die gesamte Zugriffssteuerung. Lieber sieht
      // jemand zu wenig als versehentlich alles.
      setRoles([]);
      setCurrentUserRole('User');
      setRolesReadStatus(service.lastRolesReadStatus === 403 ? 'forbidden' : 'error');
      setIsRolesLoading(false);
      console.warn('[AIUC] Rollenliste nicht lesbar — Rolle bleibt „User".');
      return;
    }

    setRolesReadStatus('ok');

    if (rows.length === 0) {
      await erstinstallationDurchfuehren();
      return;
    }

    const mapped = mapRows(rows);
    setRoles(mapped);
    // Ueber ALLE Schreibweisen der angemeldeten Person, nicht nur
    // `pageContext.user.email` — steht die Zeile unter dem Alias, waere die
    // Person sonst „User" (DEX v30.81).
    const mine = mapped.filter(r => isCurrentUser(props.context, r.userEmail))[0];
    setCurrentUserRole(mine ? mine.role : 'User');
    setIsRolesLoading(false);
  }

  /**
   * Erstinstallation: Die Rollenliste ist LEER und WAR LESBAR.
   *
   * In DEX haengt dieser Zweig zusaetzlich an `isNewlyCreated`, und das aus
   * gutem Grund: Dort lieferte `getRoles` bei einem 403 ein leeres Array,
   * „leer" hiess also auch „darf nicht lesen", und jeder Aufrufer wurde Admin
   * (v6.34). Hier kann das nicht passieren: `getRoles` liefert bei JEDEM Fehler
   * `null`, und dieser Zweig ist erst hinter der `rows === null`-Pruefung
   * erreichbar. „Leer" heisst hier also wirklich leer.
   *
   * Warum das nicht mehr nur an `isNewlyCreated` haengt: Beim ersten Live-
   * Versuch am 10.09.2026 wurde die Liste angelegt, das Anlegen der Spalten
   * scheiterte still, und der Aufrufer blieb „User" — ohne Zugang zur
   * Rollenverwaltung, mit der er sich haette helfen koennen. Eine leere
   * Rollenliste ohne Admin ist eine Sackgasse.
   *
   * v1.3 (Review-Fund 12): Das Ergebnis des Speicherns wird AUSGEWERTET. Bis
   * dahin setzte der Zweig `currentUserRole = 'Admin'` samt lokaler Zeile mit
   * Id 0, auch wenn `addRole` gescheitert war — die Person sah die volle
   * Rollenverwaltung, nichts war gespeichert, und nach „Neu laden" war sie
   * ohne Meldung wieder User. Jetzt: Scheitert das Speichern, wird NIEMAND
   * Admin, und `erstinstallation` sagt „nicht-gespeichert". Die Rechte laufen
   * über dieselbe Vergabe wie für jeden Admin (Rollenliste, Use Cases UND
   * Protokoll), nicht über eine Einzelvergabe, die zwei Listen vergaß.
   */
  async function erstinstallationDurchfuehren(): Promise<void> {
    console.warn('[AIUC] Rollenliste ist leer — die angemeldete Person wird Admin (Erstinstallation).');
    const gespeichert = await service.addRole(myEmail, myName, 'Admin', 'System (Erstinstallation)');
    if (!gespeichert) {
      console.warn('[AIUC] Erstinstallation: Der Eintrag ließ sich nicht speichern — niemand wird Admin.');
      setErstinstallation('nicht-gespeichert');
      setErstinstallationFehlt([]);
      setRoles([]);
      setCurrentUserRole('User');
      // „error", nicht „ok": Die Rolle der Person ist damit NICHT bestimmt, und die
      // bestehenden Hinweise der App („Rolle konnte nicht geprüft werden") greifen.
      setRolesReadStatus('error');
      setIsRolesLoading(false);
      return;
    }
    // Bei der allerersten Installation gibt es die Use-Case-Liste und das
    // Protokoll womöglich noch nicht (der `UseCaseProvider` legt sie gerade an —
    // beide teilen sich den Lauf). Ohne sie wäre jede Vergabe darauf ein
    // aussichtsloser Versuch mit 5,5 s Wartezeit.
    try {
      await service.ensureUseCaseList();
      await service.ensureLogList();
    } catch (e) {
      console.warn('[AIUC] Erstinstallation: Listen konnten nicht sichergestellt werden:', e);
    }
    const fehlt = await service.grantRoleRights(myEmail, 'Admin');
    await service.log(0, 'rolle-vergeben', `${myEmail} · User → Admin (Erstinstallation)`);
    setErstinstallationFehlt(fehlt);
    setErstinstallation(fehlt.length > 0 ? 'rechte-fehlen' : 'gespeichert');
    // Die echte Zeile lesen (mit ihrer Id) statt eine lokale mit Id 0 zu erfinden:
    // Änderungen und Entfernen laufen über die Id.
    const rows = await service.getRoles();
    if (rows === null) {
      setRoles([]);
      setCurrentUserRole('Admin');
      setRolesReadStatus('error');
    } else {
      const mapped = mapRows(rows);
      setRoles(mapped);
      const mine = mapped.filter(r => isCurrentUser(props.context, r.userEmail))[0];
      setCurrentUserRole(mine ? mine.role : 'User');
    }
    setIsRolesLoading(false);
  }

  const erstinstallationWiederholen = async (): Promise<void> => {
    setIsRolesLoading(true);
    setErstinstallation('keine');
    try {
      await init();
    } catch {
      setIsRolesLoading(false);
    }
  };

  const refreshRoles = React.useCallback(async (): Promise<void> => {
    const rows = await service.getRoles();
    if (rows === null) {
      // Bestehenden Stand STEHEN LASSEN. Ein Netzwerkfehler mitten in der
      // Nutzung darf niemanden herabstufen.
      console.warn('[AIUC] Rollen nicht lesbar — bestehender Stand bleibt.');
      // Antwortet die Liste jetzt eindeutig mit 403, war der frühere „error" nur ein
      // Fehlversuch — die Person darf die Liste nicht lesen und ist User; der Hinweis
      // „Rolle konnte nicht gelesen werden" gilt dann nicht mehr.
      if (service.lastRolesReadStatus === 403) setRolesReadStatus(prev => (prev === 'error' ? 'forbidden' : prev));
      return;
    }
    setRolesReadStatus('ok');
    const mapped = mapRows(rows);
    setRoles(mapped);
    const mine = mapped.filter(r => isCurrentUser(props.context, r.userEmail))[0];
    setCurrentUserRole(mine ? mine.role : 'User');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [service]);

  // ---- Selbstschutz über das aufgelöste Konto --------------------------------

  // Zeilen mit demselben Namen wie die angemeldete Person: Steht sie zusätzlich
  // unter einem Alias in der Liste, löst dessen Adresse zum SELBEN Konto auf —
  // `isCurrentUser` erkennt das nicht, die Oberfläche würde Auswahl und
  // Papierkorb freigeben. Aufgelöst wird nur, wo der Name passt (ein Aufruf je
  // Zeile, einmal).
  React.useEffect(() => {
    if (currentUserRole !== 'Admin') return undefined;
    let abgebrochen = false;
    (async (): Promise<void> => {
      const kandidaten = roles.filter(r => {
        const lc = norm(r.userEmail);
        return !!lc && !isCurrentUser(props.context, r.userEmail) && !selbstGeprueft.current[lc] && gleicherNameWie(r.userName, myName);
      });
      for (let i = 0; i < kandidaten.length; i++) {
        const lc = norm(kandidaten[i].userEmail);
        // eslint-disable-next-line no-await-in-loop
        const ich = await service.istMeinKonto(kandidaten[i].userEmail);
        // `null` = nicht auflösbar: nicht als „geprüft" merken, beim nächsten Mal noch einmal.
        if (ich !== null) selbstGeprueft.current[lc] = true;
        if (ich === true && selbstRef.current.indexOf(lc) < 0) {
          selbstRef.current = selbstRef.current.concat([lc]);
          if (!abgebrochen) setSelbstKonten(selbstRef.current);
        }
        if (abgebrochen) return;
      }
    })().catch(() => undefined);
    return () => { abgebrochen = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [roles, currentUserRole]);

  const isSelf = React.useCallback(
    (email: string): boolean => isCurrentUser(props.context, email) || selbstKonten.indexOf(norm(email)) >= 0,
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [selbstKonten],
  );

  /** Vor einer Änderung: Ist diese Adresse mein Konto? (Bekannte Schreibweisen sofort, sonst über die Auflösung.) */
  async function istSelbst(email: string): Promise<boolean> {
    if (isCurrentUser(props.context, email) || selbstRef.current.indexOf(norm(email)) >= 0) return true;
    return (await service.istMeinKonto(email)) === true;
  }

  // ---- Rollen schreiben ------------------------------------------------------

  const protokolliere = async (aktion: 'rolle-vergeben' | 'rolle-geaendert' | 'rolle-entfernt', email: string, von: string, nach: string): Promise<void> => {
    // Best-effort wie das ganze Protokoll: Ein fehlendes Protokoll verhindert keine Aktion.
    await service.log(0, aktion, `${email} · ${von} → ${nach}`);
  };

  /** Rechte auf das Maß einer niedrigeren Rolle zurückführen — Ergebnis als Grund (`''` = bestätigt). */
  async function rechteHerabsetzen(email: string, nach: UserRole): Promise<AktionsGrund> {
    const r = await service.reduziereRechte(email, nach === 'Organizer' ? 'Organizer' : 'User');
    if (r === 'ok') return '';
    return r === 'selbst' ? 'selbst' : 'entzug';
  }

  /**
   * Die GÜLTIGE Rolle einer Adresse ändern (`zeilenAktion` = ändern oder
   * löschen der Zeile) — in der Reihenfolge, die Rechte ohne Zeile ausschließt:
   *  - Herabstufen / Entfernen: ERST die Rechte zurückführen (nachgelesen
   *    bestätigt), dann die Zeile. Scheitert der Entzug, bleibt die Zeile.
   *  - Erhöhen: erst die Zeile, dann die Rechte (additiv, nachgelesen).
   */
  async function aendereGueltigeRolle(email: string, von: UserRole, nach: UserRole, zeilenAktion: () => Promise<boolean>, entfernen = false): Promise<boolean> {
    // `entfernen`: die LETZTE Zeile dieser Adresse geht weg. Dann werden die
    // Rechte auch bei der Rolle `User` entzogen — eine User-Zeile kann noch
    // Reste einer früheren, gescheiterten Herabstufung tragen (genau das meldet
    // der Audit als „zu viele Rechte"), und mit der Zeile verschwände die Person
    // aus dem Audit, der nur Zeilen abklappert (Review 29.09.2026, Fund 2).
    if (entfernen || roleRank(nach) < roleRank(von)) {
      const grund = await rechteHerabsetzen(email, nach);
      if (grund) { grundRef.current = grund; return false; }
      if (!(await zeilenAktion())) { grundRef.current = 'zeile-nach-entzug'; return false; }
      // Was die neue Rolle sonst noch braucht (bei Organizer: alles, was der Admin
      // nicht schon als Organizer-Stufe mitbrachte), additiv und nachgelesen.
      if (nach !== 'User') rightsMissingRef.current = await service.grantRoleRights(email, nach);
      return true;
    }
    if (!(await zeilenAktion())) { grundRef.current = 'zeile'; return false; }
    if (nach !== 'User') rightsMissingRef.current = await service.grantRoleRights(email, nach);
    return true;
  }

  /** Eine bestehende Zeile ändern — mit frisch gelesener Liste `rows`. */
  async function aendereZeile(rows: RoleAssignment[], before: RoleAssignment, role: UserRole): Promise<boolean> {
    const lc = norm(before.userEmail);
    if (await istSelbst(before.userEmail)) { grundRef.current = 'selbst'; return false; }
    const gueltig = lc ? rows.filter(r => norm(r.userEmail) === lc)[0] : before;
    if (!lc || (gueltig && gueltig.id !== before.id)) {
      // Überzählige Zeile derselben Adresse (oder Zeile ohne Adresse): Es gilt die
      // erste — die Rechte der Person hängen nicht an dieser Zeile und bleiben
      // unberührt. Nur die Zeile ändern.
      const ok = await service.updateRole(before.id, role, myName);
      if (ok) {
        grundRef.current = 'nur-zeile';
        await protokolliere('rolle-geaendert', before.userEmail, before.role, role);
      } else {
        grundRef.current = 'zeile';
      }
      return ok;
    }
    const ok = await aendereGueltigeRolle(before.userEmail, before.role, role, () => service.updateRole(before.id, role, myName));
    if (ok) await protokolliere('rolle-geaendert', before.userEmail, before.role, role);
    return ok;
  }

  const addRole = React.useCallback(async (email: string, name: string, role: UserRole): Promise<boolean> => {
    rightsMissingRef.current = [];
    grundRef.current = '';
    if (!echterAdminRef.current) { grundRef.current = 'keine-berechtigung'; return false; }
    // Frisch lesen statt dem zwischengespeicherten State zu glauben: Zwei Admins
    // gleichzeitig hätten sonst zwei Zeilen für dieselbe Adresse angelegt.
    const rows = await frischeRollen();
    if (rows === null) { grundRef.current = 'lesefehler'; return false; }
    const lc = norm(email);
    const existing = rows.filter(r => norm(r.userEmail) === lc)[0];

    if (existing) {
      // Nie implizit herabstufen: `addRole` heisst „diese Rechte
      // sicherstellen". Wer schon mehr hat, verloere sie sonst still
      // (DEX v30.67).
      if (roleRank(existing.role) > roleRank(role)) {
        console.warn(`[AIUC] ${lc} hat bereits ${existing.role} — ${role} waere eine Herabstufung und wird NICHT gesetzt.`);
        grundRef.current = 'herabstufung';
        return false;
      }
      if (existing.role === role) {
        // Rechte trotzdem nachsetzen: Sie koennen beim ersten Mal an der
        // Drosselung gescheitert sein.
        rightsMissingRef.current = await service.grantRoleRights(email, role);
        return true;
      }
      const ok = await aendereZeile(rows, existing, role);
      await refreshRoles();
      return ok;
    }

    const ok = await service.addRole(email, name, role, myName);
    if (!ok) { grundRef.current = 'zeile'; return false; }
    rightsMissingRef.current = await service.grantRoleRights(email, role);
    await protokolliere('rolle-vergeben', email, 'User', role);
    await refreshRoles();
    return true;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [service, refreshRoles, myName]);

  const updateRole = React.useCallback(async (itemId: number, role: UserRole): Promise<boolean> => {
    rightsMissingRef.current = [];
    grundRef.current = '';
    if (!echterAdminRef.current) { grundRef.current = 'keine-berechtigung'; return false; }
    const rows = await frischeRollen();
    if (rows === null) { grundRef.current = 'lesefehler'; return false; }
    const before = rows.filter(r => r.id === itemId)[0];
    if (!before) { grundRef.current = 'nicht-gefunden'; return false; }
    const ok = await aendereZeile(rows, before, role);
    await refreshRoles();
    return ok;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [service, refreshRoles, myName]);

  const removeRole = React.useCallback(async (itemId: number): Promise<boolean> => {
    rightsMissingRef.current = [];
    grundRef.current = '';
    if (!echterAdminRef.current) { grundRef.current = 'keine-berechtigung'; return false; }
    const rows = await frischeRollen();
    if (rows === null) { grundRef.current = 'lesefehler'; return false; }
    const entry = rows.filter(r => r.id === itemId)[0];
    if (!entry) {
      // Die Zeile ist schon weg (zweiter Admin, zweiter Tab): Es bleibt nichts zu tun.
      await refreshRoles();
      return true;
    }
    if (await istSelbst(entry.userEmail)) { grundRef.current = 'selbst'; return false; }
    const lc = norm(entry.userEmail);
    const gleiche = lc ? rows.filter(r => norm(r.userEmail) === lc) : [entry];

    let ok: boolean;
    if (!lc || gleiche[0].id !== entry.id) {
      // Überzählige Zeile: Es gilt die erste. Die Rechte der Person hängen an
      // DIESER Adresse, nicht an dieser Zeile — sie zu entziehen nähme der Person
      // die Rechte, die die gültige Zeile ihr gibt.
      ok = await service.deleteRole(entry.id);
      grundRef.current = ok ? 'nur-zeile' : 'zeile';
    } else {
      // Die gültige Zeile. Gibt es weitere mit derselben Adresse, wird die nächste
      // gültig — die Rechte richten sich dann nach ihrer Rolle, sie werden nicht
      // entzogen.
      const naechste = gleiche.filter(r => r.id !== entry.id)[0];
      ok = await aendereGueltigeRolle(entry.userEmail, entry.role, naechste ? naechste.role : 'User', () => service.deleteRole(entry.id), !naechste);
    }
    if (ok) await protokolliere('rolle-entfernt', entry.userEmail, entry.role, 'User');
    await refreshRoles();
    return ok;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [service, refreshRoles]);

  // ---- v1.3: Personensuche und Rechte-Prüfung (Durchreichen an die Rollenverwaltung) ----

  const searchUsers = React.useCallback(
    (query: string, includeInternational?: boolean): Promise<PersonenSuche | null> => service.searchUsers(query, !!includeInternational),
    [service],
  );

  const cancelSearch = React.useCallback((): void => service.sucheAbbrechen(), [service]);

  const auditRoleRights = React.useCallback(async (onProgress?: (done: number, total: number) => void): Promise<RechteBericht | null> => {
    // Die Rollenliste FRISCH lesen, nicht aus dem State nehmen: Der Audit soll
    // sagen, was JETZT gilt, und der State kann veraltet sein (ein früheres
    // Nachladen ist gescheitert und hat den alten Stand stehen lassen).
    const rows = await service.getRoles();
    if (rows === null) {
      service.lastRightsReadFehler = {
        code: service.lastRolesReadCode === 'zu-gross' ? 'zu-gross' : (service.lastRolesReadStatus ? 'http' : 'keine-antwort'),
        liste: 'rollen',
        status: service.lastRolesReadStatus,
      };
      return null;
    }
    const bericht = await service.auditRoleRights(mapRows(rows).map(r => ({ email: r.userEmail, name: r.userName, rolle: r.role })), onProgress);
    // Zeilen, die der Audit beim Auflösen als EIGENES Konto erkannt hat (Alias),
    // sperrt die Oberfläche ab jetzt ebenfalls.
    if (bericht && bericht.eigeneKonten.length > 0) {
      const neu = bericht.eigeneKonten.filter(a => selbstRef.current.indexOf(a) < 0);
      if (neu.length > 0) {
        selbstRef.current = selbstRef.current.concat(neu);
        setSelbstKonten(selbstRef.current);
      }
    }
    return bericht;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [service]);

  const repairRoleRights = React.useCallback(async (
    luecken: RechteLuecke[],
    onProgress?: (done: number, total: number) => void,
  ): Promise<RepairErgebnis> => {
    // Der Bericht kann Minuten alt sein (zweiter Admin, zweiter Tab). Nachgesetzt
    // wird nur für Zeilen, deren Rolle noch dieselbe ist — sonst bekäme eine
    // herabgestufte oder entfernte Person ihre Rechte zurück, und die
    // Herabstufung wäre wirkungslos, ohne dass es jemand merkt. Nicht lesbar
    // heißt: nichts anfassen.
    const rows = await service.getRoles();
    if (rows === null) return { behoben: [], offen: [], rolleGeaendert: [], lesefehler: true };
    const aktuell = mapRows(rows);
    const gueltig: RechteLuecke[] = [];
    const geaendert: RechteLuecke[] = [];
    luecken.forEach(l => {
      const lc = norm(l.email);
      const jetzt = aktuell.filter(r => norm(r.userEmail) === lc)[0];
      if (jetzt && jetzt.role === l.rolle) gueltig.push(l); else geaendert.push(l);
    });
    const res = gueltig.length > 0 ? await service.repairRoleRights(gueltig, onProgress) : { behoben: [], offen: [] };
    return { ...res, rolleGeaendert: geaendert, lesefehler: false };
  }, [service]);

  const revokeExcessRights = React.useCallback(async (
    items: RechteUeberschuss[],
    onProgress?: (done: number, total: number) => void,
  ): Promise<EntzugsErgebnis> => {
    // Die Rollen frisch lesen: Der Bericht kann alt sein. Wurde aus dem
    // Organizer inzwischen ein Admin, würde „Vollzugriff entziehen, Lesen
    // zurückgeben" ihn auf der Rollenliste HERABSTUFEN. Nicht lesbar heißt hier
    // nichts anfassen — unbekannt sperrt.
    const rows = await service.getRoles();
    if (rows === null) return { erledigt: 0, offen: items.map(it => ({ ...it, grund: 'lesefehler' as UeberschussGrund })), lesenFehlt: [] };
    const aktuell = mapRows(rows);
    const offen: EntzugsErgebnis['offen'] = [];
    const lesenFehlt: RechteUeberschuss[] = [];
    let erledigt = 0;
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const lc = norm(it.email);
      const jetzt = aktuell.filter(r => norm(r.userEmail) === lc)[0];
      if (!jetzt || jetzt.role !== it.rolle) {
        offen.push({ ...it, grund: 'rolle-geaendert' });
      } else {
        const res = await service.entzieheUeberschuss(it.email, it.rolle, it.listen);
        if (res.selbst) {
          offen.push({ ...it, grund: 'selbst' });
        } else {
          if (res.offen.length > 0) offen.push({ ...it, listen: res.offen, grund: res.lesefehler ? 'lesefehler' : 'nicht-bestaetigt' });
          if (res.lesenFehlt.length > 0) lesenFehlt.push({ ...it, listen: res.lesenFehlt });
          if (res.offen.length === 0 && res.lesenFehlt.length === 0) erledigt++;
        }
      }
      if (onProgress) onProgress(i + 1, items.length);
    }
    return { erledigt, offen, lesenFehlt };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [service]);

  const effectiveRole: UserRole = previewAsUser ? 'User' : currentUserRole;
  const isAdmin = effectiveRole === 'Admin';
  const isOrganizer = effectiveRole === 'Organizer' || isAdmin;

  echterAdminRef.current = currentUserRole === 'Admin';

  const value = React.useMemo<RoleContextType>(() => ({
    roles, currentUserRole, isRolesLoading, rolesReadStatus,
    isAdmin, isOrganizer, previewAsUser, setPreviewAsUser,
    originalIsAdmin: currentUserRole === 'Admin',
    originalIsOrganizer: currentUserRole === 'Organizer' || currentUserRole === 'Admin',
    addRole, updateRole, removeRole, refreshRoles,
    lastRightsMissing: () => rightsMissingRef.current,
    lastAktionGrund: () => grundRef.current,
    erstinstallation, erstinstallationFehlt, erstinstallationWiederholen,
    searchUsers, cancelSearch, lastPersonSearchStatus: () => service.lastPersonSearchStatus,
    auditRoleRights, lastRightsReadFehler: () => service.lastRightsReadFehler,
    repairRoleRights, revokeExcessRights, isSelf,
    siteUrl: props.context.pageContext.web.absoluteUrl,
    service,
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [roles, currentUserRole, isRolesLoading, rolesReadStatus, previewAsUser, erstinstallation, erstinstallationFehlt, addRole, updateRole, removeRole, refreshRoles, service, searchUsers, cancelSearch, auditRoleRights, repairRoleRights, revokeExcessRights, isSelf]);

  return React.createElement(RoleContext.Provider, { value }, props.children);
}

export function useRoles(): RoleContextType {
  const ctx = React.useContext(RoleContext);
  if (!ctx) throw new Error('useRoles muss innerhalb des RoleProvider stehen');
  return ctx;
}
