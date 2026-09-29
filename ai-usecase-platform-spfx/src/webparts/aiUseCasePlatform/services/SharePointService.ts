/**
 * SharePoint-Zugriff der AI Use Case Platform.
 *
 * Aufbau wie `services/SharePointService.ts` in DEX — dieselben REST-Muster,
 * dieselbe Fehlerhaltung. Vier Lehren aus DEX sind hier von Anfang an
 * eingebaut statt nachtraeglich gelernt:
 *
 * 1. **Ein Lesefehler ist keine Null.** `getRoles` und `getUseCases` liefern
 *    bei einem Fehler `null`, nie `[]`. In DEX hat die Verwechslung 1045
 *    „verwaiste" Verweise erfunden (v29.0) und einer Co-Organizerin ein Event
 *    mit 77 Anmeldungen als komplett leer gezeigt (v30.37). Wer aus einer
 *    leeren Liste auf „nichts da" schliesst, braucht den geprueften Status.
 *
 * 2. **Eine Rechtevergabe ist erst gesetzt, wenn sie nachgelesen wurde.**
 *    `_grantVerified` wiederholt und liest danach `roledefinitionbindings` —
 *    ein `addroleassignment`-POST allein ist keine Vergabe. In DEX fehlten
 *    18 von 126 Rollen-Eintraegen mindestens ein Recht, und jede dieser
 *    Zuweisungen hatte „erfolgreich" gemeldet (v30.85).
 *
 * 3. **Wer Rechte vergibt, baut den Entzug im selben Commit.**
 *    `revokeAccessOnRolesList` liest danach nach; SharePoint antwortet auf
 *    einen DELETE ohne vorhandene Zuweisung mit 404 ODER 500, der Status
 *    allein traegt also nicht (v30.67).
 *
 * 4. **Erst pruefen, dann anlegen.** Spalten werden nur ergaenzt, wenn sie
 *    fehlen. Der Anlage-Versuch bei jedem Start beantwortet SharePoint mit
 *    HTTP 500, der Browser schreibt eine rote Zeile in die Konsole, bevor
 *    unser `catch` sie sieht — und macht echte Meldungen daneben
 *    unglaubwuerdig (v30.65).
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { withThrottleRetry, isThrottled } from '../utils/spThrottle';
import { isCurrentUser } from '../utils/sessionIdentities';
import { rolleFuerSpeicher } from '../utils/rollen';
import { LIST } from '../constants';
import { UserRole, UseCase, UseCaseStatus, Bewertung, AufrufArt, LogEintrag } from '../types';

/** SharePoint-Rechtestufen. Sprachunabhaengige Standard-IDs. */
const ROLE_DEF = {
  read: 1073741826,
  contribute: 1073741827,
  edit: 1073741830,
  design: 1073741828,
  full: 1073741829,
};

/** Feldtypen der SharePoint-REST-API. */
const FIELD = {
  text: 2,
  note: 3,
  number: 9,
  boolean: 8,
  dateTime: 4,
  choice: 6,
  url: 11,
};

/**
 * Der Merker im Protokoll, der sagt: Die Start-Use-Cases wurden schon einmal
 * angelegt. Ohne ihn kaeme der Bestand zurueck, sobald jemand alle Eintraege
 * absichtlich geloescht hat.
 */
const SEED_MARKER = 'erstbefuellung';

/**
 * v1.3: Präfix der Kachelbild-Anhänge an einer Use-Case-Zeile.
 *
 * Nur Anhänge mit diesem Präfix gehören der App. Was jemand in SharePoint von
 * Hand an die Zeile hängt, wird beim Bildwechsel nie angefasst — dieselbe
 * Trennung wie `__eventimage__` in DEX.
 */
const BILD_PREFIX = '__ucbild__';

/** Was beim Lesen herauskam — `ok` heisst: die Daten sind belastbar. */
export type LeseStatus = 'ok' | 'forbidden' | 'notfound' | 'error';

interface SpUseCaseRow {
  Id: number;
  Title: string;
  Kurzbeschreibung?: string;
  Beschreibung?: string;
  Bereich?: string;
  UcStatus?: string;
  SalesRelevanz?: string;
  Machbarkeit?: string;
  DemoTauglichkeit?: string;
  AufrufArt?: string;
  LinkSourceCode?: string;
  LinkDeployment?: string;
  LinkGuide?: string;
  LinkWiki?: string;
  LinkVideo?: string;
  BildUrl?: string;
  Reihenfolge?: number;
  BetreuerEmails?: string;
  BetreuerNamen?: string;
  Schlagworte?: string;
  Modified?: string;
  Editor?: { Title?: string };
}

/** Eine Zeile der Protokollliste, so wie SharePoint sie liefert. */
interface SpLogRow {
  Id: number;
  Title?: string;
  UseCaseId?: number;
  Aktion?: string;
  Detail?: string;
  Wer?: string;
  Created?: string;
}

/**
 * v1.3: Die Listen, auf denen die App Rechte vergibt — und was jede Rolle dort
 * mindestens braucht. Das ist die EINE Tabelle, gegen die „Rechte prüfen"
 * (`auditRoleRights`) liest und nach der `repairRoleRights` nachsetzt.
 *
 * Sie muss zu den Vergaben passen: `grantFullControlOnRolesList` /
 * `grantReadOnRolesList` (Rollenliste) und `grantOrganizerPermissions`
 * (Use-Case-Liste, Protokoll). Steht hier etwas anderes als dort, meldet der
 * Audit „fehlt" für ein Recht, das die Vergabe nie setzt — und die Meldung
 * kommt bei jedem Lauf wieder. Die Matrix in `data/rollenMatrix.ts` (Kategorie
 * „SharePoint") beschreibt dieselben Zeilen in Sätzen.
 */
export type RechteListe = 'roles' | 'useCases' | 'log';

interface RechteSoll { key: RechteListe; liste: string; admin: number; organizer: number }

const RECHTE_SOLL: RechteSoll[] = [
  { key: 'roles', liste: LIST.roles, admin: ROLE_DEF.full, organizer: ROLE_DEF.read },
  { key: 'useCases', liste: LIST.useCases, admin: ROLE_DEF.edit, organizer: ROLE_DEF.edit },
  { key: 'log', liste: LIST.log, admin: ROLE_DEF.contribute, organizer: ROLE_DEF.contribute },
];

/** Rechtestufen aufsteigend: jede enthält die davor (Read < Contribute < Edit < Design < Full Control). */
const STUFEN = [ROLE_DEF.read, ROLE_DEF.contribute, ROLE_DEF.edit, ROLE_DEF.design, ROLE_DEF.full];

/** „Limited Access" hängt SharePoint von selbst an, sobald jemand eine einzelne Zeile sehen darf — es ist kein Recht auf der Liste. */
const LIMITED_ACCESS = 1073741825;

/**
 * Deckt eine der Stufen `ids` das geforderte Recht `need` ab? Höhere Stufen
 * zählen mit — wer Full Control hat, braucht kein zusätzliches Edit.
 *
 * (`grantVerified` zählt beim Nachlesen enger: nur `need` selbst und Full
 * Control. Das ist Absicht — dort geht es darum, dass die EIGENE Vergabe
 * ankam; hier darum, ob die Person das Recht hat, gleich wie.)
 */
function deckt(ids: number[] | undefined, need: number): boolean {
  if (!ids) return false;
  const ab = STUFEN.indexOf(need);
  return ids.some(id => { const i = STUFEN.indexOf(id); return i >= 0 && i >= ab; });
}

/** Zwei Stufen-Listen zusammenlegen (Adresse und Konto derselben Person). */
function vereint(a: number[] | undefined, b: number[] | undefined): number[] {
  const out: number[] = [];
  (a || []).concat(b || []).forEach(x => { if (out.indexOf(x) < 0) out.push(x); });
  return out;
}

/** Ein Treffer der Personensuche — mit der Adresse, die SharePoint führt. */
export interface PersonenTreffer {
  email: string;
  displayName: string;
  /** Position; der People-Picker liefert sie unter `EntityData.Title`. */
  jobTitle: string;
  department: string;
}

export interface PersonenSuche {
  treffer: PersonenTreffer[];
  /**
   * Treffer, die der Adressfilter aussortiert hat. Ohne diese Zahl hieße
   * „keine Treffer" auch dort, wo es welche gab — nur mit anderer Domain.
   */
  ausgeblendet: number;
}

/** Eine Person, bei der ein Recht fehlt, das ihre Rolle braucht. */
export interface RechteLuecke {
  email: string;
  name: string;
  rolle: UserRole;
  fehlt: RechteListe[];
  /** false = `ensureuser` kannte das Konto nicht (ausgeschieden?) oder wurde gedrosselt — nachsetzen geht dann nicht. */
  kontoAufloesbar: boolean;
}

/** Eine Person, die MEHR hat, als ihre Rolle vorsieht (z. B. ein früherer Admin). */
export interface RechteUeberschuss {
  email: string;
  name: string;
  rolle: UserRole;
  listen: RechteListe[];
}

export interface RechteBericht {
  /** Wie viele Rollen-Zeilen angesehen wurden (ohne doppelte Adressen und ohne leere). */
  geprueft: number;
  /** Davon ohne Befund. */
  ok: number;
  /**
   * Die Adressen (klein, getrimmt) der Zeilen ohne Befund. Die Zahl `ok` allein
   * genügt der Oberfläche nicht: Sie will je Zeile sagen können „geprüft, in
   * Ordnung" — und bei einer Zeile, die erst nach dem Lauf entstand, ehrlich
   * „nicht geprüft".
   */
  okAdressen: string[];
  luecken: RechteLuecke[];
  ueberschuss: RechteUeberschuss[];
  /** Rechte da — aber unter einer anderen Schreibweise der Adresse als in der Rollenliste. */
  aliase: Array<{ name: string; email: string; spEmail: string }>;
  /** Zeilen ohne Adresse: nicht prüfbar, nicht „in Ordnung". */
  ohneAdresse: number;
}

export interface RechteReparatur {
  behoben: Array<{ email: string; name: string }>;
  /** Was NACH dem Versuch noch fehlt (`fehlt` nennt nur die verbliebenen Listen). */
  offen: RechteLuecke[];
}

/** Die Zuweisungen EINER Liste, so gelesen, dass Adresse und Konto beide taugen. */
interface ListenZuweisungen {
  /** Hat die Liste eigene Berechtigungen (true) oder erbt sie von der Site (false)? */
  eigene: boolean;
  nachAdresse: Map<string, number[]>;
  nachId: Map<number, number[]>;
  adresseZuId: Map<number, string>;
}

/** Ein Eintrag der Antwort des People-Pickers — nur die Felder, die wir lesen. */
interface PickerEintrag {
  Key?: string;
  DisplayText?: string;
  Description?: string;
  EntityData?: { Email?: string; Title?: string; Department?: string };
}

/** Eine Zuweisung, wie `roleassignments?$expand=…` sie liefert (nometadata oder verbose). */
interface SpZuweisung {
  PrincipalId?: number;
  Member?: { Email?: string; LoginName?: string; PrincipalType?: number };
  RoleDefinitionBindings?: Array<{ Id: number }> | { results?: Array<{ Id: number }> };
}

export class SharePointService {
  private context: WebPartContext;
  private siteUrl: string;

  /** HTTP-Status des letzten Rollen-Lesens (0 = Ausnahme oder nie gelesen). */
  public lastRolesReadStatus = 0;
  /** HTTP-Status des letzten Use-Case-Lesens. */
  public lastUseCasesReadStatus = 0;
  /** Klartext der letzten fehlgeschlagenen Antwort — fuer die Meldung in der App. */
  public lastReadError = '';
  /**
   * v1.3: HTTP-Status und Klartext des letzten Protokoll-Lesens — getrennt von
   * `lastReadError`. Die Kachelwand zeigt `lastReadError` an; ein Fehler beim
   * Protokoll darf dort keinen Fehler der Use-Case-Liste überschreiben.
   */
  public lastLogReadStatus = 0;
  public lastLogReadError = '';
  /** Spalten, die beim Sicherstellen der Liste NICHT entstanden sind. */
  public fehlendeSpalten: string[] = [];

  /**
   * Ersatz fuer `context.spHttpClient` — gleiche Signatur.
   *
   * **Nur `post` wiederholt.** Die Drosselung trifft das Schreiben. Lesen
   * ebenfalls durch die Schranke zu schicken war in DEX ein Fehler: Der Start
   * besteht fast nur aus GETs, und ein einziges 429 legte den ganzen
   * Bootvorgang stumm still (v29.50).
   */
  private _sp = {
    get: (url: string, cfg: SPHttpClientConfiguration, options?: ISPHttpClientOptions): Promise<SPHttpClientResponse> =>
      this.context.spHttpClient.get(url, cfg, options),
    post: (url: string, cfg: SPHttpClientConfiguration, options?: ISPHttpClientOptions): Promise<SPHttpClientResponse> =>
      withThrottleRetry(() => this.context.spHttpClient.post(url, cfg, options), url),
  };

  public constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  // ===================================================================
  // Grundlagen
  // ===================================================================

  /**
   * Der Entitaetstyp einer Liste fuer `__metadata` — gelesen, nicht geraten.
   *
   * v1.1 (Fehlerbehebung): Hier stand `SP.Data.${name}ListItem`, also
   * `SP.Data.AIUC_UseCasesListItem`. Den Typ gibt es nicht: SharePoint
   * kodiert den Unterstrich im Listennamen als `_x005f_`. Da `_post` mit
   * `odata=verbose` sendet, prueft SharePoint den Typ — JEDES Anlegen einer
   * Zeile endete mit HTTP 400. Sichtbar war davon nichts: `createUseCase`
   * gibt bei `!r.ok` einfach `null` zurueck, die Erstbefuellung lief also
   * durch und legte fuenfmal nichts an. Dieselbe Zeile stand auch in
   * `addRole`, `updateRole` und `log` — Rollen liessen sich nicht vergeben,
   * das Protokoll blieb leer.
   *
   * Die Kodierung nachzubauen waere die zweite Stelle, an der geraten wird.
   * SharePoint nennt den Typ selbst: `ListItemEntityTypeFullName`. Er wird
   * einmal je Liste gelesen und gemerkt; scheitert das Lesen, greift die
   * Kodier-Regel als Notnagel — besser als gar kein Typ.
   */
  private _entityTypes: Record<string, string> = {};

  public async entityType(listName: string): Promise<string> {
    const merker = this._entityTypes[listName];
    if (merker) return merker;
    const notnagel = `SP.Data.${listName.replace(/_/g, '_x005f_')}ListItem`;
    try {
      const r = await this._sp.get(
        `${this.list(listName)}?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      if (r.ok) {
        const d = await r.json();
        const typ = d.ListItemEntityTypeFullName || d.d?.ListItemEntityTypeFullName;
        if (typeof typ === 'string' && typ) {
          this._entityTypes[listName] = typ;
          return typ;
        }
      }
    } catch { /* Notnagel */ }
    this._entityTypes[listName] = notnagel;
    return notnagel;
  }

  /** Eine Zeile anlegen — mit dem Typ, den die Liste selbst nennt. */
  private async _postItem(listName: string, body: Record<string, unknown>): Promise<SPHttpClientResponse> {
    const typ = await this.entityType(listName);
    return this._post(`${this.list(listName)}/items`, { ...body, '__metadata': { 'type': typ } });
  }

  /** Eine Zeile aendern — mit dem Typ, den die Liste selbst nennt. */
  private async _mergeItem(listName: string, id: number, body: Record<string, unknown>): Promise<SPHttpClientResponse> {
    const typ = await this.entityType(listName);
    return this._merge(`${this.list(listName)}/items(${id})`, { ...body, '__metadata': { 'type': typ } });
  }

  private async _post(url: string, body: object): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
      },
      body: JSON.stringify(body),
    };
    return this._sp.post(url, SPHttpClient.configurations.v1, options);
  }

  /** MERGE auf ein bestehendes Element. */
  private async _merge(url: string, body: object): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(body),
    };
    return this._sp.post(url, SPHttpClient.configurations.v1, options);
  }

  private async _delete(url: string): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE',
      },
    };
    return this._sp.post(url, SPHttpClient.configurations.v1, options);
  }

  private list(name: string): string {
    return `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(name)}')`;
  }

  public async listExists(listName: string): Promise<boolean> {
    try {
      const r = await this._sp.get(this.list(listName), SPHttpClient.configurations.v1);
      return r.ok;
    } catch {
      return false;
    }
  }

  /** Die Antwort von SharePoint lesbar machen — der Text sagt, was fehlt. */
  private async fehlertext(r: SPHttpClientResponse): Promise<string> {
    try {
      const t = await r.clone().text();
      // SharePoint packt die Ursache tief ein; der Klartext steht in `value`.
      const m = /"value"\s*:\s*"([^"]{5,400})"/.exec(t);
      return m ? m[1] : t.slice(0, 300);
    } catch {
      return `HTTP ${r.status}`;
    }
  }

  /**
   * Eine Spalte anlegen — aber nur, wenn sie fehlt.
   *
   * Zwei Dinge, die beim ersten Live-Versuch am 10.09.2026 wehgetan haben:
   *
   * 1. **Der Feldtyp braucht seinen EIGENEN `__metadata`-Typ.** Ein generisches
   *    `SP.Field` mit `FieldTypeKind` legt eine Textspalte an oder scheitert,
   *    je nach Feldart — mehrzeilige Felder und Auswahlfelder brauchen
   *    `SP.FieldMultiLineText` bzw. `SP.FieldChoice`. Scheitert das, existiert
   *    die Spalte nicht, und der spaetere `$select` darauf antwortet mit
   *    HTTP 400 „field does not exist" — genau der Fehler auf dem Screenshot.
   *
   * 2. **Die Standardansicht heisst nicht ueberall „All Items".** In einem
   *    deutschsprachigen Tenant ist sie „Alle Elemente". Der Aufruf lief ins
   *    Leere; das ist harmlos (die Spalte existiert trotzdem), aber er darf
   *    nicht wie ein Fehler des Anlegens aussehen. Deshalb ueber
   *    `DefaultView` statt ueber den Namen.
   *
   * Rueckgabe: true = die Spalte ist danach da (angelegt oder war schon da).
   */
  private async ensureField(
    listName: string,
    internalName: string,
    payload: Record<string, unknown>,
  ): Promise<boolean> {
    try {
      const probe = await this._sp.get(
        `${this.list(listName)}/fields/getbytitle('${encodeURIComponent(internalName)}')`,
        SPHttpClient.configurations.v1,
      );
      if (probe.ok) return true;
    } catch { /* nicht lesbar -> Anlegen versuchen */ }

    try {
      const r = await this._post(`${this.list(listName)}/fields`, {
        'Title': internalName,
        'Required': false,
        ...payload,
      });
      if (!r.ok) {
        console.warn(`[AIUC] Spalte ${internalName} auf ${listName}: ${await this.fehlertext(r)}`);
        return false;
      }
      // In die Standardansicht aufnehmen, sonst ist die Spalte in SharePoint
      // selbst unsichtbar. Ueber DefaultView, nicht ueber den Namen.
      await this._post(
        `${this.list(listName)}/DefaultView/viewfields/addviewfield('${encodeURIComponent(internalName)}')`,
        {},
      ).catch(() => undefined);
      return true;
    } catch (e) {
      console.warn(`[AIUC] Spalte ${internalName} auf ${listName} konnte nicht angelegt werden:`, e);
      return false;
    }
  }

  /** Textspalte (einzeilig). */
  private feldText(liste: string, name: string): Promise<boolean> {
    return this.ensureField(liste, name, {
      '__metadata': { 'type': 'SP.FieldText' }, 'FieldTypeKind': FIELD.text, 'MaxLength': 255,
    });
  }

  /** Mehrzeiliges Textfeld. `richText` nur dort, wo wirklich HTML hineinsoll. */
  private feldNote(liste: string, name: string, richText = false): Promise<boolean> {
    return this.ensureField(liste, name, {
      '__metadata': { 'type': 'SP.FieldMultiLineText' }, 'FieldTypeKind': FIELD.note,
      'NumberOfLines': 6, 'RichText': richText, 'AllowHyperlink': false, 'AppendOnly': false,
    });
  }

  private feldZahl(liste: string, name: string): Promise<boolean> {
    return this.ensureField(liste, name, {
      '__metadata': { 'type': 'SP.FieldNumber' }, 'FieldTypeKind': FIELD.number,
    });
  }

  private feldAuswahl(liste: string, name: string, werte: string[]): Promise<boolean> {
    return this.ensureField(liste, name, {
      '__metadata': { 'type': 'SP.FieldChoice' }, 'FieldTypeKind': FIELD.choice,
      'Choices': { 'results': werte }, 'EditFormat': 0,
    });
  }

  private feldDatum(liste: string, name: string): Promise<boolean> {
    return this.ensureField(liste, name, {
      '__metadata': { 'type': 'SP.FieldDateTime' }, 'FieldTypeKind': FIELD.dateTime,
    });
  }

  private async createList(listName: string, description: string): Promise<void> {
    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': description,
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });
  }

  // ===================================================================
  // Rollen
  // ===================================================================

  /**
   * Rollenliste sicherstellen.
   *
   * `isNewlyCreated` ist sicherheitsrelevant und kein Komfort: Nur bei einer
   * echten Erstinstallation darf der aufrufende Mensch automatisch Admin
   * werden. In DEX lief das einmal ueber „getRoles ist leer" — und ein 403
   * beim Lesen machte damit JEDEN Aufrufer zum Admin (v6.34).
   */
  public async ensureRolesList(): Promise<{ isNewlyCreated: boolean }> {
    const name = LIST.roles;
    if (await this.listExists(name)) {
      await this.ensureRolesListPermissions(name);
      return { isNewlyCreated: false };
    }

    await this.createList(name, 'Rollenverwaltung der AI Use Case Platform');
    await this.feldText(name, 'UserName');
    // Die Auswahlwerte bleiben die des Altbestands (`Kurator` steht für den Use
    // Case Organizer) — `utils/rollen.ts` erklärt, warum.
    await this.feldAuswahl(name, 'Role', ['Admin', 'Kurator', 'User']);
    await this.feldText(name, 'AssignedBy');
    await this.feldDatum(name, 'AssignedDate');
    await this.ensureRolesListPermissions(name);
    return { isNewlyCreated: true };
  }

  /**
   * Die Rollenliste bekommt EIGENE Rechte: Owners Full Control, sonst
   * niemand. Sonst koennte sich jeder Mitwirkende selbst zum Admin machen.
   *
   * Folge davon — und der Grund, warum jede Zuweisung unten Read nachsetzt:
   * Eine Rolle wirkt nur, wenn die Person die Rollenliste LESEN darf. Fehlt
   * das Recht, antwortet `getRoles` mit 403, und die Person ist trotz Eintrag
   * ein normaler Nutzer (DEX v30.80).
   */
  private async ensureRolesListPermissions(listName: string): Promise<void> {
    try {
      const r = await this._sp.get(
        `${this.list(listName)}?$select=HasUniqueRoleAssignments`,
        SPHttpClient.configurations.v1,
      );
      if (!r.ok) return;
      const d = await r.json();
      const has = d.HasUniqueRoleAssignments ?? d.d?.HasUniqueRoleAssignments;
      if (has) return;

      await this._post(`${this.list(listName)}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`, {});

      const owners = await this._sp.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1,
      );
      if (owners.ok) {
        const od = await owners.json();
        const ownerId = od.Id ?? od.d?.Id;
        if (ownerId) {
          await this._post(
            `${this.list(listName)}/roleassignments/addroleassignment(principalid=${ownerId}, roledefid=${ROLE_DEF.full})`,
            {},
          );
        }
      }
      // Die anlegende Person braucht Full Control, sonst sperrt sie sich
      // selbst aus der gerade erzeugten Liste aus.
      const me = this.context.pageContext.legacyPageContext?.userId;
      if (me) {
        await this._post(
          `${this.list(listName)}/roleassignments/addroleassignment(principalid=${me}, roledefid=${ROLE_DEF.full})`,
          {},
        );
      }
    } catch (e) {
      console.warn('[AIUC] Rechte der Rollenliste konnten nicht gesetzt werden:', e);
    }
  }

  /**
   * Rollen lesen.
   *
   * `null` heisst NICHT LESBAR und ist etwas anderes als `[]` (keine Rollen
   * vergeben). Der Aufrufer muss beides unterscheiden — sonst wird aus einem
   * Netzwerkfehler eine Aussage ueber Berechtigungen.
   */
  public async getRoles(): Promise<Array<{ Id: number; Title: string; UserName: string; Role: string; AssignedBy: string; AssignedDate: string }> | null> {
    this.lastRolesReadStatus = 0;
    try {
      // Auch hier ohne `$select` — aus demselben Grund wie bei den Use Cases.
      // Fehlt eine der Spalten, antwortet SharePoint mit HTTP 400, `getRoles`
      // liefert `null`, und die Person ist „User", obwohl sie Admin sein
      // sollte. Das war der Zustand auf dem Screenshot vom 10.09.2026.
      const r = await this._sp.get(
        `${this.list(LIST.roles)}/items?$top=500`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      this.lastRolesReadStatus = r.status;
      if (!r.ok) {
        this.lastReadError = await this.fehlertext(r);
        console.warn(`[AIUC] Rollen lesen: HTTP ${r.status} — ${this.lastReadError}`);
        return null;
      }
      const d = await r.json();
      return (d.value || []) as Array<{ Id: number; Title: string; UserName: string; Role: string; AssignedBy: string; AssignedDate: string }>;
    } catch {
      return null;
    }
  }

  public async addRole(userEmail: string, userName: string, role: UserRole, assignedBy: string): Promise<boolean> {
    try {
      const r = await this._postItem(LIST.roles, {
        'Title': userEmail,
        'UserName': userName,
        'Role': rolleFuerSpeicher(role),
        'AssignedBy': assignedBy,
        'AssignedDate': new Date().toISOString(),
      });
      return r.ok;
    } catch (e) {
      console.error('[AIUC] addRole:', e);
      return false;
    }
  }

  public async updateRole(itemId: number, role: UserRole): Promise<boolean> {
    try {
      const r = await this._mergeItem(LIST.roles, itemId, { 'Role': rolleFuerSpeicher(role) });
      return r.ok;
    } catch (e) {
      console.error('[AIUC] updateRole:', e);
      return false;
    }
  }

  public async deleteRole(itemId: number): Promise<boolean> {
    try {
      const r = await this._delete(`${this.list(LIST.roles)}/items(${itemId})`);
      return r.ok;
    } catch (e) {
      console.error('[AIUC] deleteRole:', e);
      return false;
    }
  }

  // ===================================================================
  // Rechte vergeben und entziehen — spiegelbildlich
  // ===================================================================

  private async resolveUserId(userEmail: string): Promise<number | null> {
    // `ensureuser` legt das Konto an, falls die Person die Site noch nie
    // besucht hat — deshalb nicht `siteusers/getbyemail`, das dann 404 liefert
    // (DEX v31.84). v1.3: Zweiter Versuch im Claims-Format wie in DEX; manche
    // Konten löst SharePoint nur so auf.
    const namen = [userEmail, `i:0#.f|membership|${userEmail}`];
    for (let i = 0; i < namen.length; i++) {
      try {
        const r = await this._post(
          `${this.siteUrl}/_api/web/ensureuser`,
          { 'logonName': namen[i] },
        );
        if (!r.ok) continue;
        const d = await r.json();
        const id = d.d?.Id ?? d.Id;
        if (typeof id === 'number') return id;
      } catch {
        // nächster Versuch
      }
    }
    return null;
  }

  /**
   * Recht setzen UND nachlesen, mit Wiederholung.
   *
   * Ein `addroleassignment`-POST allein ist keine Vergabe: Ein 429 kostet das
   * Recht still, und es faellt erst auf, wenn jemand seine Kachel vermisst.
   * Full Control deckt jedes niedrigere Recht ab, deshalb zaehlt es beim
   * Nachlesen mit.
   */
  private async grantVerified(base: string, userId: number, roleDefId: number, label: string): Promise<boolean> {
    const sleep = (ms: number): Promise<void> => new Promise(res => setTimeout(res, ms));
    const verify = async (): Promise<boolean> => {
      try {
        const chk = await this._sp.get(
          `${base}/roleassignments/getbyprincipalid(${userId})/roledefinitionbindings?$select=Id`,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata' } },
        );
        if (!chk.ok) return false;
        const d = await chk.json();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ids: number[] = ((d.value || d.d?.results || []) as any[]).map(b => Number(b.Id));
        return ids.indexOf(roleDefId) >= 0 || ids.indexOf(ROLE_DEF.full) >= 0;
      } catch { return false; }
    };

    const delays = [1500, 4000];
    let lastStatus = 0;
    for (let i = 0; i <= delays.length; i++) {
      try {
        const resp = await this._post(`${base}/roleassignments/addroleassignment(principalid=${userId}, roledefid=${roleDefId})`, {});
        lastStatus = resp.status;
        // Die Zuweisung kann schon dagestanden haben (SharePoint antwortet
        // dann trotzdem 200) — deshalb immer nachlesen, nie dem POST glauben.
        if (await verify()) return true;
      } catch (e) {
        console.warn(`[AIUC] ${label}: Versuch ${i + 1} Ausnahme`, e);
      }
      if (i < delays.length) await sleep(delays[i]);
    }
    console.error(`[AIUC] ${label}: Recht ${roleDefId} fuer Principal ${userId} auf ${base} nach 3 Versuchen NICHT gesetzt (letzter HTTP-Status ${lastStatus}).`);
    return false;
  }

  /** Leserecht auf der Rollenliste — ohne das ist jede Rolle wirkungslos. */
  public async grantReadOnRolesList(userEmail: string): Promise<boolean> {
    const id = await this.resolveUserId(userEmail);
    if (!id) return false;
    return this.grantVerified(this.list(LIST.roles), id, ROLE_DEF.read, 'grantReadOnRolesList');
  }

  /** Vollzugriff auf der Rollenliste — nur fuer Admins. */
  public async grantFullControlOnRolesList(userEmail: string): Promise<boolean> {
    const id = await this.resolveUserId(userEmail);
    if (!id) return false;
    return this.grantVerified(this.list(LIST.roles), id, ROLE_DEF.full, 'grantFullControlOnRolesList');
  }

  /** Use Case Organizer duerfen Use Cases pflegen: Edit auf der Use-Case-Liste. */
  public async grantOrganizerPermissions(userEmail: string): Promise<string[]> {
    const fehlend: string[] = [];
    const id = await this.resolveUserId(userEmail);
    if (!id) return ['Person nicht aufloesbar'];
    // v1.3: Eine Liste, die noch von der Site erbt, nimmt keine eigene
    // Zuweisung an — der POST scheitert, oder das Nachlesen findet nichts. DEX
    // ruft dasselbe vor jeder Vergabe (`ensureListHasUniquePermissions`). Das
    // Ergebnis zählt hier nicht: Ob es geklappt hat, sagt `grantVerified`.
    await this.ensureListeEigeneRechte(LIST.useCases);
    await this.ensureListeEigeneRechte(LIST.log);
    if (!(await this.grantVerified(this.list(LIST.useCases), id, ROLE_DEF.edit, 'grantOrganizer/useCases'))) {
      fehlend.push('Use-Case-Liste (Bearbeiten)');
    }
    if (!(await this.grantVerified(this.list(LIST.log), id, ROLE_DEF.contribute, 'grantOrganizer/log'))) {
      fehlend.push('Protokoll (Beitragen)');
    }
    return fehlend;
  }

  /**
   * Eine direkte Zuweisung entfernen und NACHLESEN.
   *
   * Der DELETE-Status traegt nicht: SharePoint antwortet ohne vorhandene
   * Zuweisung mit 404 ODER 500. Erfolg heisst deshalb „der Principal steht
   * danach nicht mehr dran", nicht „der Aufruf war 200".
   */
  private async revokeVerified(base: string, userEmail: string, label: string): Promise<boolean> {
    const id = await this.resolveUserId(userEmail);
    if (!id) return false;
    try {
      await this._post(`${base}/roleassignments/removeroleassignment(principalid=${id}, roledefid=${ROLE_DEF.full})`, {}).catch(() => undefined);
      await this._delete(`${base}/roleassignments/getbyprincipalid(${id})`).catch(() => undefined);
      const chk = await this._sp.get(
        `${base}/roleassignments/getbyprincipalid(${id})/roledefinitionbindings?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      if (!chk.ok) {
        // v1.3: Nur 404 und 500 heißen „keine Zuweisung mehr da" — SharePoint
        // antwortet so, wenn der Principal nicht (mehr) in der Liste steht. Ein
        // 403, 429 oder 503 heißt: NICHT gelesen. Bis v1.2 stand hier
        // bedingungslos `true`, ein Lesefehler galt damit als gelungener Entzug
        // (dieselbe Verwechslung wie „leer" und „nicht lesbar").
        if (chk.status === 404 || chk.status === 500) return true;
        console.warn(`[AIUC] ${label}: Entzug für ${userEmail} nicht prüfbar (HTTP ${chk.status}) — gilt als NICHT entzogen.`);
        return false;
      }
      const d = await chk.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ids: number[] = ((d.value || d.d?.results || []) as any[]).map(b => Number(b.Id));
      if (ids.length === 0) return true;
      console.warn(`[AIUC] ${label}: ${userEmail} hat auf ${base} weiterhin Rechte (${ids.join(', ')}).`);
      return false;
    } catch {
      return false;
    }
  }

  /**
   * Alle direkten Rechte einer Person entziehen.
   *
   * Selbstschutz: Die angemeldete Person entzieht sich nicht selbst — ein
   * Admin, der die Rollenverwaltung an sich ausprobiert, kaeme sonst nicht
   * mehr an die Liste (DEX v30.67).
   */
  public async revokeAllAccess(userEmail: string): Promise<boolean> {
    if (isCurrentUser(this.context, userEmail)) {
      console.warn(`[AIUC] Rechte von ${userEmail} werden NICHT entzogen — das ist die angemeldete Person (Selbstschutz).`);
      return true;
    }
    const a = await this.revokeVerified(this.list(LIST.roles), userEmail, 'revoke/roles');
    const b = await this.revokeVerified(this.list(LIST.useCases), userEmail, 'revoke/useCases');
    const c = await this.revokeVerified(this.list(LIST.log), userEmail, 'revoke/log');
    return a && b && c;
  }

  /**
   * v1.3: Nur die direkte Zuweisung auf der ROLLENLISTE entziehen.
   *
   * Der Fall: Herabstufung Admin → Organizer. `addroleassignment` ist ADDITIV —
   * Read auf der Rollenliste käme NEBEN das bestehende Full Control, und der
   * frühere Admin könnte die Liste weiter bearbeiten und sich selbst wieder
   * hochstufen (DEX v30.67). Deshalb erst dieser Entzug, dann Read vergeben.
   * Selbstschutz wie bei `revokeAllAccess`.
   */
  public async revokeRolesListAccess(userEmail: string): Promise<boolean> {
    if (isCurrentUser(this.context, userEmail)) {
      console.warn(`[AIUC] Rollenliste: Rechte von ${userEmail} werden NICHT entzogen — das ist die angemeldete Person (Selbstschutz).`);
      return true;
    }
    return this.revokeVerified(this.list(LIST.roles), userEmail, 'revoke/roles');
  }

  // ===================================================================
  // Personensuche und Rechte-Prüfung (v1.3)
  // ===================================================================

  /** HTTP-Status der letzten Personensuche (0 = Ausnahme oder nie gesucht). */
  public lastPersonSearchStatus = 0;
  /** Warum das letzte Rechte-Lesen scheiterte — Klartext für die Meldung in der App. */
  public lastRightsReadError = '';

  /**
   * Personensuche über den People-Picker von SharePoint — derselbe Weg wie
   * `searchUsers` in DEX, ohne Graph. Graph braucht eine Admin-Freigabe im
   * Tenant; der Picker läuft mit dem, was jede Person auf der Site ohnehin hat.
   *
   * **`null` heißt: Die Suche ging NICHT.** Das ist etwas anderes als `treffer:
   * []`. DEX liefert bei jedem Fehler `[]`, und die Oberfläche sagt dann „Keine
   * Treffer" — für einen Admin, der eine Person sucht, ist das die Aussage
   * „die gibt es nicht", obwohl nur die Suche gedrosselt war. Hier: scheitern
   * ALLE Anfragen, oder scheitern einige und die übrigen finden nichts, ist
   * das Ergebnis `null`, und `lastPersonSearchStatus` nennt den Grund.
   *
   * Der Picker tokenisiert auf Leerzeichen; „Nachname, Vorname" und
   * „Vorname Nachname" finden ihn unterschiedlich gut. Deshalb mehrere
   * Schreibweisen der Eingabe PARALLEL, Treffer über die Adresse entdoppelt.
   */
  public async searchUsers(query: string, includeInternational = false): Promise<PersonenSuche | null> {
    this.lastPersonSearchStatus = 0;
    const roh = (query || '').trim();
    if (roh.length < 2) return { treffer: [], ausgeblendet: 0 };
    // Drosselt SharePoint gerade, würden bis zu sechs Anfragen je Tastendruck an
    // der Schranke warten und beim Öffnen GEMEINSAM losgehen — genau das Muster,
    // das aus einer Drosselung eine Sperre macht (`utils/spThrottle`). Lieber
    // sofort „nicht möglich" melden; die Person tippt gleich noch einmal.
    if (isThrottled()) {
      this.lastPersonSearchStatus = 429;
      return null;
    }

    const varianten: string[] = [];
    const merke = (v: string): void => {
      const s = v.trim();
      if (s && varianten.indexOf(s) < 0) varianten.push(s);
    };
    merke(roh);
    if (roh.indexOf(',') >= 0) {
      const teile = roh.split(',').map(s => s.trim()).filter(Boolean);
      if (teile.length >= 2) {
        const nach = teile[0];
        const vor = teile.slice(1).join(' ');
        merke(`${vor} ${nach}`);
        merke(`${nach} ${vor}`);
        if (nach.length >= 2) merke(nach);
      } else if (teile.length === 1 && teile[0].length >= 2) {
        merke(teile[0]);
      }
    } else if (roh.indexOf(' ') >= 0 && roh.indexOf('@') < 0) {
      // In deutschen Tenants heißt der Anzeigename „Nachname, Vorname"; wer
      // „Vorname Nachname" tippt, findet ihn über die Komma-Form besser.
      const teilworte = roh.split(/\s+/).filter(Boolean);
      if (teilworte.length === 2) {
        merke(`${teilworte[1]}, ${teilworte[0]}`);
        merke(`${teilworte[0]}, ${teilworte[1]}`);
        merke(`${teilworte[1]} ${teilworte[0]}`);
        if (teilworte[0].length >= 2) merke(teilworte[0]);
        if (teilworte[1].length >= 2) merke(teilworte[1]);
      } else if (teilworte.length > 2) {
        merke(teilworte.slice().reverse().join(' '));
        const letztes = teilworte[teilworte.length - 1];
        if (letztes.length >= 2) merke(letztes);
      }
    }

    const antworten = await Promise.all(varianten.map(async (variante): Promise<{ ok: boolean; status: number; eintraege: PickerEintrag[] }> => {
      try {
        const r = await this._post(
          `${this.siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
          {
            'queryParams': {
              '__metadata': { 'type': 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
              'AllowEmailAddresses': true,
              'AllowMultipleEntities': false,
              'MaximumEntitySuggestions': 30,
              'QueryString': variante,
              'PrincipalType': 1, // nur Personen
              'PrincipalSource': 15,
              'SharePointGroupID': 0,
            },
          },
        );
        if (!r.ok) return { ok: false, status: r.status, eintraege: [] };
        const d = await r.json();
        const text = d.d?.ClientPeoplePickerSearchUser || d.ClientPeoplePickerSearchUser || '[]';
        const liste = JSON.parse(text);
        return { ok: true, status: r.status, eintraege: Array.isArray(liste) ? liste as PickerEintrag[] : [] };
      } catch {
        return { ok: false, status: 0, eintraege: [] };
      }
    }));

    const gescheitert = antworten.filter(a => !a.ok);
    if (gescheitert.length === antworten.length) {
      this.lastPersonSearchStatus = gescheitert[0].status;
      console.warn(`[AIUC] Personensuche: alle ${antworten.length} Anfragen gescheitert (HTTP ${gescheitert[0].status}).`);
      return null;
    }
    this.lastPersonSearchStatus = 200;

    const istDeloitte = (mail: string): boolean => {
      const at = mail.lastIndexOf('@');
      // „deloitte" direkt am Anfang der Domain oder eines Labels: deckt
      // deloitte.com, deloitte.at, deloitteCE.com und Subdomains ab, aber keine
      // fremden Domains, die nur den Namen enthalten.
      return at >= 0 && /(^|\.)deloitte[a-z0-9-]*\./.test(mail.slice(at + 1));
    };

    const gesehen: string[] = [];
    const alle: PersonenTreffer[] = [];
    let ausgeblendet = 0;
    antworten.forEach(a => a.eintraege.forEach(e => {
      // Der Picker liefert manchmal Treffer mit leerem `EntityData.Email` (Person
      // ohne Profil-Seite); die Adresse steckt dann im Schlüssel oder in der
      // Beschreibung. Ohne diesen Umweg fehlen genau diese Personen.
      let email = String(e.EntityData?.Email || '').toLowerCase().trim();
      if (!email) {
        const kandidaten = [String(e.Key || ''), String(e.Description || '')];
        for (let i = 0; i < kandidaten.length && !email; i++) {
          const m = kandidaten[i].match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
          if (m) email = m[0].toLowerCase();
        }
      }
      if (!email || email.indexOf('@') < 0 || gesehen.indexOf(email) >= 0) return;
      gesehen.push(email);
      const passt = includeInternational ? istDeloitte(email) : /@deloitte\.de$/.test(email);
      if (!passt) { ausgeblendet++; return; }
      alle.push({
        email,
        displayName: e.DisplayText || e.EntityData?.Title || '',
        jobTitle: e.EntityData?.Title || '',
        department: e.EntityData?.Department || '',
      });
    }));

    // Mehrere Suchwörter: ALLE müssen in Name oder Adresse vorkommen. Die
    // Einzelwort-Varianten oben ziehen sonst Personen an, die nur den Vornamen
    // teilen. Nur übernehmen, wenn dabei etwas übrig bleibt.
    const woerter = roh.replace(/,/g, ' ').split(/\s+/).map(w => w.trim().toLowerCase()).filter(w => w.length >= 2);
    let treffer = alle;
    if (woerter.length >= 2) {
      const enger = alle.filter(u => {
        const heu = (u.displayName + ' ' + u.email).toLowerCase();
        return woerter.every(w => heu.indexOf(w) >= 0);
      });
      if (enger.length > 0) treffer = enger;
    }
    if (treffer.length > 10) treffer = treffer.slice(0, 10);

    // Einige Anfragen sind gescheitert und die übrigen finden nichts: Das ist
    // keine Aussage über die Person.
    if (gescheitert.length > 0 && treffer.length === 0 && ausgeblendet === 0) {
      this.lastPersonSearchStatus = gescheitert[0].status;
      return null;
    }
    return { treffer, ausgeblendet };
  }

  /**
   * Eine Liste auf eigene Berechtigungen bringen — falls sie noch von der Site erbt.
   *
   * `copyRoleAssignments=true`: Was heute gilt, gilt danach weiter; gekappt
   * wird nur die Vererbung, damit die Liste überhaupt eine Zuweisung an eine
   * einzelne Person annimmt. Für die ROLLENLISTE gilt das nicht — sie soll
   * eigene Rechte OHNE die Kopie haben (Owners, sonst niemand);
   * `ensureRolesListPermissions` macht das.
   *
   * true = die Liste hat danach eigene Rechte.
   */
  private async ensureListeEigeneRechte(listName: string): Promise<boolean> {
    try {
      const r = await this._sp.get(
        `${this.list(listName)}?$select=HasUniqueRoleAssignments`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      if (!r.ok) return false;
      const d = await r.json();
      if (d.HasUniqueRoleAssignments ?? d.d?.HasUniqueRoleAssignments) return true;
      const b = await this._post(`${this.list(listName)}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`, {});
      return b.ok;
    } catch (e) {
      console.warn(`[AIUC] ${listName}: Vererbung konnte nicht gekappt werden:`, e);
      return false;
    }
  }

  /**
   * Alle Zuweisungen EINER Liste lesen — dem `nextLink` folgend.
   *
   * `null` = nicht lesbar (Grund in `lastRightsReadError`). Auch eine
   * abgeschnittene Antwort (mehr als 20 Seiten) ist `null`: Ein gekappter
   * Lauf darf nie „steht nicht mehr drin" bedeuten.
   */
  private async leseZuweisungen(listName: string): Promise<ListenZuweisungen | null> {
    const base = this.list(listName);
    const kopf = { headers: { 'Accept': 'application/json;odata=nometadata' } };
    try {
      const flag = await this._sp.get(`${base}?$select=HasUniqueRoleAssignments`, SPHttpClient.configurations.v1, kopf);
      if (!flag.ok) { this.lastRightsReadError = `${listName}: HTTP ${flag.status}`; return null; }
      const fd = await flag.json();
      const out: ListenZuweisungen = {
        eigene: !!(fd.HasUniqueRoleAssignments ?? fd.d?.HasUniqueRoleAssignments),
        nachAdresse: new Map<string, number[]>(),
        nachId: new Map<number, number[]>(),
        adresseZuId: new Map<number, string>(),
      };

      let url: string | null = `${base}/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Email,Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Id&$top=5000`;
      let seiten = 0;
      while (url && seiten < 20) {
        seiten++;
        const r = await this._sp.get(url, SPHttpClient.configurations.v1, kopf);
        if (!r.ok) { this.lastRightsReadError = `${listName}: HTTP ${r.status}`; return null; }
        const d = await r.json();
        const zeilen = (d.value || d.d?.results || []) as SpZuweisung[];
        zeilen.forEach(z => {
          const m = z.Member || {};
          // Gruppen tragen keine Person — nur ausschließen, wenn SharePoint es
          // ausdrücklich sagt (PrincipalType 8 = SharePoint-Gruppe). Fehlt das
          // Feld, zählt der Eintrag: sonst sähe jede Person „ohne Recht" aus.
          if (typeof m.PrincipalType === 'number' && m.PrincipalType !== 1) return;
          const bindings = Array.isArray(z.RoleDefinitionBindings) ? z.RoleDefinitionBindings : (z.RoleDefinitionBindings?.results || []);
          const ids = bindings.map(b => Number(b.Id));
          const adressen: string[] = [];
          if (m.Email) adressen.push(String(m.Email).toLowerCase().trim());
          const lm = String(m.LoginName || '').toLowerCase().match(/[^|]+@[^|\s]+$/);
          if (lm) adressen.push(lm[0].trim());
          adressen.forEach(a => out.nachAdresse.set(a, vereint(out.nachAdresse.get(a), ids)));
          const pid = Number(z.PrincipalId || 0);
          if (pid > 0) {
            out.nachId.set(pid, vereint(out.nachId.get(pid), ids));
            if (adressen[0] && !out.adresseZuId.has(pid)) out.adresseZuId.set(pid, adressen[0]);
          }
        });
        url = d['odata.nextLink'] || d['@odata.nextLink'] || (d.d && d.d.__next) || null;
      }
      if (url) { this.lastRightsReadError = `${listName}: mehr als ${seiten} Seiten Zuweisungen`; return null; }
      return out;
    } catch (e) {
      this.lastRightsReadError = `${listName}: ${e instanceof Error ? e.message : 'Lesen gescheitert'}`;
      console.warn(`[AIUC] Zuweisungen von ${listName} nicht lesbar:`, e);
      return null;
    }
  }

  /**
   * „Rechte prüfen": Liest je Rollen-Zeile nach, ob die Rechte auf den drei
   * Listen WIRKLICH gesetzt sind — schreibt aber nichts. Nachsetzen ist ein
   * eigener Schritt (`repairRoleRights`), damit der Admin erst sieht, was los
   * ist, und dann entscheidet.
   *
   * Gelesen werden die Zuweisungen der drei Listen EINMAL (drei Aufrufe, egal
   * bei wie vielen Personen); aufgelöst (`ensureuser`) wird nur, wenn nach der
   * Adresse etwas fehlt. So kostet der Lauf bei 130 Zeilen ein paar Aufrufe
   * statt Hunderter — die Drosselung ist das Risiko dieser Seite.
   *
   * Drei Befunde, getrennt gemeldet:
   *  - `luecken`: Ein Recht fehlt, das die Rolle braucht.
   *  - `aliase`: Die Rechte sind da, aber unter einer anderen Schreibweise der
   *    Adresse (SMTP gegen UPN/Alias) — kein Fehler der Rechte, aber die Zeile
   *    in der Rollenliste stimmt nicht mit dem Konto überein. DEX v31.67: Ohne
   *    diese Trennung stand eine Person bei JEDEM Lauf als „Lücke" da und
   *    wurde jedes Mal „nachgesetzt".
   *  - `ueberschuss`: Mehr, als die Rolle vorsieht — ein Organizer mit mehr
   *    als Lesen auf der Rollenliste (früherer Admin), oder eine Person mit
   *    Rolle User, die noch direkte Rechte trägt. Nur auf Listen, die eigene
   *    Berechtigungen haben: Auf einer vererbenden Liste stammen die Einträge
   *    von der Site und gehören nicht der App.
   *
   * `null` = die Zuweisungen waren nicht lesbar; `lastRightsReadError` sagt,
   * welche Liste und warum. Das ist KEIN „alles in Ordnung".
   */
  public async auditRoleRights(
    zeilen: Array<{ email: string; name: string; rolle: UserRole }>,
    onProgress?: (done: number, total: number) => void,
  ): Promise<RechteBericht | null> {
    this.lastRightsReadError = '';
    const zuw: Record<string, ListenZuweisungen> = {};
    for (let i = 0; i < RECHTE_SOLL.length; i++) {
      const z = await this.leseZuweisungen(RECHTE_SOLL[i].liste);
      if (!z) return null;
      zuw[RECHTE_SOLL[i].key] = z;
    }

    const pause = (ms: number): Promise<void> => new Promise(res => setTimeout(res, ms));
    // Die Adresse ist der Schlüssel, und in der Rollenliste kann sie zweimal
    // stehen; es gilt die erste Zeile (`RoleContext` nimmt ebenfalls die erste).
    const gesehen: string[] = [];
    const relevant = zeilen.filter(z => {
      const em = (z.email || '').toLowerCase().trim();
      if (!em || gesehen.indexOf(em) >= 0) return false;
      gesehen.push(em);
      return true;
    });
    const bericht: RechteBericht = {
      geprueft: 0, ok: 0, okAdressen: [], luecken: [], ueberschuss: [], aliase: [],
      ohneAdresse: zeilen.filter(z => !(z.email || '').trim()).length,
    };

    for (let i = 0; i < relevant.length; i++) {
      const z = relevant[i];
      const em = z.email.toLowerCase().trim();
      bericht.geprueft++;
      const stufen = (key: RechteListe, id: number | null): number[] =>
        vereint(zuw[key].nachAdresse.get(em), id ? zuw[key].nachId.get(id) : undefined);
      let auffaellig = false;

      if (z.rolle === 'User') {
        // Wer nur „User" ist, braucht nichts — und trägt auch nichts mehr.
        const uebrig = RECHTE_SOLL
          .filter(s => zuw[s.key].eigene && stufen(s.key, null).filter(id => id !== LIMITED_ACCESS).length > 0)
          .map(s => s.key);
        if (uebrig.length > 0) {
          bericht.ueberschuss.push({ email: z.email, name: z.name, rolle: z.rolle, listen: uebrig });
          auffaellig = true;
        }
      } else {
        const braucht = (s: RechteSoll): number => (z.rolle === 'Admin' ? s.admin : s.organizer);
        let fehlt = RECHTE_SOLL.filter(s => !deckt(stufen(s.key, null), braucht(s)));
        let aufloesbar = true;
        let id: number | null = null;
        if (fehlt.length > 0) {
          id = await this.resolveUserId(z.email);
          if (id) {
            const vorher = fehlt.length;
            const konto = id;
            const rest = RECHTE_SOLL.filter(s => !deckt(stufen(s.key, konto), braucht(s)));
            if (rest.length < vorher) {
              const sp = zuw.roles.adresseZuId.get(konto) || zuw.useCases.adresseZuId.get(konto) || zuw.log.adresseZuId.get(konto) || '';
              if (sp && sp !== em) bericht.aliase.push({ name: z.name, email: z.email, spEmail: sp });
            }
            fehlt = rest;
          } else {
            aufloesbar = false;
          }
          await pause(150);
        }
        if (fehlt.length > 0) {
          bericht.luecken.push({ email: z.email, name: z.name, rolle: z.rolle, fehlt: fehlt.map(s => s.key), kontoAufloesbar: aufloesbar });
          auffaellig = true;
        }
        // Ein Organizer soll die Rollenliste LESEN, nicht bearbeiten. Mehr ist
        // meist der Rest einer früheren Admin-Rolle — und damit der Weg, sich
        // selbst wieder hochzustufen.
        if (z.rolle === 'Organizer' && zuw.roles.eigene && deckt(stufen('roles', id), ROLE_DEF.contribute)) {
          bericht.ueberschuss.push({ email: z.email, name: z.name, rolle: z.rolle, listen: ['roles'] });
          auffaellig = true;
        }
      }
      if (!auffaellig) { bericht.ok++; bericht.okAdressen.push(em); }
      if (onProgress) onProgress(i + 1, relevant.length);
    }
    return bericht;
  }

  /**
   * „Fehlende Rechte nachsetzen": Setzt genau die Rechte, die `auditRoleRights`
   * als fehlend gemeldet hat — jedes mit Wiederholung UND Nachlesen
   * (`grantVerified`). Die Person wird EINMAL aufgelöst, nicht je Liste.
   *
   * Vor der Vergabe bekommt die Liste eigene Berechtigungen, falls sie noch
   * erbt (Use-Case-Liste, Protokoll: mit Kopie; Rollenliste: ohne — sie ist
   * ausdrücklich nur für die Owners gedacht).
   *
   * Nur additiv. Was zu VIEL da ist, fasst dieser Schritt nicht an.
   */
  public async repairRoleRights(
    luecken: RechteLuecke[],
    onProgress?: (done: number, total: number) => void,
  ): Promise<RechteReparatur> {
    const ergebnis: RechteReparatur = { behoben: [], offen: [] };
    const pause = (ms: number): Promise<void> => new Promise(res => setTimeout(res, ms));
    const vorbereitet: string[] = [];
    const vorbereiten = async (s: RechteSoll): Promise<void> => {
      if (vorbereitet.indexOf(s.key) >= 0) return;
      vorbereitet.push(s.key);
      if (s.key === 'roles') await this.ensureRolesListPermissions(s.liste);
      else await this.ensureListeEigeneRechte(s.liste);
    };

    for (let i = 0; i < luecken.length; i++) {
      const l = luecken[i];
      // Erneut auflösen, auch wenn der Audit „nicht auflösbar" sagte: Das kann
      // eine Drosselung gewesen sein und ist jetzt vorbei.
      const id = await this.resolveUserId(l.email);
      if (!id) {
        ergebnis.offen.push({ ...l, kontoAufloesbar: false });
      } else {
        const noch: RechteListe[] = [];
        for (let k = 0; k < l.fehlt.length; k++) {
          const s = RECHTE_SOLL.filter(x => x.key === l.fehlt[k])[0];
          if (!s) continue;
          await vorbereiten(s);
          const need = l.rolle === 'Admin' ? s.admin : s.organizer;
          if (!(await this.grantVerified(this.list(s.liste), id, need, `repair/${s.key}`))) noch.push(s.key);
          await pause(150);
        }
        if (noch.length === 0) ergebnis.behoben.push({ email: l.email, name: l.name });
        else ergebnis.offen.push({ ...l, fehlt: noch, kontoAufloesbar: true });
      }
      if (onProgress) onProgress(i + 1, luecken.length);
      await pause(250);
    }
    return ergebnis;
  }

  /**
   * „Überzählige Rechte entziehen": Nimmt einer Person auf den genannten Listen
   * die direkte Zuweisung — und vergibt danach zurück, was ihre Rolle
   * vorsieht (ein Organizer behält Lesen auf der Rollenliste).
   *
   * Rückgabe: die Listen, bei denen es NICHT geklappt hat. Leer = alles weg.
   * Selbstschutz: für die angemeldete Person wird nichts entzogen (und sie
   * steht dann als nicht erledigt in der Rückgabe, statt Erfolg zu behaupten).
   */
  public async entzieheUeberschuss(email: string, rolle: UserRole, listen: RechteListe[]): Promise<RechteListe[]> {
    if (isCurrentUser(this.context, email)) {
      console.warn(`[AIUC] Überschüssige Rechte von ${email} werden NICHT entzogen — das ist die angemeldete Person (Selbstschutz).`);
      return listen;
    }
    const fehler: RechteListe[] = [];
    for (let i = 0; i < listen.length; i++) {
      const s = RECHTE_SOLL.filter(x => x.key === listen[i])[0];
      if (!s) continue;
      if (!(await this.revokeVerified(this.list(s.liste), email, `ueberschuss/${s.key}`))) { fehler.push(s.key); continue; }
      if (rolle === 'Organizer' && s.key === 'roles' && !(await this.grantReadOnRolesList(email))) fehler.push(s.key);
    }
    return fehler;
  }

  // ===================================================================
  // Use Cases
  // ===================================================================

  public async ensureUseCaseList(): Promise<{ isNewlyCreated: boolean }> {
    const name = LIST.useCases;
    const existed = await this.listExists(name);
    if (!existed) {
      await this.createList(name, 'Die Agent-Demos der AI Use Case Platform');
    }

    // Auch auf einer bestehenden Liste: fehlende Spalten nachziehen. Das ist
    // der Weg, auf dem ein spaeteres Feld auf Bestandsdaten kommt, ohne dass
    // jemand SharePoint von Hand anfassen muss.
    //
    // Das Ergebnis wird MITGEZAEHLT. Bis v1.0.1 lief das als best-effort mit
    // einem console.warn — und wenn die Spalten nicht entstanden, antwortete
    // der spaetere `$select` mit HTTP 400 „field does not exist". Der Fehler
    // stand dann in der Oberflaeche als „Die Use Cases konnten nicht geladen
    // werden", und niemand konnte sehen, dass die LISTE das Problem war.
    const fehlend: string[] = [];
    const merke = async (name: string, p: Promise<boolean>): Promise<void> => {
      if (!(await p)) fehlend.push(name);
    };

    await merke('Kurzbeschreibung', this.feldNote(name, 'Kurzbeschreibung'));
    await merke('Beschreibung', this.feldNote(name, 'Beschreibung', true));
    await merke('Bereich', this.feldText(name, 'Bereich'));
    await merke('UcStatus', this.feldAuswahl(name, 'UcStatus', ['Geplant', 'InArbeit', 'Live', 'Archiviert']));
    for (const dim of ['SalesRelevanz', 'Machbarkeit', 'DemoTauglichkeit']) {
      await merke(dim, this.feldAuswahl(name, dim, ['Hoch', 'Mittel', 'Niedrig']));
    }
    await merke('AufrufArt', this.feldAuswahl(name, 'AufrufArt', ['fenster', 'eingebettet']));
    for (const link of ['LinkSourceCode', 'LinkDeployment', 'LinkGuide', 'LinkWiki', 'LinkVideo', 'BildUrl']) {
      await merke(link, this.feldNote(name, link));
    }
    await merke('Reihenfolge', this.feldZahl(name, 'Reihenfolge'));
    for (const f of ['BetreuerEmails', 'BetreuerNamen', 'Schlagworte']) {
      await merke(f, this.feldNote(name, f));
    }
    this.fehlendeSpalten = fehlend;
    if (fehlend.length > 0) {
      console.warn(`[AIUC] Diese Spalten fehlen auf ${name}: ${fehlend.join(', ')}`);
    }

    return { isNewlyCreated: !existed };
  }

  public async ensureLogList(): Promise<void> {
    const name = LIST.log;
    if (!(await this.listExists(name))) {
      await this.createList(name, 'Aenderungsprotokoll der AI Use Case Platform');
    }
    await this.feldZahl(name, 'UseCaseId');
    await this.feldText(name, 'Aktion');
    await this.feldNote(name, 'Detail');
    await this.feldText(name, 'Wer');
  }

  private mapUseCase(row: SpUseCaseRow): UseCase {
    const bew = (v?: string): Bewertung =>
      (v === 'Hoch' || v === 'Mittel' || v === 'Niedrig') ? v : 'unbewertet';
    const liste = (v?: string): string[] =>
      (v || '').split(';').map(s => s.trim()).filter(Boolean);
    return {
      id: row.Id,
      titel: row.Title || '',
      kurzbeschreibung: row.Kurzbeschreibung || '',
      beschreibung: row.Beschreibung || '',
      bereich: row.Bereich || '',
      status: (['Geplant', 'InArbeit', 'Live', 'Archiviert'].indexOf(row.UcStatus || '') >= 0
        ? row.UcStatus as UseCaseStatus
        : 'Geplant'),
      salesRelevanz: bew(row.SalesRelevanz),
      machbarkeit: bew(row.Machbarkeit),
      demoTauglichkeit: bew(row.DemoTauglichkeit),
      // Vorgabe ist `fenster` — siehe types/index.ts, Permissions Policy.
      aufrufArt: (row.AufrufArt === 'eingebettet' ? 'eingebettet' : 'fenster') as AufrufArt,
      ressourcen: {
        sourceCode: row.LinkSourceCode || '',
        deployment: row.LinkDeployment || '',
        deploymentGuide: row.LinkGuide || '',
        wiki: row.LinkWiki || '',
        video: row.LinkVideo || '',
      },
      bildUrl: row.BildUrl || '',
      reihenfolge: typeof row.Reihenfolge === 'number' ? row.Reihenfolge : 999,
      betreuerEmails: liste(row.BetreuerEmails),
      betreuerNamen: liste(row.BetreuerNamen),
      schlagworte: liste(row.Schlagworte),
      geaendertAm: row.Modified || '',
      // Ohne `$expand` gibt es den Namen nicht mehr. Das ist der Preis
      // dafuer, dass das Lesen nicht an einer fehlenden Spalte scheitert —
      // und ein fehlender Name ist harmloser als eine leere Plattform.
      geaendertVon: row.Editor?.Title || '',
    };
  }

  /**
   * Alle Use Cases lesen.
   *
   * `null` heisst nicht lesbar. Eine leere Kachelwand und ein Rechte-Fehler
   * sehen sonst gleich aus — und die Kachelwand ist die ganze App.
   *
   * **Kein `$select`, kein `$expand`, kein `$orderby`.** Bis v1.0.1 stand hier
   * eine Liste von zwanzig Spaltennamen. Fehlt EINE davon — weil das Anlegen
   * still gescheitert ist —, antwortet SharePoint mit HTTP 400 „field does not
   * exist", und die ganze Seite ist leer. Genau so gemeldet am 10.09.2026.
   * Die Abfrage nennt jetzt keine Spalte mehr: Was da ist, wird gelesen; was
   * fehlt, faellt in `mapUseCase` auf seinen Vorgabewert. Sortiert wird im
   * Browser — bei ein paar Dutzend Demos kostet das nichts und kann nicht
   * scheitern, weil `Reihenfolge` mal fehlt.
   */
  public async getUseCases(): Promise<UseCase[] | null> {
    this.lastUseCasesReadStatus = 0;
    this.lastReadError = '';
    try {
      const r = await this._sp.get(
        `${this.list(LIST.useCases)}/items?$top=500`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      this.lastUseCasesReadStatus = r.status;
      if (!r.ok) {
        this.lastReadError = await this.fehlertext(r);
        console.warn(`[AIUC] Use Cases lesen: HTTP ${r.status} — ${this.lastReadError}`);
        return null;
      }
      const d = await r.json();
      const rows = ((d.value || []) as SpUseCaseRow[]).map(row => this.mapUseCase(row));
      rows.sort((a, b) => (a.reihenfolge - b.reihenfolge) || a.titel.localeCompare(b.titel, 'de'));
      return rows;
    } catch (e) {
      this.lastReadError = String(e);
      return null;
    }
  }

  private toRow(uc: Partial<UseCase>): Record<string, unknown> {
    // Kein `__metadata` mehr: Den Typ setzen `_postItem`/`_mergeItem`, und
    // zwar den, den die Liste selbst nennt (s. `entityType`).
    const body: Record<string, unknown> = {};
    if (uc.titel !== undefined) body.Title = uc.titel;
    if (uc.kurzbeschreibung !== undefined) body.Kurzbeschreibung = uc.kurzbeschreibung;
    if (uc.beschreibung !== undefined) body.Beschreibung = uc.beschreibung;
    if (uc.bereich !== undefined) body.Bereich = uc.bereich;
    if (uc.status !== undefined) body.UcStatus = uc.status;
    // `unbewertet` ist in SharePoint kein Choice-Wert, sondern die Abwesenheit
    // eines Werts — sonst stuende „unbewertet" als vierte Option im Filter.
    if (uc.salesRelevanz !== undefined) body.SalesRelevanz = uc.salesRelevanz === 'unbewertet' ? null : uc.salesRelevanz;
    if (uc.machbarkeit !== undefined) body.Machbarkeit = uc.machbarkeit === 'unbewertet' ? null : uc.machbarkeit;
    if (uc.demoTauglichkeit !== undefined) body.DemoTauglichkeit = uc.demoTauglichkeit === 'unbewertet' ? null : uc.demoTauglichkeit;
    if (uc.aufrufArt !== undefined) body.AufrufArt = uc.aufrufArt;
    if (uc.ressourcen) {
      body.LinkSourceCode = uc.ressourcen.sourceCode || '';
      body.LinkDeployment = uc.ressourcen.deployment || '';
      body.LinkGuide = uc.ressourcen.deploymentGuide || '';
      body.LinkWiki = uc.ressourcen.wiki || '';
      body.LinkVideo = uc.ressourcen.video || '';
    }
    if (uc.bildUrl !== undefined) body.BildUrl = uc.bildUrl;
    if (uc.reihenfolge !== undefined) body.Reihenfolge = uc.reihenfolge;
    if (uc.betreuerEmails !== undefined) body.BetreuerEmails = uc.betreuerEmails.join('; ');
    if (uc.betreuerNamen !== undefined) body.BetreuerNamen = uc.betreuerNamen.join('; ');
    if (uc.schlagworte !== undefined) body.Schlagworte = uc.schlagworte.join('; ');
    return body;
  }

  public async createUseCase(uc: Partial<UseCase>): Promise<number | null> {
    try {
      const r = await this._postItem(LIST.useCases, this.toRow(uc));
      if (!r.ok) {
        // Der Grund gehoert in die Konsole. Vorher stand hier nur `null`, und
        // die Erstbefuellung legte stumm fuenfmal nichts an.
        this.lastReadError = await this.fehlertext(r);
        console.error(`[AIUC] createUseCase: HTTP ${r.status} — ${this.lastReadError}`);
        return null;
      }
      const d = await r.json();
      const id = d.d?.Id ?? d.Id;
      return typeof id === 'number' ? id : null;
    } catch (e) {
      console.error('[AIUC] createUseCase:', e);
      return null;
    }
  }

  public async updateUseCase(id: number, uc: Partial<UseCase>): Promise<boolean> {
    try {
      const r = await this._mergeItem(LIST.useCases, id, this.toRow(uc));
      if (!r.ok) {
        this.lastReadError = await this.fehlertext(r);
        console.error(`[AIUC] updateUseCase: HTTP ${r.status} — ${this.lastReadError}`);
      }
      return r.ok;
    } catch (e) {
      console.error('[AIUC] updateUseCase:', e);
      return false;
    }
  }

  /**
   * Einen Use Case löschen — in den Papierkorb, nicht endgültig.
   *
   * v1.3: Bis v1.2 lief das über einen REST-DELETE, und der landet NICHT im
   * Papierkorb. Eine Kachel trägt Wochen an Pflegearbeit (lange Beschreibung,
   * fünf Links, Bewertung) und ab v1.3 ein hochgeladenes Bild als Anhang —
   * alles weg mit einem Klick auf „Löschen". In DEX hat genau das am
   * 15.09.2026 einen Termin mit 64 Eingeladenen den einzigen Verweis
   * gekostet (v31.62). `recycle` kostet nichts: 93 Tage Papierkorb der Site,
   * ein Admin holt die Zeile samt Anhang zurück.
   */
  public async deleteUseCase(id: number): Promise<boolean> {
    try {
      const r = await this._post(`${this.list(LIST.useCases)}/items(${id})/recycle`, {});
      return r.ok;
    } catch (e) {
      console.error('[AIUC] deleteUseCase:', e);
      return false;
    }
  }

  // ===================================================================
  // Kachelbild (Anhang an der Use-Case-Zeile)
  // ===================================================================

  /**
   * Ein Kachelbild als Anhang an die Use-Case-Zeile hängen.
   *
   * Warum ein Anhang und keine Datei in SiteAssets: Wer die Zeile bearbeiten
   * darf, darf auch Anhänge hinzufügen — ein Organizer (Edit auf der Liste)
   * braucht damit kein zweites Recht auf einer Dokumentbibliothek. Und der
   * Anhang hängt an der Zeile: Recycelt man den Use Case, geht das Bild mit
   * in den Papierkorb und kommt beim Wiederherstellen mit zurück.
   *
   * Der Dateiname ist jedes Mal neu (Zeitstempel). Das ist Absicht: Er macht
   * die URL eindeutig, und ein Browser, der das alte Bild gecacht hat, holt
   * das neue, statt den alten Stand zu zeigen.
   *
   * Rückgabe: URL und Dateiname — oder `null`, wenn das Bild NICHT angekommen
   * ist. Eine URL wird nie geraten: Die geratene Adresse zeigte in DEX auf
   * eine Datei, die nie ankam, und die Kachel blieb dauerhaft weiß (v29.34).
   */
  public async uploadBild(useCaseId: number, file: File): Promise<{ url: string; name: string } | null> {
    const ext = file.type === 'image/png' ? 'png' : file.type === 'image/webp' ? 'webp' : 'jpg';
    const name = `${BILD_PREFIX}${Date.now().toString(36)}.${ext}`;
    const zeile = `${this.list(LIST.useCases)}/items(${useCaseId})`;
    try {
      const r = await this._sp.post(
        `${zeile}/AttachmentFiles/add(FileName='${encodeURIComponent(name)}')`,
        SPHttpClient.configurations.v1,
        // Dieselben Kopfzeilen wie der Bild-Upload in DEX (`eventAssets.ts`) —
        // dort seit Monaten im Einsatz, und bei Binärdaten ist „sieht
        // gleichwertig aus" genau die Art Abweichung, die erst im Betrieb
        // auffällt.
        { headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' }, body: file } as ISPHttpClientOptions,
      );
      if (!r.ok) {
        this.lastReadError = await this.fehlertext(r);
        console.warn(`[AIUC] Kachelbild hochladen: HTTP ${r.status} — ${this.lastReadError}`);
        return null;
      }
      let rel = '';
      try {
        const d = await r.json();
        rel = d.ServerRelativeUrl || d.d?.ServerRelativeUrl || '';
      } catch { /* Antwort ohne lesbaren Körper — unten nachschlagen */ }
      if (!rel) {
        // Angekommen, aber ohne URL in der Antwort: die Adresse NACHSCHLAGEN.
        const l = await this._sp.get(
          `${zeile}/AttachmentFiles`,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata' } },
        );
        if (l.ok) {
          const ld = await l.json();
          const treffer = ((ld.value || ld.d?.results || []) as Array<{ FileName?: string; ServerRelativeUrl?: string }>)
            .filter(f => f.FileName === name)[0];
          rel = treffer && treffer.ServerRelativeUrl ? treffer.ServerRelativeUrl : '';
        }
      }
      if (!rel) {
        console.warn('[AIUC] Kachelbild hochgeladen, aber die Adresse ließ sich nicht ermitteln.');
        return null;
      }
      return { url: `${window.location.origin}${rel}`, name };
    } catch (e) {
      console.warn('[AIUC] uploadBild:', e);
      return null;
    }
  }

  /**
   * Kachelbild-Anhänge einer Use-Case-Zeile in den Papierkorb legen.
   *
   * `behalte` nimmt genau EIN Bild aus (das gerade hochgeladene), `nur` räumt
   * genau EINES ab (Rückbau nach einem fehlgeschlagenen Speichern). Ohne
   * Angabe gehen alle Bilder der App weg. Angefasst werden nur Anhänge mit
   * dem Präfix `BILD_PREFIX`.
   *
   * Rückgabe false heißt „mindestens ein Bild ist liegen geblieben" — harmlos
   * (ein verwaister Anhang kostet Platz, nichts sonst), deshalb ruft das
   * niemand als Bedingung für Erfolg auf.
   */
  public async entferneBilder(useCaseId: number, opts: { behalte?: string; nur?: string } = {}): Promise<boolean> {
    const zeile = `${this.list(LIST.useCases)}/items(${useCaseId})`;
    try {
      const l = await this._sp.get(
        `${zeile}/AttachmentFiles`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      if (!l.ok) return false;
      const ld = await l.json();
      const namen = ((ld.value || ld.d?.results || []) as Array<{ FileName?: string }>).map(f => f.FileName || '');
      let alleWeg = true;
      for (const fn of namen) {
        if (fn.indexOf(BILD_PREFIX) !== 0) continue;
        if (opts.behalte && fn === opts.behalte) continue;
        if (opts.nur && fn !== opts.nur) continue;
        // Nacheinander — mehrere gleichzeitige POSTs gegen dieselbe Zeile sind
        // das Muster, das SharePoint drosselt.
        // eslint-disable-next-line no-await-in-loop
        const rr = await this._post(`${zeile}/AttachmentFiles/getByFileName('${encodeURIComponent(fn)}')/recycleObject`, {});
        if (!rr.ok) alleWeg = false;
      }
      return alleWeg;
    } catch (e) {
      console.warn('[AIUC] entferneBilder:', e);
      return false;
    }
  }

  // ===================================================================
  // Protokoll
  // ===================================================================

  /**
   * Das Änderungsprotokoll lesen — die letzten 500 Einträge, neueste zuerst.
   *
   * `null` heißt nicht lesbar und ist etwas anderes als `[]` (noch nichts
   * protokolliert) — dieselbe Regel wie bei den Use Cases.
   *
   * **Kein `$filter` auf `UseCaseId`.** Die Spalte wird in `ensureLogList`
   * best-effort angelegt; fehlt sie, antwortet SharePoint auf einen `$filter`
   * mit HTTP 400 statt mit „0 Treffer" (dieselbe Falle wie `StarterType` in
   * den DEX-Flows, 01.09.2026). Deshalb wird alles gelesen und im Browser
   * eingeengt. Sortiert wird nach `Id` — ein Systemfeld, das es immer gibt.
   */
  public async getLog(): Promise<LogEintrag[] | null> {
    this.lastLogReadStatus = 0;
    this.lastLogReadError = '';
    try {
      const r = await this._sp.get(
        `${this.list(LIST.log)}/items?$top=500&$orderby=Id%20desc`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      this.lastLogReadStatus = r.status;
      if (!r.ok) {
        this.lastLogReadError = await this.fehlertext(r);
        console.warn(`[AIUC] Protokoll lesen: HTTP ${r.status} — ${this.lastLogReadError}`);
        return null;
      }
      const d = await r.json();
      return ((d.value || []) as SpLogRow[]).map(row => ({
        id: row.Id,
        useCaseId: typeof row.UseCaseId === 'number' ? row.UseCaseId : 0,
        // Fehlt die Spalte `Aktion`, steht dieselbe Angabe noch im Titel
        // (`log` schreibt beides).
        aktion: row.Aktion || row.Title || '',
        detail: row.Detail || '',
        wer: row.Wer || '',
        wann: row.Created || '',
      }));
    } catch (e) {
      this.lastLogReadError = String(e);
      return null;
    }
  }

  /**
   * Darf die Plattform mit den Start-Use-Cases befuellt werden?
   *
   * Drei Bedingungen, und ALLE drei muessen belegt sein — nicht bloss nicht
   * widerlegt: Die Use-Case-Liste ist lesbar UND leer, und im Protokoll steht
   * kein Merker einer frueheren Befuellung. Ist eine der beiden Listen nicht
   * lesbar, ist die Antwort `false`: „Ich weiss es nicht" darf hier nicht wie
   * „ist leer" wirken, sonst liegen fuenf Kacheln neben einem Bestand, den
   * gerade nur niemand sehen konnte (die Lehre aus DEX v30.37).
   */
  public async darfErstbefuellen(): Promise<boolean> {
    const rows = await this.getUseCases();
    if (rows === null) {
      console.warn('[AIUC] Erstbefüllung übersprungen — die Use-Case-Liste war nicht lesbar.');
      return false;
    }
    if (rows.length > 0) return false;
    try {
      const r = await this._sp.get(
        `${this.list(LIST.log)}/items?$filter=Aktion eq '${SEED_MARKER}'&$select=Id&$top=1`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      // 404 = Protokollliste gibt es (noch) nicht. Dann kann es auch keine
      // frühere Befüllung gegeben haben.
      if (r.status === 404) return true;
      if (!r.ok) {
        console.warn(`[AIUC] Erstbefüllung übersprungen — Protokoll nicht lesbar (HTTP ${r.status}).`);
        return false;
      }
      const d = await r.json();
      return (d.value || []).length === 0;
    } catch {
      return false;
    }
  }

  /** Den Merker setzen, damit die Erstbefuellung genau einmal passiert. */
  public async merkeErstbefuellung(anzahl: number): Promise<void> {
    await this.log(0, SEED_MARKER, `${anzahl} Start-Use-Cases angelegt`);
  }

  /** Protokollzeile schreiben. Best-effort — ein fehlendes Protokoll darf keine Aktion verhindern. */
  public async log(useCaseId: number, aktion: string, detail: string): Promise<void> {
    try {
      await this._postItem(LIST.log, {
        'Title': aktion,
        'UseCaseId': useCaseId,
        'Aktion': aktion,
        'Detail': detail,
        'Wer': this.context.pageContext.user.email,
      });
    } catch { /* Protokoll ist Nebenbuchhaltung */ }
  }
}
