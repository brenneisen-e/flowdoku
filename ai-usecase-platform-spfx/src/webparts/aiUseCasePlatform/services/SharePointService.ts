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
 *    `grantVerified` wiederholt und liest danach `roledefinitionbindings` —
 *    ein `addroleassignment`-POST allein ist keine Vergabe. In DEX fehlten
 *    18 von 126 Rollen-Eintraegen mindestens ein Recht, und jede dieser
 *    Zuweisungen hatte „erfolgreich" gemeldet (v30.85).
 *
 * 3. **Wer Rechte vergibt, baut den Entzug im selben Commit.**
 *    `entzieheListe` und `reduziereListe` lesen danach nach; SharePoint antwortet auf
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
 * Sie muss zu den Vergaben passen: `grantRoleRights` und `grantListRights`
 * lesen diese Tabelle selbst — es gibt keine zweite Liste von Stufen. Steht
 * irgendwo im Code eine andere Stufe als hier, meldet der
 * Audit „fehlt" für ein Recht, das die Vergabe nie setzt — und die Meldung
 * kommt bei jedem Lauf wieder. Die Matrix in `data/rollenMatrix.ts` (Kategorie
 * „SharePoint") beschreibt dieselben Zeilen in Sätzen.
 */
export type RechteListe = 'roles' | 'useCases' | 'log';

interface RechteSoll { key: RechteListe; liste: string; admin: number; organizer: number }

/**
 * v1.3 (Review-Nachzug): Ein Admin bekommt auf ALLEN drei Listen Vollzugriff.
 * `addroleassignment`, `removeroleassignment`, `breakroleinheritance` und das
 * Lesen der `roleassignments` verlangen „Berechtigungen verwalten" auf der
 * jeweiligen Liste — das steckt in Full Control (oder Site-Owner), NICHT in
 * Edit/Contribute. Mit Edit auf den Use Cases und Contribute im Protokoll
 * scheiterte ein Admin, der nicht Site-Owner ist, beim Vergeben und beim
 * „Rechte prüfen" auf genau diesen zwei Listen (403, je 5,5 s Wartezeit).
 * Ein Organizer bekommt weiter nur, was er zum Pflegen braucht.
 */
const RECHTE_SOLL: RechteSoll[] = [
  { key: 'roles', liste: LIST.roles, admin: ROLE_DEF.full, organizer: ROLE_DEF.read },
  // Contribute reicht: Zeilen anlegen, ändern, recyceln und Anhänge tragen nur Add/Edit/Delete Items.
  // „Edit“ enthielte zusätzlich „Manage Lists“ — ein Organizer könnte Spalten oder die ganze Liste
  // löschen (Sicherheits-Review 29.09.2026).
  { key: 'useCases', liste: LIST.useCases, admin: ROLE_DEF.full, organizer: ROLE_DEF.contribute },
  { key: 'log', liste: LIST.log, admin: ROLE_DEF.full, organizer: ROLE_DEF.contribute },
];

/**
 * Ein Recht, das gesetzt oder entzogen werden sollte — als CODE, nicht als
 * Satz. Der Service kennt keine Sprache; die Oberfläche formuliert mit `t()`
 * (Review-Fund 19: deutsche Klartexte und ASCII-Umschreibungen wie
 * „unvollstaendig" gelangten bis in englische Meldungen).
 *
 * `<liste>:<stufe>` = diese Stufe konnte nicht gesetzt werden; `konto` = die
 * Person ließ sich nicht auflösen (ausgeschieden oder gedrosselt).
 */
export type RechteCode =
  | 'roles:read' | 'roles:full'
  | 'useCases:edit' | 'useCases:contribute' | 'useCases:full'
  | 'log:contribute' | 'log:full'
  | 'konto';

/** Ergebnis eines Entzugs — Codes statt `true`/`false`, damit „Selbstschutz" nicht wie „erledigt" aussieht. */
export type EntzugsCode =
  /** Entzogen UND nachgelesen. */
  | 'ok'
  /** Das ist das eigene Konto (auch unter einer anderen Schreibweise der Adresse) — nichts wurde angefasst. */
  | 'selbst'
  /** Das Konto ließ sich nicht auflösen und war auf den Listen nicht zu finden — der Entzug ist nicht belegt. */
  | 'konto'
  /** Die Zuweisungen waren nicht lesbar — nicht belegt. */
  | 'lesefehler'
  /** Der Entzug ging durch, das Nachlesen zeigt die Rechte aber noch. */
  | 'offen';

/** Wie eine Liste beim Nachschlagen dasteht: es gibt sie, es gibt sie nicht (404), oder wir wissen es nicht. */
export type ListenStatus = 'ja' | 'nein' | 'unbekannt';

/** Warum das Lesen der Rechte scheiterte — Code, keine Prosa. */
export interface LeseFehler {
  code: 'http' | 'zu-gross' | 'ausnahme' | 'keine-antwort';
  /** `rollen` = die Rollenliste selbst (Inhalt), sonst die Liste, deren Zuweisungen gelesen wurden. */
  liste: RechteListe | 'rollen';
  status: number;
}

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

/** Zwei Adress-Listen zusammenlegen (ohne Doppelte). */
function vereint2(a: string[] | undefined, b: string[] | undefined): string[] {
  const out: string[] = [];
  (a || []).concat(b || []).forEach(x => { if (out.indexOf(x) < 0) out.push(x); });
  return out;
}

/** Ergebnis von `entzieheUeberschuss` — getrennte Zustände je Liste (siehe dort). */
export interface UeberschussErgebnis {
  entzogen: RechteListe[];
  lesenFehlt: RechteListe[];
  offen: RechteListe[];
  /** Das war das eigene Konto: nichts wurde angefasst. */
  selbst: boolean;
  /** Mindestens eine Liste war nicht lesbar — der Entzug ist dort nicht belegt. */
  lesefehler: boolean;
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
  /**
   * Rechte, die die Person NICHT direkt, sondern über eine SharePoint-Gruppe der
   * Liste hat (Owners, Members …). Kein Befund und keine Lücke: „Nachsetzen"
   * würde dort nur ein redundantes Direktrecht danebenlegen. Eigene Rubrik,
   * damit die Oberfläche „über Gruppe gedeckt" sagen kann statt „ok" oder
   * „Rechte fehlen".
   */
  ueberGruppe: Array<{ email: string; name: string; rolle: UserRole; listen: RechteListe[] }>;
  /**
   * Adressen (klein, getrimmt) von Zeilen, die beim Auflösen zum ANGEMELDETEN
   * Konto führten, obwohl sie nicht die bekannten Schreibweisen der Person sind
   * (ein Alias in der Rollenliste). Die Oberfläche sperrt diese Zeilen.
   */
  eigeneKonten: string[];
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
  /** Alle bekannten Schreibweisen je Konto (Email und Adresse aus dem LoginName). */
  adressenZuId: Map<number, string[]>;
  /** SharePoint-Gruppen (PrincipalType 8) mit ihren Stufen auf dieser Liste. */
  gruppen: Array<{ id: number; stufen: number[] }>;
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

/**
 * Gehört diese Domain zu Deloitte? Geprüft wird die REGISTRIERBARE Domain, nicht
 * irgendein Label darin (Review-Fund 17): Der frühere Filter `(^|\.)deloitte…\.`
 * ließ `deloitte.evil.com`, `deloitte-fake.net` und `deloittecom.evil.org`
 * durch — ein Gast-Konto mit solcher Adresse erschien als normaler Treffer.
 *
 * Erlaubt sind (mit beliebigen Subdomains davor):
 *  - `deloitte.<endung>` — `deloitte.de`, `deloitte.com`, `deloitte.co.uk` (Endungen
 *    mit zweiter Ebene wie `co.uk` / `com.au` zählen als EINE Endung);
 *  - `deloitte<zusatz>.com` — der Zusatz nur aus Buchstaben und Ziffern, ohne
 *    Bindestrich (deloitteCE.com).
 * Restrisiko: Eine Endung wie `co.cc` wäre selbst frei registrierbar; das ist
 * hier nicht ausgeschlossen. Eine feste Liste der Member-Firm-Domains wäre die
 * strengere Lösung, aber auch die, die ein Mensch pflegen müsste.
 */
export function istDeloitteDomain(domain: string): boolean {
  const teile = (domain || '').toLowerCase().trim().split('.');
  if (teile.length < 2 || teile.some(t => !/^[a-z0-9-]+$/.test(t))) return false;
  const endung = teile[teile.length - 1];
  const zweite = teile[teile.length - 2];
  const mitZweiterEbene = teile.length >= 3 && endung.length === 2 && /^(co|com|org|net|ac|gov|edu)$/.test(zweite);
  const name = mitZweiterEbene ? teile[teile.length - 3] : zweite;
  const tld = mitZweiterEbene ? `${zweite}.${endung}` : endung;
  if (name === 'deloitte') return true;
  return tld === 'com' && /^deloitte[a-z0-9]{1,12}$/.test(name);
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
   * Warum `getRoles` null lieferte, wenn es kein HTTP-Status ist: die Liste ist
   * zu groß für den Lesepfad (`zu-gross`). Ein Code, keine Prosa.
   */
  public lastRolesReadCode: '' | 'zu-gross' = '';

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
    /**
     * v1.3 (Review-Fund 20): POST OHNE Wiederholung — für Anfragen, die nur
     * LESEN, obwohl sie ein POST sind (Personensuche des People-Pickers). Die
     * Schranke wiederholt bis zu zweimal mit Retry-After (bis 45 s); für eine
     * Suche, deren Ergebnis beim nächsten Tastendruck verworfen wird, wäre das
     * reine Wartezeit — und jede Wiederholung zählt aufs Kontingent.
     */
    postEinmal: (url: string, cfg: SPHttpClientConfiguration, options?: ISPHttpClientOptions): Promise<SPHttpClientResponse> =>
      this.context.spHttpClient.post(url, cfg, options),
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

  /** POST ohne Wiederholung (Suche) — dieselben Kopfzeilen wie `_post`. */
  private async _postEinmal(url: string, body: object): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
      },
      body: JSON.stringify(body),
    };
    return this._sp.postEinmal(url, SPHttpClient.configurations.v1, options);
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

  /**
   * Gibt es die Liste? DREI Antworten, nicht zwei (Review-Fund 11).
   *
   * Bis v1.3 machte `listExists` aus JEDEM Fehlerstatus „gibt es nicht". Für
   * einen gewöhnlichen User ist die Rollenliste unlesbar (403) — sie galt damit
   * als nicht vorhanden, und jeder Seitenaufruf jeder Person löste `createList`
   * plus je vier Spalten-Probes und -Anlagen aus, alle aussichtslos. Das
   * verletzt „erst prüfen, dann anlegen" und „ein Lesefehler ist keine Null".
   *
   *  - `ja`        — 200: die Liste ist da und lesbar.
   *  - `nein`      — 404, und NUR 404: die Liste gibt es nicht. Nur dann darf angelegt werden.
   *  - `unbekannt` — 403, 429, 5xx, Ausnahme: wir wissen es nicht. Es wird nichts geschrieben.
   */
  public async listStatus(listName: string): Promise<ListenStatus> {
    try {
      const r = await this._sp.get(this.list(listName), SPHttpClient.configurations.v1);
      if (r.ok) return 'ja';
      return r.status === 404 ? 'nein' : 'unbekannt';
    } catch {
      return 'unbekannt';
    }
  }

  private _basisRechteMerker: Promise<number | null> | null = null;

  /**
   * Darf die angemeldete Person auf dieser Site Listen verwalten (Berechtigung
   * „Manage Lists")? Gilt für das Anlegen von Listen und Spalten. Wer es nicht
   * darf, dem sagt SharePoint 403 — die Anfrage vorher zu stellen ist nur Lärm
   * und Kontingent. Nicht lesbar heißt „nein": im Zweifel nichts schreiben.
   * Einmal je Sitzung gelesen und gemerkt (lazy — nur dort, wo wirklich
   * geschrieben werden soll).
   */
  private kannListenVerwalten(): Promise<boolean> {
    // Bit 11 (0x800) = ManageLists. `&` rechnet auf 32 Bit — die unteren Bits
    // bleiben auch bei Werten über 2^31 erhalten.
    return this.basisRechte().then(low => low !== null && (low & 0x800) !== 0);
  }

  /**
   * Darf dieses Konto Berechtigungen setzen (ManagePermissions, Bit 25 = 0x2000000)?
   *
   * `breakroleinheritance` und `addroleassignment` brauchen dieses Recht, NICHT
   * ManageLists. Wer die Listen anlegen darf (Edit-Mitglied), kann die Rollenliste
   * deshalb noch lange nicht sperren — der Guard über `kannListenVerwalten` allein
   * ließ solche Konten bis an den POST kommen, der dann mit 403 scheiterte
   * (Review 29.09.2026, Fund 8).
   */
  private kannRechteVerwalten(): Promise<boolean> {
    return this.basisRechte().then(low => low !== null && (low & 0x2000000) !== 0);
  }

  /** Die unteren 32 Bit der effektiven Web-Rechte — `null` = nicht lesbar. Einmal je Sitzung gemerkt. */
  private basisRechte(): Promise<number | null> {
    if (!this._basisRechteMerker) {
      this._basisRechteMerker = (async (): Promise<number | null> => {
        try {
          const r = await this._sp.get(
            `${this.siteUrl}/_api/web/effectivebasepermissions`,
            SPHttpClient.configurations.v1,
            { headers: { 'Accept': 'application/json;odata=nometadata' } },
          );
          // Nicht lesbar heißt „nein" für DIESEN Aufruf, wird aber nicht gemerkt:
          // Eine Drosselung beim ersten Start würde sonst die Erstinstallation
          // für die ganze Sitzung sperren.
          if (!r.ok) { this._basisRechteMerker = null; return null; }
          const d = await r.json();
          const low = Number(d.Low ?? d.d?.EffectiveBasePermissions?.Low ?? d.d?.Low);
          if (isNaN(low)) { this._basisRechteMerker = null; return null; }
          return low;
        } catch {
          this._basisRechteMerker = null;
          return null;
        }
      })();
    }
    return this._basisRechteMerker;
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
      // 403 und 429 sagen NICHTS über die Spalte (kein Recht, gedrosselt) — weder anlegen
      // noch als fehlend melden: Ein Anlegeversuch auf eine vorhandene Spalte scheitert bei
      // SharePoint mit 400/500 (Lehre 4 im Kopfkommentar), und „Spalte fehlt" behauptete
      // bei normalen Nutzern etwas Unbelegtes (Review 29.09.2026).
      if (probe.status === 403) return true;
      if (probe.status === 429) {
        // Gedrosselt: EIN neuer Versuch nach kurzer Pause. Ohne ihn galt die Spalte als „vorhanden",
        // obwohl niemand es wusste — fehlte sie, scheiterten später Anlegen und Protokoll mit 400
        // (Gegenprüfung 29.09.2026). Bleibt es bei 429, zählt sie als vorhanden: nichts anlegen,
        // nichts Unbelegtes melden; der nächste Start eines Organizers prüft erneut.
        await new Promise<void>(res => setTimeout(res, 1500));
        const zweite = await this._sp.get(
          `${this.list(listName)}/fields/getbytitle('${encodeURIComponent(internalName)}')`,
          SPHttpClient.configurations.v1,
        );
        if (zweite.ok || zweite.status === 429 || zweite.status === 403) return true;
      }
    } catch { /* nicht lesbar -> Anlegen versuchen */ }

    // v1.3 (Fund 11): Wer keine Listen verwalten darf, bekommt auf das Anlegen
    // ein sicheres 403 — die Spalte fehlt dann eben, und `fehlendeSpalten` sagt es.
    if (!(await this.kannListenVerwalten())) return false;

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

  /**
   * Eine Liste anlegen — aber nur, wenn das möglich ist. `true` = der Aufruf
   * wurde abgeschickt und von SharePoint angenommen. Ob die Liste danach da
   * ist, klärt der nächste `listStatus`; hier zählt nur, dass eine Person ohne
   * Recht (Fund 11) gar nicht erst anfragt.
   */
  private async createList(listName: string, description: string): Promise<boolean> {
    if (!(await this.kannListenVerwalten())) {
      console.warn(`[AIUC] ${listName} fehlt, aber dieses Konto darf keine Listen anlegen — es wird nichts geschrieben.`);
      return false;
    }
    try {
      const r = await this._post(`${this.siteUrl}/_api/web/lists`, {
        '__metadata': { 'type': 'SP.List' },
        'Title': listName,
        'Description': description,
        'BaseTemplate': 100,
        'AllowContentTypes': false,
      });
      return r.ok;
    } catch (e) {
      console.warn(`[AIUC] ${listName} konnte nicht angelegt werden:`, e);
      return false;
    }
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
  public ensureRolesList(): Promise<{ isNewlyCreated: boolean; status: ListenStatus }> {
    return this.einmalig('roles', () => this.ensureRolesListNeu());
  }

  /**
   * Zwei gleichzeitige Aufrufe teilen sich EINEN Lauf. Beim allerersten Start
   * rufen `RoleProvider` (Erstinstallation) und `UseCaseProvider` dieselben
   * `ensure*`-Methoden gleichzeitig; ohne das würden beide eine fehlende Liste
   * anlegen, und der zweite Aufruf bekäme von SharePoint einen Fehler
   * („Liste existiert schon") und ließe die Spalten aus.
   */
  private _laeufe: Record<string, Promise<unknown>> = {};

  private einmalig<T>(key: string, fn: () => Promise<T>): Promise<T> {
    const laufend = this._laeufe[key];
    if (laufend) return laufend as Promise<T>;
    const p = fn().then(
      r => { delete this._laeufe[key]; return r; },
      e => { delete this._laeufe[key]; throw e; },
    );
    this._laeufe[key] = p;
    return p;
  }

  private async ensureRolesListNeu(): Promise<{ isNewlyCreated: boolean; status: ListenStatus }> {
    const name = LIST.roles;
    const status = await this.listStatus(name);
    if (status === 'ja') {
      await this.ensureRolesListPermissions(name);
      return { isNewlyCreated: false, status };
    }
    // Nur bei einem EINDEUTIGEN 404 wird angelegt (Review-Fund 11). „Unbekannt"
    // ist für jeden gewöhnlichen User der Normalfall — die Rollenliste ist für
    // ihn unlesbar (403) —, und dann darf nichts geschrieben werden.
    if (status !== 'nein') return { isNewlyCreated: false, status };

    if (!(await this.createList(name, 'Rollenverwaltung der AI Use Case Platform'))) {
      return { isNewlyCreated: false, status };
    }
    await this.feldText(name, 'UserName');
    // Die Auswahlwerte bleiben die des Altbestands (`Kurator` steht für den Use
    // Case Organizer) — `utils/rollen.ts` erklärt, warum.
    await this.feldAuswahl(name, 'Role', ['Admin', 'Kurator', 'User']);
    await this.feldText(name, 'AssignedBy');
    await this.feldDatum(name, 'AssignedDate');
    await this.ensureRolesListPermissions(name);
    return { isNewlyCreated: true, status };
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

      // Nur wer Listen verwalten darf, kappt die Vererbung. Diese Funktion läuft
      // bei JEDEM Start jeder Person, die die Rollenliste lesen kann; erbt sie
      // (der Rechte-Schritt bei der Installation ist gescheitert), löste sonst
      // jeder Nutzer aussichtslose Schreibanfragen aus — und ein Site-Owner, der
      // in der App nur „User" ist, vergäbe sich beim Start selbst Full Control auf
      // der Rollenliste, ohne dass es eine Rolle gibt (Review 29.09.2026).
      if (!(await this.kannRechteVerwalten())) return;

      // Ohne bekannte eigene Id gäbe es nach dem Kappen niemanden, dem man Full Control
      // nachweisbar zurückgeben könnte — dann wird gar nicht erst gekappt.
      const me = this.context.pageContext.legacyPageContext?.userId;
      if (!me) return;

      // Reihenfolge nach dem Sicherheits-Review (29.09.2026, Fund 4): Vor dem Kappen
      // steht fest, wer Owner ist; nach dem Kappen kommt die anlegende Person ZUERST
      // (mit Nachlesen), dann die Owners. Bis v1.3 waren das zwei rohe POSTs ohne Prüfung
      // von `ok` — schlug der Lesevorgang auf die Owners-Gruppe fehl oder scheiterten
      // beide POSTs, stand die Liste ohne jede Zuweisung da, alle Rollenlesungen
      // antworteten 403, und nur ein Site-Collection-Admin konnte helfen.
      let ownerId = 0;
      try {
        const owners = await this._sp.get(
          `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
          SPHttpClient.configurations.v1,
        );
        if (owners.ok) {
          const od = await owners.json();
          ownerId = Number(od.Id ?? od.d?.Id) || 0;
        }
      } catch { /* ohne Owners-Gruppe wird nur die anlegende Person eingetragen */ }

      const brk = await this._post(`${this.list(listName)}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`, {});
      if (!brk.ok) return;

      // Die anlegende Person braucht Full Control, sonst sperrt sie sich selbst aus.
      const ichOk = await this.grantVerified(this.list(listName), me, ROLE_DEF.full, 'ensureRolesListPermissions:me');
      if (ownerId) await this.grantVerified(this.list(listName), ownerId, ROLE_DEF.full, 'ensureRolesListPermissions:owners');
      if (!ichOk) {
        // Nicht gesetzt oder nicht nachlesbar: die Vererbung zurückholen, statt eine Liste
        // ohne Zugriff zu hinterlassen. Das Zurücksetzen darf nur, wer noch Rechte hat —
        // scheitert es, steht die Warnung in der Konsole, und der Audit meldet die Lücke.
        //
        // Bewusste Abwägung (Gegenprüfung 29.09.2026): Bei einem nur nicht LESBAREN Ergebnis
        // (429 beim Nachlesen) öffnet das die Liste kurz wieder für alle mit Schreibrecht auf der
        // Site. Die Alternative — geschlossen lassen — kann eine Liste hinterlassen, die niemand
        // mehr lesen kann (alle „User", keine Erstinstallation, nur ein Site-Collection-Admin hilft).
        // Das Öffnen ist sichtbar („Rechte prüfen": Rollenliste erbt) und heilt sich beim nächsten
        // Start einer Person mit „Manage Permissions" von selbst; die Aussperrung nicht.
        console.warn('[AIUC] Rollenliste: Full Control der anlegenden Person ließ sich nicht bestätigen — die Vererbung wird zurückgeholt.');
        await this._post(`${this.list(listName)}/resetroleinheritance`, {}).catch(() => undefined);
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
    this.lastRolesReadCode = '';
    type Zeile = { Id: number; Title: string; UserName: string; Role: string; AssignedBy: string; AssignedDate: string };
    try {
      // Auch hier ohne `$select` — aus demselben Grund wie bei den Use Cases.
      // Fehlt eine der Spalten, antwortet SharePoint mit HTTP 400, `getRoles`
      // liefert `null`, und die Person ist „User", obwohl sie Admin sein
      // sollte. Das war der Zustand auf dem Screenshot vom 10.09.2026.
      //
      // v1.3 (Fund 15): Mit Paging. `$top=500` liefert bei mehr Zeilen nur die
      // ersten 500 — wer in Zeile 501 als Admin steht, war für die App „User",
      // und der Audit meldete „geprüft, in Ordnung", ohne ihn gelesen zu haben.
      // Der `nextLink` wird verfolgt; ist die Liste größer als der Lesepfad
      // (5 000 Zeilen), ist das Ergebnis `null` mit Code — kein stiller
      // Teil-Erfolg (`leseZuweisungen` macht es ebenso).
      let url: string | null = `${this.list(LIST.roles)}/items?$top=500`;
      const alle: Zeile[] = [];
      let seiten = 0;
      while (url) {
        seiten++;
        const r: SPHttpClientResponse = await this._sp.get(
          url,
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
        ((d.value || []) as Zeile[]).forEach(z => alle.push(z));
        url = d['odata.nextLink'] || d['@odata.nextLink'] || null;
        if (url && (alle.length >= 5000 || seiten >= 10)) {
          this.lastRolesReadCode = 'zu-gross';
          console.warn(`[AIUC] Rollen lesen: mehr als ${alle.length} Zeilen — der Lesepfad ist gekappt, das Ergebnis gilt als NICHT lesbar.`);
          return null;
        }
      }
      return alle;
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

  /**
   * Die Rolle einer Zeile ändern.
   *
   * v1.3 (Fund 18): Mit `assignedBy` schreibt derselbe MERGE auch `AssignedBy`
   * und `AssignedDate` — die Zeile zeigt „Vergeben von … · Datum", und bei
   * einer Erhöhung stand dort sonst weiter die Erstvergabe. Wer wann jemanden
   * zum Admin gemacht hat, wäre nirgends zu finden.
   */
  public async updateRole(itemId: number, role: UserRole, assignedBy?: string): Promise<boolean> {
    try {
      const body: Record<string, unknown> = { 'Role': rolleFuerSpeicher(role) };
      if (assignedBy !== undefined) {
        body.AssignedBy = assignedBy;
        body.AssignedDate = new Date().toISOString();
      }
      const r = await this._mergeItem(LIST.roles, itemId, body);
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

  /**
   * Status der zuletzt gescheiterten `ensureuser`-Anfrage (0 = Ausnahme). Er
   * trennt „dieses Konto gibt es nicht" (400/404/500) von „SharePoint hat nicht
   * geantwortet" (0/429/503/504): Nur im ersten Fall ist eine Person, die auf
   * den Listen nirgends steht, belegt ohne Rechte — im zweiten wissen wir es nicht.
   */
  private lastResolveStatus = 0;

  private resolveUnklar(): boolean {
    const s = this.lastResolveStatus;
    return s === 0 || s === 429 || s === 503 || s === 504 || s === 408;
  }

  private async resolveUserId(userEmail: string): Promise<number | null> {
    // `ensureuser` legt das Konto an, falls die Person die Site noch nie
    // besucht hat — deshalb nicht `siteusers/getbyemail`, das dann 404 liefert
    // (DEX v31.84). v1.3: Zweiter Versuch im Claims-Format wie in DEX; manche
    // Konten löst SharePoint nur so auf.
    const namen = [userEmail, `i:0#.f|membership|${userEmail}`];
    this.lastResolveStatus = 0;
    for (let i = 0; i < namen.length; i++) {
      try {
        const r = await this._post(
          `${this.siteUrl}/_api/web/ensureuser`,
          { 'logonName': namen[i] },
        );
        if (!r.ok) { this.lastResolveStatus = r.status; continue; }
        const d = await r.json();
        const id = d.d?.Id ?? d.Id;
        if (typeof id === 'number') return id;
      } catch {
        this.lastResolveStatus = 0;
        // nächster Versuch
      }
    }
    return null;
  }

  /** Die Id des angemeldeten Kontos auf DIESER Site (0 = unbekannt — dann gilt nichts als „ich"). */
  private meineId(): number {
    const id = Number(this.context.pageContext.legacyPageContext?.userId);
    return isNaN(id) || id <= 0 ? 0 : id;
  }

  /**
   * Ist diese Adresse — in irgendeiner Schreibweise — die angemeldete Person?
   *
   * v1.3 (Review-Fund 4): Der Selbstschutz verglich bis dahin nur die
   * ADRESS-ZEICHENKETTE (`isCurrentUser`: `user.email` und die Adresse im
   * `loginName`). Steht dieselbe Person zusätzlich unter einem Alias in der
   * Rollenliste — genau der Fall, den „Rechte prüfen" als „andere
   * Schreibweise" meldet —, lösen beide Adressen zum SELBEN Konto auf, und der
   * Entzug nahm dem Admin die eigene Zuweisung (Full Control auf der
   * Rollenliste); beim nächsten Laden: 403, ausgesperrt. Deshalb zählt zuletzt
   * das AUFGELÖSTE Konto. `null` = nicht auflösbar (dann weiß man es nicht; die
   * Entzugspfade prüfen die Id noch einmal selbst).
   */
  public async istMeinKonto(email: string): Promise<boolean | null> {
    if (isCurrentUser(this.context, email)) return true;
    const id = await this.resolveUserId(email);
    if (!id) return null;
    return id === this.meineId();
  }

  /**
   * Recht setzen UND nachlesen, mit Wiederholung.
   *
   * Ein `addroleassignment`-POST allein ist keine Vergabe: Ein 429 kostet das
   * Recht still, und es faellt erst auf, wenn jemand seine Kachel vermisst.
   * Full Control deckt jedes niedrigere Recht ab, deshalb zaehlt es beim
   * Nachlesen mit — das bleibt so.
   *
   * `exakt`: Full Control zählt NICHT als Deckung, gelesen wird genau diese
   * Stufe. Das braucht nur der Rückbau (`reduziereListe`): Wer von Full Control
   * auf Lesen zurückgestuft wird, muss Lesen wirklich als EIGENE Stufe tragen,
   * sonst bliebe nach dem Entzug von Full Control nichts übrig.
   */
  private async grantVerified(base: string, userId: number, roleDefId: number, label: string, exakt = false): Promise<boolean> {
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
        return ids.indexOf(roleDefId) >= 0 || (!exakt && ids.indexOf(ROLE_DEF.full) >= 0);
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

  private sollFuer(key: RechteListe): RechteSoll {
    return RECHTE_SOLL.filter(x => x.key === key)[0];
  }

  /** Der Code für „diese Stufe auf dieser Liste konnte nicht gesetzt werden". */
  private codeFuer(key: RechteListe, stufe: number): RechteCode {
    const name = stufe === ROLE_DEF.full ? 'full' : stufe === ROLE_DEF.edit ? 'edit' : stufe === ROLE_DEF.contribute ? 'contribute' : 'read';
    return `${key}:${name}` as RechteCode;
  }

  /** Rechte auf den angegebenen Listen für eine schon aufgelöste Person — je Recht mit Nachlesen. */
  private async grantListRights(id: number, rolle: UserRole, keys: RechteListe[]): Promise<RechteCode[]> {
    const fehlend: RechteCode[] = [];
    for (let i = 0; i < RECHTE_SOLL.length; i++) {
      const s = RECHTE_SOLL[i];
      if (keys.indexOf(s.key) < 0) continue;
      // Die Rollenliste hat ihre eigenen Rechte seit dem Anlegen (Owners, sonst
      // niemand); die anderen zwei erben noch — sie brauchen erst eigene.
      if (s.key !== 'roles') await this.ensureListeEigeneRechte(s.liste);
      const need = rolle === 'Admin' ? s.admin : s.organizer;
      if (!(await this.grantVerified(this.list(s.liste), id, need, `grant/${s.key}`))) fehlend.push(this.codeFuer(s.key, need));
    }
    return fehlend;
  }

  /**
   * Die vollständige Vergabe zur Rolle: Rollenliste, Use-Case-Liste, Protokoll —
   * je Recht nachgelesen. Rückgabe: was NICHT gesetzt werden konnte (Codes).
   * `User` bekommt nichts.
   */
  public async grantRoleRights(userEmail: string, rolle: UserRole): Promise<RechteCode[]> {
    if (rolle === 'User') return [];
    const id = await this.resolveUserId(userEmail);
    if (!id) return ['konto'];
    return this.grantListRights(id, rolle, ['roles', 'useCases', 'log']);
  }

  /** Leseversuch der Zuweisungen mit EINER Wiederholung nach kurzer Pause — außer bei 401/403/404, die nichts ändert. */
  private async leseZuweisungenGewiss(listName: string): Promise<ListenZuweisungen | null> {
    const z = await this.leseZuweisungen(listName);
    if (z) return z;
    const f = this.lastRightsReadFehler;
    if (f && f.code === 'http' && (f.status === 401 || f.status === 403 || f.status === 404)) return null;
    await new Promise<void>(res => setTimeout(res, 1500));
    return this.leseZuweisungen(listName);
  }

  /** Die Konten, auf die `em` (klein, getrimmt) in dieser Liste zeigt: das aufgelöste Konto plus jedes, das dort unter dieser Adresse steht. */
  private kontenAufListe(z: ListenZuweisungen, em: string, id: number | null): number[] {
    const out: number[] = [];
    if (id) out.push(id);
    z.adressenZuId.forEach((adressen, pid) => {
      if (adressen.indexOf(em) >= 0 && out.indexOf(pid) < 0) out.push(pid);
    });
    return out;
  }

  /** Stufen ohne „Limited Access" — das hängt SharePoint von selbst an und ist kein Recht auf der Liste. */
  private echteStufen(ids: number[] | undefined): number[] {
    return (ids || []).filter(x => x !== LIMITED_ACCESS);
  }

  /**
   * Eine direkte Zuweisung auf EINER Liste ganz entfernen und NACHLESEN.
   *
   * Erfolg heißt „nachgelesen steht das Konto nicht mehr auf der Liste" — nicht
   * „der Aufruf war 200". Früher las ein `getbyprincipalid/roledefinitionbindings`
   * nach, und ein 404 ODER 500 galt als „Zuweisung weg" (Fund 16): 500 ist auch
   * das, was SharePoint bei einer echten Störung liefert, ein DELETE, der nicht
   * durchging, plus ein 500 beim Kontrolllesen ergab „Entzug bestätigt". Jetzt
   * belegt die GELESENE Zuweisungsliste der Liste das Fehlen (bei einem
   * Lesefehler einmal nach kurzer Pause erneut, dann `lesefehler`).
   *
   * Eine Liste, die noch von der Site erbt, trägt nichts, was die App vergeben
   * hätte — dort gibt es nichts zu entziehen.
   */
  private async entzieheListe(key: RechteListe, email: string, id: number | null): Promise<EntzugsCode> {
    const s = this.sollFuer(key);
    const base = this.list(s.liste);
    const em = (email || '').trim().toLowerCase();
    const z = await this.leseZuweisungenGewiss(s.liste);
    if (!z) return 'lesefehler';
    if (!z.eigene) return 'ok';
    const konten = this.kontenAufListe(z, em, id);
    // Nichts unter dieser Adresse und nicht aufgelöst: Ist das Auflösen nur
    // an einer Drosselung gescheitert, ist es NICHT belegt (Alias-Konten
    // fielen sonst durch); war das Konto nur unbekannt (ausgeschieden), gibt
    // es nichts zu entziehen — sonst ließe sich die Zeile nie mehr entfernen.
    if (konten.length === 0) return id === null && this.resolveUnklar() ? 'konto' : 'ok';
    const meine = this.meineId();
    if (meine && konten.indexOf(meine) >= 0) return 'selbst';
    // Nur löschen, was dort wirklich steht: Ein DELETE auf eine nicht vorhandene
    // Zuweisung beantwortet SharePoint mit 404/500 — eine rote Zeile in der
    // Konsole für nichts (und bei jeder Person ohne Rechte dreimal).
    const vorhanden = konten.filter(k => z.nachId.has(k));
    if (vorhanden.length === 0) return 'ok';

    for (let i = 0; i < vorhanden.length; i++) {
      await this._delete(`${base}/roleassignments/getbyprincipalid(${vorhanden[i]})`).catch(() => undefined);
    }
    const danach = await this.leseZuweisungenGewiss(s.liste);
    if (!danach) return 'lesefehler';
    const rest = vorhanden.filter(k => this.echteStufen(danach.nachId.get(k)).length > 0);
    if (rest.length === 0) return 'ok';
    console.warn(`[AIUC] revoke/${key}: ${email} hat auf ${base} weiterhin Rechte (Konto ${rest.join(', ')}).`);
    return 'offen';
  }

  /**
   * Auf einer Liste nur die ÜBERZÄHLIGEN Stufen entziehen (Review-Fund 9).
   *
   * `behalten` = die Stufe, die bleiben muss (Organizer: `s.organizer`), oder
   * `null` = nichts soll bleiben. Reihenfolge, damit ein Fehler Richtung WENIGER
   * Zugriff scheitert und nie zu „nichts mehr": erst die benötigte Stufe
   * sicherstellen (STRIKT als eigene Stufe — Full Control zählt hier nicht),
   * dann mit `removeroleassignment(roledefid = überzählige Stufe)` genau die
   * überzähligen entfernen, dann nachlesen. Früher wurde die GANZE Zuweisung
   * gelöscht und Lesen danach neu vergeben; scheiterte die Rückvergabe, stand
   * dort „nicht entzogen", obwohl der Entzug gelungen war und die Person auf
   * der Rollenliste gar nichts mehr hatte.
   */
  private async reduziereListe(key: RechteListe, email: string, id: number | null, behalten: number | null): Promise<'ok' | 'lesenFehlt' | 'lesefehler' | 'offen' | 'selbst'> {
    const s = this.sollFuer(key);
    const base = this.list(s.liste);
    const em = (email || '').trim().toLowerCase();
    const z = await this.leseZuweisungenGewiss(s.liste);
    if (!z) return 'lesefehler';
    if (!z.eigene) return 'ok';
    const konten = this.kontenAufListe(z, em, id);
    const meine = this.meineId();
    if (meine && konten.indexOf(meine) >= 0) return 'selbst';
    const ueber = (x: number): boolean =>
      x !== LIMITED_ACCESS && (behalten === null || STUFEN.indexOf(x) > STUFEN.indexOf(behalten));

    let ergebnis: 'ok' | 'lesenFehlt' | 'lesefehler' | 'offen' = 'ok';
    for (let i = 0; i < konten.length; i++) {
      const k = konten[i];
      const stufen = z.nachId.get(k) || [];
      const ueberzaehlig = stufen.filter(ueber);
      if (ueberzaehlig.length === 0) continue;
      if (behalten !== null && stufen.indexOf(behalten) < 0) {
        if (!(await this.grantVerified(base, k, behalten, `reduziere/${key}`, true))) { ergebnis = 'offen'; continue; }
      }
      for (let j = 0; j < ueberzaehlig.length; j++) {
        await this._post(`${base}/roleassignments/removeroleassignment(principalid=${k}, roledefid=${ueberzaehlig[j]})`, {}).catch(() => undefined);
      }
      const danach = await this.leseZuweisungenGewiss(s.liste);
      if (!danach) { ergebnis = 'lesefehler'; continue; }
      const jetzt = danach.nachId.get(k) || [];
      if (jetzt.filter(ueber).length > 0) { ergebnis = 'offen'; continue; }
      if (behalten !== null && jetzt.indexOf(behalten) < 0 && ergebnis === 'ok') ergebnis = 'lesenFehlt';
    }
    return ergebnis;
  }

  /** Das schlimmere von zwei Entzugs-Ergebnissen (`ok` < `konto` < `lesefehler` < `offen`). */
  private schlimmer(a: EntzugsCode, b: EntzugsCode): EntzugsCode {
    const rang = ['ok', 'konto', 'lesefehler', 'offen'];
    return rang.indexOf(b) > rang.indexOf(a) ? b : a;
  }

  /**
   * Alle direkten Rechte einer Person entziehen.
   *
   * Selbstschutz über das AUFGELÖSTE Konto (Fund 4): Die angemeldete Person
   * entzieht sich nicht selbst — auch nicht über einen Alias in der Rollenliste
   * (DEX v30.67: sonst kaeme ein Admin, der die Rollenverwaltung an sich
   * ausprobiert, nicht mehr an die Liste). Das Ergebnis ist `selbst`, NICHT
   * `ok`: Bis v1.3 antwortete der Selbstschutz mit `true` („erledigt"), obwohl
   * nichts entzogen wurde.
   */
  public async revokeAllAccess(userEmail: string): Promise<EntzugsCode> {
    if (isCurrentUser(this.context, userEmail)) {
      console.warn(`[AIUC] Rechte von ${userEmail} werden NICHT entzogen — das ist die angemeldete Person (Selbstschutz).`);
      return 'selbst';
    }
    const id = await this.resolveUserId(userEmail);
    if (id && id === this.meineId()) {
      console.warn(`[AIUC] Rechte von ${userEmail} werden NICHT entzogen — die Adresse löst zum eigenen Konto auf (Selbstschutz).`);
      return 'selbst';
    }
    let ergebnis: EntzugsCode = 'ok';
    for (let i = 0; i < RECHTE_SOLL.length; i++) {
      const r = await this.entzieheListe(RECHTE_SOLL[i].key, userEmail, id);
      if (r === 'selbst') return 'selbst';
      ergebnis = this.schlimmer(ergebnis, r);
    }
    return ergebnis;
  }

  /**
   * v1.3 (Fund 2/9): Rechte auf das Maß einer niedrigeren Rolle zurückführen,
   * BEVOR die Zeile geändert wird — Admin → Organizer: auf allen drei Listen
   * die Organizer-Stufe eigenständig sicherstellen und alles darüber entziehen;
   * nach User: alles entziehen. Ist das Ergebnis nicht `ok`, bleibt die Zeile
   * unverändert („nichts wurde geändert"): im Zweifel scheitert es Richtung
   * WENIGER Zugriff, nie Richtung Rechte ohne Zeile.
   */
  public async reduziereRechte(userEmail: string, rolle: 'Organizer' | 'User'): Promise<EntzugsCode> {
    if (rolle === 'User') return this.revokeAllAccess(userEmail);
    if (isCurrentUser(this.context, userEmail)) return 'selbst';
    const id = await this.resolveUserId(userEmail);
    if (!id) return 'konto';
    if (id === this.meineId()) return 'selbst';
    let ergebnis: EntzugsCode = 'ok';
    for (let i = 0; i < RECHTE_SOLL.length; i++) {
      const s = RECHTE_SOLL[i];
      const r = await this.reduziereListe(s.key, userEmail, id, s.organizer);
      if (r === 'selbst') return 'selbst';
      ergebnis = this.schlimmer(ergebnis, r === 'lesenFehlt' || r === 'offen' ? 'offen' : r);
    }
    return ergebnis;
  }

  // ===================================================================
  // Personensuche und Rechte-Prüfung (v1.3)
  // ===================================================================

  /** HTTP-Status der letzten Personensuche (0 = Ausnahme oder nie gesucht). */
  public lastPersonSearchStatus = 0;
  /**
   * Warum das letzte Rechte-Lesen scheiterte — als CODE (Liste, Art, Status),
   * nicht als Satz: Die Oberfläche formuliert zweisprachig mit `t()` (Fund 19).
   */
  public lastRightsReadFehler: LeseFehler | null = null;
  /**
   * Zähler der Personensuche (Fund 20). Jede neue Suche überholt die davor; deren
   * Antworten werden verworfen und dürfen `lastPersonSearchStatus` nicht mehr
   * überschreiben.
   */
  private _sucheNr = 0;

  /** Laufende Suchen für überholt erklären — z. B. beim Verlassen der Seite. */
  public sucheAbbrechen(): void {
    this._sucheNr++;
  }

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
    const nr = ++this._sucheNr;
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
        // Nicht wiederholend (Fund 20): Das ist ein LESEN, das nur als POST
        // geschickt wird. Über `_post` hätte jede überholte Suche ihre bis zu
        // zwei Wiederholungen mit Retry-After (bis 45 s) zu Ende gewartet.
        const r = await this._postEinmal(
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

    // Überholt: Eine neuere Suche läuft. Diese Antworten gehören niemandem mehr
    // und dürfen den Status der neueren nicht überschreiben.
    if (nr !== this._sucheNr) return { treffer: [], ausgeblendet: 0 };

    const gescheitert = antworten.filter(a => !a.ok);
    if (gescheitert.length === antworten.length) {
      this.lastPersonSearchStatus = gescheitert[0].status;
      console.warn(`[AIUC] Personensuche: alle ${antworten.length} Anfragen gescheitert (HTTP ${gescheitert[0].status}).`);
      return null;
    }
    this.lastPersonSearchStatus = 200;

    const istDeloitte = (mail: string): boolean => {
      const at = mail.lastIndexOf('@');
      return at >= 0 && istDeloitteDomain(mail.slice(at + 1));
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
   * **Nebenwirkung, die man kennen muss:** Was danach am WEB vergeben wird — eine
   * neue AD-Gruppe, „Everyone except external users", eine neu geteilte Site —,
   * erreicht diese Liste NICHT mehr. Neue Leser müssen dann auf der Liste selbst
   * eingetragen werden. Das ist gewollt (eine Liste nimmt sonst keine Zuweisung
   * an einzelne Personen an); die Rechte-Matrix (Kategorie „SharePoint") und die
   * README sagen es ebenfalls.
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
   * `null` = nicht lesbar (Grund in `lastRightsReadFehler`). Auch eine
   * abgeschnittene Antwort (mehr als 20 Seiten) ist `null`: Ein gekappter
   * Lauf darf nie „steht nicht mehr drin" bedeuten.
   *
   * Personen (PrincipalType 1) landen in `nachAdresse`/`nachId`; SharePoint-
   * Gruppen (PrincipalType 8) in `gruppen` — der Audit zählt ihre Stufen als
   * Deckung („über Gruppe gedeckt"), statt „Rechte fehlen" zu melden.
   */
  private async leseZuweisungen(listName: string): Promise<ListenZuweisungen | null> {
    const base = this.list(listName);
    const kopf = { headers: { 'Accept': 'application/json;odata=nometadata' } };
    const liste: RechteListe | 'rollen' = RECHTE_SOLL.filter(x => x.liste === listName)[0]?.key ?? 'rollen';
    this.lastRightsReadFehler = null;
    try {
      const flag = await this._sp.get(`${base}?$select=HasUniqueRoleAssignments`, SPHttpClient.configurations.v1, kopf);
      if (!flag.ok) { this.lastRightsReadFehler = { code: 'http', liste, status: flag.status }; return null; }
      const fd = await flag.json();
      const out: ListenZuweisungen = {
        eigene: !!(fd.HasUniqueRoleAssignments ?? fd.d?.HasUniqueRoleAssignments),
        nachAdresse: new Map<string, number[]>(),
        nachId: new Map<number, number[]>(),
        adresseZuId: new Map<number, string>(),
        adressenZuId: new Map<number, string[]>(),
        gruppen: [],
      };

      let url: string | null = `${base}/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Email,Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Id&$top=5000`;
      let seiten = 0;
      while (url && seiten < 20) {
        seiten++;
        const r: SPHttpClientResponse = await this._sp.get(url, SPHttpClient.configurations.v1, kopf);
        if (!r.ok) { this.lastRightsReadFehler = { code: 'http', liste, status: r.status }; return null; }
        const d = await r.json();
        const zeilen = (d.value || d.d?.results || []) as SpZuweisung[];
        zeilen.forEach(z => {
          const m = z.Member || {};
          const bindings = Array.isArray(z.RoleDefinitionBindings) ? z.RoleDefinitionBindings : (z.RoleDefinitionBindings?.results || []);
          const ids = bindings.map(b => Number(b.Id));
          const pid = Number(z.PrincipalId || 0);
          // SharePoint-Gruppe (Owners, Members …): trägt keine Person, deckt aber
          // Rechte ihrer Mitglieder ab.
          if (m.PrincipalType === 8) {
            if (pid > 0) out.gruppen.push({ id: pid, stufen: ids });
            return;
          }
          // Andere Nicht-Personen (Sicherheitsgruppen) nur ausschließen, wenn
          // SharePoint es ausdrücklich sagt. Fehlt das Feld, zählt der Eintrag:
          // sonst sähe jede Person „ohne Recht" aus.
          if (typeof m.PrincipalType === 'number' && m.PrincipalType !== 1) return;
          const adressen: string[] = [];
          if (m.Email) adressen.push(String(m.Email).toLowerCase().trim());
          const lm = String(m.LoginName || '').toLowerCase().match(/[^|]+@[^|\s]+$/);
          if (lm && adressen.indexOf(lm[0].trim()) < 0) adressen.push(lm[0].trim());
          adressen.forEach(a => out.nachAdresse.set(a, vereint(out.nachAdresse.get(a), ids)));
          if (pid > 0) {
            out.nachId.set(pid, vereint(out.nachId.get(pid), ids));
            out.adressenZuId.set(pid, vereint2(out.adressenZuId.get(pid), adressen));
            if (adressen[0] && !out.adresseZuId.has(pid)) out.adresseZuId.set(pid, adressen[0]);
          }
        });
        url = d['odata.nextLink'] || d['@odata.nextLink'] || (d.d && d.d.__next) || null;
      }
      if (url) { this.lastRightsReadFehler = { code: 'zu-gross', liste, status: 0 }; return null; }
      return out;
    } catch (e) {
      this.lastRightsReadFehler = { code: 'ausnahme', liste, status: 0 };
      console.warn(`[AIUC] Zuweisungen von ${listName} nicht lesbar:`, e);
      return null;
    }
  }

  /** Die Mitglieder einer SharePoint-Gruppe der Site (Id und Adressen) — `null` = nicht lesbar. */
  private async leseGruppenMitglieder(gruppenId: number): Promise<Array<{ id: number; adressen: string[] }> | null> {
    try {
      let url: string | null = `${this.siteUrl}/_api/web/sitegroups/getbyid(${gruppenId})/users?$select=Id,Email,LoginName,PrincipalType&$top=5000`;
      const out: Array<{ id: number; adressen: string[] }> = [];
      let seiten = 0;
      while (url && seiten < 5) {
        seiten++;
        const r: SPHttpClientResponse = await this._sp.get(url, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
        if (!r.ok) return null;
        const d = await r.json();
        ((d.value || d.d?.results || []) as Array<{ Id?: number; Email?: string; LoginName?: string; PrincipalType?: number }>).forEach(u => {
          if (typeof u.PrincipalType === 'number' && u.PrincipalType !== 1) return;
          const adressen: string[] = [];
          if (u.Email) adressen.push(String(u.Email).toLowerCase().trim());
          const lm = String(u.LoginName || '').toLowerCase().match(/[^|]+@[^|\s]+$/);
          if (lm && adressen.indexOf(lm[0].trim()) < 0) adressen.push(lm[0].trim());
          out.push({ id: Number(u.Id || 0), adressen });
        });
        url = d['odata.nextLink'] || d['@odata.nextLink'] || null;
      }
      // Eine gekappte Mitgliederliste sagt nicht „nicht Mitglied".
      return url ? null : out;
    } catch {
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
   * Adresse etwas fehlt — oder wenn auf den Listen Konten stehen, die zu keiner
   * Zeile passen (dann kann ein Alias darunter stecken). So kostet der Lauf bei
   * 130 Zeilen ein paar Aufrufe statt Hunderter — die Drosselung ist das Risiko
   * dieser Seite.
   *
   * Vier Befunde, getrennt gemeldet:
   *  - `luecken`: Ein Recht fehlt, das die Rolle braucht.
   *  - `ueberGruppe`: Das Recht steht nicht direkt da, sondern über eine
   *    SharePoint-Gruppe (Owners/Members): gedeckt, keine Lücke (Fund 14).
   *  - `aliase`: Die Rechte sind da, aber unter einer anderen Schreibweise der
   *    Adresse (SMTP gegen UPN/Alias) — kein Fehler der Rechte, aber die Zeile
   *    in der Rollenliste stimmt nicht mit dem Konto überein. DEX v31.67: Ohne
   *    diese Trennung stand eine Person bei JEDEM Lauf als „Lücke" da und
   *    wurde jedes Mal „nachgesetzt".
   *  - `ueberschuss`: Mehr, als die Rolle vorsieht — ein Organizer mit mehr
   *    als der Organizer-Stufe (Rollenliste: mehr als Lesen; Use Cases: mehr als
   *    Bearbeiten; Protokoll: mehr als Beitragen — meist der Rest einer früheren
   *    Admin-Rolle), oder eine Person mit Rolle User, die noch direkte Rechte
   *    trägt. Nur auf Listen, die eigene Berechtigungen haben: Auf einer
   *    vererbenden Liste stammen die Einträge von der Site und gehören nicht der
   *    App. Auch hier wird über das AUFGELÖSTE Konto gegen Alias-Zuweisungen
   *    geprüft, nicht nur über die Adresse (Fund 14).
   *
   * `null` = die Zuweisungen waren nicht lesbar; `lastRightsReadFehler` sagt,
   * welche Liste und warum. Das ist KEIN „alles in Ordnung".
   */
  public async auditRoleRights(
    zeilen: Array<{ email: string; name: string; rolle: UserRole }>,
    onProgress?: (done: number, total: number) => void,
  ): Promise<RechteBericht | null> {
    this.lastRightsReadFehler = null;
    const zuw: Record<string, ListenZuweisungen> = {};
    for (let i = 0; i < RECHTE_SOLL.length; i++) {
      const z = await this.leseZuweisungen(RECHTE_SOLL[i].liste);
      if (!z) return null;
      zuw[RECHTE_SOLL[i].key] = z;
    }

    const pause = (ms: number): Promise<void> => new Promise(res => setTimeout(res, ms));
    const norm = (a: string): string => (a || '').toLowerCase().trim();
    // Die Adresse ist der Schlüssel, und in der Rollenliste kann sie zweimal
    // stehen; es gilt die erste Zeile (`RoleContext` nimmt ebenfalls die erste).
    const gesehen: string[] = [];
    const relevant = zeilen.filter(z => {
      const em = norm(z.email);
      if (!em || gesehen.indexOf(em) >= 0) return false;
      gesehen.push(em);
      return true;
    });
    const bericht: RechteBericht = {
      geprueft: 0, ok: 0, okAdressen: [], luecken: [], ueberschuss: [], aliase: [], ueberGruppe: [], eigeneKonten: [],
      ohneAdresse: zeilen.filter(z => !(z.email || '').trim()).length,
    };

    // Konten, die auf den eigenen Listen etwas tragen, aber unter keiner Adresse
    // einer Zeile stehen: Nur dort kann ein Alias stecken. Gibt es keine, spart
    // sich der Lauf das Auflösen der Zeilen, die schon ohne Befund sind.
    const zeilenAdressen = zeilen.map(z => norm(z.email)).filter(Boolean);
    const fremde: number[] = [];
    RECHTE_SOLL.forEach(s => {
      const z = zuw[s.key];
      if (!z.eigene) return;
      z.nachId.forEach((stufen, pid) => {
        if (this.echteStufen(stufen).length === 0 || fremde.indexOf(pid) >= 0) return;
        const adr = z.adressenZuId.get(pid) || [];
        // Ohne jede Adresse (App-Konten, Systemkonten) kann es kein Alias einer
        // Person sein — und ein solches Konto würde sonst das Auflösen JEDER
        // Zeile auslösen.
        if (adr.length === 0 || adr.some(a => zeilenAdressen.indexOf(a) >= 0)) return;
        fremde.push(pid);
      });
    });

    const kontoCache: Record<string, number | null> = {};
    const kontoVon = async (em: string, roh: string): Promise<number | null> => {
      if (Object.prototype.hasOwnProperty.call(kontoCache, em)) return kontoCache[em];
      const id = await this.resolveUserId(roh);
      kontoCache[em] = id;
      if (id && id === this.meineId() && !isCurrentUser(this.context, roh)) bericht.eigeneKonten.push(em);
      await pause(150);
      return id;
    };
    const gruppenCache: Record<number, Array<{ id: number; adressen: string[] }> | null> = {};
    const mitglieder = async (gid: number): Promise<Array<{ id: number; adressen: string[] }> | null> => {
      if (!Object.prototype.hasOwnProperty.call(gruppenCache, gid)) gruppenCache[gid] = await this.leseGruppenMitglieder(gid);
      return gruppenCache[gid];
    };
    /** Deckt eine SharePoint-Gruppe der Liste dieses Recht — und ist die Person darin Mitglied? */
    const ueberGruppeGedeckt = async (key: RechteListe, need: number, em: string, id: number | null): Promise<boolean> => {
      const gruppen = zuw[key].gruppen;
      for (let g = 0; g < gruppen.length; g++) {
        if (!deckt(gruppen[g].stufen, need)) continue;
        const m = await mitglieder(gruppen[g].id);
        if (m && m.some(u => (id !== null && u.id === id) || u.adressen.indexOf(em) >= 0)) return true;
      }
      return false;
    };

    for (let i = 0; i < relevant.length; i++) {
      const z = relevant[i];
      const em = norm(z.email);
      bericht.geprueft++;
      let id: number | null = null;
      let versucht = false;
      const aufloesen = async (): Promise<number | null> => {
        if (!versucht) { versucht = true; id = await kontoVon(em, z.email); }
        return id;
      };
      const stufen = (key: RechteListe, konto: number | null): number[] =>
        vereint(zuw[key].nachAdresse.get(em), konto ? zuw[key].nachId.get(konto) : undefined);
      let auffaellig = false;

      if (z.rolle === 'User') {
        // Wer nur „User" ist, braucht nichts — und trägt auch nichts mehr.
        const uebrigMit = (konto: number | null): RechteListe[] => RECHTE_SOLL
          .filter(s => zuw[s.key].eigene && this.echteStufen(stufen(s.key, konto)).length > 0)
          .map(s => s.key);
        let uebrig = uebrigMit(null);
        if (uebrig.length === 0 && fremde.length > 0) {
          const k = await aufloesen();
          if (k) uebrig = uebrigMit(k);
        }
        if (uebrig.length > 0) {
          bericht.ueberschuss.push({ email: z.email, name: z.name, rolle: z.rolle, listen: uebrig });
          auffaellig = true;
        }
      } else {
        const braucht = (s: RechteSoll): number => (z.rolle === 'Admin' ? s.admin : s.organizer);
        let fehlt = RECHTE_SOLL.filter(s => !deckt(stufen(s.key, null), braucht(s)));
        let aufloesbar = true;
        if (fehlt.length > 0) {
          const k = await aufloesen();
          if (k) {
            const vorher = fehlt.length;
            const rest = RECHTE_SOLL.filter(s => !deckt(stufen(s.key, k), braucht(s)));
            if (rest.length < vorher) {
              const sp = zuw.roles.adresseZuId.get(k) || zuw.useCases.adresseZuId.get(k) || zuw.log.adresseZuId.get(k) || '';
              if (sp && sp !== em) bericht.aliase.push({ name: z.name, email: z.email, spEmail: sp });
            }
            fehlt = rest;
          } else {
            aufloesbar = false;
          }
        }
        // Was direkt fehlt, kann eine Gruppe der Liste decken (Owners, Members):
        // dann ist es keine Lücke, und „Nachsetzen" würde nur ein redundantes
        // Direktrecht danebenlegen (Fund 14).
        if (fehlt.length > 0) {
          const ueber: RechteListe[] = [];
          const bleibt: RechteSoll[] = [];
          for (let f = 0; f < fehlt.length; f++) {
            if (await ueberGruppeGedeckt(fehlt[f].key, braucht(fehlt[f]), em, id)) ueber.push(fehlt[f].key);
            else bleibt.push(fehlt[f]);
          }
          if (ueber.length > 0) bericht.ueberGruppe.push({ email: z.email, name: z.name, rolle: z.rolle, listen: ueber });
          fehlt = bleibt;
        }
        if (fehlt.length > 0) {
          bericht.luecken.push({ email: z.email, name: z.name, rolle: z.rolle, fehlt: fehlt.map(s => s.key), kontoAufloesbar: aufloesbar });
          auffaellig = true;
        }
        // Ein Organizer soll die Rollenliste LESEN, nicht bearbeiten — und auf den
        // anderen zwei Listen nicht mehr als Beitragen (seit v1.3; früher Edit). Mehr ist
        // meist der Rest einer früheren Admin-Rolle — und damit der Weg, sich
        // selbst wieder hochzustufen.
        if (z.rolle === 'Organizer') {
          const k = fremde.length > 0 ? await aufloesen() : id;
          const ueberListen = RECHTE_SOLL
            .filter(s => zuw[s.key].eigene && stufen(s.key, k).some(x => STUFEN.indexOf(x) > STUFEN.indexOf(s.organizer)))
            .map(s => s.key);
          if (ueberListen.length > 0) {
            bericht.ueberschuss.push({ email: z.email, name: z.name, rolle: z.rolle, listen: ueberListen });
            auffaellig = true;
          }
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
   * erbt (Use-Case-Liste, Protokoll: mit Kopie — Nebenwirkung siehe
   * `ensureListeEigeneRechte`; Rollenliste: ohne — sie ist ausdrücklich nur
   * für die Owners gedacht).
   *
   * Nur additiv. Was zu VIEL da ist, fasst dieser Schritt nicht an. Ob die
   * Rolle noch stimmt, für die nachgesetzt wird, prüft der Aufrufer
   * (`RoleContext.repairRoleRights`) — hier steht sie im Auftrag.
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
   * die überzähligen STUFEN — und lässt, was ihre Rolle vorsieht, durchgehend
   * bestehen (ein Organizer behält Lesen auf der Rollenliste, Bearbeiten auf
   * den Use Cases, Beitragen im Protokoll). Das benötigte Recht wird ZUERST
   * sichergestellt, dann werden nur die überzähligen Stufen entfernt und
   * nachgelesen (`reduziereListe`).
   *
   * Drei getrennte Zustände je Liste (Fund 9), damit die Meldung stimmt:
   *  - `entzogen`   — überzählige Rechte weg, das Vorgesehene steht noch.
   *  - `lesenFehlt` — überzählige Rechte weg, aber das Vorgesehene fehlt jetzt
   *    (nachgelesen): Der Entzug ist NICHT der Fehler, der Rest muss nachgesetzt werden.
   *  - `offen`      — der Entzug ist nicht belegt (Rechte noch da, nicht lesbar,
   *    Konto nicht auflösbar): nichts als erledigt behaupten.
   *
   * Selbstschutz über das aufgelöste Konto: für die angemeldete Person wird
   * nichts entzogen (`selbst`, alle Listen `offen`).
   */
  public async entzieheUeberschuss(email: string, rolle: UserRole, listen: RechteListe[]): Promise<UeberschussErgebnis> {
    const ergebnis: UeberschussErgebnis = { entzogen: [], lesenFehlt: [], offen: [], selbst: false, lesefehler: false };
    if (isCurrentUser(this.context, email)) {
      console.warn(`[AIUC] Überschüssige Rechte von ${email} werden NICHT entzogen — das ist die angemeldete Person (Selbstschutz).`);
      return { ...ergebnis, offen: listen.slice(), selbst: true };
    }
    const id = await this.resolveUserId(email);
    if (id && id === this.meineId()) {
      console.warn(`[AIUC] Überschüssige Rechte von ${email} werden NICHT entzogen — die Adresse löst zum eigenen Konto auf (Selbstschutz).`);
      return { ...ergebnis, offen: listen.slice(), selbst: true };
    }
    for (let i = 0; i < listen.length; i++) {
      const s = this.sollFuer(listen[i]);
      if (!s) continue;
      // Ein Admin trägt nie „Überschuss"; ein Organizer behält seine Stufe, ein
      // User nichts.
      const behalten = rolle === 'User' ? null : (rolle === 'Admin' ? s.admin : s.organizer);
      const r = await this.reduziereListe(s.key, email, id, behalten);
      if (r === 'ok') ergebnis.entzogen.push(s.key);
      else if (r === 'lesenFehlt') ergebnis.lesenFehlt.push(s.key);
      else if (r === 'selbst') return { ...ergebnis, offen: listen.slice(), entzogen: [], lesenFehlt: [], selbst: true };
      else {
        ergebnis.offen.push(s.key);
        if (r === 'lesefehler') ergebnis.lesefehler = true;
      }
    }
    return ergebnis;
  }

  // ===================================================================
  // Use Cases
  // ===================================================================

  public ensureUseCaseList(): Promise<{ isNewlyCreated: boolean; status: ListenStatus }> {
    return this.einmalig('useCases', () => this.ensureUseCaseListNeu());
  }

  private async ensureUseCaseListNeu(): Promise<{ isNewlyCreated: boolean; status: ListenStatus }> {
    const name = LIST.useCases;
    const status = await this.listStatus(name);
    // Nur bei 404 anlegen; bei „unbekannt" (403, 429, 5xx) nichts schreiben
    // und nichts nachziehen — ein Lesefehler ist keine fehlende Liste (Fund 11).
    if (status === 'unbekannt') return { isNewlyCreated: false, status };
    const existed = status === 'ja';
    if (!existed && !(await this.createList(name, 'Die Agent-Demos der AI Use Case Platform'))) {
      return { isNewlyCreated: false, status };
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

    return { isNewlyCreated: !existed, status };
  }

  public ensureLogList(): Promise<ListenStatus> {
    return this.einmalig('log', () => this.ensureLogListNeu());
  }

  private async ensureLogListNeu(): Promise<ListenStatus> {
    const name = LIST.log;
    const status = await this.listStatus(name);
    if (status === 'unbekannt') return status;
    if (status === 'nein' && !(await this.createList(name, 'Änderungsprotokoll der AI Use Case Platform'))) return status;
    await this.feldZahl(name, 'UseCaseId');
    await this.feldText(name, 'Aktion');
    await this.feldNote(name, 'Detail');
    await this.feldText(name, 'Wer');
    return status;
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
      // Mit Paging, wie `getRoles`: `$top=500` allein lieferte ab Zeile 501 eine stille
      // Teilliste mit Status ok. Ist die Liste größer als der Lesepfad (5 000 Zeilen),
      // gilt das Ergebnis als NICHT lesbar — kein stiller Teil-Erfolg.
      let url: string | null = `${this.list(LIST.useCases)}/items?$top=500`;
      const alle: SpUseCaseRow[] = [];
      let seiten = 0;
      while (url) {
        seiten++;
        const r: SPHttpClientResponse = await this._sp.get(
          url,
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
        ((d.value || []) as SpUseCaseRow[]).forEach(z => alle.push(z));
        url = d['odata.nextLink'] || d['@odata.nextLink'] || null;
        if (url && (alle.length >= 5000 || seiten >= 10)) {
          this.lastUseCasesReadStatus = 0;
          this.lastReadError = `Mehr als ${alle.length} Use Cases — der Lesepfad ist gekappt.`;
          console.warn(`[AIUC] ${this.lastReadError}`);
          return null;
        }
      }
      const rows = alle.map(row => this.mapUseCase(row));
      rows.sort((a, b) => (a.reihenfolge - b.reihenfolge) || a.titel.localeCompare(b.titel, 'de'));
      return rows;
    } catch (e) {
      // Status 0: Ein Fehler nach einer angenommenen Seite (200) soll nicht als „HTTP 200" neben
      // der Fehlermeldung stehen.
      this.lastUseCasesReadStatus = 0;
      this.lastReadError = String(e);
      return null;
    }
  }

  /**
   * Ein neues Feld hier heißt: auch in `SCHREIBBAR` in `utils/aenderungen.ts` eintragen — sonst
   * schreibt `saveUseCase` es beim Ändern nie (dort wird nur Geändertes aus dieser Liste gesendet).
   */
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
    // Je Link nur, was mitkommt: `nurGeaendertes` schickt beim Ändern nur die geänderten Schlüssel.
    // Alle fünf ungefragt zu schreiben löschte den Link, den eine andere Person inzwischen nachgetragen
    // hatte (Gegenprüfung 29.09.2026).
    if (uc.ressourcen) {
      const r = uc.ressourcen;
      if (r.sourceCode !== undefined) body.LinkSourceCode = r.sourceCode || '';
      if (r.deployment !== undefined) body.LinkDeployment = r.deployment || '';
      if (r.deploymentGuide !== undefined) body.LinkGuide = r.deploymentGuide || '';
      if (r.wiki !== undefined) body.LinkWiki = r.wiki || '';
      if (r.video !== undefined) body.LinkVideo = r.video || '';
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
      // 404: Die Zeile ist schon weg (eine andere Person war schneller). Das Ziel ist
      // erreicht — als Fehler gemeldet, blieb die Kachel stehen und das Protokoll
      // bekam eine falsche „Löschen fehlgeschlagen"-Zeile.
      return r.ok || r.status === 404;
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
        const rr = await this._post(`${zeile}/AttachmentFiles/getByFileName('${encodeURIComponent(fn.replace(/'/g, "''"))}')/recycleObject`, {});
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

  /** Den Merker setzen, damit die Erstbefuellung genau einmal passiert. `false` = nicht gespeichert. */
  public async merkeErstbefuellung(anzahl: number): Promise<boolean> {
    return this.log(0, SEED_MARKER, `${anzahl} Start-Use-Cases angelegt`);
  }

  /**
   * Protokollzeile schreiben.
   *
   * Ein fehlendes Protokoll darf keine Aktion verhindern — deshalb wirft das nie.
   * Aber die Rückgabe sagt die Wahrheit: `true` nur, wenn SharePoint die Zeile
   * angenommen hat. Bis v1.3 wertete die Funktion den Status nicht aus; bei 400
   * (Spalte fehlt), 403 oder 429 galt der Eintrag als geschrieben, und das
   * Löschen lief ohne Protokollzeile weiter, obwohl die Reihenfolge „prüfbare
   * Nebenbuchhaltung zuerst" genau das verhindern soll (Review 29.09.2026).
   */
  public async log(useCaseId: number, aktion: string, detail: string): Promise<boolean> {
    try {
      const r = await this._postItem(LIST.log, {
        'Title': aktion,
        'UseCaseId': useCaseId,
        'Aktion': aktion,
        'Detail': detail,
        'Wer': this.context.pageContext.user.email,
      });
      if (!r.ok) console.warn(`[AIUC] Protokoll schreiben (${aktion}): HTTP ${r.status}`);
      return r.ok;
    } catch {
      return false;
    }
  }
}
