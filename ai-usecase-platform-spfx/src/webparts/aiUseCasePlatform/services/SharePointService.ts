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
import { withThrottleRetry } from '../utils/spThrottle';
import { isCurrentUser } from '../utils/sessionIdentities';
import { LIST } from '../constants';
import { UserRole, UseCase, UseCaseStatus, Bewertung, AufrufArt } from '../types';

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

export class SharePointService {
  private context: WebPartContext;
  private siteUrl: string;

  /** HTTP-Status des letzten Rollen-Lesens (0 = Ausnahme oder nie gelesen). */
  public lastRolesReadStatus = 0;
  /** HTTP-Status des letzten Use-Case-Lesens. */
  public lastUseCasesReadStatus = 0;

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

  /**
   * Eine Spalte anlegen — aber nur, wenn sie fehlt.
   *
   * Das Pruefen vorweg ist kein Feinschliff: Ohne es laeuft bei JEDEM Start
   * ein Anlage-Versuch gegen eine existierende Spalte, SharePoint antwortet
   * mit 500, und der Browser protokolliert das rot, bevor unser `catch` es
   * sieht (DEX v30.65).
   */
  private async ensureField(
    listName: string,
    internalName: string,
    fieldTypeKind: number,
    extra?: Record<string, unknown>,
  ): Promise<void> {
    try {
      const probe = await this._sp.get(
        `${this.list(listName)}/fields/getbytitle('${encodeURIComponent(internalName)}')`,
        SPHttpClient.configurations.v1,
      );
      if (probe.ok) return;
    } catch { /* nicht lesbar -> Anlegen versuchen */ }

    try {
      await this._post(`${this.list(listName)}/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': internalName,
        'FieldTypeKind': fieldTypeKind,
        'Required': false,
        ...(extra || {}),
      });
      // In die Standardansicht aufnehmen, sonst ist die Spalte in SharePoint
      // selbst unsichtbar und niemand kann Daten von Hand nachsehen.
      await this._post(
        `${this.list(listName)}/views/getbytitle('All Items')/viewfields/addviewfield('${encodeURIComponent(internalName)}')`,
        {},
      ).catch(() => undefined);
    } catch (e) {
      console.warn(`[AIUC] Spalte ${internalName} auf ${listName} konnte nicht angelegt werden:`, e);
    }
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
    await this.ensureField(name, 'UserName', FIELD.text);
    await this.ensureField(name, 'Role', FIELD.choice, {
      '__metadata': { 'type': 'SP.FieldChoice' },
      'Choices': { 'results': ['Admin', 'Kurator', 'User'] },
    });
    await this.ensureField(name, 'AssignedBy', FIELD.text);
    await this.ensureField(name, 'AssignedDate', FIELD.dateTime);
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
      const r = await this._sp.get(
        `${this.list(LIST.roles)}/items?$select=Id,Title,UserName,Role,AssignedBy,AssignedDate&$top=500`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      this.lastRolesReadStatus = r.status;
      if (!r.ok) return null;
      const d = await r.json();
      return (d.value || []) as Array<{ Id: number; Title: string; UserName: string; Role: string; AssignedBy: string; AssignedDate: string }>;
    } catch {
      return null;
    }
  }

  public async addRole(userEmail: string, userName: string, role: UserRole, assignedBy: string): Promise<boolean> {
    try {
      const r = await this._post(`${this.list(LIST.roles)}/items`, {
        '__metadata': { 'type': `SP.Data.${LIST.roles}ListItem` },
        'Title': userEmail,
        'UserName': userName,
        'Role': role,
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
      const r = await this._merge(`${this.list(LIST.roles)}/items(${itemId})`, {
        '__metadata': { 'type': `SP.Data.${LIST.roles}ListItem` },
        'Role': role,
      });
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
    try {
      const r = await this._post(
        `${this.siteUrl}/_api/web/ensureuser`,
        { 'logonName': userEmail },
      );
      if (!r.ok) return null;
      const d = await r.json();
      const id = d.d?.Id ?? d.Id;
      return typeof id === 'number' ? id : null;
    } catch {
      return null;
    }
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

  /** Kuratoren duerfen Use Cases pflegen: Edit auf der Use-Case-Liste. */
  public async grantCuratorPermissions(userEmail: string): Promise<string[]> {
    const fehlend: string[] = [];
    const id = await this.resolveUserId(userEmail);
    if (!id) return ['Person nicht aufloesbar'];
    if (!(await this.grantVerified(this.list(LIST.useCases), id, ROLE_DEF.edit, 'grantCurator/useCases'))) {
      fehlend.push('Use-Case-Liste (Bearbeiten)');
    }
    if (!(await this.grantVerified(this.list(LIST.log), id, ROLE_DEF.contribute, 'grantCurator/log'))) {
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
      if (!chk.ok) return true; // 404/500 = keine Zuweisung mehr da
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
    await this.ensureField(name, 'Kurzbeschreibung', FIELD.note);
    await this.ensureField(name, 'Beschreibung', FIELD.note, { 'RichText': true });
    await this.ensureField(name, 'Bereich', FIELD.text);
    await this.ensureField(name, 'UcStatus', FIELD.choice, {
      '__metadata': { 'type': 'SP.FieldChoice' },
      'Choices': { 'results': ['Geplant', 'InArbeit', 'Live', 'Archiviert'] },
    });
    for (const dim of ['SalesRelevanz', 'Machbarkeit', 'DemoTauglichkeit']) {
      await this.ensureField(name, dim, FIELD.choice, {
        '__metadata': { 'type': 'SP.FieldChoice' },
        'Choices': { 'results': ['Hoch', 'Mittel', 'Niedrig'] },
      });
    }
    await this.ensureField(name, 'AufrufArt', FIELD.choice, {
      '__metadata': { 'type': 'SP.FieldChoice' },
      'Choices': { 'results': ['fenster', 'eingebettet'] },
    });
    for (const link of ['LinkSourceCode', 'LinkDeployment', 'LinkGuide', 'LinkWiki', 'LinkVideo', 'BildUrl']) {
      await this.ensureField(name, link, FIELD.note);
    }
    await this.ensureField(name, 'Reihenfolge', FIELD.number);
    await this.ensureField(name, 'BetreuerEmails', FIELD.note);
    await this.ensureField(name, 'BetreuerNamen', FIELD.note);
    await this.ensureField(name, 'Schlagworte', FIELD.note);

    return { isNewlyCreated: !existed };
  }

  public async ensureLogList(): Promise<void> {
    const name = LIST.log;
    if (!(await this.listExists(name))) {
      await this.createList(name, 'Aenderungsprotokoll der AI Use Case Platform');
    }
    await this.ensureField(name, 'UseCaseId', FIELD.number);
    await this.ensureField(name, 'Aktion', FIELD.text);
    await this.ensureField(name, 'Detail', FIELD.note);
    await this.ensureField(name, 'Wer', FIELD.text);
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
      geaendertVon: row.Editor?.Title || '',
    };
  }

  /**
   * Alle Use Cases lesen.
   *
   * `null` heisst nicht lesbar. Eine leere Kachelwand und ein Rechte-Fehler
   * sehen sonst gleich aus — und die Kachelwand ist die ganze App.
   */
  public async getUseCases(): Promise<UseCase[] | null> {
    this.lastUseCasesReadStatus = 0;
    try {
      const select = 'Id,Title,Kurzbeschreibung,Beschreibung,Bereich,UcStatus,SalesRelevanz,Machbarkeit,'
        + 'DemoTauglichkeit,AufrufArt,LinkSourceCode,LinkDeployment,LinkGuide,LinkWiki,LinkVideo,'
        + 'BildUrl,Reihenfolge,BetreuerEmails,BetreuerNamen,Schlagworte,Modified,Editor/Title';
      const r = await this._sp.get(
        `${this.list(LIST.useCases)}/items?$select=${select}&$expand=Editor&$orderby=Reihenfolge asc,Title asc&$top=500`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } },
      );
      this.lastUseCasesReadStatus = r.status;
      if (!r.ok) return null;
      const d = await r.json();
      return ((d.value || []) as SpUseCaseRow[]).map(row => this.mapUseCase(row));
    } catch {
      return null;
    }
  }

  private toRow(uc: Partial<UseCase>): Record<string, unknown> {
    const body: Record<string, unknown> = {
      '__metadata': { 'type': `SP.Data.${LIST.useCases}ListItem` },
    };
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
      const r = await this._post(`${this.list(LIST.useCases)}/items`, this.toRow(uc));
      if (!r.ok) return null;
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
      const r = await this._merge(`${this.list(LIST.useCases)}/items(${id})`, this.toRow(uc));
      return r.ok;
    } catch (e) {
      console.error('[AIUC] updateUseCase:', e);
      return false;
    }
  }

  public async deleteUseCase(id: number): Promise<boolean> {
    try {
      const r = await this._delete(`${this.list(LIST.useCases)}/items(${id})`);
      return r.ok;
    } catch (e) {
      console.error('[AIUC] deleteUseCase:', e);
      return false;
    }
  }

  /** Protokollzeile schreiben. Best-effort — ein fehlendes Protokoll darf keine Aktion verhindern. */
  public async log(useCaseId: number, aktion: string, detail: string): Promise<void> {
    try {
      await this._post(`${this.list(LIST.log)}/items`, {
        '__metadata': { 'type': `SP.Data.${LIST.log}ListItem` },
        'Title': aktion,
        'UseCaseId': useCaseId,
        'Aktion': aktion,
        'Detail': detail,
        'Wer': this.context.pageContext.user.email,
      });
    } catch { /* Protokoll ist Nebenbuchhaltung */ }
  }
}
