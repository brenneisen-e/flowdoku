/**
 * SharePoint Service - zentrale Klasse für alle SP REST API Aufrufe
 *
 * Erstellt Listen automatisch, liest/schreibt Einträge.
 * Nutzt den SPFx-Context für authentifizierte Aufrufe.
 *
 * - Eike, Maerz 2026
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { withThrottleRetry } from '../utils/spThrottle';
import { dlog } from '../utils/debugLog';
import { isCurrentUser } from '../utils/sessionIdentities';

export class SharePointService {
  private context: WebPartContext;
  private siteUrl: string;

    /**
   * v29.48/v29.50: Ersatz für `this.context.spHttpClient` — gleiche Signatur.
   *
   * **Nur `post` wiederholt.** Die Drosselung, um die es ging, trifft das
   * SPEICHERN (21 Sub-Events in einem Save). Lesezugriffe hier ebenfalls durch
   * die Schranke zu schicken, war ein Fehler: Der Start der App besteht fast
   * nur aus GETs, und ein einziges 429 legte damit den gesamten Bootvorgang
   * still — die Seite stand bei 8 %, der Browser meldete „reagiert nicht".
   * Ein abgelehnter Lesezugriff war vorher ein fehlendes Stück Anzeige; das
   * ist unschön, aber die App lebt. Deshalb geht `get` wieder direkt raus.
   *
   * Wer neu dazuschreibt, nimmt `this._sp`.
   */
  private _sp = {
    get: (url: string, cfg: SPHttpClientConfiguration, options?: ISPHttpClientOptions): Promise<SPHttpClientResponse> =>
      this.context.spHttpClient.get(url, cfg, options),
    post: (url: string, cfg: SPHttpClientConfiguration, options?: ISPHttpClientOptions): Promise<SPHttpClientResponse> =>
      withThrottleRetry(() => this.context.spHttpClient.post(url, cfg, options), url),
  };

  constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  /**
   * Prüfen ob eine Liste existiert
   */
  public async listExists(listName: string): Promise<boolean> {
    try {
      const response: SPHttpClientResponse = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')`,
        SPHttpClient.configurations.v1
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Rollen-Liste auf dem SharePoint erstellen (falls noch nicht vorhanden)
   *
   * Spalten:
   * - Title (built-in) -> wird als UserEmail genutzt
   * - UserName (Text)
   * - Role (Choice: SuperAdmin, EventAdmin, User)
   * - UserLocation (Text)
   * - AssignedBy (Text)
   * - AssignedDate (DateTime)
   */
  // v6.34: Return-Wert hinzugefügt — Caller muss wissen, ob die Liste gerade
  // frisch angelegt wurde (Erstinstallation). Sonst gab es einen kritischen
  // Auto-Admin-Bug: leere `getRoles`-Antwort (z.B. durch 403 Forbidden) wurde
  // fälschlich als "Erstinstallation" interpretiert und der aufrufende User
  // zum Admin befördert.
  public async ensureRolesList(): Promise<{ isNewlyCreated: boolean }> {
    const listName = 'DEX_Roles';
    const exists = await this.listExists(listName);

    if (exists) {
      // Berechtigungen prüfen und ggf. nachträglich setzen
      await this.ensureRolesListPermissions(listName);

      // Choice-Werte des Role-Feldes auf neue Benennung migrieren
      try {
        await this._sp.post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Role')`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE',
            },
            body: JSON.stringify({
              '__metadata': { 'type': 'SP.FieldChoice' },
              'Choices': { 'results': ['Admin', 'IT-Admin', 'Organizer', 'F&A', 'User'] },
            }),
          }
        );
      } catch { /* Choice-Update optional */ }

      // v18.5: IsPowerUser-Spalte auf bestehenden Listen nachziehen.
      //
      // v30.65: ERST PRÜFEN, dann anlegen. Vorher lief der Anlage-Versuch bei
      // JEDEM App-Start und bei JEDEM Nutzer — die Spalte existiert längst,
      // SharePoint antwortet mit HTTP 500, der `catch` verschluckt es, und der
      // Browser schreibt trotzdem eine rote Zeile in die Konsole (er protokolliert
      // die Antwort, bevor unser Code sie sieht). Das war der wiederkehrende
      // „DEX_Roles/fields 500" im gemeldeten Log: kein Fehler in der Sache,
      // aber einer, der bei jedem Start nach einem aussah — und der echte
      // Meldungen daneben unglaubwürdig macht.
      let hasPowerUser = false;
      try {
        const probe = await this._sp.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('IsPowerUser')`,
          SPHttpClient.configurations.v1
        );
        hasPowerUser = probe.ok;
      } catch { hasPowerUser = false; }
      try {
        if (hasPowerUser) throw new Error('__skip__');
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
          {
            '__metadata': { 'type': 'SP.Field' },
            'Title': 'IsPowerUser',
            'FieldTypeKind': 8, // Boolean
            'Required': false,
          }
        );
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/views/getbytitle('All Items')/viewfields/addviewfield('IsPowerUser')`,
          {}
        );
      } catch { /* '__skip__' = Spalte ist schon da; sonst best-effort */ }

      return { isNewlyCreated: false };
    }

    // Liste erstellen
    const listPayload = {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Rollenverwaltung für die DEX Event Experience Platform',
      'BaseTemplate': 100, // Generic List
      'AllowContentTypes': false,
    };

    await this._post(
      `${this.siteUrl}/_api/web/lists`,
      listPayload
    );

    // Spalten hinzufügen
    const columns = [
      {
        '__metadata': { 'type': 'SP.Field' },
        'Title': 'UserName',
        'FieldTypeKind': 2, // Text
        'Required': true,
      },
      {
        '__metadata': { 'type': 'SP.FieldChoice' },
        'Title': 'Role',
        'FieldTypeKind': 6, // Choice
        'Required': true,
        'Choices': { 'results': ['Admin', 'Organizer', 'User'] },
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'Title': 'UserLocation',
        'FieldTypeKind': 2, // Text
        'Required': false,
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'Title': 'AssignedBy',
        'FieldTypeKind': 2, // Text
        'Required': false,
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'Title': 'AssignedDate',
        'FieldTypeKind': 4, // DateTime
        'Required': false,
      },
      {
        // v18.5: Power-User-Flag (Zusatz auf einem Organizer, keine eigene Rolle).
        '__metadata': { 'type': 'SP.Field' },
        'Title': 'IsPowerUser',
        'FieldTypeKind': 8, // Boolean (Yes/No)
        'Required': false,
      },
    ];

    for (const col of columns) {
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
        col
      );
    }

    // Default View aktualisieren: alle Spalten anzeigen
    try {
      const viewFields = ['Title', 'UserName', 'Role', 'UserLocation', 'AssignedBy', 'AssignedDate', 'IsPowerUser'];
      for (const field of viewFields) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/views/getbytitle('All Items')/viewfields/addviewfield('${field}')`,
          {}
        );
      }
    } catch {
      // View-Update ist optional
    }

    // Berechtigungen setzen: nur Site-Owners (SuperAdmins) dürfen die Liste sehen
    await this.setRolesListPermissions(listName);
    return { isNewlyCreated: true };
  }

  /**
   * Prüfen ob die Liste bereits eigene Berechtigungen hat, sonst setzen
   */
  private async ensureRolesListPermissions(listName: string): Promise<void> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=HasUniqueRoleAssignments`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        if (!data.HasUniqueRoleAssignments) {
          await this.setRolesListPermissions(listName);
        }
      }
    } catch {
      // Ignorieren
    }
  }

  /**
   * Eigene Berechtigungen für die Rollen-Liste setzen.
   *
   * Berechtigungskonzept:
   *   - SuperAdmin (Site Owners): Full Control (lesen + schreiben)
   *   - EventAdmin: Read (nur lesen, damit die App die Rolle erkennt)
   *   - User: kein Zugriff (App defaulted auf "User")
   */
  private async setRolesListPermissions(listName: string): Promise<void> {
    try {
      // 1. Berechtigungsvererbung aufheben
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );

      // 2. Site-Owners-Gruppe: Full Control (1073741829)
      const ownersResponse = await this._sp.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (ownersResponse.ok) {
        const ownersData = await ownersResponse.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${ownersData.Id}, roledefid=1073741829)`,
          {}
        );
      }

      // 3. Aktuellen User individuell berechtigen (falls nicht in Owners-Gruppe)
      const currentUserId = await this.getUserIdByEmail(this.context.pageContext.user.email);
      if (currentUserId) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${currentUserId}, roledefid=1073741829)`,
          {}
        );
      }
    } catch {
      // Berechtigungen konnten nicht gesetzt werden
    }
  }

  /**
   * Einem User Leseberechtigung auf die Rollen-Liste geben (für EventAdmins).
   * Ermittelt die User-ID per E-Mail und setzt Read-Berechtigung.
   */
  public async grantReadOnRolesList(userEmail: string): Promise<void> {
    try {
      // User-ID per E-Mail ermitteln
      const userId = await this.getUserIdByEmail(userEmail);
      if (!userId) {
        console.error('[DEX] grantReadOnRolesList: Keine User-ID für', userEmail);
        return;
      }

      // Sicherstellen dass Liste unique permissions hat
      await this.ensureListHasUniquePermissions('DEX_Roles');

      // Read = 1073741826 (Standard SharePoint ID, sprachunabhängig)
      const readRoleId = 1073741826;

      // Leseberechtigung setzen
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/roleassignments/addroleassignment(principalid=${userId}, roledefid=${readRoleId})`,
        {}
      );
    } catch (e) {
      console.error('[DEX] grantReadOnRolesList Error:', e);
    }
  }

  /**
   * Einem SuperAdmin Full Control auf die Rollen-Liste geben.
   */
  public async grantFullControlOnRolesList(userEmail: string): Promise<void> {
    try {
      const userId = await this.getUserIdByEmail(userEmail);
      if (!userId) {
        console.error('[DEX] grantFullControlOnRolesList: Keine User-ID für', userEmail);
        return;
      }

      // Sicherstellen dass DEX_Roles unique permissions hat
      await this.ensureListHasUniquePermissions('DEX_Roles');

      // Full Control = 1073741829
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
        {}
      );
      if (!response.ok) {
        const errorText = await response.text().catch(() => 'no body');
        console.error('[DEX] grantFullControlOnRolesList Fehler-Response:', response.status, errorText);
      }
    } catch (e) {
      console.error('[DEX] grantFullControlOnRolesList Error:', e);
    }
  }

  /**
   * Einem Admin Full Control auf die DEX_Events-Liste geben.
   * Stellt sicher dass die Liste eigene Berechtigungen hat bevor Rollen zugewiesen werden.
   */
  public async grantFullControlOnEventsList(userEmail: string): Promise<void> {
    try {
      const userId = await this.getUserIdByEmail(userEmail);
      if (!userId) {
        console.error('[DEX] grantFullControlOnEventsList: Keine User-ID für', userEmail);
        return;
      }

      // Sicherstellen dass DEX_Events unique permissions hat
      await this.ensureListHasUniquePermissions('DEX_Events');

      // Full Control = 1073741829
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
        {}
      );
      if (!response.ok) {
        const errorText = await response.text().catch(() => 'no body');
        console.error('[DEX] grantFullControlOnEventsList Fehler-Response:', response.status, errorText);
      }
    } catch (e) {
      console.error('[DEX] grantFullControlOnEventsList Error:', e);
    }
  }

  /**
   * Einem Organizer Full Control auf die DEX_Events-Liste + Manage auf Site-Ebene geben.
   * Manage auf Site-Ebene erlaubt Subsite-Erstellung (webs/add).
   */
  public async grantOrganizerPermissions(userEmail: string): Promise<void> {
    try {
      const userId = await this.getUserIdByEmail(userEmail);
      if (!userId) return;

      // 1. Full Control auf DEX_Events (wie Admin)
      await this.ensureListHasUniquePermissions('DEX_Events');
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
        {}
      );

      // 2. Full Control auf Site-Ebene (für Subsite-Erstellung)
      // Full Control = 1073741829
      await this._post(
        `${this.siteUrl}/_api/web/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
        {}
      );
    } catch (e) {
      console.error('[DEX] grantOrganizerPermissions Error:', e);
    }
  }

  /**
   * Einem User Contribute-Rechte auf die DEX_Events-Liste geben (für Organizer).
   */
  public async grantContributeOnEventsList(userEmail: string): Promise<void> {
    try {
      const userId = await this.getUserIdByEmail(userEmail);
      if (!userId) return;
      await this.ensureListHasUniquePermissions('DEX_Events');
      // Contribute = 1073741827
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741827)`,
        {}
      );
    } catch (e) {
      console.error('[DEX] grantContributeOnEventsList Error:', e);
    }
  }

  /**
   * Einem User die Berechtigung auf die DEX_Events-Liste entziehen.
   *
   * v30.67: Liefert, ob der Entzug WIRKT (nachgelesen, nicht nur abgeschickt).
   * Vorher `void` mit stillem catch — ein gescheiterter Entzug sah für die App
   * genauso aus wie ein gelungener, und die Rollenverwaltung meldete „Rolle
   * entfernt", während die Rechte stehen blieben.
   */
  public async revokeAccessOnEventsList(userEmail: string): Promise<boolean> {
    return this._revokePrincipalOn(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')`, userEmail, 'revokeAccessOnEventsList');
  }

  /**
   * Einem User die Berechtigung auf die Rollen-Liste entziehen.
   * v30.67: Rückgabe wie `revokeAccessOnEventsList`.
   */
  public async revokeAccessOnRolesList(userEmail: string): Promise<boolean> {
    return this._revokePrincipalOn(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')`, userEmail, 'revokeAccessOnRolesList');
  }

  /**
   * v30.67: Spiegelbild zu `grantOrganizerPermissions`, das ZWEI Rechte
   * vergibt — Full Control auf DEX_Events UND Full Control auf dem Web-Root
   * (`/_api/web/roleassignments`, nötig für `webs/add`). Für das zweite gab es
   * bis v30.66 kein Gegenstück: Wer auf „User" zurückgestuft oder aus
   * DEX_Roles gelöscht wurde, behielt Full Control auf der gesamten
   * Site-Collection — inklusive aller Teilnehmer-Subsites und „Manage
   * Permissions", womit sich die Person die entzogenen Listenrechte selbst
   * zurückgeben konnte.
   *
   * Entzogen wird NUR die direkte Zuweisung dieses einen Principals
   * (`getbyprincipalid`), nie eine Gruppe. Rechte über Owners/Members/
   * Visitors bleiben unberührt.
   *
   * Selbstschutz: Für die angemeldete Person läuft der Entzug NICHT — ein
   * Admin, der die Rollenverwaltung an sich selbst testet, würde sich sonst
   * aus der Site aussperren. Der Aufrufer (RoleContext) prüft das ebenfalls;
   * hier ist die letzte Sperre für jeden anderen Aufrufer.
   */
  public async revokeSiteAccess(userEmail: string): Promise<boolean> {
    if (isCurrentUser(this.context, userEmail)) {
      console.warn(`[DEX] revokeSiteAccess: Selbstschutz — Web-Rechte von ${userEmail} (angemeldete Person) werden NICHT entzogen.`);
      return false;
    }
    return this._revokePrincipalOn(`${this.siteUrl}/_api/web`, userEmail, 'revokeSiteAccess');
  }

  /**
   * v30.67: Direkte Rollenzuweisung EINES Principals auf einem Securable
   * (Web oder Liste) entfernen und das Ergebnis NACHLESEN.
   *
   * Warum nachlesen statt nur den DELETE-Status zu prüfen: SharePoint
   * antwortet auf `getbyprincipalid(...)` ohne vorhandene Zuweisung je nach
   * Fall mit 404 oder 500 — das ist dann KEIN Fehler, es gab nichts zu
   * entziehen. Und ein 403 (kein „Manage Permissions") ließe die Rechte
   * stehen. Die einzige belastbare Aussage ist deshalb der Blick in die
   * Zuweisungsliste danach: Steht der Principal noch drin, ist der Entzug
   * gescheitert — egal, was der DELETE gemeldet hat.
   *
   * @returns true = die Person hat auf diesem Securable keine direkte
   *   Zuweisung mehr. false = steht noch drin ODER nicht prüfbar; in beiden
   *   Fällen mit `console.warn('[DEX] …')`, denn ein still gescheiterter
   *   Entzug ist derselbe Zustand wie gar keiner — nur ohne dass es jemand
   *   weiß.
   */
  private async _revokePrincipalOn(scopeBase: string, userEmail: string, label: string): Promise<boolean> {
    let userId: number | null = null;
    try {
      userId = await this.getUserIdByEmail(userEmail);
    } catch (e) {
      console.warn(`[DEX] ${label}: User-Id für ${userEmail} nicht auflösbar — nichts entzogen.`, e);
      return false;
    }
    if (!userId) {
      console.warn(`[DEX] ${label}: Keine User-Id für ${userEmail} — nichts entzogen.`);
      return false;
    }

    let deleteStatus = 0;
    try {
      const del = await this._sp.post(
        `${scopeBase}/roleassignments/getbyprincipalid(${userId})`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=verbose', 'X-HTTP-Method': 'DELETE' } }
      );
      deleteStatus = del.status;
    } catch (e) {
      console.warn(`[DEX] ${label}: DELETE für ${userEmail} nicht abgesetzt — nichts entzogen.`, e);
      return false;
    }

    // Kontrolle: die Zuweisungen dieses Securables lesen und dem nextLink
    // folgen. Direkte Zuweisungen sind wenige — aber ein Kappen darf hier
    // nie „steht nicht mehr drin" bedeuten.
    try {
      let url: string | null = `${scopeBase}/roleassignments?$select=PrincipalId&$top=5000`;
      let guard = 0;
      while (url && guard < 20) {
        guard++;
        const chk = await this._sp.get(url, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
        if (!chk.ok) {
          console.warn(`[DEX] ${label}: Entzug für ${userEmail} nicht prüfbar (DELETE HTTP ${deleteStatus}, Kontrolle HTTP ${chk.status}) — gilt als NICHT entzogen.`);
          return false;
        }
        const d = await chk.json();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const rows: any[] = d.value || d.d?.results || [];
        if (rows.some(r => Number(r.PrincipalId) === userId)) {
          console.warn(`[DEX] ${label}: ${userEmail} steht nach DELETE (HTTP ${deleteStatus}) noch in den Zuweisungen von ${scopeBase} — NICHT entzogen.`);
          return false;
        }
        url = d['odata.nextLink'] || d['@odata.nextLink'] || (d.d && d.d.__next) || null;
      }
      if (url) {
        console.warn(`[DEX] ${label}: Kontrolle nach ${guard} Seiten abgebrochen — gilt als NICHT entzogen.`);
        return false;
      }
      return true;
    } catch (e) {
      console.warn(`[DEX] ${label}: Kontrolle für ${userEmail} fehlgeschlagen — gilt als NICHT entzogen.`, e);
      return false;
    }
  }

  /**
   * User-ID per E-Mail aus SharePoint ermitteln.
   * Nutzt ensureuser um den User anzulegen falls er die Site noch nie besucht hat.
   */
  private async getUserIdByEmail(email: string): Promise<number | null> {
    // Versuch 1: ensureuser mit Email direkt
    try {
      const ensureResponse = await this._post(
        `${this.siteUrl}/_api/web/ensureuser`,
        { 'logonName': email }
      );
      if (ensureResponse.ok) {
        const ensureData = await ensureResponse.json();
        const id = ensureData.d?.Id || ensureData.Id;
        if (id) {
          return id;
        }
      } else {
        console.warn('[DEX] ensureuser fehlgeschlagen:', email, ensureResponse.status);
      }
    } catch (e) {
      console.warn('[DEX] ensureuser Error:', email, e);
    }

    // Versuch 2: ensureuser mit Claims-Format
    try {
      const ensureResponse2 = await this._post(
        `${this.siteUrl}/_api/web/ensureuser`,
        { 'logonName': `i:0#.f|membership|${email}` }
      );
      if (ensureResponse2.ok) {
        const ensureData2 = await ensureResponse2.json();
        const id2 = ensureData2.d?.Id || ensureData2.Id;
        if (id2) {
          return id2;
        }
      }
    } catch { /* ignore */ }

    // Versuch 3: direkt per Email suchen
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(email)}')?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        const id3 = data.Id || null;
        return id3;
      }
    } catch { /* User nicht gefunden */ }

    console.error('[DEX] getUserIdByEmail: User nicht gefunden:', email);
    return null;
  }

  /**
   * Sicherstellen dass eine Liste eigene (unique) Berechtigungen hat.
   * Falls die Liste noch von der Site erbt, wird die Vererbung aufgehoben
   * und die bestehenden Berechtigungen kopiert.
   */
  private async ensureListHasUniquePermissions(listName: string): Promise<void> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=HasUniqueRoleAssignments`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        if (!data.HasUniqueRoleAssignments) {
          // copyRoleAssignments=true: bestehende Berechtigungen beibehalten
          await this._post(
            `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=true, clearSubscopes=true)`,
            {}
          );
        } else {
          // Liste hat bereits eigene Berechtigungen
        }
      }
    } catch (e) {
      console.error('[DEX] ensureListHasUniquePermissions Error für', listName, ':', e);
    }
  }

  /**
   * RoleDefinition-ID per Name ermitteln (z.B. "Read", "Full Control")
   */
  private async getRoleDefinitionId(roleName: string): Promise<number | null> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/roledefinitions/getbyname('${encodeURIComponent(roleName)}')?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        return data.Id || null;
      }
    } catch { /* Role nicht gefunden */ }
    return null;
  }

  /**
   * Alle Rollen-Einträge lesen
   */
  // v6.34: Rückgabetyp erweitert — `null` bei API-Fehler (z.B. 403 Forbidden
  // oder Netzwerk-Fehler), `[]` nur bei wirklich leerer Liste. Der Caller
  // muss unterscheiden können, sonst löst ein Read-Fehler fälschlich die
  // "Erstinstallation"-Logik aus und macht den aktuellen User zum Admin.
  public async getRoles(): Promise<Array<{
    Id: number;
    Title: string; // UserEmail
    UserName: string;
    Role: string;
    UserLocation: string;
    AssignedBy: string;
    AssignedDate: string;
    IsPowerUser?: boolean;
  }> | null> {
    // v18.5: IsPowerUser mitlesen. Falls die Spalte auf einem alten Tenant
    // noch fehlt (ensureRolesList noch nicht durchgelaufen), würde der
    // $select mit IsPowerUser einen HTTP 400 werfen — daher Fallback auf den
    // Select OHNE IsPowerUser.
    const baseSelect = 'Id,Title,UserName,Role,UserLocation,AssignedBy,AssignedDate';
    // v30.67: `$top=5000` UND dem nextLink folgen. Ohne `$top` liefert
    // SharePoint 100 Zeilen und den Rest nur über `odata.nextLink` — die
    // Methode gab `data.value` 1:1 zurück. Ab der 101. Zeile (sortiert nach
    // Role,UserName) fehlte eine Organizerin damit im Ergebnis, `initRoles`
    // fand sie nicht und setzte sie auf 'User': kein Organizer Center, keine
    // eigenen Events — obwohl ihr Eintrag in DEX_Roles korrekt stand. Jede
    // andere Item-Abfrage der Service-Schicht setzt ein `$top`; diese war die
    // einzige Ausnahme, und ausgerechnet an ihr hängt die Rechteermittlung.
    //
    // Fail-closed: Bricht das Lesen NACH der ersten Seite ab, kommt `null`
    // (= „nicht lesbar"), nie die halbe Liste — eine halbe Rollenliste sähe
    // für die Aufrufer wie eine vollständige aus.
    const itemsUrl = (withPowerUser: boolean): string =>
      `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$select=${baseSelect}${withPowerUser ? ',IsPowerUser' : ''}&$orderby=Role,UserName&$top=5000`;
    try {
      let response = await this._sp.get(itemsUrl(true), SPHttpClient.configurations.v1);
      if (!response.ok) {
        // Retry ohne IsPowerUser (Spalte existiert evtl. noch nicht).
        response = await this._sp.get(itemsUrl(false), SPHttpClient.configurations.v1);
        if (!response.ok) return null;
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const out: any[] = [];
      let guard = 0;
      for (;;) {
        const data = await response.json();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const page: any[] = data.value || data.d?.results || [];
        for (const r of page) out.push(r);
        const next: string | null = data['odata.nextLink'] || data['@odata.nextLink'] || (data.d && data.d.__next) || null;
        if (!next) break;
        guard++;
        if (guard >= 20) {
          console.warn('[DEX] getRoles: mehr als 20 Seiten — Lesen abgebrochen, DEX_Roles gilt als nicht lesbar.');
          return null;
        }
        response = await this._sp.get(next, SPHttpClient.configurations.v1);
        if (!response.ok) {
          console.warn(`[DEX] getRoles: Folgeseite nicht lesbar (HTTP ${response.status}) — DEX_Roles gilt als nicht lesbar, nicht als leer.`);
          return null;
        }
      }
      return out;
    } catch (e) {
      console.warn('[DEX] getRoles: Lesen fehlgeschlagen — DEX_Roles gilt als nicht lesbar.', e);
      return null;
    }
  }

  /**
   * v28.65: Claims-Login-Tokens in der Rollenliste reparieren.
   *
   * Betroffen ist `UserName` — die Namen sind beim Zuweisen der Rolle aus dem
   * Anzeigenamen des Zuweisenden bzw. der Personensuche entstanden und können
   * dasselbe Token enthalten wie die Teilnehmerzeilen (Hintergrund in
   * `utils/displayName.ts`). Die E-Mail steht im Feld `Title`.
   */
  public async repairClaimNamesInRoles(): Promise<{ scanned: number; hits: number; fixed: number; failed: number }> {
    const out = { scanned: 0, hits: 0, fixed: 0, failed: 0 };
    const looksLikeClaim = (s: string): boolean => /\|membership\b|^i:0[#|]|^c:0|0#\.[a-z]\||^\d+#\./i.test((s || '').trim());
    const rows = await this.getRoles();
    if (!rows) return out;
    out.scanned = rows.length;
    const affected = rows.filter(r => looksLikeClaim(r.UserName || '') || looksLikeClaim(r.AssignedBy || ''));
    out.hits = affected.length;
    for (const r of affected) {
      const email = (r.Title || '').trim();
      let name = '';
      try {
        const prof = await this.searchUserByEmail(email);
        name = (prof && prof.displayName && !looksLikeClaim(prof.displayName)) ? prof.displayName.trim() : '';
      } catch { /* nicht auflösbar */ }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const patch: Record<string, any> = {};
      if (looksLikeClaim(r.UserName || '')) patch['UserName'] = name || email;
      if (looksLikeClaim(r.AssignedBy || '')) patch['AssignedBy'] = '';
      if (Object.keys(patch).length === 0) continue;
      try {
        const resp = await this._sp.post(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items(${r.Id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
              'odata-version': '',
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE',
            },
            body: JSON.stringify(patch),
          },
        );
        if (resp.ok || resp.status === 204 || resp.status === 406) out.fixed++; else out.failed++;
      } catch { out.failed++; }
    }
    return out;
  }

  /**
   * v18.5: Power-User-Flag eines Rollen-Eintrags setzen/entfernen.
   */
  public async setPowerUser(itemId: number, isPowerUser: boolean): Promise<boolean> {
    try {
      const response = await this._sp.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({ IsPowerUser: isPowerUser }),
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Rolle eines bestimmten Users abfragen (per E-Mail)
   */
  public async getUserRole(email: string): Promise<string | null> {
    try {
      const response = await this._sp.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(email)}'&$select=Role&$top=1`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) return null;

      const data = await response.json();
      if (data.value && data.value.length > 0) {
        return data.value[0].Role;
      }
      return null;
    } catch {
      return null;
    }
  }

  /**
   * Neue Rolle zuweisen
   */
  public async addRole(
    userEmail: string,
    userName: string,
    role: string,
    userLocation: string,
    assignedBy: string
  ): Promise<boolean> {
    try {
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_RolesListItem' },
        'Title': userEmail,
        'UserName': userName,
        'Role': role,
        'UserLocation': userLocation,
        'AssignedBy': assignedBy,
        'AssignedDate': new Date().toISOString(),
      };

      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items`,
        payload
      );

      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Rolle aktualisieren
   */
  public async updateRole(itemId: number, role: string): Promise<boolean> {
    try {
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_RolesListItem' },
        'Role': role,
      };

      const headers: HeadersInit = {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      };

      const options: ISPHttpClientOptions = {
        headers: headers,
        body: JSON.stringify(payload),
      };

      const response = await this._sp.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        options
      );

      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Location eines Rollen-Eintrags aktualisieren
   */
  public async updateRoleLocation(itemId: number, location: string): Promise<boolean> {
    try {
      const response = await this._sp.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.Data.DEX_x005f_RolesListItem' },
            'UserLocation': location,
          }),
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Rolle entfernen
   */
  public async deleteRole(itemId: number): Promise<boolean> {
    try {
      const headers: HeadersInit = {
        'Accept': 'application/json;odata=verbose',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE',
      };

      const options: ISPHttpClientOptions = {
        headers: headers,
      };

      const response = await this._sp.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        options
      );

      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * User per E-Mail in Microsoft 365 suchen — robust gegen UPN != SMTP-Mismatches.
   * Gibt Name, Standort und JobTitle zurück falls gefunden.
   *
   * Strategie:
   *   1. Direkter Claim-Lookup mit `i:0#.f|membership|<email>` (schnell).
   *   2. Wenn leer oder kein DisplayName: per `siteusers/getbyemail` den echten
   *      LoginName auflösen (UPN-Claim), dann GetPropertiesFor damit erneut
   *      aufrufen. Deckt UPN != SMTP, Guest-Accounts, Alias-SMTP-Adressen ab.
   */
  // v28.11: Unternehmenszugehörigkeit („Deloitte Consulting" etc.) für eine
  // FREMDE Person via Graph — die SP-UserProfile-Property „Company" ist im
  // Tenant nicht zuverlässig gefüllt (gleiche Erkenntnis wie v24.33 beim
  // eigenen Profil, dort /me?$select=companyName). Best-effort, leer bei
  // fehlender Berechtigung o.ä.
  private async fetchCompanyViaGraph(email: string): Promise<string> {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = this.context as any;
      if (!ctx.msGraphClientFactory) return '';
      const client = await ctx.msGraphClientFactory.getClient('3');
      const me = await client.api(`/users/${encodeURIComponent(email)}`).select('companyName').get();
      return ((me && me.companyName) || '').trim();
    } catch { return ''; }
  }

  public async searchUserByEmail(email: string): Promise<{
    displayName: string;
    location: string;
    jobTitle: string;
    // v11.97: Department + Mobile zusätzlich aus dem SP-Profil — für die
    // Personal-Info-Card auf der Registration-Page wenn jemand für eine
    // andere Person registriert.
    department?: string;
    mobilePhone?: string;
    // v28.11: Unternehmenszugehörigkeit (SP-Profil, Fallback Graph).
    company?: string;
  } | null> {
    if (!email) return null;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const extract = (data: any): { displayName: string; location: string; jobTitle: string; department: string; mobilePhone: string; company: string } | null => {
      if (!data || !data.DisplayName) return null;
      const props: Array<{ Key: string; Value: string }> = data.UserProfileProperties || [];
      const getProp = (keys: string[]): string => {
        for (const k of keys) {
          const p = props.find(x => x.Key === k);
          if (p && p.Value) return p.Value;
        }
        return '';
      };
      return {
        displayName: data.DisplayName,
        location: getProp(['Office', 'SPS-Location', 'SPS-City', 'City']),
        jobTitle: getProp(['Title', 'SPS-JobTitle']),
        department: getProp(['Department', 'SPS-Department']),
        mobilePhone: getProp(['CellPhone', 'SPS-MobilePhone', 'MobilePhone']),
        company: getProp(['Company', 'SPS-Company', 'CompanyName', 'msOnline-CompanyName']),
      };
    };
    // v28.11: Company nachladen, wenn das SP-Profil sie nicht liefert.
    // (displayName im Constraint — ein rein-optionaler „weak type" würde
    // TS-Assignability für Objekte ohne company-Property brechen.)
    const withCompany = async <T extends { displayName: string; company?: string }>(hit: T): Promise<T> => {
      if (hit.company) return hit;
      const c = await this.fetchCompanyViaGraph(email);
      return c ? { ...hit, company: c } : hit;
    };

    // 1) Direkter Lookup per SMTP
    try {
      const directResp = await this._sp.get(
        `${this.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|${encodeURIComponent(email)}'&$select=DisplayName,UserProfileProperties`,
        SPHttpClient.configurations.v1
      );
      if (directResp.ok) {
        const data = await directResp.json();
        const hit = extract(data);
        if (hit && (hit.jobTitle || hit.location)) return await withCompany(hit);
        // DisplayName ohne Properties? Merken für Fallback-Default.
        if (hit) {
          // weiter zum LoginName-Pfad - vielleicht bringt der jobTitle
        }
      }
    } catch { /* weiter */ }

    // 2) Fallback: echten LoginName (UPN-Claim) via siteusers/getbyemail
    try {
      const siteUserResp = await this._sp.get(
        `${this.siteUrl}/_api/web/siteusers/getbyemail('${email.replace(/'/g, "''")}')?$select=LoginName,Title`,
        SPHttpClient.configurations.v1
      );
      if (siteUserResp.ok) {
        const su = await siteUserResp.json();
        const loginName: string = su.LoginName || su.d?.LoginName || '';
        const fallbackDisplayName: string = su.Title || su.d?.Title || '';
        if (loginName) {
          try {
            const profileResp = await this._sp.get(
              `${this.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(loginName)}'&$select=DisplayName,UserProfileProperties`,
              SPHttpClient.configurations.v1
            );
            if (profileResp.ok) {
              const pData = await profileResp.json();
              const hit = extract(pData);
              if (hit) return await withCompany(hit);
            }
          } catch { /* */ }
        }
        if (fallbackDisplayName) {
          return await withCompany({ displayName: fallbackDisplayName, location: '', jobTitle: '' });
        }
      }
    } catch { /* User nicht gefunden */ }
    return null;
  }

  /**
   * User-Suche mit Autocomplete: sucht nach Name oder Email-Fragment.
   * Nutzt die SharePoint ClientPeoplePickerSearchUser API.
   */
  public async searchUsers(query: string, includeInternational: boolean = false): Promise<Array<{
    email: string;
    displayName: string;
    location: string;
    jobTitle: string;
  }>> {
    if (!query || query.length < 2) return [];

    /**
     * SharePoint ClientPeoplePicker tokenisiert auf Whitespace. „Lastname, Firstname"
     * wird intern oft nur teilweise gematched, weil das Komma als Wort-Boundary zählt
     * aber nicht abgestrippt wird → der Suchbegriff „Mustermann," bringt weniger
     * Treffer als „Mustermann". Außerdem bevorzugt SP-Picker die Reihenfolge
     * `Firstname Lastname`-Convention; bei „Lastname Firstname" (ohne Komma, falsche
     * Reihenfolge) findet er trotzdem oft zu wenig.
     *
     * Lösung: bei jedem Such-Request wird die Query in mehrere Varianten zerlegt und
     * parallel durchsucht; Resultate werden per Email dedupliziert. Damit findet der
     * Picker:
     *   - „Mustermann, Max"      (mit Komma — SP-Standard für Tenants mit DE-Locale)
     *   - „Max Mustermann"       (Firstname Lastname, ohne Komma)
     *   - „Mustermann Max"       (Lastname Firstname, ohne Komma)
     *   - „Mustermann"           (nur Lastname)
     *   - „Mu"                   (Anfang, wird durch Picker-Built-in abgefangen)
     */
    const cleanQuery = query.trim();
    const variants = new Set<string>();
    variants.add(cleanQuery);
    if (cleanQuery.indexOf(',') >= 0) {
      // „Lastname, Firstname" → drei zusätzliche Varianten
      const parts = cleanQuery.split(',').map(s => s.trim()).filter(Boolean);
      if (parts.length >= 2) {
        const lastname = parts[0];
        const firstname = parts.slice(1).join(' ').trim();
        variants.add(`${firstname} ${lastname}`); // Firstname Lastname
        variants.add(`${lastname} ${firstname}`); // Lastname Firstname (ohne Komma)
        if (lastname.length >= 2) variants.add(lastname); // nur Lastname
      } else if (parts.length === 1 && parts[0].length >= 2) {
        variants.add(parts[0]);
      }
    } else if (cleanQuery.indexOf(' ') >= 0 && cleanQuery.indexOf('@') < 0) {
      // v11.77: „Firstname Lastname" (ohne Komma) → SP-Picker findet das
      // nicht zuverlässig. In DE-Tenants ist das Standard-Display-Name-
      // Format „Lastname, Firstname" — die Such-API matched besser darauf.
      // Daher zusätzliche Varianten generieren: swap + Komma-Variante +
      // nur-Lastname + nur-Firstname.
      const tokens = cleanQuery.split(/\s+/).map(s => s.trim()).filter(Boolean);
      if (tokens.length === 2) {
        const a = tokens[0];
        const b = tokens[1];
        // Wir wissen nicht, ob „a b" Firstname-Lastname oder Lastname-Firstname
        // ist — probieren beide swap-Varianten mit Komma.
        variants.add(`${b}, ${a}`); // wenn a=Firstname, b=Lastname → „Lastname, Firstname"
        variants.add(`${a}, ${b}`); // wenn a=Lastname, b=Firstname → ebenfalls Komma-Form
        variants.add(`${b} ${a}`);  // umgekehrte Reihenfolge ohne Komma
        if (a.length >= 2) variants.add(a);
        if (b.length >= 2) variants.add(b);
      } else if (tokens.length > 2) {
        // Mehrere Worte (z.B. Doppelname) — als zweiten Versuch alle Worte
        // umgedreht und nur das letzte Wort (vermutlich Lastname) probieren.
        const reversed = [...tokens].reverse().join(' ');
        variants.add(reversed);
        const last = tokens[tokens.length - 1];
        if (last.length >= 2) variants.add(last);
      }
    }

    const seen = new Set<string>();
    const all: Array<{ email: string; displayName: string; location: string; jobTitle: string }> = [];

    for (const variant of Array.from(variants)) {
      try {
        const body = {
          'queryParams': {
            '__metadata': { 'type': 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
            'AllowEmailAddresses': true,
            'AllowMultipleEntities': false,
            'MaximumEntitySuggestions': 10,
            'QueryString': variant,
            'PrincipalType': 1, // Users only
            'PrincipalSource': 15,
            'SharePointGroupID': 0,
          },
        };

        const response = await this._post(
          `${this.siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
          body
        );

        if (!response.ok) continue;

        const data = await response.json();
        const resultsStr = data.d?.ClientPeoplePickerSearchUser || data.ClientPeoplePickerSearchUser || '[]';
        const results = JSON.parse(resultsStr);
        // eslint-disable-next-line no-console
        dlog('perf', `[DEX][perf][searchUsers] variant="${variant}" raw=${results.length}`);

        // v11.75: vor allem Variante 1 — SP-Picker liefert manchmal Treffer
        // mit leerem EntityData.Email (z.B. wenn der User noch keine SP-
        // Personalwand hat / Profil noch nicht angelegt). Früher wurde der
        // Treffer mit `filter(x => x.EntityData?.Email)` rausgeworfen — Namens-
        // Suchen wie „Inga Fuhr" kamen dann mit 0 Vorschlägen zurück.
        // Fallback: Email aus `Key` (z.B. „i:0#.f|membership|email@domain")
        // oder `Description` ziehen.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        for (const r of results as any[]) {
          let email = String(r.EntityData?.Email || '').toLowerCase();
          if (!email) {
            const candidates: string[] = [String(r.Key || ''), String(r.Description || '')];
            for (const c of candidates) {
              const m = c.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
              if (m) { email = m[0].toLowerCase(); break; }
            }
          }
          if (!email || !email.includes('@') || seen.has(email)) continue;
          seen.add(email);
          all.push({
            email,
            displayName: r.DisplayText || r.EntityData?.Title || '',
            location: '',
            jobTitle: r.EntityData?.Title || '',
          });
          if (all.length >= 10) break;
        }
        if (all.length >= 10) break;
      } catch {
        // Variant failed — continue with next variant
      }
    }

    // v11.85: Client-side Filter — die Variants-Logik (Single-Token-
    // Fallback wie „Nils" oder „Felt" einzeln) zog auch False-Positives an
    // (z.B. „Nilmara Santos" matched auf „Nils"-Prefix). Wenn der User
    // mehrere Tokens eingegeben hat, MUESSEN ALLE Tokens im displayName
    // ODER der Email der Person vorkommen (case-insensitive Substring).
    // Single-Token-Queries (1 Wort, ohne Komma) bleiben unangetastet —
    // dort ist Substring-Match auf einen Begriff genau das gewünschte
    // Verhalten.
    const queryTokens = cleanQuery
      .replace(/,/g, ' ')
      .split(/\s+/)
      .map(t => t.trim().toLowerCase())
      .filter(t => t.length >= 2);
    if (queryTokens.length >= 2) {
      const filtered = all.filter(u => {
        const hay = (u.displayName + ' ' + u.email).toLowerCase();
        return queryTokens.every(tok => hay.indexOf(tok) >= 0);
      });
      // Nur übernehmen wenn überhaupt etwas matched — sonst sind die
      // unscharfen Single-Token-Treffer immer noch besser als gar nichts.
      if (filtered.length > 0) {
        all.length = 0;
        all.push(...filtered);
      }
    }

    // v13.6: Member-Firm-Filter. Default: nur @deloitte.de (DEALL-Equivalent).
    // v26.57: Mit includeInternational=true sind ALLE Deloitte-Member-Firm-
    // Domains erlaubt — nicht nur @deloitte.com. Internationale Member-Firms
    // haben eigene Länder-Domains (Österreich @deloitte.at, Niederlande
    // @deloitte.nl, UK @deloitte.co.uk) und teils zusammengesetzte Domains.
    // v26.58: auch Domains mit Suffix direkt am Wort „deloitte" zulassen —
    // z. B. @deloitteCE.com (Deloitte Central Europe) oder @deloittedigital.com;
    // ebenso Subdomains wie @xy.deloitte.com. Nicht-Deloitte-Domains
    // (Gast-Accounts fremder Firmen, externe Tenants) bleiben in beiden Modi
    // geblockt — „deloitte" muss direkt am Domain-/Label-Anfang stehen.
    const isDeloitteDomain = (mail: string): boolean => {
      const at = mail.lastIndexOf('@');
      if (at < 0) return false;
      const domain = mail.slice(at + 1);
      return /(^|\.)deloitte[a-z0-9-]*\./.test(domain);
    };
    const memberFirmFiltered = all.filter(u => {
      const mail = (u.email || '').toLowerCase();
      return includeInternational ? isDeloitteDomain(mail) : mail.endsWith('@deloitte.de');
    });
    all.length = 0;
    all.push(...memberFirmFiltered);

    // Location + JobTitle per User Profile nachladen.
    // v20.0 (Audit): parallel statt sequentiell — vorher bis zu N serielle
    // Profil-Roundtrips pro Tipp-Suche (spürbare Picker-Latenz). Reine
    // Lese-Calls auf max. ~10 Treffer, daher unkritisch fürs Throttling.
    await Promise.all(all.map(async user => {
      try {
        const profile = await this.searchUserByEmail(user.email);
        if (profile) {
          if (profile.location) user.location = profile.location;
          if (profile.jobTitle) user.jobTitle = profile.jobTitle;
        }
      } catch { /* ignore */ }
    }));

    return all;
  }

  /**
   * v8.9: Alle User eines Standorts via Microsoft Graph laden. Genutzt vom
   * Exclude-Modal, damit der Organizer auch bei nur-Standortfilter-Events
   * eine Personenliste zum Aushaken bekommt (statt alle Personen einzeln
   * über die Suche finden zu müssen).
   *
   * Pagination: Graph $top=999 ist Maximum pro Page. Folgt @odata.nextLink
   * um alle Personen einzusammeln. Hard-Cap bei 5000 damit die UI nicht
   * hängt bei sehr grossen Standorten.
   *
   * Match: officeLocation eq '<location>' (exakt). Fällt auf
   * startsWith zurück, wenn das nichts liefert (z.B. 'DE - Köln' vs
   * 'Koeln, Germany' Schreibweisen-Drift).
   */
  public async searchUsersByLocation(location: string): Promise<Array<{
    email: string;
    displayName: string;
    firstName: string;
    lastName: string;
    location: string;
    jobTitle: string;
  }>> {
    if (!location) return [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const collected: any[] = [];
    const HARD_CAP = 5000;
    const escaped = location.replace(/'/g, "''");

    // Schritt 0 (v8.10b): Schnellster Pfad — Standorte haben in Deloitte-DE
    // jeweils eine eigene Verteilergruppe nach dem Schema DE<STADT>@deloitte.com
    // (z.B. 'DE - Köln' -> DEKOELN@deloitte.com, 'DE - Düsseldorf' ->
    // DEDUESSELDORF@deloitte.com). Wir extrahieren die Stadt aus dem
    // Standort-String, normalisieren Umlaute zu ASCII-Substitutionen und
    // versuchen den Verteiler über getGroupMembers aufzulösen. Damit
    // sparen wir uns Graph-Permission-Overhead vollständig für den
    // typischen DE-Office-Fall. Andere Fälle (kein 'DE'-Prefix oder
    // unbekannte Stadt) fallen weiter unten auf Graph + SP-Search zurück.
    try {
      const cityRaw = location
        .replace(/^DE\s*[-–—]\s*/i, '') // 'DE - Köln' -> 'Köln'
        .trim();
      const cityNormalized = cityRaw
        .toLowerCase()
        .replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ß/g, 'ss')
        .replace(/[^a-z0-9]/g, ''); // alles ausser ASCII-Buchstaben/-Ziffern raus
      if (cityNormalized) {
        const distMail = `DE${cityNormalized.toUpperCase()}@deloitte.com`;
        try {
          const grp = await this.getGroupMembers(distMail);
          if (grp && grp.members && grp.members.length > 0) {
            for (const m of grp.members) {
              if (m.email) {
                collected.push({
                  mail: m.email,
                  userPrincipalName: m.email,
                  displayName: m.displayName,
                  givenName: m.firstName,
                  surname: m.lastName,
                  jobTitle: m.jobTitle,
                  officeLocation: m.location || location,
                });
              }
            }
          }
        } catch (e) { console.warn('[DEX] searchUsersByLocation Konvention-Verteiler failed:', distMail, e); }
      }
    } catch (e) { console.warn('[DEX] searchUsersByLocation City-Extract failed:', e); }

    // Schritt 1: Versuch über Microsoft Graph mit ConsistencyLevel:eventual
    // (Advanced Query). Funktioniert nur wenn die App User.Read.All oder
    // Directory.Read.All Permission hat. Überspringen wenn der Verteiler-
    // Pfad schon Treffer hatte.
    if (collected.length === 0) try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = this.context as any;
      if (ctx.msGraphClientFactory) {
        const client = await ctx.msGraphClientFactory.getClient('3');
        const select = 'id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,officeLocation,city';
        const fetchPage = async (filterExpr: string): Promise<void> => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          let resp: any = await client.api('/users')
            .header('ConsistencyLevel', 'eventual')
            .filter(filterExpr)
            .count(true)
            .select(select)
            .top(999)
            .get();
          while (resp) {
            if (resp.value) collected.push(...resp.value);
            if (collected.length >= HARD_CAP) break;
            const next = resp['@odata.nextLink'];
            if (!next) break;
            resp = await client.api(next).get();
          }
        };
        try { await fetchPage(`officeLocation eq '${escaped}'`); }
        catch (e) { console.warn('[DEX] searchUsersByLocation Graph exact-match failed:', e); }
        if (collected.length === 0) {
          try { await fetchPage(`startsWith(officeLocation, '${escaped}')`); }
          catch (e) { console.warn('[DEX] searchUsersByLocation Graph startsWith failed:', e); }
        }
      }
    } catch (err) {
      console.warn('[DEX] searchUsersByLocation Graph block failed:', err);
    }

    // Schritt 2: Fallback über SharePoint Search People Index. Greift wenn
    // Graph keine Treffer hatte (z.B. wegen fehlender User.Read.All-Permission).
    // SourceId b09a7990-05ea-4af9-81ef-edfab16c4e31 = People-Search-Result-Source.
    if (collected.length === 0) {
      try {
        const refinementFilter = encodeURIComponent(`OfficeNumber:equals("${location}")`);
        const queryUrl = `${this.siteUrl}/_api/search/query?querytext='*'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'&refinementfilters='${refinementFilter}'&rowlimit=500&selectproperties='WorkEmail,PreferredName,FirstName,LastName,JobTitle,BaseOfficeLocation,Office'`;
        const resp = await this._sp.get(queryUrl, SPHttpClient.configurations.v1);
        if (resp.ok) {
          const data = await resp.json();
          const rows = data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows
            || data?.d?.query?.PrimaryQueryResult?.RelevantResults?.Table?.Rows?.results
            || [];
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          for (const row of rows) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const cells = row.Cells?.results || row.Cells || [];
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const find = (key: string): string => (cells.find((c: any) => c.Key === key) || {}).Value || '';
            const email = find('WorkEmail');
            if (!email) continue;
            collected.push({
              displayName: find('PreferredName'),
              givenName: find('FirstName'),
              surname: find('LastName'),
              mail: email,
              userPrincipalName: email,
              jobTitle: find('JobTitle'),
              officeLocation: find('BaseOfficeLocation') || find('Office') || location,
            });
          }
        } else {
          console.warn('[DEX] searchUsersByLocation SP-Search fallback HTTP', resp.status);
        }
      } catch (err) {
        console.warn('[DEX] searchUsersByLocation SP-Search fallback failed:', err);
      }
    }

    if (collected.length === 0) {
      console.warn('[DEX] searchUsersByLocation: keine User gefunden für Standort', location, '— Graph User.Read.All-Permission und SP-Search-People-Index prüfen.');
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return collected.map((u: any) => ({
      email: u.mail || u.userPrincipalName || '',
      displayName: u.displayName || '',
      firstName: u.givenName || '',
      lastName: u.surname || '',
      jobTitle: u.jobTitle || '',
      location: u.officeLocation || u.city || '',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    })).filter((u: any) => u.email);
  }

  /**
   * Verteilerlisten + Security-Groups aus Entra suchen via ClientPeoplePicker.
   * PrincipalType 6 = 2 (Distribution List) + 4 (Security Group).
   * Liefert Email (= mail des Verteilers) und DisplayName.
   */
  public async searchGroups(query: string): Promise<Array<{ email: string; displayName: string }>> {
    if (!query || query.length < 2) return [];
    try {
      const body = {
        'queryParams': {
          '__metadata': { 'type': 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
          'AllowEmailAddresses': true,
          'AllowMultipleEntities': false,
          'MaximumEntitySuggestions': 12,
          'QueryString': query,
          'PrincipalType': 6, // 2|4 = Distribution List + Security Group
          'PrincipalSource': 15,
          'SharePointGroupID': 0,
        },
      };
      const response = await this._post(
        `${this.siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
        body
      );
      if (!response.ok) return [];
      const data = await response.json();
      const resultsStr = data.d?.ClientPeoplePickerSearchUser || data.ClientPeoplePickerSearchUser || '[]';
      const results = JSON.parse(resultsStr);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return results
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .filter((r: any) => r.EntityData?.Email || r.Key)
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .map((r: any) => ({
          email: r.EntityData?.Email || r.Key || '',
          displayName: r.DisplayText || r.EntityData?.Title || r.EntityData?.AccountName || '',
        }))
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .filter((g: any) => g.email);
    } catch {
      return [];
    }
  }

  /**
   * Mitglieder einer Entra-Gruppe via Microsoft Graph laden.
   * Benötigt Group.Read.All Berechtigung im SharePoint App Catalog (admin-consent).
   * Nutzt MSGraphClientV3 aus dem WebPartContext.
   */
  public async getGroupMembers(groupEmail: string): Promise<{ groupName: string; members: Array<{ email: string; displayName: string; firstName?: string; lastName?: string; jobTitle?: string; location?: string }> } | null> {
    if (!groupEmail) return null;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = this.context as any;
      if (!ctx.msGraphClientFactory) return null;
      const client = await ctx.msGraphClientFactory.getClient('3');
      // 1. Gruppe anhand der Mail-Adresse finden
      const escaped = groupEmail.replace(/'/g, "''");
      const groupResp = await client.api(`/groups`)
        .filter(`mail eq '${escaped}' or proxyAddresses/any(p:p eq 'smtp:${escaped}')`)
        .select('id,displayName')
        .top(1)
        .get();
      const groups = groupResp?.value || [];
      if (groups.length === 0) return { groupName: groupEmail, members: [] };
      const group = groups[0];
      // 2. Transitive Members (inkl. verschachtelte Gruppen) holen — v8.8:
      // mit zusätzlichen Profil-Feldern. v8.12: PAGINATION über
      // @odata.nextLink, sonst stoppt Graph nach 999 Einträgen (oder
      // 200 wie vor v8.12) — bei Standort-Verteilern wie DEKOELN sind
      // das schnell 1000+ Personen.
      const HARD_CAP = 5000;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const collected: any[] = [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let resp: any = await client.api(`/groups/${group.id}/transitiveMembers/microsoft.graph.user`)
        .select('id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,officeLocation,city')
        .top(999)
        .get();
      while (resp) {
        if (resp.value) collected.push(...resp.value);
        if (collected.length >= HARD_CAP) break;
        const next = resp['@odata.nextLink'];
        if (!next) break;
        resp = await client.api(next).get();
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const members = collected.map((u: any) => ({
        email: u.mail || u.userPrincipalName || '',
        displayName: u.displayName || '',
        firstName: u.givenName || '',
        lastName: u.surname || '',
        jobTitle: u.jobTitle || '',
        // officeLocation ist häufig 'DE - Koeln'-Format; city ist Fallback
        location: u.officeLocation || u.city || '',
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      })).filter((m: any) => m.email);
      return { groupName: group.displayName || groupEmail, members };
    } catch (err) {
      console.warn('[DEX] getGroupMembers failed:', err);
      return null;
    }
  }

  /**
   * v30.61: Personal-Stammdaten mehrerer Personen auf einmal aus dem
   * Verzeichnis (Microsoft Graph).
   *
   * Anlass: Die Personalnummer wurde in v30.60 von Hand nachgetragen — pro
   * Person ein Klick in die Backoffice-Liste, ablesen, abtippen. Bei 100
   * Teilnehmern ist das eine Stunde stumpfe Arbeit mit Tippfehlerrisiko an
   * genau der Stelle, an der eine falsche Nummer auf der Rechnung landet.
   * Entra ID kennt diese Felder bereits:
   *
   *  - `employeeId`                    → Personalnummer
   *  - `employeeOrgData.costCenter`    → Kostenstelle des Mitarbeiters
   *  - `companyName`                   → Company Name (F&A-Spalte)
   *  - `country`                       → Country (F&A-Spalte)
   *
   * **Zwei Dinge, die man wissen muss, bevor man sich darauf verlässt:**
   *
   * (1) `employeeId` und `employeeOrgData` liefert Graph NUR mit der
   *     Berechtigung `User.Read.All`. Sie steht in `package-solution.json`,
   *     muss aber im SharePoint Admin Center unter „API-Zugriff" von einem
   *     Tenant-Admin freigegeben werden. Ohne Freigabe antwortet Graph mit
   *     403 — dann liefert diese Methode eine leere Map, und die Oberfläche
   *     sagt das auch. Sie tut NICHT so, als gäbe es die Daten nicht.
   * (2) Ob die Felder im Tenant überhaupt gepflegt sind, entscheidet HR, nicht
   *     DEX. Kommt für eine Person nichts zurück, bleibt der manuelle Weg über
   *     die Backoffice-Liste — deshalb bleibt der „Nachschlagen"-Knopf stehen.
   *
   * Abgefragt wird per JSON-Batch (20 Anfragen pro Aufruf, Graph-Limit) statt
   * einzeln: 100 Personen sind damit fünf Roundtrips statt hundert.
   */
  public async getEmployeeData(emails: string[]): Promise<Record<string, {
    employeeId?: string;
    costCenter?: string;
    companyName?: string;
    country?: string;
    department?: string;
  }>> {
    const out: Record<string, { employeeId?: string; costCenter?: string; companyName?: string; country?: string; department?: string }> = {};
    const list = Array.from(new Set((emails || []).map(e => (e || '').trim().toLowerCase()).filter(Boolean)));
    if (list.length === 0) return out;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = this.context as any;
      if (!ctx.msGraphClientFactory) return out;
      const client = await ctx.msGraphClientFactory.getClient('3');
      const SELECT = 'mail,userPrincipalName,employeeId,employeeOrgData,companyName,country,department';
      for (let i = 0; i < list.length; i += 20) {
        const chunk = list.slice(i, i + 20);
        const requests = chunk.map((mail, n) => ({
          id: String(n),
          method: 'GET',
          // Der Abruf über /users/{upn} trifft nur, wenn die Adresse der UPN
          // ist. Die Teilnehmerlisten tragen aber auch SMTP-Aliase (siehe die
          // Doppel-Adressen-Falle in CLAUDE.md) — deshalb der Filter über
          // mail UND userPrincipalName statt eines direkten Zugriffs.
          url: `/users?$select=${SELECT}&$filter=mail eq '${mail.replace(/'/g, "''")}' or userPrincipalName eq '${mail.replace(/'/g, "''")}'&$top=1`,
        }));
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const resp: any = await client.api('/$batch').post({ requests });
        for (const r of (resp?.responses || [])) {
          const idx = parseInt(r.id, 10);
          const mail = chunk[idx];
          if (!mail || r.status !== 200) continue;
          const u = (r.body && r.body.value && r.body.value[0]) || null;
          if (!u) continue;
          out[mail] = {
            employeeId: (u.employeeId || '').trim(),
            costCenter: ((u.employeeOrgData && u.employeeOrgData.costCenter) || '').trim(),
            companyName: (u.companyName || '').trim(),
            country: (u.country || '').trim(),
            department: (u.department || '').trim(),
          };
        }
      }
    } catch (err) {
      // Bewusst kein Wegwerfen: Der Aufrufer unterscheidet „nichts gefunden"
      // von „gar nicht gefragt" an der Größe der Map — deshalb hier melden.
      console.warn('[DEX] getEmployeeData (Graph) fehlgeschlagen:', err);
    }
    return out;
  }

  /**
   * Hilfsmethode für POST-Requests
   */
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
}
