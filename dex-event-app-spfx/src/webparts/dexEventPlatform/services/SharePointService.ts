/**
 * SharePoint Service - zentrale Klasse fuer alle SP REST API Aufrufe
 *
 * Erstellt Listen automatisch, liest/schreibt Eintraege.
 * Nutzt den SPFx-Context fuer authentifizierte Aufrufe.
 *
 * - Eike, Maerz 2026
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

export class SharePointService {
  private context: WebPartContext;
  private siteUrl: string;

  constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  /**
   * Pruefen ob eine Liste existiert
   */
  public async listExists(listName: string): Promise<boolean> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
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
      // Berechtigungen pruefen und ggf. nachtraeglich setzen
      await this.ensureRolesListPermissions(listName);

      // Choice-Werte des Role-Feldes auf neue Benennung migrieren
      try {
        await this.context.spHttpClient.post(
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
              'Choices': { 'results': ['Admin', 'Organizer', 'User'] },
            }),
          }
        );
      } catch { /* Choice-Update optional */ }

      return { isNewlyCreated: false };
    }

    // Liste erstellen
    const listPayload = {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Rollenverwaltung fuer die DEX Event Experience Platform',
      'BaseTemplate': 100, // Generic List
      'AllowContentTypes': false,
    };

    await this._post(
      `${this.siteUrl}/_api/web/lists`,
      listPayload
    );

    // Spalten hinzufuegen
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
    ];

    for (const col of columns) {
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
        col
      );
    }

    // Default View aktualisieren: alle Spalten anzeigen
    try {
      const viewFields = ['Title', 'UserName', 'Role', 'UserLocation', 'AssignedBy', 'AssignedDate'];
      for (const field of viewFields) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/views/getbytitle('All Items')/viewfields/addviewfield('${field}')`,
          {}
        );
      }
    } catch {
      // View-Update ist optional
    }

    // Berechtigungen setzen: nur Site-Owners (SuperAdmins) duerfen die Liste sehen
    await this.setRolesListPermissions(listName);
    return { isNewlyCreated: true };
  }

  /**
   * Pruefen ob die Liste bereits eigene Berechtigungen hat, sonst setzen
   */
  private async ensureRolesListPermissions(listName: string): Promise<void> {
    try {
      const response = await this.context.spHttpClient.get(
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
   * Eigene Berechtigungen fuer die Rollen-Liste setzen.
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
      const ownersResponse = await this.context.spHttpClient.get(
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
   * Einem User Leseberechtigung auf die Rollen-Liste geben (fuer EventAdmins).
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

      // Read = 1073741826 (Standard SharePoint ID, sprachunabhaengig)
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

      // 2. Full Control auf Site-Ebene (fuer Subsite-Erstellung)
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
   * Einem User Contribute-Rechte auf die DEX_Events-Liste geben (fuer Organizer).
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
   */
  public async revokeAccessOnEventsList(userEmail: string): Promise<void> {
    try {
      const userId = await this.getUserIdByEmail(userEmail);
      if (!userId) return;

      const headers: HeadersInit = {
        'Accept': 'application/json;odata=verbose',
        'X-HTTP-Method': 'DELETE',
      };

      await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/roleassignments/getbyprincipalid(${userId})`,
        SPHttpClient.configurations.v1,
        { headers }
      );
    } catch (e) {
      console.warn('[DEX] revokeAccessOnEventsList Error:', e);
    }
  }

  /**
   * Einem User die Berechtigung auf die Rollen-Liste entziehen.
   */
  public async revokeAccessOnRolesList(userEmail: string): Promise<void> {
    try {
      const userId = await this.getUserIdByEmail(userEmail);
      if (!userId) return;

      const headers: HeadersInit = {
        'Accept': 'application/json;odata=verbose',
        'X-HTTP-Method': 'DELETE',
      };

      await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/roleassignments/getbyprincipalid(${userId})`,
        SPHttpClient.configurations.v1,
        { headers }
      );
    } catch (e) {
      console.warn('[DEX] revokeAccessOnRolesList Error:', e);
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
      const response = await this.context.spHttpClient.get(
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
      const response = await this.context.spHttpClient.get(
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
      const response = await this.context.spHttpClient.get(
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
   * Alle Rollen-Eintraege lesen
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
  }> | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$select=Id,Title,UserName,Role,UserLocation,AssignedBy,AssignedDate&$orderby=Role,UserName`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) return null;

      const data = await response.json();
      return data.value || [];
    } catch {
      return null;
    }
  }

  /**
   * Rolle eines bestimmten Users abfragen (per E-Mail)
   */
  public async getUserRole(email: string): Promise<string | null> {
    try {
      const response = await this.context.spHttpClient.get(
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

      const response = await this.context.spHttpClient.post(
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
      const response = await this.context.spHttpClient.post(
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

      const response = await this.context.spHttpClient.post(
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
   * Gibt Name, Standort und JobTitle zurueck falls gefunden.
   *
   * Strategie:
   *   1. Direkter Claim-Lookup mit `i:0#.f|membership|<email>` (schnell).
   *   2. Wenn leer oder kein DisplayName: per `siteusers/getbyemail` den echten
   *      LoginName aufloesen (UPN-Claim), dann GetPropertiesFor damit erneut
   *      aufrufen. Deckt UPN != SMTP, Guest-Accounts, Alias-SMTP-Adressen ab.
   */
  public async searchUserByEmail(email: string): Promise<{
    displayName: string;
    location: string;
    jobTitle: string;
  } | null> {
    if (!email) return null;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const extract = (data: any): { displayName: string; location: string; jobTitle: string } | null => {
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
      };
    };

    // 1) Direkter Lookup per SMTP
    try {
      const directResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|${encodeURIComponent(email)}'&$select=DisplayName,UserProfileProperties`,
        SPHttpClient.configurations.v1
      );
      if (directResp.ok) {
        const data = await directResp.json();
        const hit = extract(data);
        if (hit && (hit.jobTitle || hit.location)) return hit;
        // DisplayName ohne Properties? Merken fuer Fallback-Default.
        if (hit) {
          // weiter zum LoginName-Pfad - vielleicht bringt der jobTitle
        }
      }
    } catch { /* weiter */ }

    // 2) Fallback: echten LoginName (UPN-Claim) via siteusers/getbyemail
    try {
      const siteUserResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/siteusers/getbyemail('${email.replace(/'/g, "''")}')?$select=LoginName,Title`,
        SPHttpClient.configurations.v1
      );
      if (siteUserResp.ok) {
        const su = await siteUserResp.json();
        const loginName: string = su.LoginName || su.d?.LoginName || '';
        const fallbackDisplayName: string = su.Title || su.d?.Title || '';
        if (loginName) {
          try {
            const profileResp = await this.context.spHttpClient.get(
              `${this.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(loginName)}'&$select=DisplayName,UserProfileProperties`,
              SPHttpClient.configurations.v1
            );
            if (profileResp.ok) {
              const pData = await profileResp.json();
              const hit = extract(pData);
              if (hit) return hit;
            }
          } catch { /* */ }
        }
        if (fallbackDisplayName) {
          return { displayName: fallbackDisplayName, location: '', jobTitle: '' };
        }
      }
    } catch { /* User nicht gefunden */ }
    return null;
  }

  /**
   * User-Suche mit Autocomplete: sucht nach Name oder Email-Fragment.
   * Nutzt die SharePoint ClientPeoplePickerSearchUser API.
   */
  public async searchUsers(query: string): Promise<Array<{
    email: string;
    displayName: string;
    location: string;
    jobTitle: string;
  }>> {
    if (!query || query.length < 2) return [];

    try {
      const body = {
        'queryParams': {
          '__metadata': { 'type': 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
          'AllowEmailAddresses': true,
          'AllowMultipleEntities': false,
          'MaximumEntitySuggestions': 10,
          'QueryString': query,
          'PrincipalType': 1, // Users only
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
      const mapped: Array<{ email: string; displayName: string; location: string; jobTitle: string }> = results
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .filter((r: any) => r.EntityData?.Email)
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .map((r: any) => ({
          email: r.EntityData.Email || '',
          displayName: r.DisplayText || r.EntityData.Title || '',
          location: '', // Wird per User Profile nachgeladen
          jobTitle: r.EntityData.Title || '', // EntityData.Title ist bei SharePoint manchmal schon JobTitle;
                                               // Wird unten durch UserProfile-Lookup ueberschrieben.
        }));

      // Location + JobTitle per User Profile nachladen
      for (const user of mapped) {
        try {
          const profile = await this.searchUserByEmail(user.email);
          if (profile) {
            if (profile.location) user.location = profile.location;
            if (profile.jobTitle) user.jobTitle = profile.jobTitle;
          }
        } catch { /* ignore */ }
      }

      return mapped;
    } catch {
      return [];
    }
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
   * Benoetigt Group.Read.All Berechtigung im SharePoint App Catalog (admin-consent).
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
      // mit zusaetzlichen Profil-Feldern (Vor-/Nachname, JobTitle, Office-
      // Location), damit das Exclude-Modal eine sortierbare/filterbare
      // Tabelle zeigen kann ohne pro Member einen weiteren Graph-Call.
      const membersResp = await client.api(`/groups/${group.id}/transitiveMembers/microsoft.graph.user`)
        .select('id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,officeLocation,city')
        .top(200)
        .get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const members = (membersResp?.value || []).map((u: any) => ({
        email: u.mail || u.userPrincipalName || '',
        displayName: u.displayName || '',
        firstName: u.givenName || '',
        lastName: u.surname || '',
        jobTitle: u.jobTitle || '',
        // officeLocation ist haeufig 'DE - Koeln'-Format; city ist Fallback
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

    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  }
}
