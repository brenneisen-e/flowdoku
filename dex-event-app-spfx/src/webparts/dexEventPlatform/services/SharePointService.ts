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
  public async ensureRolesList(): Promise<void> {
    const listName = 'DEX_Roles';
    const exists = await this.listExists(listName);

    if (exists) {
      // Berechtigungen pruefen und ggf. nachtraeglich setzen
      await this.ensureRolesListPermissions(listName);
      return;
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
        'Choices': { 'results': ['SuperAdmin', 'EventAdmin', 'User'] },
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
   * Vererbung aufheben, dann nur die Site-Owners-Gruppe mit Vollzugriff hinzufuegen.
   */
  private async setRolesListPermissions(listName: string): Promise<void> {
    try {
      // 1. Berechtigungsvererbung aufheben (eigene Berechtigungen)
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );

      // 2. Site-Owners-Gruppe ID ermitteln
      const ownersResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (ownersResponse.ok) {
        const ownersData = await ownersResponse.json();
        const ownersGroupId = ownersData.Id;

        // 3. Full Control RoleDefinition ID ermitteln (1073741829 = Full Control)
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${ownersGroupId}, roledefid=1073741829)`,
          {}
        );
      }
    } catch {
      // Berechtigungen konnten nicht gesetzt werden - Liste ist trotzdem nutzbar
    }
  }

  /**
   * Alle Rollen-Eintraege lesen
   */
  public async getRoles(): Promise<Array<{
    Id: number;
    Title: string; // UserEmail
    UserName: string;
    Role: string;
    UserLocation: string;
    AssignedBy: string;
    AssignedDate: string;
  }>> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$select=Id,Title,UserName,Role,UserLocation,AssignedBy,AssignedDate&$orderby=Role,UserName`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) return [];

      const data = await response.json();
      return data.value || [];
    } catch {
      return [];
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
   * Hilfsmethode fuer POST-Requests
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
