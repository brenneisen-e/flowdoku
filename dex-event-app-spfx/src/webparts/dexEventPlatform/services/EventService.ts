/**
 * Event Service - SharePoint-Operationen fuer Events und Teilnehmerlisten
 *
 * Erstellt DEX_Events-Liste automatisch beim ersten Start.
 * Erstellt pro Event eine Subsite mit einer "Teilnehmer"-Liste.
 *
 * Struktur auf SharePoint:
 *   DOL-c-DE-B2Run (Hauptsite)
 *   ├── DEX_Events (zentrale Event-Liste)
 *   ├── DEX_Roles (Rollenverwaltung)
 *   ├── b2run-frankfurt-2026 (Subsite)
 *   │   └── Teilnehmer (Registrierungsliste)
 *   ├── jpmorgan-muenchen-2026 (Subsite)
 *   │   └── Teilnehmer (Registrierungsliste)
 *   └── ...
 *
 * - Eike, Maerz 2026
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

// Fester Listenname auf jeder Subsite
const REG_LIST_NAME = 'Teilnehmer';
const REG_LIST_ITEM_TYPE = 'SP.Data.TeilnehmerListItem';

export interface SPEvent {
  Id: number;
  Title: string;
  EventStatus: string;
  EventType: string;
  Description: string;
  Location: string;
  LocationFilter: string;
  Audience: string;
  StartDate: string;
  EndDate: string;
  RegistrationDeadline: string;
  LastDeregisterDate: string;
  MaxParticipants: number;
  WaitlistEnabled: boolean;
  EventImageUrl: string;
  Organizer: string;
  OrganizerEmail: string;
  OutlookEventId: string;
  CustomFields: string; // JSON-String mit konfigurierbaren Feldern
  RegistrationListName: string;
  SubsiteUrl: string; // Absolute URL der Event-Subsite
}

export interface CustomField {
  id: string;
  label: string;
  type: 'text' | 'select' | 'number' | 'checkbox';
  required: boolean;
  options?: string[]; // fuer select-Felder
  visible: boolean;
}

export interface SPRegistration {
  Id: number;
  Title: string; // Teilnehmer-ID
  ParticipantName: string;
  ParticipantEmail: string;
  Status: string;
  RegistrationDate: string;
  CancellationDate: string;
  CustomData: string; // JSON mit Custom Field Werten
}

export class EventService {
  private context: WebPartContext;
  private siteUrl: string;

  constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  // ==================== DEX_Events Liste ====================

  /**
   * Events-Liste erstellen falls nicht vorhanden
   */
  public async ensureEventsList(): Promise<void> {
    const listName = 'DEX_Events';
    const exists = await this.listExists(listName);
    if (exists) {
      await this.ensureMissingFields(listName);

      await this.configureDefaultView(listName, [
        'EventStatus', 'EventType', 'Location', 'LocationFilter',
        'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
        'WaitlistEnabled', 'Organizer', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
      ]);
      try {
        const listInfo = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=HasUniqueRoleAssignments`,
          SPHttpClient.configurations.v1
        );
        if (listInfo.ok) {
          const data = await listInfo.json();
          if (!data.HasUniqueRoleAssignments) {
            await this.setEventsListPermissions(listName);
          }
        }
      } catch { /* ignore */ }
      return;
    }

    // Liste erstellen
    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Events der DEX Event Experience Platform',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    // Spalten hinzufuegen
    const fields = this.getEventsFieldDefinitions();
    for (const f of fields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if (f.choices) {
        payload['Choices'] = { 'results': f.choices };
      }
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    }

    await this.configureDefaultView(listName, [
      'EventStatus', 'EventType', 'Location', 'LocationFilter',
      'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
      'WaitlistEnabled', 'Organizer', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
    ]);

    await this.setEventsListPermissions(listName);
  }

  /**
   * Feld-Definitionen fuer DEX_Events Liste
   */
  private getEventsFieldDefinitions(): Array<{ title: string; type: number; choices?: string[]; metaType?: string }> {
    return [
      { title: 'EventStatus', type: 6, choices: ['Under Construction', 'Active', 'Completed', 'Cancelled'], metaType: 'SP.FieldChoice' },
      { title: 'EventType', type: 6, choices: ['B2Run', 'JPMorgan', 'Other'], metaType: 'SP.FieldChoice' },
      { title: 'Description', type: 3 },
      { title: 'Location', type: 2 },
      { title: 'LocationFilter', type: 2 },
      { title: 'Audience', type: 2 },
      { title: 'StartDate', type: 4 },
      { title: 'EndDate', type: 4 },
      { title: 'RegistrationDeadline', type: 4 },
      { title: 'LastDeregisterDate', type: 4 },
      { title: 'MaxParticipants', type: 9 },
      { title: 'WaitlistEnabled', type: 8 },
      { title: 'EventImageUrl', type: 2 },
      { title: 'Organizer', type: 2 },
      { title: 'OrganizerEmail', type: 2 },
      { title: 'OutlookEventId', type: 2 },
      { title: 'CustomFields', type: 3 },
      { title: 'RegistrationListName', type: 2 },
      { title: 'RegistrationListUrl', type: 2 },
      { title: 'SubsiteUrl', type: 2 },
    ];
  }

  /**
   * Fehlende Spalten auf einer bestehenden DEX_Events-Liste nachtraeglich hinzufuegen.
   */
  private async ensureMissingFields(listName: string): Promise<void> {
    const requiredFields = this.getEventsFieldDefinitions();

    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName&$filter=Hidden eq false&$top=200`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return;

      const data = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const existingFields = new Set((data.value || []).map((f: any) => f.InternalName));

      for (const f of requiredFields) {
        if (!existingFields.has(f.title)) {
          console.log('[DEX] Fehlende Spalte nachtraeglich hinzufuegen:', f.title);
          const payload: Record<string, unknown> = {
            '__metadata': { 'type': f.metaType || 'SP.Field' },
            'Title': f.title,
            'FieldTypeKind': f.type,
            'Required': false,
          };
          if (f.choices) {
            payload['Choices'] = { 'results': f.choices };
          }
          try {
            await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
          } catch {
            console.warn('[DEX] Konnte Spalte nicht hinzufuegen:', f.title);
          }
        }
      }
    } catch (e) {
      console.warn('[DEX] ensureMissingFields Error:', e);
    }
  }

  /**
   * Berechtigungen fuer DEX_Events setzen
   */
  private async setEventsListPermissions(listName: string): Promise<void> {
    try {
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );

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

      const membersResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedmembergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (membersResponse.ok) {
        const membersData = await membersResponse.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${membersData.Id}, roledefid=1073741826)`,
          {}
        );
      }
    } catch {
      // Berechtigungen konnten nicht gesetzt werden
    }
  }

  /**
   * Default View einer Liste konfigurieren
   */
  private async configureDefaultView(listName: string, fieldNames: string[], baseUrl?: string): Promise<void> {
    const url = baseUrl || this.siteUrl;
    try {
      for (const fieldName of fieldNames) {
        await this._post(
          `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/addviewfield('${fieldName}')`,
          {}
        );
      }
    } catch {
      // View-Konfiguration ist optional
    }
  }

  // ==================== Events CRUD ====================

  private static readonly EVENT_SELECT = 'Id,Title,EventStatus,EventType,Description,Location,LocationFilter,Audience,StartDate,EndDate,RegistrationDeadline,LastDeregisterDate,MaxParticipants,WaitlistEnabled,EventImageUrl,Organizer,OrganizerEmail,OutlookEventId,CustomFields,RegistrationListName,SubsiteUrl';

  /**
   * Alle Events laden
   */
  public async getEvents(): Promise<SPEvent[]> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EventService.EVENT_SELECT}&$orderby=StartDate desc&$top=100`,
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
   * Einzelnes Event laden
   */
  public async getEvent(eventId: number): Promise<SPEvent | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=${EventService.EVENT_SELECT}`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      return await response.json();
    } catch {
      return null;
    }
  }

  /**
   * Neues Event erstellen + Subsite mit Teilnehmerliste anlegen
   */
  public async createEvent(event: {
    title: string;
    status: string;
    type: string;
    description: string;
    location: string;
    locationFilter: string;
    audience: string;
    startDate: string;
    endDate: string;
    registrationDeadline: string;
    lastDeregisterDate: string;
    maxParticipants: number;
    waitlistEnabled: boolean;
    eventImageUrl: string;
    organizer: string;
    organizerEmail: string;
    outlookEventId: string;
    customFields: CustomField[];
  }): Promise<number | null> {
    try {
      // 1. Subsite fuer das Event erstellen
      const subsiteUrl = await this.createEventSubsite(event.title, event.description);
      if (!subsiteUrl) {
        console.error('[DEX] Subsite konnte nicht erstellt werden');
        return null;
      }

      // 2. Teilnehmerliste auf der Subsite erstellen
      await this.createRegistrationList(subsiteUrl, event.customFields, event.organizerEmail);

      // 3. Event in DEX_Events eintragen
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
        'Title': event.title,
        'EventStatus': event.status,
        'EventType': event.type,
        'Description': event.description,
        'Location': event.location,
        'LocationFilter': event.locationFilter,
        'Audience': event.audience,
        'StartDate': event.startDate || null,
        'EndDate': event.endDate || null,
        'RegistrationDeadline': event.registrationDeadline || null,
        'LastDeregisterDate': event.lastDeregisterDate || null,
        'MaxParticipants': event.maxParticipants,
        'WaitlistEnabled': event.waitlistEnabled,
        'EventImageUrl': event.eventImageUrl,
        'Organizer': event.organizer,
        'OrganizerEmail': event.organizerEmail,
        'OutlookEventId': event.outlookEventId,
        'CustomFields': JSON.stringify(event.customFields),
        'RegistrationListName': REG_LIST_NAME,
        'RegistrationListUrl': `${subsiteUrl}/Lists/${REG_LIST_NAME}/AllItems.aspx`,
        'SubsiteUrl': subsiteUrl,
      };

      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`,
        payload
      );

      if (!response.ok) return null;
      const result = await response.json();
      return result.d?.Id || result.Id;
    } catch {
      return null;
    }
  }

  /**
   * Event aktualisieren
   */
  public async updateEvent(eventId: number, updates: Record<string, unknown>): Promise<boolean> {
    try {
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
        ...updates,
      };

      const response = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify(payload),
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Event vollstaendig loeschen:
   * 1. Subsite loeschen (inkl. Teilnehmerliste) - fuer neue Events
   * 2. Alte Registrierungsliste loeschen (DEX_Reg_*) - fuer alte Events
   * 3. Event-Eintrag aus DEX_Events loeschen
   */
  public async deleteEvent(eventId: number): Promise<boolean> {
    try {
      // Event-Daten laden um SubsiteUrl und RegistrationListName zu bekommen
      const event = await this.getEvent(eventId);
      if (!event) return false;

      // 1. Subsite loeschen (neue Events)
      if (event.SubsiteUrl) {
        try {
          await this.context.spHttpClient.post(
            `${event.SubsiteUrl}/_api/web`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'X-HTTP-Method': 'DELETE',
              },
            }
          );
        } catch {
          console.warn('[DEX] Subsite konnte nicht geloescht werden:', event.SubsiteUrl);
        }
      }

      // 2. Alte Registrierungsliste loeschen (alte Events ohne Subsite)
      if (event.RegistrationListName && event.RegistrationListName !== 'Teilnehmer') {
        try {
          await this.context.spHttpClient.post(
            `${this.siteUrl}/_api/web/lists/getbytitle('${event.RegistrationListName.replace(/'/g, "''")}')`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'DELETE',
              },
            }
          );
        } catch {
          console.warn('[DEX] Alte Registrierungsliste konnte nicht geloescht werden:', event.RegistrationListName);
        }
      }

      // 3. Event-Eintrag aus DEX_Events loeschen
      const response = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE',
          },
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== Subsites ====================

  /**
   * URL-Suffix aus Event-Titel generieren.
   * "B2Run Frankfurt 2026" → "b2run-frankfurt-2026-k8f3a"
   */
  private generateSubsiteUrl(title: string): string {
    const slug = title
      .toLowerCase()
      .replace(/[äÄ]/g, 'ae').replace(/[öÖ]/g, 'oe').replace(/[üÜ]/g, 'ue').replace(/ß/g, 'ss')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .substring(0, 40);
    const suffix = Date.now().toString(36).slice(-5);
    return `${slug}-${suffix}`;
  }

  /**
   * Subsite fuer ein Event erstellen.
   * Gibt die absolute URL der neuen Subsite zurueck.
   */
  private async createEventSubsite(title: string, description: string): Promise<string | null> {
    try {
      const urlSuffix = this.generateSubsiteUrl(title);

      const payload = {
        'parameters': {
          '__metadata': { 'type': 'SP.WebCreationInformation' },
          'Title': title,
          'Url': urlSuffix,
          'Description': description || `Event-Subsite: ${title}`,
          'Language': 1031, // Deutsch
          'WebTemplate': 'STS#3', // Modern Team Site ohne Microsoft 365 Group
          'UseUniquePermissions': true,
        }
      };

      const response = await this._post(`${this.siteUrl}/_api/web/webs/add`, payload);
      if (!response.ok) {
        console.error('[DEX] Subsite-Erstellung fehlgeschlagen:', response.status);
        return null;
      }

      const result = await response.json();
      // Die API gibt die absolute URL im Url-Feld zurueck
      const subsiteAbsoluteUrl = result.d?.Url || result.Url;
      if (subsiteAbsoluteUrl) return subsiteAbsoluteUrl;

      // Fallback: URL manuell zusammenbauen
      return `${this.siteUrl}/${urlSuffix}`;
    } catch (e) {
      console.error('[DEX] Subsite-Erstellung Fehler:', e);
      return null;
    }
  }

  // ==================== Teilnehmerlisten (auf Subsites) ====================

  /**
   * Teilnehmerliste auf einer Subsite erstellen.
   * Liste heisst immer "Teilnehmer".
   */
  private async createRegistrationList(
    subsiteUrl: string,
    customFields: CustomField[],
    organizerEmail: string
  ): Promise<void> {
    // Liste erstellen
    await this._post(`${subsiteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': REG_LIST_NAME,
      'Description': 'Teilnehmerliste fuer dieses Event',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    // Basis-Spalten
    const baseFields = [
      { title: 'ParticipantName', type: 2 },
      { title: 'ParticipantEmail', type: 2 },
      { title: 'Status', type: 6, choices: ['Angemeldet', 'Warteliste', 'Eingecheckt', 'Abgemeldet'], metaType: 'SP.FieldChoice' },
      { title: 'RegistrationDate', type: 4 },
      { title: 'CancellationDate', type: 4 },
      { title: 'CustomData', type: 3 },
    ];

    for (const f of baseFields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if ((f as { choices?: string[] }).choices) {
        payload['Choices'] = { 'results': (f as { choices: string[] }).choices };
      }
      await this._post(`${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, payload);
    }

    // Default View konfigurieren
    await this.configureDefaultView(REG_LIST_NAME, [
      'ParticipantName', 'ParticipantEmail', 'Status', 'RegistrationDate', 'CancellationDate',
    ], subsiteUrl);

    // Item-Level Permissions
    await this.setItemLevelPermissions(subsiteUrl);

    // Berechtigungen
    await this.setRegistrationListPermissions(subsiteUrl, organizerEmail);
  }

  /**
   * Item-Level Permissions auf der Teilnehmerliste setzen.
   */
  private async setItemLevelPermissions(subsiteUrl: string): Promise<void> {
    try {
      await this.context.spHttpClient.post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.List' },
            'ReadSecurity': 2,
            'WriteSecurity': 2,
          }),
        }
      );
    } catch {
      // Item-Level Permissions konnten nicht gesetzt werden
    }
  }

  /**
   * Berechtigungen fuer Teilnehmerliste auf der Subsite setzen.
   */
  private async setRegistrationListPermissions(subsiteUrl: string, organizerEmail: string): Promise<void> {
    try {
      await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );

      // Site Owners der Hauptsite: Full Control
      const ownersResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (ownersResponse.ok) {
        const ownersData = await ownersResponse.json();
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${ownersData.Id}, roledefid=1073741829)`,
          {}
        );
      }

      // Site Members: Contribute (damit sie sich registrieren koennen)
      const membersResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedmembergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (membersResponse.ok) {
        const membersData = await membersResponse.json();
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${membersData.Id}, roledefid=1073741827)`,
          {}
        );
      }

      // Organizer: Full Control
      if (organizerEmail) {
        try {
          const userResponse = await this.context.spHttpClient.get(
            `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(organizerEmail)}')?$select=Id`,
            SPHttpClient.configurations.v1
          );
          if (userResponse.ok) {
            const userData = await userResponse.json();
            await this._post(
              `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${userData.Id}, roledefid=1073741829)`,
              {}
            );
          }
        } catch { /* Organizer-Berechtigung optional */ }
      }
    } catch {
      // Berechtigungen konnten nicht gesetzt werden
    }
  }

  // ==================== Registrierungen ====================

  /**
   * Registrierung fuer ein Event erstellen.
   * Operiert auf der Subsite des Events.
   */
  public async registerForEvent(
    subsiteUrl: string,
    participantName: string,
    participantEmail: string,
    customData: Record<string, string>,
    status: string = 'Angemeldet'
  ): Promise<boolean> {
    try {
      const payload = {
        '__metadata': { 'type': REG_LIST_ITEM_TYPE },
        'Title': participantEmail,
        'ParticipantName': participantName,
        'ParticipantEmail': participantEmail,
        'Status': status,
        'RegistrationDate': new Date().toISOString(),
        'CustomData': JSON.stringify(customData),
      };

      const response = await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`,
        payload
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Bestehende abgemeldete Registrierung reaktivieren.
   * Setzt Status zurueck auf Angemeldet/Warteliste, loescht CancellationDate,
   * aktualisiert RegistrationDate und CustomData.
   */
  public async reactivateRegistration(
    subsiteUrl: string,
    itemId: number,
    participantName: string,
    customData: Record<string, string>,
    status: string = 'Angemeldet'
  ): Promise<boolean> {
    try {
      const response = await this.context.spHttpClient.post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': REG_LIST_ITEM_TYPE },
            'ParticipantName': participantName,
            'Status': status,
            'RegistrationDate': new Date().toISOString(),
            'CancellationDate': null,
            'CustomData': JSON.stringify(customData),
          }),
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Eigene Registrierung fuer ein Event laden
   */
  public async getMyRegistration(
    subsiteUrl: string,
    email: string
  ): Promise<SPRegistration | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=ParticipantEmail eq '${email.replace(/'/g, "''")}'&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      const data = await response.json();
      if (data.value && data.value.length > 0) return data.value[0];
      return null;
    } catch {
      return null;
    }
  }

  /**
   * Alle Registrierungen fuer ein Event laden (nur fuer Organizer/Admin)
   */
  public async getAllRegistrations(subsiteUrl: string): Promise<SPRegistration[]> {
    try {
      const response = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,Title,ParticipantName,ParticipantEmail,Status,RegistrationDate,CancellationDate,CustomData&$orderby=RegistrationDate&$top=500`,
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
   * Registrierung stornieren
   */
  public async cancelRegistration(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const response = await this.context.spHttpClient.post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': REG_LIST_ITEM_TYPE },
            'Status': 'Abgemeldet',
            'CancellationDate': new Date().toISOString(),
          }),
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Aktuelle Teilnehmeranzahl ermitteln
   */
  public async getRegistrationCount(subsiteUrl: string): Promise<{ registered: number; waitlist: number }> {
    try {
      const response = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Status&$top=500`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return { registered: 0, waitlist: 0 };
      const data = await response.json();
      const items = data.value || [];
      const registered = items.filter((i: { Status: string }) => i.Status === 'Angemeldet' || i.Status === 'Eingecheckt').length;
      const waitlist = items.filter((i: { Status: string }) => i.Status === 'Warteliste').length;
      return { registered, waitlist };
    } catch {
      return { registered: 0, waitlist: 0 };
    }
  }

  /**
   * Title-Feld (= Teilnehmer-ID) aktualisieren
   */
  public async updateRegistrationTitle(
    subsiteUrl: string,
    itemId: number,
    newTitle: string
  ): Promise<boolean> {
    try {
      const response = await this.context.spHttpClient.post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': REG_LIST_ITEM_TYPE },
            'Title': newTitle,
          }),
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== Hilfsmethoden ====================

  private async listExists(listName: string): Promise<boolean> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')`,
        SPHttpClient.configurations.v1
      );
      return response.ok;
    } catch {
      return false;
    }
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
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  }
}
