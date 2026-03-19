/**
 * Event Service - SharePoint-Operationen fuer Events und Teilnehmerlisten
 *
 * Erstellt DEX_Events-Liste automatisch beim ersten Start.
 * Erstellt pro Event eine separate Teilnehmerliste mit Item-Level Permissions.
 *
 * - Eike, Maerz 2026
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

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
  Title: string; // Email
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
      // Bestehende Liste: Default View sicherstellen
      await this.configureDefaultView(listName, [
        'EventStatus', 'EventType', 'Location', 'LocationFilter',
        'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
        'WaitlistEnabled', 'Organizer',
      ]);
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
    const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
      { title: 'EventStatus', type: 6, choices: ['Under Construction', 'Active', 'Completed', 'Cancelled'], metaType: 'SP.FieldChoice' },
      { title: 'EventType', type: 6, choices: ['B2Run', 'JPMorgan', 'Other'], metaType: 'SP.FieldChoice' },
      { title: 'Description', type: 3 }, // Note (multiline)
      { title: 'Location', type: 2 },
      { title: 'LocationFilter', type: 2 },
      { title: 'Audience', type: 2 },
      { title: 'StartDate', type: 4 },
      { title: 'EndDate', type: 4 },
      { title: 'RegistrationDeadline', type: 4 },
      { title: 'LastDeregisterDate', type: 4 },
      { title: 'MaxParticipants', type: 9 }, // Number
      { title: 'WaitlistEnabled', type: 8 }, // Boolean
      { title: 'EventImageUrl', type: 2 },
      { title: 'Organizer', type: 2 },
      { title: 'OrganizerEmail', type: 2 },
      { title: 'OutlookEventId', type: 2 },
      { title: 'CustomFields', type: 3 }, // Note (JSON)
      { title: 'RegistrationListName', type: 2 },
    ];

    for (const f of fields) {
      const payload: Record<string, unknown> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if (f.choices) {
        (payload as Record<string, unknown>)['Choices'] = { 'results': f.choices };
      }
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    }

    // Default View konfigurieren: wichtige Spalten einblenden
    await this.configureDefaultView(listName, [
      'EventStatus', 'EventType', 'Location', 'LocationFilter',
      'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
      'WaitlistEnabled', 'Organizer', 'RegistrationListName',
    ]);
  }

  /**
   * Default View einer Liste konfigurieren: Spalten hinzufuegen
   */
  private async configureDefaultView(listName: string, fieldNames: string[]): Promise<void> {
    try {
      for (const fieldName of fieldNames) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/addviewfield('${fieldName}')`,
          {}
        );
      }
    } catch {
      // View-Konfiguration ist optional
    }
  }

  /**
   * Alle Events laden
   */
  public async getEvents(): Promise<SPEvent[]> {
    try {
      const select = 'Id,Title,EventStatus,EventType,Description,Location,LocationFilter,Audience,StartDate,EndDate,RegistrationDeadline,LastDeregisterDate,MaxParticipants,WaitlistEnabled,EventImageUrl,Organizer,OrganizerEmail,OutlookEventId,CustomFields,RegistrationListName';
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${select}&$orderby=StartDate desc&$top=100`,
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
      const select = 'Id,Title,EventStatus,EventType,Description,Location,LocationFilter,Audience,StartDate,EndDate,RegistrationDeadline,LastDeregisterDate,MaxParticipants,WaitlistEnabled,EventImageUrl,Organizer,OrganizerEmail,OutlookEventId,CustomFields,RegistrationListName';
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=${select}`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return null;
      return await response.json();
    } catch {
      return null;
    }
  }

  /**
   * Neues Event erstellen + Teilnehmerliste anlegen
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
      // 1. Eindeutigen Listennamen generieren
      const listSuffix = Date.now().toString(36);
      const regListName = `DEX_Reg_${listSuffix}`;

      // 2. Event in DEX_Events eintragen
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
        'RegistrationListName': regListName,
      };

      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`,
        payload
      );

      if (!response.ok) return null;
      const result = await response.json();
      const eventId = result.d?.Id || result.Id;

      // 3. Teilnehmerliste erstellen
      await this.createRegistrationList(regListName, event.customFields, event.organizerEmail);

      return eventId;
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

  // ==================== Teilnehmerlisten ====================

  /**
   * Teilnehmerliste fuer ein Event erstellen.
   * Basis-Spalten + dynamische Spalten aus den CustomFields.
   * Item-Level Permissions: jeder sieht nur seinen eigenen Eintrag.
   */
  private async createRegistrationList(
    listName: string,
    customFields: CustomField[],
    organizerEmail: string
  ): Promise<void> {
    // Liste erstellen
    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Teilnehmerliste',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    // Basis-Spalten
    const baseFields = [
      { title: 'ParticipantName', type: 2 },
      { title: 'ParticipantEmail', type: 2 },
      { title: 'Status', type: 6, choices: ['Registered', 'Waitlist', 'Checked-In', 'Cancelled'], metaType: 'SP.FieldChoice' },
      { title: 'RegistrationDate', type: 4 },
      { title: 'CancellationDate', type: 4 },
      { title: 'CustomData', type: 3 }, // JSON mit allen Custom-Field-Werten
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
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
    }

    // Item-Level Permissions setzen: jeder sieht nur seine eigenen Eintraege
    await this.setItemLevelPermissions(listName);

    // Berechtigungen: Vererbung aufheben, Organizer + Owners Vollzugriff, Members Contribute
    await this.setRegistrationListPermissions(listName, organizerEmail);
  }

  /**
   * Item-Level Permissions auf einer Liste setzen.
   * ReadSecurity=2: Nur eigene Eintraege lesen
   * WriteSecurity=2: Nur eigene Eintraege bearbeiten
   */
  private async setItemLevelPermissions(listName: string): Promise<void> {
    try {
      const response = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`,
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
            'ReadSecurity': 2, // Nur eigene Eintraege lesen
            'WriteSecurity': 2, // Nur eigene Eintraege bearbeiten
          }),
        }
      );
      if (!response.ok) {
        // Fallback ignorieren
      }
    } catch {
      // Item-Level Permissions konnten nicht gesetzt werden
    }
  }

  /**
   * Berechtigungen fuer Teilnehmerliste setzen.
   * - Vererbung aufheben
   * - Site Owners: Full Control
   * - Organizer: Full Control
   * - Site Members: Contribute (damit sie sich eintragen koennen)
   */
  private async setRegistrationListPermissions(listName: string, organizerEmail: string): Promise<void> {
    try {
      // Vererbung aufheben, bestehende Berechtigungen kopieren
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );

      // Site Owners: Full Control (1073741829)
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

      // Site Members: Contribute (1073741827)
      const membersResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedmembergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (membersResponse.ok) {
        const membersData = await membersResponse.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${membersData.Id}, roledefid=1073741827)`,
          {}
        );
      }

      // Organizer: Full Control (individuell)
      if (organizerEmail) {
        try {
          const userResponse = await this.context.spHttpClient.get(
            `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(organizerEmail)}')?$select=Id`,
            SPHttpClient.configurations.v1
          );
          if (userResponse.ok) {
            const userData = await userResponse.json();
            await this._post(
              `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${userData.Id}, roledefid=1073741829)`,
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
   * Registrierung fuer ein Event erstellen
   */
  public async registerForEvent(
    registrationListName: string,
    participantName: string,
    participantEmail: string,
    customData: Record<string, string>,
    status: string = 'Registered'
  ): Promise<boolean> {
    try {
      // Listennamen escaped fuer SP
      const escapedName = registrationListName.replace(/'/g, "''");
      const payload = {
        '__metadata': { 'type': `SP.Data.${registrationListName.replace(/_/g, '_x005f_')}ListItem` },
        'Title': participantEmail,
        'ParticipantName': participantName,
        'ParticipantEmail': participantEmail,
        'Status': status,
        'RegistrationDate': new Date().toISOString(),
        'CustomData': JSON.stringify(customData),
      };

      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${escapedName}')/items`,
        payload
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
    registrationListName: string,
    email: string
  ): Promise<SPRegistration | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${registrationListName}')/items?$filter=ParticipantEmail eq '${encodeURIComponent(email)}'&$top=1`,
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
  public async getAllRegistrations(registrationListName: string): Promise<SPRegistration[]> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${registrationListName}')/items?$select=Id,Title,ParticipantName,ParticipantEmail,Status,RegistrationDate,CancellationDate,CustomData&$orderby=RegistrationDate&$top=500`,
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
    registrationListName: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const response = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${registrationListName}')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': `SP.Data.${registrationListName.replace(/_/g, '_x005f_')}ListItem` },
            'Status': 'Cancelled',
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
   * Aktuelle Teilnehmeranzahl ermitteln (nur Registered + Checked-In)
   */
  public async getRegistrationCount(registrationListName: string): Promise<{ registered: number; waitlist: number }> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${registrationListName}')/items?$select=Status&$top=500`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return { registered: 0, waitlist: 0 };
      const data = await response.json();
      const items = data.value || [];
      const registered = items.filter((i: { Status: string }) => i.Status === 'Registered' || i.Status === 'Checked-In').length;
      const waitlist = items.filter((i: { Status: string }) => i.Status === 'Waitlist').length;
      return { registered, waitlist };
    } catch {
      return { registered: 0, waitlist: 0 };
    }
  }

  /**
   * Title-Feld (= Teilnehmer-ID) eines Registrierungseintrags aktualisieren
   */
  public async updateRegistrationTitle(
    registrationListName: string,
    itemId: number,
    newTitle: string
  ): Promise<boolean> {
    try {
      const response = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${registrationListName}')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': `SP.Data.${registrationListName.replace(/_/g, '_x005f_')}ListItem` },
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
