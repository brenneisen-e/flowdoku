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
  Audience: string; // Zielgruppen-Filter (Gruppen + Emails, kommasepariert)
  FilterMode: string; // 'AND' | 'OR' - Verknüpfung Standort+Zielgruppe
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
  CalendarLink: string;
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

  // ==================== DEX_Emails Liste ====================

  /**
   * E-Mail-Queue-Liste erstellen falls nicht vorhanden.
   * Power Automate reagiert auf neue Eintraege und versendet Mails.
   *
   * Spalten:
   * - Title: Betreff der E-Mail
   * - Recipient: Empfänger E-Mail-Adresse
   * - RecipientName: Name des Empfängers
   * - Body: HTML-Inhalt der E-Mail
   * - EmailType: Art der E-Mail (Anmeldung, Abmeldung, Warteliste, Nachrücken, Info)
   * - EventTitle: Name des Events
   * - EventId: ID des Events (Referenz)
   * - Status: Pending, Sent, Failed
   * - SentDate: Wann wurde die Mail versendet
   */
  public async ensureEmailsList(): Promise<void> {
    const listName = 'DEX_Emails';
    const exists = await this.listExists(listName);
    if (exists) {
      // Berechtigungen pruefen und ggf. nachtraeglich setzen
      try {
        const listInfo = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=HasUniqueRoleAssignments`,
          SPHttpClient.configurations.v1
        );
        if (listInfo.ok) {
          const data = await listInfo.json();
          if (!data.HasUniqueRoleAssignments) {
            await this.setEmailsListPermissions(listName);
          }
        }
      } catch { /* ignore */ }
      return;
    }

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'E-Mail-Queue für automatischen Versand via Power Automate',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
      { title: 'Recipient', type: 2 },
      { title: 'RecipientName', type: 2 },
      { title: 'Body', type: 3 }, // Note (multiline/HTML)
      { title: 'EmailType', type: 6, choices: ['Anmeldung', 'Abmeldung', 'Warteliste', 'Nachruecken', 'Info'], metaType: 'SP.FieldChoice' },
      { title: 'EventTitle', type: 2 },
      { title: 'EventId', type: 2 },
      { title: 'Status', type: 6, choices: ['Pending', 'Sent', 'Failed'], metaType: 'SP.FieldChoice' },
      { title: 'SentDate', type: 4 },
    ];

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

    // Default View
    await this.configureDefaultView(listName, [
      'Recipient', 'RecipientName', 'EmailType', 'EventTitle', 'Status', 'SentDate',
    ]);

    await this.setEmailsListPermissions(listName);
  }

  /**
   * Berechtigungen fuer DEX_Emails: Owners Full Control, Members Contribute, Item-Level Security
   */
  private async setEmailsListPermissions(listName: string): Promise<void> {
    try {
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );
      const ownersResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`, SPHttpClient.configurations.v1
      );
      if (ownersResp.ok) {
        const d = await ownersResp.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741829)`, {}
        );
      }
      const membersResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedmembergroup?$select=Id`, SPHttpClient.configurations.v1
      );
      if (membersResp.ok) {
        const d = await membersResp.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741827)`, {}
        );
      }
    } catch { /* */ }

    // Item-Level Security
    try {
      await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE',
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.List' },
            'ReadSecurity': 2, 'WriteSecurity': 2,
          }),
        }
      );
    } catch { /* */ }
  }

  /**
   * E-Mail in die Queue eintragen (wird von Power Automate versendet).
   */
  public async queueEmail(
    subject: string,
    recipient: string,
    recipientName: string,
    body: string,
    emailType: string,
    eventTitle: string,
    eventId: string
  ): Promise<boolean> {
    try {
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Emails')/items`,
        {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailsListItem' },
          'Title': subject,
          'Recipient': recipient,
          'RecipientName': recipientName,
          'Body': body,
          'EmailType': emailType,
          'EventTitle': eventTitle,
          'EventId': eventId,
          'Status': 'Pending',
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== DEX_Outlook Liste ====================

  /**
   * Outlook-Termin-Queue-Liste erstellen falls nicht vorhanden.
   * Power Automate reagiert auf neue Eintraege und erstellt Outlook-Termine.
   *
   * Spalten:
   * - Title: Betreff des Termins
   * - Attendee: E-Mail-Adresse des Teilnehmers
   * - AttendeeName: Name des Teilnehmers
   * - EventTitle: Name des Events
   * - EventId: ID des Events (Referenz)
   * - StartDate: Beginn des Termins
   * - EndDate: Ende des Termins
   * - Location: Veranstaltungsort
   * - Body: Beschreibung / HTML-Inhalt
   * - ActionType: Erstellen, Absagen (fuer Abmeldung)
   * - Status: Pending, Sent, Failed
   * - SentDate: Wann wurde der Termin versendet
   */
  public async ensureOutlookList(): Promise<void> {
    const listName = 'DEX_Outlook';
    const exists = await this.listExists(listName);
    if (exists) return;

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Outlook-Termin-Queue für automatischen Versand via Power Automate',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
      { title: 'Attendee', type: 2 },
      { title: 'AttendeeName', type: 2 },
      { title: 'EventTitle', type: 2 },
      { title: 'EventId', type: 2 },
      { title: 'StartDate', type: 4 },
      { title: 'EndDate', type: 4 },
      { title: 'Location', type: 2 },
      { title: 'Body', type: 3 },
      { title: 'ActionType', type: 6, choices: ['Erstellen', 'Absagen'], metaType: 'SP.FieldChoice' },
      { title: 'Status', type: 6, choices: ['Pending', 'Sent', 'Failed'], metaType: 'SP.FieldChoice' },
      { title: 'SentDate', type: 4 },
    ];

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
      'Attendee', 'AttendeeName', 'ActionType', 'EventTitle', 'StartDate', 'EndDate', 'Status', 'SentDate',
    ]);
  }

  /**
   * Outlook-Termin in die Queue eintragen (wird von Power Automate versendet).
   */
  public async queueOutlookEvent(
    subject: string,
    attendee: string,
    attendeeName: string,
    eventTitle: string,
    eventId: string,
    startDate: string,
    endDate: string,
    location: string,
    body: string,
    actionType: 'Erstellen' | 'Absagen'
  ): Promise<boolean> {
    try {
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`,
        {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_OutlookListItem' },
          'Title': subject,
          'Attendee': attendee,
          'AttendeeName': attendeeName,
          'EventTitle': eventTitle,
          'EventId': eventId,
          'StartDate': startDate,
          'EndDate': endDate,
          'Location': location,
          'Body': body,
          'ActionType': actionType,
          'Status': 'Pending',
        }
      );
      return response.ok;
    } catch {
      return false;
    }
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
        'WaitlistEnabled', 'Organizer', 'EventImageUrl', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
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
      { title: 'FilterMode', type: 6, choices: ['AND', 'OR'], metaType: 'SP.FieldChoice' },
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
      { title: 'CalendarLink', type: 2 },
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

  private static readonly EVENT_SELECT = 'Id,Title,EventStatus,EventType,Description,Location,LocationFilter,Audience,FilterMode,StartDate,EndDate,RegistrationDeadline,LastDeregisterDate,MaxParticipants,WaitlistEnabled,EventImageUrl,Organizer,OrganizerEmail,OutlookEventId,CalendarLink,CustomFields,RegistrationListName,SubsiteUrl';

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
    filterMode: string;
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

      // FieldMap aus createRegistrationList auslesen
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const fieldMap: Record<string, string> = (this as any)._lastFieldMap || {};

      // Custom Fields mit SP InternalName anreichern
      const enrichedCustomFields = event.customFields.map(cf => ({
        ...cf,
        spInternalName: fieldMap[cf.id] || '',
      }));

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
        'FilterMode': event.filterMode || 'OR',
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
        'CustomFields': JSON.stringify(enrichedCustomFields),
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
          await this._delete(`${event.SubsiteUrl}/_api/web`);
        } catch {
          console.warn('[DEX] Subsite konnte nicht geloescht werden:', event.SubsiteUrl);
        }
      }

      // 2. Event-Bild loeschen (wenn in DEX_EventImages)
      if (event.EventImageUrl) {
        try {
          const url = new URL(event.EventImageUrl);
          const serverRelUrl = url.pathname;
          if (serverRelUrl.indexOf('DEX_EventImages') >= 0) {
            await this._delete(
              `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelUrl}')`
            );
          }
        } catch {
          console.warn('[DEX] Event-Bild konnte nicht geloescht werden');
        }
      }

      // 3. Alte Registrierungsliste loeschen (alte Events ohne Subsite) (alte Events ohne Subsite)
      if (event.RegistrationListName && event.RegistrationListName !== 'Teilnehmer') {
        try {
          await this._delete(
            `${this.siteUrl}/_api/web/lists/getbytitle('${event.RegistrationListName.replace(/'/g, "''")}')`
          );
        } catch {
          console.warn('[DEX] Alte Registrierungsliste konnte nicht geloescht werden:', event.RegistrationListName);
        }
      }

      // 3. Event-Eintrag aus DEX_Events loeschen
      const response = await this._delete(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`
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
   * Versucht mehrere Templates falls eines fehlschlaegt.
   * Gibt die absolute URL der neuen Subsite zurueck.
   */
  private async createEventSubsite(title: string, description: string): Promise<string | null> {
    const urlSuffix = this.generateSubsiteUrl(title);
    const desc = description || `Event-Subsite: ${title}`;

    // Templates in Reihenfolge versuchen:
    // STS#3 = Modern ohne Group, STS#0 = Classic Team Site, STS = Blank
    const templates = ['STS#3', 'STS#0', 'STS'];

    for (const template of templates) {
      try {
        const payload = {
          'parameters': {
            '__metadata': { 'type': 'SP.WebCreationInformation' },
            'Title': title,
            'Url': urlSuffix,
            'Description': desc,
            'Language': 1031,
            'WebTemplate': template,
            'UseSamePermissionsAsParentSite': false,
          }
        };

        const response = await this._post(`${this.siteUrl}/_api/web/webs/add`, payload);
        if (response.ok) {
          const result = await response.json();
          const subsiteAbsoluteUrl = result.d?.Url || result.Url;
          console.log(`[DEX] Subsite erstellt mit Template ${template}:`, subsiteAbsoluteUrl);
          return subsiteAbsoluteUrl || `${this.siteUrl}/${urlSuffix}`;
        }

        // Fehlerdetails loggen
        try {
          const err = await response.json();
          console.warn(`[DEX] Template ${template} fehlgeschlagen (${response.status}):`, err.error?.message?.value || err);
        } catch {
          console.warn(`[DEX] Template ${template} fehlgeschlagen: ${response.status}`);
        }
      } catch (e) {
        console.warn(`[DEX] Template ${template} Fehler:`, e);
      }
    }

    console.error('[DEX] Subsite konnte mit keinem Template erstellt werden');
    return null;
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
      { title: 'TeilnehmerID', type: 9 }, // Number - fortlaufende ID
      { title: 'ParticipantName', type: 2 },
      { title: 'ParticipantEmail', type: 2 },
      { title: 'Department', type: 2 },
      { title: 'Location', type: 2 },
      { title: 'JobTitle', type: 2 },
      { title: 'Phone', type: 2 },
      { title: 'Status', type: 6, choices: ['Angemeldet', 'Warteliste', 'Eingecheckt', 'Abgemeldet'], metaType: 'SP.FieldChoice' },
      { title: 'RegistrationDate', type: 4 },
      { title: 'LastModifiedDate', type: 4 },
      { title: 'ChangeLog', type: 3 }, // Note (multiline) - Aenderungshistorie
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

    // Custom Fields als eigene Spalten anlegen + InternalName merken
    const customFieldViewNames: string[] = [];
    const fieldMap: Record<string, string> = {}; // cf.id -> SP InternalName
    for (const cf of customFields) {
      if (!cf.label) continue;
      let fieldPayload: Record<string, unknown>;

      if (cf.type === 'select' && cf.options && cf.options.length > 0) {
        fieldPayload = {
          '__metadata': { 'type': 'SP.FieldChoice' },
          'Title': cf.label,
          'FieldTypeKind': 6,
          'Required': false,
          'Choices': { 'results': cf.options },
        };
      } else if (cf.type === 'number') {
        fieldPayload = {
          '__metadata': { 'type': 'SP.Field' },
          'Title': cf.label,
          'FieldTypeKind': 9,
          'Required': false,
        };
      } else if (cf.type === 'checkbox') {
        fieldPayload = {
          '__metadata': { 'type': 'SP.Field' },
          'Title': cf.label,
          'FieldTypeKind': 8,
          'Required': false,
        };
      } else {
        fieldPayload = {
          '__metadata': { 'type': 'SP.Field' },
          'Title': cf.label,
          'FieldTypeKind': 2,
          'Required': false,
        };
      }

      try {
        const fieldResponse = await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`,
          fieldPayload
        );
        if (fieldResponse.ok) {
          const fieldResult = await fieldResponse.json();
          const internalName = fieldResult.d?.InternalName || fieldResult.InternalName || cf.label;
          fieldMap[cf.id] = internalName;
          customFieldViewNames.push(internalName);
        }
      } catch {
        console.warn('[DEX] Custom Field konnte nicht angelegt werden:', cf.label);
      }
    }

    // FieldMap auf der Subsite speichern (als Property Bag oder in einem Hidden Field)
    // Wir speichern die Map spaeter im CustomFields JSON des Events
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (this as any)._lastFieldMap = fieldMap;

    // Default View konfigurieren (Basis + Custom Fields)
    await this.configureDefaultView(REG_LIST_NAME, [
      'TeilnehmerID', 'ParticipantName', 'ParticipantEmail', 'Department', 'Location', 'JobTitle', 'Phone', 'Status', 'RegistrationDate', 'CancellationDate',
      ...customFieldViewNames,
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
    status: string = 'Angemeldet',
    customFieldMap?: Record<string, string> // cf.id -> SP InternalName
  ): Promise<boolean> {
    try {
      // Naechste TeilnehmerID ermitteln
      let nextId = 1;
      try {
        const counts = await this.getRegistrationCount(subsiteUrl);
        nextId = counts.registered + counts.waitlist + 1;
      } catch { /* Fallback: 1 */ }

      // Profildaten laden
      const profile = await this.getCurrentUserProfile();

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': REG_LIST_ITEM_TYPE },
        'Title': participantEmail,
        'TeilnehmerID': nextId,
        'ParticipantName': participantName,
        'ParticipantEmail': participantEmail,
        'Department': profile.department,
        'Location': profile.location,
        'JobTitle': profile.jobTitle,
        'Phone': profile.phone,
        'Status': status,
        'RegistrationDate': new Date().toISOString(),
        'CustomData': JSON.stringify(customData),
      };

      // Custom Field Werte in die echten SP-Spalten schreiben
      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName && customData[cfId]) {
            payload[spFieldName] = customData[cfId];
          }
        }
      }

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
    status: string = 'Angemeldet',
    customFieldMap?: Record<string, string>
  ): Promise<boolean> {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = {
        'ParticipantName': participantName,
        'Status': status,
        'RegistrationDate': new Date().toISOString(),
        'CancellationDate': null,
        'CustomData': JSON.stringify(customData),
      };

      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName && customData[cfId]) {
            body[spFieldName] = customData[cfId];
          }
        }
      }

      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Custom Data einer Registrierung aktualisieren (Teilnehmer aendert eigene Angaben).
   */
  public async updateRegistrationData(
    subsiteUrl: string,
    itemId: number,
    customData: Record<string, string>,
    customFieldMap?: Record<string, string>,
    oldCustomData?: Record<string, string>,
    fieldLabelMap?: Record<string, string> // cf.id -> label
  ): Promise<boolean> {
    try {
      // Änderungen ermitteln
      const changes: string[] = [];
      const now = new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
      if (oldCustomData && fieldLabelMap) {
        for (const key of Object.keys(customData)) {
          if (key === 'salutation') continue;
          const label = fieldLabelMap[key] || key;
          const oldVal = oldCustomData[key] || '';
          const newVal = customData[key] || '';
          if (oldVal !== newVal) {
            changes.push(`${label}: "${oldVal}" → "${newVal}"`);
          }
        }
      }
      const changeEntry = changes.length > 0 ? `[${now}] ${changes.join(', ')}` : '';

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = {
        'CustomData': JSON.stringify(customData),
        'LastModifiedDate': new Date().toISOString(),
      };

      // ChangeLog anhängen (bestehenden Log behalten)
      if (changeEntry) {
        try {
          const existing = await this.context.spHttpClient.get(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog`,
            SPHttpClient.configurations.v1
          );
          if (existing.ok) {
            const data = await existing.json();
            const oldLog = data.ChangeLog || '';
            body['ChangeLog'] = oldLog ? `${changeEntry}\n${oldLog}` : changeEntry;
          }
        } catch {
          body['ChangeLog'] = changeEntry;
        }
      }

      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName && customData[cfId]) {
            body[spFieldName] = customData[cfId];
          }
        }
      }

      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
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
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        {
          'Status': 'Abgemeldet',
          'CancellationDate': new Date().toISOString(),
          'TeilnehmerID': null,
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
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'Title': newTitle }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== Bild-Upload ====================

  /**
   * Event-Bild in die SiteAssets-Bibliothek hochladen.
   * Gibt die absolute URL des hochgeladenen Bildes zurueck.
   */
  /**
   * Sicherstellen dass der DEX_EventImages Ordner in SiteAssets existiert.
   */
  /**
   * SiteAssets-Unterordner sicherstellen:
   * - DEX_EventImages (Event-Bilder)
   * - DEX_Logos (Deloitte-Logo fuer E-Mail-Templates, manuell hochgeladen)
   */
  public async ensureAssetsFolders(): Promise<void> {
    const baseUrl = this.context.pageContext.web.serverRelativeUrl;
    const folders = ['DEX_EventImages', 'DEX_Logos'];

    for (const folder of folders) {
      const folderUrl = `${baseUrl}/SiteAssets/${folder}`;
      try {
        const check = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')`,
          SPHttpClient.configurations.v1
        );
        if (check.ok) continue;
      } catch { /* */ }

      try {
        await this._post(`${this.siteUrl}/_api/web/folders`, {
          '__metadata': { 'type': 'SP.Folder' },
          'ServerRelativeUrl': folderUrl,
        });
        console.log(`[DEX] Ordner erstellt: SiteAssets/${folder}`);
      } catch {
        console.warn(`[DEX] Konnte ${folder} Ordner nicht erstellen`);
      }
    }
  }

  public async uploadEventImage(file: File, eventTitle: string): Promise<string | null> {
    try {
      // Ordner sicherstellen
      await this.ensureAssetsFolders();

      const ext = file.name.split('.').pop() || 'jpg';
      const safeName = eventTitle
        .replace(/[^a-zA-Z0-9]/g, '_')
        .substring(0, 30) + '_' + Date.now().toString(36) + '.' + ext;

      const folderUrl = `${this.context.pageContext.web.serverRelativeUrl}/SiteAssets/DEX_EventImages`;
      const uploadUrl = `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/Files/add(url='${safeName}',overwrite=true)`;

      // Request Digest holen fuer nativen fetch
      const digestResp = await fetch(`${this.siteUrl}/_api/contextinfo`, {
        method: 'POST',
        headers: { 'Accept': 'application/json;odata=verbose' },
        credentials: 'same-origin',
      });
      const digestData = await digestResp.json();
      const requestDigest = digestData.d?.GetContextWebInformation?.FormDigestValue || '';

      // Nativen fetch nutzen (SPHttpClient ueberschreibt Content-Type)
      const arrayBuffer = await file.arrayBuffer();
      const response = await fetch(uploadUrl, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'X-RequestDigest': requestDigest,
        },
        body: arrayBuffer,
        credentials: 'same-origin',
      });

      if (response.ok) {
        const result = await response.json();
        const serverRelUrl = result.d?.ServerRelativeUrl || result.ServerRelativeUrl;
        if (serverRelUrl) {
          return `${window.location.origin}${serverRelUrl}`;
        }
      }
      return null;
    } catch (e) {
      console.error('[DEX] Bild-Upload fehlgeschlagen:', e);
      return null;
    }
  }

  // ==================== Profil-Daten ====================

  /**
   * Profildaten des aktuellen Users laden fuer die Teilnehmerliste.
   */
  public async getCurrentUserProfile(): Promise<{ department: string; location: string; jobTitle: string; phone: string }> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return { department: '', location: '', jobTitle: '', phone: '' };

      const data = await response.json();
      const props: Array<{ Key: string; Value: string }> = data.UserProfileProperties || [];
      const get = (key: string): string => {
        const p = props.find(x => x.Key === key);
        return p && p.Value ? p.Value : '';
      };

      return {
        department: get('Department'),
        location: get('Office'),
        jobTitle: get('Title'),
        phone: get('WorkPhone') || get('CellPhone'),
      };
    } catch {
      return { department: '', location: '', jobTitle: '', phone: '' };
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

  /**
   * MERGE-Request ohne __metadata (fuer Subsite-Listen wo der ListItemType unbekannt ist)
   */
  private async _merge(url: string, body: object): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(body),
    };
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  }

  private async _delete(url: string): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE',
      },
    };
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  }
}
