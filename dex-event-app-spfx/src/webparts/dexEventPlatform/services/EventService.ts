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
import { buildOutlookBody } from './EmailTemplates';

// Fester Listenname auf jeder Subsite
const REG_LIST_NAME = 'Teilnehmer';
const REG_LIST_ITEM_TYPE = 'SP.Data.TeilnehmerListItem';

export interface SPEvent {
  Id: number;
  Title: string;
  EventStatus: string;
  EventType: string;
  EventNumber: number;
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
  EmailImageBase64: string; // Base64 Event-Bild fuer E-Mails/Outlook
  Organizer: string;
  OrganizerEmail: string;
  OutlookEventId: string;
  CalendarLink: string;
  OutlookBody: string; // Text fuer den Outlook-Kalendereintrag
  EmailLanguage: string; // DE oder EN
  EmailTemplateOverrides: string; // JSON mit Event-spezifischen Template-Anpassungen
  CustomFields: string; // JSON-String mit konfigurierbaren Feldern
  Agenda: string; // JSON-Array mit Agenda-Eintraegen
  Transfers: string; // JSON-Array mit Transferzeiten
  Documents: string; // JSON-Array mit Dokumenten
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
  Title: string; // Email
  TeilnehmerID?: number;
  Vorname?: string;
  Nachname?: string;
  ParticipantName: string;
  ParticipantEmail: string;
  Status: string;
  RegistrationDate: string;
  CancellationDate: string;
  CustomData: string; // JSON mit Custom Field Werten
}

export interface SPParticipant {
  Id: number;
  Title: string; // Email
  Vorname: string;
  Nachname: string;
  Email: string;
  EventRegistered: string; // Kommaseparierte EventNumbers
  EventOnWaitlist: string; // Kommaseparierte EventNumbers
}

export class EventService {
  private context: WebPartContext;
  public siteUrl: string;

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
      const deallId = await this.getVisitorsGroupId();
      if (deallId) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${deallId}, roledefid=1073741827)`, {}
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
   * Berechtigungen fuer Queue-Listen (DEX_Outlook, DEX_IDReorder):
   * Owners Full Control, Members Contribute, Item-Level Security
   */
  private async setQueueListPermissions(listName: string): Promise<void> {
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
      const deallId = await this.getVisitorsGroupId();
      if (deallId) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${deallId}, roledefid=1073741827)`, {}
        );
      }
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
   * Power Automate reagiert auf neue Eintraege und laedt Teilnehmer
   * zum Outlook-Termin ein oder aus. Der Flow holt sich alle Event-Details
   * (Titel, Datum, Ort, CalendarLink) aus der DEX_Events-Liste via EventId.
   *
   * Spalten:
   * - Title: Kurzbeschreibung (z.B. "Einladung: B2Run")
   * - Attendee: E-Mail-Adresse des Teilnehmers
   * - EventId: ID des Events in DEX_Events (Referenz)
   * - ActionType: Einladen, Ausladen
   * - Status: Pending, Sent, Failed
   * - SentDate: Wann wurde die Aktion ausgefuehrt
   */
  public async ensureOutlookList(): Promise<void> {
    const listName = 'DEX_Outlook';
    const exists = await this.listExists(listName);
    if (exists) return;

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Outlook-Termin-Queue: Power Automate laedt Teilnehmer ein/aus',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
      { title: 'Attendee', type: 2 },
      { title: 'EventId', type: 2 },
      { title: 'ActionType', type: 6, choices: ['Einladen', 'Ausladen'], metaType: 'SP.FieldChoice' },
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
      'Attendee', 'EventId', 'ActionType', 'Status', 'SentDate',
    ]);

    await this.setQueueListPermissions(listName);
  }

  /**
   * Outlook-Termin-Einladung in die Queue eintragen.
   * Flow holt Event-Details (Datum, Ort, CalendarLink) aus DEX_Events via EventId.
   */
  public async queueOutlookEvent(
    attendee: string,
    eventId: string,
    eventTitle: string,
    actionType: 'Einladen' | 'Ausladen'
  ): Promise<boolean> {
    try {
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`,
        {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_OutlookListItem' },
          'Title': `${actionType}: ${eventTitle}`,
          'Attendee': attendee,
          'EventId': eventId,
          'ActionType': actionType,
          'Status': 'Pending',
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== DEX_IDReorder Queue ====================

  /**
   * Queue-Liste fuer ID-Neuvergabe erstellen falls nicht vorhanden.
   * Power Automate reagiert auf neue Eintraege und vergibt TeilnehmerIDs
   * auf der jeweiligen Subsite-Teilnehmerliste lueckenlos neu.
   *
   * Spalten:
   * - Title: Kurzbeschreibung (z.B. "Reorder: Test_20260408")
   * - EventId: SP Item-ID des Events in DEX_Events
   * - EventNumber: Hochlaufende EventNumber
   * - SubsiteUrl: Absolute URL der Event-Subsite
   * - Status: Pending, Processing, Done, Failed
   */
  public async ensureIDReorderList(): Promise<void> {
    const listName = 'DEX_IDReorder';
    const exists = await this.listExists(listName);
    if (exists) return;

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Queue fuer TeilnehmerID-Neuvergabe nach Abmeldungen',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields = [
      { title: 'EventId', type: 2 },
      { title: 'EventNumber', type: 9 },
      { title: 'SubsiteUrl', type: 2 },
      { title: 'Status', type: 6, choices: ['Pending', 'Processing', 'Done', 'Failed'], metaType: 'SP.FieldChoice' },
    ];

    for (const f of fields) {
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

    await this.configureDefaultView(listName, [
      'EventId', 'EventNumber', 'SubsiteUrl', 'Status',
    ]);

    await this.setQueueListPermissions(listName);
  }

  /**
   * ID-Reorder in Queue eintragen (nach Abmeldung).
   */
  public async queueIDReorder(
    eventId: string,
    eventNumber: number,
    subsiteUrl: string,
    eventTitle: string
  ): Promise<boolean> {
    try {
      // ListItemEntityTypeFullName dynamisch ermitteln
      let listItemType = 'SP.Data.DEX_x005f_IDReorderListItem';
      try {
        const typeResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')?$select=ListItemEntityTypeFullName`,
          SPHttpClient.configurations.v1
        );
        if (typeResp.ok) {
          const typeData = await typeResp.json();
          listItemType = typeData.ListItemEntityTypeFullName || typeData.d?.ListItemEntityTypeFullName || listItemType;
        }
      } catch { /* Fallback auf Standard-Name */ }

      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_IDReorder')/items`,
        {
          '__metadata': { 'type': listItemType },
          'Title': `Reorder: ${eventTitle}`,
          'EventId': eventId,
          'EventNumber': eventNumber,
          'SubsiteUrl': subsiteUrl,
          'Status': 'Pending',
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  // ==================== DEX_EmailTemplates Liste ====================

  /**
   * Email-Templates-Liste erstellen und Default-Templates einfuegen.
   * Templates koennen pro Event ueberschrieben werden (im Event JSON).
   *
   * Platzhalter: {{Name}}, {{EventTitle}}, {{AppUrl}}
   */
  public async ensureEmailTemplatesList(): Promise<void> {
    const listName = 'DEX_EmailTemplates';
    const exists = await this.listExists(listName);
    if (exists) {
      // Liste existiert - pruefen ob _Config Zeile und Logo-Spalten vorhanden
      await this.ensureEmailTemplatesConfig(listName);
      return;
    }

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Email-Vorlagen für die DEX Event Experience Platform (DE + EN)',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields = [
      { title: 'TemplateType', type: 2 },
      { title: 'Language', type: 2 },
      { title: 'Subject', type: 2 },
      { title: 'HeadingColor', type: 2 },
      { title: 'Heading', type: 2 },
      { title: 'BodyHtml', type: 3 },
      { title: 'LogoBase64', type: 3 },           // Base64 Deloitte Logo (Deloitte_Logo.png)
      { title: 'DefaultImageBase64', type: 3 },    // Base64 Default-Bild (dex-orb.png)
    ];

    for (const f of fields) {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      });
    }

    // Default-Templates: DE + EN fuer jeden Typ
    const defaults = [
      // ===== ENGLISCH =====
      { TemplateType: 'Anmeldung', Language: 'EN', Subject: 'Registration confirmation: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Registration successful',
        BodyHtml: '<p><strong>Dear {{Name}},</strong></p><p>you have successfully registered for the event <strong>{{EventTitle}}</strong>.</p><p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="{{AppUrl}}">Event Experience Platform</a> (\u201EMy Events\u201C) so the spot might be allocated to another participant.</p><p>For organizational questions about the event, please contact <strong>{{Organizer}}</strong>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Warteliste', Language: 'EN', Subject: 'Waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Waitlist confirmation',
        BodyHtml: '<p><strong>Dear {{Name}},</strong></p><p>you have been placed on the <strong>waitlist</strong> for the event <strong>{{EventTitle}}</strong>.</p><p>Your current position: <strong>#{{WaitlistPosition}}</strong></p><p>We will notify you as soon as a spot becomes available. You can always check your current position in the <a href="{{AppUrl}}">Event Experience Platform</a> under \u201EMy Events\u201C.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Abmeldung', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
        BodyHtml: '<p><strong>Dear {{Name}},</strong></p><p>your registration for the event <strong>{{EventTitle}}</strong> has been <strong>cancelled</strong>.</p><p>If you change your mind, you can register again via the <a href="{{AppUrl}}">Event Experience Platform</a>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Nachruecken', Language: 'EN', Subject: 'Spot available: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'You got a spot!',
        BodyHtml: '<p><strong>Dear {{Name}},</strong></p><p>Great news! A spot has become available and you have been <strong>moved from the waitlist to a confirmed participant</strong> for the event <strong>{{EventTitle}}</strong>.</p><p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="{{AppUrl}}">Event Experience Platform</a>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'EventErstellt', Language: 'EN', Subject: '[Deloitte Eventmanager] - New event created: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event Created',
        BodyHtml: '<p><strong>Dear {{Name}},</strong></p><p>your event <strong>{{EventTitle}}</strong> has been successfully created.</p><p>You can manage participants in the <a href="{{AppUrl}}">Event Experience Platform</a>.</p><p>Regards,<br>Team Event Experience Platform</p>' },
      // ===== DEUTSCH =====
      { TemplateType: 'Anmeldung', Language: 'DE', Subject: 'Anmeldebestätigung: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Anmeldung erfolgreich',
        BodyHtml: '<p><strong>Hallo {{Name}},</strong></p><p>du hast dich erfolgreich für das Event <strong>{{EventTitle}}</strong> angemeldet.</p><p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="{{AppUrl}}">Event Experience Platform</a> (\u201EMeine Events\u201C) ab, damit der Platz an einen anderen Teilnehmer vergeben werden kann.</p><p>Zu organisatorischen Fragen zum Event wende dich bitte an <strong>{{Organizer}}</strong>.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Warteliste', Language: 'DE', Subject: 'Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Warteliste-Bestätigung',
        BodyHtml: '<p><strong>Hallo {{Name}},</strong></p><p>du stehst auf der <strong>Warteliste</strong> für das Event <strong>{{EventTitle}}</strong>.</p><p>Deine aktuelle Position: <strong>#{{WaitlistPosition}}</strong></p><p>Wir benachrichtigen dich, sobald ein Platz frei wird. Deinen aktuellen Warteliste-Platz kannst du jederzeit in der <a href="{{AppUrl}}">Event Experience Platform</a> unter \u201EMeine Events\u201C sehen.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Abmeldung', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
        BodyHtml: '<p><strong>Hallo {{Name}},</strong></p><p>deine Anmeldung für das Event <strong>{{EventTitle}}</strong> wurde <strong>storniert</strong>.</p><p>Du kannst dich jederzeit erneut über die <a href="{{AppUrl}}">Event Experience Platform</a> anmelden.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Nachruecken', Language: 'DE', Subject: 'Platz frei: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Du bist nachgerückt!',
        BodyHtml: '<p><strong>Hallo {{Name}},</strong></p><p>Gute Nachrichten! Ein Platz ist frei geworden und du bist von der Warteliste <strong>als Teilnehmer bestätigt</strong> für das Event <strong>{{EventTitle}}</strong>.</p><p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="{{AppUrl}}">Event Experience Platform</a> ab.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'EventErstellt', Language: 'DE', Subject: '[Deloitte Eventmanager] - Neues Event erstellt: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event erstellt',
        BodyHtml: '<p><strong>Hallo {{Name}},</strong></p><p>dein Event <strong>{{EventTitle}}</strong> wurde erfolgreich erstellt.</p><p>Du kannst die Teilnehmer in der <a href="{{AppUrl}}">Event Experience Platform</a> verwalten.</p><p>Viele Grüße,<br>Team Event Experience Platform</p>' },
    ];

    let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
    try {
      const typeResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1
      );
      if (typeResp.ok) {
        const typeData = await typeResp.json();
        listItemType = typeData.d?.ListItemEntityTypeFullName || typeData.ListItemEntityTypeFullName || listItemType;
      }
    } catch { /* Fallback */ }

    for (const t of defaults) {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
        '__metadata': { 'type': listItemType },
        'Title': `${t.TemplateType}_${t.Language}`,
        'TemplateType': t.TemplateType,
        'Language': t.Language,
        'Subject': t.Subject,
        'HeadingColor': t.HeadingColor,
        'Heading': t.Heading,
        'BodyHtml': t.BodyHtml,
      });
    }

    // _Config Eintrag fuer Logos erstellen (Base64 muss manuell in SharePoint eingetragen werden)
    await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
      '__metadata': { 'type': listItemType },
      'Title': '_Config',
      'TemplateType': '_Config',
      'Language': '',
      'Subject': '',
      'HeadingColor': '',
      'Heading': '',
      'BodyHtml': '',
      'LogoBase64': '',           // Manuell: Base64 Data-URI von Deloitte_Logo.png eintragen
      'DefaultImageBase64': '',   // Manuell: Base64 Data-URI von dex-orb.png eintragen
    });

    await this.configureDefaultView(listName, ['TemplateType', 'Language', 'Subject', 'Heading', 'HeadingColor']);
  }

  /**
   * Sicherstellen dass LogoBase64/DefaultImageBase64 Spalten und _Config Zeile existieren.
   * Wird aufgerufen wenn die Liste bereits existiert (nachtraegliches Upgrade).
   */
  private async ensureEmailTemplatesConfig(listName: string): Promise<void> {
    try {
      // 1. Logo-Spalten nachtraeglich anlegen falls fehlend
      const logoFields = [
        { title: 'LogoBase64', type: 3 },
        { title: 'DefaultImageBase64', type: 3 },
      ];
      for (const f of logoFields) {
        try {
          const check = await this.context.spHttpClient.get(
            `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${f.title}')`,
            SPHttpClient.configurations.v1
          );
          if (!check.ok) {
            await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
              '__metadata': { 'type': 'SP.Field' },
              'Title': f.title,
              'FieldTypeKind': f.type,
              'Required': false,
            });
          }
        } catch { /* Spalte existiert oder Fehler - ignorieren */ }
      }

      // 2. _Config Zeile pruefen und ggf. anlegen
      const configResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (configResp.ok) {
        const configData = await configResp.json();
        const items = configData.value || configData.d?.results || [];
        if (items.length === 0) {
          // _Config Zeile fehlt - anlegen
          let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
          try {
            const typeResp = await this.context.spHttpClient.get(
              `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
              SPHttpClient.configurations.v1
            );
            if (typeResp.ok) {
              const typeData = await typeResp.json();
              listItemType = typeData.d?.ListItemEntityTypeFullName || typeData.ListItemEntityTypeFullName || listItemType;
            }
          } catch { /* Fallback */ }

          await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
            '__metadata': { 'type': listItemType },
            'Title': '_Config',
            'TemplateType': '_Config',
          });
        }
      }
    } catch (err) { console.warn('[DEX] ensureEmailTemplatesConfig fehlgeschlagen:', err); }
  }

  /**
   * Logos als Base64 in die _Config Zeile schreiben (fuer Power Automate Flows).
   * Laedt Deloitte_Logo.png und dex-orb.png aus SiteAssets/DEX_Logos,
   * konvertiert zu Base64 Data-URI und speichert in LogoBase64/DefaultImageBase64.
   */
  public async ensureLogosInConfig(): Promise<void> {
    try {
      // 1. _Config Zeile lesen
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      const configItem = items[0];
      if (!configItem) return;

      // 2. Wenn LogoBase64 schon korrekt befuellt ist (mit image/ MIME-Type), nichts tun
      if (configItem.LogoBase64 && configItem.LogoBase64.startsWith('data:image/')) return;

      // 3. Bilder aus SiteAssets laden
      const logoBase64 = await this.loadFileAsBase64('DEX_Logos/Deloitte_Logo.png');
      const orbBase64 = await this.loadFileAsBase64('DEX_Logos/dex-orb.png');
      if (!logoBase64 && !orbBase64) return;

      // 4. In _Config Zeile schreiben (ueber die getestete _post/_merge Methode)
      const configId = configItem.Id || configItem.d?.Id;
      if (!configId) return;

      const updatePayload: Record<string, unknown> = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailTemplatesListItem' },
      };
      if (logoBase64) updatePayload['LogoBase64'] = logoBase64;
      if (orbBase64) updatePayload['DefaultImageBase64'] = orbBase64;

      await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items(${configId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
            'odata-version': '',
          },
          body: JSON.stringify(updatePayload),
        }
      );
    } catch (err) {
      console.warn('[DEX] ensureLogosInConfig fehlgeschlagen:', err);
    }
  }

  /**
   * Datei aus SiteAssets als Base64 Data-URI laden.
   */
  private async loadFileAsBase64(path: string): Promise<string> {
    try {
      const serverRelativeUrl = this.context.pageContext.web.serverRelativeUrl;
      const fileUrl = `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}/SiteAssets/${path}')/$value`;

      // SPHttpClient mit binaryStringResponseBody fuer Binary-Downloads
      const resp = await this.context.spHttpClient.get(fileUrl, SPHttpClient.configurations.v1, {
        headers: { 'Accept': '*/*' },
      } as ISPHttpClientOptions);
      if (!resp.ok) {
        console.warn('[DEX] loadFileAsBase64 fehlgeschlagen:', path, resp.status);
        return '';
      }
      const blob = await resp.blob();
      if (!blob || blob.size === 0) return '';
      // MIME-Type aus Dateiendung ableiten (SPHttpClient gibt oft application/octet-stream)
      const ext = path.split('.').pop()?.toLowerCase();
      const mimeType = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : blob.type;
      const correctBlob = (blob.type !== mimeType) ? new Blob([blob], { type: mimeType }) : blob;
      return await new Promise<string>((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result as string || '');
        reader.onerror = () => resolve('');
        reader.readAsDataURL(correctBlob);
      });
    } catch (err) {
      console.warn('[DEX] loadFileAsBase64 Error:', path, err);
      return '';
    }
  }

  /**
   * Email-Template aus DEX_EmailTemplates laden.
   * Fallback auf eingebautes Template wenn nicht gefunden.
   */
  public async getEmailTemplate(templateType: string, language: string = 'EN'): Promise<{ subject: string; headingColor: string; heading: string; bodyHtml: string } | null> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '${templateType.replace(/'/g, "''")}' and Language eq '${language.replace(/'/g, "''")}'&$select=Subject,HeadingColor,Heading,BodyHtml&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const data = await resp.json();
        const items = data.value || data.d?.results || [];
        if (items.length > 0) {
          return {
            subject: items[0].Subject || '',
            headingColor: items[0].HeadingColor || '#86bc25',
            heading: items[0].Heading || '',
            bodyHtml: items[0].BodyHtml || '',
          };
        }
      }
    } catch { /* */ }
    return null;
  }

  /**
   * Alle Email-Templates laden (fuer Event-Erstellung / Admin).
   */
  public async getAllEmailTemplates(): Promise<Array<{ id: number; templateType: string; language: string; subject: string; headingColor: string; heading: string; bodyHtml: string }>> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$select=Id,TemplateType,Language,Subject,HeadingColor,Heading,BodyHtml&$orderby=TemplateType,Language&$top=50`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const data = await resp.json();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return (data.value || data.d?.results || []).map((item: any) => ({
          id: item.Id,
          templateType: item.TemplateType || '',
          language: item.Language || 'EN',
          subject: item.Subject || '',
          headingColor: item.HeadingColor || '#86bc25',
          heading: item.Heading || '',
          bodyHtml: item.BodyHtml || '',
        }));
      }
    } catch { /* */ }
    return [];
  }

  // ==================== DEX_Participants Liste ====================

  /**
   * Zentrale Teilnehmer-Liste erstellen falls nicht vorhanden.
   * Speichert pro Person die EventNumbers fuer Registrierung und Warteliste.
   */
  public async ensureParticipantsList(): Promise<void> {
    const listName = 'DEX_Participants';
    const exists = await this.listExists(listName);
    if (exists) {
      await this.ensureMissingParticipantsFields(listName);
      await this.configureDefaultView(listName, [
        'Vorname', 'Nachname', 'Email', 'EventRegistered', 'EventOnWaitlist',
      ]);
      return;
    }

    await this._post(`${this.siteUrl}/_api/web/lists`, {
      '__metadata': { 'type': 'SP.List' },
      'Title': listName,
      'Description': 'Zentrale Teilnehmerliste der DEX Event Experience Platform',
      'BaseTemplate': 100,
      'AllowContentTypes': false,
    });

    const fields = [
      { title: 'Vorname', type: 2 },
      { title: 'Nachname', type: 2 },
      { title: 'Email', type: 2 },
      { title: 'EventRegistered', type: 3 }, // Note fuer beliebig viele EventNumbers
      { title: 'EventOnWaitlist', type: 3 }, // Note fuer beliebig viele EventNumbers
    ];

    for (const f of fields) {
      await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      });
    }

    await this.configureDefaultView(listName, [
      'Vorname', 'Nachname', 'Email', 'EventRegistered', 'EventOnWaitlist',
    ]);

    await this.setEmailsListPermissions(listName);
  }

  private async ensureMissingParticipantsFields(listName: string): Promise<void> {
    const requiredFields = [
      { title: 'Vorname', type: 2 },
      { title: 'Nachname', type: 2 },
      { title: 'Email', type: 2 },
      { title: 'EventRegistered', type: 3 },
      { title: 'EventOnWaitlist', type: 3 },
    ];

    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName,Title&$filter=Hidden eq false&$top=200`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return;
      const data = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const existingNames = new Set((data.value || []).flatMap((f: any) => [f.InternalName, f.Title]));

      for (const f of requiredFields) {
        if (!existingNames.has(f.title)) {
          await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
            '__metadata': { 'type': 'SP.Field' },
            'Title': f.title,
            'FieldTypeKind': f.type,
            'Required': false,
          });
        }
      }
    } catch { /* optional */ }
  }

  /**
   * Teilnehmer-Eintrag per Email suchen
   */
  public async getParticipantByEmail(email: string): Promise<SPParticipant | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items?$filter=Email eq '${email.replace(/'/g, "''")}'&$select=Id,Title,Vorname,Nachname,Email,EventRegistered,EventOnWaitlist&$top=1`,
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
   * Teilnehmer anlegen oder aktualisieren bei Registrierung.
   * Fuegt eventNumber zu EventRegistered oder EventOnWaitlist hinzu.
   */
  public async upsertParticipant(
    vorname: string,
    nachname: string,
    email: string,
    eventNumber: number,
    status: string // 'Angemeldet' | 'Warteliste'
  ): Promise<boolean> {
    try {
      const existing = await this.getParticipantByEmail(email);

      if (existing) {
        // EventNumber zu richtigem Feld hinzufuegen
        let registered = existing.EventRegistered ? existing.EventRegistered.split(',').map(s => s.trim()).filter(s => s) : [];
        let waitlist = existing.EventOnWaitlist ? existing.EventOnWaitlist.split(',').map(s => s.trim()).filter(s => s) : [];
        const en = eventNumber.toString();

        // Erst aus beiden entfernen
        registered = registered.filter(n => n !== en);
        waitlist = waitlist.filter(n => n !== en);

        if (status === 'Warteliste') {
          waitlist.push(en);
        } else {
          registered.push(en);
        }

        await this._merge(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${existing.Id})`,
          {
            'Vorname': vorname,
            'Nachname': nachname,
            'EventRegistered': registered.join(','),
            'EventOnWaitlist': waitlist.join(','),
          }
        );
      } else {
        // Neuen Eintrag erstellen
        const isWaitlist = status === 'Warteliste';
        await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items`, {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_ParticipantsListItem' },
          'Title': email,
          'Vorname': vorname,
          'Nachname': nachname,
          'Email': email,
          'EventRegistered': isWaitlist ? '' : eventNumber.toString(),
          'EventOnWaitlist': isWaitlist ? eventNumber.toString() : '',
        });
      }
      return true;
    } catch {
      return false;
    }
  }

  /**
   * EventNumber aus den Feldern eines Teilnehmers entfernen (bei Abmeldung).
   */
  public async removeParticipantEvent(email: string, eventNumber: number): Promise<boolean> {
    try {
      const existing = await this.getParticipantByEmail(email);
      if (!existing) return false;

      const en = eventNumber.toString();
      const registered = existing.EventRegistered ? existing.EventRegistered.split(',').map(s => s.trim()).filter(s => s && s !== en) : [];
      const waitlist = existing.EventOnWaitlist ? existing.EventOnWaitlist.split(',').map(s => s.trim()).filter(s => s && s !== en) : [];

      await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items(${existing.Id})`,
        {
          'EventRegistered': registered.join(','),
          'EventOnWaitlist': waitlist.join(','),
        }
      );
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Alle Teilnehmer laden (fuer Admin-Seite).
   */
  public async getAllParticipants(): Promise<SPParticipant[]> {
    const allItems: SPParticipant[] = [];
    let url: string | null = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items?$select=Id,Title,Vorname,Nachname,Email,EventRegistered,EventOnWaitlist&$orderby=Nachname,Vorname&$top=500`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }
    return allItems;
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
        'EventNumber', 'EventStatus', 'EventType', 'Location', 'LocationFilter',
        'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
        'WaitlistEnabled', 'Organizer', 'EventImageUrl', 'CalendarLink', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
      ]);
      await this.setColumnFormatting(listName, 'EventImageUrl', {
        '$schema': 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
        'elmType': 'img',
        'attributes': { 'src': '@currentField' },
        'style': { 'max-height': '50px', 'max-width': '120px', 'border-radius': '4px' },
      });
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
      'EventNumber', 'EventStatus', 'EventType', 'Location', 'LocationFilter',
      'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
      'WaitlistEnabled', 'Organizer', 'EventImageUrl', 'CalendarLink', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
    ]);
    await this.setColumnFormatting(listName, 'EventImageUrl', {
      '$schema': 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
      'elmType': 'img',
      'attributes': { 'src': '@currentField' },
      'style': { 'max-height': '50px', 'max-width': '120px', 'border-radius': '4px' },
    });

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
      { title: 'EmailImageBase64', type: 3 }, // Base64 Event-Bild fuer E-Mails/Outlook (Flow ersetzt {{ORB_URL}})
      { title: 'Organizer', type: 2 },
      { title: 'OrganizerEmail', type: 2 },
      { title: 'EventNumber', type: 9 },
      { title: 'OutlookEventId', type: 2 },
      { title: 'CalendarLink', type: 2 },
      { title: 'OutlookBody', type: 3 }, // Multiline - Text fuer Outlook-Termin
      { title: 'EmailLanguage', type: 2 }, // DE oder EN
      { title: 'EmailTemplateOverrides', type: 3 }, // JSON mit Event-spezifischen Template-Anpassungen
      { title: 'CustomFields', type: 3 },
      { title: 'Agenda', type: 3 }, // JSON-Array mit Agenda-Eintraegen
      { title: 'Transfers', type: 3 }, // JSON-Array mit Transferzeiten
      { title: 'Documents', type: 3 }, // JSON-Array mit Dokumenten
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
          // Fehlende Spalte nachtraeglich hinzufuegen
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

      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741826)`,
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
      // Bestehende View-Felder laden um Duplikate zu vermeiden
      const existingResponse = await this.context.spHttpClient.get(
        `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields`,
        SPHttpClient.configurations.v1
      );
      let existingFields: string[] = [];
      if (existingResponse.ok) {
        const existingData = await existingResponse.json();
        existingFields = existingData.Items || existingData.d?.Items || existingData.SchemaXml ? [] : [];
        // Fallback: Versuche Items Array
        if (existingData.Items) {
          existingFields = existingData.Items;
        } else if (existingData.d?.Items) {
          existingFields = existingData.d.Items;
        } else if (existingData.value) {
          existingFields = existingData.value;
        }
      }

      for (const fieldName of fieldNames) {
        // Nur hinzufuegen wenn noch nicht in der View
        if (existingFields.indexOf(fieldName) < 0) {
          await this._post(
            `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/addviewfield('${fieldName}')`,
            {}
          );
        }
      }
    } catch {
      // View-Konfiguration ist optional
    }
  }

  /**
   * Column Formatting auf ein Feld setzen (z.B. Bild-Vorschau fuer URL-Spalten)
   */
  private async setColumnFormatting(listName: string, fieldName: string, formatJson: object): Promise<void> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=InternalName eq '${fieldName}'&$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return;
      const data = await response.json();
      const field = data.value?.[0];
      if (!field) return;

      await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields('${field.Id}')`,
        { CustomFormatter: JSON.stringify(formatJson) }
      );
      // Column Formatting gesetzt
    } catch {
      // Column Formatting ist optional
    }
  }

  // ==================== Events CRUD ====================

  private static readonly EVENT_SELECT = 'Id,Title,EventStatus,EventType,EventNumber,Description,Location,LocationFilter,Audience,FilterMode,StartDate,EndDate,RegistrationDeadline,LastDeregisterDate,MaxParticipants,WaitlistEnabled,EventImageUrl,EmailImageBase64,Organizer,OrganizerEmail,OutlookEventId,CalendarLink,OutlookBody,EmailLanguage,EmailTemplateOverrides,CustomFields,Agenda,Transfers,Documents,RegistrationListName,SubsiteUrl';

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
    outlookBody: string;
    agenda?: string; // JSON-Array mit Agenda-Eintraegen
    transfers?: string; // JSON-Array mit Transferzeiten
    documents?: string; // JSON-Array mit Dokumenten
    emailLanguage?: string;
    emailTemplateOverrides?: string;
    customFields: CustomField[];
  }): Promise<number | null> {
    try {
      // 0. Naechste EventNumber ermitteln
      let nextEventNumber = 1;
      try {
        const enResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=EventNumber&$orderby=EventNumber desc&$top=1`,
          SPHttpClient.configurations.v1
        );
        if (enResp.ok) {
          const enData = await enResp.json();
          if (enData.value && enData.value.length > 0 && enData.value[0].EventNumber) {
            nextEventNumber = enData.value[0].EventNumber + 1;
          }
        }
      } catch { /* Fallback: 1 */ }

      // 1. Subsite fuer das Event erstellen
      const subsiteUrl = await this.createEventSubsite(event.title, event.description);
      if (!subsiteUrl) {
        console.error('[DEX] Subsite konnte nicht erstellt werden');
        throw new Error('Subsite konnte nicht erstellt werden. Fehlende Berechtigung? Bitte wende dich an einen Site-Administrator.');
      }

      // 2. Subsite-Berechtigungen: Members der Parent-Site auf der Subsite berechtigen
      await this.setSubsitePermissions(subsiteUrl, event.organizerEmail);

      // 3. Teilnehmerliste auf der Subsite erstellen
      const fieldMap: Record<string, string> = await this.createRegistrationList(subsiteUrl, event.customFields, event.organizerEmail);

      // Custom Fields mit SP InternalName anreichern
      const enrichedCustomFields = event.customFields.map(cf => ({
        ...cf,
        spInternalName: fieldMap[cf.id] || '',
      }));

      // 3. Event in DEX_Events eintragen
      const payload = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
        'Title': event.title,
        'EventNumber': nextEventNumber,
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
        'EmailImageBase64': '', // Wird ggf. separat gesetzt, Flow nutzt Default aus Config
        'Organizer': event.organizer,
        'OrganizerEmail': event.organizerEmail,
        'OutlookEventId': event.outlookEventId,
        'OutlookBody': event.outlookBody ? buildOutlookBody(event.title, event.outlookBody) : '',
        'EmailLanguage': event.emailLanguage || 'EN',
        'EmailTemplateOverrides': event.emailTemplateOverrides || '',
        'CustomFields': JSON.stringify(enrichedCustomFields),
        'Agenda': event.agenda || '[]',
        'Transfers': event.transfers || '[]',
        'Documents': event.documents || '[]',
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
    } catch (err) {
      if (err instanceof Error) throw err;
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
            'odata-version': '',
          },
          body: JSON.stringify(payload),
        }
      );
      if (!response.ok) {
        const errText = await response.text().catch(() => '');
        console.warn('[DEX] updateEvent failed:', response.status, errText.substring(0, 200));
      }
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

      // 3. DEX_Participants aufraeumen: EventNumber aus allen Teilnehmern entfernen
      if (event.EventNumber) {
        try {
          const allParticipants = await this.getAllParticipants();
          // Parallelize participant cleanup for better performance
          const updatePromises = allParticipants
            .filter(p => {
              const en = String(event.EventNumber);
              const hasRegistered = p.EventRegistered?.split(',').map(s => s.trim()).includes(en);
              const hasWaitlist = p.EventOnWaitlist?.split(',').map(s => s.trim()).includes(en);
              return hasRegistered || hasWaitlist;
            })
            .map(p => this.removeParticipantEvent(p.Email, event.EventNumber));
          // Promise.all mit individueller Fehlerbehandlung (Promise.allSettled nicht verfuegbar in ES2017)
          const safePromises = updatePromises.map(p => p.catch(() => null));
          await Promise.all(safePromises);
        } catch {
          console.warn('[DEX] DEX_Participants konnte nicht aufgeraeumt werden');
        }
      }

      // 4. Event-Dokumente loeschen (SiteAssets/DEX_EventDocs/Event_{number}_*)
      if (event.EventNumber) {
        try {
          const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
          const safeName = (event.Title || '').replace(/[#%&*:<>?/\\|"']/g, '').replace(/\s+/g, '_').substring(0, 50);
          const folderName = safeName ? `Event_${event.EventNumber}_${safeName}` : `Event_${event.EventNumber}`;
          await this._delete(`${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelUrl}/SiteAssets/DEX_EventDocs/${folderName}')`);
        } catch {
          // Fallback: alten Ordnernamen ohne Titel probieren
          try {
            const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
            await this._delete(`${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelUrl}/SiteAssets/DEX_EventDocs/Event_${event.EventNumber}')`);
          } catch { /* Ordner nicht gefunden */ }
        }
      }

      // 5. Event-Eintrag aus DEX_Events loeschen
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
          // Subsite erfolgreich erstellt
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
  ): Promise<Record<string, string>> {
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
      { title: 'Vorname', type: 2 },
      { title: 'Nachname', type: 2 },
      { title: 'ParticipantName', type: 2 }, // Backward compat
      { title: 'ParticipantEmail', type: 2 },
      { title: 'Department', type: 2 },
      { title: 'Location', type: 2 },
      { title: 'JobTitle', type: 2 },
      { title: 'Phone', type: 2 },
      { title: 'Status', type: 6, choices: ['Angemeldet', 'QR versendet', 'Warteliste', 'Eingecheckt', 'Abgemeldet'], metaType: 'SP.FieldChoice' },
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

    // FieldMap wird als Rueckgabewert an den Caller zurueckgegeben

    // Default View konfigurieren (Basis + Custom Fields)
    await this.configureDefaultView(REG_LIST_NAME, [
      'TeilnehmerID', 'Vorname', 'Nachname', 'ParticipantEmail', 'Department', 'Location', 'JobTitle', 'Phone', 'Status', 'RegistrationDate', 'CancellationDate',
      ...customFieldViewNames,
    ], subsiteUrl);

    // Item-Level Permissions
    await this.setItemLevelPermissions(subsiteUrl);

    // Berechtigungen
    await this.setRegistrationListPermissions(subsiteUrl, organizerEmail);

    return fieldMap;
  }

  /**
   * Subsite-Berechtigungen: Owners Full Control, Members Read (damit User die Subsite betreten koennen).
   */
  private async setSubsitePermissions(subsiteUrl: string, organizerEmail: string): Promise<void> {
    try {
      // Owners der Hauptsite: Full Control auf der Subsite
      const ownersResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (ownersResponse.ok) {
        const ownersData = await ownersResponse.json();
        const ownersId = ownersData.d?.Id || ownersData.Id;
        await this._post(
          `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${ownersId}, roledefid=1073741829)`,
          {}
        );
      }

      // Visitors der Hauptsite: Read auf der Subsite (damit User die Subsite betreten koennen)
      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741826)`,
          {}
        );
      }

      // Organizer: Full Control auf der Subsite
      if (organizerEmail) {
        try {
          const userResponse = await this.context.spHttpClient.get(
            `${this.siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(organizerEmail)}')?$select=Id`,
            SPHttpClient.configurations.v1
          );
          if (userResponse.ok) {
            const userData = await userResponse.json();
            const userId = userData.d?.Id || userData.Id;
            await this._post(
              `${subsiteUrl}/_api/web/roleassignments/addroleassignment(principalid=${userId}, roledefid=1073741829)`,
              {}
            );
          }
        } catch { /* Organizer-Berechtigung optional */ }
      }
    } catch {
      console.warn('[DEX] Subsite-Berechtigungen konnten nicht gesetzt werden');
    }
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

      // Visitors: Contribute (damit User sich registrieren koennen)
      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741827)`,
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
    firstName: string,
    surname: string,
    participantEmail: string,
    customData: Record<string, string>,
    status: string = 'Angemeldet',
    customFieldMap?: Record<string, string> // cf.id -> SP InternalName
  ): Promise<boolean> {
    try {
      // Naechste TeilnehmerID ermitteln
      let nextId = 1;
      try {
        const maxResp = await this.context.spHttpClient.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items?$select=TeilnehmerID&$orderby=TeilnehmerID desc&$top=1`,
          SPHttpClient.configurations.v1
        );
        if (maxResp.ok) {
          const maxData = await maxResp.json();
          const items = maxData.value || maxData.d?.results || [];
          if (items.length > 0 && items[0].TeilnehmerID != null) {
            nextId = items[0].TeilnehmerID + 1;
          }
        }
      } catch { /* Fallback: 1 */ }

      // Profildaten laden
      const profile = await this.getCurrentUserProfile();

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': REG_LIST_ITEM_TYPE },
        'Title': participantEmail,
        'TeilnehmerID': nextId,
        'Vorname': firstName,
        'Nachname': surname,
        'ParticipantName': `${firstName} ${surname}`,
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
    firstName: string,
    surname: string,
    customData: Record<string, string>,
    status: string = 'Angemeldet',
    customFieldMap?: Record<string, string>
  ): Promise<boolean> {
    try {
      // Naechste TeilnehmerID ermitteln
      let nextId = 1;
      try {
        const maxResp = await this.context.spHttpClient.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items?$select=TeilnehmerID&$orderby=TeilnehmerID desc&$top=1`,
          SPHttpClient.configurations.v1
        );
        if (maxResp.ok) {
          const maxData = await maxResp.json();
          const items = maxData.value || maxData.d?.results || [];
          if (items.length > 0 && items[0].TeilnehmerID != null) {
            nextId = items[0].TeilnehmerID + 1;
          }
        }
      } catch { /* Fallback: 1 */ }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = {
        'Vorname': firstName,
        'Nachname': surname,
        'ParticipantName': `${firstName} ${surname}`,
        'TeilnehmerID': nextId,
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
    const allItems: SPRegistration[] = [];
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,Title,TeilnehmerID,Vorname,Nachname,ParticipantName,ParticipantEmail,Status,RegistrationDate,CancellationDate,CustomData&$orderby=RegistrationDate&$top=500`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }
    return allItems;
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
   * Teilnehmer einchecken (Status auf 'Eingecheckt' setzen)
   */
  public async checkInParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'Status': 'Eingecheckt' }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Teilnehmer auschecken (Status zurueck auf 'Angemeldet' setzen)
   */
  public async checkOutParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'Status': 'Angemeldet' }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Status eines Teilnehmers auf 'QR versendet' setzen
   */
  public async setQRSentStatus(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        { 'Status': 'QR versendet' }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Registrierung per Email auf einer Subsite finden
   */
  public async getRegistrationByEmail(
    subsiteUrl: string,
    email: string
  ): Promise<SPRegistration | null> {
    try {
      const response = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=ParticipantEmail eq '${email.replace(/'/g, "''")}'&$select=Id,Title,Vorname,Nachname,ParticipantName,ParticipantEmail,Status,RegistrationDate,CancellationDate,CustomData,Department,JobTitle,Location&$top=1`,
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
   * Aktuelle Teilnehmeranzahl ermitteln
   */
  public async getRegistrationCount(subsiteUrl: string): Promise<{ registered: number; waitlist: number }> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const allItems: any[] = [];
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Status&$top=500`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }

    const registered = allItems.filter((i: { Status: string }) => i.Status === 'Angemeldet' || i.Status === 'QR versendet' || i.Status === 'Eingecheckt').length;
    const waitlist = allItems.filter((i: { Status: string }) => i.Status === 'Warteliste').length;
    return { registered, waitlist };
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
        // Ordner erstellt
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

  /**
   * Dokument als Attachment an ein DEX_Events-Item anfuegen.
   * Nutzt native SharePoint List Item Attachments - keine Ordner noetig.
   */
  public async uploadEventDocument(eventId: number, file: File): Promise<string> {
    try {
      const fileName = file.name.replace(/[#%&*:<>?/\\|]/g, '_');

      const response = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/add(FileName='${encodeURIComponent(fileName)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
          body: file,
        } as ISPHttpClientOptions
      );

      if (response.ok) {
        const data = await response.json();
        const relUrl = data.d?.ServerRelativeUrl || data.ServerRelativeUrl || '';
        if (relUrl) return `${window.location.origin}${relUrl}`;
      } else {
        console.warn('[DEX] Attachment upload status:', response.status);
      }

      // Fallback: URL aus bekanntem Pfad
      const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
      return `${window.location.origin}${serverRelUrl}/Lists/DEX_Events/Attachments/${eventId}/${fileName}`;
    } catch (err) {
      console.warn('[DEX] uploadEventDocument error:', err);
    }
    return '';
  }

  /**
   * Attachments eines DEX_Events-Items laden.
   */
  public async getEventAttachments(eventId: number): Promise<Array<{ name: string; url: string; size: number }>> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        const files = data.value || data.d?.results || [];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return files.map((f: any) => ({
          name: f.FileName || '',
          url: `${window.location.origin}${f.ServerRelativeUrl || ''}`,
          size: 0,
        }));
      }
    } catch { /* Attachments nicht verfuegbar */ }
    return [];
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
      // 200 = OK, 403 = existiert aber kein Zugriff (User hat nur Read)
      return response.ok || response.status === 403;
    } catch {
      return false;
    }
  }

  /**
   * ID der SharePoint-Gruppe "DEALL" (alle Deloitte-Mitarbeiter) ermitteln.
   */
  /**
   * ID der Visitors-Gruppe ermitteln (dort ist DEALL / alle Deloitte-Mitarbeiter hinterlegt).
   */
  private async getVisitorsGroupId(): Promise<number | null> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedvisitorgroup?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const d = await resp.json();
        return d.d?.Id || d.Id || null;
      }
    } catch { /* */ }
    return null;
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
   * MERGE-Request fuer Item-Updates.
   * SharePoint verarbeitet den MERGE korrekt, antwortet aber auf manchen
   * Subsite-Listen mit 406 (Accept-Format nicht unterstuetzt). Da bei MERGE
   * kein Response-Body benoetigt wird, behandeln wir 406 als Erfolg.
   */
  private async _merge(url: string, body: object): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(body),
    };
    const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    // 406: SharePoint hat den MERGE ausgefuehrt, kann aber nicht im
    // gewuenschten Format antworten. Daten sind trotzdem gespeichert.
    if (response.status === 406) {
      return { ok: true, status: 204, statusText: 'No Content' } as unknown as SPHttpClientResponse;
    }
    return response;
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
