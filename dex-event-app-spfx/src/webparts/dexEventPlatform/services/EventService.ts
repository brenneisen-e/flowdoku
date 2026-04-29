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
import { wrapTemplateForStorage } from './EmailTemplates';

/**
 * HTML-Body fuer die OutlookDeclineReminder-Mail (EN) - komplett im
 * Deloitte-Design gewrappt, damit er vom DEX_SEND_MAIL-Flow direkt versendet
 * werden kann (der Flow ersetzt nur {{LOGO_URL}} und {{ORB_URL}}, wickelt aber
 * keinen Template-Wrapper mehr drumherum).
 *
 * Der Hinweis auf die Warteliste wurde bewusst durch eine neutrale Formulierung
 * ("your spot can be offered to someone else") ersetzt, damit die Mail auch
 * dann korrekt wirkt, wenn das konkrete Event keine Warteliste hat.
 */
// Standard-Decline-Reminder: Person hat selbst direkt abgelehnt (kein OnBehalfOf).
// Cancel-Button funktioniert direkt — die eingeloggte Person ist auch der Teilnehmer.
const OUTLOOK_DECLINE_BODY_EN = wrapTemplateForStorage(
  '#ed8b00',
  'You declined the Outlook invite',
  'Event {{EventTitle}}',
  `<p>Dear {{Name}},</p>
<p>we noticed that you declined the Outlook calendar invitation for <strong>{{EventTitle}}</strong>, but you are still listed as a confirmed participant.</p>
<p>If you no longer want to attend, please also cancel your registration.</p>
<p style="margin:24px 0;text-align:center;"><a href="{{CancelUrl}}" style="display:inline-block;padding:12px 28px;background:#da291c;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Cancel my registration</a></p>
<p>If you clicked decline by accident, you can simply ignore this message.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);

const OUTLOOK_DECLINE_BODY_DE = wrapTemplateForStorage(
  '#ed8b00',
  'Du hast den Outlook-Termin abgelehnt',
  'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>wir haben gesehen, dass du die Outlook-Kalendereinladung f\u00FCr <strong>{{EventTitle}}</strong> abgelehnt hast \u2013 du bist aber noch als offiziell angemeldet gelistet.</p>
<p>Falls du nicht mehr teilnehmen m\u00F6chtest, melde dich bitte auch offiziell ab.</p>
<p style="margin:24px 0;text-align:center;"><a href="{{CancelUrl}}" style="display:inline-block;padding:12px 28px;background:#da291c;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Anmeldung stornieren</a></p>
<p>Falls du versehentlich abgesagt hast, kannst du diese Mail einfach ignorieren.</p>
<p style="margin-top:24px;"><strong>Viele Gr\u00FC\u00DFe</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// OnBehalfOf-Decline-Reminder: Outlook-Decline kam von einer Assistenz, die den
// Termin im Namen einer anderen Person (Partner/Director) abgelehnt hat. Die Mail
// geht an die Mailbox der Assistenz/des Partners. Der Cancel-Button funktioniert
// nur fuer den registrierten Partner selbst (SP-Item-Level-Security). Fuer die
// Assistenz gibt's einen zweiten Button, der via mailto: eine Bitte um Abmeldung
// an die Event-Organizer schickt — die haben Full Control auf der Teilnehmerliste
// und koennen den Eintrag direkt im Admin Center loeschen.
const OUTLOOK_DECLINE_BODY_ONBEHALF_EN = wrapTemplateForStorage(
  '#ed8b00',
  'You declined the Outlook invite',
  'Event {{EventTitle}}',
  `<p>Dear {{Name}},</p>
<p>we noticed that the Outlook calendar invitation for <strong>{{EventTitle}}</strong> was declined on your behalf, but you are still listed as a confirmed participant.</p>
<p>If you no longer want to attend, please also cancel your registration.</p>
<p style="margin:24px 0;text-align:center;"><a href="{{CancelUrl}}" style="display:inline-block;padding:12px 28px;background:#da291c;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Cancel my registration</a></p>
<p style="font-size:12px;color:#666;text-align:center;margin:0 0 24px;">This button only works when clicked by the registered participant themselves (Partner/Director).</p>
<p style="margin:8px 0 12px;"><strong>Are you the assistant handling this for the participant?</strong> Use the button below to forward this request to the event organizer(s) — they will cancel the registration on the participant's behalf:</p>
<p style="margin:0 0 24px;text-align:center;"><a href="{{AssistantForwardUrl}}" style="display:inline-block;padding:10px 24px;background:#0d6efd;color:#fff;text-decoration:none;border-radius:6px;font-weight:600;font-size:0.95rem;">Forward to organizer (as assistant)</a></p>
<p>If the decline was sent by accident, you can simply ignore this message.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);

const OUTLOOK_DECLINE_BODY_ONBEHALF_DE = wrapTemplateForStorage(
  '#ed8b00',
  'Outlook-Termin abgelehnt',
  'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>die Outlook-Kalendereinladung f\u00FCr <strong>{{EventTitle}}</strong> wurde in deinem Namen abgelehnt — du bist aber noch als offiziell angemeldet gelistet.</p>
<p>Falls du nicht mehr teilnehmen m\u00F6chtest, melde dich bitte auch offiziell ab.</p>
<p style="margin:24px 0;text-align:center;"><a href="{{CancelUrl}}" style="display:inline-block;padding:12px 28px;background:#da291c;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Anmeldung stornieren</a></p>
<p style="font-size:12px;color:#666;text-align:center;margin:0 0 24px;">Dieser Button funktioniert nur, wenn der/die registrierte Teilnehmer/in selbst (Partner/Director) klickt.</p>
<p style="margin:8px 0 12px;"><strong>Bist du die Assistenz der angemeldeten Person?</strong> Bitte nutze den folgenden Button, um diese Anfrage an die Event-Organisator:innen weiterzuleiten — sie melden den/die Teilnehmer/in dann ab:</p>
<p style="margin:0 0 24px;text-align:center;"><a href="{{AssistantForwardUrl}}" style="display:inline-block;padding:10px 24px;background:#0d6efd;color:#fff;text-decoration:none;border-radius:6px;font-weight:600;font-size:0.95rem;">An Organizer weiterleiten (als Assistenz)</a></p>
<p>Falls die Absage versehentlich verschickt wurde, kannst du diese Mail einfach ignorieren.</p>
<p style="margin-top:24px;"><strong>Viele Gr\u00FC\u00DFe</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// Meeting-Forward-Notification-FYI: Ein Teilnehmer hat den Outlook-Termin an eine
// dritte Person weitergeleitet, die NICHT in der SharePoint-Teilnehmerliste steht.
// Die Mail geht an den Organizer. Platzhalter:
//   {{OrganizerFirstName}} — Vorname des Organizers (Anrede)
//   {{Forwarder}}          — Person, die den Termin weitergeleitet hat
//   {{Recipient}}          — Name der hinzugefuegten Person
//   {{RecipientEmail}}     — Email der hinzugefuegten Person ('nicht aufgeloest' bei Externen)
//   {{EventTitle}}         — Event-Titel
//   {{AppUrl}}             — DEX-App-URL (Organizer kann dort manuell registrieren)
const OUTLOOK_FORWARD_BODY_EN = wrapTemplateForStorage(
  '#0d6efd',
  'Meeting was forwarded',
  'Event {{EventTitle}}',
  `<p>Hi {{OrganizerFirstName}},</p>
<p>FYI: <strong>{{Forwarder}}</strong> forwarded the Outlook invitation for <strong>{{EventTitle}}</strong> to <strong>{{Recipient}}</strong> ({{RecipientEmail}}).</p>
<p><strong>{{Recipient}} is currently NOT registered in the DEX participant list.</strong> This person still needs to register via the app in order to get a ParticipantID and QR code and to appear in the official participant list.</p>
<p style="margin:24px 0;text-align:center;"><a href="{{AppUrl}}" style="display:inline-block;padding:12px 28px;background:#86bc25;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">Open DEX App</a></p>
<p style="font-size:13px;color:#555;margin-top:24px;">Possible next steps:</p>
<ul style="font-size:13px;color:#555;margin:0 0 24px 16px;padding:0;">
<li>Ask {{Recipient}} to register themselves via the app.</li>
<li>Or: register {{Recipient}} manually as organizer via "Register for another person".</li>
<li>Or: remove {{Recipient}} from the Outlook meeting if they should not attend.</li>
</ul>
<p style="font-size:12px;color:#999;">This message was generated automatically (Microsoft Outlook Meeting Forward Notification).</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);

const OUTLOOK_FORWARD_BODY_DE = wrapTemplateForStorage(
  '#0d6efd',
  'Termin wurde weitergeleitet',
  'Event {{EventTitle}}',
  `<p>Hallo {{OrganizerFirstName}},</p>
<p>zur Info: <strong>{{Forwarder}}</strong> hat die Outlook-Einladung f\u00FCr <strong>{{EventTitle}}</strong> an <strong>{{Recipient}}</strong> ({{RecipientEmail}}) weitergeleitet.</p>
<p><strong>{{Recipient}} ist aktuell NICHT in der DEX-Teilnehmerliste registriert.</strong> Die Person muss sich ggf. noch selbst \u00FCber die App anmelden, damit sie eine TeilnehmerID und einen QR-Code bekommt und in der offiziellen Teilnehmerliste erscheint.</p>
<p style="margin:24px 0;text-align:center;"><a href="{{AppUrl}}" style="display:inline-block;padding:12px 28px;background:#86bc25;color:#fff;text-decoration:none;border-radius:6px;font-weight:700;">DEX-App \u00F6ffnen</a></p>
<p style="font-size:13px;color:#555;margin-top:24px;">M\u00F6gliche Handlungsoptionen:</p>
<ul style="font-size:13px;color:#555;margin:0 0 24px 16px;padding:0;">
<li>{{Recipient}} bitten, sich selbst \u00FCber die App zu registrieren.</li>
<li>Oder: {{Recipient}} als Organizer manuell \u00FCber "F\u00FCr andere Person registrieren" eintragen.</li>
<li>Oder: {{Recipient}} aus dem Outlook-Termin entfernen, falls nicht gew\u00FCnscht.</li>
</ul>
<p style="font-size:12px;color:#999;">Diese Mail wurde automatisch erzeugt (Microsoft Outlook Meeting Forward Notification).</p>
<p style="margin-top:24px;"><strong>Viele Gr\u00FC\u00DFe</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// Nachruecken-Mail (PA-Flow DEX_IDReorder queued sie) — muss pre-wrapped sein,
// weil die Flow-seite den BodyHtml raw verwendet (ohne wrapTemplate). Client-Code
// erkennt die Pre-Wrap in buildEmailFromTemplate() und skippt den Wrap dann.
const NACHRUECKEN_BODY_EN = wrapTemplateForStorage(
  '#86bc25',
  'You got a spot!',
  'Event {{EventTitle}}',
  `<p>Dear {{Name}},</p>
<p>Great news! A spot has become available and you have been <strong>moved from the waitlist to a confirmed participant</strong> for the event <strong>{{EventTitle}}</strong>.</p>
<p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="{{AppUrl}}">Event Experience Platform</a>.</p>
<p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>`
);

const NACHRUECKEN_BODY_DE = wrapTemplateForStorage(
  '#86bc25',
  'Du bist nachger\u00FCckt!',
  'Event {{EventTitle}}',
  `<p>Hallo {{Name}},</p>
<p>gute Nachrichten! Ein Platz ist frei geworden und du bist von der Warteliste <strong>als Teilnehmer best\u00E4tigt</strong> f\u00FCr das Event <strong>{{EventTitle}}</strong>.</p>
<p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig \u00FCber die <a href="{{AppUrl}}">Event Experience Platform</a> ab.</p>
<p style="margin-top:24px;"><strong>Viele Gr\u00FC\u00DFe</strong><br><br><strong>Dein Event-Team</strong></p>`
);

// Fester Listenname auf jeder Subsite
const REG_LIST_NAME = 'Teilnehmer';
const REG_LIST_ITEM_TYPE = 'SP.Data.TeilnehmerListItem';
// v7.28: Counter-Liste fuer atomare TeilnehmerID-Vergabe (ETag-basiert).
// Pro Subsite eine Liste mit genau einem Item, dessen NextValue beim
// Anmelden via If-Match-Header inkrementiert wird. So koennen mehrere
// User parallel registrieren ohne dass IDs doppelt vergeben werden.
//
// Listenname mit Unterstrich → SharePoint kodiert das in der Item-Type-
// Bezeichnung als `_x005f_`. Genauso wie wir das schon bei DEX_Events
// machen ('SP.Data.DEX_x005f_EventsListItem'). Ohne dieses Encoding
// schlaegt der POST stillschweigend mit HTTP 400 fehl, weil SP den
// Typ-Namen nicht aufloesen kann.
const COUNTER_LIST_NAME = 'DEX_TeilnehmerCounter';
const COUNTER_LIST_ITEM_TYPE = 'SP.Data.DEX_x005f_TeilnehmerCounterListItem';

export interface DeclinedAttendee {
  email: string;
  name: string;
}

export type DeclineCheckReason = 'no-pointer' | 'not-found' | 'forbidden' | 'error';

// Flaches Ergebnis-Shape, weil der TS-4.7-Compiler bei Discriminated Unions
// ueber eine async-Funktion + React-Callback teils nicht korrekt narrowed.
// Erfolg: ok=true, attendees gefuellt, reason/message leer.
// Fehler: ok=false, reason gesetzt, attendees leeres Array.
export interface DeclineCheckResult {
  ok: boolean;
  attendees: DeclinedAttendee[];
  reason?: DeclineCheckReason;
  message?: string;
}

export interface SPEvent {
  Id: number;
  Title: string;
  EventStatus: string;
  EventType?: string; // @deprecated seit v5.2 — wird aus CustomFields abgeleitet; Spalte kann aus DEX_Events entfernt werden.
  EventNumber: number;
  Description: string;
  Location: string;
  LocationAddress: string; // JSON-String: { street, houseNo, zip, city }
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
  DisableEmails: boolean; // true = keine E-Mails bei An-/Abmeldung
  DisableOutlook: boolean; // true = keine Outlook-Kalendereintraege
  // v8.5: Granulare Organizer-BCC-Modi
  NotifyOrgRegisterMode?: string; // 'Never' | 'Always' | 'FromDate'
  NotifyOrgRegisterFromDate?: string;
  NotifyOrgCancelMode?: string; // 'Never' | 'Always' | 'AfterDeadline'
  ExcludedUsers?: string; // v8.6: semikolon-separierte User-Mails die das Event NICHT sehen sollen
  IsFictive?: boolean; // true = Test-Event (nur Admin + eigene Organizer sichtbar)
  DurchstarterCapacity?: number; // B2Run: getrennte Kapazitaet
  FunstarterCapacity?: number;   // B2Run: getrennte Kapazitaet
  CustomFields: string; // JSON-String mit konfigurierbaren Feldern
  Agenda: string; // JSON-Array mit Agenda-Eintraegen
  Transfers: string; // JSON-Array mit Transferzeiten
  Documents: string; // JSON-Array mit Dokumenten
  FunZone: string; // JSON-Array mit Quiz-Fragen
  QuizClusterSize?: number; // 1..4 - wie viele Fragen pro Quiz-Ansicht. Optional, Default 1.
  ParentEventId?: string; // Seit v6.4: wenn gesetzt, ist dies ein Sub-Event und zeigt auf das Parent-Event. Leer = Top-Level-Event.
  RegistrationListName: string;
  SubsiteUrl: string; // Absolute URL der Event-Subsite
}

export interface CustomField {
  id: string;
  label: string;
  type: 'text' | 'select' | 'number' | 'checkbox' | 'user' | 'roommate';
  required: boolean;
  options?: string[]; // fuer select-Felder
  visible: boolean;
  /** v7.20: Optionaler Hilfe-/Beschreibungstext, der im Registrierungs-
   *  Formular als "i"-Tooltip neben dem Feld-Label sichtbar ist. */
  helpText?: string;
  /** v7.21: Sichtbarkeitsbedingung — Feld nur anzeigen wenn das Quell-Feld
   *  einen der `values` als Antwort hat. */
  showIf?: { fieldId: string; values: string[] };
  /** Optionale externe Links, die unter dem Feld als klickbare Links erscheinen.
   * Aktuell vor allem fuer B2Run-Zustimmung (AGB + Datenschutz von b2run.de). */
  externalLinks?: Array<{ label: string; url: string }>;
  /** v7.11: Bei type=select erlaubt true Mehrfachauswahl. Wert wird in der
   *  Teilnehmerliste " | "-getrennt gespeichert. */
  multi?: boolean;
}

export interface SPRegistration {
  Id: number;
  Title: string; // Email
  TeilnehmerID?: number;
  Anrede?: string;
  Vorname?: string;
  Nachname?: string;
  StarterType?: string; // B2Run: 'Durchstarter' | 'Funstarter'
  PreferredStarterType?: string; // B2Run: Wunsch (bei Fallback/Warteliste)
  QuizScore?: number; // Anzahl richtige Antworten
  QuizAnswers?: string; // JSON-Array der gegebenen Antworten
  QuizCompletedAt?: string; // ISO-DateTime
  ParticipantName: string;
  ParticipantEmail: string;
  Status: string;
  RegistrationDate: string;
  RegisteredByName?: string;   // Audit: Name des Users der die Anmeldung durchfuehrte
  RegisteredByEmail?: string;  // Audit: E-Mail des Users der die Anmeldung durchfuehrte
  CancellationDate: string;
  CancelledByName?: string;    // Audit: Name des Users der die Abmeldung ausgeloest hat
  CancelledByEmail?: string;   // Audit: E-Mail des Users der die Abmeldung ausgeloest hat
  /** v7.16: Check-In-Audit — gesetzt sobald checkInParticipant() aufgerufen wird. */
  CheckedInDate?: string;      // ISO-DateTime, wann der Teilnehmer eingecheckt wurde
  CheckedInByName?: string;    // Name des Helfers, der den Check-In ausgeloest hat
  CheckedInByEmail?: string;   // E-Mail des Helfers, der den Check-In ausgeloest hat
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
      // Recipient-Feld auf Plain Text (RichText=false) setzen, falls es
      // noch im alten RichText-Modus ist. SharePoint wrappt sonst den Wert
      // in <div class="ExternalClassXXXX">...</div>, was den Power Automate
      // Flow "emailMessage/To must be String/email" Fehler ausloest.
      try {
        await this.setRecipientFieldPlainText(listName);
      } catch { /* ignore */ }

      // Cc-Feld nachtraeglich anlegen (fuer Anfrage-Mails von der Landing Page).
      // Bestehende Listen aus aelteren App-Versionen haben das Feld noch nicht.
      try {
        await this.ensureCcFieldExists(listName);
      } catch { /* ignore */ }

      // v8.5: Bcc-Feld nachtraeglich anlegen — wird genutzt um Organizer
      // automatisch in Anmelde-/Abmelde-Bestaetigungen zu BCC'en, ohne den
      // Teilnehmer den Verteiler zu zeigen.
      try {
        await this.ensureBccFieldExists(listName);
      } catch { /* ignore */ }

      // Berechtigungen pruefen
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

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fields: Array<Record<string, any>> = [
      // Recipient als Plain-Text Note-Feld (RichText=false), damit der Flow
      // die Email-Adresse(n) ohne HTML-Wrapping bekommt.
      { title: 'Recipient', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 },
      { title: 'Cc', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 },
      { title: 'Bcc', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 },
      { title: 'RecipientName', type: 2 },
      { title: 'Body', type: 3 }, // Body darf Rich/HTML bleiben (wird als HTML gerendert)
      { title: 'EmailType', type: 6, choices: ['Anmeldung', 'Abmeldung', 'Warteliste', 'Nachruecken', 'Info'], metaType: 'SP.FieldChoice' },
      { title: 'EventTitle', type: 2 },
      { title: 'EventId', type: 2 },
      { title: 'Status', type: 6, choices: ['Pending', 'Sent', 'Failed'], metaType: 'SP.FieldChoice' },
      { title: 'SentDate', type: 4 },
    ];

    for (const f of fields) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': f.metaType || 'SP.Field' },
        'Title': f.title,
        'FieldTypeKind': f.type,
        'Required': false,
      };
      if (f.choices) {
        payload['Choices'] = { 'results': f.choices };
      }
      if (f.metaType === 'SP.FieldMultiLineText') {
        payload['RichText'] = !!f.richText;
        if (typeof f.numberOfLines === 'number') payload['NumberOfLines'] = f.numberOfLines;
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
   * Recipient-Feld auf Plain Text (RichText=false) umstellen.
   * Idempotent: Wenn schon Plain Text, macht nichts.
   * Nur moeglich wenn der Current User Manage Lists Rechte hat (Owner/Admin).
   */
  private async setRecipientFieldPlainText(listName: string): Promise<void> {
    const fieldUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Recipient')`;
    const resp = await this.context.spHttpClient.get(
      `${fieldUrl}?$select=FieldTypeKind,RichText`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return;
    const data = await resp.json();
    const kind = data.FieldTypeKind ?? data.d?.FieldTypeKind;
    const richText = data.RichText ?? data.d?.RichText;
    // Wenn bereits Note + Plain Text: nichts zu tun
    if (kind === 3 && richText === false) return;
    // Feld auf Note + RichText=false patchen
    await this._merge(
      fieldUrl,
      {
        '__metadata': { 'type': 'SP.FieldMultiLineText' },
        'FieldTypeKind': 3,
        'RichText': false,
        'NumberOfLines': 3,
      }
    );
  }

  /**
   * Cc-Feld auf DEX_Emails anlegen, falls noch nicht vorhanden.
   * Multi-line Plain-Text damit auch ;-separierte Mehrfach-Adressen passen.
   */
  private async ensureCcFieldExists(listName: string): Promise<void> {
    const probeUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Cc')?$select=Id`;
    const probe = await this.context.spHttpClient.get(probeUrl, SPHttpClient.configurations.v1);
    if (probe.ok) return;
    await this._post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
      {
        '__metadata': { 'type': 'SP.FieldMultiLineText' },
        'Title': 'Cc',
        'FieldTypeKind': 3,
        'Required': false,
        'RichText': false,
        'NumberOfLines': 3,
      }
    );
  }

  private async ensureBccFieldExists(listName: string): Promise<void> {
    const probeUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Bcc')?$select=Id`;
    const probe = await this.context.spHttpClient.get(probeUrl, SPHttpClient.configurations.v1);
    if (probe.ok) return;
    await this._post(
      `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
      {
        '__metadata': { 'type': 'SP.FieldMultiLineText' },
        'Title': 'Bcc',
        'FieldTypeKind': 3,
        'Required': false,
        'RichText': false,
        'NumberOfLines': 3,
      }
    );
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
    eventId: string,
    cc?: string,
    bcc?: string
  ): Promise<boolean> {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailsListItem' },
        'Title': subject,
        'Recipient': recipient,
        'RecipientName': recipientName,
        'Body': body,
        'EmailType': emailType,
        'EventTitle': eventTitle,
        'EventId': eventId,
        'Status': 'Pending',
      };
      if (cc) payload['Cc'] = cc;
      if (bcc) payload['Bcc'] = bcc;
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Emails')/items`,
        payload
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
      // ActionType:
      //  - Einladen / Ausladen: einzelnen Attendee zum Outlook-Termin hinzufügen/entfernen
      //  - UpdateEvent: Titel/Start/Ende aktualisieren (kein Attendee)
      //  - DeleteEvent: kompletten Kalender-Termin löschen (wird beim Löschen eines Events
      //    aus der App abgesetzt, inkl. CalendarLink damit der Flow nicht auf DEX_Events
      //    angewiesen ist - das Event-Item wird direkt danach aus DEX_Events gelöscht).
      { title: 'ActionType', type: 6, choices: ['Einladen', 'Ausladen', 'UpdateEvent', 'DeleteEvent'], metaType: 'SP.FieldChoice' },
      { title: 'Status', type: 6, choices: ['Pending', 'Sent', 'Failed'], metaType: 'SP.FieldChoice' },
      { title: 'SentDate', type: 4 },
      // CalendarLink (iCalUId) - nur für DeleteEvent nötig, damit der Flow das Outlook-
      // Event auch dann noch finden kann, wenn das DEX_Events-Item schon gelöscht wurde.
      { title: 'CalendarLink', type: 3 },
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
      'Attendee', 'EventId', 'SubEventId', 'ActionType', 'Status', 'SentDate', 'CalendarLink',
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
    actionType: 'Einladen' | 'Ausladen' | 'UpdateEvent'
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

  /**
   * DeleteEvent in die DEX_Outlook-Queue eintragen. Wird vom deleteEvent-Flow
   * aufgerufen, BEVOR das DEX_Events-Item gelöscht wird. Der DEX_Outlook_Einladungen-
   * Flow findet den Outlook-Termin über CalendarLink (iCalUId) und löscht ihn.
   * Attendee bleibt leer - DeleteEvent wirkt event-weit.
   */
  public async queueOutlookDeleteEvent(
    eventId: string,
    eventTitle: string,
    calendarLink: string
  ): Promise<boolean> {
    try {
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`,
        {
          '__metadata': { 'type': 'SP.Data.DEX_x005f_OutlookListItem' },
          'Title': `DeleteEvent: ${eventTitle}`,
          'Attendee': '',
          'EventId': eventId,
          'ActionType': 'DeleteEvent',
          'Status': 'Pending',
          'CalendarLink': calendarLink,
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  /**
   * Liefert alle Attendees, die den Outlook-Kalendertermin abgelehnt haben.
   *
   * Liest den Termin im Postfach der Shared Mailbox `no_reply.events@deloitte.de`
   * via Microsoft Graph. Der Admin-User braucht dafuer delegate/shared access
   * auf das Postfach (Mailbox-Permission) + die SPFx-App muss `Calendars.Read.Shared`
   * im Admin Center genehmigt bekommen.
   *
   * Holt `OutlookEventId` + `CalendarLink` via `GET` auf DEX_Events/{id}. Primaerer
   * Lookup des Outlook-Events ueber `OutlookEventId`. Wenn leer (alte Events):
   * Fallback ueber `iCalUId` per `$filter`.
   *
   * Rueckgabe-Status:
   * - `ok: true`, `attendees: [...]` - Termin gefunden, Declines extrahiert
   * - `ok: false`, `reason: 'no-pointer'` - DEX_Events hat weder OutlookEventId noch CalendarLink
   * - `ok: false`, `reason: 'not-found'` - Outlook-Termin existiert nicht (mehr)
   * - `ok: false`, `reason: 'forbidden'` - Admin hat keine Mailbox-Permission oder Tenant-Admin hat Calendars.Read.Shared nicht genehmigt
   * - `ok: false`, `reason: 'error'` - unerwarteter Fehler
   */
  public async getDeclinedAttendees(
    eventId: number | string
  ): Promise<DeclineCheckResult> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = this.context as any;
    if (!ctx.msGraphClientFactory) return { ok: false, attendees: [], reason: 'error', message: 'Graph-Client nicht verfügbar.' };

    // 1. OutlookEventId + CalendarLink aus DEX_Events holen. Nutzt den bewaehrten
    // `getEvent()`-Path (gleiche Abfrage wie der Rest der App). Direktes
    // $select=OutlookEventId,CalendarLink hatte in v5.18 zu leeren Strings
    // gefuehrt obwohl die Spalten in SharePoint gefuellt waren.
    let outlookEventId = '';
    let calendarLink = '';
    let loadedEvent = false;
    try {
      const numericId = Number(eventId);
      const spEvent = await this.getEvent(numericId);
      if (spEvent) {
        loadedEvent = true;
        outlookEventId = String(spEvent.OutlookEventId || '');
        calendarLink = String(spEvent.CalendarLink || '');
        console.warn('[DEX] getDeclinedAttendees: Event geladen', {
          id: numericId,
          outlookEventIdLen: outlookEventId.length,
          calendarLinkLen: calendarLink.length,
        });
      } else {
        console.warn('[DEX] getDeclinedAttendees: getEvent() lieferte null', { eventId });
      }
    } catch (err) {
      console.warn('[DEX] getDeclinedAttendees: getEvent() warf', err);
    }
    if (!outlookEventId && !calendarLink) {
      return {
        ok: false,
        attendees: [],
        reason: 'no-pointer',
        message: loadedEvent
          ? `Event-Item (Id=${eventId}) enthält weder OutlookEventId noch CalendarLink.`
          : `Event-Item (Id=${eventId}) konnte nicht aus DEX_Events geladen werden (403/404?). Details siehe Browser-Console.`,
      };
    }

    // 2. Outlook-Termin via Graph laden
    const mailbox = 'no_reply.events@deloitte.de';
    try {
      const client = await ctx.msGraphClientFactory.getClient('3');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let ev: any = null;
      if (outlookEventId) {
        ev = await client.api(`/users/${mailbox}/events/${outlookEventId}`)
          .select('id,subject,attendees')
          .get();
      } else {
        const escaped = calendarLink.replace(/'/g, "''");
        const resp = await client.api(`/users/${mailbox}/events`)
          .filter(`iCalUId eq '${escaped}'`)
          .select('id,subject,attendees')
          .top(1)
          .get();
        ev = (resp?.value || [])[0] || null;
      }
      if (!ev) return { ok: false, attendees: [], reason: 'not-found' };
      if (!Array.isArray(ev.attendees)) return { ok: true, attendees: [] };
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const declined = ev.attendees
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .filter((a: any) => a?.status?.response === 'declined')
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .map((a: any) => ({
          email: String(a?.emailAddress?.address || '').toLowerCase(),
          name: String(a?.emailAddress?.name || ''),
        }))
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .filter((a: any) => a.email);
      return { ok: true, attendees: declined };
    } catch (err) {
      console.warn('[DEX] getDeclinedAttendees failed:', err);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const status = (err as any)?.statusCode || (err as any)?.status;
      if (status === 403 || status === 401) return { ok: false, attendees: [], reason: 'forbidden' };
      if (status === 404) return { ok: false, attendees: [], reason: 'not-found' };
      return { ok: false, attendees: [], reason: 'error', message: err instanceof Error ? err.message : String(err) };
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

  // ==================== DEX_ChangeLog (v9.0) ====================
  // Audit-Liste fuer alle Aenderungen an Events und Teilnehmern. Read-
  // Berechtigung fuer Organizer/Admin (gleiche Permission-Pattern wie
  // DEX_Roles), Schreibrechte fuer alle (damit User-Aktionen wie
  // Anmeldung/Abmeldung mitloggen koennen).
  public async ensureChangeLogList(): Promise<void> {
    const listName = 'DEX_ChangeLog';
    try {
      const exists = await this.listExists(listName);
      if (exists) {
        try { await this.ensureChangeLogPermissions(listName); } catch { /* */ }
        return;
      }
      // Liste existiert nicht — versuchen sie anzulegen. Schlaegt fehl
      // wenn der aktuelle User keine Owner-Permissions hat. Dann wird der
      // App-Start nicht blockiert (Audit-Log ist best-effort fuer User
      // ohne Schreibrechte auf der Liste-Erstellung).
      const createResp = await this._post(`${this.siteUrl}/_api/web/lists`, {
        '__metadata': { 'type': 'SP.List' },
        'Title': listName,
        'Description': 'Audit-Log fuer alle Aenderungen an Events und Teilnehmern (v9.0)',
        'BaseTemplate': 100,
        'AllowContentTypes': false,
      });
      if (!createResp.ok) {
        console.warn('[DEX] DEX_ChangeLog konnte nicht angelegt werden — vermutlich fehlen dem User Owner-Rechte. App laeuft weiter, Audit-Eintraege fehlen aber.');
        return;
      }
      const fields = [
        { title: 'Action', type: 6, choices: ['EventCreated', 'EventUpdated', 'EventArchived', 'EventRestored', 'EventDeletedPermanent', 'EventDeletedTest', 'ParticipantRegistered', 'ParticipantCancelled', 'ParticipantReactivated', 'ParticipantUpdated', 'ParticipantCheckedIn', 'ParticipantCheckedOut', 'IDReorder', 'Other'], metaType: 'SP.FieldChoice' },
        { title: 'TargetType', type: 6, choices: ['Event', 'Participant', 'Subsite', 'Other'], metaType: 'SP.FieldChoice' },
        { title: 'TargetId', type: 2 },
        { title: 'TargetName', type: 2 },
        { title: 'EventId', type: 2 },
        { title: 'EventTitle', type: 2 },
        { title: 'ActorName', type: 2 },
        { title: 'ActorEmail', type: 2 },
        { title: 'Details', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 6 },
      ];
      for (const f of fields) {
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const payload: Record<string, any> = {
            '__metadata': { 'type': f.metaType || 'SP.Field' },
            'Title': f.title,
            'FieldTypeKind': f.type,
            'Required': false,
          };
          if (f.choices) payload['Choices'] = { 'results': f.choices };
          if (f.metaType === 'SP.FieldMultiLineText') {
            payload['RichText'] = !!f.richText;
            if (typeof f.numberOfLines === 'number') payload['NumberOfLines'] = f.numberOfLines;
          }
          await this._post(`${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
        } catch { /* einzelne Feld-Fehler ignorieren */ }
      }
      try {
        await this.configureDefaultView(listName, [
          'Created', 'Action', 'TargetType', 'TargetName', 'EventTitle', 'ActorName', 'Details',
        ]);
      } catch { /* View-Setup ist optional */ }
      try { await this.ensureChangeLogPermissions(listName); } catch { /* */ }
    } catch (err) {
      console.warn('[DEX] ensureChangeLogList failed (best-effort, App laeuft weiter):', err);
    }
  }

  // Berechtigungen: Site-Members und alle authentifizierten User koennen
  // Eintraege HINZUFUEGEN (damit Self-Reg/Cancel mitschreibt), aber nur
  // Organizer/Admin koennen LESEN. Setzt Item-Level-Read = "Only their own".
  private async ensureChangeLogPermissions(listName: string): Promise<void> {
    try {
      // 1. Inheritance brechen
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
        {}
      );
      // 2. Owners (Admin-Group) → Full Control
      const ownersResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedownergroup?$select=Id`, SPHttpClient.configurations.v1
      );
      if (ownersResp.ok) {
        const d = await ownersResp.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741829)`, {}
        );
      }
      // 3. Members (Organizer-Group typischerweise) → Contribute (sollen
      //    auch schreiben koennen, damit Organizer-Aktionen wie Event-
      //    Updates protokolliert werden).
      const membersResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/associatedmembergroup?$select=Id`, SPHttpClient.configurations.v1
      );
      if (membersResp.ok) {
        const d = await membersResp.json();
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741827)`, {} // 1073741827 = Contribute
        );
      }
      // 4. Visitors (DEALL / Authenticated Users) → Contribute (damit
      //    User-Aktionen wie Self-Anmeldung mitloggen koennen).
      const visitorsId = await this.getVisitorsGroupId();
      if (visitorsId) {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${visitorsId}, roledefid=1073741827)`, {} // Contribute
        );
      }
      // 5. Item-Level-Read = "ReadAllItems" (1) — Organizer und Admin
      //    muessen ALLE Eintraege sehen, nicht nur eigene. Eigene
      //    Lese-Beschraenkung waere fuer Audit-Log nutzlos.
      await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')`,
        { '__metadata': { 'type': 'SP.List' }, 'ReadSecurity': 1 }
      );
    } catch (e) {
      console.warn('[DEX] ensureChangeLogPermissions failed:', e);
    }
  }

  /**
   * v9.0: Audit-Eintrag schreiben. Best-effort — Fehler werden nur
   * geloggt, blocken die aufrufende Aktion nie. Wird automatisch von
   * createEvent / updateEvent / deleteEvent / registerForEvent /
   * cancelRegistration / adminUpdateRegistration / etc. gerufen.
   */
  public async writeChangeLog(entry: {
    action: string;
    targetType: 'Event' | 'Participant' | 'Subsite' | 'Other';
    targetId?: string;
    targetName?: string;
    eventId?: string;
    eventTitle?: string;
    actorName?: string;
    actorEmail?: string;
    details?: string | Record<string, unknown>;
  }): Promise<void> {
    try {
      const me = this.context.pageContext.user;
      const actorName = entry.actorName || me.displayName || '';
      const actorEmail = (entry.actorEmail || me.email || '').toLowerCase();
      const detailsStr = typeof entry.details === 'string'
        ? entry.details
        : entry.details
          ? JSON.stringify(entry.details)
          : '';
      // CLAUDE.md-Hinweis: bei odata=nometadata KEIN __metadata im Body —
      // SP leitet den Typ aus der URL ab. Robust gegen List-Type-Encoding-
      // Quirks (Bug-Story v7.28 → v7.29). Nutzen wir hier statt verbose-POST
      // damit ein verschmierter Type-Name den ChangeLog-Insert nicht
      // stillschweigend killt.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        'Title': `${entry.action}: ${entry.targetName || entry.targetId || '-'}`,
        'Action': entry.action,
        'TargetType': entry.targetType,
        'TargetId': entry.targetId || '',
        'TargetName': entry.targetName || '',
        'EventId': entry.eventId || '',
        'EventTitle': entry.eventTitle || '',
        'ActorName': actorName,
        'ActorEmail': actorEmail,
        'Details': detailsStr.substring(0, 30000),
      };
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_ChangeLog')/items`;
      const options: ISPHttpClientOptions = {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': '',
        },
        body: JSON.stringify(payload),
      };
      await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    } catch (err) {
      console.warn('[DEX] writeChangeLog failed:', err);
    }
  }

  /**
   * Audit-Log lesen — Organizer/Admin only (durch SP-Permissions
   * geschuetzt). Liefert die letzten N Eintraege, optional gefiltert
   * nach EventId.
   */
  public async readChangeLog(opts?: { eventId?: string; top?: number }): Promise<Array<{
    Id: number; Created: string; Action: string; TargetType: string;
    TargetId: string; TargetName: string; EventId: string; EventTitle: string;
    ActorName: string; ActorEmail: string; Details: string;
  }>> {
    const top = opts?.top || 200;
    let url = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_ChangeLog')/items?$select=Id,Created,Action,TargetType,TargetId,TargetName,EventId,EventTitle,ActorName,ActorEmail,Details&$orderby=Created desc&$top=${top}`;
    if (opts?.eventId) {
      url += `&$filter=EventId eq '${String(opts.eventId).replace(/'/g, "''")}'`;
    }
    try {
      const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return [];
      const data = await resp.json();
      return data.value || data.d?.results || [];
    } catch (err) {
      console.warn('[DEX] readChangeLog failed:', err);
      return [];
    }
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
      // Neuere Templates nachruesten (falls die Liste vor v3.0.27 angelegt wurde
      // und OutlookDeclineReminder noch nicht existiert)
      await this.ensureMissingEmailTemplates(listName);
      // Standard-Templates auf aktuelle Version upgraden (uerschreibt User-Customizing!)
      // Damit Platzhalter wie {{WaitlistPosition}} bei aelteren Tenants nachgezogen werden.
      await this.upgradeStandardEmailTemplates(listName);
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
        BodyHtml: '<p>Dear {{Name}},</p><p>you have successfully registered for the event <strong>{{EventTitle}}</strong>.</p><p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="{{AppUrl}}">Event Experience Platform</a> (\u201EMy Events\u201C).</p><p>For organizational questions about the event, please contact <strong>{{Organizer}}</strong>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Warteliste', Language: 'EN', Subject: 'Waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Waitlist confirmation',
        BodyHtml: '<p>Dear {{Name}},</p><p>you have been placed on the <strong>waitlist</strong> for the event <strong>{{EventTitle}}</strong>.</p><p>Your current position: <strong>#{{WaitlistPosition}}</strong></p><p>We will notify you as soon as a spot becomes available. You can always check your current position in the <a href="{{AppUrl}}">Event Experience Platform</a> under \u201EMy Events\u201C.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Abmeldung', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
        BodyHtml: '<p>Dear {{Name}},</p><p>your registration for the event <strong>{{EventTitle}}</strong> has been <strong>cancelled</strong>.</p><p>If you change your mind, you can register again via the <a href="{{AppUrl}}">Event Experience Platform</a>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Nachruecken', Language: 'EN', Subject: 'Spot available: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'You got a spot!',
        BodyHtml: NACHRUECKEN_BODY_EN },
      { TemplateType: 'EventErstellt', Language: 'EN', Subject: '[Deloitte Eventmanager] - New event created: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event Created',
        BodyHtml: '<p>Dear {{Name}},</p><p>your event <strong>{{EventTitle}}</strong> has been successfully created.</p><p>You can manage participants in the <a href="{{AppUrl}}">Event Experience Platform</a>.</p><p>Regards,<br>Team Event Experience Platform</p>' },
      { TemplateType: 'OutlookDeclineReminder', Language: 'EN', Subject: 'Action Required: Do you also want to cancel your registration? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'You declined the Outlook invite',
        BodyHtml: OUTLOOK_DECLINE_BODY_EN },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'EN', Subject: 'Action Required: Cancel registration for {{EventTitle}}?', HeadingColor: '#ed8b00', Heading: 'Outlook invite declined on behalf',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_EN },
      // ===== DEUTSCH =====
      { TemplateType: 'Anmeldung', Language: 'DE', Subject: 'Anmeldebestätigung: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Anmeldung erfolgreich',
        BodyHtml: '<p>Hallo {{Name}},</p><p>du hast dich erfolgreich für das Event <strong>{{EventTitle}}</strong> angemeldet.</p><p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="{{AppUrl}}">Event Experience Platform</a> (\u201EMeine Events\u201C) ab.</p><p>Zu organisatorischen Fragen zum Event wende dich bitte an <strong>{{Organizer}}</strong>.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Warteliste', Language: 'DE', Subject: 'Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Warteliste-Bestätigung',
        BodyHtml: '<p>Hallo {{Name}},</p><p>du stehst auf der <strong>Warteliste</strong> für das Event <strong>{{EventTitle}}</strong>.</p><p>Deine aktuelle Position: <strong>#{{WaitlistPosition}}</strong></p><p>Wir benachrichtigen dich, sobald ein Platz frei wird. Deinen aktuellen Warteliste-Platz kannst du jederzeit in der <a href="{{AppUrl}}">Event Experience Platform</a> unter \u201EMeine Events\u201C sehen.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Abmeldung', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
        BodyHtml: '<p>Hallo {{Name}},</p><p>deine Anmeldung für das Event <strong>{{EventTitle}}</strong> wurde <strong>storniert</strong>.</p><p>Du kannst dich jederzeit erneut über die <a href="{{AppUrl}}">Event Experience Platform</a> anmelden.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Nachruecken', Language: 'DE', Subject: 'Platz frei: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Du bist nachgerückt!',
        BodyHtml: NACHRUECKEN_BODY_DE },
      { TemplateType: 'EventErstellt', Language: 'DE', Subject: '[Deloitte Eventmanager] - Neues Event erstellt: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event erstellt',
        BodyHtml: '<p>Hallo {{Name}},</p><p>dein Event <strong>{{EventTitle}}</strong> wurde erfolgreich erstellt.</p><p>Du kannst die Teilnehmer in der <a href="{{AppUrl}}">Event Experience Platform</a> verwalten.</p><p>Viele Grüße,<br>Team Event Experience Platform</p>' },
      { TemplateType: 'OutlookDeclineReminder', Language: 'DE', Subject: 'Action Required: Möchtest du dich auch offiziell abmelden? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Du hast den Outlook-Termin abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_DE },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'EN', Subject: 'Action Required: Cancel registration for {{EventTitle}}?', HeadingColor: '#ed8b00', Heading: 'Outlook invite declined on behalf',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_EN },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'DE', Subject: 'Action Required: Anmeldung für {{EventTitle}} stornieren?', HeadingColor: '#ed8b00', Heading: 'Outlook-Termin in deinem Namen abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_DE },
      // Meeting-Forward-Notification: FYI an Organizer wenn weitergeleitete Person nicht registriert ist
      { TemplateType: 'OutlookForwardNotification', Language: 'EN', Subject: 'FYI: Meeting was forwarded — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Meeting was forwarded',
        BodyHtml: OUTLOOK_FORWARD_BODY_EN },
      { TemplateType: 'OutlookForwardNotification', Language: 'DE', Subject: 'FYI: Termin wurde weitergeleitet — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Termin wurde weitergeleitet',
        BodyHtml: OUTLOOK_FORWARD_BODY_DE },
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
   * Fuer Tenants wo DEX_EmailTemplates schon vor v3.0.27 angelegt wurde:
   * neuere Templates (z.B. OutlookDeclineReminder DE+EN) nachruesten, ohne
   * bestehende zu ueberschreiben.
   */
  private async ensureMissingEmailTemplates(listName: string): Promise<void> {
    const newTemplates = [
      { TemplateType: 'OutlookDeclineReminder', Language: 'EN', Subject: 'Action Required: Do you also want to cancel your registration? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'You declined the Outlook invite',
        BodyHtml: OUTLOOK_DECLINE_BODY_EN },
      { TemplateType: 'OutlookDeclineReminder', Language: 'DE', Subject: 'Action Required: Möchtest du dich auch offiziell abmelden? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Du hast den Outlook-Termin abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_DE },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'EN', Subject: 'Action Required: Cancel registration for {{EventTitle}}?', HeadingColor: '#ed8b00', Heading: 'Outlook invite declined on behalf',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_EN },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'DE', Subject: 'Action Required: Anmeldung für {{EventTitle}} stornieren?', HeadingColor: '#ed8b00', Heading: 'Outlook-Termin in deinem Namen abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_DE },
      { TemplateType: 'OutlookForwardNotification', Language: 'EN', Subject: 'FYI: Meeting was forwarded — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Meeting was forwarded',
        BodyHtml: OUTLOOK_FORWARD_BODY_EN },
      { TemplateType: 'OutlookForwardNotification', Language: 'DE', Subject: 'FYI: Termin wurde weitergeleitet — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Termin wurde weitergeleitet',
        BodyHtml: OUTLOOK_FORWARD_BODY_DE },
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

    for (const t of newTemplates) {
      try {
        // Existiert das Template bereits? (TemplateType + Language)
        const checkResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '${t.TemplateType}' and Language eq '${t.Language}'&$top=1&$select=Id`,
          SPHttpClient.configurations.v1
        );
        if (checkResp.ok) {
          const checkData = await checkResp.json();
          const items = checkData.value || checkData.d?.results || [];
          if (items.length > 0) continue; // Schon vorhanden - nicht ueberschreiben
        }
        // Template fehlt - nachlegen
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
      } catch { /* Einzelnen Fehler nicht kritisch */ }
    }
  }

  /**
   * Standard-Email-Templates auf die aktuelle Version aktualisieren.
   * Wird bei jedem App-Start aufgerufen, wenn die Liste schon existiert.
   *
   * !! ACHTUNG !! Ueberschreibt User-Customizing.
   *
   * Hintergrund: Templates wie 'Warteliste' wurden ueber die Zeit erweitert
   * (z.B. {{WaitlistPosition}}-Platzhalter). Aelter angelegte Tenants haben
   * noch die OOTB-Version ohne diese Felder. Diese Funktion zieht den BodyHtml
   * (sowie Subject + Heading) auf den aktuellen Code-Stand nach.
   */
  private async upgradeStandardEmailTemplates(listName: string): Promise<void> {
    const APP_URL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';
    void APP_URL; // Reserviert fuer spaetere Templates die {{AppUrl}} hardcoden
    const standards = [
      // ========== EN ==========
      { TemplateType: 'Anmeldung', Language: 'EN', Subject: 'Registration confirmation: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Registration successful',
        BodyHtml: '<p>Dear {{Name}},</p><p>you have successfully registered for the event <strong>{{EventTitle}}</strong>.</p><p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="{{AppUrl}}">Event Experience Platform</a> (\u201EMy Events\u201C).</p><p>For organizational questions about the event, please contact <strong>{{Organizer}}</strong>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Warteliste', Language: 'EN', Subject: 'Waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Waitlist confirmation',
        BodyHtml: '<p>Dear {{Name}},</p><p>you have been placed on the <strong>waitlist</strong> for the event <strong>{{EventTitle}}</strong>.</p><p>Your current position: <strong>#{{WaitlistPosition}}</strong></p><p>We will notify you as soon as a spot becomes available. You can always check your current position in the <a href="{{AppUrl}}">Event Experience Platform</a> under \u201EMy Events\u201C.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Abmeldung', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
        BodyHtml: '<p>Dear {{Name}},</p><p>your registration for the event <strong>{{EventTitle}}</strong> has been <strong>cancelled</strong>.</p><p>If you change your mind, you can register again via the <a href="{{AppUrl}}">Event Experience Platform</a>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
      { TemplateType: 'Nachruecken', Language: 'EN', Subject: 'Spot available: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'You got a spot!',
        BodyHtml: NACHRUECKEN_BODY_EN },
      { TemplateType: 'EventErstellt', Language: 'EN', Subject: '[Deloitte Eventmanager] - New event created: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event Created',
        BodyHtml: '<p>Dear {{Name}},</p><p>your event <strong>{{EventTitle}}</strong> has been successfully created.</p><p>You can manage participants in the <a href="{{AppUrl}}">Event Experience Platform</a>.</p><p>Regards,<br>Team Event Experience Platform</p>' },
      { TemplateType: 'OutlookDeclineReminder', Language: 'EN', Subject: 'Action Required: Do you also want to cancel your registration? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'You declined the Outlook invite',
        BodyHtml: OUTLOOK_DECLINE_BODY_EN },
      // ========== DE ==========
      { TemplateType: 'Anmeldung', Language: 'DE', Subject: 'Anmeldebestätigung: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Anmeldung erfolgreich',
        BodyHtml: '<p>Hallo {{Name}},</p><p>du hast dich erfolgreich für das Event <strong>{{EventTitle}}</strong> angemeldet.</p><p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="{{AppUrl}}">Event Experience Platform</a> (\u201EMeine Events\u201C) ab.</p><p>Zu organisatorischen Fragen zum Event wende dich bitte an <strong>{{Organizer}}</strong>.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Warteliste', Language: 'DE', Subject: 'Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Warteliste-Bestätigung',
        BodyHtml: '<p>Hallo {{Name}},</p><p>du stehst auf der <strong>Warteliste</strong> für das Event <strong>{{EventTitle}}</strong>.</p><p>Deine aktuelle Position: <strong>#{{WaitlistPosition}}</strong></p><p>Wir benachrichtigen dich, sobald ein Platz frei wird. Deinen aktuellen Warteliste-Platz kannst du jederzeit in der <a href="{{AppUrl}}">Event Experience Platform</a> unter \u201EMeine Events\u201C sehen.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Abmeldung', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
        BodyHtml: '<p>Hallo {{Name}},</p><p>deine Anmeldung für das Event <strong>{{EventTitle}}</strong> wurde <strong>storniert</strong>.</p><p>Du kannst dich jederzeit erneut über die <a href="{{AppUrl}}">Event Experience Platform</a> anmelden.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
      { TemplateType: 'Nachruecken', Language: 'DE', Subject: 'Platz frei: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Du bist nachgerückt!',
        BodyHtml: NACHRUECKEN_BODY_DE },
      { TemplateType: 'EventErstellt', Language: 'DE', Subject: '[Deloitte Eventmanager] - Neues Event erstellt: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event erstellt',
        BodyHtml: '<p>Hallo {{Name}},</p><p>dein Event <strong>{{EventTitle}}</strong> wurde erfolgreich erstellt.</p><p>Du kannst die Teilnehmer in der <a href="{{AppUrl}}">Event Experience Platform</a> verwalten.</p><p>Viele Grüße,<br>Team Event Experience Platform</p>' },
      { TemplateType: 'OutlookDeclineReminder', Language: 'DE', Subject: 'Action Required: Möchtest du dich auch offiziell abmelden? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Du hast den Outlook-Termin abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_DE },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'EN', Subject: 'Action Required: Cancel registration for {{EventTitle}}?', HeadingColor: '#ed8b00', Heading: 'Outlook invite declined on behalf',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_EN },
      { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'DE', Subject: 'Action Required: Anmeldung für {{EventTitle}} stornieren?', HeadingColor: '#ed8b00', Heading: 'Outlook-Termin in deinem Namen abgelehnt',
        BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_DE },
      { TemplateType: 'OutlookForwardNotification', Language: 'EN', Subject: 'FYI: Meeting was forwarded — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Meeting was forwarded',
        BodyHtml: OUTLOOK_FORWARD_BODY_EN },
      { TemplateType: 'OutlookForwardNotification', Language: 'DE', Subject: 'FYI: Termin wurde weitergeleitet — {{EventTitle}}', HeadingColor: '#0d6efd', Heading: 'Termin wurde weitergeleitet',
        BodyHtml: OUTLOOK_FORWARD_BODY_DE },
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

    for (const t of standards) {
      try {
        // Bestehendes Item finden
        const checkResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '${t.TemplateType}' and Language eq '${t.Language}'&$top=1&$select=Id,BodyHtml`,
          SPHttpClient.configurations.v1
        );
        if (!checkResp.ok) continue;
        const checkData = await checkResp.json();
        const items = checkData.value || checkData.d?.results || [];
        if (items.length === 0) {
          // existiert nicht -> anlegen (uebernimmt ensureMissingEmailTemplates fuer einige; hier sicherheitshalber auch)
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
        } else {
          // existiert -> updaten falls BodyHtml vom Default abweicht
          const item = items[0];
          if (item.BodyHtml !== t.BodyHtml) {
            await this._merge(
              `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${item.Id})`,
              {
                'Title': `${t.TemplateType}_${t.Language}`,
                'TemplateType': t.TemplateType,
                'Language': t.Language,
                'Subject': t.Subject,
                'HeadingColor': t.HeadingColor,
                'Heading': t.Heading,
                'BodyHtml': t.BodyHtml,
              }
            );
          }
        }
      } catch { /* einzelnes Template ueberspringen */ }
    }
  }

  /**
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
    // $top=5000 statt 500 — bei mehr als 500 jemals registrierten Personen
    // im Tenant lieferte SharePoint mit $orderby+$top in Kombination mit
    // ILS nicht zuverlaessig nextLink, sodass Eintraege fehlten.
    let url: string | null = `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')/items?$select=Id,Title,Vorname,Nachname,Email,EventRegistered,EventOnWaitlist&$orderby=Nachname,Vorname&$top=5000`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
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

      // Default-View komplett neu aufbauen: ID, Title, EventImageUrl, dann Rest
      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/removeallviewfields`,
          {}
        );
      } catch { /* ignore */ }
      await this.configureDefaultView(listName, [
        'ID', 'LinkTitle', 'EventImageUrl',
        'EventNumber', 'EventStatus', 'Location', 'LocationFilter',
        'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
        'WaitlistEnabled', 'Organizer', 'DisableEmails', 'DisableOutlook',
        'CalendarLink', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
      ], undefined, { rebuild: true });
      await this.setColumnFormatting(listName, 'EventImageUrl', {
        '$schema': 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
        'elmType': 'img',
        'attributes': { 'src': '@currentField' },
        'style': { 'max-height': '60px', 'max-width': '120px', 'border-radius': '6px', 'box-shadow': '0 1px 3px rgba(0,0,0,0.15)' },
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
      'EventNumber', 'EventStatus', 'Location', 'LocationFilter',
      'StartDate', 'EndDate', 'RegistrationDeadline', 'MaxParticipants',
      'WaitlistEnabled', 'Organizer', 'EventImageUrl', 'CalendarLink', 'RegistrationListName', 'RegistrationListUrl', 'SubsiteUrl',
    ], undefined, { rebuild: true });
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
  private getEventsFieldDefinitions(): Array<{ title: string; type: number; choices?: string[]; metaType?: string; richText?: boolean; numberOfLines?: number }> {
    return [
      { title: 'EventStatus', type: 6, choices: ['Under Construction', 'Active', 'Completed', 'Cancelled'], metaType: 'SP.FieldChoice' },
      // EventType-Spalte ab v5.2 deprecated (Feld wird nicht mehr angelegt/aktualisiert).
      // Typ wird beim Laden aus CustomFields abgeleitet. Bestehende Spalte in DEX_Events
      // kann manuell entfernt werden.
      { title: 'Description', type: 3 },
      { title: 'Location', type: 2 },
      { title: 'LocationAddress', type: 2 }, // JSON-String: { street, houseNo, zip, city }
      { title: 'LocationFilter', type: 2 },
      // Audience ist Multi-Line-Text (Note) damit es auch bei 100+ E-Mail-Adressen
      // nicht abgeschnitten wird (Single-Line-Text ist auf 255 Zeichen limitiert).
      // Fuer bestehende Events siehe upgradeAudienceFieldToNote() — migriert die
      // alte Text-Spalte zu Note ohne Datenverlust.
      { title: 'Audience', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 },
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
      { title: 'DisableEmails', type: 8, metaType: 'SP.Field' }, // Boolean - keine E-Mails versenden
      { title: 'DisableOutlook', type: 8, metaType: 'SP.Field' }, // Boolean - keine Outlook-Kalendereintraege
      { title: 'NotifyOrgRegisterMode', type: 6, choices: ['Never', 'Always', 'FromDate'], metaType: 'SP.FieldChoice' }, // v8.5
      { title: 'NotifyOrgRegisterFromDate', type: 4 }, // v8.5: ISO-Date, nur fuer Mode='FromDate'
      { title: 'NotifyOrgCancelMode', type: 6, choices: ['Never', 'Always', 'AfterDeadline'], metaType: 'SP.FieldChoice' }, // v8.5
      { title: 'ExcludedUsers', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 4 }, // v8.6: explizit ausgeschlossene User
      { title: 'IsFictive', type: 8, metaType: 'SP.Field' }, // Boolean - Test-Event (nur Admin + eigene Organizer sichtbar)
      { title: 'DurchstarterCapacity', type: 9 }, // B2Run: Kapazitaet fuer Durchstarter (Number)
      { title: 'FunstarterCapacity', type: 9 }, // B2Run: Kapazitaet fuer Funstarter (Number)
      { title: 'CustomFields', type: 3 },
      { title: 'Agenda', type: 3 }, // JSON-Array mit Agenda-Eintraegen
      { title: 'Transfers', type: 3 }, // JSON-Array mit Transferzeiten
      { title: 'Documents', type: 3 }, // JSON-Array mit Dokumenten
      { title: 'FunZone', type: 3 }, // JSON-Array mit Quiz-Fragen
      { title: 'QuizClusterSize', type: 9 }, // Number - 1..4 Fragen pro Quiz-Ansicht
      { title: 'ParentEventId', type: 2 }, // Seit v6.4: ID des Parent-Events (wenn dies ein Sub-Event ist)
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
          if (f.metaType === 'SP.FieldMultiLineText') {
            payload['RichText'] = !!f.richText;
            if (typeof f.numberOfLines === 'number') payload['NumberOfLines'] = f.numberOfLines;
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
   * Migration: alte Audience-Spalte (Type 2, Single-Line-Text, 255 Zeichen Limit)
   * auf Multi-Line-Text (Type 3, Plain-Text, 63.999 Zeichen) umstellen.
   *
   * Noetig weil bei Zielgruppen mit vielen Email-Adressen (~10+) der 255-Zeichen-
   * Cutoff schon griff und Adressen stumm abgeschnitten wurden.
   *
   * Ablauf (idempotent):
   *   1. Check TypeAsString der Audience-Spalte. Wenn schon 'Note' -> skip.
   *   2. Backup aller Event-Werte (id -> audience) im Speicher.
   *   3. Alte Spalte loeschen.
   *   4. Neue Spalte als SP.FieldMultiLineText (RichText=false, NumberOfLines=4) anlegen.
   *   5. Werte aus dem Backup zurueckschreiben (MERGE pro Event).
   *
   * Laeuft beim App-Start (nur fuer Admins, weil wir Write-Rechte auf DEX_Events brauchen).
   */
  public async upgradeAudienceFieldToNote(): Promise<void> {
    const listName = 'DEX_Events';
    try {
      // 1. TypeAsString abfragen
      const fieldResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Audience')?$select=TypeAsString,FieldTypeKind`,
        SPHttpClient.configurations.v1
      );
      if (!fieldResp.ok) return;
      const fieldData = await fieldResp.json();
      const typeAsString: string = fieldData.TypeAsString || fieldData.d?.TypeAsString || '';
      const fieldTypeKind: number = fieldData.FieldTypeKind ?? fieldData.d?.FieldTypeKind ?? 0;
      if (typeAsString === 'Note' || fieldTypeKind === 3) {
        // Schon migriert
        return;
      }
      if (typeAsString !== 'Text' && fieldTypeKind !== 2) {
        // Unerwarteter Typ - nicht anfassen
        console.warn(`[DEX] upgradeAudienceFieldToNote: Audience hat unerwarteten Typ '${typeAsString}' (kind=${fieldTypeKind}) - skip.`);
        return;
      }

      // 2. Alle Event-Werte laden und backuppen
      const itemsResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Audience&$top=2000`,
        SPHttpClient.configurations.v1
      );
      if (!itemsResp.ok) return;
      const itemsData = await itemsResp.json();
      const items: Array<{ Id: number; Audience: string }> = itemsData.value || itemsData.d?.results || [];
      const backup: Record<number, string> = {};
      for (const it of items) {
        if (it.Audience) backup[it.Id] = it.Audience;
      }
      console.warn(`[DEX] upgradeAudienceFieldToNote: Backup ${Object.keys(backup).length} von ${items.length} Event-Audience-Werten.`);

      // 3. Alte Spalte loeschen
      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Audience')/deleteObject`,
          {}
        );
      } catch (e) {
        console.warn('[DEX] upgradeAudienceFieldToNote: Delete alte Audience-Spalte fehlgeschlagen, Migration abgebrochen:', e);
        return;
      }

      // 4. Neue Spalte als Multi-Line-Text anlegen
      try {
        await this._post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
          {
            '__metadata': { 'type': 'SP.FieldMultiLineText' },
            'Title': 'Audience',
            'FieldTypeKind': 3,
            'Required': false,
            'RichText': false,
            'NumberOfLines': 4,
          }
        );
      } catch (e) {
        console.error('[DEX] upgradeAudienceFieldToNote: Konnte neue Audience-Note-Spalte nicht anlegen - Daten koennten verloren gehen:', e, backup);
        return;
      }

      // 5. Werte zurueckschreiben per _merge (odata=nometadata, daher kein __metadata im Body noetig)
      let restored = 0;
      let failed = 0;
      for (const idStr of Object.keys(backup)) {
        const id = Number(idStr);
        try {
          const resp = await this._merge(
            `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
            { 'Audience': backup[id] }
          );
          if (resp.ok) restored += 1;
          else failed += 1;
        } catch { failed += 1; }
      }
      console.warn(`[DEX] upgradeAudienceFieldToNote: Migration fertig — ${restored} Werte zurueckgeschrieben, ${failed} Fehler.`);
    } catch (e) {
      console.warn('[DEX] upgradeAudienceFieldToNote Error:', e);
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
  private async configureDefaultView(listName: string, fieldNames: string[], baseUrl?: string, opts?: { rebuild?: boolean }): Promise<void> {
    const url = baseUrl || this.siteUrl;
    try {
      let existingFields: string[] = [];
      if (opts?.rebuild) {
        // Komplett neu aufbauen — SP-Defaults (Modified, Created, ID, Type,
        // Compliance-Tag, App Created By, ...) werden rausgeworfen.
        try {
          await this._post(
            `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/removeallviewfields`,
            {}
          );
        } catch { /* ignore */ }
      } else {
        // Nur hinzufuegen (behaelt bestehende SP-Felder bei). Duplikate vermeiden.
        const existingResponse = await this.context.spHttpClient.get(
          `${url}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields`,
          SPHttpClient.configurations.v1
        );
        if (existingResponse.ok) {
          const existingData = await existingResponse.json();
          if (existingData.Items) existingFields = existingData.Items;
          else if (existingData.d?.Items) existingFields = existingData.d.Items;
          else if (existingData.value) existingFields = existingData.value;
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

  private static readonly EVENT_SELECT = 'Id,Title,EventStatus,EventNumber,Description,Location,LocationAddress,LocationFilter,Audience,FilterMode,StartDate,EndDate,RegistrationDeadline,LastDeregisterDate,MaxParticipants,WaitlistEnabled,EventImageUrl,EmailImageBase64,Organizer,OrganizerEmail,OutlookEventId,CalendarLink,OutlookBody,EmailLanguage,EmailTemplateOverrides,DisableEmails,DisableOutlook,NotifyOrgRegisterMode,NotifyOrgRegisterFromDate,NotifyOrgCancelMode,ExcludedUsers,IsFictive,DurchstarterCapacity,FunstarterCapacity,CustomFields,Agenda,Transfers,Documents,FunZone,QuizClusterSize,ParentEventId,RegistrationListName,SubsiteUrl';

  /**
   * Seed-Events anlegen falls sie nicht existieren (einmalig beim ersten Start).
   */
  public async seedEvents(): Promise<void> {
    try {
      // Pruefen ob "Assistenz Meeting 2026" schon existiert
      const check = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=Title eq 'Assistenz Meeting 2026'&$top=1&$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (check.ok) {
        const data = await check.json();
        const items = data.value || data.d?.results || [];
        if (items.length > 0) return; // Existiert bereits
      }

      // Event anlegen
      await this.createEvent({
        title: 'Assistenz Meeting 2026',
        type: 'Other',
        status: 'Active',
        description: 'Assistenz Meeting Mai 2026 - Frankfurt am Main',
        location: 'Frankfurt am Main',
        locationFilter: '',
        audience: 'All',
        filterMode: 'OR',
        startDate: '2026-05-07T11:00:00.000Z',
        endDate: '2026-05-08T15:00:00.000Z',
        registrationDeadline: '2026-04-09T00:00:00.000Z',
        lastDeregisterDate: '',
        maxParticipants: 130,
        waitlistEnabled: true,
        eventImageUrl: '',
        organizer: 'Maerzluft, Petra; Schwartz, Eva',
        organizerEmail: 'pmaerzluft@deloitte.de',
        outlookEventId: '',
        outlookBody: '',
        emailLanguage: 'EN',
        emailTemplateOverrides: '',
        customFields: [
          { id: 'travel', label: 'You will travel with?', type: 'select', required: false, visible: true, options: ['Train', 'Car', 'Public Transport'] },
          { id: 'deutschlandticket', label: 'Do you own a Deutschlandticket?', type: 'select', required: false, visible: true, options: ['Yes', 'No'] },
          { id: 'expenses', label: 'Please insert the total amount of your travel expenses!', type: 'text', required: false, visible: true },
        ],
        agenda: '[]',
        transfers: '[]',
        documents: '[]',
      });
    } catch { /* Seed fehlgeschlagen - nicht kritisch */ }
  }

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
    locationAddress?: string; // JSON-String: { street, houseNo, zip, city }
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
    funZone?: string; // JSON-Array mit Quiz-Fragen
    quizClusterSize?: number; // 1..4 - Fragen pro Quiz-Ansicht
    /** Seit v6.4: wenn gesetzt, wird dieses Event als Sub-Event angelegt und zeigt auf das angegebene Parent-Event. */
    parentEventId?: string;
    emailLanguage?: string;
    emailTemplateOverrides?: string;
    disableEmails?: boolean;
    disableOutlook?: boolean;
    notifyOrgRegisterMode?: 'never' | 'always' | 'fromDate';
    notifyOrgRegisterFromDate?: string;
    notifyOrgCancelMode?: 'never' | 'always' | 'afterDeadline';
    excludedUsers?: string[];
    isFictive?: boolean;
    durchstarterCapacity?: number;
    funstarterCapacity?: number;
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
        'Description': event.description,
        'Location': event.location,
        'LocationAddress': event.locationAddress || '',
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
        // Custom-Event-Logo aus emailTemplateOverrides._eventLogo extrahieren (falls
        // vorhanden) und als EmailImageBase64 persistieren — damit der Power-Automate-Flow
        // es als {{ORB_URL}} in Mail + Outlook-Termin einsetzt.
        'EmailImageBase64': (() => {
          try {
            const o = JSON.parse(event.emailTemplateOverrides || '{}');
            return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
          } catch { return ''; }
        })(),
        'Organizer': event.organizer,
        'OrganizerEmail': event.organizerEmail,
        'OutlookEventId': event.outlookEventId,
        // outlookBody kommt bereits vollstaendig gewickelt + mit aufgeloesten Variablen
        // aus EventCreationPage — hier nur durchreichen.
        'OutlookBody': event.outlookBody || '',
        'EmailLanguage': event.emailLanguage || 'EN',
        'EmailTemplateOverrides': event.emailTemplateOverrides || '',
        'DisableEmails': !!event.disableEmails,
        'DisableOutlook': !!event.disableOutlook,
        'NotifyOrgRegisterMode': (() => {
          const m = event.notifyOrgRegisterMode || 'never';
          return m === 'always' ? 'Always' : m === 'fromDate' ? 'FromDate' : 'Never';
        })(),
        'NotifyOrgRegisterFromDate': event.notifyOrgRegisterFromDate || null,
        'NotifyOrgCancelMode': (() => {
          const m = event.notifyOrgCancelMode || 'never';
          return m === 'always' ? 'Always' : m === 'afterDeadline' ? 'AfterDeadline' : 'Never';
        })(),
        'ExcludedUsers': (event.excludedUsers || []).filter(Boolean).join(';'),
        'IsFictive': !!event.isFictive,
        'DurchstarterCapacity': typeof event.durchstarterCapacity === 'number' ? event.durchstarterCapacity : null,
        'FunstarterCapacity': typeof event.funstarterCapacity === 'number' ? event.funstarterCapacity : null,
        'CustomFields': JSON.stringify(enrichedCustomFields),
        'Agenda': event.agenda || '[]',
        'Transfers': event.transfers || '[]',
        'Documents': event.documents || '[]',
        'FunZone': event.funZone || '[]',
        'QuizClusterSize': typeof event.quizClusterSize === 'number' ? event.quizClusterSize : null,
        'ParentEventId': event.parentEventId || '',
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
   * Admin-Cleanup beim App-Start: alle Events mit EventStatus='Active' und EndDate < jetzt
   * werden automatisch auf 'Completed' gesetzt. Liefert die Anzahl der aktualisierten Events.
   */
  public async markExpiredEventsAsCompleted(): Promise<number> {
    try {
      // SharePoint OData Filter: Active + EndDate < jetzt
      const nowIso = new Date().toISOString();
      const filter = `EventStatus eq 'Active' and EndDate lt datetime'${nowIso}'`;
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=${encodeURIComponent(filter)}&$select=Id,Title,EndDate&$top=500`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return 0;
      const data = await resp.json();
      const items: Array<{ Id: number; Title: string }> = data.value || data.d?.results || [];
      if (items.length === 0) return 0;

      let updated = 0;
      for (const it of items) {
        try {
          const ok = await this.updateEvent(it.Id, { 'EventStatus': 'Completed' });
          if (ok) updated += 1;
        } catch { /* einzelnes Update ueberspringen */ }
      }
      return updated;
    } catch (err) {
      console.warn('[DEX] markExpiredEventsAsCompleted failed:', err);
      return 0;
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

      // 0. Outlook-Kalendereintrag per Queue loeschen (VOR allem anderen, damit
      //    CalendarLink noch vorhanden ist). Der DEX_Outlook_Einladungen-Flow
      //    greift den DeleteEvent-Eintrag auf und loescht den Kalender-Termin
      //    im Shared Mailbox ueber den Flow-Service-Account.
      //    Fehler hier ignorieren - Event-Delete soll trotzdem durchlaufen.
      if (event.CalendarLink) {
        try {
          await this.queueOutlookDeleteEvent(String(eventId), event.Title || '', event.CalendarLink);
        } catch { /* Queue-Fehler ignorieren */ }
      }
      // 1. Subsite RECYCEN (v9.0: nicht mehr per DELETE, sonst landet die
      //    Subsite permanent weg ohne Recycle-Bin-Eintrag. recycle() legt
      //    die Subsite mitsamt Teilnehmerliste 93 Tage in den Site
      //    Collection Recycle Bin → ein Tenant-Admin / Site Collection
      //    Admin kann sie dort wiederherstellen falls noetig.
      if (event.SubsiteUrl) {
        try {
          await this._post(`${event.SubsiteUrl}/_api/web/recycle`, {});
        } catch {
          console.warn('[DEX] Subsite konnte nicht in den Recycle Bin verschoben werden:', event.SubsiteUrl);
        }
      }

      // 2. Event-Bild ebenfalls RECYCEN statt loeschen.
      if (event.EventImageUrl) {
        try {
          const url = new URL(event.EventImageUrl);
          const serverRelUrl = url.pathname;
          if (serverRelUrl.indexOf('DEX_EventImages') >= 0) {
            await this._post(
              `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelUrl}')/recycle`,
              {}
            );
          }
        } catch {
          console.warn('[DEX] Event-Bild konnte nicht in den Recycle Bin verschoben werden');
        }
      }

      // 3. Alte Registrierungsliste recyceln (legacy Events ohne Subsite).
      if (event.RegistrationListName && event.RegistrationListName !== 'Teilnehmer') {
        try {
          await this._post(
            `${this.siteUrl}/_api/web/lists/getbytitle('${event.RegistrationListName.replace(/'/g, "''")}')/recycle`,
            {}
          );
        } catch {
          console.warn('[DEX] Alte Registrierungsliste konnte nicht recycelt werden:', event.RegistrationListName);
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

      // 5. Event-Eintrag aus DEX_Events RECYCEN (v9.0: per recycle() statt
      //    delete(), damit ein Admin via SharePoint-Recycle-Bin das Item
      //    bei Bedarf 93 Tage lang wiederherstellen kann).
      const response = await this._post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/recycle`,
        {}
      );

      // 6. Audit-Eintrag in DEX_ChangeLog (v9.0). Best-effort, blockt
      //    den Loesch-Vorgang nicht falls Logging fehlschlaegt.
      try {
        await this.writeChangeLog({
          action: 'EventDeletedTest', // wird vom Aufrufer ueberschrieben
          targetType: 'Event',
          targetId: String(eventId),
          targetName: event.Title || '',
          eventId: String(eventId),
          eventTitle: event.Title || '',
          details: {
            subsiteUrl: event.SubsiteUrl || '',
            eventNumber: event.EventNumber,
            recycledTo: 'SharePoint Recycle Bin (93 Tage)',
          },
        });
      } catch { /* */ }

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
      { title: 'Anrede', type: 6, choices: ['Frau', 'Herr', 'Divers'], metaType: 'SP.FieldChoice' },
      { title: 'Vorname', type: 2 },
      { title: 'Nachname', type: 2 },
      { title: 'ParticipantName', type: 2 }, // Backward compat
      { title: 'ParticipantEmail', type: 2 },
      { title: 'Department', type: 2 },
      { title: 'Location', type: 2 },
      { title: 'JobTitle', type: 2 },
      { title: 'Phone', type: 2 },
      { title: 'Status', type: 6, choices: ['Angemeldet', 'QR versendet', 'Warteliste', 'Eingecheckt', 'Abgemeldet'], metaType: 'SP.FieldChoice' },
      { title: 'StarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' }, // B2Run: Typ-Auswahl
      { title: 'PreferredStarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' }, // B2Run: Wunsch-Typ (wenn Fallback oder Warteliste)
      { title: 'QuizScore', type: 9 }, // Number - Anzahl richtiger Antworten
      { title: 'QuizAnswers', type: 3 }, // Note - JSON der Antworten (fuer Statistik)
      { title: 'QuizCompletedAt', type: 4 }, // DateTime
      { title: 'RegistrationDate', type: 4 },
      { title: 'RegisteredByName', type: 2 },  // Audit: Name des Users der die Anmeldung durchgefuehrt hat
      { title: 'RegisteredByEmail', type: 2 }, // Audit: E-Mail des Users der die Anmeldung durchgefuehrt hat
      { title: 'LastModifiedDate', type: 4 },
      { title: 'ChangeLog', type: 3 }, // Note (multiline) - Aenderungshistorie
      { title: 'CancellationDate', type: 4 },
      { title: 'CancelledByName', type: 2 },   // Audit: Name des Users der die Abmeldung ausgeloest hat
      { title: 'CancelledByEmail', type: 2 },  // Audit: E-Mail des Users der die Abmeldung ausgeloest hat
      { title: 'CheckedInDate', type: 4 },     // v7.16: Check-In-Audit — Zeitpunkt
      { title: 'CheckedInByName', type: 2 },   // v7.16: Check-In-Audit — Helfer-Name
      { title: 'CheckedInByEmail', type: 2 },  // v7.16: Check-In-Audit — Helfer-E-Mail
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

    // Default View komplett neu aufbauen (Basis + Custom Fields). Mit rebuild:true
    // werden alle SP-Default-Spalten (Modified, Created, ID, Type, Compliance Asset,
    // App Created By, ...) aus der View rausgeworfen — nur funktionelle Felder.
    await this.configureDefaultView(REG_LIST_NAME, [
      'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'ParticipantEmail', 'Department', 'Location', 'JobTitle', 'Phone', 'StarterType', 'PreferredStarterType', 'Status', 'RegistrationDate', 'RegisteredByName', 'RegisteredByEmail', 'CancellationDate', 'CancelledByName', 'CancelledByEmail',
      ...customFieldViewNames,
    ], subsiteUrl, { rebuild: true });

    // Item-Level Permissions
    await this.setItemLevelPermissions(subsiteUrl);

    // Berechtigungen
    await this.setRegistrationListPermissions(subsiteUrl, organizerEmail);

    // v7.28: Counter-Liste fuer atomare TeilnehmerID-Vergabe anlegen
    // (Race-Condition-Schutz bei parallelen Anmeldungen).
    try {
      await this.ensureCounterList(subsiteUrl);
    } catch {
      // Nicht kritisch — falls das schiefgeht, fallback auf max+1 in upsertParticipant.
    }

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
    customFieldMap?: Record<string, string>, // cf.id -> SP InternalName
    starterType?: string, // B2Run: effektiver Typ (nach Fallback)
    preferredStarterType?: string, // B2Run: Wunsch-Typ (was der User eigentlich wollte)
    registeredByName?: string, // Audit: Name des Users der die Anmeldung ausloest
    registeredByEmail?: string // Audit: E-Mail des Users der die Anmeldung ausloest
  ): Promise<boolean> {
    try {
      // ---- Permission-Checks (v3.9.2 / v3.9.3) ----
      // Serverseitige Pruefungen — nicht perfekt (SPFx laeuft im Browser),
      // aber fangt naiven App-Bypass (F12, direkter Service-Aufruf) ab.
      const sessionEmail = (this.context.pageContext.user.email || '').toLowerCase();
      const targetEmail = (participantEmail || '').toLowerCase();

      // Event-Metadaten laden (Deadline + OrganizerEmail) ueber SubsiteUrl.
      // Beide Checks nutzen die gleiche Abfrage — einmal laden, mehrfach pruefen.
      let eventDeadline = '';
      let eventOrganizerEmails: string[] = [];
      try {
        const subsiteEsc = encodeURIComponent(subsiteUrl.replace(/'/g, "''"));
        const evResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${subsiteEsc}'&$top=1&$select=RegistrationDeadline,OrganizerEmail`,
          SPHttpClient.configurations.v1
        );
        if (evResp.ok) {
          const evData = await evResp.json();
          const items = evData.value || evData.d?.results || [];
          if (items.length > 0) {
            eventDeadline = items[0].RegistrationDeadline || '';
            const orgStr: string = items[0].OrganizerEmail || '';
            eventOrganizerEmails = orgStr.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
          }
        }
      } catch { /* Bei Load-Fehler konservativ weitermachen — andere Checks greifen */ }

      // Check A: Darf der User fuer eine andere Person registrieren?
      if (targetEmail && targetEmail !== sessionEmail) {
        const allowed = await this.canRegisterForOthers(subsiteUrl, participantEmail);
        if (!allowed) {
          console.warn(`[DEX] registerForEvent DENIED: ${sessionEmail} versuchte ${targetEmail} zu registrieren — weder Organizer noch Admin noch erlaubter Assistant-Fall.`);
          return false;
        }
      }

      // Check B: Deadline abgelaufen? Nur Event-Organizer + Admin duerfen nach
      // Deadline registrieren (auch fuer sich selbst). Assistant NICHT — das ist
      // wie ein normaler User.
      if (eventDeadline) {
        const deadlineDate = new Date(eventDeadline);
        if (!isNaN(deadlineDate.getTime()) && deadlineDate < new Date()) {
          const isEventOrganizer = eventOrganizerEmails.indexOf(sessionEmail) >= 0;
          let isAdmin = false;
          try {
            const esc = sessionEmail.replace(/'/g, "''");
            const roleResp = await this.context.spHttpClient.get(
              `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
              SPHttpClient.configurations.v1
            );
            if (roleResp.ok) {
              const rd = await roleResp.json();
              const rItems = rd.value || rd.d?.results || [];
              if (rItems.length > 0 && rItems[0].Role === 'Admin') isAdmin = true;
            }
          } catch { /* ignore */ }

          if (!isEventOrganizer && !isAdmin) {
            console.warn(`[DEX] registerForEvent DENIED (deadline): ${sessionEmail} versuchte nach Deadline ${eventDeadline} zu registrieren — weder Event-Organizer noch Admin.`);
            return false;
          }
        }
      }
      // ---- Ende Permission-Checks ----

      // v7.28: Naechste TeilnehmerID atomar ueber den Subsite-Counter holen
      // (verhindert Race-Conditions bei parallelen Anmeldungen). Fallback
      // auf das alte max+1, wenn die Counter-Liste fuer dieses Event noch
      // nicht existiert (legacy ohne "Spalten fixen"-Lauf).
      let nextId = await this.getNextTeilnehmerId(subsiteUrl);
      if (nextId === undefined) {
        nextId = (await this.getCurrentMaxTeilnehmerId(subsiteUrl)) + 1;
      }

      // Profildaten laden - fuer den TATSAECHLICHEN Teilnehmer (nicht den eingeloggten User!)
      // Wenn jemand fuer eine andere Person registriert, muss deren Profil geladen werden,
      // sonst wird der eigene JobTitle/Department/Office in deren Teilnehmer-Eintrag geschrieben.
      const myEmail = (this.context.pageContext.user.email || '').toLowerCase();
      const profile = participantEmail.toLowerCase() === myEmail
        ? await this.getCurrentUserProfile()
        : await this.getUserProfileByEmail(participantEmail);

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: Record<string, any> = {
        '__metadata': { 'type': REG_LIST_ITEM_TYPE },
        'Title': participantEmail,
        'TeilnehmerID': nextId,
        'Anrede': customData.salutation || '',
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

      // Audit: wer hat die Anmeldung ausgeloest?
      // Bei Self-Registration = der User selbst. Bei "Fuer andere Person registrieren"
      // = der Organizer/Admin der geklickt hat. Fallback wenn nichts uebergeben: aus pageContext.
      const auditName = registeredByName || this.context.pageContext.user.displayName || '';
      const auditEmail = (registeredByEmail || this.context.pageContext.user.email || '').toLowerCase();
      if (auditName) payload['RegisteredByName'] = auditName;
      if (auditEmail) payload['RegisteredByEmail'] = auditEmail;

      // B2Run: Starter-Typ + Wunsch-Typ schreiben (bei normalen Events null)
      if (starterType) payload['StarterType'] = starterType;
      if (preferredStarterType) payload['PreferredStarterType'] = preferredStarterType;

      // Custom Field Werte in die echten SP-Spalten schreiben.
      // Wichtig: Wenn spInternalName fehlt (z.B. weil der Admin das Feld spaeter
      // ergaenzt hat ohne Spalte in der Teilnehmerliste), wuerde der Wert
      // SILENT VERLOREN GEHEN — deshalb ein console.warn damit der Admin im
      // Admin Center per "Custom Fields pruefen" das Mapping fixen kann.
      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          if (!customData[cfId]) continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName) {
            payload[spFieldName] = customData[cfId];
          } else {
            console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields pruefen' ausfuehren.`);
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
    customFieldMap?: Record<string, string>,
    registeredByName?: string, // Audit: Name des Users der die Re-Anmeldung ausloest
    registeredByEmail?: string // Audit: E-Mail des Users der die Re-Anmeldung ausloest
  ): Promise<boolean> {
    try {
      // ---- Permission-Checks (v3.9.2 / v3.9.3) ----
      // Lade die ParticipantEmail aus dem zu reaktivierenden Item und pruefe,
      // ob der aktuelle User dafuer berechtigt ist. Plus Deadline-Check.
      // v7.30: Wir laden hier zusaetzlich die existierende TeilnehmerID, damit
      // beim Reaktivieren die alte ID erhalten bleibt — Counter wird NUR
      // dann angefasst, wenn die alte ID null/0 ist (Legacy-Edge).
      let existingTeilnehmerId = 0;
      try {
        const itemResp = await this.context.spHttpClient.get(
          `${subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${itemId})?$select=ParticipantEmail,TeilnehmerID`,
          SPHttpClient.configurations.v1
        );
        const sessionEmail = (this.context.pageContext.user.email || '').toLowerCase();
        let targetEmail = '';
        if (itemResp.ok) {
          const itemData = await itemResp.json();
          targetEmail = (itemData.ParticipantEmail || itemData.d?.ParticipantEmail || '').toLowerCase();
          const tnId = itemData.TeilnehmerID ?? itemData.d?.TeilnehmerID;
          if (typeof tnId === 'number' && tnId > 0) existingTeilnehmerId = tnId;
        }

        // Check A: fuer andere Person registrieren?
        if (targetEmail && targetEmail !== sessionEmail) {
          const allowed = await this.canRegisterForOthers(subsiteUrl, targetEmail);
          if (!allowed) {
            console.warn(`[DEX] reactivateRegistration DENIED: ${sessionEmail} versuchte ${targetEmail} zu reaktivieren — nicht berechtigt.`);
            return false;
          }
        }

        // Check B: Deadline-Check (Event ueber SubsiteUrl finden)
        const subsiteEsc = encodeURIComponent(subsiteUrl.replace(/'/g, "''"));
        const evResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${subsiteEsc}'&$top=1&$select=RegistrationDeadline,OrganizerEmail`,
          SPHttpClient.configurations.v1
        );
        if (evResp.ok) {
          const evData = await evResp.json();
          const items = evData.value || evData.d?.results || [];
          if (items.length > 0) {
            const deadline = items[0].RegistrationDeadline || '';
            const orgStr: string = items[0].OrganizerEmail || '';
            const orgEmails = orgStr.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
            if (deadline) {
              const deadlineDate = new Date(deadline);
              if (!isNaN(deadlineDate.getTime()) && deadlineDate < new Date()) {
                const isEventOrganizer = orgEmails.indexOf(sessionEmail) >= 0;
                let isAdmin = false;
                try {
                  const esc = sessionEmail.replace(/'/g, "''");
                  const roleResp = await this.context.spHttpClient.get(
                    `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
                    SPHttpClient.configurations.v1
                  );
                  if (roleResp.ok) {
                    const rd = await roleResp.json();
                    const rItems = rd.value || rd.d?.results || [];
                    if (rItems.length > 0 && rItems[0].Role === 'Admin') isAdmin = true;
                  }
                } catch { /* ignore */ }
                if (!isEventOrganizer && !isAdmin) {
                  console.warn(`[DEX] reactivateRegistration DENIED (deadline): ${sessionEmail} versuchte nach Deadline ${deadline} zu reaktivieren.`);
                  return false;
                }
              }
            }
          }
        }
      } catch { /* bei Load-Fehler konservativ: weitermachen */ }
      // ---- Ende Permission-Checks ----

      // Reaktivierung = funktional eine Neuanmeldung mit existierendem Listen-
      // Item. Deshalb wird hier — analog zu registerForEvent — atomar eine
      // neue TeilnehmerID am Counter gezogen. Wer mal #12 war und reaktiviert,
      // bekommt jetzt z.B. die #87, also die naechst-freie ID am Ende der
      // Liste — exakt wie ein Neuzugang. Ohne diesen Schritt blieb der
      // Eintrag mit TeilnehmerID=null haengen, weil im Reaktivierungs-Pfad
      // niemand den DEX_IDReorder-Flow triggert.
      void existingTeilnehmerId;
      let nextId = await this.getNextTeilnehmerId(subsiteUrl);
      if (nextId === undefined) {
        nextId = (await this.getCurrentMaxTeilnehmerId(subsiteUrl)) + 1;
      }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = {
        'Vorname': firstName,
        'Nachname': surname,
        'ParticipantName': `${firstName} ${surname}`,
        'Status': status,
        'TeilnehmerID': nextId,
        'RegistrationDate': new Date().toISOString(),
        'CancellationDate': null,
        'CustomData': JSON.stringify(customData),
      };

      // Audit: wer hat die Re-Anmeldung ausgeloest? (ueberschreibt den Wert von
      // der urspruenglichen Anmeldung, weil das faktisch eine neue Anmeldung ist)
      const auditName = registeredByName || this.context.pageContext.user.displayName || '';
      const auditEmail = (registeredByEmail || this.context.pageContext.user.email || '').toLowerCase();
      if (auditName) body['RegisteredByName'] = auditName;
      if (auditEmail) body['RegisteredByEmail'] = auditEmail;

      if (customFieldMap) {
        for (const cfId of Object.keys(customData)) {
          if (cfId === 'salutation') continue;
          if (!customData[cfId]) continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName) {
            body[spFieldName] = customData[cfId];
          } else {
            console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields pruefen' ausfuehren.`);
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
   * Schlanker MERGE-Helper auf ein einzelnes Teilnehmerlisten-Item — baut
   * keine ChangeLog-Logik, keine FieldMap-Auflösung, keine Default-Felder ein.
   * Genutzt für One-Shot-Migrationen (z.B. T-Shirt-Größen-Import), die direkt
   * bestimmte Felder (inkl. CustomData-JSON + einzelne SP-Spalten) setzen wollen.
   */
  public async mergeRegistrationFields(
    subsiteUrl: string,
    itemId: number,
    body: Record<string, unknown>
  ): Promise<boolean> {
    try {
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
          if (!customData[cfId]) continue;
          const spFieldName = customFieldMap[cfId];
          if (spFieldName) {
            body[spFieldName] = customData[cfId];
          } else {
            console.warn(`[DEX] Custom-Field '${cfId}' hat keine SharePoint-Spalte (spInternalName fehlt) — Wert wird nicht in Teilnehmerliste geschrieben. Bitte im Admin Center 'Custom Fields pruefen' ausfuehren.`);
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
   * v8.0: Admin/Organizer kann Teilnehmerdaten direkt aus dem Admin Center
   * editieren. Erlaubt das Aendern von Anrede/Vorname/Nachname/Email/Phone/
   * Department/Location/JobTitle/Status sowie aller Custom-Felder. Schreibt
   * automatisch ChangeLog-Eintrag mit Wer/Wann/Was-Diff und setzt
   * LastModifiedDate.
   *
   * patch: Nur die echten Spalten-Werte (keine __metadata noetig — _merge
   * sendet odata=nometadata).
   * actor: Audit-Info des aufrufenden Users.
   * oldValues: zum Diff-Bauen, nur Felder mit oldValues[key] !== patch[key]
   * landen im ChangeLog.
   * fieldLabelMap: optional, mappt internal column name -> display label.
   */
  public async adminUpdateRegistration(
    subsiteUrl: string,
    itemId: number,
    patch: Record<string, unknown>,
    actor: { name: string; email: string },
    oldValues?: Record<string, unknown>,
    fieldLabelMap?: Record<string, string>
  ): Promise<boolean> {
    try {
      const changes: string[] = [];
      const now = new Date().toLocaleString('de-DE', {
        day: '2-digit', month: '2-digit', year: 'numeric',
        hour: '2-digit', minute: '2-digit',
      });
      if (oldValues) {
        for (const key of Object.keys(patch)) {
          const oldV = oldValues[key];
          const newV = patch[key];
          // Vergleich als String, damit number vs string nicht stoert
          const oldStr = oldV === null || oldV === undefined ? '' : String(oldV);
          const newStr = newV === null || newV === undefined ? '' : String(newV);
          if (oldStr !== newStr) {
            const label = (fieldLabelMap && fieldLabelMap[key]) || key;
            changes.push(`${label}: "${oldStr}" → "${newStr}"`);
          }
        }
      }
      const changeEntry = changes.length > 0
        ? `[${now}] ${actor.name || actor.email}: ${changes.join(', ')}`
        : '';

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: Record<string, any> = { ...patch, 'LastModifiedDate': new Date().toISOString() };

      // ChangeLog anhaengen (bestehenden Log behalten, neuestes oben)
      if (changeEntry) {
        try {
          const existing = await this.context.spHttpClient.get(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})?$select=ChangeLog`,
            SPHttpClient.configurations.v1
          );
          if (existing.ok) {
            const data = await existing.json();
            const oldLog = data.ChangeLog || data.d?.ChangeLog || '';
            body['ChangeLog'] = oldLog ? `${changeEntry}\n${oldLog}` : changeEntry;
          } else {
            body['ChangeLog'] = changeEntry;
          }
        } catch {
          body['ChangeLog'] = changeEntry;
        }
      }

      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        body
      );
      return response.ok;
    } catch (err) {
      console.warn('[DEX] adminUpdateRegistration error:', err);
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
    // $top=5000 ist das SP-REST-Maximum pro Page. Damit fallen Events bis zu
    // 5000 Teilnehmern in einen einzigen Response — keine Pagination-Edgecases
    // mit fehlendem nextLink. Bei groesseren Listen folgen wir dem nextLink
    // weiter (Schleife unten). Vorher stand hier $top=500, was bei Events mit
    // ≥500 Teilnehmern zu fehlenden Eintraegen fuehrte: SharePoint liefert
    // bei $orderby+$top in Kombination mit Item-Level-Security nicht
    // zuverlaessig nextLink, wenn die erste Page exakt voll ist.
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$orderby=Id asc&$top=5000`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        // Beide OData-Formate abdecken: nometadata (data.value) UND verbose
        // (data.d.results). Vorher nur data.value — bei verbose-Response
        // waeren null Items dazugekommen.
        const page = data.value || data.d?.results || [];
        allItems.push(...page);
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch {
        break;
      }
    }
    return allItems;
  }

  /**
   * TeilnehmerIDs sequentiell neu vergeben (1, 2, 3, ...).
   *
   * Seit v6.5: Zwei-Pass-Reorder.
   * 1. Erst alle **Angemeldeten** (Status ∈ Angemeldet / QR versendet / Eingecheckt)
   *    in Reihenfolge der Registrierung (SP-ItemId asc) → bekommen IDs 1..N.
   * 2. Danach alle **Warteliste**-Teilnehmer in Reihenfolge der Registrierung
   *    → bekommen IDs N+1..N+M (Warteliste hängt lückenlos hinten an).
   * 3. Abgemeldete bekommen TeilnehmerID = null.
   *
   * Damit stehen im Teilnehmerlisten-Grid die Angemeldeten sauber oben, die
   * Warteliste sauber unten — kein Durchmischen mehr.
   */
  public async reorderParticipantIDs(subsiteUrl: string): Promise<{ success: number; errors: number }> {
    // Alle Items laden, sortiert nach SP Id (Erstellungsreihenfolge = Reihenfolge der Registrierung)
    const allItems: Array<{ Id: number; Status: string; TeilnehmerID: number | null }> = [];
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Id,Status,TeilnehmerID&$orderby=Id asc&$top=5000`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
        url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
      } catch { break; }
    }

    // Ziel-IDs in einem ersten Durchlauf berechnen: erst Angemeldete, dann Warteliste.
    const targetIds = new Map<number, number | null>();
    let nextId = 1;
    // Pass 1: Angemeldete / QR versendet / Eingecheckt
    for (const item of allItems) {
      if (item.Status === 'Angemeldet' || item.Status === 'QR versendet' || item.Status === 'Eingecheckt') {
        targetIds.set(item.Id, nextId++);
      }
    }
    // Pass 2: Warteliste
    for (const item of allItems) {
      if (item.Status === 'Warteliste') {
        targetIds.set(item.Id, nextId++);
      }
    }
    // Pass 3: Abgemeldete (TeilnehmerID=null)
    for (const item of allItems) {
      if (!targetIds.has(item.Id)) {
        targetIds.set(item.Id, null);
      }
    }

    let success = 0;
    let errors = 0;
    for (const item of allItems) {
      const newId = targetIds.get(item.Id) ?? null;
      if (newId === item.TeilnehmerID) { success++; continue; }
      try {
        const resp = await this._merge(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${item.Id})`,
          { 'TeilnehmerID': newId }
        );
        if (resp.ok || resp.status === 406) { success++; } else { errors++; }
      } catch { errors++; }
    }

    // v7.31: Counter auf den neuen Max-Wert syncen — syncCounterToMax liest
    // intern den aktuellen Max neu, damit eine parallele Anmeldung (die
    // gerade waehrend des Reorders eingetroffen ist) den Counter nicht
    // wieder auf einen niedrigen Wert ueberschreibt.
    try { await this.syncCounterToMax(subsiteUrl); } catch { /* best-effort */ }

    return { success, errors };
  }

  /**
   * Spalten der Teilnehmerliste fixen: fehlende Spalten anlegen, View-Reihenfolge korrigieren.
   * Kann auf bestehenden Events ausgefuehrt werden um die Liste nachtraeglich zu aktualisieren.
   */
  public async fixRegistrationListColumns(
    subsiteUrl: string,
    eventContext?: {
      isB2Run?: boolean;  // Event hat Durchstarter/Funstarter Kapazitaet
      hasQuiz?: boolean;  // Event hat Quizfragen
      customFields?: CustomField[]; // Event-spezifische Custom-Fields — fehlende SP-Spalten werden angelegt
    }
  ): Promise<{ added: string[]; removed: string[]; viewFixed: boolean; customFieldMap?: Record<string, string> }> {
    const added: string[] = [];
    const removed: string[] = [];

    // Bestehende Felder laden
    const fieldsResp = await this.context.spHttpClient.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=Hidden eq false&$top=200&$select=InternalName,Title`,
      SPHttpClient.configurations.v1
    );
    const existingFieldsList: string[] = [];
    if (fieldsResp.ok) {
      const fieldsData = await fieldsResp.json();
      const fields = fieldsData.value || fieldsData.d?.results || [];
      for (const f of fields) { existingFieldsList.push(f.InternalName); }
    }

    // Fehlende Basis-Spalten anlegen. StarterType/Quiz-Felder sind feature-spezifisch:
    // - StarterType/PreferredStarterType: nur fuer B2Run-Events mit Split-Kapazitaet
    // - QuizScore/QuizAnswers/QuizCompletedAt: nur fuer Events mit Quizfragen
    // Wird das Event ohne eventContext gefixt (kein Aufrufer-seitiger Flag), lassen wir
    // feature-spezifische Spalten raus, damit sie nicht unbegruendet auf jedem
    // Teilnehmerlisten-Schema auftauchen.
    const requiredFields: Array<{ title: string; type: number; choices?: string[]; metaType?: string }> = [
      { title: 'Anrede', type: 6, choices: ['Frau', 'Herr', 'Divers'], metaType: 'SP.FieldChoice' },
      { title: 'RegisteredByName', type: 2 },  // Audit: Name des Users der die Anmeldung durchgefuehrt hat
      { title: 'RegisteredByEmail', type: 2 }, // Audit: E-Mail des Users der die Anmeldung durchgefuehrt hat
      { title: 'CancelledByName', type: 2 },   // Audit: Name des Users der die Abmeldung ausgeloest hat
      { title: 'CancelledByEmail', type: 2 },  // Audit: E-Mail des Users der die Abmeldung ausgeloest hat
      { title: 'CheckedInDate', type: 4 },     // v7.16: Check-In-Audit — Zeitpunkt
      { title: 'CheckedInByName', type: 2 },   // v7.16: Check-In-Audit — Helfer-Name
      { title: 'CheckedInByEmail', type: 2 },  // v7.16: Check-In-Audit — Helfer-E-Mail
    ];
    if (eventContext?.isB2Run) {
      requiredFields.push(
        { title: 'StarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' },
        { title: 'PreferredStarterType', type: 6, choices: ['Durchstarter', 'Funstarter'], metaType: 'SP.FieldChoice' }
      );
    }
    if (eventContext?.hasQuiz) {
      requiredFields.push(
        { title: 'QuizScore', type: 9 },
        { title: 'QuizAnswers', type: 3 },
        { title: 'QuizCompletedAt', type: 4 }
      );
    }

    // Feature-spezifische Spalten, die auf diesem Event NICHT gebraucht werden,
    // aktiv loeschen (z.B. StarterType auf einem Nicht-B2Run-Event). Das ist
    // irreversibel — eventuelle Daten in diesen Spalten gehen verloren. Ist aber
    // vom User explizit gewuenscht, damit die Teilnehmerliste pro Event-Typ
    // sauber bleibt.
    const deletableFields: string[] = [];
    if (!eventContext?.isB2Run) {
      deletableFields.push('StarterType', 'PreferredStarterType');
    }
    if (!eventContext?.hasQuiz) {
      deletableFields.push('QuizScore', 'QuizAnswers', 'QuizCompletedAt');
    }
    for (const fieldName of deletableFields) {
      if (existingFieldsList.indexOf(fieldName) >= 0) {
        try {
          const delResp = await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields/getbytitle('${fieldName}')/deleteObject`,
            {}
          );
          if (delResp.ok) {
            removed.push(fieldName);
            // aus existingFieldsList rausziehen damit die View-Logik weiter unten
            // den Feldnamen nicht mehr als "noch vorhanden" betrachtet.
            const idx = existingFieldsList.indexOf(fieldName);
            if (idx >= 0) existingFieldsList.splice(idx, 1);
          }
        } catch { /* Feld konnte nicht geloescht werden - weitermachen */ }
      }
    }

    for (const f of requiredFields) {
      if (existingFieldsList.indexOf(f.title) < 0) {
        const payload: Record<string, unknown> = {
          '__metadata': { 'type': f.metaType || 'SP.Field' },
          'Title': f.title,
          'FieldTypeKind': f.type,
          'Required': false,
        };
        if (f.choices) {
          payload['Choices'] = { 'results': f.choices };
        }
        const resp = await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, payload
        );
        if (resp.ok) added.push(f.title);
      }
    }

    // Custom Fields pro Event: wenn spInternalName leer oder die Spalte fehlt,
    // jetzt anlegen. Der Aufrufer bekommt customFieldMap zurueck und kann das
    // Event-Item persistieren (spInternalName fuer jede cf.id).
    const customFieldMap: Record<string, string> = {};
    if (eventContext?.customFields && eventContext.customFields.length > 0) {
      // Post-Fix Felder-Snapshot nach Basis-Anlage
      let currentFields = [...existingFieldsList, ...added];
      for (const cf of eventContext.customFields) {
        if (!cf.label || !cf.label.trim()) continue;
        // Wenn spInternalName schon gesetzt und Feld existiert: uebernehmen.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const existingSp: string = String((cf as any).spInternalName || '');
        if (existingSp && currentFields.indexOf(existingSp) >= 0) {
          customFieldMap[cf.id] = existingSp;
          continue;
        }
        // Feld-Payload je nach Typ
        let fieldPayload: Record<string, unknown>;
        if (cf.type === 'select' && cf.options && cf.options.length > 0) {
          fieldPayload = { '__metadata': { 'type': 'SP.FieldChoice' }, 'Title': cf.label, 'FieldTypeKind': 6, 'Required': false, 'Choices': { 'results': cf.options } };
        } else if (cf.type === 'number') {
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 9, 'Required': false };
        } else if (cf.type === 'checkbox') {
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 8, 'Required': false };
        } else if (cf.type === 'user') {
          // user-Picker wird als Text gespeichert ("Name <email>").
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 3, 'Required': false };
        } else {
          fieldPayload = { '__metadata': { 'type': 'SP.Field' }, 'Title': cf.label, 'FieldTypeKind': 2, 'Required': false };
        }
        try {
          const resp = await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`, fieldPayload
          );
          if (resp.ok) {
            const createdField = await resp.json().catch(() => null);
            const internalName = createdField?.InternalName || createdField?.d?.InternalName || '';
            if (internalName) {
              customFieldMap[cf.id] = internalName;
              added.push(internalName);
              currentFields = currentFields.concat([internalName]);
            }
          }
        } catch { /* naechstes Feld */ }
      }
    }

    // Default View komplett neu aufbauen (Reihenfolge: TeilnehmerID, Anrede, Vorname, Nachname, ...)
    let viewFixed = false;
    try {
      // Alle bestehenden Felder aus der View entfernen
      await this._post(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/defaultview/viewfields/removeallviewfields`,
        {}
      );

      // Felder in gewuenschter Reihenfolge hinzufuegen. StarterType/Quiz-Spalten
      // werden nur eingebaut, wenn sie tatsaechlich auf der Liste existieren —
      // auf Nicht-B2Run- bzw. Nicht-Quiz-Events sollen sie nicht auftauchen.
      const viewFieldsCore = [
        'TeilnehmerID', 'Anrede', 'Vorname', 'Nachname', 'ParticipantEmail',
        'Department', 'Location', 'JobTitle', 'Phone',
      ];
      const viewFields: string[] = [...viewFieldsCore];
      // Post-Fix-Feldliste (bestehende + gerade hinzugefuegte) fuer die Existenz-Checks
      const postFixFields = existingFieldsList.concat(added);
      if (postFixFields.indexOf('StarterType') >= 0) viewFields.push('StarterType');
      if (postFixFields.indexOf('PreferredStarterType') >= 0) viewFields.push('PreferredStarterType');
      viewFields.push('Status', 'RegistrationDate');
      if (postFixFields.indexOf('RegisteredByName') >= 0) viewFields.push('RegisteredByName');
      if (postFixFields.indexOf('RegisteredByEmail') >= 0) viewFields.push('RegisteredByEmail');
      viewFields.push('CancellationDate');

      // Wir blenden SP-System-Spalten komplett aus (Modified, Created, ID, Type,
      // Compliance Asset Id, Retention Label, etc.) — nur funktionelle Felder der
      // App + Custom Fields kommen in die View.
      const systemBlocklist = new Set([
        'ID', '_UIVersionString', 'Edit', 'LinkTitle', 'LinkTitleNoMenu',
        'LinkFilename', 'LinkFilenameNoMenu', 'DocIcon', 'FileLeafRef',
        'Modified', 'Created', 'Editor', 'Author', 'CreatedBy', 'ModifiedBy',
        'Title', 'ParticipantName',
        'ContentType', 'ContentTypeId', 'Attachments',
        'AppAuthor', 'AppEditor', 'App Created By', 'App Modified By',
        'Type', 'ItemChildCount', 'FolderChildCount',
        'ComplianceAssetId', '_ComplianceTag', '_ComplianceTagWrittenTime',
        '_ComplianceTagUserId', 'TaxCatchAll', 'TaxCatchAllLabel',
        'SMTotalFileStreamSize', 'SMTotalSize', 'SortBehavior',
        'OData__UIVersionString', 'OData__HasCopyDestinations',
        'LastModifiedDate', 'ChangeLog', 'CustomData',
        '_CopySource', 'owshiddenversion', 'WorkflowVersion', 'WorkflowInstanceID',
        'ItemIsRecord', '_HasEncryptedContent', '_IsRecord', '_IsRecordApplied',
        'InstanceID', 'Order', 'GUID', 'FileSizeDisplay', 'MetaInfo',
        'ParentUniqueId', 'AccessPolicy', 'HasUniqueRoleAssignments',
        'Restricted', 'Type0', 'ServerUrl', 'EncodedAbsUrl', 'BaseName',
        'FileType', 'HTML_x0020_File_x0020_Type', '_EditMenuTableStart',
        '_EditMenuTableStart2', '_EditMenuTableEnd', 'PermMask',
      ]);
      // Bereits zur View hinzugefuegt — nicht doppelt anfassen
      const alreadyAdded = new Set(viewFields);
      // Kompletter Feld-Stand NACH dem Fix (bestehende + neu angelegte),
      // damit neu angelegte Custom-Fields auch in die View kommen.
      for (const fn of postFixFields) {
        if (alreadyAdded.has(fn)) continue;
        if (systemBlocklist.has(fn)) continue;
        if (fn.charAt(0) === '_') continue;
        viewFields.push(fn);
        alreadyAdded.add(fn);
      }

      for (const fn of viewFields) {
        await this._post(
          `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/defaultview/viewfields/addviewfield('${fn}')`,
          {}
        );
      }
      viewFixed = true;
    } catch {
      console.warn('[DEX] View-Reihenfolge konnte nicht gesetzt werden');
    }

    // v7.28: Counter-Liste fuer atomare TeilnehmerID-Vergabe anlegen
    // (oder seeden mit dem aktuellen Max-Wert wenn schon vorhanden).
    try {
      const counterResult = await this.ensureCounterList(subsiteUrl);
      if (counterResult.created) {
        added.push(`Counter-Liste ${COUNTER_LIST_NAME} (atomare TeilnehmerID-Vergabe, seeded mit ${counterResult.seededValue})`);
      } else if (counterResult.seededValue !== undefined) {
        added.push(`Counter-Item nachgeseedet (NextValue=${counterResult.seededValue})`);
      }
    } catch {
      console.warn('[DEX] Counter-Liste konnte nicht angelegt werden');
    }

    return { added, removed, viewFixed, customFieldMap: Object.keys(customFieldMap).length > 0 ? customFieldMap : undefined };
  }


  /**
   * Quiz-Fortschritt in die Registrierung eines Teilnehmers schreiben.
   *
   * - answers: ausgewaehlte Antwort-Indices pro Frage (Array von Arrays, weil
   *   Fragen mehrere richtige Antworten haben koennen). Unbeantwortete Fragen
   *   bleiben als leeres Array `[]` stehen, damit der Index-Offset erhalten bleibt.
   * - score: aktuell erreichte Punkte (Anzahl korrekt beantworteter Fragen).
   * - isComplete: true wenn alle Fragen beantwortet sind — dann wird auch
   *   `QuizCompletedAt` gesetzt. Andernfalls bleibt QuizCompletedAt unveraendert
   *   (null/leer), sodass der Teilnehmer als "teilweise beantwortet" gelistet wird.
   *
   * Ersetzt die frueher nur-am-Ende aufgerufene `saveQuizResult()`. Wird jetzt
   * bei jedem "Weiter"-Klick im QuizPlayer aufgerufen (Auto-Save), damit der
   * Teilnehmer beim spaeteren Wiedereintritt an derselben Stelle weitermachen kann.
   */
  public async saveQuizProgress(
    subsiteUrl: string,
    itemId: number,
    score: number,
    answers: number[][],
    isComplete: boolean
  ): Promise<boolean> {
    try {
      // Vor dem Schreiben sicherstellen, dass die Quiz-Spalten auf der
      // Teilnehmer-Liste existieren. Bei Bestandsevents (vor Quiz-Feature
      // angelegt) fehlen sie oft; _merge mit odata=nometadata schluckt
      // unbekannte Felder stumm und das Save wirkt wie "gespeichert",
      // persistiert aber nichts.
      // Silent: wenn der aktuelle User keine Manage-Lists-Permission auf
      // der Subsite hat, schlaegt das Anlegen fehl (Regular User). Dann
      // kann nur ein Admin/Organizer die Spalten anlegen — dafuer gibt's
      // die "Spalten fixen"-Funktion im Admin Center.
      await this.ensureQuizColumnsOnRegList(subsiteUrl);

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const payload: { [key: string]: any } = {
        'QuizScore': score,
        'QuizAnswers': JSON.stringify(answers),
      };
      if (isComplete) {
        payload.QuizCompletedAt = new Date().toISOString();
      }
      const resp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        payload
      );
      return resp.ok || resp.status === 406;
    } catch {
      return false;
    }
  }

  /**
   * Quiz-Spalten auf der Teilnehmer-Liste einer Event-Subsite anlegen,
   * falls sie fehlen. Idempotent und silent: bei fehlender Permission
   * kein Crash, einfach kein-op.
   */
  private async ensureQuizColumnsOnRegList(subsiteUrl: string): Promise<void> {
    try {
      const fieldsResp = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields?$filter=Hidden eq false&$top=200&$select=InternalName`,
        SPHttpClient.configurations.v1
      );
      if (!fieldsResp.ok) return;
      const fieldsData = await fieldsResp.json();
      const existing = new Set<string>(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (fieldsData.value || fieldsData.d?.results || []).map((f: any) => f.InternalName)
      );
      const required: Array<{ title: string; type: number }> = [
        { title: 'QuizScore', type: 9 },      // Number
        { title: 'QuizAnswers', type: 3 },    // Note (multiline)
        { title: 'QuizCompletedAt', type: 4 } // DateTime
      ];
      for (const f of required) {
        if (existing.has(f.title)) continue;
        try {
          await this._post(
            `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/fields`,
            {
              '__metadata': { 'type': 'SP.Field' },
              'Title': f.title,
              'FieldTypeKind': f.type,
              'Required': false,
            }
          );
          console.warn(`[DEX] ensureQuizColumnsOnRegList: ${f.title} nachgelegt auf ${subsiteUrl}`);
        } catch {
          // keine Permission -> silent. User braucht Admin der "Spalten fixen" macht.
        }
      }
    } catch { /* silent */ }
  }

  /**
   * Ersten Warteliste-Teilnehmer nachruecken: Status -> Angemeldet.
   * Wenn inheritStarterType uebergeben wird (B2Run Split-Capacity), wird dieser Typ
   * dem Nachruecker zugewiesen (er erbt den Platz des Abgemeldeten).
   *
   * Wird **client-seitig** ausgefuehrt (von der App beim Abmelden), damit der
   * Power Automate DEX_IDReorder-Flow keinen doppelten Nachrueck-Versuch macht.
   * Liefert den nachgerueckten Teilnehmer (Email + Name) zurueck fuer die
   * Nachrueck-E-Mail.
   *
   * Schutz gegen Ueberbuchung: Wenn maxParticipants gesetzt ist und die Anzahl
   * der aktuell Angemeldeten (nach der Abmeldung) >= maxParticipants ist, wird
   * NICHT nachgerueckt. Das verhindert, dass nach einer frueheren Ueberbuchung
   * der Abbruch der Abmeldung nicht zu einer weiteren Ueberbuchung fuehrt.
   */
  public async promoteFirstWaitlistItem(
    subsiteUrl: string,
    inheritStarterType?: string,
    maxParticipants?: number,
    /** Seit v6.5: Bei B2Run-Events mit getrennten Durchstarter-/Funstarter-Wartelisten
     * hier den freigewordenen Starter-Typ mitgeben — dann wird NUR der erste
     * Warteliste-Teilnehmer mit passendem PreferredStarterType nachgerückt.
     * Wenn leer: Default-Verhalten (beliebiger Warteliste-Teilnehmer). */
    onlyWithPreferredType?: string
  ): Promise<{ success: boolean; email?: string; name?: string; skippedOverbooked?: boolean }> {
    try {
      // Ueberbuchungs-Schutz: Nur nachruecken, wenn tatsaechlich ein Platz frei ist.
      // Bei unlimited (maxParticipants === 0 oder undefined) immer nachruecken.
      //
      // WICHTIG: '>' statt '>='. Die Abmeldung (Status->Abgemeldet) ist kurz vor
      // diesem Call passiert; falls SharePoint den Statuswechsel noch nicht in
      // getRegistrationCount reflektiert (stale read), wuerden wir bei einem
      // vollen Event (z.B. 128/128) mit '>=' faelschlich skippen. Mit '>' ist
      // 'registered == max' noch erlaubt (= genau ein Platz wird nachgerueckt),
      // und eine echte Ueberbuchung (401 > 128) wird weiterhin abgefangen.
      if (maxParticipants && maxParticipants > 0) {
        const counts = await this.getRegistrationCount(subsiteUrl);
        if (counts.registered > maxParticipants) {
          console.warn(`[DEX] promoteFirstWaitlistItem: skipping promotion - event is overbooked (${counts.registered}/${maxParticipants} registered).`);
          return { success: false, skippedOverbooked: true };
        }
      }

      // Ersten Warteliste-Teilnehmer finden (aelteste RegistrationDate zuerst).
      // Bei B2Run-Split-Kapazitäten: nur die passende Warteliste durchsuchen
      // (PreferredStarterType == onlyWithPreferredType).
      let filter = `Status eq 'Warteliste'`;
      if (onlyWithPreferredType) {
        const esc = onlyWithPreferredType.replace(/'/g, "''");
        filter += ` and PreferredStarterType eq '${esc}'`;
      }
      const resp = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=${encodeURIComponent(filter)}&$orderby=RegistrationDate asc&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return { success: false };
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length === 0) return { success: false };

      const firstWaiting = items[0];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mergeBody: Record<string, any> = { 'Status': 'Angemeldet' };
      if (inheritStarterType) {
        mergeBody['StarterType'] = inheritStarterType;
      }
      const mergeResp = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${firstWaiting.Id})`,
        mergeBody
      );
      if (!(mergeResp.ok || mergeResp.status === 406)) return { success: false };

      const vorname = firstWaiting.Vorname || '';
      const nachname = firstWaiting.Nachname || '';
      const name = (vorname && nachname) ? `${vorname} ${nachname}` : (firstWaiting.ParticipantName || '');
      const email = firstWaiting.ParticipantEmail || firstWaiting.Title || '';
      console.warn(`[DEX] promoteFirstWaitlistItem: promoted ${name} <${email}> (item ${firstWaiting.Id}) to Angemeldet.`);
      return { success: true, email, name };
    } catch {
      return { success: false };
    }
  }

  /**
   * Registrierung stornieren
   */
  public async cancelRegistration(
    subsiteUrl: string,
    itemId: number,
    cancelledByName?: string,
    cancelledByEmail?: string
  ): Promise<boolean> {
    try {
      // Audit: wer hat die Abmeldung ausgeloest?
      // Bei Self-Cancel = der User selbst. Bei "Teilnehmer abmelden" durch den
      // Organizer/Admin im Admin Center = der eingeloggte Organizer/Admin.
      const auditName = cancelledByName || this.context.pageContext.user.displayName || '';
      const auditEmail = (cancelledByEmail || this.context.pageContext.user.email || '').toLowerCase();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const corePayload: Record<string, any> = {
        'Status': 'Abgemeldet',
        'CancellationDate': new Date().toISOString(),
        'TeilnehmerID': null,
      };
      // Audit-Felder optional dazu - aeltere Subsites haben die Spalten evtl. noch
      // nicht (kommt erst mit Commit a10a608). Ein 400 von SP wuerde dann die
      // ganze Abmeldung blocken. Strategie: erst mit Audit-Feldern versuchen,
      // bei Misserfolg ohne sie nochmal probieren.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const fullPayload: Record<string, any> = { ...corePayload };
      if (auditName) fullPayload['CancelledByName'] = auditName;
      if (auditEmail) fullPayload['CancelledByEmail'] = auditEmail;
      const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`;
      let response = await this._merge(url, fullPayload);
      if (!response.ok && (auditName || auditEmail)) {
        // Fallback ohne Audit-Felder (Subsite-Liste hat die Spalten noch nicht)
        console.warn('[DEX] cancelRegistration with audit failed (' + response.status + '), retrying without audit fields');
        response = await this._merge(url, corePayload);
      }
      // v7.31: Counter mit aktuellem Max syncen, damit er nicht "davonrast"
      // wenn der hoechste ID-Inhaber sich abmeldet. Best-effort, blockiert
      // die Abmeldung nicht wenn's fehlschlaegt. syncCounterToMax liest den
      // Max-Wert intern frisch (race-frei gegen parallele Anmeldungen).
      if (response.ok) {
        try { await this.syncCounterToMax(subsiteUrl); } catch { /* */ }
      }
      return response.ok;
    } catch (err) {
      console.warn('[DEX] cancelRegistration error:', err);
      return false;
    }
  }

  /**
   * Teilnehmer einchecken (Status auf 'Eingecheckt' setzen).
   * v7.16: Erfasst zusaetzlich, WANN und VON WEM der Check-In ausgeloest
   * wurde (CheckedInDate / CheckedInByName / CheckedInByEmail). Diese
   * Spalten werden bei neuen Events ueber createRegistrationList() automatisch
   * angelegt; fuer bestehende Events muss der Admin einmalig die Kachel
   * "Spalten fixen" im Admin-Center klicken, damit der Check-In nicht mit
   * HTTP 400 fehlschlaegt.
   */
  public async checkInParticipant(
    subsiteUrl: string,
    itemId: number
  ): Promise<boolean> {
    try {
      const me = this.context.pageContext.user;
      const response = await this._merge(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})`,
        {
          'Status': 'Eingecheckt',
          'CheckedInDate': new Date().toISOString(),
          'CheckedInByName': me.displayName || '',
          'CheckedInByEmail': me.email || me.loginName || '',
        }
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
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$filter=ParticipantEmail eq '${email.replace(/'/g, "''")}'&$select=Id,Title,Vorname,Nachname,ParticipantName,ParticipantEmail,Status,RegistrationDate,RegisteredByName,RegisteredByEmail,CancellationDate,CancelledByName,CancelledByEmail,CustomData,Department,JobTitle,Location&$top=1`,
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
    // $top=5000 (SP-REST-Maximum) statt 500 — sonst werden bei Events mit
    // ≥500 Eintraegen die Counts auf den Event-Karten falsch berechnet,
    // weil SharePoint bei $orderby+$top mit ILS nicht zuverlaessig nextLink
    // liefert wenn die Page exakt voll ist.
    let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=Status&$top=5000`;

    while (url) {
      try {
        const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) break;
        const data = await response.json();
        allItems.push(...(data.value || data.d?.results || []));
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

  /**
   * Event-Bild als Attachment an ein DEX_Events-Item anhaengen.
   * Loescht zuerst alle bestehenden Bild-Attachments (Praefix __eventimage__),
   * dann wird das neue Bild hochgeladen. Liefert die ServerRelativeUrl als absolute URL.
   * Vorteil: keine SiteAssets-Berechtigungen noetig - wer das Item editieren darf,
   * darf auch Attachments hinzufuegen.
   */
  public async uploadEventImageAsAttachment(eventId: number, file: File): Promise<string> {
    const IMAGE_PREFIX = '__eventimage__';
    try {
      // 1. Bestehende Bild-Attachments loeschen (nur __eventimage__-Praefixe)
      try {
        const listResp = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles`,
          SPHttpClient.configurations.v1
        );
        if (listResp.ok) {
          const listData = await listResp.json();
          const files = listData.value || listData.d?.results || [];
          for (const f of files) {
            const fn: string = f.FileName || '';
            if (fn.indexOf(IMAGE_PREFIX) === 0) {
              try {
                await this._delete(
                  `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/getByFileName('${encodeURIComponent(fn)}')`
                );
              } catch { /* ignore */ }
            }
          }
        }
      } catch { /* ignore */ }

      // 2. Neues Bild hochladen mit Praefix
      const ext = (file.name.split('.').pop() || 'jpg').toLowerCase();
      const safeName = `${IMAGE_PREFIX}${Date.now().toString(36)}.${ext}`;

      const response = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/add(FileName='${encodeURIComponent(safeName)}')`,
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
        console.warn('[DEX] Image attachment upload status:', response.status);
      }

      // Fallback: URL aus bekanntem Pfad
      const serverRelUrl = this.context.pageContext.web.serverRelativeUrl;
      return `${window.location.origin}${serverRelUrl}/Lists/DEX_Events/Attachments/${eventId}/${safeName}`;
    } catch (err) {
      console.warn('[DEX] uploadEventImageAsAttachment error:', err);
    }
    return '';
  }

  /**
   * EventImageUrl-Feld eines DEX_Events-Items setzen (kleines MERGE).
   */
  public async updateEventImageUrl(eventId: number, url: string): Promise<boolean> {
    try {
      const resp = await this._merge(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
        { 'EventImageUrl': url }
      );
      return resp.ok || resp.status === 406;
    } catch {
      return false;
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
   * Dokument-Attachment von einem DEX_Events-Item loeschen.
   * Wird beim Edit verwendet, wenn der User ein bestehendes Dokument entfernt.
   */
  public async deleteEventDocument(eventId: number, fileName: string): Promise<boolean> {
    try {
      const resp = await this._delete(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`
      );
      return resp.ok || resp.status === 200 || resp.status === 204;
    } catch (err) {
      console.warn('[DEX] deleteEventDocument error:', err);
      return false;
    }
  }

  /**
   * Attachments eines DEX_Events-Items laden.
   * Bilder mit Praefix __eventimage__ werden ausgefiltert (nur fuer EventImageUrl).
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
        return files
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .filter((f: any) => (f.FileName || '').indexOf('__eventimage__') !== 0)
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .map((f: any) => ({
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
   * Permission-Check: Darf der aktuell eingeloggte User einen anderen Teilnehmer
   * registrieren? Wird in registerForEvent() und reactivateRegistration() aufgerufen,
   * wenn ParticipantEmail !== session-Email.
   *
   * Erlaubt wenn (OR):
   *   - DEX_Roles enthaelt den User als 'Admin'
   *   - Der User ist in event.OrganizerEmail fuer das Event auf der zugehoerigen
   *     Subsite eingetragen (Event-scope Organizer)
   *   - Der User ist Assistant (JobTitle enthaelt 'assistant') UND der Target
   *     ist Partner oder Director (JobTitle enthaelt 'partner' oder 'director')
   *
   * Bei Fehlern lieber konservativ `false` zurueckgeben statt durchlassen.
   */
  private async canRegisterForOthers(subsiteUrl: string, targetParticipantEmail: string): Promise<boolean> {
    const sessionEmail = (this.context.pageContext.user.email || '').toLowerCase();
    if (!sessionEmail) return false;

    // 1. DEX_Roles pruefen: Admin-Rolle haben?
    try {
      const esc = sessionEmail.replace(/'/g, "''");
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Roles')/items?$filter=Title eq '${encodeURIComponent(esc)}'&$top=1&$select=Role`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const data = await resp.json();
        const items = data.value || data.d?.results || [];
        if (items.length > 0 && items[0].Role === 'Admin') return true;
      }
    } catch { /* ignore - fallback auf weitere Checks */ }

    // 2. Event-Organizer? OrganizerEmail aus DEX_Events finden ueber SubsiteUrl-Match
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=SubsiteUrl eq '${encodeURIComponent(subsiteUrl.replace(/'/g, "''"))}'&$top=1&$select=OrganizerEmail`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const data = await resp.json();
        const items = data.value || data.d?.results || [];
        if (items.length > 0) {
          const orgStr: string = items[0].OrganizerEmail || '';
          const orgEmails = orgStr.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
          if (orgEmails.indexOf(sessionEmail) >= 0) return true;
        }
      }
    } catch { /* ignore */ }

    // 3. Assistant-Ausnahme: User-JobTitle enthaelt 'assistant' UND Target ist Partner/Director
    try {
      const sessionProfile = await this.getCurrentUserProfile();
      const sessionJt = (sessionProfile.jobTitle || '').toLowerCase();
      if (sessionJt.indexOf('assistant') >= 0) {
        const targetProfile = await this.getUserProfileByEmail(targetParticipantEmail);
        const targetJt = (targetProfile.jobTitle || '').toLowerCase();
        if (targetJt.indexOf('partner') >= 0 || targetJt.indexOf('director') >= 0) {
          return true;
        }
        // Assistant darf nicht fuer Non-Partner/Director registrieren
        return false;
      }
    } catch { /* ignore */ }

    return false;
  }

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

  /**
   * Cleanup: bei den N juengsten Teilnehmer-Eintraegen jedes Events JobTitle, Department,
   * Location und Phone aus dem aktuellen Benutzerprofil neu laden und ueberschreiben.
   * Notwendig weil bis v3.0.x diese Felder versehentlich vom EINGELOGGTEN User (statt
   * vom registrierten Teilnehmer) gezogen wurden, wenn jemand fuer eine andere Person
   * registriert hat.
   *
   * Idempotent: wenn die Daten bereits stimmen (Profil-Lookup liefert dasselbe), passiert
   * nichts. Wird typisch einmalig per LocalStorage-Flag in EventContext getriggert.
   *
   * Liefert die Anzahl tatsaechlich aktualisierter Items.
   */
  /**
   * Cleanup nur fuer EIN Event: lae alle Teilnehmer-Profile per Email nachladen
   * und JobTitle/Department/Location/Phone updaten falls abweichend.
   * Wird per Admin-Button im Admin Center pro Event getriggert.
   */
  public async fixEventParticipantsProfileData(subsiteUrl: string, n: number = 1000): Promise<{ scanned: number; updated: number; failedLookups: number }> {
    let scanned = 0;
    let updated = 0;
    let failedLookups = 0;
    const sleep = (ms: number): Promise<void> => new Promise(res => setTimeout(res, ms));
    if (!subsiteUrl) return { scanned, updated, failedLookups };
    try {
      const listResp = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items?$select=Id,ParticipantEmail,JobTitle,Department,Location,Phone&$orderby=RegistrationDate desc&$top=${n}`,
        SPHttpClient.configurations.v1
      );
      if (!listResp.ok) return { scanned, updated, failedLookups };
      const listData = await listResp.json();
      const items = listData.value || listData.d?.results || [];
      for (const it of items) {
        scanned += 1;
        const email: string = (it.ParticipantEmail || '').trim();
        if (!email) continue;
        let profile = { department: '', location: '', jobTitle: '', phone: '' };
        let success = false;
        for (let attempt = 0; attempt < 3; attempt += 1) {
          try {
            const p = await this.getUserProfileByEmail(email);
            if (p && (p.jobTitle || p.department || p.location)) {
              profile = p; success = true; break;
            }
          } catch { /* */ }
          await sleep(500 * (attempt + 1));
        }
        if (!success) { failedLookups += 1; continue; }
        try {
          const needsUpdate =
            (profile.jobTitle && profile.jobTitle !== (it.JobTitle || '')) ||
            (profile.department && profile.department !== (it.Department || '')) ||
            (profile.location && profile.location !== (it.Location || '')) ||
            (profile.phone && profile.phone !== (it.Phone || ''));
          if (needsUpdate) {
            const ok = await this._merge(
              `${subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${it.Id})`,
              {
                'JobTitle': profile.jobTitle || it.JobTitle || '',
                'Department': profile.department || it.Department || '',
                'Location': profile.location || it.Location || '',
                'Phone': profile.phone || it.Phone || '',
              }
            );
            if (ok && (ok as { ok: boolean }).ok) updated += 1;
          }
        } catch { /* */ }
        await sleep(200);
      }
    } catch { /* */ }
    return { scanned, updated, failedLookups };
  }

  public async fixRecentParticipantsProfileData(n: number): Promise<{ scanned: number; updated: number; failedLookups: number }> {
    let scanned = 0;
    let updated = 0;
    let failedLookups = 0;
    const sleep = (ms: number): Promise<void> => new Promise(res => setTimeout(res, ms));
    try {
      const events = await this.getEvents();
      for (const evt of events) {
        if (!evt.SubsiteUrl) continue;
        try {
          const listResp = await this.context.spHttpClient.get(
            `${evt.SubsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items?$select=Id,ParticipantEmail,JobTitle,Department,Location,Phone&$orderby=RegistrationDate desc&$top=${n}`,
            SPHttpClient.configurations.v1
          );
          if (!listResp.ok) continue;
          const listData = await listResp.json();
          const items = listData.value || listData.d?.results || [];
          for (const it of items) {
            scanned += 1;
            const email: string = (it.ParticipantEmail || '').trim();
            if (!email) continue;
            // Profil-Lookup mit Retry on Failure (max 3 Versuche, exponential backoff)
            let profile = { department: '', location: '', jobTitle: '', phone: '' };
            let success = false;
            for (let attempt = 0; attempt < 3; attempt += 1) {
              try {
                const p = await this.getUserProfileByEmail(email);
                if (p && (p.jobTitle || p.department || p.location)) {
                  profile = p;
                  success = true;
                  break;
                }
              } catch { /* */ }
              await sleep(500 * (attempt + 1));
            }
            if (!success) { failedLookups += 1; continue; }
            try {
              const needsUpdate =
                (profile.jobTitle && profile.jobTitle !== (it.JobTitle || '')) ||
                (profile.department && profile.department !== (it.Department || '')) ||
                (profile.location && profile.location !== (it.Location || '')) ||
                (profile.phone && profile.phone !== (it.Phone || ''));
              if (!needsUpdate) {
                await sleep(200); continue;
              }
              const ok = await this._merge(
                `${evt.SubsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${it.Id})`,
                {
                  'JobTitle': profile.jobTitle || it.JobTitle || '',
                  'Department': profile.department || it.Department || '',
                  'Location': profile.location || it.Location || '',
                  'Phone': profile.phone || it.Phone || '',
                }
              );
              if (ok && (ok as { ok: boolean }).ok) updated += 1;
            } catch { /* einzelnen ueberspringen */ }
            // Throttle gegen Rate-Limit der UserProfile-API
            await sleep(200);
          }
        } catch { /* */ }
      }
    } catch { /* */ }
    return { scanned, updated, failedLookups };
  }

  /**
   * Profildaten eines bestimmten Users via Email laden (fuer "Register for someone else"
   * und "Profile neu laden"). Robust gegen UPN != SMTP-Mismatches.
   *
   * Strategie:
   *   1. Direkter Lookup mit Claim `i:0#.f|membership|<email>` (funktioniert wenn UPN==SMTP).
   *   2. Wenn leer: per `siteusers/getbyemail` den echten LoginName aufloesen
   *      (deckt UPN != SMTP und Guest-Accounts ab) und GetPropertiesFor mit
   *      diesem LoginName erneut aufrufen.
   *
   * Rueckgabe ist gefuellt sobald einer der Wege Properties liefert, sonst leer.
   */
  public async getUserProfileByEmail(email: string): Promise<{ department: string; location: string; jobTitle: string; phone: string }> {
    const empty = { department: '', location: '', jobTitle: '', phone: '' };
    if (!email) return empty;

    const extractProfile = (props: Array<{ Key: string; Value: string }>): { department: string; location: string; jobTitle: string; phone: string } => {
      const get = (key: string): string => {
        const p = props.find(x => x.Key === key);
        return p && p.Value ? p.Value : '';
      };
      return {
        department: get('Department'),
        location: get('Office') || get('SPS-Location'),
        jobTitle: get('Title') || get('SPS-JobTitle'),
        phone: get('WorkPhone') || get('CellPhone'),
      };
    };

    // 1) Direkter Claim per SMTP-Email (schnell, funktioniert fuer Standard-Tenants)
    try {
      const directUrl = `${this.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|${encodeURIComponent(email)}'`;
      const response = await this.context.spHttpClient.get(directUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        const props: Array<{ Key: string; Value: string }> = data.UserProfileProperties || [];
        const profile = extractProfile(props);
        if (profile.jobTitle || profile.department || profile.location || profile.phone) {
          return profile;
        }
      }
    } catch { /* weiter zu Fallback */ }

    // 2) Fallback: echten LoginName (UPN-Claim) ueber siteusers/getbyemail aufloesen
    // Deckt UPN != SMTP, Guest-Accounts und Alias-SMTP-Adressen ab.
    try {
      const siteUserUrl = `${this.siteUrl}/_api/web/siteusers/getbyemail('${email.replace(/'/g, "''")}')?$select=LoginName`;
      const siteUserResp = await this.context.spHttpClient.get(siteUserUrl, SPHttpClient.configurations.v1);
      if (!siteUserResp.ok) return empty;
      const siteUserData = await siteUserResp.json();
      const loginName: string = siteUserData.LoginName || siteUserData.d?.LoginName || '';
      if (!loginName) return empty;

      const profileUrl = `${this.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(loginName)}'`;
      const profileResp = await this.context.spHttpClient.get(profileUrl, SPHttpClient.configurations.v1);
      if (!profileResp.ok) return empty;
      const profileData = await profileResp.json();
      const props: Array<{ Key: string; Value: string }> = profileData.UserProfileProperties || [];
      return extractProfile(props);
    } catch {
      return empty;
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

  // v7.28: Variante von _merge, die den uebergebenen ETag im IF-MATCH-Header
  // mitsendet (statt '*'). SharePoint vergleicht den ETag mit dem aktuellen
  // Stand des Items und antwortet mit HTTP 412 (Precondition Failed), wenn
  // ein anderer Client zwischenzeitlich geschrieben hat. So koennen wir
  // optimistic-concurrency-Pattern fuer den TeilnehmerID-Counter umsetzen.
  private async _mergeIfMatch(url: string, body: object, etag: string): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': etag,
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(body),
    };
    const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    if (response.status === 406) {
      return { ok: true, status: 204, statusText: 'No Content' } as unknown as SPHttpClientResponse;
    }
    return response;
  }

  // v7.28: Atomare TeilnehmerID-Vergabe ueber die DEX_TeilnehmerCounter-Liste
  // pro Event-Subsite. Verhindert Race-Conditions wenn viele User gleichzeitig
  // anmelden — ohne das vorher passieren konnte, dass zwei User dieselbe ID
  // bekommen (siehe Bug-Report v7.27 → v7.28).
  //
  // Ablauf:
  //   1. Counter-Item GET'en, ETag aus Response-Header lesen.
  //   2. NextValue + 1 mit IF-MATCH: <etag> via MERGE schreiben.
  //   3. Bei 412 (ETag-Mismatch = jemand war schneller) → kurzes Jitter +
  //      Retry, max 8x.
  //   4. Bei Erfolg: NextValue+1 ist die neue TeilnehmerID des aufrufenden Users.
  //
  // Fallback: Wenn die Counter-Liste nicht existiert (z.B. legacy event ohne
  // "Spalten fixen"-Lauf), kommt undefined zurueck — der Aufrufer faellt dann
  // auf das alte (race-anfaellige) max+1-Verfahren zurueck. Bestandsschutz.
  private async getNextTeilnehmerId(subsiteUrl: string): Promise<number | undefined> {
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    for (let attempt = 0; attempt < 8; attempt++) {
      let getResp: SPHttpClientResponse;
      try {
        getResp = await this.context.spHttpClient.get(counterItemUrl, SPHttpClient.configurations.v1);
      } catch {
        return undefined;
      }
      if (!getResp.ok) {
        // 404 = Counter-Liste oder Item existiert nicht → Legacy-Fallback
        if (getResp.status === 404) return undefined;
        // Andere Fehler (403 etc.) → Fallback ebenso
        return undefined;
      }
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return undefined;
      let data;
      try { data = await getResp.json(); } catch { return undefined; }
      const current = typeof data?.NextValue === 'number' ? data.NextValue : 0;
      const next = current + 1;
      const patchResp = await this._mergeIfMatch(counterItemUrl, { 'NextValue': next }, etag);
      if (patchResp.ok) return next;
      if (patchResp.status !== 412) {
        // Anderer Fehler (z.B. 500) → kein Sinn weiter zu retry'n
        return undefined;
      }
      // 412 = jemand war schneller → kurzes Jitter, dann retry
      await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
    }
    // Nach 8 Retries aufgeben → Aufrufer faellt auf max+1 zurueck.
    return undefined;
  }

  // Hilfsroutine: prueft ob auf der Counter-Liste die NextValue-Spalte
  // existiert und legt sie an wenn sie fehlt. Idempotent.
  private async ensureCounterListField(subsiteUrl: string): Promise<void> {
    const fieldUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/fields/getbytitle('NextValue')`;
    const probe = await this.context.spHttpClient.get(fieldUrl, SPHttpClient.configurations.v1);
    if (probe.ok) return;
    const fieldsUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/fields`;
    await this._post(
      fieldsUrl,
      { '__metadata': { 'type': 'SP.Field' }, 'Title': 'NextValue', 'FieldTypeKind': 9, 'Required': false }
    );
  }

  // v7.28: Counter-Liste fuer ein Event anlegen (1 Liste mit 1 Item) und
  // direkt mit dem aktuellen Max-Wert seeden — damit bestehende Events ohne
  // ID-Lueckenproduktion umsteigen koennen.
  // Idempotent: tut nichts wenn die Liste schon existiert.
  // v7.29-Fix: Item-Inserts nutzen den korrekt _x005f_-encodeten Type-Namen
  // (genauso wie wir das fuer DEX_Events machen). Vorher wurde der Listen-
  // name 1:1 in den Type uebernommen, was bei Unterstrich stillschweigend
  // zu HTTP 400 fuehrt → leere Counter-Liste.
  private async ensureCounterList(subsiteUrl: string): Promise<{ created: boolean; seededValue?: number }> {
    const probe = await this.context.spHttpClient.get(
      `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')`,
      SPHttpClient.configurations.v1
    );
    const itemsUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items`;
    const seedItem = async (): Promise<number> => {
      const maxId = await this.getCurrentMaxTeilnehmerId(subsiteUrl);
      const resp = await this._post(itemsUrl, {
        '__metadata': { 'type': COUNTER_LIST_ITEM_TYPE },
        'Title': 'TeilnehmerID',
        'NextValue': maxId,
      });
      if (!resp.ok) {
        const errBody = await resp.text().catch(() => '');
        console.warn('[DEX] ensureCounterList: Seed-Item fehlgeschlagen, status=', resp.status, errBody.substring(0, 300));
      }
      return maxId;
    };

    if (probe.ok) {
      // Liste existiert — sicherstellen dass das Schema komplett ist und ein Item drin liegt.
      try { await this.ensureCounterListField(subsiteUrl); } catch { /* */ }
      const itemListResp = await this.context.spHttpClient.get(
        `${itemsUrl}?$top=1`,
        SPHttpClient.configurations.v1
      );
      if (itemListResp.ok) {
        const data = await itemListResp.json();
        const list = data.value || data.d?.results || [];
        if (list.length > 0) return { created: false }; // alles ok
      }
      // Liste ohne Item → nachseeden
      const seededValue = await seedItem();
      return { created: false, seededValue };
    }

    // Liste neu anlegen
    await this._post(
      `${subsiteUrl}/_api/web/lists`,
      {
        '__metadata': { 'type': 'SP.List' },
        'BaseTemplate': 100,
        'Title': COUNTER_LIST_NAME,
        'Description': 'Atomarer Counter fuer TeilnehmerID-Vergabe (ETag-basiert). Nicht manuell editieren.',
        'AllowContentTypes': false,
        'ContentTypesEnabled': false,
        'EnableVersioning': false,
        'EnableMinorVersions': false,
        'OnQuickLaunch': false,
      }
    );
    await this.ensureCounterListField(subsiteUrl);
    const seededValue = await seedItem();
    return { created: true, seededValue };
  }

  // v7.28: Aktuellen Max-Wert von TeilnehmerID in der Teilnehmerliste lesen.
  // Wird beim Seeden des Counters genutzt + als Fallback wenn der Counter
  // (noch) nicht existiert.
  private async getCurrentMaxTeilnehmerId(subsiteUrl: string): Promise<number> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items?$select=TeilnehmerID&$orderby=TeilnehmerID desc&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return 0;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      if (items.length > 0 && items[0].TeilnehmerID != null) return items[0].TeilnehmerID;
    } catch { /* */ }
    return 0;
  }

  // v7.31: Counter auf den aktuellen Max-Wert syncen — wird bei Cancel und
  // nach Reorder gerufen, damit der NextValue nicht "davonrast" ohne Bezug
  // zur tatsaechlich vergebenen hoechsten ID. ETag-basiertes optimistic
  // concurrency mit Retry.
  //
  // WICHTIG: newMax wird in JEDEM Retry-Versuch neu aus der Teilnehmerliste
  // gelesen — nicht als Parameter uebergeben. Damit kann eine Anmeldung,
  // die WAEHREND des Syncs eingetroffen ist (und dabei den Counter schon
  // hochgesetzt hat), nicht durch eine falsche niedrigere Sync-Schreibung
  // ueberschrieben werden. Ohne dieses "live re-read" wuerde die Counter-
  // Vergabe spaeter den Wert kollidieren lassen mit der just-eingetragenen
  // neuen ID.
  //
  // Wir setzen nur "down": wenn der Counter eh schon <= aktuellem Max ist,
  // ist nichts zu tun (eine Anmeldung ist gleich schnell oder schneller
  // gewesen — sie hat eine hoehere ID, der Counter darf nicht runter).
  private async syncCounterToMax(subsiteUrl: string): Promise<void> {
    const counterItemUrl = `${subsiteUrl}/_api/web/lists/getbytitle('${COUNTER_LIST_NAME}')/items(1)`;
    for (let attempt = 0; attempt < 5; attempt++) {
      // Re-Read auf JEDEM Versuch: aktueller Max kommt frisch aus der
      // Teilnehmerliste, NICHT vom Aufrufer als Parameter.
      const liveMax = await this.getCurrentMaxTeilnehmerId(subsiteUrl);
      let getResp: SPHttpClientResponse;
      try {
        getResp = await this.context.spHttpClient.get(counterItemUrl, SPHttpClient.configurations.v1);
      } catch {
        return; // Counter-Liste fehlt → kein Sync moeglich, ignorieren
      }
      if (!getResp.ok) return;
      const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
      if (!etag) return;
      let data;
      try { data = await getResp.json(); } catch { return; }
      const current = typeof data?.NextValue === 'number' ? data.NextValue : 0;
      if (current <= liveMax) return; // bereits konsistent oder Anmeldung war schneller
      const patchResp = await this._mergeIfMatch(counterItemUrl, { 'NextValue': liveMax }, etag);
      if (patchResp.ok) return;
      if (patchResp.status !== 412) return;
      // 412 = jemand war zwischenzeitlich schneller (Anmeldung oder Sync) → retry
      await new Promise(res => setTimeout(res, 50 + Math.floor(Math.random() * 100)));
    }
    // Nach 5 Retries aufgeben — der Sync ist nur "Komfort", keine Pflicht.
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
