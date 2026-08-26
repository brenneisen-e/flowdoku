/**
 * v30.11 — Modularisierung Stufe 2: Thema „DEX_EmailTemplates + _Config".
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 *
 * Die Liste trägt zwei Dinge: die Mail-Vorlagen (DE/EN, pro TemplateType,
 * vom Admin-Center editierbar, Reseed/Upgrade der Standard-Templates) und
 * die _Config-Zeile als App-weiter Kleinspeicher — KPI-Cache (Teilnehmer-/
 * Event-Zähler), App-Aufrufzähler, Logos, Test-Team-Verteiler und seit
 * v30.5 die _FAConfig-Zeile (F&A-Verteiler + Protokoll im BodyHtml-JSON).
 */

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import type { EventService, ReseedSummary } from '../EventService';
import {
  OUTLOOK_DECLINE_BODY_EN,
  OUTLOOK_DECLINE_BODY_DE,
  OUTLOOK_DECLINE_BODY_ONBEHALF_EN,
  OUTLOOK_DECLINE_BODY_ONBEHALF_DE,
  OUTLOOK_FORWARD_BODY_EN,
  OUTLOOK_FORWARD_BODY_DE,
  OUTLOOK_DECLINE_DIGEST_BODY_EN,
  OUTLOOK_DECLINE_DIGEST_BODY_DE,
  NACHRUECKEN_BODY_EN,
  NACHRUECKEN_BODY_DE,
  ORG_NACHRUECKER_BODY_EN,
  ORG_NACHRUECKER_BODY_DE,
  CANCEL_BANNER_HTML,
  ABMELDUNG_AUTO_BODY_EN,
  ABMELDUNG_AUTO_BODY_DE,
  TEAM_MEMBER_JOINED_BODY_EN,
  TEAM_MEMBER_JOINED_BODY_DE,
  TEAM_JOIN_REQUEST_BODY_EN,
  TEAM_JOIN_REQUEST_BODY_DE,
  TEAM_JOIN_REJECTED_BODY_EN,
  TEAM_JOIN_REJECTED_BODY_DE,
  TEAM_LEAD_TRANSFERRED_BODY_EN,
  TEAM_LEAD_TRANSFERRED_BODY_DE,
  TEAM_MEMBER_CANCELLED_BODY_EN,
  TEAM_MEMBER_CANCELLED_BODY_DE,
  ROOMMATE_REQUEST_BODY_EN,
  ROOMMATE_REQUEST_BODY_DE,
  GROUP_SWITCH_CONFIRMED_BODY_EN,
  GROUP_SWITCH_CONFIRMED_BODY_DE,
  GROUP_SWITCH_WAITLIST_BODY_EN,
  GROUP_SWITCH_WAITLIST_BODY_DE,
  OVERBOOK_APOLOGY_BODY_EN,
  OVERBOOK_APOLOGY_BODY_DE,
} from '../mailBodies';

// ==================== DEX_EmailTemplates Liste ====================

/**
 * Email-Templates-Liste erstellen und Default-Templates einfügen.
 * Templates können pro Event überschrieben werden (im Event JSON).
 *
 * Platzhalter: {{Name}}, {{EventTitle}}, {{AppUrl}}
 */
export async function ensureEmailTemplatesList(svc: EventService): Promise<void> {
  const listName = 'DEX_EmailTemplates';
  const exists = await svc.listExists(listName);
  if (exists) {
    // Liste existiert - prüfen ob _Config Zeile und Logo-Spalten vorhanden
    await ensureEmailTemplatesConfig(svc, listName);
    // Neuere Templates nachrüsten (falls die Liste vor v3.0.27 angelegt wurde
    // und OutlookDeclineReminder noch nicht existiert)
    await ensureMissingEmailTemplates(svc, listName);
    // Standard-Templates auf aktuelle Version upgraden (uerschreibt User-Customizing!)
    // Damit Platzhalter wie {{WaitlistPosition}} bei aelteren Tenants nachgezogen werden.
    await upgradeStandardEmailTemplates(svc, listName);
    return;
  }

  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
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
    // v15.17: Subheading editierbar (vorher hart als „Event {{EventTitle}}").
    // Leer/nicht-gesetzt → Fallback im Code auf {{EventTitle}} ohne Präfix.
    { title: 'Subheading', type: 2 },
    { title: 'BodyHtml', type: 3 },
    { title: 'LogoBase64', type: 3 },           // Base64 Deloitte Logo (Deloitte_Logo.png)
    { title: 'DefaultImageBase64', type: 3 },    // Base64 Default-Bild (dex-orb.png)
  ];

  for (const f of fields) {
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
      '__metadata': { 'type': 'SP.Field' },
      'Title': f.title,
      'FieldTypeKind': f.type,
      'Required': false,
    });
  }

  // Default-Templates: DE + EN für jeden Typ
  const defaults = [
    // ===== ENGLISCH =====
    { TemplateType: 'Anmeldung', Language: 'EN', Subject: 'Registration confirmation: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Registration successful',
      BodyHtml: '<p>Dear {{Name}},</p><p>you have successfully registered for the event <strong>{{EventTitle}}</strong>.</p><p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="{{AppUrl}}">DEX App</a> (\u201EMy Events\u201C).</p><p>For organizational questions about the event, please contact {{OrganizerHtml}}.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
    { TemplateType: 'Warteliste', Language: 'EN', Subject: 'Waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Waitlist confirmation',
      BodyHtml: '<p>Dear {{Name}},</p><p>you have been placed on the <strong>waitlist</strong> for the event <strong>{{EventTitle}}</strong>.</p><p>Your current position: <strong>#{{WaitlistPosition}}</strong></p><p>We will notify you as soon as a spot becomes available. You can always check your current position in the <a href="{{AppUrl}}">DEX App</a> under \u201EMy Events\u201C.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
    { TemplateType: 'Abmeldung', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
      BodyHtml: '<p>Dear {{Name}},</p>' + CANCEL_BANNER_HTML + '<p>your registration for the event above has been <strong>cancelled</strong>. The Outlook calendar entry will be removed from your calendar shortly.</p><p>If you change your mind, you can register again via the <a href="{{AppUrl}}">DEX App</a>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
    { TemplateType: 'Nachruecken', Language: 'EN', Subject: 'You’ve got a spot: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'You’ve got a spot!',
      BodyHtml: NACHRUECKEN_BODY_EN },
    { TemplateType: 'EventErstellt', Language: 'EN', Subject: '[Deloitte Eventmanager] - New event created: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event Created',
      BodyHtml: '<p>Dear {{Name}},</p><p>your event <strong>{{EventTitle}}</strong> has been successfully created.</p><p>You can manage participants in the <a href="{{AppUrl}}">DEX App</a>.</p><p>Regards,<br>Team DEX App</p>' },
    { TemplateType: 'OutlookDeclineReminder', Language: 'EN', Subject: 'Action Required: Do you also want to cancel your registration? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'You declined the Outlook invite',
      BodyHtml: OUTLOOK_DECLINE_BODY_EN },
    { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'EN', Subject: 'Action Required: Cancel registration for {{EventTitle}}?', HeadingColor: '#ed8b00', Heading: 'Outlook invite declined on behalf',
      BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_EN },
    // ===== DEUTSCH =====
    { TemplateType: 'Anmeldung', Language: 'DE', Subject: 'Anmeldebestätigung: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Anmeldung erfolgreich',
      BodyHtml: '<p>Hallo {{Name}},</p><p>du hast dich erfolgreich für das Event <strong>{{EventTitle}}</strong> angemeldet.</p><p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="{{AppUrl}}">DEX App</a> (\u201EMeine Events\u201C) ab.</p><p>Zu organisatorischen Fragen zum Event wende dich bitte an {{OrganizerHtml}}.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
    { TemplateType: 'Warteliste', Language: 'DE', Subject: 'Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Warteliste-Bestätigung',
      BodyHtml: '<p>Hallo {{Name}},</p><p>du stehst auf der <strong>Warteliste</strong> für das Event <strong>{{EventTitle}}</strong>.</p><p>Deine aktuelle Position: <strong>#{{WaitlistPosition}}</strong></p><p>Wir benachrichtigen dich, sobald ein Platz frei wird. Deinen aktuellen Warteliste-Platz kannst du jederzeit in der <a href="{{AppUrl}}">DEX App</a> unter \u201EMeine Events\u201C sehen.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
    { TemplateType: 'Abmeldung', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
      BodyHtml: '<p>Hallo {{Name}},</p>' + CANCEL_BANNER_HTML + '<p>deine Anmeldung für das oben genannte Event wurde <strong>storniert</strong>. Der Outlook-Termin wird in Kürze aus deinem Kalender entfernt.</p><p>Du kannst dich jederzeit erneut über die <a href="{{AppUrl}}">DEX App</a> anmelden.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
    { TemplateType: 'Nachruecken', Language: 'DE', Subject: 'Du hast einen Platz: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Du hast einen Platz!',
      BodyHtml: NACHRUECKEN_BODY_DE },
    // v19.25: pre-wrapped Abmelde-Bestätigung für die Flow-getriebene
    // Auto-Abmeldung (DEX_OutlookDeclineHandler), eigener Type damit die
    // App-eigene `Abmeldung` (unwrapped) unberührt bleibt.
    { TemplateType: 'AbmeldungAuto', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
      BodyHtml: ABMELDUNG_AUTO_BODY_EN },
    { TemplateType: 'AbmeldungAuto', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
      BodyHtml: ABMELDUNG_AUTO_BODY_DE },
    // v18.63: Organizer-Benachrichtigung bei Abmeldung mit Nachrücker (vom DEX_IDReorder-Flow gequeued).
    { TemplateType: 'OrgNachruecker', Language: 'EN', Subject: 'Cancellation with waitlist move-up: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Cancellation — waitlist move-up',
      BodyHtml: ORG_NACHRUECKER_BODY_EN },
    { TemplateType: 'OrgNachruecker', Language: 'DE', Subject: 'Abmeldung mit Nachrücker: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Abmeldung — Nachrücker',
      BodyHtml: ORG_NACHRUECKER_BODY_DE },
    { TemplateType: 'EventErstellt', Language: 'DE', Subject: '[Deloitte Eventmanager] - Neues Event erstellt: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event erstellt',
      BodyHtml: '<p>Hallo {{Name}},</p><p>dein Event <strong>{{EventTitle}}</strong> wurde erfolgreich erstellt.</p><p>Du kannst die Teilnehmer in der <a href="{{AppUrl}}">DEX App</a> verwalten.</p><p>Viele Grüße,<br>Team DEX App</p>' },
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
    // v9.38: OutlookDeclineDigest — geht an Organizer nach jedem Decline mit Liste aller noch-angemeldeten Decliner.
    { TemplateType: 'OutlookDeclineDigest', Language: 'EN', Subject: 'FYI: {{DeclineCount}} attendees declined Outlook — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: attendees declined the Outlook invite',
      BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_EN },
    { TemplateType: 'OutlookDeclineDigest', Language: 'DE', Subject: 'FYI: {{DeclineCount}} Teilnehmer haben Outlook abgelehnt — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: Teilnehmer haben den Outlook-Termin abgelehnt',
      BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_DE },
    // v12.13: Team-bezogene Templates (vorher inline in EventContext.tsx).
    { TemplateType: 'TeamMemberJoined', Language: 'EN', Subject: 'New team member — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team update',
      BodyHtml: TEAM_MEMBER_JOINED_BODY_EN },
    { TemplateType: 'TeamMemberJoined', Language: 'DE', Subject: 'Neues Team-Mitglied — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Update',
      BodyHtml: TEAM_MEMBER_JOINED_BODY_DE },
    { TemplateType: 'TeamJoinRequest', Language: 'EN', Subject: 'Team join request — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team join request',
      BodyHtml: TEAM_JOIN_REQUEST_BODY_EN },
    { TemplateType: 'TeamJoinRequest', Language: 'DE', Subject: 'Team-Beitritts-Anfrage — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Beitritts-Anfrage',
      BodyHtml: TEAM_JOIN_REQUEST_BODY_DE },
    { TemplateType: 'TeamJoinRejected', Language: 'EN', Subject: 'Team join request declined — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team join request declined',
      BodyHtml: TEAM_JOIN_REJECTED_BODY_EN },
    { TemplateType: 'TeamJoinRejected', Language: 'DE', Subject: 'Team-Beitritts-Anfrage abgelehnt — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team-Beitritts-Anfrage abgelehnt',
      BodyHtml: TEAM_JOIN_REJECTED_BODY_DE },
    { TemplateType: 'TeamLeadTransferred', Language: 'EN', Subject: 'Team lead change — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team lead change',
      BodyHtml: TEAM_LEAD_TRANSFERRED_BODY_EN },
    { TemplateType: 'TeamLeadTransferred', Language: 'DE', Subject: 'Team-Lead-Wechsel — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Lead-Wechsel',
      BodyHtml: TEAM_LEAD_TRANSFERRED_BODY_DE },
    { TemplateType: 'TeamMemberCancelled', Language: 'EN', Subject: 'Team update — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team update',
      BodyHtml: TEAM_MEMBER_CANCELLED_BODY_EN },
    { TemplateType: 'TeamMemberCancelled', Language: 'DE', Subject: 'Team-Update — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team-Update',
      BodyHtml: TEAM_MEMBER_CANCELLED_BODY_DE },
    // v13.0: Restliche bisher-inline-Mails (Zimmerpartner, Gruppen-Wechsel, Überbuchung).
    { TemplateType: 'RoommateRequest', Language: 'EN', Subject: '{{RegistrantName}} selected you as roommate — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Roommate request',
      BodyHtml: ROOMMATE_REQUEST_BODY_EN },
    { TemplateType: 'RoommateRequest', Language: 'DE', Subject: '{{RegistrantName}} hat dich als Zimmerpartner gewählt — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Zimmerpartner-Anfrage',
      BodyHtml: ROOMMATE_REQUEST_BODY_DE },
    { TemplateType: 'GroupSwitchConfirmed', Language: 'EN', Subject: 'Group switch confirmed — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Group switch',
      BodyHtml: GROUP_SWITCH_CONFIRMED_BODY_EN },
    { TemplateType: 'GroupSwitchConfirmed', Language: 'DE', Subject: 'Gruppen-Wechsel bestätigt — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Gruppen-Wechsel',
      BodyHtml: GROUP_SWITCH_CONFIRMED_BODY_DE },
    { TemplateType: 'GroupSwitchWaitlist', Language: 'EN', Subject: 'Group switch — on waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Group switch — on waitlist',
      BodyHtml: GROUP_SWITCH_WAITLIST_BODY_EN },
    { TemplateType: 'GroupSwitchWaitlist', Language: 'DE', Subject: 'Gruppen-Wechsel — auf Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Gruppen-Wechsel — auf Warteliste',
      BodyHtml: GROUP_SWITCH_WAITLIST_BODY_DE },
    { TemplateType: 'OverbookingApology', Language: 'EN', Subject: 'Important: correction of your registration — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Registration corrected',
      BodyHtml: OVERBOOK_APOLOGY_BODY_EN },
    { TemplateType: 'OverbookingApology', Language: 'DE', Subject: 'Wichtig: Korrektur deiner Anmeldung — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Anmeldung korrigiert',
      BodyHtml: OVERBOOK_APOLOGY_BODY_DE },
  ];

  let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
  try {
    const typeResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
      SPHttpClient.configurations.v1
    );
    if (typeResp.ok) {
      const typeData = await typeResp.json();
      listItemType = typeData.d?.ListItemEntityTypeFullName || typeData.ListItemEntityTypeFullName || listItemType;
    }
  } catch { /* Fallback */ }

  for (const t of defaults) {
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
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

  // _Config Eintrag für Logos erstellen (Base64 muss manuell in SharePoint eingetragen werden)
  await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
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

  await svc.configureDefaultView(listName, ['TemplateType', 'Language', 'Subject', 'Heading', 'HeadingColor']);
}

/**
 * Sicherstellen dass LogoBase64/DefaultImageBase64 Spalten und _Config Zeile existieren.
 * Für Tenants wo DEX_EmailTemplates schon vor v3.0.27 angelegt wurde:
 * neuere Templates (z.B. OutlookDeclineReminder DE+EN) nachrüsten, ohne
 * bestehende zu überschreiben.
 */
async function ensureMissingEmailTemplates(svc: EventService, listName: string): Promise<void> {
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
    // v9.38: OutlookDeclineDigest — wird bei bestehenden Tenants nachgerüstet.
    { TemplateType: 'OutlookDeclineDigest', Language: 'EN', Subject: 'FYI: {{DeclineCount}} attendees declined Outlook — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: attendees declined the Outlook invite',
      BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_EN },
    { TemplateType: 'OutlookDeclineDigest', Language: 'DE', Subject: 'FYI: {{DeclineCount}} Teilnehmer haben Outlook abgelehnt — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: Teilnehmer haben den Outlook-Termin abgelehnt',
      BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_DE },
  ];

  let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
  try {
    const typeResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
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
      const checkResp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '${t.TemplateType}' and Language eq '${t.Language}'&$top=1&$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (checkResp.ok) {
        const checkData = await checkResp.json();
        const items = checkData.value || checkData.d?.results || [];
        if (items.length > 0) continue; // Schon vorhanden - nicht überschreiben
      }
      // Template fehlt - nachlegen
      await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
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
 * !! ACHTUNG !! Überschreibt User-Customizing.
 *
 * Hintergrund: Templates wie 'Warteliste' wurden über die Zeit erweitert
 * (z.B. {{WaitlistPosition}}-Platzhalter). Aelter angelegte Tenants haben
 * noch die OOTB-Version ohne diese Felder. Diese Funktion zieht den BodyHtml
 * (sowie Subject + Heading) auf den aktuellen Code-Stand nach.
 */
/**
 * v12.12: Öffentliche Re-Seed-Funktion für Admins. Stößt das Update aller
 * Standard-Templates an — überschreibt eventuelle individuelle Änderungen
 * in DEX_EmailTemplates mit den aktuellen Default-Texten aus dem Code.
 */
export async function reseedDefaultEmailTemplates(svc: EventService): Promise<ReseedSummary> {
  return upgradeStandardEmailTemplates(svc, 'DEX_EmailTemplates');
}

async function upgradeStandardEmailTemplates(svc: EventService, listName: string): Promise<ReseedSummary> {
  const summary: ReseedSummary = { created: 0, updated: 0, skipped: 0, failed: 0, errors: [] };
  const APP_URL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-EventExperiencePlatform/SitePages/DEX.aspx?env=WebView';
  void APP_URL; // Reserviert für spätere Templates die {{AppUrl}} hardcoden
  const standards = [
    // ========== EN ==========
    { TemplateType: 'Anmeldung', Language: 'EN', Subject: 'Registration confirmation: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Registration successful',
      BodyHtml: '<p>Dear {{Name}},</p><p>you have successfully registered for the event <strong>{{EventTitle}}</strong>.</p><p>If you are unable to attend, please cancel your registration as soon as possible via the <a href="{{AppUrl}}">DEX App</a> (\u201EMy Events\u201C).</p><p>For organizational questions about the event, please contact {{OrganizerHtml}}.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
    { TemplateType: 'Warteliste', Language: 'EN', Subject: 'Waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Waitlist confirmation',
      BodyHtml: '<p>Dear {{Name}},</p><p>you have been placed on the <strong>waitlist</strong> for the event <strong>{{EventTitle}}</strong>.</p><p>Your current position: <strong>#{{WaitlistPosition}}</strong></p><p>We will notify you as soon as a spot becomes available. You can always check your current position in the <a href="{{AppUrl}}">DEX App</a> under \u201EMy Events\u201C.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
    { TemplateType: 'Abmeldung', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
      BodyHtml: '<p>Dear {{Name}},</p>' + CANCEL_BANNER_HTML + '<p>your registration for the event above has been <strong>cancelled</strong>. The Outlook calendar entry will be removed from your calendar shortly.</p><p>If you change your mind, you can register again via the <a href="{{AppUrl}}">DEX App</a>.</p><p style="margin-top:24px;"><strong>Best</strong><br><br><strong>Your Event-Team</strong></p>' },
    { TemplateType: 'Nachruecken', Language: 'EN', Subject: 'You’ve got a spot: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'You’ve got a spot!',
      BodyHtml: NACHRUECKEN_BODY_EN },
    { TemplateType: 'EventErstellt', Language: 'EN', Subject: '[Deloitte Eventmanager] - New event created: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event Created',
      BodyHtml: '<p>Dear {{Name}},</p><p>your event <strong>{{EventTitle}}</strong> has been successfully created.</p><p>You can manage participants in the <a href="{{AppUrl}}">DEX App</a>.</p><p>Regards,<br>Team DEX App</p>' },
    { TemplateType: 'OutlookDeclineReminder', Language: 'EN', Subject: 'Action Required: Do you also want to cancel your registration? {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'You declined the Outlook invite',
      BodyHtml: OUTLOOK_DECLINE_BODY_EN },
    // ========== DE ==========
    { TemplateType: 'Anmeldung', Language: 'DE', Subject: 'Anmeldebestätigung: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Anmeldung erfolgreich',
      BodyHtml: '<p>Hallo {{Name}},</p><p>du hast dich erfolgreich für das Event <strong>{{EventTitle}}</strong> angemeldet.</p><p>Falls du nicht teilnehmen kannst, melde dich bitte rechtzeitig über die <a href="{{AppUrl}}">DEX App</a> (\u201EMeine Events\u201C) ab.</p><p>Zu organisatorischen Fragen zum Event wende dich bitte an {{OrganizerHtml}}.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
    { TemplateType: 'Warteliste', Language: 'DE', Subject: 'Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Warteliste-Bestätigung',
      BodyHtml: '<p>Hallo {{Name}},</p><p>du stehst auf der <strong>Warteliste</strong> für das Event <strong>{{EventTitle}}</strong>.</p><p>Deine aktuelle Position: <strong>#{{WaitlistPosition}}</strong></p><p>Wir benachrichtigen dich, sobald ein Platz frei wird. Deinen aktuellen Warteliste-Platz kannst du jederzeit in der <a href="{{AppUrl}}">DEX App</a> unter \u201EMeine Events\u201C sehen.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
    { TemplateType: 'Abmeldung', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
      BodyHtml: '<p>Hallo {{Name}},</p>' + CANCEL_BANNER_HTML + '<p>deine Anmeldung für das oben genannte Event wurde <strong>storniert</strong>. Der Outlook-Termin wird in Kürze aus deinem Kalender entfernt.</p><p>Du kannst dich jederzeit erneut über die <a href="{{AppUrl}}">DEX App</a> anmelden.</p><p style="margin-top:24px;"><strong>Viele Grüße</strong><br><br><strong>Dein Event-Team</strong></p>' },
    { TemplateType: 'Nachruecken', Language: 'DE', Subject: 'Du hast einen Platz: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Du hast einen Platz!',
      BodyHtml: NACHRUECKEN_BODY_DE },
    // v19.25: pre-wrapped Abmelde-Bestätigung für die Flow-getriebene
    // Auto-Abmeldung (DEX_OutlookDeclineHandler), eigener Type damit die
    // App-eigene `Abmeldung` (unwrapped) unberührt bleibt.
    { TemplateType: 'AbmeldungAuto', Language: 'EN', Subject: 'Cancellation confirmation: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Cancellation confirmed',
      BodyHtml: ABMELDUNG_AUTO_BODY_EN },
    { TemplateType: 'AbmeldungAuto', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: '#da291c', Heading: 'Abmeldung bestätigt',
      BodyHtml: ABMELDUNG_AUTO_BODY_DE },
    // v18.63: Organizer-Benachrichtigung bei Abmeldung mit Nachrücker (vom DEX_IDReorder-Flow gequeued).
    { TemplateType: 'OrgNachruecker', Language: 'EN', Subject: 'Cancellation with waitlist move-up: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Cancellation — waitlist move-up',
      BodyHtml: ORG_NACHRUECKER_BODY_EN },
    { TemplateType: 'OrgNachruecker', Language: 'DE', Subject: 'Abmeldung mit Nachrücker: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Abmeldung — Nachrücker',
      BodyHtml: ORG_NACHRUECKER_BODY_DE },
    { TemplateType: 'EventErstellt', Language: 'DE', Subject: '[Deloitte Eventmanager] - Neues Event erstellt: {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Event erstellt',
      BodyHtml: '<p>Hallo {{Name}},</p><p>dein Event <strong>{{EventTitle}}</strong> wurde erfolgreich erstellt.</p><p>Du kannst die Teilnehmer in der <a href="{{AppUrl}}">DEX App</a> verwalten.</p><p>Viele Grüße,<br>Team DEX App</p>' },
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
    // v9.38: OutlookDeclineDigest
    { TemplateType: 'OutlookDeclineDigest', Language: 'EN', Subject: 'FYI: {{DeclineCount}} attendees declined Outlook — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: attendees declined the Outlook invite',
      BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_EN },
    { TemplateType: 'OutlookDeclineDigest', Language: 'DE', Subject: 'FYI: {{DeclineCount}} Teilnehmer haben Outlook abgelehnt — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'FYI: Teilnehmer haben den Outlook-Termin abgelehnt',
      BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_DE },
    // v12.13: Team-Templates auch im Re-Seed-Pfad, sonst greift der Admin-
    // Reseed-Button die Texte nicht.
    { TemplateType: 'TeamMemberJoined', Language: 'EN', Subject: 'New team member — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team update',
      BodyHtml: TEAM_MEMBER_JOINED_BODY_EN },
    { TemplateType: 'TeamMemberJoined', Language: 'DE', Subject: 'Neues Team-Mitglied — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Update',
      BodyHtml: TEAM_MEMBER_JOINED_BODY_DE },
    { TemplateType: 'TeamJoinRequest', Language: 'EN', Subject: 'Team join request — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team join request',
      BodyHtml: TEAM_JOIN_REQUEST_BODY_EN },
    { TemplateType: 'TeamJoinRequest', Language: 'DE', Subject: 'Team-Beitritts-Anfrage — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Beitritts-Anfrage',
      BodyHtml: TEAM_JOIN_REQUEST_BODY_DE },
    { TemplateType: 'TeamJoinRejected', Language: 'EN', Subject: 'Team join request declined — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team join request declined',
      BodyHtml: TEAM_JOIN_REJECTED_BODY_EN },
    { TemplateType: 'TeamJoinRejected', Language: 'DE', Subject: 'Team-Beitritts-Anfrage abgelehnt — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team-Beitritts-Anfrage abgelehnt',
      BodyHtml: TEAM_JOIN_REJECTED_BODY_DE },
    { TemplateType: 'TeamLeadTransferred', Language: 'EN', Subject: 'Team lead change — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team lead change',
      BodyHtml: TEAM_LEAD_TRANSFERRED_BODY_EN },
    { TemplateType: 'TeamLeadTransferred', Language: 'DE', Subject: 'Team-Lead-Wechsel — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Team-Lead-Wechsel',
      BodyHtml: TEAM_LEAD_TRANSFERRED_BODY_DE },
    { TemplateType: 'TeamMemberCancelled', Language: 'EN', Subject: 'Team update — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team update',
      BodyHtml: TEAM_MEMBER_CANCELLED_BODY_EN },
    { TemplateType: 'TeamMemberCancelled', Language: 'DE', Subject: 'Team-Update — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Team-Update',
      BodyHtml: TEAM_MEMBER_CANCELLED_BODY_DE },
    // v13.0: Zimmerpartner, Gruppen-Wechsel, Überbuchung (vorher inline).
    { TemplateType: 'RoommateRequest', Language: 'EN', Subject: '{{RegistrantName}} selected you as roommate — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Roommate request',
      BodyHtml: ROOMMATE_REQUEST_BODY_EN },
    { TemplateType: 'RoommateRequest', Language: 'DE', Subject: '{{RegistrantName}} hat dich als Zimmerpartner gewählt — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Zimmerpartner-Anfrage',
      BodyHtml: ROOMMATE_REQUEST_BODY_DE },
    { TemplateType: 'GroupSwitchConfirmed', Language: 'EN', Subject: 'Group switch confirmed — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Group switch',
      BodyHtml: GROUP_SWITCH_CONFIRMED_BODY_EN },
    { TemplateType: 'GroupSwitchConfirmed', Language: 'DE', Subject: 'Gruppen-Wechsel bestätigt — {{EventTitle}}', HeadingColor: '#86bc25', Heading: 'Gruppen-Wechsel',
      BodyHtml: GROUP_SWITCH_CONFIRMED_BODY_DE },
    { TemplateType: 'GroupSwitchWaitlist', Language: 'EN', Subject: 'Group switch — on waitlist: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Group switch — on waitlist',
      BodyHtml: GROUP_SWITCH_WAITLIST_BODY_EN },
    { TemplateType: 'GroupSwitchWaitlist', Language: 'DE', Subject: 'Gruppen-Wechsel — auf Warteliste: {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Gruppen-Wechsel — auf Warteliste',
      BodyHtml: GROUP_SWITCH_WAITLIST_BODY_DE },
    { TemplateType: 'OverbookingApology', Language: 'EN', Subject: 'Important: correction of your registration — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Registration corrected',
      BodyHtml: OVERBOOK_APOLOGY_BODY_EN },
    { TemplateType: 'OverbookingApology', Language: 'DE', Subject: 'Wichtig: Korrektur deiner Anmeldung — {{EventTitle}}', HeadingColor: '#ed8b00', Heading: 'Anmeldung korrigiert',
      BodyHtml: OVERBOOK_APOLOGY_BODY_DE },
  ];

  let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
  try {
    const typeResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
      SPHttpClient.configurations.v1
    );
    if (typeResp.ok) {
      const typeData = await typeResp.json();
      listItemType = typeData.d?.ListItemEntityTypeFullName || typeData.ListItemEntityTypeFullName || listItemType;
    }
  } catch { /* Fallback */ }

  for (const t of standards) {
    const label = `${t.TemplateType}_${t.Language}`;
    try {
      // Bestehendes Item finden
      const checkResp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '${t.TemplateType}' and Language eq '${t.Language}'&$top=1&$select=Id,BodyHtml`,
        SPHttpClient.configurations.v1
      );
      if (!checkResp.ok) {
        summary.failed++;
        summary.errors.push(`${label}: Prüfung fehlgeschlagen (HTTP ${checkResp.status})`);
        continue;
      }
      const checkData = await checkResp.json();
      const items = checkData.value || checkData.d?.results || [];
      if (items.length === 0) {
        // existiert nicht -> anlegen (übernimmt ensureMissingEmailTemplates für einige; hier sicherheitshalber auch)
        const postResp = await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
          '__metadata': { 'type': listItemType },
          'Title': label,
          'TemplateType': t.TemplateType,
          'Language': t.Language,
          'Subject': t.Subject,
          'HeadingColor': t.HeadingColor,
          'Heading': t.Heading,
          'BodyHtml': t.BodyHtml,
        });
        // v18.66: _post wirft NICHT bei HTTP 4xx/5xx — Status explizit prüfen,
        // sonst scheitern Inserts (z.B. neue Templates wie OrgNachruecker)
        // stillschweigend und der Reseed meldet fälschlich Erfolg.
        if (postResp.ok || postResp.status === 201 || postResp.status === 204) {
          summary.created++;
        } else {
          summary.failed++;
          let detail = '';
          try { detail = (await postResp.text()).slice(0, 200); } catch { /* ignore */ }
          summary.errors.push(`${label}: Anlegen fehlgeschlagen (HTTP ${postResp.status})${detail ? ' — ' + detail : ''}`);
        }
      } else {
        // existiert -> updaten falls BodyHtml vom Default abweicht
        const item = items[0];
        if (item.BodyHtml !== t.BodyHtml) {
          const mergeResp = await svc._merge(
            `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${item.Id})`,
            {
              'Title': label,
              'TemplateType': t.TemplateType,
              'Language': t.Language,
              'Subject': t.Subject,
              'HeadingColor': t.HeadingColor,
              'Heading': t.Heading,
              'BodyHtml': t.BodyHtml,
            }
          );
          if (mergeResp.ok || mergeResp.status === 204) {
            summary.updated++;
          } else {
            summary.failed++;
            summary.errors.push(`${label}: Aktualisieren fehlgeschlagen (HTTP ${mergeResp.status})`);
          }
        } else {
          summary.skipped++;
        }
      }
    } catch (e) {
      summary.failed++;
      summary.errors.push(`${label}: ${e instanceof Error ? e.message : 'Unbekannter Fehler'}`);
    }
  }
  return summary;
}

/**
 * Wird aufgerufen wenn die Liste bereits existiert (nachträgliches Upgrade).
 */
async function ensureEmailTemplatesConfig(svc: EventService, listName: string): Promise<void> {
  try {
    // 1. Logo-Spalten nachträglich anlegen falls fehlend
    // v9.16: TestTeamEmails ergänzt — globale Liste (";"-separiert) der
    // User die Test-Events sehen + sich anmelden dürfen, auch wenn sie
    // keine Organizer/Admin-Rolle haben.
    const logoFields = [
      { title: 'LogoBase64', type: 3 },
      { title: 'DefaultImageBase64', type: 3 },
      { title: 'TestTeamEmails', type: 3 }, // Note (multi-line text), ";"-separiert
      // v11.47: App-Aufruf-Counter für die KPI-Boxen auf der LandingPage.
      // Wird pro Browser-Session genau einmal inkrementiert (Session-Guard
      // in LandingPage), ETag-CAS-Retry im incrementAppViewCount().
      { title: 'AppViewCount', type: 9 }, // Number
      // v11.52: gecachter Total-Teilnehmer-Counter für das LandingPage-KPI.
      // Live-Zählung über alle Event-Subsites war zu langsam — stattdessen
      // liest der Boot-Loader diesen einen Wert (schneller REST-Call), und
      // sobald die App fertig geladen hat, schreiben wir den frischen Wert
      // im Hintergrund zurück. Eventual consistency, für KPI-Anzeige ok.
      { title: 'TotalParticipantsCount', type: 9 }, // Number
      { title: 'TotalEventsCount', type: 9 }, // Number — analog für 'Events'
      // v15.17: Subheading-Spalte für die untere Headline-Zeile pro
      // Template (vorher hart als „Event {{EventTitle}}" im Code).
      // Leerwert = Fallback im Code auf {{EventTitle}} ohne Präfix.
      { title: 'Subheading', type: 2 }, // Single line text
    ];
    for (const f of logoFields) {
      try {
        const check = await svc._sp.get(
          `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${f.title}')`,
          SPHttpClient.configurations.v1
        );
        if (!check.ok) {
          await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
            '__metadata': { 'type': 'SP.Field' },
            'Title': f.title,
            'FieldTypeKind': f.type,
            'Required': false,
          });
        }
      } catch { /* Spalte existiert oder Fehler - ignorieren */ }
    }

    // 2. _Config Zeile prüfen und ggf. anlegen
    const configResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
      SPHttpClient.configurations.v1
    );
    if (configResp.ok) {
      const configData = await configResp.json();
      const items = configData.value || configData.d?.results || [];
      if (items.length === 0) {
        // _Config Zeile fehlt - anlegen
        let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
        try {
          const typeResp = await svc._sp.get(
            `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
            SPHttpClient.configurations.v1
          );
          if (typeResp.ok) {
            const typeData = await typeResp.json();
            listItemType = typeData.d?.ListItemEntityTypeFullName || typeData.ListItemEntityTypeFullName || listItemType;
          }
        } catch { /* Fallback */ }

        await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
          '__metadata': { 'type': listItemType },
          'Title': '_Config',
          'TemplateType': '_Config',
        });
      }
    }
  } catch (err) { console.warn('[DEX] ensureEmailTemplatesConfig fehlgeschlagen:', err); }
}

/**
 * v11.52: Gecachte KPI-Werte (TotalParticipantsCount + TotalEventsCount)
 * aus der _Config-Zeile von DEX_EmailTemplates lesen. Ein einziger REST-
 * Call, kein Subsite-Roundtrip — Boot-Loader zeigt das innerhalb von ms.
 * Liefert null bei Fehler, sonst { participants, events } mit 0 als
 * Default für leere Felder.
 */
export async function getKpiCache(svc: EventService): Promise<{ participants: number; events: number } | null> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id,TotalParticipantsCount,TotalEventsCount`,
      SPHttpClient.configurations.v1,
    );
    if (!resp.ok) return null;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    if (items.length === 0) return { participants: 0, events: 0 };
    const it = items[0];
    const pRaw = it.TotalParticipantsCount;
    const eRaw = it.TotalEventsCount;
    const p = (pRaw === null || pRaw === undefined) ? 0 : (typeof pRaw === 'number' ? pRaw : (parseInt(String(pRaw), 10) || 0));
    const e = (eRaw === null || eRaw === undefined) ? 0 : (typeof eRaw === 'number' ? eRaw : (parseInt(String(eRaw), 10) || 0));
    return { participants: p, events: e };
  } catch { return null; }
}

/**
 * v30.5: F&A-Verteiler (Fachkonzept 8.1) — eigene Zeile in
 * DEX_EmailTemplates (TemplateType '_FAConfig'), JSON im BodyHtml-Feld.
 * Gleiche Ablage wie die _Config-Zeile: EIN REST-Call, keine neue Liste,
 * und Admins können den Stand notfalls direkt in SharePoint einsehen.
 */
export async function getFAConfig(svc: EventService): Promise<{ infoRecipients: string[]; listRecipients: string[]; log: Array<{ ts: string; by: string; action: string; old?: string; neu?: string }> }> {
  const empty = { infoRecipients: [], listRecipients: [], log: [] };
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_FAConfig'&$top=1&$select=Id,BodyHtml`,
      SPHttpClient.configurations.v1,
    );
    if (!resp.ok) return empty;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    if (items.length === 0) return empty;
    const parsed = JSON.parse(items[0].BodyHtml || '{}');
    return {
      infoRecipients: Array.isArray(parsed.infoRecipients) ? parsed.infoRecipients : [],
      listRecipients: Array.isArray(parsed.listRecipients) ? parsed.listRecipients : [],
      log: Array.isArray(parsed.log) ? parsed.log : [],
    };
  } catch { return empty; }
}

export async function saveFAConfig(svc: EventService, cfg: { infoRecipients: string[]; listRecipients: string[]; log: Array<{ ts: string; by: string; action: string; old?: string; neu?: string }> }): Promise<boolean> {
  try {
    const listName = 'DEX_EmailTemplates';
    const body = JSON.stringify({ ...cfg, log: cfg.log.slice(-100) });
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=TemplateType eq '_FAConfig'&$top=1&$select=Id`,
      SPHttpClient.configurations.v1,
    );
    if (!resp.ok) return false;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    if (items.length > 0) {
      const r = await svc._merge(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${items[0].Id})`,
        { 'BodyHtml': body }
      );
      return r.ok;
    }
    // Zeile fehlt (Bestandsinstallation) — anlegen. Entity-Typ wie beim
    // Template-Seeding ermitteln, mit Fallback auf den Standardnamen.
    let listItemType = 'SP.Data.DEX_x005f_EmailTemplatesListItem';
    try {
      const typeResp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
      );
      if (typeResp.ok) {
        const typeData = await typeResp.json();
        listItemType = typeData.ListItemEntityTypeFullName || typeData.d?.ListItemEntityTypeFullName || listItemType;
      }
    } catch { /* Fallback bleibt */ }
    const create = await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
      '__metadata': { 'type': listItemType },
      'Title': '_FAConfig',
      'TemplateType': '_FAConfig',
      'BodyHtml': body,
    });
    return create.ok;
  } catch { return false; }
}

/**
 * v11.53: KPI-Counter um delta hochzählen (Anmeldung +1, Cancel -1,
 * createEvent +1, deleteEvent -N). ETag-CAS-Retry, race-safe bei 10k+
 * parallelen Usern. Liefert den neuen Wert oder null bei Fehler.
 */
export async function bumpKpiParticipants(svc: EventService, delta: number): Promise<number | null> {
  return bumpKpiField(svc, 'TotalParticipantsCount', delta);
}
export async function bumpKpiEvents(svc: EventService, delta: number): Promise<number | null> {
  return bumpKpiField(svc, 'TotalEventsCount', delta);
}

/**
 * v26.63: NUR die Events-Kennzahl der Startseite neu berechnen — allein aus
 * DEX_Events (ein Read, paginiert), OHNE die teure Subsite-Teilnehmer-
 * Schleife. Möglich, weil DEX_Events pro Zeile alles Nötige trägt: IsFictive
 * (Entwurf), EventStatus (Cancelled/Under Construction) und ParentEventId
 * (Sub-Event). Gezählt werden veröffentlichte Haupt-Events (inkl. abgelaufener).
 * Der Teilnehmer-Zählerwert bleibt unverändert erhalten. Liefert die neue
 * Events-Zahl oder null bei Fehler.
 */
export async function recomputeEventKpiOnly(svc: EventService): Promise<number | null> {
  const all = await svc.getAllEventsForKpi();
  if (all.length === 0) return null;
  const events = all.filter(e =>
    e.status !== 'Cancelled' && e.status !== 'Under Construction' && !e.isFictive && !e.parentEventId
  ).length;
  const cache = await svc.getKpiCache();
  const ok = await svc.updateKpiCache({ events, participants: cache?.participants ?? 0 });
  return ok ? events : null;
}

/**
 * v26.63: Denormalisierte Teilnehmerzahl `CurrentParticipants` am DEX_Events-
 * Item aktualisieren. Best-effort — der MERGE klappt nur für Organizer/Admins
 * (Schreibrecht auf DEX_Events); bei normalen Usern (nur Lesen) schlägt er
 * still fehl, was gewollt ist. Liefert true bei Erfolg. Kein Fehler-Throw.
 */
export async function persistCurrentParticipants(svc: EventService, eventId: number, count: number): Promise<boolean> {
  if (!Number.isFinite(eventId) || eventId <= 0 || !Number.isFinite(count) || count < 0) return false;
  try {
    const resp = await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
      { 'CurrentParticipants': count }
    );
    return resp.ok || resp.status === 406;
  } catch { return false; }
}
async function bumpKpiField(svc: EventService, field: string, delta: number): Promise<number | null> {
  if (!Number.isFinite(delta) || delta === 0) return null;
  const itemUrl = await getConfigItemUrl(svc);
  if (!itemUrl) return null;
  const MAX_RETRIES = 8;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    let getResp: SPHttpClientResponse;
    try {
      getResp = await svc._sp.get(itemUrl, SPHttpClient.configurations.v1);
    } catch { return null; }
    if (!getResp.ok) return null;
    const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
    if (!etag) return null;
    let data;
    try { data = await getResp.json(); } catch { return null; }
    const raw = data?.[field] ?? data?.d?.[field];
    const current = (raw === null || raw === undefined)
      ? 0
      : (typeof raw === 'number' ? raw : (parseInt(String(raw), 10) || 0));
    const next = Math.max(0, current + delta);
    const patchResp = await svc._mergeIfMatch(itemUrl, { [field]: next }, etag);
    if (patchResp.ok) return next;
    if (patchResp.status !== 412) return null;
    await new Promise(res => setTimeout(res, 40 + Math.floor(Math.random() * 80)));
  }
  return null;
}

/**
 * v11.52: Gecachte KPI-Werte zurückschreiben. Wird nach vollem App-Load
 * im Hintergrund aufgerufen (DexEventPlatform), damit der nächste Boot-
 * Loader frische Zahlen sieht. Best-effort, kein ETag-CAS nötig — bei
 * gleichzeitigen Schreibern gewinnt der letzte, was für KPI ok ist.
 */
export async function updateKpiCache(svc: EventService, values: { participants: number; events: number }): Promise<boolean> {
  const itemUrl = await getConfigItemUrl(svc);
  if (!itemUrl) return false;
  try {
    const resp = await svc._mergeIfMatch(itemUrl, {
      'TotalParticipantsCount': Math.max(0, Math.floor(values.participants || 0)),
      'TotalEventsCount': Math.max(0, Math.floor(values.events || 0)),
    }, '*');
    return resp.ok;
  } catch { return false; }
}

/**
 * v26.4: ALLE DEX_Events-Zeilen (paginiert, NICHT auf 100 begrenzt) — nur die
 * für die KPI nötigen Felder. getEvents() lädt aus Performance-Gründen nur die
 * 100 neuesten; für den „bisher genutzt für"-Gesamtwert brauchen wir aber
 * wirklich alle Events.
 */
export async function getAllEventsForKpi(svc: EventService): Promise<Array<{ id: number; parentEventId: string; status: string; subsiteUrl: string; isFictive: boolean }>> {
  const out: Array<{ id: number; parentEventId: string; status: string; subsiteUrl: string; isFictive: boolean }> = [];
  let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=Id,ParentEventId,EventStatus,SubsiteUrl,IsFictive&$top=5000`;
  let guard = 0;
  while (url && guard < 20) {
    guard++;
    let resp: SPHttpClientResponse;
    try { resp = await svc._sp.get(url, SPHttpClient.configurations.v1); }
    catch { break; }
    if (!resp.ok) break;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let data: any;
    try { data = await resp.json(); } catch { break; }
    const items = data.value || data.d?.results || [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    for (const it of items as any[]) {
      out.push({ id: it.Id, parentEventId: it.ParentEventId || '', status: it.EventStatus || '', subsiteUrl: it.SubsiteUrl || '', isFictive: !!it.IsFictive });
    }
    url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
  }
  return out;
}

/**
 * v26.4: KPI-Gesamtwerte über ALLE Events (nicht nur die geladenen 100):
 * participants = Summe der aktiven Anmeldungen über alle Events, die NICHT
 * abgesagt und KEIN Entwurf sind (inkl. abgelaufener/Completed + Sub-Events);
 * events = Anzahl dieser Haupt-Events (ohne Sub-Events). Best-effort,
 * sequentiell (SP-Throttling) — wird im Hintergrund 1×/Session aufgerufen.
 * Liefert null bei komplettem Fehler (dann KEIN Cache-Überschreiben).
 */
export async function getKpiTotals(svc: EventService): Promise<{ participants: number; events: number } | null> {
  const LOG = '[DEX KPI]';
  try {
    // eslint-disable-next-line no-console
    console.log(`${LOG} Recompute „bisher genutzt für" gestartet — zähle über ALLE Events (nicht nur die geladenen 100) …`);
    const all = await svc.getAllEventsForKpi();
    if (all.length === 0) {
      // eslint-disable-next-line no-console
      console.warn(`${LOG} Keine Event-Zeilen geladen — Recompute übersprungen (Cache bleibt unverändert).`);
      return null;
    }
    // Abgesagte Events + Entwürfe zählen NICHT, alles andere (Active/
    // Completed, inkl. abgelaufener) schon. v26.52 BUG-FIX: Entwürfe sind
    // seit v11.89 das IsFictive-FLAG — vorher wurde nur der Legacy-Status
    // 'Under Construction' ausgefiltert, moderne Entwürfe/Test-Events
    // zählten fälschlich mit (Events UND deren Test-Teilnehmer).
    const counted = all.filter(e => e.status !== 'Cancelled' && e.status !== 'Under Construction' && !e.isFictive);
    const events = counted.filter(e => !e.parentEventId).length;
    // eslint-disable-next-line no-console
    console.log(`${LOG} ${all.length} Event-Zeilen geladen → ${counted.length} werden gezählt (inkl. abgelaufener), davon ${events} Haupt-Events. Summiere Teilnehmer pro Subsite …`);
    let participants = 0;
    let scanned = 0;
    let failed = 0;
    for (const e of counted) {
      if (!e.subsiteUrl) continue;
      try { const c = await svc.getRegistrationCount(e.subsiteUrl); participants += c.registered; scanned++; }
      catch { failed++; /* einzelne Subsite-Fehler ignorieren — Gesamtwert bleibt best-effort */ }
    }
    // eslint-disable-next-line no-console
    console.log(`${LOG} Ergebnis über ALLE Events: ${participants} Teilnehmer / ${events} Events (${scanned} Teilnehmerlisten gezählt${failed ? `, ${failed} nicht lesbar` : ''}).`);
    return { participants, events };
  } catch (err) {
    // eslint-disable-next-line no-console
    console.warn(`${LOG} Recompute fehlgeschlagen (best-effort, Cache bleibt unverändert):`, err);
    return null;
  }
}

/**
 * v11.50: Anzahl Items in DEX_Participants (= unique User, die jemals für
 * irgendein Event angemeldet/auf Warteliste waren). Liest nur das ItemCount-
 * Metadatum der Liste, nicht alle Items — schnell und cheap. Liefert null
 * bei Fehler.
 */
export async function getParticipantsListCount(svc: EventService): Promise<number | null> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Participants')?$select=ItemCount`,
      SPHttpClient.configurations.v1,
    );
    if (!resp.ok) return null;
    const data = await resp.json();
    const raw = data?.ItemCount ?? data?.d?.ItemCount;
    if (raw === null || raw === undefined) return null;
    return typeof raw === 'number' ? raw : (parseInt(String(raw), 10) || 0);
  } catch { return null; }
}

/**
 * v11.47: Aktuellen App-Aufruf-Counter aus der _Config-Zeile von
 * DEX_EmailTemplates lesen. Liefert 0 wenn das Feld leer / nicht
 * vorhanden ist. null bei Lese-Fehler.
 */
export async function getAppViewCount(svc: EventService): Promise<number | null> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id,AppViewCount`,
      SPHttpClient.configurations.v1,
    );
    if (!resp.ok) return null;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    if (items.length === 0) return 0;
    const raw = items[0].AppViewCount;
    if (raw === null || raw === undefined) return 0;
    return typeof raw === 'number' ? raw : (parseInt(String(raw), 10) || 0);
  } catch { return null; }
}

/**
 * v11.47: App-Aufruf-Counter um 1 inkrementieren — ETag-CAS-Retry analog
 * zum reserveSeat-Muster. Liefert den neuen Wert nach Inkrement, oder
 * null bei Fehler / Retry-Erschöpfung.
 */
export async function incrementAppViewCount(svc: EventService): Promise<number | null> {
  const itemUrl = await getConfigItemUrl(svc);
  if (!itemUrl) return null;
  const MAX_RETRIES = 8;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    let getResp: SPHttpClientResponse;
    try {
      getResp = await svc._sp.get(itemUrl, SPHttpClient.configurations.v1);
    } catch { return null; }
    if (!getResp.ok) return null;
    const etag = getResp.headers.get('ETag') || getResp.headers.get('etag') || '';
    if (!etag) return null;
    let data;
    try { data = await getResp.json(); } catch { return null; }
    const raw = data?.AppViewCount ?? data?.d?.AppViewCount;
    const current = (raw === null || raw === undefined)
      ? 0
      : (typeof raw === 'number' ? raw : (parseInt(String(raw), 10) || 0));
    const next = current + 1;
    const patchResp = await svc._mergeIfMatch(itemUrl, { 'AppViewCount': next }, etag);
    if (patchResp.ok) return next;
    if (patchResp.status !== 412) return null;
    // 412 = stale ETag → kurzer Backoff + neu lesen
    await new Promise(res => setTimeout(res, 40 + Math.floor(Math.random() * 80)));
  }
  return null;
}

/**
 * v11.47: Helper — URL des _Config-Items in DEX_EmailTemplates ermitteln.
 * Liefert null, wenn die Liste/Zeile noch nicht existiert.
 */
async function getConfigItemUrl(svc: EventService): Promise<string | null> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
      SPHttpClient.configurations.v1,
    );
    if (!resp.ok) return null;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    if (items.length === 0) return null;
    const id = items[0].Id;
    return `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items(${id})`;
  } catch { return null; }
}

/**
 * Logos als Base64 in die _Config Zeile schreiben (für Power Automate Flows).
 * Lädt Deloitte_Logo.png und dex-orb.png aus SiteAssets/DEX_Logos,
 * konvertiert zu Base64 Data-URI und speichert in LogoBase64/DefaultImageBase64.
 */
export async function ensureLogosInConfig(svc: EventService): Promise<void> {
  try {
    // 1. _Config Zeile lesen
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    const configItem = items[0];
    if (!configItem) return;

    // 2. Wenn LogoBase64 schon korrekt befüllt ist (mit image/ MIME-Type), nichts tun
    if (configItem.LogoBase64 && configItem.LogoBase64.startsWith('data:image/')) return;

    // 3. Bilder aus SiteAssets laden
    const logoBase64 = await loadFileAsBase64(svc, 'DEX_Logos/Deloitte_Logo.png');
    const orbBase64 = await loadFileAsBase64(svc, 'DEX_Logos/dex-orb.png');
    if (!logoBase64 && !orbBase64) return;

    // 4. In _Config Zeile schreiben (über die getestete _post/_merge Methode)
    const configId = configItem.Id || configItem.d?.Id;
    if (!configId) return;

    const updatePayload: Record<string, unknown> = {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailTemplatesListItem' },
    };
    if (logoBase64) updatePayload['LogoBase64'] = logoBase64;
    if (orbBase64) updatePayload['DefaultImageBase64'] = orbBase64;

    await svc._sp.post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items(${configId})`,
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
export async function loadFileAsBase64(svc: EventService, path: string): Promise<string> {
  try {
    const serverRelativeUrl = svc.context.pageContext.web.serverRelativeUrl;
    const fileUrl = `${svc.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}/SiteAssets/${path}')/$value`;

    // SPHttpClient mit binaryStringResponseBody für Binary-Downloads
    const resp = await svc._sp.get(fileUrl, SPHttpClient.configurations.v1, {
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
// v9.16: Test-Team — globale ";"-separierte E-Mail-Liste, gespeichert auf
// dem _Config-Eintrag der DEX_EmailTemplates-Liste. Erlaubt nicht-Admin/
// -Organizer-Usern Test-Events zu sehen + sich anzumelden.
export async function getTestTeamEmails(svc: EventService): Promise<string[]> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=TestTeamEmails`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return [];
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    if (items.length === 0) return [];
    const raw: string = (items[0].TestTeamEmails || '').toString();
    return raw.split(/[;,\n]/).map(s => s.trim().toLowerCase()).filter(s => !!s && s.includes('@'));
  } catch { return []; }
}

export async function setTestTeamEmails(svc: EventService, emails: string[]): Promise<boolean> {
  try {
    const cleaned = (emails || []).map(s => (s || '').trim()).filter(s => !!s && s.includes('@'));
    const value = cleaned.join(';');
    // _Config-Item-ID lookup
    const lookup = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '_Config'&$top=1&$select=Id`,
      SPHttpClient.configurations.v1
    );
    if (!lookup.ok) return false;
    const data = await lookup.json();
    const items = data.value || data.d?.results || [];
    if (items.length === 0) return false;
    const itemId = items[0].Id;
    const resp = await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items(${itemId})`,
      { 'TestTeamEmails': value }
    );
    return resp.ok;
  } catch { return false; }
}

export async function getEmailTemplate(svc: EventService, templateType: string, language: string = 'EN'): Promise<{ subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string } | null> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$filter=TemplateType eq '${templateType.replace(/'/g, "''")}' and Language eq '${language.replace(/'/g, "''")}'&$select=Subject,HeadingColor,Heading,Subheading,BodyHtml&$top=1`,
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
          // v15.17: Subheading (untere Headline-Zeile, vorher hart als
          // „Event {{EventTitle}}" geschrieben) jetzt aus dem Template.
          // Leer = Fallback auf reinen EventTitle ohne „Event "-Präfix.
          subheading: items[0].Subheading || '',
          bodyHtml: items[0].BodyHtml || '',
        };
      }
    }
  } catch { /* */ }
  return null;
}

/**
 * Alle Email-Templates laden (für Event-Erstellung / Admin).
 */
export async function getAllEmailTemplates(svc: EventService): Promise<Array<{ id: number; templateType: string; language: string; subject: string; headingColor: string; heading: string; subheading: string; bodyHtml: string }>> {
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items?$select=Id,TemplateType,Language,Subject,HeadingColor,Heading,Subheading,BodyHtml&$orderby=TemplateType,Language&$top=500`,
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
        subheading: item.Subheading || '',
        bodyHtml: item.BodyHtml || '',
      }));
    }
  } catch { /* */ }
  return [];
}

/**
 * Ein globales Email-Template (DEX_EmailTemplates) aktualisieren — Admin-Tool
 * (globaler Vorlagen-Editor). Nur die übergebenen Felder werden per MERGE
 * geschrieben.
 */
export async function updateEmailTemplate(svc: EventService, id: number, fields: { subject?: string; heading?: string; subheading?: string; headingColor?: string; bodyHtml?: string }): Promise<boolean> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: Record<string, any> = {};
    if (fields.subject !== undefined) body['Subject'] = fields.subject;
    if (fields.heading !== undefined) body['Heading'] = fields.heading;
    if (fields.subheading !== undefined) body['Subheading'] = fields.subheading;
    if (fields.headingColor !== undefined) body['HeadingColor'] = fields.headingColor;
    if (fields.bodyHtml !== undefined) body['BodyHtml'] = fields.bodyHtml;
    if (Object.keys(body).length === 0) return true;
    const resp = await svc._merge(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_EmailTemplates')/items(${id})`, body);
    return resp.ok;
  } catch { return false; }
}


