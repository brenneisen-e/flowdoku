/**
 * v32.52 — Die Standard-Mail-Vorlagen für DEX_EmailTemplates an EINER Stelle.
 *
 * Bis v32.51 standen Betreff, Überschrift und Text in `emailTemplatesList.ts`
 * dreimal: Erst-Befüllung einer neuen Liste, Nachrüsten fehlender Vorlagen und
 * „Standard-Vorlagen neu einspielen". Die Kopien waren schon auseinander-
 * gelaufen (v30.67 fand eine doppelte Zeile, v31.5 musste einen Text an zwei
 * Stellen ändern). Alle drei Wege lesen jetzt diese Liste.
 *
 * Wichtig für Änderungen: Die Vorlagen Nachruecken, OrgNachruecker,
 * AbmeldungAuto, OutlookDeclineReminder(_OnBehalfOf), OutlookForwardNotification
 * und OutlookDeclineDigest verschickt ein FLOW. Er ersetzt in Betreff und Text
 * nur die Platzhalter, die er kennt (docs/flow-jsons.md) — ein neuer
 * Platzhalter käme dort roh beim Empfänger an.
 */
import {
  OUTLOOK_DECLINE_BODY_EN, OUTLOOK_DECLINE_BODY_DE,
  OUTLOOK_DECLINE_BODY_ONBEHALF_EN, OUTLOOK_DECLINE_BODY_ONBEHALF_DE,
  OUTLOOK_FORWARD_BODY_EN, OUTLOOK_FORWARD_BODY_DE,
  OUTLOOK_DECLINE_DIGEST_BODY_EN, OUTLOOK_DECLINE_DIGEST_BODY_DE,
  NACHRUECKEN_BODY_EN, NACHRUECKEN_BODY_DE,
  ORG_NACHRUECKER_BODY_EN, ORG_NACHRUECKER_BODY_DE,
  ABMELDUNG_AUTO_BODY_EN, ABMELDUNG_AUTO_BODY_DE,
  TEAM_MEMBER_JOINED_BODY_EN, TEAM_MEMBER_JOINED_BODY_DE,
  TEAM_JOIN_REQUEST_BODY_EN, TEAM_JOIN_REQUEST_BODY_DE,
  TEAM_JOIN_REJECTED_BODY_EN, TEAM_JOIN_REJECTED_BODY_DE,
  TEAM_LEAD_TRANSFERRED_BODY_EN, TEAM_LEAD_TRANSFERRED_BODY_DE,
  TEAM_MEMBER_CANCELLED_BODY_EN, TEAM_MEMBER_CANCELLED_BODY_DE,
  ROOMMATE_REQUEST_BODY_EN, ROOMMATE_REQUEST_BODY_DE,
  GROUP_SWITCH_CONFIRMED_BODY_EN, GROUP_SWITCH_CONFIRMED_BODY_DE,
  GROUP_SWITCH_WAITLIST_BODY_EN, GROUP_SWITCH_WAITLIST_BODY_DE,
  OVERBOOK_APOLOGY_BODY_EN, OVERBOOK_APOLOGY_BODY_DE,
  ANMELDUNG_BODY_DE, ANMELDUNG_BODY_EN,
  WARTELISTE_BODY_DE, WARTELISTE_BODY_EN,
  ABMELDUNG_BODY_DE, ABMELDUNG_BODY_EN,
  EVENT_ERSTELLT_BODY_DE, EVENT_ERSTELLT_BODY_EN,
} from '../mailBodies';

export interface StandardVorlage {
  TemplateType: string;
  Language: 'DE' | 'EN';
  Subject: string;
  HeadingColor: string;
  Heading: string;
  BodyHtml: string;
}

const GRUEN = '#86bc25';
const ORANGE = '#ed8b00';
const ROT = '#da291c';
const BLAU = '#0076a8';

export const STANDARD_VORLAGEN: StandardVorlage[] = [
  // ===== Anmeldung, Warteliste, Abmeldung (App, Rahmen kommt von wrapTemplate) =====
  { TemplateType: 'Anmeldung', Language: 'DE', Subject: 'Anmeldebestätigung: {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Du bist angemeldet', BodyHtml: ANMELDUNG_BODY_DE },
  { TemplateType: 'Anmeldung', Language: 'EN', Subject: 'Registration confirmed: {{EventTitle}}', HeadingColor: GRUEN, Heading: 'You’re registered', BodyHtml: ANMELDUNG_BODY_EN },
  { TemplateType: 'Warteliste', Language: 'DE', Subject: 'Warteliste: {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Du stehst auf der Warteliste', BodyHtml: WARTELISTE_BODY_DE },
  { TemplateType: 'Warteliste', Language: 'EN', Subject: 'Waitlist: {{EventTitle}}', HeadingColor: ORANGE, Heading: 'You’re on the waitlist', BodyHtml: WARTELISTE_BODY_EN },
  { TemplateType: 'Abmeldung', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: ROT, Heading: 'Abmeldung bestätigt', BodyHtml: ABMELDUNG_BODY_DE },
  { TemplateType: 'Abmeldung', Language: 'EN', Subject: 'Cancellation confirmed: {{EventTitle}}', HeadingColor: ROT, Heading: 'Cancellation confirmed', BodyHtml: ABMELDUNG_BODY_EN },
  { TemplateType: 'EventErstellt', Language: 'DE', Subject: 'Dein Event ist angelegt: {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Event angelegt', BodyHtml: EVENT_ERSTELLT_BODY_DE },
  { TemplateType: 'EventErstellt', Language: 'EN', Subject: 'Your event has been created: {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Event created', BodyHtml: EVENT_ERSTELLT_BODY_EN },

  // ===== Flow-Vorlagen (fertig gerahmt; Platzhalter NICHT erweitern) =====
  { TemplateType: 'Nachruecken', Language: 'DE', Subject: 'Nachgerückt — du bist angemeldet: {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Du hast einen Platz!', BodyHtml: NACHRUECKEN_BODY_DE },
  { TemplateType: 'Nachruecken', Language: 'EN', Subject: 'Moved up — you’re registered: {{EventTitle}}', HeadingColor: GRUEN, Heading: 'You’ve got a spot!', BodyHtml: NACHRUECKEN_BODY_EN },
  { TemplateType: 'OrgNachruecker', Language: 'DE', Subject: 'Abmeldung mit Nachrücker: {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Abmeldung mit Nachrücker', BodyHtml: ORG_NACHRUECKER_BODY_DE },
  { TemplateType: 'OrgNachruecker', Language: 'EN', Subject: 'Cancellation with move-up: {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Cancellation with move-up', BodyHtml: ORG_NACHRUECKER_BODY_EN },
  { TemplateType: 'AbmeldungAuto', Language: 'DE', Subject: 'Abmeldebestätigung: {{EventTitle}}', HeadingColor: ROT, Heading: 'Abmeldung bestätigt', BodyHtml: ABMELDUNG_AUTO_BODY_DE },
  { TemplateType: 'AbmeldungAuto', Language: 'EN', Subject: 'Cancellation confirmed: {{EventTitle}}', HeadingColor: ROT, Heading: 'Cancellation confirmed', BodyHtml: ABMELDUNG_AUTO_BODY_EN },
  { TemplateType: 'OutlookDeclineReminder', Language: 'DE', Subject: 'Bitte prüfen: Möchtest du dich auch abmelden? {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Du hast den Outlook-Termin abgelehnt', BodyHtml: OUTLOOK_DECLINE_BODY_DE },
  { TemplateType: 'OutlookDeclineReminder', Language: 'EN', Subject: 'Please check: do you also want to cancel? {{EventTitle}}', HeadingColor: ORANGE, Heading: 'You declined the Outlook invite', BodyHtml: OUTLOOK_DECLINE_BODY_EN },
  { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'DE', Subject: 'Bitte prüfen: Anmeldung für {{EventTitle}} abmelden?', HeadingColor: ORANGE, Heading: 'Outlook-Termin in deinem Namen abgelehnt', BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_DE },
  { TemplateType: 'OutlookDeclineReminder_OnBehalfOf', Language: 'EN', Subject: 'Please check: cancel your registration for {{EventTitle}}?', HeadingColor: ORANGE, Heading: 'Outlook invite declined on your behalf', BodyHtml: OUTLOOK_DECLINE_BODY_ONBEHALF_EN },
  { TemplateType: 'OutlookForwardNotification', Language: 'DE', Subject: 'Hinweis: Termin weitergeleitet — {{EventTitle}}', HeadingColor: BLAU, Heading: 'Termin wurde weitergeleitet', BodyHtml: OUTLOOK_FORWARD_BODY_DE },
  { TemplateType: 'OutlookForwardNotification', Language: 'EN', Subject: 'Note: meeting forwarded — {{EventTitle}}', HeadingColor: BLAU, Heading: 'Meeting was forwarded', BodyHtml: OUTLOOK_FORWARD_BODY_EN },
  { TemplateType: 'OutlookDeclineDigest', Language: 'DE', Subject: '{{DeclineCount}} Absagen in Outlook, noch angemeldet — {{EventTitle}}', HeadingColor: ORANGE, Heading: 'In Outlook abgelehnt, noch angemeldet', BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_DE },
  { TemplateType: 'OutlookDeclineDigest', Language: 'EN', Subject: '{{DeclineCount}} declined in Outlook, still registered — {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Declined in Outlook, still registered', BodyHtml: OUTLOOK_DECLINE_DIGEST_BODY_EN },

  // ===== Team, Zimmer, Gruppe, Überbuchung (App, fertig gerahmt) =====
  { TemplateType: 'TeamMemberJoined', Language: 'DE', Subject: 'Neues Team-Mitglied — {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Neues Team-Mitglied', BodyHtml: TEAM_MEMBER_JOINED_BODY_DE },
  { TemplateType: 'TeamMemberJoined', Language: 'EN', Subject: 'New team member — {{EventTitle}}', HeadingColor: GRUEN, Heading: 'New team member', BodyHtml: TEAM_MEMBER_JOINED_BODY_EN },
  { TemplateType: 'TeamJoinRequest', Language: 'DE', Subject: 'Beitritts-Anfrage für dein Team — {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Team-Beitritts-Anfrage', BodyHtml: TEAM_JOIN_REQUEST_BODY_DE },
  { TemplateType: 'TeamJoinRequest', Language: 'EN', Subject: 'Request to join your team — {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Team join request', BodyHtml: TEAM_JOIN_REQUEST_BODY_EN },
  { TemplateType: 'TeamJoinRejected', Language: 'DE', Subject: 'Team-Beitritt abgelehnt — {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Team-Beitritts-Anfrage abgelehnt', BodyHtml: TEAM_JOIN_REJECTED_BODY_DE },
  { TemplateType: 'TeamJoinRejected', Language: 'EN', Subject: 'Team join request declined — {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Team join request declined', BodyHtml: TEAM_JOIN_REJECTED_BODY_EN },
  { TemplateType: 'TeamLeadTransferred', Language: 'DE', Subject: 'Neuer Team-Lead — {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Neuer Team-Lead', BodyHtml: TEAM_LEAD_TRANSFERRED_BODY_DE },
  { TemplateType: 'TeamLeadTransferred', Language: 'EN', Subject: 'New team lead — {{EventTitle}}', HeadingColor: GRUEN, Heading: 'New team lead', BodyHtml: TEAM_LEAD_TRANSFERRED_BODY_EN },
  { TemplateType: 'TeamMemberCancelled', Language: 'DE', Subject: 'Team-Mitglied abgemeldet — {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Team-Mitglied abgemeldet', BodyHtml: TEAM_MEMBER_CANCELLED_BODY_DE },
  { TemplateType: 'TeamMemberCancelled', Language: 'EN', Subject: 'Team member cancelled — {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Team member cancelled', BodyHtml: TEAM_MEMBER_CANCELLED_BODY_EN },
  { TemplateType: 'RoommateRequest', Language: 'DE', Subject: '{{RegistrantName}} möchte ein Zimmer mit dir teilen — {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Zimmerpartner-Anfrage', BodyHtml: ROOMMATE_REQUEST_BODY_DE },
  { TemplateType: 'RoommateRequest', Language: 'EN', Subject: '{{RegistrantName}} would like to share a room with you — {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Roommate request', BodyHtml: ROOMMATE_REQUEST_BODY_EN },
  { TemplateType: 'GroupSwitchConfirmed', Language: 'DE', Subject: 'Gruppenwechsel bestätigt — {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Gruppenwechsel bestätigt', BodyHtml: GROUP_SWITCH_CONFIRMED_BODY_DE },
  { TemplateType: 'GroupSwitchConfirmed', Language: 'EN', Subject: 'Group switch confirmed — {{EventTitle}}', HeadingColor: GRUEN, Heading: 'Group switch confirmed', BodyHtml: GROUP_SWITCH_CONFIRMED_BODY_EN },
  { TemplateType: 'GroupSwitchWaitlist', Language: 'DE', Subject: 'Gruppenwechsel — auf der Warteliste: {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Gruppenwechsel — auf der Warteliste', BodyHtml: GROUP_SWITCH_WAITLIST_BODY_DE },
  { TemplateType: 'GroupSwitchWaitlist', Language: 'EN', Subject: 'Group switch — on the waitlist: {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Group switch — on the waitlist', BodyHtml: GROUP_SWITCH_WAITLIST_BODY_EN },
  { TemplateType: 'OverbookingApology', Language: 'DE', Subject: 'Wichtig: Korrektur deiner Anmeldung — {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Anmeldung korrigiert', BodyHtml: OVERBOOK_APOLOGY_BODY_DE },
  { TemplateType: 'OverbookingApology', Language: 'EN', Subject: 'Important: correction of your registration — {{EventTitle}}', HeadingColor: ORANGE, Heading: 'Registration corrected', BodyHtml: OVERBOOK_APOLOGY_BODY_EN },
];
