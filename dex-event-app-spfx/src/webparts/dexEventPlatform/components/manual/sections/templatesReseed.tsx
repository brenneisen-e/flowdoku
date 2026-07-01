import * as React from 'react';
import { ManualSection } from '../types';

/**
 * v13.2: Handbuch-Sektion für den Mail-Template-Reseed-Button (v12.11+).
 * Beschreibt für Admins, wann der Button sinnvoll ist und welche
 * Templates beim Reseed überschrieben werden.
 */
export function templatesReseedSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'templates-reseed',
    title: isDe ? 'Mail-Templates: Defaults zurücksetzen' : 'Mail templates: reset defaults',
    category: 'admin',
    description: isDe
      ? 'Alle Standard-Mail-Vorlagen in DEX_EmailTemplates auf die im Code definierten Default-Texte zurücksetzen — z.B. nach App-Updates.'
      : 'Reset all standard mail templates in DEX_EmailTemplates back to the defaults defined in code — e.g. after app updates.',
    visibleFor: ['Admin'],
    keywords: 'reseed re-seed zurücksetzen reset defaults default-vorlagen default-templates mail-vorlagen email templates überschreiben overwrite standard-mails admin-hub admin center vorlagen zurücksetzen dex_emailtemplates nach app-update seed neu seeden idempotent',
    perspectives: [
      {
        perspective: 'admin',
        steps: [
          {
            number: 1,
            title: isDe ? 'Wann brauche ich das?' : 'When do I need this?',
            description: (
              <>
                {isDe
                  ? 'Wenn die App in einer neuen Version mit überarbeiteten Standard-Mail-Texten ausgeliefert wurde (z.B. v12.11 Nachrücken-Mail, v12.13 Team-Mails, v13.0 Roommate/Group-Switch/Überbuchung), und du diese neuen Texte in deinem Tenant haben möchtest. Die App seeded neue Templates nur beim allerersten Setup einer Liste — bestehende Einträge bleiben sonst unangetastet.'
                  : 'When the app ships a new version with revised default mail texts (e.g. v12.11 promotion mail, v12.13 team mails, v13.0 roommate/group-switch/overbooking), and you want those new texts in your tenant. The app only seeds new templates on the very first list setup — existing entries remain untouched otherwise.'}
              </>
            ),
            tip: isDe
              ? 'In den Release Notes ist pro App-Version aufgelistet, welche Vorlagen neu hinzugekommen oder verändert wurden.'
              : 'The release notes list per app version which templates were added or modified.',
          },
          {
            number: 2,
            title: isDe ? 'Reseed-Button öffnen' : 'Open the reseed button',
            description: (
              <>
                {isDe
                  ? 'Öffne den Admin-Hub (Kachel „Admin" auf der Startseite). Scroll zum Abschnitt „E-Mails & Berichte" und dort zur Karte „Default-Mail-Vorlagen zurücksetzen". Klick auf den Button „Vorlagen zurücksetzen".'
                  : 'Open the Admin Hub (the „Admin" tile on the start page). Scroll to the „Emails & reports" section and there to the card „Reset default mail templates". Click the „Reset templates" button.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Confirm-Dialog bestätigen' : 'Confirm the dialog',
            description: (
              <>
                {isDe
                  ? 'Ein Confirm-Dialog (rot markiert) erinnert daran: alle individuellen Tenant-Anpassungen an Subject / Heading / HeadingColor / BodyHtml gehen verloren. Wenn ihr eure eigenen Mails customized habt, vorher manuell sichern (Liste DEX_EmailTemplates öffnen, Werte kopieren). Mit „Überschreiben" bestätigen.'
                  : 'A confirm dialog (marked as dangerous) reminds you: all tenant customizations to Subject / Heading / HeadingColor / BodyHtml will be lost. If you have customized your own mails, back them up first (open DEX_EmailTemplates, copy values). Confirm with "Overwrite".'}
              </>
            ),
            tip: isDe
              ? 'Der Reseed läuft pro Template idempotent: existiert das Template und weicht sein BodyHtml vom Code-Default ab, werden Subject, Heading, HeadingColor und BodyHtml auf den Default zurückgesetzt (fehlende Templates werden neu angelegt). Ist der Body schon identisch, wird der Eintrag übersprungen.'
              : 'The reseed runs idempotently per template: if a template exists and its BodyHtml differs from the code default, Subject, Heading, HeadingColor and BodyHtml are reset to the default (missing templates are created). If the body already matches, the entry is skipped.',
          },
          {
            number: 4,
            title: isDe ? 'Welche Templates werden überschrieben?' : 'Which templates get overwritten?',
            description: (
              <>
                {isDe ? (
                  <>
                    Alle Standard-Templates, jeweils in DE und EN:<br/>
                    Anmeldung, Warteliste, Abmeldung, AbmeldungAuto, Nachrücken,
                    OrgNachruecker, EventErstellt, OutlookDeclineReminder,
                    OutlookDeclineReminder_OnBehalfOf, OutlookForwardNotification,
                    OutlookDeclineDigest, TeamMemberJoined, TeamJoinRequest,
                    TeamJoinRejected, TeamLeadTransferred, TeamMemberCancelled,
                    RoommateRequest, GroupSwitchConfirmed, GroupSwitchWaitlist,
                    OverbookingApology.
                  </>
                ) : (
                  <>
                    All standard templates, each in DE and EN:<br/>
                    Anmeldung (registration), Warteliste (waitlist), Abmeldung
                    (cancellation), AbmeldungAuto (auto-cancellation), Nachrücken
                    (promotion), OrgNachruecker (organizer move-up notice),
                    EventErstellt (event-created notification), OutlookDeclineReminder,
                    OutlookDeclineReminder_OnBehalfOf, OutlookForwardNotification,
                    OutlookDeclineDigest, TeamMemberJoined, TeamJoinRequest,
                    TeamJoinRejected, TeamLeadTransferred, TeamMemberCancelled,
                    RoommateRequest, GroupSwitchConfirmed, GroupSwitchWaitlist,
                    OverbookingApology.
                  </>
                )}
              </>
            ),
          },
          {
            number: 5,
            title: isDe ? 'Beim nächsten Mail-Versand greift es' : 'Takes effect on next mail send',
            description: (
              <>
                {isDe
                  ? 'Der DEX_SEND_MAIL-Flow liest die Templates bei jedem Lauf neu. Beim nächsten gequeueten Mail nutzt er automatisch die zurückgesetzten Texte — kein App-Reload nötig.'
                  : 'The DEX_SEND_MAIL flow reads the templates fresh on every run. The next queued email automatically uses the reset texts — no app reload needed.'}
              </>
            ),
          },
        ],
      },
    ],
  };
}
