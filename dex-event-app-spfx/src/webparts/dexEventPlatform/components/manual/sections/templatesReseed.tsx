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
                  ? 'Geh in das Admin-Panel (Header oben rechts → Avatar → „Rollenverwaltung"). Scroll runter zur Karte „Default-Mail-Templates re-seed". Klick auf den grauen Button „Default-Templates zurücksetzen".'
                  : 'Open the Admin Panel (top-right header → avatar → "Role management"). Scroll down to the card "Default email templates reset". Click the grey button "Reset default templates".'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Confirm-Dialog bestätigen' : 'Confirm the dialog',
            description: (
              <>
                {isDe
                  ? 'Ein Confirm-Dialog erinnert daran: alle individuellen Tenant-Anpassungen an Subject / Heading / BodyHtml gehen verloren. Wenn ihr eure eigenen Mails customized habt, vorher manuell sichern (Liste DEX_EmailTemplates öffnen, Werte kopieren). Mit „OK" bestätigen.'
                  : 'A confirm dialog reminds you: all tenant customizations to Subject / Heading / BodyHtml will be lost. If you have customized your own mails, back them up first (open DEX_EmailTemplates, copy values). Confirm with "OK".'}
              </>
            ),
            tip: isDe
              ? 'Der Reseed läuft pro Template idempotent: existiert das Template, wird BodyHtml/Subject/Heading überschrieben — Reihenfolge oder andere Felder bleiben.'
              : 'The reseed runs idempotently per template: if a template exists, BodyHtml/Subject/Heading get overwritten — order or other fields stay.',
          },
          {
            number: 4,
            title: isDe ? 'Welche Templates werden überschrieben?' : 'Which templates get overwritten?',
            description: (
              <>
                {isDe ? (
                  <>
                    Alle Standard-Templates, sowohl DE als auch EN:<br/>
                    Anmeldung, Warteliste, Abmeldung, Nachrücken, EventErstellt,
                    OutlookDeclineReminder, OutlookDeclineReminder_OnBehalfOf,
                    OutlookForwardNotification, OutlookDeclineDigest,
                    TeamMemberJoined, TeamJoinRequest, TeamJoinRejected,
                    TeamLeadTransferred, TeamMemberCancelled, RoommateRequest,
                    GroupSwitchConfirmed, GroupSwitchWaitlist, OverbookingApology.
                  </>
                ) : (
                  <>
                    All standard templates, both DE and EN:<br/>
                    Anmeldung (registration), Warteliste (waitlist), Abmeldung
                    (cancellation), Nachrücken (promotion), EventErstellt
                    (event-created notification), OutlookDeclineReminder,
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
