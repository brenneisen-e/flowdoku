import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';
import { AppPreview } from '../previews/AppPreview';
import Header from '../../Header';
import SettingsPage from '../../SettingsPage';

export function rolesSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'roles',
    title: isDe ? 'Rollen verwalten' : 'Manage roles',
    category: 'admin',
    description: isDe
      ? 'User zu Organizern oder Admins machen — oder Rollen entziehen.'
      : 'Promote users to Organizer or Admin — or revoke roles.',
    visibleFor: ['Admin'],
    keywords: isDe
      ? 'Rollen Rechte Berechtigungen Permissions IT-Admin ITAdmin Admin Organizer User Power-User Poweruser Ticket-Beantworter Rollenverwaltung Rollenmatrix Onboarding-Mail Begrüßungsmail Cc DEX_Roles'
      : 'roles rights permissions IT-Admin ITAdmin Admin Organizer User Power-User poweruser ticket answerer role management role matrix onboarding email welcome mail Cc DEX_Roles',
    perspectives: [
      {
        perspective: 'admin',
        steps: [
          {
            number: 1,
            title: isDe ? 'Rollen-Liste öffnen' : 'Open the roles list',
            description: (
              <>
                {isDe
                  ? 'In den Einstellungen (Zahnrad-Icon) → "Rollen-Verwaltung". Du siehst alle Nutzer, die explizit Organizer, Admin oder IT-Admin sind. Nutzer ohne Eintrag sind automatisch "User".'
                  : 'In Settings (gear icon) → "Role management". You see all users who are explicitly Organizer, Admin or IT-Admin. Users without an entry are automatically "User".'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Settings → Rollenverwaltung (echte Ansicht)' : 'Settings → role management (real view)'}
                role="Admin"
                page="settings"
                width={1024}
                device="laptop"
              >
                <Header />
                <SettingsPage />
              </AppPreview>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Rolle hinzufügen' : 'Add a role',
            description: (
              <>
                {isDe
                  ? 'Per "Neue Rolle" öffnest du einen Dialog. Name + E-Mail via Personen-Suche wählen, Rolle festlegen (Organizer, Admin oder IT-Admin), optional Standort setzen. Speichern.'
                  : 'Click "Add role" to open a dialog. Pick name + email via the person search, set role (Organizer, Admin or IT-Admin), optionally set a location. Save.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Das Rollen-Modell verstehen' : 'Understand the role model',
            description: (
              <>
                {isDe
                  ? 'Es gibt vier Rollen und ein Zusatz-Flag:'
                  : 'There are four roles and one add-on flag:'}
                <ul style={{ paddingLeft: 18, margin: '8px 0 0 0', lineHeight: 1.7 }}>
                  <li><strong>User</strong> — {isDe ? 'Standard. Anmelden, eigene Events verwalten. Kein Eintrag in DEX_Roles nötig.' : 'default. Register and manage own registrations. No DEX_Roles entry needed.'}</li>
                  <li><strong>Organizer</strong> — {isDe ? 'darf Events anlegen und die eigenen Events verwalten.' : 'may create events and manage their own events.'}</li>
                  <li><strong>Admin</strong> — {isDe ? 'volle Rechte inkl. Rollenverwaltung, Admin Center für alle Events und alle Benachrichtigungs-Mails (Ticket-Fragen, Wochenreport, Organizer-Anfragen, Inaktivitäts-Hinweise).' : 'full rights incl. role management, admin center for all events and all notification mails (ticket questions, weekly report, organizer requests, inactivity notices).'}</li>
                  <li><strong>IT-Admin</strong> — {isDe ? 'exakt dieselben App-Rechte wie ein Admin (isAdmin=true), erhält aber KEINE der Benachrichtigungs-Mails. Gedacht für technische Betreuer, die vollen Zugriff brauchen, aber nicht im operativen Mail-Verteiler stehen sollen. Für IT-Admin (und reine User) wird KEINE Onboarding-Mail angeboten.' : 'exactly the same app rights as an Admin (isAdmin=true), but receives NONE of the notification mails. Meant for technical maintainers who need full access without being on the operational mail distribution. No onboarding mail is offered for IT-Admin (or plain User).'}</li>
                  <li><strong>Power-User</strong> {isDe ? '(Flag)' : '(flag)'} — {isDe ? 'ein Stern-Toggle, den du auf einem Organizer oder Admin setzen kannst. Power-User (und Admins) beantworten die Fragen aus dem Ticket-System über die „Tickets"-Kachel. Das Flag ist eine Ergänzung, keine eigene Rolle — eine Person hat genau EINEN Rollen-Eintrag und kann gleichzeitig Organizer UND Power-User sein.' : 'a star toggle you can set on an Organizer or Admin. Power-Users (and Admins) answer ticket-system questions via the „Tickets" tile. The flag is an add-on, not its own role — a person has exactly ONE role entry and can be both Organizer AND Power-User.'}</li>
                </ul>
              </>
            ),
            tip: isDe
              ? 'IT-Admin unterscheidet sich vom Admin NUR beim Mailversand: alle Empfänger-Listen filtern exakt auf Role=Admin, wodurch ein IT-Admin automatisch nicht angeschrieben wird.'
              : 'IT-Admin differs from Admin ONLY in mailing: every recipient list filters exactly on Role=Admin, so an IT-Admin is automatically left out.',
          },
          {
            number: 4,
            title: isDe ? 'Onboarding-Mail verschicken' : 'Send onboarding email',
            description: (
              <>
                {isDe
                  ? 'Sobald du eine Person als Organizer oder Admin gespeichert hast, fragt dich die App, ob du eine Onboarding-Mail verschicken willst. Die Mail kommt im Deloitte-Layout, enthält Links zur App und zum Handbuch sowie eine kurze Anleitung zum ersten Test-Event. ebrenneisen@deloitte.de und nifelten@deloitte.de stehen automatisch im Cc, damit das DEX-Team weiß, wer neu dazugekommen ist. Bei reinen "User"- und IT-Admin-Rollen wird die Mail nicht angeboten — die Begrüßung ist nur für Organizer/Admins relevant.'
                  : 'As soon as you save someone as Organizer or Admin, the app asks whether to send an onboarding email. It comes in the Deloitte layout with links to the app and the handbook plus a short guide for the first test event. ebrenneisen@deloitte.de and nifelten@deloitte.de are automatically in Cc so the DEX team knows who joined. The dialog is skipped for plain "User" and IT-Admin roles — the greeting only makes sense for Organizers/Admins.'}
              </>
            ),
            mockup: <Callout variant="tip">{isDe ? 'Du kannst die Mail jederzeit auch ablehnen — sie ist optional.' : 'You can always decline — the email is optional.'}</Callout>,
          },
          {
            number: 5,
            title: isDe ? 'Rolle ändern oder entfernen' : 'Change or remove a role',
            description: (
              <>
                {isDe
                  ? 'Per Klick auf "Ändern" kannst du die Rolle umstellen oder den Eintrag löschen. Entferntst du einen Admin-Eintrag, wird der Nutzer wieder normaler User (kein Zugriff mehr auf Admin-Bereich).'
                  : 'Click "Edit" to switch the role or delete the entry. Removing an Admin entry reverts that user to a normal User (losing access to the admin area).'}
              </>
            ),
            warning: isDe
              ? 'Entferne dich NICHT selbst als einzigen Admin — danach kommt niemand mehr an die Rollenverwaltung. Immer mindestens 2 Admins einplanen.'
              : 'Do NOT remove yourself as the sole Admin — nobody would be able to manage roles after that. Always keep at least 2 Admins.',
          },
          {
            number: 6,
            title: isDe ? 'Rollen-Matrix einsehen' : 'View the role matrix',
            description: (
              <>
                {isDe
                  ? 'Als Admin kannst du die "Rollen-Matrix" öffnen — eine Übersicht aller Features und welche Rolle welche davon sehen/nutzen darf. Gute Referenz, wenn du dich fragst "darf ein Organizer das?".'
                  : 'As Admin you can open the "Role matrix" — an overview of all features and which role may see/use them. Great reference when wondering "can an Organizer do this?".'}
              </>
            ),
            mockup: <Callout variant="tip">{isDe ? 'Die Matrix ist die einzige Quelle der Wahrheit für Permissions.' : 'The matrix is the single source of truth for permissions.'}</Callout>,
          },
        ],
      },
    ],
  };
}
