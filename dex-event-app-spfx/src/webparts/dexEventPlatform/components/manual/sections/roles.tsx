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
                  ? 'In den Einstellungen (Zahnrad-Icon) → "Rollen-Verwaltung". Du siehst alle Nutzer, die explizit Organizer oder Admin sind. Nutzer ohne Eintrag sind automatisch "User".'
                  : 'In Settings (gear icon) → "Role management". You see all users who are explicitly Organizer or Admin. Users without an entry are automatically "User".'}
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
                  ? 'Per "Neue Rolle" öffnest du einen Dialog. Name + E-Mail via Personen-Suche wählen, Rolle festlegen (Organizer oder Admin), optional Standort setzen. Speichern.'
                  : 'Click "Add role" to open a dialog. Pick name + email via the person search, set role (Organizer or Admin), optionally set a location. Save.'}
              </>
            ),
          },
          {
            number: 3,
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
            number: 4,
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
