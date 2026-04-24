import * as React from 'react';
import { ManualSection } from '../types';
import { AppPreview } from '../previews/AppPreview';
import Header from '../../Header';
import ProfilePage from '../../ProfilePage';

export function profileSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'profile',
    title: isDe ? 'Profil & Einstellungen' : 'Profile & Settings',
    category: 'general',
    description: isDe
      ? 'Profilbild, Kontakt-Infos und Sprach-Einstellung.'
      : 'Profile picture, contact info and language setting.',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'user',
        steps: [
          {
            number: 1,
            title: isDe ? 'Profil öffnen' : 'Open your profile',
            description: (
              <>
                {isDe
                  ? 'Klick auf dein Avatar-Icon oben rechts im Header. Es öffnet sich ein Pop-up mit deinen wichtigsten Daten und einem Button "Vollständiges Profil anzeigen".'
                  : 'Click your avatar in the top-right corner. A pop-up opens with your key details and a button "View full profile".'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Sprache ändern' : 'Change language',
            description: (
              <>
                {isDe
                  ? 'Die UI-Sprache wird auf der Landing-Page über den DE/EN-Toggle oben rechts gesetzt. Die Auswahl wird im Browser gespeichert (localStorage) und gilt nur für dich.'
                  : 'The UI language is set on the landing page via the DE/EN toggle in the top right. The choice is stored in your browser (localStorage) and only applies to you.'}
              </>
            ),
            tip: isDe
              ? 'Die Sprache der System-E-Mails wird vom Organizer pro Event festgelegt — nicht von deiner UI-Sprache.'
              : 'System email language is set by the organizer per event — not by your UI language.',
          },
          {
            number: 3,
            title: isDe ? 'Profildaten kommen aus dem AD' : 'Profile data comes from AD',
            description: (
              <>
                {isDe
                  ? 'Name, E-Mail, Standort und Job-Titel werden automatisch aus deinem SharePoint-User-Profil gezogen. Die App kann diese Felder nicht ändern — bitte wende dich an den Deloitte IT-Support, wenn Daten nicht stimmen.'
                  : 'Name, email, location and job title are pulled automatically from your SharePoint user profile. The app cannot modify these fields — please contact Deloitte IT support if anything is incorrect.'}
              </>
            ),
            mockup: (
              <AppPreview
                label={isDe ? 'Profilseite (echte Ansicht)' : 'Profile page (real view)'}
                role="User"
                page="profile"
                width={430}
                device="phone"
              >
                <Header />
                <ProfilePage />
              </AppPreview>
            ),
            tip: isDe
              ? 'Die Plattform fragt dein Profil live bei SharePoint ab. Änderungen im AD sind in der App meist innerhalb weniger Minuten sichtbar.'
              : 'The platform queries your profile from SharePoint live. AD changes usually appear within minutes.',
          },
        ],
      },
    ],
  };
}
