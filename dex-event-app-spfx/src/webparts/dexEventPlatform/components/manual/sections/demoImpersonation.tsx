import * as React from 'react';
import { ManualSection } from '../types';

/**
 * v13.1: Handbuch-Sektion für die Admin-Demo-Impersonation (v12.7).
 * Beschreibt für Admins, wie sie die App durch die Brille eines beliebigen
 * Users testen können — inklusive Standortfilter, Permission-Caveats und
 * dem orangen Sticky-Banner zum Beenden.
 */
export function demoImpersonationSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'demo-impersonation',
    title: isDe ? 'Demo-Modus: als User testen' : 'Demo mode: act as a user',
    category: 'admin',
    description: isDe
      ? 'Als Admin testweise als beliebiger User mit gewähltem Standort durch die App navigieren — ohne dich selbst auszuloggen.'
      : 'As an admin, act as any chosen user with a selected location to test the app — without logging out.',
    visibleFor: ['Admin'],
    perspectives: [
      {
        perspective: 'admin',
        steps: [
          {
            number: 1,
            title: isDe ? 'Demo-Modus aktivieren' : 'Activate demo mode',
            description: (
              <>
                {isDe
                  ? 'Klick oben rechts auf dein Avatar-Icon → unter „Rollenverwaltung" findest du den Button „Demo: als User testen". Ein Modal öffnet sich mit People-Picker und Standort-Eingabe.'
                  : 'Click your avatar in the top-right corner → under "Role management" you\'ll find a button "Demo: act as a user". A modal opens with a people picker and a location input.'}
              </>
            ),
            tip: isDe
              ? 'Der Demo-Button erscheint nur für echte Admins — ein als-User-impersonierter Admin sieht ihn weiterhin (originalIsAdmin), kann den Modus also wieder verlassen.'
              : 'The demo button only appears for real admins — an impersonated admin still sees it (originalIsAdmin), so they can exit the mode.',
          },
          {
            number: 2,
            title: isDe ? 'Zielperson + Standort wählen' : 'Pick target person + location',
            description: (
              <>
                {isDe
                  ? 'Suche per People-Picker eine Person aus dem Tenant. Gib zusätzlich den Standort an (Pflicht) — der steuert den Standortfilter aus Schritt 3 des Event-Wizards, also welche Events der Demo-User in der Liste sieht.'
                  : 'Pick a person from the tenant via the people picker. Also enter the location (required) — it controls the location filter from step 3 of the event wizard, i.e. which events the demo user sees in their list.'}
              </>
            ),
            tip: isDe
              ? 'Standort und Person müssen nicht zusammenpassen — der Standort ist nur der Filter, der Name + E-Mail bestimmen die Identität.'
              : 'Location and person do not need to match — the location is just the filter, the name + email define the identity.',
          },
          {
            number: 3,
            title: isDe ? 'Was sich ändert' : 'What changes',
            description: (
              <>
                {isDe
                  ? 'Nach „Demo-Modus starten" reloaded die App. Die Rolle wird auf „User" gesetzt (isAdmin=false, canCreateEvents=false), Name + E-Mail + Standort werden aus deinen Demo-Werten gezogen. Du siehst die Event-Liste so wie die Zielperson sie sehen würde — inklusive Filter, Anmelde-Buttons, Mein-Profil-Daten.'
                  : 'After "Start demo mode" the app reloads. The role is set to "User" (isAdmin=false, canCreateEvents=false), name + email + location are pulled from your demo values. You see the event list as the target person would — including filters, register buttons, profile data.'}
              </>
            ),
            tip: isDe
              ? 'Admin-Seiten (Einstellungen, Rollenmatrix) bleiben für dich erreichbar, weil sie auf originalIsAdmin prüfen — du verlierst nichts.'
              : 'Admin pages (Settings, Role Matrix) remain accessible for you because they check originalIsAdmin — you don\'t lose anything.',
          },
          {
            number: 4,
            title: isDe ? 'Demo-Modus beenden' : 'End demo mode',
            description: (
              <>
                {isDe
                  ? 'Oben in der App siehst du jederzeit einen orangen Sticky-Banner: „DEMO-MODUS · agierst als ... · Standort ...". Klick rechts auf „Demo-Modus beenden" — die Impersonation wird gelöscht, die App reloaded, du bist wieder du selbst.'
                  : 'At the top of the app you always see an orange sticky banner: "DEMO MODE · acting as ... · location ...". Click "End demo mode" on the right — the impersonation is cleared, the app reloads, you\'re yourself again.'}
              </>
            ),
            tip: isDe
              ? 'Der Demo-Status wird in localStorage gespeichert (Key: dex_demo_impersonation) und überlebt einen Tab-Reload. Beim Beenden wird der Key gelöscht.'
              : 'Demo status is stored in localStorage (key: dex_demo_impersonation) and survives a tab reload. Exiting clears the key.',
          },
        ],
      },
    ],
  };
}
