import * as React from 'react';
import { ManualSection } from '../types';

/**
 * v26.33: Mobile Nutzung — DEX am Smartphone über die Microsoft-SharePoint-App
 * öffnen. Beantwortet die häufige Frage „Gibt es DEX als App / geht das am
 * Handy?" und den Android-Fall, in dem sich die SharePoint-Seite im mobilen
 * Browser nicht sauber öffnet. Mit Such-Stichwörtern für Suche + Frage-Vorschläge.
 */
export function mobileUsageSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'mobile-usage',
    title: isDe ? 'DEX am Handy nutzen (mobile Nutzung)' : 'Using DEX on mobile',
    category: 'general',
    description: isDe
      ? 'DEX läuft auf SharePoint. Am Smartphone öffnest du DEX am besten über die Microsoft-SharePoint-App — besonders unter Android, wo der mobile Browser die Seite teilweise nicht korrekt lädt.'
      : 'DEX runs on SharePoint. On a smartphone it is best opened via the Microsoft SharePoint app — especially on Android, where the mobile browser sometimes fails to load the page.',
    visibleFor: ['User', 'Organizer', 'Admin'],
    keywords: 'mobil mobile Handy Smartphone Telefon unterwegs App SharePoint-App SharePoint App Microsoft SharePoint Android iOS iPhone iPad AppStore App Store Play Store Playstore Google Play herunterladen installieren öffnet nicht Fehler',
    perspectives: [
      {
        perspective: 'user',
        intro: isDe
          ? 'Ja, DEX lässt sich am Handy nutzen. DEX ist eine SharePoint-Seite — am Desktop öffnest du sie im Browser, am Smartphone am zuverlässigsten über die kostenlose Microsoft-SharePoint-App.'
          : 'Yes, DEX can be used on mobile. DEX is a SharePoint page — on desktop you open it in the browser, on a smartphone it works most reliably via the free Microsoft SharePoint app.',
        steps: [
          {
            number: 1,
            title: isDe ? 'Microsoft-SharePoint-App installieren' : 'Install the Microsoft SharePoint app',
            description: (
              <>
                {isDe
                  ? 'Lade die App „Microsoft SharePoint" herunter — im Apple App Store (iPhone/iPad) bzw. im Google Play Store (Android). Melde dich anschließend mit deinem Deloitte-Konto an.'
                  : 'Download the "Microsoft SharePoint" app — from the Apple App Store (iPhone/iPad) or the Google Play Store (Android). Then sign in with your Deloitte account.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'DEX in der App öffnen' : 'Open DEX in the app',
            description: (
              <>
                {isDe
                  ? 'Öffne in der SharePoint-App die Event-Experience-Platform-Seite und wähle „DEX" aus. Tipp: Rufe DEX am Desktop einmal auf und markiere die Seite in SharePoint als „Folgen"/Favorit — dann findest du sie in der App sofort unter „Folge ich".'
                  : 'In the SharePoint app open the Event Experience Platform site and select "DEX". Tip: open DEX once on desktop and mark the page as "Follow"/favourite in SharePoint — then it appears immediately under "Following" in the app.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Warum die App statt Browser?' : 'Why the app instead of the browser?',
            description: (
              <>
                {isDe
                  ? 'Auf manchen Geräten — vor allem unter Android — öffnet sich die DEX-Seite im mobilen Browser nicht korrekt bzw. führt zu Fehlern (leere Seite, Anmelde-Schleife). Über die Microsoft-SharePoint-App wird die Seite zuverlässig geladen. Nutze am Handy daher immer die App.'
                  : 'On some devices — especially on Android — the DEX page does not open correctly in the mobile browser or throws errors (blank page, login loop). Via the Microsoft SharePoint app the page loads reliably. So on mobile always use the app.'}
              </>
            ),
            warning: isDe
              ? 'Android: Wenn sich DEX im Handy-Browser nicht öffnet, ist das kein Fehler in deinem Konto — bitte über die Microsoft-SharePoint-App öffnen.'
              : 'Android: if DEX does not open in the mobile browser, it is not a problem with your account — please open it via the Microsoft SharePoint app.',
          },
        ],
      },
    ],
  };
}
