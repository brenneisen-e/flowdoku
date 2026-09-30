/**
 * Die Konstanten, die an EINER Stelle stehen muessen.
 *
 * Warum eine eigene Datei dafuer: In DEX steht die Site-Adresse an NEUN
 * Stellen in fuenf Dateien, und nur eine davon ist eine benannte Konstante
 * (`services/EmailTemplates.ts:15`). Die uebrigen acht sind ins HTML von
 * Mailtexten und in Wizard-Komponenten gewandert. Ein Umzug auf eine andere
 * Site waere dort keine Aenderung, sondern eine Suche — und eine vergessene
 * Fundstelle faellt erst auf, wenn jemand auf einen toten Link klickt.
 *
 * Hier importiert alles von hier. Wer eine zehnte Stelle braucht, importiert
 * sie ebenfalls; wer eine Adresse direkt hinschreibt, macht denselben Fehler
 * noch einmal.
 */

/** Die SharePoint-Site, auf der die Plattform lebt. Ohne Schraegstrich am Ende. */
export const SITE_URL = 'https://deudeloitte.sharepoint.com/sites/DOL-c-DE-AIUseCasePlatform';

/** Die Seite, auf der das Webpart eingebunden ist — für Links von außen. */
export const APP_URL = `${SITE_URL}/SitePages/AIUseCases.aspx?env=WebView`;

/**
 * Der Name des Abfrageparameters, der einen einzelnen Use Case adressiert
 * (`…AIUseCases.aspx?env=WebView&uc=7`).
 *
 * Steht hier und nicht in einer der beiden Stellen, die ihn brauchen: der
 * NavigationContext LIEST ihn beim Start, der „Link kopieren"-Knopf der
 * Detailseite SCHREIBT ihn. Ein Tippfehler an nur einer von beiden erzeugt
 * Links, die auf die Startseite führen — ohne Fehlermeldung.
 */
export const DEEPLINK_PARAM = 'uc';

/** Der Link, der direkt auf die Detailseite eines Use Cases führt. */
export function linkZumUseCase(id: number): string {
  // `APP_URL` trägt schon ein `?`, deshalb `&`.
  return `${APP_URL}&${DEEPLINK_PARAM}=${id}`;
}

/**
 * Ansprechperson für Fragen zur Plattform — steht im „Hast du Fragen?"-Dialog
 * und im Dialog „Organizer werden?".
 *
 * Die Adresse stammt aus der Galerie im Repo `kiarbeitsplatz`
 * (`galerie/usecases.js`, `PROFILE.contact`) und ist hier die EINZIGE Stelle.
 * Ein Postfach der Plattform (Gruppe statt Person) wäre die bessere Adresse,
 * sobald es eines gibt: Dann ist es eine Änderung an dieser Zeile.
 */
export const KONTAKT_EMAIL = 'ebrenneisen@deloitte.de';

/** Anzeigename der Plattform. Steht im Kopf, im Titel und in Meldungen. */
export const APP_NAME = 'AI Use Case Platform';

/** Untertitel — die fachliche Einordnung aus dem Konzept. */
export const APP_SUBTITLE_DE = 'Agentic Banking Demo Hub';
export const APP_SUBTITLE_EN = 'Agentic Banking Demo Hub';

/**
 * Die SharePoint-Listen dieser App.
 *
 * Praefix `AIUC_` wie `DEX_` in der Event-Plattform: Wer auf der Site die
 * Listenuebersicht oeffnet, sieht sofort, was zur App gehoert und was jemand
 * daneben angelegt hat.
 */
export const LIST = {
  /** Die Use Cases selbst — eine Zeile je Demo. */
  useCases: 'AIUC_UseCases',
  /** Rollenverwaltung. Aufbau wie DEX_Roles. */
  roles: 'AIUC_Roles',
  /** Nachvollziehbarkeit: wer hat wann was geaendert. */
  log: 'AIUC_Log',
};
