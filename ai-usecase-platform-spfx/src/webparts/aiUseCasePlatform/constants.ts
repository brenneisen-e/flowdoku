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

/** Die Seite, auf der das Webpart eingebunden ist — fuer Links von aussen. */
export const APP_URL = `${SITE_URL}/SitePages/AIUseCases.aspx?env=WebView`;

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
