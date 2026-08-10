/**
 * Globale Typdeklarationen für Nicht-TS-Module.
 *
 * SCSS-Module: SPFx bundelt SCSS-Dateien zur Build-Zeit in JS-Module, die ein
 * Record von Klassen-Namen exportieren. TypeScript kennt diesen Mechanismus
 * nicht out-of-the-box, daher hier eine generische Modul-Deklaration.
 */

declare module '*.module.scss' {
  const classes: { [key: string]: string };
  export default classes;
}
