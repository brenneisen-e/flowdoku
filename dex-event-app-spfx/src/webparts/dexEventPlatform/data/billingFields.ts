/**
 * v30.5: Abrechnungsrelevante Pflichtfelder (F&A-Fachkonzept, Abschnitt 4.2/5).
 *
 * Aus EventCreationPage herausgezogen, weil jetzt DREI Stellen dieselbe
 * Definition brauchen: der Wizard-Schritt 10, das F&A Center und die
 * F&A-Mails (Abrechnungsinfos an F&A). Labels bewusst nur deutsch — das
 * Konzept ist deutsch, der Pilot ist intern.
 */

export interface BillingFieldDef {
  id: string;
  label: string;
  /**
   * v30.45: `user` = Person-Picker statt Freitext (F&A-Fachkonzept,
   * „Die Kontaktperson wird nicht als Freitext erfasst"). Gespeichert wird
   * weiter ein String im Format `Name <email>` — dasselbe Format, das
   * `UserFieldPicker` überall sonst in der App schreibt und wieder einliest.
   * Damit bleibt `missingBillingFields`, der F&A-Mailtext und alles, was
   * `_billing` liest, unverändert: Es ist und bleibt ein Textwert.
   */
  type?: 'date' | 'select' | 'user';
  options?: string[];
}

export const BILLING_FIELDS: BillingFieldDef[] = [
  { id: 'contact', label: 'Kontaktperson für Rückfragen', type: 'user' },
  { id: 'docNo', label: 'Dokumenten-Nr. (SH Swift Launchpad)' },
  { id: 'vendor', label: 'Lieferantenname' },
  { id: 'mice', label: 'MICE Project Nummer' },
  { id: 'ariba', label: 'Ariba Bestellnummer' },
  { id: 'company', label: 'Gesellschaft, die die Rechnung erhalten hat' },
  { id: 'category', label: 'Kategorie', type: 'select', options: ['Arbeitsessen', 'Belohnungsessen', 'Sonstiges', 'Geschenk'] },
  { id: 'date', label: 'Veranstaltungs- bzw. Bewirtungsdatum', type: 'date' },
  { id: 'place', label: 'Ort der Bewirtung bzw. Veranstaltung' },
  { id: 'wbs', label: 'WBS-Code / Kostenstelle' },
  { id: 'name', label: 'Name der Veranstaltung bzw. Anlass der Bewirtung oder des Geschenks' },
];
