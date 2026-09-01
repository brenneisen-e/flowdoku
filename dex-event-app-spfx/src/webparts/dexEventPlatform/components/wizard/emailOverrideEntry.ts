// v30.66: Aus EventCreationPage.tsx ausgelagert. Der Typ steckt in den
// Props-Vertraegen der ausgelagerten Wizard-Schritte; ein Import aus der
// Seite selbst waere ein Modul-Zyklus.
// v18.22: Pro-Event-Override eines Mail-Templates (Subject/Headings/Body +
// Formatierung). Vorher als gleiche Inline-Form an ~7 Stellen wiederholt —
// jetzt ein zentraler Alias, damit neue Felder nur hier ergänzt werden.
export type EmailOverrideEntry = {
  subject: string;
  heading: string;
  subheading?: string;
  bodyHtml: string;
  headingColor?: string;
  headingFontSize?: string;
  /** v18.22: Fett/Kursiv für die Überschrift (h1). */
  headingBold?: boolean;
  headingItalic?: boolean;
  /** v18.22: Unter-Überschrift (h2) frei formatierbar. */
  subheadingColor?: string;
  subheadingFontSize?: string;
  subheadingBold?: boolean;
  subheadingItalic?: boolean;
};
