/**
 * v26.48: B2Run-Köln-Vorlage.
 *
 * Zentrale Definition für Events, deren Titel „B2Run Köln" enthält:
 *  1. Vorgeschlagene Abfragefelder im Wizard (Schritt „Felder") — per
 *     „Alle übernehmen" in einem Klick ins Event übernehmbar.
 *  2. Exakte Struktur der OFFIZIELLEN B2Run-Köln-Excel für den Export
 *     (Sheet „B2Run Köln <Jahr>", 16 Spalten, Schreibweisen 1:1 aus der
 *     Original-Datei — inkl. „Straße" mit ß, kleingeschriebener Anrede
 *     „männlich/weiblich" und der Spalte „Nordic Walker").
 *
 * Quelle: Deloitte_Teilnehmer_-innen_b2run-koeln-2026.xlsx (offizielle
 * Meldedatei des Veranstalters). Gruppe + private Adresse + Mobilnummer
 * dürfen leer bleiben.
 */
import { CustomField } from '../services/EventService';

/** Greift, wenn der Event-Titel „B2Run Köln" enthält (auch „B2Run-Koeln" etc.). */
export function isB2RunKoelnTitle(title: string | undefined | null): boolean {
  return /b2run[\s\-_]*k(ö|oe)ln/i.test(title || '');
}

/** Exakte Header der offiziellen Excel — Reihenfolge und Schreibweise 1:1. */
export const B2RUN_KOELN_HEADERS: string[] = [
  'Nr.',
  'Anrede',
  'Vorname',
  'Nachname',
  'E-Mail',
  'Startblock',
  'Zustimmung AGB & Datenschutzhinweise',
  'Anonym',
  'Gruppe',
  'Straße und Hausnummer (privat)',
  'PLZ (privat)',
  'Stadt (privat)',
  'Mobilnummer',
  'Verwendung Infoservice',
  'Altersklasse',
  'Nordic Walker',
];

/** Exakte Startblock-Werte aus der Original-Datei. */
export const B2RUN_KOELN_STARTBLOCK_DURCHSTARTER = '17:00 Uhr Durchstarter (Orange)';
export const B2RUN_KOELN_STARTBLOCK_FUNSTARTER = '17:00 Uhr Funstarter (Grün)';

/** Einheitswert der Original-Datei. */
export const B2RUN_KOELN_ALTERSKLASSE = 'offene Klasse';

/** DEX-Anrede (Herr/Frau/Divers) → B2Run-Schreibweise (klein). */
export function mapAnredeToB2Run(anrede: string | undefined | null): string {
  const a = (anrede || '').trim().toLowerCase();
  if (a === 'herr' || a === 'männlich') return 'männlich';
  if (a === 'frau' || a === 'weiblich') return 'weiblich';
  if (a === 'divers') return 'divers';
  return '';
}

/** StarterType (Split-Capacity) → exakter Startblock-Text der Excel. */
export function mapStarterTypeToStartblock(starterType: string | undefined | null): string {
  const s = (starterType || '').trim().toLowerCase();
  if (s === 'durchstarter') return B2RUN_KOELN_STARTBLOCK_DURCHSTARTER;
  if (s === 'funstarter') return B2RUN_KOELN_STARTBLOCK_FUNSTARTER;
  return '';
}

/**
 * Die vorgeschlagenen Abfragefelder der Vorlage. Deterministische IDs — der
 * Export liest exakt diese IDs aus CustomData. `b2run_startblock`/
 * `b2run_datenschutz`/`b2run_infoservice`/`b2run_mobilnummer` sind bewusst
 * identisch mit den bestehenden B2Run-Katalog-IDs (Render-Spezialfälle in
 * der Anmeldemaske — Datenschutz-Links, Mobilnummer-showIf — greifen damit
 * automatisch auch hier).
 */
export function b2runKoelnTemplateFields(isDe: boolean): CustomField[] {
  return [
    {
      id: 'b2run_geschlecht',
      label: isDe ? 'Geschlecht (laut B2Run-Meldung)' : 'Gender (as per B2Run entry)',
      type: 'select', required: true, visible: true,
      options: ['männlich', 'weiblich', 'divers'],
      helpText: isDe
        ? 'Wird 1:1 in die offizielle B2Run-Meldedatei übernommen (Spalte „Anrede").'
        : 'Copied 1:1 into the official B2Run entry file (column "Anrede").',
    },
    {
      id: 'b2run_startblock',
      label: 'Startblock',
      type: 'select', required: true, visible: true,
      options: [B2RUN_KOELN_STARTBLOCK_DURCHSTARTER, B2RUN_KOELN_STARTBLOCK_FUNSTARTER],
    },
    {
      id: 'b2run_datenschutz',
      label: isDe ? 'Zustimmung AGB & Datenschutzhinweise (b2run.de)' : 'Consent to T&C & privacy notice (b2run.de)',
      type: 'checkbox', required: true, visible: true,
    },
    {
      id: 'b2run_anonym',
      label: isDe ? 'Anonym starten (Name erscheint nicht in Ergebnislisten)' : 'Run anonymously (name not shown in result lists)',
      type: 'checkbox', required: false, visible: true,
    },
    {
      id: 'b2run_infoservice',
      label: isDe ? 'B2Run-Infoservice nutzen (SMS-Infos am Lauftag)' : 'Use the B2Run info service (SMS updates on race day)',
      type: 'checkbox', required: false, visible: true,
    },
    {
      id: 'b2run_mobilnummer',
      label: isDe ? 'Mobilnummer (für den Infoservice)' : 'Mobile number (for the info service)',
      type: 'text', required: false, visible: true,
      showIf: { fieldId: 'b2run_infoservice', values: ['true'] },
    },
    {
      id: 'b2run_nordicwalker',
      label: isDe ? 'Als Nordic Walker teilnehmen' : 'Take part as a Nordic walker',
      type: 'checkbox', required: false, visible: true,
    },
  ];
}
