/**
 * Modul-Ebene aus AdminPage.tsx ausgelagert (v30.66) — die Typen, die sowohl
 * die Seite als auch die ausgelagerten Teilansichten brauchen. Ein Typ im
 * Funktionskoerper waere von aussen nicht referenzierbar, deshalb hier.
 */
import { SPRegistration } from '../../services/EventService';

export type ConsolidatedRow = {
  emailKey: string;
  email: string;
  vorname: string;
  nachname: string;
  jobTitle: string;
  location: string;
  company: string;
  teilnehmerId: number | null;
  /** v15.23: Früheste RegistrationDate über alle Sub-Event-
   *  Registrierungen der Person — Default-Sortierschlüssel im
   *  konsolidierten View (chronologisch nach erster Anmeldung). */
  earliestRegistrationTs: number;
  perChild: Record<string, SPRegistration | undefined>;
  activeCount: number;
};

/** v29.36: Eine Person aus dem Sichtbarkeits-Kreis eines Events. Name, Position
 *  und Standort kommen aus derselben Verteiler-Abfrage wie die Adresse — für
 *  einzeln eingetragene Personen bleiben sie leer. */
export type AudiencePerson = { email: string; displayName?: string; jobTitle?: string; location?: string };

/** v30.66: Zielgruppe der Massenmail. Stand als lokaler Typ IM Komponenten-
 *  koerper und war damit von aussen nicht referenzierbar — das ausgelagerte
 *  Zielgruppen-Modal braucht ihn aber. */
export type MassmailAudience = 'active' | 'activePlusWait' | 'waitOnly' | 'nachruecker' | 'custom';
