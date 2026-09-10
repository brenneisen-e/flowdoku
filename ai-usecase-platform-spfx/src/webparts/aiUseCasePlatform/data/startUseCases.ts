/**
 * Die fuenf Start-Use-Cases aus dem Konzept-Deck „Agentic Banking Demo Hub"
 * (Folie 3: „Top-Use-Cases aus dem Deloitte AI Dossier — Banking").
 *
 * Warum das im Code steht und nicht nur in SharePoint: Eine frisch
 * installierte Plattform ist leer, und eine leere Kachelwand erklaert
 * niemandem, wofuer sie da ist. Ein Kurator legt die fuenf mit einem Klick
 * an und hat sofort etwas zu zeigen — die Teams tragen ihre Links spaeter
 * selbst nach.
 *
 * Status ist ueberall `Geplant`: Die Demos entstehen erst auf dem Hackathon.
 * Ein Eintrag auf „Live" ohne Deployment-Link waere eine Kachel, die nichts
 * tut.
 *
 * Die Bewertungen stehen so im Deck (drei Punkte = hoch, zwei = mittel).
 */

import { UseCase } from '../types';

export const START_USE_CASES: Array<Partial<UseCase>> = [
  {
    titel: 'Portfolio- / Anlage-Advisory-Agent',
    bereich: 'Investment / Private Banking',
    kurzbeschreibung: 'Kundenprofil und Depot analysieren, Anlagevorschlag erstellen, Suitability prüfen und Beratungsprotokoll entwerfen.',
    salesRelevanz: 'Hoch', machbarkeit: 'Mittel', demoTauglichkeit: 'Hoch',
    status: 'Geplant', aufrufArt: 'fenster', reihenfolge: 10,
    schlagworte: ['Depot', 'Suitability', 'Beratungsprotokoll', 'Advisory'],
  },
  {
    titel: 'Kredit- & Credit-Memo-Agent',
    bereich: 'Sales / Corporate Banking',
    kurzbeschreibung: 'Finanzdaten sammeln, Vor-Bonitätsprüfung durchführen und einen Credit-Memo-Entwurf generieren.',
    salesRelevanz: 'Hoch', machbarkeit: 'Hoch', demoTauglichkeit: 'Hoch',
    status: 'Geplant', aufrufArt: 'fenster', reihenfolge: 20,
    schlagworte: ['Bonität', 'Credit Memo', 'Kredit', 'Corporate'],
  },
  {
    titel: 'Regulatory-Change-Agent',
    bereich: 'Compliance & Risk',
    kurzbeschreibung: 'Regulatorische Updates überwachen, auf interne Policies mappen und Umsetzungs-Reports entwerfen.',
    salesRelevanz: 'Hoch', machbarkeit: 'Mittel', demoTauglichkeit: 'Mittel',
    status: 'Geplant', aufrufArt: 'fenster', reihenfolge: 30,
    schlagworte: ['Regulatorik', 'Policy', 'Compliance', 'Reporting'],
  },
  {
    titel: 'Fraud-Investigation-Agent',
    bereich: 'Operations / Risk',
    kurzbeschreibung: 'Auffällige Transaktionen triagieren, Fälle untersuchen und einen Verdachtsmeldungs-Entwurf erstellen.',
    salesRelevanz: 'Mittel', machbarkeit: 'Mittel', demoTauglichkeit: 'Hoch',
    status: 'Geplant', aufrufArt: 'fenster', reihenfolge: 40,
    schlagworte: ['Fraud', 'Transaktionen', 'Verdachtsmeldung', 'AML'],
  },
  {
    titel: 'Retail-Service-Agent',
    bereich: 'Customer Experience',
    kurzbeschreibung: 'Konto- und Transaktionsauskünfte, Produktberatung und nahtloses Handover an Mitarbeitende.',
    salesRelevanz: 'Mittel', machbarkeit: 'Hoch', demoTauglichkeit: 'Hoch',
    status: 'Geplant', aufrufArt: 'fenster', reihenfolge: 50,
    schlagworte: ['Retail', 'Service', 'Handover', 'Kundenservice'],
  },
];
