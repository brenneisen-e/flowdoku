/**
 * wizardStepContext (v26.30)
 *
 * Leichter Marker für den aktuell geöffneten Event-Wizard-Schritt.
 *
 * `EventCreationPage` meldet hier seinen aktiven Schritt (1-basiert, passend zur
 * DexTicket-Konvention `answerWizardStep`/`askWizardStep`). Der grüne
 * „Hast du Fragen?"-Header-Button (`QuestionButton`) liest den Wert beim Öffnen
 * des Modals, um die Frage eines Organizers direkt dem Wizard-Schritt zuzuordnen.
 *
 * Bewusst als Modul-Singleton (kein React-State): der Wert wird ausschließlich
 * imperativ beim Öffnen/Absenden gelesen — es ist kein Re-Render nötig, und der
 * Header-Button und die Wizard-Seite teilen sich keinen nahen React-Kontext.
 */

let activeStep: number | null = null;

/** 1-basierter Wizard-Schritt setzen (null = kein Wizard offen). */
export function setActiveWizardStep(step: number | null): void {
  activeStep = (typeof step === 'number' && step >= 1) ? step : null;
}

/** Aktuellen 1-basierten Wizard-Schritt lesen (null = kein Wizard offen). */
export function getActiveWizardStep(): number | null {
  return activeStep;
}
