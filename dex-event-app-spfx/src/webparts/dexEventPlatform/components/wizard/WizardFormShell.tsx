/* WizardFormShell — aus EventCreationPage.tsx ausgelagert (Zeilen 4208-4790 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; die Komponente
 * gibt ein Fragment zurueck, damit die Geschwister-Reihenfolge im Elternbaum
 * exakt bleibt. */
import * as React from 'react';
import { STEP_HINTS_DE, STEP_HINTS_EN } from '../../data/wizardHints';
import { BasicsStep, BasicsStepProps } from '../wizard/steps/BasicsStep';
import { DetailsStep, DetailsStepProps } from '../wizard/steps/DetailsStep';
import { LocationProgramStep, LocationProgramStepProps } from '../wizard/steps/LocationProgramStep';
import { SubEventsSection, SubEventsSectionProps } from '../wizard/steps/SubEventsSection';
import { CapacityStep, CapacityStepProps } from '../wizard/steps/CapacityStep';
import { TeamStep } from '../wizard/steps/TeamStep';
import { FieldsStep, FieldsStepProps } from '../wizard/steps/FieldsStep';
import { CommunicationStep, CommunicationStepProps } from '../wizard/steps/CommunicationStep';
import { DocumentsStep } from '../wizard/steps/DocumentsStep';
import { FunZoneStep } from '../wizard/steps/FunZoneStep';
import { BillingStep } from '../wizard/steps/BillingStep';
import { Send, Trash2 } from '../Icons';

export interface WizardFormShellProps {
  actionRowRef: React.MutableRefObject<HTMLDivElement>;
  actionRowVisible: boolean;
  activeScopeIdx: number;
  addQuizQuestion: () => void;
  allowAttendeeUpload: boolean;
  askTeamName: boolean;
  attemptSubmitGuarded: () => void;
  attendeeUploadHint: string;
  attendeeUploadLabel: string;
  basicsStepProps: Omit<BasicsStepProps, 'visible'>;
  billingFields: Record<string, string>;
  billingPromptOpen: boolean;
  billingRelevant: boolean;
  billingSendMode: "manual" | "auto";
  canBilling: boolean;
  communicationStepProps: Omit<CommunicationStepProps, 'visible'>;
  currentStep: number;
  detailsStepProps: Omit<DetailsStepProps, 'visible'>;
  documents: { name: string; file?: File; url: string; size: number; }[];
  draftSavedAt: number;
  draggedQuestionId: string;
  error: string;
  fieldsStepProps: Omit<FieldsStepProps, 'visible'>;
  getStepErrorsFor: (step: number) => string[];
  goBack: () => void;
  hintStepIdx: number;
  isDe: boolean;
  isEditMode: boolean;
  isSubmitting: boolean;
  locationProgramStepProps: Omit<LocationProgramStepProps, 'visible'>;
  pendingSections: string[];
  proceedNext: () => void;
  progress: number;
  progressLabel: string;
  quiz: { id: string; question: string; options: string[]; correctIndices: number[]; imageBase64?: string; section?: string; }[];
  removeQuizQuestion: (id: string) => void;
  renderGlobalScopeBar: () => React.ReactElement | null;
  renderStepIntro: (_bulletsDe: string[], _bulletsEn: string[]) => React.ReactElement | null;
  setAllowAttendeeUpload: React.Dispatch<React.SetStateAction<boolean>>;
  setAskTeamName: React.Dispatch<React.SetStateAction<boolean>>;
  setAttendeeUploadHint: React.Dispatch<React.SetStateAction<string>>;
  setAttendeeUploadLabel: React.Dispatch<React.SetStateAction<string>>;
  setBillingFields: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  setBillingPromptOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setBillingRelevant: React.Dispatch<React.SetStateAction<boolean>>;
  setBillingSendMode: React.Dispatch<React.SetStateAction<"manual" | "auto">>;
  setCurrentStep: React.Dispatch<React.SetStateAction<number>>;
  setDocuments: React.Dispatch<React.SetStateAction<{ name: string; file?: File; url: string; size: number; }[]>>;
  setDraggedQuestionId: React.Dispatch<React.SetStateAction<string>>;
  setHintStepIdx: React.Dispatch<React.SetStateAction<number>>;
  setNewSectionError: React.Dispatch<React.SetStateAction<string>>;
  setNewSectionModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setNewSectionName: React.Dispatch<React.SetStateAction<string>>;
  setPendingSections: React.Dispatch<React.SetStateAction<string[]>>;
  setShowConfigCheck: React.Dispatch<React.SetStateAction<boolean>>;
  setShowRegisterPreview: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamJoinRequiresApproval: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamMembersCannotCreate: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamOpenSlotsVisible: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamPartialAllowed: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamRegistrationEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setTeamSize: React.Dispatch<React.SetStateAction<number>>;
  setTeamTermPlural: React.Dispatch<React.SetStateAction<string>>;
  setTeamTermSingular: React.Dispatch<React.SetStateAction<string>>;
  setTriedNext: React.Dispatch<React.SetStateAction<boolean>>;
  steps: ({ label: string; icon: string; dim?: undefined; } | { label: string; icon: string; dim: boolean; })[];
  subEventsSectionProps: Omit<SubEventsSectionProps, 'visible'>;
  t: (key: string) => string;
  teamJoinRequiresApproval: boolean;
  teamMembersCannotCreate: boolean;
  teamOpenSlotsVisible: boolean;
  teamPartialAllowed: boolean;
  teamRegistrationEnabled: boolean;
  teamSize: number;
  teamTermPlural: string;
  teamTermSingular: string;
  title: string;
  updateQuizQuestion: (id: string, updates: Partial<{    question: string;    options: string[];    correctIndices: number[];    imageBase64: string | undefined;    section: string | undefined;}>) => void;
}

export const WizardFormShell: React.FC<WizardFormShellProps> = (p) => {
  const { actionRowRef, actionRowVisible, activeScopeIdx, addQuizQuestion, allowAttendeeUpload, askTeamName, attemptSubmitGuarded, attendeeUploadHint, attendeeUploadLabel, basicsStepProps, billingFields, billingPromptOpen, billingRelevant, billingSendMode, canBilling, capacityStepProps, communicationStepProps, currentStep, detailsStepProps, documents, draftSavedAt, draggedQuestionId, error, fieldsStepProps, getStepErrorsFor, goBack, hintStepIdx, isDe, isEditMode, isSubmitting, locationProgramStepProps, pendingSections, proceedNext, progress, progressLabel, quiz, removeQuizQuestion, renderGlobalScopeBar, renderStepIntro, setAllowAttendeeUpload, setAskTeamName, setAttendeeUploadHint, setAttendeeUploadLabel, setBillingFields, setBillingPromptOpen, setBillingRelevant, setBillingSendMode, setCurrentStep, setDocuments, setDraggedQuestionId, setHintStepIdx, setNewSectionError, setNewSectionModalOpen, setNewSectionName, setPendingSections, setShowConfigCheck, setShowRegisterPreview, setTeamJoinRequiresApproval, setTeamMembersCannotCreate, setTeamOpenSlotsVisible, setTeamPartialAllowed, setTeamRegistrationEnabled, setTeamSize, setTeamTermPlural, setTeamTermSingular, setTriedNext, steps, subEventsSectionProps, t, teamJoinRequiresApproval, teamMembersCannotCreate, teamOpenSlotsVisible, teamPartialAllowed, teamRegistrationEnabled, teamSize, teamTermPlural, teamTermSingular, title, updateQuizQuestion } = p;
  return (
    <>
      <div>
        {/* ===== Step Progress Bar =====
            v14.8: drei Layout-Fixes für das 9-Schritt-Layout:
            (1) Linie endet exakt auf der Mittelachse des ersten/letzten
                Kreises — vorher fix `left/right: 10%`, was zufällig nur
                für n=8 stimmte; jetzt dynamisch über `100 / (steps.length * 2)`.
            (2) Linie etwas dicker (5 statt 3 px) + abgerundet — sonst
                verschwindet sie bei 9 Schritten optisch.
            (3) `alignItems: flex-start` statt `center` — sonst rutschen
                Kreise nach unten, wenn ein Label (z.B. „Kapazität &
                Sichtbarkeit") umbricht.
            Die Linie sitzt bei top=17, height=5 (Mitte bei 19.5 px) —
            das deckt sich exakt mit der Mitte der 40-px-Kreise. */}
        <div style={{ marginBottom: 32 }}>
          {/* v22.22: Hover-Effekt auf den Schritt-Punkten — hebt den Schritt
              leicht an und färbt Kreis-Rand + Label grün, damit die
              Klickbarkeit sofort erkennbar ist. */}
          <style>{`
            .dex-wizard-step { transition: transform 0.15s ease; }
            .dex-wizard-step:hover { transform: translateY(-2px); }
            .dex-wizard-step:hover .dex-step-circle { border-color: var(--dex-green, #86bc25) !important; box-shadow: 0 4px 12px rgba(134,188,37,0.35) !important; }
            .dex-wizard-step:hover .dex-step-label { color: var(--dex-green-dark, #4a7c1f) !important; }
            /* v22.30: Gefüllter grüner Schritt-Header — sitzt bündig als
               Kopf der weißen Karte (negative Margins überbrücken das
               Karten-Padding): oben rund wie die Karte, unten gerade Kante,
               darunter beginnt der Schritt-Inhalt. */
            .dex-step-head-title {
              margin: -32px -32px 0; padding: 16px 24px 4px;
              background: var(--dex-green, #86bc25); color: #fff;
              font-size: 1.3rem; font-weight: 700;
              border-radius: 15px 15px 0 0;
            }
            .dex-step-head-lead {
              margin: 0 -32px 20px; padding: 0 24px 14px;
              background: var(--dex-green, #86bc25); color: rgba(255,255,255,0.95);
              font-size: 0.85rem; line-height: 1.55;
              border-radius: 0;
            }
            @media (max-width: 768px) {
              .dex-step-head-title { margin: -20px -16px 0; padding: 14px 16px 4px; }
              .dex-step-head-lead { margin: 0 -16px 16px; padding: 0 16px 12px; }
            }
            /* v29.7: Zwischen-Trenner INNERHALB eines Schritts — gleicher
               grüner Balken wie der Schritt-Kopf, nur ohne die runden Ecken
               oben (die gehören dem Kartenanfang) und eine Spur kleiner, damit
               der Schritt-Kopf die Überschrift bleibt. Der frühere Trenner war
               eine graue Haarlinie mit Kleinschrift; die trennt zu leise für
               den Themenwechsel „was ist das Event" → „woraus besteht es". */
            .dex-step-sub-head {
              margin: 32px -32px 0; padding: 13px 24px 3px;
              background: var(--dex-green, #86bc25); color: #fff;
              font-size: 1.1rem; font-weight: 700;
            }
            .dex-step-sub-lead {
              margin: 0 -32px 20px; padding: 0 24px 13px;
              background: var(--dex-green, #86bc25); color: rgba(255,255,255,0.95);
              font-size: 0.85rem; line-height: 1.55;
            }
            @media (max-width: 768px) {
              .dex-step-sub-head { margin: 24px -16px 0; padding: 12px 16px 3px; }
              .dex-step-sub-lead { margin: 0 -16px 16px; padding: 0 16px 11px; }
            }
            /* v22.36: Ausgefüllte Eingaben — pastellgrün wie auf der
               Anmeldeseite (Klasse wird per Sweep/Listener getoggelt). */
            .dex-filled:not(:focus) {
              border-color: var(--dex-green, #86bc25) !important;
              background: rgba(134,188,37,0.07) !important;
            }
          `}</style>
          {(() => {
            const sidePct = 100 / (steps.length * 2);
            const spanPct = 100 - 2 * sidePct;
            return (
          <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', position: 'relative' }}>
            {/* Verbindungslinie */}
            <div style={{ position: 'absolute', top: 17, left: `${sidePct}%`, right: `${sidePct}%`, height: 5, background: 'var(--dex-gray-200)', borderRadius: 3, zIndex: 0 }} />
            <div style={{ position: 'absolute', top: 17, left: `${sidePct}%`, height: 5, background: 'var(--dex-green)', borderRadius: 3, zIndex: 1, width: `${(currentStep / Math.max(1, steps.length - 1)) * spanPct}%`, transition: 'width 0.4s ease' }} />
            {steps.map((step, idx) => (
              <div
                key={idx}
                className="dex-wizard-step"
                data-tour={`wizard-step-${idx}`}
                onClick={() => {
                  // v29.21 (Audit B3): Zurück ist immer frei; nach vorn nur,
                  // wenn ALLE übersprungenen Schritte fehlerfrei sind. Vorher
                  // prüfte der Klick nur den aktuellen Schritt — ein Sprung
                  // von Schritt 1 direkt auf 9 umging z.B. die Organizer-
                  // Pflicht aus Schritt 2, und handleSubmitInner prüft nur den
                  // Titel: Ein Event ohne Organizer war anlegbar. Bei einem
                  // Fehler springt der Wizard auf den ersten fehlerhaften
                  // Schritt, damit die Markierungen sichtbar sind.
                  if (idx <= currentStep) { setCurrentStep(idx); return; }
                  setTriedNext(true);
                  for (let st = currentStep; st < idx; st++) {
                    if (getStepErrorsFor(st).length > 0) { setCurrentStep(st); return; }
                  }
                  setTriedNext(false);
                  setCurrentStep(idx);
                }}
                style={{
                  display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8,
                  zIndex: 2, cursor: 'pointer',
                  flex: 1,
                  // v29.66: ausgegrauter Abrechnungs-Schritt (s. steps-Array).
                  // v30.28: NIE den Schritt ausgrauen, auf dem man gerade steht.
                  // Der aktive Kreis ist grün mit weißer Ziffer — bei 50 %
                  // Deckkraft wird daraus Weiß auf Blassgrün, also unlesbar.
                  // Das Ausgrauen soll „optional" sagen, nicht „unlesbar".
                  opacity: ((step as { dim?: boolean }).dim && idx !== currentStep) ? 0.5 : 1,
                }}
              >
                <div className="dex-step-circle" style={{
                  width: 40, height: 40, borderRadius: '50%',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  fontWeight: 700, fontSize: '1rem',
                  background: idx <= currentStep ? 'var(--dex-green)' : '#fff',
                  color: idx <= currentStep ? '#fff' : 'var(--dex-gray-400)',
                  border: idx <= currentStep ? '3px solid var(--dex-green)' : '3px solid var(--dex-gray-200)',
                  transition: 'all 0.3s ease',
                  boxShadow: idx === currentStep ? '0 0 0 4px rgba(134,188,37,0.2)' : 'none',
                }}>
                  {idx < currentStep ? '✓' : step.icon}
                </div>
                <span className="dex-step-label" style={{
                  fontSize: '0.75rem', fontWeight: idx === currentStep ? 700 : 500,
                  color: idx <= currentStep ? 'var(--dex-green)' : 'var(--dex-gray-400)',
                  transition: 'color 0.3s ease',
                  textAlign: 'center',
                }}>
                  {step.label}
                </span>
                {/* v9.27/v9.37: i-Icon UNTER dem Step-Label (vorher inline rechts daneben).
                    Hover zeigt die Hints für diesen Step.
                    v9.37: Styling identisch zur InfoTooltip-Komponente (serif, 20x20,
                    1.5px-Border) — sonst wirkt das wizard-i im Vergleich klobig. */}
                <span
                  role="button"
                  tabIndex={0}
                  aria-label={isDe ? 'Hinweise zu diesem Schritt' : 'Hints for this step'}
                  onMouseEnter={() => setHintStepIdx(idx)}
                  onMouseLeave={() => setHintStepIdx(null)}
                  onFocus={() => setHintStepIdx(idx)}
                  onBlur={() => setHintStepIdx(null)}
                  onClick={e => { e.stopPropagation(); setHintStepIdx(prev => prev === idx ? null : idx); }}
                  style={{
                    position: 'relative',
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    width: 20, height: 20, borderRadius: '50%',
                    background: hintStepIdx === idx ? 'var(--dex-gray-100, #f0f0f0)' : 'transparent',
                    color: 'var(--dex-gray-700, #555)',
                    border: `1.5px solid ${hintStepIdx === idx ? 'var(--dex-gray-700, #555)' : 'var(--dex-gray-500, #888)'}`,
                    fontSize: '0.7rem', fontWeight: 700, fontFamily: 'serif',
                    cursor: 'help',
                    marginTop: 4,
                    userSelect: 'none',
                    transition: 'background 0.15s, border-color 0.15s',
                  }}
                >
                  i
                  {hintStepIdx === idx && (
                    <div
                      role="tooltip"
                      style={{
                        // v9.40: Styling 1:1 wie InfoTooltip (siehe InfoTooltip.tsx),
                        // damit die zwei Tooltip-Varianten optisch konsistent wirken.
                        position: 'absolute',
                        top: 'calc(100% + 8px)',
                        left: '50%',
                        transform: 'translateX(-50%)',
                        width: 'max-content',
                        maxWidth: 480,
                        minWidth: 280,
                        background: 'rgba(40,40,40,0.96)',
                        color: '#fff',
                        padding: '12px 16px',
                        borderRadius: 8,
                        boxShadow: '0 6px 18px rgba(0,0,0,0.28)',
                        // v15: explizit Sans-Serif — vorher 'inherit', was
                        // den serif-Font des parent „i"-Icons (s. unten
                        // fontFamily:'serif' für das i-Glyph) übernommen
                        // hat und den ganzen Tooltip Times-artig erscheinen
                        // ließ. Jetzt 1:1 wie InfoTooltip.
                        fontFamily: 'Aptos, "Open Sans", "Segoe UI", Arial, Helvetica, sans-serif',
                        fontSize: '0.82rem',
                        lineHeight: 1.55,
                        fontWeight: 400,
                        fontStyle: 'normal',
                        textAlign: 'left',
                        whiteSpace: 'normal',
                        zIndex: 1500,
                        pointerEvents: 'none',
                      }}
                    >
                      <div style={{ fontWeight: 700, marginBottom: 8, color: 'rgba(255,255,255,0.92)' }}>
                        {isDe ? 'Was ich hier einstellen kann' : 'What I can configure here'}
                      </div>
                      <ul style={{ margin: 0, paddingLeft: 18 }}>
                        {(isDe ? STEP_HINTS_DE : STEP_HINTS_EN)[idx]?.map((b, bi) => (
                          <li key={bi} style={{ marginBottom: 4 }}>{b}</li>
                        ))}
                      </ul>
                    </div>
                  )}
                </span>
              </div>
            ))}
          </div>
            );
          })()}
        </div>

        {/* v28.78: Scope-Karte zwischen Schritt-Leiste und Formular — eine
            Ebene für „für wen gilt das hier?", die durch alle Schritte trägt. */}
        {renderGlobalScopeBar()}

        {/* v30.1: Autosave-Anzeige der Neu-Anlage — Speicher-Symbol plus
            Zeitstempel der letzten Zwischenspeicherung, auf jedem Schritt
            sichtbar. Nur Neu-Anlage: im Edit-Modus gibt es keinen
            Entwurfs-Zwischenspeicher (s. v30.0-Block). */}
        {!isEditMode && draftSavedAt !== null && (
          <div style={{ display: 'flex', justifyContent: 'flex-end', margin: '6px 2px 2px' }}>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" style={{ flexShrink: 0 }}>
                <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
                <polyline points="17 21 17 13 7 13 7 21" />
                <polyline points="7 3 7 8 15 8" />
              </svg>
              {isDe
                ? `Zwischengespeichert am ${new Date(draftSavedAt).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })} um ${new Date(draftSavedAt).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })} Uhr`
                : `Auto-saved on ${new Date(draftSavedAt).toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })} at ${new Date(draftSavedAt).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })}`}
            </span>
          </div>
        )}

        {/* ===== Formular ===== */}
        <div>
          <div className="card" style={{ borderRadius: 16, boxShadow: '0 2px 12px rgba(0,0,0,0.06)' }}>
            <div className="creation-form">
              {/* v22.30: marginBottom 48 kompensiert die -32px-Top-Margin
                  des grünen Schritt-Headers darunter (netto 16px Abstand). */}
              {error && (
                <div style={{ padding: '10px 16px', background: '#fce4ec', color: '#c62828', borderRadius: 8, marginBottom: 48, fontSize: '0.85rem' }}>
                  {error}
                </div>
              )}

              {/* ===== Schritt 1: Grundlagen =====
                  v9.32: 1-basierte UI-Nummerierung (in der Logik bleibt
                  currentStep 0-basiert) — siehe ENTWICKLUNG.md. */}
              <BasicsStep visible={currentStep === 0} {...basicsStepProps} />

              {/* v24.12: Organizer-Einstellungen als eigener Wizard-Schritt (Schritt 2). */}
              <DetailsStep visible={currentStep === 1} {...detailsStepProps} />

              {/* ===== Step 3 (v15.0: vormals Step 2): Ort & Programm ===== */}
              <LocationProgramStep visible={currentStep === 2} {...locationProgramStepProps} />

              {/* ===== Step 2 (v15.0, vormals Step 3): Sub-Events =====
                  Sub-Events (Workshops / Sessions / Programmpunkte), plus
                  Bezeichnungs-Dropdown und Anmelde-Modus (Hauptevent +
                  Sub-Events vs. nur Sub-Events).
                  v15.0: vorgezogen vor „Ort & Programm", damit die folgenden
                  Steps pro-Sub-Event-Tabs anbieten können. */}
              {/* v28.87: Frueher Schritt 3 („Sub-Events"). Der Schritt ist
                  entfallen; sein Inhalt hängt jetzt unten an Schritt 1
                  (Grundlagen). Der Block bleibt als Ganzes bestehen — nur
                  seine Anzeige-Bedingung zeigt auf Schritt 1.
                  v28.89: …und nur auf der Klammer-/Hauptevent-Ebene. Die Liste
                  ist die Übersicht ÜBER die Sub-Events; auf dem Reiter eines
                  einzelnen Sub-Events stünde sie unter dessen eigenen
                  Grundlagen und läse sich wie eine Verschachtelung. */}
              <SubEventsSection visible={currentStep === 0 && activeScopeIdx === 0} {...subEventsSectionProps} />

              {/* ===== Step 4 (v14.8: vormals Step 3): Kapazität, Fristen & Sichtbarkeit ===== */}
              <CapacityStep visible={currentStep === 3} {...capacityStepProps} />

              {/* ===== Step 5 (v14.8: vormals Step 4): Team-Anmeldung =====
                  Renderblock für den Wizard-Schritt Team-Anmeldung.
                  Konfiguriert Team-Anmeldung-Toggle + Teamgröße +
                  Team-Name-Frage. v15: Index 4 → 6 (Team kommt jetzt nach
                  Kommunikation). */}
              <TeamStep
                visible={currentStep === 6}
                teamRegistrationEnabled={teamRegistrationEnabled}
                setTeamRegistrationEnabled={setTeamRegistrationEnabled}
                teamSize={teamSize}
                setTeamSize={setTeamSize}
                askTeamName={askTeamName}
                setAskTeamName={setAskTeamName}
                teamTermSingular={teamTermSingular}
                setTeamTermSingular={setTeamTermSingular}
                teamTermPlural={teamTermPlural}
                setTeamTermPlural={setTeamTermPlural}
                teamMembersCannotCreate={teamMembersCannotCreate}
                setTeamMembersCannotCreate={setTeamMembersCannotCreate}
                teamPartialAllowed={teamPartialAllowed}
                setTeamPartialAllowed={setTeamPartialAllowed}
                teamOpenSlotsVisible={teamOpenSlotsVisible}
                setTeamOpenSlotsVisible={setTeamOpenSlotsVisible}
                teamJoinRequiresApproval={teamJoinRequiresApproval}
                setTeamJoinRequiresApproval={setTeamJoinRequiresApproval}
              />

              {/* ===== Step 5 (v15: vormals Step 6): Registrierungsfelder ===== */}
              <FieldsStep visible={currentStep === 4} {...fieldsStepProps} />

              {/* ===== Step 6 (v15: vormals Step 7): Kommunikation ===== */}
              <CommunicationStep visible={currentStep === 5} {...communicationStepProps} />

              {/* ===== Step 8 (v14.8: vormals Step 7): Dokumente ===== */}
              <DocumentsStep
                visible={currentStep === 7}
                documents={documents}
                setDocuments={setDocuments}
                allowAttendeeUpload={allowAttendeeUpload}
                setAllowAttendeeUpload={setAllowAttendeeUpload}
                attendeeUploadLabel={attendeeUploadLabel}
                setAttendeeUploadLabel={setAttendeeUploadLabel}
                attendeeUploadHint={attendeeUploadHint}
                setAttendeeUploadHint={setAttendeeUploadHint}
                renderStepIntro={renderStepIntro}
              />

              {/* ===== Step 9 (v14.8: vormals Step 8): Fun-Zone ===== */}
              <FunZoneStep
                visible={currentStep === 8}
                quiz={quiz}
                addQuizQuestion={addQuizQuestion}
                removeQuizQuestion={removeQuizQuestion}
                updateQuizQuestion={updateQuizQuestion}
                pendingSections={pendingSections}
                setPendingSections={setPendingSections}
                draggedQuestionId={draggedQuestionId}
                setDraggedQuestionId={setDraggedQuestionId}
                setNewSectionName={setNewSectionName}
                setNewSectionError={setNewSectionError}
                setNewSectionModalOpen={setNewSectionModalOpen}
                renderStepIntro={renderStepIntro}
              />

              {/* v29.66: F&A-Pilot — Schritt 10 „Abrechnung" (nur Admins). Haengt
                  als LETZTER Schritt an, damit kein bestehender Index wandert.
                  Nicht abrechnungsrelevant: nur die Frage; „Ja" blendet die
                  Abschnitte sofort ein (reiner State, kein Neuladen).
                  v30.38: Stand bis hierher AUSSERHALB von `.card`/`.creation-form`
                  — der grüne Schritt-Kopf lebt von `margin: -32px` gegen die
                  Karten-Polsterung, und ohne Karte bleute er gegen die Seite.
                  Sichtbar war das als leerer weißer Streifen über dem Kopf: die
                  Karte mit den Schritten 1–9 (alle `display:none`) rendert dann
                  nur noch ihre eigene Polsterung. */}
              {canBilling && (
                <BillingStep
                  visible={currentStep === 9}
                  billingRelevant={billingRelevant}
                  setBillingRelevant={setBillingRelevant}
                  billingSendMode={billingSendMode}
                  setBillingSendMode={setBillingSendMode}
                  billingFields={billingFields}
                  setBillingFields={setBillingFields}
                />
              )}

            </div>{/* close creation-form */}
          </div>{/* close card */}

          {/* v29.66: F&A-Pilot — Frage-Dialog nach den Nutzungsbedingungen.
              Keine Vorauswahl: Das Konzept verlangt eine AKTIVE Entscheidung,
              deshalb zwei gleichrangige Knoepfe statt Radio mit Default. */}
          {canBilling && billingPromptOpen && (
            <div style={{
              position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1300,
              display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
            }}>
              <div className="card" style={{
                width: '100%', maxWidth: 560, padding: 28, borderRadius: 16,
                background: '#fff', boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
              }}>
                {/* v29.69: Pilot-Badge — der Dialog erscheint nur Admins;
                    das soll man ihm ansehen, damit im Test niemand glaubt,
                    Organizer bekaemen diese Frage bereits. */}
                <span style={{
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                  background: 'rgba(237,139,0,0.14)', color: 'var(--dex-orange-dark, #b96a00)',
                  border: '1px solid var(--dex-orange, #ed8b00)',
                  borderRadius: 999, padding: '3px 12px', marginBottom: 10,
                  fontSize: '0.72rem', fontWeight: 700, letterSpacing: 0.4, textTransform: 'uppercase',
                }}>
                  Pilot — aktuell nur für Admins sichtbar
                </span>
                <h2 style={{ margin: '0 0 10px', fontSize: '1.15rem' }}>
                  Handelt es sich um ein abrechnungsrelevantes Event?
                </h2>
                <p style={{ fontSize: '0.9rem', color: 'var(--dex-gray-600)', margin: '0 0 18px' }}>
                  Abrechnungsrelevante Events sind Veranstaltungen, deren Kosten oder
                  Bewirtungsaufwendungen gegenüber Finance &amp; Accounting dokumentiert
                  oder abgerechnet werden müssen. Das ist der Fall, wenn im Nachgang
                  <strong> Rechnungen über die Kreditorenbuchhaltung eingereicht
                  werden</strong> — etwa für Catering, eine externe Raumbuchung oder
                  Anmeldegebühren (z.B. Startgelder für Läufer) — oder wenn für das
                  Event <strong>Ariba-Bestellungen</strong> ausgelöst werden.
                </p>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', margin: '0 0 18px' }}>
                  Die Entscheidung lässt sich später jederzeit im Schritt „Abrechnung&ldquo; ändern.
                </p>
                <div style={{ display: 'flex', gap: 12, justifyContent: 'flex-end' }}>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    onClick={() => { setBillingRelevant(false); setBillingPromptOpen(false); }}
                  >
                    Nein
                  </button>
                  <button
                    type="button"
                    className="btn btn-primary"
                    onClick={() => { setBillingRelevant(true); setBillingPromptOpen(false); }}
                  >
                    Ja, abrechnungsrelevant
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Fortschrittsanzeige */}
          {isSubmitting && (
            <div className="mt-24" style={{ padding: '20px 0' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
                  {progressLabel}
                </span>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--dex-green)' }}>
                  {progress}%
                </span>
              </div>
              <div style={{
                width: '100%', height: 8, background: 'var(--dex-gray-200)',
                borderRadius: 4, overflow: 'hidden',
              }}>
                {/* v29.57: Derselbe Schimmer wie im Boot-Balken (v29.41). Beim
                    Speichern eines Events mit vielen Terminen steht der Balken
                    zwischen zwei Abschnitten sekundenlang fast still — ohne
                    Lebenszeichen liest sich das als Hänger, und Organizer
                    klicken dann ein zweites Mal auf Speichern. Bewusst schwach
                    (weiß auf Grün, 45 %) und langsam, kein Blinken. */}
                <div style={{
                  width: `${progress}%`, height: '100%',
                  background: progress === 100
                    ? 'var(--dex-green)'
                    : 'linear-gradient(90deg, var(--dex-green), #0076a8)',
                  borderRadius: 4,
                  transition: 'width 0.5s ease',
                  position: 'relative',
                  overflow: 'hidden',
                }}>
                  {progress < 100 && (
                    <div style={{
                      position: 'absolute', top: 0, bottom: 0, left: 0, width: '35%',
                      background: 'linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.45) 50%, rgba(255,255,255,0) 100%)',
                      animation: 'dexProgressShimmer 2.1s ease-in-out infinite',
                    }} />
                  )}
                </div>
              </div>
            </div>
          )}

          {!isSubmitting && (
            <>
            {/* v22.36: Aufgeräumte Aktions-Leiste — klare Hierarchie statt
                drei loser Gruppen: links die Rück-Navigation, rechts erst die
                ruhigen Werkzeuge (Vorschau / Prüfen), dann durch einen
                Trenner abgesetzt die Haupt-Aktionen (Speichern als Outline,
                Weiter/Anlegen als EINZIGER gefüllter Primär-Button).
                v22.22: ref für den schwebenden Weiter-Button. */}
            <div ref={actionRowRef} style={{
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              gap: 12, marginTop: 28, paddingTop: 16,
              borderTop: '1px solid var(--dex-gray-200)', flexWrap: 'wrap',
            }}>
              <div style={{ display: 'flex', gap: 8 }}>
                {currentStep === 0 ? (
                  <button className="btn btn-danger" onClick={() => goBack()}><Trash2 size={16} /> {t('create.cancel')}</button>
                ) : (
                  <button className="btn btn-secondary" onClick={() => setCurrentStep(currentStep - 1)}>
                    {t('general.back')}
                  </button>
                )}
              </div>

              <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setShowRegisterPreview(true)}
                  disabled={!title}
                  style={{ opacity: title ? 1 : 0.5 }}
                  title={title ? 'So sehen Teilnehmer die Registrierungsseite' : 'Event-Titel eingeben, um die Vorschau zu öffnen'}
                >
                  {t('create.registerpreview')}
                </button>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setShowConfigCheck(true)}
                  disabled={!title}
                  style={{ opacity: title ? 1 : 0.5 }}
                  title={isDe ? 'Alle Einstellungen des Events im Überblick prüfen' : 'Review all event settings at a glance'}
                >
                  {isDe ? 'Prüfen' : 'Review'}
                </button>

                {/* Trenner zwischen Werkzeugen und Haupt-Aktionen */}
                <span aria-hidden="true" style={{ width: 1, height: 26, background: 'var(--dex-gray-200)', margin: '0 4px' }} />

                {/* v17.5: Im Edit-Modus immer speichern können, ohne durch
                    alle Steps zu klicken. v22.36: als Outline-Variante —
                    der gefüllte Primär-Button bleibt Weiter/Anlegen. */}
                {isEditMode && currentStep < steps.length - 1 && (
                  <button
                    className="btn btn-outline"
                    disabled={!title}
                    onClick={attemptSubmitGuarded}
                    style={{ opacity: !title ? 0.5 : 1 }}
                  >
                    <Send size={16} /> {isDe ? 'Speichern & zurück zum Event' : 'Save & return to event'}
                  </button>
                )}

                {currentStep < steps.length - 1 ? (
                  <button
                    className="btn btn-primary"
                    onClick={proceedNext}
                  >
                    {t('create.next')}
                  </button>
                ) : (
                  <button
                    className="btn btn-primary"
                    data-tour="wizard-submit"
                    disabled={!title}
                    onClick={attemptSubmitGuarded}
                    style={{ opacity: !title ? 0.5 : 1 }}
                  >
                    <Send size={16} /> {isEditMode ? t('create.save') : t('create.submit')}
                  </button>
                )}
              </div>
            </div>

            {/* v22.22: Schwebender Weiter-Button — sichtbar nur, solange die
                Aktions-Zeile unten noch nicht im Viewport ist. Blendet beim
                Erreichen des Seitenendes weich aus (der echte Button übernimmt).
                v22.26: Position + Look exakt wie die Jump-Buttons im Organizer
                Center (unten MITTIG, grüne Pille) — bewährtes fixed-Muster,
                kollidiert nicht mit dem Chat-Icon unten rechts. */}
            {currentStep < steps.length - 1 && (
              <button
                type="button"
                aria-hidden={actionRowVisible}
                tabIndex={actionRowVisible ? -1 : 0}
                onClick={proceedNext}
                style={{
                  position: 'fixed', left: '50%', bottom: 20, zIndex: 900,
                  background: 'var(--dex-green, #86bc25)', color: '#fff',
                  border: 'none', padding: '10px 16px', borderRadius: 999,
                  cursor: 'pointer', fontSize: '0.85rem', fontWeight: 600,
                  boxShadow: '0 4px 12px rgba(0,0,0,0.18)',
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                  fontFamily: 'inherit',
                  transition: 'opacity 0.3s ease, transform 0.3s ease',
                  opacity: actionRowVisible ? 0 : 1,
                  transform: actionRowVisible ? 'translate(-50%, 14px)' : 'translate(-50%, 0)',
                  pointerEvents: actionRowVisible ? 'none' : 'auto',
                }}
              >
                {t('create.next')}
              </button>
            )}
            </>
          )}
        </div>
      </div>
    </>
  );
};
