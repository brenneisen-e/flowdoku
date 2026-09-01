/* FieldsStep — aus EventCreationPage.tsx ausgelagert (Zeilen 14861-16968 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 4` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher. */
import * as React from 'react';
import WizardHint from '../../WizardHint';
import { b2runKoelnTemplateFields, isB2RunKoelnTitle } from '../../../data/b2runKoeln';
import { CustomField } from '../../../services/EventService';
import { Check, Plus, Trash2, X } from '../../Icons';
import { InfoTooltip } from '../../InfoTooltip';
import { Icon } from '@fluentui/react/lib/Icon';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { CustomFieldInput } from '../../wizard/customFieldInput';
import { FieldTypeSuggestion } from '../../wizard/FieldTypeSuggestion';
import { StepBadge } from '../../wizard/StepBadge';
import { FieldDescEditor } from '../../wizard/FieldDescEditor';
import { de } from 'date-fns/locale';
export interface FieldsStepProps {
  visible: boolean;
  activeFieldsTabIdx: number;
  addCustomField: () => void;
  addStartblock: () => void;
  addSubEventCustomField: (subEventId: string) => void;
  askSalutation: boolean;
  b2runStartblocks: string[];
  bilingualFields: boolean;
  childTermPlural: string;
  confirmDialogEnabled: boolean;
  confirmDialogMode: string;
  confirmDialogText: string;
  copyParentFieldsToSubEvent: (subEventId: string) => void;
  currentStep: number;
  customFields: CustomFieldInput[];
  dragFieldId: string;
  dragOverFieldId: string;
  fieldExpandOverride: Record<string, boolean>;
  isDe: boolean;
  moveCustomField: (id: string, direction: 'up' | 'down') => void;
  newStartblock: string;
  openSuggestedModal: () => void;
  registrationLanguage: "" | "de" | "en";
  removeCustomField: (id: string) => void;
  removeStartblock: (block: string) => void;
  removeSubEventCustomField: (subEventId: string, fieldId: string) => void;
  renderShowIfConfig: (field: CustomFieldInput, idx: number, allFields: CustomFieldInput[], onUpdate: (u: Partial<CustomFieldInput>) => void) => React.ReactElement;
  renderStepIntro: (_bulletsDe: string[], _bulletsEn: string[]) => React.ReactElement | null;
  reorderMode: boolean;
  setAskSalutation: React.Dispatch<React.SetStateAction<boolean>>;
  setBilingualFields: React.Dispatch<React.SetStateAction<boolean>>;
  setConfirmDialogEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setConfirmDialogMode: React.Dispatch<React.SetStateAction<string>>;
  setConfirmDialogText: React.Dispatch<React.SetStateAction<string>>;
  setCustomFields: React.Dispatch<React.SetStateAction<CustomFieldInput[]>>;
  setDragFieldId: React.Dispatch<React.SetStateAction<string>>;
  setDragOverFieldId: React.Dispatch<React.SetStateAction<string>>;
  setNewStartblock: React.Dispatch<React.SetStateAction<string>>;
  setRegistrationLanguage: React.Dispatch<React.SetStateAction<"" | "de" | "en">>;
  setReorderMode: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  splitLabelA: string;
  splitLabelB: string;
  subEvents: SubEventDraft[];
  subEventsOnlyMode: boolean;
  t: (key: string) => string;
  title: string;
  toggleFieldExpand: (id: string, current: boolean) => void;
  updateCustomField: (id: string, updates: Partial<CustomFieldInput>) => void;
  updateSubEventCustomField: (subEventId: string, fieldId: string, updates: Partial<CustomFieldInput>) => void;
  useSplitCapacities: boolean;
}
export const FieldsStep: React.FC<FieldsStepProps> = (p) => {
  const { visible } = p;
  const { activeFieldsTabIdx, addCustomField, addStartblock, addSubEventCustomField, askSalutation, b2runStartblocks, bilingualFields, childTermPlural, confirmDialogEnabled, confirmDialogMode, confirmDialogText, copyParentFieldsToSubEvent, currentStep, customFields, dragFieldId, dragOverFieldId, fieldExpandOverride, isDe, moveCustomField, newStartblock, openSuggestedModal, registrationLanguage, removeCustomField, removeStartblock, removeSubEventCustomField, renderShowIfConfig, renderStepIntro, reorderMode, setAskSalutation, setBilingualFields, setConfirmDialogEnabled, setConfirmDialogMode, setConfirmDialogText, setCustomFields, setDragFieldId, setDragOverFieldId, setNewStartblock, setRegistrationLanguage, setReorderMode, setSubEvents, splitLabelA, splitLabelB, subEvents, subEventsOnlyMode, t, title, toggleFieldExpand, updateCustomField, updateSubEventCustomField, useSplitCapacities } = p;
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              <h2 className="dex-step-head-title">
                {isDe ? 'Schritt 5 — Felder' : 'Step 5 — Fields'}
              </h2>
              <p className="dex-step-head-lead">
                {isDe
                  ? <><strong>Optional</strong> — hier ergänzt du eigene Abfrage-Felder für das Anmeldeformular. Braucht dein Event keine Zusatzfragen, lässt du den Schritt einfach leer.</>
                  : <><strong>Optional</strong> — here you add custom questions to the registration form. If your event needs no extra questions, simply leave this step empty.</>}
              </p>
              {/* v22.30: Detail-Erklärung als graue, eingeklappte Beschreibungs-
                  Box (einheitliche Farb-Logik: grau = Beschreibung). */}
              <WizardHint
                isDe={isDe}
                variant="description"
                title={isDe ? 'Welche Daten werden automatisch erfasst?' : 'Which data is captured automatically?'}
                style={{ marginBottom: 16 }}
              >
                {isDe
                  ? <>
                      <span style={{ display: 'block', marginBottom: 6 }}>Diese Daten werden bei jeder Anmeldung <strong>automatisch erfasst</strong> — du musst sie nicht abfragen:</span>
                      <ul style={{ margin: '0 0 8px', paddingLeft: 18 }}>
                        <li><strong>Vorname</strong> (aus dem Deloitte-Profil)</li>
                        <li><strong>Nachname</strong> (aus dem Deloitte-Profil)</li>
                        <li><strong>E-Mail</strong> (aus dem Deloitte-Profil)</li>
                        <li><strong>Job Title</strong> (aus dem Deloitte-Profil)</li>
                        <li><strong>Standort</strong> (aus dem Deloitte-Profil)</li>
                        <li><strong>Department</strong> (aus dem Deloitte-Profil)</li>
                      </ul>
                      <span style={{ display: 'block' }}>Hier ergänzt du <strong>nur zusätzliche Fragen</strong>, die du speziell für dieses Event brauchst — vom T-Shirt-Größen-Dropdown bis zur Pflicht-Checkbox für AGB / Datenschutz. Optional kannst du unten das <strong>Anrede</strong>-Dropdown einblenden.</span>
                    </>
                  : <>
                      <span style={{ display: 'block', marginBottom: 6 }}>This data is captured <strong>automatically</strong> for every registration — no need to ask for it:</span>
                      <ul style={{ margin: '0 0 8px', paddingLeft: 18 }}>
                        <li><strong>First name</strong> (from the Deloitte profile)</li>
                        <li><strong>Last name</strong> (from the Deloitte profile)</li>
                        <li><strong>Email</strong> (from the Deloitte profile)</li>
                        <li><strong>Job title</strong> (from the Deloitte profile)</li>
                        <li><strong>Location</strong> (from the Deloitte profile)</li>
                        <li><strong>Department</strong> (from the Deloitte profile)</li>
                      </ul>
                      <span style={{ display: 'block' }}>Here you only add <strong>extra questions</strong> specific to this event — from a T-shirt size dropdown to a privacy / terms required checkbox. Optionally enable the <strong>salutation</strong> dropdown below.</span>
                    </>}
              </WizardHint>

              {/* v15: Anrede-Toggle ist nach UNTEN den Datenschutz-Hinweis
                  gewandert — Organizer soll erst den Sammle-keine-sensiblen-
                  Daten-Hinweis lesen, dann erst die Optional-Anrede-Checkbox
                  setzen. Siehe weiter unten. */}

              {renderStepIntro(
                [
                  'Feldtyp wählen: Text, Zahl, Dropdown, Checkbox, Personen-Suche oder Roommate (Doppelzimmer)',
                  'Mehrfachauswahl bei Dropdowns (z.B. mehrere Allergien anhaken)',
                  'Pflichtfeld setzen (rotes Sternchen, Anmeldung blockiert wenn leer)',
                  'Beschreibung pro Feld — landet als „i"-Tooltip neben dem Feld-Label',
                  'Sichtbarkeitsbedingung: Feld nur dann anzeigen wenn eine andere Frage einen bestimmten Wert hat (z.B. „Zimmerart nur fragen wenn Hotel = ja")',
                  'Reihenfolge per Drag oder Pfeilen — die Nummerierung passt sich automatisch an',
                ],
                [
                  'Pick a field type: text, number, dropdown, checkbox, people search or roommate (double room)',
                  'Multi-select for dropdowns (e.g. tick multiple allergies)',
                  'Mark required (red asterisk, blocks submit when empty)',
                  'Description per field — appears as „i" tooltip next to the field label',
                  'Visibility condition: only show this field when another question has a specific value (e.g. „Only ask room type if Hotel = yes")',
                  'Reordering via drag or arrows — numbering updates automatically',
                ]
              )}

              {/* Datenschutz-Hinweis über der Template-Auswahl — links
                  angeordnet, orangener Akzent damit der Organizer beim Anlegen
                  neuer Felder bewusst entscheidet, was wirklich abgefragt
                  werden muss. Seit v7.35 deckungsgleich mit dem Hinweis aus
                  den Nutzungsbedingungen (Sammeln keiner sensiblen Daten),
                  damit Organizer den selben Wortlaut wie bei der initialen
                  Bestätigung sehen. */}
              <WizardHint
                isDe={isDe}
                title={isDe ? 'Sammle keine sensiblen personenbezogenen Daten' : 'Do not collect sensitive personal data'}
                style={{ marginBottom: 16 }}
              >
                {isDe
                  ? <>Das heißt: keine Daten bezüglich Rasse oder ethnischer Herkunft, religiöser oder philosophischer Überzeugungen, Gewerkschaftsmitgliedschaft, politischer Meinungen, medizinischer oder gesundheitlicher Zustände oder Informationen über das Sexualleben oder die sexuelle Orientierung einer Person. Falls sensible personenbezogene Daten gesammelt werden müssen, kontaktiere zuerst das Team unter <a href="mailto:privacy@deloitte.de" style={{ color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600 }}>privacy@deloitte.de</a>.</>
                  : <>That means: no data on race or ethnic origin, religious or philosophical beliefs, trade-union membership, political opinions, medical or health conditions, or information about a person&apos;s sex life or sexual orientation. If sensitive personal data must be collected, contact the team first at <a href="mailto:privacy@deloitte.de" style={{ color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600 }}>privacy@deloitte.de</a>.</>}
              </WizardHint>

              {/* v26.48: B2Run-Köln-Vorlage — Vorschlags-Box, erscheint nur
                  wenn der Event-Titel „B2Run Köln" enthält (greift auch im
                  Edit-Modus, da `title` aus editEvent vorbefüllt ist). Rein
                  additiv: ergänzt fehlende Template-Felder per Klick, der
                  bestehende Suggested-Felder-Katalog bleibt unberührt. */}
              {isB2RunKoelnTitle(title) && (() => {
                const b2rkTemplate = b2runKoelnTemplateFields(isDe);
                const b2rkMissing = b2rkTemplate.filter(f => !customFields.some(p => p.id === f.id));
                const b2rkTypeTag = (ty: CustomField['type']): string =>
                  ty === 'select' ? (isDe ? 'Auswahl' : 'Select') : ty === 'checkbox' ? 'Checkbox' : 'Text';
                return (
                  <div className="form-group" style={{ marginBottom: 16, padding: 16, background: 'rgba(134,188,37,0.08)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-green)' }}>
                    <label className="form-label" style={{ marginBottom: 4 }}>
                      {isDe ? 'B2Run-Köln-Vorlage' : 'B2Run Köln template'}
                    </label>
                    <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600, #4b5563)', marginTop: 0, marginBottom: 12 }}>
                      {isDe
                        ? 'Dieses Event sieht nach dem B2Run Köln aus — übernimm die offiziellen Meldefelder mit einem Klick. Der Excel-Export im Organizer Center füllt damit die offizielle Meldedatei exakt aus.'
                        : 'This event looks like the B2Run Köln — adopt the official entry fields with one click. The Excel export in the Organizer Center then fills in the official entry file exactly.'}
                    </p>
                    <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600, #4b5563)', marginTop: -6, marginBottom: 12 }}>
                      {isDe
                        ? <><strong>Startblock:</strong> wird NICHT als Feld abgefragt — die Teilnehmenden wählen Durchstarter/Funstarter über die <strong>Gruppen-Auswahl</strong> (Split-Kapazität im Schritt &bdquo;Kapazität &amp; Sichtbarkeit&ldquo;). Der Export übersetzt die Gruppe automatisch in die offiziellen Startblock-Texte.</>
                        : <><strong>Start block:</strong> is NOT asked as a field — attendees pick Durchstarter/Funstarter via the <strong>group selection</strong> (split capacity in the &ldquo;Capacity &amp; visibility&rdquo; step). The export automatically translates the group into the official start-block texts.</>}
                    </p>
                    <ul style={{ listStyle: 'none', margin: '0 0 12px', padding: 0, display: 'flex', flexDirection: 'column', gap: 4 }}>
                      {b2rkTemplate.map(f => {
                        const exists = customFields.some(p => p.id === f.id);
                        return (
                          <li key={f.id} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.8rem' }}>
                            <span style={{ width: 16, display: 'inline-flex', justifyContent: 'center', color: 'var(--dex-green)' }} title={exists ? (isDe ? 'Bereits im Event' : 'Already in the event') : undefined}>
                              {exists ? <Check size={14} /> : null}
                            </span>
                            <span style={{ fontWeight: exists ? 400 : 500 }}>{f.label}</span>
                            <span style={{ fontSize: '0.68rem', padding: '1px 8px', borderRadius: 999, background: '#fff', border: '1px solid var(--dex-green)', color: 'var(--dex-green-dark, #15803d)', whiteSpace: 'nowrap' }}>
                              {b2rkTypeTag(f.type)}
                            </span>
                          </li>
                        );
                      })}
                    </ul>
                    <button
                      type="button"
                      className="btn"
                      disabled={b2rkMissing.length === 0}
                      onClick={() => setCustomFields(prev => [
                        ...prev,
                        ...b2runKoelnTemplateFields(isDe)
                          .filter(f => !prev.some(p => p.id === f.id))
                          .map(f => ({ ...f, options: f.options || [] })),
                      ])}
                      style={{
                        fontSize: '0.85rem', padding: '6px 14px', border: 'none', color: '#fff',
                        background: b2rkMissing.length === 0 ? 'var(--dex-gray-300, #d1d5db)' : 'var(--dex-green)',
                        cursor: b2rkMissing.length === 0 ? 'default' : 'pointer',
                      }}
                    >
                      {b2rkMissing.length === 0
                        ? (isDe ? 'Alles übernommen' : 'All adopted')
                        : (isDe ? 'Alle übernehmen' : 'Adopt all')}
                    </button>
                  </div>
                );
              })()}

              {/* v18.57: Anrede-Abfrage-Toggle nach unten verschoben — sitzt jetzt
                  direkt unter den „Vorgeschlagene Felder"-Buttons. */}

              {/* v18.35: Anmeldesprache vorgeben — v22.32: nach unten zu den
                  Formular-Optionen (Anrede / Deutsch+Englisch) verschoben,
                  gleicher Look + gleiche Schriftgrößen wie die Toggles dort. */}

              {/* v18.57: Deutsch/Englisch-Toggle nach unten verschoben — sitzt jetzt
                  direkt unter den „Vorgeschlagene Felder"-Buttons. */}

              {/* v10.21: Template-Dropdown ist entfallen — der Organizer
                  pickt B2Run-Felder einzeln per Suggested-Felder-Modal
                  (eingeklappte Sektion "B2Run-spezifische Felder"). Damit
                  führt kein Weg mehr über ein hartes B2Run-Template, das
                  zusätzlich Logik (Auto-Split-Capacity etc.) auslöste —
                  saubere Trennung zwischen Feld-Konfiguration und
                  Kapazitäts-Modell. */}

              {/* B2Run Startblöcke - moderne Liste mit + Button. Wird
                  unverändert angezeigt, sobald das b2run_startblock-Feld in
                  customFields steht (über das Suggested-Felder-Modal
                  ausgewählt oder beim Edit eines Legacy-Events vorhanden). */}
              {customFields.some(f => f.id === 'b2run_startblock') && (
                <div className="form-group" style={{ marginBottom: 24, padding: 16, background: 'rgba(134,188,37,0.08)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-green)' }}>
                  <label className="form-label" style={{ marginBottom: 4 }}>
                    {t('create.startblocks')}
                    <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> die <strong>Startblöcke</strong>, in denen Teilnehmer ihren Lauf starten — z.B. Block A (schnell), Block B (mittel), Block C (Walking). Pro Block ein eigener Eintrag.<br /><br />
                      <strong>Anzeige in der App:</strong> bei der Anmeldung erscheint ein <strong>Dropdown Startblock</strong>, das diese Liste enthält. Falls du oben in Schritt 4 eine Starter-Typ-Zuordnung gemacht hast, ist das Dropdown automatisch gefüllt und disabled.<br /><br />
                      <strong>Auswirkung für Teilnehmer:</strong> der gewählte Startblock landet in der Bestätigungs-Mail und im Admin Center — wichtig damit ihr beim Veranstalter wisst, in welcher Welle wer startet.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>start blocks</strong> in which attendees begin their run — e.g. block A (fast), block B (medium), block C (walking). One entry per block.<br /><br />
                      <strong>Shown in the app:</strong> at registration, a <strong>start-block dropdown</strong> appears that contains this list. If you set up a starter-type mapping in step 4 above, the dropdown is auto-filled and disabled.<br /><br />
                      <strong>Effect for attendees:</strong> the selected start block ends up in the confirmation mail and in the admin center — important so you know which wave each attendee is in when coordinating with the organiser.
                    </>
                  )} />
                  </label>
                  <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                    {t('create.startblocks.hint')}
                  </p>

                  {/* Bestehende Startblöcke als Liste */}
                  {b2runStartblocks.length > 0 && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 12 }}>
                      {b2runStartblocks.map((block, idx) => (
                        <div
                          key={idx}
                          style={{
                            display: 'flex', alignItems: 'center', gap: 8,
                            padding: '8px 12px', borderRadius: 'var(--dex-radius, 12px)',
                            background: '#fff', border: '1px solid var(--dex-gray-200)',
                          }}
                        >
                          <Icon iconName="Running" style={{ fontSize: 16, color: 'var(--dex-green-dark, #6b9a1e)', flexShrink: 0 }} />
                          <span style={{ flex: 1, fontSize: '0.88rem' }}>{block}</span>
                          <button
                            type="button"
                            onClick={() => removeStartblock(block)}
                            title={t('create.startblocks.remove')}
                            style={{
                              border: 'none', background: 'transparent', cursor: 'pointer',
                              color: 'var(--dex-red, #c00)', padding: 4, borderRadius: 4,
                              display: 'inline-flex', alignItems: 'center',
                            }}
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}

                  {/* Neues Startblock hinzufügen */}
                  <div style={{ display: 'flex', gap: 8 }}>
                    <input
                      type="text"
                      className="form-input"
                      value={newStartblock}
                      onChange={e => setNewStartblock(e.target.value)}
                      onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); addStartblock(); } }}
                      placeholder={t('create.startblocks.placeholder')}
                      style={{ flex: 1 }}
                    />
                    <button
                      type="button"
                      className="btn btn-primary"
                      onClick={addStartblock}
                      disabled={!newStartblock.trim()}
                      style={{ whiteSpace: 'nowrap' }}
                    >
                      <Plus size={14} /> {t('create.startblocks.add')}
                    </button>
                  </div>
                </div>
              )}

              {/* v15.0: pro-Sub-Event-Tabs für Felder. Tab 0 = Haupt-Event
                  (komplette Custom-Fields-Liste, B2Run-Startblöcke etc.).
                  Tabs N>0 = schlanke per-Sub-Event-Felder-UI mit Inheritance-
                  Toggle. Im subEventsOnlyMode wird Tab 0 zu „Übergreifende
                  Felder" / „Cross-cutting fields" — die wirken dann auf alle
                  Sub-Event-Anmeldungen. */}
              {/* v28.78: Der Scope-Umschalter steht jetzt global unter der
                  Schritt-Leiste (renderGlobalScopeBar) — nicht mehr je Schritt. */}

              {activeFieldsTabIdx > 0 && (() => {
                const seIdx = activeFieldsTabIdx - 1;
                const se = subEvents[seIdx];
                if (!se) return null;
                const inherit = false;  // v15.3: inheritance entfernt — Sub-Events haben eigene Felder
                const seFields = se.customFields || [];
                const updateSub = (patch: Partial<SubEventDraft>): void => {
                  setSubEvents(prev => prev.map((x, i) => i === seIdx ? { ...x, ...patch } : x));
                };
                return (
                  <div>
                    {/* v15.6: Lead-paragraph analog Hauptevent-Tab. */}
                    <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                      {isDe
                        ? <><strong>Optional</strong> — die Standard-Teilnehmerdaten (Vorname, Nachname, E-Mail) und Profil-Daten (Job Title, Standort, Department, Telefon) werden automatisch erfasst. Hier ergänzt du <strong>nur Zusatzfragen speziell für dieses Sub-Event</strong>. Wenn das Sub-Event keine eigenen Fragen braucht, kannst du diese Sektion leer lassen.</>
                        : <><strong>Optional</strong> — the standard attendee data (first name, last name, email) and profile data (job title, location, department, phone) are captured automatically. Here you only add <strong>extra questions specific to this sub-event</strong>. If the sub-event needs no extra questions, you can leave this section empty.</>}
                    </p>

                    {/* v18.62: Datenschutz-Hinweis hier ENTFERNT — er steht bereits
                        einmal oben in Schritt 5 (über der Tab-Leiste). Eine
                        Wiederholung pro Sub-Event-Tab ist redundant. */}

                    {/* v15.3: „Anrede abfragen"-Toggle pro Sub-Event */}
                    <div style={{
                      background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                      padding: '14px 16px', marginBottom: 16,
                      border: '1px solid var(--dex-gray-200)',
                      display: 'flex', alignItems: 'center', gap: 10, justifyContent: 'space-between', flexWrap: 'wrap',
                    }}>
                      <label style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={!!se.askSalutation}
                          onChange={e => updateSub({ askSalutation: e.target.checked })}
                          style={{ cursor: 'pointer' }}
                        />
                        <span style={{ fontSize: '0.9rem', fontWeight: 600 }}>
                          {isDe ? 'Anrede für dieses Sub-Event abfragen' : 'Ask for salutation on this sub-event'}
                        </span>
                      </label>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 14px' }}
                        onClick={() => updateSub({ askSalutation: askSalutation })}
                        title={isDe
                          ? 'Übernimmt die Anrede-Abfrage-Einstellung vom Hauptevent'
                          : 'Copies the salutation toggle from the main event'}
                      >
                        {isDe ? 'Vom Hauptevent kopieren' : 'Copy from main event'}
                      </button>
                    </div>
                    <div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12, flexWrap: 'wrap' }}>
                        <strong style={{ fontSize: '0.95rem' }}>
                          {se.title || (isDe ? '(unbenanntes Sub-Event)' : '(unnamed sub-event)')}
                        </strong>
                        <button
                          type="button"
                          className="btn btn-outline"
                          style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                          onClick={() => addSubEventCustomField(se.id)}
                        >
                          <Plus size={12} /> {isDe ? 'Feld hinzufügen' : 'Add field'}
                        </button>
                        {customFields.length > 0 && seFields.length === 0 && (
                          <button
                            type="button"
                            className="btn btn-outline"
                            style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                            onClick={() => copyParentFieldsToSubEvent(se.id)}
                            title={isDe ? `Dupliziert die ${customFields.length} Felder vom Hauptevent als Startpunkt` : 'Duplicates the main-event fields as a starting point'}
                          >
                            {isDe ? `Felder vom Hauptevent kopieren (${customFields.length})` : `Copy fields from main event (${customFields.length})`}
                          </button>
                        )}
                      </div>
                      {seFields.length === 0 ? (
                        <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)', fontStyle: 'italic' }}>
                          {isDe
                            ? 'Keine zusätzlichen Felder definiert. Du kannst Felder hinzufügen oder vom Hauptevent kopieren.'
                            : 'No additional fields defined. You can add fields or copy them from the main event.'}
                        </p>
                      ) : (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                          {seFields.map((field, idx) => (
                            <div
                              key={field.id}
                              style={{
                                background: 'var(--dex-gray-50, #fafafa)',
                                borderRadius: 12,
                                padding: 16,
                                border: '1px solid var(--dex-gray-200)',
                              }}
                            >
                              <div style={{ display: 'flex', gap: 10, alignItems: 'center', marginBottom: 12, flexWrap: 'wrap' }}>
                                <span style={{
                                  flexShrink: 0, width: 26, height: 26, borderRadius: '50%',
                                  background: 'var(--dex-green, #86bc25)', color: '#fff',
                                  display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                  fontWeight: 700, fontSize: '0.78rem',
                                }}>{idx + 1}</span>
                                <input
                                  className="form-input"
                                  value={field.label}
                                  placeholder={isDe ? 'Frage / Feld-Label (z.B. „Welche Strecke?")' : 'Question / field label (e.g. „Which distance?")'}
                                  onChange={e => updateSubEventCustomField(se.id, field.id, { label: e.target.value })}
                                  disabled={inherit}
                                  style={{
                                    flex: '0 1 320px', minWidth: 180, maxWidth: 320,
                                    fontSize: '1rem', fontWeight: 600,
                                    padding: '8px 12px',
                                    color: field.label ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
                                  }}
                                />
                                <select
                                  className="form-select"
                                  value={field.type}
                                  disabled={inherit}
                                  onChange={e => updateSubEventCustomField(se.id, field.id, { type: e.target.value as CustomFieldInput['type'] })}
                                  title={isDe ? 'Art des Feldes' : 'Field type'}
                                  style={{
                                    flex: '0 0 200px', maxWidth: 200,
                                    background: 'rgba(134,188,37,0.08)',
                                    border: '1px solid var(--dex-green, #86bc25)',
                                    color: 'var(--dex-green-dark, #4a7c1f)',
                                    fontWeight: 600,
                                    padding: '8px 10px',
                                  }}
                                >
                                  <option value="text">{isDe ? 'Text (Freitext)' : 'Text (free text)'}</option>
                                  <option value="select">{isDe ? 'Dropdown' : 'Dropdown'}</option>
                                  <option value="number">{isDe ? 'Zahl' : 'Number'}</option>
                                  <option value="checkbox">{isDe ? 'Checkbox' : 'Checkbox'}</option>
                                  <option value="date">{isDe ? 'Datum (Kalender)' : 'Date (calendar)'}</option>
                        <option value="daterange">{isDe ? 'Übernachtungs-Zeitraum (Kalender + Nächte)' : 'Stay period (calendar + nights)'}</option>
                                </select>
                                <label
                                  style={{
                                    display: 'inline-flex', alignItems: 'center', gap: 6,
                                    padding: '6px 12px', borderRadius: 999,
                                    fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap',
                                    cursor: inherit ? 'not-allowed' : 'pointer', userSelect: 'none',
                                    border: `1px solid ${field.required ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                    background: field.required ? 'rgba(134,188,37,0.10)' : '#fff',
                                    color: field.required ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                                  }}
                                >
                                  <input
                                    type="checkbox"
                                    checked={field.required}
                                    disabled={inherit}
                                    onChange={e => updateSubEventCustomField(se.id, field.id, { required: e.target.checked })}
                                    style={{ display: 'none' }}
                                  />
                                  <span style={{ fontSize: '0.85rem', lineHeight: 1 }}>{field.required ? '✓' : '○'}</span>
                                  {t('create.required')}
                                </label>
                                <button
                                  type="button"
                                  onClick={() => removeSubEventCustomField(se.id, field.id)}
                                  disabled={inherit}
                                  title={isDe ? 'Feld entfernen' : 'Remove field'}
                                  style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4, cursor: inherit ? 'not-allowed' : 'pointer', flexShrink: 0 }}
                                >
                                  <X size={18} />
                                </button>
                              </div>
                              {/* v24.25: Datum-Feld → optional Uhrzeit mit abfragen. */}
                              {field.type === 'date' && (
                                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, marginLeft: 32, marginTop: 8, cursor: inherit ? 'default' : 'pointer', fontSize: '0.82rem', color: 'var(--dex-gray-700)' }}>
                                  <input
                                    type="checkbox"
                                    checked={!!field.withTime}
                                    disabled={inherit}
                                    onChange={e => updateSubEventCustomField(se.id, field.id, { withTime: e.target.checked })}
                                    style={{ accentColor: 'var(--dex-green, #86bc25)' }}
                                  />
                                  {isDe ? 'Auch Uhrzeit abfragen?' : 'Also ask for the time?'}
                                </label>
                              )}
                              {field.type === 'daterange' && (
                                <div style={{ marginLeft: 32, marginTop: 8, display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'flex-end' }}>
                                  <div>
                                    <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>{isDe ? 'Buchbar ab' : 'Bookable from'}</div>
                                    <input type="date" className="form-input" disabled={inherit} style={{ padding: '6px 10px', fontSize: '0.85rem', width: 160 }}
                                      value={field.rangeStart || ''} onChange={e => updateSubEventCustomField(se.id, field.id, { rangeStart: e.target.value })} />
                                  </div>
                                  <div>
                                    <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>{isDe ? 'Buchbar bis' : 'Bookable until'}</div>
                                    <input type="date" className="form-input" disabled={inherit} style={{ padding: '6px 10px', fontSize: '0.85rem', width: 160 }}
                                      value={field.rangeEnd || ''} onChange={e => updateSubEventCustomField(se.id, field.id, { rangeEnd: e.target.value })} />
                                  </div>
                                  <div>
                                    <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>{isDe ? 'Max. Nächte' : 'Max. nights'}</div>
                                    <input type="number" min={0} className="form-input" disabled={inherit} style={{ padding: '6px 10px', fontSize: '0.85rem', width: 100 }}
                                      placeholder={isDe ? 'offen' : 'open'}
                                      value={field.maxNights || ''} onChange={e => updateSubEventCustomField(se.id, field.id, { maxNights: parseInt(e.target.value, 10) || 0 })} />
                                  </div>
                                </div>
                              )}
                              {/* v24.25: Feldart-Empfehlung (nur Datum — Sub-Event-Felder
                                  kennen den People-Picker-Typ nicht). */}
                              <FieldTypeSuggestion
                                field={field}
                                isDe={isDe}
                                allowPerson={false}
                                disabled={inherit}
                                onApply={(t) => updateSubEventCustomField(se.id, field.id, { type: t })}
                              />
                              <div style={{ marginLeft: 32, marginTop: 10 }}>
                                <input
                                  className="form-input"
                                  placeholder={isDe
                                    ? 'Beschreibung (optional, erscheint als „i"-Tooltip neben dem Feld)'
                                    : 'Description (optional, shown as „i" tooltip next to the field)'}
                                  value={field.helpText || ''}
                                  disabled={inherit}
                                  onChange={e => updateSubEventCustomField(se.id, field.id, { helpText: e.target.value })}
                                  style={{ width: '100%', fontSize: '0.82rem', padding: '6px 10px' }}
                                />
                                {/* v24.14 BUG-FIX: helpTextStyle-Wahl fehlte bei Sub-Event-Feldern
                                    (nur das Hauptevent hatte sie) — „Text unter dem Feld" wurde
                                    deshalb beim Sub-Event nie gespeichert. */}
                                {field.helpText && field.helpText.trim() && (
                                  <div style={{ display: 'flex', gap: 16, marginTop: 6, fontSize: '0.78rem', color: 'var(--dex-gray-600)', flexWrap: 'wrap' }}>
                                    <span style={{ fontWeight: 600 }}>{isDe ? 'Anzeige:' : 'Display:'}</span>
                                    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 5, cursor: inherit ? 'default' : 'pointer' }}>
                                      <input
                                        type="radio"
                                        name={`seHelpStyle-${se.id}-${field.id}`}
                                        disabled={inherit}
                                        checked={(field.helpTextStyle || 'tooltip') !== 'inline'}
                                        onChange={() => updateSubEventCustomField(se.id, field.id, { helpTextStyle: 'tooltip' })}
                                      />
                                      {isDe ? '„i"-Info-Box (Hover)' : '„i" info box (hover)'}
                                    </label>
                                    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 5, cursor: inherit ? 'default' : 'pointer' }}>
                                      <input
                                        type="radio"
                                        name={`seHelpStyle-${se.id}-${field.id}`}
                                        disabled={inherit}
                                        checked={field.helpTextStyle === 'inline'}
                                        onChange={() => updateSubEventCustomField(se.id, field.id, { helpTextStyle: 'inline' })}
                                      />
                                      {isDe ? 'Text unter dem Feld-Titel' : 'Text below the field title'}
                                    </label>
                                  </div>
                                )}
                              </div>
                              {field.type === 'select' && (
                                <div style={{
                                  marginTop: 10, marginLeft: 32, padding: '12px 14px',
                                  background: '#fff',
                                  border: '1px solid var(--dex-gray-200)',
                                  borderRadius: 8,
                                }}>
                                  <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-700)', fontWeight: 600, marginBottom: 8 }}>
                                    {isDe ? 'Antwort-Optionen' : 'Answer options'}
                                  </div>
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                                    {field.options.map((opt, oidx) => (
                                      <div key={oidx} style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                                        <span style={{
                                          flexShrink: 0, width: 22, height: 22, borderRadius: '50%',
                                          background: 'var(--dex-gray-200)', color: 'var(--dex-gray-700)',
                                          display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                          fontWeight: 700, fontSize: '0.72rem',
                                        }}>{oidx + 1}</span>
                                        <input
                                          className="form-input"
                                          placeholder={isDe ? `Option ${oidx + 1}` : `Option ${oidx + 1}`}
                                          value={opt}
                                          disabled={inherit}
                                          onChange={e => {
                                            const next = field.options.slice();
                                            next[oidx] = e.target.value;
                                            updateSubEventCustomField(se.id, field.id, { options: next });
                                          }}
                                          style={{ flex: 1, fontSize: '0.85rem', padding: '6px 10px' }}
                                        />
                                        {field.options.length > 1 && (
                                          <button
                                            type="button"
                                            disabled={inherit}
                                            onClick={() => {
                                              const next = field.options.filter((_, i) => i !== oidx);
                                              updateSubEventCustomField(se.id, field.id, { options: next });
                                            }}
                                            title={isDe ? 'Option entfernen' : 'Remove option'}
                                            style={{ background: 'none', border: 'none', color: 'var(--dex-gray-500)', padding: 4, cursor: inherit ? 'not-allowed' : 'pointer' }}
                                          >
                                            <X size={14} />
                                          </button>
                                        )}
                                      </div>
                                    ))}
                                    <button
                                      type="button"
                                      disabled={inherit}
                                      onClick={() => updateSubEventCustomField(se.id, field.id, { options: [...field.options, ''] })}
                                      style={{ alignSelf: 'flex-start', background: 'none', border: '1px dashed var(--dex-gray-300)', padding: '4px 12px', fontSize: '0.78rem', borderRadius: 6, cursor: inherit ? 'not-allowed' : 'pointer', color: 'var(--dex-gray-700)', display: 'inline-flex', alignItems: 'center', gap: 4, marginTop: 4 }}
                                    >
                                      <Plus size={12} /> {isDe ? 'Option hinzufügen' : 'Add option'}
                                    </button>
                                  </div>
                                </div>
                              )}
                              {/* v24.16 BUG-FIX: Sichtbarkeitsbedingung (showIf)
                                  fehlte bei Sub-Event-Feldern — bedingte Fragen
                                  wurden deshalb auf dem Anmeldeformular IMMER
                                  angezeigt. Gleiche UI wie beim Hauptevent. */}
                              {!inherit && renderShowIfConfig(field, idx, seFields, (u) => updateSubEventCustomField(se.id, field.id, u))}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                );
              })()}

              {/* Tab 0 (Haupt-Event bzw. Übergreifende Felder): die
                  bestehende Hauptevent-Felder-UI bleibt unverändert; nur
                  die Sektion „Felder pro Sub-Event" weiter unten wird in
                  v15.0 ausgeblendet, weil pro Sub-Event jetzt einen
                  eigenen Tab. */}
              <div style={{ display: activeFieldsTabIdx === 0 ? 'block' : 'none' }}>
              {/* v16.5: In Step 5 (Felder) ist im subEventsOnlyMode KEIN
                  Greyout — die Felder im ersten Tab sind „übergreifend"
                  und werden bei JEDER Sub-Event-Anmeldung abgefragt, also
                  in dem Modus besonders relevant. Stattdessen ein info-
                  blauer Hinweis-Banner mit der korrekten Erklärung. */}
              {/* v22.35: Die frühere „Was sind übergreifende Felder?"-Box ist
                  entfallen — die Sektions-Beschreibung direkt darunter erklärt
                  dasselbe (Redundanz). */}
              <div>
              {/* Dynamische Felder */}
              <div>
                {/* Bereich-Header: trennt Hauptevent-Felder visuell vom
                    Sub-Event-Block weiter unten (v10.11+).
                    v15.0: im subEventsOnlyMode lautet die Überschrift
                    „Übergreifend für alle <childTermPlural>". */}
                <h3 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.05rem', fontWeight: 700 }}>
                  {subEventsOnlyMode
                    ? (isDe
                        ? `Übergreifend für alle ${(childTermPlural || 'Sub-Events').trim() || 'Sub-Events'}`
                        : `Across all ${(childTermPlural || 'sub-events').trim() || 'sub-events'}`)
                    : (isDe ? 'Felder für das Hauptevent' : 'Fields for the main event')}
                </h3>
                <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                  {isDe
                    ? 'Diese Felder werden bei jeder Anmeldung abgefragt — egal ob das Event Sub-Events hat oder nicht. Für Sub-Event-spezifische Fragen wechsle oben auf den jeweiligen Sub-Event-Tab.'
                    : 'These fields are asked at every registration — regardless of whether the event has sub-events. For sub-event-specific questions switch to the respective sub-event tab above.'}
                </p>
                {/* v22.38: Sub-Überschrift „Einstellungen" — die generellen
                    Formular-Optionen (Zweisprachigkeit, Formular-Sprache)
                    stehen VOR der Feld-Liste. „Eigene Abfragen / Felder"
                    sitzt weiter unten DIREKT über den Feld-Zeilen. Der
                    frühere „Anrede abfragen?"-Toggle ist als vorgeschlagenes
                    Feld in den Katalog gewandert (Sonder-Key 'salutation'). */}
                <h3 style={{ margin: '0 0 10px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.05rem', fontWeight: 700 }}>
                  {isDe ? 'Einstellungen' : 'Settings'}
                </h3>

              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '12px 16px', marginBottom: 14,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={bilingualFields}
                    onChange={e => setBilingualFields(e.target.checked)}
                    style={{ marginTop: 3, cursor: 'pointer' }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Deutsch und Englisch ermöglichen' : 'Offer German and English'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob du pro Custom-Field, das du unten anlegst, <strong>eine englische Variante</strong> der Texte hinterlegen kannst — also Feld-Name, Beschreibung (i-Tooltip), Checkbox-Bestätigungs-Text und Dropdown-Optionen jeweils auf Deutsch UND auf Englisch. Default: <strong>aus</strong>.<br /><br />
                          <strong>Anzeige in der App:</strong> wenn aktiviert, blendet jede Feld-Karte einen zweiten Eingabe-Block für die EN-Variante ein. Teilnehmer mit App-Sprache <strong>Englisch</strong> bekommen automatisch die EN-Texte zu sehen. Wer als App-Sprache Deutsch eingestellt hat, sieht weiterhin die DE-Texte. Zusätzlich folgt das Standard-Anmelde-Formular (Platzhalter, Hinweis-Boxen, Sub-Event-Sektion) ab dann der <strong>App-Spracheinstellung des Teilnehmers</strong> statt der Mail-Sprache des Events.<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> internationale Kolleg:innen, die kein Deutsch sprechen, sehen das komplette Anmelde-Formular sauber auf Englisch. Wer als Organizer keine EN-Variante einträgt, fällt im EN-Modus still auf den DE-Wert zurück — die App bricht also nichts kaputt, falls du nur einige Felder übersetzt.
                        </>
                      : <>
                          <strong>What this controls:</strong> whether, for each custom field you create below, you can store <strong>an English variant</strong> of the texts — i.e. field name, description (i-tooltip), checkbox confirmation text and dropdown options in both German AND English. Default: <strong>off</strong>.<br /><br />
                          <strong>Where you see it:</strong> when enabled, each field card shows a second input row for the EN variant. Attendees with app language set to <strong>English</strong> automatically see the EN texts. Attendees with German keep seeing the DE texts. In addition, the standard registration form chrome (placeholders, hint boxes, sub-event section) follows the <strong>attendee&apos;s app language</strong> instead of the event&apos;s email language.<br /><br />
                          <strong>For attendees:</strong> international colleagues who do not speak German see the whole registration form cleanly in English. If an organizer leaves the EN variant empty for some field, the app silently falls back to the DE value — nothing breaks if you only translate a subset of fields.
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? <>Default: aus — wenn aktiviert, hinterlegst du deine Fragen <strong>zweisprachig</strong> (deutsche UND englische Texte pro Feld); Teilnehmer sehen automatisch ihre Sprache. Stellst du deine Fragen nur <strong>einsprachig</strong> (z.&nbsp;B. nur Englisch), lass den Schalter aus und stelle unten die Formular-Sprache fest ein.</>
                        : <>Default: off — when enabled, you maintain your questions <strong>bilingually</strong> (German AND English texts per field); attendees automatically see their language. If your questions are <strong>monolingual</strong> (e.g. English only), leave this off and fix the form language below instead.</>}
                    </span>
                  </span>
                </label>
              </div>

              {/* v22.32: Sprache des Anmeldeformulars — gehört inhaltlich zu
                  den Formular-Optionen hier (vorher eigene Karte weiter oben),
                  gleicher Look + gleiche Schriftgrößen wie die Toggles. */}
              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '12px 16px', marginBottom: 14,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <span style={{ display: 'block' }}>
                  <strong>{isDe ? 'Sprache des Anmeldeformulars' : 'Registration form language'}</strong>
                  <InfoTooltip text={isDe
                    ? <>
                        <strong>Was du hier einstellst:</strong> in welcher Sprache die <strong>komplette Anmeldeseite</strong> (alle Texte, Buttons und der <strong>Datenschutz-Disclaimer</strong>) angezeigt wird.<br /><br />
                        <strong>Anzeige in der App:</strong> bei <strong>Automatisch</strong> folgt die Anmeldeseite der App-Sprache des Teilnehmers. Wählst du <strong>Immer Deutsch</strong> oder <strong>Immer Englisch</strong>, wird die Anmeldeseite <strong>fest in dieser Sprache</strong> angezeigt — auch wenn der Teilnehmer die App z.&nbsp;B. auf Deutsch nutzt. Ein kleiner Hinweis im Kopfbereich zeigt das an.<br /><br />
                        <strong>Unsere Empfehlung:</strong> Stellst du deine eigenen Fragen unten <strong>nur in einer Sprache</strong> (z.&nbsp;B. Englisch), dann stelle das Formular <strong>fest auf diese Sprache</strong>. Sonst mischen sich bei Teilnehmern mit anderer App-Sprache die deutschen Standard-Texte (Buttons, Hinweise, Datenschutz) mitten zwischen deine englischen Fragen — das wirkt unsauber.<br /><br />
                        <strong>Auswirkung für Teilnehmer:</strong> bei einem englischsprachigen Event sehen sie die Anmeldung samt Disclaimer komplett auf Englisch, egal welche App-Sprache eingestellt ist.
                      </>
                    : <>
                        <strong>What you set here:</strong> the language in which the <strong>entire registration page</strong> (all texts, buttons and the <strong>privacy disclaimer</strong>) is shown.<br /><br />
                        <strong>Where you see it:</strong> with <strong>Automatic</strong> the page follows the attendee&apos;s app language. Choosing <strong>Always German</strong> or <strong>Always English</strong> forces the registration page into that language — even if the attendee uses the app in another language. A small hint in the header indicates this.<br /><br />
                        <strong>Our recommendation:</strong> if your own questions below are written in <strong>one language only</strong> (e.g. English), <strong>fix the form to that language</strong>. Otherwise attendees with a different app language get the German standard texts (buttons, hints, privacy) mixed in between your English questions — which looks messy.<br /><br />
                        <strong>For attendees:</strong> for an English-language event they see the registration and disclaimer fully in English regardless of their app language.
                      </>
                  } />
                  <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                    {isDe
                      ? <>Default: Automatisch — folgt der App-Sprache des Teilnehmers. <strong>Empfehlung:</strong> Stellst du deine Fragen nur in einer Sprache (z.&nbsp;B. Englisch), stelle das Formular fest auf diese Sprache — sonst stehen deutsche Standard-Texte mitten zwischen deinen englischen Fragen.</>
                      : <>Default: Automatic — follows the attendee&apos;s app language. <strong>Recommendation:</strong> if your questions are in one language only (e.g. English), fix the form to that language — otherwise German standard texts appear in between your English questions.</>}
                  </span>
                  <select
                    className="form-input"
                    value={registrationLanguage}
                    onChange={e => setRegistrationLanguage(e.target.value as '' | 'de' | 'en')}
                    style={{ width: '100%', maxWidth: 380, marginTop: 8, minHeight: 0, padding: '8px 12px', fontSize: '0.85rem' }}
                  >
                    <option value="">{isDe ? 'Automatisch (App-Sprache des Teilnehmers)' : 'Automatic (attendee\'s app language)'}</option>
                    <option value="de">{isDe ? 'Immer Deutsch' : 'Always German'}</option>
                    <option value="en">{isDe ? 'Immer Englisch' : 'Always English'}</option>
                  </select>
                </span>
              </div>

                {/* v22.38: „Eigene Abfragen / Felder" sitzt jetzt DIREKT über
                    der Feld-Liste (vorher standen die Einstellungs-Karten
                    dazwischen). */}
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12, marginTop: 6 }}>
                  <StepBadge n={24} />
                  {isDe ? 'Eigene Abfragen / Felder' : 'Custom fields'}
                </label>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 16, flexWrap: 'wrap' }}>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    onClick={openSuggestedModal}
                    style={{ fontSize: '0.85rem', padding: '6px 14px' }}
                    title={isDe ? 'Felder aus einem Katalog wählen' : 'Pick fields from a catalog'}
                  >
                    {isDe ? 'Vorgeschlagene Felder' : 'Suggested fields'}
                  </button>
                  <button className="btn btn-outline" onClick={addCustomField} style={{ fontSize: '0.85rem', padding: '6px 14px' }}>
                    <Plus size={14} /> {t('create.addfield')}
                  </button>
                  {customFields.length > 1 && (
                    <button
                      type="button"
                      className={reorderMode ? 'btn btn-primary' : 'btn btn-outline'}
                      onClick={() => setReorderMode(prev => !prev)}
                      style={{ fontSize: '0.85rem', padding: '6px 14px' }}
                      title={isDe ? 'Felder per Hoch/Runter-Pfeile sortieren' : 'Reorder fields with up/down arrows'}
                    >
                      {reorderMode
                        ? (isDe ? 'Fertig' : 'Done')
                        : (isDe ? 'Reihenfolge ändern' : 'Reorder')}
                    </button>
                  )}
                </div>
                {/* v24.25: Erklär-Box zwischen den Feld-Buttons und der Feld-Liste —
                    welche Feldarten es gibt und was sie tun (aufklappbar, grau). */}
                <WizardHint
                  isDe={isDe}
                  variant="description"
                  title={isDe ? 'Welche Feldarten gibt es?' : 'Which field types are available?'}
                  style={{ marginBottom: 12 }}
                >
                  <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.5 }}>
                    <li><strong>{isDe ? 'Text (Freitext)' : 'Text (free text)'}</strong> — {isDe ? 'freie Eingabe, z.B. eine Anmerkung.' : 'free input, e.g. a note.'}</li>
                    <li><strong>{isDe ? 'Dropdown' : 'Dropdown'}</strong> — {isDe ? 'Auswahl aus festen Optionen; optional Mehrfachauswahl.' : 'pick from preset options; optionally multi-select.'}</li>
                    <li><strong>{isDe ? 'Zahl' : 'Number'}</strong> — {isDe ? 'nur Zahlen, z.B. eine Anzahl.' : 'numbers only, e.g. a quantity.'}</li>
                    <li><strong>{isDe ? 'Checkbox' : 'Checkbox'}</strong> — {isDe ? 'einfache Ja/Nein-Bestätigung.' : 'simple yes/no confirmation.'}</li>
                    <li><strong>{isDe ? 'Datum (Kalender)' : 'Date (calendar)'}</strong> — {isDe ? 'Datum über einen Kalender; optional zusätzlich die Uhrzeit.' : 'a date via a calendar; optionally with time.'}</li>
                    <li><strong>{isDe ? 'Person' : 'Person'}</strong> — {isDe ? 'Personensuche mit Foto und Standort; die gewählte Person kann optional die An-/Abmelde-Mail in Kopie (CC) bekommen.' : 'person search with photo and location; the chosen person can optionally be CC’d on the emails.'}</li>
                    <li><strong>{isDe ? 'Roommate' : 'Roommate'}</strong> — {isDe ? 'wie „Person“, löst zusätzlich eine Zimmerpartner-Mail an die gewählte Person aus.' : 'like „Person“, additionally triggers a roommate email to the selected person.'}</li>
                    <li><strong>{isDe ? 'Dokument (Upload)' : 'Document (upload)'}</strong> — {isDe ? 'Teilnehmer lädt eine Datei (PDF/Bild) hoch, die an die Anmeldung angehängt wird.' : 'attendee uploads a file (PDF/image) attached to the registration.'}</li>
                  </ul>
                  <p style={{ margin: '8px 0 0' }}>
                    {isDe
                      ? 'Pro Feld kannst du zusätzlich „Pflicht“ verlangen, eine Beschreibung hinterlegen (als „i“-Box oder als Text unter dem Feld) und eine Sichtbarkeitsbedingung setzen — das Feld erscheint dann nur, wenn eine andere Frage bestimmt beantwortet wurde.'
                      : 'Per field you can also require it, add a description (as an „i“ box or text below the field) and set a visibility condition — the field then only appears when another question has a specific answer.'}
                  </p>
                </WizardHint>
                {/* v22.38: Anrede als Standard-Feld-Zeile — aktiviert über das
                    Vorgeschlagene-Felder-Modal (Eintrag „Anrede"), entfernbar
                    über das X. Kein Custom-Field (eigener askSalutation-Flag). */}
                {askSalutation && (
                  <div style={{
                    display: 'flex', alignItems: 'center', gap: 10,
                    padding: '10px 16px', marginBottom: 12,
                    background: 'var(--dex-gray-50, #fafafa)',
                    borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)',
                  }}>
                    <span style={{ flexShrink: 0, width: 26, height: 26, borderRadius: '50%', background: 'var(--dex-green, #86bc25)', color: '#fff', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: '0.78rem', lineHeight: 1 }}>A</span>
                    <span style={{ flex: 1, fontSize: '0.88rem' }}>
                      <strong>{isDe ? 'Anrede' : 'Salutation'}</strong>{' '}
                      <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>
                        {isDe
                          ? '— Standard-Feld: Pflicht-Dropdown (Frau / Herr / Divers / Keine Angabe) über dem Vornamen.'
                          : '— standard field: required dropdown (Mrs / Mr / Diverse / Prefer not to say) above the first name.'}
                      </span>
                    </span>
                    <button
                      type="button"
                      onClick={() => setAskSalutation(false)}
                      title={isDe ? 'Anrede-Abfrage entfernen' : 'Remove salutation'}
                      style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)', padding: 4 }}
                    >
                      <X size={16} />
                    </button>
                  </div>
                )}
                {customFields.map((field, idx) => {
                  const isExpanded = !!fieldExpandOverride[field.id];
                  return (
                  <div
                    key={field.id}
                    draggable
                    onDragStart={() => setDragFieldId(field.id)}
                    onDragOver={(e) => { e.preventDefault(); setDragOverFieldId(field.id); }}
                    onDragLeave={() => { if (dragOverFieldId === field.id) setDragOverFieldId(null); }}
                    onDrop={() => {
                      if (dragFieldId && dragFieldId !== field.id) {
                        const fromIdx = customFields.findIndex(f => f.id === dragFieldId);
                        const toIdx = customFields.findIndex(f => f.id === field.id);
                        if (fromIdx >= 0 && toIdx >= 0) {
                          const updated = [...customFields];
                          const [moved] = updated.splice(fromIdx, 1);
                          updated.splice(toIdx, 0, moved);
                          setCustomFields(updated);
                        }
                      }
                      setDragFieldId(null);
                      setDragOverFieldId(null);
                    }}
                    onDragEnd={() => { setDragFieldId(null); setDragOverFieldId(null); }}
                    style={{
                      opacity: dragFieldId === field.id ? 0.4 : 1,
                      borderTop: dragOverFieldId === field.id ? '3px solid var(--dex-green)' : undefined,
                      background: 'var(--dex-gray-50, #fafafa)',
                      borderRadius: 12,
                      padding: 16,
                      marginBottom: 12,
                      border: '1px solid var(--dex-gray-200)',
                    }}
                  >
                    {/* v10.25: Konsolidierter Header — Label-Input prominent
                        als Titel + Typ-Dropdown rechts daneben + Pflicht-Pill
                        + Lösch-X. Reorder-Pfeile nur im Reorder-Modus. Die
                        separate Art/Beschriftung-Box ist entfallen, weil Typ
                        und Label hier direkt sichtbar / editierbar sind. */}
                    <div style={{ display: 'flex', gap: 10, alignItems: 'center', marginBottom: 12, flexWrap: 'wrap' }}>
                      {reorderMode && (
                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 1, padding: '0 4px' }}>
                          <button
                            type="button"
                            onClick={() => moveCustomField(field.id, 'up')}
                            disabled={idx === 0}
                            style={{ background: 'none', border: 'none', padding: 0, color: idx === 0 ? 'var(--dex-gray-300)' : 'var(--dex-gray-600)', cursor: idx === 0 ? 'default' : 'pointer', fontSize: '0.85rem', lineHeight: 1 }}
                            title={isDe ? 'Nach oben' : 'Move up'}
                          >▲</button>
                          <button
                            type="button"
                            onClick={() => moveCustomField(field.id, 'down')}
                            disabled={idx === customFields.length - 1}
                            style={{ background: 'none', border: 'none', padding: 0, color: idx === customFields.length - 1 ? 'var(--dex-gray-300)' : 'var(--dex-gray-600)', cursor: idx === customFields.length - 1 ? 'default' : 'pointer', fontSize: '0.85rem', lineHeight: 1 }}
                            title={isDe ? 'Nach unten' : 'Move down'}
                          >▼</button>
                        </div>
                      )}
                      <span style={{
                        flexShrink: 0, width: 26, height: 26, borderRadius: '50%',
                        background: 'var(--dex-green, #86bc25)', color: '#fff',
                        display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                        fontWeight: 700, fontSize: '0.78rem',
                      }}>{idx + 1}</span>
                      {/* v18.56: Textarea statt Input — lange Fragen brechen jetzt
                          um statt abgeschnitten zu werden. Auto-Höhe via ref
                          (height = scrollHeight). resize:none + overflow:hidden,
                          damit es wie ein wachsendes Eingabefeld wirkt. */}
                      <textarea
                        className="form-input"
                        value={field.label}
                        rows={1}
                        placeholder={isDe ? 'Feld-Name (z.B. T-Shirt Größe)' : 'Field name (e.g. T-shirt size)'}
                        onChange={e => updateCustomField(field.id, { label: e.target.value })}
                        ref={(el) => { if (el) { el.style.height = 'auto'; el.style.height = `${el.scrollHeight}px`; } }}
                        style={{
                          flex: '1 1 280px', minWidth: 180, maxWidth: 360,
                          // v22.30: minHeight 0 hebt die 48px-Mindesthöhe der
                          // .form-input-Klasse auf — die Auto-Höhe (scrollHeight)
                          // umschließt den Text dann exakt, der Feld-Name steht
                          // vertikal zentriert in der Box (vorher klebte er oben).
                          // Schrift dazu eine Stufe kleiner.
                          minHeight: 0,
                          fontSize: '0.9rem', fontWeight: 600,
                          padding: '10px 12px',
                          resize: 'none', overflow: 'hidden', lineHeight: 1.35,
                          fontFamily: 'inherit',
                          color: field.label ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
                        }}
                      />
                      <select
                        className="form-select"
                        value={field.type}
                        onChange={e => updateCustomField(field.id, { type: e.target.value as CustomFieldInput['type'] })}
                        title={isDe ? 'Art des Feldes' : 'Field type'}
                        style={{
                          /* v11.4: feste Breite, damit Label-Input + Typ-
                             Selector + Pflicht-Pill + X immer in einer Zeile
                             passen. Vorher: flex 0 0 auto = intrinsic Width
                             — bei langen Options wie 'Roommate (Doppelzimmer)'
                             wurde der Selector breiter und drückte das X in
                             eine zweite Zeile. */
                          flex: '0 0 200px', maxWidth: 200,
                          background: 'rgba(134,188,37,0.08)',
                          border: '1px solid var(--dex-green, #86bc25)',
                          color: 'var(--dex-green-dark, #4a7c1f)',
                          fontWeight: 600,
                          // v24.25: gleiche Höhe wie das Label-Eingabefeld
                          // (Textarea, padding 10px 12px, 0.9rem, minHeight 0).
                          padding: '10px 12px', fontSize: '0.9rem', minHeight: 0,
                        }}
                      >
                        <option value="text">{isDe ? 'Text (Freitext)' : 'Text (free text)'}</option>
                        <option value="select">{isDe ? 'Dropdown' : 'Dropdown'}</option>
                        <option value="number">{isDe ? 'Zahl' : 'Number'}</option>
                        <option value="checkbox">{isDe ? 'Checkbox' : 'Checkbox'}</option>
                        <option value="date">{isDe ? 'Datum (Kalender)' : 'Date (calendar)'}</option>
                        <option value="daterange">{isDe ? 'Übernachtungs-Zeitraum (Kalender + Nächte)' : 'Stay period (calendar + nights)'}</option>
                        <option value="user">{isDe ? 'Person' : 'Person'}</option>
                        <option value="roommate">{isDe ? 'Roommate' : 'Roommate'}</option>
                        <option value="document">{isDe ? 'Dokument (Upload)' : 'Document (upload)'}</option>
                      </select>
                      <label
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: 6,
                          padding: '6px 12px', borderRadius: 999,
                          fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap',
                          cursor: 'pointer', userSelect: 'none',
                          border: `1px solid ${field.required ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                          background: field.required ? 'rgba(134,188,37,0.10)' : '#fff',
                          color: field.required ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                          transition: 'all 0.15s ease',
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={field.required}
                          onChange={e => updateCustomField(field.id, { required: e.target.checked })}
                          style={{ display: 'none' }}
                        />
                        <span style={{ fontSize: '0.85rem', lineHeight: 1 }}>{field.required ? '✓' : '○'}</span>
                        {t('create.required')}
                      </label>
                      <button
                        type="button"
                        onClick={() => toggleFieldExpand(field.id, isExpanded)}
                        title={isExpanded ? (isDe ? 'Details einklappen' : 'Collapse details') : (isDe ? 'Details bearbeiten' : 'Edit details')}
                        aria-expanded={isExpanded}
                        style={{ background: 'none', border: 'none', color: 'var(--dex-gray-500)', padding: 4, cursor: 'pointer', flexShrink: 0, marginLeft: 'auto', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                      >
                        {!isExpanded && (
                          <span style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                            {isDe ? 'Details' : 'Details'}
                          </span>
                        )}
                        <Icon iconName={isExpanded ? 'ChevronUp' : 'ChevronDown'} style={{ fontSize: 14 }} />
                      </button>
                      <button
                        onClick={() => removeCustomField(field.id)}
                        title={isDe ? 'Feld löschen' : 'Delete field'}
                        style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4, cursor: 'pointer', flexShrink: 0 }}
                      >
                        <X size={18} />
                      </button>
                    </div>

                    {/* v24.25: Datum-Feld → optional auch die Uhrzeit abfragen
                        (immer sichtbar, nicht in „Details" versteckt). */}
                    {field.type === 'date' && (
                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, marginTop: 8, marginLeft: 32, cursor: 'pointer', fontSize: '0.82rem', color: 'var(--dex-gray-700)' }}>
                        <input
                          type="checkbox"
                          checked={!!field.withTime}
                          onChange={e => updateCustomField(field.id, { withTime: e.target.checked })}
                          style={{ accentColor: 'var(--dex-green, #86bc25)' }}
                        />
                        {isDe ? 'Auch Uhrzeit abfragen?' : 'Also ask for the time?'}
                      </label>
                    )}
                    {/* v28.63: Übernachtungs-Zeitraum — buchbares Fenster und Nächte-Limit.
                        Ohne Fenster kann der Teilnehmer jedes Datum wählen; mit Fenster
                        (z.B. 22.09.–26.09.) bleibt die Auswahl an eurem Kontingent. */}
                    {field.type === 'daterange' && (
                      <div style={{ marginLeft: 32, marginTop: 8, display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'flex-end' }}>
                        <div>
                          <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>
                            {isDe ? 'Buchbar ab' : 'Bookable from'}
                          </div>
                          <input type="date" className="form-input"
                            style={{ padding: '6px 10px', fontSize: '0.85rem', width: 160 }}
                            value={field.rangeStart || ''}
                            onChange={e => updateCustomField(field.id, { rangeStart: e.target.value })} />
                        </div>
                        <div>
                          <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>
                            {isDe ? 'Buchbar bis' : 'Bookable until'}
                          </div>
                          <input type="date" className="form-input"
                            style={{ padding: '6px 10px', fontSize: '0.85rem', width: 160 }}
                            value={field.rangeEnd || ''}
                            onChange={e => updateCustomField(field.id, { rangeEnd: e.target.value })} />
                        </div>
                        <div>
                          <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>
                            {isDe ? 'Max. Nächte' : 'Max. nights'}
                          </div>
                          <input type="number" min={0} className="form-input"
                            style={{ padding: '6px 10px', fontSize: '0.85rem', width: 100 }}
                            placeholder={isDe ? 'offen' : 'open'}
                            value={field.maxNights || ''}
                            onChange={e => updateCustomField(field.id, { maxNights: parseInt(e.target.value, 10) || 0 })} />
                        </div>
                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', flex: '1 1 220px', lineHeight: 1.45, paddingBottom: 6 }}>
                          {isDe
                            ? 'Der Teilnehmer wählt Anreise und Abreise; die Nächte werden angezeigt. „Ich brauche kein Hotel" ist immer mit dabei. Die Hotel-Planung übernimmt den Zeitraum direkt.'
                            : 'The attendee picks arrival and departure; the nights are shown. „I don’t need a hotel" is always offered. Hotel planning takes the period from here.'}
                        </div>
                      </div>
                    )}
                    {/* v24.25: Feldart-Empfehlung (Datum/Person) anhand des Labels. */}
                    <FieldTypeSuggestion
                      field={field}
                      isDe={isDe}
                      allowPerson
                      onApply={(t) => updateCustomField(field.id, { type: t })}
                    />

                    {isExpanded && (<>
                    {/* v26.60: Roommate-Feld → separate Zimmerpartner-Anfrage-Mail
                        an die ausgewählte Person abschaltbar (Default: an). Diese
                        Mail ist UNABHÄNGIG vom CC-Schalter darunter — das war der
                        gemeldete Bug: „CC ausgestellt, Mails kommen trotzdem". */}
                    {field.type === 'roommate' && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginTop: 8, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={field.notifyRoommate !== false}
                          onChange={e => updateCustomField(field.id, { notifyRoommate: e.target.checked ? undefined : false })}
                          style={{ marginTop: 2 }}
                        />
                        <span style={{ flex: 1 }}>
                          <span style={{ fontSize: '0.85rem', fontWeight: 600 }}>
                            {isDe ? 'Ausgewählte Person automatisch benachrichtigen (Zimmerpartner-Anfrage-Mail)' : 'Automatically notify the selected person (roommate request email)'}
                          </span>
                          <InfoTooltip text={isDe ? (
                            <>
                              <strong>Was du hier einstellst:</strong> ob die als Zimmerpartner <strong>ausgewählte Person</strong> direkt nach der Anmeldung eine eigene <strong>&bdquo;Zimmerpartner-Anfrage&ldquo;-Mail</strong> bekommt (&bdquo;X hat dich als Zimmerpartner angegeben&ldquo;).<br /><br />
                              <strong>Unabhängig vom CC-Schalter:</strong> Diese Benachrichtigung ist eine EIGENE Mail — der CC-Schalter darunter steuert nur die Kopie der An-/Abmelde-Mail.<br /><br />
                              <strong>Ausgeschaltet:</strong> die ausgewählte Person bekommt gar keine automatische Mail mehr; auch der Hinweis dazu im Anmeldeformular wird ausgeblendet.
                            </>
                          ) : (
                            <>
                              <strong>What this controls:</strong> whether the person <strong>selected as roommate</strong> receives a dedicated <strong>roommate request email</strong> right after registration.<br /><br />
                              <strong>Independent of the CC toggle:</strong> this notification is a SEPARATE email — the CC toggle below only controls the copy of the registration/cancellation email.<br /><br />
                              <strong>When off:</strong> the selected person receives no automatic email at all; the hint in the registration form is hidden too.
                            </>
                          )} />
                        </span>
                      </label>
                    )}
                    {/* v29.40: Personen-Feld → Suche auf den Verteilerkreis des
                        Events begrenzen. Anlass: Zimmerpartner liessen sich
                        ausserhalb des Verteilers auswaehlen. */}
                    {(field.type === 'user' || field.type === 'roommate') && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginTop: 8, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={!!field.audienceOnly}
                          onChange={e => updateCustomField(field.id, { audienceOnly: e.target.checked ? true : undefined })}
                          style={{ marginTop: 2 }}
                        />
                        <span style={{ flex: 1 }}>
                          <span style={{ fontSize: '0.85rem', fontWeight: 600 }}>
                            {isDe ? 'Nur Personen zulassen, die dieses Event sehen können' : 'Only allow people who can see this event'}
                          </span>
                          <InfoTooltip text={isDe ? (
                            <>
                              <strong>Was du hier einstellst:</strong> ob die Personensuche in diesem Feld <strong>nur Personen aus dem Verteilerkreis</strong> des Events findet (Mailverteiler, Standortfilter, abzüglich ausgeschlossener Personen) — dieselbe Prüfung wie beim Anmelden für andere.<br /><br />
                              <strong>Aus (Standard):</strong> alle Kolleg:innen sind wählbar, auch wenn sie gar nicht eingeladen sind.<br /><br />
                              <strong>Typischer Fall:</strong> Zimmerpartner. Ohne diese Option lässt sich jemand angeben, der beim Event gar nicht dabei ist.<br /><br />
                              <strong>Hinweis:</strong> Läuft die Sichtbarkeit deines Events nur über einen Standortfilter, greift die Prüfung genauso — nur über einen reinen &bdquo;alle&ldquo;-Kreis gibt es nichts einzugrenzen.
                            </>
                          ) : (
                            <>
                              <strong>What this controls:</strong> whether the people search in this field only finds people <strong>within the event audience</strong> (distribution lists, location filter, minus excluded people) — the same check as registering on behalf of someone.<br /><br />
                              <strong>Off (default):</strong> every colleague can be picked, even if they are not invited at all.<br /><br />
                              <strong>Typical case:</strong> roommates. Without this option someone can name a person who is not attending.
                            </>
                          )} />
                        </span>
                      </label>
                    )}
                    {/* v18.41: People-Picker-Feld → ausgewählte Person bei
                        An-/Abmelde-Mail auf CC (nur für user/roommate-Felder). */}
                    {(field.type === 'user' || field.type === 'roommate') && (
                      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginTop: 8, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={!!field.ccOnEmails}
                          onChange={e => updateCustomField(field.id, { ccOnEmails: e.target.checked })}
                          style={{ marginTop: 2 }}
                        />
                        <span style={{ flex: 1 }}>
                          <span style={{ fontSize: '0.85rem', fontWeight: 600 }}>
                            {isDe ? 'Ausgewählte Person bei An-/Abmelde-Mail auf CC setzen' : 'CC the selected person on registration/cancellation email'}
                          </span>
                          <InfoTooltip text={isDe ? (
                            <>
                              <strong>Was du hier einstellst:</strong> ob die in diesem Feld <strong>ausgewählte Person</strong> (z.&nbsp;B. die Assistenz) die <strong>Anmelde- und Abmelde-Mail</strong> des Teilnehmers <strong>in Kopie (CC)</strong> bekommt.<br /><br />
                              <strong>Anzeige in der App:</strong> ändert nichts an der Anzeige — wirkt nur beim Mail-Versand.<br /><br />
                              <strong>Automatismen:</strong> die im Feld gewählte Person wird automatisch auf CC der Bestätigungs- bzw. Abmelde-Mail gesetzt. <strong>Der Outlook-Termin ist davon nicht betroffen</strong> — die Person wird also NICHT in den Kalendereintrag eingeladen.<br /><br />
                              <strong>Auswirkung für Teilnehmer:</strong> seine Assistenz ist bei An- und Abmeldung automatisch informiert, ohne dass er sie manuell weiterleiten muss.
                            </>
                          ) : (
                            <>
                              <strong>What you set here:</strong> whether the <strong>person selected in this field</strong> (e.g. the assistant) receives the attendee&apos;s <strong>registration and cancellation email</strong> in <strong>CC</strong>.<br /><br />
                              <strong>Where you see it:</strong> no visible change — only affects email sending.<br /><br />
                              <strong>Automations:</strong> the chosen person is automatically added to CC of the confirmation / cancellation mail. <strong>The Outlook event is not affected</strong> — the person is NOT invited to the calendar entry.<br /><br />
                              <strong>For attendees:</strong> their assistant is automatically kept in the loop on registration and cancellation without manual forwarding.
                            </>
                          )} />
                          <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                            {isDe
                              ? 'Nur die Mails — der Outlook-Termin wird nicht an diese Person gesendet.'
                              : 'Emails only — the Outlook event is not sent to this person.'}
                          </span>
                        </span>
                      </label>
                    )}

                    {/* v17.20: EN-Feld-Name — sichtbar wenn der Bilingual-
                        Toggle oben aktiviert wurde. Sitzt direkt unter dem
                        DE-Feld-Namen, damit der Organizer beide Sprachen in
                        einer Linie liest. Flagge + Placeholder machen klar,
                        welche Sprache gemeint ist. */}
                    {bilingualFields && (
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        marginLeft: 64, marginBottom: 6,
                      }}>
                        <span style={{
                          flexShrink: 0, fontSize: '0.7rem',
                          padding: '2px 8px', borderRadius: 8,
                          background: 'rgba(0,90,156,0.10)',
                          color: '#005a9c', fontWeight: 700, letterSpacing: 0.5,
                        }}>EN</span>
                        <input
                          className="form-input"
                          value={field.labelEn || ''}
                          placeholder={isDe
                            ? 'Englischer Feld-Name (optional — leer = fällt auf den deutschen Text zurück)'
                            : 'English field name (optional — empty = falls back to the German text)'}
                          onChange={e => updateCustomField(field.id, { labelEn: e.target.value })}
                          style={{ flex: 1, fontSize: '0.88rem', padding: '6px 10px' }}
                        />
                      </div>
                    )}

                    {/* v10.24: Pro-Gruppe-Sichtbarkeit — nur sichtbar wenn die
                        Split-Capacity in Schritt 3 aktiv ist. Der Organizer
                        kann ein Feld auf Gruppe A oder Gruppe B beschränken
                        (Beispiel: Pflicht-Checkbox „Leistungsnachweis vorhanden"
                        nur für Durchstarter / Gruppe A). 'all' = beide
                        Gruppen sehen das Feld (Default). Bei Gruppe-A/B-only
                        wird das Feld in der RegistrationPage entsprechend nur
                        beim passenden Wunsch-Typ gerendert. */}
                    {useSplitCapacities && (() => {
                      const labelA = (splitLabelA || '').trim() || 'Durchstarter';
                      const labelB = (splitLabelB || '').trim() || 'Funstarter';
                      const current = field.onlyForGroup || 'all';
                      const pillStyle = (active: boolean): React.CSSProperties => ({
                        display: 'inline-flex', alignItems: 'center', gap: 6,
                        padding: '6px 12px', borderRadius: 999,
                        fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap',
                        cursor: 'pointer', userSelect: 'none',
                        border: `1px solid ${active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                        background: active ? 'rgba(134,188,37,0.10)' : '#fff',
                        color: active ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                        transition: 'all 0.15s ease',
                      });
                      return (
                        <div style={{
                          marginTop: 10, padding: '12px 14px',
                          background: '#fff',
                          border: '1px solid var(--dex-gray-200)',
                          borderRadius: 8,
                        }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8, flexWrap: 'wrap' }}>
                            <label style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
                              {isDe ? 'Sichtbar für Teilnehmergruppe' : 'Visible for attendee group'}
                            </label>
                            <InfoTooltip text={isDe ? (
                              <>
                                <strong>Was du hier einstellst:</strong> ob dieses Feld bei der Anmeldung für <strong>alle Teilnehmer</strong> oder nur für eine der zwei Kapazitäts-Gruppen sichtbar ist.<br /><br />
                                <strong>Beispiel:</strong> bei einem Lauf-Event ist die Pflicht-Checkbox &bdquo;Leistungsnachweis vorhanden&ldquo; nur für die <strong>Durchstarter-Gruppe</strong> sinnvoll, nicht für Funstarter / Walker. Stelle das Feld dann auf <strong>Nur {labelA}</strong> — Funstarter sehen es gar nicht erst.<br /><br />
                                <strong>Auswirkung in der App:</strong> die Anmelde-Seite blendet das Feld dynamisch ein/aus, sobald der Teilnehmer eine der zwei Boxen wählt. Pflichtfeld-Validierung greift natürlich nur wenn das Feld auch sichtbar ist.<br /><br />
                                <strong>Vorraussetzung:</strong> in Schritt 4 (Kapazität &amp; Sichtbarkeit) muss der Toggle &bdquo;Geteilte Kapazität&ldquo; aktiv sein. Sonst gibt&apos;s keine Gruppen — dieser Selector ist dann ausgeblendet.
                              </>
                            ) : (
                              <>
                                <strong>What you set here:</strong> whether this field is visible to <strong>all attendees</strong> or only to one of the two capacity groups during registration.<br /><br />
                                <strong>Example:</strong> on a running event, a required checkbox &ldquo;Performance proof available&rdquo; only makes sense for the <strong>fast-runner group</strong>, not for fun-runners / walkers. Set the field to <strong>{labelA} only</strong> — fun-runners won&apos;t even see it.<br /><br />
                                <strong>Effect in the app:</strong> the registration page dynamically shows / hides the field as the attendee picks one of the two boxes. Required-field validation only fires when the field is actually visible.<br /><br />
                                <strong>Requirement:</strong> the &ldquo;Split capacity&rdquo; toggle in step 4 (Capacity &amp; Visibility) must be active. Otherwise there are no groups — this selector is then hidden.
                              </>
                            )} />
                          </div>
                          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                            {([
                              { v: 'all', text: isDe ? 'Beide Gruppen' : 'Both groups' },
                              { v: 'A', text: isDe ? `Nur ${labelA}` : `${labelA} only` },
                              { v: 'B', text: isDe ? `Nur ${labelB}` : `${labelB} only` },
                            ] as const).map(opt => (
                              <label key={opt.v} style={pillStyle(current === opt.v)}>
                                <input
                                  type="radio"
                                  name={`onlyForGroup-${field.id}`}
                                  checked={current === opt.v}
                                  onChange={() => updateCustomField(field.id, { onlyForGroup: opt.v as 'all' | 'A' | 'B' })}
                                  style={{ display: 'none' }}
                                />
                                <span style={{ fontSize: '0.85rem', lineHeight: 1 }}>{current === opt.v ? '●' : '○'}</span>
                                {opt.text}
                              </label>
                            ))}
                          </div>
                        </div>
                      );
                    })()}

                    {/* v7.20: Beschreibung pro Feld. v18.18: Darstellung
                        wählbar — „i"-Box neben dem Label ODER Erklär-Text
                        unter dem Label. */}
                    <div style={{ marginLeft: 32, marginTop: 10 }}>
                      {/* v27.4: Kompakter Editor mit dauerhaft sichtbarer Leiste
                          (Fett + Link) statt der pnp-RichText-Bubble. */}
                      <FieldDescEditor
                        value={field.helpText || ''}
                        onChange={text => updateCustomField(field.id, { helpText: text })}
                        isDe={isDe}
                      />
                      {field.helpText && field.helpText.trim() && (
                        <div style={{ display: 'flex', gap: 16, marginTop: 6, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                          <span style={{ fontWeight: 600 }}>{isDe ? 'Anzeige:' : 'Display:'}</span>
                          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 5, cursor: 'pointer' }}>
                            <input
                              type="radio"
                              name={`helpStyle-${field.id}`}
                              checked={(field.helpTextStyle || 'tooltip') !== 'inline'}
                              onChange={() => updateCustomField(field.id, { helpTextStyle: 'tooltip' })}
                            />
                            {isDe ? '„i"-Info-Box (Hover)' : '„i" info box (hover)'}
                          </label>
                          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 5, cursor: 'pointer' }}>
                            <input
                              type="radio"
                              name={`helpStyle-${field.id}`}
                              checked={field.helpTextStyle === 'inline'}
                              onChange={() => updateCustomField(field.id, { helpTextStyle: 'inline' })}
                            />
                            {isDe ? 'Text unter dem Feld-Titel' : 'Text below the field title'}
                          </label>
                        </div>
                      )}
                    </div>
                    {/* v17.20: EN-Variante der Beschreibung. */}
                    {bilingualFields && (
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        marginLeft: 64, marginTop: 4,
                      }}>
                        <span style={{
                          flexShrink: 0, fontSize: '0.65rem',
                          padding: '1px 6px', borderRadius: 6,
                          background: 'rgba(0,90,156,0.10)',
                          color: '#005a9c', fontWeight: 700, letterSpacing: 0.5,
                        }}>EN</span>
                        <input
                          className="form-input"
                          value={field.helpTextEn || ''}
                          placeholder={isDe
                            ? 'Englische Beschreibung (optional)'
                            : 'English description (optional)'}
                          onChange={e => updateCustomField(field.id, { helpTextEn: e.target.value })}
                          style={{ flex: 1, fontSize: '0.78rem', padding: '5px 9px' }}
                        />
                      </div>
                    )}

                    {/* v11.94: Bei Checkbox-Feldern kann der Organizer den
                        Text neben der Checkbox individuell setzen — Default
                        ist „Ja, bestätigen" / „Yes, confirm". */}
                    {field.type === 'checkbox' && (
                      <div style={{ marginLeft: 32, marginTop: 6 }}>
                        <input
                          className="form-input"
                          placeholder={isDe
                            ? 'Text neben Checkbox (optional, Default: „Ja, bestätigen")'
                            : 'Text next to checkbox (optional, default: „Yes, confirm")'}
                          value={field.confirmLabel || ''}
                          onChange={e => updateCustomField(field.id, { confirmLabel: e.target.value })}
                          style={{ width: '100%', fontSize: '0.82rem', padding: '6px 10px' }}
                        />
                      </div>
                    )}
                    {/* v17.20: EN-Variante des Checkbox-Bestätigungstexts. */}
                    {field.type === 'checkbox' && bilingualFields && (
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 8,
                        marginLeft: 64, marginTop: 4,
                      }}>
                        <span style={{
                          flexShrink: 0, fontSize: '0.65rem',
                          padding: '1px 6px', borderRadius: 6,
                          background: 'rgba(0,90,156,0.10)',
                          color: '#005a9c', fontWeight: 700, letterSpacing: 0.5,
                        }}>EN</span>
                        <input
                          className="form-input"
                          value={field.confirmLabelEn || ''}
                          placeholder={isDe
                            ? 'Englischer Text neben Checkbox (optional, Default: „Yes, confirm")'
                            : 'English text next to checkbox (optional, default: „Yes, confirm")'}
                          onChange={e => updateCustomField(field.id, { confirmLabelEn: e.target.value })}
                          style={{ flex: 1, fontSize: '0.78rem', padding: '5px 9px' }}
                        />
                      </div>
                    )}


                    {/* v10.23: Dropdown-Optionen als gelisteter Editor mit
                        eigener Box. Pro Option eine Zeile mit Nummer +
                        Eingabefeld + Minus-Button. Plus-Button ans Ende
                        zum Hinzufügen. Mehrfachauswahl-Toggle direkt in
                        diesem Block (Kontext: betrifft nur die Optionsliste).
                        Nur sichtbar wenn type === 'select'. */}
                    {field.type === 'select' && (
                      <div style={{
                        marginTop: 10, padding: '12px 14px',
                        background: '#fff',
                        border: '1px solid var(--dex-gray-200)',
                        borderRadius: 8,
                      }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8, flexWrap: 'wrap', gap: 8 }}>
                          <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-700)', fontWeight: 600 }}>
                            {isDe ? 'Optionen' : 'Options'}
                          </div>
                          <label
                            style={{
                              display: 'inline-flex', alignItems: 'center', gap: 6,
                              padding: '4px 10px', borderRadius: 999,
                              fontSize: '0.75rem', fontWeight: 600, whiteSpace: 'nowrap',
                              cursor: 'pointer', userSelect: 'none',
                              border: `1px solid ${field.multi ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                              background: field.multi ? 'rgba(134,188,37,0.10)' : '#fff',
                              color: field.multi ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                              transition: 'all 0.15s ease',
                            }}
                            title={isDe
                              ? 'Wenn aktiv, kann der Teilnehmer mehrere Optionen gleichzeitig auswählen (z.B. mehrere Allergien).'
                              : 'When enabled, attendees can select multiple options at the same time (e.g. multiple allergies).'}
                          >
                            <input
                              type="checkbox"
                              checked={!!field.multi}
                              onChange={e => updateCustomField(field.id, { multi: e.target.checked, ...(e.target.checked ? { optionCategories: undefined, prefilterLabel: undefined, defaultValue: undefined } : {}) })}
                              style={{ display: 'none' }}
                            />
                            <span style={{ fontSize: '0.8rem', lineHeight: 1 }}>{field.multi ? '✓' : '○'}</span>
                            {isDe ? 'Mehrfachauswahl möglich' : 'Allow multiple selection'}
                          </label>
                          {/* v26.75: Vorfilter — nur bei Single-Select. Aktiviert
                              pro Option ein Kategorie-Feld; die Anmeldeseite zeigt
                              dann zuerst ein Kategorie-Dropdown und filtert die
                              Optionsliste darauf (z.B. „Herren"/„Damen" → Größen). */}
                          {!field.multi && (
                            <label
                              style={{
                                display: 'inline-flex', alignItems: 'center', gap: 6,
                                padding: '4px 10px', borderRadius: 999,
                                fontSize: '0.75rem', fontWeight: 600, whiteSpace: 'nowrap',
                                cursor: 'pointer', userSelect: 'none',
                                border: `1px solid ${field.optionCategories ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                background: field.optionCategories ? 'rgba(134,188,37,0.10)' : '#fff',
                                color: field.optionCategories ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                              }}
                              title={isDe
                                ? 'Zu jeder Option eine Kategorie hinterlegen. Der Teilnehmer wählt erst die Kategorie, dann sieht er nur die passenden Optionen (kürzere Liste).'
                                : 'Give each option a category. The attendee picks the category first and then only sees the matching options (shorter list).'}
                            >
                              <input
                                type="checkbox"
                                checked={!!field.optionCategories}
                                onChange={e => updateCustomField(field.id, e.target.checked
                                  ? { optionCategories: (field.options || []).map(() => '') }
                                  : { optionCategories: undefined, prefilterLabel: undefined })}
                                style={{ display: 'none' }}
                              />
                              <span style={{ fontSize: '0.8rem', lineHeight: 1 }}>{field.optionCategories ? '✓' : '○'}</span>
                              {isDe ? 'Vorfilter nach Kategorie' : 'Pre-filter by category'}
                            </label>
                          )}
                        </div>
                        {/* v26.75: Beschriftung des Vorfilter-Dropdowns. */}
                        {!field.multi && field.optionCategories && (
                          <div style={{ marginBottom: 10 }}>
                            <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-600)', fontWeight: 600, marginBottom: 3 }}>
                              {isDe ? 'Bezeichnung des Vorfilters' : 'Pre-filter label'}
                            </label>
                            <input
                              className="form-input"
                              value={field.prefilterLabel || ''}
                              onChange={e => updateCustomField(field.id, { prefilterLabel: e.target.value })}
                              placeholder={isDe ? 'z.B. Größentabelle, Kategorie' : 'e.g. size chart, category'}
                              maxLength={40}
                              style={{ fontSize: '0.85rem', padding: '6px 10px' }}
                            />
                          </div>
                        )}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                          {/* v26.92: Bei aktivem Vorfilter werden die Optionen
                              GRUPPIERT nach Kategorie bearbeitet — Kategorie oben,
                              darunter die Auswahl, plus eine „Ohne Kategorie"-Gruppe
                              (immer sichtbar). Ohne Vorfilter bleibt die flache
                              Liste. Das Datenmodell (options/optionsEn/optionCategories
                              als Parallel-Arrays) bleibt unverändert. */}
                          {field.optionCategories ? (() => {
                            const opts = field.options || [];
                            const optsEn = field.optionsEn || [];
                            const cats = field.optionCategories || [];
                            const named: Array<{ category: string; items: Array<{ opt: string; optEn: string }> }> = [];
                            const noCat: Array<{ opt: string; optEn: string }> = [];
                            const idxByCat = new Map<string, number>();
                            opts.forEach((o, i) => {
                              const c = (cats[i] || '').trim();
                              const item = { opt: o, optEn: optsEn[i] || '' };
                              if (!c) { noCat.push(item); return; }
                              if (!idxByCat.has(c)) { idxByCat.set(c, named.length); named.push({ category: c, items: [] }); }
                              named[idxByCat.get(c) as number].items.push(item);
                            });
                            const apply = (nm: typeof named, nc: typeof noCat): void => {
                              const nOpts: string[] = []; const nEn: string[] = []; const nCats: string[] = [];
                              nm.forEach(g => g.items.forEach(it => { nOpts.push(it.opt); nEn.push(it.optEn); nCats.push(g.category); }));
                              nc.forEach(it => { nOpts.push(it.opt); nEn.push(it.optEn); nCats.push(''); });
                              updateCustomField(field.id, { options: nOpts, optionsEn: nEn, optionCategories: nCats });
                            };
                            const catBadge = <span style={{ flexShrink: 0, fontSize: '0.65rem', padding: '1px 6px', borderRadius: 6, background: 'rgba(134,188,37,0.14)', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, letterSpacing: 0.3 }}>{isDe ? 'KAT' : 'CAT'}</span>;
                            const removeBtnStyle: React.CSSProperties = { flexShrink: 0, width: 26, height: 26, borderRadius: 6, background: '#fff', border: '1px solid var(--dex-gray-300)', color: 'var(--dex-red, #c00)', cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '1rem', lineHeight: 1, fontWeight: 700 };
                            const addBtnStyle: React.CSSProperties = { alignSelf: 'flex-start', marginTop: 2, display: 'inline-flex', alignItems: 'center', gap: 5, background: 'rgba(134,188,37,0.08)', border: '1px dashed var(--dex-green, #86bc25)', color: 'var(--dex-green-dark, #4a7c1f)', borderRadius: 6, padding: '5px 10px', fontSize: '0.76rem', fontWeight: 600, cursor: 'pointer' };
                            const optionRows = (items: Array<{ opt: string; optEn: string }>, onOpt: (ii: number, v: string) => void, onOptEn: (ii: number, v: string) => void, onRemove: (ii: number) => void): React.ReactNode => (
                              items.map((it, ii) => (
                                <div key={ii} style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                    <span style={{ flexShrink: 0, fontSize: '0.72rem', color: 'var(--dex-gray-400)', width: 16, textAlign: 'right' }}>{ii + 1}.</span>
                                    <input className="form-input" value={it.opt} placeholder={isDe ? 'Auswahl (z.B. S)' : 'Choice (e.g. S)'} onChange={e => onOpt(ii, e.target.value)} style={{ flex: 1, fontSize: '0.85rem', padding: '6px 10px' }} />
                                    <button type="button" onClick={() => onRemove(ii)} title={isDe ? 'Entfernen' : 'Remove'} style={removeBtnStyle}>−</button>
                                  </div>
                                  {bilingualFields && (
                                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginLeft: 24 }}>
                                      <span style={{ flexShrink: 0, fontSize: '0.65rem', padding: '1px 6px', borderRadius: 6, background: 'rgba(0,90,156,0.10)', color: '#005a9c', fontWeight: 700, letterSpacing: 0.5 }}>EN</span>
                                      <input className="form-input" value={it.optEn} placeholder={isDe ? 'Englische Variante (optional)' : 'English variant (optional)'} onChange={e => onOptEn(ii, e.target.value)} style={{ flex: 1, fontSize: '0.78rem', padding: '5px 9px' }} />
                                    </div>
                                  )}
                                </div>
                              ))
                            );
                            return (
                              <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                                {named.map((g, gi) => (
                                  <div key={gi} style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: '10px 12px', background: 'rgba(134,188,37,0.04)', display: 'flex', flexDirection: 'column', gap: 6 }}>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                      {catBadge}
                                      <input className="form-input" value={g.category} placeholder={isDe ? 'Kategorie (z.B. Männergrößen)' : 'Category (e.g. men’s sizes)'} onChange={e => apply(named.map((x, i) => i === gi ? { ...x, category: e.target.value } : x), noCat)} style={{ flex: 1, fontSize: '0.85rem', fontWeight: 600, padding: '6px 10px' }} />
                                      <button type="button" onClick={() => apply(named.filter((_, i) => i !== gi), noCat)} title={isDe ? 'Kategorie entfernen' : 'Remove category'} style={removeBtnStyle}>−</button>
                                    </div>
                                    {optionRows(
                                      g.items,
                                      (ii, v) => apply(named.map((x, i) => i === gi ? { ...x, items: x.items.map((y, j) => j === ii ? { ...y, opt: v } : y) } : x), noCat),
                                      (ii, v) => apply(named.map((x, i) => i === gi ? { ...x, items: x.items.map((y, j) => j === ii ? { ...y, optEn: v } : y) } : x), noCat),
                                      (ii) => apply(named.map((x, i) => i === gi ? { ...x, items: x.items.filter((_, j) => j !== ii) } : x), noCat),
                                    )}
                                    <button type="button" onClick={() => apply(named.map((x, i) => i === gi ? { ...x, items: [...x.items, { opt: '', optEn: '' }] } : x), noCat)} style={addBtnStyle}>+ {isDe ? 'Auswahl hinzufügen' : 'Add choice'}</button>
                                  </div>
                                ))}
                                <div style={{ border: '1px dashed var(--dex-gray-300)', borderRadius: 8, padding: '10px 12px', display: 'flex', flexDirection: 'column', gap: 6 }}>
                                  <div style={{ fontSize: '0.75rem', fontWeight: 600, color: 'var(--dex-gray-600)' }}>
                                    {isDe ? 'Ohne Kategorie — immer sichtbar (z.B. „T-Shirt bereits vorhanden")' : 'No category — always shown (e.g. „already have a shirt")'}
                                  </div>
                                  {optionRows(
                                    noCat,
                                    (ii, v) => apply(named, noCat.map((y, j) => j === ii ? { ...y, opt: v } : y)),
                                    (ii, v) => apply(named, noCat.map((y, j) => j === ii ? { ...y, optEn: v } : y)),
                                    (ii) => apply(named, noCat.filter((_, j) => j !== ii)),
                                  )}
                                  <button type="button" onClick={() => apply(named, [...noCat, { opt: '', optEn: '' }])} style={addBtnStyle}>+ {isDe ? 'Option ohne Kategorie' : 'Option without category'}</button>
                                </div>
                                <button type="button" onClick={() => apply([...named, { category: isDe ? `Kategorie ${named.length + 1}` : `Category ${named.length + 1}`, items: [{ opt: '', optEn: '' }] }], noCat)} style={{ ...addBtnStyle, marginTop: 0, borderStyle: 'solid', background: 'var(--dex-green, #86bc25)', color: '#fff', border: 'none', padding: '7px 14px', fontSize: '0.82rem' }}>+ {isDe ? 'Kategorie hinzufügen' : 'Add category'}</button>
                              </div>
                            );
                          })() : <>
                          {(field.options || []).map((opt, optIdx) => (
                            <div key={optIdx} style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                <span style={{ flexShrink: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontWeight: 600, width: 24, textAlign: 'right' }}>
                                  {optIdx + 1}.
                                </span>
                                <input
                                  className="form-input"
                                  value={opt}
                                  placeholder={isDe ? `Option ${optIdx + 1}` : `Option ${optIdx + 1}`}
                                  onChange={e => {
                                    const opts = [...(field.options || [])];
                                    opts[optIdx] = e.target.value;
                                    updateCustomField(field.id, { options: opts });
                                  }}
                                  style={{ flex: 1, fontSize: '0.85rem', padding: '6px 10px' }}
                                />
                                <button
                                  type="button"
                                  onClick={() => {
                                    const opts = [...(field.options || [])];
                                    opts.splice(optIdx, 1);
                                    // v17.20: EN-Optionsliste positional mit-zurücksetzen,
                                    // damit Index-Mapping konsistent bleibt.
                                    const optsEn = [...(field.optionsEn || [])];
                                    if (optsEn.length > optIdx) optsEn.splice(optIdx, 1);
                                    // v26.75: Vorfilter-Kategorien positional mit-splicen.
                                    const upd: Partial<CustomFieldInput> = { options: opts, optionsEn: optsEn };
                                    if (field.optionCategories) {
                                      const cats = [...field.optionCategories];
                                      if (cats.length > optIdx) cats.splice(optIdx, 1);
                                      upd.optionCategories = cats;
                                    }
                                    updateCustomField(field.id, upd);
                                  }}
                                  title={isDe ? 'Option entfernen' : 'Remove option'}
                                  style={{
                                    flexShrink: 0, width: 28, height: 28, borderRadius: 6,
                                    background: '#fff', border: '1px solid var(--dex-gray-300)',
                                    color: 'var(--dex-red, #c00)', cursor: 'pointer',
                                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                    fontSize: '1rem', lineHeight: 1, fontWeight: 700,
                                  }}
                                >−</button>
                              </div>
                              {/* v17.20: Positional gemappte EN-Option. */}
                              {bilingualFields && (
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginLeft: 32 }}>
                                  <span style={{
                                    flexShrink: 0, fontSize: '0.65rem',
                                    padding: '1px 6px', borderRadius: 6,
                                    background: 'rgba(0,90,156,0.10)',
                                    color: '#005a9c', fontWeight: 700, letterSpacing: 0.5,
                                  }}>EN</span>
                                  <input
                                    className="form-input"
                                    value={(field.optionsEn || [])[optIdx] || ''}
                                    placeholder={isDe ? 'Englische Variante (optional)' : 'English variant (optional)'}
                                    onChange={e => {
                                      const optsEn = [...(field.optionsEn || [])];
                                      while (optsEn.length <= optIdx) optsEn.push('');
                                      optsEn[optIdx] = e.target.value;
                                      updateCustomField(field.id, { optionsEn: optsEn });
                                    }}
                                    style={{ flex: 1, fontSize: '0.78rem', padding: '5px 9px' }}
                                  />
                                </div>
                              )}
                              {/* v26.75: Kategorie pro Option (nur wenn Vorfilter aktiv). */}
                              {field.optionCategories && (
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginLeft: 32 }}>
                                  <span style={{
                                    flexShrink: 0, fontSize: '0.65rem',
                                    padding: '1px 6px', borderRadius: 6,
                                    background: 'rgba(134,188,37,0.14)',
                                    color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, letterSpacing: 0.3,
                                  }}>{isDe ? 'KAT' : 'CAT'}</span>
                                  <input
                                    className="form-input"
                                    value={(field.optionCategories || [])[optIdx] || ''}
                                    placeholder={isDe ? 'Kategorie (z.B. Herren) — leer = immer sichtbar' : 'Category (e.g. men) — empty = always shown'}
                                    onChange={e => {
                                      const cats = [...(field.optionCategories || [])];
                                      while (cats.length <= optIdx) cats.push('');
                                      cats[optIdx] = e.target.value;
                                      updateCustomField(field.id, { optionCategories: cats });
                                    }}
                                    style={{ flex: 1, fontSize: '0.78rem', padding: '5px 9px' }}
                                  />
                                </div>
                              )}
                            </div>
                          ))}
                          <button
                            type="button"
                            onClick={() => updateCustomField(field.id, { options: [...(field.options || []), ''], ...(field.optionCategories ? { optionCategories: [...field.optionCategories, ''] } : {}) })}
                            style={{
                              alignSelf: 'flex-start', marginTop: 4,
                              display: 'inline-flex', alignItems: 'center', gap: 6,
                              background: 'rgba(134,188,37,0.08)',
                              border: '1px dashed var(--dex-green, #86bc25)',
                              color: 'var(--dex-green-dark, #4a7c1f)',
                              borderRadius: 6, padding: '6px 12px',
                              fontSize: '0.8rem', fontWeight: 600, cursor: 'pointer',
                            }}
                          >
                            <span style={{ fontSize: '1rem', lineHeight: 1, fontWeight: 700 }}>+</span>
                            {isDe ? 'Option hinzufügen' : 'Add option'}
                          </button>
                          </>}
                          {/* v26.74: Vorauswahl (nur Single-Select) — optional
                              eine Option, die im Anmeldeformular vorausgewählt ist. */}
                          {!field.multi && (field.options || []).filter(o => (o || '').trim()).length > 0 && (
                            <div style={{ marginTop: 12, paddingTop: 10, borderTop: '1px solid var(--dex-gray-100)' }}>
                              <label style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-700)', fontWeight: 600, marginBottom: 4 }}>
                                {isDe ? 'Vorauswahl (optional)' : 'Pre-selected option (optional)'}
                              </label>
                              <select
                                className="form-input"
                                value={field.defaultValue || ''}
                                onChange={e => updateCustomField(field.id, { defaultValue: e.target.value })}
                                style={{ fontSize: '0.85rem', padding: '6px 10px' }}
                              >
                                <option value="">{isDe ? '— Keine Vorauswahl („Bitte wählen") —' : '— No pre-selection („Please choose") —'}</option>
                                {(field.options || []).map(o => (o || '').trim()).filter(Boolean).map((o, i) => (
                                  <option key={i} value={o}>{o}</option>
                                ))}
                              </select>
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                                {isDe ? 'Leer = Teilnehmer muss selbst wählen.' : 'Empty = attendee must choose themselves.'}
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    )}
                    {/* v10.23: Roommate-Erklärungs-Box — wird IMMER bei
                        type=roommate angezeigt (nicht mehr nur konditional bei
                        fehlendem Zimmerart-Feld). Erklärt dem Organizer, was
                        das Feld bei der Anmeldung tatsächlich tut: Personen-
                        Picker mit Mailbenachrichtigung an den ausgewählten
                        Roommate, Match-Erkennung im Admin Center wenn beide
                        sich gegenseitig wählen. Plus: Tipp wenn noch kein
                        Zimmerart-Feld da ist. */}
                    {field.type === 'roommate' && (() => {
                      const roomKeywords = ['einzelzimmer', 'doppelzimmer', 'single room', 'double room', 'zimmerart', 'room type'];
                      const hasRoomTypeField = customFields.some(other => {
                        if (other.id === field.id) return false;
                        const lbl = (other.label || '').toLowerCase();
                        const opts = (other.options || []).join(' ').toLowerCase();
                        return roomKeywords.some(k => lbl.indexOf(k) >= 0 || opts.indexOf(k) >= 0);
                      });
                      return (
                        <div style={{
                          marginTop: 10, padding: '12px 14px',
                          background: 'rgba(21,101,192,0.06)',
                          border: '1px solid rgba(21,101,192,0.4)',
                          borderRadius: 8, fontSize: '0.82rem', color: 'var(--dex-gray-700)',
                          lineHeight: 1.5,
                        }}>
                          <div style={{ fontWeight: 700, color: 'var(--dex-blue, #1565c0)', marginBottom: 6, display: 'flex', alignItems: 'center', gap: 6 }}>
                            <Icon iconName="Info" style={{ fontSize: 16 }} />
                            {isDe ? 'So funktioniert das Roommate-Feld' : 'How the roommate field works'}
                          </div>
                          {isDe ? (
                            <>
                              <p style={{ margin: '0 0 6px' }}>
                                Bei der Anmeldung sieht der Teilnehmer einen <strong>Personen-Suchfeld</strong> (Live-Suche im Deloitte-Tenant). Er tippt einen Namen ein, wählt die gewünschte Person als Zimmerpartner und schließt die Anmeldung ab.
                              </p>
                              <p style={{ margin: '0 0 6px' }}>
                                <strong>Was direkt passiert:</strong> die ausgewählte Person bekommt automatisch eine <strong>Roommate-Anfrage-Mail</strong> im Deloitte-Layout — mit Hinweis, dass <em>X den Wunsch geäußert hat, mit ihr/ihm das Zimmer zu teilen</em> und dem Link zur Event-Anmeldung. Der Empfänger kann sich dann selbst anmelden und seinerseits den Anfragenden als Wunsch-Roommate eintragen.
                              </p>
                              <p style={{ margin: '0 0 6px' }}>
                                <strong>Match-Erkennung im Admin Center:</strong> wenn zwei Teilnehmer sich <em>gegenseitig</em> als Wunsch-Roommate eingetragen haben, markiert das Admin Center das Paar als bestätigtes Match (grünes Häkchen). Einseitige Wünsche werden grau angezeigt — der Organizer kann dann selbst entscheiden, ob er die Person trotzdem zuteilt.
                              </p>
                              <p style={{ margin: 0 }}>
                                Sinnvoll <strong>kombiniert mit einem zweiten Feld vom Typ Dropdown</strong> namens &bdquo;Zimmerart&ldquo; mit Optionen &bdquo;Einzelzimmer / Doppelzimmer&ldquo;. Über die <strong>Sichtbarkeitsbedingung</strong> kannst du das Roommate-Feld dann ausblenden, wenn jemand &bdquo;Einzelzimmer&ldquo; wählt — sonst geben auch Einzelzimmer-Bucher einen Roommate an.
                              </p>
                              {!hasRoomTypeField && (
                                <p style={{ margin: '8px 0 0', padding: '8px 10px', background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 6, color: 'var(--dex-gray-700)' }}>
                                  <strong style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>Tipp:</strong> aktuell hast du noch kein Zimmerart-Feld angelegt. Ein Dropdown &bdquo;Zimmerart&ldquo; mit &bdquo;Einzelzimmer / Doppelzimmer&ldquo; wäre eine sinnvolle Ergänzung — sonst können auch Teilnehmer ohne Doppelzimmer-Wunsch einen Roommate angeben.
                                </p>
                              )}
                            </>
                          ) : (
                            <>
                              <p style={{ margin: '0 0 6px' }}>
                                On the registration page the attendee sees a <strong>person picker</strong> (live search of the Deloitte tenant). They type a name, pick their preferred roommate and submit the registration.
                              </p>
                              <p style={{ margin: '0 0 6px' }}>
                                <strong>What happens immediately:</strong> the selected person automatically receives a <strong>roommate-request email</strong> in the Deloitte layout — letting them know that <em>X requested to share a room with them</em> and including a link to the event registration. The recipient can then register and pick the requester back as their preferred roommate.
                              </p>
                              <p style={{ margin: '0 0 6px' }}>
                                <strong>Match detection in the admin center:</strong> when two attendees pick <em>each other</em> as preferred roommate, the admin center marks the pair as a confirmed match (green check). One-sided wishes are shown in grey — the organizer can still assign them manually if desired.
                              </p>
                              <p style={{ margin: 0 }}>
                                Best <strong>combined with a separate Dropdown field</strong> called &ldquo;Room type&rdquo; with options &ldquo;Single / Double&rdquo;. Use the <strong>visibility condition</strong> to hide the roommate field when someone picks &ldquo;Single&rdquo; — otherwise even single-room bookers will be asked to name a roommate.
                              </p>
                              {!hasRoomTypeField && (
                                <p style={{ margin: '8px 0 0', padding: '8px 10px', background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)', borderRadius: 6, color: 'var(--dex-gray-700)' }}>
                                  <strong style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>Tip:</strong> you don&apos;t have a room-type field yet. A dropdown &ldquo;Room type&rdquo; with &ldquo;Single / Double&rdquo; would be a useful addition — otherwise attendees without a double-room wish can still pick a roommate.
                                </p>
                              )}
                            </>
                          )}
                        </div>
                      );
                    })()}
                    {/* v7.21: Sichtbarkeitsbedingung — Feld nur anzeigen wenn
                        eine andere Frage einen bestimmten Wert hat. Quelle
                        kann nur ein Feld VOR diesem hier sein (idx < aktuell)
                        und muss vom Typ select oder checkbox sein. */}
                    {renderShowIfConfig(field, idx, customFields, (u) => updateCustomField(field.id, u))}
                    </>)}
                  </div>
                  );
                })}
              </div>

              {/* === Bereich 2: Felder pro Sub-Event (v10.11+) ============
                  Jedes Sub-Event hat eine eigene Custom-Fields-Liste, die bei
                  der Anmeldung auf einem Sub-Event zusätzlich zu den Hauptevent-
                  Feldern gerendert wird.
                  v15.0: dieser Bereich wird per pro-Sub-Event-Tab oben
                  ersetzt — Block bleibt im Code (mit display:none), um die
                  Inline-Helper-Funktionen `addSubEventCustomField` /
                  `updateSubEventCustomField` etc. nicht entfernen zu müssen
                  und sicherzustellen, dass der Tab-N>0-Code identische
                  Verhaltens-Garantien hat. */}
              <div style={{ display: 'none', marginTop: 32, paddingTop: 24, borderTop: '2px solid var(--dex-gray-200)' }}>
                <h3 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.05rem', fontWeight: 700 }}>
                  {isDe ? 'Felder pro Sub-Event' : 'Fields per sub-event'}
                </h3>
                <p style={{ margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                  {isDe
                    ? 'Optional: pro Sub-Event eine eigene Auswahl-Frage (z.B. „Welche Strecke läufst du?" mit Optionen 5/10/Halbmarathon). Diese Felder erscheinen NUR wenn der Teilnehmer das jeweilige Sub-Event wählt — zusätzlich zu den Feldern des Hauptevents oben.'
                    : 'Optional: a per-sub-event question (e.g. „Which distance?" with options 5/10/Halfmarathon). These fields appear ONLY when an attendee picks that sub-event — in addition to the main-event fields above.'}
                </p>

                {subEvents.length === 0 ? (
                  <WizardHint
                    isDe={isDe}
                    title={isDe ? 'Noch keine Sub-Events angelegt' : 'No sub-events yet'}
                  >
                    {isDe
                      ? 'Sub-Events legst du in Schritt 1 (Grundlagen) an — danach kannst du hier pro Sub-Event eigene Anmelde-Felder definieren.'
                      : 'Add sub-events in Step 3 (Sub-events) — then come back here to define per-sub-event registration fields.'}
                  </WizardHint>
                ) : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
                    {subEvents.map(se => {
                      const seFields = se.customFields || [];
                      return (
                        <div
                          key={se.id}
                          style={{
                            border: '1px solid var(--dex-gray-200)',
                            borderRadius: 8,
                            padding: '14px 16px',
                            background: '#fafafa',
                          }}
                        >
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 8, marginBottom: 10 }}>
                            <strong style={{ fontSize: '0.95rem' }}>
                              {se.title || (isDe ? '(unbenanntes Sub-Event)' : '(unnamed sub-event)')}
                            </strong>
                            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                              <button
                                type="button"
                                className="btn btn-outline"
                                style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                                onClick={() => addSubEventCustomField(se.id)}
                              >
                                <Plus size={12} /> {isDe ? 'Feld hinzufügen' : 'Add field'}
                              </button>
                              {customFields.length > 0 && seFields.length === 0 && (
                                <button
                                  type="button"
                                  className="btn btn-outline"
                                  style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                                  onClick={() => copyParentFieldsToSubEvent(se.id)}
                                  title={isDe ? `Dupliziert die ${customFields.length} Felder vom Hauptevent als Startpunkt` : 'Duplicates the main-event fields as a starting point'}
                                >
                                  {isDe ? `Vom Hauptevent kopieren (${customFields.length})` : `Copy from main event (${customFields.length})`}
                                </button>
                              )}
                            </div>
                          </div>

                          {seFields.length === 0 ? (
                            <p style={{ margin: 0, fontSize: '0.8rem', color: 'var(--dex-gray-500)', fontStyle: 'italic' }}>
                              {isDe
                                ? 'Keine zusätzlichen Felder — Teilnehmer wählen nur das Sub-Event ohne weitere Frage.'
                                : 'No additional fields — attendees just pick this sub-event with no further question.'}
                            </p>
                          ) : (
                            // v11.96: Sub-Event-Felder-Layout = exakt gleicher
                            // Look wie die Hauptevent-Felder (Step 5 oben):
                            // numerierte grüne Badge + prominentes Label-Input
                            // + grüner Typ-Selector + Pflicht-Pill + Lösch-X +
                            // gleiches Helptext / Confirm-Label / Optionen-
                            // Layout. Vorher kompakte Mini-Variante mit kleinen
                            // Schriften — visuell inkonsistent zum Hauptevent.
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                              {seFields.map((field, idx) => (
                                <div
                                  key={field.id}
                                  style={{
                                    background: 'var(--dex-gray-50, #fafafa)',
                                    borderRadius: 12,
                                    padding: 16,
                                    border: '1px solid var(--dex-gray-200)',
                                  }}
                                >
                                  {/* Header: Badge + Label + Typ + Pflicht + X */}
                                  <div style={{ display: 'flex', gap: 10, alignItems: 'center', marginBottom: 12, flexWrap: 'wrap' }}>
                                    <span style={{
                                      flexShrink: 0, width: 26, height: 26, borderRadius: '50%',
                                      background: 'var(--dex-green, #86bc25)', color: '#fff',
                                      display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                      fontWeight: 700, fontSize: '0.78rem',
                                    }}>{idx + 1}</span>
                                    <input
                                      className="form-input"
                                      value={field.label}
                                      placeholder={isDe ? 'Frage / Feld-Label (z.B. „Welche Strecke?")' : 'Question / field label (e.g. „Which distance?")'}
                                      onChange={e => updateSubEventCustomField(se.id, field.id, { label: e.target.value })}
                                      style={{
                                        flex: '0 1 320px', minWidth: 180, maxWidth: 320,
                                        fontSize: '1rem', fontWeight: 600,
                                        padding: '8px 12px',
                                        color: field.label ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
                                      }}
                                    />
                                    <select
                                      className="form-select"
                                      value={field.type}
                                      onChange={e => updateSubEventCustomField(se.id, field.id, { type: e.target.value as CustomFieldInput['type'] })}
                                      title={isDe ? 'Art des Feldes' : 'Field type'}
                                      style={{
                                        flex: '0 0 200px', maxWidth: 200,
                                        background: 'rgba(134,188,37,0.08)',
                                        border: '1px solid var(--dex-green, #86bc25)',
                                        color: 'var(--dex-green-dark, #4a7c1f)',
                                        fontWeight: 600,
                                        padding: '8px 10px',
                                      }}
                                    >
                                      <option value="text">{isDe ? 'Text (Freitext)' : 'Text (free text)'}</option>
                                      <option value="select">{isDe ? 'Dropdown' : 'Dropdown'}</option>
                                      <option value="number">{isDe ? 'Zahl' : 'Number'}</option>
                                      <option value="checkbox">{isDe ? 'Checkbox' : 'Checkbox'}</option>
                                      <option value="date">{isDe ? 'Datum (Kalender)' : 'Date (calendar)'}</option>
                        <option value="daterange">{isDe ? 'Übernachtungs-Zeitraum (Kalender + Nächte)' : 'Stay period (calendar + nights)'}</option>
                                    </select>
                                    <label
                                      style={{
                                        display: 'inline-flex', alignItems: 'center', gap: 6,
                                        padding: '6px 12px', borderRadius: 999,
                                        fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap',
                                        cursor: 'pointer', userSelect: 'none',
                                        border: `1px solid ${field.required ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                        background: field.required ? 'rgba(134,188,37,0.10)' : '#fff',
                                        color: field.required ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                                        transition: 'all 0.15s ease',
                                      }}
                                    >
                                      <input
                                        type="checkbox"
                                        checked={field.required}
                                        onChange={e => updateSubEventCustomField(se.id, field.id, { required: e.target.checked })}
                                        style={{ display: 'none' }}
                                      />
                                      <span style={{ fontSize: '0.85rem', lineHeight: 1 }}>{field.required ? '✓' : '○'}</span>
                                      {t('create.required')}
                                    </label>
                                    {field.type === 'select' && (
                                      <label
                                        style={{
                                          display: 'inline-flex', alignItems: 'center', gap: 6,
                                          padding: '6px 12px', borderRadius: 999,
                                          fontSize: '0.78rem', fontWeight: 600, whiteSpace: 'nowrap',
                                          cursor: 'pointer', userSelect: 'none',
                                          border: `1px solid ${field.multi ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                          background: field.multi ? 'rgba(134,188,37,0.10)' : '#fff',
                                          color: field.multi ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                                          transition: 'all 0.15s ease',
                                        }}
                                      >
                                        <input
                                          type="checkbox"
                                          checked={!!field.multi}
                                          onChange={e => updateSubEventCustomField(se.id, field.id, { multi: e.target.checked })}
                                          style={{ display: 'none' }}
                                        />
                                        <span style={{ fontSize: '0.85rem', lineHeight: 1 }}>{field.multi ? '✓' : '○'}</span>
                                        {isDe ? 'Mehrfachauswahl' : 'Multi-select'}
                                      </label>
                                    )}
                                    <button
                                      type="button"
                                      onClick={() => removeSubEventCustomField(se.id, field.id)}
                                      title={isDe ? 'Feld entfernen' : 'Remove field'}
                                      style={{ background: 'none', border: 'none', color: 'var(--dex-red)', padding: 4, cursor: 'pointer', flexShrink: 0 }}
                                    >
                                      <X size={18} />
                                    </button>
                                  </div>

                                  {/* v24.25: Datum-Feld → optional Uhrzeit mit abfragen. */}
                                  {field.type === 'date' && (
                                    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, marginLeft: 32, marginTop: 8, cursor: 'pointer', fontSize: '0.82rem', color: 'var(--dex-gray-700)' }}>
                                      <input
                                        type="checkbox"
                                        checked={!!field.withTime}
                                        onChange={e => updateSubEventCustomField(se.id, field.id, { withTime: e.target.checked })}
                                        style={{ accentColor: 'var(--dex-green, #86bc25)' }}
                                      />
                                      {isDe ? 'Auch Uhrzeit abfragen?' : 'Also ask for the time?'}
                                    </label>
                                  )}
                                  {field.type === 'daterange' && (
                                    <div style={{ marginLeft: 32, marginTop: 8, display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'flex-end' }}>
                                      <div>
                                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>{isDe ? 'Buchbar ab' : 'Bookable from'}</div>
                                        <input type="date" className="form-input" style={{ padding: '6px 10px', fontSize: '0.85rem', width: 160 }}
                                          value={field.rangeStart || ''} onChange={e => updateSubEventCustomField(se.id, field.id, { rangeStart: e.target.value })} />
                                      </div>
                                      <div>
                                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>{isDe ? 'Buchbar bis' : 'Bookable until'}</div>
                                        <input type="date" className="form-input" style={{ padding: '6px 10px', fontSize: '0.85rem', width: 160 }}
                                          value={field.rangeEnd || ''} onChange={e => updateSubEventCustomField(se.id, field.id, { rangeEnd: e.target.value })} />
                                      </div>
                                      <div>
                                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 3 }}>{isDe ? 'Max. Nächte' : 'Max. nights'}</div>
                                        <input type="number" min={0} className="form-input" style={{ padding: '6px 10px', fontSize: '0.85rem', width: 100 }}
                                          placeholder={isDe ? 'offen' : 'open'}
                                          value={field.maxNights || ''} onChange={e => updateSubEventCustomField(se.id, field.id, { maxNights: parseInt(e.target.value, 10) || 0 })} />
                                      </div>
                                    </div>
                                  )}
                                  {/* v24.25: Feldart-Empfehlung (nur Datum). */}
                                  <FieldTypeSuggestion
                                    field={field}
                                    isDe={isDe}
                                    allowPerson={false}
                                    onApply={(t) => updateSubEventCustomField(se.id, field.id, { type: t })}
                                  />

                                  {/* Beschreibung pro Feld */}
                                  <div style={{ marginLeft: 32, marginTop: 10 }}>
                                    <input
                                      className="form-input"
                                      placeholder={isDe
                                        ? 'Beschreibung (optional, erscheint als „i"-Tooltip neben dem Feld)'
                                        : 'Description (optional, shown as „i" tooltip next to the field)'}
                                      value={field.helpText || ''}
                                      onChange={e => updateSubEventCustomField(se.id, field.id, { helpText: e.target.value })}
                                      style={{ width: '100%', fontSize: '0.82rem', padding: '6px 10px' }}
                                    />
                                  </div>

                                  {/* Checkbox-Confirm-Label */}
                                  {field.type === 'checkbox' && (
                                    <div style={{ marginLeft: 32, marginTop: 6 }}>
                                      <input
                                        className="form-input"
                                        placeholder={isDe
                                          ? 'Text neben Checkbox (optional, Default: „Ja, bestätigen")'
                                          : 'Text next to checkbox (optional, default: „Yes, confirm")'}
                                        value={field.confirmLabel || ''}
                                        onChange={e => updateSubEventCustomField(se.id, field.id, { confirmLabel: e.target.value })}
                                        style={{ width: '100%', fontSize: '0.82rem', padding: '6px 10px' }}
                                      />
                                    </div>
                                  )}

                                  {/* Optionen-Editor für Dropdown-Felder */}
                                  {field.type === 'select' && (
                                    <div style={{
                                      marginTop: 10, marginLeft: 32, padding: '12px 14px',
                                      background: '#fff',
                                      border: '1px solid var(--dex-gray-200)',
                                      borderRadius: 8,
                                    }}>
                                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-700)', fontWeight: 600, marginBottom: 8 }}>
                                        {isDe ? 'Antwort-Optionen' : 'Answer options'}
                                      </div>
                                      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                                        {field.options.map((opt, oidx) => (
                                          <div key={oidx} style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                                            <span style={{
                                              flexShrink: 0, width: 22, height: 22, borderRadius: '50%',
                                              background: 'var(--dex-gray-200)', color: 'var(--dex-gray-700)',
                                              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                              fontWeight: 700, fontSize: '0.72rem',
                                            }}>{oidx + 1}</span>
                                            <input
                                              className="form-input"
                                              placeholder={isDe ? `Option ${oidx + 1}` : `Option ${oidx + 1}`}
                                              value={opt}
                                              onChange={e => {
                                                const next = field.options.slice();
                                                next[oidx] = e.target.value;
                                                updateSubEventCustomField(se.id, field.id, { options: next });
                                              }}
                                              style={{ flex: 1, fontSize: '0.85rem', padding: '6px 10px' }}
                                            />
                                            {field.options.length > 1 && (
                                              <button
                                                type="button"
                                                onClick={() => {
                                                  const next = field.options.filter((_, i) => i !== oidx);
                                                  updateSubEventCustomField(se.id, field.id, { options: next });
                                                }}
                                                title={isDe ? 'Option entfernen' : 'Remove option'}
                                                style={{ background: 'none', border: 'none', color: 'var(--dex-gray-500)', padding: 4, cursor: 'pointer' }}
                                              >
                                                <X size={14} />
                                              </button>
                                            )}
                                          </div>
                                        ))}
                                        <button
                                          type="button"
                                          onClick={() => updateSubEventCustomField(se.id, field.id, { options: [...field.options, ''] })}
                                          style={{ alignSelf: 'flex-start', background: 'none', border: '1px dashed var(--dex-gray-300)', padding: '4px 12px', fontSize: '0.78rem', borderRadius: 6, cursor: 'pointer', color: 'var(--dex-gray-700)', display: 'inline-flex', alignItems: 'center', gap: 4, marginTop: 4 }}
                                        >
                                          <Plus size={12} /> {isDe ? 'Option hinzufügen' : 'Add option'}
                                        </button>
                                      </div>
                                    </div>
                                  )}
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              </div>{/* v16.5: close plain wrapper div (Step 5) — kein Greyout mehr */}
              </div>{/* v15.0: close activeFieldsTabIdx===0 wrapper (Top-Level Felder + hidden Bereich 2) */}

              {/* v18.75: Sicherheitshinweis vor dem Absenden — eigene Section
                  ganz unten in Schritt 5 (gilt event-weit, daher außerhalb der
                  Feld-Tabs). */}
              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '12px 16px', marginTop: 18, marginBottom: 4,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={confirmDialogEnabled}
                    onChange={e => setConfirmDialogEnabled(e.target.checked)}
                    style={{ marginTop: 3, cursor: 'pointer' }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Sicherheitshinweis vor dem Absenden anzeigen?' : 'Show a confirmation prompt before submitting?'}</strong>
                    <InfoTooltip text={isDe
                      ? <>
                          <strong>Was du hier einstellst:</strong> ob nach dem Klick auf <strong>„Anmelden“</strong> noch ein <strong>Bestätigungs-Dialog</strong> erscheint, bevor die Anmeldung wirklich abgeschickt wird. Default: <strong>nein</strong>.<br /><br />
                          <strong>Zwei Varianten:</strong> die <strong>Auswahl-Übersicht</strong> listet Haupt-Event und gewählte Sub-Events auf — der Teilnehmer kann vor dem Absenden einzelne Punkte noch ab- oder zuwählen. Der <strong>eigene Hinweistext</strong> zeigt stattdessen einen frei formulierten Hinweis (z.B. zu Verbindlichkeit oder Storno-Fristen), den der Teilnehmer bestätigen muss.<br /><br />
                          <strong>Auswirkung für Teilnehmer:</strong> ein zusätzlicher, bewusster Bestätigungsschritt — schützt vor versehentlichen Anmeldungen.
                        </>
                      : <>
                          <strong>What this controls:</strong> whether a <strong>confirmation dialog</strong> appears after clicking <strong>“Register”</strong>, before the registration is actually submitted. Default: <strong>no</strong>.<br /><br />
                          <strong>Two variants:</strong> the <strong>selection summary</strong> lists the main event and selected sub-events — the attendee can de-/select items before submitting. The <strong>custom hint text</strong> instead shows a free-text note (e.g. about binding registration or cancellation deadlines) the attendee must acknowledge.<br /><br />
                          <strong>For attendees:</strong> an extra, deliberate confirmation step — protects against accidental registrations.
                        </>
                    } />
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? 'Default: nein — wenn aktiviert, muss der Teilnehmer nach „Anmelden" noch einen Dialog bestätigen.'
                        : 'Default: no — when enabled, the attendee has to confirm a dialog after clicking „Register".'}
                    </span>
                  </span>
                </label>
                {confirmDialogEnabled && (
                  <div style={{ marginTop: 12, paddingLeft: 30 }}>
                    <div style={{ fontSize: '0.82rem', fontWeight: 600, marginBottom: 8, color: 'var(--dex-gray-700)' }}>
                      {isDe ? 'Was soll der Dialog zeigen?' : 'What should the dialog show?'}
                    </div>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', marginBottom: 10 }}>
                      <input type="radio" name="confirmDialogMode" checked={confirmDialogMode !== 'freetext'} onChange={() => setConfirmDialogMode('summary')} style={{ marginTop: 3 }} />
                      <span style={{ fontSize: '0.85rem' }}>
                        <strong>{isDe ? 'Auswahl-Übersicht' : 'Selection summary'}</strong> — {isDe
                          ? 'listet Haupt-Event und gewählte Sub-Events auf; der Teilnehmer kann vor dem Absenden einzelne Punkte noch ab- oder zuwählen.'
                          : 'lists the main event and selected sub-events; the attendee can de-/select items before submitting.'}
                      </span>
                    </label>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer' }}>
                      <input type="radio" name="confirmDialogMode" checked={confirmDialogMode === 'freetext'} onChange={() => setConfirmDialogMode('freetext')} style={{ marginTop: 3 }} />
                      <span style={{ fontSize: '0.85rem' }}>
                        <strong>{isDe ? 'Eigener Hinweistext' : 'Custom hint text'}</strong> — {isDe
                          ? 'zeigt einen frei formulierten Hinweis, den der Teilnehmer bestätigen muss.'
                          : 'shows a free-text note the attendee must acknowledge.'}
                      </span>
                    </label>
                    {confirmDialogMode === 'freetext' && (
                      <textarea
                        className="form-input"
                        value={confirmDialogText}
                        onChange={e => setConfirmDialogText(e.target.value)}
                        rows={3}
                        placeholder={isDe
                          ? 'z.B. „Bitte beachte: Die Anmeldung ist verbindlich. Eine Stornierung ist nur bis 3 Tage vor dem Event möglich."'
                          : 'e.g. „Please note: registration is binding. Cancellation is only possible up to 3 days before the event."'}
                        style={{ marginTop: 10, width: '100%', resize: 'vertical' }}
                      />
                    )}
                  </div>
                )}
              </div>

              </div>
  );
};
