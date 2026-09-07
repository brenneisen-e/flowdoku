/* FieldsStep — aus EventCreationPage.tsx ausgelagert (Zeilen 14861-16968 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 4` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher.
 *
 * v31.2: Modernisiert nach docs/ui-leitfaden.md. Reihenfolge jetzt: Fragen
 * (das Eigentliche) → Sprache des Formulars → Bestätigung vor dem Absenden.
 * Der Feld-Editor folgt dem Erleben des Teilnehmers: Fragetext → Antworten →
 * Typ-Folgen (Mails, Suchkreis) → Sichtbarkeit → Hilfetext. Alle Handler,
 * Bedingungen und Setter sind unverändert; die Datei schreibt weiterhin nur
 * customFields / subEvents[].customFields und die Formular-Schalter. */
import * as React from 'react';
import WizardHint from '../../WizardHint';
import { b2runKoelnTemplateFields, isB2RunKoelnTitle } from '../../../data/b2runKoeln';
import { CustomField } from '../../../services/EventService';
import { AlertCircle, Check, ChevronDown, ChevronUp, Info, Plus, Trash2, X } from '../../Icons';
import { InfoTooltip } from '../../InfoTooltip';
import { Icon } from '@fluentui/react/lib/Icon';
// v31.2: Gemeinsame UI-Klassen (Karten, Chips, Schalter-Zeilen, Hover) —
// siehe docs/ui-leitfaden.md. Inline-Styles können kein :hover.
import { cx } from '../../dexUi';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { CustomFieldInput } from '../../wizard/customFieldInput';
import { FieldTypeSuggestion } from '../../wizard/FieldTypeSuggestion';
import { StepBadge } from '../../wizard/StepBadge';
import { FieldDescEditor } from '../../wizard/FieldDescEditor';
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
  const { activeFieldsTabIdx, addCustomField, addStartblock, addSubEventCustomField, askSalutation, b2runStartblocks, bilingualFields, childTermPlural, confirmDialogEnabled, confirmDialogMode, confirmDialogText, copyParentFieldsToSubEvent, customFields, dragFieldId, dragOverFieldId, fieldExpandOverride, isDe, moveCustomField, newStartblock, openSuggestedModal, registrationLanguage, removeCustomField, removeStartblock, removeSubEventCustomField, renderShowIfConfig, renderStepIntro, reorderMode, setAskSalutation, setBilingualFields, setConfirmDialogEnabled, setConfirmDialogMode, setConfirmDialogText, setCustomFields, setDragFieldId, setDragOverFieldId, setNewStartblock, setRegistrationLanguage, setReorderMode, setSubEvents, splitLabelA, splitLabelB, subEvents, subEventsOnlyMode, t, title, toggleFieldExpand, updateCustomField, updateSubEventCustomField, useSplitCapacities } = p;
  // v31.2: Wiederkehrende Bausteine des Schritts an EINER Stelle — Typ-
  // Beschriftung, Nummern-Kreis, Sprach-/Kategorie-Marke, Breite des Typ-
  // Dropdowns. Vorher stand jeder davon drei- bis sechsmal als Inline-Style-
  // Block im JSX (Hauptevent, Sub-Event-Reiter, toter Bereich 2). Keine
  // Hooks, nur Konstanten — die Hook-Reihenfolge bleibt leer wie zuvor.
  const typeLabel = (ty: CustomFieldInput['type']): string => {
    switch (ty) {
      case 'text': return isDe ? 'Text (Freitext)' : 'Text (free text)';
      case 'select': return 'Dropdown';
      case 'number': return isDe ? 'Zahl' : 'Number';
      case 'checkbox': return 'Checkbox';
      case 'date': return isDe ? 'Datum (Kalender)' : 'Date (calendar)';
      case 'daterange': return isDe ? 'Übernachtungs-Zeitraum (Kalender + Nächte)' : 'Stay period (calendar + nights)';
      case 'user': return 'Person';
      case 'roommate': return 'Roommate';
      default: return isDe ? 'Dokument (Upload)' : 'Document (upload)';
    }
  };
  // Sub-Event-Felder kennen weder People-Picker noch Upload (v15.3).
  const SUB_TYPES: CustomFieldInput['type'][] = ['text', 'select', 'number', 'checkbox', 'date', 'daterange'];
  const MAIN_TYPES: CustomFieldInput['type'][] = [...SUB_TYPES, 'user', 'roommate', 'document'];
  const numBadge: React.CSSProperties = {
    flexShrink: 0, width: 26, height: 26, borderRadius: '50%',
    background: 'var(--dex-green, #86bc25)', color: '#fff',
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    fontWeight: 700, fontSize: '0.78rem', lineHeight: 1,
  };
  const enBadge = <span className="dex-ui-pill dex-ui-pill--blue" style={{ fontSize: '0.66rem', padding: '2px 7px', flexShrink: 0 }}>EN</span>;
  const catBadge = <span className="dex-ui-pill dex-ui-pill--green" style={{ fontSize: '0.66rem', padding: '2px 7px', flexShrink: 0 }}>{isDe ? 'KAT' : 'CAT'}</span>;
  // v11.4: feste Breite, damit Frage + Typ + Pflicht + X in einer Zeile bleiben.
  const typeSelectStyle: React.CSSProperties = { flex: '0 0 210px', width: 210, fontWeight: 600, color: 'var(--dex-green-darker, #4a7c1f)' };
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              <h2 className="dex-step-head-title">
                <span className="dex-step-eyebrow">{isDe ? 'Schritt 5 von 9' : 'Step 5 of 9'}</span>
                {isDe ? 'Fragen im Anmeldeformular' : 'Questions on the registration form'}
              </h2>
              <p className="dex-step-head-lead">
                {isDe
                  ? <><strong>Optional</strong> — hier legst du fest, was das Anmeldeformular über die Profildaten hinaus fragt: von der T-Shirt-Größe bis zur Pflicht-Checkbox für AGB oder Datenschutz. Braucht dein Event keine Zusatzfragen, lässt du den Schritt einfach leer.</>
                  : <><strong>Optional</strong> — here you decide what the registration form asks beyond the profile data: from the T-shirt size to a required privacy / terms checkbox. If your event needs no extra questions, simply leave this step empty.</>}
              </p>

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

              {/* v31.2: Der Datenschutz-Hinweis ist kein Aufklapper mehr, sondern
                  ein sichtbarer Kasten — der Organizer soll ihn lesen, BEVOR er
                  die erste Frage anlegt. Der Wortlaut bleibt deckungsgleich mit
                  den Nutzungsbedingungen (v7.35). */}
              <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginBottom: 16 }}>
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>
                  <strong>{isDe ? 'Sammle keine sensiblen personenbezogenen Daten.' : 'Do not collect sensitive personal data.'}</strong>{' '}
                  {isDe
                    ? <>Das heißt: keine Daten bezüglich Rasse oder ethnischer Herkunft, religiöser oder philosophischer Überzeugungen, Gewerkschaftsmitgliedschaft, politischer Meinungen, medizinischer oder gesundheitlicher Zustände oder Informationen über das Sexualleben oder die sexuelle Orientierung einer Person. Falls sensible personenbezogene Daten gesammelt werden müssen, kontaktiere zuerst das Team unter <a href="mailto:privacy@deloitte.de" style={{ color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600 }}>privacy@deloitte.de</a>.</>
                    : <>That means: no data on race or ethnic origin, religious or philosophical beliefs, trade-union membership, political opinions, medical or health conditions, or information about a person&apos;s sex life or sexual orientation. If sensitive personal data must be collected, contact the team first at <a href="mailto:privacy@deloitte.de" style={{ color: 'var(--dex-orange-dark, #b35a00)', fontWeight: 600 }}>privacy@deloitte.de</a>.</>}
                </span>
              </div>

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
                  <div className="dex-ui-card dex-ui-card--accent" style={{ marginBottom: 16 }}>
                    <div className="dex-ui-label">
                      {isDe ? 'B2Run-Köln-Vorlage' : 'B2Run Köln template'}
                      <span className={cx('dex-ui-pill', b2rkMissing.length === 0 ? 'dex-ui-pill--green' : 'dex-ui-pill--orange')}>
                        {b2rkMissing.length === 0
                          ? (isDe ? 'vollständig' : 'complete')
                          : (isDe ? `${b2rkMissing.length} fehlen noch` : `${b2rkMissing.length} missing`)}
                      </span>
                    </div>
                    <p className="dex-ui-help" style={{ margin: '0 0 8px' }}>
                      {isDe
                        ? 'Dieses Event sieht nach dem B2Run Köln aus — übernimm die offiziellen Meldefelder mit einem Klick. Der Excel-Export im Organizer Center füllt damit die offizielle Meldedatei exakt aus.'
                        : 'This event looks like the B2Run Köln — adopt the official entry fields with one click. The Excel export in the Organizer Center then fills in the official entry file exactly.'}
                    </p>
                    <p className="dex-ui-help" style={{ margin: '0 0 10px' }}>
                      {isDe
                        ? <><strong>Startblock:</strong> wird NICHT als Feld abgefragt — die Teilnehmenden wählen Durchstarter/Funstarter über die <strong>Gruppen-Auswahl</strong> (Split-Kapazität im Schritt &bdquo;Kapazität &amp; Sichtbarkeit&ldquo;). Der Export übersetzt die Gruppe automatisch in die offiziellen Startblock-Texte.</>
                        : <><strong>Start block:</strong> is NOT asked as a field — attendees pick Durchstarter/Funstarter via the <strong>group selection</strong> (split capacity in the &ldquo;Capacity &amp; visibility&rdquo; step). The export automatically translates the group into the official start-block texts.</>}
                    </p>
                    <div style={{ marginBottom: 12 }}>
                      {b2rkTemplate.map(f => {
                        const exists = customFields.some(p => p.id === f.id);
                        return (
                          <div key={f.id} className="dex-ui-row dex-ui-row--bordered" style={{ padding: '6px 8px' }}>
                            <span style={{ width: 16, display: 'inline-flex', justifyContent: 'center', color: 'var(--dex-green)' }} title={exists ? (isDe ? 'Bereits im Event' : 'Already in the event') : undefined}>
                              {exists ? <Check size={14} /> : null}
                            </span>
                            <span className="dex-ui-row-main" style={{ fontSize: '0.84rem', fontWeight: exists ? 400 : 600 }}>{f.label}</span>
                            <span className="dex-ui-pill dex-ui-pill--gray">{b2rkTypeTag(f.type)}</span>
                          </div>
                        );
                      })}
                    </div>
                    <button
                      type="button"
                      className="btn btn-secondary dex-ui-btn-sm"
                      disabled={b2rkMissing.length === 0}
                      onClick={() => setCustomFields(prev => [
                        ...prev,
                        ...b2runKoelnTemplateFields(isDe)
                          .filter(f => !prev.some(p => p.id === f.id))
                          .map(f => ({ ...f, options: f.options || [] })),
                      ])}
                    >
                      {b2rkMissing.length === 0
                        ? (isDe ? 'Alles übernommen' : 'All adopted')
                        : (isDe ? 'Fehlende Felder übernehmen' : 'Adopt missing fields')}
                    </button>
                  </div>
                );
              })()}

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
                <div className="dex-ui-card dex-ui-card--accent" style={{ marginBottom: 16 }}>
                  <div className="dex-ui-label">
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
                  </div>
                  <p className="dex-ui-help" style={{ margin: '0 0 10px' }}>
                    {t('create.startblocks.hint')}
                  </p>

                  {/* Bestehende Startblöcke als Liste */}
                  {b2runStartblocks.length > 0 && (
                    <div style={{ marginBottom: 10 }}>
                      {b2runStartblocks.map((block, idx) => (
                        <div key={idx} className="dex-ui-row dex-ui-row--bordered">
                          <Icon iconName="Running" style={{ fontSize: 16, color: 'var(--dex-green-dark, #6b9a1e)', flexShrink: 0 }} />
                          <span className="dex-ui-row-main" style={{ fontSize: '0.88rem' }}>{block}</span>
                          <span className="dex-ui-row-actions">
                            <button
                              type="button"
                              className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                              onClick={() => removeStartblock(block)}
                              title={t('create.startblocks.remove')}
                              aria-label={t('create.startblocks.remove')}
                            >
                              <Trash2 size={14} />
                            </button>
                          </span>
                        </div>
                      ))}
                    </div>
                  )}

                  {/* Neues Startblock hinzufügen */}
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                    <input
                      type="text"
                      className="dex-ui-input"
                      value={newStartblock}
                      onChange={e => setNewStartblock(e.target.value)}
                      onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); addStartblock(); } }}
                      placeholder={t('create.startblocks.placeholder')}
                      style={{ flex: '1 1 220px', width: 'auto' }}
                    />
                    <button
                      type="button"
                      className="btn btn-secondary dex-ui-btn-sm"
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
                  Tabs N>0 = schlanke per-Sub-Event-Felder-UI. Im
                  subEventsOnlyMode wird Tab 0 zu „Übergreifende Felder" —
                  die wirken dann auf alle Sub-Event-Anmeldungen. */}
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
                const seName = se.title || (isDe ? '(unbenanntes Sub-Event)' : '(unnamed sub-event)');
                return (
                  <div className="dex-ui-section">
                    <h3 className="dex-ui-section-title">
                      {isDe ? `Fragen für ${seName}` : `Questions for ${seName}`}
                    </h3>
                    {/* v15.6: Lead-paragraph analog Hauptevent-Tab. */}
                    <p className="dex-ui-section-desc">
                      {isDe
                        ? <><strong>Optional</strong> — Vorname, Nachname, E-Mail und Profil-Daten (Job Title, Standort, Department, Telefon) werden automatisch erfasst. Hier ergänzt du <strong>nur Zusatzfragen speziell für dieses Sub-Event</strong>; braucht es keine, bleibt die Liste leer.</>
                        : <><strong>Optional</strong> — first name, last name, email and profile data (job title, location, department, phone) are captured automatically. Here you only add <strong>extra questions specific to this sub-event</strong>; if it needs none, leave the list empty.</>}
                    </p>

                    {/* v18.62: Datenschutz-Hinweis hier ENTFERNT — er steht bereits
                        einmal oben in Schritt 5 (über der Tab-Leiste). Eine
                        Wiederholung pro Sub-Event-Tab ist redundant. */}

                    <div className="dex-ui-inline" style={{ marginBottom: 12 }}>
                      <button
                        type="button"
                        className="btn btn-primary dex-ui-btn-sm"
                        onClick={() => addSubEventCustomField(se.id)}
                      >
                        <Plus size={14} /> {isDe ? 'Frage hinzufügen' : 'Add question'}
                      </button>
                      {customFields.length > 0 && seFields.length === 0 && (
                        <button
                          type="button"
                          className="btn btn-secondary dex-ui-btn-sm"
                          onClick={() => copyParentFieldsToSubEvent(se.id)}
                          title={isDe ? `Dupliziert die ${customFields.length} Felder vom Hauptevent als Startpunkt` : 'Duplicates the main-event fields as a starting point'}
                        >
                          {isDe ? `Fragen vom Hauptevent übernehmen (${customFields.length})` : `Copy questions from main event (${customFields.length})`}
                        </button>
                      )}
                    </div>

                    {/* v15.3: „Anrede abfragen"-Toggle pro Sub-Event.
                        v31.2: als Schalter-Zeile mit Folge-Satz; der Kopier-Knopf
                        steht NEBEN der Zeile, nicht im Label — ein Knopf im Label
                        würde beim Klick zugleich die Checkbox umschalten. */}
                    <div style={{ display: 'flex', gap: 8, alignItems: 'stretch', flexWrap: 'wrap', marginBottom: 12 }}>
                      <label className={cx('dex-ui-toggle-row', !!se.askSalutation && 'is-active')} style={{ flex: '1 1 280px' }}>
                        <input
                          type="checkbox"
                          checked={!!se.askSalutation}
                          onChange={e => updateSub({ askSalutation: e.target.checked })}
                        />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">
                            {isDe ? 'Anrede für dieses Sub-Event abfragen' : 'Ask for salutation on this sub-event'}
                          </span>
                          <span className="dex-ui-toggle-row-desc">
                            {isDe
                              ? 'Dann steht über dem Vornamen ein Pflicht-Dropdown (Frau / Herr / Divers / Keine Angabe).'
                              : 'Then a required dropdown (Mrs / Mr / Diverse / Prefer not to say) appears above the first name.'}
                          </span>
                        </span>
                      </label>
                      <button
                        type="button"
                        className="dex-ui-textbtn dex-ui-textbtn--muted"
                        onClick={() => updateSub({ askSalutation: askSalutation })}
                        title={isDe
                          ? 'Übernimmt die Anrede-Abfrage-Einstellung vom Hauptevent'
                          : 'Copies the salutation toggle from the main event'}
                      >
                        {isDe ? 'Wie Hauptevent' : 'Same as main event'}
                      </button>
                    </div>

                    {seFields.length === 0 ? (
                      <div className="dex-ui-empty">
                        <div className="dex-ui-empty-title">{isDe ? 'Noch keine eigenen Fragen' : 'No questions of its own yet'}</div>
                        {isDe
                          ? 'Füge oben eine Frage hinzu oder übernimm die Fragen vom Hauptevent.'
                          : 'Add a question above or copy the questions from the main event.'}
                      </div>
                    ) : (
                      <div className="dex-ui-stack" style={{ gap: 12 }}>
                        {seFields.map((field, idx) => (
                          <div key={field.id} className="dex-ui-card" style={{ padding: '14px 16px' }}>
                            <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
                              <span style={numBadge}>{idx + 1}</span>
                              <input
                                className="form-input"
                                value={field.label}
                                placeholder={isDe ? 'Wie lautet die Frage? (z.B. „Welche Strecke läufst du?")' : 'What is the question? (e.g. „Which distance will you run?")'}
                                onChange={e => updateSubEventCustomField(se.id, field.id, { label: e.target.value })}
                                disabled={inherit}
                                style={{ flex: '1 1 260px', minWidth: 180, fontSize: '0.95rem', fontWeight: 600, padding: '8px 12px', minHeight: 0 }}
                              />
                              <select
                                className="dex-ui-select"
                                value={field.type}
                                disabled={inherit}
                                onChange={e => updateSubEventCustomField(se.id, field.id, { type: e.target.value as CustomFieldInput['type'] })}
                                title={isDe ? 'Art der Antwort' : 'Answer type'}
                                style={typeSelectStyle}
                              >
                                {SUB_TYPES.map(ty => <option key={ty} value={ty}>{typeLabel(ty)}</option>)}
                              </select>
                              <label
                                className={cx('dex-ui-chip', field.required && 'is-active', inherit && 'is-disabled')}
                                title={isDe ? 'Pflicht: ohne Antwort lässt sich die Anmeldung nicht absenden' : 'Required: the registration cannot be submitted without an answer'}
                              >
                                <input
                                  type="checkbox"
                                  checked={field.required}
                                  disabled={inherit}
                                  onChange={e => updateSubEventCustomField(se.id, field.id, { required: e.target.checked })}
                                  style={{ display: 'none' }}
                                />
                                {field.required && <Check size={12} />}
                                {t('create.required')}
                              </label>
                              <button
                                type="button"
                                className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                                onClick={() => removeSubEventCustomField(se.id, field.id)}
                                disabled={inherit}
                                title={isDe ? 'Frage entfernen' : 'Remove question'}
                                aria-label={isDe ? 'Frage entfernen' : 'Remove question'}
                                style={{ marginLeft: 'auto' }}
                              >
                                <X size={16} />
                              </button>
                            </div>
                            {/* v24.25: Datum-Feld → optional Uhrzeit mit abfragen. */}
                            {field.type === 'date' && (
                              <label className="dex-ui-inline" style={{ marginLeft: 36, marginTop: 8, cursor: inherit ? 'default' : 'pointer', fontSize: '0.82rem', color: 'var(--dex-gray-700)' }}>
                                <input
                                  type="checkbox"
                                  checked={!!field.withTime}
                                  disabled={inherit}
                                  onChange={e => updateSubEventCustomField(se.id, field.id, { withTime: e.target.checked })}
                                  style={{ accentColor: 'var(--dex-green, #86bc25)' }}
                                />
                                {isDe ? 'Auch die Uhrzeit abfragen' : 'Also ask for the time'}
                              </label>
                            )}
                            {field.type === 'daterange' && (
                              <div className="dex-ui-inline" style={{ marginLeft: 36, marginTop: 8, alignItems: 'flex-end', gap: 10 }}>
                                <div>
                                  <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 3 }}>{isDe ? 'Buchbar ab' : 'Bookable from'}</div>
                                  <input type="date" className="dex-ui-input dex-ui-input--sm" disabled={inherit} style={{ width: 160 }}
                                    value={field.rangeStart || ''} onChange={e => updateSubEventCustomField(se.id, field.id, { rangeStart: e.target.value })} />
                                </div>
                                <div>
                                  <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 3 }}>{isDe ? 'Buchbar bis' : 'Bookable until'}</div>
                                  <input type="date" className="dex-ui-input dex-ui-input--sm" disabled={inherit} style={{ width: 160 }}
                                    value={field.rangeEnd || ''} onChange={e => updateSubEventCustomField(se.id, field.id, { rangeEnd: e.target.value })} />
                                </div>
                                <div>
                                  <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 3 }}>{isDe ? 'Max. Nächte' : 'Max. nights'}</div>
                                  <input type="number" min={0} className="dex-ui-input dex-ui-input--sm" disabled={inherit} style={{ width: 100 }}
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
                            {/* v31.2: Reihenfolge wie beim Hauptevent — erst die
                                Antwortmöglichkeiten, dann wann die Frage erscheint,
                                zuletzt der Hilfetext. */}
                            <div className="dex-ui-stack" style={{ marginLeft: 36, marginTop: 12 }}>
                              {field.type === 'select' && (
                                <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '12px 14px' }}>
                                  <div className="dex-ui-label" style={{ marginBottom: 8 }}>
                                    {isDe ? 'Antwortmöglichkeiten' : 'Answer options'}
                                  </div>
                                  <div className="dex-ui-stack" style={{ gap: 6 }}>
                                    {field.options.map((opt, oidx) => (
                                      <div key={oidx} style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                                        <span className="dex-ui-muted" style={{ flexShrink: 0, width: 20, textAlign: 'right', fontWeight: 600 }}>{oidx + 1}.</span>
                                        <input
                                          className="dex-ui-input dex-ui-input--sm"
                                          placeholder={isDe ? `Option ${oidx + 1}` : `Option ${oidx + 1}`}
                                          value={opt}
                                          disabled={inherit}
                                          onChange={e => {
                                            const next = field.options.slice();
                                            next[oidx] = e.target.value;
                                            updateSubEventCustomField(se.id, field.id, { options: next });
                                          }}
                                          style={{ flex: 1, width: 'auto' }}
                                        />
                                        {field.options.length > 1 && (
                                          <button
                                            type="button"
                                            className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                                            disabled={inherit}
                                            onClick={() => {
                                              const next = field.options.filter((_, i) => i !== oidx);
                                              updateSubEventCustomField(se.id, field.id, { options: next });
                                            }}
                                            title={isDe ? 'Option entfernen' : 'Remove option'}
                                            aria-label={isDe ? 'Option entfernen' : 'Remove option'}
                                          >
                                            <X size={14} />
                                          </button>
                                        )}
                                      </div>
                                    ))}
                                    <button
                                      type="button"
                                      className="dex-ui-textbtn"
                                      disabled={inherit}
                                      onClick={() => updateSubEventCustomField(se.id, field.id, { options: [...field.options, ''] })}
                                      style={{ alignSelf: 'flex-start' }}
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
                              <div>
                                <div className="dex-ui-label" style={{ marginBottom: 6 }}>
                                  {isDe ? 'Hilfetext für Teilnehmer' : 'Help text for attendees'}
                                  <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                                </div>
                                <input
                                  className="dex-ui-input dex-ui-input--sm"
                                  placeholder={isDe
                                    ? 'z.B. „Die Strecke kannst du bis 3 Tage vorher ändern." — erscheint neben der Frage'
                                    : 'e.g. „You can change the distance up to 3 days before." — shown next to the question'}
                                  value={field.helpText || ''}
                                  disabled={inherit}
                                  onChange={e => updateSubEventCustomField(se.id, field.id, { helpText: e.target.value })}
                                />
                                {/* v24.14 BUG-FIX: helpTextStyle-Wahl fehlte bei Sub-Event-Feldern
                                    (nur das Hauptevent hatte sie) — „Text unter dem Feld" wurde
                                    deshalb beim Sub-Event nie gespeichert. */}
                                {field.helpText && field.helpText.trim() && (
                                  <div className="dex-ui-inline" style={{ marginTop: 8 }}>
                                    <span className="dex-ui-muted" style={{ fontWeight: 600 }}>{isDe ? 'Wo erscheint er?' : 'Where does it show?'}</span>
                                    <label className={cx('dex-ui-chip', (field.helpTextStyle || 'tooltip') !== 'inline' && 'is-active', inherit && 'is-disabled')}>
                                      <input
                                        type="radio"
                                        name={`seHelpStyle-${se.id}-${field.id}`}
                                        disabled={inherit}
                                        checked={(field.helpTextStyle || 'tooltip') !== 'inline'}
                                        onChange={() => updateSubEventCustomField(se.id, field.id, { helpTextStyle: 'tooltip' })}
                                        style={{ display: 'none' }}
                                      />
                                      {isDe ? 'Als „i"-Info-Box (Hover)' : 'As „i" info box (hover)'}
                                    </label>
                                    <label className={cx('dex-ui-chip', field.helpTextStyle === 'inline' && 'is-active', inherit && 'is-disabled')}>
                                      <input
                                        type="radio"
                                        name={`seHelpStyle-${se.id}-${field.id}`}
                                        disabled={inherit}
                                        checked={field.helpTextStyle === 'inline'}
                                        onChange={() => updateSubEventCustomField(se.id, field.id, { helpTextStyle: 'inline' })}
                                        style={{ display: 'none' }}
                                      />
                                      {isDe ? 'Als Text unter der Frage' : 'As text below the question'}
                                    </label>
                                  </div>
                                )}
                              </div>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })()}

              {/* Tab 0 (Haupt-Event bzw. Übergreifende Felder). v31.2: Zwei
                  Sektionen in der Reihenfolge, in der ein Organizer denkt —
                  erst die Fragen (das Eigentliche), dann die Sprache des
                  Formulars (Feineinstellung). Vorher standen die Einstellungs-
                  Karten VOR der Feld-Liste. */}
              <div style={{ display: activeFieldsTabIdx === 0 ? 'block' : 'none' }}>
              {/* v16.5: In Step 5 (Felder) ist im subEventsOnlyMode KEIN
                  Greyout — die Felder im ersten Tab sind „übergreifend"
                  und werden bei JEDER Sub-Event-Anmeldung abgefragt, also
                  in dem Modus besonders relevant. */}
              <div className="dex-ui-section">
                {/* v15.0: im subEventsOnlyMode lautet die Überschrift
                    „Übergreifend für alle <childTermPlural>". */}
                <h3 className="dex-ui-section-title">
                  <StepBadge n={24} />
                  {subEventsOnlyMode
                    ? (isDe
                        ? `Übergreifend für alle ${(childTermPlural || 'Sub-Events').trim() || 'Sub-Events'}`
                        : `Across all ${(childTermPlural || 'sub-events').trim() || 'sub-events'}`)
                    : (isDe ? 'Fragen für das Hauptevent' : 'Questions for the main event')}
                </h3>
                <p className="dex-ui-section-desc">
                  {isDe
                    ? <>Diese Fragen stellt jede Anmeldung — egal ob mit oder ohne Sub-Event. Vorname, Nachname, E-Mail, Job Title, Standort und Department kommen automatisch aus dem Deloitte-Profil und musst du nicht abfragen.</>
                    : <>Every registration asks these questions — with or without a sub-event. First name, last name, email, job title, location and department come from the Deloitte profile automatically; no need to ask for them.</>}
                  <InfoTooltip text={isDe
                    ? <>
                        <strong>Automatisch erfasst</strong> (aus dem Deloitte-Profil, bei jeder Anmeldung): Vorname, Nachname, E-Mail, Job Title, Standort, Department.<br /><br />
                        Hier ergänzt du <strong>nur zusätzliche Fragen</strong>, die du speziell für dieses Event brauchst — vom T-Shirt-Größen-Dropdown bis zur Pflicht-Checkbox für AGB / Datenschutz. Die <strong>Anrede</strong> holst du bei Bedarf über &bdquo;Vorgeschlagene Felder&ldquo; dazu.<br /><br />
                        Fragen nur für einen Termin stellst du im Reiter des jeweiligen Sub-Events oben.
                      </>
                    : <>
                        <strong>Captured automatically</strong> (from the Deloitte profile, for every registration): first name, last name, email, job title, location, department.<br /><br />
                        Here you only add <strong>extra questions</strong> specific to this event — from a T-shirt size dropdown to a required privacy / terms checkbox. Add the <strong>salutation</strong> via &ldquo;Suggested fields&rdquo; if you need it.<br /><br />
                        Questions for one date only go into that sub-event&apos;s tab above.
                      </>} />
                </p>

                <div className="dex-ui-inline" style={{ marginBottom: 12 }}>
                  <button className="btn btn-primary dex-ui-btn-sm" onClick={addCustomField}>
                    <Plus size={14} /> {t('create.addfield')}
                  </button>
                  <button
                    type="button"
                    className="btn btn-secondary dex-ui-btn-sm"
                    onClick={openSuggestedModal}
                    title={isDe ? 'Fertige Fragen aus dem Katalog übernehmen (Anrede, T-Shirt, Allergien …)' : 'Adopt ready-made questions from the catalog (salutation, T-shirt, allergies …)'}
                  >
                    {isDe ? 'Vorgeschlagene Felder' : 'Suggested fields'}
                  </button>
                  {customFields.length > 1 && (
                    <button
                      type="button"
                      className={cx('dex-ui-chip', reorderMode && 'is-active')}
                      onClick={() => setReorderMode(prev => !prev)}
                      title={isDe ? 'Felder per Hoch/Runter-Pfeile sortieren' : 'Reorder fields with up/down arrows'}
                    >
                      {reorderMode ? <Check size={12} /> : null}
                      {reorderMode
                        ? (isDe ? 'Fertig sortiert' : 'Done reordering')
                        : (isDe ? 'Reihenfolge ändern' : 'Reorder')}
                    </button>
                  )}
                </div>
                {/* v24.25: Erklär-Box — welche Feldarten es gibt und was sie tun
                    (aufklappbar, grau). */}
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
                {/* v31.2: Leerer Zustand statt leerer Fläche — sagt, was das
                    Formular ohne eigene Fragen tut und wo es weitergeht. */}
                {customFields.length === 0 && !askSalutation && (
                  <div className="dex-ui-empty" style={{ marginBottom: 12 }}>
                    <div className="dex-ui-empty-icon"><Plus size={18} /></div>
                    <div className="dex-ui-empty-title">{isDe ? 'Noch keine Zusatzfragen' : 'No extra questions yet'}</div>
                    {isDe
                      ? 'Das Formular fragt bisher nur die Profildaten ab. Füge oben eine Frage hinzu oder wähle aus den Vorschlägen.'
                      : 'So far the form only asks for the profile data. Add a question above or pick from the suggestions.'}
                  </div>
                )}
                {/* v22.38: Anrede als Standard-Feld-Zeile — aktiviert über das
                    Vorgeschlagene-Felder-Modal (Eintrag „Anrede"), entfernbar
                    über das X. Kein Custom-Field (eigener askSalutation-Flag). */}
                {askSalutation && (
                  <div className="dex-ui-row dex-ui-card" style={{ marginBottom: 12, padding: '10px 14px', borderRadius: 12 }}>
                    <span style={numBadge}>A</span>
                    <span className="dex-ui-row-main">
                      <span className="dex-ui-row-title">{isDe ? 'Anrede' : 'Salutation'}</span>
                      <span className="dex-ui-row-sub">
                        {isDe
                          ? 'Standard-Feld: Pflicht-Dropdown (Frau / Herr / Divers / Keine Angabe) über dem Vornamen.'
                          : 'Standard field: required dropdown (Mrs / Mr / Diverse / Prefer not to say) above the first name.'}
                      </span>
                    </span>
                    <span className="dex-ui-row-actions">
                      <button
                        type="button"
                        className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                        onClick={() => setAskSalutation(false)}
                        title={isDe ? 'Anrede-Abfrage entfernen' : 'Remove salutation'}
                        aria-label={isDe ? 'Anrede-Abfrage entfernen' : 'Remove salutation'}
                      >
                        <X size={16} />
                      </button>
                    </span>
                  </div>
                )}
                {customFields.map((field, idx) => {
                  const isExpanded = !!fieldExpandOverride[field.id];
                  const isPeople = field.type === 'user' || field.type === 'roommate';
                  return (
                  <div
                    key={field.id}
                    className="dex-ui-card"
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
                      padding: '14px 16px',
                      marginBottom: 12,
                    }}
                  >
                    {/* v10.25: Konsolidierter Header — Label-Input prominent
                        als Titel + Typ-Dropdown rechts daneben + Pflicht-Pill
                        + Lösch-X. Reorder-Pfeile nur im Reorder-Modus.
                        v31.2: Griff-Symbol davor, damit man sieht, dass die
                        Karte ziehbar ist (draggable war sie schon). */}
                    <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
                      {reorderMode ? (
                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                          <button
                            type="button"
                            className="dex-ui-iconbtn"
                            onClick={() => moveCustomField(field.id, 'up')}
                            disabled={idx === 0}
                            style={{ width: 24, height: 20 }}
                            title={isDe ? 'Nach oben' : 'Move up'}
                            aria-label={isDe ? 'Nach oben' : 'Move up'}
                          ><ChevronUp size={14} /></button>
                          <button
                            type="button"
                            className="dex-ui-iconbtn"
                            onClick={() => moveCustomField(field.id, 'down')}
                            disabled={idx === customFields.length - 1}
                            style={{ width: 24, height: 20 }}
                            title={isDe ? 'Nach unten' : 'Move down'}
                            aria-label={isDe ? 'Nach unten' : 'Move down'}
                          ><ChevronDown size={14} /></button>
                        </div>
                      ) : customFields.length > 1 && (
                        <span className="dex-ui-drag-handle" title={isDe ? 'Ziehen, um die Reihenfolge zu ändern' : 'Drag to reorder'} aria-hidden="true">≡</span>
                      )}
                      <span style={numBadge}>{idx + 1}</span>
                      {/* v18.56: Textarea statt Input — lange Fragen brechen jetzt
                          um statt abgeschnitten zu werden. Auto-Höhe via ref
                          (height = scrollHeight). resize:none + overflow:hidden,
                          damit es wie ein wachsendes Eingabefeld wirkt. */}
                      <textarea
                        className="form-input"
                        value={field.label}
                        rows={1}
                        placeholder={isDe ? 'Wie lautet die Frage? (z.B. „Welche T-Shirt-Größe brauchst du?")' : 'What is the question? (e.g. „Which T-shirt size do you need?")'}
                        onChange={e => updateCustomField(field.id, { label: e.target.value })}
                        ref={(el) => { if (el) { el.style.height = 'auto'; el.style.height = `${el.scrollHeight}px`; } }}
                        style={{
                          flex: '1 1 260px', minWidth: 180,
                          // v22.30: minHeight 0 hebt die 48px-Mindesthöhe der
                          // .form-input-Klasse auf — die Auto-Höhe (scrollHeight)
                          // umschließt den Text dann exakt.
                          minHeight: 0,
                          fontSize: '0.95rem', fontWeight: 600,
                          padding: '10px 12px',
                          resize: 'none', overflow: 'hidden', lineHeight: 1.35,
                          fontFamily: 'inherit',
                          color: field.label ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
                        }}
                      />
                      {/* v11.4: feste Breite, damit Frage + Typ + Pflicht + X in
                          einer Zeile bleiben — sonst drückte ein langer Typ-Text
                          das X in die zweite Zeile. */}
                      <select
                        className="dex-ui-select"
                        value={field.type}
                        onChange={e => updateCustomField(field.id, { type: e.target.value as CustomFieldInput['type'] })}
                        title={isDe ? 'Art der Antwort' : 'Answer type'}
                        style={typeSelectStyle}
                      >
                        {MAIN_TYPES.map(ty => <option key={ty} value={ty}>{typeLabel(ty)}</option>)}
                      </select>
                      <label
                        className={cx('dex-ui-chip', field.required && 'is-active')}
                        title={isDe ? 'Pflicht: ohne Antwort lässt sich die Anmeldung nicht absenden' : 'Required: the registration cannot be submitted without an answer'}
                      >
                        <input
                          type="checkbox"
                          checked={field.required}
                          onChange={e => updateCustomField(field.id, { required: e.target.checked })}
                          style={{ display: 'none' }}
                        />
                        {field.required && <Check size={12} />}
                        {t('create.required')}
                      </label>
                      <button
                        type="button"
                        className="dex-ui-textbtn"
                        onClick={() => toggleFieldExpand(field.id, isExpanded)}
                        title={isExpanded ? (isDe ? 'Details einklappen' : 'Collapse details') : (isDe ? 'Details bearbeiten' : 'Edit details')}
                        aria-expanded={isExpanded}
                        style={{ marginLeft: 'auto' }}
                      >
                        {isExpanded ? (isDe ? 'Weniger' : 'Less') : (isDe ? 'Details' : 'Details')}
                        {isExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                      </button>
                      <button
                        className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                        onClick={() => removeCustomField(field.id)}
                        title={isDe ? 'Frage löschen' : 'Delete question'}
                        aria-label={isDe ? 'Frage löschen' : 'Delete question'}
                      >
                        <X size={16} />
                      </button>
                    </div>

                    {/* v24.25: Datum-Feld → optional auch die Uhrzeit abfragen
                        (immer sichtbar, nicht in „Details" versteckt). */}
                    {field.type === 'date' && (
                      <label className="dex-ui-inline" style={{ marginTop: 8, marginLeft: 36, cursor: 'pointer', fontSize: '0.82rem', color: 'var(--dex-gray-700)' }}>
                        <input
                          type="checkbox"
                          checked={!!field.withTime}
                          onChange={e => updateCustomField(field.id, { withTime: e.target.checked })}
                          style={{ accentColor: 'var(--dex-green, #86bc25)' }}
                        />
                        {isDe ? 'Auch die Uhrzeit abfragen' : 'Also ask for the time'}
                      </label>
                    )}
                    {/* v28.63: Übernachtungs-Zeitraum — buchbares Fenster und Nächte-Limit.
                        Ohne Fenster kann der Teilnehmer jedes Datum wählen; mit Fenster
                        (z.B. 22.09.–26.09.) bleibt die Auswahl an eurem Kontingent. */}
                    {field.type === 'daterange' && (
                      <div className="dex-ui-inline" style={{ marginLeft: 36, marginTop: 8, alignItems: 'flex-end', gap: 10 }}>
                        <div>
                          <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 3 }}>{isDe ? 'Buchbar ab' : 'Bookable from'}</div>
                          <input type="date" className="dex-ui-input dex-ui-input--sm"
                            style={{ width: 160 }}
                            value={field.rangeStart || ''}
                            onChange={e => updateCustomField(field.id, { rangeStart: e.target.value })} />
                        </div>
                        <div>
                          <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 3 }}>{isDe ? 'Buchbar bis' : 'Bookable until'}</div>
                          <input type="date" className="dex-ui-input dex-ui-input--sm"
                            style={{ width: 160 }}
                            value={field.rangeEnd || ''}
                            onChange={e => updateCustomField(field.id, { rangeEnd: e.target.value })} />
                        </div>
                        <div>
                          <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 3 }}>{isDe ? 'Max. Nächte' : 'Max. nights'}</div>
                          <input type="number" min={0} className="dex-ui-input dex-ui-input--sm"
                            style={{ width: 100 }}
                            placeholder={isDe ? 'offen' : 'open'}
                            value={field.maxNights || ''}
                            onChange={e => updateCustomField(field.id, { maxNights: parseInt(e.target.value, 10) || 0 })} />
                        </div>
                        <div className="dex-ui-help" style={{ flex: '1 1 220px', marginTop: 0, paddingBottom: 6 }}>
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

                    {/* v31.2: Reihenfolge im Editor = Reihenfolge, in der der
                        Teilnehmer die Frage erlebt: Fragetext (EN, Checkbox-Text)
                        → Antwortmöglichkeiten → was der Typ auslöst (Mails,
                        Suchkreis) → wann die Frage erscheint (Gruppe, Bedingung)
                        → Hilfetext. Vorher standen die Mail-Schalter oben und
                        die Optionen unter der Beschreibung. */}
                    {isExpanded && (
                    <div className="dex-ui-stack" style={{ marginLeft: 36, marginTop: 12 }}>
                    {/* v17.20: EN-Feld-Name — sichtbar wenn der Bilingual-
                        Toggle aktiviert wurde. */}
                    {bilingualFields && (
                      <div className="dex-ui-inline" style={{ flexWrap: 'nowrap' }}>
                        {enBadge}
                        <input
                          className="dex-ui-input dex-ui-input--sm"
                          value={field.labelEn || ''}
                          placeholder={isDe
                            ? 'Englischer Fragetext (optional — leer = fällt auf den deutschen Text zurück)'
                            : 'English question text (optional — empty = falls back to the German text)'}
                          onChange={e => updateCustomField(field.id, { labelEn: e.target.value })}
                          style={{ flex: 1, width: 'auto' }}
                        />
                      </div>
                    )}
                    {/* v11.94: Bei Checkbox-Feldern kann der Organizer den
                        Text neben der Checkbox individuell setzen — Default
                        ist „Ja, bestätigen" / „Yes, confirm". */}
                    {field.type === 'checkbox' && (
                      <div className="dex-ui-stack" style={{ gap: 6 }}>
                        <div className="dex-ui-label" style={{ marginBottom: 0 }}>
                          {isDe ? 'Text neben der Checkbox' : 'Text next to the checkbox'}
                          <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                        </div>
                        <input
                          className="dex-ui-input dex-ui-input--sm"
                          placeholder={isDe
                            ? 'z.B. „Ich habe die Teilnahmebedingungen gelesen" — leer = „Ja, bestätigen"'
                            : 'e.g. „I have read the terms" — empty = „Yes, confirm"'}
                          value={field.confirmLabel || ''}
                          onChange={e => updateCustomField(field.id, { confirmLabel: e.target.value })}
                        />
                        {/* v17.20: EN-Variante des Checkbox-Bestätigungstexts. */}
                        {bilingualFields && (
                          <div className="dex-ui-inline" style={{ flexWrap: 'nowrap' }}>
                            {enBadge}
                            <input
                              className="dex-ui-input dex-ui-input--sm"
                              value={field.confirmLabelEn || ''}
                              placeholder={isDe
                                ? 'Englischer Text neben der Checkbox (optional, Default: „Yes, confirm")'
                                : 'English text next to the checkbox (optional, default: „Yes, confirm")'}
                              onChange={e => updateCustomField(field.id, { confirmLabelEn: e.target.value })}
                              style={{ flex: 1, width: 'auto' }}
                            />
                          </div>
                        )}
                      </div>
                    )}

                    {/* v10.23: Dropdown-Optionen als gelisteter Editor mit
                        eigener Box. Mehrfachauswahl-Toggle direkt in diesem
                        Block (Kontext: betrifft nur die Optionsliste).
                        Nur sichtbar wenn type === 'select'. */}
                    {field.type === 'select' && (
                      <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '12px 14px' }}>
                        <div className="dex-ui-inline" style={{ justifyContent: 'space-between', marginBottom: 8 }}>
                          <span className="dex-ui-label" style={{ marginBottom: 0 }}>
                            {isDe ? 'Antwortmöglichkeiten' : 'Answer options'}
                          </span>
                          <span className="dex-ui-inline">
                            <label
                              className={cx('dex-ui-chip', !!field.multi && 'is-active')}
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
                              {field.multi && <Check size={12} />}
                              {isDe ? 'Mehrfachauswahl möglich' : 'Allow multiple selection'}
                            </label>
                            {/* v26.75: Vorfilter — nur bei Single-Select. Aktiviert
                                pro Option ein Kategorie-Feld; die Anmeldeseite zeigt
                                dann zuerst ein Kategorie-Dropdown und filtert die
                                Optionsliste darauf (z.B. „Herren"/„Damen" → Größen). */}
                            {!field.multi && (
                              <label
                                className={cx('dex-ui-chip', !!field.optionCategories && 'is-active')}
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
                                {field.optionCategories && <Check size={12} />}
                                {isDe ? 'Vorfilter nach Kategorie' : 'Pre-filter by category'}
                              </label>
                            )}
                          </span>
                        </div>
                        {/* v26.75: Beschriftung des Vorfilter-Dropdowns. */}
                        {!field.multi && field.optionCategories && (
                          <div className="dex-ui-field" style={{ marginBottom: 10 }}>
                            <label className="dex-ui-label">
                              {isDe ? 'Wie heißt der Vorfilter für Teilnehmer?' : 'What is the pre-filter called for attendees?'}
                            </label>
                            <input
                              className="dex-ui-input dex-ui-input--sm"
                              value={field.prefilterLabel || ''}
                              onChange={e => updateCustomField(field.id, { prefilterLabel: e.target.value })}
                              placeholder={isDe ? 'z.B. Größentabelle, Kategorie' : 'e.g. size chart, category'}
                              maxLength={40}
                            />
                          </div>
                        )}
                        <div className="dex-ui-stack" style={{ gap: 6 }}>
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
                            const optionRows = (items: Array<{ opt: string; optEn: string }>, onOpt: (ii: number, v: string) => void, onOptEn: (ii: number, v: string) => void, onRemove: (ii: number) => void): React.ReactNode => (
                              items.map((it, ii) => (
                                <div key={ii} className="dex-ui-stack" style={{ gap: 4 }}>
                                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                    <span className="dex-ui-muted" style={{ flexShrink: 0, width: 20, textAlign: 'right' }}>{ii + 1}.</span>
                                    <input className="dex-ui-input dex-ui-input--sm" value={it.opt} placeholder={isDe ? 'Auswahl (z.B. S)' : 'Choice (e.g. S)'} onChange={e => onOpt(ii, e.target.value)} style={{ flex: 1, width: 'auto' }} />
                                    <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => onRemove(ii)} title={isDe ? 'Entfernen' : 'Remove'} aria-label={isDe ? 'Entfernen' : 'Remove'}><X size={14} /></button>
                                  </div>
                                  {bilingualFields && (
                                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginLeft: 28 }}>
                                      {enBadge}
                                      <input className="dex-ui-input dex-ui-input--sm" value={it.optEn} placeholder={isDe ? 'Englische Variante (optional)' : 'English variant (optional)'} onChange={e => onOptEn(ii, e.target.value)} style={{ flex: 1, width: 'auto' }} />
                                    </div>
                                  )}
                                </div>
                              ))
                            );
                            return (
                              <div className="dex-ui-stack" style={{ gap: 10 }}>
                                {named.map((g, gi) => (
                                  <div key={gi} className="dex-ui-card dex-ui-stack" style={{ padding: '10px 12px', gap: 6, background: 'rgba(134,188,37,0.04)' }}>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                      {catBadge}
                                      <input className="dex-ui-input dex-ui-input--sm" value={g.category} placeholder={isDe ? 'Kategorie (z.B. Männergrößen)' : 'Category (e.g. men’s sizes)'} onChange={e => apply(named.map((x, i) => i === gi ? { ...x, category: e.target.value } : x), noCat)} style={{ flex: 1, width: 'auto', fontWeight: 600 }} />
                                      <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => apply(named.filter((_, i) => i !== gi), noCat)} title={isDe ? 'Kategorie entfernen' : 'Remove category'} aria-label={isDe ? 'Kategorie entfernen' : 'Remove category'}><X size={14} /></button>
                                    </div>
                                    {optionRows(
                                      g.items,
                                      (ii, v) => apply(named.map((x, i) => i === gi ? { ...x, items: x.items.map((y, j) => j === ii ? { ...y, opt: v } : y) } : x), noCat),
                                      (ii, v) => apply(named.map((x, i) => i === gi ? { ...x, items: x.items.map((y, j) => j === ii ? { ...y, optEn: v } : y) } : x), noCat),
                                      (ii) => apply(named.map((x, i) => i === gi ? { ...x, items: x.items.filter((_, j) => j !== ii) } : x), noCat),
                                    )}
                                    <button type="button" className="dex-ui-textbtn" onClick={() => apply(named.map((x, i) => i === gi ? { ...x, items: [...x.items, { opt: '', optEn: '' }] } : x), noCat)} style={{ alignSelf: 'flex-start' }}><Plus size={12} /> {isDe ? 'Auswahl hinzufügen' : 'Add choice'}</button>
                                  </div>
                                ))}
                                <div className="dex-ui-stack" style={{ border: '1px dashed var(--dex-gray-300)', borderRadius: 12, padding: '10px 12px', gap: 6 }}>
                                  <div className="dex-ui-muted" style={{ fontWeight: 600 }}>
                                    {isDe ? 'Ohne Kategorie — immer sichtbar (z.B. „T-Shirt bereits vorhanden")' : 'No category — always shown (e.g. „already have a shirt")'}
                                  </div>
                                  {optionRows(
                                    noCat,
                                    (ii, v) => apply(named, noCat.map((y, j) => j === ii ? { ...y, opt: v } : y)),
                                    (ii, v) => apply(named, noCat.map((y, j) => j === ii ? { ...y, optEn: v } : y)),
                                    (ii) => apply(named, noCat.filter((_, j) => j !== ii)),
                                  )}
                                  <button type="button" className="dex-ui-textbtn" onClick={() => apply(named, [...noCat, { opt: '', optEn: '' }])} style={{ alignSelf: 'flex-start' }}><Plus size={12} /> {isDe ? 'Option ohne Kategorie' : 'Option without category'}</button>
                                </div>
                                <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => apply([...named, { category: isDe ? `Kategorie ${named.length + 1}` : `Category ${named.length + 1}`, items: [{ opt: '', optEn: '' }] }], noCat)} style={{ alignSelf: 'flex-start' }}><Plus size={14} /> {isDe ? 'Kategorie hinzufügen' : 'Add category'}</button>
                              </div>
                            );
                          })() : <>
                          {(field.options || []).map((opt, optIdx) => (
                            <div key={optIdx} className="dex-ui-stack" style={{ gap: 4 }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                <span className="dex-ui-muted" style={{ flexShrink: 0, width: 20, textAlign: 'right', fontWeight: 600 }}>
                                  {optIdx + 1}.
                                </span>
                                <input
                                  className="dex-ui-input dex-ui-input--sm"
                                  value={opt}
                                  placeholder={isDe ? `Option ${optIdx + 1}` : `Option ${optIdx + 1}`}
                                  onChange={e => {
                                    const opts = [...(field.options || [])];
                                    opts[optIdx] = e.target.value;
                                    updateCustomField(field.id, { options: opts });
                                  }}
                                  style={{ flex: 1, width: 'auto' }}
                                />
                                <button
                                  type="button"
                                  className="dex-ui-iconbtn dex-ui-iconbtn--danger"
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
                                  aria-label={isDe ? 'Option entfernen' : 'Remove option'}
                                >
                                  <X size={14} />
                                </button>
                              </div>
                              {/* v17.20: Positional gemappte EN-Option. */}
                              {bilingualFields && (
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginLeft: 28 }}>
                                  {enBadge}
                                  <input
                                    className="dex-ui-input dex-ui-input--sm"
                                    value={(field.optionsEn || [])[optIdx] || ''}
                                    placeholder={isDe ? 'Englische Variante (optional)' : 'English variant (optional)'}
                                    onChange={e => {
                                      const optsEn = [...(field.optionsEn || [])];
                                      while (optsEn.length <= optIdx) optsEn.push('');
                                      optsEn[optIdx] = e.target.value;
                                      updateCustomField(field.id, { optionsEn: optsEn });
                                    }}
                                    style={{ flex: 1, width: 'auto' }}
                                  />
                                </div>
                              )}
                              {/* v26.75: Kategorie pro Option (nur wenn Vorfilter aktiv). */}
                              {field.optionCategories && (
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginLeft: 28 }}>
                                  {catBadge}
                                  <input
                                    className="dex-ui-input dex-ui-input--sm"
                                    value={(field.optionCategories || [])[optIdx] || ''}
                                    placeholder={isDe ? 'Kategorie (z.B. Herren) — leer = immer sichtbar' : 'Category (e.g. men) — empty = always shown'}
                                    onChange={e => {
                                      const cats = [...(field.optionCategories || [])];
                                      while (cats.length <= optIdx) cats.push('');
                                      cats[optIdx] = e.target.value;
                                      updateCustomField(field.id, { optionCategories: cats });
                                    }}
                                    style={{ flex: 1, width: 'auto' }}
                                  />
                                </div>
                              )}
                            </div>
                          ))}
                          <button
                            type="button"
                            className="dex-ui-textbtn"
                            onClick={() => updateCustomField(field.id, { options: [...(field.options || []), ''], ...(field.optionCategories ? { optionCategories: [...field.optionCategories, ''] } : {}) })}
                            style={{ alignSelf: 'flex-start' }}
                          >
                            <Plus size={12} /> {isDe ? 'Option hinzufügen' : 'Add option'}
                          </button>
                          </>}
                          {/* v26.74: Vorauswahl (nur Single-Select) — optional
                              eine Option, die im Anmeldeformular vorausgewählt ist. */}
                          {!field.multi && (field.options || []).filter(o => (o || '').trim()).length > 0 && (
                            <div className="dex-ui-field" style={{ marginTop: 10, paddingTop: 10, borderTop: '1px solid var(--dex-gray-200)' }}>
                              <label className="dex-ui-label">
                                {isDe ? 'Soll eine Antwort vorausgewählt sein?' : 'Should one answer be pre-selected?'}
                                <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                              </label>
                              <select
                                className="dex-ui-select"
                                value={field.defaultValue || ''}
                                onChange={e => updateCustomField(field.id, { defaultValue: e.target.value })}
                                style={{ maxWidth: 360 }}
                              >
                                <option value="">{isDe ? '— Keine Vorauswahl („Bitte wählen") —' : '— No pre-selection („Please choose") —'}</option>
                                {(field.options || []).map(o => (o || '').trim()).filter(Boolean).map((o, i) => (
                                  <option key={i} value={o}>{o}</option>
                                ))}
                              </select>
                              <div className="dex-ui-help">
                                {isDe ? 'Leer = der Teilnehmer muss selbst wählen.' : 'Empty = the attendee must choose themselves.'}
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    )}

                    {/* v26.60: Roommate-Feld → separate Zimmerpartner-Anfrage-Mail
                        an die ausgewählte Person abschaltbar (Default: an). Diese
                        Mail ist UNABHÄNGIG vom CC-Schalter darunter — das war der
                        gemeldete Bug: „CC ausgestellt, Mails kommen trotzdem". */}
                    {field.type === 'roommate' && (
                      <label className={cx('dex-ui-toggle-row', field.notifyRoommate !== false && 'is-active')}>
                        <input
                          type="checkbox"
                          checked={field.notifyRoommate !== false}
                          onChange={e => updateCustomField(field.id, { notifyRoommate: e.target.checked ? undefined : false })}
                        />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">
                            {isDe ? 'Ausgewählte Person automatisch benachrichtigen' : 'Automatically notify the selected person'}
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
                          <span className="dex-ui-toggle-row-desc">
                            {isDe
                              ? 'Sie bekommt direkt nach der Anmeldung eine eigene Zimmerpartner-Anfrage-Mail — unabhängig vom CC-Schalter.'
                              : 'They receive a dedicated roommate request email right after the registration — independent of the CC toggle.'}
                          </span>
                        </span>
                      </label>
                    )}
                    {/* v29.40: Personen-Feld → Suche auf den Verteilerkreis des
                        Events begrenzen. Anlass: Zimmerpartner liessen sich
                        ausserhalb des Verteilers auswaehlen. */}
                    {isPeople && (
                      <label className={cx('dex-ui-toggle-row', !!field.audienceOnly && 'is-active')}>
                        <input
                          type="checkbox"
                          checked={!!field.audienceOnly}
                          onChange={e => updateCustomField(field.id, { audienceOnly: e.target.checked ? true : undefined })}
                        />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">
                            {isDe ? 'Nur Personen zulassen, die dieses Event sehen können' : 'Only allow people who can see this event'}
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
                          <span className="dex-ui-toggle-row-desc">
                            {isDe
                              ? 'Die Suche findet dann nur den Verteilerkreis des Events — niemanden, der nicht eingeladen ist.'
                              : 'The search then only finds the event audience — nobody who is not invited.'}
                          </span>
                        </span>
                      </label>
                    )}
                    {/* v18.41: People-Picker-Feld → ausgewählte Person bei
                        An-/Abmelde-Mail auf CC (nur für user/roommate-Felder). */}
                    {isPeople && (
                      <label className={cx('dex-ui-toggle-row', !!field.ccOnEmails && 'is-active')}>
                        <input
                          type="checkbox"
                          checked={!!field.ccOnEmails}
                          onChange={e => updateCustomField(field.id, { ccOnEmails: e.target.checked })}
                        />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">
                            {isDe ? 'Ausgewählte Person bei An-/Abmelde-Mail auf CC setzen' : 'CC the selected person on registration/cancellation email'}
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
                          </span>
                          <span className="dex-ui-toggle-row-desc">
                            {isDe
                              ? 'Nur die Mails — der Outlook-Termin wird nicht an diese Person gesendet.'
                              : 'Emails only — the Outlook event is not sent to this person.'}
                          </span>
                        </span>
                      </label>
                    )}
                    {/* v10.23: Roommate-Erklärung — wird IMMER bei type=roommate
                        angezeigt. v31.2: eine sichtbare Zeile, der Rest im
                        InfoTooltip; der Tipp zum fehlenden Zimmerart-Feld bleibt
                        ein eigener Warn-Kasten. */}
                    {field.type === 'roommate' && (() => {
                      const roomKeywords = ['einzelzimmer', 'doppelzimmer', 'single room', 'double room', 'zimmerart', 'room type'];
                      const hasRoomTypeField = customFields.some(other => {
                        if (other.id === field.id) return false;
                        const lbl = (other.label || '').toLowerCase();
                        const opts = (other.options || []).join(' ').toLowerCase();
                        return roomKeywords.some(k => lbl.indexOf(k) >= 0 || opts.indexOf(k) >= 0);
                      });
                      return (
                        <div className="dex-ui-stack" style={{ gap: 8 }}>
                          <div className="dex-ui-callout dex-ui-callout--info">
                            <span className="dex-ui-callout-icon"><Info size={16} /></span>
                            <span>
                              <strong>{isDe ? 'So funktioniert das Roommate-Feld:' : 'How the roommate field works:'}</strong>{' '}
                              {isDe
                                ? 'Der Teilnehmer sucht seinen Wunsch-Zimmerpartner per Personensuche; die Person bekommt eine Anfrage-Mail. Wählen sich zwei gegenseitig, zeigt das Admin Center ein bestätigtes Match.'
                                : 'The attendee picks their preferred roommate via people search; that person gets a request email. If two pick each other, the admin center shows a confirmed match.'}
                              <InfoTooltip text={isDe ? (
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
                                </>
                              )} />
                            </span>
                          </div>
                          {!hasRoomTypeField && (
                            <div className="dex-ui-callout dex-ui-callout--warn">
                              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                              <span>
                                {isDe
                                  ? <><strong>Tipp:</strong> aktuell hast du noch kein Zimmerart-Feld angelegt. Ein Dropdown &bdquo;Zimmerart&ldquo; mit &bdquo;Einzelzimmer / Doppelzimmer&ldquo; wäre eine sinnvolle Ergänzung — sonst können auch Teilnehmer ohne Doppelzimmer-Wunsch einen Roommate angeben.</>
                                  : <><strong>Tip:</strong> you don&apos;t have a room-type field yet. A dropdown &ldquo;Room type&rdquo; with &ldquo;Single / Double&rdquo; would be a useful addition — otherwise attendees without a double-room wish can still pick a roommate.</>}
                              </span>
                            </div>
                          )}
                        </div>
                      );
                    })()}

                    {/* v10.24: Pro-Gruppe-Sichtbarkeit — nur sichtbar wenn die
                        Split-Capacity in Schritt 3 aktiv ist. Der Organizer
                        kann ein Feld auf Gruppe A oder Gruppe B beschränken
                        (Beispiel: Pflicht-Checkbox „Leistungsnachweis vorhanden"
                        nur für Durchstarter / Gruppe A). 'all' = beide
                        Gruppen sehen das Feld (Default). */}
                    {useSplitCapacities && (() => {
                      const labelA = (splitLabelA || '').trim() || 'Durchstarter';
                      const labelB = (splitLabelB || '').trim() || 'Funstarter';
                      const current = field.onlyForGroup || 'all';
                      return (
                        <div>
                          <div className="dex-ui-label" style={{ marginBottom: 6 }}>
                            {isDe ? 'Wer bekommt diese Frage?' : 'Who gets this question?'}
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
                          <div className="dex-ui-inline">
                            {([
                              { v: 'all', text: isDe ? 'Beide Gruppen' : 'Both groups' },
                              { v: 'A', text: isDe ? `Nur ${labelA}` : `${labelA} only` },
                              { v: 'B', text: isDe ? `Nur ${labelB}` : `${labelB} only` },
                            ] as const).map(opt => (
                              <label key={opt.v} className={cx('dex-ui-chip', current === opt.v && 'is-active')}>
                                <input
                                  type="radio"
                                  name={`onlyForGroup-${field.id}`}
                                  checked={current === opt.v}
                                  onChange={() => updateCustomField(field.id, { onlyForGroup: opt.v as 'all' | 'A' | 'B' })}
                                  style={{ display: 'none' }}
                                />
                                {current === opt.v && <Check size={12} />}
                                {opt.text}
                              </label>
                            ))}
                          </div>
                        </div>
                      );
                    })()}
                    {/* v7.21: Sichtbarkeitsbedingung — Feld nur anzeigen wenn
                        eine andere Frage einen bestimmten Wert hat. Quelle
                        kann nur ein Feld VOR diesem hier sein (idx < aktuell)
                        und muss vom Typ select oder checkbox sein. */}
                    {renderShowIfConfig(field, idx, customFields, (u) => updateCustomField(field.id, u))}

                    {/* v7.20: Beschreibung pro Feld. v18.18: Darstellung
                        wählbar — „i"-Box neben dem Label ODER Erklär-Text
                        unter dem Label. */}
                    <div>
                      <div className="dex-ui-label" style={{ marginBottom: 6 }}>
                        {isDe ? 'Hilfetext für Teilnehmer' : 'Help text for attendees'}
                        <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                      </div>
                      {/* v27.4: Kompakter Editor mit dauerhaft sichtbarer Leiste
                          (Fett + Link) statt der pnp-RichText-Bubble. */}
                      <FieldDescEditor
                        value={field.helpText || ''}
                        onChange={text => updateCustomField(field.id, { helpText: text })}
                        isDe={isDe}
                      />
                      {field.helpText && field.helpText.trim() && (
                        <div className="dex-ui-inline" style={{ marginTop: 8 }}>
                          <span className="dex-ui-muted" style={{ fontWeight: 600 }}>{isDe ? 'Wo erscheint er?' : 'Where does it show?'}</span>
                          <label className={cx('dex-ui-chip', (field.helpTextStyle || 'tooltip') !== 'inline' && 'is-active')}>
                            <input
                              type="radio"
                              name={`helpStyle-${field.id}`}
                              checked={(field.helpTextStyle || 'tooltip') !== 'inline'}
                              onChange={() => updateCustomField(field.id, { helpTextStyle: 'tooltip' })}
                              style={{ display: 'none' }}
                            />
                            {isDe ? 'Als „i"-Info-Box (Hover)' : 'As „i" info box (hover)'}
                          </label>
                          <label className={cx('dex-ui-chip', field.helpTextStyle === 'inline' && 'is-active')}>
                            <input
                              type="radio"
                              name={`helpStyle-${field.id}`}
                              checked={field.helpTextStyle === 'inline'}
                              onChange={() => updateCustomField(field.id, { helpTextStyle: 'inline' })}
                              style={{ display: 'none' }}
                            />
                            {isDe ? 'Als Text unter der Frage' : 'As text below the question'}
                          </label>
                        </div>
                      )}
                      {/* v17.20: EN-Variante der Beschreibung. */}
                      {bilingualFields && (
                        <div className="dex-ui-inline" style={{ flexWrap: 'nowrap', marginTop: 6 }}>
                          {enBadge}
                          <input
                            className="dex-ui-input dex-ui-input--sm"
                            value={field.helpTextEn || ''}
                            placeholder={isDe
                              ? 'Englische Beschreibung (optional)'
                              : 'English description (optional)'}
                            onChange={e => updateCustomField(field.id, { helpTextEn: e.target.value })}
                            style={{ flex: 1, width: 'auto' }}
                          />
                        </div>
                      )}
                    </div>
                    </div>
                    )}
                  </div>
                  );
                })}
              </div>

              {/* v22.38/v31.2: Die generellen Formular-Optionen (Zweisprachigkeit,
                  Formular-Sprache) stehen jetzt HINTER der Feld-Liste — sie
                  verfeinern das Formular, die Fragen sind das Eigentliche. Der
                  frühere „Anrede abfragen?"-Toggle ist als vorgeschlagenes Feld
                  im Katalog (Sonder-Key 'salutation'). */}
              <div className="dex-ui-section">
                <h3 className="dex-ui-section-title">{isDe ? 'Sprache des Formulars' : 'Form language'}</h3>
                <div className="dex-ui-stack" style={{ gap: 12 }}>
                  <label className={cx('dex-ui-toggle-row', bilingualFields && 'is-active')}>
                    <input
                      type="checkbox"
                      checked={bilingualFields}
                      onChange={e => setBilingualFields(e.target.checked)}
                    />
                    <span className="dex-ui-toggle-row-body">
                      <span className="dex-ui-toggle-row-title">
                        {isDe ? 'Fragen zweisprachig anlegen (Deutsch und Englisch)' : 'Write the questions bilingually (German and English)'}
                        <InfoTooltip text={isDe
                          ? <>
                              <strong>Was du hier einstellst:</strong> ob du pro Custom-Field, das du oben anlegst, <strong>eine englische Variante</strong> der Texte hinterlegen kannst — also Feld-Name, Beschreibung (i-Tooltip), Checkbox-Bestätigungs-Text und Dropdown-Optionen jeweils auf Deutsch UND auf Englisch. Default: <strong>aus</strong>.<br /><br />
                              <strong>Anzeige in der App:</strong> wenn aktiviert, blendet jede Feld-Karte einen zweiten Eingabe-Block für die EN-Variante ein. Teilnehmer mit App-Sprache <strong>Englisch</strong> bekommen automatisch die EN-Texte zu sehen. Wer als App-Sprache Deutsch eingestellt hat, sieht weiterhin die DE-Texte. Zusätzlich folgt das Standard-Anmelde-Formular (Platzhalter, Hinweis-Boxen, Sub-Event-Sektion) ab dann der <strong>App-Spracheinstellung des Teilnehmers</strong> statt der Mail-Sprache des Events.<br /><br />
                              <strong>Auswirkung für Teilnehmer:</strong> internationale Kolleg:innen, die kein Deutsch sprechen, sehen das komplette Anmelde-Formular sauber auf Englisch. Wer als Organizer keine EN-Variante einträgt, fällt im EN-Modus still auf den DE-Wert zurück — die App bricht also nichts kaputt, falls du nur einige Felder übersetzt.
                            </>
                          : <>
                              <strong>What this controls:</strong> whether, for each custom field you create above, you can store <strong>an English variant</strong> of the texts — i.e. field name, description (i-tooltip), checkbox confirmation text and dropdown options in both German AND English. Default: <strong>off</strong>.<br /><br />
                              <strong>Where you see it:</strong> when enabled, each field card shows a second input row for the EN variant. Attendees with app language set to <strong>English</strong> automatically see the EN texts. Attendees with German keep seeing the DE texts. In addition, the standard registration form chrome (placeholders, hint boxes, sub-event section) follows the <strong>attendee&apos;s app language</strong> instead of the event&apos;s email language.<br /><br />
                              <strong>For attendees:</strong> international colleagues who do not speak German see the whole registration form cleanly in English. If an organizer leaves the EN variant empty for some field, the app silently falls back to the DE value — nothing breaks if you only translate a subset of fields.
                            </>
                        } />
                      </span>
                      <span className="dex-ui-toggle-row-desc">
                        {isDe
                          ? <>Standard: aus. Eingeschaltet bekommt jede Frage einen zweiten Eintrag für den englischen Text (Frage, Beschreibung, Optionen); Teilnehmer sehen automatisch ihre Sprache. Stellst du deine Fragen nur <strong>einsprachig</strong> (z.&nbsp;B. nur Englisch), lass das aus und lege unten die Formular-Sprache fest.</>
                          : <>Default: off. When on, every question gets a second entry for the English text (question, description, options); attendees automatically see their language. If your questions are <strong>monolingual</strong> (e.g. English only), leave this off and fix the form language below instead.</>}
                      </span>
                    </span>
                  </label>

                  {/* v22.32: Sprache des Anmeldeformulars — gehört inhaltlich zu
                      den Formular-Optionen. v31.2: drei Chips statt Dropdown mit
                      drei Einträgen; dieselben Werte ('' / 'de' / 'en'). */}
                  <div className="dex-ui-field">
                    <div className="dex-ui-label">
                      {isDe ? 'In welcher Sprache erscheint das Anmeldeformular?' : 'Which language does the registration form appear in?'}
                      <InfoTooltip text={isDe
                        ? <>
                            <strong>Was du hier einstellst:</strong> in welcher Sprache die <strong>komplette Anmeldeseite</strong> (alle Texte, Buttons und der <strong>Datenschutz-Disclaimer</strong>) angezeigt wird.<br /><br />
                            <strong>Anzeige in der App:</strong> bei <strong>Automatisch</strong> folgt die Anmeldeseite der App-Sprache des Teilnehmers. Wählst du <strong>Immer Deutsch</strong> oder <strong>Immer Englisch</strong>, wird die Anmeldeseite <strong>fest in dieser Sprache</strong> angezeigt — auch wenn der Teilnehmer die App z.&nbsp;B. auf Deutsch nutzt. Ein kleiner Hinweis im Kopfbereich zeigt das an.<br /><br />
                            <strong>Unsere Empfehlung:</strong> Stellst du deine eigenen Fragen oben <strong>nur in einer Sprache</strong> (z.&nbsp;B. Englisch), dann stelle das Formular <strong>fest auf diese Sprache</strong>. Sonst mischen sich bei Teilnehmern mit anderer App-Sprache die deutschen Standard-Texte (Buttons, Hinweise, Datenschutz) mitten zwischen deine englischen Fragen — das wirkt unsauber.<br /><br />
                            <strong>Auswirkung für Teilnehmer:</strong> bei einem englischsprachigen Event sehen sie die Anmeldung samt Disclaimer komplett auf Englisch, egal welche App-Sprache eingestellt ist.
                          </>
                        : <>
                            <strong>What you set here:</strong> the language in which the <strong>entire registration page</strong> (all texts, buttons and the <strong>privacy disclaimer</strong>) is shown.<br /><br />
                            <strong>Where you see it:</strong> with <strong>Automatic</strong> the page follows the attendee&apos;s app language. Choosing <strong>Always German</strong> or <strong>Always English</strong> forces the registration page into that language — even if the attendee uses the app in another language. A small hint in the header indicates this.<br /><br />
                            <strong>Our recommendation:</strong> if your own questions above are written in <strong>one language only</strong> (e.g. English), <strong>fix the form to that language</strong>. Otherwise attendees with a different app language get the German standard texts (buttons, hints, privacy) mixed in between your English questions — which looks messy.<br /><br />
                            <strong>For attendees:</strong> for an English-language event they see the registration and disclaimer fully in English regardless of their app language.
                          </>
                      } />
                    </div>
                    <div className="dex-ui-inline">
                      {([
                        { v: '', text: isDe ? 'Automatisch (App-Sprache des Teilnehmers)' : 'Automatic (attendee\'s app language)' },
                        { v: 'de', text: isDe ? 'Immer Deutsch' : 'Always German' },
                        { v: 'en', text: isDe ? 'Immer Englisch' : 'Always English' },
                      ] as Array<{ v: '' | 'de' | 'en'; text: string }>).map(opt => (
                        <button
                          key={opt.v || 'auto'}
                          type="button"
                          className={cx('dex-ui-chip', registrationLanguage === opt.v && 'is-active')}
                          onClick={() => setRegistrationLanguage(opt.v)}
                        >
                          {registrationLanguage === opt.v && <Check size={12} />}
                          {opt.text}
                        </button>
                      ))}
                    </div>
                    <div className="dex-ui-help">
                      {isDe
                        ? <>Standard: Automatisch. <strong>Empfehlung:</strong> Stellst du deine Fragen nur in einer Sprache (z.&nbsp;B. Englisch), stelle das Formular fest auf diese Sprache — sonst stehen deutsche Standard-Texte mitten zwischen deinen englischen Fragen.</>
                        : <>Default: Automatic. <strong>Recommendation:</strong> if your questions are in one language only (e.g. English), fix the form to that language — otherwise German standard texts appear in between your English questions.</>}
                    </div>
                  </div>
                </div>
              </div>

              {/* v15.0/v31.2: Der frühere Bereich „Felder pro Sub-Event" stand
                  hier seit v15.0 mit display:none im Baum (ersetzt durch die
                  Sub-Event-Reiter oben). Er ist entfallen — 320 Zeilen toter
                  Editor-Code, den niemand mehr sehen konnte. */}
              </div>{/* v15.0: close activeFieldsTabIdx===0 wrapper */}

              {/* v18.75: Sicherheitshinweis vor dem Absenden — eigene Section
                  ganz unten in Schritt 5 (gilt event-weit, daher außerhalb der
                  Feld-Tabs). v31.2: Schalter-Zeile plus zwei Auswahl-Kacheln
                  statt Checkbox + Radio-Paar. */}
              <div className="dex-ui-section">
                <h3 className="dex-ui-section-title">{isDe ? 'Vor dem Absenden' : 'Before submitting'}</h3>
                <label className={cx('dex-ui-toggle-row', confirmDialogEnabled && 'is-active')}>
                  <input
                    type="checkbox"
                    checked={confirmDialogEnabled}
                    onChange={e => setConfirmDialogEnabled(e.target.checked)}
                  />
                  <span className="dex-ui-toggle-row-body">
                    <span className="dex-ui-toggle-row-title">
                      {isDe ? 'Teilnehmer bestätigen ihre Anmeldung noch einmal' : 'Attendees confirm their registration once more'}
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
                    </span>
                    <span className="dex-ui-toggle-row-desc">
                      {isDe
                        ? 'Standard: aus. Eingeschaltet erscheint nach „Anmelden" ein Dialog, den der Teilnehmer bestätigen muss — schützt vor versehentlichen Anmeldungen.'
                        : 'Default: off. When on, a dialog appears after „Register" that the attendee has to confirm — protects against accidental registrations.'}
                    </span>
                  </span>
                </label>
                {confirmDialogEnabled && (
                  <div className="dex-ui-fade-in" style={{ marginTop: 12 }}>
                    <div className="dex-ui-label">
                      {isDe ? 'Was soll der Dialog zeigen?' : 'What should the dialog show?'}
                    </div>
                    <div className="dex-ui-grid-2">
                      <button
                        type="button"
                        className={cx('dex-ui-choice', confirmDialogMode !== 'freetext' && 'is-active')}
                        onClick={() => setConfirmDialogMode('summary')}
                        aria-pressed={confirmDialogMode !== 'freetext'}
                      >
                        <span className="dex-ui-choice-body">
                          <span className="dex-ui-choice-title">{isDe ? 'Auswahl-Übersicht' : 'Selection summary'}</span>
                          <span className="dex-ui-choice-desc">
                            {isDe
                              ? 'Listet Haupt-Event und gewählte Sub-Events auf; der Teilnehmer kann vor dem Absenden einzelne Punkte noch ab- oder zuwählen.'
                              : 'Lists the main event and selected sub-events; the attendee can de-/select items before submitting.'}
                          </span>
                        </span>
                        <span className="dex-ui-choice-check">{confirmDialogMode !== 'freetext' && <Check size={12} />}</span>
                      </button>
                      <button
                        type="button"
                        className={cx('dex-ui-choice', confirmDialogMode === 'freetext' && 'is-active')}
                        onClick={() => setConfirmDialogMode('freetext')}
                        aria-pressed={confirmDialogMode === 'freetext'}
                      >
                        <span className="dex-ui-choice-body">
                          <span className="dex-ui-choice-title">{isDe ? 'Eigener Hinweistext' : 'Custom hint text'}</span>
                          <span className="dex-ui-choice-desc">
                            {isDe
                              ? 'Zeigt einen frei formulierten Hinweis (z.B. Verbindlichkeit, Storno-Frist), den der Teilnehmer bestätigen muss.'
                              : 'Shows a free-text note (e.g. binding registration, cancellation deadline) the attendee must acknowledge.'}
                          </span>
                        </span>
                        <span className="dex-ui-choice-check">{confirmDialogMode === 'freetext' && <Check size={12} />}</span>
                      </button>
                    </div>
                    {confirmDialogMode === 'freetext' && (
                      <textarea
                        className="dex-ui-textarea"
                        value={confirmDialogText}
                        onChange={e => setConfirmDialogText(e.target.value)}
                        rows={3}
                        placeholder={isDe
                          ? 'z.B. „Bitte beachte: Die Anmeldung ist verbindlich. Eine Stornierung ist nur bis 3 Tage vor dem Event möglich."'
                          : 'e.g. „Please note: registration is binding. Cancellation is only possible up to 3 days before the event."'}
                        style={{ marginTop: 10 }}
                      />
                    )}
                  </div>
                )}
              </div>

              </div>
  );
};
