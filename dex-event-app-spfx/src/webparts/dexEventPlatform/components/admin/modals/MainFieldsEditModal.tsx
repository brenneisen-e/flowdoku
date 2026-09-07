/* MainFieldsEditModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14488-14614 des
 * Stands vor dem Schnitt). Die Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.2: Optik nach docs/ui-leitfaden.md — Modal-Kopf/-Fuß statt eigenem <h3>
 * und Knopfzeile, dex-ui-Klassen statt Inline-Styles (Hover, Fokus-Ring),
 * Checkbox-Felder als Ja/Nein-Zeile, ausgeblendete showIf-Felder werden
 * gezählt und erklärt. Verhalten (Form-State, Speichern, Schließen) unverändert.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Pencil, AlertCircle, Info } from '../../Icons';
import { MultiSelectDropdown } from '../../MultiSelectDropdown';
import { InfoTooltip } from '../../InfoTooltip';
import { cx } from '../../dexUi';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import { FieldSelectInput, fieldVisibleByShowIf } from './FieldSelectInput';

export interface MainFieldsEditModalProps {
  closeMainFieldsEdit: () => void;
  isDe: boolean;
  mainFieldsEditError: string;
  mainFieldsEditForm: Record<string, string>;
  mainFieldsEditName: string;
  mainFieldsEditReg: SPRegistration;
  mainFieldsEditSaving: boolean;
  saveMainFieldsEdit: () => Promise<void>;
  selectedEvent: DeloitteEvent;
  setMainFieldsEditForm: React.Dispatch<React.SetStateAction<Record<string, string>>>;
}

export const MainFieldsEditModal: React.FC<MainFieldsEditModalProps> = (p) => {
  const { closeMainFieldsEdit, isDe, mainFieldsEditError, mainFieldsEditForm, mainFieldsEditName, mainFieldsEditReg, mainFieldsEditSaving, saveMainFieldsEdit, selectedEvent, setMainFieldsEditForm } = p;
  // v31.2: Feldliste einmal ableiten (keine Hooks, reine Rechnung). Die
  // showIf-Regel blendete Felder bis v31.1 stumm aus — der Organizer suchte
  // dann ein Feld, das „in der Tabelle doch steht". Deshalb zählen wir die
  // ausgeblendeten und sagen unter dem Raster, warum sie fehlen.
  const parentFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'document' && f.label && f.label.trim());
  // v30.95: Sichtbarkeitsregel + Kombibox mit Kategorien wie auf der
  // Anmeldeseite (s. FieldSelectInput).
  const visibleFields = parentFields.filter(cf => fieldVisibleByShowIf(cf, fid => mainFieldsEditForm[fid] || ''));
  const hiddenCount = parentFields.length - visibleFields.length;
  const requiredMark = (required?: boolean): React.ReactNode => required
    ? <span style={{ color: 'var(--dex-red, #da291c)' }} title={isDe ? 'Pflichtfeld' : 'Required'}> *</span>
    : <span className="dex-ui-label-optional">(optional)</span>;
  return (
        <Modal
          open={!!mainFieldsEditReg}
          onClose={() => { if (!mainFieldsEditSaving) closeMainFieldsEdit(); }}
          maxWidth={760}
          dismissable={!mainFieldsEditSaving}
          ariaLabel={isDe ? 'Hauptevent-Felder bearbeiten' : 'Edit main-event fields'}
          icon={<Pencil size={20} />}
          title={<>{isDe ? 'Antworten zum Hauptevent bearbeiten' : 'Edit main-event answers'}{' — '}<span style={{ color: 'var(--dex-green-dark, #6b9a1e)' }}>{mainFieldsEditName}</span></>}
          subtitle={isDe
            ? 'Die Antworten liegen auf der Anmeldung zum Hauptevent — nicht je Sub-Event.'
            : 'The answers live on the person’s main-event registration — not per sub-event.'}
          footer={<>
            <button type="button" className="btn btn-secondary" onClick={closeMainFieldsEdit} disabled={mainFieldsEditSaving}>
              {isDe ? 'Abbrechen' : 'Cancel'}
            </button>
            <button type="button" className="btn btn-primary" onClick={saveMainFieldsEdit} disabled={mainFieldsEditSaving}>
              {mainFieldsEditSaving ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Speichern' : 'Save')}
            </button>
          </>}
        >
          <div className="dex-ui-section">
            <div className="dex-ui-section-title">
              {isDe ? 'Felder des Hauptevents' : 'Main-event fields'}
              <InfoTooltip text={isDe
                ? 'Was du hier einstellst: die Antworten dieser Person auf die Felder des Hauptevents. Anzeige in der App: die hellblauen Spalten „Felder des Hauptevents“ in der Teilnehmerliste. Auswirkung: gespeichert auf der Anmeldung zum Hauptevent, nicht je Sub-Event.'
                : 'What you set here: this person’s answers to the main-event fields. Shown in the app: the light-blue “Main-event fields” columns of the participant list. Effect: stored on the main-event registration, not per sub-event.'} />
            </div>
            {parentFields.length === 0 ? (
              <div className="dex-ui-empty">
                <div className="dex-ui-empty-title">{isDe ? 'Keine bearbeitbaren Felder' : 'No editable fields'}</div>
                {isDe ? 'Dieses Hauptevent hat keine bearbeitbaren Felder.' : 'This main event has no editable fields.'}
              </div>
            ) : (
              <div className="dex-ui-grid-2">
                {visibleFields.map(cf => {
                  const value = mainFieldsEditForm[cf.id] || '';
                  const setVal = (v: string): void => setMainFieldsEditForm(prev => ({ ...prev, [cf.id]: v }));
                  const inputId = `mfe-${cf.id}`;
                  // htmlFor nur beim eigenen <input> — Kombibox und Mehrfachauswahl tragen keine Id.
                  const labelEl = (withFor?: boolean): React.ReactElement => <label className="dex-ui-label" htmlFor={withFor ? inputId : undefined}>{cf.label}{requiredMark(cf.required)}</label>;
                  if (cf.type === 'select' && !cf.multi && cf.options && cf.options.length > 0) {
                    return (
                      <div key={cf.id} className="dex-ui-field" style={{ marginBottom: 0 }}>
                        {labelEl()}
                        <FieldSelectInput field={cf} value={value} onChange={setVal} isDe={isDe} />
                      </div>
                    );
                  }
                  if (cf.type === 'select' && cf.multi && cf.options && cf.options.length > 0) {
                    const selected = value.split(' | ').map(s => s.trim()).filter(Boolean);
                    return (
                      <div key={cf.id} className="dex-ui-field" style={{ marginBottom: 0, gridColumn: '1 / -1' }}>
                        {labelEl()}
                        <MultiSelectDropdown
                          options={cf.options}
                          value={selected}
                          onChange={next => setVal(next.join(' | '))}
                          placeholder={isDe ? '— bitte wählen —' : '— please choose —'}
                        />
                      </div>
                    );
                  }
                  if (cf.type === 'checkbox') {
                    const isChecked = value === 'true' || value === '1';
                    // v31.2: Frage und Haken in EINER klickbaren Zeile (Leitfaden 2b,
                    // Ja/Nein) statt Beschriftung oben und loser Checkbox darunter.
                    return (
                      <label key={cf.id} className={cx('dex-ui-toggle-row', isChecked && 'is-active')}>
                        <input type="checkbox" checked={isChecked} onChange={e => setVal(e.target.checked ? 'true' : 'false')} />
                        <span className="dex-ui-toggle-row-body">
                          <span className="dex-ui-toggle-row-title">{cf.label}{requiredMark(cf.required)}</span>
                          <span className="dex-ui-toggle-row-desc">
                            {isChecked ? (isDe ? 'Ja — angehakt' : 'Yes — checked') : (isDe ? 'Nein — nicht angehakt' : 'No — not checked')}
                          </span>
                        </span>
                      </label>
                    );
                  }
                  return (
                    <div key={cf.id} className="dex-ui-field" style={{ marginBottom: 0 }}>
                      {labelEl(true)}
                      <input id={inputId} className="dex-ui-input" type={cf.type === 'number' ? 'number' : 'text'} value={value} onChange={e => setVal(e.target.value)} />
                    </div>
                  );
                })}
              </div>
            )}
            {hiddenCount > 0 && (
              <p className="dex-ui-help" style={{ marginTop: 10 }}>
                {isDe
                  ? `${hiddenCount} ${hiddenCount === 1 ? 'weiteres Feld erscheint' : 'weitere Felder erscheinen'} erst, wenn die Antwort passt, von der ${hiddenCount === 1 ? 'es' : 'sie'} abhängt.`
                  : `${hiddenCount} more ${hiddenCount === 1 ? 'field appears' : 'fields appear'} once the answer ${hiddenCount === 1 ? 'it depends' : 'they depend'} on matches.`}
              </p>
            )}
          </div>
          {mainFieldsEditError && (
            <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span>{mainFieldsEditError}</span>
            </div>
          )}
          <div className="dex-ui-callout dex-ui-callout--neutral">
            <span className="dex-ui-callout-icon"><Info size={16} /></span>
            <span>{isDe ? 'Nach dem Speichern steht jede Änderung mit deinem Namen und Datum im Änderungsprotokoll.' : 'After saving, every change is recorded in the audit log with your name and date.'}</span>
          </div>
        </Modal>
  );
};

