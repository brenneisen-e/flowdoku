/* MainFieldsEditModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14488-14614 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Pencil, X } from '../../Icons';
import { MultiSelectDropdown } from '../../MultiSelectDropdown';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

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
  return (
        <Modal
          open={!!mainFieldsEditReg}
          onClose={() => { if (!mainFieldsEditSaving) closeMainFieldsEdit(); }}
          maxWidth={760}
          dismissable={!mainFieldsEditSaving}
          ariaLabel={isDe ? 'Hauptevent-Felder bearbeiten' : 'Edit main-event fields'}
        >
          <div className="flex-between">
            <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Pencil size={18} />{' '}
              {isDe ? 'Felder des Hauptevents bearbeiten' : 'Edit main-event fields'}
              {' — '}
              <span style={{ color: 'var(--dex-green-dark)' }}>{mainFieldsEditName}</span>
            </h3>
            <button
              onClick={closeMainFieldsEdit}
              style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}
              aria-label={isDe ? 'Schließen' : 'Close'}
              disabled={mainFieldsEditSaving}
            ><X size={20} /></button>
          </div>
          <p style={{ margin: 0, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
            {isDe
              ? 'Du bearbeitest hier die Antworten auf die Felder des Hauptevents (die hellblauen Spalten „Felder des Hauptevents"). Sie werden in der Registrierung der Person auf dem Hauptevent gespeichert — nicht pro Sub-Event. Jede Änderung wird im Änderungsprotokoll mit deinem Namen und Datum festgehalten.'
              : 'You are editing the answers to the main-event fields (the light-blue „Main-event fields" columns). They are stored in the person’s registration on the main event — not per sub-event. Every change is recorded in the audit log with your name and date.'}
          </p>
          {(() => {
            const parentFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.type !== 'document' && f.label && f.label.trim());
            if (parentFields.length === 0) {
              return (
                <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
                  {isDe ? 'Dieses Hauptevent hat keine bearbeitbaren Felder.' : 'This main event has no editable fields.'}
                </p>
              );
            }
            return (
              <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                {parentFields.map(cf => {
                  const value = mainFieldsEditForm[cf.id] || '';
                  const setVal = (v: string): void => setMainFieldsEditForm(prev => ({ ...prev, [cf.id]: v }));
                  const labelEl = (
                    <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                      {cf.label}{cf.required && <span style={{ color: 'var(--dex-red, #c00)' }}> *</span>}
                    </label>
                  );
                  if (cf.type === 'select' && !cf.multi && cf.options && cf.options.length > 0) {
                    return (
                      <div key={cf.id}>
                        {labelEl}
                        <select className="form-select" value={value} onChange={e => setVal(e.target.value)} style={{ width: '100%' }}>
                          <option value="">{isDe ? '— bitte wählen —' : '— please choose —'}</option>
                          {cf.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                        </select>
                      </div>
                    );
                  }
                  if (cf.type === 'select' && cf.multi && cf.options && cf.options.length > 0) {
                    const selected = value.split(' | ').map(s => s.trim()).filter(Boolean);
                    return (
                      <div key={cf.id} style={{ gridColumn: '1 / -1' }}>
                        {labelEl}
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
                    return (
                      <div key={cf.id}>
                        {labelEl}
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, padding: '6px 10px', background: isChecked ? 'rgba(134,188,37,0.12)' : 'var(--dex-gray-50)', borderRadius: 6, cursor: 'pointer', fontSize: '0.85rem' }}>
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={e => setVal(e.target.checked ? 'true' : 'false')}
                            style={{ accentColor: 'var(--dex-green)' }}
                          />
                          {isChecked ? (isDe ? 'Ja' : 'Yes') : (isDe ? 'Nein' : 'No')}
                        </label>
                      </div>
                    );
                  }
                  if (cf.type === 'number') {
                    return (
                      <div key={cf.id}>
                        {labelEl}
                        <input className="form-input" type="number" value={value} onChange={e => setVal(e.target.value)} style={{ width: '100%' }} />
                      </div>
                    );
                  }
                  return (
                    <div key={cf.id}>
                      {labelEl}
                      <input className="form-input" value={value} onChange={e => setVal(e.target.value)} style={{ width: '100%' }} />
                    </div>
                  );
                })}
              </div>
            );
          })()}
          {mainFieldsEditError && (
            <div style={{ padding: '8px 12px', background: 'rgba(218,41,28,0.08)', border: '1px solid var(--dex-red, #c00)', borderRadius: 6, fontSize: '0.85rem', color: 'var(--dex-red, #c00)' }}>
              {mainFieldsEditError}
            </div>
          )}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
            <button type="button" className="btn btn-outline" onClick={closeMainFieldsEdit} disabled={mainFieldsEditSaving}>
              {isDe ? 'Abbrechen' : 'Cancel'}
            </button>
            <button
              type="button"
              className="btn btn-primary"
              onClick={saveMainFieldsEdit}
              disabled={mainFieldsEditSaving}
              style={{ opacity: mainFieldsEditSaving ? 0.6 : 1 }}
            >
              {mainFieldsEditSaving ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Speichern' : 'Save')}
            </button>
          </div>
        </Modal>
  );
};

