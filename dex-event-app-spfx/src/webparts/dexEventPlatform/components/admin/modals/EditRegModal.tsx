/* EditRegModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14077-14370 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { Pencil, X } from '../../Icons';
import { MultiSelectDropdown } from '../../MultiSelectDropdown';
import { DeloitteEvent } from '../../../types';

export interface EditRegModalProps {
  closeEditModal: () => void;
  editError: string;
  editForm: Record<string, string>;
  isDe: boolean;
  isSavingEdit: boolean;
  saveEdit: () => Promise<void>;
  selectedEvent: DeloitteEvent;
  setEditForm: React.Dispatch<React.SetStateAction<Record<string, string>>>;
}

export const EditRegModal: React.FC<EditRegModalProps> = (p) => {
  const { closeEditModal, editError, editForm, isDe, isSavingEdit, saveEdit, selectedEvent, setEditForm } = p;
  return (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
          onClick={() => { if (!isSavingEdit) closeEditModal(); }}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 920, maxHeight: '90vh', overflow: 'auto',
              padding: 24, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
                <Pencil size={18} />{' '}
                {isDe ? 'Teilnehmer bearbeiten' : 'Edit attendee'}
                {' — '}
                <span style={{ color: 'var(--dex-green-dark)' }}>
                  {editForm.Vorname} {editForm.Nachname}
                </span>
              </h3>
              <button
                onClick={closeEditModal}
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}
                aria-label={isDe ? 'Schließen' : 'Close'}
                disabled={isSavingEdit}
              ><X size={20} /></button>
            </div>
            <p style={{ margin: '0 0 16px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
              {isDe
                ? 'Hier kannst du Vorname, Nachname und E-Mail-Adresse korrigieren (z.B. nach einem Tippfehler bei der manuellen Anlage) sowie die event-spezifischen Felder anpassen. Beim Ändern der E-Mail wird geprüft, ob die Adresse zum Deloitte-Tenant gehört und die Person dort existiert — externe Adressen sind nicht erlaubt. Phone, Department, Standort und Job Title kommen aus dem M365-Profil und sind read-only — sie werden bei einem Mail-Wechsel automatisch nachgezogen. Den Status änderst du über die Aktions-Buttons in der Liste. Jede Änderung wird im Audit-Log und im ChangeLog des Teilnehmers mit Datum und deinem Namen protokolliert.'
                : 'You can fix first name, last name and email address here (e.g. after a typo during manual creation) and adjust event-specific fields. When changing the email, the app verifies that the address belongs to the Deloitte tenant and that the person exists there — external addresses are not allowed. Phone, Department, Location and Job Title come from the M365 profile and are read-only — they are refreshed automatically when the email changes. The status is changed via the action buttons in the list. Every change is logged in the audit log and in the attendee\'s ChangeLog with date and your name.'}
            </p>

            {(() => {
              // Vorname / Nachname / E-Mail sind seit v9.7 editierbar (mit
              // Deloitte-Domain- und Tenant-Existenz-Check beim Speichern).
              // Die uebrigen Profil-Felder bleiben read-only — sie kommen
              // aus dem M365-Profil und werden bei einer Mail-Änderung
              // mit den Profil-Daten der neuen Person überschrieben.
              const editableStammFields: Array<{ key: string; label: string; type?: string }> = [
                { key: 'Vorname', label: isDe ? 'Vorname' : 'First name' },
                { key: 'Nachname', label: isDe ? 'Nachname' : 'Last name' },
                { key: 'ParticipantEmail', label: 'E-Mail', type: 'email' },
              ];
              const readOnlyFields: Array<{ key: string; label: string }> = [
                { key: 'Anrede', label: isDe ? 'Anrede' : 'Salutation' },
                { key: 'Phone', label: isDe ? 'Telefon' : 'Phone' },
                { key: 'Department', label: 'Department' },
                { key: 'Location', label: isDe ? 'Standort' : 'Location' },
                { key: 'JobTitle', label: 'Job Title' },
                { key: 'Status', label: 'Status' },
              ];
              const renderReadOnly = (label: string, value: string): React.ReactNode => (
                <div>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-500)', marginBottom: 4 }}>
                    {label}
                  </label>
                  <div style={{
                    width: '100%', padding: '8px 12px',
                    background: 'var(--dex-gray-50, #fafafa)',
                    border: '1px solid var(--dex-gray-200, #e5e7eb)',
                    borderRadius: 6, fontSize: '0.88rem',
                    color: value ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
                    minHeight: 38, lineHeight: 1.5,
                  }}>
                    {value || (isDe ? '— nicht gesetzt —' : '— not set —')}
                  </div>
                </div>
              );
              const renderEditable = (key: string, label: string, type?: string): React.ReactNode => (
                <div>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                    {label}
                  </label>
                  <input
                    className="form-input"
                    type={type || 'text'}
                    value={editForm[key] || ''}
                    onChange={e => setEditForm(prev => ({ ...prev, [key]: e.target.value }))}
                    style={{ fontSize: '0.88rem' }}
                  />
                </div>
              );
              return (
                <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                  {/* Anrede zuerst (read-only) */}
                  <div>{renderReadOnly(isDe ? 'Anrede' : 'Salutation', editForm.Anrede || '')}</div>
                  {/* Vorname + Nachname editierbar */}
                  {editableStammFields.filter(f => f.key !== 'ParticipantEmail').map(f => (
                    <div key={f.key}>
                      {renderEditable(f.key, f.label, f.type)}
                    </div>
                  ))}
                  {/* E-Mail editierbar (volle Breite) */}
                  <div style={{ gridColumn: '1 / -1' }}>
                    {renderEditable('ParticipantEmail', 'E-Mail', 'email')}
                    <p style={{ margin: '4px 0 0', fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                      {isDe
                        ? 'Nur Deloitte-Adressen (@deloitte.de / @deloitte.com). Beim Speichern wird die Person im Tenant verifiziert.'
                        : 'Only Deloitte addresses (@deloitte.de / @deloitte.com). The person is verified in the tenant on save.'}
                    </p>
                  </div>
                  {/* Restliche Profil-Felder read-only */}
                  {readOnlyFields.filter(f => f.key !== 'Anrede').map(f => (
                    <div key={f.key}>
                      {renderReadOnly(f.label, editForm[f.key] || '')}
                    </div>
                  ))}

                  {/* B2Run-Starter-Typ (Funstarter / Durchstarter). Hardcoded
                      SP-Spalte auf der Teilnehmerliste (kein regulärer
                      Custom-Field-Eintrag), daher explizit hier gerendert.
                      Updates BEIDE intern getrackten Felder zugleich
                      (StarterType + PreferredStarterType) — die getrennte
                      Speicherung von „aktuell vs. Wunsch" ist Implementierungs-
                      Detail für die Warteliste-Nachrück-Logik und braucht im
                      Edit-Modal keine UI-Komplexität. v10.15+ */}
                  {selectedEvent.durchstarterCapacity !== undefined
                    && selectedEvent.funstarterCapacity !== undefined
                    && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0) && (
                    <div style={{ gridColumn: '1 / -1', marginTop: 12, paddingTop: 16, borderTop: '1px solid var(--dex-gray-200)' }}>
                      <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                        {isDe ? 'B2Run-Starter-Typ' : 'B2Run starter type'}
                      </label>
                      <select
                        value={editForm.StarterType || ''}
                        onChange={e => {
                          const v = e.target.value;
                          // Beide Felder synchron halten — der Aktuelle wechselt
                          // mit, der Wunsch ebenfalls (User-Erwartung: „ich ändere
                          // den Starter-Typ" = beides ändert sich).
                          setEditForm(prev => ({ ...prev, StarterType: v, PreferredStarterType: v }));
                        }}
                        className="form-input"
                        style={{ maxWidth: 320 }}
                      >
                        <option value="">{isDe ? '— bitte wählen —' : '— please select —'}</option>
                        <option value="Durchstarter">Durchstarter</option>
                        <option value="Funstarter">Funstarter</option>
                      </select>
                    </div>
                  )}

                  {/* Custom Fields des Events — DAS ist der editierbare Teil.
                      Renderer abhängig vom Field-Type (text/number/select/
                      checkbox). Multi-Select speichert Werte als " | "-
                      getrennten String, identisch zum Registrierungs-Pfad. */}
                  {selectedEvent.eventSpecificFields && selectedEvent.eventSpecificFields.length > 0 && (
                    <div style={{ gridColumn: '1 / -1', marginTop: 12, paddingTop: 16, borderTop: '1px solid var(--dex-gray-200)' }}>
                      <h4 style={{ margin: '0 0 4px', fontSize: '0.92rem', color: 'var(--dex-gray-800)' }}>
                        {isDe ? 'Event-spezifische Felder (editierbar)' : 'Event-specific fields (editable)'}
                      </h4>
                      <p style={{ margin: '0 0 12px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                        {isDe
                          ? 'Nur diese Felder werden gespeichert.'
                          : 'Only these fields will be saved.'}
                      </p>
                      <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                        {selectedEvent.eventSpecificFields.map(cf => {
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const sp = (cf as any).spInternalName || '';
                          if (!sp) return null;
                          const value = editForm[sp] || '';
                          const setVal = (v: string): void => setEditForm(prev => ({ ...prev, [sp]: v }));
                          const labelEl = (
                            <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
                              {cf.label}{cf.required && <span style={{ color: 'var(--dex-red, #c00)' }}> *</span>}
                            </label>
                          );

                          // Single-Select-Dropdown
                          if (cf.type === 'select' && !cf.multi && cf.options && cf.options.length > 0) {
                            return (
                              <div key={cf.id}>
                                {labelEl}
                                <select
                                  className="form-select"
                                  value={value}
                                  onChange={e => setVal(e.target.value)}
                                  style={{ width: '100%' }}
                                >
                                  <option value="">{isDe ? '— bitte wählen —' : '— please choose —'}</option>
                                  {cf.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                                </select>
                              </div>
                            );
                          }

                          // v11.89: Multi-Select-Dropdown (vorher Checkbox-Liste).
                          // Werte werden weiterhin ' | '-getrennt gespeichert.
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

                          // Checkbox (true/false)
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

                          // Number
                          if (cf.type === 'number') {
                            return (
                              <div key={cf.id}>
                                {labelEl}
                                <input
                                  className="form-input"
                                  type="number"
                                  value={value}
                                  onChange={e => setVal(e.target.value)}
                                  style={{ width: '100%' }}
                                />
                              </div>
                            );
                          }

                          // Default: text-Input (auch für 'text', 'user', 'roommate')
                          return (
                            <div key={cf.id}>
                              {labelEl}
                              <input
                                className="form-input"
                                value={value}
                                onChange={e => setVal(e.target.value)}
                                style={{ width: '100%' }}
                              />
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </div>
              );
            })()}

            {editError && (
              <div style={{ marginTop: 12, padding: '8px 12px', background: 'rgba(218,41,28,0.08)', border: '1px solid var(--dex-red, #c00)', borderRadius: 6, fontSize: '0.85rem', color: 'var(--dex-red, #c00)' }}>
                {editError}
              </div>
            )}

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 20 }}>
              <button
                type="button"
                className="btn btn-outline"
                onClick={closeEditModal}
                disabled={isSavingEdit}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={saveEdit}
                disabled={isSavingEdit}
                style={{ opacity: isSavingEdit ? 0.6 : 1 }}
              >
                {isSavingEdit ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Speichern' : 'Save')}
              </button>
            </div>
          </div>
        </div>
  );
};

