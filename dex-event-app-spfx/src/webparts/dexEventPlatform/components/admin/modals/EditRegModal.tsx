/* EditRegModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14077-14370 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.2: Auf das gemeinsame `Modal` (Kopf, Fuß, Escape, Backdrop) und die
 * dex-ui-Klassen umgestellt. Die Reihenfolge folgt jetzt dem, was der
 * Organizer hier tut: Person korrigieren → Antworten zum Event anpassen →
 * Profil nachsehen. Das M365-Profil ist nicht änderbar und steht deshalb im
 * Aufklapper statt als sechs graue Kästen zwischen den Eingaben. Der alte
 * Erklär-Absatz (sieben Zeilen) steht als je ein Satz dort, wo er gilt:
 * Tenant-Prüfung unter der E-Mail, Profil-Herkunft im Aufklapper,
 * Protokollierung im Untertitel. Gespeichert wird exakt dasselbe wie vorher.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { AlertCircle, ChevronDown, Pencil } from '../../Icons';
import { cx } from '../../dexUi';
import { MultiSelectDropdown } from '../../MultiSelectDropdown';
import { DeloitteEvent } from '../../../types';
import { FieldSelectInput, fieldVisibleByShowIf } from './FieldSelectInput';

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
  // v31.2: Der Aufrufer mountet den Dialog je Öffnen neu — der Aufklapper
  // startet deshalb verlässlich geschlossen.
  const [profileOpen, setProfileOpen] = React.useState(false);

  // Vorname / Nachname / E-Mail sind seit v9.7 editierbar (mit
  // Deloitte-Domain- und Tenant-Existenz-Check beim Speichern).
  // Die uebrigen Profil-Felder bleiben read-only — sie kommen
  // aus dem M365-Profil und werden bei einer Mail-Änderung
  // mit den Profil-Daten der neuen Person überschrieben.
  const readOnlyFields: Array<{ key: string; label: string }> = [
    { key: 'Anrede', label: isDe ? 'Anrede' : 'Salutation' },
    { key: 'Phone', label: isDe ? 'Telefon' : 'Phone' },
    { key: 'Department', label: 'Department' },
    { key: 'Location', label: isDe ? 'Standort' : 'Location' },
    { key: 'JobTitle', label: 'Job Title' },
    { key: 'Status', label: 'Status' },
  ];

  // B2Run-Starter-Typ (Funstarter / Durchstarter). Hardcoded SP-Spalte auf
  // der Teilnehmerliste (kein regulärer Custom-Field-Eintrag), daher explizit
  // hier gerendert. Updates BEIDE intern getrackten Felder zugleich
  // (StarterType + PreferredStarterType) — die getrennte Speicherung von
  // „aktuell vs. Wunsch" ist Implementierungs-Detail für die Warteliste-
  // Nachrück-Logik und braucht im Edit-Modal keine UI-Komplexität. v10.15+
  const hasStarterType = selectedEvent.durchstarterCapacity !== undefined
    && selectedEvent.funstarterCapacity !== undefined
    && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
  // v31.2: Die Gruppen heißen je Event anders (splitLabelA/B) — der Wert in
  // der Spalte bleibt „Durchstarter"/„Funstarter", nur die Beschriftung folgt
  // dem Event, wie in Tabelle und KPI-Kacheln.
  const starterOptions = [
    { value: 'Durchstarter', label: (selectedEvent.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter' },
    { value: 'Funstarter', label: (selectedEvent.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter' },
  ];
  const customFields = selectedEvent.eventSpecificFields || [];
  const hasAnswers = hasStarterType || customFields.length > 0;

  const spOf = (fieldId: string): string => {
    const src = (selectedEvent.eventSpecificFields || []).find(f => f.id === fieldId);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return src ? ((src as any).spInternalName || '') : '';
  };

  const renderEditable = (key: string, label: string, type?: string, help?: React.ReactNode): React.ReactNode => (
    <div className="dex-ui-field">
      <label className="dex-ui-label" htmlFor={`editreg-${key}`}>{label}</label>
      <input
        id={`editreg-${key}`}
        className="dex-ui-input"
        type={type || 'text'}
        value={editForm[key] || ''}
        onChange={e => setEditForm(prev => ({ ...prev, [key]: e.target.value }))}
      />
      {help && <p className="dex-ui-help">{help}</p>}
    </div>
  );
  const renderReadOnly = (label: string, value: string): React.ReactNode => (
    <div className="dex-ui-field">
      <span className="dex-ui-label" style={{ color: 'var(--dex-gray-500)' }}>{label}</span>
      <div style={{
        padding: '8px 12px', background: 'var(--dex-gray-50, #fafafa)',
        border: '1px solid var(--dex-gray-200, #e8e8e8)', borderRadius: 10,
        fontSize: '0.88rem', lineHeight: 1.5, minHeight: 38, boxSizing: 'border-box',
        color: value ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
      }}>
        {value || (isDe ? '— nicht gesetzt —' : '— not set —')}
      </div>
    </div>
  );

  return (
    <Modal
      open
      onClose={closeEditModal}
      dismissable={!isSavingEdit}
      maxWidth={920}
      ariaLabel={isDe ? 'Teilnehmer bearbeiten' : 'Edit attendee'}
      icon={<Pencil size={20} />}
      title={<>
        {isDe ? 'Teilnehmer bearbeiten' : 'Edit attendee'}
        {' — '}
        <span style={{ color: 'var(--dex-green-dark)' }}>{editForm.Vorname} {editForm.Nachname}</span>
        {editForm.Status && (
          <span className="dex-ui-pill dex-ui-pill--gray" style={{ marginLeft: 10, verticalAlign: 'middle' }}>{editForm.Status}</span>
        )}
      </>}
      subtitle={isDe
        ? 'Name, E-Mail und die Antworten dieser Person korrigieren. Jede Änderung landet mit Datum und deinem Namen im Audit-Log und im ChangeLog der Person.'
        : 'Fix this person’s name, email and answers. Every change is logged in the audit log and in the attendee’s ChangeLog with date and your name.'}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={closeEditModal} disabled={isSavingEdit}>
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        <button type="button" className="btn btn-primary" onClick={saveEdit} disabled={isSavingEdit}>
          {isSavingEdit ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Speichern' : 'Save')}
        </button>
      </>}
    >
      {/* 1) Person — die Felder, die man hier am häufigsten korrigiert */}
      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{isDe ? 'Person' : 'Person'}</div>
        <p className="dex-ui-section-desc">
          {isDe ? 'Zum Beispiel nach einem Tippfehler bei der manuellen Anlage.' : 'For example after a typo during manual creation.'}
        </p>
        <div className="dex-ui-grid-2">
          {renderEditable('Vorname', isDe ? 'Vorname' : 'First name')}
          {renderEditable('Nachname', isDe ? 'Nachname' : 'Last name')}
        </div>
        <div style={{ marginTop: 16 }}>
          {renderEditable('ParticipantEmail', 'E-Mail', 'email', isDe
            ? 'Nur Deloitte-Adressen (@deloitte.de / @deloitte.com). Beim Speichern prüft die App, ob die Person im Tenant existiert — externe Adressen sind nicht erlaubt. Bei neuer Adresse zieht die App Telefon, Department, Standort und Job Title der neuen Person automatisch nach.'
            : 'Only Deloitte addresses (@deloitte.de / @deloitte.com). On save the app verifies that the person exists in the tenant — external addresses are not allowed. With a new address, phone, department, location and job title of the new person are refreshed automatically.')}
        </div>
      </div>

      {/* 2) Antworten zum Event — DAS ist der editierbare Teil. Renderer
          abhängig vom Field-Type (text/number/select/checkbox). Multi-Select
          speichert Werte als " | "-getrennten String, identisch zum
          Registrierungs-Pfad. */}
      {hasAnswers && (
        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{isDe ? 'Antworten zu diesem Event' : 'Answers for this event'}</div>
          <p className="dex-ui-section-desc">
            {isDe
              ? 'Diese Antworten werden gespeichert — das M365-Profil unten nicht.'
              : 'These answers are saved — the M365 profile below is not.'}
          </p>
          <div className="dex-ui-grid-2">
            {hasStarterType && (
              <div className="dex-ui-field" style={{ gridColumn: '1 / -1' }}>
                <span className="dex-ui-label" id="editreg-startertype-label">
                  {isDe ? 'Zu welcher Gruppe gehört die Person?' : 'Which group does this person belong to?'}
                </span>
                <div className="dex-ui-inline" role="group" aria-labelledby="editreg-startertype-label">
                  {starterOptions.map(o => {
                    const active = (editForm.StarterType || '') === o.value;
                    return (
                      <button
                        key={o.value}
                        type="button"
                        className={cx('dex-ui-chip', active && 'is-active')}
                        aria-pressed={active}
                        onClick={() => {
                          // Beide Felder synchron halten — der Aktuelle wechselt
                          // mit, der Wunsch ebenfalls (User-Erwartung: „ich ändere
                          // den Starter-Typ" = beides ändert sich). v31.2: Nochmal
                          // anklicken leert die Auswahl — das war vorher die
                          // Option „— bitte wählen —" im Dropdown.
                          const v = active ? '' : o.value;
                          setEditForm(prev => ({ ...prev, StarterType: v, PreferredStarterType: v }));
                        }}
                      >
                        {o.label}
                      </button>
                    );
                  })}
                </div>
                <p className="dex-ui-help">
                  {isDe
                    ? 'Starter-Typ für die Zuteilung: aktueller Typ und Wunsch wechseln zusammen. Noch einmal anklicken setzt die Auswahl zurück.'
                    : 'Starter type for the allocation: current type and preference change together. Click again to clear the selection.'}
                </p>
              </div>
            )}

            {customFields.map(cf => {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const sp = (cf as any).spInternalName || '';
              if (!sp) return null;
              // v30.95: Sichtbarkeitsregel wie auf der Anmeldeseite —
              // „Mobilnummer" hing als Pflichtfeld mit Stern im Dialog,
              // obwohl der Infoservice auf Nein stand.
              if (!fieldVisibleByShowIf(cf, fid => editForm[spOf(fid)] || '')) return null;
              const value = editForm[sp] || '';
              const setVal = (v: string): void => setEditForm(prev => ({ ...prev, [sp]: v }));
              const inputId = `editreg-cf-${cf.id}`;
              const labelEl = (
                <label className="dex-ui-label" htmlFor={inputId}>
                  {cf.label}{cf.required && <span style={{ color: 'var(--dex-red, #c00)' }} aria-hidden="true">*</span>}
                </label>
              );

              // Single-Select-Dropdown. v30.95: dieselbe Kombibox wie
              // die Anmeldeseite (Kategorien, Wert „Kategorie Option").
              if (cf.type === 'select' && !cf.multi && cf.options && cf.options.length > 0) {
                return (
                  <div key={cf.id} className="dex-ui-field">
                    {labelEl}
                    <FieldSelectInput field={cf} value={value} onChange={setVal} isDe={isDe} />
                  </div>
                );
              }

              // v11.89: Multi-Select-Dropdown (vorher Checkbox-Liste).
              // Werte werden weiterhin ' | '-getrennt gespeichert.
              if (cf.type === 'select' && cf.multi && cf.options && cf.options.length > 0) {
                const selected = value.split(' | ').map(s => s.trim()).filter(Boolean);
                return (
                  <div key={cf.id} className="dex-ui-field" style={{ gridColumn: '1 / -1' }}>
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

              // Checkbox (true/false). v31.2: als Schalter mit Ja/Nein daneben —
              // eine nackte Checkbox sagt nicht, was der Haken bedeutet.
              if (cf.type === 'checkbox') {
                const isChecked = value === 'true' || value === '1';
                return (
                  <div key={cf.id} className="dex-ui-field">
                    {labelEl}
                    <label className="dex-ui-switch" style={{ marginTop: 6 }}>
                      <input
                        id={inputId}
                        type="checkbox"
                        checked={isChecked}
                        onChange={e => setVal(e.target.checked ? 'true' : 'false')}
                      />
                      <span className="dex-ui-switch-track" />
                      <span className="dex-ui-switch-label">{isChecked ? (isDe ? 'Ja' : 'Yes') : (isDe ? 'Nein' : 'No')}</span>
                    </label>
                  </div>
                );
              }

              // Number
              if (cf.type === 'number') {
                return (
                  <div key={cf.id} className="dex-ui-field">
                    {labelEl}
                    <input id={inputId} className="dex-ui-input" type="number" value={value} onChange={e => setVal(e.target.value)} />
                  </div>
                );
              }

              // Default: text-Input (auch für 'text', 'user', 'roommate')
              return (
                <div key={cf.id} className="dex-ui-field">
                  {labelEl}
                  <input id={inputId} className="dex-ui-input" value={value} onChange={e => setVal(e.target.value)} />
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* 3) Profil aus M365 — nur lesen, deshalb zu */}
      <div className="dex-ui-section">
        <button
          type="button"
          className={cx('dex-ui-disclosure', profileOpen && 'is-open')}
          onClick={() => setProfileOpen(o => !o)}
          aria-expanded={profileOpen}
        >
          <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
          {isDe ? 'Profil aus M365 — nicht änderbar' : 'M365 profile — read-only'}
          <span className="dex-ui-disclosure-count">{readOnlyFields.length}</span>
        </button>
        {profileOpen && (
          <div className="dex-ui-disclosure-body dex-ui-fade-in">
            <p className="dex-ui-section-desc" style={{ margin: '0 0 12px' }}>
              {isDe
                ? 'Anrede, Telefon, Department, Standort und Job Title kommen aus dem M365-Profil und werden bei einem Mail-Wechsel automatisch nachgezogen. Den Status änderst du über die Aktions-Knöpfe in der Liste.'
                : 'Salutation, phone, department, location and job title come from the M365 profile and are refreshed automatically when the email changes. The status is changed via the action buttons in the list.'}
            </p>
            <div className="dex-ui-grid-3">
              {readOnlyFields.map(f => (
                <React.Fragment key={f.key}>{renderReadOnly(f.label, editForm[f.key] || '')}</React.Fragment>
              ))}
            </div>
          </div>
        )}
      </div>

      {editError && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>{editError}</span>
        </div>
      )}
    </Modal>
  );
};
