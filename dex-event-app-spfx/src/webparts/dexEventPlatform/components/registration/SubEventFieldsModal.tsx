/* SubEventFieldsModal — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Custom-Fields eines Sub-Events (v10.12): erst ausfuellen, dann wandert die
 * Session in `selectedSessions`. Inhalt zeichengleich uebernommen; die
 * Anzeige-Bedingung (`pendingSubEventModal`) ist beim Aufrufer geblieben. */
import * as React from 'react';
import { DeloitteEvent, EventSpecificField } from '../../types';
import Modal from '../Modal';
import { InfoTooltip } from '../InfoTooltip';
import { Locale } from '../../context/LanguageContext';

/** Custom-Fields eines Sub-Events (v10.12). */
export interface SubEventFieldsModalProps {
  childEvents: DeloitteEvent[];
  childTermSingular: string;
  locale: Locale;
  pendingSubEventModal: { subEventId: string; draftValues: Record<string, string>; };
  setPendingSubEventModal: React.Dispatch<React.SetStateAction<{ subEventId: string; draftValues: Record<string, string>; }>>;
  setSelectedSessions: React.Dispatch<React.SetStateAction<Set<string>>>;
  setSessionFieldValues: React.Dispatch<React.SetStateAction<Record<string, Record<string, string>>>>;
}
export const SubEventFieldsModal: React.FC<SubEventFieldsModalProps> = (p) => {
  const { childEvents, childTermSingular, locale, pendingSubEventModal, setPendingSubEventModal, setSelectedSessions, setSessionFieldValues } = p;
        const ce = childEvents.find(c => c.id === pendingSubEventModal.subEventId);
        if (!ce) return null;
        const draft = pendingSubEventModal.draftValues;
        // v24.16 BUG-FIX: showIf (Sichtbarkeitsbedingung) auch im Sub-Event-
        // Modal anwenden — bedingte Fragen wurden vorher IMMER angezeigt und
        // blockierten als Pflichtfeld die Bestätigung. Quell-Antwort steht im
        // Sub-Event-eigenen `draft`.
        const fields = (ce.eventSpecificFields || [])
          .filter(f => f && f.label)
          .filter(f => {
            if (!f.showIf || !f.showIf.fieldId) return true;
            const raw = (draft[f.showIf.fieldId] || '').trim();
            if (!raw) return false;
            const answers = raw.indexOf(' | ') >= 0
              ? raw.split(' | ').map(s => s.trim()).filter(Boolean)
              : [raw];
            return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
          });
        const setDraft = (next: Record<string, string>): void => {
          setPendingSubEventModal(prev => prev ? { ...prev, draftValues: next } : prev);
        };
        const updateFieldValue = (fieldId: string, value: string): void => {
          setDraft({ ...draft, [fieldId]: value });
        };
        // v17.22: EN-Varianten auch im Sub-Event-Modal respektieren — geknüpft
        // an die Bilingual-Einstellung DES Sub-Events (ce), nicht des Parents.
        const useEnHere = locale === 'en' && !!ce.bilingualFields;
        const fLabel = (f: EventSpecificField): string =>
          (useEnHere && f.labelEn && f.labelEn.trim()) ? f.labelEn : f.label;
        const fHelp = (f: EventSpecificField): string | undefined =>
          (useEnHere && f.helpTextEn && f.helpTextEn.trim()) ? f.helpTextEn : f.helpText;
        const fOpt = (f: EventSpecificField, opt: string, idx: number): string =>
          (useEnHere && f.optionsEn && f.optionsEn[idx] && f.optionsEn[idx].trim()) ? f.optionsEn[idx] : opt;
        const requiredMissing = fields.filter(f => f.required && !((draft[f.id] || '').trim())).map(f => fLabel(f));
        const canSubmit = requiredMissing.length === 0;
        const onConfirm = (): void => {
          if (!canSubmit) return;
          setSessionFieldValues(prev => ({ ...prev, [ce.id]: { ...draft } }));
          setSelectedSessions(prev => {
            const next = new Set(prev);
            next.add(ce.id);
            return next;
          });
          setPendingSubEventModal(null);
        };
        const onCancel = (): void => setPendingSubEventModal(null);

        return (
          <Modal
            open={true}
            onClose={onCancel}
            maxWidth={520}
            padding={24}
            ariaLabel={ce.title || childTermSingular || (locale === 'de' ? 'Sub-Event' : 'Sub-event')}
          >
              <h3 style={{ margin: '0 0 6px', fontSize: '1.1rem' }}>
                {ce.title || childTermSingular || (locale === 'de' ? 'Sub-Event' : 'Sub-event')}
              </h3>
              <p style={{ margin: '0 0 18px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {locale === 'de'
                  ? `Bitte beantworte die Fragen für dieses ${childTermSingular || 'Sub-Event'}:`
                  : `Please answer the questions for this ${childTermSingular || 'sub-event'}:`}
              </p>

              <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginBottom: 20 }}>
                {fields.map(f => {
                  const val = draft[f.id] || '';
                  return (
                    <div key={f.id}>
                      <label className="form-label" style={{ display: 'block', fontSize: '0.85rem', marginBottom: 4 }}>
                        {fLabel(f)}
                        {f.required && <span style={{ color: 'var(--dex-red, #c00)', marginLeft: 4 }}>*</span>}
                        {/* v11.16: konsistenter InfoTooltip statt grauer
                            Inline-Beschreibung — gleicher Look wie auf
                            der Haupt-Register-Page. */}
                        {fHelp(f) && <InfoTooltip text={fHelp(f)} />}
                      </label>
                      {f.type === 'select' && f.multi ? (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          {(f.options || []).map((opt, optIdx) => {
                            const current = val.split(' | ').map(s => s.trim()).filter(Boolean);
                            const checked = current.indexOf(opt) >= 0;
                            return (
                              <label key={opt} style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={e => {
                                    const next = e.target.checked
                                      ? [...current, opt]
                                      : current.filter(x => x !== opt);
                                    updateFieldValue(f.id, next.join(' | '));
                                  }}
                                />
                                {fOpt(f, opt, optIdx)}
                              </label>
                            );
                          })}
                        </div>
                      ) : f.type === 'select' ? (
                        <select
                          className="form-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                          style={{ width: '100%', fontSize: '0.9rem' }}
                        >
                          <option value="">{locale === 'de' ? '— bitte wählen —' : '— please select —'}</option>
                          {(f.options || []).map((opt, optIdx) => <option key={opt} value={opt}>{fOpt(f, opt, optIdx)}</option>)}
                        </select>
                      ) : f.type === 'checkbox' ? (
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
                          <input
                            type="checkbox"
                            checked={val === 'true'}
                            onChange={e => updateFieldValue(f.id, e.target.checked ? 'true' : 'false')}
                          />
                          {locale === 'de' ? 'Ja' : 'Yes'}
                        </label>
                      ) : f.type === 'number' ? (
                        <input
                          type="number"
                          className="form-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                          style={{ width: '100%', fontSize: '0.9rem' }}
                        />
                      ) : (
                        <input
                          type="text"
                          className="form-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                          style={{ width: '100%', fontSize: '0.9rem' }}
                        />
                      )}
                    </div>
                  );
                })}
              </div>

              {!canSubmit && requiredMissing.length > 0 && (
                <div style={{ fontSize: '0.78rem', color: 'var(--dex-red, #c00)', marginBottom: 12 }}>
                  {locale === 'de' ? 'Pflichtfelder fehlen: ' : 'Required fields missing: '}{requiredMissing.join(', ')}
                </div>
              )}

              <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button type="button" className="btn btn-secondary" onClick={onCancel}>
                  {locale === 'de' ? 'Abbrechen' : 'Cancel'}
                </button>
                <button type="button" className="btn btn-primary" onClick={onConfirm} disabled={!canSubmit}>
                  {locale === 'de' ? 'Bestätigen' : 'Confirm'}
                </button>
              </div>
          </Modal>
        );
};
