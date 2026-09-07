/* SubEventFieldsModal — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Custom-Fields eines Sub-Events (v10.12): erst ausfuellen, dann wandert die
 * Session in `selectedSessions`. Inhalt zeichengleich uebernommen; die
 * Anzeige-Bedingung (`pendingSubEventModal`) ist beim Aufrufer geblieben.
 * v31.2: Optik auf dex-ui-Klassen umgestellt; Logik, Werteformat, Speichern gleich. */
import * as React from 'react';
import { DeloitteEvent, EventSpecificField } from '../../types';
import Modal from '../Modal';
import { InfoTooltip } from '../InfoTooltip';
import { Locale } from '../../context/LanguageContext';
// v31.2: Gemeinsame UI-Klassen statt Inline-Styles — Inline kann kein :hover,
// und die Optionen eines Multi-Select lasen sich vorher wie Beschriftungen.
import { cx } from '../dexUi';
import { MessageSquare, AlertCircle, Check } from '../Icons';

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

        const isDe = locale === 'de';
        const termLc = childTermSingular || (isDe ? 'Sub-Event' : 'sub-event');
        const modalTitle = ce.title || childTermSingular || (isDe ? 'Sub-Event' : 'Sub-event');
        // v31.2: Beschriftung — Frage, Pflicht-Stern bzw. „(optional)", InfoTooltip
        // (v11.16: Tooltip statt grauer Inline-Beschreibung, wie auf der Anmeldeseite).
        // Einmal gebaut, weil die Schalter-Frage sie IN der Zeile trägt, alle anderen darüber.
        const labelBody = (f: EventSpecificField): React.ReactNode => (
          <>
            {fLabel(f)}
            {f.required
              ? <span style={{ color: 'var(--dex-red, #da291c)' }} aria-hidden="true">*</span>
              : <span className="dex-ui-label-optional">(optional)</span>}
            {fHelp(f) && <InfoTooltip text={fHelp(f)} />}
          </>
        );

        return (
          <Modal
            open={true}
            onClose={onCancel}
            maxWidth={520}
            ariaLabel={modalTitle}
            icon={<MessageSquare size={20} strokeWidth={2} />}
            title={modalTitle}
            // v31.2: Untertitel nennt die Folge — ausgewählt ja, angemeldet erst mit dem Formular.
            subtitle={isDe
              ? `Beantworte kurz die Fragen zu diesem ${termLc} — danach ist es ausgewählt.`
              : `Answer a few questions about this ${termLc} — then it is selected.`}
            footer={<>
              <button type="button" className="btn btn-secondary" onClick={onCancel}>
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button type="button" className="btn btn-primary" onClick={onConfirm} disabled={!canSubmit}>
                {isDe ? 'Übernehmen' : 'Apply'}
              </button>
            </>}
          >
              <div>
                {fields.map(f => {
                  const val = draft[f.id] || '';
                  // v31.2: Ja/Nein-Frage als Schalter-Zeile — die Frage steht IN
                  // der Zeile, nicht als Beschriftung über einem nackten „Ja".
                  if (f.type === 'checkbox') {
                    const on = val === 'true';
                    return (
                      <div key={f.id} className="dex-ui-field">
                        <label className={cx('dex-ui-toggle-row', on && 'is-active')}>
                          <input
                            type="checkbox"
                            checked={on}
                            onChange={e => updateFieldValue(f.id, e.target.checked ? 'true' : 'false')}
                          />
                          <span className="dex-ui-toggle-row-body">
                            <span className="dex-ui-toggle-row-title">{labelBody(f)}</span>
                            <span className="dex-ui-toggle-row-desc">{on ? (isDe ? 'Ja' : 'Yes') : (isDe ? 'Nein' : 'No')}</span>
                          </span>
                        </label>
                      </div>
                    );
                  }
                  return (
                    <div key={f.id} className="dex-ui-field">
                      <label className="dex-ui-label">{labelBody(f)}</label>
                      {f.type === 'select' && f.multi ? (
                        // v31.2: Mehrfachauswahl als Chips — jede Option ist ein
                        // Knopf mit Hover; der gespeicherte Wert bleibt „A | B".
                        <div className="dex-ui-inline" role="group" aria-label={fLabel(f)}>
                          {(f.options || []).map((opt, optIdx) => {
                            const current = val.split(' | ').map(s => s.trim()).filter(Boolean);
                            const checked = current.indexOf(opt) >= 0;
                            return (
                              <button
                                key={opt}
                                type="button"
                                className={cx('dex-ui-chip', checked && 'is-active')}
                                aria-pressed={checked}
                                onClick={() => {
                                  const next = !checked
                                    ? [...current, opt]
                                    : current.filter(x => x !== opt);
                                  updateFieldValue(f.id, next.join(' | '));
                                }}
                              >
                                {checked && <Check size={12} />}
                                {fOpt(f, opt, optIdx)}
                              </button>
                            );
                          })}
                        </div>
                      ) : f.type === 'select' ? (
                        <select
                          className="dex-ui-select"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                        >
                          <option value="">{isDe ? '— bitte wählen —' : '— please select —'}</option>
                          {(f.options || []).map((opt, optIdx) => <option key={opt} value={opt}>{fOpt(f, opt, optIdx)}</option>)}
                        </select>
                      ) : f.type === 'number' ? (
                        <input
                          type="number"
                          className="dex-ui-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                        />
                      ) : (
                        <input
                          type="text"
                          className="dex-ui-input"
                          value={val}
                          onChange={e => updateFieldValue(f.id, e.target.value)}
                        />
                      )}
                    </div>
                  );
                })}
              </div>

              {!canSubmit && requiredMissing.length > 0 && (
                <div className="dex-ui-callout dex-ui-callout--warn" role="status">
                  <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                  <span>
                    {isDe ? 'Noch offen: ' : 'Still open: '}
                    <strong>{requiredMissing.join(', ')}</strong>
                    {isDe ? ' — danach kannst du übernehmen.' : ' — then you can apply.'}
                  </span>
                </div>
              )}
          </Modal>
        );
};
