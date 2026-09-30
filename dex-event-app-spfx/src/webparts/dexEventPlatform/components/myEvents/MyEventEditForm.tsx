/**
 * v32.53 — „Meine Events → Angaben bearbeiten" wie das Anmeldeformular.
 *
 * Bis v32.52 hatte der Bearbeiten-Modus eine eigene, verkürzte Feldlogik:
 * Datum als Freitext, Mehrfachauswahl als Einzel-Select (zeigte „—", obwohl
 * eine Antwort da war), keine Bedingungen (Zimmerpartner-Frage trotz
 * „Einzelzimmer"), keine Beschreibungen, keine Pflichtprüfung (Nutzer-Befund
 * 30.09.2026). Jetzt dieselben Bausteine wie auf der Anmeldeseite:
 * SingleSelectDropdown/MultiSelectDropdown (Mehrfachwerte „ | "-getrennt),
 * CustomDateInput (gleiches Speicherformat), StayRangePicker, UserFieldPicker,
 * `fieldVisibleByShowIf` und die Gruppen-Regel. Datei-Fragen bleiben bei
 * MyEventDocField.
 */
import * as React from 'react';
import { DeloitteEvent, EventSpecificField } from '../../types';
import { SPRegistration } from '../../services/EventService';
import { SingleSelectDropdown } from '../SingleSelectDropdown';
import { MultiSelectDropdown } from '../MultiSelectDropdown';
import { UserFieldPicker } from '../UserFieldPicker';
import { CustomDateInput } from '../CustomDateInput';
import { InfoTooltip } from '../InfoTooltip';
import { cx } from '../dexUi';
import { AlertCircle } from '../Icons';
import { fieldVisibleByShowIf } from '../admin/modals/FieldSelectInput';
import { renderFieldDescHtml } from '../registration/regHelpers';
import { isEventVisibleForUser } from '../EventListPage';

const LazyStayRangePicker = React.lazy(async () => {
  const m = await import('../StayRangePicker');
  return { default: m.StayRangePicker };
});

type UserHit = { email: string; displayName: string; location: string; jobTitle: string };

export interface MyEventEditFormProps {
  event: DeloitteEvent;
  registration: SPRegistration;
  editData: Record<string, string>;
  setEditData: (next: Record<string, string>) => void;
  locale: string;
  isSaving: boolean;
  onSave: () => void;
  onCancel: () => void;
  saveLabel: string;
  cancelLabel: string;
  searchUsers: (query: string, includeInternational?: boolean) => Promise<UserHit[]>;
  searchUser: (email: string) => Promise<{ displayName: string; location: string; jobTitle: string; department?: string; mobilePhone?: string; company?: string }>;
}

export function MyEventEditForm(p: MyEventEditFormProps): React.ReactElement {
  const { event, registration, editData, setEditData, locale, isSaving, onSave, onCancel, saveLabel, cancelLabel, searchUsers, searchUser } = p;
  const isDe = locale === 'de';
  const [versucht, setVersucht] = React.useState(false);
  // v17.22: EN-Varianten im bilingualen Modus (Label, Optionen, Beschreibung).
  const useEn = locale === 'en' && !!event.bilingualFields;
  const label = (f: EventSpecificField): string => (useEn && f.labelEn && f.labelEn.trim()) ? f.labelEn : f.label;
  const hilfe = (f: EventSpecificField): string => ((useEn && f.helpTextEn && f.helpTextEn.trim()) ? f.helpTextEn : f.helpText) || '';
  const optLabels = (f: EventSpecificField): string[] => (f.options || []).map((o, i) =>
    (useEn && f.optionsEn && f.optionsEn[i] && f.optionsEn[i].trim()) ? f.optionsEn[i] : o);

  const gruppe = (registration.StarterType || registration.PreferredStarterType || '').trim();
  const gruppeErlaubt = (f: EventSpecificField): boolean => {
    const g = f.onlyForGroup;
    if (!g || g === 'all') return true;
    return g === 'A' ? gruppe === 'Durchstarter' : gruppe === 'Funstarter';
  };
  // Live gegen die gerade bearbeiteten Antworten — wie auf der Anmeldeseite
  // blendet ein Wechsel auf „Einzelzimmer" die Zimmerpartner-Frage sofort aus.
  const felder = (event.eventSpecificFields || []).filter((f: EventSpecificField) =>
    f.label && f.type !== 'document' && gruppeErlaubt(f) && fieldVisibleByShowIf(f, id => editData[id] || ''));
  const leer = (f: EventSpecificField): boolean => {
    const v = (editData[f.id] || '').trim();
    return f.type === 'checkbox' ? v !== 'true' : !v;
  };
  const fehlend = felder.filter(f => f.required && leer(f));
  const setze = (id: string, v: string): void => setEditData({ ...editData, [id]: v });

  const speichern = (): void => {
    setVersucht(true);
    if (fehlend.length > 0) return;
    onSave();
  };

  const eingabe = (f: EventSpecificField, fehler: boolean): React.ReactNode => {
    const v = editData[f.id] || '';
    if (f.type === 'select' && f.multi) {
      return (
        <MultiSelectDropdown
          options={f.options || []}
          optionLabels={useEn ? optLabels(f) : undefined}
          value={v.split(' | ').map(s => s.trim()).filter(Boolean)}
          onChange={next => setze(f.id, next.join(' | '))}
          placeholder={isDe ? 'Bitte wählen' : 'Please select'}
          error={fehler}
        />
      );
    }
    if (f.type === 'select') {
      const labels = optLabels(f);
      return (
        <SingleSelectDropdown
          options={(f.options || []).map((o, i) => ({ value: o, label: labels[i] || o }))}
          value={v}
          onChange={next => setze(f.id, next)}
          placeholder={isDe ? 'Bitte wählen' : 'Please select'}
          error={fehler}
          ariaLabel={label(f)}
        />
      );
    }
    if (f.type === 'user' || f.type === 'roommate') {
      return (
        <UserFieldPicker
          value={v}
          onChange={next => setze(f.id, next)}
          // v29.40: Die Verteiler-Begrenzung des Feldes gilt auch beim Nachtragen.
          searchUsers={f.audienceOnly
            ? (async (q: string, intl?: boolean) => {
              const res = await searchUsers(q, intl);
              return res.filter(u => isEventVisibleForUser(event, u.email, u.location || '', [], u.jobTitle || ''));
            })
            : searchUsers}
          searchUserByEmail={searchUser}
          placeholder={isDe ? 'Name oder E-Mail eingeben…' : 'Type a name or email…'}
          errorStyle={fehler ? { borderColor: 'var(--dex-red, #da291c)' } : {}}
        />
      );
    }
    if (f.type === 'checkbox') {
      return (
        <label className={cx('dex-ui-toggle-row', v === 'true' && 'is-active')}>
          <input type="checkbox" checked={v === 'true'} onChange={e => setze(f.id, e.target.checked ? 'true' : 'false')} />
          <span className="dex-ui-toggle-row-body">
            <span className="dex-ui-toggle-row-title">
              {(useEn && f.confirmLabelEn && f.confirmLabelEn.trim() ? f.confirmLabelEn : f.confirmLabel) || label(f)}
            </span>
          </span>
        </label>
      );
    }
    if (f.type === 'date') {
      return <CustomDateInput value={v} onChange={next => setze(f.id, next)} withTime={!!f.withTime} isDe={isDe} error={fehler} required={!!f.required} />;
    }
    if (f.type === 'daterange') {
      return (
        <React.Suspense fallback={<div className="form-input" aria-hidden="true" style={{ minHeight: 48 }} />}>
          <LazyStayRangePicker
            value={v}
            onChange={(next: string) => setze(f.id, next)}
            isDe={isDe}
            rangeStart={f.rangeStart}
            rangeEnd={f.rangeEnd}
            maxNights={f.maxNights}
            required={f.required}
          />
        </React.Suspense>
      );
    }
    return (
      <input
        className={cx('form-input', fehler && 'dex-ui-input--error')}
        value={v}
        onChange={e => setze(f.id, e.target.value)}
        placeholder={label(f)}
        type={f.type === 'number' ? 'number' : 'text'}
      />
    );
  };

  return (
    <div className="dex-ui-section">
      <div className="dex-ui-section-title">{isDe ? 'Angaben bearbeiten' : 'Edit your details'}</div>
      <div className="dex-ui-grid-2 dex-answer-grid">
        {felder.map(f => {
          const text = hilfe(f);
          const inline = f.helpTextStyle === 'inline';
          const fehler = versucht && !!f.required && leer(f);
          // Mehrzeilige Eingaben (Zeitraum, Personen, Häkchen mit Text) über die
          // volle Breite, wie auf der Anmeldeseite.
          const breit = f.type === 'daterange' || f.type === 'checkbox';
          return (
            <div key={f.id} className="dex-answer-field" style={breit ? { gridColumn: '1 / -1' } : undefined}>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontWeight: 600, fontSize: '0.88rem', color: 'var(--dex-gray-800)', marginBottom: 6, lineHeight: 1.35, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  <span>{f.required && <span className="required" style={{ color: 'var(--dex-red, #da291c)', marginRight: 4 }}>*</span>}{label(f)}</span>
                  {text && !inline && <InfoTooltip text={text.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim()} />}
                </div>
                {text && inline && (
                  <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', lineHeight: 1.45, marginTop: -2, marginBottom: 6 }}
                    dangerouslySetInnerHTML={{ __html: renderFieldDescHtml(text) }} />
                )}
              </div>
              <div style={{ minWidth: 0 }}>{eingabe(f, fehler)}</div>
            </div>
          );
        })}
      </div>
      {versucht && fehlend.length > 0 && (
        <div className="dex-ui-callout dex-ui-callout--warn" style={{ margin: '0 0 12px' }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>
            {isDe ? 'Bitte fülle noch die Pflichtfragen aus: ' : 'Please answer the required questions: '}
            <strong>{fehlend.map(label).join(', ')}</strong>
          </span>
        </div>
      )}
      <div style={{ display: 'flex', gap: 8 }}>
        <button className="btn btn-primary" style={{ fontSize: '0.82rem' }} disabled={isSaving} onClick={speichern}>{saveLabel}</button>
        <button className="btn btn-secondary" style={{ fontSize: '0.82rem' }} onClick={onCancel}>{cancelLabel}</button>
      </div>
    </div>
  );
}

export default MyEventEditForm;
