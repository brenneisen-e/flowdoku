/* v31.99: „Als Serie anlegen" — Eingabe der Wiederholungsregel in Schritt 1,
 * direkt unter dem Zeitraum (Nutzer-Ansage 28.09.2026: beim Datum ankreuzen,
 * dann werden Serientermine angelegt).
 *
 * Die Maske hält einen ENTWURF der Regel. Erst „Termine anlegen" bzw.
 * „Änderung übernehmen" gibt ihn an die Seite (`onApply`), die den Abgleich
 * rechnet und bei bestehenden Terminen nachfragt. So löst nicht jeder
 * Tastendruck in „alle [2] Wochen" ein Anlegen und Löschen von Terminen aus,
 * und der Organizer sieht vor der Entscheidung, was passieren wird.
 *
 * Datum über react-datepicker (dd.MM.yyyy), Zeit als Text mit
 * `normalizeTimeInput` — keine nativen Date/Time-Felder (CLAUDE.md, v30.94). */
import * as React from 'react';
import DatePicker from 'react-datepicker';
import { SeriesRule } from '../../types';
import { cx } from '../dexUi';
import { normalizeTimeInput } from './AgendaEditor';
import { Check, AlertCircle, Info, RefreshCw } from '../Icons';
import { seriesDates, seriesRuleError, seriesSummary, SERIES_MAX, defaultSeriesRule } from '../../utils/seriesRule';

export interface SeriesEditorProps {
  isDe: boolean;
  /** Schalter „Als Serie anlegen". */
  on: boolean;
  onToggle: (on: boolean) => void;
  /** Zuletzt angewendete Regel (null = noch keine Termine aus der Serie). */
  applied: SeriesRule | null;
  onApply: (rule: SeriesRule) => void;
  /** Vorbelegung aus dem Zeitraum des Hauptevents (Berliner Lokalzeit). */
  startDate: string;
  endDate: string;
  allDay: boolean;
  /** Anzahl der vorhandenen Termine (Sub-Events). */
  terminCount: number;
  /** Springt zur Terminliste unter dem Kalender. */
  onShowList: () => void;
}

const keyToDate = (k: string): Date | null => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(k || '');
  return m ? new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10)) : null;
};
const dateToKey = (d: Date | null): string => (d
  ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
  : '');
const fmtKey = (k: string, isDe: boolean): string => {
  const d = keyToDate(k);
  return d ? d.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { weekday: 'short', day: '2-digit', month: '2-digit' }) : k;
};
/** Vergleich ohne Ausnahmen — die pflegt der Abgleich, nicht die Maske. */
const ruleKey = (r: SeriesRule | null): string => {
  if (!r) return '';
  const { exceptions, ...rest } = r;
  void exceptions;
  return JSON.stringify({ ...rest, weekdays: (rest.weekdays || []).slice().sort() });
};

/** Uhrzeit als Text; übernommen wird beim Verlassen des Felds. */
const TimeField: React.FC<{ value: string; onChange: (_v: string) => void; label: string; isDe: boolean }> = ({ value, onChange, label, isDe }) => {
  const [txt, setTxt] = React.useState(value);
  const [bad, setBad] = React.useState(false);
  React.useEffect(() => { setTxt(value); setBad(false); }, [value]);
  return (
    <div className="dex-ui-field" style={{ minWidth: 110 }}>
      <label className="dex-ui-label">{label}</label>
      <input
        type="text"
        inputMode="numeric"
        className={cx('dex-ui-input', bad && 'dex-ui-input--error')}
        value={txt}
        placeholder="09:00"
        onChange={e => setTxt(e.target.value)}
        onBlur={() => {
          const n = normalizeTimeInput(txt);
          if (n === null || n === '') { setBad(true); return; }
          setBad(false); setTxt(n); onChange(n);
        }}
        aria-invalid={bad}
        title={isDe ? 'z.B. 9, 930 oder 09:30' : 'e.g. 9, 930 or 09:30'}
      />
    </div>
  );
};

export const SeriesEditor: React.FC<SeriesEditorProps> = (p) => {
  const { isDe, on, onToggle, applied, onApply, startDate, endDate, allDay, terminCount, onShowList } = p;
  const makeDefault = (): SeriesRule => defaultSeriesRule(
    (startDate || '').slice(0, 10) || dateToKey(new Date()),
    (startDate || '').slice(11, 16), (endDate || '').slice(11, 16), allDay,
  );
  const [draft, setDraft] = React.useState<SeriesRule>(() => applied || makeDefault());
  // Wird die Regel von außen neu gesetzt (angewendet, Entwurf geladen,
  // Serie beendet), folgt die Maske — sonst zeigte sie einen Stand, der
  // nirgends mehr gilt.
  const appliedKey = ruleKey(applied);
  React.useEffect(() => {
    if (applied) setDraft(applied);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [appliedKey]);
  // Noch keine Serie angewendet: Beim Einschalten den aktuellen Zeitraum des
  // Hauptevents als Vorbelegung nehmen, nicht den vom ersten Render.
  React.useEffect(() => {
    if (on && !applied) setDraft(makeDefault());
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [on]);

  const set = (patch: Partial<SeriesRule>): void => setDraft(prev => ({ ...prev, ...patch }));
  const err = seriesRuleError(draft, isDe);
  const preview = React.useMemo(() => seriesDates(draft), [draft]);
  const dirty = !!applied && ruleKey(applied) !== ruleKey(draft);
  const unit = draft.freq === 'daily'
    ? (isDe ? (draft.interval === 1 ? 'Tag' : 'Tage') : (draft.interval === 1 ? 'day' : 'days'))
    : draft.freq === 'weekly'
      ? (isDe ? (draft.interval === 1 ? 'Woche' : 'Wochen') : (draft.interval === 1 ? 'week' : 'weeks'))
      : (isDe ? (draft.interval === 1 ? 'Monat' : 'Monate') : (draft.interval === 1 ? 'month' : 'months'));
  // Mo … So — deutscher Kalender, Montag zuerst.
  const weekdayOrder = [1, 2, 3, 4, 5, 6, 0];
  const wdName = (d: number): string => (isDe ? ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'] : ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'])[d];

  return (
    <div style={{ marginTop: 12 }}>
      <label className={cx('dex-ui-toggle-row', on && 'is-active')}>
        <input type="checkbox" checked={on} onChange={e => onToggle(e.target.checked)} />
        <span className="dex-ui-toggle-row-body">
          <span className="dex-ui-toggle-row-title">{isDe ? 'Als Serie anlegen' : 'Create as a series'}</span>
          <span className="dex-ui-toggle-row-desc">
            {isDe
              ? 'Wiederkehrende Termine (z.B. jeden Dienstag) automatisch anlegen. Jeder Termin bleibt einzeln änderbar und löschbar.'
              : 'Create recurring dates (e.g. every Tuesday) automatically. Each date stays editable and removable on its own.'}
          </span>
        </span>
      </label>

      {on && (
        <div className="dex-ui-card" style={{ marginTop: 10, padding: '14px 16px' }}>
          {applied && (
            <div className="dex-ui-callout dex-ui-callout--success dex-ui-callout--sm" style={{ marginBottom: 12 }}>
              <span className="dex-ui-callout-icon"><Check size={16} /></span>
              <div style={{ flex: 1, minWidth: 0 }}>
                <strong>{isDe ? 'Aktive Serie:' : 'Active series:'}</strong> {seriesSummary(applied, isDe)} · {isDe
                  ? `${terminCount} ${terminCount === 1 ? 'Termin' : 'Termine'} angelegt.`
                  : `${terminCount} ${terminCount === 1 ? 'date' : 'dates'} created.`}{' '}
                <button type="button" className="dex-ui-textbtn" style={{ display: 'inline-flex', padding: 0 }} onClick={onShowList}>
                  {isDe ? 'Einzelne Termine bearbeiten oder löschen' : 'Edit or remove single dates'}
                </button>
              </div>
            </div>
          )}

          {/* Wie oft? */}
          <div className="dex-ui-label">{isDe ? 'Wie oft wiederholt sich der Termin?' : 'How often does it repeat?'}</div>
          <div className="dex-ui-inline" style={{ gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
            {([
              ['daily', isDe ? 'Täglich' : 'Daily'],
              ['weekly', isDe ? 'Wöchentlich' : 'Weekly'],
              ['monthly', isDe ? 'Monatlich' : 'Monthly'],
            ] as Array<[SeriesRule['freq'], string]>).map(([f, lbl]) => (
              <button key={f} type="button" className={cx('dex-ui-chip', draft.freq === f && 'is-active')} aria-pressed={draft.freq === f} onClick={() => set({ freq: f })}>
                {lbl}
              </button>
            ))}
            <span className="dex-ui-muted" style={{ marginLeft: 8 }}>{isDe ? 'alle' : 'every'}</span>
            <input
              type="number"
              min={1}
              max={52}
              className="dex-ui-input dex-ui-input--sm"
              style={{ width: 70 }}
              value={draft.interval}
              onChange={e => set({ interval: Math.max(1, Math.min(52, parseInt(e.target.value, 10) || 1)) })}
              aria-label={isDe ? 'Abstand' : 'Interval'}
            />
            <span className="dex-ui-muted">{unit}</span>
          </div>

          {draft.freq === 'weekly' && (
            <div style={{ marginTop: 10 }}>
              <div className="dex-ui-label">{isDe ? 'An welchen Wochentagen?' : 'On which weekdays?'}</div>
              <div className="dex-ui-inline" style={{ gap: 6, flexWrap: 'wrap' }}>
                {weekdayOrder.map(d => {
                  const act = (draft.weekdays || []).indexOf(d) >= 0;
                  return (
                    <button
                      key={d}
                      type="button"
                      className={cx('dex-ui-chip', act && 'is-active')}
                      aria-pressed={act}
                      onClick={() => set({ weekdays: act ? (draft.weekdays || []).filter(x => x !== d) : [...(draft.weekdays || []), d] })}
                    >
                      {wdName(d)}
                    </button>
                  );
                })}
              </div>
            </div>
          )}
          {draft.freq === 'daily' && (
            <label className={cx('dex-ui-toggle-row', draft.skipWeekends && 'is-active')} style={{ marginTop: 10 }}>
              <input type="checkbox" checked={!!draft.skipWeekends} onChange={e => set({ skipWeekends: e.target.checked })} />
              <span className="dex-ui-toggle-row-body">
                <span className="dex-ui-toggle-row-title">{isDe ? 'Nur Werktage (ohne Samstag und Sonntag)' : 'Weekdays only (no Saturday or Sunday)'}</span>
              </span>
            </label>
          )}
          {draft.freq === 'monthly' && (
            <p className="dex-ui-help" style={{ marginTop: 8 }}>
              {isDe
                ? 'Immer am selben Tag des Monats wie der erste Termin. Gibt es den Tag in einem Monat nicht (z.B. 31.), entfällt der Termin dort.'
                : 'Always on the same day of the month as the first date. If a month has no such day (e.g. the 31st), there is no date that month.'}
            </p>
          )}

          <div className="dex-ui-grid-2" style={{ marginTop: 12 }}>
            <div className="dex-ui-field">
              <label className="dex-ui-label">{isDe ? 'Erster Termin' : 'First date'}</label>
              <DatePicker
                selected={keyToDate(draft.start)}
                onChange={(d: Date | null) => set({ start: dateToKey(d) })}
                dateFormat="dd.MM.yyyy"
                locale="de"
                className="form-input"
                wrapperClassName="dex-datepicker-wrapper"
                calendarClassName="dex-datepicker-calendar"
                popperPlacement="bottom-start"
                autoComplete="off"
              />
            </div>
            <div className="dex-ui-field">
              <label className="dex-ui-label">{isDe ? 'Wann endet die Serie?' : 'When does the series end?'}</label>
              <div className="dex-ui-inline" style={{ gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
                <button type="button" className={cx('dex-ui-chip', draft.endMode === 'count' && 'is-active')} aria-pressed={draft.endMode === 'count'} onClick={() => set({ endMode: 'count' })}>
                  {isDe ? 'Nach Anzahl' : 'After a number'}
                </button>
                <button type="button" className={cx('dex-ui-chip', draft.endMode === 'until' && 'is-active')} aria-pressed={draft.endMode === 'until'} onClick={() => set({ endMode: 'until' })}>
                  {isDe ? 'Am Datum' : 'On a date'}
                </button>
              </div>
              <div style={{ marginTop: 6 }}>
                {draft.endMode === 'count' ? (
                  <div className="dex-ui-inline" style={{ gap: 6, alignItems: 'center' }}>
                    <input
                      type="number"
                      min={1}
                      max={SERIES_MAX}
                      className="dex-ui-input dex-ui-input--sm"
                      style={{ width: 80 }}
                      value={draft.count || ''}
                      onChange={e => set({ count: Math.max(0, Math.min(SERIES_MAX, parseInt(e.target.value, 10) || 0)) })}
                      aria-label={isDe ? 'Anzahl Termine' : 'Number of dates'}
                    />
                    <span className="dex-ui-muted">{isDe ? 'Termine' : 'dates'}</span>
                  </div>
                ) : (
                  <DatePicker
                    selected={keyToDate(draft.until || '')}
                    onChange={(d: Date | null) => set({ until: dateToKey(d) })}
                    dateFormat="dd.MM.yyyy"
                    locale="de"
                    minDate={keyToDate(draft.start) || undefined}
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    autoComplete="off"
                  />
                )}
              </div>
            </div>
          </div>

          <label className={cx('dex-ui-toggle-row', draft.allDay && 'is-active')} style={{ marginTop: 12 }}>
            <input type="checkbox" checked={!!draft.allDay} onChange={e => set({ allDay: e.target.checked })} />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">{isDe ? 'Ganztägige Termine' : 'All-day dates'}</span>
            </span>
          </label>
          {!draft.allDay && (
            <div className="dex-ui-inline" style={{ gap: 12, marginTop: 8, flexWrap: 'wrap', alignItems: 'flex-start' }}>
              <TimeField isDe={isDe} label={isDe ? 'Beginn' : 'Start'} value={draft.startTime} onChange={v => set({ startTime: v })} />
              <TimeField isDe={isDe} label={isDe ? 'Ende' : 'End'} value={draft.endTime} onChange={v => set({ endTime: v })} />
            </div>
          )}

          {/* Vorschau — der Organizer sieht VOR dem Anlegen, welche Tage entstehen. */}
          {err ? (
            <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" style={{ marginTop: 12 }}>
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <div>{err}</div>
            </div>
          ) : (
            <div className="dex-ui-callout dex-ui-callout--info dex-ui-callout--sm" style={{ marginTop: 12 }}>
              <span className="dex-ui-callout-icon"><Info size={16} /></span>
              <div style={{ flex: 1, minWidth: 0 }}>
                <strong>{seriesSummary(draft, isDe)}</strong>{' — '}
                {isDe
                  ? `${preview.dates.length} ${preview.dates.length === 1 ? 'Termin' : 'Termine'}`
                  : `${preview.dates.length} ${preview.dates.length === 1 ? 'date' : 'dates'}`}
                {preview.dates.length > 0 && (
                  <>: {preview.dates.length <= 8
                    ? preview.dates.map(k => fmtKey(k, isDe)).join(', ')
                    : `${preview.dates.slice(0, 5).map(k => fmtKey(k, isDe)).join(', ')} … ${fmtKey(preview.dates[preview.dates.length - 1], isDe)}`}</>
                )}
                {preview.truncated && (
                  <div style={{ marginTop: 4 }}>
                    {isDe
                      ? `Begrenzt auf ${SERIES_MAX} Termine — jeder Termin bekommt eine eigene Teilnehmerliste.`
                      : `Limited to ${SERIES_MAX} dates — each date gets its own attendee list.`}
                  </div>
                )}
              </div>
            </div>
          )}

          <div className="dex-ui-inline" style={{ marginTop: 12, gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
            <button
              type="button"
              className="btn btn-primary dex-ui-btn-sm"
              disabled={!!err || preview.dates.length === 0 || (!!applied && !dirty)}
              onClick={() => onApply({ ...draft, exceptions: applied ? (applied.exceptions || []) : [] })}
            >
              <RefreshCw size={14} /> {applied
                ? (isDe ? 'Änderung übernehmen' : 'Apply change')
                : (isDe ? `${preview.dates.length} Termine anlegen` : `Create ${preview.dates.length} dates`)}
            </button>
            {dirty && (
              <>
                <span className="dex-ui-muted">{isDe ? 'Änderung noch nicht übernommen.' : 'Change not applied yet.'}</span>
                <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => applied && setDraft(applied)}>
                  {isDe ? 'Verwerfen' : 'Discard'}
                </button>
              </>
            )}
          </div>
          <p className="dex-ui-help" style={{ marginTop: 10 }}>
            {isDe
              ? 'Jeder Termin wird ein eigenes Event mit eigener Teilnehmerliste, eigenen Plätzen und eigenem Outlook-Termin. Teilnehmer wählen ihre Termine auf der Anmeldeseite im Kalender. Ändert sich die Serie später (z.B. Verlängerung), fragt der Assistent, welche Termine angepasst werden sollen.'
              : 'Each date becomes its own event with its own attendee list, seats and Outlook entry. Attendees pick their dates in a calendar on the registration page. If the series changes later (e.g. extended), the wizard asks which dates should be adjusted.'}
          </p>
        </div>
      )}
    </div>
  );
};
