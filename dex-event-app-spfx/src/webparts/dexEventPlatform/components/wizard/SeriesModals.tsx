/* v31.99: Die zwei Rückfragen der Serien-Termine.
 *
 * SeriesApplyModal — die Serie wurde geändert (z.B. verlängert): Welche
 * Termine kommen dazu, welche fallen weg, wessen Uhrzeit wird angepasst?
 * Jede der drei Folgen ist einzeln abwählbar; nichts passiert, bevor der
 * Organizer bestätigt, und gespeichert wird ohnehin erst mit „Speichern".
 *
 * SeriesPropagateModal — ein einzelner Termin wurde geändert: Sollen die
 * anderen Termine das übernehmen (alle oder nur die folgenden)? Gefragt wird
 * beim Verlassen des Termin-Reiters und vor dem Speichern. */
import * as React from 'react';
import Modal from '../Modal';
import { cx } from '../dexUi';
import { RefreshCw, Copy, AlertCircle } from '../Icons';
import { SubEventDraft } from './wizardTypes';
import { SeriesRule } from '../../types';
import { SeriesPlan, seriesSummary } from '../../utils/seriesRule';
import { ApplySeriesChoice, SERIES_PROPAGATE_GROUPS } from './logic/seriesActions';

const fmtKey = (k: string, isDe: boolean): string => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(k || '');
  if (!m) return k;
  const d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  return d.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { weekday: 'short', day: '2-digit', month: '2-digit', year: '2-digit' });
};

const DayList: React.FC<{ days: string[]; isDe: boolean }> = ({ days, isDe }) => (
  <div className="dex-ui-inline" style={{ gap: 4, flexWrap: 'wrap', marginTop: 6 }}>
    {days.slice(0, 24).map(k => <span key={k} className="dex-ui-pill dex-ui-pill--gray dex-ui-pill--sm">{fmtKey(k, isDe)}</span>)}
    {days.length > 24 && <span className="dex-ui-muted">{isDe ? `und ${days.length - 24} weitere` : `and ${days.length - 24} more`}</span>}
  </div>
);

const CheckRow: React.FC<{ checked: boolean; onChange: (_v: boolean) => void; title: React.ReactNode; desc?: React.ReactNode; children?: React.ReactNode }> = ({ checked, onChange, title, desc, children }) => (
  <label className={cx('dex-ui-toggle-row', checked && 'is-active')} style={{ alignItems: 'flex-start' }}>
    <input type="checkbox" checked={checked} onChange={e => onChange(e.target.checked)} />
    <span className="dex-ui-toggle-row-body" style={{ minWidth: 0 }}>
      <span className="dex-ui-toggle-row-title">{title}</span>
      {desc && <span className="dex-ui-toggle-row-desc">{desc}</span>}
      {children}
    </span>
  </label>
);

export interface SeriesApplyState {
  rule: SeriesRule;
  plan: SeriesPlan<SubEventDraft>;
  /** Bestehender Termin, dessen Einstellungen neue Termine übernehmen können. */
  templateLabel: string;
  dayKeyOf: (s: SubEventDraft) => string;
}

export const SeriesApplyModal: React.FC<{
  state: SeriesApplyState | null;
  isDe: boolean;
  onCancel: () => void;
  onConfirm: (_choice: ApplySeriesChoice) => void;
}> = ({ state, isDe, onCancel, onConfirm }) => {
  const [choice, setChoice] = React.useState<ApplySeriesChoice>({ add: true, remove: true, retime: true, restore: false, copyTemplate: true });
  React.useEffect(() => {
    if (state) setChoice({ add: true, remove: true, retime: true, restore: false, copyTemplate: !!state.templateLabel });
  }, [state]);
  if (!state) return null;
  const { plan, rule, dayKeyOf } = state;
  const set = (patch: Partial<ApplySeriesChoice>): void => setChoice(prev => ({ ...prev, ...patch }));
  const savedRemovals = plan.remove.filter(s => !!s.dbId).length;
  const nothing = plan.add.length === 0 && plan.remove.length === 0 && plan.retime.length === 0 && plan.restorable.length === 0;
  const time = rule.allDay ? (isDe ? 'ganztägig' : 'all day') : `${rule.startTime}–${rule.endTime}`;
  return (
    <Modal
      open
      onClose={onCancel}
      maxWidth={620}
      backdropClose={false}
      ariaLabel={isDe ? 'Serie geändert' : 'Series changed'}
      icon={<RefreshCw size={18} />}
      title={isDe ? 'Serie geändert — Termine anpassen?' : 'Series changed — adjust dates?'}
      subtitle={seriesSummary(rule, isDe)}
      footer={(
        <>
          <button type="button" className="btn btn-secondary" onClick={onCancel}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
          <button type="button" className="btn btn-primary" onClick={() => onConfirm(choice)}>{isDe ? 'Übernehmen' : 'Apply'}</button>
        </>
      )}
    >
      <div className="dex-ui-stack" style={{ gap: 10 }}>
        {nothing && (
          <p className="dex-ui-help" style={{ margin: 0 }}>
            {isDe
              ? 'Die bestehenden Termine passen bereits zur geänderten Serie. Übernommen wird nur die neue Regel.'
              : 'The existing dates already match the changed series. Only the new rule is stored.'}
          </p>
        )}
        {plan.add.length > 0 && (
          <CheckRow
            checked={choice.add}
            onChange={v => set({ add: v })}
            title={isDe ? `${plan.add.length} neue ${plan.add.length === 1 ? 'Termin' : 'Termine'} anlegen` : `Create ${plan.add.length} new ${plan.add.length === 1 ? 'date' : 'dates'}`}
            desc={isDe ? `Uhrzeit ${time}.` : `Time ${time}.`}
          >
            <DayList days={plan.add} isDe={isDe} />
          </CheckRow>
        )}
        {plan.add.length > 0 && choice.add && !!state.templateLabel && (
          <CheckRow
            checked={choice.copyTemplate}
            onChange={v => set({ copyTemplate: v })}
            title={isDe ? `Neue Termine übernehmen die Einstellungen von „${state.templateLabel}"` : `New dates take the settings of "${state.templateLabel}"`}
            desc={isDe
              ? 'Beschreibung, Bild, Ort, Plätze, Programm, Formularfragen und Sichtbarkeit. Ohne Haken starten die neuen Termine leer.'
              : 'Description, image, location, seats, agenda, form questions and visibility. Unticked, the new dates start empty.'}
          />
        )}
        {plan.retime.length > 0 && (
          <CheckRow
            checked={choice.retime}
            onChange={v => set({ retime: v })}
            title={isDe ? `Uhrzeit bei ${plan.retime.length} bestehenden ${plan.retime.length === 1 ? 'Termin' : 'Terminen'} auf ${time} setzen` : `Set the time of ${plan.retime.length} existing ${plan.retime.length === 1 ? 'date' : 'dates'} to ${time}`}
            desc={isDe ? 'Nur künftige Termine. Bei gespeicherten Terminen fragt der Assistent beim Speichern, ob die Teilnehmer ein Outlook-Update bekommen.' : 'Future dates only. For saved dates the wizard asks on save whether attendees get an Outlook update.'}
          >
            <DayList days={plan.retime.map(dayKeyOf)} isDe={isDe} />
          </CheckRow>
        )}
        {plan.remove.length > 0 && (
          <CheckRow
            checked={choice.remove}
            onChange={v => set({ remove: v })}
            title={isDe ? `${plan.remove.length} ${plan.remove.length === 1 ? 'Termin entfällt' : 'Termine entfallen'} — entfernen` : `${plan.remove.length} ${plan.remove.length === 1 ? 'date no longer matches' : 'dates no longer match'} — remove`}
            desc={savedRemovals > 0
              ? (isDe
                ? `${savedRemovals} davon ${savedRemovals === 1 ? 'ist' : 'sind'} schon gespeichert und ${savedRemovals === 1 ? 'wird' : 'werden'} beim Speichern samt Teilnehmerliste gelöscht (93 Tage im Papierkorb). Bis dahin im Kalender orange und wiederherstellbar.`
                : `${savedRemovals} of them ${savedRemovals === 1 ? 'is' : 'are'} already saved and will be deleted with the attendee list on save (93 days in the recycle bin). Until then orange in the calendar and restorable.`)
              : undefined}
          >
            <DayList days={plan.remove.map(dayKeyOf)} isDe={isDe} />
          </CheckRow>
        )}
        {plan.restorable.length > 0 && (
          <CheckRow
            checked={choice.restore}
            onChange={v => set({ restore: v })}
            title={isDe ? `${plan.restorable.length} früher einzeln ${plan.restorable.length === 1 ? 'gelöschten Termin' : 'gelöschte Termine'} wieder anlegen` : `Re-create ${plan.restorable.length} previously removed ${plan.restorable.length === 1 ? 'date' : 'dates'}`}
            desc={isDe ? 'Ohne Haken bleiben diese Tage Ausnahmen der Serie.' : 'Unticked, these days stay exceptions of the series.'}
          >
            <DayList days={plan.restorable} isDe={isDe} />
          </CheckRow>
        )}
        {plan.outside > 0 && (
          <p className="dex-ui-help" style={{ margin: 0 }}>
            {isDe
              ? `${plan.outside} ${plan.outside === 1 ? 'Termin liegt' : 'Termine liegen'} außerhalb der Regel (vergangen oder von Hand angelegt) und ${plan.outside === 1 ? 'bleibt' : 'bleiben'} unverändert.`
              : `${plan.outside} ${plan.outside === 1 ? 'date is' : 'dates are'} outside the rule (past or added by hand) and stay unchanged.`}
          </p>
        )}
        <p className="dex-ui-help" style={{ margin: 0 }}>
          {isDe ? 'Vergangene Termine werden nie verändert. Gespeichert wird erst mit „Speichern".' : 'Past dates are never changed. Nothing is saved until you click “Save”.'}
        </p>
      </div>
    </Modal>
  );
};

export interface SeriesPropagateState {
  subId: string;
  terminLabel: string;
  groups: string[];
  allCount: number;
  followingCount: number;
  /** Nach der Entscheidung speichern (Frage kam aus „Speichern"). */
  thenSubmit: boolean;
}

export const SeriesPropagateModal: React.FC<{
  state: SeriesPropagateState | null;
  isDe: boolean;
  onDecide: (_apply: { groups: string[]; target: 'all' | 'following' } | null) => void;
}> = ({ state, isDe, onDecide }) => {
  const [sel, setSel] = React.useState<string[]>([]);
  const [target, setTarget] = React.useState<'all' | 'following'>('all');
  React.useEffect(() => {
    if (!state) return;
    setSel(state.groups.filter(k => { const g = SERIES_PROPAGATE_GROUPS.find(x => x.key === k); return !!g && g.defaultOn; }));
    setTarget('all');
  }, [state]);
  if (!state) return null;
  const n = target === 'all' ? state.allCount : state.followingCount;
  return (
    <Modal
      open
      onClose={() => onDecide(null)}
      maxWidth={560}
      backdropClose={false}
      ariaLabel={isDe ? 'Änderung auf andere Termine übertragen' : 'Apply change to other dates'}
      icon={<Copy size={18} />}
      title={isDe ? 'Auch für die anderen Termine übernehmen?' : 'Apply to the other dates as well?'}
      subtitle={isDe ? `Du hast den Termin „${state.terminLabel}" geändert.` : `You changed the date "${state.terminLabel}".`}
      footer={(
        <>
          <button type="button" className="btn btn-secondary" onClick={() => onDecide(null)}>{isDe ? 'Nur dieser Termin' : 'This date only'}</button>
          <button type="button" className="btn btn-primary" disabled={sel.length === 0 || n === 0} onClick={() => onDecide({ groups: sel, target })}>
            {isDe ? `Für ${n} ${n === 1 ? 'Termin' : 'Termine'} übernehmen` : `Apply to ${n} ${n === 1 ? 'date' : 'dates'}`}
          </button>
        </>
      )}
    >
      <div className="dex-ui-stack" style={{ gap: 10 }}>
        <div className="dex-ui-label" style={{ margin: 0 }}>{isDe ? 'Was soll übernommen werden?' : 'What should be applied?'}</div>
        {state.groups.map(k => {
          const g = SERIES_PROPAGATE_GROUPS.find(x => x.key === k);
          if (!g) return null;
          const on = sel.indexOf(k) >= 0;
          return (
            <CheckRow
              key={k}
              checked={on}
              onChange={v => setSel(prev => v ? [...prev, k] : prev.filter(x => x !== k))}
              title={isDe ? g.de : g.en}
              desc={k === 'times'
                ? (isDe ? 'Nur die Uhrzeit — jeder Termin bleibt an seinem Tag.' : 'Only the time of day — each date keeps its day.')
                : k === 'title'
                  ? (isDe ? 'Alle Termine hießen danach gleich; üblich ist der Tag als Titel.' : 'All dates would then share one title; usually the day is the title.')
                  : undefined}
            />
          );
        })}
        <div className="dex-ui-label" style={{ margin: '4px 0 0' }}>{isDe ? 'Für welche Termine?' : 'For which dates?'}</div>
        <div className="dex-ui-inline" style={{ gap: 6, flexWrap: 'wrap' }}>
          <button type="button" className={cx('dex-ui-chip', target === 'all' && 'is-active')} aria-pressed={target === 'all'} onClick={() => setTarget('all')}>
            {isDe ? `Alle anderen (${state.allCount})` : `All others (${state.allCount})`}
          </button>
          <button type="button" className={cx('dex-ui-chip', target === 'following' && 'is-active')} aria-pressed={target === 'following'} disabled={state.followingCount === 0} onClick={() => setTarget('following')}>
            {isDe ? `Nur die folgenden (${state.followingCount})` : `Following only (${state.followingCount})`}
          </button>
        </div>
        {state.thenSubmit && (
          <div className="dex-ui-callout dex-ui-callout--info dex-ui-callout--sm">
            <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
            <div>{isDe ? 'Danach wird gespeichert.' : 'Saving continues afterwards.'}</div>
          </div>
        )}
      </div>
    </Modal>
  );
};
