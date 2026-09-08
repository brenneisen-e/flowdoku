/* ChangeLogModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 4516-4691).
 *
 * v31.3 (docs/ui-leitfaden.md): Kopf, Backdrop und Kreuz macht jetzt `Modal`.
 * Die drei Filter standen als ZWEITE Kopfzeile IN der Tabelle — ein Feld in
 * einer Kopfzelle liest sich nicht als Filter, und die Aktion musste man
 * buchstabengenau tippen. Jetzt Werkzeugleiste über der Tabelle, Aktion als
 * Auswahl. Die Filter-LOGIK bleibt (Teilstring, kleingeschrieben) — auch die
 * Sammel-Einträge der Auswahl laufen genau darüber.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { ChevronDown, FileText, Search } from '../../Icons';
import { InfoTooltip } from '../../InfoTooltip';

export interface ChangeLogModalProps {
  changeLogEntries: { Id: number; Created: string; Action: string; TargetType: string; TargetId: string; TargetName: string; EventId: string; EventTitle: string; ActorName: string; ActorEmail: string; Details: string; }[];
  changeLogFilterAction: string;
  changeLogFilterActor: string;
  changeLogFilterEvent: string;
  changeLogHideSelf: boolean;
  changeLogLoading: boolean;
  isDe: boolean;
  setChangeLogFilterAction: React.Dispatch<React.SetStateAction<string>>;
  setChangeLogFilterActor: React.Dispatch<React.SetStateAction<string>>;
  setChangeLogFilterEvent: React.Dispatch<React.SetStateAction<string>>;
  setChangeLogHideSelf: React.Dispatch<React.SetStateAction<boolean>>;
  setShowChangeLogModal: React.Dispatch<React.SetStateAction<boolean>>;
}

export const ChangeLogModal: React.FC<ChangeLogModalProps> = (p) => {
  const { changeLogEntries, changeLogFilterAction, changeLogFilterActor, changeLogFilterEvent, changeLogHideSelf, changeLogLoading, isDe, setChangeLogFilterAction, setChangeLogFilterActor, setChangeLogFilterEvent, setChangeLogHideSelf, setShowChangeLogModal } = p;
  // v31.3: Welche Zeile zeigt ihre Feld-Änderungen? Bis zu zwölf „Feld:
  // alt → neu"-Zeilen in EINER Zelle machten die Tabelle unlesbar.
  const [openDetailsId, setOpenDetailsId] = React.useState<number>(null);
  const fa = changeLogFilterAction.toLowerCase().trim();
  const fe = changeLogFilterEvent.toLowerCase().trim();
  const fac = changeLogFilterActor.toLowerCase().trim();
  // Self-Action-Erkennung: Marker im Details-JSON ODER (als Fallback)
  // Actor-E-Mail == Target-E-Mail (User hat sich selbst registriert/abgemeldet).
  const isSelfAction = (e: typeof changeLogEntries[number]): boolean => {
    const d = (e.Details || '').toLowerCase();
    if (d.indexOf('"asactor":"self"') >= 0) return true;
    // Fallback: bei Participant-Aktionen ohne expliziten Marker prüfen wir,
    // ob Actor und Ziel dieselbe Person sind (Target trägt den Namen des
    // Participants, ActorName ist "Nachname, Vorname").
    const action = (e.Action || '').toLowerCase();
    if (action.indexOf('participant') < 0) return false;
    const tgt = (e.TargetName || '').toLowerCase().trim();
    const actorName = (e.ActorName || '').toLowerCase().trim();
    if (!tgt || !actorName) return false;
    // ActorName-Format "Nachname, Vorname" → in "Vorname Nachname" umdrehen
    const parts = actorName.split(',').map(s => s.trim());
    const flipped = parts.length === 2 ? `${parts[1]} ${parts[0]}` : actorName;
    return tgt === flipped || tgt === actorName;
  };
  const filtered = changeLogEntries.filter(e =>
    (!fa || (e.Action || '').toLowerCase().indexOf(fa) >= 0) &&
    (!fe || ((e.EventTitle || '').toLowerCase().indexOf(fe) >= 0 || (e.TargetName || '').toLowerCase().indexOf(fe) >= 0)) &&
    (!fac || (e.ActorName || '').toLowerCase().indexOf(fac) >= 0 || (e.ActorEmail || '').toLowerCase().indexOf(fac) >= 0) &&
    (!changeLogHideSelf || !isSelfAction(e))
  );
  const fmtDate = (iso: string): string => {
    if (!iso) return '';
    try { return new Date(iso).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); }
    catch { return iso; }
  };
  // v31.3: Aus der Textfarbe wird eine Statuspille — dieselbe Bedeutung (rot
  // gelöscht, grün angelegt, orange abgemeldet), aber sofort erkennbar.
  const actionPill = (a: string): string => {
    if (a.indexOf('Deleted') >= 0) return 'dex-ui-pill--red';
    if (a.indexOf('Created') >= 0) return 'dex-ui-pill--green';
    if (a.indexOf('Cancelled') >= 0) return 'dex-ui-pill--orange';
    return 'dex-ui-pill--gray';
  };
  // v31.3: Auswählen statt tippen — die Aktionen, die im geladenen Protokoll
  // wirklich vorkommen, davor die Sammel-Filter. Die zwei Bereichs-Einträge
  // (Teilnehmer, Events) sind Teilstrings wie alle anderen und halten damit
  // offen, nach einem Wortteil statt nach EINER Aktion zu filtern.
  const actionsInLog: string[] = [];
  changeLogEntries.forEach(e => { const a = (e.Action || '').trim(); if (a && actionsInLog.indexOf(a) < 0) actionsInLog.push(a); });
  actionsInLog.sort();
  const groups = [
    { v: 'Created', label: isDe ? 'Alles Angelegte' : 'Everything created' },
    { v: 'Updated', label: isDe ? 'Alles Geänderte' : 'Everything changed' },
    { v: 'Cancelled', label: isDe ? 'Alle Abmeldungen' : 'All cancellations' },
    { v: 'Deleted', label: isDe ? 'Alle Löschungen' : 'All deletions' },
    { v: 'Participant', label: isDe ? 'Alles zu Teilnehmern' : 'Everything about attendees' },
    { v: 'Event', label: isDe ? 'Alles zu Events' : 'Everything about events' },
  ].filter(g => actionsInLog.some(a => a.indexOf(g.v) >= 0));
  // Ein Filter, den die Auswahl nicht kennt (Protokoll neu geladen), bliebe
  // sonst unsichtbar wirksam — deshalb als eigener Eintrag.
  const knownAction = !changeLogFilterAction || actionsInLog.indexOf(changeLogFilterAction) >= 0 || groups.some(g => g.v === changeLogFilterAction);
  const anyFilter = !!(fa || fe || fac);
  const resetFilters = (): void => { setChangeLogFilterAction(''); setChangeLogFilterEvent(''); setChangeLogFilterActor(''); };
  // v19.30 (Feature D): Details lesbar rendern. Bei ParticipantUpdated &
  // ähnlichen Aktionen steht im Details-JSON `{ changes: { Feld: { old, new } } }`.
  // Wir zeigen pro Feld eine „Feld: alt → neu"-Zeile statt rohes JSON. Bei
  // anderen/unstrukturierten Details fallen wir auf den Klartext zurück.
  const fmtVal = (v: unknown): string => {
    if (v === undefined || v === null || v === '') return '—';
    return String(v);
  };
  // v31.3: Anzahl geänderter Felder — entscheidet, ob die Zeile aufklappt.
  const changeCount = (raw: string): number => {
    if (!raw) return 0;
    try { const c = (JSON.parse(raw) as { changes?: Record<string, unknown> }).changes; return c && typeof c === 'object' ? Object.keys(c).length : 0; }
    catch { return 0; }
  };
  const renderDetails = (raw: string): React.ReactNode => {
    if (!raw) return <span style={{ color: 'var(--dex-gray-400)' }}>—</span>;
    let parsed: unknown;
    try { parsed = JSON.parse(raw); } catch { return <span>{raw}</span>; }
    if (!parsed || typeof parsed !== 'object') return <span>{raw}</span>;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const obj = parsed as any;
    const changes = obj.changes;
    if (changes && typeof changes === 'object' && Object.keys(changes).length > 0) {
      return (
        <div className="dex-ui-stack" style={{ gap: 4 }}>
          {Object.keys(changes).map(field => {
            const c = changes[field] || {};
            return (
              <div key={field} style={{ fontSize: '0.78rem', lineHeight: 1.4 }}>
                <strong style={{ color: 'var(--dex-gray-800)' }}>{field}:</strong>{' '}
                <span style={{ color: 'var(--dex-gray-500)', textDecoration: 'line-through' }}>{fmtVal(c.old)}</span>
                {' → '}
                <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>{fmtVal(c.new)}</span>
              </div>
            );
          })}
        </div>
      );
    }
    // Kein changes-Block: übrige aussagekräftige Schlüssel kompakt zeigen
    // (z.B. asActor / via / scope), sonst das rohe JSON.
    const keys = Object.keys(obj).filter(k => k !== 'asActor');
    if (keys.length === 0) {
      return <span style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{isDe ? '(keine Detailänderungen)' : '(no detail changes)'}</span>;
    }
    return (
      <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>
        {keys.map(k => `${k}: ${fmtVal(obj[k])}`).join(' · ')}
      </span>
    );
  };
  // v31.3: Personen-Zelle des Leitfadens — Kürzel, Name, darunter die E-Mail.
  const initialsOf = (name: string, mail: string): string => {
    const ps = (name || '').split(/[\s,]+/).filter(Boolean);
    if (ps.length > 0) return (ps[0].charAt(0) + (ps[1] ? ps[1].charAt(0) : '')).toUpperCase();
    return (mail || '?').charAt(0).toUpperCase();
  };
  const resetBtn = (cls: string): React.ReactElement => (
    <button type="button" className={cls} onClick={resetFilters}>{isDe ? 'Filter zurücksetzen' : 'Reset filters'}</button>
  );
  // v31.3: Der Rechte-Satz steht in der Fußzeile UND im Warnkasten (dort ist
  // er die wahrscheinlichste Ursache) — eine Quelle, damit beide gleich lauten.
  const rightsNote = isDe
    ? 'Schreibrechte: alle angemeldeten Personen · Leserechte: Organizer und Admins.'
    : 'Write access: all signed-in users · read access: organizers and admins.';
  const countNote = isDe
    ? `${filtered.length} von ${changeLogEntries.length} geladenen Einträgen sichtbar`
    : `${filtered.length} of ${changeLogEntries.length} loaded entries shown`;
  return (
    <Modal
      open onClose={() => setShowChangeLogModal(false)} maxWidth={1200} ariaLabel={isDe ? 'Änderungsprotokoll' : 'Change log'}
      title={isDe ? 'Änderungsprotokoll (Audit-Log)' : 'Change log (audit log)'}
      subtitle={isDe ? 'Wer hat wann was an Events und Anmeldungen geändert? Neueste Einträge zuerst, aus der Liste DEX_ChangeLog.' : 'Who changed what, and when? Newest entries first, from the DEX_ChangeLog list.'}
      icon={<FileText size={20} />}
      footer={<button type="button" className="btn btn-secondary" onClick={() => setShowChangeLogModal(false)}>{isDe ? 'Schließen' : 'Close'}</button>}
    >
      {/* v31.3: Alle vier Filter links beieinander (Leitfaden 5a: rechts steht
          nur, was von der Liste WEG führt). Ohne geladene Einträge gibt es
          nichts zu filtern — dann bleibt die Leiste weg. */}
      {!changeLogLoading && changeLogEntries.length > 0 && (
      <div className="dex-ui-toolbar">
        <span className="dex-ui-searchbar">
          <span className="dex-ui-searchbar-icon"><Search size={14} /></span>
          <input className="dex-ui-input dex-ui-input--sm" value={changeLogFilterEvent} onChange={e => setChangeLogFilterEvent(e.target.value)} placeholder={isDe ? 'Event oder Ziel suchen…' : 'Search event or target…'} aria-label={isDe ? 'Nach Event oder Ziel suchen' : 'Search event or target'} />
        </span>
        <select className="dex-ui-select dex-ui-select--sm" style={{ width: 'auto', minWidth: 190 }} value={changeLogFilterAction} onChange={e => setChangeLogFilterAction(e.target.value)} aria-label={isDe ? 'Nach Aktion filtern' : 'Filter by action'}>
          <option value="">{isDe ? 'Alle Aktionen' : 'All actions'}</option>
          {!knownAction && <option value={changeLogFilterAction}>{changeLogFilterAction}</option>}
          {groups.length > 0 && <optgroup label={isDe ? 'Sammel-Filter' : 'Groups'}>{groups.map(g => <option key={g.v} value={g.v}>{g.label}</option>)}</optgroup>}
          {actionsInLog.length > 0 && <optgroup label={isDe ? 'Einzelne Aktionen' : 'Single actions'}>{actionsInLog.map(a => <option key={a} value={a}>{a}</option>)}</optgroup>}
        </select>
        <input className="dex-ui-input dex-ui-input--sm" style={{ width: 190 }} value={changeLogFilterActor} onChange={e => setChangeLogFilterActor(e.target.value)} placeholder={isDe ? 'Wer? Name oder E-Mail' : 'Who? name or email'} aria-label={isDe ? 'Nach handelnder Person filtern' : 'Filter by actor'} />
        {anyFilter && resetBtn('dex-ui-textbtn dex-ui-textbtn--muted')}
        <label className="dex-ui-switch">
          <input type="checkbox" checked={changeLogHideSelf} onChange={e => setChangeLogHideSelf(e.target.checked)} />
          <span className="dex-ui-switch-track" />
          <span className="dex-ui-switch-label">{isDe ? 'Eigenaktionen ausblenden' : 'Hide self-actions'}</span>
        </label>
        <InfoTooltip text={isDe ? 'Blendet aus, was Teilnehmer selbst getan haben (eigene Anmeldung, eigene Abmeldung). Übrig bleibt, was Organizer und Admins auf andere angewendet haben.' : 'Hides what attendees did themselves (own registration, own cancellation). What remains is what organizers and admins did to others.'} />
        <span className="dex-ui-toolbar-spacer" />
      </div>
      )}
      {changeLogLoading ? (
        <p className="dex-ui-muted" style={{ margin: 0, padding: '16px 0' }}>{isDe ? 'Einträge werden geladen…' : 'Loading entries…'}</p>
      ) : changeLogEntries.length === 0 ? (
        // v31.3: Leer ist hier KEINE Aussage — readChangeLog liefert auch bei
        // jedem Fehlschlag [] (CLAUDE.md: „ein Lesefehler ist keine Null").
        <div className="dex-ui-callout dex-ui-callout--warn">
          <span className="dex-ui-callout-icon"><FileText size={16} /></span>
          <div>
            {isDe ? 'Keine Einträge geladen. Entweder wurde noch nichts protokolliert — oder DEX_ChangeLog war nicht lesbar (fehlende Leserechte, Drosselung). Beides sieht hier gleich aus.' : 'No entries loaded. Either nothing has been logged yet — or DEX_ChangeLog could not be read (missing read access, throttling). Both look the same here.'}
            <div className="dex-ui-muted" style={{ marginTop: 6 }}>{rightsNote}</div>
          </div>
        </div>
      ) : filtered.length === 0 ? (
        <div className="dex-ui-empty">
          <span className="dex-ui-empty-icon"><Search size={20} /></span>
          <div className="dex-ui-empty-title">{isDe ? 'Keine Einträge passen zum Filter.' : 'No entries match the filter.'}</div>
          {anyFilter && resetBtn('dex-ui-textbtn')}
        </div>
      ) : (
        <div className="dex-ui-table-wrap dex-ui-table-wrap--sticky">
          <table className="dex-ui-table dex-ui-table--compact">
            {/* Spaltenfolge Kopf = Zeile: Wann · Aktion · Ziel · Event · Wer · Details */}
            <thead>
              <tr>
                <th>{isDe ? 'Wann' : 'When'}</th><th>{isDe ? 'Aktion' : 'Action'}</th><th>{isDe ? 'Ziel' : 'Target'}</th>
                <th>{isDe ? 'Event' : 'Event'}</th><th>{isDe ? 'Wer' : 'Who'}</th><th>{isDe ? 'Details' : 'Details'}</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(e => {
                const n = changeCount(e.Details);
                const open = openDetailsId === e.Id;
                return (
                  <React.Fragment key={e.Id}>
                    <tr>
                      <td style={{ color: 'var(--dex-gray-600)', whiteSpace: 'nowrap' }}>{fmtDate(e.Created)}</td>
                      <td><span className={cx('dex-ui-pill', actionPill(e.Action))}>{e.Action}</span></td>
                      <td>{e.TargetName || e.TargetId || '—'}</td>
                      <td style={{ color: 'var(--dex-gray-700)' }}>{e.EventTitle || '—'}</td>
                      <td><span className="dex-ui-person">
                        <span className="dex-ui-avatar" aria-hidden="true">{initialsOf(e.ActorName, e.ActorEmail)}</span>
                        <span style={{ minWidth: 0 }}>
                          <span className="dex-ui-person-name" style={{ display: 'block' }}>{e.ActorName || e.ActorEmail || '—'}</span>
                          {e.ActorEmail && e.ActorName && <span className="dex-ui-person-sub" style={{ display: 'block' }}>{e.ActorEmail}</span>}
                        </span>
                      </span></td>
                      <td style={{ maxWidth: 360, wordBreak: 'break-word' }}>{n > 0 ? (
                        <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" aria-expanded={open} onClick={() => setOpenDetailsId(open ? null : e.Id)}>
                          <span className={cx('dex-ui-disclosure-chevron', open && 'is-open')}><ChevronDown size={14} /></span>
                          {n} {isDe ? (n === 1 ? 'Änderung' : 'Änderungen') : (n === 1 ? 'change' : 'changes')}
                        </button>
                      ) : renderDetails(e.Details)}</td>
                    </tr>
                    {open && <tr className="dex-ui-fade-in"><td colSpan={6} style={{ background: 'var(--dex-gray-50, #fafafa)' }}>{renderDetails(e.Details)}</td></tr>}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {/* Auch bei leerem Filter-Ergebnis: „0 von N" beantwortet „warum leer?". */}
      {!changeLogLoading && changeLogEntries.length > 0 && (
        <div className="dex-ui-table-foot">
          <span>{countNote}</span>
          <span>{rightsNote}</span>
        </div>
      )}
    </Modal>
  );
};
