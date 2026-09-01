/* ChangeLogModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 4516-4691 des Stands
 * vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { FileText, X } from '../../Icons';

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
    const actionColor = (a: string): string => {
      if (a.indexOf('Deleted') >= 0) return 'var(--dex-red, #c00)';
      if (a.indexOf('Created') >= 0) return 'var(--dex-green-dark)';
      if (a.indexOf('Cancelled') >= 0) return 'var(--dex-orange)';
      return 'var(--dex-gray-700)';
    };
    // v19.30 (Feature D): Details lesbar rendern. Bei ParticipantUpdated &
    // ähnlichen Aktionen steht im Details-JSON `{ changes: { Feld: { old, new } } }`.
    // Wir zeigen pro Feld eine „Feld: alt → neu"-Zeile statt rohes JSON. Bei
    // anderen/unstrukturierten Details fallen wir auf den Klartext zurück.
    const fmtVal = (v: unknown): string => {
      if (v === undefined || v === null || v === '') return '—';
      return String(v);
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
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {Object.keys(changes).map(field => {
              const c = changes[field] || {};
              return (
                <div key={field} style={{ fontFamily: 'inherit', fontSize: '0.78rem', lineHeight: 1.4 }}>
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
    return (
      <div
        style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}
        onClick={() => setShowChangeLogModal(false)}
      >
        <div
          className="card"
          style={{ width: '100%', maxWidth: 1200, maxHeight: '90vh', overflow: 'auto', padding: 24, borderRadius: 16, background: '#fff' }}
          onClick={e => e.stopPropagation()}
        >
          <div className="flex-between mb-16">
            <h3 style={{ margin: 0 }}>
              <FileText size={18} /> {isDe ? 'Audit-Log (DEX_ChangeLog)' : 'Audit log (DEX_ChangeLog)'}
            </h3>
            <button onClick={() => setShowChangeLogModal(false)} style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-500)' }}>
              <X size={20} />
            </button>
          </div>
          <p style={{ margin: '0 0 8px', fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
            {isDe
              ? <>Letzte <strong>{changeLogEntries.length}</strong> Einträge ({filtered.length} sichtbar). Schreibrechte: alle authentifizierten User; Leserechte: Organizer + Admin.</>
              : <>Last <strong>{changeLogEntries.length}</strong> entries ({filtered.length} visible). Write access: all authenticated users; read access: organizer + admin.</>}
          </p>
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.82rem', color: 'var(--dex-gray-700)', marginBottom: 12, cursor: 'pointer' }}>
            <input
              type="checkbox"
              checked={changeLogHideSelf}
              onChange={e => setChangeLogHideSelf(e.target.checked)}
            />
            {isDe
              ? 'Eigenaktionen der User ausblenden (nur Aktionen von Organizer/Admin anzeigen)'
              : 'Hide user self-actions (show only actions performed by organizer/admin)'}
          </label>
          {changeLogLoading && (
            <p style={{ textAlign: 'center', padding: 16, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
              {isDe ? 'Lade Einträge…' : 'Loading entries…'}
            </p>
          )}
          {!changeLogLoading && (
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem' }}>
                <thead style={{ position: 'sticky', top: 0, background: 'var(--dex-gray-50)' }}>
                  <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Datum' : 'Date'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Aktion' : 'Action'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Ziel' : 'Target'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Event' : 'Event'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Wer' : 'Actor'}</th>
                    <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Details' : 'Details'}</th>
                  </tr>
                  <tr style={{ borderBottom: '1px solid var(--dex-gray-200)', background: '#fff' }}>
                    <th style={{ padding: 4 }} />
                    <th style={{ padding: 4 }}>
                      <input value={changeLogFilterAction} onChange={e => setChangeLogFilterAction(e.target.value)} placeholder={isDe ? 'z.B. Deleted' : 'e.g. Deleted'} style={{ width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' }} />
                    </th>
                    <th style={{ padding: 4 }} />
                    <th style={{ padding: 4 }}>
                      <input value={changeLogFilterEvent} onChange={e => setChangeLogFilterEvent(e.target.value)} placeholder={isDe ? 'Event-/Ziel-Name' : 'event/target'} style={{ width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' }} />
                    </th>
                    <th style={{ padding: 4 }}>
                      <input value={changeLogFilterActor} onChange={e => setChangeLogFilterActor(e.target.value)} placeholder={isDe ? 'Name/E-Mail' : 'name/email'} style={{ width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' }} />
                    </th>
                    <th style={{ padding: 4 }} />
                  </tr>
                </thead>
                <tbody>
                  {filtered.map(e => (
                    <tr key={e.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                      <td style={{ padding: 6, color: 'var(--dex-gray-600)', whiteSpace: 'nowrap' }}>{fmtDate(e.Created)}</td>
                      <td style={{ padding: 6, color: actionColor(e.Action), fontWeight: 600 }}>{e.Action}</td>
                      <td style={{ padding: 6 }}>{e.TargetName || e.TargetId || '-'}</td>
                      <td style={{ padding: 6, color: 'var(--dex-gray-700)' }}>{e.EventTitle || '-'}</td>
                      <td style={{ padding: 6 }}>
                        {e.ActorName || e.ActorEmail || '-'}
                        {e.ActorEmail && e.ActorName && (
                          <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{e.ActorEmail}</div>
                        )}
                      </td>
                      <td style={{ padding: 6, color: 'var(--dex-gray-600)', fontSize: '0.75rem', maxWidth: 360, wordBreak: 'break-word' }}>{renderDetails(e.Details)}</td>
                    </tr>
                  ))}
                  {filtered.length === 0 && (
                    <tr><td colSpan={6} style={{ padding: 16, textAlign: 'center', color: 'var(--dex-gray-500)' }}>
                      {isDe ? 'Keine Einträge passen zum Filter.' : 'No entries match the filter.'}
                    </td></tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    );
};

