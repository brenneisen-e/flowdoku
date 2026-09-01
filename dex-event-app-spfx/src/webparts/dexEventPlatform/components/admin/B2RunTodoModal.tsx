/**
 * v30.54: „Offen beim Veranstalter" — die Aufgabenliste zum B2Run.
 *
 * Warum als eigenes Fenster und nicht als Teil des Imports: Der Import ist ein
 * Ereignis, die Aufgaben sind ein Zustand. Nach dem Rücklauf gehen die
 * Abmeldungen weiter, und jede erzeugt eine Ummeldung beim Veranstalter — zu
 * einem Zeitpunkt, an dem niemand eine Import-Maske offen hat. Diese Liste
 * rechnet deshalb bei jedem Öffnen neu aus der Teilnehmerliste (s.
 * utils/b2runTodos) statt einen Stand von damals zu zeigen.
 *
 * **Gespeichert wird nur das Abhaken**, nicht die Liste selbst. Eine
 * gespeicherte Aufgabenliste wäre eine zweite Wahrheit neben den Anmeldungen:
 * Wird eine Abmeldung zurückgenommen oder eine Nummer korrigiert, stimmt sie
 * nicht mehr, und niemand merkt es. Der Haken liegt als Schlüssel-Liste im
 * Piggyback `_b2runTodoDone` der Event-Zeile; verschwindet die Aufgabe von
 * selbst (weil der Fall gelöst ist), ist der Haken bedeutungslos und stört
 * nicht.
 */
import * as React from 'react';
import Modal from '../Modal';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { DeloitteEvent } from '../../types';
import { EventService, SPRegistration } from '../../services/EventService';
import { buildB2RunTodos, B2RunTodo, B2RUN_TODO_LABELS } from '../../utils/b2runTodos';

const nameOf = (r?: SPRegistration): string =>
  r ? `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || (r.ParticipantEmail || '') : '—';

const KIND_COLOR: Record<B2RunTodo['kind'], { bg: string; fg: string }> = {
  transfer: { bg: 'rgba(237,139,0,0.12)', fg: 'var(--dex-orange-dark, #b35a00)' },
  assign: { bg: 'rgba(237,139,0,0.12)', fg: 'var(--dex-orange-dark, #b35a00)' },
  unregister: { bg: 'rgba(218,41,28,0.10)', fg: 'var(--dex-red, #da291c)' },
  register: { bg: 'rgba(21,101,192,0.10)', fg: '#1565c0' },
};

export default function B2RunTodoModal(props: {
  event: DeloitteEvent;
  service: EventService;
  onClose: () => void;
}): React.ReactElement {
  const { getAllRegistrations } = useEvents();
  const { showAlert } = useDialog();
  const [loading, setLoading] = React.useState(true);
  const [todos, setTodos] = React.useState<B2RunTodo[]>([]);
  const [done, setDone] = React.useState<string[]>([]);
  const [saving, setSaving] = React.useState(false);
  const [showDone, setShowDone] = React.useState(false);
  const [xlsxBusy, setXlsxBusy] = React.useState(false);

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const regs = await getAllRegistrations(props.event.id);
        if (cancelled) return;
        setTodos(buildB2RunTodos(regs));
        let stored: string[] = [];
        try {
          const o = JSON.parse(props.event.emailTemplateOverrides || '{}');
          if (Array.isArray(o?._b2runTodoDone)) stored = o._b2runTodoDone.filter((x: unknown) => typeof x === 'string');
        } catch { /* kein Piggyback — dann ist nichts abgehakt */ }
        setDone(stored);
      } catch {
        if (!cancelled) await showAlert('Die Teilnehmerliste konnte nicht gelesen werden.', { variant: 'error' });
      } finally { if (!cancelled) setLoading(false); }
    })().catch(() => { /* oben behandelt */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.event.id]);

  const toggle = async (key: string): Promise<void> => {
    const next = done.indexOf(key) >= 0 ? done.filter(k => k !== key) : done.concat(key);
    setDone(next);
    setSaving(true);
    // Nur den Haken schreiben — patchEventOverridesValue liest den aktuellen
    // Stand und merged, damit parallele Änderungen an anderen Piggybacks
    // (Abrechnung, Hotels) nicht verloren gehen.
    const ok = await props.service.patchEventOverridesValue(Number(props.event.id), '_b2runTodoDone', next)
      .catch(() => false);
    setSaving(false);
    if (!ok) {
      setDone(done); // zurückrollen — sonst sieht es erledigt aus und ist es nicht
      await showAlert('Der Haken konnte nicht gespeichert werden — bitte erneut versuchen.', { variant: 'error' });
    }
  };

  const open = todos.filter(t => done.indexOf(t.key) < 0);
  const closed = todos.filter(t => done.indexOf(t.key) >= 0);

  const downloadXlsx = async (): Promise<void> => {
    if (xlsxBusy) return;
    setXlsxBusy(true);
    try {
      const rows: string[][] = [['Status', 'Was ist zu tun', 'Startnummer', 'Von', 'E-Mail (von)', 'Auf', 'E-Mail (auf)', 'Sicherheit']];
      for (const t of todos) {
        rows.push([
          done.indexOf(t.key) >= 0 ? 'erledigt' : 'offen',
          `${B2RUN_TODO_LABELS[t.kind]} — ${t.action}`,
          t.bib || '',
          nameOf(t.from), t.from?.ParticipantEmail || '',
          nameOf(t.to), t.to?.ParticipantEmail || '',
          t.certain ? 'in DEX aufgezeichnet' : 'erschlossen — bitte prüfen',
        ]);
      }
      const XLSX = await import('xlsx');
      const ws = XLSX.utils.aoa_to_sheet(rows);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = [{ wch: 10 }, { wch: 76 }, { wch: 13 }, { wch: 24 }, { wch: 30 }, { wch: 24 }, { wch: 30 }, { wch: 26 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Offen beim Veranstalter');
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `B2Run_offene_Aufgaben_${new Date().toISOString().slice(0, 10)}.xlsx`;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
    } catch (err) {
      console.warn('[DEX] B2Run-Todo-Export fehlgeschlagen:', err);
      await showAlert('Die Excel-Datei konnte nicht erzeugt werden.', { variant: 'error' });
    } finally { setXlsxBusy(false); }
  };

  const renderTodo = (t: B2RunTodo, isDone: boolean): React.ReactElement => (
    <div
      key={t.key}
      style={{
        display: 'flex', alignItems: 'flex-start', gap: 10,
        padding: '10px 12px', border: '1px solid var(--dex-gray-200)',
        borderRadius: 10, marginBottom: 8,
        background: isDone ? 'var(--dex-gray-50, #fafafa)' : '#fff',
        opacity: isDone ? 0.65 : 1,
      }}
    >
      <input
        type="checkbox"
        checked={isDone}
        disabled={saving}
        onChange={() => { void toggle(t.key); }}
        style={{ marginTop: 3, flexShrink: 0, width: 18, height: 18, cursor: saving ? 'wait' : 'pointer' }}
        aria-label={isDone ? 'Als offen markieren' : 'Als erledigt markieren'}
      />
      <div style={{ minWidth: 0, flex: 1 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginBottom: 3 }}>
          <span style={{
            fontSize: '0.7rem', fontWeight: 700, padding: '2px 8px', borderRadius: 999,
            background: KIND_COLOR[t.kind].bg, color: KIND_COLOR[t.kind].fg,
          }}>{B2RUN_TODO_LABELS[t.kind]}</span>
          {t.bib && (
            <span style={{ fontFamily: "'Courier New',Courier,monospace", fontWeight: 700, fontSize: '0.95rem' }}>{t.bib}</span>
          )}
          {!t.certain && (
            <span style={{ fontSize: '0.7rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
              erschlossen, nicht aufgezeichnet
            </span>
          )}
        </div>
        <div style={{ fontSize: '0.83rem', lineHeight: 1.5, textDecoration: isDone ? 'line-through' : undefined }}>
          {t.action}
        </div>
        {(t.from || t.to) && (
          <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
            {t.from ? <>von {t.from.ParticipantEmail}</> : null}
            {t.from && t.to ? ' · ' : ''}
            {t.to ? <>auf {t.to.ParticipantEmail}</> : null}
          </div>
        )}
      </div>
    </div>
  );

  return (
    <Modal open onClose={props.onClose} maxWidth={780} ariaLabel="Offen beim Veranstalter">
      <div style={{ fontSize: '0.87rem', lineHeight: 1.55 }}>
        <h3 style={{ margin: '0 0 4px', fontSize: '1.05rem' }}>Offen beim Veranstalter (B2Run)</h3>
        <p style={{ margin: '0 0 14px', color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>
          Wird bei jedem Öffnen neu aus der Teilnehmerliste berechnet — spätere Abmeldungen tauchen
          also von selbst hier auf. Abgehakte Aufgaben bleiben gespeichert.
        </p>

        {loading ? (
          <p style={{ color: 'var(--dex-gray-500)' }}>Teilnehmerliste wird gelesen…</p>
        ) : todos.length === 0 ? (
          <p style={{
            padding: '12px 14px', borderRadius: 10,
            background: 'rgba(134,188,37,0.09)', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600,
          }}>
            Nichts offen — alle Startnummern sind zugeordnet.
          </p>
        ) : (
          <>
            {open.length === 0 ? (
              <p style={{
                padding: '12px 14px', borderRadius: 10, marginBottom: 12,
                background: 'rgba(134,188,37,0.09)', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600,
              }}>
                Alles abgehakt — {closed.length} Aufgabe{closed.length === 1 ? '' : 'n'} erledigt.
              </p>
            ) : (
              <>
                <div style={{ fontWeight: 700, marginBottom: 8 }}>Offen ({open.length})</div>
                {open.map(t => renderTodo(t, false))}
              </>
            )}

            {closed.length > 0 && (
              <div style={{ marginTop: 10 }}>
                <button
                  type="button"
                  onClick={() => setShowDone(v => !v)}
                  style={{
                    background: 'none', border: 'none', padding: 0, cursor: 'pointer',
                    color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600, fontSize: '0.8rem',
                    textDecoration: 'underline',
                  }}
                >
                  {showDone ? 'Erledigte ausblenden' : `Erledigte anzeigen (${closed.length})`}
                </button>
                {showDone && <div style={{ marginTop: 8 }}>{closed.map(t => renderTodo(t, true))}</div>}
              </div>
            )}
          </>
        )}

        <div style={{ display: 'flex', gap: 10, marginTop: 16, flexWrap: 'wrap' }}>
          <button
            type="button"
            className="btn btn-secondary"
            disabled={xlsxBusy || todos.length === 0}
            onClick={() => { void downloadXlsx(); }}
          >
            {xlsxBusy ? 'Wird erzeugt…' : 'Als Excel laden'}
          </button>
          <span style={{ flex: 1 }} />
          <button type="button" className="btn btn-primary" onClick={props.onClose}>Schließen</button>
        </div>
      </div>
    </Modal>
  );
}
