/**
 * v30.54: „Offen beim Veranstalter" — die Aufgabenliste zum B2Run.
 *
 * **v30.55 korrigiert den Ansatz aus v30.54.** Dort wurde die Liste
 * ausschließlich aus den Teilnehmerzeilen abgeleitet — und war direkt nach
 * einem Import leer, obwohl beim Veranstalter neun Ummeldungen warteten. Der
 * Grund steht in `utils/b2runTodos`: Die Startnummer einer abgemeldeten Person
 * existiert nur in der Datei des Veranstalters, nie in DEX. Was der Import
 * erzeugt hat, wird deshalb FESTGEHALTEN (`_b2runTodo`); was seither passiert
 * ist, kommt weiterhin aus der Ableitung. Beides wird hier zusammengeführt.
 *
 * Der Haken liegt getrennt davon in `_b2runTodoDone`.
 */
import * as React from 'react';
import Modal from '../Modal';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { DeloitteEvent } from '../../types';
import { EventService, SPRegistration } from '../../services/EventService';
import { mergeB2RunTodos, B2RunTodo, StoredB2RunTodo, B2RUN_TODO_LABELS } from '../../utils/b2runTodos';

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
  const [regs, setRegs] = React.useState<SPRegistration[]>([]);
  const [bibBusy, setBibBusy] = React.useState<string>('');
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
        setRegs(regs);
        let storedTodos: StoredB2RunTodo[] = [];
        let storedDone: string[] = [];
        try {
          const o = JSON.parse(props.event.emailTemplateOverrides || '{}');
          if (Array.isArray(o?._b2runTodo)) storedTodos = o._b2runTodo.filter((x: unknown) => !!x && typeof x === 'object');
          if (Array.isArray(o?._b2runTodoDone)) storedDone = o._b2runTodoDone.filter((x: unknown) => typeof x === 'string');
        } catch { /* kein Piggyback — dann ist nichts festgehalten */ }
        setTodos(mergeB2RunTodos(storedTodos, regs));
        setDone(storedDone);
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

  /**
   * v30.55: Die Startnummer in DEX auf die Person umtragen, die sie beim
   * Veranstalter bekommt.
   *
   * Die Ummeldung beim Veranstalter ist das eine — steht die Nummer in DEX
   * weiter bei der abgemeldeten Person (oder bei niemandem), zeigt der
   * Check-in für die Person, die wirklich läuft, keine Nummer an. Der Knopf
   * setzt sie deshalb um und räumt sie bei der Vorgängerin weg; danach
   * verschwindet die Aufgabe von selbst aus der abgeleiteten Liste.
   */
  const moveBibInDex = async (t: B2RunTodo): Promise<void> => {
    if (bibBusy || !t.toReg || !t.bib || !props.event.subsiteUrl) return;
    setBibBusy(t.key);
    try {
      const actor = { name: 'DEX', email: '' };
      await props.service.adminUpdateRegistration(props.event.subsiteUrl, t.toReg.Id, { Startnummer: t.bib }, actor);
      // Bei der Vorgängerin leeren, damit die Nummer nicht zweimal in der
      // Liste steht — sonst hält die Ableitung sie für weiterhin belegt.
      const from = regs.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === (t.fromEmail || '').toLowerCase().trim()
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        && String((r as any).Startnummer || '').trim() === t.bib);
      if (from) {
        await props.service.adminUpdateRegistration(props.event.subsiteUrl, from.Id, { Startnummer: '' }, actor);
      }
      const fresh = await getAllRegistrations(props.event.id);
      setRegs(fresh);
      setTodos(prev => prev.map(x => x.key === t.key
        ? { ...x, bibInDex: true, toReg: fresh.find(r => r.Id === t.toReg!.Id) || x.toReg }
        : x));
    } catch {
      await showAlert('Die Startnummer konnte nicht übertragen werden — bitte erneut versuchen.', { variant: 'error' });
    } finally { setBibBusy(''); }
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
          t.fromName || '', t.fromEmail || '',
          t.toName || '', t.toEmail || '',
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
        {(t.fromEmail || t.toEmail) && (
          <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
            {t.fromEmail ? <>von {t.fromEmail}</> : null}
            {t.fromEmail && t.toEmail ? ' · ' : ''}
            {t.toEmail ? <>auf {t.toEmail}</> : null}
          </div>
        )}
        {/* v30.55: Die Ummeldung beim Veranstalter ist das eine — die Nummer
            muss aber auch in DEX bei der richtigen Person stehen, sonst zeigt
            der Check-in für die Person, die wirklich läuft, gar keine Nummer. */}
        {!isDone && t.bib && t.toReg && !t.bibInDex && (
          <button
            type="button"
            className="btn btn-secondary"
            disabled={!!bibBusy}
            onClick={() => { void moveBibInDex(t); }}
            style={{ fontSize: '0.75rem', padding: '4px 12px', marginTop: 6 }}
          >
            {bibBusy === t.key ? 'Wird übertragen…' : `Startnummer ${t.bib} in DEX auf ${t.toName} übertragen`}
          </button>
        )}
        {t.bib && t.bibInDex && (
          <div style={{ fontSize: '0.72rem', color: 'var(--dex-green-dark, #4a7c1f)', marginTop: 4 }}>
            In DEX steht die Nummer bereits bei {t.toName}.
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
