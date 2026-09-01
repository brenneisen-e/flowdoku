/**
 * v30.60: „Benötigte T-Shirts" — die Bestellliste zum Event.
 *
 * Nutzer-Ansage: „Und ich brauche eine Aktion mit benötigte T-Shirts." Gemeint
 * ist die Zahl, die beim Ausstatter bestellt wird — pro Größe, nicht pro
 * Person. Bisher stand die Angabe nur verstreut in den Abfrage-Antworten der
 * Teilnehmerliste; wer bestellen wollte, exportierte nach Excel und zählte von
 * Hand.
 *
 * Zwei Entscheidungen, die die Zahl belastbar machen (Begründung in
 * `utils/checkInExtras.shirtTally`):
 *  - Gezählt wird über DIESELBE Feld-Erkennung wie am Check-in-Tisch.
 *  - Wer keine Größe angegeben hat, taucht namentlich auf, statt aus der
 *    Summe zu verschwinden.
 *
 * Bei einem Klammer-Event mit Terminen wird über ALLE Termine gezählt: Das
 * Trikot bekommt die Person, nicht die Anmeldung. Doppelt gezählt wird
 * niemand — die Zusammenführung läuft über die E-Mail-Adresse (CLAUDE.md:
 * „Die E-Mail-Adresse ist der einzige Schlüssel").
 */
import * as React from 'react';
import Modal from '../Modal';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { DeloitteEvent } from '../../types';
import { SPRegistration } from '../../services/EventService';
import { shirtTally, ShirtTallyResult } from '../../utils/checkInExtras';

export default function ShirtSizeModal(props: {
  event: DeloitteEvent;
  onClose: () => void;
}): React.ReactElement {
  const { getAllRegistrations, events } = useEvents();
  const { showAlert } = useDialog();
  const [loading, setLoading] = React.useState(true);
  const [result, setResult] = React.useState<ShirtTallyResult | null>(null);
  const [skipped, setSkipped] = React.useState<string[]>([]);
  const [openSize, setOpenSize] = React.useState<string | null>(null);
  const [xlsxBusy, setXlsxBusy] = React.useState(false);

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const children = events.filter(e => e.parentEventId === props.event.id);
        const targets = [props.event, ...children];
        const all: SPRegistration[] = [];
        const failed: string[] = [];
        const seen: Record<string, true> = {};
        for (const ev of targets) {
          if (!ev.subsiteUrl) continue;
          let ok = true;
          // v30.37-Lehre: Ein leeres Ergebnis ohne geprüften Status ist keine
          // Aussage über die Daten. Ein gesperrter Termin wird deshalb NAMENTLICH
          // gemeldet, statt als „0 Trikots" durchzugehen — sonst bestellt man zu
          // wenig und erfährt den Grund nie.
          const regs = await getAllRegistrations(ev.id, () => { ok = false; });
          if (!ok) { failed.push(ev.title); continue; }
          for (const r of regs) {
            const key = (r.ParticipantEmail || '').toLowerCase().trim();
            // Ohne Adresse lässt sich nichts zusammenführen — die Zeile zählt
            // dann einzeln, das ist ehrlicher als sie wegzulassen.
            if (key && seen[key]) continue;
            if (key) seen[key] = true;
            all.push(r);
          }
        }
        if (cancelled) return;
        // Die Feld-Definitionen des Hauptevents plus die der Termine: Das
        // Trikot-Feld kann auf beiden Ebenen stehen (CLAUDE.md: Antworten
        // stehen dort, wo angemeldet wurde).
        const fields = targets
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .reduce<Array<{ id: string; label: string }>>((acc, ev) => acc.concat(((ev as any).eventSpecificFields || []) as Array<{ id: string; label: string }>), []);
        setResult(shirtTally(fields, all));
        setSkipped(failed);
      } catch (err) {
        console.warn('[DEX] Trikot-Auswertung fehlgeschlagen:', err);
        showAlert('Die Trikotgrößen konnten nicht gelesen werden.', { variant: 'error' });
      } finally { if (!cancelled) setLoading(false); }
    })().catch(() => { /* im finally behandelt */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.event.id]);

  const downloadXlsx = async (): Promise<void> => {
    if (xlsxBusy || !result) return;
    setXlsxBusy(true);
    try {
      const rows: string[][] = [['Größe', 'Anzahl', 'Personen']];
      for (const r of result.rows) {
        rows.push([r.size || 'ohne Angabe', String(r.count), r.names.join(', ')]);
      }
      rows.push([]);
      rows.push(['Summe', String(result.total), '']);
      const XLSX = await import('xlsx');
      const ws = XLSX.utils.aoa_to_sheet(rows);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = [{ wch: 16 }, { wch: 10 }, { wch: 90 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Trikots');
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Trikots_${new Date().toISOString().slice(0, 10)}.xlsx`;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
    } catch (err) {
      console.warn('[DEX] Trikot-Export fehlgeschlagen:', err);
      showAlert('Die Excel-Datei konnte nicht erzeugt werden.', { variant: 'error' });
    } finally { setXlsxBusy(false); }
  };

  const maxCount = result ? result.rows.reduce((m, r) => Math.max(m, r.count), 0) : 0;

  return (
    <Modal open onClose={props.onClose} maxWidth={720} ariaLabel="Benötigte T-Shirts">
      <h3 style={{ margin: '0 0 4px' }}>Benötigte T-Shirts</h3>
      <p style={{ margin: '0 0 16px', color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
        {props.event.title}
      </p>

      {loading && <p style={{ color: 'var(--dex-gray-500)' }}>Trikotgrößen werden gelesen…</p>}

      {!loading && result && !result.fieldLabel && (
        <div style={{ padding: '14px 16px', borderRadius: 8, background: 'rgba(237,139,0,0.09)', fontSize: '0.85rem', lineHeight: 1.6 }}>
          Dieses Event hat kein Abfragefeld, das nach einer Trikot- oder Konfektionsgröße aussieht.
          Lege im Assistenten unter <strong>Felder</strong> ein Feld an, dessen Bezeichnung die Größe
          benennt (z.B. &bdquo;T-Shirt Größe&ldquo; oder &bdquo;Trikotgröße&ldquo;) — danach zählt diese Ansicht
          automatisch mit, und die Größe steht auch am Check-in-Tisch.
        </div>
      )}

      {!loading && result && result.fieldLabel && (
        <>
          <div style={{ display: 'flex', alignItems: 'baseline', gap: 12, flexWrap: 'wrap', marginBottom: 12 }}>
            <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', flex: '1 1 240px' }}>
              Gezählt über das Feld <strong>{result.fieldLabel}</strong> · {result.total} angemeldete{result.total === 1 ? '' : ''} {result.total === 1 ? 'Person' : 'Personen'}
              {result.missing > 0 && <> · <strong>{result.missing} ohne Angabe</strong></>}
            </div>
            <button
              type="button"
              className="btn btn-primary"
              style={{ fontSize: '0.8rem', padding: '6px 16px' }}
              disabled={xlsxBusy}
              onClick={() => { void downloadXlsx(); }}
            >
              {xlsxBusy ? 'Wird erzeugt…' : 'Als Excel laden'}
            </button>
          </div>

          {skipped.length > 0 && (
            <div style={{ marginBottom: 12, padding: '10px 12px', borderRadius: 8, background: 'rgba(218,41,28,0.08)', fontSize: '0.8rem', lineHeight: 1.5 }}>
              Für {skipped.length === 1 ? 'diesen Termin' : 'diese Termine'} konnte die Teilnehmerliste nicht gelesen werden —
              die Zahlen unten sind deshalb unvollständig: <strong>{skipped.join(', ')}</strong>.
            </div>
          )}

          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            {result.rows.map((r, i) => {
              const key = r.size || '__none__';
              const open = openSize === key;
              const pct = maxCount > 0 ? Math.round((r.count / maxCount) * 100) : 0;
              const none = !r.size;
              return (
                <div key={i} style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, overflow: 'hidden' }}>
                  <button
                    type="button"
                    onClick={() => setOpenSize(open ? null : key)}
                    style={{
                      width: '100%', display: 'flex', alignItems: 'center', gap: 12,
                      padding: '10px 14px', border: 'none', cursor: 'pointer', textAlign: 'left',
                      background: open ? 'var(--dex-gray-100)' : '#fff',
                    }}
                  >
                    <span style={{
                      minWidth: 76, fontWeight: 700, fontSize: '0.95rem',
                      color: none ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-800)',
                    }}>{r.size || 'ohne Angabe'}</span>
                    {/* Balken statt nur Zahl: Beim Bestellen zählt vor allem,
                        welche Größen die Masse ausmachen. */}
                    <span style={{ flex: 1, minWidth: 60, height: 8, borderRadius: 999, background: 'var(--dex-gray-100)' }}>
                      <span style={{
                        display: 'block', height: 8, width: `${pct}%`, borderRadius: 999,
                        background: none ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)',
                      }} />
                    </span>
                    <strong style={{ minWidth: 34, textAlign: 'right', fontSize: '1rem' }}>{r.count}</strong>
                    <span style={{ color: 'var(--dex-gray-400)', fontSize: '0.8rem' }}>{open ? '▾' : '▸'}</span>
                  </button>
                  {open && (
                    <div style={{ padding: '8px 14px 12px', fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.6, borderTop: '1px solid var(--dex-gray-100)' }}>
                      {r.names.join(' · ')}
                    </div>
                  )}
                </div>
              );
            })}
          </div>

          {result.missing > 0 && (
            <div style={{ marginTop: 12, padding: '10px 12px', borderRadius: 8, background: 'rgba(237,139,0,0.09)', fontSize: '0.8rem', lineHeight: 1.55 }}>
              <strong>{result.missing} {result.missing === 1 ? 'Person hat' : 'Personen haben'} keine Größe angegeben.</strong>{' '}
              Die Namen stehen in der Zeile &bdquo;ohne Angabe&ldquo; — frag dort nach, bevor du bestellst.
              Sonst fehlt am Lauftag genau diese Anzahl Trikots.
            </div>
          )}
        </>
      )}

      <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 20 }}>
        <button type="button" className="btn btn-secondary" onClick={props.onClose}>Schließen</button>
      </div>
    </Modal>
  );
}
