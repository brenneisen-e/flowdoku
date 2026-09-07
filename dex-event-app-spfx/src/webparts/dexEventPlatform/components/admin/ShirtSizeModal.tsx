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
 *
 * v30.88: Ist-Bestand je Größe + Gegenvorschläge. Nutzer-Ansage (07.09.2026,
 * B2Run Köln): „neben der Anzahl an Shirts auch ermöglichen, dass man angibt,
 * wie viele Shirts man wirklich hat, und dann wird geguckt, ob es überhaupt
 * passt — und bei jeder Person, wo es nicht mehr passt, ein Gegenvorschlag.
 * Die Funktion brauche ich dann für den Check-in bzw. die Abholung." Der
 * Bestand liegt als Piggyback `_shirtStock` am Hauptevent; die Verteilung
 * rechnet `utils/checkInExtras.shirtAllocate` — dieselbe Funktion, die die
 * Check-in-Seite je Person anzeigt, damit Tisch und Liste dasselbe sagen.
 */
import * as React from 'react';
import Modal from '../Modal';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { DeloitteEvent } from '../../types';
import { EventService, SPRegistration } from '../../services/EventService';
import { shirtTally, ShirtTallyResult, shirtAllocate, ShirtAllocationResult, ShirtStock, parseShirtStock, shirtSizeKey } from '../../utils/checkInExtras';

export default function ShirtSizeModal(props: {
  event: DeloitteEvent;
  onClose: () => void;
}): React.ReactElement {
  const { getAllRegistrations, events, refreshEvents } = useEvents();
  const { showAlert } = useDialog();
  const [loading, setLoading] = React.useState(true);
  const [result, setResult] = React.useState<ShirtTallyResult | null>(null);
  const [regs, setRegs] = React.useState<SPRegistration[]>([]);
  const [fields, setFields] = React.useState<Array<{ id: string; label: string }>>([]);
  const [skipped, setSkipped] = React.useState<string[]>([]);
  const [openSize, setOpenSize] = React.useState<string | null>(null);
  const [xlsxBusy, setXlsxBusy] = React.useState(false);
  // v30.88: Bestand — Eingabe als Text je Größe (leer = kein Bestand für die
  // Größe), gespeichert als Zahlen. `stockDirty` hält den Speichern-Knopf an.
  const [stockInput, setStockInput] = React.useState<Record<string, string>>(() => {
    const st = parseShirtStock(props.event.emailTemplateOverrides);
    const out: Record<string, string> = {};
    Object.keys(st).forEach(k => { out[k] = String(st[k]); });
    return out;
  });
  const [savedStock, setSavedStock] = React.useState<ShirtStock>(() => parseShirtStock(props.event.emailTemplateOverrides));
  const [stockSaving, setStockSaving] = React.useState(false);
  const [newSize, setNewSize] = React.useState('');

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
          const rs = await getAllRegistrations(ev.id, () => { ok = false; });
          if (!ok) { failed.push(ev.title); continue; }
          for (const r of rs) {
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
        const flds = targets
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .reduce<Array<{ id: string; label: string }>>((acc, ev) => acc.concat(((ev as any).eventSpecificFields || []) as Array<{ id: string; label: string }>), []);
        setFields(flds);
        setRegs(all);
        setResult(shirtTally(flds, all));
        setSkipped(failed);
      } catch (err) {
        console.warn('[DEX] Trikot-Auswertung fehlgeschlagen:', err);
        showAlert('Die Trikotgrößen konnten nicht gelesen werden.', { variant: 'error' });
      } finally { if (!cancelled) setLoading(false); }
    })().catch(() => { /* im finally behandelt */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.event.id]);

  // Bestand aus den Eingaben (nur gültige Zahlen zählen).
  const stockFromInput = React.useMemo((): ShirtStock => {
    const out: ShirtStock = {};
    Object.keys(stockInput).forEach(k => {
      const v = stockInput[k].trim();
      if (v === '') return;
      const n = parseInt(v, 10);
      if (isFinite(n) && n >= 0) out[k] = n;
    });
    return out;
  }, [stockInput]);
  const stockDirty = React.useMemo(() => {
    const a = Object.keys(stockFromInput).sort(); const b = Object.keys(savedStock).sort();
    if (a.join('|') !== b.join('|')) return true;
    return a.some(k => stockFromInput[k] !== savedStock[k]);
  }, [stockFromInput, savedStock]);
  // Die Verteilung rechnet LIVE mit den Eingaben — der Organizer sieht sofort,
  // was eine Zahl mehr oder weniger bedeutet; gespeichert wird bewusst extra.
  const alloc: ShirtAllocationResult | null = React.useMemo(
    () => (result && result.fieldLabel) ? shirtAllocate(fields, regs, stockFromInput) : null,
    [result, fields, regs, stockFromInput],
  );

  const saveStock = async (): Promise<void> => {
    if (stockSaving) return;
    setStockSaving(true);
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (!ctx) throw new Error('Kein SPFx-Kontext');
      const svc = new EventService(ctx);
      const ok = await svc.patchEventOverridesValue(Number(props.event.id), '_shirtStock', Object.keys(stockFromInput).length ? stockFromInput : null);
      if (!ok) throw new Error('patch failed');
      setSavedStock({ ...stockFromInput });
      await refreshEvents();
      showAlert('Bestand gespeichert — die Check-in-Seite zeigt Gegenvorschläge jetzt je Person an.', { variant: 'success' });
    } catch (err) {
      console.warn('[DEX] Trikot-Bestand speichern fehlgeschlagen:', err);
      showAlert('Der Bestand konnte nicht gespeichert werden — bitte erneut versuchen.', { variant: 'error' });
    } finally { setStockSaving(false); }
  };

  const downloadXlsx = async (): Promise<void> => {
    if (xlsxBusy || !result) return;
    setXlsxBusy(true);
    try {
      const hasStock = !!(alloc && alloc.hasStock);
      const rows: string[][] = [hasStock ? ['Größe', 'Benötigt', 'Bestand', 'Fehlt', 'Reserve', 'Personen'] : ['Größe', 'Anzahl', 'Personen']];
      for (const r of result.rows) {
        const a = alloc ? alloc.rows.find(x => x.key === shirtSizeKey(r.size)) : undefined;
        rows.push(hasStock
          ? [r.size || 'ohne Angabe', String(r.count), String(a ? a.stock : 0), String(a ? a.missing : 0), String(a ? a.spare : 0), r.names.join(', ')]
          : [r.size || 'ohne Angabe', String(r.count), r.names.join(', ')]);
      }
      rows.push([]);
      rows.push(['Summe', String(result.total), '']);
      const XLSX = await import('xlsx');
      const ws = XLSX.utils.aoa_to_sheet(rows);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = hasStock ? [{ wch: 16 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 8 }, { wch: 90 }] : [{ wch: 16 }, { wch: 10 }, { wch: 90 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Trikots');
      if (hasStock && alloc) {
        const prop: string[][] = [['Person', 'E-Mail', 'Wunschgröße', 'Vorschlag']];
        for (const r of regs) {
          const em = (r.ParticipantEmail || '').toLowerCase().trim();
          const a = alloc.byEmail[em];
          if (!a || !a.short) continue;
          prop.push([(r.ParticipantName || em).trim(), r.ParticipantEmail || '', a.wish, a.proposal || 'keine Größe mehr vorrätig']);
        }
        const ws2 = XLSX.utils.aoa_to_sheet(prop);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (ws2 as any)['!cols'] = [{ wch: 30 }, { wch: 34 }, { wch: 14 }, { wch: 26 }];
        XLSX.utils.book_append_sheet(wb, ws2, 'Gegenvorschläge');
      }
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
  // Zeilen für den Bestand: alle gewünschten Größen plus alle, für die schon
  // ein Bestand eingetragen ist (auch wenn niemand sie wollte — die Reserve).
  const stockKeys = React.useMemo((): Array<{ key: string; label: string }> => {
    const seen: Record<string, string> = {};
    (result ? result.rows : []).forEach(r => { if (r.size) seen[shirtSizeKey(r.size)] = r.size; });
    Object.keys(stockInput).forEach(k => { if (!seen[k]) seen[k] = k.toUpperCase(); });
    return (alloc ? alloc.rows.map(r => r.key) : Object.keys(seen))
      .filter((k, i, arr) => arr.indexOf(k) === i && !!seen[k])
      .map(k => ({ key: k, label: seen[k] }));
  }, [result, stockInput, alloc]);
  const proposals = React.useMemo(() => {
    if (!alloc || !alloc.hasStock) return [];
    return regs
      .map(r => {
        const em = (r.ParticipantEmail || '').toLowerCase().trim();
        const a = alloc.byEmail[em];
        return a && a.short ? { name: (r.ParticipantName || em).trim(), tid: r.TeilnehmerID, wish: a.wish, proposal: a.proposal } : null;
      })
      .filter((x): x is { name: string; tid: number | undefined; wish: string; proposal: string | null } => !!x);
  }, [alloc, regs]);

  return (
    <Modal open onClose={props.onClose} maxWidth={760} ariaLabel="Benötigte T-Shirts">
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
              Gezählt über das Feld <strong>{result.fieldLabel}</strong> · {result.total} angemeldete {result.total === 1 ? 'Person' : 'Personen'}
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
              const a = (alloc && alloc.hasStock && r.size) ? alloc.rows.find(x => x.key === shirtSizeKey(r.size)) : undefined;
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
                    {/* v30.88: Soll/Ist-Abgleich je Zeile, sobald ein Bestand eingetragen ist. */}
                    {a && (
                      a.missing > 0
                        ? <span style={{ fontSize: '0.72rem', fontWeight: 700, padding: '2px 8px', borderRadius: 999, background: 'rgba(218,41,28,0.10)', color: 'var(--dex-red, #da291c)', whiteSpace: 'nowrap' }}>fehlt {a.missing}</span>
                        : <span style={{ fontSize: '0.72rem', fontWeight: 700, padding: '2px 8px', borderRadius: 999, background: 'rgba(134,188,37,0.15)', color: 'var(--dex-green-dark, #4a7c1f)', whiteSpace: 'nowrap' }}>reicht{a.spare > 0 ? ` (+${a.spare})` : ''}</span>
                    )}
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

          {/* v30.88: Ist-Bestand je Größe + Gegenvorschläge. */}
          <div style={{ marginTop: 20, paddingTop: 16, borderTop: '2px solid var(--dex-gray-200)' }}>
            <div style={{ display: 'flex', alignItems: 'baseline', gap: 12, flexWrap: 'wrap', marginBottom: 8 }}>
              <h4 style={{ margin: 0, fontSize: '0.98rem' }}>Vorhandene T-Shirts (Bestand)</h4>
              <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', flex: '1 1 240px' }}>
                Trag ein, wie viele Shirts du je Größe wirklich hast. Die App prüft, ob es reicht, und macht je Person einen
                Gegenvorschlag — den sieht das Check-in-Team bei der Abholung direkt an der Person.
              </span>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(120px, 1fr))', gap: 8 }}>
              {stockKeys.map(s => {
                const a = alloc ? alloc.rows.find(x => x.key === s.key) : undefined;
                return (
                  <label key={s.key} style={{ display: 'flex', flexDirection: 'column', gap: 3, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                    <span><strong style={{ color: 'var(--dex-gray-800)' }}>{s.label}</strong>{a && a.need > 0 ? ` · benötigt ${a.need}` : ' · niemand gewünscht'}</span>
                    <input
                      type="number"
                      min={0}
                      inputMode="numeric"
                      className="form-input"
                      value={stockInput[s.key] ?? ''}
                      placeholder="0"
                      onChange={e => setStockInput(prev => ({ ...prev, [s.key]: e.target.value.replace(/[^\d]/g, '') }))}
                      style={{ padding: '6px 10px', fontSize: '0.95rem', fontWeight: 700, textAlign: 'center' }}
                    />
                  </label>
                );
              })}
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginTop: 10 }}>
              <input
                type="text"
                className="form-input"
                value={newSize}
                onChange={e => setNewSize(e.target.value)}
                placeholder="weitere Größe, z.B. XXL"
                style={{ padding: '6px 10px', fontSize: '0.85rem', maxWidth: 200 }}
              />
              <button
                type="button"
                className="btn btn-secondary"
                style={{ fontSize: '0.78rem', padding: '6px 12px' }}
                disabled={!shirtSizeKey(newSize)}
                onClick={() => {
                  const k = shirtSizeKey(newSize);
                  if (!k) return;
                  setStockInput(prev => (prev[k] !== undefined ? prev : { ...prev, [k]: '' }));
                  setNewSize('');
                }}
              >
                Größe ergänzen
              </button>
              <span style={{ flex: 1 }} />
              <button
                type="button"
                className="btn btn-primary"
                style={{ fontSize: '0.8rem', padding: '6px 16px' }}
                disabled={!stockDirty || stockSaving}
                onClick={() => { void saveStock(); }}
              >
                {stockSaving ? 'Speichert…' : (stockDirty ? 'Bestand speichern' : 'Bestand gespeichert')}
              </button>
            </div>

            {alloc && alloc.hasStock && (
              <div style={{ marginTop: 14 }}>
                {proposals.length === 0 && alloc.noneLeft.length === 0 ? (
                  <div style={{ padding: '10px 12px', borderRadius: 8, background: 'rgba(134,188,37,0.10)', fontSize: '0.82rem', lineHeight: 1.5 }}>
                    <strong>Der Bestand reicht für alle Wünsche.</strong> Jede Person bekommt ihre Größe.
                  </div>
                ) : (
                  <>
                    <div style={{ fontSize: '0.85rem', fontWeight: 700, marginBottom: 6 }}>
                      Gegenvorschläge ({proposals.length})
                      <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)', marginLeft: 8, fontSize: '0.76rem' }}>
                        Reihenfolge nach Teilnehmer-ID: Wer zuerst angemeldet war, bekommt seine Wunschgröße. Ausweichgröße = nächste mit Rest, zuerst eine Nummer größer.
                      </span>
                    </div>
                    <div style={{ maxHeight: 260, overflow: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 8 }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem' }}>
                        <thead>
                          <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                            <th style={{ textAlign: 'left', padding: 8, position: 'sticky', top: 0, background: '#fff' }}>ID</th>
                            <th style={{ textAlign: 'left', padding: 8, position: 'sticky', top: 0, background: '#fff' }}>Person</th>
                            <th style={{ textAlign: 'left', padding: 8, position: 'sticky', top: 0, background: '#fff' }}>Wunsch</th>
                            <th style={{ textAlign: 'left', padding: 8, position: 'sticky', top: 0, background: '#fff' }}>Vorschlag</th>
                          </tr>
                        </thead>
                        <tbody>
                          {proposals.map((p, i) => (
                            <tr key={i} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                              <td style={{ padding: 8, color: 'var(--dex-gray-400)' }}>{p.tid ?? '—'}</td>
                              <td style={{ padding: 8, fontWeight: 600 }}>{p.name}</td>
                              <td style={{ padding: 8 }}><span style={{ textDecoration: 'line-through', color: 'var(--dex-gray-500)' }}>{p.wish}</span></td>
                              <td style={{ padding: 8 }}>
                                {p.proposal
                                  ? <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>{p.proposal}</strong>
                                  : <strong style={{ color: 'var(--dex-red, #da291c)' }}>keine Größe mehr vorrätig</strong>}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </>
                )}
              </div>
            )}
          </div>
        </>
      )}

      <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 20 }}>
        <button type="button" className="btn btn-secondary" onClick={props.onClose}>Schließen</button>
      </div>
    </Modal>
  );
}
