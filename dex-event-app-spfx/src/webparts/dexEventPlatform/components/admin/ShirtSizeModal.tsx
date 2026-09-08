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
 * v31.3: Diese Zusammenführung behielt je Person die ZUERST gelesene Zeile.
 * Damit gewann bei einem Klammer-Event die Schattenzeile (oft ohne
 * CustomData) und bei mehreren Terminen die zufällig vorne stehende Ebene —
 * eine im Organizer Center korrigierte Größe kam in dieser Liste nie an
 * (Befund 08.09.2026: „L"/„XL" blieben stehen, obwohl längst auf
 * „Herrengröße L"/„XL" geändert). Welche Zeile die Frage beantwortet,
 * entscheidet jetzt `pickShirtAnswerRows` — mit Antwort vor ohne, aktiv vor
 * abgemeldet, sonst die zuletzt geänderte.
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
import { useLocaleSafe } from '../../context/LanguageContext';
import { DeloitteEvent } from '../../types';
import { EventService, SPRegistration } from '../../services/EventService';
import {
  shirtTally, ShirtTallyResult, shirtAllocate, ShirtAllocationResult, ShirtStock, parseShirtStock, shirtSizeKey,
  pickShirtAnswerRows, ShirtAnswerConflict, shirtSizeLabel, splitShirtSize, isShirtSizeKey,
} from '../../utils/checkInExtras';
import { cx } from '../dexUi';
import { Shirt, Download, Plus, Check, ChevronDown, AlertCircle } from '../Icons';

export default function ShirtSizeModal(props: {
  event: DeloitteEvent;
  onClose: () => void;
}): React.ReactElement {
  // v31.2: Zweisprachig wie jeder andere Dialog — bisher nur Deutsch. Die
  // Props bleiben unverändert (kein `isDe`-Prop), die Sprache kommt aus dem
  // Context; ohne Provider fällt sie auf Deutsch zurück.
  const isDe = useLocaleSafe() === 'de';
  const { getAllRegistrations, events, refreshEvents } = useEvents();
  const { showAlert } = useDialog();
  const [loading, setLoading] = React.useState(true);
  const [result, setResult] = React.useState<ShirtTallyResult | null>(null);
  const [regs, setRegs] = React.useState<SPRegistration[]>([]);
  const [fields, setFields] = React.useState<Array<{ id: string; label: string }>>([]);
  const [skipped, setSkipped] = React.useState<string[]>([]);
  // v31.3: Dieselbe Person mit zwei verschiedenen Größen auf zwei Zeilen —
  // gerechnet wird mit der maßgeblichen, gesagt wird es trotzdem.
  const [conflicts, setConflicts] = React.useState<ShirtAnswerConflict[]>([]);
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
        for (const ev of targets) {
          if (!ev.subsiteUrl) continue;
          let ok = true;
          // v30.37-Lehre: Ein leeres Ergebnis ohne geprüften Status ist keine
          // Aussage über die Daten. Ein gesperrter Termin wird deshalb NAMENTLICH
          // gemeldet, statt als „0 Trikots" durchzugehen — sonst bestellt man zu
          // wenig und erfährt den Grund nie.
          const rs = await getAllRegistrations(ev.id, () => { ok = false; });
          if (!ok) { failed.push(ev.title); continue; }
          // v31.3: Hier wird NICHT mehr zusammengeführt — welche Zeile die
          // Größenfrage beantwortet, lässt sich erst sagen, wenn das Feld
          // bekannt ist und alle Ebenen gelesen sind (s. unten).
          for (const r of rs) all.push(r);
        }
        if (cancelled) return;
        // Die Feld-Definitionen des Hauptevents plus die der Termine: Das
        // Trikot-Feld kann auf beiden Ebenen stehen (CLAUDE.md: Antworten
        // stehen dort, wo angemeldet wurde).
        const flds = targets
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .reduce<Array<{ id: string; label: string }>>((acc, ev) => acc.concat(((ev as any).eventSpecificFields || []) as Array<{ id: string; label: string }>), []);
        // v31.3: Eine Zeile je Person — die, die die Größenfrage wirklich
        // beantwortet (Antwort vor keiner Antwort, aktiv vor abgemeldet, sonst
        // die zuletzt geänderte). Vorher gewann die zuerst gelesene Ebene.
        const picked = pickShirtAnswerRows(flds, all);
        setFields(flds);
        setRegs(picked.rows);
        setConflicts(picked.conflicts);
        setResult(shirtTally(flds, picked.rows));
        setSkipped(failed);
      } catch (err) {
        console.warn('[DEX] Trikot-Auswertung fehlgeschlagen:', err);
        showAlert(isDe ? 'Die Trikotgrößen konnten nicht gelesen werden.' : 'The shirt sizes could not be read.', { variant: 'error' });
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
      showAlert(isDe ? 'Bestand gespeichert — die Check-in-Seite zeigt Gegenvorschläge jetzt je Person an.' : 'Stock saved — the check-in page now shows the alternative size per person.', { variant: 'success' });
    } catch (err) {
      console.warn('[DEX] Trikot-Bestand speichern fehlgeschlagen:', err);
      showAlert(isDe ? 'Der Bestand konnte nicht gespeichert werden — bitte erneut versuchen.' : 'The stock could not be saved — please try again.', { variant: 'error' });
    } finally { setStockSaving(false); }
  };

  const downloadXlsx = async (): Promise<void> => {
    if (xlsxBusy || !result) return;
    setXlsxBusy(true);
    try {
      const hasStock = !!(alloc && alloc.hasStock);
      const rows: string[][] = [hasStock ? ['Größe', 'Benötigt', 'Bestand', 'Fehlt', 'Reserve', 'Personen'] : ['Größe', 'Anzahl', 'Personen']];
      for (const r of result.rows) {
        // v31.3: Eine Abwahl-Antwort ist keine Größe — sie steht mit ihrem
        // Wortlaut in der Datei, bekommt aber keine Bestands-Spalten, damit
        // niemand sie beim Ausstatter bestellt.
        const a = (alloc && r.size && !r.optOut) ? alloc.rows.find(x => x.key === shirtSizeKey(r.size)) : undefined;
        const label = r.size ? (r.optOut ? `${r.size} (kein Shirt nötig)` : r.size) : 'ohne Angabe';
        rows.push(hasStock
          ? [label, String(r.count), a ? String(a.stock) : '—', a ? String(a.missing) : '—', a ? String(a.spare) : '—', r.names.join(', ')]
          : [label, String(r.count), r.names.join(', ')]);
      }
      rows.push([]);
      rows.push(['Summe', String(result.total), '']);
      // v31.3: Was davon wirklich bestellt wird — Abwahl und fehlende Angaben
      // sind keine Shirts, standen aber bisher stillschweigend in derselben Summe.
      rows.push(['davon kein Shirt nötig', String(result.optOut), '']);
      rows.push(['davon ohne Größenangabe', String(result.missing), '']);
      rows.push(['Shirts zu bestellen', String(result.sizeTotal), '']);
      if (skipped.length > 0) {
        // v31.3: Der Hinweis gehört IN die Datei — eine Bestellliste wird
        // weitergereicht, der rote Kasten im Dialog bleibt zurück.
        rows.push([]);
        rows.push(['ACHTUNG: Die Zahlen oben sind UNVOLLSTÄNDIG — für diese Termine konnte die Teilnehmerliste nicht gelesen werden:']);
        skipped.forEach(t => rows.push([`· ${t}`]));
        rows.push(['Die Werte sind damit Untergrenzen. Vor dem Bestellen die Leserechte prüfen und neu laden.']);
      }
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
        if (skipped.length > 0) {
          prop.push([]);
          prop.push([`ACHTUNG: unvollständig — nicht lesbare Termine: ${skipped.join(', ')}`]);
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
      showAlert(isDe ? 'Die Excel-Datei konnte nicht erzeugt werden.' : 'The Excel file could not be created.', { variant: 'error' });
    } finally { setXlsxBusy(false); }
  };

  const maxCount = result ? result.rows.reduce((m, r) => Math.max(m, r.count), 0) : 0;
  // Zeilen für den Bestand: alle gewünschten Größen plus alle, für die schon
  // ein Bestand eingetragen ist (auch wenn niemand sie wollte — die Reserve).
  const stockKeys = React.useMemo((): Array<{ key: string; label: string }> => {
    const seen: Record<string, string> = {};
    // v31.3: Nur echte Größen bekommen eine Bestandszeile. Eine Abwahl
    // („T-Shirt bereits aus Vorjahren vorhanden") hatte bis dahin ein
    // Bestandsfeld, als wäre sie eine Bestellposition — 30 von 90 Personen im
    // B2Run-Event standen so in der Bestellung.
    (result ? result.rows : []).forEach(r => { if (r.size && !r.optOut) seen[shirtSizeKey(r.size)] = r.size; });
    Object.keys(stockInput).forEach(k => { if (!seen[k] && isShirtSizeKey(k)) seen[k] = shirtSizeLabel(k); });
    // v31.2: Die Schlüssel aus `seen` hinten anhängen — eine über „Größe
    // ergänzen" neu angelegte Größe steht mit leerem Wert noch nicht in
    // `stockFromInput`, also auch nicht in `alloc.rows`, und fiel bis dahin
    // aus der Liste, bevor man ihr eine Zahl geben konnte.
    return (alloc ? alloc.rows.map(r => r.key).concat(Object.keys(seen)) : Object.keys(seen))
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

  // v31.2: Bedarf und Bestand stehen in EINER Tabelle je Größe (vorher zwei
  // Blöcke zum selben Thema: Balkenliste oben, Eingabe-Raster unten). Die
  // Zählung je Größe wird dafür über den Schlüssel nachgeschlagen; die Zeile
  // „ohne Angabe" hat keinen Schlüssel und kommt zuletzt.
  const hasStock = !!(alloc && alloc.hasStock);
  const ready = !loading && !!result && !!result.fieldLabel;
  const tallyByKey: Record<string, { count: number; names: string[] }> = {};
  (result ? result.rows : []).forEach(r => { if (r.size && !r.optOut) tallyByKey[shirtSizeKey(r.size)] = { count: r.count, names: r.names }; });
  // v31.3: Abwahl-Antworten sind eine eigene Gruppe zwischen Größen und
  // „ohne Angabe" — sichtbar (der Organizer will wissen, wie viele keins
  // brauchen), aber ohne Bestand, ohne Soll/Ist und ohne Bestellposition.
  const optOutRows = result ? result.rows.filter(r => !!r.size && !!r.optOut) : [];
  const optOutTotal = result ? result.optOut : 0;
  const noneRow = result ? result.rows.filter(r => !r.size)[0] : undefined;
  const totalMissing = alloc ? alloc.rows.reduce((s, r) => s + r.missing, 0) : 0;
  // v31.3: Solange ein Termin nicht lesbar ist, ist JEDE dieser Zahlen eine
  // Untergrenze — dann wird sie auch so beschriftet (CLAUDE.md: ein
  // Lesefehler ist keine Null).
  const partial = skipped.length > 0;
  const atLeast = (n: number): string => partial ? (isDe ? `mind. ${n}` : `at least ${n}`) : String(n);
  const pillInTitle: React.CSSProperties = { textTransform: 'none', letterSpacing: 0 };
  const stickyTh: React.CSSProperties = { position: 'sticky', top: 0 };

  // Eine Zeile der Größen-Tabelle; `key` ist der Aufklapp-Schlüssel für die Namen.
  // v31.3: drei Arten — echte Größe, Abwahl („habe schon eins"), ohne Angabe.
  const renderSizeRow = (key: string, label: string, count: number, names: string[], kind: 'size' | 'optout' | 'none'): React.ReactElement => {
    const open = openSize === key;
    const pct = maxCount > 0 ? Math.round((count / maxCount) * 100) : 0;
    const a = (kind === 'size' && hasStock && alloc) ? alloc.rows.filter(x => x.key === key)[0] : undefined;
    const barColor = kind === 'none' ? 'var(--dex-orange, #ed8b00)' : kind === 'optout' ? 'var(--dex-gray-400, #9e9e9e)' : 'var(--dex-green, #86bc25)';
    return (
      <React.Fragment key={key}>
        <tr>
          <td style={{ fontWeight: 700, fontSize: '0.95rem', color: kind === 'none' ? 'var(--dex-orange-dark, #b35a00)' : kind === 'optout' ? 'var(--dex-gray-600)' : undefined, whiteSpace: kind === 'size' ? 'nowrap' : undefined }}>{label}</td>
          <td>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <strong style={{ minWidth: 24, textAlign: 'right', fontSize: '0.95rem' }}>{count}</strong>
              {/* Balken statt nur Zahl: Beim Bestellen zählt vor allem, welche Größen die Masse ausmachen. */}
              {count > 0
                ? <span style={{ flex: 1, minWidth: 50, height: 8, borderRadius: 999, background: 'var(--dex-gray-100)' }}>
                  <span style={{ display: 'block', height: 8, width: `${pct}%`, borderRadius: 999, background: barColor, transition: 'width 0.18s ease' }} />
                </span>
                : <span className="dex-ui-muted">{isDe ? 'niemand gewünscht' : 'nobody asked'}</span>}
            </div>
          </td>
          <td>
            {kind !== 'size' ? <span className="dex-ui-muted">—</span> : (
              <input type="number" min={0} inputMode="numeric" className="dex-ui-input dex-ui-input--sm" placeholder="0"
                aria-label={`${isDe ? 'Bestand' : 'Stock'} ${label}`} value={stockInput[key] ?? ''} style={{ width: 76, textAlign: 'center', fontWeight: 700 }}
                onChange={e => setStockInput(prev => ({ ...prev, [key]: e.target.value.replace(/[^\d]/g, '') }))} />
            )}
          </td>
          <td>
            {/* v30.88: Soll/Ist-Abgleich je Zeile, sobald ein Bestand eingetragen ist. */}
            {kind === 'none'
              ? <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'nachfragen' : 'ask them'}</span>
              : kind === 'optout'
                ? <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'kein Shirt nötig' : 'no shirt needed'}</span>
                : a
                  ? (a.missing > 0
                    ? <span className="dex-ui-pill dex-ui-pill--red">{isDe ? `fehlt ${a.missing}` : `${a.missing} short`}</span>
                    // v31.3: Mit einem gesperrten Termin ist „reicht" eine Zusage,
                    // die die Daten nicht decken — dann heißt es „reicht bisher".
                    : partial
                      ? <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'reicht bisher' : 'enough so far'}{a.spare > 0 ? ` (+${a.spare})` : ''}</span>
                      : <span className="dex-ui-pill dex-ui-pill--green"><Check size={12} /> {isDe ? 'reicht' : 'enough'}{a.spare > 0 ? ` (+${a.spare})` : ''}</span>)
                  : <span className="dex-ui-muted">—</span>}
          </td>
          <td style={{ textAlign: 'right', width: 40 }}>
            {count > 0 && (
              <button type="button" className="dex-ui-iconbtn" aria-expanded={open} title={isDe ? 'Personen anzeigen' : 'Show people'}
                aria-label={kind === 'size'
                  ? (isDe ? `Personen mit Größe ${label} anzeigen` : `Show people with size ${label}`)
                  : (isDe ? `Personen anzeigen: ${label}` : `Show people: ${label}`)} onClick={() => setOpenSize(open ? null : key)}>
                <span style={{ display: 'inline-flex', transition: 'transform 0.2s ease', transform: open ? 'rotate(180deg)' : 'none' }}><ChevronDown size={16} /></span>
              </button>
            )}
          </td>
        </tr>
        {open && (
          <tr><td colSpan={5} style={{ background: 'var(--dex-gray-50, #fafafa)', color: 'var(--dex-gray-600)', fontSize: '0.82rem', lineHeight: 1.6 }}>{names.join(' · ')}</td></tr>
        )}
      </React.Fragment>
    );
  };

  // v31.2: Genau ein Primär-Knopf — Speichern schreibt Daten, Excel ist die
  // Nebenaktion links im Fuß. Vorher standen zwei grüne Knöpfe im Inhalt.
  const footer = (
    <>
      {ready && (
        <span className="dex-ui-modal-foot-left">
          <button type="button" className="btn btn-outline dex-ui-btn-sm" disabled={xlsxBusy} onClick={() => { void downloadXlsx(); }}>
            <Download size={14} /> {xlsxBusy ? (isDe ? 'Wird erzeugt…' : 'Generating…') : (isDe ? 'Als Excel laden' : 'Download as Excel')}
          </button>
        </span>
      )}
      <button type="button" className="btn btn-secondary" onClick={props.onClose}>{isDe ? 'Schließen' : 'Close'}</button>
      {ready && (
        <button type="button" className="btn btn-primary" disabled={!stockDirty || stockSaving} onClick={() => { void saveStock(); }}>
          {stockSaving ? (isDe ? 'Speichert…' : 'Saving…') : (stockDirty ? (isDe ? 'Bestand speichern' : 'Save stock') : (isDe ? 'Bestand gespeichert' : 'Stock saved'))}
        </button>
      )}
    </>
  );

  return (
    <Modal open onClose={props.onClose} maxWidth={760} icon={<Shirt size={20} />} footer={footer}
      ariaLabel={isDe ? 'Benötigte T-Shirts' : 'T-shirts needed'} title={isDe ? 'Benötigte T-Shirts' : 'T-shirts needed'}
      subtitle={<>{props.event.title}{result && result.fieldLabel && <> · {isDe ? 'gezählt über das Feld' : 'counted via the field'} <strong>{result.fieldLabel}</strong></>}</>}>
      {loading && <p className="dex-ui-muted" style={{ margin: 0 }}>{isDe ? 'Trikotgrößen werden gelesen…' : 'Reading shirt sizes…'}</p>}

      {!loading && result && !result.fieldLabel && (
        <div className="dex-ui-callout dex-ui-callout--warn">
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>
            {isDe
              ? <>Dieses Event hat kein Abfragefeld, das nach einer Trikot- oder Konfektionsgröße aussieht.
                Lege im Assistenten unter <strong>Felder</strong> ein Feld an, dessen Bezeichnung die Größe
                benennt (z.B. &bdquo;T-Shirt Größe&ldquo; oder &bdquo;Trikotgröße&ldquo;) — danach zählt diese Ansicht
                automatisch mit, und die Größe steht auch am Check-in-Tisch.</>
              : <>This event has no form field that looks like a shirt or clothing size.
                In the wizard, under <strong>Fields</strong>, add a field whose label names the size
                (e.g. &ldquo;T-shirt size&rdquo; or &ldquo;Jersey size&rdquo;) — this view then counts automatically,
                and the size also shows at the check-in desk.</>}
          </span>
        </div>
      )}

      {!loading && result && result.fieldLabel && (
        <>
          {skipped.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--danger">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span>
                {isDe
                  ? <>Für {skipped.length === 1 ? 'diesen Termin' : 'diese Termine'} konnte die Teilnehmerliste nicht gelesen werden —
                    die Zahlen unten sind deshalb unvollständig: <strong>{skipped.join(', ')}</strong>.</>
                  : <>The attendee list of {skipped.length === 1 ? 'this date' : 'these dates'} could not be read —
                    the numbers below are therefore incomplete: <strong>{skipped.join(', ')}</strong>.</>}
              </span>
            </div>
          )}

          {/* v31.3: Widersprüchliche Zeilen benennen, statt still zu entscheiden. */}
          {conflicts.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--warn">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span>
                {isDe
                  ? <><strong>{conflicts.length === 1 ? 'Eine Person hat' : `${conflicts.length} Personen haben`} auf zwei Anmeldezeilen zwei verschiedene Größen.</strong>{' '}
                    Gerechnet wird mit der zuletzt geänderten Zeile (jeweils zuerst genannt) — bitte kurz nachfragen:{' '}
                    {conflicts.map(c => `${c.name}: ${c.values.join(' / ')}`).join(' · ')}</>
                  : <><strong>{conflicts.length === 1 ? 'One person has' : `${conflicts.length} people have`} two different sizes on two registration rows.</strong>{' '}
                    The most recently changed row is used (named first) — please double-check:{' '}
                    {conflicts.map(c => `${c.name}: ${c.values.join(' / ')}`).join(' · ')}</>}
              </span>
            </div>
          )}

          {/* v31.2: Kennzahlen zuerst — die Antwort auf „reicht es?" steht oben, bevor die Tabelle ins Detail geht.
              v31.3: „Shirts zu bestellen" statt „Größen gewünscht" — die Abwahl-Antworten sind keine Bestellung,
              und bei einem gesperrten Termin steht vor jeder Zahl „mind.". */}
          <div className="dex-ui-grid-3">
            <div className="dex-ui-kpi"><div className="dex-ui-kpi-value">{atLeast(result.total)}</div>
              <div className="dex-ui-kpi-label">{isDe ? (result.total === 1 ? 'Person angemeldet' : 'Personen angemeldet') : (result.total === 1 ? 'person registered' : 'people registered')}</div></div>
            <div className="dex-ui-kpi"><div className="dex-ui-kpi-value">{atLeast(result.sizeTotal)}</div>
              <div className="dex-ui-kpi-label">{isDe ? 'Shirts zu bestellen' : 'shirts to order'}</div>
              {optOutTotal > 0 && <div className="dex-ui-kpi-sub">{isDe ? `${optOutTotal} brauchen keins` : `${optOutTotal} need none`}</div>}</div>
            {hasStock
              ? <div className={cx('dex-ui-kpi', totalMissing > 0 ? 'dex-ui-kpi--orange' : 'dex-ui-kpi--green')}><div className="dex-ui-kpi-value">{atLeast(totalMissing)}</div>
                <div className="dex-ui-kpi-label">{isDe ? 'ohne Wunschgröße' : 'wished size short'}</div>
                {result.missing > 0 && <div className="dex-ui-kpi-sub">{isDe ? `+ ${result.missing} ohne Angabe` : `+ ${result.missing} without a size`}</div>}</div>
              : <div className={cx('dex-ui-kpi', result.missing > 0 && 'dex-ui-kpi--orange')}><div className="dex-ui-kpi-value">{atLeast(result.missing)}</div>
                <div className="dex-ui-kpi-label">{isDe ? 'ohne Größenangabe' : 'without a size'}</div></div>}
          </div>

          <div className="dex-ui-section">
            <div className="dex-ui-section-title">
              {isDe ? 'Bedarf und Bestand je Größe' : 'Need and stock per size'}
              {stockDirty && <span className="dex-ui-pill dex-ui-pill--orange" style={pillInTitle}>{isDe ? 'Bestand noch nicht gespeichert' : 'stock not saved yet'}</span>}
            </div>
            <p className="dex-ui-section-desc">
              {isDe
                ? <>Trag unter <strong>Bestand</strong> ein, wie viele Shirts du je Größe wirklich hast. Die App prüft, ob es reicht, und macht je Person einen
                  Gegenvorschlag — den sieht das Check-in-Team bei der Abholung direkt an der Person.</>
                : <>Under <strong>Stock</strong>, enter how many shirts you really have per size. The app checks whether that is enough and proposes an
                  alternative per person — the check-in team sees it right at the person when handing out.</>}
            </p>
            <div className="dex-ui-table-wrap">
              <table className="dex-ui-table">
                <thead>
                  <tr>
                    <th>{isDe ? 'Größe' : 'Size'}</th>
                    <th style={{ width: '36%' }}>{isDe ? 'Benötigt' : 'Needed'}</th>
                    <th>{isDe ? 'Bestand' : 'Stock'}</th>
                    <th>{isDe ? 'Reicht es?' : 'Enough?'}</th>
                    <th><span className="dex-ui-sr-only">{isDe ? 'Personen' : 'People'}</span></th>
                  </tr>
                </thead>
                <tbody>
                  {stockKeys.map(s => { const t = tallyByKey[s.key]; return renderSizeRow(s.key, s.label, t ? t.count : 0, t ? t.names : [], 'size'); })}
                  {/* v31.3: Abwahl-Antworten stehen mit ihrem echten Wortlaut in der
                      Liste — der Organizer sieht, wie viele keins brauchen, ohne dass
                      daraus eine Bestellposition wird. */}
                  {optOutRows.map(r => renderSizeRow(`__opt__${shirtSizeKey(r.size)}`, r.size, r.count, r.names, 'optout'))}
                  {noneRow && renderSizeRow('__none__', isDe ? 'ohne Angabe' : 'no answer', noneRow.count, noneRow.names, 'none')}
                </tbody>
              </table>
            </div>
            <div className="dex-ui-inline" style={{ marginTop: 10 }}>
              <input type="text" className="dex-ui-input dex-ui-input--sm" value={newSize} onChange={e => setNewSize(e.target.value)} style={{ maxWidth: 200 }}
                placeholder={isDe ? 'Weitere Größe, z.B. XXL' : 'Another size, e.g. XXL'} aria-label={isDe ? 'Weitere Größe' : 'Another size'} />
              <button
                type="button"
                className="dex-ui-textbtn"
                // v31.3: Nur echte Größen — sonst legt man eine Bestandszeile für
                // etwas an, das niemand tragen kann.
                disabled={!splitShirtSize(newSize).isSize}
                onClick={() => {
                  if (!splitShirtSize(newSize).isSize) return;
                  const k = shirtSizeKey(newSize);
                  if (!k) return;
                  setStockInput(prev => (prev[k] !== undefined ? prev : { ...prev, [k]: '' }));
                  setNewSize('');
                }}
              >
                <Plus size={14} /> {isDe ? 'Größe ergänzen' : 'Add size'}
              </button>
              <span className="dex-ui-muted">
                {isDe
                  ? 'Erscheint als neue Zeile — für Reserve-Größen, die niemand gewünscht hat. Muss auf ein Größenkürzel enden (S, M, L, XL …), gern mit Vorsatz („Herrengröße XL").'
                  : 'Adds a row — for spare sizes nobody asked for. Has to end in a size (S, M, L, XL …), a prefix is fine (“Men XL”).'}
              </span>
            </div>

            {result.missing > 0 && (
              <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 12 }}>
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>
                  {isDe
                    ? <><strong>{result.missing} {result.missing === 1 ? 'Person hat' : 'Personen haben'} keine Größe angegeben.</strong>{' '}
                      Die Namen stehen in der Zeile &bdquo;ohne Angabe&ldquo; — frag dort nach, bevor du bestellst.
                      Sonst fehlt am Lauftag genau diese Anzahl Trikots.</>
                    : <><strong>{result.missing} {result.missing === 1 ? 'person has' : 'people have'} not given a size.</strong>{' '}
                      Their names are in the &ldquo;no answer&rdquo; row — ask them before you order.
                      Otherwise exactly that many shirts will be missing on race day.</>}
                </span>
              </div>
            )}
          </div>

          {/* v30.88: Gegenvorschläge je Person — erst, wenn ein Bestand eingetragen ist. */}
          {alloc && alloc.hasStock && (
            <div className="dex-ui-section">
              <div className="dex-ui-section-title">
                {isDe ? 'Wer bekommt eine andere Größe?' : 'Who gets a different size?'}
                {proposals.length > 0 && <span className="dex-ui-pill dex-ui-pill--gray" style={pillInTitle}>{proposals.length}</span>}
              </div>
              {proposals.length === 0 && alloc.noneLeft.length === 0 ? (
                /* v31.3: „Reicht für alle" nur, wenn es auch stimmt. Der Kasten hing
                   bis dahin allein an den Gegenvorschlägen — Personen ohne Angabe
                   verbrauchen nichts und tauchen dort nie auf, ein gesperrter Termin
                   ebenso wenig. Beides macht die Aussage falsch, nicht ungenau. */
                (partial || result.missing > 0) ? (
                  <div className="dex-ui-callout dex-ui-callout--warn">
                    <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                    <span>
                      {isDe
                        ? <><strong>Für die abgegebenen Wünsche reicht der Bestand.</strong>{' '}
                          {result.missing > 0 && <>Ob er für ALLE reicht, lässt sich nicht sagen: {result.missing} {result.missing === 1 ? 'Person hat' : 'Personen haben'} keine Größe angegeben —
                            frag über die Zeile &bdquo;ohne Angabe&ldquo; nach und trag die Antwort ein. </>}
                          {partial && <>Außerdem {skipped.length === 1 ? 'ist ein Termin' : 'sind Termine'} nicht lesbar ({skipped.join(', ')}) — die dort Angemeldeten fehlen in dieser Rechnung.</>}</>
                        : <><strong>The stock covers every size that was given.</strong>{' '}
                          {result.missing > 0 && <>Whether it covers EVERYONE cannot be said: {result.missing} {result.missing === 1 ? 'person has' : 'people have'} not given a size —
                            ask them via the &ldquo;no answer&rdquo; row and record the answer. </>}
                          {partial && <>On top of that, {skipped.length === 1 ? 'one date is' : 'some dates are'} unreadable ({skipped.join(', ')}) — those registrations are missing from this calculation.</>}</>}
                    </span>
                  </div>
                ) : (
                  <div className="dex-ui-callout dex-ui-callout--success">
                    <span className="dex-ui-callout-icon"><Check size={16} /></span>
                    <span>
                      {isDe
                        ? <><strong>Der Bestand reicht für alle Wünsche.</strong> Jede Person bekommt ihre Größe.</>
                        : <><strong>The stock covers every wish.</strong> Everyone gets their size.</>}
                    </span>
                  </div>
                )
              ) : (
                <>
                  <p className="dex-ui-section-desc">
                    {isDe
                      ? 'Reihenfolge nach Teilnehmer-ID: Wer zuerst angemeldet war, bekommt seine Wunschgröße. Ausweichgröße = nächste mit Rest, zuerst eine Nummer größer.'
                      : 'Order by attendee ID: whoever registered first gets their wished size. Alternative = next size with stock left, one size up first.'}
                  </p>
                  <div className="dex-ui-table-wrap" style={{ maxHeight: 260, overflowY: 'auto' }}>
                    <table className="dex-ui-table">
                      <thead>
                        <tr>
                          <th style={stickyTh}>ID</th>
                          <th style={stickyTh}>{isDe ? 'Person' : 'Person'}</th>
                          <th style={stickyTh}>{isDe ? 'Wunsch' : 'Wish'}</th>
                          <th style={stickyTh}>{isDe ? 'Vorschlag' : 'Proposal'}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {proposals.map((p, i) => (
                          <tr key={i}>
                            <td style={{ color: 'var(--dex-gray-400)' }}>{p.tid ?? '—'}</td>
                            <td style={{ fontWeight: 600 }}>{p.name}</td>
                            <td><span style={{ textDecoration: 'line-through', color: 'var(--dex-gray-500)' }}>{p.wish}</span></td>
                            <td>
                              {p.proposal
                                ? <span className="dex-ui-pill dex-ui-pill--green">{p.proposal}</span>
                                : <span className="dex-ui-pill dex-ui-pill--red">{isDe ? 'keine Größe mehr vorrätig' : 'no size left in stock'}</span>}
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
        </>
      )}
    </Modal>
  );
}
