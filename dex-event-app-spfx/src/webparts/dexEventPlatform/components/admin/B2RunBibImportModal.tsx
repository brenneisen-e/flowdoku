/**
 * v30.48: Import der offiziellen B2Run-Startnummern.
 *
 * Zwei Schritte, bewusst getrennt: erst **prüfen und anzeigen**, dann auf
 * Knopfdruck schreiben. Der Grund ist der Nachrücker-Fall — dort wandert eine
 * Startnummer von einer Person zur anderen, und das muss der Organizer sehen
 * BEVOR es in den Daten steht, weil er dieselbe Ummeldung anschließend beim
 * Veranstalter machen muss. Ein Import, der sofort schreibt, würde ihm genau
 * die Liste vorenthalten, für die er den Import gemacht hat.
 *
 * Die Zuordnung „wer ersetzt wen" wird nicht geraten: Sie steht seit v17.15 in
 * `ReplacedByParticipantEmail` / `ReplacedParticipantEmail` auf den Zeilen.
 * Siehe `utils/b2runBibImport.ts`.
 */
import * as React from 'react';
import Modal from '../Modal';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { useCurrentUser } from '../../context/UserContext';
import { DeloitteEvent } from '../../types';
import { SPRegistration, EventService } from '../../services/EventService';
import { parseBibSheet, buildBibReport, suggestOrphanPairs, BibImportReport, BibMatch } from '../../utils/b2runBibImport';
import { StoredB2RunTodo, b2runNameOf } from '../../utils/b2runTodos';
import { cx } from '../dexUi';
import { InfoTooltip } from '../InfoTooltip';
import { AlertCircle, Check, Download, Hash } from '../Icons';

const nameOf = (r?: SPRegistration): string =>
  r ? `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || (r.ParticipantEmail || '') : '—';

// v31.2: Jede Gruppe des Abgleichs ist eine Tabelle mit denselben Spalten-
// Regeln (ruhiger Kopf, Zeilen-Hover, Status-Pill rechts). Vorher waren es
// sechs verschieden gebaute Kästen mit Fließtext-Zeilen — dieselbe Information
// in drei Schriftgrößen. Die kleinen Helfer halten das JSX unten lesbar.
const Tbl = (p: { head: string[]; children: React.ReactNode }): React.ReactElement => (
  <div className="dex-ui-table-wrap">
    <table className="dex-ui-table">
      <thead><tr>{p.head.map(h => <th key={h}>{h}</th>)}</tr></thead>
      <tbody>{p.children}</tbody>
    </table>
  </div>
);
/** Eine Zeile — jedes Kind wird eine Zelle. */
const Row = (p: { children: React.ReactNode }): React.ReactElement => (
  <tr>{React.Children.map(p.children, (c, i) => <td key={i}>{c}</td>)}</tr>
);
/** Startnummer — dicktengleich, damit die Spalte bündig bleibt. */
const Bib = (p: { children: React.ReactNode }): React.ReactElement => (
  <span style={{ fontFamily: 'monospace', fontWeight: 700 }}>{p.children}</span>
);
/** Abgemeldete Person — durchgestrichen, wie in der Excel „gemeldet, läuft nicht". */
const Gone = (p: { children: React.ReactNode }): React.ReactElement => (
  <span style={{ textDecoration: 'line-through', color: 'var(--dex-gray-500)' }}>{p.children}</span>
);
/** Ablauf-Zeile: Nummer (Haken, wenn erledigt), Titel, Hinweis, rechts der Knopf. */
const Step = (p: { n: number; done?: boolean; pending?: boolean; title: string; hint: React.ReactNode; children: React.ReactNode }): React.ReactElement => (
  <div className={cx('dex-ui-step', p.pending && 'is-pending', p.done && 'is-done')}>
    <span className="dex-ui-step-num">{p.done ? <Check size={14} /> : p.n}</span>
    <div className="dex-ui-step-body">
      <div className="dex-ui-step-title">{p.title}</div>
      <div className="dex-ui-step-hint">{p.hint}</div>
    </div>
    <div className="dex-ui-step-action">{p.children}</div>
  </div>
);
const Kpi = (p: { value: number; label: string; tone?: string | false }): React.ReactElement => (
  <div className={cx('dex-ui-kpi', p.tone)}>
    <div className="dex-ui-kpi-value">{p.value}</div>
    <div className="dex-ui-kpi-label">{p.label}</div>
  </div>
);

export default function B2RunBibImportModal(props: {
  event: DeloitteEvent;
  service: EventService;
  onClose: () => void;
  onDone: () => void;
}): React.ReactElement {
  const { getAllRegistrations, refreshEvents } = useEvents();
  const { showAlert } = useDialog();
  const { currentUser } = useCurrentUser();
  const [busy, setBusy] = React.useState(false);
  const [progress, setProgress] = React.useState('');
  const [report, setReport] = React.useState<BibImportReport | null>(null);
  const [fileName, setFileName] = React.useState('');
  const [written, setWritten] = React.useState<{ ok: number; failed: number; todos: number; todoSaved: boolean } | null>(null);
  /**
   * v30.54: Zuordnung „freie Nummer → Person ohne Nummer" (Startnummer → E-Mail).
   *
   * `''` heißt ausdrücklich „niemand — Nummer verfällt". Vorbelegt wird der
   * zeitliche Vorschlag aus `suggestOrphanPairs`; die Entscheidung trifft der
   * Organizer, weil die Paarung im Gegensatz zur Nachrück-Kette NICHT in den
   * Daten steht (s. dort).
   */
  const [orphanAssign, setOrphanAssign] = React.useState<Record<string, string>>({});

  const readFile = async (file: File): Promise<void> => {
    setBusy(true); setProgress('Datei wird gelesen…'); setReport(null); setWritten(null);
    try {
      const buf = await file.arrayBuffer();
      // Bundle-Regel: xlsx MUSS dynamisch importiert werden (s. HotelImportModal).
      const XLSX = await import('xlsx');
      const wb = XLSX.read(buf, { type: 'array', cellDates: false });
      const ws = wb.Sheets[wb.SheetNames[0]];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const table = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' }) as any[][];
      const parsed = parseBibSheet(table);
      if (parsed.error) { await showAlert(parsed.error, { variant: 'error' }); return; }
      setProgress('Teilnehmer werden geladen…');
      // v30.67 (Review): Auch eine TEILWEISE gelesene Liste ist keine Basis —
      // die fehlenden Personen würden als „nicht gefunden" gelten.
      let readFailed = false;
      const regs = await getAllRegistrations(props.event.id, () => { readFailed = true; });
      if (readFailed) {
        await showAlert('Die Teilnehmerliste konnte gerade nicht (vollständig) gelesen werden — ohne sie kann nichts zugeordnet werden. Bitte später erneut versuchen.', { variant: 'error' });
        return;
      }
      if (regs.length === 0) {
        await showAlert('Die Teilnehmerliste dieses Events ist leer — ohne sie kann nichts zugeordnet werden.', { variant: 'error' });
        return;
      }
      setFileName(file.name);
      const rep = buildBibReport(parsed.rows, regs);
      setReport(rep);
      setOrphanAssign(suggestOrphanPairs(rep));
    } catch (err) {
      console.warn('[DEX] B2Run-Import failed:', err);
      await showAlert('Die Datei konnte nicht gelesen werden. Bitte die unveränderte Excel des Veranstalters verwenden.', { variant: 'error' });
    } finally { setBusy(false); setProgress(''); }
  };

  const write = async (): Promise<void> => {
    if (!report || !props.event.subsiteUrl) return;
    // v30.54: direkte + Nachrücker-Zuordnungen PLUS die vom Organizer
    // bestätigten Zuordnungen freier Nummern.
    const toWrite: Array<{ id: number; bib: string }> = report.matches
      .filter(m => (m.kind === 'direct' || m.kind === 'transfer') && m.target && m.row.bib)
      .map(m => ({ id: m.target!.Id, bib: m.row.bib }));
    for (const [bib, email] of Object.entries(orphanAssign)) {
      if (!email) continue;
      const person = report.missingFromFile.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === email);
      if (person) toWrite.push({ id: person.Id, bib });
    }
    setBusy(true);
    // Die Spalte gibt es in Bestands-Listen nicht — erst anlegen, sonst laufen
    // alle Schreibvorgänge auf einen Feldfehler und der Import meldet 0/300.
    setProgress('Spalte „Startnummer" wird geprüft…');
    const colOk = await props.service.ensureStartNumberColumn(props.event.subsiteUrl);
    if (!colOk) {
      setBusy(false); setProgress('');
      await showAlert('Die Spalte „Startnummer" konnte in der Teilnehmerliste nicht angelegt werden. Ohne sie kann nichts geschrieben werden.', { variant: 'error' });
      return;
    }
    let ok = 0; let failed = 0;
    // Sequentiell: Die Teilnehmerliste laeuft mit Item-Level-Security, und ein
    // Promise.all ueber 300 Zeilen ist die 429-Welle, vor der CLAUDE.md warnt.
    for (let i = 0; i < toWrite.length; i++) {
      const m = toWrite[i];
      setProgress(`Startnummern werden geschrieben… ${i + 1} / ${toWrite.length}`);
      try {
        await props.service.adminUpdateRegistration(
          props.event.subsiteUrl, m.id,
          { Startnummer: m.bib },
          { name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email, email: currentUser.email },
        );
        ok++;
      } catch { failed++; }
    }
    // v30.55: Die Aufgaben beim Veranstalter FESTHALTEN.
    //
    // Sie lassen sich hinterher nicht mehr ableiten: Die Startnummer einer
    // abgemeldeten Person steht nur in dieser Datei, nie in DEX — geschrieben
    // wird sie ausschließlich der Person, die läuft. Nach dem Import sieht der
    // Datenstand deshalb aus, als wäre nie etwas zu tun gewesen. Genau daran
    // ist die Aufgabenliste in v30.54 gescheitert: „Nichts offen", während beim
    // Veranstalter neun Ummeldungen warteten. Dass DEX die Nummer geschrieben
    // hat, heißt nicht, dass der Veranstalter davon weiß.
    setProgress('Aufgabenliste wird gespeichert…');
    const nowIso = new Date().toISOString();
    const todos: StoredB2RunTodo[] = [];
    for (const m of report.matches) {
      if (m.kind === 'transfer' && m.target) {
        todos.push({
          key: `transfer|${m.row.bib}|${(m.target.ParticipantEmail || '').toLowerCase().trim()}`,
          kind: 'transfer', bib: m.row.bib, certain: true, ts: nowIso,
          fromName: b2runNameOf(m.listed!), fromEmail: m.listed?.ParticipantEmail,
          toName: b2runNameOf(m.target), toEmail: m.target.ParticipantEmail,
          action: `Startnummer ${m.row.bib} beim Veranstalter von ${b2runNameOf(m.listed!)} auf ${b2runNameOf(m.target)} ummelden.`,
        });
      } else if (m.kind === 'orphan') {
        const em = orphanAssign[m.row.bib] || '';
        const target = em ? report.missingFromFile.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === em) : undefined;
        todos.push(target
          ? {
            key: `assign|${m.row.bib}|${em}`,
            kind: 'assign', bib: m.row.bib, certain: false, ts: nowIso,
            fromName: b2runNameOf(m.listed!), fromEmail: m.listed?.ParticipantEmail,
            toName: b2runNameOf(target), toEmail: target.ParticipantEmail,
            action: `Startnummer ${m.row.bib} beim Veranstalter von ${b2runNameOf(m.listed!)} auf ${b2runNameOf(target)} ummelden — von dir zugeordnet, nicht aus der Nachrück-Kette.`,
          }
          : {
            key: `unregister|${m.row.bib}`,
            kind: 'unregister', bib: m.row.bib, certain: true, ts: nowIso,
            fromName: b2runNameOf(m.listed!), fromEmail: m.listed?.ParticipantEmail,
            action: `Startnummer ${m.row.bib} (${b2runNameOf(m.listed!)}) beim Veranstalter abmelden — es ist niemand da, der sie übernimmt.`,
          });
      }
    }
    for (const r of stillWithoutBib) {
      todos.push({
        key: `register|${(r.ParticipantEmail || '').toLowerCase().trim()}`,
        kind: 'register', bib: '', certain: true, ts: nowIso,
        toName: b2runNameOf(r), toEmail: r.ParticipantEmail,
        action: `${b2runNameOf(r)} beim Veranstalter nachmelden — angemeldet, aber ohne Startnummer.`,
      });
    }
    let todoSaved = true;
    if (todos.length > 0) {
      todoSaved = await props.service
        .patchEventOverridesValue(Number(props.event.id), '_b2runTodo', todos)
        .catch(() => false);
      // v30.56: Den lokalen Event-Stand nachziehen. `patchEventOverridesValue`
      // schreibt NUR nach SharePoint — das Event-Objekt im Speicher trägt
      // danach weiter die alten Overrides. Die Aufgabenliste liest genau von
      // dort und zeigte deshalb „Nichts offen", obwohl der Import gerade neun
      // Aufgaben gespeichert hatte.
      if (todoSaved) { try { await refreshEvents(); } catch { /* beim nächsten Laden */ } }
    }
    setBusy(false); setProgress('');
    setWritten({ ok, failed, todos: todos.length, todoSaved });
    props.onDone();
  };

  /**
   * v30.53: Den Abgleich als Excel herunterladen.
   *
   * Die Vorschau am Bildschirm beantwortet „was passiert jetzt?" — die Datei
   * beantwortet „was muss ich beim Veranstalter noch tun?". Das sind zwei
   * verschiedene Momente: Die Ummeldungen macht man später, an einem anderen
   * Rechner, womöglich am Telefon. Ein Fenster, das man dafür offen halten
   * muss, ist genau der Grund, warum Dinge liegen bleiben.
   *
   * EIN Blatt mit einer Spalte „Was ist zu tun" statt vier Blättern: Die
   * Liste wird abgearbeitet, nicht ausgewertet — und beim Abarbeiten will
   * man filtern, nicht zwischen Reitern springen.
   */
  const downloadReport = async (): Promise<void> => {
    if (!report || busy) return;
    setBusy(true);
    setProgress('Excel wird erzeugt…');
    try {
      const rows: (string | number)[][] = [[
        'Was ist zu tun', 'Startnummer', 'Gemeldete Person', 'E-Mail (gemeldet)',
        'Nummer geht an', 'E-Mail (neu)', 'Startblock laut Datei', 'Hinweis',
      ]];
      for (const m of report.matches) {
        if (m.kind === 'direct') {
          rows.push([
            'Nichts — Nummer bleibt bei der Person', m.row.bib, nameOf(m.listed), m.row.email,
            '', '', m.row.block,
            m.blockMismatch ? `Startblock in DEX: ${m.blockMismatch.dex}` : '',
          ]);
        } else if (m.kind === 'transfer') {
          rows.push([
            'UMMELDEN beim Veranstalter', m.row.bib, nameOf(m.listed), m.row.email,
            nameOf(m.target), m.target?.ParticipantEmail || '', m.row.block,
            (m.chain && m.chain.length > 1) ? `über ${m.chain.length - 1} weitere Abmeldung(en)` : '',
          ]);
        } else if (m.kind === 'orphan') {
          const em = orphanAssign[m.row.bib] || '';
          const target = em ? report.missingFromFile.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === em) : undefined;
          rows.push(target
            ? [
              'UMMELDEN beim Veranstalter', m.row.bib, nameOf(m.listed), m.row.email,
              nameOf(target), target.ParticipantEmail || '', m.row.block,
              'Von dir zugeordnet — nicht aus der Nachrück-Kette',
            ]
            : [
              'ABMELDEN beim Veranstalter', m.row.bib, nameOf(m.listed), m.row.email,
              '', '', m.row.block, 'Abgemeldet, niemand übernimmt die Nummer',
            ]);
        } else {
          rows.push([
            'Prüfen — Adresse in DEX unbekannt', m.row.bib,
            `${m.row.firstName} ${m.row.lastName}`.trim(), m.row.email,
            '', '', m.row.block, 'Verdacht: zweite Schreibweise derselben Person',
          ]);
        }
      }
      for (const r of stillWithoutBib) {
        rows.push([
          'NACHMELDEN beim Veranstalter — keine Startnummer', '', nameOf(r), r.ParticipantEmail || '',
          '', '', '', 'In DEX angemeldet, steht nicht in der Datei',
        ]);
      }
      for (const bib of report.duplicateBibs) {
        rows.push(['Prüfen — Startnummer doppelt in der Datei', bib, '', '', '', '', '', '']);
      }
      const XLSX = await import('xlsx');
      const ws = XLSX.utils.aoa_to_sheet(rows);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = [{ wch: 44 }, { wch: 13 }, { wch: 26 }, { wch: 32 }, { wch: 26 }, { wch: 32 }, { wch: 30 }, { wch: 40 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Startnummern-Abgleich');
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const safeName = (props.event.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Startnummern_Abgleich_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
    } catch (err) {
      console.warn('[DEX] Abgleich-Export fehlgeschlagen:', err);
      await showAlert('Die Excel-Datei konnte nicht erzeugt werden — bitte erneut versuchen.', { variant: 'error' });
    } finally { setBusy(false); setProgress(''); }
  };

  const group = (kind: BibMatch['kind']): BibMatch[] => (report ? report.matches.filter(m => m.kind === kind) : []);
  // v30.54: Wer eine freie Nummer zugeordnet bekommen hat, zählt nicht mehr
  // als „ohne Startnummer" — und die Zuordnungen zählen mit beim Schreiben.
  const assignedEmails = new Set(Object.values(orphanAssign).filter(Boolean));
  const stillWithoutBib = report
    ? report.missingFromFile.filter(r => !assignedEmails.has((r.ParticipantEmail || '').toLowerCase().trim()))
    : [];
  const assignedCount = assignedEmails.size;
  const transfers = group('transfer');
  const orphans = group('orphan');
  const unknowns = group('unknown');
  const directs = group('direct');
  const mismatches = directs.filter(m => m.blockMismatch);

  // v31.2: Zähler an EINER Stelle — Kennzahlen, Schritt-Hinweis und
  // Zusammenfassung lesen dieselben Werte, sonst zeigt das Fenster drei
  // verschiedene Summen.
  const writeCount = directs.length + transfers.length + assignedCount;
  const remeldCount = transfers.length + assignedCount;
  const checkCount = unknowns.length + mismatches.length + (report ? report.duplicateBibs.length : 0);
  const nothingToDo = !!report && orphans.length === 0 && transfers.length === 0 && stillWithoutBib.length === 0 && checkCount === 0;
  const waitLabel = progress || 'Bitte warten…';

  return (
    <Modal open onClose={props.onClose} maxWidth={860} dismissable={!busy}
      ariaLabel="Startnummern importieren (B2Run)" title="Startnummern importieren (B2Run)" icon={<Hash size={20} />}
      subtitle={<>Die Rücklauf-Datei des Veranstalters — deine gemeldete Liste, ergänzt um die Spalte <strong>Startnummer</strong>. DEX ordnet über die E-Mail-Adresse zu und zeigt dir, was passieren würde, bevor etwas geschrieben wird.</>}
      footer={
        <button type="button" className="btn btn-secondary" disabled={busy} onClick={props.onClose}>
          {written ? 'Schließen' : 'Abbrechen'}
        </button>
      }
    >
      <div className="dex-ui-stack" style={{ fontSize: '0.87rem', lineHeight: 1.55 }}>
        {/* v31.2: Der Ablauf als drei Zeilen, der Knopf sitzt in seiner Zeile
            („Der Knopf ist der Schritt"). Vorher stand die Erklärung oben,
            die Knöpfe unten — und dazwischen bis zu sechs Kästen. */}
        <Step n={1} done={!!report} title="Datei des Veranstalters wählen"
          hint={report
            ? <>Gelesen: <strong>{fileName}</strong> · {report.matches.length} Zeilen</>
            : <>Die unveränderte Excel (.xlsx / .csv). Wer sich nach der Meldung abgemeldet hat, steht dort noch — DEX zeigt dir, <strong>wer nachgerückt ist</strong> und die Nummer übernimmt.</>}
        >
          <label
            className={cx('btn', report ? 'btn-secondary' : 'btn-primary', 'dex-ui-btn-sm')}
            style={{ cursor: busy ? 'wait' : written ? 'not-allowed' : 'pointer', opacity: written ? 0.55 : 1 }}
          >
            {busy && !report ? waitLabel : report ? 'Andere Datei' : 'Datei wählen'}
            <input type="file" accept=".xlsx,.xlsm,.xls,.csv" style={{ display: 'none' }} disabled={busy || !!written}
              onChange={e => { const f = e.target.files?.[0]; if (f) void readFile(f); e.currentTarget.value = ''; }}
            />
          </label>
        </Step>

        <Step n={2} pending={!report} done={!!written} title="Zuordnung prüfen"
          hint={report
            ? <>Ummeldungen, freie Nummern und Nachmeldungen stehen unten. Die Liste kannst du als Excel mitnehmen — mit einer Spalte &bdquo;Was ist zu tun&ldquo; zum Abarbeiten beim Veranstalter.</>
            : 'Erscheint, sobald die Datei gelesen ist.'}
        >
          <button type="button" className="btn btn-secondary dex-ui-btn-sm" disabled={!report || busy}
            onClick={() => { void downloadReport(); }}
            title="Lädt den kompletten Abgleich als Excel — mit einer Spalte „Was ist zu tun“ zum Abarbeiten beim Veranstalter."
          >
            <Download size={14} /> Als Excel laden
          </button>
        </Step>

        <Step n={3} pending={!report} done={!!written} title="Startnummern schreiben"
          hint={report
            ? <>Schreibt <strong>{writeCount}</strong> Startnummer{writeCount === 1 ? '' : 'n'} in die Teilnehmerliste und merkt sich die Aufgaben für den Veranstalter unter &bdquo;Offen beim Veranstalter&ldquo;.</>
            : 'Erst nach dem Prüfen.'}
        >
          <button type="button" className={cx('btn', report ? 'btn-primary' : 'btn-secondary', 'dex-ui-btn-sm')}
            disabled={!report || busy || !!written} onClick={() => { void write(); }}
          >
            {busy && report ? waitLabel : 'Jetzt schreiben'}
          </button>
        </Step>

        {written && (
          <div className={cx('dex-ui-callout', written.failed ? 'dex-ui-callout--danger' : 'dex-ui-callout--success')}>
            <span className="dex-ui-callout-icon">{written.failed ? <AlertCircle size={16} /> : <Check size={16} />}</span>
            <div>
              <strong>{written.ok} Startnummer(n) geschrieben{written.failed ? `, ${written.failed} fehlgeschlagen` : ''}.</strong>
              {written.todos > 0 && (
                <div style={{ marginTop: 4 }}>
                  {written.todoSaved
                    ? <>{written.todos} Aufgabe{written.todos === 1 ? '' : 'n'} für den Veranstalter stehen jetzt unter <strong>&bdquo;Offen beim Veranstalter&ldquo;</strong> — dort abhaken, wenn du sie erledigt hast.</>
                    : <span style={{ color: 'var(--dex-red)' }}>Die Aufgabenliste konnte nicht gespeichert werden — bitte den Abgleich als Excel laden, damit nichts verlorengeht.</span>}
                </div>
              )}
            </div>
          </div>
        )}

        {report && (
          <>
            <div className="dex-ui-grid-auto">
              <Kpi value={writeCount} label="Wird geschrieben" tone="dex-ui-kpi--green" />
              <Kpi value={remeldCount} label="Ummelden" tone={remeldCount > 0 && 'dex-ui-kpi--orange'} />
              <Kpi value={stillWithoutBib.length} label="Nachmelden" tone={stillWithoutBib.length > 0 && 'dex-ui-kpi--orange'} />
              <Kpi value={checkCount} label="Prüfen" />
            </div>
            <p className="dex-ui-muted" style={{ margin: 0 }}>
              {directs.length} direkt zugeordnet, {transfers.length} an Nachrücker übertragen
              {assignedCount > 0 ? `, ${assignedCount} freie Nummer${assignedCount === 1 ? '' : 'n'} von dir zugeordnet` : ''}.
              {remeldCount > 0 && (
                <> Beim Veranstalter musst du <strong>{remeldCount}</strong> Ummeldung{remeldCount === 1 ? '' : 'en'} vornehmen.</>
              )}
            </p>

            {nothingToDo && (
              <div className="dex-ui-empty">
                <div className="dex-ui-empty-icon"><Check size={18} /></div>
                <div className="dex-ui-empty-title">Alles passt</div>
                Jede Nummer in der Datei gehört einer angemeldeten Person — beim Veranstalter ist nichts zu tun.
              </div>
            )}

            {/* v31.2: Die Entscheidung zuerst — sie verändert die Zähler oben und die
                Nachmelde-Liste unten. v30.54: Diese Nummern verfallen NICHT automatisch.
                Wer den Platz eines Abgemeldeten bekommen hat, steht nur dann in den Daten,
                wenn er über die Warteliste nachgerückt ist. Bei einer Direktanmeldung in
                die frei gewordene Kapazität — oder wenn der Organizer jemanden von Hand
                angelegt hat — gibt es keine Kette. Dann ist der Platz trotzdem besetzt:
                Genau die Personen zur Auswahl sind in DEX angemeldet, stehen aber nicht
                in der Datei. Deshalb hier zuordnen statt verfallen lassen. */}
            {orphans.length > 0 && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  Wer übernimmt die freie Nummer?
                  <span className="dex-ui-pill dex-ui-pill--orange">{orphans.length} · deine Entscheidung</span>
                </div>
                <p className="dex-ui-section-desc">
                  Diese Personen haben sich nach der Meldung abgemeldet, ohne dass DEX einen Nachrücker aufgezeichnet hat.
                  Die Nummer ist deshalb <strong>nicht</strong> automatisch verfallen — wähle, wer den Platz übernommen hat.
                  {' '}
                  <InfoTooltip text={<>Zur Auswahl stehen alle, die in DEX angemeldet sind, aber in der Datei fehlen. Der Vorschlag folgt der zeitlichen Reihenfolge: früheste Abmeldung, früheste Neuanmeldung. <strong>Bitte prüfen</strong> — anders als beim Nachrücken steht diese Zuordnung nicht in den Daten.</>} />
                </p>
                <Tbl head={['Startnummer', 'Abgemeldet', 'Nummer geht an']}>
                  {orphans.map(m => {
                    const bib = m.row.bib;
                    const chosen = orphanAssign[bib] || '';
                    // Bereits an eine andere Nummer vergebene Personen ausblenden,
                    // damit dieselbe Person nicht zwei Startnummern bekommt.
                    const takenElsewhere = new Set(
                      Object.entries(orphanAssign)
                        .filter(([k, v]) => k !== bib && v)
                        .map(([, v]) => v)
                    );
                    return (
                      <Row key={m.row.rowNo}>
                        <Bib>{bib}</Bib>
                        <Gone>{nameOf(m.listed)}</Gone>
                        <select
                          className="dex-ui-select dex-ui-input--sm"
                          value={chosen}
                          disabled={busy || !!written}
                          onChange={e => setOrphanAssign(prev => ({ ...prev, [bib]: e.target.value }))}
                          style={{ minWidth: 240 }}
                        >
                          <option value="">— niemand, Nummer verfällt —</option>
                          {report.missingFromFile
                            .filter(r => {
                              const em = (r.ParticipantEmail || '').toLowerCase().trim();
                              return em === chosen || !takenElsewhere.has(em);
                            })
                            .map(r => (
                              <option key={r.Id} value={(r.ParticipantEmail || '').toLowerCase().trim()}>
                                {nameOf(r)} · {r.ParticipantEmail}
                              </option>
                            ))}
                        </select>
                      </Row>
                    );
                  })}
                </Tbl>
              </div>
            )}

            {transfers.length > 0 && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  Beim Veranstalter ummelden
                  <span className="dex-ui-pill dex-ui-pill--orange">{transfers.length}</span>
                </div>
                <p className="dex-ui-section-desc">
                  Diese Personen stehen in der Datei, haben sich aber abgemeldet. In DEX geht die Nummer beim Schreiben
                  an die Person, die nachgerückt ist — beim Veranstalter musst du sie noch ummelden.
                </p>
                <Tbl head={['Startnummer', 'Abgemeldet', 'Nachgerückt', 'Status']}>
                  {transfers.map(m => (
                    <Row key={m.row.rowNo}>
                      <Bib>{m.row.bib}</Bib>
                      <Gone>{nameOf(m.listed)}</Gone>
                      <><strong>{nameOf(m.target)}</strong><span className="dex-ui-muted"> · {m.target?.ParticipantEmail}</span></>
                      <>
                        <span className="dex-ui-pill dex-ui-pill--orange">Ummelden</span>
                        {m.chain && m.chain.length > 1 && (
                          <span className="dex-ui-muted"> über {m.chain.length - 1} weitere Abmeldung{m.chain.length - 1 === 1 ? '' : 'en'}</span>
                        )}
                      </>
                    </Row>
                  ))}
                </Tbl>
              </div>
            )}

            {/* v30.54: Wer oben eine freie Nummer zugeordnet bekommen hat, ist
                hier nicht mehr „ohne Startnummer" — sonst widerspricht sich das
                Fenster in zwei Kästen übereinander. */}
            {stillWithoutBib.length > 0 && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  Beim Veranstalter nachmelden
                  <span className="dex-ui-pill dex-ui-pill--orange">{stillWithoutBib.length}</span>
                </div>
                <p className="dex-ui-section-desc">
                  In DEX angemeldet, aber nicht in der Datei — und ohne freie Nummer. Diese Personen musst du beim Veranstalter nachmelden.
                </p>
                <Tbl head={['Person', 'E-Mail', 'Status']}>
                  {stillWithoutBib.map(r => (
                    <Row key={r.Id}>
                      <strong>{nameOf(r)}</strong>
                      <span className="dex-ui-muted">{r.ParticipantEmail}</span>
                      <span className="dex-ui-pill dex-ui-pill--orange">Nachmelden</span>
                    </Row>
                  ))}
                </Tbl>
              </div>
            )}

            {checkCount > 0 && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  Ansehen, nichts blockiert
                  <span className="dex-ui-pill dex-ui-pill--gray">{checkCount}</span>
                </div>
                <p className="dex-ui-section-desc">
                  Die Startnummern werden trotzdem geschrieben — diese Zeilen solltest du aber einmal ansehen.
                </p>
                <Tbl head={['Startnummer', 'Person', 'Hinweis', 'Status']}>
                  {unknowns.map(m => (
                    <Row key={`u${m.row.rowNo}`}>
                      <Bib>{m.row.bib}</Bib>
                      <>{`${m.row.firstName} ${m.row.lastName}`.trim()}<span className="dex-ui-muted"> · {m.row.email || '(keine E-Mail)'}</span></>
                      <>Adresse steht in keiner DEX-Zeile — häufigster Grund: eine zweite Schreibweise derselben Person (SMTP-Adresse gegen Alias).</>
                      <span className="dex-ui-pill dex-ui-pill--gray">In DEX unbekannt</span>
                    </Row>
                  ))}
                  {mismatches.map(m => (
                    <Row key={`m${m.row.rowNo}`}>
                      <Bib>{m.row.bib}</Bib>
                      <>{nameOf(m.listed)}</>
                      <>Startblock DEX &bdquo;{m.blockMismatch!.dex}&ldquo; · Liste &bdquo;{m.blockMismatch!.file}&ldquo; — nur ein Hinweis, die Nummer wird zugeordnet.</>
                      <span className="dex-ui-pill dex-ui-pill--blue">Startblock weicht ab</span>
                    </Row>
                  ))}
                  {report.duplicateBibs.map(bib => (
                    <Row key={`d${bib}`}>
                      <Bib>{bib}</Bib>
                      <>—</>
                      <>Diese Startnummer steht mehrfach in der Datei.</>
                      <span className="dex-ui-pill dex-ui-pill--red">Doppelt in der Datei</span>
                    </Row>
                  ))}
                </Tbl>
              </div>
            )}
          </>
        )}
      </div>
    </Modal>
  );
}
