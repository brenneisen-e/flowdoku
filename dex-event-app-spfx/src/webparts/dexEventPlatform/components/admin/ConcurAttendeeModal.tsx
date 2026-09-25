/**
 * v31.96: „Concur-Teilnehmerliste" — alle eingecheckten Personen eines
 * Events in die Concur-Vorlage AttendeeImportTemplate.xls.
 *
 * Nutzer-Ansage 25.09.2026 (B2Run Köln): „nicht neu erstellen, sondern in das
 * Template schreiben". Deshalb wird die Original-Datei gefüllt
 * (utils/concurAttendeeXls) — Anleitung, Typ-Liste, Formate und das
 * versteckte Blatt bleiben, wie Deloitte sie ausliefert.
 *
 * Wer drinsteht: Status „Eingecheckt" auf dem Event UND auf seinen Terminen.
 * Bei QR-Versand auf der Klammer liegt der Check-in auf der Klammer-Liste,
 * sonst auf den Termin-Listen (CLAUDE.md, v31.93) — also beide lesen und über
 * die Adresse zusammenführen. Ist eine Liste nicht lesbar, gibt es KEINE
 * Datei: Concur rechnet die Kosten auf die Liste um, eine stillschweigend
 * kürzere Liste wäre eine falsche Abrechnung.
 */
import * as React from 'react';
import Modal from '../Modal';
import { useEvents } from '../../context/EventContext';
import { DeloitteEvent } from '../../types';
import { SPRegistration } from '../../services/EventService';
import { splitName } from '../../utils/pastedRecipients';
import { CONCUR_ATTENDEE_COLUMNS, CONCUR_MAX_ATTENDEES, fillConcurAttendeeTemplate } from '../../utils/concurAttendeeXls';
import { AlertCircle, Check, Download, FileText } from '../Icons';

interface AttendeeRow {
  email: string;
  type: string;
  first: string;
  last: string;
}

/** Deloitte-Adressen sind Mitarbeitende (SYSEMP); alles andere ist Gast. */
function isDeloitteAddress(email: string): boolean {
  return /@([a-z0-9-]+\.)*deloitte\.[a-z.]+$/i.test(email);
}

function nameOf(r: SPRegistration): { first: string; last: string } {
  const v = (r.Vorname || '').trim();
  const n = (r.Nachname || '').trim();
  if (v || n) return { first: v, last: n };
  // „Nachname, Vorname" (Adressbuch) und „Vorname Nachname" — splitName kennt beide.
  const s = splitName(r.ParticipantName || '');
  return { first: s.vorname, last: s.nachname };
}

function downloadBytes(bytes: Uint8Array, filename: string): void {
  const blob = new Blob([bytes], { type: 'application/vnd.ms-excel' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
}

function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export default function ConcurAttendeeModal(props: {
  event: DeloitteEvent;
  isDe: boolean;
  onClose: () => void;
}): React.ReactElement {
  const { getAllRegistrations, events } = useEvents();
  const isDe = props.isDe;
  const [loading, setLoading] = React.useState(true);
  const [rows, setRows] = React.useState<AttendeeRow[]>([]);
  const [failed, setFailed] = React.useState<string[]>([]);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState('');
  const [ownTemplate, setOwnTemplate] = React.useState<{ name: string; bytes: Uint8Array } | null>(null);
  const [doneParts, setDoneParts] = React.useState<number[]>([]);

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      const children = events.filter(e => e.parentEventId === props.event.id);
      const targets = [props.event, ...children];
      const byEmail: Record<string, AttendeeRow> = {};
      const fail: string[] = [];
      for (const ev of targets) {
        if (!ev.subsiteUrl) continue;
        let ok = true;
        // Ein Lesefehler ist keine leere Liste (CLAUDE.md, v30.37/v30.67).
        const rs = await getAllRegistrations(ev.id, () => { ok = false; });
        if (!ok) { fail.push(ev.title); continue; }
        rs.forEach(r => {
          if (r.Status !== 'Eingecheckt') return;
          const email = (r.ParticipantEmail || r.Title || '').toLowerCase().trim();
          if (!email || byEmail[email]) return;
          const n = nameOf(r);
          byEmail[email] = { email, type: isDeloitteAddress(email) ? 'SYSEMP' : 'BUSGUEST', first: n.first, last: n.last };
        });
      }
      if (cancelled) return;
      const list = Object.keys(byEmail).map(k => byEmail[k]);
      list.sort((a, b) => (a.last + ' ' + a.first).localeCompare(b.last + ' ' + b.first, 'de'));
      setRows(list);
      setFailed(fail);
      setLoading(false);
    })().catch(err => {
      console.warn('[DEX] Concur-Liste: Lesen fehlgeschlagen', err);
      if (!cancelled) { setFailed([props.event.title]); setLoading(false); }
    });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.event.id]);

  // Gleicher Name unter zwei Adressen (SMTP vs. Alias, CLAUDE.md): Concur
  // erkennt keine Dubletten und bucht dann doppelt (Anleitungsblatt, Punkt 5).
  const nameDupes = React.useMemo(() => {
    const seen: Record<string, number> = {};
    rows.forEach(r => { const k = (r.first + ' ' + r.last).toLowerCase().trim(); seen[k] = (seen[k] || 0) + 1; });
    return Object.keys(seen).filter(k => k && seen[k] > 1);
  }, [rows]);
  const guests = rows.filter(r => r.type !== 'SYSEMP');
  const noName = rows.filter(r => !r.first || !r.last);
  const parts = Math.max(1, Math.ceil(rows.length / CONCUR_MAX_ATTENDEES));
  const blocked = loading || failed.length > 0 || rows.length === 0;

  const pickTemplate = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!f) return;
    f.arrayBuffer().then(buf => setOwnTemplate({ name: f.name, bytes: new Uint8Array(buf) })).catch(() => setError(isDe ? 'Die Datei konnte nicht gelesen werden.' : 'Could not read the file.'));
  };

  const build = async (part: number): Promise<void> => {
    if (busy || blocked) return;
    setBusy(true);
    setError('');
    try {
      const XLSX = await import('xlsx');
      const tpl = ownTemplate
        ? ownTemplate.bytes
        : base64ToBytes((await import('../../data/concurAttendeeTemplate')).CONCUR_ATTENDEE_TEMPLATE_B64);
      // Erst prüfen, ob es wirklich die Concur-Vorlage ist — eine falsche
      // Datei fällt sonst erst beim Upload in Concur auf.
      const wb = XLSX.read(tpl, { type: 'array', sheetRows: 2 });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const head = (XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })[0] || []) as string[];
      if (CONCUR_ATTENDEE_COLUMNS.some((c, i) => String(head[i] || '').trim() !== c)) {
        throw new Error(isDe
          ? `Das erste Blatt hat nicht die Concur-Spalten (${CONCUR_ATTENDEE_COLUMNS.join(', ')}).`
          : `The first sheet does not have the Concur columns (${CONCUR_ATTENDEE_COLUMNS.join(', ')}).`);
      }
      const slice = rows.slice(part * CONCUR_MAX_ATTENDEES, (part + 1) * CONCUR_MAX_ATTENDEES);
      // SYSEMP: Title und Company bleiben leer (Anleitungsblatt, Punkt 3 —
      // sonst NOATNMATCH). Custom5 (Client List) füllt DEX nie.
      const data = slice.map(r => [r.type, r.first, r.last, '', '', '']);
      // CFB ist im ESM-Build exportiert, in den Typen von xlsx 0.18 aber nicht.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const res = fillConcurAttendeeTemplate((XLSX as any).CFB, tpl, data);
      const safe = (props.event.title || 'Event').replace(/[^\w-]+/g, '_');
      const suffix = parts > 1 ? `_Teil${part + 1}` : '';
      downloadBytes(res.bytes, `Concur_Attendees_${safe}${suffix}_${new Date().toISOString().slice(0, 10)}.xls`);
      setDoneParts(d => (d.indexOf(part) >= 0 ? d : d.concat(part)));
    } catch (err) {
      console.warn('[DEX] Concur-Liste: Erzeugen fehlgeschlagen', err);
      setError((isDe ? 'Die Datei konnte nicht erzeugt werden: ' : 'Could not create the file: ') + String((err as Error)?.message || err));
    } finally {
      setBusy(false);
    }
  };

  const partLabel = (i: number): string => {
    const from = i * CONCUR_MAX_ATTENDEES + 1;
    const to = Math.min(rows.length, (i + 1) * CONCUR_MAX_ATTENDEES);
    return parts > 1
      ? (isDe ? `Teil ${i + 1} laden (${from}–${to})` : `Download part ${i + 1} (${from}–${to})`)
      : (isDe ? 'Concur-Datei laden' : 'Download Concur file');
  };

  return (
    <Modal
      open
      onClose={props.onClose}
      maxWidth={760}
      ariaLabel={isDe ? 'Concur-Teilnehmerliste' : 'Concur attendee list'}
      title={isDe ? 'Concur-Teilnehmerliste' : 'Concur attendee list'}
      subtitle={isDe
        ? 'Schreibt alle eingecheckten Personen in die Concur-Vorlage „AttendeeImportTemplate.xls" — die Datei, die du in Concur beim Beleg über „Import" hochlädst.'
        : 'Writes everyone checked in into the Concur template “AttendeeImportTemplate.xls” — the file you upload in Concur via “Import” on the expense.'}
      icon={<FileText size={20} />}
      footer={<>
        <div className="dex-ui-modal-foot-left">
          <label className="btn btn-secondary" style={{ cursor: 'pointer' }}>
            <FileText size={16} /> {isDe ? 'Eigene Vorlage verwenden' : 'Use own template'}
            <input type="file" accept=".xls" style={{ display: 'none' }} onChange={pickTemplate} />
          </label>
        </div>
        {Array.from({ length: parts }).map((_, i) => (
          <button key={i} type="button" className="btn btn-primary" disabled={busy || blocked} onClick={() => { void build(i); }}>
            {doneParts.indexOf(i) >= 0 ? <Check size={16} /> : <Download size={16} />} {busy ? (isDe ? 'Wird erzeugt…' : 'Creating…') : partLabel(i)}
          </button>
        ))}
      </>}
    >
      <div className="dex-ui-modal-body">
        {loading ? (
          <p className="dex-ui-muted" style={{ margin: 0 }}>{isDe ? 'Teilnehmerlisten werden gelesen…' : 'Reading participant lists…'}</p>
        ) : (
          <>
            {failed.length > 0 && (
              <div className="dex-ui-callout dex-ui-callout--danger">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>
                  <strong>{isDe ? 'Nicht jede Liste war lesbar' : 'Not every list was readable'}</strong> — {failed.join(', ')}.{' '}
                  {isDe
                    ? 'Eine Datei gibt es erst, wenn alle Listen gelesen sind: Eine unvollständige Liste würde in Concur die Kosten falsch verteilen. Dialog schließen und erneut öffnen; hilft das nicht, „Organizer-Berechtigungen reparieren".'
                    : 'No file until every list is read: an incomplete list would split the costs wrongly in Concur. Close and reopen; if that does not help, run “Repair organizer permissions”.'}
                </span>
              </div>
            )}

            {failed.length === 0 && rows.length === 0 && (
              <div className="dex-ui-empty">
                <div className="dex-ui-empty-icon"><AlertCircle size={20} /></div>
                <div className="dex-ui-empty-title">{isDe ? 'Noch niemand eingecheckt' : 'Nobody checked in yet'}</div>
                {isDe ? 'Die Liste enthält nur Personen mit Status „Eingecheckt".' : 'The list only contains people with status “Checked in”.'}
              </div>
            )}

            {rows.length > 0 && (
              <div className="dex-ui-callout dex-ui-callout--info">
                <span className="dex-ui-callout-icon"><FileText size={16} /></span>
                <span>
                  <strong>{rows.length} {isDe ? (rows.length === 1 ? 'Person eingecheckt' : 'Personen eingecheckt') : 'checked in'}</strong>
                  {' — '}
                  {isDe
                    ? `Deloitte-Adressen als SYSEMP (Titel und Firma bleiben leer, so verlangt es Concur), andere als BUSGUEST. Vorlage: ${ownTemplate ? ownTemplate.name : 'AttendeeImportTemplate.xls (Stand 07.04.2025)'}.`
                    : `Deloitte addresses as SYSEMP (title and company stay empty, as Concur requires), others as BUSGUEST. Template: ${ownTemplate ? ownTemplate.name : 'AttendeeImportTemplate.xls (as of 2025-04-07)'}.`}
                </span>
              </div>
            )}

            {rows.length > CONCUR_MAX_ATTENDEES && (
              <div className="dex-ui-callout dex-ui-callout--warn">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>{isDe
                  ? `Concur nimmt höchstens ${CONCUR_MAX_ATTENDEES} Personen je Beleg. Die Liste ist deshalb auf ${parts} Dateien verteilt — je Datei ein Beleg.`
                  : `Concur accepts at most ${CONCUR_MAX_ATTENDEES} attendees per expense. The list is split into ${parts} files — one expense each.`}</span>
              </div>
            )}

            {guests.length > 0 && (
              <div className="dex-ui-callout dex-ui-callout--warn">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>{isDe
                  ? `${guests.length} ${guests.length === 1 ? 'Person hat' : 'Personen haben'} keine Deloitte-Adresse (${guests.map(g => `${g.first} ${g.last}`.trim() || g.email).join(', ')}). Für Gäste verlangt Concur „Attendee Title" und „Company" — bitte in der Datei ergänzen.`
                  : `${guests.length} ${guests.length === 1 ? 'person has' : 'people have'} no Deloitte address (${guests.map(g => `${g.first} ${g.last}`.trim() || g.email).join(', ')}). Concur requires “Attendee Title” and “Company” for guests — please add them in the file.`}</span>
              </div>
            )}

            {noName.length > 0 && (
              <div className="dex-ui-callout dex-ui-callout--warn">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>{isDe
                  ? `Ohne vollständigen Namen: ${noName.map(r => r.email).join(', ')}. Concur gleicht Mitarbeitende über Vor- und Nachnamen ab — bitte in der Datei ergänzen.`
                  : `Without a full name: ${noName.map(r => r.email).join(', ')}. Concur matches employees by first and last name — please complete them in the file.`}</span>
              </div>
            )}

            {nameDupes.length > 0 && (
              <div className="dex-ui-callout dex-ui-callout--warn">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>{isDe
                  ? `Gleicher Name unter zwei Adressen: ${nameDupes.join(', ')}. Concur erkennt keine Dubletten und bucht dann doppelt — prüfen und eine Zeile löschen, falls es dieselbe Person ist.`
                  : `Same name under two addresses: ${nameDupes.join(', ')}. Concur cannot detect duplicates and would allocate twice — check and delete one row if it is the same person.`}</span>
              </div>
            )}

            {error && (
              <div className="dex-ui-callout dex-ui-callout--danger">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>{error}</span>
              </div>
            )}

            {rows.length > 0 && (
              <div className="dex-ui-table-wrap dex-ui-table-wrap--sticky" style={{ maxHeight: 360 }}>
                <table className="dex-ui-table dex-ui-table--compact">
                  <thead>
                    <tr>
                      <th>Attendee Type</th>
                      <th>First Name</th>
                      <th>Last Name</th>
                      <th>{isDe ? 'E-Mail (nicht in der Datei)' : 'Email (not in the file)'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map(r => (
                      <tr key={r.email}>
                        <td><span className={r.type === 'SYSEMP' ? 'dex-ui-pill dex-ui-pill--sm dex-ui-pill--gray' : 'dex-ui-pill dex-ui-pill--sm dex-ui-pill--orange'}>{r.type}</span></td>
                        <td>{r.first}</td>
                        <td>{r.last}</td>
                        <td className="dex-ui-muted">{r.email}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </>
        )}
      </div>
    </Modal>
  );
}
