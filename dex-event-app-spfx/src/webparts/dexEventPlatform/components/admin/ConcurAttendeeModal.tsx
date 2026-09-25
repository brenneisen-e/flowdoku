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
import { EventService, SPRegistration } from '../../services/EventService';
import { splitName } from '../../utils/pastedRecipients';
import { CONCUR_ATTENDEE_COLUMNS, CONCUR_MAX_ATTENDEES, fillConcurAttendeeTemplate } from '../../utils/concurAttendeeXls';
import { mayBeTransliterated, NameParts, profileUmlautName, suggestUmlauts } from '../../utils/umlautNames';
import { AlertCircle, Check, Download, FileText } from '../Icons';

interface AttendeeRow {
  email: string;
  type: string;
  first: string;
  last: string;
  /** v31.97: Schreibweise mit Umlaut — aus dem Profil belegt oder nach Regeln vorgeschlagen. */
  sug?: NameParts;
  sugSource?: 'profil' | 'regel';
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
  /** v31.97: fürs Nachschlagen der Schreibweise im Profil; ohne Dienst nur Regel-Vorschläge. */
  service?: EventService | null;
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
  // v31.97: Umlaute (s. utils/umlautNames). Haken je Person, Vorgabe an;
  // eigene Korrektur in der Tabelle gewinnt immer.
  const [lookup, setLookup] = React.useState<{ done: number; total: number } | null>(null);
  const [useSug, setUseSug] = React.useState<Record<string, boolean>>({});
  const [edits, setEdits] = React.useState<Record<string, NameParts>>({});

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

      // v31.97: Befund erster Concur-Import (25.09.2026) — „Benoehr",
      // „Gaessler" … abgelehnt, weil Concur „Benöhr" führt. Erst das Profil
      // fragen (Beleg), sonst ein Regel-Vorschlag. Nur Mitarbeitende: Gäste
      // gleicht Concur nicht über den Namen ab. Vier Anfragen parallel.
      const todo = list.filter(r => r.type === 'SYSEMP' && mayBeTransliterated(r.first + ' ' + r.last));
      if (todo.length === 0) return;
      setLookup({ done: 0, total: todo.length });
      let done = 0;
      let next = 0;
      const worker = async (): Promise<void> => {
        while (next < todo.length && !cancelled) {
          const r = todo[next++];
          let found: NameParts | null = null;
          if (props.service) {
            try {
              const p = await props.service.getUserProfileByEmail(r.email);
              const d = splitName(p.displayName || '');
              found = profileUmlautName(r, [{ first: p.firstName, last: p.lastName }, { first: d.vorname, last: d.nachname }]);
            } catch { /* nicht auflösbar — Regel-Vorschlag */ }
          }
          if (found) { r.sug = found; r.sugSource = 'profil'; }
          else {
            const sf = suggestUmlauts(r.first);
            const sl = suggestUmlauts(r.last);
            if (sf !== r.first || sl !== r.last) { r.sug = { first: sf, last: sl }; r.sugSource = 'regel'; }
          }
          done++;
          if (!cancelled) setLookup({ done, total: todo.length });
        }
      };
      await Promise.all([worker(), worker(), worker(), worker()]);
      if (cancelled) return;
      setRows(list.slice());
      setLookup(null);
    })().catch(err => {
      console.warn('[DEX] Concur-Liste: Lesen fehlgeschlagen', err);
      if (!cancelled) { setFailed([props.event.title]); setLoading(false); }
    });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.event.id]);

  /** Der Name, der in die Datei geht: eigene Korrektur > bestätigter Umlaut > DEX. */
  const finalName = (r: AttendeeRow): NameParts =>
    edits[r.email] || (r.sug && useSug[r.email] !== false ? r.sug : { first: r.first, last: r.last });

  // Gleicher Name unter zwei Adressen (SMTP vs. Alias, CLAUDE.md): Concur
  // erkennt keine Dubletten und bucht dann doppelt (Anleitungsblatt, Punkt 5).
  const nameDupes = React.useMemo(() => {
    const seen: Record<string, number> = {};
    rows.forEach(r => { const n = finalName(r); const k = (n.first + ' ' + n.last).toLowerCase().trim(); seen[k] = (seen[k] || 0) + 1; });
    return Object.keys(seen).filter(k => k && seen[k] > 1);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rows, edits, useSug]);
  const guests = rows.filter(r => r.type !== 'SYSEMP');
  const noName = rows.filter(r => { const n = finalName(r); return !n.first || !n.last; });
  const sugProfil = rows.filter(r => r.sugSource === 'profil').length;
  const sugRegel = rows.filter(r => r.sugSource === 'regel').length;
  const parts = Math.max(1, Math.ceil(rows.length / CONCUR_MAX_ATTENDEES));
  const blocked = loading || !!lookup || failed.length > 0 || rows.length === 0;
  const setName = (r: AttendeeRow, key: 'first' | 'last', v: string): void => {
    setEdits(e => ({ ...e, [r.email]: { ...finalName(r), [key]: v } }));
  };

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
      const data = slice.map(r => { const n = finalName(r); return [r.type, n.first.trim(), n.last.trim(), '', '', '']; });
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
      maxWidth={920}
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

            {/* v31.98: Die Typen erklärt statt nur benannt (Nutzer-Ansage
                25.09.2026: „erklären, dass SYSEMP der Status für Deloitte
                Employee ist — schön übersichtlich"). Zwei Karten mit Zahl,
                Bedeutung und dem, was Concur je Typ verlangt; darunter die
                drei Schritte in Concur. */}
            {rows.length > 0 && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  {isDe ? 'Wer kommt in die Datei?' : 'Who goes into the file?'}{' '}
                  <span className="dex-ui-pill dex-ui-pill--green">{rows.length} {isDe ? 'eingecheckt' : 'checked in'}</span>
                </div>
                <div className="dex-ui-grid-2">
                  <div className="dex-ui-card" style={{ padding: 14 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                      <span className="dex-ui-pill dex-ui-pill--gray">SYSEMP</span>
                      <strong>Deloitte Employee</strong>
                      <span className="dex-ui-muted" style={{ marginLeft: 'auto' }}>{rows.length - guests.length}</span>
                    </div>
                    <div className="dex-ui-muted" style={{ fontSize: 13, lineHeight: 1.45 }}>
                      {isDe
                        ? 'Alle mit Deloitte-Adresse. Concur findet sie über Vor- und Nachname — der muss exakt stimmen, mit Umlauten. „Attendee Title“ und „Company“ bleiben leer, sonst meldet Concur NOATNMATCH.'
                        : 'Everyone with a Deloitte address. Concur finds them by first and last name — it must match exactly, including umlauts. “Attendee Title” and “Company” stay empty, otherwise Concur reports NOATNMATCH.'}
                    </div>
                  </div>
                  <div className="dex-ui-card" style={{ padding: 14 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                      <span className="dex-ui-pill dex-ui-pill--orange">BUSGUEST</span>
                      <strong>Business Guest</strong>
                      <span className="dex-ui-muted" style={{ marginLeft: 'auto' }}>{guests.length}</span>
                    </div>
                    <div className="dex-ui-muted" style={{ fontSize: 13, lineHeight: 1.45 }}>
                      {isDe
                        ? 'Alle ohne Deloitte-Adresse. Für Gäste verlangt Concur „Attendee Title“ und „Company“ — die kennt DEX nicht, bitte in der Datei ergänzen. Andere Typen (z. B. CLIENT) stehen im Blatt „Attendee Types“.'
                        : 'Everyone without a Deloitte address. For guests Concur requires “Attendee Title” and “Company” — DEX does not know them, please add them in the file. Other types (e.g. CLIENT) are listed on the “Attendee Types” sheet.'}
                    </div>
                  </div>
                </div>
                <div className="dex-ui-muted" style={{ fontSize: 13, marginTop: 10 }}>
                  {isDe
                    ? <>In Concur: <strong>Beleg öffnen</strong> → <strong>Attendees</strong> → <strong>Import</strong> → Datei wählen. Vorlage: {ownTemplate ? ownTemplate.name : 'AttendeeImportTemplate.xls (Stand 07.04.2025)'}.</>
                    : <>In Concur: <strong>open the expense</strong> → <strong>Attendees</strong> → <strong>Import</strong> → pick the file. Template: {ownTemplate ? ownTemplate.name : 'AttendeeImportTemplate.xls (as of 2025-04-07)'}.</>}
                </div>
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

            {lookup && (
              <p className="dex-ui-muted" style={{ margin: 0 }}>
                {isDe
                  ? `Schreibweisen mit Umlaut werden im Profil nachgeschlagen … ${lookup.done} / ${lookup.total}`
                  : `Looking up spellings with umlauts in the profile … ${lookup.done} / ${lookup.total}`}
              </p>
            )}

            {!lookup && (sugProfil + sugRegel) > 0 && (
              <div className="dex-ui-callout dex-ui-callout--warn">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>
                  <strong>{isDe ? `${sugProfil + sugRegel} Namen mit Umlaut` : `${sugProfil + sugRegel} names with umlauts`}</strong>
                  {' — '}
                  {isDe
                    ? `In DEX stehen sie umschrieben („Benoehr"), Concur führt sie mit Umlaut („Benöhr") und lehnt jede Abweichung ab. ${sugProfil} ${sugProfil === 1 ? 'Schreibweise ist' : 'Schreibweisen sind'} im Profil belegt, ${sugRegel} ${sugRegel === 1 ? 'ist ein Vorschlag' : 'sind Vorschläge'} — bitte in der Spalte „Umlaut" prüfen. Meldet Concur trotzdem „Cannot import", den Namen unten direkt korrigieren und die Datei neu laden.`
                    : `DEX has them transliterated (“Benoehr”), Concur has them with umlauts (“Benöhr”) and rejects any difference. ${sugProfil} confirmed by the profile, ${sugRegel} suggested — please check the “Umlaut” column. If Concur still reports “Cannot import”, correct the name below and download again.`}
                </span>
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
                      <th>Umlaut</th>
                      <th>{isDe ? 'E-Mail (nicht in der Datei)' : 'Email (not in the file)'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map(r => {
                      const n = finalName(r);
                      return (
                        <tr key={r.email}>
                          <td style={{ whiteSpace: 'nowrap' }}>
                            <span className={r.type === 'SYSEMP' ? 'dex-ui-pill dex-ui-pill--sm dex-ui-pill--gray' : 'dex-ui-pill dex-ui-pill--sm dex-ui-pill--orange'}>{r.type}</span>
                            <div className="dex-ui-muted" style={{ fontSize: 11, marginTop: 2 }}>{r.type === 'SYSEMP' ? 'Deloitte Employee' : 'Business Guest'}</div>
                          </td>
                          {/* v31.97: direkt korrigierbar — für die Zeilen, die Concur ablehnt. */}
                          <td><input className="dex-ui-input dex-ui-input--sm" value={n.first} style={{ minWidth: 110 }} aria-label="First Name" onChange={e => setName(r, 'first', e.target.value)} /></td>
                          <td><input className="dex-ui-input dex-ui-input--sm" value={n.last} style={{ minWidth: 130 }} aria-label="Last Name" onChange={e => setName(r, 'last', e.target.value)} /></td>
                          <td style={{ whiteSpace: 'nowrap' }}>
                            {r.sug && !edits[r.email] && (
                              <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer' }} title={isDe ? `In DEX: ${r.first} ${r.last}` : `In DEX: ${r.first} ${r.last}`}>
                                <input
                                  type="checkbox"
                                  className="dex-ui-checkbox"
                                  checked={useSug[r.email] !== false}
                                  onChange={e => { const v = e.target.checked; setUseSug(u => ({ ...u, [r.email]: v })); }}
                                />
                                <span className={r.sugSource === 'profil' ? 'dex-ui-pill dex-ui-pill--sm dex-ui-pill--green' : 'dex-ui-pill dex-ui-pill--sm dex-ui-pill--orange'}>
                                  {r.sugSource === 'profil' ? (isDe ? 'aus Profil' : 'from profile') : (isDe ? 'Vorschlag' : 'suggested')}
                                </span>
                              </label>
                            )}
                            {edits[r.email] && (
                              <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => setEdits(e => { const x = { ...e }; delete x[r.email]; return x; })}>
                                {isDe ? 'Zurücksetzen' : 'Reset'}
                              </button>
                            )}
                          </td>
                          <td className="dex-ui-muted">{r.email}</td>
                        </tr>
                      );
                    })}
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
