/* AgendaAttendanceSection — v30.92 (Programmpunkte, Stufe 3;
 * docs/konzept-programmpunkte.md, 3.6).
 *
 * Wer war bei welchem Programmpunkt? Zwei Sichten auf dieselben Daten:
 *  - „nach Person": Matrix Person × Punkt, je Zelle ✓ mit Uhrzeit.
 *  - „nach Punkt": je Punkt Anzahl, Anteil, aufklappbar wer da war und wer
 *    fehlte.
 * Dazu Excel-Export (beide Sichten als Blätter) und — für Organizer — das
 * manuelle Setzen/Entfernen einer Anwesenheit mit Rückfrage und Audit-Eintrag
 * (ChangeLog `AgendaAttendanceSet`/`AgendaAttendanceRemoved`).
 *
 * Datenquelle ist `registrations` des Organizer Centers (Spalte
 * `AgendaCheckIns`, s. utils/agendaCheckIns). Nicht lesbar (`regsUnknown`) →
 * kein „0 anwesend", sondern der Zustand wird benannt (CLAUDE.md: Ein
 * Lesefehler ist keine Null). Die Bezeichnung kommt IMMER aus
 * `agendaTermSingular/Plural` des Events, nie fest „Programmpunkt". */
import * as React from 'react';
import { DeloitteEvent, AgendaItem } from '../../../types';
import { EventService, SPRegistration } from '../../../services/EventService';
import { parseAgendaCheckIns, formatMarkTime } from '../../../utils/agendaCheckIns';
import { agendaGroups, sortAgenda, groupLabel, groupDateLabel } from '../../../utils/agendaGroups';
import { downloadAttendanceCertificate, downloadAttendanceCertificates } from '../../../utils/attendanceCertificatePdf';
import { PersonContactHover } from '../../PersonContactHover';

export interface AgendaAttendanceSectionProps {
  event: DeloitteEvent;
  registrations: SPRegistration[];
  regsUnknown: boolean;
  isDe: boolean;
  canEdit: boolean;
  eventServiceRef: EventService | null;
  reloadRegistrations: () => Promise<unknown>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
}

const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

export const AgendaAttendanceSection: React.FC<AgendaAttendanceSectionProps> = (p) => {
  const { event, registrations, regsUnknown, isDe, canEdit, eventServiceRef, reloadRegistrations, showAlert, confirmDialog } = p;
  const [open, setOpen] = React.useState<boolean>(true);
  const [view, setView] = React.useState<'person' | 'point'>('point');
  const [openPoint, setOpenPoint] = React.useState<string | null>(null);
  const [busyKey, setBusyKey] = React.useState<string>('');
  const [xlsxBusy, setXlsxBusy] = React.useState(false);

  const termS = event.agendaTermSingular || (isDe ? 'Programmpunkt' : 'Agenda item');
  const termP = event.agendaTermPlural || (isDe ? 'Programmpunkte' : 'Agenda items');
  const items: AgendaItem[] = React.useMemo(() => sortAgenda(event.agenda || []), [event.agenda]);
  // v30.94: Cluster (utils/agendaGroups) für die Sicht „nach Punkt".
  const groups = React.useMemo(() => agendaGroups(items), [items]);
  const active = React.useMemo(
    () => registrations.filter(r => ACTIVE.indexOf(r.Status || '') >= 0).slice().sort((a, b) => {
      const na = (a.Nachname || a.ParticipantName || '').toLowerCase();
      const nb = (b.Nachname || b.ParticipantName || '').toLowerCase();
      return na.localeCompare(nb, 'de');
    }),
    [registrations],
  );
  const marksOf = React.useMemo(() => {
    const m = new Map<number, ReturnType<typeof parseAgendaCheckIns>>();
    active.forEach(r => m.set(r.Id, parseAgendaCheckIns(r.AgendaCheckIns)));
    return m;
  }, [active]);
  const countFor = (itemId: string): number => active.reduce((n, r) => n + ((marksOf.get(r.Id) || {})[itemId] ? 1 : 0), 0);
  const nameOf = (r: SPRegistration): string => `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || r.ParticipantName || r.ParticipantEmail || '—';

  const toggle = async (r: SPRegistration, it: AgendaItem): Promise<void> => {
    if (!canEdit || !eventServiceRef || !event.subsiteUrl) return;
    const key = `${r.Id}:${it.id}`;
    if (busyKey) return;
    const has = !!(marksOf.get(r.Id) || {})[it.id];
    const nm = nameOf(r);
    const ok = await confirmDialog(has
      ? (isDe ? `Anwesenheit von ${nm} bei „${it.title}“ entfernen?` : `Remove ${nm}'s attendance at “${it.title}”?`)
      : (isDe ? `${nm} bei „${it.title}“ als anwesend eintragen?` : `Record ${nm} as present at “${it.title}”?`),
      { confirmLabel: has ? (isDe ? 'Entfernen' : 'Remove') : (isDe ? 'Eintragen' : 'Record') });
    if (!ok) return;
    setBusyKey(key);
    try {
      const res = has
        ? await eventServiceRef.removeAgendaCheckIn(event.subsiteUrl, r.Id, it.id)
        : await eventServiceRef.checkInAgendaItem(event.subsiteUrl, r.Id, it.id);
      if (!res.ok) {
        showAlert(isDe
          ? `Konnte nicht gespeichert werden${res.status ? ` (HTTP ${res.status})` : ''}${res.status === 400 ? ' — fehlt die Spalte AgendaCheckIns? „Spalten fixen" ausführen.' : '.'}`
          : `Could not be saved${res.status ? ` (HTTP ${res.status})` : '.'}`, { variant: 'error' });
        return;
      }
      try {
        await eventServiceRef.writeChangeLog({
          action: has ? 'AgendaAttendanceRemoved' : 'AgendaAttendanceSet',
          targetType: 'Participant', targetId: r.ParticipantEmail || String(r.Id), targetName: nm,
          eventId: event.id, eventTitle: event.title,
          details: { agendaItemId: it.id, agendaItemTitle: it.title, manual: true },
        });
      } catch { /* Audit best-effort */ }
      await reloadRegistrations();
    } finally { setBusyKey(''); }
  };

  // v30.96: Teilnahmebescheinigungen (utils/attendanceCertificatePdf).
  const [pdfBusy, setPdfBusy] = React.useState(false);
  const certEvent = (): import('../../../utils/attendanceCertificatePdf').CertificateEvent => ({
    title: event.title, startDate: event.startDate, endDate: event.endDate, location: event.location, organizers: event.organizers,
    agenda: items, agendaTermSingular: event.agendaTermSingular, agendaTermPlural: event.agendaTermPlural,
  });
  const certPerson = (r: SPRegistration): import('../../../utils/attendanceCertificatePdf').CertificatePerson =>
    ({ name: nameOf(r), email: r.ParticipantEmail || '', marks: marksOf.get(r.Id) || {} });
  const downloadCertificate = async (r: SPRegistration): Promise<void> => {
    if (pdfBusy) return;
    setPdfBusy(true);
    try { await downloadAttendanceCertificate(certEvent(), certPerson(r), isDe); }
    catch (err) { console.warn('[DEX] Bescheinigung fehlgeschlagen:', err); showAlert(isDe ? 'Die Bescheinigung konnte nicht erzeugt werden.' : 'The certificate could not be created.', { variant: 'error' }); }
    finally { setPdfBusy(false); }
  };
  const downloadCertificates = async (): Promise<void> => {
    if (pdfBusy) return;
    const withMarks = active.filter(r => items.some(it => !!(marksOf.get(r.Id) || {})[it.id]));
    if (withMarks.length === 0) { showAlert(isDe ? 'Noch keine Anwesenheit erfasst — es gibt nichts zu bescheinigen.' : 'No attendance recorded yet — nothing to certify.'); return; }
    setPdfBusy(true);
    try { await downloadAttendanceCertificates(certEvent(), withMarks.map(certPerson), isDe); }
    catch (err) { console.warn('[DEX] Bescheinigungen fehlgeschlagen:', err); showAlert(isDe ? 'Die PDF-Datei konnte nicht erzeugt werden.' : 'The PDF could not be created.', { variant: 'error' }); }
    finally { setPdfBusy(false); }
  };

  const downloadXlsx = async (): Promise<void> => {
    if (xlsxBusy) return;
    setXlsxBusy(true);
    try {
      const XLSX = await import('xlsx');
      const head = ['Teilnehmer-ID', 'Name', 'E-Mail', ...items.map(it => `${it.date || ''} ${it.time || ''} ${it.title || ''}`.trim()), isDe ? 'Anwesend (Anzahl)' : 'Present (count)'];
      const rows: string[][] = [head];
      for (const r of active) {
        const m = marksOf.get(r.Id) || {};
        const cells = items.map(it => m[it.id] ? formatMarkTime(m[it.id].at) : '');
        rows.push([String(r.TeilnehmerID ?? ''), nameOf(r), r.ParticipantEmail || '', ...cells, String(cells.filter(Boolean).length)]);
      }
      const ws = XLSX.utils.aoa_to_sheet(rows);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = [{ wch: 12 }, { wch: 28 }, { wch: 34 }, ...items.map(() => ({ wch: 14 })), { wch: 12 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, isDe ? 'Nach Person' : 'By person');
      const rows2: string[][] = [[termS, isDe ? 'Datum' : 'Date', isDe ? 'Zeit' : 'Time', isDe ? 'Anwesend' : 'Present', isDe ? 'Angemeldet' : 'Registered', isDe ? 'Anteil' : 'Share']];
      for (const it of items) {
        const c = countFor(it.id);
        rows2.push([it.title || '', it.date || '', `${it.time || ''}${it.endTime ? '–' + it.endTime : ''}`, String(c), String(active.length), active.length ? `${Math.round(100 * c / active.length)} %` : '']);
      }
      const ws2 = XLSX.utils.aoa_to_sheet(rows2);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws2 as any)['!cols'] = [{ wch: 34 }, { wch: 12 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 8 }];
      XLSX.utils.book_append_sheet(wb, ws2, isDe ? 'Nach Punkt' : 'By item');
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = `Anwesenheit_${(event.title || 'Event').replace(/[^\w-]+/g, '_')}_${new Date().toISOString().slice(0, 10)}.xlsx`; a.style.display = 'none';
      document.body.appendChild(a); a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
    } catch (err) {
      console.warn('[DEX] Anwesenheits-Export fehlgeschlagen:', err);
      showAlert(isDe ? 'Die Excel-Datei konnte nicht erzeugt werden.' : 'The Excel file could not be created.', { variant: 'error' });
    } finally { setXlsxBusy(false); }
  };

  if (items.length === 0) return null;
  const totalMarks = active.reduce((n, r) => n + Object.keys(marksOf.get(r.Id) || {}).filter(k => items.some(it => it.id === k)).length, 0);

  return (
    <div className="card" style={{ padding: 20, marginBottom: 24 }}>
      <button type="button" onClick={() => setOpen(v => !v)} style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 10, background: 'none', border: 'none', padding: 0, cursor: 'pointer', textAlign: 'left' }}>
        <h3 style={{ margin: 0, fontSize: '1rem' }}>{isDe ? 'Anwesenheit je ' : 'Attendance per '}{termS}</h3>
        <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
          {regsUnknown
            ? (isDe ? '— Teilnehmerliste nicht lesbar' : '— attendee list not readable')
            : (isDe ? `— ${items.length} ${termP}, ${active.length} Angemeldete, ${totalMarks} Anwesenheiten` : `— ${items.length} ${termP}, ${active.length} registered, ${totalMarks} attendances`)}
        </span>
        <span style={{ marginLeft: 'auto', color: 'var(--dex-gray-400)' }}>{open ? '▾' : '▸'}</span>
      </button>
      {open && (
        <div style={{ marginTop: 14 }}>
          {regsUnknown ? (
            <div style={{ padding: '10px 12px', borderRadius: 8, background: 'rgba(218,41,28,0.08)', fontSize: '0.82rem' }}>
              {isDe ? 'Die Teilnehmerliste konnte nicht gelesen werden — es gibt hier keine Zahlen, keine Nullen.' : 'The attendee list could not be read — no numbers here, and no zeros.'}
            </div>
          ) : (
            <>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginBottom: 12 }}>
                <div role="radiogroup" style={{ display: 'inline-flex', gap: 3, padding: 3, border: '1px solid var(--dex-gray-300)', borderRadius: 999, background: '#fff' }}>
                  {([{ k: 'point' as const, de: `Nach ${termS}`, en: `By ${termS.toLowerCase()}` }, { k: 'person' as const, de: 'Nach Person', en: 'By person' }]).map(o => (
                    <button key={o.k} type="button" role="radio" aria-checked={view === o.k} onClick={() => setView(o.k)} style={{
                      border: 0, borderRadius: 999, padding: '6px 14px', cursor: 'pointer', fontSize: '0.82rem', fontWeight: 600,
                      background: view === o.k ? 'var(--dex-green-dark, #4a7c1f)' : 'transparent', color: view === o.k ? '#fff' : 'var(--dex-gray-700)',
                    }}>{isDe ? o.de : o.en}</button>
                  ))}
                </div>
                <span style={{ flex: 1 }} />
                {/* v30.96: Bescheinigungen für alle mit mindestens einer Erfassung — eine PDF, eine Seite je Person. */}
                <button type="button" className="btn btn-secondary" disabled={pdfBusy || totalMarks === 0} onClick={() => { void downloadCertificates(); }} style={{ fontSize: '0.8rem', padding: '6px 14px' }}
                  title={isDe ? 'Teilnahmebescheinigungen als PDF (eine Seite je Person mit erfasster Anwesenheit)' : 'Certificates of attendance as PDF (one page per person with recorded attendance)'}>
                  {pdfBusy ? (isDe ? 'Wird erzeugt…' : 'Creating…') : (isDe ? 'Bescheinigungen (PDF)' : 'Certificates (PDF)')}
                </button>
                <button type="button" className="btn btn-secondary" disabled={xlsxBusy} onClick={() => { void downloadXlsx(); }} style={{ fontSize: '0.8rem', padding: '6px 14px' }}>
                  {xlsxBusy ? (isDe ? 'Wird erzeugt…' : 'Creating…') : (isDe ? 'Als Excel laden' : 'Download Excel')}
                </button>
              </div>
              {canEdit && (
                <p style={{ margin: '0 0 10px', fontSize: '0.76rem', color: 'var(--dex-gray-500)' }}>
                  {isDe ? 'Klick auf eine Zelle bzw. einen Namen trägt die Anwesenheit nach oder nimmt sie zurück — mit Rückfrage und Eintrag im Event-Log.' : 'Click a cell or a name to record or remove attendance — with confirmation and an event-log entry.'}
                </p>
              )}

              {view === 'point' && totalMarks === 0 && (
                <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', padding: '8px 12px', borderRadius: 8, background: 'var(--dex-gray-50, #fafafa)', border: '1px dashed var(--dex-gray-300)' }}>
                  {isDe
                    ? `Noch keine Anwesenheit erfasst. Sobald am ersten ${termS} eingecheckt wird (Check-in-Seite oder Klick hier), erscheinen Anzahl und Anteil.`
                    : `No attendance recorded yet. Once someone checks in at the first ${termS.toLowerCase()} (check-in page or a click here), counts and shares appear.`}
                </p>
              )}
              {/* v30.94: kompakt und nach Cluster gegliedert. Vorher stand jeder
                  der 27 Punkte als eigene Vollbreite-Zeile mit leerem Balken —
                  „was zur Hölle ist das" (Nutzer, 07.09.2026). Jetzt: ein
                  Kasten je Cluster mit Tages-Kopf und Summe, darin schmale
                  Zeilen: Zeit · Titel · Raum · n / N · kurzer Balken. */}
              {/* v30.97: Cluster nebeneinander (Raster, auto-fit) statt
                  untereinander — vier Tage à sechs Punkte streckten die Seite
                  auf eine Bildschirmhöhe je Cluster (Nutzer-Screenshot 07.09.). */}
              {view === 'point' && (
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))', gap: 12, alignItems: 'start' }}>
                  {groups.map((g, gi) => {
                    const gMarks = g.items.reduce((n, it) => n + countFor(it.id), 0);
                    const gMax = g.items.length * active.length;
                    return (
                  <div key={g.key} style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, overflow: 'hidden' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '7px 12px', background: 'var(--dex-gray-50, #fafafa)', borderBottom: '1px solid var(--dex-gray-200)', fontSize: '0.8rem' }}>
                      <span style={{ fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)' }}>{groupLabel(g, gi, isDe)}</span>
                      <span style={{ color: 'var(--dex-gray-500)' }}>{groupDateLabel(g, isDe)}</span>
                      <span style={{ color: 'var(--dex-gray-400)' }}>· {g.items.length} {g.items.length === 1 ? termS : termP}</span>
                      <span style={{ flex: 1 }} />
                      <span style={{ color: 'var(--dex-gray-600)', fontVariantNumeric: 'tabular-nums' }} title={isDe ? 'Anwesenheiten in diesem Cluster / mögliche' : 'attendances in this cluster / possible'}>
                        <strong>{gMarks}</strong> / {gMax}{gMax ? ` · ${Math.round(100 * gMarks / gMax)} %` : ''}
                      </span>
                    </div>
                  {g.items.map(it => {
                    const c = countFor(it.id);
                    const pct = active.length ? Math.round(100 * c / active.length) : 0;
                    const isOpen = openPoint === it.id;
                    return (
                      <div key={it.id} style={{ borderTop: '1px solid var(--dex-gray-100)' }}>
                        <button type="button" onClick={() => setOpenPoint(isOpen ? null : it.id)} style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 10, padding: '6px 12px', border: 'none', cursor: 'pointer', textAlign: 'left', background: isOpen ? 'var(--dex-gray-100)' : '#fff' }}>
                          <span style={{ width: 92, flexShrink: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontVariantNumeric: 'tabular-nums' }}>{it.time}{it.endTime ? `–${it.endTime}` : ''}</span>
                          <span style={{ fontWeight: 600, fontSize: '0.86rem', flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{it.title || (isDe ? '(ohne Titel)' : '(untitled)')}{it.location ? <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}> · {it.location}</span> : null}</span>
                          <strong style={{ width: 64, flexShrink: 0, textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: c ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)' }}>{c}<span style={{ fontWeight: 400, color: 'var(--dex-gray-400)' }}> / {active.length}</span></strong>
                          <span style={{ width: 110, flexShrink: 0, height: 6, borderRadius: 999, background: 'var(--dex-gray-100)' }}>
                            <span style={{ display: 'block', height: 6, width: `${pct}%`, borderRadius: 999, background: 'var(--dex-green, #86bc25)' }} />
                          </span>
                          <span style={{ width: 12, color: 'var(--dex-gray-400)', fontSize: '0.8rem' }}>{isOpen ? '▾' : '▸'}</span>
                        </button>
                        {isOpen && (
                          <div style={{ padding: '8px 14px 12px', borderTop: '1px solid var(--dex-gray-100)', display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 12, fontSize: '0.82rem' }}>
                            {[true, false].map(present => {
                              const list = active.filter(r => !!(marksOf.get(r.Id) || {})[it.id] === present);
                              return (
                                <div key={String(present)}>
                                  <div style={{ fontWeight: 700, marginBottom: 4, color: present ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)' }}>
                                    {present ? (isDe ? `Anwesend (${list.length})` : `Present (${list.length})`) : (isDe ? `Nicht erfasst (${list.length})` : `Not recorded (${list.length})`)}
                                  </div>
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: 3, maxHeight: 260, overflow: 'auto' }}>
                                    {list.map(r => {
                                      const m = (marksOf.get(r.Id) || {})[it.id];
                                      return (
                                        <button key={r.Id} type="button" disabled={!canEdit || !!busyKey} onClick={() => { void toggle(r, it); }} title={canEdit ? (present ? (isDe ? 'Anwesenheit entfernen' : 'Remove attendance') : (isDe ? 'Als anwesend eintragen' : 'Record as present')) : undefined} style={{
                                          display: 'flex', alignItems: 'center', gap: 8, padding: '3px 6px', border: 'none', background: 'transparent', cursor: canEdit ? 'pointer' : 'default', textAlign: 'left', borderRadius: 6, color: 'inherit', fontSize: 'inherit',
                                        }}>
                                          <PersonContactHover email={r.ParticipantEmail || ''} name={nameOf(r)} size={22} isDe={isDe} />
                                          <span style={{ flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{nameOf(r)}</span>
                                          {m && <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontVariantNumeric: 'tabular-nums' }}>{formatMarkTime(m.at)}</span>}
                                        </button>
                                      );
                                    })}
                                    {list.length === 0 && <span style={{ color: 'var(--dex-gray-400)' }}>—</span>}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                  </div>
                    );
                  })}
                </div>
              )}

              {view === 'person' && (
                <div style={{ maxHeight: '70vh', overflow: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 8 }}>
                  <table style={{ borderCollapse: 'collapse', fontSize: '0.8rem', minWidth: '100%' }}>
                    <thead>
                      <tr>
                        <th style={{ position: 'sticky', top: 0, left: 0, zIndex: 6, background: '#fff', textAlign: 'left', padding: 8, borderBottom: '2px solid var(--dex-gray-200)', minWidth: 220 }}>{isDe ? 'Teilnehmer' : 'Attendee'}</th>
                        {items.map((it, i) => (
                          <th key={it.id} title={`${it.date || ''} ${it.time || ''} ${it.title || ''}`} style={{ position: 'sticky', top: 0, zIndex: 5, background: '#fff', padding: '6px 8px', borderBottom: '2px solid var(--dex-gray-200)', textAlign: 'center', minWidth: 64, fontWeight: 600, whiteSpace: 'nowrap' }}>
                            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{i + 1}</div>
                            <div style={{ fontSize: '0.72rem', maxWidth: 110, overflow: 'hidden', textOverflow: 'ellipsis' }}>{it.title || '—'}</div>
                            <div style={{ fontSize: '0.68rem', color: 'var(--dex-gray-400)' }}>{countFor(it.id)}</div>
                          </th>
                        ))}
                        <th style={{ position: 'sticky', top: 0, zIndex: 5, background: '#fff', padding: 8, borderBottom: '2px solid var(--dex-gray-200)', textAlign: 'right' }}>Σ</th>
                        <th style={{ position: 'sticky', top: 0, zIndex: 5, background: '#fff', padding: 8, borderBottom: '2px solid var(--dex-gray-200)', textAlign: 'center', fontSize: '0.72rem', fontWeight: 600 }}>{isDe ? 'Bescheinigung' : 'Certificate'}</th>
                      </tr>
                    </thead>
                    <tbody>
                      {active.map(r => {
                        const m = marksOf.get(r.Id) || {};
                        const sum = items.filter(it => !!m[it.id]).length;
                        return (
                          <tr key={r.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                            <td style={{ position: 'sticky', left: 0, background: '#fff', padding: 6, zIndex: 1 }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
                                <PersonContactHover email={r.ParticipantEmail || ''} name={nameOf(r)} size={26} isDe={isDe} />
                                <span style={{ fontWeight: 600, whiteSpace: 'nowrap' }}>{nameOf(r)}</span>
                                {r.TeilnehmerID ? <span style={{ color: 'var(--dex-gray-400)', fontSize: '0.72rem' }}>#{r.TeilnehmerID}</span> : null}
                              </div>
                            </td>
                            {items.map(it => {
                              const mk = m[it.id];
                              const key = `${r.Id}:${it.id}`;
                              return (
                                <td key={it.id} style={{ padding: 4, textAlign: 'center' }}>
                                  <button type="button" disabled={!canEdit || !!busyKey} onClick={() => { void toggle(r, it); }} title={mk ? `${isDe ? 'anwesend' : 'present'} ${formatMarkTime(mk.at)}${mk.by ? ` · ${mk.by}` : ''}` : (canEdit ? (isDe ? 'Als anwesend eintragen' : 'Record as present') : '')} style={{
                                    width: 56, padding: '4px 0', borderRadius: 6, cursor: canEdit ? 'pointer' : 'default', fontSize: '0.72rem', fontWeight: 700,
                                    border: `1px solid ${mk ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                                    background: busyKey === key ? 'var(--dex-gray-100)' : mk ? 'rgba(134,188,37,0.15)' : '#fff',
                                    color: mk ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-300)',
                                  }}>{mk ? `✓ ${formatMarkTime(mk.at)}` : '·'}</button>
                                </td>
                              );
                            })}
                            <td style={{ padding: 6, textAlign: 'right', fontWeight: 700, color: sum === items.length ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-700)' }}>{sum}/{items.length}</td>
                            <td style={{ padding: 4, textAlign: 'center' }}>
                              {sum > 0 && (
                                <button type="button" disabled={pdfBusy} onClick={() => { void downloadCertificate(r); }} title={isDe ? 'Teilnahmebescheinigung (PDF)' : 'Certificate of attendance (PDF)'}
                                  style={{ border: '1px solid var(--dex-gray-200)', background: '#fff', borderRadius: 6, padding: '3px 8px', cursor: 'pointer', fontSize: '0.72rem', color: 'var(--dex-gray-700)' }}>PDF</button>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                      {active.length === 0 && (
                        <tr><td colSpan={items.length + 3} style={{ padding: 12, color: 'var(--dex-gray-400)' }}>{isDe ? 'Keine aktiven Anmeldungen.' : 'No active registrations.'}</td></tr>
                      )}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          )}
        </div>
      )}
    </div>
  );
};

export default AgendaAttendanceSection;
