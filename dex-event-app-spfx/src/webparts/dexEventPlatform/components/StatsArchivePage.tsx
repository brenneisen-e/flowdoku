/**
 * v26.33: Statistik-Archiv (DEX_EventStats).
 *
 * Reine Anzeige-Seite für die KPIs von Events, deren Teilnehmerliste nach dem
 * 3-Monats-Löschkonzept gelöscht wurde. Enthält KEINE personenbezogenen Daten
 * mehr — nur noch die Kennzahlen (wie viele angemeldet, abgemeldet, Warteliste,
 * QR-Code, eingecheckt, No-Show), plus welches Event wann von wem organisiert
 * wurde und wer die Löschung ausgelöst hat.
 *
 * Erreichbar über die Kachel „Statistik-Archiv" im Admin Center (nur Admins).
 */
import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents, EventStatsRow } from '../context/EventContext';
import { useLanguage } from '../context/LanguageContext';
import { BarChart3, Download, RefreshCw, Search, ChevronLeft } from './Icons';

function fmtDate(iso: string, withTime = false): string {
  if (!iso) return '–';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '–';
  return d.toLocaleDateString('de-DE', withTime
    ? { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }
    : { day: '2-digit', month: '2-digit', year: 'numeric' });
}

export default function StatsArchivePage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { isAdmin, originalIsAdmin } = useRoles();
  const { getEventStats } = useEvents();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const adminLike = isAdmin || originalIsAdmin;

  const [rows, setRows] = React.useState<EventStatsRow[] | null>(null);
  const [loading, setLoading] = React.useState(true);
  const [search, setSearch] = React.useState('');

  const load = React.useCallback(async () => {
    setLoading(true);
    try { setRows(await getEventStats()); }
    catch { setRows([]); }
    finally { setLoading(false); }
  }, [getEventStats]);

  React.useEffect(() => { void load(); }, [load]);

  const filtered = React.useMemo(() => {
    const list = rows || [];
    const q = search.trim().toLowerCase();
    if (!q) return list;
    return list.filter(r =>
      (r.eventTitle || '').toLowerCase().indexOf(q) >= 0 ||
      (r.organizer || '').toLowerCase().indexOf(q) >= 0 ||
      (r.location || '').toLowerCase().indexOf(q) >= 0 ||
      String(r.eventNumber).indexOf(q) >= 0
    );
  }, [rows, search]);

  const exportCsv = (): void => {
    const list = filtered;
    if (list.length === 0) return;
    const headers = [
      'EventNr', isDe ? 'Event' : 'Event', isDe ? 'Typ' : 'Type', isDe ? 'Ort' : 'Location',
      isDe ? 'Start' : 'Start', isDe ? 'Ende' : 'End', isDe ? 'Organizer' : 'Organizer',
      isDe ? 'Max' : 'Max', isDe ? 'Angemeldet' : 'Registered', 'QR', isDe ? 'Eingecheckt' : 'Checked in',
      'No-Show', isDe ? 'Warteliste' : 'Waitlist', isDe ? 'Abgemeldet' : 'Deregistered',
      isDe ? 'Archiviert am' : 'Archived on', isDe ? 'Archiviert von' : 'Archived by',
    ];
    const esc = (v: string | number | null): string => {
      const s = v === null || v === undefined ? '' : String(v);
      return /[";\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const lines = [headers.join(';')];
    for (const r of list) {
      lines.push([
        r.eventNumber, r.eventTitle, r.eventType, r.location,
        fmtDate(r.startDate), fmtDate(r.endDate), r.organizer,
        r.maxParticipants === null ? '' : r.maxParticipants, r.registeredCount, r.qrSentCount, r.checkedInCount,
        r.noShowCount, r.waitlistCount, r.deregisteredCount,
        fmtDate(r.archivedDate, true), r.archivedByEmail,
      ].map(esc).join(';'));
    }
    const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `DEX_Statistik-Archiv.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  if (!adminLike) {
    return (
      <div className="page-container">
        <h1 style={{ marginTop: 0 }}>{isDe ? 'Statistik-Archiv' : 'Statistics archive'}</h1>
        <p style={{ color: 'var(--dex-gray-600)' }}>
          {isDe ? 'Diese Ansicht ist nur für Admins.' : 'This view is for admins only.'}
        </p>
        <button className="btn btn-secondary" onClick={() => navigate('admin-hub')}>
          {isDe ? 'Zurück zum Admin Center' : 'Back to admin center'}
        </button>
      </div>
    );
  }

  const th: React.CSSProperties = { padding: '10px 12px', textAlign: 'left', fontWeight: 600, fontSize: '0.78rem', color: 'var(--dex-gray-700)', whiteSpace: 'nowrap', borderBottom: '2px solid var(--dex-gray-200)' };
  const thNum: React.CSSProperties = { ...th, textAlign: 'right' };
  const td: React.CSSProperties = { padding: '9px 12px', fontSize: '0.82rem', color: 'var(--dex-gray-800)', verticalAlign: 'top', borderBottom: '1px solid var(--dex-gray-100)' };
  const tdNum: React.CSSProperties = { ...td, textAlign: 'right', fontVariantNumeric: 'tabular-nums' };

  return (
    <div className="page-container">
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 4 }}>
        <span style={{ color: 'var(--dex-green, #86bc25)', display: 'inline-flex' }}><BarChart3 size={30} /></span>
        <h1 style={{ margin: 0 }}>{isDe ? 'Statistik-Archiv' : 'Statistics archive'}</h1>
      </div>
      <p style={{ color: 'var(--dex-gray-600)', marginTop: 0, maxWidth: 780 }}>
        {isDe
          ? 'Kennzahlen von Events, deren Teilnehmerliste nach dem 3-Monats-Löschkonzept gelöscht wurde. Diese Ansicht enthält keine personenbezogenen Daten mehr — nur noch die Kennzahlen: welches Event, wann, von wem organisiert und mit welcher Teilnehmerzahl.'
          : 'KPIs of events whose participant list was deleted per the 3-month retention concept. This view no longer contains any personal data — only the key figures: which event, when, organized by whom, and how many participants.'}
      </p>

      <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap', margin: '16px 0' }}>
        <div style={{ position: 'relative', flex: '1 1 260px', maxWidth: 360 }}>
          <span style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--dex-gray-400)', display: 'inline-flex' }}><Search size={16} /></span>
          <input
            type="text"
            value={search}
            onChange={e => setSearch(e.target.value)}
            placeholder={isDe ? 'Event, Organizer, Ort suchen…' : 'Search event, organizer, location…'}
            style={{ width: '100%', padding: '8px 10px 8px 32px', borderRadius: 8, border: '1px solid var(--dex-gray-300)', fontSize: '0.85rem' }}
          />
        </div>
        <button className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '8px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => { void load(); }} disabled={loading}>
          <RefreshCw size={15} /> {isDe ? 'Aktualisieren' : 'Refresh'}
        </button>
        <button className="btn btn-secondary" style={{ fontSize: '0.82rem', padding: '8px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={exportCsv} disabled={loading || filtered.length === 0}>
          <Download size={15} /> {isDe ? 'CSV exportieren' : 'Export CSV'}
        </button>
      </div>

      {loading ? (
        <p style={{ color: 'var(--dex-gray-500)' }}>{isDe ? 'Wird geladen…' : 'Loading…'}</p>
      ) : filtered.length === 0 ? (
        <div style={{ background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 28, textAlign: 'center', color: 'var(--dex-gray-500)' }}>
          {(rows || []).length === 0
            ? (isDe ? 'Noch keine archivierten Statistiken vorhanden. Sobald die erste Teilnehmerliste nach dem 3-Monats-Löschkonzept gelöscht wurde, erscheinen die Kennzahlen hier.' : 'No archived statistics yet. Once the first participant list is deleted per the 3-month retention concept, the KPIs will appear here.')
            : (isDe ? 'Keine Treffer für die Suche.' : 'No matches for your search.')}
        </div>
      ) : (
        <>
          <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', margin: '0 0 8px' }}>
            {isDe ? `${filtered.length} Event${filtered.length === 1 ? '' : 's'} im Archiv` : `${filtered.length} event${filtered.length === 1 ? '' : 's'} in the archive`}
          </p>
          <div style={{ overflowX: 'auto', background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 12 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 980 }}>
              <thead>
                <tr>
                  <th style={th}>{isDe ? 'Event' : 'Event'}</th>
                  <th style={th}>{isDe ? 'Wann' : 'When'}</th>
                  <th style={th}>{isDe ? 'Ort' : 'Location'}</th>
                  <th style={th}>{isDe ? 'Organizer' : 'Organizer'}</th>
                  <th style={thNum}>{isDe ? 'Max' : 'Max'}</th>
                  <th style={thNum}>{isDe ? 'Angem.' : 'Reg.'}</th>
                  <th style={thNum}>QR</th>
                  <th style={thNum}>{isDe ? 'Eingech.' : 'Check-in'}</th>
                  <th style={thNum}>No-Show</th>
                  <th style={thNum}>{isDe ? 'Wartel.' : 'Waitl.'}</th>
                  <th style={thNum}>{isDe ? 'Abgem.' : 'Dereg.'}</th>
                  <th style={th}>{isDe ? 'Gelöscht' : 'Deleted'}</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map(r => (
                  <tr key={r.id}>
                    <td style={td}>
                      <span style={{ fontWeight: 600, display: 'block' }}>{r.eventTitle || '–'}</span>
                      <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
                        #{r.eventNumber}{r.eventType ? ` · ${r.eventType}` : ''}
                      </span>
                    </td>
                    <td style={td}>
                      {fmtDate(r.startDate)}
                      {r.endDate && fmtDate(r.endDate) !== fmtDate(r.startDate) ? <><br /><span style={{ color: 'var(--dex-gray-500)' }}>– {fmtDate(r.endDate)}</span></> : null}
                    </td>
                    <td style={td}>{r.location || '–'}</td>
                    <td style={{ ...td, maxWidth: 220 }}>{r.organizer || '–'}</td>
                    <td style={tdNum}>{r.maxParticipants === null ? '–' : r.maxParticipants}</td>
                    <td style={tdNum}>{r.registeredCount}</td>
                    <td style={tdNum}>{r.qrSentCount}</td>
                    <td style={tdNum}>{r.checkedInCount}</td>
                    <td style={tdNum}>{r.noShowCount}</td>
                    <td style={tdNum}>{r.waitlistCount}</td>
                    <td style={tdNum}>{r.deregisteredCount}</td>
                    <td style={td}>
                      <span style={{ display: 'block' }}>{fmtDate(r.archivedDate, true)}</span>
                      {r.archivedByEmail ? <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{r.archivedByEmail}</span> : null}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      <div style={{ marginTop: 20 }}>
        <button className="btn btn-secondary" style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => navigate('admin-hub')}>
          <ChevronLeft size={18} /> {isDe ? 'Zurück zum Admin Center' : 'Back to admin center'}
        </button>
      </div>
    </div>
  );
}
