/**
 * v30.5: F&A Center (Fachkonzept Abschnitte 8–11).
 *
 * Zentrale Sicht der Finance & Accounting Abteilung auf alle
 * abrechnungsrelevanten Events: Verteiler-Pflege, Dashboard-Kennzahlen,
 * durchsuchbare Tabelle und Detailansicht mit Kommunikationshistorie und
 * „Als abgerechnet markieren". Zugriff: Rolle „F&A" und Admins. Texte
 * bewusst nur deutsch — das Konzept ist deutsch, der Pilot ist intern.
 *
 * Datenhaltung: alles hängt am Piggyback `_billing` der Events (siehe
 * utils/faBilling.ts) plus der Verteiler-Zeile `_FAConfig` in
 * DEX_EmailTemplates. Die Detailansicht zeigt laut Konzept NUR den an F&A
 * ÜBERMITTELTEN Stand (Snapshots), nie den Live-Stand des Wizards.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useRoles } from '../context/RoleContext';
import { useCurrentUser } from '../context/UserContext';
import { useDialog } from '../context/DialogContext';
import { DeloitteEvent } from '../types';
import { BILLING_FIELDS } from '../data/billingFields';
import {
  parseBillingOf, faStatusOf, FAStatus, FA_STATUS_LABELS, FA_STATUS_COLORS,
  FAConfig, BillingLogEntry,
} from '../utils/faBilling';

const fmtDate = (iso: string): string => {
  const d = new Date(iso || '');
  return isFinite(d.getTime()) ? d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '—';
};
const fmtDateTime = (iso: string): string => {
  const d = new Date(iso || '');
  return isFinite(d.getTime())
    ? d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ', ' + d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
    : '—';
};

/** E-Mail-Liste aus einem Freitext (Komma, Semikolon oder Zeilenumbruch). */
const parseRecipientInput = (raw: string): string[] => {
  const seen = new Set<string>();
  return raw.split(/[,;\n]/).map(s => s.trim()).filter(Boolean).filter(s => {
    const lc = s.toLowerCase();
    if (seen.has(lc)) return false;
    seen.add(lc);
    return true;
  });
};

const statusPill = (s: FAStatus): React.ReactElement => (
  <span style={{
    display: 'inline-block', padding: '3px 10px', borderRadius: 999,
    fontSize: '0.72rem', fontWeight: 700,
    background: FA_STATUS_COLORS[s].bg, color: FA_STATUS_COLORS[s].fg,
  }}>{FA_STATUS_LABELS[s]}</span>
);

export default function FACenterPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { isFA, isAdmin } = useRoles();
  const { events, getFAConfig, saveFAConfig, markEventSettled } = useEvents();
  const { confirmDialog, showAlert } = useDialog();
  const { currentUser } = useCurrentUser();

  const allowed = isFA || isAdmin;

  const [cfg, setCfg] = React.useState<FAConfig | null>(null);
  const [infoInput, setInfoInput] = React.useState('');
  const [listInput, setListInput] = React.useState('');
  const [cfgBusy, setCfgBusy] = React.useState(false);
  const [cfgLogOpen, setCfgLogOpen] = React.useState(false);
  const [search, setSearch] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState<'' | FAStatus>('');
  const [selectedId, setSelectedId] = React.useState<string | null>(null);
  const [settleBusy, setSettleBusy] = React.useState(false);
  const [openMailIdx, setOpenMailIdx] = React.useState<number | null>(null);

  React.useEffect(() => {
    if (!allowed) return;
    let cancelled = false;
    getFAConfig().then(c => {
      if (cancelled) return;
      setCfg(c);
      setInfoInput(c.infoRecipients.join('\n'));
      setListInput(c.listRecipients.join('\n'));
    }).catch(() => { /* leerer Stand bleibt */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allowed]);

  // Nur Hauptevents tragen eine Abrechnung; Sub-Events und Demo bleiben draußen.
  const billingEvents = React.useMemo(() => {
    return events
      .filter(e => !e.parentEventId && !e.isDemoShowcase)
      .map(e => ({ ev: e, b: parseBillingOf(e) }))
      .filter(x => x.b && x.b.relevant)
      .map(x => ({ ...x, status: faStatusOf(x.ev, x.b) }))
      .sort((a, b2) => new Date(b2.ev.startDate || 0).getTime() - new Date(a.ev.startDate || 0).getTime());
  }, [events]);

  const counts = React.useMemo(() => {
    const c: Record<FAStatus, number> = { incomplete: 0, upcoming: 0, listPending: 0, sentAwaitSettle: 0, settled: 0 };
    billingEvents.forEach(x => { c[x.status]++; });
    return c;
  }, [billingEvents]);

  const filtered = React.useMemo(() => {
    const q = search.trim().toLowerCase();
    return billingEvents.filter(x => {
      if (statusFilter && x.status !== statusFilter) return false;
      if (!q) return true;
      const hay = [
        String(x.ev.eventNumber || ''), x.ev.title,
        (x.b?.fields || {}).contact || '', (x.b?.fields || {}).ariba || '',
        fmtDate(x.ev.startDate), FA_STATUS_LABELS[x.status],
      ].join(' ').toLowerCase();
      return hay.indexOf(q) >= 0;
    });
  }, [billingEvents, search, statusFilter]);

  const selected = selectedId ? billingEvents.find(x => x.ev.id === selectedId) : undefined;

  if (!allowed) {
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 560, margin: '0 auto' }}>
          <h2 style={{ marginTop: 0 }}>Kein Zugriff</h2>
          <p style={{ color: 'var(--dex-gray-600)' }}>
            Das F&amp;A Center steht Personen mit der Rolle &bdquo;F&amp;A&ldquo; und Admins zur Verfügung.
          </p>
          <button className="btn btn-primary mt-24" onClick={() => navigate('start')}>Zurück</button>
        </div>
      </div>
    );
  }

  const saveRecipients = async (): Promise<void> => {
    if (!cfg || cfgBusy) return;
    setCfgBusy(true);
    try {
      const nextInfo = parseRecipientInput(infoInput);
      const nextList = parseRecipientInput(listInput);
      const by = `${currentUser?.firstName || ''} ${currentUser?.surname || ''}`.trim() || currentUser?.email || '';
      const log = [...cfg.log];
      // Änderungen protokollieren (Konzept 8.1: „Änderungen werden protokolliert").
      if (nextInfo.join(';') !== cfg.infoRecipients.join(';')) {
        log.push({ ts: new Date().toISOString(), by, action: 'Verteiler „Abrechnungsinformationen" geändert', old: cfg.infoRecipients.join('; '), neu: nextInfo.join('; ') });
      }
      if (nextList.join(';') !== cfg.listRecipients.join(';')) {
        log.push({ ts: new Date().toISOString(), by, action: 'Verteiler „Teilnehmerlisten" geändert', old: cfg.listRecipients.join('; '), neu: nextList.join('; ') });
      }
      const next: FAConfig = { infoRecipients: nextInfo, listRecipients: nextList, log };
      const ok = await saveFAConfig(next);
      if (ok) {
        setCfg(next);
        showAlert('Verteiler gespeichert — sie gelten ab sofort für alle zukünftigen Versendungen.', { variant: 'success' });
      } else {
        showAlert('Speichern fehlgeschlagen — bitte später erneut versuchen.', { variant: 'error' });
      }
    } finally { setCfgBusy(false); }
  };

  const doSettle = async (ev: DeloitteEvent): Promise<void> => {
    if (settleBusy) return;
    const ok = await confirmDialog(
      `„${ev.title}" als abgerechnet markieren? Der Status bleibt dauerhaft bestehen; Datum und Person werden protokolliert.`,
      { confirmLabel: 'Als abgerechnet markieren' },
    );
    if (!ok) return;
    setSettleBusy(true);
    try {
      const done = await markEventSettled(ev);
      if (!done) showAlert('Markieren fehlgeschlagen — bitte später erneut versuchen.', { variant: 'error' });
    } finally { setSettleBusy(false); }
  };

  // ---------- Detailansicht (Konzept Abschnitt 10) ----------
  if (selected) {
    const { ev, b, status } = selected;
    const mails = (b?.log || []).filter(l => l.mailType);
    return (
      <div className="page-container">
        <button className="btn btn-secondary" style={{ marginBottom: 16 }} onClick={() => { setSelectedId(null); setOpenMailIdx(null); }}>
          ← Zurück zur Übersicht
        </button>
        <div className="card" style={{ padding: '24px 28px', marginBottom: 16 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 16, flexWrap: 'wrap' }}>
            <div>
              <h2 style={{ margin: '0 0 6px' }}>{ev.title}</h2>
              <div style={{ marginBottom: 10 }}>{statusPill(status)}</div>
            </div>
            {status !== 'settled' && (
              <button className="btn btn-primary" disabled={settleBusy} onClick={() => { void doSettle(ev); }}>
                {settleBusy ? 'Wird markiert…' : 'Als abgerechnet markieren'}
              </button>
            )}
          </div>
          {b?.settled && (
            <p style={{ margin: '0 0 12px', fontSize: '0.85rem', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>
              Abgerechnet am {fmtDateTime(b.settled.ts)} durch {b.settled.by}.
            </p>
          )}
          {/* Bereich 1: Eventinformationen */}
          <h3 style={{ margin: '10px 0 8px', fontSize: '0.95rem' }}>Eventinformationen</h3>
          <table style={{ borderCollapse: 'collapse', fontSize: '0.85rem' }}>
            <tbody>
              {[
                ['Event-ID', String(ev.eventNumber || ev.id)],
                ['Eventdatum', `${fmtDateTime(ev.startDate)}${ev.endDate && ev.endDate !== ev.startDate ? ' – ' + fmtDateTime(ev.endDate) : ''}`],
                ['Ort', ev.location || '—'],
                ['Organizer', (ev.organizers || []).join('; ') || '—'],
                ['Versandart', (b?.sendMode === 'auto') ? 'Automatisiert (7 Tage vor/nach dem Event)' : 'Manuell über die Organizer-Ansicht'],
              ].map(([k, v]) => (
                <tr key={k}>
                  <td style={{ padding: '4px 16px 4px 0', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap', verticalAlign: 'top' }}>{k}</td>
                  <td style={{ padding: '4px 0' }}>{v}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Bereich 2: An F&A übermittelte Abrechnungsinformationen */}
        <div className="card" style={{ padding: '20px 28px', marginBottom: 16 }}>
          <h3 style={{ margin: '0 0 10px', fontSize: '0.95rem' }}>Abrechnungsrelevante Informationen</h3>
          {b?.infoSnapshot ? (
            <>
              <p style={{ margin: '0 0 10px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                Stand des letzten Versands an F&amp;A ({b.infoSentAt ? fmtDateTime(b.infoSentAt) : '—'}).
              </p>
              <table style={{ borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                <tbody>
                  {BILLING_FIELDS.map(f => (
                    <tr key={f.id}>
                      <td style={{ padding: '4px 16px 4px 0', color: 'var(--dex-gray-500)', verticalAlign: 'top' }}>{f.label}</td>
                      <td style={{ padding: '4px 0' }}>{(b.infoSnapshot || {})[f.id] || '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </>
          ) : (
            <p style={{ margin: 0, color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
              Es wurden bisher keine Abrechnungsinformationen an F&amp;A versendet.
            </p>
          )}
        </div>

        {/* Bereich 3: An F&A übermittelte Teilnehmerliste */}
        <div className="card" style={{ padding: '20px 28px', marginBottom: 16 }}>
          <h3 style={{ margin: '0 0 10px', fontSize: '0.95rem' }}>Teilnehmerliste</h3>
          {b?.listSnapshot && b.listSnapshot.length > 0 ? (
            <>
              <p style={{ margin: '0 0 10px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                {b.listSnapshot.length} Personen — Stand des letzten Versands an F&amp;A ({b.listSentAt ? fmtDateTime(b.listSentAt) : '—'}).
              </p>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ borderCollapse: 'collapse', fontSize: '0.85rem', minWidth: 420 }}>
                  <thead>
                    <tr>
                      {['#', 'Name', 'E-Mail', 'Status'].map(h => (
                        <th key={h} style={{ textAlign: 'left', padding: '6px 12px 6px 0', borderBottom: '2px solid var(--dex-gray-200)' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {b.listSnapshot.map((p, i) => (
                      <tr key={i}>
                        <td style={{ padding: '4px 12px 4px 0', color: 'var(--dex-gray-400)' }}>{i + 1}</td>
                        <td style={{ padding: '4px 12px 4px 0' }}>{p.name}</td>
                        <td style={{ padding: '4px 12px 4px 0' }}>{p.email}</td>
                        <td style={{ padding: '4px 0' }}>{p.status}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          ) : (
            <p style={{ margin: 0, color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
              Es wurde bisher keine Teilnehmerliste an F&amp;A versendet.
            </p>
          )}
        </div>

        {/* Bereich 4: Kommunikationshistorie */}
        <div className="card" style={{ padding: '20px 28px', marginBottom: 16 }}>
          <h3 style={{ margin: '0 0 10px', fontSize: '0.95rem' }}>Kommunikationshistorie</h3>
          {mails.length === 0 ? (
            <p style={{ margin: 0, color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
              Über die Plattform wurden noch keine abrechnungsrelevanten E-Mails versendet.
            </p>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {mails.map((m: BillingLogEntry, i: number) => (
                <div key={i} style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: '10px 12px', fontSize: '0.82rem' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap' }}>
                    <strong>{m.mailType === 'info' ? 'Abrechnungsinformationen' : 'Teilnehmerliste'}</strong>
                    <span style={{ color: 'var(--dex-gray-500)' }}>{fmtDateTime(m.ts)}</span>
                  </div>
                  <div style={{ color: 'var(--dex-gray-600)', marginTop: 4, lineHeight: 1.5 }}>
                    Absender: {m.by} · Empfänger: {m.to || '—'}{m.cc ? <> · CC: {m.cc}</> : null}<br />
                    Betreff: {m.subject || '—'}
                  </div>
                  {m.body ? (
                    <button
                      type="button"
                      className="btn btn-secondary"
                      style={{ marginTop: 8, fontSize: '0.75rem', padding: '4px 12px' }}
                      onClick={() => setOpenMailIdx(openMailIdx === i ? null : i)}
                    >
                      {openMailIdx === i ? 'Inhalt ausblenden' : 'Vollständigen Inhalt anzeigen'}
                    </button>
                  ) : (
                    <div style={{ marginTop: 6, color: 'var(--dex-gray-400)', fontSize: '0.75rem' }}>
                      Inhalt aus Platzgründen nicht mehr gespeichert (nur die letzten 15 Mails behalten ihren Volltext).
                    </div>
                  )}
                  {openMailIdx === i && m.body && (
                    <div
                      style={{ marginTop: 10, border: '1px dashed var(--dex-gray-300)', borderRadius: 8, padding: 12, background: '#fff', overflowX: 'auto' }}
                      dangerouslySetInnerHTML={{ __html: m.body }}
                    />
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    );
  }

  // ---------- Übersicht ----------
  return (
    <div className="page-container">
      <h2 style={{ marginBottom: 4 }}>F&amp;A Center</h2>
      <p style={{ color: 'var(--dex-gray-600)', marginTop: 0, marginBottom: 20, fontSize: '0.9rem' }}>
        Zentrale Verwaltung und Nachverfolgung aller abrechnungsrelevanten Veranstaltungen.
      </p>

      {/* 8.1 Verteiler */}
      <div className="card" style={{ padding: '20px 24px', marginBottom: 20 }}>
        <h3 style={{ margin: '0 0 6px', fontSize: '0.95rem' }}>Empfängeradressen</h3>
        <p style={{ margin: '0 0 12px', fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          An diese Adressen gehen die Versendungen — getrennt für Abrechnungsinformationen und Teilnehmerlisten
          (eine Adresse pro Zeile). Änderungen wirken sofort auf alle zukünftigen Versendungen und werden protokolliert.
        </p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 16 }}>
          <div>
            <label className="form-label" style={{ fontSize: '0.8rem' }}>Verteiler Abrechnungsinformationen</label>
            <textarea className="form-input" rows={3} value={infoInput} onChange={e => setInfoInput(e.target.value)} placeholder={'fa-abrechnung@deloitte.de'} style={{ resize: 'vertical' }} />
          </div>
          <div>
            <label className="form-label" style={{ fontSize: '0.8rem' }}>Verteiler Teilnehmerlisten</label>
            <textarea className="form-input" rows={3} value={listInput} onChange={e => setListInput(e.target.value)} placeholder={'fa-teilnehmerlisten@deloitte.de'} style={{ resize: 'vertical' }} />
          </div>
        </div>
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginTop: 12, flexWrap: 'wrap' }}>
          <button className="btn btn-primary" disabled={cfgBusy || !cfg} onClick={() => { void saveRecipients(); }}>
            {cfgBusy ? 'Wird gespeichert…' : 'Verteiler speichern'}
          </button>
          {cfg && cfg.log.length > 0 && (
            <button className="btn btn-secondary" style={{ fontSize: '0.8rem' }} onClick={() => setCfgLogOpen(!cfgLogOpen)}>
              {cfgLogOpen ? 'Protokoll ausblenden' : `Änderungsprotokoll (${cfg.log.length})`}
            </button>
          )}
        </div>
        {cfgLogOpen && cfg && (
          <div style={{ marginTop: 12, fontSize: '0.78rem', color: 'var(--dex-gray-600)', display: 'flex', flexDirection: 'column', gap: 6 }}>
            {[...cfg.log].reverse().map((l, i) => (
              <div key={i} style={{ borderLeft: '3px solid var(--dex-gray-200)', paddingLeft: 10 }}>
                <strong>{fmtDateTime(l.ts)}</strong> · {l.by} · {l.action}
                {l.old !== undefined && <><br />Alt: {l.old || '—'} → Neu: {l.neu || '—'}</>}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* 8.2 Dashboard */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 12, marginBottom: 20 }}>
        {([
          ['incomplete', 'Abrechnungsrelevante Informationen unvollständig'],
          ['upcoming', 'Event ausstehend'],
          ['listPending', 'Teilnehmerlistenversand ausstehend'],
          ['settled', 'Abgerechnet'],
        ] as Array<[FAStatus, string]>).map(([key, label]) => (
          <button
            key={key}
            type="button"
            className="card"
            onClick={() => setStatusFilter(statusFilter === key ? '' : key)}
            style={{
              padding: '16px 18px', textAlign: 'left', cursor: 'pointer',
              border: statusFilter === key ? `2px solid ${FA_STATUS_COLORS[key].fg}` : '1px solid var(--dex-gray-200)',
            }}
          >
            <div style={{ fontSize: '1.8rem', fontWeight: 800, color: FA_STATUS_COLORS[key].fg }}>{counts[key]}</div>
            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.4 }}>{label}</div>
          </button>
        ))}
      </div>

      {/* 9. Tabelle */}
      <div className="card" style={{ padding: '20px 24px' }}>
        <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 14 }}>
          <input
            className="form-input"
            style={{ flex: '1 1 260px' }}
            placeholder="Suchen: Event-ID, Name, Ansprechpartner, Ariba-Nr., Datum …"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
          <select className="form-input" style={{ width: 'auto' }} value={statusFilter} onChange={e => setStatusFilter(e.target.value as '' | FAStatus)}>
            <option value="">Alle Status</option>
            {(Object.keys(FA_STATUS_LABELS) as FAStatus[]).map(s => <option key={s} value={s}>{FA_STATUS_LABELS[s]}</option>)}
          </select>
        </div>
        {filtered.length === 0 ? (
          <p style={{ margin: 0, color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
            {billingEvents.length === 0
              ? 'Es gibt noch keine abrechnungsrelevanten Events.'
              : 'Keine Treffer für die aktuelle Suche/Filterung.'}
          </p>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: '0.85rem' }}>
              <thead>
                <tr>
                  {['Event-ID', 'Eventname', 'Eventdatum', 'Ansprechpartner', 'Status'].map(h => (
                    <th key={h} style={{ textAlign: 'left', padding: '8px 14px 8px 0', borderBottom: '2px solid var(--dex-gray-200)', whiteSpace: 'nowrap' }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map(x => (
                  <tr
                    key={x.ev.id}
                    onClick={() => setSelectedId(x.ev.id)}
                    style={{ cursor: 'pointer' }}
                    onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'var(--dex-gray-50, #fafafa)'; }}
                    onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = ''; }}
                  >
                    <td style={{ padding: '8px 14px 8px 0', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>{x.ev.eventNumber || x.ev.id}</td>
                    <td style={{ padding: '8px 14px 8px 0', fontWeight: 600 }}>{x.ev.title}</td>
                    <td style={{ padding: '8px 14px 8px 0', whiteSpace: 'nowrap' }}>{fmtDate(x.ev.startDate)}</td>
                    <td style={{ padding: '8px 14px 8px 0' }}>{(x.b?.fields || {}).contact || (x.ev.organizers || [])[0] || '—'}</td>
                    <td style={{ padding: '8px 0' }}>{statusPill(x.status)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
