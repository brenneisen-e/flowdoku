/**
 * v30.5: Organizer-Center-Aktion „Event-Abrechnung" (Fachkonzept Abschnitt 6).
 *
 * Modal mit den beiden Versand-Funktionen („Informationen an F&A versenden",
 * „Teilnehmerliste an F&A versenden") und der Versandhistorie. Der eigentliche
 * Versand + die Protokollierung liegen in EventContext.sendFAMail — hier steht
 * nur die Bedienung. Erscheint ausschließlich bei abrechnungsrelevanten
 * Events (Gate in AdminPage).
 */

import * as React from 'react';
import { useEvents } from '../../context/EventContext';
import { useDialog } from '../../context/DialogContext';
import { DeloitteEvent } from '../../types';
import {
  parseBillingOf, missingBillingFields, faStatusOf, FA_STATUS_LABELS, FA_STATUS_COLORS,
} from '../../utils/faBilling';

const fmtDateTime = (iso: string): string => {
  const d = new Date(iso || '');
  return isFinite(d.getTime())
    ? d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ', ' + d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
    : '—';
};

export default function BillingActionPanel(props: { event: DeloitteEvent; onClose: () => void }): React.ReactElement | null {
  const { event } = props;
  const { events, sendFAMail, getFAConfig, getAllRegistrations } = useEvents();
  const { confirmDialog, showAlert } = useDialog();
  // Frischen Stand aus dem Context nehmen — nach einem Versand patcht
  // sendFAMail den lokalen Event-State, das Modal soll sofort mitziehen.
  const liveEvent = events.find(e => e.id === event.id) || event;
  const b = parseBillingOf(liveEvent);
  const status = faStatusOf(liveEvent, b);
  const missing = missingBillingFields(b);
  const mails = (b?.log || []).filter(l => l.mailType);

  const [recipients, setRecipients] = React.useState<{ info: string[]; list: string[] } | null>(null);
  const [busy, setBusy] = React.useState<'' | 'info' | 'list' | 'xlsx'>('');
  const [openMailIdx, setOpenMailIdx] = React.useState<number | null>(null);

  React.useEffect(() => {
    let cancelled = false;
    getFAConfig().then(c => { if (!cancelled) setRecipients({ info: c.infoRecipients, list: c.listRecipients }); }).catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (!b || !b.relevant) return null;

  const doSend = async (kind: 'info' | 'list'): Promise<void> => {
    if (busy) return;
    const target = kind === 'info' ? (recipients?.info || []) : (recipients?.list || []);
    if (target.length === 0) {
      showAlert('Für diesen Versand ist noch kein F&A-Verteiler hinterlegt — bitte im F&A Center pflegen (lassen).', { variant: 'error' });
      return;
    }
    const ok = await confirmDialog(
      kind === 'info'
        ? `Alle aktuell gepflegten Abrechnungsinformationen jetzt an F&A senden?\n\nEmpfänger: ${target.join(', ')} — die Organizer stehen in CC. Der Versand wird protokolliert.`
        : `Die aktuelle Teilnehmerliste jetzt an F&A senden?\n\nEmpfänger: ${target.join(', ')} — die Organizer stehen in CC. Der Versand wird protokolliert.`,
      { confirmLabel: 'Jetzt senden' },
    );
    if (!ok) return;
    setBusy(kind);
    try {
      const r = await sendFAMail(liveEvent, kind);
      if (r.ok) {
        showAlert(kind === 'info' ? 'Abrechnungsinformationen wurden an F&A gesendet.' : 'Teilnehmerliste wurde an F&A gesendet.', { variant: 'success' });
      } else {
        const msg = r.reason === 'incomplete'
          ? 'Versand nicht möglich: Die Pflichtangaben sind noch unvollständig.'
          : r.reason === 'no-recipients'
            ? 'Versand nicht möglich: Kein F&A-Verteiler hinterlegt.'
            : 'Versand fehlgeschlagen — bitte später erneut versuchen.';
        showAlert(msg, { variant: 'error' });
      }
    } finally { setBusy(''); }
  };

  /**
   * v30.23 (F&A-Nachlieferung §2): Teilnehmerliste als Excel.
   *
   * Das Fachkonzept verlangt die Liste „als Excel-Datei im Anhang der
   * E-Mail". Das geht im Deloitte-Tenant NICHT: Die Mail-Flow-Regel blockt
   * JEDE über Power Automate versendete Mail, die einen Anhang trägt (NDR
   * „Power Apps and Power Automate cannot be used to … send email
   * attachments") — die Mail käme dann gar nicht erst an. Genau daran ist
   * schon der .eml-Anhang gescheitert (v26.71, in docs/flow-jsons.md als
   * OBSOLET dokumentiert). Deshalb derselbe Weg wie dort: Die Mail trägt
   * die Liste als HTML-Tabelle (kommt garantiert an, wird protokolliert),
   * und wer die Datei braucht, lädt sie hier herunter und hängt sie aus
   * dem eigenen Postfach an. Datenbasis identisch zur Mail
   * (getAllRegistrations ohne Abgemeldete), damit beide nie auseinanderlaufen.
   */
  const downloadXlsx = async (): Promise<void> => {
    if (busy) return;
    setBusy('xlsx');
    try {
      const regs = await getAllRegistrations(liveEvent.id);
      const rows = (regs || [])
        .filter(r => r.Status !== 'Abgemeldet')
        .map((r, i) => [
          i + 1,
          (r.ParticipantName || `${r.Vorname || ''} ${r.Nachname || ''}`.trim()) || r.ParticipantEmail || '—',
          r.ParticipantEmail || '',
          r.Status || '',
        ]);
      if (rows.length === 0) {
        showAlert('Für dieses Event gibt es aktuell keine aktiven Anmeldungen.', { variant: 'info' });
        return;
      }
      const headers = ['#', 'Name', 'E-Mail', 'Status'];
      const safeName = (liveEvent.title || 'event').replace(/[^a-zA-Z0-9]/g, '_');
      const fileName = `Teilnehmerliste_FA_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`;
      // xlsx erst beim Klick nachladen (schwerste Dependency, s. AdminPage).
      const XLSX = await import('xlsx');
      const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ws as any)['!cols'] = [{ wch: 5 }, { wch: 30 }, { wch: 34 }, { wch: 14 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Teilnehmer');
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      // Manueller Anker-Download — im SPFx-Iframe ist saveAs oft blockiert.
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = fileName; a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
    } catch (err) {
      console.warn('[DEX] F&A-XLSX-Export fehlgeschlagen:', err);
      showAlert('Die Excel-Datei konnte nicht erzeugt werden — bitte erneut versuchen.', { variant: 'error' });
    } finally { setBusy(''); }
  };

  /**
   * v30.23 (F&A-Nachlieferung §2): „F&A-Team kontaktieren".
   *
   * Bewusst ein mailto: aus dem Postfach des Organizers und NICHT über die
   * DEX-Mail-Queue: Es geht um Rückfragen — die Antwort von F&A muss beim
   * Organizer landen. Eine Queue-Mail käme von no_reply.events@deloitte.de,
   * eine Antwort darauf ginge ins Leere.
   */
  const contactFA = (): void => {
    const to = (recipients?.info || []).join(';');
    if (!to) {
      showAlert('Es ist noch kein F&A-Verteiler hinterlegt — bitte im F&A Center pflegen (lassen).', { variant: 'error' });
      return;
    }
    const subject = `Frage zur Abrechnung: ${liveEvent.title} (Event-ID ${liveEvent.eventNumber || liveEvent.id})`;
    const bodyLines = [
      'Guten Tag,',
      '',
      'ich habe eine Frage zur Abrechnung der folgenden Veranstaltung:',
      '',
      `Event: ${liveEvent.title}`,
      `Event-ID: ${liveEvent.eventNumber || liveEvent.id}`,
      `Datum: ${fmtDateTime(liveEvent.startDate || '')}`,
      `Status der Angaben: ${FA_STATUS_LABELS[status]}`,
      '',
      '[Deine Frage]',
      '',
      'Viele Grüße',
    ];
    window.location.href = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyLines.join('\r\n'))}`;
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Event-Abrechnung"
      style={{
        position: 'fixed', inset: 0, zIndex: 2000, background: 'rgba(0,0,0,0.55)',
        display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16,
      }}
      onClick={props.onClose}
    >
      <div
        style={{
          background: '#fff', borderRadius: 12, padding: '24px 28px',
          maxWidth: 680, width: '100%', maxHeight: '86vh', overflowY: 'auto',
          boxShadow: '0 16px 48px rgba(0,0,0,0.35)',
        }}
        onClick={e => e.stopPropagation()}
      >
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 12 }}>
          <div>
            <h3 style={{ margin: '0 0 6px' }}>Event-Abrechnung</h3>
            <span style={{
              display: 'inline-block', padding: '3px 10px', borderRadius: 999, fontSize: '0.72rem', fontWeight: 700,
              background: FA_STATUS_COLORS[status].bg, color: FA_STATUS_COLORS[status].fg,
            }}>{FA_STATUS_LABELS[status]}</span>
          </div>
          <button type="button" className="btn btn-secondary" style={{ fontSize: '0.8rem', padding: '4px 12px' }} onClick={props.onClose}>Schließen</button>
        </div>

        <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, margin: '12px 0 16px' }}>
          Versandart: <strong>{b.sendMode === 'auto' ? 'Automatisiert' : 'Manuell'}</strong>
          {b.sendMode === 'auto'
            ? ' — die Mails gehen automatisch 7 Kalendertage vor bzw. nach dem Event raus. Du kannst hier zusätzlich jederzeit manuell senden.'
            : ' — beide Versendungen löst du hier aktiv aus.'}
          {b.settled && <> · <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>Abgerechnet am {fmtDateTime(b.settled.ts)} durch {b.settled.by}.</strong></>}
        </p>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginBottom: 18 }}>
          <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: '12px 14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
              <div style={{ flex: '1 1 260px' }}>
                <strong style={{ fontSize: '0.88rem' }}>Informationen an F&amp;A versenden</strong>
                <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                  {missing.length > 0
                    ? `Noch nicht möglich: ${missing.length} von 11 Pflichtangaben fehlen (Schritt „Abrechnung" im Wizard).`
                    : b.infoSentAt
                      ? `Zuletzt gesendet: ${fmtDateTime(b.infoSentAt)}.`
                      : 'Alle Pflichtangaben sind gepflegt — noch nicht versendet.'}
                </div>
              </div>
              <button
                type="button"
                className="btn btn-primary"
                style={{ fontSize: '0.8rem', padding: '6px 16px', flexShrink: 0 }}
                disabled={busy !== '' || missing.length > 0}
                onClick={() => { void doSend('info'); }}
              >
                {busy === 'info' ? 'Wird gesendet…' : 'Senden'}
              </button>
            </div>
          </div>
          <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: '12px 14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
              <div style={{ flex: '1 1 260px' }}>
                <strong style={{ fontSize: '0.88rem' }}>Teilnehmerliste an F&amp;A versenden</strong>
                <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                  {b.listSentAt
                    ? `Zuletzt gesendet: ${fmtDateTime(b.listSentAt)}${b.listSnapshot ? ` (${b.listSnapshot.length} Personen)` : ''}.`
                    : 'Sendet den aktuellen Stand aller nicht abgemeldeten Teilnehmer.'}
                </div>
              </div>
              <div style={{ display: 'inline-flex', gap: 8, flexShrink: 0, flexWrap: 'wrap' }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  style={{ fontSize: '0.8rem', padding: '6px 14px' }}
                  disabled={busy !== ''}
                  onClick={() => { void downloadXlsx(); }}
                  title="Lädt dieselbe Liste als Excel-Datei herunter — zum Anhängen aus deinem eigenen Postfach."
                >
                  {busy === 'xlsx' ? 'Wird erzeugt…' : 'Als Excel laden'}
                </button>
                <button
                  type="button"
                  className="btn btn-primary"
                  style={{ fontSize: '0.8rem', padding: '6px 16px' }}
                  disabled={busy !== ''}
                  onClick={() => { void doSend('list'); }}
                >
                  {busy === 'list' ? 'Wird gesendet…' : 'Senden'}
                </button>
              </div>
            </div>
            {/* v30.23: Warum kein Datei-Anhang? Der Tenant blockt Mails mit
                Anhang komplett (s. Kommentar an downloadXlsx). */}
            <div style={{ marginTop: 8, fontSize: '0.72rem', color: 'var(--dex-gray-500)', lineHeight: 1.5 }}>
              Die Mail enthält die Liste als Tabelle und einen <strong>Download-Link ins F&amp;A Center</strong>,
              über den F&amp;A sie als Excel-Datei zieht. Ein Datei-Anhang ist bewusst nicht dabei: Deloitte
              blockt Mails mit Anhang, die aus Power Automate kommen — die Mail käme dann gar nicht an.
              Über den Link bleibt die Teilnehmerliste außerdem in DEX (Zugriff nur mit F&amp;A- oder
              Admin-Rolle) statt in Postfächern zu kursieren. &bdquo;Als Excel laden&ldquo; ist dein eigener Download.
            </div>
          </div>
          {/* v30.23 (F&A-Nachlieferung §2): dritte Aktion — Rückfragen an F&A. */}
          <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: '12px 14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
              <div style={{ flex: '1 1 260px' }}>
                <strong style={{ fontSize: '0.88rem' }}>F&amp;A-Team kontaktieren</strong>
                <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                  Öffnet eine vorbereitete E-Mail an F&amp;A in deinem Outlook — mit Event, Event-ID und
                  Status. Bewusst aus deinem Postfach, damit die Antwort bei dir ankommt.
                </div>
              </div>
              <button
                type="button"
                className="btn btn-secondary"
                style={{ fontSize: '0.8rem', padding: '6px 16px', flexShrink: 0 }}
                onClick={contactFA}
              >
                E-Mail schreiben
              </button>
            </div>
          </div>
        </div>

        <h4 style={{ margin: '0 0 8px', fontSize: '0.9rem' }}>Versandhistorie</h4>
        {mails.length === 0 ? (
          <p style={{ margin: 0, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>Noch keine abrechnungsrelevanten E-Mails versendet.</p>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {mails.map((m, i) => (
              <div key={i} style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: '10px 12px', fontSize: '0.78rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap' }}>
                  <strong>{m.mailType === 'info' ? 'Abrechnungsinformationen' : 'Teilnehmerliste'}</strong>
                  <span style={{ color: 'var(--dex-gray-500)' }}>{fmtDateTime(m.ts)}</span>
                </div>
                <div style={{ color: 'var(--dex-gray-600)', marginTop: 4, lineHeight: 1.5 }}>
                  Empfänger: {m.to || '—'}{m.cc ? <> · CC: {m.cc}</> : null}<br />
                  Betreff: {m.subject || '—'} · Ausgelöst von {m.by}
                </div>
                {m.body && (
                  <>
                    <button
                      type="button"
                      className="btn btn-secondary"
                      style={{ marginTop: 8, fontSize: '0.72rem', padding: '3px 10px' }}
                      onClick={() => setOpenMailIdx(openMailIdx === i ? null : i)}
                    >
                      {openMailIdx === i ? 'Inhalt ausblenden' : 'Vollständigen Inhalt anzeigen'}
                    </button>
                    {openMailIdx === i && (
                      <div
                        style={{ marginTop: 10, border: '1px dashed var(--dex-gray-300)', borderRadius: 8, padding: 12, overflowX: 'auto' }}
                        dangerouslySetInnerHTML={{ __html: m.body }}
                      />
                    )}
                  </>
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
