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
  const { events, sendFAMail, getFAConfig } = useEvents();
  const { confirmDialog, showAlert } = useDialog();
  // Frischen Stand aus dem Context nehmen — nach einem Versand patcht
  // sendFAMail den lokalen Event-State, das Modal soll sofort mitziehen.
  const liveEvent = events.find(e => e.id === event.id) || event;
  const b = parseBillingOf(liveEvent);
  const status = faStatusOf(liveEvent, b);
  const missing = missingBillingFields(b);
  const mails = (b?.log || []).filter(l => l.mailType);

  const [recipients, setRecipients] = React.useState<{ info: string[]; list: string[] } | null>(null);
  const [busy, setBusy] = React.useState<'' | 'info' | 'list'>('');
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
              <button
                type="button"
                className="btn btn-primary"
                style={{ fontSize: '0.8rem', padding: '6px 16px', flexShrink: 0 }}
                disabled={busy !== ''}
                onClick={() => { void doSend('list'); }}
              >
                {busy === 'list' ? 'Wird gesendet…' : 'Senden'}
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
