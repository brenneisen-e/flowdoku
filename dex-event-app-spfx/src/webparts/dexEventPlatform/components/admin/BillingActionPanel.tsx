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
import Modal from '../Modal';
import { cx } from '../dexUi';
import { ChevronDown, Download, FileText, Mail, Send } from '../Icons';
import {
  parseBillingOf, missingBillingFields, faStatusOf, FA_STATUS_LABELS, FA_STATUS_COLORS, FA_STATUS_NEXT,
  faRowsFromRegistrations, downloadFAParticipantXlsx,
} from '../../utils/faBilling';

/**
 * v30.53: „Anzeige der Empfänger vor dem Versand" (Fachkonzept, Bereiche 1–3).
 *
 * Die Adressen standen bisher NUR im Bestätigungsdialog — also erst, nachdem
 * man auf „Senden" gedrückt hat. Das Konzept meint aber die Anzeige davor:
 * Man soll sehen, wohin es geht, BEVOR man sich zum Klick entscheidet.
 *
 * v31.3: `loading` unterscheidet „Verteiler ist leer" von „noch nicht
 * gelesen". Bis dahin zeigte die Zeile während des laufenden getFAConfig()
 * dieselbe orange Warnung wie ein wirklich leerer Verteiler — ein
 * unbekannter Stand, der als Aussage über die Daten gelesen wurde.
 */
function RecipientLine(props: { label: string; addrs: string[]; withCc?: boolean; loading?: boolean }): React.ReactElement {
  return (
    <div className="dex-ui-help" style={{ marginTop: 6 }}>
      <strong style={{ color: 'var(--dex-gray-700)' }}>{props.label}: </strong>
      {props.loading
        ? <span>Verteiler wird geladen …</span>
        : props.addrs.length === 0
          ? <span style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>noch kein Verteiler hinterlegt — bitte im F&amp;A Center pflegen (lassen).</span>
          : <span style={{ wordBreak: 'break-word' }}>{props.addrs.join(', ')}{props.withCc ? ' · Organizer in CC' : ''}</span>}
    </div>
  );
}

const fmtDateTime = (iso: string): string => {
  const d = new Date(iso || '');
  return isFinite(d.getTime())
    ? d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ', ' + d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
    : '—';
};

export default function BillingActionPanel(props: { event: DeloitteEvent; onClose: () => void }): React.ReactElement | null {
  const { event } = props;
  const { events, sendFAMail, getFAConfig, getAllRegistrations, logFAContact } = useEvents();
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
  // v31.3: Der Absatz „warum kein Datei-Anhang" ist die Antwort auf EINE
  // Rückfrage — er stand als sechs Zeilen Fließtext dauerhaft in der Karte
  // und schob den Versand-Knopf nach unten. Jetzt zugeklappt.
  const [whyNoAttachment, setWhyNoAttachment] = React.useState(false);

  React.useEffect(() => {
    let cancelled = false;
    getFAConfig().then(c => { if (!cancelled) setRecipients({ info: c.infoRecipients, list: c.listRecipients }); }).catch(() => { /* */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (!b || !b.relevant) return null;

  // v31.3: Die beiden Versendungen sind eine Reihenfolge, kein Paar — erst die
  // Informationen, nach dem Event die Teilnehmerliste. Welcher Schritt gerade
  // dran ist, sagt der Status; nur DER trägt den Primär-Knopf (Leitfaden 1.5:
  // genau ein Primär-Knopf je Ansicht). Zwei gleich laute grüne Knöpfe haben
  // die Frage „was soll ich jetzt drücken?" offen gelassen.
  const nextSend: '' | 'info' | 'list' = status === 'readyToSend' ? 'info' : status === 'listPending' ? 'list' : '';

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
        // v30.67: sendFAMail unterscheidet jetzt, WARUM nichts (oder nur die
        // Haelfte) passiert ist. Vor allem 'stamp-failed' braucht eine eigene
        // Meldung: Die Mail IST raus, nur der Vermerk fehlt — die generische
        // "Versand fehlgeschlagen" laedt zum zweiten Klick und damit zur
        // Doppel-Mail an F&A ein.
        const reason = r.reason || '';
        const msg = reason === 'incomplete'
          ? 'Versand nicht möglich: Die Pflichtangaben sind noch unvollständig.'
          : reason === 'no-recipients'
            ? 'Versand nicht möglich: Kein F&A-Verteiler hinterlegt.'
            : reason === 'read-failed'
              ? 'Die Teilnehmerliste konnte nicht gelesen werden — Versand abgebrochen, es wurde nichts gesendet. Bitte später erneut versuchen.'
              : reason === 'no-participants'
                ? 'Keine aktiven Anmeldungen — es wurde nichts versendet.'
                : reason.indexOf('stamp-failed') === 0
                  ? 'Die Mail wurde versendet, aber der Vermerk in der Versand-Historie konnte NICHT gespeichert werden. Bitte NICHT erneut senden — sonst bekommt F&A die Liste doppelt. Details in der Browser-Konsole unter [DEX].'
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
      // v30.67 (Review): Ein Lesefehler ergab eine Datei mit null Zeilen —
      // die F&A für „keine Anmeldungen" gehalten hätte.
      let readFailed = false;
      const regs = await getAllRegistrations(liveEvent.id, () => { readFailed = true; });
      if (readFailed) {
        showAlert('Die Teilnehmerliste konnte gerade nicht gelesen werden — es wurde keine Datei erzeugt. Bitte später erneut versuchen.', { variant: 'error' });
        return;
      }
      // v30.50: Zeilenaufbau UND Dateiaufbau liegen in utils/faBilling —
      // dieselbe Datei, die auch das F&A Center herunterlädt. Vorher waren
      // es zwei Implementierungen, die schon im Spaltensatz auseinanderliefen.
      const rows = faRowsFromRegistrations(regs);
      if (rows.length === 0) {
        showAlert('Für dieses Event gibt es aktuell keine aktiven Anmeldungen.', { variant: 'info' });
        return;
      }
      await downloadFAParticipantXlsx(liveEvent, b, rows);
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
  const contactFA = async (): Promise<void> => {
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
    // v30.53: Das Konzept verlangt „Speicherung der Kommunikation in der
    // Kommunikationshistorie". Bei einem mailto: sieht DEX den Text NIE — die
    // Mail entsteht im Outlook des Organizers. Festgehalten wird deshalb, was
    // belegbar ist: dass und wann eine Rückfrage an wen begonnen wurde. Das
    // als „Mail gesendet" zu protokollieren wäre eine Behauptung über etwas,
    // das die App weder auslöst noch bestätigt bekommt.
    try { await logFAContact(liveEvent, to, subject); } catch { /* Protokoll ist best-effort */ }
  };

  // v30.53: eigenes Overlay durch das gemeinsame <Modal> ersetzt. Das
  // handgebaute schloss bei JEDEM Klick auf den Hintergrund — inklusive der
  // Klicks, die innen begannen und außen endeten (v30.51). Nebenbei erbt der
  // Dialog damit Escape, Portal und die Button-Styles.
  return (
    <Modal
      open onClose={props.onClose} maxWidth={680} ariaLabel="Event-Abrechnung"
      icon={<FileText size={20} />} title="Event-Abrechnung"
      // v31.3: Der Status stand als Pille neben der Überschrift, die nächste
      // Handlung nirgends — beides gehört in den Kopf.
      subtitle={
        <>
          <span className="dex-ui-pill" style={{ background: FA_STATUS_COLORS[status].bg, color: FA_STATUS_COLORS[status].fg }}>
            {FA_STATUS_LABELS[status]}
          </span>
          {FA_STATUS_NEXT[status] ? <div style={{ marginTop: 6 }}>{FA_STATUS_NEXT[status]}</div> : null}
          {b.settled && (
            <div style={{ marginTop: 4, color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>
              Abgerechnet am {fmtDateTime(b.settled.ts)} durch {b.settled.by}.
            </div>
          )}
        </>
      }
      footer={<button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={props.onClose}>Schließen</button>}
    >
      <div>
        <div className="dex-ui-section">
          <div className="dex-ui-section-title">Versand an F&amp;A</div>
          <p className="dex-ui-section-desc">
            Versandart: <strong>{b.sendMode === 'auto' ? 'Automatisiert' : 'Manuell'}</strong>
            {b.sendMode === 'auto'
              ? ' — die Mails gehen automatisch 7 Kalendertage vor bzw. nach dem Event raus. Du kannst hier zusätzlich jederzeit manuell senden.'
              : ' — beide Versendungen löst du hier aktiv aus.'}
          </p>

          <div className="dex-ui-stack">
            {/* v31.3: Zwei Karten mit Knopf am rechten Rand werden zu zwei
                Ablauf-Zeilen mit dem Knopf IN der Zeile (Leitfaden 1.4);
                `is-done` färbt die Nummer, sobald der Versand raus ist. */}
            <div className={cx('dex-ui-step', b.infoSentAt && 'is-done', missing.length > 0 && 'is-pending')} style={{ flexWrap: 'wrap' }}>
              <span className="dex-ui-step-num">1</span>
              <div className="dex-ui-step-body">
                <div className="dex-ui-step-title">Informationen an F&amp;A senden</div>
                <div className="dex-ui-step-hint">
                  {missing.length > 0
                    ? `Noch nicht möglich: ${missing.length} von 11 Pflichtangaben fehlen (Schritt „Abrechnung" im Wizard).`
                    : b.infoSentAt
                      ? `Zuletzt gesendet: ${fmtDateTime(b.infoSentAt)}.`
                      : 'Alle Pflichtangaben sind gepflegt — noch nicht versendet.'}
                </div>
                <RecipientLine label="Geht an" addrs={recipients?.info || []} withCc loading={recipients === null} />
              </div>
              <div className="dex-ui-step-action">
                <button type="button" className={cx('btn', nextSend === 'info' ? 'btn-primary' : 'btn-outline', 'dex-ui-btn-sm')}
                  disabled={busy !== '' || missing.length > 0} onClick={() => { void doSend('info'); }}>
                  <Send size={14} />{busy === 'info' ? 'Wird gesendet…' : b.infoSentAt ? 'Erneut senden' : 'Senden'}
                </button>
              </div>
            </div>

            <div className={cx('dex-ui-step', b.listSentAt && 'is-done')} style={{ flexWrap: 'wrap' }}>
              <span className="dex-ui-step-num">2</span>
              <div className="dex-ui-step-body">
                <div className="dex-ui-step-title">Teilnehmerliste an F&amp;A senden</div>
                <div className="dex-ui-step-hint">
                  {b.listSentAt
                    ? `Zuletzt gesendet: ${fmtDateTime(b.listSentAt)}${b.listSnapshot ? ` (${b.listSnapshot.length} Personen)` : ''}.`
                    : 'Sendet den aktuellen Stand aller nicht abgemeldeten Teilnehmer.'}
                </div>
                <RecipientLine label="Geht an" addrs={recipients?.list || []} withCc loading={recipients === null} />
                {/* v30.23: Warum kein Datei-Anhang? Der Tenant blockt Mails mit
                    Anhang komplett (s. Kommentar an downloadXlsx). */}
                <button type="button" className={cx('dex-ui-disclosure', whyNoAttachment && 'is-open')}
                  onClick={() => setWhyNoAttachment(!whyNoAttachment)}>
                  <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                  Warum ohne Excel-Anhang?
                </button>
                {whyNoAttachment && (
                  <div className="dex-ui-disclosure-body dex-ui-help">
                    Die Mail enthält die Liste als Tabelle und einen <strong>Download-Link ins F&amp;A Center</strong>,
                    über den F&amp;A sie als Excel-Datei zieht. Ein Datei-Anhang ist bewusst nicht dabei: Deloitte
                    blockt Mails mit Anhang, die aus Power Automate kommen — die Mail käme dann gar nicht an.
                    Über den Link bleibt die Teilnehmerliste außerdem in DEX (Zugriff nur mit F&amp;A- oder
                    Admin-Rolle) statt in Postfächern zu kursieren. &bdquo;Als Excel laden&ldquo; ist dein eigener Download.
                  </div>
                )}
              </div>
              <div className="dex-ui-step-action dex-ui-inline">
                <button type="button" className="btn btn-secondary dex-ui-btn-sm"
                  disabled={busy !== ''} onClick={() => { void downloadXlsx(); }}
                  title="Lädt dieselbe Liste als Excel-Datei herunter — zum Anhängen aus deinem eigenen Postfach.">
                  <Download size={14} />{busy === 'xlsx' ? 'Wird erzeugt…' : 'Als Excel laden'}
                </button>
                <button type="button" className={cx('btn', nextSend === 'list' ? 'btn-primary' : 'btn-outline', 'dex-ui-btn-sm')}
                  disabled={busy !== ''} onClick={() => { void doSend('list'); }}>
                  <Send size={14} />{busy === 'list' ? 'Wird gesendet…' : b.listSentAt ? 'Erneut senden' : 'Senden'}
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* v30.23 (F&A-Nachlieferung §2): dritte Aktion — Rückfragen an F&A.
            v31.3: Keine dritte Ablauf-Zeile, denn eine Rückfrage ist kein
            Schritt der Reihenfolge, sondern eine eigene Absicht — deshalb ein
            eigener Abschnitt mit der Aktion als Kachel (Knopf links am Text). */}
        <div className="dex-ui-section">
          <div className="dex-ui-section-title">Rückfragen</div>
          <button type="button" className="dex-ui-action" onClick={() => { void contactFA(); }}>
            <span className="dex-ui-action-icon"><Mail size={16} /></span>
            <span className="dex-ui-action-body">
              <span className="dex-ui-action-title">F&amp;A-Team kontaktieren</span>
              <span className="dex-ui-action-desc">
                Öffnet eine vorbereitete E-Mail an F&amp;A in deinem Outlook — mit Event, Event-ID und Status.
                Bewusst aus deinem Postfach, damit die Antwort bei dir ankommt.
              </span>
            </span>
          </button>
          <RecipientLine label="Ansprechpartner" addrs={recipients?.info || []} loading={recipients === null} />
          <div className="dex-ui-help">In der Historie unten wird festgehalten, dass du eine Rückfrage begonnen hast — den Text selbst sieht DEX nicht.</div>
        </div>

        <div className="dex-ui-section">
          <div className="dex-ui-card-head" style={{ marginBottom: 10 }}>
            <h4 className="dex-ui-card-head-title">Versandhistorie</h4>
            {mails.length > 0 && <span className="dex-ui-card-head-meta">{mails.length} {mails.length === 1 ? 'Eintrag' : 'Einträge'}</span>}
          </div>
          {mails.length === 0 ? (
            <div className="dex-ui-empty">
              <span className="dex-ui-empty-icon"><Mail size={20} /></span>
              <div className="dex-ui-empty-title">Noch keine abrechnungsrelevanten E-Mails versendet.</div>
              <div>Was du oben sendest, steht hier mit Zeitpunkt, Empfänger und vollständigem Inhalt.</div>
            </div>
          ) : (
            /* v31.3: Aus gestapelten Kärtchen wird eine kompakte Tabelle — drei
               Vorgänge zu vergleichen hieß vorher scrollen. Der Mail-Inhalt
               bleibt aufklappbar, jetzt als eigene Zeile über die volle Breite. */
            <div className="dex-ui-table-wrap">
              <table className="dex-ui-table dex-ui-table--compact">
                <thead>
                  <tr>
                    <th>Zeitpunkt</th>
                    <th>Art</th>
                    <th>Empfänger</th>
                    <th className="is-actions"><span className="dex-ui-sr-only">Inhalt</span></th>
                  </tr>
                </thead>
                <tbody>
                  {mails.map((m, i) => (
                    <React.Fragment key={i}>
                      <tr>
                        <td>
                          {fmtDateTime(m.ts)}
                          <div className="dex-ui-row-sub">Ausgelöst von {m.by}</div>
                        </td>
                        <td>
                          <span className={cx('dex-ui-pill', m.mailType === 'info' ? 'dex-ui-pill--blue' : 'dex-ui-pill--green')}>
                            {m.mailType === 'info' ? 'Abrechnungsinformationen' : 'Teilnehmerliste'}
                          </span>
                          <div className="dex-ui-row-sub">Betreff: {m.subject || '—'}</div>
                        </td>
                        <td style={{ wordBreak: 'break-word' }}>
                          {m.to || '—'}
                          {m.cc ? <div className="dex-ui-row-sub">CC: {m.cc}</div> : null}
                        </td>
                        <td className="is-actions">
                          {m.body && (
                            <button type="button" className="dex-ui-textbtn" title="Vollständigen Inhalt der Mail anzeigen"
                              onClick={() => setOpenMailIdx(openMailIdx === i ? null : i)}>
                              {openMailIdx === i ? 'Inhalt ausblenden' : 'Inhalt anzeigen'}
                            </button>
                          )}
                        </td>
                      </tr>
                      {m.body && openMailIdx === i && (
                        <tr>
                          <td colSpan={4}>
                            <div style={{ border: '1px dashed var(--dex-gray-300)', borderRadius: 8, padding: 12, overflowX: 'auto' }}
                              dangerouslySetInnerHTML={{ __html: m.body }} />
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </Modal>
  );
}
