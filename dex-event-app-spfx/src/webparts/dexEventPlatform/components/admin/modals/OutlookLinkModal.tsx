/**
 * v31.52: „Outlook-Verknüpfung anzeigen / ändern" — zeigt, welche iCalUId
 * (Spalte CalendarLink) in SharePoint an der gewählten Zeile hängt, prüft auf
 * Wunsch über Graph, ob es den Termin gibt, und überschreibt den Wert.
 * Regeln in `logic/recycleBinOutlook.ts` (`verknuepfeTermin`: erst der
 * umkehrbare MERGE, dann — nur auf Wunsch — die Absage des bisherigen Termins).
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Link2 } from '../../Icons';
import type { EventService } from '../../../services/EventService';
import type { DeloitteEvent } from '../../../types';
import {
  LinkPruefung, NO_REPLY_MAILBOX, OutlookVerknuepfung, VerknuepfErgebnis,
  leseOutlookVerknuepfung, pruefeCalendarLink, verknuepfeTermin,
} from '../logic/recycleBinOutlook';

export interface OutlookLinkModalProps {
  isDe: boolean;
  svc: EventService;
  event: DeloitteEvent;
  /** Hauptevent zuerst, dann die Sub-Events. */
  zeilen: DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  refreshEvents: () => Promise<void>;
  onClose: () => void;
}

const fmtDatum = (iso: string, isDe: boolean): string => {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
};

/** Grobe Plausibilität: iCalUIds sind lang und ohne Leerzeichen. */
const siehtAusWieUid = (s: string): boolean => s.length >= 20 && !/\s/.test(s);

export const OutlookLinkModal: React.FC<OutlookLinkModalProps> = (p) => {
  const { isDe, svc, event, zeilen, confirmDialog, refreshEvents, onClose } = p;
  const [zeileId, setZeileId] = React.useState<string>(zeilen[0]?.id || '');
  const [stand, setStand] = React.useState<OutlookVerknuepfung | null | 'laedt' | 'fehler'>('laedt');
  const [neu, setNeu] = React.useState('');
  const [pruefung, setPruefung] = React.useState<LinkPruefung | null>(null);
  const [pruefBusy, setPruefBusy] = React.useState(false);
  const [absagen, setAbsagen] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [ergebnis, setErgebnis] = React.useState<VerknuepfErgebnis | null>(null);
  const [kopiert, setKopiert] = React.useState(false);
  const zeile = zeilen.find(z => z.id === zeileId) || null;

  const laden = React.useCallback(async (id: string): Promise<void> => {
    setStand('laedt');
    setErgebnis(null);
    setPruefung(null);
    setNeu('');
    try { setStand(await leseOutlookVerknuepfung(svc, id)); } catch { setStand('fehler'); }
  }, [svc]);
  React.useEffect(() => { if (zeileId) { void laden(zeileId); } }, [zeileId, laden]);

  const aktuell = stand && typeof stand === 'object' ? stand : null;
  // v31.53: Leerzeichen und Zeilenumbrüche fallen weg — die UID-Zeile einer
  // .ics ist nach 75 Zeichen umgebrochen (Folgezeile mit führendem Blank),
  // und genau so wird sie eingefügt. Der Nutzer sieht, dass bereinigt wurde.
  const neuTrim = neu.replace(/\s+/g, '');
  const bereinigt = neuTrim !== neu.trim() && neuTrim.length > 0;
  const formatOk = siehtAusWieUid(neuTrim);
  const hexartig = /^[0-9A-Fa-f]+$/.test(neuTrim);
  const gleichWieHinterlegt = !!aktuell && neuTrim.length > 0 && neuTrim === aktuell.calendarLink;
  const kannUebernehmen = !!aktuell && formatOk && !gleichWieHinterlegt && !busy;

  const kopieren = (): void => {
    if (!aktuell || !aktuell.calendarLink) return;
    const done = (): void => { setKopiert(true); setTimeout(() => setKopiert(false), 1400); };
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(aktuell.calendarLink).then(done, () => { /* Nutzer markiert von Hand */ }); }
    } catch { /* s. o. */ }
  };

  const pruefen = async (): Promise<void> => {
    if (!siehtAusWieUid(neuTrim) || pruefBusy) return;
    setPruefBusy(true);
    try { setPruefung(await pruefeCalendarLink(svc, neuTrim)); } finally { setPruefBusy(false); }
  };

  const uebernehmen = async (): Promise<void> => {
    if (!kannUebernehmen || !aktuell || !zeile) return;
    const gefunden = pruefung && pruefung.status === 'gefunden' ? pruefung : null;
    const ok = await confirmDialog(
      isDe
        ? <>
          <p style={{ margin: '0 0 10px' }}>
            Die Zeile <strong>{aktuell.id}</strong> (&bdquo;{aktuell.title}&ldquo;) bekommt als Outlook-Verknüpfung
            {gefunden ? <> den Termin <strong>&bdquo;{gefunden.subject}&ldquo;</strong> ({fmtDatum(gefunden.start, isDe)}, {gefunden.attendees} eingeladen)</> : <> die eingegebene iCalUId <strong>(nicht geprüft)</strong></>}.
          </p>
          <p style={{ margin: 0 }}>
            Ab dann treffen Einladungen, Absagen und Aktualisierungen aus DEX diesen Termin.
            {' '}{aktuell.calendarLink
              ? (absagen ? 'Der bisher verknüpfte Termin wird abgesagt — nur an die, die auf ihm stehen.' : 'Der bisher verknüpfte Termin bleibt stehen, hängt aber an keiner DEX-Zeile mehr.')
              : 'Bisher war kein Termin verknüpft.'}
          </p>
        </>
        : <>
          <p style={{ margin: '0 0 10px' }}>
            Row <strong>{aktuell.id}</strong> (&ldquo;{aktuell.title}&rdquo;) will be linked to
            {gefunden ? <> the appointment <strong>&ldquo;{gefunden.subject}&rdquo;</strong> ({fmtDatum(gefunden.start, isDe)}, {gefunden.attendees} invited)</> : <> the entered iCalUId <strong>(not verified)</strong></>}.
          </p>
          <p style={{ margin: 0 }}>
            From then on, invites, cancellations and updates from DEX hit this appointment.
            {' '}{aktuell.calendarLink
              ? (absagen ? 'The previously linked appointment is cancelled — only for its attendees.' : 'The previously linked appointment stays but is no longer linked to any DEX row.')
              : 'No appointment was linked before.'}
          </p>
        </>,
      { confirmLabel: isDe ? 'Überschreiben' : 'Overwrite', title: isDe ? 'Outlook-Verknüpfung ändern' : 'Change Outlook link', danger: !gefunden },
    );
    if (!ok) return;
    setBusy(true);
    try {
      const e = await verknuepfeTermin(svc, event, zeile, { iCalUId: neuTrim, subject: gefunden?.subject, attendees: gefunden?.attendees }, absagen);
      setErgebnis(e);
      if (e.status === 'fertig') {
        try { await refreshEvents(); } catch { /* Anzeige zieht beim nächsten Laden nach */ }
        try { setStand(await leseOutlookVerknuepfung(svc, zeile.id)); } catch { /* bleibt */ }
        setNeu('');
        setPruefung(null);
      }
    } finally {
      setBusy(false);
    }
  };

  const ergebnisText = (e: VerknuepfErgebnis): { text: string; fehler: boolean } => {
    const bisherDe: Record<string, string> = {
      eingereiht: 'Die Absage des bisher verknüpften Termins ist eingereiht.',
      behalten: 'Der bisher verknüpfte Termin bleibt stehen.',
      keiner: '',
      fehlgeschlagen: 'ACHTUNG: Die Absage des bisher verknüpften Termins konnte nicht eingereiht werden — bitte in Outlook von Hand absagen.',
    };
    const bisherEn: Record<string, string> = {
      eingereiht: 'The cancellation of the previously linked appointment is queued.',
      behalten: 'The previously linked appointment stays.',
      keiner: '',
      fehlgeschlagen: 'ATTENTION: the cancellation of the previously linked appointment could not be queued — please cancel it in Outlook manually.',
    };
    const de: Record<string, string> = {
      'fertig': `Überschrieben. ${bisherDe[e.bisheriger]} Öffne das Event einmal neu im Assistenten, damit die Anzeige den Termin kennt.`,
      'schon-verknuepft': 'Dieser Wert steht bereits in der Zeile — nichts zu tun.',
      'zeile-nicht-lesbar': 'Die Zeile war nicht lesbar — es wurde nichts geschrieben.',
      'schreiben-fehlgeschlagen': 'SharePoint hat das Schreiben abgelehnt — es wurde nichts verändert. Bitte gleich noch einmal versuchen.',
    };
    const en: Record<string, string> = {
      'fertig': `Overwritten. ${bisherEn[e.bisheriger]} Open the event in the wizard once so the display picks up the appointment.`,
      'schon-verknuepft': 'This value is already in the row — nothing to do.',
      'zeile-nicht-lesbar': 'The row could not be read — nothing was written.',
      'schreiben-fehlgeschlagen': 'SharePoint rejected the write — nothing was changed. Please try again in a moment.',
    };
    return { text: ((isDe ? de : en)[e.status] || e.status).trim(), fehler: e.status !== 'fertig' || e.bisheriger === 'fehlgeschlagen' };
  };

  const mono: React.CSSProperties = { fontFamily: 'monospace', fontSize: '0.78rem', wordBreak: 'break-all', color: 'var(--dex-gray-800)' };
  const labelStyle: React.CSSProperties = { fontSize: '0.74rem', letterSpacing: '0.04em', textTransform: 'uppercase', color: 'var(--dex-gray-600, #666)', marginBottom: 4 };

  return (
    <Modal
      open={true}
      onClose={() => { if (!busy) onClose(); }}
      dismissable={!busy}
      backdropClose={false}
      maxWidth={640}
      ariaLabel={isDe ? 'Outlook-Verknüpfung' : 'Outlook link'}
      title={isDe ? 'Outlook-Verknüpfung anzeigen / ändern' : 'Show / change Outlook link'}
      subtitle={isDe
        ? <>Die Spalte <code>CalendarLink</code> (iCalUId) der gewählten Zeile in DEX_Events — frisch aus SharePoint gelesen. Einladungen, Absagen und Aktualisierungen aus DEX gehen an genau diesen Termin.</>
        : <>The <code>CalendarLink</code> column (iCalUId) of the chosen DEX_Events row — read fresh from SharePoint. Invites, cancellations and updates from DEX go to exactly this appointment.</>}
      icon={<Link2 size={20} />}
      footer={(
        <button type="button" className="btn btn-secondary" disabled={busy} onClick={onClose}>
          {isDe ? 'Schließen' : 'Close'}
        </button>
      )}
    >
      <div className="dex-ui-stack" style={{ gap: 16 }}>
        <div>
          <div style={labelStyle}>{isDe ? 'Zeile' : 'Row'}</div>
          <select
            value={zeileId}
            disabled={busy}
            onChange={ev => setZeileId(ev.target.value)}
            style={{ width: '100%', padding: '8px 10px', borderRadius: 8, border: '1px solid var(--dex-gray-300, #ccc)', fontSize: '0.88rem' }}
          >
            {zeilen.map((z, i) => (
              <option key={z.id} value={z.id}>{i === 0 ? (isDe ? 'Hauptevent · ' : 'Main event · ') : (isDe ? 'Sub-Event · ' : 'Sub-event · ')}{z.title || (isDe ? 'ohne Titel' : 'untitled')} · {isDe ? 'Zeile' : 'row'} {z.id}</option>
            ))}
          </select>
        </div>

        <div className="dex-ui-card dex-ui-card--soft dex-ui-stack">
          <div style={labelStyle}>{isDe ? 'Aktuell in SharePoint' : 'Currently in SharePoint'}</div>
          {stand === 'laedt' ? (
            <div className="dex-ui-muted">{isDe ? 'Zeile wird gelesen …' : 'Reading row …'}</div>
          ) : stand === 'fehler' || stand === null ? (
            <div className="dex-ui-callout dex-ui-callout--danger dex-ui-callout--sm"><span>{isDe ? 'Die Zeile war nicht lesbar — der Stand ist unbekannt.' : 'The row could not be read — the state is unknown.'}</span></div>
          ) : (
            <>
              <div className="dex-ui-inline" style={{ justifyContent: 'space-between', gap: 8, flexWrap: 'wrap' }}>
                <strong style={{ fontSize: '0.9rem', minWidth: 0 }}>{aktuell!.title}</strong>
                <span className="dex-ui-inline" style={{ gap: 6 }}>
                  {aktuell!.disableOutlook && <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'Outlook aus' : 'Outlook off'}</span>}
                  {aktuell!.outlookEventId === 'FAILED' && <span className="dex-ui-pill dex-ui-pill--red">{isDe ? 'Anlegen war gescheitert' : 'creation had failed'}</span>}
                </span>
              </div>
              {aktuell!.calendarLink ? (
                <>
                  <div style={mono}>{aktuell!.calendarLink}</div>
                  <div className="dex-ui-inline" style={{ gap: 8, flexWrap: 'wrap' }}>
                    <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={kopieren}>{kopiert ? (isDe ? 'Kopiert' : 'Copied') : (isDe ? 'Kopieren' : 'Copy')}</button>
                    <span className="dex-ui-muted" style={{ fontSize: '0.8rem' }}>{isDe ? 'Zeile geändert' : 'Row modified'} {fmtDatum(aktuell!.modified, isDe)}</span>
                  </div>
                </>
              ) : (
                <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm"><span>{isDe ? 'Kein Termin verknüpft (CalendarLink leer).' : 'No appointment linked (CalendarLink empty).'}</span></div>
              )}
            </>
          )}
        </div>

        <div className="dex-ui-stack">
          <div style={labelStyle}>{isDe ? 'Neuer Wert (iCalUId)' : 'New value (iCalUId)'}</div>
          <textarea
            id="dex-outlook-link-neu"
            value={neu}
            disabled={!aktuell || busy}
            onChange={ev => { setNeu(ev.target.value); setPruefung(null); setErgebnis(null); }}
            rows={3}
            placeholder={isDe ? 'iCalUId des Termins einfügen — z. B. aus „Outlook-Termin wiederfinden", der Run history von DEX_Outlook_Einladungen oder der UID-Zeile einer .ics-Datei' : 'Paste the appointment’s iCalUId — e.g. from “Recover the Outlook appointment”, the run history of DEX_Outlook_Einladungen or the UID line of an .ics file'}
            style={{ width: '100%', boxSizing: 'border-box', padding: '8px 10px', borderRadius: 8, border: '1px solid var(--dex-gray-300, #ccc)', fontFamily: 'monospace', fontSize: '0.8rem' }}
          />
          {neuTrim.length > 0 && (
            <div className="dex-ui-inline" style={{ gap: 6, flexWrap: 'wrap', fontSize: '0.8rem' }}>
              <span className={`dex-ui-pill ${formatOk ? 'dex-ui-pill--green' : 'dex-ui-pill--red'}`}>
                {formatOk
                  ? (isDe ? `Format ok · ${neuTrim.length} Zeichen${hexartig ? ' · hexadezimal (Outlook)' : ''}` : `Format ok · ${neuTrim.length} chars${hexartig ? ' · hexadecimal (Outlook)' : ''}`)
                  : (isDe ? `Format falsch · nur ${neuTrim.length} Zeichen` : `Bad format · only ${neuTrim.length} chars`)}
              </span>
              {aktuell && (
                <span className={`dex-ui-pill ${gleichWieHinterlegt ? 'dex-ui-pill--gray' : aktuell.calendarLink ? 'dex-ui-pill--blue' : 'dex-ui-pill--green'}`}>
                  {gleichWieHinterlegt
                    ? (isDe ? 'identisch mit dem hinterlegten Wert' : 'identical to the stored value')
                    : aktuell.calendarLink
                      ? (isDe ? 'anderer Wert als hinterlegt' : 'differs from the stored value')
                      : (isDe ? 'bisher nichts hinterlegt' : 'nothing stored so far')}
                </span>
              )}
              {bereinigt && <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'Leerzeichen/Umbrüche entfernt' : 'whitespace/line breaks removed'}</span>}
            </div>
          )}
          <div className="dex-ui-inline" style={{ gap: 8, flexWrap: 'wrap' }}>
            <button type="button" className="btn btn-secondary" disabled={!siehtAusWieUid(neuTrim) || pruefBusy || busy} onClick={() => { void pruefen(); }}>
              {pruefBusy ? (isDe ? 'Wird geprüft …' : 'Checking …') : (isDe ? `Im Kalender von ${NO_REPLY_MAILBOX} prüfen` : `Check in the ${NO_REPLY_MAILBOX} calendar`)}
            </button>
          </div>
          {pruefung && (
            pruefung.status === 'gefunden' ? (
              <div className="dex-ui-callout dex-ui-callout--success dex-ui-callout--sm"><span>
                {isDe ? 'Gefunden: ' : 'Found: '}<strong>{pruefung.subject || (isDe ? '(ohne Betreff)' : '(no subject)')}</strong> · {isDe ? 'Beginn' : 'Start'} {fmtDatum(pruefung.start, isDe)} · {pruefung.attendees} {isDe ? 'eingeladen' : 'invited'} · {isDe ? 'angelegt' : 'created'} {fmtDatum(pruefung.created, isDe)}
              </span></div>
            ) : (
              <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm"><span>
                {pruefung.status === 'nicht-gefunden'
                  ? (isDe ? 'Kein Termin mit dieser iCalUId im Kalender der Shared Mailbox. Überschreiben geht trotzdem — dann aber auf eigene Verantwortung.' : 'No appointment with this iCalUId in the shared mailbox calendar. You can still overwrite — at your own risk.')
                  : pruefung.status === 'kein-zugriff'
                    ? (isDe ? `Kein Zugriff auf den Kalender von ${NO_REPLY_MAILBOX} — die Prüfung ist nicht möglich, Überschreiben geht trotzdem.` : `No access to the ${NO_REPLY_MAILBOX} calendar — cannot verify, overwriting still works.`)
                    : pruefung.status === 'kein-graph'
                      ? (isDe ? 'Graph ist hier nicht verfügbar — Prüfung nicht möglich.' : 'Graph is not available here — cannot verify.')
                      : (isDe ? `Prüfung fehlgeschlagen: ${pruefung.message || 'unbekannter Fehler'}` : `Check failed: ${pruefung.message || 'unknown error'}`)}
              </span></div>
            )
          )}
          {aktuell && aktuell.calendarLink && (
            <label className={`dex-ui-toggle-row${absagen ? ' is-active' : ''}${busy ? ' is-disabled' : ''}`}>
              <input type="checkbox" checked={absagen} disabled={busy} onChange={ev => setAbsagen(ev.target.checked)} />
              <span style={{ fontSize: '0.84rem' }}>
                {isDe
                  ? <>Den <strong>bisher verknüpften</strong> Termin absagen — die Absage geht nur an die, die auf ihm stehen. Standard aus: Wer hier von Hand einen Wert einträgt, will meist nur die Verknüpfung richten.</>
                  : <><strong>Cancel the previously linked</strong> appointment — only its attendees get the cancellation. Off by default: entering a value by hand usually just means fixing the link.</>}
              </span>
            </label>
          )}
          {ergebnis && (() => { const t = ergebnisText(ergebnis); return (
            <div className={`dex-ui-callout dex-ui-callout--sm ${t.fehler ? 'dex-ui-callout--danger' : 'dex-ui-callout--success'}`}><span>{t.text}</span></div>
          ); })()}
          <div>
            <button type="button" className="btn btn-primary" disabled={!kannUebernehmen} onClick={() => { void uebernehmen(); }}>
              {busy ? (isDe ? 'Wird geschrieben …' : 'Writing …') : (isDe ? 'CalendarLink überschreiben' : 'Overwrite CalendarLink')}
            </button>
          </div>
        </div>
      </div>
    </Modal>
  );
};
