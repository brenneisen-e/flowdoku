/**
 * v31.49: Dialog zu „Gelöschte Termine im Papierkorb suchen" — listet die
 * gefundenen DEX_Events-Zeilen und holt sie auf Klick zurück.
 * v31.50: Zweiter Weg im selben Dialog — Termine aus dem Kalender der Shared
 * Mailbox lesen und den richtigen per Klick verknüpfen (wenn die alte Zeile
 * nicht mehr im Papierkorb liegt). Regeln und Reihenfolge stehen in
 * `logic/recycleBinOutlook.ts`.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Trash2 } from '../../Icons';
import type { EventService } from '../../../services/EventService';
import type { DeloitteEvent } from '../../../types';
import {
  KalenderSuche, KalenderTermin, NO_REPLY_MAILBOX, PapierkorbSuche, PapierkorbTreffer, RueckholErgebnis,
  VerknuepfErgebnis, sucheKalender, verknuepfeTermin,
} from '../logic/recycleBinOutlook';

export interface RecycleBinOutlookModalProps {
  isDe: boolean;
  svc: EventService;
  event: DeloitteEvent;
  kinder: DeloitteEvent[];
  /** null = Suche läuft noch. */
  suche: PapierkorbSuche | null;
  /** binId des Eintrags, der gerade zurückgeholt wird. */
  busyId: string;
  /** Ergebnis je binId (nach dem Zurückholen). */
  ergebnisse: Record<string, RueckholErgebnis>;
  /** absagen = Nutzer-Wahl: den überzähligen Termin der neuen Zeile absagen (true) oder behalten. */
  onZurueckholen: (t: PapierkorbTreffer, absagen: boolean) => void;
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

const NEU_DE: Record<string, string> = {
  eingereiht: 'Die Absage des überzähligen Termins ist eingereiht.',
  behalten: 'Der überzählige Termin bleibt auf deinen Wunsch stehen — er hängt an keiner DEX-Zeile mehr und muss, wenn er weg soll, in Outlook von Hand abgesagt werden.',
  keiner: 'Es gab keinen zweiten Termin abzusagen.',
  fehlgeschlagen: 'ACHTUNG: Die Absage des überzähligen Termins konnte nicht eingereiht werden — bitte den Termin in Outlook von Hand absagen.',
};
const NEU_EN: Record<string, string> = {
  eingereiht: 'The cancellation of the surplus appointment is queued.',
  behalten: 'As requested, the surplus appointment stays — it is no longer linked to any DEX row and must be cancelled in Outlook manually if it should go.',
  keiner: 'There was no second appointment to cancel.',
  fehlgeschlagen: 'ATTENTION: the cancellation could not be queued — please cancel the appointment in Outlook manually.',
};

export const RecycleBinOutlookModal: React.FC<RecycleBinOutlookModalProps> = (p) => {
  const { isDe, svc, event, kinder, suche, busyId, ergebnisse, onZurueckholen, confirmDialog, refreshEvents, onClose } = p;
  // Nutzer-Wahl je Treffer: überzähligen Termin absagen (Vorgabe) oder behalten.
  const [absagen, setAbsagen] = React.useState<Record<string, boolean>>({});
  const absagenFuer = (binId: string): boolean => absagen[binId] !== false;

  // v31.50: Kalender-Suche (zweiter Weg).
  const [kalKindId, setKalKindId] = React.useState<string>(kinder[0]?.id || '');
  const [kalSuche, setKalSuche] = React.useState<KalenderSuche | null>(null);
  const [kalBusy, setKalBusy] = React.useState(false);
  const [kalLinkBusy, setKalLinkBusy] = React.useState('');
  const [kalAbsagen, setKalAbsagen] = React.useState(true);
  const [kalErgebnis, setKalErgebnis] = React.useState<{ uid: string; e: VerknuepfErgebnis } | null>(null);
  const kalKind = kinder.find(k => k.id === kalKindId) || null;
  const laeuft = !!busyId || kalBusy || !!kalLinkBusy;

  const runKalSuche = async (): Promise<void> => {
    if (!kalKind || kalBusy) return;
    setKalBusy(true);
    setKalSuche(null);
    setKalErgebnis(null);
    try { setKalSuche(await sucheKalender(svc, kalKind)); } finally { setKalBusy(false); }
  };
  const runVerknuepfen = async (t: KalenderTermin): Promise<void> => {
    if (!kalKind || laeuft) return;
    const bisher = (kalKind.calendarLink || '').trim();
    const ok = await confirmDialog(
      isDe
        ? <>
          <p style={{ margin: '0 0 10px' }}>
            Der Termin <strong>&bdquo;{t.subject}&ldquo;</strong> ({fmtDatum(t.start, isDe)}, {t.attendees} {t.attendees === 1 ? 'Eingeladene:r' : 'Eingeladene'}) wird mit dem Sub-Event <strong>&bdquo;{kalKind.title}&ldquo;</strong> (Zeile {kalKind.id}) verknüpft.
          </p>
          <p style={{ margin: 0 }}>
            Ab dann treffen Einladungen, Absagen und Aktualisierungen aus DEX diesen Termin.
            {' '}{bisher
              ? (kalAbsagen
                ? 'Der bisher verknüpfte Termin wird abgesagt — die Absage geht nur an die, die auf ihm stehen.'
                : 'Der bisher verknüpfte Termin bleibt stehen, hängt aber an keiner DEX-Zeile mehr.')
              : 'Bisher war kein Termin verknüpft.'}
          </p>
        </>
        : <>
          <p style={{ margin: '0 0 10px' }}>
            The appointment <strong>&ldquo;{t.subject}&rdquo;</strong> ({fmtDatum(t.start, isDe)}, {t.attendees} attendee{t.attendees === 1 ? '' : 's'}) will be linked to the sub-event <strong>&ldquo;{kalKind.title}&rdquo;</strong> (row {kalKind.id}).
          </p>
          <p style={{ margin: 0 }}>
            From then on, invites, cancellations and updates from DEX hit this appointment.
            {' '}{bisher
              ? (kalAbsagen
                ? 'The previously linked appointment is cancelled — only its attendees get the cancellation.'
                : 'The previously linked appointment stays but is no longer linked to any DEX row.')
              : 'No appointment was linked before.'}
          </p>
        </>,
      { confirmLabel: isDe ? 'Verknüpfen' : 'Link', title: isDe ? 'Termin verknüpfen' : 'Link appointment' },
    );
    if (!ok) return;
    setKalLinkBusy(t.iCalUId);
    try {
      const e = await verknuepfeTermin(svc, event, kalKind, t, kalAbsagen);
      setKalErgebnis({ uid: t.iCalUId, e });
      if (e.status === 'fertig') {
        try { await refreshEvents(); } catch { /* Anzeige zieht beim nächsten Laden nach */ }
        // Liste neu bewerten: der verknüpfte Termin trägt jetzt die Marke.
        setKalSuche(prev => prev && prev.status === 'ok'
          ? { ...prev, termine: prev.termine.map(x => ({ ...x, verknuepft: x.iCalUId === t.iCalUId })) }
          : prev);
      }
    } finally {
      setKalLinkBusy('');
    }
  };

  const ergebnisText = (e: RueckholErgebnis): { text: string; fehler: boolean } => {
    const de: Record<string, string> = {
      'fertig': `Zurückgeholt. Zeile ${e.oldItemId} ist wieder das Sub-Event, der Termin (${e.calendarLink.slice(0, 12)}…) hängt wieder daran; Zeile ${e.newItemId} ist entfernt. ${NEU_DE[e.neuerTermin]}`,
      'restore-fehlgeschlagen': 'SharePoint hat das Wiederherstellen abgelehnt — es wurde nichts verändert. Bitte gleich noch einmal versuchen.',
      'alte-zeile-nicht-lesbar': `Die Zeile ${e.oldItemId} wurde wiederhergestellt, ist aber nicht lesbar — es wurde nichts ersetzt. Bitte in der Liste DEX_Events nachsehen.`,
      'falsches-event': `Die wiederhergestellte Zeile gehört zu einem anderen Event (ParentEventId passt nicht) — ${e.zurueckgenommen ? 'sie liegt wieder im Papierkorb.' : 'ACHTUNG: sie konnte nicht zurückgelegt werden und steht jetzt in DEX_Events.'}`,
      'andere-subsite': `Die wiederhergestellte Zeile zeigt auf eine andere Teilnehmer-Subsite als das aktuelle Sub-Event — ${e.zurueckgenommen ? 'sie liegt wieder im Papierkorb.' : 'ACHTUNG: sie konnte nicht zurückgelegt werden und steht jetzt in DEX_Events.'}`,
      'kein-termin': `Die alte Zeile trägt keinen Outlook-Termin (CalendarLink leer) — es gibt nichts zu übernehmen; ${e.zurueckgenommen ? 'sie liegt wieder im Papierkorb.' : 'ACHTUNG: sie konnte nicht zurückgelegt werden und steht jetzt in DEX_Events.'}`,
      'ziel-nicht-lesbar': `Das aktuelle Sub-Event war nicht lesbar — ${e.zurueckgenommen ? 'die alte Zeile liegt wieder im Papierkorb, nichts verändert.' : 'ACHTUNG: die alte Zeile konnte nicht zurückgelegt werden und steht jetzt zusätzlich in DEX_Events.'}`,
      'neue-zeile-bleibt': `Die alte Zeile ${e.oldItemId} ist zurück, aber die neue Zeile ${e.newItemId} ließ sich nicht entfernen — jetzt stehen ZWEI Zeilen auf derselben Teilnehmerliste. Bitte melde dich beim DEX-Team; lösche keine der beiden selbst über DEX.`,
    };
    const en: Record<string, string> = {
      'fertig': `Restored. Row ${e.oldItemId} is the sub-event again with its appointment (${e.calendarLink.slice(0, 12)}…); row ${e.newItemId} has been removed. ${NEU_EN[e.neuerTermin]}`,
      'restore-fehlgeschlagen': 'SharePoint rejected the restore — nothing was changed. Please try again in a moment.',
      'alte-zeile-nicht-lesbar': `Row ${e.oldItemId} was restored but could not be read — nothing was replaced. Please check the DEX_Events list.`,
      'falsches-event': `The restored row belongs to another event (ParentEventId mismatch) — ${e.zurueckgenommen ? 'it is back in the recycle bin.' : 'ATTENTION: it could not be put back and is now in DEX_Events.'}`,
      'andere-subsite': `The restored row points to a different participant subsite than the current sub-event — ${e.zurueckgenommen ? 'it is back in the recycle bin.' : 'ATTENTION: it could not be put back and is now in DEX_Events.'}`,
      'kein-termin': `The old row has no Outlook appointment (CalendarLink empty) — nothing to take over; ${e.zurueckgenommen ? 'it is back in the recycle bin.' : 'ATTENTION: it could not be put back and is now in DEX_Events.'}`,
      'ziel-nicht-lesbar': `The current sub-event could not be read — ${e.zurueckgenommen ? 'the old row is back in the recycle bin, nothing changed.' : 'ATTENTION: the old row could not be put back and is now an extra row in DEX_Events.'}`,
      'neue-zeile-bleibt': `Old row ${e.oldItemId} is back, but new row ${e.newItemId} could not be removed — TWO rows now share the same participant list. Please contact the DEX team; do not delete either of them via DEX.`,
    };
    const fehler = e.status !== 'fertig' || e.neuerTermin === 'fehlgeschlagen';
    return { text: (isDe ? de : en)[e.status] || e.status, fehler };
  };

  const verknuepfText = (e: VerknuepfErgebnis): { text: string; fehler: boolean } => {
    const de: Record<string, string> = {
      'fertig': `Verknüpft. ${NEU_DE[e.bisheriger]} Öffne das Event einmal neu im Assistenten, damit die Anzeige den Termin kennt.`,
      'schon-verknuepft': 'Dieser Termin ist bereits verknüpft — nichts zu tun.',
      'zeile-nicht-lesbar': 'Die Zeile des Sub-Events war nicht lesbar — es wurde nichts geschrieben.',
      'schreiben-fehlgeschlagen': 'SharePoint hat das Schreiben abgelehnt — es wurde nichts verändert. Bitte gleich noch einmal versuchen.',
    };
    const en: Record<string, string> = {
      'fertig': `Linked. ${NEU_EN[e.bisheriger]} Open the event in the wizard once so the display picks up the appointment.`,
      'schon-verknuepft': 'This appointment is already linked — nothing to do.',
      'zeile-nicht-lesbar': 'The sub-event row could not be read — nothing was written.',
      'schreiben-fehlgeschlagen': 'SharePoint rejected the write — nothing was changed. Please try again in a moment.',
    };
    return { text: (isDe ? de : en)[e.status] || e.status, fehler: e.status !== 'fertig' || e.bisheriger === 'fehlgeschlagen' };
  };

  const h3: React.CSSProperties = { margin: '0 0 6px', fontSize: '0.8rem', letterSpacing: '0.04em', textTransform: 'uppercase', color: 'var(--dex-gray-600, #666)' };

  return (
    <Modal
      open={true}
      onClose={() => { if (!laeuft) onClose(); }}
      dismissable={!laeuft}
      maxWidth={680}
      ariaLabel={isDe ? 'Outlook-Termin wiederfinden' : 'Recover the Outlook appointment'}
      title={isDe ? 'Outlook-Termin wiederfinden' : 'Recover the Outlook appointment'}
      subtitle={isDe
        ? <>Zwei Wege für &bdquo;{event.title}&ldquo;: die gelöschte Zeile aus dem Papierkorb zurückholen — oder den richtigen Termin im Kalender von {NO_REPLY_MAILBOX} suchen und verknüpfen.</>
        : <>Two ways for &ldquo;{event.title}&rdquo;: restore the deleted row from the recycle bin — or find the right appointment in the {NO_REPLY_MAILBOX} calendar and link it.</>}
      icon={<Trash2 size={20} />}
      footer={(
        <button type="button" className="btn btn-secondary" disabled={laeuft} onClick={onClose}>
          {isDe ? 'Schließen' : 'Close'}
        </button>
      )}
    >
      <div className="dex-ui-stack" style={{ gap: 18 }}>
        {/* ---------- Weg 1: Papierkorb ---------- */}
        <section>
          <h3 style={h3}>{isDe ? '1 · Papierkorb der Site' : '1 · Site recycle bin'}</h3>
          {!suche ? (
            <div className="dex-ui-empty">{isDe ? 'Papierkorb wird gelesen …' : 'Reading the recycle bin …'}</div>
          ) : suche.status === 'nicht-lesbar' ? (
            <div className="dex-ui-callout dex-ui-callout--danger"><span>
              {isDe
                ? 'Der Papierkorb der Site war nicht lesbar — ob dort etwas liegt, ist unbekannt. Bitte gleich noch einmal versuchen; ohne Leserecht auf den Papierkorb (Site Owner) geht es nicht.'
                : 'The site recycle bin could not be read — whether anything is there is unknown. Please try again; reading the recycle bin requires site owner rights.'}
            </span></div>
          ) : (
            <div className="dex-ui-stack">
              {!suche.zweiteStufeGelesen && (
                <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm"><span>
                  {isDe
                    ? 'Nur die erste Stufe des Papierkorbs war lesbar; die zweite Stufe (Site Collection) nicht.'
                    : 'Only the first-stage recycle bin was readable; the second stage (site collection) was not.'}
                </span></div>
              )}
              {suche.treffer.length === 0 ? (
                <div className="dex-ui-callout dex-ui-callout--neutral"><span>
                  {isDe
                    ? 'Keine gelöschte Zeile mit passendem Titel im Papierkorb — die alte Zeile wurde endgültig gelöscht. Dann weiter mit Weg 2.'
                    : 'No deleted row with a matching title in the recycle bin — the old row was deleted for good. Continue with way 2.'}
                </span></div>
              ) : suche.treffer.map(t => {
                const e = ergebnisse[t.binId];
                const et = e ? ergebnisText(e) : null;
                const kann = t.zielGrund === 'ok' && !e;
                const grund = t.zielGrund === 'mehrdeutig'
                  ? (isDe ? 'Mehrere Sub-Events tragen diesen Titel — Ziel nicht eindeutig.' : 'Several sub-events carry this title — target is ambiguous.')
                  : t.zielGrund === 'keine-id'
                    ? (isDe ? 'Item-Id der gelöschten Zeile nicht lesbar.' : 'Item id of the deleted row not readable.')
                    : t.zielGrund === 'kein-subevent'
                      ? (isDe ? 'Kein aktuelles Sub-Event mit diesem Titel.' : 'No current sub-event with this title.')
                      : '';
                return (
                  <div key={t.binId} className="dex-ui-card dex-ui-card--soft dex-ui-stack">
                    <div className="dex-ui-inline" style={{ justifyContent: 'space-between', gap: 8 }}>
                      <strong style={{ fontSize: '0.92rem', color: 'var(--dex-gray-800)', minWidth: 0 }}>{t.title}</strong>
                      <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'Zeile' : 'Row'} {t.oldItemId ?? '?'}</span>
                    </div>
                    <div className="dex-ui-muted" style={{ fontSize: '0.82rem' }}>
                      {isDe ? 'Gelöscht' : 'Deleted'} {fmtDatum(t.deletedDate, isDe)}{t.deletedBy ? ` · ${t.deletedBy}` : ''}
                      {' · '}{t.stufe === 'web' ? (isDe ? '1. Stufe' : '1st stage') : (isDe ? '2. Stufe' : '2nd stage')}
                    </div>
                    {t.ziel && (
                      <div style={{ fontSize: '0.84rem', color: 'var(--dex-gray-800)' }}>
                        {isDe ? 'Ersetzt die aktuelle Zeile' : 'Replaces current row'} <strong>{t.ziel.id}</strong> (&bdquo;{t.ziel.title}&ldquo;).
                        {' '}{isDe ? 'Änderungen, die mit dem Neuanlegen an dieser Zeile gespeichert wurden, gehen dabei verloren.' : 'Changes saved to that row when it was recreated are lost.'}
                      </div>
                    )}
                    {grund && <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm"><span>{grund}</span></div>}
                    {et && (
                      <div className={`dex-ui-callout dex-ui-callout--sm ${et.fehler ? 'dex-ui-callout--danger' : 'dex-ui-callout--success'}`}><span>{et.text}</span></div>
                    )}
                    {!e && kann && (
                      <label className={`dex-ui-toggle-row${absagenFuer(t.binId) ? ' is-active' : ''}${laeuft ? ' is-disabled' : ''}`}>
                        <input
                          type="checkbox"
                          checked={absagenFuer(t.binId)}
                          disabled={laeuft}
                          onChange={ev => setAbsagen(prev => ({ ...prev, [t.binId]: ev.target.checked }))}
                        />
                        <span style={{ fontSize: '0.84rem' }}>
                          {isDe
                            ? <>Überzähligen Termin der neuen Zeile <strong>absagen</strong> — die Absage geht nur an die, die auf diesem Termin stehen. Ohne Haken bleibt er in deren Kalendern stehen, hängt aber an keiner DEX-Zeile mehr.</>
                            : <><strong>Cancel</strong> the surplus appointment of the new row — the cancellation only reaches its attendees. Unticked, it stays in their calendars but is no longer linked to any DEX row.</>}
                        </span>
                      </label>
                    )}
                    {!e && (
                      <div>
                        <button type="button" className="btn btn-primary" disabled={!kann || laeuft} onClick={() => onZurueckholen(t, absagenFuer(t.binId))}>
                          {busyId === t.binId
                            ? (isDe ? 'Wird zurückgeholt …' : 'Restoring …')
                            : (isDe ? 'Zurückholen und Termin übernehmen' : 'Restore and take over the appointment')}
                        </button>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </section>

        {/* ---------- Weg 2: Kalender der Shared Mailbox ---------- */}
        <section>
          <h3 style={h3}>{isDe ? `2 · Kalender von ${NO_REPLY_MAILBOX}` : `2 · Calendar of ${NO_REPLY_MAILBOX}`}</h3>
          <div className="dex-ui-stack">
            <p className="dex-ui-muted" style={{ margin: 0, fontSize: '0.84rem' }}>
              {isDe
                ? 'DEX liest die Termine rund um den Start des gewählten Sub-Events aus dem Kalender der Shared Mailbox (du brauchst Zugriff auf das Postfach — den hast du, wenn du es in Outlook siehst). Der richtige Termin ist der mit den vielen Eingeladenen; „verknüpft" markiert den, an dem die Zeile gerade hängt.'
                : 'DEX reads the appointments around the start of the chosen sub-event from the shared mailbox calendar (you need access to the mailbox — you have it if you see it in Outlook). The right one is the appointment with the many attendees; “linked” marks the one the row currently points to.'}
            </p>
            {kinder.length === 0 ? (
              <div className="dex-ui-empty">{isDe ? 'Dieses Event hat keine Sub-Events.' : 'This event has no sub-events.'}</div>
            ) : (
              <div className="dex-ui-inline" style={{ gap: 8, flexWrap: 'wrap' }}>
                <select
                  value={kalKindId}
                  disabled={laeuft}
                  onChange={ev => { setKalKindId(ev.target.value); setKalSuche(null); setKalErgebnis(null); }}
                  style={{ flex: '1 1 260px', minWidth: 0, padding: '8px 10px', borderRadius: 8, border: '1px solid var(--dex-gray-300, #ccc)', fontSize: '0.88rem' }}
                >
                  {kinder.map(k => (
                    <option key={k.id} value={k.id}>{k.title || (isDe ? 'ohne Titel' : 'untitled')} · {isDe ? 'Zeile' : 'row'} {k.id}{(k.calendarLink || '').trim() ? '' : (isDe ? ' · kein Termin verknüpft' : ' · no appointment linked')}</option>
                  ))}
                </select>
                <button type="button" className="btn btn-primary" disabled={!kalKind || laeuft} onClick={() => { void runKalSuche(); }}>
                  {kalBusy ? (isDe ? 'Kalender wird gelesen …' : 'Reading calendar …') : (isDe ? 'Termine im Kalender suchen' : 'Find appointments in the calendar')}
                </button>
              </div>
            )}
            {kalSuche && kalSuche.status !== 'ok' && (
              <div className="dex-ui-callout dex-ui-callout--danger"><span>
                {kalSuche.status === 'kein-graph'
                  ? (isDe ? 'Graph-Zugriff ist in dieser Umgebung nicht verfügbar.' : 'Graph access is not available in this environment.')
                  : kalSuche.status === 'kein-zugriff'
                    ? (isDe ? `Kein Zugriff auf den Kalender von ${NO_REPLY_MAILBOX} (401/403). Dein Konto braucht Vollzugriff auf das Postfach, und die App die Berechtigung „Calendars.Read.Shared".` : `No access to the ${NO_REPLY_MAILBOX} calendar (401/403). Your account needs full access to the mailbox, and the app the “Calendars.Read.Shared” permission.`)
                    : (isDe ? `Kalender konnte nicht gelesen werden: ${kalSuche.message || 'unbekannter Fehler'}` : `Calendar could not be read: ${kalSuche.message || 'unknown error'}`)}
              </span></div>
            )}
            {kalSuche && kalSuche.status === 'ok' && (
              kalSuche.termine.length === 0 ? (
                <div className="dex-ui-callout dex-ui-callout--neutral"><span>
                  {isDe
                    ? `Keine Termine ${kalSuche.fenster ? `im Zeitraum ${kalSuche.fenster}` : 'mit diesem Betreff'} im Kalender.`
                    : `No appointments ${kalSuche.fenster ? `in the range ${kalSuche.fenster}` : 'with this subject'} in the calendar.`}
                </span></div>
              ) : (
                <>
                  {kalSuche.fenster && (
                    <div className="dex-ui-muted" style={{ fontSize: '0.8rem' }}>
                      {isDe ? `Zeitraum ${kalSuche.fenster} · ${kalSuche.termine.length} Termine` : `Range ${kalSuche.fenster} · ${kalSuche.termine.length} appointments`}
                    </div>
                  )}
                  {kalKind && (
                    <label className={`dex-ui-toggle-row${kalAbsagen ? ' is-active' : ''}${laeuft ? ' is-disabled' : ''}`}>
                      <input type="checkbox" checked={kalAbsagen} disabled={laeuft} onChange={ev => setKalAbsagen(ev.target.checked)} />
                      <span style={{ fontSize: '0.84rem' }}>
                        {isDe
                          ? <>Beim Verknüpfen den <strong>bisher verknüpften</strong> Termin absagen — die Absage geht nur an die, die auf ihm stehen. Ohne Haken bleibt er stehen, hängt aber an keiner DEX-Zeile mehr.</>
                          : <>When linking, <strong>cancel the previously linked</strong> appointment — the cancellation only reaches its attendees. Unticked, it stays but is no longer linked to any DEX row.</>}
                      </span>
                    </label>
                  )}
                  {kalSuche.termine.map(t => {
                    const erg = kalErgebnis && kalErgebnis.uid === t.iCalUId ? verknuepfText(kalErgebnis.e) : null;
                    return (
                      <div key={t.graphId || t.iCalUId} className="dex-ui-card dex-ui-card--soft dex-ui-stack" style={t.passt ? undefined : { opacity: 0.75 }}>
                        <div className="dex-ui-inline" style={{ justifyContent: 'space-between', gap: 8, flexWrap: 'wrap' }}>
                          <strong style={{ fontSize: '0.92rem', color: 'var(--dex-gray-800)', minWidth: 0 }}>{t.subject || (isDe ? '(ohne Betreff)' : '(no subject)')}</strong>
                          <span className="dex-ui-inline" style={{ gap: 6 }}>
                            <span className={`dex-ui-pill ${t.attendees >= 10 ? 'dex-ui-pill--green' : 'dex-ui-pill--gray'}`}>{t.attendees} {isDe ? 'eingeladen' : 'invited'}</span>
                            {t.verknuepft && <span className="dex-ui-pill dex-ui-pill--blue">{isDe ? 'verknüpft' : 'linked'}</span>}
                            {!t.passt && <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'anderer Betreff' : 'other subject'}</span>}
                          </span>
                        </div>
                        <div className="dex-ui-muted" style={{ fontSize: '0.82rem' }}>
                          {isDe ? 'Beginn' : 'Start'} {fmtDatum(t.start, isDe)} · {isDe ? 'angelegt' : 'created'} {fmtDatum(t.created, isDe)}
                        </div>
                        <div style={{ fontSize: '0.76rem', fontFamily: 'monospace', color: 'var(--dex-gray-700, #555)', wordBreak: 'break-all' }}>
                          iCalUId: {t.iCalUId}
                        </div>
                        {erg && (
                          <div className={`dex-ui-callout dex-ui-callout--sm ${erg.fehler ? 'dex-ui-callout--danger' : 'dex-ui-callout--success'}`}><span>{erg.text}</span></div>
                        )}
                        {!t.verknuepft && (
                          <div>
                            <button type="button" className="btn btn-primary" disabled={laeuft} onClick={() => { void runVerknuepfen(t); }}>
                              {kalLinkBusy === t.iCalUId
                                ? (isDe ? 'Wird verknüpft …' : 'Linking …')
                                : (isDe ? 'Diesen Termin verknüpfen' : 'Link this appointment')}
                            </button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </>
              )
            )}
          </div>
        </section>
      </div>
    </Modal>
  );
};
