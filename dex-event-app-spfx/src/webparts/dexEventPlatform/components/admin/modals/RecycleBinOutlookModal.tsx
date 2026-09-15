/**
 * v31.49: Dialog zu „Gelöschte Termine im Papierkorb suchen" — listet die
 * gefundenen DEX_Events-Zeilen und holt sie auf Klick zurück. Regeln und
 * Reihenfolge stehen in `logic/recycleBinOutlook.ts`.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Trash2 } from '../../Icons';
import type { PapierkorbSuche, PapierkorbTreffer, RueckholErgebnis } from '../logic/recycleBinOutlook';

export interface RecycleBinOutlookModalProps {
  isDe: boolean;
  eventTitle: string;
  /** null = Suche läuft noch. */
  suche: PapierkorbSuche | null;
  /** binId des Eintrags, der gerade zurückgeholt wird. */
  busyId: string;
  /** Ergebnis je binId (nach dem Zurückholen). */
  ergebnisse: Record<string, RueckholErgebnis>;
  /** absagen = Nutzer-Wahl: den überzähligen Termin der neuen Zeile absagen (true) oder behalten. */
  onZurueckholen: (t: PapierkorbTreffer, absagen: boolean) => void;
  onClose: () => void;
}

const fmtDatum = (iso: string, isDe: boolean): string => {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
};

export const RecycleBinOutlookModal: React.FC<RecycleBinOutlookModalProps> = (p) => {
  const { isDe, eventTitle, suche, busyId, ergebnisse, onZurueckholen, onClose } = p;
  const laeuft = !!busyId;
  // Nutzer-Wahl je Treffer: überzähligen Termin absagen (Vorgabe) oder behalten.
  const [absagen, setAbsagen] = React.useState<Record<string, boolean>>({});
  const absagenFuer = (binId: string): boolean => absagen[binId] !== false;

  const ergebnisText = (e: RueckholErgebnis): { text: string; fehler: boolean } => {
    const neuDe: Record<string, string> = {
      eingereiht: 'Die Absage des überzähligen Termins ist eingereiht.',
      behalten: 'Der überzählige Termin bleibt auf deinen Wunsch stehen — er hängt an keiner DEX-Zeile mehr und muss, wenn er weg soll, in Outlook von Hand abgesagt werden.',
      keiner: 'Die neue Zeile hatte keinen eigenen Termin.',
      fehlgeschlagen: 'ACHTUNG: Die Absage des überzähligen Termins konnte nicht eingereiht werden — bitte den neuen Termin in Outlook von Hand absagen.',
    };
    const neuEn: Record<string, string> = {
      eingereiht: 'The cancellation of the surplus appointment is queued.',
      behalten: 'As requested, the surplus appointment stays — it is no longer linked to any DEX row and must be cancelled in Outlook manually if it should go.',
      keiner: 'The new row had no appointment of its own.',
      fehlgeschlagen: 'ATTENTION: the cancellation could not be queued — please cancel the new appointment in Outlook manually.',
    };
    const de: Record<string, string> = {
      'fertig': `Zurückgeholt. Zeile ${e.oldItemId} ist wieder das Sub-Event, der Termin (${e.calendarLink.slice(0, 12)}…) hängt wieder daran; Zeile ${e.newItemId} ist entfernt. ${neuDe[e.neuerTermin]}`,
      'restore-fehlgeschlagen': 'SharePoint hat das Wiederherstellen abgelehnt — es wurde nichts verändert. Bitte gleich noch einmal versuchen.',
      'alte-zeile-nicht-lesbar': `Die Zeile ${e.oldItemId} wurde wiederhergestellt, ist aber nicht lesbar — es wurde nichts ersetzt. Bitte in der Liste DEX_Events nachsehen.`,
      'falsches-event': `Die wiederhergestellte Zeile gehört zu einem anderen Event (ParentEventId passt nicht) — ${e.zurueckgenommen ? 'sie liegt wieder im Papierkorb.' : 'ACHTUNG: sie konnte nicht zurückgelegt werden und steht jetzt in DEX_Events.'}`,
      'andere-subsite': `Die wiederhergestellte Zeile zeigt auf eine andere Teilnehmer-Subsite als das aktuelle Sub-Event — ${e.zurueckgenommen ? 'sie liegt wieder im Papierkorb.' : 'ACHTUNG: sie konnte nicht zurückgelegt werden und steht jetzt in DEX_Events.'}`,
      'kein-termin': `Die alte Zeile trägt keinen Outlook-Termin (CalendarLink leer) — es gibt nichts zu übernehmen; ${e.zurueckgenommen ? 'sie liegt wieder im Papierkorb.' : 'ACHTUNG: sie konnte nicht zurückgelegt werden und steht jetzt in DEX_Events.'}`,
      'ziel-nicht-lesbar': `Das aktuelle Sub-Event war nicht lesbar — ${e.zurueckgenommen ? 'die alte Zeile liegt wieder im Papierkorb, nichts verändert.' : 'ACHTUNG: die alte Zeile konnte nicht zurückgelegt werden und steht jetzt zusätzlich in DEX_Events.'}`,
      'neue-zeile-bleibt': `Die alte Zeile ${e.oldItemId} ist zurück, aber die neue Zeile ${e.newItemId} ließ sich nicht entfernen — jetzt stehen ZWEI Zeilen auf derselben Teilnehmerliste. Bitte melde dich beim DEX-Team; lösche keine der beiden selbst über DEX.`,
    };
    const en: Record<string, string> = {
      'fertig': `Restored. Row ${e.oldItemId} is the sub-event again with its appointment (${e.calendarLink.slice(0, 12)}…); row ${e.newItemId} has been removed. ${neuEn[e.neuerTermin]}`,
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

  return (
    <Modal
      open={true}
      onClose={() => { if (!laeuft) onClose(); }}
      dismissable={!laeuft}
      maxWidth={640}
      ariaLabel={isDe ? 'Gelöschte Termine im Papierkorb' : 'Deleted dates in the recycle bin'}
      title={isDe ? 'Gelöschte Termine im Papierkorb' : 'Deleted dates in the recycle bin'}
      subtitle={isDe
        ? <>Gelöschte DEX_Events-Zeilen, deren Titel zu einem Sub-Event von &bdquo;{eventTitle}&ldquo; passt. Zurückholen stellt die alte Zeile samt Outlook-Termin wieder her und entfernt die neue.</>
        : <>Deleted DEX_Events rows whose title matches a sub-event of &ldquo;{eventTitle}&rdquo;. Restoring brings the old row back with its Outlook appointment and removes the new one.</>}
      icon={<Trash2 size={20} />}
      footer={(
        <button type="button" className="btn btn-secondary" disabled={laeuft} onClick={onClose}>
          {isDe ? 'Schließen' : 'Close'}
        </button>
      )}
    >
      {!suche ? (
        <div className="dex-ui-empty">{isDe ? 'Papierkorb wird gelesen …' : 'Reading the recycle bin …'}</div>
      ) : suche.status === 'nicht-lesbar' ? (
        <div className="dex-ui-callout dex-ui-callout--danger">
          {isDe
            ? 'Der Papierkorb der Site war nicht lesbar — ob dort etwas liegt, ist unbekannt. Bitte gleich noch einmal versuchen; ohne Leserecht auf den Papierkorb (Site Owner) geht es nicht.'
            : 'The site recycle bin could not be read — whether anything is there is unknown. Please try again; reading the recycle bin requires site owner rights.'}
        </div>
      ) : (
        <div className="dex-ui-stack">
          {!suche.zweiteStufeGelesen && (
            <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm">
              {isDe
                ? 'Nur die erste Stufe des Papierkorbs war lesbar; die zweite Stufe (Site Collection) nicht.'
                : 'Only the first-stage recycle bin was readable; the second stage (site collection) was not.'}
            </div>
          )}
          {suche.treffer.length === 0 ? (
            <div className="dex-ui-callout dex-ui-callout--neutral">
              {isDe
                ? <>Keine gelöschte Zeile mit einem passenden Titel im Papierkorb. Dann bleibt der zweite Weg: Die iCalUId des alten Termins aus der Run history von <strong>DEX_Outlook_Einladungen</strong> (Action <strong>Get Event Details</strong>, Feld <strong>CalendarLink</strong>) in die aktuelle Zeile eintragen.</>
                : <>No deleted row with a matching title in the recycle bin. The alternative: take the old appointment&rsquo;s iCalUId from the run history of <strong>DEX_Outlook_Einladungen</strong> (action <strong>Get Event Details</strong>, field <strong>CalendarLink</strong>) and put it into the current row&rsquo;s CalendarLink.</>}
            </div>
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
                {grund && <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm">{grund}</div>}
                {et && (
                  <div className={`dex-ui-callout dex-ui-callout--sm ${et.fehler ? 'dex-ui-callout--danger' : 'dex-ui-callout--success'}`}>{et.text}</div>
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
    </Modal>
  );
};
