/**
 * v31.4: „QR-Nummern nachtragen" — die gedruckten Nummern der bereits
 * verschickten QR-Mails in die Spalte `QrSentId` zurückschreiben.
 *
 * Warum es das gibt: Die Spalte ist neu. Für jedes Event, dessen QR-Mails
 * vorher rausgingen, steht nirgends, welche Nummer in welcher Mail stand —
 * und genau dort tut der Fehler weh, weil dort schon Abmeldungen die
 * `TeilnehmerID` verschoben haben. Die Nummern stehen aber noch in der
 * Mail-Warteschlange `DEX_Emails` (der Body enthält das komplette HTML, die
 * Zeilen werden erst rund einen Monat nach dem Event archiviert).
 *
 * Drei Dinge, die diesen Dialog belastbar machen:
 *  - **Ein Lesefehler ist keine Null.** Bei einem HTTP-Fehler bricht der Lauf
 *    ab und nennt den Status. Und weil `DEX_Emails` zeilenweise gesichert ist
 *    (`ReadSecurity=2`), heißt „0 Mails" auch ohne Fehler nicht „es gab
 *    keine" — die Meldung nennt deshalb IMMER beide Ursachen.
 *  - **Vorschau vor jedem Schreibvorgang.** Wer keinen Fund hat, steht
 *    namentlich da: für diese Personen bleibt der Check-in bei der laufenden
 *    Nummer, und das darf nicht unter den Tisch fallen.
 *  - **Die Klammer ist nie der ganze Pfad.** Gelesen und geschrieben wird über
 *    das Event UND jeden Termin; `EventId` in `DEX_Emails` ist die Id des
 *    Events, zu dem die Mail gehört.
 */
import * as React from 'react';
import Modal from '../Modal';
import { DeloitteEvent } from '../../types';
import { EventService } from '../../services/EventService';
import { QrMailHit } from '../../services/events/qrSentBackfill';
import { useEvents } from '../../context/EventContext';
import { SAMPLE_QR_ID } from './adminConstants';
import { cx } from '../dexUi';
import { AlertCircle, QrCode } from '../Icons';

/** Nur diese Zeilen bekommen eine QR-Mail — Abgemeldete und Wartelistler nicht. */
const BACKFILL_STATI = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'No-Show'];

type Verdict = 'diff' | 'same' | 'organizer' | 'none';

interface BackfillRow {
  key: string;
  eventTitle: string;
  subsiteUrl: string;
  regId: number;
  name: string;
  email: string;
  /** Die heutige laufende Nummer. */
  current: number | null;
  /** Die Nummer aus der Mail — `null` = kein Fund. */
  qrId: number | null;
  verdict: Verdict;
  /** Nur „Organizer — bitte prüfen"-Zeilen sind abwählbar (s. unten). */
  include: boolean;
}

/** Dreistellig wie in der Mail. padStart gibt es im ES5-Target nicht. */
function pad3(n: number): string { let s = String(n); while (s.length < 3) s = `0${s}`; return s; }

export default function QrSentIdBackfillModal(props: {
  event: DeloitteEvent;
  childEvents: DeloitteEvent[];
  service: EventService;
  isDe: boolean;
  onClose: () => void;
  /** Nach dem Schreiben: die Teilnehmerliste der Seite neu laden. */
  onDone?: () => void;
}): React.ReactElement {
  const { event, childEvents, service, isDe, onClose, onDone } = props;
  const { getAllRegistrations } = useEvents();
  const [phase, setPhase] = React.useState<'scan' | 'preview' | 'writing' | 'done'>('scan');
  const [abortMsg, setAbortMsg] = React.useState('');
  const [rows, setRows] = React.useState<BackfillRow[]>([]);
  const [testSkipped, setTestSkipped] = React.useState(0);
  const [scannedMails, setScannedMails] = React.useState(0);
  const [unparsedMails, setUnparsedMails] = React.useState(0);
  const [progress, setProgress] = React.useState({ done: 0, total: 0 });
  const [resultMsg, setResultMsg] = React.useState('');
  const [resultIsError, setResultIsError] = React.useState(false);

  const targets = React.useMemo(
    () => [event].concat(childEvents).filter(e => !!e && !!e.subsiteUrl),
    [event, childEvents],
  );

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      const collected: BackfillRow[] = [];
      let skipped = 0;
      let scanned = 0;
      let unparsed = 0;
      for (const ev of targets) {
        // 1) Teilnehmerliste — ein Lesefehler beendet den Lauf. Eine leere
        //    Liste wegen fehlender Rechte würde sonst als „niemand angemeldet"
        //    durchgehen und die Vorschau wäre eine Aussage über nichts.
        let readErr = -1;
        const regs = await getAllRegistrations(ev.id, (s: number) => { readErr = s; });
        if (cancelled) return;
        if (readErr >= 0) {
          // Status 0 heißt „kein HTTP-Versuch möglich" (keine Subsite oder
          // Netzfehler) — „HTTP 0" wäre dafür eine irreführende Angabe.
          const why = readErr > 0 ? `HTTP ${readErr}` : (isDe ? 'keine Teilnehmerliste erreichbar' : 'attendee list not reachable');
          setAbortMsg(isDe
            ? `Die Teilnehmerliste von „${ev.title}" konnte nicht gelesen werden (${why}) — abgebrochen, es wurde nichts geschrieben. Bitte Leserechte prüfen und erneut versuchen.`
            : `The attendee list of “${ev.title}” could not be read (${why}) — aborted, nothing was written. Please check read access and try again.`);
          setPhase('done');
          setResultIsError(true);
          return;
        }
        // 2) Die verschickten QR-Mails dieses Events.
        const scan = await service.scanQrMailsForEvent(ev.id);
        if (cancelled) return;
        if (!scan.ok) {
          setAbortMsg(isDe
            ? `Die Mail-Warteschlange DEX_Emails konnte nicht gelesen werden (${scan.status ? `HTTP ${scan.status}` : 'Netzwerkfehler'}) — abgebrochen, es wurde nichts geschrieben.`
            : `The mail queue DEX_Emails could not be read (${scan.status ? `HTTP ${scan.status}` : 'network error'}) — aborted, nothing was written.`);
          setPhase('done');
          setResultIsError(true);
          return;
        }
        scanned += scan.scanned;
        unparsed += scan.unparsed;

        // 3) Organizer-Adressen dieses Events UND des Elternevents — die
        //    Test-Mail geht an die Organizer, und bei einem Termin stehen sie
        //    oft nur auf der Klammer.
        const orgSet: Record<string, boolean> = {};
        const parent = ev.parentEventId === event.id ? event : null;
        [ev, parent].forEach(e => {
          if (!e) return;
          ((e.organizerEmails || []) as string[]).concat((e.coOrganizerEmails || []) as string[])
            .forEach(a => { const k = (a || '').trim().toLowerCase(); if (k) orgSet[k] = true; });
        });

        // 4) Test-Mails aussortieren: `qrTestSendAction` schickt Mails vom Typ
        //    QRCode an die Organizer, mit der Beispiel-ID für alle außer der
        //    angemeldeten Person. Eine Mail an eine Organizer-Adresse mit genau
        //    dieser Nummer ist deshalb sicher ein Test — alles andere nicht.
        const usable: QrMailHit[] = [];
        for (const h of scan.hits) {
          if (!h.redirected && orgSet[h.recipient] && h.qrId === SAMPLE_QR_ID) { skipped++; continue; }
          usable.push(h);
        }
        // 5) Je Person gewinnt die Mail mit der HÖCHSTEN Id — ein zweiter
        //    Versand ersetzt den ersten.
        const byEmail: Record<string, QrMailHit> = {};
        for (const h of usable) {
          const k = h.email;
          if (!k) continue;
          if (!byEmail[k] || h.mailId > byEmail[k].mailId) byEmail[k] = h;
        }
        // 6) Zeilen bilden.
        for (const r of regs) {
          if (BACKFILL_STATI.indexOf(r.Status || '') < 0) continue;
          const email = (r.ParticipantEmail || '').toLowerCase().trim();
          const hit = email ? byEmail[email] : undefined;
          const cur = (r.TeilnehmerID !== undefined && r.TeilnehmerID !== null && isFinite(Number(r.TeilnehmerID)))
            ? Number(r.TeilnehmerID) : null;
          let verdict: Verdict = 'none';
          if (hit) {
            // Ist die Person selbst Organizer und ging die Mail direkt an sie,
            // lässt sich Test von echt nicht sicher trennen — dann wird die
            // Zeile markiert und der Mensch entscheidet.
            if (orgSet[email] && !hit.redirected) verdict = 'organizer';
            else verdict = (cur !== null && hit.qrId === cur) ? 'same' : 'diff';
          }
          collected.push({
            key: `${ev.id}:${r.Id}`,
            eventTitle: ev.title,
            subsiteUrl: ev.subsiteUrl as string,
            regId: r.Id,
            name: (r.Vorname && r.Nachname) ? `${r.Vorname} ${r.Nachname}` : (r.ParticipantName || r.ParticipantEmail || '-'),
            email: r.ParticipantEmail || '',
            current: cur,
            qrId: hit ? hit.qrId : null,
            verdict,
            include: !!hit,
          });
        }
      }
      if (cancelled) return;
      // Reihenfolge: erst was abweicht (dort passiert etwas), dann die zu
      // prüfenden, dann die ohne Fund, zuletzt die unveränderten.
      const rank = (v: Verdict): number => (v === 'diff' ? 0 : v === 'organizer' ? 1 : v === 'none' ? 2 : 3);
      collected.sort((a, b) => (rank(a.verdict) - rank(b.verdict)) || a.name.localeCompare(b.name));
      setRows(collected);
      setTestSkipped(skipped);
      setScannedMails(scanned);
      setUnparsedMails(unparsed);
      setPhase('preview');
    })().catch(() => {
      if (cancelled) return;
      setAbortMsg(isDe
        ? 'Unerwarteter Fehler beim Lesen — abgebrochen, es wurde nichts geschrieben.'
        : 'Unexpected error while reading — aborted, nothing was written.');
      setResultIsError(true);
      setPhase('done');
    });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const found = rows.filter(r => r.qrId !== null);
  const diffCount = rows.filter(r => r.verdict === 'diff').length;
  const sameCount = rows.filter(r => r.verdict === 'same').length;
  const orgCount = rows.filter(r => r.verdict === 'organizer').length;
  const missing = rows.filter(r => r.verdict === 'none');
  const writable = rows.filter(r => r.qrId !== null && r.include);

  const runWrite = async (): Promise<void> => {
    setPhase('writing');
    setProgress({ done: 0, total: writable.length });
    let ok = 0;
    let errors = 0;
    for (let i = 0; i < writable.length; i++) {
      const row = writable[i];
      const r = await service.setQrSentId(row.subsiteUrl, row.regId, row.qrId as number);
      if (!r.ok) {
        if (r.status === 400) {
          // Genau eine Abhilfe — und der Lauf bricht ab, statt 87-mal
          // denselben Fehler zu erzeugen.
          setResultIsError(true);
          setResultMsg(isDe
            ? `Abgebrochen bei „${row.eventTitle}": Auf dieser Teilnehmerliste fehlt die Spalte QrSentId. Bitte einmal „Spalten fixen" für dieses Event ausführen und danach erneut nachtragen. Bis dahin geschrieben: ${ok}.`
            : `Aborted at “${row.eventTitle}”: the attendee list is missing the QrSentId column. Please run “Fix columns” for this event once, then backfill again. Written so far: ${ok}.`);
          setPhase('done');
          if (onDone) onDone();
          return;
        }
        errors++;
      } else ok++;
      setProgress({ done: i + 1, total: writable.length });
    }
    setResultIsError(errors > 0);
    setResultMsg(isDe
      ? `${ok} Nummer(n) nachgetragen${errors > 0 ? `, ${errors} fehlgeschlagen` : ''}${missing.length > 0 ? ` · ${missing.length} Person(en) ohne Fund — für sie bleibt der Check-in bei der laufenden Nummer` : ''}.`
      : `${ok} number(s) backfilled${errors > 0 ? `, ${errors} failed` : ''}${missing.length > 0 ? ` · ${missing.length} person(s) without a match — check-in stays on the running number for them` : ''}.`);
    setPhase('done');
    if (onDone) onDone();
  };

  const verdictPill = (v: Verdict): React.ReactElement => {
    if (v === 'diff') return <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? 'weicht ab' : 'differs'}</span>;
    if (v === 'same') return <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'gleich' : 'same'}</span>;
    if (v === 'organizer') return <span className="dex-ui-pill dex-ui-pill--blue">{isDe ? 'Organizer — prüfen' : 'organizer — check'}</span>;
    return <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? 'kein Fund' : 'no match'}</span>;
  };

  const footer = phase === 'preview'
    ? (<>
        <button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" disabled={writable.length === 0} onClick={() => { void runWrite(); }}>
          {isDe ? `${writable.length} Nummer(n) schreiben` : `Write ${writable.length} number(s)`}
        </button>
      </>)
    : (<button type="button" className="btn btn-secondary" onClick={onClose}>{isDe ? 'Schließen' : 'Close'}</button>);

  return (
    <Modal
      open
      onClose={onClose}
      maxWidth={860}
      dismissable={phase !== 'writing'}
      ariaLabel={isDe ? 'QR-Nummern nachtragen' : 'Backfill QR numbers'}
      title={isDe ? 'QR-Nummern nachtragen' : 'Backfill QR numbers'}
      subtitle={isDe
        ? 'Holt aus den bereits verschickten QR-Mails zurück, welche Nummer bei welcher Person gedruckt stand. Danach greift am Check-in die Nummer aus der Mail — auch wenn Abmeldungen die laufende Nummer längst verschoben haben.'
        : 'Recovers from the QR emails already sent which number was printed for which person. After that, check-in resolves the number from the email — even when cancellations have long since shifted the running number.'}
      icon={<QrCode size={20} />}
      footer={footer}
    >
      {phase === 'scan' && (
        <p className="dex-ui-muted" style={{ margin: 0 }}>
          {isDe
            ? `Mail-Warteschlange und Teilnehmerlisten werden gelesen (${targets.length} Liste(n))…`
            : `Reading the mail queue and attendee lists (${targets.length} list(s))…`}
        </p>
      )}

      {phase === 'preview' && (
        <>
          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{isDe ? 'Ergebnis der Suche' : 'Search result'}</div>
            <p style={{ margin: '0 0 8px' }}>
              {isDe
                ? <><strong>{found.length}</strong> von <strong>{rows.length}</strong> angemeldeten Personen konnten aus den QR-Mails belegt werden.</>
                : <><strong>{found.length}</strong> of <strong>{rows.length}</strong> registered people could be matched from the QR emails.</>}
            </p>
            <div className="dex-ui-inline">
              <span className="dex-ui-pill dex-ui-pill--orange">{isDe ? `${diffCount} weichen ab` : `${diffCount} differ`}</span>
              <span className="dex-ui-pill dex-ui-pill--green">{isDe ? `${sameCount} gleich` : `${sameCount} same`}</span>
              {orgCount > 0 && <span className="dex-ui-pill dex-ui-pill--blue">{isDe ? `${orgCount} Organizer — prüfen` : `${orgCount} organizer — check`}</span>}
              {missing.length > 0 && <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? `${missing.length} ohne Fund` : `${missing.length} without a match`}</span>}
              {testSkipped > 0 && <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? `${testSkipped} Test-Mails übersprungen` : `${testSkipped} test emails skipped`}</span>}
            </div>
          </div>

          {found.length === 0 && (
            /* Kein Fund heißt hier NICHT „es gab keine Mails". DEX_Emails ist
               zeilenweise gesichert — wer die Zeilen nicht angelegt hat, sieht
               sie schlicht nicht, ohne dass ein Fehler entsteht. Beide Ursachen
               gehören deshalb in denselben Satz. */
            <div className="dex-ui-callout dex-ui-callout--warn">
              <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>
              <div>
                {isDe
                  ? `In der Mail-Warteschlange wurde keine QR-Mail zu diesem Event gefunden (${scannedMails} Zeile(n) gelesen). Zwei mögliche Ursachen: Es wurden noch keine QR-Mails verschickt — oder du hast keine Leserechte auf DEX_Emails. Die Liste zeigt jeder Person nur die Zeilen, die sie selbst angelegt hat; wenn der Massenversand von einem anderen Organizer lief, sieht ihn nur diese Person oder ein Site-Owner.`
                  : `No QR email for this event was found in the mail queue (${scannedMails} row(s) read). Two possible causes: no QR emails have been sent yet — or you lack read access to DEX_Emails. The list shows each person only the rows they created themselves; if the mass send was run by another organizer, only that person or a site owner can see it.`}
              </div>
            </div>
          )}

          {unparsedMails > 0 && (
            <p className="dex-ui-muted" style={{ margin: 0, fontSize: '0.76rem' }}>
              {isDe
                ? `${unparsedMails} QR-Mail(s) enthielten keine gedruckte Nummer — die stammen aus der Zeit vor v30.35, als die ID noch nicht neben dem Code stand.`
                : `${unparsedMails} QR email(s) contained no printed number — those predate v30.35, when the ID was not yet shown next to the code.`}
            </p>
          )}

          {rows.length > 0 && (
            <div className="dex-ui-table-wrap" style={{ maxHeight: 340, overflowY: 'auto' }}>
              <table className="dex-ui-table dex-ui-table--compact">
                <thead>
                  <tr>
                    <th>{isDe ? 'Person' : 'Person'}</th>
                    {targets.length > 1 && <th>{isDe ? 'Termin' : 'Date'}</th>}
                    <th className="is-num">{isDe ? 'aus der Mail' : 'from the email'}</th>
                    <th className="is-num">{isDe ? 'laufend heute' : 'running today'}</th>
                    <th>{isDe ? 'Bewertung' : 'Assessment'}</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map(r => (
                    <tr key={r.key} className={cx(r.verdict === 'none' && 'is-muted')}>
                      <td>
                        <div className="dex-ui-person-name">{r.name}</div>
                        <div className="dex-ui-person-sub">{r.email}</div>
                      </td>
                      {targets.length > 1 && <td>{r.eventTitle}</td>}
                      <td className="is-num" style={{ fontFamily: "'Courier New',Courier,monospace" }}>
                        {r.qrId !== null ? pad3(r.qrId) : '–'}
                      </td>
                      <td className="is-num" style={{ fontFamily: "'Courier New',Courier,monospace" }}>
                        {r.current !== null ? pad3(r.current) : '–'}
                      </td>
                      <td>
                        {verdictPill(r.verdict)}
                        {/* Nur die unsicheren Zeilen sind abwählbar — alles
                            andere wird geschrieben, damit `QrSentId` danach ein
                            vollständiger Datensatz ist und nicht halb gefüllt. */}
                        {r.verdict === 'organizer' && (
                          <label className="dex-ui-inline" style={{ marginTop: 4, fontSize: '0.74rem' }}>
                            <input
                              type="checkbox"
                              checked={r.include}
                              onChange={e => {
                                const on = e.target.checked;
                                setRows(prev => prev.map(x => x.key === r.key ? { ...x, include: on } : x));
                              }}
                            />
                            <span>{isDe ? 'mitschreiben' : 'include'}</span>
                          </label>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {missing.length > 0 && (
            <p className="dex-ui-muted" style={{ margin: 0, fontSize: '0.78rem' }}>
              {isDe
                ? `Für die ${missing.length} Person(en) ohne Fund bleibt der Check-in bei der laufenden Nummer — sie stehen oben namentlich in der Tabelle.`
                : `For the ${missing.length} person(s) without a match, check-in stays on the running number — they are listed by name in the table above.`}
            </p>
          )}
        </>
      )}

      {phase === 'writing' && (
        <p style={{ margin: 0 }}>
          {isDe ? `Wird geschrieben: ${progress.done} von ${progress.total}…` : `Writing: ${progress.done} of ${progress.total}…`}
        </p>
      )}

      {phase === 'done' && (
        <div className={cx('dex-ui-callout', resultIsError ? 'dex-ui-callout--warn' : 'dex-ui-callout--success')}>
          {resultIsError && <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>}
          <div>{abortMsg || resultMsg}</div>
        </div>
      )}
    </Modal>
  );
}
