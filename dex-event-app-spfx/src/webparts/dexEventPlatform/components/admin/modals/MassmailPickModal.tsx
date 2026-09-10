/* MassmailPickModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14939-15018 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { MassmailAudience } from '../adminTypes';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { Users } from '../../Icons';
import { useLocaleSafe } from '../../../context/LanguageContext';
import { SPRegistration } from '../../../services/EventService';

export interface MassmailPickModalProps {
  massmailAudience: MassmailAudience;
  massmailStatuses: Set<string>;
  registrations: SPRegistration[];
  setMassmailAudience: React.Dispatch<React.SetStateAction<MassmailAudience>>;
  setMassmailMode: React.Dispatch<React.SetStateAction<"closed" | "pick" | "paste" | "editor">>;
  setMassmailPasteRaw: React.Dispatch<React.SetStateAction<string>>;
  setMassmailStatuses: React.Dispatch<React.SetStateAction<Set<string>>>;
  setShowEmailModal: React.Dispatch<React.SetStateAction<boolean>>;
}

export const MassmailPickModal: React.FC<MassmailPickModalProps> = (p) => {
  const { massmailAudience, massmailStatuses, registrations, setMassmailAudience, setMassmailMode, setMassmailPasteRaw, setMassmailStatuses, setShowEmailModal } = p;
        // v31.2: Die Props kennen kein isDe (Schnittstelle bleibt) — die Sprache
        // kommt wie in Modal.tsx aus dem Kontext.
        const isDe = useLocaleSafe() === 'de';
        const closeAll = (): void => { setMassmailMode('closed'); setMassmailPasteRaw(''); };
        const proceed = (): void => {
          if (massmailAudience === 'custom' && massmailStatuses.size === 0) return;
          if (massmailAudience === 'nachruecker') setMassmailMode('paste');
          else { setShowEmailModal(true); setMassmailMode('editor'); }
        };
        // v31.9.6: Abgemeldet und No-Show gehören dazu. Sie fehlten nicht aus
        // einem Grund, sondern weil nie jemand danach gefragt hat — und der
        // Organizer braucht sie: „Termin verschoben, kommst du doch?" geht
        // genau an die, die abgesagt haben.
        const STATUS_OPTIONS = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste', 'Abgemeldet', 'No-Show'];
        const toggleStatus = (st: string): void => {
          setMassmailStatuses(prev => {
            const next = new Set(prev);
            if (next.has(st)) next.delete(st); else next.add(st);
            return next;
          });
        };
        // v31.2: Zähler je Gruppe stehen direkt an der Auswahl — „an wie viele
        // geht das?" ist die Frage, die der Organizer hier beantwortet, nicht
        // erst im Editor. Reine Anzeige; die Empfängerlogik bleibt im Aufrufer.
        const countOf = (stati: string[]): number => registrations.filter(r => stati.indexOf(r.Status) >= 0).length;
        const nActive = countOf(['Angemeldet', 'QR versendet', 'Eingecheckt']);
        const nWait = countOf(['Warteliste']);
        // Alle Zeilen der Liste — nicht `countOf` mit einer Status-Aufzählung,
        // sonst fehlt jeder Status, den jemand später in SharePoint ergänzt.
        const nEveryone = registrations.length;
        const nInactive = nEveryone - nActive - nWait;
        const customEmpty = massmailAudience === 'custom' && massmailStatuses.size === 0;
        const Row = (props: { value: MassmailAudience; label: string; desc: string; count?: number }): React.ReactElement => {
          const on = massmailAudience === props.value;
          return (
            <label className={cx('dex-ui-toggle-row', on && 'is-active')}>
              <input type="radio" name="massmail-target" checked={on} onChange={() => setMassmailAudience(props.value)} />
              <span className="dex-ui-toggle-row-body">
                <span className="dex-ui-toggle-row-title">
                  {props.label}
                  {props.count !== undefined && <span className={cx('dex-ui-pill', on ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')}>{props.count} {isDe ? 'Pers.' : 'people'}</span>}
                </span>
                <span className="dex-ui-toggle-row-desc">{props.desc}</span>
              </span>
            </label>
          );
        };
        // v31.2: Reihenfolge nach Häufigkeit und Mechanik — erst die drei
        // Status-Gruppen, dann die eigene Status-Auswahl (Verfeinerung derselben
        // Frage), zuletzt der manuelle Abgleich, weil er als einziger einen
        // zweiten Schritt hat. Vorher stand er zwischen den Status-Gruppen.
        return (
          <Modal open={true} onClose={closeAll} maxWidth={560}
            ariaLabel={isDe ? 'Empfänger wählen' : 'Choose recipients'}
            title={isDe ? 'An wen soll die Mail gehen?' : 'Who should get the mail?'}
            subtitle={isDe ? 'Wähle die Empfängergruppe — den Text schreibst du danach im Mail-Editor.' : 'Pick the recipient group — you write the text in the mail editor afterwards.'}
            icon={<Users size={20} />}
            footer={<>
              <button type="button" className="btn btn-secondary" onClick={closeAll}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
              <button type="button" className="btn btn-primary" onClick={proceed} disabled={customEmpty}>
                {massmailAudience === 'nachruecker' ? (isDe ? 'Weiter: Liste einfügen' : 'Next: paste list') : (isDe ? 'Weiter zum Mail-Editor' : 'Continue to mail editor')}
              </button>
            </>}>
            <div className="dex-ui-section">
              <div className="dex-ui-section-title">{isDe ? 'Nach Status' : 'By status'}</div>
              <div className="dex-ui-stack">
                <Row value="active" count={nActive} label={isDe ? 'Alle aktiven Teilnehmer' : 'All active participants'} desc={isDe ? 'Status Angemeldet, QR versendet oder Eingecheckt — der Normalfall für Info-Mails.' : 'Status registered, QR sent or checked in — the usual choice for info mails.'} />
                <Row value="activePlusWait" count={nActive + nWait} label={isDe ? 'Aktive Teilnehmer + Warteliste' : 'Active participants + waitlist'} desc={isDe ? 'Beide zusammen — z.B. wenn Plätze frei werden und du die Warteliste vorwarnen willst.' : 'Both together — e.g. when seats free up and you want to give the waitlist a heads-up.'} />
                <Row value="waitOnly" count={nWait} label={isDe ? 'Nur Warteliste' : 'Waitlist only'} desc={isDe ? 'Nur die Wartenden — z.B. „Es wird wahrscheinlich keinen Platz mehr geben“.' : 'Only those waiting — e.g. "There will probably be no more seats".'} />
                {/* v31.9.6: „An alle" inklusive Abgemeldete und No-Shows.
                    Bewusst als EIGENE Zeile mit eigener Zahl und nicht still in
                    „alle aktiven" hineingerechnet: Wer abgesagt hat, bekommt
                    normalerweise nichts mehr — das muss man ausdrücklich
                    wählen, nicht versehentlich treffen. Die Zeile nennt
                    deshalb, wie viele davon nicht mehr dabei sind. */}
                <Row value="everyone" count={nEveryone} label={isDe ? 'Alle — auch Abgemeldete' : 'Everyone — including cancellations'} desc={isDe ? `Jede Person in der Teilnehmerliste, auch Abgemeldete und No-Shows (${nInactive} davon nicht mehr dabei) — z.B. „Termin verschoben, kommst du doch?“.` : `Everyone in the participant list, including cancellations and no-shows (${nInactive} of them no longer attending) — e.g. "Date moved, joining after all?".`} />
                {/* v22.9: Eigene Status-Auswahl — einzelne Status getrennt anhaken. */}
                <div>
                  <Row value="custom" count={massmailAudience === 'custom' ? countOf(Array.from(massmailStatuses)) : undefined} label={isDe ? 'Eigene Auswahl nach Status' : 'Custom selection by status'} desc={isDe ? 'Du wählst unten, welche Status die Mail bekommen — z.B. nur „QR versendet“.' : 'You pick below which statuses get the mail — e.g. only "QR sent".'} />
                  {/* v31.10: Der Einzug richtet die Chips unter dem Titel der Zeile
                      darüber aus — auf dem Handy kostete er ein Sechstel der Breite
                      und drängte jeden Chip in eine eigene Reihe. `clamp` statt einer
                      Media-Query, weil ein Inline-Style keine haben kann: schmal
                      5vw, ab Tablet wieder die vollen 44 px. */}
                  {massmailAudience === 'custom' && (
                    <div className="dex-ui-inline" style={{ padding: '10px 0 2px', paddingLeft: 'clamp(14px, 5vw, 44px)' }}>
                      {STATUS_OPTIONS.map(st => {
                        const count = registrations.filter(r => r.Status === st).length;
                        return (
                          <button key={st} type="button" className={cx('dex-ui-chip', massmailStatuses.has(st) && 'is-active')} aria-pressed={massmailStatuses.has(st)} onClick={() => toggleStatus(st)}>
                            {st} <span style={{ opacity: 0.75 }}>({count})</span>
                          </button>
                        );
                      })}
                      {customEmpty && (
                        <span className="dex-ui-help" style={{ width: '100%', margin: 0, color: 'var(--dex-orange-dark, #b35a00)' }}>{isDe ? 'Wähle mindestens einen Status — sonst bekommt niemand die Mail.' : 'Pick at least one status — otherwise nobody gets the mail.'}</span>
                      )}
                    </div>
                  )}
                </div>
              </div>
            </div>
            <div className="dex-ui-section">
              <div className="dex-ui-section-title">{isDe ? 'Abgleich mit deiner Liste' : 'Compare with your list'}</div>
              <Row value="nachruecker" label={isDe ? 'Nachrücker — nur wer deine letzte Mail noch nicht hat' : 'Late joiners — only those who missed your last mail'} desc={isDe ? 'Im nächsten Schritt fügst du deine bisherige Empfänger-Liste ein (Verteiler, „Vorname Nachname <mail>“, beliebig formatiert). Die App erkennt die Adressen und schreibt alle aktiven Teilnehmer an, die dort NICHT stehen.' : 'In the next step you paste your existing recipient list (distribution list, "First Last <mail>", any format). The app picks out the addresses and mails every active participant who is NOT on it.'} />
            </div>
          </Modal>
        );
};

