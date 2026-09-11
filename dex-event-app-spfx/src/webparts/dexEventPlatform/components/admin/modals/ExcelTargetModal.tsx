/* ExcelTargetModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 15129-15241 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { Check, Download } from '../../Icons';
import { cx } from '../../dexUi';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import { ExcelExportAudience, ExcelExportMode } from '../logic/createExportActions';

export interface ExcelTargetModalProps {
  consolidatedChildren: DeloitteEvent[];
  excelAudience: ExcelExportAudience;
  excelIncludeMatrix: boolean;
  excelSubIds: Set<string>;
  excelTargetModal: { mode: ExcelExportMode; chooseMode?: boolean; };
  exportConsolidatedExcel: (audience: ExcelExportAudience, includeMatrix: boolean, subIds: string[]) => void;
  exportCsv: (mode: ExcelExportMode, audience?: ExcelExportAudience) => void;
  /** v31.13: Die F&A-Liste zusätzlich als Mail — `null`, wenn der Weg fehlt. */
  sendFaListMail: ((empfaenger: string, nurEingecheckt: boolean) => Promise<boolean>) | null;
  /** Vorbelegung des Empfängerfelds: die eigene Adresse. */
  myEmail: string;
  isConsolidatedMode: boolean;
  isDe: boolean;
  selectedEvent: DeloitteEvent;
  setExcelAudience: React.Dispatch<React.SetStateAction<ExcelExportAudience>>;
  setExcelIncludeMatrix: React.Dispatch<React.SetStateAction<boolean>>;
  setExcelSubIds: React.Dispatch<React.SetStateAction<Set<string>>>;
  setExcelTargetModal: React.Dispatch<React.SetStateAction<{ mode: ExcelExportMode; chooseMode?: boolean; }>>;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
}

export const ExcelTargetModal: React.FC<ExcelTargetModalProps> = (p) => {
  const { consolidatedChildren, excelAudience, excelIncludeMatrix, excelSubIds, excelTargetModal, exportConsolidatedExcel, exportCsv, isConsolidatedMode, isDe, myEmail, selectedEvent, sendFaListMail, setExcelAudience, setExcelIncludeMatrix, setExcelSubIds, setExcelTargetModal, subEventRegsByEventId } = p;
        const closeAll = (): void => setExcelTargetModal(null);
        const istFa = excelTargetModal.mode === 'fa';
        // v31.13: Die F&A-Liste zusaetzlich per Mail. Eigener Zustand im
        // Dialog — der Versand ist eine Zugabe zum Download, kein zweiter
        // Weg daneben (ein zweiter Knopf waere die naechste „welcher gilt
        // jetzt"-Frage).
        const [mailAn, setMailAn] = React.useState(false);
        const [mailAdresse, setMailAdresse] = React.useState(myEmail || '');
        const [mailLaeuft, setMailLaeuft] = React.useState(false);
        const [mailMeldung, setMailMeldung] = React.useState('');
        // v20.4: Im Klammer-Modus entscheidet das Modal, WAS exportiert wird —
        // konsolidierte Matrix und/oder einzelne Sub-Event-Blätter.
        const consolidatedExportPossible = isConsolidatedMode && excelTargetModal.mode === 'deloitte' && consolidatedChildren.length > 0;
        const proceed = (): void => {
          const mode = excelTargetModal.mode;
          // v31.13: Erst die Mail (sie kann fehlschlagen und muss das sagen
          // koennen), dann der Download und das Schliessen. Andersherum waere
          // der Dialog weg, bevor die Meldung erscheint — und der Organizer
          // wuesste nie, ob die Mail rausging.
          if (istFa && mailAn && sendFaListMail) {
            const adresse = mailAdresse.trim();
            if (!adresse) { setMailMeldung(isDe ? 'Bitte trag eine Adresse ein, an die die Liste gehen soll.' : 'Please enter an address for the list.'); return; }
            setMailLaeuft(true);
            setMailMeldung('');
            sendFaListMail(adresse, excelAudience === 'checkedIn').then(ok => {
              setMailLaeuft(false);
              if (!ok) {
                setMailMeldung(isDe
                  ? 'Die Mail konnte nicht in die Warteschlange gelegt werden — es wurde nichts versendet. Die Datei kannst du trotzdem herunterladen.'
                  : 'The email could not be queued — nothing was sent. You can still download the file.');
                return;
              }
              setExcelTargetModal(null);
              exportCsv(mode, excelAudience);
            }, () => {
              setMailLaeuft(false);
              setMailMeldung(isDe ? 'Die Mail konnte nicht versendet werden.' : 'The email could not be sent.');
            });
            return;
          }
          setExcelTargetModal(null);
          if (consolidatedExportPossible && (excelIncludeMatrix || excelSubIds.size > 0)) {
            exportConsolidatedExcel(excelAudience, excelIncludeMatrix, Array.from(excelSubIds));
          } else {
            exportCsv(mode, excelAudience);
          }
        };
        const toggleSubId = (id: string): void => {
          setExcelSubIds(prev => {
            const next = new Set(prev);
            if (next.has(id)) next.delete(id); else next.add(id);
            return next;
          });
        };
        // v31.2: Eine Kachel je Alternative (Leitfaden 2b) statt Radio-Zeile —
        // Titel plus eine Zeile, was die Wahl für die Datei bedeutet. Der Hover
        // kommt aus der Klasse; vorher hatte keine der Zeilen einen.
        const Choice = (props: { active: boolean; onPick: () => void; label: string; desc: string }): React.ReactElement => (
          <button type="button" role="radio" aria-checked={props.active} className={cx('dex-ui-choice', props.active && 'is-active')} onClick={props.onPick}>
            <span className="dex-ui-choice-body">
              <span className="dex-ui-choice-title">{props.label}</span>
              <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{props.desc}</span>
            </span>
            <span className="dex-ui-choice-check" aria-hidden="true">{props.active && <Check size={12} />}</span>
          </button>
        );
        const Row = (props: { value: 'active' | 'activePlusWait' | 'waitOnly' | 'withCancelled'; label: string; desc: string }): React.ReactElement => (
          <Choice active={excelAudience === props.value} onPick={() => setExcelAudience(props.value)} label={props.label} desc={props.desc} />
        );
        return (
          <Modal open={true} onClose={closeAll} maxWidth={560} ariaLabel="Excel-Export Zielgruppe"
            title={isDe ? 'Excel-Export' : 'Excel export'}
            subtitle={isDe ? 'Wähle, wer und was in die Datei kommt — der Download startet sofort.' : 'Choose who and what goes into the file — the download starts right away.'}
            icon={<Download size={20} />}
            footer={<>
              <button type="button" className="btn btn-secondary" onClick={closeAll}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
              {/* Der Knopf sagt, was er TUT: Mit eingeschaltetem Versand
                  laedt er nicht nur herunter. Ein Knopf, der mehr tut als
                  er verspricht, ist die naechste Rueckfrage. */}
              <button type="button" className="btn btn-primary" onClick={proceed} disabled={mailLaeuft}>
                <Download size={16} />
                {mailLaeuft
                  ? (isDe ? 'Wird gesendet …' : 'Sending …')
                  : (istFa && mailAn
                    ? (isDe ? 'Herunterladen und Mail senden' : 'Download and send email')
                    : (isDe ? 'Excel herunterladen' : 'Download Excel'))}
              </button>
            </>}>
            {/* v27.9: Format-Auswahl (Deloitte/B2Run) direkt im Modal — vorher
                im Anker-Dropdown, das vom „Aktion auswählen"-Menü abgeschnitten
                wurde. v31.2: Steht zuerst, weil das Format entscheidet, ob der
                Klammer-Block unten überhaupt erscheint. */}
            {excelTargetModal.chooseMode && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Welches Format brauchst du?' : 'Which format do you need?'}</div>
                <div className="dex-ui-grid-2" role="radiogroup">
                  {([
                    { m: 'b2run' as const, label: 'B2Run View', desc: isDe ? 'Das Veranstalter-Format: ein Arbeitsblatt, 16 Spalten (Nr., Anrede, Name, E-Mail, Startblock …) — direkt bei b2run.com importierbar.' : 'The organizer format: one worksheet, 16 columns (no., salutation, name, e-mail, start block …) — importable at b2run.com.' },
                    { m: 'deloitte' as const, label: isDe ? 'Deloitte-Felder' : 'Deloitte fields', desc: isDe ? 'Alle internen Spalten plus alle eigenen Felder des Events — für den internen Gebrauch.' : 'All internal columns plus all custom fields of the event — for internal use.' },
                    // v31.13: die dritte Ansicht. Bewusst OHNE Rollen-Sperre:
                    // Nutzer-Ansage 11.09.2026 „soll auch direkt für alle
                    // Organizer zur Verfügung gestellt werden". Der Abrechnungs-
                    // SCHRITT im Assistenten bleibt davon unberührt — der hängt
                    // weiter an FA_BILLING_STEP_FOR_ORGANIZERS.
                    { m: 'fa' as const, label: isDe ? 'F&A-Liste' : 'F&A list', desc: isDe ? 'Der Aufbau, den Finance & Accounting einliest: oben die Abrechnungsangaben des Events, darunter die Personen (E-Mail, Vor-/Nachname, Land, Gesellschaft).' : 'The layout Finance & Accounting reads: the event billing details on top, the people below (e-mail, first/last name, country, company).' },
                  ]).map(opt => (
                    <Choice key={opt.m} active={excelTargetModal.mode === opt.m} onPick={() => {
                      setExcelTargetModal({ mode: opt.m, chooseMode: true });
                      // Die Zielgruppe muss zum Format passen. Ohne das stuende
                      // beim Wechsel auf F&A noch „Nur Warteliste" im Zustand,
                      // waehrend der Abschnitt darunter „Alle, die sich nicht
                      // abgemeldet haben" markiert zeigt — und exportiert wuerde
                      // die Warteliste. Ein Dialog, der etwas anderes tut als er
                      // anzeigt, ist schlimmer als eine Auswahl zu viel.
                      if (opt.m === 'fa') {
                        if (excelAudience !== 'checkedIn') setExcelAudience('active');
                      } else if (excelAudience === 'checkedIn') {
                        setExcelAudience('active');
                      }
                    }} label={opt.label} desc={opt.desc} />
                  ))}
                </div>
              </div>
            )}
            {istFa ? (
              /* v31.13 — Nutzer-Ansage 11.09.2026: „bei der F&A-Liste soll man
                 auch gefragt werden, ob dort nur die Leute drauf kommen sollen
                 die eingecheckt wurden oder alle die nicht abgemeldet sind."
                 Für einen Bewirtungsbeleg ist das keine Kleinigkeit: Wer nicht
                 da war, gehört nicht darauf. Deshalb steht hier KEINE Vorgabe
                 aus dem übrigen Export, sondern genau diese zwei Möglichkeiten
                 — und jede sagt, welche Status sie einschließt. */
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Wer kommt auf die F&A-Liste?' : 'Who goes on the F&A list?'}</div>
                <div className="dex-ui-stack" style={{ gap: 8 }}>
                  <Choice
                    active={excelAudience === 'checkedIn'}
                    onPick={() => setExcelAudience('checkedIn')}
                    label={isDe ? 'Nur Eingecheckte' : 'Checked-in only'}
                    desc={isDe
                      ? 'Nur Personen mit Status „Eingecheckt" — also die, deren Anwesenheit erfasst wurde. Das ist die Auswahl für einen Bewirtungsbeleg.'
                      : 'Only people with status “checked in” — those whose attendance was recorded. This is the selection for a hospitality receipt.'}
                  />
                  <Choice
                    active={excelAudience !== 'checkedIn'}
                    onPick={() => setExcelAudience('active')}
                    label={isDe ? 'Alle, die sich nicht abgemeldet haben' : 'Everyone who did not cancel'}
                    desc={isDe
                      ? 'Angemeldet, QR versendet und Eingecheckt. Wartende und Abgemeldete stehen nicht darauf.'
                      : 'Registered, QR sent and checked in. Waitlisted and cancelled people are not included.'}
                  />
                </div>
                {excelAudience === 'checkedIn' && (
                  <div className="dex-ui-help">
                    {isDe
                      ? 'Wurde bei diesem Event nicht eingecheckt, ist die Liste leer — dann nimm die zweite Auswahl.'
                      : 'If nobody was checked in at this event the list is empty — use the second option then.'}
                  </div>
                )}
              </div>
            ) : (
            <div className="dex-ui-section">
              <div className="dex-ui-section-title">{isDe ? 'Wer kommt in die Datei?' : 'Who goes into the file?'}</div>
              <div className="dex-ui-grid-2" role="radiogroup">
                <Row value="active" label={isDe ? 'Nur aktive Teilnehmer' : 'Active attendees only'} desc={isDe ? 'Angemeldet, QR versendet, Eingecheckt — die Vorgabe für Check-in und Vor-Ort-Liste.' : 'Registered, QR sent, checked in — the default for check-in and the on-site list.'} />
                <Row value="activePlusWait" label={isDe ? 'Teilnehmer + Warteliste' : 'Attendees + waitlist'} desc={isDe ? 'Aktive und Wartende in einem Blatt, sortiert nach Teilnehmer-ID.' : 'Active and waitlisted in one sheet, sorted by participant ID.'} />
                <Row value="waitOnly" label={isDe ? 'Nur Warteliste' : 'Waitlist only'} desc={isDe ? 'Nur die Wartenden — z. B. für ein Briefing.' : 'Only the waitlisted — e.g. for a briefing.'} />
                <Row value="withCancelled" label={isDe ? 'Alles inkl. Abmeldungen' : 'Everything incl. cancellations'} desc={isDe ? 'Alle Einträge, auch abgemeldete Personen — der Status steht je Zeile in der Status-Spalte.' : 'All entries including cancelled people — the status is in each row’s status column.'} />
              </div>
            </div>
            )}

            {/* v31.13: Die Liste zusätzlich per Mail — mit dem Dank fürs
                Organisieren. Der Text steht in `renderOrganizerFAListMailBody`,
                damit Mail und Datei dieselbe Auswahl beschreiben. */}
            {istFa && sendFaListMail && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Soll die Liste auch per Mail rausgehen?' : 'Should the list also go out by email?'}</div>
                <label className={cx('dex-ui-toggle-row', mailAn && 'is-active')}>
                  <input type="checkbox" checked={mailAn} disabled={mailLaeuft} onChange={e => { setMailAn(e.target.checked); setMailMeldung(''); }} />
                  <span className="dex-ui-toggle-row-body">
                    <span className="dex-ui-toggle-row-title">{isDe ? 'Ja, zusätzlich als Mail senden' : 'Yes, also send it as an email'}</span>
                    <span className="dex-ui-toggle-row-desc" style={{ display: 'block' }}>
                      {isDe
                        ? 'Die Liste steht dann als Tabelle IM Mailtext — kein Anhang, weil Anhänge aus DEX im Tenant blockiert sind und stillschweigend fehlen würden.'
                        : 'The list then sits as a table INSIDE the email text — no attachment, because attachments from DEX are blocked in this tenant and would silently go missing.'}
                    </span>
                  </span>
                </label>
                {mailAn && (
                  <div className="dex-ui-field" style={{ marginTop: 10 }}>
                    <label className="dex-ui-label" htmlFor="fa-mail-an">{isDe ? 'An welche Adresse?' : 'To which address?'}</label>
                    <input
                      id="fa-mail-an"
                      className="dex-ui-input"
                      type="email"
                      value={mailAdresse}
                      disabled={mailLaeuft}
                      onChange={e => { setMailAdresse(e.target.value); setMailMeldung(''); }}
                      placeholder="vorname.nachname@deloitte.de"
                    />
                    <div className="dex-ui-help">
                      {isDe
                        ? 'Vorbelegt ist deine eigene Adresse. Mehrere Adressen mit Semikolon trennen.'
                        : 'Your own address is prefilled. Separate several addresses with a semicolon.'}
                    </div>
                  </div>
                )}
                {mailMeldung && (
                  <div className="dex-ui-callout dex-ui-callout--warn" role="status" style={{ marginTop: 10 }}>
                    <span className="dex-ui-callout-body">{mailMeldung}</span>
                  </div>
                )}
              </div>
            )}
            {/* v20.4: Klammer-Modus — wählen, was in die Datei kommt. */}
            {consolidatedExportPossible && (
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Was kommt in die Datei?' : 'What goes into the file?'}</div>
                <label className={cx('dex-ui-toggle-row', excelIncludeMatrix && 'is-active')}>
                  <input type="checkbox" checked={excelIncludeMatrix} onChange={e => setExcelIncludeMatrix(e.target.checked)} />
                  <span className="dex-ui-toggle-row-body">
                    <span className="dex-ui-toggle-row-title">{isDe ? 'Konsolidierte Matrix' : 'Consolidated matrix'}</span>
                    <span className="dex-ui-toggle-row-desc" style={{ display: 'block' }}>
                      {isDe ? 'Eine Zeile pro Person: die übergreifenden Felder plus je Sub-Event Status und Antworten — wie die Tabelle in der Klammer-Ansicht.' : 'One row per person: the cross-cutting fields plus status and answers per sub-event — like the table in the consolidated view.'}
                    </span>
                  </span>
                </label>
                <div className="dex-ui-field" style={{ marginTop: 12 }}>
                  <div className="dex-ui-label">
                    {isDe ? 'Dazu ein eigenes Blatt je Sub-Event' : 'Plus a separate sheet per sub-event'}
                    <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                    <span className="dex-ui-inline" style={{ marginLeft: 'auto', gap: 2 }}>
                      <button type="button" className="dex-ui-textbtn" onClick={() => setExcelSubIds(new Set(consolidatedChildren.map(c => c.id)))}>{isDe ? 'Alle' : 'All'}</button>
                      <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => setExcelSubIds(new Set())}>{isDe ? 'Keine' : 'None'}</button>
                    </span>
                  </div>
                  {/* v31.2: Chips statt Checkbox-Zeilen — „mehrere aus vielen" (Leitfaden 2b);
                      die Zahl im Chip ist die Anzahl der Einträge des Termins. */}
                  <div className="dex-ui-inline" style={{ maxHeight: 180, overflowY: 'auto' }}>
                    {consolidatedChildren.map(child => {
                      const checked = excelSubIds.has(child.id);
                      const short = shortSubEventTitle(child.title, selectedEvent.title) || child.title || '?';
                      const n = (subEventRegsByEventId[child.id] || []).length;
                      return (
                        <button key={child.id} type="button" aria-pressed={checked} className={cx('dex-ui-chip', checked && 'is-active')} onClick={() => toggleSubId(child.id)} title={`${n} ${isDe ? 'Einträge' : 'entries'}`}>
                          {checked && <Check size={12} />}{short}
                          <span style={{ opacity: 0.75, fontWeight: 500 }}>· {n}</span>
                        </button>
                      );
                    })}
                  </div>
                  <div className="dex-ui-help">
                    {isDe ? 'Ohne Haken bei Matrix oder Sub-Event bekommst du die einfache Teilnehmerliste ohne Sub-Event-Blätter.' : 'With neither the matrix nor a sub-event ticked you get the plain attendee list without sub-event sheets.'}
                  </div>
                </div>
              </div>
            )}
          </Modal>
        );
};

