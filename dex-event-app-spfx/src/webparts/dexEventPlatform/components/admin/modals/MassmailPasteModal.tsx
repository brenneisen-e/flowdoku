/* MassmailPasteModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 15021-15103 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.70: Derselbe Dialog für ZWEI Abgleiche (Nutzer-Ansage 17.09.2026):
 *  - nachruecker: aktive Teilnehmer, die NICHT im Verteiler stehen (bisher)
 *  - reminder:    Personen aus dem Verteiler, die NICHT aktiv angemeldet sind
 *    — die Erinnerung an alle, die noch nicht reagiert haben. Ihre Namen
 *    kommen aus dem Verteiler (utils/pastedRecipients), nicht aus der Liste.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { cx } from '../../dexUi';
import { Check, ChevronDown, Mail } from '../../Icons';
import { useLocaleSafe } from '../../../context/LanguageContext';
import { SPRegistration } from '../../../services/EventService';
import { MassmailAudience } from '../adminTypes';
import { parsePastedRecipients, splitName } from '../../../utils/pastedRecipients';

export interface MassmailPasteModalProps {
  /** v31.70: 'nachruecker' | 'reminder' — alles andere verhält sich wie nachruecker. */
  massmailAudience: MassmailAudience;
  massmailPasteRaw: string;
  registrations: SPRegistration[];
  setMassmailMode: React.Dispatch<React.SetStateAction<"closed" | "pick" | "paste" | "editor">>;
  setMassmailPasteRaw: React.Dispatch<React.SetStateAction<string>>;
  setShowEmailModal: React.Dispatch<React.SetStateAction<boolean>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

/**
 * v31.70: Die Erinnerungs-Empfänger aus einem Verteiler — EINE Rechnung für
 * Paste-Dialog und Composer (dieselbe Menge muss angezeigt und verschickt
 * werden). Aktiv angemeldete Adressen fallen raus; wer auf der Warteliste
 * steht, abgesagt hat oder gar nicht in der Liste ist, bleibt drin.
 */
export function reminderRecipientsAus(raw: string, registrations: SPRegistration[]): Array<{ ParticipantEmail: string; Vorname: string; Nachname: string; name: string }> {
  const aktiv = new Set(registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0).map(r => (r.ParticipantEmail || '').toLowerCase()));
  return parsePastedRecipients(raw)
    .filter(p => !aktiv.has(p.email))
    .map(p => { const n = splitName(p.name); return { ParticipantEmail: p.email, Vorname: n.vorname, Nachname: n.nachname, name: p.name }; });
}

export const MassmailPasteModal: React.FC<MassmailPasteModalProps> = (p) => {
  const { massmailAudience, massmailPasteRaw, registrations, setMassmailMode, setMassmailPasteRaw, setShowEmailModal, showAlert } = p;
        // v31.2: Die Props kennen kein isDe (Schnittstelle bleibt) — die Sprache
        // kommt wie in Modal.tsx aus dem Kontext.
        const isDe = useLocaleSafe() === 'de';
        const reminder = massmailAudience === 'reminder';
        const [showMissing, setShowMissing] = React.useState(false);
        const closeAll = (): void => { setMassmailMode('closed'); setMassmailPasteRaw(''); };
        const back = (): void => { setMassmailMode('pick'); };
        // E-Mail-Adressen aus dem Rohtext extrahieren — robust gegen Vorname
        // Nachname <mail@…> / mail@…; sep / Outlook-Verteiler-Dumps.
        // v31.70: über utils/pastedRecipients (Adresse UND Name).
        const pastedAll = parsePastedRecipients(massmailPasteRaw);
        const pasted = pastedAll.map(x => x.email);
        const active = registrations.filter(r => ACTIVE.indexOf(r.Status) >= 0);
        const pastedSet = new Set(pasted);
        // Nachrücker: aktive Teilnehmer, die NICHT im Verteiler stehen.
        const missing = active.filter(r => !pastedSet.has((r.ParticipantEmail || '').toLowerCase()));
        // Erinnerung: Verteiler-Personen, die NICHT aktiv angemeldet sind.
        const reminderList = reminder ? reminderRecipientsAus(massmailPasteRaw, registrations) : [];
        const nAlreadyActive = reminder ? pastedAll.length - reminderList.length : 0;
        const count = reminder ? reminderList.length : missing.length;
        const continueAction = (): void => {
          if (count === 0) {
            showAlert(reminder
              ? (isDe ? 'Alle Personen aus deinem Verteiler sind schon aktiv angemeldet — niemand zu erinnern.' : 'Everyone on your list is already actively registered — nobody to remind.')
              : (isDe ? 'Alle aktiven Teilnehmer stehen bereits in deiner Liste — niemand zum Anschreiben übrig.' : 'All active participants are already on your list — nobody left to write to.'));
            return;
          }
          setShowEmailModal(true);
          setMassmailMode('editor');
        };
        // v31.2: Frage statt Feldname im Kopf, drei Kennzahlen statt drei Fließtext-
        // Zeilen, die Empfängerliste hinter einem Aufklapper; der Primär-Knopf nennt
        // die Zahl. backdropClose aus: Ein Klick neben die Karte warf den
        // eingefügten Verteiler weg (Modal.tsx, v30.51).
        const kpis: [number, string, string][] = reminder
          ? [
            [pastedAll.length, isDe ? 'Adressen im Verteiler' : 'addresses on your list', ''],
            [nAlreadyActive, isDe ? 'davon schon aktiv angemeldet' : 'of them already registered', ''],
            [reminderList.length, isDe ? 'noch nicht angemeldet — werden erinnert' : 'not registered yet — will be reminded', 'dex-ui-kpi--orange'],
          ]
          : [
            [pasted.length, isDe ? 'Adressen erkannt' : 'addresses found', ''],
            [active.length, isDe ? 'aktive Teilnehmer' : 'active participants', ''],
            [missing.length, isDe ? 'nicht in deiner Liste — werden angeschrieben' : 'not on your list — will be mailed', 'dex-ui-kpi--orange'],
          ];
        return (
          <Modal open={true} onClose={closeAll} maxWidth={680} backdropClose={false}
            ariaLabel={reminder ? (isDe ? 'Erinnerung — Verteiler einfügen' : 'Reminder — paste list') : (isDe ? 'Nachrücker — Liste einfügen' : 'Late joiners — paste list')}
            title={reminder ? (isDe ? 'Wen hast du eingeladen?' : 'Who did you invite?') : (isDe ? 'Wer hat die Mail schon bekommen?' : 'Who already received the mail?')}
            subtitle={reminder
              ? (isDe ? 'Schritt 2 von 2 — füge deinen Einladungs-Verteiler ein. Erinnert wird, wer dort steht, aber im Event nicht aktiv angemeldet ist.' : 'Step 2 of 2 — paste your invitation list. The reminder goes to everyone on it who is not actively registered.')
              : (isDe ? 'Schritt 2 von 2 — füge deine bisherige Empfänger-Liste ein. Angeschrieben wird, wer im Event aktiv ist, aber dort nicht steht.' : 'Step 2 of 2 — paste your existing recipient list. The mail goes to everyone active in the event who is not on it.')}
            icon={<Mail size={20} />}
            footer={<>
              <span className="dex-ui-modal-foot-left"><button type="button" className="btn btn-secondary" onClick={back}>{isDe ? 'Zurück' : 'Back'}</button></span>
              <button type="button" className="btn btn-secondary" onClick={closeAll}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
              <button type="button" className="btn btn-primary" disabled={count === 0} onClick={continueAction}>{isDe ? `Weiter zum Mail-Editor (${count})` : `Continue to mail editor (${count})`}</button>
            </>}>
            <div className="dex-ui-field">
              <label className="dex-ui-label" htmlFor="massmail-paste-raw">{reminder ? (isDe ? 'Dein Einladungs-Verteiler' : 'Your invitation list') : (isDe ? 'Deine bisherige Empfänger-Liste' : 'Your existing recipient list')}</label>
              <textarea id="massmail-paste-raw" className="dex-ui-textarea" value={massmailPasteRaw} onChange={e => setMassmailPasteRaw(e.target.value)}
                placeholder={'Max Mustermann <mmustermann@deloitte.de>; anna.schmidt@deloitte.de; ...'}
                style={{ minHeight: 160, fontFamily: 'monospace', fontSize: '0.82rem' }} />
              <div className="dex-ui-help">
                {isDe ? 'Hau alles rein, was du hast: Verteiler-Export, Outlook-To-Zeile, „Vorname Nachname <mail@deloitte.de>“, komma-, semikolon- oder zeilengetrennt — die App pickt die E-Mail-Adressen automatisch heraus.' : 'Paste whatever you have: distribution list export, Outlook To line, "First Last <mail@deloitte.de>", separated by comma, semicolon or line break — the app picks out the email addresses automatically.'}
                {reminder && (isDe ? ' Die Namen nimmt die App aus dem Verteiler — diese Personen stehen ja nicht in der Teilnehmerliste.' : ' Names are taken from the list — these people are not in the participant list.')}
              </div>
            </div>
            <div className="dex-ui-grid-3">
              {kpis.map(([v, label, mod]) => (
                <div key={label} className={cx('dex-ui-kpi', mod)}><div className="dex-ui-kpi-value">{v}</div><div className="dex-ui-kpi-label">{label}</div></div>
              ))}
            </div>
            {count > 0 ? (
              <div>
                <button type="button" className={cx('dex-ui-disclosure', showMissing && 'is-open')} aria-expanded={showMissing} onClick={() => setShowMissing(v => !v)}>
                  <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                  {isDe ? 'Empfänger anzeigen' : 'Show recipients'}
                  <span className="dex-ui-disclosure-count">{count}</span>
                </button>
                {showMissing && (
                  <div className="dex-ui-disclosure-body">
                    <div className="dex-ui-table-wrap" style={{ maxHeight: 320, overflowY: 'auto' }}>
                      <table className="dex-ui-table">
                        {reminder ? (
                          <>
                            <thead><tr><th>{isDe ? 'Name (aus dem Verteiler)' : 'Name (from your list)'}</th><th>{isDe ? 'E-Mail' : 'Email'}</th><th>{isDe ? 'Stand im Event' : 'Status in event'}</th></tr></thead>
                            <tbody>
                              {reminderList.map(r => {
                                const row = registrations.find(x => (x.ParticipantEmail || '').toLowerCase() === r.ParticipantEmail);
                                return (
                                  <tr key={r.ParticipantEmail}>
                                    <td>{r.name || '-'}</td>
                                    <td style={{ color: 'var(--dex-gray-600)' }}>{r.ParticipantEmail}</td>
                                    <td style={{ color: 'var(--dex-gray-600)' }}>{row ? row.Status : (isDe ? 'nicht in der Liste' : 'not in the list')}</td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </>
                        ) : (
                          <>
                            <thead><tr><th>{isDe ? 'Vorname' : 'First name'}</th><th>{isDe ? 'Nachname' : 'Last name'}</th><th>{isDe ? 'Position' : 'Job title'}</th><th>{isDe ? 'E-Mail' : 'Email'}</th></tr></thead>
                            <tbody>
                              {missing.map(r => (
                                <tr key={r.Id}>
                                  <td>{r.Vorname || '-'}</td><td>{r.Nachname || '-'}</td>
                                  <td style={{ color: 'var(--dex-gray-600)' }}>{(r as SPRegistration & { JobTitle?: string }).JobTitle || '-'}</td>
                                  <td style={{ color: 'var(--dex-gray-600)' }}>{r.ParticipantEmail}</td>
                                </tr>
                              ))}
                            </tbody>
                          </>
                        )}
                      </table>
                    </div>
                  </div>
                )}
              </div>
            ) : (
              <div className="dex-ui-callout dex-ui-callout--success"><span className="dex-ui-callout-icon"><Check size={16} /></span>
                <span>{reminder
                  ? (pastedAll.length === 0
                    ? (isDe ? 'Noch keine Adresse erkannt — füge oben deinen Verteiler ein.' : 'No address recognised yet — paste your list above.')
                    : (isDe ? 'Alle Personen aus deinem Verteiler sind schon aktiv angemeldet — niemand zu erinnern.' : 'Everyone on your list is already actively registered — nobody to remind.'))
                  : (isDe ? 'Alle aktiven Teilnehmer stehen schon in deiner Liste — es bleibt niemand zum Anschreiben.' : 'All active participants are already on your list — nobody left to write to.')}</span></div>
            )}
          </Modal>
        );
};

