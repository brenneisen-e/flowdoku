/* DuplicateRegHintBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10632-10735 des
 * Stands vor dem Schnitt). Welche Zeilen als Dublette gelten, ist unverändert.
 *
 * v31.3: Nach docs/ui-leitfaden.md 5a (Punkt 2) umgebaut — je ein
 * `dex-ui-callout` (Klammer-Dubletten `--info`, weil sie ausdrücklich harmlos
 * sind; echte Doppel-Anmeldungen `--danger`, weil sie einen Platz kosten), die
 * Zahl als Pille, der Bereinigen-Knopf IN der Zeile statt unter dem Fließtext
 * (Leitfaden 2a′), und die Personenliste hinter einem Aufklapper: Bei einem
 * Dutzend Betroffener schob sie die Teilnehmerliste bisher vom Bildschirm.
 */
import * as React from 'react';
import { SPRegistration } from '../../../services/EventService';
import { DUP_ACTIVE_STATI } from '../../admin/adminConstants';
import { AlertCircle, ChevronDown, Info } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';
import { DeloitteEvent } from '../../../types';

export interface DuplicateRegHintBoxProps {
  cleanupShadowDuplicates: () => Promise<void>;
  duplicateEmails: Set<string>;
  isAdmin: boolean;
  isConsolidatedMode: boolean;
  isDe: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
  shadowDupBusy: boolean;
}

/** v31.3: Eine Person je Zeile im Aufklapper — ruhig, ohne Hover (nichts klickbar). */
const LINE: React.CSSProperties = { display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'baseline', padding: '8px 0', fontSize: '0.84rem' };

export const DuplicateRegHintBox: React.FC<DuplicateRegHintBoxProps> = (p) => {
  const { cleanupShadowDuplicates, duplicateEmails, isAdmin, isConsolidatedMode, isDe, isOrganizerFor, registrations, selectedEvent, shadowDupBusy } = p;
          // v31.3: Hook vor jedem frühen Return (rules-of-hooks ist `error`);
          // ensureDexUiStyles ist idempotent und nötig, weil der Kasten ohne
          // Modal/WizardFormShell auf der Seite steht.
          const [listOpen, setListOpen] = React.useState(false);
          ensureDexUiStyles();
          // v23.2: Doppel-Anmelde-Hinweis. Listet jede Person, die mit
          // derselben E-Mail ≥2 nicht-abgemeldete Zeilen hat (z.B. dieselbe
          // Person in zwei Teams). Die betroffenen Zeilen sind in der Tabelle
          // zusätzlich rot markiert; pro Person kann der Organizer über den
          // „Abmelden"-Button das Duplikat still entfernen.
          if (duplicateEmails.size === 0 || !selectedEvent) return null;
          // Pro betroffener E-Mail die aktiven Zeilen sammeln (Name + Teams).
          const dupGroups: Array<{ email: string; rows: SPRegistration[] }> = [];
          duplicateEmails.forEach(em => {
            const rows = registrations.filter(r => DUP_ACTIVE_STATI.indexOf(r.Status || '') >= 0 && (r.ParticipantEmail || '').trim().toLowerCase() === em);
            if (rows.length > 1) dupGroups.push({ email: em, rows });
          });
          if (dupGroups.length === 0) return null;
          const nameOf = (g: { email: string; rows: SPRegistration[] }): string =>
            (g.rows[0].Vorname && g.rows[0].Nachname) ? `${g.rows[0].Vorname} ${g.rows[0].Nachname}` : (g.rows[0].ParticipantName || g.email);
          // v31.3: Ein Aufklapper-Knopf für beide Varianten — gleiche Bedienung,
          // egal ob harmlose Klammer-Zeile oder echte Doppel-Anmeldung.
          const disclosure = (
            <button type="button" className={cx('dex-ui-disclosure', listOpen && 'is-open')} aria-expanded={listOpen} onClick={() => setListOpen(o => !o)}>
              <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
              {isDe ? `Wen es betrifft (${dupGroups.length})` : `Who is affected (${dupGroups.length})`}
            </button>
          );
          // v28.21: Klammer-Modus — doppelte Zeilen auf der KLAMMER sind keine
          // Doppel-Anmeldungen. Die Klammer-Zeile ist nur eine Schatten-Zeile
          // zur Datenvollständigkeit (kein Platz, keine Mail, kein Outlook);
          // die echten Anmeldungen liegen in den Sub-Events, und alle Zähler
          // rechnen ohnehin pro Person entdoppelt. Solche Zeilen entstehen
          // z.B., wenn zwei verschiedene Assistenzen dieselbe Person nach-
          // einander anmelden — die Vorab-Prüfung sieht die fremde Zeile
          // wegen der Zeilen-Berechtigungen nicht. Deshalb hier ein neutraler
          // technischer Hinweis statt der roten Doppel-Anmelde-Warnung.
          if (isConsolidatedMode) {
            return (
              <div className="dex-ui-callout dex-ui-callout--info" style={{ marginBottom: 20 }}>
                <span className="dex-ui-callout-icon" aria-hidden="true"><Info size={16} /></span>
                <div style={{ minWidth: 0, flex: 1 }}>
                  <div style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                    {isDe ? 'Doppelte Klammer-Zeilen' : 'Duplicate overall-event rows'}
                    <span className="dex-ui-pill dex-ui-pill--blue">{isDe ? `${dupGroups.length} Personen` : `${dupGroups.length} people`}</span>
                  </div>
                  <div style={{ marginTop: 2 }}>
                    {isDe
                      ? <>Das ist <strong>keine Doppel-Anmeldung</strong>: Die Klammer-Zeile hält nur die Antworten auf die Hauptevent-Felder, belegt keinen Platz und löst weder Mail noch Outlook-Termin aus. Aufräumen ist optional.</>
                      : <>This is <strong>not a duplicate registration</strong>: the overall-event row only holds the answers to the main-event fields, takes no seat and triggers neither email nor calendar invite. Cleaning up is optional.</>}
                  </div>
                  {(isAdmin || isOrganizerFor(selectedEvent)) && (
                    <div className="dex-ui-inline" style={{ marginTop: 10 }}>
                      <button type="button" className="btn btn-secondary dex-ui-btn-sm" disabled={shadowDupBusy}
                        onClick={() => { cleanupShadowDuplicates().catch(() => { /* */ }); }}
                        title={isDe
                          ? 'Entfernt je Person die überzählige Klammer-Zeile — still, ohne Mail/Outlook/Nachrücken.'
                          : 'Removes the surplus overall-event row per person — silently, no email/Outlook/promotion.'}>
                        {shadowDupBusy
                          ? (isDe ? 'Wird bereinigt…' : 'Cleaning up…')
                          : (isDe ? `Doppelte Zeilen bereinigen (${dupGroups.length})` : `Clean up duplicate rows (${dupGroups.length})`)}
                      </button>
                    </div>
                  )}
                  {disclosure}
                  {listOpen && (
                    <div className="dex-ui-disclosure-body">
                      <div className="dex-ui-card" style={{ padding: '2px 12px' }}>
                        {dupGroups.map(g => (
                          <div key={g.email} className="dex-ui-row--bordered" style={LINE}>
                            <strong>{nameOf(g)}</strong>
                            <span className="dex-ui-muted">{g.email}</span>
                            <span className="dex-ui-pill dex-ui-pill--gray">{isDe ? `${g.rows.length} Klammer-Zeilen` : `${g.rows.length} overall-event rows`}</span>
                          </div>
                        ))}
                      </div>
                      <div className="dex-ui-muted" style={{ marginTop: 8 }}>
                        {isDe
                          ? `Die echten Anmeldungen stehen in den ${selectedEvent.childEventTermPlural || 'Sub-Events'}, und alle Zähler rechnen pro Person entdoppelt. Typische Ursache: zwei verschiedene Assistenzen haben dieselbe Person nacheinander angemeldet.`
                          : 'The real registrations are in the sub-events, and all counters de-duplicate per person. Typical cause: two different assistants registered the same person one after another.'}
                      </div>
                    </div>
                  )}
                </div>
              </div>
            );
          }
          return (
            <div className="dex-ui-callout dex-ui-callout--danger" style={{ marginBottom: 20 }}>
              <span className="dex-ui-callout-icon" aria-hidden="true"><AlertCircle size={16} /></span>
              <div style={{ minWidth: 0, flex: 1 }}>
                <div style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  {isDe ? 'Doppel-Anmeldungen erkannt' : 'Duplicate registrations detected'}
                  <span className="dex-ui-pill dex-ui-pill--red">{isDe ? `${dupGroups.length} Personen` : `${dupGroups.length} people`}</span>
                </div>
                <div style={{ marginTop: 2 }}>
                  {isDe
                    ? 'Dieselbe Person ist mehrfach angemeldet und belegt dadurch zwei Plätze. Die betroffenen Zeilen sind unten in der Tabelle rot markiert — über „Abmelden“ entfernst du die doppelte Zeile still.'
                    : 'The same person is registered more than once and therefore takes two seats. The affected rows are marked red in the table below — use „Cancel“ to silently remove the duplicate row.'}
                </div>
                {disclosure}
                {listOpen && (
                  <div className="dex-ui-disclosure-body dex-ui-card" style={{ padding: '2px 12px' }}>
                    {dupGroups.map(g => (
                      <div key={g.email} className="dex-ui-row--bordered" style={LINE}>
                        <strong>{nameOf(g)}</strong>
                        <span className="dex-ui-muted">{g.email}</span>
                        <span className="dex-ui-pill dex-ui-pill--red">{isDe ? `${g.rows.length}× angemeldet` : `${g.rows.length}× registered`}</span>
                        <span className="dex-ui-muted">— {g.rows.map(r => r.TeamName ? `„${r.TeamName}“` : (r.TeamId ? (isDe ? '(Team ohne Namen)' : '(unnamed team)') : (isDe ? '(Einzel-Anmeldung)' : '(individual)'))).join(', ')}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          );
};

