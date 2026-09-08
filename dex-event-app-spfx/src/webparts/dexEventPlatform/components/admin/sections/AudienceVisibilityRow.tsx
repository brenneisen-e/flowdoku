/* AudienceVisibilityRow — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 10282-10437 des
 * Stands vor dem Schnitt). Die Rechnung ist unverändert; seit v31.3 ist der
 * Kopf die Frage „Wer sieht das Event?" mit Pillen für Standorte und
 * Einträge. Die Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { ChevronDown, Users } from '../../Icons';
import { DeloitteEvent } from '../../../types';
import { AudiencePerson } from '../../admin/adminTypes';
import { cx } from '../../dexUi';

export interface AudienceVisibilityRowProps {
  isAdmin: boolean;
  isDe: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  openPendingReminder: () => Promise<void>;
  orgPastLock: boolean;
  pendingCheckBusy: boolean;
  resolveAudienceEmails: (ev: DeloitteEvent) => Promise<AudiencePerson[]>;
  selectedEvent: DeloitteEvent;
  setVisibilityAllAddresses: React.Dispatch<React.SetStateAction<boolean>>;
  setVisibilityBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setVisibilityOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setVisibilityResolved: React.Dispatch<React.SetStateAction<AudiencePerson[]>>;
  visibilityAllAddresses: boolean;
  visibilityBusy: boolean;
  visibilityOpen: boolean;
  visibilityResolved: AudiencePerson[];
}

export const AudienceVisibilityRow: React.FC<AudienceVisibilityRowProps> = (p) => {
  const { isAdmin, isDe, isOrganizerFor, openPendingReminder, orgPastLock, pendingCheckBusy, resolveAudienceEmails, selectedEvent, setVisibilityAllAddresses, setVisibilityBusy, setVisibilityOpen, setVisibilityResolved, visibilityAllAddresses, visibilityBusy, visibilityOpen, visibilityResolved } = p;
        const locs = (selectedEvent.locationAudience || []).filter(Boolean);
        const audEntries = (selectedEvent.audienceFilter || []).map(s => (s || '').trim()).filter(Boolean);
        const mailable = audEntries.filter(e => e.indexOf('@') > 0);
        const excludedCount = (selectedEvent.excludedUsers || []).filter(Boolean).length;
        const frozen = (selectedEvent.audienceResolvedEmails || []).length;
        const resolvedCount = visibilityResolved ? visibilityResolved.length : frozen;
        const modeAnd = selectedEvent.filterMode !== 'OR';
        const openToAll = locs.length === 0 && audEntries.length === 0;
        // Zahl nur zeigen, wenn sie etwas bedeutet. Ohne Filter sieht das Event
        // jede:r; bei reinem Standortfilter gibt es keine abzählbare Liste.
        const countable = mailable.length > 0 && resolvedCount > 0;
        const summary = openToAll
          ? (isDe ? 'Ohne Einschränkung — alle Mitarbeitenden sehen dieses Event' : 'No restriction — all employees can see this event')
          : countable
            ? (isDe
              ? `${resolvedCount} ${resolvedCount === 1 ? 'Person kann' : 'Personen können'} dieses Event sehen${locs.length > 0 && modeAnd ? ` · davon nur die am Standort ${locs.join(', ')}` : ''}`
              : `${resolvedCount} ${resolvedCount === 1 ? 'person can' : 'people can'} see this event${locs.length > 0 && modeAnd ? ` · of those only the ones at ${locs.join(', ')}` : ''}`)
            : (isDe
              ? `Sichtbar für ${locs.length > 0 ? `alle am Standort ${locs.join(', ')}` : 'ausgewählte Personen'} — wie viele das sind, weiß DEX hier nicht`
              : `Visible to ${locs.length > 0 ? `everyone at ${locs.join(', ')}` : 'selected people'} — DEX cannot tell how many that is`);
        const canManageVis = isAdmin || isOrganizerFor(selectedEvent);
        return (
          <div className="dex-ui-card" style={{ marginBottom: 12 }}>
            {/* v31.3: Die Ueberschrift ist die Frage, die der Organizer hier
                stellt. Der ganze Kopf ist der Aufklapper — Hover nur, weil ein
                Klick darauf wirklich etwas tut. */}
            <button
              type="button"
              className={cx('dex-ui-disclosure', visibilityOpen && 'is-open')}
              onClick={() => setVisibilityOpen(v => !v)}
              aria-expanded={visibilityOpen}
            >
              <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
              <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', display: 'inline-flex' }}><Users size={16} /></span>
              {isDe ? 'Wer sieht das Event?' : 'Who can see this event?'}
              {/* v31.3 (Nachzug): „Details"/„weniger" stand im Zähler-Slot, der
                  `margin-left: auto` trägt — das Wort hing allein am rechten
                  Rand (genau die Form, die Leitfaden 2a′ vermeidet). Als
                  Kleintext direkt hinter der Frage gehört es zum Aufklapper;
                  der Chevron rechts sagt ohnehin, in welche Richtung es geht. */}
              <span className="dex-ui-muted" style={{ fontSize: '0.76rem' }}>
                {visibilityOpen ? (isDe ? 'weniger' : 'less') : (isDe ? 'Details' : 'details')}
              </span>
            </button>
            <div className="dex-ui-muted" style={{ marginTop: 2 }}>{summary}</div>
            {/* Standorte und Eintraege auf einen Blick. Reine Anzeige, deshalb
                Pillen und keine Chips — hier gibt es nichts umzuschalten. */}
            {(locs.length > 0 || audEntries.length > 0 || excludedCount > 0) && (
              <div className="dex-ui-inline" style={{ marginTop: 8 }}>
                {/* Index als Key: derselbe Standort kann in den Daten zweimal stehen. */}
                {locs.map((l, i) => <span key={`loc${i}`} className="dex-ui-pill dex-ui-pill--blue">{l}</span>)}
                {audEntries.length > 0 && (
                  <span className="dex-ui-pill dex-ui-pill--gray">
                    {audEntries.length} {isDe
                      ? (audEntries.length === 1 ? 'Verteiler oder Person' : 'Verteiler & Personen')
                      : (audEntries.length === 1 ? 'list or person' : 'lists & people')}
                  </span>
                )}
                {excludedCount > 0 && (
                  <span className="dex-ui-pill dex-ui-pill--orange">
                    {excludedCount} {isDe ? 'ausgeschlossen' : 'excluded'}
                  </span>
                )}
              </div>
            )}
            {/* v29.36: Der Knopf stand rechts außen am Kartenrand — weit weg von
                der Zahl, auf die er sich bezieht, und ohne sichtbare Erklärung
                (sie steckte im title-Tooltip). Jetzt links unter der Zeile, mit
                dem Satz daneben, der sagt, was passiert. */}
            {canManageVis && !orgPastLock && mailable.length > 0 && (
              <div className="dex-ui-inline" style={{ marginTop: 10 }}>
                <button
                  type="button"
                  className="btn btn-secondary dex-ui-btn-sm"
                  style={{ flexShrink: 0 }}
                  disabled={pendingCheckBusy}
                  onClick={() => { void openPendingReminder(); }}
                >
                  {pendingCheckBusy
                    ? (isDe ? 'Wird geprüft…' : 'Checking…')
                    : (isDe ? 'Wer hat noch nicht geantwortet?' : 'Who has not responded yet?')}
                </button>
                <span className="dex-ui-muted" style={{ flex: '1 1 260px', minWidth: 0 }}>
                  {isDe
                    ? 'Listet zuerst alle, die das Event sehen, sich aber weder angemeldet noch abgemeldet noch abgesagt haben — erinnern kannst du sie im Schritt danach.'
                    : 'First lists everyone who can see this event but has neither registered nor cancelled nor declined — you can remind them in the next step.'}
                </span>
              </div>
            )}
            {visibilityOpen && (
              <div className="dex-ui-disclosure-body" style={{ marginTop: 10, paddingTop: 10, borderTop: '1px solid var(--dex-gray-200)', fontSize: '0.82rem', color: 'var(--dex-gray-700)', display: 'flex', flexDirection: 'column', gap: 6 }}>
                <div>
                  <strong>{isDe ? 'Standortfilter: ' : 'Location filter: '}</strong>
                  {locs.length > 0 ? locs.join(', ') : (isDe ? 'keiner' : 'none')}
                </div>
                {/* v29.36: 100 Adressen als Fließtext waren nicht lesbar — erst
                    die ersten zwölf, den Rest auf Klick. */}
                <div>
                  <strong>{isDe ? 'Eingetragen sind: ' : 'Entered here: '}</strong>
                  {audEntries.length === 0
                    ? (isDe ? 'keine Verteiler und keine einzelnen Personen' : 'no distribution lists and no individual people')
                    : (visibilityAllAddresses || audEntries.length <= 12
                      ? audEntries.join(', ')
                      : `${audEntries.slice(0, 12).join(', ')} …`)}
                  {audEntries.length > 12 && (
                    <button
                      type="button"
                      className="dex-ui-textbtn"
                      onClick={() => setVisibilityAllAddresses(v => !v)}
                      style={{ marginLeft: 4 }}
                    >
                      {visibilityAllAddresses
                        ? (isDe ? 'weniger anzeigen' : 'show less')
                        : (isDe ? `alle ${audEntries.length} anzeigen` : `show all ${audEntries.length}`)}
                    </button>
                  )}
                </div>
                <div style={{ color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Steht dort ein Mailverteiler, zählt DEX dessen Mitglieder einzeln — auch aus Verteilern, die selbst wieder Verteiler enthalten.'
                    : 'If a distribution list is entered, DEX counts its members individually — including lists nested inside other lists.'}
                </div>
                {locs.length > 0 && audEntries.length > 0 && (
                  <div style={{ color: 'var(--dex-gray-600)' }}>
                    {modeAnd
                      ? (isDe ? 'Verknüpfung: UND — es sehen nur Personen das Event, auf die BEIDES zutrifft.' : 'Combination: AND — only people matching BOTH can see the event.')
                      : (isDe ? 'Verknüpfung: ODER — es genügt eines von beidem.' : 'Combination: OR — either one is enough.')}
                  </div>
                )}
                {excludedCount > 0 && (
                  <div>
                    <strong>{isDe ? 'Ausgeschlossen: ' : 'Excluded: '}</strong>
                    {excludedCount} {isDe ? (excludedCount === 1 ? 'Person' : 'Personen') : (excludedCount === 1 ? 'person' : 'people')}
                  </div>
                )}
                {mailable.length > 0 && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                    {/* v29.36: „Aufgelöste Personen" war Fachsprache aus dem
                        Code — gemeint ist schlicht, wie viele Menschen am Ende
                        dahinterstehen. */}
                    <span>
                      <strong>{isDe ? 'Ergibt zusammen: ' : 'Adds up to: '}</strong>
                      {resolvedCount} {isDe ? (resolvedCount === 1 ? 'Person' : 'Personen') : (resolvedCount === 1 ? 'person' : 'people')}
                      <span style={{ color: 'var(--dex-gray-500)' }}>
                        {visibilityResolved
                          ? (isDe ? ' (gerade nachgezählt)' : ' (just counted)')
                          : (isDe ? ' (Stand: als das Event zuletzt gespeichert wurde)' : ' (as of the last time the event was saved)')}
                      </span>
                    </span>
                    <button
                      type="button"
                      className="btn btn-secondary dex-ui-btn-sm"
                      disabled={visibilityBusy}
                      onClick={() => {
                        void (async () => {
                          setVisibilityBusy(true);
                          try { setVisibilityResolved(await resolveAudienceEmails(selectedEvent)); }
                          catch { /* Anzeige bleibt beim eingefrorenen Stand */ }
                          finally { setVisibilityBusy(false); }
                        })();
                      }}
                    >
                      {visibilityBusy ? (isDe ? 'Wird gezählt…' : 'Counting…') : (isDe ? 'Jetzt neu nachzählen' : 'Count again now')}
                    </button>
                  </div>
                )}
                {locs.length > 0 && mailable.length === 0 && (
                  <div style={{ color: 'var(--dex-gray-600)' }}>
                    {isDe
                      ? 'Wenn die Sichtbarkeit nur über den Standort läuft, kennt DEX keine Namen dahinter — es gibt dann weder eine Personenzahl noch jemanden, den DEX erinnern könnte. Wenn du das brauchst, trage in Schritt 3 des Event-Edits zusätzlich einen Mailverteiler oder einzelne Personen ein.'
                      : 'If visibility runs via location only, DEX does not know the names behind it — so there is neither a headcount nor anyone DEX could remind. If you need that, add a distribution list or individual people in step 3 of the event edit.'}
                  </div>
                )}
              </div>
            )}
          </div>
        );
};

