/* ConsolidatedView — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 5377-6469 des Stands
 * vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { DeloitteEvent } from '../../../types';
import { isEventOver } from '../../../utils/eventFormat';
import { ConsolidatedRow } from '../../admin/adminTypes';
import { Icon } from '@fluentui/react/lib/Icon';
import { PersonContactHover } from '../../PersonContactHover';
import { formatDate, translateStatus } from '../../../utils/eventStatus';
import { SPRegistration } from '../../../services/EventService';
import { Check, ExternalLink, Plus, Trash2 } from '../../Icons';
import { shortSubEventTitle } from '../../../utils/subEventTitle';

export interface ConsolidatedViewProps {
  addAllToKlammer: (rows: ConsolidatedRow[]) => Promise<void>;
  addingToKlammer: string;
  addToKlammer: (row: ConsolidatedRow) => Promise<void>;
  bulkKlammerProgress: string;
  colToggleHover: boolean;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  consolidatedChildren: DeloitteEvent[];
  consolidatedFiltered: ConsolidatedRow[];
  consolidatedRows: ConsolidatedRow[];
  consolidatedSort: string;
  consolidatedSortAsc: boolean;
  expandedConsolidatedEmail: string;
  highlightMatch: (text: unknown) => React.ReactNode;
  inactiveAccounts: string[];
  isAdmin: boolean;
  isConsolidatedMode: boolean;
  isDe: boolean;
  isLoadingSubEventRegs: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  missingReminderKey: string;
  openDeregModal: (row: ConsolidatedRow) => void;
  openMainFieldsEdit: (emailKey: string, displayName: string) => void;
  orgPastLock: boolean;
  performSilentDuplicateDelete: (reg: SPRegistration) => Promise<void>;
  personalColsCollapsed: boolean;
  registrations: SPRegistration[];
  reminderBusyId: number;
  searchQuery: string;
  selectedEvent: DeloitteEvent;
  sendCompleteRegistrationReminder: (args: { eventId: string; eventTitle: string; participantEmail: string; participantName: string; registeredByEmail?: string; registeredByName?: string; }) => Promise<boolean>;
  setAssignAssistRow: React.Dispatch<React.SetStateAction<ConsolidatedRow>>;
  setAssignAssistValue: React.Dispatch<React.SetStateAction<string>>;
  setColToggleHover: React.Dispatch<React.SetStateAction<boolean>>;
  setConsolidatedSort: React.Dispatch<React.SetStateAction<string>>;
  setConsolidatedSortAsc: React.Dispatch<React.SetStateAction<boolean>>;
  setExpandedConsolidatedEmail: React.Dispatch<React.SetStateAction<string>>;
  setMissingReminderKey: React.Dispatch<React.SetStateAction<string>>;
  setParticipantDetail: React.Dispatch<React.SetStateAction<{ name: string; email: string; jobTitle: string; location: string; company: string; department: string; phone: string; status: string; tid: number; }>>;
  setPersonalColsCollapsed: React.Dispatch<React.SetStateAction<boolean>>;
  setReminderBusyId: React.Dispatch<React.SetStateAction<number>>;
  setSelectedEvent: React.Dispatch<React.SetStateAction<DeloitteEvent>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  stripLocPrefix: (loc: string) => string;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
}

export const ConsolidatedView: React.FC<ConsolidatedViewProps> = (p) => {
  const { addAllToKlammer, addingToKlammer, addToKlammer, bulkKlammerProgress, colToggleHover, confirmDialog, consolidatedChildren, consolidatedFiltered, consolidatedRows, consolidatedSort, consolidatedSortAsc, expandedConsolidatedEmail, highlightMatch, inactiveAccounts, isAdmin, isConsolidatedMode, isDe, isLoadingSubEventRegs, isOrganizerFor, missingReminderKey, openDeregModal, openMainFieldsEdit, orgPastLock, performSilentDuplicateDelete, personalColsCollapsed, registrations, reminderBusyId, searchQuery, selectedEvent, sendCompleteRegistrationReminder, setAssignAssistRow, setAssignAssistValue, setColToggleHover, setConsolidatedSort, setConsolidatedSortAsc, setExpandedConsolidatedEmail, setMissingReminderKey, setParticipantDetail, setPersonalColsCollapsed, setReminderBusyId, setSelectedEvent, showAlert, stripLocPrefix, subEventRegsByEventId } = p;
    if (!selectedEvent) return null;
    if (isLoadingSubEventRegs) {
      return <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{isDe ? 'Lade Sub-Event-Teilnehmer...' : 'Loading sub-event participants...'}</p>;
    }
    if (consolidatedRows.length === 0) {
      return <p style={{ color: 'var(--dex-gray-400)' }}>{isDe ? 'Noch keine Anmeldungen in den Sub-Events.' : 'No registrations in the sub-events yet.'}</p>;
    }
    // v14.11: pastel A = event-level (parent) fields, pastel B = sub-event-specific fields
    const PASTEL_A_HEADER: React.CSSProperties = { background: 'rgba(0, 118, 168, 0.15)' };
    const PASTEL_A_CELL: React.CSSProperties = { background: 'rgba(0, 118, 168, 0.08)' };
    const PASTEL_B_HEADER: React.CSSProperties = { background: 'rgba(255, 191, 0, 0.18)' };
    const PASTEL_B_CELL: React.CSSProperties = { background: 'rgba(255, 191, 0, 0.10)' };
    const parentCustomFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim());
    // v23.32: People-Picker-Felder des Hauptevents (z.B. „Assistenz") werden
    // jetzt als eigene Spalten mit Foto + Name gezeigt (vorher ganz ausgeblendet).
    const parentUserFields = (selectedEvent.eventSpecificFields || []).filter(f => (f.type === 'user' || f.type === 'roommate') && f.label && f.label.trim());
    const parentIds = new Set(parentCustomFields.map(f => f.id));
    const childCustomFieldsByChild: Array<{ child: DeloitteEvent; fields: typeof parentCustomFields }> = consolidatedChildren.map(c => {
      const own = (c.eventSpecificFields || []).filter(f => f.type !== 'user' && f.label && f.label.trim() && !parentIds.has(f.id));
      return { child: c, fields: own };
    });
    // v30.17: Spalten-Zustand je Termin — vergangene Tage und (bei aktiver
    // „Anmeldung ab"-Regel) noch nicht anmeldbare Tage werden in Kopf,
    // Summenzeile und Zellen gedimmt. Dieselbe openFrom-Rechnung wie auf der
    // Anmeldeseite und in „Meine Events" (fixed/day/week) — keine zweite
    // Logik aufmachen.
    const childOpenFrom = (c: DeloitteEvent): Date | null => {
      const rule = selectedEvent.subEventOpenRule;
      if (!rule) return null;
      if (rule.mode === 'fixed') {
        const d = new Date(rule.date || '');
        return isFinite(d.getTime()) ? d : null;
      }
      if (!((rule.days || 0) > 0)) return null;
      const base = new Date(c.startDate || '');
      if (!isFinite(base.getTime())) return null;
      const dd = new Date(base.getFullYear(), base.getMonth(), base.getDate());
      if (rule.mode === 'week') dd.setDate(dd.getDate() - ((dd.getDay() + 6) % 7));
      dd.setDate(dd.getDate() - (rule.days || 0));
      dd.setHours(0, 0, 0, 0);
      return dd;
    };
    const hasOpenRule = !!selectedEvent.subEventOpenRule;
    const childColState: Record<string, { past: boolean; notYetOpen: boolean; openFrom: Date | null }> = {};
    for (const { child } of childCustomFieldsByChild) {
      const opensAt = childOpenFrom(child);
      childColState[child.id] = { past: isEventOver(child), notYetOpen: !!opensAt && new Date() < opensAt, openFrom: opensAt };
    }
    const dimColStyle = (id: string): React.CSSProperties =>
      (childColState[id] && (childColState[id].past || childColState[id].notYetOpen)) ? { opacity: 0.45 } : {};
    const handleSortConsolidated = (key: string): void => {
      if (consolidatedSort === key) setConsolidatedSortAsc(!consolidatedSortAsc);
      else { setConsolidatedSort(key); setConsolidatedSortAsc(true); }
    };
    const sortArrow = (key: string): string => key === consolidatedSort ? (consolidatedSortAsc ? ' ▲' : ' ▼') : '';
    const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];
    const abbreviate = (s: string, max: number): string => s.length > max ? s.substring(0, max - 1) + '…' : s;
    // v23.5: 6 Personen-Spalten (#, Vorname, Nachname, Email, Job Title,
    // Standort) — eingeklappt nur 2 (#, „Teilnehmer").
    const personalColCount = personalColsCollapsed ? 2 : 6;
    // v26.84: +1 zusätzliche Spalte „Registriert von" (Akteur) neben „Details".
    const totalColSpan = personalColCount + parentCustomFields.length + parentUserFields.length + childCustomFieldsByChild.reduce((sum, x) => sum + 1 + x.fields.length, 0) + 2;
    // v19.30: Aktionen (Hauptevent-Felder bearbeiten / abmelden) nur für
    // berechtigte Rollen (Admin oder Organizer dieses Events).
    const canManage = isAdmin || isOrganizerFor(selectedEvent);
    // v19.30 (Feature A): Anzahl der bearbeitbaren Hauptevent-Felder (ohne
    // People-Picker und Dokument-Uploads, die keinen editierbaren Textwert
    // haben). Nur wenn > 0 erscheint der „Felder"-Button.
    const editableParentFieldCount = parentCustomFields.filter(f => f.type !== 'document').length;
    // v19.30 (Feature A): Hat die Person eine Registrierung auf der
    // Hauptevent-Teilnehmerliste? Nur dann gibt es Hauptevent-Antworten zum
    // Bearbeiten. (Im subEventsOnlyMode kann jemand nur in Sub-Events
    // angemeldet sein.)
    const hasParentReg = (emailKey: string): boolean =>
      registrations.some(r => (r.ParticipantEmail || '').toLowerCase().trim() === emailKey);
    // v26.85: Ist ein bestimmtes Hauptevent-Feld für diese Person befüllt?
    // Gleiche Auflösung wie die Tabelle (Parent-Reg-Spalte → Parent-CustomData
    // → Sub-Event-CustomData-Fallback), damit die „Infos fehlen"-Erkennung nicht
    // fälschlich anschlägt.
    const parentFieldFilled = (row: ConsolidatedRow, f: { id: string; spInternalName?: string }): boolean => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const parentReg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey) as any;
      const spName = (f as { spInternalName?: string }).spInternalName || '';
      if (parentReg) {
        let v: unknown = spName ? parentReg[spName] : undefined;
        if ((v === undefined || v === null || v === '') && parentReg.CustomData) { try { v = JSON.parse(parentReg.CustomData)[f.id]; } catch { /* */ } }
        if (v !== undefined && v !== null && v !== '') return true;
      }
      for (const ch of consolidatedChildren) {
        const r = row.perChild[ch.id];
        if (!r) continue;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let v: any = spName ? (r as any)[spName] : undefined;
        if ((v === undefined || v === null || v === '') && r.CustomData) { try { v = JSON.parse(r.CustomData)[f.id]; } catch { /* */ } }
        if (v !== undefined && v !== null && v !== '') return true;
      }
      return false;
    };
    // v26.85: Personen mit Klammer-Anmeldung (Hauptevent) + Sub-Event, bei denen
    // aber PFLICHT-Hauptevent-Felder leer sind (typisch: Anmeldung im
    // Hauptevent-Schritt abgebrochen). requiredMainFields leer → nie anschlagen.
    const requiredMainFields = (selectedEvent.eventSpecificFields || []).filter(f => f.type !== 'user' && f.required && f.label && f.label.trim());
    // v23.7: „Verwaiste" Klammer-/Schatten-Anmeldungen erkennen — eine aktive
    // Zeile auf der KLAMMER-Subsite (registrations), deren Person in KEINEM
    // Sub-Event aktiv angemeldet ist. Das ist typischerweise ein Geist aus einer
    // abgebrochenen Anmeldung: unsichtbar in der Matrix (die nur Sub-Event-
    // Anmeldungen zeigt), blockiert aber jede neue (auch stellvertretende)
    // Anmeldung, weil die Doppel-Anmelde-Prüfung die Klammer-Subsite liest.
    // v23.7: Geist = aktive Klammer-Zeile, deren Person in KEINEM Sub-Event
    // IRGENDEINE Zeile hat (auch keine abgemeldete). Wer sich aus allen
    // Sub-Events ABGEMELDET hat, hat dort abgemeldete Zeilen → wird bewusst
    // NICHT als Geist erkannt (das ist eine normale Voll-Abmeldung, kein Rest
    // aus einer abgebrochenen Anmeldung). So vermeiden wir Fehlalarme.
    const anySubEmails = new Set<string>();
    for (const ch of consolidatedChildren) {
      for (const r of (subEventRegsByEventId[ch.id] || [])) {
        const em = (r.ParticipantEmail || '').toLowerCase().trim();
        if (em) anySubEmails.add(em);
      }
    }
    const orphanShadowRegs = registrations.filter(r => {
      if ((r.Status || '') === 'Abgemeldet') return false;
      const em = (r.ParticipantEmail || '').toLowerCase().trim();
      return !!em && !anySubEmails.has(em);
    });
    // v26.68: Bei aktiver Suche eine zusätzliche „ID"-Spalte mit der echten
    // TeilnehmerID aus der Klammer-/Hauptevent-Liste einblenden. Die „#"-Spalte
    // bleibt die laufende Durchzählung der aktuellen Ansicht — die ID hilft, die
    // gefilterte Person eindeutig in der SharePoint-Liste wiederzufinden.
    const searchActive = !!(searchQuery || '').trim();
    const bracketTidByEmail: Record<string, number> = {};
    for (const rr of registrations) {
      const em = (rr.ParticipantEmail || '').toLowerCase().trim();
      if (em && typeof rr.TeilnehmerID === 'number' && !(em in bracketTidByEmail)) bracketTidByEmail[em] = rr.TeilnehmerID;
    }
    return (
      // v28.53: Eigener Scroll-Container mit Höhenbegrenzung — analog zur
      // Sub-Event-Teilnehmerliste (renderTable, maxHeight 70vh). Vorher hatte
      // der Klammer-View nur overflowX, lief also über die volle Zeilenhöhe
      // inline mit: Bei 400+ Teilnehmern musste man an der ganzen Tabelle
      // vorbeiscrollen, und der sticky-thead hätte keinen Bezugsrahmen.
      <div style={{ maxHeight: '70vh', overflow: 'auto' }}>
        {/* v23.7: Unvollständige Klammer-Anmeldungen (nur Klammer, kein Sub-Event)
            sichtbar machen — mit Erinnerungs- oder Entfernen-Option, damit eine
            blockierte (Neu-)Anmeldung wieder möglich wird. */}
        {orphanShadowRegs.length > 0 && (
          <div style={{ marginBottom: 16, padding: 14, borderRadius: 10, border: '1px solid var(--dex-red, #c00)', background: 'rgba(200,0,0,0.06)' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
              <Icon iconName="Warning" style={{ fontSize: 16, color: 'var(--dex-red, #c00)' }} />
              <strong style={{ color: 'var(--dex-red, #c00)', fontSize: '0.9rem' }}>
                {isDe ? `Unvollständige Anmeldungen (${orphanShadowRegs.length})` : `Incomplete registrations (${orphanShadowRegs.length})`}
              </strong>
            </div>
            <p style={{ margin: '0 0 10px', fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
              {isDe
                ? 'Diese Personen haben eine Klammer-Anmeldung, sind aber in keinem Sub-Event aktiv angemeldet — in der Regel ein unvollständiger Rest aus einer abgebrochenen Anmeldung. Solche Rest-Anmeldungen erscheinen in der Teilnehmerliste unten nicht, blockieren aber eine erneute (auch stellvertretende) Anmeldung. Du hast zwei Möglichkeiten: über „Erinnerung senden“ der Person – bzw. der Person, die sie angemeldet hat – einen Link zum Abschließen der Anmeldung schicken, oder die Rest-Anmeldung entfernen, sodass eine Neuanmeldung wieder möglich ist.'
                : 'These people have an umbrella registration but are not actively registered for any sub-event — usually an incomplete leftover from an interrupted registration. Such leftover registrations don’t appear in the participant list below, but they block a new (or on-behalf) registration. You have two options: use „Send reminder“ to send the person – or whoever registered them – a link to complete the registration, or remove the leftover registration so a new one becomes possible.'}
            </p>
            <p style={{ margin: '0 0 10px', padding: '8px 10px', borderRadius: 6, background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)', fontSize: '0.8rem', color: 'var(--dex-orange-dark, #b35a00)', lineHeight: 1.5 }}>
              {isDe
                ? <><strong>Hinweis:</strong> Nach dem Entfernen ist die Person <strong>nicht</strong> angemeldet. Informiere die Person – bzw. die Person, die sie angemeldet hat –, dass eine <strong>erneute Anmeldung</strong> nötig ist. Alternativ kannst du über <strong>„Erinnerung senden“</strong> direkt einen Link zum Abschließen verschicken.</>
                : <><strong>Note:</strong> After removal the person is <strong>not</strong> registered. Let the person – or whoever registered them – know that a <strong>new registration</strong> is required. Alternatively, use <strong>„Send reminder“</strong> to send a completion link directly.</>}
            </p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>
              {orphanShadowRegs.map((r, oi) => {
                const nm = (r.Vorname && r.Nachname) ? `${r.Vorname} ${r.Nachname}` : (r.ParticipantName || r.ParticipantEmail);
                // v23.8: Bei stellvertretender Anmeldung (RegisteredBy ≠ Teilnehmer)
                // zeigen, WER die Anmeldung durchgeführt hat — hilft, die Assistenz
                // bzw. den Organizer zu identifizieren.
                const actorEmail = (r.RegisteredByEmail || '').trim();
                const isProxy = !!actorEmail && actorEmail.toLowerCase() !== (r.ParticipantEmail || '').trim().toLowerCase();
                const actorLabel = (r.RegisteredByName || '').trim() || actorEmail;
                // v26.68: Profil-Zusatzinfos (Position, Standort, Unternehmen) —
                // liegen als Property auf der Registrierung, hier für Foto-Subline
                // und die Meta-Zeile genutzt.
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const anyR = r as any;
                const jobTitle = String(anyR.JobTitle || '').trim();
                const loc = stripLocPrefix(String(r.Location || ''));
                const subline = [jobTitle, loc, String(r.Company || '').trim()].filter(Boolean).join(' • ');
                return (
                  <div key={r.Id} style={{ display: 'flex', alignItems: 'flex-start', gap: 12, padding: '10px 0', borderTop: oi > 0 ? '1px solid rgba(200,0,0,0.14)' : 'none' }}>
                    {/* v26.68: Foto mit Hover-Kontaktkarte, wie in der Teilnehmerliste. */}
                    {(r.ParticipantEmail || '') ? (
                      <PersonContactHover email={r.ParticipantEmail || ''} name={nm} size={40} subline={subline} isDe={isDe} />
                    ) : null}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 3, minWidth: 0, flex: 1 }}>
                      <div style={{ fontSize: '0.86rem' }}>
                        <strong>{nm}</strong>
                        <span style={{ color: 'var(--dex-gray-500)', marginLeft: 8 }}>{r.ParticipantEmail}</span>
                      </div>
                      {(jobTitle || loc) && (
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>{[jobTitle, loc].filter(Boolean).join(' • ')}</div>
                      )}
                      <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)' }}>
                        {translateStatus(r.Status, isDe)}
                        {r.RegistrationDate ? <> · {isDe ? 'angemeldet am' : 'registered on'} {formatDate(r.RegistrationDate)}</> : null}
                        {' · '}
                        {isProxy
                          ? <>{isDe ? 'angemeldet durch' : 'registered by'} <strong>{actorLabel}</strong></>
                          : (isDe ? 'selbst angemeldet' : 'self-registered')}
                      </div>
                      {canManage && (
                        <div style={{ display: 'inline-flex', gap: 8, flexWrap: 'wrap', marginTop: 6 }}>
                          {/* v26.67 (A): Erinnerung senden — die Person (bzw. die
                              anmeldende Person) bitten, die Anmeldung abzuschließen,
                              statt sie nur zu entfernen. */}
                          <button
                            type="button"
                            className="btn btn-outline"
                            style={{ fontSize: '0.75rem', padding: '3px 10px', color: 'var(--dex-green-dark, #4a7c1f)', borderColor: 'var(--dex-green, #86bc25)' }}
                            disabled={reminderBusyId === r.Id}
                            onClick={async () => {
                              if (!selectedEvent) return;
                              setReminderBusyId(r.Id);
                              const ok = await sendCompleteRegistrationReminder({
                                eventId: selectedEvent.id,
                                eventTitle: selectedEvent.title,
                                participantEmail: r.ParticipantEmail || '',
                                participantName: nm,
                                registeredByEmail: r.RegisteredByEmail || '',
                                registeredByName: r.RegisteredByName || '',
                              }).catch(() => false);
                              setReminderBusyId(null);
                              showAlert(
                                ok
                                  ? (isProxy
                                      ? (isDe ? `Erinnerung an ${actorLabel} gesendet (${nm} auf Kopie) — mit Link zum Abschließen der Anmeldung.` : `Reminder sent to ${actorLabel} (${nm} on copy) — with a link to complete the registration.`)
                                      : (isDe ? `Erinnerung an ${nm} gesendet — mit Link zum Abschließen der Anmeldung.` : `Reminder sent to ${nm} — with a link to complete the registration.`))
                                  : (isDe ? 'Erinnerung konnte nicht gesendet werden.' : 'The reminder could not be sent.'),
                                { variant: ok ? 'success' : 'error' });
                            }}
                          >
                            {reminderBusyId === r.Id ? (isDe ? 'Wird gesendet…' : 'Sending…') : (isDe ? 'Erinnerung senden' : 'Send reminder')}
                          </button>
                          <button
                            type="button"
                            className="btn btn-outline"
                            style={{ fontSize: '0.75rem', padding: '3px 10px', color: 'var(--dex-red, #c00)', borderColor: 'var(--dex-red, #c00)' }}
                            onClick={async () => {
                              if (!(await confirmDialog(isDe ? `Unvollständige Anmeldung von ${nm} entfernen? Die Person kann sich danach neu anmelden.` : `Remove the incomplete registration of ${nm}? The person can register again afterwards.`, { danger: true, confirmLabel: isDe ? 'Entfernen' : 'Remove' }))) return;
                              await performSilentDuplicateDelete(r);
                              showAlert(isDe ? 'Rest-Anmeldung entfernt — die Person kann jetzt wieder angemeldet werden.' : 'Leftover registration removed — the person can be registered again now.', { variant: 'success' });
                            }}
                          >
                            {isDe ? 'Rest-Anmeldung entfernen' : 'Remove leftover'}
                          </button>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
        {/* v24.38: Fehlende Klammer-Anmeldung — Personen mit Sub-Event-Zeilen,
            aber OHNE Hauptevent-/Klammer-Anmeldung (Daten-Anomalie). Als Fehler
            ausweisen + pro Person „Zur Klammer hinzufügen". */}
        {canManage && (() => {
          /**
           * v30.56: ERST prüfen, ob die Klammer-Zeile nur unter einer ANDEREN
           * Schreibweise der Adresse steht — dann fehlt sie nämlich gar nicht.
           *
           * Der Kasten hat bis hier ausschließlich `ParticipantEmail` exakt
           * verglichen. Dieselbe Person kann aber unter zwei Adressen in den
           * Listen stehen (SMTP-Adresse gegen UPN/Alias) — die
           * Doppel-Adressen-Falle, die in CLAUDE.md als erster Verdacht bei
           * widersprüchlichen Ansichten steht. Trifft sie zu, meldete der
           * Kasten einen Fehler, den es nicht gibt.
           *
           * **Und der Reparatur-Knopf hätte ihn zu einem echten gemacht:**
           * „Zur Klammer hinzufügen" legt eine Zeile unter der Sub-Event-
           * Adresse an — die Person stünde danach ZWEIMAL am Hauptevent, mit
           * zwei Teilnehmer-IDs. Ein Knopf, der Daten repariert, darf keine
           * Dubletten erzeugen.
           *
           * Erkannt wird die zweite Schreibweise über den lokalen Teil der
           * Adresse (vor dem @) und über Vor-/Nachname. Beides ist eine
           * Heuristik — deshalb wird der Fall NICHT still geschluckt, sondern
           * getrennt ausgewiesen: Der Organizer sieht, unter welcher Adresse
           * die Zeile steht, und entscheidet selbst.
           */
          const normName = (v: string): string => (v || '').toLowerCase().replace(/[^a-zäöüß]/g, '');
          const localPart = (v: string): string => (v || '').toLowerCase().split('@')[0].trim();
          const altParentRegOf = (row: typeof consolidatedRows[number]): SPRegistration | undefined => {
            const lp = localPart(row.emailKey);
            const nm = normName(row.vorname) + normName(row.nachname);
            return registrations.find(r => {
              const rEmail = (r.ParticipantEmail || '').toLowerCase().trim();
              if (rEmail === row.emailKey) return false; // exakt wurde oben schon geprüft
              if (lp && localPart(rEmail) === lp) return true;
              const rNm = normName(r.Vorname || '') + normName(r.Nachname || '');
              return !!nm && rNm === nm;
            });
          };
          const flagged = consolidatedRows.filter(r => r.activeCount > 0 && !hasParentReg(r.emailKey));
          const aliasCases = flagged
            .map(r => ({ row: r, alt: altParentRegOf(r) }))
            .filter(x => !!x.alt) as Array<{ row: typeof consolidatedRows[number]; alt: SPRegistration }>;
          const aliasKeys = new Set(aliasCases.map(x => x.row.emailKey));
          const missing = flagged.filter(r => !aliasKeys.has(r.emailKey));
          if (missing.length === 0 && aliasCases.length === 0) return null;
          if (missing.length === 0) {
            // Nur Adress-Dubletten — kein Fehler, aber erklärungsbedürftig.
            return (
              <div style={{ marginBottom: 16, padding: 14, borderRadius: 10, border: '1px solid var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.07)' }}>
                <strong style={{ color: 'var(--dex-orange-dark, #b35a00)', fontSize: '0.9rem' }}>
                  {isDe ? `Klammer-Zeile unter anderer Adresse (${aliasCases.length})` : `Umbrella row under a different address (${aliasCases.length})`}
                </strong>
                <p style={{ margin: '6px 0 8px', fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                  {isDe
                    ? 'Diese Personen haben eine Klammer-Anmeldung — sie steht nur unter einer anderen Schreibweise ihrer E-Mail-Adresse (SMTP-Adresse gegen UPN/Alias). Es fehlt also nichts. Trag sie NICHT über „Zur Klammer hinzufügen" nach, das würde eine zweite Zeile mit einer zweiten Teilnehmer-ID erzeugen. Wenn die beiden Schreibweisen stören, korrigiere die Adresse in der Teilnehmerzeile.'
                    : 'These people do have an umbrella registration — it is just stored under a different spelling of their email address. Nothing is missing; do NOT use „Add to umbrella", it would create a duplicate row.'}
                </p>
                {aliasCases.map(x => (
                  <div key={x.row.emailKey} style={{ fontSize: '0.82rem', marginBottom: 2 }}>
                    <strong>{x.row.vorname} {x.row.nachname}</strong>{' '}
                    <span style={{ color: 'var(--dex-gray-600)' }}>Sub-Events: {x.row.email} · Klammer: {x.alt.ParticipantEmail}</span>
                  </div>
                ))}
              </div>
            );
          }
          return (
            <div style={{ marginBottom: 16, padding: 14, borderRadius: 10, border: '1px solid var(--dex-red, #c00)', background: 'rgba(200,0,0,0.06)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                <Icon iconName="Warning" style={{ fontSize: 16, color: 'var(--dex-red, #c00)' }} />
                <strong style={{ color: 'var(--dex-red, #c00)', fontSize: '0.9rem' }}>
                  {isDe ? `Fehlende Klammer-Anmeldung (${missing.length})` : `Missing umbrella registration (${missing.length})`}
                </strong>
              </div>
              <p style={{ margin: '0 0 10px', fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {isDe
                  ? 'Diese Personen sind in einem oder mehreren Sub-Events angemeldet, fehlen aber am Klammer-/Hauptevent selbst (z.B. durch eine abgebrochene Anmeldung). Dadurch fehlen u.a. die übergreifenden Hauptevent-Angaben. Du hast zwei Möglichkeiten: über „Erinnerung senden“ bittest du die Person (bzw. die anmeldende Person) per Mail mit Direkt-Link, die fehlenden Hauptevent-Angaben in der App nachzutragen — oder du trägst die fehlende Klammer-Anmeldung mit „Zur Klammer hinzufügen“ selbst nach (versendet KEINE Mail und KEINEN Outlook-Termin, reine Datenkorrektur).'
                  : 'These people are registered for one or more sub-events but are missing on the umbrella/main event itself (e.g. due to an interrupted registration), so the cross-cutting main-event details are missing. You have two options: use „Send reminder“ to ask the person (or whoever registered them) via email with a direct link to add the missing main-event details in the app — or add the missing umbrella registration yourself with „Add to umbrella“ (sends NO email and NO Outlook invite, data correction only).'}
              </p>
              {/* v30.56: Adress-Dubletten auch hier benennen, wenn es
                  DANEBEN echte Lücken gibt — sonst verschwinden sie
                  kommentarlos aus dem Kasten und der Organizer fragt sich,
                  wo die dritte Person geblieben ist. */}
              {aliasCases.length > 0 && (
                <p style={{ margin: '0 0 10px', padding: '8px 12px', borderRadius: 8, background: 'rgba(237,139,0,0.10)', fontSize: '0.8rem', color: 'var(--dex-orange-dark, #b35a00)', lineHeight: 1.5 }}>
                  {isDe
                    ? <>Nicht aufgeführt, weil dort nichts fehlt: {aliasCases.map(x => `${x.row.vorname} ${x.row.nachname}`).join(', ')} — die Klammer-Zeile steht unter einer anderen Schreibweise der Adresse ({aliasCases.map(x => x.alt.ParticipantEmail).join(', ')}).</>
                    : <>Not listed because nothing is missing there: {aliasCases.map(x => `${x.row.vorname} ${x.row.nachname}`).join(', ')} — the umbrella row exists under a different spelling of the address.</>}
                </p>
              )}
              {/* v30.14: Sammel-Fix — alle auf einmal, still, sequentiell. */}
              <div style={{ marginBottom: 10 }}>
                <button
                  type="button"
                  className="btn btn-outline"
                  style={{ fontSize: '0.8rem', padding: '5px 14px', color: 'var(--dex-orange, #ed8b00)', borderColor: 'var(--dex-orange, #ed8b00)' }}
                  disabled={!!bulkKlammerProgress || !!addingToKlammer}
                  onClick={() => { void addAllToKlammer(missing); }}
                >
                  <Plus size={12} />{' '}
                  {bulkKlammerProgress
                    ? (isDe ? `Wird nachgetragen… (${bulkKlammerProgress})` : `Adding… (${bulkKlammerProgress})`)
                    : (isDe ? `Alle ${missing.length} still zur Klammer hinzufügen` : `Silently add all ${missing.length} to the umbrella`)}
                </button>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {missing.map(r => {
                  const nm = `${r.vorname || ''} ${r.nachname || ''}`.trim() || r.email;
                  // v26.85: Akteur (selbst/stellvertretend) aus den Sub-Event-
                  // Registrierungen ableiten — für die Reminder-Empfänger.
                  let byEmail = '', byName = '';
                  for (const ck of Object.keys(r.perChild)) {
                    const cr = r.perChild[ck];
                    if (cr && (cr.RegisteredByEmail || '').trim()) { byEmail = (cr.RegisteredByEmail || '').trim(); byName = (cr.RegisteredByName || '').trim(); break; }
                  }
                  const isProxy = !!byEmail && byEmail.toLowerCase() !== (r.email || '').toLowerCase();
                  return (
                    <div key={r.emailKey} style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', fontSize: '0.84rem' }}>
                      <strong>{nm}</strong>
                      <span style={{ color: 'var(--dex-gray-500)' }}>{r.email}</span>
                      <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem' }}>· {isDe ? `${r.activeCount} Sub-Event(s)` : `${r.activeCount} sub-event(s)`}</span>
                      <div style={{ marginLeft: 'auto', display: 'inline-flex', gap: 8, flexWrap: 'wrap' }}>
                        {/* v26.85: Erinnerung senden — Person (bzw. Anmeldende:r) bitten,
                            die fehlenden Hauptevent-Angaben in der App nachzutragen. */}
                        <button
                          type="button"
                          className="btn btn-outline"
                          style={{ fontSize: '0.75rem', padding: '3px 10px', color: 'var(--dex-green-dark, #4a7c1f)', borderColor: 'var(--dex-green, #86bc25)' }}
                          disabled={missingReminderKey === r.emailKey}
                          onClick={async () => {
                            if (!selectedEvent) return;
                            setMissingReminderKey(r.emailKey);
                            const ok = await sendCompleteRegistrationReminder({
                              eventId: selectedEvent.id,
                              eventTitle: selectedEvent.title,
                              participantEmail: r.email || '',
                              participantName: nm,
                              registeredByEmail: byEmail,
                              registeredByName: byName,
                            }).catch(() => false);
                            setMissingReminderKey(null);
                            showAlert(
                              ok
                                ? (isProxy
                                    ? (isDe ? `Erinnerung an ${byName || byEmail} gesendet (${nm} auf Kopie) — mit Link zum Nachtragen der Hauptevent-Angaben.` : `Reminder sent to ${byName || byEmail} (${nm} on copy) — with a link to add the main-event details.`)
                                    : (isDe ? `Erinnerung an ${nm} gesendet — mit Link zum Nachtragen der Hauptevent-Angaben.` : `Reminder sent to ${nm} — with a link to add the main-event details.`))
                                : (isDe ? 'Erinnerung konnte nicht gesendet werden.' : 'The reminder could not be sent.'),
                              { variant: ok ? 'success' : 'error' });
                          }}
                        >
                          {missingReminderKey === r.emailKey ? (isDe ? 'Wird gesendet…' : 'Sending…') : (isDe ? 'Erinnerung senden' : 'Send reminder')}
                        </button>
                        <button
                          type="button"
                          className="btn btn-outline"
                          style={{ fontSize: '0.75rem', padding: '3px 10px', color: 'var(--dex-orange, #ed8b00)', borderColor: 'var(--dex-orange, #ed8b00)' }}
                          disabled={addingToKlammer === r.emailKey}
                          onClick={() => { void addToKlammer(r); }}
                        >
                          <Plus size={12} /> {addingToKlammer === r.emailKey ? '…' : (isDe ? 'Zur Klammer hinzufügen' : 'Add to umbrella')}
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })()}
        {/* v26.85: Klammer-Anmeldung vorhanden (+ Sub-Event), aber PFLICHT-
            Hauptevent-Angaben fehlen — typisch nach einer abgebrochenen
            Anmeldung. Oben als Hinweis, mit „Erinnerung senden" (Person bzw.
            anmeldende Person bitten, die Angaben in der App nachzutragen) und
            „Hauptevent-Felder bearbeiten" (selbst nachtragen). */}
        {canManage && requiredMainFields.length > 0 && (() => {
          const incomplete = consolidatedRows.filter(row => hasParentReg(row.emailKey) && row.activeCount > 0 && requiredMainFields.some(f => !parentFieldFilled(row, f)));
          if (incomplete.length === 0) return null;
          return (
            <div style={{ marginBottom: 16, padding: 14, borderRadius: 10, border: '1px solid var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.07)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                <Icon iconName="Warning" style={{ fontSize: 16, color: 'var(--dex-orange-dark, #b35a00)' }} />
                <strong style={{ color: 'var(--dex-orange-dark, #b35a00)', fontSize: '0.9rem' }}>
                  {isDe ? `Unvollständige Hauptevent-Angaben (${incomplete.length})` : `Incomplete main-event details (${incomplete.length})`}
                </strong>
              </div>
              <p style={{ margin: '0 0 10px', fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {isDe
                  ? 'Diese Personen haben eine Anmeldung (Hauptevent + Sub-Event), es fehlen aber Pflicht-Angaben aus dem Hauptevent-Schritt — meist, weil die Anmeldung vorzeitig abgebrochen wurde. Über „Erinnerung senden“ bittest du die Person (bzw. die anmeldende Person) per Mail mit Direkt-Link, die fehlenden Angaben in der App nachzutragen. Alternativ kannst du sie über „Hauptevent-Felder bearbeiten“ direkt selbst ergänzen.'
                  : 'These people have a registration (main event + sub-event), but required answers from the main-event step are missing — usually because the registration was interrupted. Use „Send reminder“ to ask the person (or whoever registered them) via email with a direct link to add the missing answers in the app. Alternatively, add them yourself via „Edit main-event fields“.'}
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {incomplete.map(row => {
                  const nm = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email;
                  const missingLabels = requiredMainFields.filter(f => !parentFieldFilled(row, f)).map(f => f.label);
                  let byEmail = '', byName = '';
                  for (const ck of Object.keys(row.perChild)) {
                    const cr = row.perChild[ck];
                    if (cr && (cr.RegisteredByEmail || '').trim()) { byEmail = (cr.RegisteredByEmail || '').trim(); byName = (cr.RegisteredByName || '').trim(); break; }
                  }
                  const isProxy = !!byEmail && byEmail.toLowerCase() !== (row.email || '').toLowerCase();
                  return (
                    <div key={row.emailKey} style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', fontSize: '0.84rem' }}>
                      <strong>{nm}</strong>
                      <span style={{ color: 'var(--dex-gray-500)' }}>{row.email}</span>
                      <span style={{ color: 'var(--dex-orange-dark, #b35a00)', fontSize: '0.76rem' }} title={missingLabels.join(', ')}>· {isDe ? 'fehlt: ' : 'missing: '}{missingLabels.slice(0, 3).join(', ')}{missingLabels.length > 3 ? ` +${missingLabels.length - 3}` : ''}</span>
                      <div style={{ marginLeft: 'auto', display: 'inline-flex', gap: 8, flexWrap: 'wrap' }}>
                        <button
                          type="button"
                          className="btn btn-outline"
                          style={{ fontSize: '0.75rem', padding: '3px 10px', color: 'var(--dex-green-dark, #4a7c1f)', borderColor: 'var(--dex-green, #86bc25)' }}
                          disabled={missingReminderKey === row.emailKey}
                          onClick={async () => {
                            if (!selectedEvent) return;
                            setMissingReminderKey(row.emailKey);
                            const ok = await sendCompleteRegistrationReminder({
                              eventId: selectedEvent.id,
                              eventTitle: selectedEvent.title,
                              participantEmail: row.email || '',
                              participantName: nm,
                              registeredByEmail: byEmail,
                              registeredByName: byName,
                            }).catch(() => false);
                            setMissingReminderKey(null);
                            showAlert(
                              ok
                                ? (isProxy
                                    ? (isDe ? `Erinnerung an ${byName || byEmail} gesendet (${nm} auf Kopie) — mit Link zum Nachtragen der Angaben.` : `Reminder sent to ${byName || byEmail} (${nm} on copy) — with a link to add the details.`)
                                    : (isDe ? `Erinnerung an ${nm} gesendet — mit Link zum Nachtragen der Angaben.` : `Reminder sent to ${nm} — with a link to add the details.`))
                                : (isDe ? 'Erinnerung konnte nicht gesendet werden.' : 'The reminder could not be sent.'),
                              { variant: ok ? 'success' : 'error' });
                          }}
                        >
                          {missingReminderKey === row.emailKey ? (isDe ? 'Wird gesendet…' : 'Sending…') : (isDe ? 'Erinnerung senden' : 'Send reminder')}
                        </button>
                        <button
                          type="button"
                          className="btn btn-outline"
                          style={{ fontSize: '0.75rem', padding: '3px 10px', color: 'var(--dex-orange, #ed8b00)', borderColor: 'var(--dex-orange, #ed8b00)' }}
                          onClick={() => openMainFieldsEdit(row.emailKey, nm)}
                        >
                          {isDe ? 'Hauptevent-Felder bearbeiten' : 'Edit main-event fields'}
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })()}
        {/* v29.2: Dieselbe Person unter ZWEI E-Mail-Adressen.
            Die Matrix aggregiert strikt nach ParticipantEmail (lowercase).
            Steht jemand in einem Sub-Event unter der SMTP-Adresse und im
            anderen unter der UPN-/Alias-Adresse, entstehen ZWEI Zeilen — und
            jede zeigt beim jeweils anderen Sub-Event „—". Von außen sieht das
            aus, als sei die Person „nicht angemeldet", obwohl die Anmeldung
            existiert; andere Ansichten (z.B. die Hotelplanung, die über die
            Klammer-Zeile matcht) zeigen dann den Haken. Dass die beiden
            Schreibweisen auseinanderlaufen können, ist in dieser Codebasis
            belegt — siehe canRegisterForOthers in EventService, das dieselbe
            Person bewusst über pageContext.user.email UND die E-Mail aus dem
            loginName sucht. */}
        {canManage && (() => {
          const normName = (s: string): string => s.toLowerCase().replace(/\s+/g, ' ').trim();
          const byName: Record<string, ConsolidatedRow[]> = {};
          for (const r of consolidatedRows) {
            const key = normName(`${r.vorname || ''} ${r.nachname || ''}`);
            if (!key) continue;
            (byName[key] = byName[key] || []).push(r);
          }
          // consolidatedRows ist bereits pro E-Mail eindeutig — mehr als eine
          // Zeile zum selben Namen heißt also: verschiedene Adressen.
          const groups = Object.keys(byName).map(k => byName[k]).filter(rows => rows.length > 1);
          if (groups.length === 0) return null;
          const subsOfRow = (r: ConsolidatedRow): string =>
            consolidatedChildren
              .filter(ch => { const cr = r.perChild[ch.id]; return !!cr && ACTIVE.indexOf(cr.Status) >= 0; })
              .map(ch => shortSubEventTitle(ch.title, selectedEvent?.title))
              .join(', ') || (isDe ? 'keine' : 'none');
          return (
            <div style={{ marginBottom: 16, padding: 14, borderRadius: 10, border: '1px solid var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.07)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                <Icon iconName="Warning" style={{ fontSize: 16, color: 'var(--dex-orange-dark, #b35a00)' }} />
                <strong style={{ color: 'var(--dex-orange-dark, #b35a00)', fontSize: '0.9rem' }}>
                  {isDe ? `Gleiche Person, mehrere E-Mail-Adressen (${groups.length})` : `Same person, several email addresses (${groups.length})`}
                </strong>
              </div>
              <p style={{ margin: '0 0 10px', fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {isDe
                  ? 'Diese Namen kommen in der Tabelle mehrfach vor — jeweils mit einer anderen E-Mail-Adresse. Die Teilnehmerliste fasst pro E-Mail zusammen, deshalb wird die Person auf zwei Zeilen aufgeteilt und jede Zeile zeigt beim Sub-Event der anderen Zeile ein „—“. Die Anmeldungen selbst sind vorhanden. Typische Ursache: Die eine Anmeldung lief über die Anmeldeseite (SMTP-Adresse), die andere stellvertretend über die Personenauswahl (UPN-/Alias-Adresse). Prüfe unten, welche Adresse die richtige ist, und melde die Person über die falsche Adresse ab und über die richtige neu an.'
                  : 'These names appear more than once in the table — each with a different email address. The participant list aggregates per email, so the person is split across two rows and each row shows a „—“ for the other row’s sub-event. The registrations themselves exist. Typical cause: one registration came from the registration page (SMTP address), the other on-behalf via the people picker (UPN/alias address). Check below which address is the correct one, then cancel the registration on the wrong address and re-register on the correct one.'}
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {groups.map(rows => (
                  <div key={rows.map(r => r.emailKey).join('|')} style={{ fontSize: '0.84rem' }}>
                    <strong>{`${rows[0].vorname || ''} ${rows[0].nachname || ''}`.trim()}</strong>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, marginTop: 2 }}>
                      {rows.map(r => (
                        <div key={r.emailKey} style={{ display: 'flex', gap: 10, flexWrap: 'wrap', fontSize: '0.8rem' }}>
                          <span style={{ color: 'var(--dex-gray-600)', minWidth: 260 }}>{r.email}</span>
                          <span style={{ color: 'var(--dex-gray-500)' }}>{isDe ? 'angemeldet für: ' : 'registered for: '}{subsOfRow(r)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          );
        })()}
        {/* v15.3.1: Legende für die Pastell-Spalten — sonst rät der Organizer,
            was die zwei Hintergrundfarben bedeuten. */}
        {(parentCustomFields.length > 0 || childCustomFieldsByChild.some(x => x.fields.length > 0)) && (
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, marginBottom: 10, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
            {parentCustomFields.length > 0 && (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                <span style={{ width: 14, height: 14, borderRadius: 3, ...PASTEL_A_HEADER, border: '1px solid rgba(0, 118, 168, 0.3)' }} />
                {isDe ? 'Felder des Hauptevents' : 'Main-event fields'}
              </span>
            )}
            {childCustomFieldsByChild.some(x => x.fields.length > 0) && (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                <span style={{ width: 14, height: 14, borderRadius: 3, ...PASTEL_B_HEADER, border: '1px solid rgba(255, 191, 0, 0.4)' }} />
                {isDe ? 'Felder eines Sub-Events' : 'Sub-event fields'}
              </span>
            )}
          </div>
        )}
        {/* v26.84: minWidth max-content, damit die Tabelle bei vielen Spalten
            NICHT gestaucht wird, sondern über den overflowX:auto-Wrapper (oben)
            horizontal scrollbar wird — wie in der Sub-Event-Teilnehmerliste. */}
        <table style={{ width: '100%', minWidth: 'max-content', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
          {/* v30.38: Kopf bleibt beim Scrollen stehen. Bei 19 Termin-Spalten
              und 77 Zeilen war nach wenigen Zeilen nicht mehr erkennbar, zu
              welchem Tag eine Haken-Spalte gehört — man musste hochscrollen,
              zählen, zurückscrollen. Sticky sitzt am `<thead>` und nicht an den
              einzelnen `<tr>`: Die drei Kopfzeilen (Spaltentitel, „∑ angemeldet",
              „Anmeldung ab") sind unterschiedlich hoch, für zeilenweises Sticky
              müsste man die Offsets messen. Dieselbe Lösung wie in der
              Sub-Event-Teilnehmerliste (`renderTable`).
              `background` ist Pflicht — ohne ihn scrollen die Datenzeilen
              sichtbar durch den Kopf hindurch. Die Pastell-Kopfzellen setzen
              ihren eigenen Hintergrund inline und gewinnen dadurch. */}
          {/* Der Schatten ersetzt die Kopf-Unterkante: Bei `border-collapse:
              collapse` bleiben die Rahmen der Kopfzeilen beim Ankleben zurück
              (Browser-Verhalten, nicht abstellbar) — ohne ihn schwebt der Kopf
              ohne Abgrenzung über den Daten. */}
          <thead style={{ position: 'sticky', top: 0, zIndex: 2, background: 'var(--dex-gray-50, #fafafa)', boxShadow: '0 2px 4px rgba(0,0,0,0.08)' }}>
            <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
              {/* v26.65: Header-Tooltip stellt klar, dass „#" die laufende Zeilen-
                  nummer dieser Ansicht ist — NICHT die Teilnehmer-ID der SharePoint-
                  Liste (die pro Sub-Event unterschiedlich ist). Sortiert nach
                  Erst-Anmeldung. */}
              <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('id')}
                title={isDe ? 'Laufende Nummer in dieser Ansicht (nicht die Teilnehmer-ID der Liste — die ist pro Sub-Event unterschiedlich)' : 'Row number in this view (not the SharePoint participant ID — that differs per sub-event)'}>#{sortArrow('id')}</th>
              {/* v26.68: echte Teilnehmer-ID (Klammer-Liste) — nur bei aktiver Suche. */}
              {searchActive && (
                <th style={{ textAlign: 'left', padding: 8, verticalAlign: 'bottom' }} title={isDe ? 'Teilnehmer-ID in der Hauptevent-/Klammer-Teilnehmerliste' : 'Participant ID in the main-event / bracket list'}>ID</th>
              )}
              {personalColsCollapsed ? (
                // v26.65 BUG-FIX: Sortier-Klick auf das GANZE <th> (vorher nur auf
                // den kleinen Text-<span> — daneben klicken sortierte nicht).
                // v30.21: Der Klapp-Knopf sitzt jetzt als beschriftete Pille in
                // einer EIGENEN Zeile ÜBER der Spaltenüberschrift — der kleine
                // runde Knopf lag direkt neben dem Sortier-Klickziel, ein leicht
                // versetzter Klick sortierte statt zu klappen (Nutzer-Befund).
                // Hover-Effekt über colToggleHover (Inline-Styles können kein :hover).
                <th style={{ textAlign: 'left', padding: 8, userSelect: 'none', whiteSpace: 'nowrap', cursor: 'pointer', verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('nachname')}>
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 3 }}>
                    <button
                      type="button"
                      onClick={(e) => { e.stopPropagation(); setPersonalColsCollapsed(false); setColToggleHover(false); }}
                      onMouseEnter={() => setColToggleHover(true)}
                      onMouseLeave={() => setColToggleHover(false)}
                      title={isDe ? 'Vorname, Nachname, E-Mail, Job Title, Standort und Unternehmen als eigene Spalten anzeigen' : 'Show first/last name, email, job title, location and company as separate columns'}
                      style={{
                        display: 'inline-flex', alignItems: 'center', gap: 4,
                        padding: '2px 10px', borderRadius: 999,
                        border: '1px solid var(--dex-green, #86bc25)',
                        background: colToggleHover ? 'var(--dex-green, #86bc25)' : '#fff',
                        color: colToggleHover ? '#fff' : 'var(--dex-green-dark, #4a7c1f)',
                        fontSize: '0.7rem', fontWeight: 700, cursor: 'pointer', lineHeight: 1.4,
                        transition: 'background 120ms ease, color 120ms ease',
                      }}
                    >» {isDe ? 'Aufklappen' : 'Expand'}</button>
                    <span>{isDe ? 'Teilnehmer' : 'Participant'}{sortArrow('nachname')}</span>
                  </div>
                </th>
              ) : (
                <>
                  <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('vorname')}>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 3 }}>
                      <button
                        type="button"
                        onClick={(e) => { e.stopPropagation(); setPersonalColsCollapsed(true); setColToggleHover(false); }}
                        onMouseEnter={() => setColToggleHover(true)}
                        onMouseLeave={() => setColToggleHover(false)}
                        title={isDe ? 'Personen-Spalten einklappen (nur Foto + Name)' : 'Collapse personal columns (photo + name only)'}
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: 4,
                          padding: '2px 10px', borderRadius: 999,
                          border: '1px solid var(--dex-green, #86bc25)',
                          background: colToggleHover ? 'var(--dex-green, #86bc25)' : '#fff',
                          color: colToggleHover ? '#fff' : 'var(--dex-green-dark, #4a7c1f)',
                          fontSize: '0.7rem', fontWeight: 700, cursor: 'pointer', lineHeight: 1.4,
                          transition: 'background 120ms ease, color 120ms ease',
                        }}
                      >« {isDe ? 'Zuklappen' : 'Collapse'}</button>
                      <span>{isDe ? 'Vorname' : 'First name'}{sortArrow('vorname')}</span>
                    </div>
                  </th>
                  <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('nachname')}>{isDe ? 'Nachname' : 'Last name'}{sortArrow('nachname')}</th>
                  <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('email')}>Email{sortArrow('email')}</th>
                  <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('jobTitle')}>Job Title{sortArrow('jobTitle')}</th>
                  <th style={{ textAlign: 'left', padding: 8, cursor: 'pointer', userSelect: 'none', verticalAlign: 'bottom' }} onClick={() => handleSortConsolidated('location')}>{isDe ? 'Standort' : 'Location'}{sortArrow('location')}</th>
                  <th style={{ textAlign: 'left', padding: 8, verticalAlign: 'bottom' }}>{isDe ? 'Unternehmen' : 'Company'}</th>
                </>
              )}
              {/* v26.84: „Registriert von" auch im Klammer-View — selbst /
                  Assistenz / stellvertretend. */}
              <th style={{ textAlign: 'left', padding: 8, whiteSpace: 'nowrap', verticalAlign: 'bottom' }}>{isDe ? 'Registriert von' : 'Registered by'}</th>
              {parentCustomFields.map(f => (
                <th key={`pf-${f.id}`} onClick={() => handleSortConsolidated(`pf:${f.id}`)} style={{ textAlign: 'left', padding: 8, fontSize: '0.78rem', whiteSpace: 'normal', overflowWrap: 'break-word', maxWidth: 150, verticalAlign: 'top', lineHeight: 1.25, cursor: 'pointer', userSelect: 'none', ...PASTEL_A_HEADER }} title={`${f.label} — ${isDe ? 'Hauptevent-Feld' : 'main-event field'}`}>
                  {f.label}{sortArrow(`pf:${f.id}`)}
                </th>
              ))}
              {/* v23.32: People-Picker-Felder des Hauptevents (Foto + Name). */}
              {parentUserFields.map(f => (
                <th key={`puf-${f.id}`} style={{ textAlign: 'left', padding: 8, fontSize: '0.78rem', whiteSpace: 'normal', overflowWrap: 'break-word', maxWidth: 170, verticalAlign: 'top', lineHeight: 1.25, ...PASTEL_A_HEADER }} title={`${f.label} — ${isDe ? 'Hauptevent-Feld' : 'main-event field'}`}>
                  {f.label}
                </th>
              ))}
              {childCustomFieldsByChild.map(({ child, fields }) => (
                <React.Fragment key={`sub-${child.id}`}>
                  <th
                    style={{ textAlign: 'center', padding: 8, cursor: 'pointer', userSelect: 'none', borderLeft: '1px solid var(--dex-gray-200)', ...dimColStyle(child.id) }}
                    onClick={() => handleSortConsolidated(`child:${child.id}`)}
                    title={child.title}
                  >
                    <div style={{ fontSize: '0.78rem', fontWeight: 600 }}>{abbreviate(shortSubEventTitle(child.title, selectedEvent?.title) || '?', 16)}</div>
                    <div style={{ fontSize: '0.68rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>{isDe ? 'angemeldet?' : 'registered?'}{sortArrow(`child:${child.id}`)}</div>
                  </th>
                  {fields.map(f => (
                    <th key={`scf-${child.id}-${f.id}`} onClick={() => handleSortConsolidated(`cf:${child.id}|${f.id}`)} style={{ textAlign: 'left', padding: 8, fontSize: '0.78rem', whiteSpace: 'normal', overflowWrap: 'break-word', maxWidth: 150, verticalAlign: 'top', lineHeight: 1.25, cursor: 'pointer', userSelect: 'none', ...PASTEL_B_HEADER }} title={`${f.label} — ${child.title}`}>
                      <div style={{ color: 'var(--dex-gray-500)', fontWeight: 400, fontSize: '0.68rem' }}>{abbreviate(shortSubEventTitle(child.title, selectedEvent?.title) || '?', 18)}</div>
                      <div style={{ fontWeight: 600 }}>{f.label}{sortArrow(`cf:${child.id}|${f.id}`)}</div>
                    </th>
                  ))}
                </React.Fragment>
              ))}
              <th style={{ textAlign: 'left', padding: 8, verticalAlign: 'bottom' }}>{isDe ? 'Details' : 'Details'}</th>
            </tr>
            {/* v30.15: Summenzeile je Termin-Spalte — bei einer Office-Tage-
                Reihe sieht man sonst nicht, wie voll ein Tag ist. Zählt über
                die GLEICHE Logik wie die Haken-Zellen darunter (ACTIVE inkl.
                Warteliste; Warteliste separat als „+N W") und über die
                gefilterten Zeilen — mit aktiver Suche also die Teilsumme,
                sonst alle. Kapazität aus dem Sub-Event als „/max".
                Spalten-Vorlauf MUSS der Kopfzeile folgen (v28.53-Falle:
                Kopf- und Zeilen-Reihenfolge nebeneinanderlegen!). */}
            <tr style={{ borderBottom: '2px solid var(--dex-gray-200)', background: 'var(--dex-gray-50, #fafafa)' }}>
              <th
                colSpan={1 + (searchActive ? 1 : 0) + (personalColsCollapsed ? 1 : 6) + 1 + parentCustomFields.length + parentUserFields.length}
                style={{ textAlign: 'right', padding: '4px 8px', fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}
              >
                {isDe ? '∑ angemeldet:' : '∑ registered:'}
              </th>
              {childCustomFieldsByChild.map(({ child, fields }) => {
                let regCount = 0;
                let wlCount = 0;
                for (const row of consolidatedFiltered) {
                  const r = row.perChild[child.id];
                  if (!r || ACTIVE.indexOf(r.Status) < 0) continue;
                  if (r.Status === 'Warteliste') wlCount++; else regCount++;
                }
                const cap = (typeof child.maxParticipants === 'number' && child.maxParticipants > 0) ? child.maxParticipants : 0;
                return (
                  <React.Fragment key={`sum-${child.id}`}>
                    <th style={{ textAlign: 'center', padding: '4px 8px', borderLeft: '1px solid var(--dex-gray-200)', fontSize: '0.8rem', whiteSpace: 'nowrap', ...dimColStyle(child.id) }}
                        title={isDe
                          ? `${regCount} angemeldet${cap ? ` von ${cap} Plätzen` : ''}${wlCount ? ` · ${wlCount} auf der Warteliste` : ''}`
                          : `${regCount} registered${cap ? ` of ${cap} seats` : ''}${wlCount ? ` · ${wlCount} on the waitlist` : ''}`}>
                      <span style={{ fontWeight: 700, color: (cap > 0 && regCount >= cap) ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)' }}>{regCount}</span>
                      {cap > 0 && <span style={{ color: 'var(--dex-gray-400)', fontWeight: 400 }}>/{cap}</span>}
                      {wlCount > 0 && <span style={{ color: 'var(--dex-orange, #ed8b00)', fontWeight: 600 }}> +{wlCount} W</span>}
                    </th>
                    {fields.map(f => <th key={`sum-${child.id}-${f.id}`} style={{ padding: 0 }} />)}
                  </React.Fragment>
                );
              })}
              <th style={{ padding: 0 }} />
            </tr>
            {/* v30.17: „Anmeldung ab"-Zeile — nur bei aktiver Freischalt-Regel.
                Vergangene Tage heißen „vorbei", noch gesperrte zeigen ihr
                Öffnungsdatum (orange), offene das Datum in grün. */}
            {hasOpenRule && (
              <tr style={{ borderBottom: '2px solid var(--dex-gray-200)', background: 'var(--dex-gray-50, #fafafa)' }}>
                <th
                  colSpan={1 + (searchActive ? 1 : 0) + (personalColsCollapsed ? 1 : 6) + 1 + parentCustomFields.length + parentUserFields.length}
                  style={{ textAlign: 'right', padding: '2px 8px', fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}
                >
                  {isDe ? 'Anmeldung ab:' : 'Opens:'}
                </th>
                {childCustomFieldsByChild.map(({ child, fields }) => {
                  const st = childColState[child.id];
                  return (
                    <React.Fragment key={`open-${child.id}`}>
                      <th style={{
                        textAlign: 'center', padding: '2px 8px', borderLeft: '1px solid var(--dex-gray-200)',
                        fontSize: '0.72rem', whiteSpace: 'nowrap', fontWeight: 600,
                        color: st.past
                          ? 'var(--dex-gray-400)'
                          : st.notYetOpen ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green-dark, #4a7c1f)',
                      }}>
                        {st.past
                          ? (isDe ? 'vorbei' : 'past')
                          : st.openFrom
                          ? `${isDe ? 'ab ' : 'from '}${st.openFrom.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit' })}`
                          : '—'}
                      </th>
                      {fields.map(f => <th key={`open-${child.id}-${f.id}`} style={{ padding: 0 }} />)}
                    </React.Fragment>
                  );
                })}
                <th style={{ padding: 0 }} />
              </tr>
            )}
          </thead>
          <tbody>
            {consolidatedFiltered.map((row, idx) => {
              const isExpanded = expandedConsolidatedEmail === row.emailKey;
              return (
                <React.Fragment key={row.emailKey}>
                  <tr style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                    {/* v15.20: Im konsolidierten View einfach fortlaufend
                        durchnummerieren (idx+1). Die Sub-Event-TeilnehmerID
                        macht hier keinen Sinn, weil jede Person eine eigene
                        TID pro Sub-Event hat — sortbar bleibt es über
                        Vorname/Nachname/Email-Spalten. */}
                    <td style={{ padding: 8, color: 'var(--dex-gray-400)' }}>{idx + 1}</td>
                    {/* v26.68: echte Teilnehmer-ID (Klammer-Liste) — nur bei Suche. */}
                    {searchActive && (
                      <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontWeight: 600 }}>{bracketTidByEmail[row.emailKey] ?? row.teilnehmerId ?? '–'}</td>
                    )}
                    {personalColsCollapsed ? (
                      <td style={{ padding: 8 }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
                          {/* v24.56: Foto-Hover = Kontaktkarte (E-Mail + Teams). */}
                          {(() => {
                            const nm = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email || '-';
                            const sl = [row.jobTitle || '', stripLocPrefix(row.location || ''), row.company || ''].filter(Boolean).join(' • ');
                            return <PersonContactHover email={row.email || ''} name={nm} size={30} subline={sl} isDe={isDe} />;
                          })()}
                          {/* v23.32: zweizeilig — Name fett, darunter „Position • Standort" (ohne DE). */}
                          {(() => {
                            const fullName = `${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email || '-';
                            const loc = stripLocPrefix(row.location || '');
                            const sub = [row.jobTitle || '', loc, row.company || ''].filter(Boolean).join(' • ');
                            return (
                              <div
                                style={{ display: 'flex', flexDirection: 'column', minWidth: 0, lineHeight: 1.25, cursor: 'pointer' }}
                                title={isDe ? 'Detailinfos anzeigen' : 'Show details'}
                                onClick={() => setParticipantDetail({ name: fullName, email: row.email || '', jobTitle: row.jobTitle || '', location: row.location || '', company: row.company || '', department: '', phone: '', status: '', tid: row.teilnehmerId })}
                              >
                                <span style={{ fontWeight: 700, whiteSpace: 'nowrap' }}>{highlightMatch(fullName)}</span>
                                {sub && <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>{highlightMatch(sub)}</span>}
                              </div>
                            );
                          })()}
                        </div>
                      </td>
                    ) : (
                      <>
                        {/* v26.68: auch in der aufgeklappten Ansicht ein Foto mit
                            Hover-Kontaktkarte zeigen (vorher nur im eingeklappten
                            2-Spalten-Modus) — links neben dem Vornamen. */}
                        <td style={{ padding: 8, fontWeight: 500 }}>
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
                            <PersonContactHover
                              email={row.email || ''}
                              name={`${row.vorname || ''} ${row.nachname || ''}`.trim() || row.email || '-'}
                              size={28}
                              subline={[row.jobTitle || '', stripLocPrefix(row.location || ''), row.company || ''].filter(Boolean).join(' • ')}
                              isDe={isDe}
                            />
                            <span>{highlightMatch(row.vorname || '-')}</span>
                          </span>
                        </td>
                        <td style={{ padding: 8, fontWeight: 500 }}>{highlightMatch(row.nachname || '-')}</td>
                        <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{highlightMatch(row.email)}</td>
                        <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{highlightMatch(row.jobTitle || '-')}</td>
                        <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{row.location ? highlightMatch(stripLocPrefix(row.location)) : '-'}</td>
                        <td style={{ padding: 8, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{row.company ? highlightMatch(row.company) : '-'}</td>
                      </>
                    )}
                    {/* v26.84: „Registriert von" — Akteur aus den Sub-Event-
                        Registrierungen der Person ableiten (erste mit
                        RegisteredByEmail). Proxy = Name/Mail des Anmeldenden,
                        sonst „Selbst". */}
                    {(() => {
                      let actorEmail = '', actorName = '', isProxy = false;
                      for (const ck of Object.keys(row.perChild)) {
                        const cr = row.perChild[ck];
                        const ae = (cr && cr.RegisteredByEmail || '').trim();
                        if (cr && ae) {
                          actorEmail = ae; actorName = (cr.RegisteredByName || '').trim();
                          isProxy = ae.toLowerCase() !== (cr.ParticipantEmail || '').trim().toLowerCase();
                          break;
                        }
                      }
                      return (
                        <td style={{ padding: 8, fontSize: '0.8rem', whiteSpace: 'nowrap', color: isProxy ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-gray-500)' }} title={isProxy ? actorEmail : ''}>
                          {isProxy ? (actorName || actorEmail) : (isDe ? 'Selbst' : 'Self')}
                        </td>
                      );
                    })()}
                    {parentCustomFields.map(f => {
                      let val = '';
                      // v15.3.1: Parent-Level-Custom-Fields zuerst aus der
                      // PARENT-Teilnehmerliste auflösen (registrations =
                      // selectedEvent.id-Regs), erst danach Fallback auf
                      // Sub-Event-CustomData. Vorher wurde nur Sub-Event-
                      // CustomData gelesen — Parent-Felder waren immer leer.
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const parentReg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey) as any;
                      if (parentReg) {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const spName = (f as any).spInternalName || '';
                        let v: unknown = spName ? parentReg[spName] : undefined;
                        if ((v === undefined || v === null || v === '') && parentReg.CustomData) {
                          try { v = JSON.parse(parentReg.CustomData)[f.id]; } catch { /* */ }
                        }
                        if (v !== undefined && v !== null && v !== '') val = String(v);
                      }
                      // Fallback: Sub-Event-CustomData durchsuchen (Legacy-Events,
                      // bei denen Parent-Felder in Sub-Event-CustomData kopiert
                      // wurden — z.B. bei Wizard-„Vom Hauptevent kopieren").
                      if (!val) {
                        for (const ch of consolidatedChildren) {
                          const r = row.perChild[ch.id];
                          if (!r) continue;
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          const spName = (f as any).spInternalName || '';
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          let v: any = spName ? (r as any)[spName] : undefined;
                          if ((v === undefined || v === null || v === '') && r.CustomData) {
                            try { v = JSON.parse(r.CustomData)[f.id]; } catch { /* */ }
                          }
                          if (v !== undefined && v !== null && v !== '') { val = String(v); break; }
                        }
                      }
                      return (
                        <td key={`pcv-${f.id}`} style={{ padding: 8, fontSize: '0.8rem', color: 'var(--dex-gray-700)', whiteSpace: 'nowrap', maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', ...PASTEL_A_CELL }} title={val}>
                          {val ? highlightMatch(val) : '-'}
                        </td>
                      );
                    })}
                    {/* v23.32: People-Picker-Felder des Hauptevents — Foto + Name. */}
                    {parentUserFields.map(f => {
                      let raw = '';
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const parentReg = registrations.find(r => (r.ParticipantEmail || '').toLowerCase().trim() === row.emailKey) as any;
                      if (parentReg && parentReg.CustomData) {
                        try { const v = JSON.parse(parentReg.CustomData)[f.id]; if (v !== undefined && v !== null && v !== '') raw = String(v); } catch { /* */ }
                      }
                      // Fallback: Sub-Event-CustomData (kopierte Parent-Felder).
                      if (!raw) {
                        for (const ch of consolidatedChildren) {
                          const r = row.perChild[ch.id];
                          if (!r || !r.CustomData) continue;
                          try { const v = JSON.parse(r.CustomData)[f.id]; if (v !== undefined && v !== null && v !== '') { raw = String(v); break; } } catch { /* */ }
                        }
                      }
                      // Wert-Format „Anzeigename <email>", ggf. mehrere per „;".
                      const persons = raw.split(';').map(s => s.trim()).filter(Boolean).map(part => {
                        const m = part.match(/<([^>]+@[^>]+)>/);
                        const email = m ? m[1].trim() : '';
                        const name = (m ? part.slice(0, part.indexOf('<')) : part).trim() || email;
                        return { name, email };
                      });
                      return (
                        <td key={`puv-${f.id}`} style={{ padding: 8, fontSize: '0.8rem', color: 'var(--dex-gray-700)', maxWidth: 190, ...PASTEL_A_CELL }} title={raw}>
                          {persons.length === 0 ? '-' : (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              {persons.map((p, pi) => (
                                <span key={pi} style={{ display: 'inline-flex', alignItems: 'center', gap: 6, whiteSpace: 'nowrap' }}>
                                  {/* v26.68: Foto mit Kontaktkarte (Hover → E-Mail/Teams),
                                      wie bei allen anderen Personen-Fotos der Liste — vorher
                                      war es ein reines <img> ohne Mouse-over. */}
                                  {p.email ? (
                                    <PersonContactHover email={p.email} name={p.name} size={24} isDe={isDe} />
                                  ) : null}
                                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{highlightMatch(p.name)}</span>
                                </span>
                              ))}
                            </div>
                          )}
                        </td>
                      );
                    })}
                    {childCustomFieldsByChild.map(({ child, fields }) => {
                      const r = row.perChild[child.id];
                      const isReg = !!r && ACTIVE.indexOf(r.Status) >= 0;
                      return (
                        <React.Fragment key={`scv-${child.id}`}>
                          <td style={{ padding: 8, textAlign: 'center', borderLeft: '1px solid var(--dex-gray-200)', ...dimColStyle(child.id) }}
                              title={r ? `${translateStatus(r.Status, isDe)} — TID ${r.TeilnehmerID || '?'}` : (isDe ? 'Nicht angemeldet' : 'Not registered')}>
                            {isReg ? (
                              r.Status === 'Warteliste'
                                ? <span style={{ color: 'var(--dex-orange, #ed8b00)', fontSize: '0.78rem' }} title={translateStatus(r.Status, isDe)}>W</span>
                                : <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', display: 'inline-flex' }}><Check size={16} /></span>
                            ) : (
                              <span style={{ color: 'var(--dex-gray-300)' }}>—</span>
                            )}
                          </td>
                          {fields.map(f => {
                            let val = '';
                            if (r) {
                              // eslint-disable-next-line @typescript-eslint/no-explicit-any
                              const spName = (f as any).spInternalName || '';
                              // eslint-disable-next-line @typescript-eslint/no-explicit-any
                              let v: any = spName ? (r as any)[spName] : undefined;
                              if ((v === undefined || v === null || v === '') && r.CustomData) {
                                try { v = JSON.parse(r.CustomData)[f.id]; } catch { /* */ }
                              }
                              if (v !== undefined && v !== null && v !== '') val = String(v);
                            }
                            return (
                              <td key={`scv-${child.id}-${f.id}`} style={{ padding: 8, fontSize: '0.8rem', color: 'var(--dex-gray-700)', whiteSpace: 'nowrap', maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', ...PASTEL_B_CELL, ...dimColStyle(child.id) }} title={val}>
                                {val ? highlightMatch(val) : (r ? '-' : '')}
                              </td>
                            );
                          })}
                        </React.Fragment>
                      );
                    })}
                    <td style={{ padding: 8 }}>
                      {/* v26.68: Die früheren Einzel-Buttons (Details / Felder /
                          Zur Klammer / Assistenz zuordnen / Abmelden) in EIN
                          Auswahl-Dropdown zusammengefasst — spart Platz und hält
                          die Aktions-Spalte schmal. Ein natives <select> ist im
                          horizontal scrollbaren Tabellen-Container am robustesten
                          (kein Abschneiden durch overflow). */}
                      {(() => {
                        const showFelder = canManage && !orgPastLock && editableParentFieldCount > 0 && hasParentReg(row.emailKey);
                        const showKlammer = canManage && !orgPastLock && isConsolidatedMode && !hasParentReg(row.emailKey);
                        const showAssist = canManage && !orgPastLock;
                        const showAbmelden = canManage && !orgPastLock;
                        return (
                          <select
                            className="btn btn-secondary"
                            value=""
                            style={{ fontSize: '0.75rem', padding: '4px 10px', cursor: 'pointer', maxWidth: 170 }}
                            aria-label={isDe ? 'Aktionen für diese Person' : 'Actions for this person'}
                            onChange={(e) => {
                              const v = e.target.value;
                              if (v === 'details') setExpandedConsolidatedEmail(isExpanded ? null : row.emailKey);
                              else if (v === 'felder') openMainFieldsEdit(row.emailKey, `${row.vorname} ${row.nachname}`.trim() || row.email);
                              else if (v === 'klammer') { if (addingToKlammer !== row.emailKey) void addToKlammer(row); }
                              else if (v === 'assist') { setAssignAssistRow(row); setAssignAssistValue(''); }
                              else if (v === 'abmelden') openDeregModal(row);
                            }}
                          >
                            <option value="" disabled>{isDe ? 'Aktionen…' : 'Actions…'}</option>
                            <option value="details">{isExpanded ? (isDe ? 'Details schließen' : 'Close details') : (isDe ? 'Details anzeigen' : 'Show details')}</option>
                            {showFelder && <option value="felder">{isDe ? 'Hauptevent-Felder bearbeiten' : 'Edit main-event fields'}</option>}
                            {showKlammer && <option value="klammer">{addingToKlammer === row.emailKey ? '…' : (isDe ? 'Zur Klammer hinzufügen' : 'Add to umbrella')}</option>}
                            {showAssist && <option value="assist">{isDe ? 'Assistenz zuordnen' : 'Assign assistant'}</option>}
                            {showAbmelden && <option value="abmelden">{isDe ? 'Abmelden' : 'Cancel registration'}</option>}
                          </select>
                        );
                      })()}
                      {/* v29.30: Ist das Konto als inaktiv gemeldet (Person hat
                          womöglich das Unternehmen verlassen), steht das
                          Abmelden als eigener Knopf neben dem Klappmenü — der
                          Hinweis oben fordert genau dazu auf, und im
                          „Aktionen…"-Menü findet man es erst nach dem Öffnen. */}
                      {canManage && !orgPastLock && inactiveAccounts.indexOf(row.emailKey) >= 0 && (
                        <button
                          type="button"
                          className="btn btn-secondary"
                          onClick={() => openDeregModal(row)}
                          title={isDe
                            ? 'Kein aktives Konto — Abmelde-Dialog öffnen (still, inklusive Klammer und aller Sub-Events)'
                            : 'No active account — open the cancellation dialog (silent, including umbrella and all sub-events)'}
                          style={{ marginTop: 4, fontSize: '0.72rem', padding: '3px 8px', color: 'var(--dex-red, #c00)', borderColor: 'var(--dex-red, #c00)', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                        >
                          <Trash2 size={11} /> {isDe ? 'Abmelden' : 'Deregister'}
                        </button>
                      )}
                    </td>
                  </tr>
                  {isExpanded && (
                    <tr style={{ background: 'var(--dex-gray-50, #f7f7f7)' }}>
                      <td colSpan={totalColSpan} style={{ padding: '12px 16px' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                          <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
                            {isDe ? 'Anmeldungen von' : 'Registrations of'} {row.vorname} {row.nachname}
                          </div>
                          {consolidatedChildren.map(ch => {
                            const r = row.perChild[ch.id];
                            if (!r) {
                              return (
                                <div key={`exp-${ch.id}`} style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                                  {shortSubEventTitle(ch.title, selectedEvent?.title)} — {isDe ? 'nicht angemeldet' : 'not registered'}
                                </div>
                              );
                            }
                            return (
                              <div key={`exp-${ch.id}`} style={{ display: 'flex', alignItems: 'center', gap: 12, fontSize: '0.82rem' }}>
                                <span style={{ fontWeight: 500, minWidth: 200 }}>{shortSubEventTitle(ch.title, selectedEvent?.title)}</span>
                                <span className={`badge ${r.Status === 'Eingecheckt' ? 'badge-green' : 'badge-gray'}`}>{translateStatus(r.Status, isDe)}</span>
                                <span style={{ color: 'var(--dex-gray-500)' }}>TID {r.TeilnehmerID || '?'}</span>
                                <span style={{ color: 'var(--dex-gray-400)', fontSize: '0.75rem' }}>{formatDate(r.RegistrationDate)}</span>
                                <button
                                  className="btn btn-secondary"
                                  style={{ marginLeft: 'auto', fontSize: '0.72rem', padding: '3px 8px', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                                  onClick={() => setSelectedEvent(ch)}
                                >
                                  {isDe ? 'In Sub-Event öffnen' : 'Open in sub-event'} <ExternalLink size={12} />
                                </button>
                              </div>
                            );
                          })}
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    );
};

