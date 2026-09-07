/**
 * v29.26: „Teilnehmer hinzufügen" im Organizer Center.
 *
 * Teilnehmer registrieren sich normalerweise selbst — dieser Dialog ist der
 * Ausnahme-Weg für Organizer (nachträgliche Zusagen, VIP-Listen, Übernahme
 * aus anderen Quellen). Ablauf in einem Dialog, von oben nach unten:
 *
 *  1. Ziel wählen: Haupt-Event (sofern buchbar) und/oder Sub-Events
 *     (Mehrfachauswahl).
 *  2. Personen einsammeln — über denselben Massenimport-Matcher wie im
 *     Wizard (BulkUserImportModal: E-Mails direkt, Namen per Tenant-Suche,
 *     Mehrdeutige zum Auflösen).
 *  3. Falls ein gewähltes Ziel Formular-Felder hat: die Angaben pro Person
 *     ausfüllen (optional — der Organizer kennt nicht jede Antwort, deshalb
 *     seit v31.2 im Aufklapper). Ohne Felder entfällt der Schritt.
 *  4. Benachrichtigung: Bestätigungs-Mail ja/nein, Outlook-Termin ja/nein
 *     (registerForEvent kennt suppressMail/suppressOutlook).
 *
 * Die Anmeldungen laufen sequentiell (Drosselungs-Schonung) über
 * registerForEvent — derselbe Pfad wie die stellvertretende Anmeldung auf
 * der Anmeldeseite, inkl. Warteliste, Platz-Reservierung und AccessFix.
 */
import * as React from 'react';
import { DeloitteEvent, EventSpecificField } from '../../types';
import BulkUserImportModal, { BulkImportItem } from '../BulkUserImportModal';
import { formatDateTimeRange } from '../myEvents/myEventsHelpers';
import { X, Users, Check, Plus, ChevronDown } from '../Icons';
// v31.2: Gemeinsamer Modal-Rahmen (Kopf/Fuß/Portal) und die dex-ui-Klassen
// statt eines eigenen Overlays mit Inline-Styles — siehe docs/ui-leitfaden.md.
import Modal from '../Modal';
import { cx } from '../dexUi';
import { InfoTooltip } from '../InfoTooltip';

interface SearchHit { email: string; displayName: string; location?: string }

export interface AddParticipantsModalProps {
  open: boolean;
  onClose: () => void;
  /** Nach erfolgreichem Lauf: Teilnehmerliste neu laden. */
  onDone: () => void;
  /** Haupt-/Klammerevent der Familie (Top-Level). */
  mainEvent: DeloitteEvent;
  childEvents: DeloitteEvent[];
  /** Das im Organizer Center geöffnete Event — wird vorausgewählt. */
  preselectedId?: string;
  searchUsers: (q: string, includeIntl?: boolean) => Promise<SearchHit[]>;
  registerForEvent: (
    eventId: string,
    customData: Record<string, string>,
    participantFirstName?: string,
    participantLastName?: string,
    participantEmail?: string,
    preferredStarterType?: string,
    opts?: { suppressMail?: boolean; suppressOutlook?: boolean; skipReload?: boolean },
  ) => Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste'; reason?: string }>;
  isDe: boolean;
}

const splitName = (displayName: string, email: string): { first: string; last: string } => {
  const dn = (displayName || '').trim();
  if (!dn || dn === email) return { first: '', last: email };
  if (dn.indexOf(',') >= 0) {
    const p = dn.split(',').map(s => s.trim());
    return { last: p[0] || dn, first: p[1] || '' };
  }
  const w = dn.split(/\s+/);
  if (w.length >= 2) return { last: w[w.length - 1], first: w.slice(0, -1).join(' ') };
  return { last: dn, first: '' };
};

/** Formular-Felder, die sich hier sinnvoll abfragen lassen (kein Upload). */
const askableFields = (ev: DeloitteEvent): EventSpecificField[] =>
  (ev.eventSpecificFields || []).filter(f => f && f.label && f.label.trim() && f.type !== 'document');

// v31.2: Ablauf-Zeile „Nummer · Frage · Knopf" mit Platz darunter für Chips,
// Karten oder Schalter (Leitfaden: „Der Knopf ist der Schritt"). Erledigte
// Schritte zeigen einen Haken — so sieht man, was dem Primär-Knopf noch fehlt.
const StepRow = (p: { num: number; title: React.ReactNode; hint?: React.ReactNode; action?: React.ReactNode; done?: boolean; children?: React.ReactNode }): React.ReactElement => (
  <div className={cx('dex-ui-step', p.done && 'is-done')} style={{ alignItems: 'flex-start', flexWrap: 'wrap' }}>
    <span className="dex-ui-step-num" aria-hidden="true">{p.done ? <Check size={14} /> : p.num}</span>
    <div className="dex-ui-step-body">
      <div className="dex-ui-step-title">{p.title}</div>
      {p.hint && <div className="dex-ui-step-hint">{p.hint}</div>}
    </div>
    {p.action && <div className="dex-ui-step-action">{p.action}</div>}
    {p.children && <div style={{ flexBasis: '100%', paddingLeft: 40, boxSizing: 'border-box' }}>{p.children}</div>}
  </div>
);

export default function AddParticipantsModal(props: AddParticipantsModalProps): React.ReactElement | null {
  const { open, onClose, onDone, mainEvent, childEvents, preselectedId, searchUsers, registerForEvent, isDe } = props;

  const mainBookable = !mainEvent.subEventsOnlyMode;
  const [targetIds, setTargetIds] = React.useState<string[]>([]);
  const [people, setPeople] = React.useState<BulkImportItem[]>([]);
  const [bulkOpen, setBulkOpen] = React.useState(false);
  const [sendMail, setSendMail] = React.useState(true);
  const [sendOutlook, setSendOutlook] = React.useState(true);
  // fieldValues[eventId][personEmailLc][fieldId] = Wert
  const [fieldValues, setFieldValues] = React.useState<Record<string, Record<string, Record<string, string>>>>({});
  const [running, setRunning] = React.useState(false);
  const [progress, setProgress] = React.useState('');
  const [report, setReport] = React.useState<Array<{ person: string; target: string; status: string; ok: boolean }> | null>(null);
  // v31.2: Die optionalen Formular-Antworten liegen in einem Aufklapper — bei
  // fünf Personen × vier Feldern füllten sie sonst den ganzen Dialog.
  const [answersOpen, setAnswersOpen] = React.useState(false);

  React.useEffect(() => {
    if (!open) return;
    // Bei jedem Öffnen frisch starten; das geöffnete Event vorauswählen.
    const pre = (preselectedId && (preselectedId === mainEvent.id ? mainBookable : true)) ? [preselectedId] : [];
    setTargetIds(pre.length ? pre : (mainBookable && childEvents.length === 0 ? [mainEvent.id] : []));
    setPeople([]);
    setSendMail(true);
    setSendOutlook(true);
    setFieldValues({});
    setReport(null);
    setProgress('');
    setRunning(false);
    setAnswersOpen(false);
  }, [open, preselectedId, mainEvent.id, mainBookable, childEvents.length]);

  if (!open) return null;

  const targets: DeloitteEvent[] = [];
  if (mainBookable) targets.push(mainEvent);
  childEvents.forEach(ce => targets.push(ce));
  const selectedTargets = targets.filter(t => targetIds.indexOf(t.id) >= 0);
  const targetsWithFields = selectedTargets.filter(t => askableFields(t).length > 0);

  const toggleTarget = (id: string): void => {
    setTargetIds(prev => prev.indexOf(id) >= 0 ? prev.filter(x => x !== id) : prev.concat(id));
  };
  const setFieldValue = (evId: string, emailLc: string, fieldId: string, value: string): void => {
    setFieldValues(prev => {
      const evMap = { ...(prev[evId] || {}) };
      const pMap = { ...(evMap[emailLc] || {}) };
      pMap[fieldId] = value;
      evMap[emailLc] = pMap;
      return { ...prev, [evId]: evMap };
    });
  };

  const renderFieldInput = (evId: string, emailLc: string, f: EventSpecificField): React.ReactElement => {
    const val = ((fieldValues[evId] || {})[emailLc] || {})[f.id] || '';
    // v31.2: kompakte dex-ui-Eingaben statt form-input mit Inline-Maßen; Auswahl-
    // Optionen als Chips (mehrfach aktiv) — ein Klick, und der Hover zeigt, was klickbar ist.
    const cls = 'dex-ui-input dex-ui-input--sm';
    if (f.type === 'select' && f.multi) {
      // Mehrfachauswahl wird ' | '-getrennt gespeichert (Konvention v7.11).
      const chosen = val ? val.split(' | ').map(s => s.trim()).filter(Boolean) : [];
      return (
        <div className="dex-ui-inline" style={{ gap: 6 }}>
          {(f.options || []).map(opt => {
            const on = chosen.indexOf(opt) >= 0;
            return (
              <button key={opt} type="button" className={cx('dex-ui-chip', on && 'is-active')} aria-pressed={on}
                onClick={() => setFieldValue(evId, emailLc, f.id, (on ? chosen.filter(c => c !== opt) : chosen.concat(opt)).join(' | '))}>
                {on && <Check size={12} />}{opt}
              </button>
            );
          })}
        </div>
      );
    }
    if (f.type === 'select') {
      return (
        <select className="dex-ui-select" style={{ padding: '6px 32px 6px 10px', fontSize: '0.84rem' }} value={val} onChange={e => setFieldValue(evId, emailLc, f.id, e.target.value)}>
          <option value="">{isDe ? '— keine Angabe —' : '— no answer —'}</option>
          {(f.options || []).map(opt => <option key={opt} value={opt}>{opt}</option>)}
        </select>
      );
    }
    if (f.type === 'checkbox') {
      const on = val === 'Ja' || val === 'Yes' || val === 'true';
      return (
        <button type="button" className={cx('dex-ui-chip', on && 'is-active')} aria-pressed={on}
          onClick={() => setFieldValue(evId, emailLc, f.id, on ? '' : 'Ja')}>
          {on && <Check size={12} />}{isDe ? 'Ja' : 'Yes'}
        </button>
      );
    }
    if (f.type === 'date') {
      // Bestand aus v29.26 — kein neues natives Datumsfeld, nur neu eingekleidet.
      return (
        <input type={f.withTime ? 'datetime-local' : 'date'} className={cls} value={val}
          onChange={e => setFieldValue(evId, emailLc, f.id, e.target.value)} />
      );
    }
    if (f.type === 'number') {
      return <input type="number" className={cls} value={val} onChange={e => setFieldValue(evId, emailLc, f.id, e.target.value)} />;
    }
    // text, user, roommate, daterange: freie Eingabe (user/roommate = E-Mail;
    // daterange im Antwort-Format 'YYYY-MM-DD – YYYY-MM-DD').
    return (
      <input type="text" className={cls} value={val}
        placeholder={f.type === 'user' || f.type === 'roommate' ? 'email@deloitte.de' : (f.type === 'daterange' ? 'YYYY-MM-DD – YYYY-MM-DD' : '')}
        onChange={e => setFieldValue(evId, emailLc, f.id, e.target.value)} />
    );
  };

  // v30.67 (Review): `res.reason` ist ein Code ('dup-check-failed', 'full', …)
  // und stand roh in der Status-Spalte des Ergebnisberichts. Die Klartexte
  // entsprechen denen der Anmeldeseite (`regFailMessage` in
  // registration/submitFlow.ts — dort eine lokale, nicht exportierte Funktion,
  // deshalb hier lokal abgebildet und kurz gehalten für die Tabellenzelle).
  // Unbekannte Codes bleiben roh sichtbar — besser als ein pauschales
  // „Fehlgeschlagen", das die Ursache verschluckt.
  const reasonText = (reason?: string): string => {
    switch (reason) {
      case 'dup-check-failed':
        return isDe
          ? 'Nicht angelegt — Anmeldeliste gerade nicht lesbar (Drosselung), bitte später erneut hinzufügen'
          : 'Not created — registration list not readable right now (throttling), please add again later';
      case 'insert-failed':
        return isDe
          ? 'Nicht gespeichert — technischer Fehler an der Teilnehmerliste, bitte erneut versuchen (hält es an: „Spalten fixen“)'
          : 'Not saved — technical error on the participant list, please try again (if it persists: „Fix columns“)';
      case 'already-registered':
        return isDe ? 'Bereits angemeldet' : 'Already registered';
      case 'full':
        return isDe ? 'Alle Plätze belegt, Warteliste deaktiviert' : 'All seats taken, waitlist disabled';
      case 'deadline':
        return isDe ? 'Anmeldefrist abgelaufen' : 'Registration deadline has passed';
      case 'not-allowed':
        return isDe ? 'Nicht berechtigt, diese Person für dieses Event anzumelden' : 'Not allowed to register this person for this event';
      case 'event-not-found':
        return isDe ? 'Event nicht gefunden (keine Teilnehmerliste)' : 'Event not found (no participant list)';
      default:
        return reason || (isDe ? 'Fehlgeschlagen' : 'Failed');
    }
  };

  const run = async (): Promise<void> => {
    if (running || people.length === 0 || selectedTargets.length === 0) return;
    setRunning(true);
    setReport(null);
    const rows: Array<{ person: string; target: string; status: string; ok: boolean }> = [];
    const total = people.length * selectedTargets.length;
    let done = 0;
    for (const p of people) {
      const emailLc = (p.email || '').toLowerCase();
      const { first, last } = splitName(p.displayName, p.email);
      // v30.14: mind. eine erfolgreiche SUB-Event-Anmeldung? Dann braucht die
      // Person im Klammer-Modus auch die Schatten-Klammer-Zeile (s.u.).
      for (const t of selectedTargets) {
        done++;
        setProgress(`${done}/${total} — ${p.displayName || p.email} → ${t.title}`);
        const values = (fieldValues[t.id] || {})[emailLc] || {};
        const cleaned: Record<string, string> = {};
        Object.keys(values).forEach(k => { if ((values[k] || '').trim()) cleaned[k] = values[k]; });
        try {
          // v30.14: skipReload — vorher zog JEDE Person×Ziel-Kombination einen
          // kompletten loadEvents nach sich (Massen-Hinzufügen = 429-Welle).
          // Der eine Refresh kommt vom Aufrufer über onDone.
          const res = await registerForEvent(t.id, cleaned, first, last, p.email, undefined,
            { suppressMail: !sendMail, suppressOutlook: !sendOutlook, skipReload: true });
          rows.push({
            person: p.displayName || p.email,
            target: t.title,
            status: res.ok
              ? (res.status === 'Warteliste' ? (isDe ? 'Warteliste' : 'Waitlist') : (isDe ? 'Angemeldet' : 'Registered'))
              : reasonText(res.reason),
            ok: res.ok,
          });
        } catch (err) {
          rows.push({ person: p.displayName || p.email, target: t.title, status: String((err as Error)?.message || err), ok: false });
        }
      }
      // v30.14 → v30.68: Die Klammer-Zeile stellt registerForEvent selbst
      // sicher — VOR jedem Termin, mit zweitem Versuch danach und Nachzug-
      // Merker (utils/shadowHeal). Der eigene Nachzug hier ist entfallen.
    }
    setProgress('');
    setReport(rows);
    setRunning(false);
    onDone();
  };

  // v31.2: Pill-Farbe aus dem Status-Text, den `run` schreibt — kein zweites
  // Feld im Bericht, der Bericht selbst bleibt unverändert.
  const isWaitRow = (r: { ok: boolean; status: string }): boolean => r.ok && (r.status === 'Warteliste' || r.status === 'Waitlist');
  const okCount = report ? report.filter(r => r.ok && !isWaitRow(r)).length : 0;
  const waitCount = report ? report.filter(isWaitRow).length : 0;
  const failCount = report ? report.filter(r => !r.ok).length : 0;
  const allSelected = targets.length > 0 && selectedTargets.length === targets.length;
  const showAnswers = targetsWithFields.length > 0 && people.length > 0;
  const answerFieldCount = targetsWithFields.reduce((n, t) => n + askableFields(t).length, 0);
  const notifyNum = showAnswers ? 4 : 3;
  const notifyRows = [
    { on: sendMail, set: setSendMail, title: isDe ? 'Bestätigungs-Mail senden' : 'Send confirmation mail',
      desc: isDe ? 'Die Person bekommt die Bestätigungs-Mail des Events — wie nach einer eigenen Anmeldung.' : 'The person gets the event’s confirmation mail — as after registering themselves.' },
    { on: sendOutlook, set: setSendOutlook, title: isDe ? 'Outlook-Termin senden' : 'Send Outlook invitation',
      desc: isDe ? 'Die Kalendereinladung landet im Outlook der Person.' : 'The calendar invitation lands in the person’s Outlook.' },
  ];
  const runLabel = running ? (isDe ? 'Anmeldungen laufen…' : 'Registering…') : (isDe
    ? `${people.length || '–'} ${people.length === 1 ? 'Person' : 'Personen'} anmelden`
    : `Register ${people.length || '–'} ${people.length === 1 ? 'person' : 'people'}`);

  return (
    <Modal
      open={open}
      onClose={onClose}
      // v31.2: Läuft die Anmeldung oder ist der Personen-Dialog offen, schließt weder
      // Escape noch der Backdrop — sonst ginge die Auswahl hinter dem zweiten Dialog verloren.
      dismissable={!running && !bulkOpen}
      maxWidth={860}
      ariaLabel={isDe ? 'Teilnehmer hinzufügen' : 'Add attendees'}
      icon={<Users size={20} />}
      title={isDe ? 'Teilnehmer hinzufügen' : 'Add attendees'}
      subtitle={isDe
        ? 'Der Ausnahme-Weg für nachträgliche Zusagen oder übernommene Listen — normalerweise melden sich Teilnehmer selbst über die Anmeldeseite an.'
        : 'The exception path for late confirmations or imported lists — attendees normally register themselves via the registration page.'}
      footer={<>
        {running && <span className="dex-ui-modal-foot-left dex-ui-muted" aria-live="polite">{progress}</span>}
        <button type="button" className="btn btn-secondary" disabled={running} onClick={onClose}>
          {report ? (isDe ? 'Schließen' : 'Close') : (isDe ? 'Abbrechen' : 'Cancel')}
        </button>
        {!report && (
          <button type="button" className="btn btn-primary" disabled={running || people.length === 0 || selectedTargets.length === 0} onClick={() => { void run(); }}>
            <Users size={16} /> {runLabel}
          </button>
        )}
      </>}
    >
      <div className="dex-ui-stack">
        {/* 1 · Ziel — mehrere aus vielen → Chips, mehrfach aktiv */}
        <StepRow
          num={1}
          done={selectedTargets.length > 0}
          title={isDe ? 'Wofür anmelden?' : 'Register for what?'}
          hint={isDe ? 'Mehrfachauswahl möglich — jede Person wird für jedes gewählte Ziel angemeldet.' : 'Pick one or more — each person is registered for every selected target.'}
          action={targets.length > 2 ? (
            <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" disabled={running}
              onClick={() => setTargetIds(allSelected ? [] : targets.map(t => t.id))}>
              {allSelected ? (isDe ? 'Auswahl aufheben' : 'Clear selection') : (isDe ? 'Alle auswählen' : 'Select all')}
            </button>
          ) : undefined}
        >
          {targets.length === 0 ? (
            <div className="dex-ui-callout dex-ui-callout--warn">{isDe ? 'Kein buchbares Ziel vorhanden.' : 'No bookable target available.'}</div>
          ) : (
            <div className="dex-ui-inline" style={{ gap: 6 }}>
              {targets.map(t => {
                const on = targetIds.indexOf(t.id) >= 0;
                // v30.79: von–bis statt nur Datum.
                const meta = t.id === mainEvent.id ? (isDe ? 'Haupt-Event' : 'main event') : (t.startDate ? formatDateTimeRange(t.startDate, t.endDate, isDe) : '');
                return (
                  <button key={t.id} type="button" className={cx('dex-ui-chip', on && 'is-active')} aria-pressed={on} disabled={running} onClick={() => toggleTarget(t.id)}>
                    {on && <Check size={12} />}
                    {t.title || (isDe ? 'Sub-Event ohne Titel' : 'Untitled sub-event')}
                    {meta && <span style={{ fontWeight: 500, opacity: 0.8 }}>· {meta}</span>}
                  </button>
                );
              })}
            </div>
          )}
        </StepRow>

        {/* 2 · Personen — der Knopf sitzt in der Zeile, die Auswahl darunter */}
        <StepRow
          num={2}
          done={people.length > 0}
          title={<>{isDe ? 'Wen anmelden?' : 'Whom to register?'}
            {people.length > 0 && <span className="dex-ui-pill dex-ui-pill--green" style={{ marginLeft: 8 }}>{people.length} {people.length === 1 ? (isDe ? 'Person' : 'person') : (isDe ? 'Personen' : 'people')}</span>}</>}
          hint={isDe
            ? 'Namen oder E-Mails einfügen, z.B. aus Outlook oder Excel kopiert. E-Mails werden direkt übernommen, Namen im Tenant gesucht — wie beim Massenimport im Assistenten.'
            : 'Paste names or emails, e.g. copied from Outlook or Excel. Emails are taken directly, names are matched in the tenant — same as the bulk import in the wizard.'}
          action={
            <button type="button" className="btn btn-secondary dex-ui-btn-sm" disabled={running} onClick={() => setBulkOpen(true)}>
              <Plus size={14} /> {people.length === 0 ? (isDe ? 'Personen einfügen' : 'Add people') : (isDe ? 'Weitere einfügen' : 'Add more')}
            </button>
          }
        >
          {people.length > 0 ? (
            <div className="dex-ui-inline" style={{ gap: 6 }}>
              {people.map(p => (
                <span key={p.email} className="dex-ui-pill dex-ui-pill--green" style={{ paddingRight: 4 }}>
                  {p.displayName || p.email}
                  <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" style={{ width: 20, height: 20 }} disabled={running}
                    onClick={() => setPeople(prev => prev.filter(x => x.email !== p.email))} aria-label={isDe ? 'Entfernen' : 'Remove'} title={isDe ? 'Entfernen' : 'Remove'}>
                    <X size={12} />
                  </button>
                </span>
              ))}
            </div>
          ) : (
            <div className="dex-ui-muted">{isDe ? 'Noch niemand ausgewählt.' : 'Nobody selected yet.'}</div>
          )}
        </StepRow>

        {/* 3 · Formular-Antworten (nur wenn ein Ziel Felder hat) — optional, deshalb im Aufklapper */}
        {showAnswers && (
          <StepRow
            num={3}
            title={<>{isDe ? 'Antworten zum Anmeldeformular' : 'Registration-form answers'}
              <span className="dex-ui-label-optional" style={{ marginLeft: 6 }}>(optional)</span>
              <InfoTooltip text={isDe ? 'Diese Felder fragt das Anmeldeformular normalerweise ab. Datei-Upload-Felder lassen sich hier nicht befüllen.' : 'The registration form normally asks these questions. File-upload fields cannot be filled here.'} /></>}
            hint={isDe
              ? 'Fülle aus, was du weißt — Leeres wird ohne Antwort gespeichert und lässt sich später in der Teilnehmerliste über „Bearbeiten“ nachtragen.'
              : 'Fill in what you know — empty fields are saved without an answer and can be completed later via “Edit” in the attendee list.'}
            action={
              <button type="button" className={cx('dex-ui-disclosure', answersOpen && 'is-open')} style={{ width: 'auto', margin: 0 }} aria-expanded={answersOpen} onClick={() => setAnswersOpen(o => !o)}>
                <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                {answersOpen ? (isDe ? 'Ausblenden' : 'Hide') : (isDe ? `${answerFieldCount} ${answerFieldCount === 1 ? 'Feld' : 'Felder'} eintragen` : `Fill in ${answerFieldCount} ${answerFieldCount === 1 ? 'field' : 'fields'}`)}
              </button>
            }
          >
            {answersOpen && (
              <div className="dex-ui-stack dex-ui-fade-in">
                {targetsWithFields.map(t => (
                  <div key={t.id} className="dex-ui-stack" style={{ gap: 8 }}>
                    {targetsWithFields.length > 1 || selectedTargets.length > 1 ? (
                      <div className="dex-ui-section-title" style={{ margin: 0 }}>{t.title}</div>
                    ) : null}
                    {people.map(p => {
                      const emailLc = (p.email || '').toLowerCase();
                      return (
                        <div key={p.email} className="dex-ui-card dex-ui-card--soft" style={{ padding: '10px 14px' }}>
                          <div className="dex-ui-row-title" style={{ marginBottom: 8 }}>{p.displayName || p.email}</div>
                          <div className="dex-ui-grid-auto" style={{ gap: 10 }}>
                            {askableFields(t).map(f => (
                              <div key={f.id}>
                                <div className="dex-ui-muted" style={{ fontWeight: 600, fontSize: '0.74rem', marginBottom: 3 }}>{f.label}{f.required ? ' *' : ''}</div>
                                {renderFieldInput(t.id, emailLc, f)}
                              </div>
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>
            )}
          </StepRow>
        )}

        {/* 3/4 · Benachrichtigung — Ja/Nein mit Folge → Schalter-Zeilen */}
        <StepRow
          num={notifyNum}
          done
          title={isDe ? 'Was bekommt die Person mit?' : 'What does the person receive?'}
          hint={isDe
            ? 'Angemeldet wird in jedem Fall regulär — mit Platz- und Wartelisten-Logik. Du entscheidest nur über die Benachrichtigung.'
            : 'Registration always runs the regular way — with seat and waitlist logic. You only decide about the notification.'}
        >
          <div className="dex-ui-grid-2" style={{ gap: 10 }}>
            {notifyRows.map(n => (
              <label key={n.title} className={cx('dex-ui-toggle-row', n.on && 'is-active', running && 'is-disabled')}>
                <input type="checkbox" checked={n.on} onChange={e => n.set(e.target.checked)} disabled={running} />
                <span className="dex-ui-toggle-row-body">
                  <span className="dex-ui-toggle-row-title">{n.title}</span>
                  <span className="dex-ui-toggle-row-desc">{n.desc}</span>
                </span>
              </label>
            ))}
          </div>
        </StepRow>

        {/* Ergebnis — Zähler als Pills, Zeilen als Tabelle mit Status-Pill */}
        {report && (
          <div className="dex-ui-section dex-ui-fade-in">
            <div className="dex-ui-section-title">{isDe ? 'Ergebnis' : 'Result'}</div>
            <div className="dex-ui-inline" style={{ marginBottom: 10 }}>
              <span className="dex-ui-pill dex-ui-pill--green">{okCount} {isDe ? 'angemeldet' : 'registered'}</span>
              {waitCount > 0 && <span className="dex-ui-pill dex-ui-pill--orange">{waitCount} {isDe ? 'auf der Warteliste' : 'on the waitlist'}</span>}
              {failCount > 0 && <span className="dex-ui-pill dex-ui-pill--red">{failCount} {isDe ? 'nicht angemeldet' : 'not registered'}</span>}
            </div>
            <div className="dex-ui-table-wrap">
              <table className="dex-ui-table">
                <thead><tr><th>Person</th><th>{isDe ? 'Ziel' : 'Target'}</th><th>Status</th></tr></thead>
                <tbody>
                  {report.map((r, i) => (
                    <tr key={i}>
                      <td style={{ fontWeight: 600 }}>{r.person}</td><td>{r.target}</td>
                      <td><span className={cx('dex-ui-pill', r.ok ? (isWaitRow(r) ? 'dex-ui-pill--orange' : 'dex-ui-pill--green') : 'dex-ui-pill--red')} style={{ whiteSpace: 'normal' }}>
                        {r.ok ? <Check size={12} /> : <X size={12} />}{r.status}
                      </span></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      <BulkUserImportModal
        open={bulkOpen}
        onClose={() => setBulkOpen(false)}
        title={isDe ? 'Personen einfügen' : 'Add people'}
        description={isDe
          ? 'Namen oder E-Mail-Adressen einfügen (z.B. aus Outlook oder Excel kopiert). Die Personen landen als Auswahl im Dialog — angemeldet wird erst über den Knopf unten.'
          : 'Paste names or email addresses (e.g. copied from Outlook or Excel). People are collected in the dialog — registration only happens via the button below.'}
        existingEmails={people.map(p => p.email)}
        searchUsers={searchUsers}
        onAdd={item => setPeople(prev => prev.some(x => x.email.toLowerCase() === item.email.toLowerCase()) ? prev : prev.concat(item))}
      />
    </Modal>
  );
}
