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
 *  3. Optionen: Bestätigungs-Mail ja/nein, Outlook-Termin ja/nein
 *     (registerForEvent kennt suppressMail/suppressOutlook).
 *  4. Falls ein gewähltes Ziel Formular-Felder hat: die Angaben pro Person
 *     ausfüllen (optional — der Organizer kennt nicht jede Antwort). Ohne
 *     Felder wird direkt angemeldet.
 *
 * Die Anmeldungen laufen sequentiell (Drosselungs-Schonung) über
 * registerForEvent — derselbe Pfad wie die stellvertretende Anmeldung auf
 * der Anmeldeseite, inkl. Warteliste, Platz-Reservierung und AccessFix.
 */
import * as React from 'react';
import { DeloitteEvent, EventSpecificField } from '../../types';
import BulkUserImportModal, { BulkImportItem } from '../BulkUserImportModal';
import { X, Users } from '../Icons';

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
    const small: React.CSSProperties = { fontSize: '0.8rem', padding: '5px 8px' };
    if (f.type === 'select' && f.multi) {
      // Mehrfachauswahl wird ' | '-getrennt gespeichert (Konvention v7.11).
      const chosen = val ? val.split(' | ').map(s => s.trim()).filter(Boolean) : [];
      return (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
          {(f.options || []).map(opt => (
            <label key={opt} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: '0.78rem', cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={chosen.indexOf(opt) >= 0}
                onChange={e => {
                  const next = e.target.checked ? chosen.concat(opt) : chosen.filter(c => c !== opt);
                  setFieldValue(evId, emailLc, f.id, next.join(' | '));
                }}
              />
              {opt}
            </label>
          ))}
        </div>
      );
    }
    if (f.type === 'select') {
      return (
        <select className="form-input" style={small} value={val} onChange={e => setFieldValue(evId, emailLc, f.id, e.target.value)}>
          <option value="">{isDe ? '— keine Angabe —' : '— no answer —'}</option>
          {(f.options || []).map(opt => <option key={opt} value={opt}>{opt}</option>)}
        </select>
      );
    }
    if (f.type === 'checkbox') {
      return (
        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.8rem', cursor: 'pointer' }}>
          <input type="checkbox" checked={val === 'Ja' || val === 'Yes' || val === 'true'} onChange={e => setFieldValue(evId, emailLc, f.id, e.target.checked ? 'Ja' : '')} />
          {isDe ? 'Ja' : 'Yes'}
        </label>
      );
    }
    if (f.type === 'date') {
      return (
        <input type={f.withTime ? 'datetime-local' : 'date'} className="form-input" style={small} value={val}
          onChange={e => setFieldValue(evId, emailLc, f.id, e.target.value)} />
      );
    }
    if (f.type === 'number') {
      return <input type="number" className="form-input" style={small} value={val} onChange={e => setFieldValue(evId, emailLc, f.id, e.target.value)} />;
    }
    // text, user, roommate, daterange: freie Eingabe (user/roommate = E-Mail;
    // daterange im Antwort-Format 'YYYY-MM-DD – YYYY-MM-DD').
    return (
      <input type="text" className="form-input" style={small} value={val}
        placeholder={f.type === 'user' || f.type === 'roommate' ? 'email@deloitte.de' : (f.type === 'daterange' ? 'YYYY-MM-DD – YYYY-MM-DD' : '')}
        onChange={e => setFieldValue(evId, emailLc, f.id, e.target.value)} />
    );
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
      let anySubOk = false;
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
          if (res.ok && t.id !== mainEvent.id) anySubOk = true;
          rows.push({
            person: p.displayName || p.email,
            target: t.title,
            status: res.ok
              ? (res.status === 'Warteliste' ? (isDe ? 'Warteliste' : 'Waitlist') : (isDe ? 'Angemeldet' : 'Registered'))
              : (res.reason || (isDe ? 'Fehlgeschlagen' : 'Failed')),
            ok: res.ok,
          });
        } catch (err) {
          rows.push({ person: p.displayName || p.email, target: t.title, status: String((err as Error)?.message || err), ok: false });
        }
      }
      // v30.14: Klammer-Schatten-Zeile nachziehen. Dieser Pfad meldete Personen
      // NUR in die Sub-Events — die Klammer-Zeile fehlte, und im Organizer
      // Center lief die Box „Fehlende Klammer-Anmeldung" voll (Befund mit 24
      // Personen im Soft Opening). registerForEvent ist für die Klammer
      // idempotent: existiert bereits eine aktive Schatten-Zeile, wird nichts
      // eingefügt. Still (keine Mail, kein Outlook) — reine Datenvollständigkeit.
      if (anySubOk && mainEvent.subEventsOnlyMode) {
        try {
          const shadow = await registerForEvent(mainEvent.id, {}, first, last, p.email, undefined,
            { suppressMail: true, suppressOutlook: true, skipReload: true });
          if (!shadow.ok) {
            rows.push({ person: p.displayName || p.email, target: isDe ? 'Klammer-Event (Schattenzeile)' : 'Umbrella event (shadow row)', status: shadow.reason || (isDe ? 'Fehlgeschlagen' : 'Failed'), ok: false });
          }
        } catch (err) {
          rows.push({ person: p.displayName || p.email, target: isDe ? 'Klammer-Event (Schattenzeile)' : 'Umbrella event (shadow row)', status: String((err as Error)?.message || err), ok: false });
        }
      }
    }
    setProgress('');
    setReport(rows);
    setRunning(false);
    onDone();
  };

  const box: React.CSSProperties = { border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: '12px 14px', marginBottom: 14 };
  const secTitle: React.CSSProperties = { fontWeight: 700, fontSize: '0.85rem', marginBottom: 8, color: 'var(--dex-gray-800, #333)' };

  return (
    <div
      style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 1050, display: 'flex', alignItems: 'center', justifyContent: 'center' }}
      onClick={() => { if (!running) onClose(); }}
    >
      <div className="card" style={{ width: '92%', maxWidth: 860, maxHeight: '88vh', overflow: 'auto', padding: 24 }} onClick={e => e.stopPropagation()}>
        <div className="flex-between mb-16">
          <h3 style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
            <Users size={18} /> {isDe ? 'Teilnehmer hinzufügen' : 'Add attendees'}
          </h3>
          <button type="button" onClick={() => { if (!running) onClose(); }} style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4 }} aria-label="Schließen">
            <X size={18} />
          </button>
        </div>
        <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginTop: 0, marginBottom: 14, lineHeight: 1.5 }}>
          {isDe
            ? 'Teilnehmer registrieren sich normalerweise selbst über die Anmeldeseite — dieser Weg ist für Ausnahmen gedacht (nachträgliche Zusagen, übernommene Listen). Die Personen werden regulär angemeldet: mit Platz-/Wartelisten-Logik und, je nach Auswahl unten, mit Bestätigungs-Mail und Outlook-Termin.'
            : 'Attendees normally register themselves via the registration page — this path is for exceptions (late confirmations, imported lists). People are registered through the regular flow: seat/waitlist logic and, depending on the options below, confirmation mail and Outlook invite.'}
        </p>

        {/* 1 · Ziel */}
        <div style={box}>
          <div style={secTitle}>{isDe ? '1 · Wofür anmelden?' : '1 · Register for what?'}</div>
          {targets.length === 0 && (
            <div style={{ fontSize: '0.8rem', color: 'var(--dex-red, #c00)' }}>
              {isDe ? 'Kein buchbares Ziel vorhanden.' : 'No bookable target available.'}
            </div>
          )}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            {mainBookable && (
              <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.84rem', cursor: 'pointer' }}>
                <input type="checkbox" checked={targetIds.indexOf(mainEvent.id) >= 0} onChange={() => toggleTarget(mainEvent.id)} disabled={running} />
                <strong>{mainEvent.title}</strong>
                <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem' }}>{isDe ? '(Haupt-Event)' : '(main event)'}</span>
              </label>
            )}
            {childEvents.map(ce => (
              <label key={ce.id} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.84rem', cursor: 'pointer', marginLeft: mainBookable ? 18 : 0 }}>
                <input type="checkbox" checked={targetIds.indexOf(ce.id) >= 0} onChange={() => toggleTarget(ce.id)} disabled={running} />
                {ce.title || (isDe ? 'Sub-Event ohne Titel' : 'Untitled sub-event')}
                {ce.startDate && (
                  <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem' }}>
                    {new Date(ce.startDate).toLocaleDateString(isDe ? 'de-DE' : 'en-GB')}
                  </span>
                )}
              </label>
            ))}
          </div>
        </div>

        {/* 2 · Personen */}
        <div style={box}>
          <div style={secTitle}>{isDe ? '2 · Wen anmelden?' : '2 · Whom to register?'}</div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: people.length > 0 ? 10 : 0 }}>
            {people.map(p => (
              <span key={p.email} style={{
                display: 'inline-flex', alignItems: 'center', gap: 6, padding: '3px 10px', borderRadius: 999,
                background: 'rgba(134,188,37,0.12)', border: '1px solid var(--dex-green, #86bc25)', fontSize: '0.78rem',
              }}>
                {p.displayName || p.email}
                <button type="button" onClick={() => setPeople(prev => prev.filter(x => x.email !== p.email))} disabled={running}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, display: 'inline-flex' }} aria-label={isDe ? 'Entfernen' : 'Remove'}>
                  <X size={12} />
                </button>
              </span>
            ))}
          </div>
          <button type="button" className="btn btn-secondary" style={{ fontSize: '0.8rem', padding: '6px 12px' }} disabled={running} onClick={() => setBulkOpen(true)}>
            {people.length === 0
              ? (isDe ? 'Personen einfügen (Namen oder E-Mails)' : 'Add people (names or emails)')
              : (isDe ? 'Weitere Personen einfügen' : 'Add more people')}
          </button>
          <span style={{ marginLeft: 10, fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
            {isDe
              ? 'E-Mails werden direkt übernommen, Namen im Tenant gesucht — wie beim Massenimport im Assistenten.'
              : 'Emails are taken directly, names are matched in the tenant — same as the bulk import in the wizard.'}
          </span>
        </div>

        {/* 3 · Optionen */}
        <div style={box}>
          <div style={secTitle}>{isDe ? '3 · Benachrichtigung' : '3 · Notifications'}</div>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.84rem', cursor: 'pointer', marginBottom: 6 }}>
            <input type="checkbox" checked={sendMail} onChange={e => setSendMail(e.target.checked)} disabled={running} />
            {isDe ? 'Bestätigungs-Mail an die Person senden' : 'Send the confirmation mail to the person'}
          </label>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.84rem', cursor: 'pointer' }}>
            <input type="checkbox" checked={sendOutlook} onChange={e => setSendOutlook(e.target.checked)} disabled={running} />
            {isDe ? 'Outlook-Kalendereinladung erzeugen' : 'Create the Outlook calendar invitation'}
          </label>
        </div>

        {/* 4 · Felder (nur wenn ein Ziel welche hat) */}
        {targetsWithFields.length > 0 && people.length > 0 && (
          <div style={box}>
            <div style={secTitle}>{isDe ? '4 · Angaben zu den Formular-Feldern (optional)' : '4 · Registration-form answers (optional)'}</div>
            <p style={{ fontSize: '0.76rem', color: 'var(--dex-gray-600)', marginTop: 0, marginBottom: 10, lineHeight: 1.5 }}>
              {isDe
                ? 'Diese Felder fragt das Anmeldeformular normalerweise ab. Fülle aus, was du weißt — leere Felder werden ohne Antwort gespeichert und können später über „Bearbeiten" in der Teilnehmerliste nachgetragen werden. Datei-Upload-Felder können hier nicht befüllt werden.'
                : 'The registration form normally asks these questions. Fill in what you know — empty fields are saved without an answer and can be completed later via “Edit” in the attendee list. File-upload fields cannot be filled here.'}
            </p>
            {targetsWithFields.map(t => (
              <div key={t.id} style={{ marginBottom: 12 }}>
                {targetsWithFields.length > 1 || selectedTargets.length > 1 ? (
                  <div style={{ fontWeight: 700, fontSize: '0.8rem', color: 'var(--dex-green-dark, #4a7c1f)', marginBottom: 6 }}>{t.title}</div>
                ) : null}
                {people.map(p => {
                  const emailLc = (p.email || '').toLowerCase();
                  return (
                    <div key={p.email} style={{ border: '1px solid var(--dex-gray-100)', borderRadius: 8, padding: '8px 10px', marginBottom: 6, background: 'var(--dex-gray-50, #fafafa)' }}>
                      <div style={{ fontWeight: 600, fontSize: '0.8rem', marginBottom: 6 }}>{p.displayName || p.email}</div>
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))', gap: 8 }}>
                        {askableFields(t).map(f => (
                          <div key={f.id}>
                            <div style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 2 }}>
                              {f.label}{f.required ? ' *' : ''}
                            </div>
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

        {/* Ergebnis */}
        {report && (
          <div style={{ ...box, background: 'var(--dex-gray-50, #fafafa)' }}>
            <div style={secTitle}>{isDe ? 'Ergebnis' : 'Result'}</div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 3, fontSize: '0.8rem' }}>
              {report.map((r, i) => (
                <div key={i} style={{ color: r.ok ? 'var(--dex-gray-700)' : 'var(--dex-red, #c00)' }}>
                  {r.ok ? '✓' : '✗'} <strong>{r.person}</strong> → {r.target}: {r.status}
                </div>
              ))}
            </div>
          </div>
        )}

        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', alignItems: 'center', flexWrap: 'wrap' }}>
          {running && <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginRight: 'auto' }}>{progress}</span>}
          <button type="button" className="btn btn-secondary" disabled={running} onClick={onClose}>
            {report ? (isDe ? 'Schließen' : 'Close') : (isDe ? 'Abbrechen' : 'Cancel')}
          </button>
          {!report && (
            <button
              type="button"
              className="btn btn-primary"
              disabled={running || people.length === 0 || selectedTargets.length === 0}
              onClick={() => { void run(); }}
            >
              {running
                ? (isDe ? 'Anmeldungen laufen…' : 'Registering…')
                : (isDe
                  ? `${people.length || '–'} ${people.length === 1 ? 'Person' : 'Personen'} anmelden`
                  : `Register ${people.length || '–'} ${people.length === 1 ? 'person' : 'people'}`)}
            </button>
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
      </div>
    </div>
  );
}
