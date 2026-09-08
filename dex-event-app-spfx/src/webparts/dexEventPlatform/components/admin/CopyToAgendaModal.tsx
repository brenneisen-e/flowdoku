/**
 * CopyToAgendaModal — v30.93 (Programmpunkte, Stufe 4;
 * docs/konzept-programmpunkte.md, Abschnitt 5).
 *
 * „Als neues Event mit Programmpunkten kopieren": Ein bestehendes Event mit
 * Sub-Events wird in ein NEUES Event überführt, dessen Programmpunkte die
 * bisherigen Sub-Events sind. Nutzer-Entscheidung 07.09.2026: „Migration
 * würde ich gerne als neues Event machen, damit das alte nicht verloren geht."
 * Deshalb wird am alten Event NICHTS gelöscht, nichts ausgeladen, nichts
 * umgehängt — es bleibt vollständig stehen. Einzige optionale Änderung: seine
 * Anmeldefristen auf „jetzt", damit sich niemand doppelt anmeldet.
 *
 * Ablauf (Reihenfolge nach CLAUDE.md: prüfbar zuerst, der unumkehrbare
 * Schritt zuletzt — hier gibt es keinen unumkehrbaren):
 *  1. Prüfen: alle Termin-Listen MIT onHttpError lesen. Eine nicht lesbare
 *     Liste bricht ab — sie zählt nicht als leer.
 *  2. Neues Event anlegen (Test-Event/Entwurf, Agenda aus den Sub-Events,
 *     _agendaCheckIn, Kommunikation der Klammer).
 *  3. Anmeldungen still kopieren (registerForEvent mit suppressMail/
 *     suppressOutlook — die Person ist längst eingeladen), sequentiell mit
 *     Fehlerzähler.
 *  4. Check-ins der Sub-Events als Anwesenheit am jeweiligen Punkt.
 *  5. Optional: Anmeldung am alten Event schließen.
 * Warteliste wird NICHT kopiert (kein Platz), sondern namentlich gemeldet.
 */
import * as React from 'react';
import Modal from '../Modal';
import { useDialog } from '../../context/DialogContext';
import { useEvents } from '../../context/EventContext';
import { DeloitteEvent, AgendaItem } from '../../types';
import { EventService, SPRegistration, CustomField } from '../../services/EventService';
import { shortSubEventTitle } from '../../utils/subEventTitle';
import { subEventGroupKey, stripGroupPrefix } from '../../utils/subEventGroups';
import { suggestClusterName, agendaGroups, groupLabel, groupDateLabel } from '../../utils/agendaGroups';
import { useCurrentUser } from '../../context/UserContext';
// v31.2: Gemeinsame Klassen (Kennzahlen, Zeilen, Schalter-Zeilen, Aufklapper) statt Inline-Styles — Hover inklusive.
import { cx } from '../dexUi';
import { Copy, Check, Calendar, AlertCircle, ChevronDown } from '../Icons';

const ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

interface Person {
  email: string;
  first: string;
  last: string;
  name: string;
  /** Sub-Event-Ids, an denen die Person eingecheckt war. */
  checkedIn: string[];
  /** Antworten aus CustomData, zusammengeführt (Klammer zuerst, dann Termine). */
  customData: Record<string, string>;
}

interface Analysis {
  items: Array<AgendaItem & { childId: string }>;
  persons: Person[];
  waitlisted: string[];
  unreadable: string[];
  checkInCount: number;
}

const pad = (n: number): string => String(n).padStart(2, '0');
const localDate = (iso: string): string => { const d = new Date(iso); return isNaN(d.getTime()) ? '' : `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const localTime = (iso: string): string => { const d = new Date(iso); return isNaN(d.getTime()) ? '' : `${pad(d.getHours())}:${pad(d.getMinutes())}`; };

export default function CopyToAgendaModal(props: {
  event: DeloitteEvent;
  childEvents: DeloitteEvent[];
  isDe: boolean;
  onClose: () => void;
  onDone: () => void;
}): React.ReactElement {
  const { event, childEvents, isDe, onClose, onDone } = props;
  const { getAllRegistrations, createEvent, registerForEvent, refreshEvents, updateEvent } = useEvents();
  const { showAlert } = useDialog();
  // v30.94 (Nutzer-Ansage 07.09.2026): Organizer des neuen Events ist NUR die
  // Person, die überführt — nicht das ganze Team des alten Events. Sonst
  // bekommen alle bei jedem Test-Lauf die Organizer-Mails.
  const { currentUser } = useCurrentUser();
  // registerForEvent liest die Events aus dem Context-State; nach dem Anlegen
  // und refreshEvents muss die FRISCHE Closure genutzt werden, nicht die vom
  // Start der Aktion (React 17 rendert nach setState außerhalb von Handlern
  // synchron — die Ref trägt dann schon die neue Funktion).
  const registerRef = React.useRef(registerForEvent);
  registerRef.current = registerForEvent;

  const [phase, setPhase] = React.useState<'analyzing' | 'ready' | 'running' | 'done' | 'blocked'>('analyzing');
  const [analysis, setAnalysis] = React.useState<Analysis | null>(null);
  const [newTitle, setNewTitle] = React.useState<string>(`${event.title} (${isDe ? 'Programmpunkte' : 'Agenda items'})`);
  const [asTest, setAsTest] = React.useState<boolean>(true);
  const [closeOld, setCloseOld] = React.useState<boolean>(false);
  const [progress, setProgress] = React.useState<{ done: number; total: number; label: string }>({ done: 0, total: 0, label: '' });
  const [report, setReport] = React.useState<{ newId: string; registered: number; failed: Array<{ name: string; reason: string }>; checkIns: number; checkInFailed: number; oldClosed: boolean } | null>(null);
  // v31.2: „Was nicht mitkopiert wird" ist Erklärtext für den seltenen Fall — zu, bis jemand ihn braucht.
  const [showSkipped, setShowSkipped] = React.useState<boolean>(false);

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      const kids = childEvents.slice().sort((a, b) => (a.startDate || '').localeCompare(b.startDate || ''));
      // v30.94: Cluster mitnehmen. Trägt der Termin-Titel ein Präfix („Day 1 -
      // Welcome"), wird das Präfix der Cluster und der Titel verliert es; sonst
      // heißt der Cluster nach dem Tag („Tag 1", „Tag 2" — je Datum eins).
      const days = kids.map(k => localDate(k.startDate)).filter((d, i, arr) => d && arr.indexOf(d) === i).sort();
      const items: Analysis['items'] = kids.map((k, i) => {
        const shortTitle = shortSubEventTitle(k.title, event.title) || k.title;
        const prefix = subEventGroupKey(shortTitle);
        const date = localDate(k.startDate);
        return {
          childId: k.id,
          id: `ag-${Date.now().toString(36)}-${i}-${Math.random().toString(36).slice(2, 6)}`,
          date,
          time: localTime(k.startDate),
          endTime: k.endDate ? localTime(k.endDate) : '',
          icon: 'Calendar',
          title: prefix ? (stripGroupPrefix(shortTitle, prefix) || shortTitle) : shortTitle,
          description: '',
          location: k.location || '',
          cluster: prefix || (date ? suggestClusterName(days.indexOf(date), isDe) : undefined),
        };
      });
      const byEmail = new Map<string, Person>();
      const waitlisted = new Set<string>();
      const unreadable: string[] = [];
      let checkInCount = 0;
      const absorb = (regs: SPRegistration[], childId: string | null): void => {
        for (const r of regs) {
          const email = (r.ParticipantEmail || '').toLowerCase().trim();
          if (!email) continue;
          const nm = `${r.Vorname || ''} ${r.Nachname || ''}`.trim() || r.ParticipantName || email;
          if (r.Status === 'Warteliste') { waitlisted.add(nm); continue; }
          if (ACTIVE.indexOf(r.Status || '') < 0) continue;
          let p = byEmail.get(email);
          if (!p) {
            p = { email, first: r.Vorname || (r.ParticipantName || '').split(' ')[0] || '', last: r.Nachname || (r.ParticipantName || '').split(' ').slice(1).join(' '), name: nm, checkedIn: [], customData: {} };
            byEmail.set(email, p);
          }
          try {
            const cd = JSON.parse(r.CustomData || '{}');
            if (cd && typeof cd === 'object') Object.keys(cd).forEach(k => { if (p!.customData[k] === undefined && cd[k] !== null && cd[k] !== '') p!.customData[k] = String(cd[k]); });
          } catch { /* */ }
          if (childId && r.Status === 'Eingecheckt') { p.checkedIn.push(childId); checkInCount++; }
        }
      };
      // Klammer zuerst (übergreifende Antworten), dann die Termine.
      if (event.subsiteUrl) {
        let ok = true;
        const regs = await getAllRegistrations(event.id, () => { ok = false; });
        if (!ok) unreadable.push(event.title); else absorb(regs, null);
      }
      for (const k of kids) {
        if (!k.subsiteUrl) continue;
        let ok = true;
        const regs = await getAllRegistrations(k.id, () => { ok = false; });
        if (!ok) { unreadable.push(shortSubEventTitle(k.title, event.title) || k.title); continue; }
        absorb(regs, k.id);
      }
      if (cancelled) return;
      const persons = Array.from(byEmail.values()).sort((a, b) => a.last.localeCompare(b.last, 'de') || a.first.localeCompare(b.first, 'de'));
      setAnalysis({ items, persons, waitlisted: Array.from(waitlisted), unreadable, checkInCount });
      setPhase(unreadable.length > 0 ? 'blocked' : 'ready');
    })().catch(err => {
      console.warn('[DEX] CopyToAgenda analyze failed:', err);
      if (!cancelled) { setAnalysis({ items: [], persons: [], waitlisted: [], unreadable: [event.title], checkInCount: 0 }); setPhase('blocked'); }
    });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [event.id]);

  const run = async (): Promise<void> => {
    if (!analysis || phase !== 'ready') return;
    setPhase('running');
    const total = 2 + analysis.persons.length + analysis.checkInCount + (closeOld ? 1 : 0);
    let done = 0;
    const step = (label: string): void => { done++; setProgress({ done, total, label }); };
    const failed: Array<{ name: string; reason: string }> = [];
    let registered = 0; let checkIns = 0; let checkInFailed = 0; let oldClosed = false;
    try {
      // 2) Neues Event — Kommunikation der Klammer, Sub-Event-Flags raus.
      setProgress({ done, total, label: isDe ? 'Neues Event wird angelegt…' : 'Creating the new event…' });
      let ov: Record<string, unknown> = {};
      try { ov = JSON.parse(event.emailTemplateOverrides || '{}') || {}; } catch { ov = {}; }
      for (const k of ['_subEventsOnlyMode', '_subEventsDisabled', '_requireSubEventSelection', '_subEventCalendar', '_subEventSingleChoice', '_klammerDeadline', '_subEventOpenRule', '_subDeadlineRule', '_visAllSubs', '_commBundledMail', '_commBundledOutlook', '_commBundledQr', '_commShared', '_childEventTerm', '_mainEventLabel', '_hotels', '_hotelStays', '_hotelVisible', '_hotelRules', '_billing', '_shirtStock']) delete ov[k];
      ov._agendaCheckIn = true;
      if (event.childEventTermSingular || event.childEventTermPlural) ov._agendaTerm = { singular: event.childEventTermSingular || '', plural: event.childEventTermPlural || '' };
      const addr = event.locationAddress;
      const newId = await createEvent({
        title: newTitle.trim() || `${event.title} (Programmpunkte)`,
        type: event.type,
        status: 'Active',
        description: event.description || '',
        location: event.location || '',
        locationAddress: addr ? JSON.stringify(addr) : undefined,
        locationFilter: (event.locationAudience || []).join(','),
        audience: (event.audienceFilter || []).join(','),
        audienceResolvedEmails: (event.audienceResolvedEmails || []).join(';') || undefined,
        filterMode: event.filterMode || 'OR',
        startDate: event.startDate,
        endDate: event.endDate || event.startDate,
        registrationDeadline: event.registrationDeadline || '',
        lastDeregisterDate: event.lastDeregisterDate || '',
        maxParticipants: event.maxParticipants && event.maxParticipants > 0 ? Math.max(event.maxParticipants, analysis.persons.length) : 0,
        waitlistEnabled: !!event.waitlistEnabled,
        eventImageUrl: '',
        // v30.94: nur die überführende Person (s. useCurrentUser oben). Das
        // alte Team lässt sich im Wizard des neuen Events jederzeit nachtragen.
        organizer: (`${currentUser.surname || ''}, ${currentUser.firstName || ''}`.replace(/^,\s*|,\s*$/g, '').trim()) || (event.organizers || [])[0] || '',
        organizerEmail: (currentUser.email || '').trim() || (event.organizerEmails || [])[0] || '',
        contactName: event.contactName || '', contactEmail: event.contactEmail || '', contactInfo: event.contactInfo || '', contactOrganizerEmail: event.contactOrganizerEmail || '',
        outlookEventId: '',
        outlookBody: event.outlookBody || '',
        outlookSubject: event.outlookSubject || '',
        // Abfragefelder der Klammer und aller Termine (je Id einmal) — die
        // kopierten Antworten (customData) brauchen ihre Felder, sonst zeigt
        // sie keine Ansicht.
        customFields: ((): CustomField[] => {
          const seen = new Set<string>();
          const out: CustomField[] = [];
          [event, ...childEvents].forEach(ev => (ev.eventSpecificFields || []).forEach(f => {
            if (!f || !f.id || seen.has(f.id)) return;
            seen.add(f.id);
            out.push({ id: f.id, label: f.label, type: f.type, required: !!f.required, options: f.options, withTime: f.withTime, rangeStart: f.rangeStart, rangeEnd: f.rangeEnd, maxNights: f.maxNights, visible: true });
          }));
          return out;
        })(),
        agenda: JSON.stringify(analysis.items.map(({ childId: _c, ...it }) => it)),
        transfers: JSON.stringify(event.transferTimes || []),
        documents: '[]',
        funZone: '[]',
        emailLanguage: event.emailLanguage,
        registrationLanguage: event.registrationLanguage,
        emailTemplateOverrides: JSON.stringify(ov),
        disableEmails: !!event.disableEmails,
        disableOutlook: !!event.disableOutlook,
        excludedUsers: event.excludedUsers || [],
        isFictive: asTest,
        askSalutation: !!event.askSalutation,
      });
      if (!newId) throw new Error(isDe ? 'Das neue Event konnte nicht angelegt werden.' : 'The new event could not be created.');
      step(isDe ? 'Event angelegt — Liste wird geladen…' : 'Event created — loading list…');
      await refreshEvents();
      await new Promise(r => setTimeout(r, 400));
      step(isDe ? 'Anmeldungen werden kopiert…' : 'Copying registrations…');
      const newIdStr = String(newId);

      // 3) Anmeldungen still — sequentiell, mit Fehlerzähler.
      for (const p of analysis.persons) {
        try {
          const res = await registerRef.current(newIdStr, p.customData, p.first, p.last, p.email, undefined,
            { suppressMail: true, suppressOutlook: true, skipReload: true, skipShadowParent: true });
          if (res.ok) registered++; else failed.push({ name: p.name, reason: res.reason || res.status || 'unbekannt' });
        } catch (err) {
          failed.push({ name: p.name, reason: String((err as Error)?.message || err) });
        }
        step(p.name);
      }

      // 4) Check-ins → Anwesenheit am Punkt.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      const svc = ctx ? new EventService(ctx) : null;
      if (svc && analysis.checkInCount > 0) {
        const all = await svc.getEvents();
        const created = all.find(e => String(e.Id) === newIdStr);
        const newSubsite = created ? created.SubsiteUrl : '';
        if (newSubsite) {
          let readable = true;
          const newRegs = await getAllRegistrations(newIdStr, () => { readable = false; });
          const rowByEmail = new Map<string, number>();
          if (readable) newRegs.forEach(r => { const em = (r.ParticipantEmail || '').toLowerCase().trim(); if (em && ACTIVE.indexOf(r.Status || '') >= 0) rowByEmail.set(em, r.Id); });
          const itemByChild = new Map<string, string>();
          analysis.items.forEach(it => itemByChild.set(it.childId, it.id));
          for (const p of analysis.persons) {
            for (const childId of p.checkedIn) {
              const rowId = rowByEmail.get(p.email);
              const agendaId = itemByChild.get(childId);
              if (!rowId || !agendaId) { checkInFailed++; step(p.name); continue; }
              const r = await svc.checkInAgendaItem(newSubsite, rowId, agendaId);
              if (r.ok) checkIns++; else checkInFailed++;
              step(`${p.name} · ${isDe ? 'Anwesenheit' : 'attendance'}`);
            }
          }
        } else {
          checkInFailed = analysis.checkInCount;
        }
      }

      // 5) Optional: altes Event schließen (Fristen = jetzt). Kein Löschen.
      if (closeOld) {
        const nowIso = new Date().toISOString();
        let ok = await updateEvent(event.id, { 'RegistrationDeadline': nowIso }, { skipReload: true });
        for (const k of childEvents) {
          const r = await updateEvent(k.id, { 'RegistrationDeadline': nowIso }, { skipReload: true });
          ok = ok && r;
        }
        oldClosed = ok;
        step(isDe ? 'Altes Event geschlossen' : 'Old event closed');
      }

      // Audit an beiden Events.
      if (svc) {
        try {
          await svc.writeChangeLog({ action: 'CopiedToAgendaEvent', targetType: 'Event', targetId: newIdStr, targetName: newTitle, eventId: event.id, eventTitle: event.title, details: { persons: analysis.persons.length, registered, failed: failed.length, checkIns, checkInFailed, oldClosed } });
          await svc.writeChangeLog({ action: 'CreatedFromSubEvents', targetType: 'Event', targetId: event.id, targetName: event.title, eventId: newIdStr, eventTitle: newTitle, details: { items: analysis.items.length, persons: analysis.persons.length, registered, checkIns } });
        } catch { /* Audit best-effort */ }
      }
      await refreshEvents();
      setReport({ newId: newIdStr, registered, failed, checkIns, checkInFailed, oldClosed });
      setPhase('done');
      onDone();
    } catch (err) {
      console.warn('[DEX] CopyToAgenda failed:', err);
      showAlert(String((err as Error)?.message || err), { variant: 'error' });
      setPhase('ready');
    }
  };

  const termP = isDe ? 'Programmpunkte' : 'agenda items';
  const running = phase === 'running';
  // v31.2: Organizer des neuen Events ist seit v30.94 nur die kopierende Person — der Dialog soll das SAGEN (gleiche Auflösung wie in run()).
  const meName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || (currentUser.email || '').trim() || (event.organizers || [])[0] || '';
  const groups = analysis ? agendaGroups(analysis.items) : [];
  const checkedInAt = (childId: string): number => analysis ? analysis.persons.filter(p => p.checkedIn.indexOf(childId) >= 0).length : 0;
  const pct = progress.total ? Math.round(100 * progress.done / progress.total) : 0;
  const closeBtn = (cls: string): React.ReactElement => <button type="button" className={`btn ${cls}`} onClick={onClose}>{isDe ? 'Schließen' : 'Close'}</button>;

  // v31.2: Genau ein Primärknopf je Phase — der Fuß gehört dem Modal, nicht dem Inhalt.
  const footer = phase === 'done' ? closeBtn('btn-primary')
    : phase === 'blocked' ? closeBtn('btn-secondary')
    : (phase === 'ready' || running) && analysis ? <>
        <button type="button" className="btn btn-secondary" onClick={onClose} disabled={running}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" onClick={() => { void run(); }} disabled={running || analysis.items.length === 0}>
          {running ? (isDe ? 'Läuft…' : 'Running…') : (isDe ? 'Neues Event anlegen und kopieren' : 'Create new event and copy')}
        </button>
      </>
    : undefined;

  return (
    <Modal open onClose={running ? () => { /* läuft */ } : onClose} dismissable={!running} maxWidth={760}
      ariaLabel={isDe ? 'Als neues Event mit Programmpunkten kopieren' : 'Copy as new event with agenda items'}
      title={isDe ? 'Als neues Event mit Programmpunkten kopieren' : 'Copy as a new event with agenda items'}
      subtitle={isDe
        ? `„${event.title}" bleibt vollständig erhalten — Sub-Events, Listen, Mails, Check-ins. Es entsteht ein NEUES Event, dessen ${termP} die bisherigen Sub-Events sind.`
        : `“${event.title}” stays completely intact — sub-events, lists, mails, check-ins. A NEW event is created whose ${termP} are the former sub-events.`}
      icon={<Copy size={20} />} footer={footer}>

      {phase === 'analyzing' && (
        <div className="dex-ui-empty"><div className="dex-ui-empty-title">{isDe ? 'Teilnehmerlisten werden gelesen…' : 'Reading attendee lists…'}</div>
          {isDe ? 'Klammer zuerst, dann jeder Termin — eine nicht lesbare Liste zählt nicht als leer.' : 'Parent first, then each session — an unreadable list does not count as empty.'}</div>
      )}

      {phase === 'blocked' && analysis && (
        <div className="dex-ui-callout dex-ui-callout--danger">
          <span className="dex-ui-callout-icon"><AlertCircle size={18} /></span>
          <div><strong>{isDe ? 'Nicht gestartet.' : 'Not started.'}</strong>{' '}
            {isDe ? 'Diese Teilnehmerlisten konnten nicht gelesen werden — eine nicht lesbare Liste zählt nicht als leer:' : 'These attendee lists could not be read — an unreadable list does not count as empty:'}{' '}
            <strong>{analysis.unreadable.join(', ')}</strong>. {isDe ? 'Bitte Rechte prüfen („Organizer-Berechtigungen reparieren") und erneut versuchen.' : 'Please check permissions and try again.'}</div>
        </div>
      )}

      {(phase === 'ready' || running) && analysis && (
        <div className="dex-ui-modal-body">
          {/* 1) Was übernommen wird — Kennzahlen, Warteliste, Vorschau je Cluster */}
          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{isDe ? 'Was übernommen wird' : 'What is carried over'}</div>
            <div className="dex-ui-grid-auto" style={{ gap: 8 }}>
              {[
                { n: analysis.items.length, l: isDe ? `Sub-Events → ${termP}` : `Sub-events → ${termP}`, g: true },
                { n: analysis.persons.length, l: isDe ? 'Personen (angemeldet)' : 'People (registered)', g: true },
                { n: analysis.checkInCount, l: isDe ? 'Check-ins → Anwesenheiten' : 'Check-ins → attendances', g: true },
                { n: analysis.waitlisted.length, l: isDe ? 'Warteliste (nicht kopiert)' : 'Waitlist (not copied)', g: false },
              ].map((k, i) => (
                <div key={i} className={cx('dex-ui-kpi', k.n > 0 && (k.g ? 'dex-ui-kpi--green' : 'dex-ui-kpi--orange'))}>
                  <div className="dex-ui-kpi-value">{k.n}</div><div className="dex-ui-kpi-label">{k.l}</div>
                </div>
              ))}
            </div>
            <div className="dex-ui-help" style={{ marginTop: 8 }}>{isDe ? 'Die Anmeldungen werden still kopiert — keine Mail, kein Outlook-Termin. Formularantworten wandern in die neue Anmeldung.' : 'Registrations are copied silently — no mail, no Outlook invite. Form answers move to the new registration.'}</div>
            {analysis.waitlisted.length > 0 && (
              <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 10 }}><span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <div>{isDe ? 'Auf einer Warteliste stehen und werden NICHT kopiert (sie hatten keinen Platz): ' : 'On a waitlist and NOT copied (they had no seat): '}<strong>{analysis.waitlisted.join(', ')}</strong></div>
              </div>
            )}
            {/* v31.2: Vorschau der Punkte je Cluster — dieselbe Gruppierung wie im Editor (agendaGroups). */}
            <div style={{ marginTop: 12, maxHeight: 240, overflowY: 'auto', border: '1px solid var(--dex-gray-200, #e8e8e8)', borderRadius: 12 }}>
              {groups.length === 0 && <div className="dex-ui-muted" style={{ padding: 12 }}>{isDe ? 'Keine Sub-Events — es gibt nichts zu überführen.' : 'No sub-events — nothing to convert.'}</div>}
              {groups.map((g, gi) => (
                <div key={g.key}>
                  <div className="dex-ui-section-title" style={{ margin: 0, padding: '8px 12px 4px', background: 'var(--dex-gray-50, #fafafa)' }}><span>{groupLabel(g, gi, isDe)}</span>
                    <span style={{ fontWeight: 500, letterSpacing: 0, textTransform: 'none' }}>{groupDateLabel(g, isDe)}</span></div>
                  {g.items.map(it => { const ci = checkedInAt((it as Analysis['items'][number]).childId); return (
                    <div key={it.id} className="dex-ui-row dex-ui-row--bordered">
                      <span className="dex-ui-choice-icon" style={{ width: 30, height: 30 }}><Calendar size={16} strokeWidth={2} /></span>
                      <div className="dex-ui-row-main">
                        <div className="dex-ui-row-title">{it.title}</div>
                        <div className="dex-ui-row-sub">{[it.time && `${it.time}${it.endTime ? `–${it.endTime}` : ''}`, it.location].filter(Boolean).join(' · ')}</div>
                      </div>
                      {ci > 0 && <span className="dex-ui-pill dex-ui-pill--green">{ci} {isDe ? 'eingecheckt' : 'checked in'}</span>}
                    </div>
                  ); })}
                </div>
              ))}
            </div>
          </div>

          {/* 2) Das neue Event — Titel (Pflicht), dann die zwei Schalter */}
          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{isDe ? 'Das neue Event' : 'The new event'}</div>
            <div className="dex-ui-field">
              <label className="dex-ui-label" htmlFor="dex-copy-agenda-title">{isDe ? 'Wie soll das neue Event heißen?' : 'What should the new event be called?'}</label>
              <input id="dex-copy-agenda-title" type="text" className="dex-ui-input" value={newTitle} onChange={e => setNewTitle(e.target.value)} disabled={running} />
              <div className="dex-ui-help">{isDe ? <>Organizer des neuen Events bist nur du{meName ? <> (<strong>{meName}</strong>)</> : null} — das alte Team trägst du im Wizard des neuen Events nach.</>
                : <>You{meName ? <> (<strong>{meName}</strong>)</> : null} are the only organizer of the new event — add the old team in the new event&rsquo;s wizard.</>}</div>
            </div>
            <div className="dex-ui-stack">
              <label className={cx('dex-ui-toggle-row', asTest && 'is-active', running && 'is-disabled')}>
                <input type="checkbox" checked={asTest} onChange={e => setAsTest(e.target.checked)} disabled={running} />
                <span className="dex-ui-toggle-row-body"><span className="dex-ui-toggle-row-title">{isDe ? 'Als Test-Event anlegen' : 'Create as test event'}<span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'empfohlen' : 'recommended'}</span></span>
                  <span className="dex-ui-toggle-row-desc">{isDe ? 'Nur Admins und Organizer sehen es, bis du es live schaltest — Zeit, Programmpunkte und Anmeldungen zu prüfen.' : 'Only admins and organizers see it until you publish — time to check items and registrations.'}</span></span>
              </label>
              <label className={cx('dex-ui-toggle-row', closeOld && 'is-active', running && 'is-disabled')}>
                <input type="checkbox" checked={closeOld} onChange={e => setCloseOld(e.target.checked)} disabled={running} />
                <span className="dex-ui-toggle-row-body"><span className="dex-ui-toggle-row-title">{isDe ? 'Anmeldung am alten Event schließen' : 'Close registration on the old event'}</span>
                  <span className="dex-ui-toggle-row-desc">{isDe ? 'Setzt die Anmeldefrist des alten Events und seiner Sub-Events auf jetzt. Nichts wird gelöscht, keine Outlook-Termine werden ausgeladen.' : 'Sets the registration deadline of the old event and its sub-events to now. Nothing is deleted, no Outlook invites are cancelled.'}</span></span>
              </label>
            </div>
          </div>

          {/* 3) Selten gebraucht: was NICHT mitkommt — zu, bis jemand fragt */}
          <div>
            <button type="button" className={cx('dex-ui-disclosure', showSkipped && 'is-open')} onClick={() => setShowSkipped(v => !v)} aria-expanded={showSkipped}><span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
              {isDe ? 'Was nicht mitkopiert wird' : 'What is not copied'}<span className="dex-ui-disclosure-count">3</span></button>
            {showSkipped && <div className="dex-ui-disclosure-body dex-ui-muted dex-ui-fade-in">{isDe
              ? 'Nicht mitkopiert: das Event-Bild und Dokumente (bitte im neuen Event neu hochladen), eigene Kommunikationstexte der Sub-Events (das neue Event hat eine Kommunikation).'
              : 'Not copied: event image and documents (please re-upload), per-sub-event communication texts (the new event has one communication).'}</div>}
          </div>

          {running && (
            <div className="dex-ui-card dex-ui-card--soft" aria-live="polite">
              <div className="dex-ui-inline" style={{ justifyContent: 'space-between', marginBottom: 8 }}><span className="dex-ui-label" style={{ margin: 0 }}>{isDe ? 'Kopiert…' : 'Copying…'}</span>
                <span className="dex-ui-pill dex-ui-pill--green">{progress.done}/{progress.total} · {pct} %</span></div>
              <div style={{ height: 8, borderRadius: 999, background: 'var(--dex-gray-200, #e8e8e8)', overflow: 'hidden' }}><div style={{ height: 8, width: `${pct}%`, background: 'var(--dex-green, #86bc25)', transition: 'width 0.2s' }} /></div>
              <div className="dex-ui-help">{progress.label}</div>
            </div>
          )}
        </div>
      )}

      {phase === 'done' && report && analysis && (
        <div className="dex-ui-modal-body">
          <div className="dex-ui-callout dex-ui-callout--success">
            <span className="dex-ui-callout-icon"><Check size={18} /></span>
            <div><strong>{isDe ? 'Fertig.' : 'Done.'}</strong>{' '}
              {isDe
                ? `Neues Event „${newTitle}" mit ${analysis.items.length} ${termP} angelegt${asTest ? ' (Test-Event)' : ''}. ${report.registered} von ${analysis.persons.length} Anmeldungen kopiert, ${report.checkIns} Anwesenheiten übernommen${report.checkInFailed ? ` (${report.checkInFailed} nicht übertragbar)` : ''}.${report.oldClosed ? ' Anmeldung am alten Event geschlossen.' : ''}`
                : `New event “${newTitle}” with ${analysis.items.length} ${termP} created${asTest ? ' (test event)' : ''}. ${report.registered} of ${analysis.persons.length} registrations copied, ${report.checkIns} attendances carried over${report.checkInFailed ? ` (${report.checkInFailed} not transferable)` : ''}.${report.oldClosed ? ' Registration on the old event closed.' : ''}`}</div>
          </div>
          {report.failed.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--danger" style={{ maxHeight: 200, overflow: 'auto' }}>
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <div><strong>{isDe ? `${report.failed.length} Anmeldung(en) nicht kopiert:` : `${report.failed.length} registration(s) not copied:`}</strong>
                <ul style={{ margin: '4px 0 0', paddingLeft: 18 }}>{report.failed.map((f, i) => <li key={i}>{f.name} — {f.reason}</li>)}</ul>
                <div style={{ marginTop: 6 }}>{isDe ? 'Diese Personen kannst du im neuen Event über „Teilnehmer hinzufügen" nachtragen.' : 'You can add these people in the new event via “Add participants”.'}</div></div>
            </div>
          )}
          {/* v31.2: Die nächsten Schritte als nummerierte Zeilen — lesbar als Liste, nicht als Fließtext. */}
          <div className="dex-ui-stack">{(isDe
            ? ['Das neue Event im Organizer Center öffnen', 'Programmpunkte und Anwesenheit prüfen', 'Bild hochladen, dann live schalten']
            : ['Open the new event in the Organizer Center', 'Check items and attendance', 'Upload the image, then publish']
          ).map((t, i) => <div key={i} className="dex-ui-step is-pending"><span className="dex-ui-step-num">{i + 1}</span><div className="dex-ui-step-body"><div className="dex-ui-step-title">{t}</div></div></div>)}</div>
        </div>
      )}
    </Modal>
  );
}
