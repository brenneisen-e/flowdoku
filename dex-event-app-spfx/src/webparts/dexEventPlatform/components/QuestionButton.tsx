/**
 * QuestionButton (v26.0.0)
 *
 * Grüner „Hast du Fragen?"-Button im Header (alle eingeloggten User, seit
 * v26.34 auch auf der Landing Page). Öffnet ein Modal, in dem man:
 *  - eine oder MEHRERE Fragen stellen kann (per „+"),
 *  - schon beim Tippen passende Handbuch-Artikel vorgeschlagen bekommt
 *    (Selbsthilfe, bevor ein Ticket rausgeht — auch für normale User),
 *  - optional einen Screenshot des aktuellen Bildschirms anhängt,
 *  - unter „Deine Fragen" Status + Antworten der eigenen Tickets sieht.
 *
 * Ohne TicketProvider (z.B. in Handbuch-Previews) rendert die Komponente nichts.
 */
import * as React from 'react';
import Modal from './Modal';
import { Icon } from '@fluentui/react/lib/Icon';
import { TicketContext } from '../context/TicketContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { searchManual, openManualArticle, getManualSection, ManualArticle } from '../utils/manualSearch';
import { captureScreen } from '../utils/screenshot';
import { getActiveWizardStep } from '../utils/wizardStepContext';
import { DexTicket, DeloitteEvent } from '../types';
import PersonContactHover from './PersonContactHover';
import { renderTicketThread, contactSubline } from './tickets/ticketThread';
import ImageAnnotateModal from './ImageAnnotateModal';
import InquiryModal from './InquiryModal';

// v26.52: Live-Wizard-Vorschau der Antwort (mit Markierungsbox). MUSS lazy
// bleiben — QuestionButton ist im Main-Bundle, die Modal-Datei zieht aber
// EventCreationPage rein (im Haupt-Router ebenfalls lazy geladen).
const WizardStepPreviewModal = React.lazy(() => import('./tickets/WizardStepPreviewModal'));

interface ShotRef { file: File; url: string; }

export default function QuestionButton(props: { isMobile?: boolean }): React.ReactElement | null {
  const ticketCtx = React.useContext(TicketContext);
  const { locale } = useLanguage();
  const { navigate, selectedEventId } = useNavigation();
  const { currentUserRole, isAdmin } = useRoles();
  const { events, topLevelEvents, getMyEventNumbers, childEventsOf } = useEvents();
  const isDe = locale === 'de';

  const [open, setOpen] = React.useState(false);
  const [tab, setTab] = React.useState<'ask' | 'mine'>('ask');
  const [questions, setQuestions] = React.useState<string[]>(['']);
  // v26.30: Explizite Event-Zuordnung der Frage + eigene Event-Nummern (für die
  // Gruppe „Deine Events" im Auswahl-Dropdown).
  const [selEventId, setSelEventId] = React.useState<string>('');
  const [myNums, setMyNums] = React.useState<{ registered: number[]; waitlisted: number[] }>({ registered: [], waitlisted: [] });
  const [askWizardStep, setAskWizardStep] = React.useState<number | null>(null);
  // v26.60: Organizer unterscheiden beim Einreichen zwischen inhaltlicher
  // Frage (→ Power-User) und Bug-Report (→ DEX-Maintainer Nils + Eike).
  const [askCategory, setAskCategory] = React.useState<'question' | 'bug'>('question');
  const [shots, setShots] = React.useState<ShotRef[]>([]);
  // v29.40: Hinweis, wenn beim Anhängen etwas aussortiert wurde (kein Bild/zu groß).
  const [uploadNote, setUploadNote] = React.useState('');
  // v26.10: Index des Screenshots, der gerade groß markiert wird (null = keiner).
  const [annotateIdx, setAnnotateIdx] = React.useState<number | null>(null);
  // v26.10: „DEX selbst nutzen?"-Anfrage-Modal (nur für normale User).
  const [showInquiry, setShowInquiry] = React.useState(false);
  const [capturing, setCapturing] = React.useState(false);
  const [submitting, setSubmitting] = React.useState(false);
  const [done, setDone] = React.useState(false);
  const [hits, setHits] = React.useState<ManualArticle[]>([]);
  // v26.3: Vorschlag inline ausklappen (Antwort lesen, ohne ins Handbuch zu springen).
  const [openHitId, setOpenHitId] = React.useState<string | null>(null);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [hitSection, setHitSection] = React.useState<any>(null);
  const [hitLoading, setHitLoading] = React.useState(false);
  // v26.8: Rückfrage des Fragestellers auf eine Antwort (pro Ticket ein Entwurf).
  const [replyDrafts, setReplyDrafts] = React.useState<Record<number, string>>({});
  const [replyBusyId, setReplyBusyId] = React.useState<number | null>(null);
  // v26.52: Ticket, dessen Antwort-Wizard-Schritt gerade als Live-Vorschau
  // (mit Markierungsbox) geöffnet ist. null = zu.
  const [wizardPreviewTicket, setWizardPreviewTicket] = React.useState<DexTicket | null>(null);

  // Live-Handbuch-Suche (debounced) auf den eingegebenen Fragetext.
  const queryText = questions.join(' ').trim();
  React.useEffect(() => {
    if (!open || tab !== 'ask') { setHits([]); return; }
    if (queryText.length < 4) { setHits([]); return; }
    let alive = true;
    const tmr = window.setTimeout(() => {
      searchManual(queryText, isDe ? 'de' : 'en', 4)
        .then((r) => { if (alive) setHits(r); })
        .catch(() => { if (alive) setHits([]); });
    }, 350);
    return () => { alive = false; window.clearTimeout(tmr); };
  }, [queryText, open, tab, isDe]);

  // „Deine Fragen" beim Öffnen / Tab-Wechsel nachladen.
  React.useEffect(() => {
    if (open && tab === 'mine' && ticketCtx) ticketCtx.reloadMyTickets().catch(() => { /* */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, tab]);

  // v26.30: Beim Öffnen die eigenen Event-Nummern laden (für „Deine Events") und
  // das aktuell betrachtete Event (falls sichtbar) vorauswählen.
  React.useEffect(() => {
    if (!open) return;
    getMyEventNumbers().then(setMyNums).catch(() => { /* */ });
    const ctxId = selectedEventId || '';
    setSelEventId(ctxId && (events || []).some((e) => e.id === ctxId) ? ctxId : '');
    // v26.30: Wird aus dem Event-Wizard heraus gefragt, den Schritt zuordnen.
    setAskWizardStep(getActiveWizardStep());
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  // Object-URLs der Screenshots beim Unmount freigeben.
  React.useEffect(() => () => { shots.forEach((s) => { try { URL.revokeObjectURL(s.url); } catch { /* */ } }); }, [shots]);

  // v26.7: Deep-Link aus der Antwort-Mail (?action=ask) öffnet das Modal direkt
  // auf „Deine Fragen".
  React.useEffect(() => {
    const onOpen = (): void => { setTab('mine'); setOpen(true); };
    window.addEventListener('dex-open-questions', onOpen);
    return () => window.removeEventListener('dex-open-questions', onOpen);
  }, []);

  if (!ticketCtx) return null;

  const resetForm = (): void => {
    shots.forEach((s) => { try { URL.revokeObjectURL(s.url); } catch { /* */ } });
    setQuestions(['']); setShots([]); setHits([]); setDone(false); setUploadNote('');
  };
  const openModal = (): void => { resetForm(); setTab('ask'); setOpen(true); };
  const closeModal = (): void => { if (submitting || capturing) return; setOpen(false); resetForm(); };

  // Routing-Hinweis: wohin geht die Frage? (v26.30: explizit gewähltes Event)
  const ctxEvent = selEventId ? (events || []).find((e) => e.id === selEventId) : undefined;
  const askerIsOrgLike = currentUserRole === 'Organizer' || isAdmin;
  const ctxOrgs = ctxEvent
    ? Array.from(new Set([...(ctxEvent.organizerEmails || []), ...(ctxEvent.coOrganizerEmails || [])].map((x) => (x || '').trim()).filter(Boolean)))
    : [];
  const goesToOrganizer = !askerIsOrgLike && !!ctxEvent && ctxOrgs.length > 0;
  // v26.30: Auswahl-Optionen — „Deine Events" (angemeldet/Warteliste) vs. weitere
  // sichtbare Events. Test-Events (isFictive) nur für Organizer/Admins.
  const isMineFor = (ev: DeloitteEvent): boolean => {
    if (myNums.registered.indexOf(ev.eventNumber) >= 0 || myNums.waitlisted.indexOf(ev.eventNumber) >= 0) return true;
    if (ev.subEventsOnlyMode) {
      return childEventsOf(ev.id).some((k) => myNums.registered.indexOf(k.eventNumber) >= 0 || myNums.waitlisted.indexOf(k.eventNumber) >= 0);
    }
    return false;
  };
  // v26.34: Anzeige „Name | Datum" pro Option.
  const fmtEventOpt = (e: DeloitteEvent): string => {
    const d = e.startDate ? new Date(e.startDate) : null;
    const dateStr = d && !isNaN(d.getTime())
      ? d.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })
      : '';
    return dateStr ? `${e.title} | ${dateStr}` : e.title;
  };
  // v26.34: Nur AKTIVE Events, alphabetisch sortiert.
  const pickableEvents = (topLevelEvents || [])
    .filter((e) => (!e.isFictive || askerIsOrgLike) && e.status === 'Active')
    .slice()
    .sort((a, b) => (a.title || '').localeCompare(b.title || '', isDe ? 'de' : 'en'));
  const myEventsList = pickableEvents.filter(isMineFor);
  const otherEventsList = pickableEvents.filter((e) => !isMineFor(e));
  // v26.30: Vorausgewähltes Event (z.B. ein Sub-Event auf dem Register-Screen),
  // das nicht in der Top-Level-Liste steht, trotzdem als Option anbieten.
  const selExtra = (ctxEvent && !pickableEvents.some((e) => e.id === ctxEvent.id)) ? ctxEvent : undefined;
  const routingText = goesToOrganizer
    ? (isDe ? `Deine Frage geht an die Organisator:innen von „${ctxEvent!.title}".` : `Your question goes to the organizers of "${ctxEvent!.title}".`)
    : (isDe ? 'Deine Frage geht an das DEX-Support-Team (Power-User).' : 'Your question goes to the DEX support team (power users).');

  const doCapture = async (): Promise<void> => {
    setCapturing(true);
    // Modal kurz ausblenden (open && !capturing), damit der Screenshot den
    // Bildschirm DAHINTER erfasst, nicht das Modal.
    await new Promise((r) => window.setTimeout(r, 90));
    const f = await captureScreen(document.body);
    setCapturing(false);
    if (f) {
      const url = URL.createObjectURL(f);
      setShots((s) => [...s, { file: f, url }]);
    }
  };
  // v29.40: Bild vom Gerät anhängen. Der Live-Screenshot erfasst nur, was
  // GERADE im Portal steht — ein Foto von einer anderen Seite, einer Mail oder
  // dem Handy ging damit nicht (Feedback 21.08.2026). Beide Wege füllen
  // dieselbe Liste, das Markieren funktioniert danach genauso.
  const MAX_SHOT_BYTES = 8 * 1024 * 1024;
  const addFiles = (files: FileList | null): void => {
    if (!files || files.length === 0) return;
    const accepted: Array<{ file: File; url: string }> = [];
    let skipped = 0;
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      if (f.type.indexOf('image/') !== 0 || f.size > MAX_SHOT_BYTES) { skipped++; continue; }
      accepted.push({ file: f, url: URL.createObjectURL(f) });
    }
    if (accepted.length > 0) setShots((s) => [...s, ...accepted]);
    setUploadNote(skipped > 0
      ? (isDe
        ? `${skipped} Datei(en) übersprungen — es gehen nur Bilder bis 8 MB.`
        : `${skipped} file(s) skipped — images up to 8 MB only.`)
      : '');
  };
  const removeShot = (idx: number): void => {
    setShots((s) => { const copy = [...s]; const [rm] = copy.splice(idx, 1); if (rm) { try { URL.revokeObjectURL(rm.url); } catch { /* */ } } return copy; });
  };
  // v26.10: markiertes Bild übernehmen → ersetzt den Screenshot an annotateIdx.
  const onAnnotateSave = (file: File): void => {
    const idx = annotateIdx;
    if (idx == null) return;
    const url = URL.createObjectURL(file);
    setShots((s) => s.map((sh, i) => {
      if (i !== idx) return sh;
      try { URL.revokeObjectURL(sh.url); } catch { /* */ }
      return { file, url };
    }));
    setAnnotateIdx(null);
  };

  const setQuestion = (idx: number, val: string): void => setQuestions((qs) => qs.map((q, i) => (i === idx ? val : q)));
  const addQuestion = (): void => setQuestions((qs) => [...qs, '']);
  const removeQuestion = (idx: number): void => setQuestions((qs) => (qs.length <= 1 ? qs : qs.filter((_, i) => i !== idx)));

  const canSubmit = questions.some((q) => q.trim().length > 0) && !submitting && !capturing;

  const submit = async (): Promise<void> => {
    if (!canSubmit) return;
    setSubmitting(true);
    const ok = await ticketCtx.createTicket({
      questions: questions.map((q) => q.trim()).filter(Boolean),
      screenshots: shots.map((s) => s.file),
      eventId: selEventId,
      askWizardStep,
      category: askCategory,
    });
    setSubmitting(false);
    if (ok) {
      shots.forEach((s) => { try { URL.revokeObjectURL(s.url); } catch { /* */ } });
      setShots([]); setQuestions(['']); setHits([]); setDone(true); setAskCategory('question');
    }
  };

  const openArticle = (id: string): void => { setOpen(false); openManualArticle(id, navigate); };

  // Vorschlag auf-/zuklappen und den vollen Artikel-Inhalt lazy nachladen.
  const toggleHit = (id: string): void => {
    if (openHitId === id) { setOpenHitId(null); setHitSection(null); return; }
    setOpenHitId(id); setHitSection(null); setHitLoading(true);
    getManualSection(id, isDe ? 'de' : 'en')
      .then((sec) => { setHitSection(sec); setHitLoading(false); })
      .catch(() => setHitLoading(false));
  };

  // v26.8: Der Fragesteller antwortet auf die Antwort — die Rückfrage geht NUR
  // an die Person, die geantwortet hat (Routing in TicketContext.replyToTicket).
  const sendReply = async (tk: DexTicket): Promise<void> => {
    const txt = (replyDrafts[tk.id] || '').trim();
    if (!txt || !ticketCtx) return;
    setReplyBusyId(tk.id);
    const ok = await ticketCtx.replyToTicket(tk, txt);
    setReplyBusyId(null);
    if (ok) setReplyDrafts((d) => ({ ...d, [tk.id]: '' }));
  };

  // ---- Status-Badge für „Deine Fragen" ----
  const statusBadge = (st: DexTicket['status'], claimedBy: string): React.ReactElement => {
    const map: Record<string, { bg: string; col: string; label: string }> = {
      Open: { bg: '#fff3e0', col: '#b35a00', label: isDe ? 'Offen' : 'Open' },
      InProgress: { bg: '#e3f2fd', col: '#1565c0', label: isDe ? `In Bearbeitung${claimedBy ? ` (${claimedBy})` : ''}` : `In progress${claimedBy ? ` (${claimedBy})` : ''}` },
      Closed: { bg: '#e8f5e9', col: '#2e7d32', label: isDe ? 'Beantwortet' : 'Answered' },
    };
    const c = map[st] || map.Open;
    return <span style={{ padding: '2px 9px', borderRadius: 10, background: c.bg, color: c.col, fontSize: 12, fontWeight: 600 }}>{c.label}</span>;
  };

  const btnStyle: React.CSSProperties = props.isMobile
    ? { display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: 36, height: 36, borderRadius: 8, background: 'var(--dex-green, #86bc25)', color: '#fff', border: 'none', cursor: 'pointer' }
    : { display: 'inline-flex', alignItems: 'center', gap: 8, height: 36, padding: '0 14px', borderRadius: 8, background: 'var(--dex-green, #86bc25)', color: '#fff', border: 'none', cursor: 'pointer', fontWeight: 600, fontSize: '0.85rem', fontFamily: 'inherit', boxShadow: '0 1px 3px rgba(0,0,0,0.12)' };

  return (
    <>
      <button
        type="button"
        onClick={openModal}
        title={isDe ? 'Hast du Fragen? Frag das DEX-Team' : 'Have a question? Ask the DEX team'}
        aria-label={isDe ? 'Hast du Fragen?' : 'Have a question?'}
        style={btnStyle}
      >
        <Icon iconName="Chat" style={{ fontSize: 16 }} />
        {!props.isMobile && <span style={{ lineHeight: 1 }}>{isDe ? 'Hast du Fragen?' : 'Have a question?'}</span>}
      </button>

      <Modal open={open && !capturing} onClose={closeModal} maxWidth={640} dismissable={!submitting} ariaLabel={isDe ? 'Hast du Fragen?' : 'Have a question?'}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <Icon iconName="Chat" style={{ fontSize: 22, color: 'var(--dex-green, #86bc25)' }} />
          <h2 style={{ margin: 0, fontSize: '1.25rem' }}>{isDe ? 'Hast du Fragen?' : 'Have a question?'}</h2>
        </div>

        {/* Tabs */}
        <div style={{ display: 'flex', gap: 6, borderBottom: '1px solid var(--dex-gray-200,#e8e8e8)' }}>
          {(['ask', 'mine'] as const).map((tb) => (
            <button
              key={tb}
              type="button"
              onClick={() => setTab(tb)}
              style={{
                background: 'transparent', border: 'none', cursor: 'pointer', fontFamily: 'inherit',
                padding: '8px 12px', fontSize: '0.9rem', fontWeight: 600,
                color: tab === tb ? 'var(--dex-green-dark,#4a7c1f)' : 'var(--dex-gray-500,#808080)',
                borderBottom: tab === tb ? '2px solid var(--dex-green,#86bc25)' : '2px solid transparent', marginBottom: -1,
              }}
            >
              {tb === 'ask' ? (isDe ? 'Frage stellen' : 'Ask a question') : (isDe ? 'Deine Fragen' : 'Your questions')}
            </button>
          ))}
        </div>

        {tab === 'ask' && !done && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <p style={{ margin: 0, fontSize: '0.86rem', color: 'var(--dex-gray-600,#666)' }}>
              {isDe
                ? 'Stell deine Frage(n) — wir schauen direkt, ob das Handbuch schon weiterhilft. Wenn nicht, schickst du die Frage als Ticket ab.'
                : 'Ask your question(s) — we instantly check whether the manual already helps. If not, you submit it as a ticket.'}
            </p>

            {/* v26.30: Zu welchem Event gehört die Frage? Steuert das Routing —
                Event mit Organizern → an dessen Organizer; „kein Event" oder ein
                Event ohne Organizer → an das DEX-Support-Team (Power-User). */}
            {/* v26.60: Organizer/Admins unterscheiden beim Einreichen zwischen
                inhaltlicher Frage (→ Power-User) und Bug (→ DEX-Maintainer). */}
            {(currentUserRole === 'Organizer' || isAdmin) && (
              <div>
                <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700,#444)', marginBottom: 4 }}>
                  {isDe ? 'Worum geht es?' : 'What is it about?'}
                </label>
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                  <button type="button" onClick={() => setAskCategory('question')}
                    style={{ flex: '1 1 180px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 7, padding: '9px 12px', borderRadius: 8, cursor: 'pointer', fontFamily: 'inherit', fontSize: '0.85rem', fontWeight: 600,
                      border: askCategory === 'question' ? '2px solid var(--dex-green,#86bc25)' : '1px solid var(--dex-gray-300,#d1d1d1)',
                      background: askCategory === 'question' ? '#f1f7e8' : '#fff',
                      color: askCategory === 'question' ? 'var(--dex-green-dark,#4a7c1f)' : 'var(--dex-gray-600,#666)' }}>
                    <Icon iconName="Help" style={{ fontSize: 14 }} />
                    {isDe ? 'Inhaltliche Frage' : 'Content question'}
                  </button>
                  <button type="button" onClick={() => setAskCategory('bug')}
                    style={{ flex: '1 1 180px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 7, padding: '9px 12px', borderRadius: 8, cursor: 'pointer', fontFamily: 'inherit', fontSize: '0.85rem', fontWeight: 600,
                      border: askCategory === 'bug' ? '2px solid #ed8b00' : '1px solid var(--dex-gray-300,#d1d1d1)',
                      background: askCategory === 'bug' ? '#fff8ef' : '#fff',
                      color: askCategory === 'bug' ? '#b35a00' : 'var(--dex-gray-600,#666)' }}>
                    <Icon iconName="Bug" style={{ fontSize: 14 }} />
                    {isDe ? 'Bug / Fehler melden' : 'Report a bug'}
                  </button>
                </div>
                <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500,#808080)', marginTop: 4 }}>
                  {askCategory === 'bug'
                    ? (isDe ? 'Bug-Reports gehen direkt an das DEX-Entwicklungsteam.' : 'Bug reports go directly to the DEX development team.')
                    : (isDe ? 'Inhaltliche Fragen beantworten die DEX-Power-User.' : 'Content questions are answered by the DEX power users.')}
                </div>
              </div>
            )}

            <div>
              <label htmlFor="dex-q-event" style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700,#444)', marginBottom: 4 }}>
                {isDe ? 'Zu welchem Event ist deine Frage?' : 'Which event is your question about?'}
              </label>
              <select
                id="dex-q-event"
                value={selEventId}
                onChange={(e) => setSelEventId(e.target.value)}
                style={{ width: '100%', padding: '9px 11px', borderRadius: 8, border: '1px solid var(--dex-gray-300,#d1d1d1)', fontFamily: 'inherit', fontSize: '0.9rem', background: '#fff' }}
              >
                <option value="">{isDe ? 'Kein bestimmtes Event – allgemeine Frage' : 'No specific event – general question'}</option>
                {selExtra && (<option value={selExtra.id}>{fmtEventOpt(selExtra)}</option>)}
                {myEventsList.length > 0 && (
                  <optgroup label={isDe ? 'Deine Events' : 'Your events'}>
                    {myEventsList.map((e) => (<option key={e.id} value={e.id}>{fmtEventOpt(e)}</option>))}
                  </optgroup>
                )}
                {otherEventsList.length > 0 && (
                  <optgroup label={isDe ? 'Weitere Events' : 'Other events'}>
                    {otherEventsList.map((e) => (<option key={e.id} value={e.id}>{fmtEventOpt(e)}</option>))}
                  </optgroup>
                )}
              </select>
            </div>

            {/* v26.30: Herkunft aus dem Event-Wizard — wird der Antwort mitgegeben. */}
            {askWizardStep != null && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, background: '#f1f7e8', border: '1px solid var(--dex-green,#86bc25)', borderRadius: 8, padding: '7px 10px' }}>
                <Icon iconName="DocumentManagement" style={{ fontSize: 14, color: 'var(--dex-green-dark,#4a7c1f)' }} />
                <span style={{ flex: 1, fontSize: '0.82rem', color: 'var(--dex-gray-700,#444)' }}>
                  {isDe ? `Frage aus dem Event-Wizard – Schritt ${askWizardStep} (wird für die Antwort mitgegeben).` : `Question from the event wizard – step ${askWizardStep} (shared with the responder).`}
                </span>
                <button type="button" onClick={() => setAskWizardStep(null)} title={isDe ? 'Wizard-Schritt nicht mitschicken' : 'Do not include the wizard step'}
                  style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400,#a0a0a0)', padding: 2 }}>
                  <Icon iconName="Cancel" style={{ fontSize: 13 }} />
                </button>
              </div>
            )}

            {questions.map((q, idx) => (
              <div key={idx} style={{ display: 'flex', gap: 8, alignItems: 'flex-start' }}>
                <textarea
                  value={q}
                  onChange={(e) => setQuestion(idx, e.target.value)}
                  placeholder={isDe ? (idx === 0 ? 'Deine Frage …' : 'Weitere Frage …') : (idx === 0 ? 'Your question …' : 'Another question …')}
                  rows={2}
                  style={{ flex: 1, padding: '9px 11px', borderRadius: 8, border: '1px solid var(--dex-gray-300,#d1d1d1)', fontFamily: 'inherit', fontSize: '0.9rem', resize: 'vertical' }}
                />
                {questions.length > 1 && (
                  <button type="button" onClick={() => removeQuestion(idx)} title={isDe ? 'Frage entfernen' : 'Remove question'}
                    style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400,#a0a0a0)', padding: 4, marginTop: 4 }}>
                    <Icon iconName="Cancel" style={{ fontSize: 14 }} />
                  </button>
                )}
              </div>
            ))}
            <button type="button" onClick={addQuestion}
              style={{ alignSelf: 'flex-start', display: 'inline-flex', alignItems: 'center', gap: 6, background: 'transparent', border: '1px dashed var(--dex-green,#86bc25)', color: 'var(--dex-green-dark,#4a7c1f)', borderRadius: 8, padding: '6px 12px', cursor: 'pointer', fontWeight: 600, fontSize: '0.82rem', fontFamily: 'inherit' }}>
              <Icon iconName="Add" style={{ fontSize: 13 }} /> {isDe ? 'Weitere Frage' : 'Add another question'}
            </button>

            {/* Live-Handbuch-Vorschläge */}
            {hits.length > 0 && (
              <div style={{ background: '#f1f7e8', border: '1px solid var(--dex-green,#86bc25)', borderRadius: 8, padding: '10px 12px' }}>
                <div style={{ fontSize: '0.82rem', fontWeight: 700, color: 'var(--dex-green-dark,#4a7c1f)', marginBottom: 6 }}>
                  {isDe ? 'Vielleicht hilft dir das schon weiter:' : 'This might already help you:'}
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                  {hits.map((h) => {
                    const isOpenHit = openHitId === h.id;
                    return (
                      <div key={h.id} style={{ background: '#fff', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 6, overflow: 'hidden' }}>
                        <button type="button" onClick={() => toggleHit(h.id)}
                          style={{ width: '100%', textAlign: 'left', background: 'transparent', border: 'none', padding: '7px 10px', cursor: 'pointer', fontFamily: 'inherit', display: 'flex', alignItems: 'flex-start', gap: 8 }}>
                          <Icon iconName="ReadingMode" style={{ fontSize: 13, color: 'var(--dex-green,#86bc25)', marginTop: 2 }} />
                          <span style={{ flex: 1 }}>
                            <strong style={{ fontSize: '0.85rem' }}>{h.title}</strong>
                            {h.description && <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500,#808080)', marginTop: 2 }}>{h.description}</div>}
                          </span>
                          <Icon iconName={isOpenHit ? 'ChevronUp' : 'ChevronDown'} style={{ fontSize: 12, color: 'var(--dex-gray-400,#a0a0a0)', marginTop: 3 }} />
                        </button>
                        {isOpenHit && (
                          <div style={{ borderTop: '1px solid var(--dex-gray-100,#f5f5f5)', padding: '6px 10px 10px' }}>
                            {hitLoading && <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400,#a0a0a0)', padding: '6px 0' }}>{isDe ? 'Lädt …' : 'Loading …'}</div>}
                            {!hitLoading && hitSection && renderArticleBody(hitSection, isDe)}
                            {!hitLoading && !hitSection && <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400,#a0a0a0)' }}>{isDe ? 'Inhalt nicht verfügbar.' : 'Content unavailable.'}</div>}
                            <button type="button" onClick={() => openArticle(h.id)}
                              style={{ marginTop: 8, background: 'transparent', border: 'none', color: 'var(--dex-green-dark,#4a7c1f)', cursor: 'pointer', fontWeight: 600, fontSize: '0.8rem', padding: 0, fontFamily: 'inherit', display: 'inline-flex', alignItems: 'center', gap: 5 }}>
                              <Icon iconName="OpenInNewWindow" style={{ fontSize: 12 }} /> {isDe ? 'Ganzen Artikel im Handbuch öffnen' : 'Open full article in the manual'}
                            </button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Screenshots */}
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                <button type="button" onClick={doCapture} disabled={capturing}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: 7, background: 'var(--dex-gray-100,#f5f5f5)', border: '1px solid var(--dex-gray-300,#d1d1d1)', borderRadius: 8, padding: '7px 12px', cursor: 'pointer', fontWeight: 600, fontSize: '0.82rem', fontFamily: 'inherit', color: 'var(--dex-gray-700,#444)' }}>
                  <Icon iconName="Camera" style={{ fontSize: 14 }} /> {isDe ? 'Screenshot des Bildschirms anhängen' : 'Attach a screenshot of the screen'}
                </button>
                {/* v29.40: Zweiter Weg — Bild vom Gerät. Als <label> um ein
                    verstecktes File-Input, damit der Knopf wie der daneben
                    aussieht (ein nacktes File-Input tut das nie). */}
                <label
                  style={{ display: 'inline-flex', alignItems: 'center', gap: 7, background: 'var(--dex-gray-100,#f5f5f5)', border: '1px solid var(--dex-gray-300,#d1d1d1)', borderRadius: 8, padding: '7px 12px', cursor: 'pointer', fontWeight: 600, fontSize: '0.82rem', color: 'var(--dex-gray-700,#444)' }}>
                  <Icon iconName="Photo2" style={{ fontSize: 14 }} /> {isDe ? 'Bild vom Gerät anhängen' : 'Attach an image from your device'}
                  <input
                    type="file"
                    accept="image/*"
                    multiple
                    style={{ display: 'none' }}
                    onChange={(e) => { addFiles(e.target.files); e.target.value = ''; }}
                  />
                </label>
                <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400,#a0a0a0)' }}>{isDe ? 'optional' : 'optional'}</span>
              </div>
              {uploadNote && (
                <div style={{ fontSize: '0.74rem', color: 'var(--dex-red, #da291c)', marginTop: 6 }}>{uploadNote}</div>
              )}
              {shots.length > 0 && (
                <>
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 10 }}>
                    {shots.map((s, idx) => (
                      <div key={idx} style={{ position: 'relative', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 8, overflow: 'hidden' }}>
                        <button type="button" onClick={() => setAnnotateIdx(idx)} title={isDe ? 'Vergrößern & markieren' : 'Enlarge & mark up'}
                          style={{ display: 'block', padding: 0, border: 'none', background: 'transparent', cursor: 'pointer' }}>
                          <img src={s.url} alt={`Screenshot ${idx + 1}`} style={{ width: 120, height: 76, objectFit: 'cover', display: 'block' }} />
                          <span style={{ position: 'absolute', bottom: 2, left: 2, background: 'rgba(0,0,0,0.55)', color: '#fff', borderRadius: 4, padding: '1px 5px', display: 'inline-flex', alignItems: 'center', gap: 3, fontSize: 10 }}>
                            <Icon iconName="InsertTextBox" style={{ fontSize: 11 }} /> {isDe ? 'markieren' : 'mark up'}
                          </span>
                        </button>
                        <button type="button" onClick={() => removeShot(idx)} title={isDe ? 'Entfernen' : 'Remove'}
                          style={{ position: 'absolute', top: 2, right: 2, background: 'rgba(0,0,0,0.55)', color: '#fff', border: 'none', borderRadius: '50%', width: 20, height: 20, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}>
                          <Icon iconName="Cancel" style={{ fontSize: 11 }} />
                        </button>
                      </div>
                    ))}
                  </div>
                  <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-400,#a0a0a0)', marginTop: 6 }}>
                    {isDe ? 'Tipp: Auf ein Bild klicken, um es groß anzuzeigen und die gemeinte Stelle zu markieren.' : 'Tip: click an image to enlarge it and mark the relevant area.'}
                  </div>
                </>
              )}
            </div>

            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500,#808080)', display: 'flex', alignItems: 'center', gap: 7 }}>
              <Icon iconName="Send" style={{ fontSize: 13, color: 'var(--dex-green,#86bc25)' }} /> {routingText}
            </div>

            {/* v26.10: Vorschlag NUR für normale User (Organizer/Admins können
                Events ohnehin selbst anlegen) — „DEX selbst nutzen?" öffnet das
                bestehende DEX-Anfrage-Modal. */}
            {!askerIsOrgLike && (
              <div style={{ background: '#f1f7e8', border: '1px solid var(--dex-green,#86bc25)', borderRadius: 8, padding: '10px 12px', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                <Icon iconName="Lightbulb" style={{ fontSize: 18, color: 'var(--dex-green-dark,#4a7c1f)' }} />
                <span style={{ flex: 1, minWidth: 200, fontSize: '0.84rem', color: 'var(--dex-gray-700,#444)' }}>
                  {isDe ? 'Möchtest du DEX selbst für ein Event nutzen?' : 'Would you like to use DEX for your own event?'}
                </span>
                <button type="button" onClick={() => setShowInquiry(true)}
                  style={{ background: 'var(--dex-green,#86bc25)', color: '#fff', border: 'none', borderRadius: 8, padding: '7px 14px', cursor: 'pointer', fontWeight: 600, fontSize: '0.82rem', fontFamily: 'inherit' }}>
                  {isDe ? 'DEX anfragen' : 'Request DEX'}
                </button>
              </div>
            )}

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 4 }}>
              <button className="btn btn-secondary" onClick={closeModal} disabled={submitting}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
              <button className="btn btn-primary" onClick={submit} disabled={!canSubmit}>
                {submitting ? (isDe ? 'Wird gesendet …' : 'Sending …') : (isDe ? 'Frage absenden' : 'Send question')}
              </button>
            </div>
          </div>
        )}

        {tab === 'ask' && done && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12, alignItems: 'center', padding: '12px 0' }}>
            <Icon iconName="CompletedSolid" style={{ fontSize: 40, color: 'var(--dex-green,#86bc25)' }} />
            <p style={{ margin: 0, textAlign: 'center', fontSize: '0.95rem' }}>
              {isDe ? 'Danke! Deine Frage wurde übermittelt. Du bekommst die Antwort per E-Mail — und findest sie auch hier unter „Deine Fragen".' : 'Thanks! Your question was submitted. You will get the answer by email — and also here under "Your questions".'}
            </p>
            <div style={{ display: 'flex', gap: 10 }}>
              <button className="btn btn-secondary" onClick={() => { resetForm(); setDone(false); }}>{isDe ? 'Weitere Frage' : 'Ask another'}</button>
              <button className="btn btn-primary" onClick={() => { setDone(false); setTab('mine'); }}>{isDe ? 'Deine Fragen ansehen' : 'View your questions'}</button>
            </div>
          </div>
        )}

        {tab === 'mine' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10, maxHeight: '55vh', overflowY: 'auto' }}>
            {ticketCtx.myTickets.length === 0 && (
              <p style={{ margin: 0, color: 'var(--dex-gray-500,#808080)', fontSize: '0.88rem' }}>
                {isDe ? 'Du hast noch keine Fragen gestellt.' : 'You have not asked any questions yet.'}
              </p>
            )}
            {ticketCtx.myTickets.map((tk) => (
              <div key={tk.id} style={{ border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 8, padding: '10px 12px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                  {statusBadge(tk.status, tk.claimedByName)}
                  {tk.eventTitle && <span style={{ fontSize: 12, color: 'var(--dex-gray-400,#a0a0a0)' }}>{tk.eventTitle}</span>}
                </div>
                {tk.questions.map((q, i) => (<div key={i} style={{ fontSize: '0.88rem', marginBottom: 2 }}>• {q}</div>))}
                {!!tk.answeredAt && (
                  <div style={{ marginTop: 8, background: '#f1f7e8', borderRadius: 6, padding: '8px 10px' }}>
                    {/* v26.8: Foto-Kontaktkarte der/des Beantwortenden (Hover → Teams-Chat). */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                      <span style={{ fontSize: 12, fontWeight: 700, color: 'var(--dex-green-dark,#4a7c1f)' }}>{isDe ? 'Antwort' : 'Answer'}</span>
                      {tk.answeredByEmail && (
                        <PersonContactHover email={tk.answeredByEmail} name={tk.answeredByName || tk.answeredByEmail} size={26}
                          subline={contactSubline(tk.answeredByJobTitle, tk.answeredByLocation)} isDe={isDe} />
                      )}
                      {tk.answeredByName && (
                        <span style={{ fontSize: 12, color: 'var(--dex-gray-600,#666)' }}>
                          {tk.answeredByName}
                          {contactSubline(tk.answeredByJobTitle, tk.answeredByLocation) ? ` · ${contactSubline(tk.answeredByJobTitle, tk.answeredByLocation)}` : ''}
                        </span>
                      )}
                    </div>
                    {tk.answerText && <div style={{ fontSize: '0.85rem', whiteSpace: 'pre-wrap' }}>{tk.answerText}</div>}
                    {tk.answerArticleIds.length > 0 && (
                      <div style={{ marginTop: 6, display: 'flex', flexDirection: 'column', gap: 3 }}>
                        {tk.answerArticleIds.map((aid) => (
                          <button key={aid} type="button" onClick={() => openArticle(aid)}
                            style={{ textAlign: 'left', background: 'transparent', border: 'none', color: 'var(--dex-green-dark,#4a7c1f)', cursor: 'pointer', fontWeight: 600, fontSize: '0.82rem', padding: 0, fontFamily: 'inherit' }}>
                            <Icon iconName="ReadingMode" style={{ fontSize: 12, marginRight: 5 }} />{isDe ? 'Handbuch-Artikel öffnen' : 'Open manual article'}
                          </button>
                        ))}
                      </div>
                    )}
                    {tk.answerWizardStep != null && (
                      <div style={{ marginTop: 6, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                        <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600,#666)' }}>
                          <Icon iconName="DocumentManagement" style={{ fontSize: 12, marginRight: 5 }} />
                          {isDe ? `Im Event-Wizard: Schritt ${tk.answerWizardStep}` : `In the event wizard: step ${tk.answerWizardStep}`}
                        </span>
                        {/* v26.52: Echten Wizard-Schritt als Read-only-Vorschau öffnen —
                            inkl. Markierungsbox („hier klicken"), falls gesetzt. */}
                        <button type="button" onClick={() => setWizardPreviewTicket(tk)}
                          style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: tk.answerWizardMarker ? '#fff8ef' : '#fff', border: `1px solid ${tk.answerWizardMarker ? '#ed8b00' : 'var(--dex-green,#86bc25)'}`, borderRadius: 6, padding: '4px 9px', cursor: 'pointer', fontFamily: 'inherit', fontSize: '0.78rem', fontWeight: 600, color: tk.answerWizardMarker ? '#b35a00' : 'var(--dex-green-dark,#4a7c1f)' }}>
                          <Icon iconName="Preview" style={{ fontSize: 11 }} />
                          {tk.answerWizardMarker
                            ? (isDe ? 'Vorschau mit Markierung ansehen' : 'View preview with marker')
                            : (isDe ? 'Vorschau ansehen' : 'View preview')}
                        </button>
                      </div>
                    )}
                  </div>
                )}
                {/* v26.8: Rückfragen-Verlauf */}
                {renderTicketThread(tk, isDe)}
                {/* v26.8: Auf die Antwort erneut antworten — geht NUR an die Person,
                    die geantwortet hat. */}
                {!!tk.answeredAt && (
                  <div style={{ marginTop: 8, display: 'flex', flexDirection: 'column', gap: 6 }}>
                    <textarea value={replyDrafts[tk.id] || ''} onChange={(e) => setReplyDrafts((d) => ({ ...d, [tk.id]: e.target.value }))} rows={2}
                      placeholder={isDe ? 'Auf die Antwort antworten …' : 'Reply to the answer …'}
                      style={{ padding: '8px 10px', borderRadius: 8, border: '1px solid var(--dex-gray-300,#d1d1d1)', fontFamily: 'inherit', fontSize: '0.85rem', resize: 'vertical' }} />
                    <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                      <button type="button" onClick={() => sendReply(tk)} disabled={replyBusyId === tk.id || !(replyDrafts[tk.id] || '').trim()}
                        style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: 'var(--dex-green,#86bc25)', color: '#fff', border: 'none', borderRadius: 8, padding: '6px 14px', cursor: 'pointer', fontWeight: 600, fontSize: '0.82rem', fontFamily: 'inherit', opacity: (replyBusyId === tk.id || !(replyDrafts[tk.id] || '').trim()) ? 0.6 : 1 }}>
                        <Icon iconName="Send" style={{ fontSize: 13 }} />
                        {replyBusyId === tk.id ? (isDe ? 'Wird gesendet …' : 'Sending …') : (isDe ? 'Antwort senden' : 'Send reply')}
                      </button>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </Modal>

      {/* v26.52: Live-Wizard-Vorschau der Antwort (read-only, mit Markierung). */}
      {wizardPreviewTicket && wizardPreviewTicket.answerWizardStep != null && (
        <React.Suspense fallback={null}>
          <WizardStepPreviewModal
            step={wizardPreviewTicket.answerWizardStep}
            isDe={isDe}
            editable={false}
            initialMarker={wizardPreviewTicket.answerWizardMarker}
            onClose={() => setWizardPreviewTicket(null)}
          />
        </React.Suspense>
      )}

      {/* v26.10: Screenshot vergrößern + transparente Boxen darübersetzen. */}
      <ImageAnnotateModal
        open={annotateIdx != null && !!shots[annotateIdx as number]}
        src={annotateIdx != null && shots[annotateIdx] ? shots[annotateIdx].url : ''}
        fileName={annotateIdx != null && shots[annotateIdx] ? shots[annotateIdx].file.name : 'screenshot.png'}
        isDe={isDe}
        onClose={() => setAnnotateIdx(null)}
        onSave={onAnnotateSave}
      />

      {/* v26.10: DEX-Anfrage-Modal (nur für normale User aus der Vorschlags-Box). */}
      <InquiryModal open={showInquiry} onClose={() => setShowInquiry(false)} />
    </>
  );
}

/* eslint-disable @typescript-eslint/no-explicit-any */
// v26.3: Inhalt einer Handbuch-Sektion kompakt inline rendern (Perspektiven →
// Schritte). Nur Text-/Prosa-Teile (intro, title, description, tip, warning) —
// KEINE liveDemo/mockup-Renders, damit nichts auf fehlende Kontexte zugreift.
function renderArticleBody(section: any, isDe: boolean): React.ReactElement {
  const persps: any[] = section.perspectives || [];
  const perspLabel = (p: any): string =>
    p.title || (p.perspective === 'organizer'
      ? (isDe ? 'Als Organizer' : 'As organizer')
      : p.perspective === 'admin'
        ? (isDe ? 'Als Admin' : 'As admin')
        : (isDe ? 'Als Teilnehmer' : 'As participant'));
  return (
    <div style={{ maxHeight: 300, overflowY: 'auto', fontSize: '0.84rem', lineHeight: 1.5 }}>
      {section.description && <p style={{ margin: '4px 0 10px', color: 'var(--dex-gray-700,#444)' }}>{section.description}</p>}
      {persps.map((p: any, pi: number) => (
        <div key={pi} style={{ marginBottom: 12 }}>
          {persps.length > 1 && (
            <div style={{ fontSize: '0.72rem', fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.4, color: 'var(--dex-green-dark,#4a7c1f)', marginBottom: 5 }}>{perspLabel(p)}</div>
          )}
          {p.intro && <div style={{ marginBottom: 7, color: 'var(--dex-gray-700,#444)' }}>{p.intro}</div>}
          {(p.steps || []).map((st: any, si: number) => (
            <div key={si} style={{ display: 'flex', gap: 8, marginBottom: 7 }}>
              <span style={{ flexShrink: 0, width: 19, height: 19, borderRadius: '50%', background: 'var(--dex-green,#86bc25)', color: '#fff', fontSize: 11, fontWeight: 700, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', marginTop: 1 }}>{st.number}</span>
              <div>
                {st.title && <strong>{st.title}</strong>}
                {st.description && <div style={{ color: 'var(--dex-gray-700,#444)' }}>{st.description}</div>}
                {st.tip && <div style={{ marginTop: 3, fontSize: '0.78rem', color: 'var(--dex-orange-dark,#b35a00)' }}>{isDe ? 'Tipp: ' : 'Tip: '}{st.tip}</div>}
                {st.warning && <div style={{ marginTop: 3, fontSize: '0.78rem', color: 'var(--dex-red,#da291c)' }}>{st.warning}</div>}
              </div>
            </div>
          ))}
        </div>
      ))}
    </div>
  );
}
/* eslint-enable @typescript-eslint/no-explicit-any */
