/**
 * QuestionButton (v26.0.0)
 *
 * Grüner „Hast du Fragen?"-Button im Header (alle eingeloggten User, nicht auf
 * der Landing Page). Öffnet ein Modal, in dem man:
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
import { searchManual, openManualArticle, ManualArticle } from '../utils/manualSearch';
import { captureScreen } from '../utils/screenshot';
import { DexTicket } from '../types';

interface ShotRef { file: File; url: string; }

export default function QuestionButton(props: { isMobile?: boolean }): React.ReactElement | null {
  const ticketCtx = React.useContext(TicketContext);
  const { locale } = useLanguage();
  const { navigate, selectedEventId } = useNavigation();
  const { currentUserRole, isAdmin } = useRoles();
  const { events } = useEvents();
  const isDe = locale === 'de';

  const [open, setOpen] = React.useState(false);
  const [tab, setTab] = React.useState<'ask' | 'mine'>('ask');
  const [questions, setQuestions] = React.useState<string[]>(['']);
  const [shots, setShots] = React.useState<ShotRef[]>([]);
  const [capturing, setCapturing] = React.useState(false);
  const [submitting, setSubmitting] = React.useState(false);
  const [done, setDone] = React.useState(false);
  const [hits, setHits] = React.useState<ManualArticle[]>([]);

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

  // Object-URLs der Screenshots beim Unmount freigeben.
  React.useEffect(() => () => { shots.forEach((s) => { try { URL.revokeObjectURL(s.url); } catch { /* */ } }); }, [shots]);

  if (!ticketCtx) return null;

  const resetForm = (): void => {
    shots.forEach((s) => { try { URL.revokeObjectURL(s.url); } catch { /* */ } });
    setQuestions(['']); setShots([]); setHits([]); setDone(false);
  };
  const openModal = (): void => { resetForm(); setTab('ask'); setOpen(true); };
  const closeModal = (): void => { if (submitting || capturing) return; setOpen(false); resetForm(); };

  // Routing-Hinweis: wohin geht die Frage?
  const ctxEvent = selectedEventId ? (events || []).find((e) => e.id === selectedEventId) : undefined;
  const askerIsOrgLike = currentUserRole === 'Organizer' || isAdmin;
  const ctxOrgs = ctxEvent
    ? Array.from(new Set([...(ctxEvent.organizerEmails || []), ...(ctxEvent.coOrganizerEmails || [])].map((x) => (x || '').trim()).filter(Boolean)))
    : [];
  const goesToOrganizer = !askerIsOrgLike && !!ctxEvent && ctxOrgs.length > 0;
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
  const removeShot = (idx: number): void => {
    setShots((s) => { const copy = [...s]; const [rm] = copy.splice(idx, 1); if (rm) { try { URL.revokeObjectURL(rm.url); } catch { /* */ } } return copy; });
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
    });
    setSubmitting(false);
    if (ok) {
      shots.forEach((s) => { try { URL.revokeObjectURL(s.url); } catch { /* */ } });
      setShots([]); setQuestions(['']); setHits([]); setDone(true);
    }
  };

  const openArticle = (id: string): void => { setOpen(false); openManualArticle(id, navigate); };

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
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  {hits.map((h) => (
                    <button key={h.id} type="button" onClick={() => openArticle(h.id)}
                      style={{ textAlign: 'left', background: '#fff', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 6, padding: '7px 10px', cursor: 'pointer', fontFamily: 'inherit' }}>
                      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 7 }}>
                        <Icon iconName="ReadingMode" style={{ fontSize: 13, color: 'var(--dex-green,#86bc25)' }} />
                        <strong style={{ fontSize: '0.85rem' }}>{h.title}</strong>
                      </span>
                      {h.description && <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500,#808080)', marginTop: 2 }}>{h.description}</div>}
                    </button>
                  ))}
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
                <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400,#a0a0a0)' }}>{isDe ? 'optional' : 'optional'}</span>
              </div>
              {shots.length > 0 && (
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 10 }}>
                  {shots.map((s, idx) => (
                    <div key={idx} style={{ position: 'relative', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 8, overflow: 'hidden' }}>
                      <img src={s.url} alt={`Screenshot ${idx + 1}`} style={{ width: 120, height: 76, objectFit: 'cover', display: 'block' }} />
                      <button type="button" onClick={() => removeShot(idx)} title={isDe ? 'Entfernen' : 'Remove'}
                        style={{ position: 'absolute', top: 2, right: 2, background: 'rgba(0,0,0,0.55)', color: '#fff', border: 'none', borderRadius: '50%', width: 20, height: 20, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}>
                        <Icon iconName="Cancel" style={{ fontSize: 11 }} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>

            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500,#808080)', display: 'flex', alignItems: 'center', gap: 7 }}>
              <Icon iconName="Send" style={{ fontSize: 13, color: 'var(--dex-green,#86bc25)' }} /> {routingText}
            </div>

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
                {tk.status === 'Closed' && (
                  <div style={{ marginTop: 8, background: '#f1f7e8', borderRadius: 6, padding: '8px 10px' }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: 'var(--dex-green-dark,#4a7c1f)', marginBottom: 3 }}>
                      {isDe ? 'Antwort' : 'Answer'}{tk.answeredByName ? ` · ${tk.answeredByName}` : ''}
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
                      <div style={{ marginTop: 6, fontSize: '0.8rem', color: 'var(--dex-gray-600,#666)' }}>
                        <Icon iconName="DocumentManagement" style={{ fontSize: 12, marginRight: 5 }} />
                        {isDe ? `Im Event-Wizard: Schritt ${tk.answerWizardStep}` : `In the event wizard: step ${tk.answerWizardStep}`}
                      </div>
                    )}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </Modal>
    </>
  );
}
