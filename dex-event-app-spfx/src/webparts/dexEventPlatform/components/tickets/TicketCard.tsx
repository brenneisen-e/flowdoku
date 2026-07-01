/**
 * TicketCard (v26.0.0)
 *
 * Stellt EIN Ticket dar und enthält den Antwort-Composer. Wird sowohl auf der
 * eigenständigen „Tickets"-Seite (Power-User/Admin) als auch in der
 * Event-Übersicht (Organizer, Box „Offene Fragen (User)") verwendet.
 *
 * Antwort = Freitext + ausgewählte Handbuch-Artikel + optional ein
 * Event-Wizard-Schritt (als beschrifteter Verweis/Visual mit Direkt-Sprung in
 * den Wizard) + optionales Bild. Wer ein Ticket öffnet, übernimmt es („in
 * Bearbeitung"); per „Wieder freigeben" kann es ein anderer übernehmen.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DexTicket } from '../../types';
import { useTickets } from '../../context/TicketContext';
import { useLanguage } from '../../context/LanguageContext';
import { useNavigation } from '../../context/NavigationContext';
import { useCurrentUser } from '../../context/UserContext';
import { listManualArticles, openManualArticle, ManualArticle } from '../../utils/manualSearch';
import PersonContactHover from '../PersonContactHover';
import { renderTicketThread, contactSubline } from './ticketThread';
import ImageAnnotateModal from '../ImageAnnotateModal';

function wizardStepLabels(isDe: boolean): string[] {
  // 1-basiert; Reihenfolge identisch zum Event-Wizard (EventCreationPage.steps).
  return isDe
    ? ['Grundlagen', 'Organizer & Team', 'Sub-Events', 'Ort & Programm', 'Kapazität & Sichtbarkeit', 'Felder', 'Kommunikation', 'Team-Anmeldung', 'Dokumente', 'Fun-Zone']
    : ['Basics', 'Organizers & team', 'Sub-events', 'Location & programme', 'Capacity & visibility', 'Fields', 'Communication', 'Team registration', 'Documents', 'Fun zone'];
}

export default function TicketCard(props: { ticket: DexTicket; defaultExpanded?: boolean }): React.ReactElement {
  const { ticket } = props;
  const { claimTicket, releaseTicket, answerTicket, replyToTicket, closeTicketNoAnswer } = useTickets();
  const { locale } = useLanguage();
  const { navigate } = useNavigation();
  const { currentUser } = useCurrentUser();
  const isDe = locale === 'de';
  const myEmailLc = (currentUser.email || '').toLowerCase();

  const [expanded, setExpanded] = React.useState<boolean>(!!props.defaultExpanded);
  const [answerText, setAnswerText] = React.useState('');
  const [selected, setSelected] = React.useState<ManualArticle[]>([]);
  const [wizardStep, setWizardStep] = React.useState<number>(0); // 0 = keiner, sonst 1..10
  const [articles, setArticles] = React.useState<ManualArticle[]>([]);
  const [articleQuery, setArticleQuery] = React.useState('');
  const [showArticlePicker, setShowArticlePicker] = React.useState(false);
  const [imgFile, setImgFile] = React.useState<File | null>(null);
  const [imgUrl, setImgUrl] = React.useState<string>('');
  const [annotateOpen, setAnnotateOpen] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  // v26.8: Folge-Antwort auf eine Rückfrage des Fragestellers.
  const [replyText, setReplyText] = React.useState('');
  const [replyBusy, setReplyBusy] = React.useState(false);

  const stepLabels = wizardStepLabels(isDe);

  React.useEffect(() => {
    if (!showArticlePicker || articles.length > 0) return;
    listManualArticles(isDe ? 'de' : 'en').then(setArticles).catch(() => { /* */ });
  }, [showArticlePicker, isDe, articles.length]);

  React.useEffect(() => () => { if (imgUrl) { try { URL.revokeObjectURL(imgUrl); } catch { /* */ } } }, [imgUrl]);

  const askShots = ticket.attachments.filter((a) => a.kind === 'ask');
  const ansShots = ticket.attachments.filter((a) => a.kind === 'ans');

  const claimedByMe = (ticket.claimedByEmail || '').toLowerCase() === myEmailLc;
  // v26.8: Wurde schon einmal beantwortet? Dann ist ein nicht-geschlossenes
  // Ticket durch eine Rückfrage des Fragestellers WIEDER GEÖFFNET — dann zeigen
  // wir eine schlanke Antwort-Box statt des vollen Erst-Antwort-Composers.
  const alreadyAnswered = !!ticket.answeredAt;
  const reopened = alreadyAnswered && ticket.status !== 'Closed';

  const sendReply = async (): Promise<void> => {
    if (!replyText.trim()) return;
    setReplyBusy(true);
    const ok = await replyToTicket(ticket, replyText.trim());
    setReplyBusy(false);
    if (ok) setReplyText('');
  };

  const startAnswering = async (): Promise<void> => {
    if (ticket.status === 'Open') { await claimTicket(ticket.id); }
    else if (ticket.status === 'InProgress' && !claimedByMe) { await claimTicket(ticket.id); }
    setExpanded(true);
  };

  const onPickImage = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const f = e.target.files && e.target.files[0];
    if (!f) return;
    if (imgUrl) { try { URL.revokeObjectURL(imgUrl); } catch { /* */ } }
    setImgFile(f); setImgUrl(URL.createObjectURL(f));
  };

  const toggleArticle = (a: ManualArticle): void => {
    setSelected((sel) => sel.some((s) => s.id === a.id) ? sel.filter((s) => s.id !== a.id) : [...sel, a]);
  };

  const submit = async (): Promise<void> => {
    if (!answerText.trim() && selected.length === 0 && !wizardStep && !imgFile) return;
    setBusy(true);
    const ok = await answerTicket({
      ticket,
      answerText: answerText.trim(),
      articles: selected.map((s) => ({ id: s.id, title: s.title })),
      wizardStep: wizardStep || null,
      wizardStepLabel: wizardStep ? stepLabels[wizardStep - 1] : undefined,
      screenshots: imgFile ? [imgFile] : [],
    });
    setBusy(false);
    if (ok) { setExpanded(false); setAnswerText(''); setSelected([]); setWizardStep(0); setImgFile(null); }
  };

  const jumpToWizardStep = (step1: number): void => {
    // In den Wizard navigieren und auf den gewählten Schritt springen
    // (gleicher Mechanismus wie das geführte Tutorial: dex-tutorial-wizard-step).
    navigate('create-event');
    window.setTimeout(() => { try { window.dispatchEvent(new CustomEvent('dex-tutorial-wizard-step', { detail: step1 - 1 })); } catch { /* */ } }, 600);
  };

  const statusInfo: Record<string, { bg: string; col: string; label: string }> = {
    Open: { bg: '#fff3e0', col: '#b35a00', label: isDe ? 'Offen' : 'Open' },
    InProgress: { bg: '#e3f2fd', col: '#1565c0', label: isDe ? 'In Bearbeitung' : 'In progress' },
    Closed: { bg: '#e8f5e9', col: '#2e7d32', label: isDe ? 'Beantwortet' : 'Answered' },
  };
  const si = statusInfo[ticket.status] || statusInfo.Open;
  const created = (() => { try { return new Date(ticket.created).toLocaleString(isDe ? 'de-DE' : 'en-GB'); } catch { return ''; } })();

  const filteredArticles = articleQuery.trim()
    ? articles.filter((a) => `${a.title} ${a.description}`.toLowerCase().indexOf(articleQuery.toLowerCase()) >= 0)
    : articles;

  return (
    <div style={{ border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 10, padding: '12px 14px', background: '#fff' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 10, flexWrap: 'wrap' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
          <span style={{ padding: '2px 9px', borderRadius: 10, background: si.bg, color: si.col, fontSize: 12, fontWeight: 700 }}>{si.label}</span>
          {/* v26.8: Foto-Kontaktkarte des Fragestellers (Hover → Teams-Chat). */}
          <PersonContactHover email={ticket.askerEmail} name={ticket.askerName || ticket.askerEmail} size={30}
            subline={contactSubline(ticket.askerJobTitle, ticket.askerLocation)} isDe={isDe} />
          <span style={{ display: 'inline-flex', flexDirection: 'column', lineHeight: 1.25 }}>
            <strong style={{ fontSize: '0.92rem' }}>{ticket.askerName || ticket.askerEmail}</strong>
            <span style={{ fontSize: 12, color: 'var(--dex-gray-400,#a0a0a0)' }}>
              {ticket.askerRole}
              {contactSubline(ticket.askerJobTitle, ticket.askerLocation) ? ` · ${contactSubline(ticket.askerJobTitle, ticket.askerLocation)}` : ''}
              {ticket.eventTitle ? ` · ${ticket.eventTitle}` : ''}
            </span>
          </span>
        </div>
        {created && <span style={{ fontSize: 11, color: 'var(--dex-gray-400,#a0a0a0)' }}>{created}</span>}
      </div>

      {/* Fragen */}
      <div style={{ marginTop: 8 }}>
        {ticket.questions.map((q, i) => (
          <div key={i} style={{ fontSize: '0.92rem', marginBottom: 3, display: 'flex', gap: 7 }}>
            <Icon iconName="Help" style={{ fontSize: 13, color: 'var(--dex-green,#86bc25)', marginTop: 3 }} />
            <span>{q}</span>
          </div>
        ))}
      </div>

      {/* v26.30: Frage aus dem Event-Wizard → direkt zum Schritt springen. */}
      {ticket.askWizardStep != null && ticket.askWizardStep >= 1 && (
        <button type="button" onClick={() => jumpToWizardStep(ticket.askWizardStep as number)}
          style={{ marginTop: 8, textAlign: 'left', background: '#fff', border: '1px solid var(--dex-green,#86bc25)', borderRadius: 6, padding: '6px 9px', cursor: 'pointer', fontFamily: 'inherit', fontSize: '0.82rem', color: 'var(--dex-green-dark,#4a7c1f)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: 5 }}>
          <Icon iconName="DocumentManagement" style={{ fontSize: 12 }} />
          {isDe ? `Frage aus Event-Wizard · Schritt ${ticket.askWizardStep}: ${stepLabels[ticket.askWizardStep - 1] || ''} öffnen` : `Question from event wizard · step ${ticket.askWizardStep}: ${stepLabels[ticket.askWizardStep - 1] || ''} — open`}
        </button>
      )}

      {/* Screenshots des Fragestellers */}
      {askShots.length > 0 && (
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 8 }}>
          {askShots.map((s, i) => (
            <a key={i} href={s.url} target="_blank" rel="noreferrer" style={{ display: 'block', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 8, overflow: 'hidden' }}>
              <img src={s.url} alt={`Screenshot ${i + 1}`} style={{ width: 130, height: 82, objectFit: 'cover', display: 'block' }} />
            </a>
          ))}
        </div>
      )}

      {/* In-Bearbeitung-Hinweis */}
      {ticket.status === 'InProgress' && (
        <div style={{ marginTop: 8, fontSize: '0.8rem', color: 'var(--dex-gray-600,#666)', display: 'flex', alignItems: 'center', gap: 6 }}>
          <Icon iconName="Edit" style={{ fontSize: 12, color: '#1565c0' }} />
          {claimedByMe
            ? (isDe ? 'Du bearbeitest dieses Ticket gerade.' : 'You are working on this ticket.')
            : (isDe ? `Wird gerade von ${ticket.claimedByName || 'jemandem'} bearbeitet.` : `Currently being handled by ${ticket.claimedByName || 'someone'}.`)}
        </div>
      )}

      {/* Bereits beantwortet → Antwort anzeigen (auch wenn durch eine Rückfrage
          wieder geöffnet — die ursprüngliche Antwort bleibt sichtbar). */}
      {!!ticket.answeredAt && (ticket.answerText || ticket.answerArticleIds.length > 0 || ticket.answerWizardStep != null || ansShots.length > 0) && (
        <div style={{ marginTop: 10, background: '#f1f7e8', borderRadius: 8, padding: '9px 11px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
            <span style={{ fontSize: 12, fontWeight: 700, color: 'var(--dex-green-dark,#4a7c1f)' }}>{isDe ? 'Antwort' : 'Answer'}</span>
            {ticket.answeredByEmail && (
              <PersonContactHover email={ticket.answeredByEmail} name={ticket.answeredByName || ticket.answeredByEmail} size={26}
                subline={contactSubline(ticket.answeredByJobTitle, ticket.answeredByLocation)} isDe={isDe} />
            )}
            {ticket.answeredByName && (
              <span style={{ fontSize: 12, color: 'var(--dex-gray-600,#666)' }}>
                {ticket.answeredByName}
                {contactSubline(ticket.answeredByJobTitle, ticket.answeredByLocation) ? ` · ${contactSubline(ticket.answeredByJobTitle, ticket.answeredByLocation)}` : ''}
              </span>
            )}
          </div>
          {ticket.answerText && <div style={{ fontSize: '0.88rem', whiteSpace: 'pre-wrap' }}>{ticket.answerText}</div>}
          {ticket.answerArticleIds.length > 0 && (
            <div style={{ marginTop: 6, display: 'flex', flexDirection: 'column', gap: 3 }}>
              {ticket.answerArticleIds.map((aid) => (
                <button key={aid} type="button" onClick={() => openManualArticle(aid, navigate)}
                  style={{ textAlign: 'left', background: 'transparent', border: 'none', color: 'var(--dex-green-dark,#4a7c1f)', cursor: 'pointer', fontWeight: 600, fontSize: '0.82rem', padding: 0, fontFamily: 'inherit' }}>
                  <Icon iconName="ReadingMode" style={{ fontSize: 12, marginRight: 5 }} />{isDe ? 'Handbuch-Artikel öffnen' : 'Open manual article'}
                </button>
              ))}
            </div>
          )}
          {ticket.answerWizardStep != null && ticket.answerWizardStep >= 1 && (
            <button type="button" onClick={() => jumpToWizardStep(ticket.answerWizardStep as number)}
              style={{ marginTop: 6, textAlign: 'left', background: '#fff', border: '1px solid var(--dex-green,#86bc25)', borderRadius: 6, padding: '6px 9px', cursor: 'pointer', fontFamily: 'inherit', fontSize: '0.82rem', color: 'var(--dex-green-dark,#4a7c1f)', fontWeight: 600 }}>
              <Icon iconName="DocumentManagement" style={{ fontSize: 12, marginRight: 5 }} />
              {isDe ? `Event-Wizard · Schritt ${ticket.answerWizardStep}: ${stepLabels[ticket.answerWizardStep - 1] || ''} öffnen` : `Open event wizard · step ${ticket.answerWizardStep}: ${stepLabels[ticket.answerWizardStep - 1] || ''}`}
            </button>
          )}
          {ansShots.length > 0 && (
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 8 }}>
              {ansShots.map((s, i) => (
                <a key={i} href={s.url} target="_blank" rel="noreferrer" style={{ display: 'block', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 8, overflow: 'hidden' }}>
                  <img src={s.url} alt={`Antwort-Bild ${i + 1}`} style={{ width: 130, height: 82, objectFit: 'cover', display: 'block' }} />
                </a>
              ))}
            </div>
          )}
        </div>
      )}

      {/* v26.8: Rückfragen-Verlauf */}
      {renderTicketThread(ticket, isDe)}

      {/* v26.8: Wieder geöffnet durch eine Rückfrage → schlanke Folge-Antwort
          ODER „keine Antwort nötig" (z.B. wenn die Rückfrage nur ein Danke war). */}
      {reopened && (
        <div style={{ marginTop: 12, borderTop: '1px solid var(--dex-gray-200,#e8e8e8)', paddingTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700,#444)' }}>{isDe ? 'Auf die Rückfrage antworten' : 'Reply to the follow-up'}</label>
          <textarea value={replyText} onChange={(e) => setReplyText(e.target.value)} rows={3}
            placeholder={isDe ? 'Deine Antwort an den Fragesteller …' : 'Your reply to the asker …'}
            style={{ padding: '9px 11px', borderRadius: 8, border: '1px solid var(--dex-gray-300,#d1d1d1)', fontFamily: 'inherit', fontSize: '0.9rem', resize: 'vertical' }} />
          <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8, flexWrap: 'wrap' }}>
            <button className="btn btn-secondary" style={{ padding: '7px 14px' }} onClick={() => closeTicketNoAnswer(ticket.id)} disabled={replyBusy}>
              {isDe ? 'Keine Antwort nötig' : 'No answer needed'}
            </button>
            <button className="btn btn-primary" style={{ padding: '7px 16px' }} onClick={sendReply} disabled={replyBusy || !replyText.trim()}>
              {replyBusy ? (isDe ? 'Wird gesendet …' : 'Sending …') : (isDe ? 'Antwort senden' : 'Send reply')}
            </button>
          </div>
        </div>
      )}

      {/* Aktionen (Erst-Antwort) */}
      {ticket.status !== 'Closed' && !alreadyAnswered && !expanded && (
        <div style={{ display: 'flex', gap: 8, marginTop: 10, flexWrap: 'wrap' }}>
          <button className="btn btn-primary" style={{ padding: '7px 16px' }} onClick={startAnswering}>
            {ticket.status === 'Open' ? (isDe ? 'Übernehmen & beantworten' : 'Take & answer') : (claimedByMe ? (isDe ? 'Antworten' : 'Answer') : (isDe ? 'Übernehmen' : 'Take over'))}
          </button>
          {ticket.status === 'InProgress' && (
            <button className="btn btn-secondary" style={{ padding: '7px 16px' }} onClick={() => releaseTicket(ticket.id)}>
              {isDe ? 'Wieder freigeben' : 'Release'}
            </button>
          )}
        </div>
      )}

      {/* Antwort-Composer (Erst-Antwort) */}
      {ticket.status !== 'Closed' && !alreadyAnswered && expanded && (
        <div style={{ marginTop: 12, borderTop: '1px solid var(--dex-gray-200,#e8e8e8)', paddingTop: 12, display: 'flex', flexDirection: 'column', gap: 10 }}>
          <label style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700,#444)' }}>{isDe ? 'Deine Antwort' : 'Your answer'}</label>
          <textarea value={answerText} onChange={(e) => setAnswerText(e.target.value)} rows={4}
            placeholder={isDe ? 'Antwort an den Fragesteller …' : 'Answer to the asker …'}
            style={{ padding: '9px 11px', borderRadius: 8, border: '1px solid var(--dex-gray-300,#d1d1d1)', fontFamily: 'inherit', fontSize: '0.9rem', resize: 'vertical' }} />

          {/* Handbuch-Artikel */}
          <div>
            <button type="button" onClick={() => setShowArticlePicker((v) => !v)}
              style={{ display: 'inline-flex', alignItems: 'center', gap: 7, background: 'transparent', border: '1px solid var(--dex-gray-300,#d1d1d1)', borderRadius: 8, padding: '6px 11px', cursor: 'pointer', fontWeight: 600, fontSize: '0.82rem', fontFamily: 'inherit', color: 'var(--dex-gray-700,#444)' }}>
              <Icon iconName="ReadingMode" style={{ fontSize: 13 }} /> {isDe ? 'Handbuch-Artikel verlinken' : 'Link manual articles'}{selected.length > 0 ? ` (${selected.length})` : ''}
            </button>
            {selected.length > 0 && (
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 8 }}>
                {selected.map((s) => (
                  <span key={s.id} style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: '#f1f7e8', border: '1px solid var(--dex-green,#86bc25)', borderRadius: 14, padding: '3px 10px', fontSize: '0.78rem', color: 'var(--dex-green-dark,#4a7c1f)' }}>
                    {s.title}
                    <button type="button" onClick={() => toggleArticle(s)} style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'var(--dex-green-dark,#4a7c1f)', padding: 0, display: 'inline-flex' }}><Icon iconName="Cancel" style={{ fontSize: 10 }} /></button>
                  </span>
                ))}
              </div>
            )}
            {showArticlePicker && (
              <div style={{ marginTop: 8, border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 8, padding: 8 }}>
                <input value={articleQuery} onChange={(e) => setArticleQuery(e.target.value)} placeholder={isDe ? 'Artikel suchen …' : 'Search articles …'}
                  style={{ width: '100%', padding: '7px 10px', borderRadius: 6, border: '1px solid var(--dex-gray-300,#d1d1d1)', fontFamily: 'inherit', fontSize: '0.85rem', marginBottom: 8 }} />
                <div style={{ maxHeight: 180, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 2 }}>
                  {filteredArticles.map((a) => {
                    const on = selected.some((s) => s.id === a.id);
                    return (
                      <button key={a.id} type="button" onClick={() => toggleArticle(a)}
                        style={{ display: 'flex', alignItems: 'center', gap: 8, textAlign: 'left', background: on ? '#f1f7e8' : 'transparent', border: 'none', borderRadius: 6, padding: '6px 8px', cursor: 'pointer', fontFamily: 'inherit' }}>
                        <Icon iconName={on ? 'CheckboxComposite' : 'Checkbox'} style={{ fontSize: 15, color: on ? 'var(--dex-green,#86bc25)' : 'var(--dex-gray-400,#a0a0a0)' }} />
                        <span style={{ fontSize: '0.83rem' }}>{a.title}</span>
                      </button>
                    );
                  })}
                  {articles.length === 0 && <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400,#a0a0a0)', padding: 6 }}>{isDe ? 'Lädt …' : 'Loading …'}</span>}
                </div>
              </div>
            )}
          </div>

          {/* Event-Wizard-Schritt */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
            <label style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700,#444)' }}>{isDe ? 'Event-Wizard-Schritt einbinden:' : 'Embed event-wizard step:'}</label>
            <select value={wizardStep} onChange={(e) => setWizardStep(Number(e.target.value))}
              style={{ padding: '6px 10px', borderRadius: 6, border: '1px solid var(--dex-gray-300,#d1d1d1)', fontFamily: 'inherit', fontSize: '0.85rem' }}>
              <option value={0}>{isDe ? '— keiner —' : '— none —'}</option>
              {stepLabels.map((lbl, i) => (<option key={i} value={i + 1}>{`${isDe ? 'Schritt' : 'Step'} ${i + 1}: ${lbl}`}</option>))}
            </select>
            {wizardStep > 0 && (
              <button type="button" onClick={() => jumpToWizardStep(wizardStep)}
                style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: 'transparent', border: '1px solid var(--dex-green,#86bc25)', color: 'var(--dex-green-dark,#4a7c1f)', borderRadius: 6, padding: '5px 10px', cursor: 'pointer', fontSize: '0.78rem', fontWeight: 600, fontFamily: 'inherit' }}>
                <Icon iconName="OpenInNewWindow" style={{ fontSize: 12 }} /> {isDe ? 'Schritt ansehen' : 'View step'}
              </button>
            )}
          </div>

          {/* Optionales Bild */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
            <label style={{ display: 'inline-flex', alignItems: 'center', gap: 7, background: 'var(--dex-gray-100,#f5f5f5)', border: '1px solid var(--dex-gray-300,#d1d1d1)', borderRadius: 8, padding: '6px 11px', cursor: 'pointer', fontWeight: 600, fontSize: '0.82rem', color: 'var(--dex-gray-700,#444)' }}>
              <Icon iconName="Camera" style={{ fontSize: 14 }} /> {isDe ? 'Bild anhängen' : 'Attach image'}
              <input type="file" accept="image/*" onChange={onPickImage} style={{ display: 'none' }} />
            </label>
            {imgUrl && (
              <div style={{ position: 'relative', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 8, overflow: 'hidden' }}>
                <button type="button" onClick={() => setAnnotateOpen(true)} title={isDe ? 'Vergrößern & markieren' : 'Enlarge & mark up'}
                  style={{ display: 'block', padding: 0, border: 'none', background: 'transparent', cursor: 'pointer' }}>
                  <img src={imgUrl} alt="Anhang" style={{ width: 110, height: 70, objectFit: 'cover', display: 'block' }} />
                  <span style={{ position: 'absolute', bottom: 2, left: 2, background: 'rgba(0,0,0,0.55)', color: '#fff', borderRadius: 4, padding: '1px 5px', display: 'inline-flex', alignItems: 'center', gap: 3, fontSize: 10 }}>
                    <Icon iconName="InsertTextBox" style={{ fontSize: 11 }} /> {isDe ? 'markieren' : 'mark up'}
                  </span>
                </button>
                <button type="button" onClick={() => { if (imgUrl) { try { URL.revokeObjectURL(imgUrl); } catch { /* */ } } setImgFile(null); setImgUrl(''); }}
                  style={{ position: 'absolute', top: 2, right: 2, background: 'rgba(0,0,0,0.55)', color: '#fff', border: 'none', borderRadius: '50%', width: 18, height: 18, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}>
                  <Icon iconName="Cancel" style={{ fontSize: 10 }} />
                </button>
              </div>
            )}
          </div>

          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
            <button className="btn btn-secondary" style={{ padding: '7px 16px' }} onClick={() => setExpanded(false)} disabled={busy}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
            <button className="btn btn-primary" style={{ padding: '7px 16px' }} onClick={submit} disabled={busy}>
              {busy ? (isDe ? 'Wird gesendet …' : 'Sending …') : (isDe ? 'Antwort senden & schließen' : 'Send answer & close')}
            </button>
          </div>
        </div>
      )}

      {/* v26.10: Antwort-Bild vergrößern + transparente Boxen darübersetzen. */}
      <ImageAnnotateModal
        open={annotateOpen && !!imgUrl}
        src={imgUrl}
        fileName={imgFile?.name || 'bild.png'}
        isDe={isDe}
        onClose={() => setAnnotateOpen(false)}
        onSave={(f) => {
          if (imgUrl) { try { URL.revokeObjectURL(imgUrl); } catch { /* */ } }
          setImgFile(f); setImgUrl(URL.createObjectURL(f)); setAnnotateOpen(false);
        }}
      />
    </div>
  );
}
