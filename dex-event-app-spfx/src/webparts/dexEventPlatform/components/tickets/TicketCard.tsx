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
 *
 * v31.3: Nach docs/ui-leitfaden.md umgebaut — Kopf als Personen-Zelle
 * (`dex-ui-person`) mit Status-Pille, der Composer als zwei Abschnitte („Deine
 * Antwort" zuerst, Verweise optional darunter), Knöpfe linksbündig beim Inhalt
 * statt rechts außen. Handler, Bedingungen und ticketThread sind unverändert;
 * neu ist der Hover auf allem Klickbaren und je Eingabe eine Zeile, was danach
 * passiert.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DexTicket } from '../../types';
import { useTickets } from '../../context/TicketContext';
import { useLanguage } from '../../context/LanguageContext';
import { useNavigation } from '../../context/NavigationContext';
import { useCurrentUser } from '../../context/UserContext';
import { useDialog } from '../../context/DialogContext';
import { listManualArticles, openManualArticle, ManualArticle } from '../../utils/manualSearch';
import PersonContactHover from '../PersonContactHover';
import { renderTicketThread, contactSubline } from './ticketThread';
import ImageAnnotateModal from '../ImageAnnotateModal';
import { cx, ensureDexUiStyles } from '../dexUi';
import { ChevronDown } from '../Icons';

// v26.52: Live-Wizard-Vorschau mit Markierungsbox. MUSS lazy bleiben — die
// Modal-Datei importiert EventCreationPage statisch, und die ist im
// Haupt-Router ebenfalls lazy (sonst landet der ganze Wizard im Main-Bundle).
const WizardStepPreviewModal = React.lazy(() => import('./WizardStepPreviewModal'));

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
  const { showAlert } = useDialog();
  const isDe = locale === 'de';
  const myEmailLc = (currentUser.email || '').toLowerCase();
  // Idempotent — die Karte steht auch ohne Modal/WizardFormShell auf der Seite
  // (Tickets-Seite, Event-Übersicht), ohne den Aufruf greifen die Klassen nicht.
  ensureDexUiStyles();

  const [expanded, setExpanded] = React.useState<boolean>(!!props.defaultExpanded);
  const [answerText, setAnswerText] = React.useState('');
  const [selected, setSelected] = React.useState<ManualArticle[]>([]);
  const [wizardStep, setWizardStep] = React.useState<number>(0); // 0 = keiner, sonst 1..10
  // v26.52: Markierungsbox auf der Live-Wizard-Vorschau (Prozent-Koordinaten).
  const [wizardMarker, setWizardMarker] = React.useState<{ x: number; y: number; w: number; h: number } | null>(null);
  const [markerModalOpen, setMarkerModalOpen] = React.useState(false);
  const [answerPreviewOpen, setAnswerPreviewOpen] = React.useState(false);
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
    // v26.32: Beim Übernehmen eines OFFENEN Tickets hart gegen Race absichern —
    // war ein anderer Power-User schneller, kommt ein Konflikt-Hinweis statt
    // stillem Überschreiben. Das bewusste Übernehmen eines bereits „in
    // Bearbeitung" befindlichen Tickets (Takeover) bleibt wie gehabt erlaubt.
    if (ticket.status === 'Open') {
      const res = await claimTicket(ticket.id, { onlyIfOpen: true });
      if (!res.ok) {
        if (res.conflict) {
          const who = res.claimedByName
            ? (isDe ? `„${res.claimedByName}" hat dieses Ticket gerade übernommen.` : `"${res.claimedByName}" just took this ticket.`)
            : (isDe ? 'Dieses Ticket wurde gerade von jemand anderem übernommen.' : 'This ticket was just taken by someone else.');
          showAlert(who, { variant: 'info' });
        }
        return; // NICHT aufklappen — das Ticket gehört jemand anderem.
      }
    } else if (ticket.status === 'InProgress' && !claimedByMe) {
      await claimTicket(ticket.id);
    }
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
      wizardMarker: wizardStep ? wizardMarker : null,
      screenshots: imgFile ? [imgFile] : [],
    });
    setBusy(false);
    if (ok) { setExpanded(false); setAnswerText(''); setSelected([]); setWizardStep(0); setWizardMarker(null); setImgFile(null); }
  };

  const jumpToWizardStep = (step1: number): void => {
    // In den Wizard navigieren und auf den gewählten Schritt springen
    // (gleicher Mechanismus wie das geführte Tutorial: dex-tutorial-wizard-step).
    navigate('create-event');
    window.setTimeout(() => { try { window.dispatchEvent(new CustomEvent('dex-tutorial-wizard-step', { detail: step1 - 1 })); } catch { /* */ } }, 600);
  };

  // v31.3: Statusfarben als Pillen-Modifier statt eigener Inline-Farben — damit
  // sieht der Ticket-Status genauso aus wie jeder andere Status im Organizer
  // Center (orange offen, blau in Arbeit, grün erledigt).
  const statusInfo: Record<string, { pill: string; label: string }> = {
    Open: { pill: 'dex-ui-pill--orange', label: isDe ? 'Offen' : 'Open' },
    InProgress: { pill: 'dex-ui-pill--blue', label: isDe ? 'In Bearbeitung' : 'In progress' },
    Closed: { pill: 'dex-ui-pill--green', label: isDe ? 'Beantwortet' : 'Answered' },
  };
  const si = statusInfo[ticket.status] || statusInfo.Open;
  const created = (() => { try { return new Date(ticket.created).toLocaleString(isDe ? 'de-DE' : 'en-GB'); } catch { return ''; } })();

  const filteredArticles = articleQuery.trim()
    ? articles.filter((a) => `${a.title} ${a.description}`.toLowerCase().indexOf(articleQuery.toLowerCase()) >= 0)
    : articles;

  return (
    <div className="dex-ui-card">
      {/* v31.3: Kopf = wer fragt (Personen-Zelle), in welchem Zustand (Pille),
          seit wann (rechts, weil Zeitstempel keine Aktion ist). */}
      <div className="dex-ui-card-head">
        <span className="dex-ui-person" style={{ flex: '1 1 220px' }}>
          {/* v26.8: Foto-Kontaktkarte des Fragestellers (Hover → Teams-Chat). */}
          <PersonContactHover email={ticket.askerEmail} name={ticket.askerName || ticket.askerEmail} size={32}
            subline={contactSubline(ticket.askerJobTitle, ticket.askerLocation)} isDe={isDe} />
          <span style={{ minWidth: 0 }}>
            <span className="dex-ui-person-name" style={{ display: 'block' }}>{ticket.askerName || ticket.askerEmail}</span>
            {/* v26.36: Sub-Zeile abschneiden statt die Karte über den Handy-
                Viewport hinauszuschieben (role · jobTitle · location · event). */}
            <span className="dex-ui-person-sub" style={{ display: 'block' }}>
              {ticket.askerRole}
              {contactSubline(ticket.askerJobTitle, ticket.askerLocation) ? ` · ${contactSubline(ticket.askerJobTitle, ticket.askerLocation)}` : ''}
              {ticket.eventTitle ? ` · ${ticket.eventTitle}` : ''}
            </span>
          </span>
        </span>
        <span className={cx('dex-ui-pill', si.pill)}>{si.label}</span>
        {/* v26.60: Bug-Reports sichtbar markieren (gehen an die DEX-Maintainer). */}
        {ticket.category === 'bug' && (
          <span className="dex-ui-pill dex-ui-pill--orange">
            <Icon iconName="Bug" style={{ fontSize: 11 }} /> Bug
          </span>
        )}
        {created && <span className="dex-ui-card-head-meta" style={{ marginLeft: 'auto' }}>{created}</span>}
      </div>

      {/* Fragen — der Inhalt, um den es geht. */}
      <div style={{ marginTop: 10 }}>
        {ticket.questions.map((q, i) => (
          <div key={i} style={{ fontSize: '0.92rem', marginBottom: 3, display: 'flex', gap: 7, lineHeight: 1.5 }}>
            <Icon iconName="Help" style={{ fontSize: 13, color: 'var(--dex-green,#86bc25)', marginTop: 3 }} />
            <span>{q}</span>
          </div>
        ))}
      </div>

      {/* v26.30: Frage aus dem Event-Wizard → direkt zum Schritt springen. */}
      {ticket.askWizardStep != null && ticket.askWizardStep >= 1 && (
        <div style={{ marginTop: 6 }}>
          <button type="button" className="dex-ui-textbtn" onClick={() => jumpToWizardStep(ticket.askWizardStep as number)}>
            <Icon iconName="DocumentManagement" style={{ fontSize: 12 }} />
            {isDe ? `Frage aus Event-Wizard · Schritt ${ticket.askWizardStep}: ${stepLabels[ticket.askWizardStep - 1] || ''} öffnen` : `Question from event wizard · step ${ticket.askWizardStep}: ${stepLabels[ticket.askWizardStep - 1] || ''} — open`}
          </button>
        </div>
      )}

      {/* Screenshots des Fragestellers — anklickbar, deshalb mit Hover. */}
      {askShots.length > 0 && (
        <div className="dex-ui-inline" style={{ marginTop: 8 }}>
          {askShots.map((s, i) => (
            <a key={i} href={s.url} target="_blank" rel="noreferrer" className="dex-ui-card dex-ui-card--hover"
              style={{ padding: 0, borderRadius: 8, overflow: 'hidden', display: 'block', lineHeight: 0 }}>
              <img src={s.url} alt={`Screenshot ${i + 1}`} style={{ width: 130, height: 82, objectFit: 'cover', display: 'block' }} />
            </a>
          ))}
        </div>
      )}

      {/* In-Bearbeitung-Hinweis */}
      {ticket.status === 'InProgress' && (
        <div className="dex-ui-callout dex-ui-callout--info" style={{ marginTop: 10 }}>
          <span className="dex-ui-callout-icon"><Icon iconName="Edit" style={{ fontSize: 12 }} /></span>
          <span>
            {claimedByMe
              ? (isDe ? 'Du bearbeitest dieses Ticket gerade.' : 'You are working on this ticket.')
              : (isDe ? `Wird gerade von ${ticket.claimedByName || 'jemandem'} bearbeitet.` : `Currently being handled by ${ticket.claimedByName || 'someone'}.`)}
          </span>
        </div>
      )}

      {/* Bereits beantwortet → Antwort anzeigen (auch wenn durch eine Rückfrage
          wieder geöffnet — die ursprüngliche Antwort bleibt sichtbar). */}
      {!!ticket.answeredAt && (ticket.answerText || ticket.answerArticleIds.length > 0 || ticket.answerWizardStep != null || ansShots.length > 0) && (
        <div className="dex-ui-card dex-ui-card--soft dex-ui-card--accent" style={{ marginTop: 10, padding: '12px 14px' }}>
          <div className="dex-ui-card-head" style={{ marginBottom: 6 }}>
            <span className="dex-ui-pill dex-ui-pill--green">{isDe ? 'Antwort' : 'Answer'}</span>
            {ticket.answeredByEmail && (
              <PersonContactHover email={ticket.answeredByEmail} name={ticket.answeredByName || ticket.answeredByEmail} size={26}
                subline={contactSubline(ticket.answeredByJobTitle, ticket.answeredByLocation)} isDe={isDe} />
            )}
            {ticket.answeredByName && (
              <span className="dex-ui-card-head-meta">
                {ticket.answeredByName}
                {contactSubline(ticket.answeredByJobTitle, ticket.answeredByLocation) ? ` · ${contactSubline(ticket.answeredByJobTitle, ticket.answeredByLocation)}` : ''}
              </span>
            )}
          </div>
          {ticket.answerText && <div style={{ fontSize: '0.88rem', whiteSpace: 'pre-wrap' }}>{ticket.answerText}</div>}
          {ticket.answerArticleIds.length > 0 && (
            <div style={{ marginTop: 6, display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 2 }}>
              {ticket.answerArticleIds.map((aid) => (
                <button key={aid} type="button" className="dex-ui-textbtn" onClick={() => openManualArticle(aid, navigate)}>
                  <Icon iconName="ReadingMode" style={{ fontSize: 12 }} />{isDe ? 'Handbuch-Artikel öffnen' : 'Open manual article'}
                </button>
              ))}
            </div>
          )}
          {ticket.answerWizardStep != null && ticket.answerWizardStep >= 1 && (
            <div className="dex-ui-inline" style={{ marginTop: 6 }}>
              <button type="button" className="dex-ui-textbtn" onClick={() => jumpToWizardStep(ticket.answerWizardStep as number)}>
                <Icon iconName="DocumentManagement" style={{ fontSize: 12 }} />
                {isDe ? `Event-Wizard · Schritt ${ticket.answerWizardStep}: ${stepLabels[ticket.answerWizardStep - 1] || ''} öffnen` : `Open event wizard · step ${ticket.answerWizardStep}: ${stepLabels[ticket.answerWizardStep - 1] || ''}`}
              </button>
              {/* v26.52: Vorschau mit gespeicherter Markierungsbox („hier klicken"). */}
              {ticket.answerWizardMarker && (
                <button type="button" className="dex-ui-textbtn" onClick={() => setAnswerPreviewOpen(true)}>
                  <Icon iconName="Preview" style={{ fontSize: 12 }} />
                  {isDe ? 'Markierung ansehen' : 'View marked spot'}
                </button>
              )}
            </div>
          )}
          {ansShots.length > 0 && (
            <div className="dex-ui-inline" style={{ marginTop: 8 }}>
              {ansShots.map((s, i) => (
                <a key={i} href={s.url} target="_blank" rel="noreferrer" className="dex-ui-card dex-ui-card--hover"
                  style={{ padding: 0, borderRadius: 8, overflow: 'hidden', display: 'block', lineHeight: 0 }}>
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
        <div className="dex-ui-section" style={{ marginTop: 12, borderTop: '1px solid var(--dex-gray-200,#e8e8e8)', paddingTop: 12 }}>
          <label className="dex-ui-label">{isDe ? 'Deine Antwort auf die Rückfrage' : 'Your reply to the follow-up'}</label>
          <textarea className="dex-ui-textarea" value={replyText} onChange={(e) => setReplyText(e.target.value)} rows={3}
            placeholder={isDe ? 'Deine Antwort an den Fragesteller …' : 'Your reply to the asker …'} />
          <div className="dex-ui-help">{isDe ? 'Die Antwort geht per Mail an den Fragesteller; danach ist das Ticket wieder geschlossen.' : 'It is emailed to the asker; after that the ticket is closed again.'}</div>
          {/* v31.3: Beide Knöpfe links beim Feld — vorher trieb `space-between`
              sie an die gegenüberliegenden Ränder, mit leerer Fläche dazwischen. */}
          <div className="dex-ui-inline" style={{ marginTop: 10 }}>
            <button className="btn btn-primary dex-ui-btn-sm" onClick={sendReply} disabled={replyBusy || !replyText.trim()}>
              {replyBusy ? (isDe ? 'Wird gesendet …' : 'Sending …') : (isDe ? 'Antwort senden' : 'Send reply')}
            </button>
            <button className="btn btn-secondary dex-ui-btn-sm" onClick={() => closeTicketNoAnswer(ticket.id)} disabled={replyBusy}>
              {isDe ? 'Keine Antwort nötig' : 'No answer needed'}
            </button>
          </div>
        </div>
      )}

      {/* Aktionen (Erst-Antwort) */}
      {ticket.status !== 'Closed' && !alreadyAnswered && !expanded && (
        <div style={{ marginTop: 12 }}>
          <div className="dex-ui-inline">
            <button className="btn btn-primary dex-ui-btn-sm" onClick={startAnswering}>
              {ticket.status === 'Open' ? (isDe ? 'Übernehmen & beantworten' : 'Take & answer') : (claimedByMe ? (isDe ? 'Antworten' : 'Answer') : (isDe ? 'Übernehmen' : 'Take over'))}
            </button>
            {ticket.status === 'InProgress' && (
              <button className="btn btn-secondary dex-ui-btn-sm" onClick={() => releaseTicket(ticket.id)}>
                {isDe ? 'Wieder freigeben' : 'Release'}
              </button>
            )}
          </div>
          {/* v31.3: Eine Zeile Folge — was das Übernehmen für die anderen heißt. */}
          {ticket.status === 'Open' && (
            <div className="dex-ui-help">
              {isDe
                ? 'Du übernimmst das Ticket — die anderen sehen dann, dass du dran bist.'
                : 'You take the ticket — the others then see that you are on it.'}
            </div>
          )}
        </div>
      )}

      {/* Antwort-Composer (Erst-Antwort) */}
      {ticket.status !== 'Closed' && !alreadyAnswered && expanded && (
        <div style={{ marginTop: 12, borderTop: '1px solid var(--dex-gray-200,#e8e8e8)', paddingTop: 12 }}>
          {/* v31.3: Zuerst der Antworttext, alles Zusätzliche (Artikel,
              Wizard-Schritt, Bild) darunter als „optional" beschriftet. */}
          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{isDe ? 'Deine Antwort' : 'Your answer'}</div>
            <textarea className="dex-ui-textarea" value={answerText} onChange={(e) => setAnswerText(e.target.value)} rows={4}
              placeholder={isDe ? 'Antwort an den Fragesteller …' : 'Answer to the asker …'} />
            <div className="dex-ui-help">{isDe ? 'Die Antwort geht per Mail an den Fragesteller und steht ihm in der App unter „Deine Fragen".' : 'It is emailed to the asker and stays available in the app under "Your questions".'}</div>
          </div>

          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{isDe ? 'Verweise und Anhänge (optional)' : 'References and attachments (optional)'}</div>

            {/* Handbuch-Artikel */}
            <div className="dex-ui-field">
              <button type="button" className={cx('dex-ui-disclosure', showArticlePicker && 'is-open')}
                onClick={() => setShowArticlePicker((v) => !v)} aria-expanded={showArticlePicker}>
                <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                {isDe ? 'Handbuch-Artikel verlinken' : 'Link manual articles'}
                {selected.length > 0 && <span className="dex-ui-disclosure-count">{selected.length}</span>}
              </button>
              {selected.length > 0 && (
                <div className="dex-ui-inline" style={{ marginTop: 8 }}>
                  {selected.map((s) => (
                    // Der ganze Chip entfernt den Verweis (kein Knopf im Knopf).
                    <button key={s.id} type="button" className="dex-ui-chip is-active" onClick={() => toggleArticle(s)}
                      title={isDe ? 'Verweis entfernen' : 'Remove link'}>
                      {s.title}
                      <span className="dex-ui-chip-remove" aria-hidden="true"><Icon iconName="Cancel" style={{ fontSize: 9 }} /></span>
                    </button>
                  ))}
                </div>
              )}
              {showArticlePicker && (
                <div className="dex-ui-card dex-ui-card--soft dex-ui-card--list" style={{ marginTop: 8 }}>
                  <input className="dex-ui-input dex-ui-input--sm" value={articleQuery} onChange={(e) => setArticleQuery(e.target.value)}
                    placeholder={isDe ? 'Artikel suchen …' : 'Search articles …'} style={{ marginBottom: 6 }} />
                  <div style={{ maxHeight: 180, overflowY: 'auto' }}>
                    {filteredArticles.map((a) => {
                      const on = selected.some((s) => s.id === a.id);
                      return (
                        <button key={a.id} type="button" onClick={() => toggleArticle(a)}
                          className={cx('dex-ui-rowbtn', 'dex-ui-row', on && 'is-active')}>
                          <Icon iconName={on ? 'CheckboxComposite' : 'Checkbox'} style={{ fontSize: 15, color: on ? 'var(--dex-green,#86bc25)' : 'var(--dex-gray-400,#a0a0a0)' }} />
                          <span className="dex-ui-row-main" style={{ fontSize: '0.83rem' }}>{a.title}</span>
                        </button>
                      );
                    })}
                    {articles.length === 0 && <span className="dex-ui-muted" style={{ display: 'inline-block', padding: 6 }}>{isDe ? 'Lädt …' : 'Loading …'}</span>}
                    {/* v31.3: Eine Suche ohne Treffer sah aus wie ein leerer Kasten. */}
                    {articles.length > 0 && filteredArticles.length === 0 && (
                      <span className="dex-ui-muted" style={{ display: 'inline-block', padding: 6 }}>
                        {isDe ? 'Kein Artikel passt zu deiner Suche.' : 'No article matches your search.'}
                      </span>
                    )}
                  </div>
                </div>
              )}
            </div>

            {/* Event-Wizard-Schritt */}
            <div className="dex-ui-field">
              <label className="dex-ui-label" htmlFor={`dex-ticket-wizard-step-${ticket.id}`}>
                {isDe ? 'Auf welchen Wizard-Schritt soll die Antwort zeigen?' : 'Which wizard step should the answer point to?'}
              </label>
              <div className="dex-ui-inline">
                <select id={`dex-ticket-wizard-step-${ticket.id}`} className="dex-ui-select dex-ui-select--sm" style={{ width: 'auto', maxWidth: '100%' }}
                  value={wizardStep} onChange={(e) => { setWizardStep(Number(e.target.value)); setWizardMarker(null); }}>
                  <option value={0}>{isDe ? '— keiner —' : '— none —'}</option>
                  {stepLabels.map((lbl, i) => (<option key={i} value={i + 1}>{`${isDe ? 'Schritt' : 'Step'} ${i + 1}: ${lbl}`}</option>))}
                </select>
                {/* v26.52: Statt in den echten Wizard wegzunavigieren, öffnet der
                    Button die Live-Vorschau des Schritts im Modal — dort kann per
                    Drag eine Markierungsbox („hier klicken") gesetzt werden. */}
                {wizardStep > 0 && (
                  <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => setMarkerModalOpen(true)}>
                    <Icon iconName="Preview" style={{ fontSize: 12 }} /> {isDe ? 'Wizard anzeigen & markieren' : 'Show wizard & mark spot'}
                  </button>
                )}
                {wizardStep > 0 && wizardMarker && (
                  <>
                    <span className="dex-ui-pill dex-ui-pill--orange">
                      <Icon iconName="SingleColumnEdit" style={{ fontSize: 11 }} />
                      {isDe ? 'Markierung gesetzt' : 'Marker set'}
                    </span>
                    <button type="button" className="dex-ui-textbtn dex-ui-textbtn--muted" onClick={() => setWizardMarker(null)}>
                      {isDe ? 'Markierung entfernen' : 'Remove marker'}
                    </button>
                  </>
                )}
              </div>
              {wizardStep > 0 && (
                <div className="dex-ui-help">
                  {isDe
                    ? 'Der Fragesteller springt aus der Antwort direkt in diesen Schritt.'
                    : 'From the answer the asker jumps straight into this step.'}
                </div>
              )}
            </div>

            {/* Optionales Bild */}
            <div className="dex-ui-field" style={{ marginBottom: 0 }}>
              <div className="dex-ui-inline">
                <label className="dex-ui-textbtn" style={{ cursor: 'pointer' }}>
                  <Icon iconName="Camera" style={{ fontSize: 14 }} /> {isDe ? 'Bild anhängen' : 'Attach image'}
                  <input type="file" accept="image/*" onChange={onPickImage} style={{ display: 'none' }} />
                </label>
                {imgUrl && (
                  <div style={{ position: 'relative', border: '1px solid var(--dex-gray-200,#e8e8e8)', borderRadius: 8, overflow: 'hidden' }}>
                    <button type="button" className="dex-ui-btn-reset" onClick={() => setAnnotateOpen(true)} title={isDe ? 'Vergrößern & markieren' : 'Enlarge & mark up'}
                      style={{ display: 'block', lineHeight: 0 }}>
                      <img src={imgUrl} alt="Anhang" style={{ width: 110, height: 70, objectFit: 'cover', display: 'block' }} />
                      <span style={{ position: 'absolute', bottom: 2, left: 2, background: 'rgba(0,0,0,0.55)', color: '#fff', borderRadius: 4, padding: '1px 5px', display: 'inline-flex', alignItems: 'center', gap: 3, fontSize: 10, lineHeight: 1.4 }}>
                        <Icon iconName="InsertTextBox" style={{ fontSize: 11 }} /> {isDe ? 'markieren' : 'mark up'}
                      </span>
                    </button>
                    <button type="button" aria-label={isDe ? 'Bild entfernen' : 'Remove image'} title={isDe ? 'Bild entfernen' : 'Remove image'}
                      onClick={() => { if (imgUrl) { try { URL.revokeObjectURL(imgUrl); } catch { /* */ } } setImgFile(null); setImgUrl(''); }}
                      style={{ position: 'absolute', top: 2, right: 2, background: 'rgba(0,0,0,0.55)', color: '#fff', border: 'none', borderRadius: '50%', width: 18, height: 18, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}>
                      <Icon iconName="Cancel" style={{ fontSize: 10 }} />
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* v31.3: Knöpfe links beim Formular statt rechts außen. */}
          <div className="dex-ui-inline" style={{ marginTop: 16 }}>
            <button className="btn btn-primary dex-ui-btn-sm" onClick={submit} disabled={busy}>
              {busy ? (isDe ? 'Wird gesendet …' : 'Sending …') : (isDe ? 'Antwort senden & schließen' : 'Send answer & close')}
            </button>
            <button className="btn btn-secondary dex-ui-btn-sm" onClick={() => setExpanded(false)} disabled={busy}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
          </div>
        </div>
      )}

      {/* v26.52: Live-Wizard-Vorschau — Composer (markieren) bzw. Antwort (ansehen).
          Conditional render + Suspense, damit der Chunk erst beim Öffnen lädt. */}
      {markerModalOpen && wizardStep > 0 && (
        <React.Suspense fallback={null}>
          <WizardStepPreviewModal
            step={wizardStep}
            stepLabel={stepLabels[wizardStep - 1]}
            isDe={isDe}
            editable={true}
            initialMarker={wizardMarker}
            onClose={() => setMarkerModalOpen(false)}
            onSave={(m) => setWizardMarker(m)}
          />
        </React.Suspense>
      )}
      {answerPreviewOpen && ticket.answerWizardStep != null && (
        <React.Suspense fallback={null}>
          <WizardStepPreviewModal
            step={ticket.answerWizardStep}
            stepLabel={stepLabels[ticket.answerWizardStep - 1]}
            isDe={isDe}
            editable={false}
            initialMarker={ticket.answerWizardMarker}
            onClose={() => setAnswerPreviewOpen(false)}
          />
        </React.Suspense>
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
