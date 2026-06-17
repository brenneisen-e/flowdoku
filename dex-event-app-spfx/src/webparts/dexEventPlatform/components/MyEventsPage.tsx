/**
 * Meine Events - zeigt alle Events für die der User registriert ist.
 * Lädt Registrierungen aus den jeweiligen Teilnehmerlisten.
 * Ermöglicht Abmeldung mit Zwei-Schritt-Bestätigung.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import OrganizerList from './OrganizerList';
import { CachedImg } from './CachedImage';

import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useRoles } from '../context/RoleContext';
import { UserFieldPicker } from './UserFieldPicker';
import { useCurrentUser } from '../context/UserContext';
// v22.13: Sub-Events in „Meine Events" nach ihrer EIGENEN Sichtbarkeit filtern
// (gleiche Logik wie Anmeldeseite/Event-Liste).
import { isEventVisibleForUser } from './EventListPage';
import { DeloitteEvent, EventSpecificField, AgendaItem, TransferTime, QuizQuestion } from '../types';
import { SPRegistration } from '../services/EventService';
import { wrapTemplate } from '../services/EmailTemplates';
import { isEventOver } from '../utils/eventFormat';
import { useLanguage } from '../context/LanguageContext';
// v20.4: moderne Confirm-/Alert-Modals statt window.confirm/alert.
import { useDialog } from '../context/DialogContext';
// v11.99: RefreshCw nicht mehr benötigt (Page-Level-Refresh-Button entfernt).
import { X, Pencil, QrCode } from './Icons';
import { InfoTooltip } from './InfoTooltip';
import Modal from './Modal';
import InternationalSearchToggle from './InternationalSearchToggle';
import { buildDemoShowcaseEvents, buildDemoMyRegistration } from '../services/demoShowcaseEvent';

// v20.0 (Audit): PdfViewer zieht react-pdf (+pdfjs) ins Bundle — lazy laden,
// der Viewer wird nur beim Öffnen eines Dokuments gebraucht.
const PdfViewer = React.lazy(() => import('./PdfViewer'));

// v19.34: People-Picker-Antworten (Feldtyp `user`/`roommate`) im „Meine
// Events"-Antwort-Tag mit Profilfoto statt als Rohtext „Name <email>"
// anzeigen — analog zum Chip im People-Picker selbst.
const parsePersonAnswer = (v: string): { name: string; email: string } | null => {
  const m = (v || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
  if (!m) return null;
  return { name: m[1].trim(), email: m[2].trim() };
};

function FieldAnswerTag(props: { label: string; value: string; type?: string; small?: boolean }): React.ReactElement {
  const { label, value, type, small } = props;
  const person = (type === 'user' || type === 'roommate') ? parsePersonAnswer(value) : null;
  const baseStyle: React.CSSProperties = {
    fontSize: small ? '0.72rem' : '0.78rem',
    padding: small ? '3px 8px' : '4px 10px',
    borderRadius: 4,
    background: 'rgba(134,188,37,0.14)',
    color: 'var(--dex-green-dark, #4a7c1f)',
    border: '1px solid rgba(134,188,37,0.30)',
  };
  if (person) {
    return (
      <span style={{ ...baseStyle, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        {label}:
        <img
          src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(person.email)}&size=L`}
          alt={person.name}
          onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
          style={{ width: 22, height: 22, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', transition: 'transform 0.15s', transformOrigin: 'center' }}
          onMouseEnter={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.6)'; (e.currentTarget as HTMLImageElement).style.zIndex = '20'; (e.currentTarget as HTMLImageElement).style.position = 'relative'; (e.currentTarget as HTMLImageElement).style.boxShadow = '0 6px 18px rgba(0,0,0,0.18)'; }}
          onMouseLeave={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)'; (e.currentTarget as HTMLImageElement).style.boxShadow = 'none'; }}
        />
        <strong>{person.name}</strong>
      </span>
    );
  }
  return (
    <span style={baseStyle}>
      {label}: <strong>{value}</strong>
    </span>
  );
}

interface MyEventEntry {
  event: DeloitteEvent;
  registration: SPRegistration;
  /** Seit v6.14: wenn true, ist der User NICHT direkt fürs Parent-Event angemeldet,
   *  sondern nur für mindestens eine Sub-Event-Session. Die Parent-Karte dient dann
   *  nur als Container, damit die Session-Registrierungen sichtbar und verwaltbar
   *  bleiben. Das Status-Badge zeigt "Nur Sessions" statt dem echten Parent-Status. */
  sessionsOnly?: boolean;
  /** v11.31: Titel der Sub-Events für die der User aktiv angemeldet ist —
   *  wird bei sessionsOnly-Entries befüllt, damit der Hinweis-Text die
   *  konkreten Sub-Event-Namen in Klammern ausgeben kann. */
  subEventTitles?: string[];
}

function formatDate(iso: string): string {
  if (!iso) return '-';
  return new Date(iso).toLocaleDateString('de-DE', {
    day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
  });
}

function formatDateRange(start: string, end: string): string {
  if (!start) return '-';
  const s = new Date(start);
  const e = end ? new Date(end) : null;
  const sDate = s.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const sTime = s.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  if (!e) return `${sDate}, ${sTime}`;
  const eDate = e.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const eTime = e.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  // Gleicher Tag: "14.04.2026, 14:00 – 18:00"
  if (sDate === eDate) return `${sDate}, ${sTime} – ${eTime}`;
  // Verschiedene Tage
  return `${sDate}, ${sTime} – ${eDate}, ${eTime}`;
}

function getStatusBadgeClass(status: string): string {
  switch (status) {
    case 'Angemeldet': return 'badge-green';
    case 'QR versendet': return 'badge-green';
    case 'Warteliste': return 'badge-orange';
    case 'Abgemeldet': return 'badge-red';
    case 'Eingecheckt': return 'badge-green';
    default: return 'badge-gray';
  }
}

function getStatusLabel(status: string, t: (key: string) => string): string {
  switch (status) {
    case 'Angemeldet': return t('status.registered');
    case 'Warteliste': return t('status.waitlist');
    case 'Abgemeldet': return t('status.cancelled');
    case 'Eingecheckt': return t('status.checkedin');
    default: return status;
  }
}

function getDocIconName(name: string): string {
  const ext = name.split('.').pop()?.toLowerCase() || '';
  switch (ext) {
    case 'pdf': return 'PDF';
    case 'doc': case 'docx': return 'WordDocument';
    case 'xls': case 'xlsx': return 'ExcelDocument';
    case 'ppt': case 'pptx': return 'PowerPointDocument';
    case 'jpg': case 'jpeg': case 'png': case 'gif': return 'FileImage';
    default: return 'Page';
  }
}

/**
 * Berechnet wie viele Fragen korrekt beantwortet wurden.
 * Eine Frage gilt als korrekt, wenn die Menge der gewählten Indices exakt
 * gleich der Menge der `correctIndices` ist.
 */
function computeQuizScore(quiz: QuizQuestion[], answers: number[][]): number {
  let score = 0;
  for (let i = 0; i < quiz.length; i++) {
    const q = quiz[i];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const correct = q.correctIndices || [(q as any).correctIndex || 0];
    const given = answers[i] || [];
    if (given.length === 0) continue;
    if (given.length === correct.length && correct.every(c => given.includes(c))) {
      score++;
    }
  }
  return score;
}

/**
 * Zählt wie viele Fragen (mindestens teilweise) beantwortet wurden.
 */
function countAnswered(answers: number[][]): number {
  let n = 0;
  for (const a of answers) {
    if (a && a.length > 0) n++;
  }
  return n;
}

/**
 * QuizPlayer mit Resume + Cluster + Progress-Bar + Auto-Save.
 *
 * - initialAnswers: zuvor gespeicherte Antworten (leere innere Arrays = nicht beantwortet)
 * - clusterSize: 1..4 Fragen pro Ansicht (Default 1)
 * - onProgress: wird bei jedem "Weiter"/"Zurück" aufgerufen + am Ende mit isComplete=true
 */
function QuizPlayer({
  quiz,
  t,
  clusterSize = 1,
  initialAnswers,
  onProgress,
}: {
  quiz: QuizQuestion[];
  t: (key: string) => string;
  clusterSize?: number;
  initialAnswers?: number[][];
  onProgress?: (score: number, answers: number[][], isComplete: boolean) => void;
}): React.ReactElement {
  const size = Math.min(Math.max(1, clusterSize || 1), 4);
  const isDe = t('myevents.agenda') === 'Programm';

  // Antworten initialisieren: für jede Frage ein Array (leer wenn unbeantwortet).
  // Pads initialAnswers auf quiz.length, damit Zugriff per Index immer sicher ist.
  const buildInitial = (): number[][] => {
    const padded: number[][] = [];
    for (let i = 0; i < quiz.length; i++) {
      const prev = initialAnswers && Array.isArray(initialAnswers[i]) ? initialAnswers[i] : [];
      padded.push(Array.isArray(prev) ? prev.slice() : []);
    }
    return padded;
  };
  const [allAnswers, setAllAnswers] = React.useState<number[][]>(buildInitial);

  const hadResumeData = !!initialAnswers && countAnswered(buildInitial()) > 0;
  const completedAllInitial = hadResumeData && countAnswered(buildInitial()) === quiz.length;

  // Gruppen = "Seiten" des Quiz. Wenn mindestens eine Frage ein Feld `section`
  // hat, wird pro Section eine Gruppe (alle Fragen der Section zusammen). Fragen
  // ohne Section landen im Anschluss in einer "Ohne Bereich"-Gruppe. Ohne
  // Sections fällt alles auf Cluster-Größe zurück.
  const groups: Array<{ title?: string; indices: number[] }> = (() => {
    const hasAnySection = quiz.some(q => !!q.section);
    if (!hasAnySection) {
      // Klassisches Cluster-Verhalten
      const out: Array<{ title?: string; indices: number[] }> = [];
      for (let i = 0; i < quiz.length; i += size) {
        out.push({ indices: Array.from({ length: Math.min(size, quiz.length - i) }, (_, k) => i + k) });
      }
      return out;
    }
    // Sections in Reihenfolge der ersten Frage-Erwähnung
    const sectionsInOrder: string[] = [];
    for (const q of quiz) {
      if (q.section && sectionsInOrder.indexOf(q.section) < 0) sectionsInOrder.push(q.section);
    }
    const out: Array<{ title?: string; indices: number[] }> = [];
    for (const sec of sectionsInOrder) {
      const indices: number[] = [];
      for (let i = 0; i < quiz.length; i++) {
        if (quiz[i].section === sec) indices.push(i);
      }
      out.push({ title: sec, indices });
    }
    // Unsortierte am Ende als eigene Gruppe (nur wenn vorhanden)
    const unsorted: number[] = [];
    for (let i = 0; i < quiz.length; i++) {
      if (!quiz[i].section) unsorted.push(i);
    }
    if (unsorted.length > 0) out.push({ title: isDe ? 'Ohne Bereich' : 'No section', indices: unsorted });
    return out;
  })();

  // Welche Gruppe enthält die erste unbeantwortete Frage?
  const firstUnansweredGroupIdx = (answers: number[][]): number => {
    for (let g = 0; g < groups.length; g++) {
      for (const qi of groups[g].indices) {
        if (!answers[qi] || answers[qi].length === 0) return g;
      }
    }
    return groups.length; // alles beantwortet
  };

  const initialGroupIdx = Math.min(firstUnansweredGroupIdx(buildInitial()), Math.max(0, groups.length - 1));

  const [currentGroupIdx, setCurrentGroupIdx] = React.useState(initialGroupIdx);
  const [showQuiz, setShowQuiz] = React.useState(false);
  const [showSummary, setShowSummary] = React.useState(completedAllInitial);
  const [isSaving, setIsSaving] = React.useState(false);

  // Prop-Sync wenn initialAnswers nachkommt (z.B. nach Save-Reload)
  const initialAnswersKey = initialAnswers ? JSON.stringify(initialAnswers) : '';
  React.useEffect(() => {
    const fresh = buildInitial();
    setAllAnswers(fresh);
    if (!showQuiz) {
      setCurrentGroupIdx(Math.min(firstUnansweredGroupIdx(fresh), Math.max(0, groups.length - 1)));
      setShowSummary(countAnswered(fresh) === quiz.length && quiz.length > 0);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [initialAnswersKey, quiz.length]);

  const toggleOption = (qIdx: number, optIdx: number): void => {
    setAllAnswers(prev => {
      const next = prev.slice();
      const curr = next[qIdx] ? next[qIdx].slice() : [];
      const pos = curr.indexOf(optIdx);
      if (pos >= 0) curr.splice(pos, 1); else curr.push(optIdx);
      next[qIdx] = curr;
      return next;
    });
  };

  const answeredCount = countAnswered(allAnswers);
  const progressPct = quiz.length > 0 ? Math.round((answeredCount / quiz.length) * 100) : 0;

  const saveProgress = async (answers: number[][], markComplete: boolean): Promise<void> => {
    if (!onProgress) return;
    setIsSaving(true);
    try {
      await onProgress(computeQuizScore(quiz, answers), answers, markComplete);
    } catch { /* Save-Fehler ignorieren - State bleibt lokal */ }
    setIsSaving(false);
  };

  const goNext = async (): Promise<void> => {
    const nextIdx = currentGroupIdx + 1;
    if (nextIdx >= groups.length) {
      const done = countAnswered(allAnswers) === quiz.length;
      await saveProgress(allAnswers, done);
      setShowSummary(true);
    } else {
      await saveProgress(allAnswers, false);
      setCurrentGroupIdx(nextIdx);
    }
  };

  const goBack = async (): Promise<void> => {
    if (showSummary) {
      setShowSummary(false);
      return;
    }
    if (currentGroupIdx === 0) return;
    await saveProgress(allAnswers, false);
    setCurrentGroupIdx(Math.max(0, currentGroupIdx - 1));
  };

  if (!showQuiz) {
    const resumeLabel = completedAllInitial
      ? (isDe ? 'Antworten ansehen / ändern' : 'Review / edit answers')
      : hadResumeData
        ? (isDe ? `Quiz fortsetzen (${answeredCount}/${quiz.length})` : `Resume quiz (${answeredCount}/${quiz.length})`)
        : (isDe ? `Quiz starten (${quiz.length} ${quiz.length === 1 ? 'Frage' : 'Fragen'})` : `Start quiz (${quiz.length} ${quiz.length === 1 ? 'question' : 'questions'})`);
    return (
      <div style={{ marginTop: 12 }}>
        <button
          onClick={() => setShowQuiz(true)}
          style={{
            width: '100%', padding: '14px 20px', borderRadius: 12,
            border: '2px solid var(--dex-green)', background: 'rgba(134,188,37,0.06)',
            cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10,
            fontSize: '0.92rem', fontWeight: 600, color: 'var(--dex-green-dark)',
          }}
        >
          <Icon iconName="Game" style={{ fontSize: 20 }} />
          {resumeLabel}
        </button>
      </div>
    );
  }

  // ===== Summary =====
  if (showSummary) {
    const score = computeQuizScore(quiz, allAnswers);
    return (
      <div style={{ marginTop: 12, padding: 20, background: 'var(--dex-gray-50)', borderRadius: 12 }}>
        <div style={{ textAlign: 'center', marginBottom: 16 }}>
          <div style={{ fontSize: '2rem', marginBottom: 8 }}>{score === quiz.length ? '🎉' : score >= quiz.length / 2 ? '👏' : '💪'}</div>
          <div style={{ fontSize: '1.1rem', fontWeight: 700, marginBottom: 4 }}>
            {score} / {quiz.length} {isDe ? 'richtig' : 'correct'}
          </div>
          <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
            {score === quiz.length
              ? (isDe ? 'Perfekt! Alle richtig!' : 'Perfect! All correct!')
              : (isDe ? 'Gut gemacht!' : 'Well done!')}
          </div>
        </div>
        {/* Einzel-Auflistung: Fragen + eigene Antwort + Haken */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 16 }}>
          {quiz.map((q, qi) => {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const correct = q.correctIndices || [(q as any).correctIndex || 0];
            const given = allAnswers[qi] || [];
            const isCorrect = given.length > 0 && given.length === correct.length && correct.every(c => given.includes(c));
            const unanswered = given.length === 0;
            return (
              <div key={q.id || qi} style={{
                padding: 10, borderRadius: 8,
                background: unanswered ? 'rgba(237,139,0,0.08)' : isCorrect ? 'rgba(134,188,37,0.1)' : 'rgba(218,41,28,0.08)',
                border: `1px solid ${unanswered ? 'var(--dex-orange)' : isCorrect ? 'var(--dex-green)' : 'var(--dex-red)'}`,
                fontSize: '0.82rem',
              }}>
                <div style={{ fontWeight: 600, marginBottom: 4 }}>
                  {qi + 1}. {q.question}{' '}
                  <Icon
                    iconName={unanswered ? 'Warning' : isCorrect ? 'CheckMark' : 'Cancel'}
                    style={{ fontSize: 14, color: unanswered ? 'var(--dex-orange)' : isCorrect ? 'var(--dex-green)' : 'var(--dex-red)' }}
                  />
                </div>
                {q.imageBase64 && (
                  <img
                    src={q.imageBase64}
                    alt=""
                    loading="lazy"
                    decoding="async"
                    style={{
                      maxWidth: '100%', maxHeight: 120, display: 'block',
                      borderRadius: 6, marginBottom: 6, border: '1px solid var(--dex-gray-200)',
                    }}
                  />
                )}
                <div style={{ color: 'var(--dex-gray-600)', fontSize: '0.78rem' }}>
                  {unanswered
                    ? (isDe ? 'Nicht beantwortet' : 'Not answered')
                    : `${isDe ? 'Deine Antwort' : 'Your answer'}: ${given.map(i => q.options[i]).join(', ')}`}
                </div>
              </div>
            );
          })}
        </div>
        <div style={{ display: 'flex', gap: 8, justifyContent: 'center' }}>
          <button className="btn btn-secondary" onClick={goBack} style={{ fontSize: '0.82rem' }}>
            {isDe ? 'Antworten ändern' : 'Edit answers'}
          </button>
          <button className="btn btn-secondary" onClick={() => setShowQuiz(false)} style={{ fontSize: '0.82rem' }}>
            {t('create.templates.close')}
          </button>
        </div>
      </div>
    );
  }

  // ===== Aktive Gruppe (Section oder Cluster) =====
  const currentGroup = groups[currentGroupIdx] || { indices: [] };
  const isLastGroup = currentGroupIdx >= groups.length - 1;

  return (
    <div style={{ marginTop: 12, padding: 16, background: 'var(--dex-gray-50)', borderRadius: 12 }}>
      {/* Progress-Bar */}
      <div style={{ marginBottom: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
          <span>
            {isDe ? 'Fortschritt' : 'Progress'}: {answeredCount} / {quiz.length} ({progressPct}%)
          </span>
          <span style={{ color: 'var(--dex-gray-400)' }}>
            {isDe ? 'Seite' : 'Page'} {currentGroupIdx + 1} / {groups.length}
          </span>
        </div>
        <div style={{ width: '100%', height: 8, background: 'var(--dex-gray-200)', borderRadius: 4, overflow: 'hidden' }}>
          <div style={{
            width: `${progressPct}%`, height: '100%',
            background: progressPct === 100 ? 'var(--dex-green)' : 'var(--dex-green-dark, #4a7c1f)',
            transition: 'width 0.3s ease',
          }} />
        </div>
      </div>

      {/* Bereichs-Titel, wenn vorhanden */}
      {currentGroup.title && (
        <h3 style={{ margin: '0 0 14px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1.15rem' }}>
          {currentGroup.title}
        </h3>
      )}

      {/* Fragen der aktuellen Gruppe */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        {currentGroup.indices.map(qIdx => {
          const question = quiz[qIdx];
          if (!question) return null;
          const given = allAnswers[qIdx] || [];
          return (
            <div key={question.id || qIdx}>
              <div style={{ fontSize: '1rem', fontWeight: 600, marginBottom: 10 }}>
                {qIdx + 1}. {question.question}
              </div>
              {question.imageBase64 && (
                <img
                  src={question.imageBase64}
                  alt=""
                  loading="lazy"
                  decoding="async"
                  style={{
                    maxWidth: '100%', maxHeight: 240, display: 'block',
                    borderRadius: 8, marginBottom: 10, border: '1px solid var(--dex-gray-200)',
                  }}
                />
              )}
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                {question.options.map((opt, i) => {
                  const isSelected = given.includes(i);
                  return (
                    <button
                      key={i}
                      onClick={() => toggleOption(qIdx, i)}
                      style={{
                        padding: '10px 14px', borderRadius: 10,
                        border: isSelected ? '2px solid var(--dex-green)' : '1px solid var(--dex-gray-200)',
                        background: isSelected ? 'rgba(134,188,37,0.08)' : 'var(--dex-white)',
                        color: 'var(--dex-gray-800)',
                        cursor: 'pointer', textAlign: 'left',
                        fontSize: '0.88rem',
                        transition: 'all 0.15s', display: 'flex', alignItems: 'center', gap: 10,
                      }}
                    >
                      <Icon iconName={isSelected ? 'CheckboxComposite' : 'Checkbox'} style={{ fontSize: 16, flexShrink: 0 }} />
                      {opt}
                    </button>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>

      {/* Navigation */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 16, gap: 8 }}>
        <button
          className="btn btn-secondary"
          onClick={goBack}
          disabled={currentGroupIdx === 0 || isSaving}
          style={{ fontSize: '0.82rem', visibility: currentGroupIdx === 0 ? 'hidden' : 'visible' }}
        >
          {isDe ? 'Zurück' : 'Back'}
        </button>
        <button
          className="btn btn-primary"
          onClick={goNext}
          disabled={isSaving}
          style={{ fontSize: '0.82rem' }}
        >
          {isSaving
            ? (isDe ? 'Speichere…' : 'Saving…')
            : isLastGroup
              ? (isDe ? 'Ergebnis anzeigen' : 'Show result')
              : (isDe ? 'Weiter' : 'Next')}
        </button>
      </div>
    </div>
  );
}

function DocumentsViewer({ documents, t }: { documents: Array<{name: string; url: string; size?: number}>; t: (key: string) => string }): React.ReactElement {
  const [expandedDoc, setExpandedDoc] = React.useState<string | null>(null);
  const [blobUrl, setBlobUrl] = React.useState<string>('');
  const [pdfBlob, setPdfBlob] = React.useState<Blob | null>(null);
  const [loading, setLoading] = React.useState(false);
  // Mobile-Erkennung: Auf Mobile nutzen wir react-pdf (Canvas), auf Desktop bleibt iframe (bewährt)
  const [isMobile, setIsMobile] = React.useState<boolean>(
    typeof window !== 'undefined' && window.matchMedia ? window.matchMedia('(max-width: 768px)').matches : false
  );
  React.useEffect(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return;
    const mq = window.matchMedia('(max-width: 768px)');
    const handler = (e: MediaQueryListEvent): void => setIsMobile(e.matches);
    if (mq.addEventListener) mq.addEventListener('change', handler);
    else mq.addListener(handler);
    return () => {
      if (mq.removeEventListener) mq.removeEventListener('change', handler);
      else mq.removeListener(handler);
    };
  }, []);

  const toggleDoc = async (doc: { url: string; name: string }): Promise<void> => {
    if (expandedDoc === doc.url) {
      setExpandedDoc(null);
      if (blobUrl) { URL.revokeObjectURL(blobUrl); setBlobUrl(''); }
      setPdfBlob(null);
      return;
    }
    setExpandedDoc(doc.url);
    setLoading(true);
    setBlobUrl('');
    setPdfBlob(null);

    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = (window as any).__dexSpfxContext;
      if (!ctx) { setLoading(false); return; }

      // Datei per SPHttpClient REST API als Binary laden
      const siteUrl = ctx.pageContext.web.absoluteUrl;
      const origin = doc.url.match(/^https?:\/\/[^/]+/)?.[0] || '';
      const serverRelPath = decodeURIComponent(doc.url.replace(origin, ''));

      // Pfad-Segmente einzeln encoden (Leerzeichen, Klammern etc.)
      const encodedPath = serverRelPath.split('/').map(s => encodeURIComponent(s)).join('/');
      const apiUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodedPath}')/$value`;

      // XHR für Binary-Download (zuverlässiger als fetch für SharePoint)
      const blob = await new Promise<Blob | null>((resolve) => {
        const xhr = new XMLHttpRequest();
        xhr.open('GET', apiUrl, true);
        xhr.responseType = 'blob';
        xhr.withCredentials = true;
        xhr.setRequestHeader('Accept', '*/*');
        xhr.onload = () => {
          if (xhr.status === 200 && xhr.response) {
            resolve(xhr.response as Blob);
          } else {
            console.warn('[DEX] Doc XHR failed:', xhr.status, apiUrl);
            resolve(null);
          }
        };
        xhr.onerror = () => { console.warn('[DEX] Doc XHR error'); resolve(null); };
        xhr.send();
      });

      if (blob && blob.size > 0) {
        const ext = doc.name.split('.').pop()?.toLowerCase() || '';
        const mimeMap: Record<string, string> = { pdf: 'application/pdf', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' };
        const correctBlob = (mimeMap[ext] && blob.type !== mimeMap[ext]) ? new Blob([blob], { type: mimeMap[ext] }) : blob;
        if (ext === 'pdf' && isMobile) {
          // Mobile: PDF via react-pdf (Canvas) - funktioniert wo iframe versagt
          setPdfBlob(correctBlob);
        } else {
          // Desktop oder Bilder: Blob-URL + iframe (bewährt)
          setBlobUrl(URL.createObjectURL(correctBlob));
        }
      }
    } catch (err) { console.warn('[DEX] Doc viewer error:', err); }
    setLoading(false);
  };

  // Cleanup blob URLs bei Unmount
  React.useEffect(() => {
    return () => { if (blobUrl) URL.revokeObjectURL(blobUrl); };
  }, []);

  return (
    <div style={{ marginTop: 12 }}>
      <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 6 }}>
        {t('myevents.documents')}
      </div>
      {documents.map((doc, i) => {
        const isExpanded = expandedDoc === doc.url;

        return (
          <div key={i} style={{ marginBottom: 6 }}>
            <div style={{
              display: 'flex', alignItems: 'center', gap: 10, padding: '10px 14px',
              background: isExpanded ? 'var(--dex-green-light, #f0fdf4)' : 'var(--dex-gray-100)',
              borderRadius: isExpanded ? '8px 8px 0 0' : 8,
              cursor: 'pointer', fontSize: '0.85rem', color: 'var(--dex-gray-700)',
              transition: 'background 0.15s',
            }} onClick={() => toggleDoc(doc)}>
              <span style={{ width: 32, height: 32, borderRadius: '50%', background: 'var(--dex-green-dark, #6b9a1e)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                <Icon iconName={getDocIconName(doc.name)} style={{ fontSize: 16, color: '#fff' }} />
              </span>
              <span style={{ flex: 1, fontWeight: isExpanded ? 600 : 400 }}>{doc.name}</span>
              {doc.size ? <span style={{ color: 'var(--dex-gray-400)', fontSize: '0.75rem' }}>{(doc.size / 1024).toFixed(0)} KB</span> : null}
              {doc.url && doc.url.startsWith('http') && (
                <a href={doc.url} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()} style={{ color: 'var(--dex-green-dark)', fontSize: '0.72rem', textDecoration: 'none' }}>
                  <Icon iconName="Download" style={{ fontSize: 14 }} />
                </a>
              )}
              <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-400)' }}>{isExpanded ? '▲' : '▼'}</span>
            </div>
            {isExpanded && (
              <div style={{
                border: '1px solid var(--dex-gray-200)', borderTop: 'none',
                borderRadius: '0 0 8px 8px', overflow: 'hidden', background: '#fff',
              }}>
                {loading ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--dex-gray-400)' }}>
                    {t('myevents.agenda') === 'Programm' ? 'Vorschau wird geladen...' : 'Loading preview...'}
                  </div>
                ) : pdfBlob ? (
                  /* PDF via react-pdf (Canvas) - funktioniert Desktop + Mobile, eigenes Scrolling.
                     v20.0: lazy Chunk — Suspense zeigt kurz den Lade-Hinweis. */
                  <React.Suspense fallback={<div style={{ padding: 40, textAlign: 'center', color: 'var(--dex-gray-400)' }}>…</div>}>
                    <PdfViewer blob={pdfBlob} height={600} />
                  </React.Suspense>
                ) : blobUrl ? (
                  /* Desktop-PDF + Bilder via iframe.
                     #view=FitH zwingt das Browser-PDF-Plugin in vertikalen Scroll-Modus
                     (sonst wird oft "Fit page" angenommen und der Scrollbalken fehlt). */
                  <iframe
                    src={doc.name.toLowerCase().endsWith('.pdf') ? `${blobUrl}#view=FitH&toolbar=1` : blobUrl}
                    scrolling="auto"
                    style={{ width: '100%', height: '75vh', minHeight: 600, border: 'none', display: 'block' }}
                    title={doc.name}
                  />
                ) : (
                  <div style={{ padding: 24, textAlign: 'center' }}>
                    <a href={doc.url} target="_blank" rel="noopener noreferrer" style={{ color: 'var(--dex-green-dark)' }}>
                      {t('myevents.agenda') === 'Programm' ? 'Im Browser öffnen' : 'Open in browser'}
                    </a>
                  </div>
                )}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

export default function MyEventsPage(): React.ReactElement {
  const { navigate, selectedEventId, navIntent, clearIntent } = useNavigation();
  const { topLevelEvents, childEventsOf, isEventsLoading, getMyRegistration, getMyEventNumbers, cancelRegistration, cancelTeamMember, updateMyRegistration, switchSplitGroup, listMyEventAttachments, uploadMyEventAttachment, deleteMyEventAttachment, uploadFieldDocument, listFieldDocuments, deleteFieldDocument, registerForEvent, getAllRegistrations, getTeamMembers, addTeamMember, listTeamJoinRequestsForEvent, decideTeamJoinRequest } = useEvents();
  const { currentUser } = useCurrentUser();
  const currentUserEmail = (currentUser?.email || '').toLowerCase();
  // v11.82: Team-Mitglieder pro Event-Karte cachen — Lazy-Load via getTeamMembers.
  // Key = `${eventId}|${teamId}`. Belastet das initiale loadMyRegistrations
  // nicht — nur für Events mit gesetzter TeamId im eigenen Eintrag.
  const [teamMembersCache, setTeamMembersCache] = React.useState<Record<string, SPRegistration[]>>({});
  const teamCacheKeyRef = React.useRef<Set<string>>(new Set());
  const enqueueTeamFetch = React.useCallback((eventId: string, teamId: string): void => {
    if (!eventId || !teamId) return;
    const key = `${eventId}|${teamId}`;
    if (teamCacheKeyRef.current.has(key)) return;
    teamCacheKeyRef.current.add(key);
    getTeamMembers(eventId, teamId).then(list => {
      setTeamMembersCache(prev => ({ ...prev, [key]: list }));
    }).catch(() => {
      setTeamMembersCache(prev => ({ ...prev, [key]: [] }));
    });
  }, [getTeamMembers]);
  // v11.99: Page-Level-Refresh-State entfernt — Header-Button übernimmt.

  // v11.83: Add-Member-Modal + Join-Requests-Cache + Helpers.
  // searchUsers wird für den Add-Member-Picker gebraucht (gleiche API wie
  // im Registrierungs-Formular).
  const { searchUsers, searchUser, isImpersonating } = useRoles();
  const [addMemberDialog, setAddMemberDialog] = React.useState<{
    eventId: string;
    teamId: string;
    teamName: string;
    freeSlots: number;
  } | null>(null);
  const [addMemberPick, setAddMemberPick] = React.useState<{ email: string; displayName: string } | null>(null);
  const [addMemberQuery, setAddMemberQuery] = React.useState('');
  const [addMemberResults, setAddMemberResults] = React.useState<Array<{ email: string; displayName: string }>>([]);
  const [addMemberSearching, setAddMemberSearching] = React.useState(false);
  const [addMemberConsent, setAddMemberConsent] = React.useState(false);
  const [addMemberBusy, setAddMemberBusy] = React.useState(false);
  const [addMemberError, setAddMemberError] = React.useState('');
  const [addMemberIncludeIntl, setAddMemberIncludeIntl] = React.useState(false);
  const addMemberQueryTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  React.useEffect(() => {
    const q = addMemberQuery.trim();
    if (q.length >= 2 && !addMemberPick) {
      (async () => {
        setAddMemberSearching(true);
        try {
          const res = await searchUsers(q, addMemberIncludeIntl);
          setAddMemberResults(res.map(r => ({ email: r.email, displayName: r.displayName })));
        } catch { setAddMemberResults([]); }
        setAddMemberSearching(false);
      })().catch(() => {});
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [addMemberIncludeIntl]);
  const closeAddMemberDialog = (): void => {
    setAddMemberDialog(null);
    setAddMemberPick(null);
    setAddMemberQuery('');
    setAddMemberResults([]);
    setAddMemberConsent(false);
    setAddMemberError('');
    setAddMemberBusy(false);
  };
  const submitAddMember = async (): Promise<void> => {
    if (!addMemberDialog || !addMemberPick || !addMemberConsent || addMemberBusy) return;
    setAddMemberBusy(true);
    setAddMemberError('');
    try {
      const res = await addTeamMember(
        addMemberDialog.eventId,
        addMemberDialog.teamId,
        addMemberDialog.teamName || undefined,
        addMemberPick
      );
      if (!res.ok) {
        if (res.reason && res.reason.startsWith('already-registered')) {
          setAddMemberError(isDe
            ? 'Diese Person ist bereits beim Event angemeldet — bitte abmelden lassen, bevor du sie zum Team hinzufügst.'
            : 'This person is already registered for the event — please have them cancel first before adding to the team.');
        } else if (res.reason === 'team-full') {
          setAddMemberError(isDe ? 'Das Team ist bereits voll.' : 'The team is already full.');
        } else {
          setAddMemberError(isDe ? 'Hinzufügen fehlgeschlagen.' : 'Adding failed.');
        }
        setAddMemberBusy(false);
        return;
      }
      // Team-Cache invalidieren, damit das Badge neu lädt.
      const key = `${addMemberDialog.eventId}|${addMemberDialog.teamId}`;
      teamCacheKeyRef.current.delete(key);
      setTeamMembersCache(prev => {
        const next = { ...prev };
        delete next[key];
        return next;
      });
      closeAddMemberDialog();
    } catch {
      setAddMemberError(isDe ? 'Hinzufügen fehlgeschlagen.' : 'Adding failed.');
      setAddMemberBusy(false);
    }
  };

  // Join-Request-Cache pro (eventId|teamId). Nur für Leads relevant.
  const [joinRequestsCache, setJoinRequestsCache] = React.useState<Record<string, Array<{ Id: number; RequesterEmail: string; RequesterDisplayName: string; Status: string; Created: string }>>>({});
  const joinReqKeyRef = React.useRef<Set<string>>(new Set());
  const enqueueJoinReqFetch = React.useCallback((eventId: string, teamId: string): void => {
    if (!eventId || !teamId) return;
    const key = `${eventId}|${teamId}`;
    if (joinReqKeyRef.current.has(key)) return;
    joinReqKeyRef.current.add(key);
    listTeamJoinRequestsForEvent(eventId, teamId).then(list => {
      setJoinRequestsCache(prev => ({ ...prev, [key]: list }));
    }).catch(() => {
      setJoinRequestsCache(prev => ({ ...prev, [key]: [] }));
    });
  }, [listTeamJoinRequestsForEvent]);
  const [joinReqBusyId, setJoinReqBusyId] = React.useState<number | null>(null);
  const handleDecideJoinRequest = async (eventId: string, teamId: string, requestId: number, decision: 'Approved' | 'Rejected'): Promise<void> => {
    if (joinReqBusyId !== null) return;
    setJoinReqBusyId(requestId);
    try {
      await decideTeamJoinRequest(requestId, decision);
      const key = `${eventId}|${teamId}`;
      joinReqKeyRef.current.delete(key);
      teamCacheKeyRef.current.delete(key);
      setJoinRequestsCache(prev => {
        const next = { ...prev };
        delete next[key];
        return next;
      });
      setTeamMembersCache(prev => {
        const next = { ...prev };
        delete next[key];
        return next;
      });
    } finally {
      setJoinReqBusyId(null);
    }
  };
  // v11.86: „Team verwalten"-Modal. Lead kann pro Mitglied einen
  // Trash-Button anklicken — ein zweites Confirm-Modal fordert die
  // Bestätigung an und ruft anschliessend cancelTeamMember auf.
  const [manageTeamDialog, setManageTeamDialog] = React.useState<{
    eventId: string;
    teamId: string;
    teamName: string;
    teamSize: number;
  } | null>(null);
  const [manageTeamMembers, setManageTeamMembers] = React.useState<SPRegistration[]>([]);
  const [manageTeamConfirm, setManageTeamConfirm] = React.useState<SPRegistration | null>(null);
  const [manageTeamBusyId, setManageTeamBusyId] = React.useState<number | null>(null);
  const closeManageTeamDialog = (): void => {
    setManageTeamDialog(null);
    setManageTeamMembers([]);
    setManageTeamConfirm(null);
    setManageTeamBusyId(null);
  };
  const openManageTeamDialog = (eventId: string, teamId: string, teamName: string, teamSize: number, members: SPRegistration[]): void => {
    setManageTeamDialog({ eventId, teamId, teamName, teamSize });
    setManageTeamMembers(members);
  };
  const performManageTeamCancel = async (member: SPRegistration): Promise<void> => {
    if (!manageTeamDialog || manageTeamBusyId !== null) return;
    setManageTeamBusyId(member.Id);
    try {
      const ok = await cancelTeamMember(manageTeamDialog.eventId, member);
      if (ok) {
        // Cache invalidieren, damit das Team-Badge neu lädt.
        const key = `${manageTeamDialog.eventId}|${manageTeamDialog.teamId}`;
        teamCacheKeyRef.current.delete(key);
        setTeamMembersCache(prev => {
          const next = { ...prev };
          delete next[key];
          return next;
        });
        // Aktualisierte Mitgliederliste neu laden.
        try {
          const fresh = await getTeamMembers(manageTeamDialog.eventId, manageTeamDialog.teamId);
          setTeamMembersCache(prev => ({ ...prev, [key]: fresh }));
          setManageTeamMembers(fresh);
        } catch { /* */ }
        setManageTeamConfirm(null);
      }
    } finally {
      setManageTeamBusyId(null);
    }
  };

  const { t, locale } = useLanguage();
  // v20.4: App-Modals statt nativer Browser-Dialoge.
  const { confirmDialog, showAlert } = useDialog();
  // v20.7: Persönlicher Check-in-QR-Code unter „Meine Events" — gleicher
  // Payload wie die QR-Mail (DEX|<EventNr>|<E-Mail>), client-seitig erzeugt.
  // Damit hat jeder aktive Teilnehmer seinen QR jederzeit zur Hand, auch
  // ohne die Mail zu suchen.
  const [myQrModal, setMyQrModal] = React.useState<{ dataUrl: string; name: string; tid?: number; eventTitle: string } | null>(null);
  const openMyQr = async (ev: DeloitteEvent, reg: SPRegistration): Promise<void> => {
    try {
      const QRCode = await import('qrcode');
      const email = reg.ParticipantEmail || currentUser.email || '';
      const dataUrl = await QRCode.toDataURL(`DEX|${ev.eventNumber}|${email}`, { width: 320, margin: 2 });
      const name = `${reg.Vorname || currentUser.firstName || ''} ${reg.Nachname || currentUser.surname || ''}`.trim() || email;
      setMyQrModal({ dataUrl, name, tid: reg.TeilnehmerID || undefined, eventTitle: ev.title || '' });
    } catch {
      showAlert(isDe ? 'QR-Code konnte nicht erzeugt werden.' : 'QR code could not be generated.', { variant: 'error' });
    }
  };
  const isDe = locale === 'de';
  const [myEvents, setMyEvents] = React.useState<MyEventEntry[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [cancellingId, setCancellingId] = React.useState<string | null>(null);
  const [isCancelling, setIsCancelling] = React.useState(false);
  // v22.46: Erfolgs-Screen nach erfolgreicher Selbst-Abmeldung — analog zum
  // Anmelde-Success-Screen (persönliche Ansprache, Event-Bild, Organizer).
  const [cancelSuccess, setCancelSuccess] = React.useState<null | {
    title: string; imageUrl?: string; organizers: string[]; organizerEmails: string[];
  }>(null);
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [editData, setEditData] = React.useState<Record<string, string>>({});
  const [isSaving, setIsSaving] = React.useState(false);
  const [loadError, setLoadError] = React.useState('');
  // v17.23: Event-Beschreibung auf der MyEvents-Karte ist standardmäßig
  // eingeklappt — pro Event-Id gemerkt, ob der User sie aufgeklappt hat.
  const [descExpanded, setDescExpanded] = React.useState<Record<string, boolean>>({});
  // v11.34: Cascade-Cancel-Dialog (Parent → Sub-Events). State hält
  // den Dialog-Inhalt + die Resolver-Funktion, damit performCancel auf
  // die User-Wahl awaiten kann statt window.confirm.
  const [cascadeDialog, setCascadeDialog] = React.useState<{
    parentTitle: string;
    // v15.8: zusätzlich Start-Datum und Ort pro Sub-Event mitliefern,
    // damit das Modal eine vollständige Info-Liste zeigen kann (vorher
    // nur Titel). Felder optional, fallbacks im Render.
    subEvents: { id: string; title: string; startDate?: string; location?: string }[];
    resolve: (_choice: 'cascade' | 'parent-only' | 'abort') => void;
    /** v14.7: Wenn das Event `requireSubEventSelection` hat, ist „nur
     *  Hauptevent abmelden, Sub-Events behalten" sinnlos (Teilnehmer
     *  konnte sich gar nicht „nur Hauptevent" anmelden). Stattdessen
     *  zeigen wir nur die zwei sinnvollen Optionen — alles abmelden ODER
     *  abbrechen und stattdessen einzelne Sections über „Anmeldung
     *  ändern" abwählen. Die Begriffe „Haupt-/Sub-Event" werden in
     *  diesem Modus durch „Event" + „Event-Sections" ersetzt. */
    isSectionedEvent?: boolean;
  } | null>(null);

  React.useEffect(() => {
    // Warten bis Events fertig geladen sind, sonst zeigen wir Fehler obwohl nur noch geladen wird
    if (isEventsLoading) {
      setIsLoading(true);
      setLoadError('');
      return;
    }
    // v11.79: Stale-while-revalidate. Wenn vor < 60 s schon mal MyEvents
    // geladen wurden, sofort den letzten Stand aus dem sessionStorage
    // rendern (Skeleton-Spinner uebersprungen) und im Hintergrund frisch
    // nachladen. Beim erneuten Klick auf "Meine Events" fühlt sich die
    // Seite damit instantan an, ohne Stale-Daten zu riskieren.
    try {
      const raw = window.sessionStorage.getItem('dex:myevents:cache');
      if (raw) {
        const cache = JSON.parse(raw) as { ts: number; entries: MyEventEntry[] };
        if (cache && Array.isArray(cache.entries) && (Date.now() - cache.ts) < 60_000) {
          setMyEvents(cache.entries);
          setIsLoading(false);
          // im Hintergrund refreshen, kein "loading"-Flag
          loadMyRegistrations(true).catch(() => { /* ignore */ });
          return;
        }
      }
    } catch { /* sessionStorage kann disabled sein — dann normaler Pfad */ }
    loadMyRegistrations();
  }, [topLevelEvents, isEventsLoading]);

  async function loadMyRegistrations(silent: boolean = false): Promise<void> {
    // v11.79: Performance-Logs + Promise.all-Parallelisierung.
    // Vorher: pro angemeldetem Event eine sequentielle getMyRegistration —
    // bei N Anmeldungen N Roundtrips in Serie. Jetzt: alle parallel via
    // Promise.all, dazu Phase-Timer pro Block für ein Vorher/Nachher-
    // Profil in der Browser-Console.
    const tStart = performance.now();
    if (!silent) {
      setIsLoading(true);
      setLoadError('');
    }
    const entries: MyEventEntry[] = [];

    // v18: Im Demo-Modus immer einen Demo-Eintrag in „Meine Events" zeigen
    // (zweite der drei Demo-Sichten). Nicht an eine echte Anmeldung gekoppelt
    // — die Register-Anmeldung ist im Demo deaktiviert; dieser Eintrag zeigt
    // nur, wie eine Anmeldung in „Meine Events" aussieht.
    if (isImpersonating) {
      try {
        const demoEvent = buildDemoShowcaseEvents(isDe ? 'de' : 'en')[0];
        entries.push({
          event: demoEvent,
          registration: buildDemoMyRegistration(currentUser?.email || 'demo.user@deloitte.de', `${currentUser?.firstName || 'Demo'} ${currentUser?.surname || 'User'}`.trim()),
        });
      } catch { /* */ }
    }

    // Schneller Pfad: DEX_Participants abfragen
    const tNums = performance.now();
    const myNumbers = await getMyEventNumbers();
    const allMyNumbers = [...myNumbers.registered, ...myNumbers.waitlisted];
    // v20.0 (Audit): Set statt wiederholtem Array.indexOf in den Filter-Schleifen
    // (vorher O(Events × Anmeldungen) pro Ladevorgang).
    const myNumSet = new Set(allMyNumbers);
    // eslint-disable-next-line no-console
    console.log(`[DEX][perf][myevents] getMyEventNumbers = ${Math.round(performance.now() - tNums)} ms (n=${allMyNumbers.length})`);

    if (allMyNumbers.length > 0) {
      // Nur Events laden die in DEX_Participants stehen
      // Seit v6.4: Sub-Events sind eigene DEX_Events-Items. In "My Events" zeigen
      // wir nur Top-Level-Events; Sub-Event-Anmeldungen erscheinen verschachtelt
      // unter ihrem Parent über childEventsOf().
      const relevantEvents = topLevelEvents.filter(e => e.eventNumber && myNumSet.has(e.eventNumber));
      const tRel = performance.now();
      // v11.79: parallele getMyRegistration-Calls statt sequentielle Schleife.
      const relevantRegs = await Promise.all(relevantEvents.map(async (event) => {
        try { return { event, reg: await getMyRegistration(event.id) }; }
        catch { return { event, reg: null as SPRegistration | null }; }
      }));
      // eslint-disable-next-line no-console
      console.log(`[DEX][perf][myevents] getMyRegistration relevant n=${relevantEvents.length} = ${Math.round(performance.now() - tRel)} ms (parallel)`);
      for (const { event, reg } of relevantRegs) {
        if (!reg) continue;
        // v10.22: Sonderfall — Parent-Reg ist 'Abgemeldet', aber der User
        // hat noch aktive Sub-Event-Registrierungen (Hauptevent abgemeldet,
        // Sub-Events behalten). Dann zeigen wir das Parent als
        // sessionsOnly-Container damit die Sub-Event-Liste sichtbar
        // bleibt und der User weitere Sub-Events nachbuchen oder einzeln
        // abmelden kann. Wenn alles abgemeldet ist, normales Cancelled-
        // Verhalten.
        if (reg.Status === 'Abgemeldet') {
          const kids = childEventsOf(event.id);
          const activeKids = kids.filter(k => k.eventNumber && myNumSet.has(k.eventNumber));
          if (activeKids.length > 0) {
            entries.push({
              event,
              registration: { ...reg, Status: 'Angemeldet' },
              sessionsOnly: true,
              subEventTitles: activeKids.map(k => k.title || (isDe ? 'ohne Titel' : 'untitled')),
            });
          } else {
            entries.push({ event, registration: reg });
          }
        } else {
          entries.push({ event, registration: reg });
        }
      }
      // v6.14: Sessions-Only-Entries. Wenn der User nur Sub-Events eines Parent-Events
      // angemeldet hat (nicht das Parent selbst), zeigen wir den Parent trotzdem als
      // Container mit Badge "Nur Sessions", damit die Session-Anmeldungen sichtbar und
      // managebar bleiben. Wir iterieren über alle Top-Level-Events, für die der Parent
      // KEINE eigene aktive Registrierung hat, und prüfen, ob mindestens ein Child-Event
      // des Parents in allMyNumbers enthalten ist.
      const parentsAlreadyAdded = new Set<string>(entries.map(e => e.event.id));
      for (const parent of topLevelEvents) {
        if (parentsAlreadyAdded.has(parent.id)) continue;
        const kids = childEventsOf(parent.id);
        if (kids.length === 0) continue;
        const activeKids = kids.filter(k => k.eventNumber && myNumSet.has(k.eventNumber));
        if (activeKids.length === 0) continue;
        // Virtuelle Registration — reicht als Platzhalter. Status 'Abgemeldet' würde den
        // Eintrag in den "past"-Bucket werfen; wir nutzen 'Angemeldet' mit sessionsOnly=true
        // als Marker, damit er in "activeEntries" erscheint aber die Parent-Aktionen ausgeblendet werden.
        const virtualReg: SPRegistration = {
          Id: 0,
          Title: '',
          ParticipantName: '',
          ParticipantEmail: '',
          Status: 'Angemeldet',
          RegistrationDate: '',
          CancellationDate: '',
          CustomData: '',
        };
        // v18.53/v18.56: Im subEventsOnlyMode hält die Schatten-Parent-
        // Registrierung die übergreifenden Hauptevent-Antworten (Food
        // Preferences, Hotel, Assistenz etc.) — die brauchen wir für die
        // „Angaben ergänzen"-Vorbefüllung. ABER: der Aktiv/Abgemeldet-Bucket
        // dieser Karte richtet sich nach den AKTIVEN Sub-Event-Anmeldungen
        // (sessionsOnly), NICHT nach dem Status der Schatten-Parent-Zeile (die
        // kann z.B. aus einer alten Abmeldung 'Abgemeldet' sein). Daher NUR die
        // CustomData (+ Id für den Edit-Pfad) übernehmen und den aktiven
        // virtualReg-Status behalten. v18.56-Fix: vorher wurde die komplette
        // realParentReg übernommen → bei Status='Abgemeldet' verschwand das
        // Event fälschlich aus „Meine Events".
        let parentRegForEntry: SPRegistration = virtualReg;
        if (parent.subEventsOnlyMode) {
          try {
            const realParentReg = await getMyRegistration(parent.id);
            if (realParentReg && realParentReg.CustomData) {
              parentRegForEntry = { ...virtualReg, CustomData: realParentReg.CustomData, Id: realParentReg.Id };
            }
          } catch { /* Fallback virtualReg */ }
        }
        entries.push({
          event: parent,
          registration: parentRegForEntry,
          sessionsOnly: true,
          subEventTitles: activeKids.map(k => k.title || (isDe ? 'ohne Titel' : 'untitled')),
        });
      }
      // Zusatzschleife: abgemeldete Events finden. DEX_Participants hält nur
      // EventRegistered/EventOnWaitlist - bei Abmeldung wird die EventNumber dort
      // entfernt. Ohne diese Schleife wären alte Abmeldungen im "My Events"-Tab
      // unsichtbar, sobald der User noch für mind. ein Event angemeldet ist.
      // v10.22: Skip-Check für Events, die wir oben bereits als sessionsOnly-
      // Container eingetragen haben (sonst Doppel-Eintrag mit Status 'Angemeldet'
      // UND 'Abgemeldet' für dasselbe Parent-Event).
      const handledParentIds = new Set(entries.map(e => e.event.id));
      const remainingEvents = topLevelEvents.filter(e =>
        (!e.eventNumber || !myNumSet.has(e.eventNumber)) && !handledParentIds.has(e.id)
      );
      const tRem = performance.now();
      // v11.79: auch die Abgemeldet-Suche parallelisieren. Vorher: pro nicht-
      // angemeldetem Event ein Roundtrip — bei vielen Events der teuerste
      // Block, weil oft 80%+ "kein Eintrag" zurückkommen. Parallel schneidet
      // die Gesamtdauer auf max(slowest), nicht sum().
      const remainingRegs = await Promise.all(remainingEvents.map(async (event) => {
        try { return { event, reg: await getMyRegistration(event.id) }; }
        catch { return { event, reg: null as SPRegistration | null }; }
      }));
      // eslint-disable-next-line no-console
      console.log(`[DEX][perf][myevents] getMyRegistration remaining n=${remainingEvents.length} = ${Math.round(performance.now() - tRem)} ms (parallel, Abgemeldet-Suche)`);
      for (const { event, reg } of remainingRegs) {
        if (reg && reg.Status === 'Abgemeldet') {
          // v10.22: Auch hier den hasActiveChild-Sonderfall prüfen — falls
          // DEX_Participants die Parent-EventNumber noch nicht entfernt hat
          // bzw. der User nur Sub-Event-Anmeldungen hat und kein DEX_Participants-
          // Eintrag für das Parent existiert.
          const kids = childEventsOf(event.id);
          const activeKids = kids.filter(k => k.eventNumber && myNumSet.has(k.eventNumber));
          if (activeKids.length > 0) {
            entries.push({
              event,
              registration: { ...reg, Status: 'Angemeldet' },
              sessionsOnly: true,
              subEventTitles: activeKids.map(k => k.title || (isDe ? 'ohne Titel' : 'untitled')),
            });
          } else {
            entries.push({ event, registration: reg });
          }
        }
      }
    } else {
      // Fallback: Alter Weg für Altdaten ohne DEX_Participants-Eintrag.
      // v11.79: ebenfalls parallelisiert.
      const tFb = performance.now();
      const fbRegs = await Promise.all(topLevelEvents.map(async (event) => {
        try { return { event, reg: await getMyRegistration(event.id) }; }
        catch { return { event, reg: null as SPRegistration | null }; }
      }));
      // eslint-disable-next-line no-console
      console.log(`[DEX][perf][myevents] getMyRegistration fallback n=${topLevelEvents.length} = ${Math.round(performance.now() - tFb)} ms (parallel, Altdaten-Pfad)`);
      for (const { event, reg } of fbRegs) {
        if (reg) entries.push({ event, registration: reg });
      }
    }

    if (entries.length === 0 && allMyNumbers.length > 0) {
      setLoadError('Registrierungen konnten nicht geladen werden.');
    }
    setMyEvents(entries);
    if (!silent) setIsLoading(false);
    // v11.79: Cache für Stale-while-revalidate. Beim nächsten Mount
    // innerhalb 60 s wird sofort der letzte Stand gerendert.
    try {
      window.sessionStorage.setItem('dex:myevents:cache', JSON.stringify({ ts: Date.now(), entries }));
    } catch { /* ignore */ }
    const tTotal = Math.round(performance.now() - tStart);
    // eslint-disable-next-line no-console
    console.log(`[DEX][perf][myevents] total = ${tTotal} ms (entries=${entries.length}, silent=${silent})`);
  }

  // Eigentliche Cancel-Logik (direkt ausführen, ohne 2-Klick-Bestätigung).
  // Wird sowohl von handleCancel (beim 2. Klick) als auch vom Auto-Cancel-
  // Deep-Link (direkt nach Navigation) genutzt.
  const performCancel = async (eventId: string): Promise<void> => {
    // v22.22: Abmeldung von bereits vergangenen Events ist nicht mehr
    // möglich — auch nicht über den Auto-Cancel-Deep-Link aus der Mail
    // (der läuft ebenfalls über performCancel).
    const guardEntry = myEvents.find(e => e.event.id === eventId);
    if (guardEntry && isEventOver(guardEntry.event)) {
      showAlert(isDe
        ? 'Dieses Event liegt bereits in der Vergangenheit — eine Abmeldung ist nicht mehr möglich.'
        : 'This event is already in the past — cancelling is no longer possible.',
        { variant: 'error' });
      setCancellingId(null);
      setIsCancelling(false);
      return;
    }
    setCancellingId(eventId);
    setIsCancelling(true);

    // Check if this is a late cancellation
    const entry = myEvents.find(e => e.event.id === eventId);
    const isLateCancellation = entry?.event.lastDeregisterDate && new Date(entry.event.lastDeregisterDate) < new Date();

    // v11.33/v11.34: Cascade-Prompt via gestyltem Modal — wenn das
    // Hauptevent Sub-Events hat für die der User aktiv angemeldet ist,
    // fragen ob diese auch abgemeldet werden sollen. Drei Auswahlen:
    // - 'cascade'      → Parent + alle Sub-Events abmelden
    // - 'parent-only'  → nur Parent abmelden, Sub-Events behalten
    // - 'abort'        → den ganzen Cancel-Vorgang abbrechen
    const childIdsToCancel: string[] = [];
    if (entry) {
      const kids = childEventsOf(eventId);
      // v15.8: zusätzlich startDate und location pro Sub-Event mitschicken
      // damit das Cancel-Modal Ort + Zeit anzeigen kann.
      const activeKids: { id: string; title: string; startDate?: string; location?: string }[] = [];
      for (const ce of kids) {
        try {
          const reg = await getMyRegistration(ce.id);
          if (reg && reg.Status !== 'Abgemeldet') {
            activeKids.push({
              id: ce.id,
              title: ce.title || (isDe ? 'Sub-Event' : 'Sub-event'),
              startDate: ce.startDate || '',
              location: ce.location || '',
            });
          }
        } catch { /* ignore */ }
      }
      if (activeKids.length > 0) {
        const choice = await new Promise<'cascade' | 'parent-only' | 'abort'>(resolve => {
          setCascadeDialog({
            parentTitle: entry.event.title,
            subEvents: activeKids,
            resolve,
            isSectionedEvent: !!entry.event.requireSubEventSelection,
          });
        });
        setCascadeDialog(null);
        if (choice === 'abort') {
          setCancellingId(null);
          setIsCancelling(false);
          return;
        }
        if (choice === 'cascade') {
          for (const k of activeKids) childIdsToCancel.push(k.id);
        }
      }
    }

    const success = await cancelRegistration(eventId);
    if (success) {
      // Late cancellation: alle Organizer zusammen benachrichtigen (EINE Mail an
      // die semikolon-separierte Liste), im Deloitte-Layout via wrapTemplate,
      // in der konfigurierten Event-Sprache.
      // entry.event.organizers = NAMEN, entry.event.organizerEmails = E-Mails.
      if (isLateCancellation && entry && entry.event.organizerEmails.length > 0) {
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ctx = (window as any).__dexSpfxContext;
          if (ctx) {
            const { EventService } = await import('../services/EventService');
            const svc = new EventService(ctx);
            const userName = `${entry.registration.Vorname || ''} ${entry.registration.Nachname || ''}`.trim() || entry.registration.ParticipantEmail;
            const userEmail = entry.registration.ParticipantEmail || entry.registration.Title;
            const isDe = (entry.event.emailLanguage || 'EN').toUpperCase() === 'DE';
            const deadlineStr = new Date(entry.event.lastDeregisterDate).toLocaleDateString(isDe ? 'de-DE' : 'en-GB');
            const subject = isDe
              ? `Verspätete Abmeldung: ${entry.event.title}`
              : `Late cancellation: ${entry.event.title}`;
            const innerBody = isDe
              ? `<p><strong>${userName}</strong> hat die Anmeldung für <strong>${entry.event.title}</strong> nach Ablauf der Abmeldefrist (${deadlineStr}) storniert.</p><p><strong>E-Mail:</strong> <a href="mailto:${userEmail}">${userEmail}</a></p>`
              : `<p><strong>${userName}</strong> has cancelled their registration for <strong>${entry.event.title}</strong> after the cancellation deadline (${deadlineStr}).</p><p><strong>E-Mail:</strong> <a href="mailto:${userEmail}">${userEmail}</a></p>`;
            const heading = isDe ? 'Verspätete Abmeldung' : 'Late cancellation';
            const subheading = entry.event.title;
            const body = wrapTemplate('#ed8b00', heading, subheading, innerBody);
            // EINE Mail mit ';'-separierter Recipient-Liste - die Recipient-Spalte
            // ist Multi-Line (Note), kann also mehrere E-Mails enthalten. So sehen
            // alle Organizer die Mail gemeinsam (statt N separate Einzel-Mails).
            const toList = entry.event.organizerEmails.join(';');
            const toNames = entry.event.organizers.join(', ') || toList;
            // EmailType 'Info' (Choice-Feld lässt nur Anmeldung/Abmeldung/
            // Warteliste/Nachrücken/Info zu).
            await svc.queueEmail(subject, toList, toNames, body, 'Info', entry.event.title, eventId)
              .catch(err => console.warn('[DEX] LateCancel queueEmail failed:', err));
          }
        } catch { /* Email-Fehler ignorieren */ }
      }
      // v11.33: nach erfolgreicher Parent-Abmeldung optional die ausgewählten
      // Sub-Events kaskadieren. Best-effort — Fehler einzelner Sub-Events
      // brechen den Reload nicht ab.
      for (const childId of childIdsToCancel) {
        try { await cancelRegistration(childId); }
        catch (err) { console.warn('[DEX] cascade-cancel sub-event failed:', childId, err); }
      }
      await loadMyRegistrations();
      // v22.46: Erfolgs-Screen mit persönlicher Ansprache anzeigen.
      if (entry) {
        setCancelSuccess({
          title: entry.event.title,
          imageUrl: entry.event.imageUrl,
          organizers: entry.event.organizers || [],
          organizerEmails: entry.event.organizerEmails || [],
        });
      }
    }
    setCancellingId(null);
    setIsCancelling(false);
  };

  const handleCancel = async (eventId: string): Promise<void> => {
    if (cancellingId === eventId) {
      await performCancel(eventId);
    } else {
      setCancellingId(eventId);
    }
  };

  // Auto-Cancel: wenn die Seite via Deep-Link mit Intent 'auto-cancel' geöffnet
  // wurde (z.B. aus einer Outlook-Decline-Reminder-Mail), die Registrierung
  // direkt stornieren - OHNE dass der User zusätzlich auf "Abmeldung
  // bestätigen" klicken muss. Der Klick auf den Link in der Mail gilt als
  // Bestätigung. Da der User eingeloggt sein muss und nur seine eigene
  // Registrierung cancelt, ist das sicher.
  const didAutoOpen = React.useRef(false);
  React.useEffect(() => {
    if (didAutoOpen.current) return;
    if (navIntent !== 'auto-cancel' || !selectedEventId) return;
    // Warten bis die Registrierungen geladen sind, sonst findet performCancel
    // den entry nicht (late-cancel-check schlägt fehl).
    if (isLoading) return;
    didAutoOpen.current = true;
    clearIntent();
    // Event-Karte einscrollen damit der User den aktualisierten Status sieht
    const el = document.getElementById(`dex-myevent-${selectedEventId}`);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    // Direkt cancellen
    performCancel(selectedEventId).catch(err => console.warn('[DEX] auto-cancel failed:', err));
  }, [navIntent, selectedEventId, isLoading]);

  const activeEntries = myEvents.filter(e => e.registration.Status !== 'Abgemeldet');
  const cancelledEntries = myEvents.filter(e => e.registration.Status === 'Abgemeldet');
  // v22.22: Cluster „Aktive Events" / „Vergangene Events" — gleiche Karte,
  // getrennte Sektionen (plus die bestehende „Abgemeldete Events"-Sektion).
  const upcomingEntries = activeEntries.filter(e => !isEventOver(e.event));
  const pastEntries = activeEntries.filter(e => isEventOver(e.event));

  if (isLoading) {
    // v11.79: Border-Ring-Spinner (v11.33/v11.70) durch eine saubere
    // indeterminierte Progress-Bar ersetzt — der border-radius/border-top-
    // Trick erzeugte in der SP-Hostpage trotz inline-block-Fix einen
    // mitrotierenden vertikalen Strich im Zentrum. Stattdessen jetzt
    // dasselbe Markup wie im App-Boot-Loader (siehe DexEventPlatform.tsx,
    // dexProgressSlide-Keyframe): eine endlos durchlaufende grüne Zone
    // über einer grauen Spur. Kein Rotations-Element, kein Strich-Quirk.
    return (
      <div className="page-container text-center" style={{ padding: '64px 16px' }}>
        <div style={{
          display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 16,
          width: 'min(320px, 80%)',
        }}>
          <div
            aria-hidden="true"
            style={{
              width: '100%', height: 6, borderRadius: 3,
              background: '#e5e5e5',
              overflow: 'hidden', position: 'relative',
            }}
          >
            <div style={{
              position: 'absolute', top: 0, bottom: 0,
              width: '40%',
              background: '#86bc25',
              borderRadius: 3,
              animation: 'dexProgressSlide 1.2s ease-in-out infinite',
            }} />
          </div>
          <p style={{ color: 'var(--dex-gray-600)', margin: 0, fontSize: '0.95rem' }}>
            {t('myevents.loading')}
          </p>
        </div>
        <style>{`@keyframes dexProgressSlide { 0% { left: -40%; } 100% { left: 100%; } }`}</style>
      </div>
    );
  }

  // v22.46: Abmelde-Erfolgs-Screen — gleiche Optik wie der Anmelde-Success
  // (Event-Bild, persönliche Ansprache, Organizer, Buttons).
  if (cancelSuccess) {
    const greet = (currentUser?.firstName || '').trim();
    const orgs = (cancelSuccess.organizers || [])
      .reduce<string[]>((acc, o) => [...acc, ...o.split(';')], [])
      .map(o => {
        const trimmed = o.trim();
        const parts = trimmed.split(',').map(s => s.trim());
        return parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
      })
      .filter(Boolean);
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 720, margin: '0 auto' }}>
          {cancelSuccess.imageUrl && (
            <div style={{
              width: '100%', maxWidth: 480, height: 200, margin: '0 auto 24px',
              borderRadius: 'var(--dex-radius-lg)',
              background: `url(${cancelSuccess.imageUrl}) center/cover no-repeat`,
              filter: 'grayscale(0.4)', opacity: 0.85,
            }} />
          )}
          <h2 style={{ marginTop: 0, color: 'var(--dex-red, #c00)' }}>
            {isDe ? 'Abmeldung bestätigt' : 'Cancellation confirmed'}
          </h2>
          <div className="mt-8" style={{ color: 'var(--dex-gray-700)', textAlign: 'left', maxWidth: 520, margin: '8px auto 0', lineHeight: 1.6 }}>
            <p style={{ margin: '0 0 10px' }}>
              {isDe
                ? <>Hallo{greet ? <> <strong>{greet}</strong></> : ''},</>
                : <>Hi{greet ? <> <strong>{greet}</strong></> : ''},</>}
            </p>
            <p style={{ margin: '0 0 10px' }}>
              {isDe
                ? <>du hast dich erfolgreich vom Event <strong>„{cancelSuccess.title}“</strong> abgemeldet. Dein Platz ist wieder frei{cancelSuccess.imageUrl ? '' : ''} — falls eine Warteliste besteht, rückt automatisch die nächste Person nach.</>
                : <>you have successfully cancelled your registration for <strong>“{cancelSuccess.title}”</strong>. Your spot is free again — if there is a waitlist, the next person is promoted automatically.</>}
            </p>
            <p style={{ margin: 0 }}>
              {isDe
                ? <>Falls du es dir anders überlegst, kannst du dich jederzeit über den Anmelde-Bereich erneut anmelden.</>
                : <>If you change your mind, you can register again anytime via the registration area.</>}
            </p>
          </div>
          {orgs.length > 0 && (
            <div style={{ marginTop: 24, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8 }}>
              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, fontWeight: 600 }}>Organizer</div>
              <OrganizerList names={orgs} emails={cancelSuccess.organizerEmails} size="md" />
            </div>
          )}
          <div style={{ marginTop: 32, display: 'flex', gap: 16, justifyContent: 'center', flexWrap: 'wrap' }}>
            <button className="btn btn-primary" onClick={() => setCancelSuccess(null)}>
              {t('myevents.title')}
            </button>
            <button className="btn btn-secondary" onClick={() => { setCancelSuccess(null); navigate('register'); }}>
              {isDe ? 'Zur Anmeldung' : 'Go to registration'}
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    // v9.9: max-width damit "Meine Events" auf Desktop nicht die volle Breite
    // einnimmt — die einspaltige Karten-Liste sieht sonst auf >1400px-Screens
    // unangenehm gestreckt aus. Inline-style damit das globale .page-container
    // (max-width:100%) andere Seiten nicht beeinflusst.
    <div className="page-container" style={{ maxWidth: 1100, marginLeft: 'auto', marginRight: 'auto' }}>
      <style>{`
        @keyframes dex-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        /* v9.9: Hover-Effekt für den Abmelden-Button — füllt rot, hebt leicht
           an und wirft einen weicheren Schatten, damit die Affordance klar ist.
           Im 2. Klick-Zustand (cancellingId === event.id) bleibt er ohnehin
           rot — nur der idle-State braucht Feedback. */
        .dex-cancel-btn { transition: background 120ms ease, color 120ms ease, transform 120ms ease, box-shadow 120ms ease; }
        .dex-cancel-btn:not(.dex-cancel-btn--armed):hover { background: var(--dex-red) !important; color: #fff !important; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(218,41,28,0.25); }
        .dex-cancel-btn--armed:hover { transform: translateY(-1px); box-shadow: 0 4px 14px rgba(218,41,28,0.4) !important; }
      `}</style>
      {/* v11.99: Page-Level-Refresh-Button entfernt — der Header oben
          rechts hat bereits einen Aktualisieren-Button, doppelt verwirrt. */}
      <div style={{ marginBottom: 16 }}>
        <h2 style={{ margin: 0 }}>{t('myevents.title')}</h2>
      </div>

      {loadError && (
        <div className="card" style={{ padding: 16, marginBottom: 16, color: 'var(--dex-red)' }}>
          {loadError}
        </div>
      )}

      {activeEntries.length === 0 && cancelledEntries.length === 0 && !loadError && (
        <div className="card text-center" style={{ padding: 48 }}>
          <p style={{ color: 'var(--dex-gray-400)' }}>{t('myevents.empty')}</p>
          <button className="btn btn-primary mt-24" onClick={() => navigate('register')}>{t('myevents.browse')}</button>
        </div>
      )}

      {activeEntries.length > 0 && (() => {
        // v22.22: Der Karten-Renderer ist die frühere map-Callback-Funktion —
        // unverändert, nur extrahiert, damit „Aktive Events" und „Vergangene
        // Events" dieselbe Karte in getrennten Clustern rendern können.
        const renderMyEventCard = ({ event, registration, sessionsOnly, subEventTitles }: MyEventEntry): React.ReactElement | null => {
            // Custom Data parsen und IDs zu Labels mappen
            let customData: Record<string, string> = {};
            try {
              if (registration.CustomData) customData = JSON.parse(registration.CustomData);
            } catch { /* */ }

            // Feld-ID zu Label-Map aus den Event-Feldern erstellen.
            // v17.22: im Bilingual-Modus + EN-Locale die EN-Variante des
            // Labels nehmen — vorher rein DE, obwohl der Teilnehmer sich
            // bei der Anmeldung die EN-Labels angesehen hatte.
            const useEnDisplay = locale === 'en' && !!event.bilingualFields;
            const fieldLabelMap: Record<string, string> = {};
            // v17.22: Wert-Uebersetzung für Select-Optionen (DE-Wert →
            // EN-Anzeige) pro Feld, damit auch die Antwort selbst zweisprachig
            // erscheint. Map: fieldId → (DE-Option → EN-Option).
            const fieldOptionEnMap: Record<string, Record<string, string>> = {};
            // v19.34: Feldtyp pro ID merken, damit People-Picker-Antworten
            // (`user`/`roommate`) als Foto-Tag gerendert werden können.
            const fieldTypeMap: Record<string, string> = {};
            for (const field of event.eventSpecificFields) {
              fieldTypeMap[field.id] = field.type;
              fieldLabelMap[field.id] = (useEnDisplay && field.labelEn && field.labelEn.trim())
                ? field.labelEn
                : field.label;
              if (useEnDisplay && field.options && field.optionsEn) {
                const m: Record<string, string> = {};
                field.options.forEach((opt, i) => {
                  const en = (field.optionsEn || [])[i];
                  if (en && en.trim()) m[opt] = en;
                });
                fieldOptionEnMap[field.id] = m;
              }
            }
            // Fallback-Labels für ad-hoc Keys, die NICHT als EventSpecificField
            // registriert sind (z.B. Leistungsnachweis, der direkt aus der Starter-
            // Typ-Sektion kommt). Sonst würde der technische Key angezeigt.
            const adHocLabels: Record<string, string> = {
              b2run_leistungsnachweis: t('reg.starter.proof') || 'Leistungsnachweis vorhanden',
            };

            // "salutation" überspringen (wird schon im Namen angezeigt)
            // Boolean-Werte ('true'/'false') zu lesbarem Ja/Nein konvertieren,
            // damit das UI nicht "true" als technischen String zeigt.
            const yesLabel = t('general.yes') || 'Ja';
            const noLabel = t('general.no') || 'Nein';
            const displayData = Object.keys(customData)
              .filter(key => key !== 'salutation' && customData[key])
              .map(key => {
                const raw = customData[key];
                let value: string = raw;
                if (raw === 'true') value = yesLabel;
                else if (raw === 'false') value = noLabel;
                else if (fieldOptionEnMap[key]) {
                  // v17.22: Select-Antworten im EN-Modus uebersetzen. Multi-
                  // Select-Werte sind " | "-getrennt — jeden Teil einzeln mappen.
                  value = raw.split(' | ').map(part => {
                    const trimmed = part.trim();
                    return fieldOptionEnMap[key][trimmed] || trimmed;
                  }).join(' | ');
                }
                return {
                  label: fieldLabelMap[key] || adHocLabels[key] || key,
                  value,
                  type: fieldTypeMap[key],
                };
              });

            return (
              <div key={event.id} id={`dex-myevent-${event.id}`} className="card my-event-card">
                {/* Header-Zeile: Thumbnail links + Titel/Details rechts */}
                <div style={{ display: 'flex', gap: 16, alignItems: 'flex-start' }}>
                  {event.imageUrl && (
                    <div
                      className="my-event-card__thumb"
                      style={{
                        flexShrink: 0,
                        width: 140,
                        height: 100,
                        borderRadius: 'var(--dex-radius, 12px)',
                        background: 'var(--dex-gray-50, #fafafa)',
                        border: '1px solid var(--dex-gray-200)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        overflow: 'hidden',
                      }}
                    >
                      <CachedImg
                        src={event.imageUrl}
                        alt={event.title}
                        loading="lazy"
                        decoding="async"
                        style={{
                          maxWidth: '100%',
                          maxHeight: '100%',
                          objectFit: 'contain',
                          display: 'block',
                        }}
                      />
                    </div>
                  )}
                  <div style={{ flex: 1, minWidth: 0 }}>
                    {/* Header: Titel + Status Badge */}
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                      <h3 style={{ margin: 0, fontSize: '1.1rem' }}>{event.title}</h3>
                      {sessionsOnly ? (() => {
                        // v15.15: Im subEventsOnlyMode komplett ausblenden —
                        // Badge UND Hinweisbox sind dort redundant, weil
                        // „nur Sub-Events"-Anmeldung der einzig mögliche
                        // Zustand ist (kein Hauptevent zum Anmelden).
                        if (event.subEventsOnlyMode) return null;
                        const term = event.childEventTermPlural || '';
                        const badgeText = term
                          ? (isDe ? `Nur ${term}` : `${term} only`)
                          : (t('myevents.sessionsonly.badge') || 'Nur Sub-Events');
                        return (
                          <span
                            className="badge"
                            title={t('myevents.sessionsonly.hint')}
                            style={{
                              flexShrink: 0, marginLeft: 12,
                              background: 'var(--dex-orange, #ed8b00)', color: '#fff',
                            }}
                          >
                            {badgeText}
                          </span>
                        );
                      })() : (
                        <span className={`badge ${getStatusBadgeClass(registration.Status)}`} style={{ flexShrink: 0, marginLeft: 12 }}>
                          {registration.Status === 'Warteliste' && registration.TeilnehmerID && event.maxParticipants > 0
                            ? `${getStatusLabel(registration.Status, t)} #${registration.TeilnehmerID - event.maxParticipants}`
                            : getStatusLabel(registration.Status, t)}
                        </span>
                      )}
                    </div>
                    {/* v15.15: Hinweisbox „nur für Sub-Events angemeldet"
                        nur außerhalb des subEventsOnlyMode anzeigen — dort
                        ist sie redundant, weil es gar keine andere Option
                        gibt. */}
                    {sessionsOnly && !event.subEventsOnlyMode && (() => {
                      const term = event.childEventTermPlural || '';
                      const subList = (subEventTitles && subEventTitles.length > 0)
                        ? ` (${subEventTitles.join(', ')})`
                        : '';
                      const hintText = term
                        ? (isDe
                            ? `Du bist für ${term} dieses Events angemeldet, aber NICHT für das Haupt-Event selbst`
                            : `You are registered for ${term} of this event but NOT for the main event itself`)
                        : t('myevents.sessionsonly.hint');
                      return (
                        <div style={{
                          marginTop: 6, padding: '6px 10px', borderRadius: 6,
                          background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                          color: 'var(--dex-orange)', fontSize: '0.78rem',
                        }}>
                          {hintText}{subList}.
                        </div>
                      );
                    })()}

                    {/* v11.82: Team-Badge — sichtbar wenn die eigene Anmeldung
                        eine TeamId hat. Lazy-Load der anderen Mitglieder via
                        getTeamMembers. Zeigt: Team-Name (falls vorhanden),
                        Belegungs-Anzahl, Liste der Mitglieder (Lead zuerst). */}
                    {registration.TeamId && (() => {
                      const cacheKey = `${event.id}|${registration.TeamId}`;
                      const cached = teamMembersCache[cacheKey];
                      if (!cached) {
                        // Async laden (nur einmal pro Karte, idempotent ueber ref-Set).
                        enqueueTeamFetch(event.id, registration.TeamId);
                        return null;
                      }
                      const activeMembers = cached.filter(m => m.Status !== 'Abgemeldet');
                      const total = activeMembers.length;
                      const teamSizeCfg = event.teamSize || total;
                      const tn = registration.TeamName || (cached.find(m => m.TeamName)?.TeamName) || '';
                      const teamTermS = event.teamTermSingular || 'Team';
                      const isLead = !!registration.TeamLead;
                      // v11.86: Sortierung — Lead zuerst, dann nach TeilnehmerID,
                      // dann nach Id. Abgemeldete Mitglieder werden ausgegraut
                      // mit eigenem Badge weiter unten gerendert.
                      const sortedAll = [...cached].sort((a, b) => {
                        const aLead = a.TeamLead ? 0 : 1;
                        const bLead = b.TeamLead ? 0 : 1;
                        if (aLead !== bLead) return aLead - bLead;
                        const aTid = typeof a.TeilnehmerID === 'number' ? a.TeilnehmerID : Number.MAX_SAFE_INTEGER;
                        const bTid = typeof b.TeilnehmerID === 'number' ? b.TeilnehmerID : Number.MAX_SAFE_INTEGER;
                        if (aTid !== bTid) return aTid - bTid;
                        return a.Id - b.Id;
                      });
                      return (
                        <div style={{
                          marginTop: 8, padding: '10px 14px', borderRadius: 6,
                          background: 'rgba(134,188,37,0.08)', border: '1px solid var(--dex-green, #86bc25)',
                          color: 'var(--dex-green-dark, #3f5f10)', fontSize: '0.82rem', lineHeight: 1.45,
                        }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                            <Icon iconName="People" style={{ fontSize: 14 }} />
                            <strong>
                              {isDe
                                ? `${teamTermS} „${tn || 'Unbenannt'}" — ${total}/${teamSizeCfg} belegt`
                                : `${teamTermS} „${tn || 'Unnamed'}" — ${total}/${teamSizeCfg} taken`}
                            </strong>
                            {isLead && (
                              <span style={{
                                padding: '1px 8px', borderRadius: 999,
                                background: 'var(--dex-green, #86bc25)', color: '#fff',
                                fontSize: '0.7rem', fontWeight: 600,
                              }}>
                                {isDe ? `du bist ${teamTermS}-Lead` : `you are ${teamTermS} lead`}
                              </span>
                            )}
                          </div>
                          {/* v11.86: Mitglieder-Karten — pro Person Foto + Name + Email + Standort. */}
                          {sortedAll.length > 0 && (
                            <div style={{ marginTop: 10, display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {sortedAll.map(m => {
                                const isCancelled = m.Status === 'Abgemeldet';
                                const isMemberLead = !!m.TeamLead && !isCancelled;
                                const fullName = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantEmail;
                                const loc = (m.Location || '').trim();
                                return (
                                  <div
                                    key={m.Id}
                                    style={{
                                      display: 'flex', alignItems: 'center', gap: 12,
                                      padding: '6px 8px', borderRadius: 6,
                                      background: isCancelled ? 'rgba(0,0,0,0.04)' : '#fff',
                                      border: '1px solid var(--dex-gray-200)',
                                      opacity: isCancelled ? 0.55 : 1,
                                      position: 'relative',
                                    }}
                                  >
                                    <div
                                      style={{
                                        width: 40, height: 40, borderRadius: '50%',
                                        overflow: 'visible', flexShrink: 0,
                                        position: 'relative',
                                      }}
                                    >
                                      <img
                                        src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=L`}
                                        alt={fullName}
                                        onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                        style={{
                                          width: 40, height: 40, borderRadius: '50%',
                                          objectFit: 'cover', background: 'var(--dex-gray-100)',
                                          transition: 'transform 160ms ease, box-shadow 160ms ease',
                                          transformOrigin: 'left center',
                                          /* v11.94: kein zoom-in-Cursor */
                                        }}
                                        onMouseEnter={e => {
                                          if (isCancelled) return;
                                          const img = e.currentTarget as HTMLImageElement;
                                          img.style.transform = 'scale(2.4)';
                                          img.style.boxShadow = '0 8px 24px rgba(0,0,0,0.35)';
                                          img.style.zIndex = '50';
                                          img.style.position = 'relative';
                                        }}
                                        onMouseLeave={e => {
                                          const img = e.currentTarget as HTMLImageElement;
                                          img.style.transform = 'scale(1)';
                                          img.style.boxShadow = 'none';
                                          img.style.zIndex = '';
                                          img.style.position = '';
                                        }}
                                      />
                                    </div>
                                    <div style={{ flex: 1, minWidth: 0 }}>
                                      <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)', fontSize: '0.85rem' }}>
                                        {fullName}
                                      </div>
                                      <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)' }}>
                                        {m.ParticipantEmail}{loc ? ` · ${loc}` : ''}
                                      </div>
                                    </div>
                                    {isCancelled ? (
                                      <span style={{
                                        padding: '2px 8px', borderRadius: 999,
                                        background: 'var(--dex-gray-300, #c8c8c8)', color: '#fff',
                                        fontSize: '0.68rem', fontWeight: 600, flexShrink: 0,
                                      }}>
                                        {isDe ? 'abgemeldet' : 'cancelled'}
                                      </span>
                                    ) : isMemberLead && (
                                      <span style={{
                                        padding: '2px 8px', borderRadius: 999,
                                        background: 'var(--dex-green, #86bc25)', color: '#fff',
                                        fontSize: '0.68rem', fontWeight: 600, flexShrink: 0,
                                      }}>
                                        Lead
                                      </span>
                                    )}
                                  </div>
                                );
                              })}
                            </div>
                          )}
                          {/* v11.83/v11.86: Aktion-Buttons — Edit (alle Leads) +
                              Add (nur bei freien Slots). */}
                          {isLead && (
                            <div style={{ marginTop: 10, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                              <button
                                type="button"
                                className="btn btn-secondary"
                                onClick={() => openManageTeamDialog(
                                  event.id,
                                  registration.TeamId!,
                                  tn || '',
                                  teamSizeCfg,
                                  cached
                                )}
                                style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                              >
                                <Icon iconName="Edit" style={{ fontSize: 12, marginRight: 6 }} />
                                {isDe ? 'Team bearbeiten' : 'Edit team'}
                              </button>
                              {total < teamSizeCfg && (
                                <button
                                  type="button"
                                  className="btn btn-secondary"
                                  onClick={() => {
                                    setAddMemberDialog({
                                      eventId: event.id,
                                      teamId: registration.TeamId!,
                                      teamName: tn || '',
                                      freeSlots: teamSizeCfg - total,
                                    });
                                    setAddMemberPick(null);
                                    setAddMemberQuery('');
                                    setAddMemberResults([]);
                                    setAddMemberConsent(false);
                                    setAddMemberError('');
                                  }}
                                  style={{ fontSize: '0.78rem', padding: '4px 10px' }}
                                >
                                  <Icon iconName="AddFriend" style={{ fontSize: 12, marginRight: 6 }} />
                                  {isDe
                                    ? `Mitglied hinzufügen (${teamSizeCfg - total} Slot${(teamSizeCfg - total) === 1 ? '' : 's'} frei)`
                                    : `Add member (${teamSizeCfg - total} slot${(teamSizeCfg - total) === 1 ? '' : 's'} free)`}
                                </button>
                              )}
                            </div>
                          )}
                          {/* v11.83: Beitritts-Anfragen-Block — nur für Leads, wenn das
                              Event Approval aktiviert hat UND es Pending-Anfragen gibt. */}
                          {isLead && event.teamJoinRequiresApproval && (() => {
                            const jKey = `${event.id}|${registration.TeamId}`;
                            const jReqs = joinRequestsCache[jKey];
                            if (jReqs === undefined) {
                              enqueueJoinReqFetch(event.id, registration.TeamId!);
                              return null;
                            }
                            if (jReqs.length === 0) return null;
                            return (
                              <div style={{
                                marginTop: 10,
                                padding: '8px 12px',
                                borderRadius: 6,
                                background: 'rgba(237,139,0,0.10)',
                                border: '1px solid var(--dex-orange, #ed8b00)',
                                color: '#7a4a00',
                              }}>
                                <div style={{ fontWeight: 700, marginBottom: 6 }}>
                                  {isDe ? 'Beitritts-Anfragen' : 'Join requests'} ({jReqs.length})
                                </div>
                                {jReqs.map(r => {
                                  const busy = joinReqBusyId === r.Id;
                                  return (
                                    <div key={r.Id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', borderTop: '1px solid rgba(237,139,0,0.25)' }}>
                                      <div style={{ flex: 1, fontSize: '0.82rem' }}>
                                        <div style={{ fontWeight: 600 }}>{r.RequesterDisplayName || r.RequesterEmail}</div>
                                        <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-600)' }}>{r.RequesterEmail}</div>
                                      </div>
                                      <button
                                        type="button"
                                        className="btn btn-primary"
                                        disabled={busy}
                                        onClick={() => handleDecideJoinRequest(event.id, registration.TeamId!, r.Id, 'Approved')}
                                        style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                      >
                                        {isDe ? 'Bestätigen' : 'Approve'}
                                      </button>
                                      <button
                                        type="button"
                                        className="btn btn-secondary"
                                        disabled={busy}
                                        onClick={() => handleDecideJoinRequest(event.id, registration.TeamId!, r.Id, 'Rejected')}
                                        style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                                      >
                                        {isDe ? 'Ablehnen' : 'Reject'}
                                      </button>
                                    </div>
                                  );
                                })}
                              </div>
                            );
                          })()}
                        </div>
                      );
                    })()}

                    {/* Kompakte Info-Zeile: Location + Datum inline, umbricht auf schmalen Bildschirmen */}
                    <div style={{ marginTop: 8, display: 'flex', flexWrap: 'wrap', alignItems: 'flex-start', gap: '6px 24px', fontSize: '0.88rem', color: 'var(--dex-gray-700)' }}>
                      <div style={{ display: 'flex', alignItems: 'flex-start', gap: 6 }}>
                        <Icon iconName="MapPin" style={{ fontSize: 14, color: 'var(--dex-gray-500)', marginTop: 2 }} />
                        <div>
                          <div style={{ fontWeight: 700, color: 'var(--dex-gray-800)' }}>{event.location || '-'}</div>
                          {event.locationAddress && (event.locationAddress.street || event.locationAddress.city) && (
                            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 1 }}>
                              {[event.locationAddress.street, event.locationAddress.houseNo].filter(Boolean).join(' ')}
                              {(event.locationAddress.zip || event.locationAddress.city) ? ', ' : ''}
                              {[event.locationAddress.zip, event.locationAddress.city].filter(Boolean).join(' ')}
                            </div>
                          )}
                        </div>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}><Icon iconName="Calendar" style={{ fontSize: 14, color: 'var(--dex-gray-500)' }} /> {formatDateRange(event.startDate, event.endDate)}</div>
                    </div>

                    {/* Organizer mit Foto (Hover vergrößert). v18.9: optional ausgeblendet. */}
                    {!event.hideOrganizer && event.organizers.length > 0 && (
                      <div style={{ marginTop: 10 }}>
                        <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4 }}>Organizer</div>
                        <OrganizerList
                          names={event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => {
                            const trimmed = o.trim();
                            const parts = trimmed.split(',').map(s => s.trim());
                            return parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
                          }).filter(Boolean)}
                          emails={event.organizerEmails}
                          size="sm"
                        />
                      </div>
                    )}
                    {/* v10.26: Optionaler Ansprechpartner — frei eingegebene Person
                        außerhalb des App-User-Pools. Reines Anzeige-Feld; Mailto-Link
                        wenn Email gesetzt. Wird nur gerendert wenn mindestens Name
                        oder Email gepflegt sind. Spiegelt das Verhalten der
                        Registration-Page in My Events wider. */}
                    {(event.contactName || event.contactEmail || event.contactInfo) && (
                      <div style={{ marginTop: 10, padding: '8px 10px', background: 'var(--dex-gray-50, #f7f7f7)', borderRadius: 6, border: '1px solid var(--dex-gray-200)' }}>
                        <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4 }}>
                          {(event.emailLanguage || 'EN').toUpperCase() === 'DE' ? 'Ansprechpartner' : 'Contact'}
                        </div>
                        {event.contactName && (
                          <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>{event.contactName}</div>
                        )}
                        {event.contactEmail && (
                          <div style={{ fontSize: '0.78rem', marginTop: 2 }}>
                            <a href={`mailto:${event.contactEmail}`} style={{ color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'none' }}>{event.contactEmail}</a>
                          </div>
                        )}
                        {event.contactInfo && (
                          <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-700)', marginTop: 4, whiteSpace: 'pre-wrap', lineHeight: 1.45 }}>{event.contactInfo}</div>
                        )}
                      </div>
                    )}
                  </div>
                </div>

                {/* v17.22: Event-Beschreibung auch unter „Meine Events"
                    anzeigen (vorher nur auf der Anmeldeseite). RichText-HTML
                    aus dem eigenen Tenant — gleiche Render-Logik wie auf der
                    RegistrationPage (HTML erlaubt, sonst \n→<br>).
                    v17.23: standardmäßig eingeklappt, per Button aufklappbar. */}
                {event.description && (!editingId || editingId !== event.id) && (() => {
                  const isOpen = !!descExpanded[event.id];
                  return (
                    <div style={{ marginTop: 10 }}>
                      <button
                        type="button"
                        onClick={() => setDescExpanded(prev => ({ ...prev, [event.id]: !prev[event.id] }))}
                        aria-expanded={isOpen}
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: 6,
                          background: 'none', border: 'none', cursor: 'pointer', padding: 0,
                          fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-green-dark, #4a7c1f)',
                        }}
                      >
                        <span style={{
                          display: 'inline-block', transition: 'transform 0.15s',
                          transform: isOpen ? 'rotate(90deg)' : 'rotate(0deg)', fontSize: '0.7rem',
                        }}>▶</span>
                        {isDe ? 'Beschreibung' : 'Description'}
                      </button>
                      {isOpen && (
                        <div
                          className="dex-event-desc"
                          style={{
                            marginTop: 6, padding: '10px 14px',
                            color: 'var(--dex-gray-700)', background: 'var(--dex-gray-50, #fafafa)',
                            borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-gray-200)',
                            wordBreak: 'break-word',
                          }}
                          dangerouslySetInnerHTML={{
                            __html: (() => {
                              const raw = event.description || '';
                              const isHtml = /<[a-z][\s\S]*>/i.test(raw);
                              return isHtml
                                ? raw
                                : raw.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
                            })(),
                          }}
                        />
                      )}
                    </div>
                  );
                })()}

                {/* v10.26: Custom-Field-Antworten als rechteckige
                    pastellgrüne Tags — visuell eindeutig von den
                    abgerundeten grauen Organizer-Chips getrennt, damit der
                    User sofort sieht: das ist „was ich angegeben habe", nicht
                    „wer organisiert das". */}
                {!editingId || editingId !== event.id ? (
                  // v18.38: Edit-Bereich anzeigen, sobald das Event ueberhaupt
                  // bearbeitbare Felder hat — NICHT mehr nur wenn schon Werte
                  // ausgefüllt sind. Sonst kann ein Teilnehmer, der ein
                  // optionales Feld leer gelassen hat (z.B. „zusätzliche
                  // Nacht"), es später nicht mehr nachtragen.
                  // v20.9 BUG-FIX: Der Block rendert jetzt AUCH, wenn das Event
                  // keine Custom-Felder hat, aber die Anmeldung QR-fähig ist —
                  // sonst fehlte der „Mein QR-Code"-Button bei Events ohne
                  // Abfragefelder (z.B. einfaches Sommerfest).
                  (displayData.length > 0 || (event.eventSpecificFields || []).filter((f: EventSpecificField) => f.label).length > 0 || (!sessionsOnly && !!event.eventNumber && ['Angemeldet', 'QR versendet', 'Eingecheckt'].indexOf(registration.Status) >= 0)) && (
                    <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 6, alignItems: 'center' }}>
                      {displayData.map(({ label, value, type }) => (
                        <FieldAnswerTag key={label} label={label} value={value} type={type} />
                      ))}
                      {/* v11.30: Edit-Button direkt neben den Angaben-Tags
                          (statt unten in der Aktions-Zeile). Näher am Inhalt
                          den er bearbeitet.
                          v18.38: zeigt jetzt „Angaben ergänzen", wenn noch
                          nichts ausgefüllt wurde — sonst „Angaben bearbeiten". */}
                      {(event.eventSpecificFields || []).filter((f: EventSpecificField) => f.label).length > 0 && (
                      <button
                        type="button"
                        className="btn btn-outline"
                        onClick={() => { setEditData(customData); setEditingId(event.id); }}
                        style={{
                          fontSize: '0.78rem', padding: '5px 12px', borderRadius: 6,
                          width: 'auto', cursor: 'pointer',
                          display: 'inline-flex', alignItems: 'center', gap: 4,
                        }}
                        title={displayData.length > 0 ? t('myevents.edit') : (isDe ? 'Angaben ergänzen' : 'Add details')}
                      >
                        <Pencil size={12} /> {displayData.length > 0 ? t('myevents.edit') : (isDe ? 'Angaben ergänzen' : 'Add details')}
                      </button>
                      )}
                      {/* v20.7: Persönlicher Check-in-QR — für aktive
                          Anmeldungen jederzeit abrufbar (gleicher Code wie
                          in der QR-Mail). */}
                      {!sessionsOnly && !!event.eventNumber && ['Angemeldet', 'QR versendet', 'Eingecheckt'].indexOf(registration.Status) >= 0 && (
                        <button
                          type="button"
                          className="btn btn-outline"
                          onClick={() => { openMyQr(event, registration).catch(() => { /* */ }); }}
                          style={{
                            fontSize: '0.78rem', padding: '5px 12px', borderRadius: 6,
                            width: 'auto', cursor: 'pointer',
                            display: 'inline-flex', alignItems: 'center', gap: 4,
                          }}
                          title={isDe ? 'Deinen persönlichen Check-in-QR-Code anzeigen' : 'Show your personal check-in QR code'}
                        >
                          <QrCode size={12} /> {isDe ? 'Mein QR-Code' : 'My QR code'}
                        </button>
                      )}
                    </div>
                  )
                ) : (
                  <div style={{ marginTop: 12 }}>
                    {/* v17.22: EN-Varianten auch im „Meine Events"-Edit-Formular
                        berücksichtigen — vorher rein DE, obwohl der Teilnehmer
                        sich auf der Anmeldeseite die EN-Labels angesehen hatte. */}
                    {(() => {
                      const useEnEdit = locale === 'en' && !!event.bilingualFields;
                      const eLabel = (f: EventSpecificField): string =>
                        (useEnEdit && f.labelEn && f.labelEn.trim()) ? f.labelEn : f.label;
                      const eOpt = (f: EventSpecificField, opt: string, idx: number): string =>
                        (useEnEdit && f.optionsEn && f.optionsEn[idx] && f.optionsEn[idx].trim()) ? f.optionsEn[idx] : opt;
                      return event.eventSpecificFields.map((field: EventSpecificField) => (
                        <div className="form-group" key={field.id} style={{ marginBottom: 10 }}>
                          <label className="form-label" style={{ fontSize: '0.82rem', marginBottom: 2 }}>
                            {field.required && <span className="required">*</span>}
                            {eLabel(field)}
                          </label>
                          {field.type === 'select' ? (
                            <select className="form-select" value={editData[field.id] || ''} onChange={e => setEditData({ ...editData, [field.id]: e.target.value })}>
                              <option value="">—</option>
                              {field.options && field.options.map((opt, optIdx) => <option key={opt} value={opt}>{eOpt(field, opt, optIdx)}</option>)}
                            </select>
                          ) : (field.type === 'user' || field.type === 'roommate') ? (
                            /* v18.61: People-Picker-Felder beim Bearbeiten wieder als
                               echter People-Picker mit Profilfoto (vorher Rohtext). */
                            <UserFieldPicker
                              value={editData[field.id] || ''}
                              onChange={v => setEditData({ ...editData, [field.id]: v })}
                              searchUsers={searchUsers}
                              searchUserByEmail={searchUser}
                              placeholder={locale === 'de' ? 'Name oder E-Mail eingeben…' : 'Type a name or email…'}
                              errorStyle={{}}
                            />
                          ) : field.type === 'checkbox' ? (
                            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
                              <input
                                type="checkbox"
                                checked={editData[field.id] === 'true'}
                                onChange={e => setEditData({ ...editData, [field.id]: e.target.checked ? 'true' : 'false' })}
                                style={{ marginTop: 2 }}
                              />
                              <span>{(useEnEdit && field.confirmLabelEn && field.confirmLabelEn.trim() ? field.confirmLabelEn : field.confirmLabel) || eLabel(field)}</span>
                            </label>
                          ) : (
                            <input className="form-input" value={editData[field.id] || ''} onChange={e => setEditData({ ...editData, [field.id]: e.target.value })} placeholder={eLabel(field)} type={field.type === 'number' ? 'number' : 'text'} />
                          )}
                        </div>
                      ));
                    })()}
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button className="btn btn-primary" style={{ fontSize: '0.82rem' }} disabled={isSaving} onClick={async () => { setIsSaving(true); await updateMyRegistration(event.id, editData); await loadMyRegistrations(); setEditingId(null); setIsSaving(false); }}>
                        {isSaving ? t('myevents.saving') : t('myevents.save')}
                      </button>
                      <button className="btn btn-secondary" style={{ fontSize: '0.82rem' }} onClick={() => setEditingId(null)}>{t('general.cancel')}</button>
                    </div>
                  </div>
                )}

                {/* Agenda / Timeline - mehrspaltig bei mehreren Tagen, horizontal scrollbar auf Mobile */}
                {event.agenda && event.agenda.length > 0 && (() => {
                  const grouped = Object.entries(
                    event.agenda.reduce((groups: Record<string, AgendaItem[]>, item: AgendaItem) => {
                      const key = item.date || 'TBD';
                      if (!groups[key]) groups[key] = [];
                      groups[key].push(item);
                      return groups;
                    }, {} as Record<string, AgendaItem[]>)
                  ).sort(([a], [b]) => a.localeCompare(b));
                  const dayCount = grouped.length;
                  // v22.36: Durchlaufende Nummerierung der Agenda-Schritte
                  // (über alle Tage, sortiert nach Datum + Uhrzeit) — die
                  // Einzel-Schritte zeigen keine Icons mehr, nur die Sektion.
                  const agendaOrderIds = event.agenda
                    .slice()
                    .sort((a: AgendaItem, b: AgendaItem) => ((a.date || '') + (a.time || '')).localeCompare((b.date || '') + (b.time || '')))
                    .map((x: AgendaItem) => x.id);

                  return (
                    <div style={{ marginTop: 12 }}>
                      <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 8, display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Icon iconName="Calendar" style={{ fontSize: 14, color: 'var(--dex-green-dark, #6b9a1e)' }} />
                        {t('myevents.agenda')} {dayCount > 1 && <span style={{ fontWeight: 400, fontSize: '0.72rem', color: 'var(--dex-gray-400)' }}>· {dayCount} {t('myevents.agenda') === 'Programm' ? 'Tage (seitwärts scrollen)' : 'days (swipe)'}</span>}
                      </div>
                      {/* Horizontal scrollbarer Container - funktioniert auf Desktop und Mobile */}
                      <div
                        style={{
                          display: 'flex',
                          flexWrap: 'nowrap',
                          gap: 16,
                          overflowX: 'auto',
                          // eslint-disable-next-line @typescript-eslint/no-explicit-any
                          WebkitOverflowScrolling: 'touch' as any,
                          scrollSnapType: 'x mandatory',
                          paddingBottom: 4,
                        }}
                      >
                        {grouped.map(([date, items]) => (
                          <div key={date} style={{
                            flex: `0 0 ${dayCount === 1 ? '100%' : dayCount === 2 ? 'calc(50% - 8px)' : 'min(280px, 85%)'}`,
                            scrollSnapAlign: 'start',
                            background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, padding: 12,
                            border: '1px solid var(--dex-gray-200)',
                            minWidth: 0,
                          }}>
                            <div style={{
                              fontSize: '0.78rem', fontWeight: 700, color: '#fff', marginBottom: 8,
                              background: 'var(--dex-green-dark, #6b9a1e)', borderRadius: 8, padding: '6px 12px',
                              textAlign: 'center',
                            }}>
                              {date !== 'TBD' ? new Date(date + 'T00:00').toLocaleDateString('de-DE', { weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric' }) : 'TBD'}
                            </div>
                            {items.sort((a: AgendaItem, b: AgendaItem) => (a.time || '').localeCompare(b.time || '')).map((item: AgendaItem) => (
                              <div key={item.id} style={{
                                display: 'flex', alignItems: 'flex-start', gap: 8, padding: '6px 0',
                                borderLeft: '2px solid var(--dex-green)', marginLeft: 4, paddingLeft: 10,
                              }}>
                                <span style={{ width: 26, height: 26, borderRadius: '50%', background: 'var(--dex-green-dark, #6b9a1e)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, color: '#fff', fontWeight: 700, fontSize: '0.72rem', lineHeight: 1 }}>
                                  {agendaOrderIds.indexOf(item.id) + 1}
                                </span>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontSize: '0.8rem', fontWeight: 600 }}>
                                    {item.time}{item.endTime ? ` – ${item.endTime}` : ''}
                                  </div>
                                  <div style={{ fontSize: '0.8rem', wordBreak: 'break-word' }}>{item.title}</div>
                                  {item.description && (
                                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 1, wordBreak: 'break-word' }}>{item.description}</div>
                                  )}
                                </div>
                              </div>
                            ))}
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })()}

                {/* Transferzeiten */}
                {event.transferTimes && event.transferTimes.length > 0 && (
                  <div style={{ marginTop: 12 }}>
                    <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 6 }}>
                      {t('myevents.transfers')}
                    </div>
                    {event.transferTimes.sort((a: TransferTime, b: TransferTime) => (a.date + a.departureTime).localeCompare(b.date + b.departureTime)).map((tr: TransferTime) => (
                      <div key={tr.id} style={{
                        display: 'flex', gap: 10, padding: '8px 12px', marginBottom: 6, fontSize: '0.82rem',
                        background: 'var(--dex-gray-50, #fafafa)', borderRadius: 10,
                        borderLeft: '3px solid var(--dex-orange)',
                      }}>
                        <span style={{ width: 28, height: 28, borderRadius: '50%', background: 'var(--dex-orange)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                          <Icon iconName="Bus" style={{ fontSize: 13, color: '#fff' }} />
                        </span>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontWeight: 600 }}>
                            {tr.location}{tr.meetingPoint ? ` – ${tr.meetingPoint}` : ''}
                          </div>
                          {tr.address && (
                            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>
                              <Icon iconName="MapPin" style={{ fontSize: 11, marginRight: 4 }} />{tr.address}
                            </div>
                          )}
                          <div style={{ marginTop: 2 }}>
                            <Icon iconName="Calendar" style={{ fontSize: 11, color: 'var(--dex-gray-400)', marginRight: 4 }} />
                            {new Date(tr.date + 'T00:00').toLocaleDateString('de-DE', {weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric'})},
                            {' '}{tr.departureTime}{tr.arrivalTime ? ` → ${tr.arrivalTime}` : ''} Uhr
                          </div>
                          {tr.description && <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>{tr.description}</div>}
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                {/* Dokumente mit Viewer */}
                {event.documents && event.documents.length > 0 && (
                  <DocumentsViewer documents={event.documents} t={t} />
                )}

                {/* Fun-Zone Quiz */}
                {!sessionsOnly && event.quiz && event.quiz.length > 0 && (
                  <QuizPlayer
                    quiz={event.quiz}
                    t={t}
                    clusterSize={event.quizClusterSize}
                    initialAnswers={(() => {
                      // Zuvor gespeicherte Antworten aus der Teilnehmer-Registrierung laden
                      // (QuizAnswers ist JSON-String eines number[][]).
                      try {
                        const raw = registration.QuizAnswers;
                        if (!raw) return undefined;
                        const parsed = JSON.parse(raw);
                        return Array.isArray(parsed) ? parsed : undefined;
                      } catch { return undefined; }
                    })()}
                    onProgress={async (score: number, answers: number[][], isComplete: boolean) => {
                      // Nach jedem Cluster-Wechsel in die Subsite-Teilnehmerliste schreiben.
                      // isComplete=true setzt zusätzlich QuizCompletedAt (für Statistik-Filter).
                      if (!event.subsiteUrl) {
                        console.warn('[DEX] saveQuizProgress: event.subsiteUrl leer - Save uebersprungen', { eventId: event.id });
                        return;
                      }
                      try {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const ctx = (window as any).__dexSpfxContext;
                        if (!ctx) {
                          console.warn('[DEX] saveQuizProgress: __dexSpfxContext fehlt - Save uebersprungen');
                          return;
                        }
                        const { EventService } = await import('../services/EventService');
                        const svc = new EventService(ctx);
                        const ok = await svc.saveQuizProgress(event.subsiteUrl, registration.Id, score, answers, isComplete);
                        const answeredNow = answers.filter(a => Array.isArray(a) && a.length > 0).length;
                        console.warn(`[DEX] saveQuizProgress ok=${ok} regId=${registration.Id} score=${score} answered=${answeredNow}/${answers.length} complete=${isComplete} subsite=${event.subsiteUrl}`);
                        // VERIFY: re-read das Item und checke ob QuizAnswers tatsächlich in SP landete
                        try {
                          const verifyResp = await ctx.spHttpClient.get(
                            `${event.subsiteUrl}/_api/web/lists/getbytitle('Teilnehmer')/items(${registration.Id})?$select=QuizScore,QuizAnswers,QuizCompletedAt`,
                            (await import('@microsoft/sp-http')).SPHttpClient.configurations.v1
                          );
                          if (verifyResp.ok) {
                            const data = await verifyResp.json();
                            console.warn(`[DEX] saveQuizProgress VERIFY: QuizScore=${data.QuizScore} QuizAnswersLen=${(data.QuizAnswers || '').length} QuizCompletedAt=${data.QuizCompletedAt} QuizAnswersSample=${String(data.QuizAnswers || '').substring(0, 120)}`);
                          } else {
                            console.warn(`[DEX] saveQuizProgress VERIFY read failed: ${verifyResp.status}`);
                          }
                        } catch (vErr) {
                          console.warn('[DEX] saveQuizProgress VERIFY error:', vErr);
                        }
                        // Lokale myEvents-Liste nach erfolgreichem Save aktualisieren, damit die
                        // registration im Parent-State die frischen QuizScore/QuizAnswers hat —
                        // sonst sieht der User beim Wiedereintritt noch das alte (leere) Feld.
                        if (ok) {
                          setMyEvents(prev => prev.map(entry => {
                            if (entry.event.id !== event.id) return entry;
                            return {
                              ...entry,
                              registration: {
                                ...entry.registration,
                                QuizScore: score,
                                QuizAnswers: JSON.stringify(answers),
                                ...(isComplete ? { QuizCompletedAt: new Date().toISOString() } : {}),
                              },
                            };
                          }));
                        }
                      } catch (err) { console.warn('[DEX] saveQuizProgress failed:', err); }
                    }}
                  />
                )}

                {/* Sub-Events (Trainingssessions etc.) — nur wenn Event welche hat.
                    Seit v6.4: Sub-Events sind eigene DEX_Events-Items, werden über
                    childEventsOf(parentId) aus dem Context gezogen. */}
                {childEventsOf(event.id).length > 0 && (
                  <MyEventSubEvents
                    parentEvent={event}
                    childEvents={childEventsOf(event.id)}
                    registerForEvent={registerForEvent}
                    cancelRegistration={cancelRegistration}
                    getMyRegistration={getMyRegistration}
                    getAllRegistrations={getAllRegistrations}
                    updateMyRegistration={updateMyRegistration}
                    onMutated={loadMyRegistrations}
                  />
                )}

                {/* v11.0: Datei-Upload-Block — wird nur gerendert, wenn der
                    Organizer beim Event den Upload erlaubt hat und die
                    Anmeldung aktiv ist (nicht sessionsOnly oder abgemeldet). */}
                {event.allowAttendeeUpload && !sessionsOnly && (
                  <MyEventUpload
                    event={event}
                    list={listMyEventAttachments}
                    upload={uploadMyEventAttachment}
                    remove={deleteMyEventAttachment}
                  />
                )}

                {/* v19.0: Dokument-Custom-Felder — pro Feld ein Upload-Block,
                    damit der User die Datei auch nachträglich ergänzen/ersetzen
                    kann. */}
                {!sessionsOnly && (event.eventSpecificFields || []).filter(f => f.type === 'document').map(df => (
                  <MyEventDocField
                    key={df.id}
                    event={event}
                    field={df}
                    list={listFieldDocuments}
                    upload={uploadFieldDocument}
                    remove={deleteFieldDocument}
                  />
                ))}
                {/* Registriert am + Aktionen — im Sessions-Only-Modus ausblenden,
                    weil es keine echte Parent-Registrierung gibt. Sessions werden
                    über die Sub-Event-Sektion oben gemanagt. */}
                {!sessionsOnly && (
                  <div style={{ marginTop: 12, paddingTop: 10, borderTop: '1px solid var(--dex-gray-200)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                      {t('myevents.registeredon')}: {formatDate(registration.RegistrationDate)}
                    </span>
                    <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                      {cancellingId === event.id && !isCancelling && event.lastDeregisterDate && new Date(event.lastDeregisterDate) < new Date() && (
                        // v9.17: prominenter Late-Cancel-Hinweis — der User soll
                        // klar sehen, dass der Organizer durch die Abmeldung
                        // automatisch informiert wird (entscheidend wenn z.B.
                        // Hotel/Catering noch reagieren muss).
                        <span style={{
                          fontSize: '0.82rem', color: 'var(--dex-orange-dark, #b35a00)',
                          background: 'var(--dex-orange-light, #fff3e0)',
                          border: '1px solid var(--dex-orange, #ed8b00)',
                          padding: '6px 10px', borderRadius: 6,
                          display: 'block', marginBottom: 6, width: '100%',
                          fontWeight: 500,
                        }}>
                          {t('myevents.latecancel')}
                        </span>
                      )}
                      {/* v11.30: „Angaben bearbeiten"-Button wandert nach
                          oben neben die Angaben-Tags. Hier in der Aktions-
                          Zeile nur noch der Cancel-Edit-Button während
                          aktiver Bearbeitung — sonst leer. */}
                      {editingId === event.id && (
                        <button className="btn btn-secondary" style={{ fontSize: '0.85rem', padding: '8px 16px' }} onClick={() => setEditingId(null)}>
                          {t('general.cancel')}
                        </button>
                      )}
                      {/* v10.27: Gruppe wechseln bei Split-Capacity-Events.
                          Sichtbar nur wenn beide Kapazitäten > 0 sind und der
                          User aktiv angemeldet (nicht abgemeldet) ist. Ein
                          Klick öffnet einen window.confirm-Dialog mit klarem
                          Hinweis, dass der Wechsel evtl. auf die Warteliste
                          der Ziel-Gruppe führt, falls diese voll ist. */}
                      {(() => {
                        const dCap = event.durchstarterCapacity || 0;
                        const fCap = event.funstarterCapacity || 0;
                        if (dCap <= 0 || fCap <= 0) return null;
                        const labelA = (event.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
                        const labelB = (event.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
                        const currentType = registration.StarterType || registration.PreferredStarterType || '';
                        const currentLabel = currentType === 'Durchstarter' ? labelA : currentType === 'Funstarter' ? labelB : '?';
                        const targetType: 'Durchstarter' | 'Funstarter' = currentType === 'Durchstarter' ? 'Funstarter' : 'Durchstarter';
                        const targetLabel = targetType === 'Durchstarter' ? labelA : labelB;
                        return (
                          <button
                            className="btn btn-secondary"
                            style={{ fontSize: '0.85rem', padding: '8px 16px' }}
                            title={t('myevents.switchgroup.title') || `Aktuell in: ${currentLabel}`}
                            onClick={async () => {
                              const msg = `${t('myevents.switchgroup.confirm') || 'Gruppe wechseln zu'} „${targetLabel}"?\n\n` +
                                ((t('myevents.switchgroup.hint') || 'Falls die Ziel-Gruppe bereits voll ist, kommst du auf deren Warteliste und rückst nach, sobald ein Platz frei wird.'));
                              if (!(await confirmDialog(msg, { confirmLabel: isDe ? 'Wechseln' : 'Switch' }))) return;
                              const r = await switchSplitGroup(event.id, targetType);
                              if (!r.ok) {
                                showAlert(t('myevents.switchgroup.failed') || 'Gruppen-Wechsel fehlgeschlagen.', { variant: 'error' });
                                return;
                              }
                              const okMsg = r.full
                                ? `${t('myevents.switchgroup.waitlist') || 'Wechsel registriert — du stehst auf der Warteliste der Gruppe'} „${targetLabel}".`
                                : `${t('myevents.switchgroup.success') || 'Wechsel erfolgreich — du bist jetzt in Gruppe'} „${targetLabel}".`;
                              showAlert(okMsg, { variant: 'success' });
                              await loadMyRegistrations();
                            }}
                          >
                            {(t('myevents.switchgroup.btn') || 'Gruppe wechseln')} → {targetLabel}
                          </button>
                        );
                      })()}
                      {/* Abmelden-Button: prominent ausgelegt damit er auf der Karte
                          sofort gefunden wird (User-Feedback v9.8). 2-Klick-Confirm
                          bleibt — der erste Klick färbt rot und blendet den
                          "Doch behalten"-Button daneben ein.
                          v22.22: Bei bereits vergangenen Events entfällt der Button —
                          stattdessen ein grauer Hinweis (performCancel blockt
                          zusätzlich, auch für den Auto-Cancel-Deep-Link). */}
                      {isEventOver(event) ? (
                        <span style={{
                          fontSize: '0.8rem', color: 'var(--dex-gray-500)',
                          padding: '8px 12px', borderRadius: 8,
                          background: 'var(--dex-gray-50, #fafafa)',
                          border: '1px solid var(--dex-gray-200)',
                          lineHeight: 1.4,
                        }}>
                          {isDe
                            ? 'Dieses Event liegt in der Vergangenheit — eine Abmeldung ist nicht mehr möglich.'
                            : 'This event is in the past — cancelling is no longer possible.'}
                        </span>
                      ) : (
                        <>
                      <button
                        className={`btn dex-cancel-btn${cancellingId === event.id ? ' dex-cancel-btn--armed' : ''}`}
                        onClick={() => handleCancel(event.id)}
                        disabled={isCancelling}
                        style={{
                          fontSize: '0.95rem',
                          fontWeight: 600,
                          padding: '10px 20px',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: 8,
                          background: cancellingId === event.id ? 'var(--dex-red)' : '#fff',
                          color: cancellingId === event.id ? '#fff' : 'var(--dex-red)',
                          border: `2px solid var(--dex-red)`,
                          borderRadius: 8,
                          boxShadow: cancellingId === event.id ? '0 2px 8px rgba(218,41,28,0.3)' : 'none',
                          cursor: isCancelling ? 'not-allowed' : 'pointer',
                        }}
                      >
                        <X size={16} />
                        {cancellingId === event.id ? (isCancelling ? '...' : t('myevents.confirmcancel')) : t('myevents.cancel')}
                      </button>
                      {cancellingId === event.id && !isCancelling && (
                        <button className="btn btn-secondary" onClick={() => setCancellingId(null)} style={{ fontSize: '0.85rem', padding: '8px 16px' }}>{t('myevents.keepreg')}</button>
                      )}
                        </>
                      )}
                    </div>
                  </div>
                )}
              </div>
            );
          };
        const clusterHeadingStyle: React.CSSProperties = {
          margin: '0 0 12px', fontSize: '1.05rem',
          color: 'var(--dex-gray-700)', display: 'flex', alignItems: 'center', gap: 8,
        };
        return (
          <>
            {upcomingEntries.length > 0 && (
              <>
                <h3 style={clusterHeadingStyle}>
                  {isDe ? 'Aktive Events' : 'Active events'}
                  <span style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-400)' }}>({upcomingEntries.length})</span>
                </h3>
                <div className="my-events-list">{upcomingEntries.map(renderMyEventCard)}</div>
              </>
            )}
            {pastEntries.length > 0 && (
              <>
                <h3 style={{ ...clusterHeadingStyle, marginTop: upcomingEntries.length > 0 ? 28 : 0 }}>
                  {isDe ? 'Vergangene Events' : 'Past events'}
                  <span style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-400)' }}>({pastEntries.length})</span>
                </h3>
                <p style={{ margin: '-6px 0 12px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                  {isDe
                    ? 'Diese Events liegen in der Vergangenheit — eine Abmeldung ist hier nicht mehr möglich.'
                    : 'These events are in the past — cancelling is no longer possible here.'}
                </p>
                <div className="my-events-list">{pastEntries.map(renderMyEventCard)}</div>
              </>
            )}
          </>
        );
      })()}

      {cancelledEntries.length > 0 && (
        // v11.97: Cancelled-Liste einklappbar — Default eingeklappt, damit
        // die Liste nicht den ganzen Screen mit abgemeldeten Events flutet.
        <CancelledEventsCollapsible
          count={cancelledEntries.length}
          title={t('myevents.cancelledevents')}
          locale={t('myevents.cancelledon')}
          entries={cancelledEntries}
          formatDate={formatDate}
          statusLabel={t('status.cancelled')}
          hintText={locale === 'de'
            ? 'Diese Events hast du abgemeldet. Möchtest du wieder teilnehmen, melde dich im Anmelde-Bereich einfach neu an.'
            : 'You cancelled these events. If you want to attend again, simply register anew in the registration area.'}
          reRegisterLabel={locale === 'de' ? 'Zur Anmeldung' : 'Go to registration'}
          onReRegister={() => navigate('register')}
        />
      )}

      {/* v11.34: Cascade-Cancel-Modal — ersetzt das früher genutzte
          window.confirm. Drei klare Aktionen: alles abmelden, nur
          Hauptevent, Abbrechen. */}
      {/* v11.83: Add-Member-Modal (Team-Lead fügt nachträglich ein
          Mitglied hinzu). Layout-Logik analog zum cascadeDialog. */}
      {addMemberDialog && (
        <Modal
          open={true}
          onClose={closeAddMemberDialog}
          dismissable={!addMemberBusy}
          maxWidth={540}
          ariaLabel={isDe ? 'Mitglied zum Team hinzufügen' : 'Add member to team'}
        >
            <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Mitglied zum Team hinzufügen' : 'Add member to team'}
            </h3>
            {/* Pflicht-Hinweisbox — orange, analog zur Initial-Team-Anmeldung. */}
            <div style={{
              padding: '14px 16px',
              background: 'rgba(237,139,0,0.10)',
              border: '2px solid var(--dex-orange, #ed8b00)',
              borderRadius: 8,
              color: '#7a4a00',
              fontSize: '0.88rem',
              lineHeight: 1.5,
            }}>
              <div style={{ fontWeight: 700, marginBottom: 6 }}>
                {isDe ? 'Vorab Zustimmung einholen' : 'Get consent up front'}
              </div>
              <div style={{ marginBottom: 6 }}>
                {isDe
                  ? 'Die ausgewählte Person wird sofort und ohne weitere Rückfrage zum Team hinzugefügt. Sie bekommt automatisch:'
                  : 'The selected person is added to the team immediately, without further confirmation. They automatically receive:'}
              </div>
              <ul style={{ margin: '0 0 4px 18px', padding: 0 }}>
                <li>{isDe ? 'eine Anmeldebestätigung per Mail' : 'a confirmation email'}</li>
                <li>{isDe ? 'einen Outlook-Termin im Kalender' : 'an Outlook calendar invite'}</li>
                <li>{isDe ? 'den Event in „Meine Events"' : 'the event in „My Events"'}</li>
              </ul>
              <div style={{ marginTop: 4 }}>
                {isDe
                  ? <>Bitte stelle sicher, dass die Person ihrer Anmeldung <strong>vorher zugestimmt</strong> hat.</>
                  : <>Make sure the person has <strong>consented up front</strong>.</>}
              </div>
            </div>
            {/* People-Picker — simple Inline-Variante mit der searchUsers-API. */}
            <div>
              <label style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)', display: 'block', marginBottom: 4 }}>
                <span style={{ color: 'var(--dex-red)' }}>*</span> {isDe ? 'Person auswählen' : 'Pick a person'}
              </label>
              {addMemberPick ? (
                <div style={{
                  display: 'inline-flex', alignItems: 'center', gap: 10,
                  padding: '6px 10px 6px 6px',
                  border: '1px solid var(--dex-gray-200)',
                  borderRadius: 'var(--dex-radius)',
                  background: 'var(--dex-gray-50, #f7f7f7)',
                  maxWidth: '100%',
                }}>
                  <img
                    src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(addMemberPick.email)}&size=S`}
                    alt={addMemberPick.displayName}
                    onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                    style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                  />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: '0.9rem' }}>{addMemberPick.displayName}</div>
                    <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem' }}>{addMemberPick.email}</div>
                  </div>
                  <button
                    type="button"
                    onClick={() => { setAddMemberPick(null); setAddMemberQuery(''); setAddMemberResults([]); }}
                    title={isDe ? 'Auswahl entfernen' : 'Remove selection'}
                    style={{
                      background: 'var(--dex-gray-200)', border: 'none', color: 'var(--dex-gray-700)',
                      width: 22, height: 22, borderRadius: '50%', cursor: 'pointer',
                      fontSize: '0.9rem', lineHeight: 1,
                    }}
                  >×</button>
                </div>
              ) : (
                <div style={{ position: 'relative' }}>
                  <input
                    className="form-input"
                    value={addMemberQuery}
                    placeholder={isDe ? 'Name oder E-Mail eingeben…' : 'Type a name or email…'}
                    onChange={e => {
                      const val = e.target.value;
                      setAddMemberQuery(val);
                      if (addMemberQueryTimer.current) clearTimeout(addMemberQueryTimer.current);
                      if (val.length >= 2) {
                        addMemberQueryTimer.current = setTimeout(async () => {
                          setAddMemberSearching(true);
                          try {
                            const res = await searchUsers(val, addMemberIncludeIntl);
                            setAddMemberResults(res.map(r => ({ email: r.email, displayName: r.displayName })));
                          } catch { setAddMemberResults([]); }
                          setAddMemberSearching(false);
                        }, 300);
                      } else {
                        setAddMemberResults([]);
                      }
                    }}
                  />
                  <div style={{ marginTop: 2 }}>
                    <InternationalSearchToggle checked={addMemberIncludeIntl} onChange={setAddMemberIncludeIntl} isDe={isDe} />
                  </div>
                  {(addMemberResults.length > 0 || addMemberSearching) && (
                    <div style={{
                      position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 10,
                      background: '#fff', border: '1px solid var(--dex-gray-200)',
                      borderRadius: 6, marginTop: 4, maxHeight: 220, overflowY: 'auto',
                      boxShadow: '0 4px 12px rgba(0,0,0,0.12)',
                    }}>
                      {addMemberSearching && (
                        <div style={{ padding: 10, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
                          {isDe ? 'Suche…' : 'Searching…'}
                        </div>
                      )}
                      {addMemberResults.map(r => (
                        <button
                          key={r.email}
                          type="button"
                          onClick={() => { setAddMemberPick(r); setAddMemberResults([]); setAddMemberQuery(''); }}
                          style={{
                            display: 'flex', alignItems: 'center', gap: 8,
                            width: '100%', padding: '6px 10px', border: 'none',
                            background: '#fff', cursor: 'pointer', textAlign: 'left',
                          }}
                        >
                          <img
                            src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(r.email)}&size=S`}
                            alt={r.displayName}
                            onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                            style={{ width: 28, height: 28, borderRadius: '50%' }}
                          />
                          <div>
                            <div style={{ fontSize: '0.85rem', fontWeight: 600 }}>{r.displayName}</div>
                            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{r.email}</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
              <input
                type="checkbox"
                checked={addMemberConsent}
                onChange={e => setAddMemberConsent(e.target.checked)}
                style={{ marginTop: 3 }}
              />
              <span>
                <span style={{ color: 'var(--dex-red)', marginRight: 4 }}>*</span>
                {isDe
                  ? 'Ich bestätige, dass die ausgewählte Person ihrer Anmeldung zugestimmt hat.'
                  : 'I confirm that the selected person has consented to this registration.'}
              </span>
            </label>
            {addMemberError && (
              <div style={{ padding: 10, borderRadius: 6, background: 'rgba(220,38,38,0.10)', color: '#b91c1c', fontSize: '0.85rem' }}>
                {addMemberError}
              </div>
            )}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={closeAddMemberDialog}
                disabled={addMemberBusy}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={submitAddMember}
                disabled={!addMemberPick || !addMemberConsent || addMemberBusy}
              >
                {addMemberBusy
                  ? (isDe ? 'Wird hinzugefügt…' : 'Adding…')
                  : (isDe ? 'Hinzufügen' : 'Add')}
              </button>
            </div>
        </Modal>
      )}
      {/* v11.86: Manage-Team-Modal — Lead bearbeitet sein Team
          (Mitglieder abmelden). Pflicht-Confirm vor dem eigentlichen
          Cancel; der Cancel selbst läuft ueber cancelTeamMember im
          EventContext. */}
      {manageTeamDialog && (() => {
        const activeMembers = manageTeamMembers.filter(m => m.Status !== 'Abgemeldet');
        const sortedAll = [...manageTeamMembers].sort((a, b) => {
          const aLead = a.TeamLead ? 0 : 1;
          const bLead = b.TeamLead ? 0 : 1;
          if (aLead !== bLead) return aLead - bLead;
          const aTid = typeof a.TeilnehmerID === 'number' ? a.TeilnehmerID : Number.MAX_SAFE_INTEGER;
          const bTid = typeof b.TeilnehmerID === 'number' ? b.TeilnehmerID : Number.MAX_SAFE_INTEGER;
          if (aTid !== bTid) return aTid - bTid;
          return a.Id - b.Id;
        });
        return (
          <Modal
            open={true}
            onClose={closeManageTeamDialog}
            dismissable={manageTeamBusyId === null}
            maxWidth={620}
            ariaLabel={isDe ? 'Team verwalten' : 'Manage team'}
          >
              <div>
                <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
                  {isDe
                    ? `Team „${manageTeamDialog.teamName || 'Unbenannt'}" verwalten`
                    : `Manage team „${manageTeamDialog.teamName || 'Unnamed'}"`}
                </h3>
                <div style={{ marginTop: 4, fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? `Belegung: ${activeMembers.length}/${manageTeamDialog.teamSize}`
                    : `Occupancy: ${activeMembers.length}/${manageTeamDialog.teamSize}`}
                </div>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {sortedAll.map(m => {
                  const isCancelled = m.Status === 'Abgemeldet';
                  const isMemberLead = !!m.TeamLead && !isCancelled;
                  const isSelf = (m.ParticipantEmail || '').toLowerCase() === (currentUserEmail || '').toLowerCase();
                  const fullName = `${m.Vorname || ''} ${m.Nachname || ''}`.trim() || m.ParticipantEmail;
                  const loc = (m.Location || '').trim();
                  const busy = manageTeamBusyId === m.Id;
                  return (
                    <div
                      key={m.Id}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 12,
                        padding: '8px 10px', borderRadius: 6,
                        background: isCancelled ? 'rgba(0,0,0,0.04)' : '#fff',
                        border: '1px solid var(--dex-gray-200)',
                        opacity: isCancelled ? 0.55 : 1,
                      }}
                    >
                      <div style={{ width: 40, height: 40, borderRadius: '50%', flexShrink: 0, position: 'relative' }}>
                        <img
                          src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(m.ParticipantEmail)}&size=L`}
                          alt={fullName}
                          onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                          style={{
                            width: 40, height: 40, borderRadius: '50%',
                            objectFit: 'cover', background: 'var(--dex-gray-100)',
                            transition: 'transform 160ms ease, box-shadow 160ms ease',
                            transformOrigin: 'left center',
                            /* v11.94: kein zoom-in-Cursor */
                          }}
                          onMouseEnter={e => {
                            if (isCancelled) return;
                            const img = e.currentTarget as HTMLImageElement;
                            img.style.transform = 'scale(2.4)';
                            img.style.boxShadow = '0 8px 24px rgba(0,0,0,0.35)';
                            img.style.zIndex = '50';
                            img.style.position = 'relative';
                          }}
                          onMouseLeave={e => {
                            const img = e.currentTarget as HTMLImageElement;
                            img.style.transform = 'scale(1)';
                            img.style.boxShadow = 'none';
                            img.style.zIndex = '';
                            img.style.position = '';
                          }}
                        />
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)', fontSize: '0.9rem' }}>
                          {fullName}
                        </div>
                        <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>
                          {m.ParticipantEmail}{loc ? ` · ${loc}` : ''}
                        </div>
                      </div>
                      {isCancelled ? (
                        <span style={{
                          padding: '2px 8px', borderRadius: 999,
                          background: 'var(--dex-gray-300, #c8c8c8)', color: '#fff',
                          fontSize: '0.68rem', fontWeight: 600, flexShrink: 0,
                        }}>
                          {isDe ? 'abgemeldet' : 'cancelled'}
                        </span>
                      ) : isMemberLead ? (
                        <span style={{
                          padding: '2px 8px', borderRadius: 999,
                          background: 'var(--dex-green, #86bc25)', color: '#fff',
                          fontSize: '0.68rem', fontWeight: 600, flexShrink: 0,
                        }}>
                          Lead
                        </span>
                      ) : null}
                      {/* Trash-Button — nur für aktive Nicht-Lead-Nicht-Self-Mitglieder. */}
                      {!isCancelled && !isMemberLead && !isSelf && (
                        <button
                          type="button"
                          onClick={() => setManageTeamConfirm(m)}
                          disabled={busy || manageTeamBusyId !== null}
                          title={isDe ? 'Diese Person aus dem Team abmelden' : 'Remove this person from the team'}
                          style={{
                            background: 'var(--dex-red, #d62828)', color: '#fff', border: 'none',
                            width: 32, height: 32, borderRadius: 6, cursor: 'pointer',
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            flexShrink: 0, opacity: busy ? 0.6 : 1,
                          }}
                        >
                          <Icon iconName="Delete" style={{ fontSize: 14 }} />
                        </button>
                      )}
                    </div>
                  );
                })}
              </div>
              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.45 }}>
                {isDe
                  ? <>{'Du als Team-Lead kannst dich nicht selbst über diesen Pfad abmelden — nutze dafür den normalen „Abmelden"-Button auf deiner Event-Karte. Dort übernimmt die App automatisch den Lead-Wechsel an das früheste verbleibende Team-Mitglied.'}</>
                  : <>{'As team lead you cannot cancel yourself via this dialog — use the normal „Cancel" button on your event card instead. The app then automatically transfers the lead role to the earliest remaining team member.'}</>}
              </div>
              <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={closeManageTeamDialog}
                  disabled={manageTeamBusyId !== null}
                >
                  {isDe ? 'Schließen' : 'Close'}
                </button>
              </div>
            {/* Confirm-Modal (zweite Ebene) — sitzt auf demselben z-index-Layer. */}
            {manageTeamConfirm && (() => {
              const cm = manageTeamConfirm;
              const fullName = `${cm.Vorname || ''} ${cm.Nachname || ''}`.trim() || cm.ParticipantEmail;
              const busy = manageTeamBusyId === cm.Id;
              return (
                <Modal
                  open={true}
                  onClose={() => setManageTeamConfirm(null)}
                  dismissable={!busy}
                  maxWidth={480}
                  ariaLabel={isDe ? 'Person aus dem Team abmelden?' : 'Cancel this team member?'}
                >
                    <h3 style={{ margin: 0, fontSize: '1.05rem', color: 'var(--dex-gray-800)' }}>
                      {isDe ? 'Person aus dem Team abmelden?' : 'Cancel this team member?'}
                    </h3>
                    <div style={{ fontSize: '0.9rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                      {isDe ? (
                        <>
                          <p style={{ marginTop: 0 }}>
                            <strong>{fullName}</strong> wird vom Event abgemeldet. Die Person bekommt eine Abmelde-Bestätigungs-Mail und (falls Outlook aktiv) eine Termin-Absage.
                          </p>
                          <p style={{ marginBottom: 0, color: 'var(--dex-gray-600)' }}>
                            {'Hinweis: die Person könnte sich auch selbst über „Meine Events" abmelden — du machst das nur stellvertretend.'}
                          </p>
                        </>
                      ) : (
                        <>
                          <p style={{ marginTop: 0 }}>
                            <strong>{fullName}</strong> will be cancelled from the event. They will receive a cancellation confirmation email and (if Outlook is active) a calendar removal.
                          </p>
                          <p style={{ marginBottom: 0, color: 'var(--dex-gray-600)' }}>
                            {'Note: this person could also cancel themselves via „My Events" — you are doing it on their behalf.'}
                          </p>
                        </>
                      )}
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        onClick={() => setManageTeamConfirm(null)}
                        disabled={busy}
                      >
                        {isDe ? 'Abbrechen' : 'Cancel'}
                      </button>
                      <button
                        type="button"
                        className="btn btn-primary"
                        onClick={() => performManageTeamCancel(cm)}
                        disabled={busy}
                        style={{ background: 'var(--dex-red, #d62828)', borderColor: 'var(--dex-red, #d62828)' }}
                      >
                        {busy
                          ? (isDe ? 'Wird abgemeldet…' : 'Cancelling…')
                          : (isDe ? 'Person abmelden' : 'Cancel member')}
                      </button>
                    </div>
                </Modal>
              );
            })()}
          </Modal>
        );
      })()}
      {cascadeDialog && (() => {
        const dlg = cascadeDialog;
        const choose = (c: 'cascade' | 'parent-only' | 'abort'): void => dlg.resolve(c);
        const isSec = !!dlg.isSectionedEvent;
        // v14.7: Bei „Sectioned"-Events (requireSubEventSelection an) sprechen
        // wir nicht von „Sub-Event" sondern von „Event-Section". Außerdem
        // wird die Option „nur Hauptevent abmelden, Sub-Events behalten"
        // weggelassen — der Teilnehmer war ja gar nicht „nur Hauptevent"
        // angemeldet (die Pflichtwahl hat das verhindert). Stattdessen
        // Hinweis aufs Anmelde-Bearbeiten für einzelne Sections.
        const title = isSec
          ? (isDe ? 'Komplett vom Event abmelden?' : 'Cancel registration entirely?')
          : (isDe ? 'Auch von Sub-Events abmelden?' : 'Cancel sub-events too?');
        return (
          <Modal
            open={true}
            onClose={() => choose('abort')}
            maxWidth={520}
            ariaLabel={title}
          >
              <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
                {title}
              </h3>
              <p style={{ margin: 0, fontSize: '0.9rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                {isSec
                  ? (isDe
                      ? <>Du bist für <strong>{dlg.subEvents.length}</strong> Event-Section{dlg.subEvents.length === 1 ? '' : 's'} von <strong>&bdquo;{dlg.parentTitle}&ldquo;</strong> angemeldet:</>
                      : <>You are registered for <strong>{dlg.subEvents.length}</strong> event section{dlg.subEvents.length === 1 ? '' : 's'} of <strong>&bdquo;{dlg.parentTitle}&ldquo;</strong>:</>)
                  : (isDe
                      ? <>Du bist für <strong>{dlg.subEvents.length}</strong> Sub-Event{dlg.subEvents.length === 1 ? '' : 's'} von <strong>&bdquo;{dlg.parentTitle}&ldquo;</strong> angemeldet:</>
                      : <>You are registered for <strong>{dlg.subEvents.length}</strong> sub-event{dlg.subEvents.length === 1 ? '' : 's'} of <strong>&bdquo;{dlg.parentTitle}&ldquo;</strong>:</>)
                }
              </p>
              <ul style={{ margin: 0, paddingLeft: 20, fontSize: '0.85rem', color: 'var(--dex-gray-700)', maxHeight: 200, overflowY: 'auto' }}>
                {/* v15.8: pro Sub-Event Titel + Datum + Ort listen, damit
                    der User auf einen Blick sieht was er da abmeldet. */}
                {dlg.subEvents.map(s => {
                  const dateStr = s.startDate
                    ? new Date(s.startDate).toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })
                    : '';
                  const subParts = [dateStr, s.location].filter(Boolean).join(' · ');
                  return (
                    <li key={s.id} style={{ marginBottom: 4 }}>
                      <div style={{ fontWeight: 600 }}>{s.title}</div>
                      {subParts && (
                        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>{subParts}</div>
                      )}
                    </li>
                  );
                })}
              </ul>
              {isSec && (
                <div style={{
                  padding: '10px 12px', borderRadius: 8,
                  background: 'rgba(0,118,168,0.08)',
                  border: '1px solid var(--dex-blue, #0076a8)',
                  fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.45,
                }}>
                  {isDe
                    ? <>Du willst nur <strong>einzelne Sections</strong> abwählen (z.B. &bdquo;Dinner absagen, aber Meeting bleibt&ldquo;)? Dann breche hier ab und nutze stattdessen <strong>&bdquo;Anmeldung bearbeiten&ldquo;</strong> auf der Event-Karte — dort kannst du gezielt einzelne Sections an- oder abwählen, ohne die ganze Teilnahme zu stornieren.</>
                    : <>Want to drop only <strong>individual sections</strong> (e.g. &bdquo;cancel dinner, keep meeting&ldquo;)? Then abort here and use <strong>&bdquo;Edit registration&ldquo;</strong> on the event card instead — there you can pick / deselect individual sections without cancelling the whole registration.</>}
                </div>
              )}
              <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.4 }}>
                {isDe
                  ? 'Wähle wie weiter:'
                  : 'How do you want to proceed?'}
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                <button
                  className="btn btn-primary"
                  onClick={() => choose('cascade')}
                  style={{ fontSize: '0.9rem', padding: '10px 16px' }}
                >
                  {isSec
                    ? (isDe
                        ? `Ja, komplett abmelden (Event + ${dlg.subEvents.length} Section${dlg.subEvents.length === 1 ? '' : 's'})`
                        : `Yes, cancel entirely (event + ${dlg.subEvents.length} section${dlg.subEvents.length === 1 ? '' : 's'})`)
                    : (isDe
                        ? `Alles abmelden (Hauptevent + ${dlg.subEvents.length} Sub-Event${dlg.subEvents.length === 1 ? '' : 's'})`
                        : `Cancel everything (main event + ${dlg.subEvents.length} sub-event${dlg.subEvents.length === 1 ? '' : 's'})`)}
                </button>
                {!isSec && (
                  <button
                    className="btn btn-secondary"
                    onClick={() => choose('parent-only')}
                    style={{ fontSize: '0.9rem', padding: '10px 16px' }}
                  >
                    {isDe
                      ? 'Nur Hauptevent abmelden, Sub-Events behalten'
                      : 'Cancel main event only, keep sub-events'}
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => choose('abort')}
                  style={{
                    background: 'transparent', border: 'none',
                    color: 'var(--dex-gray-500)', cursor: 'pointer',
                    fontSize: '0.85rem', textDecoration: 'underline',
                    padding: '6px 16px',
                  }}
                >
                  {isDe ? 'Abbrechen — nichts abmelden' : 'Cancel — keep everything'}
                </button>
              </div>
          </Modal>
        );
      })()}

      {/* v20.7: Modal mit dem persönlichen Check-in-QR-Code (wie in der
          QR-Mail: großer Code + Name + Teilnehmer-Nr. als Fallback für den
          manuellen Check-in). */}
      {myQrModal && (
        <Modal
          open={true}
          onClose={() => setMyQrModal(null)}
          maxWidth={420}
          ariaLabel={isDe ? 'Mein QR-Code' : 'My QR code'}
        >
          <h3 style={{ margin: 0, fontSize: '1.05rem', textAlign: 'center' }}>
            {isDe ? 'Mein Check-in-QR-Code' : 'My check-in QR code'}
          </h3>
          <p style={{ margin: 0, fontSize: '0.82rem', color: 'var(--dex-gray-500)', textAlign: 'center' }}>
            {myQrModal.eventTitle}
          </p>
          <div style={{ textAlign: 'center' }}>
            <img
              src={myQrModal.dataUrl}
              alt="QR-Code"
              style={{ width: 280, maxWidth: '90%', height: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: 10, background: '#fff' }}
            />
          </div>
          <p style={{ margin: 0, textAlign: 'center', fontWeight: 700, fontSize: '0.95rem' }}>
            {myQrModal.name}
            {myQrModal.tid ? <span style={{ fontWeight: 400, color: 'var(--dex-gray-500)' }}> · Nr. {myQrModal.tid}</span> : null}
          </p>
          <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)', textAlign: 'center', lineHeight: 1.5 }}>
            {isDe
              ? 'Zeig diesen Code am Event-Tag dem Check-in-Team — er ist derselbe wie in deiner QR-Mail. Falls der Scan nicht klappt, reicht dein Name.'
              : 'Show this code to the check-in team on event day — it is the same as in your QR email. If the scan fails, your name is enough.'}
          </p>
          <div style={{ textAlign: 'right' }}>
            <button className="btn btn-secondary" onClick={() => setMyQrModal(null)} style={{ fontSize: '0.85rem' }}>
              {isDe ? 'Schließen' : 'Close'}
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ==================== Datei-Upload ("Meine Events") ====================
// v11.0: Wenn der Organizer beim Event den Toggle „Teilnehmer-Upload
// erlauben" gesetzt hat, sieht der Teilnehmer hier einen Upload-Block.
// v11.97: Einklappbare Liste der abgemeldeten Events. Default eingeklappt,
// damit lange Cancelled-Listen die Hauptliste nicht überlagern. Header
// zeigt Count + Chevron, Klick togglt die Liste.
function CancelledEventsCollapsible(props: {
  count: number;
  title: string;
  locale: string;
  entries: Array<{ event: { id: string; title: string }; registration: { CancellationDate?: string } }>;
  formatDate: (iso: string) => string;
  statusLabel: string;
  // v18.62: Hinweis + Button, dass man sich für eine erneute Teilnahme
  // im Anmelde-Bereich neu registrieren muss (Abmeldung ist endgültig).
  hintText: string;
  reRegisterLabel: string;
  onReRegister: () => void;
}): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  return (
    <div className="mt-24">
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        style={{
          background: 'none', border: 'none', padding: 0, cursor: 'pointer',
          display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12,
          color: 'var(--dex-gray-600)', fontSize: '1rem', fontWeight: 600,
        }}
        aria-expanded={open}
      >
        <span style={{
          display: 'inline-block', width: 14, textAlign: 'center',
          transform: open ? 'rotate(90deg)' : 'rotate(0deg)',
          transition: 'transform 0.15s ease',
          fontSize: '0.85rem',
        }}>▶</span>
        <span>{props.title} ({props.count})</span>
      </button>
      {open && (
        <div className="my-events-list">
          <div style={{
            fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginBottom: 12,
            padding: '10px 14px', borderRadius: 8,
            background: 'var(--dex-gray-50, #fafafa)', border: '1px solid var(--dex-gray-200)',
            display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap',
          }}>
            <span style={{ flex: '1 1 auto' }}>{props.hintText}</span>
            <button
              type="button"
              className="btn btn-outline"
              style={{ fontSize: '0.78rem', padding: '5px 14px', whiteSpace: 'nowrap' }}
              onClick={props.onReRegister}
            >{props.reRegisterLabel}</button>
          </div>
          {props.entries.map(({ event, registration }) => (
            <div
              key={event.id}
              className="card my-event-card"
              style={{
                opacity: 0.6,
                display: 'flex',
                alignItems: 'center',
                gap: 12,
                padding: '8px 16px',
                flexWrap: 'wrap',
              }}
            >
              <strong style={{ flex: '1 1 auto', fontSize: '0.95rem', margin: 0 }}>{event.title}</strong>
              <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                {props.locale}: {registration.CancellationDate ? props.formatDate(registration.CancellationDate) : '-'}
              </span>
              <span className="badge badge-red">{props.statusLabel}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// Dateien werden als SP-Item-Attachment direkt an die eigene Teilnehmer-
// zeile gehängt — der Admin sieht sie im Admin Center.
function MyEventUpload(props: {
  event: DeloitteEvent;
  list: (eventId: string) => Promise<Array<{ fileName: string; serverRelativeUrl: string }>>;
  upload: (eventId: string, file: File) => Promise<boolean>;
  remove: (eventId: string, fileName: string) => Promise<boolean>;
}): React.ReactElement | null {
  const { event } = props;
  const isDe = (event.emailLanguage || 'EN').toUpperCase() === 'DE';
  // v20.4: App-Modal statt window.confirm.
  const { confirmDialog } = useDialog();
  const [files, setFiles] = React.useState<Array<{ fileName: string; serverRelativeUrl: string }>>([]);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>('');

  const refresh = React.useCallback(async () => {
    try {
      const list = await props.list(event.id);
      setFiles(list);
    } catch { /* ignore */ }
  }, [event.id]);

  React.useEffect(() => { refresh().catch(() => { /* */ }); }, [refresh]);

  if (!event.allowAttendeeUpload) return null;

  const label = (event.attendeeUploadLabel || '').trim() || (isDe ? 'Dokumenten-Upload' : 'Document upload');
  const hint = (event.attendeeUploadHint || '').trim();

  const onPick = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!f) return;
    // Soft-Limit 10 MB — SharePoint Item-Attachments sind technisch bis 250 MB,
    // aber große PDFs sprengen die App-UX und Power-Automate-Workflows.
    if (f.size > 10 * 1024 * 1024) {
      setError(isDe ? 'Datei ist größer als 10 MB. Bitte komprimieren oder kleinere Variante hochladen.' : 'File is larger than 10 MB. Please compress or upload a smaller version.');
      return;
    }
    setError('');
    setBusy(true);
    try {
      const ok = await props.upload(event.id, f);
      if (!ok) setError(isDe ? 'Upload fehlgeschlagen.' : 'Upload failed.');
      await refresh();
    } finally {
      setBusy(false);
    }
  };

  const onDelete = async (fileName: string): Promise<void> => {
    if (!(await confirmDialog(isDe ? `Datei „${fileName}" wirklich löschen?` : `Really delete file „${fileName}"?`, { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' }))) return;
    setBusy(true);
    try {
      await props.remove(event.id, fileName);
      await refresh();
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Icon iconName="Attach" style={{ fontSize: 14 }} />
        {label}
      </div>
      {hint && (
        <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginBottom: 10, lineHeight: 1.45 }}>
          {hint}
        </div>
      )}
      {files.length > 0 && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 10 }}>
          {files.map(f => (
            <div key={f.fileName} style={{
              display: 'flex', alignItems: 'center', gap: 10,
              padding: '6px 10px', borderRadius: 6,
              background: 'rgba(134,188,37,0.08)',
              border: '1px solid rgba(134,188,37,0.30)',
              fontSize: '0.82rem',
            }}>
              <Icon iconName="PDF" style={{ fontSize: 16, color: 'var(--dex-red, #c00)' }} />
              <a
                href={f.serverRelativeUrl}
                target="_blank"
                rel="noopener noreferrer"
                style={{ flex: 1, color: 'var(--dex-gray-800)', textDecoration: 'none', wordBreak: 'break-all' }}
              >
                {f.fileName}
              </a>
              <button
                type="button"
                className="btn btn-secondary"
                style={{ fontSize: '0.72rem', padding: '2px 10px' }}
                disabled={busy}
                onClick={() => onDelete(f.fileName)}
                title={isDe ? 'Löschen' : 'Delete'}
              >
                ✕
              </button>
            </div>
          ))}
        </div>
      )}
      <label className="btn btn-outline" style={{ fontSize: '0.82rem', padding: '6px 14px', cursor: busy ? 'wait' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Icon iconName="Upload" style={{ fontSize: 14 }} />
        {busy ? (isDe ? 'Wird übertragen…' : 'Uploading…') : files.length > 0 ? (isDe ? 'Weitere Datei hochladen' : 'Upload another file') : (isDe ? 'Datei hochladen' : 'Upload file')}
        <input
          type="file"
          accept="application/pdf,image/*,.doc,.docx"
          style={{ display: 'none' }}
          onChange={onPick}
          disabled={busy}
        />
      </label>
      {error && (
        <div style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-red, #c00)' }}>{error}</div>
      )}
    </div>
  );
}

// v19.0: Upload-Block für EIN Dokument-Custom-Feld. Wird pro Dokument-Feld des
// Events gerendert, damit der Teilnehmer die Datei auch nachträglich
// ergänzen/ersetzen/löschen kann. Datei = SP-Item-Attachment (pro Feld über
// Dateinamen-Präfix zugeordnet) — der Organizer sieht sie im Admin Center.
function MyEventDocField(props: {
  event: DeloitteEvent;
  field: { id: string; label: string; required?: boolean };
  list: (eventId: string, fieldId: string, participantEmail?: string) => Promise<Array<{ fileName: string; serverRelativeUrl: string; displayName: string }>>;
  upload: (eventId: string, fieldId: string, file: File, participantEmail?: string) => Promise<boolean>;
  remove: (eventId: string, fileName: string, participantEmail?: string) => Promise<boolean>;
}): React.ReactElement {
  const { event, field } = props;
  const isDe = (event.emailLanguage || 'EN').toUpperCase() === 'DE';
  // v20.4: App-Modal statt window.confirm.
  const { confirmDialog } = useDialog();
  const [files, setFiles] = React.useState<Array<{ fileName: string; serverRelativeUrl: string; displayName: string }>>([]);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState('');
  const refresh = React.useCallback(async () => {
    try { setFiles(await props.list(event.id, field.id)); } catch { /* ignore */ }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [event.id, field.id]);
  React.useEffect(() => { refresh().catch(() => { /* */ }); }, [refresh]);
  const onPick = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!f) return;
    if (f.size > 10 * 1024 * 1024) { setError(isDe ? 'Datei ist größer als 10 MB. Bitte kleinere Variante hochladen.' : 'File is larger than 10 MB. Please upload a smaller version.'); return; }
    setError(''); setBusy(true);
    try {
      const ok = await props.upload(event.id, field.id, f);
      if (!ok) setError(isDe ? 'Upload fehlgeschlagen.' : 'Upload failed.');
      await refresh();
    } finally { setBusy(false); }
  };
  const onDelete = async (fileName: string, displayName: string): Promise<void> => {
    // eslint-disable-next-line no-alert
    if (!(await confirmDialog(isDe ? `Datei „${displayName}" wirklich löschen?` : `Really delete file „${displayName}"?`, { danger: true, confirmLabel: isDe ? 'Löschen' : 'Delete' }))) return;
    setBusy(true);
    try { await props.remove(event.id, fileName); await refresh(); } finally { setBusy(false); }
  };
  return (
    <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 8, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Icon iconName="Attach" style={{ fontSize: 14 }} />
        {field.label}{field.required ? ' *' : ''}
      </div>
      {files.length > 0 && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 10 }}>
          {files.map(f => (
            <div key={f.fileName} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '6px 10px', borderRadius: 6, background: 'rgba(134,188,37,0.08)', border: '1px solid rgba(134,188,37,0.30)', fontSize: '0.82rem' }}>
              <Icon iconName="Page" style={{ fontSize: 16, color: 'var(--dex-green-dark, #4a7c1f)' }} />
              <a href={f.serverRelativeUrl} target="_blank" rel="noopener noreferrer" style={{ flex: 1, color: 'var(--dex-gray-800)', textDecoration: 'none', wordBreak: 'break-all' }}>{f.displayName}</a>
              <button type="button" className="btn btn-secondary" style={{ fontSize: '0.72rem', padding: '2px 10px' }} disabled={busy} onClick={() => onDelete(f.fileName, f.displayName)} title={isDe ? 'Löschen' : 'Delete'}>✕</button>
            </div>
          ))}
        </div>
      )}
      <label className="btn btn-outline" style={{ fontSize: '0.82rem', padding: '6px 14px', cursor: busy ? 'wait' : 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Icon iconName="Upload" style={{ fontSize: 14 }} />
        {busy ? (isDe ? 'Wird übertragen…' : 'Uploading…') : files.length > 0 ? (isDe ? 'Weitere Datei hochladen' : 'Upload another file') : (isDe ? 'Datei hochladen (PDF/Bild)' : 'Upload file (PDF/image)')}
        <input type="file" accept="application/pdf,image/png,image/jpeg,.pdf,.png,.jpg,.jpeg" style={{ display: 'none' }} onChange={onPick} disabled={busy} />
      </label>
      {error && <div style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-red, #c00)' }}>{error}</div>}
    </div>
  );
}

// ==================== Sub-Events im "My Events"-Tab ====================
// Seit v6.4: Sub-Events sind eigene DEX_Events-Items (childEventsOf(parentId)).
// Anmeldung/Abmeldung läuft über den normalen registerForEvent/cancelRegistration-
// Pfad — identisch zu einem Top-Level-Event. Eigene Teilnehmerliste pro Sub-Event.
// v10.27: Plus Anzeige + Bearbeitung der Sub-Event-spezifischen Antworten
// (eventSpecificData) — analog zum Edit-Modus auf der Hauptevent-Karte.
function MyEventSubEvents(props: {
  parentEvent: DeloitteEvent;
  childEvents: DeloitteEvent[];
  registerForEvent: (eventId: string, customData: Record<string, string>) => Promise<{ ok: boolean; status: 'Angemeldet' | 'Warteliste' }>;
  cancelRegistration: (eventId: string, opts?: { suppressNotifications?: boolean }) => Promise<boolean>;
  getMyRegistration: (eventId: string) => Promise<SPRegistration | null>;
  getAllRegistrations: (eventId: string) => Promise<SPRegistration[]>;
  updateMyRegistration: (eventId: string, customData: Record<string, string>) => Promise<boolean>;
  onMutated: () => Promise<void>;
}): React.ReactElement | null {
  // v15.13: isDe MUSS aus dem UI-Locale kommen, nicht aus event.emailLanguage —
  // letzteres ist die Sprache der Bestätigungsmails (organizer-konfiguriert)
  // und sagt nichts darüber aus, in welcher Sprache der User die App sieht.
  const { locale: __uiLocale } = useLanguage();
  const isDe = __uiLocale === 'de';
  // v22.13: Sub-Events auch in „Meine Events" nach ihrer EIGENEN Sichtbarkeit
  // filtern (gleiche Logik wie Anmeldeseite/Event-Liste, Fix v22.10). Eigene
  // aktive Anmeldungen bleiben IMMER sichtbar (verwalten/abmelden); Organizer
  // des Events und Admins sehen weiterhin alles.
  const { currentUser, groupEmails } = useCurrentUser();
  const { isAdmin } = useRoles();
  const [busyId, setBusyId] = React.useState<string | null>(null);
  // v15.21: Voll-Bild-Progress-Modal beim (Peer-)Cancel + Anmeldung. Vorher
  // war nur die einzelne Karte mit „…" markiert — bei Peer-Cancels haben die
  // peers visuell nicht reagiert, weil busyId nur die EINE id hält.
  const [processingMessage, setProcessingMessage] = React.useState<string>('');
  const [registeredSet, setRegisteredSet] = React.useState<Set<string>>(new Set());
  const [counts, setCounts] = React.useState<Record<string, number>>({});
  // v10.27: pro Sub-Event die geparsten Custom-Field-Antworten aus
  // myReg.CustomData. Wird nur für aktive Registrierungen befüllt.
  const [seData, setSeData] = React.useState<Record<string, Record<string, string>>>({});
  // v10.27: aktuell ge-editiertes Sub-Event + Draft der Werte. Beim Save
  // gehen die Werte ueber updateMyRegistration in die Subsite zurück.
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [editDraft, setEditDraft] = React.useState<Record<string, string>>({});
  const [savingEdit, setSavingEdit] = React.useState(false);
  // v11.34: Cascade-Cancel-Dialog für Sub-Event-Cancel — fragt ob auch
  // die anderen aktiven Sub-Events des gleichen Parents abgemeldet
  // werden sollen (Peer-Cancel).
  // v15.8: Peer-Cancel-Modal mit Checkbox-Liste statt drei-Buttons —
  // User hakt direkt an, welche zusätzlichen Sub-Events er mitabmelden
  // will (Target ist immer fix abgemeldet, weil er ja explizit den
  // Cancel-Button geklickt hat). resolve liefert das Set der peer-IDs
  // zurück oder 'abort' wenn der Modal-User komplett abbricht.
  const [peerCancelDialog, setPeerCancelDialog] = React.useState<{
    targetTitle: string;
    peers: { id: string; title: string; startDate?: string; location?: string }[];
    resolve: (_choice: { peerIds: string[] } | 'abort') => void;
  } | null>(null);

  const refresh = React.useCallback(async (): Promise<void> => {
    try {
      const regs = await Promise.all(props.childEvents.map(async (ce) => {
        const reg = await props.getMyRegistration(ce.id);
        const isActive = !!reg && reg.Status !== 'Abgemeldet';
        // v10.27: Custom-Field-Werte aus reg.CustomData (JSON) parsen.
        let data: Record<string, string> = {};
        if (isActive && reg && reg.CustomData) {
          try { data = JSON.parse(reg.CustomData) as Record<string, string>; } catch { /* fall back to empty */ }
        }
        return { id: ce.id, isActive, data };
      }));
      const s = new Set<string>();
      const dataMap: Record<string, Record<string, string>> = {};
      for (const r of regs) {
        if (r.isActive) {
          s.add(r.id);
          dataMap[r.id] = r.data;
        }
      }
      setRegisteredSet(s);
      setSeData(dataMap);

      const countPairs = await Promise.all(props.childEvents.map(async (ce) => {
        const all = await props.getAllRegistrations(ce.id);
        const active = (all || []).filter(r => {
          const st = r.Status || '';
          return st === 'Angemeldet' || st === 'QR versendet' || st === 'Eingecheckt';
        }).length;
        return { id: ce.id, count: active };
      }));
      const map: Record<string, number> = {};
      for (const p of countPairs) map[p.id] = p.count;
      setCounts(map);
    } catch { /* ignore */ }
  }, [props.childEvents.map(c => c.id).join(',')]);

  React.useEffect(() => { refresh().catch(() => { /* ignore */ }); }, [refresh]);

  const fmt = (iso: string): string => {
    if (!iso) return '';
    try {
      const d = new Date(iso);
      if (isNaN(d.getTime())) return iso;
      return d.toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    } catch { return iso; }
  };

  const handleToggle = async (childEventId: string, currentlyRegistered: boolean): Promise<void> => {
    // v11.34: Beim Cancel eines Sub-Events fragen, ob die anderen aktiven
    // Sub-Events des gleichen Parents auch abgemeldet werden sollen
    // (Peer-Cancel-Cascade) — gleicher Pattern wie der Parent-Cancel.
    let peerIdsToCancel: string[] = [];
    if (currentlyRegistered) {
      const peers = props.childEvents
        .filter(ce => ce.id !== childEventId && registeredSet.has(ce.id))
        .map(ce => ({
          id: ce.id,
          title: ce.title || (isDe ? 'Sub-Event' : 'Sub-event'),
          startDate: ce.startDate || '',
          location: ce.location || '',
        }));
      if (peers.length > 0) {
        const target = props.childEvents.find(ce => ce.id === childEventId);
        const targetTitle = (target && target.title) || (isDe ? 'dieses Sub-Event' : 'this sub-event');
        const choice = await new Promise<{ peerIds: string[] } | 'abort'>(resolve => {
          setPeerCancelDialog({ targetTitle, peers, resolve });
        });
        setPeerCancelDialog(null);
        if (choice === 'abort') return;
        peerIdsToCancel = choice.peerIds;
      }
    }
    setBusyId(childEventId);
    // v15.21: Voll-Bild-Progress mit Beschreibung — bei Peer-Cancels
    // zählen wir die Fortschritte fortlaufend mit, damit der User sieht
    // dass auch die peers verarbeitet werden.
    const totalSteps = (currentlyRegistered ? 1 + peerIdsToCancel.length : 1);
    const initialMsg = currentlyRegistered
      ? (isDe
          ? `Abmeldung wird verarbeitet… (1/${totalSteps})`
          : `Cancellation in progress… (1/${totalSteps})`)
      : (isDe ? 'Anmeldung wird verarbeitet…' : 'Registration in progress…');
    setProcessingMessage(initialMsg);
    try {
      if (currentlyRegistered) {
        await props.cancelRegistration(childEventId);
        let done = 1;
        for (const peerId of peerIdsToCancel) {
          done++;
          setProcessingMessage(isDe
            ? `Abmeldung wird verarbeitet… (${done}/${totalSteps})`
            : `Cancellation in progress… (${done}/${totalSteps})`);
          try { await props.cancelRegistration(peerId); }
          catch (err) { console.warn('[DEX] peer-cancel failed:', peerId, err); }
        }
        // v15.26: Im subEventsOnlyMode war das Hauptevent als Schatten-
        // Registrierung angelegt. Wenn die letzte aktive Sub-Event-
        // Registrierung jetzt entfernt wurde, soll auch die Schatten-
        // Parent-Reg wegfallen — sonst behauptet „Meine Events"
        // weiterhin „Registered" obwohl der User kein einziges Sub-
        // Event mehr hat.
        if (props.parentEvent.subEventsOnlyMode) {
          const remainingActive = props.childEvents
            .map(ce => ce.id)
            .filter(id => id !== childEventId && peerIdsToCancel.indexOf(id) < 0)
            .filter(id => registeredSet.has(id));
          if (remainingActive.length === 0) {
            setProcessingMessage(isDe ? 'Hauptevent-Eintrag wird entfernt…' : 'Removing main-event entry…');
            // v17.22: Schatten-Parent-Cancel → Notifications unterdrücken
            // (die Sub-Event-Abmeldung hat ihre eigene Mail schon geschickt).
            try { await props.cancelRegistration(props.parentEvent.id, { suppressNotifications: true }); }
            catch (err) { console.warn('[DEX] shadow-parent cancel failed:', err); }
          }
        }
      } else {
        await props.registerForEvent(childEventId, {});
      }
      setProcessingMessage(isDe ? 'Aktualisiere…' : 'Refreshing…');
      await refresh();
      await props.onMutated();
    } finally {
      setBusyId(null);
      setProcessingMessage('');
    }
  };

  // v22.13: nur Sub-Events zeigen, die der User laut Sub-Event-Sichtbarkeit
  // sehen darf — ODER in denen er bereits aktiv angemeldet ist.
  const myEmailLc = (currentUser.email || '').toLowerCase();
  const isParentOrganizer = (props.parentEvent.organizerEmails || []).some(e => (e || '').toLowerCase() === myEmailLc);
  const visibleChildren = (isAdmin || isParentOrganizer)
    ? props.childEvents
    : props.childEvents.filter(ce =>
        registeredSet.has(ce.id)
        || isEventVisibleForUser(ce, currentUser.email, currentUser.location, groupEmails, currentUser.jobTitle));

  if (visibleChildren.length === 0) return null;

  // v15.13: Bezeichnung kommt jetzt direkt aus event.childEventTermPlural /
  // childEventTermSingular (Wizard-Setting). Fallback: „Sub-Events"-Begriff.
  const termPlural = props.parentEvent.childEventTermPlural || '';
  const termSingular = props.parentEvent.childEventTermSingular || '';
  const headerLabel = termPlural || (isDe ? 'Sub-Events' : 'Sub-events');
  const hintLabel = termSingular
    ? (isDe
        ? `Eine ${termSingular} kannst du jederzeit nachträglich an- oder abmelden. Bei jeder Aktion bekommst du eine Bestätigungs-Mail und der Termin wird in Outlook angelegt bzw. zurückgezogen.`
        : `You can register or cancel a ${termSingular} at any time. Every action triggers a confirmation mail and creates or removes the Outlook calendar entry.`)
    : (isDe
        ? 'Sub-Events kannst du jederzeit nachträglich an- oder abmelden. Bei jeder Aktion bekommst du eine Bestätigungs-Mail und der Termin wird in Outlook angelegt bzw. zurückgezogen.'
        : 'You can register or cancel sub-events at any time. Every action triggers a confirmation mail and creates or removes the Outlook calendar entry.');
  return (
    <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4 }}>
        {headerLabel}
      </div>
      <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginBottom: 10, lineHeight: 1.45 }}>
        {hintLabel}
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
        {visibleChildren.map(ce => {
          const isReg = registeredSet.has(ce.id);
          const isBusy = busyId === ce.id;
          const deadlinePassed = !!(ce.registrationDeadline && new Date(ce.registrationDeadline) < new Date());
          const count = counts[ce.id] || 0;
          const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
          const isFull = hasCap && count >= (ce.maxParticipants || 0);
          // v22.22: Vergangene Sessions sind weder an- noch abmeldbar
          // (Abmelde-Sperre für vergangene Events; EventContext blockt zusätzlich).
          const disabled = isBusy || (deadlinePassed && !isReg) || (isFull && !isReg) || isEventOver(ce);
          // v11.31: Custom-Field-Antworten gehören INS Sub-Event-Karten-
          // Layout, nicht ausserhalb. Maintainer-Wunsch: Tags zwischen
          // der Datums-/Adress-Zeile und den Action-Buttons (rechts) als
          // kleiner grüner Pastell-Stripe IN der Karte.
          const filledFieldTags = (() => {
            if (!isReg) return null;
            const data = seData[ce.id] || {};
            const filled = (ce.eventSpecificFields || [])
              .filter(f => f.label && (data[f.id] || '').trim())
              .map(f => ({ label: f.label, value: data[f.id], type: f.type }));
            if (filled.length === 0) return null;
            return (
              <div style={{ marginTop: 6, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                {filled.map(({ label, value, type }) => (
                  <FieldAnswerTag key={label} label={label} value={value} type={type} small />
                ))}
              </div>
            );
          })();
          return (
            <div key={ce.id} style={{
              display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12,
              padding: '8px 12px', borderRadius: 8,
              background: isReg ? 'rgba(134,188,37,0.08)' : 'var(--dex-gray-50, #fafafa)',
              border: `1px solid ${isReg ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
            }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 600, fontSize: '0.88rem' }}>{ce.title || (isDe ? 'Session ohne Titel' : 'Untitled session')}</div>
                <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                  {ce.startDate && <>{fmt(ce.startDate)}{ce.endDate ? ` – ${fmt(ce.endDate)}` : ''}</>}
                  {ce.location && <>&nbsp;·&nbsp;{ce.location}</>}
                  {hasCap && (
                    <>&nbsp;·&nbsp;<span style={{ color: isFull ? 'var(--dex-red, #c00)' : 'inherit' }}>{count}/{ce.maxParticipants}</span></>
                  )}
                  {isReg && (
                    <span style={{ marginLeft: 8, color: 'var(--dex-green)', fontWeight: 600 }}>
                      ({isDe ? 'Angemeldet' : 'Registered'})
                    </span>
                  )}
                  {isFull && !isReg && (
                    <span style={{ marginLeft: 8, color: 'var(--dex-red, #c00)', fontWeight: 600 }}>
                      ({isDe ? 'voll' : 'full'})
                    </span>
                  )}
                </div>
                {filledFieldTags}
              </div>
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                {isReg && (ce.eventSpecificFields || []).filter(f => f.label).length > 0 && (
                  <button
                    className="btn btn-secondary"
                    style={{ fontSize: '0.75rem', padding: '4px 10px' }}
                    disabled={isBusy}
                    onClick={() => {
                      setEditingId(ce.id);
                      setEditDraft({ ...(seData[ce.id] || {}) });
                    }}
                  >
                    {isDe ? 'Bearbeiten' : 'Edit'}
                  </button>
                )}
                <button
                  className={`btn ${isReg ? 'btn-secondary' : 'btn-primary'}`}
                  style={{ fontSize: '0.75rem', padding: '4px 12px', minWidth: 92 }}
                  disabled={disabled}
                  onClick={() => handleToggle(ce.id, isReg)}
                >
                  {isBusy ? '...' : (isReg ? (isDe ? 'Abmelden' : 'Cancel') : (isDe ? 'Anmelden' : 'Register'))}
                </button>
              </div>
            </div>
          );
        })}
      </div>

      {/* v10.27: Sub-Event-Edit-Modal — analog zum Hauptevent-Inline-Edit
          (siehe weiter oben), aber als Modal damit es uebersichtlich
          bleibt. Beim Speichern geht der Draft per
          updateMyRegistration(subEventId, ...) in die Subsite zurück. */}
      {editingId && (() => {
        const ce = props.childEvents.find(c => c.id === editingId);
        if (!ce) return null;
        const fields = (ce.eventSpecificFields || []).filter(f => f && f.label);
        const closeModal = (): void => { setEditingId(null); setEditDraft({}); };
        const saveEdit = async (): Promise<void> => {
          setSavingEdit(true);
          try {
            await props.updateMyRegistration(editingId, editDraft);
            await refresh();
            await props.onMutated();
            closeModal();
          } finally {
            setSavingEdit(false);
          }
        };
        return (
          <Modal
            open={true}
            onClose={closeModal}
            dismissable={!savingEdit}
            maxWidth={520}
            padding={24}
            ariaLabel={ce.title || (isDe ? 'Sub-Event' : 'Sub-event')}
          >
              <h3 style={{ margin: '0 0 6px', fontSize: '1.1rem' }}>
                {ce.title || (isDe ? 'Sub-Event' : 'Sub-event')}
              </h3>
              <p style={{ margin: '0 0 18px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {isDe
                  ? 'Hier kannst du deine Antworten zu diesem Sub-Event aktualisieren.'
                  : 'Update your answers for this sub-event here.'}
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12, marginBottom: 18 }}>
                {fields.map(f => {
                  const val = editDraft[f.id] || '';
                  const setVal = (v: string): void => setEditDraft(prev => ({ ...prev, [f.id]: v }));
                  return (
                    <div key={f.id}>
                      <label className="form-label" style={{ display: 'block', fontSize: '0.85rem', marginBottom: 4 }}>
                        {f.label}
                        {f.required && <span style={{ color: 'var(--dex-red, #c00)', marginLeft: 4 }}>*</span>}
                        {/* v11.16: konsistenter InfoTooltip — gleicher
                            Look wie auf der Register-Page und im
                            Sub-Event-Modal. */}
                        {f.helpText && <InfoTooltip text={f.helpText} />}
                      </label>
                      {f.type === 'select' ? (
                        f.multi ? (
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                            {(f.options || []).map(opt => {
                              const current = val.split(' | ').map(s => s.trim()).filter(Boolean);
                              const checked = current.indexOf(opt) >= 0;
                              return (
                                <label key={opt} style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                                  <input
                                    type="checkbox"
                                    checked={checked}
                                    onChange={e => {
                                      const next = e.target.checked
                                        ? [...current, opt]
                                        : current.filter(x => x !== opt);
                                      setVal(next.join(' | '));
                                    }}
                                  />
                                  {opt}
                                </label>
                              );
                            })}
                          </div>
                        ) : (
                          <select className="form-select" value={val} onChange={e => setVal(e.target.value)}>
                            <option value="">{isDe ? '— bitte wählen —' : '— please select —'}</option>
                            {(f.options || []).map(opt => <option key={opt} value={opt}>{opt}</option>)}
                          </select>
                        )
                      ) : f.type === 'checkbox' ? (
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', cursor: 'pointer' }}>
                          <input
                            type="checkbox"
                            checked={val === 'true'}
                            onChange={e => setVal(e.target.checked ? 'true' : 'false')}
                          />
                          {f.label}
                        </label>
                      ) : f.type === 'number' ? (
                        <input className="form-input" type="number" value={val} onChange={e => setVal(e.target.value)} />
                      ) : (
                        <input className="form-input" type="text" value={val} onChange={e => setVal(e.target.value)} />
                      )}
                    </div>
                  );
                })}
              </div>
              <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button className="btn btn-secondary" onClick={closeModal} disabled={savingEdit}>
                  {isDe ? 'Abbrechen' : 'Cancel'}
                </button>
                <button className="btn btn-primary" onClick={saveEdit} disabled={savingEdit}>
                  {savingEdit ? '…' : (isDe ? 'Speichern' : 'Save')}
                </button>
              </div>
          </Modal>
        );
      })()}

      {/* v15.8: Peer-Cancel-Modal mit Checkbox-Liste. User hakt direkt
          an, welche zusätzlichen Sub-Events er mitabmelden will.
          Default: target ist immer abgemeldet (lock-Checkbox); peers
          unchecked. Klick auf „Abmelden" verarbeitet die Auswahl. */}
      {peerCancelDialog && <PeerCancelCheckboxModal dlg={peerCancelDialog} isDe={isDe} isSectionedEvent={!!props.parentEvent.requireSubEventSelection} />}
      {/* v15.21: Globaler Progress-Overlay während (Peer-)Cancel +
          Registrierung — blockiert die ganze Seite, damit der User nicht
          versehentlich nochmal klickt und sieht, dass die Aktion läuft. */}
      {processingMessage && (
        <div
          role="dialog"
          aria-modal="true"
          aria-label={isDe ? 'Wird verarbeitet' : 'Processing'}
          style={{
            position: 'fixed', inset: 0, zIndex: 2000,
            background: 'rgba(0,0,0,0.55)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            padding: 16,
          }}
        >
          <div style={{
            background: '#fff', borderRadius: 12, padding: '28px 32px',
            maxWidth: 420, width: '100%',
            boxShadow: '0 16px 48px rgba(0,0,0,0.35)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16,
            textAlign: 'center',
          }}>
            <div aria-hidden="true" style={{
              width: '100%', height: 6, borderRadius: 3,
              background: '#e5e5e5',
              overflow: 'hidden', position: 'relative',
            }}>
              <div style={{
                position: 'absolute', top: 0, bottom: 0,
                width: '40%',
                background: 'var(--dex-green, #86bc25)',
                borderRadius: 3,
                animation: 'dexProgressSlide 1.2s ease-in-out infinite',
              }} />
            </div>
            <div style={{ fontWeight: 700, fontSize: '1rem', color: 'var(--dex-gray-800)' }}>
              {processingMessage}
            </div>
            <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>
              {isDe ? 'Bitte einen Moment Geduld…' : 'Please wait a moment…'}
            </div>
          </div>
          <style>{`@keyframes dexProgressSlide { 0% { left: -40%; } 100% { left: 100%; } }`}</style>
        </div>
      )}
    </div>
  );
}

// v15.8: Peer-Cancel-Modal mit Checkbox-Liste pro Sub-Event.
// Target ist locked und immer ausgewählt (User hat ja explizit Cancel
// geklickt); Peers sind initial unchecked. Eine „Bestätigen"-Aktion
// schickt die peer-IDs der angehakten Sub-Events an den resolver.
function PeerCancelCheckboxModal(props: {
  dlg: {
    targetTitle: string;
    peers: { id: string; title: string; startDate?: string; location?: string }[];
    resolve: (_choice: { peerIds: string[] } | 'abort') => void;
  };
  isDe: boolean;
  isSectionedEvent: boolean;
}): React.ReactElement {
  const { dlg, isDe, isSectionedEvent } = props;
  const [selectedPeerIds, setSelectedPeerIds] = React.useState<Set<string>>(new Set());
  const peerTermPl = isSectionedEvent
    ? (isDe ? 'Sections' : 'sections')
    : (isDe ? 'Sub-Events' : 'sub-events');
  const togglePeer = (id: string): void => {
    setSelectedPeerIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  const formatLine = (startDate?: string, location?: string): string => {
    const dateStr = startDate
      ? new Date(startDate).toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })
      : '';
    return [dateStr, location].filter(Boolean).join(' · ');
  };
  const totalCount = 1 + selectedPeerIds.size;
  return (
    <Modal
      open={true}
      onClose={() => dlg.resolve('abort')}
      maxWidth={560}
      ariaLabel={isDe ? 'Abmeldung — Auswahl' : 'Cancellation — selection'}
    >
      <h3 style={{ margin: 0, fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
        {isDe ? `Welche ${peerTermPl} möchtest du abmelden?` : `Which ${peerTermPl} do you want to cancel?`}
      </h3>
      <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
        {isDe
          ? 'Wähle die Einträge an, die du abmelden willst. Nicht angewählte Anmeldungen bleiben erhalten.'
          : 'Tick the entries you want to cancel. Unticked registrations stay active.'}
      </p>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 320, overflowY: 'auto' }}>
        <label style={{
          display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
          borderRadius: 8,
          border: '1px solid var(--dex-red, #d62828)',
          background: 'rgba(214,40,40,0.06)',
        }}>
          <input type="checkbox" checked={true} disabled style={{ marginTop: 3, accentColor: 'var(--dex-red, #d62828)' }} />
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)' }}>{dlg.targetTitle}</div>
            <div style={{ fontSize: '0.75rem', color: 'var(--dex-red-dark, #8b1414)', marginTop: 2 }}>
              {isDe ? 'Dieses hast du gerade zur Abmeldung gewählt' : 'You just selected this one to cancel'}
            </div>
          </div>
        </label>
        {dlg.peers.map(p => {
          const checked = selectedPeerIds.has(p.id);
          const meta = formatLine(p.startDate, p.location);
          return (
            <label key={p.id} style={{
              display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
              borderRadius: 8,
              border: `1px solid ${checked ? 'var(--dex-red, #d62828)' : 'var(--dex-gray-200)'}`,
              background: checked ? 'rgba(214,40,40,0.04)' : '#fff',
              cursor: 'pointer',
              transition: 'border-color 0.15s, background 0.15s',
            }}>
              <input
                type="checkbox"
                checked={checked}
                onChange={() => togglePeer(p.id)}
                style={{ marginTop: 3, accentColor: 'var(--dex-red, #d62828)', cursor: 'pointer' }}
              />
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 600, color: 'var(--dex-gray-800)' }}>{p.title}</div>
                {meta && (
                  <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>{meta}</div>
                )}
              </div>
            </label>
          );
        })}
      </div>
      <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 4 }}>
        <button
          type="button"
          className="btn btn-secondary"
          onClick={() => dlg.resolve('abort')}
        >
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        <button
          type="button"
          className="btn btn-primary"
          onClick={() => dlg.resolve({ peerIds: Array.from(selectedPeerIds) })}
          style={{ background: 'var(--dex-red, #d62828)', borderColor: 'var(--dex-red, #d62828)' }}
        >
          {isDe ? `Abmelden (${totalCount})` : `Cancel ${totalCount}`}
        </button>
      </div>
    </Modal>
  );
}
