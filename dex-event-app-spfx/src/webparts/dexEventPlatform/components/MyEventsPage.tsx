/**
 * Meine Events - zeigt alle Events fuer die der User registriert ist.
 * Laedt Registrierungen aus den jeweiligen Teilnehmerlisten.
 * Ermoeglicht Abmeldung mit Zwei-Schritt-Bestaetigung.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import OrganizerList from './OrganizerList';

import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { DeloitteEvent, EventSpecificField, AgendaItem, TransferTime, QuizQuestion } from '../types';
import { SPRegistration } from '../services/EventService';
import { wrapTemplate } from '../services/EmailTemplates';
import { useLanguage } from '../context/LanguageContext';
import PdfViewer from './PdfViewer';

interface MyEventEntry {
  event: DeloitteEvent;
  registration: SPRegistration;
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
 * Eine Frage gilt als korrekt, wenn die Menge der gewaehlten Indices exakt
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
 * Zaehlt wie viele Fragen (mindestens teilweise) beantwortet wurden.
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
 * - onProgress: wird bei jedem "Weiter"/"Zurueck" aufgerufen + am Ende mit isComplete=true
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

  // Antworten initialisieren: fuer jede Frage ein Array (leer wenn unbeantwortet).
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
  // Sections faellt alles auf Cluster-Groesse zurueck.
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
    // Sections in Reihenfolge der ersten Frage-Erwaehnung
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

  // Welche Gruppe enthaelt die erste unbeantwortete Frage?
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
        ? (isDe ? `Weiter (${answeredCount}/${quiz.length})` : `Continue (${answeredCount}/${quiz.length})`)
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
          Fun-Zone: {resumeLabel}
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
  // Mobile-Erkennung: Auf Mobile nutzen wir react-pdf (Canvas), auf Desktop bleibt iframe (bewaehrt)
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

      // XHR fuer Binary-Download (zuverlaessiger als fetch fuer SharePoint)
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
          // Desktop oder Bilder: Blob-URL + iframe (bewaehrt)
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
                  /* PDF via react-pdf (Canvas) - funktioniert Desktop + Mobile, eigenes Scrolling */
                  <PdfViewer blob={pdfBlob} height={600} />
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
  const { events, isEventsLoading, getMyRegistration, getMyEventNumbers, cancelRegistration, updateMyRegistration } = useEvents();
  const { t } = useLanguage();
  const [myEvents, setMyEvents] = React.useState<MyEventEntry[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [cancellingId, setCancellingId] = React.useState<string | null>(null);
  const [isCancelling, setIsCancelling] = React.useState(false);
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [editData, setEditData] = React.useState<Record<string, string>>({});
  const [isSaving, setIsSaving] = React.useState(false);
  const [loadError, setLoadError] = React.useState('');

  React.useEffect(() => {
    // Warten bis Events fertig geladen sind, sonst zeigen wir Fehler obwohl nur noch geladen wird
    if (isEventsLoading) {
      setIsLoading(true);
      setLoadError('');
      return;
    }
    loadMyRegistrations();
  }, [events, isEventsLoading]);

  async function loadMyRegistrations(): Promise<void> {
    setIsLoading(true);
    setLoadError('');
    const entries: MyEventEntry[] = [];

    // Schneller Pfad: DEX_Participants abfragen
    const myNumbers = await getMyEventNumbers();
    const allMyNumbers = [...myNumbers.registered, ...myNumbers.waitlisted];

    if (allMyNumbers.length > 0) {
      // Nur Events laden die in DEX_Participants stehen
      const relevantEvents = events.filter(e => e.eventNumber && allMyNumbers.indexOf(e.eventNumber) >= 0);
      for (const event of relevantEvents) {
        try {
          const reg = await getMyRegistration(event.id);
          if (reg) {
            entries.push({ event, registration: reg });
          }
        } catch { /* */ }
      }
      // Zusatzschleife: abgemeldete Events finden. DEX_Participants haelt nur
      // EventRegistered/EventOnWaitlist - bei Abmeldung wird die EventNumber dort
      // entfernt. Ohne diese Schleife waeren alte Abmeldungen im "My Events"-Tab
      // unsichtbar, sobald der User noch fuer mind. ein Event angemeldet ist.
      const remainingEvents = events.filter(e => !e.eventNumber || allMyNumbers.indexOf(e.eventNumber) < 0);
      for (const event of remainingEvents) {
        try {
          const reg = await getMyRegistration(event.id);
          if (reg && reg.Status === 'Abgemeldet') {
            entries.push({ event, registration: reg });
          }
        } catch { /* */ }
      }
    } else {
      // Fallback: Alter Weg fuer Altdaten ohne DEX_Participants-Eintrag
      for (const event of events) {
        try {
          const reg = await getMyRegistration(event.id);
          if (reg) {
            entries.push({ event, registration: reg });
          }
        } catch { /* */ }
      }
    }

    if (entries.length === 0 && allMyNumbers.length > 0) {
      setLoadError('Registrierungen konnten nicht geladen werden.');
    }
    setMyEvents(entries);
    setIsLoading(false);
  }

  // Eigentliche Cancel-Logik (direkt ausfuehren, ohne 2-Klick-Bestaetigung).
  // Wird sowohl von handleCancel (beim 2. Klick) als auch vom Auto-Cancel-
  // Deep-Link (direkt nach Navigation) genutzt.
  const performCancel = async (eventId: string): Promise<void> => {
    setCancellingId(eventId);
    setIsCancelling(true);

    // Check if this is a late cancellation
    const entry = myEvents.find(e => e.event.id === eventId);
    const isLateCancellation = entry?.event.lastDeregisterDate && new Date(entry.event.lastDeregisterDate) < new Date();

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
            // EmailType 'Info' (Choice-Feld laesst nur Anmeldung/Abmeldung/
            // Warteliste/Nachruecken/Info zu).
            await svc.queueEmail(subject, toList, toNames, body, 'Info', entry.event.title, eventId)
              .catch(err => console.warn('[DEX] LateCancel queueEmail failed:', err));
          }
        } catch { /* Email-Fehler ignorieren */ }
      }
      await loadMyRegistrations();
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

  // Auto-Cancel: wenn die Seite via Deep-Link mit Intent 'auto-cancel' geoeffnet
  // wurde (z.B. aus einer Outlook-Decline-Reminder-Mail), die Registrierung
  // direkt stornieren - OHNE dass der User zusaetzlich auf "Abmeldung
  // bestaetigen" klicken muss. Der Klick auf den Link in der Mail gilt als
  // Bestaetigung. Da der User eingeloggt sein muss und nur seine eigene
  // Registrierung cancelt, ist das sicher.
  const didAutoOpen = React.useRef(false);
  React.useEffect(() => {
    if (didAutoOpen.current) return;
    if (navIntent !== 'auto-cancel' || !selectedEventId) return;
    // Warten bis die Registrierungen geladen sind, sonst findet performCancel
    // den entry nicht (late-cancel-check schlaegt fehl).
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

  if (isLoading) {
    return (
      <div className="page-container text-center">
        <p style={{ color: 'var(--dex-gray-400)', padding: 48 }}>{t('myevents.loading')}</p>
      </div>
    );
  }

  return (
    <div className="page-container">
      <h2 className="mb-16">{t('myevents.title')}</h2>

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

      {activeEntries.length > 0 && (
        <div className="my-events-list">
          {activeEntries.map(({ event, registration }) => {
            // Custom Data parsen und IDs zu Labels mappen
            let customData: Record<string, string> = {};
            try {
              if (registration.CustomData) customData = JSON.parse(registration.CustomData);
            } catch { /* */ }

            // Feld-ID zu Label-Map aus den Event-Feldern erstellen
            const fieldLabelMap: Record<string, string> = {};
            for (const field of event.eventSpecificFields) {
              fieldLabelMap[field.id] = field.label;
            }

            // "salutation" überspringen (wird schon im Namen angezeigt)
            const displayData = Object.keys(customData)
              .filter(key => key !== 'salutation' && customData[key])
              .map(key => ({
                label: fieldLabelMap[key] || key,
                value: customData[key],
              }));

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
                      <img
                        src={event.imageUrl}
                        alt={event.title}
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
                      <span className={`badge ${getStatusBadgeClass(registration.Status)}`} style={{ flexShrink: 0, marginLeft: 12 }}>
                        {registration.Status === 'Warteliste' && registration.TeilnehmerID && event.maxParticipants > 0
                          ? `${getStatusLabel(registration.Status, t)} #${registration.TeilnehmerID - event.maxParticipants}`
                          : getStatusLabel(registration.Status, t)}
                      </span>
                    </div>

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

                    {/* Organizer mit Foto (Hover vergroessert) */}
                    {event.organizers.length > 0 && (
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
                  </div>
                </div>

                {/* Custom Fields als kompakte Tags */}
                {!editingId || editingId !== event.id ? (
                  displayData.length > 0 && (
                    <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                      {displayData.map(({ label, value }) => (
                        <span key={label} style={{
                          fontSize: '0.78rem', padding: '3px 10px', borderRadius: 12,
                          background: 'var(--dex-gray-100)', color: 'var(--dex-gray-700)',
                        }}>
                          {label}: <strong>{value}</strong>
                        </span>
                      ))}
                    </div>
                  )
                ) : (
                  <div style={{ marginTop: 12 }}>
                    {event.eventSpecificFields.map((field: EventSpecificField) => (
                      <div className="form-group" key={field.id} style={{ marginBottom: 10 }}>
                        <label className="form-label" style={{ fontSize: '0.82rem', marginBottom: 2 }}>
                          {field.required && <span className="required">*</span>}
                          {field.label}
                        </label>
                        {field.type === 'select' ? (
                          <select className="form-select" value={editData[field.id] || ''} onChange={e => setEditData({ ...editData, [field.id]: e.target.value })}>
                            <option value="">—</option>
                            {field.options && field.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                          </select>
                        ) : (
                          <input className="form-input" value={editData[field.id] || ''} onChange={e => setEditData({ ...editData, [field.id]: e.target.value })} placeholder={field.label} type={field.type === 'number' ? 'number' : 'text'} />
                        )}
                      </div>
                    ))}
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

                  return (
                    <div style={{ marginTop: 12 }}>
                      <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-600)', marginBottom: 8 }}>
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
                                <span style={{ width: 26, height: 26, borderRadius: '50%', background: 'var(--dex-green-dark, #6b9a1e)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                                  <Icon iconName={item.icon || 'Calendar'} style={{ fontSize: 12, color: '#fff' }} />
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
                {event.quiz && event.quiz.length > 0 && (
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
                      // isComplete=true setzt zusaetzlich QuizCompletedAt (fuer Statistik-Filter).
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
                        // VERIFY: re-read das Item und checke ob QuizAnswers tatsaechlich in SP landete
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

                {/* Registriert am + Aktionen */}
                <div style={{ marginTop: 12, paddingTop: 10, borderTop: '1px solid var(--dex-gray-200)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-400)' }}>
                    {t('myevents.registeredon')}: {formatDate(registration.RegistrationDate)}
                  </span>
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                    {cancellingId === event.id && !isCancelling && event.lastDeregisterDate && new Date(event.lastDeregisterDate) < new Date() && (
                      <span style={{ fontSize: '0.75rem', color: 'var(--dex-orange)', display: 'block', marginBottom: 4, width: '100%' }}>
                        {t('myevents.latecancel')}
                      </span>
                    )}
                    <button className="btn btn-secondary" style={{ fontSize: '0.78rem', padding: '4px 12px' }} onClick={() => { if (editingId === event.id) { setEditingId(null); } else { setEditData(customData); setEditingId(event.id); } }}>
                      {editingId === event.id ? t('general.cancel') : t('myevents.edit')}
                    </button>
                    <button className="btn" onClick={() => handleCancel(event.id)} disabled={isCancelling} style={{ fontSize: '0.78rem', padding: '4px 12px', background: cancellingId === event.id ? 'var(--dex-red)' : 'var(--dex-gray-200)', color: cancellingId === event.id ? '#fff' : 'var(--dex-gray-700)' }}>
                      {cancellingId === event.id ? (isCancelling ? '...' : t('myevents.confirmcancel')) : t('myevents.cancel')}
                    </button>
                    {cancellingId === event.id && !isCancelling && (
                      <button className="btn btn-secondary" onClick={() => setCancellingId(null)} style={{ fontSize: '0.78rem', padding: '4px 12px' }}>{t('myevents.keepreg')}</button>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {cancelledEntries.length > 0 && (
        <div>
          <h3 className="mt-24 mb-16" style={{ color: 'var(--dex-gray-400)' }}>{t('myevents.cancelledevents')}</h3>
          <div className="my-events-list">
            {cancelledEntries.map(({ event, registration }) => (
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
                  {t('myevents.cancelledon')}: {registration.CancellationDate ? formatDate(registration.CancellationDate) : '-'}
                </span>
                <span className="badge badge-red">{t('status.cancelled')}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
