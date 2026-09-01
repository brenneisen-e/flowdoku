/* QuizPlayer — aus MyEventsPage.tsx ausgelagert (Zeilen 168-550 des
 * urspruenglichen Stands, v30.65). Der Spieler samt seinen beiden Rechen-
 * helfern; beide werden ausserhalb nicht gebraucht und bleiben deshalb
 * modul-intern. Der Code ist zeichengleich uebernommen.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { QuizQuestion } from '../../types';

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
export default function QuizPlayer({
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
