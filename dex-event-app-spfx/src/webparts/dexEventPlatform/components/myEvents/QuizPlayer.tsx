/* QuizPlayer — aus MyEventsPage.tsx ausgelagert (Zeilen 168-550 des
 * urspruenglichen Stands, v30.65). Der Spieler samt seinen beiden Rechen-
 * helfern; beide werden ausserhalb nicht gebraucht und bleiben deshalb
 * modul-intern. Der Code ist zeichengleich uebernommen.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { QuizQuestion } from '../../types';
// v31.8: Optik über die gemeinsamen dex-ui-Klassen statt handgebauter Inline-
// Kästen — Inline-Styles können kein `:hover`, und die Antwortknöpfe waren
// damit von einer Beschriftung nicht zu unterscheiden. `ensureDexUiStyles()`
// ruft die Seiten-Komponente (`MyEventsPage`), Unterkomponenten nie.
import { cx } from '../dexUi';
import { AlertCircle, Check, X } from '../Icons';

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
  isDe,
  clusterSize = 1,
  initialAnswers,
  onProgress,
}: {
  quiz: QuizQuestion[];
  t: (key: string) => string;
  isDe: boolean;
  clusterSize?: number;
  initialAnswers?: number[][];
  onProgress?: (score: number, answers: number[][], isComplete: boolean) => void;
}): React.ReactElement {
  const size = Math.min(Math.max(1, clusterSize || 1), 4);

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
    // v31.8: Die Zahlen standen bisher in Klammern im Knopftext. Jetzt trägt
    // der Titel die Handlung, die Folgezeile den Stand — die Zahl selbst
    // zusätzlich als Pille rechts, damit sie auf einen Blick lesbar ist.
    const resumeLabel = completedAllInitial
      ? (isDe ? 'Antworten ansehen oder ändern' : 'Review or edit answers')
      : hadResumeData
        ? (isDe ? 'Quiz fortsetzen' : 'Resume quiz')
        : (isDe ? 'Quiz starten' : 'Start quiz');
    const resumeDesc = completedAllInitial
      ? (isDe
        ? `Du hast alle ${quiz.length} Fragen beantwortet — dein Ergebnis siehst du sofort.`
        : `You answered all ${quiz.length} questions — your result shows right away.`)
      : hadResumeData
        ? (isDe
          ? `${answeredCount} von ${quiz.length} Fragen beantwortet — du machst dort weiter, wo du aufgehört hast.`
          : `${answeredCount} of ${quiz.length} questions answered — you carry on where you left off.`)
        // v31.8 (Nachzug): Hier stand „deine Antworten werden unterwegs
        // gespeichert" — ein Versprechen, das der Code nicht hält. `onProgress`
        // ist optional; der Aufrufer (MyEventCard) bricht ohne `subsiteUrl` oder
        // ohne `__dexSpfxContext` still ab; `saveQuizProgress` liefert bei
        // fehlenden Quiz-Spalten (Teilnehmer ohne Manage-Lists-Recht) `false`,
        // und `saveProgress` schluckt jeden Fehler bewusst („State bleibt
        // lokal"). Gespeichert wird ausserdem nur beim Seitenwechsel, nicht bei
        // jeder Antwort. Der Satz nennt jetzt nur noch, was sicher eintritt.
        : (isDe
          ? `${quiz.length} ${quiz.length === 1 ? 'Frage' : 'Fragen'} — dein Ergebnis siehst du am Ende.`
          : `${quiz.length} ${quiz.length === 1 ? 'question' : 'questions'} — you see your result at the end.`);
    return (
      <div style={{ marginTop: 12 }}>
        <button type="button" className="dex-ui-action" onClick={() => setShowQuiz(true)}>
          <span className="dex-ui-action-icon"><Icon iconName="Game" style={{ fontSize: 16 }} /></span>
          <span className="dex-ui-action-body">
            <span className="dex-ui-action-title">{resumeLabel}</span>
            <span className="dex-ui-action-desc">{resumeDesc}</span>
          </span>
          {hadResumeData && (
            <span className="dex-ui-action-badge">
              <span className={cx('dex-ui-pill', completedAllInitial ? 'dex-ui-pill--green' : 'dex-ui-pill--orange')}>
                {answeredCount} / {quiz.length}
              </span>
            </span>
          )}
        </button>
      </div>
    );
  }

  // ===== Summary =====
  if (showSummary) {
    const score = computeQuizScore(quiz, allAnswers);
    return (
      <div className="dex-ui-card dex-ui-card--soft" style={{ marginTop: 12 }}>
        <div style={{ textAlign: 'center', marginBottom: 16 }}>
          <div style={{ fontSize: '2rem', marginBottom: 8 }} aria-hidden="true">{score === quiz.length ? '🎉' : score >= quiz.length / 2 ? '👏' : '💪'}</div>
          <div className="dex-ui-kpi-value">
            {score} / {quiz.length} {isDe ? 'richtig' : 'correct'}
          </div>
          <div className="dex-ui-muted" style={{ marginTop: 4 }}>
            {score === quiz.length
              ? (isDe ? 'Perfekt! Alle richtig!' : 'Perfect! All correct!')
              : (isDe ? 'Gut gemacht!' : 'Well done!')}
          </div>
        </div>
        {/* Einzel-Auflistung: Fragen + eigene Antwort + Haken.
            v31.8: je Frage ein `dex-ui-callout` — die Farbe trägt dieselbe
            Bedeutung wie überall (grün richtig, rot falsch, orange offen),
            statt dreier handgebauter Kästen mit eigenen Farbwerten. */}
        <div className="dex-ui-stack" style={{ gap: 8, marginBottom: 16 }}>
          {quiz.map((q, qi) => {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const correct = q.correctIndices || [(q as any).correctIndex || 0];
            const given = allAnswers[qi] || [];
            const isCorrect = given.length > 0 && given.length === correct.length && correct.every(c => given.includes(c));
            const unanswered = given.length === 0;
            return (
              <div
                key={q.id || qi}
                className={cx('dex-ui-callout', unanswered ? 'dex-ui-callout--warn' : isCorrect ? 'dex-ui-callout--success' : 'dex-ui-callout--danger')}
              >
                <span className="dex-ui-callout-icon">
                  {unanswered ? <AlertCircle size={16} /> : isCorrect ? <Check size={16} /> : <X size={16} />}
                </span>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 600, marginBottom: 4 }}>
                    {qi + 1}. {q.question}
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
                  <div style={{ fontSize: '0.78rem' }}>
                    {unanswered
                      ? (isDe ? 'Nicht beantwortet' : 'Not answered')
                      : `${isDe ? 'Deine Antwort' : 'Your answer'}: ${given.map(i => q.options[i]).join(', ')}`}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
        {/* v31.8: Knöpfe links beim Inhalt (Leitfaden 2a′) statt zentriert. */}
        <div className="dex-ui-inline">
          <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={goBack}>
            {isDe ? 'Antworten ändern' : 'Edit answers'}
          </button>
          <button type="button" className="btn btn-secondary dex-ui-btn-sm" onClick={() => setShowQuiz(false)}>
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
    <div className="dex-ui-card dex-ui-card--soft" style={{ marginTop: 12 }}>
      {/* Progress-Bar */}
      <div style={{ marginBottom: 16 }}>
        <div className="dex-ui-inline" style={{ justifyContent: 'space-between', marginBottom: 6 }}>
          <span className="dex-ui-muted">
            {isDe ? 'Fortschritt' : 'Progress'}: {answeredCount} / {quiz.length} ({progressPct}%)
          </span>
          <span className="dex-ui-muted">
            {isDe ? 'Seite' : 'Page'} {currentGroupIdx + 1} / {groups.length}
          </span>
        </div>
        <div className="dex-ui-progress">
          <div className="dex-ui-progress-bar" style={{ width: `${progressPct}%` }} />
        </div>
      </div>

      {/* Bereichs-Titel, wenn vorhanden. v31.8: keine Versalien — der Titel
          kommt vom Organizer und wird angezeigt, wie er ihn geschrieben hat. */}
      {currentGroup.title && (
        <h3 className="dex-ui-card-head-title" style={{ margin: '0 0 14px' }}>
          {currentGroup.title}
        </h3>
      )}

      {/* Fragen der aktuellen Gruppe */}
      <div className="dex-ui-stack" style={{ gap: 18 }}>
        {currentGroup.indices.map(qIdx => {
          const question = quiz[qIdx];
          if (!question) return null;
          const given = allAnswers[qIdx] || [];
          return (
            <div key={question.id || qIdx}>
              <div className="dex-ui-label" style={{ fontSize: '0.95rem', marginBottom: 10 }}>
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
              {/* v31.8: `dex-ui-choice` statt Inline-Kasten — die Antwort ist
                  eine Auswahl und braucht Hover, Fokusring und einen
                  sichtbaren Haken; vorher war sie von Text nicht zu
                  unterscheiden. Mehrfachauswahl bleibt möglich, deshalb
                  `aria-pressed` statt Radio-Semantik. */}
              <div className="dex-ui-stack" style={{ gap: 8 }}>
                {question.options.map((opt, i) => {
                  const isSelected = given.includes(i);
                  return (
                    <button
                      key={i}
                      type="button"
                      className={cx('dex-ui-choice', 'dex-ui-choice--multi', isSelected && 'is-active')}
                      aria-pressed={isSelected}
                      onClick={() => toggleOption(qIdx, i)}
                    >
                      {/* v31.8: `dex-ui-choice-label` statt `-title` — eine Liste
                          von Antwortoptionen, in der alles fett ist, hat keine
                          Gewichtung mehr. */}
                      <span className="dex-ui-choice-body">
                        <span className="dex-ui-choice-label">{opt}</span>
                      </span>
                      {/* v31.8: `--multi` macht aus dem runden Haken ein eckiges
                          Kästchen. Rund bedeutet in dieser App überall
                          &bdquo;genau eine&ldquo;; hier nimmt `toggleOption` einen
                          zweiten Klick wieder zurück. */}
                      <span className="dex-ui-choice-check">{isSelected && <Check size={12} />}</span>
                    </button>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>

      {/* Navigation. v31.8: beide Knöpfe links beim Inhalt (Leitfaden 2a′).
          „Zurück" entfällt auf der ersten Seite ganz — er war dort schon
          vorher abgeschaltet und unsichtbar, hinterließ aber eine Lücke. */}
      <div className="dex-ui-inline" style={{ marginTop: 16 }}>
        {currentGroupIdx > 0 && (
          <button
            type="button"
            className="btn btn-secondary dex-ui-btn-sm"
            onClick={goBack}
            disabled={isSaving}
          >
            {isDe ? 'Zurück' : 'Back'}
          </button>
        )}
        <button
          type="button"
          className="btn btn-primary dex-ui-btn-sm"
          onClick={goNext}
          disabled={isSaving}
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
