/* QuizStatsSection — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 9624-9879 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich übernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3: Auswertungs-Karte nach docs/ui-leitfaden.md 5a Punkt 8 — eingeklappt,
 * Kopf mit Zähler und Chevron rechts, darin Kennzahlen (`dex-ui-kpi`), die
 * Fragen-Balken als `dex-ui-progress` und die Top 10 als `dex-ui-table` mit
 * Personen-Zelle (5b). Die Rechenwege sind unverändert.
 */
import * as React from 'react';
import { ChevronDown, FileText } from '../../Icons';
import { formatDate } from '../../../utils/eventStatus';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import { cx, ensureDexUiStyles } from '../../dexUi';
import { PersonContactHover } from '../../PersonContactHover';

export interface QuizStatsSectionProps {
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
}

export const QuizStatsSection: React.FC<QuizStatsSectionProps> = (p) => {
  const { registrations, selectedEvent } = p;
        // v31.3: eigener Aufklapp-Zustand statt <details>/<summary>. Der native
        // Marker liess sich weder ausrichten noch drehen; jetzt trägt die
        // Kopfzeile denselben Chevron wie alle anderen Auswertungs-Karten.
        const [open, setOpen] = React.useState(false);
        ensureDexUiStyles();
        // Teilnehmer mit mindestens einer beantworteten Frage (nicht nur "komplett durchgeführt").
        // Dadurch erscheinen auch Teilnehmer, die mittendrin aufgehört haben.
        const regsWithQuiz = registrations.filter(r => {
          if (!r.QuizAnswers) return false;
          try {
            const parsed = JSON.parse(r.QuizAnswers);
            return Array.isArray(parsed) && parsed.some((a: number[]) => Array.isArray(a) && a.length > 0);
          } catch { return false; }
        });
        const regsCompleted = regsWithQuiz.filter(r => typeof r.QuizCompletedAt === 'string' && r.QuizCompletedAt);
        const totalQuizzes = regsWithQuiz.length;
        const totalCompleted = regsCompleted.length;

        // Pro Frage: wie viele haben sie überhaupt beantwortet, wie viele richtig
        const perQuestion = selectedEvent.quiz.map((q, qIdx) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const correct = (q as any).correctIndices || [(q as any).correctIndex || 0];
          let correctCount = 0;
          let answeredCount = 0;
          for (const reg of regsWithQuiz) {
            try {
              const answers = JSON.parse(reg.QuizAnswers || '[]');
              const given: number[] = Array.isArray(answers[qIdx]) ? answers[qIdx] : [];
              if (given.length === 0) continue;
              answeredCount++;
              const isRight = correct.length === given.length && correct.every((c: number) => given.indexOf(c) >= 0);
              if (isRight) correctCount++;
            } catch { /* skip */ }
          }
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const imageBase64 = (q as any).imageBase64 as string | undefined;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const section = (q as any).section as string | undefined;
          return { question: q.question, imageBase64, section, correctCount, answeredCount, total: totalQuizzes };
        });

        // Top 10 nach Score, bei Gleichstand: abgeschlossene vor nicht-abgeschlossenen, dann Zeitpunkt
        const top10 = regsWithQuiz.slice().sort((a, b) => {
          const sa = a.QuizScore || 0;
          const sb = b.QuizScore || 0;
          if (sb !== sa) return sb - sa;
          const aDone = !!a.QuizCompletedAt;
          const bDone = !!b.QuizCompletedAt;
          if (aDone !== bDone) return aDone ? -1 : 1;
          const ta = new Date(a.QuizCompletedAt || 0).getTime();
          const tb = new Date(b.QuizCompletedAt || 0).getTime();
          return ta - tb;
        }).slice(0, 10);

        // Gruppen in Reihenfolge der ersten Erwähnung
        const hasSections = perQuestion.some(pq => !!pq.section);
        const sectionsInOrder: string[] = [];
        for (const pq of perQuestion) {
          if (pq.section && sectionsInOrder.indexOf(pq.section) < 0) sectionsInOrder.push(pq.section);
        }

        // v31.3: EINE Zeilen-Darstellung für alle drei Fälle (mit Bereich, ohne
        // Bereich, ganz ohne Bereiche) — vorher stand derselbe Block dreimal im
        // JSX und lief bei jeder Änderung auseinander.
        const renderQuestion = (pq: (typeof perQuestion)[number], idx: number): React.ReactElement => {
          const pct = pq.answeredCount > 0 ? Math.round((pq.correctCount / pq.answeredCount) * 100) : 0;
          return (
            <div key={idx} className="dex-ui-card dex-ui-card--soft" style={{ padding: '8px 10px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6 }}>
                {pq.imageBase64 && (
                  <img src={pq.imageBase64} alt="" style={{ width: 60, height: 44, objectFit: 'cover', borderRadius: 6, flexShrink: 0, border: '1px solid var(--dex-gray-200)' }} />
                )}
                <span style={{ fontSize: '0.85rem', fontWeight: 500, flex: 1, minWidth: 0 }}>{idx + 1}. {pq.question}</span>
                <span className="dex-ui-muted" style={{ whiteSpace: 'nowrap', fontVariantNumeric: 'tabular-nums' }}>
                  {pq.correctCount} / {pq.answeredCount} richtig ({pct} %)
                </span>
              </div>
              <div className="dex-ui-progress">
                <div
                  className={cx('dex-ui-progress-bar', pct < 40 && 'dex-ui-progress-bar--red', pct >= 40 && pct < 70 && 'dex-ui-progress-bar--orange')}
                  style={{ width: `${pct}%` }}
                />
              </div>
            </div>
          );
        };

        return (
          <div className="dex-ui-card" style={{ padding: 12, marginBottom: 16 }}>
            {/* v31.3: Die Kopfzeile IST der Aufklapper — Hover über `dex-ui-row`,
                Zähler links beim Titel, Chevron rechts (Leitfaden 5a.8). */}
            <button
              type="button"
              onClick={() => setOpen(o => !o)}
              aria-expanded={open}
              className="dex-ui-rowbtn dex-ui-row dex-ui-card-head"
            >
              <span className="dex-ui-card-head-title"><FileText size={18} /> Quiz-Statistik</span>
              <span className="dex-ui-card-head-meta">
                {totalQuizzes === 0
                  ? 'Noch niemand teilgenommen'
                  : `${totalCompleted} abgeschlossen · ${totalQuizzes - totalCompleted} teilweise`}
              </span>
              <span className={cx('dex-ui-disclosure-chevron', open && 'is-open')} style={{ marginLeft: 'auto' }}>
                <ChevronDown size={18} />
              </span>
            </button>
            {open && (
            <div style={{ padding: '4px 6px 6px' }}>
              {totalQuizzes === 0 ? (
                <div className="dex-ui-empty">
                  <div className="dex-ui-empty-title">Noch kein Ergebnis</div>
                  Noch kein Teilnehmer hat das Quiz gestartet.
                </div>
              ) : (
                <>
                  <div className="dex-ui-kpi-row">
                    <div className="dex-ui-kpi dex-ui-kpi--green">
                      <div className="dex-ui-kpi-value">{totalCompleted}</div>
                      <div className="dex-ui-kpi-label">Abgeschlossen</div>
                    </div>
                    <div className="dex-ui-kpi dex-ui-kpi--orange">
                      <div className="dex-ui-kpi-value">{totalQuizzes - totalCompleted}</div>
                      <div className="dex-ui-kpi-label">Teilweise</div>
                    </div>
                    <div className="dex-ui-kpi">
                      <div className="dex-ui-kpi-value">{selectedEvent.quiz.length}</div>
                      <div className="dex-ui-kpi-label">Fragen</div>
                    </div>
                    <div className="dex-ui-kpi">
                      <div className="dex-ui-kpi-value">
                        {totalQuizzes > 0
                          ? (regsWithQuiz.reduce((sum, r) => sum + (r.QuizScore || 0), 0) / totalQuizzes).toFixed(1)
                          : '0'}
                      </div>
                      <div className="dex-ui-kpi-label">Ø Punkte</div>
                      <div className="dex-ui-kpi-sub">von {selectedEvent.quiz.length} möglichen</div>
                    </div>
                  </div>

                  {/* Pro Frage - gruppiert nach Bereich falls vorhanden */}
                  <div className="dex-ui-section">
                    <div className="dex-ui-section-title">Wie oft richtig beantwortet?</div>
                    {hasSections ? (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
                        {sectionsInOrder.map(sec => (
                          <div key={`stat-sec-${sec}`}>
                            <h5 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.92rem' }}>
                              Bereich: {sec}
                            </h5>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {perQuestion.map((pq, idx) => pq.section === sec ? renderQuestion(pq, idx) : null)}
                            </div>
                          </div>
                        ))}
                        {/* Fragen ohne Bereich */}
                        {perQuestion.some(pq => !pq.section) && (
                          <div>
                            <h5 style={{ margin: '0 0 6px', color: 'var(--dex-gray-600)', fontSize: '0.92rem' }}>Ohne Bereich</h5>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {perQuestion.map((pq, idx) => !pq.section ? renderQuestion(pq, idx) : null)}
                            </div>
                          </div>
                        )}
                      </div>
                    ) : (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                        {perQuestion.map((pq, idx) => renderQuestion(pq, idx))}
                      </div>
                    )}
                  </div>

                  {/* Top 10 */}
                  <div className="dex-ui-section">
                    <div className="dex-ui-section-title">Top 10 Teilnehmer</div>
                    <div className="dex-ui-table-wrap">
                      <table className="dex-ui-table">
                        <thead>
                          <tr>
                            <th style={{ width: 44 }}>#</th>
                            <th>Teilnehmer</th>
                            <th className="is-num" style={{ width: 90 }}>Punkte</th>
                            <th>Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {top10.map((reg, i) => {
                            const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
                            const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `${i + 1}.`;
                            const done = !!reg.QuizCompletedAt;
                            // Beantwortete Fragen zählen (für Partial)
                            let answeredN = 0;
                            try {
                              const parsed = JSON.parse(reg.QuizAnswers || '[]');
                              if (Array.isArray(parsed)) answeredN = parsed.filter((a: number[]) => Array.isArray(a) && a.length > 0).length;
                            } catch { /* */ }
                            return (
                              <tr key={reg.Id}>
                                <td style={{ fontWeight: 700 }}>{medal}</td>
                                {/* v31.3: Name und E-Mail sind EINE Personen-Zelle (Leitfaden 5b). */}
                                <td>
                                  <span className="dex-ui-person">
                                    <PersonContactHover email={reg.ParticipantEmail || ''} name={name || ''} size={26} />
                                    <span style={{ minWidth: 0 }}>
                                      <span className="dex-ui-person-name" style={{ display: 'block' }}>{name}</span>
                                      <span className="dex-ui-person-sub" style={{ display: 'block' }}>{reg.ParticipantEmail}</span>
                                    </span>
                                  </span>
                                </td>
                                <td className="is-num" style={{ fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)' }}>
                                  {reg.QuizScore ?? 0} / {selectedEvent.quiz.length}
                                </td>
                                <td>
                                  {done ? (
                                    <span className="dex-ui-inline">
                                      <span className="dex-ui-pill dex-ui-pill--green">Abgeschlossen</span>
                                      <span className="dex-ui-muted">{formatDate(reg.QuizCompletedAt || '')}</span>
                                    </span>
                                  ) : (
                                    <span className="dex-ui-pill dex-ui-pill--orange">
                                      Teilweise ({answeredN}/{selectedEvent.quiz.length})
                                    </span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </>
              )}
            </div>
            )}
          </div>
        );
};

