/* QuizStatsSection — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 9624-9879 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { FileText } from '../../Icons';
import { formatDate } from '../../../utils/eventStatus';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';

export interface QuizStatsSectionProps {
  registrations: SPRegistration[];
  selectedEvent: DeloitteEvent;
}

export const QuizStatsSection: React.FC<QuizStatsSectionProps> = (p) => {
  const { registrations, selectedEvent } = p;
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

        return (
          <details className="card" style={{ padding: 0, marginBottom: 16 }}>
            <summary style={{
              padding: '16px 24px', cursor: 'pointer', listStyle: 'none',
              display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12,
              fontSize: '1rem', fontWeight: 600,
            }}>
              <span style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <FileText size={18} /> Quiz-Statistik
              </span>
              <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>
                {totalQuizzes === 0
                  ? 'Keine Daten'
                  : `${totalCompleted} abgeschlossen, ${totalQuizzes - totalCompleted} teilweise (Klick zum Ausklappen)`}
              </span>
            </summary>
            <div style={{ padding: '0 24px 24px 24px' }}>
              {totalQuizzes === 0 ? (
                <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', margin: 0 }}>
                  Noch kein Teilnehmer hat das Quiz gestartet.
                </p>
              ) : (
                <>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 12, marginBottom: 24 }}>
                    <div style={{ padding: 16, background: 'var(--dex-green-light, #f0fdf4)', borderRadius: 12, textAlign: 'center' }}>
                      <div style={{ fontSize: '1.6rem', fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)' }}>{totalCompleted}</div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Abgeschlossen</div>
                    </div>
                    <div style={{ padding: 16, background: 'var(--dex-orange-light, #fff7ed)', borderRadius: 12, textAlign: 'center' }}>
                      <div style={{ fontSize: '1.6rem', fontWeight: 700, color: 'var(--dex-orange, #ed8b00)' }}>{totalQuizzes - totalCompleted}</div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Teilweise</div>
                    </div>
                    <div style={{ padding: 16, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, textAlign: 'center' }}>
                      <div style={{ fontSize: '1.6rem', fontWeight: 700 }}>{selectedEvent.quiz.length}</div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Fragen</div>
                    </div>
                    <div style={{ padding: 16, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, textAlign: 'center' }}>
                      <div style={{ fontSize: '1.6rem', fontWeight: 700 }}>
                        {totalQuizzes > 0
                          ? (regsWithQuiz.reduce((sum, r) => sum + (r.QuizScore || 0), 0) / totalQuizzes).toFixed(1)
                          : '0'}
                      </div>
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>Ø Score</div>
                    </div>
                  </div>

                  {/* Pro Frage - gruppiert nach Bereich falls vorhanden */}
                  <h4 style={{ marginTop: 0, marginBottom: 12 }}>Pro Frage</h4>
                  {(() => {
                    const hasSections = perQuestion.some(pq => !!pq.section);
                    if (!hasSections) return null;
                    // Gruppen in Reihenfolge der ersten Erwähnung
                    const sectionsInOrder: string[] = [];
                    for (const pq of perQuestion) {
                      if (pq.section && sectionsInOrder.indexOf(pq.section) < 0) sectionsInOrder.push(pq.section);
                    }
                    return (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 16, marginBottom: 24 }}>
                        {sectionsInOrder.map(sec => (
                          <div key={`stat-sec-${sec}`}>
                            <h5 style={{ margin: '0 0 6px', color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '0.92rem' }}>
                              Bereich: {sec}
                            </h5>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {perQuestion.map((pq, idx) => pq.section === sec ? (() => {
                                const pct = pq.answeredCount > 0 ? Math.round((pq.correctCount / pq.answeredCount) * 100) : 0;
                                return (
                                  <div key={idx} style={{ padding: 10, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 8 }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4, gap: 12 }}>
                                      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flex: 1 }}>
                                        {pq.imageBase64 && (
                                          <img src={pq.imageBase64} alt="" style={{ width: 60, height: 44, objectFit: 'cover', borderRadius: 6, flexShrink: 0, border: '1px solid var(--dex-gray-200)' }} />
                                        )}
                                        <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>{idx + 1}. {pq.question}</span>
                                      </div>
                                      <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
                                        {pq.correctCount} / {pq.answeredCount} richtig ({pct}%)
                                      </span>
                                    </div>
                                    <div style={{ height: 6, background: 'var(--dex-gray-200)', borderRadius: 3, overflow: 'hidden' }}>
                                      <div style={{ height: '100%', width: `${pct}%`, background: pct >= 70 ? 'var(--dex-green, #86bc25)' : pct >= 40 ? 'var(--dex-orange, #ff8c00)' : 'var(--dex-red, #c00)', transition: 'width 0.3s' }} />
                                    </div>
                                  </div>
                                );
                              })() : null)}
                            </div>
                          </div>
                        ))}
                        {/* Fragen ohne Bereich */}
                        {perQuestion.some(pq => !pq.section) && (
                          <div>
                            <h5 style={{ margin: '0 0 6px', color: 'var(--dex-gray-600)', fontSize: '0.92rem' }}>Ohne Bereich</h5>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {perQuestion.map((pq, idx) => !pq.section ? (() => {
                                const pct = pq.answeredCount > 0 ? Math.round((pq.correctCount / pq.answeredCount) * 100) : 0;
                                return (
                                  <div key={idx} style={{ padding: 10, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 8 }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4, gap: 12 }}>
                                      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flex: 1 }}>
                                        {pq.imageBase64 && (
                                          <img src={pq.imageBase64} alt="" style={{ width: 60, height: 44, objectFit: 'cover', borderRadius: 6, flexShrink: 0, border: '1px solid var(--dex-gray-200)' }} />
                                        )}
                                        <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>{idx + 1}. {pq.question}</span>
                                      </div>
                                      <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
                                        {pq.correctCount} / {pq.answeredCount} richtig ({pct}%)
                                      </span>
                                    </div>
                                    <div style={{ height: 6, background: 'var(--dex-gray-200)', borderRadius: 3, overflow: 'hidden' }}>
                                      <div style={{ height: '100%', width: `${pct}%`, background: pct >= 70 ? 'var(--dex-green, #86bc25)' : pct >= 40 ? 'var(--dex-orange, #ff8c00)' : 'var(--dex-red, #c00)', transition: 'width 0.3s' }} />
                                    </div>
                                  </div>
                                );
                              })() : null)}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })()}
                  {!perQuestion.some(pq => !!pq.section) && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 24 }}>
                    {perQuestion.map((pq, idx) => {
                      const pct = pq.answeredCount > 0 ? Math.round((pq.correctCount / pq.answeredCount) * 100) : 0;
                      return (
                        <div key={idx} style={{ padding: 10, background: 'var(--dex-gray-50, #fafafa)', borderRadius: 8 }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4, gap: 12 }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 10, flex: 1 }}>
                              {pq.imageBase64 && (
                                <img
                                  src={pq.imageBase64}
                                  alt=""
                                  style={{ width: 60, height: 44, objectFit: 'cover', borderRadius: 6, flexShrink: 0, border: '1px solid var(--dex-gray-200)' }}
                                />
                              )}
                              <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>
                                {idx + 1}. {pq.question}
                              </span>
                            </div>
                            <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>
                              {pq.correctCount} / {pq.answeredCount} richtig ({pct}%)
                            </span>
                          </div>
                          <div style={{ height: 6, background: 'var(--dex-gray-200)', borderRadius: 3, overflow: 'hidden' }}>
                            <div style={{
                              height: '100%',
                              width: `${pct}%`,
                              background: pct >= 70 ? 'var(--dex-green, #86bc25)' : pct >= 40 ? 'var(--dex-orange, #ff8c00)' : 'var(--dex-red, #c00)',
                              transition: 'width 0.3s',
                            }} />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  )}

                  {/* Top 10 */}
                  <h4 style={{ marginTop: 0, marginBottom: 12 }}>Top 10 Teilnehmer</h4>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                          <th style={{ textAlign: 'left', padding: 8, width: 40 }}>#</th>
                          <th style={{ textAlign: 'left', padding: 8 }}>Name</th>
                          <th style={{ textAlign: 'left', padding: 8 }}>E-Mail</th>
                          <th style={{ textAlign: 'left', padding: 8, width: 80 }}>Score</th>
                          <th style={{ textAlign: 'left', padding: 8 }}>Status</th>
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
                            <tr key={reg.Id} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                              <td style={{ padding: 8, fontWeight: 700 }}>{medal}</td>
                              <td style={{ padding: 8, fontWeight: 500 }}>{name}</td>
                              <td style={{ padding: 8, color: 'var(--dex-gray-600)' }}>{reg.ParticipantEmail}</td>
                              <td style={{ padding: 8, fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)' }}>
                                {reg.QuizScore ?? 0} / {selectedEvent.quiz.length}
                              </td>
                              <td style={{ padding: 8, color: done ? 'var(--dex-gray-500)' : 'var(--dex-orange, #ed8b00)' }}>
                                {done
                                  ? `Abgeschlossen ${formatDate(reg.QuizCompletedAt || '')}`
                                  : `Teilweise (${answeredN}/${selectedEvent.quiz.length})`}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          </details>
        );
};

