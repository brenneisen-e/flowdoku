/**
 * v30.13 — Modularisierung Stufe 3: Schritt „Fun-Zone" (Step 9, Index 8)
 * als eigene Komponente. JSX 1:1 aus EventCreationPage; die Helfer
 * (handleDrop, renderQuestionCard, Section-Ableitung) lebten schon vorher
 * in einer lokalen IIFE des Blocks und wandern unverändert mit. Von außen
 * kommen nur die Quiz-Liste samt Mutatoren, die Bereichs-/Drag-States und
 * die drei Setter des „Neuer Bereich"-Modals (das Modal selbst bleibt am
 * Wizard). `visible` ersetzt `currentStep === 8` — display:none statt
 * unmount, damit Eingaben beim Schrittwechsel erhalten bleiben.
 */
import * as React from 'react';
import { Plus, X } from '../../Icons';
import { StepBadge } from '../StepBadge';
import { useLanguage } from '../../../context/LanguageContext';
import { useDialog } from '../../../context/DialogContext';

export interface QuizQuestionDraft {
  id: string; question: string; options: string[];
  correctIndices: number[]; imageBase64?: string; section?: string;
}

export interface FunZoneStepProps {
  visible: boolean;
  quiz: QuizQuestionDraft[];
  addQuizQuestion: () => void;
  removeQuizQuestion: (id: string) => void;
  updateQuizQuestion: (id: string, updates: Partial<{ question: string; options: string[]; correctIndices: number[]; imageBase64: string | undefined; section: string | undefined }>) => void;
  pendingSections: string[];
  setPendingSections: React.Dispatch<React.SetStateAction<string[]>>;
  draggedQuestionId: string | null;
  setDraggedQuestionId: (v: string | null) => void;
  setNewSectionName: (v: string) => void;
  setNewSectionError: (v: string) => void;
  setNewSectionModalOpen: (v: boolean) => void;
  renderStepIntro: (bulletsDe: string[], bulletsEn: string[]) => React.ReactElement | null;
}

export const FunZoneStep: React.FC<FunZoneStepProps> = ({
  visible, quiz, addQuizQuestion, removeQuizQuestion, updateQuizQuestion,
  pendingSections, setPendingSections, draggedQuestionId, setDraggedQuestionId,
  setNewSectionName, setNewSectionError, setNewSectionModalOpen, renderStepIntro,
}) => {
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  const { confirmDialog, showAlert } = useDialog();
  return (
    <div style={{ display: visible ? 'block' : 'none' }}>
      <h2 className="dex-step-head-title">
        {isDe ? 'Schritt 9 — Fun-Zone' : 'Step 9 — Fun Zone'}
      </h2>
      <p className="dex-step-head-lead">
        {isDe
          ? 'Optional: ein Quiz für die Teilnehmer — Multiple-Choice-Fragen mit Live-Highscore. Perfekt für Networking, Tagungs-Pausen oder Foto-Quiz.'
          : 'Optional: a quiz for attendees — multiple-choice questions with live highscore. Perfect for networking, breaks at conferences, or photo quizzes.'}
      </p>
      {renderStepIntro(
        [
          'Quiz-Fragen für das Event anlegen — Multiple-Choice mit beliebig vielen Antwortoptionen',
          'Pro Frage optional ein Bild hochladen (Logo, Foto-Quiz, etc.)',
          'Mehrere richtige Antworten möglich (Mehrfachauswahl) — werden alle für volle Punktzahl gebraucht',
          'Bereiche anlegen und Fragen per Drag & Drop zuordnen — alle Fragen eines Bereichs werden im Quiz zusammen auf einer Seite angezeigt',
          'Live-Highscore + Statistik im Admin Center sehen (welche Fragen am häufigsten falsch beantwortet werden)',
        ],
        [
          'Create quiz questions for the event — multiple choice with any number of answer options',
          'Optionally upload an image per question (logo, photo quiz, etc.)',
          'Multiple correct answers are supported — all of them must be picked for full points',
          'Create sections and assign questions via drag & drop — all questions in a section are shown together on one page in the quiz',
          'See live highscore + statistics in the admin center (which questions are most often answered incorrectly)',
        ]
      )}
      <h3 className="mb-16">{t('create.step.funzone')}</h3>
      <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 16 }}>
        {t('create.funzone.hint')}
      </p>

      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
        <StepBadge n={34} />
        {isDe ? 'Quiz-Bereiche' : 'Quiz sections'}
      </label>
      {/* Bereiche: Header + "+ Bereich"-Button. Fragen können per Drag&Drop
          in Bereiche gezogen werden; jeder Bereich wird im Quiz zusammen
          auf einer Seite angezeigt. */}
      <div style={{
        display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12, flexWrap: 'wrap',
      }}>
        <button
          type="button"
          className="btn btn-outline"
          style={{ fontSize: '0.82rem', padding: '6px 14px' }}
          onClick={() => {
            setNewSectionName('');
            setNewSectionError('');
            setNewSectionModalOpen(true);
          }}
        >
          <Plus size={14} /> Bereich
        </button>
        <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
          Fragen per Drag &amp; Drop in einen Bereich ziehen — alle Fragen eines Bereichs werden im Quiz zusammen angezeigt.
        </span>
      </div>

      {(() => {
        // Section-Reihenfolge: zuerst die in Fragen verwendeten (nach erster Erwähnung),
        // dann die noch leeren pendingSections.
        const used: string[] = [];
        for (const q of quiz) {
          if (q.section && used.indexOf(q.section) < 0) used.push(q.section);
        }
        const allSections: string[] = [...used];
        for (const p of pendingSections) {
          if (allSections.indexOf(p) < 0) allSections.push(p);
        }

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const handleDrop = (ev: any, targetSection: string | undefined): void => {
          ev.preventDefault();
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const id = ev.dataTransfer?.getData?.('text/plain') as string | undefined;
          const qid = id || draggedQuestionId;
          if (!qid) return;
          updateQuizQuestion(qid, { section: targetSection });
          setDraggedQuestionId(null);
        };

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const renderQuestionCard = (q: any, qi: number): React.ReactElement => (
          <div
            key={q.id}
            draggable={true}
            onDragStart={ev => {
              setDraggedQuestionId(q.id);
              try { ev.dataTransfer.setData('text/plain', q.id); } catch { /* some browsers restrict */ }
              ev.dataTransfer.effectAllowed = 'move';
            }}
            onDragEnd={() => setDraggedQuestionId(null)}
            style={{
              padding: 16, marginBottom: 10, background: 'var(--dex-gray-50, #fafafa)',
              borderRadius: 12, border: '1px solid var(--dex-gray-200)',
              cursor: 'grab',
              opacity: draggedQuestionId === q.id ? 0.5 : 1,
            }}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <label style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)', display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ cursor: 'grab', color: 'var(--dex-gray-400)' }} title="Ziehen, um in einen Bereich zu verschieben">⋮⋮</span>
                {t('create.funzone.question')} {qi + 1}
              </label>
              <button type="button" onClick={() => removeQuizQuestion(q.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red)', padding: 4 }}>
                <X size={16} />
              </button>
            </div>
            <input
              className="form-input"
              value={q.question}
              onChange={e => updateQuizQuestion(q.id, { question: e.target.value })}
              placeholder={t('create.funzone.questionplaceholder')}
              style={{ marginBottom: 10 }}
            />
            <div style={{ marginBottom: 10, display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
              {q.imageBase64 ? (
                <>
                  <img
                    src={q.imageBase64}
                    alt="Frage-Bild"
                    style={{ maxHeight: 80, maxWidth: 160, borderRadius: 8, border: '1px solid var(--dex-gray-200)' }}
                  />
                  <button
                    type="button"
                    onClick={() => updateQuizQuestion(q.id, { imageBase64: undefined })}
                    style={{
                      fontSize: '0.72rem', padding: '4px 10px',
                      border: '1px solid var(--dex-gray-300)', borderRadius: 6,
                      background: '#fff', color: 'var(--dex-red)', cursor: 'pointer',
                    }}
                  >
                    Bild entfernen
                  </button>
                </>
              ) : (
                <label style={{
                  fontSize: '0.78rem', padding: '6px 12px',
                  border: '1px dashed var(--dex-gray-300)', borderRadius: 8,
                  cursor: 'pointer', color: 'var(--dex-gray-600)',
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                }}>
                  Bild hochladen (optional)
                  <input
                    type="file"
                    accept="image/*"
                    style={{ display: 'none' }}
                    onChange={async e => {
                      const file = e.target.files && e.target.files[0];
                      if (!file) return;
                      try {
                        const dataUrl = await new Promise<string>((resolve, reject) => {
                          const reader = new FileReader();
                          reader.onload = () => resolve(String(reader.result || ''));
                          reader.onerror = reject;
                          reader.readAsDataURL(file);
                        });
                        const img = new Image();
                        await new Promise<void>((resolve, reject) => {
                          img.onload = () => resolve();
                          img.onerror = reject;
                          img.src = dataUrl;
                        });
                        const maxW = 800;
                        const scale = img.width > maxW ? maxW / img.width : 1;
                        const w = Math.round(img.width * scale);
                        const h = Math.round(img.height * scale);
                        const canvas = document.createElement('canvas');
                        canvas.width = w;
                        canvas.height = h;
                        const ctx = canvas.getContext('2d');
                        if (!ctx) return;
                        ctx.drawImage(img, 0, 0, w, h);
                        const compressed = canvas.toDataURL('image/jpeg', 0.8);
                        updateQuizQuestion(q.id, { imageBase64: compressed });
                      } catch {
                        showAlert('Bild konnte nicht verarbeitet werden.');
                      }
                      e.target.value = '';
                    }}
                  />
                </label>
              )}
            </div>
            <label style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginBottom: 4, display: 'block' }}>
              {t('create.funzone.options')}
            </label>
            {q.options.map((opt: string, oi: number) => (
              <div key={oi} style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
                <input
                  type="checkbox"
                  checked={q.correctIndices?.includes(oi) || false}
                  onChange={() => {
                    const indices = q.correctIndices || [];
                    const newIndices = indices.includes(oi) ? indices.filter((x: number) => x !== oi) : [...indices, oi];
                    updateQuizQuestion(q.id, { correctIndices: newIndices.length > 0 ? newIndices : [0] });
                  }}
                  title={t('create.funzone.correct')}
                  style={{ accentColor: 'var(--dex-green)' }}
                />
                <input
                  className="form-input"
                  value={opt}
                  onChange={e => {
                    const newOpts = [...q.options];
                    newOpts[oi] = e.target.value;
                    updateQuizQuestion(q.id, { options: newOpts });
                  }}
                  placeholder={`${t('create.funzone.option')} ${oi + 1}`}
                  style={{ flex: 1, padding: '6px 10px', fontSize: '0.85rem' }}
                />
                {q.options.length > 2 && (
                  <button type="button" onClick={() => {
                    const newOpts = q.options.filter((_: string, i: number) => i !== oi);
                    const newCorrect = (q.correctIndices || []).filter((ci: number) => ci !== oi).map((ci: number) => ci > oi ? ci - 1 : ci);
                    updateQuizQuestion(q.id, { options: newOpts, correctIndices: newCorrect.length > 0 ? newCorrect : [0] });
                  }} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-gray-400)', padding: 2 }}>
                    <X size={14} />
                  </button>
                )}
              </div>
            ))}
            <button type="button" onClick={() => updateQuizQuestion(q.id, { options: [...q.options, ''] })} style={{
              fontSize: '0.78rem', padding: '4px 12px', border: '1px dashed var(--dex-gray-300)',
              borderRadius: 8, background: 'none', color: 'var(--dex-green-dark)', cursor: 'pointer', marginTop: 4,
            }}>
              + {t('create.funzone.addoption')}
            </button>
            <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-400)', marginTop: 6 }}>
              {t('create.funzone.correcthint')}
            </div>
          </div>
        );

        const unsortedQuiz = quiz.filter(q => !q.section);
        const globalIndexOf = (qid: string): number => quiz.findIndex(x => x.id === qid);

        return (
          <>
            {allSections.map(sec => {
              const inSec = quiz.filter(q => q.section === sec);
              return (
                <div
                  key={`sec-${sec}`}
                  onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; }}
                  onDrop={e => handleDrop(e, sec)}
                  style={{
                    padding: 12, marginBottom: 14, borderRadius: 12,
                    border: '2px dashed var(--dex-green)',
                    background: 'rgba(134,188,37,0.04)',
                  }}
                >
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10, gap: 8, flexWrap: 'wrap' }}>
                    <h4 style={{ margin: 0, color: 'var(--dex-green-dark, #4a7c1f)', fontSize: '1rem' }}>
                      Bereich: {sec} <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>({inSec.length} {inSec.length === 1 ? 'Frage' : 'Fragen'})</span>
                    </h4>
                    <button
                      type="button"
                      onClick={() => {
                        confirmDialog(isDe ? `Bereich "${sec}" entfernen? Die Fragen bleiben erhalten und landen in "Ohne Bereich".` : `Remove section "${sec}"? The questions are kept and move to "No section".`, { confirmLabel: isDe ? 'Entfernen' : 'Remove' }).then(ok => {
                          if (!ok) return;
                          for (const qq of quiz) {
                            if (qq.section === sec) updateQuizQuestion(qq.id, { section: undefined });
                          }
                          setPendingSections(prev => prev.filter(p => p !== sec));
                        }).catch(() => { /* */ });
                      }}
                      style={{
                        fontSize: '0.72rem', padding: '4px 10px',
                        border: '1px solid var(--dex-gray-300)', borderRadius: 6,
                        background: '#fff', color: 'var(--dex-red)', cursor: 'pointer',
                      }}
                    >
                      Bereich entfernen
                    </button>
                  </div>
                  {inSec.length === 0 ? (
                    <div style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', fontSize: '0.82rem', padding: '12px 8px', textAlign: 'center' }}>
                      Fragen hierher ziehen
                    </div>
                  ) : (
                    inSec.map(q => renderQuestionCard(q, globalIndexOf(q.id)))
                  )}
                </div>
              );
            })}

            {/* Ohne Bereich */}
            <div
              onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; }}
              onDrop={e => handleDrop(e, undefined)}
              style={allSections.length > 0 ? {
                padding: 12, marginBottom: 14, borderRadius: 12,
                border: '2px dashed var(--dex-gray-300)',
                background: 'var(--dex-gray-50, #fafafa)',
              } : undefined}
            >
              {allSections.length > 0 && (
                <h4 style={{ margin: '0 0 10px', color: 'var(--dex-gray-600)', fontSize: '0.95rem' }}>
                  Ohne Bereich <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>({unsortedQuiz.length} {unsortedQuiz.length === 1 ? 'Frage' : 'Fragen'})</span>
                </h4>
              )}
              {allSections.length > 0 && unsortedQuiz.length === 0 ? (
                <div style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', fontSize: '0.82rem', padding: '8px', textAlign: 'center' }}>
                  (leer)
                </div>
              ) : (
                unsortedQuiz.map(q => renderQuestionCard(q, globalIndexOf(q.id)))
              )}
            </div>
          </>
        );
      })()}

      <button type="button" className="btn btn-outline" onClick={addQuizQuestion} style={{ fontSize: '0.85rem', padding: '8px 20px' }}>
        <Plus size={14} /> {t('create.funzone.addquestion')}
      </button>
    </div>
  );
};
