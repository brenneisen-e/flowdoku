/**
 * v30.13 — Modularisierung Stufe 3: Schritt „Fun-Zone" (Step 9, Index 8)
 * als eigene Komponente. JSX 1:1 aus EventCreationPage; die Helfer
 * (handleDrop, renderQuestionCard, Section-Ableitung) lebten schon vorher
 * in einer lokalen IIFE des Blocks und wandern unverändert mit. Von außen
 * kommen nur die Quiz-Liste samt Mutatoren, die Bereichs-/Drag-States und
 * die drei Setter des „Neuer Bereich"-Modals (das Modal selbst bleibt am
 * Wizard). `visible` ersetzt `currentStep === 8` — display:none statt
 * unmount, damit Eingaben beim Schrittwechsel erhalten bleiben.
 *
 * v31.2: Modernisiert nach docs/ui-leitfaden.md — Fragen vor Bereichen, Karte
 * in Ausfüll-Reihenfolge, „Richtig" als Chip, Bereichs-Auswahl an der Frage
 * (Ergänzung zum Drag & Drop), alle Texte zweisprachig. Props/Verhalten gleich.
 */
import * as React from 'react';
import { Check, ImageIcon, Plus, Star, X } from '../../Icons';
import { cx } from '../../dexUi';
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
  // v31.2: Zähler-Text der Pills („3 Fragen") — eine Singular/Plural-Weiche statt drei.
  const countLabel = (n: number): string => `${n} ${n === 1 ? (isDe ? 'Frage' : 'question') : (isDe ? 'Fragen' : 'questions')}`;
  return (
    <div style={{ display: visible ? 'block' : 'none' }}>
      <h2 className="dex-step-head-title">
        <span className="dex-step-eyebrow">{isDe ? 'Schritt 9 von 9' : 'Step 9 of 9'}</span>
        {t('create.step.funzone')}
      </h2>
      <p className="dex-step-head-lead">
        {isDe
          ? <><strong>Optional</strong> — ein Quiz für die Teilnehmer: Multiple-Choice-Fragen mit Live-Highscore, gespielt unter &bdquo;Meine Events&ldquo;. Perfekt für Networking, Tagungs-Pausen oder ein Foto-Quiz.</>
          : <><strong>Optional</strong> — a quiz for attendees: multiple-choice questions with a live highscore, played under &bdquo;My Events&ldquo;. Perfect for networking, conference breaks or a photo quiz.</>}
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
      {/* v31.2: Erst die Fragen (der Kern), dann die Gliederung — bis v31.1 war die
          erste angebotene Handlung der optionale Bereichs-Knopf. h3 + Hint stehen im Kopf. */}
      <div className="dex-ui-section">
        <div className="dex-ui-section-title">
          {isDe ? 'Quiz-Fragen' : 'Quiz questions'}
          {quiz.length > 0 && (
            <span className="dex-ui-pill dex-ui-pill--gray" style={{ textTransform: 'none', letterSpacing: 0 }}>{countLabel(quiz.length)}</span>
          )}
        </div>
        <p className="dex-ui-section-desc">
          {isDe
            ? 'Jede Frage hat beliebig viele Antworten, mehrere richtige sind möglich. Live-Highscore und Statistik siehst du später im Organizer Center.'
            : 'Each question can have any number of answers, several correct ones are fine. You will see the live highscore and statistics in the Organizer Center later.'}
        </p>

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

        // v31.2: Karte in Ausfüll-Reihenfolge: Frage → Antworten → richtige markieren →
        // Bild/Bereich in der Fußzeile (der Upload stand vorher ZWISCHEN Frage und
        // Antworten). Bereich zusätzlich als Auswahl — Drag & Drop gibt es auf Touch nicht.
        const renderQuestionCard = (q: QuizQuestionDraft, qi: number): React.ReactElement => (
          <div
            key={q.id}
            className="dex-ui-card dex-ui-card--hover"
            draggable={true}
            onDragStart={ev => {
              setDraggedQuestionId(q.id);
              try { ev.dataTransfer.setData('text/plain', q.id); } catch { /* some browsers restrict */ }
              ev.dataTransfer.effectAllowed = 'move';
            }}
            onDragEnd={() => setDraggedQuestionId(null)}
            style={{ marginBottom: 10, opacity: draggedQuestionId === q.id ? 0.5 : 1 }}
          >
            <div className="dex-ui-inline" style={{ justifyContent: 'space-between', marginBottom: 10 }}>
              <div className="dex-ui-inline" style={{ gap: 6 }}>
                <span className="dex-ui-drag-handle" aria-hidden="true" title={isDe ? 'Ziehen, um die Frage in einen Bereich zu verschieben' : 'Drag to move the question into a section'}>⋮⋮</span>
                <span style={{ fontSize: '0.88rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>
                  {t('create.funzone.question')} {qi + 1}
                </span>
              </div>
              <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => removeQuizQuestion(q.id)}
                title={isDe ? 'Frage entfernen' : 'Remove question'} aria-label={isDe ? 'Frage entfernen' : 'Remove question'}>
                <X size={16} />
              </button>
            </div>
            <input
              className="form-input"
              value={q.question}
              onChange={e => updateQuizQuestion(q.id, { question: e.target.value })}
              placeholder={t('create.funzone.questionplaceholder')}
              style={{ marginBottom: 14 }}
            />
            <div className="dex-ui-label" style={{ marginBottom: 2 }}>
              {isDe ? 'Welche Antworten stehen zur Wahl?' : 'Which answers can be chosen?'}
            </div>
            <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 8 }}>
              {isDe
                ? 'Markiere jede richtige Antwort mit „Richtig" — auch mehrere. Für volle Punkte müssen Teilnehmer alle richtigen wählen.'
                : 'Mark every correct answer with "Correct" — several are fine. For full points, attendees must pick all correct ones.'}
            </div>
            {q.options.map((opt: string, oi: number) => {
              const isCorrect = q.correctIndices?.includes(oi) || false;
              return (
                <div key={oi} className="dex-ui-row" style={{ padding: '4px 6px', gap: 8, marginLeft: -6, marginRight: -6 }}>
                  <button
                    type="button"
                    className={cx('dex-ui-chip', isCorrect && 'is-active')}
                    aria-pressed={isCorrect}
                    title={t('create.funzone.correct')}
                    onClick={() => {
                      const indices = q.correctIndices || [];
                      const newIndices = indices.includes(oi) ? indices.filter((x: number) => x !== oi) : [...indices, oi];
                      updateQuizQuestion(q.id, { correctIndices: newIndices.length > 0 ? newIndices : [0] });
                    }}
                  >
                    <Check size={12} /> {isDe ? 'Richtig' : 'Correct'}
                  </button>
                  <input
                    className="dex-ui-input"
                    value={opt}
                    onChange={e => {
                      const newOpts = [...q.options];
                      newOpts[oi] = e.target.value;
                      updateQuizQuestion(q.id, { options: newOpts });
                    }}
                    placeholder={`${t('create.funzone.option')} ${oi + 1}`}
                    style={{ flex: 1 }}
                  />
                  {q.options.length > 2 && (
                    <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                      title={isDe ? 'Antwort entfernen' : 'Remove answer'} aria-label={isDe ? 'Antwort entfernen' : 'Remove answer'}
                      onClick={() => {
                        const newOpts = q.options.filter((_: string, i: number) => i !== oi);
                        const newCorrect = (q.correctIndices || []).filter((ci: number) => ci !== oi).map((ci: number) => ci > oi ? ci - 1 : ci);
                        updateQuizQuestion(q.id, { options: newOpts, correctIndices: newCorrect.length > 0 ? newCorrect : [0] });
                      }}
                    >
                      <X size={14} />
                    </button>
                  )}
                </div>
              );
            })}
            <button type="button" className="dex-ui-textbtn" onClick={() => updateQuizQuestion(q.id, { options: [...q.options, ''] })} style={{ marginTop: 4 }}>
              <Plus size={14} /> {t('create.funzone.addoption')}
            </button>
            <div className="dex-ui-divider" style={{ margin: '12px 0' }} />
            <div className="dex-ui-inline" style={{ justifyContent: 'space-between' }}>
              <div className="dex-ui-inline">
                {q.imageBase64 ? (
                  <>
                    <img
                      src={q.imageBase64}
                      alt={isDe ? 'Frage-Bild' : 'Question image'}
                      style={{ maxHeight: 64, maxWidth: 140, borderRadius: 8, border: '1px solid var(--dex-gray-200)' }}
                    />
                    <button type="button" className="dex-ui-textbtn dex-ui-textbtn--danger" onClick={() => updateQuizQuestion(q.id, { imageBase64: undefined })}>
                      <X size={14} /> {isDe ? 'Bild entfernen' : 'Remove image'}
                    </button>
                  </>
                ) : (
                  <label className="dex-ui-textbtn dex-ui-textbtn--muted" title={isDe ? 'Erscheint über der Frage im Quiz' : 'Shown above the question in the quiz'}>
                    <ImageIcon size={14} /> {isDe ? 'Bild zur Frage (optional)' : 'Image for the question (optional)'}
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
                          showAlert(isDe ? 'Bild konnte nicht verarbeitet werden.' : 'The image could not be processed.');
                        }
                        e.target.value = '';
                      }}
                    />
                  </label>
                )}
              </div>
              {allSections.length > 0 && (
                <label className="dex-ui-inline" style={{ gap: 6, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                  {isDe ? 'Bereich' : 'Section'}
                  <select
                    className="dex-ui-select dex-ui-input--sm"
                    value={q.section || ''}
                    onChange={e => updateQuizQuestion(q.id, { section: e.target.value || undefined })}
                    style={{ width: 'auto', minWidth: 150, paddingRight: 28 }}
                  >
                    <option value="">{isDe ? 'Ohne Bereich' : 'No section'}</option>
                    {allSections.map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
                </label>
              )}
            </div>
          </div>
        );

        const unsortedQuiz = quiz.filter(q => !q.section);
        const globalIndexOf = (qid: string): number => quiz.findIndex(x => x.id === qid);
        // v31.2: Ablageflächen sind die zentrale dex-ui-dropzone (gestrichelt, „hier kann
        // etwas hinein"); während eine Frage gezogen wird, tragen ALLE Ziele is-over —
        // vorher sah man das nicht. Die Klasse zentriert für Datei-Ablagen; hier liegen
        // Kopfzeile und Fragekarten drin, deshalb Layout auf Block/links zurückgestellt.
        const dragging = !!draggedQuestionId;
        const zoneClass = cx('dex-ui-dropzone', dragging && 'is-over');
        const zoneStyle: React.CSSProperties = {
          alignItems: 'stretch', textAlign: 'left', cursor: 'default', gap: 0,
          padding: 12, marginBottom: 12,
        };

        return (
          <>
            {allSections.map(sec => {
              const inSec = quiz.filter(q => q.section === sec);
              return (
                <div
                  key={`sec-${sec}`}
                  onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; }}
                  onDrop={e => handleDrop(e, sec)}
                  className={zoneClass}
                  style={zoneStyle}
                >
                  <div className="dex-ui-inline" style={{ justifyContent: 'space-between', marginBottom: 10 }}>
                    <div className="dex-ui-inline">
                      <h4 style={{ margin: 0, color: 'var(--dex-green-darker, #4a7c1f)', fontSize: '0.95rem' }}>
                        {isDe ? 'Bereich' : 'Section'}: {sec}
                      </h4>
                      <span className="dex-ui-pill dex-ui-pill--green">{countLabel(inSec.length)}</span>
                    </div>
                    <button
                      type="button"
                      className="dex-ui-textbtn dex-ui-textbtn--danger"
                      onClick={() => {
                        confirmDialog(isDe ? `Bereich "${sec}" entfernen? Die Fragen bleiben erhalten und landen in "Ohne Bereich".` : `Remove section "${sec}"? The questions are kept and move to "No section".`, { confirmLabel: isDe ? 'Entfernen' : 'Remove' }).then(ok => {
                          if (!ok) return;
                          for (const qq of quiz) {
                            if (qq.section === sec) updateQuizQuestion(qq.id, { section: undefined });
                          }
                          setPendingSections(prev => prev.filter(p => p !== sec));
                        }).catch(() => { /* */ });
                      }}
                    >
                      <X size={14} /> {isDe ? 'Bereich entfernen' : 'Remove section'}
                    </button>
                  </div>
                  {inSec.length === 0 ? (
                    <div className="dex-ui-muted" style={{ fontStyle: 'italic', padding: '12px 8px', textAlign: 'center' }}>
                      {isDe ? 'Fragen am Griff hierher ziehen — oder unten an der Frage den Bereich wählen.' : 'Drag questions here by their handle — or pick the section at the bottom of a question.'}
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
              className={allSections.length > 0 ? zoneClass : undefined}
              style={allSections.length > 0 ? zoneStyle : undefined}
            >
              {allSections.length > 0 && (
                <div className="dex-ui-inline" style={{ marginBottom: 10 }}>
                  <h4 style={{ margin: 0, color: 'var(--dex-gray-600)', fontSize: '0.95rem' }}>{isDe ? 'Ohne Bereich' : 'No section'}</h4>
                  <span className="dex-ui-pill dex-ui-pill--gray">{countLabel(unsortedQuiz.length)}</span>
                </div>
              )}
              {allSections.length > 0 && unsortedQuiz.length === 0 ? (
                <div className="dex-ui-muted" style={{ fontStyle: 'italic', padding: 8, textAlign: 'center' }}>
                  {isDe ? '(leer)' : '(empty)'}
                </div>
              ) : (
                unsortedQuiz.map(q => renderQuestionCard(q, globalIndexOf(q.id)))
              )}
            </div>

            {/* v31.2: Leerer Zustand statt leerer Fläche — man soll sehen, was der erste Klick bewirkt. */}
            {quiz.length === 0 && allSections.length === 0 && (
              <div className="dex-ui-empty" style={{ marginBottom: 12 }}>
                <span className="dex-ui-empty-icon"><Star size={20} /></span>
                <div className="dex-ui-empty-title">{isDe ? 'Noch keine Quiz-Frage' : 'No quiz question yet'}</div>
                <div>
                  {isDe
                    ? 'Leg die erste Frage an — Multiple-Choice, gern mit Bild. Ohne Fragen sehen Teilnehmer keine Fun-Zone.'
                    : 'Create the first question — multiple choice, with an image if you like. Without questions, attendees see no Fun Zone.'}
                </div>
              </div>
            )}
          </>
        );
      })()}

      <button type="button" className="btn btn-outline dex-ui-btn-sm" onClick={addQuizQuestion}>
        <Plus size={14} /> {t('create.funzone.addquestion')}
      </button>
      </div>

      {/* v31.2: Bereiche als optionaler Abschnitt HINTER den Fragen — erst was gefragt wird,
          dann wie es gegliedert wird. Der Satz nennt, WO der neue Bereich auftaucht (oben). */}
      <div className="dex-ui-section">
        <div className="dex-ui-section-title">
          <StepBadge n={34} />
          {isDe ? 'Quiz in Bereiche gliedern' : 'Group the quiz into sections'}
          <span className="dex-ui-label-optional" style={{ textTransform: 'none', letterSpacing: 0 }}>{isDe ? '(optional)' : '(optional)'}</span>
        </div>
        <p className="dex-ui-section-desc">
          {isDe
            ? <>Alle Fragen eines Bereichs erscheinen im Quiz zusammen auf einer Seite. Ein neuer Bereich taucht oben in der Fragenliste als Ablagefläche auf — zieh Fragen am Griff <span aria-hidden="true">⋮⋮</span> hinein oder wähl den Bereich direkt an der Frage.</>
            : <>All questions of a section are shown together on one page in the quiz. A new section appears at the top of the question list as a drop area — drag questions in by their handle <span aria-hidden="true">⋮⋮</span> or pick the section right on the question.</>}
        </p>
        <button type="button" className="btn btn-outline dex-ui-btn-sm"
          onClick={() => { setNewSectionName(''); setNewSectionError(''); setNewSectionModalOpen(true); }}>
          <Plus size={14} /> {isDe ? 'Bereich anlegen' : 'Create section'}
        </button>
      </div>
    </div>
  );
};
