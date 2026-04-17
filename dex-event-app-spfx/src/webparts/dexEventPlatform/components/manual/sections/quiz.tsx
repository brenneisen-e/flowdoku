import * as React from 'react';
import { ManualSection } from '../types';
import { DemoQuizQuestion, Callout } from '../ManualMockups';

export function quizSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'quiz',
    title: isDe ? 'Quiz / Fun-Zone' : 'Quiz / Fun-Zone',
    category: 'organizer',
    description: isDe
      ? 'Ein Quiz für dein Event erstellen und Ergebnisse auswerten.'
      : 'Create a quiz for your event and review the results.',
    visibleFor: ['Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        title: isDe ? 'Als Organizer:in' : 'As organizer',
        steps: [
          {
            number: 1,
            title: isDe ? 'Quiz-Fragen anlegen' : 'Create quiz questions',
            description: (
              <>
                {isDe
                  ? 'Im Event-Wizard Schritt 7 "Fun-Zone" fügst du Fragen hinzu. Pro Frage: Fragetext + 4 Antwortoptionen + Markierung der richtigen Antwort. Reihenfolge per Drag-and-Drop.'
                  : 'In wizard step 7 "Fun-Zone", add questions. Per question: question text + 4 answer options + mark the correct one. Reorder via drag-and-drop.'}
              </>
            ),
            mockup: <DemoQuizQuestion />,
          },
          {
            number: 2,
            title: isDe ? 'Quiz aktivieren' : 'Activate the quiz',
            description: (
              <>
                {isDe
                  ? 'Quiz werden automatisch aktiv, sobald mindestens eine Frage gespeichert ist. Teilnehmer sehen das Quiz auf der Event-Details-Seite und können es beliebig oft spielen — gewertet wird nur der beste Durchlauf.'
                  : 'Quizzes become active as soon as at least one question is saved. Attendees see the quiz on the event details page and can play as often as they like — only the best attempt counts.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Ergebnisse einsehen' : 'Review the results',
            description: (
              <>
                {isDe
                  ? 'Im Admin Center des Events wird pro Teilnehmer die höchste Punktzahl angezeigt. Nutze das für Gewinn-Verlosungen oder Engagement-Tracking.'
                  : 'The admin center shows each attendee\'s highest score. Use this for giveaways or engagement tracking.'}
              </>
            ),
            mockup: (
              <Callout variant="tip">
                {isDe
                  ? 'Quiz-Stand wird im Teilnehmer-Item gespeichert (Spalte "QuizScore") und beim Spalten-Fix automatisch angelegt, falls fehlend.'
                  : 'Quiz state is stored in the attendee item (column "QuizScore") and auto-created by the column-fix tool if missing.'}
              </Callout>
            ),
          },
        ],
      },
      {
        perspective: 'user',
        title: isDe ? 'Als Teilnehmer:in' : 'As attendee',
        steps: [
          {
            number: 1,
            title: isDe ? 'Quiz starten' : 'Start the quiz',
            description: (
              <>
                {isDe
                  ? 'Auf der Event-Details-Seite findest du den Button "Quiz starten". Die Fragen erscheinen nacheinander.'
                  : 'On the event details page you\'ll find a "Start quiz" button. Questions appear one by one.'}
              </>
            ),
            mockup: <DemoQuizQuestion showAnswered />,
          },
          {
            number: 2,
            title: isDe ? 'Noch mal versuchen' : 'Try again',
            description: (
              <>
                {isDe
                  ? 'Am Ende siehst du deinen Score. Du kannst das Quiz beliebig oft wiederholen — gewertet wird dein bester Versuch.'
                  : 'At the end you see your score. You can replay as often as you like — only your best attempt counts.'}
              </>
            ),
          },
        ],
      },
    ],
  };
}
