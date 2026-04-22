import * as React from 'react';
import { ManualSection } from '../types';
import {
  DemoQuizQuestion, DemoQuizEditor, DemoQuizResult, DemoQuizScoreboard, DemoQuizStart,
  DemoWizardProgress, Callout, ClickPath, NumberedList,
} from '../ManualMockups';

export function quizSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'quiz',
    title: isDe ? 'Quiz / Fun-Zone' : 'Quiz / Fun-Zone',
    category: 'organizer',
    description: isDe
      ? 'Quiz-Fragen erstellen, Teilnehmer spielen lassen und Ergebnisse auswerten.'
      : 'Create quiz questions, let attendees play, and review the results.',
    visibleFor: ['User', 'Organizer', 'Admin'],
    perspectives: [
      {
        perspective: 'organizer',
        title: isDe ? 'Quiz erstellen & veröffentlichen' : 'Create & publish the quiz',
        intro: (
          <>
            {isDe
              ? 'Das Quiz ist Teil von Schritt 7 "Fun-Zone" im Event-Wizard. Du kannst beliebig viele Fragen mit je 4 Antwortoptionen anlegen und eine davon als richtig markieren. Teilnehmer spielen das Quiz auf der Event-Details-Seite — sobald mindestens eine Frage gespeichert ist, erscheint der Start-Button automatisch.'
              : 'The quiz lives in wizard step 7 "Fun-Zone". You can add any number of questions, each with 4 answer options, and mark one as correct. Attendees play the quiz on the event details page — as soon as at least one question is saved, the start button appears automatically.'}
          </>
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Wizard Schritt 7 öffnen' : 'Open wizard step 7',
            description: (
              <>
                {isDe
                  ? 'Nach den üblichen Schritten (Grundlagen, Zeit & Ort, Kapazität, Felder, Kommunikation, Dokumente) erreichst du den letzten Schritt "Fun-Zone". Standardmäßig ist noch keine Frage angelegt — das Quiz ist also optional.'
                  : 'After the usual steps (Basics, Date & Location, Capacity, Fields, Communication, Documents) you reach the final step "Fun-Zone". By default no question is set — the quiz is optional.'}
              </>
            ),
            mockup: (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                <DemoWizardProgress activeStep={6} />
                <ClickPath label={isDe ? 'Schritt 7' : 'Step 7'} label2={isDe ? 'Frage hinzufügen' : 'Add question'} />
              </div>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Fragetext & Antworten pflegen' : 'Write question & answers',
            description: (
              <>
                <p style={{ margin: '0 0 8px 0' }}>
                  {isDe
                    ? 'Pro Frage trägst du den Fragetext und bis zu 4 Antwortoptionen ein. Mit einem Klick auf den leeren Kreis links markierst du eine Option als richtig (grüner Haken). Fragen lassen sich per Drag-and-Drop verschieben.'
                    : 'Per question, enter the question text and up to 4 answer options. Click the empty circle on the left to mark one as correct (green check). Questions can be reordered via drag-and-drop.'}
                </p>
                <NumberedList items={[
                  isDe ? 'Fragetext eingeben (Pflicht).' : 'Enter the question text (required).',
                  isDe ? 'Antwortoptionen A–D ausfüllen (mindestens 2 müssen gefüllt sein).' : 'Fill in options A–D (at least 2 must be filled).',
                  isDe ? 'Genau eine Option als richtig markieren.' : 'Mark exactly one option as correct.',
                  isDe ? '"Nächste Frage" klicken, um weitere Fragen zu ergänzen.' : 'Click "Next question" to add more.',
                ]} />
              </>
            ),
            mockup: <DemoQuizEditor />,
            warning: isDe
              ? 'Ohne markierte richtige Antwort wird die Frage beim Speichern ignoriert. Die Teilnehmer sehen sie nicht.'
              : 'Without a marked correct answer, the question is skipped when saving. Attendees won\'t see it.',
          },
          {
            number: 3,
            title: isDe ? 'Reihenfolge anpassen & Fragen löschen' : 'Reorder & delete questions',
            description: (
              <>
                {isDe
                  ? 'Jede Frage hat einen Drag-Griff (⋮⋮) auf der linken Seite — damit ziehst du sie an eine andere Position. Das rote Mülleimer-Icon rechts löscht die Frage unwiderruflich (kein Papierkorb).'
                  : 'Each question has a drag handle (⋮⋮) on the left — drag it to reposition. The red trash icon on the right permanently deletes the question (no undo).'}
              </>
            ),
            tip: isDe
              ? 'Reihenfolge-Tipp: Leichte Fragen zuerst (Motivation), kniffelige in der Mitte, humorvolle zum Abschluss.'
              : 'Ordering tip: easy questions first (to motivate), tricky ones in the middle, something fun at the end.',
          },
          {
            number: 4,
            title: isDe ? 'Mehrere richtige Antworten pro Frage' : 'Multiple correct answers per question',
            description: (
              <>
                {isDe
                  ? 'Seit v5.x erlaubt der Editor mehrere richtige Antworten pro Frage. Setze einfach mehrere Haken im Feld "Richtige Antwort(en)" — der Teilnehmer muss dann alle richtigen Antworten treffen (und keine falschen). Gut für Wissensfragen mit "wähle alle passenden Antworten"-Charakter. Intern wird das als correctIndices-Array gespeichert.'
                  : 'Since v5.x the editor allows multiple correct answers per question. Just tick multiple checkboxes in the "Correct answer(s)" field — the participant must then pick all correct answers (and none wrong). Great for knowledge questions with "pick all that apply" style. Stored internally as correctIndices array.'}
              </>
            ),
          },
          {
            number: 5,
            title: isDe ? 'Bild zur Frage hinzufügen' : 'Add image to question',
            description: (
              <>
                {isDe
                  ? 'Pro Frage kann ein Bild als JPG/PNG hochgeladen werden. Die App komprimiert es automatisch (empfohlen: max. ~80 KB final, daher reicht normale Handyfoto-Qualität nach Komprimierung). Das Bild wird als data:image/... Base64-URL inline im FunZone-JSON des Events gespeichert — keine externe Storage-Referenz, kein Broken-Image-Risiko. Im Quiz-Player wird es oberhalb des Fragetextes angezeigt. Gut für "Welches Logo siehst du?" oder "Welches Foto zeigt Event XYZ?"-Fragen.'
                  : 'A JPG/PNG image can be uploaded per question. The app compresses it automatically (recommended: max ~80 KB final — standard phone photo quality after compression is plenty). The image is stored inline as data:image/... base64 URL in the event\'s FunZone JSON — no external storage reference, no broken-image risk. In the quiz player it is displayed above the question text. Great for "Which logo is this?" or "Which photo shows event XYZ?"-style questions.'}
              </>
            ),
          },
          {
            number: 6,
            title: isDe ? 'Fragen clustern (Cluster-Size)' : 'Cluster questions (cluster size)',
            description: (
              <>
                {isDe
                  ? 'Normalerweise zeigt der Quiz-Player eine Frage pro Bildschirm. Für längere Quizzes kannst du die Cluster-Size auf 2, 3 oder 4 setzen — dann werden mehrere Fragen auf einer Seite zusammen angezeigt. Praktisch, wenn die Fragen thematisch zusammenhängen oder der User schneller durchspielen soll. Default ist 1.'
                  : 'Normally the quiz player shows one question per screen. For longer quizzes you can set the cluster size to 2, 3 or 4 — multiple questions then appear on one page together. Useful when questions are thematically related or you want the user to play through faster. Default is 1.'}
              </>
            ),
          },
          {
            number: 7,
            title: isDe ? 'Bereiche (Sections) zur thematischen Gliederung' : 'Sections for thematic grouping',
            description: (
              <>
                {isDe
                  ? 'Große Quizzes kannst du per "+ Bereich"-Button in thematische Sections gliedern. Fragen mit gleichem Bereich-Namen werden im Quiz-Player automatisch auf einer gemeinsamen Seite angezeigt (überschreibt die Cluster-Size innerhalb des Bereichs). So kannst du z.B. einen Block "Deloitte-Wissen", einen Block "Event-Locations" und einen Block "Fun-Fakten" bauen, die jeweils als eigene Section erscheinen.'
                  : 'For large quizzes you can use the "+ Section" button to group questions into thematic sections. Questions with the same section name appear together on one page in the quiz player (overriding the cluster size within that section). E.g. you can build a "Deloitte knowledge" block, an "Event locations" block and a "Fun facts" block, each as its own section.'}
              </>
            ),
          },
          {
            number: 8,
            title: isDe ? 'Quiz veröffentlichen = speichern' : 'Publish = save',
            description: (
              <>
                {isDe
                  ? 'Beim Speichern des Events werden alle Quiz-Fragen automatisch mit veröffentlicht. Es gibt keinen separaten "Veröffentlichen"-Schalter. Ist mindestens eine gültige Frage vorhanden, zeigt die Event-Details-Seite den Teilnehmern einen "Quiz starten"-Button.'
                  : 'Saving the event automatically publishes all quiz questions. There\'s no separate "Publish" toggle. If at least one valid question exists, the event details page shows attendees a "Start quiz" button.'}
              </>
            ),
            mockup: (
              <Callout variant="success" title={isDe ? 'Live-Änderungen möglich' : 'Live changes supported'}>
                {isDe
                  ? 'Du kannst Fragen auch nach der Veröffentlichung noch ändern, hinzufügen oder löschen. Bereits erzielte Punktzahlen bleiben erhalten, aber die Punktzahl bezieht sich dann auf die aktuelle Fragen-Anzahl.'
                  : 'You can still edit, add, or remove questions after publishing. Already-achieved scores remain, but they now relate to the current question count.'}
              </Callout>
            ),
          },
          {
            number: 5,
            title: isDe ? 'Vorschau selbst spielen' : 'Play the preview yourself',
            description: (
              <>
                {isDe
                  ? 'Als Organizer kannst du das Quiz selbst im Event öffnen und testen. Deine Versuche werden im Admin Center getrennt markiert, damit sie die Statistik nicht verfälschen.'
                  : 'As an organizer you can open your event and play the quiz yourself to test it. Your attempts are flagged separately in the admin center so they don\'t distort the statistics.'}
              </>
            ),
            mockup: <DemoQuizQuestion showAnswered />,
          },
        ],
      },
      {
        perspective: 'user',
        title: isDe ? 'Quiz spielen' : 'Play the quiz',
        intro: (
          <>
            {isDe
              ? 'Wenn der Organizer ein Quiz eingebaut hat, siehst du es auf der Event-Details-Seite. Du kannst es so oft spielen, wie du willst — gewertet wird immer dein bester Versuch.'
              : 'If the organizer added a quiz, you see it on the event details page. Play as often as you like — only your best attempt counts.'}
          </>
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Start-Screen' : 'Start screen',
            description: (
              <>
                {isDe
                  ? 'Auf der Event-Seite findest du eine "Quiz starten"-Kachel mit kurzer Info, wie viele Fragen dich erwarten und dass es keine Zeitbegrenzung gibt.'
                  : 'On the event page you\'ll find a "Start quiz" tile with a quick info on how many questions await you and that there\'s no time limit.'}
              </>
            ),
            mockup: <DemoQuizStart />,
          },
          {
            number: 2,
            title: isDe ? 'Fragen beantworten' : 'Answer questions',
            description: (
              <>
                {isDe
                  ? 'Die Fragen erscheinen einzeln. Mit Klick auf eine Antwort wird sie ausgewählt und du bekommst sofort Feedback: grün = richtig, rot = falsch. Danach klickst du auf "Weiter" um zur nächsten Frage zu kommen.'
                  : 'Questions appear one at a time. Clicking an answer selects it and you immediately see feedback: green = correct, red = wrong. Click "Next" to move on.'}
              </>
            ),
            mockup: <DemoQuizQuestion showAnswered />,
          },
          {
            number: 3,
            title: isDe ? 'Ergebnis-Screen' : 'Result screen',
            description: (
              <>
                {isDe
                  ? 'Am Ende siehst du deine Punktzahl als Kreisdiagramm, den Prozentwert und eine personalisierte Nachricht. Zwei Buttons: "Noch mal" startet neu, "Zurück zum Event" schließt das Quiz.'
                  : 'At the end you see your score as a circular chart, the percentage, and a personalized message. Two buttons: "Retry" starts over, "Back to event" closes the quiz.'}
              </>
            ),
            mockup: <DemoQuizResult score={4} total={5} />,
          },
          {
            number: 4,
            title: isDe ? 'Mehrmals spielen erlaubt' : 'Multiple attempts allowed',
            description: (
              <>
                {isDe
                  ? 'Klicke "Noch mal" um einen weiteren Versuch zu starten. Die Reihenfolge der Fragen bleibt gleich, aber du kannst jetzt falsche Antworten korrigieren. Es wird immer nur der beste Versuch gespeichert — alte schlechtere Ergebnisse werden überschrieben.'
                  : 'Click "Retry" to start another attempt. Question order stays the same but you can correct wrong answers. Only the best attempt is stored — worse older results are overwritten.'}
              </>
            ),
            tip: isDe
              ? 'Du siehst deinen bisher besten Score klein unter dem Quiz-Start-Button — so weißt du was du schlagen musst.'
              : 'You see your best score so far in small text under the quiz start button — so you know what to beat.',
          },
        ],
      },
      {
        perspective: 'organizer',
        title: isDe ? 'Ergebnisse auswerten' : 'Review the results',
        intro: (
          <>
            {isDe
              ? 'Sobald Teilnehmer anfangen zu spielen, erscheinen die Ergebnisse im Admin Center des Events.'
              : 'As soon as attendees start playing, results show up in the event\'s admin center.'}
          </>
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Admin Center → Quiz-Tab öffnen' : 'Admin Center → Quiz tab',
            description: (
              <>
                {isDe
                  ? 'Im Admin Center des Events findest du oben den Tab "Quiz-Ergebnisse". Darunter siehst du eine sortierte Rangliste, Gesamt-Teilnehmer, Durchschnitts-Punktzahl und die Spanne.'
                  : 'In the event admin center you\'ll find a "Quiz results" tab at the top. Below it you see a sorted leaderboard, total participants, average score, and range.'}
              </>
            ),
            mockup: <ClickPath label="Admin Center" label2={isDe ? 'Quiz-Tab' : 'Quiz tab'} />,
          },
          {
            number: 2,
            title: isDe ? 'Rangliste & Statistiken' : 'Leaderboard & stats',
            description: (
              <>
                {isDe
                  ? 'Die Tabelle zeigt alle Teilnehmer sortiert nach Punktzahl (bei Gleichstand: schnellere Zeit gewinnt). Auf Platz 1 steht ein 🏆. In der Spalte "Punkte" werden volle Punktzahlen grün hervorgehoben — das sind deine Gewinner-Kandidaten für eine Verlosung.'
                  : 'The table lists all attendees sorted by score (tie-break: faster time wins). Position 1 gets a 🏆. In the "Score" column, perfect scores are highlighted green — these are your giveaway candidates.'}
              </>
            ),
            mockup: <DemoQuizScoreboard />,
          },
          {
            number: 3,
            title: isDe ? 'Pro-Frage-Auswertung' : 'Per-question analysis',
            description: (
              <>
                {isDe
                  ? 'Wenn du auf eine einzelne Frage klickst, siehst du eine Balken-Statistik der Antworten: welche Option wurde wie oft gewählt, wie viele haben richtig geantwortet. Gut geeignet, um schlecht formulierte Fragen zu identifizieren.'
                  : 'Click a single question to see a bar chart of the answers: how often each option was picked, how many were correct. Useful to identify badly phrased questions.'}
              </>
            ),
            tip: isDe
              ? 'Wenn eine Option >80% falsch beantwortet wurde, ist die Frage oft mehrdeutig — Wortlaut nochmal prüfen.'
              : 'If an option is answered wrong >80% of the time, the question is often ambiguous — check the wording.',
          },
          {
            number: 4,
            title: isDe ? 'Ergebnisse exportieren' : 'Export results',
            description: (
              <>
                {isDe
                  ? 'Mit dem "CSV exportieren"-Button lädst du alle Ergebnisse als Tabelle herunter: Name, E-Mail, beste Punktzahl, Anzahl Versuche, Zeit. Nützlich für Verlosungen, externe Analyse oder Archivierung.'
                  : 'The "Export CSV" button downloads all results as a spreadsheet: name, email, best score, number of attempts, time. Useful for giveaways, external analysis, or archiving.'}
              </>
            ),
          },
          {
            number: 5,
            title: isDe ? 'Gewinner ziehen' : 'Draw a winner',
            description: (
              <>
                {isDe
                  ? 'Über den Button "Zufälligen Gewinner ziehen" wählt die App automatisch einen Teilnehmer mit voller Punktzahl aus. Das Ergebnis ist direkt teilbar oder kopierbar — praktisch für Bühnen-Moderation.'
                  : 'The "Draw random winner" button picks an attendee with a perfect score automatically. The result can be shared or copied directly — handy for stage moderation.'}
              </>
            ),
            mockup: (
              <Callout variant="tip" title={isDe ? 'Fair-Play' : 'Fair play'}>
                {isDe
                  ? 'Die Zufallsauswahl erfolgt clientseitig aus der aktuellen Rangliste — transparent und nachvollziehbar.'
                  : 'The random draw is computed client-side from the current leaderboard — transparent and reproducible.'}
              </Callout>
            ),
          },
        ],
      },
    ],
  };
}
