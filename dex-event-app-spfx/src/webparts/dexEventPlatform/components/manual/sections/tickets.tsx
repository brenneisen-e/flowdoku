import * as React from 'react';
import { ManualSection } from '../types';
import { Callout } from '../ManualMockups';

export function ticketsSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'tickets',
    title: isDe ? 'Tickets — Fragen & Antworten' : 'Tickets — questions & answers',
    category: 'general',
    description: isDe
      ? 'Stelle direkt aus der App eine Frage — mit optionalem Screenshot. Organisator:innen und das Support-Team beantworten sie ohne Umweg über die Mail.'
      : 'Ask a question right from the app — with an optional screenshot. Organizers and the support team answer it without detouring through email.',
    visibleFor: ['User', 'Organizer', 'Admin'],
    keywords: 'frage fragen hilfe support ticket tickets question questions help screenshot bildschirmfoto handbuch problem feedback kontakt antwort answer hast du fragen power-user power user inbox postfach',
    perspectives: [
      {
        perspective: 'user',
        title: isDe ? 'Eine Frage stellen' : 'Asking a question',
        intro: isDe
          ? 'Wenn du beim Anmelden, beim Anlegen eines Events oder irgendwo sonst in der App nicht weiterkommst, musst du niemanden suchen oder eine Mail schreiben — du fragst direkt aus der App heraus. Deine Frage landet automatisch bei der richtigen Person, und die Antwort kommt zurück in die App (und per Mail).'
          : 'When you get stuck — registering, creating an event, or anywhere else in the app — you do not have to look someone up or write an email. You ask straight from the app. Your question is routed to the right person automatically, and the answer comes back into the app (and by email).',
        steps: [
          {
            number: 1,
            title: isDe ? 'Auf „Hast du Fragen?“ klicken' : 'Click „Do you have questions?“',
            description: (
              <>
                {isDe
                  ? 'Oben rechts im Kopfbereich findest du auf jeder Seite (außer der Startseite vor dem Login) einen grünen Button „Hast du Fragen?“. Ein Klick öffnet das Fragen-Fenster. Der Button ist für alle eingeloggten Nutzer da — egal ob du Teilnehmer:in, Organisator:in oder Admin bist.'
                  : 'In the top-right header you will find a green „Do you have questions?“ button on every page (except the landing page before login). One click opens the questions window. The button is available to every signed-in user — whether you are an attendee, an organizer or an admin.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Frage(n) eingeben' : 'Type your question(s)',
            description: (
              <>
                {isDe
                  ? 'Schreibe deine Frage in eigenen Worten in das Textfeld. Du kannst über das „+“ mehrere Fragen auf einmal stellen, falls dir gleich mehrere Dinge unklar sind — jede landet als eigener Punkt. Beschreibe ruhig konkret, was du erreichen wolltest und wo es hakt.'
                  : 'Write your question in your own words into the text field. Use the „+“ to add several questions at once if more than one thing is unclear — each becomes its own item. Feel free to be specific about what you were trying to do and where it stalls.'}
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Erst die App selbst um Rat fragen' : 'Let the app help you first',
            description: (
              <>
                {isDe
                  ? 'Schon während du tippst, schlägt dir die App passende Handbuch-Artikel vor. Oft steht die Antwort dort schon — ein Klick öffnet den Artikel, und du sparst dir das Warten. Hilft dir der Vorschlag nicht, schickst du die Frage einfach ab.'
                  : 'As you type, the app already suggests matching handbook articles. Often the answer is already there — one click opens the article and you skip the wait. If the suggestion does not help, you simply send the question.'}
              </>
            ),
            mockup: (
              <Callout variant="tip">
                {isDe
                  ? 'Die Vorschläge sind echte Handbuch-Artikel — Selbsthilfe in Sekunden, bevor jemand antworten muss.'
                  : 'The suggestions are real handbook articles — self-service in seconds, before anyone has to answer.'}
              </Callout>
            ),
          },
          {
            number: 4,
            title: isDe ? 'Screenshot anhängen (optional)' : 'Attach a screenshot (optional)',
            description: (
              <>
                {isDe
                  ? 'Mit einem Klick fügst du einen Screenshot deines aktuellen Bildschirms hinzu. Das hilft enorm, wenn du eine Fehlermeldung oder eine bestimmte Stelle zeigen willst — die beantwortende Person sieht dann genau, was du siehst, statt es sich vorstellen zu müssen.'
                  : 'With one click you attach a screenshot of your current screen. This helps a lot when you want to show an error message or a particular spot — the person answering then sees exactly what you see instead of having to imagine it.'}
              </>
            ),
          },
          {
            number: 5,
            title: isDe ? 'Absenden — und wer deine Frage bekommt' : 'Send — and who receives your question',
            description: (
              <>
                {isDe
                  ? 'Nach dem Absenden geht deine Frage automatisch an die richtige Stelle: Stellst du als Teilnehmer:in eine Frage zu einem konkreten Event, landet sie direkt bei den Organisator:innen dieses Events (Admins werden zur Info in Kopie gesetzt). Allgemeine Fragen sowie Fragen von Organisator:innen und Admins gehen an das Support-Team („Power-User“). Du musst dir also keine Gedanken machen, an wen du dich wenden sollst.'
                  : 'After sending, your question is routed automatically: if you — as an attendee — ask about a specific event, it goes straight to the organizers of that event (admins are copied for information). General questions and questions from organizers or admins go to the support team (the „power users“). So you never have to wonder whom to address.'}
              </>
            ),
          },
          {
            number: 6,
            title: isDe ? 'Antwort im Tab „Deine Fragen“ verfolgen' : 'Follow up in the „Your questions“ tab',
            description: (
              <>
                {isDe
                  ? 'Im Fragen-Fenster gibt es den Tab „Deine Fragen“. Dort siehst du alle deine bisherigen Tickets mit ihrem Status — von „offen“ über „in Bearbeitung“ bis „beantwortet“ — und liest die Antwort direkt in der App. Zusätzlich bekommst du eine Mail, sobald jemand geantwortet hat.'
                  : 'The questions window has a „Your questions“ tab. There you see all your past tickets with their status — from „open“ through „in progress“ to „answered“ — and read the reply right inside the app. You also receive an email as soon as someone has replied.'}
              </>
            ),
          },
          {
            number: 7,
            title: isDe ? 'Nicht auf die Mail antworten' : 'Do not reply to the email',
            description: (
              <>
                {isDe
                  ? 'Die Benachrichtigungs-Mails dienen nur als Hinweis. Antworten und Rückfragen bitte immer in der App über das Fragen-Fenster — eine Antwort auf die Mail selbst erreicht niemanden. Das steht auch in jeder Mail noch einmal als Hinweis.'
                  : 'The notification emails are just a heads-up. Always reply and ask follow-ups inside the app via the questions window — replying to the email itself reaches no one. Every email repeats this hint as well.'}
              </>
            ),
            warning: isDe
              ? 'Antworte NICHT per Mail — die Mail ist nur eine Benachrichtigung. Nutze immer das Fragen-Fenster in der App.'
              : 'Do NOT reply by email — the email is only a notification. Always use the questions window in the app.',
          },
        ],
      },
      {
        perspective: 'organizer',
        title: isDe ? 'Fragen beantworten' : 'Answering questions',
        intro: isDe
          ? 'Als Organisator:in oder als Mitglied des Support-Teams beantwortest du eingehende Fragen ohne Umweg über das Postfach — direkt in der App. So bleibt der ganze Verlauf an einem Ort, und mehrere Beantwortende kommen sich nicht in die Quere.'
          : 'As an organizer or as a member of the support team you answer incoming questions without detouring through your mailbox — directly in the app. That keeps the whole thread in one place and stops multiple responders from stepping on each other.',
        steps: [
          {
            number: 1,
            title: isDe ? 'Wo die Fragen auftauchen' : 'Where the questions show up',
            description: (
              <>
                {isDe
                  ? 'Wo du Fragen findest, hängt von deiner Rolle ab: Das Support-Team („Power-User“) und Admins sehen auf der Startseite eine neue Kachel „Tickets“ — das ist das gemeinsame Postfach für alle allgemeinen Fragen. Organisator:innen finden die Fragen von Teilnehmer:innen zu ihrem Event direkt im Organizer Center, in einer einklappbaren Box „Offene Fragen (User)“ mit einem Zähler-Badge.'
                  : 'Where you find questions depends on your role: the support team (the „power users“) and admins see a new „Tickets“ tile on the start page — the shared inbox for all general questions. Organizers find the attendee questions about their event right in the Organizer Center, in a collapsible box „Open questions (users)“ with a counter badge.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Ticket öffnen = übernehmen' : 'Opening a ticket = taking it',
            description: (
              <>
                {isDe
                  ? 'Sobald du ein Ticket öffnest, gilt es als „in Bearbeitung“ und ist dir zugeordnet. Andere Beantwortende sehen diesen Status nahezu in Echtzeit — so vermeidet ihr, dass zwei Personen dieselbe Frage gleichzeitig bearbeiten und doppelt antworten.'
                  : 'As soon as you open a ticket it becomes „in progress“ and is assigned to you. Other responders see this status almost in real time — so you avoid two people working the same question at once and answering twice.'}
              </>
            ),
            mockup: (
              <Callout variant="info">
                {isDe
                  ? 'Der „in Bearbeitung“-Status verhindert Doppel-Antworten, ohne dass ihr euch abstimmen müsst.'
                  : 'The „in progress“ status prevents duplicate answers without any need to coordinate.'}
              </Callout>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Wieder freigeben, wenn jemand anderes besser passt' : 'Release it when someone else fits better',
            description: (
              <>
                {isDe
                  ? 'Merkst du, dass jemand anderes die Frage besser beantworten kann, gibst du das Ticket per „Wieder freigeben“ zurück. Es ist dann wieder offen, und eine andere Person kann es übernehmen. Kein Ticket bleibt also bei dir „kleben“, nur weil du es einmal geöffnet hast.'
                  : 'If you notice someone else can answer the question better, hand the ticket back via „Release“. It becomes open again and another person can take it over. So no ticket gets „stuck“ with you just because you once opened it.'}
              </>
            ),
          },
          {
            number: 4,
            title: isDe ? 'Antwort schreiben — mit nützlichen Beigaben' : 'Write the answer — with helpful extras',
            description: (
              <>
                {isDe
                  ? 'Deine Antwort besteht zuerst aus Freitext. Zusätzlich kannst du sie anreichern: passende Handbuch-Artikel verlinken (damit die Person das Thema vollständig nachlesen kann), auf einen bestimmten Schritt im Event-Wizard verweisen (als Orientierung mit Visual) und ein Bild anhängen. So wird aus der Antwort eine echte Anleitung statt nur eines Satzes.'
                  : 'Your answer starts with free text. On top of that you can enrich it: link matching handbook articles (so the person can read up on the whole topic), point to a particular step in the event wizard (as guidance with a visual), and attach an image. That turns the answer into a real walkthrough rather than a single sentence.'}
              </>
            ),
            tip: isDe
              ? 'Ein verlinkter Handbuch-Artikel hilft oft mehr als ein langer Text — er beantwortet auch die nächste, noch ungestellte Frage gleich mit.'
              : 'A linked handbook article often helps more than a long text — it also answers the next, as-yet-unasked question right away.',
          },
          {
            number: 5,
            title: isDe ? 'Senden — Ticket ist beantwortet' : 'Send — the ticket is answered',
            description: (
              <>
                {isDe
                  ? 'Nach dem Senden gilt das Ticket als „beantwortet“ und wird geschlossen. Die fragende Person bekommt eine Mail, dass eine Antwort vorliegt, und liest sie im Tab „Deine Fragen“ in der App nach.'
                  : 'After sending, the ticket counts as „answered“ and is closed. The person who asked gets an email that a reply is available and reads it in the „Your questions“ tab in the app.'}
              </>
            ),
          },
          {
            number: 6,
            title: isDe ? 'Benachrichtigungen — und der Verweis in die App' : 'Notifications — and the pointer back into the app',
            description: (
              <>
                {isDe
                  ? 'Bei einer neuen Frage bekommen die Beantwortenden eine Mail mit einem Direkt-Link, der das passende Ticket in der App öffnet. Bei deiner Antwort geht eine Mail an die fragende Person. In jeder dieser Mails steht ausdrücklich: nicht per Mail antworten, sondern in der App — so bleibt der gesamte Austausch an einem Ort und geht nicht in Postfächern verloren.'
                  : 'On a new question the responders receive an email with a direct link that opens the matching ticket in the app. On your reply an email goes to the person who asked. Every one of these emails explicitly states: do not reply by email, use the app instead — that keeps the whole exchange in one place and out of scattered mailboxes.'}
              </>
            ),
            warning: isDe
              ? 'Auch als Beantwortende:r gilt: alles in der App erledigen. Die Mail ist nur der Wegweiser zum Ticket.'
              : 'As a responder too: do everything in the app. The email is just the signpost to the ticket.',
          },
        ],
      },
      {
        perspective: 'admin',
        title: isDe ? 'Als Admin im Überblick' : 'Admin overview',
        intro: isDe
          ? 'Als Admin behältst du das ganze Ticket-Geschehen im Blick und springst überall ein, wo es nötig ist.'
          : 'As an admin you keep an eye on the whole ticket activity and step in wherever needed.',
        steps: [
          {
            number: 1,
            title: isDe ? 'Die Kachel „Tickets“ auf der Startseite' : 'The „Tickets“ tile on the start page',
            description: (
              <>
                {isDe
                  ? 'Wie das Support-Team siehst du auf der Startseite die Kachel „Tickets“ — das gemeinsame Postfach für alle allgemeinen Fragen. Von dort öffnest, übernimmst und beantwortest du Tickets genau wie ein:e Beantwortende:r.'
                  : 'Like the support team you see the „Tickets“ tile on the start page — the shared inbox for all general questions. From there you open, take over and answer tickets exactly like any responder.'}
              </>
            ),
          },
          {
            number: 2,
            title: isDe ? 'Mitlesen bei Event-Fragen' : 'Visibility on event questions',
            description: (
              <>
                {isDe
                  ? 'Wenn ein:e Teilnehmer:in eine Frage zu einem konkreten Event stellt, geht diese an die Organisator:innen des Events — und du wirst als Admin zur Information in Kopie gesetzt. So bekommst du mit, was bei den Events gefragt wird, ohne jedes Ticket selbst beantworten zu müssen, und kannst bei Bedarf eingreifen.'
                  : 'When an attendee asks about a specific event, the question goes to that event’s organizers — and you, as admin, are copied for information. That way you stay aware of what is being asked across events without having to answer every ticket yourself, and you can step in when needed.'}
              </>
            ),
          },
        ],
      },
    ],
  };
}
