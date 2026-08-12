import * as React from 'react';
import { ManualSection } from '../types';

/**
 * v29.10: Systemarchitektur im Handbuch.
 *
 * Es gab bisher nur den Flow-Artikel — der beschreibt die sieben Flows im
 * Detail, beantwortet aber nicht die Frage davor: Wo läuft die App überhaupt,
 * welche Listen trägt sie, und warum liegt zwischen Klick und Bestätigungsmail
 * eine Minute? Diese Sektion liefert das Gesamtbild und übergibt für die
 * Details an den Flow-Artikel.
 *
 * Das Schaubild ist bewusst Inline-SVG mit den App-Variablen (--dex-*): Es
 * folgt damit dem Theme, skaliert mit der Spaltenbreite und braucht keine
 * Bilddatei, die beim nächsten Umbau veraltet im Bundle liegen bleibt.
 * Konventionen wie im ausführlichen Dokument (docs/architektur.html):
 * durchgezogen = synchron, gestrichelt = asynchron über eine Warteschlange.
 */

const SystemDiagram = (props: { isDe: boolean }): React.ReactElement => {
  const { isDe } = props;
  const mono = 'Consolas, "Cascadia Mono", monospace';
  const line = 'var(--dex-gray-300, #c8c8c8)';
  const ink = 'var(--dex-gray-800, #222)';
  const ink2 = 'var(--dex-gray-600, #555)';
  const muted = 'var(--dex-gray-500, #777)';
  const green = 'var(--dex-green-dark, #4a7c1f)';
  const amber = 'var(--dex-orange-dark, #b35a00)';
  const amberBg = 'rgba(237,139,0,0.09)';
  const surface = 'var(--dex-white, #fff)';
  const surface2 = 'var(--dex-gray-50, #f7f7f5)';
  return (
    <div style={{ overflowX: 'auto', border: `1px solid ${line}`, borderRadius: 8, background: surface, padding: 14 }}>
      <svg
        viewBox="0 0 900 560"
        role="img"
        aria-label={isDe
          ? 'Systembild: Das SPFx-Webpart im Browser greift per REST auf SharePoint-Listen zu. Warteschlangen-Listen lösen Power-Automate-Flows aus, die Mails und Kalendereinträge erzeugen.'
          : 'System overview: the SPFx web part in the browser calls SharePoint lists via REST. Queue lists trigger Power Automate flows that send mail and create calendar entries.'}
        style={{ display: 'block', minWidth: 640, width: '100%', height: 'auto' }}
      >
        <defs>
          <marker id="dexArchArrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
            <path d="M0,0 L10,5 L0,10 z" fill={ink2} />
          </marker>
          <marker id="dexArchArrowA" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
            <path d="M0,0 L10,5 L0,10 z" fill={amber} />
          </marker>
        </defs>

        {/* Browser */}
        <rect x="16" y="16" width="360" height="76" rx="6" fill={surface2} stroke={line} />
        <text x="32" y="38" fontSize="10" fill={muted} fontFamily={mono} letterSpacing="1.2">
          {isDe ? 'BROWSER DES NUTZERS' : 'USER’S BROWSER'}
        </text>
        <text x="32" y="60" fontSize="14" fontWeight="600" fill={ink}>SPFx-Webpart</text>
        <text x="32" y="80" fontSize="11.5" fill={ink2} fontFamily={mono}>/SitePages/DEX.aspx</text>

        {/* SharePoint */}
        <rect x="16" y="132" width="868" height="228" rx="6" fill="none" stroke={line} />
        <rect x="16" y="132" width="868" height="26" rx="6" fill={surface2} />
        <text x="32" y="150" fontSize="10" fill={muted} fontFamily={mono} letterSpacing="1.2">SHAREPOINT ONLINE</text>

        <line x1="196" y1="92" x2="196" y2="126" stroke={ink2} strokeWidth="1.6" markerEnd="url(#dexArchArrow)" />
        <text x="210" y="112" fontSize="11.5" fill={ink2}>
          {isDe ? 'REST — mit den Rechten des angemeldeten Nutzers' : 'REST — with the signed-in user’s permissions'}
        </text>

        <rect x="36" y="176" width="400" height="112" rx="5" fill={surface} stroke={line} />
        <text x="52" y="196" fontSize="10" fill={green} fontFamily={mono} letterSpacing="1.1">
          {isDe ? 'STAMMDATEN' : 'MASTER DATA'}
        </text>
        <g fontSize="11.5" fill={ink} fontFamily={mono}>
          <text x="52" y="218">DEX_Events</text>
          <text x="52" y="238">DEX_Participants</text>
          <text x="52" y="258">DEX_Roles</text>
          <text x="52" y="278">DEX_EmailTemplates</text>
          <text x="248" y="218">DEX_EventComms</text>
          <text x="248" y="238">DEX_EventStats</text>
          <text x="248" y="258">DEX_ChangeLog</text>
          <text x="248" y="278">…</text>
        </g>

        <rect x="464" y="176" width="400" height="112" rx="5" fill={amberBg} stroke={amber} />
        <text x="480" y="196" fontSize="10" fill={amber} fontFamily={mono} letterSpacing="1.1">
          {isDe ? 'WARTESCHLANGEN' : 'QUEUES'}
        </text>
        <g fontSize="11.5" fill={ink} fontFamily={mono}>
          <text x="480" y="218">DEX_Emails</text>
          <text x="480" y="238">DEX_Outlook</text>
          <text x="480" y="258">DEX_IDReorder</text>
          <text x="480" y="278">DEX_AccessFix</text>
        </g>
        <g fontSize="11" fill={muted}>
          <text x="660" y="218">{isDe ? 'Mail versenden' : 'send mail'}</text>
          <text x="660" y="238">{isDe ? 'Termin ein-/ausladen' : 'invite / uninvite'}</text>
          <text x="660" y="258">{isDe ? 'IDs + Nachrücken' : 'IDs + waitlist'}</text>
          <text x="660" y="278">{isDe ? 'Zeilen-Autor' : 'row author'}</text>
        </g>

        <rect x="36" y="304" width="828" height="42" rx="5" fill={surface} stroke={line} />
        <text x="52" y="322" fontSize="10" fill={green} fontFamily={mono} letterSpacing="1.1">
          {isDe ? 'EVENT-SUBSITE — EINE JE EVENT' : 'EVENT SUBSITE — ONE PER EVENT'}
        </text>
        <g fontSize="11.5" fill={ink} fontFamily={mono}>
          <text x="52" y="340">Teilnehmer</text>
          <text x="248" y="340">DEX_TeilnehmerCounter</text>
        </g>
        <text x="480" y="340" fontSize="11" fill={muted}>
          {isDe ? 'Anmeldungen (nur eigene Elemente) · Platzzähler für alle lesbar' : 'registrations (own items only) · seat counter readable by everyone'}
        </text>

        {/* Power Automate */}
        <rect x="16" y="412" width="500" height="132" rx="6" fill="none" stroke={line} />
        <rect x="16" y="412" width="500" height="26" rx="6" fill={amberBg} />
        <text x="32" y="430" fontSize="10" fill={amber} fontFamily={mono} letterSpacing="1.2">POWER AUTOMATE</text>
        <g fontSize="11.5" fill={ink} fontFamily={mono}>
          <text x="32" y="458">DEX_SEND_MAIL</text>
          <text x="32" y="478">DEX_Outlook_Einladungen</text>
          <text x="32" y="498">DEX_CreateOutlookEvent</text>
          <text x="32" y="518">DEX_IDReorder_TeilnehmerIDs</text>
          <text x="32" y="538">DEX_AccessFix_Autor</text>
          <text x="266" y="458">DEX_OutlookDeclineHandler</text>
          <text x="266" y="478">DEX_OutlookForwardHandler</text>
        </g>

        <rect x="548" y="412" width="336" height="132" rx="6" fill={surface2} stroke={line} />
        <text x="564" y="430" fontSize="10" fill={muted} fontFamily={mono} letterSpacing="1.2">EXCHANGE ONLINE</text>
        <text x="564" y="456" fontSize="13" fontWeight="600" fill={ink}>Shared Mailbox</text>
        <text x="564" y="474" fontSize="11" fill={ink2} fontFamily={mono}>no_reply.events@…</text>
        <text x="564" y="502" fontSize="13" fontWeight="600" fill={ink}>{isDe ? 'Outlook-Kalender' : 'Outlook calendar'}</text>
        <text x="564" y="522" fontSize="11" fill={ink2}>
          {isDe ? 'Absagen laufen als Mail zurück' : 'declines come back as mail'}
        </text>

        {/* Queue -> Flow */}
        <polyline points="664,288 664,384 266,384 266,406" fill="none" stroke={amber} strokeWidth="1.6" strokeDasharray="6 5" markerEnd="url(#dexArchArrowA)" />
        <text x="654" y="378" fontSize="11.5" fill={amber} textAnchor="end">
          {isDe ? 'neuer Eintrag startet den Flow — Abfrage ~1×/Minute' : 'a new item starts the flow — polled about once a minute'}
        </text>

        {/* Flow -> SharePoint */}
        <polyline points="90,406 90,384 60,384 60,362" fill="none" stroke={amber} strokeWidth="1.6" strokeDasharray="6 5" markerEnd="url(#dexArchArrowA)" />
        <text x="104" y="400" fontSize="11.5" fill={amber}>{isDe ? 'schreibt zurück' : 'writes back'}</text>

        {/* Flow -> Exchange */}
        <line x1="524" y1="466" x2="542" y2="466" stroke={ink2} strokeWidth="1.6" markerEnd="url(#dexArchArrow)" />
        <line x1="542" y1="510" x2="524" y2="510" stroke={amber} strokeWidth="1.6" strokeDasharray="6 5" markerEnd="url(#dexArchArrowA)" />
      </svg>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 18, marginTop: 12, paddingTop: 10, borderTop: `1px solid ${line}`, fontSize: '0.78rem', color: ink2 }}>
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
          <svg viewBox="0 0 34 10" style={{ width: 34, height: 10 }} aria-hidden="true"><line x1="1" y1="5" x2="33" y2="5" stroke={ink2} strokeWidth="1.6" /></svg>
          {isDe ? 'synchron — die App wartet auf die Antwort' : 'synchronous — the app waits for the response'}
        </span>
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
          <svg viewBox="0 0 34 10" style={{ width: 34, height: 10 }} aria-hidden="true"><line x1="1" y1="5" x2="33" y2="5" stroke={amber} strokeWidth="1.6" strokeDasharray="6 5" /></svg>
          {isDe ? 'asynchron über eine Warteschlange' : 'asynchronous via a queue'}
        </span>
      </div>
    </div>
  );
};

export function architectureSection(locale: 'de' | 'en'): ManualSection {
  const isDe = locale === 'de';
  return {
    id: 'architecture',
    title: isDe ? 'Systemarchitektur' : 'System architecture',
    category: 'architecture',
    description: isDe
      ? 'Wo die App läuft, welche Listen sie trägt und warum zwischen Klick und Bestätigungsmail eine Minute liegt.'
      : 'Where the app runs, which lists carry it, and why a minute passes between click and confirmation mail.',
    visibleFor: ['Admin'],
    keywords: isDe
      ? 'Architektur Systemarchitektur Aufbau Schaubild Diagramm SPFx Webpart SharePoint Listen Subsite Teilnehmerliste Queue Warteschlange Power Automate Item-Level-Security Zeilen-Sicherheit Counter Platzzähler DEX_Events DEX_Participants Datenmodell Technik Überblick'
      : 'architecture system architecture diagram overview SPFx web part SharePoint lists subsite participant list queue Power Automate item level security counter DEX_Events DEX_Participants data model technical overview',
    perspectives: [
      {
        perspective: 'admin',
        intro: (
          <>
            {isDe
              ? 'DEX hat keinen eigenen Server. Die Anwendung läuft als SPFx-Webpart im Browser, hält ihre Daten in SharePoint-Listen und lässt alles, was der Browser nicht darf — Mails, Kalendereinträge, Rechte —, von Power-Automate-Flows über Warteschlangen erledigen. Die ausführliche Fassung mit allen Abläufen steht in docs/architektur.html; die Flows im Detail stehen im Artikel „Power Automate Flows".'
              : 'DEX has no server of its own. The app runs as an SPFx web part in the browser, keeps its data in SharePoint lists and hands everything the browser is not allowed to do — mail, calendar entries, permissions — to Power Automate flows via queues. The long form with all the workflows is in docs/architektur.html; the flows in detail are in the „Power Automate Flows" article.'}
          </>
        ),
        steps: [
          {
            number: 1,
            title: isDe ? 'Das Gesamtbild' : 'The whole picture',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Drei Schichten, zwei Arten von Verbindung. Der Browser spricht synchron mit SharePoint. Alles Weitere passiert asynchron: Die App legt einen Auftrag in eine Warteschlangen-Liste, ein Flow findet ihn dort.'
                    : 'Three layers, two kinds of connection. The browser talks to SharePoint synchronously. Everything else happens asynchronously: the app drops a job into a queue list and a flow picks it up there.'}
                </p>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Wichtig: Es gibt keine Verbindung vom Browser zu Power Automate und keine zurück. Beide Seiten treffen sich ausschliesslich in den Warteschlangen.'
                    : 'Important: there is no connection from the browser to Power Automate and none back. The two sides only ever meet in the queues.'}
                </p>
              </>
            ),
            mockup: <SystemDiagram isDe={isDe} />,
          },
          {
            number: 2,
            title: isDe ? 'Die App handelt im Namen des Nutzers' : 'The app acts as the user',
            description: (
              <p style={{ margin: 0 }}>
                {isDe
                  ? 'Jeder Lese- und Schreibvorgang läuft mit den Berechtigungen der Person, die gerade davorsitzt. Die App kann sich keine höheren Rechte verschaffen. Wer eine Teilnehmerliste nicht sehen darf, sieht sie auch in der App nicht — und genau deshalb gibt es die Flows: Sie laufen mit einer Dienst-Identität und erledigen, wofür der Nutzer keine Rechte hat.'
                  : 'Every read and write runs with the permissions of whoever is using the app. It cannot grant itself more. Anyone who may not see an attendee list will not see it in the app either — and that is exactly why the flows exist: they run with a service identity and do what the user is not allowed to do.'}
              </p>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Eine Subsite je Event' : 'One subsite per event',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Die Anmeldungen liegen nicht in einer gemeinsamen Liste, sondern je Event in einer eigenen Subsite mit der Liste „Teilnehmer". Das hält die Listen klein, erlaubt Rechte pro Event und macht das Löschen nach der Aufbewahrungsfrist möglich, ohne das Event zu verlieren.'
                    : 'Registrations do not live in one shared list but in a subsite per event with its own „Teilnehmer" list. That keeps lists small, allows per-event permissions and makes deletion after the retention period possible without losing the event.'}
                </p>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Die Liste läuft mit Zeilen-Sicherheit („nur eigene Elemente"). Weil ein Teilnehmer damit die Belegung nicht zählen kann, gibt es daneben DEX_TeilnehmerCounter — eine Liste ohne Personenbezug, die jeder lesen darf.'
                    : 'The list uses item-level security („only their own items"). Since an attendee cannot count occupancy that way, DEX_TeilnehmerCounter sits next to it — a list without personal data that everyone may read.'}
                </p>
              </>
            ),
            tip: isDe
              ? 'Folge daraus: Es gibt keine Abfrage über alle Anmeldungen hinweg. Wer wissen will, wo eine Person überall angemeldet ist, liest das Register DEX_Participants — oder jede Subsite einzeln.'
              : 'Consequence: there is no query across all registrations. To find every event a person is registered for, read the DEX_Participants registry — or each subsite one by one.',
          },
          {
            number: 4,
            title: isDe ? 'Warteschlangen statt Aufrufe' : 'Queues instead of calls',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Jeder Auftrag ist eine Listenzeile mit Statusfeld. Die App legt sie mit Status „Pending" an und ist fertig — sie wartet nicht. Der Flow setzt auf „Processing", erledigt die Arbeit und schreibt „Sent" bzw. „Done" zurück. Bleibt eine Zeile auf „Processing" stehen, ist genau dieser Auftrag gescheitert — sichtbar in der Liste und im Ausführungsverlauf des Flows.'
                    : 'Every job is a list item with a status field. The app creates it with status „Pending" and is done — it does not wait. The flow sets „Processing", does the work and writes back „Sent" or „Done". An item stuck on „Processing" is exactly that one failed job — visible in the list and in the flow run history.'}
                </p>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Für Nutzer heisst das: Die Anmeldung ist sofort gültig, die Mail kommt gleich. Scheitert der Versand, bleibt die Anmeldung trotzdem bestehen — der Platz zählt, nicht die Benachrichtigung.'
                    : 'For users: the registration is valid immediately, the mail follows. If sending fails the registration still stands — the seat counts, not the notification.'}
                </p>
              </>
            ),
            warning: isDe
              ? 'Die Flows reagieren nur auf NEUE Listeneinträge. Das Ändern eines bestehenden Eintrags löst nichts aus — ein fehlender Outlook-Termin lässt sich deshalb nicht durch erneutes Speichern des Events nachziehen.'
              : 'The flows only react to NEW list items. Updating an existing item triggers nothing — a missing Outlook appointment cannot be fixed by simply saving the event again.',
          },
          {
            number: 5,
            title: isDe ? 'Wo die Wahrheit liegt' : 'Where the truth lives',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Dieselbe Information steht an drei Stellen, mit klarer Rangfolge:'
                    : 'The same information exists in three places, with a clear order of precedence:'}
                </p>
                <ol style={{ paddingLeft: 18, margin: '0 0 6px 0', lineHeight: 1.7 }}>
                  <li>
                    <strong>Teilnehmer</strong> {isDe ? '(Subsite) — die Anmeldung selbst. Hat immer recht.' : '(subsite) — the registration itself. Always right.'}
                  </li>
                  <li>
                    <strong>DEX_TeilnehmerCounter</strong> {isDe ? '— die Zahl für die Anzeige. Wird aus der Teilnehmerliste neu berechnet.' : '— the number shown in the UI. Recomputed from the attendee list.'}
                  </li>
                  <li>
                    <strong>DEX_Participants</strong> {isDe ? '— das Register für „Meine Events". Wird über die Wartungsaktion im Admin Center bereinigt.' : '— the registry behind „My events". Cleaned via the maintenance action in the Admin Center.'}
                  </li>
                </ol>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Widersprechen sich zwei Ansichten über „angemeldet ja/nein", ist der erste Verdacht: dieselbe Person steht unter zwei E-Mail-Adressen in den Listen. Die Adresse ist der einzige Schlüssel — und sie ist nicht eindeutig.'
                    : 'If two views disagree about „registered yes/no", the first suspicion is: the same person appears under two email addresses. The address is the only key — and it is not unique.'}
                </p>
              </>
            ),
          },
        ],
      },
    ],
  };
}
