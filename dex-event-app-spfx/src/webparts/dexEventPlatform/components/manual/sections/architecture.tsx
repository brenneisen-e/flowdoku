import * as React from 'react';
import { ManualSection } from '../types';

/**
 * v29.10: Systemarchitektur im Handbuch — seit v29.12 bewusst KURZ.
 *
 * Warum kurz: Es gibt in der App längst eine Architekturseite (ArchitecturePage,
 * v26.28, Kachel im Admin Hub) mit PDF-Export. Dieser Artikel ist v29.10 dazu
 * entstanden, ohne das zu prüfen — zwei Darstellungen derselben Sache, die
 * auseinanderlaufen. Genau davor warnt CLAUDE.md bei zwei Bedienwegen für
 * dieselbe Auswahl.
 *
 * Aufgeteilt statt gedoppelt: Der Artikel liefert das SCHAUBILD (das der
 * Architekturseite fehlt) und die drei Sätze, die man zum Verstehen braucht.
 * Alles Aufzählende — Listen, Flows, Umgebungen, Sicherung, offene
 * Sicherheitspunkte — steht auf der Architekturseite und wächst dort ins PDF
 * mit. Wer hier etwas ergänzen will, prüft zuerst, ob es dorthin gehört.
 *
 * Das Schaubild ist Inline-SVG mit den App-Variablen (--dex-*): folgt dem
 * Theme, skaliert mit der Spalte und veraltet nicht still als Bilddatei im
 * Bundle. Konventionen wie im Langdokument: durchgezogen = synchron,
 * gestrichelt = asynchron über eine Warteschlange.
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
      ? 'Das Systembild in einem Schaubild, plus die drei Sätze, die man zum Verstehen braucht. Einzelheiten und PDF: Admin Center → Architektur.'
      : 'The system overview as one diagram, plus the three sentences you need. Details and PDF: Admin Center → Architecture.',
    visibleFor: ['Admin'],
    keywords: isDe
      ? 'Architektur Systemarchitektur Aufbau Schaubild Diagramm SPFx Webpart SharePoint Listen Subsite Teilnehmerliste Queue Warteschlange Power Automate Item-Level-Security Zeilen-Sicherheit Counter Platzzähler DEX_Events DEX_Participants Datenmodell Technik Überblick Deployment Rollback Umgebungen DEV PROD Service Principal Sicherung Backup Security Requirements Abweichungen Zielmodell'
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
            title: isDe ? 'Die drei Sätze, auf die es ankommt' : 'The three sentences that matter',
            description: (
              <>
                <ul style={{ paddingLeft: 18, margin: 0, lineHeight: 1.7 }}>
                  <li>{isDe
                    ? 'Die App handelt immer im Namen der Person, die sie bedient — sie kann sich keine höheren Rechte verschaffen. Deshalb gibt es die Flows: Sie laufen mit einer eigenen Identität und erledigen, wofür der Nutzer keine Rechte hat.'
                    : 'The app always acts as the person using it — it cannot grant itself more rights. That is why the flows exist: they run under their own identity and do what the user is not allowed to do.'}</li>
                  <li>{isDe
                    ? 'Jedes Event hat eine eigene Subsite mit eigener Teilnehmerliste („nur eigene Elemente"). Weil ein Teilnehmer damit die Belegung nicht zählen kann, gibt es daneben den Platzzähler — eine Liste ohne Personenbezug, die jeder lesen darf.'
                    : 'Every event has its own subsite with its own attendee list (“own items only”). Since an attendee cannot count occupancy that way, the seat counter sits next to it — a list without personal data that everyone may read.'}</li>
                  <li>{isDe
                    ? 'Zwischen Klick und Bestätigungsmail liegt bis zu eine Minute, weil kein Aufruf stattfindet: Die App legt einen Auftrag in eine Warteschlange, ein Flow findet ihn dort. Flows reagieren nur auf NEUE Einträge — ein fehlender Outlook-Termin lässt sich deshalb nicht durch erneutes Speichern nachziehen.'
                    : 'Up to a minute passes between click and confirmation mail because there is no call: the app drops a job into a queue and a flow picks it up. Flows only react to NEW items — a missing Outlook appointment cannot be fixed by saving again.'}</li>
                </ul>
              </>
            ),
          },
          {
            number: 3,
            title: isDe ? 'Wo die Einzelheiten stehen' : 'Where the details live',
            description: (
              <>
                <p style={{ margin: '0 0 6px 0' }}>
                  {isDe
                    ? 'Alles Aufzählende steht im Admin Center unter „Architektur“ — jede Liste mit ihrer Funktion, alle acht Flows mit Auslöser und Wirkung, die Umgebungen mit Ist-Stand und Zielmodell, Sicherung und Wiederherstellung sowie die offengelegten Sicherheitspunkte. Dort gibt es auch den Knopf „Als PDF herunterladen“.'
                    : 'Everything enumerable is in the Admin Center under “Architecture” — every list with its function, all eight flows with trigger and effect, the environments with current and target state, backup and recovery, and the disclosed security items. That page also has the “Download as PDF” button.'}
                </p>
                <p style={{ margin: 0 }}>
                  {isDe
                    ? 'Die Flows im Detail beschreibt der Artikel „Power Automate Flows“. Die ausführliche Fassung mit allen Abläufen liegt in der Projektdokumentation (docs/architektur.html), die Bewertung der 49 Security Requirements unter docs/security-requirements.html.'
                    : 'The flows in detail are covered by the „Power Automate Flows“ article. The long form with all workflows is in the project documentation (docs/architektur.html), the assessment of the 49 security requirements at docs/security-requirements.html.'}
                </p>
              </>
            ),
            tip: isDe
              ? 'Neue Einzelheiten gehören auf die Architekturseite, nicht in diesen Artikel — dort wachsen sie automatisch ins PDF mit.'
              : 'New details belong on the architecture page, not in this article — there they automatically flow into the PDF.',
          },
        ],
      },
    ],
  };
}
