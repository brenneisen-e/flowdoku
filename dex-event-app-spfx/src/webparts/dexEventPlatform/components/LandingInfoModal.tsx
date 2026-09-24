/**
 * Info-Modal auf der Landing-Page („Über die App").
 *
 * Erklärt die DEX Event Experience Platform für neue User, die noch nicht
 * wissen, was die App kann. Abschnitte: Einsatzbereich - Ablauf - Funktionen
 * (Aufklapper) - Status - Kontakt.
 *
 * v31.2: Auf das gemeinsame `Modal` (Kopf, Fuß, Escape, Backdrop) und die
 * `dex-ui-*`-Klassen umgestellt. Vorher baute die Datei ihr eigenes Overlay
 * mit Farbverlauf-Hero, klebendem X und drei Inline-Style-Helfern — der
 * einzige Dialog der App, der anders aussah als alle anderen.
 *
 * Zweisprachig DE + EN.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Mail, Info, AlertCircle, ChevronDown } from './Icons';
import Modal from './Modal';
import { cx } from './dexUi';
import { APP_VERSION } from '../version';
import { DEX_TEAM_EMAIL } from '../utils/supportContact';
import { useNavigation } from '../context/NavigationContext';

interface Props {
  open: boolean;
  locale: 'de' | 'en';
  onClose: () => void;
  /** v30.25: Startet die geführte Tour (Header reicht openTutorial durch). */
  onStartTutorial?: () => void;
}

interface Feature {
  icon: string;
  title: string;
  body: string;
}

/** v31.2: Ein Ablauf-Schritt — kurzer Titel plus eine Zeile, was dabei passiert. v31.95: mit Symbol. */
interface Step { icon: string; title: string; hint: string; }

/** v31.2: Ein Einsatzbeispiel — Kategorie plus konkrete Beispiele. v31.95: mit Symbol für die Kachel. */
interface UseCase { icon: string; title: string; sub: string; }

const EVENT_MGMT_URL = 'https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx';

export default function LandingInfoModal({ open, locale, onClose, onStartTutorial }: Props): React.ReactElement | null {
  // v31.2: Die 15 Funktions-Kacheln sind der längste Block des Dialogs. Zum
  // Einstieg reichen Einsatzbereich und Ablauf; die Kacheln bleiben zu, bis
  // jemand sie sehen will — sonst zeigt der Dialog beim Öffnen drei Bildschirme.
  const [showFeatures, setShowFeatures] = React.useState(false);
  // v31.94: „Handbuch" im Self-Service-Kasten — Hook VOR dem frühen Return.
  const { navigate } = useNavigation();
  if (!open) return null;

  const isDE = locale === 'de';
  // v31.95: „Hast du Fragen?" direkt aus dem Self-Service-Kasten — derselbe
  // Dialog wie der Knopf im Kopf (QuestionButton hört auf das Fenster-Ereignis,
  // `tab: 'ask'` öffnet gleich „Frage stellen"). Erst schließen, dann öffnen,
  // sonst liegen zwei Dialoge übereinander.
  const openQuestions = (): void => {
    onClose();
    window.setTimeout(() => { try { window.dispatchEvent(new CustomEvent('dex-open-questions', { detail: { tab: 'ask' } })); } catch { /* */ } }, 150);
  };

  // v31.2: Der Punkt „Wofür DEX gedacht ist" (v28.40) steht jetzt im Hinweis-
  // kasten unter den Einsatzbeispielen — zwei Blöcke zum selben Thema waren
  // getrennt, die Aussage ist vollständig dort.
  const featuresDE: Feature[] = [
    { icon: 'Calendar', title: 'Event in Minuten anlegen', body: 'Geführter Wizard durch alle Schritte: Titel, strukturierte Adresse, Zeit, Teilnehmerlimit, Deadline, Bild, Organizer. Fertig.' },
    { icon: 'People', title: 'Zielgruppen-Steuerung', body: 'Sichtbar für bestimmte Standorte (z.B. Köln+Düsseldorf), Entra-Verteilerlisten (z.B. SAPAlliance@) oder einzeln eingeladene Personen. UND/ODER-verknüpft.' },
    { icon: 'CheckList', title: 'Custom Fields pro Event', body: 'Allergien, T-Shirt-Größe, Hotelwunsch, Transfer-Bedarf… beliebige Zusatzfragen als Text, Dropdown, Zahl oder Checkbox — einzeln als Pflichtfeld markierbar.' },
    { icon: 'CheckboxComposite', title: 'Warteliste mit Nachrück-Automatik', body: 'Volles Event? User landen automatisch auf der Warteliste mit Positionsanzeige. Wird ein Platz frei, rückt der Nächste automatisch nach — inkl. Mail + Outlook.' },
    { icon: 'Mail', title: 'Automatische Emails im Deloitte-Design', body: 'Anmelde-, Abmelde-, Warteliste-, Nachrück-Mails automatisch verschickt. Templates pro Event anpassbar (Logo, Farben, Texte DE/EN).' },
    { icon: 'OutlookLogoInverse', title: 'Outlook-Kalender-Integration', body: 'Jede Anmeldung erzeugt eine Outlook-Einladung aus no_reply.events@deloitte.de. Änderungen am Event aktualisieren alle Termine automatisch.' },
    { icon: 'AccountManagement', title: 'Organizer Center', body: 'Teilnehmer suchen, sortieren, ein-/auschecken, manuell an-/abmelden. Massen-Mails mit RichText-Editor im Deloitte-Design an alle Bestätigten.' },
    { icon: 'QRCode', title: 'QR-Code Check-In', body: 'Per Klick generiert die App personalisierte QR-Codes und verschickt sie an alle Teilnehmer. Am Event scannen → Check-In + Live-Statistik.' },
    { icon: 'DietPlanNotebook', title: 'Quiz / Fun-Zone', body: 'Optional eigene Quizfragen pro Event. Teilnehmer beantworten sie in der App. Statistik pro Frage + Top-10-Liste.' },
    { icon: 'Attach', title: 'Dokumente & PDFs', body: 'Beliebig viele Dokumente pro Event (Agenda, Anfahrt, Catering-Menü, Sponsoren). Teilnehmer sehen & laden sie; PDF-Vorschau direkt im Browser.' },
    { icon: 'Running', title: 'B2Run-Spezialfunktionen', body: 'Getrennte Kapazitäten für Durchstarter + Funstarter, Startblock-Verwaltung, T-Shirt-Größen, B2Run-CSV-Export, eigene Mail-Templates.' },
    { icon: 'Permissions', title: 'Rollen & Berechtigungen', body: 'User: eigene An-/Abmeldung. Organizer: eigene Events verwalten. Admin: Vollzugriff inkl. globale Templates und Rollen-Matrix.' },
    { icon: 'Globe', title: 'Deutsch + Englisch', body: 'Kompletter Language-Switch oben links. Mails werden in der Sprache verschickt, die der Organizer pro Event festgelegt hat.' },
    { icon: 'CellPhone', title: 'Mobile-optimiert', body: 'Responsive Layout, große Touch-Targets, optimierte PDF-Vorschau für iOS/Android. Check-in geht mit Handy + QR-Scan.' },
    { icon: 'Shield', title: 'Sicher & DSGVO-konform', body: 'Läuft komplett im Deloitte-SharePoint-Tenant. Keine externen APIs. Item-Level-Security: jeder sieht nur seine eigenen Registrierungen.' },
  ];

  const featuresEN: Feature[] = [
    { icon: 'Calendar', title: 'Create an event in minutes', body: 'Guided wizard through every step: title, structured address, time, capacity, deadline, image, organizer. Done.' },
    { icon: 'People', title: 'Audience targeting', body: 'Visible for selected offices (e.g. Cologne+Düsseldorf), Entra distribution lists (e.g. SAPAlliance@) or individually invited people. AND/OR combined.' },
    { icon: 'CheckList', title: 'Event-specific custom fields', body: 'Allergies, t-shirt size, hotel request, transfer need… any custom question as text, dropdown, number or checkbox — individually required or optional.' },
    { icon: 'CheckboxComposite', title: 'Waitlist with auto-promotion', body: 'Event full? Users are added to the waitlist with their position. When someone cancels, the next is promoted automatically — incl. email + Outlook invite.' },
    { icon: 'Mail', title: 'Automated emails in Deloitte design', body: 'Registration, cancellation, waitlist and promotion emails sent automatically. Per-event template overrides (logo, colors, text in EN/DE).' },
    { icon: 'OutlookLogoInverse', title: 'Outlook calendar integration', body: 'Every registration creates an Outlook invite from no_reply.events@deloitte.de. Event changes update all invites automatically.' },
    { icon: 'AccountManagement', title: 'Organizer Center', body: 'Search, sort, check-in/out participants, register/cancel manually. Mass emails with RichText editor in Deloitte design to all confirmed guests.' },
    { icon: 'QRCode', title: 'QR-Code check-in', body: 'One click generates personalized QR codes for all participants. Scan at the event to check them in — live statistics show who\'s already there.' },
    { icon: 'DietPlanNotebook', title: 'Quiz / Fun-Zone', body: 'Optional per-event quiz questions. Participants answer in the app. Per-question stats + Top-10 leaderboard.' },
    { icon: 'Attach', title: 'Documents & PDFs', body: 'Unlimited documents per event (agenda, directions, catering menu, sponsors). Participants view & download; PDF preview directly in-browser.' },
    { icon: 'Running', title: 'B2Run specific features', body: 'Separate capacities for "Durchstarter" and "Funstarter" runners, start-block management, t-shirt sizes, B2Run CSV export, dedicated mail templates.' },
    { icon: 'Permissions', title: 'Roles & permissions', body: 'User: own registration. Organizer: manage own events. Admin: full access incl. global templates and role matrix.' },
    { icon: 'Globe', title: 'German + English', body: 'Full language switch top left. Emails are sent in the language the organizer set on the event.' },
    { icon: 'CellPhone', title: 'Mobile-optimized', body: 'Responsive layout, large touch targets, optimized PDF preview for iOS/Android. Check-in works on-the-go with phone + QR scan.' },
    { icon: 'Shield', title: 'Secure & GDPR-compliant', body: 'Runs entirely inside the Deloitte SharePoint tenant. No external APIs. Item-level security: users only see their own registrations.' },
  ];

  const useCases: UseCase[] = isDE ? [
    { icon: 'Presentation', title: 'Leadership- & Strategie-Meetings', sub: 'z. B. SR&T P/MD/D Meeting mit 450 Teilnehmenden' },
    { icon: 'Emoji2', title: 'Firmen-Events', sub: 'Sommerfeste, Weihnachtsfeiern, Bereichs-Offsites' },
    { icon: 'Hotel', title: 'Assistenz- & Team-Meetings', sub: 'mit Transfer- und Hotelbuchung' },
    { icon: 'Running', title: 'Lauf-Events', sub: 'B2Run, JPMorgan Corporate Challenge — mit Startblöcken und geteilten Kapazitäten' },
    { icon: 'CalendarAgenda', title: 'Alles dazwischen', sub: 'vom Lunch mit 10 Personen bis zur Großveranstaltung mit über 500' },
  ] : [
    { icon: 'Presentation', title: 'Leadership & strategy meetings', sub: 'e.g. SR&T P/MD/D meeting with 450 participants' },
    { icon: 'Emoji2', title: 'Company events', sub: 'summer parties, Christmas celebrations, team offsites' },
    { icon: 'Hotel', title: 'Assistant & team meetings', sub: 'with transfer and hotel booking' },
    { icon: 'Running', title: 'Running events', sub: 'B2Run, JPMorgan Corporate Challenge — with start blocks and split capacities' },
    { icon: 'CalendarAgenda', title: 'Everything in between', sub: 'from a lunch with 10 people to a flagship event with 500+' },
  ];

  // v31.2: Ablauf als nummerierte Schritt-Zeilen (Titel + Folge) statt einer
  // Aufzählung mit Fettdruck mitten im Satz — man sieht auf einen Blick, wer
  // was tut und was die App daraufhin automatisch erledigt.
  // v31.94: Auf den heutigen Stand gebracht (Nutzer 24.09.2026: „optimiere
  // diese Formulierungen gemäß der aktuellen Logik"): Du-Form, neun Wizard-
  // Schritte, Termine und Programmpunkte, Organizer Center statt Admin Center,
  // Check-in-Team, Teilnehmer-ID und Walk-in. „Wir helfen beim ersten Mal" ist
  // raus — DEX ist Self-Service, siehe Kasten darunter.
  const steps: Step[] = isDE ? [
    { icon: 'PageEdit', title: 'Du legst dein Event an', hint: 'Der Wizard führt in neun Schritten durch Termin und Ort, Kapazität und Fristen, Sichtbarkeit, Anmeldeformular sowie E-Mail- und Outlook-Kommunikation. Mehrere Termine oder Programmpunkte lassen sich direkt mit anlegen.' },
    { icon: 'Settings', title: 'Die App richtet die Infrastruktur ein', hint: 'Je Event und Termin eine eigene Teilnehmerliste, Berechtigungen für dein Team sowie E-Mail-Vorlagen im Deloitte-Design.' },
    { icon: 'Group', title: 'Nur die eingeladene Zielgruppe sieht das Event', hint: 'Gesteuert über Standort, Zielgruppe, Verteilerliste oder persönliche Einladung — für diese Personen erscheint das Event unter „Aktuelle Events".' },
    { icon: 'Touch', title: 'Anmeldung mit einem Klick — auch stellvertretend', hint: 'Bestätigung per E-Mail und Outlook-Termin erfolgen automatisch. Assistenzen melden stellvertretend an; eine Abmeldung ist jederzeit unter „Meine Events" möglich.' },
    { icon: 'Clock', title: 'Bei voller Kapazität greift die Warteliste', hint: 'Wird ein Platz frei, rückt die nächste Person automatisch nach — inklusive E-Mail und Outlook-Termin.' },
    { icon: 'AccountManagement', title: 'Verwaltung im Organizer Center', hint: 'Teilnehmerliste, Massen-Mails, Dokumente, Hotel- und Zimmerplanung sowie Auswertungen — mit jederzeitigem Überblick über den Anmeldestand.' },
    { icon: 'QRCode', title: 'Am Event-Tag: Check-in per QR-Code', hint: 'Die QR-Codes werden per E-Mail versendet. Das Check-in-Team scannt sie oder erfasst die Teilnehmer-ID; Walk-ins werden direkt vor Ort registriert. Der Live-Zähler zeigt den aktuellen Stand.' },
  ] : [
    { icon: 'PageEdit', title: 'You create your event', hint: 'The wizard guides you through nine steps: date and venue, capacity and deadlines, visibility, registration form, and email and Outlook communication. Multiple sessions or agenda items can be added right away.' },
    { icon: 'Settings', title: 'The app sets up the infrastructure', hint: 'A dedicated participant list per event and session, permissions for your team, and email templates in Deloitte design.' },
    { icon: 'Group', title: 'Only the invited audience sees the event', hint: 'Controlled by office, audience, distribution list or personal invitation — for these people the event appears under "Current events".' },
    { icon: 'Touch', title: 'Registration with one click — also on behalf of others', hint: 'Confirmation by email and Outlook invite are sent automatically. Assistants can register on behalf of others; cancelling is possible at any time under "My events".' },
    { icon: 'Clock', title: 'At full capacity, the waitlist takes over', hint: 'When a seat becomes available, the next person is promoted automatically — including email and Outlook invite.' },
    { icon: 'AccountManagement', title: 'Management in the Organizer Center', hint: 'Participant list, mass emails, documents, hotel and room planning, and reports — with an overview of the registration status at any time.' },
    { icon: 'QRCode', title: 'On event day: check-in by QR code', hint: 'QR codes are sent by email. The check-in team scans them or enters the participant ID; walk-ins are registered on site. The live counter shows the current status.' },
  ];

  const features = isDE ? featuresDE : featuresEN;

  return (
    <Modal
      open
      onClose={onClose}
      maxWidth={860}
      ariaLabel={isDE ? 'Über die App' : 'About the app'}
      icon={<Info size={20} />}
      title={<>DEX Event Experience Platform <span className="dex-ui-pill dex-ui-pill--gray" style={{ verticalAlign: 'middle', marginLeft: 6 }}>v{APP_VERSION}</span></>}
      // v28.45: „alle Deloitte-Events" war zu weit gefasst — DEX ist auf
      // interne Events ausgelegt, nicht auf externe mit externen Gästen.
      subtitle={isDE
        ? 'Deine zentrale Plattform für interne Deloitte Events: von der Ausschreibung bis zum Check-in in einer App. Registrierung, Outlook-Einladungen, Warteliste, Massen-Mails, QR-Codes, Dokumente — alles im Deloitte-SharePoint-Tenant.'
        : 'Your central platform for internal Deloitte events: from announcement to check-in in a single app. Registration, Outlook invites, waitlist, mass emails, QR codes, documents — all inside the Deloitte SharePoint tenant.'}
      footer={<>
        {/* v30.25: Tutorial-Einstieg für JEDE Rolle. Die „Neu hier?"-Pille auf
            der Startseite wird Organizern und Admins nicht mehr angeboten (sie
            kennen die App) — über „Über die App" bleibt die geführte Tour aber
            jederzeit erreichbar, z.B. um sie neuen Kolleg:innen zu zeigen. */}
        {onStartTutorial && (
          <span className="dex-ui-modal-foot-left">
            <button type="button" className="btn btn-outline" onClick={() => { onClose(); window.setTimeout(() => onStartTutorial(), 250); }}>
              {isDE ? 'Geführtes Tutorial starten' : 'Start the guided tutorial'}
            </button>
          </span>
        )}
        <button type="button" className="btn btn-secondary" onClick={onClose}>{isDE ? 'Schließen' : 'Close'}</button>
        <a className="btn btn-primary" href={`mailto:${DEX_TEAM_EMAIL}?subject=DEX Event Experience Platform – Interesse`}>
          <Mail size={16} /> {isDE ? 'Kontakt aufnehmen' : 'Get in touch'}
        </a>
      </>}
    >
      <div>
        {/* Einsatzbereich */}
        <section className="dex-ui-section">
          <h4 className="dex-ui-section-title">{isDE ? 'Für diese Events ist DEX gemacht' : 'DEX is built for these events'}</h4>
          {/* v31.95: Kacheln mit Symbol statt Haken-Zeilen — die Liste sah
              aus wie die Schritte darunter (Nutzer 24.09.2026: „hier sieht
              irgendwie alles gleich aus"). */}
          <div className="dex-ui-grid-auto" style={{ gap: 10 }}>
            {useCases.map((u, i) => (
              <div key={i} className="dex-ui-card dex-ui-card--soft" style={{ padding: 12, display: 'flex', gap: 10, alignItems: 'flex-start' }}>
                <span className="dex-ui-choice-icon" aria-hidden="true"><Icon iconName={u.icon} style={{ fontSize: 16 }} /></span>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontSize: '0.86rem', color: 'var(--dex-gray-800)', lineHeight: 1.3 }}>{u.title}</div>
                  <div className="dex-ui-muted" style={{ fontSize: '0.76rem', lineHeight: 1.45, marginTop: 2 }}>{u.sub}</div>
                </div>
              </div>
            ))}
          </div>
          <div className="dex-ui-callout dex-ui-callout--neutral" style={{ marginTop: 10 }}>
            <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
            <span>
              {isDE
                ? <>Gedacht für <strong>interne Deloitte Events</strong> und die Koordination der Deloitte-Teilnahme an externen Veranstaltungen. <strong>Nicht</strong> für externe Events mit externen Teilnehmern — alles dazu findest du im <a href={EVENT_MGMT_URL} target="_blank" rel="noopener noreferrer">Event Management im DeloitteNet</a>.</>
                : <>Built for <strong>Deloitte-internal events</strong> and for coordinating Deloitte participation in external events. <strong>Not</strong> for external events with external attendees — everything about those is on <a href={EVENT_MGMT_URL} target="_blank" rel="noopener noreferrer">Event Management on DeloitteNet</a>.</>}
            </span>
          </div>
        </section>

        {/* So funktioniert es */}
        <section className="dex-ui-section">
          <h4 className="dex-ui-section-title">{isDE ? 'So läuft ein Event mit DEX' : 'How an event runs with DEX'}</h4>
          {/* v31.95: Prozess-Zeitstrahl statt sieben gleicher Karten — Nummern
              auf einer Linie, je Schritt ein Symbol. */}
          <div className="dex-ui-process">
            {steps.map((s, i) => (
              <div key={i} className="dex-ui-process-item">
                <span className="dex-ui-process-num">{i + 1}</span>
                <span className="dex-ui-process-icon" aria-hidden="true"><Icon iconName={s.icon} style={{ fontSize: 16 }} /></span>
                <div className="dex-ui-process-body">
                  <div className="dex-ui-process-title">{s.title}</div>
                  <div className="dex-ui-process-hint">{s.hint}</div>
                </div>
              </div>
            ))}
          </div>
        </section>

        {/* v31.94: Self-Service — Nutzer-Ansage 24.09.2026: „bei Infos der App
            soll auch stehen, dass DEX als Self-Service gedacht ist … Fragen
            über die Fragen-Funktion oben oder in einem der monatlichen Calls
            oder gerne selber im Handbuch". */}
        <section className="dex-ui-section">
          <h4 className="dex-ui-section-title">{isDE ? 'DEX ist Self-Service' : 'DEX is self-service'}</h4>
          <div className="dex-ui-callout dex-ui-callout--neutral">
            <span className="dex-ui-callout-icon"><Info size={16} /></span>
            <span>
              {isDE
                ? <>DEX ist als <strong>Self-Service-Plattform</strong> konzipiert: Du legst dein Event eigenständig an und verwaltest es über den gesamten Ablauf — der Wizard führt dich dabei Schritt für Schritt. Unterstützung findest du im <button type="button" className="dex-ui-textlink" onClick={() => { onClose(); navigate('manual'); }}>Handbuch</button>, über <button type="button" className="dex-ui-textlink" onClick={openQuestions}>&bdquo;Hast du Fragen?&ldquo;</button> in der Kopfzeile der App oder in einem der <strong>monatlichen DEX-Calls</strong>.</>
                : <>DEX is designed as a <strong>self-service platform</strong>: you create your event independently and manage it end to end — the wizard guides you step by step. Support is available in the <button type="button" className="dex-ui-textlink" onClick={() => { onClose(); navigate('manual'); }}>manual</button>, via <button type="button" className="dex-ui-textlink" onClick={openQuestions}>“Have a question?”</button> in the app header, or in one of the <strong>monthly DEX calls</strong>.</>}
            </span>
          </div>
        </section>

        {/* Features — Aufklapper, Standard zu */}
        <section className="dex-ui-section">
          <button
            type="button"
            className={cx('dex-ui-disclosure', showFeatures && 'is-open')}
            aria-expanded={showFeatures}
            onClick={() => setShowFeatures(v => !v)}
          >
            <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
            {isDE ? 'Alle Funktionen im Überblick' : 'All features at a glance'}
            <span className="dex-ui-disclosure-count">{features.length}</span>
          </button>
          {showFeatures && (
            <div className="dex-ui-disclosure-body dex-ui-fade-in">
              <div className="dex-ui-grid-auto">
                {features.map((f, i) => (
                  <div key={i} className="dex-ui-card dex-ui-card--soft dex-ui-card--hover" style={{ padding: 14 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6 }}>
                      <span className="dex-ui-choice-icon" aria-hidden="true"><Icon iconName={f.icon} style={{ fontSize: 16 }} /></span>
                      <div style={{ fontWeight: 700, fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>{f.title}</div>
                    </div>
                    <div className="dex-ui-muted" style={{ lineHeight: 1.5 }}>{f.body}</div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </section>

        {/* Status */}
        <section className="dex-ui-section">
          <h4 className="dex-ui-section-title">{isDE ? 'Wo DEX heute steht' : 'Where DEX stands today'}</h4>
          <p className="dex-ui-muted" style={{ margin: 0, fontSize: '0.86rem', lineHeight: 1.6 }}>
            {isDE
              ? 'DEX ist aktuell in der Pilotphase mit mehreren Flagship-Events: SAP All Hands Event (ca. 1000 Teilnehmer), SR&T P/MD/D Meeting (450 Teilnehmer), Assistenz Meeting 2026 (130 Teilnehmer), Sommerfest Berlin 2026, verschiedene B2Run-Läufe. Neue Events und Funktionen kommen laufend dazu.'
              : 'DEX is currently in pilot with several flagship events: SAP All Hands Event (~1000 participants), SR&T P/MD/D Meeting (450 participants), Assistenz Meeting 2026 (130 participants), Summer party Berlin 2026, several B2Run races. New events and features are added continuously.'}
          </p>
        </section>

        {/* Interesse? — der Knopf dazu sitzt im Fuß */}
        <section className="dex-ui-section">
          <h4 className="dex-ui-section-title">{isDE ? 'Interesse?' : 'Interested?'}</h4>
          <p className="dex-ui-muted" style={{ margin: 0, fontSize: '0.86rem', lineHeight: 1.6 }}>
            {isDE
              ? 'Dein Event oder dein Bereich möchte DEX nutzen? Der Knopf unten öffnet eine E-Mail an das DEX-Team. Wir schalten dich als Organizer frei — das Event legst du anschließend eigenständig an.'
              : 'Your event or department would like to use DEX? The button below opens an email to the DEX team. We will set you up as an organizer — you then create the event independently.'}
          </p>
          <p className="dex-ui-muted" style={{ margin: '16px 0 0', fontSize: '0.76rem', textAlign: 'center' }}>
            {isDE ? 'Entwickelt von ' : 'Built by '}
            <strong>Eike Brenneisen</strong> {isDE ? 'und' : 'and'} <strong>Nils Felten</strong>.
          </p>
        </section>
      </div>
    </Modal>
  );
}
