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
import { Mail, Info, Check, AlertCircle, ChevronDown } from './Icons';
import Modal from './Modal';
import { cx } from './dexUi';
import { APP_VERSION } from '../version';
import { DEX_TEAM_EMAIL } from '../utils/supportContact';

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

/** v31.2: Ein Ablauf-Schritt — kurzer Titel plus eine Zeile, was dabei passiert. */
interface Step { title: string; hint: string; }

/** v31.2: Ein Einsatzbeispiel — Kategorie plus konkrete Beispiele. */
interface UseCase { title: string; sub: string; }

const EVENT_MGMT_URL = 'https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx';

export default function LandingInfoModal({ open, locale, onClose, onStartTutorial }: Props): React.ReactElement | null {
  // v31.2: Die 15 Funktions-Kacheln sind der längste Block des Dialogs. Zum
  // Einstieg reichen Einsatzbereich und Ablauf; die Kacheln bleiben zu, bis
  // jemand sie sehen will — sonst zeigt der Dialog beim Öffnen drei Bildschirme.
  const [showFeatures, setShowFeatures] = React.useState(false);
  if (!open) return null;

  const isDE = locale === 'de';

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
    { icon: 'AccountManagement', title: 'Admin Center für Organizer', body: 'Teilnehmer suchen, sortieren, ein-/auschecken, manuell an-/abmelden. Massen-Mails mit RichText-Editor im Deloitte-Design an alle Bestätigten.' },
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
    { icon: 'AccountManagement', title: 'Admin Center for organizers', body: 'Search, sort, check-in/out participants, register/cancel manually. Mass emails with RichText editor in Deloitte design to all confirmed guests.' },
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
    { title: 'Leadership & Strategy Meetings', sub: 'z.B. SR&T P/MD/D Meeting mit 450 Teilnehmern' },
    { title: 'Firmen-Events', sub: 'Sommerfeste, Weihnachtsfeiern, Abteilungs-Offsites' },
    { title: 'Assistenz- & Team-Meetings', sub: 'mit Transfer- und Hotelbuchung' },
    { title: 'Lauf-Events', sub: 'B2Run, JPMorgan Corporate Challenge (mit Startblöcken, Split-Capacity)' },
    { title: 'Alles dazwischen', sub: 'von 10 Leuten am Lunch bis 500+ Personen auf einer Großveranstaltung' },
  ] : [
    { title: 'Leadership & strategy meetings', sub: 'e.g. SR&T P/MD/D with 450+ participants' },
    { title: 'Company events', sub: 'summer parties, Christmas celebrations, team offsites' },
    { title: 'Assistant & team meetings', sub: 'with transfer and hotel booking' },
    { title: 'Running events', sub: 'B2Run, JPMorgan Corporate Challenge (with start-blocks & split capacity)' },
    { title: 'Everything in between', sub: 'from 10 people at lunch to 500+ at a flagship event' },
  ];

  // v31.2: Ablauf als nummerierte Schritt-Zeilen (Titel + Folge) statt einer
  // Aufzählung mit Fettdruck mitten im Satz — man sieht auf einen Blick, wer
  // was tut und was die App daraufhin automatisch erledigt.
  const steps: Step[] = isDE ? [
    { title: 'Ein Organizer legt das Event an', hint: 'Selbst im Wizard — oder wir helfen beim ersten Mal.' },
    { title: 'Die App legt eine eigene SharePoint-Subsite an', hint: 'Mit Teilnehmerliste nur für dieses Event.' },
    { title: 'Nur berechtigte Kollegen sehen das Event', hint: 'Je nach Standort- und Zielgruppen-Filter erscheint es in ihrer App.' },
    { title: 'Teilnehmer melden sich mit einem Klick an', hint: 'Bestätigungsmail und Outlook-Termin kommen automatisch.' },
    { title: 'Ist das Event voll, geht es auf die Warteliste', hint: 'Bei Absagen rückt automatisch die nächste Person nach.' },
    { title: 'Der Organizer verwaltet alles im Admin Center', hint: 'Teilnehmer, Massen-Mails, Dokumente, Quiz.' },
    { title: 'Am Event-Tag: QR-Code-Check-in', hint: 'Mit Live-Statistik. Fertig.' },
  ] : [
    { title: 'An organizer creates the event', hint: 'Via the wizard — or we help on the first try.' },
    { title: 'The app creates a dedicated SharePoint subsite', hint: 'With a participant list just for this event.' },
    { title: 'Only eligible colleagues see the event', hint: 'Based on location and audience filter it shows up in their app.' },
    { title: 'Participants register with one click', hint: 'Confirmation email and Outlook invite arrive automatically.' },
    { title: 'If the event is full, they join the waitlist', hint: 'When someone cancels, the next person is promoted automatically.' },
    { title: 'The organizer manages everything in the Admin Center', hint: 'Attendees, mass emails, documents, quizzes.' },
    { title: 'On event day: QR-code check-in', hint: 'With live statistics. Done.' },
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
          <div className="dex-ui-card dex-ui-card--soft" style={{ padding: '4px 6px' }}>
            {useCases.map((u, i) => (
              <div key={i} className="dex-ui-row dex-ui-row--bordered">
                <span className="dex-ui-choice-check" aria-hidden="true" style={{ borderColor: 'var(--dex-green)', background: 'var(--dex-green)' }}><Check size={12} /></span>
                <div className="dex-ui-row-main">
                  <div className="dex-ui-row-title">{u.title}</div>
                  <div className="dex-ui-row-sub">{u.sub}</div>
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
          <div className="dex-ui-stack" style={{ gap: 8 }}>
            {steps.map((s, i) => (
              <div key={i} className="dex-ui-step">
                <span className="dex-ui-step-num">{i + 1}</span>
                <div className="dex-ui-step-body">
                  <div className="dex-ui-step-title">{s.title}</div>
                  <div className="dex-ui-step-hint">{s.hint}</div>
                </div>
              </div>
            ))}
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
              ? 'Dein Event oder deine Abteilung will DEX nutzen? Schreib uns — der Knopf unten öffnet eine Mail an uns. Wir melden uns schnell und helfen beim Einrichten.'
              : 'Your event or department wants to use DEX? Drop us a line — the button below opens an email to us. We\'ll respond quickly and help with setup.'}
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
