/**
 * Info-Modal auf der Landing-Page.
 *
 * Erklärt die DEX Event Experience Platform für neue User, die noch nicht
 * wissen was die App kann. Strukturiert in Abschnitte: Pitch - Zielgruppe -
 * Features - Sicherheit - Status - Kontakt.
 *
 * Zweisprachig DE + EN.
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { X, Mail } from './Icons';
import { APP_VERSION } from '../version';

interface Props {
  open: boolean;
  locale: 'de' | 'en';
  onClose: () => void;
}

interface Feature {
  icon: string;
  title: string;
  body: string;
}

const GREEN = 'var(--dex-green)';
const GREEN_DARK = 'var(--dex-green-dark, #6b9a1e)';
const GRAY_700 = 'var(--dex-gray-700)';

function FeatureList({ items }: { items: Feature[] }): React.ReactElement {
  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 14 }}>
      {items.map((f, i) => (
        <div key={i} style={{
          background: 'var(--dex-gray-50, #f9fafb)',
          border: '1px solid var(--dex-gray-200)',
          borderRadius: 10, padding: 14,
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6 }}>
            <span style={{
              width: 34, height: 34, borderRadius: 8, flexShrink: 0,
              background: GREEN, color: '#fff',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
            }}>
              <Icon iconName={f.icon} style={{ fontSize: 16, color: '#fff' }} />
            </span>
            <div style={{ fontWeight: 700, fontSize: '0.92rem', color: GRAY_700 }}>{f.title}</div>
          </div>
          <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>{f.body}</div>
        </div>
      ))}
    </div>
  );
}

function SectionHeading({ children }: { children: React.ReactNode }): React.ReactElement {
  return (
    <h3 style={{
      fontSize: '1.05rem', fontWeight: 700, color: GREEN_DARK,
      borderBottom: `2px solid ${GREEN}`, paddingBottom: 6, marginTop: 24, marginBottom: 14,
    }}>
      {children}
    </h3>
  );
}

export default function LandingInfoModal({ open, locale, onClose }: Props): React.ReactElement | null {
  if (!open) return null;

  const isDE = locale === 'de';

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

  return (
    <div
      style={{
        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 10000,
        display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24,
      }}
      onClick={onClose}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: '#fff', borderRadius: 16, width: '100%', maxWidth: 960,
          maxHeight: '90vh', overflowY: 'auto',
          boxShadow: '0 24px 60px rgba(0,0,0,0.28)',
          position: 'relative',
        }}
      >
        {/* Sticky Close-Button */}
        <button
          onClick={onClose}
          aria-label="Close"
          style={{
            position: 'sticky', top: 12, float: 'right', marginRight: 12, zIndex: 5,
            background: 'rgba(0,0,0,0.6)', color: '#fff', border: 'none', borderRadius: '50%',
            width: 36, height: 36, display: 'flex', alignItems: 'center', justifyContent: 'center',
            cursor: 'pointer',
          }}
        >
          <X size={18} />
        </button>

        {/* Hero */}
        <div style={{
          padding: '32px 32px 24px',
          background: 'linear-gradient(135deg, #86bc25 0%, #0076a8 55%, #00bcd4 100%)',
          color: '#fff', borderRadius: '16px 16px 0 0',
        }}>
          <div style={{ fontSize: '0.78rem', letterSpacing: 2, textTransform: 'uppercase', opacity: 0.85, marginBottom: 6 }}>
            DEX · {isDE ? 'Event Experience Platform' : 'Event Experience Platform'} · v{APP_VERSION}
          </div>
          <h2 style={{ fontSize: '1.8rem', fontWeight: 700, margin: 0, marginBottom: 8, lineHeight: 1.2 }}>
            {isDE
              ? 'Deine zentrale Plattform für alle Deloitte-Events'
              : 'Your central platform for all Deloitte events'}
          </h2>
          <p style={{ fontSize: '0.95rem', lineHeight: 1.5, opacity: 0.95, margin: 0, maxWidth: 720 }}>
            {isDE
              ? 'Von der Ausschreibung bis zum Check-In in einer App. Registrierung, Outlook-Einladungen, Warteliste, Mass-Mails, QR-Codes, Dokumente — alles integriert in den Deloitte-SharePoint-Tenant.'
              : 'From announcement to check-in in a single app. Registration, Outlook invites, waitlist, mass emails, QR codes, documents — all integrated in the Deloitte SharePoint tenant.'}
          </p>
        </div>

        <div style={{ padding: '8px 32px 32px' }}>
          {/* Für wen eignet sich DEX */}
          <SectionHeading>{isDE ? 'Für welche Events eignet sich DEX?' : 'What kind of events is DEX good for?'}</SectionHeading>
          <ul style={{ margin: 0, paddingLeft: 22, lineHeight: 1.7, fontSize: '0.9rem', color: GRAY_700 }}>
            {isDE ? (
              <>
                <li><strong>Leadership & Strategy Meetings</strong> — z.B. SR&amp;T P/MD/D Meeting mit 450 Teilnehmern</li>
                <li><strong>Firmen-Events</strong> — Sommerfeste, Weihnachtsfeiern, Abteilungs-Offsites</li>
                <li><strong>Assistenz- &amp; Team-Meetings</strong> — mit Transfer- und Hotelbuchung</li>
                <li><strong>Lauf-Events</strong> — B2Run, JPMorgan Corporate Challenge (mit Startblöcken, Split-Capacity)</li>
                <li><strong>Alles dazwischen</strong> — von 10 Leuten am Lunch bis 500+ Personen auf einer Großveranstaltung</li>
              </>
            ) : (
              <>
                <li><strong>Leadership &amp; strategy meetings</strong> — e.g. SR&amp;T P/MD/D with 450+ participants</li>
                <li><strong>Company events</strong> — summer parties, Christmas celebrations, team offsites</li>
                <li><strong>Assistant &amp; team meetings</strong> — with transfer and hotel booking</li>
                <li><strong>Running events</strong> — B2Run, JPMorgan Corporate Challenge (with start-blocks &amp; split capacity)</li>
                <li><strong>Everything in between</strong> — from 10 people at lunch to 500+ at a flagship event</li>
              </>
            )}
          </ul>

          {/* Features */}
          <SectionHeading>{isDE ? 'Alle Funktionen im Überblick' : 'All features at a glance'}</SectionHeading>
          <FeatureList items={isDE ? featuresDE : featuresEN} />

          {/* So funktioniert es */}
          <SectionHeading>{isDE ? 'So funktioniert es' : 'How it works'}</SectionHeading>
          <ol style={{ margin: 0, paddingLeft: 22, lineHeight: 1.7, fontSize: '0.9rem', color: GRAY_700 }}>
            {isDE ? (
              <>
                <li><strong>Organizer</strong> legt ein Event an — entweder selbst via Wizard oder wir helfen beim ersten Mal</li>
                <li>App erstellt automatisch eine <strong>dedizierte SharePoint-Subsite</strong> für das Event mit Teilnehmerliste</li>
                <li>Je nach Standort- und Zielgruppen-Filter sehen <strong>nur berechtigte Kollegen</strong> das Event in ihrer App</li>
                <li>Teilnehmer registrieren sich mit einem Klick — automatische Mail-Bestätigung und Outlook-Termin</li>
                <li>Bei vollem Event kommen sie auf die <strong>Warteliste</strong>; bei Absagen wird automatisch nachgerückt</li>
                <li>Der Organizer verwaltet Teilnehmer im <strong>Admin Center</strong>, versendet Massen-Mails, fügt Dokumente hinzu, erstellt Quiz</li>
                <li>Am Event-Tag: <strong>QR-Code-Check-In</strong> mit Live-Statistik. Fertig.</li>
              </>
            ) : (
              <>
                <li>An <strong>organizer</strong> creates an event — either via wizard themselves or we help on their first try</li>
                <li>The app automatically creates a <strong>dedicated SharePoint subsite</strong> with participant list</li>
                <li>Based on location and audience filter, <strong>only eligible colleagues</strong> see the event in their app</li>
                <li>Participants register with one click — automatic email confirmation and Outlook invite</li>
                <li>If the event is full, they go to the <strong>waitlist</strong>; when someone cancels, the next is promoted automatically</li>
                <li>The organizer manages attendees in the <strong>Admin Center</strong>, sends mass emails, uploads documents, sets up quizzes</li>
                <li>On event day: <strong>QR-code check-in</strong> with live statistics. Done.</li>
              </>
            )}
          </ol>

          {/* Status */}
          <SectionHeading>{isDE ? 'Aktueller Status' : 'Current status'}</SectionHeading>
          <p style={{ fontSize: '0.9rem', color: GRAY_700, lineHeight: 1.6, margin: 0 }}>
            {isDE
              ? 'DEX ist aktuell in der Pilotphase mit mehreren Flagship-Events: SAP All Hands Event (ca. 1000 Teilnehmer), SR&T P/MD/D Meeting (450 Teilnehmer), Assistenz Meeting 2026 (130 Teilnehmer), Sommerfest Berlin 2026, verschiedene B2Run-Läufe. Neue Events und Funktionen kommen laufend dazu.'
              : 'DEX is currently in pilot with several flagship events: SAP All Hands Event (~1000 participants), SR&T P/MD/D Meeting (450 participants), Assistenz Meeting 2026 (130 participants), Summer party Berlin 2026, several B2Run races. New events and features are added continuously.'}
          </p>

          {/* Interesse? */}
          <SectionHeading>{isDE ? 'Interesse?' : 'Interested?'}</SectionHeading>
          <p style={{ fontSize: '0.9rem', color: GRAY_700, lineHeight: 1.6, marginTop: 0, marginBottom: 16 }}>
            {isDE
              ? 'Dein Event oder deine Abteilung will DEX nutzen? Schreib uns — wir melden uns schnell und helfen beim Einrichten.'
              : 'Your event or department wants to use DEX? Drop us a line — we\'ll respond quickly and help with setup.'}
          </p>
          <a
            href="mailto:ebrenneisen@deloitte.de;nifelten@deloitte.de?subject=DEX Event Experience Platform – Interesse"
            style={{
              display: 'inline-flex', alignItems: 'center', gap: 8,
              padding: '10px 20px', borderRadius: 8,
              background: GREEN, color: '#fff', textDecoration: 'none',
              fontWeight: 600, fontSize: '0.9rem',
            }}
          >
            <Mail size={16} /> {isDE ? 'Kontakt aufnehmen' : 'Get in touch'}
          </a>

          {/* Credits */}
          <div style={{
            marginTop: 32, paddingTop: 16,
            borderTop: '1px solid var(--dex-gray-200)',
            fontSize: '0.78rem', color: 'var(--dex-gray-500)', textAlign: 'center',
          }}>
            {isDE ? 'Entwickelt von ' : 'Built by '}
            <strong>Eike Brenneisen</strong> {isDE ? 'und' : 'and'} <strong>Nils Felten</strong>.
          </div>
        </div>
      </div>
    </div>
  );
}
