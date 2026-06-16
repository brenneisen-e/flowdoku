/**
 * v22.77: DEX-Onepager — Querformat-Übersicht aller Funktionen „auf einen
 * Blick". Wird vom Admin über das Info-Icon der Landing Page geöffnet
 * (Button „Onepager" im LandingInfoModal). Vier Cluster:
 *   1. Event-Anlage
 *   2. An- & Abmeldung (inkl. Sichtbarkeitsprüfung)
 *   3. Kommunikation (Mail + Outlook im Deloitte-Wrapper)
 *   4. Teilnehmerliste / Organizer Center (Koordination)
 *
 * Druckbar: „Drucken / als PDF" nutzt window.print(); ein Print-Stylesheet
 * blendet die App aus und setzt A4 quer.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { X } from './Icons';
import { APP_VERSION } from '../version';
import {
  DemoEmailPreview, DemoParticipantTable, DemoEventCard, DemoQRCode,
  DemoWizardProgress, DemoAdminCenter, DemoCheckInScanner,
} from './manual/ManualMockups';
// v22.77: Echte App-Screens als interaktive Vorschau (gleicher Mechanismus wie
// im Handbuch — rendert die echten Komponenten mit Demo-Daten in einem Modal).
import { AppPreview } from './manual/previews/AppPreview';
import { DEMO_EVENT_ID } from './manual/previews/PreviewProviders';
import Header from './Header';
import RegistrationPage from './RegistrationPage';
import EventCreationPage from './EventCreationPage';
import AdminPage from './AdminPage';
import CheckInPage from './CheckInPage';

interface Props {
  open: boolean;
  locale: 'de' | 'en';
  onClose: () => void;
}

interface Cluster {
  icon: string;
  color: string;
  title: string;
  items: React.ReactNode[];
}

// Galerie-Kachel: rahmt ein Demo-Visual mit Bildunterschrift.
function GalleryCard({ caption, children }: { caption: string; children: React.ReactNode }): React.ReactElement {
  return (
    <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, overflow: 'hidden', background: '#fff', display: 'flex', flexDirection: 'column' }}>
      <div style={{ padding: '7px 12px', fontSize: '0.78rem', fontWeight: 700, color: 'var(--dex-gray-700)', background: 'var(--dex-gray-50, #f7f8f9)', borderBottom: '1px solid var(--dex-gray-200)' }}>{caption}</div>
      <div style={{ padding: 12, overflow: 'auto', maxHeight: 360, background: 'var(--dex-gray-50, #fafafa)' }}>{children}</div>
    </div>
  );
}

const PRINT_CSS = `
@media print {
  body * { visibility: hidden !important; }
  .dex-onepager-overlay, .dex-onepager-overlay * { visibility: visible !important; }
  .dex-onepager-overlay { position: absolute !important; inset: 0 !important; background: #fff !important; padding: 0 !important; overflow: visible !important; display: block !important; }
  .dex-onepager-page { box-shadow: none !important; border-radius: 0 !important; max-height: none !important; width: 100% !important; max-width: none !important; }
  .dex-onepager-noprint { display: none !important; }
  @page { size: A4 landscape; margin: 8mm; }
}
`;

export default function OnePager({ open, locale, onClose }: Props): React.ReactElement | null {
  if (!open) return null;
  const isDe = locale === 'de';

  const clusters: Cluster[] = isDe ? [
    {
      icon: 'Calendar', color: '#86bc25', title: '1 · Event anlegen',
      items: [
        <>Geführter <strong>Wizard</strong>: Grundlagen, Ort &amp; Programm, Kapazität &amp; Sichtbarkeit, Felder, Kommunikation, Dokumente, Quiz</>,
        <>Datum/Zeit, strukturierte <strong>Adresse</strong>, Bild, Organizer- &amp; Check-in-Team</>,
        <><strong>Teilnehmerlimit</strong> + optionale Warteliste; geteilte Kapazität (zwei Gruppen)</>,
        <>Eigene <strong>Abfragen/Felder</strong> (Text, Auswahl, Zahl, Checkbox, Dokument-Upload) — je Pflicht/optional</>,
        <><strong>Sub-Events / „Klammer“</strong>: mehrere Programmpunkte unter einem Event</>,
        <><strong>Team-Anmeldung</strong> (mehrere Personen auf einmal)</>,
        <>Demo-Daten, Live-Vorschau, <strong>Live/Entwurf</strong>-Schalter</>,
      ],
    },
    {
      icon: 'ContactCard', color: '#0076a8', title: '2 · An- & Abmeldung + Sichtbarkeit',
      items: [
        <><strong>Sichtbarkeit</strong>: Standortfilter, Mailverteiler (Entra), einzelne Personen, UND/ODER, Ausschluss-Liste</>,
        <><strong>„Sichtbarkeit prüfen“</strong>: zeigt alle Personen, die das Event sehen — Verteiler live aufgelöst, je Sub-Event</>,
        <><strong>1-Klick-Anmeldung</strong> mit den event-spezifischen Feldern</>,
        <><strong>Warteliste</strong> mit Position + automatischem Nachrücken</>,
        <>Selbst-<strong>Abmeldung</strong> bis Event-Ende; kommunizierte Abmelde-Frist</>,
        <><strong>Stellvertretende</strong> Anmeldung (für andere / externe Personen, mit Zustimmungsnachweis)</>,
        <>„Meine Events“ mit persönlichem <strong>QR-Code</strong></>,
      ],
    },
    {
      icon: 'MailSchedule', color: '#62b5e5', title: '3 · Kommunikation (Mail + Outlook)',
      items: [
        <>Automatische Mails: <strong>Anmeldung, Warteliste, Abmeldung, Nachrücken</strong> — im <strong>Deloitte-Layout</strong> (Logo, Farben, DE/EN)</>,
        <><strong>Outlook-Termine</strong> aus no_reply.events@deloitte.de; Event-Änderungen aktualisieren alle Termine</>,
        <><strong>Massenmail</strong> (RichText) an Teilnehmergruppen; <strong>Einladungsmail</strong> mit Anmelde-Link</>,
        <><strong>QR-Code-Mail</strong> zum Check-in; Texte pro Event anpassbar</>,
        <>Schalter pro Event: Mails/Outlook an/aus, Anmelde-/Abmelde-Mail einzeln</>,
        <>Outlook-Absage = optionale <strong>Auto-Abmeldung</strong></>,
      ],
    },
    {
      icon: 'AccountManagement', color: '#26890d', title: '4 · Teilnehmerliste / Organizer Center',
      items: [
        <>Teilnehmer <strong>suchen, sortieren</strong>, Spalten anpassen</>,
        <>Status ändern, <strong>ein-/auschecken</strong> (auch QR-Self-Check-in)</>,
        <>Manuell an-/abmelden, Felder pro Teilnehmer bearbeiten</>,
        <><strong>Excel-Export</strong> (Zielgruppen; konsolidierte Matrix bei Sub-Events)</>,
        <>Konsolidierte <strong>Sub-Event-Ansicht</strong> (Klammer): eine Zeile pro Person, ✓ je Sub-Event</>,
        <><strong>Teams</strong> verwalten, <strong>Audit-Log</strong>, Aktionen-Dropdown</>,
        <>Live-Statistik: Angemeldet / QR versendet / Eingecheckt / Abgemeldet</>,
      ],
    },
  ] : [
    {
      icon: 'Calendar', color: '#86bc25', title: '1 · Create an event',
      items: [
        <>Guided <strong>wizard</strong>: basics, location &amp; programme, capacity &amp; visibility, fields, communication, documents, quiz</>,
        <>Date/time, structured <strong>address</strong>, image, organizer &amp; check-in team</>,
        <><strong>Capacity limit</strong> + optional waitlist; split capacity (two groups)</>,
        <>Custom <strong>fields</strong> (text, choice, number, checkbox, document upload) — required/optional each</>,
        <><strong>Sub-events / “bracket”</strong>: several programme parts under one event</>,
        <><strong>Team registration</strong> (several people at once)</>,
        <>Demo data, live preview, <strong>live/draft</strong> toggle</>,
      ],
    },
    {
      icon: 'ContactCard', color: '#0076a8', title: '2 · Registration & visibility',
      items: [
        <><strong>Visibility</strong>: location filter, distribution lists (Entra), individual people, AND/OR, exclusion list</>,
        <><strong>“Check visibility”</strong>: shows everyone who can see the event — lists resolved live, per sub-event</>,
        <><strong>1-click registration</strong> with the event-specific fields</>,
        <><strong>Waitlist</strong> with position + automatic promotion</>,
        <>Self-<strong>cancellation</strong> until the event ends; communicated cancellation deadline</>,
        <><strong>Proxy</strong> registration (for others / external people, with consent record)</>,
        <>“My Events” with a personal <strong>QR code</strong></>,
      ],
    },
    {
      icon: 'MailSchedule', color: '#62b5e5', title: '3 · Communication (mail + Outlook)',
      items: [
        <>Automated emails: <strong>registration, waitlist, cancellation, promotion</strong> — in the <strong>Deloitte layout</strong> (logo, colors, EN/DE)</>,
        <><strong>Outlook invites</strong> from no_reply.events@deloitte.de; event changes update all invites</>,
        <><strong>Mass email</strong> (rich text) to participant groups; <strong>invitation email</strong> with sign-up link</>,
        <><strong>QR-code email</strong> for check-in; texts customizable per event</>,
        <>Per-event switches: emails/Outlook on/off, registration/cancellation mail individually</>,
        <>Outlook decline = optional <strong>auto-cancellation</strong></>,
      ],
    },
    {
      icon: 'AccountManagement', color: '#26890d', title: '4 · Attendee list / Organizer Center',
      items: [
        <><strong>Search, sort</strong> attendees, customize columns</>,
        <>Change status, <strong>check in/out</strong> (also QR self check-in)</>,
        <>Register/cancel manually, edit fields per attendee</>,
        <><strong>Excel export</strong> (target groups; consolidated matrix for sub-events)</>,
        <>Consolidated <strong>sub-event view</strong> (bracket): one row per person, ✓ per sub-event</>,
        <>Manage <strong>teams</strong>, <strong>audit log</strong>, actions dropdown</>,
        <>Live stats: registered / QR sent / checked-in / cancelled</>,
      ],
    },
  ];

  return (
    <div
      className="dex-onepager-overlay"
      style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)', zIndex: 10050, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16, overflowY: 'auto' }}
      onClick={onClose}
    >
      <style>{PRINT_CSS}</style>
      <div
        className="dex-onepager-page"
        onClick={e => e.stopPropagation()}
        style={{ background: '#fff', borderRadius: 12, width: '100%', maxWidth: 1180, boxShadow: '0 24px 60px rgba(0,0,0,0.3)', overflow: 'hidden', position: 'relative' }}
      >
        {/* Aktionsleiste (nicht drucken) */}
        <div className="dex-onepager-noprint" style={{ position: 'absolute', top: 12, right: 12, display: 'flex', gap: 8, zIndex: 5 }}>
          <button
            type="button"
            onClick={() => window.print()}
            style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '7px 14px', borderRadius: 8, border: 'none', background: 'rgba(255,255,255,0.92)', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, fontSize: '0.8rem', cursor: 'pointer', fontFamily: 'inherit' }}
          >
            <Icon iconName="Print" style={{ fontSize: 14 }} /> {isDe ? 'Drucken / als PDF' : 'Print / save as PDF'}
          </button>
          <button type="button" onClick={onClose} aria-label="Close" style={{ background: 'rgba(0,0,0,0.55)', color: '#fff', border: 'none', borderRadius: '50%', width: 34, height: 34, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer' }}>
            <X size={16} />
          </button>
        </div>

        {/* Kopf */}
        <div style={{ padding: '20px 28px 16px', background: 'linear-gradient(135deg, #2d5016 0%, #86bc25 55%, #0076a8 100%)', color: '#fff' }}>
          <div style={{ fontSize: '0.72rem', letterSpacing: 2, textTransform: 'uppercase', opacity: 0.85 }}>DEX · Event Experience Platform · v{APP_VERSION}</div>
          <h1 style={{ margin: '4px 0 0', fontSize: '1.5rem', fontWeight: 700 }}>
            {isDe ? 'Alle Funktionen auf einen Blick' : 'All features at a glance'}
          </h1>
          <p style={{ margin: '4px 0 0', fontSize: '0.85rem', opacity: 0.95 }}>
            {isDe
              ? 'Von der Event-Anlage über An-/Abmeldung und Kommunikation bis zur Teilnehmer-Koordination — alles im Deloitte-SharePoint-Tenant.'
              : 'From event creation through registration and communication to attendee coordination — all inside the Deloitte SharePoint tenant.'}
          </p>
        </div>

        {/* 4 Cluster im Querformat-Raster */}
        <div style={{ padding: '16px 20px 20px', display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 14 }}>
          {clusters.map((c, i) => (
            <div key={i} style={{ border: `1px solid var(--dex-gray-200)`, borderTop: `4px solid ${c.color}`, borderRadius: 10, overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 12px', background: 'var(--dex-gray-50, #f7f8f9)', borderBottom: '1px solid var(--dex-gray-200)' }}>
                <span style={{ width: 30, height: 30, borderRadius: 7, background: c.color, color: '#fff', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                  <Icon iconName={c.icon} style={{ fontSize: 15, color: '#fff' }} />
                </span>
                <span style={{ fontWeight: 700, fontSize: '0.86rem', color: 'var(--dex-gray-800)', lineHeight: 1.2 }}>{c.title}</span>
              </div>
              <ul style={{ margin: 0, padding: '10px 12px 12px 26px', display: 'flex', flexDirection: 'column', gap: 7 }}>
                {c.items.map((it, j) => (
                  <li key={j} style={{ fontSize: '0.76rem', color: 'var(--dex-gray-700)', lineHeight: 1.4 }}>{it}</li>
                ))}
              </ul>
            </div>
          ))}
        </div>

        {/* Eindrücke aus der App — echte Demo-Visuals (inkl. Deloitte-Mail) */}
        <div style={{ padding: '4px 20px 8px' }}>
          <div style={{ fontWeight: 700, fontSize: '1rem', color: 'var(--dex-green-dark, #4a7c1f)', borderBottom: '2px solid var(--dex-green)', paddingBottom: 6, marginBottom: 14 }}>
            {isDe ? 'Eindrücke aus der App' : 'Impressions from the app'}
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: 16, alignItems: 'start' }}>
            <GalleryCard caption={isDe ? 'Automatische Mail im Deloitte-Layout' : 'Automated mail in Deloitte layout'}><DemoEmailPreview /></GalleryCard>
            <GalleryCard caption={isDe ? 'Event-Karte (Anmeldung)' : 'Event card (registration)'}><DemoEventCard /></GalleryCard>
            <GalleryCard caption={isDe ? 'Persönlicher Check-in-QR' : 'Personal check-in QR'}><DemoQRCode /></GalleryCard>
            <GalleryCard caption={isDe ? 'Teilnehmerliste (Organizer Center)' : 'Attendee list (Organizer Center)'}><DemoParticipantTable /></GalleryCard>
            <GalleryCard caption={isDe ? 'Admin Center' : 'Admin Center'}><DemoAdminCenter /></GalleryCard>
            <GalleryCard caption={isDe ? 'Event-Wizard' : 'Event wizard'}><DemoWizardProgress /></GalleryCard>
            <GalleryCard caption={isDe ? 'QR-Check-in am Event-Tag' : 'QR check-in on event day'}><DemoCheckInScanner scanned /></GalleryCard>
          </div>
        </div>

        {/* Echte Screens interaktiv — gleicher Mechanismus wie im Handbuch
            (rendert die echten Komponenten mit Demo-Daten). Nicht drucken. */}
        <div className="dex-onepager-noprint" style={{ padding: '4px 20px 12px' }}>
          <div style={{ fontWeight: 700, fontSize: '1rem', color: 'var(--dex-green-dark, #4a7c1f)', borderBottom: '2px solid var(--dex-green)', paddingBottom: 6, marginBottom: 12 }}>
            {isDe ? 'Echte Screens ansehen (interaktive Vorschau)' : 'See the real screens (interactive preview)'}
          </div>
          <p style={{ margin: '0 0 12px', fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
            {isDe
              ? 'Klick auf eine Vorschau — die echte App-Ansicht öffnet sich mit Demo-Daten (read-only).'
              : 'Click a preview — the real app view opens with demo data (read-only).'}
          </p>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: 12 }}>
            <AppPreview label={isDe ? 'Anmeldeseite (Teilnehmer)' : 'Registration page (attendee)'} role="User" page="registration" selectedEventId={DEMO_EVENT_ID} width={1024} device="laptop">
              <Header />
              <RegistrationPage />
            </AppPreview>
            <AppPreview label={isDe ? 'Event-Wizard (Anlage)' : 'Event wizard (creation)'} role="Organizer" page="create-event" width={1024} device="laptop">
              <Header />
              <EventCreationPage />
            </AppPreview>
            <AppPreview label={isDe ? 'Organizer Center (Teilnehmerliste)' : 'Organizer Center (attendee list)'} role="Organizer" page="admin" selectedEventId={DEMO_EVENT_ID} width={1024} device="laptop">
              <Header />
              <AdminPage />
            </AppPreview>
            <AppPreview label={isDe ? 'Check-in am Event-Tag' : 'Check-in on event day'} role="Organizer" page="check-in" selectedEventId={DEMO_EVENT_ID} width={1024} device="laptop">
              <Header />
              <CheckInPage />
            </AppPreview>
          </div>
        </div>

        {/* Fuß */}
        <div style={{ padding: '0 20px 16px', display: 'flex', flexWrap: 'wrap', gap: 16, justifyContent: 'space-between', alignItems: 'center', fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
          <span>
            {isDe
              ? 'Rollen: User (eigene An-/Abmeldung) · Organizer (eigene Events) · Admin (Vollzugriff). Item-Level-Security: jeder sieht nur seine eigenen Anmeldungen.'
              : 'Roles: User (own registration) · Organizer (own events) · Admin (full access). Item-level security: each user sees only their own registrations.'}
          </span>
          <span style={{ fontWeight: 600, color: 'var(--dex-green-dark, #4a7c1f)' }}>deudeloitte.sharepoint.com · no_reply.events@deloitte.de</span>
        </div>
      </div>
    </div>
  );
}
