/**
 * v29.24: Einführungs-Onepager (Admin-only) — die eine Seite für die
 * Einführungsveranstaltung. Vier Abschnitte, in der Reihenfolge des Vortrags:
 *
 *  1. Der Event-Management-Zyklus (Planung → Durchführung → Abrechnung) mit
 *     Markierung, welche Teile DEX übernimmt — die Übergabe der Teilnehmer
 *     an F&A ist bewusst als „geplant" ausgewiesen, nicht als vorhanden.
 *  2. Ein Venn-Diagramm zum Einsatzbereich: interne Events mit
 *     Deloitte-Teilnehmern und die Koordination der Deloitte-Teilnahme an
 *     externen Events sind DEX; externe Events mit externen Teilnehmern
 *     gehören zum Event-Management-Team.
 *  3. Die zwei Rollen (Organizer, User) — mit dem Hinweis, dass der Fokus
 *     der Veranstaltung auf der Organizer-Rolle liegt.
 *  4. Die vier Kernfunktionen.
 *
 * Wie die Architekturseite rein deklarativ aus statischen Daten — der
 * Zyklus ist Vortrags-Wissen, kein Tenant-Zustand. Zweisprachig DE/EN.
 */
import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { Users, Mail, Calendar, QrCode, GraduationCap } from './Icons';

const GREEN = 'var(--dex-green, #86bc25)';
const GREEN_DARK = 'var(--dex-green-dark, #4a7c1f)';
const ORANGE = 'var(--dex-orange, #ed8b00)';
const ORANGE_DARK = 'var(--dex-orange-dark, #b35a00)';

export default function IntroOnePagerPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { isAdmin, originalIsAdmin } = useRoles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const adminLike = isAdmin || originalIsAdmin;

  React.useEffect(() => {
    if (!adminLike) navigate('start');
  }, [adminLike, navigate]);
  if (!adminLike) return <div className="page-container" />;

  // ---- 1 · Event-Management-Zyklus ----------------------------------------
  // dex: 'yes' = übernimmt DEX heute, 'future' = geplant, undefined = liegt
  // beim Organizer/anderen Teams (DEX ist bewusst KEIN Budget- oder
  // Locationtool — das gehört auf die Folie, sonst wird es hineingewünscht).
  type CycleItem = { de: string; en: string; dex?: 'yes' | 'future' };
  const phases: Array<{ de: string; en: string; subDe: string; subEn: string; items: CycleItem[] }> = [
    {
      de: 'Planung', en: 'Planning',
      subDe: 'vor dem Event', subEn: 'before the event',
      items: [
        { de: 'Konzept, Budget & Genehmigung', en: 'Concept, budget & approval' },
        { de: 'Ort, Termin & Dienstleister', en: 'Venue, date & vendors' },
        { de: 'Teilnehmermanagement: Einladung, An-/Abmeldung, Warteliste', en: 'Attendee management: invitation, (de)registration, waitlist', dex: 'yes' },
        { de: 'Kommunikation & Termine über Outlook', en: 'Communication & appointments via Outlook', dex: 'yes' },
      ],
    },
    {
      de: 'Durchführung', en: 'Execution',
      subDe: 'am Event-Tag', subEn: 'on event day',
      items: [
        { de: 'Aufbau, Ablauf & Betreuung vor Ort', en: 'Setup, schedule & on-site care' },
        { de: 'Check-in der Teilnehmer (QR-Code oder Liste)', en: 'Attendee check-in (QR code or list)', dex: 'yes' },
        { de: 'Aktuelle Teilnehmerliste jederzeit im Zugriff', en: 'Live attendee list at hand', dex: 'yes' },
      ],
    },
    {
      de: 'Abrechnung', en: 'Settlement',
      subDe: 'nach dem Event', subEn: 'after the event',
      items: [
        { de: 'Rechnungen, Kostenstellen & interne Verrechnung', en: 'Invoices, cost centers & internal charging' },
        { de: 'Übermittlung der Teilnehmer an F&A', en: 'Handover of attendees to F&A', dex: 'future' },
      ],
    },
  ];

  // ---- 3 · Rollen ---------------------------------------------------------
  const roles = [
    {
      focus: true,
      title: 'Organizer',
      de: 'Legt Events an und betreut sie: Anmeldeformular, Sichtbarkeit, Kapazitäten, Mails, Check-in — alles über das Organizer Center.',
      en: 'Creates and runs events: registration form, visibility, capacities, mails, check-in — all via the Organizer Center.',
    },
    {
      focus: false,
      title: 'User',
      de: 'Sieht die für sie/ihn sichtbaren Events, meldet sich an oder ab und findet alles Weitere unter „Meine Events" — ohne Einarbeitung.',
      en: 'Sees the events visible to them, registers or cancels, and finds everything else under “My events” — no training needed.',
    },
  ];

  // ---- 4 · Kernfunktionen -------------------------------------------------
  const coreFns = [
    {
      icon: <Users size={26} />,
      de: 'Teilnehmermanagement', en: 'Attendee management',
      descDe: 'Anmeldung, Abmeldung und Warteliste mit automatischem Nachrücken — inklusive eigener Anmeldeformulare je Event.',
      descEn: 'Registration, cancellation and waitlist with automatic promotion — including custom registration forms per event.',
    },
    {
      icon: <Mail size={26} />,
      de: 'Kommunikation über Outlook', en: 'Communication via Outlook',
      descDe: 'Bestätigungen, Einladungen, Erinnerungen und Rundmails — versendet über das zentrale Event-Postfach.',
      descEn: 'Confirmations, invitations, reminders and broadcasts — sent via the central event mailbox.',
    },
    {
      icon: <Calendar size={26} />,
      de: 'Terminmanagement', en: 'Appointment management',
      descDe: 'Jede Anmeldung erzeugt automatisch die Outlook-Kalendereinladung; Änderungen und Absagen laufen mit.',
      descEn: 'Every registration automatically creates the Outlook calendar invite; updates and cancellations follow along.',
    },
    {
      icon: <QrCode size={26} />,
      de: 'Check-in', en: 'Check-in',
      descDe: 'Am Event-Tag per QR-Code oder Teilnehmerliste einchecken — der Stand ist live für das ganze Orga-Team sichtbar.',
      descEn: 'Check in on event day via QR code or attendee list — the status is live for the whole organizing team.',
    },
  ];

  // ---- Styles -------------------------------------------------------------
  const section: React.CSSProperties = { border: '1px solid var(--dex-gray-200)', borderRadius: 14, background: '#fff', boxShadow: '0 2px 10px rgba(0,0,0,0.05)', padding: '20px 22px', marginBottom: 18 };
  const kicker: React.CSSProperties = { fontSize: '0.72rem', fontWeight: 700, letterSpacing: '0.08em', textTransform: 'uppercase', color: GREEN_DARK, marginBottom: 4 };
  const h2: React.CSSProperties = { margin: '0 0 4px', fontSize: '1.15rem', color: 'var(--dex-gray-800, #333)' };
  const sub: React.CSSProperties = { margin: '0 0 16px', fontSize: '0.85rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, maxWidth: 860 };

  const dexChip = (kind: 'yes' | 'future'): React.ReactElement => (
    <span style={{
      flexShrink: 0, alignSelf: 'flex-start', fontSize: '0.66rem', fontWeight: 800, letterSpacing: '0.04em',
      padding: '2px 7px', borderRadius: 8, whiteSpace: 'nowrap',
      background: kind === 'yes' ? GREEN : 'transparent',
      color: kind === 'yes' ? '#fff' : ORANGE_DARK,
      border: kind === 'yes' ? `1px solid ${GREEN}` : `1px dashed ${ORANGE}`,
    }}>
      {kind === 'yes' ? 'DEX' : (isDe ? 'DEX · geplant' : 'DEX · planned')}
    </span>
  );

  return (
    <div className="page-container">
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center', marginBottom: 14 }}>
        <button className="btn btn-secondary" onClick={() => navigate('admin-hub')}>
          {isDe ? '← Zurück zum Admin' : '← Back to admin'}
        </button>
      </div>

      {/* Kopf */}
      <div style={{ marginBottom: 18 }}>
        <div style={{ display: 'inline-flex', alignItems: 'center', gap: 8, ...kicker }}>
          <GraduationCap size={16} /> {isDe ? 'Einführungsveranstaltung' : 'Introduction session'}
        </div>
        <h1 style={{ margin: '0 0 6px' }}>{isDe ? 'DEX auf einen Blick' : 'DEX at a glance'}</h1>
        <p style={{ margin: 0, color: 'var(--dex-gray-600)', maxWidth: 860, lineHeight: 1.55 }}>
          {isDe
            ? 'Die DEX Event Experience Platform übernimmt das Teilnehmermanagement interner Events — von der Einladung über die Anmeldung bis zum Check-in am Event-Tag. Diese Seite ist der rote Faden für die Einführung: wo DEX im Event-Zyklus sitzt, wofür es gedacht ist, welche Rollen es gibt und was es kann.'
            : 'The DEX Event Experience Platform handles attendee management for internal events — from invitation through registration to check-in on event day. This page is the guiding thread for the introduction: where DEX sits in the event cycle, what it is meant for, which roles exist and what it does.'}
        </p>
      </div>

      {/* 1 · Zyklus */}
      <div style={section}>
        <div style={kicker}>1</div>
        <h2 style={h2}>{isDe ? 'Der Event-Management-Zyklus — und wo DEX übernimmt' : 'The event management cycle — and where DEX takes over'}</h2>
        <p style={sub}>
          {isDe
            ? 'Ein Event läuft von der Planung über die Durchführung bis zur Abrechnung. DEX ist das Werkzeug für alles rund um die Teilnehmer — Konzept, Budget, Location und die Abrechnung selbst bleiben beim Organizer bzw. den zuständigen Teams.'
            : 'An event runs from planning through execution to settlement. DEX is the tool for everything around the attendees — concept, budget, venue and the settlement itself stay with the organizer or the responsible teams.'}
        </p>
        <div style={{ display: 'flex', alignItems: 'stretch', gap: 0, flexWrap: 'wrap' }}>
          {phases.map((ph, i) => (
            <React.Fragment key={i}>
              {i > 0 && (
                <div style={{ display: 'flex', alignItems: 'center', padding: '0 4px', color: GREEN, fontSize: 26, fontWeight: 700 }} aria-hidden="true">›</div>
              )}
              <div style={{
                flex: '1 1 240px', minWidth: 240, border: '1px solid var(--dex-gray-200)', borderRadius: 12,
                background: 'var(--dex-gray-50, #fafafa)', overflow: 'hidden', display: 'flex', flexDirection: 'column',
              }}>
                <div style={{ padding: '10px 14px', background: '#fff', borderBottom: '1px solid var(--dex-gray-200)' }}>
                  <span style={{ fontWeight: 800, fontSize: '0.95rem', color: 'var(--dex-gray-800, #333)' }}>{isDe ? ph.de : ph.en}</span>
                  <span style={{ marginLeft: 8, fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{isDe ? ph.subDe : ph.subEn}</span>
                </div>
                <div style={{ padding: '10px 12px', display: 'flex', flexDirection: 'column', gap: 8, flex: 1 }}>
                  {ph.items.map((it, k) => (
                    <div key={k} style={{
                      display: 'flex', gap: 8, alignItems: 'flex-start', justifyContent: 'space-between',
                      padding: '7px 10px', borderRadius: 9, fontSize: '0.8rem', lineHeight: 1.4,
                      background: it.dex === 'yes' ? 'rgba(134,188,37,0.12)' : (it.dex === 'future' ? 'rgba(237,139,0,0.08)' : '#fff'),
                      border: it.dex === 'yes' ? `1px solid ${GREEN}` : (it.dex === 'future' ? `1px dashed ${ORANGE}` : '1px solid var(--dex-gray-200)'),
                      color: 'var(--dex-gray-700)',
                    }}>
                      <span>{isDe ? it.de : it.en}</span>
                      {it.dex && dexChip(it.dex)}
                    </div>
                  ))}
                </div>
              </div>
            </React.Fragment>
          ))}
        </div>
        <div style={{ display: 'flex', gap: 18, flexWrap: 'wrap', marginTop: 12, fontSize: '0.76rem', color: 'var(--dex-gray-600)' }}>
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
            <span style={{ width: 12, height: 12, borderRadius: 4, background: 'rgba(134,188,37,0.25)', border: `1px solid ${GREEN}`, display: 'inline-block' }} />
            {isDe ? 'übernimmt DEX heute' : 'DEX does this today'}
          </span>
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
            <span style={{ width: 12, height: 12, borderRadius: 4, background: 'rgba(237,139,0,0.12)', border: `1px dashed ${ORANGE}`, display: 'inline-block' }} />
            {isDe ? 'geplant — kommt in einer künftigen Ausbaustufe' : 'planned — comes in a future stage'}
          </span>
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
            <span style={{ width: 12, height: 12, borderRadius: 4, background: '#fff', border: '1px solid var(--dex-gray-300)', display: 'inline-block' }} />
            {isDe ? 'liegt beim Organizer / anderen Teams' : 'stays with the organizer / other teams'}
          </span>
        </div>
      </div>

      {/* 2 · Venn: Einsatzbereich */}
      <div style={section}>
        <div style={kicker}>2</div>
        <h2 style={h2}>{isDe ? 'Wofür ist DEX da — und wofür nicht?' : 'What is DEX for — and what not?'}</h2>
        <p style={sub}>
          {isDe
            ? 'DEX ist für Events mit Deloitte-Teilnehmern. Bei externen Veranstaltungen koordiniert DEX ausschließlich die Deloitte-Teilnahme; das externe Event selbst — mit externen Teilnehmern — betreut das Event-Management-Team.'
            : 'DEX is for events with Deloitte attendees. For external events, DEX only coordinates Deloitte participation; the external event itself — with external attendees — is run by the event management team.'}
        </p>
        <div style={{ maxWidth: 780, margin: '0 auto' }}>
          <svg viewBox="0 0 780 330" role="img" style={{ width: '100%', height: 'auto', display: 'block' }}
            aria-label={isDe ? 'Venn-Diagramm: Einsatzbereich von DEX und Event-Management-Team' : 'Venn diagram: scope of DEX and the event management team'}>
            {/* Kreis DEX (links, grün) */}
            <circle cx="285" cy="180" r="140" fill="rgba(134,188,37,0.16)" stroke="#86bc25" strokeWidth="2.5" />
            {/* Kreis Event-Management-Team (rechts, grau) */}
            <circle cx="495" cy="180" r="140" fill="rgba(120,120,120,0.10)" stroke="#9c9c9c" strokeWidth="2.5" />
            {/* Kreis-Titel */}
            <text x="240" y="26" textAnchor="middle" fontSize="17" fontWeight="800" fill="#4a7c1f">DEX</text>
            <text x="540" y="26" textAnchor="middle" fontSize="15" fontWeight="700" fill="#666">
              {isDe ? 'Event-Management-Team' : 'Event management team'}
            </text>
            <line x1="240" y1="34" x2="252" y2="52" stroke="#86bc25" strokeWidth="1.5" />
            <line x1="540" y1="34" x2="528" y2="52" stroke="#9c9c9c" strokeWidth="1.5" />
            {/* Links: nur DEX */}
            <text x="212" y="150" textAnchor="middle" fontSize="13.5" fontWeight="700" fill="#3d3d3d">
              <tspan x="212" dy="0">{isDe ? 'Interne Events' : 'Internal events'}</tspan>
              <tspan x="212" dy="19">{isDe ? 'mit Deloitte-' : 'with Deloitte'}</tspan>
              <tspan x="212" dy="19">{isDe ? 'Teilnehmern' : 'attendees'}</tspan>
            </text>
            <text x="212" y="226" textAnchor="middle" fontSize="15" fontWeight="800" fill="#4a7c1f">✓ DEX</text>
            {/* Mitte: Schnittmenge */}
            <text x="390" y="132" textAnchor="middle" fontSize="12.5" fontWeight="700" fill="#3d3d3d">
              <tspan x="390" dy="0">{isDe ? 'Externe Events:' : 'External events:'}</tspan>
              <tspan x="390" dy="17">{isDe ? 'Koordination der' : 'coordinating the'}</tspan>
              <tspan x="390" dy="17">{isDe ? 'Deloitte-Teilnahme' : 'Deloitte participation'}</tspan>
            </text>
            <text x="390" y="212" textAnchor="middle" fontSize="14" fontWeight="800" fill="#4a7c1f">✓ DEX</text>
            {/* Rechts: nur EM-Team */}
            <text x="568" y="150" textAnchor="middle" fontSize="13.5" fontWeight="700" fill="#3d3d3d">
              <tspan x="568" dy="0">{isDe ? 'Externe Events' : 'External events'}</tspan>
              <tspan x="568" dy="19">{isDe ? 'mit externen' : 'with external'}</tspan>
              <tspan x="568" dy="19">{isDe ? 'Teilnehmern' : 'attendees'}</tspan>
            </text>
            <text x="568" y="226" textAnchor="middle" fontSize="13" fontWeight="800" fill="#8a1f1f">
              {isDe ? '✕ nicht DEX' : '✕ not DEX'}
            </text>
          </svg>
        </div>
        <p style={{ margin: '10px 0 0', fontSize: '0.8rem', color: 'var(--dex-gray-600)', textAlign: 'center' }}>
          {isDe
            ? 'Merksatz: Sobald Deloitte-Mitarbeitende teilnehmen, hilft DEX — das externe Event selbst organisiert das Event-Management-Team.'
            : 'Rule of thumb: as soon as Deloitte people attend, DEX helps — the external event itself is organized by the event management team.'}
        </p>
      </div>

      {/* 3 · Rollen */}
      <div style={section}>
        <div style={kicker}>3</div>
        <h2 style={h2}>{isDe ? 'Zwei Rollen' : 'Two roles'}</h2>
        <p style={sub}>
          {isDe
            ? 'In DEX gibt es zwei Rollen: Organizer richten Events ein und betreuen sie, User nehmen teil. Der Fokus der heutigen Einführung liegt auf der Organizer-Rolle.'
            : 'DEX has two roles: organizers set up and run events, users attend. Today’s introduction focuses on the organizer role.'}
        </p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: 14 }}>
          {roles.map((r, i) => (
            <div key={i} style={{
              border: r.focus ? `2px solid ${GREEN}` : '1px solid var(--dex-gray-200)',
              borderRadius: 12, padding: '16px 18px', position: 'relative',
              background: r.focus ? 'rgba(134,188,37,0.06)' : '#fff',
            }}>
              {r.focus && (
                <span style={{
                  position: 'absolute', top: -11, right: 14, background: GREEN, color: '#fff',
                  fontSize: '0.68rem', fontWeight: 800, letterSpacing: '0.05em', padding: '3px 10px', borderRadius: 10,
                }}>
                  {isDe ? 'Fokus heute' : 'Focus today'}
                </span>
              )}
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 8 }}>
                <span style={{ color: r.focus ? GREEN_DARK : 'var(--dex-gray-500)', display: 'inline-flex' }}><Users size={22} /></span>
                <span style={{ fontWeight: 800, fontSize: '1rem', color: r.focus ? GREEN_DARK : 'var(--dex-gray-800, #333)' }}>{r.title}</span>
              </div>
              <p style={{ margin: 0, fontSize: '0.84rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>{isDe ? r.de : r.en}</p>
            </div>
          ))}
        </div>
      </div>

      {/* 4 · Kernfunktionen */}
      <div style={section}>
        <div style={kicker}>4</div>
        <h2 style={h2}>{isDe ? 'Die vier Kernfunktionen' : 'The four core functions'}</h2>
        <p style={sub}>
          {isDe
            ? 'Alles Weitere — Sichtbarkeit, Sub-Events, Formulare, Statistiken — baut auf diesen vier Kernfunktionen auf.'
            : 'Everything else — visibility, sub-events, forms, statistics — builds on these four core functions.'}
        </p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 14 }}>
          {coreFns.map((f, i) => (
            <div key={i} style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 12, padding: '16px 18px', background: '#fff' }}>
              <div style={{
                width: 46, height: 46, borderRadius: 12, background: 'rgba(134,188,37,0.14)', color: GREEN_DARK,
                display: 'inline-flex', alignItems: 'center', justifyContent: 'center', marginBottom: 10,
              }}>
                {f.icon}
              </div>
              <div style={{ fontWeight: 800, fontSize: '0.92rem', color: 'var(--dex-gray-800, #333)', marginBottom: 6 }}>{isDe ? f.de : f.en}</div>
              <p style={{ margin: 0, fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>{isDe ? f.descDe : f.descEn}</p>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
