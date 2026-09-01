/* myEventsHelpers — aus MyEventsPage.tsx ausgelagert (Zeilen 46-154 des
 * urspruenglichen Stands, v30.65). Reine Modul-Ebene: Formatierer, Status-
 * Beschriftungen, der Antwort-Chip und der Eintrags-Typ der Seite. Der Code
 * ist zeichengleich uebernommen, ergaenzt sind nur die `export`-Schluesselwoerter.
 */
import * as React from 'react';
import { DeloitteEvent } from '../../types';
import { SPRegistration } from '../../services/EventService';

// v19.34: People-Picker-Antworten (Feldtyp `user`/`roommate`) im „Meine
// Events"-Antwort-Tag mit Profilfoto statt als Rohtext „Name <email>"
// anzeigen — analog zum Chip im People-Picker selbst.
const parsePersonAnswer = (v: string): { name: string; email: string } | null => {
  const m = (v || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
  if (!m) return null;
  return { name: m[1].trim(), email: m[2].trim() };
};

export function FieldAnswerTag(props: { label: string; value: string; type?: string; small?: boolean }): React.ReactElement {
  const { label, value, type, small } = props;
  const person = (type === 'user' || type === 'roommate') ? parsePersonAnswer(value) : null;
  const baseStyle: React.CSSProperties = {
    fontSize: small ? '0.72rem' : '0.78rem',
    padding: small ? '3px 8px' : '4px 10px',
    borderRadius: 4,
    background: 'rgba(134,188,37,0.14)',
    color: 'var(--dex-green-dark, #4a7c1f)',
    border: '1px solid rgba(134,188,37,0.30)',
  };
  if (person) {
    return (
      <span style={{ ...baseStyle, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        {label}:
        <img
          src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(person.email)}&size=L`}
          alt={person.name}
          onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
          style={{ width: 22, height: 22, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', transition: 'transform 0.15s', transformOrigin: 'center' }}
          onMouseEnter={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(2.6)'; (e.currentTarget as HTMLImageElement).style.zIndex = '20'; (e.currentTarget as HTMLImageElement).style.position = 'relative'; (e.currentTarget as HTMLImageElement).style.boxShadow = '0 6px 18px rgba(0,0,0,0.18)'; }}
          onMouseLeave={e => { (e.currentTarget as HTMLImageElement).style.transform = 'scale(1)'; (e.currentTarget as HTMLImageElement).style.boxShadow = 'none'; }}
        />
        <strong>{person.name}</strong>
      </span>
    );
  }
  return (
    <span style={baseStyle}>
      {label}: <strong>{value}</strong>
    </span>
  );
}

export interface MyEventEntry {
  event: DeloitteEvent;
  registration: SPRegistration;
  /** Seit v6.14: wenn true, ist der User NICHT direkt fürs Parent-Event angemeldet,
   *  sondern nur für mindestens eine Sub-Event-Session. Die Parent-Karte dient dann
   *  nur als Container, damit die Session-Registrierungen sichtbar und verwaltbar
   *  bleiben. Das Status-Badge zeigt "Nur Sessions" statt dem echten Parent-Status. */
  sessionsOnly?: boolean;
  /** v11.31: Titel der Sub-Events für die der User aktiv angemeldet ist —
   *  wird bei sessionsOnly-Entries befüllt, damit der Hinweis-Text die
   *  konkreten Sub-Event-Namen in Klammern ausgeben kann. */
  subEventTitles?: string[];
  /** v28.23: Die Anmeldung ist im zentralen Teilnehmer-Register (DEX_Participants)
   *  belegt, die ZEILE in der Teilnehmerliste ist für die Person aber nicht
   *  lesbar — das passiert, solange eine stellvertretend angelegte Zeile noch
   *  der Assistenz gehört (Item-Level-Security „nur eigene Elemente"). Früher
   *  fiel der Eintrag dadurch komplett aus „Meine Events" heraus und die Person
   *  hielt sich für nicht angemeldet. Jetzt wird er ANGEZEIGT, aber ohne
   *  Detaildaten und ohne Selbst-Abmeldung (die scheitert an denselben
   *  Rechten) — mit Hinweis, sich an die Assistenz/Organizer zu wenden. */
  hiddenRow?: boolean;
}

export function formatDate(iso: string): string {
  if (!iso) return '-';
  return new Date(iso).toLocaleDateString('de-DE', {
    day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit',
  });
}

export function formatDateRange(start: string, end: string): string {
  if (!start) return '-';
  const s = new Date(start);
  const e = end ? new Date(end) : null;
  // v27.8: Wochentag mit anzeigen (z.B. „Mittwoch, 09.09.2026").
  const sDate = s.toLocaleDateString('de-DE', { weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric' });
  const sTime = s.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  if (!e) return `${sDate}, ${sTime}`;
  const eDate = e.toLocaleDateString('de-DE', { weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric' });
  const eTime = e.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  // Gleicher Tag: "14.04.2026, 14:00 – 18:00"
  if (sDate === eDate) return `${sDate}, ${sTime} – ${eTime}`;
  // Verschiedene Tage
  return `${sDate}, ${sTime} – ${eDate}, ${eTime}`;
}

export function getStatusBadgeClass(status: string): string {
  switch (status) {
    case 'Angemeldet': return 'badge-green';
    case 'QR versendet': return 'badge-green';
    case 'Warteliste': return 'badge-orange';
    case 'Abgemeldet': return 'badge-red';
    case 'Eingecheckt': return 'badge-green';
    default: return 'badge-gray';
  }
}

export function getStatusLabel(status: string, t: (key: string) => string): string {
  switch (status) {
    case 'Angemeldet': return t('status.registered');
    case 'Warteliste': return t('status.waitlist');
    case 'Abgemeldet': return t('status.cancelled');
    case 'Eingecheckt': return t('status.checkedin');
    default: return status;
  }
}
