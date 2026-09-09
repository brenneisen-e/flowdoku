/* myEventsHelpers — aus MyEventsPage.tsx ausgelagert (Zeilen 46-154 des
 * urspruenglichen Stands, v30.65). Reine Modul-Ebene: Formatierer, Status-
 * Beschriftungen, der Antwort-Chip und der Eintrags-Typ der Seite.
 *
 * v31.8: `FieldAnswerTag` nutzt jetzt die gemeinsamen `dex-ui-`-Klassen
 * (Leitfaden Abschnitt 6). Das Stylesheet injiziert die Seiten-Komponente
 * (`MyEventsPage`) ueber `ensureDexUiStyles()`; dieses Modul ruft es nie.
 */
import * as React from 'react';
import { DeloitteEvent } from '../../types';
import { SPRegistration } from '../../services/EventService';
import { cx } from '../dexUi';

// v19.34: People-Picker-Antworten (Feldtyp `user`/`roommate`) im „Meine
// Events"-Antwort-Tag mit Profilfoto statt als Rohtext „Name <email>"
// anzeigen — analog zum Chip im People-Picker selbst.
const parsePersonAnswer = (v: string): { name: string; email: string } | null => {
  const m = (v || '').match(/^(.+?)\s*<([^>]+@[^>]+)>\s*$/);
  if (!m) return null;
  return { name: m[1].trim(), email: m[2].trim() };
};

// v31.8: Der Antwort-Tag ist reine ANZEIGE (der echte Bearbeiten-Weg steht in
// der Aktionszeile der Karte) — also `dex-ui-pill`, nicht der handgebaute
// Kasten mit Radius 4. `--wrap`, weil eine Freitext-Antwort umbrechen darf,
// statt die Zeile aus der Karte zu schieben.
const answerPillClass = (small?: boolean): string =>
  cx('dex-ui-pill', 'dex-ui-pill--green', 'dex-ui-pill--wrap', small && 'dex-ui-pill--sm');

export function FieldAnswerTag(props: { label: string; value: string; type?: string; small?: boolean }): React.ReactElement {
  const { label, value, type, small } = props;
  const person = (type === 'user' || type === 'roommate') ? parsePersonAnswer(value) : null;
  if (person) {
    return (
      <span className={answerPillClass(small)} style={{ gap: 6 }}>
        {label}:
        {/* v31.8: Der 260-%-Zoom beim Überfahren ist ersatzlos weg — er
            versprach eine Aktion, die es nie gab (Leitfaden 6c: kein Hover
            ohne Klick). Das Foto bleibt 22 px, damit die Pille flach bleibt. */}
        <img
          className="dex-ui-avatar dex-ui-avatar--xs"
          src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(person.email)}&size=L`}
          alt={person.name}
          onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
        />
        <strong>{person.name}</strong>
      </span>
    );
  }
  return (
    <span className={answerPillClass(small)}>
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

/**
 * v30.79: Kompaktes von–bis für Termin-Zeilen in Dialogen (ohne Wochentag,
 * damit die Zeile kurz bleibt): „12.10.2026, 13:00 – 14:30" bzw. bei
 * Tageswechsel „12.10.2026, 13:00 – 13.10.2026, 09:00". Nutzer-Befund
 * 07.09.2026: „bei den Subevents fehlt irgendwie immer die zweite Zeit …
 * muss natürlich immer von bis stehen." Ungültige Werte liefern ''.
 */
export function formatDateTimeRange(start?: string, end?: string, isDe: boolean = true): string {
  const loc = isDe ? 'de-DE' : 'en-GB';
  const s = start ? new Date(start) : null;
  if (!s || !isFinite(s.getTime())) return '';
  const e = end ? new Date(end) : null;
  const sDay = s.toLocaleDateString(loc, { day: '2-digit', month: '2-digit', year: 'numeric' });
  const sTime = s.toLocaleTimeString(loc, { hour: '2-digit', minute: '2-digit' });
  if (!e || !isFinite(e.getTime())) return `${sDay}, ${sTime}`;
  const eDay = e.toLocaleDateString(loc, { day: '2-digit', month: '2-digit', year: 'numeric' });
  const eTime = e.toLocaleTimeString(loc, { hour: '2-digit', minute: '2-digit' });
  return sDay === eDay ? `${sDay}, ${sTime} – ${eTime}` : `${sDay}, ${sTime} – ${eDay}, ${eTime}`;
}

export function getStatusBadgeClass(status: string): string {
  switch (status) {
    case 'Angemeldet': return 'badge-green';
    case 'QR versendet': return 'badge-green';
    case 'Warteliste': return 'badge-orange';
    case 'Abgemeldet': return 'badge-red';
    case 'Eingecheckt': return 'badge-green';
    // v31.8: No-Show fiel bis hierher auf grau durch — ausgerechnet der
    // Zustand, den man erklärt bekommen muss.
    case 'No-Show': return 'badge-red';
    default: return 'badge-gray';
  }
}

export function getStatusLabel(status: string, t: (key: string) => string): string {
  switch (status) {
    case 'Angemeldet': return t('status.registered');
    case 'Warteliste': return t('status.waitlist');
    case 'Abgemeldet': return t('status.cancelled');
    case 'Eingecheckt': return t('status.checkedin');
    // v31.8: Die beiden Schlüssel gab es längst in beiden Sprachen, nur die
    // Fälle fehlten — im englischen UI stand deshalb deutsch „QR versendet".
    case 'QR versendet': return t('status.qrsent');
    case 'No-Show': return t('status.noshow');
    // Absichtlich der Rohwert: `Status` ist eine SharePoint-Choice, die
    // erweitert werden kann. Ein unbekannter Zustand muss sichtbar bleiben,
    // nicht zu „" oder „Unbekannt" werden.
    default: return status;
  }
}
