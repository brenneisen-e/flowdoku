/**
 * v31.15 — Wartelisten-Blocker: ein Platzhalter im Kalender, solange jemand
 * wartet.
 *
 * Nutzer-Wunsch 11.09.2026. Bis dahin bekam eine wartende Person GAR nichts
 * im Kalender — die Anmeldeseite sagte das auch so („Einen Outlook-Termin
 * gibt es dafür erst mit dem Platz."). Wer auf Platz 2 steht, hält sich den
 * Termin aber trotzdem frei; ohne Eintrag tut das niemand.
 *
 * ## Warum ein EIGENER Termin und nicht der echte
 *
 * Der naheliegende Weg wäre, die wartende Person über den vorhandenen
 * `Einladen`-Pfad in den echten Outlook-Termin zu schreiben. Das ist falsch,
 * und zwar aus zwei Richtungen: Der Organizer sähe Wartende in der
 * Teilnehmerliste des Termins (sie sind keine Teilnehmer), und die Person
 * bekäme eine Einladung, die „du bist dabei" bedeutet (ist sie nicht).
 *
 * Ein Blocker ist deshalb ein **separater Termin** mit genau einer
 * eingeladenen Person, Betreff „Warteliste: …", und er steht **mit
 * Vorbehalt** im Kalender (Nutzer-Entscheidung 11.09.2026). Das ist keine
 * Kosmetik: „Frei" würde von jedem Terminplaner überbucht, „Gebucht" wäre
 * eine Zusage, die es nicht gibt.
 *
 * ## Wo der Blocker wieder verschwindet — und warum das hier steht
 *
 * Ein Blocker, der bleibt, ist schlimmer als keiner. Er verschwindet an
 * genau EINER Stelle: `queueOutlookEvent`. Jedes `Einladen` (Nachrücken)
 * und jedes `Ausladen` (Abmeldung) räumt ihn vorher weg — dadurch hängt das
 * Aufräumen an den Wegen, die es ohnehin schon gibt, statt an fünf neuen
 * Aufrufstellen, die man beim nächsten Feature vergisst (die Lehre aus
 * `queueIDReorderChecked`, v30.80).
 *
 * ## Wo die Kennung des Blockers liegt
 *
 * Nirgends neu. Der Flow schreibt die `iCalUId` des angelegten Termins in
 * die `CalendarLink`-Spalte DERSELBEN Queue-Zeile zurück; zum Löschen sucht
 * die App die jüngste `BlockerSetzen`-Zeile dieser Person zu diesem Event
 * und gibt deren Link mit. Eine eigene Spalte auf den Teilnehmerlisten
 * wäre eine zweite Wahrheit — und müsste auf jeder Subsite nachgezogen
 * werden.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

/** Der Piggyback-Schlüssel in `EmailTemplateOverrides` des Events. */
export const WAITLIST_BLOCKER_KEY = '_waitlistBlocker';

/** Ist der Platzhalter für dieses Event eingeschaltet? Vorgabe: nein. */
export function waitlistBlockerEnabled(overridesJson: string | undefined | null): boolean {
  try {
    const o = JSON.parse(overridesJson || '{}') || {};
    return o[WAITLIST_BLOCKER_KEY] === true;
  } catch {
    return false;
  }
}

const esc = (v: string): string => (v || '').replace(/'/g, "''");

/**
 * Den Platzhalter anfordern.
 *
 * Legt KEINEN zweiten an, wenn schon einer offen ist — sonst stünden nach
 * zwei Anmeldeversuchen zwei Einträge im Kalender, und der zweite ließe sich
 * nie wieder finden.
 */
export async function queueWaitlistBlocker(
  svc: EventService,
  attendee: string,
  eventId: string,
  eventTitle: string,
): Promise<boolean> {
  const lc = (attendee || '').trim();
  if (!lc || !eventId) return false;
  try {
    const vorhanden = await findBlockerRow(svc, lc, eventId);
    if (vorhanden.gefunden) return true;
    const r = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`,
      {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_OutlookListItem' },
        'Title': `BlockerSetzen: ${eventTitle}`,
        'Attendee': lc,
        'EventId': eventId,
        'ActionType': 'BlockerSetzen',
        'Status': 'Pending',
      },
    );
    return r.ok;
  } catch (e) {
    console.warn('[DEX] queueWaitlistBlocker:', e);
    return false;
  }
}

/**
 * Den Platzhalter wieder abräumen.
 *
 * Rückgabe `false` heißt „es wurde nichts abgeräumt" — das kann bedeuten,
 * dass es keinen gab (der Normalfall bei einer Anmeldung mit freiem Platz)
 * oder dass die Queue nicht lesbar war. Der Aufrufer darf daraus NICHTS
 * über den Kalender schließen; die Funktion ist best-effort und läuft
 * neben dem eigentlichen Schreibvorgang, nicht vor ihm.
 */
export async function clearWaitlistBlocker(
  svc: EventService,
  attendee: string,
  eventId: string,
  eventTitle: string,
): Promise<boolean> {
  const lc = (attendee || '').trim();
  if (!lc || !eventId) return false;
  try {
    const treffer = await findBlockerRow(svc, lc, eventId);
    // Ohne Kennung kann der Flow den Termin nicht finden. Eine Lösch-Zeile
    // ohne CalendarLink wäre eine Zeile, die garantiert scheitert — und in
    // der Run history wie ein echter Fehler aussieht.
    if (!treffer.gefunden || !treffer.calendarLink) return false;
    const r = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`,
      {
        '__metadata': { 'type': 'SP.Data.DEX_x005f_OutlookListItem' },
        'Title': `BlockerLoeschen: ${eventTitle}`,
        'Attendee': lc,
        'EventId': eventId,
        'ActionType': 'BlockerLoeschen',
        'Status': 'Pending',
        'CalendarLink': treffer.calendarLink,
      },
    );
    return r.ok;
  } catch (e) {
    console.warn('[DEX] clearWaitlistBlocker:', e);
    return false;
  }
}

/**
 * Die jüngste Blocker-Zeile dieser Person zu diesem Event.
 *
 * `gefunden` ist getrennt von `calendarLink`, weil beides verschiedene
 * Dinge sagt: „es gibt eine Zeile" (also keine zweite anlegen) und „der
 * Flow hat den Termin angelegt und die Kennung zurückgeschrieben" (also
 * löschbar). Eine Zeile ohne Link ist eine, die der Flow noch nicht
 * abgearbeitet hat.
 */
async function findBlockerRow(
  svc: EventService,
  attendee: string,
  eventId: string,
): Promise<{ gefunden: boolean; calendarLink: string }> {
  try {
    const r = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`
      + `?$filter=EventId eq '${esc(eventId)}' and Attendee eq '${esc(attendee)}' and ActionType eq 'BlockerSetzen'`
      + `&$select=Id,CalendarLink&$orderby=Id desc&$top=1`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    if (!r.ok) return { gefunden: false, calendarLink: '' };
    const d = await r.json();
    const row = (d.value || [])[0];
    if (!row) return { gefunden: false, calendarLink: '' };
    return { gefunden: true, calendarLink: row.CalendarLink || '' };
  } catch {
    return { gefunden: false, calendarLink: '' };
  }
}
