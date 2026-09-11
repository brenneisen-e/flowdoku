/**
 * v31.16 — Der Wartelisten-Platzhalter als SCHATTEN-EVENT.
 *
 * Nutzer-Entwurf 11.09.2026, nachdem zwei eigene Vorschläge verworfen wurden:
 * „kann man nicht einfach ein zusätzliches Schattenevent erzeugen … beim
 * Schattenevent gibt es immer nur das Outlook-Ding, sonst nichts. Sieht der
 * Organizer auch nicht. Und wenn jemand nachrückt, dann wird das Schattenevent
 * für ihn abgesagt und er kriegt ja automatisch aus dem richtigen Event den
 * Outlook-Eintrag."
 *
 * Das trägt, und es ist die mit Abstand billigste Lösung — **sie braucht KEINE
 * einzige Flow-Änderung.** Der Grund liegt in zwei Eigenschaften, die DEX
 * schon hat:
 *
 *  1. `DEX_CreateOutlookEvent` triggert auf JEDE neue Zeile in `DEX_Events`
 *     und liest davon genau vierzehn Spalten (Title, StartDate, EndDate,
 *     OutlookBody/Subject/Location/Start/End, AllDay, ShowAsFree,
 *     SkipOrganizerInvite, OrganizerEmail, OutlookIsOnlineMeeting, ID). Eine
 *     Zeile mit diesen Feldern bekommt also von selbst einen Kalendertermin,
 *     und der Flow schreibt die `CalendarLink` zurück.
 *  2. `DEX_Outlook_Einladungen` lädt über `Einladen`/`Ausladen` Leute in
 *     einen Termin hinein und wieder heraus — egal, zu welchem Event die
 *     Zeile gehört.
 *
 * Damit ist der Platzhalter: ein zweites, unsichtbares Event, in das Wartende
 * ganz normal eingeladen werden. Nachrücken = `Ausladen` beim Schatten,
 * `Einladen` beim echten Event. Beides gibt es längst.
 *
 * ## Was die beiden verworfenen Wege gekostet hätten
 *
 *  - **Eigener Termin je Person über neue Auftragsarten** (v31.15): ein
 *    kompletter zweiter Flow, neue `ActionType`-Werte, und eine
 *    Buchführung darüber, welche `iCalUId` zu wem gehört.
 *  - **Wartende als `optional`-Teilnehmer des echten Termins**: zwei
 *    geänderte Ausdrücke, aber alle lesen denselben Termintext („du bist
 *    angemeldet"), und Wartende stünden in der Teilnehmerliste des
 *    Organizers.
 *
 * Der Schatten hat beides nicht: eigener Titel, eigener Text, und der
 * Organizer sieht ihn nirgends.
 *
 * ## Die eine Stelle, an der er unsichtbar wird
 *
 * `loadEvents` in `EventContext` filtert Schattenzeilen direkt nach dem
 * Mapping heraus. Das ist der Flaschenhals, durch den JEDE Ansicht der App
 * geht — Kacheln, Listen, Organizer Center, Kennzahlen. Ein Filter weiter
 * unten wäre einer von zwanzig, und der einundzwanzigste wäre der, an dem
 * das Schattenevent doch auftaucht.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';
import { DeloitteEvent } from '../../types';
import { APP_URL } from '../EmailTemplates';

/** Piggyback in `EmailTemplateOverrides` der SCHATTEN-Zeile: die Id des echten Events. */
export const WAITLIST_SHADOW_KEY = '_waitlistShadowFor';
/** Piggyback auf dem ECHTEN Event: Platzhalter AUSgeschaltet? Vorgabe: an (v31.17). */
export const WAITLIST_BLOCKER_KEY = '_waitlistBlocker';

/**
 * Ist der Platzhalter für dieses Event eingeschaltet?
 *
 * v31.17: **Vorgabe ist AN** (Nutzer-Entscheidung 11.09.2026). Deshalb steht
 * der Schlüssel nur im Blob, wenn jemand ihn AUSgeschaltet hat — nicht, wenn
 * er an ist. Andersherum wäre der Unterschied zwischen „nie entschieden" und
 * „bewusst an" nicht mehr lesbar, und Bestandsevents (die den Schlüssel gar
 * nicht haben) blieben für immer aus, während der Assistent „an" anzeigt.
 * Genau diese zwei Wahrheiten wollen wir nicht.
 */
export function waitlistBlockerEnabled(overridesJson: string | undefined | null): boolean {
  try {
    const o = JSON.parse(overridesJson || '{}') || {};
    return o[WAITLIST_BLOCKER_KEY] !== false;
  } catch {
    // Kein lesbarer Blob heisst „nie etwas eingestellt" — also die Vorgabe.
    return true;
  }
}

/**
 * Ist das eine Schatten-Zeile? Wird in `loadEvents` gefragt, also bei jedem
 * Start für jede Zeile — deshalb bewusst billig und ohne Ausnahme-Pfad.
 */
export function isWaitlistShadow(ev: Pick<DeloitteEvent, 'emailTemplateOverrides'> | null | undefined): boolean {
  try {
    const o = JSON.parse(ev?.emailTemplateOverrides || '{}') || {};
    return typeof o[WAITLIST_SHADOW_KEY] === 'string' && !!o[WAITLIST_SHADOW_KEY];
  } catch {
    return false;
  }
}

const esc = (v: string): string => (v || '').replace(/'/g, "''");

/** Der Titel, den die wartende Person im Kalender sieht. */
export function shadowTitleFor(eventTitle: string): string {
  return `Warteliste: ${eventTitle}`;
}

/**
 * Der Termintext des Platzhalters.
 *
 * v31.19, drei Nachträge aus dem ersten Live-Test (11.09.2026):
 *
 *  1. **Der DEX-Link gehört hinein.** Der Termin ist die einzige Spur, die die
 *     wartende Person von DEX hat — der Satz „melde dich in der DEX App ab"
 *     ohne Link ist eine Aufgabe, keine Hilfe.
 *  2. **Sprache je Event, Vorgabe Englisch.** `registrationLanguage` ist das
 *     Feld, an dem schon die Anmeldeseite hängt; ist es leer, gilt Englisch —
 *     ein englischer Text ist für deutschsprachige Lesende verständlich,
 *     andersherum nicht.
 *  3. **Der Platz auf der Warteliste steht in der App.** Ohne diesen Hinweis
 *     ist der Termin eine Sackgasse: Er sagt „du wartest", aber nicht, worauf
 *     man schauen könnte.
 */
function shadowBody(ev: DeloitteEvent): string {
  const t = (ev.title || '').replace(/</g, '&lt;');
  const link = '<a href="' + APP_URL + '">' + APP_URL + '</a>';
  if (ev.registrationLanguage === 'de') {
    return '<p>Du stehst bei <strong>' + t + '</strong> auf der Warteliste.</p>'
      + '<p>Dieser Eintrag h&auml;lt dir den Termin frei und ist <strong>keine Zusage</strong>. '
      + 'Sobald ein Platz frei wird, bekommst du die richtige Einladung — dieser Platzhalter '
      + 'wird dann automatisch abgesagt.</p>'
      + '<p>Deinen Platz auf der Warteliste siehst du jederzeit in der DEX App unter '
      + '&bdquo;Meine Events&ldquo;. Dort meldest du dich auch von der Warteliste ab, wenn du '
      + 'nicht mehr warten m&ouml;chtest — der Platzhalter verschwindet dann ebenfalls.</p>'
      + '<p>' + link + '</p>';
  }
  return '<p>You are on the waiting list for <strong>' + t + '</strong>.</p>'
    + '<p>This entry keeps the slot free in your calendar and is <strong>not a confirmation</strong>. '
    + 'As soon as a place opens up you will receive the real invitation — this placeholder '
    + 'is then cancelled automatically.</p>'
    + '<p>You can check your position on the waiting list at any time in the DEX app under '
    + '&ldquo;My events&rdquo;. That is also where you leave the waiting list if you no longer '
    + 'want to wait — the placeholder disappears with it.</p>'
    + '<p>' + link + '</p>';
}

/**
 * Der Betreff, den die wartende Person im Kalender liest — sprachabhängig.
 *
 * **Bewusst getrennt von `shadowTitleFor`.** Der Listen-`Title` ist der
 * Schlüssel, über den App UND Flow (`Schatten_suchen` filtert auf
 * `concat('Warteliste: ', …)`) die Schatten-Zeile wiederfinden. Würde er die
 * Sprache tragen, bräche jede Sprachumstellung die Verbindung zum Termin —
 * und der Flow müsste mitgeändert werden. Der Kalender liest `OutlookSubject`;
 * dort kostet die Übersetzung nichts.
 */
function shadowSubjectFor(ev: DeloitteEvent): string {
  return ev.registrationLanguage === 'de'
    ? `Warteliste: ${ev.title}`
    : `Waiting list: ${ev.title}`;
}

/**
 * Die Schatten-Zeile zu einem Event finden — oder anlegen.
 *
 * Angelegt wird sie ERST, wenn wirklich jemand wartet (gerufen aus dem
 * Warteliste-Zweig der Anmeldung). Ein Schatten je Event im Voraus wäre eine
 * zusätzliche Zeile in `DEX_Events` und ein zusätzlicher Outlook-Termin für
 * jedes Event, bei dem nie jemand wartet.
 *
 * Rückgabe `null` heißt „konnte nicht sichergestellt werden" — der Aufrufer
 * lässt die Anmeldung dann trotzdem laufen. Ein fehlender Platzhalter ist
 * ärgerlich; eine verhinderte Anmeldung wäre ein Schaden.
 */
export async function ensureWaitlistShadow(
  svc: EventService,
  ev: DeloitteEvent,
): Promise<{ id: string; title: string } | null> {
  const titel = shadowTitleFor(ev.title);
  try {
    // 1) Gibt es sie schon? Gesucht wird über den Titel — die Piggyback-Spalte
    //    ist ein Note-Feld und in SharePoint nicht filterbar.
    const vorhanden = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`
      + `?$filter=Title eq '${esc(titel)}'&$select=Id,EmailTemplateOverrides&$top=5`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    if (vorhanden.ok) {
      const d = await vorhanden.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const row of ((d.value || []) as any[])) {
        if (isWaitlistShadow({ emailTemplateOverrides: row.EmailTemplateOverrides })) {
          try {
            const o = JSON.parse(row.EmailTemplateOverrides || '{}');
            if (String(o[WAITLIST_SHADOW_KEY]) === String(ev.id)) {
              // v31.19: Betreff und Text der BESTEHENDEN Zeile nachziehen —
              // sie wurde beim ersten Speichern des Events angelegt und trägt
              // sonst für immer den Stand von damals (falsche Sprache, kein
              // DEX-Link). Bewusst ohne `UpdateEvent`-Auftrag: Der
              // Kalendertermin bleibt, wie er ist; erst ein neu angelegter
              // Termin liest den neuen Text. Eine „Aktualisiert"-Mail an alle
              // Wartenden nur wegen eines Textnachtrags wäre teurer als der
              // Nachtrag wert ist.
              try {
                await svc._merge(
                  `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${row.Id})`,
                  { 'OutlookSubject': shadowSubjectFor(ev), 'OutlookBody': shadowBody(ev) },
                );
              } catch { /* best-effort */ }
              return { id: String(row.Id), title: titel };
            }
          } catch { /* naechste Zeile */ }
        }
      }
    } else {
      // Nicht lesbar heisst NICHT „gibt es nicht". Eine zweite Schatten-Zeile
      // waere ein zweiter Kalendertermin, den niemand mehr zuordnen kann.
      console.warn(`[DEX] ensureWaitlistShadow: DEX_Events nicht lesbar (HTTP ${vorhanden.status}) — kein Platzhalter.`);
      return null;
    }

    // 2) Anlegen. Genau die vierzehn Spalten, die DEX_CreateOutlookEvent liest
    //    — keine Subsite, keine Teilnehmerliste, keine Anmeldung. Das
    //    Schattenevent ist ausschliesslich ein Kalendertermin.
    const ende = ev.endDate || ev.startDate;
    const r = await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`, {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
      'Title': titel,
      'StartDate': ev.startDate,
      'EndDate': ende,
      // Title = Schlüssel (immer deutsch), OutlookSubject = was im Kalender
      // steht (Sprache des Events). Siehe shadowSubjectFor.
      'OutlookSubject': shadowSubjectFor(ev),
      'OutlookBody': shadowBody(ev),
      'OutlookLocation': ev.location || '',
      'OrganizerEmail': (ev.organizerEmails || []).join('; '),
      // Der Organizer bekommt den Platzhalter NICHT — er ist die einzige
      // Person, fuer die er nichts aussagt (Nutzer-Wunsch: „sieht der
      // Organizer auch nicht").
      'SkipOrganizerInvite': true,
      'AllDay': !!ev.allDay,
      // Bewusst NICHT als „frei": Ein freier Termin wird von jedem
      // Terminplaner ueberbucht, und genau das soll der Platzhalter
      // verhindern. Wer nicht antwortet, sieht ihn in Outlook ohnehin als
      // „mit Vorbehalt" — die Entscheidung vom 11.09.2026.
      'ShowAsFree': false,
      'OutlookIsOnlineMeeting': false,
      'EmailTemplateOverrides': JSON.stringify({ [WAITLIST_SHADOW_KEY]: String(ev.id) }),
    });
    if (!r.ok) {
      console.warn(`[DEX] ensureWaitlistShadow: Anlegen fehlgeschlagen (HTTP ${r.status}).`);
      return null;
    }
    const d = await r.json();
    const neueId = d?.d?.Id ?? d?.Id;
    if (!neueId) return null;
    return { id: String(neueId), title: titel };
  } catch (e) {
    console.warn('[DEX] ensureWaitlistShadow:', e);
    return null;
  }
}

/**
 * Die Schatten-Zeile eines Events suchen, ohne eine anzulegen.
 *
 * Braucht das Abräumen: Wer nachrückt oder sich abmeldet, soll aus dem
 * Platzhalter heraus — aber wenn es keinen gibt, darf dabei keiner entstehen.
 */
export async function findWaitlistShadow(
  svc: EventService,
  eventId: string,
  eventTitle: string,
): Promise<{ id: string; title: string; calendarLink: string } | null> {
  const titel = shadowTitleFor(eventTitle);
  try {
    const r = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`
      + `?$filter=Title eq '${esc(titel)}'&$select=Id,CalendarLink,EmailTemplateOverrides&$top=5`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    if (!r.ok) return null;
    const d = await r.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    for (const row of ((d.value || []) as any[])) {
      try {
        const o = JSON.parse(row.EmailTemplateOverrides || '{}');
        if (String(o[WAITLIST_SHADOW_KEY]) === String(eventId)) {
          return { id: String(row.Id), title: titel, calendarLink: row.CalendarLink || '' };
        }
      } catch { /* naechste */ }
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * Eine wartende Person in den Platzhalter einladen.
 *
 * v31.18 — der Fehler aus dem ersten Test (11.09.2026: „hat nicht geklappt,
 * dass ich als Warteliste einen Termin bekommen habe, obwohl Termin angelegt
 * wurde"): Solange die Schatten-Zeile noch KEINE `CalendarLink` trägt, hat
 * `DEX_CreateOutlookEvent` den Kalendertermin noch nicht angelegt. Eine
 * `Einladen`-Zeile findet dann nichts, `DEX_Outlook_Einladungen` setzt sie auf
 * `Failed` — endgültig, es gibt keinen zweiten Versuch.
 *
 * Deshalb wird hier gar nicht erst eingeladen. Keine Zeile ist besser als eine
 * Zeile, die garantiert scheitert und die Run history rot färbt. Seit der
 * Platzhalter beim ANLEGEN des Events entsteht (nicht mehr beim Anmelden),
 * ist die Kennung ohnehin längst da, wenn sich jemand anmeldet.
 */
export async function inviteToWaitlistShadow(
  svc: EventService,
  attendee: string,
  eventId: string,
  eventTitle: string,
): Promise<boolean> {
  const schatten = await findWaitlistShadow(svc, eventId, eventTitle);
  if (!schatten) return false;
  if (!schatten.calendarLink) {
    console.warn('[DEX] Wartelisten-Platzhalter: Kalendertermin noch nicht angelegt — keine Einladung geschrieben.');
    return false;
  }
  return svc.queueOutlookEvent(attendee, schatten.id, schatten.title, 'Einladen');
}

/**
 * Eine Person aus dem Platzhalter ausladen.
 *
 * Schreibt die Queue-Zeile DIREKT statt über `queueOutlookEvent` — sonst
 * riefe der Aufräum-Haken, der dort drinsteckt, sich selbst wieder auf.
 */
export async function leaveWaitlistShadow(
  svc: EventService,
  attendee: string,
  eventId: string,
  eventTitle: string,
): Promise<boolean> {
  const schatten = await findWaitlistShadow(svc, eventId, eventTitle);
  if (!schatten) return false;
  try {
    const r = await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Outlook')/items`, {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_OutlookListItem' },
      'Title': `Ausladen: ${schatten.title}`,
      'Attendee': attendee,
      'EventId': schatten.id,
      'ActionType': 'Ausladen',
      'Status': 'Pending',
    });
    return r.ok;
  } catch (e) {
    console.warn('[DEX] leaveWaitlistShadow:', e);
    return false;
  }
}

/**
 * Den Platzhalter eines Events entfernen — Zeile und Kalendertermin.
 *
 * Gerufen, wenn das echte Event gelöscht wird: Ein Schattenevent ohne Event
 * ist ein Termin, den niemand mehr absagen kann.
 *
 * Reihenfolge: erst den Termin abbestellen (`DeleteEvent` in die Queue),
 * dann die Zeile recyceln. Andersherum wäre die `CalendarLink` weg, bevor
 * der Flow sie lesen kann — und der Termin bliebe für immer in den
 * Kalendern stehen.
 */
export async function removeWaitlistShadow(
  svc: EventService,
  eventId: string,
  eventTitle: string,
): Promise<boolean> {
  const schatten = await findWaitlistShadow(svc, eventId, eventTitle);
  if (!schatten) return true; // nichts da = nichts zu tun
  try {
    const r = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${schatten.id})?$select=CalendarLink`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } },
    );
    let link = '';
    if (r.ok) { const d = await r.json(); link = d.CalendarLink || ''; }
    if (link) await svc.queueOutlookDeleteEvent(schatten.id, schatten.title, link);
    const del = await svc._delete(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${schatten.id})`);
    return del.ok || del.status === 204;
  } catch (e) {
    console.warn('[DEX] removeWaitlistShadow:', e);
    return false;
  }
}
