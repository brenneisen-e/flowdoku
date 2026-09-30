/**
 * v30.66 — Modularisierung Stufe 2: Thema „Event-CRUD" — die Zeilen in
 * DEX_Events lesen, anlegen (samt Subsite und Teilnehmerliste), ändern und
 * löschen.
 *
 * Zwei Dinge, die hier hängen (siehe CLAUDE.md): `EndDate` fällt zentral auf
 * `StartDate` zurück, weil der Outlook-Flow bei leerem Ende abbricht — und
 * der Flow triggert nur auf NEUE Listeneinträge, ein MERGE stösst ihn nie an.
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import { normalizeMadeWithLink } from '../EmailTemplates';
import { buildOutlookLocation } from '../../utils/eventFormat';
import { dlog } from '../../utils/debugLog';
import { withThrottleRetry } from '../../utils/spThrottle';
import type { EventService, CustomField, SPEvent } from '../EventService';
import { REG_LIST_NAME } from '../EventService';
import { kompakteZeile } from './deleteSafety'; // v31.62

// v30.66: war `private static readonly` an der Klasse — die Spaltenliste wird
// nur von den Lese-Methoden dieses Themas gebraucht.
// v32.45: OutlookIsOnlineMeeting fehlte hier seit v30.26 — gemappt wurde es
// (eventMapping), gelesen nie. Folgen: „Meine Events" kannte keine von DEX
// erzeugte Teams-Besprechung, und der Assistent öffnete jedes solche Event im
// Modus „kein Online-Meeting" — beim Speichern schrieb er den Schalter auf
// false zurück. Die Spalte existiert sicher: jedes Wizard-Speichern schreibt sie.
const EVENT_SELECT = 'Id,Title,EventStatus,EventNumber,Description,Location,LocationAddress,LocationFilter,Audience,AudienceResolvedEmails,FilterMode,StartDate,EndDate,RegistrationDeadline,LastDeregisterDate,MaxParticipants,CurrentParticipants,WaitlistEnabled,MandatoryRegistration,EventImageUrl,EmailImageBase64,Organizer,OrganizerEmail,ContactName,ContactEmail,ContactOrganizerEmail,ContactInfo,OutlookEventId,CalendarLink,OutlookBody,OutlookSubject,OutlookStart,OutlookEnd,OutlookLocation,AllDay,ShowAsFree,OutlookIsOnlineMeeting,SkipOrganizerInvite,EmailLanguage,RegistrationLanguage,EmailTemplateOverrides,DisableEmails,DisableRegistrationEmail,DisableCancellationEmail,AutoDeregisterOnDecline,InactiveHandling,DisableOutlook,OutlookDirty,AutoSendQRCode,ActiveFrom,NotifyOrgRegisterMode,NotifyOrgRegisterFromDate,NotifyOrgCancelMode,ExcludedUsers,IsFictive,DurchstarterCapacity,FunstarterCapacity,SplitLabelA,SplitLabelB,SplitDescA,SplitDescB,SplitHelpText,SplitSectionTitle,SplitSharedWaitlist,AllowAttendeeUpload,AttendeeUploadHint,AttendeeUploadLabel,AskSalutation,ConfirmDialogEnabled,ConfirmDialogMode,ConfirmDialogText,SelfCheckInEnabled,SelfCheckInToken,SelfCheckInFrom,SelfCheckInTo,TeamRegistrationEnabled,TeamSize,AskTeamName,TeamPartialAllowed,TeamOpenSlotsVisible,TeamJoinRequiresApproval,BilingualFields,CustomFields,Agenda,Transfers,Documents,FunZone,QuizClusterSize,ParentEventId,RegistrationListName,SubsiteUrl,Modified,Created';

/**
 * v32.0.6: Die Start-Abfrage OHNE die zwei Riesenspalten. Messung vom
 * 28.09.2026 (v32.0.5): 36,7 MB für 100 Events, davon OutlookBody 23,5 MB
 * (eingebackene Bilder im Termin-Text) und EmailImageBase64 4,5 MB — 18 bis
 * 20 s Boot für JEDE Rolle. EmailImageBase64 übernimmt die App nie in den
 * Speicher (das Mail-Logo kommt aus EmailTemplateOverrides._eventLogo);
 * OutlookBody brauchen nur der Assistent und „In Programmpunkte überführen"
 * — der EventContext lädt ihn nach dem Start im Hintergrund nach
 * (`getOutlookBodies`) und sperrt diese beiden Stellen, bis er da ist.
 * Alle ANDEREN Aufrufer von getEvents() lesen weiter die volle Liste.
 */
const EVENT_SELECT_SLIM = EVENT_SELECT.split(',').filter(c => c !== 'OutlookBody' && c !== 'EmailImageBase64').join(',');

/**
 * v32.0.10: Gibt es die Spalte auf DEX_Events? Einmal je Sitzung gefragt.
 * Eine Spalte, die noch nicht angelegt ist (das Schema zieht ein Admin beim
 * ersten Start einer neuen Version nach), darf in keinem Schreibvorgang
 * stehen — SharePoint lehnt den GANZEN Save mit 400 ab.
 */
const spaltenCache: Record<string, Promise<boolean>> = {};
export function hatEventsSpalte(svc: EventService, name: string): Promise<boolean> {
  const key = `${svc.siteUrl}|${name}`;
  if (!spaltenCache[key]) {
    spaltenCache[key] = (async () => {
      try {
        const r = await svc._sp.get(
          `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/fields/getbyinternalnameortitle('${name}')?$select=InternalName`,
          SPHttpClient.configurations.v1
        );
        return r.ok;
      } catch { return false; }
    })();
    // Ein Nein nicht für immer merken — nach dem Schema-Nachzug gilt es nicht mehr.
    spaltenCache[key].then(ok => { if (!ok) setTimeout(() => { delete spaltenCache[key]; }, 60000); }).catch(() => { delete spaltenCache[key]; });
  }
  return spaltenCache[key];
}

/**
 * v32.0.10: Die zwei Bilder aus EmailTemplateOverrides in ihre Spalten
 * verlagern — am EINEN Engpass aller Schreibwege (createEvent/updateEvent).
 *
 * Messung 28.09.2026: Von 8,7 MB Start-Abfrage waren 4,5 MB `_eventLogo`
 * (dasselbe Bild steht ohnehin in EmailImageBase64, das die Flows lesen) und
 * 3,4 MB `_outlookLogo`. Beide braucht der Start nicht. Hier wandern sie in
 * EmailImageBase64 bzw. OutlookLogoBase64; die Overrides tragen den Merker
 * `_logosAusgelagert` (1 = Mail-Logo, 2 = beide). Der EventContext setzt die
 * Bilder beim Hintergrund-Nachlauf wieder ins JSON im Speicher ein — der
 * übrige Code liest weiter `o._eventLogo` und merkt davon nichts.
 *
 * Regeln, damit nie ein Bild verloren geht:
 *  - Fehlt ein Bild im JSON, wird die Spalte NICHT geleert — außer bei einem
 *    vollständigen Kommunikations-Write (Wizard: EmailImageBase64 steht mit
 *    im Update). Nur dort heißt „fehlt" auch „entfernt".
 *  - OutlookLogoBase64 nur, wenn die Spalte schon existiert; sonst bleibt
 *    `_outlookLogo` wie bisher im JSON (Merker 1).
 *  - Schreiber, die die Overrides roh aus SharePoint lesen und zurückschreiben
 *    (patchEventOverridesKey/Value), sehen den Merker ohne Bilder und lassen
 *    die Spalten in Ruhe.
 */
export async function logosAuslagern(svc: EventService, felder: Record<string, unknown>): Promise<void> {
  const raw = felder['EmailTemplateOverrides'];
  if (typeof raw !== 'string' || !raw) return;
  let o: Record<string, unknown>;
  try { o = JSON.parse(raw) || {}; } catch { return; }
  if (!o || typeof o !== 'object' || Array.isArray(o)) return;
  const vollerKommWrite = 'EmailImageBase64' in felder;
  const hatBild = typeof o._eventLogo === 'string' || typeof o._outlookLogo === 'string';
  if (!hatBild && !vollerKommWrite && !o._logosAusgelagert) return;
  const outlookSpalte = await hatEventsSpalte(svc, 'OutlookLogoBase64');
  if (typeof o._eventLogo === 'string') {
    if (!vollerKommWrite && o._eventLogo) felder['EmailImageBase64'] = o._eventLogo;
    delete o._eventLogo;
  }
  if (outlookSpalte) {
    if (typeof o._outlookLogo === 'string') {
      felder['OutlookLogoBase64'] = o._outlookLogo;
      delete o._outlookLogo;
    } else if (vollerKommWrite) {
      felder['OutlookLogoBase64'] = '';
    }
  }
  const merker = outlookSpalte ? 2 : 1;
  // Ein bereits höherer Merker bleibt (Spalte kurz nicht erkannt ≠ zurückbauen).
  o._logosAusgelagert = Math.max(merker, Number(o._logosAusgelagert) || 0);
  felder['EmailTemplateOverrides'] = JSON.stringify(o);
}

/**
 * Seed-Events anlegen falls sie nicht existieren (einmalig beim ersten Start).
 */
export async function seedEvents(svc: EventService): Promise<void> {
  try {
    // Prüfen ob "Assistenz Meeting 2026" schon existiert
    const check = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=Title eq 'Assistenz Meeting 2026'&$top=1&$select=Id`,
      SPHttpClient.configurations.v1
    );
    if (check.ok) {
      const data = await check.json();
      const items = data.value || data.d?.results || [];
      if (items.length > 0) return; // Existiert bereits
    }

    // Event anlegen
    await svc.createEvent({
      title: 'Assistenz Meeting 2026',
      type: 'Other',
      status: 'Active',
      description: 'Assistenz Meeting Mai 2026 - Frankfurt am Main',
      location: 'Frankfurt am Main',
      locationFilter: '',
      audience: 'All',
      filterMode: 'OR',
      startDate: '2026-05-07T11:00:00.000Z',
      endDate: '2026-05-08T15:00:00.000Z',
      registrationDeadline: '2026-04-09T00:00:00.000Z',
      lastDeregisterDate: '',
      maxParticipants: 130,
      waitlistEnabled: true,
      eventImageUrl: '',
      organizer: 'Maerzluft, Petra; Schwartz, Eva',
      organizerEmail: 'pmaerzluft@deloitte.de',
      outlookEventId: '',
      outlookBody: '',
      emailLanguage: 'EN',
      emailTemplateOverrides: '',
      customFields: [
        { id: 'travel', label: 'You will travel with?', type: 'select', required: false, visible: true, options: ['Train', 'Car', 'Public Transport'] },
        { id: 'deutschlandticket', label: 'Do you own a Deutschlandticket?', type: 'select', required: false, visible: true, options: ['Yes', 'No'] },
        { id: 'expenses', label: 'Please insert the total amount of your travel expenses!', type: 'text', required: false, visible: true },
      ],
      agenda: '[]',
      transfers: '[]',
      documents: '[]',
    });
  } catch { /* Seed fehlgeschlagen - nicht kritisch */ }
}

/**
 * Alle Events laden
 *
 * @param onHttpError v31.6: Meldet, dass DEX_Events NICHT gelesen werden
 *   konnte — `status` ist der HTTP-Status, `0` bei Netz-/Parse-Fehler. Ohne
 *   diesen Rückruf ist „keine Rechte auf der Liste" von „zurzeit keine
 *   Events" nicht zu unterscheiden: Beides kam als `[]` zurück, und wer aus
 *   einer Member Firm nur ein persönliches Leserecht auf der Site hat (statt
 *   der Besucher-Gruppe), sah eine leere Übersicht und meldete dem Organizer
 *   „ich sehe das Event nicht" — der dann am Event suchte statt an den
 *   Rechten. Dieselbe Falle wie bei `getAllRegistrations` (CLAUDE.md: „Ein
 *   Lesefehler ist keine Null").
 */
/**
 * v32.20: Folgt dem nextLink einer Listenabfrage. Bis v32.19 las die App
 * DEX_Events mit `$top=100` OHNE Weiterblättern — seit die Liste 100 Zeilen
 * überschritten hat, fehlten alle Events mit dem frühesten Startdatum
 * (Befund 29.09.2026: jedes Log zeigte exakt „100 Events“). Ein Fehler auf
 * einer Folgeseite ist ein Lesefehler, keine leere Menge: Status melden.
 */
async function alleSeiten(
  svc: EventService,
  ersteUrl: string,
  onHttpError?: (_status: number) => void,
): Promise<{ rows: unknown[]; bytes: number; seiten: number; ok: boolean }> {
  let url: string | null = ersteUrl;
  const rows: unknown[] = [];
  let bytes = 0; let seiten = 0;
  // Sicherung gegen Endlosschleifen: 50 Seiten à 100 = 5000 Zeilen.
  while (url && seiten < 50) {
    const response = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) { if (onHttpError) onHttpError(response.status); return { rows, bytes, seiten, ok: false }; }
    const raw = await response.text();
    bytes += raw.length; seiten++;
    const data = JSON.parse(raw);
    for (const r of (data.value || [])) rows.push(r);
    url = data['odata.nextLink'] || data['@odata.nextLink'] || (data.d && data.d.__next) || null;
  }
  return { rows, bytes, seiten, ok: true };
}

export async function getEvents(svc: EventService, onHttpError?: (_status: number) => void, slim?: boolean): Promise<SPEvent[]> {
  try {
    const t0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : 0;
    const res = await alleSeiten(svc,
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${slim ? EVENT_SELECT_SLIM : EVENT_SELECT}&$orderby=StartDate desc&$top=100`,
      onHttpError);
    if (!res.ok && res.rows.length === 0) return [];
    // v29.51 (Messpunkt): Das ist die EINZIGE blockierende Datenabfrage des
    // Starts — und EVENT_SELECT holt 79 Spalten, darunter EmailImageBase64
    // und EmailTemplateOverrides mit eingebetteten Bildern. Ob das ein paar
    // Kilobyte oder mehrere Megabyte sind, entscheidet über den nächsten
    // Optimierungsschritt; bisher wurde darüber geraten. `.text()` +
    // JSON.parse ist genau das, was `.json()` intern auch tut — der Umweg
    // kostet nichts und liefert die exakte Byte-Zahl.
    const rows = res.rows as SPEvent[];
    const raw = { length: res.bytes };
    const gesamtMs = t0 ? Math.round(performance.now() - t0) : -1;
    dlog('perf',
      `[DEX][perf][getEvents] ${rows.length} Events · ${Math.round(raw.length / 1024)} KB JSON · ${res.seiten} Seite(n) · ${gesamtMs} ms${res.ok ? '' : ' · ABGEBROCHEN'}`
    );
    // v32.0.4: Welche Spalte trägt das Gewicht? Befund 28.09.2026: 36,7 MB für
    // 100 Events, 18 s Boot — für JEDE Rolle, denn alle lesen dieselbe Abfrage.
    // Bevor Spalten aus dem Boot fliegen, messen statt raten: je Spalte die
    // Summe und die größte Zeile (mit Titel), die zehn schwersten.
    try {
      if (raw.length > 2 * 1024 * 1024) {
        const sum: Record<string, number> = {};
        const max: Record<string, { n: number; title: string }> = {};
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        for (const r of rows as any[]) {
          for (const k of Object.keys(r)) {
            const v = r[k];
            if (v === null || v === undefined) continue;
            const n = typeof v === 'string' ? v.length : (JSON.stringify(v) || '').length;
            sum[k] = (sum[k] || 0) + n;
            if (!max[k] || n > max[k].n) max[k] = { n, title: String(r.Title || r.Id || '') };
          }
        }
        const top = Object.keys(sum).sort((a, b) => sum[b] - sum[a]).slice(0, 10)
          .map(k => `${k}: ${Math.round(sum[k] / 1024)} KB (größte: ${Math.round(max[k].n / 1024)} KB „${max[k].title}")`);
        dlog('perf', `[DEX][perf][getEvents] schwerste Spalten:\n  ${top.join('\n  ')}`);
        // v32.0.8 (Stufe 1 „Start-Beschleunigung"): Was steckt IN
        // EmailTemplateOverrides? Je Schlüssel die Summe, und für die Bilder,
        // wie viele Zeilen dasselbe Bild tragen — Verdacht: Termine kopieren
        // das Logo der Klammer (persistSubEvents), eine Serie mit 20 Terminen
        // trägt es 21-mal. Gemessen wird, nichts geändert.
        const keySum: Record<string, number> = {};
        const keyRows: Record<string, number> = {};
        const bildZeilen: Record<string, Record<string, number>> = { _eventLogo: {}, _outlookLogo: {} };
        let kindKopien = 0; let kindKopienKB = 0;
        const logoVon: Record<string, string> = {};
        const fp = (v: string): string => `${v.length}:${v.slice(0, 48)}:${v.slice(-48)}`;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        for (const r of rows as any[]) {
          let o: Record<string, unknown> = {};
          try { o = JSON.parse(r.EmailTemplateOverrides || '{}') || {}; } catch { continue; }
          for (const k of Object.keys(o)) {
            const v = o[k];
            const n = typeof v === 'string' ? v.length : (JSON.stringify(v) || '').length;
            keySum[k] = (keySum[k] || 0) + n;
            keyRows[k] = (keyRows[k] || 0) + 1;
            if ((k === '_eventLogo' || k === '_outlookLogo') && typeof v === 'string' && v.length > 0) {
              const f = fp(v);
              bildZeilen[k][f] = (bildZeilen[k][f] || 0) + 1;
              if (k === '_eventLogo') logoVon[String(r.Id)] = f;
            }
          }
        }
        // Termine, deren Mail-Logo identisch mit dem ihrer Klammer ist.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        for (const r of rows as any[]) {
          const pid = String(r.ParentEventId || '');
          const f = logoVon[String(r.Id)];
          if (pid && f && logoVon[pid] === f) { kindKopien++; kindKopienKB += Number(f.split(':')[0]) / 1024; }
        }
        const keyTop = Object.keys(keySum).sort((a, b) => keySum[b] - keySum[a]).slice(0, 8)
          .map(k => `${k}: ${Math.round(keySum[k] / 1024)} KB in ${keyRows[k]} Zeilen`);
        const bildInfo = ['_eventLogo', '_outlookLogo'].map(k => {
          const z = bildZeilen[k]; const verschieden = Object.keys(z).length;
          const zeilen = Object.keys(z).reduce((a, f) => a + z[f], 0);
          return `${k}: ${zeilen} Zeilen, ${verschieden} verschiedene Bilder`;
        });
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const beideGleich = (rows as any[]).filter(r => {
          try { const o = JSON.parse(r.EmailTemplateOverrides || '{}') || {}; return typeof o._eventLogo === 'string' && o._eventLogo.length > 0 && o._eventLogo === o._outlookLogo; } catch { return false; }
        }).length;
        dlog('perf', `[DEX][perf][getEvents] EmailTemplateOverrides je Schlüssel:\n  ${keyTop.join('\n  ')}\n  ${bildInfo.join('\n  ')}\n  Termine mit Kopie des Klammer-Logos: ${kindKopien} (${Math.round(kindKopienKB)} KB)\n  Zeilen mit _eventLogo === _outlookLogo: ${beideGleich}`);
      }
    } catch { /* Messung ist Beiwerk */ }
    return rows;
  } catch {
    if (onHttpError) onHttpError(0);
    return [];
  }
}

/**
 * v32.0.6: Die Outlook-Texte derselben Zeilen wie getEvents(slim) — für den
 * Hintergrund-Nachlauf nach dem Start. `null` = nicht lesbar (dann bleiben
 * die Sperren stehen; ein Lesefehler ist kein leerer Text).
 */
export async function getOutlookBodies(svc: EventService, ids?: string[]): Promise<Record<string, { body: string; modified: string; mailLogo: string; outlookLogo: string }> | null> {
  try {
    // Mit ids: nur diese Zeilen (nach einem Speichern), sonst dieselben 100 wie der Boot.
    const nummern = (ids || []).filter(id => /^\d+$/.test(id));
    const filter = nummern.length ? `&$filter=${encodeURIComponent(nummern.map(id => `Id eq ${id}`).join(' or '))}` : '';
    // v32.0.10: dazu die ausgelagerten Bilder (s. logosAuslagern) — die
    // Outlook-Spalte nur, wenn es sie gibt (sonst lehnt SharePoint ab).
    const outlookSpalte = await hatEventsSpalte(svc, 'OutlookLogoBase64');
    const sel = `Id,Modified,OutlookBody,EmailImageBase64${outlookSpalte ? ',OutlookLogoBase64' : ''}`;
    // v32.20: alle Seiten — sonst blieben Events jenseits der ersten 100
    // für immer „Outlook-Text ausstehend", und der Assistent wartete darauf.
    const res = await alleSeiten(svc,
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${sel}${filter}&$orderby=StartDate desc&$top=100`);
    if (!res.ok) return null;
    const out: Record<string, { body: string; modified: string; mailLogo: string; outlookLogo: string }> = {};
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    for (const r of res.rows as any[]) {
      out[String(r.Id)] = { body: r.OutlookBody || '', modified: r.Modified || '', mailLogo: r.EmailImageBase64 || '', outlookLogo: r.OutlookLogoBase64 || '' };
    }
    return out;
  } catch {
    return null;
  }
}

/**
 * Einzelnes Event laden
 */
export async function getEvent(svc: EventService, eventId: number): Promise<SPEvent | null> {
  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=${EVENT_SELECT}`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return null;
    return await response.json();
  } catch {
    return null;
  }
}

/**
 * v18.33: Event anhand des Self-Check-in-Tokens finden (für den statischen
 * Check-in-Link ?action=selfcheckin&token=…). Liefert das erste Event mit
 * passendem Token. Alle eingeloggten User dürfen DEX_Events lesen.
 */
/**
 * v30.67: Alle Kind-Ids eines Klammer-Events direkt aus DEX_Events. getEvents
 * laedt nur $top=100 nach StartDate desc und laesst Zeilen mit Mapping-Fehler
 * aus — ein Loeschlauf, der die Kinder aus dem Client-State nimmt, kann Termine
 * uebersehen, und die blieben dann mit ParentEventId auf ein geloeschtes Item
 * verwaist zurueck. null = nicht lesbar (dann darf NICHT geloescht werden).
 */
export async function getChildEventIds(svc: EventService, parentId: number): Promise<number[] | null> {
  if (!parentId) return null;
  try {
    const ids: number[] = [];
    let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=Id&$filter=ParentEventId eq '${parentId}'&$top=5000`;
    while (url) {
      const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return null;
      const data = await resp.json();
      (data.value || data.d?.results || []).forEach((r: { Id: number }) => { if (typeof r.Id === 'number') ids.push(r.Id); });
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    }
    return ids;
  } catch { return null; }
}

export async function getEventBySelfCheckInToken(svc: EventService, token: string): Promise<SPEvent | null> {
  try {
    const safe = token.replace(/'/g, "''");
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EVENT_SELECT}&$filter=SelfCheckInToken eq '${safe}'&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return null;
    const data = await response.json();
    const arr = data.value || (data.d && data.d.results) || [];
    return arr.length > 0 ? arr[0] : null;
  } catch {
    return null;
  }
}

/**
 * v18.33: Event anhand der Event-Nummer finden (für den rotierenden Live-QR
 * ?action=selfcheckin&event=<Nr>&code=…&t=…).
 */
export async function getEventByEventNumber(svc: EventService, eventNumber: number): Promise<SPEvent | null> {
  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=${EVENT_SELECT}&$filter=EventNumber eq ${eventNumber}&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return null;
    const data = await response.json();
    const arr = data.value || (data.d && data.d.results) || [];
    return arr.length > 0 ? arr[0] : null;
  } catch {
    return null;
  }
}

/**
 * Neues Event erstellen + Subsite mit Teilnehmerliste anlegen
 */
export async function createEvent(svc: EventService, event: {
  title: string;
  status: string;
  type: string;
  description: string;
  location: string;
  locationAddress?: string; // JSON-String: { street, houseNo, zip, city }
  outlookSubject?: string; // v18.42: Betreff des Outlook-Termins (leer = Titel)
  outlookStart?: string; // v18.44: abweichende Start-Zeit (ISO, leer = Event-Start)
  outlookEnd?: string;   // v18.44: abweichende End-Zeit (ISO, leer = Event-Ende)
  outlookLocation?: string; // v18.40: manueller Outlook-Ort (leer = Auto aus Ort + Adresse)
  allDay?: boolean; // v29.52: ganztägiger Termin (Flow setzt daraus isAllDay)
  showAsFree?: boolean; // v29.54: Termin als „Frei" anzeigen (Flow: showAs)
  skipOrganizerInvite?: boolean; // v29.55: Organizer nicht einladen (Flow: requiredAttendees)
  outlookIsOnlineMeeting?: boolean; // v30.26: Termin als Teams-Besprechung anlegen (Flow: isOnlineMeeting)
  locationFilter: string;
  audience: string;
  /** v16.4: Vor-aufgelöste E-Mails der Audience-DLs, ';'-separiert, lowercase. */
  audienceResolvedEmails?: string;
  filterMode: string;
  startDate: string;
  endDate: string;
  registrationDeadline: string;
  lastDeregisterDate: string;
  /** v29.19: Auto-Aktivierungszeitpunkt (UTC-ISO). Der Wizard bot das Feld
   *  auch beim ANLEGEN an, persistiert wurde es aber nur im Edit-Pfad —
   *  ein als Entwurf angelegtes Event mit „Aktiv ab" ging nie von allein
   *  live. */
  activeFrom?: string;
  maxParticipants: number;
  waitlistEnabled: boolean;
  mandatoryRegistration?: boolean; // v24.64: Pflicht-Sub-Event

  eventImageUrl: string;
  organizer: string;
  organizerEmail: string;
  /** v10.16: optionaler Ansprechpartner (Anzeige-Feld). */
  contactName?: string;
  contactEmail?: string;
  contactOrganizerEmail?: string;
  contactInfo?: string;
  outlookEventId: string;
  outlookBody: string;
  agenda?: string; // JSON-Array mit Agenda-Einträgen
  transfers?: string; // JSON-Array mit Transferzeiten
  documents?: string; // JSON-Array mit Dokumenten
  funZone?: string; // JSON-Array mit Quiz-Fragen
  quizClusterSize?: number; // 1..4 - Fragen pro Quiz-Ansicht
  /** Seit v6.4: wenn gesetzt, wird dieses Event als Sub-Event angelegt und zeigt auf das angegebene Parent-Event. */
  parentEventId?: string;
  emailLanguage?: string;
  registrationLanguage?: 'de' | 'en';
  emailTemplateOverrides?: string;
  disableEmails?: boolean;
  disableRegistrationEmail?: boolean;
  disableCancellationEmail?: boolean;
  autoDeregisterOnDecline?: boolean;
  inactiveHandling?: string;
  disableOutlook?: boolean;
  notifyOrgRegisterMode?: 'never' | 'always' | 'fromDate';
  notifyOrgRegisterFromDate?: string;
  notifyOrgCancelMode?: 'never' | 'always' | 'afterDeadline';
  excludedUsers?: string[];
  isFictive?: boolean;
  durchstarterCapacity?: number;
  funstarterCapacity?: number;
  splitLabelA?: string;
  splitLabelB?: string;
  splitDescA?: string;
  splitDescB?: string;
  splitHelpText?: string;
  splitSectionTitle?: string;
  splitSharedWaitlist?: boolean;
  allowAttendeeUpload?: boolean;
  attendeeUploadHint?: string;
  attendeeUploadLabel?: string;
  /** v11.80: Anrede im Registrierungsformular abfragen (Default false). */
  askSalutation?: boolean;
  /** v18.75: Sicherheitshinweis vor dem Absenden der Anmeldung. */
  confirmDialogEnabled?: boolean;
  confirmDialogMode?: string; // 'summary' | 'freetext'
  confirmDialogText?: string;
  /** v18.33: Self-Check-in per QR-Code erlauben (Default false). */
  selfCheckInEnabled?: boolean;
  /** v18.33: Geheimer Token (statischer Link + HMAC-Schlüssel rotierender QR). */
  selfCheckInToken?: string;
  /** v18.33: optionaler Start des Check-in-Fensters (ISO). */
  selfCheckInFrom?: string;
  /** v18.33: optionales Ende des Check-in-Fensters (ISO). */
  selfCheckInTo?: string;
  /** v11.80: Team-Anmeldung erlauben (Default false). */
  teamRegistrationEnabled?: boolean;
  /** v11.80: Maximale Teamgröße (0 = nicht gesetzt). */
  teamSize?: number;
  /** v11.80: Team-Name abfragen (Default false). */
  askTeamName?: boolean;
  /** v11.81: Auch Teil-Teams zulassen (Default false = nur komplette Teams). */
  teamPartialAllowed?: boolean;
  /** v11.81: Offene Slots öffentlich für Beitritt sichtbar (Default false). */
  teamOpenSlotsVisible?: boolean;
  /** v11.81: Beitritt erfordert Bestätigung durch Team-Kapitän (Default false). */
  teamJoinRequiresApproval?: boolean;
  /** v17.20: Custom-Fields zweisprachig anbieten (DE+EN). */
  bilingualFields?: boolean;
  customFields: CustomField[];
  /** v11.69: Wenn `existingSubsiteUrl` UND `existingRegistrationListName`
   *  gesetzt sind, wird KEINE neue Subsite und KEINE neue Teilnehmer-
   *  liste angelegt — stattdessen werden die mitgegebenen Werte direkt in
   *  das neue DEX_Events-Item geschrieben. Hintergrund: Outlook-Termin
   *  nachträglich aktivieren ohne Verlust bestehender Anmeldungen — das
   *  Sub-Event wird mit `deleteEventItemOnly()` aus DEX_Events entfernt
   *  und hier neu angelegt, wobei die alte Subsite + Teilnehmerliste
   *  unangetastet bleiben und an die neue Event-Zeile angehängt werden.
   *  Damit triggert der `DEX_CreateOutlookEvent`-Flow (GetOnNewItems) auf
   *  dem neuen Item und legt den Outlook-Termin an. */
  existingSubsiteUrl?: string;
  existingRegistrationListName?: string;
  /** v11.87: Optionaler Progress-Callback. Wird zu Beginn jeder Teil-
   *  Operation aufgerufen — die UI kann darauf den Fortschrittsbalken
   *  und die Unter-Caption sichtbar bewegen, statt minutenlang auf
   *  „Event wird vorbereitet..." stehen zu bleiben. Stages decken
   *  die langsamen SP-Operationen ab (Subsite-Create, Listen-Create,
   *  Permissions, Counter, Views). */
  onProgress?: (stage:
    | 'start'
    | 'subsite-creating'
    | 'subsite-done'
    | 'permissions'
    | 'list-creating'
    | 'list-done'
    | 'item-insert'
    | 'done'
  ) => void;
}): Promise<number | null> {
  const reportProgress = (stage:
    | 'start'
    | 'subsite-creating'
    | 'subsite-done'
    | 'permissions'
    | 'list-creating'
    | 'list-done'
    | 'item-insert'
    | 'done'
  ): void => {
    try { event.onProgress?.(stage); } catch { /* */ }
  };
  try {
    reportProgress('start');
    // 0. Nächste EventNumber ermitteln
    // v30.67: Ein Lesefehler ist KEIN leeres Ergebnis. Bisher diente
    // `nextEventNumber = 1` gleichzeitig als Startwert für „Liste ist leer"
    // und als stiller Fallback für „Abfrage gescheitert" (429/500/Netz) — ab
    // da trugen zwei Events dieselbe Nummer, und DEX_Participants, „Meine
    // Events", der QR-Code und deleteEvent schlüsseln alle über diese Nummer.
    // Deshalb: bei Fehler abbrechen, BEVOR irgendetwas angelegt wird (der
    // Wizard zeigt Errors aus createEvent an). Unmittelbar vor dem Insert wird
    // die Nummer noch einmal nachgelesen und nach dem Insert auf Eindeutigkeit
    // geprüft — zwischen erstem Lesen und Schreiben liegt die Subsite-Anlage
    // (oft > 30 s), zwei parallel anlegende Organizer zogen dieselbe Nummer.
    const readNextEventNumber = async (): Promise<number> => {
      let enResp;
      try {
        // v30.67 (Review): Dieser GET ist der erste Request eines SCHREIB-
        // Vorgangs — anders als die Boot-GETs (v29.50) darf er auf eine
        // Drosselung warten. Ohne Retry wurde ein einzelnes 429 zum harten
        // Abbruch, und im Recreate-Pfad der Sub-Events war die alte Zeile
        // dann bereits gelöscht.
        enResp = await withThrottleRetry(() => svc._sp.get(
          `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=EventNumber&$orderby=EventNumber desc&$top=1`,
          SPHttpClient.configurations.v1
        ), 'EventNumber');
      } catch (err) {
        throw new Error(`Event-Nummer konnte nicht ermittelt werden (${err instanceof Error ? err.message : 'Netzwerkfehler'}) — bitte erneut versuchen.`);
      }
      if (!enResp.ok) throw new Error(`Event-Nummer konnte nicht ermittelt werden (HTTP ${enResp.status}) — bitte erneut versuchen.`);
      const enData = await enResp.json();
      const items = enData.value || enData.d?.results || [];
      const top = items.length > 0 ? Number(items[0].EventNumber) : 0;
      return (Number.isFinite(top) && top > 0) ? top + 1 : 1;
    };
    let nextEventNumber = await readNextEventNumber();

    // v11.69: Reuse-Pfad — wenn `existingSubsiteUrl` UND
    // `existingRegistrationListName` mitgegeben wurden, überspringen wir
    // 1) Subsite-Anlegen, 2) Subsite-Permissions, 3) Teilnehmerliste
    // anlegen. Die mitgegebene Subsite bleibt unangetastet inkl. aller
    // Teilnehmer-Anmeldungen. Custom-Fields werden ohne spInternalName-
    // Anreicherung übernommen — die Felder existieren bereits auf der
    // alten Teilnehmerliste mit den korrekten Internal-Names.
    const reuseSubsite = !!(event.existingSubsiteUrl && event.existingRegistrationListName);
    let subsiteUrl: string;
    let enrichedCustomFields: CustomField[];
    const coOrgEmailsForPerm: string[] = (() => {
      try {
        const o = JSON.parse(event.emailTemplateOverrides || '{}');
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const list = (o as any)._coOrganizers;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        if (Array.isArray(list)) return list.map((x: any) => String(x?.email || '')).filter(Boolean);
      } catch { /* */ }
      return [];
    })();
    const allOrgEmails = [event.organizerEmail || '', ...coOrgEmailsForPerm].filter(Boolean).join(';');
    const regListName = reuseSubsite ? (event.existingRegistrationListName as string) : REG_LIST_NAME;

    if (reuseSubsite) {
      // v11.69: bestehende Subsite + Teilnehmerliste wiederverwenden.
      subsiteUrl = event.existingSubsiteUrl as string;
      // Custom-Fields unverändert übernehmen — die Liste existiert
      // bereits, kein neues Schema nötig.
      enrichedCustomFields = event.customFields.map(cf => ({ ...cf }));
    } else {
      // 1. Subsite für das Event erstellen
      reportProgress('subsite-creating');
      const createdSubsite = await svc.createEventSubsite(event.title, event.description);
      if (!createdSubsite) {
        // v30.84: den Grund nennen. 403 = fehlender Vollzugriff auf der Site
        // (das Recht wird bei der Rollen-Zuweisung gesetzt und kann an der
        // Drosselung gescheitert sein) → ein Admin behebt es über
        // Rollenverwaltung → „Rechte prüfen". 429/503 = Drosselung → warten.
        const st = svc._lastSubsiteCreateStatus;
        console.error('[DEX] Subsite konnte nicht erstellt werden, HTTP', st);
        if (st === 403 || st === 401) {
          throw new Error('Subsite konnte nicht erstellt werden: Dein Konto hat keinen Vollzugriff auf die DEX-Site (HTTP 403). Das Recht wird bei der Organizer-Zuweisung gesetzt und ist bei dir nicht angekommen. Bitte einen Admin, in der Rollenverwaltung „Rechte prüfen" auszuführen — danach die App neu laden und erneut speichern.');
        }
        if (st === 429 || st === 503) {
          throw new Error('Subsite konnte nicht erstellt werden: SharePoint drosselt gerade (HTTP ' + st + '). Bitte ein bis zwei Minuten warten und erneut speichern — der Entwurf bleibt erhalten.');
        }
        throw new Error('Subsite konnte nicht erstellt werden' + (st ? ' (HTTP ' + st + ')' : '') + '. Bitte erneut versuchen; bleibt es dabei, einen Admin bitten, in der Rollenverwaltung „Rechte prüfen" auszuführen.');
      }
      subsiteUrl = createdSubsite;
      reportProgress('subsite-done');

      // 2. Subsite-Berechtigungen: Members der Parent-Site auf der Subsite berechtigen.
      // v9.18: Co-Organizer-Emails aus emailTemplateOverrides._coOrganizers extrahieren
      // und mit dem Hauptorganizer zusammen Full Control erteilen.
      reportProgress('permissions');
      await svc.setSubsitePermissions(subsiteUrl, allOrgEmails);

      // 3. Teilnehmerliste auf der Subsite erstellen
      reportProgress('list-creating');
      const fieldMap: Record<string, string> = await svc.createRegistrationList(subsiteUrl, event.customFields, allOrgEmails);
      reportProgress('list-done');

      // Custom Fields mit SP InternalName anreichern
      enrichedCustomFields = event.customFields.map(cf => ({
        ...cf,
        spInternalName: fieldMap[cf.id] || '',
      }));
    }

    // 3. Event in DEX_Events eintragen
    const payload = {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
      'Title': event.title,
      'EventNumber': nextEventNumber,
      'EventStatus': event.status,
      'Description': event.description,
      'Location': event.location,
      'LocationAddress': event.locationAddress || '',
      // v18.42: Outlook-Betreff (leer = Flow fällt auf Titel zurück via coalesce).
      'OutlookSubject': (event.outlookSubject && event.outlookSubject.trim()) ? event.outlookSubject.trim() : '',
      // v18.44: abweichendes Outlook-Datum (leer = Flow nutzt StartDate/EndDate).
      'OutlookStart': event.outlookStart || null,
      'OutlookEnd': event.outlookEnd || null,
      // v29.52: ganztägig. Start/Ende bleiben bewusst wie gesetzt (00:00/23:59)
      // — die Umrechnung auf die Ganztags-Grenzen macht der Flow. Solange der
      // Flow das Feld noch nicht liest, verhält sich alles wie bisher.
      'AllDay': !!event.allDay,
      'ShowAsFree': !!event.showAsFree, // v29.54
      'OutlookIsOnlineMeeting': !!event.outlookIsOnlineMeeting, // v30.26
      'SkipOrganizerInvite': !!event.skipOrganizerInvite, // v29.55
      // v18.34/v18.40: Outlook-Ort = manuelle Überschreibung, sonst
      // automatisch aus Veranstaltungsort + Adresse. Flow mappt OutlookLocation 1:1.
      // v26.54: hart auf 255 kappen (einzeilige Text-Spalte — s. updateEvent).
      'OutlookLocation': ((event.outlookLocation && event.outlookLocation.trim())
        ? event.outlookLocation.trim()
        : buildOutlookLocation(event.location, event.locationAddress)).slice(0, 255),
      'LocationFilter': event.locationFilter,
      'Audience': event.audience,
      'AudienceResolvedEmails': event.audienceResolvedEmails || '',
      'FilterMode': event.filterMode || 'OR',
      'StartDate': event.startDate || null,
      // v22.17/v28.66: EndDate darf NIE leer in DEX_Events landen — der
      // DEX_CreateOutlookEvent-Flow rechnet convertFromUtc(coalesce(
      // OutlookEnd, EndDate)); bei null stürzt „Create event (V4)" ab und es
      // entsteht kein Outlook-Termin. Die Aufrufer setzen den Fallback zwar
      // schon, hier wird er zentral erzwungen (letzte Instanz vor dem
      // Schreiben — s. auch updateEvent).
      'EndDate': event.endDate || event.startDate || null,
      'RegistrationDeadline': event.registrationDeadline || null,
      // v29.19: s. Interface — vorher nur im Edit-Pfad geschrieben.
      'ActiveFrom': event.activeFrom || null,
      'LastDeregisterDate': event.lastDeregisterDate || null,
      'MaxParticipants': event.maxParticipants,
      'WaitlistEnabled': event.waitlistEnabled,
      'MandatoryRegistration': !!event.mandatoryRegistration,
      'EventImageUrl': event.eventImageUrl,
      // Custom-Event-Logo aus emailTemplateOverrides._eventLogo extrahieren (falls
      // vorhanden) und als EmailImageBase64 persistieren — damit der Power-Automate-Flow
      // es als {{ORB_URL}} in Mail + Outlook-Termin einsetzt.
      'EmailImageBase64': (() => {
        try {
          const o = JSON.parse(event.emailTemplateOverrides || '{}');
          return (o && typeof o._eventLogo === 'string') ? o._eventLogo : '';
        } catch { return ''; }
      })(),
      'Organizer': event.organizer,
      'OrganizerEmail': event.organizerEmail,
      // v10.16: optionaler Ansprechpartner (Anzeige-Feld). Strings können
      // leer sein — leer = kein Ansprechpartner gepflegt.
      'ContactName': event.contactName || '',
      'ContactEmail': event.contactEmail || '',
      'ContactOrganizerEmail': event.contactOrganizerEmail || '',
      'ContactInfo': event.contactInfo || '',
      'OutlookEventId': event.outlookEventId,
      // outlookBody kommt bereits vollständig gewickelt + mit aufgelösten Variablen
      // aus EventCreationPage — hier nur durchreichen.
      // v29.42: auch im Termin-Text die Fußzeile auf die kanonische Adresse.
      'OutlookBody': normalizeMadeWithLink(event.outlookBody || ''),
      'EmailLanguage': event.emailLanguage || 'EN',
      'RegistrationLanguage': event.registrationLanguage || '',
      'EmailTemplateOverrides': event.emailTemplateOverrides || '',
      'DisableEmails': !!event.disableEmails,
      'DisableRegistrationEmail': !!event.disableRegistrationEmail,
      'DisableCancellationEmail': !!event.disableCancellationEmail,
      'AutoDeregisterOnDecline': !!event.autoDeregisterOnDecline,
      'InactiveHandling': event.inactiveHandling === 'autoderegister' ? 'autoderegister' : 'notify',
      'DisableOutlook': !!event.disableOutlook,
      'NotifyOrgRegisterMode': (() => {
        const m = event.notifyOrgRegisterMode || 'never';
        return m === 'always' ? 'Always' : m === 'fromDate' ? 'FromDate' : 'Never';
      })(),
      'NotifyOrgRegisterFromDate': event.notifyOrgRegisterFromDate || null,
      'NotifyOrgCancelMode': (() => {
        const m = event.notifyOrgCancelMode || 'never';
        return m === 'always' ? 'Always' : m === 'afterDeadline' ? 'AfterDeadline' : 'Never';
      })(),
      'ExcludedUsers': (event.excludedUsers || []).filter(Boolean).join(';'),
      'IsFictive': !!event.isFictive,
      'DurchstarterCapacity': typeof event.durchstarterCapacity === 'number' ? event.durchstarterCapacity : null,
      'FunstarterCapacity': typeof event.funstarterCapacity === 'number' ? event.funstarterCapacity : null,
      'SplitLabelA': event.splitLabelA || '',
      'SplitLabelB': event.splitLabelB || '',
      'SplitDescA': event.splitDescA || '',
      'SplitDescB': event.splitDescB || '',
      'SplitHelpText': event.splitHelpText || '',
      'SplitSectionTitle': event.splitSectionTitle || '',
      'SplitSharedWaitlist': !!event.splitSharedWaitlist,
      'AllowAttendeeUpload': !!event.allowAttendeeUpload,
      'AttendeeUploadHint': event.attendeeUploadHint || '',
      'AttendeeUploadLabel': event.attendeeUploadLabel || '',
      'AskSalutation': !!event.askSalutation,
      'ConfirmDialogEnabled': !!event.confirmDialogEnabled,
      'ConfirmDialogMode': event.confirmDialogMode || '',
      'ConfirmDialogText': event.confirmDialogText || '',
      'SelfCheckInEnabled': !!event.selfCheckInEnabled,
      'SelfCheckInToken': event.selfCheckInToken || '',
      'SelfCheckInFrom': event.selfCheckInFrom || null,
      'SelfCheckInTo': event.selfCheckInTo || null,
      'TeamRegistrationEnabled': !!event.teamRegistrationEnabled,
      'TeamSize': typeof event.teamSize === 'number' && event.teamSize > 0 ? event.teamSize : null,
      'AskTeamName': !!event.askTeamName,
      'TeamPartialAllowed': !!event.teamPartialAllowed,
      'TeamOpenSlotsVisible': !!event.teamOpenSlotsVisible,
      'TeamJoinRequiresApproval': !!event.teamJoinRequiresApproval,
      'BilingualFields': !!event.bilingualFields,
      'CustomFields': JSON.stringify(enrichedCustomFields),
      'Agenda': event.agenda || '[]',
      'Transfers': event.transfers || '[]',
      'Documents': event.documents || '[]',
      'FunZone': event.funZone || '[]',
      'QuizClusterSize': typeof event.quizClusterSize === 'number' ? event.quizClusterSize : null,
      'ParentEventId': event.parentEventId || '',
      'RegistrationListName': regListName,
      'RegistrationListUrl': `${subsiteUrl}/Lists/${regListName}/AllItems.aspx`,
      'SubsiteUrl': subsiteUrl,
    };

    reportProgress('item-insert');
    // v30.67: Nummer direkt vor dem Insert nachlesen (s. Kommentar oben).
    // Best-effort: Scheitert das Nachlesen, bleibt die zu Beginn geprüfte
    // Nummer — die Subsite steht schon, ein Abbruch hier hinterließe sie
    // verwaist. Im Reuse-Pfad (Recreate ohne Subsite-Anlage) liegen zwischen
    // beiden Lesevorgängen nur Millisekunden, dort entfällt der Request.
    if (!reuseSubsite) {
      try {
        const fresh = await readNextEventNumber();
        if (fresh > nextEventNumber) { nextEventNumber = fresh; payload.EventNumber = fresh; }
      } catch { /* zu Beginn geprüfte Nummer behalten */ }
    }
    // v32.0.10: Bilder aus den Overrides in ihre Spalten (s. logosAuslagern).
    await logosAuslagern(svc, payload as unknown as Record<string, unknown>);
    // v28.10: gleicher 2-MB-Schutz wie in updateEvent — zu große Payloads
    // (eingebettete Logos/Bilder) sauber abfangen statt kryptischem 400.
    if (JSON.stringify(payload).length > 1_900_000) {
      throw new Error('Die Event-Daten überschreiten das SharePoint-Limit von 2 MB. Ursache ist fast immer ein zu großes eingebettetes Bild (Mail-Logo, Outlook-Kopfbild oder ein Bild im Mail-/Termin-Text). Bitte das Bild entfernen oder neu (kleiner) hochladen.');
    }
    const response = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items`,
      payload
    );

    if (!response.ok) return null;
    const result = await response.json();
    const newItemId: number = result.d?.Id || result.Id;
    // v30.67: Eindeutigkeit nach dem Insert prüfen — analog zum Post-Insert-
    // Dedup der TeilnehmerID in registerForEvent. Haben zwei Organizer im
    // selben Fenster dieselbe Nummer gezogen, behält die ältere Zeile
    // (kleinere Id) die Nummer, unsere bekommt die nächste freie. Best-effort:
    // Ein Fehler hier darf das angelegte Event nicht mehr kippen, er wird
    // nur gemeldet.
    try {
      const dupResp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=Id&$filter=EventNumber eq ${nextEventNumber}&$top=10`,
        SPHttpClient.configurations.v1
      );
      if (dupResp.ok) {
        const dupData = await dupResp.json();
        const dupItems: Array<{ Id: number }> = dupData.value || dupData.d?.results || [];
        if (dupItems.length > 1 && newItemId !== Math.min(...dupItems.map(d => d.Id))) {
          const fresh = await readNextEventNumber();
          const fix = await svc._merge(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${newItemId})`, { 'EventNumber': fresh });
          if (fix.ok) console.warn(`[DEX] createEvent: EventNumber ${nextEventNumber} war doppelt vergeben — Item ${newItemId} auf ${fresh} korrigiert.`);
          else console.warn(`[DEX] createEvent: EventNumber ${nextEventNumber} doppelt vergeben, Korrektur fehlgeschlagen (HTTP ${fix.status}) — bitte im Admin Center prüfen.`);
        }
      }
    } catch (err) { console.warn('[DEX] createEvent: Eindeutigkeits-Prüfung der EventNumber fehlgeschlagen:', err); }
    reportProgress('done');
    return newItemId;
  } catch (err) {
    if (err instanceof Error) throw err;
    return null;
  }
}

/**
 * Admin-Cleanup beim App-Start: alle Events mit EventStatus='Active' und EndDate < jetzt
 * werden automatisch auf 'Completed' gesetzt. Liefert die Anzahl der aktualisierten Events.
 */
export async function markExpiredEventsAsCompleted(svc: EventService): Promise<number> {
  try {
    // SharePoint OData Filter: Active + EndDate < jetzt
    const nowIso = new Date().toISOString();
    const filter = `EventStatus eq 'Active' and EndDate lt datetime'${nowIso}'`;
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$filter=${encodeURIComponent(filter)}&$select=Id,Title,EndDate&$top=500`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return 0;
    const data = await resp.json();
    const items: Array<{ Id: number; Title: string }> = data.value || data.d?.results || [];
    if (items.length === 0) return 0;

    let updated = 0;
    for (const it of items) {
      try {
        const ok = await svc.updateEvent(it.Id, { 'EventStatus': 'Completed' });
        if (ok) updated += 1;
      } catch { /* einzelnes Update überspringen */ }
    }
    return updated;
  } catch (err) {
    console.warn('[DEX] markExpiredEventsAsCompleted failed:', err);
    return 0;
  }
}

/**
 * Event aktualisieren
 */
/**
 * v11.11: Versionsverlauf des Event-Items aus DEX_Events lesen, um
 * versehentlich gelöschte Custom-Fields (z.B. b2run_*-Felder nach
 * der zu aggressiven v11.9-Migration) wieder zurückzuholen.
 *
 * Liefert eine Liste der Versionen, jeweils mit dem geparsten
 * `CustomFields`-Array (sortiert: neueste zuerst). Werte ohne
 * CustomFields oder mit leerem Array fallen einfach mit raus, sind
 * aber nicht gefiltert — der Caller entscheidet, welche Version
 * relevant ist.
 */
export async function getEventCustomFieldsHistory(svc: EventService, eventId: number): Promise<Array<{
  versionLabel: string;
  modified: string;
  customFields: Array<Record<string, unknown>>;
}>> {
  try {
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/versions?$select=VersionLabel,Modified,CustomFields`;
    const response = await svc._sp.get(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' },
    });
    if (!response.ok) {
      console.warn('[DEX] getEventCustomFieldsHistory failed:', response.status);
      return [];
    }
    const data = await response.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const versions: any[] = data.value || [];
    return versions.map(v => {
      let parsed: Array<Record<string, unknown>> = [];
      try {
        const raw = (v.CustomFields || '').toString();
        if (raw.trim()) {
          const obj = JSON.parse(raw);
          if (Array.isArray(obj)) parsed = obj as Array<Record<string, unknown>>;
        }
      } catch { /* invalid JSON in old version → leer */ }
      return {
        versionLabel: String(v.VersionLabel || ''),
        modified: String(v.Modified || ''),
        customFields: parsed,
      };
    });
  } catch (err) {
    console.warn('[DEX] getEventCustomFieldsHistory error:', err);
    return [];
  }
}

/**
 * v32.5: Delta-Speichern. Gleich heißt hier STRENG gleich mit dem Wert, der
 * gerade in der Zeile steht — nur dann darf eine Spalte entfallen, denn dann
 * ändert ihr Schreiben nichts. `null` und `""` gelten als verschieden (der
 * Outlook-Flow rechnet mit coalesce, ein leerer String ist dort kein Fehlen),
 * ein anders formatiertes Datum ebenso (dann wird eben geschrieben). Eine
 * Spalte, die im Stand fehlt (nicht im $select), wird immer geschrieben.
 */
function gleicherWert(spalte: string, alt: unknown, neu: unknown): boolean {
  if (alt === undefined) return false;
  const leer = (v: unknown): boolean => v === null || v === undefined || v === '';
  // v32.6: null und "" sind für die App dasselbe. v32.9: auch in den
  // Outlook-Spalten — SharePoint legt einen leeren Text ohnehin als null ab,
  // ein geschriebenes "" liest der Flow also genauso als fehlend. Die
  // Ausnahme aus v32.6 ließ OutlookSubject bei jedem Speichern „geändert"
  // erscheinen (Tenant-Log 28.09.2026).
  void spalte;
  if (leer(alt) && leer(neu)) return true;
  if (leer(alt) || leer(neu)) return false;
  if (typeof alt === 'object' || typeof neu === 'object') {
    try { return JSON.stringify(alt) === JSON.stringify(neu); } catch { return false; }
  }
  // v32.10: Rich-Text-Spalten (z. B. Organizer, OrganizerEmail) legt
  // SharePoint verpackt ab: <div class="ExternalClass…">Inhalt</div>. Der
  // Inhalt ist derselbe, nur die Hülle nicht — ohne diesen Vergleich galten
  // beide Spalten bei jedem Speichern als geändert (Tenant-Log 28.09.2026).
  if (typeof alt === 'string' && typeof neu === 'string' && alt.indexOf('<div class="ExternalClass') === 0 && neu.indexOf('<') < 0) {
    const klar = alt.replace(/<[^>]*>/g, '')
      .replace(/&nbsp;|&#160;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
      .replace(/\u200b/g, '').trim();
    return klar === neu.trim();
  }
  // Datumswerte: SharePoint liefert „…T10:00:00Z", die App schreibt „…T10:00:00.000Z".
  const iso = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/;
  if (typeof alt === 'string' && typeof neu === 'string' && iso.test(alt) && iso.test(neu)) {
    const ta = Date.parse(alt); const tn = Date.parse(neu);
    return !isNaN(ta) && ta === tn;
  }
  return alt === neu;
}

export async function updateEvent(svc: EventService, eventId: number, updates: Record<string, unknown>, retried?: boolean, baseline?: Record<string, unknown>): Promise<boolean> {
  svc.lastUpdateEventError = '';
  try {
    // v28.66: zentraler Schutz für EndDate — analog zu createEvent. Ein
    // leeres EndDate in DEX_Events lässt den DEX_CreateOutlookEvent-Flow in
    // „Create event (V4)" mit convertFromUtc(null) abstürzen. Deshalb hier,
    // am gemeinsamen Nadelöhr aller Update-Pfade, aufräumen:
    //  - leeres EndDate + StartDate im selben Update -> Start als Ende,
    //  - sonst das Feld weglassen, statt einen gespeicherten Wert mit null
    //    zu überschreiben (leer war ohnehin nie ein gültiger Zustand).
    const safeUpdates: Record<string, unknown> = { ...updates };
    // v32.0.10: Bilder aus den Overrides in ihre Spalten (s. logosAuslagern).
    await logosAuslagern(svc, safeUpdates);
    if ('EndDate' in safeUpdates && !safeUpdates.EndDate) {
      if (safeUpdates.StartDate) {
        safeUpdates.EndDate = safeUpdates.StartDate;
      } else {
        delete safeUpdates.EndDate;
      }
    }
    // v32.5: Delta — nach allen Umformungen (Logos, EndDate) nur schreiben,
    // was sich gegenüber dem gelesenen Stand wirklich ändert. Nutzer-Befund
    // 28.09.2026: „warum dauert es überhaupt, wenn ich nichts geändert habe?"
    // Der Wizard schickt immer die komplette Zeile, samt Outlook-Text mit
    // eingebackenen Bildern (MB). Ohne Änderung entfällt der MERGE ganz —
    // und damit auch das Neu-Laden des Outlook-Texts nach dem Speichern
    // (die Zeile wird nicht neuer).
    if (baseline && Object.keys(baseline).length > 0) {
      // OutlookLogoBase64 steht nicht in EVENT_SELECT (die Spalte fehlt auf
      // älteren Sites, ein $select darauf wäre ein 400). Will logosAuslagern
      // sie schreiben, existiert sie — dann gezielt nachlesen, sonst ginge
      // das Kopfbild (bis 3,4 MB) bei jedem Speichern erneut hinaus.
      // v32.6: allgemein — jede zu schreibende Spalte, die im Stand fehlt
      // (z. B. OutlookIsOnlineMeeting), in EINEM Request nachlesen. Sie
      // existiert sicher, sonst schlüge der MERGE ohnehin fehl.
      const fehlend = Object.keys(safeUpdates).filter(k => !(k in (baseline as Record<string, unknown>)));
      if (fehlend.length > 0) {
        try {
          const r = await svc._sp.get(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=${fehlend.join(',')}`, SPHttpClient.configurations.v1);
          if (r.ok) {
            const j = await r.json();
            const nach: Record<string, unknown> = {};
            for (const k of fehlend) if (j && k in j) nach[k] = j[k];
            baseline = { ...baseline, ...nach };
          }
        } catch { /* ohne Grundlage werden die Spalten geschrieben */ }
      }
      const vorher = Object.keys(safeUpdates).length;
      for (const k of Object.keys(safeUpdates)) {
        if (gleicherWert(k, baseline[k], safeUpdates[k])) delete safeUpdates[k];
      }
      const nachher = Object.keys(safeUpdates).length;
      // eslint-disable-next-line no-console
      console.log(`[DEX][perf][updateEvent] Delta ${eventId}: ${nachher} von ${vorher} Spalten geändert${nachher ? ` (${Object.keys(safeUpdates).join(', ')})` : ' — kein Schreibvorgang'}`);
      // v32.9: Bei Textspalten zeigen, WO sich alt und neu unterscheiden —
      // Spalten, die ohne Änderung jedes Mal „geändert" sind, lassen sich
      // sonst nur raten.
      for (const k of Object.keys(safeUpdates)) {
        const a = baseline[k]; const n = safeUpdates[k];
        if (typeof a !== 'string' || typeof n !== 'string') continue;
        let i = 0; while (i < a.length && i < n.length && a.charCodeAt(i) === n.charCodeAt(i)) i++;
        // eslint-disable-next-line no-console
        console.log(`[DEX][perf][updateEvent] ${k}: Länge ${a.length} → ${n.length}, erster Unterschied bei ${i}: alt „${a.slice(Math.max(0, i - 30), i + 50)}" · neu „${n.slice(Math.max(0, i - 30), i + 50)}"`);
      }
      if (nachher === 0) return true;
    }
    const payload = {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_EventsListItem' },
      ...safeUpdates,
    };

    // v28.10: SharePoint lehnt REST-Bodies > 2 MB mit einem kryptischen
    // HTTP 400 ab („The request message is too big"). Vorab prüfen und
    // eine verständliche Meldung liefern — Verursacher ist praktisch
    // immer ein zu großes eingebettetes Bild (Mail-/Outlook-Logo oder
    // ein ins Mail-/Termin-Template eingefügtes Bild).
    const LIMIT = 1_900_000;
    const payloadStr = JSON.stringify(payload);
    if (payloadStr.length > LIMIT) {
      // v28.31: Statt aufzugeben in MEHREREN Requests nacheinander schreiben.
      // Das 2-MB-Limit gilt pro REST-Aufruf, nicht pro Item — ein Event mit
      // eingebetteten Bildern (OutlookBody + EmailTemplateOverrides +
      // EmailImageBase64 tragen dasselbe Bild je einmal) passt problemlos,
      // wenn man die Felder auf mehrere MERGEs verteilt. Vorher brach der
      // Save hier still ab: In der Konsole stand nur eine Warnung, im Wizard
      // passierte auf „Speichern" schlicht nichts.
      const FIELD_OVERHEAD = 160; // __metadata + Klammern/Kommas
      const entries = Object.keys(safeUpdates)
        .map(k => ({ k, size: JSON.stringify({ [k]: safeUpdates[k] }).length }))
        .sort((a, b) => b.size - a.size);
      // Ein EINZELNES Feld über dem Limit lässt sich nicht aufteilen — hier
      // hilft nur ein kleineres Bild. Feldname mitgeben, damit der Organizer
      // weiß, wo er suchen muss.
      const tooBig = entries.filter(e => e.size + FIELD_OVERHEAD > LIMIT);
      if (tooBig.length > 0) {
        svc.lastUpdateEventError = `Ein einzelnes Feld ist zu groß für SharePoint (${tooBig.map(e => `${e.k}: ${Math.round(e.size / 1024)} KB`).join(', ')}). Ursache ist praktisch immer ein zu großes eingebettetes Bild (Mail-Logo, Outlook-Kopfbild oder ein Bild im Mail-/Termin-Text). Bitte das Bild entfernen oder kleiner erneut hochladen.`;
        console.warn('[DEX] updateEvent: einzelnes Feld über dem Limit —', tooBig);
        return false;
      }
      const groups: Array<Record<string, unknown>> = [];
      let cur: Record<string, unknown> = {};
      let curSize = FIELD_OVERHEAD;
      for (const e of entries) {
        if (curSize + e.size > LIMIT && Object.keys(cur).length > 0) {
          groups.push(cur); cur = {}; curSize = FIELD_OVERHEAD;
        }
        cur[e.k] = safeUpdates[e.k];
        curSize += e.size;
      }
      if (Object.keys(cur).length > 0) groups.push(cur);
      console.warn(`[DEX] updateEvent: Payload ${payloadStr.length} Bytes > Limit — wird in ${groups.length} aufeinanderfolgende Schreibvorgänge aufgeteilt.`);
      for (let i = 0; i < groups.length; i++) {
        const ok = await svc.updateEvent(eventId, groups[i], retried);
        if (!ok) {
          // lastUpdateEventError kommt aus dem fehlgeschlagenen Teil-Request.
          svc.lastUpdateEventError = `Teil ${i + 1} von ${groups.length} konnte nicht gespeichert werden. ${svc.lastUpdateEventError}`.trim();
          return false;
        }
      }
      return true;
    }

    const response = await svc._sp.post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE',
          'odata-version': '',
        },
        body: payloadStr,
      }
    );
    if (!response.ok) {
      const errText = await response.text().catch(() => '');
      console.warn('[DEX] updateEvent failed:', response.status, errText.substring(0, 400));
      // SharePoint-Fehlertext extrahieren (verbose: error.message.value).
      let spMsg = '';
      try {
        const parsed = JSON.parse(errText);
        spMsg = parsed?.error?.message?.value || parsed?.['odata.error']?.message?.value || '';
      } catch { /* kein JSON */ }
      const statusHint = response.status === 403
        ? 'Keine Berechtigung — du brauchst Schreibrechte auf der Event-Liste (Organizer/Admin).'
        : response.status === 404
          ? 'Das Event wurde in der Liste nicht gefunden — womöglich wurde es zwischenzeitlich gelöscht.'
          : response.status === 409 || response.status === 412
            ? 'Das Event wurde zeitgleich von jemand anderem geändert — bitte neu laden und erneut speichern.'
            : '';
      svc.lastUpdateEventError = [`HTTP ${response.status}`, statusHint, spMsg && spMsg !== statusHint ? spMsg.slice(0, 300) : '']
        .filter(Boolean).join(' — ');

      // v26.54: „Invalid text value" = ein String-Wert passt nicht in eine
      // EINZEILIGE Text-Spalte (255-Zeichen-Limit). SharePoint nennt das
      // betroffene Feld nicht — wir diagnostizieren selbst: Payload-Werte
      // gegen die Live-Feldtypen der Liste halten. Spalten, die laut
      // Schema-Definition ohnehin mehrzeilig (Note) sein sollten, werden
      // sofort migriert und der Save EINMAL automatisch wiederholt. Alle
      // anderen Treffer werden in der Fehlermeldung beim Namen genannt.
      if (/invalid text value|text field contains invalid data/i.test(spMsg)) {
        const offenders = await findInvalidTextFields(svc, 'DEX_Events', safeUpdates);
        if (offenders.length > 0) {
          console.warn('[DEX] updateEvent: Werte passen nicht in einzeilige Text-Spalten:',
            offenders.map((o) => `${o.internalName} (${o.length} Zeichen${o.intendedNote ? ', sollte Note sein' : ''})`).join(', '));
          const healable = offenders.filter((o) => o.intendedNote);
          if (!retried && healable.length > 0) {
            for (const o of healable) {
              await svc._upgradeTextFieldToNote('DEX_Events', o.title);
            }
            return svc.updateEvent(eventId, safeUpdates, true);
          }
          svc.lastUpdateEventError += ` | ${offenders
            .map((o) => `Betroffenes Feld: „${o.title}" — ${o.length} Zeichen, die Spalte ist einzeiliger Text (max. 255 Zeichen)`)
            .join('; ')}`;
        }
      }
    }
    return response.ok;
  } catch (err) {
    svc.lastUpdateEventError = `Netzwerkfehler — keine Verbindung zu SharePoint${err instanceof Error && err.message ? ` (${err.message.slice(0, 150)})` : ''}.`;
    return false;
  }
}

/**
 * v26.54: Diagnose-Helfer für „Invalid text value. A text field contains
 * invalid data." beim Event-Update. Findet alle String-Werte im Update-
 * Payload, die zu lang für eine einzeilige Text-Spalte sind (> 255 Zeichen
 * oder mit Zeilenumbrüchen), deren Ziel-Spalte auf der LIVE-Liste aber
 * tatsächlich als einzeiliger Text ('Text') liegt. `intendedNote` markiert
 * Spalten, die laut Schema-Definition (getEventsFieldDefinitions) eigentlich
 * mehrzeilig (Typ 3, Note) sein sollten — die dürfen automatisch per
 * _upgradeTextFieldToNote geheilt werden.
 */
async function findInvalidTextFields(
  svc: EventService,
  listName: string,
  updates: Record<string, unknown>
): Promise<Array<{ internalName: string; title: string; length: number; intendedNote: boolean }>> {
  const out: Array<{ internalName: string; title: string; length: number; intendedNote: boolean }> = [];
  try {
    const resp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName,Title,TypeAsString&$filter=Hidden eq false&$top=300`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return out;
    const data = await resp.json();
    const fields: Array<{ InternalName: string; Title: string; TypeAsString: string }> = data.value || [];
    const defs = svc.getEventsFieldDefinitions();
    for (const key of Object.keys(updates)) {
      const v = updates[key];
      if (typeof v !== 'string') continue;
      if (v.length <= 255 && v.indexOf('\n') < 0) continue;
      const f = fields.filter((x) => x.InternalName === key)[0];
      if (!f || f.TypeAsString !== 'Text') continue;
      const def = defs.filter((d) => d.title === key)[0];
      out.push({ internalName: key, title: f.Title, length: v.length, intendedNote: !!def && def.type === 3 });
    }
  } catch { /* Diagnose darf den Fehlerpfad nie zusätzlich brechen */ }
  return out;
}

/**
 * Event vollständig löschen:
 * 1. Subsite löschen (inkl. Teilnehmerliste) - für neue Events
 * 2. Alte Registrierungsliste löschen (DEX_Reg_*) - für alte Events
 * 3. Event-Eintrag aus DEX_Events löschen
 */
export async function deleteEvent(svc: EventService, eventId: number): Promise<boolean> {
  try {
    // Event-Daten laden um SubsiteUrl und RegistrationListName zu bekommen
    const event = await svc.getEvent(eventId);
    if (!event) return false;

    // v30.67 (Review): Teilt eine ZWEITE DEX_Events-Zeile diese Subsite
    // (Duplikat-Zustand aus einem halb gescheiterten Recreate, s.
    // persistSubEvents.recreateWithReuse), darf die Subsite NICHT recycelt
    // werden — sie trägt die Anmeldungen der anderen Zeile. Dann wird unten
    // nur die Zeile entfernt. Ist die Abfrage nicht lesbar, wird gar nichts
    // gelöscht: unbekannt sperrt. Die Prüfung steht VOR dem Register-Lauf,
    // damit bei Abbruch nichts angefasst ist.
    let subsiteShared = false;
    if (event.SubsiteUrl) {
      try {
        const shareResp = await svc._sp.get(
          `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=Id&$filter=SubsiteUrl eq '${encodeURIComponent(event.SubsiteUrl.replace(/'/g, "''"))}' and Id ne ${eventId}&$top=1`,
          SPHttpClient.configurations.v1
        );
        if (!shareResp.ok) {
          console.warn(`[DEX] deleteEvent: Prüfung auf geteilte Subsite nicht möglich (HTTP ${shareResp.status}) — Löschung abgebrochen (Event ${eventId}).`);
          return false;
        }
        const shareData = await shareResp.json();
        const others: Array<{ Id: number }> = shareData.value || shareData.d?.results || [];
        subsiteShared = others.length > 0;
        if (subsiteShared) {
          console.warn(`[DEX] deleteEvent: Subsite ${event.SubsiteUrl} wird auch von Item ${others[0].Id} genutzt — nur die DEX_Events-Zeile ${eventId} wird entfernt, die Subsite bleibt.`);
        }
      } catch (err) {
        console.warn('[DEX] deleteEvent: Prüfung auf geteilte Subsite fehlgeschlagen — Löschung abgebrochen:', err);
        return false;
      }
    }

    // v30.67: Register ZUERST — dieselbe Reihenfolge wie deleteParticipantData
    // seit v29.3 (CLAUDE.md „Löschungen zuerst im Register"). Bisher wurde die
    // Subsite recycelt und DANACH DEX_Participants aufgeräumt, und zwar so,
    // dass ein Scheitern nicht auffiel: nicht strikt gelesen (eine abgebrochene
    // Seite kam still als vollständige Liste), alle MERGEs gleichzeitig
    // (Promise.all — die Einladung zur Drosselung) und jeder Fehler per
    // `.catch(() => null)` verworfen. Was im Register stehen blieb, ließ sich
    // nicht mehr nachrechnen, weil die Teilnehmerliste schon weg war — genau
    // die „Verweis ohne Zeile"-Rückstände der Register-Prüfung. Jetzt: strikt
    // lesen, sequentiell schreiben, Fehler zählen, und bei Fehlern mit `false`
    // abbrechen, BEVOR irgendetwas recycelt oder in die Outlook-Queue gestellt
    // wird. Das Event steht dann noch; der Löschversuch lässt sich wiederholen.
    if (event.EventNumber) {
      let registryFailed = 0;
      try {
        const allParticipants = await svc.fetchAllParticipantsOrThrow();
        const en = String(event.EventNumber);
        const affected = allParticipants.filter(p =>
          (p.EventRegistered?.split(',').map(s => s.trim()).includes(en))
          || (p.EventOnWaitlist?.split(',').map(s => s.trim()).includes(en)));
        for (const p of affected) {
          // eslint-disable-next-line no-await-in-loop
          const ok = await svc.removeParticipantEvent(p.Email, event.EventNumber).catch(() => false);
          if (!ok) registryFailed += 1;
        }
      } catch (err) {
        console.warn('[DEX] deleteEvent: Register nicht vollständig lesbar — Löschung abgebrochen:', err);
        return false;
      }
      if (registryFailed > 0) {
        console.warn(`[DEX] deleteEvent: ${registryFailed} Register-Einträge nicht aktualisiert — Löschung abgebrochen (Event ${event.EventNumber}).`);
        return false;
      }
    }

    // 0. Outlook-Kalendereintrag per Queue löschen (VOR allem anderen, damit
    //    CalendarLink noch vorhanden ist). Der DEX_Outlook_Einladungen-Flow
    //    greift den DeleteEvent-Eintrag auf und löscht den Kalender-Termin
    //    im Shared Mailbox über den Flow-Service-Account.
    //    Fehler hier ignorieren - Event-Delete soll trotzdem durchlaufen.
    if (event.CalendarLink) {
      try {
        await svc.queueOutlookDeleteEvent(String(eventId), event.Title || '', event.CalendarLink);
      } catch { /* Queue-Fehler ignorieren */ }
    }
    // v31.16: Das Wartelisten-Schattenevent mit abräumen — ebenfalls VOR
    // allem anderen. Ein Schattenevent, dessen echtes Event weg ist, ist ein
    // Termin in fremden Kalendern, den niemand mehr absagen kann; und seine
    // Zeile in DEX_Events wäre für immer da, unsichtbar (der Filter in
    // `loadEvents` blendet sie aus) und ohne Bezug. Fehler hier ignorieren —
    // der Event-Delete soll trotzdem durchlaufen.
    try {
      await svc.removeWaitlistShadow(String(eventId), event.Title || '');
    } catch { /* best-effort */ }
    // 1. Subsite RECYCEN (v9.0: nicht mehr per DELETE, sonst landet die
    //    Subsite permanent weg ohne Recycle-Bin-Eintrag. recycle() legt
    //    die Subsite mitsamt Teilnehmerliste 93 Tage in den Site
    //    Collection Recycle Bin → ein Tenant-Admin / Site Collection
    //    Admin kann sie dort wiederherstellen falls nötig.
    if (event.SubsiteUrl && !subsiteShared) {
      try {
        await svc._post(`${event.SubsiteUrl}/_api/web/recycle`, {});
      } catch {
        console.warn('[DEX] Subsite konnte nicht in den Recycle Bin verschoben werden:', event.SubsiteUrl);
      }
    }

    // 2. Event-Bild ebenfalls RECYCEN statt löschen.
    if (event.EventImageUrl) {
      try {
        const url = new URL(event.EventImageUrl);
        const serverRelUrl = url.pathname;
        if (serverRelUrl.indexOf('DEX_EventImages') >= 0) {
          await svc._post(
            `${svc.siteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelUrl}')/recycle`,
            {}
          );
        }
      } catch {
        console.warn('[DEX] Event-Bild konnte nicht in den Recycle Bin verschoben werden');
      }
    }

    // 3. Alte Registrierungsliste recyceln (legacy Events ohne Subsite).
    if (event.RegistrationListName && event.RegistrationListName !== 'Teilnehmer') {
      try {
        await svc._post(
          `${svc.siteUrl}/_api/web/lists/getbytitle('${event.RegistrationListName.replace(/'/g, "''")}')/recycle`,
          {}
        );
      } catch {
        console.warn('[DEX] Alte Registrierungsliste konnte nicht recycelt werden:', event.RegistrationListName);
      }
    }

    // (v30.67: Der DEX_Participants-Block stand hier — er läuft jetzt ganz
    // oben, VOR dem ersten Recycle, siehe Kommentar dort.)

    // 4. Event-Dokumente löschen (SiteAssets/DEX_EventDocs/Event_{number}_*)
    if (event.EventNumber) {
      try {
        const serverRelUrl = svc.context.pageContext.web.serverRelativeUrl;
        const safeName = (event.Title || '').replace(/[#%&*:<>?/\\|"']/g, '').replace(/\s+/g, '_').substring(0, 50);
        const folderName = safeName ? `Event_${event.EventNumber}_${safeName}` : `Event_${event.EventNumber}`;
        // v31.62: Ordner in den Papierkorb statt DELETE — die Dokumente des
        // Events sind 93 Tage zurückholbar. `_post` wirft nicht bei 404; ein
        // fehlender Ordner ist kein Fehler.
        const r1 = await svc._post(`${svc.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelUrl}/SiteAssets/DEX_EventDocs/${folderName}')/recycle`, {});
        if (!r1.ok) {
          // Fallback: alten Ordnernamen ohne Titel probieren
          await svc._post(`${svc.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelUrl}/SiteAssets/DEX_EventDocs/Event_${event.EventNumber}')/recycle`, {});
        }
      } catch { /* Ordner nicht gefunden */ }
    }

    // 5. Event-Eintrag aus DEX_Events RECYCEN (v9.0: per recycle() statt
    //    delete(), damit ein Admin via SharePoint-Recycle-Bin das Item
    //    bei Bedarf 93 Tage lang wiederherstellen kann).
    const response = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/recycle`,
      {}
    );

    // 6. Audit-Eintrag in DEX_ChangeLog (v9.0). Best-effort, blockt
    //    den Lösch-Vorgang nicht falls Logging fehlschlägt.
    try {
      await svc.writeChangeLog({
        action: 'EventDeletedTest', // wird vom Aufrufer überschrieben
        targetType: 'Event',
        targetId: String(eventId),
        targetName: event.Title || '',
        eventId: String(eventId),
        eventTitle: event.Title || '',
        details: {
          subsiteUrl: event.SubsiteUrl || '',
          eventNumber: event.EventNumber,
          recycledTo: 'SharePoint Recycle Bin (93 Tage)',
          // v31.62: die Zeile selbst — nach 93 Tagen ist der Papierkorb leer,
          // der Schnappschuss nicht (CalendarLink, ParentEventId, Fristen …).
          snapshot: kompakteZeile(event as unknown as Record<string, unknown>),
        },
      });
    } catch { /* */ }

    return response.ok;
  } catch {
    return false;
  }
}

/** v26.13: Versions-Historie der CustomFields-Spalte eines DEX_Events-Items
 *  (neueste zuerst). Grundlage für die Wiederherstellung versehentlich
 *  überschriebener Custom-Field-Beschreibungen (helpText etc.) aus der
 *  SharePoint-Versionshistorie. */
export async function getEventCustomFieldsVersions(svc: EventService, itemId: number): Promise<Array<{ created: string; customFields: string }>> {
  // WICHTIG (v26.15): $select=Created,CustomFields ist PFLICHT — sonst liefert
  // der versions-Endpunkt ALLE Felder pro Version (inkl. der riesigen
  // OutlookBody-/EmailTemplateOverrides-Base64-Logos). Bei stark bearbeiteten
  // Events (z.B. 188 Versionen) sprengt das die Antwortgröße und SharePoint
  // bricht nach ~51 Versionen ab → die Version MIT der Beschreibung fehlte und
  // es kam fälschlich „helpText in Historie: false". KEIN $orderby (das löst
  // auf dem versions-Endpunkt 400 aus) — wir sortieren clientseitig nach
  // Created absteigend (neueste zuerst). Folgeseiten via nextLink einsammeln.
  const out: Array<{ created: string; customFields: string }> = [];
  let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${itemId})/versions?$select=Created,CustomFields&$top=500`;
  let guard = 0;
  try {
    while (url && guard < 25) {
      guard++;
      const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) { console.warn('[DEX restore] versions HTTP', resp.status, 'für Item', itemId); break; }
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const v of (items as any[])) {
        out.push({ created: v.Created || '', customFields: typeof v.CustomFields === 'string' ? v.CustomFields : '' });
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      url = (data as any)['@odata.nextLink'] || (data.d && (data.d as any).__next) || null;
    }
  } catch (e) {
    console.warn('[DEX restore] versions fetch failed für Item', itemId, e);
  }
  out.sort((a, b) => (b.created || '').localeCompare(a.created || ''));
  return out;
}

/**
 * v32.0.10: Bestand umstellen — die Bilder aller Events aus
 * EmailTemplateOverrides in ihre Spalten verlagern (s. logosAuslagern).
 * Rollenverwaltung → „Mail-Bilder auslagern (alle Events)".
 *
 * `vorschau` liest nur und zählt. Sonst je Zeile, nacheinander: über
 * updateEvent schreiben (derselbe Engpass wie jeder Save), danach die Zeile
 * NACHLESEN und die Bildlängen vergleichen — stimmt etwas nicht, wird die
 * alte Zeile zurückgeschrieben. Drei Fehler in Folge brechen ab (Drosselung
 * oder Rechte; weiterzumachen hieße nur, dieselbe Ablehnung zu sammeln).
 */
export async function logosAuslagernAlle(
  svc: EventService,
  vorschau: boolean,
  onProgress?: (_done: number, _total: number, _label: string) => void,
): Promise<{ zeilen: number; kb: number; umgestellt: number; fehler: string[]; abgebrochen: boolean; ohneSpalte: boolean }> {
  const out = { zeilen: 0, kb: 0, umgestellt: 0, fehler: [] as string[], abgebrochen: false, ohneSpalte: false };
  if (!(await hatEventsSpalte(svc, 'OutlookLogoBase64'))) { out.ohneSpalte = true; return out; }
  // Alle Zeilen, seitenweise — die Aktion soll auch Events jenseits der
  // 100 des Starts erfassen.
  const kandidaten: Array<{ id: number; title: string; raw: string; mail: number; outlook: number }> = [];
  let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items?$select=Id,Title,EmailTemplateOverrides&$top=200`;
  while (url) {
    const r = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (!r.ok) { out.fehler.push(`DEX_Events nicht lesbar (HTTP ${r.status})`); out.abgebrochen = true; return out; }
    const d = await r.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    for (const row of (d.value || []) as any[]) {
      let o: Record<string, unknown> = {};
      try { o = JSON.parse(row.EmailTemplateOverrides || '{}') || {}; } catch { continue; }
      const m = typeof o._eventLogo === 'string' ? o._eventLogo.length : -1;
      const ol = typeof o._outlookLogo === 'string' ? o._outlookLogo.length : -1;
      if (m < 0 && ol < 0) continue;
      kandidaten.push({ id: Number(row.Id), title: String(row.Title || row.Id), raw: String(row.EmailTemplateOverrides || ''), mail: m, outlook: ol });
      out.kb += (Math.max(0, m) + Math.max(0, ol)) / 1024;
    }
    url = d['odata.nextLink'] || d['@odata.nextLink'] || null;
  }
  out.zeilen = kandidaten.length;
  out.kb = Math.round(out.kb);
  if (vorschau) return out;
  let inFolge = 0;
  for (let i = 0; i < kandidaten.length; i++) {
    const k = kandidaten[i];
    if (onProgress) onProgress(i, kandidaten.length, k.title);
    const ok = await updateEvent(svc, k.id, { 'EmailTemplateOverrides': k.raw });
    let geprueft = false;
    if (ok) {
      try {
        const r = await svc._sp.get(
          `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${k.id})?$select=EmailTemplateOverrides,EmailImageBase64,OutlookLogoBase64`,
          SPHttpClient.configurations.v1
        );
        if (r.ok) {
          const z = await r.json();
          let o: Record<string, unknown> = {};
          try { o = JSON.parse(z.EmailTemplateOverrides || '{}') || {}; } catch { o = {}; }
          const mailOk = k.mail <= 0 || String(z.EmailImageBase64 || '').length === k.mail;
          const outlookOk = k.outlook <= 0 || String(z.OutlookLogoBase64 || '').length === k.outlook;
          geprueft = mailOk && outlookOk && !!o._logosAusgelagert;
        }
      } catch { geprueft = false; }
    }
    if (ok && geprueft) { out.umgestellt++; inFolge = 0; continue; }
    // Zurück auf den alten Stand: das JSON mit Bildern, Spalten unberührt lassen.
    try { await svc._merge(`${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${k.id})`, { 'EmailTemplateOverrides': k.raw }); } catch { /* s. Fehlerliste */ }
    out.fehler.push(k.title);
    if (++inFolge >= 3) { out.abgebrochen = true; break; }
  }
  if (onProgress) onProgress(kandidaten.length, kandidaten.length, '');
  try {
    await svc.writeChangeLog({
      action: 'LogosAusgelagert', targetType: 'Event', targetId: '', targetName: 'DEX_Events',
      details: { zeilen: out.zeilen, umgestellt: out.umgestellt, fehler: out.fehler, abgebrochen: out.abgebrochen, kb: out.kb },
    });
  } catch { /* Protokoll best-effort */ }
  return out;
}
