/**
 * v30.66 — Modularisierung Stufe 2: Thema „Dateien und Overrides eines
 * einzelnen Events": Event-Bild (Item-Attachment, `EventImageUrl`), das
 * unbeschnittene Original, die Dokumente unter SiteAssets/DEX_EventDocs und
 * die Schreibzugriffe auf `EmailTemplateOverrides`.
 *
 * Die Overrides-Spalte ist ein einziges JSON, in dem mehrere Stellen der App
 * unabhängig voneinander Schlüssel ändern (Piggyback-Konfiguration, siehe
 * CLAUDE.md). Deshalb laufen alle Schreiber über `svc._ovQueue` streng
 * nacheinander — die Warteschlange ist Instanz-Zustand und bleibt an der
 * Klasse. Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import type { EventService } from '../EventService';

// v28.11: Präfix des UNBESCHNITTENEN Original-Bilds (bewusst KEIN
// '__eventimage__'-Präfix-Match, sonst würde der normale Bild-Upload es
// mitlöschen). Wird nur gespeichert, wenn ein Querformat-Original per
// App-Zuschnitt rund/quadratisch wurde — die Anmeldeseite zeigt dann
// lieber das Original im Querformat-Slot.
// v30.66: war `private static readonly` an der Klasse — nur hier gebraucht.
const ORIG_IMAGE_PREFIX = '__eventimgorig__';

/**
 * Event-Bild als Attachment an ein DEX_Events-Item anhängen.
 * Löscht zuerst alle bestehenden Bild-Attachments (Präfix __eventimage__),
 * dann wird das neue Bild hochgeladen. Liefert die ServerRelativeUrl als absolute URL.
 * Vorteil: keine SiteAssets-Berechtigungen nötig - wer das Item editieren darf,
 * darf auch Attachments hinzufügen.
 */
export async function uploadEventImageAsAttachment(svc: EventService, eventId: number, file: File): Promise<string> {
  const IMAGE_PREFIX = '__eventimage__';
  try {
    // 1. Bestehende Bild-Attachments löschen (nur __eventimage__-Präfixe)
    try {
      const listResp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles`,
        SPHttpClient.configurations.v1
      );
      if (listResp.ok) {
        const listData = await listResp.json();
        const files = listData.value || listData.d?.results || [];
        for (const f of files) {
          const fn: string = f.FileName || '';
          if (fn.indexOf(IMAGE_PREFIX) === 0) {
            try {
              await svc._delete(
                `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/getByFileName('${encodeURIComponent(fn)}')`
              );
            } catch { /* ignore */ }
          }
        }
      }
    } catch { /* ignore */ }

    // 2. Neues Bild hochladen mit Präfix
    const ext = (file.name.split('.').pop() || 'jpg').toLowerCase();
    const safeName = `${IMAGE_PREFIX}${Date.now().toString(36)}.${ext}`;

    const response = await svc._sp.post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/add(FileName='${encodeURIComponent(safeName)}')`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
        body: file,
      } as ISPHttpClientOptions
    );

    if (response.ok) {
      const data = await response.json();
      const relUrl = data.d?.ServerRelativeUrl || data.ServerRelativeUrl || '';
      if (relUrl) return `${window.location.origin}${relUrl}`;
      // Hochgeladen, aber ohne URL in der Antwort: Pfad ist bekannt und die
      // Datei liegt dort — hier darf geraten werden.
      const serverRelUrl = svc.context.pageContext.web.serverRelativeUrl;
      return `${window.location.origin}${serverRelUrl}/Lists/DEX_Events/Attachments/${eventId}/${safeName}`;
    }
    // v29.34: Bei einer FEHLER-Antwort keine URL mehr raten. Die geratene
    // Adresse zeigte auf eine Datei, die nie ankam — gespeichert wurde sie
    // trotzdem, und das Bild fehlte danach dauerhaft. Leer heißt hier
    // „fehlgeschlagen"; die Aufrufer melden das dem Organizer.
    console.warn('[DEX] Image attachment upload status:', response.status);
  } catch (err) {
    console.warn('[DEX] uploadEventImageAsAttachment error:', err);
  }
  return '';
}

/** v28.11: Alle Original-Bild-Attachments eines Events löschen (best-effort). */
export async function deleteEventOrigImageAttachment(svc: EventService, eventId: number): Promise<void> {
  try {
    const listResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles`,
      SPHttpClient.configurations.v1
    );
    if (!listResp.ok) return;
    const listData = await listResp.json();
    const files = listData.value || listData.d?.results || [];
    for (const f of files) {
      const fn: string = f.FileName || '';
      if (fn.indexOf(ORIG_IMAGE_PREFIX) === 0) {
        try {
          await svc._delete(
            `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/getByFileName('${encodeURIComponent(fn)}')`
          );
        } catch { /* ignore */ }
      }
    }
  } catch { /* ignore */ }
}

/** v28.11: Unbeschnittenes Original-Bild als zweites Attachment speichern.
 *  Löscht vorher bestehende Originale; liefert die absolute URL. */
export async function uploadEventOrigImageAsAttachment(svc: EventService, eventId: number, file: File): Promise<string> {
  try {
    await svc.deleteEventOrigImageAttachment(eventId);
    const ext = (file.name.split('.').pop() || 'jpg').toLowerCase();
    const safeName = `${ORIG_IMAGE_PREFIX}${Date.now().toString(36)}.${ext}`;
    const response = await svc._sp.post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/add(FileName='${encodeURIComponent(safeName)}')`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
        body: file,
      } as ISPHttpClientOptions
    );
    if (response.ok) {
      const data = await response.json();
      const relUrl = data.d?.ServerRelativeUrl || data.ServerRelativeUrl || '';
      if (relUrl) return `${window.location.origin}${relUrl}`;
      const serverRelUrl = svc.context.pageContext.web.serverRelativeUrl;
      return `${window.location.origin}${serverRelUrl}/Lists/DEX_Events/Attachments/${eventId}/${safeName}`;
    }
    // v29.34: Bei Fehler-Antwort NICHT raten — die geratene URL landete in
    // `_imageOrigUrl`, und die Kachel bevorzugt diese Adresse vor dem
    // Event-Bild. Ergebnis war eine weiße Kachel bei einem Event, das im
    // Organizer Center sein Bild hatte. Leer lässt den bisherigen Wert
    // stehen (der Aufrufer patcht nur bei nicht-leerem Ergebnis).
    console.warn('[DEX] Orig image attachment upload status:', response.status);
  } catch (err) {
    console.warn('[DEX] uploadEventOrigImageAsAttachment error:', err);
  }
  return '';
}

/** v28.11: EINEN Schlüssel im EmailTemplateOverrides-JSON eines Events
 *  patchen (read-modify-write). Leerer Wert entfernt den Schlüssel.
 *  Nötig für Werte, die erst NACH dem Item-Save bekannt sind (z.B. die
 *  Attachment-URL des Original-Bilds). */
export async function patchEventOverridesKey(svc: EventService, eventId: number, key: string, value: string): Promise<void> {
  try {
    const getResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=EmailTemplateOverrides`,
      SPHttpClient.configurations.v1
    );
    if (!getResp.ok) return;
    const data = await getResp.json();
    const raw = data.d?.EmailTemplateOverrides || data.EmailTemplateOverrides || '';
    let obj: Record<string, unknown> = {};
    try { obj = raw ? JSON.parse(raw) : {}; } catch { obj = {}; }
    if (value) obj[key] = value; else delete obj[key];
    await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
      { 'EmailTemplateOverrides': JSON.stringify(obj) }
    );
  } catch (err) {
    console.warn('[DEX] patchEventOverridesKey fehlgeschlagen:', key, err);
  }
}

/**
 * v28.38: Wie `patchEventOverridesKey`, aber für beliebige JSON-Werte
 * (Arrays, Booleans, Objekte). `undefined`/`null`/leeres Array löschen den
 * Schlüssel, damit die Overrides nicht mit leeren Huellen zuwachsen.
 * Liefert true bei Erfolg.
 */
export async function patchEventOverridesValue(svc: EventService, eventId: number, key: string, value: unknown): Promise<boolean> {
  const res = await svc.patchEventOverridesValueEx(eventId, key, value);
  return res.ok;
}

/**
 * v28.60: Wie `patchEventOverridesValue`, liefert aber den Grund mit.
 *
 * Vorher gab es nur true/false — bei einem Fehlschlag stand im UI „konnte
 * nicht gespeichert werden" und sonst nichts, was die Ursachensuche
 * unmöglich machte. Jetzt kommt der HTTP-Status samt SharePoint-Meldung
 * zurück, und die drei typischen Stolpersteine sind abgefangen:
 *
 *  - **Transiente Fehler** (429 Throttling, 5xx): ein Wiederholungsversuch
 *    nach kurzer Pause statt sofort aufzugeben.
 *  - **Grössenlimit**: SharePoint lehnt Requests über 2 MB ab. Das Feld
 *    trägt bei Events mit eingebetteten Logos einiges — wir prüfen vorher
 *    und sagen es klar, statt in ein nacktes HTTP 400 zu laufen.
 *  - **Parallele Schreibvorgänge**: Der Aufruf ist ein Read-Modify-Write.
 *    Zwei gleichzeitige Aufrufe (z.B. schnell hintereinander geklickte
 *    Löschungen) würden sich gegenseitig überschreiben, deshalb laufen sie
 *    über `_ovQueue` streng nacheinander.
 */
export async function patchEventOverridesValueEx(
  svc: EventService,
  eventId: number, key: string, value: unknown,
): Promise<{ ok: boolean; status: number; detail: string }> {
  const run = async (): Promise<{ ok: boolean; status: number; detail: string }> => {
    const attempt = async (): Promise<{ ok: boolean; status: number; detail: string }> => {
      const getResp = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})?$select=EmailTemplateOverrides`,
        SPHttpClient.configurations.v1
      );
      if (!getResp.ok) {
        return { ok: false, status: getResp.status, detail: `Lesen fehlgeschlagen (HTTP ${getResp.status})` };
      }
      const data = await getResp.json();
      const raw = data.d?.EmailTemplateOverrides || data.EmailTemplateOverrides || '';
      let obj: Record<string, unknown> = {};
      try { obj = raw ? JSON.parse(raw) : {}; } catch { obj = {}; }
      const empty = value === undefined || value === null || value === false
        || (Array.isArray(value) && value.length === 0);
      if (empty) { delete obj[key]; } else { obj[key] = value; }
      const payload = JSON.stringify(obj);
      if (payload.length > 1_900_000) {
        return {
          ok: false, status: 413,
          detail: `Der Event-Datensatz ist mit ${(payload.length / 1048576).toFixed(2)} MB zu gross für einen Schreibvorgang (Limit 2 MB).`,
        };
      }
      const resp = await svc._merge(
        `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
        { 'EmailTemplateOverrides': payload }
      );
      if (resp.ok || resp.status === 406) return { ok: true, status: resp.status, detail: '' };
      let body = '';
      try { body = await resp.text(); } catch { /* egal */ }
      // SharePoint verpackt die Meldung in {error:{message:{value}}}.
      let msg = '';
      try {
        const j = JSON.parse(body);
        msg = j?.error?.message?.value || j?.['odata.error']?.message?.value || '';
      } catch { msg = (body || '').substring(0, 200); }
      return { ok: false, status: resp.status, detail: msg || `HTTP ${resp.status}` };
    };

    for (let i = 0; i < 2; i++) {
      try {
        const r = await attempt();
        if (r.ok) return r;
        const transient = r.status === 429 || r.status >= 500;
        if (!transient || i === 1) {
          console.warn('[DEX] patchEventOverridesValue fehlgeschlagen:', key, r.status, r.detail);
          return r;
        }
      } catch (err) {
        if (i === 1) {
          console.warn('[DEX] patchEventOverridesValue Ausnahme:', key, err);
          return { ok: false, status: 0, detail: String((err as Error)?.message || err) };
        }
      }
      await new Promise<void>(res => setTimeout(res, 600));
    }
    return { ok: false, status: 0, detail: 'unbekannt' };
  };

  // Streng nacheinander — sonst gehen parallele Read-Modify-Writes verloren.
  const next = svc._ovQueue.then(run, run);
  svc._ovQueue = next.then(() => undefined, () => undefined);
  return next;
}

/**
 * EventImageUrl-Feld eines DEX_Events-Items setzen (kleines MERGE).
 */
export async function updateEventImageUrl(svc: EventService, eventId: number, url: string): Promise<boolean> {
  try {
    const resp = await svc._merge(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})`,
      { 'EventImageUrl': url }
    );
    return resp.ok || resp.status === 406;
  } catch {
    return false;
  }
}

/**
 * Dokument als Attachment an ein DEX_Events-Item anfügen.
 * Nutzt native SharePoint List Item Attachments - keine Ordner nötig.
 */
export async function uploadEventDocument(svc: EventService, eventId: number, file: File): Promise<string> {
  try {
    const fileName = file.name.replace(/[#%&*:<>?/\\|]/g, '_');

    const response = await svc._sp.post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/add(FileName='${encodeURIComponent(fileName)}')`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
        body: file,
      } as ISPHttpClientOptions
    );

    if (response.ok) {
      const data = await response.json();
      const relUrl = data.d?.ServerRelativeUrl || data.ServerRelativeUrl || '';
      if (relUrl) return `${window.location.origin}${relUrl}`;
    } else {
      console.warn('[DEX] Attachment upload status:', response.status);
    }

    // Fallback: URL aus bekanntem Pfad
    const serverRelUrl = svc.context.pageContext.web.serverRelativeUrl;
    return `${window.location.origin}${serverRelUrl}/Lists/DEX_Events/Attachments/${eventId}/${fileName}`;
  } catch (err) {
    console.warn('[DEX] uploadEventDocument error:', err);
  }
  return '';
}

/**
 * Dokument-Attachment von einem DEX_Events-Item löschen.
 * Wird beim Edit verwendet, wenn der User ein bestehendes Dokument entfernt.
 */
export async function deleteEventDocument(svc: EventService, eventId: number, fileName: string): Promise<boolean> {
  try {
    const resp = await svc._delete(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`
    );
    return resp.ok || resp.status === 200 || resp.status === 204;
  } catch (err) {
    console.warn('[DEX] deleteEventDocument error:', err);
    return false;
  }
}

/**
 * Attachments eines DEX_Events-Items laden.
 * Bilder mit Präfix __eventimage__ werden ausgefiltert (nur für EventImageUrl).
 */
export async function getEventAttachments(svc: EventService, eventId: number): Promise<Array<{ name: string; url: string; size: number }>> {
  try {
    const response = await svc._sp.get(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Events')/items(${eventId})/AttachmentFiles`,
      SPHttpClient.configurations.v1
    );
    if (response.ok) {
      const data = await response.json();
      const files = data.value || data.d?.results || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return files
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .filter((f: any) => (f.FileName || '').indexOf('__eventimage__') !== 0)
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .map((f: any) => ({
          name: f.FileName || '',
          url: `${window.location.origin}${f.ServerRelativeUrl || ''}`,
          size: 0,
        }));
    }
  } catch { /* Attachments nicht verfügbar */ }
  return [];
}
