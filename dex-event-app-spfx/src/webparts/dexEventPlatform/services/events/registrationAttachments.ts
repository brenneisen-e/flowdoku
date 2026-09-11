/**
 * v30.66 — Modularisierung Stufe 2: Thema „Anhänge an der Anmeldezeile"
 * (Datei-Upload-Felder des Anmeldeformulars, Einladungs-EML).
 * Anhänge hängen am Listen-Item der Teilnehmerliste, nicht an einer
 * Bibliothek — deshalb laufen sie über AttachmentFiles.
 * Herausgelöst aus EventService; dort stehen Delegations-Stubs.
 */

import { SPHttpClient } from '@microsoft/sp-http';
import type { EventService } from '../EventService';
import { REG_LIST_NAME } from '../EventService';

/**
 * v31.21: Die Anhänge ALLER Zeilen einer Teilnehmerliste in EINER Abfrage.
 *
 * Der Grund ist ein Fehlerbericht, kein Aufräumen: Das Organizer Center rief
 * bis hierher `listRegistrationAttachments` in einer `for`-Schleife — **einen
 * HTTP-Request je Teilnehmerzeile, nacheinander**. Bei einem Termin mit 426
 * Anmeldungen sind das 426 Abfragen hintereinander; SharePoint drosselt nach
 * einigen Dutzend mit 429, der Client wartet, und die Seite wirkt für Minuten
 * eingefroren (Nutzer-Befund 11.09.2026: „wenn ich hier in eins der Sub-Events
 * springe, dann hängt sich die App fast komplett auf").
 *
 * `$expand=AttachmentFiles` liefert dasselbe in einer Abfrage je 5000er-Seite.
 * Zusätzlich filtert `Attachments eq 1` auf die Zeilen, die überhaupt eine
 * Datei haben — das ist bei fast jedem Event die kleine Minderheit.
 *
 * Rückgabe `null` heißt **nicht lesbar** und ist ausdrücklich nicht dasselbe
 * wie „keine Anhänge": Wer aus einem Lesefehler eine leere Map macht, zeigt
 * dem Organizer eine Liste ohne Büroklammern und behauptet damit, es gäbe
 * keine Dateien. (Dieselbe Regel wie bei `getAllRegistrations`, v30.67.)
 */
export async function listAllRegistrationAttachments(
  svc: EventService,
  subsiteUrl: string,
): Promise<Record<number, Array<{ fileName: string; serverRelativeUrl: string }>> | null> {
  const out: Record<number, Array<{ fileName: string; serverRelativeUrl: string }>> = {};
  let url: string | null = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items`
    + `?$select=Id,AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl`
    + `&$expand=AttachmentFiles&$filter=Attachments eq 1&$top=500`;
  let guard = 0;
  while (url && guard < 30) {
    guard++;
    try {
      const resp = await svc._sp.get(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' },
      });
      if (!resp.ok) {
        console.warn(`[DEX] listAllRegistrationAttachments: HTTP ${resp.status}`);
        return null;
      }
      const data = await resp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const it of ((data.value || []) as any[])) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const files = ((it.AttachmentFiles || []) as any[])
          .map(a => ({ fileName: a.FileName || '', serverRelativeUrl: a.ServerRelativeUrl || '' }))
          .filter(x => !!x.fileName);
        if (files.length > 0) out[Number(it.Id)] = files;
      }
      url = data['odata.nextLink'] || data['@odata.nextLink'] || null;
    } catch (err) {
      console.warn('[DEX] listAllRegistrationAttachments failed:', err);
      return null;
    }
  }
  return out;
}

/**
 * v11.0: Item-Attachments einer Teilnehmer-Registrierung listen.
 * Liefert ein Array mit FileName + ServerRelativeUrl, sodass die App
 * Download-Links rendern kann. Subsite-spezifisch (jede Teilnehmerliste
 * lebt in der Event-Subsite).
 *
 * Für EINE Zeile — etwa nach einem Upload im Anhang-Dialog. Wer die Anhänge
 * einer ganzen Liste braucht, nimmt `listAllRegistrationAttachments`; ein
 * Aufruf je Zeile in einer Schleife ist der Fehler, den v31.21 behoben hat.
 */
export async function listRegistrationAttachments(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
): Promise<Array<{ fileName: string; serverRelativeUrl: string }>> {
  try {
    const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})/AttachmentFiles`;
    const resp = await svc._sp.get(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' },
    });
    if (!resp.ok) return [];
    const data = await resp.json();
    const items = data.value || [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return items.map((a: any) => ({
      fileName: a.FileName || '',
      serverRelativeUrl: a.ServerRelativeUrl || '',
    })).filter((x: { fileName: string }) => !!x.fileName);
  } catch (err) {
    console.warn('[DEX] listRegistrationAttachments failed:', err);
    return [];
  }
}

/**
 * v11.0: PDF / Datei als Item-Attachment an eine Teilnehmer-Zeile
 * hängen. SharePoint erlaubt mehrere Attachments pro Item; bei
 * gleichem Namen wirft die API einen 409, daher prefixen wir den
 * Dateinamen mit einem Timestamp wenn die App das aufruft.
 */
export async function addRegistrationAttachment(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  file: File,
  // v19.0: optionaler Präfix, um ein Attachment einem Dokument-Custom-Field
  // zuzuordnen (z.B. 'dxf-<fieldId>--'). Leer = generischer Attendee-Upload.
  fieldPrefix: string = '',
): Promise<boolean> {
  try {
    const buf = await file.arrayBuffer();
    // Dateiname säubern + Timestamp-prefix für Eindeutigkeit
    const safeName = (file.name || 'upload.pdf').replace(/[^a-zA-Z0-9._-]+/g, '_');
    const ts = new Date().toISOString().replace(/[:.]/g, '-').replace(/T/, '_').slice(0, 19);
    const finalName = `${fieldPrefix}${ts}_${safeName}`;
    const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(finalName)}')`;
    const resp = await svc._sp.post(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' },
      body: buf,
    });
    return resp.ok;
  } catch (err) {
    console.warn('[DEX] addRegistrationAttachment failed:', err);
    return false;
  }
}

/**
 * v11.0: Item-Attachment löschen. Wird sowohl vom User (eigener
 * Upload zurückziehen) als auch vom Admin (im Admin Center) genutzt.
 */
export async function deleteRegistrationAttachment(
  svc: EventService,
  subsiteUrl: string,
  itemId: number,
  fileName: string,
): Promise<boolean> {
  try {
    const url = `${subsiteUrl}/_api/web/lists/getbytitle('${REG_LIST_NAME}')/items(${itemId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`;
    const resp = await svc._sp.post(url, SPHttpClient.configurations.v1, {
      headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE', 'Accept': 'application/json;odata=nometadata' },
    });
    return resp.ok;
  } catch (err) {
    console.warn('[DEX] deleteRegistrationAttachment failed:', err);
    return false;
  }
}
