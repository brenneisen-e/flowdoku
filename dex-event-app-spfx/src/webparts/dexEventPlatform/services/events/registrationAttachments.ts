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
 * v11.0: Item-Attachments einer Teilnehmer-Registrierung listen.
 * Liefert ein Array mit FileName + ServerRelativeUrl, sodass die App
 * Download-Links rendern kann. Subsite-spezifisch (jede Teilnehmerliste
 * lebt in der Event-Subsite).
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
