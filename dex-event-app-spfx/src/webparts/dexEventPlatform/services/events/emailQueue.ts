/**
 * v30.6 — Modularisierung Stufe 2 (CLAUDE.md): Thema „DEX_Emails-Queue".
 *
 * Herausgelöst aus der EventService-Klasse (dort stehen Delegations-Stubs
 * mit unveränderter Signatur — KEINE Aufrufstelle musste angefasst werden).
 * Die Funktionen bekommen die Service-Instanz als ersten Parameter und
 * nutzen deren öffentliche Infrastruktur (`_sp`, `_post`, `_merge`,
 * `siteUrl`, listExists, configureDefaultView, …). Der Unterstrich bleibt
 * als „intern"-Signal — neue UI-Aufrufer gehen weiter über die Klasse.
 *
 * Rezept für weitere Themen: Block hierher kopieren, `this.` → `svc.`,
 * Methodenkopf → `export async function name(svc: EventService, …)`,
 * in der Klasse einen Delegations-Stub stehen lassen, `tsc` treiben lassen
 * (private Helfer, die das Modul braucht, werden public).
 */

import { SPHttpClient } from '@microsoft/sp-http';
import { normalizeMadeWithLink } from '../EmailTemplates';
import type { EventService } from '../EventService';

/**
 * E-Mail-Queue-Liste erstellen falls nicht vorhanden.
 * Power Automate reagiert auf neue Einträge und versendet Mails.
 *
 * Spalten: Title (Betreff), Recipient, RecipientName, Body (HTML),
 * EmailType, EventTitle, EventId, Status (Pending/Sent/Failed), SentDate,
 * Cc/Bcc, Importance.
 */
export async function ensureEmailsList(svc: EventService): Promise<void> {
  const listName = 'DEX_Emails';
  const exists = await svc.listExists(listName);
  if (exists) {
    // Recipient-Feld auf Plain Text (RichText=false) setzen, falls es
    // noch im alten RichText-Modus ist. SharePoint wrappt sonst den Wert
    // in <div class="ExternalClassXXXX">...</div>, was den Power Automate
    // Flow "emailMessage/To must be String/email" Fehler auslöst.
    try {
      await setRecipientFieldPlainText(svc, listName);
    } catch { /* ignore */ }

    // Cc-Feld nachträglich anlegen (für Anfrage-Mails von der Landing Page).
    // Bestehende Listen aus aelteren App-Versionen haben das Feld noch nicht.
    try {
      await ensureCcFieldExists(svc, listName);
    } catch { /* ignore */ }

    // v8.5: Bcc-Feld nachträglich anlegen — wird genutzt um Organizer
    // automatisch in Anmelde-/Abmelde-Bestätigungen zu BCC'en, ohne den
    // Teilnehmer den Verteiler zu zeigen.
    try {
      await ensureBccFieldExists(svc, listName);
    } catch { /* ignore */ }

    // v18.30: Importance-Feld nachträglich anlegen — der DEX_SEND_MAIL-Flow
    // setzt darauf basierend die Outlook-Wichtigkeit (High = rotes „!").
    // Leer/„Normal" = normale Wichtigkeit.
    try {
      await ensureImportanceFieldExists(svc, listName);
    } catch { /* ignore */ }

    // Berechtigungen prüfen
    try {
      const listInfo = await svc._sp.get(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=HasUniqueRoleAssignments`,
        SPHttpClient.configurations.v1
      );
      if (listInfo.ok) {
        const data = await listInfo.json();
        if (!data.HasUniqueRoleAssignments) {
          await setEmailsListPermissions(svc, listName);
        }
      }
    } catch { /* ignore */ }
    return;
  }

  await svc._post(`${svc.siteUrl}/_api/web/lists`, {
    '__metadata': { 'type': 'SP.List' },
    'Title': listName,
    'Description': 'E-Mail-Queue für automatischen Versand via Power Automate',
    'BaseTemplate': 100,
    'AllowContentTypes': false,
  });

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const fields: Array<Record<string, any>> = [
    // Recipient als Plain-Text Note-Feld (RichText=false), damit der Flow
    // die Email-Adresse(n) ohne HTML-Wrapping bekommt.
    { title: 'Recipient', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 },
    { title: 'Cc', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 },
    { title: 'Bcc', type: 3, metaType: 'SP.FieldMultiLineText', richText: false, numberOfLines: 3 },
    { title: 'RecipientName', type: 2 },
    { title: 'Body', type: 3 }, // Body darf Rich/HTML bleiben (wird als HTML gerendert)
    { title: 'EmailType', type: 6, choices: ['Anmeldung', 'Abmeldung', 'Warteliste', 'Nachruecken', 'Info'], metaType: 'SP.FieldChoice' },
    { title: 'EventTitle', type: 2 },
    { title: 'EventId', type: 2 },
    { title: 'Status', type: 6, choices: ['Pending', 'Sent', 'Failed'], metaType: 'SP.FieldChoice' },
    { title: 'SentDate', type: 4 },
    // v18.30: Outlook-Wichtigkeit (leer/„Normal" = normal, „High" = rotes „!").
    { title: 'Importance', type: 2 },
  ];

  for (const f of fields) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: Record<string, any> = {
      '__metadata': { 'type': f.metaType || 'SP.Field' },
      'Title': f.title,
      'FieldTypeKind': f.type,
      'Required': false,
    };
    if (f.choices) {
      payload['Choices'] = { 'results': f.choices };
    }
    if (f.metaType === 'SP.FieldMultiLineText') {
      payload['RichText'] = !!f.richText;
      if (typeof f.numberOfLines === 'number') payload['NumberOfLines'] = f.numberOfLines;
    }
    await svc._post(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, payload);
  }

  // Default View
  await svc.configureDefaultView(listName, [
    'Recipient', 'RecipientName', 'EmailType', 'EventTitle', 'Status', 'SentDate',
  ]);

  await setEmailsListPermissions(svc, listName);
}

/**
 * Recipient-Feld auf Plain Text (RichText=false) umstellen.
 * Idempotent: Wenn schon Plain Text, macht nichts.
 * Nur möglich wenn der Current User Manage Lists Rechte hat (Owner/Admin).
 */
async function setRecipientFieldPlainText(svc: EventService, listName: string): Promise<void> {
  const fieldUrl = `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Recipient')`;
  const resp = await svc._sp.get(
    `${fieldUrl}?$select=FieldTypeKind,RichText`,
    SPHttpClient.configurations.v1
  );
  if (!resp.ok) return;
  const data = await resp.json();
  const kind = data.FieldTypeKind ?? data.d?.FieldTypeKind;
  const richText = data.RichText ?? data.d?.RichText;
  // Wenn bereits Note + Plain Text: nichts zu tun
  if (kind === 3 && richText === false) return;
  // Feld auf Note + RichText=false patchen
  await svc._merge(
    fieldUrl,
    {
      '__metadata': { 'type': 'SP.FieldMultiLineText' },
      'FieldTypeKind': 3,
      'RichText': false,
      'NumberOfLines': 3,
    }
  );
}

/**
 * Cc-Feld auf DEX_Emails anlegen, falls noch nicht vorhanden.
 * Multi-line Plain-Text damit auch ;-separierte Mehrfach-Adressen passen.
 */
async function ensureCcFieldExists(svc: EventService, listName: string): Promise<void> {
  const probeUrl = `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Cc')?$select=Id`;
  const probe = await svc._sp.get(probeUrl, SPHttpClient.configurations.v1);
  if (probe.ok) return;
  await svc._post(
    `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
    {
      '__metadata': { 'type': 'SP.FieldMultiLineText' },
      'Title': 'Cc',
      'FieldTypeKind': 3,
      'Required': false,
      'RichText': false,
      'NumberOfLines': 3,
    }
  );
}

async function ensureBccFieldExists(svc: EventService, listName: string): Promise<void> {
  const probeUrl = `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Bcc')?$select=Id`;
  const probe = await svc._sp.get(probeUrl, SPHttpClient.configurations.v1);
  if (probe.ok) return;
  await svc._post(
    `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
    {
      '__metadata': { 'type': 'SP.FieldMultiLineText' },
      'Title': 'Bcc',
      'FieldTypeKind': 3,
      'Required': false,
      'RichText': false,
      'NumberOfLines': 3,
    }
  );
}

// v18.30: Importance-Spalte (Single line text) idempotent anlegen. Der
// DEX_SEND_MAIL-Flow liest sie und sendet bei „High" mit hoher Wichtigkeit.
async function ensureImportanceFieldExists(svc: EventService, listName: string): Promise<void> {
  const probeUrl = `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('Importance')?$select=Id`;
  const probe = await svc._sp.get(probeUrl, SPHttpClient.configurations.v1);
  if (probe.ok) return;
  await svc._post(
    `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
    {
      '__metadata': { 'type': 'SP.Field' },
      'Title': 'Importance',
      'FieldTypeKind': 2,
      'Required': false,
    }
  );
}

/**
 * Berechtigungen für DEX_Emails: Owners Full Control, Members Contribute, Item-Level Security
 */
export async function setEmailsListPermissions(svc: EventService, listName: string): Promise<void> {
  try {
    await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
      {}
    );
    const ownersResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/associatedownergroup?$select=Id`, SPHttpClient.configurations.v1
    );
    if (ownersResp.ok) {
      const d = await ownersResp.json();
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741829)`, {}
      );
    }
    const deallId = await svc.getVisitorsGroupId();
    if (deallId) {
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${deallId}, roledefid=1073741827)`, {}
      );
    }
  } catch { /* */ }

  // Item-Level Security (v26.87: zuverlässiger nometadata-MERGE)
  await svc._setListSecurity(`${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')`, { ReadSecurity: 2, WriteSecurity: 2 });
}

/**
 * Berechtigungen für Queue-Listen (DEX_Outlook, DEX_IDReorder):
 * Owners Full Control, Members Contribute, Item-Level Security
 */
export async function setQueueListPermissions(svc: EventService, listName: string): Promise<void> {
  try {
    await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`,
      {}
    );
    const ownersResp = await svc._sp.get(
      `${svc.siteUrl}/_api/web/associatedownergroup?$select=Id`, SPHttpClient.configurations.v1
    );
    if (ownersResp.ok) {
      const d = await ownersResp.json();
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${d.Id}, roledefid=1073741829)`, {}
      );
    }
    const deallId = await svc.getVisitorsGroupId();
    if (deallId) {
      await svc._post(
        `${svc.siteUrl}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment(principalid=${deallId}, roledefid=1073741827)`, {}
      );
    }
  } catch { /* */ }
}

/**
 * E-Mail in die Queue eintragen (wird von Power Automate versendet).
 */
export async function queueEmail(
  svc: EventService,
  subject: string,
  recipient: string,
  recipientName: string,
  body: string,
  emailType: string,
  eventTitle: string,
  eventId: string,
  cc?: string,
  bcc?: string,
  // v18.30: 'High' = Outlook hohe Wichtigkeit (rotes „!"). Default normal.
  importance?: 'High' | 'Normal',
  // v26.62: Optionaler Datei-Anhang an der Queue-Zeile (z. B. der fertige
  // .eml-Einladungs-Entwurf bei externen Anmeldungen). Der DEX_SEND_MAIL-
  // Flow muss die Item-Attachments an die ausgehende Mail anhängen
  // (Get attachments → Get attachment content → Send-Email-Attachments);
  // solange der Flow das nicht tut, wird der Anhang schlicht ignoriert.
  attachment?: { fileName: string; content: string }
): Promise<boolean> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: Record<string, any> = {
      '__metadata': { 'type': 'SP.Data.DEX_x005f_EmailsListItem' },
      'Title': subject,
      'Recipient': recipient,
      'RecipientName': recipientName,
      // v29.42: Fußzeilen-Link kurz vor dem Versand auf die kanonische
      // App-Adresse ziehen — gespeicherte Vorlagen und kopierte Events
      // schleppen die Fußzeile älterer Stände mit.
      'Body': normalizeMadeWithLink(body),
      'EmailType': emailType,
      'EventTitle': eventTitle,
      'EventId': eventId,
      'Status': 'Pending',
    };
    if (cc) payload['Cc'] = cc;
    if (bcc) payload['Bcc'] = bcc;
    if (importance === 'High') payload['Importance'] = 'High';
    const response = await svc._post(
      `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Emails')/items`,
      payload
    );
    if (!response.ok) return false;
    if (attachment && attachment.fileName && attachment.content) {
      try {
        const data = await response.json();
        const itemId = Number(data?.d?.Id ?? data?.Id ?? 0);
        if (itemId > 0) {
          const buf = new TextEncoder().encode(attachment.content);
          const safeName = attachment.fileName.replace(/[^a-zA-Z0-9._@-]+/g, '_');
          await svc._sp.post(
            `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Emails')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(safeName)}')`,
            SPHttpClient.configurations.v1,
            { headers: { 'Accept': 'application/json;odata=nometadata' }, body: buf.buffer as ArrayBuffer }
          );
        }
      } catch (attErr) {
        // Anhang best-effort — die Mail selbst ist wichtiger als der Anhang.
        console.warn('[DEX] queueEmail: Anhang konnte nicht angehängt werden:', attErr);
      }
    }
    return true;
  } catch {
    return false;
  }
}

/** v26.12: Gibt es bereits eine Mail dieses Typs für dieses Event in der
 *  DEX_Emails-Queue? Dient als serverseitiger Doppelversand-Schutz, wenn die
 *  Mail clientseitig (App-Open) ausgelöst wird und mehrere Organizer die App
 *  öffnen könnten. */
export async function hasQueuedEmail(svc: EventService, emailType: string, eventId: string): Promise<boolean> {
  try {
    const safeType = (emailType || '').replace(/'/g, "''");
    const safeId = (eventId || '').replace(/'/g, "''");
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Emails')/items?$select=Id&$filter=EmailType eq '${safeType}' and EventId eq '${safeId}'&$top=1`;
    const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (!resp.ok) return false;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    return Array.isArray(items) && items.length > 0;
  } catch {
    return false;
  }
}

/** v26.33: Wurde seit `sinceIso` (inkl.) schon eine Mail dieses Typs in die
 *  Queue gelegt? Grundlage für tages-entdoppelte Reminder (z.B. Ticket-
 *  Erinnerung höchstens einmal pro Tag, egal wie viele Power-User die App
 *  öffnen). */
export async function hasQueuedEmailSince(svc: EventService, emailType: string, sinceIso: string): Promise<boolean> {
  try {
    const safeType = (emailType || '').replace(/'/g, "''");
    const url = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Emails')/items?$select=Id&$filter=EmailType eq '${safeType}' and Created ge datetime'${encodeURIComponent(sinceIso)}'&$top=1`;
    const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
    if (!resp.ok) return false;
    const data = await resp.json();
    const items = data.value || data.d?.results || [];
    return Array.isArray(items) && items.length > 0;
  } catch {
    return false;
  }
}

/** v26.35: Für jedes Event das FRÜHESTE „ParticipantDeletionWarning"-Sendedatum
 *  (Created in der DEX_Emails-Queue). Grundlage für die 1-Wochen-Frist zwischen
 *  Vorwarnung an die Organizer und der Löschung der Teilnehmerliste. */
export async function getParticipantDeletionWarningDates(svc: EventService): Promise<Record<string, string>> {
  const out: Record<string, string> = {};
  try {
    let url: string | null = `${svc.siteUrl}/_api/web/lists/getbytitle('DEX_Emails')/items?$select=EventId,Created&$filter=EmailType eq 'ParticipantDeletionWarning'&$top=5000`;
    while (url) {
      const resp = await svc._sp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) break;
      const data = await resp.json();
      const items = data.value || data.d?.results || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const it of (items as any[])) {
        const eid = String(it.EventId || '');
        const created = it.Created || '';
        if (!eid || !created) continue;
        if (!out[eid] || new Date(created).getTime() < new Date(out[eid]).getTime()) out[eid] = created;
      }
      url = data['odata.nextLink'] || (data.d && data.d.__next) || null;
    }
  } catch { /* best-effort */ }
  return out;
}
