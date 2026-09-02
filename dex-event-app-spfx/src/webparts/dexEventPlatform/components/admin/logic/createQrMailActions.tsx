/* createQrMailActions — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 4099-4388 des
 * Stands vor dem Schnitt). Der Rumpf ist zeichengleich uebernommen; was die
 * Gruppe aus dem Komponenten-Scope liest, kommt als `ctx` herein, was sie
 * nach aussen liefert, geht als Objekt zurueck.
 */
import * as React from 'react';
import { DeloitteEvent } from '../../../types';
import { QrEmailOverride, buildQrBlockHtml, getCachedOrbBase64, injectIntoEmailContent, qrCodeEmail, qrEmailDefaults } from '../../../services/EmailTemplates';
import { SAMPLE_QR_ID } from '../../admin/adminConstants';
import { buildParticipantQrDataUrl } from '../../../utils/qrWithMark';
import { getCachedImage } from '../../../utils/imageCache';
import { MailHeaderImage, isDefaultMailHeaderImage, normalizeMailHeaderImage } from '../../../utils/mailHeaderImage';
import { isExternalEmail } from '../../../utils/deloitteDomain';
import { EventService, SPRegistration } from '../../../services/EventService';

export interface CreateQrMailActionsCtx {
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  currentUser: import("../../../types/index").User;
  eventServiceRef: EventService;
  /** v30.67 (Review): gemeinsamer Nachlade-Pfad der Seite — `null` = nicht lesbar. */
  reloadRegistrations: () => Promise<SPRegistration[] | null>;
  isDe: boolean;
  qrBlockLang: "" | "DE" | "EN";
  qrBlockNote: string;
  qrEditBody: string;
  qrEditHeading: string;
  qrEditSaving: boolean;
  qrEditSubheading: string;
  qrEditSubject: string;
  qrEditTarget: DeloitteEvent;
  qrHeaderImage: MailHeaderImage;
  refreshEvents: () => Promise<void>;
  registrations: SPRegistration[];
  sciBusy: boolean;
  sciFrom: string;
  sciTo: string;
  selectedEvent: DeloitteEvent;
  setIsSendingQR: React.Dispatch<React.SetStateAction<boolean>>;
  setQrBlockLang: React.Dispatch<React.SetStateAction<"" | "DE" | "EN">>;
  setQrBlockNote: React.Dispatch<React.SetStateAction<string>>;
  setQrEditBody: React.Dispatch<React.SetStateAction<string>>;
  setQrEditHeading: React.Dispatch<React.SetStateAction<string>>;
  setQrEditOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setQrEditSampleBlock: React.Dispatch<React.SetStateAction<string>>;
  setQrEditSampleImg: React.Dispatch<React.SetStateAction<string>>;
  setQrEditSaving: React.Dispatch<React.SetStateAction<boolean>>;
  setQrEditSubheading: React.Dispatch<React.SetStateAction<string>>;
  setQrEditSubject: React.Dispatch<React.SetStateAction<string>>;
  setQrEditTarget: React.Dispatch<React.SetStateAction<DeloitteEvent>>;
  setQrEventPhotoB64: React.Dispatch<React.SetStateAction<string>>;
  setQrHeaderImage: React.Dispatch<React.SetStateAction<MailHeaderImage>>;
  setQrPreviewHtml: React.Dispatch<React.SetStateAction<string>>;
  setQrPreviewLoading: React.Dispatch<React.SetStateAction<boolean>>;
  setQrPreviewOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setQrPreviewSubject: React.Dispatch<React.SetStateAction<string>>;
  setQrSendModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setQrSendResult: React.Dispatch<React.SetStateAction<string>>;
  setQrSentCount: React.Dispatch<React.SetStateAction<number>>;
  setSciBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setSciSaveMsg: React.Dispatch<React.SetStateAction<string>>;
  setSelectedEvent: React.Dispatch<React.SetStateAction<DeloitteEvent>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  updateEvent: (eventId: string, updates: Record<string, unknown>, opts?: { skipReload?: boolean; }) => Promise<boolean>;
}

export interface CreateQrMailActionsResult {
  closeQrMailEditor: () => void;
  getQrMailOverride: (ev: DeloitteEvent | null) => QrEmailOverride | undefined;
  openQrMailEditor: (target?: DeloitteEvent) => Promise<void>;
  qrFullSendAction: () => Promise<void>;
  qrPreviewAction: () => Promise<void>;
  qrTestSendAction: (liveOverride?: QrEmailOverride, target?: DeloitteEvent) => Promise<void>;
  saveQrMailOverride: () => Promise<void>;
  saveSelfCheckInWindow: () => Promise<void>;
}

export function createQrMailActions(ctx: CreateQrMailActionsCtx): CreateQrMailActionsResult {
  const {
    confirmDialog, currentUser, eventServiceRef, isDe, qrBlockLang,
    qrBlockNote, qrEditBody, qrEditHeading, qrEditSaving, qrEditSubheading, qrEditSubject,
    qrEditTarget, qrHeaderImage, refreshEvents, registrations, reloadRegistrations, sciBusy, sciFrom, sciTo,
    selectedEvent, setIsSendingQR, setQrBlockLang, setQrBlockNote, setQrEditBody, setQrEditHeading,
    setQrEditOpen, setQrEditSampleBlock, setQrEditSampleImg, setQrEditSaving, setQrEditSubheading,
    setQrEditSubject, setQrEditTarget, setQrEventPhotoB64, setQrHeaderImage, setQrPreviewHtml,
    setQrPreviewLoading, setQrPreviewOpen, setQrPreviewSubject, setQrSendModalOpen,
    setQrSendResult, setQrSentCount, setSciBusy, setSciSaveMsg, setSelectedEvent,
    showAlert, updateEvent,
  } = ctx;
  // v22.6: QR-Versand-Aktionen als benannte Funktionen (vorher inline im Modal) —
  // macht das neue kompakte Querformat-Layout lesbar. Verhalten unverändert.
  // v22.18: pro-Event angepasster QR-Mail-Text — Override-Key 'QRCode' im
  // EmailTemplateOverrides-JSON des Events (übersteht den Wizard-Roundtrip,
  // weil Nicht-Unterstrich-Keys dort erhalten bleiben). QR-Block bleibt fix.
  const getQrMailOverride = (ev: DeloitteEvent | null): QrEmailOverride | undefined => {
    if (!ev) return undefined;
    try {
      const all = JSON.parse(ev.emailTemplateOverrides || '{}');
      const ov = all && all['QRCode'];
      if (ov && (ov.subject || ov.heading || ov.subheading || ov.bodyHtml)) return ov as QrEmailOverride;
    } catch { /* kein Override */ }
    return undefined;
  };
  // Editor öffnen: Felder aus Override (falls vorhanden) oder den Standard-
  // Texten vorbelegen + Beispiel-QR-Block für die Live-Vorschau erzeugen.
  // v29.26: optionales target — vom Hauptevent aus lassen sich die QR-Mails
  // der Sub-Events einzeln gestalten (je Sub-Event ein eigener Override auf
  // dessen Zeile; Versand und Auto-Versand des Sub-Events lesen genau den).
  const openQrMailEditor = async (target?: DeloitteEvent): Promise<void> => {
    const tgt = target || selectedEvent;
    if (!tgt) return;
    setQrEditTarget(target || null);
    const ov = getQrMailOverride(tgt);
    const def = qrEmailDefaults(tgt.emailLanguage || 'EN');
    setQrEditSubject((ov && ov.subject) || def.subject);
    setQrEditHeading((ov && ov.heading) || def.heading);
    setQrEditSubheading((ov && ov.subheading) || def.subheading);
    setQrEditBody((ov && ov.bodyHtml) || def.body);
    // v30.52: gespeichertes Kopf-Bild laden; Event-Foto für die Auswahl
    // nachziehen (leer = „Event-Foto" bleibt deaktiviert).
    setQrHeaderImage(normalizeMailHeaderImage(ov && ov.headerImage));
    setQrBlockLang((ov && ov.blockLang) || '');
    setQrBlockNote((ov && ov.blockNote) || '');
    setQrEventPhotoB64('');
    if (tgt.imageUrl) {
      getCachedImage(tgt.imageUrl)
        .then(b64 => { if (b64 && b64.indexOf('data:') === 0) setQrEventPhotoB64(b64); })
        .catch(() => { /* Foto nicht ladbar → Option bleibt deaktiviert */ });
    }
    // Beispiel-QR (eigene Daten) für die Vorschau — gleicher Aufbau wie im Versand.
    const myName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email;
    const qrData = `DEX|${tgt.eventNumber}|${currentUser.email}`;
    // v30.33: Beispiel-ID, damit im Editor sichtbar ist, dass die
    // Teilnehmer-ID mitgeschickt wird — beim echten Versand steht dort die
    // tatsaechliche Nummer der Person.
    const qrImageHtml = await buildQrImageHtml(qrData);
    setQrEditSampleImg(qrImageHtml);
    setQrEditSampleBlock(buildQrBlockHtml(qrImageHtml, myName, SAMPLE_QR_ID, ((ov && ov.blockLang) || tgt.emailLanguage || 'EN'), (ov && ov.blockNote) || ''));
    // v22.19: Versand-Modal schließen — der Editor zeigt die Versand-Aktionen
    // in einer eigenen linken Spalte (nebeneinander statt übereinander).
    // Beim Schließen des Editors öffnet das Versand-Modal wieder.
    setQrSendModalOpen(false);
    setQrEditOpen(true);
  };
  const closeQrMailEditor = (): void => {
    setQrEditOpen(false);
    setQrEditTarget(null);
    setQrSendModalOpen(true);
  };
  // Speichern: Override in das EmailTemplateOverrides-JSON des Events mergen
  // (andere Keys + Piggybacks bleiben erhalten). Entspricht alles den
  // Standard-Texten, wird der Key entfernt (= zurück auf Standard).
  const saveQrMailOverride = async (): Promise<void> => {
    // v29.26: schreibt auf das EDITOR-ZIEL — das kann ein vom Hauptevent aus
    // geöffnetes Sub-Event sein (qrEditTarget), sonst das Event selbst.
    const tgt = qrEditTarget || selectedEvent;
    if (!tgt || qrEditSaving) return;
    setQrEditSaving(true);
    try {
      const def = qrEmailDefaults(tgt.emailLanguage || 'EN');
      // v30.52: Das Kopf-Bild zählt mit. Ohne diese Bedingung würde eine
      // geänderte Bildbreite bei sonst unveränderten Texten den Override
      // LÖSCHEN — die Einstellung wäre nach dem Speichern weg.
      const isDefault = qrEditSubject.trim() === def.subject.trim()
        && qrEditHeading.trim() === def.heading.trim()
        && qrEditSubheading.trim() === def.subheading.trim()
        && qrEditBody.trim() === def.body.trim()
        // v30.60: Ohne diese Bedingung würde eine allein geänderte
        // Block-Sprache den Override löschen — die Auswahl wäre nach dem
        // Speichern wieder weg (dieselbe Falle wie beim Kopf-Bild, v30.52).
        && !qrBlockLang
        && !qrBlockNote.trim()
        && isDefaultMailHeaderImage(qrHeaderImage);
      let all: Record<string, unknown> = {};
      try { all = JSON.parse(tgt.emailTemplateOverrides || '{}') || {}; } catch { all = {}; }
      if (isDefault) {
        delete all['QRCode'];
      } else {
        all['QRCode'] = {
          subject: qrEditSubject, heading: qrEditHeading, subheading: qrEditSubheading, bodyHtml: qrEditBody,
          // Nur Auswahl + Zahlen — NIE das Foto selbst (s. QrEmailOverride).
          headerImage: { ...qrHeaderImage },
          ...(qrBlockLang ? { blockLang: qrBlockLang } : {}),
          ...(qrBlockNote.trim() ? { blockNote: qrBlockNote.trim() } : {}),
        };
      }
      const json = JSON.stringify(all);
      const ok = await updateEvent(tgt.id, { 'EmailTemplateOverrides': json });
      if (ok) {
        if (!qrEditTarget || (selectedEvent && qrEditTarget.id === selectedEvent.id)) {
          setSelectedEvent(prev => prev ? { ...prev, emailTemplateOverrides: json } : prev);
        }
        // Lokale Kopie des Ziels nachziehen — sonst vergleicht der Editor
        // gegen den alten Stand und meldet weiter „ungespeichert".
        setQrEditTarget(prev => (prev && prev.id === tgt.id) ? { ...prev, emailTemplateOverrides: json } : prev);
        await refreshEvents();
        showAlert(isDe
          ? (isDefault
            ? 'QR-Mail auf den Standardtext zurückgesetzt.'
            : 'QR-Mail-Text gespeichert — gilt ab jetzt für Vorschau, Versand UND den automatischen QR-Versand bei neuen Anmeldungen.')
          : (isDefault ? 'QR email reset to the default text.' : 'QR email text saved — now used for preview, sending AND the automatic QR send for new registrations.'), { variant: 'success' });
      } else {
        showAlert(isDe ? 'Speichern fehlgeschlagen — vermutlich fehlen Schreibrechte auf der Event-Liste.' : 'Saving failed — you probably lack write permission on the event list.', { variant: 'error' });
      }
    } finally { setQrEditSaving(false); }
  };
  /**
   * v30.33: Unter den QR-Code kommt die Teilnehmer-ID — groß und vorlesbar.
   *
   * Sie ist nicht nur Deko: Seit v30.33 kann das Check-in-Team die ID ins
   * Suchfeld tippen und die Person damit einchecken. Das ist der Weg, der ohne
   * Kamera auskommt — und auf verwalteten Geräten ist die Kamera nicht überall
   * erreichbar (SharePoint-App-WebView, Android-Foto-Picker). Steht die ID
   * nicht in der Mail, kann niemand sie nennen, und der ganze Weg läuft leer.
   *
   * Bewusst monospace und groß: Die Zahl wird am Einlass vorgelesen und
   * abgetippt, nicht gelesen.
   */
  /**
   * v30.36: Erzeugung liegt jetzt in `utils/qrWithMark` — gemeinsam mit dem
   * Auto-Versand im EventContext. Vorher gab es zwei Erzeuger, die
   * auseinanderliefen (unterschiedliche Fehlerkorrektur), und das faellt
   * niemandem auf: Der Code scannt einfach schlechter.
   */
  const buildQrImageHtml = async (qrData: string): Promise<string> => {
    const qrDataUrl = await buildParticipantQrDataUrl(qrData, 300);
    if (!qrDataUrl) {
      return `<p style="font-family:monospace;font-size:1.2rem;background:#f5f5f5;padding:12px;border-radius:8px;text-align:center;">${qrData}</p>`;
    }
    return `<img src="${qrDataUrl}" alt="QR-Code" style="width:300px;max-width:100%;height:auto;" />`;
  };

  /**
   * v30.52: Löst das Kopf-Bild der QR-Mail auf.
   *
   * Nur wenn im Override „Event-Foto" gewählt ist, wird das Bild als Base64
   * geholt und fest eingebacken; sonst bleibt `{{ORB_URL}}` stehen und der
   * Flow setzt wie bisher das Standard-Bild. `getCachedImage` cacht — pro
   * Sitzung fällt der Abruf also einmal an, nicht je Teilnehmer.
   */
  const qrHeroPhotoFor = async (ev: DeloitteEvent, override?: QrEmailOverride): Promise<string> => {
    const hdr = normalizeMailHeaderImage(override && override.headerImage);
    if (hdr.hero !== 'event' || !ev.imageUrl) return '';
    try {
      const b64 = await getCachedImage(ev.imageUrl);
      return (b64 && b64.indexOf('data:') === 0) ? b64 : '';
    } catch { return ''; }
  };

  const qrPreviewAction = async (): Promise<void> => {
    if (!selectedEvent) return;
    setQrPreviewLoading(true);
    try {
      const orgEmail = currentUser.email;
      const orgFullName = `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || orgEmail;
      const orgFirstName = currentUser.firstName || orgFullName.split(/\s+/)[0] || orgFullName;
      const qrData = `DEX|${selectedEvent.eventNumber}|${orgEmail}`;
      const qrImageHtml = await buildQrImageHtml(qrData);
      const qrOv = getQrMailOverride(selectedEvent);
      const emailData = qrCodeEmail(orgFirstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN', orgFullName, qrOv, SAMPLE_QR_ID, await qrHeroPhotoFor(selectedEvent, qrOv));
      let eventOrb = '';
      try {
        const ov = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
        if (ov && typeof ov._eventLogo === 'string') eventOrb = ov._eventLogo;
      } catch { /* */ }
      const previewBody = emailData.body.replace(/\{\{ORB_URL\}\}/g, eventOrb || getCachedOrbBase64() || '');
      setQrPreviewSubject(emailData.subject);
      setQrPreviewHtml(previewBody);
      setQrPreviewOpen(true);
    } finally { setQrPreviewLoading(false); }
  };
  // v22.19: optionaler liveOverride — der Test-Versand aus dem Mail-Editor
  // nutzt den AKTUELLEN (ggf. ungespeicherten) Editor-Text, damit Test = Vorschau.
  // v29.26: optionales target — Test aus dem Editor eines Sub-Events nutzt
  // dessen Titel/Sprache/Event-Nummer statt der des geöffneten Events.
  const qrTestSendAction = async (liveOverride?: QrEmailOverride, target?: DeloitteEvent): Promise<void> => {
    const ev = target || selectedEvent;
    if (!eventServiceRef || !ev) return;
    setIsSendingQR(true); setQrSendResult(null); setQrSentCount(0);
    try {
      // v24.99: Test-Mail geht an ALLE Organisatoren des Events (vorher nur an
      // den eingeloggten User). Fallback: wenn keine Organisatoren hinterlegt
      // sind, an mich selbst. Jeder bekommt einen QR mit der EIGENEN Adresse.
      const orgEmails = (ev.organizerEmails || []).map(e => (e || '').trim()).filter(Boolean);
      const orgNames = ev.organizers || [];
      const recipients = orgEmails.length > 0
        ? orgEmails.map((em, i) => ({ email: em, rawName: orgNames[i] || em }))
        : [{ email: currentUser.email, rawName: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim() || currentUser.email }];
      // v30.52: Test = Vorschau — also auch beim Kopf-Bild den AKTUELLEN
      // Editor-Stand nehmen, wenn der Test aus dem Editor kommt.
      const testOverride = liveOverride || getQrMailOverride(ev);
      const testHeroPhoto = await qrHeroPhotoFor(ev, testOverride);
      let sent = 0;
      for (const r of recipients) {
        const raw = (r.rawName || '').trim();
        // Deloitte-Displayname „Nachname, Vorname" → „Vorname Nachname" + Vorname.
        const fullName = raw.indexOf(',') >= 0 ? raw.split(',').reverse().map(s => s.trim()).join(' ') : (raw || r.email);
        const firstName = raw.indexOf(',') >= 0 ? (raw.substring(raw.indexOf(',') + 1).trim().split(/\s+/)[0] || fullName) : (fullName.split(/\s+/)[0] || fullName);
        const qrData = `DEX|${ev.eventNumber}|${r.email}`;
        const qrImageHtml = await buildQrImageHtml(qrData);
        const emailData = qrCodeEmail(firstName, ev.title, qrImageHtml, ev.emailLanguage || 'EN', fullName, testOverride, SAMPLE_QR_ID, testHeroPhoto);
        await eventServiceRef.queueEmail(emailData.subject, r.email, fullName, emailData.body, 'QRCode', ev.title, ev.id);
        sent++; setQrSentCount(sent);
      }
      setQrSendResult(isDe
        ? `Test-Mail an ${sent} Organisator${sent === 1 ? '' : 'en'} verschickt — bitte im Postfach prüfen.`
        : `Test email sent to ${sent} organizer${sent === 1 ? '' : 's'} — please check the mailbox.`);
    } catch (err) {
      setQrSendResult((isDe ? 'Fehler beim Test-Versand: ' : 'Error during test send: ') + (err instanceof Error ? err.message : String(err)));
    }
    setIsSendingQR(false);
  };
  const qrFullSendAction = async (): Promise<void> => {
    if (!eventServiceRef || !selectedEvent) return;
    const eligible = registrations.filter(r => r.Status === 'Angemeldet');
    if (eligible.length === 0) {
      setQrSendResult(isDe ? 'Alle Teilnehmer haben bereits einen QR-Code — nichts zu senden.' : 'All participants already have a QR code — nothing to send.');
      return;
    }
    if (!(await confirmDialog(isDe ? `QR-Code an ${eligible.length} Teilnehmer ohne Code senden?` : `Send the QR code to ${eligible.length} participants without a code?`, { confirmLabel: isDe ? 'Senden' : 'Send' }))) return;
    setIsSendingQR(true); setQrSendResult(null); setQrSentCount(0);
    let sent = 0; let extCount = 0;
    for (const reg of eligible) {
      const qrData = `DEX|${selectedEvent.eventNumber}|${reg.ParticipantEmail}`;
      const name = (reg.Vorname && reg.Nachname) ? `${reg.Vorname} ${reg.Nachname}` : reg.ParticipantName;
      const firstName = reg.Vorname || (reg.ParticipantName || '').trim().split(/\s+/)[0] || name;
      const qrImageHtml = await buildQrImageHtml(qrData);
      const sendOv = getQrMailOverride(selectedEvent);
      const emailData = qrCodeEmail(firstName, selectedEvent.title, qrImageHtml, selectedEvent.emailLanguage || 'EN', name, sendOv, reg.TeilnehmerID, await qrHeroPhotoFor(selectedEvent, sendOv));
      // v27.11: Member-Firm-Adressen zählen als intern → QR-Mail direkt.
      const isExternal = isExternalEmail(reg.ParticipantEmail);
      if (isExternal) {
        const orgEmails = (selectedEvent.organizerEmails || []).filter(Boolean);
        const orgRecipient = orgEmails.length > 0 ? orgEmails.join(';') : currentUser.email;
        const orgSubject = `[Externer Teilnehmer] QR-Code für ${name} — ${selectedEvent.title}`;
        const qrExternalHint = `<div style="margin:0 0 16px;padding:12px 16px;background:#fff3e0;border:1px solid #ed8b00;border-radius:8px;font-size:13px;line-height:1.55;color:#7a4a00;">`
          + `<strong>QR-Code für externen Teilnehmer.</strong><br>`
          + `Eigentlich für <strong>${reg.ParticipantEmail}</strong> (${name}). Da externe Adressen keinen Mail-Versand bekommen, landet der QR-Code bei dir — drucke ihn aus oder leite die Mail intern an den Empfänger weiter (Datenschutzrichtlinien Deloitte Deutschland beachten).`
          + `</div>`;
        const qrBody = injectIntoEmailContent(emailData.body, qrExternalHint);
        await eventServiceRef.queueEmail(orgSubject, orgRecipient, 'Organizer', qrBody, 'QRCode', selectedEvent.title, selectedEvent.id);
        extCount++;
      } else {
        await eventServiceRef.queueEmail(emailData.subject, reg.ParticipantEmail, name, emailData.body, 'QRCode', selectedEvent.title, selectedEvent.id);
      }
      if (selectedEvent.subsiteUrl) {
        await eventServiceRef.setQRSentStatus(selectedEvent.subsiteUrl, reg.Id);
      }
      sent++; setQrSentCount(sent);
    }
    // v21: Erster Massen-Versand startet die QR-Phase (AutoSendQRCode=true).
    try { await eventServiceRef.updateEvent(parseInt(selectedEvent.id, 10), { AutoSendQRCode: true }); } catch { /* */ }
    // v30.67 (Review): gemeinsamer Nachlade-Pfad — nach einem Massen-Versand
    // ist die 429 auf dem Reload der Normalfall, die Liste wurde dann `[]`.
    await reloadRegistrations();
    setIsSendingQR(false);
    setQrSendResult(extCount > 0
      ? (isDe
        ? `${sent} QR-Codes verschickt (davon ${extCount} an dich/Organizer umgeleitet — externe Adressen).`
        : `${sent} QR codes sent (${extCount} of them redirected to you/the organizer — external addresses).`)
      : (isDe ? `${sent} QR-Codes verschickt.` : `${sent} QR codes sent.`));
  };

  const saveSelfCheckInWindow = async (): Promise<void> => {
    if (!selectedEvent || sciBusy) return;
    if (sciFrom && sciTo && new Date(sciFrom).getTime() >= new Date(sciTo).getTime()) {
      showAlert(isDe ? '„Bis" muss zeitlich nach „Von" liegen.' : '"Until" must be after "From".');
      return;
    }
    setSciBusy(true);
    try {
      const ok = await updateEvent(selectedEvent.id, {
        'SelfCheckInFrom': sciFrom ? new Date(sciFrom).toISOString() : null,
        'SelfCheckInTo': sciTo ? new Date(sciTo).toISOString() : null,
      });
      setSciSaveMsg(ok
        ? (isDe ? 'Zeitfenster gespeichert.' : 'Time window saved.')
        : (isDe ? 'Speichern fehlgeschlagen — bitte erneut versuchen.' : 'Saving failed — please try again.'));
    } finally { setSciBusy(false); }
  };
  return {
    closeQrMailEditor, getQrMailOverride, openQrMailEditor, qrFullSendAction, qrPreviewAction,
    qrTestSendAction, saveQrMailOverride, saveSelfCheckInWindow,
  };
}

